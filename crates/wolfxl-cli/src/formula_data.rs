//! Cached workbook values shared by one-shot formulas and appended columns.

use std::collections::{BTreeMap, HashMap};

use anyhow::{anyhow, Result};
use wolfxl_core::CellValue;
use wolfxl_reader::{Cell as RawCell, CellDataType, CellValue as RawValue};

use crate::eval::{self, DataSource, ExcelError, Extent, Value};
use crate::filter::lenient_number;
use crate::input::{delimited_value, iso_serial, Input};

pub(crate) struct WorkbookData {
    sheets: BTreeMap<String, SheetData>,
    /// Date functions read and write serials in the 1904 date system.
    date1904: bool,
}

/// One sheet's non-blank values keyed by one-based `(row, column)`, plus the
/// used extent that bounds whole row and column references. Absent cells are
/// blank, so far-apart cells never allocate the rectangle between them.
#[derive(Default)]
struct SheetData {
    cells: HashMap<(u32, u32), Value>,
    extent: Extent,
}

impl WorkbookData {
    /// Load stored values only; existing formula cells retain their caches.
    pub(crate) fn load(input: &mut Input) -> Result<Self> {
        let names = input.sheet_names().to_vec();
        let book = input.raw_book()?;
        let mut sheets = BTreeMap::new();
        for name in names {
            let data = match &book {
                // Spreadsheet numbers stay as stored, whatever their number
                // format or the workbook's date system.
                Some(book) => stored_sheet(book.cells(&name)?),
                // Delimited fields take Excel's text-import types, and money
                // text imports as numbers.
                None => {
                    let sheet = input.raw_sheet(&name)?;
                    let (top, left) = sheet.origin();
                    let rows = sheet.rows();
                    let width = rows.iter().map(Vec::len).max().unwrap_or(0);
                    let mut data = SheetData::default();
                    if width > 0 {
                        data.extent = Extent {
                            rows: top + rows.len() as u32,
                            columns: left + width as u32,
                        };
                    }
                    for (row, fields) in (top + 1..).zip(rows) {
                        for (column, cell) in (left + 1..).zip(fields) {
                            if let CellValue::String(text) = &cell.value {
                                data.cells.insert((row, column), delimited_field(text));
                            }
                        }
                    }
                    data
                }
            };
            sheets.insert(name, data);
        }
        let date1904 = book.as_ref().is_some_and(|book| book.date1904());
        Ok(Self { sheets, date1904 })
    }

    /// Coordinates are one-based, like the evaluator's references. The value
    /// replaces any stored one, a blank included.
    pub(crate) fn set_value(&mut self, sheet: &str, row: u32, column: u32, value: Value) {
        let data = self.sheets.get_mut(sheet).expect("loaded target sheet");
        data.extent.rows = data.extent.rows.max(row);
        data.extent.columns = data.extent.columns.max(column);
        data.cells.insert((row, column), value);
    }
}

/// A spreadsheet's stored cells; the extent covers every cell the reader kept,
/// formulas without a cached value included.
fn stored_sheet(cells: Vec<RawCell>) -> SheetData {
    let mut data = SheetData::default();
    for cell in cells {
        data.extent.rows = data.extent.rows.max(cell.row);
        data.extent.columns = data.extent.columns.max(cell.col);
        let value = stored_value(cell.value, cell.data_type);
        if !matches!(value, Value::Blank) {
            data.cells.insert((cell.row, cell.col), value);
        }
    }
    data
}

/// A spreadsheet cell's stored value. ISO 8601 text that a reader typed as a
/// date (ODS date values, OOXML `t="d"` cells) evaluates as its serial.
fn stored_value(value: RawValue, data_type: CellDataType) -> Value {
    match value {
        RawValue::Empty => Value::Blank,
        RawValue::String(text) if data_type == CellDataType::DateIso => {
            iso_serial(&text).map_or(Value::Text(text), Value::Number)
        }
        RawValue::String(text) => Value::Text(text),
        RawValue::SharedString(text) => Value::Text(text.to_string()),
        RawValue::Number(number) if number.is_finite() => Value::Number(number),
        RawValue::Number(_) => Value::Error(ExcelError::Num),
        RawValue::Bool(flag) => Value::Boolean(flag),
        RawValue::Error(literal) => Value::Error(match literal.as_str() {
            "#DIV/0!" => ExcelError::Div0,
            "#REF!" => ExcelError::Ref,
            "#NAME?" => ExcelError::Name,
            "#N/A" => ExcelError::Na,
            "#NUM!" => ExcelError::Num,
            "#NULL!" => ExcelError::Null,
            _ => ExcelError::Value,
        }),
    }
}

/// A delimited field typed the way Excel's text import types it; text that
/// reads as money becomes a number once here, while display paths keep the
/// source spelling.
fn delimited_field(text: &str) -> Value {
    match delimited_value(text) {
        CellValue::Bool(flag) => Value::Boolean(flag),
        CellValue::Int(number) => Value::Number(number as f64),
        CellValue::Float(number) => Value::Number(number),
        _ => lenient_number(text).map_or_else(|| Value::Text(text.to_owned()), Value::Number),
    }
}

impl DataSource for WorkbookData {
    fn sheet_names(&self) -> impl Iterator<Item = &str> {
        self.sheets.keys().map(String::as_str)
    }

    fn cell_value(&self, sheet: &str, row: u32, column: u32) -> Value {
        self.sheets
            .get(sheet)
            .and_then(|data| data.cells.get(&(row, column)))
            .cloned()
            .unwrap_or(Value::Blank)
    }

    fn used_extent(&self, sheet: &str) -> Extent {
        self.sheets
            .get(sheet)
            .map_or(Extent::default(), |data| data.extent)
    }

    fn date1904(&self) -> bool {
        self.date1904
    }
}

pub(crate) fn normalize_formula(formula: &str) -> String {
    let formula = formula.trim();
    if formula.starts_with('=') {
        formula.to_owned()
    } else {
        format!("={formula}")
    }
}

pub(crate) fn parse_error(formula: &str, error: eval::ParseError) -> anyhow::Error {
    match error {
        eval::ParseError::Unsupported { function } => anyhow!(
            "unsupported function {function} in {formula}; use `wolfxl list-functions` to list supported functions"
        ),
        error => anyhow!("cannot parse {formula}: {error}"),
    }
}

pub(crate) fn parse_formula(formula: &str) -> Result<eval::Formula> {
    eval::parse(formula).map_err(|error| parse_error(formula, error))
}

/// Excel displays at most 15 significant digits.
pub(crate) fn rounded(number: f64) -> f64 {
    format!("{number:.14e}").parse().unwrap_or(number)
}

pub(crate) fn text_value(value: &Value) -> String {
    match value {
        Value::Blank => String::new(),
        Value::Number(number) if number.is_finite() => {
            let number = rounded(*number);
            if number == 0.0 {
                "0".to_owned()
            } else {
                number.to_string()
            }
        }
        Value::Number(_) => "#NUM!".to_owned(),
        Value::Text(text) => text.clone(),
        Value::Boolean(true) => "TRUE".to_owned(),
        Value::Boolean(false) => "FALSE".to_owned(),
        Value::Error(error) => error.as_str().to_owned(),
        Value::Array(_) => "#VALUE!".to_owned(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn results_use_fifteen_significant_digits_without_negative_zero() {
        assert_eq!(text_value(&Value::Number(0.1 + 0.2)), "0.3");
        assert_eq!(text_value(&Value::Number(1_377_000.0)), "1377000");
        assert_eq!(text_value(&Value::Number(-0.0)), "0");
    }

    fn raw_cell(row: u32, col: u32, value: RawValue) -> RawCell {
        RawCell {
            row,
            col,
            coordinate: String::new(),
            style_id: None,
            data_type: CellDataType::Number,
            value,
            formula: None,
            formula_kind: None,
            formula_shared_index: None,
            array_formula: None,
            rich_text: None,
        }
    }

    fn book(cells: Vec<RawCell>) -> WorkbookData {
        WorkbookData {
            sheets: BTreeMap::from([("Data".to_owned(), stored_sheet(cells))]),
            date1904: false,
        }
    }

    fn evaluated(data: &WorkbookData, formula: &str) -> Value {
        eval::evaluate(formula, data, "Data").expect("formula parses")
    }

    #[test]
    fn far_apart_stored_cells_keep_only_their_values() {
        let data = book(vec![
            raw_cell(1, 1, RawValue::Number(2.0)),
            RawCell {
                formula: Some("A1".to_owned()),
                ..raw_cell(5, 3, RawValue::Empty)
            },
            raw_cell(1_048_576, 16_384, RawValue::Number(3.0)),
        ]);
        let sheet = &data.sheets["Data"];
        assert_eq!(sheet.cells.len(), 2);
        assert_eq!(
            sheet.extent,
            Extent {
                rows: 1_048_576,
                columns: 16_384
            }
        );
        assert_eq!(evaluated(&data, "=SUM(A:A)"), Value::Number(2.0));
        assert_eq!(evaluated(&data, "=A1+XFD1048576"), Value::Number(5.0));
        assert_eq!(evaluated(&data, "=C5"), Value::Blank);
        assert_eq!(evaluated(&data, "=COUNT(C:C)"), Value::Number(0.0));
    }

    #[test]
    fn set_values_replace_stored_cells_and_grow_the_extent() {
        let mut data = book(vec![raw_cell(1, 1, RawValue::Number(2.0))]);
        data.set_value("Data", 1, 1, Value::Blank);
        data.set_value("Data", 4, 2, Value::Number(7.0));
        assert_eq!(data.cell_value("Data", 1, 1), Value::Blank);
        assert_eq!(
            data.used_extent("Data"),
            Extent {
                rows: 4,
                columns: 2
            }
        );
        assert_eq!(evaluated(&data, "=SUM(A:B)"), Value::Number(7.0));
    }
}
