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
}

/// One sheet's values in a compact grid whose top-left cell sits at the
/// zero-based `origin`, plus values set after loading.
struct SheetData {
    origin: (u32, u32),
    rows: Vec<Vec<Value>>,
    extent: Extent,
    overrides: HashMap<(u32, u32), Value>,
}

impl SheetData {
    fn new(origin: (u32, u32), rows: Vec<Vec<Value>>) -> Self {
        let width = rows.iter().map(Vec::len).max().unwrap_or(0);
        let extent = if width == 0 {
            Extent::default()
        } else {
            Extent {
                rows: origin.0 + rows.len() as u32,
                columns: origin.1 + width as u32,
            }
        };
        Self {
            origin,
            rows,
            extent,
            overrides: HashMap::new(),
        }
    }
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
                    let rows = sheet
                        .rows()
                        .iter()
                        .map(|row| {
                            row.iter()
                                .map(|cell| match &cell.value {
                                    CellValue::String(text) => delimited_field(text),
                                    _ => Value::Blank,
                                })
                                .collect()
                        })
                        .collect();
                    SheetData::new(sheet.origin(), rows)
                }
            };
            sheets.insert(name, data);
        }
        Ok(Self { sheets })
    }

    /// Coordinates are one-based, like the evaluator's references.
    pub(crate) fn set_value(&mut self, sheet: &str, row: u32, column: u32, value: Value) {
        let data = self.sheets.get_mut(sheet).expect("loaded target sheet");
        data.extent.rows = data.extent.rows.max(row);
        data.extent.columns = data.extent.columns.max(column);
        data.overrides.insert((row, column), value);
    }
}

/// The compact grid of a spreadsheet's row-major stored cells.
fn stored_sheet(cells: Vec<RawCell>) -> SheetData {
    let (Some(first), Some(last)) = (cells.first(), cells.last()) else {
        return SheetData::new((0, 0), Vec::new());
    };
    let (top, bottom) = (first.row, last.row);
    let left = cells.iter().map(|cell| cell.col).min().unwrap_or(first.col);
    let right = cells.iter().map(|cell| cell.col).max().unwrap_or(first.col);
    let mut rows =
        vec![vec![Value::Blank; (right - left + 1) as usize]; (bottom - top + 1) as usize];
    for cell in cells {
        rows[(cell.row - top) as usize][(cell.col - left) as usize] =
            stored_value(cell.value, cell.data_type);
    }
    SheetData::new((top - 1, left - 1), rows)
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
        let Some(data) = self.sheets.get(sheet) else {
            return Value::Blank;
        };
        if let Some(value) = data.overrides.get(&(row, column)) {
            return value.clone();
        }
        let (origin_row, origin_column) = data.origin;
        let Some(row) = row.checked_sub(origin_row + 1) else {
            return Value::Blank;
        };
        let Some(column) = column.checked_sub(origin_column + 1) else {
            return Value::Blank;
        };
        data.rows
            .get(row as usize)
            .and_then(|row| row.get(column as usize))
            .cloned()
            .unwrap_or(Value::Blank)
    }

    fn used_extent(&self, sheet: &str) -> Extent {
        self.sheets
            .get(sheet)
            .map_or(Extent::default(), |data| data.extent)
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
}
