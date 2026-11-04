//! Cached workbook values shared by one-shot formulas and appended columns.

use std::collections::{BTreeMap, HashMap, HashSet};

use anyhow::{anyhow, Result};
use wolfxl_core::CellValue;
use wolfxl_reader::{CellDataType, CellValue as RawValue, WorksheetData};

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
    /// Rows the sheet marks hidden. Only the XLSX and XLSB readers report
    /// them, and neither tells a manually hidden row from a filtered one.
    hidden_rows: HashSet<u32>,
    /// Cells whose stored formula calls `SUBTOTAL`.
    subtotals: HashSet<(u32, u32)>,
    extent: Extent,
    /// The worksheet tables on this sheet, which structured references in
    /// any sheet's formulas address by name.
    tables: Vec<SheetTable>,
}

/// A worksheet table as its table part describes it.
struct SheetTable {
    name: String,
    area: eval::CellArea,
    header_rows: u32,
    totals_rows: u32,
    columns: Vec<String>,
}

impl SheetTable {
    /// `None` when the table part's `ref` is not a cell range, so references
    /// to the table evaluate to `#REF!`.
    fn new(table: wolfxl_reader::Table) -> Option<Self> {
        Some(Self {
            area: eval::cell_area(&table.ref_range)?,
            header_rows: u32::from(table.header_row),
            totals_rows: u32::from(table.totals_row),
            name: table.name,
            columns: table.columns,
        })
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
                Some(book) => stored_sheet(book.worksheet(&name)?),
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

/// A spreadsheet's stored cells; the extent covers every cell the reader kept
/// with a value or a formula, formulas without a cached value included.
fn stored_sheet(sheet: WorksheetData) -> SheetData {
    let mut data = SheetData {
        hidden_rows: sheet.hidden_rows.into_iter().collect(),
        ..SheetData::default()
    };
    for cell in sheet.cells {
        if cell.formula.is_none() && matches!(cell.value, RawValue::Empty) {
            continue;
        }
        data.extent.rows = data.extent.rows.max(cell.row);
        data.extent.columns = data.extent.columns.max(cell.col);
        if cell.formula.as_deref().is_some_and(calls_subtotal) {
            data.subtotals.insert((cell.row, cell.col));
        }
        let value = stored_value(cell.value, cell.data_type);
        if !matches!(value, Value::Blank) {
            data.cells.insert((cell.row, cell.col), value);
        }
    }
    data.tables = sheet
        .tables
        .into_iter()
        .filter_map(SheetTable::new)
        .collect();
    data
}

/// Whether formula text calls `SUBTOTAL` anywhere outside string literals and
/// quoted sheet names, as Excel checks to keep subtotals from double counting.
fn calls_subtotal(formula: &str) -> bool {
    const NAME: &[u8] = b"SUBTOTAL";
    let bytes = formula.as_bytes();
    let mut quote = None;
    let mut index = 0;
    while index < bytes.len() {
        let byte = bytes[index];
        match quote {
            // A doubled quote inside a quoted run is an escaped quote and
            // reopens the run on the next byte.
            Some(open) if byte == open => quote = None,
            Some(_) => {}
            None if byte == b'"' || byte == b'\'' => quote = Some(byte),
            None => {
                let boundary = index == 0 || {
                    let before = bytes[index - 1];
                    !(before.is_ascii_alphanumeric() || matches!(before, b'_' | b'.'))
                };
                let named = bytes
                    .get(index..index + NAME.len())
                    .is_some_and(|candidate| candidate.eq_ignore_ascii_case(NAME));
                if boundary && named {
                    let rest = &formula[index + NAME.len()..];
                    if rest.trim_start().starts_with('(') {
                        return true;
                    }
                }
            }
        }
        index += 1;
    }
    false
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

    fn row_hidden(&self, sheet: &str, row: u32) -> bool {
        self.sheets
            .get(sheet)
            .is_some_and(|data| data.hidden_rows.contains(&row))
    }

    fn is_subtotal(&self, sheet: &str, row: u32, column: u32) -> bool {
        self.sheets
            .get(sheet)
            .is_some_and(|data| data.subtotals.contains(&(row, column)))
    }

    fn table(&self, name: &str) -> Option<eval::TableInfo<'_>> {
        self.sheets.iter().find_map(|(sheet, data)| {
            let table = data
                .tables
                .iter()
                .find(|table| eval::names_match(&table.name, name))?;
            Some(eval::TableInfo {
                sheet,
                area: table.area,
                header_rows: table.header_rows,
                totals_rows: table.totals_rows,
                columns: &table.columns,
            })
        })
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
    use wolfxl_reader::Cell as RawCell;

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
        sheet_book(WorksheetData {
            cells,
            ..WorksheetData::default()
        })
    }

    fn sheet_book(sheet: WorksheetData) -> WorkbookData {
        WorkbookData {
            sheets: BTreeMap::from([("Data".to_owned(), stored_sheet(sheet))]),
            date1904: false,
        }
    }

    fn evaluated(data: &WorkbookData, formula: &str) -> Value {
        eval::evaluate(formula, data, "Data").expect("formula parses")
    }

    fn reader_table(name: &str, ref_range: &str, columns: &[&str]) -> wolfxl_reader::Table {
        wolfxl_reader::Table {
            name: name.to_owned(),
            ref_range: ref_range.to_owned(),
            header_row: true,
            totals_row: false,
            comment: None,
            table_type: None,
            totals_row_shown: None,
            style: None,
            show_first_column: false,
            show_last_column: false,
            show_row_stripes: false,
            show_column_stripes: false,
            columns: columns.iter().map(|column| (*column).to_owned()).collect(),
            auto_filter: None,
        }
    }

    #[test]
    fn table_parts_register_structured_reference_targets() {
        let data = sheet_book(WorksheetData {
            cells: vec![
                raw_cell(3, 3, RawValue::Number(5.0)),
                raw_cell(4, 3, RawValue::Number(7.0)),
            ],
            tables: vec![
                reader_table("Sales", "$B$2:$C$4", &["Item", "Amount"]),
                reader_table("Broken", "not a range", &["Amount"]),
            ],
            ..WorksheetData::default()
        });
        assert_eq!(evaluated(&data, "=SUM(sales[AMOUNT])"), Value::Number(12.0));
        assert_eq!(
            evaluated(&data, "=SUM(Broken[Amount])"),
            Value::Error(ExcelError::Ref)
        );
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

    #[test]
    fn subtotal_formulas_are_recognized_outside_quotes() {
        for formula in [
            "SUBTOTAL(109,BigCity[Population])",
            "=subtotal (9,A1:A3)*2",
            "ROUND(SUBTOTAL(1,A:A),2)",
            "\"x\"&SUBTOTAL(9,A1)",
        ] {
            assert!(calls_subtotal(formula), "{formula}");
        }
        for formula in [
            "SUM(A1:A3)",
            "\"SUBTOTAL(9,A1)\"",
            "'SUBTOTAL(1'!A1",
            "MYSUBTOTAL(1,A1)",
            "SUBTOTALS",
        ] {
            assert!(!calls_subtotal(formula), "{formula}");
        }
    }

    #[test]
    fn subtotal_skips_hidden_rows_and_nested_subtotals() {
        let data = sheet_book(WorksheetData {
            cells: vec![
                raw_cell(1, 1, RawValue::Number(10.0)),
                raw_cell(2, 1, RawValue::Number(20.0)),
                raw_cell(3, 1, RawValue::Number(30.0)),
                RawCell {
                    formula: Some("SUBTOTAL(9,A1:A3)".to_owned()),
                    ..raw_cell(4, 1, RawValue::Number(60.0))
                },
            ],
            hidden_rows: vec![2],
            ..WorksheetData::default()
        });
        assert!(data.row_hidden("Data", 2));
        assert!(!data.row_hidden("Data", 1));
        assert!(data.is_subtotal("Data", 4, 1));
        assert_eq!(evaluated(&data, "=SUM(A1:A4)"), Value::Number(120.0));
        assert_eq!(evaluated(&data, "=SUBTOTAL(9,A1:A4)"), Value::Number(60.0));
        assert_eq!(evaluated(&data, "=SUBTOTAL(109,A:A)"), Value::Number(40.0));
        assert_eq!(evaluated(&data, "=SUBTOTAL(102,A1:A4)"), Value::Number(2.0));
    }
}
