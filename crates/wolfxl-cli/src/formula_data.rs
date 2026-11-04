//! Cached workbook values shared by one-shot formulas and appended columns.

use std::collections::{BTreeMap, HashMap};

use anyhow::{anyhow, Result};
use wolfxl_core::{CellValue, Sheet};

use crate::commands::fingerprint::{date_serial, day_fraction};
use crate::eval::{self, DataSource, ExcelError, Extent, Value};
use crate::filter::lenient_number;
use crate::input::Input;

pub(crate) struct WorkbookData {
    sheets: BTreeMap<String, SheetData>,
}

struct SheetData {
    sheet: Sheet,
    extent: Extent,
    overrides: HashMap<(u32, u32), Value>,
}

impl WorkbookData {
    /// Load stored values only; existing formula cells retain their caches.
    pub(crate) fn load(input: &mut Input) -> Result<Self> {
        let names = input.sheet_names().to_vec();
        let delimited = input.is_delimited();
        let mut sheets = BTreeMap::new();
        for name in names {
            let sheet = input.sheet(&name)?;
            let (row, column) = sheet.origin();
            let width = sheet.rows().iter().map(Vec::len).max().unwrap_or(0);
            let extent = if width == 0 || sheet.rows().is_empty() {
                Extent::default()
            } else {
                Extent {
                    rows: row + sheet.rows().len() as u32,
                    columns: column + width as u32,
                }
            };
            let mut overrides = HashMap::new();
            if delimited {
                // Import money text once, without changing raw display fields
                // or the types of text cells stored in spreadsheet files.
                for (r, cells) in sheet.rows().iter().enumerate() {
                    for (c, cell) in cells.iter().enumerate() {
                        let CellValue::String(text) = &cell.value else { continue };
                        let Some(number) = lenient_number(text) else { continue };
                        overrides.insert(
                            (row + r as u32 + 1, column + c as u32 + 1),
                            Value::Number(number),
                        );
                    }
                }
            }
            sheets.insert(name, SheetData { sheet, extent, overrides });
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

impl DataSource for WorkbookData {
    fn sheet_names(&self) -> impl Iterator<Item = &str> {
        self.sheets.keys().map(String::as_str)
    }

    fn cell_value(&self, sheet: &str, row: u32, column: u32) -> Value {
        let Some(data) = self.sheets.get(sheet) else { return Value::Blank };
        if let Some(value) = data.overrides.get(&(row, column)) {
            return value.clone();
        }
        let (origin_row, origin_column) = data.sheet.origin();
        let Some(row) = row.checked_sub(origin_row + 1) else { return Value::Blank };
        let Some(column) = column.checked_sub(origin_column + 1) else { return Value::Blank };
        let Some(cell) = data.sheet.rows().get(row as usize).and_then(|row| row.get(column as usize)) else {
            return Value::Blank;
        };
        match &cell.value {
            CellValue::Empty => Value::Blank,
            CellValue::String(text) => Value::Text(text.clone()),
            CellValue::Int(number) => Value::Number(*number as f64),
            CellValue::Float(number) if number.is_finite() => Value::Number(*number),
            CellValue::Float(_) => Value::Error(ExcelError::Num),
            CellValue::Bool(value) => Value::Boolean(*value),
            CellValue::Date(date) => date_serial(*date).map_or(Value::Blank, Value::Number),
            CellValue::DateTime(moment) => date_serial(moment.date())
                .map_or(Value::Blank, |date| Value::Number(date + day_fraction(moment.time()))),
            CellValue::Time(time) => Value::Number(day_fraction(*time)),
            CellValue::Error(literal) => Value::Error(match literal.as_str() {
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

    fn used_extent(&self, sheet: &str) -> Extent {
        self.sheets.get(sheet).map_or(Extent::default(), |data| data.extent)
    }
}

pub(crate) fn normalize_formula(formula: &str) -> String {
    let formula = formula.trim();
    if formula.starts_with('=') { formula.to_owned() } else { format!("={formula}") }
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
            if number == 0.0 { "0".to_owned() } else { number.to_string() }
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
