//! Formula columns for `wolfxl peek --add NAME=FORMULA`.
//!
//! `NAME=FORMULA` splits at its first `=`, so `Margin==C2-B2` and
//! `Margin=C2-B2` name the same formula. The formula is written for the
//! first data row, the row under the header, and filled down the way Excel
//! fills a formula: each data row shifts the relative row parts of every
//! reference by its distance from the first data row, `$`-anchored rows stay,
//! and columns never shift. Whole-column references, string literals and
//! function names are left as written. Coordinates are the sheet's own, so a
//! block that starts below row 1 or right of column A keeps its addresses.
//!
//! Added columns take the free column letters after the sheet's populated
//! block, in `--add` order. Every formula is evaluated against the stored
//! values of every sheet (typed delimited fields, cached results for formula
//! cells), as `wolfxl calc` does, plus the results computed so far: every
//! earlier added column and the rows above in its own column. Later added
//! columns and lower rows of its own column read as blank.
//!
//! Results read as `wolfxl calc` prints them. An Excel error stays an error
//! value in its cell, and a result that is a range rather than one value is
//! `#VALUE!`: the evaluator never implicitly intersects.

use anyhow::{anyhow, bail, Context, Result};
use wolfxl_core::{Cell, CellValue, Sheet};
use wolfxl_formula::{translate, RefDelta, TokenizeError, MAX_COL};

use crate::commands::peek::detached;
use crate::eval::{ExcelError, Formula, Value};
use crate::formula_data::{normalize_formula, parse_formula, text_value, WorkbookData};
use crate::input::{Input, Provenance};

/// Why formula columns could not be added.
pub(crate) enum Failure {
    /// The `--add` request is unusable on this sheet: a malformed or
    /// unsupported formula, a column name already taken, or no free column.
    Usage(anyhow::Error),
    /// Reading the workbook failed.
    Input(anyhow::Error),
}

/// Every `--add NAME=FORMULA`, parsed once and appendable to any sheet.
pub(crate) struct Columns(Vec<Column>);

/// One parsed `--add NAME=FORMULA`.
struct Column {
    /// The argument as given, quoted in errors.
    spec: String,
    name: String,
    /// The formula for the first data row, with a leading `=`.
    formula: String,
    /// `formula`, parsed.
    parsed: Formula,
}

impl Columns {
    /// Parse every spec; the first malformed or unsupported one is the error.
    pub(crate) fn parse(specs: &[String]) -> Result<Self> {
        specs
            .iter()
            .map(String::as_str)
            .map(Column::parse)
            .collect::<Result<Vec<_>>>()
            .map(Self)
    }

    /// `raw` (sheet `target` of `input`, in source spelling) with one column
    /// appended per spec: the header row gets NAME, every data row through
    /// the last non-blank one gets the formula's result.
    pub(crate) fn append(
        &self,
        input: &mut Input,
        target: &str,
        raw: &Sheet,
        provenance: &Provenance,
    ) -> Result<Sheet, Failure> {
        let columns = &self.0;
        let headers = raw.headers();
        for (index, column) in columns.iter().enumerate() {
            let wanted = column.name.to_lowercase();
            let taken = headers
                .iter()
                .map(String::as_str)
                .chain(columns[..index].iter().map(|earlier| earlier.name.as_str()))
                .any(|name| name.to_lowercase() == wanted);
            if taken {
                return Err(Failure::Usage(anyhow!(
                    "--add {:?}: a column named {:?} already exists",
                    column.spec,
                    column.name
                )));
            }
        }

        let width = raw.rows().iter().map(Vec::len).max().unwrap_or(0);
        let origin_row = provenance.rows.first().copied().unwrap_or(1) - 1;
        let origin_column = provenance.columns.first().copied().flatten().unwrap_or(1) - 1;
        // 0-based sheet column of the first added column.
        let first_free = u64::from(origin_column) + width as u64;
        if first_free + columns.len() as u64 > u64::from(MAX_COL) {
            return Err(Failure::Usage(anyhow!(
                "--add: sheet {target:?} has no free column for {} more column(s)",
                columns.len()
            )));
        }
        // Data rows are block rows 1..=last_row; trailing blank rows get no
        // formula.
        let last_row = raw
            .rows()
            .iter()
            .rposition(|row| row.iter().any(|cell| !cell.value.is_empty()))
            .unwrap_or(0);

        let total = width + columns.len();
        let mut rows: Vec<Vec<Cell>> = raw
            .rows()
            .iter()
            .map(|row| {
                let mut cells = Vec::with_capacity(total);
                cells.extend(row.iter().map(|cell| detached(raw, cell)));
                cells.resize_with(total, Cell::empty);
                cells
            })
            .collect();
        if rows.is_empty() {
            rows.push(vec![Cell::empty(); total]);
        }
        for (cell, column) in rows[0].iter_mut().skip(width).zip(columns) {
            cell.value = CellValue::String(column.name.clone());
        }

        if last_row > 0 {
            let mut data = WorkbookData::load(input).map_err(Failure::Input)?;
            // 1-based sheet row of the first data row, one under the header.
            let first_row = origin_row + 2;
            for (index, column) in columns.iter().enumerate() {
                // Fits: checked against MAX_COL above.
                let sheet_column = (first_free + index as u64 + 1) as u32;
                for offset in 0..last_row {
                    let shifted;
                    let formula = if offset == 0 {
                        &column.parsed
                    } else {
                        shifted = column.filled(offset)?;
                        &shifted
                    };
                    let value = match formula.evaluate(&data, target) {
                        Value::Array(_) => Value::Error(ExcelError::Value),
                        value => value,
                    };
                    rows[offset + 1][width + index].value = cell_value(&value);
                    data.set_value(target, first_row + offset as u32, sheet_column, value);
                }
            }
        }
        Ok(Sheet::from_rows(raw.name.clone(), rows))
    }
}

impl Column {
    fn parse(spec: &str) -> Result<Self> {
        let Some((name, formula)) = spec.split_once('=') else {
            bail!("--add {spec:?}: expected NAME=FORMULA");
        };
        let name = name.trim();
        if name.is_empty() {
            bail!("--add {spec:?}: missing the column name before \"=\"");
        }
        if formula.trim().trim_start_matches('=').trim().is_empty() {
            bail!("--add {spec:?}: missing the formula after \"=\"");
        }
        let formula = normalize_formula(formula);
        let parsed = parse_formula(&formula).with_context(|| format!("--add {spec:?}"))?;
        // Filling down re-tokenizes the formula for every later row.
        wolfxl_formula::tokenize(&formula)
            .map_err(|error| anyhow!("--add {spec:?}: cannot parse {formula}: {error}"))?;
        Ok(Self {
            spec: spec.to_owned(),
            name: name.to_owned(),
            formula,
            parsed,
        })
    }

    /// The formula for the data row `offset` rows below the first, parsed.
    fn filled(&self, offset: usize) -> Result<Formula, Failure> {
        let (spec, first) = (&self.spec, &self.formula);
        let usage = |error: anyhow::Error| Failure::Usage(error.context(format!("--add {spec:?}")));
        // A sheet has at most 1,048,576 rows, so the offset fits.
        let rows = i32::try_from(offset).map_err(|_| usage(anyhow!("too many rows to fill")))?;
        let formula = fill_down(first, rows)
            .map_err(|error| usage(anyhow!("cannot parse {first}: {error}")))?;
        parse_formula(&formula).map_err(usage)
    }
}

/// `formula` (written for the first data row) as filled `rows` rows down:
/// relative row parts shift, `$` rows and every column stay.
fn fill_down(formula: &str, rows: i32) -> Result<String, TokenizeError> {
    let mut delta = RefDelta::empty();
    delta.rows = rows;
    delta.anchor_row = 1;
    delta.respect_dollar = true;
    translate(formula, &delta)
}

/// A result as a cell, spelled the way `wolfxl calc` prints it: numbers with
/// up to 15 significant digits, booleans as `TRUE`/`FALSE`.
fn cell_value(value: &Value) -> CellValue {
    match value {
        Value::Blank => CellValue::Empty,
        Value::Text(text) => CellValue::String(text.clone()),
        Value::Error(error) => CellValue::Error(error.as_str().to_owned()),
        Value::Number(number) if !number.is_finite() => {
            CellValue::Error(ExcelError::Num.as_str().to_owned())
        }
        Value::Array(_) => CellValue::Error(ExcelError::Value.as_str().to_owned()),
        value => CellValue::String(text_value(value)),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn fill_down_shifts_relative_rows_only() {
        for (formula, rows, expected) in [
            ("=C2-B2", 3, "=C5-B5"),
            ("=B2*$C$2+B$2+$B2", 2, "=B4*$C$2+B$2+$B4"),
            ("=SUM($B$2:B2)", 2, "=SUM($B$2:B4)"),
            ("=SUM(B:B)-B2", 1, "=SUM(B:B)-B3"),
            ("=Rates!B2+'Q1 Rates'!$B$2", 1, "=Rates!B3+'Q1 Rates'!$B$2"),
            ("=\"B2\"&A2", 1, "=\"B2\"&A3"),
            ("=IF(B2>0,B2,\"-\")", 1, "=IF(B3>0,B3,\"-\")"),
        ] {
            let filled = fill_down(formula, rows).expect("tokenizes");
            assert_eq!(filled, expected, "{formula} + {rows}");
        }
    }

    #[test]
    fn specs_split_at_the_first_equals_sign() {
        for spec in ["Margin=B2-C2", "Margin==B2-C2", " Margin = B2-C2 "] {
            let column = Column::parse(spec).unwrap_or_else(|error| panic!("{spec}: {error:#}"));
            assert_eq!(column.name, "Margin");
            assert_eq!(column.formula, "=B2-C2");
        }
    }

    #[test]
    fn results_read_like_calc() {
        assert!(matches!(cell_value(&Value::Blank), CellValue::Empty));
        assert!(matches!(
            cell_value(&Value::Number(0.1 + 0.2)),
            CellValue::String(text) if text == "0.3"
        ));
        assert!(matches!(
            cell_value(&Value::Boolean(true)),
            CellValue::String(text) if text == "TRUE"
        ));
        assert!(matches!(
            cell_value(&Value::Error(ExcelError::Div0)),
            CellValue::Error(text) if text == "#DIV/0!"
        ));
        assert!(matches!(
            cell_value(&Value::Array(vec![vec![Value::Number(1.0)]])),
            CellValue::Error(text) if text == "#VALUE!"
        ));
    }
}
