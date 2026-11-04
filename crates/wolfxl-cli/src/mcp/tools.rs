use std::fs;
use std::path::Path;

use anyhow::{anyhow, bail, Context, Result};
use serde_json::{json, Value};
use wolfxl_core::{CellValue, Workbook};
use wolfxl_edit::{CellEdit, EditSession, EditValue};
use wolfxl_reader::CellValue as RawValue;

use super::files::revision;
use super::Server;
use crate::{
    commands, input, CalcArgs, DiffArgs, DiffFormat, KeyTransform, PivotArgs, PivotFormat,
    StatsArgs, StatsFormat,
};

const CELL_CAP: u64 = 10_000;

fn text(arguments: &Value, field: &str) -> String {
    arguments[field].as_str().unwrap_or_default().to_owned()
}

fn optional_text(arguments: &Value, field: &str) -> Option<String> {
    arguments[field].as_str().map(str::to_owned)
}

fn strings(arguments: &Value, field: &str) -> Vec<String> {
    arguments[field]
        .as_array()
        .map(|values| {
            values
                .iter()
                .filter_map(Value::as_str)
                .map(str::to_owned)
                .collect()
        })
        .unwrap_or_default()
}

impl Server {
    pub(super) fn call(&mut self, name: &str, arguments: &Value) -> Result<Value> {
        let path = self.path(arguments, "path")?;
        let sheet = optional_text(arguments, "sheet");
        match name {
            "get_workbook" => get_workbook(&path),
            "read_range" => read_range(&path, sheet.as_deref(), &text(arguments, "range")),
            "calc" => commands::calc::structured(&CalcArgs {
                formula: text(arguments, "formula"),
                file: Some(path),
                sheet,
                from: None,
                json: true,
            }),
            "stats" => commands::stats::structured(&StatsArgs {
                file: path,
                sheet,
                header_row: None,
                from: None,
                filters: strings(arguments, "filters"),
                select: strings(arguments, "select"),
                top: arguments["top"].as_u64().unwrap_or(5) as usize,
                export: StatsFormat::Json,
            }),
            "pivot" => commands::pivot::structured(&PivotArgs {
                file: path,
                sheet,
                header_row: None,
                from: None,
                rows: strings(arguments, "rows"),
                column: optional_text(arguments, "column"),
                values: strings(arguments, "values"),
                export: PivotFormat::Json,
            }),
            "diff" => {
                let right = self.path(arguments, "right_path")?;
                let key_transform = match arguments["key_transform"].as_str().unwrap_or("none") {
                    "trim" => KeyTransform::Trim,
                    "lower" => KeyTransform::Lower,
                    "digits" => KeyTransform::Digits,
                    _ => KeyTransform::None,
                };
                commands::diff::structured(&DiffArgs {
                    left: path,
                    right,
                    key: text(arguments, "key"),
                    compare: strings(arguments, "compare"),
                    tolerance: arguments["tolerance"].as_f64().unwrap_or(0.0),
                    key_transform,
                    sheet,
                    from: None,
                    format: DiffFormat::Json,
                    strict_exit: false,
                    quiet: false,
                })
            }
            "fingerprint" => commands::fingerprint::structured(&path),
            "undo" | "redo" => {
                editable(&path)?;
                self.restore(&path, arguments, name == "redo")
            }
            "write_cells" | "set_number_format" | "insert_rows" | "delete_rows"
            | "insert_columns" | "delete_columns" | "add_sheet" | "rename_sheet" => {
                self.edit(&path, name, arguments)
            }
            _ => bail!("unknown tool {name}"),
        }
    }

    fn edit(&mut self, path: &Path, name: &str, arguments: &Value) -> Result<Value> {
        editable(path)?;
        let before = fs::read(path).context("reading workbook before edit")?;
        Self::check_revision(arguments, &revision(&before))?;
        let mut session =
            EditSession::from_bytes(before.clone()).context("opening workbook for editing")?;
        let sheet = text(arguments, "sheet");
        let change = match name {
            "write_cells" => {
                let cells = arguments["cells"]
                    .as_array()
                    .ok_or_else(|| anyhow!("cells must be an array"))?
                    .iter()
                    .map(|cell| {
                        let value = if let Some(formula) = cell["formula"].as_str() {
                            EditValue::Formula(formula.to_owned())
                        } else {
                            match &cell["value"] {
                                Value::Null => EditValue::Clear,
                                Value::Bool(value) => EditValue::Bool(*value),
                                Value::Number(value) => EditValue::Number(
                                    value
                                        .as_f64()
                                        .ok_or_else(|| anyhow!("cell number must be finite"))?,
                                ),
                                Value::String(value) => EditValue::Text(value.clone()),
                                _ => bail!("cell value must be a number, string, boolean, or null"),
                            }
                        };
                        Ok(CellEdit {
                            cell: text(cell, "cell"),
                            value,
                        })
                    })
                    .collect::<Result<Vec<_>>>()?;
                session.write_cells(&sheet, &cells)?;
                json!({"operation":name,"sheet":sheet,"cells":arguments["cells"]})
            }
            "set_number_format" => {
                let range = text(arguments, "range");
                let format_code = text(arguments, "format_code");
                session.set_number_format(&sheet, &range, &format_code)?;
                json!({"operation":name,"sheet":sheet,"range":range,"format_code":format_code})
            }
            "insert_rows" | "delete_rows" | "insert_columns" | "delete_columns" => {
                let at = arguments["at"]
                    .as_u64()
                    .ok_or_else(|| anyhow!("at must be a positive 1-based index"))?
                    as u32;
                let count = arguments["count"]
                    .as_u64()
                    .ok_or_else(|| anyhow!("count must be a positive integer"))?
                    as u32;
                match name {
                    "insert_rows" => session.insert_rows(&sheet, at, count)?,
                    "delete_rows" => session.delete_rows(&sheet, at, count)?,
                    "insert_columns" => session.insert_columns(&sheet, at, count)?,
                    _ => session.delete_columns(&sheet, at, count)?,
                }
                json!({"operation":name,"sheet":sheet,"at":at,"count":count})
            }
            "add_sheet" => {
                let new = text(arguments, "name");
                let position = arguments["position"]
                    .as_u64()
                    .map(|position| position as usize);
                session.add_sheet(&new, position)?;
                json!({"operation":name,"name":new,"position":position})
            }
            "rename_sheet" => {
                let old = text(arguments, "old");
                let new = text(arguments, "new");
                session.rename_sheet(&old, &new)?;
                json!({"operation":name,"old":old,"new":new})
            }
            _ => bail!("unknown editing tool {name}"),
        };
        let after = session.to_bytes().context("rendering edited workbook")?;
        self.commit_edit(path, before, after, arguments, change)
    }
}

fn editable(path: &Path) -> Result<()> {
    let extension = path
        .extension()
        .and_then(|ext| ext.to_str())
        .unwrap_or_default();
    if !extension.eq_ignore_ascii_case("xlsx") && !extension.eq_ignore_ascii_case("xlsm") {
        bail!("edits require an .xlsx or .xlsm workbook");
    }
    Ok(())
}

fn get_workbook(path: &Path) -> Result<Value> {
    let bytes = fs::read(path).context("reading workbook revision")?;
    let mut workbook = Workbook::open(path).context("opening workbook")?;
    let map = workbook.map().context("mapping workbook")?;
    Ok(
        json!({"path":path,"format":path.extension().and_then(|extension|extension.to_str()).unwrap_or_default().to_ascii_lowercase(),"revision":revision(&bytes),"sheets":map.sheets.iter().map(|sheet|json!({"name":sheet.name,"rows":sheet.rows,"cols":sheet.cols,"headers":sheet.headers,"class":sheet.class.as_str(),"tables":sheet.tables})).collect::<Vec<_>>(),"named_ranges":map.named_ranges.iter().map(|(name,formula)|json!({"name":name,"formula":formula})).collect::<Vec<_>>()}),
    )
}

fn read_range(path: &Path, sheet: Option<&str>, range: &str) -> Result<Value> {
    let ((top, left), (bottom, right)) = rectangle(range)?;
    let count = u64::from(bottom - top + 1) * u64::from(right - left + 1);
    if count > CELL_CAP {
        bail!("range {range:?} requests {count} cells; maximum is {CELL_CAP}. Page with adjacent smaller row or column ranges, for example A1:J1000 then A1001:J2000");
    }
    let bytes = fs::read(path).context("reading workbook revision")?;
    let mut input = input::open(Some(path), None)?;
    let name = input.sheet_name(sheet)?;
    let height = (bottom - top + 1) as usize;
    let width = (right - left + 1) as usize;
    let mut values = vec![vec![Value::Null; width]; height];
    let mut formulas = vec![vec![Value::Null; width]; height];
    if let Some(book) = input.raw_book()? {
        for cell in book.cells(&name)? {
            if cell.row < top || cell.row > bottom || cell.col < left || cell.col > right {
                continue;
            }
            let row = (cell.row - top) as usize;
            let col = (cell.col - left) as usize;
            values[row][col] = match cell.value {
                RawValue::Empty => Value::Null,
                RawValue::Number(number) => json!(number),
                RawValue::Bool(boolean) => json!(boolean),
                RawValue::String(text) | RawValue::Error(text) => json!(text),
                RawValue::SharedString(text) => json!(&*text),
            };
            formulas[row][col] = cell.formula.map_or(Value::Null, Value::String);
        }
    } else {
        let sheet = input.raw_sheet(&name)?;
        let (row_origin, col_origin) = sheet.origin();
        for (row, fields) in sheet.rows().iter().enumerate() {
            let absolute_row = row_origin + row as u32 + 1;
            if absolute_row < top || absolute_row > bottom {
                continue;
            }
            for (col, cell) in fields.iter().enumerate() {
                let absolute_col = col_origin + col as u32 + 1;
                if absolute_col < left || absolute_col > right {
                    continue;
                }
                values[(absolute_row - top) as usize][(absolute_col - left) as usize] =
                    match &cell.value {
                        CellValue::Empty => Value::Null,
                        CellValue::String(text) | CellValue::Error(text) => json!(text),
                        CellValue::Int(number) => json!(number),
                        CellValue::Float(number) => json!(number),
                        CellValue::Bool(boolean) => json!(boolean),
                        CellValue::Date(date) => json!(date.format("%Y-%m-%d").to_string()),
                        CellValue::DateTime(datetime) => {
                            json!(datetime.format("%Y-%m-%dT%H:%M:%S").to_string())
                        }
                        CellValue::Time(time) => json!(time.format("%H:%M:%S").to_string()),
                    };
            }
        }
    }
    Ok(
        json!({"path":path,"sheet":name,"range":range,"revision":revision(&bytes),"values":values,"formulas":formulas,"cell_cap":CELL_CAP}),
    )
}

fn coordinate(reference: &str) -> Result<(u32, u32)> {
    let reference = reference.as_bytes();
    let mut index = 0;
    if reference.get(index) == Some(&b'$') {
        index += 1;
    }
    let mut col = 0u32;
    let start = index;
    while let Some(byte) = reference
        .get(index)
        .filter(|byte| byte.is_ascii_alphabetic())
    {
        col = col
            .checked_mul(26)
            .and_then(|col| col.checked_add(u32::from(byte.to_ascii_uppercase() - b'A' + 1)))
            .ok_or_else(|| anyhow!("invalid A1 column"))?;
        index += 1;
    }
    if index == start {
        bail!("expected A1 cell reference");
    }
    if reference.get(index) == Some(&b'$') {
        index += 1;
    }
    let digits = reference
        .get(index..)
        .ok_or_else(|| anyhow!("expected A1 row"))?;
    if digits.is_empty() || !digits.iter().all(u8::is_ascii_digit) || digits[0] == b'0' {
        bail!("expected a positive A1 row number");
    }
    let row = std::str::from_utf8(digits)?
        .parse::<u32>()
        .context("invalid A1 row")?;
    if row > 1_048_576 || col > 16_384 {
        bail!("A1 reference exceeds Excel limits (1048576 rows, 16384 columns)");
    }
    Ok((row, col))
}

fn rectangle(range: &str) -> Result<((u32, u32), (u32, u32))> {
    let mut parts = range.split(':');
    let first = coordinate(parts.next().unwrap_or_default())
        .with_context(|| format!("invalid range {range:?}"))?;
    let last = parts.next().map(coordinate).transpose()?.unwrap_or(first);
    if parts.next().is_some() || last.0 < first.0 || last.1 < first.1 {
        bail!("invalid range {range:?}: expected top-left:bottom-right A1 rectangle");
    }
    Ok((first, last))
}
