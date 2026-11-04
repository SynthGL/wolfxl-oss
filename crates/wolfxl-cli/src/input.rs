//! Spreadsheet input for commands that read a file path or standard input.
//!
//! `wolfxl_core::Workbook` dispatches on the file extension, so piped bytes
//! and paths whose extension disagrees with `--from` are spooled into a
//! private temporary directory under a name carrying the requested
//! extension. The spool lives as long as the [`Input`] and is removed on drop.
//!
//! Delimited sources (CSV, TSV, TXT) arrive from the core reader as text.
//! [`Input::sheet`] types their fields the way Excel's text import does:
//! decimal numbers become numbers and `TRUE`/`FALSE` become booleans, so
//! aggregations and profiles see values rather than strings.

use std::fs;
use std::io::Read;
use std::path::{Path, PathBuf};
use std::time::{SystemTime, UNIX_EPOCH};

use anyhow::{anyhow, bail, Context, Result};
use clap::ValueEnum;
use wolfxl_core::{Cell, CellValue, Sheet, SourceFormat, Workbook};

/// Input format named by `--from`.
#[derive(Copy, Clone, Debug, PartialEq, Eq, ValueEnum)]
pub(crate) enum InputFormat {
    Csv,
    Tsv,
    Xlsx,
    Xlsm,
    Xls,
    Xlsb,
    Ods,
}

impl InputFormat {
    fn extension(self) -> &'static str {
        match self {
            InputFormat::Csv => "csv",
            InputFormat::Tsv => "tsv",
            InputFormat::Xlsx => "xlsx",
            InputFormat::Xlsm => "xlsm",
            InputFormat::Xls => "xls",
            InputFormat::Xlsb => "xlsb",
            InputFormat::Ods => "ods",
        }
    }

    /// The format a path's extension already selects, if `--from` names it.
    fn of_path(path: &Path) -> Option<Self> {
        let extension = path.extension()?.to_str()?.to_ascii_lowercase();
        Some(match extension.as_str() {
            "csv" | "txt" => InputFormat::Csv,
            "tsv" => InputFormat::Tsv,
            "xlsx" => InputFormat::Xlsx,
            "xlsm" => InputFormat::Xlsm,
            "xls" => InputFormat::Xls,
            "xlsb" => InputFormat::Xlsb,
            "ods" => InputFormat::Ods,
            _ => return None,
        })
    }
}

/// An opened spreadsheet input.
pub(crate) struct Input {
    workbook: Workbook,
    delimited: bool,
    // Declared after `workbook` so the backing file outlives the reader.
    _spool: Option<Spool>,
}

/// Open `file`, or standard input when `file` is `None` or `-`.
///
/// Standard input requires `from`. For a path, `from` overrides the
/// extension.
pub(crate) fn open(file: Option<&Path>, from: Option<InputFormat>) -> Result<Input> {
    let path = file.filter(|path| *path != Path::new("-"));
    let Some(path) = path else {
        let format = from.ok_or_else(|| {
            anyhow!("reading standard input requires --from csv|tsv|xlsx|xlsm|xls|xlsb|ods")
        })?;
        let mut bytes = Vec::new();
        std::io::stdin()
            .read_to_end(&mut bytes)
            .context("reading standard input")?;
        return spooled(&bytes, format, "stdin");
    };
    match from {
        Some(format) if InputFormat::of_path(path) != Some(format) => {
            let bytes = fs::read(path).with_context(|| format!("reading {}", path.display()))?;
            let stem = path
                .file_stem()
                .and_then(|stem| stem.to_str())
                .unwrap_or("Sheet1");
            spooled(&bytes, format, stem)
        }
        _ => {
            let workbook =
                Workbook::open(path).with_context(|| format!("opening {}", path.display()))?;
            Ok(Input::new(workbook, None))
        }
    }
}

impl Input {
    fn new(workbook: Workbook, spool: Option<Spool>) -> Self {
        let delimited = workbook.format() == SourceFormat::Csv;
        Self {
            workbook,
            delimited,
            _spool: spool,
        }
    }

    pub(crate) fn sheet_names(&self) -> &[String] {
        self.workbook.sheet_names()
    }

    /// Resolve `--sheet`: the named sheet, or the first one.
    pub(crate) fn sheet_name(&self, requested: Option<&str>) -> Result<String> {
        let names = self.sheet_names();
        match requested {
            Some(name) if names.iter().any(|candidate| candidate == name) => Ok(name.to_string()),
            Some(name) => bail!("no sheet named {name:?}; sheets: {}", names.join(", ")),
            None => names
                .first()
                .cloned()
                .ok_or_else(|| anyhow!("workbook has no sheets")),
        }
    }

    /// Load sheet `name` as the reader returns it: delimited fields stay text,
    /// so rendering and key matching see the source spelling (`007`, `1,200`).
    pub(crate) fn raw_sheet(&mut self, name: &str) -> Result<Sheet> {
        self.workbook
            .sheet(name)
            .with_context(|| format!("loading sheet {name:?}"))
    }

    /// Load sheet `name`, typing delimited fields.
    pub(crate) fn sheet(&mut self, name: &str) -> Result<Sheet> {
        let sheet = self.raw_sheet(name)?;
        if !self.delimited {
            return Ok(sheet);
        }
        let rows = sheet
            .rows()
            .iter()
            .map(|row| {
                row.iter()
                    .map(|cell| match &cell.value {
                        CellValue::String(text) => Cell {
                            value: delimited_value(text),
                            number_format: cell.number_format.clone(),
                        },
                        _ => cell.clone(),
                    })
                    .collect()
            })
            .collect();
        Ok(Sheet::from_rows(sheet.name, rows))
    }
}

/// Excel's text-import typing of one delimited field.
pub(crate) fn delimited_value(text: &str) -> CellValue {
    let trimmed = text.trim();
    if trimmed.eq_ignore_ascii_case("true") {
        return CellValue::Bool(true);
    }
    if trimmed.eq_ignore_ascii_case("false") {
        return CellValue::Bool(false);
    }
    if is_decimal(trimmed) {
        if let Ok(integer) = trimmed.parse::<i64>() {
            return CellValue::Int(integer);
        }
        if let Ok(float) = trimmed.parse::<f64>() {
            if float.is_finite() {
                return CellValue::Float(float);
            }
        }
    }
    CellValue::String(text.to_string())
}

/// `[+-]digits[.digits][(e|E)[+-]digits]`, with digits on at least one side
/// of the point. Rejects `inf`, `NaN`, hex, and thousands separators, which
/// `str::parse::<f64>` would otherwise accept or misread.
fn is_decimal(text: &str) -> bool {
    let bytes = text.as_bytes();
    let mut index = usize::from(matches!(bytes.first(), Some(b'+' | b'-')));
    let mut digits = 0;
    while index < bytes.len() && bytes[index].is_ascii_digit() {
        index += 1;
        digits += 1;
    }
    if index < bytes.len() && bytes[index] == b'.' {
        index += 1;
        while index < bytes.len() && bytes[index].is_ascii_digit() {
            index += 1;
            digits += 1;
        }
    }
    if digits == 0 {
        return false;
    }
    if index < bytes.len() && matches!(bytes[index], b'e' | b'E') {
        index += 1;
        if index < bytes.len() && matches!(bytes[index], b'+' | b'-') {
            index += 1;
        }
        let start = index;
        while index < bytes.len() && bytes[index].is_ascii_digit() {
            index += 1;
        }
        if index == start {
            return false;
        }
    }
    index == bytes.len()
}

fn spooled(bytes: &[u8], format: InputFormat, stem: &str) -> Result<Input> {
    let spool = Spool::create()?;
    let path = spool.0.join(format!("{stem}.{}", format.extension()));
    fs::write(&path, bytes).with_context(|| format!("spooling input to {}", path.display()))?;
    let workbook =
        Workbook::open(&path).with_context(|| format!("opening {stem} as {format:?}"))?;
    Ok(Input::new(workbook, Some(spool)))
}

/// A private temporary directory removed on drop.
struct Spool(PathBuf);

impl Spool {
    fn create() -> Result<Self> {
        let stamp = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .map(|elapsed| elapsed.as_nanos())
            .unwrap_or_default();
        let directory =
            std::env::temp_dir().join(format!("wolfxl-input-{}-{stamp}", std::process::id()));
        // `create_dir` refuses an existing path, so a pre-planted directory or
        // symlink is never reused.
        fs::create_dir(&directory)
            .with_context(|| format!("creating spool directory {}", directory.display()))?;
        Ok(Self(directory))
    }
}

impl Drop for Spool {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.0);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn delimited_fields_take_excel_import_types() {
        assert_eq!(delimited_value("420000"), CellValue::Int(420000));
        assert_eq!(delimited_value(" -1.5e3 "), CellValue::Float(-1500.0));
        assert_eq!(delimited_value(".5"), CellValue::Float(0.5));
        assert_eq!(delimited_value("TRUE"), CellValue::Bool(true));
        for text in [
            "inf", "NaN", "1,234", "0x10", "1e", "-", ".", "12a", "1.2.3",
        ] {
            assert_eq!(
                delimited_value(text),
                CellValue::String(text.to_string()),
                "{text}"
            );
        }
    }
}
