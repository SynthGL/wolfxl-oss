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
//!
//! Spreadsheet files also expose their stored cells through [`RawBook`]: the
//! core reader types numbers from their number formats and date system, which
//! rounds date-times to whole seconds and shifts 1904-system serials, so
//! fingerprints and formulas read the native readers' cells instead.

use std::fs;
use std::io::Read;
use std::path::{Path, PathBuf};
use std::time::{SystemTime, UNIX_EPOCH};

use anyhow::{anyhow, bail, Context, Result};
use chrono::{NaiveDate, NaiveDateTime, NaiveTime, Timelike};
use clap::ValueEnum;
use wolfxl_core::{Cell, CellValue, Sheet, SourceFormat, Workbook};
use wolfxl_reader::{
    Cell as RawCell, CellValue as RawValue, NativeOdsBook, NativeXlsBook, NativeXlsbBook,
    NativeXlsxBook,
};

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

    /// The stored cells of a spreadsheet file, or `None` for delimited text.
    pub(crate) fn raw_book(&self) -> Result<Option<RawBook>> {
        RawBook::open(&self.workbook)
    }
}

/// A spreadsheet file opened with its native reader, whose cells hold the
/// numbers, text, booleans, errors, and formulas exactly as stored.
pub(crate) enum RawBook {
    Xlsx(Box<NativeXlsxBook>),
    Xls(Box<NativeXlsBook>),
    Xlsb(Box<NativeXlsbBook>),
    Ods(Box<NativeOdsBook>),
}

impl RawBook {
    /// Reopen `workbook`'s file natively; `None` for delimited text, which
    /// has no stored types.
    pub(crate) fn open(workbook: &Workbook) -> Result<Option<Self>> {
        let path = workbook.path();
        let opening =
            |error: &dyn std::fmt::Display| anyhow!("opening {}: {error}", path.display());
        Ok(Some(match workbook.format() {
            SourceFormat::Xlsx => Self::Xlsx(Box::new(
                NativeXlsxBook::open_path(path).map_err(|e| opening(&e))?,
            )),
            SourceFormat::Xls => Self::Xls(Box::new(
                NativeXlsBook::open_path(path).map_err(|e| opening(&e))?,
            )),
            SourceFormat::Xlsb => Self::Xlsb(Box::new(
                NativeXlsbBook::open_path(path).map_err(|e| opening(&e))?,
            )),
            SourceFormat::Ods => Self::Ods(Box::new(
                NativeOdsBook::open_path(path).map_err(|e| opening(&e))?,
            )),
            SourceFormat::Csv => return Ok(None),
        }))
    }

    /// Sheet `name`'s cells that hold a value or a formula, row-major.
    pub(crate) fn cells(&self, name: &str) -> Result<Vec<RawCell>> {
        let loading = |error: &dyn std::fmt::Display| anyhow!("loading sheet {name:?}: {error}");
        let mut cells = match self {
            Self::Xlsx(book) => book.worksheet(name).map_err(|e| loading(&e))?,
            Self::Xls(book) => book.worksheet(name).map_err(|e| loading(&e))?,
            Self::Xlsb(book) => book.worksheet(name).map_err(|e| loading(&e))?,
            Self::Ods(book) => book.worksheet(name).map_err(|e| loading(&e))?,
        }
        .cells;
        cells.retain(|cell| cell.formula.is_some() || !matches!(cell.value, RawValue::Empty));
        cells.sort_unstable_by_key(|cell| (cell.row, cell.col));
        Ok(cells)
    }
}

/// 1900-system Excel serial of an ISO 8601 date or date-time that a reader
/// stored as text (ODS date values, OOXML `t="d"` cells), or `None` for other
/// text and for dates outside Excel's 1900-01-01 through 9999-12-31 range.
pub(crate) fn iso_serial(text: &str) -> Option<f64> {
    let moment = NaiveDateTime::parse_from_str(text, "%Y-%m-%dT%H:%M:%S%.f")
        .or_else(|_| NaiveDateTime::parse_from_str(text, "%Y-%m-%dT%H:%M:%S"))
        .or_else(|_| {
            NaiveDate::parse_from_str(text, "%Y-%m-%d").map(|date| date.and_time(NaiveTime::MIN))
        })
        .ok()?;
    Some(date_serial(moment.date())? + day_fraction(moment.time()))
}

/// Whole-day serial of a date, or `None` outside Excel's range. Serial 1 is
/// 1900-01-01, and serials from 61 on count the nonexistent 1900-02-29 that
/// Excel keeps as serial 60.
fn date_serial(date: NaiveDate) -> Option<f64> {
    let first = NaiveDate::from_ymd_opt(1900, 1, 1)?;
    let last = NaiveDate::from_ymd_opt(9999, 12, 31)?;
    if date < first || date > last {
        return None;
    }
    let days = date.signed_duration_since(first).num_days() + 1;
    Some((if days >= 60 { days + 1 } else { days }) as f64)
}

/// Fraction of a day elapsed at `time`, fractional seconds included.
fn day_fraction(time: NaiveTime) -> f64 {
    let seconds = f64::from(time.num_seconds_from_midnight());
    let fraction = f64::from(time.nanosecond()) / 1_000_000_000.0;
    (seconds + fraction) / 86_400.0
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

    #[test]
    fn iso_dates_become_1900_system_serials_inside_excels_range() {
        let date = |year, month, day| NaiveDate::from_ymd_opt(year, month, day).unwrap();
        assert_eq!(date_serial(date(1900, 1, 1)), Some(1.0));
        assert_eq!(date_serial(date(1900, 2, 28)), Some(59.0));
        assert_eq!(date_serial(date(1900, 3, 1)), Some(61.0));
        assert_eq!(date_serial(date(2000, 1, 1)), Some(36_526.0));
        assert_eq!(date_serial(date(9999, 12, 31)), Some(2_958_465.0));
        assert_eq!(date_serial(date(1899, 12, 31)), None);

        assert_eq!(iso_serial("2000-01-01"), Some(36_526.0));
        assert_eq!(iso_serial("2000-01-01T12:00:00"), Some(36_526.5));
        assert_eq!(
            iso_serial("2000-01-01T00:00:00.5"),
            Some(36_526.0 + 0.5 / 86_400.0)
        );
        assert_eq!(iso_serial("1899-12-31T12:00:00"), None);
        assert_eq!(iso_serial("PT12H30M00S"), None);
        assert_eq!(iso_serial("total"), None);
    }
}
