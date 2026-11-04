//! Spreadsheet input for commands that read a file path or standard input.
//!
//! `wolfxl_core::Workbook` dispatches on the file extension, so piped bytes
//! and paths whose extension disagrees with `--from` are streamed into a
//! private temporary directory under a name carrying the requested
//! extension. The copy is capped at the format's spool limit
//! ([`InputFormat::spool_limit`]); the spool lives as long as the [`Input`]
//! and is removed on drop or when spooling fails.
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

use std::fs::{self, File};
use std::io::{self, Read};
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicU64, Ordering};
use std::time::{SystemTime, UNIX_EPOCH};

use anyhow::{anyhow, bail, Context, Result};
use chrono::{NaiveDate, NaiveDateTime, NaiveTime, Timelike};
use clap::ValueEnum;
use wolfxl_core::{Cell, CellValue, Sheet, SourceFormat, Workbook};
use wolfxl_reader::{
    Cell as RawCell, CellValue as RawValue, NativeOdsBook, NativeXlsBook, NativeXlsbBook,
    NativeXlsxBook, WorksheetData,
};

/// Most bytes spooled for `--from xls`: the native XLS reader's source
/// admission limit, so the copy stops where the reader would refuse the
/// source anyway.
const MAX_XLS_SPOOL_BYTES: u64 = 512 * 1024 * 1024;

/// Most bytes spooled for every other format: the reader's default budget
/// for the uncompressed parts of an OOXML or ODS package
/// (`wolfxl_rels::package_guard::DEFAULT_MAX_ZIP_TOTAL_BYTES`). Delimited
/// text has no reader limit, so the same bound keeps an unending pipe from
/// filling the temporary directory.
const MAX_SPOOL_BYTES: u64 = 4 * 1024 * 1024 * 1024;

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

    /// Most bytes of this format copied into a spool.
    fn spool_limit(self) -> u64 {
        match self {
            InputFormat::Xls => MAX_XLS_SPOOL_BYTES,
            _ => MAX_SPOOL_BYTES,
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
        return spooled(std::io::stdin().lock(), "standard input", format, "stdin");
    };
    match from {
        Some(format) if InputFormat::of_path(path) != Some(format) => {
            let label = path.display().to_string();
            let file = File::open(path).with_context(|| format!("reading {label}"))?;
            // A regular file's length is known up front, so an oversized one
            // fails before any byte is copied.
            let metadata = file
                .metadata()
                .with_context(|| format!("reading {label}"))?;
            if metadata.is_file() && metadata.len() > format.spool_limit() {
                return Err(too_large(&label, format, format.spool_limit()));
            }
            let stem = path
                .file_stem()
                .and_then(|stem| stem.to_str())
                .unwrap_or("Sheet1");
            spooled(file, &label, format, stem)
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

    /// Whether the workbook numbers its date serials in the 1904 date system.
    pub(crate) fn date1904(&self) -> bool {
        match self {
            Self::Xlsx(book) => book.date1904(),
            Self::Xls(book) => book.date1904(),
            Self::Xlsb(book) => book.date1904(),
            Self::Ods(book) => book.date1904(),
        }
    }

    /// Sheet `name` as its native reader parses it, metadata included.
    pub(crate) fn worksheet(&self, name: &str) -> Result<WorksheetData> {
        let loading = |error: &dyn std::fmt::Display| anyhow!("loading sheet {name:?}: {error}");
        match self {
            Self::Xlsx(book) => book.worksheet(name).map_err(|e| loading(&e)),
            Self::Xls(book) => book.worksheet(name).map_err(|e| loading(&e)),
            Self::Xlsb(book) => book.worksheet(name).map_err(|e| loading(&e)),
            Self::Ods(book) => book.worksheet(name).map_err(|e| loading(&e)),
        }
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

/// Stream `source` (described as `label` in errors) into a fresh spool and
/// open it as `format`.
fn spooled(source: impl Read, label: &str, format: InputFormat, stem: &str) -> Result<Input> {
    let (spool, path) = spool_input(
        &std::env::temp_dir(),
        source,
        label,
        format,
        stem,
        format.spool_limit(),
    )?;
    let workbook =
        Workbook::open(&path).with_context(|| format!("opening {stem} as {format:?}"))?;
    Ok(Input::new(workbook, Some(spool)))
}

/// Copy at most `limit` bytes of `source` into `{stem}.{extension}` inside a
/// new spool under `root`, reading through a fixed-size buffer. More bytes
/// than `limit` fail, and the spool with its partial copy is removed.
fn spool_input(
    root: &Path,
    source: impl Read,
    label: &str,
    format: InputFormat,
    stem: &str,
    limit: u64,
) -> Result<(Spool, PathBuf)> {
    let spool = Spool::create(root)?;
    let path = spool.0.join(format!("{stem}.{}", format.extension()));
    let mut file =
        private_file(&path).with_context(|| format!("spooling input to {}", path.display()))?;
    // One byte past the limit tells an oversized source from one that fits.
    let copied = io::copy(&mut source.take(limit.saturating_add(1)), &mut file)
        .with_context(|| format!("spooling {label} to {}", path.display()))?;
    if copied > limit {
        return Err(too_large(label, format, limit));
    }
    Ok((spool, path))
}

/// Spooled input larger than its format's spool limit. Commands find it with
/// `anyhow::Error::downcast_ref` to report it as a usage failure.
#[derive(Debug)]
pub(crate) struct SpoolLimitExceeded {
    label: String,
    format: InputFormat,
    limit: u64,
}

impl std::fmt::Display for SpoolLimitExceeded {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(
            f,
            "{} exceeds the {} spool limit for --from {} input",
            self.label,
            byte_size(self.limit),
            self.format.extension()
        )
    }
}

impl std::error::Error for SpoolLimitExceeded {}

fn too_large(label: &str, format: InputFormat, limit: u64) -> anyhow::Error {
    SpoolLimitExceeded {
        label: label.to_string(),
        format,
        limit,
    }
    .into()
}

/// `bytes` in the largest binary unit that divides it evenly.
fn byte_size(bytes: u64) -> String {
    const MIB: u64 = 1024 * 1024;
    const GIB: u64 = 1024 * MIB;
    if bytes >= GIB && bytes.is_multiple_of(GIB) {
        format!("{} GiB", bytes / GIB)
    } else if bytes >= MIB && bytes.is_multiple_of(MIB) {
        format!("{} MiB", bytes / MIB)
    } else {
        format!("{bytes} bytes")
    }
}

/// Create directory `path`, readable only by its owner from the moment it
/// exists. Refuses an existing path, symlinks included.
fn private_dir(path: &Path) -> io::Result<()> {
    let mut builder = fs::DirBuilder::new();
    #[cfg(unix)]
    {
        use std::os::unix::fs::DirBuilderExt;
        builder.mode(0o700);
    }
    builder.create(path)
}

/// Create file `path` for writing, readable only by its owner from the
/// moment it exists. Refuses an existing path, symlinks included.
fn private_file(path: &Path) -> io::Result<File> {
    let mut options = fs::OpenOptions::new();
    options.write(true).create_new(true);
    #[cfg(unix)]
    {
        use std::os::unix::fs::OpenOptionsExt;
        options.mode(0o600);
    }
    options.open(path)
}

/// A private temporary directory removed on drop.
struct Spool(PathBuf);

impl Spool {
    fn create(root: &Path) -> Result<Self> {
        // Distinguishes spools created in the same process within one clock tick.
        static SEQUENCE: AtomicU64 = AtomicU64::new(0);
        let stamp = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .map(|elapsed| elapsed.as_nanos())
            .unwrap_or_default();
        let sequence = SEQUENCE.fetch_add(1, Ordering::Relaxed);
        let directory = root.join(format!(
            "wolfxl-input-{}-{stamp}-{sequence}",
            std::process::id()
        ));
        // The mode applies at creation, so the directory is never open to
        // other users, and a pre-planted directory or symlink is refused
        // rather than reused.
        private_dir(&directory)
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

    /// Spool `bytes` as CSV under `root` with spool limit `limit`.
    fn spool_csv(root: &Path, bytes: &[u8], limit: u64) -> Result<(Spool, PathBuf)> {
        spool_input(
            root,
            bytes,
            "standard input",
            InputFormat::Csv,
            "stdin",
            limit,
        )
    }

    #[test]
    fn spooling_stops_one_byte_past_the_limit_and_removes_the_partial_copy() {
        let root = Spool::create(&std::env::temp_dir()).unwrap();
        // An endless source shows the copy is bounded by the limit, not the input.
        let error = spool_input(
            &root.0,
            io::repeat(b'a'),
            "standard input",
            InputFormat::Xlsx,
            "stdin",
            8,
        )
        .err()
        .expect("an endless source exceeds the limit");
        assert_eq!(
            error.to_string(),
            "standard input exceeds the 8 bytes spool limit for --from xlsx input"
        );
        assert!(
            error
                .context("reading the input")
                .downcast_ref::<SpoolLimitExceeded>()
                .is_some(),
            "the limit error stays typed through added context"
        );
        assert_eq!(fs::read_dir(&root.0).unwrap().count(), 0, "spool removed");

        let (spool, path) = spool_csv(&root.0, b"a,b\n1,2", 7).expect("exactly the limit fits");
        assert_eq!(path.file_name().unwrap(), "stdin.csv");
        assert_eq!(fs::read(&path).unwrap(), b"a,b\n1,2");
        drop(spool);
        assert!(!path.exists(), "spool removed on drop");
    }

    #[test]
    fn spool_limits_name_the_reader_admission_budgets() {
        assert_eq!(byte_size(InputFormat::Xls.spool_limit()), "512 MiB");
        for format in [
            InputFormat::Csv,
            InputFormat::Tsv,
            InputFormat::Xlsx,
            InputFormat::Xlsm,
            InputFormat::Xlsb,
            InputFormat::Ods,
        ] {
            assert_eq!(byte_size(format.spool_limit()), "4 GiB", "{format:?}");
        }
        assert_eq!(byte_size(1536), "1536 bytes");
    }

    #[test]
    fn spool_entries_refuse_existing_paths_and_symlinks() {
        let root = Spool::create(&std::env::temp_dir()).unwrap();
        let existing = root.0.join("existing.csv");
        fs::write(&existing, "keep").unwrap();
        assert_eq!(
            private_file(&existing).unwrap_err().kind(),
            io::ErrorKind::AlreadyExists
        );
        assert_eq!(
            private_dir(&existing).unwrap_err().kind(),
            io::ErrorKind::AlreadyExists
        );
        assert_eq!(fs::read_to_string(&existing).unwrap(), "keep");

        #[cfg(unix)]
        {
            use std::os::unix::fs::symlink;
            let missing = root.0.join("missing.csv");
            for (link, target) in [("link.csv", &existing), ("dangling.csv", &missing)] {
                let link = root.0.join(link);
                symlink(target, &link).unwrap();
                assert_eq!(
                    private_file(&link).unwrap_err().kind(),
                    io::ErrorKind::AlreadyExists
                );
                assert_eq!(
                    private_dir(&link).unwrap_err().kind(),
                    io::ErrorKind::AlreadyExists
                );
            }
            assert_eq!(fs::read_to_string(&existing).unwrap(), "keep");
            assert!(!missing.exists(), "a dangling symlink is never followed");
        }
    }

    /// Permission bits of `path`.
    #[cfg(unix)]
    fn mode(path: &Path) -> u32 {
        use std::os::unix::fs::PermissionsExt;
        fs::symlink_metadata(path).unwrap().permissions().mode() & 0o777
    }

    #[cfg(unix)]
    #[test]
    fn spool_directory_and_file_are_owner_only() {
        let root = Spool::create(&std::env::temp_dir()).unwrap();
        let (spool, path) = spool_csv(&root.0, b"a\n1\n", 4).unwrap();
        assert_eq!(mode(&root.0), 0o700);
        assert_eq!(mode(&spool.0), 0o700);
        assert_eq!(mode(&path), 0o600);
    }

    /// Set in the child process of
    /// `spool_stays_owner_only_under_a_permissive_umask`.
    #[cfg(unix)]
    const UMASK_CHILD: &str = "WOLFXL_TEST_SPOOL_UMASK_CHILD";

    #[cfg(unix)]
    #[test]
    fn spool_stays_owner_only_under_a_permissive_umask() {
        let root = Spool::create(&std::env::temp_dir()).unwrap();
        if std::env::var_os(UMASK_CHILD).is_some() {
            // Confirms the child really runs under umask 000.
            let probe = root.0.join("probe");
            fs::write(&probe, "").unwrap();
            assert_eq!(mode(&probe), 0o666, "child umask is 000");

            let (spool, path) = spool_csv(&root.0, b"a\n1\n", 4).unwrap();
            assert_eq!(mode(&spool.0), 0o700);
            assert_eq!(mode(&path), 0o600);
            return;
        }
        // The umask is process-wide, so this test reruns itself alone in a
        // child process started under umask 000.
        let name = "input::tests::spool_stays_owner_only_under_a_permissive_umask";
        let output = std::process::Command::new("/bin/sh")
            .args(["-c", "umask 000 && exec \"$0\" \"$@\""])
            .arg(std::env::current_exe().unwrap())
            .args(["--exact", name, "--test-threads=1"])
            .env(UMASK_CHILD, "1")
            .output()
            .unwrap();
        let stdout = String::from_utf8_lossy(&output.stdout);
        assert!(
            output.status.success() && stdout.contains("1 passed"),
            "{stdout}{}",
            String::from_utf8_lossy(&output.stderr)
        );
    }
}
