//! `wolfxl fingerprint FILE` and `wolfxl verify FILE --fingerprint F`: a
//! content hash of a workbook's sheets, cell values, and formulas.
//!
//! The fingerprint is `wxf1:` followed by the lowercase hex SHA-256 of a
//! canonical `wxf1` stream. The stream is UTF-8 text, one record per line,
//! each line ending in LF. Fields are separated by TAB, and inside a field
//! `\` becomes `\\`, TAB becomes `\t`, CR becomes `\r`, and LF becomes `\n`.
//!
//! ```text
//! wxf1
//! S<TAB>sheet name                      one per sheet, in workbook order
//! V<TAB>A1<TAB>type<TAB>value           one per non-empty cell, row-major
//! F<TAB>A1<TAB>formula                  one per formula cell, row-major
//! ```
//!
//! A sheet's `V` lines follow its `S` line, and its `F` lines follow its last
//! `V` line. Types are `s` (text), `n` (number), `b` (boolean, `TRUE` or
//! `FALSE`), and `e` (Excel error literal such as `#DIV/0!`). Numbers print
//! as integer digits when integral and otherwise as the shortest decimal that
//! round-trips the IEEE double; negative zero prints as `0`. Formula text has
//! no leading `=`.
//!
//! Only content enters the stream, never styles or package bytes:
//!
//! - XLSX, XLSM, and XLAM cells hash their stored values (cached results for
//!   formula cells) exactly as written, so a number-format change leaves the
//!   fingerprint alone, and every formula adds an `F` line.
//! - XLS, XLSB, and ODS cells hash their stored values. Dates and times the
//!   reader types from number formats hash as 1900-system Excel serial
//!   numbers; a date outside Excel's serial range hashes as ISO 8601 text.
//! - CSV, TSV, and TXT fields hash as their source text (type `s`), so the
//!   fingerprint is lossless. The single synthetic sheet is named after the
//!   file, so its `S` line carries an empty name and a renamed copy keeps its
//!   fingerprint.
//!
//! `verify` exits 0 on a match, 1 on a mismatch, and 2 for input failures or
//! an expected fingerprint that is not `wxf1:` plus 64 hex digits.

use std::borrow::Cow;
use std::io::Write;
use std::path::Path;
use std::process::ExitCode;

use anyhow::{anyhow, Context, Result};
use chrono::{NaiveDate, NaiveDateTime, NaiveTime, Timelike};
use serde_json::json;
use sha2::{Digest, Sha256};
use wolfxl_core::{CellValue, SourceFormat, Workbook};
use wolfxl_reader::{CellValue as NativeValue, NativeXlsxBook};

use crate::{FingerprintArgs, VerifyArgs};

/// Stream format version, the first stream line and the fingerprint prefix.
const VERSION: &str = "wxf1";

/// A computed fingerprint and the record counts behind it.
struct Fingerprint {
    hex: String,
    sheets: usize,
    cells: usize,
    formulas: usize,
}

impl Fingerprint {
    fn text(&self) -> String {
        format!("{VERSION}:{}", self.hex)
    }
}

pub(crate) fn fingerprint(args: FingerprintArgs) -> ExitCode {
    let fingerprint = match compute(&args.file) {
        Ok(fingerprint) => fingerprint,
        Err(error) => {
            eprintln!("error: {error:#}");
            return ExitCode::from(2);
        }
    };
    let document = if args.json {
        let document = json!({
            "fingerprint": fingerprint.text(),
            "version": VERSION,
            "sheets": fingerprint.sheets,
            "cells": fingerprint.cells,
            "formulas": fingerprint.formulas,
        });
        format!("{document}\n")
    } else {
        format!("{}\n", fingerprint.text())
    };
    finish(&document, ExitCode::SUCCESS)
}

pub(crate) fn verify(args: VerifyArgs) -> ExitCode {
    let Some(expected) = expected_hex(&args.fingerprint) else {
        eprintln!(
            "error: unknown fingerprint version in {:?}; expected {VERSION}: followed by 64 hex digits",
            args.fingerprint
        );
        return ExitCode::from(2);
    };
    let actual = match compute(&args.file) {
        Ok(fingerprint) => fingerprint,
        Err(error) => {
            eprintln!("error: {error:#}");
            return ExitCode::from(2);
        }
    };
    if actual.hex == expected {
        finish(&format!("ok {}\n", actual.text()), ExitCode::SUCCESS)
    } else {
        let document = format!(
            "mismatch\nexpected {VERSION}:{expected}\nactual {}\n",
            actual.text()
        );
        finish(&document, ExitCode::from(1))
    }
}

/// The lowercase digest of a `wxf1:<64 hex>` fingerprint, or `None` for any
/// other version prefix or a malformed digest.
fn expected_hex(fingerprint: &str) -> Option<String> {
    let hex = fingerprint
        .trim()
        .strip_prefix(VERSION)?
        .strip_prefix(':')?;
    (hex.len() == 64 && hex.bytes().all(|byte| byte.is_ascii_hexdigit()))
        .then(|| hex.to_ascii_lowercase())
}

fn finish(document: &str, code: ExitCode) -> ExitCode {
    let mut stdout = std::io::stdout().lock();
    match stdout
        .write_all(document.as_bytes())
        .and_then(|()| stdout.flush())
    {
        Err(error) if error.kind() != std::io::ErrorKind::BrokenPipe => {
            eprintln!("error: writing result: {error}");
            ExitCode::from(2)
        }
        _ => code,
    }
}

fn compute(path: &Path) -> Result<Fingerprint> {
    let mut workbook =
        Workbook::open(path).with_context(|| format!("opening {}", path.display()))?;
    let names = workbook.sheet_names().to_vec();
    let mut stream = Stream::new();
    match workbook.format() {
        SourceFormat::Xlsx => {
            let book = NativeXlsxBook::open_path(path)
                .map_err(|error| anyhow!("opening {}: {error}", path.display()))?;
            for name in &names {
                stream.sheet(name);
                ooxml_sheet(&mut stream, &book, name)?;
            }
        }
        format => {
            let delimited = format == SourceFormat::Csv;
            for name in &names {
                let sheet = workbook
                    .sheet(name)
                    .with_context(|| format!("loading sheet {name:?}"))?;
                stream.sheet(if delimited { "" } else { name });
                for (row_index, row) in sheet.rows().iter().enumerate() {
                    for (column_index, cell) in row.iter().enumerate() {
                        if let Some((kind, text)) = typed(&cell.value) {
                            stream.value(&sheet.cell_address(row_index, column_index), kind, &text);
                        }
                    }
                }
            }
        }
    }
    Ok(stream.finish())
}

/// Stream one OOXML worksheet's stored values, then its formulas.
fn ooxml_sheet(stream: &mut Stream, book: &NativeXlsxBook, name: &str) -> Result<()> {
    let mut cells = book
        .worksheet(name)
        .map_err(|error| anyhow!("loading sheet {name:?}: {error}"))?
        .cells;
    cells.sort_unstable_by_key(|cell| (cell.row, cell.col));
    for cell in &cells {
        let (kind, text) = match &cell.value {
            NativeValue::Empty => continue,
            NativeValue::String(text) => ("s", Cow::Borrowed(text.as_str())),
            NativeValue::SharedString(text) => ("s", Cow::Borrowed(&**text)),
            NativeValue::Number(number) => ("n", Cow::Owned(number_text(*number))),
            NativeValue::Bool(flag) => ("b", Cow::Borrowed(bool_text(*flag))),
            NativeValue::Error(literal) => ("e", Cow::Borrowed(literal.as_str())),
        };
        stream.value(&a1(cell.row, cell.col), kind, &text);
    }
    for cell in &cells {
        if let Some(formula) = &cell.formula {
            stream.formula(&a1(cell.row, cell.col), formula);
        }
    }
    Ok(())
}

/// The stream type and value text of a reader-typed cell; `None` when empty.
fn typed(value: &CellValue) -> Option<(&'static str, Cow<'_, str>)> {
    Some(match value {
        CellValue::Empty => return None,
        CellValue::String(text) => ("s", Cow::Borrowed(text.as_str())),
        CellValue::Int(number) => ("n", Cow::Owned(number.to_string())),
        CellValue::Float(number) => ("n", Cow::Owned(number_text(*number))),
        CellValue::Bool(flag) => ("b", Cow::Borrowed(bool_text(*flag))),
        CellValue::Error(literal) => ("e", Cow::Borrowed(literal.as_str())),
        CellValue::Date(date) => match date_serial(*date) {
            Some(serial) => ("n", Cow::Owned(number_text(serial))),
            None => ("s", Cow::Owned(date.to_string())),
        },
        CellValue::DateTime(moment) => match date_serial(moment.date()) {
            Some(serial) => (
                "n",
                Cow::Owned(number_text(serial + day_fraction(moment.time()))),
            ),
            None => ("s", Cow::Owned(iso_datetime(moment))),
        },
        CellValue::Time(time) => ("n", Cow::Owned(number_text(day_fraction(*time)))),
    })
}

/// Whole-day 1900-system Excel serial of a date, or `None` outside Excel's
/// 1900-01-01 through 9999-12-31 range. Serial 1 is 1900-01-01, and serials
/// from 61 on count the nonexistent 1900-02-29 that Excel keeps as serial 60.
pub(crate) fn date_serial(date: NaiveDate) -> Option<f64> {
    let first = NaiveDate::from_ymd_opt(1900, 1, 1)?;
    let last = NaiveDate::from_ymd_opt(9999, 12, 31)?;
    if date < first || date > last {
        return None;
    }
    let days = date.signed_duration_since(first).num_days() + 1;
    Some((if days >= 60 { days + 1 } else { days }) as f64)
}

/// Fraction of a day elapsed at `time`: whole seconds since midnight plus the
/// fractional second, over 86,400.
pub(crate) fn day_fraction(time: NaiveTime) -> f64 {
    let seconds = f64::from(time.num_seconds_from_midnight());
    let fraction = f64::from(time.nanosecond()) / 1_000_000_000.0;
    (seconds + fraction) / 86_400.0
}

fn iso_datetime(moment: &NaiveDateTime) -> String {
    moment.format("%Y-%m-%dT%H:%M:%S%.f").to_string()
}

/// Canonical number text: integer digits for integral values (which matches
/// `i64` text for every value an `f64` holds exactly), the shortest
/// round-trip decimal otherwise, and `0` for negative zero.
fn number_text(number: f64) -> String {
    if number == 0.0 {
        "0".to_string()
    } else {
        number.to_string()
    }
}

fn bool_text(flag: bool) -> &'static str {
    if flag {
        "TRUE"
    } else {
        "FALSE"
    }
}

/// A1 address of a 1-based row and column.
fn a1(row: u32, column: u32) -> String {
    let mut letters = Vec::new();
    let mut remaining = column;
    while remaining > 0 {
        remaining -= 1;
        letters.push(b'A' + (remaining % 26) as u8);
        remaining /= 26;
    }
    letters.reverse();
    let mut address = String::from_utf8(letters).expect("column letters are ASCII");
    address.push_str(&row.to_string());
    address
}

/// Incremental SHA-256 over the canonical stream, with record counts.
struct Stream {
    hasher: Sha256,
    sheets: usize,
    cells: usize,
    formulas: usize,
}

impl Stream {
    fn new() -> Self {
        let mut stream = Self {
            hasher: Sha256::new(),
            sheets: 0,
            cells: 0,
            formulas: 0,
        };
        stream.line(&[VERSION]);
        stream
    }

    fn sheet(&mut self, name: &str) {
        self.sheets += 1;
        self.line(&["S", name]);
    }

    fn value(&mut self, address: &str, kind: &str, text: &str) {
        self.cells += 1;
        self.line(&["V", address, kind, text]);
    }

    fn formula(&mut self, address: &str, formula: &str) {
        self.formulas += 1;
        let formula = formula.strip_prefix('=').unwrap_or(formula);
        self.line(&["F", address, formula]);
    }

    fn line(&mut self, fields: &[&str]) {
        for (index, field) in fields.iter().enumerate() {
            if index > 0 {
                self.hasher.update(b"\t");
            }
            self.escaped(field);
        }
        self.hasher.update(b"\n");
    }

    fn escaped(&mut self, field: &str) {
        let bytes = field.as_bytes();
        let mut start = 0;
        for (index, byte) in bytes.iter().enumerate() {
            let escape: &[u8] = match byte {
                b'\\' => b"\\\\",
                b'\t' => b"\\t",
                b'\r' => b"\\r",
                b'\n' => b"\\n",
                _ => continue,
            };
            self.hasher.update(&bytes[start..index]);
            self.hasher.update(escape);
            start = index + 1;
        }
        self.hasher.update(&bytes[start..]);
    }

    fn finish(self) -> Fingerprint {
        Fingerprint {
            hex: base16ct::lower::encode_string(&self.hasher.finalize()),
            sheets: self.sheets,
            cells: self.cells,
            formulas: self.formulas,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn numbers_share_one_canonical_text() {
        let int_text = typed(&CellValue::Int(42)).map(|(_, text)| text.into_owned());
        let float_text = typed(&CellValue::Float(42.0)).map(|(_, text)| text.into_owned());
        assert_eq!(int_text.as_deref(), Some("42"));
        assert_eq!(float_text, int_text);
        assert_eq!(number_text(-0.0), "0");
        assert_eq!(number_text(0.1), "0.1");
        assert_eq!(number_text(1e21), "1000000000000000000000");
        assert_eq!(number_text(2f64.powi(53)), (1_i64 << 53).to_string());
    }

    #[test]
    fn dates_hash_as_1900_system_serials_or_iso_text() {
        let date = |year, month, day| NaiveDate::from_ymd_opt(year, month, day).unwrap();
        assert_eq!(date_serial(date(1900, 1, 1)), Some(1.0));
        assert_eq!(date_serial(date(1900, 2, 28)), Some(59.0));
        assert_eq!(date_serial(date(1900, 3, 1)), Some(61.0));
        assert_eq!(date_serial(date(2000, 1, 1)), Some(36_526.0));
        assert_eq!(date_serial(date(9999, 12, 31)), Some(2_958_465.0));
        assert_eq!(date_serial(date(1899, 12, 31)), None);

        let noon = NaiveTime::from_hms_opt(12, 0, 0).unwrap();
        assert_eq!(day_fraction(noon), 0.5);
        let moment = date(2000, 1, 1).and_time(noon);
        let text = |value| typed(&value).map(|(kind, text)| (kind, text.into_owned()));
        assert_eq!(
            text(CellValue::DateTime(moment)),
            Some(("n", "36526.5".to_string()))
        );
        assert_eq!(
            text(CellValue::Date(date(1899, 12, 31))),
            Some(("s", "1899-12-31".to_string()))
        );
        assert_eq!(
            text(CellValue::DateTime(date(1899, 12, 31).and_time(noon))),
            Some(("s", "1899-12-31T12:00:00".to_string()))
        );
        assert_eq!(text(CellValue::Time(noon)), Some(("n", "0.5".to_string())));
    }

    #[test]
    fn addresses_use_bijective_column_letters() {
        assert_eq!(a1(1, 1), "A1");
        assert_eq!(a1(7, 26), "Z7");
        assert_eq!(a1(2, 27), "AA2");
        assert_eq!(a1(1_048_576, 16_384), "XFD1048576");
    }

    #[test]
    fn expected_fingerprints_need_the_wxf1_prefix_and_64_hex_digits() {
        let digest = "AB".repeat(32);
        assert_eq!(
            expected_hex(&format!(" wxf1:{digest}\n")),
            Some("ab".repeat(32))
        );
        assert_eq!(expected_hex(&format!("wxf2:{digest}")), None);
        assert_eq!(expected_hex(&digest), None);
        assert_eq!(expected_hex("wxf1:abc"), None);
        assert_eq!(expected_hex(&format!("wxf1:{}", "g".repeat(64))), None);
    }
}
