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
//! - XLS, XLSB, and ODS cells hash their stored values (cached results for
//!   formula cells) without formula lines, so neither a number format nor
//!   the workbook's date system changes how a number hashes. ODS stores date
//!   and date-time values as ISO 8601 text; those hash as 1900-system Excel
//!   serial numbers, or as their text outside Excel's serial range.
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

use anyhow::{Context, Result};
use serde_json::{json, Value};
use sha2::{Digest, Sha256};
use wolfxl_core::{CellValue, SourceFormat, Workbook};
use wolfxl_reader::{CellDataType, CellValue as NativeValue};

use crate::input::{iso_serial, RawBook};
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

    /// The `--json` report document.
    fn json(&self) -> Value {
        json!({
            "fingerprint": self.text(),
            "version": VERSION,
            "sheets": self.sheets,
            "cells": self.cells,
            "formulas": self.formulas,
        })
    }
}

/// The `fingerprint --json` report for `path`, computed without writing
/// anything.
pub(crate) fn structured(path: &Path) -> Result<Value> {
    Ok(compute(path)?.json())
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
        format!("{}\n", fingerprint.json())
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
    let Some(book) = RawBook::open(&workbook)? else {
        for name in &names {
            let sheet = workbook
                .sheet(name)
                .with_context(|| format!("loading sheet {name:?}"))?;
            stream.sheet("");
            for (row_index, row) in sheet.rows().iter().enumerate() {
                for (column_index, cell) in row.iter().enumerate() {
                    // Delimited fields are source text or empty.
                    if let CellValue::String(text) = &cell.value {
                        stream.value(&sheet.cell_address(row_index, column_index), "s", text);
                    }
                }
            }
        }
        return Ok(stream.finish());
    };
    let ooxml = workbook.format() == SourceFormat::Xlsx;
    for name in &names {
        stream.sheet(name);
        let cells = book.cells(name)?;
        for cell in &cells {
            let (kind, text) = match &cell.value {
                NativeValue::Empty => continue,
                // OOXML `t="d"` text hashes as written.
                NativeValue::String(text) if !ooxml && cell.data_type == CellDataType::DateIso => {
                    match iso_serial(text) {
                        Some(serial) => ("n", Cow::Owned(number_text(serial))),
                        None => ("s", Cow::Borrowed(text.as_str())),
                    }
                }
                NativeValue::String(text) => ("s", Cow::Borrowed(text.as_str())),
                NativeValue::SharedString(text) => ("s", Cow::Borrowed(&**text)),
                NativeValue::Number(number) => ("n", Cow::Owned(number_text(*number))),
                NativeValue::Bool(flag) => ("b", Cow::Borrowed(bool_text(*flag))),
                NativeValue::Error(literal) => ("e", Cow::Borrowed(literal.as_str())),
            };
            stream.value(&a1(cell.row, cell.col), kind, &text);
        }
        if ooxml {
            for cell in &cells {
                if let Some(formula) = &cell.formula {
                    stream.formula(&a1(cell.row, cell.col), formula);
                }
            }
        }
    }
    Ok(stream.finish())
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
        assert_eq!(number_text(42.0), "42");
        assert_eq!(number_text(-0.0), "0");
        assert_eq!(number_text(0.1), "0.1");
        assert_eq!(number_text(45_000.123_456_789), "45000.123456789");
        assert_eq!(number_text(1e21), "1000000000000000000000");
        assert_eq!(number_text(2f64.powi(53)), (1_i64 << 53).to_string());
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
