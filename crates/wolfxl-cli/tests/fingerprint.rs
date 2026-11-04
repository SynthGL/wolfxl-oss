//! Integration tests for `wolfxl fingerprint` and `wolfxl verify`: the
//! documented `wxf1` stream, stability across package rewrites, sensitivity
//! to value and formula edits, delimited input, and the 0 / 1 / 2 exit-code
//! contract of `verify`.

use std::fs;
use std::io::{Cursor, Read, Write};
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicUsize, Ordering};
use std::time::{SystemTime, UNIX_EPOCH};

use assert_cmd::Command;
use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

static NEXT_TEMP: AtomicUsize = AtomicUsize::new(0);

const CONTENT_TYPES: &str = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>"#;

const ROOT_RELS: &str = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#;

const WORKBOOK: &str = r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"#;

const WORKBOOK_RELS: &str = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#;

/// Cell format 1 applies the built-in `m/d/yyyy` date format (id 14).
const STYLES: &str = r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs></styleSheet>"#;

/// One row covering every stream type, with a formula in C1.
const SHEET: &str = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>2</v></c><c r="B1"><v>3</v></c><c r="C1"><f>A1+B1</f><v>5</v></c><c r="D1" t="inlineStr"><is><t>x</t></is></c><c r="E1" t="b"><v>1</v></c><c r="F1" t="e"><v>#DIV/0!</v></c><c r="G1"><v>0.1</v></c></row></sheetData></worksheet>"#;

/// SHA-256 of the `SHEET` workbook's documented stream:
///
/// ```text
/// wxf1
/// S<TAB>Data
/// V<TAB>A1<TAB>n<TAB>2
/// V<TAB>B1<TAB>n<TAB>3
/// V<TAB>C1<TAB>n<TAB>5
/// V<TAB>D1<TAB>s<TAB>x
/// V<TAB>E1<TAB>b<TAB>TRUE
/// V<TAB>F1<TAB>e<TAB>#DIV/0!
/// V<TAB>G1<TAB>n<TAB>0.1
/// F<TAB>C1<TAB>A1+B1
/// ```
const SHEET_FINGERPRINT: &str =
    "wxf1:bfc7936206e5acb6e1ebd78c71f488c53b5211ea87b5dd939f32f7db65b7bea9";

/// Fields carrying a TAB, a backslash, and a quoted line break.
const CSV: &str = "id,note,memo\n7,\"a\tb\\c\",\"x\ny\"\n";

/// SHA-256 of the `CSV` file's documented stream, an unnamed sheet of
/// source text with the TAB, backslash, and LF escaped:
///
/// ```text
/// wxf1
/// S<TAB>
/// V<TAB>A1<TAB>s<TAB>id
/// V<TAB>B1<TAB>s<TAB>note
/// V<TAB>C1<TAB>s<TAB>memo
/// V<TAB>A2<TAB>s<TAB>7
/// V<TAB>B2<TAB>s<TAB>a\tb\\c
/// V<TAB>C2<TAB>s<TAB>x\ny
/// ```
const CSV_FINGERPRINT: &str =
    "wxf1:583c9e37c1d5a5104e94ed97aeffc19fcabe00935c6f5827106d1ec4ea4a0339";

fn fixture(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures")
        .join(name)
}

/// A private directory, unique per test, removed on drop.
struct TempDir(PathBuf);

impl TempDir {
    fn new(label: &str) -> Self {
        let stamp = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .map(|elapsed| elapsed.as_nanos())
            .unwrap_or_default();
        let path = std::env::temp_dir().join(format!(
            "wolfxl-fingerprint-{label}-{}-{}-{stamp}",
            std::process::id(),
            NEXT_TEMP.fetch_add(1, Ordering::Relaxed)
        ));
        fs::create_dir(&path).expect("create temp dir");
        Self(path)
    }

    fn path(&self, name: &str) -> PathBuf {
        self.0.join(name)
    }
}

impl Drop for TempDir {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.0);
    }
}

struct Run {
    code: i32,
    stdout: String,
    stderr: String,
}

fn wolfxl(args: &[&str]) -> Run {
    let out = Command::cargo_bin("wolfxl")
        .expect("wolfxl bin built")
        .args(args)
        .output()
        .expect("run wolfxl");
    Run {
        code: out.status.code().expect("exit code"),
        stdout: String::from_utf8(out.stdout).expect("utf-8 stdout"),
        stderr: String::from_utf8(out.stderr).expect("utf-8 stderr"),
    }
}

fn fingerprint(path: &Path) -> String {
    let run = wolfxl(&["fingerprint", path.to_str().unwrap()]);
    assert_eq!(run.code, 0, "fingerprint {path:?}: stderr={}", run.stderr);
    let text = run.stdout.strip_suffix('\n').expect("trailing newline");
    let hex = text.strip_prefix("wxf1:").expect("wxf1 prefix");
    assert!(
        hex.len() == 64
            && hex
                .bytes()
                .all(|byte| matches!(byte, b'0'..=b'9' | b'a'..=b'f')),
        "{text}"
    );
    text.to_string()
}

fn verify(path: &Path, expected: &str) -> Run {
    wolfxl(&["verify", path.to_str().unwrap(), "--fingerprint", expected])
}

/// Every member of a ZIP package, in archive order.
fn members(path: &Path) -> Vec<(String, Vec<u8>)> {
    let mut archive = ZipArchive::new(fs::File::open(path).expect("open package")).expect("zip");
    let count = archive.len();
    (0..count)
        .map(|index| {
            let mut member = archive.by_index(index).expect("package member");
            let mut bytes = Vec::new();
            member.read_to_end(&mut bytes).expect("read member");
            (member.name().expect("member name").to_string(), bytes)
        })
        .collect()
}

fn write_zip<'a>(
    path: &Path,
    members: impl IntoIterator<Item = (&'a str, &'a [u8])>,
    method: CompressionMethod,
) {
    let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
    let options = SimpleFileOptions::default().compression_method(method);
    for (name, bytes) in members {
        writer.start_file(name, options).expect("start member");
        writer.write_all(bytes).expect("write member");
    }
    let bytes = writer.finish().expect("finish package").into_inner();
    fs::write(path, bytes).expect("write package");
}

/// A one-sheet workbook named `Data` whose worksheet XML is `sheet`.
fn workbook(directory: &TempDir, name: &str, sheet: &str) -> PathBuf {
    let path = directory.path(name);
    write_zip(
        &path,
        [
            ("[Content_Types].xml", CONTENT_TYPES.as_bytes()),
            ("_rels/.rels", ROOT_RELS.as_bytes()),
            ("xl/workbook.xml", WORKBOOK.as_bytes()),
            ("xl/_rels/workbook.xml.rels", WORKBOOK_RELS.as_bytes()),
            ("xl/styles.xml", STYLES.as_bytes()),
            ("xl/worksheets/sheet1.xml", sheet.as_bytes()),
        ],
        CompressionMethod::Deflated,
    );
    path
}

/// `text` with its single occurrence of `from` replaced by `to`.
fn edited(text: &str, from: &str, to: &str) -> String {
    assert_eq!(text.matches(from).count(), 1, "{from:?} occurs once");
    text.replace(from, to)
}

/// The fixture with members written in reverse order and stored uncompressed.
fn rezipped_fixture(directory: &TempDir, name: &str) -> PathBuf {
    let source = fixture("sample-financials.xlsx");
    let path = directory.path(name);
    let members = members(&source);
    assert!(
        members.iter().all(|(member, _)| !member.ends_with('/')),
        "fixture has only file members"
    );
    write_zip(
        &path,
        members
            .iter()
            .rev()
            .map(|(member, bytes)| (member.as_str(), bytes.as_slice())),
        CompressionMethod::Stored,
    );
    assert_ne!(
        fs::read(&source).unwrap(),
        fs::read(&path).unwrap(),
        "the rewritten package differs byte for byte"
    );
    path
}

#[test]
fn workbook_fingerprint_hashes_the_documented_stream() {
    let directory = TempDir::new("stream");
    let path = workbook(&directory, "book.xlsx", SHEET);
    assert_eq!(fingerprint(&path), SHEET_FINGERPRINT);
}

#[test]
fn json_output_reports_version_and_record_counts() {
    let directory = TempDir::new("json");
    let path = workbook(&directory, "book.xlsx", SHEET);
    let run = wolfxl(&["fingerprint", path.to_str().unwrap(), "--json"]);
    assert_eq!(run.code, 0, "stderr={}", run.stderr);
    let document: serde_json::Value = serde_json::from_str(&run.stdout).expect("JSON output");
    assert_eq!(
        document,
        serde_json::json!({
            "fingerprint": SHEET_FINGERPRINT,
            "version": "wxf1",
            "sheets": 1,
            "cells": 7,
            "formulas": 1,
        })
    );
}

#[test]
fn rezipped_copy_with_other_order_and_compression_keeps_its_fingerprint() {
    let directory = TempDir::new("rezip");
    let original = fingerprint(&fixture("sample-financials.xlsx"));
    let copy = rezipped_fixture(&directory, "copy.xlsx");
    assert_eq!(fingerprint(&copy), original);
}

#[test]
fn value_edit_in_the_fixture_changes_its_fingerprint() {
    let directory = TempDir::new("fixture-value");
    let copy = rezipped_fixture(&directory, "copy.xlsx");
    let edited_path = directory.path("edited.xlsx");
    let members = members(&copy);
    let rewritten: Vec<(String, Vec<u8>)> = members
        .into_iter()
        .map(|(name, bytes)| {
            if name == "xl/worksheets/sheet1.xml" {
                let xml = String::from_utf8(bytes).expect("utf-8 worksheet");
                let xml = edited(
                    &xml,
                    r#"<c r="B3" t="n"><v>420000</v>"#,
                    r#"<c r="B3" t="n"><v>420001</v>"#,
                );
                (name, xml.into_bytes())
            } else {
                (name, bytes)
            }
        })
        .collect();
    write_zip(
        &edited_path,
        rewritten
            .iter()
            .map(|(name, bytes)| (name.as_str(), bytes.as_slice())),
        CompressionMethod::Stored,
    );
    assert_ne!(fingerprint(&edited_path), fingerprint(&copy));
}

#[test]
fn value_edit_changes_the_fingerprint() {
    let directory = TempDir::new("value");
    let base = workbook(&directory, "base.xlsx", SHEET);
    let sheet = edited(SHEET, r#"<c r="B1"><v>3</v>"#, r#"<c r="B1"><v>4</v>"#);
    let changed = workbook(&directory, "changed.xlsx", &sheet);
    assert_ne!(fingerprint(&changed), fingerprint(&base));
}

#[test]
fn formula_edit_with_the_same_cached_value_changes_the_fingerprint() {
    let directory = TempDir::new("formula");
    let base = workbook(&directory, "base.xlsx", SHEET);
    // Only the formula text changes; the cached result stays 5.
    let sheet = edited(SHEET, "<f>A1+B1</f>", "<f>A1*B1</f>");
    let changed = workbook(&directory, "changed.xlsx", &sheet);
    assert_ne!(fingerprint(&changed), fingerprint(&base));
}

#[test]
fn number_format_change_keeps_the_fingerprint() {
    let directory = TempDir::new("style");
    let sheet = edited(SHEET, r#"<c r="A1">"#, r#"<c r="A1" s="1">"#);
    let sheet = edited(&sheet, r#"<c r="G1">"#, r#"<c r="G1" s="1">"#);
    let styled = workbook(&directory, "styled.xlsx", &sheet);
    assert_eq!(fingerprint(&styled), SHEET_FINGERPRINT);
}

#[test]
fn csv_fingerprint_hashes_source_text_with_escapes() {
    let directory = TempDir::new("csv");
    let path = directory.path("export.csv");
    fs::write(&path, CSV).unwrap();
    assert_eq!(fingerprint(&path), CSV_FINGERPRINT);
}

#[test]
fn renamed_csv_keeps_its_fingerprint_and_text_edits_change_it() {
    let directory = TempDir::new("csv-rename");
    let first = directory.path("first.csv");
    let second = directory.path("second-name.csv");
    fs::write(&first, CSV).unwrap();
    fs::write(&second, CSV).unwrap();
    assert_eq!(fingerprint(&second), fingerprint(&first));

    // Delimited fields stay source text, so `7.0` is not `7`.
    let respelled = directory.path("respelled.csv");
    fs::write(&respelled, CSV.replace("\n7,", "\n7.0,")).unwrap();
    assert_ne!(fingerprint(&respelled), fingerprint(&first));
}

#[test]
fn verify_match_exits_zero() {
    let directory = TempDir::new("verify-match");
    let path = workbook(&directory, "book.xlsx", SHEET);
    let run = verify(&path, SHEET_FINGERPRINT);
    assert_eq!(run.code, 0, "stderr={}", run.stderr);
    assert_eq!(run.stdout, format!("ok {SHEET_FINGERPRINT}\n"));
}

#[test]
fn verify_accepts_uppercase_hex_and_surrounding_whitespace() {
    let directory = TempDir::new("verify-case");
    let path = directory.path("export.csv");
    fs::write(&path, CSV).unwrap();
    let digest = CSV_FINGERPRINT.strip_prefix("wxf1:").unwrap();
    let run = verify(&path, &format!(" wxf1:{}\n", digest.to_ascii_uppercase()));
    assert_eq!(run.code, 0, "stderr={}", run.stderr);
    assert_eq!(run.stdout, format!("ok {CSV_FINGERPRINT}\n"));
}

#[test]
fn verify_mismatch_exits_one_and_prints_both_fingerprints() {
    let directory = TempDir::new("verify-mismatch");
    let sheet = edited(SHEET, r#"<c r="B1"><v>3</v>"#, r#"<c r="B1"><v>4</v>"#);
    let changed = workbook(&directory, "changed.xlsx", &sheet);
    let actual = fingerprint(&changed);
    let run = verify(&changed, SHEET_FINGERPRINT);
    assert_eq!(run.code, 1, "stderr={}", run.stderr);
    assert_eq!(
        run.stdout,
        format!("mismatch\nexpected {SHEET_FINGERPRINT}\nactual {actual}\n")
    );
}

#[test]
fn verify_rejects_unknown_versions_and_malformed_digests_with_exit_two() {
    let directory = TempDir::new("verify-version");
    let path = workbook(&directory, "book.xlsx", SHEET);
    let digest = SHEET_FINGERPRINT.strip_prefix("wxf1:").unwrap();
    for expected in [
        format!("wxf2:{digest}"),
        format!("sha256:{digest}"),
        digest.to_string(),
        "wxf1:abc".to_string(),
        format!("wxf1:{}", "z".repeat(64)),
    ] {
        let run = verify(&path, &expected);
        assert_eq!(run.code, 2, "{expected}: stdout={}", run.stdout);
        assert!(
            run.stderr.contains("unknown fingerprint version"),
            "{expected}: stderr={}",
            run.stderr
        );
        assert!(run.stdout.is_empty(), "{expected}: stdout={}", run.stdout);
    }
}

#[test]
fn unreadable_input_exits_two() {
    let directory = TempDir::new("missing");
    let missing = directory.path("missing.xlsx");
    let run = verify(&missing, SHEET_FINGERPRINT);
    assert_eq!(run.code, 2);
    assert!(run.stderr.starts_with("error: "), "stderr={}", run.stderr);

    let run = wolfxl(&["fingerprint", missing.to_str().unwrap()]);
    assert_eq!(run.code, 2);
    assert!(run.stderr.starts_with("error: "), "stderr={}", run.stderr);
}
