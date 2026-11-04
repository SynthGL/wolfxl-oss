//! Integration tests for `wolfxl stats`: typed CSV and XLSX profiles,
//! lenient numerics, blanks, top-value order and ties, `--where` and
//! `--select`, standard input, the box/csv/json layouts, and the exit-2
//! refusals.

use std::fs;
use std::io::{Cursor, Write};
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicUsize, Ordering};
use std::time::{SystemTime, UNIX_EPOCH};

use assert_cmd::Command;
use serde_json::{json, Value};
use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipWriter};

static NEXT_TEMP: AtomicUsize = AtomicUsize::new(0);

/// Eight amounts with mean -1000 and population deviation 2000, spelled as
/// ledgers spell them, plus one non-number, one empty field, and one field
/// of spaces.
const AMOUNTS: &str = "Amount\n\
\"(4,000.00)\"\n\
-2000\n\
\"-$2,000\"\n\
(2000)\n\
-1000\n\
-1000.0\n\
\"$1,000\"\n\
3000\n\
n/a\n\
\n\
\x20\x20\x20\n";

/// Nine fruit: apple, pear, and fig tie at two; kiwi, plum, and lime at one.
const FRUIT: &str = "Fruit\napple\npear\nfig\npear\napple\nkiwi\nfig\nplum\nlime\n";

const LEDGER: &str = "\
Region,Rep,Amount
West,Ann,\"$1,200.00\"
East,bob,950
west,Cy,(500.00)
North,Dee,\"1,200\"
";

const TABLE_HEADER: &str =
    "column,non_empty,blank,number,text,bool,date,error,distinct,numeric,sum,mean,min,max,stddev,top\n";

const CONTENT_TYPES: &str = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>"#;

const ROOT_RELS: &str = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#;

const WORKBOOK: &str = r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"#;

const WORKBOOK_RELS: &str = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#;

/// Cell format 1 applies the built-in date format (id 14).
const STYLES: &str = r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs></styleSheet>"#;

/// `Name` repeats a three times; `Value` holds the numbers 10 and 30, the
/// money text `$10`, a boolean, an error, the date 2023-03-15 (serial
/// 45000), a formula whose cached result is 30, and one missing cell.
const SHEET: &str = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Value</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>a</t></is></c><c r="B2"><v>10</v></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>b</t></is></c><c r="B3"><v>30</v></c></row><row r="4"><c r="A4" t="inlineStr"><is><t>a</t></is></c><c r="B4" t="inlineStr"><is><t>$10</t></is></c></row><row r="5"><c r="A5" t="inlineStr"><is><t>c</t></is></c><c r="B5" t="b"><v>1</v></c></row><row r="6"><c r="A6" t="inlineStr"><is><t>a</t></is></c><c r="B6" t="e"><v>#DIV/0!</v></c></row><row r="7"><c r="A7" t="inlineStr"><is><t>b</t></is></c><c r="B7" s="1"><v>45000</v></c></row><row r="8"><c r="A8" t="inlineStr"><is><t>c</t></is></c><c r="B8"><f>B2+20</f><v>30</v></c></row><row r="9"><c r="A9" t="inlineStr"><is><t>d</t></is></c></row></sheetData></worksheet>"#;

/// A private directory, unique per test, removed on drop.
struct TempDir(PathBuf);

impl TempDir {
    fn new(label: &str) -> Self {
        let stamp = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .map(|elapsed| elapsed.as_nanos())
            .unwrap_or_default();
        let path = std::env::temp_dir().join(format!(
            "wolfxl-stats-{label}-{}-{}-{stamp}",
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

/// The one-sheet `SHEET` workbook, written into `directory`.
fn workbook(directory: &TempDir) -> PathBuf {
    let path = directory.path("data.xlsx");
    let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
    let options = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
    for (name, text) in [
        ("[Content_Types].xml", CONTENT_TYPES),
        ("_rels/.rels", ROOT_RELS),
        ("xl/workbook.xml", WORKBOOK),
        ("xl/_rels/workbook.xml.rels", WORKBOOK_RELS),
        ("xl/styles.xml", STYLES),
        ("xl/worksheets/sheet1.xml", SHEET),
    ] {
        writer.start_file(name, options).expect("start member");
        writer.write_all(text.as_bytes()).expect("write member");
    }
    let bytes = writer.finish().expect("finish package").into_inner();
    fs::write(&path, bytes).expect("write package");
    path
}

fn fixture(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures")
        .join(name)
}

struct Run {
    code: i32,
    stdout: String,
    stderr: String,
}

fn stats(args: &[&str], stdin: Option<&str>) -> Run {
    let mut command = Command::cargo_bin("wolfxl").expect("wolfxl bin built");
    command.arg("stats").args(args);
    if let Some(stdin) = stdin {
        command.write_stdin(stdin);
    }
    let out = command.output().expect("run wolfxl stats");
    Run {
        code: out.status.code().expect("exit code"),
        stdout: String::from_utf8(out.stdout).expect("utf-8 stdout"),
        stderr: String::from_utf8(out.stderr).expect("utf-8 stderr"),
    }
}

/// Pipe `csv` through `stats - --from csv`, appending `args`; expect exit 0.
fn piped(csv: &str, args: &[&str]) -> String {
    let mut full = vec!["-", "--from", "csv"];
    full.extend_from_slice(args);
    let run = stats(&full, Some(csv));
    assert_eq!(run.code, 0, "wolfxl stats {full:?}: stderr={}", run.stderr);
    run.stdout
}

/// Piped CSV profiled as JSON, parsed.
fn piped_json(csv: &str, args: &[&str]) -> Value {
    let mut full = vec!["-e", "json"];
    full.extend_from_slice(args);
    serde_json::from_str(&piped(csv, &full)).expect("stats JSON parses")
}

/// Run `args` with `stdin`; expect exit 2, nothing on stdout, and an
/// `error: ` line on stderr.
fn refused(args: &[&str], stdin: Option<&str>) -> String {
    let run = stats(args, stdin);
    assert_eq!(run.code, 2, "wolfxl stats {args:?}: stdout={}", run.stdout);
    assert!(run.stdout.is_empty(), "{}", run.stdout);
    assert!(run.stderr.starts_with("error: "), "stderr={}", run.stderr);
    run.stderr
}

/// The JSON column object with zero counts and no numbers.
fn empty_profile(header: &str) -> Value {
    json!({
        "header": header,
        "non_empty": 0,
        "blank": 0,
        "types": {"number": 0, "text": 0, "bool": 0, "date": 0, "error": 0},
        "distinct": 0,
        "numeric": {"count": 0, "sum": null, "mean": null, "min": null, "max": null, "stddev": null},
        "top": []
    })
}

#[test]
fn lenient_amounts_feed_the_numeric_summary_and_blanks_are_counted() {
    assert_eq!(
        piped_json(AMOUNTS, &[]),
        json!({
            "contract_version": 1,
            "sheet": "stdin",
            "rows": 11,
            "columns": [{
                "header": "Amount",
                "non_empty": 9,
                "blank": 2,
                "types": {"number": 4, "text": 5, "bool": 0, "date": 0, "error": 0},
                "distinct": 9,
                "numeric": {
                    "count": 8,
                    "sum": -8000,
                    "mean": -1000,
                    "min": -4000,
                    "max": 3000,
                    "stddev": 2000
                },
                "top": [
                    {"value": "(4,000.00)", "count": 1},
                    {"value": "-2000", "count": 1},
                    {"value": "-$2,000", "count": 1},
                    {"value": "(2000)", "count": 1},
                    {"value": "-1000", "count": 1}
                ]
            }]
        })
    );
}

#[test]
fn typed_csv_counts_numbers_and_booleans_and_keeps_source_spelling() {
    let document = piped_json("Flag,Qty\nTRUE,1\nfalse,2.5\nyes,007\n", &[]);
    assert_eq!(
        document["columns"],
        json!([
            {
                "header": "Flag",
                "non_empty": 3,
                "blank": 0,
                "types": {"number": 0, "text": 1, "bool": 2, "date": 0, "error": 0},
                "distinct": 3,
                "numeric": {"count": 0, "sum": null, "mean": null, "min": null, "max": null, "stddev": null},
                "top": [
                    {"value": "TRUE", "count": 1},
                    {"value": "false", "count": 1},
                    {"value": "yes", "count": 1}
                ]
            },
            {
                "header": "Qty",
                "non_empty": 3,
                "blank": 0,
                "types": {"number": 3, "text": 0, "bool": 0, "date": 0, "error": 0},
                "distinct": 3,
                "numeric": {
                    "count": 3,
                    "sum": 10.5,
                    "mean": 3.5,
                    "min": 1,
                    "max": 7,
                    "stddev": 2.54950975679639
                },
                "top": [
                    {"value": "1", "count": 1},
                    {"value": "2.5", "count": 1},
                    {"value": "007", "count": 1}
                ]
            }
        ])
    );
}

#[test]
fn json_prints_one_column_object_per_line_in_sheet_order() {
    let out = piped("B,A\n1,x\n", &["-e", "json"]);
    let lines: Vec<&str> = out.lines().collect();
    assert_eq!(lines[0], "{");
    assert_eq!(lines[1], "  \"contract_version\": 1,");
    assert_eq!(lines[2], "  \"sheet\": \"stdin\",");
    assert_eq!(lines[3], "  \"rows\": 1,");
    assert_eq!(lines[4], "  \"columns\": [");
    assert!(lines[5].starts_with("    {\"header\": \"B\", "), "{out}");
    assert!(lines[5].ends_with("}]},"), "{out}");
    assert!(lines[6].starts_with("    {\"header\": \"A\", "), "{out}");
    assert_eq!(&lines[7..], ["  ]", "}"]);
    assert_eq!(piped("B,A\n1,x\n", &["-e", "json"]), out);
}

#[test]
fn top_values_order_by_count_then_first_appearance() {
    let out = piped(FRUIT, &["-e", "csv", "--top", "2"]);
    assert_eq!(
        out,
        format!("{TABLE_HEADER}Fruit,9,0,0,9,0,0,0,6,0,,,,,,apple (2); pear (2)\n")
    );
    let out = piped(FRUIT, &["-e", "csv"]);
    assert!(
        out.ends_with(",apple (2); pear (2); fig (2); kiwi (1); plum (1)\n"),
        "{out}"
    );
}

#[test]
fn top_zero_lists_every_distinct_value() {
    let out = piped(FRUIT, &["-e", "csv", "--top", "0"]);
    assert!(
        out.ends_with(",apple (2); pear (2); fig (2); kiwi (1); plum (1); lime (1)\n"),
        "{out}"
    );
}

#[test]
fn csv_layout_prints_fixed_stat_columns_without_float_noise() {
    let out = piped("Price\n0.1\n0.2\n", &["-e", "csv"]);
    assert_eq!(
        out,
        format!("{TABLE_HEADER}Price,2,0,2,0,0,0,0,2,2,0.3,0.15,0.1,0.2,0.05,0.1 (1); 0.2 (1)\n")
    );
}

#[test]
fn box_layout_is_the_default() {
    let out = piped(FRUIT, &["--top", "1"]);
    assert!(
        out.contains("Sheet: stdin (stats) (1 rows × 16 columns)"),
        "{out}"
    );
    for text in ["column", "stddev", "Fruit", "apple (2)"] {
        assert!(out.contains(text), "{text:?} missing from {out}");
    }
    assert!(!out.contains("pear (2)"), "{out}");
}

#[test]
fn xlsx_profile_counts_every_type() {
    let directory = TempDir::new("xlsx");
    let path = workbook(&directory);
    let run = stats(&[path.to_str().unwrap(), "-e", "json"], None);
    assert_eq!(run.code, 0, "stderr={}", run.stderr);
    let document: Value = serde_json::from_str(&run.stdout).expect("stats JSON parses");
    assert_eq!(
        document,
        json!({
            "contract_version": 1,
            "sheet": "Data",
            "rows": 8,
            "columns": [
                {
                    "header": "Name",
                    "non_empty": 8,
                    "blank": 0,
                    "types": {"number": 0, "text": 8, "bool": 0, "date": 0, "error": 0},
                    "distinct": 4,
                    "numeric": {"count": 0, "sum": null, "mean": null, "min": null, "max": null, "stddev": null},
                    "top": [
                        {"value": "a", "count": 3},
                        {"value": "b", "count": 2},
                        {"value": "c", "count": 2},
                        {"value": "d", "count": 1}
                    ]
                },
                {
                    "header": "Value",
                    "non_empty": 7,
                    "blank": 1,
                    "types": {"number": 3, "text": 1, "bool": 1, "date": 1, "error": 1},
                    "distinct": 6,
                    "numeric": {"count": 4, "sum": 80, "mean": 20, "min": 10, "max": 30, "stddev": 10},
                    "top": [
                        {"value": "30", "count": 2},
                        {"value": "10", "count": 1},
                        {"value": "$10", "count": 1},
                        {"value": "TRUE", "count": 1},
                        {"value": "#DIV/0!", "count": 1}
                    ]
                }
            ]
        })
    );
}

#[test]
fn xlsx_profile_reads_the_named_sheet_and_filters_numerically() {
    let directory = TempDir::new("xlsx-where");
    let path = workbook(&directory);
    let run = stats(
        &[
            path.to_str().unwrap(),
            "--sheet",
            "Data",
            "--where",
            "Value>15",
            "--select",
            "name",
            "--top",
            "0",
            "-e",
            "csv",
        ],
        None,
    );
    assert_eq!(run.code, 0, "stderr={}", run.stderr);
    assert_eq!(
        run.stdout,
        format!("{TABLE_HEADER}Name,2,0,0,2,0,0,0,2,0,,,,,,b (1); c (1)\n")
    );
}

#[test]
fn where_and_select_reshape_the_profile_like_peek() {
    let document = piped_json(
        LEDGER,
        &["--where", "Region=west", "--select", "Amount,Rep"],
    );
    assert_eq!(document["rows"], 2);
    let columns = document["columns"].as_array().expect("columns array");
    let headers: Vec<&str> = columns
        .iter()
        .map(|column| column["header"].as_str().unwrap_or_default())
        .collect();
    assert_eq!(headers, ["Amount", "Rep"]);
    assert_eq!(
        columns[0]["numeric"],
        json!({"count": 2, "sum": 700, "mean": 350, "min": -500, "max": 1200, "stddev": 850})
    );
    assert_eq!(
        columns[0]["types"],
        json!({"number": 0, "text": 2, "bool": 0, "date": 0, "error": 0})
    );
    assert_eq!(
        columns[1]["top"],
        json!([{"value": "Ann", "count": 1}, {"value": "Cy", "count": 1}])
    );
}

#[test]
fn numeric_looking_headers_resolve_by_source_spelling_over_typed_data() {
    // Typed, the headers `001`, `01`, and `1` all read as 1.
    let csv = "001,01,1\n007,TRUE,1.5\n7,x,2\n8,FALSE,5\n";
    let document = piped_json(csv, &["--where", "1<3", "--select", "01,001"]);
    assert_eq!(document["rows"], 2);
    assert_eq!(
        document["columns"],
        json!([
            {
                "header": "01",
                "non_empty": 2,
                "blank": 0,
                "types": {"number": 0, "text": 1, "bool": 1, "date": 0, "error": 0},
                "distinct": 2,
                "numeric": {"count": 0, "sum": null, "mean": null, "min": null, "max": null, "stddev": null},
                "top": [{"value": "TRUE", "count": 1}, {"value": "x", "count": 1}]
            },
            {
                "header": "001",
                "non_empty": 2,
                "blank": 0,
                "types": {"number": 2, "text": 0, "bool": 0, "date": 0, "error": 0},
                "distinct": 2,
                "numeric": {"count": 2, "sum": 14, "mean": 7, "min": 7, "max": 7, "stddev": 0},
                "top": [{"value": "007", "count": 1}, {"value": "7", "count": 1}]
            }
        ])
    );

    // TSV: `1.0` and `1` stay distinct headers.
    let run = stats(
        &["-", "--from", "tsv", "-e", "json", "--select", "1.0"],
        Some("1\t1.0\n5\t6\n7\t8\n"),
    );
    assert_eq!(run.code, 0, "stderr={}", run.stderr);
    let document: Value = serde_json::from_str(&run.stdout).expect("stats JSON parses");
    assert_eq!(document["columns"][0]["header"], "1.0");
    assert_eq!(
        document["columns"][0]["numeric"],
        json!({"count": 2, "sum": 14, "mean": 7, "min": 6, "max": 8, "stddev": 1})
    );
}

#[test]
fn repeated_where_flags_combine_with_and() {
    let document = piped_json(
        LEDGER,
        &[
            "--where",
            "Amount>=1,000",
            "--where",
            "Rep~e",
            "--select",
            "Rep",
        ],
    );
    assert_eq!(document["rows"], 1);
    assert_eq!(
        document["columns"][0]["top"],
        json!([{"value": "Dee", "count": 1}])
    );
}

#[test]
fn header_only_and_empty_input_profile_zero_rows() {
    assert_eq!(
        piped_json("A,B\n", &[]),
        json!({
            "contract_version": 1,
            "sheet": "stdin",
            "rows": 0,
            "columns": [empty_profile("A"), empty_profile("B")]
        })
    );
    let filtered = piped_json(LEDGER, &["--where", "Region=Nowhere", "--select", "Rep"]);
    assert_eq!(filtered["rows"], 0);
    assert_eq!(filtered["columns"], json!([empty_profile("Rep")]));
    assert_eq!(
        piped_json("", &[]),
        json!({"contract_version": 1, "sheet": "stdin", "rows": 0, "columns": []})
    );
    assert_eq!(piped("", &["-e", "csv"]), TABLE_HEADER);
}

#[test]
fn csv_file_path_reads_without_from() {
    let path = fixture("sample-minimal.csv");
    let run = stats(
        &[
            path.to_str().unwrap(),
            "--select",
            "Q1",
            "--top",
            "1",
            "-e",
            "csv",
        ],
        None,
    );
    assert_eq!(run.code, 0, "stderr={}", run.stderr);
    assert_eq!(
        run.stdout,
        format!(
            "{TABLE_HEADER}Q1,3,0,3,0,0,0,0,3,3,2754000,918000,475500,1377000,368220.72185036,1377000 (1)\n"
        )
    );
}

#[test]
fn unknown_where_column_exits_2_naming_the_headers() {
    let stderr = refused(
        &["-", "--from", "csv", "--where", "Regin=West"],
        Some(LEDGER),
    );
    assert!(stderr.contains("unknown column \"Regin\""), "{stderr}");
    assert!(
        stderr.contains("available headers: Region, Rep, Amount"),
        "{stderr}"
    );
}

#[test]
fn unknown_select_column_exits_2_naming_the_headers() {
    let stderr = refused(
        &["-", "--from", "csv", "--select", "Rep,Total"],
        Some(LEDGER),
    );
    assert!(stderr.contains("unknown column \"Total\""), "{stderr}");
    assert!(
        stderr.contains("available headers: Region, Rep, Amount"),
        "{stderr}"
    );
}

#[test]
fn ambiguous_and_malformed_requests_exit_2() {
    let stderr = refused(
        &["-", "--from", "csv", "--where", "Id=1"],
        Some("id,ID\n1,2\n"),
    );
    assert!(stderr.contains("ambiguous column \"Id\""), "{stderr}");
    let stderr = refused(&["-", "--from", "csv", "--where", "Region"], Some(LEDGER));
    assert!(stderr.contains("expected COLUMN<op>VALUE"), "{stderr}");
    let stderr = refused(&["-", "--from", "csv", "--where", "=West"], Some(LEDGER));
    assert!(stderr.contains("missing the column name"), "{stderr}");
    let stderr = refused(
        &["-", "--from", "csv", "--select", "Rep,,Amount"],
        Some(LEDGER),
    );
    assert!(stderr.contains("empty column name"), "{stderr}");
}

#[test]
fn input_failures_exit_2() {
    let stderr = refused(&["-"], Some(LEDGER));
    assert!(stderr.contains("--from"), "{stderr}");
    let directory = TempDir::new("missing");
    let missing = directory.path("missing.csv");
    refused(&[missing.to_str().unwrap()], None);
    let path = workbook(&directory);
    let stderr = refused(&[path.to_str().unwrap(), "--sheet", "Nope"], None);
    assert!(stderr.contains("Nope"), "{stderr}");
}
