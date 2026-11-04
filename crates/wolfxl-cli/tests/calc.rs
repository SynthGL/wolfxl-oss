//! Stored-value evaluation, delimited typing, arrays, and the 0 / 1 / 2 contract.

use std::io::{Cursor, Write};
use std::path::{Path, PathBuf};

use assert_cmd::Command;
use zip::write::SimpleFileOptions;
use zip::ZipWriter;

mod xlsb;

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

fn calc(args: &[&str], stdin: Option<&str>) -> Run {
    let mut command = Command::cargo_bin("wolfxl").expect("wolfxl bin built");
    command.arg("calc").args(args);
    if let Some(stdin) = stdin {
        command.write_stdin(stdin);
    }
    let out = command.output().expect("run wolfxl calc");
    Run {
        code: out.status.code().expect("exit code"),
        stdout: String::from_utf8(out.stdout).expect("utf-8 stdout"),
        stderr: String::from_utf8(out.stderr).expect("utf-8 stderr"),
    }
}

fn value(args: &[&str], stdin: Option<&str>) -> String {
    let run = calc(args, stdin);
    assert_eq!(run.code, 0, "wolfxl calc {args:?}: stderr={}", run.stderr);
    run.stdout
}

fn json(args: &[&str], stdin: Option<&str>) -> serde_json::Value {
    let mut args = args.to_vec();
    args.push("--json");
    serde_json::from_str(&value(&args, stdin)).expect("calc JSON parses")
}

/// Text output of a formula over workbook bytes piped as `--from format`.
fn piped(formula: &str, format: &str, bytes: Vec<u8>) -> String {
    let out = Command::cargo_bin("wolfxl")
        .expect("wolfxl bin built")
        .args(["calc", formula, "--from", format])
        .write_stdin(bytes)
        .output()
        .expect("run wolfxl calc");
    let stderr = String::from_utf8_lossy(&out.stderr);
    assert_eq!(
        out.status.code(),
        Some(0),
        "wolfxl calc {formula}: stderr={stderr}"
    );
    String::from_utf8(out.stdout).expect("utf-8 stdout")
}

/// A date-time serial whose fraction of a second a whole-second
/// conversion would drop.
const SERIAL: f64 = 45_000.123_456_789;

/// An XLSX workbook in the 1904 date system whose sheet `Data` stores
/// `SERIAL` in B2 under the built-in `m/d/yy h:mm` format and in C2 as General.
fn xlsx_1904() -> Vec<u8> {
    let types = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>"#;
    let root = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#;
    let workbook = r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><workbookPr date1904="1"/><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"#;
    let rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#;
    let styles = r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="22" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs></styleSheet>"#;
    let sheet = format!(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="2"><c r="B2" s="1"><v>{SERIAL}</v></c><c r="C2"><v>{SERIAL}</v></c></row></sheetData></worksheet>"#
    );
    let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
    for (name, text) in [
        ("[Content_Types].xml", types),
        ("_rels/.rels", root),
        ("xl/workbook.xml", workbook),
        ("xl/_rels/workbook.xml.rels", rels),
        ("xl/styles.xml", styles),
        ("xl/worksheets/sheet1.xml", sheet.as_str()),
    ] {
        writer
            .start_file(name, SimpleFileOptions::default())
            .expect("start member");
        writer.write_all(text.as_bytes()).expect("write member");
    }
    writer.finish().expect("finish package").into_inner()
}

#[test]
fn whole_column_sum_over_a_csv_file_types_numeric_fields() {
    let path = fixture("sample-minimal.csv");
    assert_eq!(
        value(&["=SUM(B:B)", path.to_str().unwrap()], None),
        "840000\n"
    );
}

#[test]
fn whole_column_sum_over_piped_csv() {
    let csv = std::fs::read_to_string(fixture("sample-minimal.csv")).unwrap();
    assert_eq!(
        value(&["SUM(B:B)", "--from", "csv"], Some(&csv)),
        "840000\n"
    );
    assert_eq!(
        value(&["SUM(E:E)", "-", "--from", "csv"], Some(&csv)),
        "2754000\n"
    );
}

#[test]
fn piped_tsv_splits_on_tabs() {
    let tsv = "item\tqty\tprice\nbolt\t3\t0.1\nnut\t4\t0.2\n";
    assert_eq!(
        value(&["=B2*C2+B3*C3", "--from", "tsv"], Some(tsv)),
        "1.1\n"
    );
}

#[test]
fn workbook_formulas_see_stored_values_of_the_first_sheet() {
    let path = fixture("sample-financials.xlsx");
    let path = path.to_str().unwrap();
    assert_eq!(value(&["=SUM(B3:B5)", path], None), "692000\n");
    assert_eq!(value(&["=SUM(B3:B5)=B6", path], None), "TRUE\n");
}

#[test]
fn named_sheet_resolves_unqualified_and_cross_sheet_references() {
    let path = fixture("sample-financials.xlsx");
    let document = json(
        &[
            "='P&L'!B3*2",
            path.to_str().unwrap(),
            "--sheet",
            "Balance Sheet",
        ],
        None,
    );
    assert_eq!(
        document,
        serde_json::json!({
            "formula": "='P&L'!B3*2", "sheet": "Balance Sheet", "type": "number", "value": 840000
        })
    );
    let unknown = calc(&["=1", path.to_str().unwrap(), "--sheet", "Nope"], None);
    assert_eq!(unknown.code, 2);
    assert!(
        unknown.stderr.contains("Balance Sheet"),
        "{}",
        unknown.stderr
    );
    assert_eq!(
        value(
            &["=B4", path.to_str().unwrap(), "--sheet", "Balance Sheet"],
            None
        ),
        "1842500\n"
    );
}

#[test]
fn excel_error_result_prints_the_error_and_exits_one() {
    let csv = "a\n0\n";
    let run = calc(&["=10/A2", "--from", "csv"], Some(csv));
    assert_eq!(
        (run.code, run.stdout.as_str()),
        (1, "#DIV/0!\n"),
        "{}",
        run.stderr
    );
    let run = calc(&["=10/A2", "--from", "csv", "--json"], Some(csv));
    assert_eq!(run.code, 1);
    let document: serde_json::Value = serde_json::from_str(&run.stdout).unwrap();
    assert_eq!(document["type"], "error");
    assert_eq!(document["value"], "#DIV/0!");
}

#[test]
fn parse_and_usage_failures_exit_two_with_a_message() {
    let run = calc(&["=SUM(1,", "--from", "csv"], Some("a\n1\n"));
    assert_eq!(run.code, 2);
    assert!(run.stdout.is_empty(), "{}", run.stdout);
    assert!(run.stderr.contains("cannot parse"), "{}", run.stderr);
    let run = calc(&["=1+1"], Some("a\n1\n"));
    assert_eq!(run.code, 2);
    assert!(run.stderr.contains("--from"), "{}", run.stderr);
    let run = calc(&["=1", "file-that-does-not-exist.xlsx"], None);
    assert_eq!(run.code, 2);
    assert!(run.stdout.is_empty());
}

#[test]
fn array_results_print_csv_rows_or_a_json_matrix() {
    let csv = std::fs::read_to_string(fixture("sample-minimal.csv")).unwrap();
    assert_eq!(
        value(&["=A2:B3", "--from", "csv"], Some(&csv)),
        "Revenue,420000\nCOGS,277000\n"
    );
    let document = json(&["=A2:B3", "--from", "csv"], Some(&csv));
    assert_eq!(document["type"], "array");
    assert_eq!(
        document["value"],
        serde_json::json!([["Revenue", 420000], ["COGS", 277000]])
    );
}

#[test]
fn unsupported_functions_exit_two_and_point_to_registry() {
    for formula in [
        "=SUMPRODUCT(B2:B3,C2:C3)",
        "=SEQUENCE(2,3)",
        "=XLOOKUP(1,A:A,B:B)",
    ] {
        let run = calc(&[formula, "--from", "csv"], Some("a,b,c\n1,2,3\n"));
        assert_eq!(run.code, 2, "{formula}: {}", run.stderr);
        assert!(run.stdout.is_empty());
        let function = formula.trim_start_matches('=').split('(').next().unwrap();
        assert!(run.stderr.contains(function), "{}", run.stderr);
        assert!(run.stderr.contains("unsupported"), "{}", run.stderr);
        assert!(
            run.stderr.contains("wolfxl list-functions"),
            "{}",
            run.stderr
        );
    }
}

#[test]
fn supported_functions_and_operators_use_typed_fields() {
    let csv = "a,b\n1,TRUE\n3,FALSE\n";
    for (formula, expected) in [
        ("AVERAGE(A2:A3)", "2\n"),
        ("COUNT(A:A)", "2\n"),
        ("MIN(A:A)", "1\n"),
        ("MAX(A:A)", "3\n"),
        ("IF(B2,A2+A3,0)", "4\n"),
        ("\"row\"&A3", "row3\n"),
        ("A2<A3", "TRUE\n"),
        ("(A2+A3)^2/4", "4\n"),
    ] {
        assert_eq!(
            value(&[formula, "--from", "csv"], Some(csv)),
            expected,
            "{formula}"
        );
    }
    let document = json(&["B2", "--from", "csv"], Some(csv));
    assert_eq!(document["type"], "boolean");
    assert_eq!(document["value"], true);
    let document = json(&["Z100", "--from", "csv"], Some(csv));
    assert_eq!(document["type"], "blank");
    assert!(document["value"].is_null());
}

#[test]
fn array_csv_quotes_fields_and_json_keeps_types_and_blanks() {
    let csv = "a,b,c\n\"a,b\",\"a\"\"b\",TRUE\n-0,0.3,\n";
    assert_eq!(
        value(&["A2:C3", "--from", "csv"], Some(csv)),
        "\"a,b\",\"a\"\"b\",TRUE\n0,0.3,\n"
    );
    assert_eq!(
        json(&["A2:C3", "--from", "csv"], Some(csv))["value"],
        serde_json::json!([["a,b", "a\"b", true], [0, 0.3, null]])
    );
}

#[test]
fn money_deals_csv_fields_are_numbers_in_formula_ranges() {
    let deals = "Rep,Cost,Revenue\nana,9000,\"$12,500.00\"\nbea,8000,\"$6,500.00\"\n";
    assert_eq!(
        value(&["SUM(C:C)", "--from", "csv"], Some(deals)),
        "19000\n"
    );
    assert_eq!(value(&["COUNT(C:C)", "--from", "csv"], Some(deals)), "2\n");
    assert_eq!(value(&["C2-B2", "--from", "csv"], Some(deals)), "3500\n");
}

#[test]
fn money_ledger_csv_accounting_negatives_participate_in_sums() {
    let ledger = "Amount\n100.00\n250.50\n(75.00)\n";
    assert_eq!(
        value(&["SUM(A:A)", "--from", "csv"], Some(ledger)),
        "275.5\n"
    );
    assert_eq!(value(&["COUNT(A:A)", "--from", "csv"], Some(ledger)), "3\n");
}

#[test]
fn xlsb_date_cells_evaluate_their_stored_1904_serials() {
    let book = || {
        xlsb::book(
            true,
            &[
                (2, 2, xlsb::DATE_TIME, SERIAL),
                (2, 3, xlsb::GENERAL, SERIAL),
            ],
        )
    };
    assert_eq!(piped("=B2", "xlsb", book()), "45000.123456789\n");
    assert_eq!(piped("=B2=C2", "xlsb", book()), "TRUE\n");
    assert_eq!(piped("=SUM(B:C)-2*C2", "xlsb", book()), "0\n");
}

#[test]
fn xlsx_date_cells_evaluate_their_stored_1904_serials() {
    assert_eq!(piped("=B2", "xlsx", xlsx_1904()), "45000.123456789\n");
    assert_eq!(piped("=B2=C2", "xlsx", xlsx_1904()), "TRUE\n");
}
