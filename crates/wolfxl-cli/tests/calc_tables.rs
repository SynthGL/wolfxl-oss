//! Structured table references in `wolfxl calc` and `peek --add`, against a
//! real workbook's tables (`data-table-cities.xlsx`: `BigCity` at C2:F15 and
//! `SmallCity` at I2:L15 on sheet `Table`, each with a header and a totals
//! row, and cached structured-reference formulas on sheet `Formula`).

use std::path::{Path, PathBuf};

use assert_cmd::Command;

fn fixture(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures")
        .join(name)
}

fn cities() -> String {
    fixture("data-table-cities.xlsx")
        .to_str()
        .expect("utf-8 fixture path")
        .to_owned()
}

struct Run {
    code: i32,
    stdout: String,
    stderr: String,
}

fn wolfxl(args: &[&str], stdin: Option<&str>) -> Run {
    let mut command = Command::cargo_bin("wolfxl").expect("wolfxl bin built");
    command.args(args);
    if let Some(stdin) = stdin {
        command.write_stdin(stdin);
    }
    let out = command.output().expect("run wolfxl");
    Run {
        code: out.status.code().expect("exit code"),
        stdout: String::from_utf8(out.stdout).expect("utf-8 stdout"),
        stderr: String::from_utf8(out.stderr).expect("utf-8 stderr"),
    }
}

/// `calc formula` on the cities workbook, with extra arguments; the run must
/// succeed.
fn calc(formula: &str, extra: &[&str]) -> String {
    let file = cities();
    let mut args = vec!["calc", formula, file.as_str()];
    args.extend_from_slice(extra);
    let run = wolfxl(&args, None);
    assert_eq!(
        run.code, 0,
        "calc {formula} {extra:?}: stderr={}",
        run.stderr
    );
    run.stdout
}

#[test]
fn a_column_reference_is_the_data_body_without_the_totals_row() {
    // F15 is the totals row's SUBTOTAL; counting it would double the sum.
    assert_eq!(calc("SUM(BigCity[Population])", &[]), "198905700\n");
    assert_eq!(calc("=SUM(BigCity[[Population]])", &[]), "198905700\n");
    assert_eq!(
        calc("INDEX(BigCity[Population],12)", &[]),
        calc("F14", &["--sheet", "Table"])
    );
    assert_eq!(calc("INDEX(BigCity[Population],1)", &[]), "24256800\n");
    assert_eq!(calc("COUNTA(BigCity[City])", &[]), "12\n");
    // Table, column, and item names ignore case.
    assert_eq!(calc("SUM(bigcity[POPULATION])", &[]), "198905700\n");
}

#[test]
fn a_totals_reference_matches_the_workbooks_cached_result() {
    for (structured, cached) in [
        ("BigCity[[#Totals],[Latitude]]", "Formula!A5"),
        ("BigCity[[#Totals],[Longitude]]", "Formula!B5"),
        ("BigCity[[#Totals],[Population]]", "Formula!D5"),
        ("SmallCity[[#Totals],[Latitude]]", "Formula!A9"),
        ("SmallCity[[#totals],[Population]]", "Formula!D9"),
    ] {
        assert_eq!(calc(structured, &[]), calc(cached, &[]), "{structured}");
    }
    assert_eq!(
        calc("BigCity[[#Totals],[Latitude]]", &[]),
        "25.3966666666667\n"
    );
    assert_eq!(calc("BigCity[[#Totals],[Population]]", &[]), "198905700\n");
    // A one-cell structured reference is a scalar in arithmetic.
    assert_eq!(
        calc(
            "BigCity[[#Totals],[Population]]-SUM(BigCity[Population])",
            &[]
        ),
        "0\n"
    );
}

#[test]
fn headers_data_and_all_items_select_their_rows() {
    assert_eq!(calc("COUNTA(BigCity[[#Headers],[City]])", &[]), "1\n");
    assert_eq!(
        calc("BigCity[[#Headers],[Population]]", &[]),
        "Population\n"
    );
    assert_eq!(
        calc("BigCity[#Headers]", &[]),
        "City,Latitude,Longitude,Population\n"
    );
    assert_eq!(calc("COUNTA(BigCity[[#Data],[City]])", &[]), "12\n");
    assert_eq!(calc("COUNTA(BigCity[[#All],[City]])", &[]), "14\n");
    assert_eq!(
        calc("COUNTA(BigCity[[#Headers],[#Data],[City]])", &[]),
        "13\n"
    );
    assert_eq!(
        calc("SUM(BigCity[[#Data],[#Totals],[Population]])", &[]),
        "397811400\n"
    );
    assert_eq!(calc("COUNTA(BigCity[#All])", &[]), "56\n");
    assert_eq!(
        calc("SUM(BigCity[])", &[]),
        calc("SUM(D3:F14)", &["--sheet", "Table"])
    );
}

#[test]
fn a_column_range_covers_the_columns_between_its_ends() {
    let manual = calc("SUM(J3:L14)", &["--sheet", "Table"]);
    assert_eq!(
        calc("SUM(SmallCity[[#Data],[Latitude]:[Population]])", &[]),
        manual
    );
    assert_eq!(calc("SUM(SmallCity[[Population]:[Latitude]])", &[]), manual);
    assert_eq!(
        calc("COUNTA(SmallCity[[#Headers],[City]:[Longitude]])", &[]),
        "3\n"
    );
}

#[test]
fn table_references_are_workbook_global() {
    for formula in [
        "SUM(BigCity[Population])",
        "BigCity[[#Totals],[Latitude]]",
        "SUM(SmallCity[[#Data],[Latitude]:[Population]])",
    ] {
        let default = calc(formula, &[]);
        assert_eq!(calc(formula, &["--sheet", "Formula"]), default, "{formula}");
        assert_eq!(calc(formula, &["--sheet", "Table"]), default, "{formula}");
        assert_eq!(
            calc(formula, &["--sheet", "PivotTable"]),
            default,
            "{formula}"
        );
    }
    let json = calc(
        "SUM(BigCity[Population])",
        &["--sheet", "Formula", "--json"],
    );
    let json: serde_json::Value = serde_json::from_str(&json).expect("calc JSON");
    assert_eq!(json["sheet"], "Formula");
    assert_eq!(json["value"].as_f64(), Some(198905700.0));
}

#[test]
fn unknown_tables_and_columns_are_ref_errors() {
    let file = cities();
    for formula in [
        "SUM(NoSuchTable[x])",
        "SUM(BigCity[NoSuchColumn])",
        "SUM(BigCity[[City]:[NoSuchColumn]])",
    ] {
        let run = wolfxl(&["calc", formula, &file], None);
        assert_eq!(run.code, 1, "{formula}: stderr={}", run.stderr);
        assert_eq!(run.stdout, "#REF!\n", "{formula}");
    }
    // `@` needs the formula's own row, which a one-off formula lacks.
    let run = wolfxl(&["calc", "BigCity[@Population]", &file], None);
    assert_eq!((run.code, run.stdout.as_str()), (1, "#VALUE!\n"));
}

#[test]
fn malformed_structured_references_exit_two() {
    let file = cities();
    for (formula, message) in [
        ("SUM([Population])", "needs a table name"),
        ("SUM(BigCity[#Everything])", "unknown table item"),
        ("SUM(BigCity[[#Headers],[#Totals]])", "combine only"),
        ("SUM(BigCity[Population)", "unterminated"),
    ] {
        let run = wolfxl(&["calc", formula, &file], None);
        assert_eq!(run.code, 2, "{formula}: stdout={}", run.stdout);
        assert!(run.stderr.contains(message), "{formula}: {}", run.stderr);
    }
}

#[test]
fn inputs_without_tables_resolve_table_references_to_ref() {
    let run = wolfxl(
        &["calc", "SUM(Sales[Amount])", "-", "--from", "csv"],
        Some("Amount\n1\n2\n"),
    );
    assert_eq!(
        (run.code, run.stdout.as_str()),
        (1, "#REF!\n"),
        "{}",
        run.stderr
    );
    for name in [
        "sample-minimal.csv",
        "sample-minimal.ods",
        "sample-minimal.xls",
        "sample-financials.xlsx",
    ] {
        let path = fixture(name);
        let path = path.to_str().expect("utf-8 fixture path");
        let run = wolfxl(&["calc", "SUM(Sales[Amount])", path], None);
        assert_eq!(run.code, 1, "{name}: stderr={}", run.stderr);
        assert_eq!(run.stdout, "#REF!\n", "{name}");
        // Ordinary formulas still evaluate.
        assert_eq!(wolfxl(&["calc", "1+1", path], None).stdout, "2\n", "{name}");
    }
}

#[test]
fn added_peek_columns_resolve_table_references() {
    let file = cities();
    let run = wolfxl(
        &[
            "peek",
            &file,
            "--sheet",
            "Table",
            "-e",
            "json",
            "--add",
            "Share=F3/SUM(BigCity[Population])",
            "--add",
            "Total=SUM(BigCity[Population])",
        ],
        None,
    );
    assert_eq!(run.code, 0, "stderr={}", run.stderr);
    let document: serde_json::Value = serde_json::from_str(&run.stdout).expect("peek JSON");
    let rows = document["data"].as_array().expect("data rows");
    assert!(!rows.is_empty());
    for row in rows {
        let row = row.as_array().expect("row");
        assert_eq!(row[row.len() - 1], "198905700", "{row:?}");
    }
}

#[test]
fn added_peek_columns_accept_escaped_brackets_in_column_names() {
    // `'[` is a literal `[`, so the formula parses; BigCity has no such
    // column, so each added cell is #REF! rather than an exit-2 parse error.
    let file = cities();
    let run = wolfxl(
        &[
            "peek",
            &file,
            "--sheet",
            "Table",
            "-e",
            "json",
            "--add",
            "Cost=SUM(BigCity[Cost '[USD])",
        ],
        None,
    );
    assert_eq!(run.code, 0, "stderr={}", run.stderr);
    let document: serde_json::Value = serde_json::from_str(&run.stdout).expect("peek JSON");
    let rows = document["data"].as_array().expect("data rows");
    assert!(!rows.is_empty());
    for row in rows {
        let row = row.as_array().expect("row");
        assert_eq!(row[row.len() - 1], "#REF!", "{row:?}");
    }
}
