//! Integration tests for `wolfxl peek --where/--select` and piped input:
//! every operator, lenient numerics, AND composition, projection order,
//! standard input, and the exit-2 refusals.

use std::path::{Path, PathBuf};

use assert_cmd::Command;

/// Amounts spelled the way exported ledgers spell them, plus one non-number.
const LEDGER: &str = "\
Region,Rep,Amount,Status
West,Ann,\"$1,200.00\",Open
East,bob,950,closed
west,Cy,(500.00),Open
North,Dee,\"1,200\",Pending
South,Eve,n/a,open
";

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

fn peek(args: &[&str], stdin: Option<&str>) -> Run {
    let mut command = Command::cargo_bin("wolfxl").expect("wolfxl bin built");
    command.arg("peek").args(args);
    if let Some(stdin) = stdin {
        command.write_stdin(stdin);
    }
    let out = command.output().expect("run wolfxl peek");
    Run {
        code: out.status.code().expect("exit code"),
        stdout: String::from_utf8(out.stdout).expect("utf-8 stdout"),
        stderr: String::from_utf8(out.stderr).expect("utf-8 stderr"),
    }
}

/// Pipe `stdin` through `peek - --from csv`, appending `args`; expect exit 0.
fn piped(stdin: &str, args: &[&str]) -> String {
    let mut full = vec!["-", "--from", "csv"];
    full.extend_from_slice(args);
    let run = peek(&full, Some(stdin));
    assert_eq!(run.code, 0, "wolfxl peek {full:?}: stderr={}", run.stderr);
    run.stdout
}

/// The ledger as CSV, filtered by one `--where` expression.
fn ledger_where(expression: &str) -> String {
    piped(LEDGER, &["-e", "csv", "--where", expression])
}

/// Pipe `stdin` with `args`; expect exit 2 with nothing on stdout.
fn refused(stdin: &str, args: &[&str]) -> String {
    let mut full = vec!["-", "--from", "csv", "-e", "csv"];
    full.extend_from_slice(args);
    let run = peek(&full, Some(stdin));
    assert_eq!(run.code, 2, "wolfxl peek {full:?}: stdout={}", run.stdout);
    assert_eq!(run.stdout, "");
    assert!(run.stderr.starts_with("error: "), "stderr={}", run.stderr);
    run.stderr
}

#[test]
fn piped_csv_without_filters_renders_every_row_in_source_spelling() {
    assert_eq!(piped(LEDGER, &["-e", "csv"]), LEDGER);
}

#[test]
fn numeric_equals_reads_money_and_grouped_spellings() {
    let expected = "\
Region,Rep,Amount,Status
West,Ann,\"$1,200.00\",Open
North,Dee,\"1,200\",Pending
";
    assert_eq!(ledger_where("Amount=1200"), expected);
    assert_eq!(ledger_where("Amount=$1,200.00"), expected);
}

#[test]
fn numeric_equals_reads_accounting_negatives_on_both_sides() {
    assert_eq!(
        ledger_where("Amount=(500)"),
        "Region,Rep,Amount,Status\nwest,Cy,(500.00),Open\n"
    );
    assert_eq!(
        ledger_where("Amount=-500"),
        "Region,Rep,Amount,Status\nwest,Cy,(500.00),Open\n"
    );
}

#[test]
fn numeric_not_equals_keeps_non_numeric_cells() {
    assert_eq!(
        ledger_where("Amount!=1200"),
        "\
Region,Rep,Amount,Status
East,bob,950,closed
west,Cy,(500.00),Open
South,Eve,n/a,open
"
    );
}

#[test]
fn numeric_less_than_and_greater_than_skip_non_numeric_cells() {
    assert_eq!(
        ledger_where("Amount<1000"),
        "\
Region,Rep,Amount,Status
East,bob,950,closed
west,Cy,(500.00),Open
"
    );
    assert_eq!(
        ledger_where("Amount>1000"),
        "\
Region,Rep,Amount,Status
West,Ann,\"$1,200.00\",Open
North,Dee,\"1,200\",Pending
"
    );
}

#[test]
fn numeric_bounds_include_the_boundary() {
    assert_eq!(
        ledger_where("Amount<=950"),
        "\
Region,Rep,Amount,Status
East,bob,950,closed
west,Cy,(500.00),Open
"
    );
    assert_eq!(
        ledger_where("Amount>=950"),
        "\
Region,Rep,Amount,Status
West,Ann,\"$1,200.00\",Open
East,bob,950,closed
North,Dee,\"1,200\",Pending
"
    );
}

#[test]
fn text_equals_ignores_case() {
    assert_eq!(
        ledger_where("Region=WEST"),
        "\
Region,Rep,Amount,Status
West,Ann,\"$1,200.00\",Open
west,Cy,(500.00),Open
"
    );
}

#[test]
fn text_ordering_is_lexical_and_ignores_case() {
    assert_eq!(
        ledger_where("Region<North"),
        "Region,Rep,Amount,Status\nEast,bob,950,closed\n"
    );
    assert_eq!(
        ledger_where("Rep>=C"),
        "\
Region,Rep,Amount,Status
west,Cy,(500.00),Open
North,Dee,\"1,200\",Pending
South,Eve,n/a,open
"
    );
}

#[test]
fn contains_matches_substrings_ignoring_case() {
    assert_eq!(
        ledger_where("Status~EN"),
        "\
Region,Rep,Amount,Status
West,Ann,\"$1,200.00\",Open
west,Cy,(500.00),Open
North,Dee,\"1,200\",Pending
South,Eve,n/a,open
"
    );
}

#[test]
fn contains_with_a_numeric_right_side_still_tests_text() {
    assert_eq!(
        ledger_where("Amount~1,2"),
        "\
Region,Rep,Amount,Status
West,Ann,\"$1,200.00\",Open
North,Dee,\"1,200\",Pending
"
    );
}

#[test]
fn repeated_where_flags_combine_with_and() {
    assert_eq!(
        piped(
            LEDGER,
            &["-e", "csv", "--where", "Region=west", "--where", "Amount>0"]
        ),
        "Region,Rep,Amount,Status\nWest,Ann,\"$1,200.00\",Open\n"
    );
}

#[test]
fn where_column_names_match_case_insensitively() {
    assert_eq!(
        ledger_where("status=CLOSED"),
        "Region,Rep,Amount,Status\nEast,bob,950,closed\n"
    );
}

#[test]
fn select_chooses_and_reorders_columns() {
    assert_eq!(
        piped(LEDGER, &["-e", "csv", "--select", "Amount,region"]),
        "\
Amount,Region
\"$1,200.00\",West
950,East
(500.00),west
\"1,200\",North
n/a,South
"
    );
}

#[test]
fn where_may_reference_columns_that_select_drops() {
    assert_eq!(
        piped(
            LEDGER,
            &["-e", "csv", "--where", "Status=open", "--select", "Rep"]
        ),
        "Rep\nAnn\nCy\nEve\n"
    );
}

#[test]
fn filter_and_projection_reach_text_and_markdown_exports() {
    let args = ["--where", "Amount>1000", "--select", "Rep,Amount"];
    let mut text = vec!["-e", "text"];
    text.extend_from_slice(&args);
    assert_eq!(
        piped(LEDGER, &text),
        "Rep\tAmount\nAnn\t$1,200.00\nDee\t1,200\n"
    );
    let mut markdown = vec!["-e", "markdown"];
    markdown.extend_from_slice(&args);
    assert_eq!(
        piped(LEDGER, &markdown),
        "| Rep | Amount |\n|---|---|\n| Ann | $1,200.00 |\n| Dee | 1,200 |\n"
    );
}

#[test]
fn filter_and_projection_reach_json_export() {
    let out = piped(
        LEDGER,
        &[
            "-e",
            "json",
            "--where",
            "Amount>1000",
            "--select",
            "Rep,Amount",
        ],
    );
    let document: serde_json::Value = serde_json::from_str(&out).expect("json parses");
    assert_eq!(document["rows"], 2);
    assert_eq!(document["columns"], 2);
    assert_eq!(document["headers"], serde_json::json!(["Rep", "Amount"]));
    assert_eq!(
        document["data"],
        serde_json::json!([["Ann", "$1,200.00"], ["Dee", "1,200"]])
    );
}

#[test]
fn filtering_happens_before_the_box_row_cap() {
    let out = piped(LEDGER, &["--where", "Region=west", "-n", "1"]);
    assert!(out.contains("(2 rows × 4 columns)"), "{out}");
    assert!(out.contains("Showing 1 of 2 rows"), "{out}");
    assert!(out.contains("Ann"), "{out}");
    assert!(!out.contains("Cy"), "{out}");
}

#[test]
fn filters_apply_to_a_csv_file_path() {
    let path = fixture("sample-minimal.csv");
    let run = peek(
        &[path.to_str().unwrap(), "-e", "csv", "--where", "Q1>1000000"],
        None,
    );
    assert_eq!(run.code, 0, "stderr={}", run.stderr);
    assert_eq!(
        run.stdout,
        "Account,Jan,Feb,Mar,Q1\nRevenue,420000,445000,512000,1377000\n"
    );
}

#[test]
fn exact_header_match_wins_over_case_insensitive_ambiguity() {
    assert_eq!(
        piped("id,ID\n1,2\n3,4\n", &["-e", "csv", "--where", "ID=4"]),
        "id,ID\n3,4\n"
    );
    let stderr = refused("id,ID\n1,2\n", &["--where", "Id=1"]);
    assert!(stderr.contains("ambiguous column \"Id\""), "{stderr}");
}

#[test]
fn numeric_looking_headers_resolve_by_source_spelling() {
    // Typed, the headers `001`, `01`, and `1` would all read as 1.
    let csv = "001,01,1\n007,TRUE,1.5\n7,x,2\n8,FALSE,5\n";
    assert_eq!(
        piped(csv, &["-e", "csv", "--where", "1<3", "--select", "01,001"]),
        "01,001\nTRUE,007\nx,7\n"
    );

    // TSV: `1.0` and `1` stay distinct headers.
    let run = peek(
        &[
            "-", "--from", "tsv", "-e", "csv", "--where", "1=7", "--select", "1.0",
        ],
        Some("1\t1.0\n5\t6\n7\t8\n"),
    );
    assert_eq!(run.code, 0, "stderr={}", run.stderr);
    assert_eq!(run.stdout, "1.0\n8\n");
}

#[test]
fn unknown_where_column_exits_2_naming_the_headers() {
    let stderr = refused(LEDGER, &["--where", "Regin=West"]);
    assert!(stderr.contains("unknown column \"Regin\""), "{stderr}");
    assert!(
        stderr.contains("available headers: Region, Rep, Amount, Status"),
        "{stderr}"
    );
}

#[test]
fn unknown_select_column_exits_2_naming_the_headers() {
    let stderr = refused(LEDGER, &["--select", "Rep,Total"]);
    assert!(stderr.contains("unknown column \"Total\""), "{stderr}");
    assert!(
        stderr.contains("available headers: Region, Rep, Amount, Status"),
        "{stderr}"
    );
}

#[test]
fn malformed_where_expressions_exit_2() {
    let stderr = refused(LEDGER, &["--where", "Region"]);
    assert!(stderr.contains("expected COLUMN<op>VALUE"), "{stderr}");
    let stderr = refused(LEDGER, &["--where", "=West"]);
    assert!(stderr.contains("missing the column name"), "{stderr}");
    let stderr = refused(LEDGER, &["--select", "Rep,,Amount"]);
    assert!(stderr.contains("empty column name"), "{stderr}");
}

#[test]
fn standard_input_without_from_exits_2() {
    let run = peek(&["-", "-e", "csv"], Some(LEDGER));
    assert_eq!(run.code, 2, "stdout={}", run.stdout);
    assert!(run.stderr.contains("--from"), "stderr={}", run.stderr);
}

#[test]
fn unknown_sheet_still_exits_1() {
    let path = fixture("sample-financials.xlsx");
    let run = peek(&[path.to_str().unwrap(), "-s", "Nope"], None);
    assert_eq!(run.code, 1);
    assert!(run.stderr.contains("not found"), "stderr={}", run.stderr);
}
