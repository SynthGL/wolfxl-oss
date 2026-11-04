//! Integration tests for `wolfxl pivot`: grouping, key order, aggregations,
//! totals, output shapes, input sources, and the 0 / 2 exit-code contract.

use std::path::{Path, PathBuf};

use assert_cmd::Command;

/// Store keys 10, 9, 100, and 7 prove numeric key order; `abc` and `x` are
/// counted but not summed; the empty field is neither.
const STORES: &str = "Store,Amount\n10,4\n9,1\n100,2\n10,6\n9,abc\n100,\n7,x\n";

/// A blank Region, a 2024 Amount of 2.5, and no blank-Region 2024 row.
const SALES: &str = "Region,Year,Amount\n\
East,2024,10\n\
West,2024,5\n\
East,2023,20\n\
West,2023,7\n\
East,2024,2.5\n\
,2023,1\n";

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

fn pivot(args: &[&str], stdin: Option<&str>) -> Run {
    let mut command = Command::cargo_bin("wolfxl").expect("wolfxl bin built");
    command.arg("pivot").args(args);
    if let Some(stdin) = stdin {
        command.write_stdin(stdin);
    }
    let out = command.output().expect("run wolfxl pivot");
    Run {
        code: out.status.code().expect("exit code"),
        stdout: String::from_utf8(out.stdout).expect("utf-8 stdout"),
        stderr: String::from_utf8(out.stderr).expect("utf-8 stderr"),
    }
}

/// Pivot piped CSV and return standard output, asserting exit 0.
fn piped(csv: &str, args: &[&str]) -> String {
    let mut full = vec!["-", "--from", "csv"];
    full.extend_from_slice(args);
    let run = pivot(&full, Some(csv));
    assert_eq!(run.code, 0, "wolfxl pivot {full:?}: stderr={}", run.stderr);
    run.stdout
}

fn refused(args: &[&str], stdin: Option<&str>) -> String {
    let run = pivot(args, stdin);
    assert_eq!(run.code, 2, "wolfxl pivot {args:?}: stdout={}", run.stdout);
    assert!(run.stdout.is_empty(), "{}", run.stdout);
    run.stderr
}

#[test]
fn rows_only_pivot_applies_every_aggregation_with_numeric_key_order() {
    let out = piped(
        STORES,
        &[
            "--rows",
            "Store",
            "--values",
            "sum:Amount,count:Amount,avg:Amount,min:Amount,max:Amount",
            "-e",
            "csv",
        ],
    );
    assert_eq!(
        out,
        "Store,sum:Amount,count:Amount,avg:Amount,min:Amount,max:Amount\n\
         7,,1,,,\n\
         9,1,2,1,1,1\n\
         10,10,2,5,4,6\n\
         100,2,1,2,2,2\n\
         Total,13,6,3.25,1,6\n"
    );
}

#[test]
fn rows_only_json_reports_null_for_aggregates_without_numbers() {
    let out = piped(
        STORES,
        &[
            "--rows",
            "Store",
            "--values",
            "sum:Amount,count:Amount,avg:Amount",
            "-e",
            "json",
        ],
    );
    let document: serde_json::Value = serde_json::from_str(&out).expect("pivot JSON parses");
    assert_eq!(
        document,
        serde_json::json!({
            "contract_version": 1,
            "sheet": "stdin",
            "rows": ["Store"],
            "column": null,
            "values": ["sum:Amount", "count:Amount", "avg:Amount"],
            "column_keys": [],
            "headers": ["Store", "sum:Amount", "count:Amount", "avg:Amount"],
            "data": [
                [7, null, 1, null],
                [9, 1, 2, 1],
                [10, 10, 2, 5],
                [100, 2, 1, 2]
            ],
            "totals": [13, 6, 3.25]
        })
    );
}

#[test]
fn column_pivot_adds_grand_total_row_and_column_as_csv() {
    let out = piped(
        SALES,
        &[
            "--rows",
            "Region",
            "--column",
            "Year",
            "--values",
            "sum:Amount",
            "-e",
            "csv",
        ],
    );
    assert_eq!(
        out,
        "Region,2023,2024,Total\n\
         East,20,12.50,32.50\n\
         West,7,5,12\n\
         (blank),1,,1\n\
         Total,28,17.50,45.50\n"
    );
}

#[test]
fn column_pivot_json_has_a_stable_shape() {
    let out = piped(
        SALES,
        &[
            "--rows",
            "Region",
            "--column",
            "Year",
            "--values",
            "sum:Amount",
            "-e",
            "json",
        ],
    );
    assert_eq!(
        out,
        r#"{
  "contract_version": 1,
  "sheet": "stdin",
  "rows": ["Region"],
  "column": "Year",
  "values": ["sum:Amount"],
  "column_keys": [2023, 2024],
  "headers": ["Region", "2023", "2024", "Total"],
  "data": [
    ["East", 20, 12.5, 32.5],
    ["West", 7, 5, 12],
    ["(blank)", 1, null, 1]
  ],
  "totals": [28, 17.5, 45.5]
}
"#
    );
}

#[test]
fn multiple_values_get_one_column_per_column_key() {
    let out = piped(
        SALES,
        &[
            "--rows",
            "Region",
            "--column",
            "Year",
            "--values",
            "sum:Amount,count:Amount",
            "-e",
            "csv",
        ],
    );
    assert_eq!(
        out,
        "Region,2023 sum:Amount,2023 count:Amount,2024 sum:Amount,2024 count:Amount,\
         Total sum:Amount,Total count:Amount\n\
         East,20,1,12.50,2,32.50,3\n\
         West,7,1,5,1,12,2\n\
         (blank),1,1,,0,1,1\n\
         Total,28,3,17.50,3,45.50,6\n"
    );
}

#[test]
fn several_row_fields_sort_field_by_field() {
    let out = piped(
        SALES,
        &[
            "--rows",
            "Region,Year",
            "--values",
            "sum:Amount",
            "-e",
            "csv",
        ],
    );
    assert_eq!(
        out,
        "Region,Year,sum:Amount\n\
         East,2023,20\n\
         East,2024,12.50\n\
         West,2023,7\n\
         West,2024,5\n\
         (blank),2023,1\n\
         Total,,45.50\n"
    );
    // Repeating --rows is the same as a comma-separated list.
    let repeated = piped(
        SALES,
        &[
            "--rows",
            "Region",
            "--rows",
            "Year",
            "--values",
            "sum:Amount",
            "-e",
            "csv",
        ],
    );
    assert_eq!(repeated, out);
}

#[test]
fn headers_match_case_insensitively_when_unique() {
    let out = piped(
        SALES,
        &["--rows", "region", "--values", "MAX:amount", "-e", "csv"],
    );
    assert_eq!(
        out,
        "Region,max:Amount\nEast,20\nWest,7\n(blank),1\nTotal,20\n"
    );
}

#[test]
fn file_input_reads_a_csv_path_and_groups_thousands() {
    let path = fixture("sample-minimal.csv");
    let run = pivot(
        &[
            path.to_str().unwrap(),
            "--rows",
            "Account",
            "--values",
            "sum:Q1",
            "-e",
            "csv",
        ],
        None,
    );
    assert_eq!(run.code, 0, "{}", run.stderr);
    assert_eq!(
        run.stdout,
        "Account,sum:Q1\n\
         COGS,\"901,500\"\n\
         Gross,\"475,500\"\n\
         Revenue,\"1,377,000\"\n\
         Total,\"2,754,000\"\n"
    );
}

#[test]
fn piped_tsv_and_default_box_output() {
    let tsv = "Region\tAmount\nEast\t3\nWest\t4\nEast\t5\n";
    let run = pivot(
        &[
            "-",
            "--from",
            "tsv",
            "--rows",
            "Region",
            "--values",
            "sum:Amount",
        ],
        Some(tsv),
    );
    assert_eq!(run.code, 0, "{}", run.stderr);
    assert!(
        run.stdout
            .contains("Sheet: stdin (pivot) (3 rows × 2 columns)"),
        "{}",
        run.stdout
    );
    // Column widths are the widest cell: "Region" (6) and "sum:Amount" (10).
    for (key, sum) in [("East", "8"), ("West", "4"), ("Total", "12")] {
        let line = format!("│ {key:<6} │ {sum:<10} │");
        assert!(
            run.stdout.contains(&line),
            "missing {line:?} in\n{}",
            run.stdout
        );
    }
}

#[test]
fn unknown_headers_and_aggregations_exit_two() {
    let stderr = refused(
        &[
            "-",
            "--from",
            "csv",
            "--rows",
            "Nope",
            "--values",
            "sum:Amount",
        ],
        Some(SALES),
    );
    assert!(stderr.contains("unknown header \"Nope\""), "{stderr}");
    assert!(stderr.contains("Region, Year, Amount"), "{stderr}");

    let stderr = refused(
        &[
            "-",
            "--from",
            "csv",
            "--rows",
            "Region",
            "--column",
            "Month",
            "--values",
            "sum:Amount",
        ],
        Some(SALES),
    );
    assert!(stderr.contains("unknown header \"Month\""), "{stderr}");

    let stderr = refused(
        &[
            "-",
            "--from",
            "csv",
            "--rows",
            "Region",
            "--values",
            "sum:Price",
        ],
        Some(SALES),
    );
    assert!(stderr.contains("unknown header \"Price\""), "{stderr}");

    let stderr = refused(
        &[
            "-",
            "--from",
            "csv",
            "--rows",
            "Region",
            "--values",
            "median:Amount",
        ],
        Some(SALES),
    );
    assert!(
        stderr.contains("unknown aggregation \"median\""),
        "{stderr}"
    );
}

#[test]
fn malformed_and_ambiguous_requests_exit_two() {
    let stderr = refused(
        &[
            "-", "--from", "csv", "--rows", "Region", "--values", "Amount",
        ],
        Some(SALES),
    );
    assert!(stderr.contains("malformed --values entry"), "{stderr}");

    let stderr = refused(
        &[
            "-",
            "--from",
            "csv",
            "--rows",
            "region",
            "--values",
            "count:Amount",
        ],
        Some("Region,REGION,Amount\na,b,1\n"),
    );
    assert!(stderr.contains("use the exact header"), "{stderr}");

    let stderr = refused(
        &["-", "--rows", "Region", "--values", "sum:Amount"],
        Some(SALES),
    );
    assert!(stderr.contains("--from"), "{stderr}");

    let stderr = refused(&["-", "--from", "csv", "--rows", "Region"], Some(SALES));
    assert!(stderr.contains("--values"), "{stderr}");
}
