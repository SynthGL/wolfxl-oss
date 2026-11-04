//! Integration tests for `wolfxl diff`: keyed row matching, tolerance, key
//! transforms, input sources, and the 0 / 1 / 2 exit-code contract.

use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicUsize, Ordering};

use assert_cmd::Command;
use serde_json::{json, Value};

fn fixture(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures")
        .join(name)
}

/// A uniquely named file in the temp directory, removed on drop.
struct TempFile(PathBuf);

impl TempFile {
    fn new(name: &str, contents: &str) -> Self {
        static NEXT: AtomicUsize = AtomicUsize::new(0);
        let unique = NEXT.fetch_add(1, Ordering::Relaxed);
        let path = std::env::temp_dir().join(format!(
            "wolfxl-diff-test-{}-{unique}-{name}",
            std::process::id()
        ));
        std::fs::write(&path, contents).expect("write temp input");
        Self(path)
    }

    fn path(&self) -> &str {
        self.0.to_str().expect("utf-8 temp path")
    }
}

impl Drop for TempFile {
    fn drop(&mut self) {
        let _ = std::fs::remove_file(&self.0);
    }
}

struct Run {
    code: i32,
    stdout: String,
    stderr: String,
}

impl Run {
    fn json(&self) -> Value {
        serde_json::from_str(&self.stdout)
            .unwrap_or_else(|error| panic!("diff JSON parses ({error}): {}", self.stdout))
    }
}

fn diff(args: &[&str], stdin: Option<&str>) -> Run {
    let mut command = Command::cargo_bin("wolfxl").expect("wolfxl bin built");
    command.arg("diff").args(args);
    if let Some(stdin) = stdin {
        command.write_stdin(stdin);
    }
    let out = command.output().expect("run wolfxl diff");
    Run {
        code: out.status.code().expect("exit code"),
        stdout: String::from_utf8(out.stdout).expect("utf-8 stdout"),
        stderr: String::from_utf8(out.stderr).expect("utf-8 stderr"),
    }
}

/// Diff two inline tables written to temp files.
fn diff_tables(left: &str, right: &str, args: &[&str]) -> Run {
    let left = TempFile::new("left.csv", left);
    let right = TempFile::new("right.csv", right);
    let mut all = vec![left.path(), right.path()];
    all.extend_from_slice(args);
    diff(&all, None)
}

fn summary(matched: u64, only_left: u64, only_right: u64, diff: u64, outside: u64) -> Value {
    json!({
        "matched": matched,
        "only_left": only_left,
        "only_right": only_right,
        "diff": diff,
        "diff_outside_tolerance": outside,
    })
}

fn statuses(document: &Value) -> Vec<(String, String)> {
    document["results"]
        .as_array()
        .expect("results array")
        .iter()
        .map(|entry| {
            (
                entry["key"].as_str().expect("key").to_string(),
                entry["status"].as_str().expect("status").to_string(),
            )
        })
        .collect()
}

const LEDGER: &str = "id,name,amount\n1,Widget,100\n2,Gadget,250.50\n3,Gizmo,75\n";

#[test]
fn identical_rows_reconcile_and_exit_zero() {
    // Money spellings on the right read as the same numbers.
    let right = "id,name,amount\n1,Widget,\"$100.00\"\n2,Gadget,250.5\n3,Gizmo,75\n";
    let run = diff_tables(LEDGER, right, &["--key", "id"]);
    assert_eq!(run.code, 0, "stderr={}", run.stderr);
    let document = run.json();
    assert_eq!(document["contract_version"], 1);
    assert_eq!(document["summary"], summary(3, 0, 0, 0, 0));
    assert_eq!(document["key"], "id");
    assert_eq!(document["key_transform"], "none");
    assert_eq!(document["compare"], json!(["name", "amount"]));
    assert_eq!(
        statuses(&document),
        [("1", "match"), ("2", "match"), ("3", "match")]
            .map(|(key, status)| (key.to_string(), status.to_string()))
    );
    assert_eq!(
        document["results"][0],
        json!({"key": "1", "status": "match", "left_row": 2, "right_row": 2, "differences": []})
    );
}

#[test]
fn row_missing_from_the_right_is_only_left_and_exits_one() {
    let right = "id,name,amount\n1,Widget,100\n2,Gadget,250.50\n";
    let run = diff_tables(LEDGER, right, &["--key", "id"]);
    assert_eq!(run.code, 1, "stderr={}", run.stderr);
    let document = run.json();
    assert_eq!(document["summary"], summary(2, 1, 0, 0, 0));
    assert_eq!(
        document["results"][2],
        json!({"key": "3", "status": "only_left", "left_row": 4, "right_row": null, "differences": []})
    );
}

#[test]
fn row_missing_from_the_left_is_only_right_and_exits_one() {
    let right = "id,name,amount\n1,Widget,100\n2,Gadget,250.50\n3,Gizmo,75\n4,Doohickey,9\n";
    let run = diff_tables(LEDGER, right, &["--key", "id"]);
    assert_eq!(run.code, 1, "stderr={}", run.stderr);
    let document = run.json();
    assert_eq!(document["summary"], summary(3, 0, 1, 0, 0));
    assert_eq!(
        document["results"][3],
        json!({"key": "4", "status": "only_right", "left_row": null, "right_row": 5, "differences": []})
    );
}

#[test]
fn numeric_difference_within_tolerance_is_reported_but_reconciles() {
    let right = "id,name,amount\n1,Widget,100.004\n2,Gadget,250.50\n3,Gizmo,75\n";
    let run = diff_tables(LEDGER, right, &["--key", "id", "--tolerance", "0.01"]);
    assert_eq!(run.code, 0, "stderr={}", run.stderr);
    let document = run.json();
    assert_eq!(document["summary"], summary(3, 0, 0, 1, 0));
    assert_eq!(
        document["results"][0],
        json!({
            "key": "1",
            "status": "diff",
            "left_row": 2,
            "right_row": 2,
            "differences": [{
                "column": "amount",
                "left": "100",
                "right": "100.004",
                "delta": 0.004,
                "within_tolerance": true
            }]
        })
    );
}

#[test]
fn differences_outside_tolerance_exit_one_and_text_must_match_exactly() {
    let right = "id,name,amount\n1,widget,101\n2,Gadget,250.50\n3,Gizmo,75\n";
    let run = diff_tables(LEDGER, right, &["--key", "id", "--tolerance", "0.5"]);
    assert_eq!(run.code, 1, "stderr={}", run.stderr);
    let document = run.json();
    assert_eq!(document["summary"], summary(3, 0, 0, 1, 1));
    assert_eq!(
        document["results"][0]["differences"],
        json!([
            {"column": "name", "left": "Widget", "right": "widget", "within_tolerance": false},
            {"column": "amount", "left": "100", "right": "101", "delta": 1, "within_tolerance": false}
        ])
    );
}

#[test]
fn strict_exit_fails_on_differences_within_tolerance() {
    let right = "id,name,amount\n1,Widget,100.004\n2,Gadget,250.50\n3,Gizmo,75\n";
    let lenient = diff_tables(LEDGER, right, &["--key", "id", "--tolerance", "0.01"]);
    assert_eq!(lenient.code, 0, "stderr={}", lenient.stderr);
    let strict = diff_tables(
        LEDGER,
        right,
        &["--key", "id", "--tolerance", "0.01", "--strict-exit"],
    );
    assert_eq!(strict.code, 1, "stderr={}", strict.stderr);
    assert_eq!(strict.json()["summary"], summary(3, 0, 0, 1, 0));

    let clean = diff_tables(LEDGER, LEDGER, &["--key", "id", "--strict-exit"]);
    assert_eq!(clean.code, 0, "stderr={}", clean.stderr);
}

#[test]
fn explicit_compare_limits_the_columns_checked() {
    let right = "id,name,amount\n1,widget,100\n2,Gadget,250.50\n3,Gizmo,75\n";
    let run = diff_tables(LEDGER, right, &["--key", "ID", "--compare", "Amount"]);
    assert_eq!(run.code, 0, "stderr={}", run.stderr);
    let document = run.json();
    assert_eq!(document["compare"], json!(["amount"]));
    assert_eq!(document["summary"], summary(3, 0, 0, 0, 0));

    let missing = diff_tables(
        LEDGER,
        "id,name\n1,Widget\n",
        &["--key", "id", "--compare", "amount"],
    );
    assert_eq!(missing.code, 2);
    assert!(
        missing.stderr.contains("no column \"amount\" on right"),
        "{}",
        missing.stderr
    );
}

#[test]
fn keys_keep_their_source_spelling() {
    let run = diff_tables("id,v\n007,1\n", "id,v\n7,1\n", &["--key", "id"]);
    assert_eq!(run.code, 1, "stderr={}", run.stderr);
    let document = run.json();
    assert_eq!(document["summary"], summary(0, 1, 1, 0, 0));
    assert_eq!(
        statuses(&document),
        [("007", "only_left"), ("7", "only_right")]
            .map(|(key, status)| (key.to_string(), status.to_string()))
    );
}

#[test]
fn numeric_looking_headers_resolve_by_source_spelling() {
    // Typed, the headers `001`, `01`, and `1` would all read as 1.
    let left = TempFile::new("left.csv", "001,01,1\n007,a,10\n8,b,20\n");
    let run = diff(
        &[
            left.path(),
            "-",
            "--key",
            "001",
            "--compare",
            "1",
            "--from",
            "tsv",
        ],
        Some("001\t01\t1\n007\ta\t11\n8\tc\t20\n"),
    );
    assert_eq!(run.code, 1, "stderr={}", run.stderr);
    let document = run.json();
    assert_eq!(document["key"], "001");
    assert_eq!(document["compare"], json!(["1"]));
    assert_eq!(document["summary"], summary(2, 0, 0, 1, 1));
    assert_eq!(
        statuses(&document),
        [("007", "diff"), ("8", "match")]
            .map(|(key, status)| (key.to_string(), status.to_string()))
    );
    assert_eq!(
        document["results"][0]["differences"],
        json!([{"column": "1", "left": "10", "right": "11", "delta": 1, "within_tolerance": false}])
    );
}

#[test]
fn key_transforms_widen_matching_step_by_step() {
    let left = "key,v\n  Acme ,1\nGLOBEX,2\n";
    let right = "key,v\nAcme,1\nglobex,2\n";
    let cases = [
        ("none", summary(0, 2, 2, 0, 0), 1),
        ("trim", summary(1, 1, 1, 0, 0), 1),
        ("lower", summary(2, 0, 0, 0, 0), 0),
    ];
    for (transform, expected, code) in cases {
        let run = diff_tables(left, right, &["--key", "key", "--key-transform", transform]);
        assert_eq!(run.code, code, "{transform}: stderr={}", run.stderr);
        let document = run.json();
        assert_eq!(document["key_transform"], transform);
        assert_eq!(document["summary"], expected, "{transform}");
    }

    let run = diff_tables(
        "invoice,v\nINV-0042,3\n#0043,4\n",
        "invoice,v\n0042,3\nNo. 0043,4\n",
        &["--key", "invoice", "--key-transform", "digits"],
    );
    assert_eq!(run.code, 0, "stderr={}", run.stderr);
    let document = run.json();
    assert_eq!(document["summary"], summary(2, 0, 0, 0, 0));
    assert_eq!(
        statuses(&document),
        [("0042", "match"), ("0043", "match")]
            .map(|(key, status)| (key.to_string(), status.to_string()))
    );
}

#[test]
fn either_side_reads_standard_input_with_from() {
    let right = TempFile::new("right.csv", LEDGER);
    let run = diff(
        &["-", right.path(), "--key", "id", "--from", "csv"],
        Some(LEDGER),
    );
    assert_eq!(run.code, 0, "stderr={}", run.stderr);
    assert_eq!(run.json()["summary"], summary(3, 0, 0, 0, 0));

    // --from names the piped side's format; the path side keeps its extension.
    let left = TempFile::new("left.csv", LEDGER);
    let tsv = LEDGER.replace(',', "\t");
    let run = diff(
        &[left.path(), "-", "--key", "id", "--from", "tsv"],
        Some(tsv.as_str()),
    );
    assert_eq!(run.code, 0, "stderr={}", run.stderr);
    assert_eq!(run.json()["summary"], summary(3, 0, 0, 0, 0));

    let both = diff(&["-", "-", "--key", "id", "--from", "csv"], None);
    assert_eq!(both.code, 2);
    assert!(both.stderr.contains("only one side"), "{}", both.stderr);

    let no_format = diff(&[left.path(), "-", "--key", "id"], None);
    assert_eq!(no_format.code, 2);
    assert!(no_format.stderr.contains("--from"), "{}", no_format.stderr);
}

#[test]
fn duplicate_keys_exit_two_naming_the_key_and_rows() {
    let left = "id,amount\n1,10\n2,20\n3,30\n2,40\n";
    let run = diff_tables(left, LEDGER, &["--key", "id"]);
    assert_eq!(run.code, 2);
    assert!(run.stdout.is_empty(), "{}", run.stdout);
    assert!(
        run.stderr
            .contains("duplicate key \"2\" on left at rows 3, 5"),
        "{}",
        run.stderr
    );

    // Duplicates are judged after the key transform.
    let right = "id,amount\nA,1\na,2\n";
    let run = diff_tables(
        "id,amount\na,1\n",
        right,
        &["--key", "id", "--key-transform", "lower"],
    );
    assert_eq!(run.code, 2);
    assert!(
        run.stderr
            .contains("duplicate key \"a\" on right at rows 2, 3"),
        "{}",
        run.stderr
    );
}

#[test]
fn empty_keys_exit_two_and_blank_rows_are_skipped() {
    // A fully blank spacer row is not a record.
    let spaced = "id,name,amount\n1,Widget,100\n\n2,Gadget,250.50\n3,Gizmo,75\n";
    let run = diff_tables(spaced, LEDGER, &["--key", "id"]);
    assert_eq!(run.code, 0, "stderr={}", run.stderr);
    assert_eq!(run.json()["results"][1]["left_row"], 4);

    let keyless = "id,name,amount\n1,Widget,100\n,Gadget,250.50\n";
    let run = diff_tables(keyless, LEDGER, &["--key", "id"]);
    assert_eq!(run.code, 2);
    assert!(
        run.stderr
            .contains("empty key in column \"id\" on left at rows 3"),
        "{}",
        run.stderr
    );

    let run = diff_tables(
        LEDGER,
        "id,v\nnone,1\n",
        &["--key", "id", "--key-transform", "digits"],
    );
    assert_eq!(run.code, 2);
    assert!(run.stderr.contains("on right at rows 2"), "{}", run.stderr);
}

#[test]
fn usage_errors_exit_two() {
    let run = diff_tables(LEDGER, LEDGER, &["--key", "sku"]);
    assert_eq!(run.code, 2);
    assert!(
        run.stderr.contains("no column \"sku\" on left"),
        "{}",
        run.stderr
    );
    assert!(run.stderr.contains("id, name, amount"), "{}", run.stderr);

    let run = diff_tables(LEDGER, LEDGER, &["--key", "id", "--tolerance", "-1"]);
    assert_eq!(run.code, 2);
    assert!(run.stderr.contains("--tolerance"), "{}", run.stderr);
}

#[test]
fn text_format_lists_each_unreconciled_key_and_quiet_prints_nothing() {
    let right = "id,name,amount\n1,widget,101\n2,Gadget,250.50\n4,Doohickey,\n";
    let run = diff_tables(LEDGER, right, &["--key", "id", "--format", "text"]);
    assert_eq!(run.code, 1, "stderr={}", run.stderr);
    assert_eq!(
        run.stdout,
        "matched 2, only_left 1, only_right 1, diff 1, diff_outside_tolerance 1\n\
         diff \"1\" (left row 2, right row 2): name \"Widget\" -> \"widget\"; amount \"100\" -> \"101\" (delta 1, outside tolerance)\n\
         only_left \"3\" (left row 4)\n\
         only_right \"4\" (right row 4)\n"
    );

    let quiet = diff_tables(LEDGER, right, &["--key", "id", "--quiet"]);
    assert_eq!(quiet.code, 1);
    assert!(quiet.stdout.is_empty(), "{}", quiet.stdout);
}

/// The P&L rows with the Jan 2024 and YoY % columns, keys unindented and
/// amounts spelled as money.
const PNL_EXPORT: &str = "Account,Jan 2024,YoY %
Revenue,,
Product Sales,\"$420,000.00\",31.1%
Services,\"$180,000.00\",21.5%
Subscriptions,\"$92,000.00\",39.7%
Total Revenue,\"$692,000.00\",29.5%
Cost of Goods Sold,\"$277,000.00\",23.2%
Gross Profit,\"$415,000.00\",34.1%
Gross Margin %,60.0%,
Operating Expenses,,
Salaries & Benefits,\"$145,000.00\",17.1%
Rent,\"$18,000.00\",12.5%
Marketing,\"$32,000.00\",43.0%
Software & Tools,\"$12,400.00\",27.1%
Professional Fees,\"$8,500.00\",30.5%
Total OpEx,\"$215,900.00\",21.4%
Operating Income,\"$199,100.00\",49.9%
Operating Margin %,28.8%,
";

#[test]
fn workbook_reconciles_against_a_money_formatted_csv_export() {
    let workbook = fixture("sample-financials.xlsx");
    let workbook = workbook.to_str().unwrap();
    let export = TempFile::new("pnl.csv", PNL_EXPORT);

    let run = diff(
        &[
            workbook,
            export.path(),
            "--key",
            "Account",
            "--key-transform",
            "trim",
        ],
        None,
    );
    assert_eq!(run.code, 0, "stderr={}\nstdout={}", run.stderr, run.stdout);
    let document = run.json();
    assert_eq!(document["summary"], summary(17, 0, 0, 0, 0));
    assert_eq!(document["compare"], json!(["Jan 2024", "YoY %"]));
    assert_eq!(document["left"]["sheet"], "P&L");
    let product_sales = document["results"]
        .as_array()
        .unwrap()
        .iter()
        .find(|entry| entry["key"] == "Product Sales")
        .expect("Product Sales result");
    assert_eq!(product_sales["left_row"], 3);
    assert_eq!(product_sales["right_row"], 3);

    // Without trimming, the workbook's indented line items do not match.
    let run = diff(&[workbook, export.path(), "--key", "Account"], None);
    assert_eq!(run.code, 1, "stderr={}", run.stderr);
    assert_eq!(run.json()["summary"], summary(9, 8, 8, 0, 0));
}
