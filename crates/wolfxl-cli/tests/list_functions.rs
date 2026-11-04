//! Integration tests for `wolfxl list-functions`: the registry listing, its
//! filters, `--check` statuses, and the 0 / 1 / 2 exit-code contract.

use assert_cmd::Command;
use serde_json::{json, Value};

/// The built-in evaluator's registry, sorted by name.
const REGISTRY: [(&str, &str); 67] = [
    ("ABS", "math"),
    ("AND", "logical"),
    ("AVERAGE", "statistical"),
    ("AVERAGEIF", "statistical"),
    ("AVERAGEIFS", "statistical"),
    ("CHOOSE", "lookup"),
    ("CONCATENATE", "text"),
    ("COUNT", "statistical"),
    ("COUNTA", "statistical"),
    ("COUNTIF", "statistical"),
    ("COUNTIFS", "statistical"),
    ("DATE", "date"),
    ("DAY", "date"),
    ("DAYS", "date"),
    ("DB", "financial"),
    ("EDATE", "date"),
    ("EOMONTH", "date"),
    ("EXACT", "text"),
    ("FIND", "text"),
    ("FV", "financial"),
    ("HLOOKUP", "lookup"),
    ("HOUR", "date"),
    ("IF", "logical"),
    ("IFERROR", "logical"),
    ("INDEX", "lookup"),
    ("INT", "math"),
    ("IRR", "financial"),
    ("LEFT", "text"),
    ("LEN", "text"),
    ("LOWER", "text"),
    ("MATCH", "lookup"),
    ("MAX", "statistical"),
    ("MAXIFS", "statistical"),
    ("MID", "text"),
    ("MIN", "statistical"),
    ("MINIFS", "statistical"),
    ("MINUTE", "date"),
    ("MOD", "math"),
    ("MONTH", "date"),
    ("NOT", "logical"),
    ("NOW", "date"),
    ("NPV", "financial"),
    ("OFFSET", "lookup"),
    ("OR", "logical"),
    ("PMT", "financial"),
    ("POWER", "math"),
    ("PV", "financial"),
    ("REPT", "text"),
    ("RIGHT", "text"),
    ("ROUND", "math"),
    ("ROUNDDOWN", "math"),
    ("ROUNDUP", "math"),
    ("SECOND", "date"),
    ("SIGN", "math"),
    ("SLN", "financial"),
    ("SQRT", "math"),
    ("SUBSTITUTE", "text"),
    ("SUM", "math"),
    ("SUMIF", "math"),
    ("SUMIFS", "math"),
    ("TEXT", "text"),
    ("TODAY", "date"),
    ("TRIM", "text"),
    ("UPPER", "text"),
    ("VLOOKUP", "lookup"),
    ("XLOOKUP", "lookup"),
    ("YEAR", "date"),
];

/// The registry names in `category`, in listing order.
fn in_category(category: &str) -> Vec<&'static str> {
    REGISTRY
        .iter()
        .filter(|(_, listed)| *listed == category)
        .map(|(name, _)| *name)
        .collect()
}

struct Run {
    code: i32,
    stdout: String,
    stderr: String,
}

fn list_functions(args: &[&str]) -> Run {
    let mut command = Command::cargo_bin("wolfxl").expect("wolfxl bin built");
    command.arg("list-functions").args(args);
    let out = command.output().expect("run wolfxl list-functions");
    Run {
        code: out.status.code().expect("exit code"),
        stdout: String::from_utf8(out.stdout).expect("utf-8 stdout"),
        stderr: String::from_utf8(out.stderr).expect("utf-8 stderr"),
    }
}

fn json_with_code(args: &[&str], code: i32) -> Value {
    let mut args = args.to_vec();
    args.push("--json");
    let run = list_functions(&args);
    assert_eq!(run.code, code, "{args:?}: stderr={}", run.stderr);
    serde_json::from_str(&run.stdout).expect("list-functions JSON parses")
}

fn json(args: &[&str]) -> Value {
    json_with_code(args, 0)
}

/// `(name, category)` pairs of a listing document.
fn entries(document: &Value) -> Vec<(String, String)> {
    document["functions"]
        .as_array()
        .expect("functions array")
        .iter()
        .map(|entry| {
            (
                entry["name"].as_str().expect("name").to_owned(),
                entry["category"].as_str().expect("category").to_owned(),
            )
        })
        .collect()
}

fn names(document: &Value) -> Vec<String> {
    entries(document)
        .into_iter()
        .map(|(name, _)| name)
        .collect()
}

fn total(document: &Value) -> usize {
    document["total"].as_u64().expect("total") as usize
}

fn keys(document: &Value) -> Vec<&str> {
    let mut keys: Vec<&str> = document
        .as_object()
        .expect("JSON object")
        .keys()
        .map(String::as_str)
        .collect();
    keys.sort_unstable();
    keys
}

#[test]
fn full_listing_counts_and_sorts_the_whole_registry() {
    let document = json(&[]);
    assert_eq!(
        keys(&document),
        [
            "category",
            "contract_version",
            "count",
            "functions",
            "search",
            "total",
        ]
    );
    assert_eq!(document["contract_version"], 1);
    let listed = entries(&document);
    assert_eq!(total(&document), listed.len());
    assert_eq!(document["count"].as_u64(), Some(listed.len() as u64));
    assert!(
        listed.windows(2).all(|pair| pair[0].0 < pair[1].0),
        "names are sorted and unique"
    );
    let expected: Vec<(String, String)> = REGISTRY
        .iter()
        .map(|(name, category)| ((*name).to_owned(), (*category).to_owned()))
        .collect();
    assert_eq!(listed, expected);
    assert_eq!(document["category"], Value::Null);
    assert_eq!(document["search"], Value::Null);
}

#[test]
fn text_listing_aligns_names_and_ends_with_the_count() {
    let run = list_functions(&[]);
    assert_eq!(run.code, 0, "stderr={}", run.stderr);
    let mut expected: String = REGISTRY
        .iter()
        .map(|(name, category)| format!("{name:<11}  {category}\n"))
        .collect();
    expected.push_str("67 functions\n");
    assert_eq!(run.stdout, expected);
    assert!(run.stdout.starts_with(
        "ABS          math\n\
         AND          logical\n\
         AVERAGE      statistical\n"
    ));
    assert!(run.stderr.is_empty(), "stderr={}", run.stderr);
}

#[test]
fn category_filter_ignores_case() {
    let lower = json(&["--category", "math"]);
    let upper = json(&["--category", "MATH"]);
    let padded = json(&["--category", " Math "]);
    assert_eq!(lower["functions"], upper["functions"]);
    assert_eq!(lower["functions"], padded["functions"]);
    assert_eq!(lower["category"], "math");
    assert_eq!(upper["category"], "math");
    let listed = entries(&lower);
    assert!(listed.iter().all(|(_, category)| category == "math"));
    assert_eq!(
        names(&lower),
        [
            "ABS",
            "INT",
            "MOD",
            "POWER",
            "ROUND",
            "ROUNDDOWN",
            "ROUNDUP",
            "SIGN",
            "SQRT",
            "SUM",
            "SUMIF",
            "SUMIFS"
        ]
    );
    assert_eq!(lower["count"], 12);
    assert_eq!(total(&lower), REGISTRY.len());
    assert_eq!(lower["search"], Value::Null);

    for category in [
        "date",
        "financial",
        "logical",
        "lookup",
        "statistical",
        "text",
    ] {
        let upper = category.to_uppercase();
        assert_eq!(
            names(&json(&["--category", upper.as_str()])),
            in_category(category),
            "{category}"
        );
    }
    let logical = json(&["--category", "LOGICAL"]);
    assert_eq!(names(&logical), ["AND", "IF", "IFERROR", "NOT", "OR"]);
}

#[test]
fn unknown_category_exits_2_and_names_the_valid_ones() {
    let expected = "unknown category \"astrology\"; valid categories: date, financial, logical, lookup, math, statistical, text";
    for args in [
        &["--category", "astrology"][..],
        &["--category", "astrology", "--json"][..],
    ] {
        let run = list_functions(args);
        assert_eq!(run.code, 2, "{args:?}");
        assert!(run.stdout.is_empty(), "{args:?}: stdout={}", run.stdout);
        assert!(
            run.stderr.contains(expected),
            "{args:?}: stderr={}",
            run.stderr
        );
    }

    // Spreadsheet categories without a built-in function are not valid here.
    let run = list_functions(&["--category", "engineering"]);
    assert_eq!(run.code, 2);
    assert!(run.stdout.is_empty(), "stdout={}", run.stdout);
}

#[test]
fn search_matches_a_case_insensitive_substring() {
    let document = json(&["--search", "look"]);
    assert_eq!(names(&document), ["HLOOKUP", "VLOOKUP", "XLOOKUP"]);
    assert_eq!(document["search"], "look");
    assert_eq!(document["category"], Value::Null);
    assert_eq!(document["count"], 3);
    assert_eq!(total(&document), REGISTRY.len());

    assert_eq!(
        names(&json(&["--search", "Ou"])),
        [
            "COUNT",
            "COUNTA",
            "COUNTIF",
            "COUNTIFS",
            "HOUR",
            "ROUND",
            "ROUNDDOWN",
            "ROUNDUP"
        ]
    );
    assert_eq!(
        names(&json(&["--search", "ave"])),
        ["AVERAGE", "AVERAGEIF", "AVERAGEIFS"]
    );
    assert_eq!(
        names(&json(&["--search", "IFS"])),
        ["AVERAGEIFS", "COUNTIFS", "MAXIFS", "MINIFS", "SUMIFS"]
    );
}

#[test]
fn category_and_search_combine() {
    let document = json(&["--category", "math", "--search", "a"]);
    assert_eq!(names(&document), ["ABS"]);
    assert_eq!(document["category"], "math");
    assert_eq!(document["search"], "a");

    let document = json(&["--category", "statistical", "--search", "N"]);
    assert_eq!(
        names(&document),
        ["COUNT", "COUNTA", "COUNTIF", "COUNTIFS", "MIN", "MINIFS"]
    );

    let document = json(&["--category", "logical", "--search", "sum"]);
    assert!(names(&document).is_empty());
    assert_eq!(document["count"], 0);
}

#[test]
fn filtered_text_listing_reports_matches_out_of_total() {
    let all = REGISTRY.len();
    let run = list_functions(&["--search", "no-such-function"]);
    assert_eq!(run.code, 0, "stderr={}", run.stderr);
    assert_eq!(run.stdout, format!("0 of {all} functions\n"));

    let run = list_functions(&["--search", "sqrt"]);
    assert_eq!(run.code, 0);
    assert_eq!(run.stdout, format!("SQRT  math\n1 of {all} functions\n"));

    let run = list_functions(&["--search", "sum"]);
    assert_eq!(run.code, 0);
    assert_eq!(
        run.stdout,
        format!("SUM     math\nSUMIF   math\nSUMIFS  math\n3 of {all} functions\n")
    );

    let run = list_functions(&["--category", "logical"]);
    assert_eq!(run.code, 0);
    assert_eq!(
        run.stdout,
        format!(
            "AND      logical\nIF       logical\nIFERROR  logical\nNOT      logical\nOR       logical\n5 of {all} functions\n"
        )
    );
}

#[test]
fn check_all_supported_exits_0() {
    let document = json(&["--check", "sum,Average", "--check", " if "]);
    assert_eq!(
        keys(&document),
        [
            "checked",
            "contract_version",
            "results",
            "supported",
            "total"
        ]
    );
    assert_eq!(document["contract_version"], 1);
    assert_eq!(total(&document), REGISTRY.len());
    assert_eq!(document["checked"], 3);
    assert_eq!(document["supported"], 3);
    assert_eq!(
        document["results"],
        json!([
            {"name": "SUM", "status": "supported", "category": "math"},
            {"name": "AVERAGE", "status": "supported", "category": "statistical"},
            {"name": "IF", "status": "supported", "category": "logical"},
        ])
    );
}

#[test]
fn every_listed_function_checks_as_supported() {
    let listed = names(&json(&[]));
    let joined = listed.join(",");
    let document = json(&["--check", joined.as_str()]);
    assert_eq!(document["checked"], listed.len());
    assert_eq!(document["supported"], listed.len());
    for (result, (name, category)) in document["results"]
        .as_array()
        .expect("results array")
        .iter()
        .zip(REGISTRY)
    {
        assert_eq!(
            *result,
            json!({"name": name, "status": "supported", "category": category})
        );
    }
}

#[test]
fn check_reports_unsupported_with_exit_1() {
    let document = json_with_code(&["--check", "if,lookup,Stdev.S,NOTAFUNCTION"], 1);
    assert_eq!(document["contract_version"], 1);
    assert_eq!(document["checked"], 4);
    assert_eq!(document["supported"], 1);
    assert_eq!(
        document["results"],
        json!([
            {"name": "IF", "status": "supported", "category": "logical"},
            {"name": "LOOKUP", "status": "unsupported", "category": null},
            {"name": "STDEV.S", "status": "unsupported", "category": null},
            {"name": "NOTAFUNCTION", "status": "unsupported", "category": null},
        ])
    );

    let unsupported_only = list_functions(&["--check", "NOTAFUNCTION"]);
    assert_eq!(unsupported_only.code, 1);
    let unsupported_only = list_functions(&["--check", "sumproduct", "--json"]);
    assert_eq!(unsupported_only.code, 1);
    assert!(unsupported_only.stderr.is_empty());
}

#[test]
fn check_text_output_aligns_statuses_and_counts_supported() {
    let run = list_functions(&["--check", "sum,lookup,nope"]);
    assert_eq!(run.code, 1, "stderr={}", run.stderr);
    assert_eq!(
        run.stdout,
        "SUM     supported (math)\n\
         LOOKUP  unsupported\n\
         NOPE    unsupported\n\
         1 of 3 supported\n"
    );

    let run = list_functions(&["--check", "SUM"]);
    assert_eq!(run.code, 0);
    assert_eq!(run.stdout, "SUM  supported (math)\n1 of 1 supported\n");
}

#[test]
fn check_usage_errors_exit_2() {
    for args in [
        &["--check", "SUM", "--category", "math"][..],
        &["--check", "SUM", "--search", "S"][..],
        &["--check", "SUM,,IF"][..],
        &["--check", " "][..],
        &["--check"][..],
    ] {
        let run = list_functions(args);
        assert_eq!(run.code, 2, "{args:?}: stdout={}", run.stdout);
        assert!(run.stdout.is_empty(), "{args:?}: stdout={}", run.stdout);
        assert!(!run.stderr.is_empty(), "{args:?}");
    }

    let run = list_functions(&["--check", "SUM,,IF"]);
    assert!(
        run.stderr.contains("empty function name"),
        "stderr={}",
        run.stderr
    );
}
