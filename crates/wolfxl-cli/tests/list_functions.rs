//! Integration tests for `wolfxl list-functions`: the registry listing, its
//! filters, `--check` statuses, and the 0 / 1 / 2 exit-code contract.

use assert_cmd::Command;
use serde_json::{json, Value};

/// The built-in evaluator's registry, sorted by name.
const REGISTRY: [(&str, &str); 107] = [
    ("ABS", "math"),
    ("AND", "logical"),
    ("AVERAGE", "statistical"),
    ("AVERAGEIF", "statistical"),
    ("AVERAGEIFS", "statistical"),
    ("CHOOSE", "lookup"),
    ("CONCAT", "text"),
    ("CONCATENATE", "text"),
    ("COUNT", "statistical"),
    ("COUNTA", "statistical"),
    ("COUNTBLANK", "statistical"),
    ("COUNTIF", "statistical"),
    ("COUNTIFS", "statistical"),
    ("DATE", "date"),
    ("DAY", "date"),
    ("DAYS", "date"),
    ("DB", "financial"),
    ("EDATE", "date"),
    ("EOMONTH", "date"),
    ("EXACT", "text"),
    ("FILTER", "lookup"),
    ("FIND", "text"),
    ("FV", "financial"),
    ("HLOOKUP", "lookup"),
    ("HOUR", "date"),
    ("IF", "logical"),
    ("IFERROR", "logical"),
    ("IFNA", "logical"),
    ("IFS", "logical"),
    ("INDEX", "lookup"),
    ("INT", "math"),
    ("IRR", "financial"),
    ("LARGE", "statistical"),
    ("LEFT", "text"),
    ("LEN", "text"),
    ("LOWER", "text"),
    ("MATCH", "lookup"),
    ("MAX", "statistical"),
    ("MAXIFS", "statistical"),
    ("MEDIAN", "statistical"),
    ("MID", "text"),
    ("MIN", "statistical"),
    ("MINIFS", "statistical"),
    ("MINUTE", "date"),
    ("MOD", "math"),
    ("MODE", "statistical"),
    ("MODE.SNGL", "statistical"),
    ("MONTH", "date"),
    ("NETWORKDAYS", "date"),
    ("NOT", "logical"),
    ("NOW", "date"),
    ("NPER", "financial"),
    ("NPV", "financial"),
    ("OFFSET", "lookup"),
    ("OR", "logical"),
    ("PERCENTILE", "statistical"),
    ("PERCENTILE.EXC", "statistical"),
    ("PERCENTILE.INC", "statistical"),
    ("PMT", "financial"),
    ("POWER", "math"),
    ("PV", "financial"),
    ("QUARTILE", "statistical"),
    ("QUARTILE.EXC", "statistical"),
    ("QUARTILE.INC", "statistical"),
    ("RANK", "statistical"),
    ("RANK.AVG", "statistical"),
    ("RANK.EQ", "statistical"),
    ("RATE", "financial"),
    ("REPT", "text"),
    ("RIGHT", "text"),
    ("ROUND", "math"),
    ("ROUNDDOWN", "math"),
    ("ROUNDUP", "math"),
    ("SECOND", "date"),
    ("SIGN", "math"),
    ("SLN", "financial"),
    ("SMALL", "statistical"),
    ("SORT", "lookup"),
    ("SQRT", "math"),
    ("STDEV", "statistical"),
    ("STDEV.P", "statistical"),
    ("STDEV.S", "statistical"),
    ("STDEVP", "statistical"),
    ("SUBSTITUTE", "text"),
    ("SUBTOTAL", "math"),
    ("SUM", "math"),
    ("SUMIF", "math"),
    ("SUMIFS", "math"),
    ("SUMPRODUCT", "math"),
    ("SWITCH", "logical"),
    ("TEXT", "text"),
    ("TEXTJOIN", "text"),
    ("TODAY", "date"),
    ("TRIM", "text"),
    ("UNIQUE", "lookup"),
    ("UPPER", "text"),
    ("VALUE", "text"),
    ("VAR", "statistical"),
    ("VAR.P", "statistical"),
    ("VAR.S", "statistical"),
    ("VARP", "statistical"),
    ("VLOOKUP", "lookup"),
    ("WORKDAY", "date"),
    ("XIRR", "financial"),
    ("XLOOKUP", "lookup"),
    ("XNPV", "financial"),
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
        .map(|(name, category)| format!("{name:<14}  {category}\n"))
        .collect();
    expected.push_str("107 functions\n");
    assert_eq!(run.stdout, expected);
    assert!(run.stdout.starts_with(
        "ABS             math\n\
         AND             logical\n\
         AVERAGE         statistical\n"
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
            "SUBTOTAL",
            "SUM",
            "SUMIF",
            "SUMIFS",
            "SUMPRODUCT"
        ]
    );
    assert_eq!(lower["count"], 14);
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
    assert_eq!(
        names(&logical),
        ["AND", "IF", "IFERROR", "IFNA", "IFS", "NOT", "OR", "SWITCH"]
    );
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
            "COUNTBLANK",
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
        [
            "AVERAGEIFS",
            "COUNTIFS",
            "IFS",
            "MAXIFS",
            "MINIFS",
            "SUMIFS"
        ]
    );
}

#[test]
fn category_and_search_combine() {
    let document = json(&["--category", "math", "--search", "a"]);
    assert_eq!(names(&document), ["ABS", "SUBTOTAL"]);
    assert_eq!(document["category"], "math");
    assert_eq!(document["search"], "a");

    let document = json(&["--category", "statistical", "--search", "N"]);
    assert_eq!(
        names(&document),
        [
            "COUNT",
            "COUNTA",
            "COUNTBLANK",
            "COUNTIF",
            "COUNTIFS",
            "MEDIAN",
            "MIN",
            "MINIFS",
            "MODE.SNGL",
            "PERCENTILE",
            "PERCENTILE.EXC",
            "PERCENTILE.INC",
            "QUARTILE.INC",
            "RANK",
            "RANK.AVG",
            "RANK.EQ"
        ]
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
        format!(
            "SUM         math\nSUMIF       math\nSUMIFS      math\nSUMPRODUCT  math\n4 of {all} functions\n"
        )
    );

    let run = list_functions(&["--category", "logical"]);
    assert_eq!(run.code, 0);
    assert_eq!(
        run.stdout,
        format!(
            "AND      logical\nIF       logical\nIFERROR  logical\nIFNA     logical\nIFS      logical\nNOT      logical\nOR       logical\nSWITCH   logical\n8 of {all} functions\n"
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
    let document = json_with_code(&["--check", "if,lookup,Proper,NOTAFUNCTION"], 1);
    assert_eq!(document["contract_version"], 1);
    assert_eq!(document["checked"], 4);
    assert_eq!(document["supported"], 1);
    assert_eq!(
        document["results"],
        json!([
            {"name": "IF", "status": "supported", "category": "logical"},
            {"name": "LOOKUP", "status": "unsupported", "category": null},
            {"name": "PROPER", "status": "unsupported", "category": null},
            {"name": "NOTAFUNCTION", "status": "unsupported", "category": null},
        ])
    );

    let unsupported_only = list_functions(&["--check", "NOTAFUNCTION"]);
    assert_eq!(unsupported_only.code, 1);
    let unsupported_only = list_functions(&["--check", "sequence", "--json"]);
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
