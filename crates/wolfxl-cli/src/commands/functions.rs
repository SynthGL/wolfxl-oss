//! `wolfxl list-functions`: list the formula evaluator's functions, or check
//! names against them.
//!
//! The registry is `crate::eval::FUNCTIONS`, the functions the built-in
//! evaluator behind `wolfxl calc` and `wolfxl peek --add` implements. A listing
//! prints every entry sorted by name with its category, optionally narrowed by
//! `--category` (case-insensitive category name) and `--search`
//! (case-insensitive substring of the function name). `--check` classifies
//! each given name as supported or unsupported.
//!
//! Exit codes: 0 for a listing or when every checked name is supported, 1
//! when a checked name is unsupported, 2 for usage failures such as an unknown
//! category or an empty name.

use std::collections::BTreeSet;
use std::io::Write;
use std::process::ExitCode;

use anyhow::{bail, Result};
use serde_json::{json, Value};

use crate::eval::{FunctionInfo, FUNCTIONS};
use crate::ListFunctionsArgs;

const CONTRACT_VERSION: u64 = 1;

/// How one `--check` name relates to the evaluator.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Status {
    Supported(&'static str),
    Unsupported,
}

impl Status {
    fn label(self) -> &'static str {
        match self {
            Status::Supported(_) => "supported",
            Status::Unsupported => "unsupported",
        }
    }

    fn category(self) -> Option<&'static str> {
        match self {
            Status::Supported(category) => Some(category),
            Status::Unsupported => None,
        }
    }
}

pub(crate) fn run(args: ListFunctionsArgs) -> ExitCode {
    let result = match &args.check {
        Some(names) => check(names, args.json),
        None => list(args.category.as_deref(), args.search.as_deref(), args.json)
            .map(|document| (document, 0)),
    };
    let (document, code) = match result {
        Ok(rendered) => rendered,
        Err(error) => {
            eprintln!("error: {error:#}");
            return ExitCode::from(2);
        }
    };
    let mut stdout = std::io::stdout().lock();
    if let Err(error) = stdout
        .write_all(document.as_bytes())
        .and_then(|()| stdout.flush())
    {
        if error.kind() != std::io::ErrorKind::BrokenPipe {
            eprintln!("error: writing result: {error}");
            return ExitCode::from(2);
        }
    }
    ExitCode::from(code)
}

/// The distinct categories present in the registry, sorted.
fn categories() -> BTreeSet<&'static str> {
    FUNCTIONS.iter().map(|function| function.category).collect()
}

/// Resolve `--category` against the categories present in the registry.
fn parse_category(text: &str) -> Result<&'static str> {
    let wanted = text.trim();
    let valid = categories();
    if let Some(category) = valid
        .iter()
        .find(|category| category.eq_ignore_ascii_case(wanted))
    {
        return Ok(*category);
    }
    bail!(
        "unknown category {text:?}; valid categories: {}",
        valid.into_iter().collect::<Vec<_>>().join(", ")
    )
}

fn list(category: Option<&str>, search: Option<&str>, json: bool) -> Result<String> {
    let wanted = category.map(parse_category).transpose()?;
    let needle = search.map(str::to_ascii_uppercase);
    let mut functions: Vec<FunctionInfo> = FUNCTIONS
        .iter()
        .filter(|function| wanted.is_none_or(|wanted| function.category == wanted))
        .filter(|function| {
            needle
                .as_deref()
                .is_none_or(|needle| function.name.contains(needle))
        })
        .copied()
        .collect();
    functions.sort_by_key(|function| function.name);
    let total = FUNCTIONS.len();

    if json {
        let entries: Vec<Value> = functions
            .iter()
            .map(|function| json!({"name": function.name, "category": function.category}))
            .collect();
        let document = json!({
            "contract_version": CONTRACT_VERSION,
            "total": total,
            "count": functions.len(),
            "category": wanted,
            "search": search,
            "functions": entries,
        });
        return Ok(format!("{document}\n"));
    }

    let width = functions
        .iter()
        .map(|function| function.name.len())
        .max()
        .unwrap_or(0);
    let summary = if wanted.is_some() || search.is_some() {
        format!("{} of {total} functions\n", functions.len())
    } else {
        format!("{total} functions\n")
    };
    Ok(functions
        .iter()
        .map(|function| format!("{:<width$}  {}\n", function.name, function.category))
        .chain(std::iter::once(summary))
        .collect())
}

fn classify(name: &str) -> Status {
    FUNCTIONS
        .iter()
        .find(|function| function.name == name)
        .map_or(Status::Unsupported, |function| {
            Status::Supported(function.category)
        })
}

/// Classify every `--check` name; the exit code is 0 only when all are supported.
fn check(names: &[String], json: bool) -> Result<(String, u8)> {
    if names.is_empty() {
        bail!("--check needs at least one function name");
    }
    let mut results: Vec<(String, Status)> = Vec::with_capacity(names.len());
    for raw in names {
        let name = raw.trim().to_ascii_uppercase();
        if name.is_empty() {
            bail!(
                "--check got an empty function name in {:?}",
                names.join(",")
            );
        }
        let status = classify(&name);
        results.push((name, status));
    }
    let supported = results
        .iter()
        .filter(|(_, status)| matches!(status, Status::Supported(_)))
        .count();
    let code = if supported == results.len() { 0 } else { 1 };

    if json {
        let entries: Vec<Value> = results
            .iter()
            .map(|(name, status)| {
                json!({"name": name, "status": status.label(), "category": status.category()})
            })
            .collect();
        let document = json!({
            "contract_version": CONTRACT_VERSION,
            "total": FUNCTIONS.len(),
            "checked": results.len(),
            "supported": supported,
            "results": entries,
        });
        return Ok((format!("{document}\n"), code));
    }

    let width = results
        .iter()
        .map(|(name, _)| name.len())
        .max()
        .unwrap_or(0);
    let summary = format!("{supported} of {} supported\n", results.len());
    let text: String = results
        .iter()
        .map(|(name, status)| match status.category() {
            Some(category) => format!("{name:<width$}  {} ({category})\n", status.label()),
            None => format!("{name:<width$}  {}\n", status.label()),
        })
        .chain(std::iter::once(summary))
        .collect();
    Ok((text, code))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::eval::{parse, ParseError};

    #[test]
    fn full_listing_counts_every_registry_entry() {
        let document: Value = serde_json::from_str(&list(None, None, true).unwrap()).unwrap();
        assert_eq!(document["contract_version"], 1);
        assert_eq!(document["total"], FUNCTIONS.len());
        assert_eq!(document["count"], FUNCTIONS.len());
        assert_eq!(
            document["functions"].as_array().unwrap().len(),
            FUNCTIONS.len()
        );
    }

    #[test]
    fn text_listing_has_one_line_per_entry_plus_count() {
        let text = list(None, None, false).unwrap();
        assert_eq!(text.lines().count(), FUNCTIONS.len() + 1);
        assert_eq!(
            text.lines().last().unwrap(),
            format!("{} functions", FUNCTIONS.len())
        );
    }

    #[test]
    fn every_category_round_trips_ignoring_case() {
        for function in FUNCTIONS {
            assert_eq!(
                parse_category(function.category).unwrap(),
                function.category
            );
            assert_eq!(
                parse_category(&function.category.to_ascii_uppercase()).unwrap(),
                function.category
            );
        }
        assert!(parse_category("astrology").is_err());
    }

    #[test]
    fn classification_agrees_with_the_parser() {
        for function in FUNCTIONS {
            assert_eq!(
                classify(function.name),
                Status::Supported(function.category)
            );
            // Some argument count from zero to six is valid for every function.
            let parses = (0..=6).any(|count| {
                let arguments = vec!["1"; count].join(",");
                parse(&format!("={}({arguments})", function.name)).is_ok()
            });
            assert!(parses, "{} parses", function.name);
        }
        for name in ["LOOKUP", "STDEV.S", "NOTAFUNCTION"] {
            assert_eq!(classify(name), Status::Unsupported);
            assert!(
                matches!(
                    parse(&format!("={name}(1)")),
                    Err(ParseError::Unsupported { .. })
                ),
                "{name} is rejected as unsupported"
            );
        }
    }
}
