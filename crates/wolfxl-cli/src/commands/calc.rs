//! Evaluate one formula against stored values, never recalculating workbook caches.
//!
//! Exit codes: 0 for a value, 1 for a scalar Excel error, 2 for an invalid
//! request, unreadable input, parse failure, or unsupported function.

use std::io::Write;
use std::process::ExitCode;

use anyhow::Result;
use serde_json::{json, Value as JsonValue};

use crate::eval::{self, Value};
use crate::formula_data::{normalize_formula, parse_error, rounded, text_value, WorkbookData};
use crate::{input, CalcArgs};

pub(crate) fn run(args: CalcArgs) -> ExitCode {
    let formula = normalize_formula(&args.formula);
    let (sheet, value) = match evaluate(&args, &formula) {
        Ok(evaluated) => evaluated,
        Err(error) => {
            eprintln!("error: {error:#}");
            return ExitCode::from(2);
        }
    };
    let document = if args.json {
        format!("{}\n", report(&formula, &sheet, &value))
    } else {
        match &value {
            Value::Array(rows) => rows
                .iter()
                .map(|row| {
                    let fields: Vec<String> = row
                        .iter()
                        .map(|value| csv_field(&text_value(value)))
                        .collect();
                    format!("{}\n", fields.join(","))
                })
                .collect(),
            value => format!("{}\n", text_value(value)),
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
    if matches!(value, Value::Error(_)) {
        ExitCode::from(1)
    } else {
        ExitCode::SUCCESS
    }
}

/// The `--json` report for one formula: evaluates exactly as `run` does and
/// returns the same document without writing anything. A scalar Excel error
/// is a successful report whose `type` is `"error"`.
pub(crate) fn structured(args: &CalcArgs) -> Result<JsonValue> {
    let formula = normalize_formula(&args.formula);
    let (sheet, value) = evaluate(args, &formula)?;
    Ok(report(&formula, &sheet, &value))
}

fn evaluate(args: &CalcArgs, formula: &str) -> Result<(String, Value)> {
    let mut input = input::open(args.file.as_deref(), args.from)?;
    let current = input.sheet_name(args.sheet.as_deref())?;
    let data = WorkbookData::load(&mut input)?;
    let value =
        eval::evaluate(formula, &data, &current).map_err(|error| parse_error(formula, error))?;
    Ok((current, value))
}

fn report(formula: &str, sheet: &str, value: &Value) -> JsonValue {
    json!({
        "formula": formula, "sheet": sheet,
        "type": type_name(value), "value": json_value(value),
    })
}

fn type_name(value: &Value) -> &'static str {
    match value {
        Value::Blank => "blank",
        Value::Number(_) => "number",
        Value::Text(_) => "text",
        Value::Boolean(_) => "boolean",
        Value::Error(_) => "error",
        Value::Array(_) => "array",
    }
}

fn json_value(value: &Value) -> JsonValue {
    match value {
        Value::Blank => JsonValue::Null,
        Value::Number(number) if number.is_finite() => {
            let number = rounded(*number);
            if number.fract() == 0.0 && number.abs() < 9_007_199_254_740_992.0 {
                JsonValue::from(number as i64)
            } else {
                serde_json::Number::from_f64(number).map_or(JsonValue::Null, JsonValue::Number)
            }
        }
        Value::Boolean(boolean) => JsonValue::Bool(*boolean),
        Value::Array(rows) => JsonValue::Array(
            rows.iter()
                .map(|row| JsonValue::Array(row.iter().map(json_value).collect()))
                .collect(),
        ),
        other => JsonValue::String(text_value(other)),
    }
}

/// RFC 4180 quoting for array result fields.
fn csv_field(text: &str) -> String {
    if text.contains([',', '"', '\n', '\r']) {
        format!("\"{}\"", text.replace('"', "\"\""))
    } else {
        text.to_owned()
    }
}
