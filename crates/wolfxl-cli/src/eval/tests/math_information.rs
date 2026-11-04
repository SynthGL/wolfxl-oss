//! Expected values are cached results captured from Microsoft Excel.
//! Provenance and the exact formulas are committed with the oracle fixture.

use super::MemorySource;
use crate::eval::{ExcelError, Value};

fn native_cases(function: &str) {
    let oracle: serde_json::Value = serde_json::from_str(include_str!(
        "../../../tests/fixtures/math-information.oracle.json"
    ))
    .expect("native Excel capture JSON");
    let source = MemorySource::fixture();
    let mut arrays = std::collections::BTreeMap::new();
    for case in oracle["cases"].as_array().unwrap() {
        if case["function"] != function {
            continue;
        }
        let formula = case["formula"].as_str().unwrap();
        let actual = source.value(formula);
        let expected = &case["expected"];
        if let Some(array_formula) = case["array_formula"].as_str() {
            let rows = case["rows"].as_u64().unwrap() as usize;
            let columns = case["columns"].as_u64().unwrap() as usize;
            let row = case["row"].as_u64().unwrap() as usize - 1;
            let column = case["column"].as_u64().unwrap() as usize - 1;
            let values = arrays
                .entry(array_formula)
                .or_insert_with(|| vec![vec![Value::Blank; columns]; rows]);
            values[row][column] = Value::Boolean(expected["value"].as_bool().unwrap());
        }
        match expected["type"].as_str().unwrap() {
            "number" => {
                let expected = expected["value"].as_f64().unwrap();
                let Value::Number(actual) = actual else {
                    panic!("{formula}: expected {expected}, got {actual:?}");
                };
                let tolerance = 1e-12 * expected.abs().max(1.0);
                assert!(
                    (actual - expected).abs() <= tolerance,
                    "{formula}: expected {expected}, got {actual}"
                );
            }
            "boolean" => assert_eq!(
                actual,
                Value::Boolean(expected["value"].as_bool().unwrap()),
                "{formula}"
            ),
            "error" => assert_eq!(
                actual,
                Value::Error(ExcelError::parse(expected["value"].as_str().unwrap()).unwrap()),
                "{formula}"
            ),
            other => panic!("unrecognized native value type: {other}"),
        }
    }
    for (formula, expected) in arrays {
        assert_eq!(source.value(formula), Value::Array(expected), "{formula}");
    }
}

macro_rules! native_tests {
    ($($test:ident => $function:literal),* $(,)?) => {
        $(#[test]
        fn $test() {
            native_cases($function);
        })*
    };
}

native_tests! {
    acos_native_excel => "ACOS",
    asin_native_excel => "ASIN",
    atan_native_excel => "ATAN",
    atan2_native_excel => "ATAN2",
    cos_native_excel => "COS",
    sin_native_excel => "SIN",
    tan_native_excel => "TAN",
    degrees_native_excel => "DEGREES",
    radians_native_excel => "RADIANS",
    pi_native_excel => "PI",
    exp_native_excel => "EXP",
    ln_native_excel => "LN",
    log_native_excel => "LOG",
    log10_native_excel => "LOG10",
    ceiling_native_excel => "CEILING",
    floor_native_excel => "FLOOR",
    trunc_native_excel => "TRUNC",
    product_native_excel => "PRODUCT",
    rand_native_excel => "RAND",
    randbetween_native_excel => "RANDBETWEEN",
    isblank_native_excel => "ISBLANK",
    iserror_native_excel => "ISERROR",
    isna_native_excel => "ISNA",
    isnumber_native_excel => "ISNUMBER",
    istext_native_excel => "ISTEXT",
    na_native_excel => "NA",
}
