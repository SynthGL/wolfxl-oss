use std::cell::RefCell;
use std::collections::BTreeMap;

use super::{evaluate, parse, DataSource, ExcelError, Extent, ParseError, Value, FUNCTIONS};

#[derive(Default)]
struct Sheet {
    name: String,
    extent: Extent,
    cells: BTreeMap<(u32, u32), Value>,
}

struct MemorySource {
    sheets: Vec<Sheet>,
    reads: RefCell<Vec<(String, u32, u32)>>,
}

impl MemorySource {
    fn fixture() -> Self {
        let mut sheet = Sheet {
            name: "Sheet1".into(),
            extent: Extent {
                rows: 6,
                columns: 3,
            },
            ..Sheet::default()
        };
        for (row, value) in [
            Value::Number(2.0),
            Value::Text("3".into()),
            Value::Boolean(true),
            Value::Blank,
            Value::Number(8.0),
            Value::Text("not a number".into()),
        ]
        .into_iter()
        .enumerate()
        {
            sheet.cells.insert((row as u32 + 1, 1), value);
        }
        for (row, value) in [
            Value::Number(-4.0),
            Value::Number(0.0),
            Value::Number(6.0),
            Value::Boolean(false),
            Value::Text("7".into()),
        ]
        .into_iter()
        .enumerate()
        {
            sheet.cells.insert((row as u32 + 1, 2), value);
        }
        sheet.cells.insert((1, 3), Value::Error(ExcelError::Na));
        sheet.cells.insert((2, 3), Value::Error(ExcelError::Ref));
        let mut pnl = Sheet {
            name: "P&L".into(),
            extent: Extent {
                rows: 5,
                columns: 2,
            },
            ..Sheet::default()
        };
        for row in 3..=5 {
            pnl.cells.insert((row, 2), Value::Number(f64::from(row)));
        }
        let mut sheet2 = Sheet {
            name: "Sheet2".into(),
            extent: Extent {
                rows: 1,
                columns: 1,
            },
            ..Sheet::default()
        };
        sheet2.cells.insert((1, 1), Value::Number(42.0));
        let mut apostrophe = Sheet {
            name: "Bob's sheet".into(),
            extent: Extent {
                rows: 1,
                columns: 1,
            },
            ..Sheet::default()
        };
        apostrophe.cells.insert((1, 1), Value::Number(9.0));
        Self {
            sheets: vec![
                sheet,
                pnl,
                sheet2,
                apostrophe,
                Sheet {
                    name: "Empty".into(),
                    ..Sheet::default()
                },
            ],
            reads: RefCell::new(Vec::new()),
        }
    }

    fn value(&self, formula: &str) -> Value {
        evaluate(formula, self, "Sheet1").unwrap_or_else(|error| panic!("{formula}: {error}"))
    }
}

impl DataSource for MemorySource {
    fn sheet_names(&self) -> impl Iterator<Item = &str> {
        self.sheets.iter().map(|sheet| sheet.name.as_str())
    }

    fn cell_value(&self, sheet: &str, row: u32, column: u32) -> Value {
        self.reads
            .borrow_mut()
            .push((sheet.to_owned(), row, column));
        self.sheets
            .iter()
            .find(|candidate| candidate.name == sheet)
            .unwrap()
            .cells
            .get(&(row, column))
            .cloned()
            .unwrap_or(Value::Blank)
    }

    fn used_extent(&self, sheet: &str) -> Extent {
        self.sheets
            .iter()
            .find(|candidate| candidate.name == sheet)
            .unwrap()
            .extent
    }
}

fn assert_cases(cases: &[(&str, Value)]) {
    let source = MemorySource::fixture();
    for (formula, expected) in cases {
        assert_eq!(source.value(formula), *expected, "{formula}");
    }
}

#[test]
fn excel_precedence_and_associativity() {
    assert_cases(&[
        ("=-2^2", Value::Number(4.0)),
        ("=2+3*4", Value::Number(14.0)),
        ("=10%", Value::Number(0.1)),
        ("=2^3^2", Value::Number(64.0)),
        ("=8/4/2", Value::Number(1.0)),
        ("=10-3-2", Value::Number(5.0)),
        ("=2^-2", Value::Number(0.25)),
        ("=-(2^2)", Value::Number(-4.0)),
        ("=+--2", Value::Number(2.0)),
        ("=200*10%", Value::Number(20.0)),
        ("=100%%", Value::Number(0.01)),
        ("=2^2%", Value::Number(2.0_f64.powf(0.02))),
        ("=1+2&3*4", Value::Text("312".into())),
        ("=1&2=\"12\"", Value::Boolean(true)),
    ]);
}

#[test]
fn literals_and_optional_equals() {
    assert_cases(&[
        ("1e3 + .5", Value::Number(1000.5)),
        (" \t = 1E-3 + 1. ", Value::Number(1.001)),
        ("=\"a\"\"b\"", Value::Text("a\"b".into())),
        ("=\"\"", Value::Text(String::new())),
        ("=tRuE", Value::Boolean(true)),
        ("false", Value::Boolean(false)),
    ]);
}

#[test]
fn all_error_literals_and_display_spellings() {
    let source = MemorySource::fixture();
    for error in [
        ExcelError::Div0,
        ExcelError::Value,
        ExcelError::Ref,
        ExcelError::Name,
        ExcelError::Na,
        ExcelError::Num,
        ExcelError::Null,
    ] {
        assert_eq!(source.value(error.as_str()), Value::Error(error));
        assert_eq!(
            source.value(&error.as_str().to_ascii_lowercase()),
            Value::Error(error)
        );
        assert_eq!(error.to_string(), error.as_str());
    }
}

#[test]
fn arithmetic_coercions() {
    assert_cases(&[
        ("=\"3\"+1", Value::Number(4.0)),
        ("=\" 3.5 \"*2", Value::Number(7.0)),
        ("=\"1e2\"/2", Value::Number(50.0)),
        ("=\"x\"+1", Value::Error(ExcelError::Value)),
        ("=\"\"+1", Value::Error(ExcelError::Value)),
        ("=TRUE+FALSE", Value::Number(1.0)),
        ("=A4+2", Value::Number(2.0)),
        ("=A2+1", Value::Number(4.0)),
        ("=A3*4", Value::Number(4.0)),
        ("=-\"3\"", Value::Number(-3.0)),
        ("=+TRUE", Value::Number(1.0)),
        ("=\"10\"%", Value::Number(0.1)),
        ("=1/0", Value::Error(ExcelError::Div0)),
        ("=1/FALSE", Value::Error(ExcelError::Div0)),
        ("=1/A4", Value::Error(ExcelError::Div0)),
    ]);
}

#[test]
fn money_text_coerces_with_the_row_filter_parser() {
    assert_cases(&[
        ("=\"$1,200\"+0", Value::Number(1200.0)),
        ("=\"+$1,200.50\"*2", Value::Number(2401.0)),
        ("=\"-$1,200.50\"/2", Value::Number(-600.25)),
        ("=\"(75.00)\"+100", Value::Number(25.0)),
        ("=\" ($1,200.50) \"+0", Value::Number(-1200.5)),
        ("=\"1,234.5e2\"+0", Value::Number(123450.0)),
        ("=\"$1,20\"+0", Value::Error(ExcelError::Value)),
        ("=\"$1,200 trailing\"+0", Value::Error(ExcelError::Value)),
        ("=\"25%\"+0", Value::Error(ExcelError::Value)),
    ]);
}

#[test]
fn concatenation_coerces_scalars() {
    assert_cases(&[
        ("=A4&\"x\"", Value::Text("x".into())),
        ("=TRUE&FALSE", Value::Text("TRUEFALSE".into())),
        ("=3&\"!\"", Value::Text("3!".into())),
        ("=\"x\"&-0", Value::Text("x0".into())),
        ("=\"a\"&\"b\"&\"c\"", Value::Text("abc".into())),
    ]);
}

#[test]
fn comparisons_use_type_order_and_case_insensitive_text() {
    assert_cases(&[
        ("=2<\"1\"", Value::Boolean(true)),
        ("=\"z\"<FALSE", Value::Boolean(true)),
        ("=TRUE>999", Value::Boolean(true)),
        ("=\"3\"=3", Value::Boolean(false)),
        ("=\"HeLLo\"=\"hello\"", Value::Boolean(true)),
        ("=\"a\"<\"B\"", Value::Boolean(true)),
        ("=FALSE<TRUE", Value::Boolean(true)),
        ("=2<=2", Value::Boolean(true)),
        ("=3>=2", Value::Boolean(true)),
        ("=2<>3", Value::Boolean(true)),
        ("=2>3", Value::Boolean(false)),
        ("=A4=0", Value::Boolean(true)),
        ("=A4=\"\"", Value::Boolean(true)),
        ("=A4=FALSE", Value::Boolean(true)),
        ("=A4<1", Value::Boolean(true)),
        ("=A4<\"a\"", Value::Boolean(true)),
        ("=0=A4", Value::Boolean(true)),
        ("=\"\"=A4", Value::Boolean(true)),
        ("=FALSE=A4", Value::Boolean(true)),
        ("=1=1=TRUE", Value::Boolean(true)),
    ]);
}

#[test]
fn sum_skips_reference_text_booleans_and_blanks_but_coerces_direct_arguments() {
    assert_cases(&[
        ("=SUM(A1:A6)", Value::Number(10.0)),
        ("=SUM(A2,A3,A4)", Value::Number(0.0)),
        ("=SUM(TRUE,FALSE,\"3\",2)", Value::Number(6.0)),
        ("=SUM(A1:A6,TRUE,\"3\")", Value::Number(14.0)),
        ("=SUM(\"x\")", Value::Error(ExcelError::Value)),
        ("=SUM()", Value::Number(0.0)),
        ("=sum(a1:a6)", Value::Number(10.0)),
    ]);
}

#[test]
fn average_reference_and_scalar_rules() {
    assert_cases(&[
        ("=AVERAGE(A1:A6)", Value::Number(5.0)),
        ("=AVERAGE(A2,A3,A4)", Value::Error(ExcelError::Div0)),
        ("=AVERAGE(TRUE,FALSE,\"3\",2)", Value::Number(1.5)),
        ("=AVERAGE(A1:A6,TRUE,\"3\")", Value::Number(3.5)),
        ("=AVERAGE(\"x\")", Value::Error(ExcelError::Value)),
        ("=AVERAGE()", Value::Error(ExcelError::Div0)),
    ]);
}

#[test]
fn count_reference_and_scalar_rules() {
    assert_cases(&[
        ("=COUNT(A1:A6)", Value::Number(2.0)),
        ("=COUNT(A2,A3,A4)", Value::Number(0.0)),
        ("=COUNT(TRUE,FALSE,\"3\",2)", Value::Number(4.0)),
        ("=COUNT(A1:A6,TRUE,\"3\")", Value::Number(4.0)),
        ("=COUNT(\"x\",\"\",\"NaN\",\"inf\")", Value::Number(0.0)),
        ("=COUNT()", Value::Number(0.0)),
    ]);
}

#[test]
fn min_and_max_reference_and_scalar_rules() {
    assert_cases(&[
        ("=MIN(A1:A6)", Value::Number(2.0)),
        ("=MAX(A1:A6)", Value::Number(8.0)),
        ("=MIN(A2,A3,A4)", Value::Number(0.0)),
        ("=MAX(A2,A3,A4)", Value::Number(0.0)),
        ("=MIN(TRUE,FALSE,\"3\",2)", Value::Number(0.0)),
        ("=MAX(TRUE,FALSE,\"3\",2)", Value::Number(3.0)),
        ("=MIN(A1:A6,TRUE,\"3\")", Value::Number(1.0)),
        ("=MAX(A1:A6,TRUE,\"30\")", Value::Number(30.0)),
        ("=MIN(\"x\")", Value::Error(ExcelError::Value)),
        ("=MAX(\"x\")", Value::Error(ExcelError::Value)),
        ("=MIN()", Value::Number(0.0)),
        ("=MAX()", Value::Number(0.0)),
        ("=MIN(-5,-2)", Value::Number(-5.0)),
        ("=MAX(-5,-2)", Value::Number(-2.0)),
    ]);
}

#[test]
fn all_aggregates_propagate_errors_in_reference_and_direct_arguments() {
    let source = MemorySource::fixture();
    for function in ["SUM", "AVERAGE", "COUNT", "MIN", "MAX"] {
        for arguments in [
            "A1:C2",
            "C1",
            "1,#N/A",
            "#N/A,#REF!",
            "C1,C2",
            "IF(TRUE,C1,0)",
        ] {
            assert_eq!(
                source.value(&format!("={function}({arguments})")),
                Value::Error(ExcelError::Na),
                "{function}({arguments})"
            );
        }
        assert_eq!(
            source.value(&format!("={function}(#REF!,#N/A)")),
            Value::Error(ExcelError::Ref)
        );
    }
}

#[test]
fn errors_propagate_left_to_right_in_operators() {
    let source = MemorySource::fixture();
    for operator in [
        "+", "-", "*", "/", "^", "&", "=", "<>", "<", ">", "<=", ">=",
    ] {
        assert_eq!(
            source.value(&format!("=#REF!{operator}#N/A")),
            Value::Error(ExcelError::Ref)
        );
        assert_eq!(
            source.value(&format!("=1{operator}#N/A")),
            Value::Error(ExcelError::Na)
        );
    }
    assert_cases(&[
        ("=-#N/A", Value::Error(ExcelError::Na)),
        ("=#REF!%", Value::Error(ExcelError::Ref)),
        ("=(1/0)+#N/A", Value::Error(ExcelError::Div0)),
    ]);
}

#[test]
fn if_short_circuits_and_coerces_conditions() {
    assert_cases(&[
        ("=IF(TRUE,7,1/0)", Value::Number(7.0)),
        ("=IF(FALSE,1/0,8)", Value::Number(8.0)),
        ("=IF(FALSE,1/0)", Value::Boolean(false)),
        ("=IF(TRUE,2)", Value::Number(2.0)),
        ("=IF(-2,4,5)", Value::Number(4.0)),
        ("=IF(0,4,5)", Value::Number(5.0)),
        ("=IF(A4,4,5)", Value::Number(5.0)),
        ("=IF(\"TRUE\",4,5)", Value::Error(ExcelError::Value)),
        ("=IF(\"1\",4,5)", Value::Error(ExcelError::Value)),
        ("=IF(\"\",4,5)", Value::Error(ExcelError::Value)),
        ("=IF(#N/A,4,5)", Value::Error(ExcelError::Na)),
        ("=IF(TRUE,A4,5)", Value::Number(0.0)),
        ("=if(true,SUM(2,3),#REF!)", Value::Number(5.0)),
        ("=IF(TRUE,4,Missing!A1)", Value::Number(4.0)),
    ]);
    let source = MemorySource::fixture();
    assert_eq!(source.value("=IF(TRUE,1,C1)"), Value::Number(1.0));
    assert!(source.reads.borrow().is_empty());
}

#[test]
fn references_are_one_based_and_absolute_markers_do_not_change_reads() {
    assert_cases(&[
        ("=A1", Value::Number(2.0)),
        ("=$A$1", Value::Number(2.0)),
        ("=$A1", Value::Number(2.0)),
        ("=A$1", Value::Number(2.0)),
        ("=A4", Value::Blank),
        ("=XFD1048576", Value::Blank),
        ("=A0", Value::Error(ExcelError::Ref)),
        ("=XFE1", Value::Error(ExcelError::Ref)),
        ("=A1048577", Value::Error(ExcelError::Ref)),
    ]);
}

#[test]
fn range_results_retain_two_dimensional_shape() {
    let expected = Value::Array(vec![
        vec![Value::Number(2.0), Value::Number(-4.0)],
        vec![Value::Text("3".into()), Value::Number(0.0)],
    ]);
    let source = MemorySource::fixture();
    assert_eq!(source.value("=A1:B2"), expected);
    assert_eq!(source.value("=$B$2:$A$1"), expected);
    assert_eq!(
        source.value("=A1:A1"),
        Value::Array(vec![vec![Value::Number(2.0)]])
    );
    assert_eq!(source.value("=SUM(A1:B2)"), Value::Number(-2.0));
    assert_eq!(source.value("=A1:A2+1"), Value::Error(ExcelError::Value));
    assert_eq!(source.value("=IF(TRUE,A1:B2,1/0)"), expected);
}

#[test]
fn whole_column_references_are_bounded_by_used_extent() {
    let source = MemorySource::fixture();
    assert_eq!(source.value("=SUM(B:B)"), Value::Number(2.0));
    assert_eq!(
        *source.reads.borrow(),
        (1..=6)
            .map(|row| ("Sheet1".to_owned(), row, 2))
            .collect::<Vec<_>>()
    );
    source.reads.borrow_mut().clear();
    assert_eq!(
        source.value("=$B:$B"),
        Value::Array(vec![
            vec![Value::Number(-4.0)],
            vec![Value::Number(0.0)],
            vec![Value::Number(6.0)],
            vec![Value::Boolean(false)],
            vec![Value::Text("7".into())],
            vec![Value::Blank],
        ])
    );
    assert_eq!(source.reads.borrow().len(), 6);
    source.reads.borrow_mut().clear();
    assert_eq!(source.value("=SUM(D:D)"), Value::Number(0.0));
    assert!(source.reads.borrow().is_empty());
}

#[test]
fn whole_rows_and_empty_sheets_are_bounded() {
    let source = MemorySource::fixture();
    assert_eq!(
        source.value("=3:3"),
        Value::Array(vec![vec![
            Value::Boolean(true),
            Value::Number(6.0),
            Value::Blank
        ]])
    );
    assert_eq!(
        *source.reads.borrow(),
        (1..=3)
            .map(|column| ("Sheet1".to_owned(), 3, column))
            .collect::<Vec<_>>()
    );
    assert_eq!(source.value("=SUM($3:$3)"), Value::Number(6.0));
    assert_eq!(source.value("=SUM(6:99)"), Value::Number(0.0));
    source.reads.borrow_mut().clear();
    assert_eq!(source.value("=Empty!A:A"), Value::Array(vec![]));
    assert_eq!(source.value("=Empty!1:1"), Value::Array(vec![]));
    assert_eq!(source.value("=SUM(Empty!A:A)"), Value::Number(0.0));
    assert_eq!(
        source.value("=AVERAGE(Empty!1:1)"),
        Value::Error(ExcelError::Div0)
    );
    assert!(source.reads.borrow().is_empty());
}

#[test]
fn sheet_qualified_references_and_escaped_sheet_names() {
    assert_cases(&[
        ("=Sheet2!A1", Value::Number(42.0)),
        ("=sHeEt2!$a$1", Value::Number(42.0)),
        ("=SUM('P&L'!B3:B5)", Value::Number(12.0)),
        ("=SUM('P&L'!B:B)", Value::Number(12.0)),
        ("='Bob''s sheet'!A1", Value::Number(9.0)),
        ("=Missing!A1", Value::Error(ExcelError::Ref)),
        ("=SUM('not here'!A:A)", Value::Error(ExcelError::Ref)),
        ("=Missing!3:3", Value::Error(ExcelError::Ref)),
    ]);
    let source = MemorySource::fixture();
    assert_eq!(
        evaluate("A1", &source, "Sheet2").unwrap(),
        Value::Number(42.0)
    );
    assert_eq!(
        evaluate("A1", &source, "missing").unwrap(),
        Value::Error(ExcelError::Ref)
    );
    assert_eq!(
        evaluate("1+2", &source, "missing").unwrap(),
        Value::Number(3.0)
    );
}

#[test]
fn unsupported_calls_are_parse_errors_even_in_unchosen_branches() {
    for input in ["XLOOKUP(1,A1,A2)", "=IF(TRUE,1,XLOOKUP(1,A1,A2))"] {
        let error = parse(input).unwrap_err();
        assert_eq!(
            error,
            ParseError::Unsupported {
                function: "XLOOKUP".into()
            }
        );
        assert!(error.to_string().contains("XLOOKUP"));
    }
    assert_eq!(
        parse("=other(1)").unwrap_err(),
        ParseError::Unsupported {
            function: "other".into()
        }
    );
    assert_cases(&[
        ("=unknown_name", Value::Error(ExcelError::Name)),
        ("=XLOOKUP", Value::Error(ExcelError::Name)),
    ]);
}

#[test]
fn function_registry_lists_exact_supported_surface() {
    assert_eq!(
        FUNCTIONS
            .iter()
            .map(|info| (info.name, info.category))
            .collect::<Vec<_>>(),
        vec![
            ("SUM", "math"),
            ("AVERAGE", "statistical"),
            ("COUNT", "statistical"),
            ("MIN", "math"),
            ("MAX", "math"),
            ("IF", "logical"),
        ]
    );
    for function in FUNCTIONS {
        let args = if function.name == "IF" {
            "TRUE,1,0"
        } else {
            "1"
        };
        assert!(parse(&format!("{}({args})", function.name)).is_ok());
    }
}

#[test]
fn parser_rejects_malformed_input_without_panics() {
    for input in [
        "",
        "=",
        "1+",
        "(1",
        "1)",
        "\"unterminated",
        "'sheet",
        "1e+",
        "SUM(1,)",
        "SUM(,1)",
        "SUM(1 2)",
        "IF()",
        "IF(TRUE)",
        "IF(TRUE,1,2,3)",
        "A1:B",
        "A1:",
        "'sheet'",
        "#BOGUS!",
        "@A1",
        "1 2",
        "1..2",
        "Sheet2!",
        "A:A1",
    ] {
        assert!(
            matches!(parse(input), Err(ParseError::Syntax { .. })),
            "{input}"
        );
    }
}

#[test]
fn numeric_domain_and_overflow_errors() {
    assert_cases(&[
        ("=1e999", Value::Error(ExcelError::Num)),
        ("=1e308*10", Value::Error(ExcelError::Num)),
        ("=SUM(1e308,1e308)", Value::Error(ExcelError::Num)),
        ("=(-1)^.5", Value::Error(ExcelError::Num)),
        ("=0^0", Value::Error(ExcelError::Num)),
        ("=0^-1", Value::Error(ExcelError::Div0)),
    ]);
}

#[test]
fn parsed_formulas_can_be_reused_with_different_current_sheets() {
    let source = MemorySource::fixture();
    let formula = parse("=A1+1").unwrap();
    assert_eq!(formula.evaluate(&source, "Sheet1"), Value::Number(3.0));
    assert_eq!(formula.evaluate(&source, "Sheet2"), Value::Number(43.0));
}
