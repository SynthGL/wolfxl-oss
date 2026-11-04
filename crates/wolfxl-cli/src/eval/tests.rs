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
    date1904: bool,
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
                data_sheet(),
            ],
            reads: RefCell::new(Vec::new()),
            date1904: false,
        }
    }

    fn value(&self, formula: &str) -> Value {
        evaluate(formula, self, "Sheet1").unwrap_or_else(|error| panic!("{formula}: {error}"))
    }

    /// Evaluate on the `Data` sheet.
    fn data(&self, formula: &str) -> Value {
        evaluate(formula, self, "Data").unwrap_or_else(|error| panic!("{formula}: {error}"))
    }
}

/// A small sales table in A1:D6 beside a sorted lookup table in F2:G6.
fn data_sheet() -> Sheet {
    let mut sheet = Sheet {
        name: "Data".into(),
        extent: Extent {
            rows: 6,
            columns: 7,
        },
        ..Sheet::default()
    };
    let text = |text: &str| Value::Text(text.into());
    let rows = [
        [
            text("Region"),
            text("Product"),
            text("Amount"),
            text("Units"),
        ],
        [
            text("North"),
            text("Apple"),
            Value::Number(100.0),
            Value::Number(5.0),
        ],
        [
            text("South"),
            text("Banana"),
            Value::Number(200.0),
            Value::Number(3.0),
        ],
        [
            text("North"),
            text("Cherry"),
            Value::Number(150.0),
            Value::Blank,
        ],
        [
            text("East"),
            text("Apple"),
            Value::Number(50.0),
            Value::Number(2.0),
        ],
        [
            text("South"),
            text("12"),
            Value::Number(300.0),
            Value::Number(7.0),
        ],
    ];
    for (row, values) in rows.into_iter().enumerate() {
        for (column, value) in values.into_iter().enumerate() {
            if value != Value::Blank {
                sheet
                    .cells
                    .insert((row as u32 + 1, column as u32 + 1), value);
            }
        }
    }
    for (row, (bound, grade)) in [
        (10.0, "D"),
        (20.0, "C"),
        (30.0, "B"),
        (40.0, "A"),
        (50.0, "S"),
    ]
    .into_iter()
    .enumerate()
    {
        sheet
            .cells
            .insert((row as u32 + 2, 6), Value::Number(bound));
        sheet.cells.insert((row as u32 + 2, 7), text(grade));
    }
    sheet
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

    fn date1904(&self) -> bool {
        self.date1904
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
fn concatenation_formats_numbers_with_display_rounding() {
    assert_cases(&[
        ("=\"x\"&(0.1+0.2)", Value::Text("x0.3".into())),
        ("=1/3&\"\"", Value::Text("0.333333333333333".into())),
    ]);
    let mut source = MemorySource::fixture();
    source.sheets[0].cells.insert((2, 2), Value::Number(3.0));
    assert_eq!(
        source.value("=B2*1.1&\" USD\""),
        Value::Text("3.3 USD".into())
    );
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
fn count_ignores_errors_in_reference_direct_and_computed_arguments() {
    assert_cases(&[
        ("=COUNT(1,#N/A)", Value::Number(1.0)),
        ("=COUNT(1/0)", Value::Number(0.0)),
        ("=COUNT(#N/A,#REF!)", Value::Number(0.0)),
        ("=COUNT(C1)", Value::Number(0.0)),
        ("=COUNT(C1,C2)", Value::Number(0.0)),
        ("=COUNT(IF(TRUE,C1,0))", Value::Number(0.0)),
        ("=COUNT(A1:C2)", Value::Number(3.0)),
    ]);
    let mut source = MemorySource::fixture();
    source.sheets[0].cells.insert((2, 3), Value::Number(5.0));
    assert_eq!(source.value("=COUNT(C1:C2)"), Value::Number(1.0));
    assert_eq!(source.value("=SUM(C1:C2)"), Value::Error(ExcelError::Na));
}

#[test]
fn count_ignores_unresolvable_sheet_references() {
    assert_cases(&[
        ("=COUNT(Missing!A1)", Value::Number(0.0)),
        ("=COUNT(1,Missing!A1)", Value::Number(1.0)),
        ("=SUM(Missing!A1)", Value::Error(ExcelError::Ref)),
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
    for function in ["SUM", "AVERAGE", "MIN", "MAX"] {
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
    for input in ["LOOKUP(1,A1,A2)", "=IF(TRUE,1,LOOKUP(1,A1,A2))"] {
        let error = parse(input).unwrap_err();
        assert_eq!(
            error,
            ParseError::Unsupported {
                function: "LOOKUP".into()
            }
        );
        assert!(error.to_string().contains("LOOKUP"));
    }
    assert_eq!(
        parse("=other(1)").unwrap_err(),
        ParseError::Unsupported {
            function: "other".into()
        }
    );
    assert_cases(&[
        ("=unknown_name", Value::Error(ExcelError::Name)),
        ("=LOOKUP", Value::Error(ExcelError::Name)),
    ]);
}

#[test]
fn function_registry_lists_exact_supported_surface() {
    let expected = [
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
    assert_eq!(
        FUNCTIONS
            .iter()
            .map(|info| (info.name, info.category))
            .collect::<Vec<_>>(),
        expected
    );
    for function in FUNCTIONS {
        let parses = (0..=6).any(|count| {
            let arguments = vec!["1"; count].join(",");
            parse(&format!("{}({arguments})", function.name)).is_ok()
        });
        assert!(parses, "{}", function.name);
        assert!(
            parse(&format!("{}({})", function.name, vec!["1"; 256].join(","))).is_err(),
            "{} rejects 256 arguments",
            function.name
        );
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

fn number(value: f64) -> Value {
    Value::Number(value)
}

fn text(value: &str) -> Value {
    Value::Text(value.into())
}

fn error(error: ExcelError) -> Value {
    Value::Error(error)
}

/// Evaluate each formula on the `Data` sheet.
fn assert_data_cases(cases: &[(&str, Value)]) {
    let source = MemorySource::fixture();
    for (formula, expected) in cases {
        assert_eq!(source.data(formula), *expected, "{formula}");
    }
}

fn assert_close(cases: &[(&str, f64, f64)]) {
    let source = MemorySource::fixture();
    for &(formula, expected, tolerance) in cases {
        match source.data(formula) {
            Value::Number(actual) => assert!(
                (actual - expected).abs() <= tolerance,
                "{formula}: {actual} is not within {tolerance} of {expected}"
            ),
            other => panic!("{formula}: {other:?}"),
        }
    }
}

#[test]
fn math_functions_round_like_excel() {
    assert_cases(&[
        ("=ABS(-2.5)", number(2.5)),
        ("=INT(-2.5)", number(-3.0)),
        ("=MOD(10,3)", number(1.0)),
        ("=MOD(-10,3)", number(2.0)),
        ("=MOD(10,-3)", number(-2.0)),
        ("=MOD(1,0)", error(ExcelError::Div0)),
        ("=POWER(2,10)", number(1024.0)),
        ("=POWER(0,-1)", error(ExcelError::Div0)),
        ("=SIGN(-3)", number(-1.0)),
        ("=SIGN(0)", number(0.0)),
        ("=SQRT(16)", number(4.0)),
        ("=SQRT(-1)", error(ExcelError::Num)),
        ("=ROUND(2.675,2)", number(2.68)),
        ("=ROUND(-2.5,0)", number(-3.0)),
        ("=ROUND(1234.5678,-2)", number(1200.0)),
        ("=ROUND(5,-1)", number(10.0)),
        ("=ROUND(0.1+0.2,20)", number(0.3)),
        ("=ROUNDDOWN(-3.789,1)", number(-3.7)),
        ("=ROUNDDOWN(5,-1)", number(0.0)),
        ("=ROUNDUP(3.141,2)", number(3.15)),
        ("=ROUNDUP(-3.141,2)", number(-3.15)),
        ("=ROUNDUP(5,-2)", number(100.0)),
        ("=ROUND(A2,1)", number(3.0)),
        ("=ROUND(A6,1)", error(ExcelError::Value)),
    ]);
}

#[test]
fn array_constants_feed_aggregates_and_lookups() {
    assert_cases(&[
        (
            "={1,2;3,4}",
            Value::Array(vec![
                vec![number(1.0), number(2.0)],
                vec![number(3.0), number(4.0)],
            ]),
        ),
        (
            "={-1,\"a\",TRUE,#N/A}",
            Value::Array(vec![vec![
                number(-1.0),
                text("a"),
                Value::Boolean(true),
                error(ExcelError::Na),
            ]]),
        ),
        ("=SUM({1,2;3,4})", number(10.0)),
        ("=SUM({1,\"2\",TRUE})", number(1.0)),
        ("=COUNTA({1,\"\",TRUE})", number(3.0)),
        ("=INDEX({1,2;3,4},2,1)", number(3.0)),
        ("=MATCH(25,{50,40,30,20,10},-1)", number(3.0)),
    ]);
    for input in ["={1,2;3}", "={A1}", "={1+1}", "={}", "={1,}"] {
        assert!(
            matches!(parse(input), Err(ParseError::Syntax { .. })),
            "{input}"
        );
    }
}

#[test]
fn argument_counts_are_checked_when_parsing() {
    for input in [
        "=ROUND(1)",
        "=NOW(1)",
        "=COUNTIFS(A1:A2)",
        "=SUMIFS(A1:A2,B1:B2,1,C1:C2)",
        "=COUNTIFS(A1:A2,1,B1:B2)",
        "=MID(\"a\",1)",
        "=XLOOKUP(1,A1:A2)",
    ] {
        assert!(
            matches!(parse(input), Err(ParseError::Syntax { .. })),
            "{input}"
        );
    }
    assert!(parse("=_xlfn.XLOOKUP(1,A1:A2,B1:B2)").is_ok());
}

#[test]
fn empty_arguments_are_blanks_or_defaults() {
    assert_cases(&[
        ("=SUM(1,)", number(1.0)),
        ("=SUM(,1)", number(1.0)),
        ("=IF(FALSE,1,)", number(0.0)),
        ("=ROUND(2.5,)", number(3.0)),
        ("=LEFT(\"abc\",)", text("")),
        ("=SUBSTITUTE(\"abc\",\"b\",)", text("ac")),
    ]);
    assert_data_cases(&[
        // An empty range_lookup or match_type is FALSE or 0: exact match.
        ("=VLOOKUP(35,F2:G6,2,)", error(ExcelError::Na)),
        ("=MATCH(30,F2:F6,)", number(3.0)),
        ("=MATCH(35,F2:F6,)", error(ExcelError::Na)),
        // XLOOKUP and OFFSET give empty optional arguments their defaults.
        ("=XLOOKUP(41,F2:F6,G2:G6,,1)", text("S")),
        ("=XLOOKUP(41,F2:F6,G2:G6,,-1)", text("A")),
        ("=XLOOKUP(41,F2:F6,G2:G6,)", error(ExcelError::Na)),
        ("=XLOOKUP(\"South\",A2:A6,C2:C6,,,-1)", number(300.0)),
        ("=SUM(OFFSET(C2,0,0,,))", number(100.0)),
        ("=SUM(OFFSET(C2,0,0,3,))", number(450.0)),
    ]);
    for input in ["=IF()", "=IF(,,,)", "=ROUND(,,)", "=SUM(1,,"] {
        assert!(parse(input).is_err(), "{input}");
    }
}

#[test]
fn logical_functions_skip_referenced_text() {
    assert_cases(&[
        ("=AND(TRUE,1)", Value::Boolean(true)),
        ("=AND(TRUE,0)", Value::Boolean(false)),
        ("=OR(FALSE,0)", Value::Boolean(false)),
        ("=OR(A3)", Value::Boolean(true)),
        ("=AND(A6)", error(ExcelError::Value)),
        ("=AND(A1:A6)", Value::Boolean(true)),
        ("=AND(B1:B5)", Value::Boolean(false)),
        ("=OR(C1,TRUE)", error(ExcelError::Na)),
        ("=AND(\"true\")", Value::Boolean(true)),
        ("=AND(\"x\")", error(ExcelError::Value)),
        ("=NOT(0)", Value::Boolean(true)),
        ("=NOT(A3)", Value::Boolean(false)),
        ("=IFERROR(1/0,\"bad\")", text("bad")),
        ("=IFERROR(C1,5)", number(5.0)),
        ("=IFERROR(A1,5)", number(2.0)),
        ("=IFERROR(A4,5)", number(0.0)),
        ("=CHOOSE(2,\"a\",\"b\",\"c\")", text("b")),
        ("=CHOOSE(4,\"a\",\"b\",\"c\")", error(ExcelError::Value)),
        ("=CHOOSE(0,\"a\")", error(ExcelError::Value)),
        ("=SUM(CHOOSE(1,A1:A6,B1:B5))", number(10.0)),
        ("=COUNTA(A1:A6)", number(5.0)),
        ("=COUNTA(C1:C2)", number(2.0)),
        ("=COUNTA(A1:C6)", number(12.0)),
    ]);
}

#[test]
fn conditional_aggregates_follow_excel_criteria() {
    assert_data_cases(&[
        ("=COUNTIF(A2:A6,\"North\")", number(2.0)),
        ("=COUNTIF(A2:A6,\"n*\")", number(2.0)),
        ("=COUNTIF(A2:A6,\"?outh\")", number(2.0)),
        ("=COUNTIF(C2:C6,\">=150\")", number(3.0)),
        ("=COUNTIF(C2:C6,\"<>200\")", number(4.0)),
        ("=COUNTIF(D2:D6,\"\")", number(1.0)),
        ("=COUNTIF(D2:D6,\"<>\")", number(4.0)),
        ("=COUNTIF(B2:B6,12)", number(1.0)),
        ("=COUNTIF(D:D,\"\")", number(1_048_571.0)),
        ("=SUMIF(A2:A6,\"North\",C2:C6)", number(250.0)),
        ("=SUMIF(C2:C6,\">100\")", number(650.0)),
        ("=SUMIF(A2:A6,\"South\",C2)", number(500.0)),
        ("=AVERAGEIF(A2:A6,\"North\",C2:C6)", number(125.0)),
        ("=AVERAGEIF(A2:A6,\"West\",C2:C6)", error(ExcelError::Div0)),
        (
            "=SUMIFS(C2:C6,A2:A6,\"North\",B2:B6,\"Apple\")",
            number(100.0),
        ),
        ("=COUNTIFS(A2:A6,\"South\",C2:C6,\">250\")", number(1.0)),
        ("=AVERAGEIFS(C2:C6,A2:A6,\"<>North\")", number(550.0 / 3.0)),
        ("=MINIFS(C2:C6,A2:A6,\"South\")", number(200.0)),
        ("=MAXIFS(C2:C6,A2:A6,\"North\")", number(150.0)),
        ("=MAXIFS(C2:C6,A2:A6,\"West\")", number(0.0)),
        ("=SUMIFS(C2:C6,A2:A5,\"North\")", error(ExcelError::Value)),
        ("=SUM(SUMIF(A2:A6,\"East\",C2:C6),1)", number(51.0)),
    ]);
}

#[test]
fn lookups_find_exact_approximate_and_wildcard_matches() {
    assert_data_cases(&[
        ("=MATCH(30,F2:F6)", number(3.0)),
        ("=MATCH(35,F2:F6,1)", number(3.0)),
        ("=MATCH(5,F2:F6,1)", error(ExcelError::Na)),
        ("=MATCH(\"cherry\",B2:B6,0)", number(3.0)),
        ("=MATCH(\"B*\",B2:B6,0)", number(2.0)),
        ("=MATCH(1,A2:D6,0)", error(ExcelError::Na)),
        ("=VLOOKUP(\"South\",A2:D6,3,FALSE)", number(200.0)),
        ("=VLOOKUP(\"south\",A2:D6,4,FALSE)", number(3.0)),
        ("=VLOOKUP(\"Cherry\",B2:D6,3,FALSE)", number(0.0)),
        ("=VLOOKUP(35,F2:G6,2)", text("B")),
        ("=VLOOKUP(5,F2:G6,2)", error(ExcelError::Na)),
        ("=VLOOKUP(10,F2:G6,3)", error(ExcelError::Ref)),
        ("=VLOOKUP(10,F2:G6,0)", error(ExcelError::Value)),
        ("=HLOOKUP(\"Amount\",A1:D6,3,FALSE)", number(200.0)),
        ("=INDEX(A2:D6,2,3)", number(200.0)),
        ("=INDEX(F2:F6,4)", number(40.0)),
        ("=INDEX(A1:D1,2)", text("Product")),
        ("=INDEX(A2:D6,3,4)", number(0.0)),
        ("=INDEX(A2:D6,6,1)", error(ExcelError::Ref)),
        ("=SUM(INDEX(C2:D6,0,1))", number(800.0)),
        ("=SUM(INDEX(C2:D6,2,0))", number(203.0)),
        ("=SUM(OFFSET(C2,0,0,5,1))", number(800.0)),
        ("=OFFSET(A1,2,1)", text("Banana")),
        ("=SUM(OFFSET(C2:C3,1,0))", number(350.0)),
        ("=OFFSET(A1,-1,0)", error(ExcelError::Ref)),
        ("=OFFSET(A1,0,0,0,1)", error(ExcelError::Ref)),
        ("=XLOOKUP(\"East\",A2:A6,C2:C6)", number(50.0)),
        ("=_xlfn.XLOOKUP(\"East\",A2:A6,C2:C6)", number(50.0)),
        ("=XLOOKUP(\"West\",A2:A6,C2:C6,\"none\")", text("none")),
        ("=XLOOKUP(\"West\",A2:A6,C2:C6)", error(ExcelError::Na)),
        (
            "=XLOOKUP(\"North\",A2:A6,C2:C6,\"none\",0,-1)",
            number(150.0),
        ),
        ("=XLOOKUP(35,F2:F6,G2:G6,\"none\",-1)", text("B")),
        ("=XLOOKUP(35,F2:F6,G2:G6,\"none\",1)", text("A")),
        ("=XLOOKUP(\"Ch*\",B2:B6,C2:C6,\"none\",2)", number(150.0)),
        ("=XLOOKUP(\"Ch*\",B2:B6,C2:C6)", error(ExcelError::Na)),
        ("=SUM(XLOOKUP(\"South\",A2:A6,C2:D6))", number(203.0)),
        ("=XLOOKUP(\"East\",A2:A6,C2:C5)", error(ExcelError::Value)),
    ]);
}

#[test]
fn financial_functions_match_published_examples() {
    assert_close(&[
        ("=PMT(0.05/12,360,200000)", -1073.643246, 1e-6),
        ("=PMT(0,10,1000)", -100.0, 0.0),
        ("=PV(0.08/12,240,-500)", 59_777.15, 0.01),
        ("=FV(0.005,10,-200,-500,1)", 2_581.40, 0.01),
        ("=NPV(0.1,-10000,3000,4200,6800)", 1_188.44, 0.01),
        (
            "=IRR({-70000,12000,15000,18000,21000,26000})",
            0.086_63,
            1e-5,
        ),
        ("=SLN(30000,7500,10)", 2250.0, 0.0),
        ("=DB(1000000,100000,6,1,7)", 186_083.33, 0.01),
        ("=DB(1000000,100000,6,2,7)", 259_639.42, 0.01),
        ("=DB(1000000,100000,6,7,7)", 15_845.10, 0.01),
    ]);
    assert_data_cases(&[
        ("=NPV(-1,1)", error(ExcelError::Div0)),
        ("=IRR({1,2})", error(ExcelError::Num)),
        ("=SLN(1,1,0)", error(ExcelError::Div0)),
        ("=DB(1000,100,5,6)", error(ExcelError::Num)),
        ("=PMT(0.1,0,100)", error(ExcelError::Num)),
        (
            "=NPV(0.1,C2:C3)",
            number(100.0 / 1.1 + 200.0 / 1.1_f64.powi(2)),
        ),
    ]);
}

#[test]
fn text_functions_count_characters() {
    assert_cases(&[
        ("=CONCATENATE(\"a\",1,TRUE)", text("a1TRUE")),
        ("=CONCATENATE(A1,\"-\",A2)", text("2-3")),
        ("=CONCATENATE(A1:A2)", error(ExcelError::Value)),
        ("=EXACT(\"a\",\"A\")", Value::Boolean(false)),
        ("=FIND(\"b\",\"abcb\")", number(2.0)),
        ("=FIND(\"b\",\"abcb\",3)", number(4.0)),
        ("=FIND(\"B\",\"abc\")", error(ExcelError::Value)),
        ("=LEFT(\"hello\",2)", text("he")),
        ("=LEFT(\"hello\")", text("h")),
        ("=RIGHT(\"hello\",3)", text("llo")),
        ("=RIGHT(\"hi\",-1)", error(ExcelError::Value)),
        ("=MID(\"hello\",2,3)", text("ell")),
        ("=MID(\"hello\",0,1)", error(ExcelError::Value)),
        ("=LEN(\"h\u{e9}llo\")", number(5.0)),
        ("=LOWER(\"AbC\")", text("abc")),
        ("=UPPER(\"abc\")", text("ABC")),
        ("=REPT(\"ab\",3)", text("ababab")),
        ("=REPT(\"a\",-1)", error(ExcelError::Value)),
        ("=REPT(\"ab\",20000)", error(ExcelError::Value)),
        ("=SUBSTITUTE(\"a-b-c\",\"-\",\"+\")", text("a+b+c")),
        ("=SUBSTITUTE(\"a-b-c\",\"-\",\"+\",2)", text("a-b+c")),
        ("=SUBSTITUTE(\"aaa\",\"aa\",\"b\")", text("ba")),
        ("=TRIM(\"  a   b  \")", text("a b")),
        ("=TEXT(1234.5,\"#,##0.00\")", text("1,234.50")),
        ("=TEXT(\"12\",\"000\")", text("012")),
        ("=TEXT(\"abc\",\"0\")", text("abc")),
        ("=TEXT(45356,\"yyyy-mm-dd\")", text("2024-03-05")),
        ("=TEXT(A4,\"0.0\")", text("0.0")),
    ]);
}

#[test]
fn date_functions_use_excel_serials() {
    assert_cases(&[
        ("=DATE(2024,1,1)", number(45292.0)),
        ("=DATE(2024,14,1)", number(45689.0)),
        ("=DATE(2024,1,0)", number(45291.0)),
        ("=DATE(1900,2,29)", number(60.0)),
        ("=DATE(24,1,1)=DATE(1924,1,1)", Value::Boolean(true)),
        ("=DATE(-1,1,1)", error(ExcelError::Num)),
        ("=DATE(10000,1,1)", error(ExcelError::Num)),
        ("=YEAR(45292)", number(2024.0)),
        ("=MONTH(45350)", number(2.0)),
        ("=DAY(45350)", number(28.0)),
        ("=DAY(60)", number(29.0)),
        ("=YEAR(-1)", error(ExcelError::Num)),
        ("=HOUR(0.75)", number(18.0)),
        ("=MINUTE(45292.5+7/1440)", number(7.0)),
        ("=SECOND(9/86400)", number(9.0)),
        ("=HOUR(-0.5)", error(ExcelError::Num)),
        ("=DAYS(45300,45292)", number(8.0)),
        ("=EDATE(DATE(2024,1,31),1)", number(45351.0)),
        ("=EDATE(DATE(2024,3,31),-1)", number(45351.0)),
        ("=EOMONTH(DATE(2024,1,15),1)", number(45351.0)),
        ("=EOMONTH(DATE(2024,1,15),-1)", number(45291.0)),
    ]);
    let source = MemorySource::fixture();
    let Value::Number(today) = source.value("=TODAY()") else {
        panic!("TODAY is a number");
    };
    assert!(today > 45_000.0 && today.fract() == 0.0);
    assert_eq!(source.value("=NOW()>=TODAY()"), Value::Boolean(true));
}

#[test]
fn the_1904_date_system_shifts_serials() {
    let mut source = MemorySource::fixture();
    source.date1904 = true;
    for (formula, expected) in [
        ("=DATE(1904,1,1)", number(0.0)),
        ("=DATE(2024,1,1)", number(45292.0 - 1462.0)),
        ("=YEAR(0)", number(1904.0)),
        ("=DATE(1903,12,31)", error(ExcelError::Num)),
        ("=TEXT(0,\"yyyy-mm-dd\")", text("1904-01-01")),
    ] {
        assert_eq!(source.value(formula), expected, "{formula}");
    }
}
