//! Statistical, dynamic array, logical, text, date, and financial functions,
//! and element-wise operators. Expected values follow Microsoft's function
//! reference examples where one exists.

use super::{error, number, text, MemorySource};
use crate::eval::{evaluate, DataSource, ExcelError, Extent, Value};

/// Evaluate each formula on the `Data` sheet.
fn cases(cases: &[(&str, Value)]) {
    let source = MemorySource::fixture();
    for (formula, expected) in cases {
        assert_eq!(source.data(formula), *expected, "{formula}");
    }
}

/// Evaluate each formula on the `Data` sheet to a number near `expected`.
fn close(cases: &[(&str, f64, f64)]) {
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

fn column(values: Vec<Value>) -> Value {
    Value::Array(values.into_iter().map(|value| vec![value]).collect())
}

fn row(values: Vec<Value>) -> Value {
    Value::Array(vec![values])
}

const BREAKING: &str = "{1345,1301,1368,1322,1310,1370,1318,1350,1303,1299}";

#[test]
fn dispersion_matches_the_sample_and_population_formulas() {
    close(&[
        (&format!("=STDEV.S({BREAKING})"), 27.463_915_72, 1e-8),
        (&format!("=STDEV({BREAKING})"), 27.463_915_72, 1e-8),
        (&format!("=STDEV.P({BREAKING})"), 26.054_558_14, 1e-8),
        (&format!("=STDEVP({BREAKING})"), 26.054_558_14, 1e-8),
        (&format!("=VAR.S({BREAKING})"), 754.266_666_7, 1e-6),
        (&format!("=VAR({BREAKING})"), 754.266_666_7, 1e-6),
        (&format!("=VAR.P({BREAKING})"), 678.84, 1e-9),
        (&format!("=VARP({BREAKING})"), 678.84, 1e-9),
        ("=_xlfn.STDEV.S({1,2,3})", 1.0, 1e-12),
    ]);
    cases(&[
        // Ranges skip text and blanks: amounts 100, 200, 150, 50, 300.
        ("=VAR.S(C1:C6)", number(9250.0)),
        ("=VAR.P(C2:C6)", number(7400.0)),
        ("=STDEV.S(5)", error(ExcelError::Div0)),
        ("=STDEV.P(A1:A6)", error(ExcelError::Div0)),
        ("=VAR(1,\"x\")", error(ExcelError::Value)),
    ]);
}

#[test]
fn median_mode_and_order_statistics() {
    cases(&[
        ("=MEDIAN({1,2,3,4,5})", number(3.0)),
        ("=MEDIAN({1,2,3,4,5,6})", number(3.5)),
        ("=MEDIAN(C2:C6)", number(150.0)),
        // D4 is blank, so the units are 5, 3, 2, 7.
        ("=MEDIAN(D2:D6)", number(4.0)),
        ("=MEDIAN(A1:A6)", error(ExcelError::Num)),
        ("=MODE({5.6,4,4,3,2,4})", number(4.0)),
        ("=MODE.SNGL({1,2,2,1})", number(1.0)),
        ("=MODE({1,2,3})", error(ExcelError::Na)),
        ("=PERCENTILE.INC({1,3,2,4},0.3)", number(1.9)),
        ("=PERCENTILE({1,3,2,4},1)", number(4.0)),
        ("=PERCENTILE({1,3,2,4},1.5)", error(ExcelError::Num)),
        ("=PERCENTILE.EXC({1,2,3,6,6,6,7,8,9},0.25)", number(2.5)),
        (
            "=PERCENTILE.EXC({1,2,3,6,6,6,7,8,9},0)",
            error(ExcelError::Num),
        ),
        ("=QUARTILE({1,2,4,7,8,9,10,12},1)", number(3.5)),
        ("=QUARTILE.INC({1,2,4,7,8,9,10,12},4)", number(12.0)),
        ("=QUARTILE({1,2,4,7,8,9,10,12},5)", error(ExcelError::Num)),
        (
            "=QUARTILE.EXC({6,7,15,36,39,40,41,42,43,47,49},1)",
            number(15.0),
        ),
        (
            "=QUARTILE.EXC({6,7,15,36,39,40,41,42,43,47,49},3)",
            number(43.0),
        ),
        ("=LARGE({3,4,5,2,3,4,5,6,4,7},3)", number(5.0)),
        ("=LARGE({3,4,5,2,3,4,5,6,4,7},7)", number(4.0)),
        ("=SMALL({3,4,5,2,3,4,5,6,4,7},4)", number(4.0)),
        ("=SMALL(C2:C6,1)", number(50.0)),
        ("=LARGE(C2:C6,0)", error(ExcelError::Num)),
        ("=SMALL(C2:C6,6)", error(ExcelError::Num)),
    ]);
}

#[test]
fn rank_counts_ties_and_order() {
    cases(&[
        ("=RANK.EQ(7,{7,3.5,3.5,1,2},1)", number(5.0)),
        ("=RANK.EQ(2,{7,3.5,3.5,1,2})", number(4.0)),
        ("=RANK(3.5,{7,3.5,3.5,1,2},1)", number(3.0)),
        ("=RANK.AVG(94,{89,88,92,101,94,97,95})", number(4.0)),
        ("=RANK.AVG(3.5,{7,3.5,3.5,1,2})", number(2.5)),
        ("=RANK(5,{1,2})", error(ExcelError::Na)),
        ("=RANK(C3,C2:C6)", number(2.0)),
    ]);
}

#[test]
fn count_blank_and_sum_product() {
    cases(&[
        ("=COUNTBLANK(D2:D6)", number(1.0)),
        ("=COUNTBLANK(A1:A1)", number(0.0)),
        ("=SUMPRODUCT(C2:C6,D2:D6)", number(3300.0)),
        ("=SUMPRODUCT({1,2;3,4},{5,6;7,8})", number(70.0)),
        ("=SUMPRODUCT(C2:C6)", number(800.0)),
        ("=SUMPRODUCT({1,2},{3,4,5})", error(ExcelError::Value)),
    ]);
}

#[test]
fn element_wise_operators_broadcast_and_pad_with_na() {
    cases(&[
        ("=SUMPRODUCT((A2:A6=\"North\")*C2:C6)", number(250.0)),
        ("=SUMPRODUCT(--(C2:C6>100))", number(3.0)),
        ("=SUMPRODUCT((A2:A6=\"South\")*(C2:C6>250))", number(1.0)),
        (
            "={1,2}+{10;20}",
            Value::Array(vec![
                vec![number(11.0), number(12.0)],
                vec![number(21.0), number(22.0)],
            ]),
        ),
        (
            "={1,2,3}+{1,2}",
            row(vec![number(2.0), number(4.0), error(ExcelError::Na)]),
        ),
        (
            "={1;2}*{1,2;3,4;5,6}",
            Value::Array(vec![
                vec![number(1.0), number(2.0)],
                vec![number(6.0), number(8.0)],
                vec![error(ExcelError::Na), error(ExcelError::Na)],
            ]),
        ),
        ("=-{1,2}", row(vec![number(-1.0), number(-2.0)])),
        ("={50,100}%", row(vec![number(0.5), number(1.0)])),
        (
            "={1,\"a\"}*2",
            row(vec![number(2.0), error(ExcelError::Value)]),
        ),
        (
            "={1,2}=1",
            row(vec![Value::Boolean(true), Value::Boolean(false)]),
        ),
        ("=C2:C3&\"!\"", column(vec![text("100!"), text("200!")])),
        // Scalars behave as before.
        ("=2*3", number(6.0)),
        ("=C2>C3", Value::Boolean(false)),
        ("=1/0+{1,2}", error(ExcelError::Div0)),
    ]);
}

#[test]
fn filter_keeps_rows_or_columns_where_include_is_true() {
    cases(&[
        (
            "=FILTER(B2:B6,A2:A6=\"North\")",
            column(vec![text("Apple"), text("Cherry")]),
        ),
        (
            "=FILTER(A2:B6,C2:C6>=200)",
            Value::Array(vec![
                vec![text("South"), text("Banana")],
                vec![text("South"), text("12")],
            ]),
        ),
        (
            "=FILTER({1,2,3},{TRUE,FALSE,TRUE})",
            row(vec![number(1.0), number(3.0)]),
        ),
        ("=FILTER(D2:D6,C2:C6=150)", column(vec![Value::Blank])),
        ("=FILTER(B2:B6,C2:C6>1000)", error(ExcelError::Value)),
        ("=FILTER(B2:B6,C2:C6>1000,\"none\")", text("none")),
        ("=FILTER(B2:B6,A2:A6)", error(ExcelError::Value)),
        ("=FILTER(B2:B6,C2:C4>0)", error(ExcelError::Value)),
        ("=SUM(FILTER(C2:C6,A2:A6=\"South\"))", number(500.0)),
    ]);
}

#[test]
fn unique_and_sort() {
    cases(&[
        (
            "=UNIQUE(A2:A6)",
            column(vec![text("North"), text("South"), text("East")]),
        ),
        ("=UNIQUE(A2:A6,FALSE,TRUE)", column(vec![text("East")])),
        (
            "=UNIQUE({\"a\",\"A\",\"b\"},TRUE)",
            row(vec![text("a"), text("b")]),
        ),
        ("=UNIQUE({1,1},TRUE,TRUE)", error(ExcelError::Value)),
        (
            "=SORT(C2:C6)",
            column(vec![
                number(50.0),
                number(100.0),
                number(150.0),
                number(200.0),
                number(300.0),
            ]),
        ),
        (
            "=SORT(D2:D6,1,-1)",
            column(vec![
                number(7.0),
                number(5.0),
                number(3.0),
                number(2.0),
                Value::Blank,
            ]),
        ),
        (
            "=SORT(D2:D6)",
            column(vec![
                number(2.0),
                number(3.0),
                number(5.0),
                number(7.0),
                Value::Blank,
            ]),
        ),
        (
            "=SORT(A2:C4,3,-1)",
            Value::Array(vec![
                vec![text("South"), text("Banana"), number(200.0)],
                vec![text("North"), text("Cherry"), number(150.0)],
                vec![text("North"), text("Apple"), number(100.0)],
            ]),
        ),
        (
            "=SORT({3,1,2},1,1,TRUE)",
            row(vec![number(1.0), number(2.0), number(3.0)]),
        ),
        (
            "=_xlfn._xlws.SORT({\"b\";1;TRUE;\"A\"})",
            column(vec![
                number(1.0),
                text("A"),
                text("b"),
                Value::Boolean(true),
            ]),
        ),
        ("=SORT({1,2},3)", error(ExcelError::Value)),
        ("=SORT({1,2},1,0)", error(ExcelError::Value)),
    ]);
}

#[test]
fn ifs_switch_and_ifna_return_the_chosen_argument() {
    cases(&[
        ("=IFS(C2>150,\"big\",C2>50,\"mid\")", text("mid")),
        ("=IFS(FALSE,1)", error(ExcelError::Na)),
        ("=IFS(1/0,1)", error(ExcelError::Div0)),
        ("=IFS(TRUE,1,1/0,2)", number(1.0)),
        ("=SWITCH(2,1,\"one\",2,\"two\")", text("two")),
        ("=SWITCH(3,1,\"one\",\"other\")", text("other")),
        ("=SWITCH(3,1,\"one\")", error(ExcelError::Na)),
        ("=SWITCH(A2,\"South\",1,\"North\",2)", number(2.0)),
        (
            "=IFNA(MATCH(\"West\",A2:A6,0),\"missing\")",
            text("missing"),
        ),
        ("=IFNA(1/0,\"x\")", error(ExcelError::Div0)),
        ("=IFNA(C2,0)", number(100.0)),
        ("=IFNA(D4,0)", number(0.0)),
    ]);
}

#[test]
fn concat_textjoin_and_value() {
    cases(&[
        ("=CONCAT(A2:B3)", text("NorthAppleSouthBanana")),
        ("=CONCAT(\"a\",1,TRUE)", text("a1TRUE")),
        ("=TEXTJOIN(\", \",TRUE,D2:D6)", text("5, 3, 2, 7")),
        ("=TEXTJOIN(\", \",FALSE,D2:D6)", text("5, 3, , 2, 7")),
        ("=TEXTJOIN({\"-\",\"+\"},TRUE,1,2,3)", text("1-2+3")),
        ("=TEXTJOIN(\",\",TRUE,C2,#N/A)", error(ExcelError::Na)),
        ("=VALUE(\"1,234.5\")", number(1234.5)),
        ("=VALUE(\" 42 \")", number(42.0)),
        ("=VALUE(\"25%\")", number(0.25)),
        ("=VALUE(D4)", number(0.0)),
        ("=VALUE(\"abc\")", error(ExcelError::Value)),
        ("=VALUE(TRUE)", error(ExcelError::Value)),
        ("=VALUE(\"1/2/2024\")", number(45293.0)),
        ("=VALUE(\"2024-01-02\")", number(45293.0)),
        ("=VALUE(\"2-Jan-2024\")", number(45293.0)),
        ("=VALUE(\"Jan 2, 2024\")", number(45293.0)),
        ("=VALUE(\"January 2024\")", number(45292.0)),
        ("=VALUE(\"12:00\")", number(0.5)),
        ("=VALUE(\"3 PM\")", number(0.625)),
        ("=VALUE(\"12:30 AM\")", number(30.0 / 1440.0)),
        ("=VALUE(\"1/2/2024 18:00\")", number(45293.75)),
        ("=VALUE(\"2/30/2024\")", error(ExcelError::Value)),
        ("=VALUE(\"13:00 PM\")", error(ExcelError::Value)),
    ]);
}

#[test]
fn working_days_skip_weekends_and_holidays() {
    cases(&[
        // 45292 is Monday 2024-01-01.
        ("=NETWORKDAYS(45292,45298)", number(5.0)),
        ("=NETWORKDAYS(45298,45292)", number(-5.0)),
        ("=NETWORKDAYS(45297,45298)", number(0.0)),
        (
            "=NETWORKDAYS(45292,45305,{45292,45293,45297,45293})",
            number(8.0),
        ),
        (
            "=NETWORKDAYS(DATE(2012,10,1),DATE(2013,3,1))",
            number(110.0),
        ),
        (
            "=NETWORKDAYS(DATE(2012,10,1),DATE(2013,3,1),{41235,41247,41295})",
            number(107.0),
        ),
        ("=NETWORKDAYS(45292,45298,\"x\")", error(ExcelError::Value)),
        ("=NETWORKDAYS(-1,45298)", error(ExcelError::Num)),
        ("=WORKDAY(DATE(2008,10,1),151)", number(39933.0)),
        (
            "=WORKDAY(DATE(2008,10,1),151,{39778,39786,39834})",
            number(39938.0),
        ),
        ("=WORKDAY(45292,-1)", number(45289.0)),
        ("=WORKDAY(45297,1)", number(45299.0)),
        ("=WORKDAY(45297,0)", number(45297.0)),
        ("=WORKDAY(45292,4,{45293})", number(45297.0 + 2.0)),
        ("=WORKDAY(1,-5)", error(ExcelError::Num)),
    ]);
}

#[test]
fn dated_cash_flows_rate_and_periods() {
    let flows = "{-10000,2750,4250,3250,2750}";
    let dates = "{39448,39508,39751,39859,39904}";
    close(&[
        (&format!("=XNPV(0.09,{flows},{dates})"), 2086.647602, 1e-3),
        (&format!("=XIRR({flows},{dates})"), 0.373_362_535, 1e-7),
        (&format!("=XIRR({flows},{dates},0.5)"), 0.373_362_535, 1e-7),
        ("=RATE(48,-200,8000)", 0.007_701_472_4, 1e-9),
        ("=RATE(48,-200,8000)*12", 0.092_417_669, 1e-8),
        ("=RATE(10,0,-100,200)", 2f64.powf(0.1) - 1.0, 1e-10),
        ("=NPER(0.01,-100,-1000,10000,1)", 59.673_865_7, 1e-6),
        ("=NPER(0.01,-100,-1000,10000)", 60.082_122_9, 1e-6),
        ("=NPER(0.01,-100,-1000)", -9.578_594_04, 1e-7),
        ("=NPER(0,-100,1000)", 10.0, 1e-12),
    ]);
    cases(&[
        (
            &format!("=XNPV(0.09,{flows},{{39448,39508}})"),
            error(ExcelError::Num),
        ),
        (
            "=XNPV(0.09,{-100,110},{39448,39400})",
            error(ExcelError::Num),
        ),
        (
            "=XNPV(0.09,{-100,\"a\"},{39448,39449})",
            error(ExcelError::Value),
        ),
        ("=XIRR({100,200},{39448,39508})", error(ExcelError::Num)),
        ("=RATE(0,-200,8000)", error(ExcelError::Num)),
        ("=NPER(0,0,100)", error(ExcelError::Num)),
        ("=NPER(0.1,-5,100)", error(ExcelError::Num)),
    ]);
}

#[test]
fn subtotal_applies_the_function_number() {
    cases(&[
        ("=SUBTOTAL(9,C2:C6)", number(800.0)),
        ("=SUBTOTAL(109,C2:C6)", number(800.0)),
        ("=SUBTOTAL(1,C2:C6)", number(160.0)),
        ("=SUBTOTAL(2,A1:D6)", number(9.0)),
        ("=SUBTOTAL(3,A1:D6)", number(23.0)),
        ("=SUBTOTAL(4,C2:C6,D2:D6)", number(300.0)),
        ("=SUBTOTAL(5,C2:C6)", number(50.0)),
        ("=SUBTOTAL(6,D2:D6)", number(210.0)),
        ("=SUBTOTAL(10,C2:C6)", number(9250.0)),
        ("=SUBTOTAL(11,C2:C6)", number(7400.0)),
        ("=SUBTOTAL(12,C2:C6)", error(ExcelError::Value)),
        ("=SUBTOTAL(9,{1,2})", error(ExcelError::Value)),
    ]);
}

/// The fixture with `Data` row 3 hidden and `Data!C6` holding a subtotal.
struct Hidden(MemorySource);

impl DataSource for Hidden {
    fn sheet_names(&self) -> impl Iterator<Item = &str> {
        self.0.sheet_names()
    }

    fn cell_value(&self, sheet: &str, row: u32, column: u32) -> Value {
        self.0.cell_value(sheet, row, column)
    }

    fn used_extent(&self, sheet: &str) -> Extent {
        self.0.used_extent(sheet)
    }

    fn date1904(&self) -> bool {
        false
    }

    fn row_hidden(&self, sheet: &str, row: u32) -> bool {
        sheet == "Data" && row == 3
    }

    fn is_subtotal(&self, sheet: &str, row: u32, column: u32) -> bool {
        sheet == "Data" && (row, column) == (6, 3)
    }
}

#[test]
fn subtotal_skips_hidden_rows_and_nested_subtotals() {
    let source = Hidden(MemorySource::fixture());
    let value = |formula: &str| evaluate(formula, &source, "Data").unwrap();
    // Amounts 100, 200 (hidden row), 150, 50, 300 (a subtotal cell).
    assert_eq!(value("=SUBTOTAL(9,C2:C6)"), number(500.0));
    assert_eq!(value("=SUBTOTAL(109,C2:C6)"), number(300.0));
    assert_eq!(value("=SUBTOTAL(103,A2:A6)"), number(4.0));
    assert_eq!(value("=SUBTOTAL(102,C:C)"), number(3.0));
    // Other functions see every cell.
    assert_eq!(value("=SUM(C2:C6)"), number(800.0));
}
