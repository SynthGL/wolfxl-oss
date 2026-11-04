//! Structured references: parsing, and resolution against a source's tables.

use std::collections::BTreeMap;

use super::{
    evaluate, parse, CellArea, DataSource, ExcelError, Expr, Extent, ParseError, TableInfo,
    TableReference, TableRows, Value,
};

/// A `Data` sheet holding `Sales` at B2:E6 (a header row, three data rows,
/// and a totals row) and `Bare` at G1:G2 (no header or totals row), beside
/// an empty `Other` sheet.
struct TableSource {
    cells: BTreeMap<(u32, u32), Value>,
    sales_columns: Vec<String>,
    bare_columns: Vec<String>,
}

impl TableSource {
    fn new() -> Self {
        let text = |text: &str| Value::Text(text.into());
        let number = Value::Number;
        let rows = [
            [
                text("Region"),
                text("Amount"),
                text("Units"),
                text("Q1 [est]"),
            ],
            [text("North"), number(100.0), number(5.0), number(1.0)],
            [text("South"), number(200.0), number(3.0), number(2.0)],
            [text("East"), number(50.0), number(2.0), number(3.0)],
            [text("Total"), number(350.0), number(10.0), number(6.0)],
        ];
        let mut cells = BTreeMap::new();
        for (row, values) in (2..).zip(rows) {
            for (column, value) in (2..).zip(values) {
                cells.insert((row, column), value);
            }
        }
        cells.insert((1, 7), number(4.0));
        cells.insert((2, 7), number(6.0));
        Self {
            cells,
            sales_columns: ["Region", "Amount", "Units", "Q1 [est]"]
                .map(String::from)
                .to_vec(),
            bare_columns: vec!["Value".into()],
        }
    }

    fn on(&self, sheet: &str, formula: &str) -> Value {
        evaluate(formula, self, sheet).unwrap_or_else(|error| panic!("{formula}: {error}"))
    }
}

impl DataSource for TableSource {
    fn sheet_names(&self) -> impl Iterator<Item = &str> {
        ["Other", "Data"].into_iter()
    }

    fn cell_value(&self, sheet: &str, row: u32, column: u32) -> Value {
        match sheet {
            "Data" => self.cells.get(&(row, column)).cloned(),
            _ => None,
        }
        .unwrap_or(Value::Blank)
    }

    fn used_extent(&self, sheet: &str) -> Extent {
        match sheet {
            "Data" => Extent {
                rows: 6,
                columns: 7,
            },
            _ => Extent::default(),
        }
    }

    fn table(&self, name: &str) -> Option<TableInfo<'_>> {
        let (area, header_rows, totals_rows, columns) = if super::names_match(name, "Sales") {
            ((2, 2, 6, 5), 1, 1, &self.sales_columns)
        } else if super::names_match(name, "Bare") {
            ((1, 7, 2, 7), 0, 0, &self.bare_columns)
        } else {
            return None;
        };
        let (first_row, first_column, last_row, last_column) = area;
        Some(TableInfo {
            sheet: "Data",
            area: CellArea {
                first_row,
                first_column,
                last_row,
                last_column,
            },
            header_rows,
            totals_rows,
            columns,
        })
    }
}

fn table_reference(formula: &str) -> TableReference {
    match parse(formula).map(|formula| formula.expression) {
        Ok(Expr::Table(reference)) => reference,
        other => panic!("{formula}: {other:?}"),
    }
}

fn syntax_message(formula: &str) -> String {
    match parse(formula) {
        Err(ParseError::Syntax { message, .. }) => message,
        other => panic!("{formula}: {other:?}"),
    }
}

#[test]
fn structured_reference_grammar() {
    let columns = |first: &str, last: &str| Some((first.to_owned(), last.to_owned()));
    for (formula, rows, selected) in [
        ("T[]", TableRows::Data, None),
        ("T[Amount]", TableRows::Data, columns("Amount", "Amount")),
        ("T[[Amount]]", TableRows::Data, columns("Amount", "Amount")),
        (
            "T[Unit Count]",
            TableRows::Data,
            columns("Unit Count", "Unit Count"),
        ),
        ("T[#all]", TableRows::All, None),
        ("T[#Data]", TableRows::Data, None),
        ("T[#HEADERS]", TableRows::Headers, None),
        ("T[#Totals]", TableRows::Totals, None),
        ("T[[#Totals],[A]]", TableRows::Totals, columns("A", "A")),
        (
            "T[[#Headers],[#Data],[A]]",
            TableRows::HeadersData,
            columns("A", "A"),
        ),
        ("T[[#Data],[#Totals]]", TableRows::DataTotals, None),
        ("T[[A]:[B]]", TableRows::Data, columns("A", "B")),
        (
            "T[ [#Data] , [A] : [B] ]",
            TableRows::Data,
            columns("A", "B"),
        ),
        ("T[@A]", TableRows::ThisRow, columns("A", "A")),
        (
            "T[@[Unit Count]]",
            TableRows::ThisRow,
            columns("Unit Count", "Unit Count"),
        ),
        ("T[@]", TableRows::ThisRow, None),
        ("T[[#This Row],[A]]", TableRows::ThisRow, columns("A", "A")),
    ] {
        let reference = table_reference(formula);
        assert_eq!(reference.table, "T", "{formula}");
        assert_eq!(reference.rows, rows, "{formula}");
        assert_eq!(reference.columns, selected, "{formula}");
    }
}

#[test]
fn apostrophes_escape_brackets_hashes_and_apostrophes_in_column_names() {
    assert_eq!(
        table_reference("Sales['[Odd']' '#1 Bob''s]").columns,
        Some(("[Odd] #1 Bob's".to_owned(), "[Odd] #1 Bob's".to_owned()))
    );
    assert_eq!(
        table_reference("Sales[[Q1 '[est']]:[Q'#2]]").columns,
        Some(("Q1 [est]".to_owned(), "Q#2".to_owned()))
    );
}

#[test]
fn malformed_structured_references_are_syntax_errors() {
    assert!(syntax_message("SUM([Amount])").contains("needs a table name"));
    assert!(syntax_message("[Amount]").contains("needs a table name"));
    assert!(syntax_message("T[#Everything]").contains("unknown table item"));
    assert!(syntax_message("T[[#Headers],[#Totals]]").contains("combine only"));
    assert!(syntax_message("T[[#All],[#Data]]").contains("combine only"));
    assert!(syntax_message("T[[A],[B]]").contains("one column"));
    assert!(syntax_message("T[[A],[#Totals]]").contains("before the column"));
    assert!(syntax_message("T[Amount").contains("unterminated"));
    assert!(syntax_message("T[Am[ount]").contains("'["));
    assert!(syntax_message("T[[A]:B]").contains("after ':'"));
}

#[test]
fn structured_references_select_table_parts() {
    let source = TableSource::new();
    for (formula, expected) in [
        // The data body excludes the header and totals rows.
        ("=SUM(Sales[Amount])", Value::Number(350.0)),
        ("=SUM(Sales[])", Value::Number(366.0)),
        ("=SUM(Sales[#Data])", Value::Number(366.0)),
        ("=Sales[[#Totals],[Amount]]", Value::Number(350.0)),
        ("=Sales[[#Totals],[Amount]]*2", Value::Number(700.0)),
        ("=Sales[[#Headers],[Region]]", Value::Text("Region".into())),
        ("=COUNTA(Sales[#Headers])", Value::Number(4.0)),
        ("=COUNTA(Sales[#All])", Value::Number(20.0)),
        (
            "=COUNTA(Sales[[#Headers],[#Data],[Region]])",
            Value::Number(4.0),
        ),
        (
            "=SUM(Sales[[#Data],[#Totals],[Amount]])",
            Value::Number(700.0),
        ),
        (
            "=SUM(Sales[[#Data],[Amount]:[Units]])",
            Value::Number(360.0),
        ),
        ("=SUM(Sales[[Units]:[Amount]])", Value::Number(360.0)),
        ("=SUM(SALES[ [amount] ])", Value::Number(350.0)),
        ("=SUM(Sales[Q1 '[est']])", Value::Number(6.0)),
        ("=INDEX(Sales[Amount],2)", Value::Number(200.0)),
        (
            "=SUMIF(Sales[Region],\"South\",Sales[Amount])",
            Value::Number(200.0),
        ),
        ("=SUM(Bare[Value])", Value::Number(10.0)),
        ("=SUM(Bare[#All])", Value::Number(10.0)),
    ] {
        assert_eq!(source.on("Data", formula), expected, "{formula}");
        // Table names are workbook-global, whatever the current sheet.
        assert_eq!(source.on("Other", formula), expected, "{formula} on Other");
    }
}

#[test]
fn unresolvable_structured_references_are_errors() {
    let source = TableSource::new();
    for (formula, error) in [
        ("=SUM(NoSuchTable[Amount])", ExcelError::Ref),
        ("=SUM(Sales[NoSuchColumn])", ExcelError::Ref),
        ("=SUM(Sales[[Amount]:[NoSuchColumn]])", ExcelError::Ref),
        ("=SUM(Bare[#Headers])", ExcelError::Ref),
        ("=SUM(Bare[#Totals])", ExcelError::Ref),
        ("=SUM(Bare[[#Data],[#Totals]])", ExcelError::Ref),
        ("=Sales[@Amount]", ExcelError::Value),
        ("=Sales[[#This Row],[Amount]]", ExcelError::Value),
    ] {
        assert_eq!(source.on("Data", formula), Value::Error(error), "{formula}");
    }
}
