//! Integration tests for `wolfxl peek --add NAME=FORMULA`: fill-down with
//! relative and `$` references, functions, chained columns, error cells,
//! range results, filtering and selecting on added columns, every export
//! format and `--max-columns`, CSV and XLSX input from a path or standard
//! input, sheets whose block starts away from A1, header-only and empty
//! sheets, and the exit-2 refusals.

use std::fs;
use std::io::{Cursor, Write};
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicUsize, Ordering};
use std::time::{SystemTime, UNIX_EPOCH};

use assert_cmd::Command;
use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipWriter};

/// Four parts; `gear` has no quantity, so per-unit formulas divide by zero.
const ITEMS: &str = "\
Item,Qty,Price,Cost
bolt,10,0.1,0.5
nut,100,0.25,20
gear,0,2,3
";

static NEXT_TEMP: AtomicUsize = AtomicUsize::new(0);

fn fixture(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures")
        .join(name)
}

struct Run {
    code: i32,
    stdout: String,
    stderr: String,
}

fn peek(args: &[&str], stdin: Option<&[u8]>) -> Run {
    let mut command = Command::cargo_bin("wolfxl").expect("wolfxl bin built");
    command.arg("peek").args(args);
    if let Some(stdin) = stdin {
        command.write_stdin(stdin);
    }
    let out = command.output().expect("run wolfxl peek");
    Run {
        code: out.status.code().expect("exit code"),
        stdout: String::from_utf8(out.stdout).expect("utf-8 stdout"),
        stderr: String::from_utf8(out.stderr).expect("utf-8 stderr"),
    }
}

/// `args` with exit 0 expected; the output.
fn ok(args: &[&str], stdin: Option<&[u8]>) -> String {
    let run = peek(args, stdin);
    assert_eq!(run.code, 0, "wolfxl peek {args:?}: stderr={}", run.stderr);
    run.stdout
}

/// `csv` piped as CSV and exported as CSV, with `args` appended.
fn piped_csv(csv: &str, args: &[&str]) -> String {
    let mut full = vec!["-", "--from", "csv", "-e", "csv"];
    full.extend_from_slice(args);
    ok(&full, Some(csv.as_bytes()))
}

/// `ITEMS` piped as CSV and exported as CSV, with `args` appended.
fn items(args: &[&str]) -> String {
    piped_csv(ITEMS, args)
}

/// `ITEMS` piped with `args`; expect exit 2, nothing on stdout, and the
/// stderr message.
fn refused(args: &[&str]) -> String {
    let mut full = vec!["-", "--from", "csv", "-e", "csv"];
    full.extend_from_slice(args);
    let run = peek(&full, Some(ITEMS.as_bytes()));
    assert_eq!(run.code, 2, "wolfxl peek {full:?}: stdout={}", run.stdout);
    assert_eq!(run.stdout, "");
    assert!(run.stderr.starts_with("error: "), "stderr={}", run.stderr);
    run.stderr
}

fn json(text: &str) -> serde_json::Value {
    serde_json::from_str(text).expect("peek JSON parses")
}

#[test]
fn arithmetic_formula_fills_down_and_prints_like_calc() {
    let expected = "\
Item,Qty,Price,Cost,Margin
bolt,10,0.1,0.5,0.5
nut,100,0.25,20,5
gear,0,2,3,-3
";
    assert_eq!(items(&["--add", "Margin=B2*C2-D2"]), expected);
    // Everything after the first `=` is the formula, with or without its own `=`.
    assert_eq!(items(&["--add", "Margin==B2*C2-D2"]), expected);
}

#[test]
fn if_and_concatenation_leave_string_literals_and_function_names_alone() {
    let expected = "\
Item,Qty,Price,Cost,Size,Revenue,Tag
bolt,10,0.1,0.5,small,1,B2bolt
nut,100,0.25,20,nut,25,B2nut
gear,0,2,3,small,-,B2gear
";
    assert_eq!(
        items(&[
            "--add",
            "Size=IF(B2>=100,A2,\"small\")",
            "--add",
            "Revenue=IF(B2>0,B2*C2,\"-\")",
            "--add",
            "Tag=\"B2\"&A2",
        ]),
        expected
    );
}

#[test]
fn absolute_rows_stay_fixed_and_whole_columns_are_unchanged() {
    let expected = "\
Item,Qty,Price,Cost,Scaled,Running,Share
bolt,10,0.1,0.5,1,10,0.0909090909090909
nut,100,0.25,20,10,110,0.909090909090909
gear,0,2,3,0,110,0
";
    assert_eq!(
        items(&[
            "--add",
            "Scaled=B2*$C$2",
            "--add",
            "Running=SUM($B$2:B2)",
            "--add",
            "Share=B2/SUM(B:B)",
        ]),
        expected
    );
}

#[test]
fn later_columns_reference_earlier_ones_by_letter() {
    // Revenue takes column E, the first free letter after A:D.
    let expected = "\
Item,Qty,Price,Cost,Revenue,Double
bolt,10,0.1,0.5,1,2
nut,100,0.25,20,25,50
gear,0,2,3,0,0
";
    assert_eq!(
        items(&["--add", "Revenue=B2*C2", "--add", "Double=E2*2"]),
        expected
    );
}

#[test]
fn a_column_sees_the_rows_above_it_in_its_own_column() {
    // E1 is the added column's header cell, blank to formulas.
    assert_eq!(
        items(&["--add", "Running=B2+E1"]),
        "\
Item,Qty,Price,Cost,Running
bolt,10,0.1,0.5,10
nut,100,0.25,20,110
gear,0,2,3,110
"
    );
}

#[test]
fn formula_errors_stay_in_their_cells() {
    assert_eq!(
        items(&["--add", "Unit=D2/B2"]),
        "\
Item,Qty,Price,Cost,Unit
bolt,10,0.1,0.5,0.05
nut,100,0.25,20,0.2
gear,0,2,3,ERROR: #DIV/0!
"
    );
    let document = json(&ok(
        &["-", "--from", "csv", "-e", "json", "--add", "Unit=D2/B2"],
        Some(ITEMS.as_bytes()),
    ));
    assert_eq!(document["data"][2][4], "#DIV/0!");
    // An error in an earlier added column propagates through a later one.
    assert_eq!(
        items(&[
            "--add",
            "Unit=D2/B2",
            "--add",
            "Twice=E2*2",
            "--select",
            "Twice"
        ]),
        "Twice\n0.1\n0.4\nERROR: #DIV/0!\n"
    );
}

#[test]
fn a_range_result_is_a_value_error_without_implicit_intersection() {
    assert_eq!(
        items(&["--add", "Span=B2:C2", "--select", "Item,Span"]),
        "\
Item,Span
bolt,ERROR: #VALUE!
nut,ERROR: #VALUE!
gear,ERROR: #VALUE!
"
    );
}

#[test]
fn formulas_read_typed_fields_while_output_keeps_source_spelling() {
    let ledger = "Code,Amount\n007,1.50\n010,2.25\n";
    assert_eq!(
        piped_csv(ledger, &["--add", "Next=A2+1", "--add", "Double=B2*2"]),
        "\
Code,Amount,Next,Double
007,1.50,8,3
010,2.25,11,4.5
"
    );
}

#[test]
fn where_and_select_see_added_columns() {
    assert_eq!(
        items(&[
            "--add",
            "Margin=B2*C2-D2",
            "--where",
            "Margin>0",
            "--select",
            "Item,Margin",
        ]),
        "Item,Margin\nbolt,0.5\nnut,5\n"
    );
    assert_eq!(
        items(&[
            "--add",
            "Unit=D2/B2",
            "--where",
            "Unit=#DIV/0!",
            "--select",
            "Item,Unit",
        ]),
        "Item,Unit\ngear,ERROR: #DIV/0!\n"
    );
    assert_eq!(
        items(&[
            "--add",
            "Revenue=B2*C2",
            "--add",
            "Double=E2*2",
            "--where",
            "double>=50",
            "--select",
            "Double,Revenue,Item",
        ]),
        "Double,Revenue,Item\n50,25,nut\n"
    );
}

#[test]
fn every_export_format_shows_the_added_column() {
    let piped = |format: &str| {
        ok(
            &["-", "--from", "csv", "-e", format, "--add", "Revenue=B2*C2"],
            Some(ITEMS.as_bytes()),
        )
    };
    assert_eq!(
        piped("text"),
        "Item\tQty\tPrice\tCost\tRevenue\nbolt\t10\t0.1\t0.5\t1\nnut\t100\t0.25\t20\t25\ngear\t0\t2\t3\t0\n"
    );
    let markdown = piped("markdown");
    assert!(
        markdown.starts_with("| Item | Qty | Price | Cost | Revenue |\n"),
        "{markdown}"
    );
    assert!(
        markdown.contains("| nut | 100 | 0.25 | 20 | 25 |"),
        "{markdown}"
    );
    let document = json(&piped("json"));
    assert_eq!(document["columns"], 5);
    assert_eq!(document["headers"][4], "Revenue");
    assert_eq!(document["data"][1][4], "25");
    let boxed = piped("box");
    assert!(boxed.contains("3 rows × 5 columns"), "{boxed}");
    assert!(boxed.contains("Revenue"), "{boxed}");
}

#[test]
fn max_columns_caps_the_sheet_with_its_added_columns() {
    let document = json(&ok(
        &[
            "-",
            "--from",
            "csv",
            "-e",
            "json",
            "--max-columns",
            "4",
            "--add",
            "Revenue=B2*C2",
        ],
        Some(ITEMS.as_bytes()),
    ));
    assert_eq!(document["columns"], 5);
    assert_eq!(
        document["headers"],
        serde_json::json!(["Item", "Qty", "Price", "Cost"])
    );
    assert_eq!(document["data"][1].as_array().map(Vec::len), Some(4));
}

#[test]
fn csv_file_input_takes_ranges_relative_to_each_row() {
    let path = fixture("sample-minimal.csv");
    assert_eq!(
        ok(
            &[
                path.to_str().unwrap(),
                "-e",
                "csv",
                "--add",
                "Growth=C2-B2",
                "--add",
                "Check=SUM(B2:D2)-E2",
            ],
            None,
        ),
        "\
Account,Jan,Feb,Mar,Q1,Growth,Check
Revenue,420000,445000,512000,1377000,25000,0
COGS,277000,295000,329500,901500,18000,0
Gross,143000,150000,182500,475500,7000,0
"
    );
}

#[test]
fn a_header_only_sheet_gets_the_header_and_no_rows() {
    assert_eq!(
        piped_csv("Item,Qty\n", &["--add", "Double=B2*2"]),
        "Item,Qty,Double\n"
    );
}

/// A private directory, unique per test, removed on drop.
struct TempDir(PathBuf);

impl TempDir {
    fn new(label: &str) -> Self {
        let stamp = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .map(|elapsed| elapsed.as_nanos())
            .unwrap_or_default();
        let path = std::env::temp_dir().join(format!(
            "wolfxl-peek-add-{label}-{}-{}-{stamp}",
            std::process::id(),
            NEXT_TEMP.fetch_add(1, Ordering::Relaxed)
        ));
        fs::create_dir(&path).expect("create temp dir");
        Self(path)
    }

    fn path(&self, name: &str) -> PathBuf {
        self.0.join(name)
    }
}

impl Drop for TempDir {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.0);
    }
}

const ROOT_RELS: &str = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#;

const STYLES: &str = r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>"#;

/// A workbook of `sheets` (name, `<sheetData>` contents) in `directory`.
fn workbook(directory: &TempDir, sheets: &[(&str, &str)]) -> PathBuf {
    let mut content_types = String::from(
        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>"#,
    );
    let mut book = String::from(
        r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>"#,
    );
    let mut rels = String::from(
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    );
    let mut members = Vec::new();
    for (index, (name, data)) in sheets.iter().enumerate() {
        let number = index + 1;
        content_types.push_str(&format!(
            r#"<Override PartName="/xl/worksheets/sheet{number}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#
        ));
        book.push_str(&format!(
            r#"<sheet name="{name}" sheetId="{number}" r:id="rId{number}"/>"#
        ));
        rels.push_str(&format!(
            r#"<Relationship Id="rId{number}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{number}.xml"/>"#
        ));
        members.push((
            format!("xl/worksheets/sheet{number}.xml"),
            format!(
                r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>{data}</sheetData></worksheet>"#
            ),
        ));
    }
    content_types.push_str("</Types>");
    book.push_str("</sheets></workbook>");
    rels.push_str(&format!(
        r#"<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#,
        sheets.len() + 1
    ));
    members.extend([
        ("[Content_Types].xml".to_owned(), content_types),
        ("_rels/.rels".to_owned(), ROOT_RELS.to_owned()),
        ("xl/workbook.xml".to_owned(), book),
        ("xl/_rels/workbook.xml.rels".to_owned(), rels),
        ("xl/styles.xml".to_owned(), STYLES.to_owned()),
    ]);

    let path = directory.path("book.xlsx");
    let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
    let options = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
    for (name, text) in &members {
        writer
            .start_file(name.as_str(), options)
            .expect("start member");
        writer.write_all(text.as_bytes()).expect("write member");
    }
    let bytes = writer.finish().expect("finish package").into_inner();
    fs::write(&path, bytes).expect("write package");
    path
}

/// Total in D is a formula whose stored result for `nut` (999) disagrees with
/// B3*C3, so a formula reading D sees the stored value, not a recalculation.
const DATA_SHEET: &str = r#"<row r="1"><c r="A1" t="inlineStr"><is><t>Item</t></is></c><c r="B1" t="inlineStr"><is><t>Qty</t></is></c><c r="C1" t="inlineStr"><is><t>Price</t></is></c><c r="D1" t="inlineStr"><is><t>Total</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>bolt</t></is></c><c r="B2"><v>2</v></c><c r="C2"><v>3</v></c><c r="D2"><f>B2*C2</f><v>6</v></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>nut</t></is></c><c r="B3"><v>4</v></c><c r="C3"><v>5</v></c><c r="D3"><f>B3*C3</f><v>999</v></c></row>"#;

const RATES_SHEET: &str = r#"<row r="1"><c r="A1" t="inlineStr"><is><t>Tier</t></is></c><c r="B1" t="inlineStr"><is><t>Rate</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>low</t></is></c><c r="B2"><v>10</v></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>high</t></is></c><c r="B3"><v>100</v></c></row>"#;

/// The two-sheet `Data` + `Rates` workbook.
fn data_and_rates(directory: &TempDir) -> PathBuf {
    workbook(directory, &[("Data", DATA_SHEET), ("Rates", RATES_SHEET)])
}

const XLSX_ADDS: [&str; 6] = [
    "--add",
    "Check=D2+1",
    "--add",
    "Scaled=B2*Rates!B2",
    "--add",
    "Fixed=B2*Rates!$B$2",
];

const XLSX_EXPECTED: &str = "\
Item,Qty,Price,Total,Check,Scaled,Fixed
bolt,2,3,6,7,20,20
nut,4,5,999,1000,400,40
";

#[test]
fn xlsx_formulas_see_stored_values_and_sheet_qualified_references() {
    let directory = TempDir::new("xlsx");
    let path = data_and_rates(&directory);
    let path = path.to_str().unwrap();

    let mut args = vec![path, "-e", "csv"];
    args.extend_from_slice(&XLSX_ADDS);
    assert_eq!(ok(&args, None), XLSX_EXPECTED);

    let mut args = vec![path, "-e", "json"];
    args.extend_from_slice(&XLSX_ADDS);
    let document = json(&ok(&args, None));
    assert_eq!(
        document["data"],
        serde_json::json!([
            ["bolt", 2, 3, 6, "7", "20", "20"],
            ["nut", 4, 5, 999, "1000", "400", "40"]
        ])
    );

    let mut args = vec![
        path,
        "-e",
        "csv",
        "--where",
        "Check>=1000",
        "--select",
        "Item,Check",
    ];
    args.extend_from_slice(&XLSX_ADDS);
    assert_eq!(ok(&args, None), "Item,Check\nnut,1000\n");
}

#[test]
fn piped_xlsx_matches_the_path() {
    let directory = TempDir::new("stdin");
    let bytes = fs::read(data_and_rates(&directory)).expect("read workbook");
    let mut args = vec!["-", "--from", "xlsx", "-e", "csv"];
    args.extend_from_slice(&XLSX_ADDS);
    assert_eq!(ok(&args, Some(bytes.as_slice())), XLSX_EXPECTED);
}

/// A block at B3:E5 whose `Bad` column holds a cached `#DIV/0!` result.
const OFFSET_SHEET: &str = r#"<row r="3"><c r="B3" t="inlineStr"><is><t>Item</t></is></c><c r="C3" t="inlineStr"><is><t>Qty</t></is></c><c r="D3" t="inlineStr"><is><t>Price</t></is></c><c r="E3" t="inlineStr"><is><t>Bad</t></is></c></row><row r="4"><c r="B4" t="inlineStr"><is><t>bolt</t></is></c><c r="C4"><v>2</v></c><c r="D4"><v>3</v></c><c r="E4" t="e"><f>1/0</f><v>#DIV/0!</v></c></row><row r="5"><c r="B5" t="inlineStr"><is><t>nut</t></is></c><c r="C5"><v>4</v></c><c r="D5"><v>5</v></c><c r="E5"><v>1</v></c></row>"#;

#[test]
fn a_block_away_from_a1_keeps_its_sheet_coordinates() {
    let directory = TempDir::new("origin");
    let path = workbook(&directory, &[("Offset", OFFSET_SHEET), ("Empty", "")]);
    let path = path.to_str().unwrap();
    // The first data row is sheet row 4 and the first free column is F.
    assert_eq!(
        ok(
            &[
                path,
                "-e",
                "csv",
                "--add",
                "Total=C4*D4",
                "--add",
                "Twice=F4*2",
                "--add",
                "Anchored=C4*$C$4",
                "--add",
                "Over=E4+1",
                "--add",
                "Qtys=SUM(C:C)",
            ],
            None,
        ),
        "\
Item,Qty,Price,Bad,Total,Twice,Anchored,Over,Qtys
bolt,2,3,ERROR: Div0,6,12,4,ERROR: #DIV/0!,6
nut,4,5,1,20,40,8,2,6
"
    );
    let document = json(&ok(&[path, "-e", "json", "--add", "Over=E4+1"], None));
    assert_eq!(document["data"][0][3], "Div0");
    assert_eq!(document["data"][0][4], "#DIV/0!");
    assert_eq!(document["data"][1][4], "2");
}

#[test]
fn an_empty_sheet_gets_only_the_added_headers() {
    let directory = TempDir::new("empty");
    let path = workbook(&directory, &[("Offset", OFFSET_SHEET), ("Empty", "")]);
    let document = json(&ok(
        &[
            path.to_str().unwrap(),
            "--sheet",
            "Empty",
            "-e",
            "json",
            "--add",
            "Calc=A2*2",
        ],
        None,
    ));
    assert_eq!(document["rows"], 0);
    assert_eq!(document["columns"], 1);
    assert_eq!(document["headers"], serde_json::json!(["Calc"]));
    assert_eq!(document["data"], serde_json::json!([]));
}

#[test]
fn malformed_formulas_exit_two_naming_the_argument() {
    let stderr = refused(&["--add", "Bad=SUM(1,"]);
    assert!(stderr.contains("Bad=SUM(1,"), "{stderr}");
    assert!(stderr.contains("cannot parse"), "{stderr}");

    let stderr = refused(&["--add", "Nothing"]);
    assert!(stderr.contains("expected NAME=FORMULA"), "{stderr}");
    let stderr = refused(&["--add", "=B2*2"]);
    assert!(stderr.contains("missing the column name"), "{stderr}");
    let stderr = refused(&["--add", "Empty="]);
    assert!(stderr.contains("missing the formula"), "{stderr}");
    let stderr = refused(&["--add", "Empty=="]);
    assert!(stderr.contains("missing the formula"), "{stderr}");
}

#[test]
fn unsupported_functions_exit_two_naming_the_argument_and_function() {
    for (spec, function) in [
        ("Size=IF(B2>=100,PROPER(A2),\"small\")", "PROPER"),
        ("Log=IF(B2>0,LOG10(B2),\"-\")", "LOG10"),
        ("Pivot=GETPIVOTDATA(\"Qty\",A1)", "GETPIVOTDATA"),
    ] {
        let stderr = refused(&["--add", "Margin=B2*C2-D2", "--add", spec]);
        assert!(stderr.contains(&format!("{spec:?}")), "{stderr}");
        assert!(
            stderr.contains(&format!("unsupported function {function}")),
            "{stderr}"
        );
        assert!(stderr.contains("wolfxl list-functions"), "{stderr}");
    }
}

#[test]
fn taken_column_names_exit_two() {
    let stderr = refused(&["--add", "qty=B2*2"]);
    assert!(stderr.contains("qty=B2*2"), "{stderr}");
    assert!(stderr.contains("already exists"), "{stderr}");

    let stderr = refused(&["--add", "Total=B2*C2", "--add", "TOTAL=E2*2"]);
    assert!(stderr.contains("TOTAL=E2*2"), "{stderr}");
    assert!(stderr.contains("already exists"), "{stderr}");
}

#[test]
fn where_on_an_unknown_column_still_exits_two_after_adding() {
    let stderr = refused(&["--add", "Revenue=B2*C2", "--where", "Profit>0"]);
    assert!(
        stderr.contains("Revenue"),
        "available headers listed: {stderr}"
    );
}

#[test]
fn interactive_refusal_precedes_add_validation() {
    let file = fixture("sample-minimal.csv");
    let file = file.to_str().expect("utf-8 fixture path");
    for spec in ["Bad=SUM(1,", "Jan=B2*2", "Growth=C2-B2"] {
        let run = peek(&[file, "-i", "--add", spec], None);
        assert_eq!(run.code, 2, "{spec}: stderr={}", run.stderr);
        assert_eq!(run.stdout, "", "{spec} draws nothing");
        assert!(
            run.stderr.contains("needs a terminal"),
            "{spec}: stderr={}",
            run.stderr
        );
    }
}

#[test]
fn money_deals_csv_columns_and_filters_use_the_same_numbers() {
    let deals = "Rep,Cost,Revenue\nana,9000,\"$12,500.00\"\nbea,8000,\"$6,500.00\"\n";
    assert_eq!(
        piped_csv(deals, &["--add", "Margin=C2-B2", "--select", "Rep,Margin"]),
        "Rep,Margin\nana,3500\nbea,-1500\n"
    );
    assert_eq!(
        piped_csv(
            deals,
            &[
                "--add",
                "Margin=C2-B2",
                "--where",
                "Revenue>=$10,000",
                "--select",
                "Rep,Margin"
            ]
        ),
        "Rep,Margin\nana,3500\n"
    );
}

#[test]
fn money_ledger_csv_running_totals_include_accounting_negatives() {
    let ledger = "Rep,Amount\nana,100.00\nbea,250.50\ncam,(75.00)\n";
    assert_eq!(
        piped_csv(ledger, &["--add", "Running=SUM($B$2:B2)"]),
        "Rep,Amount,Running\nana,100.00,100\nbea,250.50,350.5\ncam,(75.00),275.5\n"
    );
}

#[test]
fn money_text_in_xlsx_stays_text_in_ranges_but_coerces_in_arithmetic() {
    let directory = TempDir::new("money-text");
    let data = r#"<row r="1"><c r="A1" t="inlineStr"><is><t>Amount</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>$1,200</t></is></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>(75.00)</t></is></c></row>"#;
    let path = workbook(&directory, &[("Money", data)]);
    let output = ok(
        &[
            path.to_str().unwrap(),
            "-e",
            "json",
            "--add",
            "Total=SUM(A:A)",
            "--add",
            "Numeric=A2+0",
        ],
        None,
    );
    let document = json(&output);
    assert_eq!(document["data"][0][0], "$1,200");
    assert_eq!(document["data"][1][0], "(75.00)");
    assert_eq!(document["data"][0][1], "0");
    assert_eq!(document["data"][1][1], "0");
    assert_eq!(document["data"][0][2], "1200");
    assert_eq!(document["data"][1][2], "-75");
}
