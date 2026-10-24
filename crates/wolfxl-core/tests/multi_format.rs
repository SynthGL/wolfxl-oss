//! Multi-format `Workbook::open` dispatch tests.
//!
//! Native format adapters retain typed values and sparse coordinates.
//! Delimited inputs remain a synthetic worksheet of strings.

use std::path::PathBuf;

use wolfxl_core::{infer_sheet_schema, CellValue, InferredType, SourceFormat, Workbook};

fn fixture(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join(name)
}

#[test]
fn opens_csv_as_single_synthetic_sheet() {
    let path = fixture("sample-minimal.csv");
    assert!(path.exists(), "csv fixture missing at {}", path.display());

    let mut wb = Workbook::open(&path).expect("open csv");
    assert_eq!(wb.format(), SourceFormat::Csv);

    let names = wb.sheet_names().to_vec();
    assert_eq!(
        names.len(),
        1,
        "csv should expose exactly one synthetic sheet, got {names:?}"
    );
    // Sheet name comes from the filename stem.
    assert_eq!(names[0], "sample-minimal");

    let sheet = wb.first_sheet().expect("read csv sheet");
    let (rows, cols) = sheet.dimensions();
    // Header row + 3 data rows, 5 columns per row.
    assert_eq!(rows, 4, "got {rows} rows");
    assert_eq!(cols, 5, "got {cols} cols");

    let headers = sheet.headers();
    assert_eq!(
        headers,
        vec!["Account", "Jan", "Feb", "Mar", "Q1"],
        "csv header row"
    );

    // CSV cells land as strings; schema inference is the source of truth
    // for "this column is actually numbers" (per invariant B4).
    for row in sheet.rows().iter().skip(1) {
        for cell in row.iter().skip(1) {
            match &cell.value {
                CellValue::String(_) => {}
                CellValue::Empty => {}
                other => panic!("expected string-valued CSV cell, got {other:?}"),
            }
        }
    }

    // Schema inference has to see through the string-valued cells to
    // classify Jan/Feb/Mar/Q1 as Int columns.
    let schema = infer_sheet_schema(&sheet);
    assert_eq!(schema.columns.len(), 5);
    assert_eq!(schema.columns[0].inferred_type, InferredType::String);
    for col in &schema.columns[1..] {
        assert_eq!(
            col.inferred_type,
            InferredType::Int,
            "column {:?} should infer Int, got {:?}",
            col.name,
            col.inferred_type
        );
    }
}

#[test]
fn opens_xls_with_native_backend() {
    let path = fixture("sample-minimal.xls");
    assert!(path.exists(), "xls fixture missing at {}", path.display());

    let mut wb = Workbook::open(&path).expect("open xls");
    assert_eq!(wb.format(), SourceFormat::Xls);

    let names = wb.sheet_names().to_vec();
    assert!(!names.is_empty(), "xls should have at least one sheet");
    assert_eq!(names[0], "P&L");

    let sheet = wb.first_sheet().expect("read xls sheet");
    let (rows, cols) = sheet.dimensions();
    assert!(
        rows >= 4 && cols >= 5,
        "dims should be >= 4x5, got {rows}x{cols}"
    );

    let headers = sheet.headers();
    assert_eq!(headers[0], "Account");
    assert_eq!(headers[1], "Jan");

    // BIFF8 numeric values remain numeric during schema inference.
    let schema = infer_sheet_schema(&sheet);
    assert_eq!(schema.columns[0].inferred_type, InferredType::String);
    assert_eq!(
        schema.columns[1].inferred_type,
        InferredType::Int,
        "Jan column should infer Int, got {:?}",
        schema.columns[1].inferred_type
    );
}

#[test]
fn opens_xlsb_with_native_backend() {
    // Fixture source: calamine's MIT-licensed `tests/date.xlsb`, copied
    // into this repo as a tiny binary workbook that exercises the xlsb
    // dispatch path without relying on a local Excel install.
    let path = fixture("sample-date.xlsb");
    assert!(path.exists(), "xlsb fixture missing at {}", path.display());

    let mut wb = Workbook::open(&path).expect("open xlsb");
    assert_eq!(wb.format(), SourceFormat::Xlsb);

    let names = wb.sheet_names().to_vec();
    assert_eq!(names, vec!["Sheet1"]);

    let sheet = wb.first_sheet().expect("read xlsb sheet");
    let (rows, cols) = sheet.dimensions();
    assert_eq!((rows, cols), (3, 2));

    let headers = sheet.headers();
    assert_eq!(headers[0], "2021-01-01");
    assert_eq!(headers[1], "15");

    let schema = infer_sheet_schema(&sheet);
    assert_eq!(schema.columns.len(), 2);
    assert_eq!(schema.columns[1].inferred_type, InferredType::Int);
    assert_eq!(
        sheet
            .row(0)
            .and_then(|row| row.first())
            .and_then(|cell| cell.number_format.as_deref()),
        Some("yyyy\\-mm\\-dd")
    );
}

#[test]
fn opens_ods_with_native_backend() {
    let path = fixture("sample-minimal.ods");
    assert!(path.exists(), "ods fixture missing at {}", path.display());

    let mut wb = Workbook::open(&path).expect("open ods");
    assert_eq!(wb.format(), SourceFormat::Ods);

    let names = wb.sheet_names().to_vec();
    assert!(!names.is_empty(), "ods should have at least one sheet");

    let sheet = wb.first_sheet().expect("read ods sheet");
    let (rows, cols) = sheet.dimensions();
    assert!(
        rows >= 4 && cols >= 5,
        "dims should be >= 4x5, got {rows}x{cols}"
    );

    let headers = sheet.headers();
    assert_eq!(headers[0], "Account");

    let schema = infer_sheet_schema(&sheet);
    assert_eq!(
        schema.columns[1].inferred_type,
        InferredType::Int,
        "Jan column on ods should infer Int, got {:?}",
        schema.columns[1].inferred_type
    );
}

#[test]
fn rejects_unsupported_extension() {
    // Point at something valid (no missing-file branch) with a bad ext.
    let bad_path = fixture("sample-minimal.csv").with_extension("unknown-ext");
    match Workbook::open(&bad_path) {
        Ok(_) => panic!("should reject .unknown-ext"),
        Err(e) => {
            let msg = format!("{e}");
            assert!(
                msg.contains("unsupported file extension"),
                "error should name the bad extension, got {msg:?}"
            );
        }
    }
}

#[test]
fn rejects_missing_extension() {
    let bare = fixture("sample-minimal").with_file_name("no-extension-here");
    match Workbook::open(&bare) {
        Ok(_) => panic!("should reject path with no extension"),
        Err(e) => assert!(format!("{e}").contains("cannot detect format"), "got {e}"),
    }
}

#[test]
fn styles_accessor_errors_for_non_xlsx_formats() {
    let mut wb = Workbook::open(fixture("sample-minimal.csv")).expect("open csv");
    match wb.styles() {
        Ok(_) => panic!("styles walker should error on non-xlsx"),
        Err(e) => assert!(format!("{e}").contains("only supports xlsx"), "got {e}"),
    }
}

#[test]
fn independently_authored_reader_fixtures_keep_their_sparse_coordinates() {
    let input = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../wolfxl-reader/qualification");
    for extension in ["xls", "ods"] {
        let mut workbook = Workbook::open(input.join(format!("fixtures.{extension}"))).unwrap();
        assert_eq!(workbook.sheet_names(), &["Values", "Second sheet"]);
        let values = workbook.sheet("Values").unwrap();
        assert_eq!(values.origin(), (2, 1));
        assert_eq!(values.cell_address(0, 0), "B3");
        assert_eq!(values.dimensions(), (1, 5));
        assert!(values.get_at_address("A1").is_none());
        assert!(matches!(&values.get_at_address("B3").unwrap().value,
            CellValue::String(text) if text == "Independent café 中文"));
        assert!(matches!(
            values.get_at_address("C3").unwrap().value,
            CellValue::Float(42.5)
        ));
        assert!(matches!(
            values.get_at_address("D3").unwrap().value,
            CellValue::Bool(true)
        ));
        assert!(
            matches!(values.get_at_address("E3").unwrap().value, CellValue::Date(date)
            if date == chrono::NaiveDate::from_ymd_opt(2024, 2, 29).unwrap())
        );
        if extension == "ods" {
            assert!(matches!(
                values.get_at_address("F3").unwrap().value,
                CellValue::Int(45)
            ));
        } else {
            // xlwt stores F3's FormulaValue as 03 00 00 00 00 00 ff ff.
            // MS-XLS 2.5.133 defines type 0x03 as a blank string, not Empty.
            assert!(matches!(
                &values.get_at_address("F3").unwrap().value,
                CellValue::String(text) if text.is_empty()
            ));
        }
        assert!(matches!(
            workbook.sheet("absent"),
            Err(wolfxl_core::Error::SheetNotFound(_))
        ));
    }
}
