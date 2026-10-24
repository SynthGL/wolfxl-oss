// SPDX-License-Identifier: MIT
// Copyright (c) 2026 Wolfie

//! Reader-only smoke against independently produced, committed XLS/ODS inputs.
//! Expected constants follow generate_fixtures.py / expected.json, not another reader.
use std::path::Path;
use wolfxl_reader::{
    Cell, CellDataType, CellValue, NativeOdsBook, NativeXlsBook, SheetState, WorksheetData,
};

fn cell<'a>(sheet: &'a WorksheetData, coordinate: &str) -> &'a Cell {
    sheet
        .cells
        .iter()
        .find(|cell| cell.coordinate == coordinate)
        .expect(coordinate)
}

fn values(sheet: &WorksheetData) {
    assert_eq!(sheet.dimension.as_deref(), Some("B3:F3"));
    assert_eq!(
        cell(sheet, "B3").value,
        CellValue::String("Independent café 中文".into())
    );
    assert_eq!(cell(sheet, "C3").value, CellValue::Number(42.5));
    assert_eq!(cell(sheet, "D3").value, CellValue::Bool(true));
    assert_eq!((cell(sheet, "B3").row, cell(sheet, "B3").col), (3, 2));
}

fn second(sheet: &WorksheetData) {
    assert_eq!(sheet.dimension.as_deref(), Some("A1:A1"));
    assert_eq!(cell(sheet, "A1").value, CellValue::String("last".into()));
}

fn main() {
    let root = Path::new(env!("CARGO_MANIFEST_DIR"));
    let xls = NativeXlsBook::open_path(root.join("fixtures.xls")).expect("XLS path");
    assert_eq!(xls.sheet_names(), ["Values", "Second sheet"]);
    assert_eq!(xls.sheet_state("Values").unwrap(), SheetState::Visible);
    assert!(!xls.date1904());
    let xls_sheet = xls.worksheet("Values").expect("XLS values");
    values(&xls_sheet);
    assert_eq!(cell(&xls_sheet, "E3").value, CellValue::Number(45351.0));
    assert_eq!(
        xls.number_format_for_style_id(cell(&xls_sheet, "E3").style_id.unwrap()),
        Some("yyyy-mm-dd")
    );
    assert_eq!(
        cell(&xls_sheet, "F3").formula.as_deref(),
        Some("SUM(C3,2.5)")
    );
    second(&xls.worksheet("Second sheet").unwrap());
    let xls_bytes =
        NativeXlsBook::open_bytes(std::fs::read(root.join("fixtures.xls")).unwrap()).unwrap();
    assert_eq!(xls_bytes.worksheet("Values").unwrap(), xls_sheet);

    let ods = NativeOdsBook::open_path(root.join("fixtures.ods")).expect("ODS path");
    assert_eq!(ods.sheet_names(), ["Values", "Second sheet"]);
    assert_eq!(ods.sheet_state("Values").unwrap(), SheetState::Visible);
    assert!(!ods.date1904());
    let ods_sheet = ods.worksheet("Values").expect("ODS values");
    values(&ods_sheet);
    assert_eq!(cell(&ods_sheet, "E3").data_type, CellDataType::DateIso);
    assert_eq!(
        cell(&ods_sheet, "E3").value,
        CellValue::String("2024-02-29".into())
    );
    assert_eq!(
        cell(&ods_sheet, "F3").formula.as_deref(),
        Some("SUM(C3,2.5)")
    );
    assert_eq!(cell(&ods_sheet, "F3").value, CellValue::Number(45.0));
    second(&ods.worksheet("Second sheet").unwrap());
    let ods_bytes =
        NativeOdsBook::open_bytes(std::fs::read(root.join("fixtures.ods")).unwrap()).unwrap();
    assert_eq!(ods_bytes.worksheet("Values").unwrap(), ods_sheet);
    println!("{{\"scope\":\"Community reader-only\",\"xls_path_bytes\":\"passed\",\"ods_path_bytes\":\"passed\",\"sparse_values_dates_formula_cache\":\"passed\"}}");
}
