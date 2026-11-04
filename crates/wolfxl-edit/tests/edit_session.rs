use sha2::{Digest, Sha256};
use std::collections::BTreeMap;
use std::io::{Cursor, Read, Write};
use wolfxl_edit::{CellEdit, EditError, EditSession, EditValue};
use wolfxl_reader::{CellValue, NativeXlsxBook};
use zip::{write::SimpleFileOptions, ZipArchive, ZipWriter};

const SHEET: &str = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:C4"/><sheetData><row r="1"><c r="A1" s="1"><v>7</v></c><c r="B1" t="s"><v>0</v></c></row><row r="2"><c r="A2"><v>11</v></c><c r="B2"><f>A2+Other!A2</f><v>13</v></c></row><row r="4"><c r="C4"><v>9</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="B3:C3"/></mergeCells><dataValidations count="1"><dataValidation type="whole" sqref="A2:A4"><formula1>1</formula1><formula2>99</formula2></dataValidation></dataValidations></worksheet>"#;
const STYLES: &str = r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font/><font><b/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>"#;

fn package() -> Vec<u8> {
    let mut zip = ZipWriter::new(Cursor::new(Vec::new()));
    for (name, body) in [
        (
            "[Content_Types].xml",
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/></Types>"#,
        ),
        (
            "_rels/.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#,
        ),
        (
            "xl/workbook.xml",
            r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Other" sheetId="2" r:id="rId2"/></sheets><definedNames><definedName name="Amount">Data!$A$2</definedName><definedName name="Local" localSheetId="0">$A$2</definedName><definedName name="OtherLocal" localSheetId="1">$A$2</definedName></definedNames><calcPr calcId="123"/></workbook>"#,
        ),
        (
            "xl/_rels/workbook.xml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>"#,
        ),
        ("xl/worksheets/sheet1.xml", SHEET),
        (
            "xl/worksheets/sheet2.xml",
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><f>Data!A2+A2+'Data'!$A$2</f><v>22</v></c></row></sheetData></worksheet>"#,
        ),
        ("xl/styles.xml", STYLES),
        (
            "xl/sharedStrings.xml",
            r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>shared original</t></si></sst>"#,
        ),
        (
            "xl/calcChain.xml",
            r#"<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="B2" i="1"/></calcChain>"#,
        ),
        (
            "docProps/app.xml",
            r#"<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>2</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="2" baseType="lpstr"><vt:lpstr>Data</vt:lpstr><vt:lpstr>Other</vt:lpstr></vt:vector></TitlesOfParts></Properties>"#,
        ),
        (
            "customXml/item1.xml",
            "<unknown untouched='yes'>&#65;  preserved</unknown>",
        ),
        ("xl/vbaProject.bin", "opaque VBA payload\0\u{1}\u{2}"),
    ] {
        zip.start_file(name, SimpleFileOptions::default()).unwrap();
        zip.write_all(body.as_bytes()).unwrap();
    }
    zip.finish().unwrap().into_inner()
}
fn parts(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let mut archive = ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut parts = BTreeMap::new();
    for index in 0..archive.len() {
        let mut entry = archive.by_index(index).unwrap();
        let name = entry.name().unwrap().into_owned();
        let mut bytes = Vec::new();
        entry.read_to_end(&mut bytes).unwrap();
        parts.insert(name, bytes);
    }
    parts
}
fn repack(entries: BTreeMap<String, Vec<u8>>) -> Vec<u8> {
    let mut archive = ZipWriter::new(Cursor::new(Vec::new()));
    for (name, bytes) in entries {
        archive
            .start_file(name, SimpleFileOptions::default())
            .unwrap();
        archive.write_all(&bytes).unwrap();
    }
    archive.finish().unwrap().into_inner()
}
/// The fixture package without its calculation chain, and optionally without formulas.
fn package_without_calc_chain(keep_formulas: bool) -> Vec<u8> {
    let mut entries = parts(&package());
    let mut strip = |name: &str, pattern: &str| {
        let text = String::from_utf8(entries.remove(name).unwrap()).unwrap();
        assert!(text.contains(pattern), "{name} lacks {pattern}");
        entries.insert(name.into(), text.replace(pattern, "").into_bytes());
    };
    if !keep_formulas {
        strip("xl/worksheets/sheet1.xml", "<f>A2+Other!A2</f>");
        strip("xl/worksheets/sheet2.xml", "<f>Data!A2+A2+'Data'!$A$2</f>");
    }
    strip(
        "xl/_rels/workbook.xml.rels",
        r#"<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>"#,
    );
    strip(
        "[Content_Types].xml",
        r#"<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>"#,
    );
    entries.remove("xl/calcChain.xml").unwrap();
    repack(entries)
}
fn xml(bytes: &[u8], name: &str) -> String {
    String::from_utf8(parts(bytes).remove(name).unwrap()).unwrap()
}
fn edit(cell: &str, value: EditValue) -> CellEdit {
    CellEdit {
        cell: cell.into(),
        value,
    }
}
fn book(session: &EditSession) -> NativeXlsxBook {
    NativeXlsxBook::open_bytes(session.to_bytes().unwrap()).unwrap()
}
fn cell(book: &NativeXlsxBook, sheet: &str, coordinate: &str) -> wolfxl_reader::Cell {
    book.worksheet(sheet)
        .unwrap()
        .cells
        .into_iter()
        .find(|c| c.coordinate == coordinate)
        .unwrap()
}
fn preserved(before: &[u8], after: &[u8], expected: &[&str]) {
    let before = parts(before);
    let after = parts(after);
    for (name, bytes) in before {
        if !expected.contains(&name.as_str()) {
            assert_eq!(
                Sha256::digest(&bytes),
                Sha256::digest(after.get(&name).unwrap()),
                "unexpected part change: {name}"
            );
        }
    }
}

#[test]
fn noop_and_deterministic_atomic_save() {
    let original = package();
    let mut session = EditSession::from_bytes(original.clone()).unwrap();
    assert!(!session.is_dirty());
    assert_eq!(session.to_bytes().unwrap(), original);
    session
        .write_cells("Data", &[edit("A1", EditValue::Number(42.5))])
        .unwrap();
    let bytes = session.to_bytes().unwrap();
    assert_eq!(bytes, session.to_bytes().unwrap());
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("output.xlsm");
    std::fs::write(&path, b"old").unwrap();
    session.save(&path).unwrap();
    assert_eq!(std::fs::read(&path).unwrap(), bytes);
    assert_eq!(
        EditSession::open(&path).unwrap().sheet_names(),
        ["Data", "Other"]
    );
}

#[test]
fn values_clear_formula_and_part_hashes() {
    let original = package();
    let mut session = EditSession::from_bytes(original.clone()).unwrap();
    session
        .write_cells(
            "Data",
            &[
                edit("A1", EditValue::Clear),
                edit("B1", EditValue::Text("New & <text>".into())),
                edit("D5", EditValue::Bool(true)),
                edit("E5", EditValue::Number(3.5)),
                edit("F5", EditValue::Formula("=SUM(A2:E5)".into())),
            ],
        )
        .unwrap();
    assert!(session.is_dirty());
    let reader = book(&session);
    assert_eq!(cell(&reader, "Data", "A1").style_id, Some(1));
    assert_eq!(cell(&reader, "Data", "A1").value, CellValue::Empty);
    assert_eq!(
        cell(&reader, "Data", "B1").value.as_text(),
        Some("New & <text>")
    );
    assert_eq!(cell(&reader, "Data", "D5").value, CellValue::Bool(true));
    assert_eq!(cell(&reader, "Data", "E5").value, CellValue::Number(3.5));
    assert_eq!(
        cell(&reader, "Data", "F5").formula.as_deref(),
        Some("SUM(A2:E5)")
    );
    assert_eq!(cell(&reader, "Data", "F5").value, CellValue::Empty);
    let bytes = session.to_bytes().unwrap();
    assert!(!parts(&bytes).contains_key("xl/calcChain.xml"));
    assert!(xml(&bytes, "xl/workbook.xml").contains("fullCalcOnLoad=\"1\""));
    assert!(!xml(&bytes, "xl/_rels/workbook.xml.rels").contains("calcChain"));
    assert!(!xml(&bytes, "[Content_Types].xml").contains("calcChain"));
    preserved(
        &original,
        &bytes,
        &[
            "xl/worksheets/sheet1.xml",
            "xl/workbook.xml",
            "xl/_rels/workbook.xml.rels",
            "[Content_Types].xml",
            "xl/calcChain.xml",
        ],
    );
}

#[test]
fn scalar_writes_request_recalculation_when_formulas_exist() {
    // With and without a calculation chain: Data!B2 caches A2+Other!A2, so
    // value edits and clears must not leave that cached 13 standing.
    for original in [package(), package_without_calc_chain(true)] {
        let mut session = EditSession::from_bytes(original.clone()).unwrap();
        session
            .write_cells(
                "Data",
                &[
                    edit("A2", EditValue::Number(42.0)),
                    edit("C4", EditValue::Clear),
                ],
            )
            .unwrap();
        let bytes = session.to_bytes().unwrap();
        assert!(xml(&bytes, "xl/workbook.xml").contains("fullCalcOnLoad=\"1\""));
        assert!(!parts(&bytes).contains_key("xl/calcChain.xml"));
        preserved(
            &original,
            &bytes,
            &[
                "xl/worksheets/sheet1.xml",
                "xl/workbook.xml",
                "xl/_rels/workbook.xml.rels",
                "[Content_Types].xml",
                "xl/calcChain.xml",
            ],
        );
        assert_eq!(
            cell(&book(&session), "Data", "A2").value,
            CellValue::Number(42.0)
        );
    }
}

#[test]
fn scalar_writes_without_formulas_only_touch_one_part() {
    let original = package_without_calc_chain(false);
    let mut session = EditSession::from_bytes(original.clone()).unwrap();
    session
        .write_cells(
            "Data",
            &[
                edit("A1", EditValue::Number(42.0)),
                edit("C4", EditValue::Clear),
            ],
        )
        .unwrap();
    preserved(
        &original,
        &session.to_bytes().unwrap(),
        &["xl/worksheets/sheet1.xml"],
    );
    assert_eq!(
        cell(&book(&session), "Data", "A1").value,
        CellValue::Number(42.0)
    );
}

#[test]
fn rows_shift_formulas_merges_dimensions_validations_and_defined_names() {
    let mut session = EditSession::from_bytes(package()).unwrap();
    session.insert_rows("Data", 2, 2).unwrap();
    let reader = book(&session);
    let data = reader.worksheet("Data").unwrap();
    assert_eq!(data.dimension.as_deref(), Some("A1:C6"));
    assert_eq!(data.merged_ranges, ["B5:C5"]);
    assert_eq!(cell(&reader, "Data", "A4").value, CellValue::Number(11.0));
    assert_eq!(
        cell(&reader, "Data", "B4").formula.as_deref(),
        Some("A4+Other!A2")
    );
    assert_eq!(
        cell(&reader, "Other", "A1").formula.as_deref(),
        Some("Data!A4+A2+'Data'!$A$4")
    );
    let bytes = session.to_bytes().unwrap();
    let wb = xml(&bytes, "xl/workbook.xml");
    assert!(wb.contains("Data!$A$4"));
    assert!(wb.contains("name=\"Local\" localSheetId=\"0\">$A$4"));
    assert!(wb.contains("name=\"OtherLocal\" localSheetId=\"1\">$A$2"));
    assert!(xml(&bytes, "xl/worksheets/sheet1.xml").contains("sqref=\"A4:A6\""));
    session.delete_rows("Data", 2, 2).unwrap();
    assert_eq!(
        cell(&book(&session), "Data", "B2").formula.as_deref(),
        Some("A2+Other!A2")
    );
    assert_eq!(
        book(&session).worksheet("Data").unwrap().merged_ranges,
        ["B3:C3"]
    );
}

#[test]
fn deleting_referenced_rows_produces_ref_errors() {
    let mut session = EditSession::from_bytes(package()).unwrap();
    session.delete_rows("Data", 2, 1).unwrap();
    assert_eq!(
        cell(&book(&session), "Other", "A1").formula.as_deref(),
        Some("#REF!+A2+#REF!")
    );
}

#[test]
fn columns_shift_and_restore() {
    let mut session = EditSession::from_bytes(package()).unwrap();
    session.insert_columns("Data", 1, 2).unwrap();
    let reader = book(&session);
    assert_eq!(cell(&reader, "Data", "C2").value, CellValue::Number(11.0));
    assert_eq!(
        cell(&reader, "Data", "D2").formula.as_deref(),
        Some("C2+Other!A2")
    );
    assert_eq!(reader.worksheet("Data").unwrap().merged_ranges, ["D3:E3"]);
    session.delete_columns("Data", 1, 2).unwrap();
    assert_eq!(
        cell(&book(&session), "Data", "A2").value,
        CellValue::Number(11.0)
    );
}

#[test]
fn add_sheet_at_position_and_rename_quoted_references() {
    let original = package();
    let mut session = EditSession::from_bytes(original.clone()).unwrap();
    session.add_sheet("New Sheet", Some(1)).unwrap();
    assert_eq!(session.sheet_names(), ["Data", "New Sheet", "Other"]);
    assert_eq!(book(&session).sheet_names(), ["Data", "New Sheet", "Other"]);
    assert!(book(&session)
        .worksheet("New Sheet")
        .unwrap()
        .cells
        .is_empty());
    let added = session.to_bytes().unwrap();
    preserved(
        &original,
        &added,
        &[
            "xl/workbook.xml",
            "xl/_rels/workbook.xml.rels",
            "[Content_Types].xml",
            "docProps/app.xml",
        ],
    );
    assert!(xml(&added, "docProps/app.xml").contains("New Sheet"));
    session.rename_sheet("Data", "O'Brien Data").unwrap();
    let reader = book(&session);
    assert_eq!(reader.sheet_names(), ["O'Brien Data", "New Sheet", "Other"]);
    assert_eq!(
        cell(&reader, "Other", "A1").formula.as_deref(),
        Some("'O''Brien Data'!A2+A2+'O''Brien Data'!$A$2")
    );
    let renamed = session.to_bytes().unwrap();
    assert!(xml(&renamed, "xl/workbook.xml")
        .contains(quick_xml::escape::escape("'O''Brien Data'!$A$2").as_ref()));
    assert!(xml(&renamed, "docProps/app.xml")
        .contains(quick_xml::escape::escape("O'Brien Data").as_ref()));
}

#[test]
fn rename_sheet_rewrites_sheet_qualified_defined_names() {
    let mut entries = parts(&package());
    entries.insert(
        "xl/worksheets/sheet2.xml".into(),
        br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><f>Data!Local+data!Local+Other!OtherLocal+Amount</f><v>22</v></c></row></sheetData></worksheet>"#.to_vec(),
    );
    let mut session = EditSession::from_bytes(repack(entries)).unwrap();
    session.rename_sheet("Data", "New Data").unwrap();
    assert_eq!(
        cell(&book(&session), "Other", "A1").formula.as_deref(),
        Some("'New Data'!Local+'New Data'!Local+Other!OtherLocal+Amount")
    );
}

#[test]
fn number_format_preserves_styles_values_and_deduplicates() {
    let original = package();
    let mut session = EditSession::from_bytes(original.clone()).unwrap();
    session
        .set_number_format("Data", "A1:B2", "$#,##0.00")
        .unwrap();
    let first = session.to_bytes().unwrap();
    let styles = xml(&first, "xl/styles.xml");
    assert_eq!(styles.matches("formatCode=").count(), 1);
    let reader = book(&session);
    let a1 = cell(&reader, "Data", "A1");
    let a2 = cell(&reader, "Data", "A2");
    assert_ne!(a1.style_id, a2.style_id);
    assert_eq!(a1.value, CellValue::Number(7.0));
    assert_eq!(
        cell(&reader, "Data", "B2").formula.as_deref(),
        Some("A2+Other!A2")
    );
    session
        .set_number_format("Data", "A1:B2", "$#,##0.00")
        .unwrap();
    assert_eq!(xml(&session.to_bytes().unwrap(), "xl/styles.xml"), styles);
    preserved(
        &original,
        &first,
        &["xl/worksheets/sheet1.xml", "xl/styles.xml"],
    );
}

#[test]
fn invalid_edits_leave_session_unchanged() {
    let original = package();
    let mut session = EditSession::from_bytes(original.clone()).unwrap();
    assert!(matches!(
        session.write_cells("Missing", &[]),
        Err(EditError::UnknownSheet(_))
    ));
    for reference in ["A0", "XFE1", "A1048577", "Data!A1", "A1:B2", "", "$A$1"] {
        assert!(matches!(
            session.write_cells("Data", &[edit(reference, EditValue::Number(1.0))]),
            Err(EditError::BadReference(_))
        ));
    }
    for name in [
        "data",
        "Other",
        "",
        "A/B",
        "[bad]",
        "'start",
        "end'",
        "History",
        "12345678901234567890123456789012",
    ] {
        assert!(session.add_sheet(name, None).is_err(), "accepted {name:?}");
    }
    assert!(session.rename_sheet("Data", "OTHER").is_err());
    assert!(session.set_number_format("Data", "B2:A1", "0").is_err());
    assert!(session.insert_rows("Data", 0, 1).is_err());
    assert!(session.delete_columns("Data", 16384, 2).is_err());
    assert!(session
        .write_cells("Data", &[edit("A1", EditValue::Number(f64::NAN))])
        .is_err());
    assert!(session
        .write_cells("Data", &[edit("A1", EditValue::Formula("=".into()))])
        .is_err());
    assert!(!session.is_dirty());
    assert_eq!(session.to_bytes().unwrap(), original);
}

#[test]
fn zero_shift_is_noop_and_duplicate_cell_last_value_wins() {
    let original = package();
    let mut session = EditSession::from_bytes(original.clone()).unwrap();
    session.insert_rows("Data", 1, 0).unwrap();
    session.delete_columns("Data", 1, 0).unwrap();
    assert_eq!(session.to_bytes().unwrap(), original);
    session
        .write_cells(
            "Data",
            &[
                edit("A1", EditValue::Number(1.0)),
                edit("A1", EditValue::Number(2.0)),
            ],
        )
        .unwrap();
    assert_eq!(
        cell(&book(&session), "Data", "A1").value,
        CellValue::Number(2.0)
    );
}

#[test]
fn chart_series_follow_structural_shifts_and_sheet_rename() {
    let original = include_bytes!("../../../tests/fixtures/charts/bar_chart_simple.xlsx");
    let mut session = EditSession::from_bytes(original.to_vec()).unwrap();
    session.insert_rows("Data", 2, 2).unwrap();
    let chart = xml(&session.to_bytes().unwrap(), "xl/charts/chart1.xml");
    assert!(
        chart.contains(quick_xml::escape::escape("'Data'!$A$4:$A$8").as_ref()),
        "{chart}"
    );
    assert!(
        chart.contains(quick_xml::escape::escape("'Data'!$B$4:$B$8").as_ref()),
        "{chart}"
    );
    session.rename_sheet("Data", "New Data").unwrap();
    let chart = xml(&session.to_bytes().unwrap(), "xl/charts/chart1.xml");
    assert!(
        chart.contains(quick_xml::escape::escape("'New Data'!$A$4:$A$8").as_ref()),
        "{chart}"
    );
    assert!(
        chart.contains(quick_xml::escape::escape("'New Data'!$B$4:$B$8").as_ref()),
        "{chart}"
    );
    assert_eq!(book(&session).sheet_names(), ["New Data"]);
}

#[test]
fn table_formulas_shift_with_sheet_scope_without_touching_foreign_references() {
    let mut entries = parts(&package());
    entries.insert("xl/worksheets/_rels/sheet1.xml.rels".into(), br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="table" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>"#.to_vec());
    entries.insert("xl/tables/table1.xml".into(), br#"<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T" displayName="T" ref="A1:B4"><autoFilter ref="A1:B4"/><tableColumns count="2"><tableColumn id="1" name="X"><calculatedColumnFormula>A2+Other!A2</calculatedColumnFormula></tableColumn><tableColumn id="2" name="Y"><totalsRowFormula>SUM(A2:A4)+Other!A2</totalsRowFormula></tableColumn></tableColumns></table>"#.to_vec());
    let mut session = EditSession::from_bytes(repack(entries)).unwrap();
    session.insert_rows("Data", 2, 1).unwrap();
    let table = xml(&session.to_bytes().unwrap(), "xl/tables/table1.xml");
    assert!(table.contains("ref=\"A1:B5\""));
    assert!(table.contains("<calculatedColumnFormula>A3+Other!A2</calculatedColumnFormula>"));
    assert!(table.contains("<totalsRowFormula>SUM(A3:A5)+Other!A2</totalsRowFormula>"));
    session.delete_rows("Data", 2, 1).unwrap();
    let table = xml(&session.to_bytes().unwrap(), "xl/tables/table1.xml");
    assert!(table.contains("<calculatedColumnFormula>A2+Other!A2</calculatedColumnFormula>"));
    assert!(!table.contains("_wolfxl_preserved"));
}
