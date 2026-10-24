//! Wave 2 integration gate: end-to-end xlsx synthesis + re-parse.
//!
//! Builds a small `Workbook` with every interesting cell type, a merge,
//! a freeze pane, a defined name, and a styled cell — runs every Wave 1+2
//! emitter + the ZIP packager, opens the resulting bytes with
//! `zip::ZipArchive`, parses every extracted XML part with
//! `quick_xml::Reader`, and asserts cell values round-trip through the SST.
//!
//! If this test fails, Wave 2 is not complete. It is deliberately verbose
//! — each assertion doubles as documentation of what one full emit cycle
//! is supposed to produce.

use std::collections::HashMap;
use std::io::{Cursor, Read};

use quick_xml::events::Event;
use quick_xml::Reader;

use wolfxl_writer::emit_xlsx;
use wolfxl_writer::model::cell::{FormulaResult, WriteCell, WriteCellValue};
use wolfxl_writer::model::comment::Comment;
use wolfxl_writer::model::conditional::{
    CellIsOperator, ConditionalFormat, ConditionalKind, ConditionalRule,
};
use wolfxl_writer::model::defined_name::DefinedName;
use wolfxl_writer::model::format::{DxfRecord, FontSpec, FormatSpec};
use wolfxl_writer::model::table::{Table, TableColumn};
use wolfxl_writer::model::validation::{
    DataValidation, ErrorStyle, ValidationOperator, ValidationType,
};
use wolfxl_writer::model::workbook::Workbook;
use wolfxl_writer::model::worksheet::{FreezePane, Merge, Worksheet};

/// Build the fixture: two sheets, mixed cells, merge, freeze, one defined
/// name, one styled cell. Returns the assembled `Workbook` *plus* the xf
/// index of the bold-red style so the caller can cross-check the `s` attr.
fn build_fixture() -> (Workbook, u32) {
    let mut wb = Workbook::new();

    // Intern a red-bold font into the styles table so we can point a cell at it.
    let bold_red = FormatSpec {
        font: Some(FontSpec {
            bold: true,
            color_rgb: Some("FFFF0000".to_string()),
            ..Default::default()
        }),
        ..Default::default()
    };
    let style_id = wb.styles.intern_format(&bold_red);

    // Sheet 1: every cell type.
    let mut s1 = Worksheet::new("Data");
    s1.set_cell(
        1,
        1,
        WriteCell::new(WriteCellValue::String("Name".to_string())),
    );
    s1.set_cell(
        1,
        2,
        WriteCell::new(WriteCellValue::String("Count".to_string())),
    );
    s1.set_cell(
        2,
        1,
        WriteCell::new(WriteCellValue::String("apples".to_string())),
    );
    s1.set_cell(2, 2, WriteCell::new(WriteCellValue::Number(42.0)));
    s1.set_cell(
        3,
        1,
        WriteCell::new(WriteCellValue::String("pears".to_string())),
    );
    s1.set_cell(3, 2, WriteCell::new(WriteCellValue::Number(3.5)));
    s1.set_cell(4, 1, WriteCell::new(WriteCellValue::Boolean(true)));
    s1.set_cell(
        4,
        2,
        WriteCell::new(WriteCellValue::Formula {
            expr: "SUM(B2:B3)".to_string(),
            result: Some(FormulaResult::Number(45.5)),
        }),
    );
    // Styled cell — points into the xf table.
    s1.set_cell(
        5,
        1,
        WriteCell::new(WriteCellValue::String("styled".to_string())).with_style(style_id),
    );
    // A genuinely blank styled cell to prove blank+style still emits.
    s1.set_cell(
        5,
        2,
        WriteCell::new(WriteCellValue::Blank).with_style(style_id),
    );

    s1.merge(Merge {
        top_row: 1,
        left_col: 1,
        bottom_row: 1,
        right_col: 2,
    });
    // Freeze cell at A2 = freeze the first row (1 row frozen, ySplit=1).
    s1.freeze = Some(FreezePane {
        freeze_row: 2,
        freeze_col: 0,
        top_left: None,
    });
    wb.add_sheet(s1);

    // Sheet 2: one string shared with sheet 1 ("apples") to exercise SST dedup.
    let mut s2 = Worksheet::new("Summary");
    s2.set_cell(
        1,
        1,
        WriteCell::new(WriteCellValue::String("apples".to_string())),
    );
    s2.set_cell(1, 2, WriteCell::new(WriteCellValue::Number(100.0)));
    wb.add_sheet(s2);

    // Workbook-scope defined name.
    wb.defined_names.push(DefinedName {
        name: "Grand_Total".to_string(),
        formula: "Summary!$B$1".to_string(),
        scope_sheet_index: None,
        builtin: None,
        hidden: false,
        ..Default::default()
    });

    (wb, style_id)
}

// `emit_full_pipeline` was promoted to `wolfxl_writer::emit_xlsx` in W4A's
// contract sub-commit so the `NativeWorkbook` pyclass can call it directly.
// The integration test now exercises the same code path Python users hit.

/// Open an xlsx byte slice and return a map of path → bytes for every entry.
fn read_archive(bytes: &[u8]) -> HashMap<String, Vec<u8>> {
    let cursor = Cursor::new(bytes);
    let mut archive = zip::ZipArchive::new(cursor).expect("open zip");
    let mut out = HashMap::new();
    for i in 0..archive.len() {
        let mut file = archive.by_index(i).expect("zip entry");
        let path = file.name().expect("valid fixture entry name").into_owned();
        let mut buf = Vec::new();
        file.read_to_end(&mut buf).expect("read entry");
        out.insert(path, buf);
    }
    out
}

/// Parse raw bytes with quick_xml end-to-end, panicking on any XML-level error.
fn assert_xml_well_formed(part: &str, bytes: &[u8]) {
    let text = std::str::from_utf8(bytes).unwrap_or_else(|e| panic!("{part} not utf8: {e}"));
    let mut reader = Reader::from_str(text);
    reader.config_mut().check_end_names = true;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Eof) => break,
            Err(e) => panic!("{part} parse error: {e}"),
            _ => (),
        }
        buf.clear();
    }
}

/// Walk sharedStrings.xml and return the interned strings in index order.
fn parse_sst(bytes: &[u8]) -> Vec<String> {
    let text = std::str::from_utf8(bytes).expect("sst utf8");
    let mut reader = Reader::from_str(text);
    reader.config_mut().check_end_names = true;
    let mut buf = Vec::new();
    let mut out = Vec::new();
    let mut in_t = false;
    let mut current = String::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.name().as_ref() == "t" => {
                in_t = true;
                current.clear();
            }
            Ok(Event::End(e)) if e.name().as_ref() == "t" => {
                in_t = false;
                out.push(std::mem::take(&mut current));
            }
            Ok(Event::Text(t)) if in_t => {
                current.push_str(&t);
            }
            Ok(Event::GeneralRef(r)) if in_t => {
                let decoded =
                    quick_xml::escape::unescape(&format!("&{};", &*r)).expect("unescape sst");
                current.push_str(&decoded);
            }
            Ok(Event::Eof) => break,
            Err(e) => panic!("sst parse error: {e}"),
            _ => (),
        }
        buf.clear();
    }
    out
}

#[test]
fn wave2_full_pipeline_roundtrip() {
    let (mut wb, style_id) = build_fixture();
    let bytes = emit_xlsx(&mut wb);

    // 1. Archive is openable and contains every required part.
    let parts = read_archive(&bytes);
    for required in [
        "[Content_Types].xml",
        "_rels/.rels",
        "xl/workbook.xml",
        "xl/_rels/workbook.xml.rels",
        "xl/worksheets/sheet1.xml",
        "xl/worksheets/sheet2.xml",
        "xl/styles.xml",
        "xl/sharedStrings.xml",
        "docProps/core.xml",
        "docProps/app.xml",
    ] {
        assert!(
            parts.contains_key(required),
            "archive missing required part: {required}"
        );
    }

    // 2. Every XML part is well-formed.
    for (path, data) in &parts {
        assert_xml_well_formed(path, data);
    }

    // 3. SST contains exactly the distinct strings we interned, in
    //    first-use order. "apples" appears on both sheets but only once
    //    in the table; the total reference count is higher than unique.
    let sst = parse_sst(&parts["xl/sharedStrings.xml"]);
    assert!(sst.contains(&"Name".to_string()), "SST missing Name");
    assert!(sst.contains(&"Count".to_string()), "SST missing Count");
    assert!(sst.contains(&"apples".to_string()), "SST missing apples");
    assert!(sst.contains(&"pears".to_string()), "SST missing pears");
    assert!(sst.contains(&"styled".to_string()), "SST missing styled");
    assert_eq!(
        sst.iter().filter(|s| *s == "apples").count(),
        1,
        "SST should dedupe 'apples' across sheets"
    );

    // 4. Sheet1 has the right structural hooks.
    let sheet1 = std::str::from_utf8(&parts["xl/worksheets/sheet1.xml"]).unwrap();
    assert!(sheet1.contains("<mergeCell ref=\"A1:B1\""), "{sheet1}");
    // Freeze pane present.
    assert!(
        sheet1.contains("<pane "),
        "sheet1 should emit <pane> for freeze: {sheet1}"
    );
    // Formula emitted with cached result.
    assert!(sheet1.contains("<f>SUM(B2:B3)</f>"), "{sheet1}");
    assert!(sheet1.contains("<v>45.5</v>"), "{sheet1}");
    // Boolean emitted as t="b" with <v>1</v>.
    assert!(sheet1.contains("t=\"b\""), "{sheet1}");
    // Styled cell points at the interned xf id.
    let s_attr = format!("s=\"{style_id}\"");
    assert!(
        sheet1.contains(&s_attr),
        "styled cell should have {s_attr}: {sheet1}"
    );

    // 5. Sheet2 re-uses the "apples" SST index. Parse the string-typed
    //    cell and verify the index maps back to "apples".
    let sheet2 = std::str::from_utf8(&parts["xl/worksheets/sheet2.xml"]).unwrap();
    let apples_idx = sst
        .iter()
        .position(|s| s == "apples")
        .expect("apples in SST");
    let expected = format!("t=\"s\"><v>{apples_idx}</v>");
    assert!(
        sheet2.contains(&expected),
        "sheet2 should reference SST index {apples_idx} for 'apples': {sheet2}"
    );

    // 6. Workbook lists the defined name and both sheets.
    let wb_xml = std::str::from_utf8(&parts["xl/workbook.xml"]).unwrap();
    assert!(wb_xml.contains("<sheet name=\"Data\""), "{wb_xml}");
    assert!(wb_xml.contains("<sheet name=\"Summary\""), "{wb_xml}");
    assert!(
        wb_xml.contains("<definedName name=\"Grand_Total\">Summary!$B$1</definedName>"),
        "{wb_xml}"
    );

    // 7. Styles part surfaces our bold-red font. The exact font index is
    //    implementation-detail, but a bold=1 + color FFFF0000 pair must
    //    appear in the <fonts> block.
    let styles = std::str::from_utf8(&parts["xl/styles.xml"]).unwrap();
    assert!(styles.contains("<b/>"), "bold tag missing: {styles}");
    assert!(
        styles.contains("FFFF0000") || styles.contains("ffff0000"),
        "red color missing from styles: {styles}"
    );
}

/// Wave 3 integration gate: build a workbook that exercises every rich
/// feature added in Wave 3 — multi-author comments (insertion-ordered),
/// a structured table with totals row, a conditional format pointing at
/// an interned dxf, and a list-type data validation. The xlsx is then
/// re-opened and every feature is asserted from the resulting parts.
fn build_wave3_fixture() -> (Workbook, u32) {
    let mut wb = Workbook::new();

    // Authors interned in deliberate order — Bob FIRST, Alice SECOND. This
    // is the test that proves the BTreeMap-ordering bug from rust_xlsxwriter
    // is gone: alphabetical sort would put Alice first.
    let bob_id = wb.comment_authors.intern("Bob");
    let alice_id = wb.comment_authors.intern("Alice");
    assert_eq!(bob_id, 0);
    assert_eq!(alice_id, 1);

    // Intern a bold-red dxf for the conditional format to point at.
    let bold_red_dxf = DxfRecord {
        font: Some(FontSpec {
            bold: true,
            color_rgb: Some("FFFF0000".to_string()),
            ..Default::default()
        }),
        ..Default::default()
    };
    let dxf_id = wb.styles.intern_dxf(&bold_red_dxf);

    let mut sheet = Worksheet::new("Data");
    sheet.set_cell(
        1,
        1,
        WriteCell::new(WriteCellValue::String("Alpha".to_string())),
    );
    sheet.set_cell(
        1,
        2,
        WriteCell::new(WriteCellValue::String("Beta".to_string())),
    );
    sheet.set_cell(
        1,
        3,
        WriteCell::new(WriteCellValue::String("Gamma".to_string())),
    );
    sheet.set_cell(2, 1, WriteCell::new(WriteCellValue::Number(50.0)));
    sheet.set_cell(2, 2, WriteCell::new(WriteCellValue::Number(150.0)));
    sheet.set_cell(2, 3, WriteCell::new(WriteCellValue::Number(75.0)));

    // Comments from each author on different cells. authorId values must
    // match the workbook-scope intern order (Bob=0, Alice=1).
    sheet.comments.insert(
        "A1".to_string(),
        Comment {
            text: "Bob's note".to_string(),
            author_id: bob_id,
            width_pt: None,
            height_pt: None,
            visible: false,
        },
    );
    sheet.comments.insert(
        "B1".to_string(),
        Comment {
            text: "Alice's note".to_string(),
            author_id: alice_id,
            width_pt: None,
            height_pt: None,
            visible: false,
        },
    );

    // Structured table with three columns and a totals row.
    sheet.tables.push(Table {
        name: "DataTable".to_string(),
        display_name: None,
        range: "A1:C3".to_string(),
        columns: vec![
            TableColumn {
                name: "Alpha".to_string(),
                totals_function: None,
                totals_label: None,
            },
            TableColumn {
                name: "Beta".to_string(),
                totals_function: Some("sum".to_string()),
                totals_label: None,
            },
            TableColumn {
                name: "Gamma".to_string(),
                totals_function: None,
                totals_label: None,
            },
        ],
        header_row: true,
        totals_row: true,
        style: None,
        autofilter: true,
    });

    // Conditional format: cells > 100 get bold-red dxf.
    sheet.conditional_formats.push(ConditionalFormat {
        sqref: "A2:C2".to_string(),
        rules: vec![ConditionalRule {
            kind: ConditionalKind::CellIs {
                operator: CellIsOperator::GreaterThan,
                formula_a: "100".to_string(),
                formula_b: None,
            },
            dxf_id: Some(dxf_id),
            stop_if_true: false,
            priority: None,
        }],
    });

    // Data validation: list of three colors.
    sheet.validations.push(DataValidation {
        sqref: "A4".to_string(),
        validation_type: ValidationType::List,
        operator: ValidationOperator::Between,
        formula_a: Some("\"Red,Green,Blue\"".to_string()),
        formula_b: None,
        allow_blank: true,
        show_dropdown: false,
        show_error_message: true,
        error_style: ErrorStyle::Stop,
        error_title: None,
        error_message: None,
        show_input_message: false,
        input_title: None,
        input_message: None,
    });

    wb.add_sheet(sheet);
    (wb, dxf_id)
}

#[test]
fn wave3_rich_features_roundtrip() {
    let (mut wb, _dxf_id) = build_wave3_fixture();
    let bytes = emit_xlsx(&mut wb);

    // 1. Archive opens and contains the Wave 3 parts.
    let parts = read_archive(&bytes);
    for required in [
        "xl/worksheets/sheet1.xml",
        "xl/comments/comments1.xml",
        "xl/drawings/vmlDrawing1.vml",
        "xl/tables/table1.xml",
        "xl/styles.xml",
    ] {
        assert!(
            parts.contains_key(required),
            "archive missing Wave 3 part: {required}; got keys={:?}",
            parts.keys().collect::<Vec<_>>()
        );
    }

    // 2. Every part is well-formed XML/VML.
    for (path, data) in &parts {
        assert_xml_well_formed(path, data);
    }

    // 3. comments1.xml emits authors in INSERTION order — Bob before Alice.
    //    This is the assertion that proves the BTreeMap bug is gone; an
    //    alphabetical sort would put Alice first.
    let comments = std::str::from_utf8(&parts["xl/comments/comments1.xml"]).unwrap();
    let bob_pos = comments
        .find("<author>Bob</author>")
        .expect("Bob in <authors>");
    let alice_pos = comments
        .find("<author>Alice</author>")
        .expect("Alice in <authors>");
    assert!(
        bob_pos < alice_pos,
        "Bob must precede Alice (insertion order): {comments}"
    );
    // authorId attributes match the intern order.
    assert!(
        comments.contains("<comment ref=\"A1\" authorId=\"0\">"),
        "Bob's comment authorId=0 (A1): {comments}"
    );
    assert!(
        comments.contains("<comment ref=\"B1\" authorId=\"1\">"),
        "Alice's comment authorId=1 (B1): {comments}"
    );

    // 4. vmlDrawing1.vml declares the three required namespaces.
    let vml = std::str::from_utf8(&parts["xl/drawings/vmlDrawing1.vml"]).unwrap();
    assert!(vml.contains("xmlns:v="), "missing v: namespace: {vml}");
    assert!(vml.contains("xmlns:o="), "missing o: namespace: {vml}");
    assert!(vml.contains("xmlns:x="), "missing x: namespace: {vml}");

    // 5. table1.xml has id=1 and three <tableColumn> entries.
    let table = std::str::from_utf8(&parts["xl/tables/table1.xml"]).unwrap();
    assert!(
        table.contains("<table xmlns=") && table.contains(" id=\"1\""),
        "table1 root id=1: {table}"
    );
    assert!(
        table.contains("<tableColumns count=\"3\">"),
        "three tableColumns: {table}"
    );
    assert!(
        table.contains("totalsRowFunction=\"sum\""),
        "totals row sum function: {table}"
    );

    // 6. sheet1.xml wires up <legacyDrawing> and <tableParts>.
    let sheet = std::str::from_utf8(&parts["xl/worksheets/sheet1.xml"]).unwrap();
    assert!(
        sheet.contains("<legacyDrawing r:id=\"rId2\"/>"),
        "legacyDrawing rId2: {sheet}"
    );
    assert!(
        sheet.contains("<tableParts count=\"1\"><tablePart r:id=\"rId3\"/></tableParts>"),
        "tableParts rId3 (after comments rId1+VML rId2): {sheet}"
    );

    // 7. sheet1.xml carries the conditional-format block with cellIs rule.
    assert!(
        sheet.contains("<conditionalFormatting sqref=\"A2:C2\">"),
        "CF wrapper present: {sheet}"
    );
    assert!(
        sheet.contains("type=\"cellIs\"") && sheet.contains("operator=\"greaterThan\""),
        "CF cellIs greaterThan: {sheet}"
    );

    // 8. sheet1.xml carries the data-validation block with the literal list.
    assert!(
        sheet.contains("<dataValidations count=\"1\">"),
        "DV wrapper count=1: {sheet}"
    );
    assert!(
        sheet.contains("<formula1>\"Red,Green,Blue\"</formula1>"),
        "DV list formula1: {sheet}"
    );

    // 9. styles.xml has a non-empty <dxfs> with our bold-red font.
    let styles = std::str::from_utf8(&parts["xl/styles.xml"]).unwrap();
    let dxfs_start = styles
        .find("<dxfs ")
        .expect("dxfs element present in styles");
    let dxfs_end = styles[dxfs_start..]
        .find("</dxfs>")
        .expect("dxfs close tag");
    let dxfs_block = &styles[dxfs_start..dxfs_start + dxfs_end];
    assert!(
        !dxfs_block.contains("count=\"0\""),
        "dxfs must not be empty when CF references one: {dxfs_block}"
    );
    assert!(
        dxfs_block.contains("<dxf>")
            && dxfs_block.contains("<font>")
            && dxfs_block.contains("<b/>"),
        "dxfs block must contain bold-font dxf: {dxfs_block}"
    );
    assert!(
        dxfs_block.contains("FFFF0000") || dxfs_block.contains("ffff0000"),
        "dxfs block must contain red color: {dxfs_block}"
    );
}
