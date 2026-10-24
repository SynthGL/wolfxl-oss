// SPDX-License-Identifier: MIT
// All XML fixtures below were independently authored for this reader directly
// from ODF 1.3; no parser output, calamine source, or proprietary fixture donor.
use std::io::{Cursor, Write};
use std::sync::atomic::{AtomicU64, Ordering};

use wolfxl_reader::{CellDataType, CellValue, NativeOdsBook, ReaderError, SheetState};
use zip::write::SimpleFileOptions;

const NS: &str = r#"xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:x="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:f="urn:oasis:names:tc:opendocument:xmlns:of:1.2""#;

fn document(body: &str) -> String {
    format!(
        r#"<o:document-content {NS}><o:body><o:spreadsheet>{body}</o:spreadsheet></o:body></o:document-content>"#
    )
}

fn package(parts: &[(&str, &[u8])]) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    for (name, bytes) in parts {
        zip.start_file(*name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

fn book(body: &str) -> NativeOdsBook {
    let xml = document(body);
    NativeOdsBook::open_bytes(package(&[("content.xml", xml.as_bytes())])).unwrap()
}

#[test]
fn values_formula_caches_names_order_states_and_iso_values() {
    let reader = book(
        r#"
      <t:calculation-settings><t:null-date t:date-value="1904-01-01"/></t:calculation-settings>
      <t:named-expressions><t:named-range t:name="Global" t:cell-range-address="$'Main sheet'.$A$1:.$B$2" t:base-cell-address="$'Main sheet'.$A$1"/></t:named-expressions>
      <t:table t:name="Main sheet"><t:named-expressions><t:named-expression t:name="Local" t:expression="f:=2+3" t:base-cell-address="$'Main sheet'.$A$1"/></t:named-expressions>
        <t:table-row>
          <t:table-cell o:value-type="float" o:value="42.5"/>
          <t:table-cell o:value-type="boolean" o:boolean-value="true"/>
          <t:table-cell o:value-type="percentage" o:value="0.25"/>
          <t:table-cell o:value-type="currency" o:value="19.95" o:currency="EUR"/>
          <t:table-cell o:value-type="date" o:date-value="2024-02-29T23:59:59.125Z"/>
          <t:table-cell o:value-type="time" o:time-value="PT12H34M56.25S"/>
          <t:table-cell o:value-type="time" o:time-value="-P2DT3H"/>
          <t:table-cell o:value-type="float" o:value="45" t:formula="f:=SUM([.A1];2.5)"/>
          <t:table-cell o:value-type="string" o:string-value="saved &amp; exact" t:formula="f:=&quot;not cache&quot;"/>
          <t:table-cell o:value-type="boolean" o:boolean-value="false" t:formula="f:=1=2"/>
          <t:table-cell o:value-type="date" o:date-value="2024-02-29" t:formula="f:=DATE(2024;2;29)"/>
          <t:table-cell t:formula="f:=1+1"/>
        </t:table-row>
      </t:table><t:table t:name="Hidden" t:display="false"/><t:table t:name="Last"/>
    "#,
    );
    assert_eq!(reader.sheet_names(), vec!["Main sheet", "Hidden", "Last"]);
    assert!(reader.date1904());
    assert_eq!(reader.sheet_state("Hidden").unwrap(), SheetState::Hidden);
    assert_eq!(reader.sheet_state("Last").unwrap(), SheetState::Visible);
    assert!(matches!(
        reader.worksheet("absent"),
        Err(ReaderError::SheetNotFound(_))
    ));
    assert!(reader.number_format_for_style_id(0).is_none());
    let names = reader.named_ranges();
    assert_eq!(names.len(), 2);
    assert_eq!(
        (
            &*names[0].name,
            &*names[0].scope,
            names[0].local_sheet_id,
            &*names[0].refers_to
        ),
        ("Global", "workbook", None, "$'Main sheet'!$A$1:$B$2")
    );
    assert_eq!(
        (
            &*names[1].name,
            &*names[1].scope,
            names[1].local_sheet_id,
            &*names[1].refers_to
        ),
        ("Local", "sheet", Some(0), "2+3")
    );
    let sheet = reader.worksheet("Main sheet").unwrap();
    assert_eq!(sheet.dimension.as_deref(), Some("A1:L1"));
    let c = &sheet.cells;
    assert_eq!(c.len(), 12);
    assert_eq!(c[0].value, CellValue::Number(42.5));
    assert_eq!(c[1].value, CellValue::Bool(true));
    assert_eq!(c[2].value, CellValue::Number(0.25));
    assert_eq!(c[3].value, CellValue::Number(19.95));
    for (index, expected) in [
        (4, "2024-02-29T23:59:59.125Z"),
        (5, "PT12H34M56.25S"),
        (6, "-P2DT3H"),
        (10, "2024-02-29"),
    ] {
        assert_eq!(c[index].data_type, CellDataType::DateIso);
        assert_eq!(c[index].value, CellValue::String(expected.into()));
    }
    assert_eq!(c[7].formula.as_deref(), Some("SUM(A1,2.5)"));
    assert_eq!(c[7].value, CellValue::Number(45.0));
    assert_eq!(c[8].formula.as_deref(), Some("\"not cache\""));
    assert_eq!(c[8].value, CellValue::String("saved & exact".into()));
    assert_eq!(c[9].value, CellValue::Bool(false));
    assert_eq!(c[11].value, CellValue::Empty);
    assert_eq!(c[11].formula.as_deref(), Some("1+1"));
    assert_eq!(reader.worksheet("Last").unwrap().dimension, None);
}

#[test]
fn namespace_aliases_default_element_namespace_and_rebound_prefixes() {
    let original = document(
        r#"<t:table t:name="Alias"><t:table-row><t:table-cell o:value-type="string"><x:p>café 中文 &amp; exact</x:p></t:table-cell></t:table-row></t:table>"#,
    );
    let aliases = original
        .replace("xmlns:o=", "xmlns:office_alias=")
        .replace("<o:", "<office_alias:")
        .replace("</o:", "</office_alias:")
        .replace(" o:", " office_alias:")
        .replace("xmlns:t=", "xmlns:tab=")
        .replace("<t:", "<tab:")
        .replace("</t:", "</tab:")
        .replace(" t:", " tab:")
        .replace("xmlns:x=", "xmlns:txt=")
        .replace("<x:", "<txt:")
        .replace("</x:", "</txt:")
        .replace(" x:", " txt:");
    let a = NativeOdsBook::open_bytes(package(&[("content.xml", original.as_bytes())])).unwrap();
    let b = NativeOdsBook::open_bytes(package(&[("content.xml", aliases.as_bytes())])).unwrap();
    assert_eq!(a.worksheet("Alias").unwrap(), b.worksheet("Alias").unwrap());
    let default = r#"<document-content xmlns="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:v="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:a="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><body><spreadsheet><table xmlns="urn:oasis:names:tc:opendocument:xmlns:table:1.0" a:name="Default"><table-row><table-cell v:value-type="float" v:value="7"/></table-row></table></spreadsheet></body></document-content>"#;
    let c = NativeOdsBook::open_bytes(package(&[("content.xml", default.as_bytes())])).unwrap();
    assert_eq!(
        c.worksheet("Default").unwrap().cells[0].value,
        CellValue::Number(7.0)
    );
    let rebound = document(
        r#"<t:table t:name="Rebound"><t:table-row><t:table-cell xmlns:o="urn:not-office" o:value-type="float" o:value="7"/></t:table-row></t:table>"#,
    );
    let d = NativeOdsBook::open_bytes(package(&[("content.xml", rebound.as_bytes())])).unwrap();
    assert!(d.worksheet("Rebound").unwrap().cells.is_empty());
}

#[test]
fn sparse_empty_covered_and_large_repeats_do_not_materialize_empty_grid() {
    let b = book(
        r#"<t:table t:name="Sparse"><t:table-row t:number-rows-repeated="999999"><t:table-cell t:number-columns-repeated="4294967295"/></t:table-row><t:table-row><t:covered-table-cell t:number-columns-repeated="999999"/><t:table-cell o:value-type="float" o:value="7"/><t:table-cell t:number-columns-repeated="1000000"/></t:table-row></t:table>"#,
    );
    let sheet = b.worksheet("Sparse").unwrap();
    assert_eq!(sheet.cells.len(), 1);
    assert_eq!(
        (sheet.cells[0].row, sheet.cells[0].col),
        (1_000_000, 1_000_000)
    );
    assert_eq!(sheet.cells[0].coordinate, "BDWGN1000000");
    assert_eq!(
        sheet.dimension.as_deref(),
        Some("BDWGN1000000:BDWGN1000000")
    );
}

#[test]
fn repeated_values_and_formulas_preserve_identical_content_and_cache() {
    let b = book(
        r#"<t:table t:name="Repeat"><t:table-row t:number-rows-repeated="2"><t:table-cell/><t:table-cell o:value-type="float" o:value="9" t:formula="f:=[.A1]+[.$A$1]" t:number-columns-repeated="2"/></t:table-row></t:table>"#,
    );
    let s = b.worksheet("Repeat").unwrap();
    assert_eq!(s.dimension.as_deref(), Some("B1:C2"));
    let actual: Vec<_> = s
        .cells
        .iter()
        .map(|c| (c.coordinate.as_str(), c.formula.as_deref(), &c.value))
        .collect();
    assert_eq!(
        actual,
        vec![
            ("B1", Some("A1+$A$1"), &CellValue::Number(9.0)),
            ("C1", Some("A1+$A$1"), &CellValue::Number(9.0)),
            ("B2", Some("A1+$A$1"), &CellValue::Number(9.0)),
            ("C2", Some("A1+$A$1"), &CellValue::Number(9.0))
        ]
    );
}

#[test]
fn text_paragraphs_spaces_cdata_and_covered_merged_values() {
    let b = book(
        r#"<t:table t:name="Text"><t:table-row><t:table-cell o:value-type="string" t:number-columns-spanned="2"><x:p>  Hello <x:span> world </x:span><x:s x:c="2"/>x<x:tab/>y<x:line-break/>z</x:p><x:p><![CDATA[café & raw]]></x:p></t:table-cell><t:covered-table-cell/><t:table-cell o:value-type="string"><x:p/></t:table-cell></t:table-row></t:table>"#,
    );
    let s = b.worksheet("Text").unwrap();
    assert_eq!(s.merged_ranges, ["A1:B1"]);
    assert_eq!(s.cells.len(), 2);
    assert_eq!(
        s.cells[0].value,
        CellValue::String("Hello world   x\ty\nz\ncafé & raw".into())
    );
    assert_eq!(s.cells[1].value, CellValue::String(String::new()));
}

#[test]
fn table_styles_from_styles_and_automatic_styles_preserve_hidden_state() {
    let styles = format!(
        r#"<o:document-styles {NS}><o:styles><s:style s:name="base" s:family="table"><s:table-properties t:display="false"/></s:style><s:style s:name="child" s:family="table" s:parent-style-name="base"/></o:styles></o:document-styles>"#
    );
    let xml = format!(
        r#"<o:document-content {NS}><o:automatic-styles><s:style s:name="visible" s:family="table"><s:table-properties t:display="true"/></s:style></o:automatic-styles><o:body><o:spreadsheet><t:table t:name="Hidden style" t:style-name="child"/><t:table t:name="Visible" t:style-name="visible"/></o:spreadsheet></o:body></o:document-content>"#
    );
    let b = NativeOdsBook::open_bytes(package(&[
        ("content.xml", xml.as_bytes()),
        ("styles.xml", styles.as_bytes()),
    ]))
    .unwrap();
    assert_eq!(b.sheet_state("Hidden style").unwrap(), SheetState::Hidden);
    assert_eq!(b.sheet_state("Visible").unwrap(), SheetState::Visible);
}

#[test]
fn malformed_values_repeat_arithmetic_and_expansion_budgets_are_errors() {
    for cell in [
        r#"<t:table-cell t:number-columns-repeated="0"/>"#,
        r#"<t:table-cell t:number-columns-repeated="4294967296"/>"#,
        r#"<t:table-cell t:number-columns-repeated="4294967295"/><t:table-cell/>"#,
        r#"<t:table-cell o:value-type="float" o:value="not-a-number"/>"#,
        r#"<t:table-cell o:value-type="float"/>"#,
        r#"<t:table-cell o:value-type="boolean" o:boolean-value="maybe"/>"#,
        r#"<t:table-cell o:value-type="date" o:date-value="2023-02-29"/>"#,
        r#"<t:table-cell o:value-type="time" o:time-value="PT1.5H"/>"#,
        r#"<t:table-cell o:value-type="string"><x:p><x:s x:c="4294967295"/></x:p></t:table-cell>"#,
        r#"<t:table-cell o:value-type="float" o:value="1" t:number-columns-repeated="2000001"/>"#,
        r#"<t:table-cell o:value-type="string" o:string-value="0123456789012345678901234567890123456789" t:number-columns-repeated="2000000"/>"#,
        r#"<t:table-cell t:formula="f:=SUM([.A1)"/>"#,
    ] {
        let xml = document(&format!(
            r#"<t:table t:name="Bad"><t:table-row>{cell}</t:table-row></t:table>"#
        ));
        assert!(
            NativeOdsBook::open_bytes(package(&[("content.xml", xml.as_bytes())])).is_err(),
            "{cell}"
        );
    }
    let overflow = document(
        r#"<t:table t:name="Bad"><t:table-row t:number-rows-repeated="4294967295"/><t:table-row/></t:table>"#,
    );
    assert!(NativeOdsBook::open_bytes(package(&[("content.xml", overflow.as_bytes())])).is_err());
}

#[test]
fn unused_members_receive_zip_crc_and_xml_admission() {
    let xml = document(r#"<t:table t:name="Fine"/>"#);
    for (name, extra) in [
        ("../bad", b"x".as_slice()),
        (
            "unused.xml",
            b"<!DOCTYPE a [<!ENTITY e 'x'>]><a>&e;</a>".as_slice(),
        ),
        ("unused.xml", b"<a><b></a>".as_slice()),
        ("CONTENT.XML", b"<a/>".as_slice()),
    ] {
        assert!(
            NativeOdsBook::open_bytes(package(&[("content.xml", xml.as_bytes()), (name, extra)]))
                .is_err(),
            "{name}"
        );
    }
    let deep = format!("{}{}", "<a>".repeat(129), "</a>".repeat(129));
    assert!(NativeOdsBook::open_bytes(package(&[
        ("content.xml", xml.as_bytes()),
        ("unused.xml", deep.as_bytes())
    ]))
    .is_err());
    let mut damaged = package(&[
        ("content.xml", xml.as_bytes()),
        ("unused.bin", b"UNIQUE_BODY"),
    ]);
    let position = damaged
        .windows(11)
        .position(|v| v == b"UNIQUE_BODY")
        .unwrap();
    damaged[position] ^= 1;
    assert!(NativeOdsBook::open_bytes(damaged).is_err());
    let mut declared = package(&[
        ("content.xml", xml.as_bytes()),
        ("unused.bin", b"UNIQUE_BODY"),
    ]);
    let central = declared
        .windows(4)
        .enumerate()
        .filter(|(_, v)| *v == b"PK\x01\x02")
        .last()
        .unwrap()
        .0;
    declared[central + 24..central + 28].copy_from_slice(&12u32.to_le_bytes());
    assert!(NativeOdsBook::open_bytes(declared).is_err());
}

#[test]
fn cached_errors_are_not_formula_text() {
    let b = book(
        r##"<t:table t:name="Errors"><t:table-row><t:table-cell xmlns:c="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" o:value-type="string" c:value-type="error" t:formula="f:=1/0"><x:p>#DIV/0!</x:p></t:table-cell></t:table-row></t:table>"##,
    );
    let s = b.worksheet("Errors").unwrap();
    assert_eq!(s.cells[0].value, CellValue::Error("#DIV/0!".into()));
    assert_eq!(s.cells[0].data_type, CellDataType::Error);
    assert_eq!(s.cells[0].formula.as_deref(), Some("1/0"));
}

#[test]
fn path_and_bytes_have_identical_observable_contracts() {
    static SERIAL: AtomicU64 = AtomicU64::new(0);
    let root = std::env::var_os("TMPDIR")
        .map(std::path::PathBuf::from)
        .unwrap_or_else(|| std::env::current_dir().unwrap());
    let path = root.join(format!(
        "wolfxl-native-ods-{}-{}.ods",
        std::process::id(),
        SERIAL.fetch_add(1, Ordering::Relaxed)
    ));
    let xml = document(
        r#"<t:table t:name="Path"><t:table-row><t:table-cell o:value-type="date" o:date-value="2024-02-29"/></t:table-row></t:table>"#,
    );
    let bytes = package(&[("content.xml", xml.as_bytes())]);
    std::fs::OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(&path)
        .unwrap()
        .write_all(&bytes)
        .unwrap();
    let path_book = NativeOdsBook::open_path(&path).unwrap();
    let bytes_book = NativeOdsBook::open_bytes(bytes).unwrap();
    std::fs::remove_file(&path).unwrap();
    assert_eq!(path_book.sheet_names(), bytes_book.sheet_names());
    assert_eq!(path_book.named_ranges(), bytes_book.named_ranges());
    assert_eq!(path_book.date1904(), bytes_book.date1904());
    assert_eq!(
        path_book.sheet_state("Path").unwrap(),
        bytes_book.sheet_state("Path").unwrap()
    );
    assert_eq!(
        path_book.worksheet("Path").unwrap(),
        bytes_book.worksheet("Path").unwrap()
    );
}

#[test]
fn xml_schema_double_special_values_are_preserved() {
    let b = book(
        r#"<t:table t:name="Special"><t:table-row><t:table-cell o:value-type="float" o:value="INF"/><t:table-cell o:value-type="float" o:value="-INF"/><t:table-cell o:value-type="float" o:value="NaN"/></t:table-row></t:table>"#,
    );
    let s = b.worksheet("Special").unwrap();
    assert!(matches!(s.cells[0].value, CellValue::Number(n) if n == f64::INFINITY));
    assert!(matches!(s.cells[1].value, CellValue::Number(n) if n == f64::NEG_INFINITY));
    assert!(matches!(s.cells[2].value, CellValue::Number(n) if n.is_nan()));
}

#[test]
fn relative_names_keep_sheet_flags_scope_and_base_origins() {
    let b = book(
        r#"<t:named-expressions>
      <t:named-range t:name="Absolute" t:cell-range-address="$'O''Brien'.$B$3" t:base-cell-address="$'O''Brien'.$B$3"/>
      <t:named-range t:name="Relative" t:cell-range-address="'O''Brien'.C4:.D$5" t:base-cell-address="$'O''Brien'.$B$3"/>
      <t:named-expression t:name="Constant" t:expression="f:=2+3"/>
      </t:named-expressions><t:table t:name="O'Brien"><t:named-expressions>
      <t:named-expression t:name="Relative" t:expression="f:=[.$A1]+[$'O''Brien'.B$2]" t:base-cell-address="$'O''Brien'.$B$3"/>
      </t:named-expressions></t:table>"#,
    );
    let names = b.named_ranges();
    assert_eq!(
        names.iter().map(|n| n.name.as_str()).collect::<Vec<_>>(),
        ["Absolute", "Relative", "Constant", "Relative"]
    );
    assert_eq!(names[0].refers_to, "$'O''Brien'!$B$3");
    assert_eq!(names[1].refers_to, "'O''Brien'!C4:D$5");
    assert_eq!(names[3].refers_to, "$A1+$'O''Brien'!B$2");
    assert_eq!(names[1].scope, "workbook");
    assert_eq!(names[1].local_sheet_id, None);
    assert_eq!(names[3].scope, "sheet");
    assert_eq!(names[3].local_sheet_id, Some(0));
    assert_eq!(b.named_range_base_cell_address(0), Some("'O''Brien'!$B$3"));
    assert_eq!(b.named_range_base_cell_address(1), Some("'O''Brien'!$B$3"));
    assert_eq!(b.named_range_base_cell_address(2), None);
    assert_eq!(b.named_range_base_cell_address(3), Some("'O''Brien'!$B$3"));
    assert_eq!(b.named_range_base_cell_address(4), None);
}

#[test]
fn annotations_ruby_annotations_and_nested_tables_do_not_pollute_values() {
    let b = book(
        r#"<t:table t:name="Outer"><t:table-row><t:table-cell o:value-type="string"><x:p>visible <x:ruby><x:ruby-base>base</x:ruby-base><x:ruby-text>not value</x:ruby-text></x:ruby></x:p><o:annotation><x:p>not value</x:p></o:annotation><t:table t:name="Nested"><t:table-row><t:table-cell><x:p>not value</x:p></t:table-cell></t:table-row></t:table><x:p>end</x:p></t:table-cell></t:table-row></t:table>"#,
    );
    assert_eq!(b.sheet_names(), ["Outer"]);
    assert_eq!(
        b.worksheet("Outer").unwrap().cells[0].value,
        CellValue::String("visible base\nend".into())
    );
}

#[test]
fn zip_declared_entry_limits_and_manifest_encryption_are_admitted_before_use() {
    let xml = document(r#"<t:table t:name="Fine"/>"#);
    let mut oversized = package(&[("content.xml", xml.as_bytes()), ("unused.bin", b"x")]);
    let central = oversized
        .windows(4)
        .enumerate()
        .filter(|(_, v)| *v == b"PK\x01\x02")
        .last()
        .unwrap()
        .0;
    oversized[central + 24..central + 28].copy_from_slice(&u32::MAX.to_le_bytes());
    assert!(NativeOdsBook::open_bytes(oversized).is_err());
    let manifest = br#"<m:manifest xmlns:m="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><m:file-entry><m:encryption-data/></m:file-entry></m:manifest>"#;
    assert!(matches!(
        NativeOdsBook::open_bytes(package(&[
            ("content.xml", xml.as_bytes()),
            ("META-INF/manifest.xml", manifest)
        ])),
        Err(ReaderError::Unsupported(_))
    ));
    let duplicate = br#"<a xmlns:x="urn:test" xmlns:y="urn:test" x:v="1" y:v="2"/>"#;
    assert!(NativeOdsBook::open_bytes(package(&[
        ("content.xml", xml.as_bytes()),
        ("unused.xml", duplicate)
    ]))
    .is_err());
}

#[test]
fn epoch_and_mimetype_semantics_do_not_convert_iso_to_serials() {
    for epoch in ["1899-12-30", "1900-01-01"] {
        let xml = document(&format!(
            r#"<t:calculation-settings><t:null-date t:date-value="{epoch}"/></t:calculation-settings><t:table t:name="Date"><t:table-row><t:table-cell o:value-type="date" o:date-value="2024-02-29"/></t:table-row></t:table>"#
        ));
        let b = NativeOdsBook::open_bytes(package(&[
            ("content.xml", xml.as_bytes()),
            (
                "mimetype",
                b"application/vnd.oasis.opendocument.spreadsheet",
            ),
        ]))
        .unwrap();
        assert!(!b.date1904());
        assert_eq!(
            b.worksheet("Date").unwrap().cells[0].value,
            CellValue::String("2024-02-29".into())
        );
    }
    let xml = document("");
    assert!(NativeOdsBook::open_bytes(package(&[
        ("content.xml", xml.as_bytes()),
        ("mimetype", b"application/vnd.oasis.opendocument.text")
    ]))
    .is_err());
}

#[test]
fn repeated_column_declarations_and_merged_repeat_boundaries_are_checked() {
    for body in [
        r#"<t:table t:name="Bad"><t:table-column t:number-columns-repeated="0"/></t:table>"#,
        r#"<t:table t:name="Bad"><t:table-column t:number-columns-repeated="4294967295"/><t:table-column/></t:table>"#,
        r#"<t:table t:name="Bad"><t:table-row><t:table-cell t:number-columns-repeated="2" t:number-columns-spanned="2"/></t:table-row></t:table>"#,
        r#"<t:table t:name="Bad"><t:table-row t:number-rows-repeated="2"><t:table-cell t:number-rows-spanned="2"/></t:table-row></t:table>"#,
    ] {
        let xml = document(body);
        assert!(NativeOdsBook::open_bytes(package(&[("content.xml", xml.as_bytes())])).is_err());
    }
}

#[test]
fn utf16_xml_is_namespace_equivalent_and_decoding_is_strict() {
    let xml = format!(
        r#"<?xml version="1.0" encoding="UTF-16"?>{}"#,
        document(
            r#"<t:table t:name="Unicode"><t:table-row><t:table-cell o:value-type="string" o:string-value="café 中文 🐺"/></t:table-row></t:table>"#
        )
    );
    for little in [true, false] {
        let mut bytes = if little {
            vec![0xff, 0xfe]
        } else {
            vec![0xfe, 0xff]
        };
        for unit in xml.encode_utf16() {
            bytes.extend_from_slice(&if little {
                unit.to_le_bytes()
            } else {
                unit.to_be_bytes()
            });
        }
        let b = NativeOdsBook::open_bytes(package(&[("content.xml", &bytes)])).unwrap();
        assert_eq!(
            b.worksheet("Unicode").unwrap().cells[0].value,
            CellValue::String("café 中文 🐺".into())
        );
    }
    for invalid in [&[0xff, 0xfe, 0x00][..], &[0xff, 0xfe, 0x00, 0xd8][..]] {
        assert!(NativeOdsBook::open_bytes(package(&[("content.xml", invalid)])).is_err());
    }
}
