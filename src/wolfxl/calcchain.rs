//! `calcchain` — rebuild `xl/calcChain.xml` from the post-mutation
//! sheet bytes (Sprint Θ Pod-C3).
//!
//! Excel writes a calc-chain hint that lists every formula cell in
//! the order they should be evaluated. WolfXL historically left the
//! source file's calcChain alone (modify-mode) or omitted it entirely
//! (write-mode); Excel transparently rebuilds it on next open, so the
//! end-user impact is a one-time recompute on the first open after
//! saving.
//!
//! This module exposes the pure scanning + emission helpers; the
//! patcher's flush phase calls them after every sheet mutation has
//! settled. The output is byte-deterministic for a given (workbook sheetId,
//! formula cells per sheet) tuple — important for the diff-test
//! infrastructure.
//!
//! # Format
//!
//! ```xml
//! <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
//! <calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
//!   <c r="A1" i="1"/>
//!   <c r="B2" i="1"/>
//!   <c r="C3" i="2"/>
//! </calcChain>
//! ```
//!
//! Where `r` is the cell A1 reference and `i` matches the workbook
//! `<sheet sheetId="N">` value from `xl/workbook.xml`.

use quick_xml::events::Event;
use quick_xml::Reader as XmlReader;
use quick_xml::Writer as XmlWriter;

/// Content type for `xl/calcChain.xml`.
pub const CT_CALC_CHAIN: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml";

/// Relationship type for the workbook → calcChain edge.
pub const REL_CALC_CHAIN: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain";

/// One calcChain entry: cell reference + workbook sheetId.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CalcChainEntry {
    /// A1 reference, e.g. `"A1"`, `"BC42"`. Verbatim copy of the
    /// `<c r="…">` attribute on the formula cell.
    pub cell_ref: String,
    /// Workbook `<sheet sheetId="N">` value.
    pub sheet_index: u32,
}

/// Scan a single sheet's XML and emit one [`CalcChainEntry`] for each
/// `<c>` element that has at least one `<f>` child.
///
/// Quick-XML based; tolerant of namespaces and arbitrary whitespace.
/// Skips cells that have no formula.
pub fn scan_sheet_for_formulas(sheet_xml: &[u8], sheet_index_1based: u32) -> Vec<CalcChainEntry> {
    if !may_contain_formula_tag(sheet_xml) {
        return Vec::new();
    }

    let mut reader = XmlReader::from_reader(sheet_xml);
    reader.config_mut().trim_text(false);

    let mut entries: Vec<CalcChainEntry> = Vec::new();
    let mut buf: Vec<u8> = Vec::new();

    // The <c> we're currently inside, if any. We commit it when we
    // see an <f> child before the closing </c>.
    let mut current_c_ref: Option<String> = None;
    let mut emitted_for_current: bool = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name();
                let local = name.local_name();
                match local.as_ref() {
                    "c" => {
                        // Pull r="..." attribute.
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == "r" {
                                if let Ok(v) = attr.normalized_value(quick_xml::XmlVersion::Implicit1_0) {
                                    current_c_ref = Some(v.into_owned());
                                }
                            }
                        }
                        emitted_for_current = false;
                    }
                    "f" => {
                        if let Some(r) = current_c_ref.as_ref() {
                            if !emitted_for_current {
                                entries.push(CalcChainEntry {
                                    cell_ref: r.clone(),
                                    sheet_index: sheet_index_1based,
                                });
                                emitted_for_current = true;
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(e)) => {
                // Self-closing tags. <c r="A1" t="str"/> with no <f>
                // is a value cell, not a formula — skip. <f/> by
                // itself is unusual but we treat it the same as Start
                // for safety.
                let name = e.name();
                let local = name.local_name();
                match local.as_ref() {
                    "c" => {
                        // Self-closing <c .../> has no children → no formula.
                        current_c_ref = None;
                        emitted_for_current = false;
                    }
                    "f" => {
                        if let Some(r) = current_c_ref.as_ref() {
                            if !emitted_for_current {
                                entries.push(CalcChainEntry {
                                    cell_ref: r.clone(),
                                    sheet_index: sheet_index_1based,
                                });
                                emitted_for_current = true;
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::End(e)) => {
                if e.name().local_name().as_ref() == "c" {
                    current_c_ref = None;
                    emitted_for_current = false;
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break, // best-effort: a parse failure means we
            // skip this sheet rather than abort the
            // save.
            _ => {}
        }
        buf.clear();
    }
    entries
}

fn may_contain_formula_tag(sheet_xml: &[u8]) -> bool {
    sheet_xml
        .windows(2)
        .any(|pair| pair == b"<f" || pair == b":f")
}

/// Extract a source calcChain `<extLst>` block so rebuilds can preserve
/// extension metadata attached to the chain.
pub fn extract_ext_lst(calc_chain_xml: &[u8]) -> Option<Vec<u8>> {
    let mut reader = XmlReader::from_reader(calc_chain_xml);
    reader.config_mut().trim_text(false);
    let mut writer = XmlWriter::new(Vec::new());
    let mut buf: Vec<u8> = Vec::new();
    let mut in_ext_lst = false;
    let mut depth = 0usize;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.name().local_name().as_ref() == "extLst" && !in_ext_lst => {
                in_ext_lst = true;
                depth = 1;
                writer.write_event(Event::Start(e.into_owned())).ok()?;
            }
            Ok(Event::Empty(e)) if e.name().local_name().as_ref() == "extLst" && !in_ext_lst => {
                writer.write_event(Event::Empty(e.into_owned())).ok()?;
                return Some(writer.into_inner());
            }
            Ok(Event::Start(e)) if in_ext_lst => {
                depth += 1;
                writer.write_event(Event::Start(e.into_owned())).ok()?;
            }
            Ok(Event::Empty(e)) if in_ext_lst => {
                writer.write_event(Event::Empty(e.into_owned())).ok()?;
            }
            Ok(Event::Text(e)) if in_ext_lst => {
                writer.write_event(Event::Text(e.into_owned())).ok()?;
            }
            // References are separate events in quick-xml 0.42; copy them verbatim.
            Ok(Event::GeneralRef(e)) if in_ext_lst => {
                writer.write_event(Event::GeneralRef(e.into_owned())).ok()?;
            }
            Ok(Event::CData(e)) if in_ext_lst => {
                writer.write_event(Event::CData(e.into_owned())).ok()?;
            }
            Ok(Event::Comment(e)) if in_ext_lst => {
                writer.write_event(Event::Comment(e.into_owned())).ok()?;
            }
            Ok(Event::End(e)) if in_ext_lst => {
                writer.write_event(Event::End(e.into_owned())).ok()?;
                depth = depth.saturating_sub(1);
                if depth == 0 {
                    return Some(writer.into_inner());
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    None
}

/// Render `xl/calcChain.xml` from a flat list of entries. Order is
/// preserved verbatim — caller is responsible for the iteration order
/// (typically sheet tab-order, then sheet-XML scan order).
///
/// Returns `None` if there are no entries — the caller should DELETE
/// the calcChain.xml part instead of writing an empty one.
pub fn render_calc_chain_with_ext_lst(
    entries: &[CalcChainEntry],
    ext_lst: Option<&[u8]>,
) -> Option<Vec<u8>> {
    if entries.is_empty() {
        return None;
    }
    let mut out = String::new();
    out.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
    out.push_str("<calcChain xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
    for e in entries {
        // Escape `&`, `<`, `"` in the cell ref (defensive — Excel
        // never emits non-A1 refs but a corrupt source could).
        let escaped = e
            .cell_ref
            .replace('&', "&amp;")
            .replace('<', "&lt;")
            .replace('"', "&quot;");
        out.push_str(&format!("<c r=\"{}\" i=\"{}\"/>", escaped, e.sheet_index));
    }
    let mut bytes = out.into_bytes();
    if let Some(ext) = ext_lst {
        bytes.extend_from_slice(ext);
    }
    bytes.extend_from_slice(b"</calcChain>");
    Some(bytes)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn extract_ext_lst_preserves_entity_spelling() {
        let xml = br#"<calcChain><extLst><ext uri="A&amp;B"><value> A&amp;B&#32;&custom; </value></ext></extLst></calcChain>"#;
        let ext = extract_ext_lst(xml).unwrap();
        assert_eq!(ext, br#"<extLst><ext uri="A&amp;B"><value> A&amp;B&#32;&custom; </value></ext></extLst>"#);
    }

    #[test]
    fn scan_finds_formula_cells_only() {
        let xml = r#"<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1"><f>SUM(C1:D1)</f><v>10</v></c>
      <c r="C1"><v>4</v></c>
      <c r="D1"><f>A1+1</f><v>2</v></c>
    </row>
  </sheetData>
</worksheet>"#;
        let entries = scan_sheet_for_formulas(xml.as_bytes(), 1);
        assert_eq!(entries.len(), 2, "expected 2 formula cells");
        assert_eq!(entries[0].cell_ref, "B1");
        assert_eq!(entries[0].sheet_index, 1);
        assert_eq!(entries[1].cell_ref, "D1");
        assert_eq!(entries[1].sheet_index, 1);
    }

    #[test]
    fn scan_skips_self_closing_value_cells() {
        let xml = br#"<sheetData>
            <row r="1"><c r="A1" t="str"/></row>
        </sheetData>"#;
        assert_eq!(scan_sheet_for_formulas(xml, 7), vec![]);
    }

    #[test]
    fn scan_fast_skips_xml_without_formula_tags() {
        let xml = br#"<sheetData>
            <row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>
        </sheetData>"#;
        assert!(!may_contain_formula_tag(xml));
        assert_eq!(scan_sheet_for_formulas(xml, 7), vec![]);
    }

    #[test]
    fn scan_prefilter_keeps_prefixed_formula_tags() {
        let xml = br#"<sheetData xmlns:x="urn:test">
            <row r="1"><c r="A1"><x:f>A1+1</x:f><v>2</v></c></row>
        </sheetData>"#;
        assert!(may_contain_formula_tag(xml));
        let entries = scan_sheet_for_formulas(xml, 3);
        assert_eq!(entries.len(), 1);
        assert_eq!(entries[0].cell_ref, "A1");
        assert_eq!(entries[0].sheet_index, 3);
    }

    #[test]
    fn scan_emits_one_entry_per_cell_even_with_multiple_f() {
        // Defensive: normally a <c> contains at most one <f>, but if a
        // shared formula and a normal <f> co-exist, we shouldn't
        // double-count.
        let xml = br#"<sheetData>
            <row r="1"><c r="A1"><f t="shared" si="0"/><f>A1+1</f><v>0</v></c></row>
        </sheetData>"#;
        let entries = scan_sheet_for_formulas(xml, 1);
        assert_eq!(entries.len(), 1);
        assert_eq!(entries[0].cell_ref, "A1");
    }

    #[test]
    fn render_empty_returns_none() {
        assert!(render_calc_chain_with_ext_lst(&[], None).is_none());
    }

    #[test]
    fn render_emits_canonical_shape() {
        let entries = vec![
            CalcChainEntry {
                cell_ref: "A1".into(),
                sheet_index: 1,
            },
            CalcChainEntry {
                cell_ref: "B2".into(),
                sheet_index: 2,
            },
        ];
        let bytes = render_calc_chain_with_ext_lst(&entries, None).expect("non-empty");
        let s = String::from_utf8(bytes).unwrap();
        assert!(s.contains("<?xml"));
        assert!(s.contains("<calcChain"));
        assert!(s.contains("<c r=\"A1\" i=\"1\"/>"));
        assert!(s.contains("<c r=\"B2\" i=\"2\"/>"));
        assert!(s.contains("</calcChain>"));
    }

    #[test]
    fn extract_ext_lst_preserves_source_extension_block() {
        let xml = br#"<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="A1" i="1"/><extLst><ext uri="{wolfxl-test}"><x:test xmlns:x="urn:test">ok</x:test></ext></extLst></calcChain>"#;
        let ext = extract_ext_lst(xml).expect("extLst present");
        let s = String::from_utf8(ext).unwrap();
        assert!(s.starts_with("<extLst>"), "{s}");
        assert!(s.contains("{wolfxl-test}"), "{s}");
        assert!(s.contains("urn:test"), "{s}");
    }

    #[test]
    fn render_can_append_preserved_ext_lst() {
        let entries = vec![CalcChainEntry {
            cell_ref: "A1".into(),
            sheet_index: 1,
        }];
        let bytes =
            render_calc_chain_with_ext_lst(&entries, Some(b"<extLst><ext uri=\"u\"/></extLst>"))
                .expect("non-empty");
        let s = String::from_utf8(bytes).unwrap();
        assert!(s.contains("<c r=\"A1\" i=\"1\"/>"));
        assert!(
            s.contains("<extLst><ext uri=\"u\"/></extLst></calcChain>"),
            "{s}"
        );
    }
}
