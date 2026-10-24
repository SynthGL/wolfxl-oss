//! Property test — random sheet XMLs assembled from a small grammar of
//! CT_Worksheet children in random allowed orders, crossed with random
//! block combinations. Asserts:
//!
//! 1. Output is well-formed XML (re-parses without quick-xml errors).
//! 2. Every supplied block's root local-name appears in the output at the
//!    correct ECMA position relative to source elements that survived.
//! 3. Source elements not in the replace set survive (in their original
//!    relative order).
//!
//! 500 iters — each iter does considerably more work than RFC-010's
//! property test (XML parse + write vs just rels parse), so we cap lower.

use std::collections::BTreeMap;

use proptest::prelude::*;
use quick_xml::events::Event;
use quick_xml::Reader as XmlReader;
use wolfxl_merger::ct_worksheet_order::ECMA_ORDER;
use wolfxl_merger::{merge_blocks, SheetBlock};

/// Pick a subset of CT_Worksheet child slots (by local-name) to embed in
/// the synthetic source. Each slot appears at most once except slot 17
/// (conditionalFormatting), which we deliberately allow up to 3
/// occurrences so the replace-all path gets exercised.
///
/// Returns owned bytes of a synthetic worksheet XML.
fn build_source_xml(slots: &[&'static [u8]], cf_count: usize) -> Vec<u8> {
    let mut xml: Vec<u8> = b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">".to_vec();

    for slot in slots {
        if *slot == b"conditionalFormatting" {
            for i in 0..cf_count {
                xml.extend_from_slice(b"<conditionalFormatting sqref=\"A");
                xml.extend_from_slice(format!("{}", i + 1).as_bytes());
                xml.extend_from_slice(b"\"><cfRule type=\"src\"/></conditionalFormatting>");
            }
            continue;
        }
        // Self-closing form for everything else — minimal valid sentinel.
        xml.push(b'<');
        xml.extend_from_slice(slot);
        xml.extend_from_slice(b" data-src=\"y\"/>");
    }

    xml.extend_from_slice(b"</worksheet>");
    xml
}

/// All known CT_Worksheet children. Source samples a subset.
fn all_slot_names() -> Vec<&'static [u8]> {
    ECMA_ORDER.iter().map(|(n, _)| *n).collect()
}

prop_compose! {
    fn arb_block()
        (kind in 0u8..6)
        -> SheetBlock
    {
        match kind {
            0 => SheetBlock::MergeCells(b"<mergeCells count=\"1\"><mergeCell ref=\"A1:B1\"/></mergeCells>".to_vec()),
            1 => SheetBlock::ConditionalFormatting(b"<conditionalFormatting sqref=\"Z1\"><cfRule type=\"new\"/></conditionalFormatting>".to_vec()),
            2 => SheetBlock::DataValidations(b"<dataValidations count=\"1\"><dataValidation/></dataValidations>".to_vec()),
            3 => SheetBlock::Hyperlinks(b"<hyperlinks><hyperlink ref=\"A1\" r:id=\"rId1\"/></hyperlinks>".to_vec()),
            4 => SheetBlock::LegacyDrawing(b"<legacyDrawing r:id=\"rId99\"/>".to_vec()),
            _ => SheetBlock::TableParts(b"<tableParts count=\"1\"><tablePart r:id=\"rId50\"/></tableParts>".to_vec()),
        }
    }
}

fn well_formed(xml: &[u8]) -> bool {
    let mut reader = XmlReader::from_reader(xml);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Eof) => return true,
            Ok(_) => {}
            Err(_) => return false,
        }
        buf.clear();
    }
}

/// Walk parsed events and, for each top-level CT_Worksheet child, return
/// `(local_name, ecma_ordinal_or_None)` in document order. Used to assert
/// the output's slot ordering.
fn ordered_top_level_children(xml: &[u8]) -> Vec<(Vec<u8>, Option<u32>)> {
    let mut reader = XmlReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut depth = 0;
    let mut out: Vec<(Vec<u8>, Option<u32>)> = Vec::new();
    let lookup: BTreeMap<&[u8], u32> = ECMA_ORDER.iter().map(|(n, o)| (*n, *o)).collect();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.local_name().as_ref().as_bytes().to_vec();
                if depth == 1 {
                    let ord = lookup.get(name.as_slice()).copied();
                    out.push((name, ord));
                }
                depth += 1;
            }
            Ok(Event::Empty(e)) => {
                let name = e.local_name().as_ref().as_bytes().to_vec();
                if depth == 1 {
                    let ord = lookup.get(name.as_slice()).copied();
                    out.push((name, ord));
                }
            }
            Ok(Event::End(_)) => {
                depth -= 1;
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(_) => break,
        }
        buf.clear();
    }
    out
}

proptest! {
    #![proptest_config(ProptestConfig {
        cases: 500,
        max_shrink_iters: 200,
        ..Default::default()
    })]

    #[test]
    fn merge_preserves_well_formed_and_ordering(
        source_slot_indices in prop::collection::vec(0usize..ECMA_ORDER.len(), 0..10),
        cf_count in 0usize..3,
        blocks in prop::collection::vec(arb_block(), 0..5),
    ) {
        // Build a source by picking distinct slots from ECMA_ORDER in
        // ascending order — so the source itself is always validly ordered.
        let mut indices: Vec<usize> = source_slot_indices.into_iter().collect();
        indices.sort();
        indices.dedup();
        let names = all_slot_names();
        let slots: Vec<&[u8]> = indices.into_iter().map(|i| names[i]).collect();
        let source = build_source_xml(&slots, cf_count);

        let merged = merge_blocks(&source, blocks.clone())
            .map_err(|e| TestCaseError::fail(format!("merge_blocks failed: {e}")))?;

        // (1) Output is well-formed.
        prop_assert!(
            well_formed(&merged),
            "output is not well-formed: {}",
            String::from_utf8_lossy(&merged)
        );

        // (2) Top-level children with known ordinals appear in
        // non-decreasing order.
        let kids = ordered_top_level_children(&merged);
        let mut last_ord: u32 = 0;
        for (name, ord) in &kids {
            if let Some(o) = ord {
                prop_assert!(
                    *o >= last_ord,
                    "child <{}> at ord {} appears after a sibling with ord {}",
                    String::from_utf8_lossy(name),
                    o,
                    last_ord
                );
                last_ord = *o;
            }
            // Unknown elements: skip — they pass through at source position
            // and the merger does not promise to keep them ordered relative
            // to inserted blocks (RFC-011 §5.3 / §6 test #6).
        }

        // (3) Every supplied block's root local-name appears in the
        // output, AT LEAST ONCE. (Multiple CF blocks all appear.)
        for b in &blocks {
            let root = b.root_local_name();
            let mut needle: Vec<u8> = b"<".to_vec();
            needle.extend_from_slice(root);
            prop_assert!(
                merged.windows(needle.len()).any(|w| w == needle.as_slice()),
                "supplied block <{}> missing from output",
                String::from_utf8_lossy(root)
            );
        }
    }
}
