//! Hyperlink extraction + block emission for the modify-mode patcher (RFC-022).
//!
//! Used by `XlsxPatcher::do_save`'s Phase 2.5e to:
//! 1. Read the existing `<hyperlinks>` block from a sheet's XML.
//! 2. Merge user-supplied [`HyperlinkOp`]s on top.
//! 3. Emit a fresh `<hyperlinks>` block whose bytes feed
//!    `wolfxl_merger::SheetBlock::Hyperlinks` (slot 19).
//! 4. Mutate the sheet's [`RelsGraph`] for added/removed external links.
//!
//! The native writer's `crates/wolfxl-writer/src/emit/sheet_xml.rs::emit_hyperlinks`
//! is the lift-and-shift template. TODO: consolidate into a shared crate when a
//! third caller appears (RFC-020 §4.2 Option 2 precedent).

use std::collections::BTreeMap;

use quick_xml::events::Event;
use quick_xml::Reader as XmlReader;

use wolfxl_rels::{rt, RelId, RelsGraph, TargetMode};

// ---------------------------------------------------------------------------
// Public types
// ---------------------------------------------------------------------------

/// One user-supplied hyperlink edit. At least one of `target` or `location`
/// must be `Some` — the patcher's `queue_hyperlink` PyO3 wrapper validates
/// that before constructing the patch.
///
/// `target` is an external URL (rendered as a `<Relationship TargetMode="External">`
/// in the sheet's rels graph). `location` is an internal sheet reference
/// (e.g. `'Sheet2'!A1`, no leading `#`). The two are mutually exclusive in
/// the serialized `<hyperlink>` element; `target` wins when both are set.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct HyperlinkPatch {
    pub coordinate: String,
    pub target: Option<String>,
    pub location: Option<String>,
    pub tooltip: Option<String>,
    pub display: Option<String>,
}

/// `Set` adds-or-overwrites the hyperlink at `coordinate`; `Delete` removes it
/// (and any associated external rId in the rels graph).
#[derive(Debug, Clone)]
pub enum HyperlinkOp {
    Set(HyperlinkPatch),
    Delete,
}

/// One hyperlink already present in the source sheet XML, with its `r:id`
/// resolved against the sheet's rels graph so callers see the URL itself
/// rather than an opaque relationship handle.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExistingHyperlink {
    pub coordinate: String,
    pub target: Option<String>,
    pub location: Option<String>,
    pub tooltip: Option<String>,
    pub display: Option<String>,
    pub rid: Option<RelId>,
}

// ---------------------------------------------------------------------------
// Extract
// ---------------------------------------------------------------------------

/// Parse the `<hyperlinks>` block out of a sheet XML, resolving each
/// `r:id` against `rels` so callers see the external URL alongside the rId.
///
/// Source-key is the raw `ref` attribute, so range refs (`A1:C3`) round-trip
/// verbatim. Returns an empty map when the sheet has no `<hyperlinks>` block
/// (which is the common case — most sheets have no hyperlinks at all).
///
/// Tolerant parser: malformed XML is bailed on early with whatever entries
/// were successfully read. Hyperlinks with no `ref` attribute are skipped.
pub fn extract_hyperlinks(
    sheet_xml: &[u8],
    rels: &RelsGraph,
) -> BTreeMap<String, ExistingHyperlink> {
    let mut out: BTreeMap<String, ExistingHyperlink> = BTreeMap::new();
    let mut reader = XmlReader::from_reader(sheet_xml);
    reader.config_mut().trim_text(true);
    let mut buf: Vec<u8> = Vec::new();
    let mut in_block = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if e.local_name().as_ref() == "hyperlinks" {
                    in_block = true;
                }
            }
            Ok(Event::End(e)) => {
                if e.local_name().as_ref() == "hyperlinks" {
                    in_block = false;
                }
            }
            Ok(Event::Empty(e)) if in_block && e.local_name().as_ref() == "hyperlink" => {
                if let Some(parsed) = parse_one(&e, rels) {
                    out.insert(parsed.coordinate.clone(), parsed);
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    out
}

fn parse_one(e: &quick_xml::events::BytesStart<'_>, rels: &RelsGraph) -> Option<ExistingHyperlink> {
    let mut coord = String::new();
    let mut rid_str: Option<String> = None;
    let mut location: Option<String> = None;
    let mut tooltip: Option<String> = None;
    let mut display: Option<String> = None;

    for a in e.attributes().with_checks(false).flatten() {
        let key = a.key.as_ref();
        let val = a
            .normalized_value(quick_xml::XmlVersion::Implicit1_0)
            .map(|v| v.into_owned())
            .unwrap_or_else(|_| a.value.to_string());
        match key {
            "ref" => coord = val,
            // Real files use the prefixed form `r:id`. The `id` fallback is
            // for documents that bind the relationships namespace to a
            // different prefix or use no prefix at all.
            "r:id" | "id" => rid_str = Some(val),
            "location" => location = Some(val),
            "tooltip" => tooltip = Some(val),
            "display" => display = Some(val),
            _ => {}
        }
    }
    if coord.is_empty() {
        return None;
    }
    let rid = rid_str.map(RelId);
    let target = rid
        .as_ref()
        .and_then(|r| rels.get(r).map(|rel| rel.target.clone()));
    Some(ExistingHyperlink {
        coordinate: coord,
        target,
        location,
        tooltip,
        display,
        rid,
    })
}

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

/// Produce the merged `<hyperlinks>` block bytes for a sheet.
///
/// Returns `(block_bytes, deleted_rids)`. Block bytes are the bytes of the
/// `<hyperlinks>...</hyperlinks>` element (or `Vec::new()` when the merged
/// set is empty — in which case the caller should not push a `SheetBlock`).
/// `deleted_rids` is the list of rIds removed from `rels` as a side-effect;
/// `rels` itself has already been mutated by the time this function returns.
///
/// Pure on `(existing, ops)` modulo the `&mut RelsGraph`. The mutation is
/// part of the contract — we allocate fresh rIds for new external links and
/// remove rIds whose hyperlinks were deleted or overwritten.
pub fn build_hyperlinks_block(
    existing: BTreeMap<String, ExistingHyperlink>,
    ops: &BTreeMap<String, HyperlinkOp>,
    rels: &mut RelsGraph,
) -> (Vec<u8>, Vec<RelId>) {
    #[derive(Debug, Clone)]
    struct Merged {
        target: Option<String>,
        location: Option<String>,
        tooltip: Option<String>,
        display: Option<String>,
        rid: Option<RelId>,
    }

    let mut merged: BTreeMap<String, Merged> = existing
        .into_iter()
        .map(|(c, e)| {
            (
                c,
                Merged {
                    target: e.target,
                    location: e.location,
                    tooltip: e.tooltip,
                    display: e.display,
                    rid: e.rid,
                },
            )
        })
        .collect();

    let mut deleted: Vec<RelId> = Vec::new();
    for (coord, op) in ops {
        match op {
            HyperlinkOp::Set(patch) => {
                if let Some(prev) = merged.get(coord) {
                    if let Some(rid) = &prev.rid {
                        deleted.push(rid.clone());
                    }
                }
                let rid = patch
                    .target
                    .as_deref()
                    .map(|t| rels.add(rt::HYPERLINK, t, TargetMode::External));
                merged.insert(
                    coord.clone(),
                    Merged {
                        target: patch.target.clone(),
                        location: patch.location.clone(),
                        tooltip: patch.tooltip.clone(),
                        display: patch.display.clone(),
                        rid,
                    },
                );
            }
            HyperlinkOp::Delete => {
                if let Some(removed) = merged.remove(coord) {
                    if let Some(rid) = removed.rid {
                        deleted.push(rid);
                    }
                }
            }
        }
    }

    for rid in &deleted {
        rels.remove(rid);
    }

    if merged.is_empty() {
        return (Vec::new(), deleted);
    }

    let mut out = String::with_capacity(64 + merged.len() * 96);
    out.push_str("<hyperlinks>");
    for (coord, hl) in &merged {
        out.push_str("<hyperlink ref=\"");
        push_xml_attr_escape(&mut out, coord);
        out.push('"');
        if let Some(rid) = &hl.rid {
            out.push_str(" r:id=\"");
            push_xml_attr_escape(&mut out, &rid.0);
            out.push('"');
        } else if let Some(loc) = &hl.location {
            out.push_str(" location=\"");
            push_xml_attr_escape(&mut out, loc);
            out.push('"');
        }
        if let Some(d) = &hl.display {
            out.push_str(" display=\"");
            push_xml_attr_escape(&mut out, d);
            out.push('"');
        }
        if let Some(t) = &hl.tooltip {
            out.push_str(" tooltip=\"");
            push_xml_attr_escape(&mut out, t);
            out.push('"');
        }
        out.push_str("/>");
    }
    out.push_str("</hyperlinks>");
    (out.into_bytes(), deleted)
}

fn push_xml_attr_escape(out: &mut String, s: &str) {
    for ch in s.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&apos;"),
            _ => out.push(ch),
        }
    }
}

// ---------------------------------------------------------------------------
// Tests
//
// Inline pure-Rust tests. The patcher's cdylib does not link standalone via
// `cargo test -p wolfxl --lib` (Python linkage), so these compile under
// `cargo build` and behavior is exercised via pytest in commit 5. Same
// precedent as `ancillary.rs` and `properties.rs`.
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn entity_encoded_hyperlink_attributes_roundtrip() {
        let xml = sheet_xml_with_hyperlinks(
            r#"<hyperlink ref="A&#49;" location="&apos;A&amp;B&apos;!A1" tooltip="A&#32;&amp;&#32;B" display="&quot;Go&quot;"/>"#,
        );
        let mut rels = RelsGraph::new();
        let existing = extract_hyperlinks(xml.as_bytes(), &rels);
        assert_eq!(existing["A1"].location.as_deref(), Some("'A&B'!A1"));
        assert_eq!(existing["A1"].tooltip.as_deref(), Some("A & B"));
        assert_eq!(existing["A1"].display.as_deref(), Some("\"Go\""));
        let (bytes, _) = build_hyperlinks_block(existing.clone(), &BTreeMap::new(), &mut rels);
        assert_eq!(extract_hyperlinks(&bytes, &rels), existing);
    }

    fn rels_with_hyperlinks(targets: &[&str]) -> RelsGraph {
        let mut g = RelsGraph::new();
        for t in targets {
            g.add(rt::HYPERLINK, t, TargetMode::External);
        }
        g
    }

    fn sheet_xml_with_hyperlinks(body: &str) -> String {
        format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheetData/>
<hyperlinks>{body}</hyperlinks>
</worksheet>"#
        )
    }

    #[test]
    fn extract_hyperlinks_resolves_rids() {
        let rels = rels_with_hyperlinks(&["https://example.com/docs", "mailto:test@example.com"]);
        let xml = sheet_xml_with_hyperlinks(
            r#"<hyperlink ref="A1" r:id="rId1"/>
<hyperlink ref="A2" r:id="rId2" tooltip="Email"/>
<hyperlink ref="A3" location="'Sheet2'!A1"/>"#,
        );
        let out = extract_hyperlinks(xml.as_bytes(), &rels);
        assert_eq!(out.len(), 3);
        assert_eq!(
            out["A1"].target.as_deref(),
            Some("https://example.com/docs")
        );
        assert_eq!(out["A1"].rid, Some(RelId("rId1".into())));
        assert_eq!(out["A2"].tooltip.as_deref(), Some("Email"));
        assert!(out["A3"].rid.is_none());
        assert_eq!(out["A3"].location.as_deref(), Some("'Sheet2'!A1"));
    }

    #[test]
    fn build_block_preserves_existing_rids() {
        // Existing rId1 + rId2 external; add a third → output has rId1, rId2,
        // and a fresh rId3. The sources of the first two never touch the rels
        // graph, so their numbers are stable.
        let mut rels = rels_with_hyperlinks(&["https://example.com/a", "https://example.com/b"]);
        let mut existing = BTreeMap::new();
        existing.insert(
            "A1".into(),
            ExistingHyperlink {
                coordinate: "A1".into(),
                target: Some("https://example.com/a".into()),
                location: None,
                tooltip: None,
                display: None,
                rid: Some(RelId("rId1".into())),
            },
        );
        existing.insert(
            "A2".into(),
            ExistingHyperlink {
                coordinate: "A2".into(),
                target: Some("https://example.com/b".into()),
                location: None,
                tooltip: None,
                display: None,
                rid: Some(RelId("rId2".into())),
            },
        );
        let mut ops = BTreeMap::new();
        ops.insert(
            "A3".into(),
            HyperlinkOp::Set(HyperlinkPatch {
                coordinate: "A3".into(),
                target: Some("https://example.com/c".into()),
                location: None,
                tooltip: None,
                display: None,
            }),
        );

        let (bytes, deleted) = build_hyperlinks_block(existing, &ops, &mut rels);
        let text = std::str::from_utf8(&bytes).unwrap();
        assert!(deleted.is_empty(), "no deletions");
        assert!(text.contains(r#"ref="A1" r:id="rId1""#));
        assert!(text.contains(r#"ref="A2" r:id="rId2""#));
        assert!(text.contains(r#"ref="A3" r:id="rId3""#));
        // Rels has all three; rId1 and rId2 untouched.
        assert_eq!(rels.len(), 3);
        assert_eq!(
            rels.get(&RelId("rId1".into())).unwrap().target,
            "https://example.com/a"
        );
    }

    #[test]
    fn delete_external_removes_rid() {
        let mut rels = rels_with_hyperlinks(&["https://example.com/a"]);
        let mut existing = BTreeMap::new();
        existing.insert(
            "A1".into(),
            ExistingHyperlink {
                coordinate: "A1".into(),
                target: Some("https://example.com/a".into()),
                location: None,
                tooltip: None,
                display: None,
                rid: Some(RelId("rId1".into())),
            },
        );
        let mut ops = BTreeMap::new();
        ops.insert("A1".into(), HyperlinkOp::Delete);

        let (bytes, deleted) = build_hyperlinks_block(existing, &ops, &mut rels);
        assert!(bytes.is_empty(), "empty merged → empty bytes");
        assert_eq!(deleted, vec![RelId("rId1".into())]);
        assert!(rels.is_empty(), "rels graph drained");
    }

    #[test]
    fn internal_link_no_rid() {
        let mut rels = RelsGraph::new();
        let mut ops = BTreeMap::new();
        ops.insert(
            "A1".into(),
            HyperlinkOp::Set(HyperlinkPatch {
                coordinate: "A1".into(),
                target: None,
                location: Some("'Sheet2'!A1".into()),
                tooltip: None,
                display: None,
            }),
        );
        let (bytes, deleted) = build_hyperlinks_block(BTreeMap::new(), &ops, &mut rels);
        let text = std::str::from_utf8(&bytes).unwrap();
        assert!(deleted.is_empty());
        assert!(rels.is_empty(), "internal link allocates no rId");
        assert!(text.contains(r#"location="&apos;Sheet2&apos;!A1""#));
        assert!(!text.contains("r:id"));
    }

    #[test]
    fn mixed_internal_external_in_one_block() {
        let mut rels = RelsGraph::new();
        let mut ops = BTreeMap::new();
        ops.insert(
            "A1".into(),
            HyperlinkOp::Set(HyperlinkPatch {
                coordinate: "A1".into(),
                target: Some("https://x.com".into()),
                location: None,
                tooltip: None,
                display: None,
            }),
        );
        ops.insert(
            "A2".into(),
            HyperlinkOp::Set(HyperlinkPatch {
                coordinate: "A2".into(),
                target: None,
                location: Some("Sheet2!A1".into()),
                tooltip: None,
                display: None,
            }),
        );
        ops.insert(
            "A3".into(),
            HyperlinkOp::Set(HyperlinkPatch {
                coordinate: "A3".into(),
                target: Some("https://y.com".into()),
                location: None,
                tooltip: None,
                display: None,
            }),
        );
        let (bytes, _) = build_hyperlinks_block(BTreeMap::new(), &ops, &mut rels);
        let text = std::str::from_utf8(&bytes).unwrap();
        // BTreeMap order keeps A1 < A2 < A3.
        let p1 = text.find(r#"ref="A1""#).unwrap();
        let p2 = text.find(r#"ref="A2""#).unwrap();
        let p3 = text.find(r#"ref="A3""#).unwrap();
        assert!(p1 < p2 && p2 < p3);
        assert!(text[p2..p3].contains("location="));
        assert!(text[p2..p3].matches("r:id").count() == 0);
    }

    #[test]
    fn attr_escape_url_with_ampersand() {
        // The URL itself is stored verbatim in memory; the escape happens at
        // serialization time. The rels-graph caller uses `wolfxl_rels`'s
        // serializer so the rels file gets `&amp;` too — separate concern,
        // tested in `wolfxl-rels`.
        let mut rels = RelsGraph::new();
        let mut ops = BTreeMap::new();
        ops.insert(
            "A1".into(),
            HyperlinkOp::Set(HyperlinkPatch {
                coordinate: "A1".into(),
                target: Some("https://x.com/path?a=1&b=2".into()),
                location: None,
                tooltip: None,
                display: None,
            }),
        );
        let (bytes, _) = build_hyperlinks_block(BTreeMap::new(), &ops, &mut rels);
        // Hyperlinks block emits an r:id only — the URL itself goes into the
        // rels file and is escaped there. So the block bytes don't carry the
        // ampersand. Confirm the rels graph holds the unescaped URL and its
        // serializer handles the escape.
        let text = std::str::from_utf8(&bytes).unwrap();
        assert!(text.contains("r:id=\"rId1\""));
        let rels_bytes = rels.serialize();
        let rels_text = std::str::from_utf8(&rels_bytes).unwrap();
        assert!(rels_text.contains("a=1&amp;b=2"));
    }

    #[test]
    fn tooltip_and_display_emitted() {
        let mut rels = RelsGraph::new();
        let mut ops = BTreeMap::new();
        ops.insert(
            "B5".into(),
            HyperlinkOp::Set(HyperlinkPatch {
                coordinate: "B5".into(),
                target: Some("https://x.com".into()),
                location: None,
                tooltip: Some("Click me".into()),
                display: Some("X Site".into()),
            }),
        );
        let (bytes, _) = build_hyperlinks_block(BTreeMap::new(), &ops, &mut rels);
        let text = std::str::from_utf8(&bytes).unwrap();
        assert!(text.contains(r#"display="X Site""#));
        assert!(text.contains(r#"tooltip="Click me""#));
    }

    #[test]
    fn empty_existing_plus_one_set_op_produces_single_link_block() {
        let mut rels = RelsGraph::new();
        let mut ops = BTreeMap::new();
        ops.insert(
            "A1".into(),
            HyperlinkOp::Set(HyperlinkPatch {
                coordinate: "A1".into(),
                target: Some("https://x.com".into()),
                location: None,
                tooltip: None,
                display: None,
            }),
        );
        let (bytes, _) = build_hyperlinks_block(BTreeMap::new(), &ops, &mut rels);
        let text = std::str::from_utf8(&bytes).unwrap();
        assert_eq!(rels.len(), 1);
        // Exactly one <hyperlink/> in the output.
        assert_eq!(text.matches("<hyperlink ").count(), 1);
        assert!(text.starts_with("<hyperlinks>"));
        assert!(text.ends_with("</hyperlinks>"));
    }

    #[test]
    fn overwrite_existing_external_removes_old_rid() {
        // Setting a new URL on a coordinate that previously had an external
        // hyperlink: the old rId is queued for deletion, a fresh rId is
        // allocated for the new URL.
        let mut rels = rels_with_hyperlinks(&["https://old.com"]);
        let mut existing = BTreeMap::new();
        existing.insert(
            "A1".into(),
            ExistingHyperlink {
                coordinate: "A1".into(),
                target: Some("https://old.com".into()),
                location: None,
                tooltip: None,
                display: None,
                rid: Some(RelId("rId1".into())),
            },
        );
        let mut ops = BTreeMap::new();
        ops.insert(
            "A1".into(),
            HyperlinkOp::Set(HyperlinkPatch {
                coordinate: "A1".into(),
                target: Some("https://new.com".into()),
                location: None,
                tooltip: None,
                display: None,
            }),
        );
        let (bytes, deleted) = build_hyperlinks_block(existing, &ops, &mut rels);
        let text = std::str::from_utf8(&bytes).unwrap();
        assert_eq!(deleted, vec![RelId("rId1".into())]);
        // rId2 was allocated by add() (next_rid is monotonic — see RFC-010).
        assert!(text.contains(r#"r:id="rId2""#));
        assert_eq!(rels.len(), 1, "rId1 removed, rId2 added");
    }
}
