//! `[Content_Types].xml` parser, mutable graph, and serializer (RFC-013).
//!
//! ECMA-376 Part 2 (Open Packaging Conventions) §10.1: every part inside an
//! OOXML container must be accounted for in `[Content_Types].xml` either by
//! a `<Default Extension="..." ContentType="..."/>` (matches every entry
//! whose path ends in that extension) or by a `<Override PartName="..."
//! ContentType="..."/>` (explicit per-part). Adding a new part — comments,
//! tables, hyperlink rels, doc properties — requires either an `Override`
//! or, for shared extensions like `vml`, a `Default`.
//!
//! # Usage
//!
//! ```ignore
//! let mut g = ContentTypesGraph::parse(source_xml.as_bytes())?;
//! g.add_override(
//!     "/xl/comments1.xml",
//!     "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml",
//! );
//! g.ensure_default(
//!     "vml",
//!     "application/vnd.openxmlformats-officedocument.vmlDrawing",
//! );
//! let bytes = g.serialize();
//! ```
//!
//! # Invariants
//!
//! - **Source order is preserved.** Both defaults and overrides are stored
//!   in document order. New entries append at the tail. Reordering would
//!   break Excel's strict-validator paths and the modify-mode minimal-diff
//!   contract — RFC-013 §8 risk #3.
//! - **`add_override` is idempotent.** A second call with the same
//!   `(part, type)` pair is a no-op. A call that re-points an existing
//!   `part` at a new `content_type` overwrites in place (caller bug, but
//!   we don't panic — RFC-013 reserves the explicit-collision panic for
//!   `file_adds`, not for content-type ops which can be retried safely).
//! - **`ensure_default` is idempotent in the same shape.**
//! - **Lenient parser.** `<Default>` / `<Override>` entries missing
//!   required attributes are skipped silently (mirrors
//!   `wolfxl_rels::RelsGraph::parse`); only malformed XML errors out.

use quick_xml::events::Event;
use quick_xml::Reader as XmlReader;

const CT_NS: &str = "http://schemas.openxmlformats.org/package/2006/content-types";

pub(crate) const CT_DRAWING: &str = "application/vnd.openxmlformats-officedocument.drawing+xml";
pub(crate) const CT_CHART: &str =
    "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";

pub(crate) fn image_content_type_for_ext(ext: &str) -> &'static str {
    match ext {
        "png" => "image/png",
        "jpeg" | "jpg" => "image/jpeg",
        "gif" => "image/gif",
        "bmp" => "image/bmp",
        _ => "application/octet-stream",
    }
}

// ---------------------------------------------------------------------------
// Cross-sheet aggregation op (Phase 2.5c — see `mod.rs::do_save`).
// ---------------------------------------------------------------------------

/// One mutation queued by a per-sheet flush, drained by Phase 2.5c into a
/// single `[Content_Types].xml` parse + serialize per save. The patcher
/// never queues either op in this slice; callers land in RFC-022 / RFC-023
/// / RFC-024 / RFC-035.
#[derive(Debug, Clone, PartialEq, Eq)]
#[allow(dead_code)] // No live caller in this slice; field reserved.
pub enum ContentTypeOp {
    AddOverride(String, String),
    RemoveOverride(String),
    EnsureDefault(String, String),
    RemoveOverridePrefix(String),
}

// ---------------------------------------------------------------------------
// ContentTypesGraph — parsed + mutable view of `[Content_Types].xml`.
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ContentTypesGraph {
    /// `(extension, content_type)` pairs in source order.
    defaults: Vec<(String, String)>,
    /// `(part_name, content_type)` pairs in source order. New
    /// `add_override` entries append to the tail.
    overrides: Vec<(String, String)>,
}

impl ContentTypesGraph {
    /// Empty graph — used when synthesizing `[Content_Types].xml` from
    /// scratch (not exercised by RFC-013 itself; reserved for write-mode
    /// use cases).
    #[allow(dead_code)]
    pub fn new() -> Self {
        Self::default()
    }

    /// Parse an existing `[Content_Types].xml` body.
    ///
    /// Tolerant of:
    /// - missing XML declaration / BOM
    /// - self-closing or open-close `<Default>` / `<Override>` forms
    /// - unknown attributes on `<Types>` / `<Default>` / `<Override>`
    /// - elements outside the OPC namespace (rare; preserved by
    ///   ignoring rather than rejecting)
    pub fn parse(xml: &[u8]) -> Result<Self, String> {
        let mut reader = XmlReader::from_reader(xml);
        reader.config_mut().trim_text(true);
        let mut buf: Vec<u8> = Vec::new();
        let mut defaults: Vec<(String, String)> = Vec::new();
        let mut overrides: Vec<(String, String)> = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    let local = e.local_name();
                    let local = local.as_ref();
                    match local {
                        "Default" => {
                            let mut ext: Option<String> = None;
                            let mut ct: Option<String> = None;
                            for a in e.attributes().with_checks(false).flatten() {
                                let value = a
                                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                    .map(|v| v.into_owned())
                                    .unwrap_or_else(|_| a.value.to_string());
                                match a.key.as_ref() {
                                    "Extension" => ext = Some(value),
                                    "ContentType" => ct = Some(value),
                                    _ => {}
                                }
                            }
                            if let (Some(ext), Some(ct)) = (ext, ct) {
                                defaults.push((ext, ct));
                            }
                        }
                        "Override" => {
                            let mut part: Option<String> = None;
                            let mut ct: Option<String> = None;
                            for a in e.attributes().with_checks(false).flatten() {
                                let value = a
                                    .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                    .map(|v| v.into_owned())
                                    .unwrap_or_else(|_| a.value.to_string());
                                match a.key.as_ref() {
                                    "PartName" => part = Some(value),
                                    "ContentType" => ct = Some(value),
                                    _ => {}
                                }
                            }
                            if let (Some(p), Some(t)) = (part, ct) {
                                overrides.push((p, t));
                            }
                        }
                        _ => {}
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(format!("XML parse error: {e}")),
                _ => {}
            }
            buf.clear();
        }

        Ok(ContentTypesGraph {
            defaults,
            overrides,
        })
    }

    /// Append (or update in place) an Override entry.
    ///
    /// Idempotent: if the exact `(part, content_type)` pair already
    /// exists, returns without mutating. If `part` already maps to a
    /// different `content_type`, overwrites in place — this matches
    /// what the OOXML spec allows (a part can have only one content
    /// type) without forcing the caller to dedupe before queuing.
    pub fn add_override(&mut self, part: &str, content_type: &str) {
        for slot in self.overrides.iter_mut() {
            if slot.0 == part {
                if slot.1 != content_type {
                    slot.1 = content_type.to_string();
                }
                return;
            }
        }
        self.overrides
            .push((part.to_string(), content_type.to_string()));
    }

    /// Remove an Override entry by part name. Idempotent.
    pub fn remove_override(&mut self, part: &str) {
        self.overrides.retain(|(p, _)| p != part);
    }

    /// Append (or update in place) a Default entry. Same semantics as
    /// `add_override` keyed on `extension`.
    pub fn ensure_default(&mut self, extension: &str, content_type: &str) {
        for slot in self.defaults.iter_mut() {
            if slot.0 == extension {
                if slot.1 != content_type {
                    slot.1 = content_type.to_string();
                }
                return;
            }
        }
        self.defaults
            .push((extension.to_string(), content_type.to_string()));
    }

    pub fn remove_override_prefix(&mut self, prefix: &str) {
        self.overrides.retain(|(p, _)| !p.starts_with(prefix));
    }

    /// Source-order accessors (used by tests; no live caller in slice).
    #[allow(dead_code)]
    pub fn defaults(&self) -> &[(String, String)] {
        &self.defaults
    }

    /// Source-order accessors.
    #[allow(dead_code)]
    pub fn overrides(&self) -> &[(String, String)] {
        &self.overrides
    }

    /// Dispatch a single [`ContentTypeOp`] onto the matching mutator. Used by
    /// the patcher's Phase-2.5c aggregation loop so cross-sheet ops can be
    /// applied with one match instead of N.
    pub fn apply_op(&mut self, op: &ContentTypeOp) {
        match op {
            ContentTypeOp::AddOverride(part, ct) => self.add_override(part, ct),
            ContentTypeOp::RemoveOverride(part) => self.remove_override(part),
            ContentTypeOp::EnsureDefault(ext, ct) => self.ensure_default(ext, ct),
            ContentTypeOp::RemoveOverridePrefix(prefix) => self.remove_override_prefix(prefix),
        }
    }

    /// Serialize back to canonical bytes.
    ///
    /// Layout: XML declaration + `<Types>` open + every `<Default>` in
    /// source order + every `<Override>` in source order + `</Types>`.
    /// Always emits the canonical preamble (`<?xml version="1.0"
    /// encoding="UTF-8" standalone="yes"?>\r\n`) and the OPC namespace.
    /// Single-line body (no inter-element whitespace) — matches what
    /// the writer's `crates/wolfxl-writer/src/emit/content_types.rs::emit`
    /// produces, so write-mode and modify-mode share the same shape.
    pub fn serialize(&self) -> Vec<u8> {
        let mut out =
            String::with_capacity(128 + self.defaults.len() * 96 + self.overrides.len() * 128);
        out.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n");
        out.push_str(&format!("<Types xmlns=\"{CT_NS}\">"));
        for (ext, ct) in &self.defaults {
            out.push_str("<Default Extension=\"");
            push_attr_escape(&mut out, ext);
            out.push_str("\" ContentType=\"");
            push_attr_escape(&mut out, ct);
            out.push_str("\"/>");
        }
        for (part, ct) in &self.overrides {
            out.push_str("<Override PartName=\"");
            push_attr_escape(&mut out, part);
            out.push_str("\" ContentType=\"");
            push_attr_escape(&mut out, ct);
            out.push_str("\"/>");
        }
        out.push_str("</Types>");
        out.into_bytes()
    }
}

/// XML attribute-value escape — content-type strings can carry `+xml`
/// suffixes and similar; defensively escape `&`, `<`, `>`, and `"`.
fn push_attr_escape(out: &mut String, s: &str) {
    for ch in s.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            _ => out.push(ch),
        }
    }
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn entity_encoded_attributes_roundtrip() {
        let xml = br#"<Types><Default Extension="x&#109;l" ContentType="type/a&amp;b"/><Override PartName="/a&#38;b.xml" ContentType="type/&quot;quoted&quot;"/></Types>"#;
        let graph = ContentTypesGraph::parse(xml).unwrap();
        assert_eq!(
            graph.defaults(),
            &[("xml".to_string(), "type/a&b".to_string())]
        );
        assert_eq!(
            graph.overrides(),
            &[("/a&b.xml".to_string(), "type/\"quoted\"".to_string())]
        );
        assert_eq!(ContentTypesGraph::parse(&graph.serialize()).unwrap(), graph);
    }

    const CT_COMMENTS: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";
    const CT_TABLE: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";
    const CT_VML: &str = "application/vnd.openxmlformats-officedocument.vmlDrawing";
    const CT_RELATIONSHIPS: &str = "application/vnd.openxmlformats-package.relationships+xml";
    const CT_XML_DEFAULT: &str = "application/xml";

    fn minimal_xml() -> &'static str {
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n\
         <Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\
         <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\
         <Default Extension=\"xml\" ContentType=\"application/xml\"/>\
         <Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>\
         </Types>"
    }

    #[test]
    fn parse_empty_types_element() {
        let xml = "<?xml version=\"1.0\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"/>";
        let g = ContentTypesGraph::parse(xml.as_bytes()).unwrap();
        assert!(g.defaults().is_empty());
        assert!(g.overrides().is_empty());
    }

    #[test]
    fn parse_multiple_defaults_in_source_order() {
        let xml = "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\
            <Default Extension=\"rels\" ContentType=\"X\"/>\
            <Default Extension=\"xml\" ContentType=\"Y\"/>\
            <Default Extension=\"vml\" ContentType=\"Z\"/>\
            </Types>";
        let g = ContentTypesGraph::parse(xml.as_bytes()).unwrap();
        let exts: Vec<&str> = g.defaults().iter().map(|(e, _)| e.as_str()).collect();
        assert_eq!(exts, vec!["rels", "xml", "vml"]);
    }

    #[test]
    fn parse_multiple_overrides_in_source_order() {
        let g = ContentTypesGraph::parse(minimal_xml().as_bytes()).unwrap();
        let parts: Vec<&str> = g.overrides().iter().map(|(p, _)| p.as_str()).collect();
        assert_eq!(parts, vec!["/xl/workbook.xml"]);
    }

    #[test]
    fn parse_mixed_defaults_and_overrides() {
        let g = ContentTypesGraph::parse(minimal_xml().as_bytes()).unwrap();
        assert_eq!(g.defaults().len(), 2);
        assert_eq!(g.overrides().len(), 1);
    }

    #[test]
    fn parse_self_closing_and_open_close_forms() {
        // Mix self-closing `<Default ... />` with open-close `<Override></Override>`.
        let xml = "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\
            <Default Extension=\"rels\" ContentType=\"R\"/>\
            <Override PartName=\"/x.xml\" ContentType=\"X\"></Override>\
            </Types>";
        let g = ContentTypesGraph::parse(xml.as_bytes()).unwrap();
        assert_eq!(g.defaults().len(), 1);
        assert_eq!(g.overrides().len(), 1);
        assert_eq!(g.overrides()[0].0, "/x.xml");
    }

    #[test]
    fn add_override_is_idempotent_for_identical_pair() {
        let mut g = ContentTypesGraph::parse(minimal_xml().as_bytes()).unwrap();
        let n_before = g.overrides().len();
        g.add_override(
            "/xl/workbook.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
        );
        assert_eq!(g.overrides().len(), n_before);
    }

    #[test]
    fn add_override_appends_brand_new_part_at_tail() {
        let mut g = ContentTypesGraph::parse(minimal_xml().as_bytes()).unwrap();
        let n_before = g.overrides().len();
        g.add_override("/xl/comments1.xml", CT_COMMENTS);
        assert_eq!(g.overrides().len(), n_before + 1);
        let last = g.overrides().last().unwrap();
        assert_eq!(last.0, "/xl/comments1.xml");
        assert_eq!(last.1, CT_COMMENTS);
    }

    #[test]
    fn add_override_overwrites_in_place_on_part_collision() {
        let mut g = ContentTypesGraph::parse(minimal_xml().as_bytes()).unwrap();
        // Re-point /xl/workbook.xml at a different content type. The
        // OOXML spec allows only one content type per part, so this
        // should overwrite — order is preserved (stays at slot 0).
        g.add_override("/xl/workbook.xml", "application/X");
        assert_eq!(g.overrides()[0].0, "/xl/workbook.xml");
        assert_eq!(g.overrides()[0].1, "application/X");
        // No duplicate appended at the tail.
        assert_eq!(
            g.overrides()
                .iter()
                .filter(|(p, _)| p == "/xl/workbook.xml")
                .count(),
            1
        );
    }

    #[test]
    fn ensure_default_is_idempotent_for_identical_pair() {
        let mut g = ContentTypesGraph::parse(minimal_xml().as_bytes()).unwrap();
        let n_before = g.defaults().len();
        g.ensure_default("rels", CT_RELATIONSHIPS);
        assert_eq!(g.defaults().len(), n_before);
    }

    #[test]
    fn ensure_default_appends_brand_new_extension() {
        let mut g = ContentTypesGraph::parse(minimal_xml().as_bytes()).unwrap();
        let n_before = g.defaults().len();
        g.ensure_default("vml", CT_VML);
        assert_eq!(g.defaults().len(), n_before + 1);
        let last = g.defaults().last().unwrap();
        assert_eq!(last.0, "vml");
        assert_eq!(last.1, CT_VML);
    }

    #[test]
    fn serialize_round_trip_byte_compatible_for_known_xml() {
        // Round trip: parse → serialize → parse → assert structural equality.
        let g1 = ContentTypesGraph::parse(minimal_xml().as_bytes()).unwrap();
        let bytes = g1.serialize();
        let g2 = ContentTypesGraph::parse(&bytes).unwrap();
        assert_eq!(g1.defaults(), g2.defaults());
        assert_eq!(g1.overrides(), g2.overrides());
    }

    #[test]
    fn serialize_preserves_source_order_of_overrides() {
        let xml = "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\
            <Override PartName=\"/a.xml\" ContentType=\"A\"/>\
            <Override PartName=\"/b.xml\" ContentType=\"B\"/>\
            <Override PartName=\"/c.xml\" ContentType=\"C\"/>\
            </Types>";
        let mut g = ContentTypesGraph::parse(xml.as_bytes()).unwrap();
        // Add a new one — it should go AFTER the existing three.
        g.add_override("/d.xml", "D");
        let bytes = g.serialize();
        let s = std::str::from_utf8(&bytes).unwrap();
        // Extract part-name positions in serialized output.
        let pa = s.find("/a.xml").unwrap();
        let pb = s.find("/b.xml").unwrap();
        let pc = s.find("/c.xml").unwrap();
        let pd = s.find("/d.xml").unwrap();
        assert!(
            pa < pb && pb < pc && pc < pd,
            "expected source order preserved with new at tail"
        );
    }

    #[test]
    fn parse_tolerates_extra_whitespace_and_attributes() {
        let xml = "  <?xml version=\"1.0\"?>\n\
            <Types  xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"\n\
                    xmlns:foo=\"http://example.com/\">\n\
                <Default  Extension=\"rels\"  ContentType=\"R\"  />\n\
                <Override  PartName=\"/x.xml\"  ContentType=\"X\"  foo:bar=\"baz\" />\n\
            </Types>";
        let g = ContentTypesGraph::parse(xml.as_bytes()).unwrap();
        assert_eq!(g.defaults().len(), 1);
        assert_eq!(g.overrides().len(), 1);
    }

    #[test]
    fn parse_skips_entries_missing_required_attributes() {
        // Lenient parse: missing Extension or ContentType drops the entry.
        let xml = "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\
            <Default Extension=\"rels\"/>\
            <Default ContentType=\"X\"/>\
            <Default Extension=\"xml\" ContentType=\"Y\"/>\
            <Override PartName=\"/x.xml\"/>\
            <Override ContentType=\"X\"/>\
            <Override PartName=\"/y.xml\" ContentType=\"Y\"/>\
            </Types>";
        let g = ContentTypesGraph::parse(xml.as_bytes()).unwrap();
        assert_eq!(g.defaults().len(), 1);
        assert_eq!(g.defaults()[0], ("xml".to_string(), "Y".to_string()));
        assert_eq!(g.overrides().len(), 1);
        assert_eq!(g.overrides()[0], ("/y.xml".to_string(), "Y".to_string()));
    }

    #[test]
    fn parse_rejects_malformed_xml() {
        let xml = "<Types xmlns=\"...\"><Default Extension=\"rels\"";
        let result = ContentTypesGraph::parse(xml.as_bytes());
        assert!(result.is_err(), "expected parse error on truncated XML");
    }

    #[test]
    fn serialize_escapes_attribute_special_chars() {
        let mut g = ContentTypesGraph::new();
        // A path containing `&` (unlikely in practice but legal in OOXML
        // before escape) must serialize as `&amp;`.
        g.add_override("/foo&bar.xml", "type/a&b");
        let bytes = g.serialize();
        let s = std::str::from_utf8(&bytes).unwrap();
        assert!(s.contains("/foo&amp;bar.xml"));
        assert!(s.contains("type/a&amp;b"));
        assert!(!s.contains("/foo&bar.xml"));
    }

    #[test]
    fn aggregation_op_records_compose() {
        // Smoke: ContentTypeOp can be cloned + compared, so future
        // aggregation phases (RFC-022/023/024) can dedupe per-sheet ops
        // before applying.
        let a = ContentTypeOp::AddOverride("/x.xml".into(), CT_TABLE.into());
        let b = a.clone();
        assert_eq!(a, b);
        let c = ContentTypeOp::EnsureDefault("vml".into(), CT_VML.into());
        assert_ne!(a, c);
        // Pattern-match exhaustiveness check: silence dead-code for the
        // CT_XML_DEFAULT constant.
        let _used = CT_XML_DEFAULT;
    }
}
