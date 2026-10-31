//! `.rels` (Open Packaging Conventions Relationships) parser, mutable graph,
//! and serializer. Shared between the writer (`wolfxl-writer`) and the
//! patcher (`src/wolfxl/`) so the OOXML rels grammar lives in exactly one
//! place.
//!
//! See ECMA-376 Part 1 §15.2 (Open Packaging Conventions — Relationships) and
//! `Plans/rfcs/010-rels-graph.md` for the full design rationale.
//!
//! # Invariants
//!
//! - **`RelId` is opaque.** Allocated ids are returned by [`RelsGraph::add`]
//!   and may be stored in any sibling part as `r:id="…"`. Callers must not
//!   parse the numeric suffix or assume contiguity. Excel does not require
//!   gap-free ids; we exploit that for monotonicity.
//! - **`next_rid` is monotonic and never decreases.** [`RelsGraph::remove`]
//!   does not free an id for re-use. This is the correctness fix over
//!   openpyxl's `f"rId{len(self)}"` allocation: in modify mode the patcher
//!   mixes existing-on-disk ids with freshly-allocated ones, so any id
//!   re-use would silently re-aim an unrelated reference.
//! - **Source order is preserved.** [`RelsGraph::parse`] pushes into a `Vec`
//!   in document order; [`RelsGraph::serialize`] iterates that `Vec`. No
//!   sort. Reordering would break the modify-mode "minimal diff" promise.
//! - **`TargetMode` semantics.** `Internal` (default) means `Target` is a
//!   ZIP-relative URI resolved against the part owning the rels file.
//!   `External` means `Target` is an opaque absolute URI (e.g. `https://`,
//!   `mailto:`); we never normalize, percent-encode, or otherwise rewrite it.
//! - **`find_by_target` returns the first match by source order.** Callers
//!   use this to dedupe — e.g. a hyperlink to the same external URL from
//!   two cells should reuse one rId rather than allocating a fresh one.
//! - **Lenient parser.** Missing `Id`/`Type`/`Target` on a `<Relationship>`
//!   skips that entry silently (matches the legacy reader behavior in
//!   `src/ooxml_util.rs::parse_relationship_targets`); only malformed XML
//!   surfaces as an error.
//! - **Deterministic serialization.** Always emits the canonical preamble,
//!   one continuous body line, fixed attribute order (Id, Type, Target,
//!   [TargetMode]). Required for `WOLFXL_TEST_EPOCH=0` golden files.

use std::fmt;

use quick_xml::events::Event;
use quick_xml::Reader as XmlReader;

pub mod package_guard;

mod part_id_allocator;
pub use part_id_allocator::PartIdAllocator;

mod subgraph_walk;
pub use subgraph_walk::{walk_sheet_subgraph, walk_sheet_subgraph_with_nested, SheetSubgraph};

// ---------------------------------------------------------------------------
// Relationship-type URIs.
//
// Single source of truth for both the writer (`crates/wolfxl-writer`) and the
// patcher (`src/wolfxl/`). New relationship types belong here, not in the
// consumers.
// ---------------------------------------------------------------------------
pub mod rt {
    pub const OFFICE_DOC: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
    pub const CORE_PROPS: &str =
        "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
    pub const EXT_PROPS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
    pub const CUSTOM_PROPS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties";
    pub const WORKSHEET: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
    pub const CHARTSHEET: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";
    pub const STYLES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
    pub const THEME: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
    pub const SHARED_STRINGS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
    pub const HYPERLINK: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
    pub const COMMENTS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
    pub const VML_DRAWING: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";
    pub const DRAWING: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
    pub const CTRL_PROP: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp";
    pub const IMAGE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
    pub const CHART: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
    pub const CHART_STYLE: &str =
        "http://schemas.microsoft.com/office/2011/relationships/chartStyle";
    pub const CHART_COLOR_STYLE: &str =
        "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle";
    pub const TABLE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
    pub const PIVOT_TABLE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable";
    pub const PIVOT_CACHE_DEF: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition";
    pub const OLE_OBJECT: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject";
    pub const VBA_PROJECT: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vbaProject";
    pub const PRINTER_SETTINGS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings";

    // RFC-071 / G18 — external workbook links.
    pub const EXTERNAL_LINK: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink";
    pub const EXTERNAL_LINK_PATH: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath";
    pub const MS_EXTERNAL_LINK_PATH_PREFIX: &str =
        "http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/";

    // RFC-068 / G08 — threaded comments. Microsoft's `/2018/...` namespace.
    pub const THREADED_COMMENTS: &str =
        "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment";
    pub const PERSON_LIST: &str =
        "http://schemas.microsoft.com/office/2017/10/relationships/person";
}

const RELS_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";

// ---------------------------------------------------------------------------
// RelId — newtype around the "rId<N>" string.
// ---------------------------------------------------------------------------

/// Relationship Id (e.g. `rId7`). A newtype so we cannot accidentally mix it
/// with a `Target` string at a call site.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct RelId(pub String);

impl RelId {
    /// Parse the numeric suffix of an `rId<N>` string. Returns `None` if the
    /// id does not follow the convention. Some legacy files (PowerPoint
    /// templates) use other id schemes; we preserve them verbatim but cannot
    /// allocate fresh ids alongside them.
    pub fn numeric_suffix(&self) -> Option<u32> {
        let n = self.0.strip_prefix("rId")?;
        if n.is_empty() {
            return None;
        }
        n.parse::<u32>().ok()
    }
}

impl fmt::Display for RelId {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(&self.0)
    }
}

impl From<RelId> for String {
    fn from(id: RelId) -> String {
        id.0
    }
}

// ---------------------------------------------------------------------------
// TargetMode + Relationship.
// ---------------------------------------------------------------------------

/// `Internal` is the default — `Target` is resolved as a ZIP-relative URI
/// against the part owning the rels file. `External` is set explicitly for
/// hyperlinks, oleObject links, etc., and `Target` is treated as an opaque
/// absolute URI (e.g. `https://...`, `mailto:...`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub enum TargetMode {
    Internal,
    External,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Relationship {
    pub id: RelId,
    pub rel_type: String,
    pub target: String,
    pub mode: TargetMode,
}

// ---------------------------------------------------------------------------
// RelsGraph — parsed + mutable view of a single `*.rels` file.
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct RelsGraph {
    rels: Vec<Relationship>,
    next_rid: u32,
}

impl RelsGraph {
    /// Empty graph. Used when synthesizing a brand-new rels file.
    pub fn new() -> Self {
        Self {
            rels: Vec::new(),
            next_rid: 1,
        }
    }

    /// Parse an existing rels file. Empty input yields an empty graph (the
    /// `next_rid` counter starts at 1).
    ///
    /// Tolerant of:
    /// - missing XML declaration / BOM
    /// - extra attributes on `<Relationships>` or `<Relationship>`
    /// - non-`rId<N>` ids (preserved verbatim, but they don't bump
    ///   `next_rid`)
    ///
    /// Skips (does not error on) `<Relationship>` entries missing required
    /// attributes — this matches the legacy reader behavior in
    /// `src/ooxml_util.rs::parse_relationship_targets`.
    pub fn parse(xml: &[u8]) -> Result<Self, String> {
        let mut reader = XmlReader::from_reader(xml);
        reader.config_mut().trim_text(true);
        let mut buf: Vec<u8> = Vec::new();
        let mut rels: Vec<Relationship> = Vec::new();
        let mut max_seen: u32 = 0;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    if e.local_name().as_ref() != "Relationship" {
                        buf.clear();
                        continue;
                    }
                    let mut id_str: Option<String> = None;
                    let mut rel_type: Option<String> = None;
                    let mut target: Option<String> = None;
                    let mut mode = TargetMode::Internal;

                    for a in e.attributes().with_checks(false).flatten() {
                        let key = a.key.as_ref();
                        let value = match quick_xml::escape::unescape(&a.value) {
                            Ok(v) => v.into_owned(),
                            Err(_) => a.value.to_string(),
                        };
                        match key {
                            "Id" => id_str = Some(value),
                            "Type" => rel_type = Some(value),
                            "Target" => target = Some(value),
                            "TargetMode" => {
                                if value == "External" {
                                    mode = TargetMode::External;
                                }
                            }
                            _ => {}
                        }
                    }

                    let (Some(id_str), Some(rel_type), Some(target)) = (id_str, rel_type, target)
                    else {
                        // Skip malformed relationships — match legacy reader.
                        buf.clear();
                        continue;
                    };

                    let id = RelId(id_str);
                    if let Some(n) = id.numeric_suffix() {
                        if n > max_seen {
                            max_seen = n;
                        }
                    }
                    rels.push(Relationship {
                        id,
                        rel_type,
                        target,
                        mode,
                    });
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(format!("XML parse error: {e}")),
                _ => {}
            }
            buf.clear();
        }

        Ok(RelsGraph {
            rels,
            next_rid: max_seen.saturating_add(1).max(1),
        })
    }

    /// Append a new relationship, allocating a fresh monotonic `rId`. Returns
    /// the allocated id. The counter never decreases — `remove` does not free
    /// an id for re-use. This is the correctness fix over openpyxl's
    /// `f"rId{len(self)}"` allocation (see RFC-010 §3 bullet 2).
    pub fn add(&mut self, rel_type: &str, target: &str, mode: TargetMode) -> RelId {
        let id = RelId(format!("rId{}", self.next_rid));
        self.next_rid += 1;
        self.rels.push(Relationship {
            id: id.clone(),
            rel_type: rel_type.to_string(),
            target: target.to_string(),
            mode,
        });
        id
    }

    /// Append a relationship with an explicit id. Used when migrating
    /// existing-on-disk rels graphs where the caller already owns the id
    /// space (e.g. rebuilding `xl/_rels/workbook.xml.rels`).
    ///
    /// Panics if `id` is already present in the graph.
    pub fn add_with_id(&mut self, id: RelId, rel_type: &str, target: &str, mode: TargetMode) {
        if self.rels.iter().any(|r| r.id == id) {
            panic!("RelsGraph::add_with_id: id {id} already present");
        }
        if let Some(n) = id.numeric_suffix() {
            if n >= self.next_rid {
                self.next_rid = n + 1;
            }
        }
        self.rels.push(Relationship {
            id,
            rel_type: rel_type.to_string(),
            target: target.to_string(),
            mode,
        });
    }

    /// Remove by id. No-op if absent. Does not renumber siblings, does not
    /// free the id for re-use.
    pub fn remove(&mut self, id: &RelId) {
        self.rels.retain(|r| &r.id != id);
    }

    /// Look up by id. O(N); rels files almost always have <50 entries.
    pub fn get(&self, id: &RelId) -> Option<&Relationship> {
        self.rels.iter().find(|r| &r.id == id)
    }

    /// Return all relationships of the given type, in source order.
    pub fn find_by_type(&self, rel_type: &str) -> Vec<&Relationship> {
        self.rels
            .iter()
            .filter(|r| r.rel_type == rel_type)
            .collect()
    }

    /// First match by `(target, mode)`, in source order. Used by callers to
    /// dedupe — e.g. a hyperlink to the same external URL from two cells
    /// should reuse one rId.
    pub fn find_by_target(&self, target: &str, mode: TargetMode) -> Option<&Relationship> {
        self.rels
            .iter()
            .find(|r| r.target == target && r.mode == mode)
    }

    /// Iterate all relationships in source order.
    pub fn iter(&self) -> std::slice::Iter<'_, Relationship> {
        self.rels.iter()
    }

    /// Number of relationships.
    pub fn len(&self) -> usize {
        self.rels.len()
    }

    /// True if no relationships have been added.
    pub fn is_empty(&self) -> bool {
        self.rels.is_empty()
    }

    /// Serialize back to canonical bytes:
    ///
    /// ```xml
    /// <?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n
    /// <Relationships xmlns="…/relationships">
    ///   <Relationship Id="…" Type="…" Target="…" [TargetMode="External"]/>
    ///   …
    /// </Relationships>
    /// ```
    ///
    /// Always emits attributes in this canonical order (Id, Type, Target,
    /// [TargetMode]), one continuous line for the body, no inter-element
    /// whitespace. Determinism is required for `WOLFXL_TEST_EPOCH=0` golden
    /// files and for byte-stable round-trips.
    pub fn serialize(&self) -> Vec<u8> {
        let mut out = String::with_capacity(256 + self.rels.len() * 220);
        out.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n");
        out.push_str(&format!("<Relationships xmlns=\"{RELS_NS}\">"));
        for r in &self.rels {
            out.push_str("<Relationship Id=\"");
            push_xml_attr_escape(&mut out, &r.id.0);
            out.push_str("\" Type=\"");
            // Type values come from the rt:: constants (or otherwise from a
            // file we just parsed); they never contain attribute-special
            // characters.
            out.push_str(&r.rel_type);
            out.push_str("\" Target=\"");
            push_xml_attr_escape(&mut out, &r.target);
            if r.mode == TargetMode::External {
                out.push_str("\" TargetMode=\"External");
            }
            out.push_str("\"/>");
        }
        out.push_str("</Relationships>");
        out.into_bytes()
    }
}

/// XML attribute-value escape — `Target` values can be URLs containing `&`,
/// `<`, etc., which would otherwise break the surrounding attribute.
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
// rels_path_for — the conventional path mapping.
// ---------------------------------------------------------------------------

/// Compute the conventional rels path for a part:
///
/// | Input                          | Output                                  |
/// |--------------------------------|-----------------------------------------|
/// | `xl/workbook.xml`              | `Some("xl/_rels/workbook.xml.rels")`    |
/// | `xl/worksheets/sheet1.xml`     | `Some("xl/worksheets/_rels/sheet1.xml.rels")` |
/// | `[Content_Types].xml`          | `None` (root content-types has no rels) |
/// | `` (empty)                     | `None`                                  |
pub fn rels_path_for(part_path: &str) -> Option<String> {
    if part_path.is_empty() || part_path == "[Content_Types].xml" {
        return None;
    }
    match part_path.rfind('/') {
        Some(idx) => {
            let dir = &part_path[..idx];
            let file = &part_path[idx + 1..];
            Some(format!("{dir}/_rels/{file}.rels"))
        }
        None => Some(format!("_rels/{part_path}.rels")),
    }
}

/// Resolve a relative or absolute OOXML relationship target against a base
/// directory.
///
/// `base_dir` is the directory containing the source part (for example
/// `xl/worksheets` for `xl/worksheets/sheet1.xml`). Absolute package targets
/// beginning with `/` resolve from the ZIP root.
pub fn resolve_target(base_dir: &str, target: &str) -> String {
    let (mut parts, target_iter): (Vec<&str>, _) = if let Some(stripped) = target.strip_prefix('/')
    {
        (Vec::new(), stripped.split('/'))
    } else {
        (
            base_dir.split('/').filter(|p| !p.is_empty()).collect(),
            target.split('/'),
        )
    };
    for seg in target_iter {
        match seg {
            "" | "." => {}
            ".." => {
                parts.pop();
            }
            other => parts.push(other),
        }
    }
    parts.join("/")
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;

    const MINIMAL_ROOT_RELS: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>"#;

    const HYPERLINK_RELS: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/docs" TargetMode="External" Id="rId1"/><Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="mailto:test@example.com" TargetMode="External" Id="rId2"/><Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/search?q=excel%20bench&amp;sort=desc#section-2" TargetMode="External" Id="rId3"/></Relationships>"#;

    fn workbook_rels_with_n_sheets(n: usize) -> Vec<u8> {
        let mut s = String::from(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
        );
        for i in 1..=n {
            s.push_str(&format!(
                r#"<Relationship Id="rId{i}" Type="{}" Target="worksheets/sheet{i}.xml"/>"#,
                rt::WORKSHEET
            ));
        }
        s.push_str(&format!(
            r#"<Relationship Id="rId{}" Type="{}" Target="styles.xml"/>"#,
            n + 1,
            rt::STYLES
        ));
        s.push_str(&format!(
            r#"<Relationship Id="rId{}" Type="{}" Target="sharedStrings.xml"/>"#,
            n + 2,
            rt::SHARED_STRINGS
        ));
        s.push_str("</Relationships>");
        s.into_bytes()
    }

    #[test]
    fn parse_minimal_root() {
        let g = RelsGraph::parse(MINIMAL_ROOT_RELS).expect("parse");
        assert_eq!(g.len(), 3);
        assert_eq!(g.iter().next().unwrap().id, RelId("rId1".into()));
        assert_eq!(
            g.find_by_type(rt::OFFICE_DOC).len(),
            1,
            "exactly one officeDocument relationship"
        );
        assert_eq!(g.find_by_type(rt::CORE_PROPS).len(), 1);
        assert_eq!(g.find_by_type(rt::EXT_PROPS).len(), 1);
    }

    #[test]
    fn parse_workbook_rels_with_styles_and_sst() {
        let bytes = workbook_rels_with_n_sheets(3);
        let g = RelsGraph::parse(&bytes).expect("parse");
        assert_eq!(
            g.find_by_type(rt::WORKSHEET).len(),
            3,
            "one worksheet relationship per sheet"
        );
        assert_eq!(g.find_by_type(rt::STYLES).len(), 1);
        assert_eq!(g.find_by_type(rt::SHARED_STRINGS).len(), 1);
    }

    #[test]
    fn parse_external_hyperlink_marks_external() {
        let g = RelsGraph::parse(HYPERLINK_RELS).expect("parse");
        assert_eq!(g.len(), 3);
        for r in g.iter() {
            assert_eq!(r.mode, TargetMode::External);
            assert_eq!(r.rel_type, rt::HYPERLINK);
        }
        // Ampersand in the third URL must round-trip unescaped in memory.
        let third = g.iter().nth(2).unwrap();
        assert!(
            third.target.contains("&sort=desc"),
            "ampersand must be unescaped after parse, got: {}",
            third.target
        );
    }

    #[test]
    fn add_after_parse_uses_strictly_greater_rid() {
        let mut g = RelsGraph::parse(&workbook_rels_with_n_sheets(3)).expect("parse");
        // Existing ids are rId1..rId5. Next allocations must be rId6, rId7.
        let id_a = g.add(rt::HYPERLINK, "https://a", TargetMode::External);
        let id_b = g.add(rt::HYPERLINK, "https://b", TargetMode::External);
        assert_eq!(id_a, RelId("rId6".into()));
        assert_eq!(id_b, RelId("rId7".into()));
    }

    #[test]
    fn remove_then_add_does_not_collide() {
        // Demonstrates the openpyxl bug fix: removing rId2 then adding must
        // NOT reuse rId2.
        let bytes = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#;
        let mut g = RelsGraph::parse(bytes).expect("parse");
        g.remove(&RelId("rId2".into()));
        let new_id = g.add(rt::HYPERLINK, "https://x", TargetMode::External);
        assert_eq!(
            new_id,
            RelId("rId4".into()),
            "monotonic counter must not reuse rId2"
        );
    }

    #[test]
    fn find_by_target_dedupe() {
        let mut g = RelsGraph::new();
        let id1 = g.add(rt::HYPERLINK, "https://example.com", TargetMode::External);
        // A caller dedupes by find_by_target before allocating.
        let found = g.find_by_target("https://example.com", TargetMode::External);
        assert_eq!(found.map(|r| r.id.clone()), Some(id1));
        // External target with same URL but Internal mode should NOT match.
        assert!(g
            .find_by_target("https://example.com", TargetMode::Internal)
            .is_none());
    }

    #[test]
    fn serialize_round_trips_idempotent() {
        // For every fixture-like input, serialize(parse(x)) is structurally
        // equal to parse(x), and parse(serialize(parse(x))) bytes match.
        for &input in &[MINIMAL_ROOT_RELS, HYPERLINK_RELS] {
            let g1 = RelsGraph::parse(input).expect("parse 1");
            let out1 = g1.serialize();
            let g2 = RelsGraph::parse(&out1).expect("parse 2");
            assert_eq!(
                g1, g2,
                "parse(serialize(parse(input))) must equal parse(input)"
            );
            let out2 = g2.serialize();
            assert_eq!(
                out1, out2,
                "serialize is byte-stable on a fixed-point graph"
            );
        }
    }

    #[test]
    fn serialize_external_xml_escape() {
        let mut g = RelsGraph::new();
        g.add(
            rt::HYPERLINK,
            "https://example.com/path?q=1&r=2",
            TargetMode::External,
        );
        let bytes = g.serialize();
        let text = std::str::from_utf8(&bytes).expect("utf8");
        assert!(text.contains("q=1&amp;r=2"), "ampersand must be escaped");
        assert!(text.contains("TargetMode=\"External\""));
        // Round-trip preserves the unescaped value in memory.
        let g2 = RelsGraph::parse(&bytes).expect("re-parse");
        assert_eq!(
            g2.iter().next().unwrap().target,
            "https://example.com/path?q=1&r=2"
        );
    }

    #[test]
    fn numeric_suffix_handles_legacy_ids() {
        let bytes = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="myCustomId123" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>"#;
        let mut g = RelsGraph::parse(bytes).expect("parse");
        assert_eq!(g.len(), 1);
        assert_eq!(g.iter().next().unwrap().id.numeric_suffix(), None);
        // Allocating after a legacy-only file starts at rId1 (counter never
        // bumped above 0 by the legacy id).
        let new_id = g.add(rt::HYPERLINK, "https://x", TargetMode::External);
        assert_eq!(new_id, RelId("rId1".into()));
    }

    #[test]
    fn rels_path_for_workbook() {
        assert_eq!(
            rels_path_for("xl/workbook.xml"),
            Some("xl/_rels/workbook.xml.rels".into())
        );
    }

    #[test]
    fn rels_path_for_sheet() {
        assert_eq!(
            rels_path_for("xl/worksheets/sheet1.xml"),
            Some("xl/worksheets/_rels/sheet1.xml.rels".into())
        );
    }

    #[test]
    fn rels_path_for_top_level_returns_root_rels() {
        // A bare path with no '/' is a root-level part; its rels lives at
        // "_rels/<name>.rels".
        assert_eq!(
            rels_path_for("workbook.xml"),
            Some("_rels/workbook.xml.rels".into())
        );
    }

    #[test]
    fn rels_path_for_content_types_returns_none() {
        assert_eq!(rels_path_for("[Content_Types].xml"), None);
        assert_eq!(rels_path_for(""), None);
    }

    #[test]
    fn resolve_target_normalizes_relative_paths() {
        assert_eq!(
            resolve_target("xl/worksheets", "../drawings/drawing1.xml"),
            "xl/drawings/drawing1.xml"
        );
        assert_eq!(
            resolve_target("xl/drawings", "../charts/chart1.xml"),
            "xl/charts/chart1.xml"
        );
        assert_eq!(
            resolve_target("xl/drawings", "/xl/media/image1.png"),
            "xl/media/image1.png"
        );
    }

    #[test]
    fn add_with_id_panics_on_duplicate() {
        let mut g = RelsGraph::new();
        g.add_with_id(
            RelId("rId7".into()),
            rt::HYPERLINK,
            "https://a",
            TargetMode::External,
        );
        let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            g.add_with_id(
                RelId("rId7".into()),
                rt::HYPERLINK,
                "https://b",
                TargetMode::External,
            );
        }));
        assert!(result.is_err(), "duplicate add_with_id must panic");
    }

    #[test]
    fn add_with_id_bumps_counter() {
        // Ensure subsequent add() doesn't collide with an explicit add_with_id.
        let mut g = RelsGraph::new();
        g.add_with_id(
            RelId("rId10".into()),
            rt::WORKSHEET,
            "worksheets/sheet1.xml",
            TargetMode::Internal,
        );
        let next = g.add(rt::STYLES, "styles.xml", TargetMode::Internal);
        assert_eq!(next, RelId("rId11".into()));
    }
}
