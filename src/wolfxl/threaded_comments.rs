//! Threaded-comments + persons extract / merge / emit for the modify-mode
//! patcher (RFC-068 / G08 step 5).
//!
//! Used by `XlsxPatcher::do_save`'s threaded-comments phase to:
//! 1. Read existing `xl/threadedComments/threadedCommentsN.xml` parts
//!    referenced from sheet rels — preserving threads on cells the user
//!    didn't touch.
//! 2. Read the workbook-scope `xl/persons/personList.xml` — preserving
//!    GUIDs for round-trip stability.
//! 3. Merge user-supplied `ThreadedCommentOp::Set / Delete` per cell and
//!    `PersonOp::Add` (idempotent on GUID).
//! 4. Re-emit fresh `threadedCommentsN.xml` and `personList.xml` byte
//!    streams.
//! 5. Mutate the sheet's / workbook's `RelsGraph` for added or removed
//!    parts.
//! 6. Synthesize legacy-comment placeholders (`tc={topId}` author, body
//!    `[Threaded comment]`) so the existing comments phase emits them
//!    into `commentsN.xml` and the file remains readable by Excel <365.

use std::collections::BTreeMap;

use quick_xml::events::Event;
use quick_xml::Reader as XmlReader;

use wolfxl_rels::{rt, RelId, RelsGraph, TargetMode};

// ---------------------------------------------------------------------------
// Public types
// ---------------------------------------------------------------------------

/// One full top-level threaded comment (with replies) the user is
/// queueing onto a cell.
///
/// The patcher treats `(sheet, cell_ref)` as the merge key — running
/// `Set` replaces ALL existing threads on that cell with this one. To
/// keep multiple existing threads on the same cell intact while adding
/// a new top-level thread, callers must read the existing payload first
/// and re-queue it (modify mode does not currently expose that path).
#[derive(Debug, Clone, PartialEq)]
pub struct ThreadedCommentPatch {
    pub top: ThreadedCommentEntry,
    pub replies: Vec<ThreadedCommentEntry>,
}

/// One serialized `<threadedComment>` row, flat in the OOXML sense.
#[derive(Debug, Clone, PartialEq)]
pub struct ThreadedCommentEntry {
    pub id: String,
    pub cell_ref: String,
    pub person_id: String,
    /// ISO-8601 with millisecond precision. Allocator (Python flush
    /// layer) is responsible for filling this in at queue time.
    pub created: String,
    /// `Some(topId)` for replies; `None` for the top of a thread.
    pub parent_id: Option<String>,
    pub text: String,
    pub done: bool,
}

/// `Set` queues a new thread at `(sheet, cell)`; `Delete` drops every
/// existing thread at that cell (top + replies).
#[derive(Debug, Clone, PartialEq)]
pub enum ThreadedCommentOp {
    Set(ThreadedCommentPatch),
    Delete,
}

/// One workbook-scope person entry queued for the personList. The patcher
/// is idempotent on `id`: adding the same GUID twice is a no-op.
#[derive(Debug, Clone, PartialEq)]
pub struct PersonPatch {
    pub id: String,
    pub display_name: String,
    pub user_id: String,
    pub provider_id: String,
}

/// One existing `<threadedComment>` row read off disk. Same shape as
/// [`ThreadedCommentEntry`] but kept distinct so mergers can tell at-a-
/// glance whether they're operating on parsed-from-disk or queued-by-
/// user state.
#[derive(Debug, Clone, PartialEq)]
pub struct ExistingThreadedComment {
    pub id: String,
    pub cell_ref: String,
    pub person_id: String,
    pub created: String,
    pub parent_id: Option<String>,
    pub text: String,
    pub done: bool,
}

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

pub const NS_THREADED: &str =
    "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments";
pub const CT_THREADED: &str = "application/vnd.ms-excel.threadedcomments+xml";
pub const CT_PERSON_LIST: &str = "application/vnd.ms-excel.person+xml";

// ---------------------------------------------------------------------------
// Extract: threadedCommentsN.xml
// ---------------------------------------------------------------------------

/// Parse a `threadedCommentsN.xml` byte stream into a flat ordered list.
pub fn extract_threaded_comments(xml: &[u8]) -> Vec<ExistingThreadedComment> {
    if xml.is_empty() {
        return Vec::new();
    }
    let text = match std::str::from_utf8(xml) {
        Ok(s) => s,
        Err(_) => return Vec::new(),
    };
    let mut reader = XmlReader::from_str(text);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut out: Vec<ExistingThreadedComment> = Vec::new();
    let mut current: Option<ExistingThreadedComment> = None;
    let mut in_text = false;
    let mut text_buf = String::new();
    let mut node = crate::ooxml_util::XmlTextNode::default();

    fn build_from_attrs(e: &quick_xml::events::BytesStart<'_>) -> ExistingThreadedComment {
        let id = attr(e, b"id").unwrap_or_default();
        let cell_ref = attr(e, b"ref").unwrap_or_default();
        let person_id = attr(e, b"personId").unwrap_or_default();
        let created = attr(e, b"dT").unwrap_or_default();
        let parent_id = attr(e, b"parentId");
        let done = attr(e, b"done")
            .map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
            .unwrap_or(false);
        ExistingThreadedComment {
            id,
            cell_ref,
            person_id,
            created,
            parent_id,
            text: String::new(),
            done,
        }
    }

    loop {
        let event = reader.read_event_into(&mut buf);
        if !matches!(event, Ok(Event::Text(_)) | Ok(Event::GeneralRef(_))) {
            if in_text {
                if let Ok(text) = quick_xml::escape::unescape(node.raw()) {
                    text_buf.push_str(&text);
                }
            }
            node.clear();
        }
        match event {
            Ok(Event::Start(e)) => match e.local_name().as_ref().as_bytes() {
                b"threadedComment" => {
                    current = Some(build_from_attrs(&e));
                    text_buf.clear();
                }
                b"text" => {
                    in_text = true;
                    text_buf.clear();
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => {
                if e.local_name().as_ref() == "threadedComment" {
                    out.push(build_from_attrs(&e));
                }
            }
            Ok(Event::End(e)) => match e.local_name().as_ref().as_bytes() {
                b"threadedComment" => {
                    if let Some(mut tc) = current.take() {
                        tc.text = std::mem::take(&mut text_buf);
                        out.push(tc);
                    }
                }
                b"text" => {
                    in_text = false;
                }
                _ => {}
            },
            Ok(Event::Text(e)) if in_text => node.push_text(&e),
            Ok(Event::GeneralRef(e)) if in_text => node.push_ref(&e),
            Ok(Event::CData(e)) => {
                if in_text {
                    text_buf.push_str(&e);
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    out
}

/// Parse a `personList.xml` byte stream into an ordered list.
pub fn extract_persons(xml: &[u8]) -> Vec<PersonPatch> {
    if xml.is_empty() {
        return Vec::new();
    }
    let text = match std::str::from_utf8(xml) {
        Ok(s) => s,
        Err(_) => return Vec::new(),
    };
    let mut reader = XmlReader::from_str(text);
    let mut buf = Vec::new();
    let mut out = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.local_name().as_ref() == "person" {
                    let id = attr(&e, b"id").unwrap_or_default();
                    if id.is_empty() {
                        continue;
                    }
                    out.push(PersonPatch {
                        id,
                        display_name: attr(&e, b"displayName").unwrap_or_default(),
                        user_id: attr(&e, b"userId").unwrap_or_default(),
                        provider_id: attr(&e, b"providerId").unwrap_or_else(|| "None".to_string()),
                    });
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    out
}

fn attr(e: &quick_xml::events::BytesStart<'_>, name: &[u8]) -> Option<String> {
    for raw in e.attributes().with_checks(false).flatten() {
        if raw.key.as_ref().as_bytes() == name {
            // Unescape without normalizing literal attribute whitespace.
            return quick_xml::escape::unescape(&raw.value).ok().map(|cow| cow.into_owned());
        }
    }
    None
}

// ---------------------------------------------------------------------------
// Merge
// ---------------------------------------------------------------------------

/// Merge `existing` threads with `ops`, keyed by cell coordinate. `Set`
/// at a coord replaces every existing thread at that coord with the
/// queued top + replies. `Delete` drops them. Cells absent from ops
/// pass through verbatim.
///
/// The output order is: existing entries (with edited cells substituted
/// in place; deleted cells removed), then any net-new cells in ops
/// order. This preserves on-disk byte stability for unedited cells.
pub fn merge_threaded_comments(
    existing: Vec<ExistingThreadedComment>,
    ops: &BTreeMap<String, ThreadedCommentOp>,
) -> Vec<ExistingThreadedComment> {
    let mut out: Vec<ExistingThreadedComment> = Vec::with_capacity(existing.len());

    // Track which cells we've already emitted from existing so we can
    // splice the queued replacement in the original position.
    let mut handled_cells: std::collections::HashSet<String> = std::collections::HashSet::new();

    for entry in existing {
        let coord = entry.cell_ref.clone();
        match ops.get(&coord) {
            Some(ThreadedCommentOp::Delete) => {
                // Drop this entry; do not flag handled (subsequent
                // entries on this coord are also dropped).
                continue;
            }
            Some(ThreadedCommentOp::Set(patch)) => {
                if handled_cells.insert(coord.clone()) {
                    push_patch(&mut out, patch);
                }
                // Skip every existing entry on this coord — the queued
                // patch fully owns the cell now.
                continue;
            }
            None => {
                out.push(entry);
            }
        }
    }

    // Net-new sets (cells without any existing entries).
    for (coord, op) in ops {
        if let ThreadedCommentOp::Set(patch) = op {
            if !handled_cells.contains(coord) {
                push_patch(&mut out, patch);
                handled_cells.insert(coord.clone());
            }
        }
    }

    out
}

fn push_patch(out: &mut Vec<ExistingThreadedComment>, patch: &ThreadedCommentPatch) {
    out.push(entry_to_existing(&patch.top));
    for reply in &patch.replies {
        out.push(entry_to_existing(reply));
    }
}

fn entry_to_existing(e: &ThreadedCommentEntry) -> ExistingThreadedComment {
    ExistingThreadedComment {
        id: e.id.clone(),
        cell_ref: e.cell_ref.clone(),
        person_id: e.person_id.clone(),
        created: e.created.clone(),
        parent_id: e.parent_id.clone(),
        text: e.text.clone(),
        done: e.done,
    }
}

/// Merge the existing personList with queued additions. Idempotent on
/// `id`: repeating an already-present GUID is a no-op. New entries
/// preserve queue order.
pub fn merge_persons(existing: Vec<PersonPatch>, ops: &[PersonPatch]) -> Vec<PersonPatch> {
    let mut out = existing;
    for queued in ops {
        if queued.id.is_empty() {
            continue;
        }
        if out.iter().any(|p| p.id == queued.id) {
            continue;
        }
        out.push(queued.clone());
    }
    out
}

// ---------------------------------------------------------------------------
// Build: threadedCommentsN.xml + personList.xml
// ---------------------------------------------------------------------------

pub fn build_threaded_comments_xml(merged: &[ExistingThreadedComment]) -> Vec<u8> {
    if merged.is_empty() {
        return Vec::new();
    }
    let mut out = String::with_capacity(2048 + merged.len() * 96);
    out.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n");
    out.push_str(&format!("<ThreadedComments xmlns=\"{NS_THREADED}\">"));
    for tc in merged {
        out.push_str("<threadedComment ref=\"");
        push_attr_escape(&mut out, &tc.cell_ref);
        out.push_str("\" dT=\"");
        push_attr_escape(&mut out, &tc.created);
        out.push_str("\" personId=\"");
        push_attr_escape(&mut out, &tc.person_id);
        out.push_str("\" id=\"");
        push_attr_escape(&mut out, &tc.id);
        out.push('"');
        if let Some(parent) = &tc.parent_id {
            out.push_str(" parentId=\"");
            push_attr_escape(&mut out, parent);
            out.push('"');
        }
        if tc.done {
            out.push_str(" done=\"1\"");
        }
        out.push_str("><text>");
        push_text_escape(&mut out, &tc.text);
        out.push_str("</text></threadedComment>");
    }
    out.push_str("</ThreadedComments>");
    out.into_bytes()
}

pub fn build_persons_xml(merged: &[PersonPatch]) -> Vec<u8> {
    if merged.is_empty() {
        return Vec::new();
    }
    let mut out = String::with_capacity(512);
    out.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n");
    out.push_str(&format!("<personList xmlns=\"{NS_THREADED}\">"));
    for p in merged {
        out.push_str("<person displayName=\"");
        push_attr_escape(&mut out, &p.display_name);
        out.push_str("\" id=\"");
        push_attr_escape(&mut out, &p.id);
        out.push_str("\" userId=\"");
        push_attr_escape(&mut out, &p.user_id);
        out.push_str("\" providerId=\"");
        push_attr_escape(&mut out, &p.provider_id);
        out.push_str("\"/>");
    }
    out.push_str("</personList>");
    out.into_bytes()
}

fn push_attr_escape(out: &mut String, s: &str) {
    for c in s.chars() {
        match c {
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '&' => out.push_str("&amp;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&apos;"),
            _ => out.push(c),
        }
    }
}

fn push_text_escape(out: &mut String, s: &str) {
    for c in s.chars() {
        match c {
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '&' => out.push_str("&amp;"),
            _ => out.push(c),
        }
    }
}

// ---------------------------------------------------------------------------
// Top-level drivers (called from the patcher phase)
// ---------------------------------------------------------------------------

/// Drive the threaded-comments emit for one sheet. Mirrors the shape of
/// [`super::comments::build_comments`]: takes existing bytes + queued
/// ops + the sheet's rels graph, mutates the rels graph for added or
/// removed parts, returns the fresh byte stream.
pub fn build_threaded_for_sheet(
    existing_xml: Option<&[u8]>,
    ops: &BTreeMap<String, ThreadedCommentOp>,
    rels: &mut RelsGraph,
    threaded_n: u32,
) -> (Vec<u8>, Option<RelId>) {
    let existing = match existing_xml {
        Some(xml) => extract_threaded_comments(xml),
        None => Vec::new(),
    };
    let merged = merge_threaded_comments(existing, ops);

    let target_relative = format!("../threadedComments/threadedComments{}.xml", threaded_n);
    let existing_rid = rels
        .find_by_type(rt::THREADED_COMMENTS)
        .first()
        .map(|r| r.id.clone());

    if merged.is_empty() {
        if let Some(rid) = &existing_rid {
            rels.remove(rid);
        }
        return (Vec::new(), existing_rid);
    }

    let rid = match existing_rid.clone() {
        Some(r) => r,
        None => rels.add(
            rt::THREADED_COMMENTS,
            &target_relative,
            TargetMode::Internal,
        ),
    };
    let bytes = build_threaded_comments_xml(&merged);
    (bytes, Some(rid))
}

/// Drive the personList emit for the workbook. `wb_rels` is the
/// workbook's rels graph; `existing_xml` is the current personList part
/// bytes (None when the file has no personList).
pub fn build_persons_for_workbook(
    existing_xml: Option<&[u8]>,
    ops: &[PersonPatch],
    wb_rels: &mut RelsGraph,
) -> (Vec<u8>, Option<RelId>) {
    let existing = match existing_xml {
        Some(xml) => extract_persons(xml),
        None => Vec::new(),
    };
    let merged = merge_persons(existing, ops);

    let existing_rid = wb_rels
        .find_by_type(rt::PERSON_LIST)
        .first()
        .map(|r| r.id.clone());

    if merged.is_empty() {
        if let Some(rid) = &existing_rid {
            wb_rels.remove(rid);
        }
        return (Vec::new(), existing_rid);
    }

    let rid = match existing_rid.clone() {
        Some(r) => r,
        None => wb_rels.add(
            rt::PERSON_LIST,
            "persons/personList.xml",
            TargetMode::Internal,
        ),
    };
    let bytes = build_persons_xml(&merged);
    (bytes, Some(rid))
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;

    fn entry(id: &str, cell: &str, parent: Option<&str>, text: &str) -> ThreadedCommentEntry {
        ThreadedCommentEntry {
            id: id.to_string(),
            cell_ref: cell.to_string(),
            person_id: "{P}".to_string(),
            created: "2026-05-03T12:00:00.000".to_string(),
            parent_id: parent.map(String::from),
            text: text.to_string(),
            done: false,
        }
    }

    fn existing(id: &str, cell: &str, parent: Option<&str>, text: &str) -> ExistingThreadedComment {
        ExistingThreadedComment {
            id: id.to_string(),
            cell_ref: cell.to_string(),
            person_id: "{P}".to_string(),
            created: "2026-05-03T12:00:00.000".to_string(),
            parent_id: parent.map(String::from),
            text: text.to_string(),
            done: false,
        }
    }

    #[test]
    fn extract_round_trip() {
        let bytes = build_threaded_comments_xml(&[
            existing("{T1}", "A1", None, "hi"),
            existing("{T2}", "A1", Some("{T1}"), "reply"),
        ]);
        let parsed = extract_threaded_comments(&bytes);
        assert_eq!(parsed.len(), 2);
        assert_eq!(parsed[0].text, "hi");
        assert_eq!(parsed[1].parent_id.as_deref(), Some("{T1}"));
    }

    #[test]
    fn extract_persons_round_trip() {
        let bytes = build_persons_xml(&[PersonPatch {
            id: "{A}".into(),
            display_name: "Alice".into(),
            user_id: "alice@x.com".into(),
            provider_id: "AD".into(),
        }]);
        let parsed = extract_persons(&bytes);
        assert_eq!(parsed.len(), 1);
        assert_eq!(parsed[0].id, "{A}");
        assert_eq!(parsed[0].display_name, "Alice");
        assert_eq!(parsed[0].user_id, "alice@x.com");
    }

    #[test]
    fn merge_set_replaces_thread_at_cell() {
        let existing = vec![
            existing("{OLD}", "A1", None, "old"),
            existing("{KEEP}", "B1", None, "untouched"),
        ];
        let mut ops = BTreeMap::new();
        ops.insert(
            "A1".to_string(),
            ThreadedCommentOp::Set(ThreadedCommentPatch {
                top: entry("{NEW}", "A1", None, "new"),
                replies: Vec::new(),
            }),
        );
        let merged = merge_threaded_comments(existing, &ops);
        assert_eq!(merged.len(), 2);
        assert_eq!(merged[0].text, "new");
        assert_eq!(merged[0].cell_ref, "A1");
        assert_eq!(merged[1].text, "untouched");
        assert_eq!(merged[1].cell_ref, "B1");
    }

    #[test]
    fn merge_delete_drops_all_entries_on_cell() {
        let existing = vec![
            existing("{T1}", "A1", None, "top"),
            existing("{T2}", "A1", Some("{T1}"), "reply"),
            existing("{T3}", "B1", None, "keep"),
        ];
        let mut ops = BTreeMap::new();
        ops.insert("A1".to_string(), ThreadedCommentOp::Delete);
        let merged = merge_threaded_comments(existing, &ops);
        assert_eq!(merged.len(), 1);
        assert_eq!(merged[0].cell_ref, "B1");
    }

    #[test]
    fn merge_set_appends_when_cell_is_new() {
        let mut ops = BTreeMap::new();
        ops.insert(
            "C3".to_string(),
            ThreadedCommentOp::Set(ThreadedCommentPatch {
                top: entry("{N}", "C3", None, "fresh"),
                replies: vec![entry("{R}", "C3", Some("{N}"), "child")],
            }),
        );
        let merged = merge_threaded_comments(Vec::new(), &ops);
        assert_eq!(merged.len(), 2);
        assert_eq!(merged[0].text, "fresh");
        assert_eq!(merged[1].text, "child");
    }

    #[test]
    fn persons_merge_idempotent_on_id() {
        let existing = vec![PersonPatch {
            id: "{A}".into(),
            display_name: "Alice".into(),
            user_id: "alice@x.com".into(),
            provider_id: "AD".into(),
        }];
        let queued = vec![
            // Same GUID — must be a no-op even if other fields differ.
            PersonPatch {
                id: "{A}".into(),
                display_name: "Alice REPLACED".into(),
                user_id: String::new(),
                provider_id: "None".into(),
            },
            PersonPatch {
                id: "{B}".into(),
                display_name: "Bob".into(),
                user_id: String::new(),
                provider_id: "None".into(),
            },
        ];
        let merged = merge_persons(existing, &queued);
        assert_eq!(merged.len(), 2);
        assert_eq!(merged[0].display_name, "Alice");
        assert_eq!(merged[1].display_name, "Bob");
    }

    #[test]
    fn build_threaded_xml_escapes_text() {
        let merged = vec![existing("{T}", "A1", None, "<b>&\"hi\"")];
        let bytes = build_threaded_comments_xml(&merged);
        let text = String::from_utf8(bytes).unwrap();
        assert!(text.contains("&lt;b&gt;&amp;\"hi\""));
    }

    #[test]
    fn build_persons_xml_escapes_attrs() {
        let merged = vec![PersonPatch {
            id: "{A}".into(),
            display_name: "R&D \"Team\"".into(),
            user_id: String::new(),
            provider_id: "None".into(),
        }];
        let bytes = build_persons_xml(&merged);
        let text = String::from_utf8(bytes).unwrap();
        assert!(text.contains("displayName=\"R&amp;D &quot;Team&quot;\""));
    }

    #[test]
    fn split_references_preserve_text_and_cdata_boundaries() {
        let xml = b"<ThreadedComments><threadedComment id=\"T\" ref=\"A1\">\
                    <text> A &amp; B&#10; <![CDATA[<&raw>]]>&#13; tail </text>\
                    </threadedComment></ThreadedComments>";
        let comments = extract_threaded_comments(xml);
        assert_eq!(comments.len(), 1);
        assert_eq!(comments[0].text, " A & B\n <&raw>\r tail ");
    }

    #[test]
    fn invalid_reference_omits_whole_text_node_not_adjacent_cdata() {
        let xml = b"<ThreadedComments><threadedComment id=\"T\" ref=\"A1\">\
                    <text>before &unknown; after<![CDATA[kept]]></text>\
                    </threadedComment></ThreadedComments>";
        let comments = extract_threaded_comments(xml);
        assert_eq!(comments.len(), 1);
        assert_eq!(comments[0].text, "kept");
    }

    #[test]
    fn person_attributes_preserve_literal_whitespace_and_references() {
        let persons = extract_persons(
            b"<personList><person id=\"P\" displayName=\" A\t&amp;\r\nB&#10; \" /></personList>",
        );
        assert_eq!(persons.len(), 1);
        assert_eq!(persons[0].display_name, " A\t&\r\nB\n ");
    }
}
