//! `xl/workbook.xml` `<workbookProtection>` + `<fileSharing>` splicer
//! for modify mode (RFC-058 Phase 2.5q).
//!
//! Used by `XlsxPatcher::do_save`'s Phase 2.5q to splice the two
//! security elements into `xl/workbook.xml` at canonical
//! CT_Workbook child positions:
//!
//! ```text
//!   fileVersion → fileSharing → workbookPr → workbookProtection
//!   → bookViews → sheets → ...
//! ```
//!
//! ## Layout invariants
//!
//! `<fileSharing>` MUST come BEFORE `<workbookPr>`; Excel rejects the
//! workbook otherwise. `<workbookProtection>` MUST come AFTER
//! `<workbookPr>` but BEFORE `<bookViews>`. The splicer handles the
//! "no existing element" case by inserting fresh elements at the
//! correct position; existing elements are replaced in place
//! (preserving every other byte of `xl/workbook.xml`).
//!
//! ## Empty input ⇒ identity
//!
//! When the queued `WorkbookSecurity` has both fields `None` the
//! function returns the source bytes unchanged. The patcher's
//! Phase 2.5q calls into this only when the user actually set
//! `wb.security` or `wb.fileSharing`.

use quick_xml::events::Event;
use quick_xml::Reader as XmlReader;

use wolfxl_writer::parse::workbook_security::{
    emit_file_sharing, emit_workbook_protection, WorkbookSecurity,
};

// ---------------------------------------------------------------------------
// Public entry point
// ---------------------------------------------------------------------------

/// Splice `security` into `workbook_xml`.
///
/// Returns the updated bytes. Errors only on malformed XML
/// (no `<workbook>` root or unbalanced tags).
pub fn merge_workbook_security(
    workbook_xml: &[u8],
    security: &WorkbookSecurity,
) -> Result<Vec<u8>, String> {
    if security.is_empty() {
        return Ok(workbook_xml.to_vec());
    }

    let layout = scan_layout(workbook_xml)?;

    // Build the two replacement bytes (may be empty if the spec is
    // also empty, even though the bundle says is_empty is false).
    let new_protection: Vec<u8> = security
        .workbook_protection
        .as_ref()
        .map(emit_workbook_protection)
        .unwrap_or_default();
    let new_sharing: Vec<u8> = security
        .file_sharing
        .as_ref()
        .map(emit_file_sharing)
        .unwrap_or_default();

    // Apply the two splices in DESCENDING byte-offset order so earlier
    // edits don't shift later offsets. workbookProtection lives later
    // in the document than fileSharing, so we splice protection first
    // then sharing.

    let mut out: Vec<u8> = workbook_xml.to_vec();
    apply_protection_splice(&mut out, &layout, &new_protection)?;
    // Re-scan because the protection splice may have moved offsets.
    let new_layout = scan_layout(&out)?;
    apply_sharing_splice(&mut out, &new_layout, &new_sharing)?;

    Ok(out)
}

/// Replace only calcPr, retaining the surrounding workbook XML verbatim.
pub fn merge_calc_properties(
    workbook_xml: &[u8],
    attributes: &[(String, Option<String>)],
) -> Result<Vec<u8>, String> {
    use quick_xml::events::BytesStart;
    use quick_xml::Writer;

    let mut reader = XmlReader::from_reader(workbook_xml);
    let mut buffer = Vec::new();
    let mut depth = 0;
    let mut range = None;
    let mut insertion = None;
    let mut tag_name = String::from("calcPr");
    let mut old_attributes = Vec::new();
    loop {
        let before = reader.buffer_position() as usize;
        match reader
            .read_event_into(&mut buffer)
            .map_err(|e| e.to_string())?
        {
            Event::Start(element) | Event::Empty(element)
                if depth == 1 && element.local_name().as_ref() == b"calcPr" =>
            {
                tag_name = String::from_utf8(element.name().as_ref().to_vec())
                    .map_err(|e| e.to_string())?;
                for attribute in element.attributes() {
                    let attribute = attribute.map_err(|e| e.to_string())?;
                    old_attributes.push((
                        String::from_utf8(attribute.key.as_ref().to_vec())
                            .map_err(|e| e.to_string())?,
                        Some(
                            attribute
                                .unescape_value()
                                .map_err(|e| e.to_string())?
                                .into_owned(),
                        ),
                    ));
                }
                // A non-empty calcPr ends in '>', not '/>'.
                let after = reader.buffer_position() as usize;
                if !workbook_xml[before..after].ends_with(b"/>") {
                    reader
                        .read_to_end_into(element.name(), &mut Vec::new())
                        .map_err(|e| e.to_string())?;
                }
                range = Some((before, reader.buffer_position() as usize));
            }
            Event::Start(element) => {
                if depth == 0 {
                    let name = element.name();
                    if let Some(colon) = name.as_ref().iter().position(|byte| *byte == b':') {
                        tag_name = format!(
                            "{}:calcPr",
                            std::str::from_utf8(&name.as_ref()[..colon])
                                .map_err(|e| e.to_string())?
                        );
                    }
                }
                if depth == 1 && is_after_calc_pr(element.local_name().as_ref()) {
                    insertion.get_or_insert(before);
                }
                depth += 1;
            }
            Event::Empty(element) if depth == 1 => {
                if is_after_calc_pr(element.local_name().as_ref()) {
                    insertion.get_or_insert(before);
                }
            }
            Event::End(_) => {
                depth -= 1;
                if depth == 0 {
                    insertion.get_or_insert(before);
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    let (start, end) = range
        .or_else(|| insertion.map(|position| (position, position)))
        .ok_or_else(|| "workbook root missing".to_string())?;
    old_attributes.retain(|(name, _)| !attributes.iter().any(|(key, _)| key == name));
    let mut element = BytesStart::new(tag_name);
    for (name, value) in old_attributes.iter().chain(attributes) {
        if let Some(value) = value {
            element.push_attribute((name.as_str(), value.as_str()));
        }
    }
    let mut writer = Writer::new(Vec::new());
    writer
        .write_event(Event::Empty(element))
        .map_err(|e| e.to_string())?;
    let replacement = writer.into_inner();
    let mut output = Vec::with_capacity(workbook_xml.len() - (end - start) + replacement.len());
    output.extend_from_slice(&workbook_xml[..start]);
    output.extend_from_slice(&replacement);
    output.extend_from_slice(&workbook_xml[end..]);
    Ok(output)
}

fn is_after_calc_pr(name: &[u8]) -> bool {
    matches!(
        name,
        b"oleSize"
            | b"customWorkbookViews"
            | b"pivotCaches"
            | b"smartTagPr"
            | b"smartTagTypes"
            | b"webPublishing"
            | b"fileRecoveryPr"
            | b"webPublishObjects"
            | b"extLst"
    )
}

// ---------------------------------------------------------------------------
// Layout discovery
// ---------------------------------------------------------------------------

/// Byte offsets discovered for the splice points.
#[derive(Debug, Default, Clone)]
struct Layout {
    /// `(start_of_<workbookProtection, end_of_close_tag)` byte range
    /// when the source has an existing element. Self-closing or full
    /// open/close form, both supported.
    protection_outer: Option<(usize, usize)>,
    /// `(start_of_<fileSharing, end_of_close_tag)` byte range.
    sharing_outer: Option<(usize, usize)>,
    /// Byte offset just AFTER the close of `<fileVersion>` (or the
    /// open of `<workbookPr>` if no `<fileVersion>` exists). Used to
    /// inject a new `<fileSharing>` when the source has none.
    sharing_inject_at: usize,
    /// Byte offset just AFTER the close of `<workbookPr>` (the
    /// element closes self-closingly in every workbook this writer
    /// has met). Used to inject a new `<workbookProtection>` when
    /// the source has none.
    protection_inject_at: usize,
}

fn scan_layout(workbook_xml: &[u8]) -> Result<Layout, String> {
    let xml_str =
        std::str::from_utf8(workbook_xml).map_err(|e| format!("workbook.xml not UTF-8: {e}"))?;
    let mut reader = XmlReader::from_str(xml_str);
    reader.config_mut().trim_text(false);

    let mut layout = Layout::default();
    let mut buf = Vec::new();

    // Track whether we've already discovered fileVersion / workbookPr
    // inject points so we can fall back to <bookViews> / <sheets> for
    // the inject point when those siblings are absent.
    let mut after_file_version: Option<usize> = None;
    let mut after_workbook_pr: Option<usize> = None;
    let mut at_book_views: Option<usize> = None;
    let mut at_sheets: Option<usize> = None;

    loop {
        let pos_before = reader.buffer_position() as usize;
        match reader.read_event_into(&mut buf) {
            Ok(Event::Eof) => break,
            Ok(Event::Empty(ref e)) => {
                let pos_after = reader.buffer_position() as usize;
                let name = e.name();
                let local = name.as_ref();
                match local {
                    b"fileVersion" => {
                        after_file_version = Some(pos_after);
                    }
                    b"workbookPr" => {
                        after_workbook_pr = Some(pos_after);
                    }
                    b"workbookProtection" => {
                        layout.protection_outer = Some((pos_before, pos_after));
                    }
                    b"fileSharing" => {
                        layout.sharing_outer = Some((pos_before, pos_after));
                    }
                    _ => {}
                }
            }
            Ok(Event::Start(ref e)) => {
                let pos_after = reader.buffer_position() as usize;
                let name = e.name();
                let local = name.as_ref();
                match local {
                    b"bookViews" => {
                        at_book_views = Some(pos_before);
                    }
                    b"sheets" => {
                        at_sheets = Some(pos_before);
                    }
                    b"workbookPr" => {
                        // Defensive: workbookPr is normally self-closing.
                        // If we ever meet the open form, treat its
                        // close as the inject-after position.
                        let _ = pos_after;
                    }
                    b"workbookProtection" => {
                        // Open form (rare). Search for the close.
                        let close = find_close_tag(workbook_xml, pos_after, b"workbookProtection")
                            .ok_or_else(|| "unclosed <workbookProtection>".to_string())?;
                        layout.protection_outer = Some((pos_before, close));
                    }
                    b"fileSharing" => {
                        let close = find_close_tag(workbook_xml, pos_after, b"fileSharing")
                            .ok_or_else(|| "unclosed <fileSharing>".to_string())?;
                        layout.sharing_outer = Some((pos_before, close));
                    }
                    _ => {}
                }
            }
            Err(e) => return Err(format!("parse error at byte {pos_before}: {e}")),
            _ => {}
        }
        buf.clear();
    }

    // Resolve inject points by falling back when the preferred sibling
    // is absent.
    layout.sharing_inject_at = after_file_version
        .or(after_workbook_pr)
        .or(at_book_views)
        .or(at_sheets)
        .ok_or_else(|| {
            "could not find <fileVersion>/<workbookPr>/<bookViews>/<sheets> in workbook.xml"
                .to_string()
        })?;

    layout.protection_inject_at = after_workbook_pr
        .or(at_book_views)
        .or(at_sheets)
        .ok_or_else(|| {
            "could not find <workbookPr>/<bookViews>/<sheets> in workbook.xml".to_string()
        })?;

    Ok(layout)
}

/// Search forward from `start` for the close of an element named
/// `local_name`, returning the byte offset just past `</name>`.
fn find_close_tag(haystack: &[u8], start: usize, local_name: &[u8]) -> Option<usize> {
    // Build the close tag bytes "</name>".
    let mut needle = Vec::with_capacity(local_name.len() + 3);
    needle.push(b'<');
    needle.push(b'/');
    needle.extend_from_slice(local_name);
    needle.push(b'>');
    haystack[start..]
        .windows(needle.len())
        .position(|w| w == needle.as_slice())
        .map(|rel| start + rel + needle.len())
}

// ---------------------------------------------------------------------------
// Splice application
// ---------------------------------------------------------------------------

fn apply_protection_splice(
    out: &mut Vec<u8>,
    layout: &Layout,
    new_bytes: &[u8],
) -> Result<(), String> {
    match (layout.protection_outer, new_bytes.is_empty()) {
        (Some((s, e)), true) => {
            // Spec slot is None ⇒ remove the existing element.
            out.drain(s..e);
        }
        (Some((s, e)), false) => {
            // Replace existing element in place.
            let mut tail = out.split_off(e);
            out.truncate(s);
            out.extend_from_slice(new_bytes);
            out.append(&mut tail);
        }
        (None, true) => {
            // No existing element, no new bytes → no-op.
        }
        (None, false) => {
            // Inject at the canonical position.
            let inject = layout.protection_inject_at;
            let mut tail = out.split_off(inject);
            out.extend_from_slice(new_bytes);
            out.append(&mut tail);
        }
    }
    Ok(())
}

fn apply_sharing_splice(
    out: &mut Vec<u8>,
    layout: &Layout,
    new_bytes: &[u8],
) -> Result<(), String> {
    match (layout.sharing_outer, new_bytes.is_empty()) {
        (Some((s, e)), true) => {
            out.drain(s..e);
        }
        (Some((s, e)), false) => {
            let mut tail = out.split_off(e);
            out.truncate(s);
            out.extend_from_slice(new_bytes);
            out.append(&mut tail);
        }
        (None, true) => {}
        (None, false) => {
            let inject = layout.sharing_inject_at;
            let mut tail = out.split_off(inject);
            out.extend_from_slice(new_bytes);
            out.append(&mut tail);
        }
    }
    Ok(())
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;
    use wolfxl_writer::parse::workbook_security::{
        FileSharingSpec, WorkbookProtectionSpec, WorkbookSecurity,
    };

    const MINIMAL_WB: &[u8] = b"<?xml version=\"1.0\"?>\
        <workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\
        <fileVersion appName=\"xl\"/>\
        <workbookPr date1904=\"false\"/>\
        <bookViews><workbookView/></bookViews>\
        <sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets>\
        </workbook>";

    #[test]
    fn empty_security_returns_identity() {
        let bytes = merge_workbook_security(MINIMAL_WB, &WorkbookSecurity::default()).unwrap();
        assert_eq!(bytes, MINIMAL_WB);
    }

    #[test]
    fn injects_protection_after_workbook_pr() {
        let security = WorkbookSecurity {
            workbook_protection: Some(WorkbookProtectionSpec {
                lock_structure: true,
                ..Default::default()
            }),
            file_sharing: None,
        };
        let out = merge_workbook_security(MINIMAL_WB, &security).unwrap();
        let text = std::str::from_utf8(&out).unwrap();
        let pr = text.find("<workbookPr ").unwrap();
        let prot = text.find("<workbookProtection ").unwrap();
        let bv = text.find("<bookViews>").unwrap();
        assert!(pr < prot && prot < bv, "ordering: {text}");
        assert!(text.contains("lockStructure=\"1\""));
    }

    #[test]
    fn injects_sharing_after_file_version() {
        let security = WorkbookSecurity {
            workbook_protection: None,
            file_sharing: Some(FileSharingSpec {
                read_only_recommended: true,
                user_name: Some("alice".into()),
                ..Default::default()
            }),
        };
        let out = merge_workbook_security(MINIMAL_WB, &security).unwrap();
        let text = std::str::from_utf8(&out).unwrap();
        let fv = text.find("<fileVersion ").unwrap();
        let fs = text.find("<fileSharing ").unwrap();
        let pr = text.find("<workbookPr ").unwrap();
        assert!(fv < fs && fs < pr, "ordering: {text}");
        assert!(text.contains("userName=\"alice\""));
    }

    #[test]
    fn injects_both_blocks_canonical_order() {
        let security = WorkbookSecurity {
            workbook_protection: Some(WorkbookProtectionSpec {
                lock_structure: true,
                ..Default::default()
            }),
            file_sharing: Some(FileSharingSpec {
                read_only_recommended: true,
                ..Default::default()
            }),
        };
        let out = merge_workbook_security(MINIMAL_WB, &security).unwrap();
        let text = std::str::from_utf8(&out).unwrap();
        let positions: Vec<usize> = [
            "<fileVersion ",
            "<fileSharing ",
            "<workbookPr ",
            "<workbookProtection ",
            "<bookViews>",
            "<sheets>",
        ]
        .iter()
        .map(|tag| {
            text.find(tag)
                .unwrap_or_else(|| panic!("missing {tag}: {text}"))
        })
        .collect();
        for window in positions.windows(2) {
            assert!(window[0] < window[1], "ordering violated: {positions:?}");
        }
    }

    #[test]
    fn replaces_existing_protection_in_place() {
        let src: &[u8] = b"<?xml version=\"1.0\"?>\
            <workbook>\
            <fileVersion appName=\"xl\"/>\
            <workbookPr/>\
            <workbookProtection lockStructure=\"1\"/>\
            <bookViews><workbookView/></bookViews>\
            <sheets><sheet name=\"A\" sheetId=\"1\" r:id=\"rId1\"/></sheets>\
            </workbook>";
        let security = WorkbookSecurity {
            workbook_protection: Some(WorkbookProtectionSpec {
                lock_structure: true,
                lock_windows: true,
                ..Default::default()
            }),
            file_sharing: None,
        };
        let out = merge_workbook_security(src, &security).unwrap();
        let text = std::str::from_utf8(&out).unwrap();
        // Only ONE occurrence — the original was replaced, not appended.
        assert_eq!(text.matches("<workbookProtection ").count(), 1, "{text}");
        assert!(text.contains("lockWindows=\"1\""));
    }
}
