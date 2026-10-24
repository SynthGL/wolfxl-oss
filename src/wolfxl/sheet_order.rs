//! `xl/workbook.xml` `<sheets>` reorder + `<definedName localSheetId>` remap
//! for modify mode (RFC-036).
//!
//! Used by `XlsxPatcher::do_save`'s Phase 2.5h to:
//!   1. Parse the existing `<sheets>` order out of `xl/workbook.xml`.
//!   2. Apply a queue of `(sheet_name, offset)` moves in order, each
//!      against the running tab list. Indices are clamped to `[0, n-1]`.
//!   3. Re-emit the `<sheets>` block with the same `<sheet …/>` byte
//!      slices in the new order — every attribute on each `<sheet>`
//!      element flows through verbatim.
//!   4. Rewrite every `<definedName>` whose `localSheetId` attribute
//!      maps to a moved position. Only the integer attribute *value*
//!      is replaced; the rest of the element's bytes survive.
//!
//! ## Why this lives here, not in `wolfxl-writer`
//!
//! Same reasoning as `defined_names.rs`: the patcher must surgically
//! rewrite `xl/workbook.xml` while preserving every other child of
//! `<workbook>` byte-for-byte. The native writer builds workbook.xml
//! from a structured model and gets the new tab order for free once
//! `Workbook._sheet_names` is updated. RFC-036 is patcher-only.
//!
//! ## RFC-021 / RFC-036 sequencing
//!
//! Both RFCs mutate `xl/workbook.xml`. Phase 2.5h (this module) runs
//! BEFORE Phase 2.5f (defined-names), so the defined-names merger
//! always sees a workbook.xml whose `<sheets>` order and
//! `localSheetId` integers already reflect the move. The two phases
//! compose without re-parsing.

use std::collections::BTreeMap;

use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader as XmlReader;
use quick_xml::Writer as XmlWriter;

// ---------------------------------------------------------------------------
// Public entry point
// ---------------------------------------------------------------------------

/// Result of applying the queued moves: the rewritten `xl/workbook.xml`
/// bytes and the new tab order (so the caller can update its
/// `sheet_order` field for downstream phases — RFC-020 app.xml,
/// RFC-026 CF aggregation).
#[derive(Debug, Clone)]
pub struct SheetOrderResult {
    pub workbook_xml: Vec<u8>,
    pub new_order: Vec<String>,
}

/// Parse `workbook_xml`, apply each `(sheet_name, offset)` move in order,
/// and return the rewritten bytes plus the resulting tab order.
///
/// - Each move is applied against the running tab list, then committed
///   before the next move runs. This composes correctly for multi-move
///   queues.
/// - `offset` may be negative. The new index `idx + offset` is clamped
///   to `[0, n-1]` where `n` is the current tab count.
/// - A move whose `sheet_name` is not in the current tab list is
///   skipped (logged in the future; today it's a silent no-op to
///   match the validate-on-the-Python-side contract).
/// - An empty `moves` slice returns the source bytes verbatim — the
///   modify-mode no-op invariant.
///
/// Returns `Err` on malformed XML (no `<sheets>` block).
pub fn merge_sheet_moves(
    workbook_xml: &[u8],
    moves: &[(String, i32)],
) -> Result<SheetOrderResult, String> {
    let layout = scan_workbook_layout(workbook_xml)?;

    // Empty queue → identity. Source bytes survive byte-for-byte.
    if moves.is_empty() {
        return Ok(SheetOrderResult {
            workbook_xml: workbook_xml.to_vec(),
            new_order: layout
                .sheet_entries
                .iter()
                .map(|e| e.name.clone())
                .collect(),
        });
    }

    // Apply each move against the running tab list.
    let mut entries: Vec<SheetEntry> = layout.sheet_entries.clone();
    for (name, offset) in moves {
        let Some(idx) = entries.iter().position(|e| &e.name == name) else {
            // Unknown sheet name — silent skip. The Python coordinator
            // is expected to validate before queueing, so reaching this
            // branch implies a bug above us; we'd rather drop the move
            // than corrupt the file.
            continue;
        };
        let n = entries.len() as i64;
        if n == 0 {
            continue;
        }
        let mut new_pos = (idx as i64) + (*offset as i64);
        if new_pos < 0 {
            new_pos = 0;
        }
        if new_pos > n - 1 {
            new_pos = n - 1;
        }
        let entry = entries.remove(idx);
        entries.insert(new_pos as usize, entry);
    }

    // Build the position remap: old_pos → new_pos. Only entries whose
    // position changed get a remap entry (saves work in Phase 2 below).
    let mut remap: BTreeMap<u32, u32> = BTreeMap::new();
    for (new_pos, entry) in entries.iter().enumerate() {
        let old_pos = entry.original_pos;
        if old_pos as usize != new_pos {
            remap.insert(old_pos, new_pos as u32);
        }
    }

    // ---- Pass 1: rewrite the <sheets> block in place. -------------------
    let mut reordered_sheets: Vec<u8> = Vec::with_capacity(layout.sheets_inner_len());
    for entry in &entries {
        reordered_sheets.extend_from_slice(&entry.raw);
    }

    // ---- Pass 2: rewrite <definedName localSheetId="N"> values. ---------
    // We only touch the localSheetId integer; everything else flows
    // through. If there's no <definedNames> block at all, the second
    // pass is a no-op.
    let with_sheets_rewritten = splice_sheets_inner(
        workbook_xml,
        layout.sheets_inner_start,
        layout.sheets_inner_end,
        &reordered_sheets,
    );

    let final_bytes = if remap.is_empty() {
        with_sheets_rewritten
    } else {
        rewrite_local_sheet_ids(&with_sheets_rewritten, &remap)?
    };

    let new_order: Vec<String> = entries.iter().map(|e| e.name.clone()).collect();

    Ok(SheetOrderResult {
        workbook_xml: final_bytes,
        new_order,
    })
}

/// Rename `<sheet name="...">` attributes in `xl/workbook.xml`.
///
/// Used by modify mode when a loaded worksheet's `.title` is changed. The
/// patcher updates its in-memory sheet maps eagerly, while this function makes
/// the workbook tab name mutation durable in the saved OOXML.
pub fn merge_sheet_renames(
    workbook_xml: &[u8],
    renames: &[(String, String)],
) -> Result<Vec<u8>, String> {
    if renames.is_empty() {
        return Ok(workbook_xml.to_vec());
    }
    let rename_map: BTreeMap<&str, &str> = renames
        .iter()
        .map(|(old, new)| (old.as_str(), new.as_str()))
        .collect();

    let mut reader = XmlReader::from_reader(workbook_xml);
    reader.config_mut().trim_text(false);
    let mut writer = XmlWriter::new(Vec::with_capacity(workbook_xml.len()));
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.local_name().as_ref() == "sheet" => {
                let renamed = rename_sheet_event(&e, &rename_map)?;
                writer
                    .write_event(Event::Start(renamed))
                    .map_err(|e| format!("workbook.xml write error: {e}"))?;
            }
            Ok(Event::Empty(e)) if e.local_name().as_ref() == "sheet" => {
                let renamed = rename_sheet_event(&e, &rename_map)?;
                writer
                    .write_event(Event::Empty(renamed))
                    .map_err(|e| format!("workbook.xml write error: {e}"))?;
            }
            Ok(Event::Eof) => break,
            Ok(event) => writer
                .write_event(event)
                .map_err(|e| format!("workbook.xml write error: {e}"))?,
            Err(e) => return Err(format!("workbook.xml parse error: {e}")),
        }
        buf.clear();
    }

    Ok(writer.into_inner())
}

fn rename_sheet_event(
    event: &BytesStart<'_>,
    rename_map: &BTreeMap<&str, &str>,
) -> Result<BytesStart<'static>, String> {
    let mut out = BytesStart::new(event.name().as_ref().to_owned());
    for attr in event.attributes().with_checks(false) {
        let attr = attr.map_err(|e| format!("workbook.xml sheet attr error: {e}"))?;
        if attr.key.as_ref() == "name" {
            let value = quick_xml::escape::unescape(&attr.value)
                .map_err(|e| format!("workbook.xml sheet name decode error: {e}"))?;
            if let Some(new_name) = rename_map.get(value.as_ref()) {
                out.push_attribute(("name", *new_name));
            } else {
                out.push_attribute(("name", value.as_ref()));
            }
        } else {
            out.push_attribute(attr);
        }
    }
    Ok(out)
}

/// Remove `<sheet>` entries from `xl/workbook.xml` and keep sheet-scoped
/// defined names aligned with the surviving tab indexes.
pub fn merge_sheet_deletes(
    workbook_xml: &[u8],
    deletes: &[String],
) -> Result<SheetOrderResult, String> {
    let layout = scan_workbook_layout(workbook_xml)?;
    if deletes.is_empty() {
        return Ok(SheetOrderResult {
            workbook_xml: workbook_xml.to_vec(),
            new_order: layout
                .sheet_entries
                .iter()
                .map(|e| e.name.clone())
                .collect(),
        });
    }

    let delete_set: std::collections::BTreeSet<&str> = deletes.iter().map(|s| s.as_str()).collect();
    let mut deleted_positions: std::collections::BTreeSet<u32> = std::collections::BTreeSet::new();
    let mut entries: Vec<SheetEntry> = Vec::new();
    for entry in &layout.sheet_entries {
        if delete_set.contains(entry.name.as_str()) {
            deleted_positions.insert(entry.original_pos);
        } else {
            entries.push(entry.clone());
        }
    }

    let mut remap: BTreeMap<u32, u32> = BTreeMap::new();
    for (new_pos, entry) in entries.iter().enumerate() {
        if entry.original_pos as usize != new_pos {
            remap.insert(entry.original_pos, new_pos as u32);
        }
    }

    let mut rewritten_sheets: Vec<u8> = Vec::with_capacity(layout.sheets_inner_len());
    for entry in &entries {
        rewritten_sheets.extend_from_slice(&entry.raw);
    }
    let with_sheets_rewritten = splice_sheets_inner(
        workbook_xml,
        layout.sheets_inner_start,
        layout.sheets_inner_end,
        &rewritten_sheets,
    );

    let without_deleted_defined_names =
        remove_defined_names_for_deleted_sheets(&with_sheets_rewritten, &deleted_positions)?;
    let final_bytes = if remap.is_empty() {
        without_deleted_defined_names
    } else {
        rewrite_local_sheet_ids(&without_deleted_defined_names, &remap)?
    };

    Ok(SheetOrderResult {
        workbook_xml: final_bytes,
        new_order: entries.iter().map(|e| e.name.clone()).collect(),
    })
}

// ---------------------------------------------------------------------------
// Internal: layout scan
// ---------------------------------------------------------------------------

#[derive(Debug, Clone)]
struct SheetEntry {
    /// `name` attribute on the `<sheet …/>` element.
    name: String,
    /// 0-based position in the source `<sheets>` block.
    original_pos: u32,
    /// Verbatim bytes covering the `<sheet …/>` element (start through
    /// either `/>` or `</sheet>`). Whitespace BETWEEN entries is NOT
    /// captured — that's the price of the splice; it's acceptable
    /// because the patcher's contract is byte-stability for unchanged
    /// queues, and a queued move is a known-byte-changing operation.
    raw: Vec<u8>,
}

#[derive(Debug, Default)]
struct WorkbookLayout {
    /// Byte offset of the first byte after `<sheets>`'s start tag.
    sheets_inner_start: usize,
    /// Byte offset of the first byte of `</sheets>`'s end tag.
    sheets_inner_end: usize,
    /// Each `<sheet …/>` element parsed out of the block, in source
    /// order. `<sheets>` is required to have at least one child by
    /// ECMA-376, but we accept zero defensively.
    sheet_entries: Vec<SheetEntry>,
}

impl WorkbookLayout {
    fn sheets_inner_len(&self) -> usize {
        self.sheets_inner_end
            .saturating_sub(self.sheets_inner_start)
    }
}

fn scan_workbook_layout(xml: &[u8]) -> Result<WorkbookLayout, String> {
    let s =
        std::str::from_utf8(xml).map_err(|e| format!("workbook.xml is not valid UTF-8: {e}"))?;
    let offset = utf8_bom_offset(xml);
    let mut reader = XmlReader::from_str(s);
    reader.config_mut().trim_text(false);
    let mut buf: Vec<u8> = Vec::new();

    let mut sheets_inner_start: Option<usize> = None;
    let mut sheets_inner_end: Option<usize> = None;
    let mut entries: Vec<SheetEntry> = Vec::new();
    let mut in_sheets = false;
    let mut current_sheet_start: Option<usize> = None;
    let mut current_sheet_name: Option<String> = None;

    loop {
        let pre = reader.buffer_position() as usize;
        let evt = reader.read_event_into(&mut buf);
        let post = reader.buffer_position() as usize;

        match evt {
            Ok(Event::Start(ref e)) => {
                if e.local_name().as_ref() == "sheets" && sheets_inner_start.is_none() {
                    sheets_inner_start = Some(post + offset);
                    in_sheets = true;
                } else if in_sheets && e.local_name().as_ref() == "sheet" {
                    current_sheet_start = Some(pre + offset);
                    current_sheet_name = Some(parse_sheet_name(e)?);
                }
            }
            Ok(Event::Empty(ref e)) => {
                if in_sheets && e.local_name().as_ref() == "sheet" {
                    let name = parse_sheet_name(e)?;
                    entries.push(SheetEntry {
                        name,
                        original_pos: entries.len() as u32,
                        raw: xml[pre + offset..post + offset].to_vec(),
                    });
                }
            }
            Ok(Event::End(ref e)) => {
                if in_sheets && e.local_name().as_ref() == "sheet" {
                    if let (Some(start), Some(name)) =
                        (current_sheet_start.take(), current_sheet_name.take())
                    {
                        entries.push(SheetEntry {
                            name,
                            original_pos: entries.len() as u32,
                            raw: xml[start..post + offset].to_vec(),
                        });
                    }
                } else if e.local_name().as_ref() == "sheets" && in_sheets {
                    sheets_inner_end = Some(pre + offset);
                    in_sheets = false;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(format!("workbook.xml parse error: {e}")),
            _ => {}
        }
        buf.clear();
    }

    let sheets_inner_start =
        sheets_inner_start.ok_or_else(|| "workbook.xml has no <sheets> opening tag".to_string())?;
    let sheets_inner_end =
        sheets_inner_end.ok_or_else(|| "workbook.xml has no </sheets> closing tag".to_string())?;

    Ok(WorkbookLayout {
        sheets_inner_start,
        sheets_inner_end,
        sheet_entries: entries,
    })
}

fn parse_sheet_name(e: &quick_xml::events::BytesStart<'_>) -> Result<String, String> {
    for a in e.attributes().with_checks(false) {
        let a = a.map_err(|e| format!("workbook.xml sheet attr error: {e}"))?;
        if a.key.as_ref() == "name" {
            return quick_xml::escape::unescape(&a.value)
                .map(|v| v.into_owned())
                .map_err(|e| format!("workbook.xml sheet name decode error: {e}"));
        }
    }
    Ok(String::new())
}

fn splice_sheets_inner(
    src: &[u8],
    inner_start: usize,
    inner_end: usize,
    new_inner: &[u8],
) -> Vec<u8> {
    let mut out: Vec<u8> = Vec::with_capacity(src.len());
    out.extend_from_slice(&src[..inner_start]);
    out.extend_from_slice(new_inner);
    out.extend_from_slice(&src[inner_end..]);
    out
}

// ---------------------------------------------------------------------------
// Internal: `<definedName localSheetId="N">` rewrite.
//
// We do NOT re-emit the entire <definedNames> block (which would lose
// inter-child whitespace). We do a streaming scan, locate each
// `<definedName>` start tag (including the `Empty` form), find its
// `localSheetId="…"` attribute byte range, and replace just the
// integer value when the integer is in the remap. Other attributes
// flow through unchanged.
// ---------------------------------------------------------------------------

fn rewrite_local_sheet_ids(src: &[u8], remap: &BTreeMap<u32, u32>) -> Result<Vec<u8>, String> {
    let s = std::str::from_utf8(src)
        .map_err(|e| format!("workbook.xml is not valid UTF-8 after sheets splice: {e}"))?;
    if !s.contains("definedName") {
        return Ok(src.to_vec());
    }
    let offset = utf8_bom_offset(src);
    let mut reader = XmlReader::from_str(s);
    reader.config_mut().trim_text(false);
    let mut buf: Vec<u8> = Vec::new();

    // (start_byte_of_attr_value, end_byte_of_attr_value, old_val, new_val).
    // We collect rewrites first, then apply them in reverse-byte order
    // so earlier offsets remain valid.
    let mut rewrites: Vec<(usize, usize, String)> = Vec::new();

    loop {
        let pre = reader.buffer_position() as usize;
        let evt = reader.read_event_into(&mut buf);
        let _post = reader.buffer_position() as usize;
        match evt {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                if e.local_name().as_ref() == "definedName" {
                    if let Some((val_start, val_end, old)) =
                        find_local_sheet_id_value(src, pre + offset)
                    {
                        if let Ok(parsed) = old.parse::<u32>() {
                            if let Some(&new_idx) = remap.get(&parsed) {
                                rewrites.push((val_start, val_end, new_idx.to_string()));
                            }
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(format!("definedName scan error: {e}")),
            _ => {}
        }
        buf.clear();
    }

    if rewrites.is_empty() {
        return Ok(src.to_vec());
    }

    // Apply in reverse so earlier offsets stay valid.
    rewrites.sort_by_key(|(s, _, _)| *s);
    let mut out: Vec<u8> = Vec::with_capacity(src.len() + 8);
    let mut cursor: usize = 0;
    for (val_start, val_end, new_val) in &rewrites {
        out.extend_from_slice(&src[cursor..*val_start]);
        out.extend_from_slice(new_val.as_bytes());
        cursor = *val_end;
    }
    out.extend_from_slice(&src[cursor..]);
    Ok(out)
}

fn remove_defined_names_for_deleted_sheets(
    src: &[u8],
    deleted_positions: &std::collections::BTreeSet<u32>,
) -> Result<Vec<u8>, String> {
    if deleted_positions.is_empty() {
        return Ok(src.to_vec());
    }
    let s = std::str::from_utf8(src)
        .map_err(|e| format!("workbook.xml is not valid UTF-8 after sheet delete: {e}"))?;
    if !s.contains("definedName") {
        return Ok(src.to_vec());
    }
    let offset = utf8_bom_offset(src);
    let mut reader = XmlReader::from_str(s);
    reader.config_mut().trim_text(false);
    let mut buf: Vec<u8> = Vec::new();
    let mut removals: Vec<(usize, usize)> = Vec::new();
    let mut active_defined_name: Option<(usize, u32)> = None;

    loop {
        let pre = reader.buffer_position() as usize;
        let evt = reader.read_event_into(&mut buf);
        let post = reader.buffer_position() as usize;
        match evt {
            Ok(Event::Empty(e)) if e.local_name().as_ref() == "definedName" => {
                if let Some((_, _, old)) = find_local_sheet_id_value(src, pre + offset) {
                    if let Ok(parsed) = old.parse::<u32>() {
                        if deleted_positions.contains(&parsed) {
                            removals.push((pre + offset, post + offset));
                        }
                    }
                }
            }
            Ok(Event::Start(e)) if e.local_name().as_ref() == "definedName" => {
                if let Some((_, _, old)) = find_local_sheet_id_value(src, pre + offset) {
                    if let Ok(parsed) = old.parse::<u32>() {
                        if deleted_positions.contains(&parsed) {
                            active_defined_name = Some((pre + offset, parsed));
                        }
                    }
                }
            }
            Ok(Event::End(e)) if e.local_name().as_ref() == "definedName" => {
                if let Some((start, _)) = active_defined_name.take() {
                    removals.push((start, post + offset));
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(format!("definedName delete scan error: {e}")),
            _ => {}
        }
        buf.clear();
    }

    if removals.is_empty() {
        return Ok(src.to_vec());
    }

    removals.sort_by_key(|(start, _)| *start);
    let mut out = Vec::with_capacity(src.len());
    let mut cursor = 0;
    for (start, end) in removals {
        out.extend_from_slice(&src[cursor..start]);
        cursor = end;
    }
    out.extend_from_slice(&src[cursor..]);
    Ok(out)
}

fn utf8_bom_offset(src: &[u8]) -> usize {
    if src.starts_with(b"\xEF\xBB\xBF") {
        3
    } else {
        0
    }
}

/// Given the byte offset of a `<definedName` start tag, locate the
/// byte range of its `localSheetId` attribute *value* (between the
/// quotes, exclusive). Returns `None` if the attribute is absent.
///
/// We can't get the byte-precise attribute span out of `quick_xml`'s
/// `Attribute` struct directly, so we do a small ASCII scan of the
/// tag bytes. Acceptable because attribute values are ASCII integers
/// here — no Unicode escapes or entity refs.
fn find_local_sheet_id_value(src: &[u8], tag_start: usize) -> Option<(usize, usize, String)> {
    // Find the matching `>` that closes this start tag (or self-close).
    // We need a quote-aware scan because `>` can appear inside attr values
    // in theory; in practice <definedName> attrs are simple, but be safe.
    let n = src.len();
    let mut i = tag_start;
    if i + 1 >= n || src[i] != b'<' {
        return None;
    }
    i += 1;
    let mut in_quote: Option<u8> = None;
    while i < n {
        let b = src[i];
        match in_quote {
            Some(q) => {
                if b == q {
                    in_quote = None;
                }
            }
            None => {
                if b == b'"' || b == b'\'' {
                    in_quote = Some(b);
                } else if b == b'>' {
                    break;
                }
            }
        }
        i += 1;
    }
    let tag_end = i;
    if tag_end <= tag_start {
        return None;
    }
    let tag_bytes = &src[tag_start..=tag_end];
    let needle = b"localSheetId";
    let mut j = 0;
    let m = tag_bytes.len();
    while j + needle.len() < m {
        if &tag_bytes[j..j + needle.len()] == needle {
            // Make sure preceding char is whitespace (so we don't
            // match `xlocalSheetId`).
            if j == 0 || !is_xml_attr_name_char(tag_bytes[j - 1]) {
                // Skip past name + optional whitespace + '='.
                let mut k = j + needle.len();
                while k < m && (tag_bytes[k] == b' ' || tag_bytes[k] == b'\t') {
                    k += 1;
                }
                if k < m && tag_bytes[k] == b'=' {
                    k += 1;
                    while k < m && (tag_bytes[k] == b' ' || tag_bytes[k] == b'\t') {
                        k += 1;
                    }
                    if k < m && (tag_bytes[k] == b'"' || tag_bytes[k] == b'\'') {
                        let quote = tag_bytes[k];
                        let value_start_local = k + 1;
                        let mut p = value_start_local;
                        while p < m && tag_bytes[p] != quote {
                            p += 1;
                        }
                        if p < m {
                            let val_bytes = &tag_bytes[value_start_local..p];
                            let val = std::str::from_utf8(val_bytes).ok()?.to_string();
                            return Some((tag_start + value_start_local, tag_start + p, val));
                        }
                    }
                }
            }
        }
        j += 1;
    }
    None
}

#[inline]
fn is_xml_attr_name_char(b: u8) -> bool {
    b.is_ascii_alphanumeric() || b == b'_' || b == b'-' || b == b':'
}

// ---------------------------------------------------------------------------
// Tests
//
// Inline pure-Rust tests. The patcher cdylib doesn't link standalone via
// `cargo test -p wolfxl --lib` (Python linkage), so these compile under
// `cargo build` and end-to-end behaviour is exercised via pytest. Same
// precedent as `defined_names.rs` and `properties.rs`.
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn sheet_rename_decodes_names_and_preserves_formula_references() {
        let xml = br#"<workbook><sheets><sheet name="A&amp;&#x42;" sheetId="1"/></sheets><definedNames><definedName name="X"> A &amp; B &#38; C &custom; </definedName></definedNames></workbook>"#;
        let out = merge_sheet_renames(xml, &[("A&B".into(), "C&\"D".into())]).unwrap();
        let text = std::str::from_utf8(&out).unwrap();
        assert!(text.contains(r#"name="C&amp;&quot;D""#));
        assert!(text.contains("> A &amp; B &#38; C &custom; </definedName>"));
        assert_eq!(merge_sheet_moves(&out, &[]).unwrap().new_order, vec!["C&\"D"]);
    }

    #[test]
    fn sheet_names_reject_unknown_entities() {
        let xml = br#"<workbook><sheets><sheet name="A&unknown;B" sheetId="1"/></sheets></workbook>"#;
        assert!(merge_sheet_moves(xml, &[]).is_err());
        assert!(merge_sheet_renames(xml, &[("A".into(), "B".into())]).is_err());
    }

    fn workbook_with_four_sheets() -> &'static str {
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="A" sheetId="1" r:id="rId1"/>
    <sheet name="B" sheetId="2" r:id="rId2"/>
    <sheet name="C" sheetId="3" r:id="rId3"/>
    <sheet name="D" sheetId="4" r:id="rId4"/>
  </sheets>
  <definedNames>
    <definedName name="X" localSheetId="0">A!$A$1</definedName>
    <definedName name="Y" localSheetId="2">C!$A$1</definedName>
    <definedName name="Z" localSheetId="3">D!$A$1</definedName>
    <definedName name="W">A!$A$1</definedName>
  </definedNames>
  <calcPr/>
</workbook>"#
    }

    fn workbook_with_two_sheets_no_defined_names() -> &'static str {
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    <sheet name="Sheet2" sheetId="2" r:id="rId2"/>
  </sheets>
  <calcPr/>
</workbook>"#
    }

    #[test]
    fn empty_moves_is_identity() {
        let xml = workbook_with_four_sheets();
        let res = merge_sheet_moves(xml.as_bytes(), &[]).expect("merge");
        assert_eq!(res.workbook_xml, xml.as_bytes());
        assert_eq!(res.new_order, vec!["A", "B", "C", "D"]);
    }

    #[test]
    fn forward_move_a_by_2_reorders_block() {
        let xml = workbook_with_four_sheets();
        let res = merge_sheet_moves(xml.as_bytes(), &[("A".to_string(), 2)]).expect("merge");
        assert_eq!(res.new_order, vec!["B", "C", "A", "D"]);
        let text = std::str::from_utf8(&res.workbook_xml).unwrap();
        // `<sheets>` order matches new_order. Match the order of `name=`
        // appearances in the rewritten block.
        let pos_b = text.find(r#"name="B""#).unwrap();
        let pos_c = text.find(r#"name="C""#).unwrap();
        let pos_a = text.find(r#"name="A""#).unwrap();
        let pos_d = text.find(r#"name="D""#).unwrap();
        assert!(pos_b < pos_c && pos_c < pos_a && pos_a < pos_d);
    }

    #[test]
    fn forward_move_remaps_local_sheet_id() {
        // [A,B,C,D] → move A by +2 → [B,C,A,D].
        // Remap: 0→2 (A moved), 1→0 (B), 2→1 (C). 3→3 (D) is identity.
        // Defined names referenced 0/2/3 → after remap: 2/1/3.
        let xml = workbook_with_four_sheets();
        let res = merge_sheet_moves(xml.as_bytes(), &[("A".to_string(), 2)]).expect("merge");
        let text = std::str::from_utf8(&res.workbook_xml).unwrap();
        assert!(
            text.contains(r#"<definedName name="X" localSheetId="2">A!$A$1</definedName>"#),
            "X should now reference position 2 (A's new position):\n{text}"
        );
        assert!(
            text.contains(r#"<definedName name="Y" localSheetId="1">C!$A$1</definedName>"#),
            "Y should now reference position 1 (C's new position):\n{text}"
        );
        assert!(
            text.contains(r#"<definedName name="Z" localSheetId="3">D!$A$1</definedName>"#),
            "Z should remain at position 3 (D unchanged):\n{text}"
        );
        assert!(
            text.contains(r#"<definedName name="W">A!$A$1</definedName>"#),
            "W has no localSheetId; must survive verbatim:\n{text}"
        );
    }

    #[test]
    fn backward_move_works() {
        // [A,B,C,D] → move D by -2 → [A,D,B,C].
        // Remap: 1→2 (B), 2→3 (C), 3→1 (D moved). 0→0 (A) identity.
        let xml = workbook_with_four_sheets();
        let res = merge_sheet_moves(xml.as_bytes(), &[("D".to_string(), -2)]).expect("merge");
        assert_eq!(res.new_order, vec!["A", "D", "B", "C"]);
        let text = std::str::from_utf8(&res.workbook_xml).unwrap();
        // Z was at 3; D moved 3 → 1, so Z should now read 1.
        assert!(text.contains(r#"localSheetId="1""#));
    }

    #[test]
    fn offset_zero_is_position_no_op_but_remap_is_empty() {
        let xml = workbook_with_four_sheets();
        let res = merge_sheet_moves(xml.as_bytes(), &[("B".to_string(), 0)]).expect("merge");
        assert_eq!(res.new_order, vec!["A", "B", "C", "D"]);
        // The block is re-emitted (the remove-then-insert at the same
        // index still touches bytes when whitespace differs). What we
        // require is that the defined-name integers are unchanged.
        let text = std::str::from_utf8(&res.workbook_xml).unwrap();
        assert!(text.contains(r#"<definedName name="X" localSheetId="0">"#));
        assert!(text.contains(r#"<definedName name="Y" localSheetId="2">"#));
        assert!(text.contains(r#"<definedName name="Z" localSheetId="3">"#));
    }

    #[test]
    fn high_offset_clamps_to_last() {
        let xml = workbook_with_four_sheets();
        let res = merge_sheet_moves(xml.as_bytes(), &[("A".to_string(), 100)]).expect("merge");
        assert_eq!(res.new_order, vec!["B", "C", "D", "A"]);
    }

    #[test]
    fn low_offset_clamps_to_first() {
        let xml = workbook_with_four_sheets();
        let res = merge_sheet_moves(xml.as_bytes(), &[("D".to_string(), -100)]).expect("merge");
        assert_eq!(res.new_order, vec!["D", "A", "B", "C"]);
    }

    #[test]
    fn workbook_with_no_defined_names_block_is_safe() {
        let xml = workbook_with_two_sheets_no_defined_names();
        let res = merge_sheet_moves(xml.as_bytes(), &[("Sheet2".to_string(), -1)]).expect("merge");
        assert_eq!(res.new_order, vec!["Sheet2", "Sheet1"]);
        let text = std::str::from_utf8(&res.workbook_xml).unwrap();
        assert!(text.contains(r#"name="Sheet2""#));
        assert!(text.contains(r#"name="Sheet1""#));
    }

    #[test]
    fn unknown_sheet_name_is_silent_skip() {
        let xml = workbook_with_four_sheets();
        let res =
            merge_sheet_moves(xml.as_bytes(), &[("DOES_NOT_EXIST".to_string(), 1)]).expect("merge");
        // Unknown moves leave order unchanged; remap is empty.
        assert_eq!(res.new_order, vec!["A", "B", "C", "D"]);
    }

    #[test]
    fn multiple_moves_compose() {
        // [A,B,C,D] → move A +2 → [B,C,A,D] → move B +2 → [C,A,B,D].
        // After first move:  A:0→2, B:1→0, C:2→1.
        // After second move: B:0→2 (in new tab list). The composite
        // for the original positions is:
        //   A's original pos 0 ends at new pos 1 ([C,A,B,D]).
        //   B's original pos 1 ends at new pos 2.
        //   C's original pos 2 ends at new pos 0.
        //   D unchanged at 3.
        let xml = workbook_with_four_sheets();
        let res = merge_sheet_moves(
            xml.as_bytes(),
            &[("A".to_string(), 2), ("B".to_string(), 2)],
        )
        .expect("merge");
        assert_eq!(res.new_order, vec!["C", "A", "B", "D"]);
        let text = std::str::from_utf8(&res.workbook_xml).unwrap();
        // X was at 0 (A); A is now at position 1. So X.localSheetId = 1.
        assert!(text.contains(r#"<definedName name="X" localSheetId="1">"#));
        // Y was at 2 (C); C is now at position 0. So Y.localSheetId = 0.
        assert!(text.contains(r#"<definedName name="Y" localSheetId="0">"#));
        // Z was at 3 (D); D is unchanged. So Z.localSheetId = 3.
        assert!(text.contains(r#"<definedName name="Z" localSheetId="3">"#));
    }

    #[test]
    fn single_sheet_workbook_is_safe() {
        let xml = r#"<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheets><sheet name="Only" sheetId="1"/></sheets>
</workbook>"#;
        let res = merge_sheet_moves(xml.as_bytes(), &[("Only".to_string(), 5)]).expect("merge");
        assert_eq!(res.new_order, vec!["Only"]);
    }

    #[test]
    fn missing_sheets_block_errors() {
        let xml = b"<?xml version=\"1.0\"?><workbook>no sheets</workbook>";
        assert!(merge_sheet_moves(xml, &[("X".to_string(), 1)]).is_err());
    }

    #[test]
    fn sheet_attributes_state_hidden_preserved() {
        let xml = r#"<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheets>
    <sheet name="Visible" sheetId="1" r:id="rId1"/>
    <sheet name="Hidden" sheetId="2" state="hidden" r:id="rId2"/>
  </sheets>
</workbook>"#;
        let res = merge_sheet_moves(xml.as_bytes(), &[("Hidden".to_string(), -1)]).expect("merge");
        let text = std::str::from_utf8(&res.workbook_xml).unwrap();
        // Hidden's state="hidden" attribute must survive verbatim.
        assert!(text.contains(r#"state="hidden""#));
        let pos_hidden = text.find(r#"name="Hidden""#).unwrap();
        let pos_visible = text.find(r#"name="Visible""#).unwrap();
        assert!(pos_hidden < pos_visible);
    }

    #[test]
    fn sheet_delete_without_defined_names_handles_utf8_bom_workbook() {
        let xml = b"\xEF\xBB\xBF<?xml version=\"1.0\" encoding=\"utf-8\"?>
<workbook xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"
          xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">
    <sheets>
        <sheet name=\"Some Sheet\" sheetId=\"1\" r:id=\"rId1\"/>
        <sheet name=\"Some Sheet Copy\" sheetId=\"2\" r:id=\"rId5\"/>
    </sheets>
</workbook>";
        let res = merge_sheet_deletes(xml, &["Some Sheet Copy".to_string()]).expect("merge");
        let text = std::str::from_utf8(&res.workbook_xml).unwrap();

        assert!(text.contains("Some Sheet"));
        assert!(!text.contains("Some Sheet Copy"));
        assert!(text.contains("<sheets>"));
        assert!(text.contains(r#"<sheet name="Some Sheet" sheetId="1" r:id="rId1"/>"#));
        assert_eq!(res.new_order, vec!["Some Sheet"]);

        let mut reader = XmlReader::from_str(text);
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Eof) => break,
                Ok(_) => {}
                Err(e) => panic!("workbook.xml should remain well-formed: {e}"),
            }
            buf.clear();
        }
    }

    #[test]
    fn find_local_sheet_id_value_basic() {
        let s = br#"<definedName name="X" localSheetId="42">x</definedName>"#;
        let (start, end, val) = find_local_sheet_id_value(s, 0).unwrap();
        assert_eq!(val, "42");
        assert_eq!(&s[start..end], b"42");
    }

    #[test]
    fn find_local_sheet_id_value_absent_returns_none() {
        let s = br#"<definedName name="X">x</definedName>"#;
        assert!(find_local_sheet_id_value(s, 0).is_none());
    }

    #[test]
    fn find_local_sheet_id_value_does_not_match_substring() {
        // "xlocalSheetId" should NOT match.
        let s = br#"<definedName name="X" xlocalSheetId="42">x</definedName>"#;
        assert!(find_local_sheet_id_value(s, 0).is_none());
    }
}
