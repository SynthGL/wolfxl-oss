//! `xl/workbook.xml` `<definedNames>` block rewriter for modify mode (RFC-021).
//!
//! Used by `XlsxPatcher::do_save`'s Phase 2.5f to:
//!   1. Parse the existing `<definedNames>` block in `xl/workbook.xml`.
//!   2. Upsert user-supplied [`DefinedNameMut`] entries by `(name, local_sheet_id)`.
//!   3. Re-emit the merged block in place, or inject a fresh `<definedNames>`
//!      element immediately after `</sheets>` if the source had none.
//!   4. Hand the updated `xl/workbook.xml` bytes back to the patcher's
//!      `file_patches` map.
//!
//! ## Why this lives here, not in `wolfxl-writer`
//!
//! The native writer (`crates/wolfxl-writer/src/emit/workbook_xml.rs`) builds
//! `xl/workbook.xml` from a structured `Workbook` model. The patcher has no
//! such model — it must surgically rewrite `xl/workbook.xml` while preserving
//! every other child of `<workbook>` (`fileVersion`, `workbookPr`,
//! `bookViews`, `sheets`, `calcPr`, `extLst`, …) byte-for-byte. The streaming
//! splice here covers the modify-mode contract; consolidation with the
//! writer's emitter is deferred until a third caller appears (RFC-020 §4.2
//! Option-2 precedent).
//!
//! ## RFC-012 / RFC-036 seam
//!
//! [`DefinedNameMut::formula`] is a plain string. RFC-036 (`move_sheet`)
//! will call RFC-012's translator on each formula BEFORE invoking
//! [`merge_defined_names`]. The merger never inspects formula contents —
//! it just escapes the text and writes it through. RFC-021 §8 risk #2.

use std::collections::BTreeMap;

use quick_xml::events::Event;
use quick_xml::Reader as XmlReader;

// ---------------------------------------------------------------------------
// Public types
// ---------------------------------------------------------------------------

/// One user-supplied defined-name mutation.
///
/// Every field except `name` and `formula` is optional for upserts.
/// `formula` is the XML text content (no leading `=`, openpyxl strips it on
/// the Python side). `local_sheet_id` is the 0-based sheet *position index*
/// (NOT a sheet name); `None` means workbook-scope.
///
/// On update of an existing name, attributes that the user did NOT provide
/// (`Option::None`) are preserved verbatim from the source XML. Phase 2
/// (G22) extended the explicit field surface to cover every ECMA-376
/// `definedName` attribute openpyxl exposes; truly unknown attributes
/// (forward-compat) still fall through the verbatim-passthrough path in
/// [`serialize_upsert_over_existing`].
///
/// When `delete` is true, the mutation removes the matching `(name,
/// local_sheet_id)` entry if it exists and ignores all attribute/formula
/// fields. Deleting a missing entry is a no-op.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct DefinedNameMut {
    pub name: String,
    pub delete: bool,
    /// XML text content. Stored verbatim modulo XML text-escape on emit.
    pub formula: String,
    /// `None` = workbook-scope; `Some(idx)` = sheet at 0-based position.
    pub local_sheet_id: Option<u32>,
    pub hidden: Option<bool>,
    pub comment: Option<String>,
    pub custom_menu: Option<String>,
    pub description: Option<String>,
    pub help: Option<String>,
    pub status_bar: Option<String>,
    pub shortcut_key: Option<String>,
    pub function: Option<bool>,
    pub function_group_id: Option<u32>,
    pub vb_procedure: Option<bool>,
    pub xlm: Option<bool>,
    pub publish_to_server: Option<bool>,
    pub workbook_parameter: Option<bool>,
}

/// One existing `<definedName>` entry parsed out of the source XML.
///
/// `raw` is the verbatim byte slice (start tag through end tag, or the
/// self-closing form) so entries that aren't being upserted flow through
/// untouched. Used internally; exposed for testability of the parser.
#[derive(Debug, Clone)]
pub(crate) struct ExistingDefinedName {
    pub name: String,
    pub local_sheet_id: Option<u32>,
    pub raw: Vec<u8>,
}

// ---------------------------------------------------------------------------
// merge_defined_names — public entry point
// ---------------------------------------------------------------------------

/// Parse `workbook_xml`, merge `names` (upsert by `(name, local_sheet_id)`),
/// and return the updated XML bytes.
///
/// - Existing names not referenced in `names` are preserved verbatim.
/// - Existing names referenced in `names` have their formula and any
///   explicitly-set optional attributes overwritten; other attributes
///   (and order in the block) are preserved.
/// - Names in `names` that don't match any existing entry are appended
///   to the end of the `<definedNames>` block.
/// - When the source has no `<definedNames>` block, one is inserted
///   immediately after `</sheets>` (per ECMA-376 §18.2.27 child ordering).
///
/// Returns `Err` on malformed XML (no `<workbook>`/`<sheets>` block);
/// well-formed inputs always succeed even when `names` is empty (the
/// idempotent identity transform).
pub fn merge_defined_names(
    workbook_xml: &[u8],
    names: &[DefinedNameMut],
) -> Result<Vec<u8>, String> {
    // Locate the existing `<definedNames>` block (if any) and the position
    // where a new block should be inserted (after `</sheets>` end tag).
    let layout = scan_workbook_layout(workbook_xml)?;

    // Empty upsert list → identity. Avoids the byte-shape change that
    // would happen if we re-emitted an existing block (whitespace
    // between children is not captured in `extract_defined_name_children`
    // and would be dropped on re-emit). Modify-mode contract: no
    // mutations queued ⇒ source bytes survive verbatim.
    if names.is_empty() {
        return Ok(workbook_xml.to_vec());
    }

    // Pull the existing entries out of any source block so we can
    // upsert by (name, local_sheet_id).
    let existing_entries: Vec<ExistingDefinedName> = match layout.defined_names_inner {
        Some((inner_start, inner_end)) => {
            extract_defined_name_children(&workbook_xml[inner_start..inner_end])
        }
        None => Vec::new(),
    };

    // Index user upserts by key. BTreeMap → deterministic iteration if
    // we ever surface "names that didn't match an existing entry".
    let mut pending: BTreeMap<(String, Option<u32>), &DefinedNameMut> = BTreeMap::new();
    for n in names {
        pending.insert((n.name.clone(), n.local_sheet_id), n);
    }

    // Pass 1: walk existing entries in source order. If there's a pending
    // mutation for an entry, either drop it (delete) or replace its bytes
    // with a freshly serialized form (upsert, preserving attributes the
    // upsert didn't override). Otherwise pass through verbatim.
    let mut merged_inner: Vec<u8> = Vec::with_capacity(256);
    for ex in &existing_entries {
        let key = (ex.name.clone(), ex.local_sheet_id);
        if let Some(mutation) = pending.remove(&key) {
            if mutation.delete {
                continue;
            }
            // Re-serialize with overrides applied to the source attrs.
            let serialized = serialize_upsert_over_existing(&ex.raw, mutation);
            merged_inner.extend_from_slice(&serialized);
        } else {
            merged_inner.extend_from_slice(&ex.raw);
        }
    }

    // Pass 2: emit any remaining upserts (new names) at the end of the block.
    // Deletes for missing names are no-ops. BTreeMap order keeps this
    // deterministic.
    let defined_name_child_name = layout.defined_name_child_name();
    for ((_name, _scope), mutation) in &pending {
        if mutation.delete {
            continue;
        }
        serialize_new_defined_name(
            &mut merged_inner,
            mutation,
            defined_name_child_name.as_bytes(),
        );
    }

    // Empty merged block + no existing block + nothing to emit → identity.
    if merged_inner.is_empty() && layout.defined_names_outer.is_none() {
        return Ok(workbook_xml.to_vec());
    }

    // Splice the merged block back into the workbook XML.
    let block_with_wrapper = if merged_inner.is_empty() {
        // All existing entries were deleted.
        Vec::new()
    } else {
        let mut wrapped: Vec<u8> = Vec::with_capacity(merged_inner.len() + 32);
        let wrapper_name = layout
            .defined_names_name
            .as_deref()
            .unwrap_or("definedNames");
        wrapped.push(b'<');
        wrapped.extend_from_slice(wrapper_name.as_bytes());
        wrapped.push(b'>');
        wrapped.extend_from_slice(&merged_inner);
        wrapped.extend_from_slice(b"</");
        wrapped.extend_from_slice(wrapper_name.as_bytes());
        wrapped.push(b'>');
        wrapped
    };

    let mut out: Vec<u8> = Vec::with_capacity(workbook_xml.len() + block_with_wrapper.len());
    match layout.defined_names_outer {
        Some((outer_start, outer_end)) => {
            // Replace the existing block in place.
            out.extend_from_slice(&workbook_xml[..outer_start]);
            out.extend_from_slice(&block_with_wrapper);
            out.extend_from_slice(&workbook_xml[outer_end..]);
        }
        None => {
            // Inject after `</sheets>`.
            let inject_at = layout.sheets_end;
            out.extend_from_slice(&workbook_xml[..inject_at]);
            out.extend_from_slice(&block_with_wrapper);
            out.extend_from_slice(&workbook_xml[inject_at..]);
        }
    }

    Ok(out)
}

// ---------------------------------------------------------------------------
// Internal: scan the workbook layout to find splice positions.
// ---------------------------------------------------------------------------

#[derive(Debug, Default)]
struct WorkbookLayout {
    /// Byte offset just after the `</sheets>` end tag. Required.
    sheets_end: usize,
    /// `(start_of_<definedNames>, end_of_</definedNames>)` byte range
    /// covering the entire existing block, or `None` if none.
    defined_names_outer: Option<(usize, usize)>,
    /// `(start_of_inner_content, end_of_inner_content)` — the bytes
    /// between `<definedNames>` and `</definedNames>` exclusive of the
    /// wrapper tags themselves. `None` when no existing block.
    defined_names_inner: Option<(usize, usize)>,
    defined_names_name: Option<String>,
}

impl WorkbookLayout {
    fn defined_name_child_name(&self) -> String {
        self.defined_names_name
            .as_deref()
            .and_then(|name| name.strip_suffix("definedNames"))
            .map(|prefix| format!("{prefix}definedName"))
            .unwrap_or_else(|| "definedName".to_string())
    }
}

fn scan_workbook_layout(xml: &[u8]) -> Result<WorkbookLayout, String> {
    let s =
        std::str::from_utf8(xml).map_err(|e| format!("workbook.xml is not valid UTF-8: {e}"))?;
    let mut reader = XmlReader::from_str(s);
    reader.config_mut().trim_text(false);
    let mut buf: Vec<u8> = Vec::new();

    let mut sheets_end: Option<usize> = None;
    let mut dn_start: Option<usize> = None;
    let mut dn_inner_start: Option<usize> = None;
    let mut dn_inner_end: Option<usize> = None;
    let mut dn_outer_end: Option<usize> = None;
    let mut dn_name: Option<String> = None;
    let mut dn_depth: u32 = 0;

    loop {
        let pre = reader.buffer_position() as usize;
        let evt = reader.read_event_into(&mut buf);
        let post = reader.buffer_position() as usize;

        match evt {
            Ok(Event::Start(ref e)) => {
                if e.local_name().as_ref() == "definedNames" && dn_start.is_none() {
                    let name = e.name().as_ref().as_bytes().to_vec();
                    let start = find_start_event_start(xml, pre, post, &name).unwrap_or(pre);
                    let end = find_tag_end(xml, start, post).unwrap_or(post);
                    dn_name = Some(String::from_utf8_lossy(&name).into_owned());
                    dn_start = Some(start);
                    dn_inner_start = Some(end);
                    dn_depth = 1;
                } else if dn_start.is_some() && e.local_name().as_ref() == "definedNames" {
                    dn_depth += 1;
                }
            }
            Ok(Event::Empty(ref e)) => {
                if e.local_name().as_ref() == "definedNames" && dn_start.is_none() {
                    // `<definedNames/>` self-closing — empty existing block.
                    let name = e.name().as_ref().as_bytes().to_vec();
                    let start = find_start_event_start(xml, pre, post, &name).unwrap_or(pre);
                    let end = find_tag_end(xml, start, post).unwrap_or(post);
                    dn_name = Some(String::from_utf8_lossy(&name).into_owned());
                    dn_start = Some(start);
                    dn_inner_start = Some(end);
                    dn_inner_end = Some(end);
                    dn_outer_end = Some(end);
                }
            }
            Ok(Event::End(ref e)) => {
                let local = e.local_name();
                if local.as_ref() == "sheets" && sheets_end.is_none() {
                    let name = e.name().as_ref().as_bytes().to_vec();
                    sheets_end = find_end_event_start(xml, pre, post, &name)
                        .and_then(|start| find_tag_end(xml, start, post))
                        .or(Some(post));
                } else if local.as_ref() == "definedNames" && dn_start.is_some() {
                    if dn_depth > 0 {
                        dn_depth -= 1;
                    }
                    if dn_depth == 0 && dn_outer_end.is_none() {
                        let name = dn_name
                            .as_ref()
                            .map(|n| n.as_bytes().to_vec())
                            .unwrap_or_else(|| e.name().as_ref().as_bytes().to_vec());
                        let start = find_end_event_start(xml, pre, post, &name).unwrap_or(pre);
                        let end = find_tag_end(xml, start, post).unwrap_or(post);
                        dn_inner_end = Some(start);
                        dn_outer_end = Some(end);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(format!("workbook.xml parse error: {e}")),
            _ => {}
        }
        buf.clear();
    }

    let sheets_end =
        sheets_end.ok_or_else(|| "workbook.xml has no </sheets> closing tag".to_string())?;
    let outer = match (dn_start, dn_outer_end) {
        (Some(s), Some(e)) => Some((s, e)),
        _ => None,
    };
    let inner = match (dn_inner_start, dn_inner_end) {
        (Some(s), Some(e)) => Some((s, e)),
        _ => None,
    };
    Ok(WorkbookLayout {
        sheets_end,
        defined_names_outer: outer,
        defined_names_inner: inner,
        defined_names_name: dn_name,
    })
}

fn find_start_event_start(
    xml: &[u8],
    event_pre: usize,
    event_post: usize,
    event_name: &[u8],
) -> Option<usize> {
    let end = event_post
        .saturating_add(event_name.len())
        .saturating_add(3)
        .min(xml.len());
    let mut i = event_pre.min(end);
    while i + event_name.len() + 1 <= end {
        if xml.get(i) == Some(&b'<')
            && xml.get(i + 1) != Some(&b'/')
            && xml.get(i + 1..i + 1 + event_name.len()) == Some(event_name)
        {
            return Some(i);
        }
        i += 1;
    }
    None
}

fn find_end_event_start(
    xml: &[u8],
    event_pre: usize,
    event_post: usize,
    event_name: &[u8],
) -> Option<usize> {
    let end = event_post
        .saturating_add(event_name.len())
        .saturating_add(3)
        .min(xml.len());
    let mut i = event_pre.min(end);
    while i + event_name.len() + 2 <= end {
        if xml.get(i) == Some(&b'<')
            && xml.get(i + 1) == Some(&b'/')
            && xml.get(i + 2..i + 2 + event_name.len()) == Some(event_name)
        {
            return Some(i);
        }
        i += 1;
    }
    None
}

fn find_tag_end(xml: &[u8], tag_start: usize, event_post: usize) -> Option<usize> {
    let search_end = event_post.saturating_add(64).min(xml.len());
    let mut i = tag_start.min(search_end);
    while i < search_end {
        if xml.get(i) == Some(&b'>') {
            return Some(i + 1);
        }
        i += 1;
    }
    None
}

// ---------------------------------------------------------------------------
// Internal: extract `<definedName>` children from the inner block bytes.
// ---------------------------------------------------------------------------

pub(crate) fn extract_defined_name_children(inner: &[u8]) -> Vec<ExistingDefinedName> {
    let s = match std::str::from_utf8(inner) {
        Ok(s) => s,
        Err(_) => return Vec::new(),
    };
    let mut reader = XmlReader::from_str(s);
    reader.config_mut().trim_text(false);
    let mut buf: Vec<u8> = Vec::new();

    let mut out: Vec<ExistingDefinedName> = Vec::new();
    let mut child_start: Option<usize> = None;
    let mut current_name: String = String::new();
    let mut current_local_id: Option<u32> = None;

    loop {
        let pre = reader.buffer_position() as usize;
        let evt = reader.read_event_into(&mut buf);
        let post = reader.buffer_position() as usize;

        match &evt {
            Ok(Event::Start(e)) if e.local_name().as_ref() == "definedName" => {
                child_start = Some(pre);
                let (name, local_id) = parse_defined_name_attrs(e);
                current_name = name;
                current_local_id = local_id;
            }
            Ok(Event::Empty(e)) if e.local_name().as_ref() == "definedName" => {
                let (name, local_id) = parse_defined_name_attrs(e);
                out.push(ExistingDefinedName {
                    name,
                    local_sheet_id: local_id,
                    raw: inner[pre..post].to_vec(),
                });
            }
            Ok(Event::End(e)) if e.local_name().as_ref() == "definedName" => {
                if let Some(start) = child_start.take() {
                    out.push(ExistingDefinedName {
                        name: std::mem::take(&mut current_name),
                        local_sheet_id: current_local_id.take(),
                        raw: inner[start..post].to_vec(),
                    });
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

fn parse_defined_name_attrs(e: &quick_xml::events::BytesStart<'_>) -> (String, Option<u32>) {
    let mut name = String::new();
    let mut local_id: Option<u32> = None;
    for a in e.attributes().with_checks(false).flatten() {
        let key = a.key.as_ref().as_bytes();
        let val = quick_xml::escape::unescape(&a.value)
            .map(|v| v.into_owned())
            .unwrap_or_else(|_| a.value.to_string());
        match key {
            b"name" => name = val,
            b"localSheetId" => local_id = val.parse::<u32>().ok(),
            _ => {}
        }
    }
    (name, local_id)
}

// ---------------------------------------------------------------------------
// Internal: serialize a brand-new `<definedName>` element from a `DefinedNameMut`.
// ---------------------------------------------------------------------------

fn serialize_new_defined_name(out: &mut Vec<u8>, dn: &DefinedNameMut, element_name: &[u8]) {
    out.push(b'<');
    out.extend_from_slice(element_name);
    out.extend_from_slice(b" name=\"");
    push_xml_attr_escape(out, &dn.name);
    out.push(b'"');
    if let Some(idx) = dn.local_sheet_id {
        out.extend_from_slice(format!(" localSheetId=\"{idx}\"").as_bytes());
    }
    // Attribute order mirrors `wolfxl-writer::emit::workbook_xml::emit_user_defined_names`
    // so write-mode and modify-mode produce byte-identical `<definedName>` shapes.
    push_opt_str_attr(out, b"comment", dn.comment.as_deref());
    push_opt_str_attr(out, b"customMenu", dn.custom_menu.as_deref());
    push_opt_str_attr(out, b"description", dn.description.as_deref());
    push_opt_str_attr(out, b"help", dn.help.as_deref());
    push_opt_str_attr(out, b"statusBar", dn.status_bar.as_deref());
    push_opt_str_attr(out, b"shortcutKey", dn.shortcut_key.as_deref());
    if dn.hidden == Some(true) {
        out.extend_from_slice(b" hidden=\"1\"");
    }
    push_opt_bool_true_attr(out, b"function", dn.function);
    push_opt_bool_true_attr(out, b"vbProcedure", dn.vb_procedure);
    push_opt_bool_true_attr(out, b"xlm", dn.xlm);
    if let Some(id) = dn.function_group_id {
        out.extend_from_slice(format!(" functionGroupId=\"{id}\"").as_bytes());
    }
    push_opt_bool_true_attr(out, b"publishToServer", dn.publish_to_server);
    push_opt_bool_true_attr(out, b"workbookParameter", dn.workbook_parameter);
    out.push(b'>');
    push_xml_text_escape(out, &dn.formula);
    out.extend_from_slice(b"</");
    out.extend_from_slice(element_name);
    out.push(b'>');
}

fn push_opt_str_attr(out: &mut Vec<u8>, key: &[u8], value: Option<&str>) {
    if let Some(v) = value {
        out.push(b' ');
        out.extend_from_slice(key);
        out.extend_from_slice(b"=\"");
        push_xml_attr_escape(out, v);
        out.push(b'"');
    }
}

fn push_opt_bool_true_attr(out: &mut Vec<u8>, key: &[u8], value: Option<bool>) {
    if value == Some(true) {
        out.push(b' ');
        out.extend_from_slice(key);
        out.extend_from_slice(b"=\"1\"");
    }
}

// ---------------------------------------------------------------------------
// Internal: rebuild a `<definedName>` element by overlaying overrides from
// `DefinedNameMut` onto the original source attributes. Attributes the
// upsert didn't override are preserved.
// ---------------------------------------------------------------------------

fn serialize_upsert_over_existing(raw: &[u8], upsert: &DefinedNameMut) -> Vec<u8> {
    let s = match std::str::from_utf8(raw) {
        Ok(s) => s,
        Err(_) => {
            // Pathological — fall back to a fresh emit.
            let mut out = Vec::new();
            serialize_new_defined_name(&mut out, upsert, b"definedName");
            return out;
        }
    };

    // Pull all attributes out of the start tag (or self-closing form) so
    // we can preserve `customMenu`, `description`, etc.
    let mut reader = XmlReader::from_str(s);
    reader.config_mut().trim_text(false);
    let mut buf: Vec<u8> = Vec::new();

    let mut existing_attrs: Vec<(Vec<u8>, String)> = Vec::new();
    let mut element_name: Vec<u8> = b"definedName".to_vec();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e))
                if e.local_name().as_ref() == "definedName" =>
            {
                element_name = e.name().as_ref().as_bytes().to_vec();
                for a in e.attributes().with_checks(false).flatten() {
                    let key = a.key.as_ref().as_bytes().to_vec();
                    let val = quick_xml::escape::unescape(&a.value)
                        .map(|v| v.into_owned())
                        .unwrap_or_else(|_| a.value.to_string());
                    existing_attrs.push((key, val));
                }
                break;
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    // Apply overrides. For each override, replace the matching attr (if
    // present) or append it. `name` always reflects the upsert's name
    // (which equals the existing entry's name by construction — same key).
    fn upsert_attr(attrs: &mut Vec<(Vec<u8>, String)>, key: &[u8], val: String) {
        if let Some(slot) = attrs.iter_mut().find(|(k, _)| k.as_slice() == key) {
            slot.1 = val;
        } else {
            attrs.push((key.to_vec(), val));
        }
    }
    fn remove_attr(attrs: &mut Vec<(Vec<u8>, String)>, key: &[u8]) {
        attrs.retain(|(k, _)| k.as_slice() != key);
    }

    fn upsert_opt_str(attrs: &mut Vec<(Vec<u8>, String)>, key: &[u8], value: Option<&str>) {
        if let Some(v) = value {
            upsert_attr(attrs, key, v.to_string());
        }
    }
    fn upsert_opt_bool_true(attrs: &mut Vec<(Vec<u8>, String)>, key: &[u8], value: Option<bool>) {
        match value {
            Some(true) => upsert_attr(attrs, key, "1".to_string()),
            Some(false) => remove_attr(attrs, key),
            None => { /* preserve source */ }
        }
    }

    upsert_attr(&mut existing_attrs, b"name", upsert.name.clone());
    match upsert.local_sheet_id {
        Some(idx) => upsert_attr(&mut existing_attrs, b"localSheetId", idx.to_string()),
        // None override: source key matched (None == None), so just
        // ensure the attribute is absent.
        None => remove_attr(&mut existing_attrs, b"localSheetId"),
    }
    match upsert.hidden {
        Some(true) => upsert_attr(&mut existing_attrs, b"hidden", "1".to_string()),
        Some(false) => remove_attr(&mut existing_attrs, b"hidden"),
        None => { /* preserve source */ }
    }
    upsert_opt_str(&mut existing_attrs, b"comment", upsert.comment.as_deref());
    upsert_opt_str(
        &mut existing_attrs,
        b"customMenu",
        upsert.custom_menu.as_deref(),
    );
    upsert_opt_str(
        &mut existing_attrs,
        b"description",
        upsert.description.as_deref(),
    );
    upsert_opt_str(&mut existing_attrs, b"help", upsert.help.as_deref());
    upsert_opt_str(
        &mut existing_attrs,
        b"statusBar",
        upsert.status_bar.as_deref(),
    );
    upsert_opt_str(
        &mut existing_attrs,
        b"shortcutKey",
        upsert.shortcut_key.as_deref(),
    );
    upsert_opt_bool_true(&mut existing_attrs, b"function", upsert.function);
    upsert_opt_bool_true(&mut existing_attrs, b"vbProcedure", upsert.vb_procedure);
    upsert_opt_bool_true(&mut existing_attrs, b"xlm", upsert.xlm);
    if let Some(id) = upsert.function_group_id {
        upsert_attr(&mut existing_attrs, b"functionGroupId", id.to_string());
    }
    upsert_opt_bool_true(
        &mut existing_attrs,
        b"publishToServer",
        upsert.publish_to_server,
    );
    upsert_opt_bool_true(
        &mut existing_attrs,
        b"workbookParameter",
        upsert.workbook_parameter,
    );

    // Re-emit the element. Attribute order: keep original order for
    // attrs that existed, then append any newly added ones.
    let mut out: Vec<u8> = Vec::with_capacity(raw.len() + upsert.formula.len());
    out.push(b'<');
    out.extend_from_slice(&element_name);
    for (key, val) in &existing_attrs {
        out.push(b' ');
        out.extend_from_slice(key);
        out.extend_from_slice(b"=\"");
        push_xml_attr_escape(&mut out, val);
        out.push(b'"');
    }
    out.push(b'>');
    push_xml_text_escape(&mut out, &upsert.formula);
    out.extend_from_slice(b"</");
    out.extend_from_slice(&element_name);
    out.push(b'>');
    out
}

// ---------------------------------------------------------------------------
// XML escape helpers — local copies of the writer's helpers. Identical
// semantics so write-mode and modify-mode produce the same byte shape.
// ---------------------------------------------------------------------------

fn push_xml_text_escape(out: &mut Vec<u8>, s: &str) {
    for ch in s.chars() {
        match ch {
            '&' => out.extend_from_slice(b"&amp;"),
            '<' => out.extend_from_slice(b"&lt;"),
            '>' => out.extend_from_slice(b"&gt;"),
            _ => {
                let mut b = [0u8; 4];
                out.extend_from_slice(ch.encode_utf8(&mut b).as_bytes());
            }
        }
    }
}

fn push_xml_attr_escape(out: &mut Vec<u8>, s: &str) {
    for ch in s.chars() {
        match ch {
            '&' => out.extend_from_slice(b"&amp;"),
            '<' => out.extend_from_slice(b"&lt;"),
            '>' => out.extend_from_slice(b"&gt;"),
            '"' => out.extend_from_slice(b"&quot;"),
            '\'' => out.extend_from_slice(b"&apos;"),
            _ => {
                let mut b = [0u8; 4];
                out.extend_from_slice(ch.encode_utf8(&mut b).as_bytes());
            }
        }
    }
}

// ---------------------------------------------------------------------------
// Tests
//
// Inline pure-Rust tests. The patcher's cdylib does not link standalone via
// `cargo test -p wolfxl --lib` (Python linkage), so these compile under
// `cargo build` and end-to-end behavior is exercised via pytest. Same
// precedent as `properties.rs` and `hyperlinks.rs`.
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;

    fn workbook_xml_no_defined_names() -> &'static str {
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <fileVersion appName="xl"/>
  <workbookPr/>
  <bookViews><workbookView/></bookViews>
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    <sheet name="Sheet2" sheetId="2" r:id="rId2"/>
  </sheets>
  <calcPr calcId="171027"/>
</workbook>"#
    }

    fn workbook_xml_with_defined_names() -> &'static str {
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
  <definedNames>
    <definedName name="Region">Sheet1!$A$1:$A$10</definedName>
    <definedName name="_xlnm.Print_Area" localSheetId="0">Sheet1!$A$1:$D$20</definedName>
  </definedNames>
  <calcPr/>
</workbook>"#
    }

    #[test]
    fn merge_into_xml_with_no_defined_names_inserts_block_after_sheets() {
        let xml = workbook_xml_no_defined_names();
        let names = vec![DefinedNameMut {
            name: "Budget".into(),
            formula: "Sheet1!$A$1:$A$100".into(),
            ..Default::default()
        }];
        let bytes = merge_defined_names(xml.as_bytes(), &names).expect("merge");
        let text = std::str::from_utf8(&bytes).unwrap();
        assert!(text.contains("<definedNames>"));
        assert!(text.contains("</definedNames>"));
        assert!(text.contains(r#"<definedName name="Budget">Sheet1!$A$1:$A$100</definedName>"#));
        let pos_sheets_end = text.find("</sheets>").unwrap();
        let pos_dn_start = text.find("<definedNames>").unwrap();
        let pos_calc = text.find("<calcPr").unwrap();
        assert!(pos_sheets_end < pos_dn_start && pos_dn_start < pos_calc);
    }

    #[test]
    fn merge_appends_to_existing_block_preserving_existing_entries() {
        let xml = workbook_xml_with_defined_names();
        let names = vec![DefinedNameMut {
            name: "Budget".into(),
            formula: "Sheet1!$B$1".into(),
            ..Default::default()
        }];
        let bytes = merge_defined_names(xml.as_bytes(), &names).expect("merge");
        let text = std::str::from_utf8(&bytes).unwrap();
        assert!(text.contains(r#"<definedName name="Region">Sheet1!$A$1:$A$10</definedName>"#));
        assert!(text.contains(r#"<definedName name="_xlnm.Print_Area" localSheetId="0">Sheet1!$A$1:$D$20</definedName>"#));
        assert!(text.contains(r#"<definedName name="Budget">Sheet1!$B$1</definedName>"#));
        assert_eq!(text.matches("<definedNames>").count(), 1);
    }

    #[test]
    fn merge_preserves_prefixed_defined_names_after_compact_sheets_block() {
        let xml = b"\xef\xbb\xbf<?xml version=\"1.0\" encoding=\"utf-8\" standalone=\"yes\"?><x:workbook xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><x:sheets><x:sheet name=\"BAL\" sheetId=\"4\" state=\"hidden\" r:id=\"rId3\" /><x:sheet name=\"COVER Copy\" sheetId=\"119\" r:id=\"rId74\"/></x:sheets><x:definedNames><x:definedName name=\"BAL_DATA\" localSheetId=\"0\">BAL!$A$1</x:definedName></x:definedNames><x:calcPr/></x:workbook>";
        let names = vec![DefinedNameMut {
            name: "CopyRange".into(),
            formula: "'COVER Copy'!$A$1".into(),
            local_sheet_id: Some(1),
            ..Default::default()
        }];
        let bytes = merge_defined_names(xml, &names).expect("merge");
        let text = std::str::from_utf8(&bytes).unwrap();
        assert!(text.contains("</x:sheets><x:definedNames>"));
        assert!(text.contains(
            r#"<x:definedName name="BAL_DATA" localSheetId="0">BAL!$A$1</x:definedName>"#
        ));
        assert!(text.contains(
            r#"<x:definedName name="CopyRange" localSheetId="1">'COVER Copy'!$A$1</x:definedName>"#
        ));
        assert!(!text.contains("</x:shee<definedNames>"));
        assert!(!text.contains("<definedNames>"));
    }

    #[test]
    fn merge_upsert_replaces_formula_preserves_other_attrs() {
        let xml = r#"<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheets><sheet name="Sheet1" sheetId="1"/></sheets>
  <definedNames>
    <definedName name="Region" comment="ignore-me">Sheet1!$A$1</definedName>
  </definedNames>
</workbook>"#;
        let names = vec![DefinedNameMut {
            name: "Region".into(),
            formula: "Sheet1!$Z$99".into(),
            ..Default::default()
        }];
        let bytes = merge_defined_names(xml.as_bytes(), &names).expect("merge");
        let text = std::str::from_utf8(&bytes).unwrap();
        assert!(text.contains("Sheet1!$Z$99"));
        assert!(!text.contains("Sheet1!$A$1<"));
        assert!(text.contains(r#"comment="ignore-me""#));
        assert_eq!(text.matches(r#"name="Region""#).count(), 1);
    }

    #[test]
    fn upsert_distinguishes_workbook_vs_sheet_scope() {
        let xml = r#"<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheets><sheet name="Sheet1" sheetId="1"/></sheets>
  <definedNames>
    <definedName name="Foo">Sheet1!$A$1</definedName>
    <definedName name="Foo" localSheetId="0">Sheet1!$B$1</definedName>
  </definedNames>
</workbook>"#;
        let names = vec![DefinedNameMut {
            name: "Foo".into(),
            formula: "Sheet1!$AA$1".into(),
            local_sheet_id: None,
            ..Default::default()
        }];
        let bytes = merge_defined_names(xml.as_bytes(), &names).expect("merge");
        let text = std::str::from_utf8(&bytes).unwrap();
        assert!(text.contains(r#"<definedName name="Foo">Sheet1!$AA$1</definedName>"#));
        assert!(
            text.contains(r#"<definedName name="Foo" localSheetId="0">Sheet1!$B$1</definedName>"#)
        );
        assert_eq!(text.matches(r#"name="Foo""#).count(), 2);
    }

    #[test]
    fn xml_special_chars_in_formula_are_escaped() {
        let xml = workbook_xml_no_defined_names();
        let names = vec![DefinedNameMut {
            name: "Weird".into(),
            formula: "Sheet1!$A$1 & \"quoted\" < other".into(),
            ..Default::default()
        }];
        let bytes = merge_defined_names(xml.as_bytes(), &names).expect("merge");
        let text = std::str::from_utf8(&bytes).unwrap();
        assert!(text.contains("&amp;"), "ampersand must be escaped");
        assert!(text.contains("&lt;"), "less-than must be escaped");
        assert!(text.contains("\"quoted\""));
    }

    #[test]
    fn empty_names_is_identity() {
        let xml = workbook_xml_with_defined_names();
        let bytes = merge_defined_names(xml.as_bytes(), &[]).expect("merge");
        // Identity: empty queue ⇒ source bytes survive verbatim.
        assert_eq!(bytes, xml.as_bytes());
    }

    #[test]
    fn empty_names_on_xml_with_no_block_is_identity() {
        let xml = workbook_xml_no_defined_names();
        let bytes = merge_defined_names(xml.as_bytes(), &[]).expect("merge");
        assert_eq!(bytes, xml.as_bytes());
    }

    #[test]
    fn builtin_print_area_round_trips() {
        let xml = workbook_xml_with_defined_names();
        let names = vec![DefinedNameMut {
            name: "Margin".into(),
            formula: "0.5".into(),
            ..Default::default()
        }];
        let bytes = merge_defined_names(xml.as_bytes(), &names).expect("merge");
        let text = std::str::from_utf8(&bytes).unwrap();
        assert!(text.contains(
            r#"<definedName name="_xlnm.Print_Area" localSheetId="0">Sheet1!$A$1:$D$20</definedName>"#
        ));
    }

    #[test]
    fn delete_existing_defined_name_preserves_unrelated_entries() {
        let xml = workbook_xml_with_defined_names();
        let names = vec![DefinedNameMut {
            name: "_xlnm.Print_Area".into(),
            local_sheet_id: Some(0),
            delete: true,
            ..Default::default()
        }];
        let bytes = merge_defined_names(xml.as_bytes(), &names).expect("merge");
        let text = std::str::from_utf8(&bytes).unwrap();
        assert!(text.contains(r#"<definedName name="Region">Sheet1!$A$1:$A$10</definedName>"#));
        assert!(!text.contains(r#"name="_xlnm.Print_Area""#));
        assert!(text.contains("<definedNames>"));
    }

    #[test]
    fn delete_only_existing_defined_name_removes_wrapper_block() {
        let xml = r#"<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheets><sheet name="Sheet1" sheetId="1"/></sheets>
  <definedNames><definedName name="Only">Sheet1!$A$1</definedName></definedNames>
  <calcPr/>
</workbook>"#;
        let names = vec![DefinedNameMut {
            name: "Only".into(),
            delete: true,
            ..Default::default()
        }];
        let bytes = merge_defined_names(xml.as_bytes(), &names).expect("merge");
        let text = std::str::from_utf8(&bytes).unwrap();
        assert!(!text.contains("<definedNames>"));
        assert!(text.contains("<calcPr/>"));
    }

    #[test]
    fn delete_missing_defined_name_is_identity_when_no_block_exists() {
        let xml = workbook_xml_no_defined_names();
        let names = vec![DefinedNameMut {
            name: "Missing".into(),
            delete: true,
            ..Default::default()
        }];
        let bytes = merge_defined_names(xml.as_bytes(), &names).expect("merge");
        assert_eq!(bytes, xml.as_bytes());
    }

    #[test]
    fn hidden_attr_emitted_when_true() {
        let xml = workbook_xml_no_defined_names();
        let names = vec![DefinedNameMut {
            name: "Internal".into(),
            formula: "Sheet1!$A$1".into(),
            hidden: Some(true),
            ..Default::default()
        }];
        let bytes = merge_defined_names(xml.as_bytes(), &names).expect("merge");
        let text = std::str::from_utf8(&bytes).unwrap();
        assert!(text.contains(r#"hidden="1""#));
    }

    #[test]
    fn local_sheet_id_emitted_for_sheet_scope() {
        let xml = workbook_xml_no_defined_names();
        let names = vec![DefinedNameMut {
            name: "S1Range".into(),
            formula: "Sheet2!$A$1".into(),
            local_sheet_id: Some(1),
            ..Default::default()
        }];
        let bytes = merge_defined_names(xml.as_bytes(), &names).expect("merge");
        let text = std::str::from_utf8(&bytes).unwrap();
        assert!(text
            .contains(r#"<definedName name="S1Range" localSheetId="1">Sheet2!$A$1</definedName>"#));
    }

    #[test]
    fn extract_children_handles_two_entries() {
        let inner = br#"
    <definedName name="A">x</definedName>
    <definedName name="B" localSheetId="2">y</definedName>
"#;
        let kids = extract_defined_name_children(inner);
        assert_eq!(kids.len(), 2);
        assert_eq!(kids[0].name, "A");
        assert_eq!(kids[0].local_sheet_id, None);
        assert_eq!(kids[1].name, "B");
        assert_eq!(kids[1].local_sheet_id, Some(2));
    }

    #[test]
    fn split_references_preserve_raw_defined_name_children() {
        let raw = br#"<definedName name="R&amp;D" localSheetId="&#50;"> 'A &amp; B'!$A$1 &amp; "&#x1F600;" </definedName>"#;
        let children = extract_defined_name_children(raw);
        assert_eq!(children.len(), 1);
        assert_eq!(children[0].name, "R&D");
        assert_eq!(children[0].local_sheet_id, Some(2));
        assert_eq!(children[0].raw, raw);
    }

    #[test]
    fn upsert_preserves_literal_attribute_whitespace() {
        let xml = b"<workbook><sheets></sheets><definedNames><definedName name=\"R&amp;D\" \
                    comment=\" a\tb\r\nc&#10; \">old</definedName></definedNames></workbook>";
        let names = [DefinedNameMut {
            name: "R&D".into(),
            formula: " A1 & \" B \" ".into(),
            ..Default::default()
        }];
        let output = merge_defined_names(xml, &names).unwrap();
        let children = extract_defined_name_children(&output);
        assert_eq!(children.len(), 1);
        assert_eq!(children[0].name, "R&D");
        let text = std::str::from_utf8(&children[0].raw).unwrap();
        assert!(text.contains("comment=\" a\tb\r\nc\n \""));
        assert!(text.contains(" A1 &amp; \" B \" "));
    }

    #[test]
    fn missing_sheets_close_tag_errors() {
        let xml = b"<?xml version=\"1.0\"?><workbook>no sheets here</workbook>";
        assert!(merge_defined_names(xml, &[]).is_err());
    }
}
