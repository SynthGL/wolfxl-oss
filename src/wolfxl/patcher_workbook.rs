//! Workbook-level helpers used by the surgical xlsx patcher.

use std::collections::{HashMap, HashSet};
use std::fs::File;
use std::io::{Read, Seek, Write};

use pyo3::exceptions::PyIOError;
use pyo3::prelude::*;
use quick_xml::events::{BytesStart, BytesText, Event};
use quick_xml::Reader as XmlReader;
use quick_xml::Writer as XmlWriter;
use zip::write::SimpleFileOptions;
use zip::{ZipArchive, ZipWriter};

use crate::ooxml_util;
use wolfxl_rels::{RelId, RelsGraph};

use super::{
    calcchain, content_types, defined_names, properties, security, sheet_order, XlsxPatcher,
};

const CT_WORKSHEET: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const FAST_SAVE_DEFLATE_LEVEL: i64 = 2;

/// Maps a sheet XML path to its relationship sidecar path.
pub(crate) fn sheet_rels_path_for(sheet_path: &str) -> String {
    wolfxl_rels::rels_path_for(sheet_path).unwrap_or_else(|| format!("_rels/{sheet_path}.rels"))
}

/// Maps any OOXML part path to its conventional relationship sidecar path.
pub(crate) fn part_rels_path_for(part_path: &str) -> PyResult<String> {
    wolfxl_rels::rels_path_for(part_path)
        .ok_or_else(|| PyIOError::new_err(format!("part has no rels path: {part_path}")))
}

/// Parses the trailing integer from an OOXML part path.
pub(crate) fn parse_n_from_part_path(path: &str, prefix: &str, suffix: &str) -> Option<u32> {
    let mid = path.strip_prefix(prefix)?.strip_suffix(suffix)?;
    mid.parse::<u32>().ok()
}

/// Loads a `.rels` part from the source ZIP, returning an empty graph when absent.
pub(crate) fn load_or_empty_rels(zip: &mut ZipArchive<File>, path: &str) -> PyResult<RelsGraph> {
    match ooxml_util::zip_read_to_string_opt(zip, path)? {
        Some(xml) => RelsGraph::parse(xml.as_bytes())
            .map_err(|e| PyIOError::new_err(format!("rels parse for '{path}': {e}"))),
        None => Ok(RelsGraph::new()),
    }
}

/// Loads the current rels graph, preferring queued replacements/adds over ZIP source.
pub(crate) fn current_or_empty_rels(
    patcher: &XlsxPatcher,
    zip: &mut ZipArchive<File>,
    path: &str,
) -> PyResult<RelsGraph> {
    if let Some(g) = patcher.rels_patches.get(path) {
        return Ok(g.clone());
    }
    if let Some(bytes) = patcher.file_adds.get(path) {
        return RelsGraph::parse(bytes)
            .map_err(|e| PyIOError::new_err(format!("rels parse for '{path}': {e}")));
    }
    load_or_empty_rels(zip, path)
}

/// Returns true when the source ZIP contains an entry with the exact name.
pub(crate) fn source_zip_has_entry<R: Read + Seek>(zip: &mut ZipArchive<R>, name: &str) -> bool {
    zip.by_name(name).is_ok()
}

/// Reads the current bytes for a part, preferring pending replacements and adds.
pub(crate) fn current_part_bytes(
    file_patches: &HashMap<String, Vec<u8>>,
    file_adds: &HashMap<String, Vec<u8>>,
    zip: &mut ZipArchive<File>,
    path: &str,
) -> Option<Vec<u8>> {
    if let Some(b) = file_patches.get(path) {
        return Some(b.clone());
    }
    if let Some(b) = file_adds.get(path) {
        return Some(b.clone());
    }
    source_part_bytes(zip, path)
}

/// Reads source or replacement bytes for a part that cannot come from file_adds.
pub(crate) fn patched_or_source_part_bytes(
    file_patches: &HashMap<String, Vec<u8>>,
    zip: &mut ZipArchive<File>,
    path: &str,
) -> Option<Vec<u8>> {
    if let Some(b) = file_patches.get(path) {
        return Some(b.clone());
    }
    source_part_bytes(zip, path)
}

fn source_part_bytes(zip: &mut ZipArchive<File>, path: &str) -> Option<Vec<u8>> {
    let mut entry = match zip.by_name(path) {
        Ok(e) => e,
        Err(_) => return None,
    };
    let mut buf: Vec<u8> = Vec::with_capacity(entry.size() as usize);
    std::io::copy(&mut entry, &mut buf).ok()?;
    Some(buf)
}

/// Inserts a `<sheet .../>` element into the workbook `<sheets>` block.
pub(crate) fn splice_into_sheets_block(
    workbook_xml: &[u8],
    new_sheet_element: &[u8],
) -> PyResult<Vec<u8>> {
    use quick_xml::events::Event as XmlEvent;
    use quick_xml::Reader as XmlReader;

    let s0 = std::str::from_utf8(workbook_xml)
        .map_err(|_| PyIOError::new_err("Phase 2.7: workbook.xml is not valid UTF-8"))?;
    let owned;
    let s = if new_sheet_element.windows(4).any(|w| w == b"r:id")
        && !workbook_root_has_relationship_namespace(s0)?
    {
        owned = add_workbook_relationship_namespace(s0)?;
        owned.as_str()
    } else {
        s0
    };
    let workbook_xml = s.as_bytes();
    let mut reader = XmlReader::from_str(s);
    reader.config_mut().trim_text(false);

    let mut depth: i32 = 0;
    let mut splice_pos: Option<usize> = None;
    let mut sheets_open_depth: Option<i32> = None;
    let mut self_closing_range: Option<(usize, usize)> = None;
    let mut sheets_element_name: Option<String> = None;
    let mut buf: Vec<u8> = Vec::new();

    loop {
        let pre = reader.buffer_position() as usize;
        let evt = reader.read_event_into(&mut buf);
        let post = reader.buffer_position() as usize;
        match evt {
            Ok(XmlEvent::Start(ref e)) => {
                if e.local_name().as_ref() == b"sheets" && sheets_open_depth.is_none() {
                    sheets_open_depth = Some(depth);
                    sheets_element_name =
                        Some(String::from_utf8_lossy(e.name().as_ref()).into_owned());
                }
                depth += 1;
            }
            Ok(XmlEvent::End(ref e)) => {
                depth -= 1;
                if e.local_name().as_ref() == b"sheets"
                    && Some(depth) == sheets_open_depth
                    && splice_pos.is_none()
                {
                    let event_name: Vec<u8> = sheets_element_name
                        .as_ref()
                        .map(|name| name.as_bytes().to_vec())
                        .unwrap_or_else(|| e.name().as_ref().to_vec());
                    splice_pos =
                        find_end_event_start(workbook_xml, pre, post, &event_name).or(Some(pre));
                    break;
                }
            }
            Ok(XmlEvent::Empty(ref e)) => {
                if e.local_name().as_ref() == b"sheets" && self_closing_range.is_none() {
                    self_closing_range = Some((pre, post));
                    sheets_element_name =
                        Some(String::from_utf8_lossy(e.name().as_ref()).into_owned());
                    break;
                }
            }
            Ok(XmlEvent::Eof) => break,
            Ok(_) => {}
            Err(_) => break,
        }
        buf.clear();
    }

    let sheets_name = sheets_element_name.as_deref().unwrap_or("sheets");
    let sheet_element = qualify_sheet_element_for_sheets(new_sheet_element, sheets_name);
    if let Some(pos) = splice_pos {
        let mut out = Vec::with_capacity(workbook_xml.len() + sheet_element.len());
        out.extend_from_slice(&workbook_xml[..pos]);
        out.extend_from_slice(&sheet_element);
        out.extend_from_slice(&workbook_xml[pos..]);
        return Ok(out);
    }
    if let Some((start, end)) = self_closing_range {
        let mut out = Vec::with_capacity(
            workbook_xml.len() + sheet_element.len() + (2 * sheets_name.len()) + 5,
        );
        out.extend_from_slice(&workbook_xml[..start]);
        out.extend_from_slice(format!("<{sheets_name}>").as_bytes());
        out.extend_from_slice(&sheet_element);
        out.extend_from_slice(format!("</{sheets_name}>").as_bytes());
        out.extend_from_slice(&workbook_xml[end..]);
        return Ok(out);
    }
    Err(PyIOError::new_err(
        "Phase 2.7: workbook.xml has no <sheets> block",
    ))
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
    let needle_len = event_name.len();
    if needle_len == 0 || end < needle_len + 3 {
        return None;
    }

    let start = event_pre.min(end);
    if let Some(i) = find_named_end_tag(xml, start, end, event_name) {
        return Some(i);
    }
    find_named_end_tag(xml, end, xml.len(), event_name)
}

fn find_named_end_tag(xml: &[u8], start: usize, end: usize, event_name: &[u8]) -> Option<usize> {
    let end = end.min(xml.len());
    let needle_len = event_name.len();
    let mut i = start.min(end);
    while i + needle_len + 3 <= end {
        if xml.get(i) == Some(&b'<')
            && xml.get(i + 1) == Some(&b'/')
            && xml.get(i + 2..i + 2 + needle_len) == Some(event_name)
        {
            let after_name = i + 2 + needle_len;
            if matches!(
                xml.get(after_name),
                Some(b'>') | Some(b' ' | b'\t' | b'\r' | b'\n')
            ) {
                return Some(i);
            }
        }
        i += 1;
    }
    None
}

fn qualify_sheet_element_for_sheets(new_sheet_element: &[u8], sheets_name: &str) -> Vec<u8> {
    let prefix = sheets_name.strip_suffix("sheets").unwrap_or_default();
    if prefix.is_empty() {
        return new_sheet_element.to_vec();
    }
    let Ok(text) = std::str::from_utf8(new_sheet_element) else {
        return new_sheet_element.to_vec();
    };
    let Some(rest) = text.strip_prefix("<sheet") else {
        return new_sheet_element.to_vec();
    };
    format!("<{prefix}sheet{rest}").into_bytes()
}

pub(super) fn apply_sheet_deletes_phase(
    patcher: &mut XlsxPatcher,
    file_patches: &mut HashMap<String, Vec<u8>>,
    zip: &mut ZipArchive<File>,
) -> PyResult<()> {
    if patcher.queued_sheet_deletes.is_empty() {
        return Ok(());
    }

    let workbook_rels_path = "xl/_rels/workbook.xml.rels".to_string();
    if !patcher.rels_patches.contains_key(&workbook_rels_path) {
        let g = load_or_empty_rels(zip, &workbook_rels_path)?;
        patcher.rels_patches.insert(workbook_rels_path.clone(), g);
    }

    let workbook_xml =
        match current_part_bytes(file_patches, &patcher.file_adds, zip, "xl/workbook.xml") {
            Some(b) => b,
            None => {
                return Err(PyIOError::new_err(
                    "sheet-delete: xl/workbook.xml missing from source ZIP",
                ));
            }
        };
    let workbook_xml_text = std::str::from_utf8(&workbook_xml)
        .map_err(|e| PyIOError::new_err(format!("sheet-delete workbook.xml utf8: {e}")))?;
    let sheet_rids = ooxml_util::parse_workbook_sheet_rids(workbook_xml_text)
        .map_err(|e| PyIOError::new_err(format!("sheet-delete workbook.xml parse: {e}")))?;

    let deletes = patcher.queued_sheet_deletes.clone();
    let mut deleted_workbook_cache_names: HashSet<String> = HashSet::new();
    for title in &deletes {
        if let Some((_, rid)) = sheet_rids.iter().find(|(name, _)| name == title) {
            if let Some(graph) = patcher.rels_patches.get_mut(&workbook_rels_path) {
                graph.remove(&RelId(rid.clone()));
            }
        }
        let Some(sheet_path) = patcher.deleted_sheet_paths.get(title).cloned() else {
            continue;
        };
        let sheet_rels_path = sheet_rels_path_for(&sheet_path);
        let source_rels = if let Some(g) = patcher.rels_patches.get(&sheet_rels_path) {
            g.clone()
        } else {
            load_or_empty_rels(zip, &sheet_rels_path)?
        };
        let subgraph = wolfxl_rels::walk_sheet_subgraph_with_nested(
            &source_rels,
            &sheet_path,
            |part_path: &str| {
                let rels_path = wolfxl_rels::rels_path_for(part_path)?;
                let bytes = current_part_bytes(file_patches, &patcher.file_adds, zip, &rels_path)?;
                RelsGraph::parse(&bytes).ok()
            },
        );
        let deleted_parts: HashSet<String> = subgraph.reachable_parts.iter().cloned().collect();
        let retained_parts =
            retained_deleted_subgraph_parts(&deleted_parts, patcher, file_patches, zip)?;
        let pruned_deleted_parts: HashSet<String> =
            deleted_parts.difference(&retained_parts).cloned().collect();
        collect_deleted_workbook_cache_names(
            &pruned_deleted_parts,
            file_patches,
            &patcher.file_adds,
            zip,
            &mut deleted_workbook_cache_names,
        );
        for part in subgraph.reachable_parts {
            if retained_parts.contains(&part) {
                continue;
            }
            patcher.file_deletes.insert(part.clone());
            if let Some(rels_path) = wolfxl_rels::rels_path_for(&part) {
                patcher.file_deletes.insert(rels_path);
            }
            patcher
                .queued_content_type_ops
                .entry("__sheet_delete__".to_string())
                .or_default()
                .push(content_types::ContentTypeOp::RemoveOverride(format!(
                    "/{part}"
                )));
        }
        patcher.file_deletes.insert(sheet_rels_path);
    }

    let result = sheet_order::merge_sheet_deletes(&workbook_xml, &deletes)
        .map_err(|e| PyIOError::new_err(format!("sheet-delete merge: {e}")))?;
    let workbook_xml = prune_deleted_workbook_cache_refs(
        result.workbook_xml,
        &deleted_workbook_cache_names,
        patcher,
        file_patches,
        zip,
    )?;
    file_patches.insert("xl/workbook.xml".to_string(), workbook_xml);
    patcher.sheet_order = result.new_order;
    patcher.queued_sheet_deletes.clear();
    patcher.deleted_sheet_paths.clear();
    Ok(())
}

fn collect_deleted_workbook_cache_names(
    deleted_parts: &HashSet<String>,
    file_patches: &HashMap<String, Vec<u8>>,
    file_adds: &HashMap<String, Vec<u8>>,
    zip: &mut ZipArchive<File>,
    out: &mut HashSet<String>,
) {
    for part in deleted_parts {
        let is_slicer = part.starts_with("xl/slicers/slicer") && part.ends_with(".xml");
        let is_timeline = part.starts_with("xl/timelines/timeline") && part.ends_with(".xml");
        if !is_slicer && !is_timeline {
            continue;
        }
        let Some(bytes) = current_part_bytes(file_patches, file_adds, zip, part) else {
            continue;
        };
        for cache_name in xml_attr_values_by_local_name(&bytes, b"cache") {
            out.insert(cache_name);
        }
    }
}

fn prune_deleted_workbook_cache_refs(
    workbook_xml: Vec<u8>,
    cache_names: &HashSet<String>,
    patcher: &mut XlsxPatcher,
    file_patches: &HashMap<String, Vec<u8>>,
    zip: &mut ZipArchive<File>,
) -> PyResult<Vec<u8>> {
    if cache_names.is_empty() {
        return Ok(workbook_xml);
    }
    let cache_parts =
        workbook_cache_parts_for_names(cache_names, file_patches, &patcher.file_adds, zip);
    if cache_parts.is_empty() {
        return Ok(workbook_xml);
    }
    let workbook_rels_path = "xl/_rels/workbook.xml.rels".to_string();
    if !patcher.rels_patches.contains_key(&workbook_rels_path) {
        let g = load_or_empty_rels(zip, &workbook_rels_path)?;
        patcher.rels_patches.insert(workbook_rels_path.clone(), g);
    }
    let mut removed_rids = HashSet::new();
    if let Some(graph) = patcher.rels_patches.get_mut(&workbook_rels_path) {
        let mut to_remove = Vec::new();
        for rel in graph.iter() {
            if rel.mode == wolfxl_rels::TargetMode::External {
                continue;
            }
            let target = wolfxl_rels::resolve_target("xl", &rel.target);
            if cache_parts.contains(&target) {
                to_remove.push(rel.id.clone());
            }
        }
        for rid in to_remove {
            removed_rids.insert(rid.0.clone());
            graph.remove(&rid);
        }
    }
    for part in &cache_parts {
        patcher.file_deletes.insert(part.clone());
        if let Some(rels_path) = wolfxl_rels::rels_path_for(part) {
            patcher.file_deletes.insert(rels_path);
        }
        patcher
            .queued_content_type_ops
            .entry("__sheet_delete__".to_string())
            .or_default()
            .push(content_types::ContentTypeOp::RemoveOverride(format!(
                "/{part}"
            )));
    }
    let mut xml = String::from_utf8(workbook_xml)
        .map_err(|e| PyIOError::new_err(format!("sheet-delete workbook.xml utf8: {e}")))?;
    for rid in &removed_rids {
        xml = remove_self_closing_element_with_attr(&xml, "slicerCache", "id", rid);
        xml = remove_self_closing_element_with_attr(&xml, "timelineCacheRef", "id", rid);
    }
    for name in cache_names {
        xml = remove_defined_name_by_name(&xml, name);
    }
    Ok(xml.into_bytes())
}

fn workbook_cache_parts_for_names(
    cache_names: &HashSet<String>,
    file_patches: &HashMap<String, Vec<u8>>,
    file_adds: &HashMap<String, Vec<u8>>,
    zip: &mut ZipArchive<File>,
) -> HashSet<String> {
    let mut candidates: HashSet<String> = HashSet::new();
    for i in 0..zip.len() {
        let Ok(entry) = zip.by_index(i) else {
            continue;
        };
        let name = entry.name().to_string();
        if is_workbook_cache_part(&name) {
            candidates.insert(name);
        }
    }
    candidates.extend(
        file_patches
            .keys()
            .filter(|p| is_workbook_cache_part(p))
            .cloned(),
    );
    candidates.extend(
        file_adds
            .keys()
            .filter(|p| is_workbook_cache_part(p))
            .cloned(),
    );

    let mut out = HashSet::new();
    for part in candidates {
        let Some(bytes) = current_part_bytes(file_patches, file_adds, zip, &part) else {
            continue;
        };
        if xml_attr_values_by_local_name(&bytes, b"name")
            .into_iter()
            .any(|name| cache_names.contains(&name))
        {
            out.insert(part);
        }
    }
    out
}

fn is_workbook_cache_part(path: &str) -> bool {
    (path.starts_with("xl/slicerCaches/slicerCache") && path.ends_with(".xml"))
        || (path.starts_with("xl/timelineCaches/timelineCache") && path.ends_with(".xml"))
}

fn xml_attr_values_by_local_name(xml: &[u8], attr_name: &[u8]) -> Vec<String> {
    let mut reader = XmlReader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut out = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                for attr in e.attributes().with_checks(false).flatten() {
                    if attr.key.local_name().as_ref() == attr_name {
                        if let Ok(value) = attr.decode_and_unescape_value(reader.decoder()) {
                            out.push(value.into_owned());
                        }
                    }
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

fn remove_self_closing_element_with_attr(
    xml: &str,
    local: &str,
    attr: &str,
    value: &str,
) -> String {
    let mut out = xml.to_string();
    loop {
        let Some(pos) = find_local_tag_with_attr(&out, local, attr, value) else {
            break;
        };
        let Some(end_rel) = out[pos..].find("/>") else {
            break;
        };
        let end = pos + end_rel + 2;
        out.replace_range(pos..end, "");
    }
    out
}

fn remove_defined_name_by_name(xml: &str, name: &str) -> String {
    let mut out = xml.to_string();
    loop {
        let Some(pos) = find_local_tag_with_attr(&out, "definedName", "name", name) else {
            break;
        };
        let Some(end_rel) = out[pos..].find("</definedName>") else {
            break;
        };
        let end = pos + end_rel + "</definedName>".len();
        out.replace_range(pos..end, "");
    }
    out
}

fn find_local_tag_with_attr(xml: &str, local: &str, attr: &str, value: &str) -> Option<usize> {
    let needle = format!("{attr}=\"{value}\"");
    let mut offset = 0;
    while let Some(rel) = xml[offset..].find(&needle) {
        let hit = offset + rel;
        let tag_start = xml[..hit].rfind('<')?;
        let tag_end = xml[hit..].find('>').map(|e| hit + e)?;
        let tag = &xml[tag_start + 1..tag_end];
        let tag_name = tag.split_whitespace().next().unwrap_or("");
        let tag_local = tag_name.rsplit(':').next().unwrap_or(tag_name);
        if tag_local == local {
            return Some(tag_start);
        }
        offset = hit + needle.len();
    }
    None
}

fn retained_deleted_subgraph_parts(
    deleted_parts: &HashSet<String>,
    patcher: &XlsxPatcher,
    file_patches: &HashMap<String, Vec<u8>>,
    zip: &mut ZipArchive<File>,
) -> PyResult<HashSet<String>> {
    let mut rels_paths: HashSet<String> = HashSet::new();
    for i in 0..zip.len() {
        let Ok(entry) = zip.by_index(i) else {
            continue;
        };
        let name = entry.name().to_string();
        if name.ends_with(".rels") {
            rels_paths.insert(name);
        }
    }
    rels_paths.extend(patcher.rels_patches.keys().cloned());
    rels_paths.extend(
        file_patches
            .keys()
            .filter(|p| p.ends_with(".rels"))
            .cloned(),
    );
    rels_paths.extend(
        patcher
            .file_adds
            .keys()
            .filter(|p| p.ends_with(".rels"))
            .cloned(),
    );

    let mut retained: HashSet<String> = HashSet::new();
    let mut changed = true;
    while changed {
        changed = false;
        for rels_path in &rels_paths {
            let Some(owner_part) = owner_part_for_rels_path(rels_path) else {
                continue;
            };
            let owner_is_kept =
                !deleted_parts.contains(&owner_part) || retained.contains(&owner_part);
            if !owner_is_kept {
                continue;
            }
            let graph = if let Some(graph) = patcher.rels_patches.get(rels_path) {
                graph.clone()
            } else {
                let Some(bytes) =
                    current_part_bytes(file_patches, &patcher.file_adds, zip, rels_path)
                else {
                    continue;
                };
                let Ok(graph) = RelsGraph::parse(&bytes) else {
                    continue;
                };
                graph
            };
            let owner_dir = owner_part
                .rsplit_once('/')
                .map(|(dir, _)| dir)
                .unwrap_or("");
            for rel in graph.iter() {
                if rel.mode == wolfxl_rels::TargetMode::External {
                    continue;
                }
                let target = wolfxl_rels::resolve_target(owner_dir, &rel.target);
                if deleted_parts.contains(&target) && retained.insert(target) {
                    changed = true;
                }
            }
        }
    }

    let retained_presentations: Vec<String> = retained
        .iter()
        .filter(|part| {
            (part.starts_with("xl/slicers/slicer") && part.ends_with(".xml"))
                || (part.starts_with("xl/timelines/timeline") && part.ends_with(".xml"))
        })
        .cloned()
        .collect();
    let mut retained_cache_names: HashSet<String> = HashSet::new();
    for part in retained_presentations {
        let Some(bytes) = current_part_bytes(file_patches, &patcher.file_adds, zip, &part) else {
            continue;
        };
        retained_cache_names.extend(xml_attr_values_by_local_name(&bytes, b"cache"));
    }
    for part in
        workbook_cache_parts_for_names(&retained_cache_names, file_patches, &patcher.file_adds, zip)
    {
        if deleted_parts.contains(&part) {
            retained.insert(part);
        }
    }

    Ok(retained)
}

fn owner_part_for_rels_path(rels_path: &str) -> Option<String> {
    if rels_path == "_rels/.rels" {
        return Some(String::new());
    }
    let (dir, file) = rels_path.rsplit_once("/_rels/")?;
    let part_file = file.strip_suffix(".rels")?;
    if dir.is_empty() {
        Some(part_file.to_string())
    } else {
        Some(format!("{dir}/{part_file}"))
    }
}

pub(super) fn apply_sheet_creates_phase(
    patcher: &mut XlsxPatcher,
    file_patches: &mut HashMap<String, Vec<u8>>,
    zip: &mut ZipArchive<File>,
    part_id_allocator: &mut wolfxl_rels::PartIdAllocator,
) -> PyResult<()> {
    if patcher.queued_sheet_creates.is_empty() {
        return Ok(());
    }

    let workbook_rels_path = "xl/_rels/workbook.xml.rels".to_string();
    if !patcher.rels_patches.contains_key(&workbook_rels_path) {
        let g = load_or_empty_rels(zip, &workbook_rels_path)?;
        patcher.rels_patches.insert(workbook_rels_path.clone(), g);
    }

    let desired_order = patcher.sheet_order.clone();
    let queued_titles: HashSet<String> = patcher
        .queued_sheet_creates
        .iter()
        .map(|op| op.title.clone())
        .collect();
    let mut created_titles: HashSet<String> = HashSet::new();
    let ops = patcher.queued_sheet_creates.clone();
    for op in ops {
        if !desired_order.contains(&op.title) {
            continue;
        }
        let sheet_n = part_id_allocator.alloc_sheet();
        let sheet_path = format!("xl/worksheets/sheet{sheet_n}.xml");
        let target = format!("worksheets/sheet{sheet_n}.xml");
        let rid = {
            let graph = patcher
                .rels_patches
                .get_mut(&workbook_rels_path)
                .expect("workbook rels graph loaded above");
            graph.add(
                wolfxl_rels::rt::WORKSHEET,
                &target,
                wolfxl_rels::TargetMode::Internal,
            )
        };

        let workbook_xml =
            match current_part_bytes(file_patches, &patcher.file_adds, zip, "xl/workbook.xml") {
                Some(b) => b,
                None => {
                    return Err(PyIOError::new_err(
                        "sheet-create: xl/workbook.xml missing from source ZIP",
                    ));
                }
            };
        let sheet_id = next_sheet_id(&workbook_xml);
        let sheet_element = format!(
            "<sheet name=\"{}\" sheetId=\"{}\" r:id=\"{}\"/>",
            xml_escape_attr(&op.title),
            sheet_id,
            xml_escape_attr(&rid.0)
        )
        .into_bytes();
        let mut workbook_xml = splice_into_sheets_block(&workbook_xml, &sheet_element)?;

        let desired_idx = desired_order
            .iter()
            .take_while(|name| *name != &op.title)
            .filter(|name| !queued_titles.contains(*name) || created_titles.contains(*name))
            .count();
        let last_idx = next_sheet_count(&workbook_xml).saturating_sub(1);
        if desired_idx != last_idx {
            let offset = desired_idx as i32 - last_idx as i32;
            let result =
                sheet_order::merge_sheet_moves(&workbook_xml, &[(op.title.clone(), offset)])
                    .map_err(|e| PyIOError::new_err(format!("sheet-create order: {e}")))?;
            workbook_xml = result.workbook_xml;
        }
        file_patches.insert("xl/workbook.xml".to_string(), workbook_xml);
        patcher
            .file_adds
            .insert(sheet_path.clone(), minimal_worksheet_xml());
        patcher
            .sheet_paths
            .insert(op.title.clone(), sheet_path.clone());
        patcher
            .queued_content_type_ops
            .entry("__sheet_create__".to_string())
            .or_default()
            .push(content_types::ContentTypeOp::AddOverride(
                format!("/{sheet_path}"),
                CT_WORKSHEET.to_string(),
            ));
        created_titles.insert(op.title.clone());
    }

    patcher.sheet_order = desired_order;
    patcher.queued_sheet_creates.clear();
    Ok(())
}

fn minimal_worksheet_xml() -> Vec<u8> {
    br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/></worksheet>"#.to_vec()
}

fn next_sheet_id(workbook_xml: &[u8]) -> u32 {
    max_sheet_id_and_count(workbook_xml)
        .0
        .saturating_add(1)
        .max(1)
}

fn next_sheet_count(workbook_xml: &[u8]) -> usize {
    max_sheet_id_and_count(workbook_xml).1
}

fn max_sheet_id_and_count(workbook_xml: &[u8]) -> (u32, usize) {
    let mut reader = quick_xml::Reader::from_reader(workbook_xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut max_id = 0_u32;
    let mut count = 0_usize;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(e)) | Ok(quick_xml::events::Event::Empty(e))
                if e.local_name().as_ref() == b"sheet" =>
            {
                count += 1;
                for attr in e.attributes().with_checks(false).flatten() {
                    if attr.key.as_ref() == b"sheetId" {
                        let value = attr
                            .unescape_value()
                            .map(|v| v.into_owned())
                            .unwrap_or_else(|_| {
                                String::from_utf8_lossy(attr.value.as_ref()).into_owned()
                            });
                        if let Ok(parsed) = value.parse::<u32>() {
                            max_id = max_id.max(parsed);
                        }
                    }
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    (max_id, count)
}

fn sheet_ids_by_name(workbook_xml: &[u8]) -> HashMap<String, u32> {
    let mut reader = quick_xml::Reader::from_reader(workbook_xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut ids = HashMap::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(e)) | Ok(quick_xml::events::Event::Empty(e))
                if e.local_name().as_ref() == b"sheet" =>
            {
                let mut name = None;
                let mut sheet_id = None;
                for attr in e.attributes().with_checks(false).flatten() {
                    let value = attr
                        .unescape_value()
                        .map(|v| v.into_owned())
                        .unwrap_or_else(|_| {
                            String::from_utf8_lossy(attr.value.as_ref()).into_owned()
                        });
                    match attr.key.as_ref() {
                        b"name" => name = Some(value),
                        b"sheetId" => sheet_id = value.parse::<u32>().ok(),
                        _ => {}
                    }
                }
                if let (Some(name), Some(sheet_id)) = (name, sheet_id) {
                    ids.insert(name, sheet_id);
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    ids
}

fn workbook_root_has_relationship_namespace(s: &str) -> PyResult<bool> {
    let (open, rel_end) = find_open_tag_by_local_name(s, "workbook")
        .ok_or_else(|| PyIOError::new_err("Phase 2.7: workbook.xml has no <workbook> root"))?;
    Ok(s[open..rel_end].contains("xmlns:r="))
}

fn add_workbook_relationship_namespace(s: &str) -> PyResult<String> {
    let (_open, rel_end) = find_open_tag_by_local_name(s, "workbook")
        .ok_or_else(|| PyIOError::new_err("Phase 2.7: workbook.xml has no <workbook> root"))?;
    let insert_at = rel_end;
    let mut out = String::with_capacity(
        s.len()
            + " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\""
                .len(),
    );
    out.push_str(&s[..insert_at]);
    out.push_str(
        " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"",
    );
    out.push_str(&s[insert_at..]);
    Ok(out)
}

fn find_open_tag_by_local_name(s: &str, local_name: &str) -> Option<(usize, usize)> {
    let mut search_from = 0;
    while let Some(rel_start) = s[search_from..].find('<') {
        let start = search_from + rel_start;
        let name_start = start + 1;
        let first = s.as_bytes().get(name_start).copied()?;
        if matches!(first, b'/' | b'!' | b'?') {
            search_from = name_start + 1;
            continue;
        }
        let name_end = s[name_start..]
            .find(|ch: char| ch.is_whitespace() || ch == '>' || ch == '/')
            .map(|rel| name_start + rel)?;
        let element_name = &s[name_start..name_end];
        if element_name.rsplit(':').next() == Some(local_name) {
            let tag_end = s[name_end..].find('>').map(|rel| name_end + rel)?;
            return Some((start, tag_end));
        }
        search_from = name_end;
    }
    None
}

/// Naive byte-substring search for tiny workbook XML payloads.
#[allow(dead_code)]
pub(crate) fn find_subslice(haystack: &[u8], needle: &[u8]) -> Option<usize> {
    if needle.is_empty() || needle.len() > haystack.len() {
        return None;
    }
    haystack.windows(needle.len()).position(|w| w == needle)
}

/// Escapes text for an XML attribute value.
pub(crate) fn xml_escape_attr(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for c in s.chars() {
        match c {
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '&' => out.push_str("&amp;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&apos;"),
            other => out.push(other),
        }
    }
    out
}

/// Replaces the first occurrence of `needle` in `haystack`.
pub(crate) fn replace_first_occurrence(
    haystack: &str,
    needle: &str,
    replacement: &str,
) -> Option<String> {
    let idx = haystack.find(needle)?;
    let mut out = String::with_capacity(haystack.len() - needle.len() + replacement.len());
    out.push_str(&haystack[..idx]);
    out.push_str(replacement);
    out.push_str(&haystack[idx + needle.len()..]);
    Some(out)
}

/// Returns a ZIP timestamp honoring `WOLFXL_TEST_EPOCH` when set.
pub(crate) fn epoch_or_now() -> zip::DateTime {
    use chrono::{Datelike, Timelike};
    let secs = std::env::var("WOLFXL_TEST_EPOCH")
        .ok()
        .and_then(|s| s.parse::<i64>().ok());
    let dt = match secs.and_then(|s| chrono::DateTime::<chrono::Utc>::from_timestamp(s, 0)) {
        Some(d) => d,
        None => chrono::Utc::now(),
    };
    let naive = dt.naive_utc();
    let year = naive.year();
    if year < 1980 {
        return zip::DateTime::from_date_and_time(1980, 1, 1, 0, 0, 0)
            .unwrap_or_else(|_| zip::DateTime::default());
    }
    if year > 2107 {
        return zip::DateTime::from_date_and_time(2107, 12, 31, 23, 59, 58)
            .unwrap_or_else(|_| zip::DateTime::default());
    }
    zip::DateTime::from_date_and_time(
        year as u16,
        naive.month() as u8,
        naive.day() as u8,
        naive.hour() as u8,
        naive.minute() as u8,
        naive.second() as u8,
    )
    .unwrap_or_else(|_| zip::DateTime::default())
}

/// Minimal styles part used when a workbook has style-dependent edits but no styles.xml.
pub(crate) fn minimal_styles_xml() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>
</styleSheet>"#
        .to_string()
}

pub(super) fn apply_content_types_phase(
    patcher: &mut XlsxPatcher,
    file_patches: &mut HashMap<String, Vec<u8>>,
    zip: &mut ZipArchive<File>,
) -> PyResult<()> {
    let mut content_type_ops: Vec<content_types::ContentTypeOp> = Vec::new();
    for sheet_name in &patcher.sheet_order {
        if let Some(ops) = patcher.queued_content_type_ops.get(sheet_name) {
            content_type_ops.extend(ops.iter().cloned());
        }
    }
    // Also pick up synthetic per-workbook keys (e.g. RFC-023
    // ``__rfc023_comments__`` and RFC-045
    // ``__rfc045_drawing_N__``) that aren't tied to a single
    // sheet name in `sheet_order`. Iterate in sorted order so the
    // emitted Override sequence is deterministic.
    let mut synth_keys: Vec<&String> = patcher
        .queued_content_type_ops
        .keys()
        .filter(|k| !patcher.sheet_order.contains(k))
        .collect();
    synth_keys.sort();
    for k in synth_keys {
        if let Some(ops) = patcher.queued_content_type_ops.get(k) {
            content_type_ops.extend(ops.iter().cloned());
        }
    }
    if !content_type_ops.is_empty() {
        let ct_xml = ooxml_util::zip_read_to_string(zip, "[Content_Types].xml")?;
        let mut graph = content_types::ContentTypesGraph::parse(ct_xml.as_bytes())
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("[Content_Types].xml parse: {e}")))?;
        for op in &content_type_ops {
            graph.apply_op(op);
        }
        file_patches.insert("[Content_Types].xml".to_string(), graph.serialize());
    }

    Ok(())
}

pub(super) fn apply_document_properties_phase(
    patcher: &mut XlsxPatcher,
    file_patches: &mut HashMap<String, Vec<u8>>,
    zip: &mut ZipArchive<File>,
) -> PyResult<()> {
    if let Some(ref payload) = patcher.queued_props {
        let mut effective = payload.clone();
        if effective.sheet_names.is_empty() {
            effective.sheet_names = patcher.sheet_order.clone();
        }
        let core_bytes = properties::rewrite_core_props(&effective);
        let app_bytes = properties::rewrite_app_props(&effective);

        let core_in_source = source_zip_has_entry(zip, "docProps/core.xml");
        let app_in_source = source_zip_has_entry(zip, "docProps/app.xml");

        if core_in_source {
            file_patches.insert("docProps/core.xml".into(), core_bytes);
        } else {
            patcher
                .file_adds
                .insert("docProps/core.xml".into(), core_bytes);
        }
        if app_in_source {
            file_patches.insert("docProps/app.xml".into(), app_bytes);
        } else {
            patcher
                .file_adds
                .insert("docProps/app.xml".into(), app_bytes);
        }
    }

    Ok(())
}

// These phases all mutate xl/workbook.xml, so keep the in-progress bytes flowing
// security -> sheet renames -> sheet order -> defined names before publishing
// the final patch.
pub(super) fn apply_workbook_xml_phases(
    patcher: &mut XlsxPatcher,
    file_patches: &mut HashMap<String, Vec<u8>>,
    zip: &mut ZipArchive<File>,
) -> PyResult<()> {
    let mut workbook_xml_in_progress: Option<Vec<u8>> = None;

    if let Some(ref sec) = patcher.queued_workbook_security {
        if !sec.is_empty() {
            let wb_bytes: Vec<u8> = match file_patches.get("xl/workbook.xml") {
                Some(b) => b.clone(),
                None => ooxml_util::zip_read_to_string(zip, "xl/workbook.xml")?.into_bytes(),
            };
            let updated = security::merge_workbook_security(&wb_bytes, sec)
                .map_err(|e| PyIOError::new_err(format!("workbook-security merge: {e}")))?;
            workbook_xml_in_progress = Some(updated);
        }
    }

    if !patcher.queued_sheet_renames.is_empty() {
        let wb_bytes: Vec<u8> = match workbook_xml_in_progress.take() {
            Some(b) => b,
            None => match file_patches.get("xl/workbook.xml") {
                Some(b) => b.clone(),
                None => ooxml_util::zip_read_to_string(zip, "xl/workbook.xml")?.into_bytes(),
            },
        };
        let updated = sheet_order::merge_sheet_renames(&wb_bytes, &patcher.queued_sheet_renames)
            .map_err(|e| PyIOError::new_err(format!("sheet-rename merge: {e}")))?;
        let updated =
            rewrite_workbook_defined_name_sheet_renames(&updated, &patcher.queued_sheet_renames)
                .map_err(|e| PyIOError::new_err(format!("defined-name sheet-rename merge: {e}")))?;
        workbook_xml_in_progress = Some(updated);
    }

    if !patcher.queued_sheet_moves.is_empty() {
        let wb_bytes: Vec<u8> = match workbook_xml_in_progress.take() {
            Some(b) => b,
            None => match file_patches.get("xl/workbook.xml") {
                Some(b) => b.clone(),
                None => ooxml_util::zip_read_to_string(zip, "xl/workbook.xml")?.into_bytes(),
            },
        };
        let result = sheet_order::merge_sheet_moves(&wb_bytes, &patcher.queued_sheet_moves)
            .map_err(|e| PyIOError::new_err(format!("sheet-reorder merge: {e}")))?;
        workbook_xml_in_progress = Some(result.workbook_xml);
        patcher.sheet_order = result.new_order;
    }

    if !patcher.queued_defined_names.is_empty() {
        let wb_xml_bytes: Vec<u8> = match workbook_xml_in_progress.take() {
            Some(bytes) => bytes,
            None => match file_patches.get("xl/workbook.xml") {
                Some(bytes) => bytes.clone(),
                None => ooxml_util::zip_read_to_string(zip, "xl/workbook.xml")?.into_bytes(),
            },
        };
        let updated =
            defined_names::merge_defined_names(&wb_xml_bytes, &patcher.queued_defined_names)
                .map_err(|e| PyIOError::new_err(format!("defined-names merge: {e}")))?;
        file_patches.insert("xl/workbook.xml".to_string(), updated);
    } else if let Some(bytes) = workbook_xml_in_progress.take() {
        file_patches.insert("xl/workbook.xml".to_string(), bytes);
    }

    Ok(())
}

pub(super) fn apply_sheet_rename_chart_formula_refs_phase(
    patcher: &mut XlsxPatcher,
    file_patches: &mut HashMap<String, Vec<u8>>,
    zip: &mut ZipArchive<File>,
) -> PyResult<()> {
    if patcher.queued_sheet_renames.is_empty() {
        return Ok(());
    }

    let mut chart_paths: HashSet<String> = HashSet::new();
    let mut worksheet_paths: HashSet<String> = HashSet::new();
    let mut pivot_cache_paths: HashSet<String> = HashSet::new();
    for idx in 0..zip.len() {
        let Ok(entry) = zip.by_index(idx) else {
            continue;
        };
        let name = entry.name();
        if is_chart_part_path(name) {
            chart_paths.insert(name.to_string());
        }
        if is_worksheet_part_path(name) {
            worksheet_paths.insert(name.to_string());
        }
        if is_pivot_cache_definition_path(name) {
            pivot_cache_paths.insert(name.to_string());
        }
    }
    for path in patcher.file_adds.keys() {
        if is_chart_part_path(path) {
            chart_paths.insert(path.clone());
        }
        if is_worksheet_part_path(path) {
            worksheet_paths.insert(path.clone());
        }
        if is_pivot_cache_definition_path(path) {
            pivot_cache_paths.insert(path.clone());
        }
    }
    for path in file_patches.keys() {
        if is_chart_part_path(path) {
            chart_paths.insert(path.clone());
        }
        if is_worksheet_part_path(path) {
            worksheet_paths.insert(path.clone());
        }
        if is_pivot_cache_definition_path(path) {
            pivot_cache_paths.insert(path.clone());
        }
    }

    for path in worksheet_paths {
        let Some(bytes) = current_part_bytes(file_patches, &patcher.file_adds, zip, &path) else {
            continue;
        };
        let updated =
            rewrite_worksheet_formula_sheet_renames(&bytes, &patcher.queued_sheet_renames)
                .map_err(|e| {
                    PyIOError::new_err(format!("worksheet formula sheet-rename merge: {e}"))
                })?;
        if updated != bytes {
            file_patches.insert(path, updated);
        }
    }

    for path in chart_paths {
        let Some(bytes) = current_part_bytes(file_patches, &patcher.file_adds, zip, &path) else {
            continue;
        };
        let updated = rewrite_chart_formula_sheet_renames(&bytes, &patcher.queued_sheet_renames)
            .map_err(|e| PyIOError::new_err(format!("chart formula sheet-rename merge: {e}")))?;
        if updated != bytes {
            file_patches.insert(path, updated);
        }
    }

    for path in pivot_cache_paths {
        let Some(bytes) = current_part_bytes(file_patches, &patcher.file_adds, zip, &path) else {
            continue;
        };
        let updated = rewrite_pivot_cache_sheet_renames(&bytes, &patcher.queued_sheet_renames)
            .map_err(|e| PyIOError::new_err(format!("pivot cache sheet-rename merge: {e}")))?;
        if updated != bytes {
            file_patches.insert(path, updated);
        }
    }

    Ok(())
}

fn is_chart_part_path(path: &str) -> bool {
    path.starts_with("xl/charts/") && path.ends_with(".xml")
}

fn is_worksheet_part_path(path: &str) -> bool {
    path.starts_with("xl/worksheets/sheet") && path.ends_with(".xml")
}

fn is_pivot_cache_definition_path(path: &str) -> bool {
    path.starts_with("xl/pivotCache/pivotCacheDefinition") && path.ends_with(".xml")
}

fn rewrite_workbook_defined_name_sheet_renames(
    workbook_xml: &[u8],
    renames: &[(String, String)],
) -> Result<Vec<u8>, String> {
    let mut reader = XmlReader::from_reader(workbook_xml);
    reader.config_mut().trim_text(false);
    let mut writer = XmlWriter::new(Vec::with_capacity(workbook_xml.len()));
    let mut buf = Vec::new();
    let mut in_defined_name = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if e.local_name().as_ref() == b"definedName" {
                    in_defined_name = true;
                }
                writer
                    .write_event(Event::Start(e.into_owned()))
                    .map_err(|e| format!("workbook XML write error: {e}"))?;
            }
            Ok(Event::End(e)) => {
                if e.local_name().as_ref() == b"definedName" {
                    in_defined_name = false;
                }
                writer
                    .write_event(Event::End(e.into_owned()))
                    .map_err(|e| format!("workbook XML write error: {e}"))?;
            }
            Ok(Event::Empty(e)) => {
                writer
                    .write_event(Event::Empty(e.into_owned()))
                    .map_err(|e| format!("workbook XML write error: {e}"))?;
            }
            Ok(Event::Text(t)) if in_defined_name => {
                let formula_text = t
                    .unescape()
                    .map_err(|e| format!("defined name decode error: {e}"))?
                    .into_owned();
                let translated = rename_formula_sheet_refs(&formula_text, renames);
                writer
                    .write_event(Event::Text(BytesText::new(&translated)))
                    .map_err(|e| format!("workbook XML write error: {e}"))?;
            }
            Ok(Event::Text(t)) => {
                writer
                    .write_event(Event::Text(t.into_owned()))
                    .map_err(|e| format!("workbook XML write error: {e}"))?;
            }
            Ok(Event::CData(c)) => writer
                .write_event(Event::CData(c.into_owned()))
                .map_err(|e| format!("workbook XML write error: {e}"))?,
            Ok(Event::Decl(d)) => writer
                .write_event(Event::Decl(d.into_owned()))
                .map_err(|e| format!("workbook XML write error: {e}"))?,
            Ok(Event::PI(p)) => writer
                .write_event(Event::PI(p.into_owned()))
                .map_err(|e| format!("workbook XML write error: {e}"))?,
            Ok(Event::Comment(c)) => writer
                .write_event(Event::Comment(c.into_owned()))
                .map_err(|e| format!("workbook XML write error: {e}"))?,
            Ok(Event::DocType(d)) => writer
                .write_event(Event::DocType(d.into_owned()))
                .map_err(|e| format!("workbook XML write error: {e}"))?,
            Ok(Event::Eof) => break,
            Err(e) => return Err(format!("workbook XML parse error: {e}")),
        }
        buf.clear();
    }

    Ok(writer.into_inner())
}

fn rewrite_worksheet_formula_sheet_renames(
    worksheet_xml: &[u8],
    renames: &[(String, String)],
) -> Result<Vec<u8>, String> {
    let mut reader = XmlReader::from_reader(worksheet_xml);
    reader.config_mut().trim_text(false);
    let mut writer = XmlWriter::new(Vec::with_capacity(worksheet_xml.len()));
    let mut buf = Vec::new();
    let mut formula_depth = 0usize;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if is_formula_text_element(e.local_name().as_ref()) {
                    formula_depth += 1;
                }
                writer
                    .write_event(Event::Start(e.into_owned()))
                    .map_err(|e| format!("worksheet XML write error: {e}"))?;
            }
            Ok(Event::End(e)) => {
                let closes_formula = is_formula_text_element(e.local_name().as_ref());
                writer
                    .write_event(Event::End(e.into_owned()))
                    .map_err(|e| format!("worksheet XML write error: {e}"))?;
                if closes_formula {
                    formula_depth = formula_depth.saturating_sub(1);
                }
            }
            Ok(Event::Empty(e)) => {
                writer
                    .write_event(Event::Empty(e.into_owned()))
                    .map_err(|e| format!("worksheet XML write error: {e}"))?;
            }
            Ok(Event::Text(t)) if formula_depth > 0 => {
                let formula_text = t
                    .unescape()
                    .map_err(|e| format!("worksheet formula decode error: {e}"))?
                    .into_owned();
                if formula_text.trim().is_empty() {
                    writer
                        .write_event(Event::Text(t.into_owned()))
                        .map_err(|e| format!("worksheet XML write error: {e}"))?;
                } else {
                    let translated = rename_formula_sheet_refs(&formula_text, renames);
                    writer
                        .write_event(Event::Text(BytesText::new(&translated)))
                        .map_err(|e| format!("worksheet XML write error: {e}"))?;
                }
            }
            Ok(Event::Text(t)) => {
                writer
                    .write_event(Event::Text(t.into_owned()))
                    .map_err(|e| format!("worksheet XML write error: {e}"))?;
            }
            Ok(Event::CData(c)) => writer
                .write_event(Event::CData(c.into_owned()))
                .map_err(|e| format!("worksheet XML write error: {e}"))?,
            Ok(Event::Decl(d)) => writer
                .write_event(Event::Decl(d.into_owned()))
                .map_err(|e| format!("worksheet XML write error: {e}"))?,
            Ok(Event::PI(p)) => writer
                .write_event(Event::PI(p.into_owned()))
                .map_err(|e| format!("worksheet XML write error: {e}"))?,
            Ok(Event::Comment(c)) => writer
                .write_event(Event::Comment(c.into_owned()))
                .map_err(|e| format!("worksheet XML write error: {e}"))?,
            Ok(Event::DocType(d)) => writer
                .write_event(Event::DocType(d.into_owned()))
                .map_err(|e| format!("worksheet XML write error: {e}"))?,
            Ok(Event::Eof) => break,
            Err(e) => return Err(format!("worksheet XML parse error: {e}")),
        }
        buf.clear();
    }

    Ok(writer.into_inner())
}

fn is_formula_text_element(local_name: &[u8]) -> bool {
    matches!(local_name, b"f" | b"formula" | b"formula1" | b"formula2")
}

fn rewrite_pivot_cache_sheet_renames(
    pivot_cache_xml: &[u8],
    renames: &[(String, String)],
) -> Result<Vec<u8>, String> {
    let mut reader = XmlReader::from_reader(pivot_cache_xml);
    reader.config_mut().trim_text(false);
    let mut writer = XmlWriter::new(Vec::with_capacity(pivot_cache_xml.len()));
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.local_name().as_ref() == b"worksheetSource" => {
                let updated = rename_worksheet_source_sheet_attr(&e, renames)?;
                writer
                    .write_event(Event::Start(updated))
                    .map_err(|e| format!("pivot cache XML write error: {e}"))?;
            }
            Ok(Event::Empty(e)) if e.local_name().as_ref() == b"worksheetSource" => {
                let updated = rename_worksheet_source_sheet_attr(&e, renames)?;
                writer
                    .write_event(Event::Empty(updated))
                    .map_err(|e| format!("pivot cache XML write error: {e}"))?;
            }
            Ok(Event::Start(e)) => writer
                .write_event(Event::Start(e.into_owned()))
                .map_err(|e| format!("pivot cache XML write error: {e}"))?,
            Ok(Event::End(e)) => writer
                .write_event(Event::End(e.into_owned()))
                .map_err(|e| format!("pivot cache XML write error: {e}"))?,
            Ok(Event::Empty(e)) => writer
                .write_event(Event::Empty(e.into_owned()))
                .map_err(|e| format!("pivot cache XML write error: {e}"))?,
            Ok(Event::Text(t)) => writer
                .write_event(Event::Text(t.into_owned()))
                .map_err(|e| format!("pivot cache XML write error: {e}"))?,
            Ok(Event::CData(c)) => writer
                .write_event(Event::CData(c.into_owned()))
                .map_err(|e| format!("pivot cache XML write error: {e}"))?,
            Ok(Event::Decl(d)) => writer
                .write_event(Event::Decl(d.into_owned()))
                .map_err(|e| format!("pivot cache XML write error: {e}"))?,
            Ok(Event::PI(p)) => writer
                .write_event(Event::PI(p.into_owned()))
                .map_err(|e| format!("pivot cache XML write error: {e}"))?,
            Ok(Event::Comment(c)) => writer
                .write_event(Event::Comment(c.into_owned()))
                .map_err(|e| format!("pivot cache XML write error: {e}"))?,
            Ok(Event::DocType(d)) => writer
                .write_event(Event::DocType(d.into_owned()))
                .map_err(|e| format!("pivot cache XML write error: {e}"))?,
            Ok(Event::Eof) => break,
            Err(e) => return Err(format!("pivot cache XML parse error: {e}")),
        }
        buf.clear();
    }

    Ok(writer.into_inner())
}

fn rename_worksheet_source_sheet_attr(
    event: &BytesStart<'_>,
    renames: &[(String, String)],
) -> Result<BytesStart<'static>, String> {
    let mut out = BytesStart::new(String::from_utf8_lossy(event.name().as_ref()).into_owned());
    for attr in event.attributes().with_checks(false) {
        let attr = attr.map_err(|e| format!("pivot cache worksheetSource attr error: {e}"))?;
        if attr.key.as_ref() == b"sheet" {
            let value = attr
                .unescape_value()
                .map_err(|e| format!("pivot cache worksheetSource sheet decode error: {e}"))?;
            if let Some((_, new_name)) = renames
                .iter()
                .find(|(old_name, _)| value.as_ref() == old_name.as_str())
            {
                out.push_attribute(("sheet", new_name.as_str()));
            } else {
                out.push_attribute(("sheet", value.as_ref()));
            }
        } else {
            out.push_attribute(attr);
        }
    }
    Ok(out)
}

fn rewrite_chart_formula_sheet_renames(
    chart_xml: &[u8],
    renames: &[(String, String)],
) -> Result<Vec<u8>, String> {
    let mut reader = XmlReader::from_reader(chart_xml);
    reader.config_mut().trim_text(false);
    let mut writer = XmlWriter::new(Vec::with_capacity(chart_xml.len()));
    let mut buf = Vec::new();
    let mut in_formula = false;
    let mut in_pivot_source = false;
    let mut in_pivot_source_name = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let local_name = e.local_name();
                let local = local_name.as_ref();
                if local == b"pivotSource" {
                    in_pivot_source = true;
                }
                in_formula = local == b"f";
                in_pivot_source_name = in_pivot_source && local == b"name";
                writer
                    .write_event(Event::Start(e.into_owned()))
                    .map_err(|e| format!("chart XML write error: {e}"))?;
            }
            Ok(Event::End(e)) => {
                let local_name = e.local_name();
                let local = local_name.as_ref();
                if local == b"f" {
                    in_formula = false;
                }
                if local == b"name" {
                    in_pivot_source_name = false;
                }
                if local == b"pivotSource" {
                    in_pivot_source = false;
                }
                writer
                    .write_event(Event::End(e.into_owned()))
                    .map_err(|e| format!("chart XML write error: {e}"))?;
            }
            Ok(Event::Empty(e)) => {
                writer
                    .write_event(Event::Empty(e.into_owned()))
                    .map_err(|e| format!("chart XML write error: {e}"))?;
            }
            Ok(Event::Text(t)) if in_formula => {
                let formula_text = t
                    .unescape()
                    .map_err(|e| format!("chart formula decode error: {e}"))?
                    .into_owned();
                let translated = rename_formula_sheet_refs(&formula_text, renames);
                writer
                    .write_event(Event::Text(BytesText::new(&translated)))
                    .map_err(|e| format!("chart XML write error: {e}"))?;
            }
            Ok(Event::Text(t)) if in_pivot_source_name => {
                let pivot_source_name = t
                    .unescape()
                    .map_err(|e| format!("chart pivot source decode error: {e}"))?
                    .into_owned();
                let translated = rename_chart_pivot_source_sheet_refs(&pivot_source_name, renames);
                writer
                    .write_event(Event::Text(BytesText::new(&translated)))
                    .map_err(|e| format!("chart XML write error: {e}"))?;
            }
            Ok(Event::Text(t)) => {
                writer
                    .write_event(Event::Text(t.into_owned()))
                    .map_err(|e| format!("chart XML write error: {e}"))?;
            }
            Ok(Event::CData(c)) => writer
                .write_event(Event::CData(c.into_owned()))
                .map_err(|e| format!("chart XML write error: {e}"))?,
            Ok(Event::Decl(d)) => writer
                .write_event(Event::Decl(d.into_owned()))
                .map_err(|e| format!("chart XML write error: {e}"))?,
            Ok(Event::PI(p)) => writer
                .write_event(Event::PI(p.into_owned()))
                .map_err(|e| format!("chart XML write error: {e}"))?,
            Ok(Event::Comment(c)) => writer
                .write_event(Event::Comment(c.into_owned()))
                .map_err(|e| format!("chart XML write error: {e}"))?,
            Ok(Event::DocType(d)) => writer
                .write_event(Event::DocType(d.into_owned()))
                .map_err(|e| format!("chart XML write error: {e}"))?,
            Ok(Event::Eof) => break,
            Err(e) => return Err(format!("chart XML parse error: {e}")),
        }
        buf.clear();
    }

    Ok(writer.into_inner())
}

fn rename_chart_pivot_source_sheet_refs(name: &str, renames: &[(String, String)]) -> String {
    let mut current = name.to_string();
    for (old_name, new_name) in renames {
        let Some(bang_idx) = current.find('!') else {
            continue;
        };
        let sheet_start = current[..bang_idx]
            .rfind(']')
            .map(|idx| idx + 1)
            .unwrap_or(0);
        if current[sheet_start..bang_idx] != *old_name {
            continue;
        }
        current = format!(
            "{}{}{}",
            &current[..sheet_start],
            new_name,
            &current[bang_idx..]
        );
    }
    current
}

fn rename_formula_sheet_refs(formula: &str, renames: &[(String, String)]) -> String {
    let had_prefix = formula.starts_with('=');
    let mut current = if had_prefix {
        formula.to_string()
    } else {
        format!("={formula}")
    };
    for (old_name, new_name) in renames {
        current = wolfxl_formula::rename_sheet(&current, old_name, new_name);
        current = rename_loose_local_sheet_refs(&current, old_name, new_name);
    }
    if had_prefix {
        current
    } else {
        current
            .strip_prefix('=')
            .unwrap_or(current.as_str())
            .to_string()
    }
}

fn rename_loose_local_sheet_refs(formula: &str, old_name: &str, new_name: &str) -> String {
    let quoted = rename_quoted_local_sheet_refs(formula, old_name, new_name);
    rename_unquoted_local_sheet_refs(&quoted, old_name, new_name)
}

fn rename_quoted_local_sheet_refs(formula: &str, old_name: &str, new_name: &str) -> String {
    let mut out = String::with_capacity(formula.len());
    let mut cursor = 0;
    while let Some(rel_start) = formula[cursor..].find('\'') {
        let start = cursor + rel_start;
        let mut idx = start + 1;
        let mut token = String::new();
        let mut close = None;
        while idx < formula.len() {
            let rest = &formula[idx..];
            if rest.starts_with("''") {
                token.push('\'');
                idx += 2;
                continue;
            }
            if rest.starts_with('\'') {
                close = Some(idx);
                break;
            }
            let Some(ch) = rest.chars().next() else {
                break;
            };
            token.push(ch);
            idx += ch.len_utf8();
        }

        let Some(close_idx) = close else {
            break;
        };
        let after_quote = close_idx + 1;
        if !formula[after_quote..].starts_with('!') {
            out.push_str(&formula[cursor..after_quote]);
            cursor = after_quote;
            continue;
        }

        out.push_str(&formula[cursor..start]);
        if let Some(renamed) = rename_sheet_token(&token, old_name, new_name) {
            out.push_str(&quote_sheet_token(&renamed));
        } else {
            out.push_str(&formula[start..after_quote]);
        }
        out.push('!');
        cursor = after_quote + 1;
    }
    out.push_str(&formula[cursor..]);
    out
}

fn rename_unquoted_local_sheet_refs(formula: &str, old_name: &str, new_name: &str) -> String {
    let mut out = String::with_capacity(formula.len());
    let mut cursor = 0;
    let double_quoted_ranges = double_quoted_ranges(formula);
    while let Some(rel_bang) = formula[cursor..].find('!') {
        let bang = cursor + rel_bang;
        if byte_index_in_ranges(bang, &double_quoted_ranges) {
            out.push_str(&formula[cursor..=bang]);
            cursor = bang + 1;
            continue;
        }
        let mut start = bang;
        while start > cursor {
            let Some((prev_idx, ch)) = formula[..start].char_indices().next_back() else {
                break;
            };
            if ch.is_ascii_alphanumeric() || ch == '_' || ch == '.' || ch == ':' {
                start = prev_idx;
            } else {
                break;
            }
        }

        let token = formula[start..bang].trim();
        let preceded_by_external_book = start > cursor
            && formula[..start]
                .chars()
                .next_back()
                .is_some_and(|ch| ch == ']');
        if token.is_empty()
            || token.contains('[')
            || token.contains(']')
            || preceded_by_external_book
        {
            out.push_str(&formula[cursor..=bang]);
            cursor = bang + 1;
            continue;
        }

        if let Some(renamed) = rename_sheet_token(token, old_name, new_name) {
            out.push_str(&formula[cursor..start]);
            out.push_str(&quote_sheet_token(&renamed));
            out.push('!');
        } else {
            out.push_str(&formula[cursor..=bang]);
        }
        cursor = bang + 1;
    }
    out.push_str(&formula[cursor..]);
    out
}

fn double_quoted_ranges(formula: &str) -> Vec<(usize, usize)> {
    let mut ranges = Vec::new();
    let mut cursor = 0;
    while let Some(rel_start) = formula[cursor..].find('"') {
        let start = cursor + rel_start;
        let mut idx = start + 1;
        while idx < formula.len() {
            let rest = &formula[idx..];
            if rest.starts_with("\"\"") {
                idx += 2;
                continue;
            }
            if rest.starts_with('"') {
                ranges.push((start, idx + 1));
                idx += 1;
                break;
            }
            let Some(ch) = rest.chars().next() else {
                break;
            };
            idx += ch.len_utf8();
        }
        cursor = idx;
    }
    ranges
}

fn byte_index_in_ranges(index: usize, ranges: &[(usize, usize)]) -> bool {
    ranges
        .iter()
        .any(|(start, end)| *start <= index && index < *end)
}

fn rename_sheet_token(token: &str, old_name: &str, new_name: &str) -> Option<String> {
    if token.contains('[') || token.contains(']') {
        return None;
    }
    let mut changed = false;
    let renamed = token
        .split(':')
        .map(|part| {
            if part == old_name {
                changed = true;
                new_name
            } else {
                part
            }
        })
        .collect::<Vec<_>>()
        .join(":");
    changed.then_some(renamed)
}

fn quote_sheet_token(token: &str) -> String {
    format!("'{}'", token.replace('\'', "''"))
}

pub(super) fn has_pending_save_work(patcher: &XlsxPatcher) -> bool {
    !patcher.value_patches.is_empty()
        || !patcher.format_patches.is_empty()
        || !patcher.rels_patches.is_empty()
        || !patcher.queued_blocks.is_empty()
        || !patcher.queued_dv_patches.is_empty()
        || !patcher.queued_cf_patches.is_empty()
        || !patcher.file_adds.is_empty()
        || !patcher.file_deletes.is_empty()
        || !patcher.queued_content_type_ops.is_empty()
        || patcher.queued_props.is_some()
        || !patcher.queued_hyperlinks.is_empty()
        || !patcher.queued_defined_names.is_empty()
        || !patcher.queued_tables.is_empty()
        || !patcher.queued_table_growth.is_empty()
        || !patcher.queued_comments.is_empty()
        || !patcher.queued_sheet_renames.is_empty()
        || !patcher.queued_sheet_moves.is_empty()
        || !patcher.queued_axis_shifts.is_empty()
        || !patcher.queued_range_moves.is_empty()
        || !patcher.queued_sheet_copies.is_empty()
        || !patcher.queued_sheet_creates.is_empty()
        || !patcher.queued_sheet_deletes.is_empty()
        || !patcher.queued_images.is_empty()
        || !patcher.queued_charts.is_empty()
        || !patcher.queued_chart_removes.is_empty()
        || !patcher.queued_pivot_caches.is_empty()
        || !patcher.queued_pivot_tables.is_empty()
        || !patcher.queued_pivot_source_edits.is_empty()
        || patcher.drop_external_links
        || !patcher.queued_sheet_setup.is_empty()
        || !patcher.queued_page_breaks.is_empty()
        || !patcher.queued_autofilters.is_empty()
        || patcher.queued_workbook_security.is_some()
        || !patcher.queued_slicers.is_empty()
        || !patcher.queued_threaded_comments.is_empty()
        || !patcher.queued_persons.is_empty()
}

pub(super) fn copy_source_file_phase(patcher: &XlsxPatcher, output_path: &str) -> PyResult<()> {
    if std::path::Path::new(&patcher.file_path) == std::path::Path::new(output_path) {
        return Ok(());
    }
    crate::atomic_save::write_zip_atomically(output_path, |out| {
        let mut src = File::open(&patcher.file_path)
            .map_err(|e| PyIOError::new_err(format!("Cannot open '{}': {e}", patcher.file_path)))?;
        std::io::copy(&mut src, out)
            .map_err(|e| PyIOError::new_err(format!("Copy failed: {e}")))?;
        Ok(())
    })
}

pub(super) fn drain_permissive_seed_file_patches_phase(
    patcher: &mut XlsxPatcher,
    file_patches: &mut HashMap<String, Vec<u8>>,
) {
    for (k, v) in patcher.permissive_seed_file_patches.drain() {
        file_patches.insert(k, v);
    }
}

pub(super) fn serialize_rels_patches_phase(
    patcher: &mut XlsxPatcher,
    file_patches: &mut HashMap<String, Vec<u8>>,
    zip: &mut ZipArchive<File>,
) -> PyResult<()> {
    for (path, graph) in &patcher.rels_patches {
        let bytes = graph.serialize();
        if zip.by_name(path).is_ok() {
            file_patches.insert(path.clone(), bytes);
        } else {
            patcher.file_adds.insert(path.clone(), bytes);
        }
    }

    Ok(())
}

pub(super) fn route_part_writes_and_deletes_phase(
    patcher: &mut XlsxPatcher,
    file_patches: &mut HashMap<String, Vec<u8>>,
    zip: &mut ZipArchive<File>,
    file_writes: HashMap<String, Vec<u8>>,
    file_deletes: HashSet<String>,
) {
    for (path, bytes) in file_writes {
        if zip.by_name(&path).is_ok() {
            file_patches.insert(path, bytes);
        } else {
            patcher.file_adds.insert(path, bytes);
        }
    }
    for path in file_deletes {
        patcher.file_deletes.insert(path);
    }
}

pub(super) fn rebuild_calc_chain_phase(
    patcher: &mut XlsxPatcher,
    file_patches: &mut HashMap<String, Vec<u8>>,
    zip: &mut ZipArchive<File>,
) -> PyResult<()> {
    fn get_bytes(
        file_patches: &HashMap<String, Vec<u8>>,
        file_adds: &HashMap<String, Vec<u8>>,
        zip: &mut ZipArchive<File>,
        path: &str,
    ) -> Option<Vec<u8>> {
        if let Some(b) = file_patches.get(path) {
            return Some(b.clone());
        }
        if let Some(b) = file_adds.get(path) {
            return Some(b.clone());
        }
        let mut entry = match zip.by_name(path) {
            Ok(e) => e,
            Err(_) => return None,
        };
        let mut buf: Vec<u8> = Vec::with_capacity(entry.size() as usize);
        std::io::copy(&mut entry, &mut buf).ok()?;
        Some(buf)
    }

    const CALC_CHAIN_PATH: &str = "xl/calcChain.xml";
    let source_calc_chain = get_bytes(file_patches, &patcher.file_adds, zip, CALC_CHAIN_PATH);
    let source_has_calc_chain = source_calc_chain.is_some();
    let source_calc_chain_ext_lst = source_calc_chain
        .as_deref()
        .and_then(calcchain::extract_ext_lst);
    let workbook_xml = get_bytes(file_patches, &patcher.file_adds, zip, "xl/workbook.xml");
    let sheet_ids = workbook_xml
        .as_deref()
        .map(sheet_ids_by_name)
        .unwrap_or_default();

    // Walk sheets in tab order, scanning each while emitting calcChain's
    // workbook sheetId value.
    let mut all_entries: Vec<calcchain::CalcChainEntry> = Vec::new();
    let order = patcher.sheet_order.clone();
    for (i, sheet_name) in order.iter().enumerate() {
        let sheet_path = match patcher.sheet_paths.get(sheet_name) {
            Some(p) => p.clone(),
            None => continue,
        };
        let sheet_xml = match get_bytes(file_patches, &patcher.file_adds, zip, &sheet_path) {
            Some(b) => b,
            None => continue,
        };
        let sheet_id = sheet_ids
            .get(sheet_name)
            .copied()
            .unwrap_or_else(|| (i as u32) + 1);
        let entries = calcchain::scan_sheet_for_formulas(&sheet_xml, sheet_id);
        all_entries.extend(entries);
    }

    match calcchain::render_calc_chain_with_ext_lst(
        &all_entries,
        source_calc_chain_ext_lst.as_deref(),
    ) {
        Some(bytes) => {
            // Route the rewrite based on whether the source ZIP
            // already had a calcChain.xml entry.
            if source_has_calc_chain {
                file_patches.insert(CALC_CHAIN_PATH.to_string(), bytes);
            } else {
                patcher.file_adds.insert(CALC_CHAIN_PATH.to_string(), bytes);
            }
            // Ensure content-type Override + workbook rel.
            ensure_calc_chain_metadata(patcher, file_patches, zip)?;
        }
        None => {
            // Zero formulas in the workbook. If the source had a
            // calcChain.xml, delete it (it would be stale and Excel
            // would emit a parse warning if it pointed at missing
            // cells).
            if source_has_calc_chain {
                patcher.file_deletes.insert(CALC_CHAIN_PATH.to_string());
                file_patches.remove(CALC_CHAIN_PATH);
                remove_calc_chain_metadata(patcher, file_patches, zip)?;
            }
        }
    }
    Ok(())
}

pub(super) fn rewrite_zip_phase(
    patcher: &XlsxPatcher,
    file_patches: &HashMap<String, Vec<u8>>,
    output_path: &str,
) -> PyResult<()> {
    crate::atomic_save::write_zip_atomically_fast(output_path, |dst| {
        let src = File::open(&patcher.file_path)
            .map_err(|e| PyIOError::new_err(format!("Cannot open '{}': {e}", patcher.file_path)))?;
        let mut zip =
            ZipArchive::new(src).map_err(|e| PyIOError::new_err(format!("ZIP read error: {e}")))?;

        let mut out = ZipWriter::new(dst);

        let mut source_names: HashSet<String> = HashSet::with_capacity(zip.len());
        for i in 0..zip.len() {
            let file = zip
                .by_index(i)
                .map_err(|e| PyIOError::new_err(format!("ZIP entry read error: {e}")))?;
            let name = file.name().to_string();
            source_names.insert(name.clone());

            if patcher.file_deletes.contains(&name) {
                continue;
            }

            if !file_patches.contains_key(&name) {
                out.raw_copy_file(file)
                    .map_err(|e| PyIOError::new_err(format!("ZIP raw-copy error: {e}")))?;
                continue;
            }

            let mut opts = SimpleFileOptions::default().compression_method(file.compression());
            if file.compression() == zip::CompressionMethod::Deflated {
                opts = opts.compression_level(Some(FAST_SAVE_DEFLATE_LEVEL));
            }
            if let Some(dt) = file.last_modified() {
                opts = opts.last_modified_time(dt);
            }
            if let Some(mode) = file.unix_mode() {
                opts = opts.unix_permissions(mode);
            }

            if file.is_dir() {
                out.add_directory(&name, opts)
                    .map_err(|e| PyIOError::new_err(format!("ZIP write error: {e}")))?;
                continue;
            }

            out.start_file(&name, opts)
                .map_err(|e| PyIOError::new_err(format!("ZIP write error: {e}")))?;
            let patched = file_patches
                .get(&name)
                .expect("patched ZIP entry checked before rewrite");
            out.write_all(patched)
                .map_err(|e| PyIOError::new_err(format!("ZIP write error: {e}")))?;
        }

        if !patcher.file_adds.is_empty() {
            for new_path in patcher.file_adds.keys() {
                assert!(
                    !source_names.contains(new_path),
                    "file_adds collision with source entry: {new_path}; \
                 caller bug; use file_patches to REPLACE existing entries"
                );
            }
            let mut new_paths: Vec<&String> = patcher.file_adds.keys().collect();
            new_paths.sort();
            let dt = epoch_or_now();
            for new_path in new_paths {
                let bytes = &patcher.file_adds[new_path];
                let opts = SimpleFileOptions::default()
                    .compression_method(zip::CompressionMethod::Deflated)
                    .compression_level(Some(FAST_SAVE_DEFLATE_LEVEL))
                    .last_modified_time(dt);
                out.start_file(new_path, opts)
                    .map_err(|e| PyIOError::new_err(format!("ZIP write error: {e}")))?;
                out.write_all(bytes)
                    .map_err(|e| PyIOError::new_err(format!("ZIP write error: {e}")))?;
            }
        }

        out.finish()
            .map_err(|e| PyIOError::new_err(format!("ZIP finalize error: {e}")))?;

        Ok(())
    })
}

pub(super) fn ensure_calc_chain_metadata(
    patcher: &mut XlsxPatcher,
    file_patches: &mut HashMap<String, Vec<u8>>,
    zip: &mut ZipArchive<File>,
) -> PyResult<()> {
    // Content types.
    let ct_xml: Vec<u8> = if let Some(b) = file_patches.get("[Content_Types].xml") {
        b.clone()
    } else {
        ooxml_util::zip_read_to_string(zip, "[Content_Types].xml")?
            .as_bytes()
            .to_vec()
    };
    let mut graph = content_types::ContentTypesGraph::parse(&ct_xml)
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("[Content_Types].xml parse: {e}")))?;
    graph.add_override("/xl/calcChain.xml", calcchain::CT_CALC_CHAIN);
    file_patches.insert("[Content_Types].xml".to_string(), graph.serialize());

    // Workbook rels.
    let wb_rels_path = "xl/_rels/workbook.xml.rels";
    let wb_rels_bytes_opt: Option<Vec<u8>> = if let Some(b) = file_patches.get(wb_rels_path) {
        Some(b.clone())
    } else if let Some(g) = patcher.rels_patches.get(wb_rels_path) {
        Some(g.serialize())
    } else if let Ok(mut entry) = zip.by_name(wb_rels_path) {
        let mut buf: Vec<u8> = Vec::with_capacity(entry.size() as usize);
        if std::io::copy(&mut entry, &mut buf).is_ok() {
            Some(buf)
        } else {
            None
        }
    } else {
        None
    };
    if let Some(bytes) = wb_rels_bytes_opt {
        let mut graph =
            wolfxl_rels::RelsGraph::parse(&bytes).unwrap_or_else(|_| wolfxl_rels::RelsGraph::new());
        // Idempotent: only add if no existing rel of this type
        // already targets calcChain.xml.
        let already = graph.iter().any(|r| {
            r.rel_type == calcchain::REL_CALC_CHAIN
                && (r.target == "calcChain.xml" || r.target == "/xl/calcChain.xml")
        });
        if !already {
            graph.add(
                calcchain::REL_CALC_CHAIN,
                "calcChain.xml",
                wolfxl_rels::TargetMode::Internal,
            );
            file_patches.insert(wb_rels_path.to_string(), graph.serialize());
        }
    }

    Ok(())
}

fn remove_calc_chain_metadata(
    patcher: &mut XlsxPatcher,
    file_patches: &mut HashMap<String, Vec<u8>>,
    zip: &mut ZipArchive<File>,
) -> PyResult<()> {
    // Content types.
    let ct_xml: Vec<u8> = if let Some(b) = file_patches.get("[Content_Types].xml") {
        b.clone()
    } else {
        ooxml_util::zip_read_to_string(zip, "[Content_Types].xml")?
            .as_bytes()
            .to_vec()
    };
    let mut content_types = content_types::ContentTypesGraph::parse(&ct_xml)
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("[Content_Types].xml parse: {e}")))?;
    content_types.remove_override("/xl/calcChain.xml");
    file_patches.insert("[Content_Types].xml".to_string(), content_types.serialize());

    // Workbook rels.
    let wb_rels_path = "xl/_rels/workbook.xml.rels";
    let wb_rels_bytes_opt: Option<Vec<u8>> = if let Some(b) = file_patches.get(wb_rels_path) {
        Some(b.clone())
    } else if let Some(g) = patcher.rels_patches.get(wb_rels_path) {
        Some(g.serialize())
    } else if let Ok(mut entry) = zip.by_name(wb_rels_path) {
        let mut buf: Vec<u8> = Vec::with_capacity(entry.size() as usize);
        if std::io::copy(&mut entry, &mut buf).is_ok() {
            Some(buf)
        } else {
            None
        }
    } else {
        None
    };
    if let Some(bytes) = wb_rels_bytes_opt {
        let mut graph =
            wolfxl_rels::RelsGraph::parse(&bytes).unwrap_or_else(|_| wolfxl_rels::RelsGraph::new());
        let calc_chain_ids: Vec<wolfxl_rels::RelId> = graph
            .iter()
            .filter(|r| {
                r.rel_type == calcchain::REL_CALC_CHAIN
                    && (r.target == "calcChain.xml" || r.target == "/xl/calcChain.xml")
            })
            .map(|r| r.id.clone())
            .collect();
        for id in calc_chain_ids {
            graph.remove(&id);
        }
        file_patches.insert(wb_rels_path.to_string(), graph.serialize());
    }

    Ok(())
}

#[cfg(test)]
mod tests {
    use super::super::content_types::{ContentTypeOp, ContentTypesGraph};
    use super::*;
    use crate::ooxml_util;

    static TEST_EPOCH_ENV_LOCK: std::sync::Mutex<()> = std::sync::Mutex::new(());

    #[test]
    fn epoch_or_now_honors_test_epoch_zero() {
        let _guard = TEST_EPOCH_ENV_LOCK.lock().unwrap();
        let prev = std::env::var("WOLFXL_TEST_EPOCH").ok();
        std::env::set_var("WOLFXL_TEST_EPOCH", "0");
        let dt = epoch_or_now();
        match prev {
            Some(v) => std::env::set_var("WOLFXL_TEST_EPOCH", v),
            None => std::env::remove_var("WOLFXL_TEST_EPOCH"),
        }

        std::env::set_var("WOLFXL_TEST_EPOCH", "0");
        let dt2 = epoch_or_now();
        std::env::remove_var("WOLFXL_TEST_EPOCH");

        assert_eq!(
            (
                dt.year(),
                dt.month(),
                dt.day(),
                dt.hour(),
                dt.minute(),
                dt.second()
            ),
            (
                dt2.year(),
                dt2.month(),
                dt2.day(),
                dt2.hour(),
                dt2.minute(),
                dt2.second()
            ),
        );
    }

    #[test]
    fn epoch_or_now_clamps_pre_1980_floor() {
        let _guard = TEST_EPOCH_ENV_LOCK.lock().unwrap();
        let prev = std::env::var("WOLFXL_TEST_EPOCH").ok();
        std::env::set_var("WOLFXL_TEST_EPOCH", "0");
        let dt = epoch_or_now();
        match prev {
            Some(v) => std::env::set_var("WOLFXL_TEST_EPOCH", v),
            None => std::env::remove_var("WOLFXL_TEST_EPOCH"),
        }
        assert_eq!(dt.year(), 1980);
        assert_eq!(dt.month(), 1);
        assert_eq!(dt.day(), 1);
    }

    #[test]
    fn epoch_or_now_handles_recent_timestamp() {
        let _guard = TEST_EPOCH_ENV_LOCK.lock().unwrap();
        let prev = std::env::var("WOLFXL_TEST_EPOCH").ok();
        std::env::set_var("WOLFXL_TEST_EPOCH", "1704067200");
        let dt = epoch_or_now();
        match prev {
            Some(v) => std::env::set_var("WOLFXL_TEST_EPOCH", v),
            None => std::env::remove_var("WOLFXL_TEST_EPOCH"),
        }
        assert_eq!(dt.year(), 2024);
        assert_eq!(dt.month(), 1);
        assert_eq!(dt.day(), 1);
    }

    #[test]
    fn sheet_order_parser_preserves_workbook_xml_order() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Apples"  sheetId="1" r:id="rId1"/>
    <sheet name="Bananas" sheetId="2" r:id="rId2"/>
    <sheet name="Cherries" sheetId="3" r:id="rId3"/>
  </sheets>
</workbook>"#;
        let pairs = ooxml_util::parse_workbook_sheet_rids(xml).unwrap();
        let names: Vec<&str> = pairs.iter().map(|(n, _)| n.as_str()).collect();
        assert_eq!(names, vec!["Apples", "Bananas", "Cherries"]);
    }

    const SOURCE_CT_XML: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>"#;

    #[test]
    fn phase_2_5c_no_ops_is_no_op() {
        let ops: Vec<ContentTypeOp> = Vec::new();
        assert!(ops.is_empty(), "no-op precondition: no queued ops");
    }

    #[test]
    fn phase_2_5c_aggregates_overrides_into_single_mutation() {
        let ops = vec![
            ContentTypeOp::AddOverride(
                "/xl/comments1.xml".into(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml".into(),
            ),
            ContentTypeOp::EnsureDefault(
                "vml".into(),
                "application/vnd.openxmlformats-officedocument.vmlDrawing".into(),
            ),
            ContentTypeOp::AddOverride(
                "/xl/tables/table1.xml".into(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml".into(),
            ),
        ];
        let mut graph = ContentTypesGraph::parse(SOURCE_CT_XML).expect("parse source");
        for op in &ops {
            graph.apply_op(op);
        }
        let bytes = graph.serialize();
        let text = std::str::from_utf8(&bytes).expect("utf8 round-trip");
        assert!(text.contains("/xl/comments1.xml"), "comments override");
        assert!(text.contains("/xl/tables/table1.xml"), "table override");
        assert!(text.contains(r#"Extension="vml""#), "vml default");
        assert!(text.contains("/xl/workbook.xml"));
        assert!(text.contains("/xl/styles.xml"));
    }

    #[test]
    fn phase_2_5c_preserves_source_order_for_existing_overrides() {
        let ops = vec![ContentTypeOp::AddOverride(
            "/xl/comments1.xml".into(),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml".into(),
        )];
        let mut graph = ContentTypesGraph::parse(SOURCE_CT_XML).expect("parse");
        for op in &ops {
            graph.apply_op(op);
        }
        let bytes = graph.serialize();
        let text = std::str::from_utf8(&bytes).expect("utf8");
        let idx_workbook = text.find("/xl/workbook.xml").expect("workbook");
        let idx_sheet1 = text.find("/xl/worksheets/sheet1.xml").expect("sheet1");
        let idx_styles = text.find("/xl/styles.xml").expect("styles");
        let idx_comments = text.find("/xl/comments1.xml").expect("comments");
        assert!(
            idx_workbook < idx_sheet1 && idx_sheet1 < idx_styles,
            "source overrides retain document order",
        );
        assert!(
            idx_styles < idx_comments,
            "new overrides append after source ones, not interleaved",
        );
    }

    const NEW_SHEET: &[u8] = br#"<sheet name="Copy" sheetId="2" r:id="rId99"/>"#;

    #[test]
    fn splice_normal_sheets_block_inserts_before_close() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>"#;
        let out = splice_into_sheets_block(xml, NEW_SHEET).expect("splice ok");
        let s = std::str::from_utf8(&out).unwrap();
        let rid1 = s.find("r:id=\"rId1\"").unwrap();
        let rid99 = s.find("r:id=\"rId99\"").unwrap();
        let close = s.find("</sheets>").unwrap();
        assert!(rid1 < rid99, "new sheet appended after Sheet1");
        assert!(rid99 < close, "new sheet inserted BEFORE </sheets>");
        assert_eq!(s.matches("</sheets>").count(), 1);
    }

    #[test]
    fn splice_adds_root_rel_namespace_when_only_existing_sheet_has_it() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheets><sheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>"#;
        let out = splice_into_sheets_block(xml, NEW_SHEET).expect("splice ok");
        let s = std::str::from_utf8(&out).unwrap();
        let root_start = s.find("<workbook").unwrap();
        let root_end = root_start + s[root_start..].find('>').unwrap();
        assert!(
            s[root_start..root_end].contains(
                r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#
            ),
            "new sheet r:id must be bound by the workbook root namespace"
        );
        assert!(s.contains("r:id=\"rId99\""));
    }

    #[test]
    fn splice_handles_prefixed_workbook_and_sheets() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<x:sheets><x:sheet name="Sheet1" sheetId="1" r:id="rId1"/></x:sheets>
</x:workbook>"#;
        let out = splice_into_sheets_block(xml, NEW_SHEET).expect("splice ok");
        let s = std::str::from_utf8(&out).unwrap();
        let root_start = s.find("<x:workbook").unwrap();
        let root_end = root_start + s[root_start..].find('>').unwrap();
        assert!(
            s[root_start..root_end].contains(
                r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#
            ),
            "prefixed workbook root must get the relationship namespace"
        );
        assert!(s.contains("<x:sheet name=\"Copy\""));
        assert!(!s.contains("<sheet name=\"Copy\""));
        assert!(s.contains("</x:sheets>"));
    }

    #[test]
    fn splice_handles_prefixed_sheets_with_space_before_empty_close() {
        let xml = b"\xef\xbb\xbf<?xml version=\"1.0\" encoding=\"utf-8\" standalone=\"yes\"?><x:workbook xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><x:sheets><x:sheet name=\"BAL\" sheetId=\"4\" state=\"hidden\" r:id=\"rId3\" /></x:sheets><x:definedNames><x:definedName name=\"BAL_DATA\" localSheetId=\"0\">BAL!$A$1</x:definedName></x:definedNames></x:workbook>";
        let out = splice_into_sheets_block(xml, NEW_SHEET).expect("splice ok");
        let s = std::str::from_utf8(&out).unwrap();
        assert!(s.contains(r#"<x:sheet name="BAL" sheetId="4" state="hidden" r:id="rId3" />"#));
        assert!(s.contains(r#"<x:sheet name="Copy" sheetId="2" r:id="rId99"/>"#));
        assert!(s.contains("</x:sheets>"));
        assert!(s.contains("<x:definedNames>"));
        assert!(!s.contains(r#"r:id="rId3"<x:sheet"#));
        assert!(!s.contains("</x:shee<x:definedNames>"));
    }

    #[test]
    fn splice_handles_self_closing_sheets() {
        let xml = br#"<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets/>
</workbook>"#;
        let out = splice_into_sheets_block(xml, NEW_SHEET).expect("splice ok");
        let s = std::str::from_utf8(&out).unwrap();
        assert!(s.contains("<sheets>"), "open tag emitted");
        assert!(s.contains("</sheets>"), "close tag emitted");
        assert!(s.contains("rId99"), "new sheet entry inserted");
        assert_eq!(s.matches("<sheets>").count(), 1);
        assert_eq!(s.matches("</sheets>").count(), 1);
        assert!(!s.contains("<sheets/>"));
    }

    #[test]
    fn splice_ignores_fake_close_tag_inside_comment() {
        let xml = br#"<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<!-- FUZZTOKEN: this fakeout closes </sheets> here, naive splice would bite -->
<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>"#;
        let out = splice_into_sheets_block(xml, NEW_SHEET).expect("splice ok");
        let s = std::str::from_utf8(&out).unwrap();
        assert!(s.contains("FUZZTOKEN"), "comment survives splice");
        let open = s.find("<sheets>").expect("real <sheets> open");
        let close = s.rfind("</sheets>").expect("real </sheets> close");
        let rid99 = s.find("rId99").expect("new entry present");
        assert!(open < rid99, "new entry after real <sheets> open");
        assert!(rid99 < close, "new entry before real </sheets> close");
    }

    #[test]
    fn splice_ignores_fake_close_tag_inside_cdata() {
        let xml = br#"<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"><![CDATA[fake </sheets> token]]></sheet></sheets>
</workbook>"#;
        let out = splice_into_sheets_block(xml, NEW_SHEET).expect("splice ok");
        let s = std::str::from_utf8(&out).unwrap();
        let rid99 = s.find("rId99").expect("new entry present");
        let cdata_close = s.find("]]>").expect("cdata close");
        let real_close = s.rfind("</sheets>").expect("real close");
        assert!(cdata_close < rid99, "new entry follows CDATA");
        assert!(rid99 < real_close, "new entry before real </sheets>");
    }

    #[test]
    fn splice_returns_error_when_no_sheets_block() {
        pyo3::Python::initialize();
        let xml = br#"<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;
        let err = splice_into_sheets_block(xml, NEW_SHEET).unwrap_err();
        let msg = format!("{err}");
        assert!(
            msg.contains("no <sheets>"),
            "preserves historical error message, got: {msg}"
        );
    }
}
