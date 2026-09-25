//! Sheet-scoped block save phases for the surgical xlsx patcher.

use std::collections::{HashMap, HashSet};
use std::fs::File;

use pyo3::exceptions::{PyIOError, PyValueError};
use pyo3::prelude::*;
use zip::ZipArchive;

use crate::ooxml_util;

use super::patcher_workbook::{
    load_or_empty_rels, minimal_styles_xml, parse_n_from_part_path, patched_or_source_part_bytes,
    sheet_rels_path_for,
};
use super::{
    autofilter, autofilter_helpers, comments, content_types, hyperlinks, sheet_patcher, tables,
    threaded_comments, XlsxPatcher,
};
use sheet_patcher::CellPatch;
use wolfxl_merger::SheetBlock;
use wolfxl_rels::{rt, PartIdAllocator, RelsGraph};

pub(super) fn apply_data_validations_phase(
    patcher: &XlsxPatcher,
    local_blocks: &mut HashMap<String, Vec<SheetBlock>>,
    zip: &mut ZipArchive<File>,
) -> PyResult<()> {
    for (sheet_name, patches) in &patcher.queued_dv_patches {
        let sheet_path = match patcher.sheet_paths.get(sheet_name) {
            Some(p) => p,
            None => continue,
        };
        let xml = ooxml_util::zip_read_to_string(zip, sheet_path)?;
        let existing = super::validations::extract_existing_dv_block(&xml);
        let block_bytes =
            super::validations::build_data_validations_block(existing.as_deref(), patches);
        local_blocks
            .entry(sheet_path.clone())
            .or_default()
            .push(SheetBlock::DataValidations(block_bytes));
    }

    Ok(())
}

pub(super) fn apply_conditional_formatting_phase(
    patcher: &XlsxPatcher,
    local_blocks: &mut HashMap<String, Vec<SheetBlock>>,
    styles_xml: &mut Option<String>,
    zip: &mut ZipArchive<File>,
) -> PyResult<()> {
    let mut new_dxfs_total: Vec<super::conditional_formatting::DxfPatch> = Vec::new();
    let mut styles_loaded: Option<String> = None;
    let mut running_dxf_count: u32 = 0;
    let mut cf_sheet_names: Vec<&String> = patcher.queued_cf_patches.keys().collect();
    cf_sheet_names.sort();

    for sheet_name in cf_sheet_names {
        let patches = &patcher.queued_cf_patches[sheet_name];
        let sheet_path = match patcher.sheet_paths.get(sheet_name) {
            Some(p) => p,
            None => continue,
        };
        let xml = ooxml_util::zip_read_to_string(zip, sheet_path)?;

        if styles_loaded.is_none() {
            let raw = ooxml_util::zip_read_to_string_opt(zip, "xl/styles.xml")?
                .unwrap_or_else(minimal_styles_xml);
            running_dxf_count = super::conditional_formatting::count_dxfs(&raw);
            styles_loaded = Some(raw);
        }

        let existing = super::conditional_formatting::extract_existing_cf_blocks(&xml);
        let pmax = super::conditional_formatting::scan_max_cf_priority(&xml);
        let element_prefix = super::conditional_formatting::main_xml_prefix(&xml, b"worksheet");
        let result = super::conditional_formatting::build_cf_blocks_with_prefix(
            &existing,
            patches,
            pmax,
            running_dxf_count,
            &element_prefix,
        );
        running_dxf_count += result.new_dxfs.len() as u32;
        new_dxfs_total.extend(result.new_dxfs);
        local_blocks
            .entry(sheet_path.clone())
            .or_default()
            .push(SheetBlock::ConditionalFormatting(result.block_bytes));
    }

    if !new_dxfs_total.is_empty() {
        let base = match styles_xml.take() {
            Some(s) => s,
            None => styles_loaded.unwrap_or_else(minimal_styles_xml),
        };
        let element_prefix = super::conditional_formatting::main_xml_prefix(&base, b"styleSheet");
        let new_dxfs_xml: String = new_dxfs_total
            .iter()
            .map(|dxf| super::conditional_formatting::dxf_to_xml_with_prefix(dxf, &element_prefix))
            .collect::<Vec<_>>()
            .join("");
        let updated = super::conditional_formatting::ensure_dxfs_section(&base, &new_dxfs_xml);
        *styles_xml = Some(updated);
    }

    Ok(())
}

pub(super) fn apply_hyperlinks_phase(
    patcher: &mut XlsxPatcher,
    local_blocks: &mut HashMap<String, Vec<SheetBlock>>,
    zip: &mut ZipArchive<File>,
) -> PyResult<()> {
    let sheet_order_local: Vec<String> = patcher.sheet_order.clone();
    for sheet_name in &sheet_order_local {
        let ops = match patcher.queued_hyperlinks.get(sheet_name) {
            Some(o) if !o.is_empty() => o.clone(),
            _ => continue,
        };
        let sheet_path = match patcher.sheet_paths.get(sheet_name).cloned() {
            Some(p) => p,
            None => continue,
        };
        let rels_path = sheet_rels_path_for(&sheet_path);
        patcher
            .ancillary
            .populate_for_sheet(zip, sheet_name, &sheet_path)
            .map_err(|e| {
                PyIOError::new_err(format!("ancillary populate for '{sheet_name}': {e}"))
            })?;
        if !patcher.rels_patches.contains_key(&rels_path) {
            let graph = load_or_empty_rels(zip, &rels_path)?;
            patcher.rels_patches.insert(rels_path.clone(), graph);
        }
        let rels = patcher
            .rels_patches
            .get_mut(&rels_path)
            .expect("just inserted above");
        let xml = ooxml_util::zip_read_to_string(zip, &sheet_path)?;
        let existing = hyperlinks::extract_hyperlinks(xml.as_bytes(), rels);
        let had_existing = !existing.is_empty();
        let (block_bytes, _deleted_rids) = hyperlinks::build_hyperlinks_block(existing, &ops, rels);

        if block_bytes.is_empty() && !had_existing {
            continue;
        }
        local_blocks
            .entry(sheet_path)
            .or_default()
            .push(SheetBlock::Hyperlinks(block_bytes));
    }

    Ok(())
}

pub(super) fn apply_tables_phase(
    patcher: &mut XlsxPatcher,
    local_blocks: &mut HashMap<String, Vec<SheetBlock>>,
    file_patches: &HashMap<String, Vec<u8>>,
    zip: &mut ZipArchive<File>,
    part_id_allocator: &mut PartIdAllocator,
    cloned_table_names: &HashSet<String>,
) -> PyResult<()> {
    if patcher.queued_tables.is_empty() {
        return Ok(());
    }

    let mut tables_inventory = tables::scan_existing_tables(zip)
        .map_err(|e| PyIOError::new_err(format!("scan tables: {e}")))?;

    for name in cloned_table_names {
        tables_inventory.names.insert(name.clone());
    }

    let sheet_order_local: Vec<String> = patcher.sheet_order.clone();
    for sheet_name in &sheet_order_local {
        let patches = match patcher.queued_tables.get(sheet_name) {
            Some(p) if !p.is_empty() => p.clone(),
            _ => continue,
        };
        let sheet_path = match patcher.sheet_paths.get(sheet_name).cloned() {
            Some(p) => p,
            None => continue,
        };
        let rels_path = sheet_rels_path_for(&sheet_path);
        if !patcher.rels_patches.contains_key(&rels_path) {
            let graph = if let Some(bytes) = patcher.file_adds.get(&rels_path) {
                RelsGraph::parse(bytes).map_err(|e| {
                    PyIOError::new_err(format!("rels parse for cloned '{rels_path}': {e}"))
                })?
            } else if let Some(bytes) = file_patches.get(&rels_path) {
                RelsGraph::parse(bytes).map_err(|e| {
                    PyIOError::new_err(format!("rels parse for patched '{rels_path}': {e}"))
                })?
            } else {
                load_or_empty_rels(zip, &rels_path)?
            };
            patcher.rels_patches.insert(rels_path.clone(), graph);
        }
        let rels = patcher
            .rels_patches
            .get_mut(&rels_path)
            .expect("just inserted above");
        let result =
            tables::build_tables(&patches, &tables_inventory, rels, Some(part_id_allocator))
                .map_err(PyValueError::new_err)?;

        for (path, _bytes) in &result.table_parts {
            tables_inventory.count += 1;
            tables_inventory.paths.push(path.clone());
        }
        for patch in &patches {
            tables_inventory.names.insert(patch.name.clone());
        }
        for (path, bytes) in result.table_parts {
            patcher.file_adds.insert(path, bytes);
        }
        for path in &tables_inventory.paths {
            if let Some(bytes) = patcher.file_adds.get(path) {
                let (id_opt, _) = tables::parse_table_root_attrs(bytes);
                if let Some(id) = id_opt {
                    tables_inventory.ids.insert(id);
                }
            }
        }

        let content_type_ops = patcher
            .queued_content_type_ops
            .entry(sheet_name.clone())
            .or_default();
        for (part_name, content_type) in result.new_content_types {
            content_type_ops.push(content_types::ContentTypeOp::AddOverride(
                part_name,
                content_type,
            ));
        }
        if !result.table_parts_block.is_empty() {
            local_blocks
                .entry(sheet_path)
                .or_default()
                .push(SheetBlock::TableParts(result.table_parts_block));
        }
    }

    Ok(())
}

/// Grow existing tables that the Python side widened or lengthened in
/// modify mode. Each queued [`tables::TableGrowthPatch`] resolves its
/// part through the sheet's table relationships and rewrites only the
/// table and autoFilter ranges and the appended column entries. Runs
/// before Phase 3 so later structural shifts see the grown ranges.
pub(super) fn apply_table_growth_phase(
    patcher: &XlsxPatcher,
    file_patches: &mut HashMap<String, Vec<u8>>,
    zip: &mut ZipArchive<File>,
) -> PyResult<()> {
    for sheet_name in &patcher.sheet_order {
        let Some(patches) = patcher.queued_table_growth.get(sheet_name) else {
            continue;
        };
        let Some(sheet_path) = patcher.sheet_paths.get(sheet_name) else {
            continue;
        };
        for patch in patches {
            let part_path = tables::find_table_part_by_name(zip, sheet_path, &patch.table_name)
                .map_err(|error| {
                    PyValueError::new_err(format!(
                        "resolve table '{}' on '{sheet_name}': {error}",
                        patch.table_name
                    ))
                })?
                .ok_or_else(|| {
                    PyValueError::new_err(format!(
                        "table '{}' was not found on '{sheet_name}'",
                        patch.table_name
                    ))
                })?;
            let xml =
                patched_or_source_part_bytes(file_patches, zip, &part_path).ok_or_else(|| {
                    PyIOError::new_err(format!("table part '{part_path}' is missing"))
                })?;
            let patched = tables::patch_table_growth(&xml, patch).map_err(|error| {
                PyValueError::new_err(format!(
                    "grow table '{}' on '{sheet_name}': {error}",
                    patch.table_name
                ))
            })?;
            file_patches.insert(part_path, patched);
        }
    }
    Ok(())
}

/// Drain `queued_threaded_comments` (per sheet) and `queued_persons`
/// (workbook scope) into fresh `threadedCommentsN.xml` + `personList.xml`
/// part bytes. RFC-068 G08 step 5.
///
/// Runs BEFORE [`apply_comments_phase`] so it can synthesize legacy
/// `tc={topId}` placeholder comments into `patcher.queued_comments`,
/// matching the writer-side behavior in
/// `crates/wolfxl-writer/src/emit/threaded_comments_xml.rs::synthesize_legacy_placeholders`.
///
/// Returns `(file_writes, file_deletes)` for the threaded-comments and
/// person-list parts. Mutates `patcher.rels_patches` for the affected
/// sheet rels and (if persons were queued or already present)
/// `xl/_rels/workbook.xml.rels`. Mutates `patcher.queued_content_type_ops`
/// for new `Override` entries on threaded-comments / person-list parts.
pub(super) fn apply_threaded_comments_phase(
    patcher: &mut XlsxPatcher,
    zip: &mut ZipArchive<File>,
    part_id_allocator: &mut PartIdAllocator,
) -> PyResult<(HashMap<String, Vec<u8>>, HashSet<String>)> {
    let sheet_order_local: Vec<String> = patcher.sheet_order.clone();
    let mut file_writes: HashMap<String, Vec<u8>> = HashMap::new();
    let mut file_deletes: HashSet<String> = HashSet::new();
    let mut content_type_ops: Vec<content_types::ContentTypeOp> = Vec::new();

    // ----- Phase A: per-sheet threadedCommentsN.xml + synthetic placeholders.
    for sheet_name in &sheet_order_local {
        let ops_for_sheet = match patcher.queued_threaded_comments.get(sheet_name) {
            Some(o) if !o.is_empty() => o.clone(),
            _ => continue,
        };
        let sheet_path = match patcher.sheet_paths.get(sheet_name).cloned() {
            Some(p) => p,
            None => continue,
        };
        let rels_path = sheet_rels_path_for(&sheet_path);
        if !patcher.rels_patches.contains_key(&rels_path) {
            let graph = load_or_empty_rels(zip, &rels_path)?;
            patcher.rels_patches.insert(rels_path.clone(), graph);
        }

        // Find existing threadedCommentsN.xml (if any) by walking the
        // sheet rels for the THREADED_COMMENTS rel-type. Targets are
        // stored as `../threadedComments/threadedCommentsN.xml`; resolve
        // relative to `xl/worksheets/`.
        let existing_path: Option<String> = {
            let rels = patcher
                .rels_patches
                .get(&rels_path)
                .expect("just inserted above");
            rels.find_by_type(rt::THREADED_COMMENTS)
                .first()
                .map(|r| ooxml_util::join_and_normalize(parent_dir_of(&sheet_path), &r.target))
        };

        let existing_xml: Option<Vec<u8>> = match &existing_path {
            Some(path) => {
                let exists = zip.index_for_name(path).is_some();
                if exists {
                    Some(ooxml_util::zip_read_to_string(zip, path)?.into_bytes())
                } else {
                    None
                }
            }
            None => None,
        };

        let threaded_n = match &existing_path {
            Some(path) => {
                parse_n_from_part_path(path, "xl/threadedComments/threadedComments", ".xml")
                    .unwrap_or_else(|| part_id_allocator.alloc_threaded_comments())
            }
            None => part_id_allocator.alloc_threaded_comments(),
        };

        let rels = patcher
            .rels_patches
            .get_mut(&rels_path)
            .expect("just inserted above");
        let (bytes, _rid) = threaded_comments::build_threaded_for_sheet(
            existing_xml.as_deref(),
            &ops_for_sheet,
            rels,
            threaded_n,
        );

        let part_path = existing_path
            .clone()
            .unwrap_or_else(|| format!("xl/threadedComments/threadedComments{threaded_n}.xml"));
        if bytes.is_empty() {
            if existing_path.is_some() {
                file_deletes.insert(part_path.clone());
            }
        } else {
            file_writes.insert(part_path.clone(), bytes);
            if existing_path.is_none() {
                content_type_ops.push(content_types::ContentTypeOp::AddOverride(
                    format!("/{}", part_path),
                    threaded_comments::CT_THREADED.to_string(),
                ));
            }
        }

        // Synthesize legacy `tc={topId}` placeholders into
        // `patcher.queued_comments` so the comments phase emits them.
        // This matches the writer's `synthesize_legacy_placeholders`.
        synthesize_legacy_placeholders_into_queue(patcher, sheet_name, &ops_for_sheet);
    }

    // ----- Phase B: workbook-scope personList.xml.
    let persons_queue = patcher.queued_persons.clone();
    if !persons_queue.is_empty() || any_threaded_payload_exists(&file_writes, &file_deletes) {
        let wb_rels_path = "xl/_rels/workbook.xml.rels".to_string();
        if !patcher.rels_patches.contains_key(&wb_rels_path) {
            let graph = load_or_empty_rels(zip, &wb_rels_path)?;
            patcher.rels_patches.insert(wb_rels_path.clone(), graph);
        }

        let existing_path: Option<String> = {
            let rels = patcher
                .rels_patches
                .get(&wb_rels_path)
                .expect("just inserted above");
            rels.find_by_type(rt::PERSON_LIST)
                .first()
                .map(|r| ooxml_util::join_and_normalize("xl/", &r.target))
        };

        let existing_xml: Option<Vec<u8>> = match &existing_path {
            Some(path) => {
                let exists = zip.index_for_name(path).is_some();
                if exists {
                    Some(ooxml_util::zip_read_to_string(zip, path)?.into_bytes())
                } else {
                    None
                }
            }
            None => None,
        };

        let wb_rels = patcher
            .rels_patches
            .get_mut(&wb_rels_path)
            .expect("just inserted above");
        let (bytes, _rid) = threaded_comments::build_persons_for_workbook(
            existing_xml.as_deref(),
            &persons_queue,
            wb_rels,
        );

        let part_path = existing_path
            .clone()
            .unwrap_or_else(|| "xl/persons/personList.xml".to_string());
        if bytes.is_empty() {
            if existing_path.is_some() {
                file_deletes.insert(part_path);
            }
        } else {
            file_writes.insert(part_path.clone(), bytes);
            if existing_path.is_none() {
                content_type_ops.push(content_types::ContentTypeOp::AddOverride(
                    format!("/{}", part_path),
                    threaded_comments::CT_PERSON_LIST.to_string(),
                ));
            }
        }
    }

    if !content_type_ops.is_empty() {
        patcher
            .queued_content_type_ops
            .entry("__rfc068_threaded__".to_string())
            .or_default()
            .extend(content_type_ops);
    }

    Ok((file_writes, file_deletes))
}

fn parent_dir_of(sheet_path: &str) -> &str {
    match sheet_path.rfind('/') {
        Some(idx) => &sheet_path[..=idx],
        None => "",
    }
}

fn any_threaded_payload_exists(
    file_writes: &HashMap<String, Vec<u8>>,
    file_deletes: &HashSet<String>,
) -> bool {
    file_writes
        .keys()
        .any(|p| p.starts_with("xl/threadedComments/"))
        || file_deletes
            .iter()
            .any(|p| p.starts_with("xl/threadedComments/"))
}

/// For each `Set` op, push a synthetic `CommentOp::Set` with author
/// `tc={topId}` and body `[Threaded comment]` into `patcher.queued_comments`
/// so the existing comments phase emits a legacy placeholder for Excel < 365.
/// For each `Delete`, queue a matching `CommentOp::Delete`.
///
/// Idempotent: if the user already queued a real `CommentOp::Set` at the
/// same coord (e.g. they want to override the placeholder body), we leave
/// it alone — matching the writer's `synthesize_legacy_placeholders`.
fn synthesize_legacy_placeholders_into_queue(
    patcher: &mut XlsxPatcher,
    sheet_name: &str,
    ops: &std::collections::BTreeMap<String, threaded_comments::ThreadedCommentOp>,
) {
    let entry = patcher
        .queued_comments
        .entry(sheet_name.to_string())
        .or_default();
    for (cell, op) in ops {
        match op {
            threaded_comments::ThreadedCommentOp::Set(patch) => {
                if entry.contains_key(cell) {
                    continue;
                }
                let synthetic_author = format!("tc={}", patch.top.id);
                entry.insert(
                    cell.clone(),
                    comments::CommentOp::Set(comments::CommentPatch {
                        coordinate: cell.clone(),
                        author: synthetic_author,
                        text: "[Threaded comment]".to_string(),
                        width_pt: None,
                        height_pt: None,
                    }),
                );
            }
            threaded_comments::ThreadedCommentOp::Delete => {
                if entry.contains_key(cell) {
                    continue;
                }
                entry.insert(cell.clone(), comments::CommentOp::Delete);
            }
        }
    }
}

pub(super) fn apply_comments_phase(
    patcher: &mut XlsxPatcher,
    local_blocks: &mut HashMap<String, Vec<SheetBlock>>,
    zip: &mut ZipArchive<File>,
    part_id_allocator: &mut PartIdAllocator,
) -> PyResult<(HashMap<String, Vec<u8>>, HashSet<String>)> {
    let sheet_order_local: Vec<String> = patcher.sheet_order.clone();
    let mut comment_authors = comments::CommentAuthorTable::new();
    for sheet_name in &sheet_order_local {
        let sheet_path = match patcher.sheet_paths.get(sheet_name).cloned() {
            Some(p) => p,
            None => continue,
        };
        let _ = patcher
            .ancillary
            .populate_for_sheet(zip, sheet_name, &sheet_path);
    }

    let mut file_writes: HashMap<String, Vec<u8>> = HashMap::new();
    let mut file_deletes: HashSet<String> = HashSet::new();
    let mut content_type_ops: Vec<content_types::ContentTypeOp> = Vec::new();
    let mut vml_default_added = false;

    for sheet_name in &sheet_order_local {
        let ops = match patcher.queued_comments.get(sheet_name) {
            Some(o) if !o.is_empty() => o.clone(),
            _ => continue,
        };
        let sheet_path = match patcher.sheet_paths.get(sheet_name).cloned() {
            Some(p) => p,
            None => continue,
        };
        let rels_path = sheet_rels_path_for(&sheet_path);
        patcher
            .ancillary
            .populate_for_sheet(zip, sheet_name, &sheet_path)
            .map_err(|e| {
                PyIOError::new_err(format!("ancillary populate for '{sheet_name}': {e}"))
            })?;
        let (existing_comments_path, existing_vml_path) = {
            let ancillary = patcher
                .ancillary
                .get(sheet_name)
                .cloned()
                .unwrap_or_default();
            (ancillary.comments_part, ancillary.vml_drawing_part)
        };
        if !patcher.rels_patches.contains_key(&rels_path) {
            let graph = load_or_empty_rels(zip, &rels_path)?;
            patcher.rels_patches.insert(rels_path.clone(), graph);
        }

        let existing_comments_xml: Option<Vec<u8>> = match &existing_comments_path {
            Some(path) => Some(ooxml_util::zip_read_to_string(zip, path)?.into_bytes()),
            None => None,
        };
        let existing_vml_xml: Option<Vec<u8>> = match &existing_vml_path {
            Some(path) => Some(ooxml_util::zip_read_to_string(zip, path)?.into_bytes()),
            None => None,
        };
        let sheet_xml = ooxml_util::zip_read_to_string(zip, &sheet_path)?;

        let comments_n = match &existing_comments_path {
            Some(path) => parse_n_from_part_path(path, "xl/comments", ".xml")
                .unwrap_or_else(|| part_id_allocator.alloc_comments()),
            None => part_id_allocator.alloc_comments(),
        };
        let vml_n = match &existing_vml_path {
            Some(path) => parse_n_from_part_path(path, "xl/drawings/vmlDrawing", ".vml")
                .unwrap_or_else(|| part_id_allocator.alloc_vml_drawing()),
            None => part_id_allocator.alloc_vml_drawing(),
        };

        let rels = patcher
            .rels_patches
            .get_mut(&rels_path)
            .expect("just inserted above");
        let (result, _comments_rid_opt, _vml_rid_opt) = comments::build_comments(
            existing_comments_xml.as_deref(),
            existing_vml_xml.as_deref(),
            &ops,
            sheet_xml.as_bytes(),
            rels,
            &mut comment_authors,
            comments_n,
            vml_n,
        );

        let comments_path = existing_comments_path
            .clone()
            .unwrap_or_else(|| format!("xl/comments{comments_n}.xml"));
        if result.comments_xml.is_empty() {
            if existing_comments_path.is_some() {
                file_deletes.insert(comments_path.clone());
            }
        } else {
            file_writes.insert(comments_path.clone(), result.comments_xml);
            if existing_comments_path.is_none() {
                content_type_ops.push(content_types::ContentTypeOp::AddOverride(
                    format!("/{}", comments_path),
                    comments::CT_COMMENTS.to_string(),
                ));
            }
        }

        let vml_path = existing_vml_path
            .clone()
            .unwrap_or_else(|| format!("xl/drawings/vmlDrawing{vml_n}.vml"));
        if result.vml_drawing.is_empty() {
            if existing_vml_path.is_some() {
                file_deletes.insert(vml_path.clone());
            }
        } else {
            file_writes.insert(vml_path.clone(), result.vml_drawing);
            if existing_vml_path.is_none() && !vml_default_added {
                content_type_ops.push(content_types::ContentTypeOp::EnsureDefault(
                    "vml".to_string(),
                    comments::CT_VML.to_string(),
                ));
                vml_default_added = true;
            }
        }

        let legacy_block: Vec<u8> = match &result.legacy_drawing_rid {
            Some(rid) => format!(r#"<legacyDrawing r:id="{}"/>"#, rid.0).into_bytes(),
            None => Vec::new(),
        };
        local_blocks
            .entry(sheet_path)
            .or_default()
            .push(SheetBlock::LegacyDrawing(legacy_block));
    }

    if !content_type_ops.is_empty() {
        patcher
            .queued_content_type_ops
            .entry("__rfc023_comments__".to_string())
            .or_default()
            .extend(content_type_ops);
    }

    Ok((file_writes, file_deletes))
}

pub(super) fn apply_sheet_setup_phase(
    patcher: &XlsxPatcher,
    local_blocks: &mut HashMap<String, Vec<SheetBlock>>,
) {
    let sheet_titles: Vec<String> = patcher.queued_sheet_setup.keys().cloned().collect();
    for sheet_title in &sheet_titles {
        let queued = match patcher.queued_sheet_setup.get(sheet_title) {
            Some(q) => q,
            None => continue,
        };
        let sheet_path = match patcher.sheet_paths.get(sheet_title) {
            Some(p) => p.clone(),
            None => continue,
        };
        let specs = &queued.specs;

        if let Some(s) = &specs.sheet_view {
            let bytes = wolfxl_writer::parse::sheet_setup::emit_sheet_views(s);
            if !bytes.is_empty() {
                local_blocks
                    .entry(sheet_path.clone())
                    .or_default()
                    .push(SheetBlock::SheetViews(bytes));
            }
        }
        if let Some(s) = &specs.sheet_protection {
            let bytes = wolfxl_writer::parse::sheet_setup::emit_sheet_protection(s);
            if !bytes.is_empty() {
                local_blocks
                    .entry(sheet_path.clone())
                    .or_default()
                    .push(SheetBlock::SheetProtection(bytes));
            }
        }
        if let Some(s) = &specs.page_margins {
            let bytes = wolfxl_writer::parse::sheet_setup::emit_page_margins(s);
            if !bytes.is_empty() {
                local_blocks
                    .entry(sheet_path.clone())
                    .or_default()
                    .push(SheetBlock::PageMargins(bytes));
            }
        }
        if let Some(s) = &specs.print_options {
            let bytes = wolfxl_writer::parse::sheet_setup::emit_print_options(s);
            if !bytes.is_empty() {
                local_blocks
                    .entry(sheet_path.clone())
                    .or_default()
                    .push(SheetBlock::PrintOptions(bytes));
            }
        }
        if let Some(s) = &specs.page_setup {
            let bytes = wolfxl_writer::parse::sheet_setup::emit_page_setup(s);
            if !bytes.is_empty() {
                local_blocks
                    .entry(sheet_path.clone())
                    .or_default()
                    .push(SheetBlock::PageSetup(bytes));
            }
        }
        if let Some(s) = &specs.header_footer {
            let bytes = wolfxl_writer::parse::sheet_setup::emit_header_footer(s);
            if !bytes.is_empty() {
                local_blocks
                    .entry(sheet_path.clone())
                    .or_default()
                    .push(SheetBlock::HeaderFooter(bytes));
            }
        }
        // print_titles is routed through workbook definedNames by the
        // Python coordinator. The patcher queue entry is informational.
    }
}

pub(super) fn apply_page_breaks_phase(
    patcher: &XlsxPatcher,
    local_blocks: &mut HashMap<String, Vec<SheetBlock>>,
) {
    let sheet_titles: Vec<String> = patcher.queued_page_breaks.keys().cloned().collect();
    for sheet_title in &sheet_titles {
        let queued = match patcher.queued_page_breaks.get(sheet_title) {
            Some(q) => q,
            None => continue,
        };
        let sheet_path = match patcher.sheet_paths.get(sheet_title) {
            Some(p) => p.clone(),
            None => continue,
        };

        if let Some(spec) = &queued.sheet_format {
            let bytes = wolfxl_writer::parse::page_breaks::emit_sheet_format_pr(spec);
            if !bytes.is_empty() {
                local_blocks
                    .entry(sheet_path.clone())
                    .or_default()
                    .push(SheetBlock::SheetFormatPr(bytes));
            }
        }
        if let Some(spec) = &queued.row_breaks {
            let bytes = wolfxl_writer::parse::page_breaks::emit_row_breaks(spec);
            if !bytes.is_empty() {
                local_blocks
                    .entry(sheet_path.clone())
                    .or_default()
                    .push(SheetBlock::RowBreaks(bytes));
            }
        }
        if let Some(spec) = &queued.col_breaks {
            let bytes = wolfxl_writer::parse::page_breaks::emit_col_breaks(spec);
            if !bytes.is_empty() {
                local_blocks
                    .entry(sheet_path.clone())
                    .or_default()
                    .push(SheetBlock::ColBreaks(bytes));
            }
        }
    }
}

pub(super) fn apply_autofilter_phase(
    patcher: &XlsxPatcher,
    local_blocks: &mut HashMap<String, Vec<SheetBlock>>,
    file_patches: &HashMap<String, Vec<u8>>,
    zip: &mut ZipArchive<File>,
) -> PyResult<HashMap<String, Vec<u32>>> {
    let mut autofilter_hidden_rows: HashMap<String, Vec<u32>> = HashMap::new();
    let sheet_titles: Vec<String> = patcher.queued_autofilters.keys().cloned().collect();
    for sheet_title in &sheet_titles {
        let queued = patcher
            .queued_autofilters
            .get(sheet_title)
            .cloned()
            .unwrap();
        let sheet_path = match patcher.sheet_paths.get(sheet_title) {
            Some(p) => p.clone(),
            None => continue,
        };

        let xml_bytes: Vec<u8> = if let Some(b) = file_patches.get(&sheet_path) {
            b.clone()
        } else if let Some(b) = patcher.file_adds.get(&sheet_path) {
            b.clone()
        } else {
            ooxml_util::zip_read_to_string(zip, &sheet_path)?.into_bytes()
        };

        let af_model = wolfxl_autofilter::parse::parse_autofilter(&queued.dict)
            .map_err(|e| PyErr::new::<PyValueError, _>(format!("Phase 2.5o: {e}")))?;
        let (start_row, end_row, start_col, end_col) = match af_model
            .ref_
            .as_deref()
            .and_then(autofilter_helpers::parse_a1_range)
        {
            Some(t) => t,
            None => {
                let block = wolfxl_autofilter::emit::emit(&af_model);
                if !block.is_empty() {
                    local_blocks
                        .entry(sheet_path.clone())
                        .or_default()
                        .push(SheetBlock::AutoFilter(block));
                }
                continue;
            }
        };

        let data_start = start_row + 1;
        let rows_data = autofilter_helpers::extract_cell_grid(
            &xml_bytes, data_start, end_row, start_col, end_col,
        )?;

        let drain = autofilter::drain_autofilter(&queued, &rows_data, None)
            .map_err(|e| PyErr::new::<PyValueError, _>(format!("Phase 2.5o: {e}")))?;

        let abs_hidden: Vec<u32> = drain
            .hidden_offsets
            .iter()
            .map(|off| data_start + off)
            .collect();
        if !abs_hidden.is_empty() {
            autofilter_hidden_rows.insert(sheet_path.clone(), abs_hidden);
        }

        if !drain.block_bytes.is_empty() {
            local_blocks
                .entry(sheet_path.clone())
                .or_default()
                .push(SheetBlock::AutoFilter(drain.block_bytes));
        }
    }

    Ok(autofilter_hidden_rows)
}

pub(super) fn apply_worksheet_xml_patch_phase(
    patcher: &mut XlsxPatcher,
    sheet_cell_patches: &HashMap<String, Vec<CellPatch>>,
    local_blocks: &HashMap<String, Vec<SheetBlock>>,
    autofilter_hidden_rows: &HashMap<String, Vec<u32>>,
    file_patches: &mut HashMap<String, Vec<u8>>,
    zip: &mut ZipArchive<File>,
) -> PyResult<()> {
    let mut all_sheet_paths: HashSet<String> = HashSet::new();
    all_sheet_paths.extend(sheet_cell_patches.keys().cloned());
    all_sheet_paths.extend(local_blocks.keys().cloned());
    all_sheet_paths.extend(autofilter_hidden_rows.keys().cloned());

    for sheet_path in &all_sheet_paths {
        // Compose cell edits, sibling OOXML blocks, and autoFilter row hiding
        // against the newest sheet bytes, including Phase 2.7 cloned sheets.
        let xml = if let Some(bytes) = file_patches.get(sheet_path) {
            String::from_utf8_lossy(bytes).into_owned()
        } else if let Some(bytes) = patcher.file_adds.get(sheet_path) {
            String::from_utf8_lossy(bytes).into_owned()
        } else {
            ooxml_util::zip_read_to_string(zip, sheet_path)?
        };

        let after_cells: Vec<u8> = if let Some(patches) = sheet_cell_patches.get(sheet_path) {
            sheet_patcher::patch_worksheet(&xml, patches)
                .map_err(|e| PyIOError::new_err(format!("Patch failed: {e}")))?
                .into_bytes()
        } else {
            xml.into_bytes()
        };

        let after_blocks = if let Some(blocks) = local_blocks.get(sheet_path) {
            if blocks.is_empty() {
                after_cells
            } else {
                wolfxl_merger::merge_blocks(&after_cells, blocks.clone())
                    .map_err(|e| PyIOError::new_err(format!("Merge failed: {e}")))?
            }
        } else {
            after_cells
        };

        let after_blocks = if let Some(rows) = autofilter_hidden_rows.get(sheet_path) {
            if rows.is_empty() {
                after_blocks
            } else {
                autofilter_helpers::stamp_row_hidden(&after_blocks, rows)?
            }
        } else {
            after_blocks
        };

        if patcher.file_adds.contains_key(sheet_path) {
            patcher.file_adds.insert(sheet_path.clone(), after_blocks);
        } else {
            file_patches.insert(sheet_path.clone(), after_blocks);
        }
    }

    Ok(())
}
