//! Pivot and slicer save phases for the surgical xlsx patcher.

use std::collections::HashMap;
use std::fs::File;

use pyo3::exceptions::{PyIOError, PyValueError};
use pyo3::prelude::*;
use zip::ZipArchive;

use crate::ooxml_util;

use super::{content_types, pivot, pivot_slicer, XlsxPatcher};

pub(super) fn apply_pivot_adds_phase(
    patcher: &mut XlsxPatcher,
    file_patches: &mut HashMap<String, Vec<u8>>,
    zip: &mut ZipArchive<File>,
) -> PyResult<()> {
    // Bootstrap a per-patcher pivot part-id counter from the
    // source ZIP so we never collide with existing pivot parts.
    let mut counters = pivot::PivotPartCounters::new(1, 1);
    for i in 0..zip.len() {
        if let Ok(entry) = zip.by_index(i) {
            let Ok(name) = entry.name() else {
                continue;
            };
            counters.observe(&name);
        }
    }
    // Also observe paths the patcher may have already written
    // earlier in the save (e.g. from a previous apply_pivot_adds
    // call within RFC-035 deep-clone — defensive).
    for path in patcher.file_adds.keys() {
        counters.observe(path);
    }

    // ---- Pass 1: drain caches ----
    let drained_caches: Vec<pivot::QueuedPivotCacheAdd> =
        std::mem::take(&mut patcher.queued_pivot_caches);

    // Map: queued cache_id → allocated part-id (cache_n) so we
    // can resolve rels targets for tables in Pass 2.
    let mut cache_id_to_part_id: HashMap<u32, u32> = HashMap::new();

    // Collect new <pivotCache> entries to splice into workbook.xml.
    let mut pivot_cache_refs: Vec<pivot::PivotCacheRef> = Vec::new();

    // Workbook rels graph mutation. Read once, mutate, persist.
    let workbook_rels_path = "xl/_rels/workbook.xml.rels";
    let mut workbook_rels: wolfxl_rels::RelsGraph =
        if let Some(g) = patcher.rels_patches.get(workbook_rels_path) {
            g.clone()
        } else if let Some(b) = patcher.file_adds.get(workbook_rels_path) {
            wolfxl_rels::RelsGraph::parse(b)
                .map_err(|e| PyErr::new::<PyIOError, _>(format!("workbook rels parse: {e}")))?
        } else {
            let s = ooxml_util::zip_read_to_string(zip, workbook_rels_path)?;
            wolfxl_rels::RelsGraph::parse(s.as_bytes())
                .map_err(|e| PyErr::new::<PyIOError, _>(format!("workbook rels parse: {e}")))?
        };

    let mut ct_ops: Vec<content_types::ContentTypeOp> = Vec::new();

    for cache in &drained_caches {
        let n = counters.alloc_cache();
        cache_id_to_part_id.insert(cache.cache_id, n);

        let def_path = format!("xl/pivotCache/pivotCacheDefinition{n}.xml");
        let rec_path = format!("xl/pivotCache/pivotCacheRecords{n}.xml");
        let cache_rels_path = format!("xl/pivotCache/_rels/pivotCacheDefinition{n}.xml.rels");

        patcher
            .file_adds
            .insert(def_path.clone(), cache.cache_def_xml.clone());
        patcher
            .file_adds
            .insert(rec_path.clone(), cache.cache_records_xml.clone());

        // Per-cache rels: definition → records.
        let mut cache_rels = wolfxl_rels::RelsGraph::new();
        // The cache definition uses `r:id="rId1"` to reference
        // its records part (matches the canonical emit from
        // wolfxl-pivot::emit::pivot_cache_definition_xml).
        cache_rels.add_with_id(
            wolfxl_rels::RelId("rId1".into()),
            wolfxl_pivot::rt::PIVOT_CACHE_RECORDS,
            &format!("pivotCacheRecords{n}.xml"),
            wolfxl_rels::TargetMode::Internal,
        );
        patcher.rels_patches.insert(cache_rels_path, cache_rels);

        // Workbook rel → cache definition.
        let rid = workbook_rels.add(
            wolfxl_rels::rt::PIVOT_CACHE_DEF,
            &format!("pivotCache/pivotCacheDefinition{n}.xml"),
            wolfxl_rels::TargetMode::Internal,
        );
        pivot_cache_refs.push(pivot::PivotCacheRef {
            cache_id: cache.cache_id,
            rid: rid.0,
        });

        // Content-type overrides.
        ct_ops.push(content_types::ContentTypeOp::AddOverride(
            format!("/{def_path}"),
            wolfxl_pivot::ct::PIVOT_CACHE_DEFINITION.to_string(),
        ));
        ct_ops.push(content_types::ContentTypeOp::AddOverride(
            format!("/{rec_path}"),
            wolfxl_pivot::ct::PIVOT_CACHE_RECORDS.to_string(),
        ));
    }

    // Persist workbook rels mutation.
    patcher
        .rels_patches
        .insert(workbook_rels_path.to_string(), workbook_rels);

    // Splice <pivotCaches> into xl/workbook.xml.
    if !pivot_cache_refs.is_empty() {
        let wb_xml: Vec<u8> = if let Some(b) = file_patches.get("xl/workbook.xml") {
            b.clone()
        } else if let Some(b) = patcher.file_adds.get("xl/workbook.xml") {
            b.clone()
        } else {
            ooxml_util::zip_read_to_string(zip, "xl/workbook.xml")?.into_bytes()
        };
        let updated = pivot::splice_pivot_caches(&wb_xml, &pivot_cache_refs)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("splice <pivotCaches>: {e}")))?;
        file_patches.insert("xl/workbook.xml".to_string(), updated);
    }

    // ---- Pass 2: drain tables ----
    let drained_tables: HashMap<String, Vec<pivot::QueuedPivotTableAdd>> =
        std::mem::take(&mut patcher.queued_pivot_tables);

    // Drain in sheet_order for stable output.
    let sheet_order_clone: Vec<String> = patcher.sheet_order.clone();
    for sheet_name in &sheet_order_clone {
        let queued = match drained_tables.get(sheet_name) {
            Some(q) if !q.is_empty() => q,
            _ => continue,
        };
        let sheet_path = patcher
            .sheet_paths
            .get(sheet_name)
            .cloned()
            .ok_or_else(|| {
                PyValueError::new_err(format!(
                    "queue_pivot_table_add: no such sheet: {sheet_name}"
                ))
            })?;

        let sheet_rels_path = format!(
            "{}/_rels/{}.rels",
            sheet_path.rsplit_once('/').map(|(d, _)| d).unwrap_or(""),
            sheet_path.rsplit('/').next().unwrap_or("")
        );

        let mut sheet_rels: wolfxl_rels::RelsGraph = if let Some(g) =
            patcher.rels_patches.get(&sheet_rels_path)
        {
            g.clone()
        } else if let Some(b) = patcher.file_adds.get(&sheet_rels_path) {
            wolfxl_rels::RelsGraph::parse(b)
                .map_err(|e| PyErr::new::<PyIOError, _>(format!("sheet rels parse: {e}")))?
        } else {
            match ooxml_util::zip_read_to_string_opt(zip, &sheet_rels_path)? {
                Some(s) => wolfxl_rels::RelsGraph::parse(s.as_bytes())
                    .map_err(|e| PyErr::new::<PyIOError, _>(format!("sheet rels parse: {e}")))?,
                None => wolfxl_rels::RelsGraph::new(),
            }
        };

        for table in queued {
            let table_n = counters.alloc_table();
            let table_path = format!("xl/pivotTables/pivotTable{table_n}.xml");
            let table_rels_path = format!("xl/pivotTables/_rels/pivotTable{table_n}.xml.rels");

            patcher
                .file_adds
                .insert(table_path.clone(), table.table_xml.clone());

            // Per-table rels: table → cache definition. Resolve
            // the matching cache part id via cache_id_to_part_id
            // (caches queued in this same save) OR fall back to
            // direct cache_id-to-part-id mapping for caches that
            // already exist on disk in the source workbook.
            let cache_n = match cache_id_to_part_id.get(&table.cache_id) {
                Some(&n) => n,
                None => {
                    // Fall back: cache lives on disk in source
                    // ZIP. The cacheId in workbook.xml's
                    // <pivotCaches> entry maps to a workbook-rel
                    // pointing at pivotCacheDefinition{N}.xml; we
                    // assume cache_id+1 == part_id here as a
                    // simplifying convention. Real-world: the
                    // patcher would parse <pivotCaches> to
                    // resolve. v2.0 caches always come from this
                    // save, so the fallback is rarely hit.
                    table.cache_id + 1
                }
            };

            let mut table_rels = wolfxl_rels::RelsGraph::new();
            table_rels.add_with_id(
                wolfxl_rels::RelId("rId1".into()),
                wolfxl_rels::rt::PIVOT_CACHE_DEF,
                &format!("../pivotCache/pivotCacheDefinition{cache_n}.xml"),
                wolfxl_rels::TargetMode::Internal,
            );
            patcher.rels_patches.insert(table_rels_path, table_rels);

            // Sheet rel → pivot table.
            sheet_rels.add(
                wolfxl_rels::rt::PIVOT_TABLE,
                &format!("../pivotTables/pivotTable{table_n}.xml"),
                wolfxl_rels::TargetMode::Internal,
            );

            ct_ops.push(content_types::ContentTypeOp::AddOverride(
                format!("/{table_path}"),
                wolfxl_pivot::ct::PIVOT_TABLE.to_string(),
            ));
        }

        patcher.rels_patches.insert(sheet_rels_path, sheet_rels);
    }

    // Queue content-type ops under a synthetic per-workbook key
    // (pivots are workbook-scope; not tied to a single sheet
    // name in `sheet_order`). Phase 2.5c picks these up via the
    // `synth_keys` aggregator.
    if !ct_ops.is_empty() {
        patcher
            .queued_content_type_ops
            .entry("__rfc047_pivots__".to_string())
            .or_default()
            .extend(ct_ops);
    }

    Ok(())
}

pub(super) fn apply_slicer_adds_phase(
    patcher: &mut XlsxPatcher,
    file_patches: &mut HashMap<String, Vec<u8>>,
    zip: &mut ZipArchive<File>,
) -> PyResult<()> {
    // Bootstrap counters from source ZIP + pre-existing file_adds
    // (so RFC-035 deep-clones never collide).
    let mut counters = pivot_slicer::SlicerPartCounters::new();
    for i in 0..zip.len() {
        if let Ok(entry) = zip.by_index(i) {
            let Ok(name) = entry.name() else {
                continue;
            };
            counters.observe(&name);
        }
    }
    for path in patcher.file_adds.keys() {
        counters.observe(path);
    }

    let drained: Vec<pivot_slicer::QueuedSlicer> = std::mem::take(&mut patcher.queued_slicers);

    // Group drainage results by sheet for sheet-side rels + extLst splices.
    let mut workbook_cache_refs: Vec<pivot_slicer::WorkbookSlicerCacheRef> = Vec::new();
    // sheet_title → list of slicer rids (one per slicer presentation file).
    let mut sheet_slicer_rids: HashMap<String, Vec<String>> = HashMap::new();
    let mut ct_ops: Vec<content_types::ContentTypeOp> = Vec::new();

    // Read workbook rels graph once, mutate, persist at the end.
    let workbook_rels_path = "xl/_rels/workbook.xml.rels";
    let mut workbook_rels: wolfxl_rels::RelsGraph =
        if let Some(g) = patcher.rels_patches.get(workbook_rels_path) {
            g.clone()
        } else if let Some(b) = patcher.file_adds.get(workbook_rels_path) {
            wolfxl_rels::RelsGraph::parse(b)
                .map_err(|e| PyErr::new::<PyIOError, _>(format!("workbook rels parse: {e}")))?
        } else {
            let s = ooxml_util::zip_read_to_string(zip, workbook_rels_path)?;
            wolfxl_rels::RelsGraph::parse(s.as_bytes())
                .map_err(|e| PyErr::new::<PyIOError, _>(format!("workbook rels parse: {e}")))?
        };

    // Per-sheet rels graphs cached so we can mutate-then-persist
    // each owner sheet exactly once.
    let mut sheet_rels_cache: HashMap<String, wolfxl_rels::RelsGraph> = HashMap::new();

    for q in &drained {
        let out = pivot_slicer::drain_one(q, &mut counters);

        // 1. Write cache + slicer parts to file_adds.
        patcher
            .file_adds
            .insert(out.cache_part_path.clone(), out.cache_xml.clone());
        patcher
            .file_adds
            .insert(out.slicer_part_path.clone(), out.slicer_xml.clone());

        // 2. Per-cache rels file: cache → source pivot cache.
        //    Convention used by Phase 2.5m: pivot_cache_id == part_id - 1.
        //    (queue_pivot_cache_add is monotonic from 0; phase 2.5m
        //    allocates part ids starting at 1.)
        let pivot_part_id = out.source_pivot_cache_id + 1;
        let mut cache_rels = wolfxl_rels::RelsGraph::new();
        cache_rels.add_with_id(
            wolfxl_rels::RelId("rId1".into()),
            wolfxl_rels::rt::PIVOT_CACHE_DEF,
            &format!("../pivotCache/pivotCacheDefinition{pivot_part_id}.xml"),
            wolfxl_rels::TargetMode::Internal,
        );
        patcher
            .rels_patches
            .insert(out.cache_rels_part_path.clone(), cache_rels);

        // 3. Workbook rel → cache.
        let cache_rid = workbook_rels.add(
            wolfxl_pivot::rt::SLICER_CACHE,
            &format!("slicerCaches/slicerCache{}.xml", out.cache_id),
            wolfxl_rels::TargetMode::Internal,
        );
        workbook_cache_refs.push(pivot_slicer::WorkbookSlicerCacheRef {
            name: out.cache_name.clone(),
            rid: cache_rid.0,
        });

        // 4. Sheet rels → slicer presentation.
        let sheet_path = match patcher.sheet_paths.get(&out.sheet_title) {
            Some(p) => p.clone(),
            None => {
                return Err(PyValueError::new_err(format!(
                    "queue_slicer_add: sheet not found: {}",
                    out.sheet_title
                )));
            }
        };
        let sheet_rels_path = format!(
            "{}/_rels/{}.rels",
            sheet_path.rsplit_once('/').map(|(d, _)| d).unwrap_or(""),
            sheet_path.rsplit('/').next().unwrap_or("")
        );

        let sheet_rels = sheet_rels_cache
            .entry(sheet_rels_path.clone())
            .or_insert_with(|| {
                if let Some(g) = patcher.rels_patches.get(&sheet_rels_path) {
                    g.clone()
                } else if let Some(b) = patcher.file_adds.get(&sheet_rels_path) {
                    wolfxl_rels::RelsGraph::parse(b).unwrap_or_default()
                } else {
                    match ooxml_util::zip_read_to_string_opt(zip, &sheet_rels_path) {
                        Ok(Some(s)) => {
                            wolfxl_rels::RelsGraph::parse(s.as_bytes()).unwrap_or_default()
                        }
                        _ => wolfxl_rels::RelsGraph::new(),
                    }
                }
            });
        let slicer_rid = sheet_rels.add(
            wolfxl_pivot::rt::SLICER,
            &format!("../slicers/slicer{}.xml", out.slicer_id),
            wolfxl_rels::TargetMode::Internal,
        );
        sheet_slicer_rids
            .entry(out.sheet_title.clone())
            .or_default()
            .push(slicer_rid.0);

        // 5. Content-type overrides.
        ct_ops.push(content_types::ContentTypeOp::AddOverride(
            format!("/{}", out.cache_part_path),
            wolfxl_pivot::ct::SLICER_CACHE.to_string(),
        ));
        ct_ops.push(content_types::ContentTypeOp::AddOverride(
            format!("/{}", out.slicer_part_path),
            wolfxl_pivot::ct::SLICER.to_string(),
        ));
    }

    // Persist workbook rels mutations.
    patcher
        .rels_patches
        .insert(workbook_rels_path.to_string(), workbook_rels);

    // Persist per-sheet rels mutations.
    for (path, graph) in sheet_rels_cache {
        patcher.rels_patches.insert(path, graph);
    }

    // Splice <x14:slicerCaches> into xl/workbook.xml.
    if !workbook_cache_refs.is_empty() {
        let wb_xml: Vec<u8> = if let Some(b) = file_patches.get("xl/workbook.xml") {
            b.clone()
        } else if let Some(b) = patcher.file_adds.get("xl/workbook.xml") {
            b.clone()
        } else {
            ooxml_util::zip_read_to_string(zip, "xl/workbook.xml")?.into_bytes()
        };
        let updated = pivot_slicer::splice_workbook_slicer_caches(&wb_xml, &workbook_cache_refs)
            .map_err(|e| {
                PyErr::new::<PyIOError, _>(format!("splice workbook <x14:slicerCaches>: {e}"))
            })?;
        file_patches.insert("xl/workbook.xml".to_string(), updated);
    }

    // Splice <x14:slicerList> into each owner sheet.
    for (sheet_title, rids) in &sheet_slicer_rids {
        let sheet_path = match patcher.sheet_paths.get(sheet_title) {
            Some(p) => p.clone(),
            None => continue,
        };
        let sheet_xml: Vec<u8> = if let Some(b) = file_patches.get(&sheet_path) {
            b.clone()
        } else if let Some(b) = patcher.file_adds.get(&sheet_path) {
            b.clone()
        } else {
            ooxml_util::zip_read_to_string(zip, &sheet_path)?.into_bytes()
        };
        // v2.0 emits one <x14:slicerList> per sheet referencing the
        // first slicer rel; additional slicer rels would aggregate
        // into the same list element. The slicer file itself can
        // hold multiple <slicer/> entries, but Pod 3.5 keeps it
        // 1-presentation-file-per-slicer, so we point at each rid.
        let mut updated = sheet_xml;
        for rid in rids {
            updated = pivot_slicer::splice_sheet_slicer_list(&updated, rid).map_err(|e| {
                PyErr::new::<PyIOError, _>(format!("splice sheet <x14:slicerList>: {e}"))
            })?;
        }
        file_patches.insert(sheet_path, updated);
    }

    // Queue content-type ops under a synthetic per-workbook key
    // (slicers are workbook-scope; not tied to one sheet name in
    // `sheet_order`). Phase 2.5c picks them up via `synth_keys`.
    if !ct_ops.is_empty() {
        patcher
            .queued_content_type_ops
            .entry("__rfc061_slicers__".to_string())
            .or_default()
            .extend(ct_ops);
    }

    Ok(())
}
