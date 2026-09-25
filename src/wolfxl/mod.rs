//! WolfXL — surgical xlsx patcher.
//!
//! Instead of parsing the entire workbook into a DOM (like openpyxl or umya),
//! WolfXL opens the xlsx ZIP, queues cell changes in memory, and on save:
//!   1. Patches only the worksheet XMLs that have dirty cells
//!   2. Patches sharedStrings/styles only if needed
//!   3. Raw-copies all other ZIP entries unchanged
//!
//! This makes modify-and-save O(modified data) instead of O(entire file).

#[allow(dead_code)] // RFC-013: registry is scaffolding-only; first caller is RFC-022
pub mod ancillary;
pub mod comments;
pub mod conditional_formatting;
pub mod content_types;
pub mod defined_names;
#[allow(dead_code)] // RFC-022: live caller wires up in commit 3 (queue_hyperlink + Phase 2.5e)
pub mod hyperlinks;
pub mod patcher_cells;
pub mod patcher_drawing;
pub mod patcher_external_links;
pub mod patcher_models;
pub mod patcher_payload;
pub mod patcher_pivot;
pub mod patcher_pivot_edit;
pub mod patcher_pivot_parse;
mod patcher_save;
pub mod patcher_sheet_blocks;
pub mod patcher_sheet_copy;
pub mod patcher_structural;
pub mod patcher_workbook;
pub mod properties;
#[allow(dead_code)] // SST parser used in Phase 3 (format patching reads existing styles)
pub mod shared_strings;
pub mod sheet_order;
pub mod sheet_patcher;
#[allow(dead_code)] // Styles parser/appender used in Phase 3 (format patching)
pub mod styles;
pub mod tables;
pub mod validations;
// RFC-035 Pod-β: Phase 2.7 (do_save) consumes plan_sheet_copy from this re-export.
pub mod sheet_copy;
// Sprint Θ Pod-C3: Phase 2.8 (do_save) rebuilds xl/calcChain.xml.
pub mod calcchain;
// Sprint Ν Pod-γ (RFC-047 / RFC-048): Phase 2.5m drains pivot adds.
pub mod pivot;
// Sprint Ο Pod 3.5 (RFC-061 §3.1): Phase 2.5p drains slicer caches +
// slicer presentations.
pub mod pivot_slicer;
// Sprint Ο Pod 1D (RFC-058): Phase 2.5q splices workbookProtection +
// fileSharing into xl/workbook.xml.
pub mod security;
// Sprint Ο Pod 1B (RFC-056): Phase 2.5o drains autoFilter evaluation +
// `<row hidden="1">` markers.
pub mod autofilter;
pub mod autofilter_helpers;
// Sprint Ο Pod 1A.5 (RFC-055): Phase 2.5n drains queued sheet-setup
// blocks (sheetView / sheetProtection / pageMargins / pageSetup /
// headerFooter) into per-sheet `local_blocks`.
pub mod sheet_setup;
// Sprint Π Pod Π-α (RFC-062): Phase 2.5r drains queued
// rowBreaks / colBreaks / sheetFormatPr mutations.
pub mod page_breaks;
// RFC-068 G08 step 5: modify-mode threaded comments + persons.
pub mod threaded_comments;
// RFC-071 / G18 — external link part + rels parsers.
pub mod external_links;

use std::collections::{BTreeMap, HashMap, HashSet};
use std::fs::File;

use pyo3::exceptions::{PyIOError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::PyDict;

use zip::ZipArchive;

use crate::ooxml_util;
use conditional_formatting::{CfRulePatch, ConditionalFormattingPatch};
use patcher_drawing::parse_queued_image_anchor;
use patcher_models::{
    AxisShift, QueuedChartAdd, QueuedChartRemove, QueuedImageAdd, RangeMove, SheetCopyOp,
    SheetCreateOp,
};
use patcher_payload::{
    dict_to_border_spec, dict_to_format_spec, dict_to_threaded_entry, extract_bool,
    extract_cf_rule, extract_f64, extract_str, extract_u32, parse_workbook_security_payload,
    py_runs_to_rust,
};
use patcher_save::{open_source_zip, SaveWorkspace};
use patcher_workbook::{
    load_or_empty_rels, replace_first_occurrence, sheet_rels_path_for, xml_escape_attr,
};
use sheet_patcher::{CellPatch, CellValue};
use styles::FormatSpec;
use validations::DataValidationPatch;
use wolfxl_merger::SheetBlock;
use wolfxl_rels::RelsGraph;

// ---------------------------------------------------------------------------
// PyO3 class
// ---------------------------------------------------------------------------

#[pyclass]
pub struct XlsxPatcher {
    file_path: String,
    /// Sheet name → ZIP entry path (e.g. "Sheet1" → "xl/worksheets/sheet1.xml").
    sheet_paths: HashMap<String, String>,
    /// Queued cell value changes: (sheet, "A1") → CellPatch.
    value_patches: HashMap<(String, String), CellPatch>,
    /// Queued cell format changes: (sheet, "A1") → FormatSpec.
    format_patches: HashMap<(String, String), FormatSpec>,
    /// Queued mutations to `*.rels` parts. Key: ZIP entry path (e.g.
    /// `xl/worksheets/_rels/sheet1.xml.rels`). The save loop serializes the
    /// graph and writes it in place of the original entry. Populated by
    /// future Phase-3 RFCs (RFC-022 hyperlinks, RFC-023 comments, RFC-024
    /// tables); empty in the current slice.
    rels_patches: HashMap<String, RelsGraph>,
    /// Queued sibling-block insertions on `xl/worksheets/sheet*.xml` parts.
    /// Key: sheet XML path (e.g. `xl/worksheets/sheet1.xml`). The save
    /// loop calls `wolfxl_merger::merge_blocks` after `sheet_patcher`
    /// runs, so cell-level patches and block-level patches compose
    /// without conflict. Populated by future Phase-3 RFCs (RFC-022
    /// hyperlinks, RFC-024 tables, RFC-026 conditional formatting);
    /// empty in the current slice.
    ///
    /// Note: RFC-025 (data validations) does NOT populate this map
    /// directly. It builds blocks on-demand inside `do_save` from
    /// `queued_dv_patches` so that the existing `<dataValidations>`
    /// block (read out of the source sheet XML at save time) can be
    /// merged with the queued patches before the merger is invoked.
    queued_blocks: HashMap<String, Vec<SheetBlock>>,
    /// Queued data-validation rules per sheet name (NOT path — we
    /// resolve to path inside `do_save`). Each entry becomes a single
    /// `<dataValidations>` block during save: any pre-existing block
    /// in the source sheet XML is read out, prepended verbatim, and
    /// the queued patches are appended. The combined block is then
    /// handed to `wolfxl_merger` as `SheetBlock::DataValidations`.
    queued_dv_patches: HashMap<String, Vec<DataValidationPatch>>,
    /// Queued conditional-formatting patches per sheet name (RFC-026).
    /// Each entry becomes one or more `<conditionalFormatting>` blocks
    /// during save. Existing CF blocks in the source sheet XML are
    /// extracted byte-for-byte and prepended (because the merger's
    /// replace-all CF semantics drop them otherwise — RFC-011 §5.5).
    /// `dxfId` allocation threads through a workbook-wide counter so
    /// CF added on multiple sheets in one save() session lands in a
    /// single coordinated `xl/styles.xml` mutation.
    queued_cf_patches: HashMap<String, Vec<ConditionalFormattingPatch>>,
    /// Sheet names in source-document order (RFC-013). Populated in
    /// `open()` from `xl/workbook.xml`'s `<sheet>` order. Replaces
    /// `sheet_paths.keys()` for any caller that needs deterministic
    /// iteration (RFC-020's `app.xml` regen, RFC-026's CF aggregation
    /// when it migrates off the temporary sorted-keys path).
    sheet_order: Vec<String>,
    /// Brand-new ZIP entries to emit on save (RFC-013). Parallel to
    /// `file_patches`: `file_patches` REPLACES an existing source entry
    /// in place; `file_adds` APPENDS a new entry that wasn't in the
    /// source ZIP. Collisions with source-ZIP names are a hard panic
    /// (caller bug — see RFC-013 §8 risk #2). First user is RFC-020's
    /// optional `docProps/core.xml` add path; RFC-022/023/024 will be
    /// the volume callers.
    file_adds: HashMap<String, Vec<u8>>,
    /// Source ZIP entries to skip during the save loop (RFC-013).
    /// Reserved for future use; v1 is unused. RFC-035 (copy_worksheet
    /// + delete-sheet) will be the first caller. Including the field
    /// now keeps the short-circuit predicate and rewrite loop forward-
    /// compatible without a follow-up patcher refactor.
    file_deletes: HashSet<String>,
    /// Per-sheet inventory of ancillary parts (comments, VML drawings,
    /// tables, hyperlinks) lazily populated from the source ZIP's
    /// `_rels/sheetN.xml.rels` files (RFC-013). Scaffolding-only this
    /// slice — `ancillary::AncillaryPartRegistry::populate_for_sheet`
    /// has no live caller yet. RFC-022 (Hyperlinks) is the first
    /// consumer; RFC-023/024 follow.
    #[allow(dead_code)]
    ancillary: ancillary::AncillaryPartRegistry,
    /// Per-sheet `[Content_Types].xml` ops queued by sheet block
    /// builders (RFC-013 Phase 2.5c). Each entry is the list of
    /// content-type adjustments that sheet's flush requires (a new
    /// comments/table part needs an `Override` entry; vmlDrawing
    /// requires `Default Extension="vml"`). Aggregated across sheets
    /// during `do_save` so a single workbook-wide
    /// `[Content_Types].xml` mutation absorbs every sheet's ops in
    /// one parse + serialize. Empty in this slice (RFC-022/023/024
    /// will be the first volume callers).
    queued_content_type_ops: HashMap<String, Vec<content_types::ContentTypeOp>>,
    /// Document properties pending flush (RFC-020). When `Some(_)`,
    /// `do_save` rewrites both `docProps/core.xml` and
    /// `docProps/app.xml` from the payload's fields. Routing depends
    /// on whether each part already exists in the source ZIP — present
    /// → patches it through `file_patches`; absent → adds it via
    /// RFC-013's `file_adds` primitive. Populated by
    /// [`Self::queue_properties`].
    queued_props: Option<properties::DocPropertiesPayload>,
    /// Per-sheet hyperlink ops pending flush (RFC-022). Outer key is
    /// sheet name; inner is coordinate → op. `BTreeMap` for the inner
    /// gives deterministic flush ordering when a single save touches
    /// multiple cells. Phase 2.5e drains this map: it reads the
    /// existing `<hyperlinks>` block + sheet rels, merges the queued
    /// ops, and pushes a `SheetBlock::Hyperlinks` plus mutates
    /// `rels_patches`. `None` value (delete sentinel) lands here as
    /// `HyperlinkOp::Delete` per INDEX decision #5.
    queued_hyperlinks: HashMap<String, BTreeMap<String, hyperlinks::HyperlinkOp>>,
    /// Defined-name upserts pending flush (RFC-021). Drained by
    /// Phase 2.5f, which parses `xl/workbook.xml`, merges these
    /// entries via `defined_names::merge_defined_names`, and writes
    /// the result back through `file_patches`. Empty queue → no
    /// rewrite of `xl/workbook.xml` (modify-mode no-op invariant).
    /// Order is insertion order from the Python coordinator (which
    /// itself iterates a regular dict — Python 3.7+ preserves
    /// insertion order). Within a save, the merger upserts by
    /// `(name, local_sheet_id)` so duplicates collapse to last-wins.
    queued_defined_names: Vec<defined_names::DefinedNameMut>,
    /// Per-sheet table-add patches pending flush (RFC-024). Drained
    /// by Phase 2.5g: scans the source ZIP for the workbook's
    /// existing-table inventory (across ALL sheets, since `id` and
    /// `name` are workbook-unique), allocates fresh ids + sequential
    /// part filenames, mutates `rels_patches`, queues the
    /// `[Content_Types].xml` Override entries through
    /// `queued_content_type_ops`, and pushes a
    /// `SheetBlock::TableParts` per sheet. Insertion order via Vec
    /// matches openpyxl's "first add → first slot" semantics.
    queued_tables: HashMap<String, Vec<tables::TablePatch>>,
    /// Per-sheet growth of existing tables (range, autoFilter range, and
    /// appended columns), keyed by sheet title. Drained before Phase 3.
    queued_table_growth: HashMap<String, Vec<tables::TableGrowthPatch>>,
    /// Per-sheet comment ops pending flush (RFC-023). Outer key is
    /// sheet name; inner is coordinate → op. `Set` adds/replaces a
    /// comment with the supplied text/author/width/height; `Delete`
    /// removes any existing comment at that coordinate. Drained by
    /// Phase 2.5h during `do_save`. Workbook-scope author dedup
    /// happens in `comments::CommentAuthorTable`, shared across all
    /// sheets touched in a single save.
    queued_comments: HashMap<String, BTreeMap<String, comments::CommentOp>>,
    /// Queued threaded-comment ops per sheet (RFC-068 G08 step 5). Keyed
    /// by cell coordinate. Drained by `apply_threaded_comments_phase`
    /// before the legacy comments phase, which then picks up the
    /// synthesized `tc={topId}` placeholders.
    queued_threaded_comments:
        HashMap<String, BTreeMap<String, threaded_comments::ThreadedCommentOp>>,
    /// Queued workbook-scope person additions (RFC-068 G08 step 5).
    /// Idempotent on `id` against both existing personList entries and
    /// previously queued additions.
    queued_persons: Vec<threaded_comments::PersonPatch>,
    /// Sheet-reorder operations pending flush (RFC-036). Insertion-
    /// ordered list of `(sheet_name, offset)` moves. Drained by
    /// Phase 2.5h, which sequences BEFORE Phase 2.5f (defined-names)
    /// because both phases mutate `xl/workbook.xml`. The reorder
    /// merger also produces the post-move `<definedName
    /// localSheetId>` integers, so the defined-names merger sees a
    /// workbook.xml whose tab indices already reflect the move.
    /// Empty queue → no `xl/workbook.xml` touch.
    queued_sheet_moves: Vec<(String, i32)>,
    /// Sheet-title rename operations pending flush. Queueing a rename also
    /// updates `sheet_paths` and `sheet_order` immediately so subsequent
    /// sheet-scoped mutations can use the new openpyxl-visible title.
    queued_sheet_renames: Vec<(String, String)>,
    /// Per-workbook structural-shift queue (RFC-030 / RFC-031). Each
    /// entry is `(sheet, axis, idx, n)` where `axis` is "row" or "col"
    /// and `n` is signed (positive = insert, negative = delete).
    /// Drained by Phase 2.5i during `do_save`. Order is append order
    /// — the Python coordinator validates `idx >= 1` and `amount >= 1`
    /// before queueing.
    queued_axis_shifts: Vec<AxisShift>,
    /// Per-workbook range-move queue (RFC-034). Each entry describes
    /// one paste-style relocation of a rectangular block. Drained by
    /// Phase 2.5j during `do_save`, AFTER axis shifts so a sequence
    /// like `insert_rows(2, 3)` then `move_range("C3:E10", rows=5)`
    /// is applied in source order against the post-shift coordinate
    /// space.
    queued_range_moves: Vec<RangeMove>,
    /// Per-workbook sheet-copy queue (RFC-035). Each entry is a
    /// `(src_title, dst_title)` pair in user-call order. Drained by
    /// Phase 2.7 during `do_save`, BEFORE every per-sheet phase so
    /// the cloned sheet is visible to downstream phases as if it
    /// had always been part of the source workbook.
    queued_sheet_copies: Vec<SheetCopyOp>,
    /// Per-workbook blank sheet creation queue for modify mode.
    queued_sheet_creates: Vec<SheetCreateOp>,
    /// Per-workbook source sheet deletion queue for modify mode.
    queued_sheet_deletes: Vec<String>,
    /// Deleted sheet title -> original worksheet part path. Kept separate
    /// because queue_sheet_delete removes the live sheet path eagerly.
    deleted_sheet_paths: HashMap<String, String>,
    /// Sprint Θ Pod-A: pre-seeded `file_patches` entries produced by
    /// permissive-mode load-time normalization (e.g. rewriting an
    /// empty `<sheets/>` block in `xl/workbook.xml`). Empty in the
    /// non-permissive path. Drained into the actual `file_patches`
    /// map at the start of `do_save` so every downstream phase
    /// (Phase 2.7 splice, defined-names merger, …) sees the
    /// rewritten bytes.
    permissive_seed_file_patches: HashMap<String, Vec<u8>>,
    /// Sprint Λ Pod-β (RFC-045) — per-sheet pending image adds.
    /// Drained by Phase 2.5k during `do_save`.
    queued_images: HashMap<String, Vec<QueuedImageAdd>>,
    /// Sprint S1 G06 — per-sheet pending image removals by index.
    /// Drained by Phase 2.5k during `do_save`, before image adds.
    queued_image_removes: HashMap<String, Vec<usize>>,
    /// Sprint Μ Pod-γ (RFC-046) — per-sheet pending chart adds.
    /// Each entry carries pre-serialized chart XML bytes plus an
    /// A1-style anchor cell. Drained by Phase 2.5l during
    /// `do_save`, BEFORE Phase 3 (cell patches) so a chart's data
    /// range can compose with cell rewrites in the same save.
    /// Phase 2.5l handles both the "fresh-drawing" case and the
    /// "merge-into-existing-drawing" case.
    queued_charts: HashMap<String, Vec<QueuedChartAdd>>,
    /// Sprint 2 — per-sheet pending loaded chart removals.
    /// Drained before chart additions so remove-and-replace composes.
    queued_chart_removes: HashMap<String, Vec<QueuedChartRemove>>,
    /// Sprint Ν Pod-γ (RFC-047) — pending pivot-cache adds. Append
    /// order is the cache_id allocation order. Drained by Phase
    /// 2.5m during `do_save` (sequenced AFTER Phase 2.5l so chart
    /// pivot-source linkage in v2.1 can resolve table names).
    queued_pivot_caches: Vec<pivot::QueuedPivotCacheAdd>,
    /// Sprint Ν Pod-γ (RFC-048) — pending pivot-table adds, keyed
    /// by sheet title. Drained by Phase 2.5m AFTER all caches are
    /// emitted (so the table → cache rels target is resolvable).
    queued_pivot_tables: HashMap<String, Vec<pivot::QueuedPivotTableAdd>>,
    /// G17 / RFC-070 — pending source-range edits for existing
    /// on-disk pivot caches. Drained by Phase 2.5m-edit (sequenced
    /// immediately AFTER `apply_pivot_adds_phase`). Each edit names a
    /// cache definition part and the new `<worksheetSource>` values.
    queued_pivot_source_edits: Vec<patcher_pivot_edit::QueuedPivotSourceEdit>,
    /// Drop workbook external-link parts/rels on the next modify-mode save.
    drop_external_links: bool,
    /// Sprint Ν Pod-γ — workbook-scope cache_id allocator. Bumps
    /// monotonically as `queue_pivot_cache_add` is called. The
    /// counter starts at 0 (cache_id = `pivotCache.cacheId` attr;
    /// ECMA-376 0-based). Initialised by `XlsxPatcher::open` if
    /// the source already has pivot caches.
    next_pivot_cache_id: u32,

    /// Sprint Ο Pod 1D (RFC-058) — pending workbook-level security
    /// blocks (workbookProtection + fileSharing). `None` = user
    /// never set wb.security or wb.fileSharing in this session;
    /// `Some(_)` = the queue was populated and Phase 2.5q must
    /// splice into `xl/workbook.xml`.
    queued_workbook_security: Option<wolfxl_writer::parse::workbook_security::WorkbookSecurity>,

    /// Sprint Ο Pod 1B (RFC-056) — pending autoFilter adds, keyed
    /// by sheet title. Drained by Phase 2.5o (sequenced AFTER pivot
    /// Phase 2.5m, BEFORE Phase 3 cell patches). The queue stores
    /// the §10 dict shape so the cdylib can lift it into the typed
    /// model + run filter evaluation at drain time.
    queued_autofilters: HashMap<String, autofilter::QueuedAutoFilter>,

    /// Sprint Ο Pod 1A.5 (RFC-055) — pending sheet-setup mutations,
    /// keyed by sheet title. Each entry is a parsed
    /// [`sheet_setup::QueuedSheetSetup`] holding typed specs for the
    /// 5 sheet-setup blocks. Drained by Phase 2.5n (sequenced AFTER
    /// pivots in Phase 2.5m and BEFORE Phase 2.5p slicers / Phase
    /// 2.5o autoFilter so a later sheet-protection toggle can lock
    /// the autoFilter range). Calling `queue_sheet_setup_update`
    /// again for the same sheet REPLACES the previous payload
    /// (matches Python `ws.page_setup = ...` semantics).
    queued_sheet_setup: HashMap<String, sheet_setup::QueuedSheetSetup>,

    /// Sprint Π Pod Π-α (RFC-062) — queued page-breaks +
    /// sheet-format-pr mutations, keyed by sheet title. Drained by
    /// Phase 2.5r (sequenced AFTER Phase 2.5n sheet-setup, BEFORE
    /// Phase 2.5p slicers per RFC-062 §6). Each non-empty slot
    /// (`row_breaks` / `col_breaks` / `sheet_format`) becomes one
    /// SheetBlock variant; the merger handles ECMA-376 §18.3.1.99
    /// ordering (slots 4 / 24 / 25). Calling
    /// `queue_page_breaks_update` REPLACES the previous payload,
    /// matching Python `ws.row_breaks = ...` semantics.
    queued_page_breaks: HashMap<String, page_breaks::QueuedPageBreaks>,

    /// Sprint Ο Pod 3.5 (RFC-061 §3.1) — pending slicer adds,
    /// in append order. Each entry pairs a slicer-cache + a slicer
    /// presentation against an owner sheet title. Drained by Phase
    /// 2.5p (after Phase 2.5n sheet-setup, before Phase 2.5o
    /// autoFilter).
    queued_slicers: Vec<pivot_slicer::QueuedSlicer>,
}

fn rename_hash_key<T>(map: &mut HashMap<String, T>, old_name: &str, new_name: &str) {
    if let Some(value) = map.remove(old_name) {
        map.insert(new_name.to_string(), value);
    }
}

fn rename_tuple_sheet_keys<T>(
    map: &mut HashMap<(String, String), T>,
    old_name: &str,
    new_name: &str,
) {
    let keys: Vec<(String, String)> = map
        .keys()
        .filter(|(sheet, _)| sheet == old_name)
        .cloned()
        .collect();
    for old_key in keys {
        if let Some(value) = map.remove(&old_key) {
            map.insert((new_name.to_string(), old_key.1), value);
        }
    }
}

#[pymethods]
impl XlsxPatcher {
    /// Open an xlsx file for surgical patching.
    ///
    /// When `permissive` is true and the parsed `xl/workbook.xml`
    /// declares no `<sheet>` children (e.g. a self-closing `<sheets/>`
    /// produced by a malformed but still loadable workbook), this
    /// fallback registers every worksheet target in
    /// `xl/_rels/workbook.xml.rels` under a synthesized title
    /// (`Sheet1`, `Sheet2`, ...). This unblocks the Phase 2.7
    /// self-closing-`<sheets/>` splice path and is gated behind the
    /// flag so well-formed inputs are unaffected. See Sprint Θ Pod-A.
    #[staticmethod]
    #[pyo3(signature = (path, permissive = false))]
    fn open(path: &str, permissive: bool) -> PyResult<Self> {
        let f = File::open(path)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Cannot open '{path}': {e}")))?;
        let mut zip = ZipArchive::new(f)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Not a valid ZIP: {e}")))?;

        // Parse workbook.xml + rels to build sheet name → XML path mapping.
        let wb_xml = ooxml_util::zip_read_to_string(&mut zip, "xl/workbook.xml")?;
        let rels_xml = ooxml_util::zip_read_to_string(&mut zip, "xl/_rels/workbook.xml.rels")?;
        let sheet_rids = ooxml_util::parse_workbook_sheet_rids(&wb_xml)?;
        let rel_targets = ooxml_util::parse_relationship_targets(&rels_xml)?;

        let mut sheet_paths: HashMap<String, String> = HashMap::new();
        // RFC-013: capture sheet names in source-document order. The
        // `parse_workbook_sheet_rids` call above returns a Vec in
        // document order; iterating it here preserves that ordering
        // and skips any sheet whose rId target is missing (mirroring
        // the legacy lenient-parse contract).
        let mut sheet_order: Vec<String> = Vec::with_capacity(sheet_rids.len());
        for (name, rid) in sheet_rids {
            if let Some(target) = rel_targets.get(&rid) {
                sheet_paths.insert(name.clone(), ooxml_util::join_and_normalize("xl/", target));
                sheet_order.push(name);
            }
        }

        // Sprint Θ Pod-A: permissive fallback for malformed workbooks
        // whose <sheets> block is self-closing (no <sheet> children)
        // even though the rels graph still references worksheet parts.
        // We synthesize "Sheet1", "Sheet2", ... in rels iteration order
        // for every worksheet relationship target. This makes the
        // Phase 2.7 splice exercisable through the public API.
        //
        // We also normalize `xl/workbook.xml` in-memory: the empty
        // `<sheets/>` block is rewritten to `<sheets>...</sheets>`
        // populated with `<sheet>` entries that mirror the synthesized
        // titles + the rIds we recovered from the rels graph. The
        // rewrite is queued through the standard `file_patches` map,
        // which means downstream phases (Phase 2.7 splice, defined-
        // names merger, etc.) all see a well-formed workbook.xml. This
        // does NOT mutate the source file on disk; it only affects the
        // copy emitted by `save()`.
        let mut file_patches: HashMap<String, Vec<u8>> = HashMap::new();
        if permissive && sheet_order.is_empty() {
            const WORKSHEET_REL_TYPE: &str =
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
            // Re-parse rels via wolfxl_rels so we can filter by type
            // and reuse the relationship rId on the synthesized
            // <sheet> element (Excel requires r:id to match a real
            // relationship).
            let graph = wolfxl_rels::RelsGraph::parse(rels_xml.as_bytes())
                .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to parse rels: {e}")))?;
            let mut idx: usize = 1;
            // Collected (synthesized_name, rId) pairs in rels order.
            let mut synthesized: Vec<(String, String)> = Vec::new();
            for r in graph.iter() {
                if r.rel_type == WORKSHEET_REL_TYPE {
                    let synth_name = format!("Sheet{idx}");
                    let path_in_zip = ooxml_util::join_and_normalize("xl/", &r.target);
                    sheet_paths.insert(synth_name.clone(), path_in_zip);
                    sheet_order.push(synth_name.clone());
                    synthesized.push((synth_name, r.id.0.clone()));
                    idx += 1;
                }
            }

            if !synthesized.is_empty() {
                // Build <sheet name="..." sheetId="N" r:id="..."/> entries.
                let mut entries = String::new();
                for (i, (name, rid)) in synthesized.iter().enumerate() {
                    let sheet_id = i + 1;
                    entries.push_str(&format!(
                        "<sheet name=\"{}\" sheetId=\"{}\" r:id=\"{}\"/>",
                        xml_escape_attr(name),
                        sheet_id,
                        xml_escape_attr(rid)
                    ));
                }
                let new_block = format!("<sheets>{entries}</sheets>");
                let rewritten = if let Some(replaced) =
                    replace_first_occurrence(&wb_xml, "<sheets/>", &new_block)
                {
                    replaced
                } else if let Some(replaced) =
                    replace_first_occurrence(&wb_xml, "<sheets />", &new_block)
                {
                    replaced
                } else {
                    // No empty <sheets> marker to replace — workbook
                    // already has an open/close form but contains no
                    // <sheet> children. Inject our entries before
                    // </sheets>.
                    if let Some(close_pos) = wb_xml.find("</sheets>") {
                        let mut s = String::with_capacity(wb_xml.len() + entries.len());
                        s.push_str(&wb_xml[..close_pos]);
                        s.push_str(&entries);
                        s.push_str(&wb_xml[close_pos..]);
                        s
                    } else {
                        // Workbook has no <sheets> block at all; this
                        // is too far gone for permissive mode. Fall
                        // through without rewriting workbook.xml; the
                        // splice will report MissingSourceTitle if it
                        // needs the synthesized name.
                        wb_xml.clone()
                    }
                };
                if rewritten != wb_xml {
                    file_patches.insert("xl/workbook.xml".to_string(), rewritten.into_bytes());
                }
            }
        }

        Ok(XlsxPatcher {
            file_path: path.to_string(),
            sheet_paths,
            value_patches: HashMap::new(),
            format_patches: HashMap::new(),
            rels_patches: HashMap::new(),
            queued_blocks: HashMap::new(),
            queued_dv_patches: HashMap::new(),
            queued_cf_patches: HashMap::new(),
            sheet_order,
            file_adds: HashMap::new(),
            file_deletes: HashSet::new(),
            ancillary: ancillary::AncillaryPartRegistry::new(),
            queued_content_type_ops: HashMap::new(),
            queued_props: None,
            queued_hyperlinks: HashMap::new(),
            queued_defined_names: Vec::new(),
            queued_tables: HashMap::new(),
            queued_table_growth: HashMap::new(),
            queued_comments: HashMap::new(),
            queued_threaded_comments: HashMap::new(),
            queued_persons: Vec::new(),
            queued_sheet_moves: Vec::new(),
            queued_sheet_renames: Vec::new(),
            queued_axis_shifts: Vec::new(),
            queued_range_moves: Vec::new(),
            queued_sheet_copies: Vec::new(),
            queued_sheet_creates: Vec::new(),
            queued_sheet_deletes: Vec::new(),
            deleted_sheet_paths: HashMap::new(),
            permissive_seed_file_patches: file_patches,
            queued_images: HashMap::new(),
            queued_image_removes: HashMap::new(),
            queued_charts: HashMap::new(),
            queued_chart_removes: HashMap::new(),
            queued_pivot_caches: Vec::new(),
            queued_pivot_tables: HashMap::new(),
            queued_pivot_source_edits: Vec::new(),
            drop_external_links: false,
            next_pivot_cache_id: 0,
            queued_workbook_security: None,
            queued_autofilters: HashMap::new(),
            queued_sheet_setup: HashMap::new(),
            queued_page_breaks: HashMap::new(),
            queued_slicers: Vec::new(),
        })
    }

    /// Sprint Ι Pod-α: queue a rich-text value for a cell.
    ///
    /// `runs` is a list of `(text, font_dict_or_none)` tuples.  Each
    /// font dict mirrors `wolfxl.cell.rich_text.InlineFont` field
    /// names (`b`, `i`, `strike`, `u`, `sz`, `color`, `rFont`,
    /// `family`, `charset`, `vertAlign`, `scheme`).  The patcher
    /// emits an inline-string cell (`t="inlineStr"`) — so the SST
    /// never has to be modified.
    fn queue_rich_text_value(
        &mut self,
        sheet: &str,
        cell: &str,
        runs: &Bound<'_, pyo3::types::PyList>,
    ) -> PyResult<()> {
        let parsed = py_runs_to_rust(runs)?;
        let (row, col) =
            crate::util::a1_to_row_col(cell).map_err(|e| PyErr::new::<PyValueError, _>(e))?;
        let patch = CellPatch {
            row: row + 1,
            col: col + 1,
            value: Some(CellValue::RichText(parsed)),
            style_index: None,
        };
        self.value_patches
            .insert((sheet.to_string(), cell.to_string()), patch);
        Ok(())
    }

    /// Queue an array-formula / data-table / spill-child cell.
    ///
    /// RFC-057 (Sprint Ο Pod 1C).  `payload` is a dict matching the
    /// shape pinned in §10:
    ///   - ``{"kind": "array", "ref": "A1:A10", "text": "B1:B10*2"}``
    ///   - ``{"kind": "data_table", "ref": "B2:F11", "ca": false,
    ///        "dt2D": true, "dtr": false, "r1": "A1", "r2": "A2"}``
    ///   - ``{"kind": "spill_child"}``
    fn queue_array_formula(
        &mut self,
        sheet: &str,
        cell: &str,
        payload: &Bound<'_, PyDict>,
    ) -> PyResult<()> {
        let kind: String = payload
            .get_item("kind")?
            .ok_or_else(|| PyErr::new::<PyValueError, _>("payload missing 'kind'"))?
            .extract()?;

        let value = match kind.as_str() {
            "array" => {
                let ref_range: String = payload
                    .get_item("ref")?
                    .ok_or_else(|| PyErr::new::<PyValueError, _>("array kind needs 'ref'"))?
                    .extract()?;
                let mut text: String = payload
                    .get_item("text")?
                    .ok_or_else(|| PyErr::new::<PyValueError, _>("array kind needs 'text'"))?
                    .extract()?;
                if let Some(stripped) = text.strip_prefix('=') {
                    text = stripped.to_string();
                }
                CellValue::ArrayFormula { ref_range, text }
            }
            "data_table" => {
                let ref_range: String = payload
                    .get_item("ref")?
                    .ok_or_else(|| PyErr::new::<PyValueError, _>("data_table kind needs 'ref'"))?
                    .extract()?;
                let ca: bool = payload
                    .get_item("ca")?
                    .map(|v| v.extract::<bool>())
                    .transpose()?
                    .unwrap_or(false);
                let dt2_d: bool = payload
                    .get_item("dt2D")?
                    .map(|v| v.extract::<bool>())
                    .transpose()?
                    .unwrap_or(false);
                let dtr: bool = payload
                    .get_item("dtr")?
                    .map(|v| v.extract::<bool>())
                    .transpose()?
                    .unwrap_or(false);
                let r1: Option<String> = payload.get_item("r1")?.and_then(|v| v.extract().ok());
                let r2: Option<String> = payload.get_item("r2")?.and_then(|v| v.extract().ok());
                let del1: bool = payload
                    .get_item("del1")?
                    .map(|v| v.extract::<bool>())
                    .transpose()?
                    .unwrap_or(false);
                let del2: bool = payload
                    .get_item("del2")?
                    .map(|v| v.extract::<bool>())
                    .transpose()?
                    .unwrap_or(false);
                CellValue::DataTableFormula {
                    ref_range,
                    ca,
                    dt2_d,
                    dtr,
                    r1,
                    r2,
                    del1,
                    del2,
                }
            }
            "spill_child" => CellValue::SpillChild,
            other => {
                return Err(PyErr::new::<PyValueError, _>(format!(
                    "Unknown array-formula kind: '{other}'"
                )))
            }
        };

        let (row, col) =
            crate::util::a1_to_row_col(cell).map_err(|e| PyErr::new::<PyValueError, _>(e))?;

        let patch = CellPatch {
            row: row + 1,
            col: col + 1,
            value: Some(value),
            style_index: None,
        };

        self.value_patches
            .insert((sheet.to_string(), cell.to_string()), patch);
        Ok(())
    }

    /// Queue a cell value change.
    ///
    /// `payload` is a dict matching the ExcelBench cell payload format:
    ///   {"type": "string"|"number"|"boolean"|"formula"|"blank", "value": ...}
    fn queue_value(
        &mut self,
        sheet: &str,
        cell: &str,
        payload: &Bound<'_, PyDict>,
    ) -> PyResult<()> {
        let cell_type = payload
            .get_item("type")?
            .map(|v| v.extract::<String>())
            .transpose()?
            .unwrap_or_default();

        let value = match cell_type.as_str() {
            "blank" => CellValue::Blank,
            "string" | "str" => {
                let v = payload
                    .get_item("value")?
                    .map(|v| v.extract::<String>())
                    .transpose()?
                    .unwrap_or_default();
                CellValue::String(v)
            }
            "number" | "float" | "int" | "integer" => {
                let v = payload
                    .get_item("value")?
                    .map(|v| v.extract::<f64>())
                    .transpose()?
                    .unwrap_or(0.0);
                CellValue::Number(v)
            }
            "boolean" | "bool" => {
                let v = payload
                    .get_item("value")?
                    .map(|v| v.extract::<bool>())
                    .transpose()?
                    .unwrap_or(false);
                CellValue::Boolean(v)
            }
            "formula" => {
                let v = payload
                    .get_item("value")?
                    .map(|v| v.extract::<String>())
                    .transpose()?
                    .unwrap_or_default();
                // Strip leading '=' if present (openpyxl convention)
                let formula = v.strip_prefix('=').unwrap_or(&v).to_string();
                CellValue::Formula(formula)
            }
            "error" => {
                let v = payload
                    .get_item("value")?
                    .map(|v| v.extract::<String>())
                    .transpose()?
                    .unwrap_or_default();
                CellValue::Error(v)
            }
            other => {
                return Err(PyErr::new::<PyValueError, _>(format!(
                    "Unknown cell type: '{other}'"
                )));
            }
        };

        let (row, col) =
            crate::util::a1_to_row_col(cell).map_err(|e| PyErr::new::<PyValueError, _>(e))?;

        let patch = CellPatch {
            row: row + 1, // a1_to_row_col returns 0-based, patcher uses 1-based
            col: col + 1,
            value: Some(value),
            style_index: None,
        };

        self.value_patches
            .insert((sheet.to_string(), cell.to_string()), patch);
        Ok(())
    }

    /// Queue a cell format change.
    ///
    /// `format_dict` matches the ExcelBench format dict:
    ///   {"bold": true, "font_size": 14, "font_name": "Arial", "font_color": "#FF0000",
    ///    "bg_color": "#00FF00", "number_format": "$#,##0", ...}
    fn queue_format(
        &mut self,
        sheet: &str,
        cell: &str,
        format_dict: &Bound<'_, PyDict>,
    ) -> PyResult<()> {
        let spec = dict_to_format_spec(format_dict)?;
        self.format_patches
            .insert((sheet.to_string(), cell.to_string()), spec);
        Ok(())
    }

    /// Queue a data-validation rule on a sheet (RFC-025).
    ///
    /// `payload` is a dict of openpyxl-shaped fields. `sqref` is required;
    /// every other key is optional. Booleans default to `false`. Unknown
    /// keys are ignored to keep the Python side forward-compatible.
    fn queue_data_validation(&mut self, sheet: &str, payload: &Bound<'_, PyDict>) -> PyResult<()> {
        let sqref = extract_str(payload, "sqref")?
            .ok_or_else(|| PyErr::new::<PyValueError, _>("data validation requires 'sqref'"))?;

        let patch = DataValidationPatch {
            validation_type: extract_str(payload, "validation_type")?
                .unwrap_or_else(|| "none".to_string()),
            operator: extract_str(payload, "operator")?,
            formula1: extract_str(payload, "formula1")?,
            formula2: extract_str(payload, "formula2")?,
            sqref,
            allow_blank: extract_bool(payload, "allow_blank")?.unwrap_or(false),
            show_dropdown: extract_bool(payload, "show_dropdown")?.unwrap_or(false),
            show_input_message: extract_bool(payload, "show_input_message")?.unwrap_or(false),
            show_error_message: extract_bool(payload, "show_error_message")?.unwrap_or(false),
            error_style: extract_str(payload, "error_style")?,
            error_title: extract_str(payload, "error_title")?,
            error: extract_str(payload, "error")?,
            prompt_title: extract_str(payload, "prompt_title")?,
            prompt: extract_str(payload, "prompt")?,
        };

        self.queued_dv_patches
            .entry(sheet.to_string())
            .or_default()
            .push(patch);
        Ok(())
    }

    /// Queue a conditional-formatting patch on a sheet (RFC-026).
    ///
    /// `payload` is a flat dict shaped like:
    ///   {"sqref": "A1:A10",
    ///    "rules": [
    ///      {"kind": "cellIs"|"expression"|"colorScale"|"dataBar",
    ///       "operator": "greaterThan",          # cellIs only
    ///       "formula_a": "5", "formula_b": "10", # cellIs / expression
    ///       "formula":   "...",                  # expression only
    ///       "stops": [{"cfvo_type": "min", "val": None,
    ///                  "color_rgb": "FFF8696B"}, ...],   # colorScale
    ///       "min_cfvo_type": "min", "min_val": None,     # dataBar
    ///       "max_cfvo_type": "max", "max_val": None,
    ///       "color_rgb": "FF638EC6",                     # dataBar
    ///       "stop_if_true": false,
    ///       "dxf": { ... } | None,
    ///      }, ...]}
    ///
    /// Mirrors the writer's `add_conditional_format` shape but nests rules
    /// under one wrapper per sqref so priority ordering within a wrapper
    /// is preserved.
    fn queue_conditional_formatting(
        &mut self,
        sheet: &str,
        payload: &Bound<'_, PyDict>,
    ) -> PyResult<()> {
        let sqref = extract_str(payload, "sqref")?.ok_or_else(|| {
            PyErr::new::<PyValueError, _>("conditional formatting requires 'sqref'")
        })?;

        let rules_obj = payload.get_item("rules")?.ok_or_else(|| {
            PyErr::new::<PyValueError, _>("conditional formatting requires 'rules'")
        })?;
        let rules_list = rules_obj
            .cast::<pyo3::types::PyList>()
            .map_err(|_| PyErr::new::<PyValueError, _>("'rules' must be a list of dicts"))?;

        let mut rules: Vec<CfRulePatch> = Vec::with_capacity(rules_list.len());
        for item in rules_list.iter() {
            let rd = item
                .cast::<PyDict>()
                .map_err(|_| PyErr::new::<PyValueError, _>("each rule must be a dict"))?;
            rules.push(extract_cf_rule(rd)?);
        }

        let patch = ConditionalFormattingPatch { sqref, rules };
        self.queued_cf_patches
            .entry(sheet.to_string())
            .or_default()
            .push(patch);
        Ok(())
    }

    /// Queue a cell border change.
    fn queue_border(
        &mut self,
        sheet: &str,
        cell: &str,
        border_dict: &Bound<'_, PyDict>,
    ) -> PyResult<()> {
        let border = dict_to_border_spec(border_dict)?;
        // Merge with existing format patch or create new one
        let key = (sheet.to_string(), cell.to_string());
        let spec = self.format_patches.entry(key).or_default();
        spec.border = Some(border);
        Ok(())
    }

    /// Return the list of sheet names discovered in the workbook.
    ///
    /// Returned in source-document order (the order Excel rendered the
    /// tabs). Switched from `sheet_paths.keys()` to `sheet_order` in
    /// RFC-013 so callers that thread the sheet list into output
    /// (RFC-020's `app.xml` `<TitlesOfParts>`) get the right ordering
    /// without re-parsing `xl/workbook.xml`.
    fn sheet_names(&self) -> Vec<String> {
        self.sheet_order.clone()
    }

    /// Queue a document-properties update (RFC-020). The payload is the
    /// flat dict produced by `python/wolfxl/_workbook.py`'s
    /// `_flush_properties_to_patcher`; absent fields stay `None` and
    /// don't appear in the rewritten core.xml.
    ///
    /// Recognized keys (all optional): `title`, `subject`, `creator`,
    /// `keywords`, `description`, `last_modified_by`, `category`,
    /// `content_status`, `created_iso`, `modified_iso`, `sheet_names`
    /// (`list[str]`).
    /// Queue a hyperlink set/update for `sheet[cell]` (RFC-022).
    ///
    /// `payload` keys (all optional but at least one of `target` /
    /// `location` MUST be present): `target` (external URL — http/mailto/
    /// file), `location` (internal sheet anchor like `'Sheet2'!A1`),
    /// `tooltip`, `display`. Drained by Phase 2.5e during `do_save`.
    fn queue_hyperlink(
        &mut self,
        sheet: &str,
        cell: &str,
        payload: &Bound<'_, PyDict>,
    ) -> PyResult<()> {
        let target = extract_str(payload, "target")?;
        let location = extract_str(payload, "location")?;
        let tooltip = extract_str(payload, "tooltip")?;
        let display = extract_str(payload, "display")?;
        if target.is_none() && location.is_none() {
            return Err(PyErr::new::<PyValueError, _>(
                "queue_hyperlink: at least one of 'target' or 'location' must be set",
            ));
        }
        let patch = hyperlinks::HyperlinkPatch {
            coordinate: cell.to_string(),
            target,
            location,
            tooltip,
            display,
        };
        self.queued_hyperlinks
            .entry(sheet.to_string())
            .or_default()
            .insert(cell.to_string(), hyperlinks::HyperlinkOp::Set(patch));
        Ok(())
    }

    /// Queue a defined-name upsert/delete (RFC-021).
    ///
    /// `payload` keys (`name` required; `formula` required for upserts; rest optional):
    ///   - `name`               (str)  — defined name. Includes any `_xlnm.` prefix verbatim.
    ///   - `formula`            (str)  — XML text content (no leading `=`).
    ///   - `delete`             (bool?)— remove matching name/scope when true.
    ///   - `local_sheet_id`     (int?) — `None` = workbook-scope; 0-based sheet position otherwise.
    ///   - `hidden`             (bool?)— `True` emits `hidden="1"`.
    ///   - `comment`            (str?) — defined-name `comment` attribute.
    ///   - Phase 2 (G22) — full ECMA-376 §18.2.5 attribute surface:
    ///     - `custom_menu`      (str?)
    ///     - `description`      (str?)
    ///     - `help`             (str?)
    ///     - `status_bar`       (str?)
    ///     - `shortcut_key`     (str?)
    ///     - `function`         (bool?)
    ///     - `function_group_id`(int?)
    ///     - `vb_procedure`     (bool?)
    ///     - `xlm`              (bool?)
    ///     - `publish_to_server`(bool?)
    ///     - `workbook_parameter`(bool?)
    ///
    /// Drained by Phase 2.5f during `do_save`. Upsert key is
    /// `(name, local_sheet_id)` — two entries with the same name but
    /// different scopes coexist independently.
    fn queue_defined_name(&mut self, payload: &Bound<'_, PyDict>) -> PyResult<()> {
        let name = extract_str(payload, "name")?.ok_or_else(|| {
            PyErr::new::<PyValueError, _>("queue_defined_name: 'name' is required")
        })?;
        let delete = extract_bool(payload, "delete")?.unwrap_or(false);
        let formula = match extract_str(payload, "formula")? {
            Some(formula) => formula,
            None if delete => String::new(),
            None => {
                return Err(PyErr::new::<PyValueError, _>(
                    "queue_defined_name: 'formula' is required",
                ));
            }
        };
        let local_sheet_id = match payload.get_item("local_sheet_id")? {
            Some(v) if !v.is_none() => Some(v.extract::<u32>()?),
            _ => None,
        };
        let hidden = extract_bool(payload, "hidden")?;
        let comment = extract_str(payload, "comment")?;
        let custom_menu = extract_str(payload, "custom_menu")?;
        let description = extract_str(payload, "description")?;
        let help = extract_str(payload, "help")?;
        let status_bar = extract_str(payload, "status_bar")?;
        let shortcut_key = extract_str(payload, "shortcut_key")?;
        let function = extract_bool(payload, "function")?;
        let function_group_id = extract_u32(payload, "function_group_id")?;
        let vb_procedure = extract_bool(payload, "vb_procedure")?;
        let xlm = extract_bool(payload, "xlm")?;
        let publish_to_server = extract_bool(payload, "publish_to_server")?;
        let workbook_parameter = extract_bool(payload, "workbook_parameter")?;
        self.queued_defined_names
            .push(defined_names::DefinedNameMut {
                name,
                delete,
                formula,
                local_sheet_id,
                hidden,
                comment,
                custom_menu,
                description,
                help,
                status_bar,
                shortcut_key,
                function,
                function_group_id,
                vb_procedure,
                xlm,
                publish_to_server,
                workbook_parameter,
            });
        Ok(())
    }

    /// Queue workbook-level security (RFC-058 Phase 2.5q).
    ///
    /// `payload` shape (RFC-058 §10):
    ///
    /// ```python
    /// {
    ///     "workbook_protection": {
    ///         "lock_structure": bool,
    ///         "lock_windows": bool,
    ///         "lock_revision": bool,
    ///         "workbook_algorithm_name": str | None,
    ///         "workbook_hash_value": str | None,
    ///         "workbook_salt_value": str | None,
    ///         "workbook_spin_count": int | None,
    ///         "revisions_algorithm_name": str | None,
    ///         "revisions_hash_value": str | None,
    ///         "revisions_salt_value": str | None,
    ///         "revisions_spin_count": int | None,
    ///     } | None,
    ///     "file_sharing": {
    ///         "read_only_recommended": bool,
    ///         "user_name": str | None,
    ///         "algorithm_name": str | None,
    ///         "hash_value": str | None,
    ///         "salt_value": str | None,
    ///         "spin_count": int | None,
    ///     } | None,
    /// }
    /// ```
    ///
    /// Either branch may be `None`. Drained by Phase 2.5q during
    /// `do_save`; the queue is single-slot — calling this again
    /// REPLACES the previous payload (matches the Python-side
    /// `wb.security = ...` semantics).
    fn queue_workbook_security(&mut self, payload: &Bound<'_, PyDict>) -> PyResult<()> {
        let security = parse_workbook_security_payload(payload)?;
        self.queued_workbook_security = Some(security);
        Ok(())
    }

    /// Queue a sheet-setup update for `sheet` (RFC-055 Phase 2.5n).
    ///
    /// `payload` is the §10 dict shape produced by
    /// `Worksheet.to_rust_setup_dict()`:
    ///
    /// ```text
    /// {
    ///   "page_setup": {...} | None,
    ///   "page_margins": {...} | None,
    ///   "header_footer": {...} | None,
    ///   "sheet_view": {...} | None,
    ///   "sheet_protection": {...} | None,
    ///   "print_titles": {"rows": "1:1" | None, "cols": "A:A" | None} | None,
    /// }
    /// ```
    ///
    /// Calling this again for the same `sheet` REPLACES the previous
    /// payload — matches Python `ws.page_setup = ...` semantics.
    /// Drained by Phase 2.5n during `do_save`, sequenced AFTER pivots
    /// (2.5m) and BEFORE autoFilter (2.5o).
    fn queue_sheet_setup_update(
        &mut self,
        sheet: &str,
        payload: &Bound<'_, PyDict>,
    ) -> PyResult<()> {
        let specs = sheet_setup::parse_sheet_setup_payload(payload)?;
        if specs.is_empty() {
            // No-op queue entry — drop any prior entry too, matching
            // "user reset everything to default" semantics.
            self.queued_sheet_setup.remove(sheet);
        } else {
            self.queued_sheet_setup
                .insert(sheet.to_string(), sheet_setup::QueuedSheetSetup { specs });
        }
        Ok(())
    }

    /// Queue a page-breaks + sheet-format update for `sheet`
    /// (RFC-062 Phase 2.5r).
    ///
    /// `payload` is the merged §10 dict shape produced by
    /// ``Worksheet.to_rust_page_breaks_dict()`` +
    /// ``Worksheet.to_rust_sheet_format_dict()``:
    ///
    /// ```text
    /// {
    ///   "row_breaks":   {...} | None,
    ///   "col_breaks":   {...} | None,
    ///   "sheet_format": {...} | None,
    /// }
    /// ```
    ///
    /// Calling this again for the same `sheet` REPLACES the previous
    /// payload — matches Python `ws.row_breaks = ...` semantics.
    /// Drained by Phase 2.5r during `do_save`, sequenced AFTER
    /// sheet-setup (2.5n) and BEFORE slicers (2.5p).
    fn queue_page_breaks_update(
        &mut self,
        sheet: &str,
        payload: &Bound<'_, PyDict>,
    ) -> PyResult<()> {
        let queued = page_breaks::parse_page_breaks_payload(payload)?;
        if queued.is_empty() {
            self.queued_page_breaks.remove(sheet);
        } else {
            self.queued_page_breaks.insert(sheet.to_string(), queued);
        }
        Ok(())
    }

    /// Queue a sheet reorder (RFC-036).
    ///
    /// `sheet` is the sheet's `name` attribute (resolved on the Python
    /// side from a `Worksheet` instance or string). `offset` is added
    /// to the sheet's current 0-based position; the resulting index is
    /// clamped to `[0, n-1]`. Drained by Phase 2.5h during `do_save`.
    /// Multiple queued moves apply in queue order against the running
    /// tab list, and Phase 2.5h re-points every `<definedName
    /// localSheetId>` whose integer maps to a moved position before
    /// the defined-names merger runs.
    fn queue_sheet_move(&mut self, sheet: &str, offset: i32) -> PyResult<()> {
        self.queued_sheet_moves.push((sheet.to_string(), offset));
        Ok(())
    }

    /// Queue a sheet-title rename for modify mode.
    ///
    /// The in-memory maps are updated immediately so any later queued
    /// sheet-scoped mutations can address the worksheet by its new title.
    /// `xl/workbook.xml` is rewritten during the workbook XML phase.
    fn queue_sheet_rename(&mut self, old_name: &str, new_name: &str) -> PyResult<()> {
        if old_name == new_name {
            return Ok(());
        }
        if new_name.is_empty() {
            return Err(PyValueError::new_err(
                "queue_sheet_rename: new sheet title must be non-empty",
            ));
        }
        if self.sheet_paths.contains_key(new_name) {
            return Err(PyValueError::new_err(format!(
                "queue_sheet_rename: destination sheet '{new_name}' already exists"
            )));
        }
        let Some(path) = self.sheet_paths.remove(old_name) else {
            return Err(PyValueError::new_err(format!(
                "queue_sheet_rename: source sheet '{old_name}' not found in workbook"
            )));
        };
        self.sheet_paths.insert(new_name.to_string(), path);
        for name in &mut self.sheet_order {
            if name == old_name {
                *name = new_name.to_string();
            }
        }
        for (sheet, _) in &mut self.queued_sheet_moves {
            if sheet == old_name {
                *sheet = new_name.to_string();
            }
        }
        rename_tuple_sheet_keys(&mut self.value_patches, old_name, new_name);
        rename_tuple_sheet_keys(&mut self.format_patches, old_name, new_name);
        rename_hash_key(&mut self.queued_dv_patches, old_name, new_name);
        rename_hash_key(&mut self.queued_cf_patches, old_name, new_name);
        rename_hash_key(&mut self.queued_content_type_ops, old_name, new_name);
        rename_hash_key(&mut self.queued_hyperlinks, old_name, new_name);
        rename_hash_key(&mut self.queued_tables, old_name, new_name);
        rename_hash_key(&mut self.queued_table_growth, old_name, new_name);
        rename_hash_key(&mut self.queued_comments, old_name, new_name);
        rename_hash_key(&mut self.queued_threaded_comments, old_name, new_name);
        rename_hash_key(&mut self.queued_images, old_name, new_name);
        rename_hash_key(&mut self.queued_image_removes, old_name, new_name);
        rename_hash_key(&mut self.queued_charts, old_name, new_name);
        rename_hash_key(&mut self.queued_pivot_tables, old_name, new_name);
        rename_hash_key(&mut self.queued_autofilters, old_name, new_name);
        rename_hash_key(&mut self.queued_sheet_setup, old_name, new_name);
        rename_hash_key(&mut self.queued_page_breaks, old_name, new_name);
        self.queued_sheet_renames
            .push((old_name.to_string(), new_name.to_string()));
        Ok(())
    }

    /// Queue creation of a blank worksheet in modify mode.
    #[pyo3(signature = (title, index=None))]
    fn queue_sheet_create(&mut self, title: &str, index: Option<usize>) -> PyResult<()> {
        if title.is_empty() {
            return Err(PyValueError::new_err(
                "queue_sheet_create: sheet title must be non-empty",
            ));
        }
        if self.sheet_paths.contains_key(title)
            || self.queued_sheet_creates.iter().any(|op| op.title == title)
        {
            return Err(PyValueError::new_err(format!(
                "queue_sheet_create: sheet '{title}' already exists"
            )));
        }
        self.queued_sheet_creates.push(SheetCreateOp {
            title: title.to_string(),
        });
        self.sheet_paths.insert(
            title.to_string(),
            format!("__wolfxl_pending_sheet_create__/{title}"),
        );
        let insert_at = index
            .unwrap_or(self.sheet_order.len())
            .min(self.sheet_order.len());
        self.sheet_order.insert(insert_at, title.to_string());
        Ok(())
    }

    /// Queue deletion of an existing worksheet in modify mode.
    fn queue_sheet_delete(&mut self, title: &str) -> PyResult<()> {
        let Some(path) = self.sheet_paths.remove(title) else {
            return Err(PyValueError::new_err(format!(
                "queue_sheet_delete: sheet '{title}' not found in workbook"
            )));
        };
        self.sheet_order.retain(|name| name != title);
        self.queued_dv_patches.remove(title);
        self.queued_cf_patches.remove(title);
        self.value_patches.retain(|(sheet, _), _| sheet != title);
        self.format_patches.retain(|(sheet, _), _| sheet != title);
        self.queued_hyperlinks.remove(title);
        self.queued_tables.remove(title);
        self.queued_table_growth.remove(title);
        self.queued_comments.remove(title);
        self.queued_threaded_comments.remove(title);
        self.queued_images.remove(title);
        self.queued_image_removes.remove(title);
        self.queued_charts.remove(title);
        self.queued_pivot_tables.remove(title);
        self.queued_autofilters.remove(title);
        self.queued_sheet_setup.remove(title);
        self.queued_page_breaks.remove(title);
        if path.starts_with("__wolfxl_pending_sheet_create__/") {
            self.queued_sheet_creates.retain(|op| op.title != title);
            return Ok(());
        }
        self.deleted_sheet_paths.insert(title.to_string(), path);
        self.queued_sheet_deletes.push(title.to_string());
        Ok(())
    }

    /// Queue a hyperlink delete for `sheet[cell]` (RFC-022). Idempotent:
    /// running on a cell that had no source hyperlink is a no-op at
    /// flush time.
    fn queue_hyperlink_delete(&mut self, sheet: &str, cell: &str) -> PyResult<()> {
        self.queued_hyperlinks
            .entry(sheet.to_string())
            .or_default()
            .insert(cell.to_string(), hyperlinks::HyperlinkOp::Delete);
        Ok(())
    }

    /// Queue a table addition on `sheet` (RFC-024).
    ///
    /// `payload` keys (`name`, `ref`, and `columns` are required;
    /// other keys default sensibly):
    ///   - `name`              (str)
    ///   - `display_name`      (str?, defaults to `name`)
    ///   - `ref`               (str)  — A1 range, e.g. `"A1:E10"`
    ///   - `columns`           (list[str]) — column names in order
    ///   - `style`             (dict?) — `name`, `show_first_column`,
    ///                          `show_last_column`, `show_row_stripes`,
    ///                          `show_column_stripes`
    ///   - `header_row_count`  (int?, defaults to 1)
    ///   - `totals_row_shown`  (bool?, defaults to `false`)
    ///   - `autofilter`        (bool?, defaults to `true`)
    ///
    /// Workbook-unique id allocation, name-collision detection,
    /// part-file emission, sheet-rels mutation, content-type
    /// override, and `<tableParts>` block insertion all happen at
    /// `save()` time during Phase-2.5f.
    fn queue_table(&mut self, sheet: &str, payload: &Bound<'_, PyDict>) -> PyResult<()> {
        let name = extract_str(payload, "name")?
            .ok_or_else(|| PyErr::new::<PyValueError, _>("table requires 'name'"))?;
        let display_name = extract_str(payload, "display_name")?.unwrap_or_else(|| name.clone());
        let ref_range = extract_str(payload, "ref")?
            .ok_or_else(|| PyErr::new::<PyValueError, _>("table requires 'ref'"))?;
        let columns_obj = payload
            .get_item("columns")?
            .ok_or_else(|| PyErr::new::<PyValueError, _>("table requires 'columns'"))?;
        let columns: Vec<String> = columns_obj.extract::<Vec<String>>()?;
        let header_row_count = extract_u32(payload, "header_row_count")?.unwrap_or(1);
        let totals_row_shown = extract_bool(payload, "totals_row_shown")?.unwrap_or(false);
        let autofilter = extract_bool(payload, "autofilter")?.unwrap_or(true);

        let style = match payload.get_item("style")? {
            Some(v) if !v.is_none() => {
                let d = v
                    .cast::<PyDict>()
                    .map_err(|_| PyErr::new::<PyValueError, _>("'style' must be a dict or None"))?;
                let style_name = extract_str(d, "name")?.unwrap_or_default();
                Some(tables::TableStylePatch {
                    name: style_name,
                    show_first_column: extract_bool(d, "show_first_column")?.unwrap_or(false),
                    show_last_column: extract_bool(d, "show_last_column")?.unwrap_or(false),
                    show_row_stripes: extract_bool(d, "show_row_stripes")?.unwrap_or(false),
                    show_column_stripes: extract_bool(d, "show_column_stripes")?.unwrap_or(false),
                })
            }
            _ => None,
        };

        let patch = tables::TablePatch {
            name,
            display_name,
            ref_range,
            columns,
            style,
            header_row_count,
            totals_row_shown,
            autofilter,
        };
        self.queued_tables
            .entry(sheet.to_string())
            .or_default()
            .push(patch);
        Ok(())
    }

    /// Queue growth of an existing table on `sheet` (modify mode).
    ///
    /// `ref_range` is the new table range; its top-left cell and the
    /// existing columns stay fixed. `auto_filter_ref` replaces the table's
    /// own autoFilter range when the table has one. `appended_columns`
    /// become new `<tableColumn>` entries after the existing ones.
    fn queue_table_growth(
        &mut self,
        sheet: &str,
        table_name: &str,
        ref_range: &str,
        auto_filter_ref: &str,
        appended_columns: Vec<String>,
    ) -> PyResult<()> {
        if table_name.is_empty() || ref_range.is_empty() {
            return Err(PyValueError::new_err(
                "table growth requires a table name and a range",
            ));
        }
        self.queued_table_growth
            .entry(sheet.to_string())
            .or_default()
            .push(tables::TableGrowthPatch {
                table_name: table_name.to_string(),
                ref_range: ref_range.to_string(),
                auto_filter_ref: auto_filter_ref.to_string(),
                appended_columns,
            });
        Ok(())
    }

    /// Sprint Λ Pod-β (RFC-045) — queue an image add for `sheet`.
    ///
    /// Payload shape mirrors `NativeWorkbook.add_image`:
    /// ```python
    /// {
    ///   "data": <bytes>,
    ///   "ext": "png" | "jpeg" | "gif" | "bmp",
    ///   "width": int,
    ///   "height": int,
    ///   "anchor": {"type": "one_cell"|"two_cell"|"absolute", ...},
    /// }
    /// ```
    /// Drained by Phase 2.5k during `do_save`. Sheets that already
    /// have a drawing rel raise `NotImplementedError` at flush time
    /// (v1.5 follow-up: append to existing drawingN.xml).
    fn queue_image_add(&mut self, sheet: &str, payload: &Bound<'_, PyDict>) -> PyResult<()> {
        let data: Vec<u8> = payload
            .get_item("data")?
            .ok_or_else(|| PyValueError::new_err("queue_image_add: missing 'data'"))?
            .extract()?;
        let ext: String = payload
            .get_item("ext")?
            .ok_or_else(|| PyValueError::new_err("queue_image_add: missing 'ext'"))?
            .extract()?;
        let width: u32 = payload
            .get_item("width")?
            .ok_or_else(|| PyValueError::new_err("queue_image_add: missing 'width'"))?
            .extract()?;
        let height: u32 = payload
            .get_item("height")?
            .ok_or_else(|| PyValueError::new_err("queue_image_add: missing 'height'"))?
            .extract()?;
        let anchor_obj = payload
            .get_item("anchor")?
            .ok_or_else(|| PyValueError::new_err("queue_image_add: missing 'anchor'"))?;
        let anchor_dict = anchor_obj
            .cast::<PyDict>()
            .map_err(|_| PyValueError::new_err("queue_image_add: 'anchor' must be a dict"))?;
        let anchor = parse_queued_image_anchor(anchor_dict)?;
        self.queued_images
            .entry(sheet.to_string())
            .or_default()
            .push(QueuedImageAdd {
                data,
                ext: ext.to_ascii_lowercase(),
                width_px: width,
                height_px: height,
                anchor,
            });
        Ok(())
    }

    /// Sprint S1 G06 — queue an image removal for `sheet`.
    ///
    /// `index` is a zero-based image index in the source drawing state,
    /// evaluated in queue order.
    fn queue_image_remove(&mut self, sheet: &str, index: usize) -> PyResult<()> {
        self.queued_image_removes
            .entry(sheet.to_string())
            .or_default()
            .push(index);
        Ok(())
    }

    /// Sprint Μ Pod-γ (RFC-046) — queue a chart add for `sheet`.
    ///
    /// The caller (Pod-α's `emit_chart_xml(&Chart)` from the writer
    /// crate, surfaced via `Workbook.add_chart_modify_mode`) supplies
    /// pre-serialized chart XML bytes plus an A1-style anchor. The
    /// patcher routes the bytes through `xl/charts/chartN.xml`, the
    /// drawing layer (fresh OR existing), the rels graphs, and the
    /// content-types map. Drained by Phase 2.5l during `do_save`.
    fn queue_chart_add(
        &mut self,
        sheet: &str,
        chart_xml: Vec<u8>,
        anchor_a1: &str,
        width_emu: i64,
        height_emu: i64,
    ) -> PyResult<()> {
        if anchor_a1.trim().is_empty() {
            return Err(PyValueError::new_err(
                "queue_chart_add: anchor_a1 must be a non-empty A1 cell coord",
            ));
        }
        self.queued_charts
            .entry(sheet.to_string())
            .or_default()
            .push(QueuedChartAdd {
                chart_xml,
                anchor_a1: anchor_a1.to_string(),
                width_emu,
                height_emu,
            });
        Ok(())
    }

    fn queue_chart_remove(
        &mut self,
        sheet: &str,
        drawing_path: &str,
        chart_rid: &str,
        chart_path: &str,
    ) -> PyResult<()> {
        if drawing_path.trim().is_empty()
            || chart_rid.trim().is_empty()
            || chart_path.trim().is_empty()
        {
            return Err(PyValueError::new_err(
                "queue_chart_remove: drawing_path, chart_rid, and chart_path are required",
            ));
        }
        self.queued_chart_removes
            .entry(sheet.to_string())
            .or_default()
            .push(QueuedChartRemove {
                drawing_path: drawing_path.to_string(),
                chart_rid: chart_rid.to_string(),
                chart_path: chart_path.to_string(),
            });
        Ok(())
    }

    /// Sprint Ν Pod-γ (RFC-047) — queue a pivot cache add. Returns
    /// the allocated 0-based `cache_id` so the caller can wire it
    /// into pivot tables that reference this cache.
    ///
    /// The XML payloads are pre-serialised by the Python coordinator
    /// via `wolfxl._rust.serialize_pivot_cache_dict` (definition)
    /// and `serialize_pivot_records_dict` (records). Drained by
    /// Phase 2.5m during `do_save`.
    fn queue_pivot_cache_add(
        &mut self,
        cache_def_xml: Vec<u8>,
        cache_records_xml: Vec<u8>,
    ) -> PyResult<u32> {
        let cache_id = self.next_pivot_cache_id;
        self.next_pivot_cache_id += 1;
        self.queued_pivot_caches.push(pivot::QueuedPivotCacheAdd {
            cache_def_xml,
            cache_records_xml,
            cache_id,
        });
        Ok(cache_id)
    }

    /// Sprint Ν Pod-γ (RFC-048) — queue a pivot table add. The
    /// `cache_id` must reference a cache previously queued via
    /// `queue_pivot_cache_add` (or already present in the source
    /// workbook). Drained by Phase 2.5m AFTER the cache drain so the
    /// table → cache rels target resolves cleanly.
    fn queue_pivot_table_add(
        &mut self,
        sheet: &str,
        table_xml: Vec<u8>,
        cache_id: u32,
    ) -> PyResult<()> {
        if !self.sheet_paths.contains_key(sheet) {
            return Err(PyValueError::new_err(format!(
                "queue_pivot_table_add: no such sheet: {sheet}"
            )));
        }
        self.queued_pivot_tables
            .entry(sheet.to_string())
            .or_default()
            .push(pivot::QueuedPivotTableAdd {
                sheet: sheet.to_string(),
                table_xml,
                cache_id,
            });
        Ok(())
    }

    /// G17 / RFC-070 — register a source-range mutation against an
    /// existing on-disk pivot cache definition. Drained by Phase
    /// 2.5m-edit (immediately AFTER `apply_pivot_adds_phase`).
    ///
    /// `cache_part_path` is the ZIP entry path of the cache
    /// definition (e.g. `xl/pivotCache/pivotCacheDefinition1.xml`).
    /// `new_ref` is the A1 range string. `new_sheet` is optional; pass
    /// `None` to leave the existing `sheet=` attribute untouched (or
    /// absent). `force_refresh_on_load` flips the
    /// `<pivotCacheDefinition refreshOnLoad="1">` flag — caller is
    /// responsible for setting it when the new range's column count
    /// differs from the original.
    fn register_pivot_source_edit(
        &mut self,
        cache_part_path: String,
        new_ref: String,
        new_sheet: Option<String>,
        force_refresh_on_load: bool,
    ) -> PyResult<()> {
        self.queued_pivot_source_edits
            .push(patcher_pivot_edit::QueuedPivotSourceEdit {
                cache_part_path,
                new_ref,
                new_sheet,
                force_refresh_on_load,
            });
        Ok(())
    }

    fn queue_external_links_drop(&mut self) -> PyResult<()> {
        self.drop_external_links = true;
        Ok(())
    }

    /// Sprint Ο Pod 1B (RFC-056) — queue an autoFilter for a sheet.
    ///
    /// `dict` is the §10 dict shape produced by
    /// `Worksheet.auto_filter.to_rust_dict()`. Drained by Phase 2.5o
    /// during `do_save` (sequenced AFTER pivots, BEFORE cells).
    fn queue_autofilter(&mut self, sheet: &str, dict: &Bound<'_, PyDict>) -> PyResult<()> {
        if !self.sheet_paths.contains_key(sheet) {
            return Err(PyValueError::new_err(format!(
                "queue_autofilter: no such sheet: {sheet}"
            )));
        }
        let dv = autofilter::pyany_to_dictvalue(&dict.as_any().clone())?;
        self.queued_autofilters.insert(
            sheet.to_string(),
            autofilter::QueuedAutoFilter {
                sheet: sheet.to_string(),
                dict: dv,
            },
        );
        Ok(())
    }

    /// Sprint Ο Pod 3.5 (RFC-061 §3.1) — queue one slicer + its
    /// cache for emit on a sheet. Each call adds a single
    /// `(slicerCache, slicer)` pair; v2.0 emits one slicer per
    /// presentation file.
    ///
    /// `cache_dict` follows the §10.1 contract;
    /// `slicer_dict` follows §10.2. The Python coordinator's
    /// `Workbook._flush_pending_slicers_to_patcher` builds these.
    ///
    /// Drained by Phase 2.5p in `do_save`.
    fn queue_slicer_add(
        &mut self,
        sheet: &str,
        cache_dict: &Bound<'_, PyDict>,
        slicer_dict: &Bound<'_, PyDict>,
    ) -> PyResult<()> {
        if !self.sheet_paths.contains_key(sheet) {
            return Err(PyValueError::new_err(format!(
                "queue_slicer_add: no such sheet: {sheet}"
            )));
        }
        let queued = pivot_slicer::parse_queued_slicer(sheet, cache_dict, slicer_dict)?;
        self.queued_slicers.push(queued);
        Ok(())
    }

    /// Queue a comment set/update for `sheet[cell]` (RFC-023).
    ///
    /// `payload` keys: `text` (required), `author` (optional — defaults
    /// to `"wolfxl"` to match the writer), `width_pt` / `height_pt`
    /// (optional, in OOXML points). Drained by Phase 2.5g during
    /// `do_save`.
    fn queue_comment(
        &mut self,
        sheet: &str,
        cell: &str,
        payload: &Bound<'_, PyDict>,
    ) -> PyResult<()> {
        let text = extract_str(payload, "text")?.unwrap_or_default();
        let author = extract_str(payload, "author")?.unwrap_or_else(|| "wolfxl".to_string());
        let width_pt = extract_f64(payload, "width_pt")?;
        let height_pt = extract_f64(payload, "height_pt")?;
        let patch = comments::CommentPatch {
            coordinate: cell.to_string(),
            author,
            text,
            width_pt,
            height_pt,
        };
        self.queued_comments
            .entry(sheet.to_string())
            .or_default()
            .insert(cell.to_string(), comments::CommentOp::Set(patch));
        Ok(())
    }

    /// Queue a comment delete for `sheet[cell]` (RFC-023). Idempotent:
    /// running on a cell that had no source comment is a no-op at
    /// flush time.
    fn queue_comment_delete(&mut self, sheet: &str, cell: &str) -> PyResult<()> {
        self.queued_comments
            .entry(sheet.to_string())
            .or_default()
            .insert(cell.to_string(), comments::CommentOp::Delete);
        Ok(())
    }

    /// Queue a threaded-comment set for `sheet[cell]` (RFC-068 G08
    /// step 5). The payload carries the top-level entry plus a flat
    /// list of replies; both are pre-resolved by the Python flush
    /// layer (GUIDs allocated, person ids resolved, ISO timestamps
    /// formatted).
    ///
    /// Payload schema (top-level):
    /// ```text
    /// {
    ///   "top": {id, cell, person_id, created, text, done},
    ///   "replies": [{id, cell, person_id, created, parent_id, text, done}, ...]
    /// }
    /// ```
    ///
    /// Set replaces every existing thread at this coordinate. To delete
    /// all threads on a cell use `queue_threaded_comment_delete`.
    fn queue_threaded_comment(
        &mut self,
        sheet: &str,
        cell: &str,
        payload: &Bound<'_, PyDict>,
    ) -> PyResult<()> {
        let top_dict = payload
            .get_item("top")?
            .ok_or_else(|| PyValueError::new_err("queue_threaded_comment: missing 'top'"))?;
        let top_bound = top_dict.cast::<PyDict>()?;
        let top = dict_to_threaded_entry(top_bound)?;

        let mut replies: Vec<threaded_comments::ThreadedCommentEntry> = Vec::new();
        if let Some(replies_obj) = payload.get_item("replies")? {
            if let Ok(list) = replies_obj.cast::<pyo3::types::PyList>() {
                for item in list.iter() {
                    let dict = item.cast::<PyDict>()?;
                    replies.push(dict_to_threaded_entry(dict)?);
                }
            }
        }

        let op =
            threaded_comments::ThreadedCommentOp::Set(threaded_comments::ThreadedCommentPatch {
                top,
                replies,
            });
        self.queued_threaded_comments
            .entry(sheet.to_string())
            .or_default()
            .insert(cell.to_string(), op);
        Ok(())
    }

    /// Queue a threaded-comment delete for `sheet[cell]` (RFC-068 G08
    /// step 5). Drops the top-level thread and every reply on that
    /// cell. Idempotent on cells with no existing threads.
    fn queue_threaded_comment_delete(&mut self, sheet: &str, cell: &str) -> PyResult<()> {
        self.queued_threaded_comments
            .entry(sheet.to_string())
            .or_default()
            .insert(
                cell.to_string(),
                threaded_comments::ThreadedCommentOp::Delete,
            );
        Ok(())
    }

    /// Queue a workbook-scope person addition for the personList
    /// (RFC-068 G08 step 5). Idempotent on `id`: the patcher merges
    /// against existing persons + the queue and skips duplicates.
    ///
    /// Payload: `{id, name, user_id?, provider_id?}`. The `id` GUID is
    /// allocated by the Python `PersonRegistry`; `provider_id` defaults
    /// to `"None"` when omitted to match Excel.
    fn queue_person(&mut self, payload: &Bound<'_, PyDict>) -> PyResult<()> {
        let id = extract_str(payload, "id")?
            .ok_or_else(|| PyValueError::new_err("queue_person: missing 'id'"))?;
        if id.is_empty() {
            return Err(PyValueError::new_err(
                "queue_person: 'id' must be non-empty",
            ));
        }
        let display_name = extract_str(payload, "name")?.unwrap_or_default();
        let user_id = extract_str(payload, "user_id")?.unwrap_or_default();
        let provider_id =
            extract_str(payload, "provider_id")?.unwrap_or_else(|| "None".to_string());
        self.queued_persons.push(threaded_comments::PersonPatch {
            id,
            display_name,
            user_id,
            provider_id,
        });
        Ok(())
    }

    /// Queue a structural axis shift for `sheet` (RFC-030 / RFC-031).
    ///
    /// `axis` must be `"row"` or `"col"`. `idx` is 1-based; `n` is
    /// signed (positive = insert; negative = delete). The Python
    /// coordinator validates `idx >= 1` and `amount >= 1` before
    /// queueing so this method does NOT re-validate.
    ///
    /// Drained by Phase 2.5i during `do_save`. Order is append order
    /// — multi-op sequencing matters (each op runs in the coordinate
    /// space produced by the previous op).
    fn queue_axis_shift(&mut self, sheet: &str, axis: &str, idx: u32, n: i32) -> PyResult<()> {
        if axis != "row" && axis != "col" {
            return Err(PyErr::new::<PyValueError, _>(format!(
                "queue_axis_shift: axis must be 'row' or 'col', got '{axis}'"
            )));
        }
        if idx < 1 {
            return Err(PyErr::new::<PyValueError, _>(
                "queue_axis_shift: idx must be >= 1",
            ));
        }
        self.queued_axis_shifts.push(AxisShift {
            sheet: sheet.to_string(),
            axis: axis.to_string(),
            idx,
            n,
        });
        Ok(())
    }

    /// Queue a paste-style range relocation for `sheet` (RFC-034).
    ///
    /// The source rectangle is given by 1-based inclusive corners
    /// `(src_min_col, src_min_row)..=(src_max_col, src_max_row)` and
    /// the destination delta by `(d_row, d_col)`. `translate` controls
    /// whether external formulas pointing INTO `src` also re-anchor
    /// (default false — matches openpyxl). The Python coordinator
    /// validates corners and destination bounds before queueing.
    ///
    /// Drained by Phase 2.5j during `do_save`. Ops apply in append
    /// order; each op runs against the post-previous-op bytes.
    #[allow(clippy::too_many_arguments)]
    fn queue_range_move(
        &mut self,
        sheet: &str,
        src_min_col: u32,
        src_min_row: u32,
        src_max_col: u32,
        src_max_row: u32,
        d_row: i32,
        d_col: i32,
        translate: bool,
    ) -> PyResult<()> {
        if src_min_col < 1 || src_min_row < 1 {
            return Err(PyErr::new::<PyValueError, _>(
                "queue_range_move: source corners must be >= 1",
            ));
        }
        if src_min_col > src_max_col || src_min_row > src_max_row {
            return Err(PyErr::new::<PyValueError, _>(
                "queue_range_move: src_min must be <= src_max on both axes",
            ));
        }
        self.queued_range_moves.push(RangeMove {
            sheet: sheet.to_string(),
            src_min_col,
            src_min_row,
            src_max_col,
            src_max_row,
            d_row,
            d_col,
            translate,
        });
        Ok(())
    }

    /// Queue a sheet-copy op (RFC-035 Phase 7.3).
    ///
    /// Validates eagerly that `src_title` exists in `self.sheet_paths`,
    /// `dst_title` is non-empty, `dst_title` is not already a sheet
    /// name in `self.sheet_paths`, and `dst_title` is not already
    /// queued by an earlier `queue_sheet_copy` call. On success
    /// appends to `queued_sheet_copies` (drained by Phase 2.7 in
    /// append order during `do_save`).
    ///
    /// `deep_copy_images` is the workbook-level
    /// `wb.copy_options.deep_copy_images` flag, snapshot at queue
    /// time. Defaults to `false` to preserve historical RFC-035 §5.3
    /// alias-by-target behaviour for callers that omit it.
    #[pyo3(signature = (src_title, dst_title, deep_copy_images=false))]
    fn queue_sheet_copy(
        &mut self,
        src_title: &str,
        dst_title: &str,
        deep_copy_images: bool,
    ) -> PyResult<()> {
        if !self.sheet_paths.contains_key(src_title) {
            return Err(PyErr::new::<PyValueError, _>(format!(
                "queue_sheet_copy: source sheet '{src_title}' not found in workbook"
            )));
        }
        if dst_title.is_empty() {
            return Err(PyErr::new::<PyValueError, _>(
                "queue_sheet_copy: destination title must be non-empty",
            ));
        }
        if self.sheet_paths.contains_key(dst_title) {
            return Err(PyErr::new::<PyValueError, _>(format!(
                "queue_sheet_copy: destination sheet '{dst_title}' already exists"
            )));
        }
        if self
            .queued_sheet_copies
            .iter()
            .any(|op| op.dst_title == dst_title)
        {
            return Err(PyErr::new::<PyValueError, _>(format!(
                "queue_sheet_copy: destination sheet '{dst_title}' is already queued"
            )));
        }
        self.queued_sheet_copies.push(SheetCopyOp {
            src_title: src_title.to_string(),
            dst_title: dst_title.to_string(),
            deep_copy_images,
        });
        Ok(())
    }

    fn queue_properties(&mut self, payload: &Bound<'_, PyDict>) -> PyResult<()> {
        let title = extract_str(payload, "title")?;
        let subject = extract_str(payload, "subject")?;
        let creator = extract_str(payload, "creator")?;
        let keywords = extract_str(payload, "keywords")?;
        let description = extract_str(payload, "description")?;
        let last_modified_by = extract_str(payload, "last_modified_by")?;
        let category = extract_str(payload, "category")?;
        let content_status = extract_str(payload, "content_status")?;
        let created_iso = extract_str(payload, "created_iso")?;
        let modified_iso = extract_str(payload, "modified_iso")?;
        let sheet_names: Vec<String> = match payload.get_item("sheet_names")? {
            Some(v) => v.extract::<Vec<String>>()?,
            None => Vec::new(),
        };
        self.queued_props = Some(properties::DocPropertiesPayload {
            title,
            subject,
            creator,
            keywords,
            description,
            last_modified_by,
            category,
            content_status,
            created_iso,
            modified_iso,
            sheet_names,
        });
        Ok(())
    }

    /// Save patched file to a new path.
    fn save(&mut self, path: &str) -> PyResult<()> {
        self.do_save(path)
    }

    /// Save in-place (atomic tmp+rename).
    fn save_in_place(&mut self) -> PyResult<()> {
        let target = self.file_path.clone();
        self.do_save(&target)
    }

    /// Return whether the Rust patcher already has queued mutations.
    fn _has_pending_save_work(&self) -> bool {
        self.has_pending_save_work()
    }

    /// RFC-072 (G19): return the raw `xl/vbaProject.bin` bytes from the
    /// source workbook, or `None` when the workbook contains no VBA
    /// archive. Read-only inspection — no authoring side effects.
    ///
    /// The patcher does not eagerly buffer the payload at `open()` time;
    /// instead it reopens the source ZIP on demand and reads the
    /// `xl/vbaProject.bin` entry if present. xlsx files (no VBA) yield
    /// `None`. This stays true to the modify-mode "preserve untouched
    /// parts via raw-copy" model — `xl/vbaProject.bin` round-trips
    /// untouched through `do_save`, and this accessor merely surfaces
    /// the same bytes for inspection.
    fn get_vba_archive_bytes<'py>(
        &self,
        py: Python<'py>,
    ) -> PyResult<Option<pyo3::Bound<'py, pyo3::types::PyBytes>>> {
        let mut zip = open_source_zip(&self.file_path)?;
        let buf: Option<Vec<u8>> = match zip.by_name("xl/vbaProject.bin") {
            Ok(mut f) => {
                let mut buf = Vec::with_capacity(f.size() as usize);
                std::io::Read::read_to_end(&mut f, &mut buf).map_err(|e| {
                    PyErr::new::<PyIOError, _>(format!("Failed to read xl/vbaProject.bin: {e}"))
                })?;
                Some(buf)
            }
            Err(zip::result::ZipError::FileNotFound) => None,
            Err(e) => {
                return Err(PyErr::new::<PyIOError, _>(format!(
                    "Zip error reading xl/vbaProject.bin: {e}"
                )));
            }
        };
        Ok(buf.map(|b| pyo3::types::PyBytes::new(py, &b)))
    }

    // -------------------------------------------------------------------
    // RFC-013 test-only hooks.
    //
    // These methods drive the new patcher primitives (`file_adds`,
    // `queued_content_type_ops`, `ancillary`) directly so pytest
    // integration tests can verify behavior end-to-end. They are
    // intentionally `_test_`-prefixed (Python convention for "internal
    // testing API") and have NO live caller in `python/wolfxl/`. RFC-022
    // / RFC-023 / RFC-024 will add the real public callers; until then,
    // these hooks are how `tests/test_patcher_infra.py` exercises the
    // plumbing.
    // -------------------------------------------------------------------

    /// Inject a brand-new ZIP entry that will be emitted on the next
    /// `save()`. Used by `tests/test_patcher_infra.py` to verify that
    /// `file_adds` round-trips through `do_save`.
    fn _test_inject_file_add(&mut self, path: &str, bytes: Vec<u8>) {
        self.file_adds.insert(path.to_string(), bytes);
    }

    /// Queue a content-type op against a sheet. `kind` is `"add_override"`
    /// or `"ensure_default"`; `key` is the part path or extension; `value`
    /// is the content type. The next `save()` aggregates queued ops
    /// across all sheets in `sheet_order` and writes one rewritten
    /// `[Content_Types].xml`.
    fn _test_queue_content_type_op(
        &mut self,
        sheet: &str,
        kind: &str,
        key: &str,
        value: &str,
    ) -> PyResult<()> {
        let op = match kind {
            "add_override" => {
                content_types::ContentTypeOp::AddOverride(key.to_string(), value.to_string())
            }
            "ensure_default" => {
                content_types::ContentTypeOp::EnsureDefault(key.to_string(), value.to_string())
            }
            other => {
                return Err(PyErr::new::<PyValueError, _>(format!(
                    "unknown ContentTypeOp kind '{other}' (expected 'add_override' or 'ensure_default')"
                )));
            }
        };
        self.queued_content_type_ops
            .entry(sheet.to_string())
            .or_default()
            .push(op);
        Ok(())
    }

    /// Lazily populate the ancillary registry for one sheet by name. After
    /// this call, `_test_ancillary_*` accessors return the classified
    /// `_rels/sheetN.xml.rels` contents.
    fn _test_populate_ancillary(&mut self, sheet: &str) -> PyResult<()> {
        let path = self
            .sheet_paths
            .get(sheet)
            .cloned()
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("no such sheet: {sheet}")))?;
        let mut zip = open_source_zip(&self.file_path)?;
        self.ancillary
            .populate_for_sheet(&mut zip, sheet, &path)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("ancillary populate: {e}")))?;
        Ok(())
    }

    /// Has the ancillary registry been populated for `sheet`? Returns
    /// `False` for unknown sheets and for sheets whose rels file has not
    /// been read yet.
    fn _test_ancillary_is_populated(&self, sheet: &str) -> bool {
        self.ancillary.get(sheet).is_some()
    }

    /// Cached comments-part path for `sheet`, or `None` if the sheet has
    /// none / has not been populated.
    fn _test_ancillary_comments_part(&self, sheet: &str) -> Option<String> {
        self.ancillary
            .get(sheet)
            .and_then(|a| a.comments_part.clone())
    }

    /// Cached VML drawing part path for `sheet`.
    fn _test_ancillary_vml_drawing_part(&self, sheet: &str) -> Option<String> {
        self.ancillary
            .get(sheet)
            .and_then(|a| a.vml_drawing_part.clone())
    }

    /// Cached table-part paths for `sheet`, in source order.
    fn _test_ancillary_table_parts(&self, sheet: &str) -> Vec<String> {
        self.ancillary
            .get(sheet)
            .map(|a| a.table_parts.clone())
            .unwrap_or_default()
    }

    /// Cached hyperlink `rId`s for `sheet`, in source order.
    fn _test_ancillary_hyperlink_rids(&self, sheet: &str) -> Vec<String> {
        self.ancillary
            .get(sheet)
            .map(|a| a.hyperlinks_rels.iter().map(|r| r.0.clone()).collect())
            .unwrap_or_default()
    }

    // -------------------------------------------------------------------
    // RFC-022 test-only hooks.
    // -------------------------------------------------------------------

    /// Inject a Set op directly into `queued_hyperlinks`. Mirrors
    /// `queue_hyperlink` but bypasses the validator so tests can set up
    /// odd shapes (e.g. tooltip-only) deliberately.
    fn _test_inject_hyperlink(
        &mut self,
        sheet: &str,
        coord: &str,
        target: Option<String>,
        location: Option<String>,
        tooltip: Option<String>,
        display: Option<String>,
    ) {
        let patch = hyperlinks::HyperlinkPatch {
            coordinate: coord.to_string(),
            target,
            location,
            tooltip,
            display,
        };
        self.queued_hyperlinks
            .entry(sheet.to_string())
            .or_default()
            .insert(coord.to_string(), hyperlinks::HyperlinkOp::Set(patch));
    }

    /// Inject a Delete op directly into `queued_hyperlinks`.
    fn _test_inject_hyperlink_delete(&mut self, sheet: &str, coord: &str) {
        self.queued_hyperlinks
            .entry(sheet.to_string())
            .or_default()
            .insert(coord.to_string(), hyperlinks::HyperlinkOp::Delete);
    }

    /// Run `extract_hyperlinks` on the source ZIP's current sheet XML
    /// and return `(coord, target_or_location)` pairs in BTreeMap order
    /// for assertion in pytest.
    fn _test_get_extracted_hyperlinks(&mut self, sheet: &str) -> PyResult<Vec<(String, String)>> {
        let sheet_path = self
            .sheet_paths
            .get(sheet)
            .cloned()
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("no such sheet: {sheet}")))?;
        let mut zip = open_source_zip(&self.file_path)?;
        let rels_path = sheet_rels_path_for(&sheet_path);
        let rels = load_or_empty_rels(&mut zip, &rels_path)?;
        let xml = ooxml_util::zip_read_to_string(&mut zip, &sheet_path)?;
        let extracted = hyperlinks::extract_hyperlinks(xml.as_bytes(), &rels);
        Ok(extracted
            .into_iter()
            .map(|(coord, h)| {
                let val = h.target.or(h.location).unwrap_or_default();
                (coord, val)
            })
            .collect())
    }
}

// ---------------------------------------------------------------------------
// Save implementation
// ---------------------------------------------------------------------------

impl XlsxPatcher {
    /// Open an existing `.xlsx` package from Rust without constructing Python
    /// dictionaries. This is intentionally small: it supports direct Rust
    /// parity checks for the same two-cell modify workload used by the Rust
    /// competitor benchmark.
    pub fn open_for_rust(path: &str, permissive: bool) -> PyResult<Self> {
        Self::open(path, permissive)
    }

    /// Queue a string cell value from Rust.
    pub fn queue_string_value(&mut self, sheet: &str, cell: &str, value: &str) -> PyResult<()> {
        self.queue_rust_value(sheet, cell, CellValue::String(value.to_string()))
    }

    /// Queue a numeric cell value from Rust.
    pub fn queue_number_value(&mut self, sheet: &str, cell: &str, value: f64) -> PyResult<()> {
        self.queue_rust_value(sheet, cell, CellValue::Number(value))
    }

    /// Save the queued Rust patches to a target path.
    pub fn save_to_path(&mut self, path: &str) -> PyResult<()> {
        self.do_save(path)
    }

    fn queue_rust_value(&mut self, sheet: &str, cell: &str, value: CellValue) -> PyResult<()> {
        if !self.sheet_paths.contains_key(sheet) {
            return Err(PyErr::new::<PyValueError, _>(format!(
                "Unknown sheet for patch: '{sheet}'"
            )));
        }
        let (row, col) = wolfxl_writer::refs::parse_a1(cell).ok_or_else(|| {
            PyErr::new::<PyValueError, _>(format!("Invalid cell reference '{cell}'"))
        })?;
        self.value_patches.insert(
            (sheet.to_string(), cell.to_string()),
            CellPatch {
                row,
                col,
                value: Some(value),
                style_index: None,
            },
        );
        Ok(())
    }

    fn do_save(&mut self, output_path: &str) -> PyResult<()> {
        if !self.has_pending_save_work() {
            patcher_workbook::copy_source_file_phase(self, output_path)?;
            return Ok(());
        }

        let mut zip = open_source_zip(&self.file_path)?;

        // Centralized part-suffix allocator (RFC-035 §5.2 / §8 risk #1).
        // Built once per save; seeded from the source ZIP's part listing
        // so freshly minted tableN / commentsN / vmlDrawingN / sheetN
        // suffixes never collide with source entries. Shared by Phase
        // 2.7 (sheet copies), Phase 2.5f (tables), and Phase 2.5g
        // (comments + VML).
        let mut save = SaveWorkspace::new(
            &mut zip,
            &self.queued_blocks,
            self.save_needs_seeded_part_id_allocator(),
        );

        // RFC-035 §8 risk #6: when Phase 2.7 clones tables onto cloned
        // sheets, those new table names must be visible to Phase 2.5f's
        // collision-scan so a user `add_table` in the same save against
        // an as-yet-unflushed cloned name surfaces a clean error rather
        // than a silent rId/file collision.
        // `file_patches` is the running map of source-ZIP entries that
        // will be REPLACED on emit. Phase 2.7 (RFC-035) is the first
        // phase to write into it (workbook.xml + workbook.xml.rels).
        // Phase 3 mutates it further with per-sheet rewrites.
        patcher_workbook::drain_permissive_seed_file_patches_phase(self, &mut save.file_patches);

        // --- Phase 2.6: Sheet deletes / creates ---
        //
        // Source sheet removal must run before blank-sheet creation so a user
        // can remove a sheet and recreate the same title in one save session.
        if !self.queued_sheet_deletes.is_empty() {
            patcher_workbook::apply_sheet_deletes_phase(self, &mut save.file_patches, &mut zip)?;
        }
        if !self.queued_sheet_creates.is_empty() {
            patcher_workbook::apply_sheet_creates_phase(
                self,
                &mut save.file_patches,
                &mut zip,
                &mut save.part_id_allocator,
            )?;
        }

        // --- Phase 2.7: Sheet copies (RFC-035) ---
        //
        // Drains `queued_sheet_copies` in append order. For each
        // `(src_title, dst_title)` op:
        //   1. Build source rels graph from the ZIP (or from
        //      `rels_patches` / `file_patches` if already mutated).
        //   2. Pre-load source ZIP parts map for the planner.
        //   3. Read workbook.xml (from `file_patches` if already
        //      mutated, else from the source ZIP).
        //   4. Call `wolfxl_structural::sheet_copy::plan_sheet_copy`.
        //   5. Allocate a workbook-rels rId for the new sheet.
        //   6. Apply mutations: insert sheet/rels/ancillary into
        //      `file_adds`; splice `<sheet>` into workbook.xml's
        //      `<sheets>` block; queue cloned sheet-scoped defined
        //      names through `queued_defined_names`; forward
        //      content-type ops through `queued_content_type_ops`;
        //      update `self.sheet_paths` + `self.sheet_order`.
        //
        // Phase-ordering invariant: any new per-sheet phase MUST run
        // AFTER 2.7 (per RFC-035 §8 risk #7). 2.7 runs before Phase
        // 2 / 2.5* so cloned sheets are visible to every downstream
        // per-sheet drain (cell patches, DV, CF, hyperlinks, tables,
        // comments, axis shifts, range moves) as if they had always
        // been part of the source workbook.
        if !self.queued_sheet_copies.is_empty() {
            patcher_sheet_copy::apply_sheet_copies_phase(
                self,
                &mut save.file_patches,
                &mut zip,
                &mut save.part_id_allocator,
                &mut save.cloned_table_names,
            )?;
        }

        // --- Phase 1 / 2: Styles + cell patches ---
        let (mut styles_xml, sheet_cell_patches) =
            patcher_cells::build_sheet_cell_patches_phase(self, &mut zip)?;

        // --- Phase 2.5: Build <dataValidations> blocks from queued DV
        // patches (RFC-025).  Each queued sheet gets exactly one
        // SheetBlock::DataValidations entry whose bytes are the
        // (existing block's children, verbatim) + (new patches,
        // freshly serialized), wrapped in a single <dataValidations
        // count="N">…</dataValidations>.  We read the source sheet
        // XML here so the existing block can flow through unchanged.
        //
        // Pushed into a *local* clone of queued_blocks rather than
        // self — do_save takes &self, and self.queued_blocks is
        // reserved for setters that produce blocks pre-save (future
        // RFCs).  A local map keeps this slice's wiring contained
        // and safe to compose with future block-producing setters.
        // Note: `save.part_id_allocator` is now built earlier (before
        // Phase 2.7) so the centralized allocator can mint cloned-sheet
        // suffixes for RFC-035. Phase 2.5f (tables) + Phase 2.5g
        // (comments + VML) consume the same instance below so
        // workbook-wide suffix uniqueness is preserved across phases.

        patcher_sheet_blocks::apply_data_validations_phase(self, &mut save.local_blocks, &mut zip)?;

        // --- Phase 2.5b: Build <conditionalFormatting> blocks from
        // queued CF patches (RFC-026). Cross-sheet coordination: a
        // single workbook-wide `dxf_count` allocates dxfId values
        // across every sheet's patches in deterministic (sorted) sheet
        // order, and the resulting new `<dxf>` entries are folded into
        // a single `xl/styles.xml` mutation at the end.
        //
        // The merger uses replace-all semantics for slot 17 (RFC-011
        // §5.5) — supplying any CF block drops every existing CF block
        // in the source. We therefore call `extract_existing_cf_blocks`
        // first and re-include them verbatim at the head of each
        // sheet's payload so byte-preservation of unchanged CF rules
        // is not a side-effect of our setter call.
        patcher_sheet_blocks::apply_conditional_formatting_phase(
            self,
            &mut save.local_blocks,
            &mut styles_xml,
            &mut zip,
        )?;

        // --- Phase 2.5e: Hyperlinks (RFC-022) ---
        //
        // Per-sheet flush. For each sheet with queued hyperlink ops:
        //   1. Lazy-populate the ancillary registry so we know which
        //      rIds in the sheet's rels are hyperlinks (vs tables /
        //      comments / vmlDrawings).
        //   2. Get-or-load the rels graph into `rels_patches`. Phase 3's
        //      rels-serialization loop picks up the mutated graph.
        //   3. Read the source sheet XML, extract any existing
        //      `<hyperlinks>` block (resolving rIds → URLs via the rels
        //      graph), and merge with the queued ops.
        //   4. Push a `SheetBlock::Hyperlinks` (slot 19) into
        //      `save.local_blocks` so Phase 3's merge_blocks call inserts it.
        //
        // No `ContentTypeOp`s are emitted here — the worksheet content
        // type is already declared in every source ZIP, and external
        // hyperlinks live in the sheet's rels (which Phase 3 already
        // serializes). An empty `block_bytes` (all hyperlinks deleted
        // and the source had a block) is signaled to the merger by
        // pushing `SheetBlock::Hyperlinks(Vec::new())` — it drops the
        // existing block with no replacement.
        //
        // Cloning `sheet_order` into a local Vec sidesteps the
        // immutable-borrow-on-self-while-mutating-self.{ancillary,
        // rels_patches} conflict (same trick as Phase 2.5d).
        patcher_sheet_blocks::apply_hyperlinks_phase(self, &mut save.local_blocks, &mut zip)?;

        // --- Phase 2.5f: Tables (RFC-024) ---
        //
        // Per-sheet flush. The workbook's existing-table inventory is
        // scanned ONCE up front (table `id` and `name` are
        // workbook-unique, not sheet-scoped, so per-sheet flushes
        // would otherwise risk allocating duplicate ids when two
        // sheets are flushed in the same save). For each sheet with
        // queued tables:
        //   1. Get-or-load the rels graph into `rels_patches` so the
        //      Phase-3 rels-serialization loop picks up the new
        //      TABLE rels we add.
        //   2. Call `tables::build_tables`, which serializes each
        //      patch into `xl/tables/tableN.xml` bytes (reusing the
        //      writer's emitter), allocates fresh rIds in the rels
        //      graph, queues `[Content_Types].xml` Override entries
        //      for each new part, and emits a merged `<tableParts>`
        //      block that includes any pre-existing TABLE rIds plus
        //      the new ones.
        //   3. Inject the new part bytes into `file_adds`.
        //   4. Forward content-type ops into `queued_content_type_ops`
        //      so Phase 2.5c aggregates them into one
        //      `[Content_Types].xml` mutation.
        //   5. Push `SheetBlock::TableParts(block_bytes)` (slot 37)
        //      into `save.local_blocks` so Phase-3's `merge_blocks` call
        //      replaces the sheet's existing `<tableParts>` (if any)
        //      with the merged block.
        //
        // Inventory + ID allocation across sheets: `build_tables`
        // takes a mutable inventory cloned per sheet only — but we
        // thread the names/ids/count manually here so concurrent
        // sheet flushes still see each others' allocations and
        // collisions surface deterministically. (Same trick as the
        // CF cross-sheet dxfId counter in Phase 2.5b.)
        patcher_sheet_blocks::apply_tables_phase(
            self,
            &mut save.local_blocks,
            &save.file_patches,
            &mut zip,
            &mut save.part_id_allocator,
            &save.cloned_table_names,
        )?;

        // --- Phase 2.5g: Comments + VML drawings (RFC-023) ---
        //
        // Per-sheet flush. For each sheet with queued comment ops:
        //   1. Lazy-populate the ancillary registry to learn the
        //      sheet's existing comments part path / VML part path
        //      (if any).
        //   2. Get-or-load the rels graph into `rels_patches`.
        //   3. Read the existing commentsN.xml + vmlDrawingN.vml (if
        //      any), merge in the queued ops, re-emit fresh bytes.
        //   4. Choose a workbook-wide unique `comments_n` / `vml_n`
        //      for sheets gaining their first comments part.
        //   5. Push a `SheetBlock::LegacyDrawing` (slot 31) into
        //      `save.local_blocks` so the merger injects it (deletes the
        //      tag if the rel was removed and the sheet had one).
        //   6. Route comment/vml part bytes:
        //      - if `merged.is_empty()` and no preserved VML shapes
        //        → schedule deletion via `file_deletes`.
        //      - otherwise patch (existing) or add (new) the bytes.
        //   7. Emit `[Content_Types].xml` ops (Override for the
        //      comments part; Default for the vml extension).
        //
        // Workbook-scope author table (`comment_authors`) lives on
        // the stack so all sheets share dedup. New `comments<N>.xml`
        // and `vmlDrawing<N>.vml` suffixes come from the shared
        // `save.part_id_allocator` (RFC-035 §5.2) — already pre-seeded by
        // a single pass over the source ZIP listing earlier in
        // Phase 2.5, so this loop only needs to populate the
        // ancillary registry for path-lookup purposes.
        // --- Phase 2.5g0: Threaded comments + person list (RFC-068 G08). ---
        //
        // Runs BEFORE the legacy comments phase so it can synthesize
        // `tc={topId}` placeholder comments into `queued_comments`, matching
        // the writer-side `synthesize_legacy_placeholders`. Drains
        // `queued_threaded_comments` (per sheet) and `queued_persons`
        // (workbook scope) into fresh `threadedCommentsN.xml` and
        // `personList.xml` part bytes.
        let (threaded_file_writes, threaded_file_deletes) =
            patcher_sheet_blocks::apply_threaded_comments_phase(
                self,
                &mut zip,
                &mut save.part_id_allocator,
            )?;

        let (comments_file_writes, comments_file_deletes) =
            patcher_sheet_blocks::apply_comments_phase(
                self,
                &mut save.local_blocks,
                &mut zip,
                &mut save.part_id_allocator,
            )?;

        // --- Phase 2.5k: Image remove/add (Sprint Λ Pod-β / RFC-045 + G06) ---
        //
        // Drain order is remove first, then add:
        //   * removals mutate existing drawingN.xml + drawing rels
        //   * adds keep the original RFC-045 "fresh drawing only" path
        //
        // Removals are sequenced first so `replace_image` on a sheet that
        // starts with exactly one image can remove the old drawing ref and
        // then create a fresh drawing for the replacement add.
        if !self.queued_image_removes.is_empty() {
            patcher_drawing::apply_image_removes_phase(self, &mut save.file_patches, &mut zip)?;
        }
        //
        // Drains `queued_images` per sheet. For each sheet that has queued images:
        //   1. Read the existing sheet rels (if any) — error if a
        //      `drawing` rel is already present (v1.5 limit:
        //      append-to-existing is a follow-up).
        //   2. Allocate a fresh `drawingN.xml` part via the shared
        //      part-id allocator.
        //   3. Allocate fresh `imageM.<ext>` media parts (one per
        //      queued image).
        //   4. Add an image rel for each one to a brand-new
        //      `xl/drawings/_rels/drawingN.xml.rels`.
        //   5. Add a drawing rel to the sheet's rels graph (creates
        //      `rels_patches` entry for that sheet's rels file).
        //   6. Splice `<drawing r:id="rIdN"/>` into the sheet XML
        //      (right before `<legacyDrawing>` if present, else
        //      before `</worksheet>`).
        //   7. Queue content-type ops: `<Default Extension="png" .../>`
        //      and `<Override PartName="/xl/drawings/drawingN.xml" .../>`.
        //
        // Phase 2.5k runs BEFORE Phase 3 so the cell/block merge
        // pass picks up any drawing-element splice we put into
        // `save.file_patches`. Sheet rels mutations land in
        // `rels_patches` which is serialized in the final emit pass.
        if !self.queued_images.is_empty() {
            patcher_drawing::apply_image_adds_phase(
                self,
                &mut save.file_patches,
                &mut zip,
                &mut save.part_id_allocator,
            )?;
        }

        // --- Phase 2.5l: Chart removes/adds ---
        //
        // Mirrors Phase 2.5k's image-add flow but with two extra
        // capabilities:
        //   * Each queued chart emits its own `xl/charts/chartN.xml`
        //     part, content-type override, and chart-rel under the
        //     drawing's nested rels graph.
        //   * Sheets that already have a `drawing` rel get the new
        //     `<xdr:graphicFrame>` SAX-merged into the existing
        //     drawing XML.
        // Phase 2.5l runs BEFORE Phase 3 so cell-range formulas in
        // chart XML can compose with cell rewrites in the same save.
        if !self.queued_chart_removes.is_empty() {
            patcher_drawing::apply_chart_removes_phase(self, &mut save.file_patches, &mut zip)?;
        }
        if !self.queued_charts.is_empty() {
            patcher_drawing::apply_chart_adds_phase(
                self,
                &mut save.file_patches,
                &mut zip,
                &mut save.part_id_allocator,
            )?;
        }

        // --- Phase 2.5m: Pivot adds (Sprint Ν Pod-γ / RFC-047 + RFC-048) ---
        //
        // Sequenced AFTER 2.5l (charts) and BEFORE Phase 3 (cell
        // patches). See `Plans/sprint-nu.md` Risk #1 for the
        // ordering rationale: charts → pivots → cells. Drainage:
        //
        //   1. For each queued cache:
        //      * Allocate `pivotCacheN` part id.
        //      * Write `xl/pivotCache/pivotCacheDefinition{N}.xml` and
        //        `xl/pivotCache/pivotCacheRecords{N}.xml` to file_adds.
        //      * Build the per-cache rels file pointing at records.
        //      * Add a workbook-rel of type PIVOT_CACHE_DEF.
        //      * Add content-type overrides for both parts.
        //   2. Splice <pivotCaches> into xl/workbook.xml.
        //   3. For each queued table:
        //      * Allocate `pivotTableN` part id.
        //      * Write `xl/pivotTables/pivotTable{N}.xml` to file_adds.
        //      * Build the table → cache rels file.
        //      * Add a sheet-rel of type PIVOT_TABLE.
        //      * Add a content-type override.
        if !self.queued_pivot_caches.is_empty() || !self.queued_pivot_tables.is_empty() {
            patcher_pivot::apply_pivot_adds_phase(self, &mut save.file_patches, &mut zip)?;
        }

        // --- Phase 2.5m-edit: G17 / RFC-070 — pivot source-range
        // mutation of *existing* pivot caches. Sequenced AFTER the
        // adds phase so any cache definition just queued in this
        // session is touched in its post-adds form. The phase is a
        // no-op when no edits have been registered.
        if !self.queued_pivot_source_edits.is_empty() {
            patcher_pivot_edit::apply_pivot_source_edits_phase(
                self,
                &mut save.file_patches,
                &mut zip,
            )?;
        }

        // --- Phase 2.5n: Sheet setup (Sprint Ο Pod 1A.5 / RFC-055) ---
        //
        // Drains queued sheet-setup mutations (sheetView /
        // sheetProtection / pageMargins / pageSetup / headerFooter)
        // into per-sheet `save.local_blocks` for splice via merge_blocks
        // in Phase 3. Sequenced AFTER pivots (2.5m) and BEFORE
        // autoFilter (2.5o) so a later sheet-protection toggle can
        // observe the pivot-table block when computing its allowed
        // operation set.
        //
        // Each non-empty block emits one SheetBlock variant; the
        // merger handles ECMA-376 §18.3.1.99 ordering. The
        // `print_titles` slot routes through workbook-scope
        // <definedNames> (RFC-021 path) — handled by the Workbook
        // coordinator on the Python side; the patcher just stashes
        // the strings on the queue for now.
        if !self.queued_sheet_setup.is_empty() {
            patcher_sheet_blocks::apply_sheet_setup_phase(self, &mut save.local_blocks);
        }

        // --- Phase 2.5r: Page breaks + sheetFormatPr (Sprint Π Pod Π-α / RFC-062) ---
        //
        // Drains queued <rowBreaks> / <colBreaks> / <sheetFormatPr>
        // mutations into per-sheet `save.local_blocks` for splice via
        // merge_blocks in Phase 3. Sequenced AFTER sheet-setup
        // (2.5n) and BEFORE slicers (2.5p) per RFC-062 §6 — page
        // breaks must land before slicer extLst entries because
        // slicer-list refs can anchor to break-bounded cells.
        //
        // Each non-empty slot emits one SheetBlock variant; the
        // merger handles ECMA-376 §18.3.1.99 ordering (slots 4 /
        // 24 / 25).
        if !self.queued_page_breaks.is_empty() {
            patcher_sheet_blocks::apply_page_breaks_phase(self, &mut save.local_blocks);
        }

        // --- Phase 2.5p: Slicer caches + presentations (Sprint Ο Pod 3.5 / RFC-061 §3.1) ---
        //
        // Sequenced AFTER pivots (2.5m) and sheet-setup (2.5n),
        // BEFORE autofilter (2.5o) per RFC-061. For each queued
        // slicer:
        //
        //   1. Allocate a slicer-cache part id + slicer-presentation
        //      part id.
        //   2. Render `xl/slicerCaches/slicerCache{N}.xml` and
        //      `xl/slicers/slicer{M}.xml` via wolfxl_pivot::emit.
        //   3. Build the per-cache rels file pointing at the source
        //      pivot-cache part.
        //   4. Add a workbook-rel of type SLICER_CACHE.
        //   5. Add a sheet-rel of type SLICER.
        //   6. Splice an `<extLst>` `<x14:slicerCaches>` block into
        //      `xl/workbook.xml`.
        //   7. Splice an `<extLst>` `<x14:slicerList>` block into the
        //      owner sheet.
        //   8. Add content-type Overrides for both parts.
        if !self.queued_slicers.is_empty() {
            patcher_pivot::apply_slicer_adds_phase(self, &mut save.file_patches, &mut zip)?;
        }

        // --- Phase 2.5o: AutoFilter (Sprint Ο Pod 1B / RFC-056) ---
        //
        // Sequenced AFTER pivots (2.5m) and BEFORE Phase 3 (cell
        // patches) per RFC-056 §5. For each sheet with a queued
        // `auto_filter`:
        //
        //   1. Lift the §10 dict into the typed `AutoFilter` model.
        //   2. Read the sheet's existing cells inside `auto_filter.ref`
        //      from the source XML (or file_adds for cloned sheets).
        //   3. Run `wolfxl_autofilter::evaluate` to compute the
        //      hidden-row offsets + sort permutation.
        //   4. Push a `SheetBlock::AutoFilter` into `save.local_blocks`
        //      (replaces any existing `<autoFilter>` element).
        //   5. Stash the hidden-row offsets in `autofilter_hidden_rows`
        //      so Phase 3 can apply `<row hidden="1">` markers AFTER
        //      sheet_patcher has rewritten the cell payloads.
        //
        // Sort permutation is computed but **not applied** in v2.0:
        // physical row reorder is deferred to v2.1 per RFC-056 §8.
        let autofilter_hidden_rows = if self.queued_autofilters.is_empty() {
            HashMap::new()
        } else {
            patcher_sheet_blocks::apply_autofilter_phase(
                self,
                &mut save.local_blocks,
                &save.file_patches,
                &mut zip,
            )?
        };

        // --- Phase 2.5t: Existing-table growth ---
        //
        // Rewrites only the grown table parts, before Phase 3 and the
        // structural axis shifts that read table parts from file_patches.
        patcher_sheet_blocks::apply_table_growth_phase(self, &mut save.file_patches, &mut zip)?;

        // --- Phase 3: Patch worksheet XMLs ---
        //
        // Two-pass per sheet: cell-level patches via `sheet_patcher`, then
        // sibling-block insertions via `wolfxl_merger`. The two passes
        // commute (cells live inside <sheetData>, blocks are siblings) so
        // composing them is straightforward.
        //
        // `save.file_patches` was declared early (before Phase 2.7) so RFC-035
        // can write workbook.xml + workbook.xml.rels into it before the
        // per-sheet phases run.

        patcher_sheet_blocks::apply_worksheet_xml_patch_phase(
            self,
            &sheet_cell_patches,
            &save.local_blocks,
            &autofilter_hidden_rows,
            &mut save.file_patches,
            &mut zip,
        )?;

        // Add styles.xml patch if modified
        if let Some(ref sxml) = styles_xml {
            save.file_patches
                .insert("xl/styles.xml".to_string(), sxml.as_bytes().to_vec());
        }

        // --- Phase 2.5q: Workbook security (Sprint Ο Pod 1D / RFC-058) ---
        //
        // Splices `<workbookProtection>` and `<fileSharing>` into
        // `xl/workbook.xml` at canonical CT_Workbook child positions:
        //
        //   fileVersion → fileSharing → workbookPr → workbookProtection
        //   → bookViews → sheets → ...
        //
        // Sequenced AFTER Phase 2.5m (pivots) and BEFORE Phase 2.5h
        // (sheet reorder) so the reorder phase sees the updated
        // workbook.xml (the splice + reorder commute, but composing
        // them through `workbook_xml_in_progress` matches the
        // RFC-035 / RFC-036 hand-off pattern exactly).
        //
        // Empty queue ⇒ identity: workbook.xml flows through
        // unchanged (no extra parse, no extra serialize).
        if self.drop_external_links {
            patcher_external_links::apply_external_links_drop_phase(
                self,
                &mut save.file_patches,
                &mut zip,
            )?;
        }
        patcher_workbook::apply_workbook_xml_phases(self, &mut save.file_patches, &mut zip)?;
        patcher_workbook::apply_sheet_rename_chart_formula_refs_phase(
            self,
            &mut save.file_patches,
            &mut zip,
        )?;

        // Serialize any mutated `*.rels` graphs. Routing depends on whether
        // the path already exists in the source ZIP:
        //   - present → `save.file_patches` replaces it in place (RFC-020 precedent)
        //   - absent  → `file_adds` appends a brand-new entry (RFC-013)
        // The "absent" branch is the common case for RFC-022 on a clean
        // file that had zero hyperlinks before.
        patcher_workbook::serialize_rels_patches_phase(self, &mut save.file_patches, &mut zip)?;

        // --- Phase 2.5c: Content-types aggregation (RFC-013) ---
        //
        // Cross-sheet collection of `ContentTypeOp`s; one parse + serialize
        // of `[Content_Types].xml` regardless of how many sheets contribute.
        // Iteration follows `sheet_order` (source-document order) so the
        // resulting Override sequence is deterministic when multiple sheets
        // each push ops.
        //
        // No live producer in the current slice — `queued_content_type_ops`
        // is always empty, so this loop short-circuits at the
        // `is_empty()` guard. RFC-022 (Hyperlinks via new
        // `xl/worksheets/_rels/sheetN.xml.rels` parts), RFC-023 (Comments
        // via new `xl/comments<N>.xml` Overrides + a vml `Default`),
        // and RFC-024 (Tables via new `xl/tables/tableN.xml` Overrides)
        // will be the first volume producers.
        patcher_workbook::apply_content_types_phase(self, &mut save.file_patches, &mut zip)?;

        // --- Phase 2.5d: Document properties (RFC-020) ---
        //
        // Full rewrite of `docProps/core.xml` + `docProps/app.xml` when
        // `queued_props` is set. Routing depends on whether each part
        // already exists in the source ZIP:
        //   - present → save.file_patches replaces it in place
        //   - absent  → file_adds appends a brand-new entry (RFC-013)
        //
        // `docProps/core.xml` is OPTIONAL in OOXML (some minimal xlsx
        // readers omit it), which is why the file_adds path matters
        // here. See RFC-020 §8 risk #3.
        //
        // If the caller didn't supply `sheet_names`, we thread the
        // patcher's `sheet_order` in so app.xml's `<TitlesOfParts>`
        // matches the workbook's tab order.
        patcher_workbook::apply_document_properties_phase(self, &mut save.file_patches, &mut zip)?;

        // Route RFC-023 comments/vml + RFC-068 threaded-comments/persons
        // part bytes into the right primitive (in-place patch vs. new
        // add) and delete dropped parts. Done after Phase 2.5d so we
        // already know which paths exist in the source ZIP.
        let mut combined_writes = comments_file_writes;
        combined_writes.extend(threaded_file_writes);
        let mut combined_deletes = comments_file_deletes;
        combined_deletes.extend(threaded_file_deletes);
        patcher_workbook::route_part_writes_and_deletes_phase(
            self,
            &mut save.file_patches,
            &mut zip,
            combined_writes,
            combined_deletes,
        );

        // --- Phase 2.5i: Structural axis shifts (RFC-030 / RFC-031) ---
        //
        // Drains `queued_axis_shifts` in append order. For each op:
        //   1. Read sheet XML from `save.file_patches` if already mutated,
        //      else from the source ZIP.
        //   2. Read every table part attached to the sheet (via the
        //      ancillary registry's source-side scan).
        //   3. Read every comments/vmlDrawing part attached to the sheet.
        //   4. Read `xl/workbook.xml` once (cached across ops in this
        //      flush block) for defined-name shifts.
        //   5. Build `wolfxl_structural::SheetXmlInputs` and call
        //      `apply_workbook_shift` with this single op.
        //   6. Merge the returned `save.file_patches` back into our
        //      `save.file_patches`.
        //
        // The empty-queue path is the no-op identity: a workbook with
        // zero queued shifts produces byte-identical output (the
        // outer `is_empty()` short-circuit at the top of `do_save`
        // handles the global no-op case; this block handles the
        // partial case where some other RFC also queued ops).
        if !self.queued_axis_shifts.is_empty() {
            patcher_structural::apply_axis_shifts_phase(self, &mut save.file_patches, &mut zip)?;
        }

        // --- Phase 2.5j: Range moves (RFC-034) ---
        //
        // Drains `queued_range_moves` in append order. Each op reads
        // the affected sheet XML from `save.file_patches` if already
        // mutated (e.g. by Phase 2.5i axis shifts), else from the
        // source ZIP, and routes through
        // `wolfxl_structural::apply_range_move`. Multi-op sequencing
        // mirrors Phase 2.5i: each op runs against the post-previous
        // bytes.
        if !self.queued_range_moves.is_empty() {
            patcher_structural::apply_range_moves_phase(self, &mut save.file_patches, &mut zip)?;
        }

        // --- Phase 2.8: calcChain.xml rebuild (Sprint Θ Pod-C3) ---
        //
        // Walk every sheet in `sheet_order`, scan each sheet's
        // post-mutation XML for formula cells, and emit a fresh
        // `xl/calcChain.xml`. Excel transparently rebuilds this file
        // on next open if it's stale, so the rebuild is a perf-only
        // hint — it never changes correctness. We still do it because
        // (a) it makes Excel's first-open faster, (b) external tools
        // that read calcChain directly see the right cells, and (c)
        // it keeps WolfXL output closer to a freshly-saved Excel
        // file.
        //
        // The no-op short-circuit at the top of `do_save` already
        // bypasses this whole flush, so byte-identical no-op saves
        // are unaffected.
        patcher_workbook::rebuild_calc_chain_phase(self, &mut save.file_patches, &mut zip)?;

        drop(zip);

        // --- Phase 4: Rewrite ZIP ---
        patcher_workbook::rewrite_zip_phase(self, &save.file_patches, output_path)
    }

    fn has_pending_save_work(&self) -> bool {
        !self.queued_image_removes.is_empty() || patcher_workbook::has_pending_save_work(self)
    }

    fn save_needs_seeded_part_id_allocator(&self) -> bool {
        !self.queued_sheet_creates.is_empty()
            || !self.queued_sheet_copies.is_empty()
            || !self.queued_tables.is_empty()
            || !self.queued_threaded_comments.is_empty()
            || !self.queued_persons.is_empty()
            || !self.queued_comments.is_empty()
            || !self.queued_images.is_empty()
            || !self.queued_charts.is_empty()
    }
}
