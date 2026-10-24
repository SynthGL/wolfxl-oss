//! `plan_sheet_copy` — pure planner for `Workbook.copy_worksheet` (RFC-035
//! Phase 7.2).
//!
//! Lives in `wolfxl-structural` (PyO3-free) so its tests can run via
//! `cargo test -p wolfxl-structural` without touching the cdylib's link
//! issue. The patcher (`src/wolfxl/sheet_copy.rs`) is a thin re-export
//! over this module per the spec's API location requirement.
//!
//! See `Plans/rfcs/035-copy-worksheet.md` §4.2, §5.1–§5.7 for the
//! authoritative contract.
//!
//! # Responsibilities
//!
//! 1. Walk the source sheet's rels graph (one level for built-in part
//!    types; nested resolver for drawings → images).
//! 2. Allocate fresh suffixes for each cloneable ancillary part via the
//!    shared [`PartIdAllocator`].
//! 3. Build the `(old_rid → new_rid)` remap and apply it as a SINGLE
//!    pass over the cloned sheet XML (RFC-035 §8 risk #4).
//! 4. Clone each ancillary part's bytes (table parts get name dedup +
//!    re-emitted bytes; comments / VML / drawings flow through verbatim
//!    in this slice).
//! 5. Scan workbook.xml for sheet-scoped defined names whose
//!    `localSheetId == src_idx` and clone each with `localSheetId ==
//!    new_idx` (RFC-035 §5.4 / OQ-c default).
//! 6. Validate (RFC-035 §5.9): error if `dst_title` already exists or if
//!    `src_title` is missing from the sheet list.
//!
//! # Sprint Ο Pod 1B (RFC-056) — autoFilter deep-clone
//!
//! `<autoFilter>` is a sheet-scoped child of `<worksheet>` (slot 11)
//! with no rels-graph dependencies. The planner clones the source
//! sheet bytes verbatim into the new `xl/worksheets/sheet{N}.xml`
//! entry, so `<autoFilter>` and its `<filterColumn>` / `<sortState>`
//! children survive the clone byte-for-byte without requiring a
//! dedicated rewrite.
//!
//! Excel and openpyxl will see the cloned autoFilter on the new
//! sheet immediately. The patcher's Phase 2.5o (Pod 1B's drainage
//! step) re-evaluates the filter on the cloned sheet's data, so if
//! the user then mutates cells on the destination sheet in the same
//! save the `<row hidden>` markers refresh automatically. See
//! `src/wolfxl/mod.rs` Phase 2.5o for the wiring and
//! RFC-056 §6 for the contract.

use std::collections::{HashMap, HashSet};

use quick_xml::events::{BytesStart, BytesText, Event};
use quick_xml::Reader as XmlReader;
use quick_xml::Writer as XmlWriter;

use wolfxl_rels::{
    rt, walk_sheet_subgraph_with_nested, PartIdAllocator, RelId, Relationship, RelsGraph,
    TargetMode,
};

use crate::xml_text::TextNode;

// ---------------------------------------------------------------------------
// Public types
// ---------------------------------------------------------------------------

/// Inputs for one sheet-copy planning call. Borrows everything the
/// planner needs from the patcher's current view of the workbook —
/// including `source_zip_parts`, which the caller pre-loads (the
/// planner does not perform ZIP I/O).
pub struct SheetCopyInputs<'a> {
    /// Source sheet title (for error messages + workbook.xml lookup).
    pub src_title: String,
    /// Destination sheet title (for the new `<sheet>` entry). Must be
    /// pre-deduped by the caller.
    pub dst_title: String,
    /// Source sheet ZIP part path, e.g. `xl/worksheets/sheet3.xml`.
    pub src_sheet_path: String,
    /// Pre-loaded source ZIP parts: `(path, bytes)`. Includes the
    /// sheet itself, every table/comments/VML/drawing part it
    /// references, and any nested-rels parts (drawing's own rels file
    /// for the image case).
    pub source_zip_parts: &'a HashMap<String, Vec<u8>>,
    /// The source sheet's parsed rels graph
    /// (`xl/worksheets/_rels/sheet<src>.xml.rels`). May be empty.
    pub source_rels: &'a RelsGraph,
    /// `xl/workbook.xml` bytes — used to scan sheet-scoped
    /// `<definedName>` entries and to compute `src_idx`.
    pub workbook_xml: &'a [u8],
    /// Centralized part-suffix allocator. Mutated by the planner.
    pub allocator: &'a mut PartIdAllocator,
    /// Workbook-wide set of table `name`/`displayName` values that are
    /// already taken (for §5.5 dedup).
    pub existing_table_names: &'a HashSet<String>,
    /// Sprint Θ Pod-C2 — when `true`, drawings reachable from the
    /// source sheet have their referenced `xl/media/imageN.<ext>`
    /// targets DEEP-CLONED into freshly numbered media parts. The
    /// cloned drawing's nested rels file is rewritten to point at
    /// the new paths and the image bytes are added to
    /// `new_ancillary_parts`.
    ///
    /// The default (`false`) preserves the historical RFC-035 §5.3
    /// alias-by-target behaviour.
    pub deep_copy_images: bool,
}

/// One cloned defined-name (RFC-035 §5.4). The planner returns the
/// XML bytes ready to splice into `<definedNames>`; the caller
/// concatenates and merges.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DefinedNameClone {
    /// `name` attribute, verbatim from the source defined name.
    pub name: String,
    /// `localSheetId`, recomputed to point at the new sheet's
    /// position in the post-copy tab list.
    pub local_sheet_id: u32,
    /// Formula text (verbatim from source — sheet copies share
    /// coordinate space, so cell references resolve correctly without
    /// translation).
    pub formula: String,
    /// Pre-rendered `<definedName …>…</definedName>` element bytes
    /// for direct splicing.
    pub element_bytes: Vec<u8>,
}

/// All mutations the planner produces for one sheet copy. The caller
/// (Phase 2.7 of `XlsxPatcher::do_save`) applies these atomically.
#[derive(Debug, Clone, Default)]
pub struct SheetCopyMutations {
    /// New sheet's ZIP part path, e.g. `xl/worksheets/sheet5.xml`.
    pub new_sheet_path: String,
    /// New sheet's XML bytes (source bytes with the `r:id` remap
    /// applied).
    pub new_sheet_xml: Vec<u8>,
    /// `(zip_path, bytes)` for every cloned ancillary part: tables,
    /// comments, VML, drawings, drawing rels files (when present).
    /// Image media is NOT in this list — images are aliased per
    /// RFC-035 §5.3 / §8 risk #2.
    pub new_ancillary_parts: Vec<(String, Vec<u8>)>,
    /// `(part_path, content_type)` overrides to add to
    /// `[Content_Types].xml`. `part_path` is leading-slashed per OPC
    /// convention (`"/xl/tables/table7.xml"`).
    pub content_type_overrides_to_add: Vec<(String, String)>,
    /// `<sheet name="…" sheetId="…" r:id="…"/>` element bytes for the
    /// caller to splice into workbook.xml's `<sheets>` block. The
    /// caller fills in `r:id` AFTER allocating the workbook-rels
    /// entry (we don't know the workbook rels graph here).
    pub workbook_sheets_append: Vec<u8>,
    /// `(rId-placeholder, rel_type, target)` for the workbook rels
    /// entry the caller must add. The `rId-placeholder` is the same
    /// string used in `workbook_sheets_append` so the caller can do a
    /// single replace-all pass after allocating the actual rId.
    pub workbook_rels_to_add: Vec<(String, String, String)>,
    /// Workbook-level slicer-cache rels to add, plus the slicer-cache
    /// name that must be referenced from workbook.xml `<x14:slicerCaches>`.
    pub workbook_slicer_cache_rels_to_add: Vec<(String, String, String)>,
    /// Workbook-level timeline-cache rels to add, plus the timeline-cache
    /// name that must be referenced from workbook.xml `<x15:timelineCacheRefs>`.
    pub workbook_timeline_cache_rels_to_add: Vec<(String, String, String)>,
    /// Sheet-scoped defined-name clones (RFC-035 §5.4).
    pub defined_names_to_add: Vec<DefinedNameClone>,
    /// Newly allocated table `name`/`displayName` values added to
    /// `existing_table_names` for collision detection across multiple
    /// in-flight copies. The caller folds these into the running set
    /// before the next `plan_sheet_copy` call.
    pub new_table_names: Vec<String>,
    /// `old_rid → new_rid` for the source sheet's local rels graph.
    /// Useful for callers that want to do additional rewrites on the
    /// cloned sheet bytes; the planner has already applied this remap
    /// to `new_sheet_xml`.
    pub rid_remap: HashMap<String, String>,
}

/// Errors `plan_sheet_copy` may return.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum SheetCopyError {
    /// `dst_title` already names an existing sheet in workbook.xml.
    DuplicateDestinationTitle(String),
    /// `src_title` was not found in workbook.xml's `<sheets>` list.
    MissingSourceTitle(String),
    /// Source sheet bytes were not in `source_zip_parts`.
    MissingSourceSheetBytes(String),
    /// XML parse / serialization error.
    Xml(String),
}

impl std::fmt::Display for SheetCopyError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            SheetCopyError::DuplicateDestinationTitle(s) => {
                write!(f, "destination sheet title {s:?} already exists")
            }
            SheetCopyError::MissingSourceTitle(s) => {
                write!(f, "source sheet title {s:?} not found in workbook.xml")
            }
            SheetCopyError::MissingSourceSheetBytes(s) => {
                write!(f, "source sheet bytes for {s:?} not provided")
            }
            SheetCopyError::Xml(e) => write!(f, "xml error: {e}"),
        }
    }
}

impl std::error::Error for SheetCopyError {}

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

const CT_WORKSHEET: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const CT_TABLE: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";
const CT_COMMENTS: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";
const CT_DRAWING: &str = "application/vnd.openxmlformats-officedocument.drawing+xml";
/// Sprint Μ Pod-γ (RFC-046 §7) — DrawingML chart content type.
const CT_CHART: &str = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
const CT_CHART_STYLE: &str = "application/vnd.ms-office.chartstyle+xml";
const CT_CHART_COLOR_STYLE: &str = "application/vnd.ms-office.chartcolorstyle+xml";
const CT_CTRL_PROP: &str = "application/vnd.ms-excel.controlproperties+xml";
const CT_TIMELINE: &str = "application/vnd.ms-excel.timeline+xml";
const CT_TIMELINE_CACHE: &str = "application/vnd.ms-excel.timelineCache+xml";
const RT_TIMELINE: &str = "http://schemas.microsoft.com/office/2011/relationships/timeline";
const RT_TIMELINE_CACHE: &str =
    "http://schemas.microsoft.com/office/2011/relationships/timelineCache";

// ---------------------------------------------------------------------------
// Planner entry point
// ---------------------------------------------------------------------------

/// Plan one sheet copy. Pure: does not mutate any caller-owned data
/// other than `inputs.allocator` (which is `&mut`).
pub fn plan_sheet_copy(inputs: SheetCopyInputs<'_>) -> Result<SheetCopyMutations, SheetCopyError> {
    // ---- Validate ----------------------------------------------------------
    let SheetMetadata {
        sheet_titles,
        sheet_ids,
        defined_names,
    } = parse_workbook_metadata(inputs.workbook_xml)?;

    let src_idx = sheet_titles
        .iter()
        .position(|t| t == &inputs.src_title)
        .ok_or_else(|| SheetCopyError::MissingSourceTitle(inputs.src_title.clone()))?
        as u32;

    if sheet_titles.iter().any(|t| t == &inputs.dst_title) {
        return Err(SheetCopyError::DuplicateDestinationTitle(
            inputs.dst_title.clone(),
        ));
    }

    let new_idx = sheet_titles.len() as u32; // appended at end
    let source_sheet_id = sheet_ids
        .get(src_idx as usize)
        .copied()
        .unwrap_or(src_idx + 1);
    let new_sheet_id = sheet_ids.iter().copied().max().unwrap_or(0) + 1;

    let src_sheet_xml = inputs
        .source_zip_parts
        .get(&inputs.src_sheet_path)
        .ok_or_else(|| SheetCopyError::MissingSourceSheetBytes(inputs.src_sheet_path.clone()))?;

    // ---- Walk the rels subgraph -------------------------------------------
    // Resolver: for any reachable part we hit, look up its `_rels` file (if
    // any) in source_zip_parts and parse it. This catches drawings → images
    // and chart parts → media.
    let zip_parts = inputs.source_zip_parts;
    let subgraph =
        walk_sheet_subgraph_with_nested(inputs.source_rels, &inputs.src_sheet_path, |part_path| {
            let rels_path = wolfxl_rels::rels_path_for(part_path)?;
            let bytes = zip_parts.get(&rels_path)?;
            RelsGraph::parse(bytes).ok()
        });

    // ---- Allocate destination part suffixes + build rid_remap -------------
    let new_sheet_n = inputs.allocator.alloc_sheet();
    let new_sheet_path = format!("xl/worksheets/sheet{new_sheet_n}.xml");

    // Build the destination rels graph alongside the rid_remap. We assign
    // freshly monotonic rIds in the new graph so the remap is stable. The
    // destination rels graph is NOT serialized here — the caller (Phase 2.7)
    // wires it into the patcher's `rels_patches` map.
    let mut dest_rels = RelsGraph::new();
    let mut rid_remap: HashMap<String, String> = HashMap::new();

    let mut new_ancillary_parts: Vec<(String, Vec<u8>)> = Vec::new();
    let mut content_type_overrides_to_add: Vec<(String, String)> = Vec::new();
    let mut new_table_names: Vec<String> = Vec::new();
    let mut table_id_remap: HashMap<String, String> = HashMap::new();
    let mut table_name_remap: HashMap<String, String> = HashMap::new();
    let mut slicer_cache_name_remap: HashMap<String, String> = HashMap::new();
    let mut slicer_name_remap: HashMap<String, String> = HashMap::new();
    let mut workbook_slicer_cache_rels_to_add: Vec<(String, String, String)> = Vec::new();
    let mut workbook_timeline_cache_rels_to_add: Vec<(String, String, String)> = Vec::new();
    let mut taken_slicer_cache_names = workbook_slicer_cache_names(zip_parts);
    let mut taken_slicer_names = slicer_names_from_parts(zip_parts);
    let mut taken_timeline_cache_names = workbook_timeline_cache_names(zip_parts);
    let mut taken_timeline_names = timeline_names_from_parts(zip_parts);
    let mut timeline_cache_name_remap: HashMap<String, String> = HashMap::new();
    let mut timeline_name_remap: HashMap<String, String> = HashMap::new();
    inputs
        .allocator
        .observe_vml_shape_id_max(max_shape_id_from_parts(zip_parts));
    let mut control_shape_id_remap: HashMap<String, String> = HashMap::new();
    for old_id in control_shape_ids_from_sheet(src_sheet_xml) {
        control_shape_id_remap
            .entry(old_id)
            .or_insert_with(|| inputs.allocator.alloc_vml_shape_id().to_string());
    }

    // For tracking the "set already used in this session" to dedup tables
    // when one copy clones two tables with the same base name.
    let mut taken_names: HashSet<String> = inputs.existing_table_names.clone();

    // Cache of per-part rels graphs keyed by part path so we don't re-parse.
    let mut nested_rels_cache: HashMap<String, RelsGraph> = HashMap::new();
    for part_path in &subgraph.reachable_parts {
        if let Some(rp) = wolfxl_rels::rels_path_for(part_path) {
            if let Some(bytes) = zip_parts.get(&rp) {
                if let Ok(g) = RelsGraph::parse(bytes) {
                    nested_rels_cache.insert(part_path.clone(), g);
                }
            }
        }
    }

    let mut table_clone_ids: HashMap<String, u32> = HashMap::new();
    for source_rel in inputs.source_rels.iter() {
        if source_rel.rel_type != rt::TABLE {
            continue;
        }
        let new_n = inputs.allocator.alloc_table();
        table_clone_ids.insert(source_rel.id.0.clone(), new_n);
        let resolved = resolve_relative(parent_dir(&inputs.src_sheet_path), &source_rel.target);
        if let Some(src_bytes) = zip_parts.get(&resolved) {
            if let Some(old_id) = table_id_from_xml(src_bytes) {
                table_id_remap.insert(old_id, new_n.to_string());
            }
        }
    }

    // Walk the source sheet's edges in source order and clone each part.
    for source_rel in inputs.source_rels.iter() {
        let old_rid = source_rel.id.0.clone();
        match source_rel.rel_type.as_str() {
            // ------ Tables (need name dedup + bytes re-emit) ------
            t if t == rt::TABLE => {
                let resolved =
                    resolve_relative(parent_dir(&inputs.src_sheet_path), &source_rel.target);
                let new_n = table_clone_ids
                    .get(&old_rid)
                    .copied()
                    .unwrap_or_else(|| inputs.allocator.alloc_table());
                let new_part_path = format!("xl/tables/table{new_n}.xml");
                // Clone bytes with name + id dedup.
                let src_bytes = zip_parts.get(&resolved).cloned().unwrap_or_default();
                let (cloned_bytes, new_name, old_name) =
                    clone_table_part(&src_bytes, &mut taken_names, new_n)?;
                if let Some(old_name) = old_name {
                    table_name_remap.insert(old_name, new_name.clone());
                }
                new_ancillary_parts.push((new_part_path.clone(), cloned_bytes));
                new_table_names.push(new_name.clone());
                taken_names.insert(new_name);
                content_type_overrides_to_add
                    .push((format!("/{}", new_part_path), CT_TABLE.into()));
                let new_target = format!("../tables/table{new_n}.xml");
                let new_rid = dest_rels.add(rt::TABLE, &new_target, TargetMode::Internal);
                rid_remap.insert(old_rid, new_rid.0);
            }
            // ------ Comments ------
            t if t == rt::COMMENTS => {
                let resolved =
                    resolve_relative(parent_dir(&inputs.src_sheet_path), &source_rel.target);
                let new_n = inputs.allocator.alloc_comments();
                let new_part_path = format!("xl/comments{new_n}.xml");
                let src_bytes = zip_parts.get(&resolved).cloned().unwrap_or_default();
                let new_target = format!("../comments{new_n}.xml");
                let mut outputs = SheetRelCloneOutputs {
                    dest_rels: &mut dest_rels,
                    rid_remap: &mut rid_remap,
                    new_ancillary_parts: &mut new_ancillary_parts,
                    content_type_overrides_to_add: &mut content_type_overrides_to_add,
                };
                add_internal_part_clone(
                    source_rel,
                    InternalPartClone {
                        bytes: src_bytes,
                        new_part_path,
                        content_type: Some(CT_COMMENTS),
                        new_target,
                    },
                    &mut outputs,
                );
            }
            // ------ VML drawings ------
            t if t == rt::VML_DRAWING => {
                let resolved =
                    resolve_relative(parent_dir(&inputs.src_sheet_path), &source_rel.target);
                let new_n = inputs.allocator.alloc_vml_drawing();
                let new_part_path = format!("xl/drawings/vmlDrawing{new_n}.vml");
                let src_bytes = zip_parts.get(&resolved).cloned().unwrap_or_default();
                // VML uses a Default content-type by extension; not an
                // Override, but we still surface it so the caller can
                // ensure_default("vml", ...) at the content-types layer.
                let new_target = format!("../drawings/vmlDrawing{new_n}.vml");
                let mut outputs = SheetRelCloneOutputs {
                    dest_rels: &mut dest_rels,
                    rid_remap: &mut rid_remap,
                    new_ancillary_parts: &mut new_ancillary_parts,
                    content_type_overrides_to_add: &mut content_type_overrides_to_add,
                };
                add_internal_part_clone(
                    source_rel,
                    InternalPartClone {
                        bytes: src_bytes,
                        new_part_path,
                        content_type: None,
                        new_target,
                    },
                    &mut outputs,
                );
            }
            // ------ Form-control property parts ------
            //
            // A sheet control's DrawingML/VML shape is sheet-local, and the
            // matching ctrlProp part must be sheet-local too. Sharing the
            // source `ctrlPropN.xml` across the source and cloned sheet opens
            // in Excel, but real Excel renders some copied controls with
            // different geometry.
            t if t == rt::CTRL_PROP => {
                let resolved =
                    resolve_relative(parent_dir(&inputs.src_sheet_path), &source_rel.target);
                let new_n = inputs.allocator.alloc_ctrl_prop();
                let new_part_path = format!("xl/ctrlProps/ctrlProp{new_n}.xml");
                let src_bytes = zip_parts.get(&resolved).cloned().unwrap_or_default();
                let new_target = format!("../ctrlProps/ctrlProp{new_n}.xml");
                let mut outputs = SheetRelCloneOutputs {
                    dest_rels: &mut dest_rels,
                    rid_remap: &mut rid_remap,
                    new_ancillary_parts: &mut new_ancillary_parts,
                    content_type_overrides_to_add: &mut content_type_overrides_to_add,
                };
                add_internal_part_clone(
                    source_rel,
                    InternalPartClone {
                        bytes: src_bytes,
                        new_part_path,
                        content_type: Some(CT_CTRL_PROP),
                        new_target,
                    },
                    &mut outputs,
                );
            }
            // ------ DrawingML drawings (own rels file → image alias OR deep-clone) ------
            t if t == rt::DRAWING => {
                let resolved =
                    resolve_relative(parent_dir(&inputs.src_sheet_path), &source_rel.target);
                let new_n = inputs.allocator.alloc_drawing();
                let new_part_path = format!("xl/drawings/drawing{new_n}.xml");
                let src_bytes = zip_parts.get(&resolved).cloned().unwrap_or_default();
                new_ancillary_parts.push((new_part_path.clone(), src_bytes));
                content_type_overrides_to_add
                    .push((format!("/{}", new_part_path), CT_DRAWING.into()));
                let new_target = format!("../drawings/drawing{new_n}.xml");
                let new_rid = dest_rels.add(rt::DRAWING, &new_target, TargetMode::Internal);
                rid_remap.insert(old_rid, new_rid.0);
                // Clone the drawing's own rels file IF present.
                //
                // - Default (alias) mode: the image target stays
                //   pointed at the same xl/media/imageN.{ext} path
                //   (RFC-035 §5.3).
                // - Deep-copy mode (Sprint Θ Pod-C2): allocate a
                //   fresh `xl/media/imageM.<ext>` per image rel,
                //   copy its bytes into `new_ancillary_parts`, and
                //   re-point the cloned drawing rel at the new
                //   target.
                if let Some(nested) = nested_rels_cache.get(&resolved) {
                    let mut cloned = RelsGraph::new();
                    // Sprint Μ Pod-γ (RFC-046 §7): track per-chart
                    // old_rid → new_rid remap so we can patch the
                    // drawing XML's <c:chart r:id="..."/> references.
                    let mut chart_rid_remap: HashMap<String, String> = HashMap::new();
                    for nrel in nested.iter() {
                        if inputs.deep_copy_images && nrel.rel_type == rt::IMAGE {
                            let resolved_image =
                                resolve_relative(parent_dir(&resolved), &nrel.target);
                            // Determine the original extension; fall
                            // back to "png" if the original had none
                            // (rare in practice but keeps us safe).
                            let ext = resolved_image
                                .rsplit_once('.')
                                .map(|(_, e)| e.to_string())
                                .unwrap_or_else(|| "png".to_string());
                            let new_image_n = inputs.allocator.alloc_image();
                            let new_image_path = format!("xl/media/image{new_image_n}.{ext}");
                            // Copy the original bytes if present.
                            if let Some(bytes) = zip_parts.get(&resolved_image).cloned() {
                                new_ancillary_parts.push((new_image_path.clone(), bytes));
                            }
                            // The rel's `target` is drawing-relative
                            // ("../media/imageN.<ext>"); we rewrite
                            // it to the new suffix.
                            let new_rel_target = format!("../media/image{new_image_n}.{ext}");
                            cloned.add(&nrel.rel_type, &new_rel_target, nrel.mode);
                        } else if nrel.rel_type == rt::CHART {
                            // Sprint Μ Pod-γ (RFC-046 §7) — deep-clone
                            // every chart referenced by the drawing.
                            // Always-on (no opt-in flag). Cell ranges
                            // in <c:f> formulas get re-pointed from
                            // src_title to dst_title using the formula
                            // translator's `rename_sheet`.
                            let resolved_chart =
                                resolve_relative(parent_dir(&resolved), &nrel.target);
                            let new_chart_n = inputs.allocator.alloc_chart();
                            let new_chart_path = format!("xl/charts/chart{new_chart_n}.xml");
                            let cloned_bytes: Vec<u8> =
                                if let Some(src_bytes) = zip_parts.get(&resolved_chart) {
                                    rewrite_chart_sheet_refs(
                                        src_bytes,
                                        &inputs.src_title,
                                        &inputs.dst_title,
                                    )
                                } else {
                                    Vec::new()
                                };
                            new_ancillary_parts.push((new_chart_path.clone(), cloned_bytes));
                            content_type_overrides_to_add
                                .push((format!("/{new_chart_path}"), CT_CHART.into()));
                            // Drawing-rels target is drawing-relative
                            // ("../charts/chartN.xml"); rewrite to the
                            // new suffix and stash the rId mapping so
                            // we can patch the drawing XML below.
                            let new_rel_target = format!("../charts/chart{new_chart_n}.xml");
                            let new_chart_rid =
                                cloned.add(&nrel.rel_type, &new_rel_target, nrel.mode);
                            chart_rid_remap.insert(nrel.id.0.clone(), new_chart_rid.0);
                            // If the chart has its own rels file, clone
                            // chart-local dependencies such as style/color
                            // parts instead of sharing them across chart
                            // parts; desktop Excel rejects some duplicated
                            // chart dependency targets on open.
                            let chart_rels_src = wolfxl_rels::rels_path_for(&resolved_chart);
                            if let Some(rels_path_src) = chart_rels_src {
                                if let Some(rb) = zip_parts.get(&rels_path_src) {
                                    let chart_rels_dst =
                                        format!("xl/charts/_rels/chart{new_chart_n}.xml.rels");
                                    let cloned_rels = clone_chart_rels_with_dependencies(
                                        rb,
                                        &resolved_chart,
                                        zip_parts,
                                        inputs.allocator,
                                        &mut new_ancillary_parts,
                                        &mut content_type_overrides_to_add,
                                    );
                                    new_ancillary_parts.push((chart_rels_dst, cloned_rels));
                                }
                            }
                        } else {
                            cloned.add(&nrel.rel_type, &nrel.target, nrel.mode);
                        }
                    }
                    let nested_rels_path = format!("xl/drawings/_rels/drawing{new_n}.xml.rels");
                    new_ancillary_parts.push((nested_rels_path, cloned.serialize()));
                    // Sprint Μ Pod-γ — patch the just-cloned drawing
                    // XML so any `<c:chart r:id="rIdOLD"/>` tokens
                    // point at the new chart rIds. We applied this
                    // AFTER the verbatim drawing-bytes push above so
                    // we mutate the entry in place.
                    if !chart_rid_remap.is_empty() {
                        if let Some(slot) = new_ancillary_parts
                            .iter_mut()
                            .find(|(p, _)| p == &new_part_path)
                        {
                            slot.1 = rewrite_drawing_chart_rids(&slot.1, &chart_rid_remap);
                        }
                    }
                }
            }
            // ------ Charts directly off the sheet (chartsheet ref) ------
            // Sprint Μ Pod-γ (RFC-046 §7) — deep-clone with cell-
            // range re-pointing. Always-on (no opt-in flag, unlike
            // `deep_copy_images`). Embedded charts on a normal
            // worksheet are NOT handled here — those hang off the
            // drawing, see the drawing-rels nested clone block above.
            t if t == rt::CHART => {
                let resolved =
                    resolve_relative(parent_dir(&inputs.src_sheet_path), &source_rel.target);
                let new_chart_n = inputs.allocator.alloc_chart();
                let new_chart_path = format!("xl/charts/chart{new_chart_n}.xml");
                let cloned_bytes: Vec<u8> = if let Some(src_bytes) = zip_parts.get(&resolved) {
                    rewrite_chart_sheet_refs(src_bytes, &inputs.src_title, &inputs.dst_title)
                } else {
                    Vec::new()
                };
                let new_target = format!("../charts/chart{new_chart_n}.xml");
                let mut outputs = SheetRelCloneOutputs {
                    dest_rels: &mut dest_rels,
                    rid_remap: &mut rid_remap,
                    new_ancillary_parts: &mut new_ancillary_parts,
                    content_type_overrides_to_add: &mut content_type_overrides_to_add,
                };
                add_internal_part_clone(
                    source_rel,
                    InternalPartClone {
                        bytes: cloned_bytes,
                        new_part_path: new_chart_path,
                        content_type: Some(CT_CHART),
                        new_target,
                    },
                    &mut outputs,
                );
                let chart_rels_src = wolfxl_rels::rels_path_for(&resolved);
                if let Some(rels_path_src) = chart_rels_src {
                    if let Some(rb) = zip_parts.get(&rels_path_src) {
                        let chart_rels_dst = format!("xl/charts/_rels/chart{new_chart_n}.xml.rels");
                        let cloned_rels = clone_chart_rels_with_dependencies(
                            rb,
                            &resolved,
                            zip_parts,
                            inputs.allocator,
                            &mut new_ancillary_parts,
                            &mut content_type_overrides_to_add,
                        );
                        new_ancillary_parts.push((chart_rels_dst, cloned_rels));
                    }
                }
            }
            // ------ Pivot tables (Sprint Ν Pod-γ / RFC-035 §10) ------
            //
            // Deep-clone strategy:
            //   * Allocate a fresh `pivotTableN` suffix.
            //   * Bytes-rewrite the table XML (rewrite
            //     `worksheetSource sheet="<src>"` → `<dst>`; this is
            //     a no-op for the common case where the cache's
            //     source range lives on a sheet *other* than the one
            //     being copied, but it's the right behaviour for the
            //     edge case where the source range itself is on the
            //     copied sheet — see §10.4 "self-cache copy").
            //   * Copy the per-table rels file verbatim (the rels
            //     point at the cache definition; the cache stays
            //     aliased).
            //   * Add a content-type override.
            //   * Emit a sheet-rel pointing at the new pivotTableN
            //     part. The cache itself is NOT deep-cloned in this
            //     pass (workbook-scope; same cache → many tables).
            t if t == rt::PIVOT_TABLE => {
                let resolved_table =
                    resolve_relative(parent_dir(&inputs.src_sheet_path), &source_rel.target);
                let new_n = inputs.allocator.alloc_pivot_table();
                let new_table_path = format!("xl/pivotTables/pivotTable{new_n}.xml");

                let cloned_bytes: Vec<u8> = if let Some(src_bytes) = zip_parts.get(&resolved_table)
                {
                    wolfxl_pivot::structural::deep_clone_pivot_table(
                        src_bytes,
                        &inputs.src_title,
                        &inputs.dst_title,
                        None,
                    )
                } else {
                    Vec::new()
                };
                new_ancillary_parts.push((new_table_path.clone(), cloned_bytes));
                content_type_overrides_to_add.push((
                    format!("/{new_table_path}"),
                    wolfxl_pivot::ct::PIVOT_TABLE.to_string(),
                ));

                // Copy per-table rels file (table → cache) verbatim.
                // The rels target stays pointed at the original cache
                // definition since we don't deep-clone caches in
                // sheet copy (RFC-047 §6: caches are workbook-scope).
                if let Some(src_table_rels_path) = wolfxl_rels::rels_path_for(&resolved_table) {
                    if let Some(rb) = zip_parts.get(&src_table_rels_path) {
                        let dst_table_rels_path =
                            format!("xl/pivotTables/_rels/pivotTable{new_n}.xml.rels");
                        new_ancillary_parts.push((dst_table_rels_path, rb.clone()));
                    }
                }

                let new_target = format!("../pivotTables/pivotTable{new_n}.xml");
                let new_rid = dest_rels.add(rt::PIVOT_TABLE, &new_target, TargetMode::Internal);
                rid_remap.insert(old_rid, new_rid.0);
            }
            // ------ Slicer presentations (Sprint Ο Pod 3.5 / RFC-061 §6) ------
            //
            // Deep-clone strategy:
            //   * Allocate a fresh `slicer{N}.xml` suffix (the
            //     destination sheet must not share the source sheet's
            //     slicer presentation file, since each presentation
            //     is anchored to one sheet).
            //   * Bytes-copy the source slicer XML verbatim. The
            //     `cache="..."` attribute references the slicer-cache
            //     by name, NOT by rId, so the cache is automatically
            //     shared without any rewrite.
            //   * Add a content-type override.
            //   * Emit a sheet-rel of type SLICER pointing at the new
            //     slicer{N}.xml part.
            //
            // The slicerCache itself is NOT deep-cloned (workbook-
            // scope per RFC-061 §6: many slicer presentations can
            // reference one cache).
            t if t == wolfxl_pivot::rt::SLICER => {
                let resolved =
                    resolve_relative(parent_dir(&inputs.src_sheet_path), &source_rel.target);
                let new_n = inputs.allocator.alloc_slicer();
                let new_part_path = format!("xl/slicers/slicer{new_n}.xml");
                let src_bytes = zip_parts.get(&resolved).cloned().unwrap_or_default();
                let (src_bytes, new_cache_refs) = clone_slicer_caches_for_presentation(
                    &src_bytes,
                    zip_parts,
                    inputs.allocator,
                    &table_id_remap,
                    source_sheet_id,
                    new_sheet_id,
                    &mut taken_slicer_cache_names,
                    &mut taken_slicer_names,
                    &mut slicer_cache_name_remap,
                    &mut slicer_name_remap,
                    &mut new_ancillary_parts,
                    &mut content_type_overrides_to_add,
                );
                workbook_slicer_cache_rels_to_add.extend(new_cache_refs);
                let new_target = format!("../slicers/slicer{new_n}.xml");
                let mut outputs = SheetRelCloneOutputs {
                    dest_rels: &mut dest_rels,
                    rid_remap: &mut rid_remap,
                    new_ancillary_parts: &mut new_ancillary_parts,
                    content_type_overrides_to_add: &mut content_type_overrides_to_add,
                };
                add_internal_part_clone(
                    source_rel,
                    InternalPartClone {
                        bytes: src_bytes,
                        new_part_path,
                        content_type: Some(wolfxl_pivot::ct::SLICER),
                        new_target,
                    },
                    &mut outputs,
                );
            }
            // ------ Timeline presentations (Excel pivot timelines) ------
            t if t == RT_TIMELINE => {
                let resolved =
                    resolve_relative(parent_dir(&inputs.src_sheet_path), &source_rel.target);
                let new_n = inputs.allocator.alloc_timeline();
                let new_part_path = format!("xl/timelines/timeline{new_n}.xml");
                let src_bytes = zip_parts.get(&resolved).cloned().unwrap_or_default();
                let (src_bytes, new_cache_refs) = clone_timeline_cache_for_presentation(
                    &src_bytes,
                    zip_parts,
                    inputs.allocator,
                    source_sheet_id,
                    new_sheet_id,
                    &mut taken_timeline_cache_names,
                    &mut taken_timeline_names,
                    &mut timeline_cache_name_remap,
                    &mut timeline_name_remap,
                    &mut new_ancillary_parts,
                    &mut content_type_overrides_to_add,
                );
                workbook_timeline_cache_rels_to_add.extend(new_cache_refs);
                let new_target = format!("../timelines/timeline{new_n}.xml");
                let mut outputs = SheetRelCloneOutputs {
                    dest_rels: &mut dest_rels,
                    rid_remap: &mut rid_remap,
                    new_ancillary_parts: &mut new_ancillary_parts,
                    content_type_overrides_to_add: &mut content_type_overrides_to_add,
                };
                add_internal_part_clone(
                    source_rel,
                    InternalPartClone {
                        bytes: src_bytes,
                        new_part_path,
                        content_type: Some(CT_TIMELINE),
                        new_target,
                    },
                    &mut outputs,
                );
            }
            // ------ Hyperlinks (External: alias by URL) ------
            t if t == rt::HYPERLINK => {
                let new_rid =
                    dest_rels.add(&source_rel.rel_type, &source_rel.target, source_rel.mode);
                rid_remap.insert(old_rid, new_rid.0);
            }
            // ------ Anything else (alias the target verbatim) ------
            _ => {
                let new_rid =
                    dest_rels.add(&source_rel.rel_type, &source_rel.target, source_rel.mode);
                rid_remap.insert(old_rid, new_rid.0);
            }
        }
    }

    // Emit the destination rels file alongside the cloned sheet so the
    // caller can drop both into file_adds in one shot.
    if !dest_rels.is_empty() {
        let dest_rels_path = wolfxl_rels::rels_path_for(&new_sheet_path)
            .unwrap_or_else(|| format!("_rels/{new_sheet_path}.rels"));
        new_ancillary_parts.push((dest_rels_path, dest_rels.serialize()));
    }

    // ---- Apply the rId remap to the cloned sheet XML in one pass ----------
    let new_sheet_xml = rewrite_rids(src_sheet_xml, &rid_remap)?;
    let new_sheet_xml = rewrite_control_shape_ids(&new_sheet_xml, &control_shape_id_remap);

    // ---- RFC-035 + RFC-057: translate any explicit-sheet refs inside
    // the cloned sheet's <f> formula text AND the <f t="array" ref="..."/>
    // attribute when the source sheet is being renamed during the
    // copy.  Plain unqualified refs (e.g. `=A1`) are sheet-local and
    // need no translation.  Qualified refs (e.g. `=Sheet1!A1`) get
    // re-pointed via the formula translator so the cloned sheet's
    // cross-sheet references continue to point at the right cells.
    let new_sheet_xml = if inputs.src_title != inputs.dst_title {
        rewrite_sheet_formula_refs(&new_sheet_xml, &inputs.src_title, &inputs.dst_title)
    } else {
        new_sheet_xml
    };
    let new_sheet_xml = if table_name_remap.is_empty() {
        new_sheet_xml
    } else {
        rewrite_table_formula_refs(&new_sheet_xml, &table_name_remap)
    };
    if !slicer_name_remap.is_empty() {
        for (path, bytes) in &mut new_ancillary_parts {
            if path.starts_with("xl/drawings/drawing") && path.ends_with(".xml") {
                *bytes = rewrite_drawing_object_names(bytes, &slicer_name_remap);
            }
        }
    }
    if !timeline_name_remap.is_empty() {
        for (path, bytes) in &mut new_ancillary_parts {
            if path.starts_with("xl/drawings/drawing") && path.ends_with(".xml") {
                *bytes = rewrite_drawing_object_names(bytes, &timeline_name_remap);
            }
        }
    }
    if !control_shape_id_remap.is_empty() {
        for (path, bytes) in &mut new_ancillary_parts {
            if path.starts_with("xl/drawings/drawing") && path.ends_with(".xml") {
                *bytes = rewrite_control_shape_ids(bytes, &control_shape_id_remap);
            } else if path.starts_with("xl/drawings/vmlDrawing") && path.ends_with(".vml") {
                let control_ids = vml_control_shape_ids(bytes, &control_shape_id_remap);
                if !control_ids.is_empty() {
                    *bytes =
                        rewrite_vml_control_shape_ids(bytes, &control_shape_id_remap, &control_ids);
                    *bytes = rewrite_vml_control_idmap_data(bytes, &control_shape_id_remap);
                }
            }
        }
    }

    // ---- Sheet-scoped defined-name clones ---------------------------------
    let defined_names_to_add: Vec<DefinedNameClone> = defined_names
        .iter()
        .filter(|dn| dn.local_sheet_id == Some(src_idx))
        .map(|dn| {
            let element_bytes = render_defined_name(&dn.name, new_idx, &dn.formula);
            DefinedNameClone {
                name: dn.name.clone(),
                local_sheet_id: new_idx,
                formula: dn.formula.clone(),
                element_bytes,
            }
        })
        .collect();

    // ---- New <sheet> entry for workbook.xml -------------------------------
    let placeholder_rid = format!("__SHEET_RID_PLACEHOLDER_{new_sheet_n}__");
    let workbook_sheets_append =
        render_sheet_entry(&inputs.dst_title, new_sheet_id, &placeholder_rid);
    let workbook_rels_to_add = vec![(
        placeholder_rid,
        rt::WORKSHEET.to_string(),
        format!("worksheets/sheet{new_sheet_n}.xml"),
    )];

    // Add the worksheet content-type override.
    content_type_overrides_to_add.insert(0, (format!("/{new_sheet_path}"), CT_WORKSHEET.into()));

    // Avoid `unused`: subgraph is informational; we acted on inputs.source_rels
    // directly. Keep a debug-friendly check that the sheet itself was visited.
    debug_assert_eq!(
        subgraph.reachable_parts.first().map(|s| s.as_str()),
        Some(inputs.src_sheet_path.as_str())
    );

    Ok(SheetCopyMutations {
        new_sheet_path,
        new_sheet_xml,
        new_ancillary_parts,
        content_type_overrides_to_add,
        workbook_sheets_append,
        workbook_rels_to_add,
        workbook_slicer_cache_rels_to_add,
        workbook_timeline_cache_rels_to_add,
        defined_names_to_add,
        new_table_names,
        rid_remap,
    })
}

struct InternalPartClone {
    bytes: Vec<u8>,
    new_part_path: String,
    content_type: Option<&'static str>,
    new_target: String,
}

struct SheetRelCloneOutputs<'a> {
    dest_rels: &'a mut RelsGraph,
    rid_remap: &'a mut HashMap<String, String>,
    new_ancillary_parts: &'a mut Vec<(String, Vec<u8>)>,
    content_type_overrides_to_add: &'a mut Vec<(String, String)>,
}

fn add_internal_part_clone(
    source_rel: &Relationship,
    clone: InternalPartClone,
    outputs: &mut SheetRelCloneOutputs<'_>,
) {
    outputs
        .new_ancillary_parts
        .push((clone.new_part_path.clone(), clone.bytes));
    if let Some(content_type) = clone.content_type {
        outputs.content_type_overrides_to_add.push((
            format!("/{}", clone.new_part_path),
            content_type.to_string(),
        ));
    }
    let new_rid = outputs.dest_rels.add(
        &source_rel.rel_type,
        &clone.new_target,
        TargetMode::Internal,
    );
    outputs.rid_remap.insert(source_rel.id.0.clone(), new_rid.0);
}

fn clone_chart_rels_with_dependencies(
    chart_rels_bytes: &[u8],
    source_chart_path: &str,
    zip_parts: &HashMap<String, Vec<u8>>,
    allocator: &mut PartIdAllocator,
    new_ancillary_parts: &mut Vec<(String, Vec<u8>)>,
    content_type_overrides_to_add: &mut Vec<(String, String)>,
) -> Vec<u8> {
    let Ok(source_rels) = RelsGraph::parse(chart_rels_bytes) else {
        return chart_rels_bytes.to_vec();
    };
    let mut cloned_rels = RelsGraph::new();
    for rel in source_rels.iter() {
        let mut target = rel.target.clone();
        if rel.mode == TargetMode::Internal && rel.rel_type == rt::CHART_STYLE {
            if let Some((new_path, new_target)) = clone_chart_sidecar_part(
                source_chart_path,
                &rel.target,
                "style",
                ".xml",
                CT_CHART_STYLE,
                zip_parts,
                allocator.alloc_chart_style(),
                new_ancillary_parts,
                content_type_overrides_to_add,
            ) {
                allocator.observe(&new_path);
                target = new_target;
            }
        } else if rel.mode == TargetMode::Internal && rel.rel_type == rt::CHART_COLOR_STYLE {
            if let Some((new_path, new_target)) = clone_chart_sidecar_part(
                source_chart_path,
                &rel.target,
                "colors",
                ".xml",
                CT_CHART_COLOR_STYLE,
                zip_parts,
                allocator.alloc_chart_color(),
                new_ancillary_parts,
                content_type_overrides_to_add,
            ) {
                allocator.observe(&new_path);
                target = new_target;
            }
        }
        cloned_rels.add_with_id(RelId(rel.id.0.clone()), &rel.rel_type, &target, rel.mode);
    }
    cloned_rels.serialize()
}

fn clone_chart_sidecar_part(
    source_chart_path: &str,
    source_target: &str,
    prefix: &str,
    suffix: &str,
    content_type: &str,
    zip_parts: &HashMap<String, Vec<u8>>,
    new_n: u32,
    new_ancillary_parts: &mut Vec<(String, Vec<u8>)>,
    content_type_overrides_to_add: &mut Vec<(String, String)>,
) -> Option<(String, String)> {
    let resolved = resolve_relative(parent_dir(source_chart_path), source_target);
    let bytes = zip_parts.get(&resolved)?.clone();
    let new_path = format!("xl/charts/{prefix}{new_n}{suffix}");
    let new_target = format!("{prefix}{new_n}{suffix}");
    new_ancillary_parts.push((new_path.clone(), bytes));
    content_type_overrides_to_add.push((format!("/{new_path}"), content_type.to_string()));
    Some((new_path, new_target))
}

// ---------------------------------------------------------------------------
// Helpers — workbook.xml metadata parse
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, Default)]
struct SheetMetadata {
    /// Sheet titles in document order (the position is the
    /// `localSheetId`).
    sheet_titles: Vec<String>,
    /// Workbook `<sheet sheetId="...">` values in document order.
    sheet_ids: Vec<u32>,
    defined_names: Vec<ParsedDefinedName>,
}

#[derive(Debug, Clone)]
struct ParsedDefinedName {
    name: String,
    /// `Some(idx)` if scoped to a sheet position; `None` if
    /// workbook-scoped.
    local_sheet_id: Option<u32>,
    /// Formula text (verbatim, unescaped).
    formula: String,
}

fn parse_workbook_metadata(workbook_xml: &[u8]) -> Result<SheetMetadata, SheetCopyError> {
    let mut reader = XmlReader::from_reader(workbook_xml);
    let mut buf: Vec<u8> = Vec::new();
    let mut meta = SheetMetadata::default();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.local_name().as_ref() {
                "sheet" => {
                    let name = attr_value(&e, b"name").unwrap_or_default();
                    let sheet_id = attr_value(&e, b"sheetId")
                        .and_then(|v| v.parse::<u32>().ok())
                        .unwrap_or((meta.sheet_titles.len() as u32) + 1);
                    meta.sheet_titles.push(name);
                    meta.sheet_ids.push(sheet_id);
                }
                "definedName" => {
                    let name = attr_value(&e, b"name").unwrap_or_default();
                    let lsid = attr_value(&e, b"localSheetId").and_then(|v| v.parse::<u32>().ok());
                    meta.defined_names.push(ParsedDefinedName {
                        name,
                        local_sheet_id: lsid,
                        formula: String::new(),
                    });
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(SheetCopyError::Xml(format!("workbook.xml: {e}"))),
            _ => {}
        }
        buf.clear();
    }

    // Second pass to populate formulas — quick-xml's nested-event API makes
    // a single pass tricky for `<definedName>foo</definedName>` text nodes.
    populate_defined_name_formulas(workbook_xml, &mut meta.defined_names)?;

    Ok(meta)
}

fn populate_defined_name_formulas(
    workbook_xml: &[u8],
    defined_names: &mut [ParsedDefinedName],
) -> Result<(), SheetCopyError> {
    let mut reader = XmlReader::from_reader(workbook_xml);
    let mut buf: Vec<u8> = Vec::new();
    let mut in_dn: Option<usize> = None;
    let mut idx = 0;
    let mut formula = TextNode::default();

    loop {
        let event = reader.read_event_into(&mut buf);
        if !formula.is_empty() && !matches!(event, Ok(Event::Text(_)) | Ok(Event::GeneralRef(_))) {
            let text = formula
                .unescaped()
                .map_err(|e| SheetCopyError::Xml(format!("dn formula: {e}")))?;
            if let Some(dn) = in_dn.and_then(|i| defined_names.get_mut(i)) {
                dn.formula.push_str(text);
            }
            formula.clear();
        }
        match event {
            Ok(Event::Start(e)) if e.local_name().as_ref() == "definedName" => {
                in_dn = Some(idx);
            }
            Ok(Event::Text(t)) if in_dn.is_some() => formula.push_text(&t),
            Ok(Event::GeneralRef(r)) if in_dn.is_some() => formula.push_ref(&r),
            Ok(Event::End(e)) if e.local_name().as_ref() == "definedName" => {
                in_dn = None;
                idx += 1;
            }
            Ok(Event::Empty(e)) if e.local_name().as_ref() == "definedName" => {
                // Self-closing definedName has no formula text.
                idx += 1;
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(SheetCopyError::Xml(format!("dn formula: {e}"))),
            _ => {}
        }
        buf.clear();
    }

    Ok(())
}

fn attr_value(e: &BytesStart<'_>, key: &[u8]) -> Option<String> {
    for a in e.attributes().with_checks(false).flatten() {
        if a.key.as_ref().as_bytes() == key {
            return Some(
                quick_xml::escape::unescape(&a.value)
                    .map(|v| v.into_owned())
                    .unwrap_or_else(|_| a.value.into_owned()),
            );
        }
    }
    None
}

// ---------------------------------------------------------------------------
// Helpers — table clone + name dedup
// ---------------------------------------------------------------------------

/// Per RFC-035 §5.5 / OQ-b default:
/// `f"{base}_{N}"` starting at N=2.
fn dedup_table_name(base: &str, taken: &HashSet<String>) -> String {
    if !taken.contains(base) {
        // Even if the base isn't taken, RFC-035's contract is: every
        // CLONED table gets a non-source name. We always append `_2`+ to
        // make the divergence loud.
        let mut suffix = 2u32;
        let mut candidate = format!("{base}_{suffix}");
        while taken.contains(&candidate) {
            suffix += 1;
            candidate = format!("{base}_{suffix}");
        }
        return candidate;
    }
    let mut suffix = 2u32;
    let mut candidate = format!("{base}_{suffix}");
    while taken.contains(&candidate) {
        suffix += 1;
        candidate = format!("{base}_{suffix}");
    }
    candidate
}

/// Re-emit a `tableN.xml` part with: (1) a deduped `name` /
/// `displayName`; (2) a fresh `id` derived from `new_n` (workbook-unique
/// suffix is sufficient since callers track every emitted id via the
/// allocator). Other attributes flow through verbatim.
fn clone_table_part(
    src_bytes: &[u8],
    taken_names: &mut HashSet<String>,
    new_n: u32,
) -> Result<(Vec<u8>, String, Option<String>), SheetCopyError> {
    let mut reader = XmlReader::from_reader(src_bytes);
    let mut writer = XmlWriter::new(Vec::with_capacity(src_bytes.len()));
    let mut buf: Vec<u8> = Vec::new();
    let mut renamed: Option<String> = None;
    let mut original_name: Option<String> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Decl(d)) => {
                writer
                    .write_event(Event::Decl(d))
                    .map_err(|e| SheetCopyError::Xml(format!("table decl: {e}")))?;
            }
            Ok(Event::Start(e)) if e.local_name().as_ref() == "table" => {
                let (new_e, new_name, old_name) = clone_table_start(&e, taken_names, new_n)?;
                renamed = Some(new_name);
                original_name = old_name;
                writer
                    .write_event(Event::Start(new_e))
                    .map_err(|e| SheetCopyError::Xml(format!("table root: {e}")))?;
            }
            Ok(Event::Empty(e)) if e.local_name().as_ref() == "table" => {
                let (new_e, new_name, old_name) = clone_table_start(&e, taken_names, new_n)?;
                renamed = Some(new_name);
                original_name = old_name;
                writer
                    .write_event(Event::Empty(new_e))
                    .map_err(|e| SheetCopyError::Xml(format!("table root: {e}")))?;
            }
            Ok(Event::Eof) => break,
            Ok(other) => {
                writer
                    .write_event(other)
                    .map_err(|e| SheetCopyError::Xml(format!("table flow: {e}")))?;
            }
            Err(e) => return Err(SheetCopyError::Xml(format!("table parse: {e}"))),
        }
        buf.clear();
    }

    let bytes = writer.into_inner();
    let new_name = renamed.unwrap_or_else(|| {
        // Source had no name attribute — should never happen in real OOXML
        // but handle gracefully.
        let synth = format!("Table{new_n}");
        taken_names.insert(synth.clone());
        synth
    });
    Ok((bytes, new_name, original_name))
}

fn clone_table_start(
    e: &BytesStart<'_>,
    taken_names: &HashSet<String>,
    new_n: u32,
) -> Result<(BytesStart<'static>, String, Option<String>), SheetCopyError> {
    let mut attrs: Vec<(String, String)> = Vec::new();
    let mut source_name: Option<String> = None;
    let mut source_display_name: Option<String> = None;
    for a in e.attributes().with_checks(false).flatten() {
        let key = a.key.as_ref().to_owned();
        let value = quick_xml::escape::unescape(&a.value)
            .map(|v| v.into_owned())
            .unwrap_or_else(|_| a.value.into_owned());
        match key.as_str() {
            "name" => source_name = Some(value.clone()),
            "displayName" => source_display_name = Some(value.clone()),
            _ => {}
        }
        attrs.push((key, value));
    }

    let original_name = source_name.or(source_display_name);
    let new_name = original_name
        .as_ref()
        .map(|name| dedup_table_name(name, taken_names))
        .unwrap_or_else(|| format!("Table{new_n}"));

    let mut new_e = BytesStart::new(e.name().as_ref().to_owned());
    let mut wrote_name = false;
    let mut wrote_display_name = false;
    let mut wrote_id = false;

    for (key, value) in attrs {
        match key.as_str() {
            "name" => {
                push_attr_escaped(&mut new_e, "name", &new_name);
                wrote_name = true;
            }
            "displayName" => {
                push_attr_escaped(&mut new_e, "displayName", &new_name);
                wrote_display_name = true;
            }
            "id" => {
                push_attr_escaped(&mut new_e, "id", &new_n.to_string());
                wrote_id = true;
            }
            _ => {
                push_attr_escaped(&mut new_e, &key, &value);
            }
        }
    }

    if !wrote_id {
        push_attr_escaped(&mut new_e, "id", &new_n.to_string());
    }
    if !wrote_name {
        push_attr_escaped(&mut new_e, "name", &new_name);
    }
    if !wrote_display_name {
        push_attr_escaped(&mut new_e, "displayName", &new_name);
    }

    Ok((new_e, new_name, original_name))
}

fn table_id_from_xml(src_bytes: &[u8]) -> Option<String> {
    let mut reader = XmlReader::from_reader(src_bytes);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) if e.local_name().as_ref() == "table" => {
                return attr_value(&e, b"id");
            }
            Ok(Event::Eof) => return None,
            Err(_) => return None,
            _ => {}
        }
        buf.clear();
    }
}

fn workbook_slicer_cache_names(zip_parts: &HashMap<String, Vec<u8>>) -> HashSet<String> {
    zip_parts
        .iter()
        .filter(|(path, _)| {
            path.starts_with("xl/slicerCaches/slicerCache") && path.ends_with(".xml")
        })
        .filter_map(|(_, bytes)| slicer_cache_name_from_xml(bytes))
        .collect()
}

fn slicer_names_from_parts(zip_parts: &HashMap<String, Vec<u8>>) -> HashSet<String> {
    let mut names = HashSet::new();
    for (path, bytes) in zip_parts {
        if !(path.starts_with("xl/slicers/slicer") && path.ends_with(".xml")) {
            continue;
        }
        let mut reader = XmlReader::from_reader(bytes.as_slice());
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) | Ok(Event::Empty(e))
                    if e.local_name().as_ref() == "slicer" =>
                {
                    if let Some(name) = attr_value(&e, b"name") {
                        names.insert(name);
                    }
                }
                Ok(Event::Eof) => break,
                Err(_) => break,
                _ => {}
            }
            buf.clear();
        }
    }
    names
}

fn slicer_cache_name_from_xml(src_bytes: &[u8]) -> Option<String> {
    let mut reader = XmlReader::from_reader(src_bytes);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e))
                if e.local_name().as_ref() == "slicerCacheDefinition" =>
            {
                return attr_value(&e, b"name");
            }
            Ok(Event::Eof) => return None,
            Err(_) => return None,
            _ => {}
        }
        buf.clear();
    }
}

fn clone_slicer_caches_for_presentation(
    slicer_xml: &[u8],
    zip_parts: &HashMap<String, Vec<u8>>,
    allocator: &mut PartIdAllocator,
    table_id_remap: &HashMap<String, String>,
    source_sheet_id: u32,
    new_sheet_id: u32,
    taken_cache_names: &mut HashSet<String>,
    taken_slicer_names: &mut HashSet<String>,
    cache_name_remap: &mut HashMap<String, String>,
    slicer_name_remap: &mut HashMap<String, String>,
    new_ancillary_parts: &mut Vec<(String, Vec<u8>)>,
    content_type_overrides_to_add: &mut Vec<(String, String)>,
) -> (Vec<u8>, Vec<(String, String, String)>) {
    let cache_parts = slicer_cache_parts_by_name(zip_parts);
    let mut workbook_refs = Vec::new();
    let mut reader = XmlReader::from_reader(slicer_xml);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) if e.local_name().as_ref() == "slicer" => {
                let Some(old_cache_name) = attr_value(&e, b"cache") else {
                    buf.clear();
                    continue;
                };
                let Some((_cache_part, cache_bytes)) = cache_parts.get(&old_cache_name) else {
                    buf.clear();
                    continue;
                };
                if !cache_name_remap.contains_key(&old_cache_name) {
                    let cloned_cache =
                        if let Some(source_table_id) = table_slicer_cache_table_id(cache_bytes) {
                            let Some(new_table_id) = table_id_remap.get(&source_table_id) else {
                                buf.clear();
                                continue;
                            };
                            let new_cache_name =
                                unique_with_numeric_suffix(&old_cache_name, taken_cache_names, "");
                            taken_cache_names.insert(new_cache_name.clone());
                            let cloned_cache = rewrite_table_slicer_cache(
                                cache_bytes,
                                &old_cache_name,
                                &new_cache_name,
                                &source_table_id,
                                new_table_id,
                            );
                            cache_name_remap.insert(old_cache_name.clone(), new_cache_name);
                            cloned_cache
                        } else if let Some(source_tab_id) = pivot_slicer_cache_tab_id(cache_bytes) {
                            if source_tab_id != source_sheet_id.to_string() {
                                buf.clear();
                                continue;
                            }
                            let new_cache_name =
                                unique_with_numeric_suffix(&old_cache_name, taken_cache_names, "");
                            taken_cache_names.insert(new_cache_name.clone());
                            let cloned_cache = rewrite_pivot_slicer_cache(
                                cache_bytes,
                                &old_cache_name,
                                &new_cache_name,
                                &source_tab_id,
                                &new_sheet_id.to_string(),
                            );
                            cache_name_remap.insert(old_cache_name.clone(), new_cache_name);
                            cloned_cache
                        } else {
                            buf.clear();
                            continue;
                        };
                    let new_cache_name = cache_name_remap
                        .get(&old_cache_name)
                        .cloned()
                        .unwrap_or_else(|| old_cache_name.clone());
                    let new_cache_n = allocator.alloc_slicer_cache();
                    let new_cache_path = format!("xl/slicerCaches/slicerCache{new_cache_n}.xml");
                    allocator.observe(&new_cache_path);
                    new_ancillary_parts.push((new_cache_path.clone(), cloned_cache));
                    content_type_overrides_to_add.push((
                        format!("/{new_cache_path}"),
                        wolfxl_pivot::ct::SLICER_CACHE.to_string(),
                    ));
                    workbook_refs.push((
                        new_cache_name.clone(),
                        wolfxl_pivot::rt::SLICER_CACHE.to_string(),
                        format!("slicerCaches/slicerCache{new_cache_n}.xml"),
                    ));
                }
                if let Some(old_slicer_name) = attr_value(&e, b"name") {
                    slicer_name_remap
                        .entry(old_slicer_name.clone())
                        .or_insert_with(|| {
                            let new_name = unique_with_numeric_suffix(
                                &old_slicer_name,
                                taken_slicer_names,
                                " ",
                            );
                            taken_slicer_names.insert(new_name.clone());
                            new_name
                        });
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    let rewritten = rewrite_slicer_presentation(slicer_xml, cache_name_remap, slicer_name_remap);
    (rewritten, workbook_refs)
}

fn slicer_cache_parts_by_name(
    zip_parts: &HashMap<String, Vec<u8>>,
) -> HashMap<String, (String, Vec<u8>)> {
    let mut out = HashMap::new();
    for (path, bytes) in zip_parts {
        if !(path.starts_with("xl/slicerCaches/slicerCache") && path.ends_with(".xml")) {
            continue;
        }
        if let Some(name) = slicer_cache_name_from_xml(bytes) {
            out.insert(name, (path.clone(), bytes.clone()));
        }
    }
    out
}

fn table_slicer_cache_table_id(src_bytes: &[u8]) -> Option<String> {
    let mut reader = XmlReader::from_reader(src_bytes);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e))
                if e.local_name().as_ref() == "tableSlicerCache" =>
            {
                return attr_value(&e, b"tableId");
            }
            Ok(Event::Eof) => return None,
            Err(_) => return None,
            _ => {}
        }
        buf.clear();
    }
}

fn pivot_slicer_cache_tab_id(src_bytes: &[u8]) -> Option<String> {
    first_attr_on_local_element(src_bytes, b"pivotTable", b"tabId")
}

fn workbook_timeline_cache_names(zip_parts: &HashMap<String, Vec<u8>>) -> HashSet<String> {
    zip_parts
        .iter()
        .filter(|(path, _)| {
            path.starts_with("xl/timelineCaches/timelineCache") && path.ends_with(".xml")
        })
        .filter_map(|(_, bytes)| timeline_cache_name_from_xml(bytes))
        .collect()
}

fn timeline_names_from_parts(zip_parts: &HashMap<String, Vec<u8>>) -> HashSet<String> {
    let mut names = HashSet::new();
    for (path, bytes) in zip_parts {
        if !(path.starts_with("xl/timelines/timeline") && path.ends_with(".xml")) {
            continue;
        }
        if let Some(name) = first_attr_on_local_element(bytes, b"timeline", b"name") {
            names.insert(name);
        }
    }
    names
}

fn timeline_cache_name_from_xml(src_bytes: &[u8]) -> Option<String> {
    first_attr_on_local_element(src_bytes, b"timelineCacheDefinition", b"name")
}

fn timeline_cache_parts_by_name(
    zip_parts: &HashMap<String, Vec<u8>>,
) -> HashMap<String, (String, Vec<u8>)> {
    let mut out = HashMap::new();
    for (path, bytes) in zip_parts {
        if !(path.starts_with("xl/timelineCaches/timelineCache") && path.ends_with(".xml")) {
            continue;
        }
        if let Some(name) = timeline_cache_name_from_xml(bytes) {
            out.insert(name, (path.clone(), bytes.clone()));
        }
    }
    out
}

fn clone_timeline_cache_for_presentation(
    timeline_xml: &[u8],
    zip_parts: &HashMap<String, Vec<u8>>,
    allocator: &mut PartIdAllocator,
    source_sheet_id: u32,
    new_sheet_id: u32,
    taken_cache_names: &mut HashSet<String>,
    taken_timeline_names: &mut HashSet<String>,
    cache_name_remap: &mut HashMap<String, String>,
    timeline_name_remap: &mut HashMap<String, String>,
    new_ancillary_parts: &mut Vec<(String, Vec<u8>)>,
    content_type_overrides_to_add: &mut Vec<(String, String)>,
) -> (Vec<u8>, Vec<(String, String, String)>) {
    let cache_parts = timeline_cache_parts_by_name(zip_parts);
    let mut workbook_refs = Vec::new();
    let mut reader = XmlReader::from_reader(timeline_xml);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) if e.local_name().as_ref() == "timeline" => {
                let Some(old_cache_name) = attr_value(&e, b"cache") else {
                    buf.clear();
                    continue;
                };
                let Some((_cache_part, cache_bytes)) = cache_parts.get(&old_cache_name) else {
                    buf.clear();
                    continue;
                };
                if !cache_name_remap.contains_key(&old_cache_name) {
                    let Some(source_tab_id) = pivot_slicer_cache_tab_id(cache_bytes) else {
                        buf.clear();
                        continue;
                    };
                    if source_tab_id != source_sheet_id.to_string() {
                        buf.clear();
                        continue;
                    }
                    let new_cache_name =
                        unique_with_numeric_suffix(&old_cache_name, taken_cache_names, "");
                    taken_cache_names.insert(new_cache_name.clone());
                    let cloned_cache = rewrite_pivot_slicer_cache(
                        cache_bytes,
                        &old_cache_name,
                        &new_cache_name,
                        &source_tab_id,
                        &new_sheet_id.to_string(),
                    );
                    let new_cache_n = allocator.alloc_timeline_cache();
                    let new_cache_path =
                        format!("xl/timelineCaches/timelineCache{new_cache_n}.xml");
                    allocator.observe(&new_cache_path);
                    new_ancillary_parts.push((new_cache_path.clone(), cloned_cache));
                    content_type_overrides_to_add
                        .push((format!("/{new_cache_path}"), CT_TIMELINE_CACHE.to_string()));
                    workbook_refs.push((
                        new_cache_name.clone(),
                        RT_TIMELINE_CACHE.to_string(),
                        format!("timelineCaches/timelineCache{new_cache_n}.xml"),
                    ));
                    cache_name_remap.insert(old_cache_name.clone(), new_cache_name);
                }
                if let Some(old_timeline_name) = attr_value(&e, b"name") {
                    timeline_name_remap
                        .entry(old_timeline_name.clone())
                        .or_insert_with(|| {
                            let new_name = unique_with_numeric_suffix(
                                &old_timeline_name,
                                taken_timeline_names,
                                " ",
                            );
                            taken_timeline_names.insert(new_name.clone());
                            new_name
                        });
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    let rewritten =
        rewrite_timeline_presentation(timeline_xml, cache_name_remap, timeline_name_remap);
    (rewritten, workbook_refs)
}

fn unique_with_numeric_suffix(base: &str, taken: &HashSet<String>, sep: &str) -> String {
    let mut suffix = 1_u32;
    loop {
        let candidate = format!("{base}{sep}{suffix}");
        if !taken.contains(&candidate) {
            return candidate;
        }
        suffix += 1;
    }
}

fn rewrite_table_slicer_cache(
    cache_xml: &[u8],
    old_name: &str,
    new_name: &str,
    old_table_id: &str,
    new_table_id: &str,
) -> Vec<u8> {
    let mut attr_maps = HashMap::new();
    attr_maps.insert(
        "name".to_string(),
        HashMap::from([(old_name.to_string(), new_name.to_string())]),
    );
    attr_maps.insert(
        "tableId".to_string(),
        HashMap::from([(old_table_id.to_string(), new_table_id.to_string())]),
    );
    rewrite_attributes_by_local_name(cache_xml, &attr_maps)
}

fn rewrite_pivot_slicer_cache(
    cache_xml: &[u8],
    old_name: &str,
    new_name: &str,
    old_tab_id: &str,
    new_tab_id: &str,
) -> Vec<u8> {
    let mut attr_maps = HashMap::new();
    attr_maps.insert(
        "name".to_string(),
        HashMap::from([(old_name.to_string(), new_name.to_string())]),
    );
    attr_maps.insert(
        "tabId".to_string(),
        HashMap::from([(old_tab_id.to_string(), new_tab_id.to_string())]),
    );
    rewrite_attributes_by_local_name(cache_xml, &attr_maps)
}

fn rewrite_slicer_presentation(
    slicer_xml: &[u8],
    cache_name_remap: &HashMap<String, String>,
    slicer_name_remap: &HashMap<String, String>,
) -> Vec<u8> {
    let mut attr_maps = HashMap::new();
    attr_maps.insert("cache".to_string(), cache_name_remap.clone());
    attr_maps.insert("name".to_string(), slicer_name_remap.clone());
    rewrite_attributes_by_local_name(slicer_xml, &attr_maps)
}

fn rewrite_timeline_presentation(
    timeline_xml: &[u8],
    cache_name_remap: &HashMap<String, String>,
    timeline_name_remap: &HashMap<String, String>,
) -> Vec<u8> {
    let mut attr_maps = HashMap::new();
    attr_maps.insert("cache".to_string(), cache_name_remap.clone());
    attr_maps.insert("name".to_string(), timeline_name_remap.clone());
    rewrite_attributes_by_local_name(timeline_xml, &attr_maps)
}

fn rewrite_drawing_object_names(xml: &[u8], name_remap: &HashMap<String, String>) -> Vec<u8> {
    let mut attr_maps = HashMap::new();
    attr_maps.insert("name".to_string(), name_remap.clone());
    rewrite_attributes_by_local_name(xml, &attr_maps)
}

fn rewrite_control_shape_ids(xml: &[u8], id_remap: &HashMap<String, String>) -> Vec<u8> {
    if id_remap.is_empty() {
        return xml.to_vec();
    }
    let mut shape_ref_remap = HashMap::new();
    for (old_id, new_id) in id_remap {
        shape_ref_remap.insert(old_id.clone(), new_id.clone());
        shape_ref_remap.insert(format!("_x0000_s{old_id}"), format!("_x0000_s{new_id}"));
    }
    let mut attr_maps = HashMap::new();
    attr_maps.insert("shapeId".to_string(), id_remap.clone());
    attr_maps.insert("id".to_string(), shape_ref_remap.clone());
    attr_maps.insert("spid".to_string(), shape_ref_remap);
    rewrite_attributes_by_local_name(xml, &attr_maps)
}

fn rewrite_vml_control_shape_ids(
    xml: &[u8],
    id_remap: &HashMap<String, String>,
    control_ids: &HashSet<String>,
) -> Vec<u8> {
    if control_ids.is_empty() {
        return xml.to_vec();
    }
    let vml_id_remap: HashMap<String, String> = id_remap
        .iter()
        .map(|(old_id, new_id)| (format!("_x0000_s{old_id}"), format!("_x0000_s{new_id}")))
        .filter(|(old_id, _)| control_ids.contains(old_id))
        .collect();
    if vml_id_remap.is_empty() {
        return xml.to_vec();
    }
    let mut attr_maps = HashMap::new();
    attr_maps.insert("id".to_string(), vml_id_remap);
    let mut reader = XmlReader::from_reader(xml);
    let mut writer = XmlWriter::new(Vec::with_capacity(xml.len()));
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.local_name().as_ref() == "shape" => {
                let Ok(new_e) = rewrite_element_attrs_by_local_name(&e, &attr_maps) else {
                    return xml.to_vec();
                };
                if writer.write_event(Event::Start(new_e)).is_err() {
                    return xml.to_vec();
                }
            }
            Ok(Event::Empty(e)) if e.local_name().as_ref() == "shape" => {
                let Ok(new_e) = rewrite_element_attrs_by_local_name(&e, &attr_maps) else {
                    return xml.to_vec();
                };
                if writer.write_event(Event::Empty(new_e)).is_err() {
                    return xml.to_vec();
                }
            }
            Ok(Event::Eof) => break,
            Ok(other) => {
                if writer.write_event(other).is_err() {
                    return xml.to_vec();
                }
            }
            Err(_) => return xml.to_vec(),
        }
        buf.clear();
    }
    writer.into_inner()
}

fn vml_control_shape_ids(xml: &[u8], id_remap: &HashMap<String, String>) -> HashSet<String> {
    let candidate_ids: HashSet<String> = id_remap
        .keys()
        .map(|old_id| format!("_x0000_s{old_id}"))
        .collect();
    let mut reader = XmlReader::from_reader(xml);
    let mut buf = Vec::new();
    let mut shape_stack: Vec<Option<String>> = Vec::new();
    let mut out = HashSet::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.local_name().as_ref() == "shape" => {
                let candidate = attr_value(&e, b"id").filter(|id| candidate_ids.contains(id));
                shape_stack.push(candidate);
            }
            Ok(Event::Empty(e)) if e.local_name().as_ref() == "ClientData" => {
                if let Some(Some(shape_id)) = shape_stack.last() {
                    if is_vml_control_client_data(&e) {
                        out.insert(shape_id.clone());
                    }
                }
            }
            Ok(Event::Start(e)) if e.local_name().as_ref() == "ClientData" => {
                if let Some(Some(shape_id)) = shape_stack.last() {
                    if is_vml_control_client_data(&e) {
                        out.insert(shape_id.clone());
                    }
                }
            }
            Ok(Event::End(e)) if e.local_name().as_ref() == "shape" => {
                shape_stack.pop();
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    out
}

fn is_vml_control_client_data(e: &BytesStart<'_>) -> bool {
    attr_value(e, b"ObjectType")
        .map(|object_type| object_type != "Note")
        .unwrap_or(false)
}

fn rewrite_vml_control_idmap_data(xml: &[u8], id_remap: &HashMap<String, String>) -> Vec<u8> {
    let Some(idmap_data) = id_remap
        .values()
        .filter_map(|value| value.parse::<u32>().ok())
        .min()
        .map(vml_shape_id_block_data)
    else {
        return xml.to_vec();
    };
    let idmap_data = idmap_data.to_string();
    let mut reader = XmlReader::from_reader(xml);
    let mut writer = XmlWriter::new(Vec::with_capacity(xml.len()));
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if e.local_name().as_ref() == "idmap" => {
                let Ok(new_e) = rewrite_element_attr_value_by_local_name(&e, "data", &idmap_data)
                else {
                    return xml.to_vec();
                };
                if writer.write_event(Event::Start(new_e)).is_err() {
                    return xml.to_vec();
                }
            }
            Ok(Event::Empty(e)) if e.local_name().as_ref() == "idmap" => {
                let Ok(new_e) = rewrite_element_attr_value_by_local_name(&e, "data", &idmap_data)
                else {
                    return xml.to_vec();
                };
                if writer.write_event(Event::Empty(new_e)).is_err() {
                    return xml.to_vec();
                }
            }
            Ok(Event::Eof) => break,
            Ok(other) => {
                if writer.write_event(other).is_err() {
                    return xml.to_vec();
                }
            }
            Err(_) => return xml.to_vec(),
        }
        buf.clear();
    }
    writer.into_inner()
}

fn control_shape_ids_from_sheet(sheet_xml: &[u8]) -> Vec<String> {
    let mut reader = XmlReader::from_reader(sheet_xml);
    let mut buf = Vec::new();
    let mut out = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) if e.local_name().as_ref() == "control" => {
                if let Some(value) = attr_value(&e, b"shapeId") {
                    if value.chars().all(|c| c.is_ascii_digit()) {
                        out.push(value);
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

fn max_shape_id_from_parts(zip_parts: &HashMap<String, Vec<u8>>) -> u32 {
    zip_parts
        .values()
        .flat_map(|bytes| shape_ids_from_xml(bytes))
        .max()
        .unwrap_or(0)
}

fn vml_shape_id_block_data(shape_id: u32) -> u32 {
    shape_id.saturating_sub(1) / 1024
}

fn shape_ids_from_xml(xml: &[u8]) -> Vec<u32> {
    let mut reader = XmlReader::from_reader(xml);
    let mut buf = Vec::new();
    let mut out = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                for attr in e.attributes().with_checks(false).flatten() {
                    let key = attr.key.as_ref();
                    let local_key = key.rsplit(':').next().unwrap_or(key);
                    let value = quick_xml::escape::unescape(&attr.value)
                        .map(|v| v.into_owned())
                        .unwrap_or_else(|_| attr.value.into_owned());
                    match local_key {
                        "shapeId" => {
                            if let Ok(id) = value.parse::<u32>() {
                                out.push(id);
                            }
                        }
                        "id" | "spid" => {
                            if let Some(id) = value
                                .strip_prefix("_x0000_s")
                                .and_then(|suffix| suffix.parse::<u32>().ok())
                            {
                                out.push(id);
                            } else if let Ok(id) = value.parse::<u32>() {
                                out.push(id);
                            }
                        }
                        _ => {}
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

fn rewrite_table_formula_refs(xml: &[u8], table_name_remap: &HashMap<String, String>) -> Vec<u8> {
    if table_name_remap.is_empty() {
        return xml.to_vec();
    }
    let Ok(mut s) = String::from_utf8(xml.to_vec()) else {
        return xml.to_vec();
    };
    for (old_name, new_name) in table_name_remap {
        s = s.replace(&format!("{old_name}["), &format!("{new_name}["));
    }
    s.into_bytes()
}

fn first_attr_on_local_element(src_bytes: &[u8], element: &[u8], attr: &[u8]) -> Option<String> {
    let mut reader = XmlReader::from_reader(src_bytes);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e))
                if e.local_name().as_ref().as_bytes() == element =>
            {
                return attr_value(&e, attr);
            }
            Ok(Event::Eof) => return None,
            Err(_) => return None,
            _ => {}
        }
        buf.clear();
    }
}

fn rewrite_attributes_by_local_name(
    xml: &[u8],
    attr_maps: &HashMap<String, HashMap<String, String>>,
) -> Vec<u8> {
    if attr_maps.values().all(HashMap::is_empty) {
        return xml.to_vec();
    }
    let mut reader = XmlReader::from_reader(xml);
    let mut writer = XmlWriter::new(Vec::with_capacity(xml.len()));
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let Ok(new_e) = rewrite_element_attrs_by_local_name(&e, attr_maps) else {
                    return xml.to_vec();
                };
                if writer.write_event(Event::Start(new_e)).is_err() {
                    return xml.to_vec();
                }
            }
            Ok(Event::Empty(e)) => {
                let Ok(new_e) = rewrite_element_attrs_by_local_name(&e, attr_maps) else {
                    return xml.to_vec();
                };
                if writer.write_event(Event::Empty(new_e)).is_err() {
                    return xml.to_vec();
                }
            }
            Ok(Event::Eof) => break,
            Ok(other) => {
                if writer.write_event(other).is_err() {
                    return xml.to_vec();
                }
            }
            Err(_) => return xml.to_vec(),
        }
        buf.clear();
    }
    writer.into_inner()
}

fn rewrite_element_attrs_by_local_name(
    e: &BytesStart<'_>,
    attr_maps: &HashMap<String, HashMap<String, String>>,
) -> Result<BytesStart<'static>, SheetCopyError> {
    let mut new_e = BytesStart::new(e.name().as_ref().to_owned());
    for a in e.attributes().with_checks(false).flatten() {
        let key = a.key.as_ref();
        let local_key = key.rsplit(':').next().unwrap_or(key);
        let value = quick_xml::escape::unescape(&a.value)
            .map(|v| v.into_owned())
            .unwrap_or_else(|_| a.value.into_owned());
        let new_value = attr_maps
            .get(local_key)
            .and_then(|m| m.get(&value))
            .cloned()
            .unwrap_or(value);
        push_attr_escaped(&mut new_e, key, &new_value);
    }
    Ok(new_e)
}

fn rewrite_element_attr_value_by_local_name(
    e: &BytesStart<'_>,
    attr_name: &str,
    attr_value: &str,
) -> Result<BytesStart<'static>, SheetCopyError> {
    let mut new_e = BytesStart::new(e.name().as_ref().to_owned());
    for a in e.attributes().with_checks(false).flatten() {
        let key = a.key.as_ref();
        let local_key = key.rsplit(':').next().unwrap_or(key);
        if local_key == attr_name {
            push_attr_escaped(&mut new_e, key, attr_value);
        } else {
            let value = quick_xml::escape::unescape(&a.value)
                .map(|v| v.into_owned())
                .unwrap_or_else(|_| a.value.into_owned());
            push_attr_escaped(&mut new_e, key, &value);
        }
    }
    Ok(new_e)
}

// ---------------------------------------------------------------------------
// Helpers — sheet-XML rId remap pass
// ---------------------------------------------------------------------------

/// Apply an `(old_rid → new_rid)` remap to every `r:id="…"` attribute
/// in the source sheet XML. SINGLE pass (RFC-035 §8 risk #4 mitigation)
/// — we walk every element and rewrite any attribute whose local name
/// is `id` AND whose namespace prefix is `r` (or whose qualified name
/// is `r:id`). Other attributes (and element text) flow through
/// verbatim.
fn rewrite_rids(
    src_xml: &[u8],
    remap: &HashMap<String, String>,
) -> Result<Vec<u8>, SheetCopyError> {
    if remap.is_empty() {
        return Ok(src_xml.to_vec());
    }
    let mut reader = XmlReader::from_reader(src_xml);
    let mut writer = XmlWriter::new(Vec::with_capacity(src_xml.len()));
    let mut buf: Vec<u8> = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let new_e = remap_element_attributes(&e, remap)?;
                writer
                    .write_event(Event::Start(new_e))
                    .map_err(|er| SheetCopyError::Xml(format!("write Start: {er}")))?;
            }
            Ok(Event::Empty(e)) => {
                let new_e = remap_element_attributes(&e, remap)?;
                writer
                    .write_event(Event::Empty(new_e))
                    .map_err(|er| SheetCopyError::Xml(format!("write Empty: {er}")))?;
            }
            Ok(Event::Eof) => break,
            Ok(other) => {
                writer
                    .write_event(other)
                    .map_err(|er| SheetCopyError::Xml(format!("write flow: {er}")))?;
            }
            Err(e) => return Err(SheetCopyError::Xml(format!("rid remap parse: {e}"))),
        }
        buf.clear();
    }
    Ok(writer.into_inner())
}

fn remap_element_attributes<'a>(
    e: &BytesStart<'a>,
    remap: &HashMap<String, String>,
) -> Result<BytesStart<'static>, SheetCopyError> {
    let mut new_e = BytesStart::new(e.name().as_ref().to_owned());
    for a in e.attributes().with_checks(false).flatten() {
        let key = a.key.as_ref();
        let value = quick_xml::escape::unescape(&a.value)
            .map(|v| v.into_owned())
            .unwrap_or_else(|_| a.value.into_owned());
        let is_r_id = key == "r:id" || (key.ends_with(":id") && key.starts_with('r'));
        let new_value = if is_r_id {
            remap.get(&value).cloned().unwrap_or(value)
        } else {
            value
        };
        push_attr_escaped(&mut new_e, key, &new_value);
    }
    Ok(new_e)
}

// ---------------------------------------------------------------------------
// Helpers — defined-name / sheet-entry rendering
// ---------------------------------------------------------------------------

fn render_defined_name(name: &str, local_sheet_id: u32, formula: &str) -> Vec<u8> {
    let mut out = String::with_capacity(64 + name.len() + formula.len());
    out.push_str("<definedName name=\"");
    push_xml_attr_escape(&mut out, name);
    out.push_str(&format!("\" localSheetId=\"{local_sheet_id}\">"));
    push_xml_text_escape(&mut out, formula);
    out.push_str("</definedName>");
    out.into_bytes()
}

fn render_sheet_entry(name: &str, sheet_n: u32, rid: &str) -> Vec<u8> {
    let mut out = String::with_capacity(64 + name.len());
    out.push_str("<sheet name=\"");
    push_xml_attr_escape(&mut out, name);
    out.push_str(&format!("\" sheetId=\"{sheet_n}\" r:id=\"{rid}\"/>"));
    out.into_bytes()
}

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

fn push_xml_text_escape(out: &mut String, s: &str) {
    for ch in s.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            _ => out.push(ch),
        }
    }
}

// ---------------------------------------------------------------------------
// Helpers — relative path resolution (mirrors subgraph_walk's helper)
// ---------------------------------------------------------------------------

fn parent_dir(part_path: &str) -> String {
    match part_path.rfind('/') {
        Some(idx) => part_path[..idx].to_string(),
        None => String::new(),
    }
}

fn resolve_relative(base_dir: String, target: &str) -> String {
    if let Some(stripped) = target.strip_prefix('/') {
        return stripped.to_string();
    }
    let mut segments: Vec<&str> = if base_dir.is_empty() {
        Vec::new()
    } else {
        base_dir.split('/').collect()
    };
    for part in target.split('/') {
        match part {
            "" | "." => {}
            ".." => {
                segments.pop();
            }
            other => segments.push(other),
        }
    }
    segments.join("/")
}

// ---------------------------------------------------------------------------
// Chart deep-clone helpers (Sprint Μ Pod-γ / RFC-046 §7)
// ---------------------------------------------------------------------------

/// RFC-057 / RFC-035: walk a worksheet XML body, translate every
/// `<f>` formula body and `<f t="array" ref="..." />` attribute so
/// any explicit reference to the source sheet name is re-pointed at
/// the destination sheet.
///
/// For RFC-057 specifically, the `ref` attribute on
/// `<f t="array" ref="...">` and `<f t="dataTable" ref="..."/>` is
/// sheet-local A1 — it never carries an explicit sheet qualifier, so
/// in practice the translator is a no-op for that attribute.  We run
/// it through anyway so future fixtures that DO ship a qualified
/// `ref` (rare but legal) round-trip cleanly.
///
/// Done via SAX (quick-xml) — never regex — so we respect quoting,
/// 3-D refs, and embedded `<` characters in formula text.
pub(crate) fn rewrite_sheet_formula_refs(
    sheet_xml: &[u8],
    src_title: &str,
    dst_title: &str,
) -> Vec<u8> {
    if src_title == dst_title {
        return sheet_xml.to_vec();
    }
    let mut reader = XmlReader::from_reader(sheet_xml);
    let mut writer = XmlWriter::new(Vec::with_capacity(sheet_xml.len()));
    let mut in_f = false;
    let mut buf: Vec<u8> = Vec::new();
    // A formula is one logical node even when XML entity references split it.
    let mut formula = TextNode::default();
    loop {
        let event = reader.read_event_into(&mut buf);
        if !formula.is_empty() && !matches!(event, Ok(Event::Text(_)) | Ok(Event::GeneralRef(_))) {
            if let Ok(text) = formula.unescaped() {
                // OOXML formula bodies omit the tokenizer's leading `=`.
                let prefixed = format!("={text}");
                let translated = wolfxl_formula::rename_sheet(&prefixed, src_title, dst_title);
                writer
                    .write_event(Event::Text(BytesText::new(
                        translated.strip_prefix('=').unwrap_or(&translated),
                    )))
                    .expect("write");
            } else {
                writer
                    .write_event(Event::Text(BytesText::from_escaped(formula.raw())))
                    .expect("write");
            }
            formula.clear();
        }
        match event {
            Ok(Event::Start(e)) => {
                if e.local_name().as_ref() == "f" {
                    in_f = true;
                    let new_e = rewrite_f_attributes(&e, src_title, dst_title);
                    writer.write_event(Event::Start(new_e)).expect("write");
                } else {
                    writer
                        .write_event(Event::Start(e.into_owned()))
                        .expect("write");
                }
            }
            Ok(Event::End(e)) => {
                if e.local_name().as_ref() == "f" {
                    in_f = false;
                }
                writer
                    .write_event(Event::End(e.into_owned()))
                    .expect("write");
            }
            Ok(Event::Empty(e)) => {
                if e.local_name().as_ref() == "f" {
                    let new_e = rewrite_f_attributes(&e, src_title, dst_title);
                    writer.write_event(Event::Empty(new_e)).expect("write");
                } else {
                    writer
                        .write_event(Event::Empty(e.into_owned()))
                        .expect("write");
                }
            }
            Ok(Event::Text(t)) if in_f => formula.push_text(&t),
            Ok(Event::GeneralRef(r)) if in_f => formula.push_ref(&r),
            Ok(Event::Text(t)) => {
                writer.write_event(Event::Text(t)).expect("write");
            }
            Ok(Event::GeneralRef(r)) => {
                writer.write_event(Event::GeneralRef(r)).expect("write");
            }
            Ok(Event::CData(c)) => {
                writer
                    .write_event(Event::CData(c.into_owned()))
                    .expect("write");
            }
            Ok(Event::Decl(d)) => {
                writer
                    .write_event(Event::Decl(d.into_owned()))
                    .expect("write");
            }
            Ok(Event::PI(p)) => {
                writer
                    .write_event(Event::PI(p.into_owned()))
                    .expect("write");
            }
            Ok(Event::Comment(c)) => {
                writer
                    .write_event(Event::Comment(c.into_owned()))
                    .expect("write");
            }
            Ok(Event::DocType(d)) => {
                writer
                    .write_event(Event::DocType(d.into_owned()))
                    .expect("write");
            }
            Ok(Event::Eof) => break,
            Err(_) => return sheet_xml.to_vec(),
        }
        buf.clear();
    }
    writer.into_inner()
}

/// RFC-057: translate the `ref`, `r1`, `r2` attributes on `<f>` so
/// any explicit-sheet refs they may carry get re-pointed.  Other
/// attributes pass through verbatim.
fn rewrite_f_attributes(
    e: &BytesStart<'_>,
    src_title: &str,
    dst_title: &str,
) -> BytesStart<'static> {
    let mut new_e = BytesStart::new(e.name().as_ref().to_owned());
    for a in e.attributes().with_checks(false).flatten() {
        let key = a.key.as_ref();
        let value = quick_xml::escape::unescape(&a.value)
            .map(|v| v.into_owned())
            .unwrap_or_else(|_| a.value.into_owned());
        let new_value = if matches!(key, "ref" | "r1" | "r2") {
            // Wrap as `=value` to feed the tokenizer, strip back.
            let prefixed = format!("={value}");
            let translated_eq = wolfxl_formula::rename_sheet(&prefixed, src_title, dst_title);
            translated_eq
                .strip_prefix('=')
                .unwrap_or(&translated_eq)
                .to_string()
        } else {
            value
        };
        push_attr_escaped(&mut new_e, key, &new_value);
    }
    new_e
}

/// Walk a chart XML body, find every `<c:f>...</c:f>` formula text
/// element, and apply
/// [`wolfxl_formula::rename_sheet`](rename_sheet) so any cell range
/// pointing at `src_title` is re-pointed at `dst_title`.
///
/// Done via SAX (quick-xml) — never regex — so we respect quoting,
/// 3-D refs, and embedded `</c:f>` characters in unlikely (but
/// legal) text.
///
/// Cell ranges that point at OTHER sheets (cross-sheet refs that
/// don't match the source title) are left untouched. Unknown / bad
/// formulas pass through unchanged (the formula translator already
/// returns the input on parse error).
pub(crate) fn rewrite_chart_sheet_refs(
    chart_xml: &[u8],
    src_title: &str,
    dst_title: &str,
) -> Vec<u8> {
    if src_title == dst_title {
        return chart_xml.to_vec();
    }
    let mut reader = XmlReader::from_reader(chart_xml);
    let mut writer = XmlWriter::new(Vec::with_capacity(chart_xml.len()));
    let mut in_f = false;
    let mut buf: Vec<u8> = Vec::new();
    let mut formula = TextNode::default();
    loop {
        let event = reader.read_event_into(&mut buf);
        if !formula.is_empty() && !matches!(event, Ok(Event::Text(_)) | Ok(Event::GeneralRef(_))) {
            if let Ok(text) = formula.unescaped() {
                // Resolve the entire chart reference before renaming its sheet.
                let prefixed = format!("={text}");
                let translated = wolfxl_formula::rename_sheet(&prefixed, src_title, dst_title);
                writer
                    .write_event(Event::Text(BytesText::new(
                        translated.strip_prefix('=').unwrap_or(&translated),
                    )))
                    .expect("write");
            } else {
                writer
                    .write_event(Event::Text(BytesText::from_escaped(formula.raw())))
                    .expect("write");
            }
            formula.clear();
        }
        match event {
            Ok(Event::Start(e)) => {
                let local = e.local_name();
                if local.as_ref() == "f" {
                    in_f = true;
                }
                writer
                    .write_event(Event::Start(e.into_owned()))
                    .expect("write");
            }
            Ok(Event::End(e)) => {
                let local = e.local_name();
                if local.as_ref() == "f" {
                    in_f = false;
                }
                writer
                    .write_event(Event::End(e.into_owned()))
                    .expect("write");
            }
            Ok(Event::Empty(e)) => {
                writer
                    .write_event(Event::Empty(e.into_owned()))
                    .expect("write");
            }
            Ok(Event::Text(t)) if in_f => formula.push_text(&t),
            Ok(Event::GeneralRef(r)) if in_f => formula.push_ref(&r),
            Ok(Event::Text(t)) => {
                writer.write_event(Event::Text(t)).expect("write");
            }
            Ok(Event::GeneralRef(r)) => {
                writer.write_event(Event::GeneralRef(r)).expect("write");
            }
            Ok(Event::CData(c)) => {
                writer
                    .write_event(Event::CData(c.into_owned()))
                    .expect("write");
            }
            Ok(Event::Decl(d)) => {
                writer
                    .write_event(Event::Decl(d.into_owned()))
                    .expect("write");
            }
            Ok(Event::PI(p)) => {
                writer
                    .write_event(Event::PI(p.into_owned()))
                    .expect("write");
            }
            Ok(Event::Comment(c)) => {
                writer
                    .write_event(Event::Comment(c.into_owned()))
                    .expect("write");
            }
            Ok(Event::DocType(d)) => {
                writer
                    .write_event(Event::DocType(d.into_owned()))
                    .expect("write");
            }
            Ok(Event::Eof) => break,
            Err(_) => {
                // On parse error, return input verbatim — same
                // pass-through contract as the formula translator.
                return chart_xml.to_vec();
            }
        }
        buf.clear();
    }
    writer.into_inner()
}

/// Walk a drawing XML body and rewrite every `<c:chart r:id="...">`
/// reference using `chart_rid_remap`. Used after the drawing's
/// nested rels file is cloned with fresh chart rIds — the drawing
/// XML's `r:id` tokens must match those new rIds.
pub(crate) fn rewrite_drawing_chart_rids(
    drawing_xml: &[u8],
    chart_rid_remap: &HashMap<String, String>,
) -> Vec<u8> {
    if chart_rid_remap.is_empty() {
        return drawing_xml.to_vec();
    }
    let mut reader = XmlReader::from_reader(drawing_xml);
    let mut writer = XmlWriter::new(Vec::with_capacity(drawing_xml.len()));
    let mut buf: Vec<u8> = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let local = e.local_name();
                if local.as_ref() == "chart" {
                    let new_e = remap_chart_attrs(&e, chart_rid_remap);
                    writer.write_event(Event::Start(new_e)).expect("write");
                } else {
                    writer
                        .write_event(Event::Start(e.into_owned()))
                        .expect("write");
                }
            }
            Ok(Event::Empty(e)) => {
                let local = e.local_name();
                if local.as_ref() == "chart" {
                    let new_e = remap_chart_attrs(&e, chart_rid_remap);
                    writer.write_event(Event::Empty(new_e)).expect("write");
                } else {
                    writer
                        .write_event(Event::Empty(e.into_owned()))
                        .expect("write");
                }
            }
            Ok(Event::End(e)) => {
                writer
                    .write_event(Event::End(e.into_owned()))
                    .expect("write");
            }
            Ok(Event::Text(t)) => {
                writer
                    .write_event(Event::Text(t.into_owned()))
                    .expect("write");
            }
            Ok(Event::GeneralRef(r)) => {
                writer.write_event(Event::GeneralRef(r)).expect("write");
            }
            Ok(Event::CData(c)) => {
                writer
                    .write_event(Event::CData(c.into_owned()))
                    .expect("write");
            }
            Ok(Event::Decl(d)) => {
                writer
                    .write_event(Event::Decl(d.into_owned()))
                    .expect("write");
            }
            Ok(Event::PI(p)) => {
                writer
                    .write_event(Event::PI(p.into_owned()))
                    .expect("write");
            }
            Ok(Event::Comment(c)) => {
                writer
                    .write_event(Event::Comment(c.into_owned()))
                    .expect("write");
            }
            Ok(Event::DocType(d)) => {
                writer
                    .write_event(Event::DocType(d.into_owned()))
                    .expect("write");
            }
            Ok(Event::Eof) => break,
            Err(_) => {
                return drawing_xml.to_vec();
            }
        }
        buf.clear();
    }
    writer.into_inner()
}

fn remap_chart_attrs<'a>(
    e: &BytesStart<'a>,
    remap: &HashMap<String, String>,
) -> BytesStart<'static> {
    let name_str = e.name().as_ref().to_owned();
    let mut new_e = BytesStart::new(name_str);
    for a in e.attributes().with_checks(false).flatten() {
        let key = a.key.as_ref();
        let value = quick_xml::escape::unescape(&a.value)
            .map(|v| v.into_owned())
            .unwrap_or_else(|_| a.value.into_owned());
        let is_r_id = key == "r:id" || (key.ends_with(":id") && key.starts_with('r'));
        let new_value = if is_r_id {
            remap.get(&value).cloned().unwrap_or(value)
        } else {
            value
        };
        push_attr_escaped(&mut new_e, key, &new_value);
    }
    new_e
}

fn push_attr_escaped(e: &mut BytesStart<'_>, key: &str, value: &str) {
    e.push_attribute((key, value));
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn defined_name_formulas_resolve_split_references() {
        let xml = br#"<workbook><definedNames><definedName name="empty"/><definedName name="range">'A&amp;B'!A&#49;&amp;&quot;A1&quot;</definedName><definedName name="entity">&#65;</definedName></definedNames></workbook>"#;
        let meta = parse_workbook_metadata(xml).unwrap();
        assert_eq!(meta.defined_names.len(), 3);
        assert_eq!(meta.defined_names[0].formula, "");
        assert_eq!(meta.defined_names[1].formula, "'A&B'!A1&\"A1\"");
        assert_eq!(meta.defined_names[2].formula, "A");
        let bad = br#"<workbook><definedNames><definedName name="bad">&unknown;</definedName></definedNames></workbook>"#;
        assert!(matches!(
            parse_workbook_metadata(bad),
            Err(SheetCopyError::Xml(_))
        ));
    }

    #[test]
    fn sheet_and_chart_renames_resolve_complete_formula_nodes() {
        let sheet = br#"<worksheet><sheetData><row r="1"><c r="A1"><f>'A&amp;B'!A&#49;&amp;&quot;A&amp;B!A1&quot;</f><is><t> &#65;&amp;&lt; </t></is></c></row></sheetData></worksheet>"#;
        let out = String::from_utf8(rewrite_sheet_formula_refs(sheet, "A&B", "C&D")).unwrap();
        assert!(
            out.contains("<f>&apos;C&amp;D&apos;!A1&amp;&quot;A&amp;B!A1&quot;</f>"),
            "{out}"
        );
        assert!(out.contains("<t> &#65;&amp;&lt; </t>"), "{out}");

        let chart = br#"<c:chart xmlns:c="urn:chart"><c:f>'A&amp;B'!$A$&#49;:$A$2</c:f><c:v> &#65;&amp;&lt; </c:v></c:chart>"#;
        let out = String::from_utf8(rewrite_chart_sheet_refs(chart, "A&B", "C&D")).unwrap();
        assert!(
            out.contains("<c:f>&apos;C&amp;D&apos;!$A$1:$A$2</c:f>"),
            "{out}"
        );
        assert!(out.contains("<c:v> &#65;&amp;&lt; </c:v>"), "{out}");
    }

    #[test]
    fn unresolvable_formula_references_remain_raw() {
        let xml = br#"<root><f>Template!A1&amp;&unknown;</f></root>"#;
        assert_eq!(rewrite_sheet_formula_refs(xml, "Template", "Copy"), xml);
        assert_eq!(rewrite_chart_sheet_refs(xml, "Template", "Copy"), xml);
    }

    #[test]
    fn drawing_rid_rewrite_preserves_text_and_escapes_attributes_once() {
        let xml =
            br#"<drawing><chart r:id="old" note="A&amp;B"/><t> &#65;&amp;&lt; </t></drawing>"#;
        let remap = HashMap::from([("old".to_string(), "new&rid".to_string())]);
        let out = String::from_utf8(rewrite_drawing_chart_rids(xml, &remap)).unwrap();
        assert!(
            out.contains(r#"r:id="new&amp;rid" note="A&amp;B""#),
            "{out}"
        );
        assert!(out.contains("<t> &#65;&amp;&lt; </t>"), "{out}");
    }

    fn minimal_sheet_xml() -> Vec<u8> {
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/></worksheet>"#.to_vec()
    }

    fn sheet_xml_with_table_part(rid: &str) -> Vec<u8> {
        format!(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/><tableParts count="1"><tablePart r:id="{rid}"/></tableParts></worksheet>"#).into_bytes()
    }

    fn workbook_xml(titles: &[&str], defined_names: &[(&str, Option<u32>, &str)]) -> Vec<u8> {
        let mut s = String::from(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>"#,
        );
        for (i, t) in titles.iter().enumerate() {
            s.push_str(&format!(
                r#"<sheet name="{}" sheetId="{}" r:id="rId{}"/>"#,
                t,
                i + 1,
                i + 1
            ));
        }
        s.push_str("</sheets>");
        if !defined_names.is_empty() {
            s.push_str("<definedNames>");
            for (n, lsid, f) in defined_names {
                match lsid {
                    Some(id) => s.push_str(&format!(
                        r#"<definedName name="{n}" localSheetId="{id}">{f}</definedName>"#
                    )),
                    None => s.push_str(&format!(r#"<definedName name="{n}">{f}</definedName>"#)),
                }
            }
            s.push_str("</definedNames>");
        }
        s.push_str("</workbook>");
        s.into_bytes()
    }

    fn rels_with(entries: &[(&str, &str, TargetMode)]) -> RelsGraph {
        let mut g = RelsGraph::new();
        for (rt_str, target, mode) in entries {
            g.add(rt_str, target, *mode);
        }
        g
    }

    fn table_part_bytes(name: &str, id: u32) -> Vec<u8> {
        format!(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="{id}" name="{name}" displayName="{name}" ref="A1:B2"><tableColumns count="2"><tableColumn id="1" name="A"/><tableColumn id="2" name="B"/></tableColumns></table>"#).into_bytes()
    }

    fn one_sheet_workbook(
        extra_titles: &[&str],
        defined_names: &[(&str, Option<u32>, &str)],
    ) -> Vec<u8> {
        let mut titles = vec!["Template"];
        titles.extend_from_slice(extra_titles);
        workbook_xml(&titles, defined_names)
    }

    #[test]
    fn basic_clone_no_rels() {
        let mut alloc =
            PartIdAllocator::from_zip_parts(["xl/worksheets/sheet1.xml"].iter().copied());
        let mut zip_parts: HashMap<String, Vec<u8>> = HashMap::new();
        zip_parts.insert("xl/worksheets/sheet1.xml".into(), minimal_sheet_xml());
        let workbook_bytes = one_sheet_workbook(&[], &[]);
        let existing_table_names: HashSet<String> = HashSet::new();
        let source_rels = RelsGraph::new();

        let mutations = plan_sheet_copy(SheetCopyInputs {
            src_title: "Template".into(),
            dst_title: "Template Copy".into(),
            src_sheet_path: "xl/worksheets/sheet1.xml".into(),
            source_zip_parts: &zip_parts,
            source_rels: &source_rels,
            workbook_xml: &workbook_bytes,
            allocator: &mut alloc,
            existing_table_names: &existing_table_names,
            deep_copy_images: false,
        })
        .expect("plan ok");

        assert_eq!(mutations.new_sheet_path, "xl/worksheets/sheet2.xml");
        // No ancillary parts (only the dest rels file, but it's empty here so
        // no rels file is emitted).
        assert!(mutations.new_ancillary_parts.is_empty());
        assert!(mutations.defined_names_to_add.is_empty());
        assert!(mutations.new_table_names.is_empty());
        // CT override for the worksheet itself.
        assert_eq!(
            mutations.content_type_overrides_to_add[0].0,
            "/xl/worksheets/sheet2.xml"
        );
        assert_eq!(mutations.content_type_overrides_to_add[0].1, CT_WORKSHEET);
        // Workbook rels has the new worksheet entry.
        assert_eq!(mutations.workbook_rels_to_add.len(), 1);
        assert_eq!(mutations.workbook_rels_to_add[0].1, rt::WORKSHEET);
        // The placeholder rId appears verbatim in workbook_sheets_append.
        let entry = std::str::from_utf8(&mutations.workbook_sheets_append).unwrap();
        assert!(entry.contains(r#"name="Template Copy""#), "{entry}");
        assert!(entry.contains("__SHEET_RID_PLACEHOLDER_"), "{entry}");
    }

    #[test]
    fn clone_with_one_table_dedups_name() {
        let mut alloc = PartIdAllocator::from_zip_parts(
            ["xl/worksheets/sheet1.xml", "xl/tables/table1.xml"]
                .iter()
                .copied(),
        );
        let mut zip_parts: HashMap<String, Vec<u8>> = HashMap::new();
        zip_parts.insert(
            "xl/worksheets/sheet1.xml".into(),
            sheet_xml_with_table_part("rId1"),
        );
        zip_parts.insert("xl/tables/table1.xml".into(), table_part_bytes("Sales", 1));
        let workbook_bytes = one_sheet_workbook(&[], &[]);
        let mut existing_table_names: HashSet<String> = HashSet::new();
        existing_table_names.insert("Sales".into());
        let source_rels = rels_with(&[(rt::TABLE, "../tables/table1.xml", TargetMode::Internal)]);

        let mutations = plan_sheet_copy(SheetCopyInputs {
            src_title: "Template".into(),
            dst_title: "Template Copy".into(),
            src_sheet_path: "xl/worksheets/sheet1.xml".into(),
            source_zip_parts: &zip_parts,
            source_rels: &source_rels,
            workbook_xml: &workbook_bytes,
            allocator: &mut alloc,
            existing_table_names: &existing_table_names,
            deep_copy_images: false,
        })
        .expect("plan ok");

        // One new table part allocated (table2.xml since seed = max+1 = 2).
        let table_part = mutations
            .new_ancillary_parts
            .iter()
            .find(|(p, _)| p == "xl/tables/table2.xml")
            .expect("table2.xml emitted");
        let body = std::str::from_utf8(&table_part.1).unwrap();
        assert!(body.contains(r#"name="Sales_2""#), "{body}");
        assert!(body.contains(r#"displayName="Sales_2""#), "{body}");
        assert_eq!(mutations.new_table_names, vec!["Sales_2".to_string()]);

        // The cloned sheet XML's r:id should now reference the new rId.
        let new_sheet = std::str::from_utf8(&mutations.new_sheet_xml).unwrap();
        let new_rid = mutations.rid_remap.get("rId1").expect("rId1 remapped");
        assert!(
            new_sheet.contains(&format!(r#"r:id="{new_rid}""#)),
            "{new_sheet}"
        );
    }

    #[test]
    fn clone_with_table_preserves_autofilter_child() {
        let mut alloc = PartIdAllocator::from_zip_parts(
            ["xl/worksheets/sheet1.xml", "xl/tables/table1.xml"]
                .iter()
                .copied(),
        );
        let mut zip_parts: HashMap<String, Vec<u8>> = HashMap::new();
        zip_parts.insert(
            "xl/worksheets/sheet1.xml".into(),
            sheet_xml_with_table_part("rId1"),
        );
        zip_parts.insert(
            "xl/tables/table1.xml".into(),
            br#"<?xml version="1.0" encoding="UTF-8"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="SalesTable" displayName="SalesTable" ref="A1:E31"><autoFilter ref="A1:E31"></autoFilter><tableColumns count="5"><tableColumn id="1" name="Month"></tableColumn><tableColumn id="2" name="Year"></tableColumn><tableColumn id="3" name="Type"></tableColumn><tableColumn id="4" name="Sales"></tableColumn><tableColumn id="5" name="Region"></tableColumn></tableColumns></table>"#.to_vec(),
        );
        let workbook_bytes = one_sheet_workbook(&[], &[]);
        let mut existing_table_names: HashSet<String> = HashSet::new();
        existing_table_names.insert("SalesTable".into());
        let source_rels = rels_with(&[(rt::TABLE, "../tables/table1.xml", TargetMode::Internal)]);

        let mutations = plan_sheet_copy(SheetCopyInputs {
            src_title: "Template".into(),
            dst_title: "Template Copy".into(),
            src_sheet_path: "xl/worksheets/sheet1.xml".into(),
            source_zip_parts: &zip_parts,
            source_rels: &source_rels,
            workbook_xml: &workbook_bytes,
            allocator: &mut alloc,
            existing_table_names: &existing_table_names,
            deep_copy_images: false,
        })
        .expect("plan ok");

        let table_part = mutations
            .new_ancillary_parts
            .iter()
            .find(|(p, _)| p == "xl/tables/table2.xml")
            .expect("table2.xml emitted");
        let body = std::str::from_utf8(&table_part.1).unwrap();
        assert!(
            body.contains(r#"<autoFilter ref="A1:E31"></autoFilter><tableColumns"#),
            "{body}"
        );
    }

    #[test]
    fn clone_with_n_tables_dedups_each() {
        let mut alloc = PartIdAllocator::new();
        let mut zip_parts: HashMap<String, Vec<u8>> = HashMap::new();
        zip_parts.insert(
            "xl/worksheets/sheet1.xml".into(),
            br#"<?xml version="1.0"?><worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><tableParts count="2"><tablePart r:id="rId1"/><tablePart r:id="rId2"/></tableParts></worksheet>"#.to_vec(),
        );
        zip_parts.insert("xl/tables/table1.xml".into(), table_part_bytes("Sales", 1));
        zip_parts.insert("xl/tables/table2.xml".into(), table_part_bytes("Costs", 2));
        let workbook_bytes = one_sheet_workbook(&[], &[]);
        let mut existing_table_names: HashSet<String> = HashSet::new();
        existing_table_names.insert("Sales".into());
        existing_table_names.insert("Costs".into());
        let source_rels = rels_with(&[
            (rt::TABLE, "../tables/table1.xml", TargetMode::Internal),
            (rt::TABLE, "../tables/table2.xml", TargetMode::Internal),
        ]);

        // Pre-seed the allocator past existing files.
        for p in zip_parts.keys() {
            alloc.observe(p);
        }

        let mutations = plan_sheet_copy(SheetCopyInputs {
            src_title: "Template".into(),
            dst_title: "Copy".into(),
            src_sheet_path: "xl/worksheets/sheet1.xml".into(),
            source_zip_parts: &zip_parts,
            source_rels: &source_rels,
            workbook_xml: &workbook_bytes,
            allocator: &mut alloc,
            existing_table_names: &existing_table_names,
            deep_copy_images: false,
        })
        .expect("plan ok");

        assert_eq!(
            mutations.new_table_names,
            vec!["Sales_2".to_string(), "Costs_2".to_string()]
        );

        // Both tables emitted with fresh suffixes.
        let table_paths: Vec<&str> = mutations
            .new_ancillary_parts
            .iter()
            .map(|(p, _)| p.as_str())
            .filter(|p| p.starts_with("xl/tables/"))
            .collect();
        assert!(
            table_paths.contains(&"xl/tables/table3.xml"),
            "{table_paths:?}"
        );
        assert!(
            table_paths.contains(&"xl/tables/table4.xml"),
            "{table_paths:?}"
        );
    }

    #[test]
    fn clone_with_external_hyperlink_alias_url() {
        let mut alloc = PartIdAllocator::new();
        let mut zip_parts: HashMap<String, Vec<u8>> = HashMap::new();
        zip_parts.insert("xl/worksheets/sheet1.xml".into(), minimal_sheet_xml());
        let workbook_bytes = one_sheet_workbook(&[], &[]);
        let existing_table_names: HashSet<String> = HashSet::new();
        let source_rels = rels_with(&[
            (rt::HYPERLINK, "https://example.com/x", TargetMode::External),
            (rt::HYPERLINK, "mailto:a@b.com", TargetMode::External),
        ]);

        let mutations = plan_sheet_copy(SheetCopyInputs {
            src_title: "Template".into(),
            dst_title: "T2".into(),
            src_sheet_path: "xl/worksheets/sheet1.xml".into(),
            source_zip_parts: &zip_parts,
            source_rels: &source_rels,
            workbook_xml: &workbook_bytes,
            allocator: &mut alloc,
            existing_table_names: &existing_table_names,
            deep_copy_images: false,
        })
        .expect("plan ok");

        // Both hyperlinks remap to fresh rIds in the new dest rels graph.
        assert_eq!(mutations.rid_remap.len(), 2);
        // The destination rels file is emitted (non-empty).
        let rels_part = mutations
            .new_ancillary_parts
            .iter()
            .find(|(p, _)| p.ends_with(".rels"))
            .expect("rels file emitted");
        let rels_body = std::str::from_utf8(&rels_part.1).unwrap();
        assert!(rels_body.contains("https://example.com/x"), "{rels_body}");
        assert!(rels_body.contains("mailto:a@b.com"), "{rels_body}");
    }

    #[test]
    fn clone_escapes_internal_hyperlink_location_after_rid_rewrite() {
        let mut alloc =
            PartIdAllocator::from_zip_parts(["xl/worksheets/sheet1.xml"].iter().copied());
        let mut zip_parts: HashMap<String, Vec<u8>> = HashMap::new();
        zip_parts.insert(
            "xl/worksheets/sheet1.xml".into(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData><row r="1"><c r="A1"><f>Template!A1</f></c></row></sheetData>
  <hyperlinks><hyperlink ref="B2" r:id="rId1" location="'Sec. 1 &amp; 2 Notes'!A1" display="A &amp; B"/></hyperlinks>
</worksheet>"#
                .to_vec(),
        );
        let workbook_bytes = one_sheet_workbook(&[], &[]);
        let existing_table_names: HashSet<String> = HashSet::new();
        let source_rels =
            rels_with(&[(rt::HYPERLINK, "https://example.com/x", TargetMode::External)]);

        let mutations = plan_sheet_copy(SheetCopyInputs {
            src_title: "Template".into(),
            dst_title: "Template Copy".into(),
            src_sheet_path: "xl/worksheets/sheet1.xml".into(),
            source_zip_parts: &zip_parts,
            source_rels: &source_rels,
            workbook_xml: &workbook_bytes,
            allocator: &mut alloc,
            existing_table_names: &existing_table_names,
            deep_copy_images: false,
        })
        .expect("plan ok");

        let new_sheet = std::str::from_utf8(&mutations.new_sheet_xml).unwrap();
        assert!(
            new_sheet.contains(r#"location="&apos;Sec. 1 &amp; 2 Notes&apos;!A1""#),
            "{new_sheet}"
        );
        assert!(new_sheet.contains(r#"display="A &amp; B""#), "{new_sheet}");
        assert!(
            new_sheet.contains("<f>&apos;Template Copy&apos;!A1</f>"),
            "{new_sheet}"
        );

        let mut reader = XmlReader::from_reader(mutations.new_sheet_xml.as_slice());
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Eof) => break,
                Ok(_) => {}
                Err(err) => panic!("cloned sheet XML is malformed: {err}; {new_sheet}"),
            }
            buf.clear();
        }
    }

    #[test]
    fn clone_with_comments_and_vml() {
        let mut alloc = PartIdAllocator::new();
        let mut zip_parts: HashMap<String, Vec<u8>> = HashMap::new();
        zip_parts.insert(
            "xl/worksheets/sheet1.xml".into(),
            br#"<?xml version="1.0"?><worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><legacyDrawing r:id="rId2"/></worksheet>"#.to_vec(),
        );
        zip_parts.insert("xl/comments1.xml".into(), b"<comments/>".to_vec());
        zip_parts.insert("xl/drawings/vmlDrawing1.vml".into(), b"<xml/>".to_vec());
        let workbook_bytes = one_sheet_workbook(&[], &[]);
        let existing_table_names: HashSet<String> = HashSet::new();
        let source_rels = rels_with(&[
            (rt::COMMENTS, "../comments1.xml", TargetMode::Internal),
            (
                rt::VML_DRAWING,
                "../drawings/vmlDrawing1.vml",
                TargetMode::Internal,
            ),
        ]);
        for p in zip_parts.keys() {
            alloc.observe(p);
        }

        let mutations = plan_sheet_copy(SheetCopyInputs {
            src_title: "Template".into(),
            dst_title: "T2".into(),
            src_sheet_path: "xl/worksheets/sheet1.xml".into(),
            source_zip_parts: &zip_parts,
            source_rels: &source_rels,
            workbook_xml: &workbook_bytes,
            allocator: &mut alloc,
            existing_table_names: &existing_table_names,
            deep_copy_images: false,
        })
        .expect("plan ok");

        let paths: Vec<&str> = mutations
            .new_ancillary_parts
            .iter()
            .map(|(p, _)| p.as_str())
            .collect();
        assert!(paths.contains(&"xl/comments2.xml"), "{paths:?}");
        assert!(paths.contains(&"xl/drawings/vmlDrawing2.vml"), "{paths:?}");
    }

    #[test]
    fn clone_with_form_control_properties_allocates_fresh_ctrl_prop() {
        let mut alloc = PartIdAllocator::new();
        let mut zip_parts: HashMap<String, Vec<u8>> = HashMap::new();
        zip_parts.insert(
            "xl/worksheets/sheet1.xml".into(),
            br#"<?xml version="1.0"?><worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><controls><control shapeId="1025" r:id="rId1"/></controls><legacyDrawing r:id="rId2"/><drawing r:id="rId3"/></worksheet>"#.to_vec(),
        );
        zip_parts.insert(
            "xl/ctrlProps/ctrlProp1.xml".into(),
            b"<formControlPr/>".to_vec(),
        );
        zip_parts.insert(
            "xl/drawings/vmlDrawing1.vml".into(),
            br#"<xml><o:shapelayout xmlns:o="urn:schemas-microsoft-com:office:office"><o:idmap data="1"/></o:shapelayout><v:shape xmlns:v="urn:schemas-microsoft-com:vml" id="_x0000_s1025"><x:ClientData xmlns:x="urn:schemas-microsoft-com:office:excel" ObjectType="List"/></v:shape></xml>"#
                .to_vec(),
        );
        zip_parts.insert(
            "xl/drawings/drawing1.xml".into(),
            br#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"><xdr:cNvPr id="1025"><a14:compatExt spid="_x0000_s1025"/></xdr:cNvPr></xdr:wsDr>"#.to_vec(),
        );
        let workbook_bytes = one_sheet_workbook(&[], &[]);
        let existing_table_names: HashSet<String> = HashSet::new();
        let source_rels = rels_with(&[
            (
                rt::CTRL_PROP,
                "../ctrlProps/ctrlProp1.xml",
                TargetMode::Internal,
            ),
            (
                rt::VML_DRAWING,
                "../drawings/vmlDrawing1.vml",
                TargetMode::Internal,
            ),
            (
                rt::DRAWING,
                "../drawings/drawing1.xml",
                TargetMode::Internal,
            ),
        ]);
        for p in zip_parts.keys() {
            alloc.observe(p);
        }

        let mutations = plan_sheet_copy(SheetCopyInputs {
            src_title: "Template".into(),
            dst_title: "T2".into(),
            src_sheet_path: "xl/worksheets/sheet1.xml".into(),
            source_zip_parts: &zip_parts,
            source_rels: &source_rels,
            workbook_xml: &workbook_bytes,
            allocator: &mut alloc,
            existing_table_names: &existing_table_names,
            deep_copy_images: false,
        })
        .expect("plan ok");

        let paths: Vec<&str> = mutations
            .new_ancillary_parts
            .iter()
            .map(|(p, _)| p.as_str())
            .collect();
        assert!(paths.contains(&"xl/ctrlProps/ctrlProp2.xml"), "{paths:?}");
        assert!(mutations
            .content_type_overrides_to_add
            .contains(&("/xl/ctrlProps/ctrlProp2.xml".into(), CT_CTRL_PROP.into())));
        let dest_rels = mutations
            .new_ancillary_parts
            .iter()
            .find(|(p, _)| p == "xl/worksheets/_rels/sheet2.xml.rels")
            .map(|(_, bytes)| std::str::from_utf8(bytes).unwrap().to_string())
            .expect("dest rels emitted");
        assert!(
            dest_rels.contains("../ctrlProps/ctrlProp2.xml"),
            "{dest_rels}"
        );
        assert!(
            !dest_rels.contains("../ctrlProps/ctrlProp1.xml"),
            "{dest_rels}"
        );
        let sheet = std::str::from_utf8(&mutations.new_sheet_xml).unwrap();
        assert!(sheet.contains(r#"shapeId="2049""#), "{sheet}");
        assert!(!sheet.contains(r#"shapeId="1025""#), "{sheet}");
        let cloned_drawing = mutations
            .new_ancillary_parts
            .iter()
            .find(|(p, _)| p == "xl/drawings/drawing2.xml")
            .map(|(_, bytes)| std::str::from_utf8(bytes).unwrap().to_string())
            .expect("cloned drawing");
        assert!(cloned_drawing.contains(r#"id="2049""#), "{cloned_drawing}");
        assert!(
            cloned_drawing.contains(r#"spid="_x0000_s2049""#),
            "{cloned_drawing}"
        );
        let cloned_vml = mutations
            .new_ancillary_parts
            .iter()
            .find(|(p, _)| p == "xl/drawings/vmlDrawing2.vml")
            .map(|(_, bytes)| std::str::from_utf8(bytes).unwrap().to_string())
            .expect("cloned vml");
        assert!(cloned_vml.contains(r#"id="_x0000_s2049""#), "{cloned_vml}");
        assert!(cloned_vml.contains(r#"data="2""#), "{cloned_vml}");

        let second = plan_sheet_copy(SheetCopyInputs {
            src_title: "Template".into(),
            dst_title: "T3".into(),
            src_sheet_path: "xl/worksheets/sheet1.xml".into(),
            source_zip_parts: &zip_parts,
            source_rels: &source_rels,
            workbook_xml: &workbook_bytes,
            allocator: &mut alloc,
            existing_table_names: &existing_table_names,
            deep_copy_images: false,
        })
        .expect("second plan ok");
        let second_sheet = std::str::from_utf8(&second.new_sheet_xml).unwrap();
        assert!(second_sheet.contains(r#"shapeId="2050""#), "{second_sheet}");
        assert!(
            !second_sheet.contains(r#"shapeId="2049""#),
            "{second_sheet}"
        );
    }

    #[test]
    fn vml_control_shape_rewrite_ignores_comment_notes() {
        let mut remap = HashMap::new();
        remap.insert("1025".to_string(), "2049".to_string());
        let comment_vml = br#"<xml><v:shape xmlns:v="urn:schemas-microsoft-com:vml" id="_x0000_s1025"><x:ClientData xmlns:x="urn:schemas-microsoft-com:office:excel" ObjectType="Note"/></v:shape></xml>"#;
        let control_vml = br#"<xml><o:shapelayout xmlns:o="urn:schemas-microsoft-com:office:office"><o:idmap data="1"/></o:shapelayout><v:shape xmlns:v="urn:schemas-microsoft-com:vml" id="_x0000_s1025"><x:ClientData xmlns:x="urn:schemas-microsoft-com:office:excel" ObjectType="List"/></v:shape></xml>"#;

        let comment_ids = vml_control_shape_ids(comment_vml, &remap);
        assert!(comment_ids.is_empty());
        let comment_out = rewrite_vml_control_shape_ids(comment_vml, &remap, &comment_ids);
        let comment_out = std::str::from_utf8(&comment_out).unwrap();
        assert!(
            comment_out.contains(r#"id="_x0000_s1025""#),
            "{comment_out}"
        );
        assert!(!comment_out.contains("_x0000_s2049"), "{comment_out}");

        let control_ids = vml_control_shape_ids(control_vml, &remap);
        assert!(control_ids.contains("_x0000_s1025"));
        let control_out = rewrite_vml_control_shape_ids(control_vml, &remap, &control_ids);
        let control_out = rewrite_vml_control_idmap_data(&control_out, &remap);
        let control_out = std::str::from_utf8(&control_out).unwrap();
        assert!(
            control_out.contains(r#"id="_x0000_s2049""#),
            "{control_out}"
        );
        assert!(control_out.contains(r#"data="2""#), "{control_out}");
    }

    #[test]
    fn clone_with_drawing_aliases_image() {
        let mut alloc = PartIdAllocator::new();
        let mut zip_parts: HashMap<String, Vec<u8>> = HashMap::new();
        zip_parts.insert("xl/worksheets/sheet1.xml".into(), minimal_sheet_xml());
        zip_parts.insert("xl/drawings/drawing1.xml".into(), b"<wsDr/>".to_vec());
        // Drawing's own rels file → image1.png.
        let drawing_rels_xml =
            rels_with(&[(rt::IMAGE, "../media/image1.png", TargetMode::Internal)]).serialize();
        zip_parts.insert(
            "xl/drawings/_rels/drawing1.xml.rels".into(),
            drawing_rels_xml,
        );
        zip_parts.insert("xl/media/image1.png".into(), b"\x89PNG".to_vec());
        let workbook_bytes = one_sheet_workbook(&[], &[]);
        let existing_table_names: HashSet<String> = HashSet::new();
        let source_rels = rels_with(&[(
            rt::DRAWING,
            "../drawings/drawing1.xml",
            TargetMode::Internal,
        )]);
        for p in zip_parts.keys() {
            alloc.observe(p);
        }

        let mutations = plan_sheet_copy(SheetCopyInputs {
            src_title: "Template".into(),
            dst_title: "T2".into(),
            src_sheet_path: "xl/worksheets/sheet1.xml".into(),
            source_zip_parts: &zip_parts,
            source_rels: &source_rels,
            workbook_xml: &workbook_bytes,
            allocator: &mut alloc,
            existing_table_names: &existing_table_names,
            deep_copy_images: false,
        })
        .expect("plan ok");

        // New drawing emitted at drawing2.xml.
        let drawing_part = mutations
            .new_ancillary_parts
            .iter()
            .find(|(p, _)| p == "xl/drawings/drawing2.xml")
            .expect("drawing emitted");
        assert_eq!(drawing_part.1.as_slice(), b"<wsDr/>");

        // Cloned drawing rels file emitted, with image still pointing at
        // the ORIGINAL image1.png (alias, not deep clone).
        let drawing_rels_part = mutations
            .new_ancillary_parts
            .iter()
            .find(|(p, _)| p == "xl/drawings/_rels/drawing2.xml.rels")
            .expect("drawing rels emitted");
        let body = std::str::from_utf8(&drawing_rels_part.1).unwrap();
        assert!(body.contains("../media/image1.png"), "{body}");

        // The image part is NOT in new_ancillary_parts (aliasing).
        assert!(!mutations
            .new_ancillary_parts
            .iter()
            .any(|(p, _)| p == "xl/media/image1.png"));
    }

    #[test]
    fn clone_with_sheet_scoped_defined_name() {
        let mut alloc = PartIdAllocator::new();
        let mut zip_parts: HashMap<String, Vec<u8>> = HashMap::new();
        zip_parts.insert("xl/worksheets/sheet1.xml".into(), minimal_sheet_xml());
        // src_idx = 0 (Template). new_idx = 2 (after Sheet2 inserts).
        let workbook_bytes = workbook_xml(
            &["Template", "Other"],
            &[
                ("_xlnm.Print_Area", Some(0), "Template!$A$1:$E$10"),
                ("WorkbookScopeName", None, "Other!$A$1"),
                ("OtherSheetName", Some(1), "Other!$B$2"),
            ],
        );
        let existing_table_names: HashSet<String> = HashSet::new();
        let source_rels = RelsGraph::new();

        let mutations = plan_sheet_copy(SheetCopyInputs {
            src_title: "Template".into(),
            dst_title: "Template Copy".into(),
            src_sheet_path: "xl/worksheets/sheet1.xml".into(),
            source_zip_parts: &zip_parts,
            source_rels: &source_rels,
            workbook_xml: &workbook_bytes,
            allocator: &mut alloc,
            existing_table_names: &existing_table_names,
            deep_copy_images: false,
        })
        .expect("plan ok");

        // Only the Template-scoped name was cloned.
        assert_eq!(mutations.defined_names_to_add.len(), 1);
        let dn = &mutations.defined_names_to_add[0];
        assert_eq!(dn.name, "_xlnm.Print_Area");
        // new_idx = sheet_titles.len() = 2.
        assert_eq!(dn.local_sheet_id, 2);
        assert_eq!(dn.formula, "Template!$A$1:$E$10");
        // Element bytes carry the new localSheetId.
        let body = std::str::from_utf8(&dn.element_bytes).unwrap();
        assert!(body.contains(r#"localSheetId="2""#), "{body}");
        assert!(body.contains("_xlnm.Print_Area"), "{body}");
    }

    #[test]
    fn clone_without_sheet_scoped_names_yields_empty_list() {
        let mut alloc = PartIdAllocator::new();
        let mut zip_parts: HashMap<String, Vec<u8>> = HashMap::new();
        zip_parts.insert("xl/worksheets/sheet1.xml".into(), minimal_sheet_xml());
        let workbook_bytes =
            workbook_xml(&["Template", "Other"], &[("WorkbookScopeName", None, "A1")]);
        let existing_table_names: HashSet<String> = HashSet::new();
        let source_rels = RelsGraph::new();

        let mutations = plan_sheet_copy(SheetCopyInputs {
            src_title: "Template".into(),
            dst_title: "Template Copy".into(),
            src_sheet_path: "xl/worksheets/sheet1.xml".into(),
            source_zip_parts: &zip_parts,
            source_rels: &source_rels,
            workbook_xml: &workbook_bytes,
            allocator: &mut alloc,
            existing_table_names: &existing_table_names,
            deep_copy_images: false,
        })
        .expect("plan ok");

        assert!(mutations.defined_names_to_add.is_empty());
    }

    #[test]
    fn validation_duplicate_dst_title_errors() {
        let mut alloc = PartIdAllocator::new();
        let mut zip_parts: HashMap<String, Vec<u8>> = HashMap::new();
        zip_parts.insert("xl/worksheets/sheet1.xml".into(), minimal_sheet_xml());
        let workbook_bytes = workbook_xml(&["Template", "Existing"], &[]);
        let existing_table_names: HashSet<String> = HashSet::new();
        let source_rels = RelsGraph::new();

        let err = plan_sheet_copy(SheetCopyInputs {
            src_title: "Template".into(),
            dst_title: "Existing".into(),
            src_sheet_path: "xl/worksheets/sheet1.xml".into(),
            source_zip_parts: &zip_parts,
            source_rels: &source_rels,
            workbook_xml: &workbook_bytes,
            allocator: &mut alloc,
            existing_table_names: &existing_table_names,
            deep_copy_images: false,
        })
        .expect_err("duplicate dst should error");
        match err {
            SheetCopyError::DuplicateDestinationTitle(s) => assert_eq!(s, "Existing"),
            other => panic!("wrong error: {other:?}"),
        }
    }

    #[test]
    fn validation_missing_source_title_errors() {
        let mut alloc = PartIdAllocator::new();
        let mut zip_parts: HashMap<String, Vec<u8>> = HashMap::new();
        zip_parts.insert("xl/worksheets/sheet1.xml".into(), minimal_sheet_xml());
        let workbook_bytes = workbook_xml(&["Template"], &[]);
        let existing_table_names: HashSet<String> = HashSet::new();
        let source_rels = RelsGraph::new();

        let err = plan_sheet_copy(SheetCopyInputs {
            src_title: "Nonexistent".into(),
            dst_title: "Foo".into(),
            src_sheet_path: "xl/worksheets/sheet1.xml".into(),
            source_zip_parts: &zip_parts,
            source_rels: &source_rels,
            workbook_xml: &workbook_bytes,
            allocator: &mut alloc,
            existing_table_names: &existing_table_names,
            deep_copy_images: false,
        })
        .expect_err("missing src should error");
        assert!(matches!(err, SheetCopyError::MissingSourceTitle(_)));
    }

    #[test]
    fn validation_missing_source_bytes_errors() {
        let mut alloc = PartIdAllocator::new();
        let zip_parts: HashMap<String, Vec<u8>> = HashMap::new(); // empty
        let workbook_bytes = workbook_xml(&["Template"], &[]);
        let existing_table_names: HashSet<String> = HashSet::new();
        let source_rels = RelsGraph::new();

        let err = plan_sheet_copy(SheetCopyInputs {
            src_title: "Template".into(),
            dst_title: "Copy".into(),
            src_sheet_path: "xl/worksheets/sheet1.xml".into(),
            source_zip_parts: &zip_parts,
            source_rels: &source_rels,
            workbook_xml: &workbook_bytes,
            allocator: &mut alloc,
            existing_table_names: &existing_table_names,
            deep_copy_images: false,
        })
        .expect_err("missing src bytes should error");
        assert!(matches!(err, SheetCopyError::MissingSourceSheetBytes(_)));
    }

    #[test]
    fn rid_remap_applied_to_sheet_xml_in_one_pass() {
        // Source sheet has TWO different r:id references; both must be
        // rewritten in a single pass to point at fresh dest rIds.
        let mut alloc = PartIdAllocator::new();
        let mut zip_parts: HashMap<String, Vec<u8>> = HashMap::new();
        let sheet_src = br#"<?xml version="1.0"?><worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><tableParts count="1"><tablePart r:id="rId1"/></tableParts><legacyDrawing r:id="rId2"/></worksheet>"#.to_vec();
        zip_parts.insert("xl/worksheets/sheet1.xml".into(), sheet_src);
        zip_parts.insert("xl/tables/table1.xml".into(), table_part_bytes("Sales", 1));
        zip_parts.insert("xl/drawings/vmlDrawing1.vml".into(), b"<xml/>".to_vec());
        let workbook_bytes = one_sheet_workbook(&[], &[]);
        let existing_table_names: HashSet<String> = HashSet::new();
        let source_rels = rels_with(&[
            (rt::TABLE, "../tables/table1.xml", TargetMode::Internal),
            (
                rt::VML_DRAWING,
                "../drawings/vmlDrawing1.vml",
                TargetMode::Internal,
            ),
        ]);
        for p in zip_parts.keys() {
            alloc.observe(p);
        }

        let mutations = plan_sheet_copy(SheetCopyInputs {
            src_title: "Template".into(),
            dst_title: "Copy".into(),
            src_sheet_path: "xl/worksheets/sheet1.xml".into(),
            source_zip_parts: &zip_parts,
            source_rels: &source_rels,
            workbook_xml: &workbook_bytes,
            allocator: &mut alloc,
            existing_table_names: &existing_table_names,
            deep_copy_images: false,
        })
        .expect("plan ok");

        // Both source rIds must be mapped.
        let new_for_table = mutations.rid_remap.get("rId1").expect("rId1 mapped");
        let new_for_vml = mutations.rid_remap.get("rId2").expect("rId2 mapped");
        assert_ne!(new_for_table, new_for_vml);

        let new_xml = std::str::from_utf8(&mutations.new_sheet_xml).unwrap();
        assert!(
            new_xml.contains(&format!(r#"r:id="{new_for_table}""#)),
            "{new_xml}"
        );
        assert!(
            new_xml.contains(&format!(r#"r:id="{new_for_vml}""#)),
            "{new_xml}"
        );
        // Old rIds should be gone.
        assert!(!new_xml.contains(r#"r:id="rId1""#) || new_for_table == "rId1");
    }

    #[test]
    fn dedup_table_name_uses_underscore_with_n_starting_at_two() {
        let mut taken = HashSet::new();
        assert_eq!(dedup_table_name("Sales", &taken), "Sales_2");
        taken.insert("Sales".into());
        assert_eq!(dedup_table_name("Sales", &taken), "Sales_2");
        taken.insert("Sales_2".into());
        assert_eq!(dedup_table_name("Sales", &taken), "Sales_3");
        taken.insert("Sales_3".into());
        assert_eq!(dedup_table_name("Sales", &taken), "Sales_4");
        // A base ending in digits still gets _N — no ambiguity.
        let taken2 = HashSet::new();
        assert_eq!(dedup_table_name("Sales2024", &taken2), "Sales2024_2");
    }

    #[test]
    fn placeholder_rid_reference_consistent_between_sheet_entry_and_workbook_rels() {
        let mut alloc = PartIdAllocator::new();
        let mut zip_parts: HashMap<String, Vec<u8>> = HashMap::new();
        zip_parts.insert("xl/worksheets/sheet1.xml".into(), minimal_sheet_xml());
        let workbook_bytes = one_sheet_workbook(&[], &[]);
        let existing_table_names: HashSet<String> = HashSet::new();
        let source_rels = RelsGraph::new();

        let mutations = plan_sheet_copy(SheetCopyInputs {
            src_title: "Template".into(),
            dst_title: "Copy".into(),
            src_sheet_path: "xl/worksheets/sheet1.xml".into(),
            source_zip_parts: &zip_parts,
            source_rels: &source_rels,
            workbook_xml: &workbook_bytes,
            allocator: &mut alloc,
            existing_table_names: &existing_table_names,
            deep_copy_images: false,
        })
        .expect("plan ok");

        // The placeholder in workbook_sheets_append matches the placeholder
        // in workbook_rels_to_add[0].0 — caller does one rename.
        let placeholder = &mutations.workbook_rels_to_add[0].0;
        let entry = std::str::from_utf8(&mutations.workbook_sheets_append).unwrap();
        assert!(
            entry.contains(placeholder),
            "entry={entry} placeholder={placeholder}"
        );
    }

    #[test]
    fn allocator_advances_across_sequential_calls() {
        let mut alloc = PartIdAllocator::new();
        let mut zip_parts: HashMap<String, Vec<u8>> = HashMap::new();
        zip_parts.insert("xl/worksheets/sheet1.xml".into(), minimal_sheet_xml());
        let workbook_bytes = one_sheet_workbook(&["Other"], &[]);
        let existing_table_names: HashSet<String> = HashSet::new();
        let source_rels = RelsGraph::new();

        let m1 = plan_sheet_copy(SheetCopyInputs {
            src_title: "Template".into(),
            dst_title: "C1".into(),
            src_sheet_path: "xl/worksheets/sheet1.xml".into(),
            source_zip_parts: &zip_parts,
            source_rels: &source_rels,
            workbook_xml: &workbook_bytes,
            allocator: &mut alloc,
            existing_table_names: &existing_table_names,
            deep_copy_images: false,
        })
        .expect("plan 1 ok");
        // Second call: the workbook_xml argument is the same (caller hasn't
        // applied m1 yet), but the allocator has already advanced.
        let m2 = plan_sheet_copy(SheetCopyInputs {
            src_title: "Template".into(),
            dst_title: "C2".into(),
            src_sheet_path: "xl/worksheets/sheet1.xml".into(),
            source_zip_parts: &zip_parts,
            source_rels: &source_rels,
            workbook_xml: &workbook_bytes,
            allocator: &mut alloc,
            existing_table_names: &existing_table_names,
            deep_copy_images: false,
        })
        .expect("plan 2 ok");

        assert_ne!(m1.new_sheet_path, m2.new_sheet_path);
        let n1 = m1
            .new_sheet_path
            .strip_prefix("xl/worksheets/sheet")
            .and_then(|s| s.strip_suffix(".xml"))
            .and_then(|s| s.parse::<u32>().ok())
            .unwrap();
        let n2 = m2
            .new_sheet_path
            .strip_prefix("xl/worksheets/sheet")
            .and_then(|s| s.strip_suffix(".xml"))
            .and_then(|s| s.parse::<u32>().ok())
            .unwrap();
        assert!(n2 > n1, "{n1} < {n2}");
    }

    #[test]
    fn workbook_with_no_defined_names_block_does_not_error() {
        let mut alloc = PartIdAllocator::new();
        let mut zip_parts: HashMap<String, Vec<u8>> = HashMap::new();
        zip_parts.insert("xl/worksheets/sheet1.xml".into(), minimal_sheet_xml());
        let workbook_bytes = workbook_xml(&["Template"], &[]);
        let existing_table_names: HashSet<String> = HashSet::new();
        let source_rels = RelsGraph::new();

        let mutations = plan_sheet_copy(SheetCopyInputs {
            src_title: "Template".into(),
            dst_title: "Copy".into(),
            src_sheet_path: "xl/worksheets/sheet1.xml".into(),
            source_zip_parts: &zip_parts,
            source_rels: &source_rels,
            workbook_xml: &workbook_bytes,
            allocator: &mut alloc,
            existing_table_names: &existing_table_names,
            deep_copy_images: false,
        })
        .expect("plan ok");
        assert!(mutations.defined_names_to_add.is_empty());
    }

    #[test]
    fn unknown_rel_type_is_aliased_not_dropped() {
        let mut alloc = PartIdAllocator::new();
        let mut zip_parts: HashMap<String, Vec<u8>> = HashMap::new();
        zip_parts.insert("xl/worksheets/sheet1.xml".into(), minimal_sheet_xml());
        let workbook_bytes = one_sheet_workbook(&[], &[]);
        let existing_table_names: HashSet<String> = HashSet::new();
        let source_rels = rels_with(&[(
            "http://example.com/rels/weird",
            "../weird/path.bin",
            TargetMode::Internal,
        )]);

        let mutations = plan_sheet_copy(SheetCopyInputs {
            src_title: "Template".into(),
            dst_title: "Copy".into(),
            src_sheet_path: "xl/worksheets/sheet1.xml".into(),
            source_zip_parts: &zip_parts,
            source_rels: &source_rels,
            workbook_xml: &workbook_bytes,
            allocator: &mut alloc,
            existing_table_names: &existing_table_names,
            deep_copy_images: false,
        })
        .expect("plan ok");
        // The rId is in the remap (aliased into dest rels graph).
        assert_eq!(mutations.rid_remap.len(), 1);
    }
}
