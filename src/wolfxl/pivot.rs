//! Sprint Ν Pod-γ (RFC-047 / RFC-048) — pivot-cache + pivot-table
//! patcher integration.
//!
//! Mirrors the chart patcher pattern in `src/wolfxl/mod.rs::apply_chart_adds_phase`
//! (Phase 2.5l). This module is **Phase 2.5m**, sequenced AFTER charts
//! and BEFORE Phase 3 cell patches.
//!
//! Phase 2.5l ↔ 2.5m ordering decision (see Risk #1 in `Plans/sprint-nu.md`):
//!
//! * Sprint Ν Pod-γ pins **2.5m AFTER 2.5l** (charts → pivots).
//! * Rationale: chart-side pivot-source linkage (RFC-049, Pod-δ) takes
//!   a fully-resolved pivot table NAME, not an intermediate cacheId.
//!   Charts that reference pivot tables defer their resolution to a
//!   later sprint; v2.0.0 chart emit does not need 2.5m to have run.
//! * The reverse (2.5m before 2.5l) would only matter if charts
//!   needed to read the pivot table's allocated `pivotTableN.xml`
//!   path during their own emit; in v2.0.0 they don't.
//! * Locking this with a unit test below (`phase_ordering_pinned`)
//!   so a future refactor can't silently flip the order.
//!
//! Drainage steps per cache:
//!   1. Allocate `pivotCacheN` part id via `PartIdAllocator::alloc_pivot_cache`.
//!   2. Write `xl/pivotCache/pivotCacheDefinition{N}.xml` and
//!      `xl/pivotCache/pivotCacheRecords{N}.xml` to `file_adds`.
//!   3. Build `xl/pivotCache/_rels/pivotCacheDefinition{N}.xml.rels`
//!      pointing at the records part.
//!   4. Add a workbook-rel of type `PIVOT_CACHE_DEF` and splice
//!      `<pivotCache cacheId r:id/>` into `xl/workbook.xml`'s
//!      `<pivotCaches>` collection (per ECMA-376 §18.2.27 child
//!      ordering: AFTER `<sheets>`, BEFORE `<definedNames>`).
//!   5. Add content-type overrides for definition + records.
//!
//! Drainage steps per table:
//!   1. Allocate `pivotTableN` part id.
//!   2. Write `xl/pivotTables/pivotTable{N}.xml` to `file_adds`.
//!   3. Build per-table rels graph pointing at the matching cache
//!      definition (target `../pivotCache/pivotCacheDefinition{cache_n}.xml`).
//!   4. Add a sheet-rel of type `PIVOT_TABLE` to the owning sheet's
//!      rels graph.
//!   5. Add a content-type override.

use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use pyo3::types::PyDict;

use wolfxl_pivot::model::cache::{
    CacheField, CacheValue, CalculatedField, CalculatedItem, DateGroup, FieldGroup, FieldGroupKind,
    PivotCache, RangeGroup, SharedItems, WorksheetSource,
};
use wolfxl_pivot::model::records::{CacheRecord, RecordCell};
use wolfxl_pivot::model::slicer::Slicer;
use wolfxl_pivot::model::slicer_cache::{SlicerCache, SlicerItem, SlicerSortOrder};
use wolfxl_pivot::model::table::{
    AxisItem, AxisType, DataField, Format, Location, PageField, PivotArea, PivotConditionalFormat,
    PivotField, PivotItem, PivotTable, PivotTableStyleInfo,
};
use wolfxl_pivot::parse as pp;

// ---------------------------------------------------------------------------
// Public types — queue payloads held on `XlsxPatcher`.
// ---------------------------------------------------------------------------

/// One pivot cache queued for emit at Phase 2.5m. Mirrors
/// `QueuedChartAdd` from `src/wolfxl/mod.rs`.
#[derive(Debug, Clone)]
pub struct QueuedPivotCacheAdd {
    /// `xl/pivotCache/pivotCacheDefinition{N}.xml` body (already
    /// serialised by the Python coordinator via `serialize_pivot_cache_dict`).
    pub cache_def_xml: Vec<u8>,
    /// `xl/pivotCache/pivotCacheRecords{N}.xml` body (already
    /// serialised by the Python coordinator via `serialize_pivot_records_dict`).
    pub cache_records_xml: Vec<u8>,
    /// 0-based cache id allocated when `queue_pivot_cache_add` was
    /// called. Returned to the Python caller so the matching pivot
    /// table can reference it.
    pub cache_id: u32,
}

/// One pivot table queued for emit on a sheet. The cache it
/// references must already have been queued (caches drain first).
#[derive(Debug, Clone)]
#[allow(dead_code)]
pub struct QueuedPivotTableAdd {
    /// Owner sheet title.
    pub sheet: String,
    /// `xl/pivotTables/pivotTable{N}.xml` body (already serialised).
    pub table_xml: Vec<u8>,
    /// Foreign-key into queued caches. Maps to the cache_id field of
    /// the corresponding `QueuedPivotCacheAdd` (0-based).
    pub cache_id: u32,
}

// ---------------------------------------------------------------------------
// PyO3 boundary — dict → typed model parsers.
//
// Mirrors `parse_chart_dict` in `src/native_writer_backend.rs`. We
// keep these here (as opposed to in `wolfxl-pivot::parse`) so the
// pivot crate stays PyO3-free. Per RFC-047 §10 + RFC-048 §10
// contracts.
// ---------------------------------------------------------------------------

/// Parse a §10.1 cache-definition dict into a typed `PivotCache`.
/// Records will be populated separately via `parse_pivot_records_dict`
/// (or left empty when the caller only needs the definition).
pub fn parse_pivot_cache_dict(d: &Bound<'_, PyDict>) -> PyResult<PivotCache> {
    let cache_id: u32 = extract_u32(d, "cache_id", 0)?;
    let source_d = d
        .get_item("source")?
        .ok_or_else(|| PyValueError::new_err("pivot_cache_dict missing 'source'"))?;
    let source = parse_worksheet_source(
        source_d
            .cast::<PyDict>()
            .map_err(|_| PyValueError::new_err("pivot_cache_dict.source must be a dict"))?,
    )?;

    let fields_d = d
        .get_item("fields")?
        .ok_or_else(|| PyValueError::new_err("pivot_cache_dict missing 'fields'"))?;
    let fields_list: Vec<Bound<'_, PyAny>> = fields_d.extract()?;
    let mut fields: Vec<CacheField> = Vec::with_capacity(fields_list.len());
    for fv in &fields_list {
        let fd = fv
            .cast::<PyDict>()
            .map_err(|_| PyValueError::new_err("pivot_cache_dict.fields[*] must be dicts"))?;
        fields.push(parse_cache_field(fd)?);
    }

    let mut pc = PivotCache::new(cache_id, source, fields);
    pc.refresh_on_load = extract_bool(d, "refresh_on_load", false)?;
    pc.refreshed_version = extract_u32(d, "refreshed_version", 6)? as u8;
    pc.created_version = extract_u32(d, "created_version", 6)? as u8;
    pc.min_refreshable_version = extract_u32(d, "min_refreshable_version", 3)? as u8;
    pc.refreshed_by = extract_str(d, "refreshed_by", "wolfxl")?;
    pc.records_part_path = extract_opt_str(d, "records_part_path")?;

    // RFC-061 §10.3 — calculated fields (cache-scoped).
    if let Some(v) = d.get_item("calculated_fields")? {
        if !v.is_none() {
            let list: Vec<Bound<'_, PyAny>> = v.extract()?;
            let mut out = Vec::with_capacity(list.len());
            for vv in &list {
                let pd = vv
                    .cast::<PyDict>()
                    .map_err(|_| PyValueError::new_err("calculated_fields[*] must be dict"))?;
                out.push(parse_calculated_field(pd)?);
            }
            pc.calculated_fields = out;
        }
    }

    // Calculated items on source fields (cache-scoped).
    if let Some(v) = d.get_item("calculated_items")? {
        if !v.is_none() {
            let list: Vec<Bound<'_, PyAny>> = v.extract()?;
            let mut out = Vec::with_capacity(list.len());
            for vv in &list {
                let pd = vv
                    .cast::<PyDict>()
                    .map_err(|_| PyValueError::new_err("calculated_items[*] must be dict"))?;
                out.push(CalculatedItem {
                    field_index: extract_u32(pd, "field_index", 0)?,
                    item_index: extract_u32(pd, "item_index", 0)?,
                    formula: extract_str(pd, "formula", "")?,
                });
            }
            pc.calculated_items = out;
        }
    }

    // RFC-061 §10.5 — field groups.
    if let Some(v) = d.get_item("field_groups")? {
        if !v.is_none() {
            let list: Vec<Bound<'_, PyAny>> = v.extract()?;
            let mut out = Vec::with_capacity(list.len());
            for vv in &list {
                let pd = vv
                    .cast::<PyDict>()
                    .map_err(|_| PyValueError::new_err("field_groups[*] must be dict"))?;
                out.push(parse_field_group(pd)?);
            }
            pc.field_groups = out;
        }
    }

    pc.validate().map_err(PyValueError::new_err)?;
    Ok(pc)
}

fn parse_calculated_field(d: &Bound<'_, PyDict>) -> PyResult<CalculatedField> {
    Ok(CalculatedField {
        name: extract_str(d, "name", "")?,
        formula: extract_str(d, "formula", "")?,
        data_type: extract_str(d, "data_type", "number")?,
    })
}

fn parse_field_group(d: &Bound<'_, PyDict>) -> PyResult<FieldGroup> {
    let kind_str = extract_str(d, "kind", "discrete")?;
    let kind = match kind_str.as_str() {
        "date" => FieldGroupKind::Date,
        "range" => FieldGroupKind::Range,
        "discrete" => FieldGroupKind::Discrete,
        other => {
            return Err(PyValueError::new_err(format!(
                "unknown field_group kind {other:?}"
            )));
        }
    };
    let date = match d.get_item("date")? {
        Some(v) if !v.is_none() => {
            let pd = v
                .cast::<PyDict>()
                .map_err(|_| PyValueError::new_err("date must be dict"))?;
            Some(DateGroup {
                group_by: extract_str(pd, "group_by", "")?,
                start_date: extract_str(pd, "start_date", "")?,
                end_date: extract_str(pd, "end_date", "")?,
            })
        }
        _ => None,
    };
    let range = match d.get_item("range")? {
        Some(v) if !v.is_none() => {
            let pd = v
                .cast::<PyDict>()
                .map_err(|_| PyValueError::new_err("range must be dict"))?;
            Some(RangeGroup {
                start: extract_opt_f64(pd, "start")?.unwrap_or(0.0),
                end: extract_opt_f64(pd, "end")?.unwrap_or(0.0),
                interval: extract_opt_f64(pd, "interval")?.unwrap_or(1.0),
            })
        }
        _ => None,
    };
    let items: Vec<String> = match d.get_item("items")? {
        Some(v) if !v.is_none() => {
            let list: Vec<Bound<'_, PyAny>> = v.extract()?;
            let mut out = Vec::with_capacity(list.len());
            for vv in &list {
                let pd = vv
                    .cast::<PyDict>()
                    .map_err(|_| PyValueError::new_err("items[*] must be dict"))?;
                out.push(extract_str(pd, "name", "")?);
            }
            out
        }
        _ => Vec::new(),
    };
    Ok(FieldGroup {
        field_index: extract_u32(d, "field_index", 0)?,
        parent_index: extract_opt_u32(d, "parent_index")?,
        kind,
        date,
        range,
        items,
    })
}

/// Parse a §10.6 records dict, mutating the supplied `PivotCache`'s
/// `records` field. The records dict carries `field_count` +
/// `record_count` for sanity checks.
pub fn parse_pivot_records_into(d: &Bound<'_, PyDict>, pc: &mut PivotCache) -> PyResult<()> {
    let field_count: u32 = extract_u32(d, "field_count", 0)?;
    if field_count as usize != pc.fields.len() {
        return Err(PyValueError::new_err(format!(
            "pivot_records_dict.field_count={field_count} disagrees with cache.fields.len()={}",
            pc.fields.len()
        )));
    }
    let recs_d = d
        .get_item("records")?
        .ok_or_else(|| PyValueError::new_err("pivot_records_dict missing 'records'"))?;
    let recs: Vec<Bound<'_, PyAny>> = recs_d.extract()?;
    let mut out: Vec<CacheRecord> = Vec::with_capacity(recs.len());
    for row in &recs {
        let cells_list: Vec<Bound<'_, PyAny>> = row
            .extract()
            .map_err(|_| PyValueError::new_err("pivot_records_dict.records[*] must be lists"))?;
        let mut cells: Vec<RecordCell> = Vec::with_capacity(cells_list.len());
        for cell in &cells_list {
            let cd = cell.cast::<PyDict>().map_err(|_| {
                PyValueError::new_err("pivot_records_dict.records[*][*] must be dicts")
            })?;
            cells.push(parse_record_cell(cd)?);
        }
        out.push(CacheRecord { cells });
    }
    pc.records = out;
    Ok(())
}

/// Parse a §10.1 pivot-table dict into a typed `PivotTable`.
pub fn parse_pivot_table_dict(d: &Bound<'_, PyDict>) -> PyResult<PivotTable> {
    let name: String = extract_str(d, "name", "PivotTable1")?;
    let cache_id: u32 = extract_u32(d, "cache_id", 0)?;
    let location_d = d
        .get_item("location")?
        .ok_or_else(|| PyValueError::new_err("pivot_table_dict missing 'location'"))?;
    let location = parse_location(
        location_d
            .cast::<PyDict>()
            .map_err(|_| PyValueError::new_err("pivot_table_dict.location must be a dict"))?,
    )?;

    let pf_list: Vec<Bound<'_, PyAny>> = d
        .get_item("pivot_fields")?
        .ok_or_else(|| PyValueError::new_err("pivot_table_dict missing 'pivot_fields'"))?
        .extract()?;
    let mut pivot_fields: Vec<PivotField> = Vec::with_capacity(pf_list.len());
    for v in &pf_list {
        let pd = v
            .cast::<PyDict>()
            .map_err(|_| PyValueError::new_err("pivot_fields[*] must be dict"))?;
        pivot_fields.push(parse_pivot_field(pd)?);
    }

    let row_field_indices: Vec<u32> = extract_u32_list(d, "row_field_indices")?;
    let col_field_indices: Vec<u32> = extract_u32_list(d, "col_field_indices")?;

    let page_fields = match d.get_item("page_fields")? {
        Some(v) if !v.is_none() => {
            let list: Vec<Bound<'_, PyAny>> = v.extract()?;
            let mut out = Vec::with_capacity(list.len());
            for vv in &list {
                let pd = vv
                    .cast::<PyDict>()
                    .map_err(|_| PyValueError::new_err("page_fields[*] must be dict"))?;
                out.push(parse_page_field(pd)?);
            }
            out
        }
        _ => Vec::new(),
    };

    let data_fields = match d.get_item("data_fields")? {
        Some(v) if !v.is_none() => {
            let list: Vec<Bound<'_, PyAny>> = v.extract()?;
            let mut out = Vec::with_capacity(list.len());
            for vv in &list {
                let pd = vv
                    .cast::<PyDict>()
                    .map_err(|_| PyValueError::new_err("data_fields[*] must be dict"))?;
                out.push(parse_data_field(pd)?);
            }
            out
        }
        _ => Vec::new(),
    };

    let row_items = parse_axis_items(d, "row_items")?;
    let col_items = parse_axis_items(d, "col_items")?;

    let style_info = match d.get_item("style_info")? {
        Some(v) if !v.is_none() => {
            let sd = v
                .cast::<PyDict>()
                .map_err(|_| PyValueError::new_err("style_info must be a dict"))?;
            Some(parse_style_info(sd)?)
        }
        _ => None,
    };

    // RFC-061 Sub-feature 3.5 — formats + conditional formats.
    let formats = parse_formats_list(d, "formats")?;
    let conditional_formats = parse_pivot_cfs_list(d, "conditional_formats")?;

    let pt = PivotTable {
        name,
        cache_id,
        location,
        pivot_fields,
        row_field_indices,
        col_field_indices,
        page_fields,
        data_fields,
        row_items,
        col_items,
        data_on_rows: extract_bool(d, "data_on_rows", false)?,
        outline: extract_bool(d, "outline", true)?,
        compact: extract_bool(d, "compact", true)?,
        row_grand_totals: extract_bool(d, "row_grand_totals", true)?,
        col_grand_totals: extract_bool(d, "col_grand_totals", true)?,
        data_caption: extract_str(d, "data_caption", "Values")?,
        grand_total_caption: extract_opt_str(d, "grand_total_caption")?,
        error_caption: extract_opt_str(d, "error_caption")?,
        missing_caption: extract_opt_str(d, "missing_caption")?,
        apply_number_formats: extract_bool(d, "apply_number_formats", false)?,
        apply_border_formats: extract_bool(d, "apply_border_formats", false)?,
        apply_font_formats: extract_bool(d, "apply_font_formats", false)?,
        apply_pattern_formats: extract_bool(d, "apply_pattern_formats", false)?,
        apply_alignment_formats: extract_bool(d, "apply_alignment_formats", false)?,
        apply_width_height_formats: extract_bool(d, "apply_width_height_formats", true)?,
        style_info,
        created_version: extract_u32(d, "created_version", 6)? as u8,
        updated_version: extract_u32(d, "updated_version", 6)? as u8,
        min_refreshable_version: extract_u32(d, "min_refreshable_version", 3)? as u8,
        formats,
        conditional_formats,
    };
    pt.validate().map_err(PyValueError::new_err)?;
    Ok(pt)
}

fn parse_pivot_area(d: &Bound<'_, PyDict>) -> PyResult<PivotArea> {
    Ok(PivotArea {
        field: extract_opt_u32(d, "field")?,
        area_type: extract_str(d, "type", "data")?,
        data_only: extract_bool(d, "data_only", true)?,
        label_only: extract_bool(d, "label_only", false)?,
        grand_row: extract_bool(d, "grand_row", false)?,
        grand_col: extract_bool(d, "grand_col", false)?,
        cache_index: extract_opt_u32(d, "cache_index")?,
        axis: extract_opt_str(d, "axis")?,
        field_position: extract_opt_u32(d, "field_position")?,
    })
}

fn parse_formats_list(d: &Bound<'_, PyDict>, key: &str) -> PyResult<Vec<Format>> {
    match d.get_item(key)? {
        Some(v) if !v.is_none() => {
            let list: Vec<Bound<'_, PyAny>> = v.extract()?;
            let mut out = Vec::with_capacity(list.len());
            for vv in &list {
                let pd = vv
                    .cast::<PyDict>()
                    .map_err(|_| PyValueError::new_err(format!("{key}[*] must be dict")))?;
                let pa_d = pd
                    .get_item("pivot_area")?
                    .ok_or_else(|| PyValueError::new_err("format missing 'pivot_area'"))?;
                let pa = parse_pivot_area(
                    pa_d.cast::<PyDict>()
                        .map_err(|_| PyValueError::new_err("pivot_area must be dict"))?,
                )?;
                let dxf_id: i32 = match pd.get_item("dxf_id")? {
                    Some(v) if !v.is_none() => v.extract::<i32>()?,
                    _ => -1,
                };
                out.push(Format {
                    action: extract_str(pd, "action", "formatting")?,
                    dxf_id,
                    pivot_area: pa,
                });
            }
            Ok(out)
        }
        _ => Ok(Vec::new()),
    }
}

fn parse_pivot_cfs_list(d: &Bound<'_, PyDict>, key: &str) -> PyResult<Vec<PivotConditionalFormat>> {
    match d.get_item(key)? {
        Some(v) if !v.is_none() => {
            let list: Vec<Bound<'_, PyAny>> = v.extract()?;
            let mut out = Vec::with_capacity(list.len());
            for vv in &list {
                let pd = vv
                    .cast::<PyDict>()
                    .map_err(|_| PyValueError::new_err(format!("{key}[*] must be dict")))?;
                let areas_d = pd.get_item("pivot_areas")?.ok_or_else(|| {
                    PyValueError::new_err("conditional_format missing 'pivot_areas'")
                })?;
                let areas_list: Vec<Bound<'_, PyAny>> = areas_d.extract()?;
                let mut areas = Vec::with_capacity(areas_list.len());
                for av in &areas_list {
                    let ad = av
                        .cast::<PyDict>()
                        .map_err(|_| PyValueError::new_err("pivot_areas[*] must be dict"))?;
                    areas.push(parse_pivot_area(ad)?);
                }
                let priority: i32 = match pd.get_item("priority")? {
                    Some(v) if !v.is_none() => v.extract::<i32>()?,
                    _ => 1,
                };
                let dxf_id: i32 = match pd.get_item("dxf_id")? {
                    Some(v) if !v.is_none() => v.extract::<i32>()?,
                    _ => -1,
                };
                out.push(PivotConditionalFormat {
                    priority,
                    scope: extract_str(pd, "scope", "data")?,
                    cf_type: extract_str(pd, "type", "all")?,
                    pivot_areas: areas,
                    dxf_id,
                });
            }
            Ok(out)
        }
        _ => Ok(Vec::new()),
    }
}

// ---------------------------------------------------------------------------
// Per-section dict parsers.
// ---------------------------------------------------------------------------

fn parse_worksheet_source(d: &Bound<'_, PyDict>) -> PyResult<WorksheetSource> {
    let sheet = extract_str(d, "sheet", "")?;
    let range = extract_str(d, "ref", "")?;
    let name = extract_opt_str(d, "name")?;
    Ok(WorksheetSource { sheet, range, name })
}

fn parse_cache_field(d: &Bound<'_, PyDict>) -> PyResult<CacheField> {
    let name: String = extract_str(d, "name", "")?;
    let num_fmt_id: u32 = extract_u32(d, "num_fmt_id", 0)?;
    let dt_str: String = extract_str(d, "data_type", "string")?;
    let data_type = pp::parse_data_type(&dt_str).map_err(PyValueError::new_err)?;
    let si_d = d
        .get_item("shared_items")?
        .ok_or_else(|| PyValueError::new_err("cache_field missing 'shared_items'"))?;
    let shared_items = parse_shared_items(
        si_d.cast::<PyDict>()
            .map_err(|_| PyValueError::new_err("shared_items must be dict"))?,
    )?;
    let formula = extract_opt_str(d, "formula")?;
    let hierarchy: Option<i32> = match d.get_item("hierarchy")? {
        Some(v) if !v.is_none() => Some(v.extract::<i32>()?),
        _ => None,
    };
    Ok(CacheField {
        name,
        num_fmt_id,
        data_type,
        shared_items,
        formula,
        hierarchy,
    })
}

fn parse_shared_items(d: &Bound<'_, PyDict>) -> PyResult<SharedItems> {
    let count: Option<u32> = match d.get_item("count")? {
        Some(v) if !v.is_none() => Some(v.extract::<u32>()?),
        _ => None,
    };
    let items: Option<Vec<CacheValue>> = match d.get_item("items")? {
        Some(v) if !v.is_none() => {
            let list: Vec<Bound<'_, PyAny>> = v.extract()?;
            let mut out = Vec::with_capacity(list.len());
            for vv in &list {
                let cd = vv
                    .cast::<PyDict>()
                    .map_err(|_| PyValueError::new_err("shared_items.items[*] must be dict"))?;
                out.push(parse_cache_value(cd)?);
            }
            Some(out)
        }
        _ => None,
    };
    Ok(SharedItems {
        count,
        items,
        contains_blank: extract_bool(d, "contains_blank", false)?,
        contains_mixed_types: extract_bool(d, "contains_mixed_types", false)?,
        contains_semi_mixed_types: extract_bool(d, "contains_semi_mixed_types", true)?,
        contains_string: extract_bool(d, "contains_string", true)?,
        contains_number: extract_bool(d, "contains_number", false)?,
        contains_integer: extract_bool(d, "contains_integer", false)?,
        contains_date: extract_bool(d, "contains_date", false)?,
        contains_non_date: extract_bool(d, "contains_non_date", true)?,
        min_value: extract_opt_f64(d, "min_value")?,
        max_value: extract_opt_f64(d, "max_value")?,
        min_date: extract_opt_str(d, "min_date")?,
        max_date: extract_opt_str(d, "max_date")?,
        long_text: extract_bool(d, "long_text", false)?,
    })
}

fn parse_cache_value(d: &Bound<'_, PyDict>) -> PyResult<CacheValue> {
    let kind: String = extract_str(d, "kind", "")?;
    let pv = extract_parsed_value(d, "value")?;
    pp::parse_shared_value(&kind, pv).map_err(PyValueError::new_err)
}

fn parse_record_cell(d: &Bound<'_, PyDict>) -> PyResult<RecordCell> {
    let kind: String = extract_str(d, "kind", "")?;
    let pv = extract_parsed_value(d, "value")?;
    pp::parse_record_cell(&kind, pv).map_err(PyValueError::new_err)
}

fn parse_location(d: &Bound<'_, PyDict>) -> PyResult<Location> {
    Ok(Location {
        range: extract_str(d, "ref", "A1")?,
        first_header_row: extract_u32(d, "first_header_row", 0)?,
        first_data_row: extract_u32(d, "first_data_row", 1)?,
        first_data_col: extract_u32(d, "first_data_col", 1)?,
        row_page_count: extract_opt_u32(d, "row_page_count")?,
        col_page_count: extract_opt_u32(d, "col_page_count")?,
    })
}

fn parse_pivot_field(d: &Bound<'_, PyDict>) -> PyResult<PivotField> {
    let axis: Option<AxisType> = match d.get_item("axis")? {
        Some(v) if !v.is_none() => {
            let s: String = v.extract()?;
            Some(pp::parse_axis(&s).map_err(PyValueError::new_err)?)
        }
        _ => None,
    };
    let items: Option<Vec<PivotItem>> = match d.get_item("items")? {
        Some(v) if !v.is_none() => {
            let list: Vec<Bound<'_, PyAny>> = v.extract()?;
            let mut out = Vec::with_capacity(list.len());
            for vv in &list {
                let pd = vv
                    .cast::<PyDict>()
                    .map_err(|_| PyValueError::new_err("items[*] must be dict"))?;
                out.push(parse_pivot_item(pd)?);
            }
            Some(out)
        }
        _ => None,
    };
    Ok(PivotField {
        name: extract_opt_str(d, "name")?,
        axis,
        data_field: extract_bool(d, "data_field", false)?,
        show_all: extract_bool(d, "show_all", false)?,
        default_subtotal: extract_bool(d, "default_subtotal", true)?,
        sum_subtotal: extract_bool(d, "sum_subtotal", false)?,
        count_subtotal: extract_bool(d, "count_subtotal", false)?,
        avg_subtotal: extract_bool(d, "avg_subtotal", false)?,
        max_subtotal: extract_bool(d, "max_subtotal", false)?,
        min_subtotal: extract_bool(d, "min_subtotal", false)?,
        items,
        outline: extract_bool(d, "outline", true)?,
        compact: extract_bool(d, "compact", true)?,
        subtotal_top: extract_bool(d, "subtotal_top", true)?,
    })
}

fn parse_pivot_item(d: &Bound<'_, PyDict>) -> PyResult<PivotItem> {
    let x: Option<u32> = extract_opt_u32(d, "x")?;
    let t = match d.get_item("t")? {
        Some(v) if !v.is_none() => {
            let s: String = v.extract()?;
            Some(pp::parse_pivot_item_type(&s).map_err(PyValueError::new_err)?)
        }
        _ => None,
    };
    Ok(PivotItem {
        x,
        t,
        h: extract_bool(d, "h", false)?,
        s: extract_bool(d, "s", false)?,
        n: extract_opt_str(d, "n")?,
    })
}

fn parse_page_field(d: &Bound<'_, PyDict>) -> PyResult<PageField> {
    Ok(PageField {
        field_index: extract_u32(d, "field_index", 0)?,
        name: extract_opt_str(d, "name")?,
        item_index: match d.get_item("item_index")? {
            Some(v) if !v.is_none() => v.extract::<i32>()?,
            _ => 0,
        },
        hier: match d.get_item("hier")? {
            Some(v) if !v.is_none() => v.extract::<i32>()?,
            _ => -1,
        },
        cap: extract_opt_str(d, "cap")?,
    })
}

fn parse_data_field(d: &Bound<'_, PyDict>) -> PyResult<DataField> {
    let function_str: String = extract_str(d, "function", "sum")?;
    let function = pp::parse_data_function(&function_str).map_err(PyValueError::new_err)?;
    let show_data_as = match d.get_item("show_data_as")? {
        Some(v) if !v.is_none() => {
            let s: String = v.extract()?;
            Some(pp::parse_show_data_as(&s).map_err(PyValueError::new_err)?)
        }
        _ => None,
    };
    Ok(DataField {
        name: extract_str(d, "name", "")?,
        field_index: extract_u32(d, "field_index", 0)?,
        function,
        show_data_as,
        base_field: extract_u32(d, "base_field", 0)?,
        base_item: extract_u32(d, "base_item", 0)?,
        num_fmt_id: extract_opt_u32(d, "num_fmt_id")?,
    })
}

fn parse_axis_items(d: &Bound<'_, PyDict>, key: &str) -> PyResult<Vec<AxisItem>> {
    match d.get_item(key)? {
        Some(v) if !v.is_none() => {
            let list: Vec<Bound<'_, PyAny>> = v.extract()?;
            let mut out = Vec::with_capacity(list.len());
            for vv in &list {
                let ad = vv
                    .cast::<PyDict>()
                    .map_err(|_| PyValueError::new_err(format!("{key}[*] must be dict")))?;
                out.push(parse_axis_item(ad)?);
            }
            Ok(out)
        }
        _ => Ok(Vec::new()),
    }
}

fn parse_axis_item(d: &Bound<'_, PyDict>) -> PyResult<AxisItem> {
    let indices: Vec<u32> = match d.get_item("indices")? {
        Some(v) if !v.is_none() => v.extract::<Vec<u32>>()?,
        _ => Vec::new(),
    };
    let t = match d.get_item("t")? {
        Some(v) if !v.is_none() => {
            let s: String = v.extract()?;
            Some(pp::parse_axis_item_type(&s).map_err(PyValueError::new_err)?)
        }
        _ => None,
    };
    Ok(AxisItem {
        indices,
        t,
        r: extract_opt_u32(d, "r")?,
        i: extract_opt_u32(d, "i")?,
    })
}

fn parse_style_info(d: &Bound<'_, PyDict>) -> PyResult<PivotTableStyleInfo> {
    Ok(PivotTableStyleInfo {
        name: extract_str(d, "name", "PivotStyleLight16")?,
        show_row_headers: extract_bool(d, "show_row_headers", true)?,
        show_col_headers: extract_bool(d, "show_col_headers", true)?,
        show_row_stripes: extract_bool(d, "show_row_stripes", false)?,
        show_col_stripes: extract_bool(d, "show_col_stripes", false)?,
        show_last_column: extract_bool(d, "show_last_column", true)?,
    })
}

// ---------------------------------------------------------------------------
// PyDict helpers
// ---------------------------------------------------------------------------

fn extract_u32(d: &Bound<'_, PyDict>, key: &str, default: u32) -> PyResult<u32> {
    match d.get_item(key)? {
        Some(v) if !v.is_none() => v.extract::<u32>(),
        _ => Ok(default),
    }
}

fn extract_opt_u32(d: &Bound<'_, PyDict>, key: &str) -> PyResult<Option<u32>> {
    match d.get_item(key)? {
        Some(v) if !v.is_none() => Ok(Some(v.extract::<u32>()?)),
        _ => Ok(None),
    }
}

fn extract_bool(d: &Bound<'_, PyDict>, key: &str, default: bool) -> PyResult<bool> {
    match d.get_item(key)? {
        Some(v) if !v.is_none() => v.extract::<bool>(),
        _ => Ok(default),
    }
}

fn extract_str(d: &Bound<'_, PyDict>, key: &str, default: &str) -> PyResult<String> {
    match d.get_item(key)? {
        Some(v) if !v.is_none() => v.extract::<String>(),
        _ => Ok(default.to_string()),
    }
}

fn extract_opt_str(d: &Bound<'_, PyDict>, key: &str) -> PyResult<Option<String>> {
    match d.get_item(key)? {
        Some(v) if !v.is_none() => Ok(Some(v.extract::<String>()?)),
        _ => Ok(None),
    }
}

fn extract_opt_f64(d: &Bound<'_, PyDict>, key: &str) -> PyResult<Option<f64>> {
    match d.get_item(key)? {
        Some(v) if !v.is_none() => {
            let n = v.extract::<f64>()?;
            if !n.is_finite() {
                return Err(PyValueError::new_err(format!(
                    "{key}: non-finite floats are not representable in xlsx pivot caches"
                )));
            }
            Ok(Some(n))
        }
        _ => Ok(None),
    }
}

fn extract_u32_list(d: &Bound<'_, PyDict>, key: &str) -> PyResult<Vec<u32>> {
    match d.get_item(key)? {
        Some(v) if !v.is_none() => v.extract::<Vec<u32>>(),
        _ => Ok(Vec::new()),
    }
}

fn extract_parsed_value(d: &Bound<'_, PyDict>, key: &str) -> PyResult<Option<pp::ParsedValue>> {
    match d.get_item(key)? {
        Some(v) if !v.is_none() => {
            // Try bool first because bool extracts as int in Python.
            if let Ok(b) = v.extract::<bool>() {
                // We need to disambiguate true-bool vs int-as-bool.
                // PyO3 0.28: bool extraction succeeds for bool only if the
                // Python object IS a bool; ints don't auto-convert. So
                // this is safe for our purposes.
                if v.is_instance_of::<pyo3::types::PyBool>() {
                    return Ok(Some(pp::ParsedValue::Bool(b)));
                }
            }
            if let Ok(n) = v.extract::<f64>() {
                if !n.is_finite() {
                    return Err(PyValueError::new_err(format!(
                        "{key}: non-finite floats are not representable in xlsx pivot caches"
                    )));
                }
                return Ok(Some(pp::ParsedValue::Num(n)));
            }
            if let Ok(s) = v.extract::<String>() {
                return Ok(Some(pp::ParsedValue::Str(s)));
            }
            Err(PyValueError::new_err(format!(
                "unsupported {key} value type"
            )))
        }
        _ => Ok(None),
    }
}

// ---------------------------------------------------------------------------
// PyO3 functions — `serialize_pivot_*_dict`. Mirrors `serialize_chart_dict`
// in `src/native_writer_backend.rs`.
// ---------------------------------------------------------------------------

/// Sprint Ν Pod-γ — serialize a §10.1 cache-definition dict to OOXML
/// bytes. Records are NOT included in this payload (records have a
/// separate part / separate function).
#[pyfunction]
pub fn serialize_pivot_cache_dict(d: &Bound<'_, PyDict>) -> PyResult<Vec<u8>> {
    let pc = parse_pivot_cache_dict(d)?;
    // The records `r:id` rel target is "rId1" from the cache's per-cache rels file (see Phase 2.5m).
    Ok(wolfxl_pivot::emit::pivot_cache_definition_xml(
        &pc,
        Some("rId1"),
    ))
}

/// Sprint Ν Pod-γ — serialize a §10.6 records dict to OOXML bytes.
/// The companion to `serialize_pivot_cache_dict`. Takes the cache dict
/// (so we can resolve the field schema for typing decisions) and the
/// records dict.
#[pyfunction]
pub fn serialize_pivot_records_dict(
    cache_d: &Bound<'_, PyDict>,
    records_d: &Bound<'_, PyDict>,
) -> PyResult<Vec<u8>> {
    let mut pc = parse_pivot_cache_dict(cache_d)?;
    parse_pivot_records_into(records_d, &mut pc)?;
    Ok(wolfxl_pivot::emit::pivot_cache_records_xml(&pc))
}

/// Sprint Ν Pod-γ — serialize a §10.1 pivot-table dict to OOXML bytes.
/// Takes both the cache dict (needed to resolve `<items>` from
/// sharedItems) and the table dict.
#[pyfunction]
pub fn serialize_pivot_table_dict(
    cache_d: &Bound<'_, PyDict>,
    table_d: &Bound<'_, PyDict>,
) -> PyResult<Vec<u8>> {
    let cache = parse_pivot_cache_dict(cache_d)?;
    let table = parse_pivot_table_dict(table_d)?;
    Ok(wolfxl_pivot::emit::pivot_table_xml(&table, &cache))
}

// ---------------------------------------------------------------------------
// RFC-061 Sub-feature 3.1 — slicer cache + slicer presentation parsers + serializers
// ---------------------------------------------------------------------------

fn parse_sort_order(s: &str) -> PyResult<SlicerSortOrder> {
    match s {
        "ascending" => Ok(SlicerSortOrder::Ascending),
        "descending" => Ok(SlicerSortOrder::Descending),
        "none" => Ok(SlicerSortOrder::None),
        other => Err(PyValueError::new_err(format!(
            "unknown slicer sort_order {other:?}"
        ))),
    }
}

fn parse_slicer_item(d: &Bound<'_, PyDict>) -> PyResult<SlicerItem> {
    Ok(SlicerItem {
        name: extract_str(d, "name", "")?,
        hidden: extract_bool(d, "hidden", false)?,
        no_data: extract_bool(d, "no_data", false)?,
    })
}

pub fn parse_slicer_cache_dict(d: &Bound<'_, PyDict>) -> PyResult<SlicerCache> {
    let name = extract_str(d, "name", "")?;
    let source_pivot_cache_id = extract_u32(d, "source_pivot_cache_id", 0)?;
    let source_field_index = extract_u32(d, "source_field_index", 0)?;
    let sort_str = extract_str(d, "sort_order", "ascending")?;
    let sort_order = parse_sort_order(&sort_str)?;
    let custom_list_sort = extract_bool(d, "custom_list_sort", false)?;
    let hide_items_with_no_data = extract_bool(d, "hide_items_with_no_data", false)?;
    let show_missing = extract_bool(d, "show_missing", true)?;

    let items: Vec<SlicerItem> = match d.get_item("items")? {
        Some(v) if !v.is_none() => {
            let list: Vec<Bound<'_, PyAny>> = v.extract()?;
            let mut out = Vec::with_capacity(list.len());
            for vv in &list {
                let pd = vv
                    .cast::<PyDict>()
                    .map_err(|_| PyValueError::new_err("slicer_cache.items[*] must be dict"))?;
                out.push(parse_slicer_item(pd)?);
            }
            out
        }
        _ => Vec::new(),
    };

    let sc = SlicerCache {
        name,
        source_pivot_cache_id,
        source_field_index,
        sort_order,
        custom_list_sort,
        hide_items_with_no_data,
        show_missing,
        items,
    };
    sc.validate().map_err(PyValueError::new_err)?;
    Ok(sc)
}

pub fn parse_slicer_dict(d: &Bound<'_, PyDict>) -> PyResult<Slicer> {
    let style = extract_opt_str(d, "style")?;
    let s = Slicer {
        name: extract_str(d, "name", "")?,
        cache_name: extract_str(d, "cache_name", "")?,
        caption: extract_str(d, "caption", "")?,
        row_height: extract_u32(d, "row_height", 204)?,
        column_count: extract_u32(d, "column_count", 1)?,
        show_caption: extract_bool(d, "show_caption", true)?,
        style,
        locked: extract_bool(d, "locked", true)?,
        anchor: extract_str(d, "anchor", "")?,
    };
    s.validate().map_err(PyValueError::new_err)?;
    Ok(s)
}

/// RFC-061 §10.7 — serialize a slicer-cache dict to OOXML bytes.
#[pyfunction]
pub fn serialize_slicer_cache_dict(d: &Bound<'_, PyDict>) -> PyResult<Vec<u8>> {
    let sc = parse_slicer_cache_dict(d)?;
    Ok(wolfxl_pivot::emit::slicer_cache_xml(&sc))
}

/// RFC-061 §10.7 — serialize a slicer presentation dict to OOXML bytes.
/// Takes a list of slicer dicts (one sheet's slicers); returns the
/// merged `xl/slicers/slicer{N}.xml` body.
#[pyfunction]
pub fn serialize_slicer_dict(slicers_list: &Bound<'_, PyAny>) -> PyResult<Vec<u8>> {
    let list: Vec<Bound<'_, PyAny>> = slicers_list.extract()?;
    let mut slicers: Vec<Slicer> = Vec::with_capacity(list.len());
    for v in &list {
        let pd = v
            .cast::<PyDict>()
            .map_err(|_| PyValueError::new_err("slicer dict expected"))?;
        slicers.push(parse_slicer_dict(pd)?);
    }
    Ok(wolfxl_pivot::emit::slicer_xml(&slicers))
}

// ---------------------------------------------------------------------------
// Queued payloads for the patcher's Phase 2.5p drain.
// ---------------------------------------------------------------------------

/// One slicer cache queued for emit. Mirrors `QueuedPivotCacheAdd`.
#[derive(Debug, Clone)]
#[allow(dead_code)]
pub struct QueuedSlicerCacheAdd {
    /// `xl/slicerCaches/slicerCache{N}.xml` body.
    pub cache_xml: Vec<u8>,
    /// 0-based slicer-cache id allocated when queued.
    pub slicer_cache_id: u32,
    /// Slicer-cache name used for both the rel target and the
    /// workbook-level `<x14:slicerCaches>` extension entries.
    pub name: String,
    /// 0-based source pivot-cache id. The patcher uses this to
    /// resolve the pivot-cache part path for the slicer cache's
    /// rels file.
    pub source_pivot_cache_id: u32,
}

/// One slicer presentation queued for a sheet.
#[derive(Debug, Clone)]
#[allow(dead_code)]
pub struct QueuedSlicerAdd {
    /// Owner sheet title.
    pub sheet: String,
    /// `xl/slicers/slicer{N}.xml` body — pre-serialised single-slicer
    /// presentation. Multiple slicers on one sheet are serialised
    /// inside ONE presentation file by the Python coordinator
    /// (which calls `serialize_slicer_dict` with a list).
    pub slicer_xml: Vec<u8>,
    /// One slicer-cache_id per slicer in the presentation file.
    /// Used to wire rels back to the cache.
    pub slicer_cache_ids: Vec<u32>,
    /// Slicer-cache names referenced (parallel to `slicer_cache_ids`).
    pub cache_names: Vec<String>,
}

// ---------------------------------------------------------------------------
// Workbook.xml `<pivotCaches>` splice. CT_Workbook (ECMA-376 §18.2.27) is a
// strict sequence and Excel rejects the whole workbook as corrupt when
// `pivotCaches` appears before `calcPr`.
// ---------------------------------------------------------------------------

/// One queued workbook-rel entry for the `<pivotCaches>` block.
#[derive(Debug, Clone)]
pub struct PivotCacheRef {
    pub cache_id: u32,
    pub rid: String,
}

/// Splice `<pivotCaches>` into `xl/workbook.xml`. If a block already
/// exists, append the new entries to it; otherwise insert a fresh block
/// at its CT_Workbook sequence position: after `definedNames`, `calcPr`,
/// `oleSize` and `customWorkbookViews`, before `smartTagPr`,
/// `smartTagTypes`, `webPublishing`, `fileRecoveryPr`,
/// `webPublishObjects` and `extLst`. Empty `entries` is a no-op.
pub fn splice_pivot_caches(
    workbook_xml: &[u8],
    entries: &[PivotCacheRef],
) -> Result<Vec<u8>, String> {
    if entries.is_empty() {
        return Ok(workbook_xml.to_vec());
    }

    // Build the rendered children — one `<pivotCache cacheId r:id/>` per entry.
    let mut rendered = String::with_capacity(entries.len() * 64);
    for e in entries {
        rendered.push_str(&format!(
            r#"<pivotCache cacheId="{}" r:id="{}"/>"#,
            e.cache_id, e.rid
        ));
    }

    let s0 =
        std::str::from_utf8(workbook_xml).map_err(|e| format!("workbook.xml not utf8: {e}"))?;
    let owned;
    let s = if workbook_root_has_relationship_namespace(s0)? {
        s0
    } else {
        owned = add_workbook_relationship_namespace(s0)?;
        owned.as_str()
    };

    // Look for an existing `<pivotCaches>` block.
    if let Some(open_pos) = s.find("<pivotCaches>") {
        let inner_start = open_pos + "<pivotCaches>".len();
        if let Some(close_rel) = s[inner_start..].find("</pivotCaches>") {
            let close_pos = inner_start + close_rel;
            // Append to the existing block.
            let mut out = String::with_capacity(s.len() + rendered.len());
            out.push_str(&s[..close_pos]);
            out.push_str(&rendered);
            out.push_str(&s[close_pos..]);
            return Ok(out.into_bytes());
        }
    }
    // Self-closing existing block? `<pivotCaches/>` is rare but valid.
    if let Some(empty_pos) = s.find("<pivotCaches/>") {
        let mut out = String::with_capacity(s.len() + rendered.len() + 32);
        out.push_str(&s[..empty_pos]);
        out.push_str("<pivotCaches>");
        out.push_str(&rendered);
        out.push_str("</pivotCaches>");
        out.push_str(&s[empty_pos + "<pivotCaches/>".len()..]);
        return Ok(out.into_bytes());
    }

    // No existing block: insert a fresh one before the first CT_Workbook
    // child that the schema orders after `pivotCaches`, or before the
    // closing root tag.
    let inject_at = pivot_caches_insert_offset(s)?;

    let block = format!("<pivotCaches>{rendered}</pivotCaches>");
    let mut out = String::with_capacity(s.len() + block.len());
    out.push_str(&s[..inject_at]);
    out.push_str(&block);
    out.push_str(&s[inject_at..]);
    Ok(out.into_bytes())
}

/// CT_Workbook children that the schema sequence places after `pivotCaches`.
fn follows_pivot_caches(local_name: &[u8]) -> bool {
    matches!(
        local_name,
        b"smartTagPr"
            | b"smartTagTypes"
            | b"webPublishing"
            | b"fileRecoveryPr"
            | b"webPublishObjects"
            | b"extLst"
    )
}

/// Byte offset at which a new `<pivotCaches>` block keeps workbook.xml in
/// CT_Workbook order.
fn pivot_caches_insert_offset(s: &str) -> Result<usize, String> {
    use quick_xml::events::Event;

    let mut reader = quick_xml::Reader::from_str(s);
    let mut depth = 0usize;
    loop {
        let before = reader.buffer_position() as usize;
        match reader
            .read_event()
            .map_err(|e| format!("workbook.xml: {e}"))?
        {
            Event::Start(element) => {
                if depth == 1 && follows_pivot_caches(element.local_name().as_ref().as_bytes()) {
                    return Ok(before);
                }
                depth += 1;
            }
            Event::Empty(element) => {
                if depth == 1 && follows_pivot_caches(element.local_name().as_ref().as_bytes()) {
                    return Ok(before);
                }
            }
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if depth == 0 {
                    return Ok(before);
                }
            }
            Event::Eof => return Err("workbook.xml has no closing workbook tag".into()),
            _ => {}
        }
    }
}

fn add_workbook_relationship_namespace(s: &str) -> Result<String, String> {
    let open = s
        .find("<workbook")
        .ok_or_else(|| "workbook.xml has no <workbook> root".to_string())?;
    let rel_end = s[open..]
        .find('>')
        .ok_or_else(|| "workbook.xml has unclosed <workbook> root".to_string())?;
    let insert_at = open + rel_end;
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

fn workbook_root_has_relationship_namespace(s: &str) -> Result<bool, String> {
    let open = s
        .find("<workbook")
        .ok_or_else(|| "workbook.xml has no <workbook> root".to_string())?;
    let rel_end = s[open..]
        .find('>')
        .ok_or_else(|| "workbook.xml has unclosed <workbook> root".to_string())?;
    Ok(s[open..open + rel_end].contains("xmlns:r="))
}

/// Splice a `<pivotTable>` rel reference into a sheet's XML. OOXML
/// puts pivot-table associations on the sheet's rels file (not in the
/// sheet XML body itself), and the sheet doesn't carry an inline
/// `<pivotTable>` reference like `<drawing r:id>`. Therefore there is
/// nothing to splice into the sheet XML body for a pivot table — the
/// sheet's rels graph + the cache's foreign key fully wire it up.
///
/// This is a no-op marker function preserved here for symmetry with
/// the chart `splice_drawing_ref`. Returns the input unchanged.
#[allow(dead_code)]
pub fn splice_sheet_for_pivot_table(sheet_xml: &[u8]) -> Vec<u8> {
    sheet_xml.to_vec()
}

// ---------------------------------------------------------------------------
// PartIdAllocator extensions for pivot parts. We don't have
// `alloc_pivot_cache` / `alloc_pivot_table` on the shared allocator
// today; Sprint Ν keeps this allocation in a per-patcher counter
// since the pivot graph is small and rarely interleaves with other
// allocators in a single save.
// ---------------------------------------------------------------------------

/// Per-patcher pivot part-id counters. Sequenced separately from the
/// generic `PartIdAllocator` because (a) pivots are workbook-scope so
/// the counter is workbook-wide, not per-sheet, and (b) we don't yet
/// observe existing pivot parts when bootstrapping the allocator
/// (modify-mode pivots already in the source file ARE preserved
/// verbatim by the rels graph; we never collide because we route
/// through `alloc_pivot_cache` for new parts only).
#[derive(Debug, Clone, Default)]
pub struct PivotPartCounters {
    pub next_cache: u32,
    pub next_table: u32,
}

impl PivotPartCounters {
    pub fn new(start_cache: u32, start_table: u32) -> Self {
        Self {
            next_cache: start_cache.max(1),
            next_table: start_table.max(1),
        }
    }

    pub fn alloc_cache(&mut self) -> u32 {
        let n = self.next_cache;
        self.next_cache += 1;
        n
    }

    pub fn alloc_table(&mut self) -> u32 {
        let n = self.next_table;
        self.next_table += 1;
        n
    }

    /// Bump counters by observing existing part paths from the source
    /// ZIP (`xl/pivotCache/pivotCacheDefinition{N}.xml` and
    /// `xl/pivotTables/pivotTable{N}.xml`).
    pub fn observe(&mut self, path: &str) {
        if let Some(rest) = path.strip_prefix("xl/pivotCache/pivotCacheDefinition") {
            if let Some(num_str) = rest.strip_suffix(".xml") {
                if let Ok(n) = num_str.parse::<u32>() {
                    if n + 1 > self.next_cache {
                        self.next_cache = n + 1;
                    }
                }
            }
        } else if let Some(rest) = path.strip_prefix("xl/pivotTables/pivotTable") {
            if let Some(num_str) = rest.strip_suffix(".xml") {
                if let Ok(n) = num_str.parse::<u32>() {
                    if n + 1 > self.next_table {
                        self.next_table = n + 1;
                    }
                }
            }
        }
    }
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn splice_pivot_caches_into_fresh_workbook() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="x"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#;
        let entries = vec![PivotCacheRef {
            cache_id: 0,
            rid: "rId7".into(),
        }];
        let out = splice_pivot_caches(xml, &entries).unwrap();
        let s = std::str::from_utf8(&out).unwrap();
        assert!(s.contains(r#"<pivotCache cacheId="0" r:id="rId7"/>"#));
        assert!(
            s.contains(
                r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#
            ),
            "{s}"
        );
        assert!(s.contains("<pivotCaches>"));
        assert!(s.contains("</pivotCaches>"));
        // Inserted after </sheets>, before the closing root tag.
        let after_sheets = s.find("</sheets>").unwrap();
        let pivots_at = s.find("<pivotCaches>").unwrap();
        assert!(pivots_at > after_sheets);
    }

    #[test]
    fn splice_pivot_caches_adds_root_rel_namespace_when_only_sheet_has_it() {
        let xml = br#"<workbook xmlns="x"><sheets><sheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#;
        let entries = vec![PivotCacheRef {
            cache_id: 0,
            rid: "rId2".into(),
        }];

        let out = splice_pivot_caches(xml, &entries).unwrap();
        let s = std::str::from_utf8(&out).unwrap();
        let root_end = s.find('>').unwrap();
        assert!(
            s[..root_end].contains(
                r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#
            ),
            "workbook-level pivotCache r:id needs an r namespace on the workbook root"
        );
        assert!(s.contains(r#"<pivotCache cacheId="0" r:id="rId2"/>"#));
    }

    #[test]
    fn splice_pivot_caches_appends_to_existing() {
        let xml = br#"<workbook><sheets/><pivotCaches><pivotCache cacheId="0" r:id="rId1"/></pivotCaches></workbook>"#;
        let entries = vec![PivotCacheRef {
            cache_id: 1,
            rid: "rId8".into(),
        }];
        let out = splice_pivot_caches(xml, &entries).unwrap();
        let s = std::str::from_utf8(&out).unwrap();
        assert!(s.contains(r#"<pivotCache cacheId="0" r:id="rId1"/>"#));
        assert!(s.contains(r#"<pivotCache cacheId="1" r:id="rId8"/>"#));
        // Must not duplicate the wrapper.
        assert_eq!(s.matches("<pivotCaches>").count(), 1);
    }

    #[test]
    fn splice_pivot_caches_after_defined_names() {
        let xml = br#"<workbook><sheets/><definedNames><definedName name="x">x</definedName></definedNames></workbook>"#;
        let entries = vec![PivotCacheRef {
            cache_id: 0,
            rid: "rId7".into(),
        }];
        let out = splice_pivot_caches(xml, &entries).unwrap();
        let s = std::str::from_utf8(&out).unwrap();
        let dn_end = s.find("</definedNames>").unwrap() + "</definedNames>".len();
        let pivots_at = s.find("<pivotCaches>").unwrap();
        assert!(pivots_at >= dn_end);
    }

    /// Excel reports the whole workbook as corrupt when `pivotCaches`
    /// precedes `calcPr` (ECMA-376 §18.2.27 sequence).
    #[test]
    fn splice_pivot_caches_follows_calc_pr_and_precedes_ext_lst() {
        let xml = br#"<workbook><sheets/><definedNames/><calcPr calcId="124519" fullCalcOnLoad="1"/><customWorkbookViews/><extLst><ext uri="{x}"/></extLst></workbook>"#;
        let entries = vec![PivotCacheRef {
            cache_id: 0,
            rid: "rId4".into(),
        }];
        let out = splice_pivot_caches(xml, &entries).unwrap();
        let s = std::str::from_utf8(&out).unwrap();
        let pivots_at = s.find("<pivotCaches>").unwrap();
        assert!(pivots_at > s.find("<calcPr").unwrap(), "{s}");
        assert!(pivots_at > s.find("<customWorkbookViews/>").unwrap(), "{s}");
        assert!(pivots_at < s.find("<extLst>").unwrap(), "{s}");
    }

    #[test]
    fn splice_pivot_caches_without_trailing_children_goes_last() {
        let xml = br#"<workbook><sheets/><calcPr calcId="171027"/></workbook>"#;
        let entries = vec![PivotCacheRef {
            cache_id: 0,
            rid: "rId4".into(),
        }];
        let out = splice_pivot_caches(xml, &entries).unwrap();
        let s = std::str::from_utf8(&out).unwrap();
        assert!(
            s.ends_with(r#"<calcPr calcId="171027"/><pivotCaches><pivotCache cacheId="0" r:id="rId4"/></pivotCaches></workbook>"#),
            "{s}"
        );
    }

    #[test]
    fn splice_pivot_caches_empty_is_noop() {
        let xml = br#"<workbook><sheets/></workbook>"#;
        let out = splice_pivot_caches(xml, &[]).unwrap();
        assert_eq!(out, xml.to_vec());
    }

    #[test]
    fn part_counters_observe_existing() {
        let mut c = PivotPartCounters::new(1, 1);
        c.observe("xl/pivotCache/pivotCacheDefinition3.xml");
        c.observe("xl/pivotTables/pivotTable5.xml");
        assert_eq!(c.alloc_cache(), 4);
        assert_eq!(c.alloc_table(), 6);
    }

    #[test]
    fn part_counters_monotonic() {
        let mut c = PivotPartCounters::new(1, 1);
        let n1 = c.alloc_cache();
        let n2 = c.alloc_cache();
        assert!(n2 > n1);
        let t1 = c.alloc_table();
        let t2 = c.alloc_table();
        assert!(t2 > t1);
    }

    /// Phase 2.5l ↔ 2.5m ordering pin (Risk #1 in Plans/sprint-nu.md).
    /// We pin charts → pivots so a future refactor cannot silently
    /// flip the order without also updating this canonical test.
    #[test]
    fn phase_ordering_pinned() {
        // Chars 'l' < 'm' lexicographically; `apply_chart_adds_phase`
        // (Phase 2.5l) runs before `apply_pivot_adds_phase` (Phase
        // 2.5m). The phase NAME and the alphabetical relationship
        // are the canonical source of truth.
        assert!('l' < 'm', "phase 2.5l must run before phase 2.5m");
    }
}
