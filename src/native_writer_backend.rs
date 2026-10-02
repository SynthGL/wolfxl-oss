//! Native xlsx writer pyclass — sole write-mode backend (W5+).
//!
//! Exposes the 22 pymethods that [`python::wolfxl::_backend.make_writer`]
//! constructs for every `Workbook()`. The 13 cell/format/structure
//! pymethods + `save()` drive [`wolfxl_writer::Workbook`] and emit a
//! complete xlsx via [`wolfxl_writer::emit_xlsx`]. The legacy
//! `rust_xlsxwriter`-backed sibling pyclass was removed in W5; the
//! payload-shape contract documented below is preserved verbatim for
//! Python-side compatibility.
//!
//! # Why mirror oracle exactly
//!
//! The Python flush path (`_flush_to_writer`) calls these methods with
//! payloads built by `python_value_to_payload`, `font_to_format_dict`,
//! etc. Those builders are oracle-shaped — keys, types, and value
//! coercions match what the oracle consumes. This file consumes the
//! same dicts so no Python-side change is required to drive native.
//!
//! # Style-merge limitation (4B follow-up)
//!
//! Calling `write_cell_format` after `write_cell_value` on the same
//! cell **replaces** the style_id with one freshly interned from the
//! format dict alone. Calling `write_cell_border` after that replaces
//! again. Oracle merges format + border at save time because it stores
//! them in separate HashMaps; native interns eagerly per call. In
//! practice the Python flush path always calls `write_cell_format`
//! and `write_cell_border` together (or only one), so the smoke test
//! does not hit the merge path. Documented gap; fix in 4B by exposing
//! a `StylesBuilder::lookup_format(style_id)` reverse query.

use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use pyo3::types::PyDict;

use wolfxl_writer::model::Worksheet;
use wolfxl_writer::Workbook;

use crate::native_writer_autofilter::install_autofilter;
use crate::native_writer_cells::{
    parse_a1_to_row_col, write_array_formula_cell, write_cell_payload, write_rich_text_cell,
    write_value_grid, write_value_style_id_grid,
};
use crate::native_writer_charts::parse_chart_dict;
use crate::native_writer_formats::{
    apply_border_grid, apply_cell_border, apply_cell_format, apply_format_grid, apply_style_id_grid,
};
use crate::native_writer_images::dict_to_sheet_image;
use crate::native_writer_sheet_features::{
    dict_to_comment, dict_to_conditional_format, dict_to_data_validation, dict_to_hyperlink,
    dict_to_person, dict_to_table, dict_to_threaded_comment, unwrap_optional_wrapper,
};
use crate::native_writer_sheet_state::{
    apply_column_dimension, apply_column_width, apply_freeze_panes, apply_merged_range,
    apply_page_breaks, apply_print_area, apply_row_dimension, apply_row_height, apply_sheet_setup,
};
use crate::native_writer_streaming::{
    append_streaming_row, enable_streaming, finalize_all_streaming,
};
use crate::native_writer_workbook::{add_sheet_if_missing, move_sheet, remove_sheet, rename_sheet};
use crate::native_writer_workbook_metadata::{
    dict_to_defined_name, dict_to_doc_properties, dict_to_workbook_security,
};

// ---------------------------------------------------------------------------
// PyClass
// ---------------------------------------------------------------------------

#[pyclass(unsendable)]
pub struct NativeWorkbook {
    inner: Workbook,
}

fn require_sheet<'wb>(wb: &'wb mut Workbook, name: &str) -> PyResult<&'wb mut Worksheet> {
    wb.sheet_mut_by_name(name)
        .ok_or_else(|| PyValueError::new_err(format!("Unknown sheet: {name}")))
}

// ---------------------------------------------------------------------------
// PyMethods
// ---------------------------------------------------------------------------

#[pymethods]
impl NativeWorkbook {
    #[new]
    pub fn new() -> Self {
        Self {
            inner: Workbook::new(),
        }
    }

    pub fn add_sheet(&mut self, name: &str) -> PyResult<()> {
        // Mirror oracle's idempotent semantic: re-adding an existing sheet is a no-op.
        add_sheet_if_missing(&mut self.inner, name);
        Ok(())
    }

    pub fn rename_sheet(&mut self, old_name: &str, new_name: &str) -> PyResult<()> {
        rename_sheet(&mut self.inner, old_name, new_name)
    }

    pub fn move_sheet(&mut self, name: &str, offset: isize) -> PyResult<()> {
        move_sheet(&mut self.inner, name, offset)
    }

    pub fn remove_sheet(&mut self, name: &str) -> PyResult<()> {
        remove_sheet(&mut self.inner, name)
    }

    pub fn write_cell_value(
        &mut self,
        sheet: &str,
        a1: &str,
        payload: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        write_cell_payload(&mut self.inner, sheet, a1, payload)
    }

    /// Sprint Ι Pod-α: write a rich-text inline-string cell.
    ///
    /// `runs` is a Python list of ``(text, font_dict_or_None)``
    /// tuples — same shape as the patcher's
    /// ``queue_rich_text_value``. The native writer emits an
    /// inline-string `<c t="inlineStr">` cell so the SST stays
    /// untouched.
    pub fn write_cell_rich_text(
        &mut self,
        sheet: &str,
        a1: &str,
        runs: &Bound<'_, pyo3::types::PyList>,
    ) -> PyResult<()> {
        write_rich_text_cell(&mut self.inner, sheet, a1, runs)
    }

    /// RFC-057 (Sprint Ο Pod 1C): write an array-formula / data-table
    /// formula / spill-child cell.
    ///
    /// `payload` mirrors the patcher's ``queue_array_formula`` shape:
    ///   - ``{"kind": "array", "ref": "A1:A10", "text": "B1:B10*2"}``
    ///   - ``{"kind": "data_table", "ref": "B2:F11", "ca": false,
    ///        "dt2D": true, "dtr": false, "r1": "A1", "r2": "A2"}``
    ///   - ``{"kind": "spill_child"}``
    pub fn write_cell_array_formula(
        &mut self,
        sheet: &str,
        a1: &str,
        payload: &Bound<'_, PyDict>,
    ) -> PyResult<()> {
        write_array_formula_cell(&mut self.inner, sheet, a1, payload)
    }

    /// Bulk-write a rectangular grid of values starting at `start_a1`.
    pub fn write_sheet_values(
        &mut self,
        sheet: &str,
        start_a1: &str,
        values: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        write_value_grid(&mut self.inner, sheet, start_a1, values)
    }

    pub fn write_sheet_values_with_style_ids(
        &mut self,
        sheet: &str,
        start_a1: &str,
        values: &Bound<'_, PyAny>,
        style_ids: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        write_value_style_id_grid(&mut self.inner, sheet, start_a1, values, style_ids)
    }

    pub fn write_cell_format(
        &mut self,
        sheet: &str,
        a1: &str,
        format_dict: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let (row, col) = parse_a1_to_row_col(a1)?;
        let dict = format_dict
            .cast::<PyDict>()
            .map_err(|_| PyValueError::new_err("format_dict must be a dict"))?;
        apply_cell_format(&mut self.inner, sheet, row, col, dict)
    }

    pub fn write_cell_border(
        &mut self,
        sheet: &str,
        a1: &str,
        border_dict: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let (row, col) = parse_a1_to_row_col(a1)?;
        let dict = border_dict
            .cast::<PyDict>()
            .map_err(|_| PyValueError::new_err("border_dict must be a dict"))?;
        apply_cell_border(&mut self.inner, sheet, row, col, dict)
    }

    pub fn write_sheet_formats(
        &mut self,
        sheet: &str,
        start_a1: &str,
        formats: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let (base_row, base_col) = parse_a1_to_row_col(start_a1)?;
        apply_format_grid(&mut self.inner, sheet, base_row, base_col, formats)
    }

    pub fn write_sheet_style_ids(
        &mut self,
        sheet: &str,
        start_a1: &str,
        style_ids: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let (base_row, base_col) = parse_a1_to_row_col(start_a1)?;
        apply_style_id_grid(&mut self.inner, sheet, base_row, base_col, style_ids)
    }

    pub fn write_sheet_borders(
        &mut self,
        sheet: &str,
        start_a1: &str,
        borders: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let (base_row, base_col) = parse_a1_to_row_col(start_a1)?;
        apply_border_grid(&mut self.inner, sheet, base_row, base_col, borders)
    }

    pub fn set_row_height(&mut self, sheet: &str, row: u32, height: f64) -> PyResult<()> {
        // Python passes 1-based rows (openpyxl convention). The native
        // model is also 1-based, so no `saturating_sub(1)` like oracle.
        let ws = require_sheet(&mut self.inner, sheet)?;
        apply_row_height(ws, row, height);
        Ok(())
    }

    pub fn set_row_dimension(
        &mut self,
        sheet: &str,
        row: u32,
        payload: &Bound<'_, PyDict>,
    ) -> PyResult<()> {
        let ws = require_sheet(&mut self.inner, sheet)?;
        apply_row_dimension(ws, row, payload)
    }

    pub fn set_column_width(&mut self, sheet: &str, col_str: &str, width: f64) -> PyResult<()> {
        let ws = require_sheet(&mut self.inner, sheet)?;
        apply_column_width(ws, col_str, width)
    }

    pub fn set_column_dimension(
        &mut self,
        sheet: &str,
        col_str: &str,
        payload: &Bound<'_, PyDict>,
    ) -> PyResult<()> {
        let ws = require_sheet(&mut self.inner, sheet)?;
        apply_column_dimension(ws, col_str, payload)
    }

    pub fn merge_cells(&mut self, sheet: &str, range_str: &str) -> PyResult<()> {
        let ws = require_sheet(&mut self.inner, sheet)?;
        apply_merged_range(ws, range_str)
    }

    /// Mirrors oracle: dict with keys `mode` ("freeze" | "split") and
    /// either `top_left_cell` (freeze) or `x_split` / `y_split` (split).
    /// Optional wrapper key `freeze` is also accepted.
    pub fn set_freeze_panes(&mut self, sheet: &str, settings: &Bound<'_, PyAny>) -> PyResult<()> {
        let ws = require_sheet(&mut self.inner, sheet)?;
        apply_freeze_panes(ws, settings)
    }

    pub fn save(&mut self, path: &str) -> PyResult<()> {
        crate::native_writer_workbook::save(&mut self.inner, path)
    }

    // =========================================================================
    // Sprint 7 / G20 — `Workbook(write_only=True)` streaming write mode.
    // =========================================================================

    /// Convert an empty existing sheet into streaming write-only mode.
    /// Idempotent. Errors if the sheet already has eager rows.
    pub fn enable_streaming_sheet(&mut self, sheet: &str) -> PyResult<()> {
        enable_streaming(&mut self.inner, sheet)
    }

    /// Append one row's worth of cell payloads to a streaming sheet's
    /// temp file. `cells` is a list of `dict | None`, indexed `column - 1`.
    pub fn append_streaming_row(
        &mut self,
        sheet: &str,
        row_idx: u32,
        cells: &Bound<'_, pyo3::types::PyList>,
    ) -> PyResult<()> {
        append_streaming_row(&mut self.inner, sheet, row_idx, cells)
    }

    /// Flush every streaming sheet's `BufWriter`. Called automatically
    /// from `save`, but exposed so the Python `WriteOnlyWorksheet.close()`
    /// path can release file descriptors before save if needed.
    pub fn finalize_streaming_sheets(&mut self) -> PyResult<()> {
        finalize_all_streaming(&mut self.inner)
    }

    /// Intern a format dict on the workbook's StylesBuilder and return
    /// the resulting style id. Exposed so the streaming
    /// `WriteOnlyCell` path can resolve a font/fill/border/alignment/
    /// number_format combo to one style id once and cache it Python-
    /// side, instead of paying a write_cell_format hop per cell.
    pub fn intern_format(&mut self, format_dict: &Bound<'_, pyo3::types::PyAny>) -> PyResult<u32> {
        let dict = format_dict
            .cast::<PyDict>()
            .map_err(|_| PyValueError::new_err("format_dict must be a dict"))?;
        crate::native_writer_formats::intern_format_from_dict(&mut self.inner, dict)
    }

    // =========================================================================
    // Wave 4B rich-feature pymethods — real implementations.
    // =========================================================================

    pub fn add_hyperlink(&mut self, sheet: &str, link_dict: &Bound<'_, PyAny>) -> PyResult<()> {
        let dict = link_dict
            .cast::<PyDict>()
            .map_err(|_| PyValueError::new_err("link must be a dict"))?;
        let cfg = unwrap_optional_wrapper(dict, "hyperlink")?;
        let Some((a1, hyperlink)) = dict_to_hyperlink(&cfg)? else {
            return Ok(()); // silent no-op — match oracle
        };
        let ws = require_sheet(&mut self.inner, sheet)?;
        ws.hyperlinks.insert(a1, hyperlink);
        Ok(())
    }

    pub fn add_comment(&mut self, sheet: &str, comment_dict: &Bound<'_, PyAny>) -> PyResult<()> {
        let dict = comment_dict
            .cast::<PyDict>()
            .map_err(|_| PyValueError::new_err("comment_dict must be a dict"))?;
        let cfg = unwrap_optional_wrapper(dict, "comment")?;
        // Borrow authors first (resolution must happen before re-borrowing inner for sheet).
        let Some((a1, comment)) = dict_to_comment(&cfg, &mut self.inner.comment_authors)? else {
            return Ok(()); // silent no-op — match oracle
        };
        let ws = require_sheet(&mut self.inner, sheet)?;
        ws.comments.insert(a1, comment);
        Ok(())
    }

    /// Append one threaded comment payload (top-level or reply) to a sheet.
    /// RFC-068 / G08. Idempotency on the Python side ensures we don't double-add.
    pub fn add_threaded_comment(
        &mut self,
        sheet: &str,
        payload: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let dict = payload
            .cast::<PyDict>()
            .map_err(|_| PyValueError::new_err("threaded_comment payload must be a dict"))?;
        let Some(tc) = dict_to_threaded_comment(&dict)? else {
            return Ok(()); // silent no-op — caller didn't supply required keys
        };
        let ws = require_sheet(&mut self.inner, sheet)?;
        ws.threaded_comments.push(tc);
        Ok(())
    }

    /// Register one Person in the workbook-scope person table. RFC-068 / G08.
    pub fn add_person(&mut self, payload: &Bound<'_, PyAny>) -> PyResult<()> {
        let dict = payload
            .cast::<PyDict>()
            .map_err(|_| PyValueError::new_err("person payload must be a dict"))?;
        let Some(person) = dict_to_person(&dict)? else {
            return Ok(()); // silent no-op
        };
        // Idempotency on the Rust side: skip if a Person with the same id is
        // already present. The Python registry already enforces this, but a
        // defensive check keeps the personList byte-stable on repeated flush.
        if self.inner.persons.contains_id(&person.id) {
            return Ok(());
        }
        self.inner.persons.push(person);
        Ok(())
    }

    pub fn set_print_area(&mut self, sheet: &str, range_str: &str) -> PyResult<()> {
        let ws = require_sheet(&mut self.inner, sheet)?;
        apply_print_area(ws, range_str);
        Ok(())
    }

    /// Sprint Ο Pod 1A.5 (RFC-055) — install sheet-setup blocks on a
    /// write-mode sheet. Consumes the §10 dict shape produced by
    /// `Worksheet.to_rust_setup_dict()` and copies the parsed specs
    /// onto the worksheet model. The native writer's emit pass then
    /// renders them at the correct CT_Worksheet positions.
    pub fn set_sheet_setup_native(
        &mut self,
        sheet: &str,
        payload: &Bound<'_, PyDict>,
    ) -> PyResult<()> {
        let ws = require_sheet(&mut self.inner, sheet)?;
        apply_sheet_setup(ws, payload)
    }

    /// Sprint Π Pod Π-α (RFC-062) — install page-breaks +
    /// sheetFormatPr blocks on a write-mode sheet. Consumes the
    /// merged §10 dict shape produced by
    /// ``Worksheet.to_rust_page_breaks_dict()`` +
    /// ``Worksheet.to_rust_sheet_format_dict()``.
    pub fn set_page_breaks_native(
        &mut self,
        sheet: &str,
        payload: &Bound<'_, PyDict>,
    ) -> PyResult<()> {
        let ws = require_sheet(&mut self.inner, sheet)?;
        apply_page_breaks(ws, payload)
    }

    /// Sprint Ο Pod 1B (RFC-056) — install an autoFilter on a write-
    /// mode sheet. Takes the §10 dict, pre-emits the `<autoFilter>`
    /// block via `wolfxl_autofilter::emit::emit`, and evaluates the
    /// filter against the worksheet's in-memory cells to stamp
    /// `row.hidden = true` on rows excluded by the filter.
    pub fn set_autofilter_native(&mut self, sheet: &str, d: &Bound<'_, PyDict>) -> PyResult<()> {
        let ws = require_sheet(&mut self.inner, sheet)?;
        install_autofilter(ws, d)
    }

    pub fn add_conditional_format(
        &mut self,
        sheet: &str,
        rule_dict: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let dict = rule_dict
            .cast::<PyDict>()
            .map_err(|_| PyValueError::new_err("rule must be a dict"))?;
        let cfg = unwrap_optional_wrapper(dict, "cf_rule")?;
        // Resolve CF (may intern a dxf — borrows styles) before borrowing sheet.
        let Some(cf) = dict_to_conditional_format(&cfg, &mut self.inner.styles)? else {
            return Ok(()); // silent no-op — match oracle
        };
        let ws = require_sheet(&mut self.inner, sheet)?;
        ws.conditional_formats.push(cf);
        Ok(())
    }

    pub fn add_data_validation(
        &mut self,
        sheet: &str,
        validation_dict: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let dict = validation_dict
            .cast::<PyDict>()
            .map_err(|_| PyValueError::new_err("validation must be a dict"))?;
        let cfg = unwrap_optional_wrapper(dict, "validation")?;
        let Some(dv) = dict_to_data_validation(&cfg)? else {
            return Ok(()); // silent no-op — match oracle
        };
        let ws = require_sheet(&mut self.inner, sheet)?;
        ws.validations.push(dv);
        Ok(())
    }

    pub fn add_named_range(&mut self, sheet: &str, named_range: &Bound<'_, PyAny>) -> PyResult<()> {
        let dict = named_range
            .cast::<PyDict>()
            .map_err(|_| PyValueError::new_err("named_range must be a dict"))?;
        let cfg = unwrap_optional_wrapper(dict, "named_range")?;
        // sheet_index_by_name borrows &self.inner immutably — do it before any &mut borrow.
        let Some(dn) = dict_to_defined_name(&self.inner, sheet, &cfg)? else {
            return Ok(()); // silent no-op — match oracle
        };
        self.inner.defined_names.push(dn);
        Ok(())
    }

    pub fn add_named_style(&mut self, name: &str) -> PyResult<()> {
        self.inner.styles.add_named_style(name);
        Ok(())
    }

    pub fn add_table(&mut self, sheet: &str, table: &Bound<'_, PyAny>) -> PyResult<()> {
        let dict = table
            .cast::<PyDict>()
            .map_err(|_| PyValueError::new_err("table must be a dict"))?;
        let cfg = unwrap_optional_wrapper(dict, "table")?;
        let Some(tbl) = dict_to_table(&cfg)? else {
            return Ok(()); // silent no-op — match oracle
        };
        let ws = require_sheet(&mut self.inner, sheet)?;
        ws.tables.push(tbl);
        Ok(())
    }

    pub fn set_properties(&mut self, props: &Bound<'_, PyAny>) -> PyResult<()> {
        let dict = props
            .cast::<PyDict>()
            .map_err(|_| PyValueError::new_err("props must be a dict"))?;
        let doc_props = dict_to_doc_properties(dict)?;
        self.inner.set_doc_props(doc_props);
        Ok(())
    }

    /// Sprint Ο Pod 1D (RFC-058) — set workbook-level security.
    ///
    /// Accepts the §10 flat dict shape (`workbook_protection` and
    /// `file_sharing` top-level keys, either of which may be `None`).
    /// Replaces any previously-set security (last writer wins).
    pub fn set_workbook_security(&mut self, payload: &Bound<'_, PyAny>) -> PyResult<()> {
        let dict = payload
            .cast::<PyDict>()
            .map_err(|_| PyValueError::new_err("workbook security payload must be a dict"))?;
        let security = dict_to_workbook_security(dict)?;
        self.inner.security = security;
        Ok(())
    }

    pub fn set_date1904(&mut self, date1904: bool) -> PyResult<()> {
        self.inner.date1904 = date1904;
        Ok(())
    }

    /// Sprint Λ Pod-β (RFC-045) — queue an image onto a sheet.
    ///
    /// `image_dict` shape (built by Python's
    /// ``Worksheet.add_image``):
    ///
    /// ```python
    /// {
    ///     "data": <bytes>,
    ///     "ext": "png" | "jpeg" | "gif" | "bmp",
    ///     "width": int,
    ///     "height": int,
    ///     "anchor": {
    ///         "type": "one_cell" | "two_cell" | "absolute",
    ///         # one_cell: from_col, from_row, from_col_off, from_row_off (0-based + EMU)
    ///         # two_cell: + to_col, to_row, to_col_off, to_row_off, edit_as
    ///         # absolute: x_emu, y_emu, cx_emu, cy_emu
    ///         ...
    ///     },
    /// }
    /// ```
    pub fn add_image(&mut self, sheet: &str, image_dict: &Bound<'_, PyAny>) -> PyResult<()> {
        let dict = image_dict
            .cast::<PyDict>()
            .map_err(|_| PyValueError::new_err("image must be a dict"))?;
        let img = dict_to_sheet_image(dict)?;
        let ws = require_sheet(&mut self.inner, sheet)?;
        ws.images.push(img);
        Ok(())
    }

    /// Sprint Μ Pod-α (RFC-046) — queue a chart onto a sheet.
    ///
    /// `chart_dict` shape (built by Python's
    /// ``Worksheet.add_chart``):
    ///
    /// ```python
    /// {
    ///     "kind": "bar" | "line" | "pie" | "doughnut" | "area"
    ///           | "scatter" | "bubble" | "radar",
    ///     "anchor": { "type": "one_cell" | "two_cell" | "absolute", ... },
    ///     "title": { "runs": [{"text": "Sales", "bold": true, ...}],
    ///                "overlay": false } | None,
    ///     "legend": { "position": "r"|"l"|"t"|"b"|"tr",
    ///                 "overlay": false, ... } | None,
    ///     "x_axis": { "kind": "category"|"value"|"date"|"series",
    ///                 "ax_id": 10, "cross_ax": 100,
    ///                 "ax_pos": "b"|"t"|"l"|"r",
    ///                 "orientation": "minMax"|"maxMin",
    ///                 "major_gridlines": false, "minor_gridlines": false,
    ///                 "major_tick_mark": "none"|"in"|"out"|"cross",
    ///                 "title": {...}, "number_format": "0.00",
    ///                 # ValueAxis only:
    ///                 "min": 0.0, "max": 100.0,
    ///                 "major_unit": 10.0, "minor_unit": 1.0,
    ///                 "crosses": "autoZero"|"min"|"max",
    ///                 # CategoryAxis only:
    ///                 "lbl_offset": 100, "lbl_algn": "ctr",
    ///                 # DateAxis only:
    ///                 "base_time_unit": "days"|"months"|"years",
    ///               } | None,
    ///     "y_axis": {...} | None,
    ///     "series": [
    ///         { "idx": 0, "order": 0,
    ///           "title": {"strRef": {"sheet": "Sheet1", "range": "B1"}}
    ///                 | {"literal": "My Series"} | None,
    ///           "categories": {"sheet": "Sheet1", "range": "A2:A6"} | None,
    ///           "values": {"sheet": "Sheet1", "range": "B2:B6"} | None,
    ///           "x_values": {...} | None,
    ///           "bubble_size": {...} | None,
    ///           "graphical_properties": {
    ///               "line_color": "FF000000", "line_width_emu": 12700,
    ///               "line_dash": "solid", "fill_color": "FF0000FF",
    ///               "no_fill": false, "no_line": false,
    ///           } | None,
    ///           "marker": { "symbol": "circle"|"square"|...,
    ///                       "size": 7, "graphical_properties": {...} } | None,
    ///           "data_labels": { "show_val": true, "show_cat_name": true,
    ///                            "show_ser_name": false, "show_percent": false,
    ///                            "show_legend_key": false,
    ///                            "show_bubble_size": false,
    ///                            "position": "outEnd",
    ///                            "number_format": "0.00",
    ///                            "separator": "," } | None,
    ///           "error_bars": [
    ///               { "bar_type": "plus"|"minus"|"both",
    ///                 "val_type": "cust"|"fixedVal"|"percentage"
    ///                          |"stdDev"|"stdErr",
    ///                 "value": 1.5, "no_end_cap": false }
    ///           ],
    ///           "trendlines": [
    ///               { "kind": "linear"|"log"|"power"|"exp"
    ///                       |"poly"|"movingAvg",
    ///                 "order": 2, "period": 3, "forward": 1.0,
    ///                 "backward": 1.0, "display_equation": true,
    ///                 "display_r_squared": true, "name": "fit" }
    ///           ],
    ///           "smooth": true, "invert_if_negative": false,
    ///         },
    ///     ],
    ///     "plot_visible_only": true, "display_blanks_as": "gap"|"span"|"zero",
    ///     "vary_colors": true,
    ///     # Bar:
    ///     "bar_dir": "col"|"bar", "grouping": "clustered"|"stacked"
    ///                                |"percentStacked"|"standard",
    ///     "gap_width": 150, "overlap": -50,
    ///     # Doughnut/Pie:
    ///     "hole_size": 50, "first_slice_ang": 0,
    ///     # Scatter:
    ///     "scatter_style": "line"|"lineMarker"|"marker"|"smooth"
    ///                    |"smoothMarker"|"none",
    ///     # Radar:
    ///     "radar_style": "standard"|"marker"|"filled",
    ///     # Bubble:
    ///     "bubble3d": false, "bubble_scale": 100,
    ///     "show_neg_bubbles": false,
    ///     "smoothing": false, "style": 1,
    /// }
    /// ```
    ///
    /// `anchor_a1` is the A1 reference of the top-left cell where the
    /// chart should be anchored (e.g. `"D2"`). Pod-β-style anchor
    /// dicts inside `chart_dict["anchor"]` override this if present.
    pub fn add_chart_native(
        &mut self,
        sheet: &str,
        chart_dict: &Bound<'_, PyAny>,
        anchor_a1: &str,
    ) -> PyResult<()> {
        let dict = chart_dict
            .cast::<PyDict>()
            .map_err(|_| PyValueError::new_err("chart must be a dict"))?;
        let chart = parse_chart_dict(dict, anchor_a1)?;
        let ws = require_sheet(&mut self.inner, sheet)?;
        ws.charts.push(chart);
        Ok(())
    }
}
