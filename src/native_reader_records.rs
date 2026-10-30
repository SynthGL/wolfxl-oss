//! Streaming-style record dispatch for `read_sheet_records` (xlsx + xlsb).
//! The shape of each record honours the include_* flags so the Python layer
//! can opt in to coordinate, format, style id, and extended format payloads.

use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList};

use wolfxl_reader::{Cell, CellValue};

use crate::native_reader_backend::{NativeXlsbBook, NativeXlsxBook};
use crate::native_reader_cell_helpers::{cell_to_plain, ensure_formula_prefix, native_data_type};
use crate::native_reader_dimensions::{is_merged_subordinate, row_col_to_a1_1based};
use crate::native_reader_styles::gradient_to_pydict;
use crate::native_reader_traits::NativeStyleResolver;

type PyObject = Py<PyAny>;

#[derive(Clone, Copy)]
pub(crate) struct NativeRecordOptions {
    pub data_only: bool,
    pub include_format: bool,
    pub include_empty: bool,
    pub include_formula_blanks: bool,
    pub include_coordinate: bool,
    pub include_style_id: bool,
    pub include_extended_format: bool,
    pub include_cached_formula_value: bool,
}

#[allow(clippy::too_many_arguments)]
pub(crate) fn read_sheet_records_xlsx(
    book: &mut NativeXlsxBook,
    py: Python<'_>,
    sheet: &str,
    cell_range: Option<&str>,
    data_only: bool,
    include_format: bool,
    include_empty: bool,
    include_formula_blanks: bool,
    include_coordinate: bool,
    include_style_id: bool,
    include_extended_format: bool,
    include_cached_formula_value: bool,
) -> PyResult<PyObject> {
    let window = book.resolve_window(sheet, cell_range)?;
    book.ensure_sheet_indexes(sheet)?;
    let _ = book.ensure_sheet(sheet)?;
    let cell_index = book.sheet_cell_indexes.get(sheet).ok_or_else(|| {
        PyErr::new::<PyValueError, _>(format!("Missing cell index for sheet: {sheet}"))
    })?;
    let merged_bounds = book.sheet_merged_bounds.get(sheet).ok_or_else(|| {
        PyErr::new::<PyValueError, _>(format!("Missing merged bounds for sheet: {sheet}"))
    })?;
    let data = book.sheet_cache.get(sheet).ok_or_else(|| {
        PyErr::new::<PyValueError, _>(format!("Missing sheet cache for sheet: {sheet}"))
    })?;
    let native_book = &book.book;
    let options = NativeRecordOptions {
        data_only,
        include_format,
        include_empty,
        include_formula_blanks,
        include_coordinate,
        include_style_id,
        include_extended_format,
        include_cached_formula_value,
    };
    let out = PyList::empty(py);

    if include_empty {
        if let Some((min_row, min_col, max_row, max_col)) = window {
            for row in min_row..=max_row {
                for col in min_col..=max_col {
                    append_native_record(
                        py,
                        &out,
                        native_book,
                        merged_bounds,
                        cell_index.get(&(row, col)).map(|idx| &data.cells[*idx]),
                        row,
                        col,
                        options,
                    )?;
                }
            }
            return Ok(out.into());
        }
    }

    for cell in &data.cells {
        if let Some((min_row, min_col, max_row, max_col)) = window {
            if cell.row < min_row || cell.row > max_row || cell.col < min_col || cell.col > max_col
            {
                continue;
            }
        }
        if !native_record_should_emit(cell, options) {
            continue;
        }
        append_native_record(
            py,
            &out,
            native_book,
            merged_bounds,
            Some(cell),
            cell.row,
            cell.col,
            options,
        )?;
    }
    Ok(out.into())
}

#[allow(clippy::too_many_arguments)]
pub(crate) fn read_sheet_records_xlsb(
    book: &mut NativeXlsbBook,
    py: Python<'_>,
    sheet: &str,
    cell_range: Option<&str>,
    data_only: bool,
    include_format: bool,
    include_empty: bool,
    include_formula_blanks: bool,
    include_coordinate: bool,
    include_style_id: bool,
    include_extended_format: bool,
    include_cached_formula_value: bool,
) -> PyResult<PyObject> {
    let window = book.resolve_window(sheet, cell_range)?;
    book.ensure_sheet_indexes(sheet)?;
    let _ = book.ensure_sheet(sheet)?;
    let cell_index = book.sheet_cell_indexes.get(sheet).ok_or_else(|| {
        PyErr::new::<PyValueError, _>(format!("Missing cell index for sheet: {sheet}"))
    })?;
    let data = book.sheet_cache.get(sheet).ok_or_else(|| {
        PyErr::new::<PyValueError, _>(format!("Missing sheet cache for sheet: {sheet}"))
    })?;
    let native_book = &book.book;
    let options = NativeRecordOptions {
        data_only,
        include_format,
        include_empty,
        include_formula_blanks,
        include_coordinate,
        include_style_id,
        include_extended_format,
        include_cached_formula_value,
    };
    let out = PyList::empty(py);
    if include_empty {
        if let Some((min_row, min_col, max_row, max_col)) = window {
            for row in min_row..=max_row {
                for col in min_col..=max_col {
                    append_native_record(
                        py,
                        &out,
                        native_book,
                        &[],
                        cell_index.get(&(row, col)).map(|idx| &data.cells[*idx]),
                        row,
                        col,
                        options,
                    )?;
                }
            }
            return Ok(out.into());
        }
    }
    for cell in &data.cells {
        if let Some((min_row, min_col, max_row, max_col)) = window {
            if cell.row < min_row || cell.row > max_row || cell.col < min_col || cell.col > max_col
            {
                continue;
            }
        }
        if !native_record_should_emit(cell, options) {
            continue;
        }
        append_native_record(
            py,
            &out,
            native_book,
            &[],
            Some(cell),
            cell.row,
            cell.col,
            options,
        )?;
    }
    Ok(out.into())
}

pub(crate) fn read_sheet_style_ids_xlsx(
    book: &mut NativeXlsxBook,
    sheet: &str,
    cell_range: Option<&str>,
) -> PyResult<Vec<(u32, u32, u32)>> {
    let window = book.resolve_window(sheet, cell_range)?;
    book.ensure_sheet_indexes(sheet)?;
    let _ = book.ensure_sheet(sheet)?;
    let merged_bounds = book.sheet_merged_bounds.get(sheet).ok_or_else(|| {
        PyErr::new::<PyValueError, _>(format!("Missing merged bounds for sheet: {sheet}"))
    })?;
    let data = book.sheet_cache.get(sheet).ok_or_else(|| {
        PyErr::new::<PyValueError, _>(format!("Missing sheet cache for sheet: {sheet}"))
    })?;
    let mut out = Vec::new();

    for cell in &data.cells {
        if let Some((min_row, min_col, max_row, max_col)) = window {
            if cell.row < min_row || cell.row > max_row || cell.col < min_col || cell.col > max_col
            {
                continue;
            }
        }
        let Some(style_id) = cell.style_id else {
            continue;
        };
        if style_id == 0 || is_merged_subordinate(merged_bounds, cell.row, cell.col) {
            continue;
        }
        out.push((cell.row, cell.col, style_id));
    }
    Ok(out)
}

pub(crate) fn native_record_should_emit(cell: &Cell, options: NativeRecordOptions) -> bool {
    let has_formula = cell.formula.is_some();
    let has_value = !matches!(cell.value, CellValue::Empty);
    let should_emit_formula =
        has_formula && !options.data_only && (options.include_formula_blanks || has_value);
    options.include_empty || should_emit_formula || has_value
}

#[allow(clippy::too_many_arguments)]
pub(crate) fn append_native_record<B: NativeStyleResolver>(
    py: Python<'_>,
    out: &Bound<'_, PyList>,
    book: &B,
    merged_bounds: &[(u32, u32, u32, u32)],
    cell: Option<&Cell>,
    row: u32,
    col: u32,
    options: NativeRecordOptions,
) -> PyResult<()> {
    let Some(cell) = cell else {
        if !options.include_empty {
            return Ok(());
        }
        let record = PyDict::new(py);
        record.set_item("row", row)?;
        record.set_item("column", col)?;
        record.set_item("data_type", "blank")?;
        record.set_item("value", py.None())?;
        if options.include_coordinate {
            record.set_item("coordinate", row_col_to_a1_1based(row, col))?;
        }
        out.append(record)?;
        return Ok(());
    };

    if !native_record_should_emit(cell, options) {
        return Ok(());
    }

    let is_merged_subordinate = is_merged_subordinate(merged_bounds, cell.row, cell.col);
    let number_format = if is_merged_subordinate {
        None
    } else {
        cell.style_id
            .and_then(|style_id| book.number_format_for_style_id(style_id))
    };
    let has_formula = cell.formula.is_some();
    let has_cached_value = !matches!(cell.value, CellValue::Empty);
    let should_emit_formula =
        has_formula && !options.data_only && (options.include_formula_blanks || has_cached_value);

    let record = PyDict::new(py);
    record.set_item("row", cell.row)?;
    record.set_item("column", cell.col)?;
    if options.include_coordinate {
        record.set_item("coordinate", &cell.coordinate)?;
    }
    if let Some(formula) = &cell.formula {
        record.set_item("formula", ensure_formula_prefix(formula))?;
        if options.include_cached_formula_value && has_cached_value {
            record.set_item(
                "cached_value",
                cell_to_plain(py, cell, true, number_format, book.date1904())?,
            )?;
        }
    }
    if should_emit_formula {
        let formula = ensure_formula_prefix(cell.formula.as_deref().unwrap_or_default());
        record.set_item("data_type", "formula")?;
        record.set_item("value", formula)?;
    } else if options.data_only && has_formula && !has_cached_value {
        // Like openpyxl, an uncached formula keeps its cell's `t=` type.
        record.set_item("data_type", native_data_type(cell, true, number_format))?;
        record.set_item("value", py.None())?;
    } else {
        record.set_item(
            "data_type",
            native_data_type(cell, options.data_only, number_format),
        )?;
        record.set_item(
            "value",
            cell_to_plain(py, cell, options.data_only, number_format, book.date1904())?,
        )?;
    }

    if options.include_format && !is_merged_subordinate {
        populate_record_format(book, &record, cell.style_id, options)?;
    }
    out.append(record)?;
    Ok(())
}

pub(crate) fn populate_record_format<B: NativeStyleResolver>(
    book: &B,
    record: &Bound<'_, PyDict>,
    style_id: Option<u32>,
    options: NativeRecordOptions,
) -> PyResult<()> {
    let Some(style_id) = style_id else {
        return Ok(());
    };
    if options.include_style_id {
        record.set_item("style_id", style_id)?;
    }
    if style_id == 0 {
        return Ok(());
    }
    if options.include_extended_format {
        if let Some(font) = book.font_for_style_id(style_id) {
            if font.bold {
                record.set_item("bold", true)?;
            }
            if font.italic {
                record.set_item("italic", true)?;
            }
            if let Some(value) = &font.underline {
                record.set_item("underline", value)?;
            }
            if font.strikethrough {
                record.set_item("strikethrough", true)?;
            }
            if let Some(value) = font
                .size
                .filter(|value| (*value - 11.0).abs() > f64::EPSILON)
            {
                record.set_item("font_size", value)?;
            }
        }
        if let Some(fill) = book.fill_for_style_id(style_id) {
            if let Some(value) = &fill.bg_color {
                record.set_item("bg_color", value)?;
            }
            if let Some(grad) = &fill.gradient {
                record.set_item("gradient", gradient_to_pydict(record.py(), grad)?)?;
            }
        }
        if let Some(alignment) = book.alignment_for_style_id(style_id) {
            if let Some(value) = &alignment.horizontal {
                record.set_item("h_align", value)?;
            }
            if let Some(value) = &alignment.vertical {
                record.set_item("v_align", value)?;
            }
            if alignment.wrap_text {
                record.set_item("wrap", true)?;
            }
            if let Some(value) = alignment.text_rotation.filter(|value| *value != 0) {
                record.set_item("rotation", value)?;
            }
            if let Some(value) = alignment.indent.filter(|value| *value != 0) {
                record.set_item("indent", value)?;
            }
        }
        if let Some(border) = book.border_for_style_id(style_id) {
            if let Some(bottom) = &border.bottom {
                record.set_item("bottom_border_style", &bottom.style)?;
                record.set_item("has_bottom_border", true)?;
                if bottom.style == "double" {
                    record.set_item("is_double_underline", true)?;
                }
            }
        }
    }
    if let Some(number_format) = book.number_format_for_style_id(style_id) {
        record.set_item("number_format", number_format)?;
    }
    if let Some(name) = book.named_style_for_style_id(style_id) {
        record.set_item("named_style", name)?;
    }
    Ok(())
}
