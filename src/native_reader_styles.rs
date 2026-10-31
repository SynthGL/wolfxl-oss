//! Styles reader logic: fonts, fills, borders, alignment, named-style and
//! cell-level format readers.

use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList};

use wolfxl_reader::{
    AlignmentInfo, ArrayFormulaInfo, BorderInfo, BorderSide, FillInfo, FontInfo, GradientInfo,
    InlineFontProps,
};

use crate::native_reader_backend::{NativeXlsbBook, NativeXlsxBook};
use crate::native_reader_dimensions::{is_merged_subordinate, row_col_to_a1_1based};
use crate::native_reader_traits::NativeStyleResolver;
use crate::util::a1_to_row_col;

// Silence unused-import lint when the trait is referenced only transitively.
#[allow(dead_code)]
fn _trait_marker<T: NativeStyleResolver>(_: &T) {}

type PyObject = Py<PyAny>;

// ---------- Rich text ----------

pub(crate) fn read_cell_rich_text_xlsx(
    book: &mut NativeXlsxBook,
    py: Python<'_>,
    sheet: &str,
    a1: &str,
) -> PyResult<PyObject> {
    let (row0, col0) = a1_to_row_col(a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
    let row = row0 + 1;
    let col = col0 + 1;
    let runs = {
        let data = book.ensure_sheet(sheet)?;
        data.cells
            .iter()
            .find(|c| c.row == row && c.col == col)
            .and_then(|cell| cell.rich_text.clone())
    };
    serialize_rich_text(py, runs, &book.shared_string_values)
}

pub(crate) fn read_cell_rich_text_xlsb(
    book: &mut NativeXlsbBook,
    py: Python<'_>,
    sheet: &str,
    a1: &str,
) -> PyResult<PyObject> {
    let (row0, col0) = a1_to_row_col(a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
    let row = row0 + 1;
    let col = col0 + 1;
    let runs = {
        let data = book.ensure_sheet(sheet)?;
        data.cells
            .iter()
            .find(|c| c.row == row && c.col == col)
            .and_then(|cell| cell.rich_text.clone())
    };
    serialize_rich_text(py, runs, &book.shared_string_values)
}

fn serialize_rich_text(
    py: Python<'_>,
    runs: Option<std::sync::Arc<[wolfxl_reader::RichTextRun]>>,
    strings: &crate::native_reader_cell_helpers::SharedStringCache,
) -> PyResult<PyObject> {
    let Some(runs) = runs else {
        return Ok(py.None());
    };
    let out = PyList::empty(py);
    for (index, run) in runs.iter().enumerate() {
        let item = PyList::empty(py);
        item.append(strings.rich_run_text(py, &runs, index))?;
        match &run.font {
            Some(font) => item.append(rich_font_to_py(py, font)?)?,
            None => item.append(py.None())?,
        }
        out.append(item)?;
    }
    Ok(out.into())
}

pub(crate) fn rich_font_to_py(py: Python<'_>, font: &InlineFontProps) -> PyResult<PyObject> {
    let d = PyDict::new(py);
    if let Some(value) = font.bold {
        d.set_item("b", value)?;
    }
    if let Some(value) = font.italic {
        d.set_item("i", value)?;
    }
    if let Some(value) = font.strike {
        d.set_item("strike", value)?;
    }
    if let Some(value) = &font.underline {
        d.set_item("u", value)?;
    }
    if let Some(value) = font.size {
        d.set_item("sz", value)?;
    }
    if let Some(value) = &font.color {
        d.set_item("color", value)?;
    }
    if let Some(value) = &font.name {
        d.set_item("rFont", value)?;
    }
    if let Some(value) = font.family {
        d.set_item("family", value)?;
    }
    if let Some(value) = font.charset {
        d.set_item("charset", value)?;
    }
    if let Some(value) = &font.vert_align {
        d.set_item("vertAlign", value)?;
    }
    if let Some(value) = &font.scheme {
        d.set_item("scheme", value)?;
    }
    Ok(d.into())
}

// ---------- Cell array formula ----------

pub(crate) fn read_cell_array_formula_xlsx(
    book: &mut NativeXlsxBook,
    py: Python<'_>,
    sheet: &str,
    a1: &str,
) -> PyResult<PyObject> {
    let (row0, col0) = a1_to_row_col(a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
    let row = row0 + 1;
    let col = col0 + 1;
    let Some(info) = book
        .ensure_sheet(sheet)?
        .array_formulas
        .get(&(row, col))
        .cloned()
    else {
        return Ok(py.None());
    };
    serialize_array_formula(py, info, false)
}

pub(crate) fn read_sheet_array_formulas_xlsx(
    book: &mut NativeXlsxBook,
    py: Python<'_>,
    sheet: &str,
) -> PyResult<PyObject> {
    let data = book.ensure_sheet(sheet)?;
    serialize_sheet_array_formulas(py, &data.array_formulas, false)
}

pub(crate) fn has_sheet_array_formulas_xlsx(
    book: &mut NativeXlsxBook,
    sheet: &str,
) -> PyResult<bool> {
    Ok(!book.ensure_sheet(sheet)?.array_formulas.is_empty())
}

pub(crate) fn read_cell_array_formula_xlsb(
    book: &mut NativeXlsbBook,
    py: Python<'_>,
    sheet: &str,
    a1: &str,
) -> PyResult<PyObject> {
    let (row0, col0) = a1_to_row_col(a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
    let row = row0 + 1;
    let col = col0 + 1;
    let Some(info) = book
        .ensure_sheet(sheet)?
        .array_formulas
        .get(&(row, col))
        .cloned()
    else {
        return Ok(py.None());
    };
    // XLSB applies formula-prefix to array text; XLSX does not.
    serialize_array_formula(py, info, true)
}

pub(crate) fn read_sheet_array_formulas_xlsb(
    book: &mut NativeXlsbBook,
    py: Python<'_>,
    sheet: &str,
) -> PyResult<PyObject> {
    let data = book.ensure_sheet(sheet)?;
    serialize_sheet_array_formulas(py, &data.array_formulas, true)
}

pub(crate) fn has_sheet_array_formulas_xlsb(
    book: &mut NativeXlsbBook,
    sheet: &str,
) -> PyResult<bool> {
    Ok(!book.ensure_sheet(sheet)?.array_formulas.is_empty())
}

fn serialize_sheet_array_formulas(
    py: Python<'_>,
    formulas: &std::collections::HashMap<(u32, u32), ArrayFormulaInfo>,
    prefix_text: bool,
) -> PyResult<PyObject> {
    let out = PyDict::new(py);
    for ((row, col), info) in formulas {
        out.set_item(
            row_col_to_a1_1based(*row, *col),
            serialize_array_formula(py, info.clone(), prefix_text)?,
        )?;
    }
    Ok(out.into())
}

fn serialize_array_formula(
    py: Python<'_>,
    info: ArrayFormulaInfo,
    prefix_text: bool,
) -> PyResult<PyObject> {
    let d = PyDict::new(py);
    match info {
        ArrayFormulaInfo::Array { ref_range, text } => {
            d.set_item("kind", "array")?;
            d.set_item("ref", ref_range)?;
            if prefix_text {
                d.set_item(
                    "text",
                    crate::native_reader_sheet_data::ensure_formula_prefix(&text),
                )?;
            } else {
                d.set_item("text", text)?;
            }
        }
        ArrayFormulaInfo::DataTable {
            ref_range,
            ca,
            dt2_d,
            dtr,
            r1,
            r2,
            del1,
            del2,
        } => {
            d.set_item("kind", "data_table")?;
            d.set_item("ref", ref_range)?;
            d.set_item("ca", ca)?;
            d.set_item("dt2D", dt2_d)?;
            d.set_item("dtr", dtr)?;
            d.set_item("r1", r1)?;
            d.set_item("r2", r2)?;
            d.set_item("del1", del1)?;
            d.set_item("del2", del2)?;
        }
        ArrayFormulaInfo::SpillChild => {
            d.set_item("kind", "spill_child")?;
        }
    }
    Ok(d.into())
}

// ---------- Cell format ----------

pub(crate) fn read_cell_format_xlsx(
    book: &mut NativeXlsxBook,
    py: Python<'_>,
    sheet: &str,
    a1: &str,
) -> PyResult<PyObject> {
    let (row0, col0) = a1_to_row_col(a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
    let row = row0 + 1;
    let col = col0 + 1;
    let style_id = {
        book.ensure_sheet_indexes(sheet)?;
        if is_merged_subordinate(
            book.sheet_merged_bounds
                .get(sheet)
                .map(Vec::as_slice)
                .unwrap_or(&[]),
            row,
            col,
        ) {
            return Ok(PyDict::new(py).into());
        }
        let index = book
            .sheet_cell_indexes
            .get(sheet)
            .and_then(|cells| cells.get(&(row, col)))
            .copied();
        let data = book.ensure_sheet(sheet)?;
        index.and_then(|idx| data.cells[idx].style_id)
    };
    let d = PyDict::new(py);
    if let Some(style_id) = style_id {
        if style_id == 0 {
            return Ok(d.into());
        }
        if let Some(font) = book.book.font_for_style_id(style_id) {
            populate_font(&d, font)?;
        }
        if let Some(fill) = book.book.fill_for_style_id(style_id) {
            populate_fill(&d, fill)?;
        }
        if let Some(number_format) = book.book.number_format_for_style_id(style_id) {
            d.set_item("number_format", number_format)?;
        }
        if let Some(alignment) = book.book.alignment_for_style_id(style_id) {
            populate_alignment(&d, alignment)?;
        }
        if let Some(protection) = book.book.protection_for_style_id(style_id) {
            d.set_item("locked", protection.locked)?;
            d.set_item("hidden", protection.hidden)?;
        }
        if let Some(name) = book.book.named_style_for_style_id(style_id) {
            d.set_item("named_style", name)?;
        }
    }
    Ok(d.into())
}

pub(crate) fn read_format_for_style_id_xlsx(
    book: &NativeXlsxBook,
    py: Python<'_>,
    style_id: u32,
) -> PyResult<PyObject> {
    let d = PyDict::new(py);
    if style_id == 0 {
        return Ok(d.into());
    }
    if let Some(font) = book.book.font_for_style_id(style_id) {
        populate_font(&d, font)?;
    }
    if let Some(fill) = book.book.fill_for_style_id(style_id) {
        populate_fill(&d, fill)?;
    }
    if let Some(number_format) = book.book.number_format_for_style_id(style_id) {
        d.set_item("number_format", number_format)?;
    }
    if let Some(alignment) = book.book.alignment_for_style_id(style_id) {
        populate_alignment(&d, alignment)?;
    }
    if let Some(protection) = book.book.protection_for_style_id(style_id) {
        d.set_item("locked", protection.locked)?;
        d.set_item("hidden", protection.hidden)?;
    }
    if let Some(name) = book.book.named_style_for_style_id(style_id) {
        d.set_item("named_style", name)?;
    }
    Ok(d.into())
}

pub(crate) fn read_cell_format_xlsb(
    book: &mut NativeXlsbBook,
    py: Python<'_>,
    sheet: &str,
    a1: &str,
) -> PyResult<PyObject> {
    let style_id = book.style_id_for_a1(sheet, a1)?;
    let d = PyDict::new(py);
    if let Some(style_id) = style_id {
        if let Some(font) = book.book.font_for_style_id(style_id) {
            populate_font(&d, font)?;
        }
        if let Some(fill) = book.book.fill_for_style_id(style_id) {
            populate_fill(&d, fill)?;
        }
        if let Some(number_format) = book.book.number_format_for_style_id(style_id) {
            d.set_item("number_format", number_format)?;
        }
        if let Some(alignment) = book.book.alignment_for_style_id(style_id) {
            populate_alignment(&d, alignment)?;
        }
        if let Some(protection) = book.book.protection_for_style_id(style_id) {
            d.set_item("locked", protection.locked)?;
            d.set_item("hidden", protection.hidden)?;
        }
    }
    Ok(d.into())
}

// ---------- Cell border ----------

pub(crate) fn read_cell_border_xlsx(
    book: &mut NativeXlsxBook,
    py: Python<'_>,
    sheet: &str,
    a1: &str,
) -> PyResult<PyObject> {
    let (row0, col0) = a1_to_row_col(a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
    let row = row0 + 1;
    let col = col0 + 1;
    let style_id = {
        book.ensure_sheet_indexes(sheet)?;
        if is_merged_subordinate(
            book.sheet_merged_bounds
                .get(sheet)
                .map(Vec::as_slice)
                .unwrap_or(&[]),
            row,
            col,
        ) {
            return Ok(PyDict::new(py).into());
        }
        let index = book
            .sheet_cell_indexes
            .get(sheet)
            .and_then(|cells| cells.get(&(row, col)))
            .copied();
        let data = book.ensure_sheet(sheet)?;
        index.and_then(|idx| data.cells[idx].style_id)
    };
    let d = PyDict::new(py);
    if let Some(border) = style_id.and_then(|id| book.book.border_for_style_id(id)) {
        populate_border(py, &d, border)?;
    }
    Ok(d.into())
}

pub(crate) fn read_cell_border_xlsb(
    book: &mut NativeXlsbBook,
    py: Python<'_>,
    sheet: &str,
    a1: &str,
) -> PyResult<PyObject> {
    let d = PyDict::new(py);
    if let Some(border) = book
        .style_id_for_a1(sheet, a1)?
        .and_then(|id| book.book.border_for_style_id(id))
    {
        populate_border(py, &d, border)?;
    }
    Ok(d.into())
}

// ---------- Style populators ----------

pub(crate) fn populate_font(d: &Bound<'_, PyDict>, font: &FontInfo) -> PyResult<()> {
    if font.bold {
        d.set_item("bold", true)?;
    }
    if font.italic {
        d.set_item("italic", true)?;
    }
    if let Some(value) = &font.underline {
        d.set_item("underline", value)?;
    }
    if font.strikethrough {
        d.set_item("strikethrough", true)?;
    }
    if let Some(value) = &font.name {
        d.set_item("font_name", value)?;
    }
    if let Some(value) = font.size {
        d.set_item("font_size", value)?;
    }
    if let Some(value) = &font.color {
        d.set_item("font_color", value)?;
    }
    Ok(())
}

pub(crate) fn populate_fill(d: &Bound<'_, PyDict>, fill: &FillInfo) -> PyResult<()> {
    if let Some(value) = &fill.bg_color {
        d.set_item("bg_color", value)?;
    }
    if let Some(grad) = &fill.gradient {
        d.set_item("gradient", gradient_to_pydict(d.py(), grad)?)?;
    }
    Ok(())
}

pub(crate) fn gradient_to_pydict<'py>(
    py: Python<'py>,
    grad: &GradientInfo,
) -> PyResult<Bound<'py, PyDict>> {
    let d = PyDict::new(py);
    d.set_item("type", &grad.gradient_type)?;
    if let Ok(v) = grad.degree.parse::<f64>() {
        d.set_item("degree", v)?;
    }
    if let Ok(v) = grad.left.parse::<f64>() {
        d.set_item("left", v)?;
    }
    if let Ok(v) = grad.right.parse::<f64>() {
        d.set_item("right", v)?;
    }
    if let Ok(v) = grad.top.parse::<f64>() {
        d.set_item("top", v)?;
    }
    if let Ok(v) = grad.bottom.parse::<f64>() {
        d.set_item("bottom", v)?;
    }
    let stops = PyList::empty(py);
    for stop in &grad.stops {
        let s = PyDict::new(py);
        if let Ok(v) = stop.position.parse::<f64>() {
            s.set_item("position", v)?;
        }
        if let Some(color) = &stop.color {
            s.set_item("color", color)?;
        }
        stops.append(s)?;
    }
    d.set_item("stops", stops)?;
    Ok(d)
}

pub(crate) fn populate_alignment(d: &Bound<'_, PyDict>, alignment: &AlignmentInfo) -> PyResult<()> {
    if let Some(value) = &alignment.horizontal {
        d.set_item("h_align", value)?;
    }
    if let Some(value) = &alignment.vertical {
        d.set_item("v_align", value)?;
    }
    if alignment.wrap_text {
        d.set_item("wrap", true)?;
    }
    if let Some(value) = alignment.text_rotation.filter(|value| *value != 0) {
        d.set_item("rotation", value)?;
    }
    if let Some(value) = alignment.indent.filter(|value| *value != 0) {
        d.set_item("indent", value)?;
    }
    Ok(())
}

pub(crate) fn populate_border(
    py: Python<'_>,
    d: &Bound<'_, PyDict>,
    border: &BorderInfo,
) -> PyResult<()> {
    if let Some(side) = &border.left {
        set_border_side(py, d, "left", side)?;
    }
    if let Some(side) = &border.right {
        set_border_side(py, d, "right", side)?;
    }
    if let Some(side) = &border.top {
        set_border_side(py, d, "top", side)?;
    }
    if let Some(side) = &border.bottom {
        set_border_side(py, d, "bottom", side)?;
    }
    if let Some(side) = &border.diagonal_up {
        set_border_side(py, d, "diagonal_up", side)?;
    }
    if let Some(side) = &border.diagonal_down {
        set_border_side(py, d, "diagonal_down", side)?;
    }
    Ok(())
}

pub(crate) fn set_border_side(
    py: Python<'_>,
    d: &Bound<'_, PyDict>,
    key: &str,
    side: &BorderSide,
) -> PyResult<()> {
    let edge = PyDict::new(py);
    edge.set_item("style", &side.style)?;
    edge.set_item("color", &side.color)?;
    d.set_item(key, edge)?;
    Ok(())
}
