//! Dimensions / merged-ranges / used-range / visibility logic for native books.

use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use pyo3::types::PyDict;

use crate::native_reader_backend::{NativeXlsbBook, NativeXlsxBook};
use crate::util::a1_to_row_col;

type PyObject = Py<PyAny>;

pub(crate) fn read_merged_ranges_xlsx(
    book: &mut NativeXlsxBook,
    sheet: &str,
) -> PyResult<Vec<String>> {
    crate::native_reader_merges::read_ranges(book, sheet)
}

pub(crate) fn read_merged_ranges_xlsb(
    book: &mut NativeXlsbBook,
    sheet: &str,
) -> PyResult<Vec<String>> {
    Ok(book.ensure_sheet(sheet)?.merged_ranges.clone())
}

pub(crate) fn read_sheet_visibility_xlsx(
    book: &mut NativeXlsxBook,
    py: Python<'_>,
    sheet: &str,
) -> PyResult<PyObject> {
    let data = book.ensure_sheet(sheet)?;
    serialize_visibility(py, data)
}

pub(crate) fn read_sheet_visibility_xlsb(
    book: &mut NativeXlsbBook,
    py: Python<'_>,
    sheet: &str,
) -> PyResult<PyObject> {
    let data = book.ensure_sheet(sheet)?;
    serialize_visibility(py, data)
}

pub(crate) fn read_column_dimensions_xlsx(
    book: &mut NativeXlsxBook,
    py: Python<'_>,
    sheet: &str,
) -> PyResult<PyObject> {
    let data = book.ensure_sheet(sheet)?;
    serialize_column_dimensions(py, data)
}

pub(crate) fn read_column_dimensions_xlsb(
    book: &mut NativeXlsbBook,
    py: Python<'_>,
    sheet: &str,
) -> PyResult<PyObject> {
    let data = book.ensure_sheet(sheet)?;
    serialize_column_dimensions(py, data)
}

pub(crate) fn read_row_dimensions_xlsx(
    book: &mut NativeXlsxBook,
    py: Python<'_>,
    sheet: &str,
) -> PyResult<PyObject> {
    let data = book.ensure_sheet(sheet)?;
    serialize_row_dimensions(py, data)
}

pub(crate) fn read_row_dimensions_xlsb(
    book: &mut NativeXlsbBook,
    py: Python<'_>,
    sheet: &str,
) -> PyResult<PyObject> {
    let data = book.ensure_sheet(sheet)?;
    serialize_row_dimensions(py, data)
}

fn serialize_visibility(py: Python<'_>, data: &wolfxl_reader::WorksheetData) -> PyResult<PyObject> {
    let d = PyDict::new(py);
    d.set_item("hidden_rows", data.hidden_rows.clone())?;
    d.set_item("hidden_columns", data.hidden_columns.clone())?;
    let row_levels = PyDict::new(py);
    for (row, level) in &data.row_outline_levels {
        row_levels.set_item(*row, *level)?;
    }
    d.set_item("row_outline_levels", row_levels)?;
    let column_levels = PyDict::new(py);
    for (col, level) in &data.column_outline_levels {
        column_levels.set_item(*col, *level)?;
    }
    d.set_item("column_outline_levels", column_levels)?;
    Ok(d.into())
}

fn serialize_row_dimensions(
    py: Python<'_>,
    data: &wolfxl_reader::WorksheetData,
) -> PyResult<PyObject> {
    let d = PyDict::new(py);
    for (row, height) in &data.row_heights {
        let entry = PyDict::new(py);
        if height.custom_height {
            entry.set_item("height", height.height)?;
            entry.set_item("custom_height", true)?;
        } else {
            entry.set_item("height", py.None())?;
            entry.set_item("custom_height", false)?;
        }
        entry.set_item("hidden", data.hidden_rows.contains(row))?;
        let outline_level = data
            .row_outline_levels
            .iter()
            .find_map(|(level_row, level)| (*level_row == *row).then_some(*level))
            .unwrap_or(0);
        entry.set_item("outline_level", outline_level)?;
        d.set_item(*row, entry)?;
    }
    for row in &data.hidden_rows {
        if d.contains(row)? {
            continue;
        }
        let entry = PyDict::new(py);
        entry.set_item("height", py.None())?;
        entry.set_item("custom_height", false)?;
        entry.set_item("hidden", true)?;
        let outline_level = data
            .row_outline_levels
            .iter()
            .find_map(|(level_row, level)| (*level_row == *row).then_some(*level))
            .unwrap_or(0);
        entry.set_item("outline_level", outline_level)?;
        d.set_item(*row, entry)?;
    }
    for (row, level) in &data.row_outline_levels {
        if d.contains(row)? {
            continue;
        }
        let entry = PyDict::new(py);
        entry.set_item("height", py.None())?;
        entry.set_item("custom_height", false)?;
        entry.set_item("hidden", false)?;
        entry.set_item("outline_level", *level)?;
        d.set_item(*row, entry)?;
    }
    Ok(d.into())
}

fn serialize_column_dimensions(
    py: Python<'_>,
    data: &wolfxl_reader::WorksheetData,
) -> PyResult<PyObject> {
    let d = PyDict::new(py);
    for width in &data.column_widths {
        for col in width.min..=width.max {
            let letter = column_letter_1based(col);
            let entry = PyDict::new(py);
            entry.set_item("width", strip_excel_padding(width.width))?;
            entry.set_item("custom_width", width.custom_width)?;
            entry.set_item("hidden", data.hidden_columns.contains(&col))?;
            entry.set_item("min", col)?;
            entry.set_item("max", col)?;
            let outline_level = data
                .column_outline_levels
                .iter()
                .find_map(|(level_col, level)| (*level_col == col).then_some(*level))
                .unwrap_or(0);
            entry.set_item("outline_level", outline_level)?;
            d.set_item(letter, entry)?;
        }
    }
    for col in &data.hidden_columns {
        let letter = column_letter_1based(*col);
        if d.contains(&letter)? {
            continue;
        }
        let entry = PyDict::new(py);
        entry.set_item("width", py.None())?;
        entry.set_item("custom_width", false)?;
        entry.set_item("hidden", true)?;
        entry.set_item("min", *col)?;
        entry.set_item("max", *col)?;
        let outline_level = data
            .column_outline_levels
            .iter()
            .find_map(|(level_col, level)| (*level_col == *col).then_some(*level))
            .unwrap_or(0);
        entry.set_item("outline_level", outline_level)?;
        d.set_item(letter, entry)?;
    }
    for (col, level) in &data.column_outline_levels {
        let letter = column_letter_1based(*col);
        if d.contains(&letter)? {
            continue;
        }
        let entry = PyDict::new(py);
        entry.set_item("width", py.None())?;
        entry.set_item("custom_width", false)?;
        entry.set_item("hidden", false)?;
        entry.set_item("min", *col)?;
        entry.set_item("max", *col)?;
        entry.set_item("outline_level", *level)?;
        d.set_item(letter, entry)?;
    }
    Ok(d.into())
}

fn column_letter_1based(mut col: u32) -> String {
    let mut letters = String::new();
    while col > 0 {
        col -= 1;
        letters.insert(0, (b'A' + (col % 26) as u8) as char);
        col /= 26;
    }
    letters
}

pub(crate) fn parse_range_1based(range: &str) -> Option<(u32, u32, u32, u32)> {
    let clean = range.replace('$', "").to_ascii_uppercase();
    let mut parts = clean.split(':');
    let start = parts.next()?;
    let end = parts.next().unwrap_or(start);
    let (start_row0, start_col0) = a1_to_row_col(start).ok()?;
    let (end_row0, end_col0) = a1_to_row_col(end).ok()?;
    let start_row = start_row0 + 1;
    let start_col = start_col0 + 1;
    let end_row = end_row0 + 1;
    let end_col = end_col0 + 1;
    Some((
        start_row.min(end_row),
        start_col.min(end_col),
        start_row.max(end_row),
        start_col.max(end_col),
    ))
}

pub(crate) fn update_bounds(bounds: &mut Option<(u32, u32, u32, u32)>, row: u32, col: u32) {
    match bounds {
        Some((min_row, min_col, max_row, max_col)) => {
            *min_row = (*min_row).min(row);
            *min_col = (*min_col).min(col);
            *max_row = (*max_row).max(row);
            *max_col = (*max_col).max(col);
        }
        None => *bounds = Some((row, col, row, col)),
    }
}

pub(crate) fn is_merged_subordinate(
    merged_bounds: &[(u32, u32, u32, u32)],
    row: u32,
    col: u32,
) -> bool {
    merged_bounds
        .iter()
        .any(|(min_row, min_col, max_row, max_col)| {
            row >= *min_row
                && row <= *max_row
                && col >= *min_col
                && col <= *max_col
                && !(row == *min_row && col == *min_col)
        })
}

pub(crate) fn row_col_to_a1_1based(row: u32, col: u32) -> String {
    let mut n = col;
    let mut letters = String::new();
    while n > 0 {
        n -= 1;
        letters.insert(0, (b'A' + (n % 26) as u8) as char);
        n /= 26;
    }
    format!("{letters}{row}")
}

pub(crate) fn col_letter_to_index_1based(col: &str) -> PyResult<u32> {
    let mut idx = 0u32;
    for ch in col.chars() {
        if !ch.is_ascii_alphabetic() {
            return Err(PyErr::new::<PyValueError, _>(format!(
                "Invalid column letter: {col}"
            )));
        }
        idx = idx
            .checked_mul(26)
            .and_then(|value| value.checked_add((ch.to_ascii_uppercase() as u8 - b'A' + 1) as u32))
            .ok_or_else(|| {
                PyErr::new::<PyValueError, _>(format!("Invalid column letter: {col}"))
            })?;
    }
    if idx == 0 {
        return Err(PyErr::new::<PyValueError, _>(format!(
            "Invalid column letter: {col}"
        )));
    }
    Ok(idx)
}

pub(crate) fn strip_excel_padding(raw: f64) -> f64 {
    const CALIBRI_WIDTH_PADDING: f64 = 0.83203125;
    const ALT_WIDTH_PADDING: f64 = 0.7109375;
    const WIDTH_TOLERANCE: f64 = 0.0005;

    let frac = raw % 1.0;
    for padding in [CALIBRI_WIDTH_PADDING, ALT_WIDTH_PADDING] {
        if (frac - padding).abs() < WIDTH_TOLERANCE {
            let adjusted = raw - padding;
            if adjusted >= 0.0 {
                return (adjusted * 10000.0).round() / 10000.0;
            }
        }
    }
    (raw * 10000.0).round() / 10000.0
}
