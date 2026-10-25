//! Sheet-data reader logic: cell value reads, sheet values (dict + plain),
//! formulas, cached formula values, and row-height/column-width lookups.
//! The cell-value primitives live in `native_reader_cell_helpers`; the
//! native-record dispatch lives in `native_reader_records`.

use std::collections::HashMap;

use pyo3::exceptions::{PyIOError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList, PyTuple};
use pyo3::IntoPyObjectExt;
use wolfxl_reader::{CellValue, ValueCell};

use crate::native_reader_backend::{NativeXlsbBook, NativeXlsxBook};
use crate::native_reader_cell_helpers::{
    cell_to_dict, cell_to_plain, cell_to_plain_unstyled_data_only, formula_to_py, number_to_py,
};
use crate::native_reader_dimensions::{parse_range_1based, update_bounds};
use crate::native_reader_traits::NativeStyleResolver;
use crate::util::{a1_to_row_col, cell_blank};

type PyObject = Py<PyAny>;

// ---------- Cell value reads ----------

pub(crate) fn read_cell_value_xlsx(
    book: &mut NativeXlsxBook,
    py: Python<'_>,
    sheet: &str,
    a1: &str,
    data_only: bool,
) -> PyResult<PyObject> {
    let (row0, col0) = a1_to_row_col(a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
    let row = row0 + 1;
    let col = col0 + 1;
    let cell = {
        book.ensure_sheet_indexes(sheet)?;
        let index = book
            .sheet_cell_indexes
            .get(sheet)
            .and_then(|cells| cells.get(&(row, col)))
            .copied();
        let data = book.ensure_sheet(sheet)?;
        index.map(|idx| data.cells[idx].clone())
    };
    let Some(cell) = cell else {
        return cell_blank(py);
    };
    let number_format = book.number_format_for_cell(&cell);
    cell_to_dict(py, &cell, data_only, number_format, book.book.date1904())
}

pub(crate) fn read_cell_value_xlsb(
    book: &mut NativeXlsbBook,
    py: Python<'_>,
    sheet: &str,
    a1: &str,
    data_only: bool,
) -> PyResult<PyObject> {
    let (row0, col0) = a1_to_row_col(a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
    let row = row0 + 1;
    let col = col0 + 1;
    let cell = {
        book.ensure_sheet_indexes(sheet)?;
        let index = book
            .sheet_cell_indexes
            .get(sheet)
            .and_then(|cells| cells.get(&(row, col)))
            .copied();
        let data = book.ensure_sheet(sheet)?;
        index.map(|idx| data.cells[idx].clone())
    };
    let Some(cell) = cell else {
        return cell_blank(py);
    };
    let number_format = book.number_format_for_cell(&cell);
    cell_to_dict(py, &cell, data_only, number_format, book.book.date1904())
}

// ---------- Sheet values (dict / plain) ----------

pub(crate) fn read_sheet_values_xlsx(
    book: &mut NativeXlsxBook,
    py: Python<'_>,
    sheet: &str,
    cell_range: Option<&str>,
    data_only: bool,
) -> PyResult<PyObject> {
    let (min_row, min_col, max_row, max_col) = match book.resolve_window(sheet, cell_range)? {
        Some(bounds) => bounds,
        None => return Ok(PyList::empty(py).into()),
    };
    book.ensure_sheet_indexes(sheet)?;
    let _ = book.ensure_sheet(sheet)?;
    let cell_index = book.sheet_cell_indexes.get(sheet).ok_or_else(|| {
        PyErr::new::<PyValueError, _>(format!("Missing cell index for sheet: {sheet}"))
    })?;
    let data = book.sheet_cache.get(sheet).ok_or_else(|| {
        PyErr::new::<PyValueError, _>(format!("Missing sheet cache for sheet: {sheet}"))
    })?;
    let native_book = &book.book;
    let outer = PyList::empty(py);
    let date1904 = native_book.date1904();
    for row in min_row..=max_row {
        let inner = PyList::empty(py);
        for col in min_col..=max_col {
            if let Some(cell) = cell_index.get(&(row, col)).map(|idx| &data.cells[*idx]) {
                let number_format = cell
                    .style_id
                    .and_then(|style_id| native_book.number_format_for_style_id(style_id));
                inner.append(cell_to_dict(py, cell, data_only, number_format, date1904)?)?;
            } else {
                inner.append(cell_blank(py)?)?;
            }
        }
        outer.append(inner)?;
    }
    Ok(outer.into())
}

pub(crate) fn read_sheet_values_xlsb(
    book: &mut NativeXlsbBook,
    py: Python<'_>,
    sheet: &str,
    cell_range: Option<&str>,
    data_only: bool,
) -> PyResult<PyObject> {
    let window = book.resolve_window(sheet, cell_range)?;
    let Some((min_row, min_col, max_row, max_col)) = window else {
        return Ok(PyList::empty(py).into());
    };
    book.ensure_sheet_indexes(sheet)?;
    let _ = book.ensure_sheet(sheet)?;
    let cell_index = book.sheet_cell_indexes.get(sheet).ok_or_else(|| {
        PyErr::new::<PyValueError, _>(format!("Missing cell index for sheet: {sheet}"))
    })?;
    let data = book.sheet_cache.get(sheet).ok_or_else(|| {
        PyErr::new::<PyValueError, _>(format!("Missing sheet cache for sheet: {sheet}"))
    })?;
    let native_book = &book.book;
    let date1904 = native_book.date1904();
    let outer = PyList::empty(py);
    for row in min_row..=max_row {
        let inner = PyList::empty(py);
        for col in min_col..=max_col {
            let cell = cell_index.get(&(row, col)).map(|idx| &data.cells[*idx]);
            match cell {
                Some(cell) => {
                    let number_format = cell
                        .style_id
                        .and_then(|style_id| native_book.number_format_for_style_id(style_id));
                    inner.append(cell_to_dict(py, cell, data_only, number_format, date1904)?)?;
                }
                None => inner.append(cell_blank(py)?)?,
            }
        }
        outer.append(inner)?;
    }
    Ok(outer.into())
}

pub(crate) fn read_sheet_values_plain_xlsx(
    book: &mut NativeXlsxBook,
    py: Python<'_>,
    sheet: &str,
    cell_range: Option<&str>,
    data_only: bool,
) -> PyResult<PyObject> {
    if cell_range.is_none() {
        return read_sheet_values_plain_full_xlsx(book, py, sheet, data_only);
    }

    let (min_row, min_col, max_row, max_col) = match book.resolve_window(sheet, cell_range)? {
        Some(bounds) => bounds,
        None => return Ok(PyList::empty(py).into()),
    };
    book.ensure_sheet_indexes(sheet)?;
    let _ = book.ensure_sheet(sheet)?;
    let cell_index = book.sheet_cell_indexes.get(sheet).ok_or_else(|| {
        PyErr::new::<PyValueError, _>(format!("Missing cell index for sheet: {sheet}"))
    })?;
    let data = book.sheet_cache.get(sheet).ok_or_else(|| {
        PyErr::new::<PyValueError, _>(format!("Missing sheet cache for sheet: {sheet}"))
    })?;
    let native_book = &book.book;
    let date1904 = native_book.date1904();
    let outer = PyList::empty(py);
    for row in min_row..=max_row {
        let mut inner: Vec<PyObject> = Vec::with_capacity((max_col - min_col + 1) as usize);
        for col in min_col..=max_col {
            if let Some(cell) = cell_index.get(&(row, col)).map(|idx| &data.cells[*idx]) {
                let number_format = cell
                    .style_id
                    .and_then(|style_id| native_book.number_format_for_style_id(style_id));
                inner.push(cell_to_plain(py, cell, data_only, number_format, date1904)?);
            } else {
                inner.push(py.None());
            }
        }
        outer.append(PyTuple::new(py, inner)?)?;
    }
    Ok(outer.into())
}

fn read_sheet_values_plain_full_xlsx(
    book: &mut NativeXlsxBook,
    py: Python<'_>,
    sheet: &str,
    data_only: bool,
) -> PyResult<PyObject> {
    let _ = book.ensure_sheet(sheet)?;
    let data = book.sheet_cache.get(sheet).ok_or_else(|| {
        PyErr::new::<PyValueError, _>(format!("Missing sheet cache for sheet: {sheet}"))
    })?;
    let Some(analysis) = analyze_sheet_bounds(data) else {
        return Ok(PyList::empty(py).into());
    };
    let (_, _, max_row, max_col) = analysis.bounds;
    let min_row = 1;
    let min_col = 1;

    let row_count = (max_row - min_row + 1) as usize;
    let col_count = (max_col - min_col + 1) as usize;
    let native_book = &book.book;
    let date1904 = native_book.date1904();
    if analysis.dense_a1 {
        if analysis.unstyled {
            if data_only {
                return read_dense_a1_unstyled_data_only_plain_values(py, data, max_row, max_col);
            }
            return read_dense_a1_unstyled_plain_values(py, data, max_row, max_col, data_only);
        }
        return read_dense_a1_plain_values(
            py,
            data,
            native_book,
            max_row,
            max_col,
            data_only,
            date1904,
        );
    }

    let mut grid: Vec<Vec<PyObject>> = (0..row_count)
        .map(|_| (0..col_count).map(|_| py.None()).collect())
        .collect();

    for cell in &data.cells {
        if cell.row < min_row || cell.row > max_row || cell.col < min_col || cell.col > max_col {
            continue;
        }
        let number_format = cell
            .style_id
            .and_then(|style_id| native_book.number_format_for_style_id(style_id));
        grid[(cell.row - min_row) as usize][(cell.col - min_col) as usize] =
            cell_to_plain(py, cell, data_only, number_format, date1904)?;
    }

    let outer = PyList::empty(py);
    for row in grid {
        outer.append(PyTuple::new(py, row)?)?;
    }
    Ok(outer.into())
}

pub(crate) fn read_sheet_value_chunks_plain_xlsx(
    book: &NativeXlsxBook,
    py: Python<'_>,
    sheet: &str,
    chunk_size: usize,
    min_row: Option<u32>,
    max_row: Option<u32>,
    min_col: Option<u32>,
    max_col: Option<u32>,
) -> PyResult<PyObject> {
    if chunk_size < 1 {
        return Err(PyErr::new::<PyValueError, _>(
            "chunk_size must be >= 1".to_string(),
        ));
    }

    let mut cells = book.book.worksheet_value_cells(sheet).map_err(|e| {
        PyErr::new::<PyIOError, _>(format!("native value-only sheet read failed: {e}"))
    })?;
    if cells.is_empty() {
        return Ok(PyList::empty(py).into());
    }
    let row_min = min_row.unwrap_or(1);
    let col_min = min_col.unwrap_or(1);
    let mut layout = analyze_value_cell_layout(&cells, row_min, max_row, col_min, max_col);
    if !layout.sorted {
        cells.sort_by_key(|cell| (cell.row, cell.col));
        layout = analyze_value_cell_layout(&cells, row_min, max_row, col_min, max_col);
    }

    let row_max = layout.row_max;
    if row_max < row_min {
        return Ok(PyList::empty(py).into());
    }
    let col_max = layout.col_max;
    let col_count = if col_max >= col_min {
        (col_max - col_min + 1) as usize
    } else {
        0
    };
    if layout.dense_rectangle {
        return read_dense_value_chunks_plain(py, &cells, chunk_size, col_count);
    }

    let chunks = PyList::empty(py);
    let mut pending_rows: Vec<PyObject> = Vec::with_capacity(chunk_size);
    let mut cell_idx = 0usize;

    for row in row_min..=row_max {
        while cell_idx < cells.len() && cells[cell_idx].row < row {
            cell_idx += 1;
        }
        let mut values: Vec<PyObject> = (0..col_count).map(|_| py.None()).collect();
        let mut scan_idx = cell_idx;
        while scan_idx < cells.len() && cells[scan_idx].row == row {
            let cell = &cells[scan_idx];
            if cell.col >= col_min && cell.col <= col_max {
                values[(cell.col - col_min) as usize] = value_cell_to_plain_unstyled(py, cell)?;
            }
            scan_idx += 1;
        }
        cell_idx = scan_idx;
        pending_rows.push(PyTuple::new(py, values)?.into());
        if pending_rows.len() >= chunk_size {
            chunks.append(PyTuple::new(py, pending_rows)?)?;
            pending_rows = Vec::with_capacity(chunk_size);
        }
    }
    if !pending_rows.is_empty() {
        chunks.append(PyTuple::new(py, pending_rows)?)?;
    }
    Ok(chunks.into())
}

struct ValueCellLayout {
    row_max: u32,
    col_max: u32,
    sorted: bool,
    dense_rectangle: bool,
}

fn analyze_value_cell_layout(
    cells: &[ValueCell],
    row_min: u32,
    max_row: Option<u32>,
    col_min: u32,
    max_col: Option<u32>,
) -> ValueCellLayout {
    let mut sorted = true;
    let mut previous_key: Option<(u32, u32)> = None;
    let mut observed_row_max = row_min.saturating_sub(1);
    for cell in cells {
        let key = (cell.row, cell.col);
        if previous_key.is_some_and(|previous| previous > key) {
            sorted = false;
        }
        previous_key = Some(key);
        if max_row.is_none() && cell.row >= row_min {
            observed_row_max = observed_row_max.max(cell.row);
        }
    }
    let row_max = max_row.unwrap_or(observed_row_max);

    let mut observed_col_max = col_min.saturating_sub(1);
    if max_col.is_none() && row_max >= row_min {
        for cell in cells {
            if cell.row >= row_min && cell.row <= row_max && cell.col >= col_min {
                observed_col_max = observed_col_max.max(cell.col);
            }
        }
    }
    let col_max = max_col.unwrap_or(observed_col_max);

    let dense_rectangle =
        value_cells_cover_dense_rectangle(cells, row_min, row_max, col_min, col_max);
    ValueCellLayout {
        row_max,
        col_max,
        sorted,
        dense_rectangle,
    }
}

fn value_cells_cover_dense_rectangle(
    cells: &[ValueCell],
    row_min: u32,
    row_max: u32,
    col_min: u32,
    col_max: u32,
) -> bool {
    if row_max < row_min || col_max < col_min {
        return false;
    }
    let row_count = (row_max - row_min + 1) as usize;
    let col_count = (col_max - col_min + 1) as usize;
    if row_count.checked_mul(col_count) != Some(cells.len()) {
        return false;
    }
    cells.iter().enumerate().all(|(idx, cell)| {
        let row_offset = idx / col_count;
        let col_offset = idx % col_count;
        cell.row == row_min + row_offset as u32 && cell.col == col_min + col_offset as u32
    })
}

fn read_dense_value_chunks_plain(
    py: Python<'_>,
    cells: &[ValueCell],
    chunk_size: usize,
    col_count: usize,
) -> PyResult<PyObject> {
    let chunks = PyList::empty(py);
    let mut pending_rows: Vec<PyObject> = Vec::with_capacity(chunk_size);
    for row_cells in cells.chunks(col_count) {
        let mut values: Vec<PyObject> = Vec::with_capacity(col_count);
        for cell in row_cells {
            values.push(value_cell_to_plain_unstyled(py, cell)?);
        }
        pending_rows.push(PyTuple::new(py, values)?.into());
        if pending_rows.len() >= chunk_size {
            chunks.append(PyTuple::new(py, pending_rows)?)?;
            pending_rows = Vec::with_capacity(chunk_size);
        }
    }
    if !pending_rows.is_empty() {
        chunks.append(PyTuple::new(py, pending_rows)?)?;
    }
    Ok(chunks.into())
}

fn value_cell_to_plain_unstyled(py: Python<'_>, cell: &ValueCell) -> PyResult<PyObject> {
    match &cell.value {
        CellValue::Empty => Ok(py.None()),
        CellValue::String(s) => Ok(s.clone().into_py_any(py)?),
        CellValue::Number(n) => number_to_py(py, *n, cell.float_text),
        CellValue::Bool(b) => Ok((*b).into_py_any(py)?),
        CellValue::Error(e) => Ok(e.clone().into_py_any(py)?),
    }
}

struct SheetBoundsAnalysis {
    bounds: (u32, u32, u32, u32),
    dense_a1: bool,
    unstyled: bool,
}

fn analyze_sheet_bounds(data: &wolfxl_reader::WorksheetData) -> Option<SheetBoundsAnalysis> {
    let mut bounds: Option<(u32, u32, u32, u32)> = None;
    let mut dense_a1 = true;
    let mut unstyled = true;
    let mut current_dense_row = 1u32;
    let mut next_dense_col = 1u32;
    let mut expected_dense_width: Option<u32> = None;
    for (idx, cell) in data.cells.iter().enumerate() {
        update_bounds(&mut bounds, cell.row, cell.col);
        if !matches!(cell.style_id, None | Some(0)) {
            unstyled = false;
        }
        if !dense_a1 {
            continue;
        }
        if idx == 0 && (cell.row != 1 || cell.col != 1) {
            dense_a1 = false;
            continue;
        }
        if cell.row == current_dense_row && cell.col == next_dense_col {
            next_dense_col += 1;
            continue;
        }
        if cell.row == current_dense_row + 1 && cell.col == 1 {
            let current_width = next_dense_col - 1;
            if let Some(width) = expected_dense_width {
                if current_width != width {
                    dense_a1 = false;
                    continue;
                }
            } else {
                expected_dense_width = Some(current_width);
            }
            current_dense_row = cell.row;
            next_dense_col = 2;
            continue;
        }
        dense_a1 = false;
    }

    for range in &data.merged_ranges {
        if let Some((min_row, min_col, max_row, max_col)) = parse_range_1based(range) {
            update_bounds(&mut bounds, min_row, min_col);
            update_bounds(&mut bounds, max_row, max_col);
        }
    }

    let bounds = match bounds {
        Some(bounds) => bounds,
        None => data.dimension.as_deref().and_then(parse_range_1based)?,
    };
    let (min_row, min_col, max_row, max_col) = bounds;
    if min_row != 1 || min_col != 1 || max_row == 0 || max_col == 0 {
        dense_a1 = false;
    }
    let expected_len = max_row as usize * max_col as usize;
    if data.cells.len() != expected_len {
        dense_a1 = false;
    }
    if let Some(width) = expected_dense_width {
        let final_width = next_dense_col - 1;
        if final_width != width || width != max_col {
            dense_a1 = false;
        }
    }

    Some(SheetBoundsAnalysis {
        bounds,
        dense_a1,
        unstyled,
    })
}

fn read_dense_a1_unstyled_data_only_plain_values(
    py: Python<'_>,
    data: &wolfxl_reader::WorksheetData,
    max_row: u32,
    max_col: u32,
) -> PyResult<PyObject> {
    let row_count = max_row as usize;
    let col_count = max_col as usize;
    let outer = PyList::empty(py);

    for row_idx in 0..row_count {
        let start = row_idx * col_count;
        let end = start + col_count;
        let mut inner: Vec<PyObject> = Vec::with_capacity(col_count);
        for cell in &data.cells[start..end] {
            inner.push(cell_to_plain_unstyled_data_only(py, cell)?);
        }
        outer.append(PyTuple::new(py, inner)?)?;
    }
    Ok(outer.into())
}

fn read_dense_a1_unstyled_plain_values(
    py: Python<'_>,
    data: &wolfxl_reader::WorksheetData,
    max_row: u32,
    max_col: u32,
    data_only: bool,
) -> PyResult<PyObject> {
    let row_count = max_row as usize;
    let col_count = max_col as usize;
    let outer = PyList::empty(py);

    for row_idx in 0..row_count {
        let start = row_idx * col_count;
        let end = start + col_count;
        let mut inner: Vec<PyObject> = Vec::with_capacity(col_count);
        for cell in &data.cells[start..end] {
            inner.push(cell_to_plain(py, cell, data_only, None, false)?);
        }
        outer.append(PyTuple::new(py, inner)?)?;
    }
    Ok(outer.into())
}

fn read_dense_a1_plain_values<B: NativeStyleResolver>(
    py: Python<'_>,
    data: &wolfxl_reader::WorksheetData,
    native_book: &B,
    max_row: u32,
    max_col: u32,
    data_only: bool,
    date1904: bool,
) -> PyResult<PyObject> {
    let row_count = max_row as usize;
    let col_count = max_col as usize;
    let outer = PyList::empty(py);

    for row_idx in 0..row_count {
        let start = row_idx * col_count;
        let end = start + col_count;
        let mut inner: Vec<PyObject> = Vec::with_capacity(col_count);
        for cell in &data.cells[start..end] {
            let number_format = cell
                .style_id
                .and_then(|style_id| native_book.number_format_for_style_id(style_id));
            inner.push(cell_to_plain(py, cell, data_only, number_format, date1904)?);
        }
        outer.append(PyTuple::new(py, inner)?)?;
    }
    Ok(outer.into())
}

pub(crate) fn read_sheet_values_plain_xlsb(
    book: &mut NativeXlsbBook,
    py: Python<'_>,
    sheet: &str,
    cell_range: Option<&str>,
    data_only: bool,
) -> PyResult<PyObject> {
    let window = book.resolve_window(sheet, cell_range)?;
    let Some((min_row, min_col, max_row, max_col)) = window else {
        return Ok(PyList::empty(py).into());
    };
    book.ensure_sheet_indexes(sheet)?;
    let _ = book.ensure_sheet(sheet)?;
    let cell_index = book.sheet_cell_indexes.get(sheet).ok_or_else(|| {
        PyErr::new::<PyValueError, _>(format!("Missing cell index for sheet: {sheet}"))
    })?;
    let data = book.sheet_cache.get(sheet).ok_or_else(|| {
        PyErr::new::<PyValueError, _>(format!("Missing sheet cache for sheet: {sheet}"))
    })?;
    let native_book = &book.book;
    let date1904 = native_book.date1904();
    let outer = PyList::empty(py);
    for row in min_row..=max_row {
        let mut inner: Vec<PyObject> = Vec::with_capacity((max_col - min_col + 1) as usize);
        for col in min_col..=max_col {
            let cell = cell_index.get(&(row, col)).map(|idx| &data.cells[*idx]);
            match cell {
                Some(cell) => {
                    let number_format = cell
                        .style_id
                        .and_then(|style_id| native_book.number_format_for_style_id(style_id));
                    inner.push(cell_to_plain(py, cell, data_only, number_format, date1904)?);
                }
                None => inner.push(py.None()),
            }
        }
        outer.append(PyTuple::new(py, inner)?)?;
    }
    Ok(outer.into())
}

// ---------- Formulas ----------

pub(crate) fn read_sheet_formulas_xlsx(
    book: &mut NativeXlsxBook,
    sheet: &str,
) -> PyResult<HashMap<(u32, u32), String>> {
    let data = book.ensure_sheet(sheet)?;
    Ok(data
        .cells
        .iter()
        .filter_map(|c| {
            c.formula
                .as_ref()
                .map(|f| ((c.row - 1, c.col - 1), f.clone()))
        })
        .collect())
}

pub(crate) fn read_sheet_formulas_xlsb(
    book: &mut NativeXlsbBook,
    sheet: &str,
) -> PyResult<HashMap<(u32, u32), String>> {
    let data = book.ensure_sheet(sheet)?;
    Ok(data
        .cells
        .iter()
        .filter_map(|c| {
            c.formula
                .as_ref()
                .map(|f| ((c.row - 1, c.col - 1), f.clone()))
        })
        .collect())
}

pub(crate) fn read_cell_formula_xlsx(
    book: &mut NativeXlsxBook,
    py: Python<'_>,
    sheet: &str,
    a1: &str,
) -> PyResult<PyObject> {
    let (row0, col0) = a1_to_row_col(a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
    let row = row0 + 1;
    let col = col0 + 1;
    let formula = book
        .ensure_sheet(sheet)?
        .cells
        .iter()
        .find(|cell| cell.row == row && cell.col == col)
        .and_then(|cell| cell.formula.as_deref())
        .map(str::to_string);
    match formula {
        Some(formula) => formula_to_py(py, &formula),
        None => Ok(py.None()),
    }
}

pub(crate) fn read_cached_formula_values_xlsx(
    book: &mut NativeXlsxBook,
    py: Python<'_>,
    sheet: &str,
) -> PyResult<PyObject> {
    let cells = book.ensure_sheet(sheet)?.cells.clone();
    let date1904 = book.book.date1904();
    let out = PyDict::new(py);
    for cell in &cells {
        if cell.formula.is_some() {
            let number_format = book.number_format_for_cell(cell);
            out.set_item(
                &cell.coordinate,
                cell_to_plain(py, cell, true, number_format, date1904)?,
            )?;
        }
    }
    Ok(out.into())
}

pub(crate) fn read_cached_formula_values_xlsb(
    book: &mut NativeXlsbBook,
    py: Python<'_>,
    sheet: &str,
) -> PyResult<PyObject> {
    let cells = book.ensure_sheet(sheet)?.cells.clone();
    let date1904 = book.book.date1904();
    let out = PyDict::new(py);
    for cell in &cells {
        if cell.formula.is_some() {
            let number_format = book.number_format_for_cell(cell);
            out.set_item(
                &cell.coordinate,
                cell_to_plain(py, cell, true, number_format, date1904)?,
            )?;
        }
    }
    Ok(out.into())
}

// ---------- Row height / column width ----------

pub(crate) fn read_row_height_xlsx(
    book: &mut NativeXlsxBook,
    sheet: &str,
    row: i64,
) -> PyResult<Option<f64>> {
    if row < 1 {
        return Ok(None);
    }
    Ok(book
        .ensure_sheet(sheet)?
        .row_heights
        .get(&(row as u32))
        .filter(|height| height.custom_height)
        .map(|height| height.height))
}

pub(crate) fn read_row_height_xlsb(
    book: &mut NativeXlsbBook,
    sheet: &str,
    row: i64,
) -> PyResult<Option<f64>> {
    if row < 1 {
        return Ok(None);
    }
    Ok(book
        .ensure_sheet(sheet)?
        .row_heights
        .get(&(row as u32))
        .filter(|height| height.custom_height)
        .map(|height| height.height))
}

pub(crate) fn read_column_width_xlsx(
    book: &mut NativeXlsxBook,
    sheet: &str,
    col_letter: &str,
) -> PyResult<Option<f64>> {
    let col = crate::native_reader_dimensions::col_letter_to_index_1based(col_letter)?;
    Ok(book
        .ensure_sheet(sheet)?
        .column_widths
        .iter()
        .find(|width| width.custom_width && col >= width.min && col <= width.max)
        .map(|width| crate::native_reader_dimensions::strip_excel_padding(width.width)))
}

pub(crate) fn read_column_width_xlsb(
    book: &mut NativeXlsbBook,
    sheet: &str,
    col_letter: &str,
) -> PyResult<Option<f64>> {
    let col = crate::native_reader_dimensions::col_letter_to_index_1based(col_letter)?;
    Ok(book
        .ensure_sheet(sheet)?
        .column_widths
        .iter()
        .find(|width| width.custom_width && col >= width.min && col <= width.max)
        .map(|width| crate::native_reader_dimensions::strip_excel_padding(width.width)))
}

// ---------- Re-export for the `native_reader_styles` array-formula fallback ----------

pub(crate) use crate::native_reader_cell_helpers::ensure_formula_prefix;
