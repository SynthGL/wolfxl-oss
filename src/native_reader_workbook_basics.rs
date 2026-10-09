//! Workbook-level basics: open constructors, sheet state, print area,
//! print titles.

use std::collections::HashMap;

use pyo3::exceptions::{PyIOError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::PyDict;

use wolfxl_reader::{
    NativeXlsbBook as NativeXlsbReaderBook, NativeXlsxBook as NativeReaderBook, SheetState,
};

use crate::native_reader_backend::{NativeXlsbBook, NativeXlsxBook};

type PyObject = Py<PyAny>;

pub(crate) fn open_xlsx_path(path: &str, permissive: bool) -> PyResult<NativeXlsxBook> {
    let book = NativeReaderBook::open_path_permissive(path, permissive)
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("native xlsx open failed: {e}")))?;
    let sheet_names = book.sheet_names().into_iter().map(str::to_string).collect();
    Ok(NativeXlsxBook {
        book,
        sheet_names,
        sheet_cache: HashMap::new(),
        sheet_cell_indexes: HashMap::new(),
        sheet_merged_bounds: HashMap::new(),
        sheet_merge_metadata: HashMap::new(),
        opened_from_bytes: false,
        source_path: Some(path.to_string()),
    })
}

pub(crate) fn open_xlsx_bytes(data: &[u8], permissive: bool) -> PyResult<NativeXlsxBook> {
    let book = NativeReaderBook::open_bytes_permissive(data.to_vec(), permissive)
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("native xlsx open failed: {e}")))?;
    let sheet_names = book.sheet_names().into_iter().map(str::to_string).collect();
    Ok(NativeXlsxBook {
        book,
        sheet_names,
        sheet_cache: HashMap::new(),
        sheet_cell_indexes: HashMap::new(),
        sheet_merged_bounds: HashMap::new(),
        sheet_merge_metadata: HashMap::new(),
        opened_from_bytes: true,
        source_path: None,
    })
}

pub(crate) fn open_xlsb_path(path: &str) -> PyResult<NativeXlsbBook> {
    let book = NativeXlsbReaderBook::open_path(path)
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("native xlsb open failed: {e}")))?;
    let sheet_names = book.sheet_names().into_iter().map(str::to_string).collect();
    Ok(NativeXlsbBook {
        book,
        sheet_names,
        sheet_cache: HashMap::new(),
        sheet_cell_indexes: HashMap::new(),
    })
}

pub(crate) fn open_xlsb_bytes(data: &[u8]) -> PyResult<NativeXlsbBook> {
    let book = NativeXlsbReaderBook::open_bytes(data.to_vec())
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("native xlsb open failed: {e}")))?;
    let sheet_names = book.sheet_names().into_iter().map(str::to_string).collect();
    Ok(NativeXlsbBook {
        book,
        sheet_names,
        sheet_cache: HashMap::new(),
        sheet_cell_indexes: HashMap::new(),
    })
}

pub(crate) fn read_sheet_state_xlsx(book: &NativeXlsxBook, sheet: &str) -> PyResult<&'static str> {
    if !book.sheet_names.iter().any(|name| name == sheet) {
        return Err(PyErr::new::<PyValueError, _>(format!(
            "Unknown sheet: {sheet}"
        )));
    }
    let state = book
        .book
        .sheet_state(sheet)
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("native sheet state read failed: {e}")))?;
    Ok(sheet_state_to_str(state))
}

pub(crate) fn read_sheet_state_xlsb(book: &NativeXlsbBook, sheet: &str) -> PyResult<&'static str> {
    let state = book
        .book
        .sheet_state(sheet)
        .map_err(|e| PyErr::new::<PyValueError, _>(format!("{e}")))?;
    Ok(sheet_state_to_str(state))
}

fn sheet_state_to_str(state: SheetState) -> &'static str {
    match state {
        SheetState::Visible => "visible",
        SheetState::Hidden => "hidden",
        SheetState::VeryHidden => "veryHidden",
    }
}

pub(crate) fn read_print_area_xlsx(book: &NativeXlsxBook, sheet: &str) -> PyResult<Option<String>> {
    if !book.sheet_names.iter().any(|name| name == sheet) {
        return Err(PyErr::new::<PyValueError, _>(format!(
            "Unknown sheet: {sheet}"
        )));
    }
    Ok(book.book.print_area(sheet).map(str::to_string))
}

pub(crate) fn read_print_area_xlsb(book: &NativeXlsbBook, sheet: &str) -> PyResult<Option<String>> {
    if !book.sheet_names.iter().any(|name| name == sheet) {
        return Err(PyErr::new::<PyValueError, _>(format!(
            "Unknown sheet: {sheet}"
        )));
    }
    Ok(book.book.print_area(sheet).map(str::to_string))
}

pub(crate) fn read_print_titles_xlsx(
    book: &NativeXlsxBook,
    py: Python<'_>,
    sheet: &str,
) -> PyResult<PyObject> {
    if !book.sheet_names.iter().any(|name| name == sheet) {
        return Err(PyErr::new::<PyValueError, _>(format!(
            "Unknown sheet: {sheet}"
        )));
    }
    let Some(titles) = book.book.print_titles(sheet) else {
        return Ok(py.None());
    };
    let d = PyDict::new(py);
    d.set_item("rows", titles.rows.as_deref())?;
    d.set_item("cols", titles.cols.as_deref())?;
    Ok(d.into())
}

pub(crate) fn read_print_titles_xlsb(
    book: &NativeXlsbBook,
    py: Python<'_>,
    sheet: &str,
) -> PyResult<PyObject> {
    if !book.sheet_names.iter().any(|name| name == sheet) {
        return Err(PyErr::new::<PyValueError, _>(format!(
            "Unknown sheet: {sheet}"
        )));
    }
    let Some(titles) = book.book.print_titles(sheet) else {
        return Ok(py.None());
    };
    let d = PyDict::new(py);
    d.set_item("rows", titles.rows.as_deref())?;
    d.set_item("cols", titles.cols.as_deref())?;
    Ok(d.into())
}
