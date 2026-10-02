//! Named-range reader logic for native XLSX/XLSB books.

use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList};
use wolfxl_reader::NamedRange;

use crate::native_reader_backend::{NativeXlsbBook, NativeXlsxBook};

type PyObject = Py<PyAny>;

fn named_range_to_dict<'py>(py: Python<'py>, nr: &NamedRange) -> PyResult<Bound<'py, PyDict>> {
    let d = PyDict::new(py);
    d.set_item("name", &nr.name)?;
    d.set_item("scope", &nr.scope)?;
    d.set_item("local_sheet_id", nr.local_sheet_id)?;
    d.set_item("refers_to", &nr.refers_to)?;
    d.set_item("comment", nr.comment.as_deref())?;
    d.set_item("hidden", nr.hidden)?;
    d.set_item("custom_menu", nr.custom_menu.as_deref())?;
    d.set_item("description", nr.description.as_deref())?;
    d.set_item("help", nr.help.as_deref())?;
    d.set_item("status_bar", nr.status_bar.as_deref())?;
    d.set_item("shortcut_key", nr.shortcut_key.as_deref())?;
    d.set_item("function", nr.function)?;
    d.set_item("function_group_id", nr.function_group_id)?;
    d.set_item("vb_procedure", nr.vb_procedure)?;
    d.set_item("xlm", nr.xlm)?;
    d.set_item("publish_to_server", nr.publish_to_server)?;
    d.set_item("workbook_parameter", nr.workbook_parameter)?;
    Ok(d)
}

pub(crate) fn read_named_ranges_xlsx(
    book: &NativeXlsxBook,
    py: Python<'_>,
    sheet: &str,
) -> PyResult<PyObject> {
    named_ranges_for_sheet(py, &book.sheet_names, book.book.named_ranges(), sheet)
}

pub(crate) fn read_named_ranges_xlsb(
    book: &NativeXlsbBook,
    py: Python<'_>,
    sheet: &str,
) -> PyResult<PyObject> {
    named_ranges_for_sheet(py, &book.sheet_names, book.book.named_ranges(), sheet)
}

/// Workbook-scoped names plus the names whose `localSheetId` is `sheet`.
fn named_ranges_for_sheet(
    py: Python<'_>,
    sheet_names: &[String],
    named_ranges: &[NamedRange],
    sheet: &str,
) -> PyResult<PyObject> {
    let Some(sheet_idx) = sheet_names.iter().position(|name| name == sheet) else {
        return Err(PyErr::new::<PyValueError, _>(format!(
            "Unknown sheet: {sheet}"
        )));
    };
    let result = PyList::empty(py);
    for named_range in named_ranges {
        if named_range.scope == "sheet" && named_range.local_sheet_id != Some(sheet_idx) {
            continue;
        }
        result.append(named_range_to_dict(py, named_range)?)?;
    }
    Ok(result.into())
}
