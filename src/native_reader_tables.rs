//! Table (`<tableParts>`) reader logic for native XLSX/XLSB books.

use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList};

use wolfxl_reader::Table;

use crate::native_reader_backend::{NativeXlsbBook, NativeXlsxBook};
use crate::native_reader_filter::serialize_auto_filter;

type PyObject = Py<PyAny>;

pub(crate) fn read_tables_xlsx(
    book: &mut NativeXlsxBook,
    py: Python<'_>,
    sheet: &str,
) -> PyResult<PyObject> {
    let tables = book.ensure_sheet(sheet)?.tables.clone();
    serialize(py, &tables)
}

pub(crate) fn read_tables_xlsb(
    book: &mut NativeXlsbBook,
    py: Python<'_>,
    sheet: &str,
) -> PyResult<PyObject> {
    let tables = book.ensure_sheet(sheet)?.tables.clone();
    serialize(py, &tables)
}

fn serialize(py: Python<'_>, tables: &[Table]) -> PyResult<PyObject> {
    let result = PyList::empty(py);
    for table in tables {
        let d = PyDict::new(py);
        d.set_item("name", &table.name)?;
        d.set_item("ref", &table.ref_range)?;
        d.set_item("header_row", table.header_row)?;
        d.set_item("totals_row", table.totals_row)?;
        d.set_item("comment", table.comment.clone())?;
        d.set_item("table_type", table.table_type.clone())?;
        d.set_item("totals_row_shown", table.totals_row_shown)?;
        match &table.style {
            Some(style) => d.set_item("style", style)?,
            None => d.set_item("style", py.None())?,
        }
        d.set_item("show_first_column", table.show_first_column)?;
        d.set_item("show_last_column", table.show_last_column)?;
        d.set_item("show_row_stripes", table.show_row_stripes)?;
        d.set_item("show_column_stripes", table.show_column_stripes)?;
        d.set_item("columns", table.columns.clone())?;
        d.set_item(
            "auto_filter",
            serialize_auto_filter(py, table.auto_filter.as_ref())?,
        )?;
        result.append(d)?;
    }
    Ok(result.into())
}
