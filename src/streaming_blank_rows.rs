//! Bounded construction of ordinary Python blank-cell coordinate snapshots.
//!
//! Use Python's own allocator and member descriptors: no object-layout casts,
//! recycled mutable cells, or process-global worksheet/type references.

use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use pyo3::types::{PyList, PyTuple, PyType};

/// Allocate a small rectangular batch using the existing Python blank class.
/// Only the three coordinate slots initialized by that class are populated.
#[pyfunction]
pub fn streaming_blank_rows<'py>(
    py: Python<'py>,
    cell_type: &Bound<'py, PyType>,
    worksheet: &Bound<'py, PyAny>,
    first_row: i64,
    row_count: usize,
    min_col: i64,
    max_col: i64,
) -> PyResult<Bound<'py, PyList>> {
    let width = max_col
        .checked_sub(min_col)
        .and_then(|n| n.checked_add(1))
        .filter(|n| *n >= 0)
        .ok_or_else(|| PyValueError::new_err("invalid blank-row column bounds"))?
        as usize;
    if row_count > 1024 || width > 32768 || row_count.saturating_mul(width) > 32768 {
        return Err(PyValueError::new_err(
            "blank-row batch exceeds 32768 cells or 1024 rows",
        ));
    }
    first_row
        .checked_add(row_count as i64)
        .ok_or_else(|| PyValueError::new_err("blank-row bounds overflow"))?;

    let object_type = py.get_type::<PyAny>();
    let allocate = object_type.getattr("__new__")?;
    let set_ws = cell_type.getattr("_ws")?.getattr("__set__")?;
    let set_row = cell_type.getattr("_row")?.getattr("__set__")?;
    let set_col = cell_type.getattr("_col")?.getattr("__set__")?;
    let columns = (min_col..=max_col)
        .map(|col| col.into_pyobject(py).map(Bound::into_any))
        .collect::<Result<Vec<_>, _>>()?;
    let rows = PyList::empty(py);
    for offset in 0..row_count {
        let row = (first_row + offset as i64).into_pyobject(py)?;
        let mut cells = Vec::with_capacity(width);
        for col in &columns {
            let cell = allocate.call1((cell_type,))?;
            set_ws.call1((&cell, worksheet))?;
            set_row.call1((&cell, &row))?;
            set_col.call1((&cell, col))?;
            cells.push(cell);
        }
        rows.append(PyTuple::new(py, cells)?)?;
    }
    Ok(rows)
}
