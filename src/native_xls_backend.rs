// SPDX-License-Identifier: MIT
//! Community XLS Python value adapter, independently authored against the
//! Community NativeXlsBook/WorksheetData interfaces and legacy Python contract.
//! No Commercial reader, binding, BIFF or CFB implementation is used here.

use crate::native_reader_cell_helpers::{
    cell_to_dict, cell_to_plain, SerialDates, SharedStringCache,
};
use crate::native_reader_dimensions::{parse_range_1based, update_bounds};
use crate::util::{a1_to_row_col, cell_blank};
use pyo3::exceptions::{PyIOError, PyNotImplementedError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList};
use std::collections::HashMap;
use wolfxl_reader::{NativeXlsBook as XlsReader, WorksheetData};

type PyObject = Py<PyAny>;
type Bounds = (u32, u32, u32, u32);

#[pyclass(unsendable, module = "wolfxl._rust")]
pub struct NativeXlsBook {
    reader: XlsReader,
    names: Vec<String>,
    path: Option<String>,
    sheets: HashMap<String, WorksheetData>,
    indexes: HashMap<String, HashMap<(u32, u32), usize>>,
    /// Python `str` objects for shared-string values, one per table entry.
    shared_string_values: SharedStringCache,
}

impl NativeXlsBook {
    fn load(&mut self, sheet: &str) -> PyResult<()> {
        if !self.names.iter().any(|name| name == sheet) {
            return Err(PyValueError::new_err(format!("Unknown sheet: {sheet}")));
        }
        if !self.sheets.contains_key(sheet) {
            let data = self
                .reader
                .worksheet(sheet)
                .map_err(|e| PyIOError::new_err(e.to_string()))?;
            let index = data
                .cells
                .iter()
                .enumerate()
                .map(|(i, cell)| ((cell.row, cell.col), i))
                .collect();
            self.indexes.insert(sheet.to_owned(), index);
            self.sheets.insert(sheet.to_owned(), data);
        }
        Ok(())
    }

    fn bounds(&mut self, sheet: &str, range: Option<&str>) -> PyResult<Option<Bounds>> {
        self.load(sheet)?;
        if let Some(range) = range.filter(|r| !r.is_empty()) {
            return parse_range_1based(range)
                .map(Some)
                .ok_or_else(|| PyValueError::new_err(format!("Invalid cell range: {range}")));
        }
        let mut bounds = None;
        for cell in &self.sheets[sheet].cells {
            update_bounds(&mut bounds, cell.row, cell.col);
        }
        Ok(bounds)
    }

    fn value(
        &self,
        py: Python<'_>,
        sheet: &str,
        row: u32,
        col: u32,
        data_only: bool,
    ) -> PyResult<PyObject> {
        match self.indexes[sheet].get(&(row, col)) {
            None => cell_blank(py),
            Some(&index) => {
                let cell = &self.sheets[sheet].cells[index];
                let format = cell
                    .style_id
                    .and_then(|id| self.reader.number_format_for_style_id(id));
                cell_to_dict(
                    py,
                    cell,
                    data_only,
                    format,
                    self.serial_dates(),
                    &self.shared_string_values,
                )
            }
        }
    }

    fn serial_dates(&self) -> SerialDates {
        SerialDates {
            date1904: self.reader.date1904(),
            durations: false,
        }
    }
}

fn no_styles<T>() -> PyResult<T> {
    Err(PyNotImplementedError::new_err(
        "styles not supported for .xls files; use .xlsx for style-aware reads",
    ))
}

#[pymethods]
impl NativeXlsBook {
    #[staticmethod]
    #[pyo3(signature = (path, _permissive = false))]
    fn open(path: &str, _permissive: bool) -> PyResult<Self> {
        let reader = XlsReader::open_path(path).map_err(|e| PyIOError::new_err(e.to_string()))?;
        let names = reader
            .sheet_names()
            .into_iter()
            .map(str::to_owned)
            .collect();
        Ok(Self {
            reader,
            names,
            path: Some(path.to_owned()),
            sheets: HashMap::new(),
            indexes: HashMap::new(),
            shared_string_values: SharedStringCache::default(),
        })
    }

    #[staticmethod]
    #[pyo3(signature = (data, _permissive = false))]
    fn open_from_bytes(data: &[u8], _permissive: bool) -> PyResult<Self> {
        let reader =
            XlsReader::open_bytes(data.to_vec()).map_err(|e| PyIOError::new_err(e.to_string()))?;
        let names = reader
            .sheet_names()
            .into_iter()
            .map(str::to_owned)
            .collect();
        Ok(Self {
            reader,
            names,
            path: None,
            sheets: HashMap::new(),
            indexes: HashMap::new(),
            shared_string_values: SharedStringCache::default(),
        })
    }

    fn sheet_names(&self) -> Vec<String> {
        self.names.clone()
    }
    fn source_path(&self) -> Option<&str> {
        self.path.as_deref()
    }
    fn opened_from_bytes(&self) -> bool {
        self.path.is_none()
    }
    fn read_sheet_dimensions(&mut self, sheet: &str) -> PyResult<Option<(u32, u32)>> {
        Ok(self.bounds(sheet, None)?.map(|(_, _, r, c)| (r, c)))
    }
    fn read_sheet_bounds(&mut self, sheet: &str) -> PyResult<Option<Bounds>> {
        self.bounds(sheet, None)
    }

    #[pyo3(signature = (sheet, a1, data_only = false))]
    fn read_cell_value(
        &mut self,
        py: Python<'_>,
        sheet: &str,
        a1: &str,
        data_only: bool,
    ) -> PyResult<PyObject> {
        let (r, c) = a1_to_row_col(a1).map_err(PyValueError::new_err)?;
        self.load(sheet)?;
        self.value(py, sheet, r + 1, c + 1, data_only)
    }

    #[pyo3(signature = (sheet, cell_range = None, data_only = false))]
    fn read_sheet_values(
        &mut self,
        py: Python<'_>,
        sheet: &str,
        cell_range: Option<&str>,
        data_only: bool,
    ) -> PyResult<PyObject> {
        let rows = PyList::empty(py);
        if let Some((r0, c0, r1, c1)) = self.bounds(sheet, cell_range)? {
            for row in r0..=r1 {
                let values = PyList::empty(py);
                for col in c0..=c1 {
                    values.append(self.value(py, sheet, row, col, data_only)?)?;
                }
                rows.append(values)?;
            }
        }
        Ok(rows.into())
    }

    fn read_sheet_formulas(&mut self, sheet: &str) -> PyResult<HashMap<(u32, u32), String>> {
        self.load(sheet)?;
        Ok(self.sheets[sheet]
            .cells
            .iter()
            .filter_map(|cell| {
                cell.formula
                    .as_ref()
                    .map(|f| ((cell.row - 1, cell.col - 1), f.clone()))
            })
            .collect())
    }

    fn read_cached_formula_values(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        self.load(sheet)?;
        let out = PyDict::new(py);
        for cell in &self.sheets[sheet].cells {
            if cell.formula.is_some() {
                let format = cell
                    .style_id
                    .and_then(|id| self.reader.number_format_for_style_id(id));
                out.set_item(
                    &cell.coordinate,
                    cell_to_plain(
                        py,
                        cell,
                        true,
                        format,
                        self.serial_dates(),
                        &self.shared_string_values,
                    )?,
                )?;
            }
        }
        Ok(out.into())
    }

    fn read_cell_font(&self, _row: u32, _col: u32, _sheet: &str) -> PyResult<()> {
        no_styles()
    }
    fn read_cell_fill(&self, _row: u32, _col: u32, _sheet: &str) -> PyResult<()> {
        no_styles()
    }
    fn read_cell_border(&self, _row: u32, _col: u32, _sheet: &str) -> PyResult<()> {
        no_styles()
    }
    fn read_cell_alignment(&self, _row: u32, _col: u32, _sheet: &str) -> PyResult<()> {
        no_styles()
    }
    fn read_cell_number_format(&self, _row: u32, _col: u32, _sheet: &str) -> PyResult<()> {
        no_styles()
    }
    fn read_column_styles(&self, _sheet: &str) -> PyResult<()> {
        no_styles()
    }
    fn read_row_styles(&self, _sheet: &str) -> PyResult<()> {
        no_styles()
    }
    #[pyo3(signature = (_sheet, _cell_range, _data_only, _include_format))]
    fn read_sheet_records(
        &self,
        _sheet: &str,
        _cell_range: Option<&str>,
        _data_only: bool,
        _include_format: bool,
    ) -> PyResult<()> {
        no_styles()
    }
}
