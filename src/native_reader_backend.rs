//! PyO3 bridge for the native XLSX/XLSB reader. Pymethods here delegate to
//! `native_reader_<feature>` sibling modules; inherent helpers below
//! (`ensure_sheet`, `resolve_window`, etc.) stay here because every feature
//! module needs them. Feature payloads emit openpyxl-compatible dicts; keep
//! new keys additive so older Python hydration can ignore unknown fields.

use std::collections::HashMap;

use pyo3::exceptions::{PyIOError, PyValueError};
use pyo3::prelude::*;

type PyObject = Py<PyAny>;

use crate::native_reader_dimensions::{parse_range_1based, update_bounds};
use crate::util::a1_to_row_col;
use wolfxl_reader::{
    Cell, NativeXlsbBook as NativeXlsbReaderBook, NativeXlsxBook as NativeReaderBook, WorksheetData,
};

#[pyclass(unsendable, module = "wolfxl._rust")]
pub struct NativeXlsxBook {
    pub(crate) book: NativeReaderBook,
    pub(crate) sheet_names: Vec<String>,
    pub(crate) sheet_cache: HashMap<String, WorksheetData>,
    pub(crate) sheet_cell_indexes: HashMap<String, HashMap<(u32, u32), usize>>,
    pub(crate) sheet_merged_bounds: HashMap<String, Vec<(u32, u32, u32, u32)>>,
    pub(crate) sheet_merge_metadata: HashMap<String, crate::native_reader_merges::MergeMetadata>,
    pub(crate) opened_from_bytes: bool,
    pub(crate) source_path: Option<String>,
}

#[pyclass(unsendable, module = "wolfxl._rust")]
pub struct NativeXlsbBook {
    pub(crate) book: NativeXlsbReaderBook,
    pub(crate) sheet_names: Vec<String>,
    pub(crate) sheet_cache: HashMap<String, WorksheetData>,
    pub(crate) sheet_cell_indexes: HashMap<String, HashMap<(u32, u32), usize>>,
}

#[pymethods]
impl NativeXlsxBook {
    /// Open an XLSX/XLSM workbook from a filesystem path.
    #[staticmethod]
    #[pyo3(signature = (path, permissive = false))]
    pub fn open(path: &str, permissive: bool) -> PyResult<Self> {
        crate::native_reader_workbook_basics::open_xlsx_path(path, permissive)
    }

    /// Open an XLSX/XLSM workbook from raw bytes.
    #[staticmethod]
    #[pyo3(signature = (data, permissive = false))]
    pub fn open_from_bytes(data: &[u8], permissive: bool) -> PyResult<Self> {
        crate::native_reader_workbook_basics::open_xlsx_bytes(data, permissive)
    }

    pub fn sheet_names(&self) -> Vec<String> {
        self.sheet_names.clone()
    }

    pub fn chartsheet_names(&self) -> Vec<String> {
        self.book
            .sheets()
            .iter()
            .filter(|sheet| sheet.path.contains("chartsheets/"))
            .map(|sheet| sheet.name.clone())
            .collect()
    }

    pub fn read_chartsheet_charts(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        let charts = self.book.chartsheet_charts(sheet).map_err(|e| {
            PyErr::new::<PyIOError, _>(format!("native chartsheet read failed: {e}"))
        })?;
        let result = pyo3::types::PyList::empty(py);
        for chart in &charts {
            result.append(crate::native_reader_drawings::chart_to_py(py, chart)?)?;
        }
        Ok(result.into())
    }

    pub fn read_sheet_state(&self, sheet: &str) -> PyResult<&'static str> {
        crate::native_reader_workbook_basics::read_sheet_state_xlsx(self, sheet)
    }

    pub fn read_print_area(&self, sheet: &str) -> PyResult<Option<String>> {
        crate::native_reader_workbook_basics::read_print_area_xlsx(self, sheet)
    }

    pub fn read_print_titles(&self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_workbook_basics::read_print_titles_xlsx(self, py, sheet)
    }

    pub fn opened_from_bytes(&self) -> bool {
        self.opened_from_bytes
    }

    pub fn source_path(&self) -> Option<String> {
        self.source_path.clone()
    }

    pub fn read_sheet_bounds(&mut self, sheet: &str) -> PyResult<Option<(u32, u32, u32, u32)>> {
        self.read_bounds_1based(sheet)
    }

    pub fn read_sheet_dimensions(&mut self, sheet: &str) -> PyResult<Option<(u32, u32)>> {
        let Some((_, _, max_row, max_col)) = self.read_bounds_1based(sheet)? else {
            return Ok(None);
        };
        Ok(Some((max_row, max_col)))
    }

    #[pyo3(signature = (sheet, a1, data_only = false))]
    pub fn read_cell_value(
        &mut self,
        py: Python<'_>,
        sheet: &str,
        a1: &str,
        data_only: bool,
    ) -> PyResult<PyObject> {
        crate::native_reader_sheet_data::read_cell_value_xlsx(self, py, sheet, a1, data_only)
    }

    #[pyo3(signature = (sheet, cell_range = None, data_only = false))]
    pub fn read_sheet_values(
        &mut self,
        py: Python<'_>,
        sheet: &str,
        cell_range: Option<&str>,
        data_only: bool,
    ) -> PyResult<PyObject> {
        crate::native_reader_sheet_data::read_sheet_values_xlsx(
            self, py, sheet, cell_range, data_only,
        )
    }

    #[pyo3(signature = (sheet, cell_range = None, data_only = false))]
    pub fn read_sheet_values_plain(
        &mut self,
        py: Python<'_>,
        sheet: &str,
        cell_range: Option<&str>,
        data_only: bool,
    ) -> PyResult<PyObject> {
        crate::native_reader_sheet_data::read_sheet_values_plain_xlsx(
            self, py, sheet, cell_range, data_only,
        )
    }

    #[pyo3(signature = (
        sheet,
        chunk_size = 1024,
        min_row = None,
        max_row = None,
        min_col = None,
        max_col = None,
    ))]
    pub fn read_sheet_value_chunks_plain(
        &self,
        py: Python<'_>,
        sheet: &str,
        chunk_size: usize,
        min_row: Option<u32>,
        max_row: Option<u32>,
        min_col: Option<u32>,
        max_col: Option<u32>,
    ) -> PyResult<PyObject> {
        crate::native_reader_sheet_data::read_sheet_value_chunks_plain_xlsx(
            self, py, sheet, chunk_size, min_row, max_row, min_col, max_col,
        )
    }

    #[pyo3(signature = (
        sheet,
        cell_range = None,
        data_only = false,
        include_format = true,
        include_empty = false,
        include_formula_blanks = true,
        include_coordinate = true,
        include_style_id = true,
        include_extended_format = true,
        include_cached_formula_value = false,
    ))]
    #[allow(clippy::too_many_arguments)]
    pub fn read_sheet_records(
        &mut self,
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
        crate::native_reader_records::read_sheet_records_xlsx(
            self,
            py,
            sheet,
            cell_range,
            data_only,
            include_format,
            include_empty,
            include_formula_blanks,
            include_coordinate,
            include_style_id,
            include_extended_format,
            include_cached_formula_value,
        )
    }

    #[pyo3(signature = (sheet, cell_range = None))]
    pub fn read_sheet_style_ids(
        &mut self,
        sheet: &str,
        cell_range: Option<&str>,
    ) -> PyResult<Vec<(u32, u32, u32)>> {
        crate::native_reader_records::read_sheet_style_ids_xlsx(self, sheet, cell_range)
    }

    pub fn read_format_for_style_id(&self, py: Python<'_>, style_id: u32) -> PyResult<PyObject> {
        crate::native_reader_styles::read_format_for_style_id_xlsx(self, py, style_id)
    }

    pub fn read_sheet_formulas(&mut self, sheet: &str) -> PyResult<HashMap<(u32, u32), String>> {
        crate::native_reader_sheet_data::read_sheet_formulas_xlsx(self, sheet)
    }

    pub fn read_cell_formula(
        &mut self,
        py: Python<'_>,
        sheet: &str,
        a1: &str,
    ) -> PyResult<PyObject> {
        crate::native_reader_sheet_data::read_cell_formula_xlsx(self, py, sheet, a1)
    }

    pub fn read_cached_formula_values(
        &mut self,
        py: Python<'_>,
        sheet: &str,
    ) -> PyResult<PyObject> {
        crate::native_reader_sheet_data::read_cached_formula_values_xlsx(self, py, sheet)
    }

    /// Read styles of arbitrary endpoints without decoding a cell model.
    pub fn read_endpoint_style_ids(
        &mut self,
        sheet: &str,
        positions: Vec<(u32, u32)>,
    ) -> PyResult<Vec<u32>> {
        crate::native_reader_merges::read_sparse_endpoint_style_ids(self, sheet, positions)
    }

    /// Merge-aware border overlay without constructing worksheet cells.
    pub fn read_merged_cell_border(
        &mut self,
        py: Python<'_>,
        sheet: &str,
        row: u32,
        col: u32,
    ) -> PyResult<PyObject> {
        crate::native_reader_merges::read_merged_cell_border(self, py, sheet, row, col)
    }

    pub fn read_merged_endpoint_style_ids(
        &mut self,
        sheet: &str,
    ) -> PyResult<Vec<(String, u32, u32)>> {
        crate::native_reader_merges::read_endpoint_style_ids(self, sheet)
    }

    /// Lazy Python hydration uses this conservative source-negative shortcut.
    pub fn read_merged_ranges_if_present(&mut self, sheet: &str) -> PyResult<Vec<String>> {
        crate::native_reader_merges::read_ranges_if_present(self, sheet)
    }

    pub fn read_merged_ranges(&mut self, sheet: &str) -> PyResult<Vec<String>> {
        crate::native_reader_dimensions::read_merged_ranges_xlsx(self, sheet)
    }

    pub fn read_sheet_visibility(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_dimensions::read_sheet_visibility_xlsx(self, py, sheet)
    }

    pub fn read_hyperlinks(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_hyperlinks::read_hyperlinks_xlsx(self, py, sheet)
    }

    pub fn read_comments(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_comments::read_comments_xlsx(self, py, sheet)
    }

    /// Threaded comments parsed from `xl/threadedComments/threadedCommentsN.xml`.
    /// The Python layer reassembles into a tree by GUID.
    pub fn read_threaded_comments(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_comments::read_threaded_comments_xlsx(self, py, sheet)
    }

    /// Workbook-scoped person list parsed from `xl/persons/personList.xml`.
    /// Insertion-order preserved.
    pub fn read_persons(&self, py: Python<'_>) -> PyResult<PyObject> {
        crate::native_reader_comments::read_persons_xlsx(self, py)
    }

    pub fn read_freeze_panes(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_sheet_view::read_freeze_panes_xlsx(self, py, sheet)
    }

    pub fn read_sheet_view(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_sheet_view::read_sheet_view_xlsx(self, py, sheet)
    }

    pub fn read_sheet_properties(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_sheet_view::read_sheet_properties_xlsx(self, py, sheet)
    }

    pub fn read_conditional_formats(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_cf::read_conditional_formats_xlsx(self, py, sheet)
    }

    pub fn read_data_validations(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_validations::read_data_validations_xlsx(self, py, sheet)
    }

    pub fn read_sheet_protection(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_page_setup::read_sheet_protection_xlsx(self, py, sheet)
    }

    pub fn read_auto_filter(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_filter::read_auto_filter_xlsx(self, py, sheet)
    }

    pub fn read_images(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_drawings::read_images_xlsx(self, py, sheet)
    }

    pub fn read_charts(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_drawings::read_charts_xlsx(self, py, sheet)
    }

    pub fn read_page_margins(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_page_setup::read_page_margins_xlsx(self, py, sheet)
    }

    pub fn read_page_setup(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_page_setup::read_page_setup_xlsx(self, py, sheet)
    }

    pub fn read_print_options(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_page_setup::read_print_options_xlsx(self, py, sheet)
    }

    pub fn read_header_footer(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_page_setup::read_header_footer_xlsx(self, py, sheet)
    }

    pub fn read_row_breaks(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_page_setup::read_row_breaks_xlsx(self, py, sheet)
    }

    pub fn read_column_breaks(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_page_setup::read_column_breaks_xlsx(self, py, sheet)
    }

    pub fn read_sheet_format(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_sheet_view::read_sheet_format_xlsx(self, py, sheet)
    }

    pub fn read_named_ranges(&self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_named_ranges::read_named_ranges_xlsx(self, py, sheet)
    }

    pub fn read_tables(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_tables::read_tables_xlsx(self, py, sheet)
    }

    pub fn read_doc_properties(&self, py: Python<'_>) -> PyResult<PyObject> {
        crate::native_reader_metadata::read_doc_properties(self, py)
    }

    pub fn read_custom_doc_properties(&self, py: Python<'_>) -> PyResult<PyObject> {
        crate::native_reader_metadata::read_custom_doc_properties(self, py)
    }

    pub fn read_workbook_security(&self, py: Python<'_>) -> PyResult<PyObject> {
        crate::native_reader_metadata::read_workbook_security(self, py)
    }

    pub fn read_workbook_properties(&self, py: Python<'_>) -> PyResult<PyObject> {
        crate::native_reader_metadata::read_workbook_properties(self, py)
    }

    pub fn read_calc_properties(&self, py: Python<'_>) -> PyResult<PyObject> {
        crate::native_reader_metadata::read_calc_properties(self, py)
    }

    pub fn read_workbook_views(&self, py: Python<'_>) -> PyResult<PyObject> {
        crate::native_reader_metadata::read_workbook_views(self, py)
    }

    pub fn read_row_height(&mut self, sheet: &str, row: i64) -> PyResult<Option<f64>> {
        crate::native_reader_sheet_data::read_row_height_xlsx(self, sheet, row)
    }

    pub fn read_row_dimensions(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_dimensions::read_row_dimensions_xlsx(self, py, sheet)
    }

    pub fn read_column_width(&mut self, sheet: &str, col_letter: &str) -> PyResult<Option<f64>> {
        crate::native_reader_sheet_data::read_column_width_xlsx(self, sheet, col_letter)
    }

    pub fn read_column_dimensions(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_dimensions::read_column_dimensions_xlsx(self, py, sheet)
    }

    pub fn read_cell_array_formula(
        &mut self,
        py: Python<'_>,
        sheet: &str,
        a1: &str,
    ) -> PyResult<PyObject> {
        crate::native_reader_styles::read_cell_array_formula_xlsx(self, py, sheet, a1)
    }

    pub fn read_sheet_array_formulas(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_styles::read_sheet_array_formulas_xlsx(self, py, sheet)
    }

    pub fn has_sheet_array_formulas(&mut self, sheet: &str) -> PyResult<bool> {
        crate::native_reader_styles::has_sheet_array_formulas_xlsx(self, sheet)
    }

    pub fn read_cell_rich_text(
        &mut self,
        py: Python<'_>,
        sheet: &str,
        a1: &str,
    ) -> PyResult<PyObject> {
        crate::native_reader_styles::read_cell_rich_text_xlsx(self, py, sheet, a1)
    }

    pub fn read_cell_format(
        &mut self,
        py: Python<'_>,
        sheet: &str,
        a1: &str,
    ) -> PyResult<PyObject> {
        crate::native_reader_styles::read_cell_format_xlsx(self, py, sheet, a1)
    }

    pub fn read_cell_border(
        &mut self,
        py: Python<'_>,
        sheet: &str,
        a1: &str,
    ) -> PyResult<PyObject> {
        crate::native_reader_styles::read_cell_border_xlsx(self, py, sheet, a1)
    }
}

#[pymethods]
impl NativeXlsbBook {
    /// Open an XLSB workbook from a filesystem path.
    #[staticmethod]
    #[pyo3(signature = (path, _permissive = false))]
    pub fn open(path: &str, _permissive: bool) -> PyResult<Self> {
        crate::native_reader_workbook_basics::open_xlsb_path(path)
    }

    /// Open an XLSB workbook from raw bytes.
    #[staticmethod]
    #[pyo3(signature = (data, _permissive = false))]
    pub fn open_from_bytes(data: &[u8], _permissive: bool) -> PyResult<Self> {
        crate::native_reader_workbook_basics::open_xlsb_bytes(data)
    }

    pub fn sheet_names(&self) -> Vec<String> {
        self.sheet_names.clone()
    }

    pub fn read_sheet_state(&self, sheet: &str) -> PyResult<&'static str> {
        crate::native_reader_workbook_basics::read_sheet_state_xlsb(self, sheet)
    }

    pub fn read_print_area(&self, sheet: &str) -> PyResult<Option<String>> {
        crate::native_reader_workbook_basics::read_print_area_xlsb(self, sheet)
    }

    pub fn read_print_titles(&self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_workbook_basics::read_print_titles_xlsb(self, py, sheet)
    }

    pub fn read_named_ranges(&self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_named_ranges::read_named_ranges_xlsb(self, py, sheet)
    }

    pub fn read_freeze_panes(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_sheet_view::read_freeze_panes_xlsb(self, py, sheet)
    }

    pub fn read_sheet_view(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_sheet_view::read_sheet_view_xlsb(self, py, sheet)
    }

    pub fn read_sheet_properties(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_sheet_view::read_sheet_properties_xlsb(self, py, sheet)
    }

    pub fn read_hyperlinks(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_hyperlinks::read_hyperlinks_xlsb(self, py, sheet)
    }

    pub fn read_comments(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_comments::read_comments_xlsb(self, py, sheet)
    }

    pub fn read_sheet_protection(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_page_setup::read_sheet_protection_xlsb(self, py, sheet)
    }

    pub fn read_auto_filter(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_filter::read_auto_filter_xlsb(self, py, sheet)
    }

    pub fn read_data_validations(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_validations::read_data_validations_xlsb(self, py, sheet)
    }

    pub fn read_conditional_formats(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_cf::read_conditional_formats_xlsb(self, py, sheet)
    }

    pub fn read_tables(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_tables::read_tables_xlsb(self, py, sheet)
    }

    pub fn read_page_margins(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_page_setup::read_page_margins_xlsb(self, py, sheet)
    }

    pub fn read_page_setup(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_page_setup::read_page_setup_xlsb(self, py, sheet)
    }

    pub fn read_print_options(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_page_setup::read_print_options_xlsb(self, py, sheet)
    }

    pub fn read_images(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_drawings::read_images_xlsb(self, py, sheet)
    }

    pub fn read_charts(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_drawings::read_charts_xlsb(self, py, sheet)
    }

    pub fn read_header_footer(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_page_setup::read_header_footer_xlsb(self, py, sheet)
    }

    pub fn read_row_breaks(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_page_setup::read_row_breaks_xlsb(self, py, sheet)
    }

    pub fn read_column_breaks(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_page_setup::read_column_breaks_xlsb(self, py, sheet)
    }

    pub fn read_sheet_format(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_sheet_view::read_sheet_format_xlsb(self, py, sheet)
    }

    pub fn read_merged_ranges(&mut self, sheet: &str) -> PyResult<Vec<String>> {
        crate::native_reader_dimensions::read_merged_ranges_xlsb(self, sheet)
    }

    pub fn read_sheet_visibility(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_dimensions::read_sheet_visibility_xlsb(self, py, sheet)
    }

    pub fn read_sheet_bounds(&mut self, sheet: &str) -> PyResult<Option<(u32, u32, u32, u32)>> {
        self.read_bounds_1based(sheet)
    }

    pub fn read_sheet_dimensions(&mut self, sheet: &str) -> PyResult<Option<(u32, u32)>> {
        let Some((_, _, max_row, max_col)) = self.read_bounds_1based(sheet)? else {
            return Ok(None);
        };
        Ok(Some((max_row, max_col)))
    }

    #[pyo3(signature = (sheet, a1, data_only = false))]
    pub fn read_cell_value(
        &mut self,
        py: Python<'_>,
        sheet: &str,
        a1: &str,
        data_only: bool,
    ) -> PyResult<PyObject> {
        crate::native_reader_sheet_data::read_cell_value_xlsb(self, py, sheet, a1, data_only)
    }

    #[pyo3(signature = (sheet, cell_range = None, data_only = false))]
    pub fn read_sheet_values(
        &mut self,
        py: Python<'_>,
        sheet: &str,
        cell_range: Option<&str>,
        data_only: bool,
    ) -> PyResult<PyObject> {
        crate::native_reader_sheet_data::read_sheet_values_xlsb(
            self, py, sheet, cell_range, data_only,
        )
    }

    #[pyo3(signature = (sheet, cell_range = None, data_only = false))]
    pub fn read_sheet_values_plain(
        &mut self,
        py: Python<'_>,
        sheet: &str,
        cell_range: Option<&str>,
        data_only: bool,
    ) -> PyResult<PyObject> {
        crate::native_reader_sheet_data::read_sheet_values_plain_xlsb(
            self, py, sheet, cell_range, data_only,
        )
    }

    #[pyo3(signature = (
        sheet,
        cell_range = None,
        data_only = false,
        include_format = true,
        include_empty = false,
        include_formula_blanks = true,
        include_coordinate = true,
        include_style_id = true,
        include_extended_format = true,
        include_cached_formula_value = false,
    ))]
    #[allow(clippy::too_many_arguments)]
    pub fn read_sheet_records(
        &mut self,
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
        crate::native_reader_records::read_sheet_records_xlsb(
            self,
            py,
            sheet,
            cell_range,
            data_only,
            include_format,
            include_empty,
            include_formula_blanks,
            include_coordinate,
            include_style_id,
            include_extended_format,
            include_cached_formula_value,
        )
    }

    pub fn read_sheet_formulas(&mut self, sheet: &str) -> PyResult<HashMap<(u32, u32), String>> {
        crate::native_reader_sheet_data::read_sheet_formulas_xlsb(self, sheet)
    }

    pub fn read_cached_formula_values(
        &mut self,
        py: Python<'_>,
        sheet: &str,
    ) -> PyResult<PyObject> {
        crate::native_reader_sheet_data::read_cached_formula_values_xlsb(self, py, sheet)
    }

    pub fn read_row_height(&mut self, sheet: &str, row: i64) -> PyResult<Option<f64>> {
        crate::native_reader_sheet_data::read_row_height_xlsb(self, sheet, row)
    }

    pub fn read_row_dimensions(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_dimensions::read_row_dimensions_xlsb(self, py, sheet)
    }

    pub fn read_column_width(&mut self, sheet: &str, col_letter: &str) -> PyResult<Option<f64>> {
        crate::native_reader_sheet_data::read_column_width_xlsb(self, sheet, col_letter)
    }

    pub fn read_column_dimensions(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_dimensions::read_column_dimensions_xlsb(self, py, sheet)
    }

    pub fn read_cell_rich_text(
        &mut self,
        py: Python<'_>,
        sheet: &str,
        a1: &str,
    ) -> PyResult<PyObject> {
        crate::native_reader_styles::read_cell_rich_text_xlsb(self, py, sheet, a1)
    }

    pub fn read_cell_array_formula(
        &mut self,
        py: Python<'_>,
        sheet: &str,
        a1: &str,
    ) -> PyResult<PyObject> {
        crate::native_reader_styles::read_cell_array_formula_xlsb(self, py, sheet, a1)
    }

    pub fn read_sheet_array_formulas(&mut self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        crate::native_reader_styles::read_sheet_array_formulas_xlsb(self, py, sheet)
    }

    pub fn has_sheet_array_formulas(&mut self, sheet: &str) -> PyResult<bool> {
        crate::native_reader_styles::has_sheet_array_formulas_xlsb(self, sheet)
    }

    pub fn read_cell_format(
        &mut self,
        py: Python<'_>,
        sheet: &str,
        a1: &str,
    ) -> PyResult<PyObject> {
        crate::native_reader_styles::read_cell_format_xlsb(self, py, sheet, a1)
    }

    pub fn read_cell_border(
        &mut self,
        py: Python<'_>,
        sheet: &str,
        a1: &str,
    ) -> PyResult<PyObject> {
        crate::native_reader_styles::read_cell_border_xlsb(self, py, sheet, a1)
    }
}

// ---------- Inherent helpers shared across feature modules ----------

impl NativeXlsxBook {
    pub(crate) fn ensure_sheet(&mut self, sheet: &str) -> PyResult<&WorksheetData> {
        if !self.sheet_names.iter().any(|name| name == sheet) {
            return Err(PyErr::new::<PyValueError, _>(format!(
                "Unknown sheet: {sheet}"
            )));
        }
        if !self.sheet_cache.contains_key(sheet) {
            let data = self.book.worksheet(sheet).map_err(|e| {
                PyErr::new::<PyIOError, _>(format!("native sheet read failed: {e}"))
            })?;
            self.sheet_cache.insert(sheet.to_string(), data);
        }
        Ok(self.sheet_cache.get(sheet).unwrap())
    }

    pub(crate) fn ensure_sheet_indexes(&mut self, sheet: &str) -> PyResult<()> {
        if self.sheet_cell_indexes.contains_key(sheet)
            && self.sheet_merged_bounds.contains_key(sheet)
        {
            return Ok(());
        }
        let (cell_index, merged_bounds) = {
            let data = self.ensure_sheet(sheet)?;
            let cell_index = data
                .cells
                .iter()
                .enumerate()
                .map(|(idx, cell)| ((cell.row, cell.col), idx))
                .collect();
            let merged_bounds = data
                .merged_ranges
                .iter()
                .filter_map(|range| parse_range_1based(range))
                .collect();
            (cell_index, merged_bounds)
        };
        self.sheet_cell_indexes
            .insert(sheet.to_string(), cell_index);
        self.sheet_merged_bounds
            .insert(sheet.to_string(), merged_bounds);
        Ok(())
    }

    pub(crate) fn read_bounds_1based(
        &mut self,
        sheet: &str,
    ) -> PyResult<Option<(u32, u32, u32, u32)>> {
        let data = self.ensure_sheet(sheet)?;
        let mut bounds: Option<(u32, u32, u32, u32)> = None;
        for cell in &data.cells {
            update_bounds(&mut bounds, cell.row, cell.col);
        }
        for range in &data.merged_ranges {
            if let Some((min_row, min_col, max_row, max_col)) = parse_range_1based(range) {
                update_bounds(&mut bounds, min_row, min_col);
                update_bounds(&mut bounds, max_row, max_col);
            }
        }
        if bounds.is_none() {
            bounds = data.dimension.as_deref().and_then(parse_range_1based);
        }
        Ok(bounds)
    }

    pub(crate) fn resolve_window(
        &mut self,
        sheet: &str,
        cell_range: Option<&str>,
    ) -> PyResult<Option<(u32, u32, u32, u32)>> {
        if let Some(range) = cell_range.filter(|s| !s.is_empty()) {
            return parse_range_1based(range)
                .ok_or_else(|| {
                    PyErr::new::<PyValueError, _>(format!("Invalid cell range: {range}"))
                })
                .map(Some);
        }
        self.read_bounds_1based(sheet)
    }

    pub(crate) fn number_format_for_cell(&self, cell: &Cell) -> Option<&str> {
        cell.style_id
            .and_then(|style_id| self.book.number_format_for_style_id(style_id))
    }
}

impl NativeXlsbBook {
    pub(crate) fn ensure_sheet(&mut self, sheet: &str) -> PyResult<&WorksheetData> {
        if !self.sheet_names.iter().any(|name| name == sheet) {
            return Err(PyErr::new::<PyValueError, _>(format!(
                "Unknown sheet: {sheet}"
            )));
        }
        if !self.sheet_cache.contains_key(sheet) {
            let data = self.book.worksheet(sheet).map_err(|e| {
                PyErr::new::<PyIOError, _>(format!("native xlsb sheet read failed: {e}"))
            })?;
            self.sheet_cache.insert(sheet.to_string(), data);
        }
        Ok(self.sheet_cache.get(sheet).unwrap())
    }

    pub(crate) fn ensure_sheet_indexes(&mut self, sheet: &str) -> PyResult<()> {
        if self.sheet_cell_indexes.contains_key(sheet) {
            return Ok(());
        }
        let cell_index = self
            .ensure_sheet(sheet)?
            .cells
            .iter()
            .enumerate()
            .map(|(idx, cell)| ((cell.row, cell.col), idx))
            .collect();
        self.sheet_cell_indexes
            .insert(sheet.to_string(), cell_index);
        Ok(())
    }

    pub(crate) fn read_bounds_1based(
        &mut self,
        sheet: &str,
    ) -> PyResult<Option<(u32, u32, u32, u32)>> {
        let data = self.ensure_sheet(sheet)?;
        let mut bounds: Option<(u32, u32, u32, u32)> = None;
        for cell in &data.cells {
            update_bounds(&mut bounds, cell.row, cell.col);
        }
        for range in &data.merged_ranges {
            if let Some((min_row, min_col, max_row, max_col)) = parse_range_1based(range) {
                update_bounds(&mut bounds, min_row, min_col);
                update_bounds(&mut bounds, max_row, max_col);
            }
        }
        if bounds.is_none() {
            bounds = data.dimension.as_deref().and_then(parse_range_1based);
        }
        Ok(bounds)
    }

    pub(crate) fn resolve_window(
        &mut self,
        sheet: &str,
        cell_range: Option<&str>,
    ) -> PyResult<Option<(u32, u32, u32, u32)>> {
        if let Some(range) = cell_range.filter(|s| !s.is_empty()) {
            return parse_range_1based(range)
                .ok_or_else(|| {
                    PyErr::new::<PyValueError, _>(format!("Invalid cell range: {range}"))
                })
                .map(Some);
        }
        self.read_bounds_1based(sheet)
    }

    pub(crate) fn number_format_for_cell(&self, cell: &Cell) -> Option<&str> {
        cell.style_id
            .and_then(|style_id| self.book.number_format_for_style_id(style_id))
    }

    pub(crate) fn style_id_for_a1(&mut self, sheet: &str, a1: &str) -> PyResult<Option<u32>> {
        let (row0, col0) = a1_to_row_col(a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
        let row = row0 + 1;
        let col = col0 + 1;
        self.ensure_sheet_indexes(sheet)?;
        let index = self
            .sheet_cell_indexes
            .get(sheet)
            .and_then(|cells| cells.get(&(row, col)))
            .copied();
        let data = self.ensure_sheet(sheet)?;
        Ok(index.and_then(|idx| data.cells[idx].style_id))
    }
}
