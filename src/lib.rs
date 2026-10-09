use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList};

type PyObject = Py<PyAny>;

mod atomic_save;
mod calamine_xlsb_xls_backend;
mod native_reader_backend;
mod native_reader_cell_helpers;
mod native_reader_cf;
mod native_reader_comments;
mod native_reader_dimensions;
mod native_reader_drawings;
mod native_reader_filter;
mod native_reader_hyperlinks;
mod native_reader_merges;
mod native_reader_metadata;
mod native_reader_named_ranges;
mod native_reader_page_setup;
mod native_reader_records;
mod native_reader_sheet_data;
mod native_reader_sheet_view;
mod native_reader_styles;
mod native_reader_tables;
mod native_reader_traits;
mod native_reader_validations;
mod native_reader_workbook_basics;
mod native_writer_anchors;
mod native_writer_autofilter;
mod native_writer_backend;
mod native_writer_cells;
mod native_writer_charts;
mod native_writer_formats;
mod native_writer_images;
mod native_writer_rich_text;
mod native_writer_sheet_features;
mod native_writer_sheet_state;
mod native_writer_streaming;
mod native_writer_workbook;
mod native_writer_workbook_metadata;
mod ooxml_util;
mod streaming;
mod streaming_blank_rows;
mod util;
mod wolfxl;
mod wolfxl_core_bridge;

pub use wolfxl::XlsxPatcher;

#[pyfunction]
fn build_info(py: Python<'_>) -> PyResult<PyObject> {
    let info = PyDict::new(py);
    info.set_item("package", "wolfxl")?;
    info.set_item("package_version", env!("CARGO_PKG_VERSION"))?;

    let enabled = PyList::new(py, ["native-xlsx", "native-xlsb", "calamine-xls", "wolfxl"])?;
    info.set_item("enabled_backends", enabled)?;

    let versions = PyDict::new(py);
    versions.set_item("calamine-xls", option_env!("WOLFXL_DEP_CALAMINE_VERSION"))?;
    info.set_item("backend_versions", versions)?;

    Ok(info.into())
}

#[pyfunction]
fn read_first_sheet_value_summary(py: Python<'_>, path: &str) -> PyResult<PyObject> {
    let summary = wolfxl_reader::read_first_sheet_value_summary_path(path).map_err(|e| {
        pyo3::exceptions::PyIOError::new_err(format!("native value summary failed: {e}"))
    })?;
    let result = PyDict::new(py);
    result.set_item("rows", summary.rows)?;
    result.set_item("numeric_checksum", summary.numeric_checksum)?;
    Ok(result.into())
}

#[pyfunction]
fn read_first_sheet_formula_summary(py: Python<'_>, path: &str) -> PyResult<PyObject> {
    let summary = wolfxl_reader::read_first_sheet_formula_summary_path(path).map_err(|e| {
        pyo3::exceptions::PyIOError::new_err(format!("native formula summary failed: {e}"))
    })?;
    let result = PyDict::new(py);
    result.set_item("rows", summary.rows)?;
    result.set_item(
        "first_col_numeric_checksum",
        summary.first_col_numeric_checksum,
    )?;
    result.set_item("formula_cells", summary.formula_cells)?;
    Ok(result.into())
}

#[pymodule]
fn _rust(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add("__version__", env!("CARGO_PKG_VERSION"))?;
    m.add_function(wrap_pyfunction!(build_info, m)?)?;
    m.add_function(wrap_pyfunction!(read_first_sheet_value_summary, m)?)?;
    m.add_function(wrap_pyfunction!(read_first_sheet_formula_summary, m)?)?;
    m.add_class::<native_reader_backend::NativeXlsxBook>()?;
    m.add_class::<native_reader_backend::NativeXlsbBook>()?;
    m.add_class::<calamine_xlsb_xls_backend::CalamineXlsBook>()?;
    m.add_function(wrap_pyfunction!(
        calamine_xlsb_xls_backend::classify_file_format,
        m
    )?)?;
    m.add_class::<native_writer_backend::NativeWorkbook>()?;
    m.add_function(wrap_pyfunction!(
        native_writer_charts::serialize_chart_dict,
        m
    )?)?;
    // Sprint Ο Pod 1D (RFC-058 §10) — workbook security serializer.
    m.add_function(wrap_pyfunction!(
        native_writer_workbook_metadata::serialize_workbook_security_dict,
        m
    )?)?;
    // Sprint Ν Pod-γ (RFC-047 / RFC-048) — pivot serialisers.
    m.add_function(wrap_pyfunction!(
        wolfxl::pivot::serialize_pivot_cache_dict,
        m
    )?)?;
    m.add_function(wrap_pyfunction!(
        wolfxl::pivot::serialize_pivot_records_dict,
        m
    )?)?;
    m.add_function(wrap_pyfunction!(
        wolfxl::pivot::serialize_pivot_table_dict,
        m
    )?)?;
    // Sprint Ο Pod 1B (RFC-056) — autoFilter serialiser + evaluator.
    m.add_function(wrap_pyfunction!(
        wolfxl::autofilter::serialize_autofilter_dict,
        m
    )?)?;
    m.add_function(wrap_pyfunction!(
        wolfxl::autofilter::evaluate_autofilter,
        m
    )?)?;
    // Sprint Ο Pod 3 (RFC-061) — slicer serialisers.
    m.add_function(wrap_pyfunction!(
        wolfxl::pivot::serialize_slicer_cache_dict,
        m
    )?)?;
    m.add_function(wrap_pyfunction!(wolfxl::pivot::serialize_slicer_dict, m)?)?;
    // G17 / RFC-070 — pivot mutation parsers (modify-mode load path).
    m.add_function(wrap_pyfunction!(
        wolfxl::patcher_pivot_parse::parse_pivot_table_meta,
        m
    )?)?;
    m.add_function(wrap_pyfunction!(
        wolfxl::patcher_pivot_parse::parse_pivot_cache_source,
        m
    )?)?;
    // G18 / RFC-071 — external link parsers (modify-mode load path).
    m.add_function(wrap_pyfunction!(
        wolfxl::external_links::parse_external_link_part,
        m
    )?)?;
    m.add_function(wrap_pyfunction!(
        wolfxl::external_links::parse_external_link_rels,
        m
    )?)?;
    // Sprint Π Pod Π-α (RFC-062) — page breaks + sheet format serialiser.
    m.add_function(wrap_pyfunction!(
        wolfxl::page_breaks::serialize_page_breaks_dict,
        m
    )?)?;
    m.add_class::<streaming::StreamingSheetReader>()?;
    m.add_function(wrap_pyfunction!(streaming_blank_rows::streaming_blank_rows, m)?)?;
    m.add_class::<wolfxl::XlsxPatcher>()?;
    wolfxl_core_bridge::register(m)?;
    Ok(())
}
