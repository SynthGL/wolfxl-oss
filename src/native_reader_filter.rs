//! AutoFilter / sort-state reader logic for native XLSX/XLSB books.

use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList};

use wolfxl_reader::{
    AutoFilterInfo, DateGroupItemInfo, FilterColumnInfo, FilterInfo, SortConditionInfo,
    SortStateInfo,
};

use crate::native_reader_backend::{NativeXlsbBook, NativeXlsxBook};

type PyObject = Py<PyAny>;

pub(crate) fn read_auto_filter_xlsx(
    book: &mut NativeXlsxBook,
    py: Python<'_>,
    sheet: &str,
) -> PyResult<PyObject> {
    let auto_filter = book.ensure_sheet(sheet)?.auto_filter.as_ref();
    serialize_auto_filter(py, auto_filter)
}

pub(crate) fn read_auto_filter_xlsb(
    book: &mut NativeXlsbBook,
    py: Python<'_>,
    sheet: &str,
) -> PyResult<PyObject> {
    let auto_filter = book.ensure_sheet(sheet)?.auto_filter.as_ref();
    serialize_auto_filter(py, auto_filter)
}

pub(crate) fn serialize_auto_filter(
    py: Python<'_>,
    auto_filter: Option<&AutoFilterInfo>,
) -> PyResult<PyObject> {
    match auto_filter {
        Some(auto_filter) => {
            let d = PyDict::new(py);
            d.set_item("ref", &auto_filter.ref_range)?;
            let columns = PyList::empty(py);
            for column in &auto_filter.filter_columns {
                columns.append(filter_column_to_py(py, column)?)?;
            }
            d.set_item("filter_columns", columns)?;
            match &auto_filter.sort_state {
                Some(sort_state) => d.set_item("sort_state", sort_state_to_py(py, sort_state)?)?,
                None => d.set_item("sort_state", py.None())?,
            }
            Ok(d.into())
        }
        None => Ok(py.None()),
    }
}

pub(crate) fn filter_column_to_py(py: Python<'_>, column: &FilterColumnInfo) -> PyResult<PyObject> {
    let d = PyDict::new(py);
    d.set_item("col_id", column.col_id)?;
    d.set_item("hidden_button", column.hidden_button)?;
    d.set_item("show_button", column.show_button)?;
    match &column.filter {
        Some(filter) => d.set_item("filter", filter_info_to_py(py, filter)?)?,
        None => d.set_item("filter", py.None())?,
    }
    let date_items = PyList::empty(py);
    for item in &column.date_group_items {
        date_items.append(date_group_item_to_py(py, item)?)?;
    }
    d.set_item("date_group_items", date_items)?;
    Ok(d.into())
}

pub(crate) fn filter_info_to_py(py: Python<'_>, filter: &FilterInfo) -> PyResult<PyObject> {
    let d = PyDict::new(py);
    match filter {
        FilterInfo::Blank => {
            d.set_item("kind", "blank")?;
        }
        FilterInfo::Color { dxf_id, cell_color } => {
            d.set_item("kind", "color")?;
            d.set_item("dxf_id", dxf_id)?;
            d.set_item("cell_color", cell_color)?;
        }
        FilterInfo::Custom { and_, filters } => {
            d.set_item("kind", "custom")?;
            d.set_item("and_", and_)?;
            let out = PyList::empty(py);
            for filter in filters {
                let item = PyDict::new(py);
                item.set_item("operator", &filter.operator)?;
                item.set_item("val", &filter.val)?;
                out.append(item)?;
            }
            d.set_item("filters", out)?;
        }
        FilterInfo::Dynamic {
            filter_type,
            val,
            val_iso,
            max_val_iso,
        } => {
            d.set_item("kind", "dynamic")?;
            d.set_item("type", filter_type)?;
            d.set_item("val", val)?;
            d.set_item("val_iso", val_iso)?;
            d.set_item("max_val_iso", max_val_iso)?;
        }
        FilterInfo::Icon { icon_set, icon_id } => {
            d.set_item("kind", "icon")?;
            d.set_item("icon_set", icon_set)?;
            d.set_item("icon_id", icon_id)?;
        }
        FilterInfo::String { values } => {
            d.set_item("kind", "string")?;
            d.set_item("values", values)?;
        }
        FilterInfo::Top10 {
            top,
            percent,
            val,
            filter_val,
        } => {
            d.set_item("kind", "top10")?;
            d.set_item("top", top)?;
            d.set_item("percent", percent)?;
            d.set_item("val", val)?;
            d.set_item("filter_val", filter_val)?;
        }
    }
    Ok(d.into())
}

pub(crate) fn date_group_item_to_py(
    py: Python<'_>,
    item: &DateGroupItemInfo,
) -> PyResult<PyObject> {
    let d = PyDict::new(py);
    d.set_item("year", item.year)?;
    d.set_item("month", item.month)?;
    d.set_item("day", item.day)?;
    d.set_item("hour", item.hour)?;
    d.set_item("minute", item.minute)?;
    d.set_item("second", item.second)?;
    d.set_item("date_time_grouping", &item.date_time_grouping)?;
    Ok(d.into())
}

pub(crate) fn sort_state_to_py(py: Python<'_>, state: &SortStateInfo) -> PyResult<PyObject> {
    let d = PyDict::new(py);
    let conditions = PyList::empty(py);
    for condition in &state.sort_conditions {
        conditions.append(sort_condition_to_py(py, condition)?)?;
    }
    d.set_item("sort_conditions", conditions)?;
    d.set_item("column_sort", state.column_sort)?;
    d.set_item("case_sensitive", state.case_sensitive)?;
    d.set_item("ref", state.ref_range.as_deref())?;
    Ok(d.into())
}

pub(crate) fn sort_condition_to_py(
    py: Python<'_>,
    condition: &SortConditionInfo,
) -> PyResult<PyObject> {
    let d = PyDict::new(py);
    d.set_item("ref", &condition.ref_range)?;
    d.set_item("descending", condition.descending)?;
    d.set_item("sort_by", &condition.sort_by)?;
    d.set_item("custom_list", condition.custom_list.as_deref())?;
    d.set_item("dxf_id", condition.dxf_id)?;
    d.set_item("icon_set", condition.icon_set.as_deref())?;
    d.set_item("icon_id", condition.icon_id)?;
    Ok(d.into())
}
