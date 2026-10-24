//! Drawings reader logic: images, charts, anchor positioning.

use pyo3::prelude::*;
use pyo3::types::{PyBytes, PyDict, PyList};

use wolfxl_reader::{
    AnchorExtentInfo, AnchorMarkerInfo, AnchorPositionInfo, ChartAxisInfo, ChartDataLabelsInfo,
    ChartErrorBarsInfo, ChartGraphicalPropertiesInfo, ChartInfo, ChartSeriesInfo,
    ChartTrendlineInfo, ImageAnchorInfo, ImageInfo,
};

use crate::native_reader_backend::{NativeXlsbBook, NativeXlsxBook};

type PyObject = Py<PyAny>;

pub(crate) fn read_images_xlsx(
    book: &mut NativeXlsxBook,
    py: Python<'_>,
    sheet: &str,
) -> PyResult<PyObject> {
    let images = book.ensure_sheet(sheet)?.images.clone();
    let result = PyList::empty(py);
    for image in &images {
        result.append(image_to_py(py, image)?)?;
    }
    Ok(result.into())
}

pub(crate) fn read_images_xlsb(
    book: &mut NativeXlsbBook,
    py: Python<'_>,
    sheet: &str,
) -> PyResult<PyObject> {
    let images = book.ensure_sheet(sheet)?.images.clone();
    let result = PyList::empty(py);
    for image in &images {
        result.append(image_to_py(py, image)?)?;
    }
    Ok(result.into())
}

pub(crate) fn read_charts_xlsx(
    book: &mut NativeXlsxBook,
    py: Python<'_>,
    sheet: &str,
) -> PyResult<PyObject> {
    let charts = book.ensure_sheet(sheet)?.charts.clone();
    let result = PyList::empty(py);
    for chart in &charts {
        result.append(chart_to_py(py, chart)?)?;
    }
    Ok(result.into())
}

pub(crate) fn read_charts_xlsb(
    book: &mut NativeXlsbBook,
    py: Python<'_>,
    sheet: &str,
) -> PyResult<PyObject> {
    let charts = book.ensure_sheet(sheet)?.charts.clone();
    let result = PyList::empty(py);
    for chart in &charts {
        result.append(chart_to_py(py, chart)?)?;
    }
    Ok(result.into())
}

pub(crate) fn image_to_py(py: Python<'_>, image: &ImageInfo) -> PyResult<PyObject> {
    let d = PyDict::new(py);
    d.set_item("data", PyBytes::new(py, &image.data))?;
    d.set_item("ext", &image.ext)?;
    d.set_item("anchor", image_anchor_to_py(py, &image.anchor)?)?;
    Ok(d.into())
}

/// Serialize native chart metadata into the Python hydration contract.
///
/// The keys mirror openpyxl-facing attributes in
/// `python/wolfxl/_worksheet_media.py`; keep this payload additive.
pub(crate) fn chart_to_py(py: Python<'_>, chart: &ChartInfo) -> PyResult<PyObject> {
    let d = PyDict::new(py);
    d.set_item("source_drawing_path", chart.source_drawing_path.as_deref())?;
    d.set_item("source_chart_rid", chart.source_chart_rid.as_deref())?;
    d.set_item("source_chart_path", chart.source_chart_path.as_deref())?;
    d.set_item("kind", &chart.kind)?;
    d.set_item("title", chart.title.as_deref())?;
    d.set_item("x_axis_title", chart.x_axis_title.as_deref())?;
    d.set_item("y_axis_title", chart.y_axis_title.as_deref())?;
    d.set_item("x_axis", chart_axis_to_py(py, chart.x_axis.as_ref())?)?;
    d.set_item("y_axis", chart_axis_to_py(py, chart.y_axis.as_ref())?)?;
    d.set_item(
        "data_labels",
        chart_data_labels_to_py(py, chart.data_labels.as_ref())?,
    )?;
    d.set_item("legend_position", chart.legend_position.as_deref())?;
    d.set_item("bar_dir", chart.bar_dir.as_deref())?;
    d.set_item("grouping", chart.grouping.as_deref())?;
    d.set_item("scatter_style", chart.scatter_style.as_deref())?;
    d.set_item("vary_colors", chart.vary_colors)?;
    d.set_item("style", chart.style)?;
    d.set_item("anchor", image_anchor_to_py(py, &chart.anchor)?)?;
    let series = PyList::empty(py);
    for item in &chart.series {
        series.append(chart_series_to_py(py, item)?)?;
    }
    d.set_item("series", series)?;
    // `subcharts[0]` is the primary family the fields above describe; a
    // combination chart lists its other families after it.
    let subcharts = PyList::empty(py);
    for item in &chart.subcharts {
        let payload = PyDict::new(py);
        payload.set_item("kind", &item.kind)?;
        payload.set_item("bar_dir", item.bar_dir.as_deref())?;
        payload.set_item("grouping", item.grouping.as_deref())?;
        payload.set_item("scatter_style", item.scatter_style.as_deref())?;
        payload.set_item("vary_colors", item.vary_colors)?;
        payload.set_item("x_axis", chart_axis_to_py(py, item.x_axis.as_ref())?)?;
        payload.set_item("y_axis", chart_axis_to_py(py, item.y_axis.as_ref())?)?;
        let item_series = PyList::empty(py);
        for series_item in &item.series {
            item_series.append(chart_series_to_py(py, series_item)?)?;
        }
        payload.set_item("series", item_series)?;
        subcharts.append(payload)?;
    }
    d.set_item("subcharts", subcharts)?;
    Ok(d.into())
}

pub(crate) fn chart_axis_to_py(py: Python<'_>, axis: Option<&ChartAxisInfo>) -> PyResult<PyObject> {
    let Some(axis) = axis else {
        return Ok(py.None());
    };
    let d = PyDict::new(py);
    d.set_item("axis_type", &axis.axis_type)?;
    d.set_item("axis_position", axis.axis_position.as_deref())?;
    d.set_item("ax_id", axis.ax_id)?;
    d.set_item("cross_ax", axis.cross_ax)?;
    d.set_item("scaling_min", axis.scaling_min)?;
    d.set_item("scaling_max", axis.scaling_max)?;
    d.set_item("scaling_orientation", axis.scaling_orientation.as_deref())?;
    d.set_item("scaling_log_base", axis.scaling_log_base)?;
    d.set_item("num_format_code", axis.num_format_code.as_deref())?;
    d.set_item("num_format_source_linked", axis.num_format_source_linked)?;
    d.set_item("major_unit", axis.major_unit)?;
    d.set_item("minor_unit", axis.minor_unit)?;
    d.set_item("tick_lbl_pos", axis.tick_lbl_pos.as_deref())?;
    d.set_item("major_tick_mark", axis.major_tick_mark.as_deref())?;
    d.set_item("minor_tick_mark", axis.minor_tick_mark.as_deref())?;
    d.set_item("crosses", axis.crosses.as_deref())?;
    d.set_item("crosses_at", axis.crosses_at)?;
    d.set_item("cross_between", axis.cross_between.as_deref())?;
    d.set_item("display_unit", axis.display_unit.as_deref())?;
    Ok(d.into())
}

pub(crate) fn chart_series_to_py(py: Python<'_>, series: &ChartSeriesInfo) -> PyResult<PyObject> {
    let d = PyDict::new(py);
    d.set_item("idx", series.idx)?;
    d.set_item("order", series.order)?;
    d.set_item("title_ref", series.title_ref.as_deref())?;
    d.set_item("title_value", series.title_value.as_deref())?;
    d.set_item(
        "graphical_properties",
        chart_graphical_properties_to_py(py, series.graphical_properties.as_ref())?,
    )?;
    d.set_item(
        "data_labels",
        chart_data_labels_to_py(py, series.data_labels.as_ref())?,
    )?;
    d.set_item(
        "trendline",
        chart_trendline_to_py(py, series.trendline.as_ref())?,
    )?;
    d.set_item(
        "error_bars",
        chart_error_bars_to_py(py, series.error_bars.as_ref())?,
    )?;
    d.set_item("cat_ref", series.cat_ref.as_deref())?;
    d.set_item("val_ref", series.val_ref.as_deref())?;
    d.set_item("x_ref", series.x_ref.as_deref())?;
    d.set_item("y_ref", series.y_ref.as_deref())?;
    d.set_item("bubble_size_ref", series.bubble_size_ref.as_deref())?;
    Ok(d.into())
}

pub(crate) fn chart_graphical_properties_to_py(
    py: Python<'_>,
    properties: Option<&ChartGraphicalPropertiesInfo>,
) -> PyResult<PyObject> {
    let Some(properties) = properties else {
        return Ok(py.None());
    };
    let d = PyDict::new(py);
    d.set_item("no_fill", properties.no_fill)?;
    d.set_item("solid_fill", properties.solid_fill.as_deref())?;
    let line = PyDict::new(py);
    line.set_item("no_fill", properties.line_no_fill)?;
    line.set_item("solid_fill", properties.line_solid_fill.as_deref())?;
    line.set_item("prst_dash", properties.line_dash.as_deref())?;
    line.set_item("w_emu", properties.line_width)?;
    d.set_item("ln", line)?;
    Ok(d.into())
}

pub(crate) fn chart_data_labels_to_py(
    py: Python<'_>,
    labels: Option<&ChartDataLabelsInfo>,
) -> PyResult<PyObject> {
    let Some(labels) = labels else {
        return Ok(py.None());
    };
    let d = PyDict::new(py);
    d.set_item("position", labels.position.as_deref())?;
    d.set_item("show_legend_key", labels.show_legend_key)?;
    d.set_item("show_val", labels.show_val)?;
    d.set_item("show_cat_name", labels.show_cat_name)?;
    d.set_item("show_ser_name", labels.show_ser_name)?;
    d.set_item("show_percent", labels.show_percent)?;
    d.set_item("show_bubble_size", labels.show_bubble_size)?;
    d.set_item("show_leader_lines", labels.show_leader_lines)?;
    Ok(d.into())
}

pub(crate) fn chart_trendline_to_py(
    py: Python<'_>,
    trendline: Option<&ChartTrendlineInfo>,
) -> PyResult<PyObject> {
    let Some(trendline) = trendline else {
        return Ok(py.None());
    };
    let d = PyDict::new(py);
    d.set_item("trendline_type", trendline.trendline_type.as_deref())?;
    d.set_item("order", trendline.order)?;
    d.set_item("period", trendline.period)?;
    d.set_item("forward", trendline.forward)?;
    d.set_item("backward", trendline.backward)?;
    d.set_item("intercept", trendline.intercept)?;
    d.set_item("display_equation", trendline.display_equation)?;
    d.set_item("display_r_squared", trendline.display_r_squared)?;
    Ok(d.into())
}

pub(crate) fn chart_error_bars_to_py(
    py: Python<'_>,
    error_bars: Option<&ChartErrorBarsInfo>,
) -> PyResult<PyObject> {
    let Some(error_bars) = error_bars else {
        return Ok(py.None());
    };
    let d = PyDict::new(py);
    d.set_item("direction", error_bars.direction.as_deref())?;
    d.set_item("bar_type", error_bars.bar_type.as_deref())?;
    d.set_item("val_type", error_bars.val_type.as_deref())?;
    d.set_item("no_end_cap", error_bars.no_end_cap)?;
    d.set_item("val", error_bars.val)?;
    Ok(d.into())
}

pub(crate) fn image_anchor_to_py(py: Python<'_>, anchor: &ImageAnchorInfo) -> PyResult<PyObject> {
    let d = PyDict::new(py);
    match anchor {
        ImageAnchorInfo::OneCell { from, ext } => {
            d.set_item("type", "one_cell")?;
            populate_marker(&d, "from", from)?;
            match ext {
                Some(ext) => populate_extent(&d, ext)?,
                None => {
                    d.set_item("cx_emu", py.None())?;
                    d.set_item("cy_emu", py.None())?;
                }
            }
        }
        ImageAnchorInfo::TwoCell { from, to, edit_as } => {
            d.set_item("type", "two_cell")?;
            populate_marker(&d, "from", from)?;
            populate_marker(&d, "to", to)?;
            d.set_item("edit_as", edit_as)?;
        }
        ImageAnchorInfo::Absolute { pos, ext } => {
            d.set_item("type", "absolute")?;
            populate_position(&d, pos)?;
            populate_extent(&d, ext)?;
        }
    }
    Ok(d.into())
}

pub(crate) fn populate_marker(
    d: &Bound<'_, PyDict>,
    prefix: &str,
    marker: &AnchorMarkerInfo,
) -> PyResult<()> {
    d.set_item(format!("{prefix}_col"), marker.col)?;
    d.set_item(format!("{prefix}_row"), marker.row)?;
    d.set_item(format!("{prefix}_col_off"), marker.col_off)?;
    d.set_item(format!("{prefix}_row_off"), marker.row_off)?;
    Ok(())
}

pub(crate) fn populate_position(d: &Bound<'_, PyDict>, pos: &AnchorPositionInfo) -> PyResult<()> {
    d.set_item("x_emu", pos.x)?;
    d.set_item("y_emu", pos.y)?;
    Ok(())
}

pub(crate) fn populate_extent(d: &Bound<'_, PyDict>, ext: &AnchorExtentInfo) -> PyResult<()> {
    d.set_item("cx_emu", ext.cx)?;
    d.set_item("cy_emu", ext.cy)?;
    Ok(())
}
