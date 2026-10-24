//! Chart payload parsing for the native writer backend.

use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use pyo3::types::PyDict;
use wolfxl_writer::model::chart::{
    Axis, AxisCommon, AxisOrientation, AxisPos, BarDir, BarGrouping, CategoryAxis, Chart,
    ChartKind, DataLabels, DataPoint, DateAxis, DisplayBlanksAs, DisplayUnits, ErrorBarType,
    ErrorBarValType, ErrorBars, GraphicalProperties, Gridlines, Layout, LayoutTarget, Legend,
    LegendPosition, Marker, MarkerSymbol, PivotSource, RadarStyle, Reference as ChartReference,
    ScatterStyle, Series, SeriesAxis, SeriesTitle, TickMark, Title as ChartTitle, TitleRun,
    Trendline, TrendlineKind, UpDownBars, ValueAxis, View3D,
};
use wolfxl_writer::model::image::ImageAnchor;

use crate::native_writer_anchors::parse_image_anchor;

/// Sprint Μ-prime — module-level PyO3 helper used by Pod-γ's Python
/// modify-mode bridge to render a chart dict to OOXML bytes without
/// going through `NativeWorkbook.add_chart_native`.
///
/// `chart_dict` is the v1.6.1 §10 canonical shape; `anchor_a1` is a
/// fallback A1 reference if the dict's `anchor` key is missing or
/// `None`. The returned bytes are a complete `xl/charts/chartN.xml`
/// part, ready for the patcher's `file_adds`.
#[pyfunction]
pub fn serialize_chart_dict(chart_dict: &Bound<'_, PyDict>, anchor_a1: &str) -> PyResult<Vec<u8>> {
    let chart = parse_chart_dict(chart_dict, anchor_a1)?;
    Ok(wolfxl_writer::emit::charts::emit_chart_xml(&chart))
}

// ---------------------------------------------------------------------------
// Sprint Μ Pod-α (RFC-046) — chart dict → typed Chart parsing
// ---------------------------------------------------------------------------

pub(crate) fn parse_chart_dict(d: &Bound<'_, PyDict>, anchor_a1: &str) -> PyResult<Chart> {
    let kind_str: String = d
        .get_item("kind")?
        .ok_or_else(|| PyValueError::new_err("chart dict missing 'kind'"))?
        .extract()?;
    let kind = match kind_str.as_str() {
        "bar" => ChartKind::Bar,
        "line" => ChartKind::Line,
        "pie" => ChartKind::Pie,
        "doughnut" => ChartKind::Doughnut,
        "area" => ChartKind::Area,
        "scatter" => ChartKind::Scatter,
        "bubble" => ChartKind::Bubble,
        "radar" => ChartKind::Radar,
        // Sprint Μ-prime (RFC-046 §11): new families.
        "bar3d" => ChartKind::Bar3D,
        "line3d" => ChartKind::Line3D,
        "pie3d" => ChartKind::Pie3D,
        "area3d" => ChartKind::Area3D,
        "surface" => ChartKind::Surface,
        "surface3d" => ChartKind::Surface3D,
        "stock" => ChartKind::Stock,
        "of_pie" => ChartKind::OfPie,
        other => {
            return Err(PyValueError::new_err(format!(
                "unknown chart kind {other:?} (expected bar/line/pie/doughnut/\
                 area/scatter/bubble/radar/bar3d/line3d/pie3d/area3d/surface/\
                 surface3d/stock/of_pie)"
            )))
        }
    };

    // Anchor: accept (a) explicit dict, (b) A1 string (Pod-β shape),
    // or (c) None / missing — fall back to the call-site `anchor_a1`.
    let anchor = if let Some(v) = d.get_item("anchor")? {
        if v.is_none() {
            a1_to_one_cell_anchor(anchor_a1)?
        } else if let Ok(ad) = v.cast::<PyDict>() {
            parse_image_anchor(ad)?
        } else if let Ok(s) = v.extract::<String>() {
            a1_to_one_cell_anchor(&s)?
        } else {
            return Err(PyValueError::new_err(
                "chart anchor must be a dict, A1 string, or None",
            ));
        }
    } else {
        a1_to_one_cell_anchor(anchor_a1)?
    };

    let mut chart = Chart::new(kind, anchor);

    if let Some(v) = d.get_item("title")? {
        if !v.is_none() {
            let td = v
                .cast::<PyDict>()
                .map_err(|_| PyValueError::new_err("chart title must be a dict"))?;
            chart.title = Some(parse_chart_title(td)?);
        }
    }

    if let Some(v) = d.get_item("legend")? {
        if v.is_none() {
            chart.legend = None;
        } else {
            let ld = v
                .cast::<PyDict>()
                .map_err(|_| PyValueError::new_err("chart legend must be a dict"))?;
            chart.legend = Some(parse_legend(ld)?);
        }
    }

    if let Some(v) = d.get_item("layout")? {
        if !v.is_none() {
            let ld = v
                .cast::<PyDict>()
                .map_err(|_| PyValueError::new_err("chart layout must be a dict"))?;
            chart.layout = Some(parse_layout(ld)?);
        }
    }

    if let Some(v) = d.get_item("x_axis")? {
        if !v.is_none() {
            let ad = v
                .cast::<PyDict>()
                .map_err(|_| PyValueError::new_err("x_axis must be a dict"))?;
            chart.x_axis = Some(parse_axis(ad)?);
        }
    }
    if let Some(v) = d.get_item("y_axis")? {
        if !v.is_none() {
            let ad = v
                .cast::<PyDict>()
                .map_err(|_| PyValueError::new_err("y_axis must be a dict"))?;
            chart.y_axis = Some(parse_axis(ad)?);
        }
    }

    if let Some(v) = d.get_item("series")? {
        let list: Vec<Bound<'_, PyAny>> = v.extract()?;
        for sv in list.iter() {
            let sd = sv
                .cast::<PyDict>()
                .map_err(|_| PyValueError::new_err("each series must be a dict"))?;
            chart.series.push(parse_series(sd)?);
        }
    }

    // RFC-046 §10.6.2: chart-level `data_labels` are emitted once inside
    // the chart-kind block, after series, matching openpyxl's `chart.dataLabels`
    // serialization. Per-series data labels remain supported via each series.
    if let Some(v) = d.get_item("data_labels")? {
        if !v.is_none() {
            let dd = v
                .cast::<PyDict>()
                .map_err(|_| PyValueError::new_err("data_labels must be a dict"))?;
            chart.data_labels = Some(parse_data_labels(dd)?);
        }
    }

    if let Some(b) = py_opt_bool(d, "plot_visible_only")? {
        chart.plot_visible_only = Some(b);
    }
    if let Some(s) = py_opt_str(d, "display_blanks_as")? {
        chart.display_blanks_as = Some(match s.as_str() {
            "gap" => DisplayBlanksAs::Gap,
            "span" => DisplayBlanksAs::Span,
            "zero" => DisplayBlanksAs::Zero,
            other => {
                return Err(PyValueError::new_err(format!(
                    "unknown display_blanks_as {other:?}"
                )))
            }
        });
    }
    if let Some(b) = py_opt_bool(d, "vary_colors")? {
        chart.vary_colors = Some(b);
    }

    if let Some(s) = py_opt_str(d, "bar_dir")? {
        chart.bar_dir = Some(match s.as_str() {
            "col" => BarDir::Col,
            "bar" => BarDir::Bar,
            other => return Err(PyValueError::new_err(format!("unknown bar_dir {other:?}"))),
        });
    }
    if let Some(s) = py_opt_str(d, "grouping")? {
        chart.grouping = Some(match s.as_str() {
            "clustered" => BarGrouping::Clustered,
            "stacked" => BarGrouping::Stacked,
            "percentStacked" => BarGrouping::PercentStacked,
            "standard" => BarGrouping::Standard,
            other => return Err(PyValueError::new_err(format!("unknown grouping {other:?}"))),
        });
    }
    if let Some(n) = py_opt_u32(d, "gap_width")? {
        chart.gap_width = Some(n);
    }
    if let Some(n) = py_opt_i32(d, "overlap")? {
        chart.overlap = Some(n);
    }
    if let Some(n) = py_opt_u32(d, "hole_size")? {
        chart.hole_size = Some(n);
    }
    if let Some(n) = py_opt_u32(d, "first_slice_ang")? {
        chart.first_slice_ang = Some(n);
    }
    if let Some(s) = py_opt_str(d, "scatter_style")? {
        chart.scatter_style = Some(parse_scatter_style(&s)?);
    }
    if let Some(s) = py_opt_str(d, "radar_style")? {
        chart.radar_style = Some(parse_radar_style(&s)?);
    }
    if let Some(b) = py_opt_bool(d, "bubble3d")? {
        chart.bubble3d = Some(b);
    }
    if let Some(n) = py_opt_u32(d, "bubble_scale")? {
        chart.bubble_scale = Some(n);
    }
    if let Some(b) = py_opt_bool(d, "show_neg_bubbles")? {
        chart.show_neg_bubbles = Some(b);
    }
    if let Some(b) = py_opt_bool(d, "smoothing")? {
        chart.smoothing = Some(b);
    }
    // Only stock charts emit these; other kinds keep ignoring the keys.
    if matches!(chart.kind, ChartKind::Stock) {
        chart.hi_low_lines = parse_optional_chart_lines(d, "hi_low_lines")?;
        chart.up_down_bars = parse_optional_up_down_bars(d)?;
    }
    if let Some(n) = py_opt_u32(d, "style")? {
        chart.style = Some(n);
    }

    // Sprint Μ-prime (RFC-046 §10.10): view_3d on 3D chart kinds.
    if let Some(v) = d.get_item("view_3d")? {
        if !v.is_none() {
            let vd = v
                .cast::<PyDict>()
                .map_err(|_| PyValueError::new_err("view_3d must be a dict"))?;
            chart.view_3d = Some(parse_view_3d(vd)?);
        }
    }

    // Surface wireframe toggle (RFC-046 §11.3).
    if let Some(b) = py_opt_bool(d, "wireframe")? {
        chart.wireframe = Some(b);
    }

    // OfPie family fields.
    if let Some(s) = py_opt_str(d, "of_pie_type")? {
        chart.of_pie_type = Some(s);
    }
    if let Some(s) = py_opt_str(d, "split_type")? {
        chart.split_type = Some(s);
    }
    if let Some(f) = py_opt_f64(d, "split_pos")? {
        chart.split_pos = Some(f);
    }
    if let Some(n) = py_opt_u32(d, "second_pie_size")? {
        chart.second_pie_size = Some(n);
    }

    // Sprint Ν Pod-δ — RFC-049 §10. Optional `pivot_source` dict
    // {"name": str, "fmt_id": int} or None. Backward-compat: chart
    // dicts without this key parse identically to v1.7 output.
    if let Some(v) = d.get_item("pivot_source")? {
        if !v.is_none() {
            let psd = v
                .cast::<PyDict>()
                .map_err(|_| PyValueError::new_err("pivot_source must be a dict or None"))?;
            chart.pivot_source = Some(parse_pivot_source(psd)?);
        }
    }

    // RFC-069 / G15 — combination charts. Optional `secondary_charts`
    // list of fully-formed chart dicts. Each becomes a sibling chart
    // family inside the same `<plotArea>`. The recursive
    // `parse_chart_dict` call uses the same `anchor_a1` for nested
    // anchors; combination-chart emit deliberately ignores secondary
    // outer-frame fields (anchor, title, legend) so the value is moot.
    if let Some(v) = d.get_item("secondary_charts")? {
        if !v.is_none() {
            let list: Vec<Bound<'_, PyAny>> = v.extract().map_err(|_| {
                PyValueError::new_err("secondary_charts must be a list of chart dicts")
            })?;
            for sv in list.iter() {
                let sd = sv.cast::<PyDict>().map_err(|_| {
                    PyValueError::new_err("each secondary_charts entry must be a dict")
                })?;
                let secondary = parse_chart_dict(sd, anchor_a1)?;
                chart.secondary_charts.push(secondary);
            }
        }
    }

    Ok(chart)
}

/// RFC-049 §10.2 — parse + validate a chart `pivot_source` dict.
/// Validation matches the Python-side ``ChartBase._validate_pivot_source``
/// so write-mode (Python validation) and modify-mode (Rust validation)
/// reject the same inputs.
fn parse_pivot_source(d: &Bound<'_, PyDict>) -> PyResult<PivotSource> {
    let name: String = d
        .get_item("name")?
        .ok_or_else(|| PyValueError::new_err("pivot_source missing 'name'"))?
        .extract()
        .map_err(|_| PyValueError::new_err("pivot_source.name must be a string"))?;
    if name.is_empty() {
        return Err(PyValueError::new_err(
            "pivot_source.name must be a non-empty string",
        ));
    }
    if !is_valid_pivot_source_name(&name) {
        return Err(PyValueError::new_err(format!(
            "pivot_source.name={name:?} does not match the OOXML \
             pivot-source name regex"
        )));
    }
    let fmt_id: u32 = match d.get_item("fmt_id")? {
        Some(v) if !v.is_none() => v
            .extract()
            .map_err(|_| PyValueError::new_err("pivot_source.fmt_id must be an int"))?,
        _ => 0,
    };
    if fmt_id > 65535 {
        return Err(PyValueError::new_err(format!(
            "pivot_source.fmt_id={fmt_id} must be in [0, 65535]"
        )));
    }
    Ok(PivotSource { name, fmt_id })
}

/// RFC-049 §10.2 name regex implemented as a manual matcher (avoids a
/// `regex` dep). Pattern:
///     `^([A-Za-z_][A-Za-z0-9_]*!)?[A-Za-z_][A-Za-z0-9_ ]*$`
fn is_valid_pivot_source_name(s: &str) -> bool {
    fn is_ident_start(b: u8) -> bool {
        b.is_ascii_alphabetic() || b == b'_'
    }
    fn is_ident_cont(b: u8) -> bool {
        b.is_ascii_alphanumeric() || b == b'_'
    }
    fn is_table_cont(b: u8) -> bool {
        b.is_ascii_alphanumeric() || b == b'_' || b == b' '
    }
    let bytes = s.as_bytes();
    if bytes.is_empty() {
        return false;
    }
    // Optional `[ident]!` sheet-name prefix.
    let mut i = 0;
    if let Some(bang) = bytes.iter().position(|&b| b == b'!') {
        let prefix = &bytes[..bang];
        if prefix.is_empty() || !is_ident_start(prefix[0]) {
            return false;
        }
        if !prefix[1..].iter().copied().all(is_ident_cont) {
            return false;
        }
        i = bang + 1;
    }
    let table = &bytes[i..];
    if table.is_empty() || !is_ident_start(table[0]) {
        return false;
    }
    table[1..].iter().copied().all(is_table_cont)
}

/// Sprint Μ-prime — parse `<c:view3D>` dict per RFC-046 §10.10.
fn parse_view_3d(d: &Bound<'_, PyDict>) -> PyResult<View3D> {
    Ok(View3D {
        rot_x: py_opt_i16(d, "rot_x")?,
        rot_y: py_opt_i16(d, "rot_y")?,
        perspective: py_opt_u8(d, "perspective")?,
        right_angle_axes: py_opt_bool(d, "right_angle_axes")?,
        auto_scale: py_opt_bool(d, "auto_scale")?,
        depth_percent: py_opt_u32(d, "depth_percent")?,
        h_percent: py_opt_u32(d, "h_percent")?,
    })
}

/// Sprint Μ-prime — parse a `gridlines` dict per RFC-046 §10.7.1.
/// An empty dict is permitted ("draw default gridlines").
fn parse_gridlines(d: &Bound<'_, PyDict>) -> PyResult<Gridlines> {
    let graphical_properties = if let Some(v) = d.get_item("graphical_properties")? {
        if v.is_none() {
            None
        } else {
            let gd = v
                .cast::<PyDict>()
                .map_err(|_| PyValueError::new_err("graphical_properties must be a dict"))?;
            Some(parse_graphical_properties(gd)?)
        }
    } else {
        None
    };
    Ok(Gridlines {
        graphical_properties,
    })
}

fn parse_optional_chart_lines(d: &Bound<'_, PyDict>, key: &str) -> PyResult<Option<Gridlines>> {
    let Some(value) = d.get_item(key)? else {
        return Ok(None);
    };
    if value.is_none() {
        return Ok(None);
    }
    let lines = value
        .cast::<PyDict>()
        .map_err(|_| PyValueError::new_err(format!("{key} must be a dict")))?;
    Ok(Some(parse_gridlines(lines)?))
}

fn parse_optional_up_down_bars(d: &Bound<'_, PyDict>) -> PyResult<Option<UpDownBars>> {
    let Some(value) = d.get_item("up_down_bars")? else {
        return Ok(None);
    };
    if value.is_none() {
        return Ok(None);
    }
    let bars = value
        .cast::<PyDict>()
        .map_err(|_| PyValueError::new_err("up_down_bars must be a dict"))?;
    let gap_width = py_opt_u32(bars, "gap_width")?;
    if gap_width.is_some_and(|gap| gap > 500) {
        return Err(PyValueError::new_err(
            "up_down_bars.gap_width must be in [0, 500]",
        ));
    }
    Ok(Some(UpDownBars {
        gap_width,
        up_bars: parse_optional_chart_lines(bars, "up_bars")?,
        down_bars: parse_optional_chart_lines(bars, "down_bars")?,
    }))
}

fn a1_to_one_cell_anchor(a1: &str) -> PyResult<ImageAnchor> {
    let ((row, col), _) = wolfxl_writer::refs::parse_range(&format!("{a1}:{a1}"))
        .ok_or_else(|| PyValueError::new_err(format!("invalid A1 anchor {a1:?}")))?;
    Ok(ImageAnchor::OneCell {
        from_col: col.saturating_sub(1),
        from_row: row.saturating_sub(1),
        from_col_off: 0,
        from_row_off: 0,
    })
}

fn parse_chart_title(d: &Bound<'_, PyDict>) -> PyResult<ChartTitle> {
    // RFC-046 §10.3: `runs` and `text` are mutually exclusive; if both
    // present, `runs` wins. A `runs` value of `None` falls through to
    // the `text` plain-text path.
    let runs_obj = d.get_item("runs")?;
    let runs_is_useful = runs_obj.as_ref().map(|v| !v.is_none()).unwrap_or(false);

    let runs = if runs_is_useful {
        let v = runs_obj.unwrap();
        let list: Vec<Bound<'_, PyAny>> = v.extract()?;
        let mut out = Vec::with_capacity(list.len());
        for rv in list.iter() {
            let rd = rv
                .cast::<PyDict>()
                .map_err(|_| PyValueError::new_err("title run must be a dict"))?;
            let text: String = rd
                .get_item("text")?
                .ok_or_else(|| PyValueError::new_err("title run missing 'text'"))?
                .extract()?;

            // Two shapes accepted: flat ({"text", "bold", ...}) per the
            // pre-v1.6.1 internal contract, OR nested {"text",
            // "font": {"name", "size", "bold", "italic", "color"}} per
            // RFC-046 §10.3. The nested form is what Pod-β' emits.
            let mut bold = py_opt_bool(rd, "bold")?;
            let mut italic = py_opt_bool(rd, "italic")?;
            let mut underline = py_opt_bool(rd, "underline")?;
            let mut size_pt = py_opt_u32(rd, "size_pt")?;
            let mut color = py_opt_str(rd, "color")?;
            let mut font_name = py_opt_str(rd, "font_name")?;

            if let Some(fv) = rd.get_item("font")? {
                if !fv.is_none() {
                    let fd = fv
                        .cast::<PyDict>()
                        .map_err(|_| PyValueError::new_err("title run 'font' must be a dict"))?;
                    if let Some(b) = py_opt_bool(fd, "bold")? {
                        bold = Some(b);
                    }
                    if let Some(i) = py_opt_bool(fd, "italic")? {
                        italic = Some(i);
                    }
                    if let Some(u) = py_opt_bool(fd, "underline")? {
                        underline = Some(u);
                    }
                    // §10.3 uses "size", we also accept "size_pt".
                    if let Some(s) = py_opt_u32(fd, "size")? {
                        size_pt = Some(s);
                    } else if let Some(s) = py_opt_u32(fd, "size_pt")? {
                        size_pt = Some(s);
                    }
                    if let Some(c) = py_opt_str(fd, "color")? {
                        color = Some(c);
                    }
                    if let Some(n) = py_opt_str(fd, "name")? {
                        font_name = Some(n);
                    } else if let Some(n) = py_opt_str(fd, "font_name")? {
                        font_name = Some(n);
                    }
                }
            }

            out.push(TitleRun {
                text,
                bold,
                italic,
                underline,
                size_pt,
                color,
                font_name,
            });
        }
        out
    } else if let Some(v) = d.get_item("text")? {
        if v.is_none() {
            return Err(PyValueError::new_err(
                "chart title must have 'runs' or 'text'",
            ));
        }
        // Convenience: {"text": "Sales"} → single plain run.
        let text: String = v.extract()?;
        vec![TitleRun::plain(text)]
    } else {
        return Err(PyValueError::new_err(
            "chart title must have 'runs' or 'text'",
        ));
    };
    Ok(ChartTitle {
        runs,
        overlay: py_opt_bool(d, "overlay")?,
        layout: parse_optional_layout(d, "layout")?,
    })
}

fn parse_legend(d: &Bound<'_, PyDict>) -> PyResult<Legend> {
    let position = if let Some(s) = py_opt_str(d, "position")? {
        match s.as_str() {
            "r" => LegendPosition::Right,
            "l" => LegendPosition::Left,
            "t" => LegendPosition::Top,
            "b" => LegendPosition::Bottom,
            "tr" => LegendPosition::TopRight,
            other => {
                return Err(PyValueError::new_err(format!(
                    "unknown legend position {other:?}"
                )))
            }
        }
    } else {
        LegendPosition::Right
    };
    Ok(Legend {
        position,
        overlay: py_opt_bool(d, "overlay")?,
        layout: parse_optional_layout(d, "layout")?,
    })
}

fn parse_optional_layout(d: &Bound<'_, PyDict>, key: &str) -> PyResult<Option<Layout>> {
    if let Some(v) = d.get_item(key)? {
        if !v.is_none() {
            let ld = v
                .cast::<PyDict>()
                .map_err(|_| PyValueError::new_err(format!("{key} must be a dict")))?;
            return Ok(Some(parse_layout(ld)?));
        }
    }
    Ok(None)
}

fn parse_layout(d: &Bound<'_, PyDict>) -> PyResult<Layout> {
    // RFC-046 §10.5: x/y/w/h are all optional floats. Missing → 0.0
    // (the Layout struct uses bare f64; an all-zero layout is mostly a
    // no-op for Excel which interprets this as "place at origin with
    // zero size", but Pod-β' is responsible for emitting `None` at the
    // chart level instead of a zero-layout dict). The `*_mode` keys
    // (x_mode/y_mode/w_mode/h_mode) are currently honored by hardcoding
    // "edge" in the emitter; future work may plumb them through.
    let x: f64 = py_opt_f64(d, "x")?.unwrap_or(0.0);
    let y: f64 = py_opt_f64(d, "y")?.unwrap_or(0.0);
    let w: f64 = py_opt_f64(d, "w")?.unwrap_or(0.0);
    let h: f64 = py_opt_f64(d, "h")?.unwrap_or(0.0);
    let layout_target = if let Some(s) = py_opt_str(d, "layout_target")? {
        Some(match s.as_str() {
            "inner" => LayoutTarget::Inner,
            "outer" => LayoutTarget::Outer,
            other => {
                return Err(PyValueError::new_err(format!(
                    "unknown layout_target {other:?}"
                )))
            }
        })
    } else {
        None
    };
    Ok(Layout {
        x,
        y,
        w,
        h,
        layout_target,
    })
}

fn parse_axis(d: &Bound<'_, PyDict>) -> PyResult<Axis> {
    // Accept both the legacy {"kind": "category"} and the RFC-046 §10.7
    // {"ax_type": "cat"} shape.
    let kind: String = if let Some(v) = d.get_item("kind")? {
        v.extract()?
    } else if let Some(v) = d.get_item("ax_type")? {
        match v.extract::<String>()?.as_str() {
            "cat" => "category".to_string(),
            "val" => "value".to_string(),
            "date" => "date".to_string(),
            "ser" => "series".to_string(),
            other => {
                return Err(PyValueError::new_err(format!(
                    "unknown ax_type {other:?} (expected cat|val|date|ser)"
                )))
            }
        }
    } else {
        return Err(PyValueError::new_err(
            "axis dict missing 'kind' or 'ax_type'",
        ));
    };

    let common = parse_axis_common(d)?;

    // RFC-046 §10.7 nests scaling under "scaling": {"min", "max", ...}.
    let (scaled_min, scaled_max) = parse_axis_scaling(d)?;

    match kind.as_str() {
        "category" => Ok(Axis::Category(CategoryAxis {
            common,
            lbl_offset: py_opt_u32(d, "lbl_offset")?,
            lbl_algn: py_opt_str(d, "lbl_algn")?,
        })),
        "value" => Ok(Axis::Value(ValueAxis {
            common,
            min: scaled_min.or(py_opt_f64(d, "min")?),
            max: scaled_max.or(py_opt_f64(d, "max")?),
            major_unit: py_opt_f64(d, "major_unit")?,
            minor_unit: py_opt_f64(d, "minor_unit")?,
            display_units: parse_display_units(d)?,
            crosses: py_opt_str(d, "crosses")?,
        })),
        "date" => Ok(Axis::Date(DateAxis {
            common,
            min: scaled_min.or(py_opt_f64(d, "min")?),
            max: scaled_max.or(py_opt_f64(d, "max")?),
            major_unit: py_opt_f64(d, "major_unit")?,
            minor_unit: py_opt_f64(d, "minor_unit")?,
            base_time_unit: py_opt_str(d, "base_time_unit")?,
        })),
        "series" => Ok(Axis::Series(SeriesAxis { common })),
        other => Err(PyValueError::new_err(format!(
            "unknown axis kind {other:?} (expected category|value|date|series)"
        ))),
    }
}

/// RFC-046 §10.7: optional "scaling" sub-dict on axis. Returns (min, max).
/// Orientation/log_base are not yet consumed by the emitter.
fn parse_axis_scaling(d: &Bound<'_, PyDict>) -> PyResult<(Option<f64>, Option<f64>)> {
    if let Some(v) = d.get_item("scaling")? {
        if !v.is_none() {
            let sd = v
                .cast::<PyDict>()
                .map_err(|_| PyValueError::new_err("scaling must be a dict"))?;
            return Ok((py_opt_f64(sd, "min")?, py_opt_f64(sd, "max")?));
        }
    }
    Ok((None, None))
}

fn parse_axis_common(d: &Bound<'_, PyDict>) -> PyResult<AxisCommon> {
    let ax_id: u32 = d
        .get_item("ax_id")?
        .ok_or_else(|| PyValueError::new_err("axis missing 'ax_id'"))?
        .extract()?;
    let cross_ax: u32 = d
        .get_item("cross_ax")?
        .ok_or_else(|| PyValueError::new_err("axis missing 'cross_ax'"))?
        .extract()?;
    // RFC-046 §10.7 calls the field `axis_position`; the legacy shape
    // used `ax_pos`. Accept both with `axis_position` taking precedence.
    let ax_pos_raw = py_opt_str(d, "axis_position")?.or(py_opt_str(d, "ax_pos")?);
    let ax_pos = match ax_pos_raw.as_deref() {
        Some("b") | None => AxisPos::Bottom,
        Some("t") => AxisPos::Top,
        Some("l") => AxisPos::Left,
        Some("r") => AxisPos::Right,
        Some(other) => return Err(PyValueError::new_err(format!("unknown ax_pos {other:?}"))),
    };
    // Orientation may live on the axis OR (RFC-046) under `scaling`.
    let mut orientation_raw = py_opt_str(d, "orientation")?;
    if orientation_raw.is_none() {
        if let Some(v) = d.get_item("scaling")? {
            if !v.is_none() {
                if let Ok(sd) = v.cast::<PyDict>() {
                    orientation_raw = py_opt_str(sd, "orientation")?;
                }
            }
        }
    }
    let orientation = match orientation_raw.as_deref() {
        Some("minMax") | None => AxisOrientation::MinMax,
        Some("maxMin") => AxisOrientation::MaxMin,
        Some(other) => {
            return Err(PyValueError::new_err(format!(
                "unknown axis orientation {other:?}"
            )))
        }
    };
    let major_tick_mark = if let Some(s) = py_opt_str(d, "major_tick_mark")? {
        Some(parse_tick_mark(&s)?)
    } else {
        None
    };
    let minor_tick_mark = if let Some(s) = py_opt_str(d, "minor_tick_mark")? {
        Some(parse_tick_mark(&s)?)
    } else {
        None
    };
    let title = if let Some(v) = d.get_item("title")? {
        if v.is_none() {
            None
        } else {
            let td = v
                .cast::<PyDict>()
                .map_err(|_| PyValueError::new_err("axis title must be a dict"))?;
            Some(parse_chart_title(td)?)
        }
    } else {
        None
    };

    // RFC-046 §10.7.1: gridlines are dicts (with optional graphical
    // properties). Empty `{}` means "draw default gridlines"; `None`
    // means "no gridlines". Legacy bool form is also accepted —
    // {"major_gridlines": true} → flag-only emit.
    let (major_grid_flag, major_grid_obj) = parse_gridlines_field(d, "major_gridlines")?;
    let (minor_grid_flag, minor_grid_obj) = parse_gridlines_field(d, "minor_gridlines")?;

    let number_format = parse_axis_number_format(d)?;

    Ok(AxisCommon {
        ax_id,
        cross_ax,
        orientation,
        ax_pos,
        delete: py_opt_bool(d, "delete")?,
        major_tick_mark,
        minor_tick_mark,
        title,
        major_gridlines: major_grid_flag,
        minor_gridlines: minor_grid_flag,
        major_gridlines_obj: major_grid_obj,
        minor_gridlines_obj: minor_grid_obj,
        number_format,
    })
}

/// Parse the gridlines slot. Accepts:
///   - `None` / missing → (false, None) (no gridlines).
///   - `True` (bool)    → (true, None) (default gridlines, legacy).
///   - `False` (bool)   → (false, None).
///   - dict (possibly empty) → (false, Some(Gridlines{...})).
///
/// Returns `(flag, obj)` to feed AxisCommon.
fn parse_gridlines_field(d: &Bound<'_, PyDict>, key: &str) -> PyResult<(bool, Option<Gridlines>)> {
    let Some(v) = d.get_item(key)? else {
        return Ok((false, None));
    };
    if v.is_none() {
        return Ok((false, None));
    }
    if let Ok(b) = v.extract::<bool>() {
        return Ok((b, None));
    }
    let gd = v
        .cast::<PyDict>()
        .map_err(|_| PyValueError::new_err(format!("{key} must be a dict, bool, or None")))?;
    Ok((false, Some(parse_gridlines(gd)?)))
}

/// Number format on axis: accept either a string (legacy) or a
/// dict {"format_code", "source_linked"} per RFC-046 §10.7.
fn parse_axis_number_format(d: &Bound<'_, PyDict>) -> PyResult<Option<String>> {
    let Some(v) = d.get_item("number_format")? else {
        return Ok(None);
    };
    if v.is_none() {
        return Ok(None);
    }
    if let Ok(s) = v.extract::<String>() {
        return Ok(Some(s));
    }
    if let Ok(nfd) = v.cast::<PyDict>() {
        return Ok(py_opt_str(nfd, "format_code")?);
    }
    Ok(None)
}

fn parse_display_units(d: &Bound<'_, PyDict>) -> PyResult<Option<DisplayUnits>> {
    let Some(v) = d.get_item("disp_units")?.or(d.get_item("display_units")?) else {
        return Ok(None);
    };
    if v.is_none() {
        return Ok(None);
    }
    let dd = v
        .cast::<PyDict>()
        .map_err(|_| PyValueError::new_err("disp_units must be a dict or None"))?;
    Ok(Some(DisplayUnits {
        built_in_unit: py_opt_str(dd, "built_in_unit")?.or(py_opt_str(dd, "builtInUnit")?),
        custom_unit: py_opt_f64(dd, "custom_unit")?
            .or(py_opt_f64(dd, "cust_unit")?)
            .or(py_opt_f64(dd, "custUnit")?),
    }))
}

fn parse_tick_mark(s: &str) -> PyResult<TickMark> {
    Ok(match s {
        "none" => TickMark::None,
        "in" => TickMark::In,
        "out" => TickMark::Out,
        "cross" => TickMark::Cross,
        other => {
            return Err(PyValueError::new_err(format!(
                "unknown tick mark {other:?}"
            )))
        }
    })
}

fn parse_series(d: &Bound<'_, PyDict>) -> PyResult<Series> {
    let idx: u32 = d
        .get_item("idx")?
        .and_then(|v| v.extract().ok())
        .unwrap_or(0);
    let order: u32 = d
        .get_item("order")?
        .and_then(|v| v.extract().ok())
        .unwrap_or(idx);
    let mut s = Series::new(idx);
    s.order = order;

    // RFC-046 §10.6: title_ref (A1 string) | title_text (literal). Also
    // accept the legacy {"strRef": {...}} / {"literal": "..."} shape.
    if let Some(s_str) = py_opt_str(d, "title_ref")? {
        s.title = Some(SeriesTitle::StrRef(reference_from_a1(&s_str)?));
    } else if let Some(s_str) = py_opt_str(d, "title_text")? {
        s.title = Some(SeriesTitle::Literal(s_str));
    } else if let Some(v) = d.get_item("title")? {
        if !v.is_none() {
            // Legacy: {"strRef": {"sheet", "range"}} or {"literal": "..."}
            if let Ok(td) = v.cast::<PyDict>() {
                if let Some(rv) = td.get_item("strRef")? {
                    let rd = rv
                        .cast::<PyDict>()
                        .map_err(|_| PyValueError::new_err("strRef must be a dict"))?;
                    s.title = Some(SeriesTitle::StrRef(parse_reference(rd)?));
                } else if let Some(lv) = td.get_item("literal")? {
                    let s_str: String = lv.extract()?;
                    s.title = Some(SeriesTitle::Literal(s_str));
                }
            } else if let Ok(s_str) = v.extract::<String>() {
                // {"title": "Plain text"} convenience.
                s.title = Some(SeriesTitle::Literal(s_str));
            }
        }
    }

    s.categories = parse_series_ref_field(d, "categories", "categories_ref")?;
    s.values = parse_series_ref_field(d, "values", "values_ref")?;
    s.x_values = parse_series_ref_field(d, "x_values", "x_values_ref")?;
    if s.values.is_none() {
        s.values = parse_series_ref_field(d, "y_values", "y_values_ref")?;
    }
    s.bubble_size = parse_series_ref_field(d, "bubble_size", "bubble_size_ref")?;

    if let Some(v) = d.get_item("graphical_properties")? {
        if !v.is_none() {
            let gd = v
                .cast::<PyDict>()
                .map_err(|_| PyValueError::new_err("graphical_properties must be a dict"))?;
            s.graphical_properties = Some(parse_graphical_properties(gd)?);
        }
    }
    if let Some(v) = d.get_item("marker")? {
        if !v.is_none() {
            let md = v
                .cast::<PyDict>()
                .map_err(|_| PyValueError::new_err("marker must be a dict"))?;
            s.marker = Some(parse_marker(md)?);
        }
    }
    if let Some(v) = d.get_item("data_points")?.or(d.get_item("dPt")?) {
        if !v.is_none() {
            let list: Vec<Bound<'_, PyAny>> = v.extract()?;
            for dv in list.iter() {
                let dd = dv
                    .cast::<PyDict>()
                    .map_err(|_| PyValueError::new_err("data point must be a dict"))?;
                s.data_points.push(parse_data_point(dd)?);
            }
        }
    }
    if let Some(v) = d.get_item("data_labels")? {
        if !v.is_none() {
            let dd = v
                .cast::<PyDict>()
                .map_err(|_| PyValueError::new_err("data_labels must be a dict"))?;
            s.data_labels = Some(parse_data_labels(dd)?);
        }
    }
    // RFC-046 §10.6.3: `err_bars` (singular dict). Legacy: `error_bars`
    // (list of dicts). Accept both.
    if let Some(v) = d.get_item("err_bars")? {
        if !v.is_none() {
            let ed = v
                .cast::<PyDict>()
                .map_err(|_| PyValueError::new_err("err_bars must be a dict"))?;
            s.error_bars.push(parse_error_bars(ed)?);
        }
    }
    if let Some(v) = d.get_item("error_bars")? {
        if !v.is_none() {
            let list: Vec<Bound<'_, PyAny>> = v.extract()?;
            for ev in list.iter() {
                let ed = ev
                    .cast::<PyDict>()
                    .map_err(|_| PyValueError::new_err("error bar must be a dict"))?;
                s.error_bars.push(parse_error_bars(ed)?);
            }
        }
    }
    if let Some(v) = d.get_item("trendlines")? {
        if !v.is_none() {
            let list: Vec<Bound<'_, PyAny>> = v.extract()?;
            for tv in list.iter() {
                let td = tv
                    .cast::<PyDict>()
                    .map_err(|_| PyValueError::new_err("trendline must be a dict"))?;
                s.trendlines.push(parse_trendline(td)?);
            }
        }
    }

    s.smooth = py_opt_bool(d, "smooth")?;
    s.invert_if_negative = py_opt_bool(d, "invert_if_negative")?;
    Ok(s)
}

/// Pick up a series reference field. Tries the RFC-046 §10.6 A1-string
/// form first (`{prefix}_ref`), then the legacy dict form (`prefix`).
fn parse_series_ref_field(
    d: &Bound<'_, PyDict>,
    legacy_key: &str,
    ref_key: &str,
) -> PyResult<Option<ChartReference>> {
    if let Some(s_str) = py_opt_str(d, ref_key)? {
        return Ok(Some(reference_from_a1(&s_str)?));
    }
    if let Some(v) = d.get_item(legacy_key)? {
        if !v.is_none() {
            // Could be either the dict form or an A1 string.
            if let Ok(s_str) = v.extract::<String>() {
                return Ok(Some(reference_from_a1(&s_str)?));
            }
            let rd = v.cast::<PyDict>().map_err(|_| {
                PyValueError::new_err(format!("{legacy_key} must be a dict or A1 string"))
            })?;
            return Ok(Some(parse_reference(rd)?));
        }
    }
    Ok(None)
}

/// Parse an A1 string of the form `Sheet!A2:B6` or `'Sheet'!A2:B6` (with
/// optional `$` markers on cells) into a ChartReference. The sheet name
/// is the LHS of the first `!`. Cell range is preserved verbatim;
/// downstream `to_formula_string()` will absolutize as needed.
fn reference_from_a1(s: &str) -> PyResult<ChartReference> {
    let trimmed = s.trim();
    let (sheet, range) = trimmed.split_once('!').ok_or_else(|| {
        PyValueError::new_err(format!("expected A1 reference 'Sheet!A1:B2', got {s:?}"))
    })?;
    let sheet = sheet.trim_matches('\'').replace("''", "'");
    Ok(ChartReference::new(sheet, range))
}

fn parse_reference(d: &Bound<'_, PyDict>) -> PyResult<ChartReference> {
    let sheet: String = d
        .get_item("sheet")?
        .ok_or_else(|| PyValueError::new_err("reference missing 'sheet'"))?
        .extract()?;
    let range: String = d
        .get_item("range")?
        .ok_or_else(|| PyValueError::new_err("reference missing 'range'"))?
        .extract()?;
    Ok(ChartReference::new(sheet, range))
}

fn parse_graphical_properties(d: &Bound<'_, PyDict>) -> PyResult<GraphicalProperties> {
    // Pod-β (RFC-046 §10.9) emits the snake_case names `solid_fill` and
    // a nested `ln` dict with `solid_fill` / `prst_dash` / `w_emu`.
    // Earlier callers (legacy) used flat `fill_color` / `line_color` /
    // `line_dash` / `line_width_emu`. Accept either; §10 form wins.
    let fill_color = py_opt_str(d, "solid_fill")?.or(py_opt_str(d, "fill_color")?);

    // Nested ln dict per §10.9
    let mut line_color = py_opt_str(d, "line_color")?;
    let mut line_width_emu = py_opt_u32(d, "line_width_emu")?;
    let mut line_dash = py_opt_str(d, "line_dash")?;
    let mut no_line = py_opt_bool(d, "no_line")?.unwrap_or(false);

    if let Some(v) = d.get_item("ln")? {
        if !v.is_none() {
            let ln = v
                .cast::<PyDict>()
                .map_err(|_| PyValueError::new_err("ln must be a dict"))?;
            if line_color.is_none() {
                line_color = py_opt_str(ln, "solid_fill")?;
            }
            if line_width_emu.is_none() {
                line_width_emu = py_opt_u32(ln, "w_emu")?;
            }
            if line_dash.is_none() {
                line_dash = py_opt_str(ln, "prst_dash")?;
            }
            if !no_line {
                if let Some(b) = py_opt_bool(ln, "no_fill")? {
                    no_line = b;
                }
            }
        }
    }

    Ok(GraphicalProperties {
        line_color,
        line_width_emu,
        line_dash,
        fill_color,
        no_fill: py_opt_bool(d, "no_fill")?.unwrap_or(false),
        no_line,
    })
}

fn parse_marker(d: &Bound<'_, PyDict>) -> PyResult<Marker> {
    let symbol = match py_opt_str(d, "symbol")?.as_deref() {
        Some("none") => MarkerSymbol::None,
        Some("circle") | None => MarkerSymbol::Circle,
        Some("square") => MarkerSymbol::Square,
        Some("diamond") => MarkerSymbol::Diamond,
        Some("triangle") => MarkerSymbol::Triangle,
        Some("plus") => MarkerSymbol::Plus,
        Some("x") => MarkerSymbol::X,
        Some("star") => MarkerSymbol::Star,
        Some("dash") => MarkerSymbol::Dash,
        Some("dot") => MarkerSymbol::Dot,
        Some("auto") => MarkerSymbol::Auto,
        Some(other) => {
            return Err(PyValueError::new_err(format!(
                "unknown marker symbol {other:?}"
            )))
        }
    };
    let graphical_properties = if let Some(v) = d.get_item("graphical_properties")? {
        if v.is_none() {
            None
        } else {
            let gd = v
                .cast::<PyDict>()
                .map_err(|_| PyValueError::new_err("graphical_properties must be a dict"))?;
            Some(parse_graphical_properties(gd)?)
        }
    } else {
        None
    };
    Ok(Marker {
        symbol,
        size: py_opt_u32(d, "size")?,
        graphical_properties,
    })
}

fn parse_data_point(d: &Bound<'_, PyDict>) -> PyResult<DataPoint> {
    let idx = match d.get_item("idx")? {
        Some(v) if !v.is_none() => v
            .extract()
            .map_err(|_| PyValueError::new_err("data point idx must be an int"))?,
        _ => 0,
    };
    let marker = if let Some(v) = d.get_item("marker")? {
        if v.is_none() {
            None
        } else {
            let md = v
                .cast::<PyDict>()
                .map_err(|_| PyValueError::new_err("data point marker must be a dict"))?;
            Some(parse_marker(md)?)
        }
    } else {
        None
    };
    let graphical_properties =
        if let Some(v) = d.get_item("graphical_properties")?.or(d.get_item("spPr")?) {
            if v.is_none() {
                None
            } else {
                let gd = v.cast::<PyDict>().map_err(|_| {
                    PyValueError::new_err("data point graphical_properties must be a dict")
                })?;
                Some(parse_graphical_properties(gd)?)
            }
        } else {
            None
        };
    Ok(DataPoint {
        idx,
        invert_if_negative: py_opt_bool(d, "invert_if_negative")?
            .or(py_opt_bool(d, "invertIfNegative")?),
        marker,
        bubble_3d: py_opt_bool(d, "bubble_3d")?.or(py_opt_bool(d, "bubble3D")?),
        explosion: py_opt_u32(d, "explosion")?,
        graphical_properties,
    })
}

fn parse_data_labels(d: &Bound<'_, PyDict>) -> PyResult<DataLabels> {
    let tx_pr_runs = parse_optional_runs(d, "tx_pr_runs")?;
    Ok(DataLabels {
        show_val: py_opt_bool(d, "show_val")?,
        show_cat_name: py_opt_bool(d, "show_cat_name")?,
        show_ser_name: py_opt_bool(d, "show_ser_name")?,
        show_percent: py_opt_bool(d, "show_percent")?,
        show_legend_key: py_opt_bool(d, "show_legend_key")?,
        show_bubble_size: py_opt_bool(d, "show_bubble_size")?,
        position: py_opt_str(d, "position")?,
        number_format: py_opt_str(d, "number_format")?,
        separator: py_opt_str(d, "separator")?,
        tx_pr_runs,
    })
}

fn parse_optional_runs(d: &Bound<'_, PyDict>, key: &str) -> PyResult<Vec<TitleRun>> {
    let Some(v) = d.get_item(key)? else {
        return Ok(Vec::new());
    };
    if v.is_none() {
        return Ok(Vec::new());
    }
    let list: Vec<Bound<'_, PyAny>> = v.extract()?;
    let mut out = Vec::with_capacity(list.len());
    for rv in list.iter() {
        let rd = rv
            .cast::<PyDict>()
            .map_err(|_| PyValueError::new_err(format!("{key} entry must be a dict")))?;
        let text: String = rd
            .get_item("text")?
            .ok_or_else(|| PyValueError::new_err(format!("{key} run missing 'text'")))?
            .extract()?;
        let mut bold = py_opt_bool(rd, "bold")?;
        let mut italic = py_opt_bool(rd, "italic")?;
        let mut underline = py_opt_bool(rd, "underline")?;
        let mut size_pt = py_opt_u32(rd, "size_pt")?;
        let mut color = py_opt_str(rd, "color")?;
        let mut font_name = py_opt_str(rd, "font_name")?;
        if let Some(fv) = rd.get_item("font")? {
            if !fv.is_none() {
                let fd = fv.cast::<PyDict>().map_err(|_| {
                    PyValueError::new_err(format!("{key} run 'font' must be a dict"))
                })?;
                if let Some(b) = py_opt_bool(fd, "bold")? {
                    bold = Some(b);
                }
                if let Some(i) = py_opt_bool(fd, "italic")? {
                    italic = Some(i);
                }
                if let Some(u) = py_opt_bool(fd, "underline")? {
                    underline = Some(u);
                }
                if let Some(s) = py_opt_u32(fd, "size")? {
                    size_pt = Some(s);
                } else if let Some(s) = py_opt_u32(fd, "size_pt")? {
                    size_pt = Some(s);
                }
                if let Some(c) = py_opt_str(fd, "color")? {
                    color = Some(c);
                }
                if let Some(n) = py_opt_str(fd, "name")? {
                    font_name = Some(n);
                } else if let Some(n) = py_opt_str(fd, "font_name")? {
                    font_name = Some(n);
                }
            }
        }
        out.push(TitleRun {
            text,
            bold,
            italic,
            underline,
            size_pt,
            color,
            font_name,
        });
    }
    Ok(out)
}

fn parse_error_bars(d: &Bound<'_, PyDict>) -> PyResult<ErrorBars> {
    // RFC-046 §10.6.3 names: err_bar_type, err_val_type. Accept both
    // legacy bar_type/val_type and the new names.
    let bar_type_str = py_opt_str(d, "err_bar_type")?.or(py_opt_str(d, "bar_type")?);
    let bar_type = match bar_type_str.as_deref() {
        Some("plus") => ErrorBarType::Plus,
        Some("minus") => ErrorBarType::Minus,
        Some("both") | None => ErrorBarType::Both,
        Some(other) => {
            return Err(PyValueError::new_err(format!(
                "unknown error bar type {other:?}"
            )))
        }
    };
    let val_type_str = py_opt_str(d, "err_val_type")?.or(py_opt_str(d, "val_type")?);
    let val_type = match val_type_str.as_deref() {
        Some("cust") => ErrorBarValType::Cust,
        Some("fixedVal") | None => ErrorBarValType::FixedVal,
        Some("percentage") => ErrorBarValType::Percentage,
        Some("stdDev") => ErrorBarValType::StdDev,
        Some("stdErr") => ErrorBarValType::StdErr,
        Some(other) => {
            return Err(PyValueError::new_err(format!(
                "unknown error bar val_type {other:?}"
            )))
        }
    };
    // §10.6.3: `val` (or legacy `value`).
    let value = py_opt_f64(d, "val")?.or(py_opt_f64(d, "value")?);
    // direction / plus_ref / minus_ref are not yet plumbed to the Rust
    // model (the underlying ErrorBars struct doesn't carry them); they
    // are accepted-and-ignored so dicts conform to the contract without
    // erroring. Future work: extend ErrorBars with direction & cust refs.
    let _ = py_opt_str(d, "direction")?;
    let _ = py_opt_str(d, "plus_ref")?;
    let _ = py_opt_str(d, "minus_ref")?;

    Ok(ErrorBars {
        bar_type,
        val_type,
        value,
        no_end_cap: py_opt_bool(d, "no_end_cap")?,
    })
}

fn parse_trendline(d: &Bound<'_, PyDict>) -> PyResult<Trendline> {
    // RFC-046 §10.6.4 calls the field `trendline_type`; legacy used `kind`.
    let kind_raw = py_opt_str(d, "trendline_type")?.or(py_opt_str(d, "kind")?);
    let kind = match kind_raw.as_deref() {
        Some("linear") | None => TrendlineKind::Linear,
        Some("log") => TrendlineKind::Log,
        Some("power") => TrendlineKind::Power,
        Some("exp") => TrendlineKind::Exp,
        Some("poly") => TrendlineKind::Polynomial,
        Some("movingAvg") => TrendlineKind::MovingAvg,
        Some(other) => {
            return Err(PyValueError::new_err(format!(
                "unknown trendline kind {other:?}"
            )))
        }
    };
    // §10.6.4: disp_eq / disp_r_sqr; legacy: display_equation / display_r_squared.
    let display_equation = py_opt_bool(d, "disp_eq")?.or(py_opt_bool(d, "display_equation")?);
    let display_r_squared = py_opt_bool(d, "disp_r_sqr")?.or(py_opt_bool(d, "display_r_squared")?);
    // intercept is in the contract but not in the underlying Trendline struct;
    // accept-and-ignore for forward-compat.
    let _ = py_opt_f64(d, "intercept")?;
    Ok(Trendline {
        kind,
        order: py_opt_u32(d, "order")?,
        period: py_opt_u32(d, "period")?,
        forward: py_opt_f64(d, "forward")?,
        backward: py_opt_f64(d, "backward")?,
        display_equation,
        display_r_squared,
        name: py_opt_str(d, "name")?,
    })
}

fn parse_scatter_style(s: &str) -> PyResult<ScatterStyle> {
    Ok(match s {
        "line" => ScatterStyle::Line,
        "lineMarker" => ScatterStyle::LineMarker,
        "marker" => ScatterStyle::Marker,
        "smooth" => ScatterStyle::Smooth,
        "smoothMarker" => ScatterStyle::SmoothMarker,
        "none" => ScatterStyle::None,
        other => {
            return Err(PyValueError::new_err(format!(
                "unknown scatter_style {other:?}"
            )))
        }
    })
}

fn parse_radar_style(s: &str) -> PyResult<RadarStyle> {
    Ok(match s {
        "standard" => RadarStyle::Standard,
        "marker" => RadarStyle::Marker,
        "filled" => RadarStyle::Filled,
        other => {
            return Err(PyValueError::new_err(format!(
                "unknown radar_style {other:?}"
            )))
        }
    })
}

fn py_opt_bool(d: &Bound<'_, PyDict>, key: &str) -> PyResult<Option<bool>> {
    if let Some(v) = d.get_item(key)? {
        if v.is_none() {
            return Ok(None);
        }
        return Ok(Some(v.extract()?));
    }
    Ok(None)
}

fn py_opt_str(d: &Bound<'_, PyDict>, key: &str) -> PyResult<Option<String>> {
    if let Some(v) = d.get_item(key)? {
        if v.is_none() {
            return Ok(None);
        }
        return Ok(Some(v.extract()?));
    }
    Ok(None)
}

fn py_opt_u32(d: &Bound<'_, PyDict>, key: &str) -> PyResult<Option<u32>> {
    if let Some(v) = d.get_item(key)? {
        if v.is_none() {
            return Ok(None);
        }
        // Accept either Python int or float (Pod-β's _BoundedNumber
        // stores as float; the contract treats these as integers).
        if let Ok(n) = v.extract::<u32>() {
            return Ok(Some(n));
        }
        if let Ok(f) = v.extract::<f64>() {
            if f.is_finite() && f >= 0.0 && f <= u32::MAX as f64 {
                return Ok(Some(f as u32));
            }
        }
        return Err(PyValueError::new_err(format!(
            "{key}: expected non-negative integer (got {})",
            v.repr()?.to_string()
        )));
    }
    Ok(None)
}

fn py_opt_i32(d: &Bound<'_, PyDict>, key: &str) -> PyResult<Option<i32>> {
    if let Some(v) = d.get_item(key)? {
        if v.is_none() {
            return Ok(None);
        }
        if let Ok(n) = v.extract::<i32>() {
            return Ok(Some(n));
        }
        if let Ok(f) = v.extract::<f64>() {
            if f.is_finite() && f >= i32::MIN as f64 && f <= i32::MAX as f64 {
                return Ok(Some(f as i32));
            }
        }
        return Err(PyValueError::new_err(format!(
            "{key}: expected integer (got {})",
            v.repr()?.to_string()
        )));
    }
    Ok(None)
}

fn py_opt_f64(d: &Bound<'_, PyDict>, key: &str) -> PyResult<Option<f64>> {
    if let Some(v) = d.get_item(key)? {
        if v.is_none() {
            return Ok(None);
        }
        return Ok(Some(v.extract()?));
    }
    Ok(None)
}

fn py_opt_i16(d: &Bound<'_, PyDict>, key: &str) -> PyResult<Option<i16>> {
    if let Some(v) = d.get_item(key)? {
        if v.is_none() {
            return Ok(None);
        }
        return Ok(Some(v.extract()?));
    }
    Ok(None)
}

fn py_opt_u8(d: &Bound<'_, PyDict>, key: &str) -> PyResult<Option<u8>> {
    if let Some(v) = d.get_item(key)? {
        if v.is_none() {
            return Ok(None);
        }
        return Ok(Some(v.extract()?));
    }
    Ok(None)
}
