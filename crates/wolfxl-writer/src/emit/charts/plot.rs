//! Per-type chart plot element emission.

use crate::model::chart::{
    BarDir, BarGrouping, Chart, ChartKind, Gridlines, RadarStyle, UpDownBars,
};
use crate::xml_escape;

use super::primitives::{bool_str, fmt_f64};
use super::series::{emit_data_labels, emit_series};
use super::style::emit_graphical_props;

pub(super) fn emit_plot_chart(out: &mut String, chart: &Chart, ax_a: u32, ax_b: u32) {
    let elem = chart.kind.plot_element_name();
    out.push_str(&format!("<c:{elem}>"));

    // varyColors comes first in plot-area children (openpyxl order).
    if let Some(v) = chart.vary_colors {
        out.push_str(&format!("<c:varyColors val=\"{}\"/>", bool_str(v)));
    }

    // Type-specific shape header.
    match chart.kind {
        ChartKind::Bar | ChartKind::Bar3D => {
            if let Some(d) = chart.bar_dir {
                out.push_str(&format!("<c:barDir val=\"{}\"/>", d.as_str()));
            } else {
                out.push_str(&format!("<c:barDir val=\"{}\"/>", BarDir::Col.as_str()));
            }
            if let Some(g) = chart.grouping {
                out.push_str(&format!("<c:grouping val=\"{}\"/>", g.as_str()));
            } else {
                out.push_str(&format!(
                    "<c:grouping val=\"{}\"/>",
                    BarGrouping::Clustered.as_str()
                ));
            }
        }
        ChartKind::Line | ChartKind::Line3D => {
            // Default grouping for Line is "standard".
            let g = chart.grouping.unwrap_or(BarGrouping::Standard);
            out.push_str(&format!("<c:grouping val=\"{}\"/>", g.as_str()));
        }
        ChartKind::Area | ChartKind::Area3D => {
            let g = chart.grouping.unwrap_or(BarGrouping::Standard);
            out.push_str(&format!("<c:grouping val=\"{}\"/>", g.as_str()));
        }
        ChartKind::Scatter => {
            if let Some(s) = chart.scatter_style {
                out.push_str(&format!("<c:scatterStyle val=\"{}\"/>", s.as_str()));
            }
        }
        ChartKind::Radar => {
            let s = chart.radar_style.unwrap_or(RadarStyle::Standard);
            out.push_str(&format!("<c:radarStyle val=\"{}\"/>", s.as_str()));
        }
        ChartKind::OfPie => {
            // ofPieType comes first inside <ofPieChart>.
            let t = chart.of_pie_type.as_deref().unwrap_or("pie");
            out.push_str(&format!("<c:ofPieType val=\"{}\"/>", xml_escape::attr(t)));
        }
        ChartKind::Surface | ChartKind::Surface3D => {
            if let Some(w) = chart.wireframe {
                out.push_str(&format!("<c:wireframe val=\"{}\"/>", bool_str(w)));
            }
        }
        // Pie/Doughnut/Pie3D/Bubble/Stock have no opening style attr beyond varyColors.
        _ => {}
    }

    // Series. When the chart has a `pivot_source`,
    // each `<c:ser>` MUST carry a `<c:fmtId val="0"/>` element matching
    // the pivot source's fmt_id; Excel rejects pivot charts whose
    // series lack `<c:fmtId>`.
    let pivot_fmt_id = chart.pivot_source.as_ref().map(|ps| ps.fmt_id);
    for ser in &chart.series {
        emit_series(out, ser, chart.kind, pivot_fmt_id);
    }
    if let Some(d) = &chart.data_labels {
        emit_data_labels(out, d);
    }

    // Type-specific trailing properties.
    match chart.kind {
        ChartKind::Bar | ChartKind::Bar3D => {
            if let Some(g) = chart.gap_width {
                out.push_str(&format!("<c:gapWidth val=\"{g}\"/>"));
            }
            if let Some(o) = chart.overlap {
                out.push_str(&format!("<c:overlap val=\"{o}\"/>"));
            }
        }
        ChartKind::Doughnut => {
            if let Some(a) = chart.first_slice_ang {
                out.push_str(&format!("<c:firstSliceAng val=\"{a}\"/>"));
            }
            if let Some(h) = chart.hole_size {
                out.push_str(&format!("<c:holeSize val=\"{h}\"/>"));
            }
        }
        ChartKind::Pie | ChartKind::Pie3D => {
            if let Some(a) = chart.first_slice_ang {
                out.push_str(&format!("<c:firstSliceAng val=\"{a}\"/>"));
            }
        }
        ChartKind::OfPie => {
            // gapWidth, splitType, splitPos, secondPieSize.
            if let Some(g) = chart.gap_width {
                out.push_str(&format!("<c:gapWidth val=\"{g}\"/>"));
            }
            let st = chart.split_type.as_deref().unwrap_or("auto");
            out.push_str(&format!("<c:splitType val=\"{}\"/>", xml_escape::attr(st)));
            if let Some(p) = chart.split_pos {
                out.push_str(&format!("<c:splitPos val=\"{}\"/>", fmt_f64(p)));
            }
            if let Some(s) = chart.second_pie_size {
                out.push_str(&format!("<c:secondPieSize val=\"{s}\"/>"));
            }
        }
        ChartKind::Bubble => {
            if let Some(s) = chart.bubble_scale {
                out.push_str(&format!("<c:bubbleScale val=\"{s}\"/>"));
            }
            if let Some(b) = chart.show_neg_bubbles {
                out.push_str(&format!("<c:showNegBubbles val=\"{}\"/>", bool_str(b)));
            }
            if let Some(b) = chart.bubble3d {
                out.push_str(&format!("<c:bubble3D val=\"{}\"/>", bool_str(b)));
            }
        }
        ChartKind::Line | ChartKind::Line3D => {
            if let Some(s) = chart.smoothing {
                out.push_str(&format!("<c:smooth val=\"{}\"/>", bool_str(s)));
            }
        }
        ChartKind::Stock => {
            // Both decorations are optional in CT_StockChart; `None` means the
            // chart has none (StockChart(hi_low_lines=False, up_down_bars=False)).
            emit_chart_lines(out, "hiLowLines", chart.hi_low_lines.as_ref());
            if let Some(bars) = &chart.up_down_bars {
                emit_up_down_bars(out, bars);
            }
        }
        _ => {}
    }

    // Axis ids — Pie/Doughnut emit none; everything else emits both.
    if !chart.kind.is_axis_free() {
        out.push_str(&format!("<c:axId val=\"{ax_a}\"/>"));
        out.push_str(&format!("<c:axId val=\"{ax_b}\"/>"));
    }

    out.push_str(&format!("</c:{elem}>"));
}

fn emit_chart_lines(out: &mut String, name: &str, lines: Option<&Gridlines>) {
    if let Some(lines) = lines {
        out.push_str(&format!("<c:{name}>"));
        if let Some(properties) = &lines.graphical_properties {
            emit_graphical_props(out, properties);
        }
        out.push_str(&format!("</c:{name}>"));
    }
}

fn emit_up_down_bars(out: &mut String, bars: &UpDownBars) {
    out.push_str(&format!(
        "<c:upDownBars><c:gapWidth val=\"{}\"/>",
        bars.gap_width.unwrap_or(150)
    ));
    emit_chart_lines(out, "upBars", bars.up_bars.as_ref());
    emit_chart_lines(out, "downBars", bars.down_bars.as_ref());
    out.push_str("</c:upDownBars>");
}
