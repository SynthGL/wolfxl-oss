//! Integration tests for the chart emit pipeline.
//!
//! Each test builds a [`Chart`] in memory, emits its
//! `xl/charts/chartN.xml` via [`wolfxl_writer::emit::charts::emit_chart_xml`],
//! then re-parses the bytes with `quick-xml` to assert structural
//! invariants. The Python API and cross-language parity suites layer their
//! own tests on top; this file owns the Rust-only contract.

use quick_xml::events::Event;
use quick_xml::Reader;

use wolfxl_writer::emit::charts::emit_chart_xml;
use wolfxl_writer::model::chart::{
    Axis, AxisCommon, AxisOrientation, AxisPos, BarGrouping, CategoryAxis, Chart, ChartKind,
    DataLabels, DataPoint, DateAxis, DisplayBlanksAs, DisplayUnits, ErrorBarType, ErrorBarValType,
    ErrorBars, GraphicalProperties, Gridlines, Layout, LayoutTarget, Legend, LegendPosition,
    Marker, MarkerSymbol, RadarStyle, Reference, Series, SeriesAxis, SeriesTitle, TickMark, Title,
    TitleRun, Trendline, TrendlineKind, UpDownBars, ValueAxis, View3D,
};
use wolfxl_writer::model::image::ImageAnchor;

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

fn parse_ok(bytes: &[u8]) {
    let text = std::str::from_utf8(bytes).expect("emit produced invalid utf-8");
    let mut reader = Reader::from_str(text);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Eof) => break,
            Err(e) => panic!("emit produced invalid xml: {e}\n\n{text}"),
            _ => (),
        }
        buf.clear();
    }
}

fn assert_contains(bytes: &[u8], needle: &str) {
    let text = std::str::from_utf8(bytes).expect("utf-8");
    assert!(
        text.contains(needle),
        "expected output to contain {needle:?}, got:\n{text}"
    );
}

fn assert_not_contains(bytes: &[u8], needle: &str) {
    let text = std::str::from_utf8(bytes).expect("utf-8");
    assert!(
        !text.contains(needle),
        "expected output NOT to contain {needle:?}, got:\n{text}"
    );
}

fn anchor_at(col: u32, row: u32) -> ImageAnchor {
    ImageAnchor::one_cell(col, row)
}

fn cat_value_pair() -> (Axis, Axis) {
    let cat = Axis::Category(CategoryAxis {
        common: AxisCommon::new(10, 100, AxisPos::Bottom),
        lbl_offset: Some(100),
        lbl_algn: None,
    });
    let val = Axis::Value(ValueAxis {
        common: {
            let mut a = AxisCommon::new(100, 10, AxisPos::Left);
            a.major_gridlines = true;
            a
        },
        min: None,
        max: None,
        major_unit: None,
        minor_unit: None,
        display_units: None,
        crosses: None,
    });
    (cat, val)
}

fn dual_value_pair() -> (Axis, Axis) {
    let x = Axis::Value(ValueAxis {
        common: AxisCommon::new(10, 100, AxisPos::Bottom),
        min: None,
        max: None,
        major_unit: None,
        minor_unit: None,
        display_units: None,
        crosses: None,
    });
    let y = Axis::Value(ValueAxis {
        common: AxisCommon::new(100, 10, AxisPos::Left),
        min: None,
        max: None,
        major_unit: None,
        minor_unit: None,
        display_units: None,
        crosses: None,
    });
    (x, y)
}

fn series_with_cat_val() -> Series {
    let mut s = Series::new(0);
    s.title = Some(SeriesTitle::StrRef(Reference::new("Sheet", "B1")));
    s.categories = Some(Reference::new("Sheet", "A2:A6"));
    s.values = Some(Reference::new("Sheet", "B2:B6"));
    s
}

// ---------------------------------------------------------------------------
// 8 chart-kind round-trips
// ---------------------------------------------------------------------------

#[test]
fn bar_chart_round_trips() {
    let mut c = Chart::new(ChartKind::Bar, anchor_at(0, 0));
    c.title = Some(Title::plain("Sales"));
    c.add_series(series_with_cat_val());
    let (x, y) = cat_value_pair();
    c.x_axis = Some(x);
    c.y_axis = Some(y);
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:barChart>");
    assert_contains(&bytes, "<c:barDir val=\"col\"/>");
    assert_contains(&bytes, "<c:grouping val=\"clustered\"/>");
    assert_contains(&bytes, "<c:gapWidth val=\"150\"/>");
    assert_contains(&bytes, "<c:catAx>");
    assert_contains(&bytes, "<c:valAx>");
    assert_contains(&bytes, "<c:axId val=\"10\"/>");
    assert_contains(&bytes, "<c:axId val=\"100\"/>");
    assert_contains(&bytes, "'Sheet'!$A$2:$A$6");
    assert_contains(&bytes, "'Sheet'!$B$2:$B$6");
}

#[test]
fn line_chart_round_trips() {
    let mut c = Chart::new(ChartKind::Line, anchor_at(0, 0));
    c.title = Some(Title::plain("Trend"));
    c.add_series(series_with_cat_val());
    let (x, y) = cat_value_pair();
    c.x_axis = Some(x);
    c.y_axis = Some(y);
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:lineChart>");
    assert_contains(&bytes, "<c:grouping val=\"standard\"/>");
    // No barDir on line.
    assert_not_contains(&bytes, "<c:barDir");
}

#[test]
fn pie_chart_round_trips() {
    let mut c = Chart::new(ChartKind::Pie, anchor_at(0, 0));
    c.vary_colors = Some(true);
    let mut s = Series::new(0);
    s.values = Some(Reference::new("Sheet", "B2:B6"));
    s.categories = Some(Reference::new("Sheet", "A2:A6"));
    c.add_series(s);
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:pieChart>");
    assert_contains(&bytes, "<c:varyColors val=\"1\"/>");
    // Pie has no axes.
    assert_not_contains(&bytes, "<c:catAx>");
    assert_not_contains(&bytes, "<c:valAx>");
    assert_not_contains(&bytes, "<c:axId");
}

#[test]
fn doughnut_chart_round_trips() {
    let mut c = Chart::new(ChartKind::Doughnut, anchor_at(0, 0));
    c.hole_size = Some(75);
    c.vary_colors = Some(true);
    let mut s = Series::new(0);
    s.values = Some(Reference::new("Sheet", "B2:B6"));
    c.add_series(s);
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:doughnutChart>");
    assert_contains(&bytes, "<c:holeSize val=\"75\"/>");
    assert_not_contains(&bytes, "<c:catAx>");
}

#[test]
fn area_chart_round_trips() {
    let mut c = Chart::new(ChartKind::Area, anchor_at(0, 0));
    c.grouping = Some(BarGrouping::PercentStacked);
    c.add_series(series_with_cat_val());
    let (x, y) = cat_value_pair();
    c.x_axis = Some(x);
    c.y_axis = Some(y);
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:areaChart>");
    assert_contains(&bytes, "<c:grouping val=\"percentStacked\"/>");
    assert_contains(&bytes, "<c:catAx>");
    assert_contains(&bytes, "<c:valAx>");
}

#[test]
fn scatter_chart_round_trips() {
    let mut c = Chart::new(ChartKind::Scatter, anchor_at(0, 0));
    let mut s = Series::new(0);
    s.x_values = Some(Reference::new("Sheet", "A2:A6"));
    s.values = Some(Reference::new("Sheet", "B2:B6"));
    c.add_series(s);
    let (x, y) = dual_value_pair();
    c.x_axis = Some(x);
    c.y_axis = Some(y);
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:scatterChart>");
    assert_not_contains(&bytes, "<c:scatterStyle");
    assert_contains(&bytes, "<c:xVal>");
    assert_contains(&bytes, "<c:yVal>");
    assert_not_contains(&bytes, "<c:cat>");
    // Scatter uses two valAx, no catAx.
    assert_not_contains(&bytes, "<c:catAx>");
    let text = std::str::from_utf8(&bytes).unwrap();
    assert_eq!(
        text.matches("<c:valAx>").count(),
        2,
        "two valAx for scatter"
    );
}

#[test]
fn bubble_chart_round_trips() {
    let mut c = Chart::new(ChartKind::Bubble, anchor_at(0, 0));
    c.bubble_scale = Some(120);
    c.show_neg_bubbles = Some(false);
    let mut s = Series::new(0);
    s.x_values = Some(Reference::new("Sheet", "A2:A6"));
    s.values = Some(Reference::new("Sheet", "B2:B6"));
    s.bubble_size = Some(Reference::new("Sheet", "C2:C6"));
    c.add_series(s);
    let (x, y) = dual_value_pair();
    c.x_axis = Some(x);
    c.y_axis = Some(y);
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:bubbleChart>");
    assert_contains(&bytes, "<c:bubbleScale val=\"120\"/>");
    assert_contains(&bytes, "<c:bubbleSize>");
    assert_contains(&bytes, "<c:xVal>");
    assert_contains(&bytes, "<c:yVal>");
}

#[test]
fn radar_chart_round_trips() {
    let mut c = Chart::new(ChartKind::Radar, anchor_at(0, 0));
    c.radar_style = Some(RadarStyle::Filled);
    c.add_series(series_with_cat_val());
    let (x, y) = cat_value_pair();
    c.x_axis = Some(x);
    c.y_axis = Some(y);
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:radarChart>");
    assert_contains(&bytes, "<c:radarStyle val=\"filled\"/>");
    assert_contains(&bytes, "<c:catAx>");
}

// ---------------------------------------------------------------------------
// 12 sub-feature tests
// ---------------------------------------------------------------------------

#[test]
fn category_axis_emits_lbl_offset() {
    let mut c = Chart::new(ChartKind::Bar, anchor_at(0, 0));
    c.add_series(series_with_cat_val());
    c.x_axis = Some(Axis::Category(CategoryAxis {
        common: AxisCommon::new(10, 100, AxisPos::Bottom),
        lbl_offset: Some(75),
        lbl_algn: Some("ctr".to_string()),
    }));
    c.y_axis = Some(Axis::Value(ValueAxis {
        common: AxisCommon::new(100, 10, AxisPos::Left),
        min: None,
        max: None,
        major_unit: None,
        minor_unit: None,
        display_units: None,
        crosses: None,
    }));
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:lblOffset val=\"75\"/>");
    assert_contains(&bytes, "<c:lblAlgn val=\"ctr\"/>");
}

#[test]
fn value_axis_emits_min_max_units() {
    let mut c = Chart::new(ChartKind::Bar, anchor_at(0, 0));
    c.add_series(series_with_cat_val());
    c.x_axis = Some(Axis::Category(CategoryAxis {
        common: AxisCommon::new(10, 100, AxisPos::Bottom),
        lbl_offset: None,
        lbl_algn: None,
    }));
    c.y_axis = Some(Axis::Value(ValueAxis {
        common: AxisCommon::new(100, 10, AxisPos::Left),
        min: Some(0.0),
        max: Some(100.0),
        major_unit: Some(20.0),
        minor_unit: Some(5.0),
        display_units: None,
        crosses: Some("autoZero".to_string()),
    }));
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:min val=\"0\"/>");
    assert_contains(&bytes, "<c:max val=\"100\"/>");
    assert_contains(&bytes, "<c:majorUnit val=\"20\"/>");
    assert_contains(&bytes, "<c:minorUnit val=\"5\"/>");
    assert_contains(&bytes, "<c:crosses val=\"autoZero\"/>");
}

#[test]
fn date_axis_emits_base_time_unit() {
    let mut c = Chart::new(ChartKind::Line, anchor_at(0, 0));
    c.add_series(series_with_cat_val());
    c.x_axis = Some(Axis::Date(DateAxis {
        common: AxisCommon::new(10, 100, AxisPos::Bottom),
        min: None,
        max: None,
        major_unit: Some(1.0),
        minor_unit: None,
        base_time_unit: Some("months".to_string()),
    }));
    c.y_axis = Some(Axis::Value(ValueAxis {
        common: AxisCommon::new(100, 10, AxisPos::Left),
        min: None,
        max: None,
        major_unit: None,
        minor_unit: None,
        display_units: None,
        crosses: None,
    }));
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:dateAx>");
    assert_contains(&bytes, "<c:baseTimeUnit val=\"months\"/>");
    assert_contains(&bytes, "<c:majorUnit val=\"1\"/>");
}

#[test]
fn title_with_rich_text_runs() {
    let mut c = Chart::new(ChartKind::Bar, anchor_at(0, 0));
    c.add_series(series_with_cat_val());
    c.x_axis = Some(Axis::Category(CategoryAxis {
        common: AxisCommon::new(10, 100, AxisPos::Bottom),
        lbl_offset: None,
        lbl_algn: None,
    }));
    c.y_axis = Some(Axis::Value(ValueAxis {
        common: AxisCommon::new(100, 10, AxisPos::Left),
        min: None,
        max: None,
        major_unit: None,
        minor_unit: None,
        display_units: None,
        crosses: None,
    }));
    c.title = Some(Title {
        runs: vec![
            TitleRun {
                text: "Q4 ".into(),
                bold: Some(true),
                italic: None,
                underline: None,
                size_pt: Some(14),
                color: Some("FF0000FF".into()),
                font_name: Some("Calibri".into()),
            },
            TitleRun {
                text: "Sales".into(),
                bold: None,
                italic: Some(true),
                underline: None,
                size_pt: Some(14),
                color: None,
                font_name: None,
            },
        ],
        overlay: Some(false),
        layout: None,
    });
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<a:t>Q4 </a:t>");
    assert_contains(&bytes, "<a:t>Sales</a:t>");
    // Bold on first run.
    assert_contains(&bytes, "b=\"1\"");
    // Italic on second run.
    assert_contains(&bytes, "i=\"1\"");
    // Color on first run (alpha stripped).
    assert_contains(&bytes, "<a:srgbClr val=\"0000FF\"/>");
    assert_contains(&bytes, "Calibri");
    // Size encoded as 14*100 = 1400.
    assert_contains(&bytes, "sz=\"1400\"");
}

#[test]
fn legend_each_position() {
    let positions = [
        (LegendPosition::Right, "r"),
        (LegendPosition::Left, "l"),
        (LegendPosition::Top, "t"),
        (LegendPosition::Bottom, "b"),
        (LegendPosition::TopRight, "tr"),
    ];
    for (pos, expected) in positions {
        let mut c = Chart::new(ChartKind::Bar, anchor_at(0, 0));
        c.add_series(series_with_cat_val());
        c.x_axis = Some(Axis::Category(CategoryAxis {
            common: AxisCommon::new(10, 100, AxisPos::Bottom),
            lbl_offset: None,
            lbl_algn: None,
        }));
        c.y_axis = Some(Axis::Value(ValueAxis {
            common: AxisCommon::new(100, 10, AxisPos::Left),
            min: None,
            max: None,
            major_unit: None,
            minor_unit: None,
            display_units: None,
            crosses: None,
        }));
        c.legend = Some(Legend {
            position: pos,
            overlay: None,
            layout: None,
        });
        let bytes = emit_chart_xml(&c);
        parse_ok(&bytes);
        assert_contains(&bytes, &format!("<c:legendPos val=\"{expected}\"/>"));
    }
}

#[test]
fn data_labels_show_val_cat_position() {
    let mut c = Chart::new(ChartKind::Bar, anchor_at(0, 0));
    let mut s = series_with_cat_val();
    s.data_labels = Some(DataLabels {
        show_val: Some(true),
        show_cat_name: Some(true),
        show_ser_name: Some(false),
        show_percent: None,
        show_legend_key: None,
        show_bubble_size: None,
        position: Some("outEnd".to_string()),
        number_format: Some("0.00".to_string()),
        separator: Some(", ".to_string()),
        tx_pr_runs: Vec::new(),
    });
    c.add_series(s);
    let (x, y) = cat_value_pair();
    c.x_axis = Some(x);
    c.y_axis = Some(y);
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:showVal val=\"1\"/>");
    assert_contains(&bytes, "<c:showCatName val=\"1\"/>");
    assert_contains(&bytes, "<c:showSerName val=\"0\"/>");
    assert_contains(&bytes, "<c:dLblPos val=\"outEnd\"/>");
    assert_contains(&bytes, "formatCode=\"0.00\"");
    assert_contains(&bytes, "<c:separator>, </c:separator>");
}

#[test]
fn error_bars_each_type_and_val_type() {
    let cases = [
        (ErrorBarType::Plus, ErrorBarValType::FixedVal, 1.5_f64),
        (ErrorBarType::Minus, ErrorBarValType::Percentage, 5.0),
        (ErrorBarType::Both, ErrorBarValType::StdDev, 2.0),
    ];
    for (bt, vt, value) in cases {
        let mut c = Chart::new(ChartKind::Bar, anchor_at(0, 0));
        let mut s = series_with_cat_val();
        s.error_bars.push(ErrorBars {
            bar_type: bt,
            val_type: vt,
            value: Some(value),
            no_end_cap: Some(false),
        });
        c.add_series(s);
        let (x, y) = cat_value_pair();
        c.x_axis = Some(x);
        c.y_axis = Some(y);
        let bytes = emit_chart_xml(&c);
        parse_ok(&bytes);
        assert_contains(&bytes, &format!("<c:errBarType val=\"{}\"/>", bt.as_str()));
        assert_contains(&bytes, &format!("<c:errValType val=\"{}\"/>", vt.as_str()));
    }
}

#[test]
fn trendlines_linear_polynomial_movingavg() {
    let cases = [
        Trendline {
            kind: TrendlineKind::Linear,
            order: None,
            period: None,
            forward: Some(2.0),
            backward: Some(1.0),
            display_equation: Some(true),
            display_r_squared: Some(true),
            name: None,
        },
        Trendline {
            kind: TrendlineKind::Polynomial,
            order: Some(3),
            period: None,
            forward: None,
            backward: None,
            display_equation: None,
            display_r_squared: None,
            name: Some("Cubic".into()),
        },
        Trendline {
            kind: TrendlineKind::MovingAvg,
            order: None,
            period: Some(3),
            forward: None,
            backward: None,
            display_equation: None,
            display_r_squared: None,
            name: None,
        },
    ];
    for tl in cases {
        let expected_kind = tl.kind.as_str().to_string();
        let order = tl.order;
        let period = tl.period;
        let mut c = Chart::new(ChartKind::Bar, anchor_at(0, 0));
        let mut s = series_with_cat_val();
        s.trendlines.push(tl);
        c.add_series(s);
        let (x, y) = cat_value_pair();
        c.x_axis = Some(x);
        c.y_axis = Some(y);
        let bytes = emit_chart_xml(&c);
        parse_ok(&bytes);
        assert_contains(
            &bytes,
            &format!("<c:trendlineType val=\"{expected_kind}\"/>"),
        );
        if let Some(o) = order {
            assert_contains(&bytes, &format!("<c:order val=\"{o}\"/>"));
        }
        if let Some(p) = period {
            assert_contains(&bytes, &format!("<c:period val=\"{p}\"/>"));
        }
    }
}

#[test]
fn manual_layout_emits_emu_style_fractions() {
    let mut c = Chart::new(ChartKind::Bar, anchor_at(0, 0));
    c.add_series(series_with_cat_val());
    let (x, y) = cat_value_pair();
    c.x_axis = Some(x);
    c.y_axis = Some(y);
    c.layout = Some(Layout {
        x: 0.1,
        y: 0.2,
        w: 0.7,
        h: 0.6,
        layout_target: None,
    });
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:manualLayout>");
    assert_contains(&bytes, "<c:xMode val=\"edge\"/>");
    assert_contains(&bytes, "<c:x val=\"0.1\"/>");
    assert_contains(&bytes, "<c:y val=\"0.2\"/>");
    assert_contains(&bytes, "<c:w val=\"0.7\"/>");
    assert_contains(&bytes, "<c:h val=\"0.6\"/>");
}

#[test]
fn major_and_minor_gridlines_emitted() {
    let mut c = Chart::new(ChartKind::Bar, anchor_at(0, 0));
    c.add_series(series_with_cat_val());
    c.x_axis = Some(Axis::Category(CategoryAxis {
        common: AxisCommon::new(10, 100, AxisPos::Bottom),
        lbl_offset: None,
        lbl_algn: None,
    }));
    c.y_axis = Some(Axis::Value(ValueAxis {
        common: {
            let mut a = AxisCommon::new(100, 10, AxisPos::Left);
            a.major_gridlines = true;
            a.minor_gridlines = true;
            a
        },
        min: None,
        max: None,
        major_unit: None,
        minor_unit: None,
        display_units: None,
        crosses: None,
    }));
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:majorGridlines/>");
    assert_contains(&bytes, "<c:minorGridlines/>");
}

#[test]
fn graphical_properties_line_and_fill() {
    let mut c = Chart::new(ChartKind::Bar, anchor_at(0, 0));
    let mut s = series_with_cat_val();
    s.graphical_properties = Some(GraphicalProperties {
        line_color: Some("FFFF0000".into()),
        line_width_emu: Some(12700),
        line_dash: Some("dash".into()),
        fill_color: Some("FF00FF00".into()),
        no_fill: false,
        no_line: false,
    });
    c.add_series(s);
    let (x, y) = cat_value_pair();
    c.x_axis = Some(x);
    c.y_axis = Some(y);
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    // Fill color (alpha stripped).
    assert_contains(&bytes, "<a:srgbClr val=\"00FF00\"/>");
    // Line width.
    assert_contains(&bytes, "<a:ln w=\"12700\">");
    // Line color (alpha stripped).
    assert_contains(&bytes, "<a:srgbClr val=\"FF0000\"/>");
    // Dash style.
    assert_contains(&bytes, "<a:prstDash val=\"dash\"/>");
}

#[test]
fn marker_symbol_size_emitted() {
    let mut c = Chart::new(ChartKind::Line, anchor_at(0, 0));
    let mut s = Series::new(0);
    s.values = Some(Reference::new("S", "B2:B6"));
    s.marker = Some(Marker {
        symbol: MarkerSymbol::Triangle,
        size: Some(9),
        graphical_properties: None,
    });
    c.add_series(s);
    let (x, y) = cat_value_pair();
    c.x_axis = Some(x);
    c.y_axis = Some(y);
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:symbol val=\"triangle\"/>");
    assert_contains(&bytes, "<c:size val=\"9\"/>");
}

#[test]
fn data_point_overrides_emitted_on_series() {
    let mut c = Chart::new(ChartKind::Pie, anchor_at(0, 0));
    let mut s = Series::new(0);
    s.values = Some(Reference::new("S", "B2:B6"));
    s.data_points.push(DataPoint {
        idx: 2,
        invert_if_negative: None,
        marker: None,
        bubble_3d: None,
        explosion: Some(15),
        graphical_properties: Some(GraphicalProperties {
            fill_color: Some("FFFF0000".into()),
            ..GraphicalProperties::default()
        }),
    });
    c.add_series(s);
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:dPt>");
    assert_contains(&bytes, "<c:idx val=\"2\"/>");
    assert_contains(&bytes, "<c:explosion val=\"15\"/>");
    assert_contains(&bytes, "<a:srgbClr val=\"FF0000\"/>");
}

#[test]
fn multi_series_chart_emits_distinct_idx_order() {
    let mut c = Chart::new(ChartKind::Line, anchor_at(0, 0));
    for i in 0..3 {
        let mut s = Series::new(i);
        s.title = Some(SeriesTitle::Literal(format!("Series {i}")));
        s.values = Some(Reference::new("Sheet", &format!("B{}:B{}", i + 2, i + 6)));
        s.categories = Some(Reference::new("Sheet", "A2:A6"));
        c.add_series(s);
    }
    let (x, y) = cat_value_pair();
    c.x_axis = Some(x);
    c.y_axis = Some(y);
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    let text = std::str::from_utf8(&bytes).unwrap();
    // Three <c:ser> entries with distinct idx values.
    assert_eq!(text.matches("<c:ser>").count(), 3);
    assert_contains(&bytes, "<c:idx val=\"0\"/>");
    assert_contains(&bytes, "<c:idx val=\"1\"/>");
    assert_contains(&bytes, "<c:idx val=\"2\"/>");
    assert_contains(&bytes, "Series 0");
    assert_contains(&bytes, "Series 1");
    assert_contains(&bytes, "Series 2");
}

#[test]
fn vary_colors_true_for_pie() {
    let mut c = Chart::new(ChartKind::Pie, anchor_at(0, 0));
    c.vary_colors = Some(true);
    let mut s = Series::new(0);
    s.values = Some(Reference::new("Sheet", "B2:B6"));
    c.add_series(s);
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:varyColors val=\"1\"/>");
}

#[test]
fn series_axis_shape_emitted() {
    // Series axis is rare but present in 3D and surface charts.
    // Even for 2D kinds the shape has to round-trip if a user
    // attaches one.
    let mut c = Chart::new(ChartKind::Line, anchor_at(0, 0));
    c.add_series(series_with_cat_val());
    c.x_axis = Some(Axis::Series(SeriesAxis {
        common: AxisCommon::new(10, 100, AxisPos::Bottom),
    }));
    c.y_axis = Some(Axis::Value(ValueAxis {
        common: AxisCommon::new(100, 10, AxisPos::Left),
        min: None,
        max: None,
        major_unit: None,
        minor_unit: None,
        display_units: None,
        crosses: None,
    }));
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:serAx>");
}

#[test]
fn axis_orientation_and_tick_marks() {
    let mut c = Chart::new(ChartKind::Bar, anchor_at(0, 0));
    c.add_series(series_with_cat_val());
    c.x_axis = Some(Axis::Category(CategoryAxis {
        common: {
            let mut a = AxisCommon::new(10, 100, AxisPos::Bottom);
            a.orientation = AxisOrientation::MaxMin;
            a.major_tick_mark = Some(TickMark::Out);
            a.minor_tick_mark = Some(TickMark::None);
            a
        },
        lbl_offset: None,
        lbl_algn: None,
    }));
    c.y_axis = Some(Axis::Value(ValueAxis {
        common: AxisCommon::new(100, 10, AxisPos::Left),
        min: None,
        max: None,
        major_unit: None,
        minor_unit: None,
        display_units: None,
        crosses: None,
    }));
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:orientation val=\"maxMin\"/>");
    assert_contains(&bytes, "<c:majorTickMark val=\"out\"/>");
    assert_contains(&bytes, "<c:minorTickMark val=\"none\"/>");
}

#[test]
fn display_blanks_as_emitted() {
    let mut c = Chart::new(ChartKind::Line, anchor_at(0, 0));
    c.add_series(series_with_cat_val());
    let (x, y) = cat_value_pair();
    c.x_axis = Some(x);
    c.y_axis = Some(y);
    c.display_blanks_as = Some(DisplayBlanksAs::Span);
    c.plot_visible_only = Some(true);
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:dispBlanksAs val=\"span\"/>");
    assert_contains(&bytes, "<c:plotVisOnly val=\"1\"/>");
}

#[test]
fn value_axis_display_units_emitted() {
    let mut c = Chart::new(ChartKind::Line, anchor_at(0, 0));
    c.add_series(series_with_cat_val());
    let (x, mut y) = cat_value_pair();
    if let Axis::Value(ref mut val) = y {
        val.display_units = Some(DisplayUnits {
            built_in_unit: Some("millions".into()),
            custom_unit: Some(1000.0),
        });
    }
    c.x_axis = Some(x);
    c.y_axis = Some(y);
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:dispUnits>");
    assert_contains(&bytes, "<c:custUnit val=\"1000\"/>");
    assert_contains(&bytes, "<c:builtInUnit val=\"millions\"/>");
    let text = std::str::from_utf8(&bytes).unwrap();
    assert!(text.find("<c:custUnit").unwrap() < text.find("<c:builtInUnit").unwrap());
}

#[test]
fn bar_overlap_emitted_when_set() {
    let mut c = Chart::new(ChartKind::Bar, anchor_at(0, 0));
    c.overlap = Some(-25);
    c.add_series(series_with_cat_val());
    let (x, y) = cat_value_pair();
    c.x_axis = Some(x);
    c.y_axis = Some(y);
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:overlap val=\"-25\"/>");
}

#[test]
fn end_to_end_emits_chart_part_via_emit_xlsx() {
    use std::io::Cursor;
    use wolfxl_writer::emit_xlsx;
    use wolfxl_writer::model::Worksheet;
    use wolfxl_writer::Workbook;

    let mut wb = Workbook::new();
    let mut sheet = Worksheet::new("Sheet");
    let mut chart = Chart::new(ChartKind::Bar, anchor_at(3, 1));
    chart.title = Some(Title::plain("End-to-end"));
    chart.add_series(series_with_cat_val());
    let (x, y) = cat_value_pair();
    chart.x_axis = Some(x);
    chart.y_axis = Some(y);
    sheet.charts.push(chart);
    wb.add_sheet(sheet);

    let bytes = emit_xlsx(&mut wb);

    // Crack open the ZIP and verify the chart part exists.
    let archive = ::zip::ZipArchive::new(Cursor::new(bytes)).expect("open archive");
    let names: Vec<String> = archive
        .file_names()
        .map(|s| s.expect("valid fixture entry name").into_owned())
        .collect();
    assert!(
        names.iter().any(|n| n == "xl/charts/chart1.xml"),
        "archive missing chart part; entries:\n{names:#?}"
    );
    assert!(names.iter().any(|n| n == "xl/drawings/drawing1.xml"));
    assert!(names
        .iter()
        .any(|n| n == "xl/drawings/_rels/drawing1.xml.rels"));
}

// ---------------------------------------------------------------------------
// Additional chart-family round-trips and sub-features.
// ---------------------------------------------------------------------------

#[test]
fn bar3d_chart_round_trips() {
    let mut c = Chart::new(ChartKind::Bar3D, anchor_at(0, 0));
    c.title = Some(Title::plain("3D Sales"));
    c.add_series(series_with_cat_val());
    let (x, y) = cat_value_pair();
    c.x_axis = Some(x);
    c.y_axis = Some(y);
    // 3D defaults set by Chart::new.
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:bar3DChart>");
    assert_contains(&bytes, "<c:view3D>");
    assert_contains(&bytes, "<c:barDir val=\"col\"/>");
}

#[test]
fn line3d_chart_round_trips() {
    let mut c = Chart::new(ChartKind::Line3D, anchor_at(0, 0));
    c.add_series(series_with_cat_val());
    let (x, y) = cat_value_pair();
    c.x_axis = Some(x);
    c.y_axis = Some(y);
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:line3DChart>");
    assert_contains(&bytes, "<c:grouping val=\"standard\"/>");
    assert_contains(&bytes, "<c:view3D>");
}

#[test]
fn pie3d_chart_round_trips() {
    let mut c = Chart::new(ChartKind::Pie3D, anchor_at(0, 0));
    c.vary_colors = Some(true);
    let mut s = Series::new(0);
    s.values = Some(Reference::new("Sheet", "B2:B6"));
    s.categories = Some(Reference::new("Sheet", "A2:A6"));
    c.add_series(s);
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:pie3DChart>");
    assert_contains(&bytes, "<c:varyColors val=\"1\"/>");
    assert_contains(&bytes, "<c:view3D>");
    // Pie3D is axis-free.
    assert_not_contains(&bytes, "<c:axId");
}

#[test]
fn area3d_chart_round_trips() {
    let mut c = Chart::new(ChartKind::Area3D, anchor_at(0, 0));
    c.add_series(series_with_cat_val());
    let (x, y) = cat_value_pair();
    c.x_axis = Some(x);
    c.y_axis = Some(y);
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:area3DChart>");
    assert_contains(&bytes, "<c:grouping val=\"standard\"/>");
    assert_contains(&bytes, "<c:view3D>");
}

#[test]
fn surface_chart_round_trips_with_wireframe() {
    let mut c = Chart::new(ChartKind::Surface, anchor_at(0, 0));
    c.wireframe = Some(true);
    c.add_series(series_with_cat_val());
    let (x, y) = cat_value_pair();
    c.x_axis = Some(x);
    c.y_axis = Some(y);
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:surfaceChart>");
    assert_contains(&bytes, "<c:wireframe val=\"1\"/>");
    // Surface is NOT 3D, so no view3D.
    assert_not_contains(&bytes, "<c:view3D>");
}

#[test]
fn surface3d_chart_round_trips() {
    let mut c = Chart::new(ChartKind::Surface3D, anchor_at(0, 0));
    c.wireframe = Some(false);
    c.add_series(series_with_cat_val());
    let (x, y) = cat_value_pair();
    c.x_axis = Some(x);
    c.y_axis = Some(y);
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:surface3DChart>");
    assert_contains(&bytes, "<c:wireframe val=\"0\"/>");
    assert_contains(&bytes, "<c:view3D>");
}

#[test]
fn stock_chart_emits_only_the_decorations_it_models() {
    let mut c = Chart::new(ChartKind::Stock, anchor_at(0, 0));
    // 4 series typical (Open, High, Low, Close), but emit accepts any
    // count; semantic validation lives in the Python API tests.
    for i in 0..4u32 {
        let mut s = Series::new(i);
        s.values = Some(Reference::new("Sheet", &format!("B{}:B{}", 2 + i, 6 + i)));
        c.add_series(s);
    }
    let (x, y) = cat_value_pair();
    c.x_axis = Some(x);
    c.y_axis = Some(y);
    let bare = emit_chart_xml(&c);
    parse_ok(&bare);
    assert_contains(&bare, "<c:stockChart>");
    assert_not_contains(&bare, "hiLowLines");
    assert_not_contains(&bare, "upDownBars");

    c.hi_low_lines = Some(Gridlines::default());
    c.up_down_bars = Some(UpDownBars {
        gap_width: Some(150),
        ..UpDownBars::default()
    });
    let decorated = emit_chart_xml(&c);
    parse_ok(&decorated);
    assert_contains(&decorated, "<c:hiLowLines>");
    assert_contains(&decorated, "<c:upDownBars><c:gapWidth val=\"150\"/>");
}

#[test]
fn of_pie_chart_emits_split_type_and_of_pie_type() {
    let mut c = Chart::new(ChartKind::OfPie, anchor_at(0, 0));
    c.of_pie_type = Some("bar".to_string());
    c.split_type = Some("percent".to_string());
    c.split_pos = Some(15.0);
    c.second_pie_size = Some(75);
    c.gap_width = Some(100);
    let mut s = Series::new(0);
    s.values = Some(Reference::new("Sheet", "B2:B8"));
    s.categories = Some(Reference::new("Sheet", "A2:A8"));
    c.add_series(s);
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:ofPieChart>");
    assert_contains(&bytes, "<c:ofPieType val=\"bar\"/>");
    assert_contains(&bytes, "<c:splitType val=\"percent\"/>");
    assert_contains(&bytes, "<c:splitPos val=\"15\"/>");
    assert_contains(&bytes, "<c:secondPieSize val=\"75\"/>");
    assert_contains(&bytes, "<c:gapWidth val=\"100\"/>");
    // OfPie is axis-free.
    assert_not_contains(&bytes, "<c:axId");
}

#[test]
fn view_3d_emits_all_fields() {
    let mut c = Chart::new(ChartKind::Bar3D, anchor_at(0, 0));
    c.view_3d = Some(View3D {
        rot_x: Some(15),
        rot_y: Some(20),
        perspective: Some(30),
        right_angle_axes: Some(true),
        auto_scale: Some(true),
        depth_percent: Some(100),
        h_percent: Some(120),
    });
    c.add_series(series_with_cat_val());
    let (x, y) = cat_value_pair();
    c.x_axis = Some(x);
    c.y_axis = Some(y);
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:view3D>");
    assert_contains(&bytes, "<c:rotX val=\"15\"/>");
    assert_contains(&bytes, "<c:rotY val=\"20\"/>");
    assert_contains(&bytes, "<c:perspective val=\"30\"/>");
    assert_contains(&bytes, "<c:rAngAx val=\"1\"/>");
    assert_contains(&bytes, "<c:autoScale val=\"1\"/>");
    assert_contains(&bytes, "<c:depthPercent val=\"100\"/>");
    assert_contains(&bytes, "<c:hPercent val=\"120\"/>");
}

#[test]
fn major_gridlines_obj_with_graphical_properties() {
    let mut c = Chart::new(ChartKind::Bar, anchor_at(0, 0));
    c.add_series(series_with_cat_val());
    c.x_axis = Some(Axis::Category(CategoryAxis {
        common: AxisCommon::new(10, 100, AxisPos::Bottom),
        lbl_offset: None,
        lbl_algn: None,
    }));
    let mut yc = AxisCommon::new(100, 10, AxisPos::Left);
    yc.major_gridlines_obj = Some(Gridlines {
        graphical_properties: Some(GraphicalProperties {
            line_color: Some("FFCCCCCC".to_string()),
            ..Default::default()
        }),
    });
    c.y_axis = Some(Axis::Value(ValueAxis {
        common: yc,
        min: None,
        max: None,
        major_unit: None,
        minor_unit: None,
        display_units: None,
        crosses: None,
    }));
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    // Rich gridlines emit the open/close pair with spPr inside.
    assert_contains(&bytes, "<c:majorGridlines>");
    assert_contains(&bytes, "<c:spPr>");
    assert_contains(&bytes, "<a:srgbClr val=\"CCCCCC\"/>");
    assert_contains(&bytes, "</c:majorGridlines>");
}

#[test]
fn empty_gridlines_obj_emits_self_closing_tag() {
    let mut c = Chart::new(ChartKind::Bar, anchor_at(0, 0));
    c.add_series(series_with_cat_val());
    c.x_axis = Some(Axis::Category(CategoryAxis {
        common: AxisCommon::new(10, 100, AxisPos::Bottom),
        lbl_offset: None,
        lbl_algn: None,
    }));
    let mut yc = AxisCommon::new(100, 10, AxisPos::Left);
    yc.major_gridlines_obj = Some(Gridlines::default());
    c.y_axis = Some(Axis::Value(ValueAxis {
        common: yc,
        min: None,
        max: None,
        major_unit: None,
        minor_unit: None,
        display_units: None,
        crosses: None,
    }));
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    // Empty gridlines emit `<c:majorGridlines/>` with no inner spPr.
    assert_contains(&bytes, "<c:majorGridlines/>");
}

#[test]
fn fixedval_error_bars_emit_val_attribute() {
    let mut c = Chart::new(ChartKind::Bar, anchor_at(0, 0));
    let mut s = series_with_cat_val();
    s.error_bars.push(ErrorBars {
        bar_type: ErrorBarType::Both,
        val_type: ErrorBarValType::FixedVal,
        value: Some(0.5),
        no_end_cap: None,
    });
    c.add_series(s);
    let (x, y) = cat_value_pair();
    c.x_axis = Some(x);
    c.y_axis = Some(y);
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:errBars>");
    assert_contains(&bytes, "<c:errBarType val=\"both\"/>");
    assert_contains(&bytes, "<c:errValType val=\"fixedVal\"/>");
    assert_contains(&bytes, "<c:val val=\"0.5\"/>");
}

#[test]
fn linear_trendline_emits_linear_type() {
    let mut c = Chart::new(ChartKind::Line, anchor_at(0, 0));
    let mut s = Series::new(0);
    s.values = Some(Reference::new("Sheet", "B2:B6"));
    s.trendlines.push(Trendline {
        kind: TrendlineKind::Linear,
        order: None,
        period: None,
        forward: None,
        backward: None,
        display_equation: Some(true),
        display_r_squared: Some(true),
        name: None,
    });
    c.add_series(s);
    let (x, y) = cat_value_pair();
    c.x_axis = Some(x);
    c.y_axis = Some(y);
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:trendline>");
    assert_contains(&bytes, "<c:trendlineType val=\"linear\"/>");
    assert_contains(&bytes, "<c:dispEq val=\"1\"/>");
    assert_contains(&bytes, "<c:dispRSqr val=\"1\"/>");
}

#[test]
fn polynomial_trendline_order_3_emits_order() {
    let mut c = Chart::new(ChartKind::Scatter, anchor_at(0, 0));
    let mut s = Series::new(0);
    s.x_values = Some(Reference::new("Sheet", "A2:A8"));
    s.values = Some(Reference::new("Sheet", "B2:B8"));
    s.trendlines.push(Trendline {
        kind: TrendlineKind::Polynomial,
        order: Some(3),
        period: None,
        forward: None,
        backward: None,
        display_equation: None,
        display_r_squared: None,
        name: Some("Cubic Fit".to_string()),
    });
    c.add_series(s);
    let (x, y) = dual_value_pair();
    c.x_axis = Some(x);
    c.y_axis = Some(y);
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:trendlineType val=\"poly\"/>");
    assert_contains(&bytes, "<c:order val=\"3\"/>");
    assert_contains(&bytes, "<c:name>Cubic Fit</c:name>");
}

#[test]
fn data_labels_position_emitted() {
    let mut c = Chart::new(ChartKind::Bar, anchor_at(0, 0));
    let mut s = series_with_cat_val();
    s.data_labels = Some(DataLabels {
        show_val: Some(true),
        position: Some("outEnd".to_string()),
        ..Default::default()
    });
    c.add_series(s);
    let (x, y) = cat_value_pair();
    c.x_axis = Some(x);
    c.y_axis = Some(y);
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:dLblPos val=\"outEnd\"/>");
    assert_contains(&bytes, "<c:showVal val=\"1\"/>");
}

#[test]
fn marker_symbol_circle_size_emitted() {
    let mut c = Chart::new(ChartKind::Line, anchor_at(0, 0));
    let mut s = Series::new(0);
    s.values = Some(Reference::new("Sheet", "B2:B6"));
    s.marker = Some(Marker {
        symbol: MarkerSymbol::Circle,
        size: Some(9),
        graphical_properties: None,
    });
    c.add_series(s);
    let (x, y) = cat_value_pair();
    c.x_axis = Some(x);
    c.y_axis = Some(y);
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:marker>");
    assert_contains(&bytes, "<c:symbol val=\"circle\"/>");
    assert_contains(&bytes, "<c:size val=\"9\"/>");
}

#[test]
fn manual_layout_with_target_inner_emitted() {
    let mut c = Chart::new(ChartKind::Bar, anchor_at(0, 0));
    c.add_series(series_with_cat_val());
    c.layout = Some(Layout {
        x: 0.1,
        y: 0.15,
        w: 0.8,
        h: 0.7,
        layout_target: Some(LayoutTarget::Inner),
    });
    let (x, y) = cat_value_pair();
    c.x_axis = Some(x);
    c.y_axis = Some(y);
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<c:manualLayout>");
    assert_contains(&bytes, "<c:layoutTarget val=\"inner\"/>");
    assert_contains(&bytes, "<c:x val=\"0.1\"/>");
    assert_contains(&bytes, "<c:y val=\"0.15\"/>");
    assert_contains(&bytes, "<c:w val=\"0.8\"/>");
    assert_contains(&bytes, "<c:h val=\"0.7\"/>");
}

#[test]
fn title_with_two_runs_emits_both() {
    let mut c = Chart::new(ChartKind::Bar, anchor_at(0, 0));
    c.title = Some(Title {
        runs: vec![
            TitleRun {
                text: "Q4 ".to_string(),
                bold: Some(true),
                italic: None,
                underline: None,
                size_pt: Some(14),
                color: Some("FF000000".to_string()),
                font_name: Some("Calibri".to_string()),
            },
            TitleRun {
                text: "Revenue".to_string(),
                bold: None,
                italic: Some(true),
                underline: None,
                size_pt: Some(12),
                color: None,
                font_name: None,
            },
        ],
        overlay: Some(false),
        layout: None,
    });
    c.add_series(series_with_cat_val());
    let (x, y) = cat_value_pair();
    c.x_axis = Some(x);
    c.y_axis = Some(y);
    let bytes = emit_chart_xml(&c);
    parse_ok(&bytes);
    assert_contains(&bytes, "<a:t>Q4 </a:t>");
    assert_contains(&bytes, "<a:t>Revenue</a:t>");
    // First run: bold=1 sz=1400.
    assert_contains(&bytes, " sz=\"1400\"");
    assert_contains(&bytes, " b=\"1\"");
    // Second run: italic=1.
    assert_contains(&bytes, " i=\"1\"");
}
