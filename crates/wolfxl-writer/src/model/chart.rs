//! Chart data model.
//!
//! Pure data only; no I/O, no XML emission. Consumed by
//! [`crate::emit::charts`] which renders one `xl/charts/chartN.xml` per
//! [`Chart`] and by [`crate::emit::drawings`] which anchors them via
//! `<xdr:graphicFrame>`.
//!
//! # Coverage
//!
//! Common openpyxl-compatible chart families are represented, including
//! 2D charts, 3D variants, Stock, Surface, and OfPie / projected pie.
//!
//! Every sub-feature is an `Option<T>` so absent → no XML element
//! emitted. Defaults match openpyxl's "leave the attribute off" rule
//! rather than "emit explicit defaults" — Excel inherits the default
//! either way and byte-parity tests rely on omission.

use super::image::ImageAnchor;

/// Re-export of [`wolfxl_pivot::PivotSource`] so the chart model owns a
/// stable name for downstream callers without forcing them to reach into
/// the pivot crate. This module cannot define the type itself; it only
/// borrows it as an `Option<PivotSource>` on [`Chart`].
pub use wolfxl_pivot::PivotSource;

/// Top-level chart kind. Each variant maps to one OOXML plot-area
/// element name (e.g. `<barChart>`, `<lineChart>`).
///
/// Includes both the original 2D chart families and the later 3D, Stock,
/// Surface, and OfPie families.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartKind {
    Bar,
    Line,
    Pie,
    Doughnut,
    Area,
    Scatter,
    Bubble,
    Radar,
    Bar3D,
    Line3D,
    Pie3D,
    Area3D,
    Surface,
    Surface3D,
    Stock,
    OfPie,
}

impl ChartKind {
    /// The element name (without namespace prefix) emitted inside
    /// `<plotArea>` for this kind.
    pub fn plot_element_name(self) -> &'static str {
        match self {
            ChartKind::Bar => "barChart",
            ChartKind::Line => "lineChart",
            ChartKind::Pie => "pieChart",
            ChartKind::Doughnut => "doughnutChart",
            ChartKind::Area => "areaChart",
            ChartKind::Scatter => "scatterChart",
            ChartKind::Bubble => "bubbleChart",
            ChartKind::Radar => "radarChart",
            ChartKind::Bar3D => "bar3DChart",
            ChartKind::Line3D => "line3DChart",
            ChartKind::Pie3D => "pie3DChart",
            ChartKind::Area3D => "area3DChart",
            ChartKind::Surface => "surfaceChart",
            ChartKind::Surface3D => "surface3DChart",
            ChartKind::Stock => "stockChart",
            ChartKind::OfPie => "ofPieChart",
        }
    }

    /// True when this kind uses a category + value axis pair (catAx +
    /// valAx). False for kinds like Pie/Doughnut which have neither and
    /// for Scatter/Bubble which use two value axes.
    pub fn has_category_axis(self) -> bool {
        matches!(
            self,
            ChartKind::Bar
                | ChartKind::Line
                | ChartKind::Area
                | ChartKind::Radar
                | ChartKind::Bar3D
                | ChartKind::Line3D
                | ChartKind::Area3D
                | ChartKind::Surface
                | ChartKind::Surface3D
                | ChartKind::Stock
        )
    }

    /// True when this kind uses two value axes (no category axis).
    pub fn has_dual_value_axes(self) -> bool {
        matches!(self, ChartKind::Scatter | ChartKind::Bubble)
    }

    /// True when this kind has no axes at all (Pie/Doughnut/Pie3D/OfPie).
    pub fn is_axis_free(self) -> bool {
        matches!(
            self,
            ChartKind::Pie | ChartKind::Doughnut | ChartKind::Pie3D | ChartKind::OfPie
        )
    }

    /// True for 3D chart variants — they emit a top-level `<c:view3D>`
    /// element and use the 3D plot-area element name.
    pub fn is_3d(self) -> bool {
        matches!(
            self,
            ChartKind::Bar3D
                | ChartKind::Line3D
                | ChartKind::Pie3D
                | ChartKind::Area3D
                | ChartKind::Surface3D
        )
    }

    /// True for Surface/Surface3D — emit `<c:wireframe/>` if requested.
    pub fn is_surface(self) -> bool {
        matches!(self, ChartKind::Surface | ChartKind::Surface3D)
    }

    /// True for Pie family (Pie, Doughnut, Pie3D, OfPie).
    pub fn is_pie_family(self) -> bool {
        matches!(
            self,
            ChartKind::Pie | ChartKind::Doughnut | ChartKind::Pie3D | ChartKind::OfPie
        )
    }
}

/// A reference to a contiguous range on one sheet. Used by series
/// titles, categories, values, x/y, and bubble-size.
///
/// `cell_range` is an A1 fragment — either a single cell (`"B1"`) or a
/// rectangular range (`"A2:A6"`). The emitter quotes the sheet name and
/// emits absolute refs; e.g. `Reference { sheet_name: "Sheet1",
/// cell_range: "A2:A6" }` → `'Sheet1'!$A$2:$A$6`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Reference {
    pub sheet_name: String,
    pub cell_range: String,
}

impl Reference {
    pub fn new(sheet_name: impl Into<String>, cell_range: impl Into<String>) -> Self {
        Self {
            sheet_name: sheet_name.into(),
            cell_range: cell_range.into(),
        }
    }

    /// Render as `'Sheet'!$A$2:$A$6`. Sheet name is always single-quoted
    /// — Excel accepts unquoted only for plain ASCII identifiers, and
    /// always-quote keeps the formula round-trip safe for spaces /
    /// non-ASCII.
    ///
    /// The cell range is upgraded to absolute by inserting `$` before
    /// every column letter and row number, *if* not already absolute.
    /// Already-absolute fragments pass through unchanged.
    pub fn to_formula_string(&self) -> String {
        let abs = absolutize_a1(&self.cell_range);
        format!("'{}'!{}", self.sheet_name.replace('\'', "''"), abs)
    }

    /// Render the reference as openpyxl serializes series title refs:
    /// single cells stay relative (`'Sheet'!B1`), while ranges use the
    /// normal absolute formula representation.
    pub fn to_series_title_formula_string(&self) -> String {
        if self.cell_range.contains(':') {
            return self.to_formula_string();
        }
        let rel = self.cell_range.replace('$', "");
        format!("'{}'!{}", self.sheet_name.replace('\'', "''"), rel)
    }
}

/// Insert `$` before column letters and row numbers in an A1 fragment if
/// not already present. Handles both `A1` and `A1:B2` shapes. Leaves
/// invalid input alone (the emitter is lenient — bad input becomes a
/// non-functional chart, not a panic).
fn absolutize_a1(s: &str) -> String {
    if let Some((lhs, rhs)) = s.split_once(':') {
        format!("{}:{}", absolutize_one(lhs), absolutize_one(rhs))
    } else {
        absolutize_one(s)
    }
}

fn absolutize_one(s: &str) -> String {
    let mut out = String::with_capacity(s.len() + 2);
    let mut chars = s.chars().peekable();
    let mut started_letters = false;
    let mut started_digits = false;
    // Skip a leading '$' if present, then re-add it.
    if chars.peek() == Some(&'$') {
        chars.next();
    }
    out.push('$');
    while let Some(ch) = chars.next() {
        if ch == '$' {
            // Skip an interior '$' before digits — we'll re-add it.
            continue;
        }
        if ch.is_ascii_alphabetic() {
            out.push(ch);
            started_letters = true;
        } else if ch.is_ascii_digit() {
            if started_letters && !started_digits {
                out.push('$');
                started_digits = true;
            }
            out.push(ch);
        } else {
            // Anything else: bail out, return original.
            return s.to_string();
        }
    }
    out
}

/// One series in a chart. `idx` and `order` are typically equal and
/// 0-based; openpyxl emits `<idx val="0"/>` and `<order val="0"/>` for
/// the first series.
#[derive(Debug, Clone, PartialEq)]
pub struct Series {
    pub idx: u32,
    pub order: u32,

    /// Series title — either a string-reference to a header cell, or a
    /// literal string. `None` produces no `<tx>` element.
    pub title: Option<SeriesTitle>,

    /// Category axis values (X for Bar/Line/Area/Radar). `None` produces
    /// no `<cat>` element. Category axes accept `strRef` (text labels)
    /// or `numRef`; we always emit `numRef` matching openpyxl.
    pub categories: Option<Reference>,

    /// Y-axis values (or numeric values for Pie/Doughnut). Required for
    /// most kinds.
    pub values: Option<Reference>,

    /// X-axis values. Required for Scatter/Bubble; ignored for others.
    pub x_values: Option<Reference>,

    /// Bubble size values. Required for Bubble.
    pub bubble_size: Option<Reference>,

    /// Per-series graphical properties (line/fill/dash on the series
    /// shape itself).
    pub graphical_properties: Option<GraphicalProperties>,

    /// Marker (Line/Scatter/Radar). Has no effect on Bar/Pie/Area.
    pub marker: Option<Marker>,

    /// Per-data-point overrides (`<c:dPt>`): point-specific colors,
    /// markers, explosion offsets, and related flags.
    pub data_points: Vec<DataPoint>,

    /// Per-series data labels.
    pub data_labels: Option<DataLabels>,

    /// Error bars (one entry per direction; openpyxl supports plus,
    /// minus, both).
    pub error_bars: Vec<ErrorBars>,

    /// Trendlines.
    pub trendlines: Vec<Trendline>,

    /// Smooth flag (Line/Scatter only).
    pub smooth: Option<bool>,

    /// Invert if negative (Bar/Bubble only — paint negative bars in a
    /// different color).
    pub invert_if_negative: Option<bool>,
}

impl Series {
    pub fn new(idx: u32) -> Self {
        Self {
            idx,
            order: idx,
            title: None,
            categories: None,
            values: None,
            x_values: None,
            bubble_size: None,
            graphical_properties: None,
            marker: None,
            data_points: Vec::new(),
            data_labels: None,
            error_bars: Vec::new(),
            trendlines: Vec::new(),
            smooth: None,
            invert_if_negative: None,
        }
    }
}

/// A series title — either pinned to a header cell or a plain literal
/// string. Pinning to a cell is more common and matches what users get
/// when they construct a chart from a labelled column in Excel.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum SeriesTitle {
    /// `<tx><strRef><f>'Sheet'!B1</f></strRef></tx>`
    StrRef(Reference),
    /// `<tx><v>literal text</v></tx>` — emitted as a `<rich>` block by
    /// the emitter.
    Literal(String),
}

/// Chart axis. The four flavours match OOXML's four element names. The
/// emitter dispatches on the variant.
#[derive(Debug, Clone, PartialEq)]
pub enum Axis {
    Category(CategoryAxis),
    Value(ValueAxis),
    Date(DateAxis),
    Series(SeriesAxis),
}

/// Common axis fields that every axis flavour shares.
#[derive(Debug, Clone, PartialEq)]
pub struct AxisCommon {
    /// 1-based numeric id used to cross-link axes (`<axId val="10"/>`).
    pub ax_id: u32,
    /// Crosses-axis id pointing at the partner axis (`<crossAx val="100"/>`).
    pub cross_ax: u32,
    /// `<orientation val="minMax"/>` (default) or `"maxMin"` (reversed).
    pub orientation: AxisOrientation,
    /// `<axPos val="…"/>` — `b` (bottom), `t` (top), `l` (left), `r` (right).
    pub ax_pos: AxisPos,
    /// `<delete val="1"/>` to hide an axis. `None` → no element.
    pub delete: Option<bool>,
    /// `<majorTickMark val="…"/>` — `none`, `in`, `out`, `cross`.
    pub major_tick_mark: Option<TickMark>,
    /// `<minorTickMark val="…"/>` — same enum.
    pub minor_tick_mark: Option<TickMark>,
    /// Axis title (rich-text label). `None` → no `<title>` block.
    pub title: Option<Title>,
    /// `<majorGridlines/>` present when true (legacy short-form flag).
    /// Set to `true` to emit a default `<c:majorGridlines/>`. To attach
    /// graphical properties, use [`Self::major_gridlines_obj`] instead;
    /// when both are set, `major_gridlines_obj` takes precedence.
    pub major_gridlines: bool,
    /// `<minorGridlines/>` present when true (legacy short-form flag).
    pub minor_gridlines: bool,
    /// `<majorGridlines>` rich form (with optional graphical properties).
    pub major_gridlines_obj: Option<Gridlines>,
    /// `<minorGridlines>` rich form.
    pub minor_gridlines_obj: Option<Gridlines>,
    /// `<numFmt formatCode="…" sourceLinked="0"/>` — explicit number format.
    pub number_format: Option<String>,
}

impl AxisCommon {
    pub fn new(ax_id: u32, cross_ax: u32, ax_pos: AxisPos) -> Self {
        Self {
            ax_id,
            cross_ax,
            orientation: AxisOrientation::MinMax,
            ax_pos,
            delete: None,
            major_tick_mark: None,
            minor_tick_mark: None,
            title: None,
            major_gridlines: false,
            minor_gridlines: false,
            major_gridlines_obj: None,
            minor_gridlines_obj: None,
            number_format: None,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AxisOrientation {
    MinMax,
    MaxMin,
}

impl AxisOrientation {
    pub fn as_str(self) -> &'static str {
        match self {
            AxisOrientation::MinMax => "minMax",
            AxisOrientation::MaxMin => "maxMin",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AxisPos {
    Bottom,
    Top,
    Left,
    Right,
}

impl AxisPos {
    pub fn as_str(self) -> &'static str {
        match self {
            AxisPos::Bottom => "b",
            AxisPos::Top => "t",
            AxisPos::Left => "l",
            AxisPos::Right => "r",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TickMark {
    None,
    In,
    Out,
    Cross,
}

impl TickMark {
    pub fn as_str(self) -> &'static str {
        match self {
            TickMark::None => "none",
            TickMark::In => "in",
            TickMark::Out => "out",
            TickMark::Cross => "cross",
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct CategoryAxis {
    pub common: AxisCommon,
    /// `<lblOffset val="100"/>` (default).
    pub lbl_offset: Option<u32>,
    /// `<lblAlgn val="ctr"/>` — `ctr`, `l`, or `r`.
    pub lbl_algn: Option<String>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct ValueAxis {
    pub common: AxisCommon,
    /// `<scaling><min val="…"/></scaling>` — explicit min override.
    pub min: Option<f64>,
    pub max: Option<f64>,
    pub major_unit: Option<f64>,
    pub minor_unit: Option<f64>,
    /// `<dispUnits>` for display units such as thousands or millions.
    pub display_units: Option<DisplayUnits>,
    /// `<crosses val="…"/>` — `autoZero`, `min`, `max`.
    pub crosses: Option<String>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct DateAxis {
    pub common: AxisCommon,
    pub min: Option<f64>,
    pub max: Option<f64>,
    pub major_unit: Option<f64>,
    pub minor_unit: Option<f64>,
    /// `<baseTimeUnit val="…"/>` — `days`, `months`, `years`.
    pub base_time_unit: Option<String>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct SeriesAxis {
    pub common: AxisCommon,
}

/// Chart title.
#[derive(Debug, Clone, PartialEq)]
pub struct Title {
    /// One or more rich-text runs. The emitter wraps them in `<rich>`.
    pub runs: Vec<TitleRun>,
    /// `<overlay val="0"/>` controls whether the title overlaps the
    /// plot area.
    pub overlay: Option<bool>,
    /// Manual layout (EMU coords). `None` → no `<layout>`.
    pub layout: Option<Layout>,
}

impl Title {
    /// Convenience constructor for a single-run title.
    pub fn plain(text: impl Into<String>) -> Self {
        Self {
            runs: vec![TitleRun::plain(text)],
            overlay: None,
            layout: None,
        }
    }
}

/// 3D chart view parameters — emitted as `<c:view3D>` at the chart level
/// before `<plotArea>` for 3D variants.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct View3D {
    /// `<c:rotX val="…"/>` — typically -90..90 (or 0..30 for Bar3D).
    pub rot_x: Option<i16>,
    /// `<c:rotY val="…"/>` — 0..360.
    pub rot_y: Option<i16>,
    /// `<c:perspective val="…"/>` — 0..240.
    pub perspective: Option<u8>,
    /// `<c:rAngAx val="1"/>` — orthogonal axes flag.
    pub right_angle_axes: Option<bool>,
    /// `<c:autoScale val="1"/>` — auto scaling.
    pub auto_scale: Option<bool>,
    /// `<c:depthPercent val="…"/>` — 20..2000.
    pub depth_percent: Option<u32>,
    /// `<c:hPercent val="…"/>` — 5..500.
    pub h_percent: Option<u32>,
}

/// `<c:majorGridlines>` / `<c:minorGridlines>` content.
///
/// `None` at the parent axis means "no gridlines". Empty `Gridlines`
/// (default) means "draw default gridlines" (an empty self-closing
/// element is emitted). Optional `graphical_properties` paints the
/// gridline shape.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct Gridlines {
    pub graphical_properties: Option<GraphicalProperties>,
}

/// `<c:upDownBars>` on a stock chart.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct UpDownBars {
    pub gap_width: Option<u32>,
    pub up_bars: Option<Gridlines>,
    pub down_bars: Option<Gridlines>,
}

/// One run of rich text inside a title (or per-cell rich label).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TitleRun {
    pub text: String,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<bool>,
    pub size_pt: Option<u32>,
    /// 8-char ARGB color (e.g. `"FF000000"`).
    pub color: Option<String>,
    pub font_name: Option<String>,
}

impl TitleRun {
    pub fn plain(text: impl Into<String>) -> Self {
        Self {
            text: text.into(),
            bold: None,
            italic: None,
            underline: None,
            size_pt: None,
            color: None,
            font_name: None,
        }
    }
}

/// Chart legend.
#[derive(Debug, Clone, PartialEq)]
pub struct Legend {
    pub position: LegendPosition,
    pub overlay: Option<bool>,
    pub layout: Option<Layout>,
}

impl Default for Legend {
    fn default() -> Self {
        Self {
            position: LegendPosition::Right,
            overlay: None,
            layout: None,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LegendPosition {
    Right,
    Left,
    Top,
    Bottom,
    TopRight,
}

impl LegendPosition {
    pub fn as_str(self) -> &'static str {
        match self {
            LegendPosition::Right => "r",
            LegendPosition::Left => "l",
            LegendPosition::Top => "t",
            LegendPosition::Bottom => "b",
            LegendPosition::TopRight => "tr",
        }
    }
}

/// Manual layout for a title, legend, or plot area. Coordinates are
/// 0..1 fractions of the chart's drawing surface (this matches OOXML —
/// it does *not* use EMU here despite the rest of drawingml).
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Layout {
    pub x: f64,
    pub y: f64,
    pub w: f64,
    pub h: f64,
    /// `inner` (default) or `outer` for layoutTarget.
    pub layout_target: Option<LayoutTarget>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LayoutTarget {
    Inner,
    Outer,
}

impl LayoutTarget {
    pub fn as_str(self) -> &'static str {
        match self {
            LayoutTarget::Inner => "inner",
            LayoutTarget::Outer => "outer",
        }
    }
}

/// Per-series data labels.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct DataLabels {
    pub show_val: Option<bool>,
    pub show_cat_name: Option<bool>,
    pub show_ser_name: Option<bool>,
    pub show_percent: Option<bool>,
    pub show_legend_key: Option<bool>,
    pub show_bubble_size: Option<bool>,
    /// `<dLblPos val="…"/>` — `ctr`, `b`, `t`, `r`, `l`, `outEnd`,
    /// `inEnd`, `inBase`, `bestFit`.
    pub position: Option<String>,
    pub number_format: Option<String>,
    /// Custom separator between fields (e.g. `","`, `";"`).
    pub separator: Option<String>,
    /// Optional rich-text runs for `<c:txPr>` so labels render with
    /// per-run formatting (bold/italic/color/size/font). Each run shares
    /// the chart-title `TitleRun` shape; an empty vector emits no
    /// `<c:txPr>` block.
    pub tx_pr_runs: Vec<TitleRun>,
}

/// Per-series error bars.
#[derive(Debug, Clone, PartialEq)]
pub struct ErrorBars {
    /// `<errBarType val="…"/>` — `plus`, `minus`, or `both`.
    pub bar_type: ErrorBarType,
    /// `<errValType val="…"/>` — `cust`, `fixedVal`, `percentage`,
    /// `stdDev`, `stdErr`.
    pub val_type: ErrorBarValType,
    /// Numeric value for `fixedVal`, `percentage`, `stdDev`. Ignored
    /// for `cust` (custom uses `plus`/`minus` references) and
    /// `stdErr`.
    pub value: Option<f64>,
    /// `<noEndCap val="1"/>` to hide the cross-bars at the ends.
    pub no_end_cap: Option<bool>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ErrorBarType {
    Plus,
    Minus,
    Both,
}

impl ErrorBarType {
    pub fn as_str(self) -> &'static str {
        match self {
            ErrorBarType::Plus => "plus",
            ErrorBarType::Minus => "minus",
            ErrorBarType::Both => "both",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ErrorBarValType {
    Cust,
    FixedVal,
    Percentage,
    StdDev,
    StdErr,
}

impl ErrorBarValType {
    pub fn as_str(self) -> &'static str {
        match self {
            ErrorBarValType::Cust => "cust",
            ErrorBarValType::FixedVal => "fixedVal",
            ErrorBarValType::Percentage => "percentage",
            ErrorBarValType::StdDev => "stdDev",
            ErrorBarValType::StdErr => "stdErr",
        }
    }
}

/// Per-series trendline.
#[derive(Debug, Clone, PartialEq)]
pub struct Trendline {
    pub kind: TrendlineKind,
    /// Polynomial order (only when `kind == Polynomial`).
    pub order: Option<u32>,
    /// Moving-average period (only when `kind == MovingAvg`).
    pub period: Option<u32>,
    /// Forecast forward (data units).
    pub forward: Option<f64>,
    /// Forecast backward.
    pub backward: Option<f64>,
    pub display_equation: Option<bool>,
    pub display_r_squared: Option<bool>,
    pub name: Option<String>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TrendlineKind {
    Linear,
    Log,
    Power,
    Exp,
    Polynomial,
    MovingAvg,
}

impl TrendlineKind {
    pub fn as_str(self) -> &'static str {
        match self {
            TrendlineKind::Linear => "linear",
            TrendlineKind::Log => "log",
            TrendlineKind::Power => "power",
            TrendlineKind::Exp => "exp",
            TrendlineKind::Polynomial => "poly",
            TrendlineKind::MovingAvg => "movingAvg",
        }
    }
}

/// Per-series marker (Line/Scatter/Radar).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Marker {
    pub symbol: MarkerSymbol,
    pub size: Option<u32>,
    pub graphical_properties: Option<GraphicalProperties>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DataPoint {
    pub idx: u32,
    pub invert_if_negative: Option<bool>,
    pub marker: Option<Marker>,
    pub bubble_3d: Option<bool>,
    pub explosion: Option<u32>,
    pub graphical_properties: Option<GraphicalProperties>,
}

impl DataPoint {
    pub fn new(idx: u32) -> Self {
        Self {
            idx,
            invert_if_negative: None,
            marker: None,
            bubble_3d: None,
            explosion: None,
            graphical_properties: None,
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct DisplayUnits {
    pub built_in_unit: Option<String>,
    pub custom_unit: Option<f64>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MarkerSymbol {
    None,
    Circle,
    Square,
    Diamond,
    Triangle,
    Plus,
    X,
    Star,
    Dash,
    Dot,
    Auto,
}

impl MarkerSymbol {
    pub fn as_str(self) -> &'static str {
        match self {
            MarkerSymbol::None => "none",
            MarkerSymbol::Circle => "circle",
            MarkerSymbol::Square => "square",
            MarkerSymbol::Diamond => "diamond",
            MarkerSymbol::Triangle => "triangle",
            MarkerSymbol::Plus => "plus",
            MarkerSymbol::X => "x",
            MarkerSymbol::Star => "star",
            MarkerSymbol::Dash => "dash",
            MarkerSymbol::Dot => "dot",
            MarkerSymbol::Auto => "auto",
        }
    }
}

/// Drawing-ml graphical properties (line + fill) shared between series
/// shapes, marker shapes, axis title boxes, etc.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct GraphicalProperties {
    /// 8-char ARGB outline color (`"FF000000"`).
    pub line_color: Option<String>,
    /// Line width in EMU (12700 EMU per pt).
    pub line_width_emu: Option<u32>,
    /// `<a:prstDash val="…"/>` — `solid`, `dash`, `dashDot`,
    /// `dot`, `lgDash`, `lgDashDot`, `lgDashDotDot`, `sysDash`,
    /// `sysDashDot`, `sysDashDotDot`, `sysDot`.
    pub line_dash: Option<String>,
    /// 8-char ARGB fill color.
    pub fill_color: Option<String>,
    /// `<a:noFill/>` instead of any fill.
    pub no_fill: bool,
    /// `<a:ln><a:noFill/></a:ln>` — explicit no-line override.
    pub no_line: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DisplayBlanksAs {
    Gap,
    Span,
    Zero,
}

impl DisplayBlanksAs {
    pub fn as_str(self) -> &'static str {
        match self {
            DisplayBlanksAs::Gap => "gap",
            DisplayBlanksAs::Span => "span",
            DisplayBlanksAs::Zero => "zero",
        }
    }
}

/// Bar-specific direction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BarDir {
    Col,
    Bar,
}

impl BarDir {
    pub fn as_str(self) -> &'static str {
        match self {
            BarDir::Col => "col",
            BarDir::Bar => "bar",
        }
    }
}

/// Bar grouping (`<grouping val="…"/>`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BarGrouping {
    Clustered,
    Stacked,
    PercentStacked,
    Standard,
}

impl BarGrouping {
    pub fn as_str(self) -> &'static str {
        match self {
            BarGrouping::Clustered => "clustered",
            BarGrouping::Stacked => "stacked",
            BarGrouping::PercentStacked => "percentStacked",
            BarGrouping::Standard => "standard",
        }
    }
}

/// Scatter style (`<scatterStyle val="…"/>`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ScatterStyle {
    Line,
    LineMarker,
    Marker,
    Smooth,
    SmoothMarker,
    None,
}

impl ScatterStyle {
    pub fn as_str(self) -> &'static str {
        match self {
            ScatterStyle::Line => "line",
            ScatterStyle::LineMarker => "lineMarker",
            ScatterStyle::Marker => "marker",
            ScatterStyle::Smooth => "smooth",
            ScatterStyle::SmoothMarker => "smoothMarker",
            ScatterStyle::None => "none",
        }
    }
}

/// Radar style (`<radarStyle val="…"/>`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RadarStyle {
    Standard,
    Marker,
    Filled,
}

impl RadarStyle {
    pub fn as_str(self) -> &'static str {
        match self {
            RadarStyle::Standard => "standard",
            RadarStyle::Marker => "marker",
            RadarStyle::Filled => "filled",
        }
    }
}

/// One chart instance attached to a sheet.
#[derive(Debug, Clone, PartialEq)]
pub struct Chart {
    pub kind: ChartKind,

    /// Optional chart title.
    pub title: Option<Title>,

    /// Optional legend (`None` → no `<legend>` element emitted).
    pub legend: Option<Legend>,

    /// Optional explicit plot-area layout.
    pub layout: Option<Layout>,

    /// Primary axes. The exact pair depends on the kind:
    ///
    /// - Bar/Line/Area/Radar: x = Category, y = Value
    /// - Scatter/Bubble: x = Value, y = Value
    /// - Pie/Doughnut: both `None` (no axes)
    pub x_axis: Option<Axis>,
    pub y_axis: Option<Axis>,

    /// One or more series.
    pub series: Vec<Series>,

    /// Chart-level data labels (`<dLbls>` inside the chart-kind block).
    pub data_labels: Option<DataLabels>,

    /// `<plotVisOnly val="1"/>` (default `true` per openpyxl).
    pub plot_visible_only: Option<bool>,

    /// `<dispBlanksAs val="gap"/>`.
    pub display_blanks_as: Option<DisplayBlanksAs>,

    /// `<varyColors val="1"/>` — Pie/Doughnut typically true so each
    /// slice is a different color.
    pub vary_colors: Option<bool>,

    /// Anchor on the worksheet — same enum as image anchors so we can
    /// pin charts to cells the same way.
    pub anchor: ImageAnchor,

    /// Bar-specific: `<barDir val="…"/>`. Required for Bar.
    pub bar_dir: Option<BarDir>,

    /// Bar / Area: `<grouping val="…"/>`.
    pub grouping: Option<BarGrouping>,

    /// Bar: `<gapWidth val="…"/>` (0..500, default 150).
    pub gap_width: Option<u32>,

    /// Bar: `<overlap val="…"/>` (-100..100, default 0). Negative values
    /// space bars apart, positive values overlap them.
    pub overlap: Option<i32>,

    /// Doughnut: `<holeSize val="…"/>` (10..90, default 50).
    pub hole_size: Option<u32>,

    /// Pie: `<firstSliceAng val="…"/>` (0..360).
    pub first_slice_ang: Option<u32>,

    /// Scatter: `<scatterStyle val="…"/>`.
    pub scatter_style: Option<ScatterStyle>,

    /// Radar: `<radarStyle val="…"/>`.
    pub radar_style: Option<RadarStyle>,

    /// Bubble: `<bubble3D val="0"/>` controls 3D bubble shading.
    pub bubble3d: Option<bool>,

    /// Bubble: `<bubbleScale val="…"/>` (0..300).
    pub bubble_scale: Option<u32>,

    /// Bubble: `<showNegBubbles val="0"/>`.
    pub show_neg_bubbles: Option<bool>,

    /// Optional `style` element on `<chart>` (1..48). Excel 2007 looks.
    pub style: Option<u32>,

    /// `<smooth val="1"/>` on Line at the chart level (rare; usually
    /// per-series).
    pub smoothing: Option<bool>,

    /// Stock chart-group decorations. `None` omits the element; an empty
    /// shape retains default styling.
    pub hi_low_lines: Option<Gridlines>,
    pub up_down_bars: Option<UpDownBars>,

    /// 3D view parameters; only emitted when `kind.is_3d()`.
    pub view_3d: Option<View3D>,

    /// Surface chart wireframe toggle.
    pub wireframe: Option<bool>,

    /// `<c:ofPieType val="bar|pie"/>` for OfPie kind.
    pub of_pie_type: Option<String>,

    /// `<c:splitType val="auto|cust|percent|pos|val"/>` for OfPie kind.
    pub split_type: Option<String>,

    /// `<c:splitPos val="…"/>` for OfPie when `split_type` requires a
    /// numeric split point.
    pub split_pos: Option<f64>,

    /// `<c:secondPieSize val="…"/>` for OfPie (5..200).
    pub second_pie_size: Option<u32>,

    /// When `Some`, the chart is a pivot chart and the emitter writes a
    /// `<c:pivotSource>` block inside `<c:chart>` (between the `<c:chart>`
    /// open and the title) and injects a `<c:fmtId val="0"/>` element on
    /// every series.
    pub pivot_source: Option<PivotSource>,

    /// Sibling chart families to emit inside the same `<plotArea>`
    /// (RFC-069 / G15 — combination charts). Each entry is a fully-formed
    /// `Chart` whose outer-frame fields (title, legend, anchor, dimensions)
    /// are ignored by the combination-chart emit path; only per-family
    /// fields (`kind`, `series`, type-specific knobs, `y_axis`) are
    /// consumed. Default is an empty vec — every existing single-family
    /// caller is unaffected.
    pub secondary_charts: Vec<Chart>,
}

impl Chart {
    /// Create a minimal chart with the given kind anchored at one cell.
    pub fn new(kind: ChartKind, anchor: ImageAnchor) -> Self {
        Self {
            kind,
            title: None,
            legend: Some(Legend::default()),
            layout: None,
            x_axis: None,
            y_axis: None,
            series: Vec::new(),
            data_labels: None,
            plot_visible_only: Some(true),
            display_blanks_as: Some(DisplayBlanksAs::Gap),
            vary_colors: None,
            anchor,
            bar_dir: if matches!(kind, ChartKind::Bar) {
                Some(BarDir::Col)
            } else {
                None
            },
            grouping: if matches!(kind, ChartKind::Bar | ChartKind::Area) {
                Some(BarGrouping::Clustered)
            } else {
                None
            },
            gap_width: if matches!(kind, ChartKind::Bar) {
                Some(150)
            } else {
                None
            },
            overlap: None,
            hole_size: if matches!(kind, ChartKind::Doughnut) {
                Some(50)
            } else {
                None
            },
            first_slice_ang: None,
            scatter_style: None,
            radar_style: if matches!(kind, ChartKind::Radar) {
                Some(RadarStyle::Standard)
            } else {
                None
            },
            bubble3d: None,
            bubble_scale: None,
            show_neg_bubbles: None,
            style: None,
            smoothing: None,
            hi_low_lines: None,
            up_down_bars: None,
            view_3d: if matches!(
                kind,
                ChartKind::Bar3D
                    | ChartKind::Line3D
                    | ChartKind::Pie3D
                    | ChartKind::Area3D
                    | ChartKind::Surface3D
            ) {
                Some(View3D::default())
            } else {
                None
            },
            wireframe: None,
            of_pie_type: if matches!(kind, ChartKind::OfPie) {
                Some("pie".to_string())
            } else {
                None
            },
            split_type: if matches!(kind, ChartKind::OfPie) {
                Some("auto".to_string())
            } else {
                None
            },
            split_pos: None,
            second_pie_size: None,
            pivot_source: None,
            secondary_charts: Vec::new(),
        }
    }

    /// Append a series.
    pub fn add_series(&mut self, series: Series) {
        self.series.push(series);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn reference_to_formula_quotes_sheet_and_absolutizes() {
        let r = Reference::new("Sheet1", "A2:A6");
        assert_eq!(r.to_formula_string(), "'Sheet1'!$A$2:$A$6");
    }

    #[test]
    fn reference_single_cell() {
        let r = Reference::new("Data", "B1");
        assert_eq!(r.to_formula_string(), "'Data'!$B$1");
    }

    #[test]
    fn reference_series_title_single_cell_is_relative() {
        let r = Reference::new("Data", "B1");
        assert_eq!(r.to_series_title_formula_string(), "'Data'!B1");
    }

    #[test]
    fn reference_already_absolute_passes_through() {
        let r = Reference::new("Sheet", "$A$2:$A$6");
        assert_eq!(r.to_formula_string(), "'Sheet'!$A$2:$A$6");
    }

    #[test]
    fn reference_sheet_with_apostrophe_doubles_it() {
        let r = Reference::new("It's mine", "A1");
        assert_eq!(r.to_formula_string(), "'It''s mine'!$A$1");
    }

    #[test]
    fn chart_kind_plot_element_names() {
        let cases = [
            (ChartKind::Bar, "barChart"),
            (ChartKind::Line, "lineChart"),
            (ChartKind::Pie, "pieChart"),
            (ChartKind::Doughnut, "doughnutChart"),
            (ChartKind::Area, "areaChart"),
            (ChartKind::Scatter, "scatterChart"),
            (ChartKind::Bubble, "bubbleChart"),
            (ChartKind::Radar, "radarChart"),
            (ChartKind::Bar3D, "bar3DChart"),
            (ChartKind::Line3D, "line3DChart"),
            (ChartKind::Pie3D, "pie3DChart"),
            (ChartKind::Area3D, "area3DChart"),
            (ChartKind::Surface, "surfaceChart"),
            (ChartKind::Surface3D, "surface3DChart"),
            (ChartKind::Stock, "stockChart"),
            (ChartKind::OfPie, "ofPieChart"),
        ];
        for (kind, element_name) in cases {
            assert_eq!(kind.plot_element_name(), element_name);
        }
    }

    #[test]
    fn chart_kind_axis_classification() {
        for kind in [
            ChartKind::Bar,
            ChartKind::Line,
            ChartKind::Area,
            ChartKind::Radar,
            ChartKind::Bar3D,
            ChartKind::Line3D,
            ChartKind::Area3D,
            ChartKind::Surface,
            ChartKind::Surface3D,
            ChartKind::Stock,
        ] {
            assert!(kind.has_category_axis(), "{kind:?}");
            assert!(!kind.has_dual_value_axes(), "{kind:?}");
            assert!(!kind.is_axis_free(), "{kind:?}");
        }
        for kind in [ChartKind::Scatter, ChartKind::Bubble] {
            assert!(!kind.has_category_axis(), "{kind:?}");
            assert!(kind.has_dual_value_axes(), "{kind:?}");
            assert!(!kind.is_axis_free(), "{kind:?}");
        }
        for kind in [
            ChartKind::Pie,
            ChartKind::Doughnut,
            ChartKind::Pie3D,
            ChartKind::OfPie,
        ] {
            assert!(!kind.has_category_axis(), "{kind:?}");
            assert!(!kind.has_dual_value_axes(), "{kind:?}");
            assert!(kind.is_axis_free(), "{kind:?}");
        }
    }

    #[test]
    fn new_bar_chart_has_default_bar_dir_grouping_gap() {
        let c = Chart::new(ChartKind::Bar, ImageAnchor::one_cell(0, 0));
        assert_eq!(c.bar_dir, Some(BarDir::Col));
        assert_eq!(c.grouping, Some(BarGrouping::Clustered));
        assert_eq!(c.gap_width, Some(150));
    }

    #[test]
    fn new_doughnut_chart_has_default_hole_size() {
        let c = Chart::new(ChartKind::Doughnut, ImageAnchor::one_cell(0, 0));
        assert_eq!(c.hole_size, Some(50));
    }

    #[test]
    fn new_scatter_omits_scatter_style_by_default() {
        let c = Chart::new(ChartKind::Scatter, ImageAnchor::one_cell(0, 0));
        assert_eq!(c.scatter_style, None);
    }

    #[test]
    fn new_radar_has_default_radar_style() {
        let c = Chart::new(ChartKind::Radar, ImageAnchor::one_cell(0, 0));
        assert_eq!(c.radar_style, Some(RadarStyle::Standard));
    }

    #[test]
    fn series_new_zeroes_optional_fields() {
        let s = Series::new(0);
        assert_eq!(s.idx, 0);
        assert_eq!(s.order, 0);
        assert!(s.title.is_none());
        assert!(s.values.is_none());
        assert!(s.error_bars.is_empty());
        assert!(s.trendlines.is_empty());
    }
}
