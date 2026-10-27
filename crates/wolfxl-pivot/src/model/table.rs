//! `PivotTable` — the layout half of an OOXML pivot.
//!
//! Mirrors RFC-048 §10. One `PivotTable` represents
//! `xl/pivotTables/pivotTable{N}.xml`; references a `PivotCache`
//! by `cache_id`.

/// `<location>` element. RFC-048 §10.2.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Location {
    /// A1-style range, e.g. `"F2:I20"`.
    pub range: String,
    pub first_header_row: u32,
    pub first_data_row: u32,
    pub first_data_col: u32,
    pub row_page_count: Option<u32>,
    pub col_page_count: Option<u32>,
}

/// Per-cache-field appearance/role in the pivot. One entry per cache
/// field, in cache-field order. RFC-048 §10.3.
#[derive(Debug, Clone, PartialEq)]
pub struct PivotField {
    pub name: Option<String>,
    pub axis: Option<AxisType>,
    pub data_field: bool,
    pub show_all: bool,
    pub default_subtotal: bool,
    pub sum_subtotal: bool,
    pub count_subtotal: bool,
    pub avg_subtotal: bool,
    pub max_subtotal: bool,
    pub min_subtotal: bool,
    pub items: Option<Vec<PivotItem>>,
    pub outline: bool,
    pub compact: bool,
    pub subtotal_top: bool,
}

impl Default for PivotField {
    fn default() -> Self {
        Self {
            name: None,
            axis: None,
            data_field: false,
            show_all: false,
            default_subtotal: true,
            sum_subtotal: false,
            count_subtotal: false,
            avg_subtotal: false,
            max_subtotal: false,
            min_subtotal: false,
            items: None,
            outline: true,
            compact: true,
            subtotal_top: true,
        }
    }
}

/// Pivot-field axis assignment. RFC-048 §10.3.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AxisType {
    Row,
    Col,
    Page,
    /// `"axisValues"` — pseudo-axis used when the field acts only as
    /// a data field.
    Values,
}

impl AxisType {
    pub fn xml_attr(&self) -> &'static str {
        match self {
            AxisType::Row => "axisRow",
            AxisType::Col => "axisCol",
            AxisType::Page => "axisPage",
            AxisType::Values => "axisValues",
        }
    }
}

/// A single `<item>` child of `<pivotField>`. RFC-048 §10.4.
#[derive(Debug, Clone, PartialEq)]
pub struct PivotItem {
    pub x: Option<u32>,
    pub t: Option<PivotItemType>,
    pub h: bool,
    pub s: bool,
    pub n: Option<String>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PivotItemType {
    Default,
    Sum,
    Count,
    Avg,
    Max,
    Min,
    Blank,
    Grand,
}

impl PivotItemType {
    pub fn xml_value(&self) -> &'static str {
        match self {
            PivotItemType::Default => "default",
            PivotItemType::Sum => "sum",
            PivotItemType::Count => "count",
            PivotItemType::Avg => "avg",
            PivotItemType::Max => "max",
            PivotItemType::Min => "min",
            PivotItemType::Blank => "blank",
            PivotItemType::Grand => "grand",
        }
    }
}

/// `<dataField>` — aggregation directive. RFC-048 §10.5.
#[derive(Debug, Clone, PartialEq)]
pub struct DataField {
    /// Display name, e.g. `"Sum of revenue"`.
    pub name: String,
    /// Index into `pivot_fields[]`.
    pub field_index: u32,
    pub function: DataFunction,
    pub show_data_as: Option<ShowDataAs>,
    pub base_field: u32,
    pub base_item: u32,
    pub num_fmt_id: Option<u32>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DataFunction {
    Sum,
    Count,
    Average,
    Max,
    Min,
    Product,
    CountNums,
    StdDev,
    StdDevp,
    Var,
    Varp,
}

impl DataFunction {
    pub fn xml_value(&self) -> &'static str {
        match self {
            DataFunction::Sum => "sum",
            DataFunction::Count => "count",
            DataFunction::Average => "average",
            DataFunction::Max => "max",
            DataFunction::Min => "min",
            DataFunction::Product => "product",
            DataFunction::CountNums => "countNums",
            DataFunction::StdDev => "stdDev",
            DataFunction::StdDevp => "stdDevp",
            DataFunction::Var => "var",
            DataFunction::Varp => "varp",
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ShowDataAs {
    Normal,
    Difference,
    Percent,
    PercentDiff,
    RunTotal,
    PercentOfRow,
    PercentOfCol,
    PercentOfTotal,
    Index,
}

/// `<pageField>`. RFC-048 §10.8.
#[derive(Debug, Clone, PartialEq)]
pub struct PageField {
    pub field_index: u32,
    pub name: Option<String>,
    /// Default `0`. `-1` for "(All)".
    pub item_index: i32,
    /// Default `-1` (no hierarchy).
    pub hier: i32,
    pub cap: Option<String>,
}

/// One pre-computed row or column position. RFC-048 §10.6.
#[derive(Debug, Clone, PartialEq)]
pub struct AxisItem {
    /// Path through axis fields — one entry per axis field, each is
    /// a 0-based shared-items index.
    pub indices: Vec<u32>,
    pub t: Option<AxisItemType>,
    /// Run-length compression for leading indices.
    pub r: Option<u32>,
    /// Data-field index when multi-data-field.
    pub i: Option<u32>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AxisItemType {
    Data,
    Default,
    Sum,
    Count,
    Avg,
    Max,
    Min,
    Product,
    CountNums,
    StdDev,
    StdDevp,
    Var,
    Varp,
    Grand,
    Blank,
}

impl AxisItemType {
    pub fn xml_value(&self) -> &'static str {
        match self {
            AxisItemType::Data => "data",
            AxisItemType::Default => "default",
            AxisItemType::Sum => "sum",
            AxisItemType::Count => "count",
            AxisItemType::Avg => "avg",
            AxisItemType::Max => "max",
            AxisItemType::Min => "min",
            AxisItemType::Product => "product",
            AxisItemType::CountNums => "countNums",
            AxisItemType::StdDev => "stdDev",
            AxisItemType::StdDevp => "stdDevp",
            AxisItemType::Var => "var",
            AxisItemType::Varp => "varp",
            AxisItemType::Grand => "grand",
            AxisItemType::Blank => "blank",
        }
    }
}

/// `<pivotTableStyleInfo>`. RFC-048 §10.7.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotTableStyleInfo {
    pub name: String,
    pub show_row_headers: bool,
    pub show_col_headers: bool,
    pub show_row_stripes: bool,
    pub show_col_stripes: bool,
    pub show_last_column: bool,
}

impl Default for PivotTableStyleInfo {
    fn default() -> Self {
        Self {
            name: "PivotStyleLight16".into(),
            show_row_headers: true,
            show_col_headers: true,
            show_row_stripes: false,
            show_col_stripes: false,
            show_last_column: true,
        }
    }
}

/// Top-level pivot-table. Mirrors RFC-048 §10.1.
#[derive(Debug, Clone, PartialEq)]
pub struct PivotTable {
    pub name: String,
    pub cache_id: u32,
    pub location: Location,
    pub pivot_fields: Vec<PivotField>,
    pub row_field_indices: Vec<u32>,
    pub col_field_indices: Vec<u32>,
    pub page_fields: Vec<PageField>,
    pub data_fields: Vec<DataField>,
    pub row_items: Vec<AxisItem>,
    pub col_items: Vec<AxisItem>,
    pub data_on_rows: bool,
    pub outline: bool,
    pub compact: bool,
    pub row_grand_totals: bool,
    pub col_grand_totals: bool,
    pub data_caption: String,
    pub grand_total_caption: Option<String>,
    pub error_caption: Option<String>,
    pub missing_caption: Option<String>,
    pub apply_number_formats: bool,
    pub apply_border_formats: bool,
    pub apply_font_formats: bool,
    pub apply_pattern_formats: bool,
    pub apply_alignment_formats: bool,
    pub apply_width_height_formats: bool,
    pub style_info: Option<PivotTableStyleInfo>,
    pub created_version: u8,
    pub updated_version: u8,
    pub min_refreshable_version: u8,
    /// RFC-061 §2.5 — pivot-area Format directives.
    pub formats: Vec<Format>,
    /// RFC-061 §2.5 — pivot-scoped CF.
    pub conditional_formats: Vec<PivotConditionalFormat>,
}

impl PivotTable {
    /// RFC-048 §10.9 validation. Returns first violation.
    pub fn validate(&self) -> Result<(), String> {
        if self.data_fields.is_empty() {
            return Err("PivotTable requires ≥1 data field".into());
        }
        let nfields = self.pivot_fields.len() as u32;
        for &i in &self.row_field_indices {
            if i >= nfields {
                return Err(format!("row field index out of range: {i}"));
            }
        }
        for &i in &self.col_field_indices {
            if i >= nfields {
                return Err(format!("col field index out of range: {i}"));
            }
        }
        // Same field on multiple axes?
        for &i in &self.row_field_indices {
            if self.col_field_indices.contains(&i) {
                return Err(format!(
                    "field at index {i} appears on multiple axes (row + col)"
                ));
            }
        }
        for df in &self.data_fields {
            if df.field_index >= nfields {
                return Err(format!(
                    "data field references field index {} out of range",
                    df.field_index
                ));
            }
        }
        Ok(())
    }
}

/// RFC-061 §10.6 — pivot-area selector.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotArea {
    pub field: Option<u32>,
    pub area_type: String,
    pub data_only: bool,
    pub label_only: bool,
    pub grand_row: bool,
    pub grand_col: bool,
    pub cache_index: Option<u32>,
    pub axis: Option<String>,
    pub field_position: Option<u32>,
}

/// RFC-061 §10.7 — pivot Format directive.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Format {
    pub action: String, // "formatting" | "blank"
    pub dxf_id: i32,
    pub pivot_area: PivotArea,
}

/// RFC-061 §2.5 — pivot-scoped CF rule wrapper.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotConditionalFormat {
    pub priority: i32,
    pub scope: String,   // "selection" | "data" | "field"
    pub cf_type: String, // "all" | "row" | "column" | "none"
    pub pivot_areas: Vec<PivotArea>,
    /// Reference to a workbook-scoped dxf entry id (-1 = unallocated).
    /// The patcher's RFC-026 dxf allocator stamps this at flush time.
    pub dxf_id: i32,
}

/// `<c:pivotSource>` block on a chart. RFC-049 §10.1. Lives here
/// because `wolfxl-charts` cannot depend on `wolfxl-pivot` (would be
/// a cycle); the chart crate gets a fully-resolved string `name` +
/// `fmt_id` from the Python coordinator's `Chart.to_rust_dict()`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotSource {
    /// `"MyPivot"` or `"Sheet1!MyPivot"`.
    pub name: String,
    pub fmt_id: u32,
}

#[cfg(test)]
mod tests {
    use super::*;

    fn dummy_table(n_fields: usize) -> PivotTable {
        PivotTable {
            name: "PivotTable1".into(),
            cache_id: 0,
            location: Location {
                range: "F2:I20".into(),
                first_header_row: 0,
                first_data_row: 1,
                first_data_col: 1,
                row_page_count: None,
                col_page_count: None,
            },
            pivot_fields: (0..n_fields).map(|_| PivotField::default()).collect(),
            row_field_indices: vec![],
            col_field_indices: vec![],
            page_fields: vec![],
            data_fields: vec![DataField {
                name: "Sum".into(),
                field_index: 0,
                function: DataFunction::Sum,
                show_data_as: None,
                base_field: 0,
                base_item: 0,
                num_fmt_id: None,
            }],
            row_items: vec![],
            col_items: vec![],
            data_on_rows: false,
            outline: true,
            compact: true,
            row_grand_totals: true,
            col_grand_totals: true,
            data_caption: "Values".into(),
            grand_total_caption: None,
            error_caption: None,
            missing_caption: None,
            apply_number_formats: false,
            apply_border_formats: false,
            apply_font_formats: false,
            apply_pattern_formats: false,
            apply_alignment_formats: false,
            apply_width_height_formats: true,
            style_info: Some(PivotTableStyleInfo::default()),
            created_version: 6,
            updated_version: 6,
            min_refreshable_version: 3,
            formats: vec![],
            conditional_formats: vec![],
        }
    }

    #[test]
    fn validate_no_data_fields_rejects() {
        let mut pt = dummy_table(2);
        pt.data_fields.clear();
        assert!(pt.validate().is_err());
    }

    #[test]
    fn validate_row_oor_rejects() {
        let mut pt = dummy_table(2);
        pt.row_field_indices = vec![5];
        assert!(pt.validate().is_err());
    }

    #[test]
    fn validate_dual_axis_rejects() {
        let mut pt = dummy_table(2);
        pt.row_field_indices = vec![0];
        pt.col_field_indices = vec![0];
        assert!(pt.validate().is_err());
    }

    #[test]
    fn validate_data_field_oor_rejects() {
        let mut pt = dummy_table(2);
        pt.data_fields[0].field_index = 5;
        assert!(pt.validate().is_err());
    }

    #[test]
    fn validate_happy() {
        let mut pt = dummy_table(2);
        pt.row_field_indices = vec![0];
        pt.col_field_indices = vec![1];
        assert!(pt.validate().is_ok());
    }
}
