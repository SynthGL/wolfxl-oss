//! Native workbook readers for WolfXL.
//!
//! This crate is the dependency-free-from-calamine reader foundation. The
//! first production target is XLSX/XLSM because those files already use ZIP +
//! OOXML helpers elsewhere in WolfXL. XLSB and XLS readers will grow beside
//! this API while preserving the same value-only public contract they have
//! today.

mod merge_metadata;
mod xlsb;

pub mod external_links;

pub use xlsb::NativeXlsbBook;

use std::collections::{HashMap, HashSet};
use std::fs::File;
use std::io::{Cursor, Read, Seek};
use std::path::{Path, PathBuf};

use quick_xml::events::attributes::Attribute;
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader as XmlReader;
use wolfxl_formula::{translate as translate_formula, RefDelta};
use wolfxl_rels::{RelId, RelsGraph};
use zip::ZipArchive;

const DEFAULT_MAX_ZIP_ENTRIES: usize = 200_000;
const DEFAULT_MAX_ZIP_ENTRY_BYTES: u64 = 512 * 1024 * 1024;
const DEFAULT_MAX_ZIP_TOTAL_BYTES: u64 = 4 * 1024 * 1024 * 1024;
const DEFAULT_MAX_COMPRESSION_RATIO: u64 = 1_000;
const MAX_VALUE_CELL_PREALLOC: usize = 1_000_000;

/// Native reader result type.
pub type Result<T> = std::result::Result<T, ReaderError>;

/// Read value-only cells from the first worksheet in an OOXML workbook on disk.
///
/// This is a one-shot fast path for Rust callers that only need worksheet
/// values. It opens the package once and skips styles, comments, drawings,
/// validations, document properties, and other metadata.
pub fn read_first_sheet_value_cells_path(path: impl AsRef<Path>) -> Result<Vec<ValueCell>> {
    let source = WorkbookSource::Path(path.as_ref().to_path_buf());
    read_first_sheet_value_cells_source(&source)
}

/// Read value-only cells from the first worksheet in OOXML workbook bytes.
///
/// This matches [`read_first_sheet_value_cells_path`] for callers that already
/// have the workbook in memory.
pub fn read_first_sheet_value_cells_bytes(bytes: impl Into<Vec<u8>>) -> Result<Vec<ValueCell>> {
    let source = WorkbookSource::Bytes(bytes.into());
    read_first_sheet_value_cells_source(&source)
}

/// Summarize numeric value cells from the first worksheet in an OOXML workbook.
///
/// This is the direct Rust fast path for value-scan workloads that only need a
/// row count and numeric checksum. It intentionally skips shared strings,
/// styles, comments, drawings, validations, and other metadata.
pub fn read_first_sheet_value_summary_path(path: impl AsRef<Path>) -> Result<ValueSheetSummary> {
    let source = WorkbookSource::Path(path.as_ref().to_path_buf());
    read_first_sheet_value_summary_source(&source)
}

/// Summarize formula text cells from the first worksheet in an OOXML workbook.
///
/// This is the direct Rust fast path for formula-audit workloads. It scans the
/// worksheet XML once and intentionally skips styles, comments, drawings,
/// validations, and other metadata.
pub fn read_first_sheet_formula_summary_path(
    path: impl AsRef<Path>,
) -> Result<FormulaSheetSummary> {
    let source = WorkbookSource::Path(path.as_ref().to_path_buf());
    read_first_sheet_formula_summary_source(&source)
}

/// Errors surfaced by native readers.
#[derive(Debug)]
pub enum ReaderError {
    Io(std::io::Error),
    Zip(zip::result::ZipError),
    Xml(String),
    MissingPart(String),
    SheetNotFound(String),
    Unsupported(String),
}

impl std::fmt::Display for ReaderError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            ReaderError::Io(e) => write!(f, "{e}"),
            ReaderError::Zip(e) => write!(f, "{e}"),
            ReaderError::Xml(e) => f.write_str(e),
            ReaderError::MissingPart(part) => write!(f, "missing workbook part: {part}"),
            ReaderError::SheetNotFound(sheet) => write!(f, "sheet not found: {sheet}"),
            ReaderError::Unsupported(msg) => f.write_str(msg),
        }
    }
}

impl std::error::Error for ReaderError {}

impl From<std::io::Error> for ReaderError {
    fn from(value: std::io::Error) -> Self {
        ReaderError::Io(value)
    }
}

impl From<zip::result::ZipError> for ReaderError {
    fn from(value: zip::result::ZipError) -> Self {
        ReaderError::Zip(value)
    }
}

/// Sheet metadata from `xl/workbook.xml` and `xl/_rels/workbook.xml.rels`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SheetInfo {
    pub name: String,
    pub sheet_id: Option<String>,
    pub state: SheetState,
    pub path: String,
}

/// Excel sheet visibility state.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SheetState {
    Visible,
    Hidden,
    VeryHidden,
}

impl Default for SheetState {
    fn default() -> Self {
        Self::Visible
    }
}

/// A decoded worksheet cell.
#[derive(Debug, Clone, PartialEq)]
pub struct Cell {
    /// 1-based row index.
    pub row: u32,
    /// 1-based column index.
    pub col: u32,
    /// A1 coordinate as written or inferred from row/column.
    pub coordinate: String,
    /// Raw `s` style id from the worksheet XML.
    pub style_id: Option<u32>,
    /// Cell type label from OOXML (`s`, `inlineStr`, `b`, `e`, `str`, or `n`).
    pub data_type: CellDataType,
    /// Cached/display value. Formula text lives in `formula`.
    pub value: CellValue,
    /// Formula text without a leading equals sign when present.
    pub formula: Option<String>,
    /// Raw OOXML formula kind (`array`, `dataTable`, `shared`, etc.).
    pub formula_kind: Option<String>,
    /// Raw OOXML shared-formula index (`si`) when present.
    pub formula_shared_index: Option<String>,
    /// Array/data-table formula metadata when this cell is the master.
    pub array_formula: Option<ArrayFormulaInfo>,
    /// Structured rich-text runs for shared-string or inline-string cells.
    pub rich_text: Option<Vec<RichTextRun>>,
}

/// A decoded value-only worksheet cell.
///
/// This is the fast path for workloads that only need cell values and do not
/// need styles, comments, drawings, validations, filters, or other worksheet
/// metadata.
#[derive(Debug, Clone, PartialEq)]
pub struct ValueCell {
    /// 1-based row index.
    pub row: u32,
    /// 1-based column index.
    pub col: u32,
    /// Cached/display value. Formula text is intentionally skipped.
    pub value: CellValue,
}

/// Compact value-scan summary for one worksheet.
#[derive(Debug, Clone, PartialEq)]
pub struct ValueSheetSummary {
    pub rows: usize,
    pub numeric_checksum: f64,
}

/// Compact formula-audit summary for one worksheet.
#[derive(Debug, Clone, PartialEq)]
pub struct FormulaSheetSummary {
    pub rows: usize,
    pub first_col_numeric_checksum: f64,
    pub formula_cells: usize,
}

/// Native cell value model shared by future readers.
#[derive(Debug, Clone, PartialEq)]
pub enum CellValue {
    Empty,
    String(String),
    Number(f64),
    Bool(bool),
    Error(String),
}

/// OOXML cell type classification.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CellDataType {
    Number,
    DateIso,
    SharedString,
    InlineString,
    FormulaString,
    Bool,
    Error,
}

/// Parsed worksheet data returned by the native XLSX reader.
#[derive(Debug, Clone, PartialEq)]
pub struct WorksheetData {
    pub dimension: Option<String>,
    pub merged_ranges: Vec<String>,
    pub hyperlinks: Vec<Hyperlink>,
    pub freeze_panes: Option<FreezePane>,
    pub sheet_properties: Option<SheetPropertiesInfo>,
    pub sheet_view: Option<SheetViewInfo>,
    pub comments: Vec<Comment>,
    pub threaded_comments: Vec<ParsedThreadedComment>,
    pub row_heights: HashMap<u32, RowHeight>,
    pub column_widths: Vec<ColumnWidth>,
    pub data_validations: Vec<DataValidation>,
    pub sheet_protection: Option<SheetProtection>,
    pub auto_filter: Option<AutoFilterInfo>,
    pub page_margins: Option<PageMarginsInfo>,
    pub page_setup: Option<PageSetupInfo>,
    pub print_options: Option<PrintOptionsInfo>,
    pub header_footer: Option<HeaderFooterInfo>,
    pub row_breaks: Option<PageBreakListInfo>,
    pub column_breaks: Option<PageBreakListInfo>,
    pub sheet_format: Option<SheetFormatInfo>,
    pub images: Vec<ImageInfo>,
    pub charts: Vec<ChartInfo>,
    pub tables: Vec<Table>,
    pub conditional_formats: Vec<ConditionalFormatRule>,
    pub hidden_rows: Vec<u32>,
    pub hidden_columns: Vec<u32>,
    pub row_outline_levels: Vec<(u32, u8)>,
    pub column_outline_levels: Vec<(u32, u8)>,
    pub array_formulas: HashMap<(u32, u32), ArrayFormulaInfo>,
    pub cells: Vec<Cell>,
}

/// Parsed worksheet hyperlink metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Hyperlink {
    pub cell: String,
    pub target: String,
    pub location: Option<String>,
    pub display: String,
    pub tooltip: Option<String>,
    pub internal: bool,
}

/// Parsed worksheet pane metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FreezePane {
    pub mode: PaneMode,
    pub top_left_cell: Option<String>,
    pub x_split: Option<i64>,
    pub y_split: Option<i64>,
    pub active_pane: Option<String>,
}

/// Parsed worksheet view metadata from the first `<sheetView>`.
#[derive(Debug, Clone, PartialEq)]
pub struct SheetViewInfo {
    pub zoom_scale: u32,
    pub zoom_scale_normal: u32,
    pub view: String,
    pub show_grid_lines: Option<bool>,
    pub show_row_col_headers: Option<bool>,
    pub show_outline_symbols: Option<bool>,
    pub show_zeros: Option<bool>,
    pub right_to_left: Option<bool>,
    pub tab_selected: Option<bool>,
    pub top_left_cell: Option<String>,
    pub workbook_view_id: u32,
    pub pane: Option<FreezePane>,
    pub selections: Vec<SelectionInfo>,
}

/// Parsed worksheet selection metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SelectionInfo {
    pub active_cell: Option<String>,
    pub sqref: Option<String>,
    pub pane: Option<String>,
    pub active_cell_id: Option<u32>,
}

/// OOXML pane mode relevant to read compatibility.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PaneMode {
    Freeze,
    Split,
}

/// Parsed worksheet properties from `<sheetPr>`.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct SheetPropertiesInfo {
    pub code_name: Option<String>,
    pub enable_format_conditions_calculation: Option<bool>,
    pub filter_mode: Option<bool>,
    pub published: Option<bool>,
    pub sync_horizontal: Option<bool>,
    pub sync_ref: Option<String>,
    pub sync_vertical: Option<bool>,
    pub transition_evaluation: Option<bool>,
    pub transition_entry: Option<bool>,
    pub tab_color: Option<String>,
    pub outline: OutlineInfo,
    pub page_setup: PageSetupPropertiesInfo,
}

/// Parsed worksheet outline display properties.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OutlineInfo {
    pub summary_below: bool,
    pub summary_right: bool,
    pub apply_styles: Option<bool>,
    pub show_outline_symbols: Option<bool>,
}

/// Parsed worksheet page-setup flags from `<sheetPr><pageSetUpPr>`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PageSetupPropertiesInfo {
    pub auto_page_breaks: Option<bool>,
    pub fit_to_page: Option<bool>,
}

impl Default for OutlineInfo {
    fn default() -> Self {
        Self {
            summary_below: true,
            summary_right: true,
            apply_styles: None,
            show_outline_symbols: None,
        }
    }
}

impl Default for PageSetupPropertiesInfo {
    fn default() -> Self {
        Self {
            auto_page_breaks: None,
            fit_to_page: None,
        }
    }
}

impl SheetPropertiesInfo {
    fn from_start(e: &BytesStart<'_>) -> Self {
        Self {
            code_name: attr_value(e, b"codeName"),
            enable_format_conditions_calculation: attr_bool(
                e,
                b"enableFormatConditionsCalculation",
            ),
            filter_mode: attr_bool(e, b"filterMode"),
            published: attr_bool(e, b"published"),
            sync_horizontal: attr_bool(e, b"syncHorizontal"),
            sync_ref: attr_value(e, b"syncRef"),
            sync_vertical: attr_bool(e, b"syncVertical"),
            transition_evaluation: attr_bool(e, b"transitionEvaluation"),
            transition_entry: attr_bool(e, b"transitionEntry"),
            ..Self::default()
        }
    }

    fn apply_tab_color(&mut self, e: &BytesStart<'_>) {
        self.tab_color = parse_ooxml_color(e);
    }

    fn apply_outline(&mut self, e: &BytesStart<'_>) {
        self.outline = OutlineInfo {
            summary_below: attr_bool_default(e, b"summaryBelow", true),
            summary_right: attr_bool_default(e, b"summaryRight", true),
            apply_styles: attr_bool(e, b"applyStyles"),
            show_outline_symbols: attr_bool(e, b"showOutlineSymbols"),
        };
    }

    fn apply_page_setup(&mut self, e: &BytesStart<'_>) {
        self.page_setup = PageSetupPropertiesInfo {
            auto_page_breaks: attr_bool(e, b"autoPageBreaks"),
            fit_to_page: attr_bool(e, b"fitToPage"),
        };
    }
}

impl SheetViewInfo {
    fn from_start(e: &BytesStart<'_>) -> Self {
        Self {
            zoom_scale: attr_u32(e, b"zoomScale").unwrap_or(100),
            zoom_scale_normal: attr_u32(e, b"zoomScaleNormal").unwrap_or(100),
            view: attr_value(e, b"view").unwrap_or_else(|| "normal".to_string()),
            show_grid_lines: attr_bool(e, b"showGridLines"),
            show_row_col_headers: attr_bool(e, b"showRowColHeaders"),
            show_outline_symbols: attr_bool(e, b"showOutlineSymbols"),
            show_zeros: attr_bool(e, b"showZeros"),
            right_to_left: attr_bool(e, b"rightToLeft"),
            tab_selected: attr_bool(e, b"tabSelected"),
            top_left_cell: attr_value(e, b"topLeftCell"),
            workbook_view_id: attr_u32(e, b"workbookViewId").unwrap_or_default(),
            pane: None,
            selections: Vec::new(),
        }
    }
}

impl SelectionInfo {
    fn from_start(e: &BytesStart<'_>) -> Self {
        Self {
            active_cell: attr_value(e, b"activeCell"),
            sqref: attr_value(e, b"sqref"),
            pane: attr_value(e, b"pane"),
            active_cell_id: attr_u32(e, b"activeCellId"),
        }
    }
}

/// Parsed worksheet comment.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Comment {
    pub cell: String,
    pub text: String,
    pub author: String,
    pub threaded: bool,
}

/// Workbook-scoped person record from `xl/persons/personList.xml` (RFC-068).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Person {
    pub display_name: String,
    pub id: String,
    pub user_id: Option<String>,
    pub provider_id: Option<String>,
}

/// One threaded-comment entry parsed from `xl/threadedComments/threadedCommentsN.xml`.
///
/// Top-level threads have ``parent_id == None``; replies carry the GUID of
/// their parent thread. Reassembly into a tree is done by callers (the
/// Python layer surfaces it as `cell.threaded_comment` with `replies`).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ParsedThreadedComment {
    pub id: String,
    pub cell: String,
    pub person_id: String,
    pub created: Option<String>,
    pub text: String,
    pub parent_id: Option<String>,
    pub done: bool,
}

/// Row dimension metadata from worksheet XML.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct RowHeight {
    pub height: f64,
    pub custom_height: bool,
}

/// Column dimension metadata from worksheet XML.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ColumnWidth {
    pub min: u32,
    pub max: u32,
    pub width: f64,
    pub custom_width: bool,
}

/// Parsed worksheet data-validation rule.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DataValidation {
    pub range: String,
    pub validation_type: String,
    pub operator: Option<String>,
    pub formula1: Option<String>,
    pub formula2: Option<String>,
    pub allow_blank: bool,
    pub show_input_message: bool,
    pub show_error_message: bool,
    pub error_title: Option<String>,
    pub error: Option<String>,
    pub prompt_title: Option<String>,
    pub prompt: Option<String>,
}

/// Parsed worksheet protection flags.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SheetProtection {
    pub sheet: bool,
    pub objects: bool,
    pub scenarios: bool,
    pub format_cells: bool,
    pub format_columns: bool,
    pub format_rows: bool,
    pub insert_columns: bool,
    pub insert_rows: bool,
    pub insert_hyperlinks: bool,
    pub delete_columns: bool,
    pub delete_rows: bool,
    pub select_locked_cells: bool,
    pub sort: bool,
    pub auto_filter: bool,
    pub pivot_tables: bool,
    pub select_unlocked_cells: bool,
    pub password_hash: Option<String>,
}

/// Parsed worksheet page margins.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct PageMarginsInfo {
    pub left: f64,
    pub right: f64,
    pub top: f64,
    pub bottom: f64,
    pub header: f64,
    pub footer: f64,
}

/// Parsed worksheet page setup.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PageSetupInfo {
    pub orientation: Option<String>,
    pub paper_size: Option<u32>,
    pub fit_to_width: Option<u32>,
    pub fit_to_height: Option<u32>,
    pub scale: Option<u32>,
    pub first_page_number: Option<u32>,
    pub horizontal_dpi: Option<u32>,
    pub vertical_dpi: Option<u32>,
    pub cell_comments: Option<String>,
    pub errors: Option<String>,
    pub use_first_page_number: Option<bool>,
    pub paper_height: Option<String>,
    pub paper_width: Option<String>,
    pub page_order: Option<String>,
    pub use_printer_defaults: Option<bool>,
    pub black_and_white: Option<bool>,
    pub draft: Option<bool>,
    pub copies: Option<u32>,
}

/// Parsed worksheet print options.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PrintOptionsInfo {
    pub horizontal_centered: Option<bool>,
    pub vertical_centered: Option<bool>,
    pub headings: Option<bool>,
    pub grid_lines: Option<bool>,
    pub grid_lines_set: Option<bool>,
}

/// Parsed worksheet header/footer settings.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct HeaderFooterInfo {
    pub odd_header: HeaderFooterItemInfo,
    pub odd_footer: HeaderFooterItemInfo,
    pub even_header: HeaderFooterItemInfo,
    pub even_footer: HeaderFooterItemInfo,
    pub first_header: HeaderFooterItemInfo,
    pub first_footer: HeaderFooterItemInfo,
    pub different_odd_even: bool,
    pub different_first: bool,
    pub scale_with_doc: bool,
    pub align_with_margins: bool,
}

/// Parsed left/center/right header/footer text.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct HeaderFooterItemInfo {
    pub left: Option<String>,
    pub center: Option<String>,
    pub right: Option<String>,
}

/// Parsed row or column page-break list.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PageBreakListInfo {
    pub count: u32,
    pub manual_break_count: u32,
    pub breaks: Vec<BreakInfo>,
}

/// Parsed single row/column page break.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BreakInfo {
    pub id: u32,
    pub min: Option<u32>,
    pub max: Option<u32>,
    pub man: bool,
    pub pt: bool,
}

/// Parsed worksheet sheet-format defaults.
#[derive(Debug, Clone, PartialEq)]
pub struct SheetFormatInfo {
    pub base_col_width: u32,
    pub default_col_width: Option<f64>,
    pub default_row_height: f64,
    pub custom_height: Option<bool>,
    pub zero_height: Option<bool>,
    pub thick_top: Option<bool>,
    pub thick_bottom: Option<bool>,
    pub outline_level_row: u32,
    pub outline_level_col: u32,
}

impl PageSetupInfo {
    fn from_start(e: &BytesStart<'_>) -> Self {
        Self {
            orientation: attr_value(e, b"orientation"),
            paper_size: attr_u32(e, b"paperSize"),
            fit_to_width: attr_u32(e, b"fitToWidth"),
            fit_to_height: attr_u32(e, b"fitToHeight"),
            scale: attr_u32(e, b"scale"),
            first_page_number: attr_u32(e, b"firstPageNumber"),
            horizontal_dpi: attr_u32(e, b"horizontalDpi"),
            vertical_dpi: attr_u32(e, b"verticalDpi"),
            cell_comments: attr_value(e, b"cellComments"),
            errors: attr_value(e, b"errors"),
            use_first_page_number: attr_bool(e, b"useFirstPageNumber"),
            paper_height: attr_value(e, b"paperHeight"),
            paper_width: attr_value(e, b"paperWidth"),
            page_order: attr_value(e, b"pageOrder"),
            use_printer_defaults: attr_bool(e, b"usePrinterDefaults"),
            black_and_white: attr_bool(e, b"blackAndWhite"),
            draft: attr_bool(e, b"draft"),
            copies: attr_u32(e, b"copies"),
        }
    }
}

impl PrintOptionsInfo {
    fn from_start(e: &BytesStart<'_>) -> Self {
        Self {
            horizontal_centered: attr_bool(e, b"horizontalCentered"),
            vertical_centered: attr_bool(e, b"verticalCentered"),
            headings: attr_bool(e, b"headings"),
            grid_lines: attr_bool(e, b"gridLines"),
            grid_lines_set: attr_bool(e, b"gridLinesSet"),
        }
    }
}

impl HeaderFooterInfo {
    fn from_start(e: &BytesStart<'_>) -> Self {
        Self {
            different_odd_even: attr_bool_default(e, b"differentOddEven", false),
            different_first: attr_bool_default(e, b"differentFirst", false),
            scale_with_doc: attr_bool_default(e, b"scaleWithDoc", true),
            align_with_margins: attr_bool_default(e, b"alignWithMargins", true),
            ..Self::default()
        }
    }

    fn set_part(&mut self, part: HeaderFooterPart, item: HeaderFooterItemInfo) {
        match part {
            HeaderFooterPart::OddHeader => self.odd_header = item,
            HeaderFooterPart::OddFooter => self.odd_footer = item,
            HeaderFooterPart::EvenHeader => self.even_header = item,
            HeaderFooterPart::EvenFooter => self.even_footer = item,
            HeaderFooterPart::FirstHeader => self.first_header = item,
            HeaderFooterPart::FirstFooter => self.first_footer = item,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum HeaderFooterPart {
    OddHeader,
    OddFooter,
    EvenHeader,
    EvenFooter,
    FirstHeader,
    FirstFooter,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum BreakListKind {
    Row,
    Column,
}

impl PageBreakListInfo {
    fn from_start(e: &BytesStart<'_>) -> Self {
        Self {
            count: attr_u32(e, b"count").unwrap_or_default(),
            manual_break_count: attr_u32(e, b"manualBreakCount").unwrap_or_default(),
            breaks: Vec::new(),
        }
    }
}

impl SheetFormatInfo {
    fn from_start(e: &BytesStart<'_>) -> Self {
        Self {
            base_col_width: attr_u32(e, b"baseColWidth").unwrap_or(8),
            default_col_width: attr_f64(e, b"defaultColWidth"),
            default_row_height: attr_f64(e, b"defaultRowHeight").unwrap_or(15.0),
            custom_height: attr_bool(e, b"customHeight"),
            zero_height: attr_bool(e, b"zeroHeight"),
            thick_top: attr_bool(e, b"thickTop"),
            thick_bottom: attr_bool(e, b"thickBottom"),
            outline_level_row: attr_u32(e, b"outlineLevelRow").unwrap_or_default(),
            outline_level_col: attr_u32(e, b"outlineLevelCol").unwrap_or_default(),
        }
    }
}

/// Parsed worksheet-level auto-filter metadata.
#[derive(Debug, Clone, PartialEq)]
pub struct AutoFilterInfo {
    pub ref_range: String,
    pub filter_columns: Vec<FilterColumnInfo>,
    pub sort_state: Option<SortStateInfo>,
}

/// Parsed worksheet-level auto-filter column.
#[derive(Debug, Clone, PartialEq)]
pub struct FilterColumnInfo {
    pub col_id: u32,
    pub hidden_button: bool,
    pub show_button: bool,
    pub filter: Option<FilterInfo>,
    pub date_group_items: Vec<DateGroupItemInfo>,
}

/// Parsed worksheet-level auto-filter predicate.
#[derive(Debug, Clone, PartialEq)]
pub enum FilterInfo {
    Blank,
    Color {
        dxf_id: u32,
        cell_color: bool,
    },
    Custom {
        and_: bool,
        filters: Vec<CustomFilterInfo>,
    },
    Dynamic {
        filter_type: String,
        val: Option<f64>,
        val_iso: Option<String>,
        max_val_iso: Option<String>,
    },
    Icon {
        icon_set: String,
        icon_id: u32,
    },
    String {
        values: Vec<String>,
    },
    Top10 {
        top: bool,
        percent: bool,
        val: f64,
        filter_val: Option<f64>,
    },
}

/// Parsed custom auto-filter predicate.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CustomFilterInfo {
    pub operator: String,
    pub val: String,
}

/// Parsed date-group auto-filter predicate.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DateGroupItemInfo {
    pub year: u32,
    pub month: Option<u32>,
    pub day: Option<u32>,
    pub hour: Option<u32>,
    pub minute: Option<u32>,
    pub second: Option<u32>,
    pub date_time_grouping: String,
}

/// Parsed worksheet-level auto-filter sort state.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SortStateInfo {
    pub sort_conditions: Vec<SortConditionInfo>,
    pub column_sort: bool,
    pub case_sensitive: bool,
    pub ref_range: Option<String>,
}

/// Parsed worksheet-level auto-filter sort condition.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SortConditionInfo {
    pub ref_range: String,
    pub descending: bool,
    pub sort_by: String,
    pub custom_list: Option<String>,
    pub dxf_id: Option<u32>,
    pub icon_set: Option<String>,
    pub icon_id: Option<u32>,
}

/// Parsed worksheet image payload and anchor metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ImageInfo {
    pub data: Vec<u8>,
    pub ext: String,
    pub anchor: ImageAnchorInfo,
}

/// Parsed worksheet chart payload and anchor metadata.
#[derive(Debug, Clone, PartialEq)]
pub struct ChartInfo {
    pub source_drawing_path: Option<String>,
    pub source_chart_rid: Option<String>,
    pub source_chart_path: Option<String>,
    pub kind: String,
    pub title: Option<String>,
    pub x_axis_title: Option<String>,
    pub y_axis_title: Option<String>,
    pub x_axis: Option<ChartAxisInfo>,
    pub y_axis: Option<ChartAxisInfo>,
    pub data_labels: Option<ChartDataLabelsInfo>,
    pub legend_position: Option<String>,
    pub bar_dir: Option<String>,
    pub grouping: Option<String>,
    pub scatter_style: Option<String>,
    pub vary_colors: Option<bool>,
    pub style: Option<u32>,
    pub anchor: ImageAnchorInfo,
    pub series: Vec<ChartSeriesInfo>,
}

/// Parsed chart data label options.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ChartDataLabelsInfo {
    pub position: Option<String>,
    pub show_legend_key: Option<bool>,
    pub show_val: Option<bool>,
    pub show_cat_name: Option<bool>,
    pub show_ser_name: Option<bool>,
    pub show_percent: Option<bool>,
    pub show_bubble_size: Option<bool>,
    pub show_leader_lines: Option<bool>,
}

/// Parsed chart axis metadata for native chart hydration.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct ChartAxisInfo {
    pub axis_type: String,
    pub axis_position: Option<String>,
    pub ax_id: Option<u32>,
    pub cross_ax: Option<u32>,
    pub scaling_min: Option<f64>,
    pub scaling_max: Option<f64>,
    pub scaling_orientation: Option<String>,
    pub scaling_log_base: Option<f64>,
    pub num_format_code: Option<String>,
    pub num_format_source_linked: Option<bool>,
    pub major_unit: Option<f64>,
    pub minor_unit: Option<f64>,
    pub tick_lbl_pos: Option<String>,
    pub major_tick_mark: Option<String>,
    pub minor_tick_mark: Option<String>,
    pub crosses: Option<String>,
    pub crosses_at: Option<f64>,
    pub cross_between: Option<String>,
    pub display_unit: Option<String>,
}

/// Parsed chart series references.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct ChartSeriesInfo {
    pub idx: Option<u32>,
    pub order: Option<u32>,
    pub title_ref: Option<String>,
    pub title_value: Option<String>,
    pub graphical_properties: Option<ChartGraphicalPropertiesInfo>,
    pub data_labels: Option<ChartDataLabelsInfo>,
    pub trendline: Option<ChartTrendlineInfo>,
    pub error_bars: Option<ChartErrorBarsInfo>,
    pub cat_ref: Option<String>,
    pub val_ref: Option<String>,
    pub x_ref: Option<String>,
    pub y_ref: Option<String>,
    pub bubble_size_ref: Option<String>,
}

/// Parsed chart graphical properties for a series.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ChartGraphicalPropertiesInfo {
    pub no_fill: Option<bool>,
    pub solid_fill: Option<String>,
    pub line_no_fill: Option<bool>,
    pub line_solid_fill: Option<String>,
    pub line_dash: Option<String>,
    pub line_width: Option<u32>,
}

/// Parsed chart trendline options for a series.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct ChartTrendlineInfo {
    pub trendline_type: Option<String>,
    pub order: Option<u32>,
    pub period: Option<u32>,
    pub forward: Option<f64>,
    pub backward: Option<f64>,
    pub intercept: Option<f64>,
    pub display_equation: Option<bool>,
    pub display_r_squared: Option<bool>,
}

/// Parsed chart error bar options for a series.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct ChartErrorBarsInfo {
    pub direction: Option<String>,
    pub bar_type: Option<String>,
    pub val_type: Option<String>,
    pub no_end_cap: Option<bool>,
    pub val: Option<f64>,
}

/// Parsed drawing anchor for an embedded image.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ImageAnchorInfo {
    OneCell {
        from: AnchorMarkerInfo,
        ext: Option<AnchorExtentInfo>,
    },
    TwoCell {
        from: AnchorMarkerInfo,
        to: AnchorMarkerInfo,
        edit_as: String,
    },
    Absolute {
        pos: AnchorPositionInfo,
        ext: AnchorExtentInfo,
    },
}

/// Parsed cell-relative drawing marker.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct AnchorMarkerInfo {
    pub col: i64,
    pub row: i64,
    pub col_off: i64,
    pub row_off: i64,
}

/// Parsed EMU drawing position.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct AnchorPositionInfo {
    pub x: i64,
    pub y: i64,
}

/// Parsed EMU drawing extent.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct AnchorExtentInfo {
    pub cx: i64,
    pub cy: i64,
}

/// Parsed worksheet table metadata.
#[derive(Debug, Clone, PartialEq)]
pub struct Table {
    pub name: String,
    pub ref_range: String,
    pub header_row: bool,
    pub totals_row: bool,
    pub comment: Option<String>,
    pub table_type: Option<String>,
    pub totals_row_shown: Option<bool>,
    pub style: Option<String>,
    pub show_first_column: bool,
    pub show_last_column: bool,
    pub show_row_stripes: bool,
    pub show_column_stripes: bool,
    pub columns: Vec<String>,
    /// The table's own `<autoFilter>` element, when present.
    pub auto_filter: Option<AutoFilterInfo>,
}

/// Parsed workbook defined name.
///
/// Phase 2 (G22) extends the struct to carry the full ECMA-376
/// `definedName` attribute surface that openpyxl exposes. `comment`
/// and `hidden` were previously parsed by the writer but never surfaced
/// from the reader; everything else is new in Phase 2.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct NamedRange {
    pub name: String,
    pub scope: String,
    /// Position of the owning sheet (`localSheetId`) for sheet-scoped names.
    pub local_sheet_id: Option<usize>,
    /// Formula text as stored in the workbook (no sheet prefix is added).
    pub refers_to: String,
    pub comment: Option<String>,
    pub hidden: bool,
    pub custom_menu: Option<String>,
    pub description: Option<String>,
    pub help: Option<String>,
    pub status_bar: Option<String>,
    pub shortcut_key: Option<String>,
    pub function: Option<bool>,
    pub function_group_id: Option<u32>,
    pub vb_procedure: Option<bool>,
    pub xlm: Option<bool>,
    pub publish_to_server: Option<bool>,
    pub workbook_parameter: Option<bool>,
}

/// Parsed worksheet print-title ranges.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PrintTitlesInfo {
    pub rows: Option<String>,
    pub cols: Option<String>,
}

/// Parsed workbook-level protection and file-sharing metadata.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct WorkbookSecurity {
    pub workbook_protection: Option<WorkbookProtection>,
    pub file_sharing: Option<FileSharing>,
}

/// Parsed `<workbookPr>` metadata.
#[derive(Debug, Clone, PartialEq)]
pub struct WorkbookPropertiesInfo {
    pub date1904: bool,
    pub date_compatibility: Option<bool>,
    pub show_objects: Option<String>,
    pub show_border_unselected_tables: Option<bool>,
    pub filter_privacy: Option<bool>,
    pub prompted_solutions: Option<bool>,
    pub show_ink_annotation: Option<bool>,
    pub backup_file: Option<bool>,
    pub save_external_link_values: Option<bool>,
    pub update_links: Option<String>,
    pub code_name: Option<String>,
    pub hide_pivot_field_list: Option<bool>,
    pub show_pivot_chart_filter: Option<bool>,
    pub allow_refresh_query: Option<bool>,
    pub publish_items: Option<bool>,
    pub check_compatibility: Option<bool>,
    pub auto_compress_pictures: Option<bool>,
    pub refresh_all_connections: Option<bool>,
    pub default_theme_version: Option<u32>,
}

/// Parsed `<calcPr>` metadata.
#[derive(Debug, Clone, PartialEq)]
pub struct CalcPropertiesInfo {
    pub calc_id: Option<u32>,
    pub calc_mode: Option<String>,
    pub full_calc_on_load: Option<bool>,
    pub ref_mode: Option<String>,
    pub iterate: Option<bool>,
    pub iterate_count: Option<u32>,
    pub iterate_delta: Option<f64>,
    pub full_precision: Option<bool>,
    pub calc_completed: Option<bool>,
    pub calc_on_save: Option<bool>,
    pub concurrent_calc: Option<bool>,
    pub concurrent_manual_count: Option<u32>,
    pub force_full_calc: Option<bool>,
}

/// Parsed workbook window view metadata from `<workbookView>`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BookViewInfo {
    pub visibility: String,
    pub minimized: bool,
    pub show_horizontal_scroll: bool,
    pub show_vertical_scroll: bool,
    pub show_sheet_tabs: bool,
    pub x_window: Option<i64>,
    pub y_window: Option<i64>,
    pub window_width: Option<u32>,
    pub window_height: Option<u32>,
    pub tab_ratio: u32,
    pub first_sheet: u32,
    pub active_tab: u32,
    pub auto_filter_date_grouping: bool,
}

/// Parsed custom document property from `docProps/custom.xml`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CustomPropertyInfo {
    pub name: String,
    pub kind: String,
    pub value: String,
}

/// Parsed `<workbookProtection>` metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WorkbookProtection {
    pub lock_structure: bool,
    pub lock_windows: bool,
    pub lock_revision: bool,
    pub workbook_password: Option<String>,
    pub workbook_password_character_set: Option<String>,
    pub workbook_algorithm_name: Option<String>,
    pub workbook_hash_value: Option<String>,
    pub workbook_salt_value: Option<String>,
    pub workbook_spin_count: Option<u32>,
    pub revisions_password: Option<String>,
    pub revisions_password_character_set: Option<String>,
    pub revisions_algorithm_name: Option<String>,
    pub revisions_hash_value: Option<String>,
    pub revisions_salt_value: Option<String>,
    pub revisions_spin_count: Option<u32>,
}

/// Parsed `<fileSharing>` metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FileSharing {
    pub read_only_recommended: bool,
    pub user_name: Option<String>,
    pub reservation_password: Option<String>,
    pub algorithm_name: Option<String>,
    pub hash_value: Option<String>,
    pub salt_value: Option<String>,
    pub spin_count: Option<u32>,
}

/// Parsed worksheet conditional-formatting rule.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ConditionalFormatRule {
    pub range: String,
    pub rule_type: String,
    pub operator: Option<String>,
    pub formula: Option<String>,
    pub dxf_id: Option<u32>,
    pub priority: Option<i64>,
    pub stop_if_true: Option<bool>,
    /// Populated for ``colorScale`` rules so callers can reconstruct the
    /// gradient anchors (cfvo type / val) and per-stop color. Empty for
    /// every other rule kind.
    pub color_scale: Option<ColorScaleInfo>,
}

/// Parsed `<colorScale>` block: ordered cfvo anchors paired with colors.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ColorScaleInfo {
    pub cfvo: Vec<CfvoInfo>,
    pub colors: Vec<String>,
}

/// One `<cfvo>` anchor: type plus optional literal value (for num / percent
/// / percentile / formula thresholds).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CfvoInfo {
    pub cfvo_type: String,
    pub val: Option<String>,
}

/// Array/data-table formula metadata for a cell.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ArrayFormulaInfo {
    Array {
        ref_range: String,
        text: String,
    },
    DataTable {
        ref_range: String,
        ca: bool,
        dt2_d: bool,
        dtr: bool,
        r1: Option<String>,
        r2: Option<String>,
        del1: bool,
        del2: bool,
    },
    SpillChild,
}

/// Inline font properties attached to a rich-text run.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct InlineFontProps {
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub strike: Option<bool>,
    pub underline: Option<String>,
    pub size: Option<f64>,
    pub color: Option<String>,
    pub name: Option<String>,
    pub family: Option<i32>,
    pub charset: Option<i32>,
    pub vert_align: Option<String>,
    pub scheme: Option<String>,
}

/// One structured rich-text run from `<si>` or inline `<is>` content.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct RichTextRun {
    pub text: String,
    pub font: Option<InlineFontProps>,
}

/// Native XLSX/XLSM workbook reader.
#[derive(Debug, Clone)]
pub struct NativeXlsxBook {
    source: WorkbookSource,
    sheets: Vec<SheetInfo>,
    named_ranges: Vec<NamedRange>,
    print_areas: HashMap<String, String>,
    print_titles: HashMap<String, PrintTitlesInfo>,
    workbook_security: WorkbookSecurity,
    workbook_properties: Option<WorkbookPropertiesInfo>,
    calc_properties: Option<CalcPropertiesInfo>,
    workbook_views: Vec<BookViewInfo>,
    custom_doc_properties: Vec<CustomPropertyInfo>,
    doc_properties: HashMap<String, String>,
    shared_strings: SharedStrings,
    styles: StyleTables,
    date1904: bool,
    /// Workbook-scoped person registry parsed from `xl/persons/personList.xml`.
    /// Empty when the workbook has no threaded-comment authorship metadata.
    persons: Vec<Person>,
}

#[derive(Debug, Clone)]
enum WorkbookSource {
    Path(PathBuf),
    Bytes(Vec<u8>),
}

enum ZipSource<'a> {
    Path(File),
    Bytes(Cursor<&'a [u8]>),
}

impl Read for ZipSource<'_> {
    fn read(&mut self, buf: &mut [u8]) -> std::io::Result<usize> {
        match self {
            ZipSource::Path(file) => file.read(buf),
            ZipSource::Bytes(cursor) => cursor.read(buf),
        }
    }
}

impl Seek for ZipSource<'_> {
    fn seek(&mut self, pos: std::io::SeekFrom) -> std::io::Result<u64> {
        match self {
            ZipSource::Path(file) => file.seek(pos),
            ZipSource::Bytes(cursor) => cursor.seek(pos),
        }
    }
}

impl NativeXlsxBook {
    /// Open an OOXML workbook from disk.
    pub fn open_path(path: impl AsRef<Path>) -> Result<Self> {
        Self::open_path_permissive(path, false)
    }

    /// Open an OOXML workbook from disk, optionally enabling malformed-topology
    /// recovery for legacy-compatible permissive loads.
    pub fn open_path_permissive(path: impl AsRef<Path>, permissive: bool) -> Result<Self> {
        let source = WorkbookSource::Path(path.as_ref().to_path_buf());
        Self::open_source_permissive(source, permissive)
    }

    /// Open an OOXML workbook from bytes.
    pub fn open_bytes(bytes: impl Into<Vec<u8>>) -> Result<Self> {
        Self::open_bytes_permissive(bytes, false)
    }

    /// Open an OOXML workbook from bytes, optionally enabling
    /// malformed-topology recovery for legacy-compatible permissive loads.
    pub fn open_bytes_permissive(bytes: impl Into<Vec<u8>>, permissive: bool) -> Result<Self> {
        Self::open_source_permissive(WorkbookSource::Bytes(bytes.into()), permissive)
    }

    fn open_source_permissive(source: WorkbookSource, permissive: bool) -> Result<Self> {
        let mut zip = zip_from_source(&source)?;
        validate_zip_archive(&mut zip)?;
        let workbook_xml = read_part_required(&mut zip, "xl/workbook.xml")?;
        let workbook_rels = read_part_required(&mut zip, "xl/_rels/workbook.xml.rels")?;
        let rels = RelsGraph::parse(workbook_rels.as_bytes())
            .map_err(|e| ReaderError::Xml(format!("failed to parse workbook rels: {e}")))?;
        let (
            sheet_refs,
            date1904,
            named_ranges,
            print_areas,
            print_titles,
            workbook_security,
            workbook_properties,
            calc_properties,
            workbook_views,
        ) = parse_workbook(&workbook_xml)?;
        let sheet_refs = if permissive && sheet_refs.is_empty() {
            synthesize_sheet_refs_from_rels(&rels)
        } else {
            sheet_refs
        };
        let sheets = resolve_sheet_paths(sheet_refs, &rels)?;
        let shared_strings = match read_part_optional(&mut zip, "xl/sharedStrings.xml")? {
            Some(xml) => parse_shared_strings(&xml)?,
            None => SharedStrings::default(),
        };
        let styles = match read_part_optional(&mut zip, "xl/styles.xml")? {
            Some(xml) => parse_style_tables(&xml)?,
            None => StyleTables::default(),
        };
        let mut doc_properties = HashMap::new();
        if let Some(xml) = read_part_optional(&mut zip, "docProps/core.xml")? {
            parse_doc_properties_into(&xml, &mut doc_properties, doc_property_core_key);
        }
        if let Some(xml) = read_part_optional(&mut zip, "docProps/app.xml")? {
            parse_doc_properties_into(&xml, &mut doc_properties, doc_property_app_key);
        }
        let custom_doc_properties = match read_part_optional(&mut zip, "docProps/custom.xml")? {
            Some(xml) => parse_custom_doc_properties(&xml)?,
            None => Vec::new(),
        };

        // RFC-068 / G08: workbook-scoped persons registry. Resolved via the
        // workbook rels graph so we honor whatever path the writer chose
        // (canonical wolfxl writes `xl/persons/personList.xml`).
        let persons = match person_list_target(&rels) {
            Some(target) => {
                let path = join_and_normalize("xl/", &target);
                match read_part_optional(&mut zip, &path)? {
                    Some(xml) => parse_person_list(&xml)?,
                    None => Vec::new(),
                }
            }
            None => Vec::new(),
        };

        Ok(Self {
            source,
            sheets,
            named_ranges,
            print_areas,
            print_titles,
            workbook_security,
            workbook_properties,
            calc_properties,
            workbook_views,
            custom_doc_properties,
            doc_properties,
            shared_strings,
            styles,
            date1904,
            persons,
        })
    }

    /// Workbook-scoped persons registry for threaded comments (RFC-068).
    ///
    /// Returns an empty slice when the workbook has no threaded-comment
    /// authorship metadata.
    pub fn persons(&self) -> &[Person] {
        &self.persons
    }

    /// Workbook sheet names in document order.
    pub fn sheet_names(&self) -> Vec<&str> {
        self.sheets.iter().map(|s| s.name.as_str()).collect()
    }

    /// Workbook sheet metadata in document order.
    pub fn sheets(&self) -> &[SheetInfo] {
        &self.sheets
    }

    /// Resolve a worksheet's workbook-tab visibility state.
    pub fn sheet_state(&self, sheet_name: &str) -> Result<SheetState> {
        self.sheets
            .iter()
            .find(|s| s.name == sheet_name)
            .map(|sheet| sheet.state)
            .ok_or_else(|| ReaderError::SheetNotFound(sheet_name.to_string()))
    }

    /// Workbook defined names.
    pub fn named_ranges(&self) -> &[NamedRange] {
        &self.named_ranges
    }

    /// Worksheet print area parsed from `_xlnm.Print_Area`, if present.
    pub fn print_area(&self, sheet_name: &str) -> Option<&str> {
        self.print_areas.get(sheet_name).map(String::as_str)
    }

    /// Worksheet print titles parsed from `_xlnm.Print_Titles`, if present.
    pub fn print_titles(&self, sheet_name: &str) -> Option<&PrintTitlesInfo> {
        self.print_titles.get(sheet_name)
    }

    /// Workbook protection and file-sharing blocks.
    pub fn workbook_security(&self) -> &WorkbookSecurity {
        &self.workbook_security
    }

    /// Workbook-level properties parsed from `<workbookPr>`.
    pub fn workbook_properties(&self) -> Option<&WorkbookPropertiesInfo> {
        self.workbook_properties.as_ref()
    }

    /// Workbook calculation properties parsed from `<calcPr>`.
    pub fn calc_properties(&self) -> Option<&CalcPropertiesInfo> {
        self.calc_properties.as_ref()
    }

    /// Workbook window views parsed from `<bookViews>`.
    pub fn workbook_views(&self) -> &[BookViewInfo] {
        &self.workbook_views
    }

    /// Custom document properties parsed from `docProps/custom.xml`.
    pub fn custom_doc_properties(&self) -> &[CustomPropertyInfo] {
        &self.custom_doc_properties
    }

    /// Workbook document properties parsed from `docProps/core.xml` and app.xml.
    pub fn doc_properties(&self) -> &HashMap<String, String> {
        &self.doc_properties
    }

    /// Whether the workbook uses the 1904 date system.
    pub fn date1904(&self) -> bool {
        self.date1904
    }

    /// Shared-string table as plain strings.
    pub fn shared_strings(&self) -> &[String] {
        &self.shared_strings.values
    }

    /// Resolve a style id to an Excel number format code.
    pub fn number_format_for_style_id(&self, style_id: u32) -> Option<&str> {
        self.styles.number_format_for_style_id(style_id)
    }

    /// Resolve a style id to cell border metadata.
    pub fn border_for_style_id(&self, style_id: u32) -> Option<&BorderInfo> {
        self.styles.border_for_style_id(style_id)
    }

    /// Resolve a style id to font metadata.
    pub fn font_for_style_id(&self, style_id: u32) -> Option<&FontInfo> {
        self.styles.font_for_style_id(style_id)
    }

    /// Resolve a style id to fill metadata.
    pub fn fill_for_style_id(&self, style_id: u32) -> Option<&FillInfo> {
        self.styles.fill_for_style_id(style_id)
    }

    /// Resolve a style id to alignment metadata.
    pub fn alignment_for_style_id(&self, style_id: u32) -> Option<&AlignmentInfo> {
        self.styles.alignment_for_style_id(style_id)
    }

    /// Resolve a style id to cell protection metadata. Returns ``None`` when
    /// the xf has no `<protection>` child (meaning "use Excel defaults":
    /// ``locked=true``, ``hidden=false``).
    pub fn protection_for_style_id(&self, style_id: u32) -> Option<&ProtectionInfo> {
        self.styles.protection_for_style_id(style_id)
    }

    /// Resolve a style id to a workbook-level named cell style.
    ///
    /// Returns `None` when the cell carries the implicit Normal style or
    /// the file has no `<cellStyles>` entry for the referenced xfId.
    /// What openpyxl exposes as `cell.style`.
    pub fn named_style_for_style_id(&self, style_id: u32) -> Option<&str> {
        self.styles.named_style_for_style_id(style_id)
    }

    /// Read merged ranges without decoding worksheet cells or related parts.
    pub fn worksheet_merged_ranges(&self, sheet_name: &str) -> Result<Vec<String>> {
        merge_metadata::read_ranges(self, sheet_name)
    }

    /// Read only the styles of merged anchors and bottom-right endpoints.
    pub fn worksheet_merge_endpoint_styles(
        &self,
        sheet_name: &str,
        ranges: &[String],
    ) -> Result<HashMap<(u32, u32), u32>> {
        merge_metadata::read_endpoint_styles(self, sheet_name, ranges)
    }

    /// Parse a worksheet into sparse decoded cells.
    pub fn worksheet(&self, sheet_name: &str) -> Result<WorksheetData> {
        let Some(info) = self.sheets.iter().find(|s| s.name == sheet_name) else {
            return Err(ReaderError::SheetNotFound(sheet_name.to_string()));
        };
        let mut zip = zip_from_source(&self.source)?;
        let xml = read_part_required(&mut zip, &info.path)?;
        let rels = read_part_optional(&mut zip, &sheet_rels_path(&info.path))?
            .map(|xml| {
                RelsGraph::parse(xml.as_bytes())
                    .map_err(|e| ReaderError::Xml(format!("failed to parse sheet rels: {e}")))
            })
            .transpose()?;
        let comments = match rels.as_ref().and_then(comments_target) {
            Some(target) => read_part_optional(
                &mut zip,
                &join_and_normalize(&part_dir(&info.path), &target),
            )?
            .map(|xml| parse_comments(&xml))
            .transpose()?
            .unwrap_or_default(),
            None => Vec::new(),
        };
        let threaded_comments = match rels.as_ref() {
            Some(graph) => {
                let mut acc: Vec<ParsedThreadedComment> = Vec::new();
                for rel in graph
                    .iter()
                    .filter(|rel| rel.rel_type == wolfxl_rels::rt::THREADED_COMMENTS)
                {
                    let path = join_and_normalize(&part_dir(&info.path), &rel.target);
                    if let Some(xml) = read_part_optional(&mut zip, &path)? {
                        acc.extend(parse_threaded_comments(&xml)?);
                    }
                }
                acc
            }
            None => Vec::new(),
        };
        let tables = read_tables(&mut zip, &info.path, &xml, rels.as_ref())?;
        let images = read_images(&mut zip, &info.path, rels.as_ref())?;
        let charts = read_charts(&mut zip, &info.path, rels.as_ref())?;
        let mut data =
            parse_worksheet(&xml, &self.shared_strings, rels.as_ref(), comments, tables)?;
        data.images = images;
        data.charts = charts;
        data.threaded_comments = threaded_comments;
        Ok(data)
    }

    /// Parse only worksheet cell values.
    ///
    /// This intentionally skips worksheet metadata after `<sheetData>`, making
    /// it the direct Rust equivalent of a values-only read workload.
    pub fn worksheet_value_cells(&self, sheet_name: &str) -> Result<Vec<ValueCell>> {
        let Some(info) = self.sheets.iter().find(|s| s.name == sheet_name) else {
            return Err(ReaderError::SheetNotFound(sheet_name.to_string()));
        };
        let mut zip = zip_from_source(&self.source)?;
        let xml = read_part_required_bytes(&mut zip, &info.path)?;
        parse_worksheet_value_cells_bytes(&xml, &self.shared_strings)
    }

    /// Parse charts attached to a chartsheet tab.
    pub fn chartsheet_charts(&self, sheet_name: &str) -> Result<Vec<ChartInfo>> {
        let Some(info) = self
            .sheets
            .iter()
            .find(|s| s.name == sheet_name && s.path.contains("chartsheets/"))
        else {
            return Err(ReaderError::SheetNotFound(sheet_name.to_string()));
        };
        let mut zip = zip_from_source(&self.source)?;
        let rels = read_part_optional(&mut zip, &sheet_rels_path(&info.path))?
            .map(|xml| {
                RelsGraph::parse(xml.as_bytes())
                    .map_err(|e| ReaderError::Xml(format!("failed to parse chartsheet rels: {e}")))
            })
            .transpose()?;
        read_charts(&mut zip, &info.path, rels.as_ref())
    }
}

#[derive(Debug)]
struct SheetRef {
    name: String,
    sheet_id: Option<String>,
    state: SheetState,
    rid: String,
}

#[derive(Debug, Clone, Default, PartialEq)]
pub(crate) struct StyleTables {
    pub(crate) custom_num_fmts: HashMap<u32, String>,
    pub(crate) cell_xfs: Vec<XfEntry>,
    pub(crate) fonts: Vec<FontInfo>,
    pub(crate) fills: Vec<FillInfo>,
    pub(crate) borders: Vec<BorderInfo>,
    /// `<cellStyles>` map: cellStyleXfs slot id (the `xfId` attr) -> name.
    /// Populated only for explicitly-named styles. The "Normal" entry is
    /// stored as well so callers don't need to special-case slot 0.
    pub(crate) cell_styles: HashMap<u32, String>,
}

impl StyleTables {
    pub(crate) fn number_format_for_style_id(&self, style_id: u32) -> Option<&str> {
        if style_id == 0 {
            return None;
        }
        let xf = self.cell_xfs.get(style_id as usize)?;
        if xf.num_fmt_id == 0 {
            return None;
        }
        if let Some(custom) = self.custom_num_fmts.get(&xf.num_fmt_id) {
            return Some(custom.as_str());
        }
        match builtin_num_fmt(xf.num_fmt_id) {
            Some("General") => None,
            other => other,
        }
    }

    pub(crate) fn border_for_style_id(&self, style_id: u32) -> Option<&BorderInfo> {
        let xf = self.cell_xfs.get(style_id as usize)?;
        self.borders.get(xf.border_id as usize)
    }

    pub(crate) fn font_for_style_id(&self, style_id: u32) -> Option<&FontInfo> {
        let xf = self.cell_xfs.get(style_id as usize)?;
        self.fonts.get(xf.font_id as usize)
    }

    pub(crate) fn fill_for_style_id(&self, style_id: u32) -> Option<&FillInfo> {
        let xf = self.cell_xfs.get(style_id as usize)?;
        self.fills.get(xf.fill_id as usize)
    }

    pub(crate) fn alignment_for_style_id(&self, style_id: u32) -> Option<&AlignmentInfo> {
        self.cell_xfs.get(style_id as usize)?.alignment.as_ref()
    }

    pub(crate) fn protection_for_style_id(&self, style_id: u32) -> Option<&ProtectionInfo> {
        self.cell_xfs.get(style_id as usize)?.protection.as_ref()
    }

    /// Resolve a cell's `s` attribute to a workbook-level named style.
    ///
    /// Walks `cellXfs[style_id].xf_id` -> `cellStyles[xf_id]`. Returns
    /// `None` for the implicit "Normal" style (slot 0) so callers can
    /// keep their default-style logic; non-zero xfIds resolve to the
    /// explicit name registered in `<cellStyles>`.
    pub(crate) fn named_style_for_style_id(&self, style_id: u32) -> Option<&str> {
        let xf = self.cell_xfs.get(style_id as usize)?;
        if xf.xf_id == 0 {
            return None;
        }
        self.cell_styles.get(&xf.xf_id).map(String::as_str)
    }
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub(crate) struct XfEntry {
    pub(crate) num_fmt_id: u32,
    pub(crate) font_id: u32,
    pub(crate) fill_id: u32,
    pub(crate) border_id: u32,
    /// `xfId` attr on `<xf>`. Points into `<cellStyleXfs>` and is what
    /// `named_style_for_style_id` cross-references against
    /// `StyleTables::cell_styles`. `0` means the implicit Normal style.
    pub(crate) xf_id: u32,
    pub(crate) alignment: Option<AlignmentInfo>,
    pub(crate) protection: Option<ProtectionInfo>,
}

/// Parsed cell font.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct FontInfo {
    pub bold: bool,
    pub italic: bool,
    pub underline: Option<String>,
    pub strikethrough: bool,
    pub name: Option<String>,
    pub size: Option<f64>,
    pub color: Option<String>,
}

/// Parsed cell fill.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct FillInfo {
    pub bg_color: Option<String>,
    /// Set when this `<fill>` was a `<gradientFill>` rather than a
    /// `<patternFill>`. Mutually exclusive with `bg_color` per OOXML.
    pub gradient: Option<GradientInfo>,
}

/// Parsed `<gradientFill>` payload (G05 follow-up).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct GradientInfo {
    pub gradient_type: String,
    pub degree: String,
    pub left: String,
    pub right: String,
    pub top: String,
    pub bottom: String,
    pub stops: Vec<GradientStopInfo>,
}

/// One `<stop position="..."><color rgb="..."/></stop>` entry.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct GradientStopInfo {
    pub position: String,
    pub color: Option<String>,
}

/// Parsed cell alignment.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct AlignmentInfo {
    pub horizontal: Option<String>,
    pub vertical: Option<String>,
    pub wrap_text: bool,
    pub text_rotation: Option<u32>,
    pub indent: Option<u32>,
}

/// Parsed cell-level protection flags.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ProtectionInfo {
    pub locked: bool,
    pub hidden: bool,
}

impl Default for ProtectionInfo {
    fn default() -> Self {
        Self {
            locked: true,
            hidden: false,
        }
    }
}

/// Parsed cell border.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct BorderInfo {
    pub left: Option<BorderSide>,
    pub right: Option<BorderSide>,
    pub top: Option<BorderSide>,
    pub bottom: Option<BorderSide>,
    pub diagonal_up: Option<BorderSide>,
    pub diagonal_down: Option<BorderSide>,
}

/// Parsed border side.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BorderSide {
    pub style: String,
    pub color: String,
}

fn parse_workbook(
    xml: &str,
) -> Result<(
    Vec<SheetRef>,
    bool,
    Vec<NamedRange>,
    HashMap<String, String>,
    HashMap<String, PrintTitlesInfo>,
    WorkbookSecurity,
    Option<WorkbookPropertiesInfo>,
    Option<CalcPropertiesInfo>,
    Vec<BookViewInfo>,
)> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut sheets = Vec::new();
    let mut raw_names = Vec::new();
    let mut raw_print_areas = Vec::new();
    let mut raw_print_titles = Vec::new();
    let mut date1904 = false;
    let mut workbook_protection = None;
    let mut file_sharing = None;
    let mut workbook_properties = None;
    let mut calc_properties = None;
    let mut workbook_views = Vec::new();
    let mut in_defined_name = false;
    let mut current_name: Option<String> = None;
    let mut current_local_id: Option<usize> = None;
    let mut current_name_text = String::new();
    // G22 — accumulate the rest of the ECMA-376 `<definedName>` attrs.
    let mut current_dn_attrs: RawNamedRangeAttrs = RawNamedRangeAttrs::default();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.local_name().as_ref() {
                b"workbookPr" => {
                    date1904 = attr_truthy(attr_value(&e, b"date1904").as_deref());
                    workbook_properties = Some(WorkbookPropertiesInfo::from_start(&e));
                }
                b"calcPr" => {
                    calc_properties = Some(CalcPropertiesInfo::from_start(&e));
                }
                b"workbookView" => {
                    workbook_views.push(BookViewInfo::from_start(&e));
                }
                b"workbookProtection" => {
                    workbook_protection = Some(WorkbookProtection::from_start(&e));
                }
                b"fileSharing" => {
                    file_sharing = Some(FileSharing::from_start(&e));
                }
                b"sheet" => {
                    let name = attr_value(&e, b"name");
                    let rid = relationship_id_attr(&e);
                    if let (Some(name), Some(rid)) = (name, rid) {
                        sheets.push(SheetRef {
                            name,
                            sheet_id: attr_value(&e, b"sheetId"),
                            state: parse_sheet_state(attr_value(&e, b"state").as_deref()),
                            rid,
                        });
                    }
                }
                b"definedName" => {
                    in_defined_name = true;
                    current_name = attr_value(&e, b"name");
                    current_local_id =
                        attr_value(&e, b"localSheetId").and_then(|v| v.parse::<usize>().ok());
                    current_name_text.clear();
                    current_dn_attrs = RawNamedRangeAttrs {
                        comment: attr_value(&e, b"comment"),
                        hidden: attr_bool_default(&e, b"hidden", false),
                        custom_menu: attr_value(&e, b"customMenu"),
                        description: attr_value(&e, b"description"),
                        help: attr_value(&e, b"help"),
                        status_bar: attr_value(&e, b"statusBar"),
                        shortcut_key: attr_value(&e, b"shortcutKey"),
                        function: attr_bool(&e, b"function"),
                        function_group_id: attr_u32(&e, b"functionGroupId"),
                        vb_procedure: attr_bool(&e, b"vbProcedure"),
                        xlm: attr_bool(&e, b"xlm"),
                        publish_to_server: attr_bool(&e, b"publishToServer"),
                        workbook_parameter: attr_bool(&e, b"workbookParameter"),
                    };
                }
                _ => {}
            },
            Ok(Event::Text(e)) => {
                if in_defined_name {
                    current_name_text
                        .push_str(&e.unescape().map_err(|err| {
                            ReaderError::Xml(format!("defined name text: {err}"))
                        })?);
                }
            }
            Ok(Event::End(e)) => {
                if e.local_name().as_ref() == b"definedName" {
                    in_defined_name = false;
                    if let Some(name) = current_name.take() {
                        let refers_to = current_name_text.trim().to_string();
                        let attrs = std::mem::take(&mut current_dn_attrs);
                        if name == "_xlnm.Print_Area" && !refers_to.is_empty() {
                            raw_print_areas.push(RawNamedRange::from_attrs(
                                name,
                                current_local_id.take(),
                                refers_to,
                                attrs,
                            ));
                        } else if name == "_xlnm.Print_Titles" && !refers_to.is_empty() {
                            raw_print_titles.push(RawNamedRange::from_attrs(
                                name,
                                current_local_id.take(),
                                refers_to,
                                attrs,
                            ));
                        } else if !name.starts_with("_xlnm.") && !refers_to.is_empty() {
                            raw_names.push(RawNamedRange::from_attrs(
                                name,
                                current_local_id.take(),
                                refers_to,
                                attrs,
                            ));
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ReaderError::Xml(format!(
                    "failed to parse xl/workbook.xml: {e}"
                )));
            }
            _ => {}
        }
        buf.clear();
    }

    let named_ranges = resolve_named_ranges(raw_names);
    let print_areas = resolve_print_areas(&sheets, raw_print_areas);
    let print_titles = resolve_print_titles(&sheets, raw_print_titles);
    Ok((
        sheets,
        date1904,
        named_ranges,
        print_areas,
        print_titles,
        WorkbookSecurity {
            workbook_protection,
            file_sharing,
        },
        workbook_properties,
        calc_properties,
        workbook_views,
    ))
}

#[derive(Debug, Default)]
struct RawNamedRange {
    name: String,
    local_id: Option<usize>,
    refers_to: String,
    /// Phase 2 (G22) — extra ECMA-376 attrs preserved through resolve.
    /// Held as a sub-struct so xlsb.rs can also build raw entries with
    /// `RawNamedRange { name, local_id, refers_to, attrs: Default::default() }`
    /// without spelling out the dozen extra fields.
    attrs: RawNamedRangeAttrs,
}

impl RawNamedRange {
    fn from_attrs(
        name: String,
        local_id: Option<usize>,
        refers_to: String,
        attrs: RawNamedRangeAttrs,
    ) -> Self {
        Self {
            name,
            local_id,
            refers_to,
            attrs,
        }
    }
}

#[derive(Debug, Default)]
struct RawNamedRangeAttrs {
    comment: Option<String>,
    hidden: bool,
    custom_menu: Option<String>,
    description: Option<String>,
    help: Option<String>,
    status_bar: Option<String>,
    shortcut_key: Option<String>,
    function: Option<bool>,
    function_group_id: Option<u32>,
    vb_procedure: Option<bool>,
    xlm: Option<bool>,
    publish_to_server: Option<bool>,
    workbook_parameter: Option<bool>,
}

impl WorkbookProtection {
    fn from_start(e: &BytesStart<'_>) -> Self {
        Self {
            lock_structure: attr_bool_default(e, b"lockStructure", false),
            lock_windows: attr_bool_default(e, b"lockWindows", false),
            lock_revision: attr_bool_default(e, b"lockRevision", false),
            workbook_password: attr_value(e, b"workbookPassword"),
            workbook_password_character_set: attr_value(e, b"workbookPasswordCharacterSet"),
            workbook_algorithm_name: attr_value(e, b"workbookAlgorithmName"),
            workbook_hash_value: attr_value(e, b"workbookHashValue"),
            workbook_salt_value: attr_value(e, b"workbookSaltValue"),
            workbook_spin_count: attr_value(e, b"workbookSpinCount")
                .and_then(|value| value.parse::<u32>().ok()),
            revisions_password: attr_value(e, b"revisionsPassword"),
            revisions_password_character_set: attr_value(e, b"revisionsPasswordCharacterSet"),
            revisions_algorithm_name: attr_value(e, b"revisionsAlgorithmName"),
            revisions_hash_value: attr_value(e, b"revisionsHashValue"),
            revisions_salt_value: attr_value(e, b"revisionsSaltValue"),
            revisions_spin_count: attr_value(e, b"revisionsSpinCount")
                .and_then(|value| value.parse::<u32>().ok()),
        }
    }
}

impl WorkbookPropertiesInfo {
    fn from_start(e: &BytesStart<'_>) -> Self {
        Self {
            date1904: attr_bool_default(e, b"date1904", false),
            date_compatibility: attr_bool(e, b"dateCompatibility"),
            show_objects: attr_value(e, b"showObjects"),
            show_border_unselected_tables: attr_bool(e, b"showBorderUnselectedTables"),
            filter_privacy: attr_bool(e, b"filterPrivacy"),
            prompted_solutions: attr_bool(e, b"promptedSolutions"),
            show_ink_annotation: attr_bool(e, b"showInkAnnotation"),
            backup_file: attr_bool(e, b"backupFile"),
            save_external_link_values: attr_bool(e, b"saveExternalLinkValues"),
            update_links: attr_value(e, b"updateLinks"),
            code_name: attr_value(e, b"codeName"),
            hide_pivot_field_list: attr_bool(e, b"hidePivotFieldList"),
            show_pivot_chart_filter: attr_bool(e, b"showPivotChartFilter"),
            allow_refresh_query: attr_bool(e, b"allowRefreshQuery"),
            publish_items: attr_bool(e, b"publishItems"),
            check_compatibility: attr_bool(e, b"checkCompatibility"),
            auto_compress_pictures: attr_bool(e, b"autoCompressPictures"),
            refresh_all_connections: attr_bool(e, b"refreshAllConnections"),
            default_theme_version: attr_u32(e, b"defaultThemeVersion"),
        }
    }
}

impl CalcPropertiesInfo {
    fn from_start(e: &BytesStart<'_>) -> Self {
        Self {
            calc_id: attr_u32(e, b"calcId"),
            calc_mode: attr_value(e, b"calcMode"),
            full_calc_on_load: attr_bool(e, b"fullCalcOnLoad"),
            ref_mode: attr_value(e, b"refMode"),
            iterate: attr_bool(e, b"iterate"),
            iterate_count: attr_u32(e, b"iterateCount"),
            iterate_delta: attr_f64(e, b"iterateDelta"),
            full_precision: attr_bool(e, b"fullPrecision"),
            calc_completed: attr_bool(e, b"calcCompleted"),
            calc_on_save: attr_bool(e, b"calcOnSave"),
            concurrent_calc: attr_bool(e, b"concurrentCalc"),
            concurrent_manual_count: attr_u32(e, b"concurrentManualCount"),
            force_full_calc: attr_bool(e, b"forceFullCalc"),
        }
    }
}

impl BookViewInfo {
    fn from_start(e: &BytesStart<'_>) -> Self {
        Self {
            visibility: attr_value(e, b"visibility").unwrap_or_else(|| "visible".to_string()),
            minimized: attr_bool_default(e, b"minimized", false),
            show_horizontal_scroll: attr_bool_default(e, b"showHorizontalScroll", true),
            show_vertical_scroll: attr_bool_default(e, b"showVerticalScroll", true),
            show_sheet_tabs: attr_bool_default(e, b"showSheetTabs", true),
            x_window: attr_i64(e, b"xWindow"),
            y_window: attr_i64(e, b"yWindow"),
            window_width: attr_u32(e, b"windowWidth"),
            window_height: attr_u32(e, b"windowHeight"),
            tab_ratio: attr_u32(e, b"tabRatio").unwrap_or(600),
            first_sheet: attr_u32(e, b"firstSheet").unwrap_or_default(),
            active_tab: attr_u32(e, b"activeTab").unwrap_or_default(),
            auto_filter_date_grouping: attr_bool_default(e, b"autoFilterDateGrouping", true),
        }
    }
}

impl FileSharing {
    fn from_start(e: &BytesStart<'_>) -> Self {
        Self {
            read_only_recommended: attr_bool_default(e, b"readOnlyRecommended", false),
            user_name: attr_value(e, b"userName"),
            reservation_password: attr_value(e, b"reservationPassword"),
            algorithm_name: attr_value(e, b"algorithmName"),
            hash_value: attr_value(e, b"hashValue"),
            salt_value: attr_value(e, b"saltValue"),
            spin_count: attr_value(e, b"spinCount").and_then(|value| value.parse::<u32>().ok()),
        }
    }
}

#[derive(Debug, Clone, Default, PartialEq)]
struct SharedStrings {
    values: Vec<String>,
    rich_text: Vec<Option<Vec<RichTextRun>>>,
}

fn resolve_named_ranges(raw_names: Vec<RawNamedRange>) -> Vec<NamedRange> {
    raw_names
        .into_iter()
        .map(|raw| NamedRange {
            name: raw.name,
            scope: if raw.local_id.is_some() {
                "sheet"
            } else {
                "workbook"
            }
            .to_string(),
            local_sheet_id: raw.local_id,
            refers_to: raw.refers_to,
            comment: raw.attrs.comment,
            hidden: raw.attrs.hidden,
            custom_menu: raw.attrs.custom_menu,
            description: raw.attrs.description,
            help: raw.attrs.help,
            status_bar: raw.attrs.status_bar,
            shortcut_key: raw.attrs.shortcut_key,
            function: raw.attrs.function,
            function_group_id: raw.attrs.function_group_id,
            vb_procedure: raw.attrs.vb_procedure,
            xlm: raw.attrs.xlm,
            publish_to_server: raw.attrs.publish_to_server,
            workbook_parameter: raw.attrs.workbook_parameter,
        })
        .collect()
}

fn resolve_print_areas(
    sheet_refs: &[SheetRef],
    raw_print_areas: Vec<RawNamedRange>,
) -> HashMap<String, String> {
    let mut out = HashMap::new();
    for raw in raw_print_areas {
        let sheet_name = raw
            .local_id
            .and_then(|index| sheet_refs.get(index).map(|sheet| sheet.name.clone()))
            .or_else(|| sheet_name_from_ref(&raw.refers_to));
        let Some(sheet_name) = sheet_name else {
            continue;
        };
        let refers_to = if raw.refers_to.contains('!') {
            raw.refers_to
        } else {
            format!("{sheet_name}!{}", raw.refers_to)
        };
        out.insert(sheet_name, refers_to);
    }
    out
}

fn resolve_print_titles(
    sheet_refs: &[SheetRef],
    raw_print_titles: Vec<RawNamedRange>,
) -> HashMap<String, PrintTitlesInfo> {
    let mut out = HashMap::new();
    for raw in raw_print_titles {
        let sheet_name = raw
            .local_id
            .and_then(|index| sheet_refs.get(index).map(|sheet| sheet.name.clone()))
            .or_else(|| sheet_name_from_ref(&raw.refers_to));
        let Some(sheet_name) = sheet_name else {
            continue;
        };
        let titles = parse_print_titles_ref(&raw.refers_to);
        if titles.rows.is_some() || titles.cols.is_some() {
            out.insert(sheet_name, titles);
        }
    }
    out
}

fn parse_print_titles_ref(refers_to: &str) -> PrintTitlesInfo {
    let mut info = PrintTitlesInfo::default();
    for part in refers_to.split(',') {
        let range = part
            .split_once('!')
            .map(|(_, range)| range)
            .unwrap_or(part)
            .replace('$', "");
        let Some((start, end)) = range.split_once(':') else {
            continue;
        };
        if start.chars().all(|ch| ch.is_ascii_digit()) && end.chars().all(|ch| ch.is_ascii_digit())
        {
            info.rows = Some(format!("{start}:{end}"));
        } else if start.chars().all(|ch| ch.is_ascii_alphabetic())
            && end.chars().all(|ch| ch.is_ascii_alphabetic())
        {
            info.cols = Some(format!(
                "{}:{}",
                start.to_ascii_uppercase(),
                end.to_ascii_uppercase()
            ));
        }
    }
    info
}

fn sheet_name_from_ref(refers_to: &str) -> Option<String> {
    let (sheet, _) = refers_to.split_once('!')?;
    let unquoted = sheet
        .strip_prefix('\'')
        .and_then(|value| value.strip_suffix('\''))
        .unwrap_or(sheet);
    Some(unquoted.replace("''", "'"))
}

fn resolve_sheet_paths(sheet_refs: Vec<SheetRef>, rels: &RelsGraph) -> Result<Vec<SheetInfo>> {
    let mut sheets = Vec::with_capacity(sheet_refs.len());
    for sheet in sheet_refs {
        let rel = rels.get(&RelId(sheet.rid.clone())).ok_or_else(|| {
            ReaderError::MissingPart(format!("workbook relationship {}", sheet.rid))
        })?;
        sheets.push(SheetInfo {
            name: sheet.name,
            sheet_id: sheet.sheet_id,
            state: sheet.state,
            path: join_and_normalize("xl/", &rel.target),
        });
    }
    Ok(sheets)
}

fn synthesize_sheet_refs_from_rels(rels: &RelsGraph) -> Vec<SheetRef> {
    rels.find_by_type(wolfxl_rels::rt::WORKSHEET)
        .into_iter()
        .enumerate()
        .map(|(idx, rel)| SheetRef {
            name: format!("Sheet{}", idx + 1),
            sheet_id: Some((idx + 1).to_string()),
            state: SheetState::Visible,
            rid: rel.id.0.clone(),
        })
        .collect()
}

fn parse_shared_strings(xml: &str) -> Result<SharedStrings> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut out = SharedStrings::default();
    let mut current = String::new();
    let mut runs: Vec<RichTextRun> = Vec::new();
    let mut current_run: Option<RichTextRun> = None;
    let mut current_props: Option<InlineFontProps> = None;
    let mut in_si = false;
    let mut in_t = false;
    let mut saw_r = false;
    let mut text_buf = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.local_name().as_ref() {
                b"si" => {
                    in_si = true;
                    current.clear();
                    runs.clear();
                    current_run = None;
                    current_props = None;
                    saw_r = false;
                    text_buf.clear();
                }
                b"r" if in_si => {
                    saw_r = true;
                    current_run = Some(RichTextRun::default());
                    current_props = None;
                }
                b"rPr" if in_si && current_run.is_some() => {
                    current_props = Some(InlineFontProps::default());
                }
                b"t" => {
                    if in_si {
                        in_t = true;
                        text_buf.clear();
                    }
                }
                other if current_props.is_some() => {
                    apply_rich_font_attr(current_props.as_mut().unwrap(), other, e.attributes());
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.local_name().as_ref() {
                b"si" => {
                    out.values.push(String::new());
                    out.rich_text.push(None);
                }
                other if current_props.is_some() => {
                    apply_rich_font_attr(current_props.as_mut().unwrap(), other, e.attributes());
                }
                _ => {}
            },
            Ok(Event::End(e)) => match e.local_name().as_ref() {
                b"si" => {
                    out.values.push(std::mem::take(&mut current));
                    out.rich_text.push(saw_r.then(|| std::mem::take(&mut runs)));
                    in_si = false;
                }
                b"t" => {
                    current.push_str(&text_buf);
                    if let Some(run) = current_run.as_mut() {
                        run.text.push_str(&text_buf);
                    }
                    text_buf.clear();
                    in_t = false;
                }
                b"rPr" => {
                    if let (Some(run), Some(props)) = (current_run.as_mut(), current_props.take()) {
                        run.font = Some(props);
                    }
                }
                b"r" => {
                    if let Some(run) = current_run.take() {
                        runs.push(run);
                    }
                }
                _ => {}
            },
            Ok(Event::Text(e)) => {
                if in_si && in_t {
                    let text = e
                        .unescape()
                        .map_err(|err| ReaderError::Xml(format!("shared string text: {err}")))?;
                    text_buf.push_str(&normalize_ooxml_text(&text));
                }
            }
            Ok(Event::CData(e)) => {
                if in_si && in_t {
                    text_buf.push_str(&normalize_ooxml_text(&String::from_utf8_lossy(e.as_ref())));
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ReaderError::Xml(format!(
                    "failed to parse xl/sharedStrings.xml: {e}"
                )));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

fn apply_rich_font_attr(
    props: &mut InlineFontProps,
    tag: &[u8],
    mut attrs: quick_xml::events::attributes::Attributes<'_>,
) {
    let mut all_attrs: Vec<(Vec<u8>, String)> = Vec::with_capacity(2);
    for a in attrs.with_checks(false) {
        if let Ok(Attribute { key, value }) = a {
            all_attrs.push((
                key.local_name().as_ref().to_vec(),
                String::from_utf8_lossy(value.as_ref()).into_owned(),
            ));
        }
    }
    let val = || -> Option<String> {
        all_attrs
            .iter()
            .rev()
            .find(|(k, _)| k.as_slice() == b"val")
            .map(|(_, v)| v.clone())
    };
    match tag {
        b"b" => props.bold = Some(parse_rich_bool(val())),
        b"i" => props.italic = Some(parse_rich_bool(val())),
        b"strike" => props.strike = Some(parse_rich_bool(val())),
        b"u" => props.underline = Some(val().unwrap_or_else(|| "single".to_string())),
        b"sz" => props.size = val().and_then(|v| v.parse::<f64>().ok()),
        b"rFont" => props.name = val(),
        b"family" => props.family = val().and_then(|v| v.parse::<i32>().ok()),
        b"charset" => props.charset = val().and_then(|v| v.parse::<i32>().ok()),
        b"vertAlign" => props.vert_align = val(),
        b"scheme" => props.scheme = val(),
        b"color" => {
            let mut rgb = None;
            let mut theme = None;
            let mut indexed = None;
            for (k, v) in &all_attrs {
                match k.as_slice() {
                    b"rgb" => rgb = Some(v.clone()),
                    b"theme" => theme = Some(v.clone()),
                    b"indexed" => indexed = Some(v.clone()),
                    _ => {}
                }
            }
            props.color = rgb
                .or_else(|| theme.map(|v| format!("theme:{v}")))
                .or_else(|| indexed.map(|v| format!("indexed:{v}")));
        }
        _ => {}
    }
}

fn parse_rich_bool(value: Option<String>) -> bool {
    match value.as_deref() {
        None => true,
        Some(raw) => {
            let trimmed = raw.trim();
            !(trimmed == "0" || trimmed.eq_ignore_ascii_case("false"))
        }
    }
}

fn parse_style_tables(xml: &str) -> Result<StyleTables> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut styles = StyleTables::default();
    let mut in_num_fmts = false;
    let mut in_fonts = false;
    let mut in_fills = false;
    let mut in_cell_xfs = false;
    let mut in_borders = false;
    let mut current_font: Option<FontBuilder> = None;
    let mut current_fill: Option<FillBuilder> = None;
    let mut in_pattern_fill = false;
    let mut in_gradient_fill = false;
    let mut current_border: Option<BorderBuilder> = None;
    let mut current_border_edge: Option<BorderEdgeKind> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.local_name().as_ref() {
                b"numFmts" => in_num_fmts = true,
                b"fonts" => in_fonts = true,
                b"fills" => in_fills = true,
                b"cellXfs" => in_cell_xfs = true,
                b"borders" => in_borders = true,
                b"font" if in_fonts => {
                    current_font = Some(FontBuilder::default());
                }
                tag if in_fonts && current_font.is_some() => {
                    if let Some(font) = current_font.as_mut() {
                        font.apply_tag(tag, &e);
                    }
                }
                b"fill" if in_fills => {
                    current_fill = Some(FillBuilder::default());
                }
                b"patternFill" if in_fills && current_fill.is_some() => {
                    in_pattern_fill = true;
                    if let Some(fill) = current_fill.as_mut() {
                        fill.pattern_type = attr_value(&e, b"patternType");
                    }
                }
                b"gradientFill" if in_fills && current_fill.is_some() => {
                    in_gradient_fill = true;
                    if let Some(fill) = current_fill.as_mut() {
                        fill.start_gradient(&e);
                    }
                }
                b"stop" if in_gradient_fill && current_fill.is_some() => {
                    if let Some(fill) = current_fill.as_mut() {
                        fill.open_stop(&e);
                    }
                }
                b"color" if in_gradient_fill && current_fill.is_some() => {
                    if let Some(fill) = current_fill.as_mut() {
                        fill.apply_stop_color(&e);
                    }
                }
                tag if in_pattern_fill && current_fill.is_some() => {
                    if let Some(fill) = current_fill.as_mut() {
                        fill.apply_color_tag(tag, &e);
                    }
                }
                b"border" if in_borders => {
                    current_border = Some(BorderBuilder::from_start(&e));
                }
                tag if in_borders && current_border.is_some() => {
                    if let Some(edge) = BorderEdgeKind::from_tag(tag) {
                        current_border_edge = Some(edge);
                        if let Some(border) = current_border.as_mut() {
                            border.set_edge(edge, parse_border_side(&e));
                        }
                    } else if tag == b"color" {
                        if let (Some(border), Some(edge)) =
                            (current_border.as_mut(), current_border_edge)
                        {
                            border.update_edge_color(edge, parse_border_color(&e));
                        }
                    }
                }
                b"numFmt" if in_num_fmts => {
                    push_num_fmt(&mut styles, &e);
                }
                b"xf" if in_cell_xfs => {
                    styles.cell_xfs.push(parse_xf_entry(&e));
                }
                b"alignment" if in_cell_xfs => {
                    if let Some(xf) = styles.cell_xfs.last_mut() {
                        xf.alignment = parse_alignment(&e);
                    }
                }
                b"protection" if in_cell_xfs => {
                    if let Some(xf) = styles.cell_xfs.last_mut() {
                        xf.protection = parse_protection(&e);
                    }
                }
                b"cellStyle" => push_cell_style(&mut styles, &e),
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.local_name().as_ref() {
                b"numFmt" if in_num_fmts => push_num_fmt(&mut styles, &e),
                b"xf" if in_cell_xfs => styles.cell_xfs.push(parse_xf_entry(&e)),
                b"cellStyle" => push_cell_style(&mut styles, &e),
                b"font" if in_fonts => styles.fonts.push(FontInfo::default()),
                tag if in_fonts && current_font.is_some() => {
                    if let Some(font) = current_font.as_mut() {
                        font.apply_tag(tag, &e);
                    }
                }
                b"fill" if in_fills => styles.fills.push(FillInfo::default()),
                b"patternFill" if in_fills && current_fill.is_some() => {
                    if let Some(fill) = current_fill.as_mut() {
                        fill.pattern_type = attr_value(&e, b"patternType");
                    }
                }
                b"color" if in_gradient_fill && current_fill.is_some() => {
                    if let Some(fill) = current_fill.as_mut() {
                        fill.apply_stop_color(&e);
                    }
                }
                tag if in_pattern_fill && current_fill.is_some() => {
                    if let Some(fill) = current_fill.as_mut() {
                        fill.apply_color_tag(tag, &e);
                    }
                }
                b"alignment" if in_cell_xfs => {
                    if let Some(xf) = styles.cell_xfs.last_mut() {
                        xf.alignment = parse_alignment(&e);
                    }
                }
                b"protection" if in_cell_xfs => {
                    if let Some(xf) = styles.cell_xfs.last_mut() {
                        xf.protection = parse_protection(&e);
                    }
                }
                b"border" if in_borders => {
                    styles.borders.push(BorderBuilder::from_start(&e).finish());
                }
                tag if in_borders && current_border.is_some() => {
                    if let Some(edge) = BorderEdgeKind::from_tag(tag) {
                        if let Some(border) = current_border.as_mut() {
                            border.set_edge(edge, parse_border_side(&e));
                        }
                    } else if tag == b"color" {
                        if let (Some(border), Some(edge)) =
                            (current_border.as_mut(), current_border_edge)
                        {
                            border.update_edge_color(edge, parse_border_color(&e));
                        }
                    }
                }
                _ => {}
            },
            Ok(Event::End(e)) => match e.local_name().as_ref() {
                b"numFmts" => in_num_fmts = false,
                b"fonts" => in_fonts = false,
                b"fills" => in_fills = false,
                b"cellXfs" => in_cell_xfs = false,
                b"borders" => in_borders = false,
                b"font" => {
                    if let Some(font) = current_font.take() {
                        styles.fonts.push(font.finish());
                    }
                }
                b"fill" => {
                    if let Some(fill) = current_fill.take() {
                        styles.fills.push(fill.finish());
                    }
                    in_pattern_fill = false;
                    in_gradient_fill = false;
                }
                b"patternFill" => in_pattern_fill = false,
                b"gradientFill" => in_gradient_fill = false,
                b"stop" => {
                    if let Some(fill) = current_fill.as_mut() {
                        fill.close_stop();
                    }
                }
                b"border" => {
                    if let Some(border) = current_border.take() {
                        styles.borders.push(border.finish());
                    }
                    current_border_edge = None;
                }
                tag if BorderEdgeKind::from_tag(tag).is_some() => {
                    current_border_edge = None;
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ReaderError::Xml(format!(
                    "failed to parse xl/styles.xml: {e}"
                )));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(styles)
}

fn push_num_fmt(styles: &mut StyleTables, e: &BytesStart<'_>) {
    let id = attr_value(e, b"numFmtId").and_then(|value| value.parse::<u32>().ok());
    let code = attr_value(e, b"formatCode");
    if let (Some(id), Some(code)) = (id, code) {
        styles.custom_num_fmts.insert(id, code);
    }
}

fn push_cell_style(styles: &mut StyleTables, e: &BytesStart<'_>) {
    // `<cellStyle name="Foo" xfId="2" builtinId="..."/>` — the xfId is what
    // cells reference (via cellXfs[s].xf_id). Skip entries missing either
    // attribute; nothing else in OOXML can resolve them anyway.
    let xf_id = attr_value(e, b"xfId").and_then(|value| value.parse::<u32>().ok());
    let name = attr_value(e, b"name");
    if let (Some(xf_id), Some(name)) = (xf_id, name) {
        styles.cell_styles.insert(xf_id, name);
    }
}

fn parse_xf_entry(e: &BytesStart<'_>) -> XfEntry {
    XfEntry {
        num_fmt_id: attr_value(e, b"numFmtId")
            .and_then(|value| value.parse::<u32>().ok())
            .unwrap_or(0),
        font_id: attr_value(e, b"fontId")
            .and_then(|value| value.parse::<u32>().ok())
            .unwrap_or(0),
        fill_id: attr_value(e, b"fillId")
            .and_then(|value| value.parse::<u32>().ok())
            .unwrap_or(0),
        border_id: attr_value(e, b"borderId")
            .and_then(|value| value.parse::<u32>().ok())
            .unwrap_or(0),
        xf_id: attr_value(e, b"xfId")
            .and_then(|value| value.parse::<u32>().ok())
            .unwrap_or(0),
        alignment: None,
        protection: None,
    }
}

fn parse_protection(e: &BytesStart<'_>) -> Option<ProtectionInfo> {
    let locked_attr = attr_value(e, b"locked");
    let hidden_attr = attr_value(e, b"hidden");
    if locked_attr.is_none() && hidden_attr.is_none() {
        return None;
    }
    let parse_bool = |s: &str| !matches!(s, "0" | "false" | "False");
    Some(ProtectionInfo {
        locked: locked_attr.as_deref().map(parse_bool).unwrap_or(true),
        hidden: hidden_attr.as_deref().map(parse_bool).unwrap_or(false),
    })
}

#[derive(Debug, Default)]
struct FontBuilder {
    font: FontInfo,
}

impl FontBuilder {
    fn apply_tag(&mut self, tag: &[u8], e: &BytesStart<'_>) {
        match tag {
            b"b" => self.font.bold = parse_bool_attr_or_true(e, b"val"),
            b"i" => self.font.italic = parse_bool_attr_or_true(e, b"val"),
            b"strike" => self.font.strikethrough = parse_bool_attr_or_true(e, b"val"),
            b"u" => {
                self.font.underline =
                    Some(attr_value(e, b"val").unwrap_or_else(|| "single".to_string()));
            }
            b"name" => self.font.name = attr_value(e, b"val"),
            b"sz" => {
                self.font.size = attr_value(e, b"val").and_then(|value| value.parse().ok());
            }
            b"color" => self.font.color = parse_ooxml_color(e),
            _ => {}
        }
    }

    fn finish(self) -> FontInfo {
        self.font
    }
}

#[derive(Debug, Default)]
struct FillBuilder {
    pattern_type: Option<String>,
    fg_color: Option<String>,
    bg_color: Option<String>,
    gradient: Option<GradientInfo>,
    /// Index into `gradient.stops` of the currently-open `<stop>` element.
    pending_stop: Option<usize>,
}

impl FillBuilder {
    fn apply_color_tag(&mut self, tag: &[u8], e: &BytesStart<'_>) {
        match tag {
            b"fgColor" => self.fg_color = parse_ooxml_color(e),
            b"bgColor" => self.bg_color = parse_ooxml_color(e),
            _ => {}
        }
    }

    fn start_gradient(&mut self, e: &BytesStart<'_>) {
        let mut gi = GradientInfo {
            gradient_type: attr_value(e, b"type").unwrap_or_else(|| "linear".to_string()),
            degree: attr_value(e, b"degree").unwrap_or_default(),
            left: attr_value(e, b"left").unwrap_or_default(),
            right: attr_value(e, b"right").unwrap_or_default(),
            top: attr_value(e, b"top").unwrap_or_default(),
            bottom: attr_value(e, b"bottom").unwrap_or_default(),
            stops: Vec::new(),
        };
        if gi.gradient_type.is_empty() {
            gi.gradient_type = "linear".to_string();
        }
        self.gradient = Some(gi);
    }

    fn open_stop(&mut self, e: &BytesStart<'_>) {
        if let Some(grad) = self.gradient.as_mut() {
            grad.stops.push(GradientStopInfo {
                position: attr_value(e, b"position").unwrap_or_default(),
                color: None,
            });
            self.pending_stop = Some(grad.stops.len() - 1);
        }
    }

    fn close_stop(&mut self) {
        self.pending_stop = None;
    }

    fn apply_stop_color(&mut self, e: &BytesStart<'_>) {
        if let (Some(grad), Some(idx)) = (self.gradient.as_mut(), self.pending_stop) {
            if let Some(stop) = grad.stops.get_mut(idx) {
                stop.color = parse_ooxml_color(e);
            }
        }
    }

    fn finish(self) -> FillInfo {
        if let Some(grad) = self.gradient {
            return FillInfo {
                bg_color: None,
                gradient: Some(grad),
            };
        }
        let is_none = self
            .pattern_type
            .as_deref()
            .is_none_or(|pattern| pattern.eq_ignore_ascii_case("none"));
        FillInfo {
            bg_color: (!is_none)
                .then(|| self.fg_color.or(self.bg_color))
                .flatten(),
            gradient: None,
        }
    }
}

fn parse_alignment(e: &BytesStart<'_>) -> Option<AlignmentInfo> {
    let horizontal = attr_value(e, b"horizontal").filter(|value| value != "general");
    let vertical = attr_value(e, b"vertical");
    let wrap_text = attr_truthy(attr_value(e, b"wrapText").as_deref());
    let text_rotation = attr_value(e, b"textRotation").and_then(|value| value.parse().ok());
    let indent = attr_value(e, b"indent").and_then(|value| value.parse().ok());
    let alignment = AlignmentInfo {
        horizontal,
        vertical,
        wrap_text,
        text_rotation,
        indent,
    };
    (alignment.horizontal.is_some()
        || alignment.vertical.is_some()
        || alignment.wrap_text
        || alignment
            .text_rotation
            .is_some_and(|rotation| rotation != 0)
        || alignment.indent.is_some_and(|indent| indent != 0))
    .then_some(alignment)
}

fn parse_bool_attr_or_true(e: &BytesStart<'_>, key: &[u8]) -> bool {
    attr_value(e, key)
        .map(|value| attr_truthy(Some(&value)))
        .unwrap_or(true)
}

fn parse_ooxml_color(e: &BytesStart<'_>) -> Option<String> {
    if let Some(rgb) = attr_value(e, b"rgb") {
        return Some(normalize_ooxml_rgb(&rgb));
    }
    if let Some(indexed) = attr_value(e, b"indexed").and_then(|value| value.parse::<usize>().ok()) {
        return Some(indexed_color_hex(indexed));
    }
    attr_value(e, b"theme").map(|_| "#000000".to_string())
}

#[derive(Debug, Clone, Copy)]
enum BorderEdgeKind {
    Left,
    Right,
    Top,
    Bottom,
    Diagonal,
}

impl BorderEdgeKind {
    fn from_tag(tag: &[u8]) -> Option<Self> {
        match tag {
            b"left" => Some(Self::Left),
            b"right" => Some(Self::Right),
            b"top" => Some(Self::Top),
            b"bottom" => Some(Self::Bottom),
            b"diagonal" => Some(Self::Diagonal),
            _ => None,
        }
    }
}

#[derive(Debug, Default)]
struct BorderBuilder {
    border: BorderInfo,
    diagonal_up: bool,
    diagonal_down: bool,
}

impl BorderBuilder {
    fn from_start(e: &BytesStart<'_>) -> Self {
        Self {
            border: BorderInfo::default(),
            diagonal_up: attr_truthy(attr_value(e, b"diagonalUp").as_deref()),
            diagonal_down: attr_truthy(attr_value(e, b"diagonalDown").as_deref()),
        }
    }

    fn set_edge(&mut self, edge: BorderEdgeKind, side: Option<BorderSide>) {
        match edge {
            BorderEdgeKind::Left => self.border.left = side,
            BorderEdgeKind::Right => self.border.right = side,
            BorderEdgeKind::Top => self.border.top = side,
            BorderEdgeKind::Bottom => self.border.bottom = side,
            BorderEdgeKind::Diagonal => {
                if self.diagonal_up {
                    self.border.diagonal_up = side.clone();
                }
                if self.diagonal_down {
                    self.border.diagonal_down = side;
                }
            }
        }
    }

    fn update_edge_color(&mut self, edge: BorderEdgeKind, color: Option<String>) {
        let Some(color) = color else {
            return;
        };
        match edge {
            BorderEdgeKind::Left => {
                if let Some(side) = self.border.left.as_mut() {
                    side.color = color;
                }
            }
            BorderEdgeKind::Right => {
                if let Some(side) = self.border.right.as_mut() {
                    side.color = color;
                }
            }
            BorderEdgeKind::Top => {
                if let Some(side) = self.border.top.as_mut() {
                    side.color = color;
                }
            }
            BorderEdgeKind::Bottom => {
                if let Some(side) = self.border.bottom.as_mut() {
                    side.color = color;
                }
            }
            BorderEdgeKind::Diagonal => {
                if let Some(side) = self.border.diagonal_up.as_mut() {
                    side.color = color.clone();
                }
                if let Some(side) = self.border.diagonal_down.as_mut() {
                    side.color = color;
                }
            }
        }
    }

    fn finish(self) -> BorderInfo {
        self.border
    }
}

fn parse_border_side(e: &BytesStart<'_>) -> Option<BorderSide> {
    let style = attr_value(e, b"style").filter(|value| !value.is_empty() && value != "none")?;
    let color = parse_border_color(e).unwrap_or_else(|| "#000000".to_string());
    Some(BorderSide { style, color })
}

fn parse_border_color(e: &BytesStart<'_>) -> Option<String> {
    parse_ooxml_color(e)
}

fn normalize_ooxml_rgb(rgb: &str) -> String {
    let raw = rgb.trim().trim_start_matches('#');
    let rgb = if raw.len() == 8 { &raw[2..] } else { raw };
    format!("#{}", rgb.to_ascii_uppercase())
}

fn indexed_color_hex(index: usize) -> String {
    const COLORS: [&str; 66] = [
        "#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF", "#FFFF00", "#FF00FF", "#00FFFF",
        "#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF", "#FFFF00", "#FF00FF", "#00FFFF",
        "#800000", "#008000", "#000080", "#808000", "#800080", "#008080", "#C0C0C0", "#808080",
        "#9999FF", "#993366", "#FFFFCC", "#CCFFFF", "#660066", "#FF8080", "#0066CC", "#CCCCFF",
        "#000080", "#FF00FF", "#FFFF00", "#00FFFF", "#800080", "#800000", "#008080", "#0000FF",
        "#00CCFF", "#CCFFFF", "#CCFFCC", "#FFFF99", "#99CCFF", "#FF99CC", "#CC99FF", "#FFCC99",
        "#3366FF", "#33CCCC", "#99CC00", "#FFCC00", "#FF9900", "#FF6600", "#666699", "#969696",
        "#003366", "#339966", "#003300", "#333300", "#993300", "#993366", "#333399", "#333333",
        "#000000", "#FFFFFF",
    ];
    COLORS.get(index).unwrap_or(&"#000000").to_string()
}

fn builtin_num_fmt(format_id: u32) -> Option<&'static str> {
    match format_id {
        0 => Some("General"),
        1 => Some("0"),
        2 => Some("0.00"),
        3 => Some("#,##0"),
        4 => Some("#,##0.00"),
        5 => Some("\"$\"#,##0_);(\"$\"#,##0)"),
        6 => Some("\"$\"#,##0_);[Red](\"$\"#,##0)"),
        7 => Some("\"$\"#,##0.00_);(\"$\"#,##0.00)"),
        8 => Some("\"$\"#,##0.00_);[Red](\"$\"#,##0.00)"),
        9 => Some("0%"),
        10 => Some("0.00%"),
        11 => Some("0.00E+00"),
        12 => Some("# ?/?"),
        13 => Some("# ??/??"),
        14 => Some("mm-dd-yy"),
        15 => Some("d-mmm-yy"),
        16 => Some("d-mmm"),
        17 => Some("mmm-yy"),
        18 => Some("h:mm AM/PM"),
        19 => Some("h:mm:ss AM/PM"),
        20 => Some("h:mm"),
        21 => Some("h:mm:ss"),
        22 => Some("m/d/yy h:mm"),
        37 => Some("#,##0_);(#,##0)"),
        38 => Some("#,##0_);[Red](#,##0)"),
        39 => Some("#,##0.00_);(#,##0.00)"),
        40 => Some("#,##0.00_);[Red](#,##0.00)"),
        41 => Some(r#"_(* #,##0_);_(* \(#,##0\);_(* "-"_);_(@_)"#),
        42 => Some(r#"_("$"* #,##0_);_("$"* \(#,##0\);_("$"* "-"_);_(@_)"#),
        43 => Some(r#"_(* #,##0.00_);_(* \(#,##0.00\);_(* "-"??_);_(@_)"#),
        44 => Some(r#"_("$"* #,##0.00_)_("$"* \(#,##0.00\)_("$"* "-"??_)_(@_)"#),
        45 => Some("mm:ss"),
        46 => Some("[h]:mm:ss"),
        47 => Some("mmss.0"),
        48 => Some("##0.0E+0"),
        49 => Some("@"),
        _ => None,
    }
}

fn parse_worksheet(
    xml: &str,
    shared_strings: &SharedStrings,
    rels: Option<&RelsGraph>,
    comments: Vec<Comment>,
    tables: Vec<Table>,
) -> Result<WorksheetData> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut dimension = None;
    let mut merged_ranges = Vec::new();
    let mut hyperlink_nodes = Vec::new();
    let mut freeze_panes = None;
    let mut sheet_properties = None;
    let mut sheet_view = None;
    let mut current_sheet_view: Option<SheetViewInfo> = None;
    let mut row_heights = HashMap::new();
    let mut column_widths = Vec::new();
    let mut data_validations = Vec::new();
    let mut sheet_protection = None;
    let mut page_margins = None;
    let mut page_setup = None;
    let mut print_options = None;
    let mut header_footer = None;
    let mut row_breaks = None;
    let mut column_breaks = None;
    let mut sheet_format = None;
    let mut conditional_formats = Vec::new();
    let mut hidden_rows: HashMap<u32, bool> = HashMap::new();
    let mut hidden_columns: HashMap<u32, bool> = HashMap::new();
    let mut row_outline_levels: HashMap<u32, u8> = HashMap::new();
    let mut column_outline_levels: HashMap<u32, u8> = HashMap::new();
    let mut current_conditional_range: Option<String> = None;
    let mut current_conditional_rule: Option<ConditionalFormatBuilder> = None;
    let mut in_conditional_formula = false;
    let mut row_index: Option<u32> = None;
    let mut current: Option<CellBuilder> = None;
    let mut active_text: Option<TextTarget> = None;
    let mut current_validation: Option<DataValidationBuilder> = None;
    let mut active_validation_text: Option<DataValidationFormula> = None;
    let mut current_header_footer: Option<HeaderFooterInfo> = None;
    let mut active_header_footer: Option<HeaderFooterPart> = None;
    let mut active_breaks: Option<BreakListKind> = None;
    let mut cells = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.local_name().as_ref() {
                b"sheetPr" => {
                    sheet_properties = Some(SheetPropertiesInfo::from_start(&e));
                }
                b"tabColor" => {
                    sheet_properties
                        .get_or_insert_with(SheetPropertiesInfo::default)
                        .apply_tab_color(&e);
                }
                b"outlinePr" => {
                    sheet_properties
                        .get_or_insert_with(SheetPropertiesInfo::default)
                        .apply_outline(&e);
                }
                b"pageSetUpPr" => {
                    sheet_properties
                        .get_or_insert_with(SheetPropertiesInfo::default)
                        .apply_page_setup(&e);
                }
                b"dimension" => {
                    dimension = attr_value(&e, b"ref");
                }
                b"mergeCell" => {
                    if let Some(range) = attr_value(&e, b"ref") {
                        merged_ranges.push(range);
                    }
                }
                b"hyperlink" => {
                    if let Some(node) = HyperlinkNode::from_start(&e) {
                        hyperlink_nodes.push(node);
                    }
                }
                b"pane" => {
                    if let Some(pane) = parse_pane(&e) {
                        freeze_panes = Some(pane.clone());
                        if let Some(view) = current_sheet_view.as_mut() {
                            view.pane = Some(pane);
                        }
                    }
                }
                b"sheetView" => {
                    if current_sheet_view.is_none() && sheet_view.is_none() {
                        current_sheet_view = Some(SheetViewInfo::from_start(&e));
                    }
                }
                b"selection" => {
                    if let Some(view) = current_sheet_view.as_mut() {
                        view.selections.push(SelectionInfo::from_start(&e));
                    }
                }
                b"sheetFormatPr" => {
                    sheet_format = Some(SheetFormatInfo::from_start(&e));
                }
                b"row" => {
                    row_index = attr_value(&e, b"r").and_then(|v| v.parse::<u32>().ok());
                    if let Some((row, height)) = parse_row_height(&e, row_index) {
                        row_heights.insert(row, height);
                    }
                    update_row_visibility(&e, row_index, &mut hidden_rows, &mut row_outline_levels);
                }
                b"c" => {
                    current = Some(CellBuilder::from_start(&e, row_index));
                }
                b"dataValidation" => {
                    current_validation = Some(DataValidationBuilder::from_start(&e));
                }
                b"sheetProtection" => {
                    sheet_protection = Some(SheetProtection::from_start(&e));
                }
                b"pageMargins" => {
                    page_margins = parse_page_margins(&e);
                }
                b"printOptions" => {
                    print_options = Some(PrintOptionsInfo::from_start(&e));
                }
                b"pageSetup" => {
                    page_setup = Some(PageSetupInfo::from_start(&e));
                }
                b"headerFooter" => {
                    current_header_footer = Some(HeaderFooterInfo::from_start(&e));
                }
                b"oddHeader" => active_header_footer = Some(HeaderFooterPart::OddHeader),
                b"oddFooter" => active_header_footer = Some(HeaderFooterPart::OddFooter),
                b"evenHeader" => active_header_footer = Some(HeaderFooterPart::EvenHeader),
                b"evenFooter" => active_header_footer = Some(HeaderFooterPart::EvenFooter),
                b"firstHeader" => active_header_footer = Some(HeaderFooterPart::FirstHeader),
                b"firstFooter" => active_header_footer = Some(HeaderFooterPart::FirstFooter),
                b"rowBreaks" => {
                    row_breaks = Some(PageBreakListInfo::from_start(&e));
                    active_breaks = Some(BreakListKind::Row);
                }
                b"colBreaks" => {
                    column_breaks = Some(PageBreakListInfo::from_start(&e));
                    active_breaks = Some(BreakListKind::Column);
                }
                b"brk" => append_break(&mut row_breaks, &mut column_breaks, active_breaks, &e),
                b"conditionalFormatting" => {
                    current_conditional_range = attr_value(&e, b"sqref");
                }
                b"cfRule" => {
                    current_conditional_rule = Some(ConditionalFormatBuilder::from_start(
                        &e,
                        current_conditional_range.clone().unwrap_or_default(),
                    ));
                }
                b"colorScale" => {
                    if let Some(rule) = current_conditional_rule.as_mut() {
                        rule.start_color_scale();
                    }
                }
                b"cfvo" => {
                    if let Some(rule) = current_conditional_rule
                        .as_mut()
                        .filter(|r| r.color_scale.is_some())
                    {
                        rule.push_cfvo(&e);
                    }
                }
                b"color" => {
                    // Only consume <color> when it's a child of <colorScale>
                    // inside the current cfRule. Otherwise fall through so
                    // inline rich-text (<rPr><color/></rPr>) is handled by
                    // the inline-string fallback below.
                    if let Some(rule) = current_conditional_rule
                        .as_mut()
                        .filter(|r| r.color_scale.is_some())
                    {
                        rule.push_color(&e);
                    } else if let Some(cell) = current
                        .as_mut()
                        .filter(|c| c.data_type == CellDataType::InlineString)
                    {
                        cell.apply_inline_font_tag(b"color", e.attributes());
                    }
                }
                b"formula" => {
                    if current_conditional_rule.is_some() {
                        in_conditional_formula = true;
                    }
                }
                b"formula1" => {
                    if current_validation.is_some() {
                        active_validation_text = Some(DataValidationFormula::Formula1);
                    }
                }
                b"formula2" => {
                    if current_validation.is_some() {
                        active_validation_text = Some(DataValidationFormula::Formula2);
                    }
                }
                b"v" => active_text = Some(TextTarget::Value),
                b"f" => {
                    if let Some(cell) = current.as_mut() {
                        cell.start_formula(&e);
                    }
                    active_text = Some(TextTarget::Formula);
                }
                b"t" => {
                    if current
                        .as_ref()
                        .is_some_and(|c| c.data_type == CellDataType::InlineString)
                    {
                        active_text = Some(TextTarget::InlineString);
                    }
                }
                b"r" => {
                    if let Some(cell) = current
                        .as_mut()
                        .filter(|c| c.data_type == CellDataType::InlineString)
                    {
                        cell.start_inline_run();
                    }
                }
                b"rPr" => {
                    if let Some(cell) = current
                        .as_mut()
                        .filter(|c| c.data_type == CellDataType::InlineString)
                    {
                        cell.start_inline_props();
                    }
                }
                other => {
                    if let Some(cell) = current
                        .as_mut()
                        .filter(|c| c.data_type == CellDataType::InlineString)
                    {
                        cell.apply_inline_font_tag(other, e.attributes());
                    }
                }
            },
            Ok(Event::Empty(e)) => match e.local_name().as_ref() {
                b"sheetPr" => {
                    sheet_properties = Some(SheetPropertiesInfo::from_start(&e));
                }
                b"tabColor" => {
                    sheet_properties
                        .get_or_insert_with(SheetPropertiesInfo::default)
                        .apply_tab_color(&e);
                }
                b"outlinePr" => {
                    sheet_properties
                        .get_or_insert_with(SheetPropertiesInfo::default)
                        .apply_outline(&e);
                }
                b"pageSetUpPr" => {
                    sheet_properties
                        .get_or_insert_with(SheetPropertiesInfo::default)
                        .apply_page_setup(&e);
                }
                b"col" => {
                    if let Some(width) = parse_column_width(&e) {
                        column_widths.push(width);
                    }
                    update_column_visibility(&e, &mut hidden_columns, &mut column_outline_levels);
                }
                b"dimension" => {
                    dimension = attr_value(&e, b"ref");
                }
                b"mergeCell" => {
                    if let Some(range) = attr_value(&e, b"ref") {
                        merged_ranges.push(range);
                    }
                }
                b"hyperlink" => {
                    if let Some(node) = HyperlinkNode::from_start(&e) {
                        hyperlink_nodes.push(node);
                    }
                }
                b"pane" => {
                    if let Some(pane) = parse_pane(&e) {
                        freeze_panes = Some(pane.clone());
                        if let Some(view) = current_sheet_view.as_mut() {
                            view.pane = Some(pane);
                        }
                    }
                }
                b"sheetView" => {
                    if sheet_view.is_none() {
                        sheet_view = Some(SheetViewInfo::from_start(&e));
                    }
                }
                b"selection" => {
                    if let Some(view) = current_sheet_view.as_mut() {
                        view.selections.push(SelectionInfo::from_start(&e));
                    }
                }
                b"sheetFormatPr" => {
                    sheet_format = Some(SheetFormatInfo::from_start(&e));
                }
                b"c" => {
                    let builder = CellBuilder::from_start(&e, row_index);
                    cells.push(builder.finish(shared_strings)?);
                }
                b"f" => {
                    if let Some(cell) = current.as_mut() {
                        cell.start_formula(&e);
                    }
                }
                b"row" => {
                    let row = attr_value(&e, b"r").and_then(|v| v.parse::<u32>().ok());
                    if let Some((row, height)) = parse_row_height(&e, row) {
                        row_heights.insert(row, height);
                    }
                    update_row_visibility(&e, row, &mut hidden_rows, &mut row_outline_levels);
                }
                b"dataValidation" => {
                    let validation = DataValidationBuilder::from_start(&e).finish();
                    if !validation.range.trim().is_empty() {
                        data_validations.push(validation);
                    }
                }
                b"sheetProtection" => {
                    sheet_protection = Some(SheetProtection::from_start(&e));
                }
                b"pageMargins" => {
                    page_margins = parse_page_margins(&e);
                }
                b"printOptions" => {
                    print_options = Some(PrintOptionsInfo::from_start(&e));
                }
                b"pageSetup" => {
                    page_setup = Some(PageSetupInfo::from_start(&e));
                }
                b"headerFooter" => {
                    header_footer = Some(HeaderFooterInfo::from_start(&e));
                }
                b"rowBreaks" => {
                    row_breaks = Some(PageBreakListInfo::from_start(&e));
                }
                b"colBreaks" => {
                    column_breaks = Some(PageBreakListInfo::from_start(&e));
                }
                b"brk" => append_break(&mut row_breaks, &mut column_breaks, active_breaks, &e),
                b"conditionalFormatting" => {
                    current_conditional_range = attr_value(&e, b"sqref");
                }
                b"cfRule" => {
                    let rule = ConditionalFormatBuilder::from_start(
                        &e,
                        current_conditional_range.clone().unwrap_or_default(),
                    )
                    .finish();
                    if !rule.range.trim().is_empty() && !rule.rule_type.trim().is_empty() {
                        conditional_formats.push(rule);
                    }
                }
                b"cfvo" => {
                    if let Some(rule) = current_conditional_rule
                        .as_mut()
                        .filter(|r| r.color_scale.is_some())
                    {
                        rule.push_cfvo(&e);
                    }
                }
                b"color" => {
                    // Only consume <color/> when inside a <colorScale> in
                    // the current cfRule; otherwise treat as an inline-font
                    // tag for rich-text inline strings.
                    if let Some(rule) = current_conditional_rule
                        .as_mut()
                        .filter(|r| r.color_scale.is_some())
                    {
                        rule.push_color(&e);
                    } else if let Some(cell) = current
                        .as_mut()
                        .filter(|c| c.data_type == CellDataType::InlineString)
                    {
                        cell.apply_inline_font_tag(b"color", e.attributes());
                    }
                }
                other => {
                    if let Some(cell) = current
                        .as_mut()
                        .filter(|c| c.data_type == CellDataType::InlineString)
                    {
                        cell.apply_inline_font_tag(other, e.attributes());
                    }
                }
            },
            Ok(Event::End(e)) => match e.local_name().as_ref() {
                b"c" => {
                    if let Some(builder) = current.take() {
                        cells.push(builder.finish(shared_strings)?);
                    }
                }
                b"formula1" => {
                    active_validation_text = None;
                    if let Some(validation) = current_validation.as_mut() {
                        validation.finish_formula1();
                    }
                }
                b"formula2" => {
                    active_validation_text = None;
                    if let Some(validation) = current_validation.as_mut() {
                        validation.finish_formula2();
                    }
                }
                b"dataValidation" => {
                    active_validation_text = None;
                    if let Some(validation) = current_validation.take() {
                        let validation = validation.finish();
                        if !validation.range.trim().is_empty() {
                            data_validations.push(validation);
                        }
                    }
                }
                b"oddHeader" | b"oddFooter" | b"evenHeader" | b"evenFooter" | b"firstHeader"
                | b"firstFooter" => {
                    active_header_footer = None;
                }
                b"headerFooter" => {
                    header_footer = current_header_footer.take();
                    active_header_footer = None;
                }
                b"rowBreaks" | b"colBreaks" => {
                    active_breaks = None;
                }
                b"sheetView" => {
                    if sheet_view.is_none() {
                        sheet_view = current_sheet_view.take();
                    } else {
                        current_sheet_view = None;
                    }
                }
                b"conditionalFormatting" => {
                    current_conditional_range = None;
                }
                b"formula" => {
                    if in_conditional_formula {
                        in_conditional_formula = false;
                        if let Some(rule) = current_conditional_rule.as_mut() {
                            rule.finish_formula();
                        }
                    } else {
                        active_text = None;
                    }
                }
                b"cfRule" => {
                    in_conditional_formula = false;
                    if let Some(rule) = current_conditional_rule.take() {
                        let rule = rule.finish();
                        if !rule.range.trim().is_empty() && !rule.rule_type.trim().is_empty() {
                            conditional_formats.push(rule);
                        }
                    }
                }
                b"v" | b"f" | b"t" => active_text = None,
                b"rPr" => {
                    if let Some(cell) = current.as_mut() {
                        cell.end_inline_props();
                    }
                }
                b"r" => {
                    if let Some(cell) = current.as_mut() {
                        cell.end_inline_run();
                    }
                }
                b"row" => row_index = None,
                _ => {}
            },
            Ok(Event::Text(e)) => {
                if in_conditional_formula {
                    if let Some(rule) = current_conditional_rule.as_mut() {
                        let text = e
                            .unescape()
                            .map_err(|err| ReaderError::Xml(format!("worksheet text: {err}")))?;
                        rule.push_formula_text(&text);
                    }
                } else if let (Some(target), Some(validation)) =
                    (active_validation_text, current_validation.as_mut())
                {
                    let text = e
                        .unescape()
                        .map_err(|err| ReaderError::Xml(format!("worksheet text: {err}")))?;
                    validation.push_text(target, &text);
                } else if let (Some(part), Some(header_footer)) =
                    (active_header_footer, current_header_footer.as_mut())
                {
                    let text = e
                        .unescape()
                        .map_err(|err| ReaderError::Xml(format!("worksheet text: {err}")))?;
                    header_footer.set_part(part, parse_header_footer_item_text(&text));
                } else if let (Some(target), Some(cell)) = (active_text, current.as_mut()) {
                    let text = e
                        .unescape()
                        .map_err(|err| ReaderError::Xml(format!("worksheet text: {err}")))?;
                    cell.push_text(target, &text);
                }
            }
            Ok(Event::CData(e)) => {
                if in_conditional_formula {
                    if let Some(rule) = current_conditional_rule.as_mut() {
                        rule.push_formula_text(&String::from_utf8_lossy(e.as_ref()));
                    }
                } else if let (Some(target), Some(validation)) =
                    (active_validation_text, current_validation.as_mut())
                {
                    validation.push_text(target, &String::from_utf8_lossy(e.as_ref()));
                } else if let (Some(part), Some(header_footer)) =
                    (active_header_footer, current_header_footer.as_mut())
                {
                    header_footer.set_part(
                        part,
                        parse_header_footer_item_text(&String::from_utf8_lossy(e.as_ref())),
                    );
                } else if let (Some(target), Some(cell)) = (active_text, current.as_mut()) {
                    cell.push_text(target, &String::from_utf8_lossy(e.as_ref()));
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(ReaderError::Xml(format!("failed to parse worksheet: {e}"))),
            _ => {}
        }
        buf.clear();
    }

    let mut hidden_rows: Vec<u32> = hidden_rows
        .into_iter()
        .filter_map(|(row, hidden)| hidden.then_some(row))
        .collect();
    let mut hidden_columns: Vec<u32> = hidden_columns
        .into_iter()
        .filter_map(|(col, hidden)| hidden.then_some(col))
        .collect();
    let mut row_outline_levels: Vec<(u32, u8)> = row_outline_levels.into_iter().collect();
    let mut column_outline_levels: Vec<(u32, u8)> = column_outline_levels.into_iter().collect();
    hidden_rows.sort_unstable();
    hidden_columns.sort_unstable();
    row_outline_levels.sort_unstable_by_key(|(row, _)| *row);
    column_outline_levels.sort_unstable_by_key(|(col, _)| *col);

    let auto_filter = parse_auto_filter(xml)?;
    resolve_shared_formulas(&mut cells);
    let array_formulas = build_array_formula_map(&cells);

    Ok(WorksheetData {
        dimension,
        merged_ranges,
        hyperlinks: resolve_hyperlinks(hyperlink_nodes, rels),
        freeze_panes,
        sheet_properties,
        sheet_view,
        comments,
        threaded_comments: Vec::new(),
        row_heights,
        column_widths,
        data_validations,
        sheet_protection,
        auto_filter,
        page_margins,
        page_setup,
        print_options,
        header_footer,
        row_breaks,
        column_breaks,
        sheet_format,
        images: Vec::new(),
        charts: Vec::new(),
        tables,
        conditional_formats,
        hidden_rows,
        hidden_columns,
        row_outline_levels,
        column_outline_levels,
        array_formulas,
        cells,
    })
}

#[cfg(test)]
fn parse_worksheet_value_cells(
    xml: &str,
    shared_strings: &SharedStrings,
) -> Result<Vec<ValueCell>> {
    parse_worksheet_value_cells_bytes(xml.as_bytes(), shared_strings)
}

fn parse_worksheet_value_cells_bytes(
    xml: &[u8],
    shared_strings: &SharedStrings,
) -> Result<Vec<ValueCell>> {
    let mut reader = XmlReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut row_index: Option<u32> = None;
    let mut current: Option<ValueCellBuilder> = None;
    let mut active_text: Option<TextTarget> = None;
    let mut in_sheet_data = false;
    let mut cells = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.local_name().as_ref() {
                b"dimension" if !in_sheet_data => reserve_value_cell_capacity(&mut cells, &e),
                b"sheetData" => in_sheet_data = true,
                b"row" if in_sheet_data => row_index = raw_attr_u32(&e, b"r"),
                b"c" if in_sheet_data => {
                    current = Some(ValueCellBuilder::from_start(&e, row_index));
                }
                b"v" if current.is_some() => active_text = Some(TextTarget::Value),
                b"t" if current
                    .as_ref()
                    .is_some_and(ValueCellBuilder::is_inline_string) =>
                {
                    active_text = Some(TextTarget::InlineString);
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.local_name().as_ref() {
                b"dimension" if !in_sheet_data => reserve_value_cell_capacity(&mut cells, &e),
                b"sheetData" => break,
                b"row" if in_sheet_data => {
                    row_index = None;
                }
                b"c" if in_sheet_data => {
                    let builder = ValueCellBuilder::from_start(&e, row_index);
                    cells.push(builder.finish(shared_strings)?);
                }
                _ => {}
            },
            Ok(Event::End(e)) => match e.local_name().as_ref() {
                b"sheetData" => break,
                b"row" if in_sheet_data => row_index = None,
                b"c" if in_sheet_data => {
                    if let Some(builder) = current.take() {
                        cells.push(builder.finish(shared_strings)?);
                    }
                }
                b"v" | b"t" => active_text = None,
                _ => {}
            },
            Ok(Event::Text(e)) => {
                if let (Some(target), Some(cell)) = (active_text, current.as_mut()) {
                    match target {
                        TextTarget::InlineString => {
                            push_inline_value_text(cell, e.as_ref(), || {
                                e.unescape().map(|text| text.into_owned()).map_err(|err| {
                                    ReaderError::Xml(format!("worksheet text: {err}"))
                                })
                            })?;
                        }
                        TextTarget::Value | TextTarget::Formula => {
                            cell.push_raw_text(target, e.as_ref());
                        }
                    }
                }
            }
            Ok(Event::CData(e)) => {
                if let (Some(target), Some(cell)) = (active_text, current.as_mut()) {
                    cell.push_text(target, &String::from_utf8_lossy(e.as_ref()));
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(ReaderError::Xml(format!("failed to parse worksheet: {e}"))),
            _ => {}
        }
        buf.clear();
    }

    Ok(cells)
}

fn read_first_sheet_value_cells_source(source: &WorkbookSource) -> Result<Vec<ValueCell>> {
    let mut zip = zip_from_source(source)?;
    validate_zip_archive(&mut zip)?;

    let workbook_xml = read_part_required(&mut zip, "xl/workbook.xml")?;
    let workbook_rels = read_part_required(&mut zip, "xl/_rels/workbook.xml.rels")?;
    let rels = RelsGraph::parse(workbook_rels.as_bytes())
        .map_err(|e| ReaderError::Xml(format!("failed to parse workbook rels: {e}")))?;
    let (sheet_refs, ..) = parse_workbook(&workbook_xml)?;
    let sheets = resolve_sheet_paths(sheet_refs, &rels)?;
    let Some(info) = sheets.first() else {
        return Err(ReaderError::SheetNotFound("first worksheet".to_string()));
    };
    let shared_strings = match read_part_optional(&mut zip, "xl/sharedStrings.xml")? {
        Some(xml) => parse_shared_strings(&xml)?,
        None => SharedStrings::default(),
    };
    let xml = read_part_required_bytes(&mut zip, &info.path)?;
    parse_worksheet_value_cells_bytes(&xml, &shared_strings)
}

fn read_first_sheet_value_summary_source(source: &WorkbookSource) -> Result<ValueSheetSummary> {
    let mut zip = zip_from_source(source)?;
    validate_zip_archive(&mut zip)?;

    let workbook_xml = read_part_required(&mut zip, "xl/workbook.xml")?;
    let workbook_rels = read_part_required(&mut zip, "xl/_rels/workbook.xml.rels")?;
    let rels = RelsGraph::parse(workbook_rels.as_bytes())
        .map_err(|e| ReaderError::Xml(format!("failed to parse workbook rels: {e}")))?;
    let (sheet_refs, ..) = parse_workbook(&workbook_xml)?;
    let sheets = resolve_sheet_paths(sheet_refs, &rels)?;
    let Some(info) = sheets.first() else {
        return Err(ReaderError::SheetNotFound("first worksheet".to_string()));
    };
    let xml = read_part_required_bytes(&mut zip, &info.path)?;
    parse_worksheet_value_summary_bytes(&xml)
}

fn read_first_sheet_formula_summary_source(source: &WorkbookSource) -> Result<FormulaSheetSummary> {
    let mut zip = zip_from_source(source)?;
    validate_zip_archive(&mut zip)?;

    let workbook_xml = read_part_required(&mut zip, "xl/workbook.xml")?;
    let workbook_rels = read_part_required(&mut zip, "xl/_rels/workbook.xml.rels")?;
    let rels = RelsGraph::parse(workbook_rels.as_bytes())
        .map_err(|e| ReaderError::Xml(format!("failed to parse workbook rels: {e}")))?;
    let (sheet_refs, ..) = parse_workbook(&workbook_xml)?;
    let sheets = resolve_sheet_paths(sheet_refs, &rels)?;
    let Some(info) = sheets.first() else {
        return Err(ReaderError::SheetNotFound("first worksheet".to_string()));
    };
    let xml = read_part_required_bytes(&mut zip, &info.path)?;
    parse_worksheet_formula_summary_bytes(&xml)
}

fn parse_worksheet_value_summary_bytes(xml: &[u8]) -> Result<ValueSheetSummary> {
    let mut reader = XmlReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut row_index: Option<u32> = None;
    let mut current: Option<FormulaSummaryCellBuilder> = None;
    let mut active_value = false;
    let mut in_sheet_data = false;
    let mut summary = ValueSheetSummary {
        rows: 0,
        numeric_checksum: 0.0,
    };

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.local_name().as_ref() {
                b"sheetData" => in_sheet_data = true,
                b"row" if in_sheet_data => row_index = raw_attr_u32(&e, b"r"),
                b"c" if in_sheet_data => {
                    current = Some(FormulaSummaryCellBuilder::from_start(&e, row_index));
                }
                b"v" if current.is_some() => active_value = true,
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.local_name().as_ref() {
                b"sheetData" => break,
                b"row" if in_sheet_data => row_index = None,
                b"c" if in_sheet_data => {
                    let cell = FormulaSummaryCellBuilder::from_start(&e, row_index);
                    cell.update_value_summary(&mut summary)?;
                }
                _ => {}
            },
            Ok(Event::End(e)) => match e.local_name().as_ref() {
                b"sheetData" => break,
                b"row" if in_sheet_data => row_index = None,
                b"c" if in_sheet_data => {
                    if let Some(cell) = current.take() {
                        cell.update_value_summary(&mut summary)?;
                    }
                }
                b"v" => active_value = false,
                _ => {}
            },
            Ok(Event::Text(e)) => {
                if active_value {
                    if let Some(cell) = current.as_mut() {
                        cell.push_raw_value_text(e.as_ref());
                    }
                }
            }
            Ok(Event::CData(e)) => {
                if active_value {
                    if let Some(cell) = current.as_mut() {
                        cell.push_raw_value_text(e.as_ref());
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(ReaderError::Xml(format!("failed to parse worksheet: {e}"))),
            _ => {}
        }
        buf.clear();
    }

    Ok(summary)
}

fn parse_worksheet_formula_summary_bytes(xml: &[u8]) -> Result<FormulaSheetSummary> {
    let mut reader = XmlReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut row_index: Option<u32> = None;
    let mut current: Option<FormulaSummaryCellBuilder> = None;
    let mut active_value = false;
    let mut in_sheet_data = false;
    let mut summary = FormulaSheetSummary {
        rows: 0,
        first_col_numeric_checksum: 0.0,
        formula_cells: 0,
    };

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.local_name().as_ref() {
                b"sheetData" => in_sheet_data = true,
                b"row" if in_sheet_data => row_index = raw_attr_u32(&e, b"r"),
                b"c" if in_sheet_data => {
                    current = Some(FormulaSummaryCellBuilder::from_start(&e, row_index));
                }
                b"f" if current.is_some() => {
                    if let Some(cell) = current.as_mut() {
                        cell.has_formula = true;
                    }
                }
                b"v" if current.is_some() => active_value = true,
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.local_name().as_ref() {
                b"sheetData" => break,
                b"row" if in_sheet_data => row_index = None,
                b"c" if in_sheet_data => {
                    let cell = FormulaSummaryCellBuilder::from_start(&e, row_index);
                    cell.update_summary(&mut summary)?;
                }
                b"f" if current.is_some() => {
                    if let Some(cell) = current.as_mut() {
                        cell.has_formula = true;
                    }
                }
                _ => {}
            },
            Ok(Event::End(e)) => match e.local_name().as_ref() {
                b"sheetData" => break,
                b"row" if in_sheet_data => row_index = None,
                b"c" if in_sheet_data => {
                    if let Some(cell) = current.take() {
                        cell.update_summary(&mut summary)?;
                    }
                }
                b"v" => active_value = false,
                _ => {}
            },
            Ok(Event::Text(e)) => {
                if active_value {
                    if let Some(cell) = current.as_mut() {
                        cell.push_raw_value_text(e.as_ref());
                    }
                }
            }
            Ok(Event::CData(e)) => {
                if active_value {
                    if let Some(cell) = current.as_mut() {
                        cell.value_text.extend_from_slice(e.as_ref());
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(ReaderError::Xml(format!("failed to parse worksheet: {e}"))),
            _ => {}
        }
        buf.clear();
    }

    Ok(summary)
}

fn push_inline_value_text<F>(cell: &mut ValueCellBuilder, raw: &[u8], unescape: F) -> Result<()>
where
    F: FnOnce() -> Result<String>,
{
    if !raw.iter().any(|byte| matches!(*byte, b'&' | b'\r')) {
        let text = std::str::from_utf8(raw)
            .map_err(|err| ReaderError::Xml(format!("worksheet text: {err}")))?;
        cell.inline_text.push_str(text);
        return Ok(());
    }
    cell.inline_text
        .push_str(&normalize_ooxml_text(&unescape()?));
    Ok(())
}

fn reserve_value_cell_capacity(cells: &mut Vec<ValueCell>, e: &BytesStart<'_>) {
    let Some(capacity) = value_cell_capacity_from_dimension(e) else {
        return;
    };
    cells.reserve_exact(capacity);
}

fn value_cell_capacity_from_dimension(e: &BytesStart<'_>) -> Option<usize> {
    for attr in e.attributes().with_checks(false).flatten() {
        if attr.key.as_ref() == b"ref" {
            let (min_row, min_col, max_row, max_col) =
                parse_range_bounds_1based_bytes(attr.value.as_ref())?;
            let row_count = max_row.checked_sub(min_row)?.checked_add(1)? as usize;
            let col_count = max_col.checked_sub(min_col)?.checked_add(1)? as usize;
            let capacity = row_count.checked_mul(col_count)?;
            return (capacity <= MAX_VALUE_CELL_PREALLOC).then_some(capacity);
        }
    }
    None
}

fn parse_range_bounds_1based_bytes(range: &[u8]) -> Option<(u32, u32, u32, u32)> {
    let mut parts = range.split(|byte| *byte == b':');
    let start = parts.next()?;
    let end = parts.next().unwrap_or(start);
    let (start_row, start_col) = a1_bytes_to_row_col(start)?;
    let (end_row, end_col) = a1_bytes_to_row_col(end)?;
    Some((
        start_row.min(end_row),
        start_col.min(end_col),
        start_row.max(end_row),
        start_col.max(end_col),
    ))
}

fn parse_row_height(e: &BytesStart<'_>, row: Option<u32>) -> Option<(u32, RowHeight)> {
    let row = row?;
    let height = attr_value(e, b"ht")?.parse::<f64>().ok()?;
    Some((
        row,
        RowHeight {
            height,
            custom_height: attr_truthy(attr_value(e, b"customHeight").as_deref()),
        },
    ))
}

fn parse_column_width(e: &BytesStart<'_>) -> Option<ColumnWidth> {
    Some(ColumnWidth {
        min: attr_value(e, b"min")?.parse::<u32>().ok()?,
        max: attr_value(e, b"max")?.parse::<u32>().ok()?,
        width: attr_value(e, b"width")?.parse::<f64>().ok()?,
        custom_width: attr_truthy(attr_value(e, b"customWidth").as_deref()),
    })
}

fn parse_page_margins(e: &BytesStart<'_>) -> Option<PageMarginsInfo> {
    Some(PageMarginsInfo {
        left: attr_f64(e, b"left")?,
        right: attr_f64(e, b"right")?,
        top: attr_f64(e, b"top")?,
        bottom: attr_f64(e, b"bottom")?,
        header: attr_f64(e, b"header")?,
        footer: attr_f64(e, b"footer")?,
    })
}

fn parse_header_footer_item_text(text: &str) -> HeaderFooterItemInfo {
    let mut item = HeaderFooterItemInfo::default();
    let mut current: Option<u8> = None;
    let mut buf = String::new();
    let mut chars = text.chars().peekable();

    while let Some(ch) = chars.next() {
        if ch == '&' {
            match chars.peek().copied() {
                Some('L') | Some('C') | Some('R') => {
                    assign_header_footer_segment(&mut item, current, std::mem::take(&mut buf));
                    current = chars.next().map(|marker| marker as u8);
                    continue;
                }
                _ => {}
            }
        }
        buf.push(ch);
    }
    assign_header_footer_segment(&mut item, current, buf);
    item
}

fn assign_header_footer_segment(
    item: &mut HeaderFooterItemInfo,
    segment: Option<u8>,
    value: String,
) {
    if value.is_empty() {
        return;
    }
    match segment.unwrap_or(b'C') {
        b'L' => item.left = Some(value),
        b'R' => item.right = Some(value),
        _ => item.center = Some(value),
    }
}

fn append_break(
    row_breaks: &mut Option<PageBreakListInfo>,
    column_breaks: &mut Option<PageBreakListInfo>,
    active: Option<BreakListKind>,
    e: &BytesStart<'_>,
) {
    let Some(active) = active else {
        return;
    };
    let Some(id) = attr_u32(e, b"id") else {
        return;
    };
    let break_info = BreakInfo {
        id,
        min: attr_u32(e, b"min"),
        max: attr_u32(e, b"max"),
        man: attr_bool_default(e, b"man", true),
        pt: attr_bool_default(e, b"pt", false),
    };
    match active {
        BreakListKind::Row => {
            row_breaks
                .get_or_insert_with(PageBreakListInfo::default)
                .breaks
                .push(break_info);
        }
        BreakListKind::Column => {
            column_breaks
                .get_or_insert_with(PageBreakListInfo::default)
                .breaks
                .push(break_info);
        }
    }
}

fn update_row_visibility(
    e: &BytesStart<'_>,
    row: Option<u32>,
    hidden_rows: &mut HashMap<u32, bool>,
    row_outline_levels: &mut HashMap<u32, u8>,
) {
    let Some(row) = row else {
        return;
    };
    hidden_rows.insert(row, attr_truthy(attr_value(e, b"hidden").as_deref()));
    let outline_level = attr_value(e, b"outlineLevel")
        .and_then(|value| value.parse::<u8>().ok())
        .unwrap_or(0);
    if outline_level > 0 {
        row_outline_levels.insert(row, outline_level);
    } else {
        row_outline_levels.remove(&row);
    }
}

fn update_column_visibility(
    e: &BytesStart<'_>,
    hidden_columns: &mut HashMap<u32, bool>,
    column_outline_levels: &mut HashMap<u32, u8>,
) {
    let min = attr_value(e, b"min")
        .and_then(|value| value.parse::<u32>().ok())
        .unwrap_or(1);
    let max = attr_value(e, b"max")
        .and_then(|value| value.parse::<u32>().ok())
        .unwrap_or(min);
    let hidden = attr_truthy(attr_value(e, b"hidden").as_deref());
    let outline_level = attr_value(e, b"outlineLevel")
        .and_then(|value| value.parse::<u8>().ok())
        .unwrap_or(0);
    for col in min..=max {
        hidden_columns.insert(col, hidden);
        if outline_level > 0 {
            column_outline_levels.insert(col, outline_level);
        } else {
            column_outline_levels.remove(&col);
        }
    }
}

fn build_array_formula_map(cells: &[Cell]) -> HashMap<(u32, u32), ArrayFormulaInfo> {
    let mut out = HashMap::new();
    for cell in cells {
        let Some(info) = &cell.array_formula else {
            continue;
        };
        out.insert((cell.row, cell.col), info.clone());
        match info {
            ArrayFormulaInfo::Array { ref_range, .. }
            | ArrayFormulaInfo::DataTable { ref_range, .. } => {
                mark_array_spill_children(&mut out, ref_range, (cell.row, cell.col));
            }
            ArrayFormulaInfo::SpillChild => {}
        }
    }
    out
}

#[derive(Debug, Clone)]
struct SharedFormulaMaster {
    row: u32,
    col: u32,
    formula: String,
}

fn resolve_shared_formulas(cells: &mut [Cell]) {
    let mut masters: HashMap<String, SharedFormulaMaster> = HashMap::new();
    for cell in cells.iter() {
        if cell.formula_kind.as_deref() != Some("shared") {
            continue;
        }
        let (Some(index), Some(formula)) = (&cell.formula_shared_index, &cell.formula) else {
            continue;
        };
        masters
            .entry(index.clone())
            .or_insert_with(|| SharedFormulaMaster {
                row: cell.row,
                col: cell.col,
                formula: formula.clone(),
            });
    }

    for cell in cells.iter_mut() {
        if cell.formula_kind.as_deref() != Some("shared") || cell.formula.is_some() {
            continue;
        }
        let Some(index) = &cell.formula_shared_index else {
            continue;
        };
        let Some(master) = masters.get(index) else {
            continue;
        };
        cell.formula = Some(translate_shared_formula(
            &master.formula,
            cell.row as i32 - master.row as i32,
            cell.col as i32 - master.col as i32,
        ));
    }
}

pub(crate) fn translate_shared_formula(formula: &str, rows: i32, cols: i32) -> String {
    if rows == 0 && cols == 0 {
        return formula.to_string();
    }
    let had_equals = formula.starts_with('=');
    let wrapped;
    let input = if had_equals {
        formula
    } else {
        wrapped = format!("={formula}");
        &wrapped
    };
    let mut delta = RefDelta::empty();
    delta.rows = rows;
    delta.cols = cols;
    delta.anchor_row = 1;
    delta.anchor_col = 1;
    delta.respect_dollar = true;
    let translated = translate_formula(input, &delta).unwrap_or_else(|_| input.to_string());
    if had_equals {
        translated
    } else {
        translated
            .strip_prefix('=')
            .unwrap_or(&translated)
            .to_string()
    }
}

fn mark_array_spill_children(
    out: &mut HashMap<(u32, u32), ArrayFormulaInfo>,
    ref_range: &str,
    master: (u32, u32),
) {
    let Some((min_row, min_col, max_row, max_col)) = parse_range_bounds_1based(ref_range) else {
        return;
    };
    for row in min_row..=max_row {
        for col in min_col..=max_col {
            if (row, col) != master {
                out.entry((row, col))
                    .or_insert(ArrayFormulaInfo::SpillChild);
            }
        }
    }
}

fn parse_range_bounds_1based(range: &str) -> Option<(u32, u32, u32, u32)> {
    let clean = range.replace('$', "").to_ascii_uppercase();
    let mut parts = clean.split(':');
    let start = parts.next()?;
    let end = parts.next().unwrap_or(start);
    let (start_row, start_col) = a1_to_row_col(start)?;
    let (end_row, end_col) = a1_to_row_col(end)?;
    Some((
        start_row.min(end_row),
        start_col.min(end_col),
        start_row.max(end_row),
        start_col.max(end_col),
    ))
}

fn parse_auto_filter(xml: &str) -> Result<Option<AutoFilterInfo>> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut auto_filter: Option<AutoFilterInfo> = None;
    let mut current_col: Option<FilterColumnInfo> = None;
    let mut filters: Option<FiltersBuilder> = None;
    let mut custom_filters: Option<CustomFiltersBuilder> = None;
    let mut sort_state: Option<SortStateInfo> = None;
    let mut in_auto_filter = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.local_name().as_ref() {
                b"autoFilter" => {
                    in_auto_filter = true;
                    auto_filter = attr_value(&e, b"ref")
                        .filter(|value| !value.trim().is_empty())
                        .map(|ref_range| AutoFilterInfo {
                            ref_range,
                            filter_columns: Vec::new(),
                            sort_state: None,
                        });
                }
                b"filterColumn" if in_auto_filter => {
                    current_col = Some(FilterColumnInfo::from_start(&e));
                }
                b"filters" if current_col.is_some() => {
                    filters = Some(FiltersBuilder::from_start(&e));
                }
                b"customFilters" if current_col.is_some() => {
                    custom_filters = Some(CustomFiltersBuilder::from_start(&e));
                }
                b"sortState" if in_auto_filter => {
                    sort_state = Some(SortStateInfo::from_start(&e));
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.local_name().as_ref() {
                b"autoFilter" => {
                    auto_filter = attr_value(&e, b"ref")
                        .filter(|value| !value.trim().is_empty())
                        .map(|ref_range| AutoFilterInfo {
                            ref_range,
                            filter_columns: Vec::new(),
                            sort_state: None,
                        });
                }
                b"filterColumn" if in_auto_filter => {
                    if let Some(auto_filter) = auto_filter.as_mut() {
                        auto_filter
                            .filter_columns
                            .push(FilterColumnInfo::from_start(&e));
                    }
                }
                b"filters" if current_col.is_some() => {
                    if let Some(col) = current_col.as_mut() {
                        let filter = FiltersBuilder::from_start(&e).finish();
                        if filter.is_some() {
                            col.filter = filter;
                        }
                    }
                }
                b"filter" if filters.is_some() => {
                    if let (Some(filters), Some(value)) = (filters.as_mut(), attr_value(&e, b"val"))
                    {
                        filters.values.push(value);
                    }
                }
                b"dateGroupItem" if current_col.is_some() => {
                    if let Some(item) = DateGroupItemInfo::from_start(&e) {
                        current_col.as_mut().unwrap().date_group_items.push(item);
                    }
                }
                b"customFilters" if current_col.is_some() => {
                    if let Some(col) = current_col.as_mut() {
                        col.filter = Some(CustomFiltersBuilder::from_start(&e).finish());
                    }
                }
                b"customFilter" if custom_filters.is_some() => {
                    if let Some(filter) = CustomFilterInfo::from_start(&e) {
                        custom_filters.as_mut().unwrap().filters.push(filter);
                    }
                }
                b"dynamicFilter" if current_col.is_some() => {
                    current_col.as_mut().unwrap().filter = Some(FilterInfo::Dynamic {
                        filter_type: attr_value(&e, b"type").unwrap_or_else(|| "null".to_string()),
                        val: attr_f64(&e, b"val"),
                        val_iso: attr_value(&e, b"valIso"),
                        max_val_iso: attr_value(&e, b"maxValIso"),
                    });
                }
                b"colorFilter" if current_col.is_some() => {
                    current_col.as_mut().unwrap().filter = Some(FilterInfo::Color {
                        dxf_id: attr_u32(&e, b"dxfId").unwrap_or(0),
                        cell_color: attr_bool_default(&e, b"cellColor", true),
                    });
                }
                b"iconFilter" if current_col.is_some() => {
                    current_col.as_mut().unwrap().filter = Some(FilterInfo::Icon {
                        icon_set: attr_value(&e, b"iconSet")
                            .unwrap_or_else(|| "3Arrows".to_string()),
                        icon_id: attr_u32(&e, b"iconId").unwrap_or(0),
                    });
                }
                b"top10" if current_col.is_some() => {
                    current_col.as_mut().unwrap().filter = Some(FilterInfo::Top10 {
                        top: attr_bool_default(&e, b"top", true),
                        percent: attr_bool_default(&e, b"percent", false),
                        val: attr_f64(&e, b"val").unwrap_or(10.0),
                        filter_val: attr_f64(&e, b"filterVal"),
                    });
                }
                b"sortCondition" if sort_state.is_some() => {
                    if let Some(condition) = SortConditionInfo::from_start(&e) {
                        sort_state.as_mut().unwrap().sort_conditions.push(condition);
                    }
                }
                _ => {}
            },
            Ok(Event::End(e)) => match e.local_name().as_ref() {
                b"autoFilter" => {
                    if let Some(pending_sort_state) = sort_state.take() {
                        if let Some(auto_filter) = auto_filter.as_mut() {
                            auto_filter.sort_state = Some(pending_sort_state);
                        }
                    }
                    break;
                }
                b"filterColumn" if in_auto_filter => {
                    if let (Some(auto_filter), Some(col)) =
                        (auto_filter.as_mut(), current_col.take())
                    {
                        auto_filter.filter_columns.push(col);
                    }
                }
                b"filters" if current_col.is_some() => {
                    if let (Some(col), Some(filters)) = (current_col.as_mut(), filters.take()) {
                        if let Some(filter) = filters.finish() {
                            col.filter = Some(filter);
                        }
                    }
                }
                b"customFilters" if current_col.is_some() => {
                    if let (Some(col), Some(filters)) =
                        (current_col.as_mut(), custom_filters.take())
                    {
                        col.filter = Some(filters.finish());
                    }
                }
                b"sortState" if auto_filter.is_some() => {
                    if let Some(auto_filter) = auto_filter.as_mut() {
                        auto_filter.sort_state = sort_state.take();
                    }
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(ReaderError::Xml(format!("failed to parse autoFilter: {e}"))),
            _ => {}
        }
        buf.clear();
    }

    Ok(auto_filter)
}

#[derive(Debug)]
struct FiltersBuilder {
    blank: bool,
    values: Vec<String>,
}

impl FiltersBuilder {
    fn from_start(e: &BytesStart<'_>) -> Self {
        Self {
            blank: attr_bool_default(e, b"blank", false),
            values: Vec::new(),
        }
    }

    fn finish(self) -> Option<FilterInfo> {
        if !self.values.is_empty() {
            Some(FilterInfo::String {
                values: self.values,
            })
        } else if self.blank {
            Some(FilterInfo::Blank)
        } else {
            None
        }
    }
}

#[derive(Debug)]
struct CustomFiltersBuilder {
    and_: bool,
    filters: Vec<CustomFilterInfo>,
}

impl CustomFiltersBuilder {
    fn from_start(e: &BytesStart<'_>) -> Self {
        Self {
            and_: attr_bool_default(e, b"and", false),
            filters: Vec::new(),
        }
    }

    fn finish(self) -> FilterInfo {
        FilterInfo::Custom {
            and_: self.and_,
            filters: self.filters,
        }
    }
}

impl FilterColumnInfo {
    fn from_start(e: &BytesStart<'_>) -> Self {
        Self {
            col_id: attr_u32(e, b"colId").unwrap_or(0),
            hidden_button: attr_bool_default(e, b"hiddenButton", false),
            show_button: attr_bool_default(e, b"showButton", true),
            filter: None,
            date_group_items: Vec::new(),
        }
    }
}

impl CustomFilterInfo {
    fn from_start(e: &BytesStart<'_>) -> Option<Self> {
        Some(Self {
            operator: attr_value(e, b"operator").unwrap_or_else(|| "equal".to_string()),
            val: attr_value(e, b"val")?,
        })
    }
}

impl DateGroupItemInfo {
    fn from_start(e: &BytesStart<'_>) -> Option<Self> {
        Some(Self {
            year: attr_u32(e, b"year")?,
            month: attr_u32(e, b"month"),
            day: attr_u32(e, b"day"),
            hour: attr_u32(e, b"hour"),
            minute: attr_u32(e, b"minute"),
            second: attr_u32(e, b"second"),
            date_time_grouping: attr_value(e, b"dateTimeGrouping")
                .unwrap_or_else(|| "year".to_string()),
        })
    }
}

impl SortStateInfo {
    fn from_start(e: &BytesStart<'_>) -> Self {
        Self {
            sort_conditions: Vec::new(),
            column_sort: attr_bool_default(e, b"columnSort", false),
            case_sensitive: attr_bool_default(e, b"caseSensitive", false),
            ref_range: attr_value(e, b"ref"),
        }
    }
}

impl SortConditionInfo {
    fn from_start(e: &BytesStart<'_>) -> Option<Self> {
        Some(Self {
            ref_range: attr_value(e, b"ref")?,
            descending: attr_bool_default(e, b"descending", false),
            sort_by: attr_value(e, b"sortBy").unwrap_or_else(|| "value".to_string()),
            custom_list: attr_value(e, b"customList"),
            dxf_id: attr_u32(e, b"dxfId"),
            icon_set: attr_value(e, b"iconSet"),
            icon_id: attr_u32(e, b"iconId"),
        })
    }
}

impl SheetProtection {
    fn from_start(e: &BytesStart<'_>) -> Self {
        Self {
            sheet: attr_bool_default(e, b"sheet", false),
            objects: attr_bool_default(e, b"objects", false),
            scenarios: attr_bool_default(e, b"scenarios", false),
            format_cells: attr_bool_default(e, b"formatCells", true),
            format_columns: attr_bool_default(e, b"formatColumns", true),
            format_rows: attr_bool_default(e, b"formatRows", true),
            insert_columns: attr_bool_default(e, b"insertColumns", true),
            insert_rows: attr_bool_default(e, b"insertRows", true),
            insert_hyperlinks: attr_bool_default(e, b"insertHyperlinks", true),
            delete_columns: attr_bool_default(e, b"deleteColumns", true),
            delete_rows: attr_bool_default(e, b"deleteRows", true),
            select_locked_cells: attr_bool_default(e, b"selectLockedCells", false),
            sort: attr_bool_default(e, b"sort", true),
            auto_filter: attr_bool_default(e, b"autoFilter", true),
            pivot_tables: attr_bool_default(e, b"pivotTables", true),
            select_unlocked_cells: attr_bool_default(e, b"selectUnlockedCells", false),
            password_hash: attr_value(e, b"password"),
        }
    }
}

#[derive(Debug)]
struct DataValidationBuilder {
    range: String,
    validation_type: String,
    operator: Option<String>,
    formula1: Option<String>,
    formula2: Option<String>,
    formula1_text: String,
    formula2_text: String,
    allow_blank: bool,
    show_input_message: bool,
    show_error_message: bool,
    error_title: Option<String>,
    error: Option<String>,
    prompt_title: Option<String>,
    prompt: Option<String>,
}

impl DataValidationBuilder {
    fn from_start(e: &BytesStart<'_>) -> Self {
        Self {
            range: attr_value(e, b"sqref").unwrap_or_default(),
            validation_type: attr_value(e, b"type").unwrap_or_else(|| "any".to_string()),
            operator: attr_value(e, b"operator"),
            formula1: None,
            formula2: None,
            formula1_text: String::new(),
            formula2_text: String::new(),
            allow_blank: attr_value(e, b"allowBlank")
                .map(|value| attr_truthy(Some(&value)))
                .unwrap_or(true),
            show_input_message: attr_value(e, b"showInputMessage")
                .map(|value| attr_truthy(Some(&value)))
                .unwrap_or(false),
            show_error_message: attr_value(e, b"showErrorMessage")
                .map(|value| attr_truthy(Some(&value)))
                .unwrap_or(false),
            error_title: attr_value(e, b"errorTitle").filter(|value| !value.is_empty()),
            error: attr_value(e, b"error").filter(|value| !value.is_empty()),
            prompt_title: attr_value(e, b"promptTitle").filter(|value| !value.is_empty()),
            prompt: attr_value(e, b"prompt").filter(|value| !value.is_empty()),
        }
    }

    fn push_text(&mut self, target: DataValidationFormula, text: &str) {
        match target {
            DataValidationFormula::Formula1 => self.formula1_text.push_str(text),
            DataValidationFormula::Formula2 => self.formula2_text.push_str(text),
        }
    }

    fn finish_formula1(&mut self) {
        let formula = self.formula1_text.trim();
        if !formula.is_empty() {
            self.formula1 = Some(formula.to_string());
        }
    }

    fn finish_formula2(&mut self) {
        let formula = self.formula2_text.trim();
        if !formula.is_empty() {
            self.formula2 = Some(formula.to_string());
        }
    }

    fn finish(self) -> DataValidation {
        DataValidation {
            range: self.range,
            validation_type: self.validation_type,
            operator: self.operator,
            formula1: self.formula1,
            formula2: self.formula2,
            allow_blank: self.allow_blank,
            show_input_message: self.show_input_message,
            show_error_message: self.show_error_message,
            error_title: self.error_title,
            error: self.error,
            prompt_title: self.prompt_title,
            prompt: self.prompt,
        }
    }
}

#[derive(Debug, Clone, Copy)]
enum DataValidationFormula {
    Formula1,
    Formula2,
}

#[derive(Debug)]
struct ConditionalFormatBuilder {
    range: String,
    rule_type: String,
    operator: Option<String>,
    formula: Option<String>,
    formula_text: String,
    dxf_id: Option<u32>,
    priority: Option<i64>,
    stop_if_true: Option<bool>,
    color_scale: Option<ColorScaleInfo>,
}

impl ConditionalFormatBuilder {
    fn from_start(e: &BytesStart<'_>, range: String) -> Self {
        Self {
            range,
            rule_type: attr_value(e, b"type").unwrap_or_default(),
            operator: attr_value(e, b"operator"),
            formula: None,
            formula_text: String::new(),
            dxf_id: attr_u32(e, b"dxfId"),
            priority: attr_value(e, b"priority").and_then(|value| value.parse::<i64>().ok()),
            stop_if_true: attr_value(e, b"stopIfTrue").map(|value| attr_truthy(Some(&value))),
            color_scale: None,
        }
    }

    fn push_formula_text(&mut self, text: &str) {
        self.formula_text.push_str(text);
    }

    fn finish_formula(&mut self) {
        let formula = self.formula_text.trim();
        if !formula.is_empty() && self.formula.is_none() {
            self.formula = Some(formula.to_string());
        }
        self.formula_text.clear();
    }

    fn start_color_scale(&mut self) {
        if self.color_scale.is_none() {
            self.color_scale = Some(ColorScaleInfo::default());
        }
    }

    fn push_cfvo(&mut self, e: &BytesStart<'_>) {
        if let Some(cs) = self.color_scale.as_mut() {
            cs.cfvo.push(CfvoInfo {
                cfvo_type: attr_value(e, b"type").unwrap_or_default(),
                val: attr_value(e, b"val"),
            });
        }
    }

    fn push_color(&mut self, e: &BytesStart<'_>) {
        if let Some(cs) = self.color_scale.as_mut() {
            // Prefer rgb attribute; theme/indexed colors are out of scope
            // here (round-tripped only when explicit ARGB is present).
            if let Some(rgb) = attr_value(e, b"rgb") {
                cs.colors.push(rgb);
            }
        }
    }

    fn finish(self) -> ConditionalFormatRule {
        ConditionalFormatRule {
            range: self.range,
            rule_type: self.rule_type,
            operator: self.operator,
            formula: self.formula,
            dxf_id: self.dxf_id,
            priority: self.priority,
            stop_if_true: self.stop_if_true,
            color_scale: self.color_scale,
        }
    }
}

fn parse_pane(e: &BytesStart<'_>) -> Option<FreezePane> {
    let mode = match attr_value(e, b"state")?.to_ascii_lowercase().as_str() {
        "split" => PaneMode::Split,
        state if state.starts_with("frozen") => PaneMode::Freeze,
        _ => return None,
    };
    Some(FreezePane {
        mode,
        top_left_cell: attr_value(e, b"topLeftCell").filter(|value| !value.is_empty()),
        x_split: attr_value(e, b"xSplit")
            .and_then(|value| value.parse::<f64>().ok())
            .map(|value| value as i64),
        y_split: attr_value(e, b"ySplit")
            .and_then(|value| value.parse::<f64>().ok())
            .map(|value| value as i64),
        active_pane: attr_value(e, b"activePane").filter(|value| !value.is_empty()),
    })
}

#[derive(Debug)]
struct HyperlinkNode {
    cell: String,
    rid: Option<String>,
    location: Option<String>,
    display: Option<String>,
    tooltip: Option<String>,
}

impl HyperlinkNode {
    fn from_start(e: &BytesStart<'_>) -> Option<Self> {
        let cell = attr_value(e, b"ref").filter(|value| !value.is_empty())?;
        Some(Self {
            cell,
            rid: attr_value(e, b"r:id"),
            location: attr_value(e, b"location"),
            display: attr_value(e, b"display"),
            tooltip: attr_value(e, b"tooltip"),
        })
    }
}

fn resolve_hyperlinks(nodes: Vec<HyperlinkNode>, rels: Option<&RelsGraph>) -> Vec<Hyperlink> {
    let mut out = Vec::new();
    for node in nodes {
        let internal = node.location.is_some() && node.rid.is_none();
        let target = if let Some(rid) = &node.rid {
            rels.and_then(|rels| rels.get(&RelId(rid.clone())))
                .map(|rel| rel.target.clone())
                .unwrap_or_default()
        } else {
            node.location.clone().unwrap_or_default()
        };
        if target.is_empty() {
            continue;
        }
        let display = node
            .display
            .filter(|value| !value.is_empty())
            .unwrap_or_default();
        let tooltip = node
            .tooltip
            .and_then(|t| if t.is_empty() { None } else { Some(t) });
        out.push(Hyperlink {
            cell: node.cell,
            target,
            location: node.location,
            display,
            tooltip,
            internal,
        });
    }
    out
}

fn comments_target(rels: &RelsGraph) -> Option<String> {
    rels.iter()
        .find(|rel| rel.rel_type.ends_with("/comments") || rel.rel_type == "comments")
        .map(|rel| rel.target.clone())
}

fn person_list_target(rels: &RelsGraph) -> Option<String> {
    rels.iter()
        .find(|rel| rel.rel_type == wolfxl_rels::rt::PERSON_LIST)
        .map(|rel| rel.target.clone())
}

/// Parse `xl/persons/personList.xml` (RFC-068).
///
/// Schema is flat: a `<personList>` root with `<person displayName id userId
/// providerId/>` children. Insertion order is preserved so the registry
/// round-trips byte-stable through wolfxl's writer.
fn parse_person_list(xml: &str) -> Result<Vec<Person>> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut persons = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.local_name().as_ref() == b"person" {
                    let display_name = attr_value(&e, b"displayName").unwrap_or_default();
                    let id = attr_value(&e, b"id").unwrap_or_default();
                    if id.is_empty() {
                        // Skip malformed entries — every Person must carry a GUID
                        // for the threaded-comment cross-link to resolve.
                        continue;
                    }
                    persons.push(Person {
                        display_name,
                        id,
                        user_id: attr_value(&e, b"userId"),
                        provider_id: attr_value(&e, b"providerId"),
                    });
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ReaderError::Xml(format!(
                    "failed to parse personList XML: {e}"
                )));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(persons)
}

/// Parse one `xl/threadedComments/threadedCommentsN.xml` part.
///
/// Top-level threads have no `parentId`; replies carry the parent's GUID.
/// Caller reassembles into a tree by matching `parent_id` against `id`.
fn parse_threaded_comments(xml: &str) -> Result<Vec<ParsedThreadedComment>> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut out: Vec<ParsedThreadedComment> = Vec::new();

    let mut current: Option<ParsedThreadedComment> = None;
    let mut in_text = false;
    let mut text_buf = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.local_name().as_ref() {
                b"threadedComment" => {
                    let id = attr_value(&e, b"id").unwrap_or_default();
                    let cell = attr_value(&e, b"ref").unwrap_or_default();
                    let person_id = attr_value(&e, b"personId").unwrap_or_default();
                    let created = attr_value(&e, b"dT");
                    let parent_id = attr_value(&e, b"parentId");
                    let done = attr_value(&e, b"done")
                        .map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
                        .unwrap_or(false);
                    current = Some(ParsedThreadedComment {
                        id,
                        cell,
                        person_id,
                        created,
                        text: String::new(),
                        parent_id,
                        done,
                    });
                    text_buf.clear();
                }
                b"text" => {
                    in_text = true;
                    text_buf.clear();
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => {
                // Self-closing `<threadedComment ... />` — empty body is valid.
                if e.local_name().as_ref() == b"threadedComment" {
                    let id = attr_value(&e, b"id").unwrap_or_default();
                    let cell = attr_value(&e, b"ref").unwrap_or_default();
                    let person_id = attr_value(&e, b"personId").unwrap_or_default();
                    let created = attr_value(&e, b"dT");
                    let parent_id = attr_value(&e, b"parentId");
                    let done = attr_value(&e, b"done")
                        .map(|v| v == "1" || v.eq_ignore_ascii_case("true"))
                        .unwrap_or(false);
                    out.push(ParsedThreadedComment {
                        id,
                        cell,
                        person_id,
                        created,
                        text: String::new(),
                        parent_id,
                        done,
                    });
                }
            }
            Ok(Event::End(e)) => match e.local_name().as_ref() {
                b"threadedComment" => {
                    if let Some(mut tc) = current.take() {
                        tc.text = std::mem::take(&mut text_buf);
                        out.push(tc);
                    }
                }
                b"text" => {
                    in_text = false;
                }
                _ => {}
            },
            Ok(Event::Text(e)) => {
                if in_text {
                    let s = e
                        .unescape()
                        .map_err(|err| ReaderError::Xml(format!("threadedComments text: {err}")))?;
                    text_buf.push_str(&s);
                }
            }
            Ok(Event::CData(e)) => {
                if in_text {
                    text_buf.push_str(&String::from_utf8_lossy(e.as_ref()));
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ReaderError::Xml(format!(
                    "failed to parse threadedComments XML: {e}"
                )));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

fn parse_comments(xml: &str) -> Result<Vec<Comment>> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut authors = Vec::new();
    let mut comments = Vec::new();
    let mut in_author = false;
    let mut in_comment = false;
    let mut in_t = false;
    let mut current_cell = String::new();
    let mut current_author_id = 0usize;
    let mut current_text = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.local_name().as_ref() {
                b"author" => in_author = true,
                b"comment" => {
                    in_comment = true;
                    current_text.clear();
                    current_cell = attr_value(&e, b"ref").unwrap_or_default();
                    current_author_id = attr_value(&e, b"authorId")
                        .and_then(|value| value.parse::<usize>().ok())
                        .unwrap_or(0);
                }
                b"t" => in_t = true,
                _ => {}
            },
            Ok(Event::End(e)) => match e.local_name().as_ref() {
                b"author" => in_author = false,
                b"comment" => {
                    in_comment = false;
                    comments.push(Comment {
                        cell: current_cell.clone(),
                        text: current_text.clone(),
                        author: authors.get(current_author_id).cloned().unwrap_or_default(),
                        threaded: false,
                    });
                }
                b"t" => in_t = false,
                _ => {}
            },
            Ok(Event::Text(e)) => {
                let text = e
                    .unescape()
                    .map_err(|err| ReaderError::Xml(format!("comments text: {err}")))?
                    .to_string();
                if in_author {
                    authors.push(text);
                } else if in_comment && in_t {
                    current_text.push_str(&text);
                }
            }
            Ok(Event::CData(e)) => {
                if in_comment && in_t {
                    current_text.push_str(&String::from_utf8_lossy(e.as_ref()));
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ReaderError::Xml(format!(
                    "failed to parse comments XML: {e}"
                )));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(comments)
}

fn parse_doc_properties_into(
    xml: &str,
    out: &mut HashMap<String, String>,
    key_for_tag: fn(&[u8]) -> Option<&'static str>,
) {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut current_tag: Option<Vec<u8>> = None;
    let mut current_text = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                current_tag = Some(e.local_name().as_ref().to_vec());
                current_text.clear();
            }
            Ok(Event::Text(e)) => {
                if current_tag.is_some() {
                    current_text.push_str(&e.unescape().unwrap_or_default());
                }
            }
            Ok(Event::End(e)) => {
                let name = e.local_name();
                let name = name.as_ref();
                if current_tag.as_deref() == Some(name) {
                    if let Some(key) = key_for_tag(name) {
                        let value = current_text.trim();
                        if !value.is_empty() {
                            out.entry(key.to_string())
                                .or_insert_with(|| value.to_string());
                        }
                    }
                }
                current_tag = None;
                current_text.clear();
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
}

fn doc_property_core_key(name: &[u8]) -> Option<&'static str> {
    match name {
        b"title" => Some("title"),
        b"subject" => Some("subject"),
        b"creator" => Some("creator"),
        b"keywords" => Some("keywords"),
        b"description" => Some("description"),
        b"lastModifiedBy" => Some("lastModifiedBy"),
        b"category" => Some("category"),
        b"contentStatus" => Some("contentStatus"),
        b"identifier" => Some("identifier"),
        b"language" => Some("language"),
        b"revision" => Some("revision"),
        b"version" => Some("version"),
        b"created" => Some("created"),
        b"modified" => Some("modified"),
        _ => None,
    }
}

fn doc_property_app_key(name: &[u8]) -> Option<&'static str> {
    match name {
        b"Company" => Some("company"),
        b"Manager" => Some("manager"),
        b"Application" => Some("application"),
        _ => None,
    }
}

fn parse_custom_doc_properties(xml: &str) -> Result<Vec<CustomPropertyInfo>> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut properties = Vec::new();
    let mut current_name: Option<String> = None;
    let mut current_link_target: Option<String> = None;
    let mut current_kind: Option<String> = None;
    let mut current_value = String::new();
    let mut active_value_tag: Option<Vec<u8>> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.local_name().as_ref() {
                b"property" => {
                    current_name = attr_value(&e, b"name");
                    current_link_target = attr_value(&e, b"linkTarget");
                    current_kind = current_link_target.as_ref().map(|_| "link".to_string());
                    current_value.clear();
                }
                tag if current_name.is_some() => {
                    active_value_tag = Some(tag.to_vec());
                    if current_kind.is_none() {
                        current_kind = Some(custom_property_kind(tag).to_string());
                    }
                    current_value.clear();
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.local_name().as_ref() {
                b"property" => {
                    if let Some(name) = attr_value(&e, b"name") {
                        if let Some(link) = attr_value(&e, b"linkTarget") {
                            properties.push(CustomPropertyInfo {
                                name,
                                kind: "link".to_string(),
                                value: link,
                            });
                        }
                    }
                }
                tag if current_name.is_some() => {
                    if current_kind.is_none() {
                        current_kind = Some(custom_property_kind(tag).to_string());
                    }
                }
                _ => {}
            },
            Ok(Event::Text(e)) => {
                if active_value_tag.is_some() {
                    current_value.push_str(
                        &e.unescape().map_err(|err| {
                            ReaderError::Xml(format!("custom property text: {err}"))
                        })?,
                    );
                }
            }
            Ok(Event::CData(e)) => {
                if active_value_tag.is_some() {
                    current_value.push_str(&String::from_utf8_lossy(e.as_ref()));
                }
            }
            Ok(Event::End(e)) => match e.local_name().as_ref() {
                b"property" => {
                    if let Some(name) = current_name.take() {
                        let kind = current_kind.take().unwrap_or_else(|| "string".to_string());
                        let value = current_link_target
                            .take()
                            .unwrap_or_else(|| current_value.clone());
                        properties.push(CustomPropertyInfo { name, kind, value });
                    }
                    active_value_tag = None;
                    current_value.clear();
                }
                tag => {
                    if active_value_tag.as_deref() == Some(tag) {
                        active_value_tag = None;
                    }
                }
            },
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ReaderError::Xml(format!(
                    "failed to parse docProps/custom.xml: {e}"
                )));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(properties)
}

fn custom_property_kind(tag: &[u8]) -> &'static str {
    match tag {
        b"i1" | b"i2" | b"i4" | b"i8" | b"int" | b"uint" | b"ui1" | b"ui2" | b"ui4" | b"ui8" => {
            "int"
        }
        b"r4" | b"r8" | b"decimal" => "float",
        b"bool" => "bool",
        b"filetime" | b"date" => "datetime",
        _ => "string",
    }
}

fn read_tables<R: Read + std::io::Seek>(
    zip: &mut ZipArchive<R>,
    sheet_path: &str,
    sheet_xml: &str,
    rels: Option<&RelsGraph>,
) -> Result<Vec<Table>> {
    let Some(rels) = rels else {
        return Ok(Vec::new());
    };
    let table_rids = parse_table_rids(sheet_xml)?;
    if table_rids.is_empty() {
        return Ok(Vec::new());
    }

    let sheet_dir = part_dir(sheet_path);
    let mut out = Vec::new();
    for rid in table_rids {
        let Some(rel) = rels.get(&RelId(rid)) else {
            continue;
        };
        let table_path = join_and_normalize(&sheet_dir, &rel.target);
        let Some(table_xml) = read_part_optional(zip, &table_path)? else {
            continue;
        };
        if let Ok(table) = parse_table_xml(&table_xml) {
            out.push(table);
        }
    }
    Ok(out)
}

pub(crate) fn read_images<R: Read + std::io::Seek>(
    zip: &mut ZipArchive<R>,
    sheet_path: &str,
    rels: Option<&RelsGraph>,
) -> Result<Vec<ImageInfo>> {
    let Some(rels) = rels else {
        return Ok(Vec::new());
    };
    let sheet_dir = part_dir(sheet_path);
    let mut out = Vec::new();

    for drawing_rel in rels.find_by_type(wolfxl_rels::rt::DRAWING) {
        let drawing_path = join_and_normalize(&sheet_dir, &drawing_rel.target);
        let Some(drawing_xml) = read_part_optional(zip, &drawing_path)? else {
            continue;
        };
        let drawing_rels_path = sheet_rels_path(&drawing_path);
        let image_rels = read_part_optional(zip, &drawing_rels_path)?
            .map(|xml| {
                RelsGraph::parse(xml.as_bytes())
                    .map_err(|e| ReaderError::Xml(format!("failed to parse drawing rels: {e}")))
            })
            .transpose()?;
        let Some(image_rels) = image_rels else {
            continue;
        };
        let image_targets: HashMap<String, String> = image_rels
            .iter()
            .filter(|rel| rel.rel_type == wolfxl_rels::rt::IMAGE)
            .map(|rel| (rel.id.0.clone(), rel.target.clone()))
            .collect();
        let drawing_dir = part_dir(&drawing_path);

        for image_ref in parse_drawing_images(&drawing_xml)? {
            let Some(target) = image_targets.get(&image_ref.rid) else {
                continue;
            };
            let image_path = join_and_normalize(&drawing_dir, target);
            let Some(data) = read_part_optional_bytes(zip, &image_path)? else {
                continue;
            };
            out.push(ImageInfo {
                data,
                ext: image_ext_from_path(&image_path),
                anchor: image_ref.anchor,
            });
        }
    }

    Ok(out)
}

pub(crate) fn read_charts<R: Read + std::io::Seek>(
    zip: &mut ZipArchive<R>,
    sheet_path: &str,
    rels: Option<&RelsGraph>,
) -> Result<Vec<ChartInfo>> {
    let Some(rels) = rels else {
        return Ok(Vec::new());
    };
    let sheet_dir = part_dir(sheet_path);
    let mut out = Vec::new();

    for drawing_rel in rels.find_by_type(wolfxl_rels::rt::DRAWING) {
        let drawing_path = join_and_normalize(&sheet_dir, &drawing_rel.target);
        let Some(drawing_xml) = read_part_optional(zip, &drawing_path)? else {
            continue;
        };
        let drawing_rels_path = sheet_rels_path(&drawing_path);
        let chart_rels = read_part_optional(zip, &drawing_rels_path)?
            .map(|xml| {
                RelsGraph::parse(xml.as_bytes())
                    .map_err(|e| ReaderError::Xml(format!("failed to parse drawing rels: {e}")))
            })
            .transpose()?;
        let Some(chart_rels) = chart_rels else {
            continue;
        };
        let chart_targets: HashMap<String, String> = chart_rels
            .iter()
            .filter(|rel| rel.rel_type == wolfxl_rels::rt::CHART)
            .map(|rel| (rel.id.0.clone(), rel.target.clone()))
            .collect();
        let drawing_dir = part_dir(&drawing_path);

        for chart_ref in parse_drawing_charts(&drawing_xml)? {
            let Some(target) = chart_targets.get(&chart_ref.rid) else {
                continue;
            };
            let chart_path = join_and_normalize(&drawing_dir, target);
            let Some(chart_xml) = read_part_optional(zip, &chart_path)? else {
                continue;
            };
            if let Some(mut chart) = parse_chart_xml(&chart_xml, chart_ref.anchor)? {
                chart.source_drawing_path = Some(drawing_path.clone());
                chart.source_chart_rid = Some(chart_ref.rid);
                chart.source_chart_path = Some(chart_path);
                out.push(chart);
            }
        }
    }

    Ok(out)
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct DrawingObjectRef {
    rid: String,
    anchor: ImageAnchorInfo,
}

type DrawingImageRef = DrawingObjectRef;
type DrawingChartRef = DrawingObjectRef;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum DrawingAnchorKind {
    OneCell,
    TwoCell,
    Absolute,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct DrawingObjectBuilder {
    kind: DrawingAnchorKind,
    from: AnchorMarkerInfo,
    to: AnchorMarkerInfo,
    pos: AnchorPositionInfo,
    ext: Option<AnchorExtentInfo>,
    edit_as: Option<String>,
    rid: Option<String>,
}

impl DrawingObjectBuilder {
    fn new(kind: DrawingAnchorKind, e: &BytesStart<'_>) -> Self {
        Self {
            kind,
            from: AnchorMarkerInfo::default(),
            to: AnchorMarkerInfo::default(),
            pos: AnchorPositionInfo::default(),
            ext: None,
            edit_as: attr_value(e, b"editAs"),
            rid: None,
        }
    }

    fn finish(self) -> Option<DrawingObjectRef> {
        let rid = self.rid?;
        let anchor = match self.kind {
            DrawingAnchorKind::OneCell => ImageAnchorInfo::OneCell {
                from: self.from,
                ext: self.ext,
            },
            DrawingAnchorKind::TwoCell => ImageAnchorInfo::TwoCell {
                from: self.from,
                to: self.to,
                edit_as: self.edit_as.unwrap_or_else(|| "oneCell".to_string()),
            },
            DrawingAnchorKind::Absolute => ImageAnchorInfo::Absolute {
                pos: self.pos,
                ext: self.ext.unwrap_or_default(),
            },
        };
        Some(DrawingObjectRef { rid, anchor })
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum MarkerSlot {
    From,
    To,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum MarkerTextTarget {
    Col,
    Row,
    ColOff,
    RowOff,
}

fn parse_drawing_images(xml: &str) -> Result<Vec<DrawingImageRef>> {
    parse_drawing_objects(xml, DrawingObjectKind::Image)
}

fn parse_drawing_charts(xml: &str) -> Result<Vec<DrawingChartRef>> {
    parse_drawing_objects(xml, DrawingObjectKind::Chart)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum DrawingObjectKind {
    Image,
    Chart,
}

fn parse_drawing_objects(
    xml: &str,
    object_kind: DrawingObjectKind,
) -> Result<Vec<DrawingObjectRef>> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut out = Vec::new();
    let mut current: Option<DrawingObjectBuilder> = None;
    let mut marker_slot: Option<MarkerSlot> = None;
    let mut marker_text: Option<MarkerTextTarget> = None;
    let mut in_target_frame = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.local_name().as_ref() {
                b"oneCellAnchor" => {
                    current = Some(DrawingObjectBuilder::new(DrawingAnchorKind::OneCell, &e));
                }
                b"twoCellAnchor" => {
                    current = Some(DrawingObjectBuilder::new(DrawingAnchorKind::TwoCell, &e));
                }
                b"absoluteAnchor" => {
                    current = Some(DrawingObjectBuilder::new(DrawingAnchorKind::Absolute, &e));
                }
                b"from" if current.is_some() => marker_slot = Some(MarkerSlot::From),
                b"to" if current.is_some() => marker_slot = Some(MarkerSlot::To),
                b"col" if marker_slot.is_some() => marker_text = Some(MarkerTextTarget::Col),
                b"row" if marker_slot.is_some() => marker_text = Some(MarkerTextTarget::Row),
                b"colOff" if marker_slot.is_some() => {
                    marker_text = Some(MarkerTextTarget::ColOff);
                }
                b"rowOff" if marker_slot.is_some() => {
                    marker_text = Some(MarkerTextTarget::RowOff);
                }
                b"pic" if object_kind == DrawingObjectKind::Image && current.is_some() => {
                    in_target_frame = true;
                }
                b"graphicFrame" if object_kind == DrawingObjectKind::Chart && current.is_some() => {
                    in_target_frame = true;
                }
                b"pos" => apply_anchor_pos(&mut current, &e),
                b"ext" if !in_target_frame => apply_anchor_ext(&mut current, &e),
                b"blip" if object_kind == DrawingObjectKind::Image => {
                    apply_blip_rid(&mut current, &e);
                }
                b"chart" if object_kind == DrawingObjectKind::Chart => {
                    apply_chart_rid(&mut current, &e);
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.local_name().as_ref() {
                b"pos" => apply_anchor_pos(&mut current, &e),
                b"ext" if !in_target_frame => apply_anchor_ext(&mut current, &e),
                b"blip" if object_kind == DrawingObjectKind::Image => {
                    apply_blip_rid(&mut current, &e);
                }
                b"chart" if object_kind == DrawingObjectKind::Chart => {
                    apply_chart_rid(&mut current, &e);
                }
                _ => {}
            },
            Ok(Event::Text(e)) => {
                if let (Some(slot), Some(target), Some(builder)) =
                    (marker_slot, marker_text, current.as_mut())
                {
                    let text = e
                        .unescape()
                        .map_err(|err| ReaderError::Xml(format!("drawing text: {err}")))?;
                    if let Ok(value) = text.parse::<i64>() {
                        apply_marker_value(builder, slot, target, value);
                    }
                }
            }
            Ok(Event::End(e)) => match e.local_name().as_ref() {
                b"col" | b"row" | b"colOff" | b"rowOff" => marker_text = None,
                b"from" | b"to" => {
                    marker_slot = None;
                    marker_text = None;
                }
                b"pic" if object_kind == DrawingObjectKind::Image => in_target_frame = false,
                b"graphicFrame" if object_kind == DrawingObjectKind::Chart => {
                    in_target_frame = false;
                }
                b"oneCellAnchor" | b"twoCellAnchor" | b"absoluteAnchor" => {
                    marker_slot = None;
                    marker_text = None;
                    in_target_frame = false;
                    if let Some(builder) = current.take().and_then(DrawingObjectBuilder::finish) {
                        out.push(builder);
                    }
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(ReaderError::Xml(format!("failed to parse drawing: {e}"))),
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

fn apply_anchor_pos(builder: &mut Option<DrawingObjectBuilder>, e: &BytesStart<'_>) {
    let Some(builder) = builder.as_mut() else {
        return;
    };
    builder.pos = AnchorPositionInfo {
        x: attr_i64(e, b"x").unwrap_or_default(),
        y: attr_i64(e, b"y").unwrap_or_default(),
    };
}

fn apply_chart_rid(builder: &mut Option<DrawingObjectBuilder>, e: &BytesStart<'_>) {
    let Some(builder) = builder.as_mut() else {
        return;
    };
    builder.rid = attr_value(e, b"r:id").or_else(|| attr_value(e, b"id"));
}

fn parse_chart_xml(xml: &str, anchor: ImageAnchorInfo) -> Result<Option<ChartInfo>> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut stack: Vec<Vec<u8>> = Vec::new();
    let mut kind: Option<String> = None;
    let mut title_parts: Vec<String> = Vec::new();
    let mut current_axis: Option<Vec<u8>> = None;
    let mut current_axis_info: Option<ChartAxisInfo> = None;
    let mut current_axis_title_parts: Vec<String> = Vec::new();
    let mut x_axis_title: Option<String> = None;
    let mut y_axis_title: Option<String> = None;
    let mut x_axis: Option<ChartAxisInfo> = None;
    let mut y_axis: Option<ChartAxisInfo> = None;
    let mut data_labels: Option<ChartDataLabelsInfo> = None;
    let mut val_axis_titles_seen = 0usize;
    let mut legend_position: Option<String> = None;
    let mut bar_dir: Option<String> = None;
    let mut grouping: Option<String> = None;
    let mut scatter_style: Option<String> = None;
    let mut vary_colors: Option<bool> = None;
    let mut style: Option<u32> = None;
    let mut current_series: Option<ChartSeriesInfo> = None;
    let mut series = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let local = e.local_name().as_ref().to_vec();
                apply_chart_start(
                    &stack,
                    &local,
                    &e,
                    &mut kind,
                    &mut legend_position,
                    &mut bar_dir,
                    &mut grouping,
                    &mut scatter_style,
                    &mut vary_colors,
                    &mut data_labels,
                    &mut style,
                    &mut current_series,
                );
                if chart_axis_name(&local).is_some() {
                    current_axis = Some(local.clone());
                    current_axis_info = Some(new_chart_axis_info(&local, &e));
                    current_axis_title_parts.clear();
                } else if let Some(axis) = current_axis_info.as_mut() {
                    apply_chart_axis_start(axis, &local, &e);
                }
                stack.push(local);
            }
            Ok(Event::Empty(e)) => {
                let local = e.local_name().as_ref().to_vec();
                apply_chart_start(
                    &stack,
                    &local,
                    &e,
                    &mut kind,
                    &mut legend_position,
                    &mut bar_dir,
                    &mut grouping,
                    &mut scatter_style,
                    &mut vary_colors,
                    &mut data_labels,
                    &mut style,
                    &mut current_series,
                );
                if let Some(axis) = current_axis_info.as_mut() {
                    apply_chart_axis_start(axis, &local, &e);
                }
            }
            Ok(Event::Text(e)) => {
                let text = e
                    .unescape()
                    .map_err(|err| ReaderError::Xml(format!("chart text: {err}")))?
                    .to_string();
                apply_chart_text(
                    &stack,
                    text.trim(),
                    &mut title_parts,
                    current_axis.as_deref(),
                    &mut current_axis_title_parts,
                    &mut current_series,
                );
            }
            Ok(Event::End(e)) => {
                let local_name = e.local_name();
                let local = local_name.as_ref();
                if local == b"ser" {
                    if let Some(ser) = current_series.take() {
                        series.push(ser);
                    }
                } else if chart_axis_name(local).is_some() {
                    let axis_info = current_axis_info.take();
                    apply_chart_axis_title(
                        local,
                        &kind,
                        &mut val_axis_titles_seen,
                        &current_axis_title_parts,
                        axis_info,
                        &mut x_axis_title,
                        &mut y_axis_title,
                        &mut x_axis,
                        &mut y_axis,
                    );
                    current_axis = None;
                    current_axis_title_parts.clear();
                }
                stack.pop();
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(ReaderError::Xml(format!("failed to parse chart: {e}"))),
            _ => {}
        }
        buf.clear();
    }

    let Some(kind) = kind else {
        return Ok(None);
    };
    let title = if title_parts.is_empty() {
        None
    } else {
        Some(title_parts.join(""))
    };
    Ok(Some(ChartInfo {
        source_drawing_path: None,
        source_chart_rid: None,
        source_chart_path: None,
        kind,
        title,
        x_axis_title,
        y_axis_title,
        x_axis,
        y_axis,
        data_labels,
        legend_position,
        bar_dir,
        grouping,
        scatter_style,
        vary_colors,
        style,
        anchor,
        series,
    }))
}

fn apply_chart_start(
    stack: &[Vec<u8>],
    local: &[u8],
    e: &BytesStart<'_>,
    kind: &mut Option<String>,
    legend_position: &mut Option<String>,
    bar_dir: &mut Option<String>,
    grouping: &mut Option<String>,
    scatter_style: &mut Option<String>,
    vary_colors: &mut Option<bool>,
    data_labels: &mut Option<ChartDataLabelsInfo>,
    style: &mut Option<u32>,
    current_series: &mut Option<ChartSeriesInfo>,
) {
    if kind.is_none() {
        *kind = chart_kind(local).map(str::to_string);
    }
    match local {
        b"style" => {
            *style = attr_u32(e, b"val");
        }
        b"legendPos" => {
            *legend_position = attr_value(e, b"val");
        }
        b"barDir" => {
            *bar_dir = attr_value(e, b"val");
        }
        b"grouping" => {
            *grouping = attr_value(e, b"val");
        }
        b"scatterStyle" => {
            *scatter_style = attr_value(e, b"val");
        }
        b"varyColors" => {
            *vary_colors = attr_bool(e, b"val");
        }
        b"spPr" => {
            if let Some(series) = current_series.as_mut() {
                series
                    .graphical_properties
                    .get_or_insert_with(ChartGraphicalPropertiesInfo::default);
            }
        }
        b"solidFill" => {
            if let Some(series) = current_series.as_mut() {
                chart_series_graphical_properties_mut(series);
            }
        }
        b"noFill" => {
            if let Some(series) = current_series.as_mut() {
                let gp = chart_series_graphical_properties_mut(series);
                if chart_path_contains(stack, b"ln") {
                    gp.line_no_fill = Some(true);
                } else {
                    gp.no_fill = Some(true);
                }
            }
        }
        b"ln" => {
            if let Some(series) = current_series.as_mut() {
                let gp = chart_series_graphical_properties_mut(series);
                gp.line_width = attr_u32(e, b"w").or(gp.line_width);
            }
        }
        b"srgbClr" => {
            if let Some(series) = current_series.as_mut() {
                let Some(color) = attr_value(e, b"val") else {
                    return;
                };
                let gp = chart_series_graphical_properties_mut(series);
                if chart_path_contains(stack, b"ln") {
                    gp.line_solid_fill = Some(color);
                } else {
                    gp.solid_fill = Some(color);
                }
            }
        }
        b"prstDash" => {
            if let Some(series) = current_series.as_mut() {
                chart_series_graphical_properties_mut(series).line_dash = attr_value(e, b"val");
            }
        }
        b"dLbls" => {
            chart_data_labels_mut(data_labels, current_series);
        }
        b"dLblPos" => {
            chart_data_labels_mut(data_labels, current_series).position = attr_value(e, b"val");
        }
        b"showLegendKey" => {
            chart_data_labels_mut(data_labels, current_series).show_legend_key =
                attr_bool(e, b"val");
        }
        b"showVal" => {
            chart_data_labels_mut(data_labels, current_series).show_val = attr_bool(e, b"val");
        }
        b"showCatName" => {
            chart_data_labels_mut(data_labels, current_series).show_cat_name = attr_bool(e, b"val");
        }
        b"showSerName" => {
            chart_data_labels_mut(data_labels, current_series).show_ser_name = attr_bool(e, b"val");
        }
        b"showPercent" => {
            chart_data_labels_mut(data_labels, current_series).show_percent = attr_bool(e, b"val");
        }
        b"showBubbleSize" => {
            chart_data_labels_mut(data_labels, current_series).show_bubble_size =
                attr_bool(e, b"val");
        }
        b"showLeaderLines" => {
            chart_data_labels_mut(data_labels, current_series).show_leader_lines =
                attr_bool(e, b"val");
        }
        b"trendline" => {
            if let Some(series) = current_series.as_mut() {
                series
                    .trendline
                    .get_or_insert_with(ChartTrendlineInfo::default);
            }
        }
        b"trendlineType" => {
            if let Some(series) = current_series.as_mut() {
                series
                    .trendline
                    .get_or_insert_with(ChartTrendlineInfo::default)
                    .trendline_type = attr_value(e, b"val");
            }
        }
        b"errBars" => {
            if let Some(series) = current_series.as_mut() {
                series
                    .error_bars
                    .get_or_insert_with(ChartErrorBarsInfo::default);
            }
        }
        b"errDir" => {
            if let Some(series) = current_series.as_mut() {
                series
                    .error_bars
                    .get_or_insert_with(ChartErrorBarsInfo::default)
                    .direction = attr_value(e, b"val");
            }
        }
        b"errBarType" => {
            if let Some(series) = current_series.as_mut() {
                series
                    .error_bars
                    .get_or_insert_with(ChartErrorBarsInfo::default)
                    .bar_type = attr_value(e, b"val");
            }
        }
        b"errValType" => {
            if let Some(series) = current_series.as_mut() {
                series
                    .error_bars
                    .get_or_insert_with(ChartErrorBarsInfo::default)
                    .val_type = attr_value(e, b"val");
            }
        }
        b"noEndCap" => {
            if let Some(series) = current_series.as_mut() {
                series
                    .error_bars
                    .get_or_insert_with(ChartErrorBarsInfo::default)
                    .no_end_cap = attr_bool(e, b"val");
            }
        }
        b"val" => {
            if let Some(series) = current_series.as_mut() {
                if let (Some(error_bars), Some(value)) =
                    (series.error_bars.as_mut(), attr_f64(e, b"val"))
                {
                    error_bars.val = Some(value);
                }
            }
        }
        b"ser" => {
            *current_series = Some(ChartSeriesInfo::default());
        }
        b"idx" => {
            if let Some(series) = current_series.as_mut() {
                series.idx = attr_u32(e, b"val");
            }
        }
        b"order" => {
            if let Some(series) = current_series.as_mut() {
                if let Some(trendline) = series.trendline.as_mut() {
                    trendline.order = attr_u32(e, b"val");
                } else {
                    series.order = attr_u32(e, b"val");
                }
            }
        }
        b"period" => {
            if let Some(series) = current_series.as_mut() {
                series
                    .trendline
                    .get_or_insert_with(ChartTrendlineInfo::default)
                    .period = attr_u32(e, b"val");
            }
        }
        b"forward" => {
            if let Some(series) = current_series.as_mut() {
                series
                    .trendline
                    .get_or_insert_with(ChartTrendlineInfo::default)
                    .forward = attr_f64(e, b"val");
            }
        }
        b"backward" => {
            if let Some(series) = current_series.as_mut() {
                series
                    .trendline
                    .get_or_insert_with(ChartTrendlineInfo::default)
                    .backward = attr_f64(e, b"val");
            }
        }
        b"intercept" => {
            if let Some(series) = current_series.as_mut() {
                series
                    .trendline
                    .get_or_insert_with(ChartTrendlineInfo::default)
                    .intercept = attr_f64(e, b"val");
            }
        }
        b"dispEq" => {
            if let Some(series) = current_series.as_mut() {
                series
                    .trendline
                    .get_or_insert_with(ChartTrendlineInfo::default)
                    .display_equation = attr_bool(e, b"val");
            }
        }
        b"dispRSqr" => {
            if let Some(series) = current_series.as_mut() {
                series
                    .trendline
                    .get_or_insert_with(ChartTrendlineInfo::default)
                    .display_r_squared = attr_bool(e, b"val");
            }
        }
        _ => {}
    }
}

fn apply_chart_text(
    stack: &[Vec<u8>],
    text: &str,
    title_parts: &mut Vec<String>,
    current_axis: Option<&[u8]>,
    current_axis_title_parts: &mut Vec<String>,
    current_series: &mut Option<ChartSeriesInfo>,
) {
    if text.is_empty() {
        return;
    }
    let Some(last) = stack.last().map(Vec::as_slice) else {
        return;
    };
    if let Some(series) = current_series.as_mut() {
        if last == b"f" {
            if chart_path_contains(stack, b"tx") && chart_path_contains(stack, b"strRef") {
                series.title_ref = Some(text.to_string());
            } else if chart_path_contains(stack, b"cat") {
                series.cat_ref = Some(text.to_string());
            } else if chart_path_contains(stack, b"val") {
                series.val_ref = Some(text.to_string());
            } else if chart_path_contains(stack, b"xVal") {
                series.x_ref = Some(text.to_string());
            } else if chart_path_contains(stack, b"yVal") {
                series.y_ref = Some(text.to_string());
            } else if chart_path_contains(stack, b"bubbleSize") {
                series.bubble_size_ref = Some(text.to_string());
            }
        } else if last == b"v" && chart_path_contains(stack, b"tx") {
            series.title_value = Some(text.to_string());
        }
        return;
    }

    if last == b"t" && chart_path_contains(stack, b"title") {
        if current_axis.is_some() {
            current_axis_title_parts.push(text.to_string());
        } else {
            title_parts.push(text.to_string());
        }
    }
}

fn chart_path_contains(stack: &[Vec<u8>], name: &[u8]) -> bool {
    stack.iter().any(|part| part.as_slice() == name)
}

fn chart_data_labels_mut<'a>(
    chart_labels: &'a mut Option<ChartDataLabelsInfo>,
    current_series: &'a mut Option<ChartSeriesInfo>,
) -> &'a mut ChartDataLabelsInfo {
    if let Some(series) = current_series.as_mut() {
        return series
            .data_labels
            .get_or_insert_with(ChartDataLabelsInfo::default);
    }
    chart_labels.get_or_insert_with(ChartDataLabelsInfo::default)
}

fn chart_series_graphical_properties_mut(
    series: &mut ChartSeriesInfo,
) -> &mut ChartGraphicalPropertiesInfo {
    series
        .graphical_properties
        .get_or_insert_with(ChartGraphicalPropertiesInfo::default)
}

fn chart_axis_name(local: &[u8]) -> Option<&'static str> {
    match local {
        b"catAx" | b"dateAx" | b"serAx" => Some("x"),
        b"valAx" => Some("value"),
        _ => None,
    }
}

fn new_chart_axis_info(local: &[u8], e: &BytesStart<'_>) -> ChartAxisInfo {
    let axis_type = match local {
        b"catAx" => "cat",
        b"dateAx" => "date",
        b"serAx" => "ser",
        b"valAx" => "val",
        _ => "unknown",
    };
    let mut axis = ChartAxisInfo {
        axis_type: axis_type.to_string(),
        ..ChartAxisInfo::default()
    };
    apply_chart_axis_start(&mut axis, local, e);
    axis
}

fn apply_chart_axis_start(axis: &mut ChartAxisInfo, local: &[u8], e: &BytesStart<'_>) {
    match local {
        b"axId" => axis.ax_id = attr_u32(e, b"val"),
        b"crossAx" => axis.cross_ax = attr_u32(e, b"val"),
        b"axPos" => axis.axis_position = attr_value(e, b"val"),
        b"min" => axis.scaling_min = attr_f64(e, b"val"),
        b"max" => axis.scaling_max = attr_f64(e, b"val"),
        b"orientation" => axis.scaling_orientation = attr_value(e, b"val"),
        b"logBase" => axis.scaling_log_base = attr_f64(e, b"val"),
        b"numFmt" => {
            axis.num_format_code = attr_value(e, b"formatCode");
            axis.num_format_source_linked = attr_bool(e, b"sourceLinked");
        }
        b"majorUnit" => axis.major_unit = attr_f64(e, b"val"),
        b"minorUnit" => axis.minor_unit = attr_f64(e, b"val"),
        b"tickLblPos" => axis.tick_lbl_pos = attr_value(e, b"val"),
        b"majorTickMark" => axis.major_tick_mark = attr_value(e, b"val"),
        b"minorTickMark" => axis.minor_tick_mark = attr_value(e, b"val"),
        b"crosses" => axis.crosses = attr_value(e, b"val"),
        b"crossesAt" => axis.crosses_at = attr_f64(e, b"val"),
        b"crossBetween" => axis.cross_between = attr_value(e, b"val"),
        b"builtInUnit" => axis.display_unit = attr_value(e, b"val"),
        _ => {}
    }
}

fn apply_chart_axis_title(
    axis: &[u8],
    kind: &Option<String>,
    val_axis_titles_seen: &mut usize,
    title_parts: &[String],
    axis_info: Option<ChartAxisInfo>,
    x_axis_title: &mut Option<String>,
    y_axis_title: &mut Option<String>,
    x_axis: &mut Option<ChartAxisInfo>,
    y_axis: &mut Option<ChartAxisInfo>,
) {
    match axis {
        b"catAx" | b"dateAx" | b"serAx" => {
            if x_axis_title.is_none() {
                apply_axis_title_parts(title_parts, x_axis_title);
            }
            if x_axis.is_none() {
                *x_axis = axis_info;
            }
        }
        b"valAx" => {
            if matches!(kind.as_deref(), Some("scatter" | "bubble")) {
                *val_axis_titles_seen += 1;
                if *val_axis_titles_seen == 1 && x_axis_title.is_none() {
                    apply_axis_title_parts(title_parts, x_axis_title);
                    if x_axis.is_none() {
                        *x_axis = axis_info;
                    }
                } else if y_axis_title.is_none() {
                    apply_axis_title_parts(title_parts, y_axis_title);
                    if y_axis.is_none() {
                        *y_axis = axis_info;
                    }
                } else if y_axis.is_none() {
                    *y_axis = axis_info;
                }
            } else {
                if y_axis_title.is_none() {
                    apply_axis_title_parts(title_parts, y_axis_title);
                }
                if y_axis.is_none() {
                    *y_axis = axis_info;
                }
            }
        }
        _ => {}
    }
}

fn apply_axis_title_parts(title_parts: &[String], target: &mut Option<String>) {
    if !title_parts.is_empty() {
        *target = Some(title_parts.join(""));
    }
}

fn chart_kind(local: &[u8]) -> Option<&'static str> {
    match local {
        b"barChart" => Some("bar"),
        b"bar3DChart" => Some("bar3d"),
        b"lineChart" => Some("line"),
        b"line3DChart" => Some("line3d"),
        b"pieChart" => Some("pie"),
        b"pie3DChart" => Some("pie3d"),
        b"ofPieChart" => Some("of_pie"),
        b"doughnutChart" => Some("doughnut"),
        b"areaChart" => Some("area"),
        b"area3DChart" => Some("area3d"),
        b"scatterChart" => Some("scatter"),
        b"bubbleChart" => Some("bubble"),
        b"radarChart" => Some("radar"),
        b"surfaceChart" => Some("surface"),
        b"surface3DChart" => Some("surface3d"),
        b"stockChart" => Some("stock"),
        _ => None,
    }
}

fn apply_anchor_ext(builder: &mut Option<DrawingObjectBuilder>, e: &BytesStart<'_>) {
    let Some(builder) = builder.as_mut() else {
        return;
    };
    let Some(cx) = attr_i64(e, b"cx") else {
        return;
    };
    let Some(cy) = attr_i64(e, b"cy") else {
        return;
    };
    builder.ext = Some(AnchorExtentInfo { cx, cy });
}

fn apply_blip_rid(builder: &mut Option<DrawingObjectBuilder>, e: &BytesStart<'_>) {
    let Some(builder) = builder.as_mut() else {
        return;
    };
    builder.rid = attr_value(e, b"r:embed").or_else(|| attr_value(e, b"embed"));
}

fn apply_marker_value(
    builder: &mut DrawingObjectBuilder,
    slot: MarkerSlot,
    target: MarkerTextTarget,
    value: i64,
) {
    let marker = match slot {
        MarkerSlot::From => &mut builder.from,
        MarkerSlot::To => &mut builder.to,
    };
    match target {
        MarkerTextTarget::Col => marker.col = value,
        MarkerTextTarget::Row => marker.row = value,
        MarkerTextTarget::ColOff => marker.col_off = value,
        MarkerTextTarget::RowOff => marker.row_off = value,
    }
}

fn image_ext_from_path(path: &str) -> String {
    PathBuf::from(path)
        .extension()
        .and_then(|ext| ext.to_str())
        .map(str::to_ascii_lowercase)
        .unwrap_or_default()
}

fn parse_table_rids(xml: &str) -> Result<Vec<String>> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut out = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.local_name().as_ref() == b"tablePart" {
                    if let Some(rid) = attr_value(&e, b"r:id").filter(|value| !value.is_empty()) {
                        out.push(rid);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ReaderError::Xml(format!(
                    "failed to parse worksheet table parts: {e}"
                )));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

fn parse_table_xml(xml: &str) -> Result<Table> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut name = String::new();
    let mut ref_range = String::new();
    let mut header_row = true;
    let mut totals_row = false;
    let mut comment = None;
    let mut table_type = None;
    let mut totals_row_shown = None;
    let mut style = None;
    let mut show_first_column = false;
    let mut show_last_column = false;
    let mut show_row_stripes = false;
    let mut show_column_stripes = false;
    let mut columns = Vec::new();
    let auto_filter = parse_auto_filter(xml)?;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.local_name().as_ref() {
                b"table" => {
                    name = attr_value(&e, b"name")
                        .or_else(|| attr_value(&e, b"displayName"))
                        .unwrap_or_default();
                    ref_range = attr_value(&e, b"ref").unwrap_or_default();
                    header_row = attr_value(&e, b"headerRowCount")
                        .map(|value| value != "0")
                        .unwrap_or(true);
                    totals_row = attr_value(&e, b"totalsRowCount")
                        .map(|value| value != "0")
                        .unwrap_or(false);
                    comment = attr_value(&e, b"comment").filter(|value| !value.is_empty());
                    table_type = attr_value(&e, b"tableType").filter(|value| !value.is_empty());
                    totals_row_shown = attr_bool(&e, b"totalsRowShown");
                }
                b"tableStyleInfo" => {
                    style = attr_value(&e, b"name").filter(|value| !value.is_empty());
                    show_first_column = attr_truthy(attr_value(&e, b"showFirstColumn").as_deref());
                    show_last_column = attr_truthy(attr_value(&e, b"showLastColumn").as_deref());
                    show_row_stripes = attr_truthy(attr_value(&e, b"showRowStripes").as_deref());
                    show_column_stripes =
                        attr_truthy(attr_value(&e, b"showColumnStripes").as_deref());
                }
                b"tableColumn" => {
                    if let Some(column) = attr_value(&e, b"name") {
                        columns.push(column);
                    }
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(ReaderError::Xml(format!("failed to parse table XML: {e}"))),
            _ => {}
        }
        buf.clear();
    }

    Ok(Table {
        name,
        ref_range,
        header_row,
        totals_row,
        comment,
        table_type,
        totals_row_shown,
        style,
        show_first_column,
        show_last_column,
        show_row_stripes,
        show_column_stripes,
        columns,
        auto_filter,
    })
}

#[derive(Debug)]
struct CellBuilder {
    row: u32,
    col: u32,
    coordinate: String,
    style_id: Option<u32>,
    data_type: CellDataType,
    value_text: String,
    inline_text: String,
    formula_text: String,
    formula_kind: Option<String>,
    formula_shared_index: Option<String>,
    formula_ref: Option<String>,
    formula_ca: bool,
    formula_dt2_d: bool,
    formula_dtr: bool,
    formula_r1: Option<String>,
    formula_r2: Option<String>,
    formula_del1: bool,
    formula_del2: bool,
    inline_runs: Vec<RichTextRun>,
    inline_current_run: Option<RichTextRun>,
    inline_current_props: Option<InlineFontProps>,
    inline_saw_r: bool,
}

#[derive(Debug)]
struct ValueCellBuilder {
    row: u32,
    col: u32,
    data_type: CellDataType,
    value_text: Vec<u8>,
    fast_number: Option<f64>,
    fast_number_text: [u8; 64],
    fast_number_text_len: u8,
    inline_text: String,
}

#[derive(Debug)]
struct FormulaSummaryCellBuilder {
    row: u32,
    col: u32,
    data_type: CellDataType,
    value_text: Vec<u8>,
    fast_number: Option<f64>,
    fast_number_text: [u8; 64],
    fast_number_text_len: u8,
    has_formula: bool,
}

impl FormulaSummaryCellBuilder {
    fn from_start(e: &BytesStart<'_>, current_row: Option<u32>) -> Self {
        let base = ValueCellBuilder::from_start(e, current_row);
        Self {
            row: base.row,
            col: base.col,
            data_type: base.data_type,
            value_text: Vec::new(),
            fast_number: None,
            fast_number_text: [0; 64],
            fast_number_text_len: 0,
            has_formula: false,
        }
    }

    fn push_raw_value_text(&mut self, text: &[u8]) {
        if self.data_type == CellDataType::Number {
            self.push_number_text(text);
        } else {
            self.value_text.extend_from_slice(text);
        }
    }

    fn push_number_text(&mut self, text: &[u8]) {
        if self.fast_number.take().is_some() {
            let len = self.fast_number_text_len as usize;
            self.value_text
                .extend_from_slice(&self.fast_number_text[..len]);
            self.fast_number_text_len = 0;
        }
        if self.value_text.is_empty() {
            let raw = trim_ascii(text);
            if raw.is_empty() {
                return;
            }
            if text.len() <= self.fast_number_text.len() {
                if let Ok(raw_text) = std::str::from_utf8(raw) {
                    if let Ok(value) = raw_text.parse::<f64>() {
                        self.fast_number = Some(value);
                        self.fast_number_text[..text.len()].copy_from_slice(text);
                        self.fast_number_text_len = text.len() as u8;
                        return;
                    }
                }
            }
        }
        self.value_text.extend_from_slice(text);
    }

    fn update_summary(self, summary: &mut FormulaSheetSummary) -> Result<()> {
        summary.rows = summary.rows.max(self.row as usize);
        if self.has_formula {
            summary.formula_cells += 1;
        }
        if self.col == 1 && self.data_type == CellDataType::Number {
            if let Some(value) = self.fast_number {
                summary.first_col_numeric_checksum += value;
            } else {
                let raw = trim_ascii(&self.value_text);
                if !raw.is_empty() {
                    let text = std::str::from_utf8(raw).map_err(|e| {
                        ReaderError::Xml(format!(
                            "invalid numeric cell {}: {e}",
                            row_col_to_a1(self.row, self.col)
                        ))
                    })?;
                    summary.first_col_numeric_checksum += text.parse::<f64>().map_err(|e| {
                        ReaderError::Xml(format!(
                            "invalid numeric cell {}: {e}",
                            row_col_to_a1(self.row, self.col)
                        ))
                    })?;
                }
            }
        }
        Ok(())
    }

    fn update_value_summary(self, summary: &mut ValueSheetSummary) -> Result<()> {
        summary.rows = summary.rows.max(self.row as usize);
        if self.data_type == CellDataType::Number {
            if let Some(value) = self.fast_number {
                summary.numeric_checksum += value;
            } else {
                let raw = trim_ascii(&self.value_text);
                if !raw.is_empty() {
                    let text = std::str::from_utf8(raw).map_err(|e| {
                        ReaderError::Xml(format!(
                            "invalid numeric cell {}: {e}",
                            row_col_to_a1(self.row, self.col)
                        ))
                    })?;
                    summary.numeric_checksum += text.parse::<f64>().map_err(|e| {
                        ReaderError::Xml(format!(
                            "invalid numeric cell {}: {e}",
                            row_col_to_a1(self.row, self.col)
                        ))
                    })?;
                }
            }
        }
        Ok(())
    }
}

impl ValueCellBuilder {
    fn from_start(e: &BytesStart<'_>, current_row: Option<u32>) -> Self {
        let mut row_col: Option<(u32, u32)> = None;
        let mut data_type = CellDataType::Number;
        for attr in e.attributes().with_checks(false).flatten() {
            match attr.key.as_ref() {
                b"r" => row_col = a1_bytes_to_row_col(attr.value.as_ref()),
                b"t" => {
                    data_type = match attr.value.as_ref() {
                        b"s" => CellDataType::SharedString,
                        b"inlineStr" => CellDataType::InlineString,
                        b"str" => CellDataType::FormulaString,
                        b"b" => CellDataType::Bool,
                        b"e" => CellDataType::Error,
                        b"d" => CellDataType::DateIso,
                        _ => CellDataType::Number,
                    };
                }
                _ => {}
            }
        }
        let (row, col) = row_col.unwrap_or((current_row.unwrap_or(1), 1));
        Self {
            row,
            col,
            data_type,
            value_text: Vec::new(),
            fast_number: None,
            fast_number_text: [0; 64],
            fast_number_text_len: 0,
            inline_text: String::new(),
        }
    }

    fn is_inline_string(&self) -> bool {
        self.data_type == CellDataType::InlineString
    }

    fn push_text(&mut self, target: TextTarget, text: &str) {
        match target {
            TextTarget::Value | TextTarget::Formula => {
                self.value_text.extend_from_slice(text.as_bytes())
            }
            TextTarget::InlineString => self.inline_text.push_str(&normalize_ooxml_text(text)),
        }
    }

    fn push_raw_text(&mut self, target: TextTarget, text: &[u8]) {
        if matches!(target, TextTarget::InlineString) {
            self.push_text(target, &String::from_utf8_lossy(text));
        } else if self.data_type == CellDataType::Number {
            self.push_number_text(text);
        } else {
            self.value_text.extend_from_slice(text);
        }
    }

    fn push_number_text(&mut self, text: &[u8]) {
        if self.fast_number.take().is_some() {
            let len = self.fast_number_text_len as usize;
            self.value_text
                .extend_from_slice(&self.fast_number_text[..len]);
            self.fast_number_text_len = 0;
        }
        if self.value_text.is_empty() {
            let raw = trim_ascii(text);
            if raw.is_empty() {
                return;
            }
            if text.len() <= self.fast_number_text.len() {
                if let Ok(raw_text) = std::str::from_utf8(raw) {
                    if let Ok(value) = raw_text.parse::<f64>() {
                        self.fast_number = Some(value);
                        self.fast_number_text[..text.len()].copy_from_slice(text);
                        self.fast_number_text_len = text.len() as u8;
                        return;
                    }
                }
            }
        }
        self.value_text.extend_from_slice(text);
    }

    fn finish(self, shared_strings: &SharedStrings) -> Result<ValueCell> {
        let shared_string_idx = if self.data_type == CellDataType::SharedString {
            parse_usize_ascii(trim_ascii(&self.value_text))
        } else {
            None
        };
        let value = match self.data_type {
            CellDataType::SharedString => shared_string_idx
                .and_then(|i| shared_strings.values.get(i).cloned())
                .map(CellValue::String)
                .unwrap_or(CellValue::Empty),
            CellDataType::InlineString => {
                if self.inline_text.is_empty() {
                    CellValue::Empty
                } else {
                    CellValue::String(self.inline_text)
                }
            }
            CellDataType::Bool => {
                CellValue::Bool(matches!(trim_ascii(&self.value_text), b"1" | b"true"))
            }
            CellDataType::Error => {
                CellValue::Error(String::from_utf8_lossy(&self.value_text).into_owned())
            }
            CellDataType::DateIso | CellDataType::FormulaString => {
                CellValue::String(String::from_utf8_lossy(&self.value_text).into_owned())
            }
            CellDataType::Number => {
                if let Some(value) = self.fast_number {
                    return Ok(ValueCell {
                        row: self.row,
                        col: self.col,
                        value: CellValue::Number(value),
                    });
                }
                let raw = trim_ascii(&self.value_text);
                if raw.is_empty() {
                    CellValue::Empty
                } else {
                    let text = std::str::from_utf8(raw).map_err(|e| {
                        ReaderError::Xml(format!(
                            "invalid numeric cell {}: {e}",
                            row_col_to_a1(self.row, self.col)
                        ))
                    })?;
                    text.parse::<f64>().map(CellValue::Number).map_err(|e| {
                        ReaderError::Xml(format!(
                            "invalid numeric cell {}: {e}",
                            row_col_to_a1(self.row, self.col)
                        ))
                    })?
                }
            }
        };
        Ok(ValueCell {
            row: self.row,
            col: self.col,
            value,
        })
    }
}

impl CellBuilder {
    fn from_start(e: &BytesStart<'_>, current_row: Option<u32>) -> Self {
        let coordinate = attr_value(e, b"r");
        let (row, col) = coordinate
            .as_deref()
            .and_then(a1_to_row_col)
            .unwrap_or((current_row.unwrap_or(1), 1));
        let coordinate = coordinate.unwrap_or_else(|| row_col_to_a1(row, col));
        let data_type = match attr_value(e, b"t").as_deref() {
            Some("s") => CellDataType::SharedString,
            Some("inlineStr") => CellDataType::InlineString,
            Some("str") => CellDataType::FormulaString,
            Some("b") => CellDataType::Bool,
            Some("e") => CellDataType::Error,
            Some("d") => CellDataType::DateIso,
            _ => CellDataType::Number,
        };
        let style_id = attr_value(e, b"s").and_then(|v| v.parse::<u32>().ok());
        Self {
            row,
            col,
            coordinate,
            style_id,
            data_type,
            value_text: String::new(),
            inline_text: String::new(),
            formula_text: String::new(),
            formula_kind: None,
            formula_shared_index: None,
            formula_ref: None,
            formula_ca: false,
            formula_dt2_d: false,
            formula_dtr: false,
            formula_r1: None,
            formula_r2: None,
            formula_del1: false,
            formula_del2: false,
            inline_runs: Vec::new(),
            inline_current_run: None,
            inline_current_props: None,
            inline_saw_r: false,
        }
    }

    fn start_formula(&mut self, e: &BytesStart<'_>) {
        self.formula_kind = attr_value(e, b"t");
        self.formula_shared_index = attr_value(e, b"si");
        self.formula_ref = attr_value(e, b"ref");
        self.formula_ca = attr_truthy(attr_value(e, b"ca").as_deref());
        self.formula_dt2_d = attr_truthy(attr_value(e, b"dt2D").as_deref());
        self.formula_dtr = attr_truthy(attr_value(e, b"dtr").as_deref());
        self.formula_r1 = attr_value(e, b"r1");
        self.formula_r2 = attr_value(e, b"r2");
        self.formula_del1 = attr_truthy(attr_value(e, b"del1").as_deref());
        self.formula_del2 = attr_truthy(attr_value(e, b"del2").as_deref());
    }

    fn push_text(&mut self, target: TextTarget, text: &str) {
        match target {
            TextTarget::Value => self.value_text.push_str(text),
            TextTarget::Formula => self.formula_text.push_str(text),
            TextTarget::InlineString => {
                let text = normalize_ooxml_text(text);
                self.inline_text.push_str(&text);
                if let Some(run) = self.inline_current_run.as_mut() {
                    run.text.push_str(&text);
                }
            }
        }
    }

    fn start_inline_run(&mut self) {
        self.inline_saw_r = true;
        self.inline_current_run = Some(RichTextRun::default());
        self.inline_current_props = None;
    }

    fn start_inline_props(&mut self) {
        if self.inline_current_run.is_some() {
            self.inline_current_props = Some(InlineFontProps::default());
        }
    }

    fn apply_inline_font_tag(
        &mut self,
        tag: &[u8],
        attrs: quick_xml::events::attributes::Attributes<'_>,
    ) {
        if let Some(props) = self.inline_current_props.as_mut() {
            apply_rich_font_attr(props, tag, attrs);
        }
    }

    fn end_inline_props(&mut self) {
        if let (Some(run), Some(props)) = (
            self.inline_current_run.as_mut(),
            self.inline_current_props.take(),
        ) {
            run.font = Some(props);
        }
    }

    fn end_inline_run(&mut self) {
        if let Some(run) = self.inline_current_run.take() {
            self.inline_runs.push(run);
        }
        self.inline_current_props = None;
    }

    fn finish(mut self, shared_strings: &SharedStrings) -> Result<Cell> {
        if self.inline_current_props.is_some() {
            self.end_inline_props();
        }
        if self.inline_current_run.is_some() {
            self.end_inline_run();
        }
        let shared_string_idx = if self.data_type == CellDataType::SharedString {
            self.value_text.trim().parse::<usize>().ok()
        } else {
            None
        };
        let rich_text = match self.data_type {
            CellDataType::SharedString => shared_string_idx
                .and_then(|idx| shared_strings.rich_text.get(idx).cloned().flatten()),
            CellDataType::InlineString if self.inline_saw_r => Some(self.inline_runs),
            _ => None,
        };
        let value = match self.data_type {
            CellDataType::SharedString => shared_string_idx
                .and_then(|i| shared_strings.values.get(i).cloned())
                .map(CellValue::String)
                .unwrap_or(CellValue::Empty),
            CellDataType::InlineString => {
                if self.inline_text.is_empty() {
                    CellValue::Empty
                } else {
                    CellValue::String(self.inline_text)
                }
            }
            CellDataType::Bool => CellValue::Bool(matches!(self.value_text.trim(), "1" | "true")),
            CellDataType::Error => CellValue::Error(self.value_text),
            CellDataType::DateIso => CellValue::String(self.value_text),
            CellDataType::FormulaString => CellValue::String(self.value_text),
            CellDataType::Number => {
                let raw = self.value_text.trim();
                if raw.is_empty() {
                    CellValue::Empty
                } else {
                    raw.parse::<f64>().map(CellValue::Number).map_err(|e| {
                        ReaderError::Xml(format!("invalid numeric cell {}: {e}", self.coordinate))
                    })?
                }
            }
        };
        let array_formula = match self.formula_kind.as_deref() {
            Some("array") => Some(ArrayFormulaInfo::Array {
                ref_range: self.formula_ref.clone().unwrap_or_default(),
                text: self.formula_text.clone(),
            }),
            Some("dataTable") => Some(ArrayFormulaInfo::DataTable {
                ref_range: self.formula_ref.clone().unwrap_or_default(),
                ca: self.formula_ca,
                dt2_d: self.formula_dt2_d,
                dtr: self.formula_dtr,
                r1: self.formula_r1.clone(),
                r2: self.formula_r2.clone(),
                del1: self.formula_del1,
                del2: self.formula_del2,
            }),
            _ => None,
        };
        Ok(Cell {
            row: self.row,
            col: self.col,
            coordinate: self.coordinate,
            style_id: self.style_id,
            data_type: self.data_type,
            value,
            formula: if self.formula_text.is_empty() {
                None
            } else {
                Some(self.formula_text)
            },
            formula_kind: self.formula_kind,
            formula_shared_index: self.formula_shared_index,
            array_formula,
            rich_text,
        })
    }
}

#[derive(Debug, Clone, Copy)]
enum TextTarget {
    Value,
    Formula,
    InlineString,
}

fn zip_from_source(source: &WorkbookSource) -> Result<ZipArchive<ZipSource<'_>>> {
    let reader = match source {
        WorkbookSource::Path(path) => ZipSource::Path(File::open(path)?),
        WorkbookSource::Bytes(bytes) => ZipSource::Bytes(Cursor::new(bytes.as_slice())),
    };
    ZipArchive::new(reader).map_err(ReaderError::Zip)
}

fn env_usize(name: &str, default: usize) -> usize {
    std::env::var(name)
        .ok()
        .and_then(|v| v.parse::<usize>().ok())
        .unwrap_or(default)
}

fn env_u64(name: &str, default: u64) -> u64 {
    std::env::var(name)
        .ok()
        .and_then(|v| v.parse::<u64>().ok())
        .unwrap_or(default)
}

fn max_zip_entries() -> usize {
    env_usize("WOLFXL_MAX_ZIP_ENTRIES", DEFAULT_MAX_ZIP_ENTRIES)
}

fn max_zip_entry_bytes() -> u64 {
    env_u64("WOLFXL_MAX_ZIP_ENTRY_BYTES", DEFAULT_MAX_ZIP_ENTRY_BYTES)
}

fn max_zip_total_bytes() -> u64 {
    env_u64("WOLFXL_MAX_ZIP_TOTAL_BYTES", DEFAULT_MAX_ZIP_TOTAL_BYTES)
}

fn max_zip_compression_ratio() -> u64 {
    env_u64(
        "WOLFXL_MAX_ZIP_COMPRESSION_RATIO",
        DEFAULT_MAX_COMPRESSION_RATIO,
    )
}

fn validate_zip_archive<R: Read + Seek>(zip: &mut ZipArchive<R>) -> Result<()> {
    if zip.len() > max_zip_entries() {
        return Err(ReaderError::Unsupported(format!(
            "OOXML package has too many ZIP entries: {} > {}",
            zip.len(),
            max_zip_entries()
        )));
    }
    let mut names = HashSet::with_capacity(zip.len());
    let mut casefolded_names = HashSet::with_capacity(zip.len());
    let mut total: u64 = 0;
    for i in 0..zip.len() {
        let file = zip.by_index(i)?;
        let name = file.name().to_string();
        validate_zip_part_name(&name)?;
        if !names.insert(name.clone()) {
            return Err(ReaderError::Unsupported(format!(
                "OOXML package contains duplicate ZIP entry: {name}"
            )));
        }
        let casefolded_name = name.to_ascii_lowercase();
        if !casefolded_names.insert(casefolded_name) {
            return Err(ReaderError::Unsupported(format!(
                "OOXML package contains case-insensitive duplicate ZIP entry: {name}"
            )));
        }
        validate_zip_entry_metadata(&name, file.size(), file.compressed_size())?;
        total = total.saturating_add(file.size());
        if total > max_zip_total_bytes() {
            return Err(ReaderError::Unsupported(format!(
                "OOXML package is too large: {total} > {} uncompressed bytes",
                max_zip_total_bytes()
            )));
        }
    }
    Ok(())
}

fn validate_zip_part_name(name: &str) -> Result<()> {
    let invalid = name.is_empty()
        || name.starts_with('/')
        || name.starts_with('\\')
        || name.contains('\\')
        || name
            .split('/')
            .any(|part| part == ".." || part.contains(':'));
    if invalid {
        return Err(ReaderError::Unsupported(format!(
            "unsafe OOXML package part path: {name}"
        )));
    }
    Ok(())
}

fn validate_zip_entry_metadata(name: &str, size: u64, compressed_size: u64) -> Result<()> {
    if size > max_zip_entry_bytes() {
        return Err(ReaderError::Unsupported(format!(
            "OOXML package part {name} is too large: {size} > {} bytes",
            max_zip_entry_bytes()
        )));
    }
    if size > 0 && compressed_size == 0 {
        return Err(ReaderError::Unsupported(format!(
            "OOXML package part {name} has invalid compressed size"
        )));
    }
    if compressed_size > 0 && size > compressed_size.saturating_mul(max_zip_compression_ratio()) {
        return Err(ReaderError::Unsupported(format!(
            "OOXML package part {name} exceeds compression ratio limit"
        )));
    }
    Ok(())
}

fn normalize_ooxml_text(text: &str) -> String {
    text.replace("\r\n", "\n").replace('\r', "\n")
}

fn read_part_required<R: Read + std::io::Seek>(
    zip: &mut ZipArchive<R>,
    name: &str,
) -> Result<String> {
    read_part_optional(zip, name)?.ok_or_else(|| ReaderError::MissingPart(name.to_string()))
}

fn read_part_required_bytes<R: Read + std::io::Seek>(
    zip: &mut ZipArchive<R>,
    name: &str,
) -> Result<Vec<u8>> {
    read_part_optional_bytes(zip, name)?.ok_or_else(|| ReaderError::MissingPart(name.to_string()))
}

fn read_part_optional<R: Read + std::io::Seek>(
    zip: &mut ZipArchive<R>,
    name: &str,
) -> Result<Option<String>> {
    match zip.by_name(name) {
        Ok(mut file) => {
            validate_zip_entry_metadata(name, file.size(), file.compressed_size())?;
            let mut out = String::new();
            file.read_to_string(&mut out)?;
            return Ok(Some(out));
        }
        Err(zip::result::ZipError::FileNotFound) => {}
        Err(e) => return Err(ReaderError::Zip(e)),
    }
    let Some(actual_name) = resolve_zip_part_name(zip, name) else {
        return Ok(None);
    };
    match zip.by_name(&actual_name) {
        Ok(mut file) => {
            validate_zip_entry_metadata(&actual_name, file.size(), file.compressed_size())?;
            let mut out = String::new();
            file.read_to_string(&mut out)?;
            Ok(Some(out))
        }
        Err(e) => Err(ReaderError::Zip(e)),
    }
}

fn read_part_optional_bytes<R: Read + std::io::Seek>(
    zip: &mut ZipArchive<R>,
    name: &str,
) -> Result<Option<Vec<u8>>> {
    match zip.by_name(name) {
        Ok(mut file) => {
            validate_zip_entry_metadata(name, file.size(), file.compressed_size())?;
            let mut out = Vec::new();
            file.read_to_end(&mut out)?;
            return Ok(Some(out));
        }
        Err(zip::result::ZipError::FileNotFound) => {}
        Err(e) => return Err(ReaderError::Zip(e)),
    }
    let Some(actual_name) = resolve_zip_part_name(zip, name) else {
        return Ok(None);
    };
    match zip.by_name(&actual_name) {
        Ok(mut file) => {
            validate_zip_entry_metadata(&actual_name, file.size(), file.compressed_size())?;
            let mut out = Vec::new();
            file.read_to_end(&mut out)?;
            Ok(Some(out))
        }
        Err(e) => Err(ReaderError::Zip(e)),
    }
}

fn resolve_zip_part_name<R: Read + std::io::Seek>(
    zip: &ZipArchive<R>,
    name: &str,
) -> Option<String> {
    zip.file_names()
        .find(|candidate| candidate.eq_ignore_ascii_case(name))
        .map(str::to_string)
}

fn attr_value(e: &BytesStart<'_>, key: &[u8]) -> Option<String> {
    for a in e.attributes().with_checks(false).flatten() {
        if a.key.as_ref() == key {
            if let Ok(v) = a.unescape_value() {
                return Some(v.to_string());
            }
            return Some(String::from_utf8_lossy(a.value.as_ref()).into_owned());
        }
    }
    None
}

fn relationship_id_attr(e: &BytesStart<'_>) -> Option<String> {
    attr_value(e, b"r:id").or_else(|| attr_value_by_suffix(e, b":id"))
}

fn attr_value_by_suffix(e: &BytesStart<'_>, suffix: &[u8]) -> Option<String> {
    for a in e.attributes().with_checks(false).flatten() {
        if a.key.as_ref().ends_with(suffix) {
            if let Ok(v) = a.unescape_value() {
                return Some(v.to_string());
            }
            return Some(String::from_utf8_lossy(a.value.as_ref()).into_owned());
        }
    }
    None
}

fn attr_truthy(value: Option<&str>) -> bool {
    matches!(value, Some(v) if v != "0" && !v.eq_ignore_ascii_case("false"))
}

fn attr_bool_default(e: &BytesStart<'_>, key: &[u8], default: bool) -> bool {
    attr_value(e, key)
        .as_deref()
        .map_or(default, |value| attr_truthy(Some(value)))
}

fn attr_bool(e: &BytesStart<'_>, key: &[u8]) -> Option<bool> {
    attr_value(e, key)
        .as_deref()
        .map(|value| attr_truthy(Some(value)))
}

fn attr_u32(e: &BytesStart<'_>, key: &[u8]) -> Option<u32> {
    attr_value(e, key).and_then(|value| value.parse::<u32>().ok())
}

fn raw_attr_u32(e: &BytesStart<'_>, key: &[u8]) -> Option<u32> {
    for attr in e.attributes().with_checks(false).flatten() {
        if attr.key.as_ref() == key {
            let mut value = 0u32;
            for byte in attr.value.as_ref() {
                if !byte.is_ascii_digit() {
                    return None;
                }
                value = value.checked_mul(10)?.checked_add((byte - b'0') as u32)?;
            }
            return (value != 0).then_some(value);
        }
    }
    None
}

fn trim_ascii(mut value: &[u8]) -> &[u8] {
    while let Some((first, rest)) = value.split_first() {
        if !first.is_ascii_whitespace() {
            break;
        }
        value = rest;
    }
    while let Some((last, rest)) = value.split_last() {
        if !last.is_ascii_whitespace() {
            break;
        }
        value = rest;
    }
    value
}

fn parse_usize_ascii(value: &[u8]) -> Option<usize> {
    let mut out = 0usize;
    for byte in value {
        if !byte.is_ascii_digit() {
            return None;
        }
        out = out.checked_mul(10)?.checked_add((byte - b'0') as usize)?;
    }
    Some(out)
}

fn attr_i64(e: &BytesStart<'_>, key: &[u8]) -> Option<i64> {
    attr_value(e, key).and_then(|value| value.parse::<i64>().ok())
}

fn attr_f64(e: &BytesStart<'_>, key: &[u8]) -> Option<f64> {
    attr_value(e, key).and_then(|value| value.parse::<f64>().ok())
}

fn parse_sheet_state(value: Option<&str>) -> SheetState {
    match value {
        Some("hidden") => SheetState::Hidden,
        Some("veryHidden") => SheetState::VeryHidden,
        _ => SheetState::Visible,
    }
}

fn normalize_zip_path(path: &str) -> String {
    let mut stack = Vec::new();
    for part in path.split('/') {
        if part.is_empty() || part == "." {
            continue;
        }
        if part == ".." {
            stack.pop();
            continue;
        }
        stack.push(part);
    }
    stack.join("/")
}

fn join_and_normalize(base_dir: &str, target: &str) -> String {
    let t = target.trim_start_matches('/');
    let combined = if t.starts_with("xl/") {
        t.to_string()
    } else {
        format!("{base_dir}{t}")
    };
    normalize_zip_path(&combined)
}

fn part_dir(path: &str) -> String {
    let normalized = normalize_zip_path(path);
    normalized
        .rsplit_once('/')
        .map(|(dir, _)| format!("{dir}/"))
        .unwrap_or_default()
}

fn sheet_rels_path(sheet_path: &str) -> String {
    let normalized = normalize_zip_path(sheet_path);
    let Some((dir, file)) = normalized.rsplit_once('/') else {
        return format!("_rels/{normalized}.rels");
    };
    format!("{dir}/_rels/{file}.rels")
}

fn a1_to_row_col(coord: &str) -> Option<(u32, u32)> {
    let mut col = 0u32;
    let mut row = 0u32;
    let mut saw_digit = false;
    for ch in coord.chars() {
        if ch == '$' {
            continue;
        }
        if ch.is_ascii_alphabetic() && !saw_digit {
            col = col
                .checked_mul(26)?
                .checked_add((ch.to_ascii_uppercase() as u8 - b'A' + 1) as u32)?;
        } else if ch.is_ascii_digit() {
            saw_digit = true;
            row = row.checked_mul(10)?.checked_add(ch.to_digit(10)?)?;
        } else {
            return None;
        }
    }
    if row == 0 || col == 0 {
        None
    } else {
        Some((row, col))
    }
}

fn a1_bytes_to_row_col(coord: &[u8]) -> Option<(u32, u32)> {
    let mut col = 0u32;
    let mut row = 0u32;
    let mut saw_digit = false;
    for byte in coord {
        match byte {
            b'$' => {}
            b'a'..=b'z' | b'A'..=b'Z' if !saw_digit => {
                let upper = byte.to_ascii_uppercase();
                col = col
                    .checked_mul(26)?
                    .checked_add((upper - b'A' + 1) as u32)?;
            }
            b'0'..=b'9' => {
                saw_digit = true;
                row = row.checked_mul(10)?.checked_add((byte - b'0') as u32)?;
            }
            _ => return None,
        }
    }
    if row == 0 || col == 0 {
        None
    } else {
        Some((row, col))
    }
}

pub(crate) fn row_col_to_a1(row: u32, col: u32) -> String {
    let mut c = col;
    let mut letters = Vec::new();
    while c > 0 {
        c -= 1;
        letters.push((b'A' + (c % 26) as u8) as char);
        c /= 26;
    }
    letters.reverse();
    format!("{}{row}", letters.into_iter().collect::<String>())
}

#[cfg(test)]
mod tests {
    use super::*;

    const XLSX_BYTES: &[u8] = include_bytes!("../../../tests/fixtures/sprint_kappa_smoke.xlsx");

    #[test]
    fn opens_committed_xlsx_fixture() {
        let book = NativeXlsxBook::open_bytes(XLSX_BYTES).expect("fixture opens");
        assert!(!book.sheet_names().is_empty());
        let first = book.sheet_names()[0].to_string();
        let sheet = book.worksheet(&first).expect("first sheet parses");
        assert!(!sheet.cells.is_empty(), "fixture should have cells");
    }

    #[test]
    fn one_shot_value_reader_opens_committed_xlsx_fixture() {
        let cells = read_first_sheet_value_cells_bytes(XLSX_BYTES).expect("fixture values parse");

        assert!(!cells.is_empty(), "fixture should have value cells");
        assert!(
            cells
                .iter()
                .any(|cell| matches!(cell.value, CellValue::Number(_))),
            "fixture should expose numeric values"
        );
    }

    #[test]
    fn value_summary_counts_rows_and_numeric_checksum_without_decoding_strings() {
        let xml = br#"<worksheet><sheetData>
            <row r="1">
                <c r="A1"><v>1</v></c>
                <c r="B1" t="s"><v>0</v></c>
                <c r="C1" t="b"><v>1</v></c>
            </row>
            <row r="2">
                <c r="A2"><v>2.5</v></c>
                <c r="B2" t="inlineStr"><is><t>Inline</t></is></c>
                <c r="C2" t="str"><v>not counted</v></c>
            </row>
        </sheetData></worksheet>"#;

        let summary = parse_worksheet_value_summary_bytes(xml).expect("summary parses");

        assert_eq!(summary.rows, 2);
        assert_eq!(summary.numeric_checksum, 3.5);
    }

    #[test]
    fn formula_summary_counts_formulas_and_first_column_numbers() {
        let xml = br#"<worksheet><sheetData>
            <row r="1">
                <c r="A1"><v>1</v></c>
                <c r="B1"><f>A1*2</f><v>2</v></c>
            </row>
            <row r="2">
                <c r="A2"><v>2.5</v></c>
                <c r="B2"><f>A2*2</f><v>5</v></c>
            </row>
        </sheetData></worksheet>"#;

        let summary = parse_worksheet_formula_summary_bytes(xml).expect("summary parses");

        assert_eq!(summary.rows, 2);
        assert_eq!(summary.first_col_numeric_checksum, 3.5);
        assert_eq!(summary.formula_cells, 2);
    }

    #[test]
    fn parses_workbook_sheet_order_and_state() {
        let xml = r#"<workbook xmlns:r="r">
        <fileSharing readOnlyRecommended="1" userName="Wolf" algorithmName="SHA-512" hashValue="FILEHASH" saltValue="FILESALT" spinCount="100000"/>
        <workbookPr date1904="1" codeName="ThisWorkbook" defaultThemeVersion="164011"/>
        <workbookProtection lockStructure="1" workbookAlgorithmName="SHA-512" workbookHashValue="HASH" workbookSaltValue="SALT" workbookSpinCount="100000"/>
        <bookViews><workbookView visibility="hidden" minimized="1" showHorizontalScroll="0" showVerticalScroll="0" showSheetTabs="0" xWindow="10" yWindow="20" windowWidth="12000" windowHeight="8000" tabRatio="750" firstSheet="1" activeTab="2" autoFilterDateGrouping="0"/></bookViews>
        <calcPr calcId="191029" calcMode="manual" fullCalcOnLoad="1" iterate="1" iterateCount="25" iterateDelta="0.01" forceFullCalc="1"/>
        <sheets>
            <sheet name="Visible" sheetId="1" r:id="rId1"/>
            <sheet name="Hidden" sheetId="2" state="hidden" r:id="rId2"/>
            <sheet name="Very" sheetId="3" state="veryHidden" r:id="rId3"/>
        </sheets><definedNames>
            <definedName name="GlobalName">Visible!$A$1</definedName>
            <definedName name="LocalName" localSheetId="1">$B$2</definedName>
            <definedName name="_xlnm.Print_Area">Visible!$A$1:$B$2</definedName>
            <definedName name="_xlnm.Print_Titles" localSheetId="0">Visible!$1:$2,Visible!$A:$B</definedName>
        </definedNames></workbook>"#;
        let (
            sheets,
            date1904,
            named_ranges,
            print_areas,
            print_titles,
            security,
            workbook_properties,
            calc_properties,
            workbook_views,
        ) = parse_workbook(xml).expect("parse workbook");
        assert!(date1904);
        let workbook_properties = workbook_properties.expect("workbookPr");
        assert!(workbook_properties.date1904);
        assert_eq!(
            workbook_properties.code_name.as_deref(),
            Some("ThisWorkbook")
        );
        assert_eq!(workbook_properties.default_theme_version, Some(164011));
        let calc_properties = calc_properties.expect("calcPr");
        assert_eq!(calc_properties.calc_id, Some(191029));
        assert_eq!(calc_properties.calc_mode.as_deref(), Some("manual"));
        assert_eq!(calc_properties.full_calc_on_load, Some(true));
        assert_eq!(calc_properties.iterate_count, Some(25));
        assert_eq!(calc_properties.iterate_delta, Some(0.01));
        assert_eq!(calc_properties.force_full_calc, Some(true));
        assert_eq!(workbook_views.len(), 1);
        assert_eq!(workbook_views[0].visibility, "hidden");
        assert!(workbook_views[0].minimized);
        assert!(!workbook_views[0].show_horizontal_scroll);
        assert_eq!(workbook_views[0].active_tab, 2);
        assert_eq!(sheets[0].name, "Visible");
        assert_eq!(sheets[1].state, SheetState::Hidden);
        assert_eq!(sheets[2].state, SheetState::VeryHidden);
        assert_eq!(
            security.workbook_protection,
            Some(WorkbookProtection {
                lock_structure: true,
                lock_windows: false,
                lock_revision: false,
                workbook_password: None,
                workbook_password_character_set: None,
                workbook_algorithm_name: Some("SHA-512".to_string()),
                workbook_hash_value: Some("HASH".to_string()),
                workbook_salt_value: Some("SALT".to_string()),
                workbook_spin_count: Some(100000),
                revisions_password: None,
                revisions_password_character_set: None,
                revisions_algorithm_name: None,
                revisions_hash_value: None,
                revisions_salt_value: None,
                revisions_spin_count: None,
            })
        );
        assert_eq!(
            security.file_sharing,
            Some(FileSharing {
                read_only_recommended: true,
                user_name: Some("Wolf".to_string()),
                reservation_password: None,
                algorithm_name: Some("SHA-512".to_string()),
                hash_value: Some("FILEHASH".to_string()),
                salt_value: Some("FILESALT".to_string()),
                spin_count: Some(100000),
            })
        );
        assert_eq!(
            named_ranges,
            vec![
                NamedRange {
                    name: "GlobalName".to_string(),
                    scope: "workbook".to_string(),
                    refers_to: "Visible!$A$1".to_string(),
                    ..Default::default()
                },
                NamedRange {
                    name: "LocalName".to_string(),
                    scope: "sheet".to_string(),
                    local_sheet_id: Some(1),
                    refers_to: "$B$2".to_string(),
                    ..Default::default()
                },
            ]
        );
        assert_eq!(
            print_areas.get("Visible").map(String::as_str),
            Some("Visible!$A$1:$B$2")
        );
        assert_eq!(
            print_titles.get("Visible"),
            Some(&PrintTitlesInfo {
                rows: Some("1:2".to_string()),
                cols: Some("A:B".to_string()),
            })
        );
    }

    #[test]
    fn synthesizes_sheet_refs_from_workbook_rels() {
        let rels = RelsGraph::parse(
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
                <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
                <Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
            </Relationships>"#,
        )
        .expect("parse rels");

        let sheet_refs = synthesize_sheet_refs_from_rels(&rels);

        assert_eq!(sheet_refs.len(), 2);
        assert_eq!(sheet_refs[0].name, "Sheet1");
        assert_eq!(sheet_refs[0].sheet_id.as_deref(), Some("1"));
        assert_eq!(sheet_refs[0].rid, "rId3");
        assert_eq!(sheet_refs[1].name, "Sheet2");
        assert_eq!(sheet_refs[1].sheet_id.as_deref(), Some("2"));
        assert_eq!(sheet_refs[1].rid, "rId9");
    }

    #[test]
    fn parses_shared_strings_plain_and_rich_text() {
        let xml = r#"<sst><si><t>Plain</t></si><si><r><rPr><b/><color rgb="FFFF0000"/></rPr><t>Rich&#13;&#10;</t></r><r><t>Text</t></r></si></sst>"#;
        let strings = parse_shared_strings(xml).expect("parse sst");
        assert_eq!(strings.values, vec!["Plain", "Rich\nText"]);
        assert_eq!(strings.rich_text[0], None);
        let rich = strings.rich_text[1].as_ref().expect("rich runs");
        assert_eq!(
            rich.iter().map(|run| run.text.as_str()).collect::<Vec<_>>(),
            vec!["Rich\n", "Text"]
        );
        assert_eq!(rich[0].font.as_ref().and_then(|font| font.bold), Some(true));
        assert_eq!(
            rich[0].font.as_ref().and_then(|font| font.color.clone()),
            Some("FFFF0000".to_string())
        );
    }

    #[test]
    fn parses_custom_and_builtin_number_formats() {
        let xml = r#"<styleSheet>
            <numFmts count="1"><numFmt numFmtId="165" formatCode="$#,##0.00"/></numFmts>
            <borders count="2">
                <border><left/><right/><top/><bottom/><diagonal/></border>
                <border diagonalUp="1"><left style="thin"><color rgb="FFFF0000"/></left><right style="medium"/><top style="double"/><bottom style="dashed"><color indexed="4"/></bottom><diagonal style="hair"><color rgb="FF00FF00"/></diagonal></border>
            </borders>
            <cellXfs count="3">
                <xf numFmtId="0"/>
                <xf numFmtId="4"/>
                <xf numFmtId="165" borderId="1"/>
            </cellXfs>
        </styleSheet>"#;
        let styles = parse_style_tables(xml).expect("parse styles");
        assert_eq!(styles.number_format_for_style_id(0), None);
        assert_eq!(styles.number_format_for_style_id(1), Some("#,##0.00"));
        assert_eq!(styles.number_format_for_style_id(2), Some("$#,##0.00"));
        let border = styles.border_for_style_id(2).expect("border");
        assert_eq!(
            border.left,
            Some(BorderSide {
                style: "thin".to_string(),
                color: "#FF0000".to_string(),
            })
        );
        assert_eq!(
            border.right,
            Some(BorderSide {
                style: "medium".to_string(),
                color: "#000000".to_string(),
            })
        );
        assert_eq!(
            border.bottom,
            Some(BorderSide {
                style: "dashed".to_string(),
                color: "#0000FF".to_string(),
            })
        );
        assert_eq!(
            border.diagonal_up,
            Some(BorderSide {
                style: "hair".to_string(),
                color: "#00FF00".to_string(),
            })
        );
        assert_eq!(border.diagonal_down, None);
    }

    #[test]
    fn parses_sheet_values_formulas_and_types() {
        let xml = r#"<worksheet><dimension ref="A1:D2"/><sheetViews><sheetView>
            <pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>
        </sheetView></sheetViews><cols><col min="2" max="3" width="18.5" customWidth="1" hidden="1" outlineLevel="2"/></cols><sheetData>
            <row r="1">
                <c r="A1" t="s"><v>0</v></c>
                <c r="B1"><v>42</v></c>
                <c r="C1" t="b"><v>1</v></c>
                <c r="D1"><f>SUM(B1:B1)</f><v>42</v></c>
            </row>
            <row r="2" ht="24" customHeight="1" hidden="1" outlineLevel="1"><c r="A2" t="inlineStr"><is><r><rPr><i/></rPr><t>In</t></r><r><t>line</t></r></is></c></row>
        </sheetData><mergeCells count="1"><mergeCell ref="A3:B3"/></mergeCells>
        <dataValidations count="1">
            <dataValidation type="whole" operator="between" allowBlank="1" sqref="B2:B5" errorTitle="Invalid" error="Use 1-10">
                <formula1>1</formula1><formula2>10</formula2>
            </dataValidation>
        </dataValidations>
        <conditionalFormatting sqref="C2:C5">
            <cfRule type="cellIs" operator="greaterThan" priority="1" stopIfTrue="1"><formula>50</formula></cfRule>
        </conditionalFormatting>
        <sheetProtection sheet="1" objects="1" formatCells="0" sort="0" password="C258"/>
        <autoFilter ref="A1:D10">
            <filterColumn colId="0"><filters><filter val="Label"/></filters></filterColumn>
            <filterColumn colId="1"><customFilters and="1"><customFilter operator="greaterThan" val="10"/></customFilters></filterColumn>
            <sortState ref="A2:D10"><sortCondition ref="B2:B10" descending="1"/></sortState>
        </autoFilter>
        </worksheet>"#;
        let shared_strings = SharedStrings {
            values: vec!["Shared".to_string()],
            rich_text: vec![None],
        };
        let sheet = parse_worksheet(xml, &shared_strings, None, Vec::new(), Vec::new())
            .expect("parse worksheet");
        assert_eq!(sheet.dimension.as_deref(), Some("A1:D2"));
        assert_eq!(sheet.merged_ranges, vec!["A3:B3"]);
        assert_eq!(
            sheet.freeze_panes,
            Some(FreezePane {
                mode: PaneMode::Freeze,
                top_left_cell: Some("B2".to_string()),
                x_split: Some(1),
                y_split: Some(1),
                active_pane: Some("bottomRight".to_string()),
            })
        );
        assert_eq!(
            sheet.row_heights.get(&2),
            Some(&RowHeight {
                height: 24.0,
                custom_height: true,
            })
        );
        assert_eq!(
            sheet.column_widths,
            vec![ColumnWidth {
                min: 2,
                max: 3,
                width: 18.5,
                custom_width: true,
            }]
        );
        assert_eq!(sheet.hidden_rows, vec![2]);
        assert_eq!(sheet.hidden_columns, vec![2, 3]);
        assert_eq!(sheet.row_outline_levels, vec![(2, 1)]);
        assert_eq!(sheet.column_outline_levels, vec![(2, 2), (3, 2)]);
        assert_eq!(
            sheet.data_validations,
            vec![DataValidation {
                range: "B2:B5".to_string(),
                validation_type: "whole".to_string(),
                operator: Some("between".to_string()),
                formula1: Some("1".to_string()),
                formula2: Some("10".to_string()),
                allow_blank: true,
                show_input_message: false,
                show_error_message: false,
                error_title: Some("Invalid".to_string()),
                error: Some("Use 1-10".to_string()),
                prompt_title: None,
                prompt: None,
            }]
        );
        assert_eq!(
            sheet.conditional_formats,
            vec![ConditionalFormatRule {
                range: "C2:C5".to_string(),
                rule_type: "cellIs".to_string(),
                operator: Some("greaterThan".to_string()),
                formula: Some("50".to_string()),
                dxf_id: None,
                priority: Some(1),
                stop_if_true: Some(true),
                color_scale: None,
            }]
        );
        assert_eq!(
            sheet.sheet_protection,
            Some(SheetProtection {
                sheet: true,
                objects: true,
                scenarios: false,
                format_cells: false,
                format_columns: true,
                format_rows: true,
                insert_columns: true,
                insert_rows: true,
                insert_hyperlinks: true,
                delete_columns: true,
                delete_rows: true,
                select_locked_cells: false,
                sort: false,
                auto_filter: true,
                pivot_tables: true,
                select_unlocked_cells: false,
                password_hash: Some("C258".to_string()),
            })
        );
        assert_eq!(
            sheet.auto_filter,
            Some(AutoFilterInfo {
                ref_range: "A1:D10".to_string(),
                filter_columns: vec![
                    FilterColumnInfo {
                        col_id: 0,
                        hidden_button: false,
                        show_button: true,
                        filter: Some(FilterInfo::String {
                            values: vec!["Label".to_string()],
                        }),
                        date_group_items: Vec::new(),
                    },
                    FilterColumnInfo {
                        col_id: 1,
                        hidden_button: false,
                        show_button: true,
                        filter: Some(FilterInfo::Custom {
                            and_: true,
                            filters: vec![CustomFilterInfo {
                                operator: "greaterThan".to_string(),
                                val: "10".to_string(),
                            }],
                        }),
                        date_group_items: Vec::new(),
                    },
                ],
                sort_state: Some(SortStateInfo {
                    sort_conditions: vec![SortConditionInfo {
                        ref_range: "B2:B10".to_string(),
                        descending: true,
                        sort_by: "value".to_string(),
                        custom_list: None,
                        dxf_id: None,
                        icon_set: None,
                        icon_id: None,
                    }],
                    column_sort: false,
                    case_sensitive: false,
                    ref_range: Some("A2:D10".to_string()),
                }),
            })
        );
        assert_eq!(
            sheet.cells[0].value,
            CellValue::String("Shared".to_string())
        );
        assert_eq!(sheet.cells[1].value, CellValue::Number(42.0));
        assert_eq!(sheet.cells[2].value, CellValue::Bool(true));
        assert_eq!(sheet.cells[3].formula.as_deref(), Some("SUM(B1:B1)"));
        assert_eq!(
            sheet.cells[4].value,
            CellValue::String("Inline".to_string())
        );
        let inline_rich = sheet.cells[4].rich_text.as_ref().expect("inline rich text");
        assert_eq!(
            inline_rich
                .iter()
                .map(|run| run.text.as_str())
                .collect::<Vec<_>>(),
            vec!["In", "line"]
        );
        assert_eq!(
            inline_rich[0].font.as_ref().and_then(|font| font.italic),
            Some(true)
        );
    }

    #[test]
    fn parses_value_only_cells_without_later_sheet_features() {
        let xml = r#"<worksheet><sheetData>
            <row r="1">
                <c r="A1" t="s"><v>0</v></c>
                <c r="B1"><v>42</v></c>
                <c r="C1" t="b"><v>1</v></c>
            </row>
            <row r="2">
                <c r="A2" t="inlineStr"><is><t>Inline</t></is></c>
            </row>
        </sheetData>
        <dataValidations count="1">
            <dataValidation type="whole" sqref="B2:B5"><formula1>1</formula1></dataValidation>
        </dataValidations>
        <conditionalFormatting sqref="C2:C5">
            <cfRule type="cellIs" priority="1"><formula>50</formula></cfRule>
        </conditionalFormatting></worksheet>"#;
        let shared_strings = SharedStrings {
            values: vec!["Shared".to_string()],
            rich_text: vec![None],
        };
        let cells =
            parse_worksheet_value_cells(xml, &shared_strings).expect("parse value-only cells");

        assert_eq!(
            cells,
            vec![
                ValueCell {
                    row: 1,
                    col: 1,
                    value: CellValue::String("Shared".to_string()),
                },
                ValueCell {
                    row: 1,
                    col: 2,
                    value: CellValue::Number(42.0),
                },
                ValueCell {
                    row: 1,
                    col: 3,
                    value: CellValue::Bool(true),
                },
                ValueCell {
                    row: 2,
                    col: 1,
                    value: CellValue::String("Inline".to_string()),
                },
            ]
        );
    }

    #[test]
    fn value_only_inline_strings_keep_escape_and_line_break_parity() {
        let xml = "<worksheet><sheetData><row r=\"1\">\
            <c r=\"A1\" t=\"inlineStr\"><is><t>AT&amp;T</t></is></c>\
            <c r=\"B1\" t=\"inlineStr\"><is><t>one\r\ntwo</t></is></c>\
        </row></sheetData></worksheet>";
        let cells =
            parse_worksheet_value_cells(xml, &SharedStrings::default()).expect("parse value cells");

        assert_eq!(cells[0].value, CellValue::String("AT&T".to_string()));
        assert_eq!(cells[1].value, CellValue::String("one\ntwo".to_string()));
    }

    #[test]
    fn value_only_cells_preallocate_from_dimension() {
        let xml = r#"<worksheet><dimension ref="A1:C4"/><sheetData/></worksheet>"#;
        let cells = parse_worksheet_value_cells(xml, &SharedStrings::default())
            .expect("parse empty value-only sheet");

        assert!(cells.is_empty());
        assert_eq!(cells.capacity(), 12);
    }

    #[test]
    fn value_only_dimension_preallocation_is_capped() {
        let xml = r#"<worksheet><dimension ref="A1:XFD1048576"/><sheetData/></worksheet>"#;
        let cells = parse_worksheet_value_cells(xml, &SharedStrings::default())
            .expect("parse empty huge value-only sheet");

        assert!(cells.is_empty());
        assert_eq!(cells.capacity(), 0);
    }

    #[test]
    fn parses_iso_date_cells_as_string_values() {
        let xml = r#"<worksheet><sheetData>
            <row r="1">
                <c r="A1" t="d"><v>2021-01-01</v></c>
                <c r="A2" t="d"><v>2021-01-01T10:10:10</v></c>
                <c r="A3" t="d"><v>10:10:10</v></c>
            </row>
        </sheetData></worksheet>"#;
        let sheet = parse_worksheet(xml, &SharedStrings::default(), None, Vec::new(), Vec::new())
            .expect("parse worksheet");

        assert_eq!(
            sheet
                .cells
                .iter()
                .map(|cell| cell.value.clone())
                .collect::<Vec<_>>(),
            vec![
                CellValue::String("2021-01-01".to_string()),
                CellValue::String("2021-01-01T10:10:10".to_string()),
                CellValue::String("10:10:10".to_string()),
            ]
        );
    }

    #[test]
    fn expands_shared_formula_children() {
        let xml = r#"<worksheet><sheetData>
            <row r="1">
                <c r="A1"><v>1</v></c>
                <c r="B1"><f t="shared" si="0" ref="B1:B3">A1*2</f><v/></c>
                <c r="C1"><f t="shared" si="1" ref="C1:C3">$A$1+A1+B$1+$A1</f><v/></c>
            </row>
            <row r="2">
                <c r="A2"><v>2</v></c>
                <c r="B2"><f t="shared" si="0"/><v/></c>
                <c r="C2"><f t="shared" si="1"/><v/></c>
            </row>
            <row r="3">
                <c r="A3"><v>3</v></c>
                <c r="B3"><f t="shared" si="0"/><v/></c>
                <c r="C3"><f t="shared" si="1"/><v/></c>
            </row>
        </sheetData></worksheet>"#;

        let sheet = parse_worksheet(xml, &SharedStrings::default(), None, Vec::new(), Vec::new())
            .expect("parse worksheet");
        let formulas: HashMap<_, _> = sheet
            .cells
            .iter()
            .filter_map(|cell| {
                cell.formula
                    .as_deref()
                    .map(|formula| (cell.coordinate.as_str(), formula))
            })
            .collect();

        assert_eq!(formulas.get("B1"), Some(&"A1*2"));
        assert_eq!(formulas.get("B2"), Some(&"A2*2"));
        assert_eq!(formulas.get("B3"), Some(&"A3*2"));
        assert_eq!(formulas.get("C1"), Some(&"$A$1+A1+B$1+$A1"));
        assert_eq!(formulas.get("C2"), Some(&"$A$1+A2+B$1+$A2"));
        assert_eq!(formulas.get("C3"), Some(&"$A$1+A3+B$1+$A3"));
    }

    #[test]
    fn parses_sheet_hyperlinks_with_relationship_targets() {
        let xml = r#"<worksheet xmlns:r="r"><sheetData><row r="1">
            <c r="A1" t="inlineStr"><is><t>Website</t></is></c>
            <c r="B1" t="inlineStr"><is><t>Internal</t></is></c>
        </row></sheetData><hyperlinks>
            <hyperlink ref="A1" r:id="rId1" tooltip="External site"/>
            <hyperlink ref="B1" location="Other!A1" display="Jump"/>
        </hyperlinks></worksheet>"#;
        let rels = RelsGraph::parse(
            br#"<Relationships>
                <Relationship Id="rId1" Type="hyperlink" Target="https://example.com" TargetMode="External"/>
            </Relationships>"#,
        )
        .expect("parse rels");

        let sheet = parse_worksheet(
            xml,
            &SharedStrings::default(),
            Some(&rels),
            Vec::new(),
            Vec::new(),
        )
        .expect("parse worksheet");
        assert_eq!(
            sheet.hyperlinks,
            vec![
                Hyperlink {
                    cell: "A1".to_string(),
                    target: "https://example.com".to_string(),
                    location: None,
                    display: String::new(),
                    tooltip: Some("External site".to_string()),
                    internal: false,
                },
                Hyperlink {
                    cell: "B1".to_string(),
                    target: "Other!A1".to_string(),
                    location: Some("Other!A1".to_string()),
                    display: "Jump".to_string(),
                    tooltip: None,
                    internal: true,
                },
            ]
        );
    }

    #[test]
    fn parses_comments_authors_and_rich_text() {
        let xml = r#"<comments><authors><author>Alice</author><author>Bob</author></authors>
            <commentList>
                <comment ref="A1" authorId="0"><text><t>First note</t></text></comment>
                <comment ref="B2" authorId="1"><text><r><t>Second</t></r><r><t> note</t></r></text></comment>
            </commentList></comments>"#;

        assert_eq!(
            parse_comments(xml).expect("parse comments"),
            vec![
                Comment {
                    cell: "A1".to_string(),
                    text: "First note".to_string(),
                    author: "Alice".to_string(),
                    threaded: false,
                },
                Comment {
                    cell: "B2".to_string(),
                    text: "Second note".to_string(),
                    author: "Bob".to_string(),
                    threaded: false,
                },
            ]
        );
    }

    #[test]
    fn parses_table_metadata() {
        let xml = r#"<table name="SalesTable" displayName="SalesTable" ref="A1:B3" headerRowCount="1" totalsRowCount="0" totalsRowShown="1" comment="table comment" tableType="worksheet">
            <autoFilter ref="A1:B3"/>
            <tableColumns count="2"><tableColumn id="1" name="Name"/><tableColumn id="2" name="Sales"/></tableColumns>
            <tableStyleInfo name="TableStyleLight9" showFirstColumn="1" showLastColumn="1" showRowStripes="1" showColumnStripes="1"/>
        </table>"#;

        assert_eq!(
            parse_table_xml(xml).expect("parse table"),
            Table {
                name: "SalesTable".to_string(),
                ref_range: "A1:B3".to_string(),
                header_row: true,
                totals_row: false,
                comment: Some("table comment".to_string()),
                table_type: Some("worksheet".to_string()),
                totals_row_shown: Some(true),
                style: Some("TableStyleLight9".to_string()),
                show_first_column: true,
                show_last_column: true,
                show_row_stripes: true,
                show_column_stripes: true,
                columns: vec!["Name".to_string(), "Sales".to_string()],
                auto_filter: Some(AutoFilterInfo {
                    ref_range: "A1:B3".to_string(),
                    filter_columns: Vec::new(),
                    sort_state: None,
                }),
            }
        );
    }

    #[test]
    fn parses_chart_axis_titles_separately_from_chart_title() {
        let xml = r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:chart>
                <c:title><c:tx><c:rich><a:p><a:r><a:t>Sales Trend</a:t></a:r></a:p></c:rich></c:tx></c:title>
                <c:plotArea>
                    <c:barChart>
                        <c:barDir val="bar"/>
                        <c:grouping val="stacked"/>
                        <c:varyColors val="1"/>
                        <c:ser>
                            <c:idx val="0"/><c:order val="0"/>
                            <c:tx><c:strRef><c:f>'Charts'!B1</c:f></c:strRef></c:tx>
                            <c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:ln w="20000"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill><a:prstDash val="dash"/></a:ln></c:spPr>
                            <c:dLbls><c:dLblPos val="outEnd"/><c:showVal val="1"/></c:dLbls>
                            <c:trendline><c:trendlineType val="poly"/><c:order val="3"/><c:dispEq val="1"/><c:dispRSqr val="1"/></c:trendline>
                            <c:errBars><c:errBarType val="both"/><c:errValType val="fixedVal"/><c:noEndCap val="1"/><c:val val="2"/></c:errBars>
                            <c:cat><c:strRef><c:f>'Charts'!$A$2:$A$4</c:f></c:strRef></c:cat>
                            <c:val><c:numRef><c:f>'Charts'!$B$2:$B$4</c:f></c:numRef></c:val>
                        </c:ser>
                        <c:dLbls><c:dLblPos val="outEnd"/><c:showVal val="1"/></c:dLbls>
                    </c:barChart>
                    <c:catAx>
                        <c:axId val="10"/><c:axPos val="b"/><c:tickLblPos val="low"/>
                        <c:title><c:tx><c:rich><a:p><a:r><a:t>Month</a:t></a:r></a:p></c:rich></c:tx></c:title>
                        <c:crossAx val="100"/>
                    </c:catAx>
                    <c:valAx>
                        <c:axId val="100"/>
                        <c:scaling><c:orientation val="minMax"/><c:min val="0"/><c:max val="40"/></c:scaling>
                        <c:axPos val="l"/>
                        <c:title><c:tx><c:rich><a:p><a:r><a:t>Sales</a:t></a:r></a:p></c:rich></c:tx></c:title>
                        <c:numFmt formatCode="0.0" sourceLinked="0"/>
                        <c:majorTickMark val="out"/><c:minorTickMark val="in"/><c:tickLblPos val="high"/>
                        <c:crossAx val="10"/><c:crossBetween val="between"/>
                        <c:majorUnit val="10"/><c:minorUnit val="5"/>
                        <c:dispUnits><c:builtInUnit val="thousands"/></c:dispUnits>
                    </c:valAx>
                </c:plotArea>
                <c:legend><c:legendPos val="t"/></c:legend>
            </c:chart>
        </c:chartSpace>"#;
        let anchor = ImageAnchorInfo::Absolute {
            pos: AnchorPositionInfo::default(),
            ext: AnchorExtentInfo::default(),
        };

        let chart = parse_chart_xml(xml, anchor)
            .expect("parse chart")
            .expect("chart info");

        assert_eq!(chart.kind, "bar");
        assert_eq!(chart.title.as_deref(), Some("Sales Trend"));
        assert_eq!(chart.x_axis_title.as_deref(), Some("Month"));
        assert_eq!(chart.y_axis_title.as_deref(), Some("Sales"));
        assert_eq!(chart.legend_position.as_deref(), Some("t"));
        assert_eq!(chart.bar_dir.as_deref(), Some("bar"));
        assert_eq!(chart.grouping.as_deref(), Some("stacked"));
        assert_eq!(chart.vary_colors, Some(true));
        let x_axis = chart.x_axis.as_ref().expect("x axis metadata");
        assert_eq!(x_axis.axis_type, "cat");
        assert_eq!(x_axis.ax_id, Some(10));
        assert_eq!(x_axis.tick_lbl_pos.as_deref(), Some("low"));
        let y_axis = chart.y_axis.as_ref().expect("y axis metadata");
        assert_eq!(y_axis.axis_type, "val");
        assert_eq!(y_axis.ax_id, Some(100));
        assert_eq!(y_axis.cross_ax, Some(10));
        assert_eq!(y_axis.scaling_min, Some(0.0));
        assert_eq!(y_axis.scaling_max, Some(40.0));
        assert_eq!(y_axis.scaling_orientation.as_deref(), Some("minMax"));
        assert_eq!(y_axis.num_format_code.as_deref(), Some("0.0"));
        assert_eq!(y_axis.num_format_source_linked, Some(false));
        assert_eq!(y_axis.major_unit, Some(10.0));
        assert_eq!(y_axis.minor_unit, Some(5.0));
        assert_eq!(y_axis.tick_lbl_pos.as_deref(), Some("high"));
        assert_eq!(y_axis.major_tick_mark.as_deref(), Some("out"));
        assert_eq!(y_axis.minor_tick_mark.as_deref(), Some("in"));
        assert_eq!(y_axis.cross_between.as_deref(), Some("between"));
        assert_eq!(y_axis.display_unit.as_deref(), Some("thousands"));
        let labels = chart.data_labels.as_ref().expect("chart labels");
        assert_eq!(labels.position.as_deref(), Some("outEnd"));
        assert_eq!(labels.show_val, Some(true));
        let series = &chart.series[0];
        assert_eq!(series.title_ref.as_deref(), Some("'Charts'!B1"));
        let gp = series
            .graphical_properties
            .as_ref()
            .expect("series graphical properties");
        assert_eq!(gp.solid_fill.as_deref(), Some("FF0000"));
        assert_eq!(gp.line_solid_fill.as_deref(), Some("00FF00"));
        assert_eq!(gp.line_dash.as_deref(), Some("dash"));
        assert_eq!(gp.line_width, Some(20000));
        let series_labels = series.data_labels.as_ref().expect("series labels");
        assert_eq!(series_labels.position.as_deref(), Some("outEnd"));
        assert_eq!(series_labels.show_val, Some(true));
        let trendline = series.trendline.as_ref().expect("series trendline");
        assert_eq!(trendline.trendline_type.as_deref(), Some("poly"));
        assert_eq!(trendline.order, Some(3));
        assert_eq!(trendline.display_equation, Some(true));
        assert_eq!(trendline.display_r_squared, Some(true));
        let error_bars = series.error_bars.as_ref().expect("series error bars");
        assert_eq!(error_bars.bar_type.as_deref(), Some("both"));
        assert_eq!(error_bars.val_type.as_deref(), Some("fixedVal"));
        assert_eq!(error_bars.no_end_cap, Some(true));
        assert_eq!(error_bars.val, Some(2.0));
    }

    #[test]
    fn parses_document_properties() {
        let mut props = HashMap::new();
        parse_doc_properties_into(
            r#"<cp:coreProperties>
                <dc:title>Q3 Report</dc:title>
                <dc:creator>Alice</dc:creator>
                <cp:lastModifiedBy>Bob</cp:lastModifiedBy>
                <dcterms:created>2024-01-01T00:00:00Z</dcterms:created>
            </cp:coreProperties>"#,
            &mut props,
            doc_property_core_key,
        );
        parse_doc_properties_into(
            r#"<Properties><Application>Excel</Application><Company>SynthGL</Company></Properties>"#,
            &mut props,
            doc_property_app_key,
        );

        assert_eq!(props.get("title").map(String::as_str), Some("Q3 Report"));
        assert_eq!(props.get("creator").map(String::as_str), Some("Alice"));
        assert_eq!(props.get("lastModifiedBy").map(String::as_str), Some("Bob"));
        assert_eq!(
            props.get("created").map(String::as_str),
            Some("2024-01-01T00:00:00Z")
        );
        assert_eq!(props.get("application").map(String::as_str), Some("Excel"));
        assert_eq!(props.get("company").map(String::as_str), Some("SynthGL"));
    }

    #[test]
    fn parses_person_list_with_optional_fields() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <person displayName="Alice" id="{A}" userId="alice@x.com" providerId="AD"/>
  <person displayName="Bob" id="{B}"/>
</personList>"#;
        let persons = parse_person_list(xml).expect("parses");
        assert_eq!(persons.len(), 2);
        assert_eq!(persons[0].display_name, "Alice");
        assert_eq!(persons[0].id, "{A}");
        assert_eq!(persons[0].user_id.as_deref(), Some("alice@x.com"));
        assert_eq!(persons[0].provider_id.as_deref(), Some("AD"));
        assert_eq!(persons[1].display_name, "Bob");
        assert_eq!(persons[1].user_id, None);
        assert_eq!(persons[1].provider_id, None);
    }

    #[test]
    fn skips_persons_missing_id() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <person displayName="No GUID"/>
  <person displayName="Has GUID" id="{C}"/>
</personList>"#;
        let persons = parse_person_list(xml).expect("parses");
        assert_eq!(persons.len(), 1);
        assert_eq!(persons[0].id, "{C}");
    }

    #[test]
    fn parses_threaded_comments_top_level_and_replies() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <threadedComment ref="A1" dT="2026-05-03T12:00:00.000" personId="{A}" id="{T1}">
    <text>Looks wrong</text>
  </threadedComment>
  <threadedComment ref="A1" dT="2026-05-03T12:01:00.000" personId="{B}" id="{T2}" parentId="{T1}">
    <text>Why?</text>
  </threadedComment>
</ThreadedComments>"#;
        let entries = parse_threaded_comments(xml).expect("parses");
        assert_eq!(entries.len(), 2);
        assert_eq!(entries[0].id, "{T1}");
        assert_eq!(entries[0].cell, "A1");
        assert_eq!(entries[0].person_id, "{A}");
        assert_eq!(entries[0].text, "Looks wrong");
        assert_eq!(entries[0].parent_id, None);
        assert_eq!(
            entries[0].created.as_deref(),
            Some("2026-05-03T12:00:00.000")
        );
        assert_eq!(entries[1].parent_id.as_deref(), Some("{T1}"));
        assert_eq!(entries[1].text, "Why?");
    }

    #[test]
    fn parses_threaded_comment_with_xml_escaped_text() {
        let xml = r#"<?xml version="1.0"?>
<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <threadedComment ref="B2" personId="{A}" id="{T1}">
    <text>&lt;b&gt; &amp; "hi"</text>
  </threadedComment>
</ThreadedComments>"#;
        let entries = parse_threaded_comments(xml).expect("parses");
        assert_eq!(entries.len(), 1);
        assert_eq!(entries[0].text, "<b> & \"hi\"");
    }

    #[test]
    fn parses_threaded_comment_done_flag() {
        let xml = r#"<?xml version="1.0"?>
<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <threadedComment ref="A1" personId="{A}" id="{T1}" done="1"><text>resolved</text></threadedComment>
  <threadedComment ref="B1" personId="{A}" id="{T2}"><text>open</text></threadedComment>
</ThreadedComments>"#;
        let entries = parse_threaded_comments(xml).expect("parses");
        assert_eq!(entries[0].done, true);
        assert_eq!(entries[1].done, false);
    }
}
