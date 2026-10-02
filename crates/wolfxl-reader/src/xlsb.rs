mod formula;
mod style_records;
mod worksheet_features;
mod worksheet_meta;

use std::collections::HashMap;
use std::fs;
use std::io::{Cursor, Read};
use std::path::Path;

use wolfxl_rels::{RelId, RelsGraph};
use zip::ZipArchive;

use crate::{
    row_col_to_a1, AlignmentInfo, AutoFilterInfo, BorderInfo, Cell, CellDataType, CellValue,
    ChartInfo, ColumnWidth, Comment, ConditionalFormatRule, DataValidation, FillInfo,
    FilterColumnInfo, FilterInfo, FontInfo, FreezePane, HeaderFooterInfo, Hyperlink, ImageInfo,
    NamedRange, PageBreakListInfo, PageMarginsInfo, PageSetupInfo, PrintOptionsInfo,
    PrintTitlesInfo, RowHeight, SheetFormatInfo, SheetPropertiesInfo, SheetProtection, SheetState,
    SheetViewInfo, StyleTables, Table, WorksheetData,
};

type Result<T> = std::result::Result<T, XlsbError>;

#[derive(Debug)]
pub enum XlsbError {
    Io(std::io::Error),
    Zip(zip::result::ZipError),
    Xml(String),
    Format(String),
    SheetNotFound(String),
}

impl std::fmt::Display for XlsbError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Io(e) => write!(f, "I/O error: {e}"),
            Self::Zip(e) => write!(f, "ZIP error: {e}"),
            Self::Xml(e) => write!(f, "XML error: {e}"),
            Self::Format(e) => write!(f, "XLSB format error: {e}"),
            Self::SheetNotFound(name) => write!(f, "sheet not found: {name}"),
        }
    }
}

impl std::error::Error for XlsbError {}

impl From<std::io::Error> for XlsbError {
    fn from(value: std::io::Error) -> Self {
        Self::Io(value)
    }
}

impl From<zip::result::ZipError> for XlsbError {
    fn from(value: zip::result::ZipError) -> Self {
        Self::Zip(value)
    }
}

#[derive(Debug, Clone)]
pub struct NativeXlsbBook {
    bytes: Vec<u8>,
    sheets: Vec<XlsbSheet>,
    named_ranges: Vec<NamedRange>,
    print_areas: HashMap<String, String>,
    print_titles: HashMap<String, PrintTitlesInfo>,
    extern_sheets: Vec<XtiRef>,
    formula_names: Vec<String>,
    shared_strings: Vec<String>,
    styles: StyleTables,
    date1904: bool,
}

#[derive(Debug, Clone)]
struct XlsbSheet {
    name: String,
    path: String,
    state: SheetState,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct XtiRef {
    first_sheet: i32,
    last_sheet: i32,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct RawNamedRange {
    name: String,
    local_id: Option<usize>,
    refers_to: String,
}

#[derive(Debug)]
struct Record<'a> {
    typ: u16,
    payload: &'a [u8],
}

impl NativeXlsbBook {
    pub fn open_path(path: impl AsRef<Path>) -> Result<Self> {
        Self::open_bytes(fs::read(path)?)
    }

    pub fn open_bytes(bytes: Vec<u8>) -> Result<Self> {
        let mut zip = ZipArchive::new(Cursor::new(bytes.clone()))?;
        let workbook_rels = read_rels(&mut zip, "xl/_rels/workbook.bin.rels")?;
        let shared_strings = read_shared_strings(&mut zip)?;
        let styles = style_records::read_styles(&mut zip)?;
        let (
            sheets,
            named_ranges,
            print_areas,
            print_titles,
            extern_sheets,
            formula_names,
            date1904,
        ) = read_workbook(&mut zip, &workbook_rels)?;
        Ok(Self {
            bytes,
            sheets,
            named_ranges,
            print_areas,
            print_titles,
            extern_sheets,
            formula_names,
            shared_strings,
            styles,
            date1904,
        })
    }

    pub fn sheet_names(&self) -> Vec<&str> {
        self.sheets.iter().map(|s| s.name.as_str()).collect()
    }

    pub fn sheet_state(&self, sheet_name: &str) -> Result<SheetState> {
        self.sheets
            .iter()
            .find(|s| s.name == sheet_name)
            .map(|s| s.state)
            .ok_or_else(|| XlsbError::SheetNotFound(sheet_name.to_string()))
    }

    pub fn date1904(&self) -> bool {
        self.date1904
    }

    pub fn named_ranges(&self) -> &[NamedRange] {
        &self.named_ranges
    }

    pub fn print_area(&self, sheet_name: &str) -> Option<&str> {
        self.print_areas.get(sheet_name).map(String::as_str)
    }

    pub fn print_titles(&self, sheet_name: &str) -> Option<&PrintTitlesInfo> {
        self.print_titles.get(sheet_name)
    }

    pub fn number_format_for_style_id(&self, style_id: u32) -> Option<&str> {
        self.styles.number_format_for_style_id(style_id)
    }

    pub fn border_for_style_id(&self, style_id: u32) -> Option<&BorderInfo> {
        self.styles.border_for_style_id(style_id)
    }

    pub fn font_for_style_id(&self, style_id: u32) -> Option<&FontInfo> {
        self.styles.font_for_style_id(style_id)
    }

    pub fn fill_for_style_id(&self, style_id: u32) -> Option<&FillInfo> {
        self.styles.fill_for_style_id(style_id)
    }

    pub fn alignment_for_style_id(&self, style_id: u32) -> Option<&AlignmentInfo> {
        self.styles.alignment_for_style_id(style_id)
    }

    pub fn worksheet(&self, sheet_name: &str) -> Result<WorksheetData> {
        let sheet = self
            .sheets
            .iter()
            .find(|s| s.name == sheet_name)
            .ok_or_else(|| XlsbError::SheetNotFound(sheet_name.to_string()))?;
        let mut zip = ZipArchive::new(Cursor::new(self.bytes.clone()))?;
        let data = read_zip_part(&mut zip, &sheet.path)?;
        let rels = read_zip_part_optional(&mut zip, &sheet_rels_path(&sheet.path))?
            .map(|xml| {
                RelsGraph::parse(&xml).map_err(|e| {
                    XlsbError::Xml(format!("failed to parse sheet rels for {sheet_name}: {e}"))
                })
            })
            .transpose()?;
        let comments = match rels.as_ref().and_then(worksheet_meta::comments_target) {
            Some(target) => read_zip_part_optional(
                &mut zip,
                &join_and_normalize(&part_dir(&sheet.path), &target),
            )?
            .map(|data| worksheet_meta::parse_comments_bin(&data))
            .transpose()?
            .unwrap_or_default(),
            None => Vec::new(),
        };
        let context = formula::FormulaContext {
            sheets: &self.sheets,
            extern_sheets: &self.extern_sheets,
            named_ranges: &self.named_ranges,
            formula_names: &self.formula_names,
        };
        let tables = read_tables_bin(&mut zip, &sheet.path, &data, rels.as_ref())?;
        let images = read_images_xml(&mut zip, &sheet.path, rels.as_ref())?;
        let charts = read_charts_xml(&mut zip, &sheet.path, rels.as_ref())?;
        let mut data = parse_worksheet_with_tables(
            &data,
            &self.shared_strings,
            rels.as_ref(),
            comments,
            tables,
            Some(&context),
        )?;
        data.images = images;
        data.charts = charts;
        Ok(data)
    }
}

fn read_images_xml(
    zip: &mut ZipArchive<Cursor<Vec<u8>>>,
    sheet_path: &str,
    rels: Option<&RelsGraph>,
) -> Result<Vec<ImageInfo>> {
    crate::read_images(zip, sheet_path, rels)
        .map_err(|e| XlsbError::Xml(format!("failed to read sheet drawings: {e}")))
}

fn read_charts_xml(
    zip: &mut ZipArchive<Cursor<Vec<u8>>>,
    sheet_path: &str,
    rels: Option<&RelsGraph>,
) -> Result<Vec<ChartInfo>> {
    crate::read_charts(zip, sheet_path, rels)
        .map_err(|e| XlsbError::Xml(format!("failed to read sheet charts: {e}")))
}

fn read_rels(zip: &mut ZipArchive<Cursor<Vec<u8>>>, path: &str) -> Result<RelsGraph> {
    let xml = read_zip_part(zip, path)?;
    RelsGraph::parse(&xml).map_err(|e| XlsbError::Xml(format!("failed to parse {path}: {e}")))
}

fn read_zip_part(zip: &mut ZipArchive<Cursor<Vec<u8>>>, path: &str) -> Result<Vec<u8>> {
    let name = zip_part_name(zip, path).unwrap_or_else(|| path.to_string());
    let mut file = zip.by_name(&name)?;
    let mut bytes = Vec::new();
    file.read_to_end(&mut bytes)?;
    Ok(bytes)
}

fn read_zip_part_optional(
    zip: &mut ZipArchive<Cursor<Vec<u8>>>,
    path: &str,
) -> Result<Option<Vec<u8>>> {
    let Some(name) = zip_part_name(zip, path) else {
        return Ok(None);
    };
    match zip.by_name(&name) {
        Ok(mut file) => {
            let mut bytes = Vec::new();
            file.read_to_end(&mut bytes)?;
            Ok(Some(bytes))
        }
        Err(zip::result::ZipError::FileNotFound) => Ok(None),
        Err(e) => Err(e.into()),
    }
}

fn zip_part_name(zip: &mut ZipArchive<Cursor<Vec<u8>>>, path: &str) -> Option<String> {
    if zip.by_name(path).is_ok() {
        return Some(path.to_string());
    }
    let needle = path.to_ascii_lowercase();
    zip.file_names()
        .find(|name| name.to_ascii_lowercase() == needle)
        .map(str::to_string)
}

fn sheet_rels_path(sheet_path: &str) -> String {
    let normalized = sheet_path.trim_start_matches('/');
    let (dir, file) = match normalized.rsplit_once('/') {
        Some((dir, file)) => (dir, file),
        None => ("", normalized),
    };
    if dir.is_empty() {
        format!("_rels/{file}.rels")
    } else {
        format!("{dir}/_rels/{file}.rels")
    }
}

fn part_dir(path: &str) -> String {
    path.trim_start_matches('/')
        .rsplit_once('/')
        .map(|(dir, _)| dir.to_string())
        .unwrap_or_default()
}

fn join_and_normalize(base_dir: &str, target: &str) -> String {
    let target = target.trim_start_matches('/');
    let raw = if target.starts_with("xl/") {
        target.to_string()
    } else if base_dir.is_empty() {
        target.to_string()
    } else {
        format!("{base_dir}/{target}")
    };
    let mut parts = Vec::new();
    for part in raw.split('/') {
        match part {
            "" | "." => {}
            ".." => {
                parts.pop();
            }
            value => parts.push(value),
        }
    }
    parts.join("/")
}

fn read_workbook(
    zip: &mut ZipArchive<Cursor<Vec<u8>>>,
    rels: &RelsGraph,
) -> Result<(
    Vec<XlsbSheet>,
    Vec<NamedRange>,
    HashMap<String, String>,
    HashMap<String, PrintTitlesInfo>,
    Vec<XtiRef>,
    Vec<String>,
    bool,
)> {
    let data = read_zip_part(zip, "xl/workbook.bin")?;
    let mut sheets = Vec::new();
    let mut raw_names = Vec::new();
    let mut raw_print_areas = Vec::new();
    let mut raw_print_titles = Vec::new();
    let mut extern_sheets = Vec::new();
    let mut formula_names = Vec::new();
    let mut date1904 = false;
    for record in Records::new(&data) {
        let record = record?;
        match record.typ {
            0x0099 => {
                date1904 = record
                    .payload
                    .first()
                    .is_some_and(|value| value & 0x01 != 0);
            }
            0x009c => {
                if record.payload.len() < 16 {
                    continue;
                }
                let state = match le_u32(&record.payload[0..4]) {
                    1 => SheetState::Hidden,
                    2 => SheetState::VeryHidden,
                    _ => SheetState::Visible,
                };
                let rel_len = le_u32(&record.payload[8..12]) as usize;
                let rel_start = 12;
                let rel_end = rel_start + rel_len * 2;
                if rel_end > record.payload.len() {
                    continue;
                }
                let rid = utf16_string(&record.payload[rel_start..rel_end]);
                let mut consumed = 0;
                let name = wide_string(&record.payload[rel_end..], &mut consumed)?;
                let target = rels
                    .get(&RelId(rid.clone()))
                    .ok_or_else(|| XlsbError::Format(format!("missing workbook rel {rid}")))?;
                let path = join_xl_target(&target.target);
                sheets.push(XlsbSheet { name, path, state });
            }
            0x016a => {
                extern_sheets = parse_extern_sheet(record.payload);
            }
            0x0027 => {
                if let Some(name) = parse_defined_name_name(record.payload) {
                    formula_names.push(name);
                }
                if let Some(raw_name) = parse_defined_name(record.payload, &sheets, &extern_sheets)
                {
                    if raw_name.name == "_xlnm.Print_Area" {
                        raw_print_areas.push(raw_name);
                    } else if raw_name.name == "_xlnm.Print_Titles" {
                        raw_print_titles.push(raw_name);
                    } else if !raw_name.name.starts_with("_xlnm.") {
                        raw_names.push(raw_name);
                    }
                }
            }
            _ => {}
        }
    }
    let named_ranges = resolve_xlsb_named_ranges(raw_names);
    let print_areas = resolve_xlsb_print_areas(&sheets, raw_print_areas);
    let print_titles = resolve_xlsb_print_titles(&sheets, raw_print_titles);
    Ok((
        sheets,
        named_ranges,
        print_areas,
        print_titles,
        extern_sheets,
        formula_names,
        date1904,
    ))
}

fn read_shared_strings(zip: &mut ZipArchive<Cursor<Vec<u8>>>) -> Result<Vec<String>> {
    let Some(data) = read_zip_part_optional(zip, "xl/sharedStrings.bin")? else {
        return Ok(Vec::new());
    };
    let mut strings = Vec::new();
    for record in Records::new(&data) {
        let record = record?;
        if record.typ == 0x0013 && record.payload.len() > 1 {
            let mut consumed = 0;
            strings.push(wide_string(&record.payload[1..], &mut consumed)?);
        }
    }
    Ok(strings)
}

fn parse_extern_sheet(payload: &[u8]) -> Vec<XtiRef> {
    let Some(count) = payload.get(0..4).map(le_u32) else {
        return Vec::new();
    };
    payload
        .get(4..)
        .unwrap_or_default()
        .chunks_exact(12)
        .take(count as usize)
        .map(|xti| XtiRef {
            first_sheet: le_i32(&xti[4..8]),
            last_sheet: le_i32(&xti[8..12]),
        })
        .collect()
}

fn parse_defined_name(
    payload: &[u8],
    sheets: &[XlsbSheet],
    extern_sheets: &[XtiRef],
) -> Option<RawNamedRange> {
    if payload.len() < 13 {
        return None;
    }
    let itab = le_u32(&payload[5..9]);
    let local_id = (itab != u32::MAX).then_some(itab as usize);
    let mut name_len = 0;
    let name = wide_string(&payload[9..], &mut name_len).ok()?;
    let formula_offset = 9 + name_len;
    let refers_to = parse_name_formula(payload.get(formula_offset..)?, sheets, extern_sheets)?;
    if refers_to.is_empty() {
        return None;
    }
    Some(RawNamedRange {
        name,
        local_id,
        refers_to,
    })
}

fn parse_defined_name_name(payload: &[u8]) -> Option<String> {
    if payload.len() < 13 {
        return None;
    }
    let mut name_len = 0;
    wide_string(&payload[9..], &mut name_len).ok()
}

fn parse_name_formula(
    payload: &[u8],
    sheets: &[XlsbSheet],
    extern_sheets: &[XtiRef],
) -> Option<String> {
    let rgce_len = payload.get(0..4).map(le_u32)? as usize;
    let rgce = payload.get(4..4 + rgce_len)?;
    let context = formula::FormulaContext {
        sheets,
        extern_sheets,
        named_ranges: &[],
        formula_names: &[],
    };
    formula::parse_formula_rgce_with_context(rgce, Some(&context))
}

fn read_tables_bin(
    zip: &mut ZipArchive<Cursor<Vec<u8>>>,
    sheet_path: &str,
    sheet_data: &[u8],
    rels: Option<&RelsGraph>,
) -> Result<Vec<Table>> {
    let Some(rels) = rels else {
        return Ok(Vec::new());
    };
    let table_rids = parse_table_part_rids(sheet_data)?;
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
        let Some(table_data) = read_zip_part_optional(zip, &table_path)? else {
            continue;
        };
        if let Some(table) = worksheet_features::parse_table_bin(&table_data) {
            out.push(table);
        }
    }
    Ok(out)
}

fn parse_table_part_rids(data: &[u8]) -> Result<Vec<String>> {
    let mut rids = Vec::new();
    for record in Records::new(data) {
        let record = record?;
        if record.typ != 0x0295 {
            continue;
        }
        let mut offset = 0;
        if let Some(rid) = worksheet_meta::read_wide_string_at(record.payload, &mut offset)
            .filter(|value| !value.is_empty())
        {
            rids.push(rid);
        }
    }
    Ok(rids)
}

fn resolve_xlsb_named_ranges(raw_names: Vec<RawNamedRange>) -> Vec<NamedRange> {
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
            // BIFF12 doesn't carry the ECMA-376 extra attrs; G22 fields
            // default to None/false so the public NamedRange has a
            // consistent shape across xlsx/xlsb readers.
            ..Default::default()
        })
        .collect()
}

fn resolve_xlsb_print_areas(
    sheets: &[XlsbSheet],
    raw_print_areas: Vec<RawNamedRange>,
) -> HashMap<String, String> {
    let mut out = HashMap::new();
    for raw in raw_print_areas {
        let sheet_name = raw
            .local_id
            .and_then(|index| sheets.get(index).map(|sheet| sheet.name.clone()))
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

fn resolve_xlsb_print_titles(
    sheets: &[XlsbSheet],
    raw_print_titles: Vec<RawNamedRange>,
) -> HashMap<String, PrintTitlesInfo> {
    let mut out = HashMap::new();
    for raw in raw_print_titles {
        let sheet_name = raw
            .local_id
            .and_then(|index| sheets.get(index).map(|sheet| sheet.name.clone()))
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

#[cfg(test)]
fn parse_worksheet(
    data: &[u8],
    shared_strings: &[String],
    rels: Option<&RelsGraph>,
    comments: Vec<Comment>,
    formula_context: Option<&formula::FormulaContext<'_>>,
) -> Result<WorksheetData> {
    parse_worksheet_with_tables(
        data,
        shared_strings,
        rels,
        comments,
        Vec::new(),
        formula_context,
    )
}

fn parse_worksheet_with_tables(
    data: &[u8],
    shared_strings: &[String],
    rels: Option<&RelsGraph>,
    comments: Vec<Comment>,
    tables: Vec<Table>,
    formula_context: Option<&formula::FormulaContext<'_>>,
) -> Result<WorksheetData> {
    let mut cells = Vec::new();
    let mut hyperlink_nodes = Vec::new();
    let mut dimension = None;
    let mut row_heights = HashMap::new();
    let mut column_widths = Vec::new();
    let mut merged_ranges = Vec::new();
    let mut hidden_rows = Vec::new();
    let mut hidden_columns = Vec::new();
    let mut row_outline_levels = Vec::new();
    let mut column_outline_levels = Vec::new();
    let mut freeze_panes = None;
    let mut sheet_view = None;
    let mut sheet_protection = None;
    let mut page_margins = None;
    let mut page_setup = None;
    let mut print_options = None;
    let mut header_footer = None;
    let mut sheet_format = None;
    let mut sheet_properties = None;
    let mut auto_filter = None;
    let mut current_filter_column: Option<FilterColumnInfo> = None;
    let mut current_filter_values: Option<(bool, Vec<String>)> = None;
    let mut data_validations = Vec::new();
    let mut pending_data_validation_list = None;
    let mut conditional_formats = Vec::new();
    let mut current_conditional_range: Option<String> = None;
    let mut current_sheet_view: Option<SheetViewInfo> = None;
    let mut shared_formula_masters = Vec::new();
    let mut last_formula_cell_index: Option<usize> = None;
    let mut row = 0u32;
    for record in Records::new(data) {
        let record = record?;
        match record.typ {
            0x0094 => {
                dimension = worksheet_meta::parse_ws_dimension(record.payload);
            }
            0x0000 => {
                if record.payload.len() >= 4 {
                    row = le_u32(&record.payload[0..4]);
                }
                if let Some(info) = worksheet_meta::parse_row_header(record.payload) {
                    if info.hidden {
                        hidden_rows.push(info.row);
                    }
                    if info.outline_level > 0 {
                        row_outline_levels.push((info.row, info.outline_level));
                    }
                    if let Some(height) = info.height {
                        row_heights.insert(info.row, height);
                    }
                }
            }
            0x0001 => {
                if let Some((col, style_id)) = parse_cell_header(record.payload) {
                    cells.push(make_cell(
                        row,
                        col,
                        style_id,
                        CellValue::Empty,
                        CellDataType::Number,
                        None,
                    ));
                }
            }
            0x0002 => {
                if let Some((col, style_id)) = parse_cell_header(record.payload) {
                    let value = parse_rk(record.payload.get(8..12).unwrap_or(&[]));
                    cells.push(make_cell(
                        row,
                        col,
                        style_id,
                        CellValue::Number(value),
                        CellDataType::Number,
                        None,
                    ));
                }
            }
            0x0003 | 0x000b => {
                if let Some((col, style_id)) = parse_cell_header(record.payload) {
                    let err = record
                        .payload
                        .get(8)
                        .copied()
                        .map(error_code)
                        .unwrap_or("#ERROR!");
                    let formula = formula::parse_formula_from_cell_record(
                        record.typ,
                        record.payload,
                        formula_context,
                        row,
                        col,
                    );
                    let shared_anchor =
                        formula::parse_shared_formula_anchor(record.typ, record.payload)
                            .filter(|_| formula.is_none());
                    let mut cell = make_cell(
                        row,
                        col,
                        style_id,
                        CellValue::Error(err.to_string()),
                        CellDataType::Error,
                        formula,
                    );
                    formula::apply_shared_formula_anchor(&mut cell, shared_anchor);
                    last_formula_cell_index = Some(cells.len());
                    cells.push(cell);
                }
            }
            0x0004 | 0x000a => {
                if let Some((col, style_id)) = parse_cell_header(record.payload) {
                    let value = record.payload.get(8).copied().unwrap_or_default() != 0;
                    let formula = formula::parse_formula_from_cell_record(
                        record.typ,
                        record.payload,
                        formula_context,
                        row,
                        col,
                    );
                    let shared_anchor =
                        formula::parse_shared_formula_anchor(record.typ, record.payload)
                            .filter(|_| formula.is_none());
                    let mut cell = make_cell(
                        row,
                        col,
                        style_id,
                        CellValue::Bool(value),
                        CellDataType::Bool,
                        formula,
                    );
                    formula::apply_shared_formula_anchor(&mut cell, shared_anchor);
                    last_formula_cell_index = Some(cells.len());
                    cells.push(cell);
                }
            }
            0x0005 | 0x0009 => {
                if let Some((col, style_id)) = parse_cell_header(record.payload) {
                    let value = record.payload.get(8..16).map(le_f64).unwrap_or_default();
                    let formula = formula::parse_formula_from_cell_record(
                        record.typ,
                        record.payload,
                        formula_context,
                        row,
                        col,
                    );
                    let shared_anchor =
                        formula::parse_shared_formula_anchor(record.typ, record.payload)
                            .filter(|_| formula.is_none());
                    let mut cell = make_cell(
                        row,
                        col,
                        style_id,
                        CellValue::Number(value),
                        CellDataType::Number,
                        formula,
                    );
                    formula::apply_shared_formula_anchor(&mut cell, shared_anchor);
                    last_formula_cell_index = Some(cells.len());
                    cells.push(cell);
                }
            }
            0x0006 | 0x0008 => {
                if let Some((col, style_id)) = parse_cell_header(record.payload) {
                    let mut consumed = 0;
                    let value = wide_string(record.payload.get(8..).unwrap_or(&[]), &mut consumed)?;
                    let formula = formula::parse_formula_from_cell_record(
                        record.typ,
                        record.payload,
                        formula_context,
                        row,
                        col,
                    );
                    let shared_anchor =
                        formula::parse_shared_formula_anchor(record.typ, record.payload)
                            .filter(|_| formula.is_none());
                    let mut cell = make_cell(
                        row,
                        col,
                        style_id,
                        CellValue::String(value),
                        CellDataType::InlineString,
                        formula,
                    );
                    formula::apply_shared_formula_anchor(&mut cell, shared_anchor);
                    last_formula_cell_index = Some(cells.len());
                    cells.push(cell);
                }
            }
            0x0007 => {
                if let Some((col, style_id)) = parse_cell_header(record.payload) {
                    let idx = record.payload.get(8..12).map(le_u32).unwrap_or_default() as usize;
                    let value = shared_strings.get(idx).cloned().unwrap_or_default();
                    cells.push(make_cell(
                        row,
                        col,
                        style_id,
                        CellValue::String(value),
                        CellDataType::SharedString,
                        None,
                    ));
                }
            }
            0x003c => {
                if let Some(info) = worksheet_meta::parse_column_info(record.payload) {
                    if info.width.custom_width {
                        column_widths.push(info.width);
                    }
                    for col in info.width.min..=info.width.max {
                        if info.hidden {
                            hidden_columns.push(col);
                        }
                        if info.outline_level > 0 {
                            column_outline_levels.push((col, info.outline_level));
                        }
                    }
                }
            }
            0x00b0 => {
                if let Some(range) = worksheet_meta::parse_merged_range(record.payload) {
                    merged_ranges.push(range);
                }
            }
            0x01ee => {
                if let Some(node) = worksheet_meta::parse_hyperlink(record.payload) {
                    hyperlink_nodes.push(node);
                }
            }
            0x0089 => {
                if sheet_view.is_none() {
                    current_sheet_view = worksheet_meta::parse_sheet_view(record.payload);
                }
            }
            0x008a => {
                if sheet_view.is_none() {
                    sheet_view = current_sheet_view.take();
                } else {
                    current_sheet_view = None;
                }
            }
            0x0097 => {
                if let Some(pane) = worksheet_meta::parse_pane(record.payload) {
                    if freeze_panes.is_none() {
                        freeze_panes = Some(pane.clone());
                    }
                    if let Some(view) = current_sheet_view.as_mut() {
                        view.pane = Some(pane);
                    }
                }
            }
            0x0098 => {
                if let (Some(view), Some(selection)) = (
                    current_sheet_view.as_mut(),
                    worksheet_meta::parse_selection(record.payload),
                ) {
                    view.selections.push(selection);
                }
            }
            0x0217 => {
                sheet_protection = worksheet_meta::parse_sheet_protection(record.payload);
            }
            0x01dc => {
                page_margins = worksheet_meta::parse_page_margins(record.payload);
            }
            0x01de => {
                page_setup = worksheet_meta::parse_page_setup(record.payload);
            }
            0x01dd => {
                print_options = worksheet_meta::parse_print_options(record.payload);
            }
            0x01df => {
                header_footer = worksheet_meta::parse_header_footer(record.payload);
            }
            0x01e5 => {
                sheet_format = worksheet_meta::parse_sheet_format(record.payload);
            }
            0x0093 => {
                sheet_properties = worksheet_meta::parse_sheet_properties(record.payload);
            }
            0x00a1 => {
                auto_filter = worksheet_features::parse_auto_filter_begin(record.payload);
            }
            0x00a2 => {
                if let (Some(filter), Some(column)) =
                    (auto_filter.as_mut(), current_filter_column.take())
                {
                    filter.filter_columns.push(column);
                }
                current_filter_values = None;
            }
            0x00a3 => {
                current_filter_column = worksheet_features::parse_filter_column(record.payload);
            }
            0x00a4 => {
                if let (Some(filter), Some(column)) =
                    (auto_filter.as_mut(), current_filter_column.take())
                {
                    filter.filter_columns.push(column);
                }
                current_filter_values = None;
            }
            0x00a5 => {
                current_filter_values = Some((
                    record.payload.get(0..4).is_some_and(|v| le_u32(v) != 0),
                    Vec::new(),
                ));
            }
            0x00a6 => {
                if let (Some(column), Some((blank, values))) =
                    (current_filter_column.as_mut(), current_filter_values.take())
                {
                    column.filter = if !values.is_empty() {
                        Some(FilterInfo::String { values })
                    } else {
                        blank.then_some(FilterInfo::Blank)
                    };
                }
            }
            0x00a7 => {
                if let Some((_, values)) = current_filter_values.as_mut() {
                    if let Some(value) = worksheet_features::parse_filter_value(record.payload) {
                        values.push(value);
                    }
                }
            }
            0x02a9 => {
                pending_data_validation_list =
                    worksheet_features::parse_filter_value(record.payload);
            }
            0x0040 => {
                if let Some(validation) = worksheet_features::parse_data_validation(
                    record.payload,
                    pending_data_validation_list.take(),
                ) {
                    data_validations.push(validation);
                }
            }
            0x01cd => {
                current_conditional_range =
                    worksheet_features::parse_conditional_formatting_begin(record.payload);
            }
            0x01cf => {
                if let (Some(range), Some(rule)) = (
                    current_conditional_range.as_deref(),
                    worksheet_features::parse_conditional_format_rule(
                        record.payload,
                        current_conditional_range.as_deref(),
                    ),
                ) {
                    if rule.range == range {
                        conditional_formats.push(rule);
                    }
                }
            }
            0x01ab => {
                if let Some(master) =
                    formula::parse_shared_formula_record(record.payload, formula_context)
                {
                    if let Some(cell) = last_formula_cell_index.and_then(|idx| cells.get_mut(idx)) {
                        cell.formula = Some(master.formula.clone());
                        cell.formula_kind = Some("shared".to_string());
                        cell.formula_shared_index = Some(formula::shared_formula_anchor_key(
                            master.master_row,
                            master.master_col,
                        ));
                    }
                    shared_formula_masters.push(master);
                }
            }
            0x01ce => {
                current_conditional_range = None;
            }
            _ => {
                last_formula_cell_index = None;
            }
        }
    }
    formula::resolve_xlsb_shared_formulas(&mut cells, &shared_formula_masters);
    hidden_rows.sort_unstable();
    hidden_rows.dedup();
    hidden_columns.sort_unstable();
    hidden_columns.dedup();
    row_outline_levels.sort_unstable_by_key(|(row, _)| *row);
    column_outline_levels.sort_unstable_by_key(|(col, _)| *col);
    Ok(worksheet_data(
        dimension.or_else(|| cells_dimension(&cells)),
        row_heights,
        column_widths,
        hidden_rows,
        hidden_columns,
        row_outline_levels,
        column_outline_levels,
        merged_ranges,
        freeze_panes,
        sheet_view,
        worksheet_meta::resolve_hyperlinks(hyperlink_nodes, rels),
        comments,
        sheet_protection,
        page_margins,
        page_setup,
        print_options,
        header_footer,
        sheet_format,
        sheet_properties,
        auto_filter,
        data_validations,
        tables,
        conditional_formats,
        cells,
    ))
}

fn parse_cell_header(payload: &[u8]) -> Option<(u32, Option<u32>)> {
    if payload.len() < 8 {
        return None;
    }
    let col = le_u32(&payload[0..4]);
    let style_id = le_u32(&payload[4..8]);
    Some((col, (style_id != 0).then_some(style_id)))
}

fn make_cell(
    row0: u32,
    col0: u32,
    style_id: Option<u32>,
    value: CellValue,
    data_type: CellDataType,
    formula: Option<String>,
) -> Cell {
    let row = row0 + 1;
    let col = col0 + 1;
    Cell {
        row,
        col,
        coordinate: row_col_to_a1(row, col),
        style_id,
        data_type,
        value,
        formula,
        formula_kind: None,
        formula_shared_index: None,
        array_formula: None,
        rich_text: None,
    }
}

fn worksheet_data(
    dimension: Option<String>,
    row_heights: HashMap<u32, RowHeight>,
    column_widths: Vec<ColumnWidth>,
    hidden_rows: Vec<u32>,
    hidden_columns: Vec<u32>,
    row_outline_levels: Vec<(u32, u8)>,
    column_outline_levels: Vec<(u32, u8)>,
    merged_ranges: Vec<String>,
    freeze_panes: Option<FreezePane>,
    sheet_view: Option<SheetViewInfo>,
    hyperlinks: Vec<Hyperlink>,
    comments: Vec<Comment>,
    sheet_protection: Option<SheetProtection>,
    page_margins: Option<PageMarginsInfo>,
    page_setup: Option<PageSetupInfo>,
    print_options: Option<PrintOptionsInfo>,
    header_footer: Option<HeaderFooterInfo>,
    sheet_format: Option<SheetFormatInfo>,
    sheet_properties: Option<SheetPropertiesInfo>,
    auto_filter: Option<AutoFilterInfo>,
    data_validations: Vec<DataValidation>,
    tables: Vec<Table>,
    conditional_formats: Vec<ConditionalFormatRule>,
    cells: Vec<Cell>,
) -> WorksheetData {
    WorksheetData {
        dimension,
        merged_ranges,
        hyperlinks,
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
        row_breaks: None::<PageBreakListInfo>,
        column_breaks: None::<PageBreakListInfo>,
        sheet_format,
        images: Vec::<ImageInfo>::new(),
        charts: Vec::new(),
        tables,
        conditional_formats,
        hidden_rows,
        hidden_columns,
        row_outline_levels,
        column_outline_levels,
        array_formulas: HashMap::new(),
        cells,
    }
}

fn cells_dimension(cells: &[Cell]) -> Option<String> {
    let mut bounds: Option<(u32, u32, u32, u32)> = None;
    for cell in cells {
        bounds = match bounds {
            Some((min_r, min_c, max_r, max_c)) => Some((
                min_r.min(cell.row),
                min_c.min(cell.col),
                max_r.max(cell.row),
                max_c.max(cell.col),
            )),
            None => Some((cell.row, cell.col, cell.row, cell.col)),
        };
    }
    bounds.map(|(min_r, min_c, max_r, max_c)| {
        format!(
            "{}:{}",
            row_col_to_a1(min_r, min_c),
            row_col_to_a1(max_r, max_c)
        )
    })
}

fn parse_rk(bytes: &[u8]) -> f64 {
    if bytes.len() < 4 {
        return 0.0;
    }
    let mut raw = [0u8; 4];
    raw.copy_from_slice(&bytes[..4]);
    let flags = raw[0] & 0x03;
    raw[0] &= 0xfc;
    let mut value = if flags & 0x02 != 0 {
        i32::from_le_bytes(raw) as f64 / 4.0
    } else {
        let mut eight = [0u8; 8];
        eight[4..8].copy_from_slice(&raw);
        f64::from_le_bytes(eight)
    };
    if flags & 0x01 != 0 {
        value /= 100.0;
    }
    value
}

fn find_trailing_wide_string(payload: &[u8]) -> Option<(usize, String)> {
    for offset in (0..payload.len().saturating_sub(4)).rev() {
        let len = le_u32(&payload[offset..offset + 4]) as usize;
        let end = offset + 4 + len * 2;
        if len > 0 && end == payload.len() {
            return Some((offset, utf16_string(&payload[offset + 4..end])));
        }
    }
    None
}

fn wide_string(bytes: &[u8], consumed: &mut usize) -> Result<String> {
    if bytes.len() < 4 {
        return Err(XlsbError::Format(
            "truncated wide string length".to_string(),
        ));
    }
    let len = le_u32(&bytes[0..4]) as usize;
    let end = 4 + len * 2;
    if bytes.len() < end {
        return Err(XlsbError::Format(
            "truncated wide string payload".to_string(),
        ));
    }
    *consumed = end;
    Ok(utf16_string(&bytes[4..end]))
}

pub(super) fn utf16_string(bytes: &[u8]) -> String {
    let units: Vec<u16> = bytes.chunks_exact(2).map(le_u16).collect();
    String::from_utf16_lossy(&units)
}

fn join_xl_target(target: &str) -> String {
    let target = target.trim_start_matches('/');
    if target.starts_with("xl/") {
        target.to_string()
    } else {
        format!("xl/{target}")
    }
}

pub(super) fn error_code(code: u8) -> &'static str {
    match code {
        0x00 => "#NULL!",
        0x07 => "#DIV/0!",
        0x0f => "#VALUE!",
        0x17 => "#REF!",
        0x1d => "#NAME?",
        0x24 => "#NUM!",
        0x2a => "#N/A",
        _ => "#ERROR!",
    }
}

pub(super) fn le_u16(bytes: &[u8]) -> u16 {
    u16::from_le_bytes([bytes[0], bytes[1]])
}

pub(super) fn le_u32(bytes: &[u8]) -> u32 {
    u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]])
}

pub(super) fn le_i32(bytes: &[u8]) -> i32 {
    i32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]])
}

pub(super) fn le_f64(bytes: &[u8]) -> f64 {
    f64::from_le_bytes([
        bytes[0], bytes[1], bytes[2], bytes[3], bytes[4], bytes[5], bytes[6], bytes[7],
    ])
}

struct Records<'a> {
    bytes: &'a [u8],
    offset: usize,
}

impl<'a> Records<'a> {
    fn new(bytes: &'a [u8]) -> Self {
        Self { bytes, offset: 0 }
    }
}

impl<'a> Iterator for Records<'a> {
    type Item = Result<Record<'a>>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.offset >= self.bytes.len() {
            return None;
        }
        let typ = match read_record_type(self.bytes, &mut self.offset) {
            Ok(value) => value,
            Err(e) => return Some(Err(e)),
        };
        let len = match read_record_len(self.bytes, &mut self.offset) {
            Ok(value) => value,
            Err(e) => return Some(Err(e)),
        };
        let end = self.offset.saturating_add(len);
        if end > self.bytes.len() {
            return Some(Err(XlsbError::Format(
                "truncated record payload".to_string(),
            )));
        }
        let payload = &self.bytes[self.offset..end];
        self.offset = end;
        Some(Ok(Record { typ, payload }))
    }
}

fn read_record_type(bytes: &[u8], offset: &mut usize) -> Result<u16> {
    let first = read_byte(bytes, offset)?;
    let mut typ = (first & 0x7f) as u16;
    if first & 0x80 != 0 {
        let second = read_byte(bytes, offset)?;
        typ += ((second & 0x7f) as u16) << 7;
    }
    Ok(typ)
}

fn read_record_len(bytes: &[u8], offset: &mut usize) -> Result<usize> {
    let mut shift = 0;
    let mut len = 0usize;
    loop {
        let byte = read_byte(bytes, offset)?;
        len += ((byte & 0x7f) as usize) << shift;
        if byte & 0x80 == 0 {
            break;
        }
        shift += 7;
        if shift > 21 {
            return Err(XlsbError::Format(
                "record length varint too long".to_string(),
            ));
        }
    }
    Ok(len)
}

fn read_byte(bytes: &[u8], offset: &mut usize) -> Result<u8> {
    let Some(value) = bytes.get(*offset).copied() else {
        return Err(XlsbError::Format("unexpected end of records".to_string()));
    };
    *offset += 1;
    Ok(value)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{HeaderFooterItemInfo, PaneMode};

    use super::worksheet_meta::{ColumnInfo, RowHeaderInfo};

    fn put_u16(out: &mut Vec<u8>, value: u16) {
        out.extend_from_slice(&value.to_le_bytes());
    }

    fn put_u32(out: &mut Vec<u8>, value: u32) {
        out.extend_from_slice(&value.to_le_bytes());
    }

    fn put_i32(out: &mut Vec<u8>, value: i32) {
        out.extend_from_slice(&value.to_le_bytes());
    }

    fn put_f64(out: &mut Vec<u8>, value: f64) {
        out.extend_from_slice(&value.to_le_bytes());
    }

    fn put_wide_string(out: &mut Vec<u8>, value: &str) {
        let units: Vec<u16> = value.encode_utf16().collect();
        put_u32(out, units.len() as u32);
        for unit in units {
            put_u16(out, unit);
        }
    }

    fn push_record(out: &mut Vec<u8>, typ: u16, payload: &[u8]) {
        if typ < 0x80 {
            out.push(typ as u8);
        } else {
            out.push(((typ & 0x7f) as u8) | 0x80);
            out.push((typ >> 7) as u8);
        }
        let mut len = payload.len();
        loop {
            let mut byte = (len & 0x7f) as u8;
            len >>= 7;
            if len != 0 {
                byte |= 0x80;
            }
            out.push(byte);
            if len == 0 {
                break;
            }
        }
        out.extend_from_slice(payload);
    }

    #[test]
    fn parses_xlsb_row_header_visibility_outline_and_height() {
        let mut payload = Vec::new();
        put_u32(&mut payload, 4); // zero-based row 5
        put_u32(&mut payload, 0);
        put_u16(&mut payload, 420); // 21pt
        put_u16(&mut payload, (3 << 8) | 0x1000 | 0x2000);

        assert_eq!(
            worksheet_meta::parse_row_header(&payload),
            Some(RowHeaderInfo {
                row: 5,
                hidden: true,
                outline_level: 3,
                height: Some(RowHeight {
                    height: 21.0,
                    custom_height: true,
                }),
            })
        );
    }

    #[test]
    fn parses_xlsb_column_info_visibility_outline_and_width() {
        let mut payload = Vec::new();
        put_u32(&mut payload, 1); // zero-based B
        put_u32(&mut payload, 2); // zero-based C
        put_u32(&mut payload, 12 * 256);
        put_u32(&mut payload, 0);
        put_u16(&mut payload, 0x0001 | 0x0002 | (2 << 8));

        assert_eq!(
            worksheet_meta::parse_column_info(&payload),
            Some(ColumnInfo {
                width: ColumnWidth {
                    min: 2,
                    max: 3,
                    width: 12.0,
                    custom_width: true,
                },
                hidden: true,
                outline_level: 2,
            })
        );
    }

    #[test]
    fn parses_xlsb_worksheet_dimension_and_merged_range_as_inclusive_bounds() {
        let mut payload = Vec::new();
        put_u32(&mut payload, 3); // B4
        put_u32(&mut payload, 6); // B7
        put_u32(&mut payload, 1);
        put_u32(&mut payload, 1);
        assert_eq!(
            worksheet_meta::parse_ws_dimension(&payload),
            Some("B4:B7".to_string())
        );

        let mut merge = Vec::new();
        put_u32(&mut merge, 0); // A1
        put_u32(&mut merge, 1); // B2
        put_u32(&mut merge, 0);
        put_u32(&mut merge, 1);
        assert_eq!(
            worksheet_meta::parse_merged_range(&merge),
            Some("A1:B2".to_string())
        );
    }

    #[test]
    fn parse_worksheet_collects_xlsb_sheet_structure_metadata() {
        let mut data = Vec::new();
        let mut row = Vec::new();
        put_u32(&mut row, 4);
        put_u32(&mut row, 0);
        put_u16(&mut row, 420);
        put_u16(&mut row, (1 << 8) | 0x1000 | 0x2000);
        put_u32(&mut row, 0);
        push_record(&mut data, 0x0000, &row);

        let mut cell = Vec::new();
        put_u32(&mut cell, 1);
        put_u32(&mut cell, 0);
        cell.extend_from_slice(&42.0f64.to_le_bytes());
        push_record(&mut data, 0x0005, &cell);

        let mut col = Vec::new();
        put_u32(&mut col, 1);
        put_u32(&mut col, 1);
        put_u32(&mut col, 20 * 256);
        put_u32(&mut col, 0);
        put_u16(&mut col, 0x0001 | 0x0002 | (2 << 8));
        push_record(&mut data, 0x003c, &col);

        let mut merge = Vec::new();
        put_u32(&mut merge, 4);
        put_u32(&mut merge, 4);
        put_u32(&mut merge, 1);
        put_u32(&mut merge, 2);
        push_record(&mut data, 0x00b0, &merge);

        let sheet =
            parse_worksheet(&data, &[], None, Vec::new(), None).expect("parse synthetic worksheet");
        assert_eq!(sheet.merged_ranges, vec!["B5:C5"]);
        assert_eq!(sheet.hidden_rows, vec![5]);
        assert_eq!(sheet.row_outline_levels, vec![(5, 1)]);
        assert_eq!(
            sheet.row_heights.get(&5),
            Some(&RowHeight {
                height: 21.0,
                custom_height: true,
            })
        );
        assert_eq!(sheet.hidden_columns, vec![2]);
        assert_eq!(sheet.column_outline_levels, vec![(2, 2)]);
        assert_eq!(
            sheet.column_widths,
            vec![ColumnWidth {
                min: 2,
                max: 2,
                width: 20.0,
                custom_width: true,
            }]
        );
    }

    #[test]
    fn parses_xlsb_freeze_pane_record() {
        let mut payload = Vec::new();
        payload.extend_from_slice(&1.0f64.to_le_bytes());
        payload.extend_from_slice(&1.0f64.to_le_bytes());
        put_u32(&mut payload, 1);
        put_u32(&mut payload, 1);
        put_u32(&mut payload, 0);
        payload.push(0x02);

        assert_eq!(
            worksheet_meta::parse_pane(&payload),
            Some(FreezePane {
                mode: PaneMode::Freeze,
                top_left_cell: Some("B2".to_string()),
                x_split: Some(1),
                y_split: Some(1),
                active_pane: Some("bottomRight".to_string()),
            })
        );
    }

    #[test]
    fn parse_worksheet_collects_xlsb_freeze_panes() {
        let mut data = Vec::new();
        let mut pane = Vec::new();
        pane.extend_from_slice(&1.0f64.to_le_bytes());
        pane.extend_from_slice(&0.0f64.to_le_bytes());
        put_u32(&mut pane, 1);
        put_u32(&mut pane, 0);
        put_u32(&mut pane, 2);
        pane.push(0x02);
        push_record(&mut data, 0x0097, &pane);

        let sheet =
            parse_worksheet(&data, &[], None, Vec::new(), None).expect("parse freeze worksheet");

        assert_eq!(
            sheet.freeze_panes,
            Some(FreezePane {
                mode: PaneMode::Freeze,
                top_left_cell: Some("A2".to_string()),
                x_split: Some(1),
                y_split: Some(0),
                active_pane: Some("bottomLeft".to_string()),
            })
        );
    }

    #[test]
    fn parses_xlsb_sheet_view_record() {
        let mut payload = Vec::new();
        put_u16(&mut payload, 0x0004 | 0x0008 | 0x0010 | 0x0040 | 0x0100);
        put_u32(&mut payload, 2);
        put_u32(&mut payload, 2);
        put_u32(&mut payload, 3);
        payload.push(0);
        payload.push(0);
        put_u16(&mut payload, 0);
        put_u16(&mut payload, 150);
        put_u16(&mut payload, 120);
        put_u16(&mut payload, 80);
        put_u16(&mut payload, 90);
        put_u32(&mut payload, 1);

        let view = worksheet_meta::parse_sheet_view(&payload).unwrap();

        assert_eq!(view.zoom_scale, 150);
        assert_eq!(view.zoom_scale_normal, 120);
        assert_eq!(view.view, "pageLayout");
        assert_eq!(view.show_grid_lines, Some(true));
        assert_eq!(view.show_row_col_headers, Some(true));
        assert_eq!(view.show_outline_symbols, Some(true));
        assert_eq!(view.show_zeros, Some(true));
        assert_eq!(view.tab_selected, Some(true));
        assert_eq!(view.top_left_cell.as_deref(), Some("D3"));
        assert_eq!(view.workbook_view_id, 1);
    }

    #[test]
    fn parses_xlsb_selection_record() {
        let mut payload = Vec::new();
        put_u32(&mut payload, 0);
        put_u32(&mut payload, 2);
        put_u32(&mut payload, 3);
        put_u32(&mut payload, 0);
        put_u32(&mut payload, 1);
        put_u32(&mut payload, 2);
        put_u32(&mut payload, 4);
        put_u32(&mut payload, 3);
        put_u32(&mut payload, 5);

        let selection = worksheet_meta::parse_selection(&payload).unwrap();

        assert_eq!(selection.pane.as_deref(), Some("bottomRight"));
        assert_eq!(selection.active_cell.as_deref(), Some("D3"));
        assert_eq!(selection.sqref.as_deref(), Some("D3:F5"));
        assert_eq!(selection.active_cell_id, Some(0));
    }

    #[test]
    fn parse_worksheet_collects_xlsb_sheet_view_metadata() {
        let mut data = Vec::new();
        let mut view = Vec::new();
        put_u16(&mut view, 0x0004 | 0x0008 | 0x0010 | 0x0100);
        put_u32(&mut view, 0);
        put_u32(&mut view, 0);
        put_u32(&mut view, 0);
        view.push(0);
        view.push(0);
        put_u16(&mut view, 0);
        put_u16(&mut view, 100);
        put_u16(&mut view, 100);
        put_u16(&mut view, 100);
        put_u16(&mut view, 100);
        put_u32(&mut view, 0);
        push_record(&mut data, 0x0089, &view);

        let mut selection = Vec::new();
        put_u32(&mut selection, 3);
        put_u32(&mut selection, 0);
        put_u32(&mut selection, 0);
        put_u32(&mut selection, 0);
        put_u32(&mut selection, 1);
        put_u32(&mut selection, 0);
        put_u32(&mut selection, 0);
        put_u32(&mut selection, 0);
        put_u32(&mut selection, 0);
        push_record(&mut data, 0x0098, &selection);
        push_record(&mut data, 0x008a, &[]);

        let sheet = parse_worksheet(&data, &[], None, Vec::new(), None)
            .expect("parse sheet view worksheet");
        let view = sheet.sheet_view.unwrap();

        assert_eq!(view.view, "normal");
        assert_eq!(view.selections.len(), 1);
        assert_eq!(view.selections[0].active_cell.as_deref(), Some("A1"));
    }

    #[test]
    fn parses_xlsb_hyperlink_record_with_relationship_target() {
        let mut payload = Vec::new();
        put_u32(&mut payload, 0);
        put_u32(&mut payload, 0);
        put_u32(&mut payload, 0);
        put_u32(&mut payload, 0);
        put_wide_string(&mut payload, "rId5");
        put_wide_string(&mut payload, "");
        put_wide_string(&mut payload, "tip");
        put_wide_string(&mut payload, "Example");
        let rels = RelsGraph::parse(
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
            </Relationships>"#,
        )
        .expect("parse rels");

        let node = worksheet_meta::parse_hyperlink(&payload).unwrap();
        let links = worksheet_meta::resolve_hyperlinks(vec![node], Some(&rels));

        assert_eq!(
            links,
            vec![Hyperlink {
                cell: "A1".to_string(),
                target: "https://example.com".to_string(),
                location: None,
                display: "Example".to_string(),
                tooltip: Some("tip".to_string()),
                internal: false,
            }]
        );
    }

    #[test]
    fn parses_xlsb_internal_hyperlink_record() {
        let mut payload = Vec::new();
        put_u32(&mut payload, 1);
        put_u32(&mut payload, 1);
        put_u32(&mut payload, 1);
        put_u32(&mut payload, 1);
        put_wide_string(&mut payload, "");
        put_wide_string(&mut payload, "Sheet2!A1");
        put_wide_string(&mut payload, "");
        put_wide_string(&mut payload, "");

        let node = worksheet_meta::parse_hyperlink(&payload).unwrap();
        let links = worksheet_meta::resolve_hyperlinks(vec![node], None);

        assert_eq!(
            links,
            vec![Hyperlink {
                cell: "B2".to_string(),
                target: "Sheet2!A1".to_string(),
                location: Some("Sheet2!A1".to_string()),
                display: String::new(),
                tooltip: None,
                internal: true,
            }]
        );
    }

    #[test]
    fn parse_worksheet_collects_xlsb_hyperlinks() {
        let mut data = Vec::new();
        let mut link = Vec::new();
        put_u32(&mut link, 0);
        put_u32(&mut link, 0);
        put_u32(&mut link, 0);
        put_u32(&mut link, 0);
        put_wide_string(&mut link, "rId2");
        put_wide_string(&mut link, "");
        put_wide_string(&mut link, "");
        put_wide_string(&mut link, "Docs");
        push_record(&mut data, 0x01ee, &link);
        let rels = RelsGraph::parse(
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://docs.example" TargetMode="External"/>
            </Relationships>"#,
        )
        .expect("parse rels");

        let sheet =
            parse_worksheet(&data, &[], Some(&rels), Vec::new(), None).expect("parse hyperlinks");

        assert_eq!(sheet.hyperlinks.len(), 1);
        assert_eq!(sheet.hyperlinks[0].cell, "A1");
        assert_eq!(sheet.hyperlinks[0].target, "https://docs.example");
        assert_eq!(sheet.hyperlinks[0].display, "Docs");
    }

    #[test]
    fn parses_xlsb_binary_comments_authors_and_text() {
        let mut data = Vec::new();
        let mut author = Vec::new();
        put_wide_string(&mut author, "Analyst");
        push_record(&mut data, 0x0278, &author);

        let mut comment = Vec::new();
        put_u32(&mut comment, 0);
        put_u32(&mut comment, 2);
        put_u32(&mut comment, 2);
        put_u32(&mut comment, 1);
        put_u32(&mut comment, 1);
        comment.extend_from_slice(&[0; 16]);
        push_record(&mut data, 0x027b, &comment);

        let mut text = vec![0];
        put_wide_string(&mut text, "Review EBITDA addback");
        push_record(&mut data, 0x027d, &text);
        push_record(&mut data, 0x027c, &[]);

        assert_eq!(
            worksheet_meta::parse_comments_bin(&data).expect("parse comments"),
            vec![Comment {
                cell: "B3".to_string(),
                text: "Review EBITDA addback".to_string(),
                author: "Analyst".to_string(),
                threaded: false,
            }]
        );
    }

    #[test]
    fn parse_worksheet_collects_xlsb_comments() {
        let comments = vec![Comment {
            cell: "A1".to_string(),
            text: "Looks right".to_string(),
            author: "Reviewer".to_string(),
            threaded: false,
        }];

        let sheet =
            parse_worksheet(&[], &[], None, comments.clone(), None).expect("parse comments");

        assert_eq!(sheet.comments, comments);
    }

    #[test]
    fn parses_xlsb_sheet_protection_record() {
        let mut payload = Vec::new();
        put_u16(&mut payload, 0xc258);
        for value in [
            true, true, false, false, true, true, true, true, true, true, true, false, false, true,
            true, false,
        ] {
            put_u32(&mut payload, u32::from(value));
        }

        assert_eq!(
            worksheet_meta::parse_sheet_protection(&payload),
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
    }

    #[test]
    fn parse_worksheet_collects_xlsb_sheet_protection() {
        let mut data = Vec::new();
        let mut protection = Vec::new();
        put_u16(&mut protection, 0);
        for value in [
            true, false, false, true, true, true, true, true, true, true, true, false, true, true,
            true, false,
        ] {
            put_u32(&mut protection, u32::from(value));
        }
        push_record(&mut data, 0x0217, &protection);

        let sheet = parse_worksheet(&data, &[], None, Vec::new(), None)
            .expect("parse sheet protection worksheet");
        let protection = sheet.sheet_protection.unwrap();

        assert!(protection.sheet);
        assert_eq!(protection.password_hash, None);
    }

    #[test]
    fn parses_xlsb_page_margins_record() {
        let mut payload = Vec::new();
        for value in [0.7, 0.7, 0.75, 0.75, 0.3, 0.3] {
            put_f64(&mut payload, value);
        }

        assert_eq!(
            worksheet_meta::parse_page_margins(&payload),
            Some(PageMarginsInfo {
                left: 0.7,
                right: 0.7,
                top: 0.75,
                bottom: 0.75,
                header: 0.3,
                footer: 0.3,
            })
        );
    }

    #[test]
    fn parse_worksheet_collects_xlsb_page_margins() {
        let mut data = Vec::new();
        let mut margins = Vec::new();
        for value in [0.5, 0.6, 0.7, 0.8, 0.2, 0.25] {
            put_f64(&mut margins, value);
        }
        push_record(&mut data, 0x01dc, &margins);

        let sheet =
            parse_worksheet(&data, &[], None, Vec::new(), None).expect("parse margins worksheet");

        assert_eq!(
            sheet.page_margins,
            Some(PageMarginsInfo {
                left: 0.5,
                right: 0.6,
                top: 0.7,
                bottom: 0.8,
                header: 0.2,
                footer: 0.25,
            })
        );
    }

    #[test]
    fn parses_xlsb_page_setup_record() {
        let mut payload = Vec::new();
        put_u32(&mut payload, 9);
        put_u32(&mut payload, 150);
        put_u32(&mut payload, 300);
        put_u32(&mut payload, 600);
        put_u32(&mut payload, 2);
        put_i32(&mut payload, 4);
        put_u32(&mut payload, 1);
        put_u32(&mut payload, 2);
        put_u16(
            &mut payload,
            (1 << 1) | (1 << 3) | (1 << 5) | (1 << 7) | (1 << 8) | (3 << 9),
        );
        put_wide_string(&mut payload, "rIdPrinter");

        assert_eq!(
            worksheet_meta::parse_page_setup(&payload),
            Some(PageSetupInfo {
                orientation: Some("landscape".to_string()),
                paper_size: Some(9),
                fit_to_width: Some(1),
                fit_to_height: Some(2),
                scale: Some(150),
                first_page_number: Some(4),
                horizontal_dpi: Some(300),
                vertical_dpi: Some(600),
                cell_comments: Some("atEnd".to_string()),
                errors: Some("NA".to_string()),
                use_first_page_number: Some(true),
                use_printer_defaults: Some(false),
                black_and_white: Some(true),
                draft: Some(false),
                ..PageSetupInfo::default()
            })
        );
    }

    #[test]
    fn parse_worksheet_collects_xlsb_page_setup() {
        let mut data = Vec::new();
        let mut setup = Vec::new();
        put_u32(&mut setup, 1);
        put_u32(&mut setup, 100);
        put_u32(&mut setup, 0);
        put_u32(&mut setup, 0);
        put_u32(&mut setup, 1);
        put_i32(&mut setup, 1);
        put_u32(&mut setup, 0);
        put_u32(&mut setup, 0);
        put_u16(&mut setup, 0);
        push_record(&mut data, 0x01de, &setup);

        let sheet =
            parse_worksheet(&data, &[], None, Vec::new(), None).expect("parse setup worksheet");

        assert_eq!(
            sheet.page_setup.unwrap().orientation.as_deref(),
            Some("portrait")
        );
    }

    #[test]
    fn parses_xlsb_print_options_record() {
        assert_eq!(
            worksheet_meta::parse_print_options(&[0x1f, 0x00]),
            Some(PrintOptionsInfo {
                horizontal_centered: Some(true),
                vertical_centered: Some(true),
                headings: Some(true),
                grid_lines: Some(true),
                grid_lines_set: Some(true),
            })
        );
    }

    #[test]
    fn parse_worksheet_collects_xlsb_print_options() {
        let mut data = Vec::new();
        push_record(&mut data, 0x01dd, &[0x10, 0x00]);

        let sheet = parse_worksheet(&data, &[], None, Vec::new(), None)
            .expect("parse print options worksheet");

        assert_eq!(
            sheet.print_options,
            Some(PrintOptionsInfo {
                horizontal_centered: Some(false),
                vertical_centered: Some(false),
                headings: Some(false),
                grid_lines: Some(false),
                grid_lines_set: Some(true),
            })
        );
    }

    #[test]
    fn parses_xlsb_header_footer_record() {
        let mut payload = Vec::new();
        put_u16(&mut payload, 0x000f);
        put_wide_string(&mut payload, "&LLeft&CHead&RRight");
        put_wide_string(&mut payload, "&CPage &P");
        for _ in 0..4 {
            put_i32(&mut payload, -1);
        }

        assert_eq!(
            worksheet_meta::parse_header_footer(&payload),
            Some(HeaderFooterInfo {
                odd_header: HeaderFooterItemInfo {
                    left: Some("Left".to_string()),
                    center: Some("Head".to_string()),
                    right: Some("Right".to_string()),
                },
                odd_footer: HeaderFooterItemInfo {
                    left: None,
                    center: Some("Page &P".to_string()),
                    right: None,
                },
                even_header: HeaderFooterItemInfo::default(),
                even_footer: HeaderFooterItemInfo::default(),
                first_header: HeaderFooterItemInfo::default(),
                first_footer: HeaderFooterItemInfo::default(),
                different_odd_even: true,
                different_first: true,
                scale_with_doc: true,
                align_with_margins: true,
            })
        );
    }

    #[test]
    fn parse_worksheet_collects_xlsb_header_footer() {
        let mut data = Vec::new();
        let mut header_footer = Vec::new();
        put_u16(&mut header_footer, 0x000c);
        put_wide_string(&mut header_footer, "&C&A");
        put_wide_string(&mut header_footer, "&CPage &P");
        for _ in 0..4 {
            put_i32(&mut header_footer, -1);
        }
        push_record(&mut data, 0x01df, &header_footer);

        let sheet = parse_worksheet(&data, &[], None, Vec::new(), None)
            .expect("parse header/footer worksheet");
        let header_footer = sheet.header_footer.unwrap();

        assert_eq!(header_footer.odd_header.center.as_deref(), Some("&A"));
        assert_eq!(header_footer.odd_footer.center.as_deref(), Some("Page &P"));
        assert!(header_footer.scale_with_doc);
        assert!(header_footer.align_with_margins);
    }

    #[test]
    fn parses_xlsb_sheet_format_record() {
        let mut payload = Vec::new();
        put_u32(&mut payload, 11 * 256);
        put_u16(&mut payload, 8);
        put_u16(&mut payload, 300);
        put_u32(&mut payload, 0x0201000d);

        assert_eq!(
            worksheet_meta::parse_sheet_format(&payload),
            Some(SheetFormatInfo {
                base_col_width: 8,
                default_col_width: Some(11.0),
                default_row_height: 15.0,
                custom_height: Some(true),
                zero_height: Some(false),
                thick_top: Some(true),
                thick_bottom: Some(true),
                outline_level_row: 1,
                outline_level_col: 2,
            })
        );
    }

    #[test]
    fn parse_worksheet_collects_xlsb_sheet_format() {
        let mut data = Vec::new();
        let mut format = Vec::new();
        put_u32(&mut format, u32::MAX);
        put_u16(&mut format, 8);
        put_u16(&mut format, 300);
        put_u32(&mut format, 0);
        push_record(&mut data, 0x01e5, &format);

        let sheet = parse_worksheet(&data, &[], None, Vec::new(), None)
            .expect("parse sheet format worksheet");

        assert_eq!(
            sheet.sheet_format,
            Some(SheetFormatInfo {
                base_col_width: 8,
                default_col_width: None,
                default_row_height: 15.0,
                custom_height: Some(false),
                zero_height: Some(false),
                thick_top: Some(false),
                thick_bottom: Some(false),
                outline_level_row: 0,
                outline_level_col: 0,
            })
        );
    }

    #[test]
    fn parses_xlsb_sheet_properties_code_name() {
        let mut payload = vec![
            0xc9, 0x04, 0x02, 0x00, 0x40, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xff, 0xff, 0xff,
            0xff, 0xff, 0xff, 0xff, 0xff, 0x00, 0x00, 0x00, 0x00,
        ];
        put_wide_string(&mut payload, "NativeSheet");

        let properties =
            worksheet_meta::parse_sheet_properties(&payload).expect("sheet properties");
        let defaults = SheetPropertiesInfo::default();

        assert_eq!(properties.code_name.as_deref(), Some("NativeSheet"));
        assert_eq!(properties.outline, defaults.outline);
        assert_eq!(properties.page_setup, defaults.page_setup);
    }

    #[test]
    fn ignores_xlsb_sheet_properties_without_code_name() {
        let payload = [
            0xc9, 0x04, 0x02, 0x00, 0x40, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xff, 0xff, 0xff,
            0xff, 0xff, 0xff, 0xff, 0xff, 0x00, 0x00, 0x00, 0x00,
        ];

        assert_eq!(worksheet_meta::parse_sheet_properties(&payload), None);
    }

    #[test]
    fn parse_worksheet_collects_xlsb_sheet_properties() {
        let mut data = Vec::new();
        let mut properties = vec![
            0xc9, 0x04, 0x02, 0x00, 0x40, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xff, 0xff, 0xff,
            0xff, 0xff, 0xff, 0xff, 0xff, 0x00, 0x00, 0x00, 0x00,
        ];
        put_wide_string(&mut properties, "NativeSheet");
        push_record(&mut data, 0x0093, &properties);

        let sheet = parse_worksheet(&data, &[], None, Vec::new(), None)
            .expect("parse sheet properties worksheet");

        assert_eq!(
            sheet
                .sheet_properties
                .and_then(|properties| properties.code_name),
            Some("NativeSheet".to_string())
        );
    }

    #[test]
    fn parses_xlsb_auto_filter_range() {
        let mut payload = Vec::new();
        put_u32(&mut payload, 3);
        put_u32(&mut payload, 7);
        put_u32(&mut payload, 2);
        put_u32(&mut payload, 2);

        assert_eq!(
            worksheet_features::parse_auto_filter_begin(&payload),
            Some(AutoFilterInfo {
                ref_range: "C4:C8".to_string(),
                filter_columns: Vec::new(),
                sort_state: None,
            })
        );
    }

    #[test]
    fn parses_xlsb_filter_column_flags() {
        let mut payload = Vec::new();
        put_u32(&mut payload, 1);
        put_u16(&mut payload, 0x0003);

        assert_eq!(
            worksheet_features::parse_filter_column(&payload),
            Some(FilterColumnInfo {
                col_id: 1,
                hidden_button: true,
                show_button: false,
                filter: None,
                date_group_items: Vec::new(),
            })
        );
    }

    #[test]
    fn parse_worksheet_collects_xlsb_auto_filter_with_string_values() {
        let mut data = Vec::new();
        let mut range = Vec::new();
        put_u32(&mut range, 0);
        put_u32(&mut range, 3);
        put_u32(&mut range, 0);
        put_u32(&mut range, 2);
        push_record(&mut data, 0x00a1, &range);

        let mut column = Vec::new();
        put_u32(&mut column, 2);
        put_u16(&mut column, 0);
        push_record(&mut data, 0x00a3, &column);

        let mut filters = Vec::new();
        put_u32(&mut filters, 0);
        push_record(&mut data, 0x00a5, &filters);

        let mut value = Vec::new();
        put_wide_string(&mut value, "Open");
        push_record(&mut data, 0x00a7, &value);
        push_record(&mut data, 0x00a6, &[]);
        push_record(&mut data, 0x00a4, &[]);
        push_record(&mut data, 0x00a2, &[]);

        let sheet = parse_worksheet(&data, &[], None, Vec::new(), None)
            .expect("parse auto filter worksheet");
        let auto_filter = sheet.auto_filter.expect("auto filter");

        assert_eq!(auto_filter.ref_range, "A1:C4");
        assert_eq!(auto_filter.filter_columns.len(), 1);
        assert_eq!(auto_filter.filter_columns[0].col_id, 2);
        assert_eq!(
            auto_filter.filter_columns[0].filter,
            Some(FilterInfo::String {
                values: vec!["Open".to_string()],
            })
        );
    }

    #[test]
    fn parses_xlsb_list_data_validation() {
        let mut payload = Vec::new();
        put_u32(&mut payload, 0x0000_0103);
        put_u32(&mut payload, 1);
        put_u32(&mut payload, 1);
        put_u32(&mut payload, 3);
        put_u32(&mut payload, 2);
        put_u32(&mut payload, 2);
        put_wide_string(&mut payload, "Status error");
        put_wide_string(&mut payload, "Choose a valid status");
        put_wide_string(&mut payload, "");
        put_wide_string(&mut payload, "");

        assert_eq!(
            worksheet_features::parse_data_validation(
                &payload,
                Some("\"Open,Closed\"".to_string())
            ),
            Some(DataValidation {
                range: "C2:C4".to_string(),
                validation_type: "list".to_string(),
                operator: None,
                formula1: Some("\"Open,Closed\"".to_string()),
                formula2: None,
                allow_blank: true,
                show_input_message: false,
                show_error_message: false,
                error_title: Some("Status error".to_string()),
                error: Some("Choose a valid status".to_string()),
                prompt_title: None,
                prompt: None,
            })
        );
    }

    #[test]
    fn parse_worksheet_collects_xlsb_data_validations() {
        let mut data = Vec::new();
        let mut list = Vec::new();
        put_wide_string(&mut list, "\"Open,Closed\"");
        push_record(&mut data, 0x02a9, &list);

        let mut payload = Vec::new();
        put_u32(&mut payload, 0x0000_0103);
        put_u32(&mut payload, 1);
        put_u32(&mut payload, 1);
        put_u32(&mut payload, 3);
        put_u32(&mut payload, 2);
        put_u32(&mut payload, 2);
        for _ in 0..4 {
            put_wide_string(&mut payload, "");
        }
        push_record(&mut data, 0x0040, &payload);

        let sheet = parse_worksheet(&data, &[], None, Vec::new(), None)
            .expect("parse data validations worksheet");

        assert_eq!(
            sheet.data_validations,
            vec![DataValidation {
                range: "C2:C4".to_string(),
                validation_type: "list".to_string(),
                operator: None,
                formula1: Some("\"Open,Closed\"".to_string()),
                formula2: None,
                allow_blank: true,
                show_input_message: false,
                show_error_message: false,
                error_title: None,
                error: None,
                prompt_title: None,
                prompt: None,
            }]
        );
    }

    #[test]
    fn parses_xlsb_conditional_format_rule() {
        let mut rgce = Vec::new();
        rgce.push(0x1e);
        put_u16(&mut rgce, 50);

        let mut payload = Vec::new();
        put_u32(&mut payload, 1);
        put_u32(&mut payload, 0);
        put_u32(&mut payload, 0);
        put_i32(&mut payload, 1);
        put_i32(&mut payload, 5);
        put_i32(&mut payload, 0);
        put_i32(&mut payload, 0);
        put_u16(&mut payload, 0x0002);
        put_u32(&mut payload, rgce.len() as u32);
        put_u32(&mut payload, 0);
        put_u32(&mut payload, 0);
        put_i32(&mut payload, -1);
        payload.extend_from_slice(&rgce);

        assert_eq!(
            worksheet_features::parse_conditional_format_rule(&payload, Some("C2:C5")),
            Some(ConditionalFormatRule {
                range: "C2:C5".to_string(),
                rule_type: "cellIs".to_string(),
                operator: Some("greaterThan".to_string()),
                formula: Some("=50".to_string()),
                dxf_id: None,
                priority: Some(1),
                stop_if_true: Some(true),
                color_scale: None,
            })
        );
    }

    #[test]
    fn parse_worksheet_collects_xlsb_conditional_formats() {
        let mut data = Vec::new();
        let mut begin = Vec::new();
        put_u32(&mut begin, 1);
        put_u32(&mut begin, 0);
        put_u32(&mut begin, 1);
        put_u32(&mut begin, 1);
        put_u32(&mut begin, 4);
        put_u32(&mut begin, 2);
        put_u32(&mut begin, 2);
        push_record(&mut data, 0x01cd, &begin);

        let mut rgce = Vec::new();
        rgce.push(0x44);
        put_u32(&mut rgce, 1);
        put_u16(&mut rgce, 0xc001);
        rgce.push(0x1e);
        put_u16(&mut rgce, 0);
        rgce.push(0x0c);

        let mut rule = Vec::new();
        put_u32(&mut rule, 2);
        put_u32(&mut rule, 1);
        put_u32(&mut rule, 0);
        put_i32(&mut rule, 3);
        put_i32(&mut rule, 0);
        put_i32(&mut rule, 0);
        put_i32(&mut rule, 0);
        put_u16(&mut rule, 0);
        put_u32(&mut rule, rgce.len() as u32);
        put_u32(&mut rule, 0);
        put_u32(&mut rule, 0);
        put_i32(&mut rule, -1);
        rule.extend_from_slice(&rgce);
        push_record(&mut data, 0x01cf, &rule);
        push_record(&mut data, 0x01d0, &[]);
        push_record(&mut data, 0x01ce, &[]);

        let sheet = parse_worksheet(&data, &[], None, Vec::new(), None)
            .expect("parse conditional formatting worksheet");

        assert_eq!(
            sheet.conditional_formats,
            vec![ConditionalFormatRule {
                range: "C2:C5".to_string(),
                rule_type: "expression".to_string(),
                operator: None,
                formula: Some("=B2>0".to_string()),
                dxf_id: None,
                priority: Some(3),
                stop_if_true: Some(false),
                color_scale: None,
            }]
        );
    }

    #[test]
    fn parses_xlsb_table_part_relationship_ids() {
        let mut data = Vec::new();
        let mut payload = Vec::new();
        put_wide_string(&mut payload, "rId3");
        push_record(&mut data, 0x0295, &payload);

        assert_eq!(parse_table_part_rids(&data).unwrap(), vec!["rId3"]);
    }

    #[test]
    fn parses_xlsb_table_part_metadata() {
        let mut data = Vec::new();
        let mut table = Vec::new();
        put_u32(&mut table, 0);
        put_u32(&mut table, 3);
        put_u32(&mut table, 0);
        put_u32(&mut table, 2);
        put_u32(&mut table, 1);
        put_u32(&mut table, 7);
        put_u32(&mut table, 1);
        put_u32(&mut table, 0);
        put_u32(&mut table, 1);
        for _ in 0..6 {
            put_u32(&mut table, u32::MAX);
        }
        put_u32(&mut table, 0);
        put_wide_string(&mut table, "SalesTable");
        put_wide_string(&mut table, "SalesTable");
        put_wide_string(&mut table, "table comment");
        put_wide_string(&mut table, "");
        put_wide_string(&mut table, "");
        put_wide_string(&mut table, "");
        push_record(&mut data, 0x0157, &table);

        let mut style = Vec::new();
        put_u16(&mut style, 0x0005);
        put_wide_string(&mut style, "TableStyleMedium2");
        push_record(&mut data, 0x0201, &style);

        for (id, name) in [(1, "Region"), (2, "Amount"), (3, "Status")] {
            let mut column = Vec::new();
            put_u32(&mut column, id);
            put_u32(&mut column, 0);
            put_u32(&mut column, u32::MAX);
            put_u32(&mut column, u32::MAX);
            put_u32(&mut column, u32::MAX);
            put_u32(&mut column, 0);
            put_wide_string(&mut column, name);
            put_wide_string(&mut column, name);
            put_wide_string(&mut column, "");
            put_wide_string(&mut column, "");
            put_wide_string(&mut column, "");
            put_wide_string(&mut column, "");
            push_record(&mut data, 0x015b, &column);
        }
        let mut auto_filter = Vec::new();
        for value in [0, 3, 0, 2] {
            put_u32(&mut auto_filter, value);
        }
        push_record(&mut data, 0x00a1, &auto_filter);

        assert_eq!(
            worksheet_features::parse_table_bin(&data),
            Some(Table {
                name: "SalesTable".to_string(),
                ref_range: "A1:C4".to_string(),
                header_row: true,
                totals_row: false,
                comment: Some("table comment".to_string()),
                table_type: Some("worksheet".to_string()),
                totals_row_shown: Some(true),
                style: Some("TableStyleMedium2".to_string()),
                show_first_column: true,
                show_last_column: false,
                show_row_stripes: true,
                show_column_stripes: false,
                columns: vec![
                    "Region".to_string(),
                    "Amount".to_string(),
                    "Status".to_string()
                ],
                auto_filter: Some(AutoFilterInfo {
                    ref_range: "A1:C4".to_string(),
                    filter_columns: Vec::new(),
                    sort_state: None,
                }),
            })
        );
    }

    #[test]
    fn parses_basic_formula_rgce_tokens() {
        let mut rgce = Vec::new();
        rgce.push(0x44); // PtgRef relative B5
        put_u32(&mut rgce, 4);
        put_u16(&mut rgce, 0xc001);
        rgce.push(0x1e); // PtgInt 2
        put_u16(&mut rgce, 2);
        rgce.push(0x03); // PtgAdd

        assert_eq!(formula::parse_formula_rgce(&rgce), Some("B5+2".to_string()));
    }

    #[test]
    fn parses_variable_function_formula_rgce_tokens() {
        let mut rgce = Vec::new();
        rgce.push(0x45); // PtgArea relative A1:B2
        put_u32(&mut rgce, 0);
        put_u32(&mut rgce, 1);
        put_u16(&mut rgce, 0xc000);
        put_u16(&mut rgce, 0xc001);
        rgce.push(0x22); // PtgFuncVar SUM(argc=1)
        rgce.push(1);
        put_u16(&mut rgce, 4);

        assert_eq!(
            formula::parse_formula_rgce(&rgce),
            Some("SUM(A1:B2)".to_string())
        );
    }

    #[test]
    fn parses_formula_area_absolute_flags() {
        let mut rgce = Vec::new();
        rgce.push(0x25); // PtgArea C$2:C6
        put_u32(&mut rgce, 1);
        put_u32(&mut rgce, 5);
        put_u16(&mut rgce, 0x4002);
        put_u16(&mut rgce, 0xc002);

        assert_eq!(
            formula::parse_formula_rgce(&rgce),
            Some("C$2:C6".to_string())
        );
    }

    #[test]
    fn parses_defined_name_formula_rgce_tokens() {
        let named_ranges = vec![NamedRange {
            name: "Revenue".to_string(),
            scope: "workbook".to_string(),
            refers_to: "Data!$A$1".to_string(),
            ..Default::default()
        }];
        let context = formula::FormulaContext {
            sheets: &[],
            extern_sheets: &[],
            named_ranges: &named_ranges,
            formula_names: &[],
        };
        let mut rgce = Vec::new();
        rgce.push(0x43);
        put_u32(&mut rgce, 1);

        assert_eq!(
            formula::parse_formula_rgce_with_context(&rgce, Some(&context)),
            Some("Revenue".to_string())
        );
    }

    #[test]
    fn parse_worksheet_attaches_formula_text_to_formula_cells() {
        let mut rgce = Vec::new();
        rgce.push(0x1e);
        put_u16(&mut rgce, 40);
        rgce.push(0x1e);
        put_u16(&mut rgce, 2);
        rgce.push(0x03);

        let mut payload = Vec::new();
        put_u32(&mut payload, 0); // A
        put_u32(&mut payload, 0);
        payload.extend_from_slice(&42.0f64.to_le_bytes());
        put_u16(&mut payload, 0);
        put_u32(&mut payload, rgce.len() as u32);
        payload.extend_from_slice(&rgce);

        let mut data = Vec::new();
        let mut row = Vec::new();
        put_u32(&mut row, 0);
        put_u32(&mut row, 0);
        put_u16(&mut row, 0);
        put_u16(&mut row, 0);
        push_record(&mut data, 0x0000, &row);
        push_record(&mut data, 0x0009, &payload);

        let sheet =
            parse_worksheet(&data, &[], None, Vec::new(), None).expect("parse formula worksheet");
        assert_eq!(sheet.cells.len(), 1);
        assert_eq!(sheet.cells[0].value, CellValue::Number(42.0));
        assert_eq!(sheet.cells[0].formula.as_deref(), Some("40+2"));
    }

    #[test]
    fn parse_worksheet_uses_extern_sheets_for_formula_text() {
        let sheets = vec![
            XlsbSheet {
                name: "Inputs".to_string(),
                path: "xl/worksheets/sheet1.bin".to_string(),
                state: SheetState::Visible,
            },
            XlsbSheet {
                name: "Calc".to_string(),
                path: "xl/worksheets/sheet2.bin".to_string(),
                state: SheetState::Visible,
            },
        ];
        let extern_sheets = vec![XtiRef {
            first_sheet: 0,
            last_sheet: 0,
        }];
        let context = formula::FormulaContext {
            sheets: &sheets,
            extern_sheets: &extern_sheets,
            named_ranges: &[],
            formula_names: &[],
        };

        let mut rgce = Vec::new();
        rgce.push(0x5a);
        put_u16(&mut rgce, 0);
        put_u32(&mut rgce, 0);
        put_u16(&mut rgce, 0);

        let mut payload = Vec::new();
        put_u32(&mut payload, 0);
        put_u32(&mut payload, 0);
        payload.extend_from_slice(&42.0f64.to_le_bytes());
        put_u16(&mut payload, 0);
        put_u32(&mut payload, rgce.len() as u32);
        payload.extend_from_slice(&rgce);

        let mut data = Vec::new();
        let mut row = Vec::new();
        put_u32(&mut row, 0);
        put_u32(&mut row, 0);
        put_u16(&mut row, 0);
        put_u16(&mut row, 0);
        push_record(&mut data, 0x0000, &row);
        push_record(&mut data, 0x0009, &payload);

        let sheet = parse_worksheet(&data, &[], None, Vec::new(), Some(&context))
            .expect("parse formula worksheet");

        assert_eq!(sheet.cells[0].formula.as_deref(), Some("Inputs!$A$1"));
    }

    #[test]
    fn parses_xlsb_extern_sheet_records() {
        let mut payload = Vec::new();
        put_u32(&mut payload, 2);
        put_u32(&mut payload, 0);
        put_i32(&mut payload, 0);
        put_i32(&mut payload, 0);
        put_u32(&mut payload, 0);
        put_i32(&mut payload, -2);
        put_i32(&mut payload, -2);

        let refs = parse_extern_sheet(&payload);

        assert_eq!(
            refs,
            vec![
                XtiRef {
                    first_sheet: 0,
                    last_sheet: 0
                },
                XtiRef {
                    first_sheet: -2,
                    last_sheet: -2
                }
            ]
        );
    }

    #[test]
    fn parses_xlsb_defined_name_area3d_formula() {
        let sheets = vec![XlsbSheet {
            name: "Data".to_string(),
            path: "xl/worksheets/sheet1.bin".to_string(),
            state: SheetState::Visible,
        }];
        let extern_sheets = vec![XtiRef {
            first_sheet: 0,
            last_sheet: 0,
        }];
        let mut payload = Vec::new();
        put_u32(&mut payload, 0);
        payload.push(0);
        put_u32(&mut payload, u32::MAX);
        put_wide_string(&mut payload, "GlobalRange");

        let mut rgce = Vec::new();
        rgce.push(0x3b);
        put_u16(&mut rgce, 0);
        put_u32(&mut rgce, 0);
        put_u32(&mut rgce, 1);
        put_u16(&mut rgce, 0);
        put_u16(&mut rgce, 0);
        put_u32(&mut payload, rgce.len() as u32);
        payload.extend_from_slice(&rgce);
        put_u32(&mut payload, 0);

        let raw = parse_defined_name(&payload, &sheets, &extern_sheets).unwrap();
        let ranges = resolve_xlsb_named_ranges(vec![raw]);

        assert_eq!(
            ranges,
            vec![NamedRange {
                name: "GlobalRange".to_string(),
                scope: "workbook".to_string(),
                refers_to: "Data!$A$1:$A$2".to_string(),
                ..Default::default()
            }]
        );
    }

    #[test]
    fn resolves_xlsb_sheet_scoped_defined_name_with_owning_sheet_index() {
        let sheets = vec![XlsbSheet {
            name: "Other".to_string(),
            path: "xl/worksheets/sheet1.bin".to_string(),
            state: SheetState::Visible,
        }];
        let mut payload = Vec::new();
        put_u32(&mut payload, 0);
        payload.push(0);
        put_u32(&mut payload, 0);
        put_wide_string(&mut payload, "LocalConstant");

        let mut rgce = Vec::new();
        rgce.push(0x1e);
        put_u16(&mut rgce, 7);
        put_u32(&mut payload, rgce.len() as u32);
        payload.extend_from_slice(&rgce);
        put_u32(&mut payload, 0);

        let raw = parse_defined_name(&payload, &sheets, &[]).unwrap();
        let ranges = resolve_xlsb_named_ranges(vec![raw]);

        assert_eq!(
            ranges[0],
            NamedRange {
                name: "LocalConstant".to_string(),
                scope: "sheet".to_string(),
                local_sheet_id: Some(0),
                refers_to: "7".to_string(),
                ..Default::default()
            }
        );
    }

    #[test]
    fn resolves_xlsb_print_area_and_titles_from_builtin_names() {
        let sheets = vec![XlsbSheet {
            name: "Data".to_string(),
            path: "xl/worksheets/sheet1.bin".to_string(),
            state: SheetState::Visible,
        }];
        let print_area = RawNamedRange {
            name: "_xlnm.Print_Area".to_string(),
            local_id: Some(0),
            refers_to: "Data!$A$1:$C$10".to_string(),
        };
        let print_titles = RawNamedRange {
            name: "_xlnm.Print_Titles".to_string(),
            local_id: Some(0),
            refers_to: "Data!$1:$2,Data!$A:$B".to_string(),
        };

        let areas = resolve_xlsb_print_areas(&sheets, vec![print_area]);
        let titles = resolve_xlsb_print_titles(&sheets, vec![print_titles]);

        assert_eq!(
            areas.get("Data").map(String::as_str),
            Some("Data!$A$1:$C$10")
        );
        assert_eq!(
            titles.get("Data"),
            Some(&PrintTitlesInfo {
                rows: Some("1:2".to_string()),
                cols: Some("A:B".to_string()),
            })
        );
    }
}
