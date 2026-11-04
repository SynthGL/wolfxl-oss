//! Number-format edits over a worksheet range.
//!
//! A cell's `s` attribute indexes `<cellXfs>` in `xl/styles.xml`, and each
//! `<xf>` there names a `numFmtId` that is either built in or defined by a
//! `<numFmt>` in `<numFmts>` (ECMA-376 Part 1, 18.8). Applying a format code to
//! a range therefore means: find or define the code, then point every affected
//! cell at a copy of its current cell format that differs only in its number
//! format. Identical copies are shared, and an existing identical `<xf>` is
//! reused instead of appending a duplicate.
//!
//! The styles part is edited by splicing: every byte outside the changed
//! collection start tags and the appended entries is copied through. Cells are
//! restyled with style-only [`CellPatch`]es, so values, formulas, and other
//! cell attributes are kept.

use std::borrow::Cow;
use std::collections::{BTreeMap, HashMap};
use std::fmt::Display;
use std::ops::Range;

use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader as XmlReader;

use crate::ooxml_util::attr_value;
use crate::sheet_patcher::{patch_worksheet, CellPatch};
use crate::EditError;

const MAX_ROW: u32 = 1_048_576;
const MAX_COL: u32 = 16_384;

/// Ids 0-163 are reserved for built-in number formats.
const FIRST_CUSTOM_NUM_FMT_ID: u32 = 164;

/// Styles part written when a package has none: the default font, the two
/// fills Excel requires, an empty border, and the Normal cell style.
const MINIMAL_STYLES: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
    "\n",
    r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
    r#"<fonts count="1"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts>"#,
    r#"<fills count="2"><fill><patternFill patternType="none"/></fill>"#,
    r#"<fill><patternFill patternType="gray125"/></fill></fills>"#,
    r#"<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>"#,
    r#"<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>"#,
    r#"<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>"#,
    r#"<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>"#,
    "</styleSheet>",
);

/// Apply `code` to every cell in `bounds` (`min_row, min_col, max_row, max_col`,
/// 1-based and inclusive).
///
/// Returns the updated styles part and worksheet XML. With no `styles` part a
/// minimal valid one is generated; the caller registers it in the package.
/// Existing cells keep their font, fill, border, alignment, and protection;
/// absent cells start from the row or column style they would inherit.
pub(crate) fn set_number_format(
    styles: Option<&str>,
    sheet: &str,
    bounds: (u32, u32, u32, u32),
    code: &str,
) -> Result<(String, String), EditError> {
    validate_bounds(bounds)?;
    let (min_row, min_col, max_row, max_col) = bounds;
    let grid = scan_worksheet(sheet, bounds)?;

    let mut source = styles.unwrap_or(MINIMAL_STYLES);
    let mut scan = scan_styles(source)?;
    if scan.root_self_closing {
        // An empty `<styleSheet/>` defines nothing a cell could keep.
        source = MINIMAL_STYLES;
        scan = scan_styles(source)?;
    }
    let mut plan = FormatPlan::new(source, &scan, code)?;

    let column_bases = grid.column_styles(min_col, max_col);
    let mut patches = Vec::new();
    for row in min_row..=max_row {
        let row_info = grid.rows.get(&row);
        let row_style = row_info.and_then(|info| info.style);
        let mut present = row_info
            .map_or(&[][..], |info| info.cells.as_slice())
            .iter()
            .peekable();
        for (col, &column_base) in (min_col..=max_col).zip(&column_bases) {
            let existing = match present.peek() {
                Some(&&(cell_col, style)) if cell_col == col => {
                    present.next();
                    Some(style)
                }
                _ => None,
            };
            let base = existing.or(row_style).unwrap_or(column_base);
            let target = plan.target_for(base, row, col)?;
            if target != base {
                patches.push(CellPatch {
                    row,
                    col,
                    value: None,
                    style_index: Some(target),
                });
            }
        }
    }

    let sheet_xml = patch_worksheet(sheet, &patches).map_err(worksheet_error)?;
    Ok((plan.finish()?, sheet_xml))
}

// ---------------------------------------------------------------------------
// Format plan
// ---------------------------------------------------------------------------

/// Maps each base cell format to its formatted copy and collects new entries.
struct FormatPlan<'a> {
    xml: &'a str,
    scan: &'a StylesScan,
    cell_xfs: &'a Section,
    code: &'a str,
    num_fmt_id: u32,
    defines_num_fmt: bool,
    /// Existing cell formats by index. An empty `<cellXfs>` gets a default
    /// format at index 0 so unstyled cells keep their meaning.
    xfs: Vec<Cow<'a, str>>,
    synthesized_default: bool,
    /// Canonical form of every reusable format, mapped to its index.
    known: HashMap<String, u32>,
    /// Base format index to formatted index.
    targets: HashMap<u32, u32>,
    appended: Vec<String>,
}

impl<'a> FormatPlan<'a> {
    fn new(xml: &'a str, scan: &'a StylesScan, code: &'a str) -> Result<Self, EditError> {
        let cell_xfs = scan.cell_xfs.as_ref().ok_or_else(|| {
            EditError::Invalid("styles part has no cellXfs collection".to_string())
        })?;
        let reused = scan
            .custom_formats
            .iter()
            .find(|(_, existing)| existing.as_str() == code)
            .map(|&(id, _)| id);
        let (num_fmt_id, defines_num_fmt) = match reused {
            Some(id) => (id, false),
            None => {
                let next = match scan.max_num_fmt_id {
                    Some(max) => max.checked_add(1).ok_or_else(|| {
                        EditError::Invalid("styles part has no free number format id".to_string())
                    })?,
                    None => FIRST_CUSTOM_NUM_FMT_ID,
                };
                (next.max(FIRST_CUSTOM_NUM_FMT_ID), true)
            }
        };

        let mut xfs: Vec<Cow<'a, str>> = scan
            .xfs
            .iter()
            .map(|xf| Cow::Borrowed(&xml[xf.outer.clone()]))
            .collect();
        let mut num_fmt_ids: Vec<u32> = scan.xfs.iter().map(|xf| xf.num_fmt_id).collect();
        let synthesized_default = xfs.is_empty();
        if synthesized_default {
            xfs.push(Cow::Owned(format!(
                r#"<{} numFmtId="0" fontId="0" fillId="0" borderId="0"/>"#,
                qualify(cell_xfs.prefix.as_deref(), "xf")
            )));
            num_fmt_ids.push(0);
        }

        // Only formats already using the target number format can match a copy.
        let mut known = HashMap::new();
        for (index, (text, &id)) in xfs.iter().zip(&num_fmt_ids).enumerate() {
            if id == num_fmt_id {
                known
                    .entry(canonical_element(text)?)
                    .or_insert(index as u32);
            }
        }

        Ok(Self {
            xml,
            scan,
            cell_xfs,
            code,
            num_fmt_id,
            defines_num_fmt,
            xfs,
            synthesized_default,
            known,
            targets: HashMap::new(),
            appended: Vec::new(),
        })
    }

    /// Index of the formatted copy of cell format `base`, used by cell `row`/`col`.
    fn target_for(&mut self, base: u32, row: u32, col: u32) -> Result<u32, EditError> {
        if let Some(&target) = self.targets.get(&base) {
            return Ok(target);
        }
        let Some(source) = self.xfs.get(base as usize) else {
            return Err(EditError::Invalid(format!(
                "cell {} uses cell format {base}, but the styles part defines {} cell format(s)",
                cell_label(row, col),
                self.scan.xfs.len()
            )));
        };
        let copy = clone_xf(source, self.num_fmt_id)?;
        let key = canonical_element(&copy)?;
        let target = match self.known.get(&key) {
            Some(&index) => index,
            None => {
                let index = (self.xfs.len() + self.appended.len()) as u32;
                self.known.insert(key, index);
                self.appended.push(copy);
                index
            }
        };
        self.targets.insert(base, target);
        Ok(target)
    }

    /// Splice the new number format and cell formats into the styles part.
    fn finish(self) -> Result<String, EditError> {
        let mut splices: Vec<(Range<usize>, String)> = Vec::new();
        if self.defines_num_fmt {
            let section = self.scan.num_fmts.as_ref();
            let prefix = section.map_or(self.scan.prefix.as_deref(), |s| s.prefix.as_deref());
            let element = format!(
                r#"<{} numFmtId="{}" formatCode="{}"/>"#,
                qualify(prefix, "numFmt"),
                self.num_fmt_id,
                escape_attr(self.code)
            );
            match section {
                Some(section) => extend_section(
                    self.xml,
                    section,
                    self.scan.num_fmt_count + 1,
                    &element,
                    &mut splices,
                )?,
                None => {
                    // `numFmts` is the first child of `styleSheet`.
                    let name = qualify(prefix, "numFmts");
                    let at = self.scan.root_tag_end;
                    splices.push((at..at, format!(r#"<{name} count="1">{element}</{name}>"#)));
                }
            }
        }
        if !self.appended.is_empty() {
            let mut children = String::new();
            if self.synthesized_default {
                children.push_str(&self.xfs[0]);
            }
            for xf in &self.appended {
                children.push_str(xf);
            }
            extend_section(
                self.xml,
                self.cell_xfs,
                self.xfs.len() + self.appended.len(),
                &children,
                &mut splices,
            )?;
        }
        Ok(apply_splices(self.xml, splices))
    }
}

/// Copy of `source` (one `<xf>` element) using number format `num_fmt_id`.
///
/// Attribute order, the other attributes, and child elements are kept.
fn clone_xf(source: &str, num_fmt_id: u32) -> Result<String, EditError> {
    let (tag, self_closing, tag_end) = first_tag(source)?;
    let id = num_fmt_id.to_string();
    let overrides = [("numFmtId", id.as_str()), ("applyNumberFormat", "1")];
    let mut out = render_start_tag(&tag, &overrides, self_closing)?;
    if !self_closing {
        out.push_str(&source[tag_end..]);
    }
    Ok(out)
}

/// Comparison form of an element: attributes sorted with decoded values,
/// boolean spellings unified, insignificant whitespace and empty-element
/// syntax ignored.
fn canonical_element(text: &str) -> Result<String, EditError> {
    let mut reader = XmlReader::from_str(text);
    reader.config_mut().trim_text(false);
    let mut out = String::with_capacity(text.len());
    loop {
        match reader.read_event().map_err(styles_error)? {
            Event::Start(tag) => push_canonical_tag(&mut out, &tag)?,
            Event::Empty(tag) => {
                push_canonical_tag(&mut out, &tag)?;
                push_canonical_end(&mut out, tag.local_name().into_inner());
            }
            Event::End(tag) => push_canonical_end(&mut out, tag.local_name().into_inner()),
            Event::Text(content) => {
                let content = content.into_inner();
                out.push_str(content.trim());
            }
            Event::Eof => return Ok(out),
            _ => {}
        }
    }
}

fn push_canonical_tag(out: &mut String, tag: &BytesStart<'_>) -> Result<(), EditError> {
    let mut attrs: Vec<(&str, String)> = Vec::new();
    for attr in tag.attributes().with_checks(false) {
        let attr = attr.map_err(styles_error)?;
        let value = match quick_xml::escape::unescape(&attr.value) {
            Ok(value) => value.into_owned(),
            Err(_) => attr.value.to_string(),
        };
        let value = if value == "true" {
            "1".to_string()
        } else if value == "false" {
            "0".to_string()
        } else {
            value
        };
        attrs.push((attr.key.into_inner(), value));
    }
    attrs.sort_unstable();
    out.push('<');
    out.push_str(tag.local_name().into_inner());
    for (key, value) in attrs {
        out.push(' ');
        out.push_str(key);
        out.push_str("=\"");
        out.push_str(&escape_attr(&value));
        out.push('"');
    }
    out.push('>');
    Ok(())
}

fn push_canonical_end(out: &mut String, local: &str) {
    out.push_str("</");
    out.push_str(local);
    out.push('>');
}

// ---------------------------------------------------------------------------
// Styles scan
// ---------------------------------------------------------------------------

#[derive(Default)]
struct StylesScan {
    /// Namespace prefix of the `styleSheet` root, if any.
    prefix: Option<String>,
    /// Byte offset just past the root start tag.
    root_tag_end: usize,
    root_self_closing: bool,
    num_fmts: Option<Section>,
    /// `<numFmt>` children of `<numFmts>`.
    num_fmt_count: usize,
    /// Decoded `(numFmtId, formatCode)` pairs from `<numFmts>`.
    custom_formats: Vec<(u32, String)>,
    /// Highest number format id defined or referenced anywhere in the part.
    max_num_fmt_id: Option<u32>,
    cell_xfs: Option<Section>,
    xfs: Vec<XfSpan>,
}

impl StylesScan {
    fn note_num_fmt_id(&mut self, id: u32) {
        self.max_num_fmt_id = Some(self.max_num_fmt_id.map_or(id, |max| max.max(id)));
    }
}

/// A `numFmts` or `cellXfs` collection.
struct Section {
    /// Start tag span, or the whole element when self-closing.
    tag: Range<usize>,
    /// Offset of the end tag; `None` when self-closing.
    close: Option<usize>,
    prefix: Option<String>,
}

/// One `<xf>` in `<cellXfs>`.
struct XfSpan {
    outer: Range<usize>,
    num_fmt_id: u32,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Collection {
    NumFmts,
    CellXfs,
    Other,
}

fn scan_styles(xml: &str) -> Result<StylesScan, EditError> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut scan = StylesScan::default();
    let mut depth = 0usize;
    let mut collection = Collection::Other;
    let mut saw_root = false;
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(styles_error)?;
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(tag) => {
                match depth {
                    0 => {
                        open_root(&mut scan, &tag, end)?;
                        saw_root = true;
                    }
                    1 => collection = open_collection(&mut scan, &tag, start..end)?,
                    _ => match tag.local_name().into_inner() {
                        "numFmt" => record_num_fmt(
                            &mut scan,
                            &tag,
                            depth == 2 && collection == Collection::NumFmts,
                        ),
                        "xf" => {
                            let num_fmt_id = xf_num_fmt_id(&tag);
                            scan.note_num_fmt_id(num_fmt_id);
                            if depth == 2 && collection == Collection::CellXfs {
                                reader.read_to_end(tag.name()).map_err(styles_error)?;
                                let outer = start..reader.buffer_position() as usize;
                                scan.xfs.push(XfSpan { outer, num_fmt_id });
                                continue;
                            }
                        }
                        _ => {}
                    },
                }
                depth += 1;
            }
            Event::Empty(tag) => match depth {
                0 => {
                    open_root(&mut scan, &tag, end)?;
                    scan.root_self_closing = true;
                    return Ok(scan);
                }
                1 => {
                    open_collection(&mut scan, &tag, start..end)?;
                }
                _ => match tag.local_name().into_inner() {
                    "numFmt" => record_num_fmt(
                        &mut scan,
                        &tag,
                        depth == 2 && collection == Collection::NumFmts,
                    ),
                    "xf" => {
                        let num_fmt_id = xf_num_fmt_id(&tag);
                        scan.note_num_fmt_id(num_fmt_id);
                        if depth == 2 && collection == Collection::CellXfs {
                            scan.xfs.push(XfSpan {
                                outer: start..end,
                                num_fmt_id,
                            });
                        }
                    }
                    _ => {}
                },
            },
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if depth == 1 {
                    let section = match collection {
                        Collection::NumFmts => scan.num_fmts.as_mut(),
                        Collection::CellXfs => scan.cell_xfs.as_mut(),
                        Collection::Other => None,
                    };
                    if let Some(section) = section {
                        section.close = Some(start);
                    }
                    collection = Collection::Other;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    if saw_root {
        Ok(scan)
    } else {
        Err(EditError::Invalid(
            "styles part has no styleSheet root element".to_string(),
        ))
    }
}

fn open_root(scan: &mut StylesScan, tag: &BytesStart<'_>, tag_end: usize) -> Result<(), EditError> {
    let local = tag.local_name().into_inner();
    if local != "styleSheet" {
        return Err(EditError::Invalid(format!(
            "styles part root element is {local}, expected styleSheet"
        )));
    }
    scan.prefix = name_prefix(tag);
    scan.root_tag_end = tag_end;
    Ok(())
}

fn open_collection(
    scan: &mut StylesScan,
    tag: &BytesStart<'_>,
    span: Range<usize>,
) -> Result<Collection, EditError> {
    let local = tag.local_name().into_inner();
    let (slot, collection) = match local {
        "numFmts" => (&mut scan.num_fmts, Collection::NumFmts),
        "cellXfs" => (&mut scan.cell_xfs, Collection::CellXfs),
        _ => return Ok(Collection::Other),
    };
    if slot.is_some() {
        return Err(EditError::Invalid(format!(
            "styles part has more than one {local} collection"
        )));
    }
    *slot = Some(Section {
        tag: span,
        close: None,
        prefix: name_prefix(tag),
    });
    Ok(collection)
}

fn record_num_fmt(scan: &mut StylesScan, tag: &BytesStart<'_>, in_table: bool) {
    let id = attr_value(tag, b"numFmtId").and_then(|value| value.trim().parse::<u32>().ok());
    if let Some(id) = id {
        scan.note_num_fmt_id(id);
    }
    if in_table {
        scan.num_fmt_count += 1;
        if let (Some(id), Some(code)) = (id, attr_value(tag, b"formatCode")) {
            scan.custom_formats.push((id, code));
        }
    }
}

fn xf_num_fmt_id(tag: &BytesStart<'_>) -> u32 {
    attr_value(tag, b"numFmtId")
        .and_then(|value| value.trim().parse().ok())
        .unwrap_or(0)
}

// ---------------------------------------------------------------------------
// Worksheet scan
// ---------------------------------------------------------------------------

/// Styles that decide the base format of each cell in the target range.
struct SheetStyles {
    /// `<col min max style>` ranges, in document order.
    col_styles: Vec<(u32, u32, u32)>,
    rows: BTreeMap<u32, RowStyles>,
}

struct RowStyles {
    /// Row style applied to absent cells (only with `customFormat`).
    style: Option<u32>,
    /// `(column, style)` of existing cells inside the range, sorted by column.
    cells: Vec<(u32, u32)>,
}

impl SheetStyles {
    /// Style an absent cell inherits from its column, for each column in range.
    ///
    /// ECMA-376 Part 1, 18.3.1.13 (`col`): `style` applies to cells not yet
    /// allocated in the column. The first matching range wins.
    fn column_styles(&self, min_col: u32, max_col: u32) -> Vec<u32> {
        (min_col..=max_col)
            .map(|col| {
                self.col_styles
                    .iter()
                    .find(|&&(min, max, _)| (min..=max).contains(&col))
                    .map_or(0, |&(_, _, style)| style)
            })
            .collect()
    }
}

fn scan_worksheet(xml: &str, bounds: (u32, u32, u32, u32)) -> Result<SheetStyles, EditError> {
    let (min_row, min_col, max_row, max_col) = bounds;
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut grid = SheetStyles {
        col_styles: Vec::new(),
        rows: BTreeMap::new(),
    };
    let mut in_cols = false;
    let mut in_sheet_data = false;
    let mut saw_sheet_data = false;
    let mut open_row: Option<u32> = None;
    loop {
        let (tag, self_closing) = match reader.read_event().map_err(worksheet_error)? {
            Event::Start(tag) => (tag, false),
            Event::Empty(tag) => (tag, true),
            Event::End(tag) => {
                match tag.local_name().into_inner() {
                    "cols" => in_cols = false,
                    "row" => open_row = None,
                    "sheetData" => break,
                    _ => {}
                }
                continue;
            }
            Event::Eof => break,
            _ => continue,
        };
        match tag.local_name().into_inner() {
            "cols" if !in_sheet_data => in_cols = !self_closing,
            "col" if in_cols => push_col_style(&mut grid.col_styles, &tag),
            "sheetData" => {
                saw_sheet_data = true;
                if self_closing {
                    break;
                }
                in_sheet_data = true;
            }
            "row" if in_sheet_data => {
                let row = row_number(&tag)?;
                if (min_row..=max_row).contains(&row) {
                    grid.rows.entry(row).or_insert_with(|| RowStyles {
                        style: row_style(&tag),
                        cells: Vec::new(),
                    });
                    if !self_closing {
                        open_row = Some(row);
                    }
                } else if !self_closing {
                    reader.read_to_end(tag.name()).map_err(worksheet_error)?;
                }
            }
            "c" => {
                if let Some(row) = open_row {
                    let (col, style) = cell_style(&tag, row)?;
                    let in_range = (min_col..=max_col).contains(&col);
                    if let Some(info) = grid.rows.get_mut(&row).filter(|_| in_range) {
                        info.cells.push((col, style));
                    }
                    if !self_closing {
                        reader.read_to_end(tag.name()).map_err(worksheet_error)?;
                    }
                }
            }
            _ => {}
        }
    }
    if !saw_sheet_data {
        return Err(EditError::Invalid(
            "worksheet has no sheetData element".to_string(),
        ));
    }
    for info in grid.rows.values_mut() {
        info.cells.sort_by_key(|&(col, _)| col);
        info.cells.dedup_by_key(|&mut (col, _)| col);
    }
    Ok(grid)
}

fn row_number(tag: &BytesStart<'_>) -> Result<u32, EditError> {
    let value = attr_value(tag, b"r").ok_or_else(|| {
        EditError::Unsupported(
            "worksheet rows without an r attribute are not supported".to_string(),
        )
    })?;
    value
        .trim()
        .parse::<u32>()
        .ok()
        .filter(|row| (1..=MAX_ROW).contains(row))
        .ok_or_else(|| EditError::Invalid(format!("worksheet row number {value} is out of range")))
}

/// `(column, style)` of a cell stored in `row`.
fn cell_style(tag: &BytesStart<'_>, row: u32) -> Result<(u32, u32), EditError> {
    let reference = attr_value(tag, b"r").ok_or_else(|| {
        EditError::Unsupported(format!(
            "row {row} has a cell without an r attribute; implicit cell positions are not supported"
        ))
    })?;
    let (cell_row, col) = parse_a1(&reference).ok_or_else(|| {
        EditError::Invalid(format!(
            "cell reference {reference} in row {row} is not valid"
        ))
    })?;
    if cell_row != row {
        return Err(EditError::Invalid(format!(
            "cell {reference} is stored in row {row}"
        )));
    }
    let style = attr_value(tag, b"s")
        .and_then(|value| value.trim().parse().ok())
        .unwrap_or(0);
    Ok((col, style))
}

/// Style an absent cell takes from its row.
///
/// ECMA-376 Part 1, 18.3.1.73 (`row`): `s` applies only when `customFormat`
/// is true. A formatted row wins over a formatted column.
fn row_style(row: &BytesStart<'_>) -> Option<u32> {
    let custom = attr_value(row, b"customFormat")?;
    if custom != "1" && custom != "true" {
        return None;
    }
    Some(
        attr_value(row, b"s")
            .and_then(|value| value.trim().parse().ok())
            .unwrap_or(0),
    )
}

fn push_col_style(col_styles: &mut Vec<(u32, u32, u32)>, col: &BytesStart<'_>) {
    let parse =
        |key: &[u8]| attr_value(col, key).and_then(|value| value.trim().parse::<u32>().ok());
    if let (Some(min), Some(max), Some(style)) = (parse(b"min"), parse(b"max"), parse(b"style")) {
        col_styles.push((min, max, style));
    }
}

// ---------------------------------------------------------------------------
// XML text helpers
// ---------------------------------------------------------------------------

/// First element tag in `text`, whether it is self-closing, and the offset
/// just past it.
fn first_tag(text: &str) -> Result<(BytesStart<'_>, bool, usize), EditError> {
    let mut reader = XmlReader::from_str(text);
    loop {
        match reader.read_event().map_err(styles_error)? {
            Event::Start(tag) => return Ok((tag, false, reader.buffer_position() as usize)),
            Event::Empty(tag) => return Ok((tag, true, reader.buffer_position() as usize)),
            Event::Eof => return Err(styles_error("expected an element")),
            _ => {}
        }
    }
}

/// Re-render a start tag, replacing or appending `overrides` (values already
/// escaped) and copying every other attribute's escaped value unchanged.
fn render_start_tag(
    tag: &BytesStart<'_>,
    overrides: &[(&str, &str)],
    self_closing: bool,
) -> Result<String, EditError> {
    let mut out = String::new();
    out.push('<');
    out.push_str(tag.name().into_inner());
    let mut applied = vec![false; overrides.len()];
    for attr in tag.attributes().with_checks(false) {
        let attr = attr.map_err(styles_error)?;
        let key = attr.key.into_inner();
        out.push(' ');
        out.push_str(key);
        out.push_str("=\"");
        match overrides.iter().position(|&(name, _)| name == key) {
            Some(index) => {
                applied[index] = true;
                out.push_str(overrides[index].1);
            }
            // Values quoted with apostrophes may hold a literal quote.
            None => out.push_str(&attr.value.replace('"', "&quot;")),
        }
        out.push('"');
    }
    for (&(name, value), done) in overrides.iter().zip(applied) {
        if !done {
            out.push(' ');
            out.push_str(name);
            out.push_str("=\"");
            out.push_str(value);
            out.push('"');
        }
    }
    out.push_str(if self_closing { "/>" } else { ">" });
    Ok(out)
}

/// Add `children` to a collection and set its `count` to `count`.
fn extend_section(
    xml: &str,
    section: &Section,
    count: usize,
    children: &str,
    splices: &mut Vec<(Range<usize>, String)>,
) -> Result<(), EditError> {
    let (tag, _, _) = first_tag(&xml[section.tag.clone()])?;
    let count = count.to_string();
    let open = render_start_tag(&tag, &[("count", count.as_str())], false)?;
    match section.close {
        Some(close) => {
            splices.push((section.tag.clone(), open));
            splices.push((close..close, children.to_string()));
        }
        None => {
            let name = tag.name().into_inner();
            splices.push((section.tag.clone(), format!("{open}{children}</{name}>")));
        }
    }
    Ok(())
}

/// Replace each range with its text; insertions at one offset keep their order.
fn apply_splices(xml: &str, mut splices: Vec<(Range<usize>, String)>) -> String {
    splices.sort_by_key(|(range, _)| range.start);
    let extra: usize = splices.iter().map(|(_, text)| text.len()).sum();
    let mut out = String::with_capacity(xml.len() + extra);
    let mut copied = 0;
    for (range, text) in splices {
        out.push_str(&xml[copied..range.start]);
        out.push_str(&text);
        copied = range.end;
    }
    out.push_str(&xml[copied..]);
    out
}

fn escape_attr(value: &str) -> Cow<'_, str> {
    if !value.contains(['&', '<', '>', '"', '\t', '\n', '\r']) {
        return Cow::Borrowed(value);
    }
    let mut out = String::with_capacity(value.len() + 16);
    for c in value.chars() {
        match c {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\t' => out.push_str("&#9;"),
            '\n' => out.push_str("&#10;"),
            '\r' => out.push_str("&#13;"),
            _ => out.push(c),
        }
    }
    Cow::Owned(out)
}

fn name_prefix(tag: &BytesStart<'_>) -> Option<String> {
    tag.name()
        .into_inner()
        .split_once(':')
        .map(|(prefix, _)| prefix.to_string())
}

fn qualify(prefix: Option<&str>, local: &str) -> String {
    match prefix {
        Some(prefix) if !prefix.is_empty() => format!("{prefix}:{local}"),
        _ => local.to_string(),
    }
}

fn styles_error(err: impl Display) -> EditError {
    EditError::Xml(format!("styles part: {err}"))
}

fn worksheet_error(err: impl Display) -> EditError {
    EditError::Xml(format!("worksheet: {err}"))
}

// ---------------------------------------------------------------------------
// References
// ---------------------------------------------------------------------------

/// A1 label of `bounds`, or an error when they are inverted or off the grid.
fn validate_bounds(bounds: (u32, u32, u32, u32)) -> Result<String, EditError> {
    let (min_row, min_col, max_row, max_col) = bounds;
    let rows = 1..=MAX_ROW;
    let cols = 1..=MAX_COL;
    if !rows.contains(&min_row)
        || !rows.contains(&max_row)
        || !cols.contains(&min_col)
        || !cols.contains(&max_col)
        || min_row > max_row
        || min_col > max_col
    {
        return Err(EditError::BadReference(format!(
            "rows {min_row}-{max_row}, columns {min_col}-{max_col} must be ordered and lie \
             within rows 1-{MAX_ROW} and columns 1-{MAX_COL}"
        )));
    }
    let start = cell_label(min_row, min_col);
    Ok(if (min_row, min_col) == (max_row, max_col) {
        start
    } else {
        format!("{start}:{}", cell_label(max_row, max_col))
    })
}

/// Parse a plain A1 reference into `(row, column)`.
fn parse_a1(reference: &str) -> Option<(u32, u32)> {
    let split = reference.find(|c: char| !c.is_ascii_alphabetic())?;
    let (letters, digits) = reference.split_at(split);
    if letters.is_empty() || letters.len() > 3 || !digits.bytes().all(|b| b.is_ascii_digit()) {
        return None;
    }
    let col = letters.bytes().fold(0u32, |acc, b| {
        acc * 26 + u32::from(b.to_ascii_uppercase() - b'A' + 1)
    });
    let row = digits.parse::<u32>().ok()?;
    ((1..=MAX_ROW).contains(&row) && (1..=MAX_COL).contains(&col)).then_some((row, col))
}

fn cell_label(row: u32, col: u32) -> String {
    let mut letters = Vec::new();
    let mut rest = col;
    while rest > 0 {
        rest -= 1;
        letters.push(char::from(b'A' + (rest % 26) as u8));
        rest /= 26;
    }
    let column: String = letters.iter().rev().collect();
    format!("{column}{row}")
}

#[cfg(test)]
mod tests {
    use super::*;

    const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    /// Three cell formats: default, bold font, and yellow fill with centered
    /// alignment.
    const STYLES: &str = concat!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
        "\n",
        r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
        r#"<numFmts count="1"><numFmt numFmtId="164" formatCode="0.000"/></numFmts>"#,
        r#"<fonts count="2"><font><sz val="11"/></font><font><b/><sz val="11"/></font></fonts>"#,
        r#"<fills count="3"><fill><patternFill patternType="none"/></fill>"#,
        r#"<fill><patternFill patternType="gray125"/></fill>"#,
        r#"<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill></fills>"#,
        r#"<borders count="1"><border/></borders>"#,
        r#"<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>"#,
        r#"<cellXfs count="3">"#,
        r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#,
        r#"<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>"#,
        r#"<xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1">"#,
        r#"<alignment horizontal="center"/></xf>"#,
        "</cellXfs>",
        r#"<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>"#,
        "</styleSheet>",
    );

    const SHEET: &str = concat!(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
        r#"<dimension ref="A1:B2"/><sheetData>"#,
        r#"<row r="1"><c r="A1" s="1"><v>42</v></c><c r="B1" s="2" t="s"><v>0</v></c></row>"#,
        r#"<row r="2"><c r="A2"><f>A1*2</f><v>84</v></c></row>"#,
        "</sheetData></worksheet>",
    );

    fn empty_sheet() -> String {
        format!(r#"<worksheet xmlns="{NS}"><sheetData/></worksheet>"#)
    }

    fn assert_well_formed(xml: &str) {
        let mut reader = XmlReader::from_str(xml);
        let (mut depth, mut roots) = (0i32, 0);
        loop {
            match reader.read_event() {
                Ok(Event::Start(_)) => {
                    roots += i32::from(depth == 0);
                    depth += 1;
                }
                Ok(Event::Empty(_)) => roots += i32::from(depth == 0),
                Ok(Event::End(_)) => depth -= 1,
                Ok(Event::Eof) => break,
                Ok(_) => {}
                Err(err) => panic!("malformed XML: {err}\n{xml}"),
            }
        }
        assert_eq!((depth, roots), (0, 1), "unbalanced XML: {xml}");
    }

    fn attrs_of(tag: &BytesStart<'_>) -> BTreeMap<String, String> {
        tag.attributes()
            .map(|attr| {
                let attr = attr.unwrap();
                let value = quick_xml::escape::unescape(&attr.value)
                    .unwrap()
                    .into_owned();
                (attr.key.into_inner().to_string(), value)
            })
            .collect()
    }

    fn attrs(pairs: &[(&str, &str)]) -> BTreeMap<String, String> {
        pairs
            .iter()
            .map(|&(key, value)| (key.to_string(), value.to_string()))
            .collect()
    }

    /// Attributes of the `<c>` element for `reference`, if it exists.
    fn cell(sheet: &str, reference: &str) -> Option<BTreeMap<String, String>> {
        let mut reader = XmlReader::from_str(sheet);
        loop {
            match reader.read_event().unwrap() {
                Event::Start(tag) | Event::Empty(tag)
                    if tag.local_name().into_inner() == "c"
                        && attr_value(&tag, b"r").as_deref() == Some(reference) =>
                {
                    return Some(attrs_of(&tag));
                }
                Event::Eof => return None,
                _ => {}
            }
        }
    }

    fn style_of(sheet: &str, reference: &str) -> Option<u32> {
        cell(sheet, reference).map(|attrs| attrs.get("s").map_or(0, |s| s.parse().unwrap()))
    }

    /// Attributes of each `<xf>` in `<cellXfs>`.
    fn cell_xfs(styles: &str) -> Vec<BTreeMap<String, String>> {
        let mut reader = XmlReader::from_str(styles);
        let mut inside = false;
        let mut out = Vec::new();
        loop {
            match reader.read_event().unwrap() {
                Event::Start(tag) if tag.local_name().into_inner() == "cellXfs" => inside = true,
                Event::End(tag) if tag.local_name().into_inner() == "cellXfs" => inside = false,
                Event::Start(tag) | Event::Empty(tag)
                    if inside && tag.local_name().into_inner() == "xf" =>
                {
                    out.push(attrs_of(&tag));
                }
                Event::Eof => return out,
                _ => {}
            }
        }
    }

    /// `(id, code)` of each `<numFmt>` in `<numFmts>`.
    fn num_fmts(styles: &str) -> Vec<(u32, String)> {
        let mut reader = XmlReader::from_str(styles);
        let mut inside = false;
        let mut out = Vec::new();
        loop {
            match reader.read_event().unwrap() {
                Event::Start(tag) if tag.local_name().into_inner() == "numFmts" => inside = true,
                Event::End(tag) if tag.local_name().into_inner() == "numFmts" => inside = false,
                Event::Start(tag) | Event::Empty(tag)
                    if inside && tag.local_name().into_inner() == "numFmt" =>
                {
                    let attrs = attrs_of(&tag);
                    out.push((
                        attrs["numFmtId"].parse().unwrap(),
                        attrs["formatCode"].clone(),
                    ));
                }
                Event::Eof => return out,
                _ => {}
            }
        }
    }

    #[test]
    fn distinct_base_styles_get_distinct_copies() {
        let (styles, sheet) =
            set_number_format(Some(STYLES), SHEET, (1, 1, 1, 2), "$#,##0.00").unwrap();
        assert_well_formed(&styles);
        assert_well_formed(&sheet);

        assert_eq!(
            num_fmts(&styles),
            vec![(164, "0.000".to_string()), (165, "$#,##0.00".to_string())]
        );
        assert!(styles.contains(r#"<numFmts count="2">"#));
        assert!(styles.contains(r#"<cellXfs count="5">"#));
        let xfs = cell_xfs(&styles);
        assert_eq!(xfs.len(), 5);
        assert_eq!(
            xfs[3],
            attrs(&[
                ("numFmtId", "165"),
                ("fontId", "1"),
                ("fillId", "0"),
                ("borderId", "0"),
                ("xfId", "0"),
                ("applyFont", "1"),
                ("applyNumberFormat", "1"),
            ])
        );
        assert_eq!(xfs[4]["fillId"], "2");
        assert_eq!(xfs[4]["applyFill"], "1");
        assert!(styles.contains(
            r#"<xf numFmtId="165" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1" applyNumberFormat="1"><alignment horizontal="center"/></xf>"#
        ));
        // Bytes before the first changed collection are untouched.
        let head = &STYLES[..STYLES.find("<numFmts").unwrap()];
        assert!(styles.starts_with(head));

        assert_eq!(style_of(&sheet, "A1"), Some(3));
        assert_eq!(style_of(&sheet, "B1"), Some(4));
        assert_eq!(cell(&sheet, "B1").unwrap()["t"], "s");
        assert!(sheet.contains("<v>42</v>"));
        assert!(sheet.contains("<v>0</v>"));
        // Outside the range nothing changes.
        assert_eq!(style_of(&sheet, "A2"), Some(0));
    }

    #[test]
    fn reuses_matching_number_format_and_cell_format() {
        // Same attributes as a formatted copy of format 1, in another order
        // and with a different boolean spelling.
        let styles_in = STYLES.replace(r#"<cellXfs count="3">"#, r#"<cellXfs count="4">"#).replace(
            "</cellXfs>",
            r#"<xf applyNumberFormat="true" applyFont="1" xfId="0" borderId="0" fillId="0" fontId="1" numFmtId="164"/></cellXfs>"#,
        );
        let (styles, sheet) =
            set_number_format(Some(styles_in.as_str()), SHEET, (1, 1, 1, 1), "0.000").unwrap();
        assert_eq!(styles, styles_in);
        assert_eq!(style_of(&sheet, "A1"), Some(3));
        assert!(sheet.contains("<v>42</v>"));
    }

    #[test]
    fn cells_sharing_a_base_share_one_new_format() {
        let (styles, sheet) = set_number_format(Some(STYLES), SHEET, (2, 1, 2, 3), "0.0%").unwrap();
        assert_well_formed(&styles);
        assert_well_formed(&sheet);
        assert_eq!(cell_xfs(&styles).len(), 4);
        for reference in ["A2", "B2", "C2"] {
            assert_eq!(style_of(&sheet, reference), Some(3), "{reference}");
        }
        assert!(sheet.contains("<f>A1*2</f>"));
        assert!(sheet.contains("<v>84</v>"));
    }

    #[test]
    fn formatting_again_is_a_no_op() {
        let (styles, sheet) =
            set_number_format(Some(STYLES), SHEET, (1, 1, 2, 2), "0.00%").unwrap();
        let (again_styles, again_sheet) =
            set_number_format(Some(styles.as_str()), &sheet, (1, 1, 2, 2), "0.00%").unwrap();
        assert_eq!(again_styles, styles);
        assert_eq!(again_sheet, sheet);
    }

    #[test]
    fn self_closing_collections_are_expanded() {
        let styles_in = format!(
            r#"<styleSheet xmlns="{NS}"><numFmts count="0"/><fonts count="1"><font/></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders><cellXfs count="0" /></styleSheet>"#
        );
        let (styles, sheet) = set_number_format(
            Some(styles_in.as_str()),
            &empty_sheet(),
            (3, 3, 3, 3),
            "0.0",
        )
        .unwrap();
        assert_well_formed(&styles);
        assert_well_formed(&sheet);
        assert!(styles
            .contains(r#"<numFmts count="1"><numFmt numFmtId="164" formatCode="0.0"/></numFmts>"#));
        assert!(styles.contains(concat!(
            r#"<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>"#,
            r#"<xf numFmtId="164" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/></cellXfs>"#,
        )));
        assert_eq!(style_of(&sheet, "C3"), Some(1));
    }

    #[test]
    fn missing_cells_keep_row_and_column_styles() {
        let sheet_in = format!(
            r#"<worksheet xmlns="{NS}"><cols><col min="2" max="2" width="12" style="2" customWidth="1"/></cols><sheetData><row r="3" s="1" customFormat="1"/><row r="5" s="1"><c r="A5"><v>1</v></c></row></sheetData></worksheet>"#
        );
        let (styles, sheet) =
            set_number_format(Some(STYLES), &sheet_in, (3, 1, 5, 2), "0.0").unwrap();
        assert_well_formed(&styles);
        assert_well_formed(&sheet);

        let xfs = cell_xfs(&styles);
        assert_eq!(xfs.len(), 6);
        assert!(xfs[3..].iter().all(|xf| xf["numFmtId"] == "165"));
        // Row 3 is formatted with style 1: it wins over the column style.
        assert_eq!(style_of(&sheet, "A3"), Some(3));
        assert_eq!(style_of(&sheet, "B3"), Some(3));
        assert_eq!(xfs[3]["fontId"], "1");
        // Row 4 is absent: column B carries style 2, column A nothing.
        assert_eq!(style_of(&sheet, "A4"), Some(4));
        assert_eq!(style_of(&sheet, "B4"), Some(5));
        assert_eq!(
            (xfs[4]["fontId"].as_str(), xfs[4]["fillId"].as_str()),
            ("0", "0")
        );
        assert_eq!(xfs[5]["fillId"], "2");
        // Row 5 has `s` without `customFormat`, so absent B5 follows column B.
        assert_eq!(style_of(&sheet, "A5"), Some(4));
        assert_eq!(style_of(&sheet, "B5"), Some(5));
        assert!(sheet.contains("<v>1</v>"));
        // The row and column definitions themselves are kept.
        assert!(sheet.contains(r#"style="2""#));
        assert!(sheet.contains(r#"customFormat="1""#));
    }

    #[test]
    fn generates_styles_when_the_part_is_missing() {
        let (styles, sheet) = set_number_format(None, &empty_sheet(), (1, 1, 1, 1), "0.0").unwrap();
        assert_well_formed(&styles);
        assert!(styles.starts_with("<?xml"));
        assert_eq!(num_fmts(&styles), vec![(164, "0.0".to_string())]);
        let xfs = cell_xfs(&styles);
        assert_eq!(xfs.len(), 2);
        assert_eq!(xfs[1]["numFmtId"], "164");
        assert_eq!(xfs[1]["xfId"], "0");
        assert!(styles.contains(r#"<cellXfs count="2">"#));
        assert_eq!(style_of(&sheet, "A1"), Some(1));
    }

    #[test]
    fn prefixed_style_sheet_gets_prefixed_entries() {
        let styles_in = format!(
            r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font/></x:fonts><x:cellXfs count="1"><x:xf numFmtId="0" fontId="0"/></x:cellXfs></x:styleSheet>"#
        );
        let (styles, sheet) = set_number_format(
            Some(styles_in.as_str()),
            &empty_sheet(),
            (1, 1, 1, 1),
            "0.0",
        )
        .unwrap();
        assert_well_formed(&styles);
        assert!(styles.contains(concat!(
            r#"<x:styleSheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
            r#"<x:numFmts count="1"><x:numFmt numFmtId="164" formatCode="0.0"/></x:numFmts><x:fonts"#,
        )));
        assert!(styles.contains(
            r#"<x:cellXfs count="2"><x:xf numFmtId="0" fontId="0"/><x:xf numFmtId="164" fontId="0" applyNumberFormat="1"/></x:cellXfs>"#
        ));
        assert_eq!(style_of(&sheet, "A1"), Some(1));
    }

    #[test]
    fn format_codes_match_decoded_and_new_ids_skip_every_used_id() {
        let styles_in = STYLES
            .replace(
                r#"<numFmts count="1"><numFmt numFmtId="164" formatCode="0.000"/></numFmts>"#,
                r#"<numFmts count="1"><numFmt numFmtId="170" formatCode="&quot;A&amp;B&quot;0"/></numFmts>"#,
            )
            .replace(
                "</cellStyles>",
                r#"</cellStyles><dxfs count="1"><dxf><numFmt numFmtId="200" formatCode="0.0"/></dxf></dxfs>"#,
            );
        let (styles, sheet) =
            set_number_format(Some(styles_in.as_str()), SHEET, (2, 1, 2, 1), "\"A&B\"0").unwrap();
        assert_eq!(num_fmts(&styles).len(), 1);
        assert_eq!(cell_xfs(&styles)[3]["numFmtId"], "170");

        let (styles, _) =
            set_number_format(Some(styles.as_str()), &sheet, (1, 1, 1, 1), "\"<x>\"0").unwrap();
        assert_well_formed(&styles);
        assert!(styles.contains(r#"<numFmt numFmtId="201" formatCode="&quot;&lt;x&gt;&quot;0"/>"#));
        assert_eq!(num_fmts(&styles)[1], (201, "\"<x>\"0".to_string()));
    }

    #[test]
    fn rejects_inverted_or_off_grid_bounds() {
        for bounds in [
            (5, 1, 2, 1),
            (1, 3, 1, 2),
            (0, 1, 1, 1),
            (1, 1, 1, MAX_COL + 1),
        ] {
            let err = set_number_format(Some(STYLES), SHEET, bounds, "0.0").unwrap_err();
            assert!(
                matches!(err, EditError::BadReference(_)),
                "{bounds:?}: {err:?}"
            );
        }
    }

    #[test]
    fn rejects_cells_that_reference_missing_formats() {
        let styles_in = STYLES
            .replace(r#"<cellXfs count="3">"#, r#"<cellXfs count="1">"#)
            .replace(
                concat!(
                    r#"<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>"#,
                    r#"<xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1">"#,
                    r#"<alignment horizontal="center"/></xf>"#,
                ),
                "",
            );
        let err =
            set_number_format(Some(styles_in.as_str()), SHEET, (1, 1, 1, 1), "0.0").unwrap_err();
        assert!(
            matches!(&err, EditError::Invalid(message) if message.contains("cell A1") && message.contains("format 1")),
            "{err:?}"
        );
    }

    #[test]
    fn rejects_styles_without_cell_formats() {
        let styles_in =
            format!(r#"<styleSheet xmlns="{NS}"><fonts count="1"><font/></fonts></styleSheet>"#);
        let err = set_number_format(
            Some(styles_in.as_str()),
            &empty_sheet(),
            (1, 1, 1, 1),
            "0.0",
        )
        .unwrap_err();
        assert!(
            matches!(&err, EditError::Invalid(message) if message.contains("cellXfs")),
            "{err:?}"
        );
    }

    #[test]
    fn rejects_cells_without_references_in_range() {
        let sheet_in = format!(
            r#"<worksheet xmlns="{NS}"><sheetData><row r="1"><c><v>1</v></c></row></sheetData></worksheet>"#
        );
        let err = set_number_format(Some(STYLES), &sheet_in, (1, 1, 1, 1), "0.0").unwrap_err();
        assert!(matches!(err, EditError::Unsupported(_)), "{err:?}");
    }
}
