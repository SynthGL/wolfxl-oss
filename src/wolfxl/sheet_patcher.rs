//! Worksheet XML stream-patcher.
//!
//! Takes a worksheet XML string and a set of cell patches, produces a new XML
//! string with those cells replaced or inserted.  Uses quick-xml's streaming
//! reader+writer to avoid building a full DOM.
//!
//! WolfXL uses **inline strings** (`t="str"`) for all new string values.  This
//! avoids modifying the shared string table for the common case.

use std::collections::{BTreeMap, BTreeSet};
use std::io::Write;

use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::Reader as XmlReader;
use quick_xml::Writer as XmlWriter;

use crate::ooxml_util::attr_value;

// ---------------------------------------------------------------------------
// Cell patch types
// ---------------------------------------------------------------------------

/// What to write into a cell.
#[derive(Debug, Clone)]
pub enum CellValue {
    /// Empty / blank cell.
    Blank,
    /// Numeric value (integer or float).
    Number(f64),
    /// String value — written as inline string (`t="str"`).
    String(String),
    /// Boolean value.
    Boolean(bool),
    /// Excel error literal such as `#N/A`.
    Error(String),
    /// Formula string (e.g. `"SUM(A1:A2)"`).
    Formula(String),
    /// Rich-text runs — emitted as `t="inlineStr"` with `<is>...</is>`
    /// containing one `<r><rPr/><t/></r>` per run.  Sprint Ι Pod-α.
    RichText(Vec<wolfxl_writer::rich_text::RichTextRun>),
    /// RFC-057 (Sprint Ο Pod 1C): array-formula master cell.
    /// Emitted as `<c r="..."><f t="array" ref="...">text</f></c>`.
    ArrayFormula {
        /// Spill / array range, e.g. `"A1:A10"`.
        ref_range: String,
        /// Formula body without the leading `=` and without the
        /// surrounding `{}` braces.
        text: String,
    },
    /// RFC-057: data-table formula master cell.
    DataTableFormula {
        ref_range: String,
        ca: bool,
        dt2_d: bool,
        dtr: bool,
        r1: Option<String>,
        r2: Option<String>,
        del1: bool,
        del2: bool,
    },
    /// RFC-057: bare placeholder cell that lives inside an
    /// array-formula's spill range (everything except the master).
    /// Emitted as `<c r="..."/>`.
    SpillChild,
}

/// A single cell modification.
#[derive(Debug, Clone)]
pub struct CellPatch {
    /// 1-based row number.
    pub row: u32,
    /// 1-based column number.
    pub col: u32,
    /// New value (or None to keep existing).
    pub value: Option<CellValue>,
    /// New style index (or None to keep existing).
    pub style_index: Option<u32>,
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/// Patch a worksheet XML string with the given cell modifications.
///
/// Cells are replaced if they already exist, or inserted at the correct
/// sorted position if they don't.  Rows are created as needed.
///
/// The `shared_strings` table is used only to resolve existing shared string
/// values in cells that aren't being patched (for context — we don't modify
/// them).
pub fn patch_worksheet(xml: &str, patches: &[CellPatch]) -> Result<String, String> {
    if patches.is_empty() {
        return Ok(xml.to_string());
    }
    if let Some(out) = patch_existing_value_cells_prefix(xml, patches)? {
        return Ok(out);
    }

    // Group patches by row for efficient lookup.
    // Within each row, map col -> patch.
    let mut row_patches: BTreeMap<u32, BTreeMap<u32, &CellPatch>> = BTreeMap::new();
    for p in patches {
        row_patches.entry(p.row).or_default().insert(p.col, p);
    }
    let patch_bounds = bounds_for_patches(patches);

    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut writer = XmlWriter::new(Vec::new());
    let mut buf: Vec<u8> = Vec::new();

    // State tracking
    let mut in_sheet_data = false;
    let mut in_cols = false;
    // `<col style>` ranges as (min, max, style); `<cols>` precedes `<sheetData>`.
    let mut col_styles: Vec<(u32, u32, u32)> = Vec::new();
    let mut current_row: Option<u32> = None;
    let mut current_row_style: Option<u32> = None;
    let mut current_row_cols_seen: BTreeSet<u32> = BTreeSet::new();
    let mut rows_seen: BTreeSet<u32> = BTreeSet::new();
    let mut skip_until_cell_end = false; // skip children of a cell being replaced
    let mut worksheet_prefix: Option<String> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let tag = e.local_name().as_ref().to_vec();
                capture_prefix(&mut worksheet_prefix, e.name().as_ref(), &tag);

                if tag == b"dimension" {
                    let dim = rewrite_dimension(e, patch_bounds, worksheet_prefix.as_deref())?;
                    write_event(&mut writer, Event::Start(dim))?;
                } else if tag == b"sheetData" {
                    in_sheet_data = true;
                    write_event(&mut writer, Event::Start(e.to_owned()))?;
                } else if tag == b"cols" && !in_sheet_data {
                    in_cols = true;
                    write_event(&mut writer, Event::Start(e.to_owned()))?;
                } else if tag == b"col" && in_cols {
                    push_col_style(&mut col_styles, e);
                    write_event(&mut writer, Event::Start(e.to_owned()))?;
                } else if tag == b"row" && in_sheet_data {
                    let row_num = attr_value(e, b"r")
                        .and_then(|s| s.parse::<u32>().ok())
                        .unwrap_or(0);

                    // Insert any missing rows that should come before this one
                    for &pr in row_patches.keys() {
                        if pr < row_num && !rows_seen.contains(&pr) {
                            write_new_row(
                                &mut writer,
                                pr,
                                row_patches.get(&pr).unwrap(),
                                None,
                                &col_styles,
                                worksheet_prefix.as_deref(),
                            )?;
                            rows_seen.insert(pr);
                        }
                    }

                    current_row = Some(row_num);
                    current_row_style = row_style(e);
                    current_row_cols_seen.clear();
                    rows_seen.insert(row_num);
                    let row_start = rewrite_row_spans(
                        e,
                        row_patches.get(&row_num),
                        worksheet_prefix.as_deref(),
                    )?;
                    write_event(&mut writer, Event::Start(row_start))?;
                } else if tag == b"c" && in_sheet_data {
                    let cell_ref = attr_value(e, b"r").unwrap_or_default();
                    let (_, col) = parse_cell_ref(&cell_ref);

                    current_row_cols_seen.insert(col);

                    if let Some(row_map) = current_row.and_then(|r| row_patches.get(&r)) {
                        if let Some(patch) = row_map.get(&col) {
                            // This cell is being patched.
                            // If it's style-only (no value change), preserve the original
                            // children (<v>, <f>, etc.) and only rewrite the <c ...> attrs.
                            if patch.value.is_none() && patch.style_index.is_some() {
                                write_style_only_cell_start(
                                    &mut writer,
                                    &cell_ref,
                                    e,
                                    patch,
                                    worksheet_prefix.as_deref(),
                                )?;
                                // Do NOT skip children.
                            } else {
                                // Value patch: replace the entire cell element.
                                write_patched_cell(
                                    &mut writer,
                                    &cell_ref,
                                    e,
                                    patch,
                                    worksheet_prefix.as_deref(),
                                )?;
                                skip_until_cell_end = true;
                            }
                        } else {
                            // Not patched — pass through
                            write_event(&mut writer, Event::Start(e.to_owned()))?;
                        }
                    } else {
                        write_event(&mut writer, Event::Start(e.to_owned()))?;
                    }
                } else {
                    if !skip_until_cell_end {
                        write_event(&mut writer, Event::Start(e.to_owned()))?;
                    }
                }
            }
            Ok(Event::Empty(ref e)) => {
                let tag = e.local_name().as_ref().to_vec();
                capture_prefix(&mut worksheet_prefix, e.name().as_ref(), &tag);

                if tag == b"dimension" {
                    let dim = rewrite_dimension(e, patch_bounds, worksheet_prefix.as_deref())?;
                    write_event(&mut writer, Event::Empty(dim))?;
                } else if tag == b"col" && in_cols {
                    push_col_style(&mut col_styles, e);
                    write_event(&mut writer, Event::Empty(e.to_owned()))?;
                } else if tag == b"row" && in_sheet_data {
                    // Self-closing empty row — handle insertions
                    let row_num = attr_value(e, b"r")
                        .and_then(|s| s.parse::<u32>().ok())
                        .unwrap_or(0);

                    for &pr in row_patches.keys() {
                        if pr < row_num && !rows_seen.contains(&pr) {
                            write_new_row(
                                &mut writer,
                                pr,
                                row_patches.get(&pr).unwrap(),
                                None,
                                &col_styles,
                                worksheet_prefix.as_deref(),
                            )?;
                            rows_seen.insert(pr);
                        }
                    }
                    rows_seen.insert(row_num);

                    // If this empty row has patches, expand it, keeping its attributes.
                    if let Some(row_map) = row_patches.get(&row_num) {
                        write_new_row(
                            &mut writer,
                            row_num,
                            row_map,
                            Some(e),
                            &col_styles,
                            worksheet_prefix.as_deref(),
                        )?;
                    } else {
                        write_event(&mut writer, Event::Empty(e.to_owned()))?;
                    }
                } else if tag == b"c" && in_sheet_data {
                    // Self-closing cell (no value/formula children)
                    let cell_ref = attr_value(e, b"r").unwrap_or_default();
                    let (_, col) = parse_cell_ref(&cell_ref);

                    current_row_cols_seen.insert(col);

                    if let Some(row_map) = current_row.and_then(|r| row_patches.get(&r)) {
                        if let Some(patch) = row_map.get(&col) {
                            write_patched_cell(
                                &mut writer,
                                &cell_ref,
                                e,
                                patch,
                                worksheet_prefix.as_deref(),
                            )?;
                        } else {
                            write_event(&mut writer, Event::Empty(e.to_owned()))?;
                        }
                    } else {
                        write_event(&mut writer, Event::Empty(e.to_owned()))?;
                    }
                } else if tag == b"sheetData" {
                    // Empty <sheetData/> — need to insert all rows
                    let sheet_data_name = qname(worksheet_prefix.as_deref(), "sheetData");
                    let start = BytesStart::new(sheet_data_name.as_str());
                    write_event(&mut writer, Event::Start(start))?;
                    for (&row_num, row_map) in &row_patches {
                        write_new_row(
                            &mut writer,
                            row_num,
                            row_map,
                            None,
                            &col_styles,
                            worksheet_prefix.as_deref(),
                        )?;
                        rows_seen.insert(row_num);
                    }
                    write_event(
                        &mut writer,
                        Event::End(BytesEnd::new(sheet_data_name.as_str())),
                    )?;
                } else {
                    if !skip_until_cell_end {
                        write_event(&mut writer, Event::Empty(e.to_owned()))?;
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                let tag = e.local_name().as_ref().to_vec();

                if tag == b"c" && skip_until_cell_end {
                    skip_until_cell_end = false;
                    // Already wrote the replacement cell — don't write end tag
                } else if tag == b"cols" && in_cols {
                    in_cols = false;
                    write_event(&mut writer, Event::End(e.to_owned()))?;
                } else if tag == b"row" && in_sheet_data {
                    // Before closing row, insert any new cells for this row
                    if let Some(r) = current_row {
                        if let Some(row_map) = row_patches.get(&r) {
                            for (&col, patch) in row_map.iter() {
                                if !current_row_cols_seen.contains(&col) {
                                    let cell_ref = col_row_to_a1(col, r);
                                    write_new_cell(
                                        &mut writer,
                                        &cell_ref,
                                        patch,
                                        current_row_style.or_else(|| col_style(&col_styles, col)),
                                        worksheet_prefix.as_deref(),
                                    )?;
                                }
                            }
                        }
                    }
                    current_row = None;
                    current_row_style = None;
                    write_event(&mut writer, Event::End(e.to_owned()))?;
                } else if tag == b"sheetData" {
                    // Before closing sheetData, insert any remaining rows
                    for (&row_num, row_map) in &row_patches {
                        if !rows_seen.contains(&row_num) {
                            write_new_row(
                                &mut writer,
                                row_num,
                                row_map,
                                None,
                                &col_styles,
                                worksheet_prefix.as_deref(),
                            )?;
                        }
                    }
                    in_sheet_data = false;
                    write_event(&mut writer, Event::End(e.to_owned()))?;
                } else if !skip_until_cell_end {
                    write_event(&mut writer, Event::End(e.to_owned()))?;
                }
            }
            Ok(event @ Event::Text(_))
            | Ok(event @ Event::CData(_))
            | Ok(event @ Event::Comment(_))
            | Ok(event @ Event::Decl(_))
            | Ok(event @ Event::PI(_))
            | Ok(event @ Event::DocType(_)) => {
                if !skip_until_cell_end {
                    write_event(&mut writer, event.into_owned())?;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(format!("XML parse error: {e}")),
        }
        buf.clear();
    }

    let out = writer.into_inner();
    String::from_utf8(out).map_err(|e| format!("Output not UTF-8: {e}"))
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

fn patch_existing_value_cells_prefix(
    xml: &str,
    patches: &[CellPatch],
) -> Result<Option<String>, String> {
    if patches.iter().any(|patch| patch.value.is_none()) {
        return Ok(None);
    }

    let mut patch_map: BTreeMap<(u32, u32), &CellPatch> = BTreeMap::new();
    for patch in patches {
        patch_map.insert((patch.row, patch.col), patch);
    }

    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut writer = XmlWriter::new(Vec::new());
    let mut buf: Vec<u8> = Vec::new();
    let mut in_sheet_data = false;
    let mut skip_until_cell_end = false;
    let mut patched_count = 0usize;
    let mut worksheet_prefix: Option<String> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let tag = e.local_name().as_ref().to_vec();
                capture_prefix(&mut worksheet_prefix, e.name().as_ref(), &tag);
                if tag == b"sheetData" {
                    in_sheet_data = true;
                    write_event(&mut writer, Event::Start(e.to_owned()))?;
                } else if tag == b"c" && in_sheet_data {
                    let cell_ref = attr_value(e, b"r").unwrap_or_default();
                    let key = parse_cell_ref(&cell_ref);
                    if let Some(patch) = patch_map.get(&key) {
                        write_patched_cell(
                            &mut writer,
                            &cell_ref,
                            e,
                            patch,
                            worksheet_prefix.as_deref(),
                        )?;
                        patched_count += 1;
                        skip_until_cell_end = true;
                    } else {
                        write_event(&mut writer, Event::Start(e.to_owned()))?;
                    }
                } else if !skip_until_cell_end {
                    write_event(&mut writer, Event::Start(e.to_owned()))?;
                }
            }
            Ok(Event::Empty(ref e)) => {
                let tag = e.local_name().as_ref().to_vec();
                capture_prefix(&mut worksheet_prefix, e.name().as_ref(), &tag);
                if tag == b"sheetData" {
                    in_sheet_data = false;
                    write_event(&mut writer, Event::Empty(e.to_owned()))?;
                } else if tag == b"c" && in_sheet_data {
                    let cell_ref = attr_value(e, b"r").unwrap_or_default();
                    let key = parse_cell_ref(&cell_ref);
                    if let Some(patch) = patch_map.get(&key) {
                        write_patched_cell(
                            &mut writer,
                            &cell_ref,
                            e,
                            patch,
                            worksheet_prefix.as_deref(),
                        )?;
                        patched_count += 1;
                        if patched_count == patch_map.len() {
                            append_xml_tail(&mut writer, xml, reader.buffer_position() as usize)?;
                            let out = writer.into_inner();
                            return String::from_utf8(out)
                                .map(Some)
                                .map_err(|e| format!("Output not UTF-8: {e}"));
                        }
                    } else {
                        write_event(&mut writer, Event::Empty(e.to_owned()))?;
                    }
                } else if !skip_until_cell_end {
                    write_event(&mut writer, Event::Empty(e.to_owned()))?;
                }
            }
            Ok(Event::End(ref e)) => {
                let tag = e.local_name().as_ref().to_vec();
                if tag == b"c" && skip_until_cell_end {
                    skip_until_cell_end = false;
                    if patched_count == patch_map.len() {
                        append_xml_tail(&mut writer, xml, reader.buffer_position() as usize)?;
                        let out = writer.into_inner();
                        return String::from_utf8(out)
                            .map(Some)
                            .map_err(|e| format!("Output not UTF-8: {e}"));
                    }
                } else {
                    if tag == b"sheetData" {
                        in_sheet_data = false;
                    }
                    if !skip_until_cell_end {
                        write_event(&mut writer, Event::End(e.to_owned()))?;
                    }
                }
            }
            Ok(event @ Event::Text(_))
            | Ok(event @ Event::CData(_))
            | Ok(event @ Event::Comment(_))
            | Ok(event @ Event::Decl(_))
            | Ok(event @ Event::PI(_))
            | Ok(event @ Event::DocType(_)) => {
                if !skip_until_cell_end {
                    write_event(&mut writer, event.into_owned())?;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(format!("XML parse error: {e}")),
        }
        buf.clear();
    }

    Ok(None)
}

fn append_xml_tail<W: Write>(
    writer: &mut XmlWriter<W>,
    xml: &str,
    start: usize,
) -> Result<(), String> {
    writer
        .get_mut()
        .write_all(&xml.as_bytes()[start.min(xml.len())..])
        .map_err(|e| format!("XML write error: {e}"))
}

fn write_event<W: Write>(writer: &mut XmlWriter<W>, event: Event<'_>) -> Result<(), String> {
    writer
        .write_event(event)
        .map_err(|e| format!("XML write error: {e}"))
}

fn bounds_for_patches(patches: &[CellPatch]) -> Option<(u32, u32, u32, u32)> {
    let mut min_row = u32::MAX;
    let mut min_col = u32::MAX;
    let mut max_row = 0;
    let mut max_col = 0;
    for patch in patches {
        min_row = min_row.min(patch.row);
        min_col = min_col.min(patch.col);
        max_row = max_row.max(patch.row);
        max_col = max_col.max(patch.col);
    }
    if max_row == 0 || max_col == 0 {
        None
    } else {
        Some((min_row, min_col, max_row, max_col))
    }
}

fn rewrite_dimension(
    original: &BytesStart<'_>,
    patch_bounds: Option<(u32, u32, u32, u32)>,
    prefix: Option<&str>,
) -> Result<BytesStart<'static>, String> {
    let dimension_name = qname_for_original(original.name().as_ref(), "dimension", prefix);
    let mut elem = BytesStart::new(dimension_name);
    let original_ref = attr_value(original, b"ref");
    let merged_ref = original_ref
        .as_deref()
        .and_then(parse_dimension_ref)
        .map(|existing| merge_bounds(existing, patch_bounds))
        .or(patch_bounds)
        .map(format_dimension_ref);

    for a in original.attributes() {
        let a = a.map_err(|e| format!("XML attr error: {e}"))?;
        if a.key.as_ref() == b"ref" {
            continue;
        }
        elem.push_attribute((a.key.as_ref(), a.value.as_ref()));
    }
    if let Some(ref_value) = merged_ref {
        elem.push_attribute(("ref", ref_value.as_str()));
    }
    Ok(elem)
}

fn merge_bounds(
    existing: (u32, u32, u32, u32),
    patch_bounds: Option<(u32, u32, u32, u32)>,
) -> (u32, u32, u32, u32) {
    if let Some((patch_min_row, patch_min_col, patch_max_row, patch_max_col)) = patch_bounds {
        (
            existing.0.min(patch_min_row),
            existing.1.min(patch_min_col),
            existing.2.max(patch_max_row),
            existing.3.max(patch_max_col),
        )
    } else {
        existing
    }
}

fn parse_dimension_ref(value: &str) -> Option<(u32, u32, u32, u32)> {
    let (start, end) = value.split_once(':').unwrap_or((value, value));
    let (start_row, start_col) = parse_cell_ref(start);
    let (end_row, end_col) = parse_cell_ref(end);
    if start_row == 0 || start_col == 0 || end_row == 0 || end_col == 0 {
        None
    } else {
        Some((
            start_row.min(end_row),
            start_col.min(end_col),
            start_row.max(end_row),
            start_col.max(end_col),
        ))
    }
}

fn format_dimension_ref((min_row, min_col, max_row, max_col): (u32, u32, u32, u32)) -> String {
    let start = col_row_to_a1(min_col, min_row);
    let end = col_row_to_a1(max_col, max_row);
    if start == end {
        start
    } else {
        format!("{start}:{end}")
    }
}

fn rewrite_row_spans(
    original: &BytesStart<'_>,
    row_patches: Option<&BTreeMap<u32, &CellPatch>>,
    prefix: Option<&str>,
) -> Result<BytesStart<'static>, String> {
    let row_name = qname_for_original(original.name().as_ref(), "row", prefix);
    let mut elem = BytesStart::new(row_name);
    let patch_bounds = row_patches.map(|patches| {
        let min_col = patches.keys().copied().min().unwrap_or(1);
        let max_col = patches.keys().copied().max().unwrap_or(1);
        (min_col, max_col)
    });
    let original_spans = attr_value(original, b"spans");
    let merged_spans = original_spans
        .as_deref()
        .and_then(parse_spans)
        .map(|existing| merge_spans(existing, patch_bounds))
        .or(patch_bounds)
        .map(|(min_col, max_col)| format!("{min_col}:{max_col}"));

    for a in original.attributes() {
        let a = a.map_err(|e| format!("XML attr error: {e}"))?;
        if a.key.as_ref() == b"spans" {
            continue;
        }
        elem.push_attribute((a.key.as_ref(), a.value.as_ref()));
    }
    if let Some(spans) = merged_spans {
        elem.push_attribute(("spans", spans.as_str()));
    }
    Ok(elem)
}

fn parse_spans(value: &str) -> Option<(u32, u32)> {
    let (start, end) = value.split_once(':')?;
    let start_col = start.parse::<u32>().ok()?;
    let end_col = end.parse::<u32>().ok()?;
    if start_col == 0 || end_col == 0 {
        None
    } else {
        Some((start_col.min(end_col), start_col.max(end_col)))
    }
}

fn merge_spans(existing: (u32, u32), patch_bounds: Option<(u32, u32)>) -> (u32, u32) {
    if let Some((patch_min_col, patch_max_col)) = patch_bounds {
        (existing.0.min(patch_min_col), existing.1.max(patch_max_col))
    } else {
        existing
    }
}

/// Write a patched `<c ...>` start tag for a style-only patch, preserving original children.
fn write_style_only_cell_start<W: Write>(
    writer: &mut XmlWriter<W>,
    cell_ref: &str,
    original: &BytesStart<'_>,
    patch: &CellPatch,
    prefix: Option<&str>,
) -> Result<(), String> {
    let cell_name = qname_for_original(original.name().as_ref(), "c", prefix);
    let mut elem = BytesStart::new(cell_name.as_str());

    // Copy all original attributes except r/s. We'll re-add r and (patched) s.
    for a in original.attributes() {
        let a = a.map_err(|e| format!("XML attr error: {e}"))?;
        let key = a.key.as_ref();
        if key == b"r" || key == b"s" {
            continue;
        }
        elem.push_attribute((key, a.value.as_ref()));
    }

    // Always set r (cell reference)
    elem.push_attribute((b"r".as_slice(), cell_ref.as_bytes()));

    // Apply (or preserve) style index
    let style = if let Some(s) = patch.style_index {
        Some(s)
    } else {
        attr_value(original, b"s").and_then(|s| s.parse().ok())
    };
    if let Some(s) = style {
        if s > 0 {
            let sval = s.to_string();
            elem.push_attribute((b"s".as_slice(), sval.as_bytes()));
        }
    }

    write_event(writer, Event::Start(elem))
}

/// Write a complete patched cell element.
fn write_patched_cell<W: Write>(
    writer: &mut XmlWriter<W>,
    cell_ref: &str,
    original: &BytesStart<'_>,
    patch: &CellPatch,
    prefix: Option<&str>,
) -> Result<(), String> {
    let original_name = original.name();
    let effective_prefix = prefix_for_original(original_name.as_ref(), b"c").or(prefix);
    let cell_name = qname(effective_prefix, "c");
    let value_name = qname(effective_prefix, "v");
    let formula_name = qname(effective_prefix, "f");
    let inline_string_name = qname(effective_prefix, "is");

    let mut elem = BytesStart::new(cell_name.as_str());
    elem.push_attribute(("r", cell_ref));

    // Style index: use patch value if set, otherwise preserve original
    let style = if let Some(s) = patch.style_index {
        Some(s)
    } else {
        attr_value(original, b"s").and_then(|s| s.parse().ok())
    };
    if let Some(s) = style {
        if s > 0 {
            elem.push_attribute(("s", s.to_string().as_str()));
        }
    }

    match &patch.value {
        Some(CellValue::Blank) | None => {
            if patch.value.is_some() {
                // Explicit blank — write empty cell
                writer
                    .write_event(Event::Empty(elem))
                    .map_err(|e| format!("XML write error: {e}"))?;
            } else {
                // No value change — need to preserve original value
                // For simplicity, write the cell with original type
                // This path means only style changed, copy original attributes
                let orig_type = attr_value(original, b"t");
                if let Some(t) = &orig_type {
                    elem.push_attribute(("t", t.as_str()));
                }
                // Write as start tag, original children will follow via skip logic...
                // Actually, since skip_until_cell_end skips children, we need to
                // read and replay them.  For now, write empty if no value patch.
                writer
                    .write_event(Event::Empty(elem))
                    .map_err(|e| format!("XML write error: {e}"))?;
            }
        }
        Some(CellValue::Number(n)) => {
            writer
                .write_event(Event::Start(elem))
                .map_err(|e| format!("XML write error: {e}"))?;
            // <v>number</v>
            let v_start = BytesStart::new(value_name.as_str());
            writer
                .write_event(Event::Start(v_start))
                .map_err(|e| format!("XML write error: {e}"))?;
            let text = if *n == (*n as i64) as f64 {
                format!("{}", *n as i64)
            } else {
                format!("{n}")
            };
            writer
                .write_event(Event::Text(BytesText::new(&text)))
                .map_err(|e| format!("XML write error: {e}"))?;
            writer
                .write_event(Event::End(BytesEnd::new(value_name.as_str())))
                .map_err(|e| format!("XML write error: {e}"))?;
            writer
                .write_event(Event::End(BytesEnd::new(cell_name.as_str())))
                .map_err(|e| format!("XML write error: {e}"))?;
        }
        Some(CellValue::String(s)) => {
            elem.push_attribute(("t", "str"));
            writer
                .write_event(Event::Start(elem))
                .map_err(|e| format!("XML write error: {e}"))?;
            let v_start = BytesStart::new(value_name.as_str());
            writer
                .write_event(Event::Start(v_start))
                .map_err(|e| format!("XML write error: {e}"))?;
            writer
                .write_event(Event::Text(BytesText::new(s)))
                .map_err(|e| format!("XML write error: {e}"))?;
            writer
                .write_event(Event::End(BytesEnd::new(value_name.as_str())))
                .map_err(|e| format!("XML write error: {e}"))?;
            writer
                .write_event(Event::End(BytesEnd::new(cell_name.as_str())))
                .map_err(|e| format!("XML write error: {e}"))?;
        }
        Some(CellValue::Error(s)) => {
            elem.push_attribute(("t", "e"));
            writer
                .write_event(Event::Start(elem))
                .map_err(|e| format!("XML write error: {e}"))?;
            let v_start = BytesStart::new(value_name.as_str());
            writer
                .write_event(Event::Start(v_start))
                .map_err(|e| format!("XML write error: {e}"))?;
            writer
                .write_event(Event::Text(BytesText::new(s)))
                .map_err(|e| format!("XML write error: {e}"))?;
            writer
                .write_event(Event::End(BytesEnd::new(value_name.as_str())))
                .map_err(|e| format!("XML write error: {e}"))?;
            writer
                .write_event(Event::End(BytesEnd::new(cell_name.as_str())))
                .map_err(|e| format!("XML write error: {e}"))?;
        }
        Some(CellValue::Boolean(b)) => {
            elem.push_attribute(("t", "b"));
            writer
                .write_event(Event::Start(elem))
                .map_err(|e| format!("XML write error: {e}"))?;
            let v_start = BytesStart::new(value_name.as_str());
            writer
                .write_event(Event::Start(v_start))
                .map_err(|e| format!("XML write error: {e}"))?;
            let val = if *b { "1" } else { "0" };
            writer
                .write_event(Event::Text(BytesText::new(val)))
                .map_err(|e| format!("XML write error: {e}"))?;
            writer
                .write_event(Event::End(BytesEnd::new(value_name.as_str())))
                .map_err(|e| format!("XML write error: {e}"))?;
            writer
                .write_event(Event::End(BytesEnd::new(cell_name.as_str())))
                .map_err(|e| format!("XML write error: {e}"))?;
        }
        Some(CellValue::Formula(f)) => {
            writer
                .write_event(Event::Start(elem))
                .map_err(|e| format!("XML write error: {e}"))?;
            // <f>formula</f> — no <v> (force recalc)
            let f_start = BytesStart::new(formula_name.as_str());
            writer
                .write_event(Event::Start(f_start))
                .map_err(|e| format!("XML write error: {e}"))?;
            writer
                .write_event(Event::Text(BytesText::new(f)))
                .map_err(|e| format!("XML write error: {e}"))?;
            writer
                .write_event(Event::End(BytesEnd::new(formula_name.as_str())))
                .map_err(|e| format!("XML write error: {e}"))?;
            writer
                .write_event(Event::End(BytesEnd::new(cell_name.as_str())))
                .map_err(|e| format!("XML write error: {e}"))?;
        }
        Some(CellValue::ArrayFormula { ref_range, text }) => {
            // RFC-057: <c r="..."><f t="array" ref="A1:A10">B1:B10*2</f></c>
            writer
                .write_event(Event::Start(elem))
                .map_err(|e| format!("XML write error: {e}"))?;
            let mut f_start = BytesStart::new(formula_name.as_str());
            f_start.push_attribute(("t", "array"));
            f_start.push_attribute(("ref", ref_range.as_str()));
            writer
                .write_event(Event::Start(f_start))
                .map_err(|e| format!("XML write error: {e}"))?;
            writer
                .write_event(Event::Text(BytesText::new(text)))
                .map_err(|e| format!("XML write error: {e}"))?;
            writer
                .write_event(Event::End(BytesEnd::new(formula_name.as_str())))
                .map_err(|e| format!("XML write error: {e}"))?;
            writer
                .write_event(Event::End(BytesEnd::new(cell_name.as_str())))
                .map_err(|e| format!("XML write error: {e}"))?;
        }
        Some(CellValue::DataTableFormula {
            ref_range,
            ca,
            dt2_d,
            dtr,
            r1,
            r2,
            del1,
            del2,
        }) => {
            // RFC-057: <c r="..."><f t="dataTable" ref=".." dt2D="1" r1=".." r2=".."/></c>
            writer
                .write_event(Event::Start(elem))
                .map_err(|e| format!("XML write error: {e}"))?;
            let mut f_empty = BytesStart::new(formula_name.as_str());
            f_empty.push_attribute(("t", "dataTable"));
            f_empty.push_attribute(("ref", ref_range.as_str()));
            if *ca {
                f_empty.push_attribute(("ca", "1"));
            }
            if *dt2_d {
                f_empty.push_attribute(("dt2D", "1"));
            }
            if *dtr {
                f_empty.push_attribute(("dtr", "1"));
            }
            if let Some(rv) = r1.as_ref() {
                f_empty.push_attribute(("r1", rv.as_str()));
            }
            if let Some(rv) = r2.as_ref() {
                f_empty.push_attribute(("r2", rv.as_str()));
            }
            if *del1 {
                f_empty.push_attribute(("del1", "1"));
            }
            if *del2 {
                f_empty.push_attribute(("del2", "1"));
            }
            writer
                .write_event(Event::Empty(f_empty))
                .map_err(|e| format!("XML write error: {e}"))?;
            writer
                .write_event(Event::End(BytesEnd::new(cell_name.as_str())))
                .map_err(|e| format!("XML write error: {e}"))?;
        }
        Some(CellValue::SpillChild) => {
            // RFC-057: bare placeholder `<c r="..."/>`.  Style is
            // already on `elem` (set above when constructing the start tag).
            writer
                .write_event(Event::Empty(elem))
                .map_err(|e| format!("XML write error: {e}"))?;
        }
        Some(CellValue::RichText(runs)) => {
            // Sprint Ι Pod-α: rich-text writes use inline strings so
            // the SST never has to be touched.  This matches openpyxl's
            // own rich-text emit path (see the test fixture in the
            // Sprint Ι Pod-α task brief).
            elem.push_attribute(("t", "inlineStr"));
            // Emit the start tag, then hand-write the `<is>` body so
            // we can leverage the run-emitter from `wolfxl_writer::rich_text`
            // verbatim (xml-escape, xml:space=preserve, etc).
            writer
                .write_event(Event::Start(elem))
                .map_err(|e| format!("XML write error: {e}"))?;
            let body = wolfxl_writer::rich_text::emit_runs(runs);
            let raw = format!("<{inline_string_name}>{body}</{inline_string_name}>");
            // BytesText would re-escape; we want the run XML emitted
            // verbatim. quick-xml's `Writer::get_mut()` lets us drop in
            // raw bytes between events without breaking the surrounding
            // structure.
            writer
                .get_mut()
                .write_all(raw.as_bytes())
                .map_err(|e| format!("XML write error: {e}"))?;
            writer
                .write_event(Event::End(BytesEnd::new(cell_name.as_str())))
                .map_err(|e| format!("XML write error: {e}"))?;
        }
    }

    Ok(())
}

/// Style index an absent cell takes from its row or column.
///
/// ECMA-376 Part 1, 18.3.1.73 (`row`): `s` is the row's style and is applied
/// only when `customFormat` is true. 18.3.1.13 (`col`): `style` is the default
/// style for cells not yet allocated in the column. A formatted row wins over a
/// formatted column, which is the style Excel gives a cell typed into such a row.
fn row_style(row: &BytesStart<'_>) -> Option<u32> {
    let custom = attr_value(row, b"customFormat")?;
    if custom != "1" && custom != "true" {
        return None;
    }
    Some(
        attr_value(row, b"s")
            .and_then(|s| s.parse().ok())
            .unwrap_or(0),
    )
}

/// Record a `<col min max style>` range that carries a style.
fn push_col_style(col_styles: &mut Vec<(u32, u32, u32)>, col: &BytesStart<'_>) {
    let parse = |key: &[u8]| attr_value(col, key).and_then(|v| v.parse::<u32>().ok());
    if let (Some(min), Some(max), Some(style)) = (parse(b"min"), parse(b"max"), parse(b"style")) {
        col_styles.push((min, max, style));
    }
}

fn col_style(col_styles: &[(u32, u32, u32)], col: u32) -> Option<u32> {
    col_styles
        .iter()
        .find(|&&(min, max, _)| min <= col && col <= max)
        .map(|&(_, _, style)| style)
}

/// Write a brand-new cell element (insertion, not replacement).
///
/// `inherited_style` is the row or column style the absent cell had; it is
/// used when the patch carries no style of its own.
fn write_new_cell<W: Write>(
    writer: &mut XmlWriter<W>,
    cell_ref: &str,
    patch: &CellPatch,
    inherited_style: Option<u32>,
    prefix: Option<&str>,
) -> Result<(), String> {
    let mut absent = BytesStart::new("c");
    if let Some(s) = inherited_style {
        absent.push_attribute(("s", s.to_string().as_str()));
    }
    write_patched_cell(writer, cell_ref, &absent, patch, prefix)
}

/// Write a `<row>` element containing only patched cells.
///
/// `original` is the self-closing `<row/>` being expanded, if any; its
/// attributes are kept and its style is inherited by the new cells.
fn write_new_row<W: Write>(
    writer: &mut XmlWriter<W>,
    row_num: u32,
    cells: &BTreeMap<u32, &CellPatch>,
    original: Option<&BytesStart<'_>>,
    col_styles: &[(u32, u32, u32)],
    prefix: Option<&str>,
) -> Result<(), String> {
    let (row_elem, inherited_row_style) = match original {
        Some(original) => (
            rewrite_row_spans(original, Some(cells), prefix)?,
            row_style(original),
        ),
        None => {
            let mut row_elem = BytesStart::new(qname(prefix, "row"));
            row_elem.push_attribute(("r", row_num.to_string().as_str()));
            (row_elem, None)
        }
    };
    writer
        .write_event(Event::Start(row_elem.borrow()))
        .map_err(|e| format!("XML write error: {e}"))?;

    for (&col, patch) in cells {
        let cell_ref = col_row_to_a1(col, row_num);
        let inherited = inherited_row_style.or_else(|| col_style(col_styles, col));
        write_new_cell(writer, &cell_ref, patch, inherited, prefix)?;
    }

    writer
        .write_event(Event::End(row_elem.to_end()))
        .map_err(|e| format!("XML write error: {e}"))?;

    Ok(())
}

fn capture_prefix(prefix: &mut Option<String>, qname: &[u8], local: &[u8]) {
    if prefix.is_some() {
        return;
    }
    if let Some(found) = prefix_for_original(qname, local) {
        *prefix = Some(found.to_string());
    }
}

fn qname(prefix: Option<&str>, local: &str) -> String {
    match prefix {
        Some(prefix) if !prefix.is_empty() => format!("{prefix}:{local}"),
        _ => local.to_string(),
    }
}

fn qname_for_original(qname_bytes: &[u8], local: &str, fallback_prefix: Option<&str>) -> String {
    let local_bytes = local.as_bytes();
    qname(
        prefix_for_original(qname_bytes, local_bytes).or(fallback_prefix),
        local,
    )
}

fn prefix_for_original<'a>(qname: &'a [u8], local: &[u8]) -> Option<&'a str> {
    if !qname.ends_with(local) {
        return None;
    }
    let prefix_len = qname.len().checked_sub(local.len() + 1)?;
    if qname.get(prefix_len) != Some(&b':') {
        return None;
    }
    std::str::from_utf8(&qname[..prefix_len]).ok()
}

/// Parse a cell reference like "B3" into (row=3, col=2) — both 1-based.
fn parse_cell_ref(cell_ref: &str) -> (u32, u32) {
    let mut col: u32 = 0;
    let mut row_str = String::new();

    for ch in cell_ref.chars() {
        if ch.is_ascii_alphabetic() {
            col = col * 26 + (ch.to_ascii_uppercase() as u32 - b'A' as u32 + 1);
        } else if ch.is_ascii_digit() {
            row_str.push(ch);
        }
    }

    let row = row_str.parse::<u32>().unwrap_or(0);
    (row, col)
}

/// Convert 1-based (col, row) to A1-style reference.
fn col_row_to_a1(col: u32, row: u32) -> String {
    let mut letters = String::new();
    let mut c = col;
    while c > 0 {
        c -= 1;
        letters.insert(0, (b'A' + (c % 26) as u8) as char);
        c /= 26;
    }
    format!("{letters}{row}")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_cell_ref() {
        assert_eq!(parse_cell_ref("A1"), (1, 1));
        assert_eq!(parse_cell_ref("B3"), (3, 2));
        assert_eq!(parse_cell_ref("AA100"), (100, 27));
        assert_eq!(parse_cell_ref("Z1"), (1, 26));
    }

    #[test]
    fn test_col_row_to_a1() {
        assert_eq!(col_row_to_a1(1, 1), "A1");
        assert_eq!(col_row_to_a1(2, 3), "B3");
        assert_eq!(col_row_to_a1(27, 100), "AA100");
        assert_eq!(col_row_to_a1(26, 1), "Z1");
    }

    #[test]
    fn test_patch_replace_value() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet><sheetData>
<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>42</v></c></row>
</sheetData></worksheet>"#;

        let patches = vec![CellPatch {
            row: 1,
            col: 2, // B1
            value: Some(CellValue::Number(99.0)),
            style_index: None,
        }];

        let result = patch_worksheet(xml, &patches).unwrap();
        assert!(result.contains("<v>99</v>"));
        // A1 should be unchanged (though type=s is preserved)
        assert!(result.contains("r=\"A1\""));
    }

    #[test]
    fn test_fast_existing_value_patch_appends_tail() {
        let xml = r#"<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="B2" s="4"><v>2</v></c></row><row r="3"><c r="C3"><v>3</v></c></row><row r="9999"><c r="A9999"><v>9999</v></c></row></sheetData><extLst><ext uri="{tail}"/></extLst></worksheet>"#;
        let patches = vec![
            CellPatch {
                row: 2,
                col: 2,
                value: Some(CellValue::String("changed".to_string())),
                style_index: None,
            },
            CellPatch {
                row: 3,
                col: 3,
                value: Some(CellValue::Number(12345.0)),
                style_index: None,
            },
        ];

        let result = patch_existing_value_cells_prefix(xml, &patches)
            .unwrap()
            .expect("existing value patches should use prefix fast path");
        assert!(result.contains(r#"<c r="B2" s="4" t="str"><v>changed</v></c>"#));
        assert!(result.contains(r#"<c r="C3"><v>12345</v></c>"#));
        assert!(result.contains(r#"<row r="9999"><c r="A9999"><v>9999</v></c></row>"#));
        assert!(result.contains(r#"<ext uri="{tail}"/>"#));
    }

    #[test]
    fn test_fast_existing_value_patch_falls_back_for_missing_cell() {
        let xml = r#"<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>"#;
        let patches = vec![CellPatch {
            row: 1,
            col: 3,
            value: Some(CellValue::String("new".to_string())),
            style_index: None,
        }];

        assert!(patch_existing_value_cells_prefix(xml, &patches)
            .unwrap()
            .is_none());
    }

    #[test]
    fn test_fast_existing_value_patch_falls_back_for_style_only() {
        let xml = r#"<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>"#;
        let patches = vec![CellPatch {
            row: 1,
            col: 1,
            value: None,
            style_index: Some(2),
        }];

        assert!(patch_existing_value_cells_prefix(xml, &patches)
            .unwrap()
            .is_none());
    }

    #[test]
    fn test_patch_insert_new_cell() {
        let xml = r#"<worksheet><sheetData>
<row r="1"><c r="A1"><v>1</v></c></row>
</sheetData></worksheet>"#;

        let patches = vec![CellPatch {
            row: 1,
            col: 3, // C1 — doesn't exist
            value: Some(CellValue::String("new".to_string())),
            style_index: None,
        }];

        let result = patch_worksheet(xml, &patches).unwrap();
        assert!(result.contains("r=\"C1\""));
        assert!(result.contains("t=\"str\""));
        assert!(result.contains("<v>new</v>"));
    }

    #[test]
    fn test_patch_insert_new_cell_expands_dimension() {
        let xml = r#"<worksheet><dimension ref="A1:K22"/><sheetData><row r="1" spans="1:11"><c r="A1"><v>1</v></c></row></sheetData></worksheet>"#;

        let patches = vec![CellPatch {
            row: 1,
            col: 26, // Z1 — outside the original A1:K22 dimension.
            value: Some(CellValue::String("new".to_string())),
            style_index: None,
        }];

        let result = patch_worksheet(xml, &patches).unwrap();
        assert!(result.contains(r#"<dimension ref="A1:Z22"/>"#), "{result}");
        assert!(result.contains(r#"<row r="1" spans="1:26">"#), "{result}");
        assert!(result.contains(r#"r="Z1""#), "{result}");
    }

    #[test]
    fn test_patch_insert_new_cell_expands_prefixed_dimension() {
        let xml = r#"<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:dimension ref="B2:C3"/><x:sheetData><x:row r="2"><x:c r="B2"><x:v>1</x:v></x:c></x:row></x:sheetData></x:worksheet>"#;

        let patches = vec![CellPatch {
            row: 10,
            col: 1,
            value: Some(CellValue::String("new".to_string())),
            style_index: None,
        }];

        let result = patch_worksheet(xml, &patches).unwrap();
        assert!(
            result.contains(r#"<x:dimension ref="A2:C10"/>"#),
            "{result}"
        );
        assert!(result.contains(r#"<x:c r="A10" t="str"><x:v>new</x:v></x:c>"#));
    }

    #[test]
    fn test_patch_insert_new_cell_preserves_prefixed_namespace() {
        let xml = r#"<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:sheetData><x:row r="1"><x:c r="A1" t="s"><x:v>0</x:v></x:c></x:row></x:sheetData></x:worksheet>"#;

        let patches = vec![CellPatch {
            row: 1,
            col: 10, // J1 — doesn't exist
            value: Some(CellValue::String("wolfxl_modify_smoke".to_string())),
            style_index: None,
        }];

        let result = patch_worksheet(xml, &patches).unwrap();
        assert!(result.contains(r#"<x:c r="J1" t="str"><x:v>wolfxl_modify_smoke</x:v></x:c>"#));
        assert!(!result.contains(r#"<c r="J1""#));
    }

    #[test]
    fn test_patch_insert_new_row() {
        let xml = r#"<worksheet><sheetData>
<row r="1"><c r="A1"><v>1</v></c></row>
<row r="3"><c r="A3"><v>3</v></c></row>
</sheetData></worksheet>"#;

        let patches = vec![CellPatch {
            row: 2,
            col: 1, // A2 — row doesn't exist
            value: Some(CellValue::String("inserted".to_string())),
            style_index: None,
        }];

        let result = patch_worksheet(xml, &patches).unwrap();
        assert!(result.contains("r=\"A2\""));
        assert!(result.contains("<v>inserted</v>"));
        // Verify row ordering: row 1 before row 2 before row 3
        let pos_r1 = result.find("r=\"A1\"").unwrap();
        let pos_r2 = result.find("r=\"A2\"").unwrap();
        let pos_r3 = result.find("r=\"A3\"").unwrap();
        assert!(pos_r1 < pos_r2);
        assert!(pos_r2 < pos_r3);
    }

    #[test]
    fn test_patch_formula() {
        let xml = r#"<worksheet><sheetData>
<row r="1"><c r="A1"><v>10</v></c></row>
</sheetData></worksheet>"#;

        let patches = vec![CellPatch {
            row: 1,
            col: 1,
            value: Some(CellValue::Formula("SUM(B1:B10)".to_string())),
            style_index: None,
        }];

        let result = patch_worksheet(xml, &patches).unwrap();
        assert!(result.contains("<f>SUM(B1:B10)</f>"));
        // No <v> — forces recalculation
        assert!(!result.contains("<v>10</v>"));
    }

    #[test]
    fn test_patch_with_style() {
        let xml = r#"<worksheet><sheetData>
<row r="1"><c r="A1"><v>42</v></c></row>
</sheetData></worksheet>"#;

        let patches = vec![CellPatch {
            row: 1,
            col: 1,
            value: Some(CellValue::Number(42.0)),
            style_index: Some(5),
        }];

        let result = patch_worksheet(xml, &patches).unwrap();
        assert!(result.contains("s=\"5\""));
    }

    #[test]
    fn test_patch_boolean() {
        let xml = r#"<worksheet><sheetData>
<row r="1"><c r="A1"><v>0</v></c></row>
</sheetData></worksheet>"#;

        let patches = vec![CellPatch {
            row: 1,
            col: 1,
            value: Some(CellValue::Boolean(true)),
            style_index: None,
        }];

        let result = patch_worksheet(xml, &patches).unwrap();
        assert!(result.contains("t=\"b\""));
        assert!(result.contains("<v>1</v>"));
    }

    #[test]
    fn test_patch_empty_sheet_data() {
        let xml = r#"<worksheet><sheetData/></worksheet>"#;

        let patches = vec![CellPatch {
            row: 1,
            col: 1,
            value: Some(CellValue::String("hello".to_string())),
            style_index: None,
        }];

        let result = patch_worksheet(xml, &patches).unwrap();
        assert!(result.contains("r=\"A1\""));
        assert!(result.contains("<v>hello</v>"));
    }

    #[test]
    fn test_no_patches_returns_unchanged() {
        let xml = r#"<worksheet><sheetData>
<row r="1"><c r="A1"><v>42</v></c></row>
</sheetData></worksheet>"#;

        let result = patch_worksheet(xml, &[]).unwrap();
        assert_eq!(result, xml);
    }

    fn string_patch(row: u32, col: u32) -> CellPatch {
        CellPatch {
            row,
            col,
            value: Some(CellValue::String("new".to_string())),
            style_index: None,
        }
    }

    const STYLED_SHEET: &str = r#"<worksheet><cols><col min="1" max="2" width="9" style="7" customWidth="1"/><col min="4" max="4" style="8"/></cols><sheetData><row r="1" s="5" customFormat="1"><c r="C1" s="2"><v>1</v></c><c r="D1"><v>4</v></c></row><row r="2" s="6"><c r="C2"><v>2</v></c></row><row r="3"><c r="C3"><v>3</v></c></row><row r="4" s="9" customFormat="1" ht="20" customHeight="1"/></sheetData></worksheet>"#;

    #[test]
    fn new_cell_in_custom_format_row_takes_row_style_over_column_style() {
        let result = patch_worksheet(STYLED_SHEET, &[string_patch(1, 1)]).unwrap();
        assert!(
            result.contains(r#"<c r="A1" s="5" t="str"><v>new</v></c>"#),
            "{result}"
        );
    }

    #[test]
    fn new_cell_takes_column_style_when_row_style_is_not_custom() {
        // Row 2 has `s` but no customFormat, so the row style does not apply.
        let result =
            patch_worksheet(STYLED_SHEET, &[string_patch(2, 2), string_patch(3, 4)]).unwrap();
        assert!(
            result.contains(r#"<c r="B2" s="7" t="str"><v>new</v></c>"#),
            "{result}"
        );
        assert!(
            result.contains(r#"<c r="D3" s="8" t="str"><v>new</v></c>"#),
            "{result}"
        );
    }

    #[test]
    fn new_cell_without_row_or_column_style_takes_default_style() {
        let result = patch_worksheet(STYLED_SHEET, &[string_patch(3, 5)]).unwrap();
        assert!(
            result.contains(r#"<c r="E3" t="str"><v>new</v></c>"#),
            "{result}"
        );
    }

    #[test]
    fn new_cell_in_new_row_takes_column_style() {
        let result = patch_worksheet(STYLED_SHEET, &[string_patch(10, 1)]).unwrap();
        assert!(
            result.contains(r#"<row r="10"><c r="A10" s="7" t="str"><v>new</v></c></row>"#),
            "{result}"
        );
    }

    #[test]
    fn expanded_empty_row_keeps_attributes_and_gives_new_cell_row_style() {
        let result = patch_worksheet(STYLED_SHEET, &[string_patch(4, 1)]).unwrap();
        assert!(
            result.contains(
                r#"<row r="4" s="9" customFormat="1" ht="20" customHeight="1" spans="1:1"><c r="A4" s="9" t="str"><v>new</v></c></row>"#
            ),
            "{result}"
        );
    }

    #[test]
    fn new_cell_with_explicit_style_keeps_it() {
        let mut patch = string_patch(1, 1);
        patch.style_index = Some(3);
        let result = patch_worksheet(STYLED_SHEET, &[patch]).unwrap();
        assert!(
            result.contains(r#"<c r="A1" s="3" t="str"><v>new</v></c>"#),
            "{result}"
        );
    }

    #[test]
    fn existing_cells_keep_their_own_style_in_styled_rows_and_columns() {
        // C1 (s=2) and D1 (no s) sit in a custom-format row; D1 also sits in
        // a styled column. Rewriting them must not pick up either style.
        let result = patch_worksheet(
            STYLED_SHEET,
            &[string_patch(1, 3), string_patch(1, 4), string_patch(1, 1)],
        )
        .unwrap();
        assert!(
            result.contains(r#"<c r="C1" s="2" t="str"><v>new</v></c>"#),
            "{result}"
        );
        assert!(
            result.contains(r#"<c r="D1" t="str"><v>new</v></c>"#),
            "{result}"
        );
        assert!(
            result.contains(r#"<c r="A1" s="5" t="str"><v>new</v></c>"#),
            "{result}"
        );
    }
}
