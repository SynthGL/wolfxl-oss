//! Sprint Ο Pod 1B (RFC-056) — Phase 2.5o helpers.
//!
//! Two pure-Rust helpers used by the patcher's `do_save` to:
//!
//! 1. Read existing cell values from a sheet's XML in a given range
//!    (`extract_cell_grid`).
//! 2. Stamp `hidden="1"` onto specific `<row>` elements in already-
//!    patched sheet XML (`stamp_row_hidden`).
//!
//! Both walk the XML stream once via `quick_xml::Reader`/`Writer` and
//! are streaming-safe over multi-MB sheets.

use pyo3::exceptions::PyIOError;
use pyo3::PyErr;
use pyo3::PyResult;

use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader as XmlReader;
use quick_xml::Writer as XmlWriter;

use wolfxl_autofilter::Cell;

use crate::ooxml_util::{attr_value, resolve_general_ref};

/// Parse an A1-notation range like `"A1:D100"` into 1-based
/// `(start_row, end_row, start_col, end_col)`. Returns `None` if the
/// range is malformed or unbounded.
pub fn parse_a1_range(s: &str) -> Option<(u32, u32, u32, u32)> {
    let s = s.trim();
    let (lhs, rhs) = match s.split_once(':') {
        Some(parts) => parts,
        None => (s, s),
    };
    let (lc, lr) = parse_a1_cell(lhs)?;
    let (rc, rr) = parse_a1_cell(rhs)?;
    Some((lr.min(rr), lr.max(rr), lc.min(rc), lc.max(rc)))
}

fn parse_a1_cell(s: &str) -> Option<(u32, u32)> {
    let s = s.trim_start_matches('$');
    let mut col: u32 = 0;
    let mut col_chars = 0;
    let mut iter = s.chars().peekable();
    while let Some(&c) = iter.peek() {
        if c.is_ascii_alphabetic() {
            col = col * 26 + (c.to_ascii_uppercase() as u32 - b'A' as u32 + 1);
            col_chars += 1;
            iter.next();
        } else if c == '$' {
            iter.next();
        } else {
            break;
        }
    }
    if col_chars == 0 {
        return None;
    }
    let row_str: String = iter.collect();
    let row: u32 = row_str.parse().ok()?;
    Some((col, row))
}

/// Read a 2D cell-value grid out of a worksheet XML byte slice.
///
/// `start_row` / `end_row` are 1-based, inclusive. `start_col` /
/// `end_col` are 1-based, inclusive. Missing cells are returned as
/// `Cell::Empty`. Cells with unsupported types coerce to `Cell::Empty`.
///
/// String resolution: `t="s"` cells (shared-string indices) cannot
/// be resolved here without the SST; we conservatively return
/// `Cell::Empty` for them. Inline-string `t="inlineStr"` and
/// `t="str"` cells are read directly. This is acceptable for a v2.0
/// AutoFilter evaluator: callers who need full SST-aware filtering
/// must currently make sure the cells in question are written as
/// inline strings (the patcher does this for its own write path).
pub fn extract_cell_grid(
    sheet_xml: &[u8],
    start_row: u32,
    end_row: u32,
    start_col: u32,
    end_col: u32,
) -> PyResult<Vec<Vec<Cell>>> {
    if start_row > end_row || start_col > end_col {
        return Ok(Vec::new());
    }
    let n_rows = (end_row - start_row + 1) as usize;
    let n_cols = (end_col - start_col + 1) as usize;
    let mut grid: Vec<Vec<Cell>> = vec![vec![Cell::Empty; n_cols]; n_rows];

    let mut reader = XmlReader::from_reader(sheet_xml);
    reader.config_mut().trim_text(false);
    let mut buf: Vec<u8> = Vec::new();

    let mut current_row: u32 = 0;
    let mut current_col: u32 = 0;
    let mut current_type: String = String::new();
    let mut in_v = false;
    let mut in_t = false; // <is><t> for inline strings
    let mut text_acc: String = String::new();
    let mut in_target_cell = false;
    let mut current_row_target = false;

    loop {
        let event = reader
            .read_event_into(&mut buf)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("autofilter read XML: {e}")))?;
        match event {
            Event::Start(e) => {
                let name = e.local_name();
                let local = name.as_ref();
                if local == "row" {
                    current_row = attr_value(&e, b"r")
                        .and_then(|s| s.parse::<u32>().ok())
                        .unwrap_or(0);
                    current_row_target = current_row >= start_row && current_row <= end_row;
                } else if local == "c" && current_row_target {
                    let cell_ref = attr_value(&e, b"r").unwrap_or_default();
                    let (col, _) = parse_cell_ref_local(&cell_ref);
                    current_col = col;
                    current_type = attr_value(&e, b"t").unwrap_or_else(|| "n".to_string());
                    in_target_cell = col >= start_col && col <= end_col;
                } else if (local == "v" || local == "t") && in_target_cell {
                    if local == "v" {
                        in_v = true;
                    } else {
                        in_t = true;
                    }
                    text_acc.clear();
                }
            }
            Event::Text(t) => {
                if (in_v || in_t) && in_target_cell {
                    text_acc.push_str(t.as_ref());
                }
            }
            // quick-xml 0.42 emits references separately from adjacent text.
            Event::GeneralRef(reference) => {
                if (in_v || in_t) && in_target_cell {
                    let value = resolve_general_ref(&reference)
                        .map_err(|e| PyErr::new::<PyIOError, _>(format!("autofilter reference: {e}")))?;
                    text_acc.push_str(&value);
                }
            }
            Event::CData(t) => {
                if (in_v || in_t) && in_target_cell {
                    text_acc.push_str(t.as_ref());
                }
            }
            Event::End(e) => {
                let name = e.local_name();
                let local = name.as_ref();
                if local == "v" || local == "t" {
                    if in_target_cell {
                        // Decode according to the cell type.
                        let cell = decode_cell(&current_type, &text_acc);
                        let row_idx = (current_row - start_row) as usize;
                        let col_idx = (current_col - start_col) as usize;
                        if row_idx < n_rows && col_idx < n_cols {
                            // Only keep the first non-empty value (handles
                            // cases where both <v> and <is><t> appear).
                            if matches!(grid[row_idx][col_idx], Cell::Empty) {
                                grid[row_idx][col_idx] = cell;
                            }
                        }
                    }
                    in_v = false;
                    in_t = false;
                } else if local == "c" {
                    in_target_cell = false;
                }
            }
            Event::Empty(e) => {
                let name = e.local_name();
                let local = name.as_ref();
                if local == "row" {
                    // self-closing row: nothing to extract.
                } else if local == "c" {
                    // empty cell, leave as Empty.
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(grid)
}

fn parse_cell_ref_local(s: &str) -> (u32, u32) {
    let mut col: u32 = 0;
    let mut row: u32 = 0;
    let mut chars = s.chars().peekable();
    while let Some(&c) = chars.peek() {
        if c.is_ascii_alphabetic() {
            col = col * 26 + (c.to_ascii_uppercase() as u32 - b'A' as u32 + 1);
            chars.next();
        } else {
            break;
        }
    }
    for c in chars {
        if let Some(d) = c.to_digit(10) {
            row = row * 10 + d;
        }
    }
    (col, row)
}

fn decode_cell(t: &str, text: &str) -> Cell {
    if text.is_empty() {
        return Cell::Empty;
    }
    match t {
        "n" | "" => text.parse::<f64>().map(Cell::Number).unwrap_or(Cell::Empty),
        "b" => match text {
            "1" | "true" | "TRUE" => Cell::Bool(true),
            _ => Cell::Bool(false),
        },
        "str" | "inlineStr" => Cell::String(text.to_string()),
        "s" => {
            // Shared-string index: we don't have the SST here, so
            // store the index as a string for downstream string
            // matching. Conservative fallback: empty.
            Cell::Empty
        }
        "d" | "date" => text.parse::<f64>().map(Cell::Date).unwrap_or(Cell::Empty),
        "e" => Cell::String(text.to_string()),
        _ => Cell::String(text.to_string()),
    }
}

/// Re-stream the sheet XML, adding `hidden="1"` to any `<row r="N">`
/// element where `N` is in `hidden_rows`. Idempotent: if the row
/// already has `hidden="1"`, it's left alone (no double-attr).
pub fn stamp_row_hidden(sheet_xml: &[u8], hidden_rows: &[u32]) -> PyResult<Vec<u8>> {
    if hidden_rows.is_empty() {
        return Ok(sheet_xml.to_vec());
    }
    let row_set: std::collections::HashSet<u32> = hidden_rows.iter().copied().collect();

    let mut reader = XmlReader::from_reader(sheet_xml);
    reader.config_mut().trim_text(false);
    let mut writer = XmlWriter::new(Vec::with_capacity(sheet_xml.len() + 32));
    let mut buf: Vec<u8> = Vec::new();

    loop {
        let event = reader
            .read_event_into(&mut buf)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("stamp_row_hidden read: {e}")))?;
        match event {
            Event::Start(e) => {
                if e.local_name().as_ref() == "row" {
                    if let Some(rewritten) = maybe_set_row_hidden(&e, &row_set) {
                        writer
                            .write_event(Event::Start(rewritten))
                            .map_err(|e| PyErr::new::<PyIOError, _>(format!("write: {e}")))?;
                    } else {
                        writer
                            .write_event(Event::Start(e.borrow()))
                            .map_err(|e| PyErr::new::<PyIOError, _>(format!("write: {e}")))?;
                    }
                } else {
                    writer
                        .write_event(Event::Start(e.borrow()))
                        .map_err(|e| PyErr::new::<PyIOError, _>(format!("write: {e}")))?;
                }
            }
            Event::Empty(e) => {
                if e.local_name().as_ref() == "row" {
                    if let Some(rewritten) = maybe_set_row_hidden(&e, &row_set) {
                        writer
                            .write_event(Event::Empty(rewritten))
                            .map_err(|e| PyErr::new::<PyIOError, _>(format!("write: {e}")))?;
                    } else {
                        writer
                            .write_event(Event::Empty(e.borrow()))
                            .map_err(|e| PyErr::new::<PyIOError, _>(format!("write: {e}")))?;
                    }
                } else {
                    writer
                        .write_event(Event::Empty(e.borrow()))
                        .map_err(|e| PyErr::new::<PyIOError, _>(format!("write: {e}")))?;
                }
            }
            Event::Eof => break,
            other => {
                writer
                    .write_event(other)
                    .map_err(|e| PyErr::new::<PyIOError, _>(format!("write: {e}")))?;
            }
        }
        buf.clear();
    }

    // Now also INSERT any rows that are in row_set but absent from the
    // source XML (so the user can hide rows that have no cell content).
    let mut bytes = writer.into_inner();
    let absent: Vec<u32> = hidden_rows
        .iter()
        .copied()
        .filter(|r| !contains_row(&bytes, *r))
        .collect();
    if !absent.is_empty() {
        bytes = insert_hidden_only_rows(&bytes, &absent)?;
    }
    Ok(bytes)
}

fn maybe_set_row_hidden(
    e: &BytesStart<'_>,
    hidden_rows: &std::collections::HashSet<u32>,
) -> Option<BytesStart<'static>> {
    let r = attr_value(e, b"r").and_then(|s| s.parse::<u32>().ok())?;
    if !hidden_rows.contains(&r) {
        return None;
    }
    // Already hidden? leave alone.
    if attr_value(e, b"hidden").as_deref() == Some("1") {
        return None;
    }
    let mut new_e = BytesStart::new(e.name().as_ref().to_owned());
    let mut had_hidden = false;
    for attr in e.attributes().with_checks(false).flatten() {
        if attr.key.as_ref() == "hidden" {
            had_hidden = true;
            new_e.push_attribute(("hidden", "1"));
        } else {
            new_e.push_attribute(attr);
        }
    }
    if !had_hidden {
        new_e.push_attribute(("hidden", "1"));
    }
    Some(new_e)
}

fn contains_row(xml: &[u8], row_num: u32) -> bool {
    let needle = format!(r#"<row r="{row_num}""#);
    xml.windows(needle.len()).any(|w| w == needle.as_bytes())
}

/// Insert `<row r="N" hidden="1"/>` markers for any rows missing
/// from the source. Inserts inside `<sheetData>` in ascending order
/// relative to existing rows.
fn insert_hidden_only_rows(xml: &[u8], rows: &[u32]) -> PyResult<Vec<u8>> {
    let mut sorted: Vec<u32> = rows.to_vec();
    sorted.sort_unstable();

    let mut reader = XmlReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut writer = XmlWriter::new(Vec::with_capacity(xml.len() + sorted.len() * 32));
    let mut buf: Vec<u8> = Vec::new();
    let mut in_sheet_data = false;
    let mut emitted: std::collections::HashSet<u32> = std::collections::HashSet::new();

    loop {
        let event = reader
            .read_event_into(&mut buf)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("insert_hidden read: {e}")))?;
        match event {
            Event::Start(e) => {
                let name = e.local_name();
                let local = name.as_ref();
                if local == "sheetData" {
                    in_sheet_data = true;
                    writer
                        .write_event(Event::Start(e.borrow()))
                        .map_err(|e| PyErr::new::<PyIOError, _>(format!("w: {e}")))?;
                } else if local == "row" && in_sheet_data {
                    let r = attr_value(&e, b"r")
                        .and_then(|s| s.parse::<u32>().ok())
                        .unwrap_or(0);
                    // Emit any pending rows BEFORE this one.
                    for &p in &sorted {
                        if p < r && !emitted.contains(&p) {
                            write_hidden_row(&mut writer, p)?;
                            emitted.insert(p);
                        }
                    }
                    writer
                        .write_event(Event::Start(e.borrow()))
                        .map_err(|e| PyErr::new::<PyIOError, _>(format!("w: {e}")))?;
                } else {
                    writer
                        .write_event(Event::Start(e.borrow()))
                        .map_err(|e| PyErr::new::<PyIOError, _>(format!("w: {e}")))?;
                }
            }
            Event::Empty(e) => {
                let name = e.local_name();
                let local = name.as_ref();
                if local == "row" && in_sheet_data {
                    let r = attr_value(&e, b"r")
                        .and_then(|s| s.parse::<u32>().ok())
                        .unwrap_or(0);
                    for &p in &sorted {
                        if p < r && !emitted.contains(&p) {
                            write_hidden_row(&mut writer, p)?;
                            emitted.insert(p);
                        }
                    }
                    writer
                        .write_event(Event::Empty(e.borrow()))
                        .map_err(|e| PyErr::new::<PyIOError, _>(format!("w: {e}")))?;
                } else if local == "sheetData" {
                    // self-closing sheetData: emit all rows now
                    let opened = quick_xml::events::BytesStart::new("sheetData");
                    writer
                        .write_event(Event::Start(opened))
                        .map_err(|e| PyErr::new::<PyIOError, _>(format!("w: {e}")))?;
                    for &p in &sorted {
                        if !emitted.contains(&p) {
                            write_hidden_row(&mut writer, p)?;
                            emitted.insert(p);
                        }
                    }
                    writer
                        .write_event(Event::End(quick_xml::events::BytesEnd::new("sheetData")))
                        .map_err(|e| PyErr::new::<PyIOError, _>(format!("w: {e}")))?;
                } else {
                    writer
                        .write_event(Event::Empty(e.borrow()))
                        .map_err(|e| PyErr::new::<PyIOError, _>(format!("w: {e}")))?;
                }
            }
            Event::End(e) => {
                if e.local_name().as_ref() == "sheetData" && in_sheet_data {
                    for &p in &sorted {
                        if !emitted.contains(&p) {
                            write_hidden_row(&mut writer, p)?;
                            emitted.insert(p);
                        }
                    }
                    in_sheet_data = false;
                }
                writer
                    .write_event(Event::End(e))
                    .map_err(|e| PyErr::new::<PyIOError, _>(format!("w: {e}")))?;
            }
            Event::Eof => break,
            other => {
                writer
                    .write_event(other)
                    .map_err(|e| PyErr::new::<PyIOError, _>(format!("w: {e}")))?;
            }
        }
        buf.clear();
    }

    Ok(writer.into_inner())
}

fn write_hidden_row<W: std::io::Write>(writer: &mut XmlWriter<W>, row_num: u32) -> PyResult<()> {
    let mut start = quick_xml::events::BytesStart::new("row");
    let r_str = row_num.to_string();
    start.push_attribute(("r", r_str.as_str()));
    start.push_attribute(("hidden", "1"));
    writer
        .write_event(Event::Empty(start))
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("write hidden row: {e}")))?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn extract_cell_grid_aggregates_references_and_cdata() {
        let xml = br#"<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t> A &amp; B &#x3C; &#65;<![CDATA[ & raw]]> </t></is></c><c r="B1"><v>&#49;2</v></c></row></sheetData>"#;
        let grid = extract_cell_grid(xml, 1, 1, 1, 2).unwrap();
        assert!(matches!(&grid[0][0], Cell::String(s) if s == " A & B < A & raw "));
        assert!(matches!(grid[0][1], Cell::Number(n) if n == 12.0));
    }

    #[test]
    fn extract_cell_grid_rejects_unknown_references() {
        let xml = br#"<sheetData><row r="1"><c r="A1" t="str"><v>A&unknown;B</v></c></row></sheetData>"#;
        assert!(extract_cell_grid(xml, 1, 1, 1, 1).is_err());
    }

    #[test]
    fn stamp_row_hidden_preserves_escaped_attributes_and_text() {
        let xml = br#"<sheetData><row r="1" custom="A&amp;B"><c r="A1" t="str"><v>A&amp;B&#32;C</v></c></row></sheetData>"#;
        let out = stamp_row_hidden(xml, &[1]).unwrap();
        let text = std::str::from_utf8(&out).unwrap();
        assert!(text.contains(r#"custom="A&amp;B""#));
        assert!(text.contains("<v>A&amp;B&#32;C</v>"));
        assert!(!text.contains("&amp;amp;"));
    }

    #[test]
    fn parse_a1_cell_simple() {
        assert_eq!(parse_a1_cell("A1"), Some((1, 1)));
        assert_eq!(parse_a1_cell("$AB$10"), Some((28, 10)));
        assert_eq!(parse_a1_cell("foo"), None);
    }

    #[test]
    fn parse_a1_range_simple() {
        assert_eq!(parse_a1_range("A1:D5"), Some((1, 5, 1, 4)));
        assert_eq!(parse_a1_range("A1"), Some((1, 1, 1, 1)));
    }

    #[test]
    fn extract_cell_grid_basic() {
        let xml = br#"<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
<row r="1"><c r="A1" t="str"><v>Header</v></c></row>
<row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>10</v></c></row>
<row r="3"><c r="A3"><v>2</v></c><c r="B3"><v>20</v></c></row>
</sheetData></worksheet>"#;
        let grid = extract_cell_grid(xml, 2, 3, 1, 2).unwrap();
        assert_eq!(grid.len(), 2);
        assert_eq!(grid[0].len(), 2);
        assert!(matches!(grid[0][0], Cell::Number(n) if n == 1.0));
        assert!(matches!(grid[1][1], Cell::Number(n) if n == 20.0));
    }

    #[test]
    fn extract_cell_grid_inline_string() {
        let xml = br#"<sheetData>
<row r="1"><c r="A1" t="inlineStr"><is><t>red</t></is></c></row>
<row r="2"><c r="A2" t="inlineStr"><is><t>blue</t></is></c></row>
</sheetData>"#;
        let grid = extract_cell_grid(xml, 1, 2, 1, 1).unwrap();
        assert!(matches!(&grid[0][0], Cell::String(s) if s == "red"));
        assert!(matches!(&grid[1][0], Cell::String(s) if s == "blue"));
    }

    #[test]
    fn stamp_row_hidden_existing() {
        let xml = br#"<sheetData>
<row r="1"><c r="A1"><v>1</v></c></row>
<row r="2"><c r="A2"><v>2</v></c></row>
<row r="3"><c r="A3"><v>3</v></c></row>
</sheetData>"#;
        let out = stamp_row_hidden(xml, &[2]).unwrap();
        let s = std::str::from_utf8(&out).unwrap();
        assert!(s.contains(r#"<row r="2" hidden="1">"#));
        assert!(!s.contains(r#"<row r="1" hidden="1""#));
    }

    #[test]
    fn stamp_row_hidden_inserts_missing() {
        let xml = br#"<sheetData>
<row r="1"><c r="A1"><v>1</v></c></row>
</sheetData>"#;
        let out = stamp_row_hidden(xml, &[5]).unwrap();
        let s = std::str::from_utf8(&out).unwrap();
        assert!(s.contains(r#"<row r="5" hidden="1"/>"#));
    }

    #[test]
    fn stamp_row_hidden_idempotent() {
        let xml = br#"<sheetData><row r="2" hidden="1"><c r="A2"><v>1</v></c></row></sheetData>"#;
        let out = stamp_row_hidden(xml, &[2]).unwrap();
        let s = std::str::from_utf8(&out).unwrap();
        // Exactly one hidden="1" attr, not two.
        assert_eq!(s.matches(r#"hidden="1""#).count(), 1);
    }
}
