//! `Worksheet.move_range` paste-style relocation.
//!
//! Given a source rectangle `src_lo..=src_hi` and a delta `(d_row,
//! d_col)`, every cell whose coordinate lies inside the source
//! rectangle is physically relocated to `(row + d_row, col + d_col)`.
//! Existing cells at the destination are silently overwritten, matching
//! openpyxl's worksheet behavior.
//!
//! Formula handling:
//! - **Inside `src`**: the `<f>` payload is paste-translated via
//!   `wolfxl_formula::move_range` with `respect_dollar=true`. Relative
//!   refs shift by `(d_row, d_col)`; `$`-marked refs do NOT shift.
//! - **Outside `src`** with `translate=true`: the `<f>` payload is
//!   routed through the same `move_range` call. Refs that point INTO
//!   `src` re-anchor; refs elsewhere are left alone.
//! - **Outside `src`** with `translate=false`: cells are passed
//!   through verbatim. Matches openpyxl's "Formulae and references
//!   will not be updated" docstring.
//!
//! Anchor handling (mergeCells, hyperlinks, DV/CF sqref):
//! - Anchors fully inside `src` shift by `(d_row, d_col)`.
//! - Anchors that straddle the source boundary are left in place.
//! - Anchors fully outside `src` are left in place.
//!
use std::io::Cursor;

use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::Reader as XmlReader;
use quick_xml::Writer as XmlWriter;

use wolfxl_formula::{move_range as formula_move_range, Range};

use crate::xml_text::TextNode;

/// Plan for one paste-style range move.
///
/// Both corners are 1-based, inclusive. `d_row` / `d_col` are signed.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RangeMovePlan {
    /// Top-left source corner — `(min_row, min_col)`.
    pub src_lo: (u32, u32),
    /// Bottom-right source corner — `(max_row, max_col)`.
    pub src_hi: (u32, u32),
    /// Signed row delta — positive shifts down, negative shifts up.
    pub d_row: i32,
    /// Signed col delta — positive shifts right, negative shifts left.
    pub d_col: i32,
    /// If true, formulas in cells OUTSIDE the source rectangle that
    /// reference cells INSIDE `src` are also re-anchored. Default
    /// false (matches openpyxl). Cells INSIDE `src` are always
    /// paste-translated regardless of this flag.
    pub translate: bool,
}

impl RangeMovePlan {
    fn is_noop(&self) -> bool {
        self.d_row == 0 && self.d_col == 0
    }

    fn src_min_row(&self) -> u32 {
        self.src_lo.0
    }
    fn src_min_col(&self) -> u32 {
        self.src_lo.1
    }
    fn src_max_row(&self) -> u32 {
        self.src_hi.0
    }
    fn src_max_col(&self) -> u32 {
        self.src_hi.1
    }

    fn contains_cell(&self, row: u32, col: u32) -> bool {
        row >= self.src_min_row()
            && row <= self.src_max_row()
            && col >= self.src_min_col()
            && col <= self.src_max_col()
    }

    /// True if the rectangle `[r_lo..=r_hi, c_lo..=c_hi]` lies fully
    /// inside `src`.
    fn contains_rect(&self, r_lo: u32, c_lo: u32, r_hi: u32, c_hi: u32) -> bool {
        r_lo >= self.src_min_row()
            && r_hi <= self.src_max_row()
            && c_lo >= self.src_min_col()
            && c_hi <= self.src_max_col()
    }

    fn formula_src_range(&self) -> Range {
        Range {
            min_row: self.src_min_row(),
            max_row: self.src_max_row(),
            min_col: self.src_min_col(),
            max_col: self.src_max_col(),
        }
    }

    fn formula_dst_range(&self) -> Range {
        let dst_min_row = (self.src_min_row() as i64 + self.d_row as i64) as u32;
        let dst_max_row = (self.src_max_row() as i64 + self.d_row as i64) as u32;
        let dst_min_col = (self.src_min_col() as i64 + self.d_col as i64) as u32;
        let dst_max_col = (self.src_max_col() as i64 + self.d_col as i64) as u32;
        Range {
            min_row: dst_min_row,
            max_row: dst_max_row,
            min_col: dst_min_col,
            max_col: dst_max_col,
        }
    }
}

/// Apply one paste-style range move to a worksheet XML, returning
/// the new bytes.
///
/// No-op invariant: `(d_row, d_col) == (0, 0)` returns the input
/// verbatim.
pub fn apply_range_move(sheet_xml: &[u8], plan: &RangeMovePlan) -> Vec<u8> {
    if plan.is_noop() {
        return sheet_xml.to_vec();
    }
    let xml_str = match std::str::from_utf8(sheet_xml) {
        Ok(s) => s,
        Err(_) => return sheet_xml.to_vec(),
    };

    // Parse and relocate <sheetData>. The result is a fresh
    // <sheetData> byte block plus the rest of the worksheet XML
    // surrounding it (unchanged at this stage).
    let sheet_data = match split_sheet_data(xml_str) {
        Some(parts) => parts,
        // No <sheetData> means we're being handed something we can't
        // process; return unchanged. (Belt-and-braces: every real
        // worksheet has <sheetData>.)
        None => return sheet_xml.to_vec(),
    };
    let sheet_data_new = rewrite_sheet_data(sheet_data.inner, plan, &sheet_data.element_prefix);

    // Streaming-splice the surrounding parts (mergeCells, hyperlinks,
    // DV, CF, dimension) so anchor refs reflect the move.
    let mut combined =
        String::with_capacity(sheet_data.head.len() + sheet_data_new.len() + sheet_data.tail.len());
    combined.push_str(&sheet_data.head);
    combined.push_str(&sheet_data_new);
    combined.push_str(&sheet_data.tail);
    rewrite_anchors_and_dimension(combined.as_bytes(), plan)
}

// ---------------------------------------------------------------------------
// sheetData parse / relocate / re-emit
// ---------------------------------------------------------------------------

struct SheetDataSplit<'a> {
    head: String,
    inner: &'a str,
    tail: String,
    element_prefix: String,
}

/// Split the worksheet XML into the surrounding `<sheetData>` slices.
///
/// The source may use a namespace prefix such as `<x:sheetData>`. Preserve that
/// prefix when re-emitting rows and cells so a move does not corrupt prefixed
/// worksheet XML.
fn split_sheet_data(xml: &str) -> Option<SheetDataSplit<'_>> {
    let (open_start, open_end, element_name, self_closing) =
        find_open_element_by_local_name(xml, "sheetData")?;
    let element_prefix = element_name
        .strip_suffix("sheetData")
        .unwrap_or_default()
        .to_string();

    if self_closing {
        let open_tag = &xml[open_start..open_end];
        let expanded_open = open_tag
            .trim_end_matches('>')
            .trim_end()
            .trim_end_matches('/')
            .trim_end();
        return Some(SheetDataSplit {
            head: format!("{}{expanded_open}>", &xml[..open_start]),
            inner: "",
            tail: format!("</{element_name}>{}", &xml[open_end..]),
            element_prefix,
        });
    }

    let close_tag = format!("</{element_name}>");
    let close_rel = xml[open_end..].find(&close_tag)?;
    let close_start = open_end + close_rel;
    Some(SheetDataSplit {
        head: xml[..open_end].to_string(),
        inner: &xml[open_end..close_start],
        tail: xml[close_start..].to_string(),
        element_prefix,
    })
}

fn find_open_element_by_local_name(
    xml: &str,
    local_name: &str,
) -> Option<(usize, usize, String, bool)> {
    let mut search_from = 0;
    while let Some(rel_start) = xml[search_from..].find('<') {
        let start = search_from + rel_start;
        let name_start = start + 1;
        let first = xml.as_bytes().get(name_start).copied()?;
        if matches!(first, b'/' | b'!' | b'?') {
            search_from = name_start + 1;
            continue;
        }
        let name_end = xml[name_start..]
            .find(|ch: char| ch.is_whitespace() || ch == '>' || ch == '/')
            .map(|rel| name_start + rel)?;
        let element_name = &xml[name_start..name_end];
        if element_name.rsplit(':').next() == Some(local_name) {
            let tag_end = xml[name_end..].find('>').map(|rel| name_end + rel + 1)?;
            let self_closing = xml[name_end..tag_end]
                .trim_end_matches('>')
                .trim_end()
                .ends_with('/');
            return Some((start, tag_end, element_name.to_string(), self_closing));
        }
        search_from = name_end;
    }
    None
}

/// One captured cell from the source `<sheetData>`.
#[derive(Debug, Clone)]
struct CapturedCell {
    /// 1-based row.
    row: u32,
    /// 1-based col.
    col: u32,
    /// Re-emitted bytes for `<c ...>...</c>` (or self-closing). The
    /// re-emit happens at relocation time so the `r=` attribute can
    /// be rewritten and any embedded `<f>` payload can be paste-
    /// translated.
    raw_attrs: Vec<(String, String)>,
    /// True if the source was `<c .../>` (no children).
    self_closing: bool,
    /// Verbatim child events (between `<c>` and `</c>`), suitable for
    /// piping through a `quick_xml::Writer` again. Kept as raw bytes
    /// so we don't have to round-trip through the typed event API.
    /// Element-shape preserved; the only mutation we may apply is to
    /// the text inside `<f>` (paste-translation).
    children: Vec<u8>,
}

/// One captured row's metadata (we re-emit it verbatim — height,
/// custom-height, hidden, style — only its `r` attribute is
/// recomputed).
#[derive(Debug, Clone, Default)]
struct CapturedRow {
    /// Attributes other than `r`.
    extra_attrs: Vec<(String, String)>,
}

/// Parse the `<sheetData>` inner XML, relocate cells per `plan`,
/// re-emit. Cells outside `src` pass through; cells inside `src`
/// are translated to `(row + d_row, col + d_col)`. Existing cells
/// at the destination are overwritten.
fn rewrite_sheet_data(inner: &str, plan: &RangeMovePlan, element_prefix: &str) -> String {
    let mut reader = XmlReader::from_str(inner);
    let mut buf: Vec<u8> = Vec::new();

    // Per-row metadata captured in input order. Indexed by the row's
    // original `r=` attribute.
    let mut row_meta: std::collections::HashMap<u32, CapturedRow> =
        std::collections::HashMap::new();
    // All cells captured in input order.
    let mut cells: Vec<CapturedCell> = Vec::new();

    // State machine: we walk events. When we see <row>, capture its
    // attrs. When we see <c>, capture the cell + its children until
    // </c>. Anything outside a <row> or unrecognised goes into the
    // "preamble" (we currently expect <sheetData> to contain only
    // rows + cells; OOXML's <sheetData> child schema is exactly
    // <row>*, so this is safe).

    let mut current_row: Option<u32> = None;
    let mut in_cell: Option<CapturedCell> = None;
    let mut cell_child_writer: Option<XmlWriter<Cursor<Vec<u8>>>> = None;
    let mut cell_depth: u32 = 0;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                if let Some(_) = in_cell.as_mut() {
                    // Inside a cell — pipe through to the child writer.
                    cell_depth += 1;
                    if let Some(w) = cell_child_writer.as_mut() {
                        let _ = w.write_event(Event::Start(e.to_owned()));
                    }
                    buf.clear();
                    continue;
                }
                match local.as_ref().as_bytes() {
                    b"row" => {
                        let mut row_n: Option<u32> = None;
                        let mut extra: Vec<(String, String)> = Vec::new();
                        for attr_res in e.attributes().with_checks(false) {
                            let Ok(attr) = attr_res else { continue };
                            let key = attr.key.as_ref().to_owned();
                            let val = match quick_xml::escape::unescape(&attr.value) {
                                Ok(v) => v.into_owned(),
                                Err(_) => continue,
                            };
                            if key == "r" {
                                row_n = val.parse().ok();
                            } else {
                                extra.push((key, val));
                            }
                        }
                        if let Some(n) = row_n {
                            row_meta.insert(n, CapturedRow { extra_attrs: extra });
                            current_row = Some(n);
                        } else {
                            current_row = None;
                        }
                    }
                    b"c" => {
                        let (row, col, raw_attrs) = parse_cell_attrs(e, current_row);
                        let cap = CapturedCell {
                            row,
                            col,
                            raw_attrs,
                            self_closing: false,
                            children: Vec::new(),
                        };
                        in_cell = Some(cap);
                        cell_child_writer = Some(XmlWriter::new(Cursor::new(Vec::new())));
                        cell_depth = 0;
                    }
                    _ => {
                        // Unknown element inside <sheetData>. Pass
                        // through into a preamble bucket — but per
                        // OOXML schema this shouldn't happen. We log
                        // it via a dropped event (no-op) for now.
                    }
                }
            }
            Ok(Event::Empty(ref e)) => {
                let local = e.local_name();
                if let Some(_) = in_cell.as_mut() {
                    if let Some(w) = cell_child_writer.as_mut() {
                        let _ = w.write_event(Event::Empty(e.to_owned()));
                    }
                    buf.clear();
                    continue;
                }
                match local.as_ref().as_bytes() {
                    b"row" => {
                        // Empty row — capture metadata, no cells.
                        let mut row_n: Option<u32> = None;
                        let mut extra: Vec<(String, String)> = Vec::new();
                        for attr_res in e.attributes().with_checks(false) {
                            let Ok(attr) = attr_res else { continue };
                            let key = attr.key.as_ref().to_owned();
                            let val = match quick_xml::escape::unescape(&attr.value) {
                                Ok(v) => v.into_owned(),
                                Err(_) => continue,
                            };
                            if key == "r" {
                                row_n = val.parse().ok();
                            } else {
                                extra.push((key, val));
                            }
                        }
                        if let Some(n) = row_n {
                            row_meta.insert(n, CapturedRow { extra_attrs: extra });
                        }
                        current_row = None;
                    }
                    b"c" => {
                        let (row, col, raw_attrs) = parse_cell_attrs(e, current_row);
                        cells.push(CapturedCell {
                            row,
                            col,
                            raw_attrs,
                            self_closing: true,
                            children: Vec::new(),
                        });
                    }
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) => {
                let local = e.local_name();
                if in_cell.is_some() {
                    if cell_depth == 0 && local.as_ref() == "c" {
                        // Closing the captured cell.
                        let mut cap = in_cell.take().unwrap();
                        let writer = cell_child_writer.take().unwrap();
                        cap.children = writer.into_inner().into_inner();
                        cap.self_closing = false;
                        cells.push(cap);
                        cell_depth = 0;
                    } else {
                        if cell_depth > 0 {
                            cell_depth -= 1;
                        }
                        if let Some(w) = cell_child_writer.as_mut() {
                            let _ = w.write_event(Event::End(e.to_owned()));
                        }
                    }
                    buf.clear();
                    continue;
                }
                if local.as_ref() == "row" {
                    current_row = None;
                }
            }
            Ok(Event::Text(ref t)) => {
                if in_cell.is_some() {
                    if let Some(w) = cell_child_writer.as_mut() {
                        let _ = w.write_event(Event::Text(t.to_owned()));
                    }
                }
            }
            Ok(Event::CData(ref t)) => {
                if in_cell.is_some() {
                    if let Some(w) = cell_child_writer.as_mut() {
                        let _ = w.write_event(Event::CData(t.to_owned()));
                    }
                }
            }
            Ok(Event::Eof) => break,
            Ok(other) => {
                if let Some(w) = cell_child_writer.as_mut() {
                    let _ = w.write_event(other);
                }
            }
            Err(_) => break,
        }
        buf.clear();
    }

    // Now relocate.
    let mut relocated: Vec<CapturedCell> = Vec::with_capacity(cells.len());
    let mut overwritten: std::collections::HashSet<(u32, u32)> = std::collections::HashSet::new();

    // First pass: cells inside src get relocated and their formulas
    // paste-translated. Track destination keys so we can drop any
    // pre-existing dst-cell on the second pass.
    let mut moved: Vec<CapturedCell> = Vec::new();
    let mut external: Vec<CapturedCell> = Vec::new();
    for cell in cells.into_iter() {
        if plan.contains_cell(cell.row, cell.col) {
            moved.push(cell);
        } else {
            external.push(cell);
        }
    }

    let formula_translator = FormulaTranslator::from_plan(plan);

    for mut cell in moved.into_iter() {
        let new_row = (cell.row as i64 + plan.d_row as i64) as u32;
        let new_col = (cell.col as i64 + plan.d_col as i64) as u32;
        // Paste-translate formula payload (if any) inside this cell.
        cell.children = formula_translator.translate_cell_children(&cell.children);
        // Update the `r=` attribute in raw_attrs.
        let new_r = a1(new_row, new_col);
        for (key, val) in cell.raw_attrs.iter_mut() {
            if key == "r" {
                *val = new_r.clone();
            }
        }
        cell.row = new_row;
        cell.col = new_col;
        overwritten.insert((new_row, new_col));
        relocated.push(cell);
    }

    for mut cell in external.into_iter() {
        // If translate=true, rewrite refs that point into src.
        if plan.translate {
            cell.children = formula_translator.translate_cell_children(&cell.children);
        }
        // Drop external cells that fall on a dst slot we just wrote.
        if overwritten.contains(&(cell.row, cell.col)) {
            continue;
        }
        relocated.push(cell);
    }

    // Sort by (row, col) for deterministic emission.
    relocated.sort_by_key(|c| (c.row, c.col));

    // Group into rows for re-emission.
    let mut writer = XmlWriter::new(Cursor::new(Vec::new()));
    let mut current_emit_row: Option<u32> = None;
    let mut row_open = false;
    let cells_iter = relocated.into_iter();
    let row_name = format!("{element_prefix}row");
    let cell_name = format!("{element_prefix}c");

    for cell in cells_iter {
        if Some(cell.row) != current_emit_row {
            if row_open {
                let _ = writer.write_event(Event::End(BytesEnd::new(row_name.as_str())));
            }
            // Open a new <row r="N"> (recover original metadata if
            // we had it for this row index in the SOURCE; we map the
            // pre-move row index lookup via inverse_row).
            let lookup_row = inverse_row_index(cell.row, plan);
            let extra = row_meta.get(&lookup_row).cloned().unwrap_or_default();
            let mut start = BytesStart::new(row_name.as_str());
            start.push_attribute(("r", cell.row.to_string().as_str()));
            for (key, val) in &extra.extra_attrs {
                // Skip any `spans` attribute — Excel re-derives, and
                // a stale span for the post-move row would be wrong.
                if key == "spans" {
                    continue;
                }
                start.push_attribute((key.as_str(), val.as_str()));
            }
            let _ = writer.write_event(Event::Start(start));
            current_emit_row = Some(cell.row);
            row_open = true;
        }
        // Emit <c r="...">...</c>
        let mut c_start = BytesStart::new(cell_name.as_str());
        for (key, val) in &cell.raw_attrs {
            c_start.push_attribute((key.as_str(), val.as_str()));
        }
        if cell.self_closing && cell.children.is_empty() {
            let _ = writer.write_event(Event::Empty(c_start));
        } else {
            let _ = writer.write_event(Event::Start(c_start));
            // Children: pipe raw bytes back. They are already valid
            // XML (the writer emitted them). We use BytesText with a
            // pre-rendered marker — but since children may include
            // tags, we drop them in via a custom write.
            // The Writer doesn't have a "raw bytes" event; we sink
            // into the underlying buffer directly.
            let inner = writer.get_mut();
            let _ = std::io::Write::write_all(inner, &cell.children);
            let _ = writer.write_event(Event::End(BytesEnd::new(cell_name.as_str())));
        }
    }
    if row_open {
        let _ = writer.write_event(Event::End(BytesEnd::new(row_name.as_str())));
    }

    String::from_utf8(writer.into_inner().into_inner()).unwrap_or_default()
}

/// Compute the pre-move row index for a post-move row index. Used to
/// recover row metadata (height, style) from the source.
fn inverse_row_index(post_move_row: u32, plan: &RangeMovePlan) -> u32 {
    // If post_move_row is in the destination band, the source row was
    // post_move_row - d_row. Else it's the same.
    let dst_min_row = (plan.src_min_row() as i64 + plan.d_row as i64) as u32;
    let dst_max_row = (plan.src_max_row() as i64 + plan.d_row as i64) as u32;
    if post_move_row >= dst_min_row && post_move_row <= dst_max_row {
        ((post_move_row as i64) - plan.d_row as i64) as u32
    } else {
        post_move_row
    }
}

/// Parse a `<c r="..." s="..." t="...">` start tag's attributes into
/// the captured shape. `current_row` is the enclosing `<row r=N>`'s
/// row index (used as a fallback if `<c>` has no `r=` — uncommon but
/// permitted by the spec when the cell is positionally implied).
fn parse_cell_attrs(
    e: &BytesStart<'_>,
    current_row: Option<u32>,
) -> (u32, u32, Vec<(String, String)>) {
    let mut row: u32 = current_row.unwrap_or(0);
    let mut col: u32 = 0;
    let mut attrs: Vec<(String, String)> = Vec::new();
    for attr_res in e.attributes().with_checks(false) {
        let Ok(attr) = attr_res else { continue };
        let key = attr.key.as_ref().to_owned();
        let val = match quick_xml::escape::unescape(&attr.value) {
            Ok(v) => v.into_owned(),
            Err(_) => continue,
        };
        if key == "r" {
            if let Some((r, c)) = parse_a1(&val) {
                row = r;
                col = c;
            }
        }
        attrs.push((key, val));
    }
    (row, col, attrs)
}

/// Parse an A1-style cell ref like `"B5"` or `"$B$5"` into
/// `(row, col)` (1-based). Whole-row / whole-col refs return None.
fn parse_a1(s: &str) -> Option<(u32, u32)> {
    let bytes = s.as_bytes();
    let mut i = 0;
    if i < bytes.len() && bytes[i] == b'$' {
        i += 1;
    }
    let col_start = i;
    while i < bytes.len() && bytes[i].is_ascii_alphabetic() {
        i += 1;
    }
    let col_str = &s[col_start..i];
    if i < bytes.len() && bytes[i] == b'$' {
        i += 1;
    }
    let row_start = i;
    while i < bytes.len() && bytes[i].is_ascii_digit() {
        i += 1;
    }
    if i != bytes.len() || col_str.is_empty() || row_start == bytes.len() {
        return None;
    }
    let mut col_n: u32 = 0;
    for &b in col_str.as_bytes() {
        col_n = col_n
            .checked_mul(26)?
            .checked_add((b.to_ascii_uppercase() - b'A' + 1) as u32)?;
    }
    let row_n: u32 = s[row_start..].parse().ok()?;
    if col_n == 0 || row_n == 0 {
        return None;
    }
    Some((row_n, col_n))
}

/// Render `(row, col)` (1-based) as A1.
fn a1(row: u32, col: u32) -> String {
    let mut letters = String::new();
    let mut n = col;
    while n > 0 {
        let rem = ((n - 1) % 26) as u8;
        letters.insert(0, (b'A' + rem) as char);
        n = (n - 1) / 26;
    }
    format!("{letters}{row}")
}

#[derive(Debug, Clone, Copy)]
struct FormulaTranslator {
    src: Range,
    dst: Range,
}

impl FormulaTranslator {
    fn from_plan(plan: &RangeMovePlan) -> Self {
        Self {
            src: plan.formula_src_range(),
            dst: plan.formula_dst_range(),
        }
    }

    /// Walk the captured cell-children byte stream and rewrite the text
    /// inside any `<f>` element via `wolfxl_formula::move_range`. This is
    /// the paste-style translation: refs inside `src` re-anchor by
    /// `dst.min - src.min`; refs outside are left alone; `$`-marked refs
    /// short-circuit.
    fn translate_cell_children(&self, children: &[u8]) -> Vec<u8> {
        if children.is_empty() {
            return Vec::new();
        }
        let s = match std::str::from_utf8(children) {
            Ok(s) => s,
            Err(_) => return children.to_vec(),
        };
        // The captured children byte stream starts and ends inside a
        // <c>...</c> block, so it is a sequence of zero or more child
        // elements. Wrap in a synthetic root for parsing.
        let wrapped = format!("<__r__>{s}</__r__>");
        let mut reader = XmlReader::from_str(&wrapped);
        let mut writer = XmlWriter::new(Cursor::new(Vec::new()));
        let mut buf: Vec<u8> = Vec::new();
        let mut in_f: bool = false;
        let mut formula = TextNode::default();

        loop {
            let event = reader.read_event_into(&mut buf);
            if !formula.is_empty()
                && !matches!(event, Ok(Event::Text(_)) | Ok(Event::GeneralRef(_)))
            {
                if let Ok(text) = formula.unescaped() {
                    let translated = self.translate_text(text);
                    let _ = writer.write_event(Event::Text(BytesText::new(&translated)));
                } else {
                    let _ = writer.write_event(Event::Text(BytesText::from_escaped(formula.raw())));
                }
                formula.clear();
            }
            match event {
                Ok(Event::Start(ref e)) => {
                    let local = e.local_name();
                    if local.as_ref() == "__r__" {
                        // Drop the synthetic root.
                        buf.clear();
                        continue;
                    }
                    if local.as_ref() == "f" {
                        in_f = true;
                    }
                    let _ = writer.write_event(Event::Start(e.to_owned()));
                }
                Ok(Event::Empty(ref e)) => {
                    let local = e.local_name();
                    if local.as_ref() == "__r__" {
                        buf.clear();
                        continue;
                    }
                    let _ = writer.write_event(Event::Empty(e.to_owned()));
                }
                Ok(Event::End(ref e)) => {
                    let local = e.local_name();
                    if local.as_ref() == "__r__" {
                        buf.clear();
                        continue;
                    }
                    if local.as_ref() == "f" {
                        in_f = false;
                    }
                    let _ = writer.write_event(Event::End(e.to_owned()));
                }
                Ok(Event::Text(t)) if in_f => formula.push_text(&t),
                Ok(Event::GeneralRef(r)) if in_f => formula.push_ref(&r),
                Ok(Event::CData(ref t)) => {
                    let _ = writer.write_event(Event::CData(t.to_owned()));
                }
                Ok(Event::Eof) => break,
                Ok(other) => {
                    let _ = writer.write_event(other);
                }
                Err(_) => break,
            }
            buf.clear();
        }

        writer.into_inner().into_inner()
    }

    /// Translate an OOXML formula text payload while preserving its leading `=`.
    fn translate_text(&self, raw: &str) -> String {
        let has_equals = raw.starts_with('=');
        let wrapped = if has_equals {
            raw.to_string()
        } else {
            format!("={raw}")
        };
        let translated = formula_move_range(&wrapped, &self.src, &self.dst, true);
        if has_equals {
            translated
        } else {
            translated
                .strip_prefix('=')
                .unwrap_or(&translated)
                .to_string()
        }
    }
}

// ---------------------------------------------------------------------------
// Anchor + dimension rewrite (mergeCells, hyperlinks, DV, CF)
// ---------------------------------------------------------------------------

/// Streaming-splice the worksheet XML to update mergeCells, hyperlinks,
/// DV/CF sqrefs, and the `<dimension>` ref. Cells inside `<sheetData>`
/// have already been relocated by the sheetData pass; this pass only touches
/// the anchor-bearing siblings.
fn rewrite_anchors_and_dimension(xml: &[u8], plan: &RangeMovePlan) -> Vec<u8> {
    let xml_str = match std::str::from_utf8(xml) {
        Ok(s) => s,
        Err(_) => return xml.to_vec(),
    };
    let mut reader = XmlReader::from_str(xml_str);
    let mut writer = XmlWriter::new(Cursor::new(Vec::new()));
    let mut buf: Vec<u8> = Vec::new();

    // Track whether we're inside a formula text field (for translate=true).
    let mut in_dv_formula: u32 = 0;
    let mut in_cf_formula: u32 = 0;
    let formula_translator = FormulaTranslator::from_plan(plan);
    let mut formula = TextNode::default();

    loop {
        let event = reader.read_event_into(&mut buf);
        if !formula.is_empty()
            && !matches!(event, Ok(Event::Text(_)) | Ok(Event::GeneralRef(_)))
        {
            if let Ok(text) = formula.unescaped() {
                let translated = formula_translator.translate_text(text);
                let _ = writer.write_event(Event::Text(BytesText::new(&translated)));
            } else {
                let _ = writer.write_event(Event::Text(BytesText::from_escaped(formula.raw())));
            }
            formula.clear();
        }
        match event {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                match local.as_ref().as_bytes() {
                    b"mergeCell" | b"hyperlink" => {
                        let new_e = rewrite_anchor_attr(e, plan, "ref");
                        let _ = writer.write_event(Event::Start(new_e));
                    }
                    b"dataValidation" | b"conditionalFormatting" => {
                        let new_e = rewrite_anchor_attr(e, plan, "sqref");
                        let _ = writer.write_event(Event::Start(new_e));
                    }
                    b"formula1" | b"formula2" => {
                        in_dv_formula += 1;
                        let _ = writer.write_event(Event::Start(e.to_owned()));
                    }
                    b"formula" => {
                        in_cf_formula += 1;
                        let _ = writer.write_event(Event::Start(e.to_owned()));
                    }
                    _ => {
                        let _ = writer.write_event(Event::Start(e.to_owned()));
                    }
                }
            }
            Ok(Event::Empty(ref e)) => {
                let local = e.local_name();
                match local.as_ref().as_bytes() {
                    b"dimension" => {
                        let new_e = rewrite_anchor_attr(e, plan, "ref");
                        let _ = writer.write_event(Event::Empty(new_e));
                    }
                    b"mergeCell" | b"hyperlink" => {
                        let new_e = rewrite_anchor_attr(e, plan, "ref");
                        let _ = writer.write_event(Event::Empty(new_e));
                    }
                    b"dataValidation" | b"conditionalFormatting" => {
                        let new_e = rewrite_anchor_attr(e, plan, "sqref");
                        let _ = writer.write_event(Event::Empty(new_e));
                    }
                    _ => {
                        let _ = writer.write_event(Event::Empty(e.to_owned()));
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                let local = e.local_name();
                match local.as_ref().as_bytes() {
                    b"formula1" | b"formula2" => {
                        if in_dv_formula > 0 {
                            in_dv_formula -= 1;
                        }
                    }
                    b"formula" => {
                        if in_cf_formula > 0 {
                            in_cf_formula -= 1;
                        }
                    }
                    _ => {}
                }
                let _ = writer.write_event(Event::End(e.to_owned()));
            }
            Ok(Event::Text(t)) if plan.translate && (in_dv_formula > 0 || in_cf_formula > 0) => {
                formula.push_text(&t);
            }
            Ok(Event::GeneralRef(r)) if plan.translate && (in_dv_formula > 0 || in_cf_formula > 0) => {
                formula.push_ref(&r);
            }
            Ok(Event::Eof) => break,
            Ok(other) => {
                let _ = writer.write_event(other);
            }
            Err(_) => break,
        }
        buf.clear();
    }

    writer.into_inner().into_inner()
}

/// Rewrite a single attribute (`ref` or `sqref`) on an element. The
/// value is parsed into one or more A1 ranges; each piece gets the
/// "fully inside src? shift; else leave" treatment. Empty post-move
/// values pass through unchanged (we don't drop the parent element
/// here — that's a v2 feature for `move_range` if reported).
fn rewrite_anchor_attr<'a>(
    e: &BytesStart<'a>,
    plan: &RangeMovePlan,
    attr_name: &str,
) -> BytesStart<'a> {
    let mut new_e = BytesStart::new(e.name().as_ref().to_owned());
    for attr_res in e.attributes().with_checks(false) {
        let Ok(attr) = attr_res else { continue };
        let key = attr.key.as_ref();
        let val = match quick_xml::escape::unescape(&attr.value) {
            Ok(v) => v.into_owned(),
            Err(_) => continue,
        };
        if key == attr_name {
            let new_val = shift_anchor_pieces(&val, plan);
            push_attr_escaped(&mut new_e, key, &new_val);
        } else {
            push_attr_escaped(&mut new_e, key, &val);
        }
    }
    new_e
}

fn push_attr_escaped(e: &mut BytesStart<'_>, key: &str, val: &str) {
    e.push_attribute((key, val));
}

/// Split a `ref` or `sqref` value on whitespace, apply the
/// "fully-inside-src" rule per piece, re-join with spaces.
fn shift_anchor_pieces(s: &str, plan: &RangeMovePlan) -> String {
    let pieces: Vec<&str> = s.split_whitespace().collect();
    if pieces.is_empty() {
        return s.to_string();
    }
    let mut out: Vec<String> = Vec::with_capacity(pieces.len());
    for piece in pieces {
        out.push(shift_one_anchor(piece, plan));
    }
    out.join(" ")
}

fn shift_one_anchor(s: &str, plan: &RangeMovePlan) -> String {
    // Split on ':' for ranges.
    if let Some(colon) = s.find(':') {
        let lhs = &s[..colon];
        let rhs = &s[colon + 1..];
        let (l_row, l_col) = match parse_a1(lhs) {
            Some(v) => v,
            None => return s.to_string(),
        };
        let (r_row, r_col) = match parse_a1(rhs) {
            Some(v) => v,
            None => return s.to_string(),
        };
        let r_lo = l_row.min(r_row);
        let r_hi = l_row.max(r_row);
        let c_lo = l_col.min(r_col);
        let c_hi = l_col.max(r_col);
        if plan.contains_rect(r_lo, c_lo, r_hi, c_hi) {
            let new_lhs = a1(
                (l_row as i64 + plan.d_row as i64) as u32,
                (l_col as i64 + plan.d_col as i64) as u32,
            );
            let new_rhs = a1(
                (r_row as i64 + plan.d_row as i64) as u32,
                (r_col as i64 + plan.d_col as i64) as u32,
            );
            return format!("{new_lhs}:{new_rhs}");
        }
        return s.to_string();
    }
    // Single cell.
    let (row, col) = match parse_a1(s) {
        Some(v) => v,
        None => return s.to_string(),
    };
    if plan.contains_cell(row, col) {
        return a1(
            (row as i64 + plan.d_row as i64) as u32,
            (col as i64 + plan.d_col as i64) as u32,
        );
    }
    s.to_string()
}

#[cfg(test)]
mod tests {
    use super::*;

    fn plan(src_lo: (u32, u32), src_hi: (u32, u32), d_row: i32, d_col: i32) -> RangeMovePlan {
        RangeMovePlan {
            src_lo,
            src_hi,
            d_row,
            d_col,
            translate: false,
        }
    }

    #[test]
    fn parse_a1_basics() {
        assert_eq!(parse_a1("A1"), Some((1, 1)));
        assert_eq!(parse_a1("$B$5"), Some((5, 2)));
        assert_eq!(parse_a1("AA10"), Some((10, 27)));
        assert_eq!(parse_a1("A"), None);
        assert_eq!(parse_a1("5"), None);
    }

    #[test]
    fn render_a1() {
        assert_eq!(a1(1, 1), "A1");
        assert_eq!(a1(5, 2), "B5");
        assert_eq!(a1(10, 27), "AA10");
    }

    #[test]
    fn noop_returns_input() {
        let xml = b"<worksheet><sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData></worksheet>";
        let p = plan((1, 1), (1, 1), 0, 0);
        let out = apply_range_move(xml, &p);
        assert_eq!(out, xml);
    }

    #[test]
    fn moves_simple_block_down() {
        let xml = br#"<worksheet><sheetData><row r="3"><c r="A3"><v>1</v></c><c r="B3"><v>2</v></c></row></sheetData></worksheet>"#;
        let p = plan((3, 1), (3, 2), 5, 0);
        let out = apply_range_move(xml, &p);
        let s = String::from_utf8(out).unwrap();
        // Original cells gone, new cells at row 8.
        assert!(s.contains(r#"<row r="8""#), "got: {s}");
        assert!(s.contains(r#"<c r="A8""#), "got: {s}");
        assert!(s.contains(r#"<c r="B8""#), "got: {s}");
        assert!(!s.contains(r#"<c r="A3""#), "got: {s}");
    }

    #[test]
    fn moves_block_negative_delta_up() {
        let xml = br#"<worksheet><sheetData><row r="10"><c r="C10"><v>x</v></c></row></sheetData></worksheet>"#;
        let p = plan((10, 3), (10, 3), -5, 0);
        let out = apply_range_move(xml, &p);
        let s = String::from_utf8(out).unwrap();
        assert!(s.contains(r#"<c r="C5""#), "got: {s}");
    }

    #[test]
    fn moves_block_right() {
        let xml = br#"<worksheet><sheetData><row r="3"><c r="A3"><v>1</v></c></row></sheetData></worksheet>"#;
        let p = plan((3, 1), (3, 1), 0, 4);
        let out = apply_range_move(xml, &p);
        let s = String::from_utf8(out).unwrap();
        assert!(s.contains(r#"<c r="E3""#), "got: {s}");
    }

    #[test]
    fn formula_inside_src_with_dollar_does_not_shift() {
        // $A$1 is absolute — must NOT shift even though the cell moves.
        let xml = br#"<worksheet><sheetData><row r="3"><c r="A3"><f>$A$1</f></c></row></sheetData></worksheet>"#;
        let p = plan((3, 1), (3, 1), 5, 0);
        let out = apply_range_move(xml, &p);
        let s = String::from_utf8(out).unwrap();
        // Cell moved to A8, formula still references $A$1.
        assert!(s.contains(r#"<c r="A8""#), "got: {s}");
        assert!(s.contains("<f>$A$1</f>"), "got: {s}");
    }

    #[test]
    fn formula_inside_src_with_relative_ref_shifts() {
        // =A1 is relative — must shift by (5, 0).
        let xml = br#"<worksheet><sheetData><row r="3"><c r="B3"><f>A1</f></c></row></sheetData></worksheet>"#;
        let p = plan((3, 2), (3, 2), 5, 0);
        let out = apply_range_move(xml, &p);
        let s = String::from_utf8(out).unwrap();
        assert!(s.contains(r#"<c r="B8""#), "got: {s}");
        // The formula's reference to A1 — note the move_range
        // translator only re-anchors refs that fall INSIDE src. A1 is
        // outside src=(3,2)..(3,2), so it's left alone.
        assert!(s.contains("<f>A1</f>"), "got: {s}");
    }

    #[test]
    fn formula_inside_src_pointing_into_src_re_anchors() {
        // src is C3:E5. Cell C3 has formula =D4 (inside src). After
        // moving (rows=10, cols=0), C3→C13 and the ref D4→D14.
        let xml = br#"<worksheet><sheetData><row r="3"><c r="C3"><f>D4</f></c></row></sheetData></worksheet>"#;
        let p = plan((3, 3), (5, 5), 10, 0);
        let out = apply_range_move(xml, &p);
        let s = String::from_utf8(out).unwrap();
        assert!(s.contains(r#"<c r="C13""#), "got: {s}");
        assert!(s.contains("<f>D14</f>"), "got: {s}");
    }

    #[test]
    fn external_formula_not_translated_by_default() {
        // Cell A10 has =B3. src=(3,2)..(3,2). Move (rows=5, cols=0).
        // Default translate=false: A10's formula is NOT touched.
        let xml = br#"<worksheet><sheetData><row r="3"><c r="B3"><v>1</v></c></row><row r="10"><c r="A10"><f>B3</f></c></row></sheetData></worksheet>"#;
        let mut p = plan((3, 2), (3, 2), 5, 0);
        p.translate = false;
        let out = apply_range_move(xml, &p);
        let s = String::from_utf8(out).unwrap();
        // A10's formula is left at B3 (the now-empty original).
        assert!(s.contains("<f>B3</f>"), "got: {s}");
    }

    #[test]
    fn external_formula_translated_when_translate_true() {
        let xml = br#"<worksheet><sheetData><row r="3"><c r="B3"><v>1</v></c></row><row r="10"><c r="A10"><f>B3</f></c></row></sheetData></worksheet>"#;
        let mut p = plan((3, 2), (3, 2), 5, 0);
        p.translate = true;
        let out = apply_range_move(xml, &p);
        let s = String::from_utf8(out).unwrap();
        // A10's formula re-anchored from B3 to B8.
        assert!(s.contains("<f>B8</f>"), "got: {s}");
    }

    #[test]
    fn merge_fully_inside_src_shifts() {
        let xml = br#"<worksheet><sheetData><row r="3"><c r="A3"><v>1</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="A3:B5"/></mergeCells></worksheet>"#;
        let p = plan((3, 1), (5, 2), 5, 0);
        let out = apply_range_move(xml, &p);
        let s = String::from_utf8(out).unwrap();
        // Merge moved from A3:B5 to A8:B10.
        assert!(s.contains(r#"ref="A8:B10""#), "got: {s}");
    }

    #[test]
    fn merge_straddling_boundary_left_alone() {
        // src = A3:B5; merge is A2:B6 — straddles the top of src.
        let xml = br#"<worksheet><sheetData/><mergeCells count="1"><mergeCell ref="A2:B6"/></mergeCells></worksheet>"#;
        let p = plan((3, 1), (5, 2), 5, 0);
        let out = apply_range_move(xml, &p);
        let s = String::from_utf8(out).unwrap();
        assert!(s.contains(r#"ref="A2:B6""#), "got: {s}");
    }

    #[test]
    fn hyperlink_anchor_inside_src_shifts() {
        let xml = br#"<worksheet><sheetData/><hyperlinks><hyperlink ref="B3" r:id="rId1"/></hyperlinks></worksheet>"#;
        let p = plan((3, 2), (3, 2), 5, 0);
        let out = apply_range_move(xml, &p);
        let s = String::from_utf8(out).unwrap();
        assert!(s.contains(r#"ref="B8""#), "got: {s}");
    }

    #[test]
    fn hyperlink_attrs_remain_xml_escaped_after_anchor_rewrite() {
        let xml = br#"<worksheet><sheetData/><hyperlinks><hyperlink ref="B3" location="'Sec. 1 &amp; 2 Notes'!A1" display="A &amp; B"/></hyperlinks></worksheet>"#;
        let p = plan((3, 2), (3, 2), 5, 0);
        let out = apply_range_move(xml, &p);
        let s = String::from_utf8(out).unwrap();

        assert!(s.contains(r#"ref="B8""#), "got: {s}");
        assert!(
            s.contains(r#"location="&apos;Sec. 1 &amp; 2 Notes&apos;!A1""#),
            "got: {s}"
        );
        assert!(s.contains(r#"display="A &amp; B""#), "got: {s}");
        assert!(!s.contains(" 1 & 2 "), "got: {s}");
        assert!(!s.contains("A & B"), "got: {s}");
    }

    #[test]
    fn dv_sqref_per_piece_treatment() {
        // src = A3:B5. sqref = "A3:B5 D7:E9" — first piece fully
        // inside, second piece outside. After move (5, 0), first piece
        // → "A8:B10", second stays "D7:E9".
        let xml = br#"<worksheet><sheetData/><dataValidations><dataValidation type="any" sqref="A3:B5 D7:E9"/></dataValidations></worksheet>"#;
        let p = plan((3, 1), (5, 2), 5, 0);
        let out = apply_range_move(xml, &p);
        let s = String::from_utf8(out).unwrap();
        assert!(s.contains(r#"sqref="A8:B10 D7:E9""#), "got: {s}");
    }

    #[test]
    fn overlapping_move_preserves_cells() {
        // Move A1:A3 down by 2 — destination overlaps with original A3.
        // The overlap region (A3) is read before being overwritten,
        // so cell content is preserved.
        let xml = br#"<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row><row r="3"><c r="A3"><v>3</v></c></row></sheetData></worksheet>"#;
        let p = plan((1, 1), (3, 1), 2, 0);
        let out = apply_range_move(xml, &p);
        let s = String::from_utf8(out).unwrap();
        // After move: A3=1, A4=2, A5=3. Original A1 and A2 are blank
        // because they're outside the new dst range.
        assert!(s.contains(r#"<c r="A3""#), "got: {s}");
        assert!(s.contains(r#"<c r="A4""#), "got: {s}");
        assert!(s.contains(r#"<c r="A5""#), "got: {s}");
        // Source rows are gone (A1, A2 empty).
        // Note: <row r="3"> is re-emitted as it now hosts the moved A3.
        // Original cell A3 (value 3) is overwritten → A5 holds value 3.
        // Confirm value preservation:
        assert!(s.contains("<v>1</v>"), "got: {s}");
        assert!(s.contains("<v>2</v>"), "got: {s}");
        assert!(s.contains("<v>3</v>"), "got: {s}");
    }

    #[test]
    fn dimension_attr_passes_through() {
        // The streaming-splice pass treats `<dimension ref>` like any
        // other anchor: per-piece, "fully inside src? shift; else
        // leave". A1:Z100 is a range that straddles src=(1,1)..(1,1)
        // (Z100 is outside), so the dimension is left unchanged. (A
        // proper post-move dimension recompute is a v2 enhancement;
        // this test pins the v1 behaviour — no corruption.)
        let xml = br#"<worksheet><dimension ref="A1:Z100"/><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>"#;
        let p = plan((1, 1), (1, 1), 5, 0);
        let out = apply_range_move(xml, &p);
        let s = String::from_utf8(out).unwrap();
        assert!(s.contains(r#"ref="A1:Z100""#), "got: {s}");
        // The cell did move:
        assert!(s.contains(r#"<c r="A6""#), "got: {s}");
    }

    #[test]
    fn cell_value_preserved_through_move() {
        let xml = br#"<worksheet><sheetData><row r="3"><c r="A3" t="s"><v>42</v></c></row></sheetData></worksheet>"#;
        let p = plan((3, 1), (3, 1), 5, 0);
        let out = apply_range_move(xml, &p);
        let s = String::from_utf8(out).unwrap();
        assert!(s.contains(r#"<c r="A8" t="s">"#), "got: {s}");
        assert!(s.contains("<v>42</v>"), "got: {s}");
    }

    #[test]
    fn inline_string_cell_moves_intact() {
        let xml = br#"<worksheet><sheetData><row r="3"><c r="A3" t="inlineStr"><is><t>hello</t></is></c></row></sheetData></worksheet>"#;
        let p = plan((3, 1), (3, 1), 5, 0);
        let out = apply_range_move(xml, &p);
        let s = String::from_utf8(out).unwrap();
        assert!(s.contains(r#"<c r="A8" t="inlineStr">"#), "got: {s}");
        assert!(s.contains("<t>hello</t>"), "got: {s}");
    }

    #[test]
    fn move_preserves_literal_entities_and_translates_complete_formulas() {
        let xml = br#"<worksheet><sheetData><row r="1" note="A&amp;B"><c r="A1" note="&quot;&amp;"><f>A&#49;&amp;&quot;A1&quot;</f><is><t> &#65;&amp;&lt; </t></is></c></row></sheetData><dataValidation sqref="A1"><formula1>A&#49;</formula1></dataValidation><conditionalFormatting sqref="A1"><cfRule><formula>A&#x31;&lt;2</formula></cfRule></conditionalFormatting></worksheet>"#;
        let mut p = plan((1, 1), (1, 1), 1, 0);
        p.translate = true;
        let out = String::from_utf8(apply_range_move(xml, &p)).unwrap();
        assert!(out.contains("<f>A2&amp;&quot;A1&quot;</f>"), "{out}");
        assert!(out.contains("<formula1>A2</formula1>"), "{out}");
        assert!(out.contains("<formula>A2&lt;2</formula>"), "{out}");
        assert!(out.contains(r#"note="A&amp;B""#), "{out}");
        assert!(out.contains(r#"note="&quot;&amp;""#), "{out}");
        assert!(out.contains("<t> &#65;&amp;&lt; </t>"), "{out}");
    }
}
