//! Comment + VML drawing extraction / merge / emission for the modify-mode
//! patcher (RFC-023).
//!
//! Used by `XlsxPatcher::do_save`'s Phase 2.5g to:
//! 1. Read existing `<authors>` and `<commentList>` from `commentsN.xml`
//!    parts referenced from the sheet rels.
//! 2. Read existing `<v:shape>` elements from `vmlDrawingN.vml` parts,
//!    preserving non-comment shapes (form controls, images) verbatim.
//! 3. Merge user-supplied `CommentOp::Set` / `CommentOp::Delete` on top.
//! 4. Re-emit fresh `commentsN.xml` and `vmlDrawingN.vml` byte streams.
//! 5. Mutate the sheet's `RelsGraph` for added/removed comments / vml
//!    rels.
//! 6. Inject a `<legacyDrawing r:id="…"/>` block into the sheet xml at
//!    ECMA §18.3.1.99 slot 31 via `wolfxl_merger::SheetBlock`.
//! 7. Add `<Override>` content-type entries (and the `vml` `<Default>`)
//!    via `content_types::ContentTypeOp`.
//!
//! ## VML pixel-positioning seam vs the native writer
//!
//! The native writer's `compute_margin` (drawings_vml.rs:38-43) hard-codes
//! `COL_WIDTH_PT = 48.0`. That is correct only on sheets whose `<cols>`
//! block declares no overrides. The patcher fixes this by parsing the
//! sheet xml's `<cols>` block and computing per-column widths in points.
//! See [`compute_margin_with_widths`] for the math. The native writer
//! retains the bug because fixing it requires threading `Worksheet.cols`
//! through to the emitter — INDEX decision #8 punts that to post-1.0.

use std::collections::BTreeMap;

use indexmap::IndexMap;
use quick_xml::events::Event;
use quick_xml::Reader as XmlReader;

use wolfxl_rels::{rt, RelId, RelsGraph, TargetMode};

// ---------------------------------------------------------------------------
// Public types
// ---------------------------------------------------------------------------

/// One user-supplied comment edit.
#[derive(Debug, Clone, PartialEq)]
pub struct CommentPatch {
    pub coordinate: String,
    pub author: String,
    pub text: String,
    pub width_pt: Option<f64>,
    pub height_pt: Option<f64>,
}

/// `Set` adds-or-overwrites the comment at `coordinate`; `Delete` removes
/// it. None sentinel pattern matches RFC-022 hyperlinks.
#[derive(Debug, Clone, PartialEq)]
pub enum CommentOp {
    Set(CommentPatch),
    Delete,
}

/// One comment already present in the source `commentsN.xml`, with author
/// resolved to a name (so the patcher's workbook-scope author table can
/// re-intern by name and assign a new index).
#[derive(Debug, Clone, PartialEq)]
pub struct ExistingComment {
    pub coordinate: String,
    pub author: String,
    /// Inner XML of `<text>…</text>` carried verbatim (preserves
    /// rich-text runs).
    pub text_inner_xml: String,
    /// `<commentPr>` subtree if present, raw bytes including open + close.
    pub comment_pr: Option<String>,
    /// Optional `<extLst>` subtree (Excel 365 threaded-comment back-ref).
    pub ext_lst: Option<String>,
    pub width_pt: Option<f64>,
    pub height_pt: Option<f64>,
}

/// Workbook-scope author registry. Insertion-ordered (IndexMap) — fixes
/// the upstream BTreeMap bug. Authors deduped by exact string
/// equality (RFC-023 §8 risk #2).
#[derive(Debug, Clone, Default)]
pub struct CommentAuthorTable {
    inner: IndexMap<String, u32>,
}

impl CommentAuthorTable {
    pub fn new() -> Self {
        Self::default()
    }

    /// Intern an author name. Returns the stable, workbook-scope authorId.
    pub fn intern(&mut self, name: &str) -> u32 {
        if let Some(&id) = self.inner.get(name) {
            return id;
        }
        let id = self.inner.len() as u32;
        self.inner.insert(name.to_string(), id);
        id
    }

    #[allow(dead_code)]
    pub fn name_of(&self, id: u32) -> Option<&str> {
        self.inner
            .iter()
            .find_map(|(k, v)| if *v == id { Some(k.as_str()) } else { None })
    }

    pub fn iter(&self) -> impl Iterator<Item = (&str, u32)> {
        self.inner.iter().map(|(k, v)| (k.as_str(), *v))
    }

    #[allow(dead_code)]
    pub fn len(&self) -> usize {
        self.inner.len()
    }

    #[allow(dead_code)]
    pub fn is_empty(&self) -> bool {
        self.inner.is_empty()
    }
}

/// Output of [`build_comments`].
#[derive(Debug, Clone, Default, PartialEq)]
pub struct CommentsResult {
    pub comments_xml: Vec<u8>,
    pub vml_drawing: Vec<u8>,
    pub legacy_drawing_rid: Option<RelId>,
}

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

pub const CT_COMMENTS: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";
pub const CT_VML: &str = "application/vnd.openxmlformats-officedocument.vmlDrawing";

const ORIGIN_LEFT_PT: f64 = 59.25;
const ORIGIN_TOP_PT: f64 = 1.5;
const DEFAULT_COL_WIDTH_PT: f64 = 48.0;
const DEFAULT_ROW_HEIGHT_PT: f64 = 12.75;

// ---------------------------------------------------------------------------
// Extract: comments
// ---------------------------------------------------------------------------

/// Parse a `commentsN.xml` byte stream into `(authors, comments)`.
pub fn extract_comments(xml: &[u8]) -> (Vec<String>, BTreeMap<String, ExistingComment>) {
    let mut authors: Vec<String> = Vec::new();
    let mut comments: BTreeMap<String, ExistingComment> = BTreeMap::new();
    if xml.is_empty() {
        return (authors, comments);
    }

    let text = match std::str::from_utf8(xml) {
        Ok(s) => s,
        Err(_) => return (authors, comments),
    };

    // Authors: stream-scan inside <authors>...</authors>.
    let mut reader = XmlReader::from_str(text);
    reader.config_mut().trim_text(false);
    let mut buf: Vec<u8> = Vec::new();
    let mut in_authors = false;
    let mut in_author = false;
    let mut current_author = String::new();
    let mut node = crate::ooxml_util::XmlTextNode::default();
    loop {
        let event = reader.read_event_into(&mut buf);
        if !matches!(event, Ok(Event::Text(_)) | Ok(Event::GeneralRef(_))) {
            if in_author {
                if let Ok(text) = quick_xml::escape::unescape(node.raw()) {
                    current_author.push_str(&text);
                }
            }
            node.clear();
        }
        match event {
            Ok(Event::Start(e)) => {
                let n = e.local_name();
                let name = n.as_ref().as_bytes();
                if name == b"authors" {
                    in_authors = true;
                } else if in_authors && name == b"author" {
                    in_author = true;
                    current_author.clear();
                }
            }
            Ok(Event::End(e)) => {
                let n = e.local_name();
                let name = n.as_ref().as_bytes();
                if name == b"author" && in_author {
                    authors.push(std::mem::take(&mut current_author));
                    in_author = false;
                } else if name == b"authors" {
                    break;
                }
            }
            Ok(Event::Text(t)) if in_author => node.push_text(&t),
            Ok(Event::GeneralRef(r)) if in_author => node.push_ref(&r),
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    // Comments: walk top-level <comment> children of <commentList>.
    let comment_list = extract_block_inner(text, "commentList").unwrap_or_default();
    if comment_list.is_empty() {
        return (authors, comments);
    }
    for raw in iter_top_level_elements(&comment_list, "comment") {
        let (attrs, inner) = split_open_inner(&raw, "comment");
        let mut coord = String::new();
        let mut author_id: u32 = 0;
        for (k, v) in attrs {
            match k.as_str() {
                "ref" => coord = v,
                "authorId" => author_id = v.parse().unwrap_or(0),
                _ => {}
            }
        }
        if coord.is_empty() {
            continue;
        }
        let author = authors.get(author_id as usize).cloned().unwrap_or_default();
        let text_inner = extract_block_inner(&inner, "text").unwrap_or_default();
        let comment_pr = extract_block_full(&inner, "commentPr");
        let ext_lst = extract_block_full(&inner, "extLst");
        comments.insert(
            coord.clone(),
            ExistingComment {
                coordinate: coord,
                author,
                text_inner_xml: text_inner,
                comment_pr,
                ext_lst,
                width_pt: None,
                height_pt: None,
            },
        );
    }
    (authors, comments)
}

// ---------------------------------------------------------------------------
// Extract: VML drawing
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PreservedVmlShape {
    pub raw_xml: String,
}

#[derive(Debug, Clone, Default)]
pub struct ExtractedVml {
    pub preserved: Vec<PreservedVmlShape>,
    pub sizes: BTreeMap<String, (Option<f64>, Option<f64>)>,
    pub idmap_data: Option<String>,
}

pub fn extract_vml(xml: &[u8]) -> ExtractedVml {
    let mut out = ExtractedVml::default();
    if xml.is_empty() {
        return out;
    }
    let text = match std::str::from_utf8(xml) {
        Ok(s) => s,
        Err(_) => return out,
    };
    let inner = match extract_block_inner(text, "xml") {
        Some(s) => s,
        None => return out,
    };

    if let Some(idmap) = find_self_closing_attr(&inner, "idmap", "data") {
        out.idmap_data = Some(idmap);
    }

    for raw in iter_top_level_elements(&inner, "v:shape") {
        let (attrs, body) = split_open_inner(&raw, "v:shape");
        let mut shape_type = String::new();
        let mut style = String::new();
        for (k, v) in &attrs {
            if k == "type" {
                shape_type = v.clone();
            } else if k == "style" {
                style = v.clone();
            }
        }
        if shape_type == "#_x0000_t202" {
            let (w, h) = parse_size_from_style(&style);
            if let Some(coord) = parse_anchor_cell_from_body(&body) {
                out.sizes.insert(coord, (w, h));
            }
        } else {
            out.preserved.push(PreservedVmlShape { raw_xml: raw });
        }
    }
    out
}

fn parse_size_from_style(style: &str) -> (Option<f64>, Option<f64>) {
    let mut w: Option<f64> = None;
    let mut h: Option<f64> = None;
    for part in style.split(';') {
        let part = part.trim();
        if let Some(rest) = part.strip_prefix("width:") {
            w = parse_pt(rest.trim());
        } else if let Some(rest) = part.strip_prefix("height:") {
            h = parse_pt(rest.trim());
        }
    }
    (w, h)
}

fn parse_pt(s: &str) -> Option<f64> {
    let s = s.trim();
    let stripped = s.strip_suffix("pt").unwrap_or(s);
    stripped.trim().parse::<f64>().ok()
}

fn parse_anchor_cell_from_body(body: &str) -> Option<String> {
    let row = extract_block_inner(body, "x:Row")?
        .trim()
        .parse::<u32>()
        .ok()?;
    let col = extract_block_inner(body, "x:Column")?
        .trim()
        .parse::<u32>()
        .ok()?;
    Some(rowcol0_to_a1(row, col))
}

fn rowcol0_to_a1(row0: u32, col0: u32) -> String {
    let mut col_letters = String::new();
    let mut n = col0 + 1;
    while n > 0 {
        let r = (n - 1) % 26;
        col_letters.insert(0, (b'A' + r as u8) as char);
        n = (n - 1) / 26;
    }
    format!("{col_letters}{}", row0 + 1)
}

// ---------------------------------------------------------------------------
// Cols parser — for VML pixel positioning math
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, Default)]
pub struct ColWidthMap {
    inner: BTreeMap<u32, f64>,
}

impl ColWidthMap {
    #[allow(dead_code)]
    pub fn new() -> Self {
        Self::default()
    }

    pub fn get(&self, col0: u32) -> Option<f64> {
        self.inner.get(&col0).copied()
    }
}

fn col_units_to_pt(units: f64) -> f64 {
    let px = ((units * 7.0 + 5.0) / 7.0 * 7.0 + 5.0).trunc();
    px * 72.0 / 96.0
}

pub fn parse_col_widths(sheet_xml: &[u8]) -> ColWidthMap {
    let mut map = ColWidthMap::default();
    let text = match std::str::from_utf8(sheet_xml) {
        Ok(s) => s,
        Err(_) => return map,
    };
    let cols_inner = match extract_block_inner(text, "cols") {
        Some(s) => s,
        None => return map,
    };
    let mut reader = XmlReader::from_str(&cols_inner);
    let mut buf: Vec<u8> = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) | Ok(Event::Start(e)) => {
                if e.local_name().as_ref() != "col" {
                    buf.clear();
                    continue;
                }
                let mut min_v: Option<u32> = None;
                let mut max_v: Option<u32> = None;
                let mut width_v: Option<f64> = None;
                for a in e.attributes().with_checks(false).flatten() {
                    let key = a.key.as_ref().as_bytes();
                    let val = quick_xml::escape::unescape(&a.value)
                        .map(|v| v.into_owned())
                        .unwrap_or_else(|_| a.value.to_string());
                    match key {
                        b"min" => min_v = val.parse().ok(),
                        b"max" => max_v = val.parse().ok(),
                        b"width" => width_v = val.parse().ok(),
                        _ => {}
                    }
                }
                if let (Some(min), Some(max), Some(w)) = (min_v, max_v, width_v) {
                    let pt = col_units_to_pt(w);
                    let lo = min.saturating_sub(1);
                    let hi = max.saturating_sub(1);
                    for c in lo..=hi {
                        map.inner.insert(c, pt);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    map
}

pub fn compute_margin_with_widths(row0: u32, col0: u32, cols: &ColWidthMap) -> (f64, f64) {
    if col0 == 0 && cols.inner.is_empty() {
        return (
            ORIGIN_LEFT_PT,
            (row0 as f64) * DEFAULT_ROW_HEIGHT_PT + ORIGIN_TOP_PT,
        );
    }
    if cols.inner.is_empty() {
        let margin_left = (col0 as f64) * DEFAULT_COL_WIDTH_PT + ORIGIN_LEFT_PT;
        let margin_top = (row0 as f64) * DEFAULT_ROW_HEIGHT_PT + ORIGIN_TOP_PT;
        return (margin_left, margin_top);
    }
    let mut margin_left = ORIGIN_LEFT_PT;
    for c in 0..col0 {
        margin_left += cols.get(c).unwrap_or(DEFAULT_COL_WIDTH_PT);
    }
    let margin_top = (row0 as f64) * DEFAULT_ROW_HEIGHT_PT + ORIGIN_TOP_PT;
    (margin_left, margin_top)
}

// ---------------------------------------------------------------------------
// Build: comments.xml
// ---------------------------------------------------------------------------

#[derive(Debug, Clone)]
pub struct MergedComment {
    pub coordinate: String,
    pub author_id: u32,
    pub text_inner_xml: String,
    pub comment_pr: Option<String>,
    pub ext_lst: Option<String>,
    pub width_pt: Option<f64>,
    pub height_pt: Option<f64>,
}

pub fn merge_comments(
    existing: BTreeMap<String, ExistingComment>,
    ops: &BTreeMap<String, CommentOp>,
    authors: &mut CommentAuthorTable,
    vml_sizes: &BTreeMap<String, (Option<f64>, Option<f64>)>,
) -> BTreeMap<String, MergedComment> {
    let mut merged: BTreeMap<String, MergedComment> = BTreeMap::new();
    for (coord, ec) in existing {
        let id = authors.intern(&ec.author);
        let (w, h) = vml_sizes
            .get(&coord)
            .copied()
            .unwrap_or((ec.width_pt, ec.height_pt));
        merged.insert(
            coord.clone(),
            MergedComment {
                coordinate: coord,
                author_id: id,
                text_inner_xml: ec.text_inner_xml,
                comment_pr: ec.comment_pr,
                ext_lst: ec.ext_lst,
                width_pt: w,
                height_pt: h,
            },
        );
    }
    for (coord, op) in ops {
        match op {
            CommentOp::Set(p) => {
                let id = authors.intern(&p.author);
                let plain_text = format!("<t>{}</t>", xml_escape_text(&p.text));
                merged.insert(
                    coord.clone(),
                    MergedComment {
                        coordinate: coord.clone(),
                        author_id: id,
                        text_inner_xml: plain_text,
                        comment_pr: None,
                        ext_lst: None,
                        width_pt: p.width_pt,
                        height_pt: p.height_pt,
                    },
                );
            }
            CommentOp::Delete => {
                merged.remove(coord);
            }
        }
    }
    merged
}

pub fn build_comments_xml(
    merged: &BTreeMap<String, MergedComment>,
    authors: &CommentAuthorTable,
) -> Vec<u8> {
    if merged.is_empty() {
        return Vec::new();
    }
    let mut out = String::with_capacity(2048 + merged.len() * 96);
    out.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n");
    out.push_str("<comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
    out.push_str("<authors>");
    for (name, _id) in authors.iter() {
        out.push_str("<author>");
        push_xml_text_escape(&mut out, name);
        out.push_str("</author>");
    }
    out.push_str("</authors>");
    out.push_str("<commentList>");
    for (_coord, c) in merged {
        out.push_str("<comment ref=\"");
        push_xml_attr_escape(&mut out, &c.coordinate);
        out.push_str("\" authorId=\"");
        out.push_str(&c.author_id.to_string());
        out.push_str("\">");
        out.push_str("<text>");
        out.push_str(&c.text_inner_xml);
        out.push_str("</text>");
        if let Some(pr) = &c.comment_pr {
            out.push_str(pr);
        }
        if let Some(ext) = &c.ext_lst {
            out.push_str(ext);
        }
        out.push_str("</comment>");
    }
    out.push_str("</commentList>");
    out.push_str("</comments>");
    out.into_bytes()
}

// ---------------------------------------------------------------------------
// Build: vmlDrawing.vml
// ---------------------------------------------------------------------------

pub fn build_vml(
    merged: &BTreeMap<String, MergedComment>,
    preserved: &[PreservedVmlShape],
    cols: &ColWidthMap,
    idmap_data: Option<&str>,
) -> Vec<u8> {
    if merged.is_empty() && preserved.is_empty() {
        return Vec::new();
    }
    let mut out = String::with_capacity(4096);
    out.push_str(
        "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" \
xmlns:o=\"urn:schemas-microsoft-com:office:office\" \
xmlns:x=\"urn:schemas-microsoft-com:office:excel\">",
    );
    let idmap = idmap_data.unwrap_or("1");
    out.push_str(&format!(
        "<o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" data=\"{}\"/></o:shapelayout>",
        idmap
    ));
    out.push_str(
        "<v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\" \
path=\"m,l,21600r21600,l21600,xe\">\
<v:stroke joinstyle=\"miter\"/>\
<v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/>\
</v:shapetype>",
    );
    for s in preserved {
        out.push_str(&s.raw_xml);
    }
    for (idx, (cell_ref, comment)) in merged.iter().enumerate() {
        let shape_num = 1025u32 + idx as u32;
        let (row0, col0) = match a1_to_row_col0(cell_ref) {
            Some(rc) => rc,
            None => continue,
        };
        let (margin_left, margin_top) = compute_margin_with_widths(row0, col0, cols);
        let width = match comment.width_pt {
            Some(w) => format!("{}pt", format_f64(w)),
            None => "96pt".to_string(),
        };
        let height = match comment.height_pt {
            Some(h) => format!("{}pt", format_f64(h)),
            None => "55.5pt".to_string(),
        };
        out.push_str(&format!(
            "<v:shape id=\"_x0000_s{}\" type=\"#_x0000_t202\" \
style=\"position:absolute; margin-left:{}pt; margin-top:{}pt; \
width:{}; height:{}; z-index:1; visibility:hidden\" \
fillcolor=\"#ffffe1\" o:insetmode=\"auto\">",
            shape_num,
            format_f64(margin_left),
            format_f64(margin_top),
            width,
            height,
        ));
        out.push_str("<v:fill color2=\"#ffffe1\"/>");
        out.push_str("<v:shadow on=\"t\" color=\"black\" obscured=\"t\"/>");
        out.push_str("<v:path o:connecttype=\"none\"/>");
        out.push_str(
            "<v:textbox style=\"mso-direction-alt:auto\"><div style=\"text-align:left\"/></v:textbox>",
        );
        out.push_str("<x:ClientData ObjectType=\"Note\">");
        out.push_str("<x:MoveWithCells/>");
        out.push_str("<x:SizeWithCells/>");
        let (cl, ol, rt, ot, cr, or_, rb, ob) = compute_anchor(row0, col0);
        out.push_str(&format!(
            "<x:Anchor>{}, {}, {}, {}, {}, {}, {}, {}</x:Anchor>",
            cl, ol, rt, ot, cr, or_, rb, ob
        ));
        out.push_str("<x:AutoFill>False</x:AutoFill>");
        out.push_str(&format!("<x:Row>{}</x:Row>", row0));
        out.push_str(&format!("<x:Column>{}</x:Column>", col0));
        out.push_str("</x:ClientData>");
        out.push_str("</v:shape>");
    }
    out.push_str("</xml>");
    out.into_bytes()
}

fn compute_anchor(row0: u32, col0: u32) -> (u32, u32, u32, u32, u32, u32, u32, u32) {
    let col_left = col0 + 1;
    let row_top = row0.saturating_sub(1);
    let col_right = col0 + 3;
    let row_bottom = row0 + 3;
    (col_left, 15, row_top, 10, col_right, 15, row_bottom, 4)
}

fn format_f64(v: f64) -> String {
    if v.fract() == 0.0 {
        format!("{}", v as i64)
    } else {
        let s = format!("{:.6}", v);
        let trimmed = s.trim_end_matches('0').trim_end_matches('.').to_string();
        if trimmed.is_empty() {
            "0".to_string()
        } else {
            trimmed
        }
    }
}

// ---------------------------------------------------------------------------
// Top-level drive
// ---------------------------------------------------------------------------

#[allow(clippy::too_many_arguments)]
pub fn build_comments(
    existing_comments_xml: Option<&[u8]>,
    existing_vml: Option<&[u8]>,
    ops: &BTreeMap<String, CommentOp>,
    sheet_xml: &[u8],
    rels: &mut RelsGraph,
    authors: &mut CommentAuthorTable,
    comments_n: u32,
    vml_n: u32,
) -> (CommentsResult, Option<RelId>, Option<RelId>) {
    let (_existing_authors, existing_comments) = match existing_comments_xml {
        Some(xml) => extract_comments(xml),
        None => (Vec::new(), BTreeMap::new()),
    };
    let extracted_vml = match existing_vml {
        Some(xml) => extract_vml(xml),
        None => ExtractedVml::default(),
    };

    for (_, ec) in &existing_comments {
        authors.intern(&ec.author);
    }

    let merged = merge_comments(existing_comments, ops, authors, &extracted_vml.sizes);

    let comments_target_relative = format!("../comments{}.xml", comments_n);
    let vml_target_relative = format!("../drawings/vmlDrawing{}.vml", vml_n);
    let comments_rid_existing = rels
        .find_by_type(rt::COMMENTS)
        .first()
        .map(|r| r.id.clone());
    let vml_rid_existing = rels
        .find_by_type(rt::VML_DRAWING)
        .first()
        .map(|r| r.id.clone());

    if merged.is_empty() {
        if let Some(rid) = &comments_rid_existing {
            rels.remove(rid);
        }
        if extracted_vml.preserved.is_empty() {
            if let Some(rid) = &vml_rid_existing {
                rels.remove(rid);
            }
        }
        let result = CommentsResult {
            comments_xml: Vec::new(),
            vml_drawing: if extracted_vml.preserved.is_empty() {
                Vec::new()
            } else {
                let cols = parse_col_widths(sheet_xml);
                build_vml(
                    &BTreeMap::new(),
                    &extracted_vml.preserved,
                    &cols,
                    extracted_vml.idmap_data.as_deref(),
                )
            },
            legacy_drawing_rid: if extracted_vml.preserved.is_empty() {
                None
            } else {
                vml_rid_existing.clone()
            },
        };
        return (result, comments_rid_existing, vml_rid_existing);
    }

    let comments_rid = match comments_rid_existing.clone() {
        Some(r) => r,
        None => rels.add(
            rt::COMMENTS,
            &comments_target_relative,
            TargetMode::Internal,
        ),
    };
    let vml_rid = match vml_rid_existing.clone() {
        Some(r) => r,
        None => rels.add(rt::VML_DRAWING, &vml_target_relative, TargetMode::Internal),
    };

    let cols = parse_col_widths(sheet_xml);
    let comments_xml = build_comments_xml(&merged, authors);
    let vml_xml = build_vml(
        &merged,
        &extracted_vml.preserved,
        &cols,
        extracted_vml.idmap_data.as_deref(),
    );
    let result = CommentsResult {
        comments_xml,
        vml_drawing: vml_xml,
        legacy_drawing_rid: Some(vml_rid.clone()),
    };
    (result, Some(comments_rid), Some(vml_rid))
}

// ---------------------------------------------------------------------------
// Helpers — string-level XML probing
// ---------------------------------------------------------------------------

fn extract_block_inner(xml: &str, name: &str) -> Option<String> {
    let open_marker = format!("<{}", name);
    let start = xml.find(&open_marker)?;
    let after_open = xml[start..].find('>')? + start + 1;
    if xml[start..after_open].ends_with("/>") {
        return Some(String::new());
    }
    let close_marker = format!("</{}>", name);
    let close = xml[after_open..].find(&close_marker)? + after_open;
    Some(xml[after_open..close].to_string())
}

fn extract_block_full(xml: &str, name: &str) -> Option<String> {
    let open_marker = format!("<{}", name);
    let start = xml.find(&open_marker)?;
    let after_open = xml[start..].find('>')? + start + 1;
    if xml[start..after_open].ends_with("/>") {
        return Some(xml[start..after_open].to_string());
    }
    let close_marker = format!("</{}>", name);
    let close_end = xml[after_open..].find(&close_marker)? + after_open + close_marker.len();
    Some(xml[start..close_end].to_string())
}

fn iter_top_level_elements(xml: &str, name: &str) -> Vec<String> {
    let mut out = Vec::new();
    let open_prefix = format!("<{}", name);
    let close_marker = format!("</{}>", name);
    let mut cursor = 0usize;
    while let Some(rel_start) = xml[cursor..].find(&open_prefix) {
        let start = cursor + rel_start;
        let next_byte = xml.as_bytes().get(start + open_prefix.len()).copied();
        if !matches!(
            next_byte,
            Some(b' ') | Some(b'\t') | Some(b'\n') | Some(b'/') | Some(b'>')
        ) {
            cursor = start + open_prefix.len();
            continue;
        }
        let after_open_rel = match xml[start..].find('>') {
            Some(p) => p,
            None => break,
        };
        let after_open = start + after_open_rel + 1;
        if xml[start..after_open].ends_with("/>") {
            out.push(xml[start..after_open].to_string());
            cursor = after_open;
            continue;
        }
        let close_rel = match xml[after_open..].find(&close_marker) {
            Some(p) => p,
            None => break,
        };
        let close_end = after_open + close_rel + close_marker.len();
        out.push(xml[start..close_end].to_string());
        cursor = close_end;
    }
    out
}

fn split_open_inner(raw: &str, name: &str) -> (Vec<(String, String)>, String) {
    let open_prefix = format!("<{}", name);
    let close_marker = format!("</{}>", name);
    let mut attrs: Vec<(String, String)> = Vec::new();
    let after_open = match raw.find('>') {
        Some(p) => p + 1,
        None => return (attrs, String::new()),
    };
    let attr_text = &raw[open_prefix.len()..after_open - 1];
    let attr_text = attr_text.trim_end_matches('/').trim();
    parse_attrs(attr_text, &mut attrs);
    let body = if raw[..after_open].ends_with("/>") {
        String::new()
    } else if let Some(close_rel) = raw[after_open..].rfind(&close_marker) {
        raw[after_open..after_open + close_rel].to_string()
    } else {
        String::new()
    };
    (attrs, body)
}

fn parse_attrs(s: &str, out: &mut Vec<(String, String)>) {
    let mut chars = s.char_indices().peekable();
    while let Some(&(_, c)) = chars.peek() {
        if c.is_whitespace() {
            chars.next();
            continue;
        }
        let key_start = chars.peek().unwrap().0;
        let mut key_end = key_start;
        while let Some(&(i, ch)) = chars.peek() {
            if ch == '=' || ch.is_whitespace() {
                key_end = i;
                break;
            }
            key_end = i + ch.len_utf8();
            chars.next();
        }
        while let Some(&(_, ch)) = chars.peek() {
            if ch.is_whitespace() || ch == '=' {
                chars.next();
            } else {
                break;
            }
        }
        let quote = match chars.peek() {
            Some(&(_, q)) if q == '"' || q == '\'' => {
                chars.next();
                q
            }
            _ => return,
        };
        let val_start = match chars.peek() {
            Some(&(i, _)) => i,
            None => return,
        };
        let mut val_end = val_start;
        while let Some(&(i, ch)) = chars.peek() {
            if ch == quote {
                val_end = i;
                chars.next();
                break;
            }
            val_end = i + ch.len_utf8();
            chars.next();
        }
        let key = s[key_start..key_end].trim().to_string();
        let val = s[val_start..val_end].to_string();
        let val = xml_unescape(&val);
        out.push((key, val));
    }
}

fn xml_unescape(s: &str) -> String {
    s.replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", "\"")
        .replace("&apos;", "'")
        .replace("&amp;", "&")
}

fn find_self_closing_attr(xml: &str, name: &str, attr: &str) -> Option<String> {
    let mut cursor = 0usize;
    while let Some(rel_start) = xml[cursor..].find('<') {
        let start = cursor + rel_start;
        if !tag_at_has_local_name(xml, start, name) {
            cursor = start + 1;
            continue;
        }
        let after_open_rel = xml[start..].find('>')?;
        let after_open = start + after_open_rel + 1;
        let tag_end = xml[start + 1..after_open - 1]
            .find(|c: char| c.is_whitespace() || c == '/' || c == '>')
            .map(|rel| start + 1 + rel)
            .unwrap_or(after_open - 1);
        let attr_text = &xml[tag_end..after_open - 1];
        let mut attrs: Vec<(String, String)> = Vec::new();
        parse_attrs(attr_text.trim_end_matches('/').trim(), &mut attrs);
        for (k, v) in attrs {
            if k == attr {
                return Some(v);
            }
        }
        cursor = after_open;
    }
    None
}

fn tag_at_has_local_name(xml: &str, start: usize, expected: &str) -> bool {
    if xml.as_bytes().get(start) != Some(&b'<') {
        return false;
    }
    let rest = &xml[start + 1..];
    if rest.starts_with('/') || rest.starts_with('!') || rest.starts_with('?') {
        return false;
    }
    let tag_end = rest
        .find(|c: char| c.is_whitespace() || c == '/' || c == '>')
        .unwrap_or(rest.len());
    let tag_name = &rest[..tag_end];
    tag_name
        .rsplit_once(':')
        .map_or(tag_name, |(_, local)| local)
        == expected
}

fn xml_escape_text(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    push_xml_text_escape(&mut out, s);
    out
}

fn push_xml_text_escape(out: &mut String, s: &str) {
    for ch in s.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            _ => out.push(ch),
        }
    }
}

fn push_xml_attr_escape(out: &mut String, s: &str) {
    for ch in s.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&apos;"),
            _ => out.push(ch),
        }
    }
}

fn a1_to_row_col0(a1: &str) -> Option<(u32, u32)> {
    let mut letters_end = 0usize;
    for (i, ch) in a1.char_indices() {
        if ch.is_ascii_alphabetic() {
            letters_end = i + 1;
        } else {
            break;
        }
    }
    if letters_end == 0 {
        return None;
    }
    let letters = &a1[..letters_end];
    let digits = &a1[letters_end..];
    let row: u32 = digits.parse().ok()?;
    if row == 0 {
        return None;
    }
    let mut col: u32 = 0;
    for ch in letters.chars() {
        if !ch.is_ascii_alphabetic() {
            return None;
        }
        col = col * 26 + (ch.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
    }
    Some((row - 1, col - 1))
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;

    fn op_set(coord: &str, author: &str, text: &str) -> CommentOp {
        CommentOp::Set(CommentPatch {
            coordinate: coord.to_string(),
            author: author.to_string(),
            text: text.to_string(),
            width_pt: None,
            height_pt: None,
        })
    }

    #[test]
    fn author_table_dedups_and_preserves_insertion_order() {
        let mut t = CommentAuthorTable::new();
        assert_eq!(t.intern("Alice"), 0);
        assert_eq!(t.intern("Bob"), 1);
        assert_eq!(t.intern("Alice"), 0);
        assert_eq!(t.intern("Charlie"), 2);
        let names: Vec<&str> = t.iter().map(|(n, _)| n).collect();
        assert_eq!(names, vec!["Alice", "Bob", "Charlie"]);
    }

    #[test]
    fn author_table_case_sensitive() {
        let mut t = CommentAuthorTable::new();
        assert_eq!(t.intern("Alice"), 0);
        assert_eq!(t.intern("Alice "), 1);
        assert_eq!(t.len(), 2);
    }

    #[test]
    fn author_entities_preserve_node_boundaries_and_whitespace() {
        let xml = b"<comments><authors><author> A &amp; B&#10;&#13; </author>\
                    <author>before &unknown; after</author></authors></comments>";
        let (authors, _) = extract_comments(xml);
        assert_eq!(authors, vec![" A & B\n\r ", ""]);
    }

    #[test]
    fn extract_comments_basic() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<authors><author>Alice</author><author>Bob</author></authors>
<commentList>
<comment ref="A1" authorId="0"><text><t>note one</t></text></comment>
<comment ref="B5" authorId="1"><text><t>note two</t></text></comment>
</commentList>
</comments>"#;
        let (authors, comments) = extract_comments(xml.as_bytes());
        assert_eq!(authors, vec!["Alice", "Bob"]);
        assert_eq!(comments.len(), 2);
        assert_eq!(comments["A1"].author, "Alice");
        assert_eq!(comments["B5"].author, "Bob");
        assert!(comments["A1"].text_inner_xml.contains("note one"));
    }

    #[test]
    fn extract_comments_preserves_extlst() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<authors><author>Alice</author></authors>
<commentList>
<comment ref="A1" authorId="0"><text><t>x</t></text><extLst><ext uri="{...}"><x14:guid>{abc}</x14:guid></ext></extLst></comment>
</commentList>
</comments>"#;
        let (_, comments) = extract_comments(xml.as_bytes());
        let c = &comments["A1"];
        assert!(c.ext_lst.is_some());
        let ext = c.ext_lst.as_ref().unwrap();
        assert!(ext.contains("{abc}"), "ext_lst preserves guid: {ext}");
    }

    #[test]
    fn build_comments_xml_round_trip_byte_close() {
        let mut authors = CommentAuthorTable::new();
        authors.intern("Alice");
        let mut merged = BTreeMap::new();
        merged.insert(
            "A1".to_string(),
            MergedComment {
                coordinate: "A1".into(),
                author_id: 0,
                text_inner_xml: "<t>hello</t>".into(),
                comment_pr: None,
                ext_lst: None,
                width_pt: None,
                height_pt: None,
            },
        );
        let bytes = build_comments_xml(&merged, &authors);
        let s = std::str::from_utf8(&bytes).unwrap();
        assert!(s.contains("<author>Alice</author>"), "{s}");
        assert!(s.contains(r#"<comment ref="A1" authorId="0">"#), "{s}");
        assert!(s.contains("<t>hello</t>"));
    }

    #[test]
    fn merge_set_dedups_existing_author() {
        let mut authors = CommentAuthorTable::new();
        let mut existing = BTreeMap::new();
        existing.insert(
            "A1".to_string(),
            ExistingComment {
                coordinate: "A1".into(),
                author: "Alice".into(),
                text_inner_xml: "<t>old</t>".into(),
                comment_pr: None,
                ext_lst: None,
                width_pt: None,
                height_pt: None,
            },
        );
        authors.intern("Alice");
        let mut ops = BTreeMap::new();
        ops.insert("B5".to_string(), op_set("B5", "Alice", "new"));
        let merged = merge_comments(existing, &ops, &mut authors, &BTreeMap::new());
        assert_eq!(authors.len(), 1);
        assert_eq!(merged["A1"].author_id, 0);
        assert_eq!(merged["B5"].author_id, 0);
    }

    #[test]
    fn merge_set_assigns_new_author_id() {
        let mut authors = CommentAuthorTable::new();
        let mut existing = BTreeMap::new();
        existing.insert(
            "A1".to_string(),
            ExistingComment {
                coordinate: "A1".into(),
                author: "Alice".into(),
                text_inner_xml: "<t>x</t>".into(),
                comment_pr: None,
                ext_lst: None,
                width_pt: None,
                height_pt: None,
            },
        );
        authors.intern("Alice");
        let mut ops = BTreeMap::new();
        ops.insert("B5".to_string(), op_set("B5", "Bob", "y"));
        let merged = merge_comments(existing, &ops, &mut authors, &BTreeMap::new());
        assert_eq!(authors.len(), 2);
        assert_eq!(merged["A1"].author_id, 0);
        assert_eq!(merged["B5"].author_id, 1);
    }

    #[test]
    fn delete_op_removes_comment() {
        let mut authors = CommentAuthorTable::new();
        let mut existing = BTreeMap::new();
        existing.insert(
            "A1".to_string(),
            ExistingComment {
                coordinate: "A1".into(),
                author: "Alice".into(),
                text_inner_xml: "<t>x</t>".into(),
                comment_pr: None,
                ext_lst: None,
                width_pt: None,
                height_pt: None,
            },
        );
        authors.intern("Alice");
        let mut ops = BTreeMap::new();
        ops.insert("A1".to_string(), CommentOp::Delete);
        let merged = merge_comments(existing, &ops, &mut authors, &BTreeMap::new());
        assert!(merged.is_empty());
    }

    #[test]
    fn build_vml_default_widths_matches_native_writer_d5() {
        let mut merged = BTreeMap::new();
        merged.insert(
            "D5".to_string(),
            MergedComment {
                coordinate: "D5".into(),
                author_id: 0,
                text_inner_xml: "<t>x</t>".into(),
                comment_pr: None,
                ext_lst: None,
                width_pt: None,
                height_pt: None,
            },
        );
        let bytes = build_vml(&merged, &[], &ColWidthMap::new(), None);
        let s = std::str::from_utf8(&bytes).unwrap();
        assert!(s.contains("margin-left:203.25pt"), "{s}");
        assert!(s.contains("margin-top:52.5pt"), "{s}");
        assert!(s.contains("<x:Anchor>4, 15, 3, 10, 6, 15, 7, 4</x:Anchor>"));
    }

    #[test]
    fn build_vml_uses_actual_col_widths() {
        let sheet_xml = br#"<worksheet><cols>
<col min="1" max="1" width="20" customWidth="1"/>
</cols><sheetData/></worksheet>"#;
        let cols = parse_col_widths(sheet_xml);
        assert!(cols.get(0).is_some(), "col 0 width parsed");
        let mut merged = BTreeMap::new();
        merged.insert(
            "B2".to_string(),
            MergedComment {
                coordinate: "B2".into(),
                author_id: 0,
                text_inner_xml: "<t>x</t>".into(),
                comment_pr: None,
                ext_lst: None,
                width_pt: None,
                height_pt: None,
            },
        );
        let bytes = build_vml(&merged, &[], &cols, None);
        let s = std::str::from_utf8(&bytes).unwrap();
        // Column A 20 units -> 149 px -> 111.75 pt.
        // Native writer would emit (1 * 48 + 59.25) = 107.25 pt regardless
        // of width. The patcher's width-aware path should reflect col A's
        // actual size: ORIGIN_LEFT_PT + 111.75 = 171.75 pt.
        assert!(
            s.contains("margin-left:171.75pt"),
            "patcher uses actual width: {s}"
        );
    }

    #[test]
    fn build_vml_preserves_non_comment_shape() {
        let preserved = vec![PreservedVmlShape {
            raw_xml: r##"<v:shape id="LH" type="#_x0000_t75" style="position:absolute"><v:imagedata/></v:shape>"##.to_string(),
        }];
        let merged = BTreeMap::new();
        let bytes = build_vml(&merged, &preserved, &ColWidthMap::new(), None);
        let s = std::str::from_utf8(&bytes).unwrap();
        assert!(s.contains(r#"id="LH""#), "preserved shape kept: {s}");
        assert!(s.contains("#_x0000_t75"));
    }

    #[test]
    fn extract_vml_filters_comment_shape_keeps_image() {
        let vml = r##"<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="31"/></o:shapelayout>
<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"/>
<v:shape id="LH" type="#_x0000_t75" style="position:absolute"><v:imagedata/></v:shape>
<v:shape id="_x0000_s1025" type="#_x0000_t202" style="margin-left:59.25pt; margin-top:1.5pt; width:96pt; height:55.5pt"><x:ClientData ObjectType="Note"><x:Row>0</x:Row><x:Column>0</x:Column></x:ClientData></v:shape>
</xml>"##;
        let out = extract_vml(vml.as_bytes());
        assert_eq!(out.preserved.len(), 1);
        assert!(out.preserved[0].raw_xml.contains(r#"id="LH""#));
        assert_eq!(out.idmap_data.as_deref(), Some("31"));
        assert_eq!(
            out.sizes.get("A1").map(|(w, h)| (w.unwrap(), h.unwrap())),
            Some((96.0, 55.5))
        );
    }

    #[test]
    fn col_widths_parser_expands_ranges() {
        let xml = br#"<worksheet><cols>
<col min="1" max="3" width="20" customWidth="1"/>
<col min="5" max="5" width="40"/>
</cols></worksheet>"#;
        let cols = parse_col_widths(xml);
        assert!(cols.get(0).is_some());
        assert!(cols.get(1).is_some());
        assert!(cols.get(2).is_some());
        assert!(cols.get(3).is_none());
        assert!(cols.get(4).is_some());
    }

    #[test]
    fn build_comments_drives_full_pipeline_minimal() {
        let mut rels = RelsGraph::new();
        let mut authors = CommentAuthorTable::new();
        let mut ops = BTreeMap::new();
        ops.insert("A1".to_string(), op_set("A1", "Wolf", "hi"));
        let sheet_xml = br#"<worksheet><sheetData/></worksheet>"#;
        let (result, comments_rid, vml_rid) =
            build_comments(None, None, &ops, sheet_xml, &mut rels, &mut authors, 1, 1);
        assert!(comments_rid.is_some());
        assert!(vml_rid.is_some());
        assert_eq!(result.legacy_drawing_rid, vml_rid);
        assert!(!result.comments_xml.is_empty());
        assert!(!result.vml_drawing.is_empty());
        let cs = std::str::from_utf8(&result.comments_xml).unwrap();
        assert!(cs.contains("<author>Wolf</author>"));
        assert!(cs.contains(r#"ref="A1""#));
        assert_eq!(rels.len(), 2);
    }

    #[test]
    fn build_comments_delete_last_drops_parts() {
        let existing_comments = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<authors><author>Alice</author></authors>
<commentList><comment ref="A1" authorId="0"><text><t>x</t></text></comment></commentList>
</comments>"#;
        let existing_vml = r##"<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>
<v:shapetype id="_x0000_t202"/>
<v:shape id="_x0000_s1025" type="#_x0000_t202" style="margin-left:59.25pt; margin-top:1.5pt; width:96pt; height:55.5pt"><x:ClientData ObjectType="Note"><x:Row>0</x:Row><x:Column>0</x:Column></x:ClientData></v:shape>
</xml>"##;
        let mut rels = RelsGraph::new();
        rels.add(
            rt::COMMENTS,
            "../comments/comments1.xml",
            TargetMode::Internal,
        );
        rels.add(
            rt::VML_DRAWING,
            "../drawings/vmlDrawing1.vml",
            TargetMode::Internal,
        );
        let mut authors = CommentAuthorTable::new();
        let mut ops = BTreeMap::new();
        ops.insert("A1".to_string(), CommentOp::Delete);
        let sheet_xml = br#"<worksheet><sheetData/></worksheet>"#;
        let (result, _crid, _vrid) = build_comments(
            Some(existing_comments.as_bytes()),
            Some(existing_vml.as_bytes()),
            &ops,
            sheet_xml,
            &mut rels,
            &mut authors,
            1,
            1,
        );
        assert!(result.comments_xml.is_empty());
        assert!(result.vml_drawing.is_empty());
        assert!(result.legacy_drawing_rid.is_none());
        assert!(rels.is_empty());
    }

    #[test]
    fn build_comments_delete_keeps_image_vml() {
        let existing_comments = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<authors><author>Alice</author></authors>
<commentList><comment ref="A1" authorId="0"><text><t>x</t></text></comment></commentList>
</comments>"#;
        let existing_vml = r##"<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="2"/></o:shapelayout>
<v:shapetype id="_x0000_t202"/>
<v:shape id="LH" type="#_x0000_t75" style="position:absolute"><v:imagedata/></v:shape>
<v:shape id="_x0000_s1025" type="#_x0000_t202" style="margin-left:59.25pt; margin-top:1.5pt; width:96pt; height:55.5pt"><x:ClientData ObjectType="Note"><x:Row>0</x:Row><x:Column>0</x:Column></x:ClientData></v:shape>
</xml>"##;
        let mut rels = RelsGraph::new();
        rels.add(
            rt::COMMENTS,
            "../comments/comments1.xml",
            TargetMode::Internal,
        );
        rels.add(
            rt::VML_DRAWING,
            "../drawings/vmlDrawing1.vml",
            TargetMode::Internal,
        );
        let mut authors = CommentAuthorTable::new();
        let mut ops = BTreeMap::new();
        ops.insert("A1".to_string(), CommentOp::Delete);
        let sheet_xml = br#"<worksheet><sheetData/></worksheet>"#;
        let (result, _crid, vml_rid) = build_comments(
            Some(existing_comments.as_bytes()),
            Some(existing_vml.as_bytes()),
            &ops,
            sheet_xml,
            &mut rels,
            &mut authors,
            1,
            1,
        );
        assert!(result.comments_xml.is_empty());
        assert!(!result.vml_drawing.is_empty(), "image preserved");
        let s = std::str::from_utf8(&result.vml_drawing).unwrap();
        assert!(s.contains(r#"id="LH""#), "{s}");
        assert!(result.legacy_drawing_rid == vml_rid);
        assert_eq!(rels.len(), 1);
    }

    #[test]
    fn rowcol_a1_round_trip() {
        assert_eq!(rowcol0_to_a1(0, 0), "A1");
        assert_eq!(rowcol0_to_a1(4, 3), "D5");
        assert_eq!(rowcol0_to_a1(99, 25), "Z100");
        assert_eq!(rowcol0_to_a1(0, 26), "AA1");
        assert_eq!(a1_to_row_col0("A1"), Some((0, 0)));
        assert_eq!(a1_to_row_col0("D5"), Some((4, 3)));
        assert_eq!(a1_to_row_col0("Z100"), Some((99, 25)));
        assert_eq!(a1_to_row_col0("AA1"), Some((0, 26)));
    }
}
