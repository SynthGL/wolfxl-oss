//! Conditional-formatting block builder for modify mode (RFC-026).
//!
//! The patcher path needs to:
//!   1. Read every existing `<conditionalFormatting>` element out of a
//!      sheet's XML and re-include them verbatim in the output (because
//!      RFC-011's merger uses **replace-all** semantics for slot 17 —
//!      any supplied CF block drops every existing CF block).
//!   2. Append zero or more new `<conditionalFormatting>` blocks built
//!      from `ConditionalFormattingPatch` records, allocating
//!      sheet-unique `priority` values and (where applicable) `dxfId`
//!      indices into `xl/styles.xml`'s `<dxfs>` collection.
//!   3. Emit the new `<dxf>` entries that the workbook coordinator will
//!      thread into a single `xl/styles.xml` mutation.
//!
//! The new-rule serializer mirrors
//! `crates/wolfxl-writer/src/emit/sheet_xml.rs::emit_conditional_formats`
//! (attribute order: `type`, `priority`, optional `operator`, optional
//! `dxfId`, optional `stopIfTrue`). RFC-026 §4.2 explicitly authorizes a
//! parallel implementation rather than refactoring the writer to share —
//! the patcher's input shape (a flat dict-of-strings via PyO3) and the
//! writer's input shape (its in-memory `Worksheet` model) don't compose
//! without an intermediate type that adds complexity, and the writer
//! crate is PyO3-free while the patcher is PyO3-heavy. Mirror the writer
//! by reading its source as a schema reference.

use quick_xml::events::Event;
use quick_xml::Reader as XmlReader;

// ---------------------------------------------------------------------------
// Public types
// ---------------------------------------------------------------------------

/// A differential format (font/fill/border override) for a CF rule.
/// Only the fields being overridden need to be populated. Maps to one
/// `<dxf>` entry in `xl/styles.xml`.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct DxfPatch {
    pub font_bold: Option<bool>,
    pub font_italic: Option<bool>,
    pub font_color_rgb: Option<String>,
    pub fill_pattern_type: Option<String>,
    pub fill_fg_color_rgb: Option<String>,
    pub border_top_style: Option<String>,
    pub border_bottom_style: Option<String>,
    pub border_left_style: Option<String>,
    pub border_right_style: Option<String>,
}

/// A cfvo threshold for colorScale / dataBar.
#[derive(Debug, Clone, PartialEq)]
pub struct CfvoPatch {
    /// `"min"`, `"max"`, `"num"`, `"percent"`, `"percentile"`, or
    /// `"formula"`.
    pub cfvo_type: String,
    /// Required for num/percent/percentile/formula; `None` for min/max.
    pub val: Option<String>,
}

/// A color-scale stop (one cfvo paired with one color).
#[derive(Debug, Clone, PartialEq)]
pub struct ColorScaleStop {
    pub cfvo: CfvoPatch,
    pub color_rgb: String,
}

#[derive(Debug, Clone, PartialEq)]
pub enum CfRuleKind {
    CellIs {
        /// `"equal"`, `"notEqual"`, `"lessThan"`, `"lessThanOrEqual"`,
        /// `"greaterThan"`, `"greaterThanOrEqual"`, `"between"`,
        /// `"notBetween"`.
        operator: String,
        formula_a: String,
        formula_b: Option<String>,
    },
    Expression {
        formula: String,
    },
    ColorScale {
        stops: Vec<ColorScaleStop>,
    },
    DataBar {
        min: CfvoPatch,
        max: CfvoPatch,
        color_rgb: String,
        show_value: bool,
        min_length: Option<u32>,
        max_length: Option<u32>,
    },
    IconSet {
        set_name: String,
        thresholds: Vec<CfvoPatch>,
        show_value: bool,
        percent: Option<bool>,
        reverse: Option<bool>,
    },
    Generic {
        type_name: String,
        attrs: Vec<(String, String)>,
        formulas: Vec<String>,
    },
}

#[derive(Debug, Clone, PartialEq)]
pub struct CfRulePatch {
    pub kind: CfRuleKind,
    /// Differential format for rules that set cell styling. `None` for
    /// colorScale / dataBar (those carry their own color inline).
    pub dxf: Option<DxfPatch>,
    pub stop_if_true: bool,
}

#[derive(Debug, Clone, PartialEq)]
pub struct ConditionalFormattingPatch {
    /// Space-separated A1 ranges. Required.
    pub sqref: String,
    pub rules: Vec<CfRulePatch>,
}

/// Output of [`build_cf_blocks`].
#[derive(Debug, Clone)]
pub struct CfResult {
    /// Concatenation of (existing blocks, verbatim) followed by
    /// (new blocks, freshly serialized). Hand directly to
    /// `wolfxl_merger::merge_blocks` as
    /// `SheetBlock::ConditionalFormatting`.
    pub block_bytes: Vec<u8>,
    /// New `<dxf>` patches that need to be appended to
    /// `xl/styles.xml`'s `<dxfs>` collection. Ordered to match the
    /// `dxfId` values already baked into `block_bytes`.
    pub new_dxfs: Vec<DxfPatch>,
}

// ---------------------------------------------------------------------------
// extract_existing_cf_blocks
// ---------------------------------------------------------------------------

/// Walk a sheet XML and return the raw byte range of each
/// `<conditionalFormatting>...</conditionalFormatting>` element (or
/// self-closing form) in source order. Empty `Vec` if no CF blocks.
///
/// Direct multi-element analog of
/// `validations::extract_existing_dv_block`. The merger's replace-all
/// CF semantics (RFC-011 §5.5) make this primitive load-bearing for
/// preservation: the patcher must re-include these bytes verbatim in
/// its supplied block payload.
pub fn extract_existing_cf_blocks(sheet_xml: &str) -> Vec<Vec<u8>> {
    let bytes = sheet_xml.as_bytes();
    let mut reader = XmlReader::from_str(sheet_xml);
    reader.config_mut().trim_text(false);
    let mut buf: Vec<u8> = Vec::new();

    let mut blocks: Vec<Vec<u8>> = Vec::new();
    let mut start_pos: Option<usize> = None;
    let mut depth: usize = 0;
    let mut cf_depth: u32 = 0;
    let mut search_from: usize = 0;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let current_depth = depth;
                if current_depth == 1 && e.local_name().as_ref() == "conditionalFormatting" {
                    if start_pos.is_none() {
                        start_pos = find_local_start(bytes, search_from, b"conditionalFormatting");
                        cf_depth = 1;
                    } else {
                        cf_depth += 1;
                    }
                } else if start_pos.is_some() && e.local_name().as_ref() == "conditionalFormatting"
                {
                    cf_depth += 1;
                }
                depth += 1;
            }
            Ok(Event::Empty(ref e)) => {
                if depth == 1
                    && e.local_name().as_ref() == "conditionalFormatting"
                    && start_pos.is_none()
                {
                    if let Some(start) =
                        find_local_start(bytes, search_from, b"conditionalFormatting")
                    {
                        if let Some(end) = find_tag_end(bytes, start) {
                            blocks.push(bytes[start..end].to_vec());
                            search_from = end;
                        }
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                if e.local_name().as_ref() == "conditionalFormatting" && start_pos.is_some() {
                    cf_depth -= 1;
                    if cf_depth == 0 {
                        let start = start_pos.take().expect("end without start");
                        if let Some(end) =
                            find_local_end_tag_end(bytes, start, b"conditionalFormatting")
                        {
                            blocks.push(bytes[start..end].to_vec());
                            search_from = end;
                        }
                    }
                }
                depth = depth.saturating_sub(1);
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    blocks
}

fn find_local_start(bytes: &[u8], from: usize, local: &[u8]) -> Option<usize> {
    let mut i = from;
    while i < bytes.len() {
        let rel = bytes[i..].iter().position(|b| *b == b'<')?;
        i += rel;
        let name_start = i + 1;
        if name_start >= bytes.len() {
            return None;
        }
        let first = bytes[name_start];
        if first == b'/' || first == b'!' || first == b'?' {
            i = name_start + 1;
            continue;
        }
        let mut local_start = name_start;
        let mut j = name_start;
        while j < bytes.len() {
            match bytes[j] {
                b':' => {
                    local_start = j + 1;
                    j += 1;
                }
                b' ' | b'\t' | b'\r' | b'\n' | b'/' | b'>' => break,
                _ => j += 1,
            }
        }
        if bytes.get(local_start..j) == Some(local) {
            return Some(i);
        }
        i = j.saturating_add(1);
    }
    None
}

fn find_tag_end(bytes: &[u8], start: usize) -> Option<usize> {
    bytes[start..]
        .iter()
        .position(|b| *b == b'>')
        .map(|pos| start + pos + 1)
}

fn find_local_end_tag_end(bytes: &[u8], from: usize, local: &[u8]) -> Option<usize> {
    let mut i = from;
    while i < bytes.len() {
        let rel = bytes[i..].iter().position(|b| *b == b'<')?;
        i += rel;
        let slash = i + 1;
        if bytes.get(slash) != Some(&b'/') {
            i = slash.saturating_add(1);
            continue;
        }
        let name_start = slash + 1;
        let mut local_start = name_start;
        let mut j = name_start;
        while j < bytes.len() {
            match bytes[j] {
                b':' => {
                    local_start = j + 1;
                    j += 1;
                }
                b' ' | b'\t' | b'\r' | b'\n' | b'>' => break,
                _ => j += 1,
            }
        }
        if bytes.get(local_start..j) == Some(local) {
            return find_tag_end(bytes, i);
        }
        i = j.saturating_add(1);
    }
    None
}

// ---------------------------------------------------------------------------
// scan_max_cf_priority
// ---------------------------------------------------------------------------

/// Scan a sheet XML and return the maximum `priority` value across all
/// `<cfRule>` elements. Returns 0 if no CF rules exist (so callers can
/// allocate the next priority as `existing_priority_max + 1` starting
/// at 1).
///
/// ECMA §18.3.1.10 requires `priority` to be unique across every
/// `<cfRule>` in the sheet — not just within one
/// `<conditionalFormatting>` group. Missing the global max here would
/// produce duplicate priorities, which Excel applies non-deterministically
/// (RFC §8 risk #2).
pub fn scan_max_cf_priority(sheet_xml: &str) -> u32 {
    let mut reader = XmlReader::from_str(sheet_xml);
    reader.config_mut().trim_text(false);
    let mut buf: Vec<u8> = Vec::new();
    let mut max_priority: u32 = 0;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                if e.local_name().as_ref() == "cfRule" {
                    for attr in e.attributes().with_checks(false).flatten() {
                        if attr.key.local_name().as_ref() == "priority" {
                            if let Ok(v) = attr.normalized_value(quick_xml::XmlVersion::Implicit1_0) {
                                if let Ok(p) = v.parse::<u32>() {
                                    if p > max_priority {
                                        max_priority = p;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    max_priority
}

// ---------------------------------------------------------------------------
// count_dxfs
// ---------------------------------------------------------------------------

/// Count the number of `<dxf>` direct children of `<dxfs>` in
/// `xl/styles.xml`. Returns 0 if the section is absent.
///
/// **Counts children, NOT the `count="N"` attribute** — some Excel
/// versions emit incorrect counts (RFC §8.1 / §8 risk #1, HIGH). An
/// off-by-one here silently applies the wrong formatting.
pub fn count_dxfs(styles_xml: &str) -> u32 {
    let mut reader = XmlReader::from_str(styles_xml);
    reader.config_mut().trim_text(false);
    let mut buf: Vec<u8> = Vec::new();
    let mut count: u32 = 0;
    let mut in_dxfs: u32 = 0;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                if e.local_name().as_ref() == "dxfs" {
                    in_dxfs += 1;
                } else if in_dxfs > 0 && e.local_name().as_ref() == "dxf" {
                    count += 1;
                }
            }
            Ok(Event::Empty(ref e)) => {
                if in_dxfs > 0 && e.local_name().as_ref() == "dxf" {
                    count += 1;
                }
            }
            Ok(Event::End(ref e)) => {
                if e.local_name().as_ref() == "dxfs" && in_dxfs > 0 {
                    in_dxfs -= 1;
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    count
}

// ---------------------------------------------------------------------------
// build_cf_blocks
// ---------------------------------------------------------------------------

/// Build the `<conditionalFormatting>` block payload for one worksheet.
///
/// `existing_blocks` is the list of source CF block byte-ranges (from
/// [`extract_existing_cf_blocks`]). They are emitted verbatim, in
/// supplied order, BEFORE any new blocks. This compensates for the
/// merger's replace-all CF semantics: any `SheetBlock::ConditionalFormatting`
/// in the merger's input drops every existing source CF block, so
/// callers must re-include them here.
///
/// `existing_priority_max` is the max `priority` value already on the
/// sheet (typically from [`scan_max_cf_priority`]). New rules are
/// allocated `existing_priority_max + 1`, `+2`, … in order.
///
/// `existing_dxf_count` is the count of `<dxf>` entries currently in
/// `xl/styles.xml` (typically from [`count_dxfs`]). New `dxfId` values
/// start at `existing_dxf_count`.
///
/// Empty-rule wrappers are skipped: if a `ConditionalFormattingPatch`
/// has no rules (or all rules are stub variants the caller filtered
/// out), no `<conditionalFormatting>` element is emitted for it
/// (RFC §8 risk #5 — Excel "repairs" empty wrappers on open).
#[allow(dead_code)]
pub fn build_cf_blocks(
    existing_blocks: &[Vec<u8>],
    patches: &[ConditionalFormattingPatch],
    existing_priority_max: u32,
    existing_dxf_count: u32,
) -> CfResult {
    build_cf_blocks_with_prefix(
        existing_blocks,
        patches,
        existing_priority_max,
        existing_dxf_count,
        "",
    )
}

pub fn build_cf_blocks_with_prefix(
    existing_blocks: &[Vec<u8>],
    patches: &[ConditionalFormattingPatch],
    existing_priority_max: u32,
    existing_dxf_count: u32,
    element_prefix: &str,
) -> CfResult {
    let mut out: Vec<u8> = Vec::with_capacity(256);
    let mut new_dxfs: Vec<DxfPatch> = Vec::new();

    for existing in existing_blocks {
        out.extend_from_slice(existing);
    }

    let mut next_priority = existing_priority_max + 1;
    let mut next_dxf_id = existing_dxf_count;

    for patch in patches {
        if patch.rules.is_empty() {
            continue;
        }

        let mut rules_buf: Vec<u8> = Vec::with_capacity(128);
        for rule in &patch.rules {
            let priority = next_priority;
            let dxf_id = if rule.dxf.is_some() {
                let id = next_dxf_id;
                next_dxf_id += 1;
                Some(id)
            } else {
                None
            };

            match &rule.kind {
                CfRuleKind::CellIs {
                    operator,
                    formula_a,
                    formula_b,
                } => {
                    rules_buf.extend_from_slice(
                        format!(
                            "<cfRule type=\"cellIs\" priority=\"{}\" operator=\"{}\"",
                            priority,
                            attr_escape(operator),
                        )
                        .as_bytes(),
                    );
                    if let Some(id) = dxf_id {
                        rules_buf.extend_from_slice(format!(" dxfId=\"{}\"", id).as_bytes());
                    }
                    if rule.stop_if_true {
                        rules_buf.extend_from_slice(b" stopIfTrue=\"1\"");
                    }
                    rules_buf.push(b'>');
                    rules_buf.extend_from_slice(
                        format!("<formula>{}</formula>", text_escape(formula_a)).as_bytes(),
                    );
                    let needs_second = matches!(operator.as_str(), "between" | "notBetween");
                    if needs_second {
                        if let Some(fb) = formula_b {
                            rules_buf.extend_from_slice(
                                format!("<formula>{}</formula>", text_escape(fb)).as_bytes(),
                            );
                        }
                    }
                    rules_buf.extend_from_slice(b"</cfRule>");
                }
                CfRuleKind::Expression { formula } => {
                    rules_buf.extend_from_slice(
                        format!("<cfRule type=\"expression\" priority=\"{}\"", priority).as_bytes(),
                    );
                    if let Some(id) = dxf_id {
                        rules_buf.extend_from_slice(format!(" dxfId=\"{}\"", id).as_bytes());
                    }
                    if rule.stop_if_true {
                        rules_buf.extend_from_slice(b" stopIfTrue=\"1\"");
                    }
                    rules_buf.push(b'>');
                    rules_buf.extend_from_slice(
                        format!("<formula>{}</formula>", text_escape(formula)).as_bytes(),
                    );
                    rules_buf.extend_from_slice(b"</cfRule>");
                }
                CfRuleKind::ColorScale { stops } => {
                    // ColorScale carries inline color — never has a dxfId.
                    rules_buf.extend_from_slice(
                        format!("<cfRule type=\"colorScale\" priority=\"{}\"", priority).as_bytes(),
                    );
                    if rule.stop_if_true {
                        rules_buf.extend_from_slice(b" stopIfTrue=\"1\"");
                    }
                    rules_buf.push(b'>');
                    rules_buf.extend_from_slice(b"<colorScale>");
                    for stop in stops {
                        emit_cfvo(&mut rules_buf, &stop.cfvo);
                    }
                    for stop in stops {
                        rules_buf.extend_from_slice(
                            format!("<color rgb=\"{}\"/>", attr_escape(&stop.color_rgb)).as_bytes(),
                        );
                    }
                    rules_buf.extend_from_slice(b"</colorScale>");
                    rules_buf.extend_from_slice(b"</cfRule>");
                }
                CfRuleKind::DataBar {
                    min,
                    max,
                    color_rgb,
                    show_value,
                    min_length,
                    max_length,
                } => {
                    rules_buf.extend_from_slice(
                        format!("<cfRule type=\"dataBar\" priority=\"{}\"", priority).as_bytes(),
                    );
                    if rule.stop_if_true {
                        rules_buf.extend_from_slice(b" stopIfTrue=\"1\"");
                    }
                    rules_buf.push(b'>');
                    rules_buf.extend_from_slice(b"<dataBar");
                    if !*show_value {
                        rules_buf.extend_from_slice(b" showValue=\"0\"");
                    }
                    if let Some(value) = min_length {
                        rules_buf.extend_from_slice(format!(" minLength=\"{}\"", value).as_bytes());
                    }
                    if let Some(value) = max_length {
                        rules_buf.extend_from_slice(format!(" maxLength=\"{}\"", value).as_bytes());
                    }
                    rules_buf.push(b'>');
                    emit_cfvo(&mut rules_buf, min);
                    emit_cfvo(&mut rules_buf, max);
                    rules_buf.extend_from_slice(
                        format!("<color rgb=\"{}\"/>", attr_escape(color_rgb)).as_bytes(),
                    );
                    rules_buf.extend_from_slice(b"</dataBar>");
                    rules_buf.extend_from_slice(b"</cfRule>");
                }
                CfRuleKind::IconSet {
                    set_name,
                    thresholds,
                    show_value,
                    percent,
                    reverse,
                } => {
                    rules_buf.extend_from_slice(
                        format!("<cfRule type=\"iconSet\" priority=\"{}\"", priority).as_bytes(),
                    );
                    if rule.stop_if_true {
                        rules_buf.extend_from_slice(b" stopIfTrue=\"1\"");
                    }
                    rules_buf.push(b'>');
                    rules_buf.extend_from_slice(
                        format!("<iconSet iconSet=\"{}\"", attr_escape(set_name)).as_bytes(),
                    );
                    if !*show_value {
                        rules_buf.extend_from_slice(b" showValue=\"0\"");
                    }
                    if let Some(value) = percent {
                        rules_buf.extend_from_slice(if *value {
                            b" percent=\"1\""
                        } else {
                            b" percent=\"0\""
                        });
                    }
                    if let Some(value) = reverse {
                        rules_buf.extend_from_slice(if *value {
                            b" reverse=\"1\""
                        } else {
                            b" reverse=\"0\""
                        });
                    }
                    rules_buf.push(b'>');
                    for threshold in thresholds {
                        emit_cfvo(&mut rules_buf, threshold);
                    }
                    rules_buf.extend_from_slice(b"</iconSet>");
                    rules_buf.extend_from_slice(b"</cfRule>");
                }
                CfRuleKind::Generic {
                    type_name,
                    attrs,
                    formulas,
                } => {
                    rules_buf.extend_from_slice(
                        format!(
                            "<cfRule type=\"{}\" priority=\"{}\"",
                            attr_escape(type_name),
                            priority
                        )
                        .as_bytes(),
                    );
                    for (name, value) in attrs {
                        if value.is_empty() {
                            continue;
                        }
                        rules_buf.extend_from_slice(
                            format!(" {}=\"{}\"", attr_escape(name), attr_escape(value)).as_bytes(),
                        );
                    }
                    if let Some(id) = dxf_id {
                        rules_buf.extend_from_slice(format!(" dxfId=\"{}\"", id).as_bytes());
                    }
                    if rule.stop_if_true {
                        rules_buf.extend_from_slice(b" stopIfTrue=\"1\"");
                    }
                    if formulas.is_empty() {
                        rules_buf.extend_from_slice(b"/>");
                    } else {
                        rules_buf.push(b'>');
                        for formula in formulas {
                            rules_buf.extend_from_slice(
                                format!("<formula>{}</formula>", text_escape(formula)).as_bytes(),
                            );
                        }
                        rules_buf.extend_from_slice(b"</cfRule>");
                    }
                }
            }

            if let Some(ref dxf) = rule.dxf {
                new_dxfs.push(dxf.clone());
            }
            next_priority += 1;
        }

        if rules_buf.is_empty() {
            continue;
        }

        let conditional_formatting_open = prefixed_tag(element_prefix, "conditionalFormatting");
        let conditional_formatting_close =
            prefixed_close_tag(element_prefix, "conditionalFormatting");
        let rules_xml = if element_prefix.is_empty() {
            rules_buf
        } else {
            prefix_generated_cf_tags(&rules_buf, element_prefix)
        };
        out.extend_from_slice(
            format!(
                "<{} sqref=\"{}\">",
                conditional_formatting_open,
                attr_escape(&patch.sqref)
            )
            .as_bytes(),
        );
        out.extend_from_slice(&rules_xml);
        out.extend_from_slice(conditional_formatting_close.as_bytes());
    }

    CfResult {
        block_bytes: out,
        new_dxfs,
    }
}

pub fn main_xml_prefix(xml: &str, root_local_name: &[u8]) -> String {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buf: Vec<u8> = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                if e.local_name().as_ref().as_bytes() == root_local_name {
                    let name = e.name();
                    if let Some((prefix, _)) = name.as_ref().split_once(':') {
                        return format!("{prefix}:");
                    }
                    return String::new();
                }
            }
            Ok(Event::Eof) | Err(_) => return String::new(),
            _ => {}
        }
        buf.clear();
    }
}

fn prefixed_tag(prefix: &str, local: &str) -> String {
    format!("{prefix}{local}")
}

fn prefixed_close_tag(prefix: &str, local: &str) -> String {
    format!("</{prefix}{local}>")
}

fn prefix_generated_cf_tags(xml: &[u8], prefix: &str) -> Vec<u8> {
    let mut s = String::from_utf8_lossy(xml).into_owned();
    for tag in [
        "cfRule",
        "formula",
        "colorScale",
        "cfvo",
        "color",
        "dataBar",
        "iconSet",
    ] {
        s = s.replace(&format!("<{tag}"), &format!("<{prefix}{tag}"));
        s = s.replace(&format!("</{tag}>"), &format!("</{prefix}{tag}>"));
    }
    s.into_bytes()
}

fn emit_cfvo(out: &mut Vec<u8>, cfvo: &CfvoPatch) {
    match cfvo.val.as_deref() {
        Some(v) => {
            out.extend_from_slice(
                format!(
                    "<cfvo type=\"{}\" val=\"{}\"/>",
                    attr_escape(&cfvo.cfvo_type),
                    attr_escape(v),
                )
                .as_bytes(),
            );
        }
        None => {
            out.extend_from_slice(
                format!("<cfvo type=\"{}\"/>", attr_escape(&cfvo.cfvo_type)).as_bytes(),
            );
        }
    }
}

// ---------------------------------------------------------------------------
// dxf_to_xml
// ---------------------------------------------------------------------------

/// Serialize one [`DxfPatch`] into a `<dxf>` element. Children are
/// emitted in openpyxl's serialization order (`font`, `fill`, `border`)
/// for downstream-tool friendliness — Excel itself accepts any order
/// but some third-party readers don't.
///
/// Empty `DxfPatch::default()` produces `<dxf/>`. Callers should avoid
/// pushing empty patches into `new_dxfs`; this function tolerates them
/// rather than panicking.
#[allow(dead_code)]
pub fn dxf_to_xml(patch: &DxfPatch) -> String {
    dxf_to_xml_with_prefix(patch, "")
}

pub fn dxf_to_xml_with_prefix(patch: &DxfPatch, element_prefix: &str) -> String {
    let mut s = String::with_capacity(64);
    let dxf = prefixed_tag(element_prefix, "dxf");
    let font = prefixed_tag(element_prefix, "font");
    let b = prefixed_tag(element_prefix, "b");
    let i = prefixed_tag(element_prefix, "i");
    let color = prefixed_tag(element_prefix, "color");
    let fill = prefixed_tag(element_prefix, "fill");
    let pattern_fill = prefixed_tag(element_prefix, "patternFill");
    let fg_color = prefixed_tag(element_prefix, "fgColor");
    let border = prefixed_tag(element_prefix, "border");
    s.push_str(&format!("<{dxf}>"));

    let has_font =
        patch.font_bold.is_some() || patch.font_italic.is_some() || patch.font_color_rgb.is_some();
    if has_font {
        s.push_str(&format!("<{font}>"));
        if patch.font_bold == Some(true) {
            s.push_str(&format!("<{b}/>"));
        } else if patch.font_bold == Some(false) {
            s.push_str(&format!("<{b} val=\"0\"/>"));
        }
        if patch.font_italic == Some(true) {
            s.push_str(&format!("<{i}/>"));
        } else if patch.font_italic == Some(false) {
            s.push_str(&format!("<{i} val=\"0\"/>"));
        }
        if let Some(ref c) = patch.font_color_rgb {
            s.push_str(&format!("<{color} rgb=\"{}\"/>", attr_escape(c)));
        }
        s.push_str(&format!("</{font}>"));
    }

    let has_fill = patch.fill_pattern_type.is_some() || patch.fill_fg_color_rgb.is_some();
    if has_fill {
        s.push_str(&format!("<{fill}><{pattern_fill}"));
        if let Some(ref pt) = patch.fill_pattern_type {
            s.push_str(&format!(" patternType=\"{}\"", attr_escape(pt)));
        }
        if let Some(ref c) = patch.fill_fg_color_rgb {
            s.push_str(&format!(
                "><{fg_color} rgb=\"{}\"/></{pattern_fill}>",
                attr_escape(c)
            ));
        } else {
            s.push_str("/>");
        }
        s.push_str(&format!("</{fill}>"));
    }

    let has_border = patch.border_top_style.is_some()
        || patch.border_bottom_style.is_some()
        || patch.border_left_style.is_some()
        || patch.border_right_style.is_some();
    if has_border {
        s.push_str(&format!("<{border}>"));
        push_border_side(&mut s, element_prefix, "left", &patch.border_left_style);
        push_border_side(&mut s, element_prefix, "right", &patch.border_right_style);
        push_border_side(&mut s, element_prefix, "top", &patch.border_top_style);
        push_border_side(&mut s, element_prefix, "bottom", &patch.border_bottom_style);
        s.push_str(&format!("</{border}>"));
    }

    s.push_str(&format!("</{dxf}>"));
    s
}

fn push_border_side(s: &mut String, prefix: &str, name: &str, style: &Option<String>) {
    let tag = prefixed_tag(prefix, name);
    match style {
        Some(v) => s.push_str(&format!("<{} style=\"{}\"/>", tag, attr_escape(v))),
        None => s.push_str(&format!("<{}/>", tag)),
    }
}

// ---------------------------------------------------------------------------
// ensure_dxfs_section
// ---------------------------------------------------------------------------

/// Append `new_dxfs_xml` (a concatenation of one or more `<dxf>…</dxf>`
/// elements) to the `<dxfs>` section of `xl/styles.xml`. If the section
/// is absent (common in simple workbooks with no prior CF), insert a
/// fresh `<dxfs count="N">{new_dxfs_xml}</dxfs>` immediately before
/// `</styleSheet>`.
///
/// The new-element count is derived from the supplied XML by counting
/// `<dxf` substrings. Caller may supply zero new dxfs — in that case
/// the input is returned unchanged.
pub fn ensure_dxfs_section(styles_xml: &str, new_dxfs_xml: &str) -> String {
    if new_dxfs_xml.is_empty() {
        return styles_xml.to_string();
    }
    let new_count = count_dxf_substrings(new_dxfs_xml);

    if styles_xml.contains("<dxfs") {
        if !styles_xml.contains("</dxfs>") {
            return expand_self_closing_dxfs_section(styles_xml, new_dxfs_xml, new_count);
        }

        // Existing section — use the established inject_into_section, but
        // call it once per child since it bumps count by 1 per call.
        let mut s = styles_xml.to_string();
        for elem in split_dxf_elements(new_dxfs_xml) {
            let (updated, _) = crate::wolfxl::styles::inject_into_section(&s, "dxfs", &elem);
            s = updated;
        }
        return s;
    }

    // Section absent — insert at the schema-valid location. In styles.xml,
    // dxfs must come after cellStyles and before tableStyles/colors/extLst.
    let prefix = main_xml_prefix(styles_xml, b"styleSheet");
    let insert_pos = match find_dxfs_insert_pos(styles_xml, &prefix) {
        Some(p) => p,
        None => return styles_xml.to_string(),
    };
    let dxfs = prefixed_tag(&prefix, "dxfs");
    let mut result = String::with_capacity(styles_xml.len() + new_dxfs_xml.len() + 32);
    result.push_str(&styles_xml[..insert_pos]);
    result.push_str(&format!("<{dxfs} count=\"{}\">", new_count));
    result.push_str(new_dxfs_xml);
    result.push_str(&format!("</{dxfs}>"));
    result.push_str(&styles_xml[insert_pos..]);
    result
}

fn find_dxfs_insert_pos(styles_xml: &str, prefix: &str) -> Option<usize> {
    for tag in ["tableStyles", "colors", "extLst"] {
        if let Some(pos) = styles_xml.find(&format!("<{prefix}{tag}")) {
            return Some(pos);
        }
    }
    styles_xml.find(&format!("</{prefix}styleSheet>"))
}

fn expand_self_closing_dxfs_section(
    styles_xml: &str,
    new_dxfs_xml: &str,
    new_count: u32,
) -> String {
    let Some(open_pos) = styles_xml.find("<dxfs") else {
        return styles_xml.to_string();
    };
    let Some(open_end_rel) = styles_xml[open_pos..].find('>') else {
        return styles_xml.to_string();
    };
    let open_end = open_pos + open_end_rel;
    let open_tag = &styles_xml[open_pos..=open_end];
    if !open_tag.trim_end().ends_with("/>") {
        return styles_xml.to_string();
    }

    let mut opening = open_tag
        .trim_end_matches('>')
        .trim_end_matches('/')
        .to_string();
    opening = update_or_insert_count_attr(&opening, new_count);

    let mut result = String::with_capacity(styles_xml.len() + new_dxfs_xml.len() + 16);
    result.push_str(&styles_xml[..open_pos]);
    result.push_str(&opening);
    result.push('>');
    result.push_str(new_dxfs_xml);
    result.push_str("</dxfs>");
    result.push_str(&styles_xml[open_end + 1..]);
    result
}

fn update_or_insert_count_attr(tag: &str, count: u32) -> String {
    let needle = "count=\"";
    if let Some(start) = tag.find(needle) {
        let val_start = start + needle.len();
        if let Some(end_offset) = tag[val_start..].find('"') {
            let end = val_start + end_offset;
            let mut result = String::with_capacity(tag.len() + 8);
            result.push_str(&tag[..val_start]);
            result.push_str(&count.to_string());
            result.push_str(&tag[end..]);
            return result;
        }
    }

    format!("{tag} count=\"{count}\"")
}

fn count_dxf_substrings(s: &str) -> u32 {
    let mut reader = XmlReader::from_str(s);
    reader.config_mut().trim_text(false);
    let mut buf: Vec<u8> = Vec::new();
    let mut count = 0;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                if e.local_name().as_ref() == "dxf" {
                    count += 1;
                }
            }
            Ok(Event::Eof) | Err(_) => return count,
            _ => {}
        }
        buf.clear();
    }
}

/// Split a concatenation of `<dxf>...</dxf>` (and `<dxf/>`) elements
/// into individual element strings. Used to feed
/// `inject_into_section` one element at a time so its count bookkeeping
/// stays correct.
fn split_dxf_elements(s: &str) -> Vec<String> {
    let mut out: Vec<String> = Vec::new();
    let bytes = s.as_bytes();
    let mut i = 0;
    while i < bytes.len() {
        if i + 4 < bytes.len() && &bytes[i..i + 4] == b"<dxf" {
            // Find the close. A `<dxf/>` self-closes; otherwise it's
            // `<dxf>...</dxf>`.
            let after_open = i + 4;
            let next = bytes.get(after_open).copied().unwrap_or(b' ');
            if next == b'/' {
                if let Some(end) = find_at(bytes, b">", after_open) {
                    out.push(s[i..end + 1].to_string());
                    i = end + 1;
                    continue;
                }
            } else if let Some(end) = find_at(bytes, b"</dxf>", after_open) {
                let final_end = end + b"</dxf>".len();
                out.push(s[i..final_end].to_string());
                i = final_end;
                continue;
            }
        }
        i += 1;
    }
    out
}

fn find_at(haystack: &[u8], needle: &[u8], from: usize) -> Option<usize> {
    if needle.is_empty() || haystack.len() < needle.len() {
        return None;
    }
    let upper = haystack.len() - needle.len();
    let mut i = from;
    while i <= upper {
        if &haystack[i..i + needle.len()] == needle {
            return Some(i);
        }
        i += 1;
    }
    None
}

// ---------------------------------------------------------------------------
// XML escaping helpers (mirror validations.rs)
// ---------------------------------------------------------------------------

fn attr_escape(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
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
    out
}

fn text_escape(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for ch in s.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            _ => out.push(ch),
        }
    }
    out
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn entity_encoded_priority_and_formula_are_preserved() {
        let block = r#"<conditionalFormatting sqref="A1"><cfRule type="expression" priority="&#55;"><formula>A1&lt;&#50;&amp;B1</formula></cfRule></conditionalFormatting>"#;
        let xml = format!("<worksheet>{block}</worksheet>");
        assert_eq!(scan_max_cf_priority(&xml), 7);
        assert_eq!(extract_existing_cf_blocks(&xml), vec![block.as_bytes().to_vec()]);
    }

    fn cellis_gt5() -> CfRulePatch {
        CfRulePatch {
            kind: CfRuleKind::CellIs {
                operator: "greaterThan".to_string(),
                formula_a: "5".to_string(),
                formula_b: None,
            },
            dxf: Some(DxfPatch {
                font_bold: Some(true),
                ..Default::default()
            }),
            stop_if_true: false,
        }
    }

    fn cellis_between_1_10() -> CfRulePatch {
        CfRulePatch {
            kind: CfRuleKind::CellIs {
                operator: "between".to_string(),
                formula_a: "1".to_string(),
                formula_b: Some("10".to_string()),
            },
            dxf: Some(DxfPatch::default()),
            stop_if_true: false,
        }
    }

    // --- extract_existing_cf_blocks ----------------------------------------

    #[test]
    fn extract_returns_empty_when_no_blocks() {
        let xml = r#"<?xml version="1.0"?><worksheet><sheetData/></worksheet>"#;
        assert!(extract_existing_cf_blocks(xml).is_empty());
    }

    #[test]
    fn extract_captures_single_block() {
        let xml = r#"<worksheet><sheetData/><conditionalFormatting sqref="A1:A10"><cfRule type="cellIs" priority="1" operator="greaterThan" dxfId="0"><formula>5</formula></cfRule></conditionalFormatting></worksheet>"#;
        let got = extract_existing_cf_blocks(xml);
        assert_eq!(got.len(), 1);
        let s = std::str::from_utf8(&got[0]).unwrap();
        assert!(s.starts_with("<conditionalFormatting"));
        assert!(s.ends_with("</conditionalFormatting>"));
        assert!(s.contains("priority=\"1\""));
        assert!(s.contains("dxfId=\"0\""));
    }

    #[test]
    fn extract_captures_two_blocks_in_order() {
        let xml = r#"<worksheet><sheetData/><conditionalFormatting sqref="A1"><cfRule type="cellIs" priority="1" operator="equal" dxfId="0"><formula>1</formula></cfRule></conditionalFormatting><conditionalFormatting sqref="B1"><cfRule type="cellIs" priority="2" operator="equal" dxfId="1"><formula>2</formula></cfRule></conditionalFormatting></worksheet>"#;
        let got = extract_existing_cf_blocks(xml);
        assert_eq!(got.len(), 2);
        let first = std::str::from_utf8(&got[0]).unwrap();
        let second = std::str::from_utf8(&got[1]).unwrap();
        assert!(first.contains("sqref=\"A1\""));
        assert!(second.contains("sqref=\"B1\""));
    }

    #[test]
    fn extract_ignores_nested_x14_conditional_formatting_extensions() {
        let xml = r#"<worksheet xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><sheetData/><conditionalFormatting sqref="A1"><cfRule type="cellIs"/></conditionalFormatting><extLst><ext><x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type="dataBar"/></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst></worksheet>"#;
        let got = extract_existing_cf_blocks(xml);
        assert_eq!(got.len(), 1);
        let s = std::str::from_utf8(&got[0]).unwrap();
        assert!(s.contains("sqref=\"A1\""));
        assert!(!s.contains("x14:conditionalFormatting"));
    }

    #[test]
    fn extract_prefixed_block_starts_at_open_tag() {
        let xml = r#"<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:sheetData></x:sheetData><x:conditionalFormatting sqref="A1"><x:cfRule type="cellIs"/></x:conditionalFormatting></x:worksheet>"#;
        let got = extract_existing_cf_blocks(xml);
        assert_eq!(got.len(), 1);
        let s = std::str::from_utf8(&got[0]).unwrap();
        assert!(s.starts_with("<x:conditionalFormatting"), "got: {s}");
        assert!(s.ends_with("</x:conditionalFormatting>"), "got: {s}");
        assert!(!s.starts_with("ta>"), "got: {s}");
    }

    #[test]
    fn extract_preserves_inner_escaping() {
        let xml = r#"<worksheet><conditionalFormatting sqref="A1"><cfRule type="expression" priority="1"><formula>A1&gt;5</formula></cfRule></conditionalFormatting></worksheet>"#;
        let got = extract_existing_cf_blocks(xml);
        assert_eq!(got.len(), 1);
        let s = std::str::from_utf8(&got[0]).unwrap();
        // The escape MUST flow through verbatim.
        assert!(s.contains("&gt;"), "got: {s}");
    }

    // --- scan_max_cf_priority ----------------------------------------------

    #[test]
    fn scan_priority_zero_when_no_rules() {
        let xml = r#"<worksheet><sheetData/></worksheet>"#;
        assert_eq!(scan_max_cf_priority(xml), 0);
    }

    #[test]
    fn scan_priority_returns_max_across_blocks() {
        let xml = r#"<worksheet><conditionalFormatting sqref="A1"><cfRule type="cellIs" priority="3" operator="equal"><formula>1</formula></cfRule></conditionalFormatting><conditionalFormatting sqref="B1"><cfRule type="cellIs" priority="7" operator="equal"><formula>2</formula></cfRule><cfRule type="cellIs" priority="2" operator="equal"><formula>3</formula></cfRule></conditionalFormatting></worksheet>"#;
        assert_eq!(scan_max_cf_priority(xml), 7);
    }

    #[test]
    fn scan_priority_skips_unrelated_attrs() {
        // A non-cfRule element with priority="99" must not be picked up.
        let xml = r#"<worksheet><randomThing priority="99"/><conditionalFormatting sqref="A1"><cfRule type="cellIs" priority="3" operator="equal"><formula>1</formula></cfRule></conditionalFormatting></worksheet>"#;
        assert_eq!(scan_max_cf_priority(xml), 3);
    }

    // --- count_dxfs --------------------------------------------------------

    #[test]
    fn count_dxfs_returns_zero_when_section_absent() {
        let xml = r#"<styleSheet><fonts count="1"><font/></fonts></styleSheet>"#;
        assert_eq!(count_dxfs(xml), 0);
    }

    #[test]
    fn count_dxfs_counts_children_not_count_attr() {
        // The wrapper claims count="99" but only has 2 actual <dxf> children.
        let xml = r#"<styleSheet><dxfs count="99"><dxf><font><b/></font></dxf><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill></dxf></dxfs></styleSheet>"#;
        assert_eq!(count_dxfs(xml), 2);
    }

    #[test]
    fn count_dxfs_handles_self_closing_dxf() {
        let xml =
            r#"<styleSheet><dxfs count="2"><dxf/><dxf><font><b/></font></dxf></dxfs></styleSheet>"#;
        assert_eq!(count_dxfs(xml), 2);
    }

    // --- build_cf_blocks ---------------------------------------------------

    #[test]
    fn build_cellis_clean_file() {
        let patches = vec![ConditionalFormattingPatch {
            sqref: "A1:A10".to_string(),
            rules: vec![cellis_gt5()],
        }];
        let result = build_cf_blocks(&[], &patches, 0, 0);
        let s = String::from_utf8(result.block_bytes).unwrap();
        assert!(s.starts_with("<conditionalFormatting sqref=\"A1:A10\">"));
        assert!(s.contains(
            "<cfRule type=\"cellIs\" priority=\"1\" operator=\"greaterThan\" dxfId=\"0\">"
        ));
        assert!(s.contains("<formula>5</formula>"));
        assert!(s.ends_with("</conditionalFormatting>"));
        assert_eq!(result.new_dxfs.len(), 1);
    }

    #[test]
    fn build_cellis_priority_starts_above_existing() {
        let patches = vec![ConditionalFormattingPatch {
            sqref: "B1".to_string(),
            rules: vec![cellis_gt5()],
        }];
        // Existing max priority = 3 → new rule gets priority 4, dxfId starts at 1.
        let result = build_cf_blocks(&[], &patches, 3, 1);
        let s = String::from_utf8(result.block_bytes).unwrap();
        assert!(s.contains("priority=\"4\""), "got: {s}");
        assert!(s.contains("dxfId=\"1\""), "got: {s}");
    }

    #[test]
    fn build_dxf_id_monotonic_across_rules_in_one_patch() {
        let patches = vec![ConditionalFormattingPatch {
            sqref: "A1:A10".to_string(),
            rules: vec![cellis_gt5(), cellis_gt5()],
        }];
        let result = build_cf_blocks(&[], &patches, 0, 0);
        let s = String::from_utf8(result.block_bytes).unwrap();
        assert!(s.contains("dxfId=\"0\""));
        assert!(s.contains("dxfId=\"1\""));
        assert!(s.contains("priority=\"1\""));
        assert!(s.contains("priority=\"2\""));
        assert_eq!(result.new_dxfs.len(), 2);
    }

    #[test]
    fn build_colorscale_omits_dxf_id() {
        let patches = vec![ConditionalFormattingPatch {
            sqref: "A1:A20".to_string(),
            rules: vec![CfRulePatch {
                kind: CfRuleKind::ColorScale {
                    stops: vec![
                        ColorScaleStop {
                            cfvo: CfvoPatch {
                                cfvo_type: "min".to_string(),
                                val: None,
                            },
                            color_rgb: "FFF8696B".to_string(),
                        },
                        ColorScaleStop {
                            cfvo: CfvoPatch {
                                cfvo_type: "max".to_string(),
                                val: None,
                            },
                            color_rgb: "FF63BE7B".to_string(),
                        },
                    ],
                },
                dxf: None,
                stop_if_true: false,
            }],
        }];
        let result = build_cf_blocks(&[], &patches, 5, 3);
        let s = String::from_utf8(result.block_bytes).unwrap();
        assert!(s.contains("type=\"colorScale\""));
        assert!(
            !s.contains("dxfId="),
            "colorScale must not emit dxfId, got: {s}"
        );
        assert_eq!(result.new_dxfs.len(), 0);
        assert!(s.contains("priority=\"6\""));
        assert!(s.contains("<cfvo type=\"min\"/>"));
        assert!(s.contains("<color rgb=\"FFF8696B\"/>"));
    }

    #[test]
    fn build_databar_omits_dxf_id() {
        let patches = vec![ConditionalFormattingPatch {
            sqref: "C1:C20".to_string(),
            rules: vec![CfRulePatch {
                kind: CfRuleKind::DataBar {
                    min: CfvoPatch {
                        cfvo_type: "min".to_string(),
                        val: None,
                    },
                    max: CfvoPatch {
                        cfvo_type: "max".to_string(),
                        val: None,
                    },
                    color_rgb: "FF638EC6".to_string(),
                    show_value: true,
                    min_length: None,
                    max_length: None,
                },
                dxf: None,
                stop_if_true: false,
            }],
        }];
        let result = build_cf_blocks(&[], &patches, 0, 0);
        let s = String::from_utf8(result.block_bytes).unwrap();
        assert!(s.contains("type=\"dataBar\""));
        assert!(!s.contains("dxfId="));
        assert_eq!(result.new_dxfs.len(), 0);
    }

    #[test]
    fn build_expression_escapes_formula_gt_lt() {
        let patches = vec![ConditionalFormattingPatch {
            sqref: "A1".to_string(),
            rules: vec![CfRulePatch {
                kind: CfRuleKind::Expression {
                    formula: "A1>B1".to_string(),
                },
                dxf: Some(DxfPatch::default()),
                stop_if_true: false,
            }],
        }];
        let result = build_cf_blocks(&[], &patches, 0, 0);
        let s = String::from_utf8(result.block_bytes).unwrap();
        assert!(s.contains("<formula>A1&gt;B1</formula>"), "got: {s}");
    }

    #[test]
    fn build_cellis_between_emits_two_formulas() {
        let patches = vec![ConditionalFormattingPatch {
            sqref: "A1".to_string(),
            rules: vec![cellis_between_1_10()],
        }];
        let result = build_cf_blocks(&[], &patches, 0, 0);
        let s = String::from_utf8(result.block_bytes).unwrap();
        assert!(s.contains("operator=\"between\""));
        assert!(s.contains("<formula>1</formula>"));
        assert!(s.contains("<formula>10</formula>"));
    }

    #[test]
    fn build_skips_empty_rule_wrappers() {
        let patches = vec![ConditionalFormattingPatch {
            sqref: "A1".to_string(),
            rules: vec![],
        }];
        let result = build_cf_blocks(&[], &patches, 0, 0);
        assert!(result.block_bytes.is_empty());
    }

    #[test]
    fn build_includes_existing_blocks_verbatim_first() {
        let existing = vec![
            br#"<conditionalFormatting sqref="OLD"><cfRule type="cellIs" priority="1" operator="equal" dxfId="0"><formula>1</formula></cfRule></conditionalFormatting>"#.to_vec(),
        ];
        let patches = vec![ConditionalFormattingPatch {
            sqref: "NEW".to_string(),
            rules: vec![cellis_gt5()],
        }];
        // existing_priority_max=1 (from the existing rule),
        // existing_dxf_count=1 (because existing rule has dxfId=0).
        let result = build_cf_blocks(&existing, &patches, 1, 1);
        let s = String::from_utf8(result.block_bytes).unwrap();
        assert!(s.contains("sqref=\"OLD\""));
        assert!(s.contains("sqref=\"NEW\""));
        // Order: existing first, new second.
        let old_pos = s.find("OLD").unwrap();
        let new_pos = s.find("NEW").unwrap();
        assert!(old_pos < new_pos);
        // New rule got priority=2, dxfId=1.
        assert!(s.contains("priority=\"2\""), "got: {s}");
        assert!(s.contains("dxfId=\"1\""), "got: {s}");
    }

    // --- dxf_to_xml --------------------------------------------------------

    #[test]
    fn dxf_to_xml_bold_only() {
        let p = DxfPatch {
            font_bold: Some(true),
            ..Default::default()
        };
        assert_eq!(dxf_to_xml(&p), "<dxf><font><b/></font></dxf>");
    }

    #[test]
    fn dxf_to_xml_solid_red_fill() {
        let p = DxfPatch {
            fill_pattern_type: Some("solid".to_string()),
            fill_fg_color_rgb: Some("FFFF0000".to_string()),
            ..Default::default()
        };
        let s = dxf_to_xml(&p);
        assert!(s.contains("patternType=\"solid\""));
        assert!(s.contains("<fgColor rgb=\"FFFF0000\"/>"));
    }

    // --- ensure_dxfs_section -----------------------------------------------

    #[test]
    fn ensure_dxfs_inserts_when_absent() {
        let xml =
            r#"<?xml version="1.0"?><styleSheet><fonts count="1"><font/></fonts></styleSheet>"#;
        let new_xml = "<dxf><font><b/></font></dxf>";
        let out = ensure_dxfs_section(xml, new_xml);
        assert!(out.contains("<dxfs count=\"1\">"));
        assert!(out.contains("<dxf><font><b/></font></dxf>"));
        assert!(out.ends_with("</styleSheet>"));
    }

    #[test]
    fn ensure_dxfs_inserts_before_colors_when_absent() {
        let xml = r#"<?xml version="1.0"?><styleSheet><fonts count="1"><font/></fonts><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><colors><indexedColors/></colors></styleSheet>"#;
        let new_xml = "<dxf><font><b/></font></dxf>";
        let out = ensure_dxfs_section(xml, new_xml);
        let dxfs_pos = out.find("<dxfs").expect("dxfs");
        let colors_pos = out.find("<colors").expect("colors");
        assert!(dxfs_pos < colors_pos, "got: {out}");
    }

    #[test]
    fn ensure_dxfs_uses_styles_prefix_when_absent() {
        let xml = r#"<?xml version="1.0"?><x:styleSheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:fonts count="1"><x:font/></x:fonts></x:styleSheet>"#;
        let new_xml = "<x:dxf><x:font><x:b/></x:font></x:dxf>";
        let out = ensure_dxfs_section(xml, new_xml);
        assert!(out.contains("<x:dxfs count=\"1\"><x:dxf>"), "got: {out}");
        assert!(out.ends_with("</x:styleSheet>"));
    }

    #[test]
    fn ensure_dxfs_appends_when_present() {
        let xml = r#"<?xml version="1.0"?><styleSheet><dxfs count="1"><dxf><font><i/></font></dxf></dxfs></styleSheet>"#;
        let new_xml = "<dxf><font><b/></font></dxf>";
        let out = ensure_dxfs_section(xml, new_xml);
        // Both old and new <dxf> present, count incremented.
        assert!(out.contains("<dxfs count=\"2\">"), "got: {out}");
        assert!(out.contains("<font><i/></font>"));
        assert!(out.contains("<font><b/></font>"));
    }

    #[test]
    fn ensure_dxfs_expands_self_closing_section() {
        let xml = r#"<?xml version="1.0"?><styleSheet><dxfs count="0"/><tableStyles count="0"/></styleSheet>"#;
        let new_xml = "<dxf><font><b/></font></dxf>";
        let out = ensure_dxfs_section(xml, new_xml);
        assert!(out.contains("<dxfs count=\"1\"><dxf><font><b/></font></dxf></dxfs>"));
        assert!(out.contains("<tableStyles count=\"0\"/>"));
    }

    #[test]
    fn ensure_dxfs_appends_two_when_present() {
        let xml = r#"<?xml version="1.0"?><styleSheet><dxfs count="1"><dxf><font><i/></font></dxf></dxfs></styleSheet>"#;
        let new_xml = "<dxf><font><b/></font></dxf><dxf/>";
        let out = ensure_dxfs_section(xml, new_xml);
        assert!(out.contains("<dxfs count=\"3\">"), "got: {out}");
    }

    #[test]
    fn dxf_to_xml_with_prefix_prefixes_children() {
        let xml = dxf_to_xml_with_prefix(
            &DxfPatch {
                font_bold: Some(true),
                ..DxfPatch::default()
            },
            "x:",
        );
        assert_eq!(xml, "<x:dxf><x:font><x:b/></x:font></x:dxf>");
    }

    #[test]
    fn build_cf_blocks_with_prefix_prefixes_generated_tags() {
        let patches = vec![ConditionalFormattingPatch {
            sqref: "A1:A2".to_string(),
            rules: vec![CfRulePatch {
                kind: CfRuleKind::CellIs {
                    operator: "greaterThan".to_string(),
                    formula_a: "0".to_string(),
                    formula_b: None,
                },
                dxf: Some(DxfPatch {
                    font_bold: Some(true),
                    ..DxfPatch::default()
                }),
                stop_if_true: false,
            }],
        }];
        let result = build_cf_blocks_with_prefix(&[], &patches, 0, 0, "x:");
        let xml = String::from_utf8(result.block_bytes).unwrap();
        assert!(xml.starts_with("<x:conditionalFormatting"), "got: {xml}");
        assert!(xml.contains("<x:cfRule"));
        assert!(xml.contains("<x:formula>0</x:formula>"));
        assert!(xml.ends_with("</x:conditionalFormatting>"));
    }
}
