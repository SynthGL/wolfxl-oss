//! Styles parser and appender for `xl/styles.xml`.
//!
//! The styles file contains four collections that cellXfs entries reference:
//!   - `<fonts>` (fontId)
//!   - `<fills>` (fillId)
//!   - `<borders>` (borderId)
//!   - `<numFmts>` (numFmtId, optional)
//!
//! Each cell has a `s` attribute (style index) pointing into `<cellXfs>`.
//! A cellXfs `<xf>` combines fontId + fillId + borderId + numFmtId.
//!
//! For patching, WolfXL appends new component entries and a new `<xf>`,
//! then sets the cell's `s` attribute to the new xf index.

use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader as XmlReader;

use crate::ooxml_util::attr_value;

// ---------------------------------------------------------------------------
// Data types
// ---------------------------------------------------------------------------

/// A parsed `<xf>` entry from `<cellXfs>`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XfEntry {
    pub num_fmt_id: u32,
    pub font_id: u32,
    pub fill_id: u32,
    pub border_id: u32,
}

/// Font specification for creating a new `<font>` element.
#[derive(Debug, Clone, Default, PartialEq, Eq, Hash)]
pub struct FontSpec {
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strikethrough: bool,
    pub name: Option<String>,
    pub size: Option<u32>,         // stored as integer points (e.g. 11)
    pub color_rgb: Option<String>, // "FFRRGGBB"
}

/// Fill specification for creating a new `<fill>` element.
#[derive(Debug, Clone, Default, PartialEq, Eq, Hash)]
pub struct FillSpec {
    pub pattern_type: String,         // "solid", "none", etc.
    pub fg_color_rgb: Option<String>, // "FFRRGGBB"
}

/// Border side.
#[derive(Debug, Clone, Default, PartialEq, Eq, Hash)]
pub struct BorderSideSpec {
    pub style: Option<String>,     // "thin", "medium", "thick", etc.
    pub color_rgb: Option<String>, // "FFRRGGBB"
}

/// Border specification for creating a new `<border>` element.
///
/// `diagonal` carries the single `<diagonal>` child shared by both
/// directions; `diagonal_up` and `diagonal_down` are the boolean
/// `diagonalUp` / `diagonalDown` attrs on the parent `<border>`,
/// gating which direction(s) Excel renders.
#[derive(Debug, Clone, Default, PartialEq, Eq, Hash)]
pub struct BorderSpec {
    pub left: BorderSideSpec,
    pub right: BorderSideSpec,
    pub top: BorderSideSpec,
    pub bottom: BorderSideSpec,
    pub diagonal: BorderSideSpec,
    pub diagonal_up: bool,
    pub diagonal_down: bool,
}

/// Alignment specification.
#[derive(Debug, Clone, Default, PartialEq, Eq, Hash)]
pub struct AlignmentSpec {
    pub horizontal: Option<String>,
    pub vertical: Option<String>,
    pub wrap_text: bool,
    pub indent: u32,
    pub text_rotation: u32,
}

/// Full format spec for a cell — used to find-or-create a style.
#[derive(Debug, Clone, Default)]
pub struct FormatSpec {
    pub font: Option<FontSpec>,
    pub fill: Option<FillSpec>,
    pub border: Option<BorderSpec>,
    pub alignment: Option<AlignmentSpec>,
    pub number_format: Option<String>,
}

// ---------------------------------------------------------------------------
// Parsing
// ---------------------------------------------------------------------------

/// Parse `<cellXfs>` entries from styles.xml.
pub fn parse_cellxfs(xml: &str) -> Vec<XfEntry> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf: Vec<u8> = Vec::new();
    let mut in_cellxfs = false;
    let mut entries: Vec<XfEntry> = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                let tag = e.local_name();
                if tag.as_ref() == "cellXfs" {
                    in_cellxfs = true;
                } else if tag.as_ref() == "xf" && in_cellxfs {
                    entries.push(parse_xf_entry(e));
                }
            }
            Ok(Event::End(ref e)) => {
                if e.local_name().as_ref() == "cellXfs" {
                    in_cellxfs = false;
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    entries
}

fn parse_xf_entry(e: &BytesStart<'_>) -> XfEntry {
    let num_fmt_id = attr_value(e, b"numFmtId")
        .and_then(|s| s.parse().ok())
        .unwrap_or(0);
    let font_id = attr_value(e, b"fontId")
        .and_then(|s| s.parse().ok())
        .unwrap_or(0);
    let fill_id = attr_value(e, b"fillId")
        .and_then(|s| s.parse().ok())
        .unwrap_or(0);
    let border_id = attr_value(e, b"borderId")
        .and_then(|s| s.parse().ok())
        .unwrap_or(0);

    XfEntry {
        num_fmt_id,
        font_id,
        fill_id,
        border_id,
    }
}

// ---------------------------------------------------------------------------
// XML generation helpers
// ---------------------------------------------------------------------------

/// Count existing elements in a section (e.g. `<fonts count="N">`).
/// Returns (section_count, byte_offset_of_closing_tag).
pub fn count_section_elements(xml: &str, section_tag: &str) -> (u32, u64) {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf: Vec<u8> = Vec::new();
    let mut in_section = false;
    let mut count: u32 = 0;
    let section_bytes = section_tag.as_bytes();

    // The child tag is the singular form: "fonts" -> "font", "fills" -> "fill", etc.
    // For cellXfs -> xf
    let child_tag: Vec<u8> = if section_tag == "cellXfs" {
        b"xf".to_vec()
    } else if section_tag.ends_with('s') {
        section_tag[..section_tag.len() - 1].as_bytes().to_vec()
    } else {
        section_tag.as_bytes().to_vec()
    };

    let mut end_offset: u64 = 0;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                if e.local_name().as_ref().as_bytes() == section_bytes {
                    in_section = true;
                    // Try to read count from attribute first
                    if let Some(c) = attr_value(e, b"count") {
                        if let Ok(n) = c.parse::<u32>() {
                            count = n;
                        }
                    }
                } else if in_section && e.local_name().as_ref().as_bytes() == child_tag.as_slice() {
                    // Count child elements if count attr wasn't present
                }
            }
            Ok(Event::Empty(ref e)) => {
                if e.local_name().as_ref().as_bytes() == section_bytes {
                    // Self-closing section, empty
                    if let Some(c) = attr_value(e, b"count") {
                        if let Ok(n) = c.parse::<u32>() {
                            count = n;
                        }
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                if e.local_name().as_ref().as_bytes() == section_bytes {
                    in_section = false;
                    end_offset = reader.buffer_position();
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    (count, end_offset)
}

/// Generate `<font>` XML element from a FontSpec.
pub fn font_to_xml(spec: &FontSpec) -> String {
    let mut parts: Vec<String> = Vec::new();
    if spec.bold {
        parts.push("<b/>".to_string());
    }
    if spec.italic {
        parts.push("<i/>".to_string());
    }
    if spec.underline {
        parts.push("<u/>".to_string());
    }
    if spec.strikethrough {
        parts.push("<strike/>".to_string());
    }
    if let Some(sz) = spec.size {
        parts.push(format!("<sz val=\"{sz}\"/>"));
    }
    if let Some(ref rgb) = spec.color_rgb {
        parts.push(format!("<color rgb=\"{rgb}\"/>"));
    }
    if let Some(ref name) = spec.name {
        parts.push(format!("<name val=\"{name}\"/>"));
    }
    format!("<font>{}</font>", parts.join(""))
}

/// Generate `<fill>` XML element from a FillSpec.
pub fn fill_to_xml(spec: &FillSpec) -> String {
    let mut inner = format!("<patternFill patternType=\"{}\"", spec.pattern_type);
    if let Some(ref rgb) = spec.fg_color_rgb {
        inner.push_str(&format!("><fgColor rgb=\"{rgb}\"/></patternFill>"));
    } else {
        inner.push_str("/>");
    }
    format!("<fill>{inner}</fill>")
}

/// Generate `<border>` XML element from a BorderSpec.
pub fn border_to_xml(spec: &BorderSpec) -> String {
    fn side_xml(tag: &str, side: &BorderSideSpec) -> String {
        match (&side.style, &side.color_rgb) {
            (Some(style), Some(rgb)) => {
                format!("<{tag} style=\"{style}\"><color rgb=\"{rgb}\"/></{tag}>")
            }
            (Some(style), None) => format!("<{tag} style=\"{style}\"/>"),
            _ => format!("<{tag}/>"),
        }
    }

    let left = side_xml("left", &spec.left);
    let right = side_xml("right", &spec.right);
    let top = side_xml("top", &spec.top);
    let bottom = side_xml("bottom", &spec.bottom);
    let diagonal = side_xml("diagonal", &spec.diagonal);

    let mut border_attrs = String::new();
    if spec.diagonal_up {
        border_attrs.push_str(" diagonalUp=\"1\"");
    }
    if spec.diagonal_down {
        border_attrs.push_str(" diagonalDown=\"1\"");
    }

    format!("<border{border_attrs}>{left}{right}{top}{bottom}{diagonal}</border>")
}

/// Generate an `<xf>` element from component IDs.
pub fn xf_to_xml(
    font_id: u32,
    fill_id: u32,
    border_id: u32,
    num_fmt_id: u32,
    alignment: Option<&AlignmentSpec>,
    apply_font: bool,
    apply_fill: bool,
    apply_border: bool,
    apply_number_format: bool,
) -> String {
    let mut attrs = format!(
        "numFmtId=\"{num_fmt_id}\" fontId=\"{font_id}\" fillId=\"{fill_id}\" borderId=\"{border_id}\""
    );
    if apply_font {
        attrs.push_str(" applyFont=\"1\"");
    }
    if apply_fill {
        attrs.push_str(" applyFill=\"1\"");
    }
    if apply_border {
        attrs.push_str(" applyBorder=\"1\"");
    }
    if apply_number_format {
        attrs.push_str(" applyNumberFormat=\"1\"");
    }

    if let Some(align) = alignment {
        let mut align_attrs = Vec::new();
        if let Some(ref h) = align.horizontal {
            align_attrs.push(format!("horizontal=\"{h}\""));
        }
        if let Some(ref v) = align.vertical {
            align_attrs.push(format!("vertical=\"{v}\""));
        }
        if align.wrap_text {
            align_attrs.push("wrapText=\"1\"".to_string());
        }
        if align.indent > 0 {
            align_attrs.push(format!("indent=\"{}\"", align.indent));
        }
        if align.text_rotation > 0 {
            align_attrs.push(format!("textRotation=\"{}\"", align.text_rotation));
        }
        if !align_attrs.is_empty() {
            return format!(
                "<xf {attrs} applyAlignment=\"1\"><alignment {}/></xf>",
                align_attrs.join(" ")
            );
        }
    }

    format!("<xf {attrs}/>")
}

// ---------------------------------------------------------------------------
// Style injection into styles.xml
// ---------------------------------------------------------------------------

/// Insert a new element just before the closing tag of a section.
///
/// E.g. insert `<font>...</font>` before `</fonts>` and bump the `count`.
/// A self-closing section such as `<numFmts count="0"/>` (as openpyxl writes
/// it) is expanded to an open/close pair around the new element.
pub fn inject_into_section(xml: &str, section_tag: &str, new_element: &str) -> (String, u32) {
    let close_tag = format!("</{section_tag}>");
    let open_prefix = format!("<{section_tag}");

    // Find the opening tag to update count attribute
    let Some(open_pos) = xml
        .match_indices(&open_prefix)
        .map(|(pos, _)| pos)
        .find(|&pos| {
            matches!(
                xml.as_bytes().get(pos + open_prefix.len()),
                Some(b' ' | b'\t' | b'\r' | b'\n' | b'>' | b'/')
            )
        })
    else {
        return (xml.to_string(), 0);
    };
    let Some(open_len) = xml[open_pos..].find('>') else {
        return (xml.to_string(), 0);
    };
    let open_end = open_pos + open_len;
    let open_tag_str = &xml[open_pos..=open_end];
    let self_closing = open_tag_str.ends_with("/>");

    // Find the closing tag position
    let close_pos = if self_closing {
        open_end + 1
    } else {
        let Some(close_offset) = xml[open_end..].find(&close_tag) else {
            // Section doesn't close — shouldn't happen for well-formed styles.xml
            return (xml.to_string(), 0);
        };
        open_end + close_offset
    };

    let existing_count = extract_count_attr(open_tag_str).unwrap_or(0);
    let new_count = existing_count + 1;
    let new_index = existing_count; // 0-based index of the appended element

    // Build new opening tag with updated count
    let updated_open = update_count_attr(open_tag_str, new_count);

    // Reconstruct XML
    let mut result = String::with_capacity(xml.len() + new_element.len() + 32);
    result.push_str(&xml[..open_pos]);
    if self_closing {
        result.push_str(
            updated_open
                .strip_suffix("/>")
                .unwrap_or(&updated_open)
                .trim_end(),
        );
        result.push('>');
        result.push_str(new_element);
        result.push_str(&close_tag);
    } else {
        result.push_str(&updated_open);
        result.push_str(&xml[open_end + 1..close_pos]);
        result.push_str(new_element);
    }
    result.push_str(&xml[close_pos..]);

    (result, new_index)
}

fn extract_count_attr(tag: &str) -> Option<u32> {
    // Simple regex-free extraction of count="N" from an opening tag string
    let needle = "count=\"";
    let start = tag.find(needle)? + needle.len();
    let end = tag[start..].find('"')? + start;
    tag[start..end].parse().ok()
}

fn update_count_attr(tag: &str, new_count: u32) -> String {
    let needle = "count=\"";
    if let Some(start) = tag.find(needle) {
        let val_start = start + needle.len();
        if let Some(end_offset) = tag[val_start..].find('"') {
            let end = val_start + end_offset;
            let mut result = String::with_capacity(tag.len() + 8);
            result.push_str(&tag[..val_start]);
            result.push_str(&new_count.to_string());
            result.push_str(&tag[end..]);
            return result;
        }
    }
    tag.to_string()
}

/// Find an existing numFmt ID for a custom format code, or return a new ID.
/// Built-in IDs 0-163 are reserved; custom start at 164.
pub fn find_or_create_num_fmt(xml: &str, format_code: &str) -> (String, u32) {
    // Check if format_code matches a built-in
    let builtin_id = builtin_num_fmt_id(format_code);
    if let Some(id) = builtin_id {
        return (xml.to_string(), id);
    }

    // Search existing custom numFmts
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf: Vec<u8> = Vec::new();
    let mut max_id: u32 = 163; // custom IDs start at 164

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                if e.local_name().as_ref() == "numFmt" {
                    let id = attr_value(e, b"numFmtId")
                        .and_then(|s| s.parse::<u32>().ok())
                        .unwrap_or(0);
                    let code = attr_value(e, b"formatCode").unwrap_or_default();
                    if code == format_code {
                        return (xml.to_string(), id);
                    }
                    if id > max_id {
                        max_id = id;
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    // Need to create a new numFmt entry
    let new_id = max_id + 1;
    let new_element = format!(
        "<numFmt numFmtId=\"{new_id}\" formatCode=\"{}\"/>",
        xml_escape(format_code)
    );

    // Try to inject into existing <numFmts> section
    if xml.contains("<numFmts") {
        let (updated, _) = inject_into_section(xml, "numFmts", &new_element);
        (updated, new_id)
    } else {
        // No numFmts section — insert one before <fonts>
        let insert_before = "<fonts";
        if let Some(pos) = xml.find(insert_before) {
            let section = format!("<numFmts count=\"1\">{new_element}</numFmts>");
            let mut result = String::with_capacity(xml.len() + section.len());
            result.push_str(&xml[..pos]);
            result.push_str(&section);
            result.push_str(&xml[pos..]);
            (result, new_id)
        } else {
            (xml.to_string(), new_id)
        }
    }
}

fn builtin_num_fmt_id(code: &str) -> Option<u32> {
    match code {
        "General" => Some(0),
        "0" => Some(1),
        "0.00" => Some(2),
        "#,##0" => Some(3),
        "#,##0.00" => Some(4),
        "\"$\"#,##0_);(\"$\"#,##0)" => Some(5),
        "\"$\"#,##0_);[Red](\"$\"#,##0)" => Some(6),
        "\"$\"#,##0.00_);(\"$\"#,##0.00)" => Some(7),
        "\"$\"#,##0.00_);[Red](\"$\"#,##0.00)" => Some(8),
        "0%" => Some(9),
        "0.00%" => Some(10),
        "0.00E+00" => Some(11),
        "# ?/?" => Some(12),
        "# ??/??" => Some(13),
        "mm-dd-yy" => Some(14),
        "d-mmm-yy" => Some(15),
        "d-mmm" => Some(16),
        "mmm-yy" => Some(17),
        "h:mm AM/PM" => Some(18),
        "h:mm:ss AM/PM" => Some(19),
        "h:mm" => Some(20),
        "h:mm:ss" => Some(21),
        "m/d/yy h:mm" => Some(22),
        "#,##0_);(#,##0)" => Some(37),
        "#,##0_);[Red](#,##0)" => Some(38),
        "#,##0.00_);(#,##0.00)" => Some(39),
        "#,##0.00_);[Red](#,##0.00)" => Some(40),
        r#"_(* #,##0_);_(* \(#,##0\);_(* "-"_);_(@_)"# => Some(41),
        r#"_("$"* #,##0_);_("$"* \(#,##0\);_("$"* "-"_);_(@_)"# => Some(42),
        r#"_(* #,##0.00_);_(* \(#,##0.00\);_(* "-"??_);_(@_)"# => Some(43),
        r#"_("$"* #,##0.00_)_("$"* \(#,##0.00\)_("$"* "-"??_)_(@_)"# => Some(44),
        "mm:ss" => Some(45),
        "[h]:mm:ss" => Some(46),
        "mmss.0" => Some(47),
        "##0.0E+0" => Some(48),
        "@" => Some(49),
        _ => None,
    }
}

fn xml_escape(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

/// Convenience: apply a full FormatSpec to styles.xml, returning updated XML and the xf index.
pub fn apply_format_spec(xml: &str, spec: &FormatSpec) -> (String, u32) {
    let mut xml = xml.to_string();

    // 1. Font
    let font_id = if let Some(ref font) = spec.font {
        let font_xml = font_to_xml(font);
        let (updated, id) = inject_into_section(&xml, "fonts", &font_xml);
        xml = updated;
        id
    } else {
        0
    };

    // 2. Fill
    let fill_id = if let Some(ref fill) = spec.fill {
        let fill_xml = fill_to_xml(fill);
        let (updated, id) = inject_into_section(&xml, "fills", &fill_xml);
        xml = updated;
        id
    } else {
        0
    };

    // 3. Border
    let border_id = if let Some(ref border) = spec.border {
        let border_xml = border_to_xml(border);
        let (updated, id) = inject_into_section(&xml, "borders", &border_xml);
        xml = updated;
        id
    } else {
        0
    };

    // 4. Number format
    let num_fmt_id = if let Some(ref code) = spec.number_format {
        let (updated, id) = find_or_create_num_fmt(&xml, code);
        xml = updated;
        id
    } else {
        0
    };

    // 5. Create cellXfs entry
    let xf_xml = xf_to_xml(
        font_id,
        fill_id,
        border_id,
        num_fmt_id,
        spec.alignment.as_ref(),
        spec.font.is_some(),
        spec.fill.is_some(),
        spec.border.is_some(),
        spec.number_format.is_some(),
    );
    let (xml, xf_index) = inject_into_section(&xml, "cellXfs", &xf_xml);

    (xml, xf_index)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn custom_num_fmt_reuses_entity_encoded_format() {
        let xml = r#"<styleSheet><numFmts count="1"><numFmt numFmtId="164" formatCode="&quot;A&amp;&#x42;&quot;0"/></numFmts></styleSheet>"#;
        let (out, id) = find_or_create_num_fmt(xml, "\"A&B\"0");
        assert_eq!(id, 164);
        assert_eq!(out, xml);
    }

    const MINIMAL_STYLES: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>
</styleSheet>"#;

    #[test]
    fn test_parse_cellxfs() {
        let entries = parse_cellxfs(MINIMAL_STYLES);
        assert_eq!(entries.len(), 1);
        assert_eq!(entries[0].font_id, 0);
        assert_eq!(entries[0].fill_id, 0);
    }

    #[test]
    fn test_inject_font() {
        let spec = FontSpec {
            bold: true,
            size: Some(14),
            name: Some("Arial".to_string()),
            ..Default::default()
        };
        let (updated, idx) = inject_into_section(MINIMAL_STYLES, "fonts", &font_to_xml(&spec));
        assert_eq!(idx, 1); // second font
        assert!(updated.contains("count=\"2\""));
        assert!(updated.contains("<b/>"));
        assert!(updated.contains("<sz val=\"14\"/>"));
        assert!(updated.contains("<name val=\"Arial\"/>"));
    }

    #[test]
    fn test_inject_fill() {
        let spec = FillSpec {
            pattern_type: "solid".to_string(),
            fg_color_rgb: Some("FFFF0000".to_string()),
        };
        let (updated, idx) = inject_into_section(MINIMAL_STYLES, "fills", &fill_to_xml(&spec));
        assert_eq!(idx, 2); // third fill (after none + gray125)
        assert!(updated.contains("count=\"3\""));
        assert!(updated.contains("FFFF0000"));
    }

    #[test]
    fn test_apply_format_spec_full() {
        let spec = FormatSpec {
            font: Some(FontSpec {
                bold: true,
                ..Default::default()
            }),
            fill: Some(FillSpec {
                pattern_type: "solid".to_string(),
                fg_color_rgb: Some("FF00FF00".to_string()),
            }),
            ..Default::default()
        };
        let (updated, xf_idx) = apply_format_spec(MINIMAL_STYLES, &spec);
        assert_eq!(xf_idx, 1); // second xf entry
                               // Verify all sections got updated
        assert!(updated.contains("fontId=\"1\""));
        assert!(updated.contains("fillId=\"2\""));
    }

    #[test]
    fn test_builtin_num_fmt() {
        let (unchanged, id) = find_or_create_num_fmt(MINIMAL_STYLES, "General");
        assert_eq!(id, 0);
        assert_eq!(unchanged, MINIMAL_STYLES);
    }

    #[test]
    fn test_custom_num_fmt() {
        let (updated, id) = find_or_create_num_fmt(MINIMAL_STYLES, "$#,##0.00");
        assert!(id >= 164);
        assert!(updated.contains("$#,##0.00"));
        assert!(updated.contains("<numFmts count=\"1\">"));
    }

    #[test]
    fn custom_num_fmts_fill_self_closing_section_and_get_distinct_ids() {
        // openpyxl writes an empty collection as `<numFmts count="0" />`.
        let xml =
            r#"<styleSheet><numFmts count="0" /><fonts count="1"><font/></fonts></styleSheet>"#;
        let (once, date_id) = find_or_create_num_fmt(xml, "yyyy-mm-dd");
        assert_eq!(date_id, 164);
        assert!(once.contains(
            r#"<numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts><fonts"#
        ));
        let (twice, datetime_id) = find_or_create_num_fmt(&once, "yyyy-mm-dd h:mm:ss");
        assert_eq!(datetime_id, 165);
        assert!(twice.contains(r#"<numFmts count="2">"#));
        assert_eq!(find_or_create_num_fmt(&twice, "yyyy-mm-dd").1, 164);
    }

    #[test]
    fn test_xf_with_alignment() {
        let align = AlignmentSpec {
            horizontal: Some("center".to_string()),
            wrap_text: true,
            ..Default::default()
        };
        let xf = xf_to_xml(0, 0, 0, 0, Some(&align), false, false, false, false);
        assert!(xf.contains("applyAlignment=\"1\""));
        assert!(xf.contains("horizontal=\"center\""));
        assert!(xf.contains("wrapText=\"1\""));
    }
}
