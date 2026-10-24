//! Data-validation block builder for modify mode (RFC-025).
//!
//! The patcher path needs to read any existing `<dataValidations>` element
//! out of a sheet's XML, append zero or more new `<dataValidation>` rules,
//! and hand the combined block to RFC-011's merger as
//! `SheetBlock::DataValidations`. This module owns that responsibility.
//!
//! Existing rules are preserved by **byte-slice capture**: each child
//! `<dataValidation>...</dataValidation>` element from the source block is
//! copied verbatim into the output. We never round-trip them through a
//! parsed model, so attributes wolfxl doesn't recognize, escaped content
//! (`&amp;`, `&lt;`), self-closing forms, and CDATA all flow through
//! unchanged. Only newly added rules go through a serializer.
//!
//! The new-rule serializer mirrors `crates/wolfxl-writer/src/emit/sheet_xml.rs`
//! `emit_data_validations` so write-mode and modify-mode emit byte-identical
//! XML for the same `DataValidationPatch` inputs. RFC-025 §4.2 explicitly
//! authorizes a parallel implementation here rather than refactoring the
//! native writer to share — the patcher's input shape (a flat dict) and
//! the writer's input shape (its in-memory model) don't compose without an
//! intermediate type that adds complexity, and the writer crate is
//! PyO3-free while the patcher is PyO3-heavy, so the dep direction is
//! wrong. If drift surfaces, extract a shared `wolfxl-validations` crate.

use quick_xml::events::Event;
use quick_xml::Reader as XmlReader;

// ---------------------------------------------------------------------------
// Public types
// ---------------------------------------------------------------------------

/// One queued data-validation rule. Mirrors the OOXML §18.3.1.32
/// `CT_DataValidation` shape, plus `sqref` (which is required in our use).
#[derive(Debug, Clone, Default)]
pub struct DataValidationPatch {
    /// `"list"`, `"whole"`, `"decimal"`, `"date"`, `"time"`, `"textLength"`,
    /// `"custom"`, or `"none"`. Required.
    pub validation_type: String,
    /// `"between"`, `"notBetween"`, `"equal"`, `"notEqual"`, `"lessThan"`,
    /// `"lessThanOrEqual"`, `"greaterThan"`, `"greaterThanOrEqual"`. Omitted
    /// for `list` and `custom` even if supplied.
    pub operator: Option<String>,
    /// Primary formula (or value). For `type="list"` with inline values this
    /// is `"\"A,B,C\""` (note the embedded quotes); for a range reference
    /// it's the range string.
    pub formula1: Option<String>,
    /// Secondary formula. Only emitted when operator is `between` or
    /// `notBetween`.
    pub formula2: Option<String>,
    /// Space-separated A1 ranges. Required.
    pub sqref: String,
    pub allow_blank: bool,
    /// OOXML's `showDropDown="1"` means **HIDE** the dropdown
    /// (counterintuitive). We pass through whatever the caller supplies.
    pub show_dropdown: bool,
    pub show_input_message: bool,
    pub show_error_message: bool,
    /// `"warning"` or `"information"`. `None` (or `"stop"`) means default,
    /// which is omitted from the output per OOXML convention.
    pub error_style: Option<String>,
    pub error_title: Option<String>,
    pub error: Option<String>,
    pub prompt_title: Option<String>,
    pub prompt: Option<String>,
}

// ---------------------------------------------------------------------------
// extract_existing_dv_block
// ---------------------------------------------------------------------------

/// Locate the main worksheet `<dataValidations>` element in a sheet XML and return its
/// raw byte range (everything from `<dataValidations` through
/// `</dataValidations>` inclusive, or the self-closing form if applicable).
///
/// Returns `None` if no such block exists. This is what the merger needs
/// to decide whether to preserve an existing block or only emit new rules.
/// Extension payloads such as `<x14:dataValidations>` are deliberately ignored:
/// they are not valid children for the main `<dataValidations>` wrapper and
/// may carry namespace declarations from an enclosing `<ext>` element.
pub fn extract_existing_dv_block(sheet_xml: &str) -> Option<Vec<u8>> {
    let bytes = sheet_xml.as_bytes();
    let mut reader = XmlReader::from_str(sheet_xml);
    reader.config_mut().trim_text(false);
    let mut buf: Vec<u8> = Vec::new();

    let mut start_pos: Option<usize> = None;
    let mut depth: u32 = 0;

    loop {
        // Snapshot position BEFORE reading the next event. quick-xml's
        // `buffer_position()` returns the offset of the byte AFTER the last
        // event consumed, which is the start of the next event.
        let pre = reader.buffer_position() as usize;

        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                if is_main_data_validations_name(e.name().as_ref()) && start_pos.is_none() {
                    start_pos = Some(pre);
                    depth = 1;
                } else if start_pos.is_some() && is_main_data_validations_name(e.name().as_ref()) {
                    // Pathological nested case — shouldn't happen, but be defensive.
                    depth += 1;
                }
            }
            Ok(Event::Empty(ref e)) => {
                if is_main_data_validations_name(e.name().as_ref()) && start_pos.is_none() {
                    // Self-closing `<dataValidations/>` — capture the single tag.
                    let end = reader.buffer_position() as usize;
                    return Some(bytes[pre..end].to_vec());
                }
            }
            Ok(Event::End(ref e)) => {
                if is_main_data_validations_name(e.name().as_ref()) {
                    if depth > 0 {
                        depth -= 1;
                    }
                    if depth == 0 {
                        let start = start_pos.expect("end without start");
                        let end = reader.buffer_position() as usize;
                        return Some(bytes[start..end].to_vec());
                    }
                }
            }
            Ok(Event::Eof) => return None,
            Err(_) => return None,
            _ => {}
        }
        buf.clear();
    }
}

fn is_main_data_validations_name(name: &str) -> bool {
    name == "dataValidations"
}

// ---------------------------------------------------------------------------
// build_data_validations_block
// ---------------------------------------------------------------------------

/// Build a complete `<dataValidations count="N">…</dataValidations>` block
/// from any existing block plus the supplied new patches.
///
/// Existing `<dataValidation>` children are preserved verbatim (byte-slice
/// copy). New patches are serialized via [`serialize_patch`]. The `count`
/// attribute always reflects the total child count.
pub fn build_data_validations_block(
    existing_block: Option<&[u8]>,
    patches: &[DataValidationPatch],
) -> Vec<u8> {
    let existing_children: Vec<Vec<u8>> =
        existing_block.map(extract_dv_children).unwrap_or_default();

    let total = existing_children.len() + patches.len();

    let mut out: Vec<u8> = Vec::with_capacity(256);
    out.extend_from_slice(format!("<dataValidations count=\"{}\">", total).as_bytes());

    for child in &existing_children {
        out.extend_from_slice(child);
    }
    for p in patches {
        serialize_patch(&mut out, p);
    }

    out.extend_from_slice(b"</dataValidations>");
    out
}

// ---------------------------------------------------------------------------
// extract_dv_children — pull each <dataValidation> child as opaque bytes
// ---------------------------------------------------------------------------

/// Walk an existing `<dataValidations>` block (as captured by
/// [`extract_existing_dv_block`]) and return each `<dataValidation>` child
/// as a verbatim byte slice. Self-closing children are kept as-is.
fn extract_dv_children(block: &[u8]) -> Vec<Vec<u8>> {
    let s = match std::str::from_utf8(block) {
        Ok(s) => s,
        Err(_) => return Vec::new(),
    };

    let mut reader = XmlReader::from_str(s);
    reader.config_mut().trim_text(false);
    let mut buf: Vec<u8> = Vec::new();

    let mut children: Vec<Vec<u8>> = Vec::new();
    let mut child_start: Option<usize> = None;
    let mut child_depth: u32 = 0;

    loop {
        let pre = reader.buffer_position() as usize;

        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                if e.local_name().as_ref() == "dataValidation" {
                    if child_start.is_none() {
                        child_start = Some(pre);
                        child_depth = 1;
                    } else {
                        child_depth += 1;
                    }
                }
            }
            Ok(Event::Empty(ref e)) => {
                if e.local_name().as_ref() == "dataValidation" && child_start.is_none() {
                    let end = reader.buffer_position() as usize;
                    children.push(block[pre..end].to_vec());
                }
            }
            Ok(Event::End(ref e)) => {
                if e.local_name().as_ref() == "dataValidation" && child_start.is_some() {
                    child_depth -= 1;
                    if child_depth == 0 {
                        let start = child_start.take().expect("end without start");
                        let end = reader.buffer_position() as usize;
                        children.push(block[start..end].to_vec());
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    children
}

// ---------------------------------------------------------------------------
// serialize_patch — emit a single new <dataValidation>
// ---------------------------------------------------------------------------

/// Append one `<dataValidation>` element to `out`, mirroring the attribute
/// order and conditional emission of
/// `crates/wolfxl-writer/src/emit/sheet_xml.rs::emit_data_validations`.
fn serialize_patch(out: &mut Vec<u8>, p: &DataValidationPatch) {
    let type_str = if p.validation_type.is_empty() {
        "none"
    } else {
        p.validation_type.as_str()
    };

    out.extend_from_slice(format!("<dataValidation type=\"{}\"", type_str).as_bytes());

    // operator — omitted for list and custom (matches writer §18.3.1.32 rule)
    let needs_operator = !matches!(type_str, "list" | "custom" | "none" | "any");
    if needs_operator {
        if let Some(ref op) = p.operator {
            out.extend_from_slice(format!(" operator=\"{}\"", op).as_bytes());
        }
    }

    if p.allow_blank {
        out.extend_from_slice(b" allowBlank=\"1\"");
    }
    if p.show_dropdown {
        out.extend_from_slice(b" showDropDown=\"1\"");
    }
    if p.show_input_message {
        out.extend_from_slice(b" showInputMessage=\"1\"");
    }
    if p.show_error_message {
        out.extend_from_slice(b" showErrorMessage=\"1\"");
    }

    // errorStyle — only emit when not the default "stop"
    if let Some(ref es) = p.error_style {
        match es.as_str() {
            "stop" | "" => {} // default — omit
            other => {
                out.extend_from_slice(format!(" errorStyle=\"{}\"", other).as_bytes());
            }
        }
    }

    if let Some(ref t) = p.error_title {
        out.extend_from_slice(format!(" errorTitle=\"{}\"", attr_escape(t)).as_bytes());
    }
    if let Some(ref m) = p.error {
        out.extend_from_slice(format!(" error=\"{}\"", attr_escape(m)).as_bytes());
    }
    if let Some(ref t) = p.prompt_title {
        out.extend_from_slice(format!(" promptTitle=\"{}\"", attr_escape(t)).as_bytes());
    }
    if let Some(ref m) = p.prompt {
        out.extend_from_slice(format!(" prompt=\"{}\"", attr_escape(m)).as_bytes());
    }

    out.extend_from_slice(format!(" sqref=\"{}\">", attr_escape(&p.sqref)).as_bytes());

    if let Some(ref f1) = p.formula1 {
        out.extend_from_slice(format!("<formula1>{}</formula1>", text_escape(f1)).as_bytes());
    }

    let is_between = matches!(p.operator.as_deref(), Some("between") | Some("notBetween"));
    if is_between {
        if let Some(ref f2) = p.formula2 {
            out.extend_from_slice(format!("<formula2>{}</formula2>", text_escape(f2)).as_bytes());
        }
    }

    out.extend_from_slice(b"</dataValidation>");
}

/// XML attribute escaping. Matches the writer's `xml_escape::attr` rules so
/// new patches are byte-identical to the writer's output for the same input.
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

/// XML text escaping. Less strict than `attr_escape` — quotes are allowed
/// in text content (which matters for `<formula1>"A,B,C"</formula1>`).
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
    fn validation_merge_preserves_split_reference_text_and_cdata() {
        let child = r#"<dataValidation type="custom" sqref="A1"><formula1> A1 &lt; &#50; &amp; " &custom; " <![CDATA[ & raw ]]></formula1></dataValidation>"#;
        let block = format!(r#"<dataValidations count="1">{child}</dataValidations>"#);
        let sheet = format!("<worksheet><sheetData/>{block}</worksheet>");
        let extracted = extract_existing_dv_block(&sheet).unwrap();
        assert_eq!(extracted, block.as_bytes());
        let out = build_data_validations_block(Some(&extracted), &[list_patch()]);
        let text = std::str::from_utf8(&out).unwrap();
        assert!(text.starts_with(r#"<dataValidations count="2">"#));
        assert!(text.contains(child));
    }

    fn list_patch() -> DataValidationPatch {
        DataValidationPatch {
            validation_type: "list".to_string(),
            sqref: "B2:B100".to_string(),
            formula1: Some("\"Apple,Banana,Cherry\"".to_string()),
            allow_blank: true,
            show_error_message: true,
            ..Default::default()
        }
    }

    fn whole_between_patch() -> DataValidationPatch {
        DataValidationPatch {
            validation_type: "whole".to_string(),
            operator: Some("between".to_string()),
            sqref: "C2:C100".to_string(),
            formula1: Some("1".to_string()),
            formula2: Some("100".to_string()),
            show_input_message: true,
            show_error_message: true,
            error_title: Some("Invalid".to_string()),
            error: Some("Enter 1-100".to_string()),
            prompt_title: Some("Hint".to_string()),
            prompt: Some("Enter a number 1-100".to_string()),
            ..Default::default()
        }
    }

    // --- extract_existing_dv_block -----------------------------------------

    #[test]
    fn extract_returns_none_when_no_block() {
        let xml = r#"<?xml version="1.0"?><worksheet><sheetData/></worksheet>"#;
        assert!(extract_existing_dv_block(xml).is_none());
    }

    #[test]
    fn extract_captures_self_closing_block() {
        let xml = r#"<worksheet><sheetData/><dataValidations/></worksheet>"#;
        let got = extract_existing_dv_block(xml).expect("should find block");
        let s = std::str::from_utf8(&got).unwrap();
        // The self-closing tag should be captured verbatim. Range-capture of
        // a self-closing element returns just the tag itself.
        assert!(s.starts_with("<dataValidations"), "got: {s}");
        assert!(s.ends_with("/>"), "got: {s}");
    }

    #[test]
    fn extract_preserves_escaped_content() {
        let xml = r#"<worksheet><sheetData/><dataValidations count="1"><dataValidation type="custom" sqref="A1"><formula1>=A&gt;5</formula1></dataValidation></dataValidations></worksheet>"#;
        let got = extract_existing_dv_block(xml).expect("should find block");
        let s = std::str::from_utf8(&got).unwrap();
        assert!(s.contains("&gt;"), "escape preserved verbatim, got: {s}");
        assert!(s.starts_with("<dataValidations count=\"1\">"));
        assert!(s.ends_with("</dataValidations>"));
    }

    #[test]
    fn extract_captures_block_with_two_children() {
        let xml = r#"<worksheet><sheetData/><dataValidations count="2"><dataValidation type="list" sqref="A1"><formula1>"A,B"</formula1></dataValidation><dataValidation type="whole" operator="between" sqref="B1"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations></worksheet>"#;
        let got = extract_existing_dv_block(xml).expect("should find block");
        let s = std::str::from_utf8(&got).unwrap();
        assert!(s.starts_with("<dataValidations count=\"2\">"));
        assert!(s.contains("type=\"list\""));
        assert!(s.contains("type=\"whole\""));
        assert!(s.ends_with("</dataValidations>"));
    }

    #[test]
    fn extract_ignores_x14_extension_data_validations() {
        let xml = r#"<worksheet><sheetData/><extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:dataValidations count="1" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:dataValidation type="list"><x14:formula1><xm:f>LOV!$A$2:$A$15</xm:f></x14:formula1><xm:sqref>A2</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst></worksheet>"#;
        assert!(extract_existing_dv_block(xml).is_none());
    }

    #[test]
    fn extract_prefers_main_block_before_x14_extension_payload() {
        let xml = r#"<worksheet><sheetData/><dataValidations count="1"><dataValidation type="whole" sqref="A1"><formula1>1</formula1></dataValidation></dataValidations><extLst><ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:dataValidations count="1"/></ext></extLst></worksheet>"#;
        let got = extract_existing_dv_block(xml).expect("should find main block");
        let s = std::str::from_utf8(&got).unwrap();
        assert!(s.starts_with("<dataValidations count=\"1\">"));
        assert!(
            !s.contains("x14:dataValidation"),
            "got extension payload: {s}"
        );
    }

    // --- build_data_validations_block --------------------------------------

    #[test]
    fn build_with_no_existing_one_patch() {
        let p = list_patch();
        let out = build_data_validations_block(None, &[p]);
        let s = String::from_utf8(out).unwrap();
        assert!(s.starts_with("<dataValidations count=\"1\">"));
        assert!(s.contains("type=\"list\""));
        assert!(s.contains("sqref=\"B2:B100\""));
        assert!(s.contains("<formula1>\"Apple,Banana,Cherry\"</formula1>"));
        assert!(s.ends_with("</dataValidations>"));
    }

    #[test]
    fn build_with_existing_two_plus_one_new() {
        // Existing block has 2 DVs; we add one. count must be 3.
        let existing = br#"<dataValidations count="2"><dataValidation type="list" sqref="A1"><formula1>"X,Y"</formula1></dataValidation><dataValidation type="whole" operator="between" sqref="B1"><formula1>1</formula1><formula2>9</formula2></dataValidation></dataValidations>"#;
        let new = list_patch();
        let out = build_data_validations_block(Some(existing), &[new]);
        let s = String::from_utf8(out).unwrap();
        assert!(s.starts_with("<dataValidations count=\"3\">"), "got: {s}");
        // Existing children preserved verbatim — the X,Y inline list should still be there.
        assert!(s.contains("\"X,Y\""), "existing list preserved, got: {s}");
        assert!(
            s.contains("\"Apple,Banana,Cherry\""),
            "new list present, got: {s}"
        );
    }

    #[test]
    fn build_dv_list_inline_values() {
        let p = list_patch();
        let out = build_data_validations_block(None, &[p]);
        let s = String::from_utf8(out).unwrap();
        // Inline list MUST include embedded quotes verbatim.
        assert!(s.contains("<formula1>\"Apple,Banana,Cherry\"</formula1>"));
    }

    #[test]
    fn build_dv_list_range_ref() {
        let p = DataValidationPatch {
            validation_type: "list".to_string(),
            sqref: "B2:B100".to_string(),
            formula1: Some("Sheet2!$A$1:$A$5".to_string()),
            ..Default::default()
        };
        let out = build_data_validations_block(None, &[p]);
        let s = String::from_utf8(out).unwrap();
        assert!(s.contains("<formula1>Sheet2!$A$1:$A$5</formula1>"));
        // operator must NOT appear for list type
        assert!(
            !s.contains("operator="),
            "operator should be omitted for list, got: {s}"
        );
    }

    #[test]
    fn build_dv_whole_between_emits_formula2() {
        let p = whole_between_patch();
        let out = build_data_validations_block(None, &[p]);
        let s = String::from_utf8(out).unwrap();
        assert!(s.contains("operator=\"between\""));
        assert!(s.contains("<formula1>1</formula1>"));
        assert!(s.contains("<formula2>100</formula2>"));
        // Plus all the optional message attributes
        assert!(s.contains("errorTitle=\"Invalid\""));
        assert!(s.contains("error=\"Enter 1-100\""));
        assert!(s.contains("promptTitle=\"Hint\""));
    }

    #[test]
    fn build_dv_custom_omits_operator() {
        let p = DataValidationPatch {
            validation_type: "custom".to_string(),
            // Even if a caller supplies an operator (mistake), custom must omit it.
            operator: Some("between".to_string()),
            sqref: "A1".to_string(),
            formula1: Some("=LEN(A1)>5".to_string()),
            ..Default::default()
        };
        let out = build_data_validations_block(None, &[p]);
        let s = String::from_utf8(out).unwrap();
        assert!(s.contains("type=\"custom\""));
        assert!(
            !s.contains("operator="),
            "custom must not emit operator, got: {s}"
        );
        // > inside text content must be escaped
        assert!(s.contains("<formula1>=LEN(A1)&gt;5</formula1>"));
    }

    #[test]
    fn build_dv_error_style_warning_appears_default_stop_omitted() {
        let warn = DataValidationPatch {
            validation_type: "whole".to_string(),
            operator: Some("equal".to_string()),
            sqref: "A1".to_string(),
            formula1: Some("5".to_string()),
            error_style: Some("warning".to_string()),
            ..Default::default()
        };
        let s = String::from_utf8(build_data_validations_block(None, &[warn])).unwrap();
        assert!(s.contains("errorStyle=\"warning\""));

        let stop = DataValidationPatch {
            validation_type: "whole".to_string(),
            operator: Some("equal".to_string()),
            sqref: "A1".to_string(),
            formula1: Some("5".to_string()),
            error_style: Some("stop".to_string()),
            ..Default::default()
        };
        let s = String::from_utf8(build_data_validations_block(None, &[stop])).unwrap();
        assert!(
            !s.contains("errorStyle="),
            "default stop must be omitted, got: {s}"
        );
    }

    #[test]
    fn build_count_attribute_matches_total_children() {
        let existing = br#"<dataValidations count="2"><dataValidation type="list" sqref="A1"/><dataValidation type="list" sqref="B1"/></dataValidations>"#;
        let p1 = list_patch();
        let p2 = whole_between_patch();
        let s = String::from_utf8(build_data_validations_block(Some(existing), &[p1, p2])).unwrap();
        // 2 existing + 2 new = 4
        assert!(s.starts_with("<dataValidations count=\"4\">"), "got: {s}");
        // sanity: the 4 child tags all appear
        let child_count = s.matches("<dataValidation ").count();
        assert_eq!(child_count, 4, "expected 4 child tags in: {s}");
    }

    #[test]
    fn build_preserves_existing_with_escaped_content_byte_for_byte() {
        // Headline correctness gate (RFC §8 risk #1).
        let existing = br#"<dataValidations count="1"><dataValidation type="custom" sqref="A1"><formula1>=A&gt;5</formula1></dataValidation></dataValidations>"#;
        let p = list_patch();
        let s = String::from_utf8(build_data_validations_block(Some(existing), &[p])).unwrap();
        // The escape MUST flow through verbatim — no over-escaping (&amp;gt;) and no
        // un-escaping (>) introduced by a parse-and-re-emit roundtrip.
        assert!(s.contains("=A&gt;5"), "verbatim escape preserved, got: {s}");
        assert!(!s.contains("&amp;gt;"), "no double-escape, got: {s}");
    }

    #[test]
    fn build_dv_empty_patches_with_existing_passes_through() {
        // No new patches, but an existing block — output should re-emit the
        // existing children with the same count.
        let existing = br#"<dataValidations count="1"><dataValidation type="list" sqref="A1"><formula1>"X,Y"</formula1></dataValidation></dataValidations>"#;
        let s = String::from_utf8(build_data_validations_block(Some(existing), &[])).unwrap();
        assert_eq!(s, "<dataValidations count=\"1\"><dataValidation type=\"list\" sqref=\"A1\"><formula1>\"X,Y\"</formula1></dataValidation></dataValidations>");
    }
}
