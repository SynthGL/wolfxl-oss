//! Pure-Rust `xl/styles.xml` reader — cellXfs + numFmts.
//!
//! Reads the cellXfs table and custom numFmts directly from styles.xml,
//! then resolves a style id to its number-format code.
//!
//! Paired with [`crate::worksheet_xml::parse_cell_style_ids`] (which builds
//! the per-cell `(row, col) → styleId` map), this resolves styles without
//! relying on an external workbook reader.

use std::collections::HashMap;

use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader as XmlReader;

use crate::error::{Error, Result};
use crate::ooxml::attr_value;

/// A parsed `<xf>` entry from `<cellXfs>`. A cell's `s` attribute points into
/// this table; each entry's `num_fmt_id` then resolves via built-in table
/// or custom `numFmts` to a format-code string.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XfEntry {
    pub num_fmt_id: u32,
    pub font_id: u32,
    pub fill_id: u32,
    pub border_id: u32,
}

/// Parse the `<cellXfs>` section of styles.xml into an ordered list of
/// [`XfEntry`]s. The ordinal position matches the `s="N"` attribute on cells.
pub fn parse_cellxfs(xml: &str) -> Vec<XfEntry> {
    let mut reader = XmlReader::from_str(xml);
    let mut buf: Vec<u8> = Vec::new();
    let mut in_cellxfs = false;
    let mut entries: Vec<XfEntry> = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let tag = e.local_name();
                if tag.as_ref() == "cellXfs" {
                    in_cellxfs = true;
                } else if tag.as_ref() == "xf" && in_cellxfs {
                    entries.push(parse_xf_entry(e));
                }
            }
            Ok(Event::Empty(ref e)) => {
                if e.local_name().as_ref() == "xf" && in_cellxfs {
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

/// Parse the `<numFmts>` section into `numFmtId → formatCode`. Custom
/// formats always live here; built-in ones (IDs < 164) are resolved via
/// [`builtin_num_fmt`] instead.
pub fn parse_num_fmts(xml: &str) -> Result<HashMap<u32, String>> {
    let mut reader = XmlReader::from_str(xml);
    let mut buf: Vec<u8> = Vec::new();

    let mut in_numfmts = false;
    let mut formats: HashMap<u32, String> = HashMap::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if e.local_name().as_ref() == "numFmts" {
                    in_numfmts = true;
                } else if in_numfmts && e.local_name().as_ref() == "numFmt" {
                    capture_num_fmt(&e, &mut formats);
                }
            }
            Ok(Event::Empty(e)) => {
                if in_numfmts && e.local_name().as_ref() == "numFmt" {
                    capture_num_fmt(&e, &mut formats);
                }
            }
            Ok(Event::End(e)) => {
                if e.local_name().as_ref() == "numFmts" {
                    in_numfmts = false;
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xlsx(format!("failed to parse styles.xml: {e}"))),
            _ => {}
        }
        buf.clear();
    }

    Ok(formats)
}

fn capture_num_fmt(e: &BytesStart<'_>, out: &mut HashMap<u32, String>) {
    let id = attr_value(e, b"numFmtId").and_then(|s| s.parse::<u32>().ok());
    let code = attr_value(e, b"formatCode");
    if let (Some(id), Some(code)) = (id, code) {
        out.insert(id, code);
    }
}

/// Excel's reserved built-in number-format codes. IDs 0..163 are reserved;
/// 164+ are always custom and live in `<numFmts>`. Only the IDs Excel
/// actually uses are listed; missing slots have no built-in meaning.
///
/// Table mirrors openpyxl's `openpyxl.styles.numbers.BUILTIN_FORMATS` so
/// we converge on the same string a host tool would display.
pub const BUILTIN_NUM_FMTS: &[(u32, &str)] = &[
    (0, "General"),
    (1, "0"),
    (2, "0.00"),
    (3, "#,##0"),
    (4, "#,##0.00"),
    (5, "\"$\"#,##0_);(\"$\"#,##0)"),
    (6, "\"$\"#,##0_);[Red](\"$\"#,##0)"),
    (7, "\"$\"#,##0.00_);(\"$\"#,##0.00)"),
    (8, "\"$\"#,##0.00_);[Red](\"$\"#,##0.00)"),
    (9, "0%"),
    (10, "0.00%"),
    (11, "0.00E+00"),
    (12, "# ?/?"),
    (13, "# ??/??"),
    (14, "mm-dd-yy"),
    (15, "d-mmm-yy"),
    (16, "d-mmm"),
    (17, "mmm-yy"),
    (18, "h:mm AM/PM"),
    (19, "h:mm:ss AM/PM"),
    (20, "h:mm"),
    (21, "h:mm:ss"),
    (22, "m/d/yy h:mm"),
    (37, "#,##0_);(#,##0)"),
    (38, "#,##0_);[Red](#,##0)"),
    (39, "#,##0.00_);(#,##0.00)"),
    (40, "#,##0.00_);[Red](#,##0.00)"),
    (41, r#"_(* #,##0_);_(* \(#,##0\);_(* "-"_);_(@_)"#),
    (42, r#"_("$"* #,##0_);_("$"* \(#,##0\);_("$"* "-"_);_(@_)"#),
    (43, r#"_(* #,##0.00_);_(* \(#,##0.00\);_(* "-"??_);_(@_)"#),
    (
        44,
        r#"_("$"* #,##0.00_)_("$"* \(#,##0.00\)_("$"* "-"??_)_(@_)"#,
    ),
    (45, "mm:ss"),
    (46, "[h]:mm:ss"),
    (47, "mmss.0"),
    (48, "##0.0E+0"),
    (49, "@"),
];

/// Resolve a built-in numFmtId to its format-code string, or `None` if the
/// ID isn't a known built-in. `0 → "General"` is returned as-is; callers who
/// treat General as "no format" must filter it out themselves.
pub fn builtin_num_fmt(id: u32) -> Option<&'static str> {
    BUILTIN_NUM_FMTS
        .iter()
        .find_map(|(i, code)| if *i == id { Some(*code) } else { None })
}

/// Resolve a numFmtId against both the custom table and the built-in list.
/// Custom entries win on conflict (Excel itself uses the custom value when
/// an ID that overlaps with a built-in is redefined).
pub fn resolve_num_fmt<'a>(id: u32, customs: &'a HashMap<u32, String>) -> Option<&'a str> {
    if let Some(custom) = customs.get(&id) {
        return Some(custom.as_str());
    }
    builtin_num_fmt(id)
}

#[cfg(test)]
mod tests {
    use super::*;

    const MINIMAL_STYLES: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="2">
  <numFmt numFmtId="164" formatCode="&quot;$&quot;#,##0.00"/>
  <numFmt numFmtId="165" formatCode="0.0%"/>
</numFmts>
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="1"><fill><patternFill patternType="none"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellXfs count="3">
  <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  <xf numFmtId="164" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
  <xf numFmtId="9" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
</cellXfs>
</styleSheet>"#;

    #[test]
    fn parse_cellxfs_returns_entries_in_order() {
        let entries = parse_cellxfs(MINIMAL_STYLES);
        assert_eq!(entries.len(), 3);
        assert_eq!(entries[0].num_fmt_id, 0);
        assert_eq!(entries[1].num_fmt_id, 164);
        assert_eq!(entries[2].num_fmt_id, 9);
    }

    #[test]
    fn self_closing_cellxfs_does_not_capture_later_xfs() {
        let xml = r#"
<styleSheet>
  <cellXfs count="0"/>
  <cellStyleXfs count="1"><xf numFmtId="164"/></cellStyleXfs>
</styleSheet>"#;
        let entries = parse_cellxfs(xml);
        assert!(entries.is_empty());
    }

    #[test]
    fn parse_num_fmts_captures_custom_formats() {
        let customs = parse_num_fmts(MINIMAL_STYLES).unwrap();
        assert_eq!(customs.get(&164).map(|s| s.as_str()), Some("\"$\"#,##0.00"));
        assert_eq!(customs.get(&165).map(|s| s.as_str()), Some("0.0%"));
    }

    #[test]
    fn prefixed_styles_decode_attribute_entities() {
        let xml = r#"<s:styleSheet xmlns:s="urn:styles"><s:numFmts><s:numFmt numFmtId="16&#52;" formatCode="&quot;A&amp;B&lt;&#x3E;&quot;0"/></s:numFmts><s:cellXfs><s:xf numFmtId="16&#52;" fontId="&#49;"/></s:cellXfs></s:styleSheet>"#;
        let formats = parse_num_fmts(xml).unwrap();
        assert_eq!(formats.get(&164).map(String::as_str), Some("\"A&B<>\"0"));
        let entries = parse_cellxfs(xml);
        assert_eq!(entries.len(), 1);
        assert_eq!(entries[0].num_fmt_id, 164);
        assert_eq!(entries[0].font_id, 1);
    }

    #[test]
    fn parse_num_fmts_empty_when_no_section() {
        let xml = r#"<styleSheet><cellXfs count="1"><xf/></cellXfs></styleSheet>"#;
        let customs = parse_num_fmts(xml).unwrap();
        assert!(customs.is_empty());
    }

    #[test]
    fn builtin_num_fmt_covers_common_ids() {
        assert_eq!(builtin_num_fmt(0), Some("General"));
        assert_eq!(builtin_num_fmt(9), Some("0%"));
        assert_eq!(builtin_num_fmt(14), Some("mm-dd-yy"));
        assert_eq!(
            builtin_num_fmt(44),
            Some(r#"_("$"* #,##0.00_)_("$"* \(#,##0.00\)_("$"* "-"??_)_(@_)"#)
        );
        assert_eq!(builtin_num_fmt(163), None);
    }

    #[test]
    fn resolve_prefers_custom_over_builtin() {
        let mut customs = HashMap::new();
        customs.insert(9, "0.0% (redefined)".to_string());
        assert_eq!(resolve_num_fmt(9, &customs), Some("0.0% (redefined)"));
    }

    #[test]
    fn resolve_falls_back_to_builtin() {
        let customs = HashMap::new();
        assert_eq!(resolve_num_fmt(14, &customs), Some("mm-dd-yy"));
        assert_eq!(resolve_num_fmt(999, &customs), None);
    }
}
