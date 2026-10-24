//! Worksheet XML walker — builds a per-cell `(row, col) → styleId` map.
//!
//! Each `<c r="A1" s="N">` element in a sheet's XML body carries a
//! reference into the workbook-level `<cellXfs>` table. Calamine-styles
//! exposes the resolved `Style` object for us when it can, but for
//! openpyxl-generated fixtures that path sometimes yields `None`. This
//! walker is the fallback: given the raw worksheet XML string, return a
//! map of populated cells with non-zero style ids so
//! [`crate::sheet::Sheet::load`] can look up their number-format codes
//! through [`crate::styles`] and recover the format.

use std::collections::HashMap;

use quick_xml::events::Event;
use quick_xml::Reader as XmlReader;

use crate::error::{Error, Result};
use crate::ooxml::{a1_to_row_col, attr_value};

/// Walk a worksheet's XML, returning only cells whose `s` attribute is
/// non-zero. Style-0 (the workbook default) is intentionally skipped — most
/// cells reference it and caching them all would balloon memory for large
/// sheets without affecting format-resolution correctness.
pub fn parse_cell_style_ids(xml: &str) -> Result<HashMap<(u32, u32), u32>> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf: Vec<u8> = Vec::new();
    let mut out: HashMap<(u32, u32), u32> = HashMap::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.local_name().as_ref() == "c" {
                    let a1 = attr_value(&e, b"r").unwrap_or_default();
                    if a1.is_empty() {
                        continue;
                    }
                    let style_id = attr_value(&e, b"s")
                        .and_then(|s| s.parse::<u32>().ok())
                        .unwrap_or(0);
                    if style_id == 0 {
                        continue;
                    }
                    if let Ok((row, col)) = a1_to_row_col(&a1) {
                        out.insert((row, col), style_id);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(Error::Xlsx(format!(
                    "failed to parse worksheet XML for style IDs: {e}"
                )))
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn walker_extracts_non_zero_style_ids() {
        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" s="0"><v>1</v></c><c r="B1" s="2"><v>2</v></c></row>
    <row r="2"><c r="A2" s="3"><v>10</v></c><c r="B2"><v>20</v></c></row>
  </sheetData>
</worksheet>"#;
        let out = parse_cell_style_ids(xml).unwrap();
        assert_eq!(out.len(), 2);
        assert_eq!(out.get(&(0, 1)), Some(&2)); // B1 style 2
        assert_eq!(out.get(&(1, 0)), Some(&3)); // A2 style 3
        assert!(!out.contains_key(&(0, 0))); // A1 style 0 skipped
        assert!(!out.contains_key(&(1, 1))); // B2 no s attr skipped
    }

    #[test]
    fn walker_handles_self_closing_cells() {
        // Cells with no value are self-closing <c .../>. They still count.
        let xml = r#"<worksheet><sheetData>
  <row r="1"><c r="A1" s="5"/><c r="B1" s="7"><v>1</v></c></row>
</sheetData></worksheet>"#;
        let out = parse_cell_style_ids(xml).unwrap();
        assert_eq!(out.get(&(0, 0)), Some(&5));
        assert_eq!(out.get(&(0, 1)), Some(&7));
    }

    #[test]
    fn walker_tolerates_missing_r_attr() {
        // A `<c>` without `r` can't be located on the grid — skip it rather
        // than bubbling an error, since calamine will have assigned it a
        // position based on document order anyway.
        let xml = r#"<worksheet><sheetData>
  <row><c s="3"><v>1</v></c></row>
</sheetData></worksheet>"#;
        let out = parse_cell_style_ids(xml).unwrap();
        assert!(out.is_empty());
    }

    #[test]
    fn walker_errors_on_broken_xml() {
        let xml = "<worksheet><sheetData><c r=\"A1\" s=\"unterminated";
        assert!(parse_cell_style_ids(xml).is_err());
    }
}
