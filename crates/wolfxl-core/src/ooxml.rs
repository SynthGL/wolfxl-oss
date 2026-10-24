//! OOXML zip/XML helpers. Pure-Rust, no PyO3 coupling.
//!
//! These functions are shared between the styles walker (cellXfs + numFmts),
//! the worksheet-cell-style-id walker, and any future reader that needs to
//! poke at raw OOXML parts without going through `calamine-styles`.

use std::collections::HashMap;
use std::fs::File;
use std::io::Read;

use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader as XmlReader;
use zip::ZipArchive;

use crate::error::{Error, Result};

/// Read a single attribute off a `<tag ...>` start event, unescaping XML
/// entities. Falls back to a lossy UTF-8 conversion on unescape failure so
/// malformed workbooks still yield something rather than failing hard.
pub fn attr_value(e: &BytesStart<'_>, key: &[u8]) -> Option<String> {
    for a in e.attributes().with_checks(false).flatten() {
        if a.key.as_ref().as_bytes() == key {
            if let Ok(v) = quick_xml::escape::unescape(&a.value) {
                return Some(v.into_owned());
            }
            return Some(a.value.into_owned());
        }
    }
    None
}

/// Read a zip entry into a UTF-8 string, returning `None` when the entry
/// is missing. Any other zip/IO error propagates.
pub fn zip_read_to_string_opt(zip: &mut ZipArchive<File>, name: &str) -> Result<Option<String>> {
    match zip.by_name(name) {
        Ok(mut f) => {
            let mut out = String::new();
            f.read_to_string(&mut out)
                .map_err(|e| Error::Xlsx(format!("failed to read {name}: {e}")))?;
            Ok(Some(out))
        }
        Err(zip::result::ZipError::FileNotFound) => Ok(None),
        Err(e) => Err(Error::Xlsx(format!("zip error reading {name}: {e}"))),
    }
}

/// Read a zip entry that is required. Errors if missing.
pub fn zip_read_to_string(zip: &mut ZipArchive<File>, name: &str) -> Result<String> {
    zip_read_to_string_opt(zip, name)?
        .ok_or_else(|| Error::Xlsx(format!("missing zip entry: {name}")))
}

/// Parse `xl/workbook.xml` → list of `(sheetName, relationshipId)` pairs in
/// workbook order. Used with [`parse_relationship_targets`] to resolve each
/// sheet's worksheet XML path.
pub fn parse_workbook_sheet_rids(xml: &str) -> Result<Vec<(String, String)>> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf: Vec<u8> = Vec::new();
    let mut out: Vec<(String, String)> = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.local_name().as_ref() == "sheet" {
                    let name = attr_value(&e, b"name");
                    let rid = attr_value(&e, b"r:id");
                    if let (Some(n), Some(r)) = (name, rid) {
                        out.push((n, r));
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xlsx(format!("failed to parse workbook.xml: {e}"))),
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

/// Parse `xl/_rels/workbook.xml.rels` → map of relationship ID → target path.
/// Targets are returned as-is (usually relative, e.g. `worksheets/sheet1.xml`);
/// combine with [`join_and_normalize`] to get absolute zip-entry paths.
pub fn parse_relationship_targets(xml: &str) -> Result<HashMap<String, String>> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf: Vec<u8> = Vec::new();
    let mut out: HashMap<String, String> = HashMap::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.local_name().as_ref() == "Relationship" {
                    let id = attr_value(&e, b"Id");
                    let target = attr_value(&e, b"Target");
                    if let (Some(i), Some(t)) = (id, target) {
                        out.insert(i, t);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xlsx(format!("failed to parse rels: {e}"))),
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

/// Normalize a zip-entry path: collapse `.` / `..` segments and leading
/// slashes, returning a canonical `xl/...` form.
pub fn normalize_zip_path(path: &str) -> String {
    let mut stack: Vec<&str> = Vec::new();
    for part in path.split('/') {
        if part.is_empty() || part == "." {
            continue;
        }
        if part == ".." {
            stack.pop();
            continue;
        }
        stack.push(part);
    }
    stack.join("/")
}

/// Join a base directory (e.g. `xl/`) with a relationship target and
/// normalize the result. Handles absolute targets (`/xl/...`) and
/// already-rooted targets (`xl/...`) correctly.
pub fn join_and_normalize(base_dir: &str, target: &str) -> String {
    let t = target.trim_start_matches('/');
    let combined = if t.starts_with("xl/") {
        t.to_string()
    } else {
        format!("{base_dir}{t}")
    };
    normalize_zip_path(&combined)
}

/// Convert an A1-style cell reference (e.g. `"B7"`) to zero-based
/// `(row, col)` coordinates. Accepts mixed case and rejects `$`-anchored
/// references (strip them upstream if needed).
pub fn a1_to_row_col(a1: &str) -> std::result::Result<(u32, u32), String> {
    let mut col: u32 = 0;
    let mut row_digits = String::new();

    for ch in a1.chars() {
        if ch.is_ascii_alphabetic() {
            let uc = ch.to_ascii_uppercase() as u8;
            let val = (uc - b'A' + 1) as u32;
            col = col * 26 + val;
        } else if ch.is_ascii_digit() {
            row_digits.push(ch);
        } else {
            return Err(format!("Invalid cell reference: {a1}"));
        }
    }

    if col == 0 || row_digits.is_empty() {
        return Err(format!("Invalid cell reference: {a1}"));
    }

    let row_1: u32 = row_digits
        .parse()
        .map_err(|_| format!("Invalid cell reference: {a1}"))?;
    if row_1 == 0 {
        return Err(format!("Invalid cell reference: {a1}"));
    }

    Ok((row_1 - 1, col - 1))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a1_basics() {
        assert_eq!(a1_to_row_col("A1").unwrap(), (0, 0));
        assert_eq!(a1_to_row_col("B7").unwrap(), (6, 1));
        assert_eq!(a1_to_row_col("AA10").unwrap(), (9, 26));
        assert_eq!(a1_to_row_col("ZZ100").unwrap(), (99, 701));
    }

    #[test]
    fn a1_case_insensitive() {
        assert_eq!(a1_to_row_col("ab3").unwrap(), a1_to_row_col("AB3").unwrap());
    }

    #[test]
    fn a1_rejects_garbage() {
        // The parser scans letters and digits independently of order, so
        // "1A" resolves the same as "A1" — matching the existing PyO3
        // util behavior. The cases below are the hard rejections:
        // row 0 (openxml is 1-based), `$`-anchors, and empty input.
        assert!(a1_to_row_col("A0").is_err());
        assert!(a1_to_row_col("$A$1").is_err());
        assert!(a1_to_row_col("").is_err());
    }

    #[test]
    fn normalize_collapses_dots() {
        assert_eq!(
            normalize_zip_path("xl/./worksheets/../sheet1.xml"),
            "xl/sheet1.xml"
        );
        assert_eq!(
            normalize_zip_path("/xl/worksheets/sheet1.xml"),
            "xl/worksheets/sheet1.xml"
        );
    }

    #[test]
    fn join_handles_rooted_target() {
        assert_eq!(
            join_and_normalize("xl/", "worksheets/sheet1.xml"),
            "xl/worksheets/sheet1.xml"
        );
        assert_eq!(
            join_and_normalize("xl/", "/xl/worksheets/sheet1.xml"),
            "xl/worksheets/sheet1.xml"
        );
        assert_eq!(
            join_and_normalize("xl/", "xl/worksheets/sheet1.xml"),
            "xl/worksheets/sheet1.xml"
        );
    }

    #[test]
    fn parse_workbook_sheets() {
        let xml = r#"<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Revenue" sheetId="1" r:id="rId1"/>
    <sheet name="Balance Sheet" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>"#;
        let out = parse_workbook_sheet_rids(xml).unwrap();
        assert_eq!(out.len(), 2);
        assert_eq!(out[0], ("Revenue".to_string(), "rId1".to_string()));
        assert_eq!(out[1], ("Balance Sheet".to_string(), "rId2".to_string()));
    }

    #[test]
    fn parse_rels_happy() {
        let xml = r#"<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Target="worksheets/sheet2.xml" Type="..."/>
</Relationships>"#;
        let out = parse_relationship_targets(xml).unwrap();
        assert_eq!(out.get("rId1"), Some(&"worksheets/sheet1.xml".to_string()));
        assert_eq!(out.get("rId2"), Some(&"worksheets/sheet2.xml".to_string()));
    }
}
