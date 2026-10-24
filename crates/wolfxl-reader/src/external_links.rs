//! Pure-Rust external-link part + rels parsers (RFC-071 / G18).
//!
//! Used by the PyO3 bridge in `src/wolfxl/external_links.rs` and unit-
//! testable without an embedded interpreter.
//!
//! The parsers are deliberately small: we only need the fields that the
//! Python `ExternalLink` dataclass exposes (RFC-071 §3). Anything else in
//! the parts is preserved by the patcher's byte-level passthrough on save
//! — we never round-trip through these structs.

use quick_xml::events::Event;
use quick_xml::Reader as XmlReader;

use wolfxl_rels::{rt, RelsGraph, TargetMode};

/// A single cached cell value from `<sheetDataSet>/<sheetData>/<row>/<cell>`.
#[derive(Debug, Clone, PartialEq)]
pub struct CachedCell {
    pub r#ref: String,
    pub value: String,
}

/// Parsed shape of `xl/externalLinks/externalLink{N}.xml`.
#[derive(Debug, Default, Clone, PartialEq)]
pub struct ExternalLinkPart {
    /// `r:id` on `<externalBook>` — points into the part's sibling rels file.
    pub book_rid: Option<String>,
    /// Sheet names referenced by formulas in the linking workbook.
    pub sheet_names: Vec<String>,
    /// One entry per `<sheetData sheetId="N">` block, in document order.
    pub cached_data: Vec<(String, Vec<CachedCell>)>,
}

/// Parsed shape of `xl/externalLinks/_rels/externalLink{N}.xml.rels`.
#[derive(Debug, Default, Clone, PartialEq)]
pub struct ExternalLinkRels {
    /// `Target` attribute of the first `externalLinkPath` rel.
    pub target: Option<String>,
    /// `External` for normal external workbook refs.
    pub target_mode: Option<TargetMode>,
    /// Relationship id (`rId1`) of the `externalLinkPath` rel.
    pub rid: Option<String>,
}

/// Parse `xl/externalLinks/externalLink{N}.xml`. Returns a `String` error
/// message on malformed XML so the PyO3 wrapper can wrap it cleanly.
pub fn parse_part(xml: &[u8]) -> Result<ExternalLinkPart, String> {
    let mut reader = XmlReader::from_reader(xml);

    let mut out = ExternalLinkPart::default();
    let mut current_sheet_id: Option<String> = None;
    let mut current_cells: Vec<CachedCell> = Vec::new();
    let mut current_ref: Option<String> = None;
    let mut value_buf = String::new();
    let mut in_value = false;

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                let name = e.name();
                let local = local_name(name.as_ref().as_bytes());
                match local {
                    b"externalBook" => {
                        for attr in e.attributes().flatten() {
                            let key = attr.key.as_ref().as_bytes();
                            if local_name(key) == b"id" && key.iter().any(|&c| c == b':') {
                                if let Ok(v) = quick_xml::escape::unescape(&attr.value) {
                                    out.book_rid = Some(v.into_owned());
                                }
                            }
                        }
                    }
                    b"sheetName" => {
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == "val" {
                                if let Ok(v) = quick_xml::escape::unescape(&attr.value) {
                                    out.sheet_names.push(v.into_owned());
                                }
                            }
                        }
                    }
                    b"sheetData" => {
                        current_sheet_id = None;
                        current_cells = Vec::new();
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == "sheetId" {
                                if let Ok(v) = quick_xml::escape::unescape(&attr.value) {
                                    current_sheet_id = Some(v.into_owned());
                                }
                            }
                        }
                    }
                    b"cell" => {
                        current_ref = None;
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == "r" {
                                if let Ok(v) = quick_xml::escape::unescape(&attr.value) {
                                    current_ref = Some(v.into_owned());
                                }
                            }
                        }
                    }
                    b"v" => {
                        in_value = true;
                        value_buf.clear();
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(ref e)) => {
                if in_value {
                    value_buf.push_str(e);
                }
            }
            Ok(Event::GeneralRef(e)) => {
                if in_value {
                    value_buf.push_str(
                        crate::resolve_xml_reference(&e, &mut [0; 4])
                            .map_err(|err| format!("external link XML text: {err}"))?,
                    );
                }
            }
            Ok(Event::CData(e)) => {
                if in_value {
                    value_buf.push_str(&e);
                }
            }
            Ok(Event::End(ref e)) => {
                let name = e.name();
                let local = local_name(name.as_ref().as_bytes());
                match local {
                    b"v" => in_value = false,
                    b"cell" => {
                        if let Some(r) = current_ref.take() {
                            current_cells.push(CachedCell {
                                r#ref: r,
                                value: std::mem::take(&mut value_buf),
                            });
                        } else {
                            value_buf.clear();
                        }
                    }
                    b"sheetData" => {
                        let sid = current_sheet_id.take().unwrap_or_default();
                        out.cached_data
                            .push((sid, std::mem::take(&mut current_cells)));
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(format!("external link XML parse error: {e}")),
            _ => {}
        }
        buf.clear();
    }
    Ok(out)
}

/// Parse `xl/externalLinks/_rels/externalLink{N}.xml.rels`. The first
/// `externalLinkPath` rel wins (Excel only emits one in practice).
pub fn parse_rels(xml: &[u8]) -> Result<ExternalLinkRels, String> {
    let graph =
        RelsGraph::parse(xml).map_err(|e| format!("external link rels parse error: {e}"))?;
    let mut out = ExternalLinkRels::default();
    for r in graph.iter() {
        if is_external_link_path_rel(&r.rel_type) {
            out.target = Some(r.target.clone());
            out.target_mode = Some(r.mode);
            out.rid = Some(r.id.0.clone());
            break;
        }
    }
    Ok(out)
}

fn is_external_link_path_rel(rel_type: &str) -> bool {
    rel_type == rt::EXTERNAL_LINK_PATH || rel_type.starts_with(rt::MS_EXTERNAL_LINK_PATH_PREFIX)
}

fn local_name(name: &[u8]) -> &[u8] {
    match name.iter().rposition(|&c| c == b':') {
        Some(i) => &name[i + 1..],
        None => name,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const STD_PART: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <externalBook r:id="rId1">
    <sheetNames>
      <sheetName val="Sheet1"/>
      <sheetName val="Sheet2"/>
    </sheetNames>
    <sheetDataSet>
      <sheetData sheetId="0">
        <row r="1">
          <cell r="A1"><v>42</v></cell>
        </row>
      </sheetData>
    </sheetDataSet>
  </externalBook>
</externalLink>"#;

    #[test]
    fn parse_part_basic() {
        let p = parse_part(STD_PART).unwrap();
        assert_eq!(p.book_rid.as_deref(), Some("rId1"));
        assert_eq!(p.sheet_names, vec!["Sheet1", "Sheet2"]);
        assert_eq!(p.cached_data.len(), 1);
        let (sid, cells) = &p.cached_data[0];
        assert_eq!(sid, "0");
        assert_eq!(cells.len(), 1);
        assert_eq!(cells[0].r#ref, "A1");
        assert_eq!(cells[0].value, "42");
    }

    #[test]
    fn parse_part_empty_sheet_names() {
        let xml = br#"<externalLink xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><externalBook r:id="rId7"><sheetNames/></externalBook></externalLink>"#;
        let p = parse_part(xml).unwrap();
        assert_eq!(p.book_rid.as_deref(), Some("rId7"));
        assert!(p.sheet_names.is_empty());
        assert!(p.cached_data.is_empty());
    }

    #[test]
    fn parse_part_no_external_book() {
        // Pathological but well-formed: the part is empty.
        let xml =
            br#"<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;
        let p = parse_part(xml).unwrap();
        assert!(p.book_rid.is_none());
        assert!(p.sheet_names.is_empty());
    }

    #[test]
    fn parse_part_malformed_rejects() {
        let xml = b"<externalLink><externalBook";
        assert!(parse_part(xml).is_err());
    }

    #[test]
    fn cached_values_resolve_split_references_once() {
        let xml = br#"<externalLink xmlns:r="urn:rels"><externalBook r:id="r&amp;1">
          <sheetNames><sheetName val="A &amp; B"/></sheetNames>
          <sheetDataSet><sheetData sheetId="0"><row r="1">
            <cell r="A1"><v> left &amp; right &#x1D11E; &amp;lt;<![CDATA[ <raw>]]></v></cell>
          </row></sheetData></sheetDataSet>
        </externalBook></externalLink>"#;
        let part = parse_part(xml).unwrap();
        assert_eq!(part.book_rid.as_deref(), Some("r&1"));
        assert_eq!(part.sheet_names, ["A & B"]);
        assert_eq!(
            part.cached_data[0].1[0].value,
            " left & right \u{1D11E} &lt; <raw>"
        );
        assert!(parse_part(b"<externalLink><v>&unknown;</v></externalLink>").is_err());
    }

    #[test]
    fn parse_rels_basic() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath"
    Target="ext.xlsx" TargetMode="External"/>
</Relationships>"#;
        let r = parse_rels(xml).unwrap();
        assert_eq!(r.target.as_deref(), Some("ext.xlsx"));
        assert_eq!(r.target_mode, Some(TargetMode::External));
        assert_eq!(r.rid.as_deref(), Some("rId1"));
    }

    #[test]
    fn parse_rels_no_external_link_path() {
        let xml = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#;
        let r = parse_rels(xml).unwrap();
        assert!(r.target.is_none());
        assert!(r.rid.is_none());
    }

    #[test]
    fn parse_rels_malformed_rejects() {
        let xml = b"<Relationships";
        assert!(parse_rels(xml).is_err());
    }
}
