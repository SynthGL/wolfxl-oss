//! Bounded merge metadata scans. Never construct a worksheet cell model.

use std::collections::{HashMap, HashSet};
use std::io::{BufRead, BufReader};

use quick_xml::events::Event;
use quick_xml::Reader;

use super::{attr_value, zip_from_source, NativeXlsxBook, ReaderError, Result};

pub(super) fn read_ranges(book: &NativeXlsxBook, sheet: &str) -> Result<Vec<String>> {
    let info = book
        .sheets
        .iter()
        .find(|s| s.name == sheet)
        .ok_or_else(|| ReaderError::SheetNotFound(sheet.to_string()))?;
    let mut zip = zip_from_source(&book.source)?;
    let part_name = zip
        .file_names()
        .find(|name| name.eq_ignore_ascii_case(&info.path))
        .map(str::to_string)
        .ok_or_else(|| ReaderError::MissingPart(info.path.clone()))?;
    let part = zip.by_name(&part_name)?;
    parse_ranges(BufReader::new(part))
}

fn parse_ranges<R: BufRead>(source: R) -> Result<Vec<String>> {
    let mut reader = Reader::from_reader(source);
    let mut buf = Vec::new();
    let mut ranges = Vec::new();
    loop {
        match reader
            .read_event_into(&mut buf)
            .map_err(|e| ReaderError::Xml(e.to_string()))?
        {
            Event::Start(ref e) | Event::Empty(ref e)
                if e.local_name().as_ref() == b"mergeCell" =>
            {
                if let Some(range) = attr_value(e, b"ref") {
                    ranges.push(range);
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(ranges)
}

pub(super) fn read_endpoint_styles(
    book: &NativeXlsxBook,
    sheet: &str,
    ranges: &[String],
) -> Result<HashMap<(u32, u32), u32>> {
    let targets: HashSet<(u32, u32)> = ranges
        .iter()
        .filter_map(|r| parse_bounds(r))
        .flat_map(|(r1, c1, r2, c2)| [(r1, c1), (r2, c2)])
        .collect();
    if targets.is_empty() {
        return Ok(HashMap::new());
    }
    let info = book
        .sheets
        .iter()
        .find(|s| s.name == sheet)
        .ok_or_else(|| ReaderError::SheetNotFound(sheet.to_string()))?;
    let mut zip = zip_from_source(&book.source)?;
    let part_name = zip
        .file_names()
        .find(|name| name.eq_ignore_ascii_case(&info.path))
        .map(str::to_string)
        .ok_or_else(|| ReaderError::MissingPart(info.path.clone()))?;
    let part = zip.by_name(&part_name)?;
    parse_endpoint_styles(BufReader::new(part), &targets)
}

fn parse_endpoint_styles<R: BufRead>(
    source: R,
    targets: &HashSet<(u32, u32)>,
) -> Result<HashMap<(u32, u32), u32>> {
    let mut reader = Reader::from_reader(source);
    let mut buf = Vec::new();
    let mut styles = HashMap::new();
    loop {
        match reader
            .read_event_into(&mut buf)
            .map_err(|e| ReaderError::Xml(e.to_string()))?
        {
            Event::Start(ref e) | Event::Empty(ref e) if e.local_name().as_ref() == b"c" => {
                if let Some(coord) = attr_value(e, b"r").and_then(|r| parse_cell(&r)) {
                    if targets.contains(&coord) {
                        let style = attr_value(e, b"s")
                            .and_then(|s| s.parse().ok())
                            .unwrap_or(0);
                        styles.insert(coord, style);
                    }
                }
            }
            Event::End(ref e) if e.local_name().as_ref() == b"sheetData" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(styles)
}

fn parse_cell(reference: &str) -> Option<(u32, u32)> {
    let mut col = 0u32;
    let mut row = 0u32;
    let mut digits = false;
    for byte in reference.bytes().filter(|b| *b != b'$') {
        if byte.is_ascii_alphabetic() && !digits {
            col = col
                .checked_mul(26)?
                .checked_add((byte.to_ascii_uppercase() - b'A' + 1) as u32)?;
        } else if byte.is_ascii_digit() {
            digits = true;
            row = row.checked_mul(10)?.checked_add((byte - b'0') as u32)?;
        } else {
            return None;
        }
    }
    (row > 0 && col > 0).then_some((row, col))
}

fn parse_bounds(reference: &str) -> Option<(u32, u32, u32, u32)> {
    let (start, end) = reference.split_once(':').unwrap_or((reference, reference));
    let (r1, c1) = parse_cell(start)?;
    let (r2, c2) = parse_cell(end)?;
    Some((r1, c1, r2, c2))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn namespaced_metadata_keeps_only_requested_endpoint_styles() {
        let xml = br#"<x:worksheet xmlns:x="urn:test"><x:sheetData><x:row r="1"><x:c r="A1" s="4"><x:v>1</x:v></x:c><x:c r="B1" s="9"/></x:row><x:row r="2"><x:c r="C2" s="7"/></x:row></x:sheetData><x:mergeCells><x:mergeCell ref="A1:C2"/></x:mergeCells></x:worksheet>"#;
        assert_eq!(parse_ranges(&xml[..]).unwrap(), vec!["A1:C2"]);
        let targets = HashSet::from([(1, 1), (2, 3)]);
        let styles = parse_endpoint_styles(&xml[..], &targets).unwrap();
        assert_eq!(styles, HashMap::from([((1, 1), 4), ((2, 3), 7)]));
    }

    #[test]
    fn metadata_rejects_malformed_xml_and_ignores_false_positive_text() {
        assert!(parse_ranges(&b"<worksheet><mergeCells>"[..]).is_ok());
        assert!(parse_ranges(&b"<worksheet><mergeCells></worksheet>"[..]).is_err());
        assert!(
            parse_ranges(&b"<worksheet><t>mergeCell</t></worksheet>"[..])
                .unwrap()
                .is_empty()
        );
    }
}
