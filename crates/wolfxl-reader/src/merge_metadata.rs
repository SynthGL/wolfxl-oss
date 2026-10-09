//! Bounded merge metadata scans. Never construct a worksheet cell model.

use std::collections::{HashMap, HashSet};
use std::io::{BufRead, BufReader, Read};

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

/// Conservative byte-only probe, matching the Python source-negative shortcut.
/// A positive includes comments, CDATA and prefixed names; callers parse XML.
/// A negative proves absence of mergeCell bytes, not XML well-formedness.
pub(super) fn may_have_ranges(book: &NativeXlsxBook, sheet: &str) -> Result<bool> {
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
    contains_marker(part)
}

fn contains_marker<R: Read>(mut source: R) -> Result<bool> {
    // Keep the final eight bytes so mergeCell can cross any short-read boundary.
    let mut buffer = [0u8; 64 * 1024 + 8];
    let mut retained = 0;
    loop {
        let count = source.read(&mut buffer[retained..])?;
        if count == 0 {
            return Ok(false);
        }
        let end = retained + count;
        if memchr::memmem::find(&buffer[..end], b"mergeCell").is_some() {
            return Ok(true);
        }
        retained = end.min(8);
        buffer.copy_within(end - retained..end, 0);
    }
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

    struct ShortReads<'a> {
        source: &'a [u8],
        limit: usize,
    }

    impl Read for ShortReads<'_> {
        fn read(&mut self, buffer: &mut [u8]) -> std::io::Result<usize> {
            let count = self.limit.min(buffer.len()).min(self.source.len());
            buffer[..count].copy_from_slice(&self.source[..count]);
            self.source = &self.source[count..];
            Ok(count)
        }
    }

    #[test]
    fn negative_probe_handles_short_reads_and_chunk_boundaries() {
        for limit in 1..=16 {
            for prefix in 0..=16 {
                let mut bytes = vec![b'x'; prefix];
                bytes.extend_from_slice(b"<m:mergeCell ref=\"A1:B2\"/>");
                assert!(contains_marker(ShortReads {
                    source: &bytes,
                    limit
                })
                .unwrap());
                assert!(!contains_marker(ShortReads {
                    source: b"<worksheet/>",
                    limit
                })
                .unwrap());
            }
        }
        for offset in 65528..=65544 {
            let mut bytes = vec![b'x'; offset];
            bytes.extend_from_slice(b"mergeCell");
            assert!(contains_marker(&bytes[..]).unwrap());
        }
    }

    #[test]
    fn positive_probe_defers_namespaces_comments_cdata_and_text_to_exact_parser() {
        for xml in [
            &br#"<worksheet><!-- mergeCell ref="A1:B2" --></worksheet>"#[..],
            &br#"<worksheet><![CDATA[<mergeCell ref="A1:B2"/>]]></worksheet>"#[..],
            &b"<worksheet><t>mergeCell</t></worksheet>"[..],
        ] {
            assert!(contains_marker(xml).unwrap());
            assert!(parse_ranges(xml).unwrap().is_empty());
        }
        let xml = br#"<m:worksheet xmlns:m="urn:test"><m:mergeCell ref="A1:B2"/></m:worksheet>"#;
        assert!(contains_marker(&xml[..]).unwrap());
        assert_eq!(parse_ranges(&xml[..]).unwrap(), vec!["A1:B2"]);
    }

    #[test]
    fn byte_negative_does_not_weaken_exact_parser_error_boundary() {
        let malformed = b"<worksheet><sheetData></worksheet>";
        assert!(!contains_marker(&malformed[..]).unwrap());
        assert!(parse_ranges(&malformed[..]).is_err());
        let positive = b"<worksheet><!-- mergeCell --><sheetData></worksheet>";
        assert!(contains_marker(&positive[..]).unwrap());
        assert!(parse_ranges(&positive[..]).is_err());
    }

    #[test]
    fn probe_propagates_source_read_errors() {
        struct Broken;
        impl Read for Broken {
            fn read(&mut self, _: &mut [u8]) -> std::io::Result<usize> {
                Err(std::io::Error::other("probe read failed"))
            }
        }
        assert!(contains_marker(Broken).is_err());
    }

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
