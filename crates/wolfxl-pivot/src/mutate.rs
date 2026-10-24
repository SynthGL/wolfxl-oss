//! G17 / RFC-070 — minimal source-range mutation of existing pivot parts.
//!
//! Pure-Rust parser + rewriter for two narrow needs:
//!
//! 1. `parse_pivot_table_meta` — read `name`, `<location ref="...">`,
//!    and `cacheId` out of an existing `xl/pivotTables/pivotTable*.xml`.
//! 2. `parse_pivot_cache_source` — read `<cacheSource><worksheetSource
//!    ref=... [sheet=...]/></cacheSource>` out of an existing
//!    `xl/pivotCache/pivotCacheDefinition*.xml`, plus the
//!    `<cacheFields count="N">` count for shape comparison.
//! 3. `rewrite_cache_source` — given the original cache definition XML
//!    bytes, replace the `worksheetSource` `ref=` (and optionally
//!    `sheet=`) attributes in place. When the new source has a
//!    different column count from the original, also stamp
//!    `refreshOnLoad="1"` onto the surrounding `<pivotCacheDefinition>`
//!    element. Every other byte is preserved verbatim, including
//!    `<extLst>`, comments, and whitespace.
//!
//! Out of scope: field placement, filters, aggregation. See RFC-070 §3.

use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

/// Minimal metadata extracted from a `pivotTable*.xml` part.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PivotTableMeta {
    /// `<pivotTableDefinition name="...">`.
    pub name: String,
    /// `<location ref="A1:E20">`.
    pub location_ref: String,
    /// `<pivotTableDefinition cacheId="N">` — the workbook-scope id.
    pub cache_id: u32,
}

/// Minimal metadata extracted from a `pivotCacheDefinition*.xml` part.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PivotCacheSourceMeta {
    /// `<worksheetSource ref="A1:B5">`.
    pub range: String,
    /// `<worksheetSource sheet="Sheet1">`. Empty when omitted (the
    /// XLSX convention is that an absent `sheet=` falls back to the
    /// pivot table's own owning sheet — caller resolves).
    pub sheet: String,
    /// `<cacheFields count="N">` — used to detect shape changes that
    /// require `refreshOnLoad="1"`.
    pub field_count: u32,
}

/// Errors surfaced from the minimal parser.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum MutateError {
    /// Top-level `<pivotTableDefinition>` not found.
    NoPivotTableDefinition,
    /// Top-level `<pivotCacheDefinition>` not found.
    NoCacheDefinition,
    /// `<location>` element missing.
    NoLocation,
    /// `<cacheSource><worksheetSource>` missing.
    NoWorksheetSource,
    /// XML parse error.
    Xml(String),
}

impl core::fmt::Display for MutateError {
    fn fmt(&self, f: &mut core::fmt::Formatter<'_>) -> core::fmt::Result {
        match self {
            MutateError::NoPivotTableDefinition => {
                write!(f, "no <pivotTableDefinition> element found")
            }
            MutateError::NoCacheDefinition => {
                write!(f, "no <pivotCacheDefinition> element found")
            }
            MutateError::NoLocation => write!(f, "no <location> element found"),
            MutateError::NoWorksheetSource => {
                write!(f, "no <cacheSource><worksheetSource> element found")
            }
            MutateError::Xml(msg) => write!(f, "xml parse error: {msg}"),
        }
    }
}

impl std::error::Error for MutateError {}

fn local_name<'a>(e: &'a BytesStart<'a>) -> &'a [u8] {
    let n = e.name();
    let bytes = n.into_inner().as_bytes();
    match bytes.iter().rposition(|b| *b == b':') {
        Some(idx) => &bytes[idx + 1..],
        None => bytes,
    }
}

fn attr_value(e: &BytesStart<'_>, key: &[u8]) -> Option<String> {
    for attr in e.attributes().with_checks(false).flatten() {
        let name = attr.key.into_inner().as_bytes();
        let local = match name.iter().rposition(|b| *b == b':') {
            Some(idx) => &name[idx + 1..],
            None => name,
        };
        if local == key {
            return quick_xml::escape::unescape(&attr.value)
                .ok()
                .map(|cow| cow.into_owned());
        }
    }
    None
}

/// Parse a pivot table part (`xl/pivotTables/pivotTable*.xml`) and
/// pull out the minimal metadata G17 cares about.
pub fn parse_pivot_table_meta(xml: &[u8]) -> Result<PivotTableMeta, MutateError> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);

    let mut meta = PivotTableMeta::default();
    let mut buf = Vec::with_capacity(256);
    let mut saw_definition = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Eof) => break,
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match local_name(&e) {
                b"pivotTableDefinition" => {
                    saw_definition = true;
                    if let Some(v) = attr_value(&e, b"name") {
                        meta.name = v;
                    }
                    if let Some(v) = attr_value(&e, b"cacheId") {
                        meta.cache_id = v.parse().unwrap_or(0);
                    }
                }
                b"location" => {
                    if let Some(v) = attr_value(&e, b"ref") {
                        meta.location_ref = v;
                    }
                }
                _ => {}
            },
            Err(err) => return Err(MutateError::Xml(err.to_string())),
            _ => {}
        }
        buf.clear();
    }

    if !saw_definition {
        return Err(MutateError::NoPivotTableDefinition);
    }
    if meta.location_ref.is_empty() {
        return Err(MutateError::NoLocation);
    }
    Ok(meta)
}

/// Parse a pivot cache definition part
/// (`xl/pivotCache/pivotCacheDefinition*.xml`) for the `cacheSource`
/// metadata + `cacheFields` count.
pub fn parse_pivot_cache_source(xml: &[u8]) -> Result<PivotCacheSourceMeta, MutateError> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);

    let mut meta = PivotCacheSourceMeta::default();
    let mut buf = Vec::with_capacity(256);
    let mut saw_definition = false;
    let mut saw_worksheet_source = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Eof) => break,
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match local_name(&e) {
                b"pivotCacheDefinition" => {
                    saw_definition = true;
                }
                b"worksheetSource" => {
                    saw_worksheet_source = true;
                    if let Some(v) = attr_value(&e, b"ref") {
                        meta.range = v;
                    }
                    if let Some(v) = attr_value(&e, b"sheet") {
                        meta.sheet = v;
                    }
                }
                b"cacheFields" => {
                    if let Some(v) = attr_value(&e, b"count") {
                        meta.field_count = v.parse().unwrap_or(0);
                    }
                }
                _ => {}
            },
            Err(err) => return Err(MutateError::Xml(err.to_string())),
            _ => {}
        }
        buf.clear();
    }

    if !saw_definition {
        return Err(MutateError::NoCacheDefinition);
    }
    if !saw_worksheet_source {
        return Err(MutateError::NoWorksheetSource);
    }
    Ok(meta)
}

/// Compute the column count of an A1 range string. The grammar we
/// accept is the subset used by openpyxl `Reference` and emitted by
/// `wolfxl-pivot`'s own `cache.rs::emit_cache_source`:
///
///   * `A1` (single cell) → 1.
///   * `A1:B5` (rectangular range) → max_col - min_col + 1.
///
/// Returns `None` for unrecognised shapes (named ranges, multi-area
/// references, etc.); the rewriter then conservatively skips the
/// `refreshOnLoad` mutation rather than guess.
pub fn column_count_of_range(range: &str) -> Option<u32> {
    let (left, right) = match range.split_once(':') {
        Some((l, r)) => (l, r),
        None => (range, range),
    };
    let l_col = leading_col(left)?;
    let r_col = leading_col(right)?;
    Some(r_col.saturating_sub(l_col).saturating_add(1))
}

fn leading_col(cell: &str) -> Option<u32> {
    let s = cell.strip_prefix('$').unwrap_or(cell);
    let mut col: u32 = 0;
    let mut letters = 0u32;
    for ch in s.chars() {
        if ch.is_ascii_alphabetic() {
            letters += 1;
            col = col * 26 + (ch.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
        } else {
            break;
        }
    }
    if letters == 0 {
        None
    } else {
        Some(col)
    }
}

/// Rewrite the `<worksheetSource ref="..." [sheet="..."]/>` attributes
/// on an existing pivot cache definition, optionally stamping
/// `refreshOnLoad="1"` on the surrounding `<pivotCacheDefinition>`.
///
/// `force_refresh_on_load` controls the second behaviour: callers who
/// have already detected a shape change (column-count divergence) pass
/// `true`. When `false`, the existing `refreshOnLoad` attribute is
/// preserved verbatim.
///
/// Bytes outside the touched attributes round-trip verbatim. The
/// implementation operates at the byte level for the surrounding
/// elements to preserve `<extLst>`, comments, and whitespace exactly.
pub fn rewrite_cache_source(
    xml: &[u8],
    new_ref: &str,
    new_sheet: Option<&str>,
    force_refresh_on_load: bool,
) -> Result<Vec<u8>, MutateError> {
    // Sanity-check: confirm the input parses cleanly via quick-xml.
    // This catches the gross malformed-input cases before we start
    // doing byte-level edits. We don't use the parsed events for the
    // rewrite — we do a manual byte-level scan to keep full control
    // over the output bytes.
    let _ = parse_pivot_cache_source(xml)?;

    // Locate `<pivotCacheDefinition ...>` open tag and
    // `<worksheetSource .../>` element by manual scanning.
    let def_span =
        find_open_tag(xml, b"pivotCacheDefinition").ok_or(MutateError::NoCacheDefinition)?;
    let ws_span = find_open_tag(xml, b"worksheetSource").ok_or(MutateError::NoWorksheetSource)?;

    // Build new tag bytes.
    let new_def_tag = if force_refresh_on_load {
        Some(rebuild_tag_with_overrides(
            &xml[def_span.start..def_span.end],
            &[(b"refreshOnLoad" as &[u8], "1")],
        ))
    } else {
        None
    };

    let mut ws_overrides: Vec<(&[u8], String)> = Vec::with_capacity(2);
    ws_overrides.push((b"ref" as &[u8], new_ref.to_string()));
    if let Some(s) = new_sheet {
        ws_overrides.push((b"sheet" as &[u8], s.to_string()));
    }
    let ws_overrides_static: Vec<(&[u8], &str)> =
        ws_overrides.iter().map(|(k, v)| (*k, v.as_str())).collect();
    let new_ws_tag =
        rebuild_tag_with_overrides(&xml[ws_span.start..ws_span.end], &ws_overrides_static);

    // Splice in span order.
    let mut spans: Vec<(usize, usize, Vec<u8>)> = Vec::with_capacity(2);
    if let Some(t) = new_def_tag {
        spans.push((def_span.start, def_span.end, t));
    }
    spans.push((ws_span.start, ws_span.end, new_ws_tag));
    spans.sort_by_key(|s| s.0);

    let mut out = Vec::with_capacity(xml.len() + 64);
    let mut cursor = 0usize;
    for (start, end, repl) in spans {
        out.extend_from_slice(&xml[cursor..start]);
        out.extend_from_slice(&repl);
        cursor = end;
    }
    out.extend_from_slice(&xml[cursor..]);
    Ok(out)
}

#[derive(Debug)]
struct TagSpan {
    /// Byte index of the leading `<`.
    start: usize,
    /// Byte index one past the trailing `>`.
    end: usize,
}

/// Locate the first open (or self-closing) tag whose local name matches.
fn find_open_tag(xml: &[u8], target_local: &[u8]) -> Option<TagSpan> {
    let mut i = 0;
    while i < xml.len() {
        if xml[i] != b'<' {
            i += 1;
            continue;
        }
        // Skip XML declarations / comments / PI / DOCTYPE / CDATA.
        if xml[i..].starts_with(b"<?") {
            i += find_subseq(&xml[i..], b"?>")? + 2;
            continue;
        }
        if xml[i..].starts_with(b"<!--") {
            i += find_subseq(&xml[i..], b"-->")? + 3;
            continue;
        }
        if xml[i..].starts_with(b"<![CDATA[") {
            i += find_subseq(&xml[i..], b"]]>")? + 3;
            continue;
        }
        if xml[i..].starts_with(b"<!") {
            i += find_subseq(&xml[i..], b">")? + 1;
            continue;
        }
        if xml.get(i + 1) == Some(&b'/') {
            i += find_subseq(&xml[i..], b">")? + 1;
            continue;
        }
        let tag_start = i;
        let tag_end_offset = find_tag_end(&xml[i..])?;
        let tag_end = i + tag_end_offset + 1;
        let inner = &xml[tag_start + 1..tag_end - 1];
        let (name_bytes, _) = read_tag_name(inner);
        if strip_prefix(name_bytes) == target_local {
            return Some(TagSpan {
                start: tag_start,
                end: tag_end,
            });
        }
        i = tag_end;
    }
    None
}

fn find_subseq(haystack: &[u8], needle: &[u8]) -> Option<usize> {
    haystack.windows(needle.len()).position(|w| w == needle)
}

fn find_tag_end(s: &[u8]) -> Option<usize> {
    // Find first `>` outside any quoted attribute value. Skips the
    // leading `<`.
    let mut in_double = false;
    let mut in_single = false;
    for (i, b) in s.iter().enumerate().skip(1) {
        match *b {
            b'"' if !in_single => in_double = !in_double,
            b'\'' if !in_double => in_single = !in_single,
            b'>' if !in_double && !in_single => return Some(i),
            _ => {}
        }
    }
    None
}

fn read_tag_name(inner: &[u8]) -> (&[u8], usize) {
    let mut end = inner.len();
    for (i, b) in inner.iter().enumerate() {
        if matches!(b, b' ' | b'\t' | b'\n' | b'\r' | b'/') {
            end = i;
            break;
        }
    }
    (&inner[..end], end)
}

fn strip_prefix(name: &[u8]) -> &[u8] {
    match name.iter().rposition(|b| *b == b':') {
        Some(idx) => &name[idx + 1..],
        None => name,
    }
}

/// Reconstruct an element tag with attribute overrides applied.
///
/// `tag_bytes` is the raw byte slice covering `<name ...>` (or
/// `<name .../>`), inclusive of both delimiters. Attribute order from
/// the original is preserved; overrides hit by matching local name
/// replace the existing value, and any override not seen in the
/// original is appended at the end.
fn rebuild_tag_with_overrides(tag_bytes: &[u8], overrides: &[(&[u8], &str)]) -> Vec<u8> {
    let inner = &tag_bytes[1..tag_bytes.len() - 1];
    let self_close = inner.last() == Some(&b'/');
    let scan_end = if self_close {
        inner.len() - 1
    } else {
        inner.len()
    };
    let scan = &inner[..scan_end];

    let (name_bytes, name_end) = read_tag_name(scan);

    let mut out = Vec::with_capacity(tag_bytes.len() + 32);
    out.push(b'<');
    out.extend_from_slice(name_bytes);

    let mut applied = vec![false; overrides.len()];

    // Walk attributes in the original tag.
    let attrs_section = &scan[name_end..];
    let mut p = 0usize;
    while p < attrs_section.len() {
        // Skip whitespace.
        while p < attrs_section.len() && matches!(attrs_section[p], b' ' | b'\t' | b'\n' | b'\r') {
            p += 1;
        }
        if p >= attrs_section.len() {
            break;
        }
        // Read attribute name.
        let n_start = p;
        while p < attrs_section.len()
            && !matches!(attrs_section[p], b' ' | b'\t' | b'\n' | b'\r' | b'=' | b'/')
        {
            p += 1;
        }
        let n_end = p;
        if n_start == n_end {
            break;
        }
        let attr_name = &attrs_section[n_start..n_end];
        // Skip whitespace + '='.
        while p < attrs_section.len() && matches!(attrs_section[p], b' ' | b'\t' | b'\n' | b'\r') {
            p += 1;
        }
        if p >= attrs_section.len() || attrs_section[p] != b'=' {
            // Attribute without value — emit as-is, skip.
            out.push(b' ');
            out.extend_from_slice(attr_name);
            continue;
        }
        p += 1;
        while p < attrs_section.len() && matches!(attrs_section[p], b' ' | b'\t' | b'\n' | b'\r') {
            p += 1;
        }
        if p >= attrs_section.len() {
            break;
        }
        let quote = attrs_section[p];
        if quote != b'"' && quote != b'\'' {
            break;
        }
        p += 1;
        let v_start = p;
        while p < attrs_section.len() && attrs_section[p] != quote {
            p += 1;
        }
        let v_end = p;
        if p < attrs_section.len() {
            p += 1; // consume closing quote
        }
        let local = strip_prefix(attr_name);
        let override_idx = overrides.iter().position(|(k, _)| *k == local);
        out.push(b' ');
        out.extend_from_slice(attr_name);
        out.push(b'=');
        out.push(quote);
        if let Some(idx) = override_idx {
            applied[idx] = true;
            write_xml_attr_escaped(&mut out, overrides[idx].1);
        } else {
            out.extend_from_slice(&attrs_section[v_start..v_end]);
        }
        out.push(quote);
    }

    // Append any overrides that weren't matched.
    for (i, applied_flag) in applied.iter().enumerate() {
        if !applied_flag {
            out.push(b' ');
            out.extend_from_slice(overrides[i].0);
            out.push(b'=');
            out.push(b'"');
            write_xml_attr_escaped(&mut out, overrides[i].1);
            out.push(b'"');
        }
    }

    if self_close {
        out.extend_from_slice(b"/>");
    } else {
        out.push(b'>');
    }
    out
}

fn write_xml_attr_escaped(out: &mut Vec<u8>, s: &str) {
    for b in s.bytes() {
        match b {
            b'<' => out.extend_from_slice(b"&lt;"),
            b'>' => out.extend_from_slice(b"&gt;"),
            b'&' => out.extend_from_slice(b"&amp;"),
            b'"' => out.extend_from_slice(b"&quot;"),
            b'\'' => out.extend_from_slice(b"&apos;"),
            _ => out.push(b),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const SAMPLE_TABLE: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PivotTable1" cacheId="0" applyNumberFormats="0">
  <location ref="A1:B5" firstHeaderRow="0" firstDataRow="1" firstDataCol="0"/>
</pivotTableDefinition>"#;

    const SAMPLE_CACHE: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" refreshOnLoad="0" recordCount="2">
  <cacheSource type="worksheet">
    <worksheetSource ref="A1:B3" sheet="Sheet"/>
  </cacheSource>
  <cacheFields count="2">
    <cacheField name="region" numFmtId="0"/>
    <cacheField name="amount" numFmtId="0"/>
  </cacheFields>
</pivotCacheDefinition>"#;

    #[test]
    fn parses_pivot_table_meta() {
        let m = parse_pivot_table_meta(SAMPLE_TABLE).expect("parses");
        assert_eq!(m.name, "PivotTable1");
        assert_eq!(m.location_ref, "A1:B5");
        assert_eq!(m.cache_id, 0);
    }

    #[test]
    fn parses_pivot_cache_source() {
        let m = parse_pivot_cache_source(SAMPLE_CACHE).expect("parses");
        assert_eq!(m.range, "A1:B3");
        assert_eq!(m.sheet, "Sheet");
        assert_eq!(m.field_count, 2);
    }

    #[test]
    fn parses_cache_with_extlst_preserved_in_rewrite() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" refreshOnLoad="0" recordCount="2">
  <cacheSource type="worksheet"><worksheetSource ref="A1:B3" sheet="Sheet"/></cacheSource>
  <cacheFields count="2"><cacheField name="region" numFmtId="0"/><cacheField name="amount" numFmtId="0"/></cacheFields>
  <extLst><ext uri="{XYZ}"><foo/></ext></extLst>
</pivotCacheDefinition>"#;
        let out = rewrite_cache_source(xml, "A1:B5", Some("Sheet"), false).expect("rewrite");
        let s = std::str::from_utf8(&out).unwrap();
        assert!(s.contains(r#"ref="A1:B5""#));
        assert!(s.contains(r#"<extLst><ext uri="{XYZ}"><foo/></ext></extLst>"#));
    }

    #[test]
    fn parses_worksheet_source_without_sheet_attr() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cacheSource type="worksheet"><worksheetSource ref="A1:B3"/></cacheSource>
  <cacheFields count="2"></cacheFields>
</pivotCacheDefinition>"#;
        let m = parse_pivot_cache_source(xml).expect("parses");
        assert_eq!(m.range, "A1:B3");
        assert_eq!(m.sheet, "");
    }

    #[test]
    fn rejects_malformed_xml() {
        let bad = b"<<<not xml";
        assert!(parse_pivot_cache_source(bad).is_err());
    }

    #[test]
    fn rewrite_sets_refresh_on_load() {
        let out =
            rewrite_cache_source(SAMPLE_CACHE, "A1:C5", Some("Sheet"), true).expect("rewrite");
        let s = std::str::from_utf8(&out).unwrap();
        assert!(s.contains(r#"refreshOnLoad="1""#));
        assert!(s.contains(r#"ref="A1:C5""#));
        // Original `refreshOnLoad="0"` must NOT survive — only the
        // single overridden value should remain.
        assert_eq!(s.matches("refreshOnLoad=").count(), 1);
    }

    #[test]
    fn rewrite_preserves_unrelated_attrs() {
        let out =
            rewrite_cache_source(SAMPLE_CACHE, "A1:B5", Some("Sheet"), false).expect("rewrite");
        let s = std::str::from_utf8(&out).unwrap();
        assert!(s.contains(r#"r:id="rId1""#));
        assert!(s.contains(r#"recordCount="2""#));
        assert!(s.contains(
            r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#
        ));
    }

    #[test]
    fn column_count_basic() {
        assert_eq!(column_count_of_range("A1"), Some(1));
        assert_eq!(column_count_of_range("A1:B5"), Some(2));
        assert_eq!(column_count_of_range("A1:C10"), Some(3));
        assert_eq!(column_count_of_range("$A$1:$D$5"), Some(4));
    }

    #[test]
    fn rewrite_appends_missing_sheet_attr() {
        let xml = br#"<?xml version="1.0"?><pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cacheSource type="worksheet"><worksheetSource ref="A1:B3"/></cacheSource><cacheFields count="2"/></pivotCacheDefinition>"#;
        let out = rewrite_cache_source(xml, "A1:B5", Some("NewSheet"), false).expect("rewrite");
        let s = std::str::from_utf8(&out).unwrap();
        assert!(s.contains(r#"ref="A1:B5""#));
        assert!(s.contains(r#"sheet="NewSheet""#));
    }
}
