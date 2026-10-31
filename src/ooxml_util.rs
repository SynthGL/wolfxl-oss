use pyo3::exceptions::PyIOError;
use pyo3::prelude::*;

use std::collections::{HashMap, HashSet};
use std::io::{Read, Seek};

use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader as XmlReader;
use zip::ZipArchive;

use wolfxl_rels::package_guard::{
    DEFAULT_MAX_COMPRESSION_RATIO, DEFAULT_MAX_ZIP_ENTRIES, DEFAULT_MAX_ZIP_ENTRY_BYTES,
    DEFAULT_MAX_ZIP_TOTAL_BYTES,
};

fn env_usize(name: &str, default: usize) -> usize {
    std::env::var(name)
        .ok()
        .and_then(|v| v.parse::<usize>().ok())
        .unwrap_or(default)
}

fn env_u64(name: &str, default: u64) -> u64 {
    std::env::var(name)
        .ok()
        .and_then(|v| v.parse::<u64>().ok())
        .unwrap_or(default)
}

fn max_entries() -> usize {
    env_usize("WOLFXL_MAX_ZIP_ENTRIES", DEFAULT_MAX_ZIP_ENTRIES)
}

fn max_entry_bytes() -> u64 {
    env_u64("WOLFXL_MAX_ZIP_ENTRY_BYTES", DEFAULT_MAX_ZIP_ENTRY_BYTES)
}

fn max_total_bytes() -> u64 {
    env_u64("WOLFXL_MAX_ZIP_TOTAL_BYTES", DEFAULT_MAX_ZIP_TOTAL_BYTES)
}

fn max_compression_ratio() -> u64 {
    env_u64(
        "WOLFXL_MAX_ZIP_COMPRESSION_RATIO",
        DEFAULT_MAX_COMPRESSION_RATIO,
    )
}

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

pub fn validate_zip_archive<R: Read + Seek>(zip: &mut ZipArchive<R>) -> PyResult<()> {
    if zip.len() > max_entries() {
        return Err(PyErr::new::<PyIOError, _>(format!(
            "OOXML package has too many ZIP entries: {} > {}",
            zip.len(),
            max_entries()
        )));
    }

    let mut names = HashSet::with_capacity(zip.len());
    let mut casefolded_names = HashSet::with_capacity(zip.len());
    let mut total: u64 = 0;
    for i in 0..zip.len() {
        let file = zip
            .by_index(i)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("ZIP entry read error: {e}")))?;
        let name = file
            .name()
            .map_err(|e| PyIOError::new_err(format!("ZIP entry name error: {e}")))?
            .into_owned();
        validate_part_name(&name)?;
        if !names.insert(name.clone()) {
            return Err(PyErr::new::<PyIOError, _>(format!(
                "OOXML package contains duplicate ZIP entry: {name}"
            )));
        }
        let casefolded_name = name.to_ascii_lowercase();
        if !casefolded_names.insert(casefolded_name) {
            return Err(PyErr::new::<PyIOError, _>(format!(
                "OOXML package contains case-insensitive duplicate ZIP entry: {name}"
            )));
        }
        validate_zip_entry_metadata(&name, file.size(), file.compressed_size())?;
        total = total.saturating_add(file.size());
        if total > max_total_bytes() {
            return Err(PyErr::new::<PyIOError, _>(format!(
                "OOXML package is too large: {total} > {} uncompressed bytes",
                max_total_bytes()
            )));
        }
    }
    Ok(())
}

pub fn validate_part_name(name: &str) -> PyResult<()> {
    let invalid = name.is_empty()
        || name.starts_with('/')
        || name.starts_with('\\')
        || name.contains('\\')
        || name
            .split('/')
            .any(|part| part == ".." || part.contains(':'));
    if invalid {
        return Err(PyErr::new::<PyIOError, _>(format!(
            "unsafe OOXML package part path: {name}"
        )));
    }
    Ok(())
}

pub fn validate_zip_entry_metadata(name: &str, size: u64, compressed_size: u64) -> PyResult<()> {
    if size > max_entry_bytes() {
        return Err(PyErr::new::<PyIOError, _>(format!(
            "OOXML package part {name} is too large: {size} > {} bytes",
            max_entry_bytes()
        )));
    }
    if size > 0 && compressed_size == 0 {
        return Err(PyErr::new::<PyIOError, _>(format!(
            "OOXML package part {name} has invalid compressed size"
        )));
    }
    if compressed_size > 0 && size > compressed_size.saturating_mul(max_compression_ratio()) {
        return Err(PyErr::new::<PyIOError, _>(format!(
            "OOXML package part {name} exceeds compression ratio limit"
        )));
    }
    Ok(())
}

pub fn join_and_normalize(base_dir: &str, target: &str) -> String {
    let t = target.trim_start_matches('/');
    let combined = if t.starts_with("xl/") {
        t.to_string()
    } else {
        format!("{base_dir}{t}")
    };
    normalize_zip_path(&combined)
}

pub fn attr_value(e: &BytesStart<'_>, key: &[u8]) -> Option<String> {
    for a in e.attributes().with_checks(false).flatten() {
        if a.key.as_ref().as_bytes() == key {
            if let Ok(v) = quick_xml::escape::unescape(&a.value) {
                return Some(v.to_string());
            }
            return Some(a.value.into_owned());
        }
    }
    None
}

fn relationship_id_attr(e: &BytesStart<'_>) -> Option<String> {
    attr_value(e, b"r:id").or_else(|| attr_value_by_suffix(e, b":id"))
}

fn attr_value_by_suffix(e: &BytesStart<'_>, suffix: &[u8]) -> Option<String> {
    for a in e.attributes().with_checks(false).flatten() {
        if a.key.as_ref().as_bytes().ends_with(suffix) {
            if let Ok(v) = quick_xml::escape::unescape(&a.value) {
                return Some(v.to_string());
            }
            return Some(a.value.into_owned());
        }
    }
    None
}

pub fn parse_workbook_sheet_rids(xml: &str) -> PyResult<Vec<(String, String)>> {
    let mut reader = XmlReader::from_str(xml);
    let mut buf: Vec<u8> = Vec::new();
    let mut out: Vec<(String, String)> = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.local_name().as_ref() == "sheet" {
                    let name = attr_value(&e, b"name");
                    let rid = relationship_id_attr(&e);
                    if let (Some(n), Some(r)) = (name, rid) {
                        out.push((n, r));
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(PyErr::new::<PyIOError, _>(format!(
                    "Failed to parse workbook.xml: {e}"
                )))
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

pub fn parse_relationship_targets(xml: &str) -> PyResult<HashMap<String, String>> {
    // Body-only swap: the wolfxl-rels crate is the single source of truth
    // for the rels grammar. Signature (and lenient skipping of malformed
    // entries) are preserved so existing patcher call sites need not change.
    let graph = wolfxl_rels::RelsGraph::parse(xml.as_bytes())
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to parse rels: {e}")))?;
    let mut out: HashMap<String, String> = HashMap::with_capacity(graph.len());
    for r in graph.iter() {
        out.insert(r.id.0.clone(), r.target.clone());
    }
    Ok(out)
}

pub fn zip_read_to_string<R: Read + Seek>(zip: &mut ZipArchive<R>, name: &str) -> PyResult<String> {
    match zip.by_name(name) {
        Ok(mut f) => {
            validate_zip_entry_metadata(name, f.size(), f.compressed_size())?;
            let mut out = String::new();
            f.read_to_string(&mut out)
                .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to read {name}: {e}")))?;
            return Ok(out);
        }
        Err(zip::result::ZipError::FileNotFound) => {}
        Err(e) => {
            return Err(PyErr::new::<PyIOError, _>(format!(
                "Missing zip entry {name}: {e}"
            )));
        }
    }
    let actual_name = resolve_zip_name_case_insensitive(zip, name)?
        .ok_or_else(|| PyErr::new::<PyIOError, _>(format!("Missing zip entry {name}")))?;
    let mut f = zip
        .by_name(&actual_name)
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("Missing zip entry {name}: {e}")))?;
    validate_zip_entry_metadata(&actual_name, f.size(), f.compressed_size())?;
    let mut out = String::new();
    f.read_to_string(&mut out)
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to read {name}: {e}")))?;
    Ok(out)
}

pub fn zip_read_to_string_opt<R: Read + Seek>(
    zip: &mut ZipArchive<R>,
    name: &str,
) -> PyResult<Option<String>> {
    match zip.by_name(name) {
        Ok(mut f) => {
            validate_zip_entry_metadata(name, f.size(), f.compressed_size())?;
            let mut out = String::new();
            f.read_to_string(&mut out)
                .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to read {name}: {e}")))?;
            return Ok(Some(out));
        }
        Err(zip::result::ZipError::FileNotFound) => {}
        Err(e) => {
            return Err(PyErr::new::<PyIOError, _>(format!(
                "Zip error reading {name}: {e}"
            )));
        }
    }
    let Some(actual_name) = resolve_zip_name_case_insensitive(zip, name)? else {
        return Ok(None);
    };
    match zip.by_name(&actual_name) {
        Ok(mut f) => {
            validate_zip_entry_metadata(&actual_name, f.size(), f.compressed_size())?;
            let mut out = String::new();
            f.read_to_string(&mut out)
                .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to read {name}: {e}")))?;
            Ok(Some(out))
        }
        Err(e) => Err(PyErr::new::<PyIOError, _>(format!(
            "Zip error reading {name}: {e}"
        ))),
    }
}

fn resolve_zip_name_case_insensitive<R: Read + Seek>(
    zip: &ZipArchive<R>,
    name: &str,
) -> PyResult<Option<String>> {
    for candidate in zip.file_names() {
        let candidate =
            candidate.map_err(|e| PyIOError::new_err(format!("ZIP entry name error: {e}")))?;
        if candidate.eq_ignore_ascii_case(name) {
            return Ok(Some(candidate.into_owned()));
        }
    }
    Ok(None)
}

/// Resolve quick-xml 0.42's split references without normalizing literal
/// whitespace. Matches the former `BytesText::unescape` contract.
pub(crate) fn resolve_general_ref(
    reference: &quick_xml::events::BytesRef<'_>,
) -> quick_xml::Result<std::borrow::Cow<'static, str>> {
    use quick_xml::escape::{resolve_predefined_entity, EscapeError};
    if let Some(ch) = reference.resolve_char_ref()? {
        return Ok(std::borrow::Cow::Owned(ch.to_string()));
    }
    let name: &str = reference;
    match resolve_predefined_entity(name) {
        Some(text) => Ok(std::borrow::Cow::Borrowed(text)),
        None => Err(EscapeError::UnrecognizedEntity(1..name.len() + 1, name.to_owned()).into()),
    }
}

/// Collect a single text node before transforming it. quick-xml 0.42 emits
/// references separately; trimming or transforming each fragment loses data.
/// Keep the raw spelling so callers can retain their existing error policy.
#[derive(Default)]
pub(crate) struct XmlTextNode {
    raw: String,
}

impl XmlTextNode {
    pub(crate) fn push_text(&mut self, text: &quick_xml::events::BytesText<'_>) {
        self.raw.push_str(text);
    }

    pub(crate) fn push_ref(&mut self, reference: &quick_xml::events::BytesRef<'_>) {
        self.raw.push('&');
        self.raw.push_str(reference);
        self.raw.push(';');
    }

    pub(crate) fn raw(&self) -> &str {
        &self.raw
    }

    pub(crate) fn clear(&mut self) {
        self.raw.clear();
    }
}

#[cfg(test)]
mod tests {
    use super::validate_zip_entry_metadata;

    #[test]
    fn compression_ratio_rejects_fractional_over_limit() {
        let result = validate_zip_entry_metadata("xl/workbook.xml", 10_001, 10);
        assert!(result.is_err());
    }

    #[test]
    fn references_resolve_without_normalizing_referenced_whitespace() {
        use quick_xml::events::BytesRef;
        for (reference, expected) in [
            ("amp", "&"),
            ("lt", "<"),
            ("gt", ">"),
            ("apos", "'"),
            ("quot", "\""),
            ("#10", "\n"),
            ("#13", "\r"),
            ("#x1F600", "😀"),
        ] {
            assert_eq!(
                super::resolve_general_ref(&BytesRef::new(reference)).unwrap(),
                expected
            );
        }
        for reference in ["unknown", "#0", "#xD800", "#x110000"] {
            assert!(super::resolve_general_ref(&BytesRef::new(reference)).is_err());
        }
    }

    #[test]
    fn workbook_attributes_keep_entities_and_literal_whitespace() {
        let xml = "<workbook><sheets><sheet name=\" A\t&amp;\r\nB&#10; \" \
                   r:id=\"rId&#49;\"/></sheets></workbook>";
        let sheets = super::parse_workbook_sheet_rids(xml).unwrap();
        assert_eq!(
            sheets,
            vec![(" A\t&\r\nB\n ".to_owned(), "rId1".to_owned())]
        );
    }
}
