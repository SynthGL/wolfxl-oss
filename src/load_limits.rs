//! Bounded, streaming admission before eager worksheet materialization.
use std::collections::HashSet;
use std::fs::File;
use std::io::{BufReader, Cursor, Read, Seek};

use pyo3::exceptions::{PyIOError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::PyBytes;
use quick_xml::events::Event;
use quick_xml::Reader;
use wolfxl_rels::package_guard::{check, LimitExceeded, Limits};
use zip::ZipArchive;

fn limit_error(error: LimitExceeded) -> PyErr {
    Python::attach(|py| {
        match py
            .import("wolfxl._limits")
            .and_then(|m| m.getattr("LimitExceeded"))
        {
            Ok(class) => match class.call1((error.to_string(),)) {
                Ok(value) => PyErr::from_value(value),
                Err(err) => err,
            },
            Err(err) => err,
        }
    })
}

#[pyfunction]
pub fn load_limit_defaults() -> (u64, u64, u64, u64, u64, usize, u64) {
    let limits = Limits::default();
    (
        limits.max_package_bytes,
        limits.max_part_bytes,
        limits.max_cells,
        limits.max_shared_string_bytes,
        limits.max_materialized_string_bytes,
        wolfxl_rels::package_guard::DEFAULT_MAX_ZIP_ENTRIES,
        wolfxl_rels::package_guard::DEFAULT_MAX_COMPRESSION_RATIO,
    )
}

#[pyfunction]
pub fn validate_load_limits(
    source: &Bound<'_, PyAny>,
    max_package_bytes: u64,
    max_part_bytes: u64,
    max_cells: u64,
    max_shared_string_bytes: u64,
    max_materialized_string_bytes: u64,
) -> PyResult<()> {
    let limits = Limits {
        max_package_bytes,
        max_part_bytes,
        max_cells,
        max_shared_string_bytes,
        max_materialized_string_bytes,
    };
    if let Ok(data) = source.cast::<PyBytes>() {
        validate(Cursor::new(data.as_bytes()), limits)
    } else if let Ok(path) = source.extract::<String>() {
        validate(
            File::open(path).map_err(|e| PyIOError::new_err(e.to_string()))?,
            limits,
        )
    } else {
        Err(PyValueError::new_err("expected workbook path or bytes"))
    }
}

fn validate<R: Read + Seek>(source: R, limits: Limits) -> PyResult<()> {
    let mut zip = ZipArchive::new(source).map_err(|e| PyIOError::new_err(e.to_string()))?;
    let mut package_bytes = 0u64;
    // Check the central directory before decompressing any entry.
    for index in 0..zip.len() {
        let entry = zip
            .by_index(index)
            .map_err(|e| PyIOError::new_err(e.to_string()))?;
        check("max_part_bytes", entry.size(), limits.max_part_bytes).map_err(limit_error)?;
        package_bytes = package_bytes.saturating_add(entry.size());
        check("max_package_bytes", package_bytes, limits.max_package_bytes).map_err(limit_error)?;
    }
    // Retain all existing path, duplicate-entry and compression-ratio checks.
    crate::ooxml_util::validate_zip_archive(&mut zip)?;
    let worksheet_parts = worksheet_parts(&mut zip)?;
    let mut cells = 0u64;
    let mut shared_bytes = 0u64;
    let mut text_bytes = 0u64;
    for index in 0..zip.len() {
        let entry = zip
            .by_index(index)
            .map_err(|e| PyIOError::new_err(e.to_string()))?;
        let name = entry
            .name()
            .map_err(|e| PyIOError::new_err(e.to_string()))?
            .to_ascii_lowercase();
        let shared = name == "xl/sharedstrings.xml";
        let worksheet = worksheet_parts.contains(&name);
        if !shared && !worksheet {
            continue;
        }
        let mut reader = Reader::from_reader(BufReader::new(entry));
        let mut buf = Vec::new();
        let mut in_text = false;
        let mut in_formula = false;
        let mut in_value = false;
        let mut string_cell = false;
        loop {
            let event = reader.read_event_into(&mut buf);
            let empty = matches!(&event, Ok(Event::Empty(_)));
            match event {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    let tag = e.local_name();
                    match tag.as_ref().as_bytes() {
                        b"c" if worksheet => {
                            cells = cells.saturating_add(1);
                            check("max_cells", cells, limits.max_cells).map_err(limit_error)?;
                            string_cell = e.attributes().flatten().any(|a| {
                                a.key.as_ref().as_bytes() == b"t"
                                    && matches!(a.value.as_ref(), "str" | "inlineStr" | "e" | "d")
                            });
                        }
                        b"t" if !empty => in_text = true,
                        b"f" if !empty => in_formula = true,
                        b"v" if !empty => in_value = true,
                        _ => {}
                    }
                }
                Ok(Event::End(e)) => match e.local_name().as_ref().as_bytes() {
                    b"t" => in_text = false,
                    b"f" => in_formula = false,
                    b"v" => in_value = false,
                    b"c" => string_cell = false,
                    _ => {}
                },
                Ok(Event::Text(e)) => {
                    let len = e.as_ref().len() as u64;
                    account(
                        len,
                        shared && in_text,
                        in_text || in_formula || (in_value && string_cell),
                        &mut shared_bytes,
                        &mut text_bytes,
                        limits,
                    )?;
                }
                Ok(Event::CData(e)) => {
                    let len = e.as_ref().len() as u64;
                    account(
                        len,
                        shared && in_text,
                        in_text || in_formula || (in_value && string_cell),
                        &mut shared_bytes,
                        &mut text_bytes,
                        limits,
                    )?;
                }
                Ok(Event::GeneralRef(e)) => {
                    let len = e
                        .resolve_char_ref()
                        .ok()
                        .flatten()
                        .map(|c| c.len_utf8() as u64)
                        .or_else(|| {
                            quick_xml::escape::resolve_predefined_entity(&e).map(|s| s.len() as u64)
                        })
                        .unwrap_or(0);
                    account(
                        len,
                        shared && in_text,
                        in_text || in_formula || (in_value && string_cell),
                        &mut shared_bytes,
                        &mut text_bytes,
                        limits,
                    )?;
                }
                // Preserve the existing reader's malformed-XML diagnostics.
                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
            buf.clear();
        }
    }
    Ok(())
}

fn worksheet_parts<R: Read + Seek>(zip: &mut ZipArchive<R>) -> PyResult<HashSet<String>> {
    let mut out = HashSet::new();
    for index in 0..zip.len() {
        let entry = zip
            .by_index(index)
            .map_err(|e| PyIOError::new_err(e.to_string()))?;
        let name = entry
            .name()
            .map_err(|e| PyIOError::new_err(e.to_string()))?;
        if !name.eq_ignore_ascii_case("xl/_rels/workbook.xml.rels") {
            continue;
        }
        let mut reader = Reader::from_reader(BufReader::new(entry));
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) | Ok(Event::Empty(e))
                    if e.local_name().as_ref() == "Relationship" =>
                {
                    let kind = crate::ooxml_util::attr_value(&e, b"Type").unwrap_or_default();
                    if kind.ends_with("/worksheet") || kind.ends_with("/chartsheet") {
                        if let Some(target) = crate::ooxml_util::attr_value(&e, b"Target") {
                            out.insert(
                                crate::ooxml_util::join_and_normalize("xl/", &target)
                                    .to_ascii_lowercase(),
                            );
                        }
                    }
                }
                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
            buf.clear();
        }
    }
    Ok(out)
}

fn account(
    len: u64,
    shared: bool,
    materialized: bool,
    shared_bytes: &mut u64,
    text_bytes: &mut u64,
    limits: Limits,
) -> PyResult<()> {
    if shared {
        *shared_bytes = shared_bytes.saturating_add(len);
        check(
            "max_shared_string_bytes",
            *shared_bytes,
            limits.max_shared_string_bytes,
        )
        .map_err(limit_error)?;
    }
    if materialized {
        *text_bytes = text_bytes.saturating_add(len);
        check(
            "max_materialized_string_bytes",
            *text_bytes,
            limits.max_materialized_string_bytes,
        )
        .map_err(limit_error)?;
    }
    Ok(())
}
