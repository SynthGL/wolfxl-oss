//! Bounded, streaming admission before eager worksheet materialization.
use std::collections::HashMap;
use std::fs::File;
use std::io::{BufReader, Read, Seek};

use pyo3::exceptions::PyIOError;
use pyo3::prelude::*;
use quick_xml::events::Event;
use quick_xml::Reader;
use wolfxl_reader::WorkbookBytesCursor;
use wolfxl_rels::package_guard::{check, LimitExceeded, Limits};
use zip::ZipArchive;

use crate::py_workbook_bytes::PyWorkbookBytes;

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
    permissive: bool,
) -> PyResult<()> {
    let limits = Limits {
        max_package_bytes,
        max_part_bytes,
        max_cells,
        max_shared_string_bytes,
        max_materialized_string_bytes,
    };
    if let Ok(path) = source.extract::<String>() {
        validate(
            File::open(path).map_err(|e| PyIOError::new_err(e.to_string()))?,
            limits,
            permissive,
        )
    } else {
        // `bytes` or a read-only mapping of the workbook file; the cursor
        // releases consumed mapped pages as the package is scanned.
        let data = PyWorkbookBytes::new(source)?;
        validate(WorkbookBytesCursor::new(&data), limits, permissive)
    }
}

fn validate<R: Read + Seek>(source: R, limits: Limits, permissive: bool) -> PyResult<()> {
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
    let worksheet_parts = worksheet_parts(&mut zip, permissive)?;
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
        let copies = worksheet_parts.get(&name).copied().unwrap_or(0);
        let worksheet = copies != 0;
        if !shared && !worksheet {
            continue;
        }
        let mut reader = Reader::from_reader(BufReader::new(entry));
        let mut buf = Vec::new();
        let mut in_text = false;
        let mut in_formula = false;
        let mut in_value = false;
        let mut string_cell = false;
        let mut formula_index = None;
        let mut formula_len = 0u64;
        let mut shared_formulas: HashMap<String, (Option<u64>, u64)> = HashMap::new();
        loop {
            let event = reader.read_event_into(&mut buf);
            let len = match &event {
                Ok(Event::Text(e)) => Some(normalized_text_len(e)),
                Ok(Event::CData(e)) => Some(normalized_text_len(e)),
                Ok(Event::GeneralRef(e)) => Some(
                    e.resolve_char_ref()
                        .ok()
                        .flatten()
                        .map(|c| c.len_utf8() as u64)
                        .or_else(|| {
                            quick_xml::escape::resolve_predefined_entity(e).map(|s| s.len() as u64)
                        })
                        .unwrap_or(0),
                ),
                _ => None,
            };
            if let Some(len) = len {
                if in_formula {
                    formula_len = formula_len.saturating_add(len);
                }
                if shared && in_text {
                    account(len, true, false, &mut shared_bytes, &mut text_bytes, limits)?;
                }
                account(
                    len.saturating_mul(copies.max(1)),
                    false,
                    in_text || in_formula || (in_value && string_cell),
                    &mut shared_bytes,
                    &mut text_bytes,
                    limits,
                )?;
            }
            let empty = matches!(&event, Ok(Event::Empty(_)));
            match event {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    let tag = e.local_name();
                    match tag.as_ref().as_bytes() {
                        b"c" if worksheet => {
                            cells = cells.saturating_add(copies);
                            check("max_cells", cells, limits.max_cells).map_err(limit_error)?;
                            string_cell = matches!(
                                crate::ooxml_util::attr_value(&e, b"t").as_deref(),
                                Some("str" | "inlineStr" | "e" | "d")
                            );
                        }
                        b"t" if !empty => in_text = true,
                        b"f" => {
                            formula_index = if crate::ooxml_util::attr_value(&e, b"t").as_deref()
                                == Some("shared")
                            {
                                crate::ooxml_util::attr_value(&e, b"si")
                            } else {
                                None
                            };
                            formula_len = 0;
                            in_formula = !empty;
                            if empty {
                                record_shared_formula(
                                    &mut shared_formulas,
                                    formula_index.take(),
                                    0,
                                );
                            }
                        }
                        b"v" if !empty => in_value = true,
                        _ => {}
                    }
                }
                Ok(Event::End(e)) => match e.local_name().as_ref().as_bytes() {
                    b"t" => in_text = false,
                    b"f" => {
                        in_formula = false;
                        record_shared_formula(
                            &mut shared_formulas,
                            formula_index.take(),
                            formula_len,
                        );
                    }
                    b"v" => in_value = false,
                    b"c" => string_cell = false,
                    _ => {}
                },
                // Preserve the existing reader's malformed-XML diagnostics.
                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
            buf.clear();
        }
        for (master_len, children) in shared_formulas.into_values() {
            account(
                master_len
                    .unwrap_or(0)
                    .saturating_mul(children)
                    .saturating_mul(copies),
                false,
                true,
                &mut shared_bytes,
                &mut text_bytes,
                limits,
            )?;
        }
    }
    Ok(())
}

fn normalized_text_len(text: &str) -> u64 {
    // XML 1.0 maps CR LF to one LF before materializing literal text.
    (text.len() - memchr::memmem::find_iter(text.as_bytes(), b"\r\n").count()) as u64
}

fn worksheet_parts<R: Read + Seek>(
    zip: &mut ZipArchive<R>,
    permissive: bool,
) -> PyResult<HashMap<String, u64>> {
    let mut workbook_part = "xl/workbook.xml".to_owned();
    let canonical = read_metadata_part(zip, &workbook_part)?;
    if canonical.is_none() {
        let root_rels = read_metadata_part(zip, "_rels/.rels")?.unwrap_or_default();
        if let Ok(rels) = wolfxl_rels::RelsGraph::parse(root_rels.as_bytes()) {
            if let Some(rel) = rels
                .iter()
                .find(|rel| rel.rel_type.ends_with("/officeDocument"))
            {
                workbook_part = rel.target.trim_start_matches('/').to_owned();
            }
        }
    }
    let workbook_xml = match canonical {
        Some(xml) => xml,
        None => read_metadata_part(zip, &workbook_part)?.unwrap_or_default(),
    };
    let (directory, basename) = workbook_part
        .rsplit_once('/')
        .unwrap_or(("", &workbook_part));
    let rels_part = if directory.is_empty() {
        format!("_rels/{basename}.rels")
    } else {
        format!("{directory}/_rels/{basename}.rels")
    };
    let workbook_rels = read_metadata_part(zip, &rels_part)?.unwrap_or_default();
    let mut out = HashMap::new();
    if let Ok(targets) =
        wolfxl_reader::worksheet_load_targets(&workbook_xml, &workbook_rels, permissive)
    {
        // The loader normalizes the workbook and its rels to xl/ before loading.
        for target in targets {
            let copies = out.entry(target.to_ascii_lowercase()).or_insert(0u64);
            *copies = copies.saturating_add(1);
        }
    }
    Ok(out)
}

fn read_metadata_part<R: Read + Seek>(
    zip: &mut ZipArchive<R>,
    part: &str,
) -> PyResult<Option<String>> {
    for index in 0..zip.len() {
        let mut entry = zip
            .by_index(index)
            .map_err(|e| PyIOError::new_err(e.to_string()))?;
        let name = entry
            .name()
            .map_err(|e| PyIOError::new_err(e.to_string()))?;
        if name.eq_ignore_ascii_case(part) {
            let mut xml = String::new();
            // Let the reader produce its existing malformed-metadata diagnostic.
            return Ok(entry.read_to_string(&mut xml).ok().map(|_| xml));
        }
    }
    Ok(None)
}

fn record_shared_formula(
    formulas: &mut HashMap<String, (Option<u64>, u64)>,
    index: Option<String>,
    len: u64,
) {
    if let Some(index) = index {
        let (master, children) = formulas.entry(index).or_default();
        if len == 0 {
            *children = children.saturating_add(1);
        } else if master.is_none() {
            *master = Some(len);
        }
    }
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
