//! SAX-based streaming sheet reader (Sprint Ι Pod-β).
//!
//! Activated by `load_workbook(path, read_only=True)` (or auto-trigger when
//! a sheet has > 50k rows). Walks `xl/worksheets/sheetN.xml` one row at a
//! time using an incremental XML reader over a bounded in-memory sheet part, or
//! a disk-spooled sheet part for larger worksheets, resolving shared-string-table
//! (SST) references upfront. Style metadata is
//! exposed as a `style_id` only — Python-side `StreamingCell` resolves the
//! actual font/fill/etc. via the eager Rust reader code path
//! (which already loads `xl/styles.xml` and exposes O(1) style lookups).
//!
//! Public surface (Python):
//!
//! - `StreamingSheetReader.open(path, sheet, ...)` — constructor.
//! - `reader.read_next_row()` → `(row_index_1based, [(col_1based, value, style_id, type), ...])`.
//! - `reader.read_next_values()` → padded value tuple.
//! - `reader.read_next_values_row()` → `(row_index_1based, padded value tuple)`.
//! - `reader.close()` — eagerly closes the XML reader and removes the temp part.
//!
//! Memory profile: SST loaded once (typically <10MB even on huge
//! workbooks); sheet XML uses a bounded in-memory fast path for small/medium
//! sheets and falls back to temp-file spooling for large sheets.

use std::fs::File;
use std::io::{BufReader, Cursor, Read, Seek, SeekFrom};

use pyo3::exceptions::{PyIOError, PyStopIteration, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList, PyTuple};
use pyo3::IntoPyObjectExt;

use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader as XmlReader;
use tempfile::NamedTempFile;
use zip::ZipArchive;

use crate::ooxml_util;

type PyObjectOwned = Py<PyAny>;
// Keep read_only value scans below openpyxl's peak RSS on ordinary table-sized
// workbooks; larger sheet XML parts stream through a temp file instead.
const IN_MEMORY_SHEET_XML_LIMIT: u64 = 1 * 1024 * 1024;

/// Parse `xl/sharedStrings.xml` into a flat `Vec<String>`. Each entry is
/// the plain-text concatenation of any nested `<r><t>...</t></r>` runs
/// (matches Excel/openpyxl's flattening for `Cell.value`).
fn load_sst(zip: &mut ZipArchive<File>) -> PyResult<Vec<String>> {
    let xml = match ooxml_util::zip_read_to_string_opt(zip, "xl/sharedStrings.xml")? {
        Some(s) => s,
        None => return Ok(Vec::new()),
    };

    let mut reader = XmlReader::from_str(&xml);
    reader.config_mut().trim_text(false);
    let mut buf: Vec<u8> = Vec::new();
    let mut out: Vec<String> = Vec::new();
    let mut current = String::new();
    let mut in_si = false;
    let mut in_t = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.local_name();
                match name.as_ref().as_bytes() {
                    b"si" => {
                        in_si = true;
                        current.clear();
                    }
                    b"t" => {
                        if in_si {
                            in_t = true;
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::End(e)) => {
                let name = e.local_name();
                match name.as_ref().as_bytes() {
                    b"si" => {
                        out.push(std::mem::take(&mut current));
                        in_si = false;
                    }
                    b"t" => {
                        in_t = false;
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(t)) => {
                if in_si && in_t {
                    current.push_str(&t);
                }
            }
            Ok(Event::GeneralRef(r)) if in_si && in_t => {
                let text = ooxml_util::resolve_general_ref(&r)
                    .map_err(|e| PyErr::new::<PyIOError, _>(format!("SST text decode: {e}")))?;
                current.push_str(&text);
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(PyErr::new::<PyIOError, _>(format!(
                    "Failed to parse sharedStrings.xml: {e}"
                )));
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(out)
}

/// Resolve `<sheet name=...>` → ZIP path (`xl/worksheets/sheetN.xml`).
fn resolve_sheet_xml_path(zip: &mut ZipArchive<File>, sheet: &str) -> PyResult<String> {
    let workbook_xml = ooxml_util::zip_read_to_string(zip, "xl/workbook.xml")?;
    let rels_xml = ooxml_util::zip_read_to_string(zip, "xl/_rels/workbook.xml.rels")?;
    let sheet_rids = ooxml_util::parse_workbook_sheet_rids(&workbook_xml)?;
    let rel_targets = ooxml_util::parse_relationship_targets(&rels_xml)?;
    for (name, rid) in sheet_rids {
        if name == sheet {
            if let Some(target) = rel_targets.get(&rid) {
                return Ok(ooxml_util::join_and_normalize("xl/", target));
            }
        }
    }
    Err(PyErr::new::<PyIOError, _>(format!(
        "Sheet not found in workbook.xml: {sheet}"
    )))
}

/// Streaming sheet reader. See module docs.
#[pyclass(unsendable, module = "wolfxl._rust")]
pub struct StreamingSheetReader {
    /// Incremental XML reader over the worksheet XML.
    reader: Option<SheetXmlReader>,
    /// Scratch buffer reused for quick-xml events.
    event_buf: Vec<u8>,
    /// Temp file that owns the decompressed sheet XML while streaming.
    temp_file: Option<NamedTempFile>,
    /// Shared-strings table loaded upfront from `xl/sharedStrings.xml`.
    sst: Vec<String>,
    /// Whether the reader has been exhausted.
    exhausted: bool,
    /// Optional `min_row` bound (1-based, inclusive). Rows below are skipped.
    min_row: Option<u32>,
    /// Optional `max_row` bound (1-based, inclusive). Iteration stops past this.
    max_row: Option<u32>,
    /// Optional `min_col` bound (1-based, inclusive).
    min_col: Option<u32>,
    /// Optional `max_col` bound (1-based, inclusive).
    max_col: Option<u32>,
}

/// One parsed cell, before Python-tuple boxing.
struct ParsedCell {
    col: u32,
    value: PyObjectOwned,
    style_id: Option<u32>,
    cell_type: &'static str,
}

#[derive(Clone, Copy)]
enum CellTypeAttr {
    Number,
    Shared,
    FormulaString,
    InlineString,
    Bool,
    Error,
    Date,
}

enum SheetXmlReader {
    Memory(XmlReader<Cursor<Vec<u8>>>),
    File(XmlReader<BufReader<File>>),
}

impl SheetXmlReader {
    fn read_event_into<'a>(&mut self, buf: &'a mut Vec<u8>) -> quick_xml::Result<Event<'a>> {
        match self {
            Self::Memory(reader) => reader.read_event_into(buf),
            Self::File(reader) => reader.read_event_into(buf),
        }
    }

    fn read_text_into<'a>(
        &mut self,
        end: quick_xml::name::QName<'_>,
        buf: &'a mut Vec<u8>,
    ) -> quick_xml::Result<quick_xml::events::BytesText<'a>> {
        match self {
            Self::Memory(reader) => reader.read_text_into(end, buf),
            Self::File(reader) => reader.read_text_into(end, buf),
        }
    }
}

#[pymethods]
impl StreamingSheetReader {
    /// Open `path` and prepare to stream `sheet`.
    #[staticmethod]
    #[pyo3(signature = (path, sheet, min_row=None, max_row=None, min_col=None, max_col=None))]
    pub fn open(
        path: &str,
        sheet: &str,
        min_row: Option<u32>,
        max_row: Option<u32>,
        min_col: Option<u32>,
        max_col: Option<u32>,
    ) -> PyResult<Self> {
        let file = File::open(path)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to open xlsx: {e}")))?;
        let mut zip = ZipArchive::new(file)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to open zip: {e}")))?;
        ooxml_util::validate_zip_archive(&mut zip)?;

        let sst = load_sst(&mut zip)?;
        let sheet_path = resolve_sheet_xml_path(&mut zip, sheet)?;
        let mut sheet_entry = zip.by_name(&sheet_path).map_err(|e| {
            PyErr::new::<PyIOError, _>(format!("Failed to open sheet part '{sheet_path}': {e}"))
        })?;
        let (reader, temp_file) = if sheet_entry.size() <= IN_MEMORY_SHEET_XML_LIMIT {
            let mut xml = Vec::with_capacity(sheet_entry.size() as usize);
            sheet_entry
                .read_to_end(&mut xml)
                .map_err(|e| PyErr::new::<PyIOError, _>(format!("read sheet XML: {e}")))?;
            let mut reader = XmlReader::from_reader(Cursor::new(xml));
            reader.config_mut().trim_text(false);
            (SheetXmlReader::Memory(reader), None)
        } else {
            let mut temp_file = NamedTempFile::new()
                .map_err(|e| PyErr::new::<PyIOError, _>(format!("streaming temp file: {e}")))?;
            std::io::copy(&mut sheet_entry, temp_file.as_file_mut())
                .map_err(|e| PyErr::new::<PyIOError, _>(format!("spool sheet XML: {e}")))?;
            temp_file
                .as_file_mut()
                .seek(SeekFrom::Start(0))
                .map_err(|e| PyErr::new::<PyIOError, _>(format!("rewind sheet XML: {e}")))?;
            let reader_file = temp_file
                .reopen()
                .map_err(|e| PyErr::new::<PyIOError, _>(format!("reopen sheet XML: {e}")))?;
            let mut reader = XmlReader::from_reader(BufReader::new(reader_file));
            reader.config_mut().trim_text(false);
            (SheetXmlReader::File(reader), Some(temp_file))
        };

        Ok(Self {
            reader: Some(reader),
            event_buf: Vec::with_capacity(8192),
            temp_file,
            sst,
            exhausted: false,
            min_row,
            max_row,
            min_col,
            max_col,
        })
    }

    /// Read the next row matching the configured bounds.
    ///
    /// Returns `None` when the stream is exhausted. Otherwise returns
    /// `(row_index_1based, [(col_1based, value, style_id_or_None, type_str), ...])`.
    pub fn read_next_row<'py>(&mut self, py: Python<'py>) -> PyResult<Option<Bound<'py, PyTuple>>> {
        if self.exhausted {
            return Ok(None);
        }
        loop {
            match self.parse_one_row(py)? {
                StepResult::Row(row_idx, cells) => {
                    let cell_list = PyList::empty(py);
                    for c in cells {
                        let style_obj = match c.style_id {
                            Some(s) => s.into_py_any(py)?,
                            None => py.None(),
                        };
                        let tup = PyTuple::new(
                            py,
                            [
                                c.col.into_py_any(py)?,
                                c.value,
                                style_obj,
                                c.cell_type.into_py_any(py)?,
                            ],
                        )?;
                        cell_list.append(tup)?;
                    }
                    let outer =
                        PyTuple::new(py, [row_idx.into_py_any(py)?, cell_list.into_py_any(py)?])?;
                    return Ok(Some(outer));
                }
                StepResult::Skip => continue,
                StepResult::Done => {
                    self.exhausted = true;
                    return Ok(None);
                }
            }
        }
    }

    /// Read the next row as a plain tuple of values, padded by column
    /// bounds. Returns `None` when exhausted.
    pub fn read_next_values<'py>(
        &mut self,
        py: Python<'py>,
    ) -> PyResult<Option<Bound<'py, PyTuple>>> {
        if self.exhausted {
            return Ok(None);
        }
        loop {
            match self.parse_one_value_row(py)? {
                ValueStepResult::Row(_row_idx, values, _needs_normalise) => {
                    return Ok(Some(values))
                }
                ValueStepResult::Skip => continue,
                ValueStepResult::Done => {
                    self.exhausted = true;
                    return Ok(None);
                }
            }
        }
    }

    /// Read the next row as `(row_index, value_tuple)` without boxing
    /// style ids into Python. Used by the values-only path when date-style
    /// conversion is unnecessary.
    pub fn read_next_values_row<'py>(
        &mut self,
        py: Python<'py>,
    ) -> PyResult<Option<Bound<'py, PyTuple>>> {
        if self.exhausted {
            return Ok(None);
        }
        loop {
            match self.parse_one_value_row(py)? {
                ValueStepResult::Row(row_idx, values, _needs_normalise) => {
                    let row_obj = row_idx.into_py_any(py)?;
                    return Ok(Some(PyTuple::new(py, [row_obj, values.into_py_any(py)?])?));
                }
                ValueStepResult::Skip => continue,
                ValueStepResult::Done => {
                    self.exhausted = true;
                    return Ok(None);
                }
            }
        }
    }

    /// Read several `(row_index, value_tuple)` records in one Python/Rust call.
    ///
    /// This keeps the public Python iterator streaming rows lazily while
    /// avoiding one FFI boundary crossing per worksheet row on values-only
    /// scans.
    #[pyo3(signature = (max_rows=1024))]
    pub fn read_next_values_chunk<'py>(
        &mut self,
        py: Python<'py>,
        max_rows: usize,
    ) -> PyResult<Option<Bound<'py, PyList>>> {
        match self.read_next_values_chunk_with_flags(py, max_rows)? {
            Some(result) => Ok(Some(result.get_item(0)?.cast_into()?)),
            None => Ok(None),
        }
    }

    /// Read a value-row chunk and report whether rows need Python normalization.
    ///
    /// The boolean is true only when the chunk contains formula/error payloads.
    /// Plain value scans can yield tuples directly without Python rescanning
    /// every row for dictionaries.
    #[pyo3(signature = (max_rows=1024))]
    pub fn read_next_values_chunk_with_flags<'py>(
        &mut self,
        py: Python<'py>,
        max_rows: usize,
    ) -> PyResult<Option<Bound<'py, PyTuple>>> {
        if self.exhausted {
            return Ok(None);
        }
        let limit = max_rows.max(1);
        let chunk = PyList::empty(py);
        let mut needs_normalise = false;
        while chunk.len() < limit {
            match self.parse_one_value_row(py)? {
                ValueStepResult::Row(row_idx, values, row_needs_normalise) => {
                    needs_normalise |= row_needs_normalise;
                    let row_obj = row_idx.into_py_any(py)?;
                    let item = PyTuple::new(py, [row_obj, values.into_py_any(py)?])?;
                    chunk.append(item)?;
                }
                ValueStepResult::Skip => continue,
                ValueStepResult::Done => {
                    self.exhausted = true;
                    break;
                }
            }
        }
        if chunk.is_empty() {
            Ok(None)
        } else {
            Ok(Some(PyTuple::new(
                py,
                [chunk.into_py_any(py)?, needs_normalise.into_py_any(py)?],
            )?))
        }
    }

    /// Read a value-row chunk with a compact dense-table representation.
    ///
    /// Dense plain chunks return only value tuples plus their first/last row
    /// indexes. Sparse chunks, or chunks needing Python formula/error
    /// normalization, fall back to the row-indexed record shape used by
    /// `read_next_values_chunk_with_flags`.
    #[pyo3(signature = (max_rows=1024))]
    pub fn read_next_values_compact_chunk_with_flags<'py>(
        &mut self,
        py: Python<'py>,
        max_rows: usize,
    ) -> PyResult<Option<Bound<'py, PyTuple>>> {
        if self.exhausted {
            return Ok(None);
        }
        let limit = max_rows.max(1);
        let mut rows: Vec<(u32, Bound<'py, PyTuple>)> = Vec::with_capacity(limit);
        let mut needs_normalise = false;
        let mut first_row = 0;
        let mut last_row = 0;
        let mut contiguous = true;

        while rows.len() < limit {
            match self.parse_one_value_row(py)? {
                ValueStepResult::Row(row_idx, values, row_needs_normalise) => {
                    if rows.is_empty() {
                        first_row = row_idx;
                    } else if row_idx != last_row + 1 {
                        contiguous = false;
                    }
                    last_row = row_idx;
                    needs_normalise |= row_needs_normalise;
                    rows.push((row_idx, values));
                }
                ValueStepResult::Skip => continue,
                ValueStepResult::Done => {
                    self.exhausted = true;
                    break;
                }
            }
        }

        if rows.is_empty() {
            return Ok(None);
        }

        let dense_plain = contiguous && !needs_normalise;
        let items = PyList::empty(py);
        if dense_plain {
            for (_row_idx, values) in rows {
                items.append(values)?;
            }
        } else {
            for (row_idx, values) in rows {
                let row_obj = row_idx.into_py_any(py)?;
                items.append(PyTuple::new(py, [row_obj, values.into_py_any(py)?])?)?;
            }
        }

        Ok(Some(PyTuple::new(
            py,
            [
                items.into_py_any(py)?,
                needs_normalise.into_py_any(py)?,
                first_row.into_py_any(py)?,
                last_row.into_py_any(py)?,
                dense_plain.into_py_any(py)?,
            ],
        )?))
    }

    /// True once the stream has been fully consumed.
    pub fn is_exhausted(&self) -> bool {
        self.exhausted
    }

    /// Eagerly drop the in-memory XML buffer (releases peak RSS).
    pub fn close(&mut self) {
        self.exhausted = true;
        self.reader = None;
        self.temp_file = None;
        self.event_buf.clear();
    }

    fn __iter__(slf: PyRef<'_, Self>) -> PyRef<'_, Self> {
        slf
    }

    fn __next__<'py>(
        mut slf: PyRefMut<'_, Self>,
        py: Python<'py>,
    ) -> PyResult<Bound<'py, PyTuple>> {
        match slf.read_next_row(py)? {
            Some(t) => Ok(t),
            None => Err(PyStopIteration::new_err(())),
        }
    }
}

enum StepResult {
    Row(u32, Vec<ParsedCell>),
    Skip,
    Done,
}

enum ValueStepResult<'py> {
    Row(u32, Bound<'py, PyTuple>, bool),
    Skip,
    Done,
}

impl StreamingSheetReader {
    fn parse_one_row(&mut self, py: Python<'_>) -> PyResult<StepResult> {
        let reader = match self.reader.as_mut() {
            Some(reader) => reader,
            None => return Ok(StepResult::Done),
        };

        loop {
            self.event_buf.clear();
            let event = reader
                .read_event_into(&mut self.event_buf)
                .map_err(|e| PyErr::new::<PyIOError, _>(format!("Streaming reader XML: {e}")))?;
            match event {
                Event::Start(e) if e.local_name().as_ref() == "row" => {
                    let row_idx = parse_row_index_from_start(&e)?;
                    drop(e);
                    let cells = read_cells_until_row_end(
                        reader,
                        &mut self.event_buf,
                        py,
                        &self.sst,
                        row_idx,
                    )?;
                    return self.row_result(row_idx, cells);
                }
                Event::Empty(e) if e.local_name().as_ref() == "row" => {
                    let row_idx = parse_row_index_from_start(&e)?;
                    return self.row_result(row_idx, Vec::new());
                }
                Event::Eof => return Ok(StepResult::Done),
                _ => {}
            }
        }
    }

    fn parse_one_value_row<'py>(&mut self, py: Python<'py>) -> PyResult<ValueStepResult<'py>> {
        let reader = match self.reader.as_mut() {
            Some(reader) => reader,
            None => return Ok(ValueStepResult::Done),
        };

        loop {
            self.event_buf.clear();
            let event = reader
                .read_event_into(&mut self.event_buf)
                .map_err(|e| PyErr::new::<PyIOError, _>(format!("Streaming reader XML: {e}")))?;
            match event {
                Event::Start(e) if e.local_name().as_ref() == "row" => {
                    let row_idx = parse_row_index_from_start(&e)?;
                    drop(e);
                    if let Some(min) = self.min_row {
                        if row_idx < min {
                            skip_row_until_end(reader, &mut self.event_buf, row_idx)?;
                            return Ok(ValueStepResult::Skip);
                        }
                    }
                    if let Some(max) = self.max_row {
                        if row_idx > max {
                            self.exhausted = true;
                            return Ok(ValueStepResult::Done);
                        }
                    }
                    let (values, needs_normalise) = read_values_until_row_end(
                        reader,
                        &mut self.event_buf,
                        py,
                        &self.sst,
                        row_idx,
                        self.min_col,
                        self.max_col,
                    )?;
                    return Ok(ValueStepResult::Row(row_idx, values, needs_normalise));
                }
                Event::Empty(e) if e.local_name().as_ref() == "row" => {
                    let row_idx = parse_row_index_from_start(&e)?;
                    return self.empty_value_row_result(py, row_idx);
                }
                Event::Eof => return Ok(ValueStepResult::Done),
                _ => {}
            }
        }
    }

    fn empty_value_row_result<'py>(
        &mut self,
        py: Python<'py>,
        row_idx: u32,
    ) -> PyResult<ValueStepResult<'py>> {
        if let Some(min) = self.min_row {
            if row_idx < min {
                return Ok(ValueStepResult::Skip);
            }
        }
        if let Some(max) = self.max_row {
            if row_idx > max {
                self.exhausted = true;
                return Ok(ValueStepResult::Done);
            }
        }
        Ok(ValueStepResult::Row(
            row_idx,
            empty_value_row(py, self.min_col, self.max_col)?,
            false,
        ))
    }

    fn row_result(&mut self, row_idx: u32, mut cells: Vec<ParsedCell>) -> PyResult<StepResult> {
        if let Some(min) = self.min_row {
            if row_idx < min {
                return Ok(StepResult::Skip);
            }
        }
        if let Some(max) = self.max_row {
            if row_idx > max {
                self.exhausted = true;
                return Ok(StepResult::Done);
            }
        }

        if let Some(cmin) = self.min_col {
            cells.retain(|c| c.col >= cmin);
        }
        if let Some(cmax) = self.max_col {
            cells.retain(|c| c.col <= cmax);
        }

        Ok(StepResult::Row(row_idx, cells))
    }
}

fn read_cells_until_row_end(
    reader: &mut SheetXmlReader,
    buf: &mut Vec<u8>,
    py: Python<'_>,
    sst: &[String],
    row_idx: u32,
) -> PyResult<Vec<ParsedCell>> {
    let mut cells: Vec<ParsedCell> = Vec::new();

    loop {
        buf.clear();
        let event = reader
            .read_event_into(buf)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Streaming reader XML: {e}")))?;
        match event {
            Event::Start(e) if e.local_name().as_ref() == "c" => {
                let (col, style_id, t_attr) = cell_attrs_from_start(&e, &cells)?;
                drop(e);
                let (value, cell_type) = read_cell_contents(reader, buf, py, &t_attr, sst)?;
                cells.push(ParsedCell {
                    col,
                    value,
                    style_id,
                    cell_type,
                });
            }
            Event::Empty(e) if e.local_name().as_ref() == "c" => {
                let (col, style_id, _t_attr) = cell_attrs_from_start(&e, &cells)?;
                cells.push(ParsedCell {
                    col,
                    value: py.None(),
                    style_id,
                    cell_type: "blank",
                });
            }
            Event::End(e) if e.local_name().as_ref() == "row" => return Ok(cells),
            Event::Eof => {
                return Err(PyErr::new::<PyIOError, _>(format!(
                    "Streaming reader: unterminated <row r=\"{row_idx}\">"
                )));
            }
            _ => {}
        }
    }
}

fn skip_row_until_end(
    reader: &mut SheetXmlReader,
    buf: &mut Vec<u8>,
    row_idx: u32,
) -> PyResult<()> {
    loop {
        buf.clear();
        let event = reader
            .read_event_into(buf)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Streaming reader XML: {e}")))?;
        match event {
            Event::End(e) if e.local_name().as_ref() == "row" => return Ok(()),
            Event::Eof => {
                return Err(PyErr::new::<PyIOError, _>(format!(
                    "Streaming reader: unterminated <row r=\"{row_idx}\">"
                )));
            }
            _ => {}
        }
    }
}

fn read_values_until_row_end<'py>(
    reader: &mut SheetXmlReader,
    buf: &mut Vec<u8>,
    py: Python<'py>,
    sst: &[String],
    row_idx: u32,
    min_col: Option<u32>,
    max_col: Option<u32>,
) -> PyResult<(Bound<'py, PyTuple>, bool)> {
    let cmin = min_col.unwrap_or(1);
    let mut out: Vec<PyObjectOwned> = Vec::new();
    let mut last_col = cmin.saturating_sub(1);
    let mut next_col = 1;
    let mut needs_normalise = false;

    loop {
        buf.clear();
        let event = reader
            .read_event_into(buf)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Streaming reader XML: {e}")))?;
        match event {
            Event::Start(e) if e.local_name().as_ref() == "c" => {
                let (col, t_attr) = cell_value_attrs_from_start(&e, next_col)?;
                next_col = col.saturating_add(1);
                drop(e);
                let (value, cell_type) = read_cell_contents(reader, buf, py, &t_attr, sst)?;
                let pushed =
                    push_bounded_value(py, &mut out, &mut last_col, col, value, min_col, max_col);
                if pushed && matches!(cell_type, "formula" | "e") {
                    needs_normalise = true;
                }
            }
            Event::Empty(e) if e.local_name().as_ref() == "c" => {
                let (col, _t_attr) = cell_value_attrs_from_start(&e, next_col)?;
                next_col = col.saturating_add(1);
                push_bounded_value(
                    py,
                    &mut out,
                    &mut last_col,
                    col,
                    py.None(),
                    min_col,
                    max_col,
                );
            }
            Event::End(e) if e.local_name().as_ref() == "row" => {
                return Ok((
                    finish_value_row(py, out, last_col, min_col, max_col)?,
                    needs_normalise,
                ));
            }
            Event::Eof => {
                return Err(PyErr::new::<PyIOError, _>(format!(
                    "Streaming reader: unterminated <row r=\"{row_idx}\">"
                )));
            }
            _ => {}
        }
    }
}

fn push_bounded_value(
    py: Python<'_>,
    out: &mut Vec<PyObjectOwned>,
    last_col: &mut u32,
    col: u32,
    value: PyObjectOwned,
    min_col: Option<u32>,
    max_col: Option<u32>,
) -> bool {
    let cmin = min_col.unwrap_or(1);
    if col < cmin || max_col.is_some_and(|cmax| col > cmax) || col <= *last_col {
        return false;
    }
    for _ in last_col.saturating_add(1)..col {
        out.push(py.None());
    }
    out.push(value);
    *last_col = col;
    true
}

fn finish_value_row<'py>(
    py: Python<'py>,
    mut out: Vec<PyObjectOwned>,
    last_col: u32,
    min_col: Option<u32>,
    max_col: Option<u32>,
) -> PyResult<Bound<'py, PyTuple>> {
    let cmin = min_col.unwrap_or(1);
    if let Some(cmax) = max_col {
        if cmax < cmin {
            return Ok(PyTuple::empty(py));
        }
        for _ in last_col.saturating_add(1)..=cmax {
            out.push(py.None());
        }
    } else if min_col.is_some() && out.is_empty() {
        out.push(py.None());
    }
    Ok(PyTuple::new(py, out)?)
}

fn empty_value_row<'py>(
    py: Python<'py>,
    min_col: Option<u32>,
    max_col: Option<u32>,
) -> PyResult<Bound<'py, PyTuple>> {
    let cmin = min_col.unwrap_or(1);
    finish_value_row(py, Vec::new(), cmin.saturating_sub(1), min_col, max_col)
}

fn read_cell_contents(
    reader: &mut SheetXmlReader,
    buf: &mut Vec<u8>,
    py: Python<'_>,
    t_attr: &CellTypeAttr,
    sst: &[String],
) -> PyResult<(PyObjectOwned, &'static str)> {
    let mut v_text: Option<String> = None;
    let mut f_text: Option<String> = None;
    let mut inline_text: Option<String> = None;
    let mut fast_number: Option<PyObjectOwned> = None;
    let mut fast_shared: Option<usize> = None;
    let mut in_f = false;
    let mut in_t = false;

    loop {
        buf.clear();
        let event = reader
            .read_event_into(buf)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Streaming reader XML: {e}")))?;
        match event {
            Event::Start(e) => match e.local_name().as_ref().as_bytes() {
                b"v" => {
                    // Read the complete value before parsing: `1&#50;3` is
                    // 123, not the first Text fragment (1). Reuse the XML
                    // buffer so ordinary numeric values remain allocation-free.
                    let end = if e.name().as_ref() == "v" {
                        std::borrow::Cow::Borrowed("v")
                    } else {
                        std::borrow::Cow::Owned(e.name().as_ref().to_owned())
                    };
                    drop(e);
                    buf.clear();
                    let raw = reader
                        .read_text_into(quick_xml::name::QName(&end), buf)
                        .map_err(|e| {
                            PyErr::new::<PyIOError, _>(format!("Streaming reader XML: {e}"))
                        })?;
                    let text = quick_xml::escape::unescape(&raw).map_err(|e| {
                        PyErr::new::<PyIOError, _>(format!("cell text decode: {e}"))
                    })?;
                    if matches!(t_attr, CellTypeAttr::Number)
                        && f_text.is_none()
                        && v_text.is_none()
                        && fast_number.is_none()
                    {
                        if let Some(value) = number_text_to_py(py, text.as_bytes())? {
                            fast_number = Some(value);
                            continue;
                        }
                    }
                    if matches!(t_attr, CellTypeAttr::Shared)
                        && v_text.is_none()
                        && fast_shared.is_none()
                    {
                        if let Some(idx) = parse_usize_bytes(text.as_bytes()) {
                            fast_shared = Some(idx);
                            continue;
                        }
                    }
                    if !text.is_empty() {
                        v_text.get_or_insert_with(String::new).push_str(&text);
                    }
                }
                b"f" => in_f = true,
                b"t" => in_t = true,
                _ => {}
            },
            Event::Empty(e) => match e.local_name().as_ref().as_bytes() {
                b"v" => {
                    v_text.get_or_insert_with(String::new);
                }
                b"f" => {
                    f_text.get_or_insert_with(String::new);
                }
                b"t" => {
                    inline_text.get_or_insert_with(String::new);
                }
                _ => {}
            },
            Event::Text(t) => {
                if in_f {
                    f_text.get_or_insert_with(String::new).push_str(&t);
                } else if in_t {
                    inline_text.get_or_insert_with(String::new).push_str(&t);
                }
            }
            Event::GeneralRef(r) => {
                let text = ooxml_util::resolve_general_ref(&r)
                    .map_err(|e| PyErr::new::<PyIOError, _>(format!("cell text decode: {e}")))?;
                if in_f {
                    f_text.get_or_insert_with(String::new).push_str(&text);
                } else if in_t {
                    inline_text.get_or_insert_with(String::new).push_str(&text);
                }
            }
            Event::End(e) => match e.local_name().as_ref().as_bytes() {
                b"c" => {
                    if f_text.is_none() {
                        if let Some(value) = fast_number {
                            return Ok((value, "n"));
                        }
                        if let Some(idx) = fast_shared {
                            let resolved = sst.get(idx).map_or("", String::as_str);
                            return Ok((resolved.into_py_any(py)?, "s"));
                        }
                    } else if let Some(idx) = fast_shared {
                        let formula = f_text.unwrap_or_default();
                        let resolved = sst.get(idx).map_or("", String::as_str);
                        return Ok((build_formula_dict(py, &formula, resolved)?, "formula"));
                    }
                    return build_cell_object(py, t_attr, f_text, v_text.or(inline_text), sst);
                }
                b"f" => in_f = false,
                b"t" => in_t = false,
                _ => {}
            },
            Event::Eof => {
                return Err(PyErr::new::<PyIOError, _>(
                    "Streaming reader: unterminated <c>".to_string(),
                ));
            }
            _ => {}
        }
    }
}

fn cell_attrs_from_start(
    e: &BytesStart<'_>,
    cells: &[ParsedCell],
) -> PyResult<(u32, Option<u32>, CellTypeAttr)> {
    let mut col = None;
    let mut style_id = None;
    let mut cell_type = CellTypeAttr::Number;
    for attr in e.attributes().with_checks(false).flatten() {
        let key = attr.key.as_ref().as_bytes();
        let local = key.rsplit(|b| *b == b':').next().unwrap_or(key);
        match local {
            b"r" => {
                col = Some(
                    parse_a1_column(attr.value.as_bytes())
                        .map_err(PyErr::new::<PyValueError, _>)?,
                );
            }
            b"s" => style_id = parse_u32_bytes(attr.value.as_bytes()),
            b"t" => cell_type = cell_type_from_bytes(attr.value.as_bytes()),
            _ => {}
        }
    }
    let col = col.unwrap_or_else(|| cells.last().map(|c| c.col + 1).unwrap_or(1));
    Ok((col, style_id, cell_type))
}

fn cell_value_attrs_from_start(
    e: &BytesStart<'_>,
    default_col: u32,
) -> PyResult<(u32, CellTypeAttr)> {
    let mut col = None;
    let mut cell_type = CellTypeAttr::Number;
    for attr in e.attributes().with_checks(false).flatten() {
        let key = attr.key.as_ref().as_bytes();
        let local = key.rsplit(|b| *b == b':').next().unwrap_or(key);
        match local {
            b"r" => {
                col = Some(
                    parse_a1_column(attr.value.as_bytes())
                        .map_err(PyErr::new::<PyValueError, _>)?,
                );
            }
            b"t" => cell_type = cell_type_from_bytes(attr.value.as_bytes()),
            _ => {}
        }
    }
    Ok((col.unwrap_or(default_col), cell_type))
}

fn parse_row_index_from_start(e: &BytesStart<'_>) -> PyResult<u32> {
    match attr_u32(e, b"r") {
        Some(row) => Ok(row),
        None => Err(PyErr::new::<PyIOError, _>(
            "Streaming reader: <row> missing r= attribute".to_string(),
        )),
    }
}

fn attr_u32(e: &BytesStart<'_>, name: &[u8]) -> Option<u32> {
    for attr in e.attributes().with_checks(false).flatten() {
        let key = attr.key.as_ref().as_bytes();
        let local = key.rsplit(|b| *b == b':').next().unwrap_or(key);
        if local == name {
            return parse_u32_bytes(attr.value.as_bytes());
        }
    }
    None
}

fn cell_type_from_bytes(value: &[u8]) -> CellTypeAttr {
    match value {
        b"s" => CellTypeAttr::Shared,
        b"str" => CellTypeAttr::FormulaString,
        b"inlineStr" => CellTypeAttr::InlineString,
        b"b" => CellTypeAttr::Bool,
        b"e" => CellTypeAttr::Error,
        b"d" => CellTypeAttr::Date,
        _ => CellTypeAttr::Number,
    }
}

fn parse_u32_bytes(bytes: &[u8]) -> Option<u32> {
    if bytes.is_empty() {
        return None;
    }
    let mut value = 0u32;
    for &byte in bytes {
        if !byte.is_ascii_digit() {
            return None;
        }
        value = value.checked_mul(10)?.checked_add((byte - b'0') as u32)?;
    }
    Some(value)
}

fn parse_usize_bytes(bytes: &[u8]) -> Option<usize> {
    if bytes.is_empty() {
        return None;
    }
    let mut value = 0usize;
    for &byte in bytes {
        if !byte.is_ascii_digit() {
            return None;
        }
        value = value.checked_mul(10)?.checked_add((byte - b'0') as usize)?;
    }
    Some(value)
}

fn parse_a1_column(bytes: &[u8]) -> Result<u32, String> {
    let mut col = 0u32;
    let mut saw_letter = false;
    for &byte in bytes {
        if byte == b'$' {
            continue;
        }
        let upper = byte.to_ascii_uppercase();
        if upper.is_ascii_uppercase() {
            saw_letter = true;
            col = col
                .checked_mul(26)
                .and_then(|value| value.checked_add((upper - b'A' + 1) as u32))
                .ok_or_else(|| {
                    format!("Bad cell coordinate: {:?}", String::from_utf8_lossy(bytes))
                })?;
            continue;
        }
        if byte.is_ascii_digit() {
            break;
        }
        return Err(format!(
            "Bad cell coordinate: {:?}",
            String::from_utf8_lossy(bytes)
        ));
    }
    if saw_letter {
        Ok(col)
    } else {
        Err(format!(
            "Bad cell coordinate: {:?}",
            String::from_utf8_lossy(bytes)
        ))
    }
}

// ----------------------------------------------------------------------
// Low-level XML scanning helpers.
// ----------------------------------------------------------------------

/// Return the index of the next `<TAG` (followed by space / `>` / `/`).
#[cfg(test)]
fn find_tag_open(bytes: &[u8], start: usize, tag: &[u8]) -> Option<usize> {
    let needle_len = 1 + tag.len();
    let mut i = start;
    while i + needle_len < bytes.len() {
        if bytes[i] == b'<' && bytes[i + 1..].starts_with(tag) {
            let after = bytes[i + 1 + tag.len()];
            if matches!(after, b' ' | b'>' | b'/' | b'\t' | b'\n' | b'\r') {
                return Some(i);
            }
        }
        i += 1;
    }
    None
}

/// Return the index of the next `</TAG>` token.
#[cfg(test)]
fn find_close_tag(bytes: &[u8], start: usize, tag: &[u8]) -> Option<usize> {
    let needle_len = 3 + tag.len();
    let mut i = start;
    while i + needle_len <= bytes.len() {
        if bytes[i] == b'<' && bytes[i + 1] == b'/' && bytes[i + 2..].starts_with(tag) {
            let after = bytes[i + 2 + tag.len()];
            if matches!(after, b'>' | b' ' | b'\t') {
                return Some(i);
            }
        }
        i += 1;
    }
    None
}

/// Parse the open tag starting at `bytes[start]` (which should be `<`).
/// Returns `(attrs_substring, byte_index_after_open_tag, self_closing)`.
#[cfg(test)]
fn read_open_tag(bytes: &[u8], start: usize) -> PyResult<(&str, usize, bool)> {
    debug_assert_eq!(bytes.get(start), Some(&b'<'));
    let mut i = start + 1;
    let mut in_quote: Option<u8> = None;
    while i < bytes.len() {
        let b = bytes[i];
        if let Some(q) = in_quote {
            if b == q {
                in_quote = None;
            }
        } else if b == b'"' || b == b'\'' {
            in_quote = Some(b);
        } else if b == b'>' {
            let self_closing = i > 0 && bytes[i - 1] == b'/';
            let attrs_end = if self_closing { i - 1 } else { i };
            let mut name_end = start + 1;
            while name_end < bytes.len()
                && !matches!(bytes[name_end], b' ' | b'/' | b'>' | b'\t' | b'\n' | b'\r')
            {
                name_end += 1;
            }
            let attrs_slice = &bytes[name_end..attrs_end];
            let attrs_str = std::str::from_utf8(attrs_slice).map_err(|e| {
                PyErr::new::<PyIOError, _>(format!("Streaming reader: invalid UTF-8 in attrs: {e}"))
            })?;
            return Ok((attrs_str, i + 1, self_closing));
        }
        i += 1;
    }
    Err(PyErr::new::<PyIOError, _>(
        "Streaming reader: unterminated open tag".to_string(),
    ))
}

/// Parse `r="A1"` from an attribute substring.
#[cfg(test)]
fn read_attr(attrs: &str, name: &str) -> Option<String> {
    // Build candidate offsets for `<sep>name=`.
    let mut idx: Option<usize> = None;
    let prefix = format!("{name}=");
    if attrs.starts_with(&prefix) {
        idx = Some(0);
    }
    if idx.is_none() {
        for sep in [' ', '\t', '\n', '\r'] {
            let needle: String = format!("{sep}{name}=");
            if let Some(p) = attrs.find(&needle) {
                idx = Some(p + 1);
                break;
            }
        }
    }
    let i = idx?;
    let after_eq = i + name.len() + 1;
    let bytes = attrs.as_bytes();
    if after_eq >= bytes.len() {
        return None;
    }
    let quote = bytes[after_eq];
    if quote != b'"' && quote != b'\'' {
        return None;
    }
    let value_start = after_eq + 1;
    let mut j = value_start;
    while j < bytes.len() && bytes[j] != quote {
        j += 1;
    }
    Some(
        std::str::from_utf8(&bytes[value_start..j])
            .ok()?
            .to_string(),
    )
}

#[cfg(test)]
fn parse_row_index(attrs: &str) -> PyResult<u32> {
    match read_attr(attrs, "r") {
        Some(s) => s.parse().map_err(|_| {
            PyErr::new::<PyIOError, _>(format!("Streaming reader: bad row index {s:?}"))
        }),
        None => Err(PyErr::new::<PyIOError, _>(
            "Streaming reader: <row> missing r= attribute".to_string(),
        )),
    }
}

#[cfg(test)]
fn parse_cell_inner(
    py: Python<'_>,
    inner: &[u8],
    t_attr: &str,
    sst: &[String],
) -> PyResult<(PyObjectOwned, &'static str)> {
    let v_text = extract_inner_text(inner, b"v");
    let f_text = extract_inner_text(inner, b"f");
    let is_text = extract_is_text(inner);
    let t_attr = cell_type_from_bytes(t_attr.as_bytes());

    build_cell_object(py, &t_attr, f_text, v_text.or(is_text), sst)
}

fn build_cell_object(
    py: Python<'_>,
    t_attr: &CellTypeAttr,
    formula: Option<String>,
    raw_value: Option<String>,
    sst: &[String],
) -> PyResult<(PyObjectOwned, &'static str)> {
    match t_attr {
        CellTypeAttr::Shared => {
            let v = raw_value.unwrap_or_default();
            let idx: usize = v
                .parse()
                .map_err(|_| PyErr::new::<PyValueError, _>(format!("Bad SST index: {v:?}")))?;
            let resolved = sst.get(idx).map_or("", String::as_str);
            if let Some(formula) = formula {
                return Ok((build_formula_dict(py, &formula, resolved)?, "formula"));
            }
            Ok((resolved.into_py_any(py)?, "s"))
        }
        CellTypeAttr::FormulaString => {
            let v = raw_value.unwrap_or_default();
            if let Some(formula) = formula {
                return Ok((build_formula_dict(py, &formula, &v)?, "formula"));
            }
            Ok((v.into_py_any(py)?, "str"))
        }
        CellTypeAttr::InlineString => {
            let v = raw_value.unwrap_or_default();
            Ok((v.into_py_any(py)?, "inlineStr"))
        }
        CellTypeAttr::Bool => {
            let v = raw_value.unwrap_or_default();
            let b = matches!(v.trim(), "1" | "true" | "TRUE");
            if let Some(formula) = formula {
                return Ok((build_formula_dict(py, &formula, &v)?, "formula"));
            }
            Ok((b.into_py_any(py)?, "b"))
        }
        CellTypeAttr::Error => {
            let v = raw_value.unwrap_or_else(|| "#ERROR!".to_string());
            let d = PyDict::new(py);
            d.set_item("type", "error")?;
            d.set_item("value", &v)?;
            Ok((d.into_py_any(py)?, "e"))
        }
        CellTypeAttr::Date => {
            let v = raw_value.unwrap_or_default();
            Ok((v.into_py_any(py)?, "d"))
        }
        CellTypeAttr::Number => {
            if let Some(formula) = formula {
                let cached = raw_value.unwrap_or_default();
                return Ok((build_formula_dict(py, &formula, &cached)?, "formula"));
            }
            let v = match raw_value {
                Some(s) => s,
                None => return Ok((py.None(), "blank")),
            };
            if let Ok(i) = v.parse::<i64>() {
                // Excel-stored ints: surface as int when round-trip safe.
                return Ok((i.into_py_any(py)?, "n"));
            }
            let f: f64 = v
                .parse()
                .map_err(|_| PyErr::new::<PyValueError, _>(format!("Bad numeric value: {v:?}")))?;
            Ok((f.into_py_any(py)?, "n"))
        }
    }
}

fn build_formula_dict(py: Python<'_>, formula: &str, cached: &str) -> PyResult<PyObjectOwned> {
    let d = PyDict::new(py);
    let f_with_eq = if formula.starts_with('=') {
        formula.to_string()
    } else {
        format!("={formula}")
    };
    d.set_item("type", "formula")?;
    d.set_item("formula", &f_with_eq)?;
    d.set_item("cached", cached)?;
    Ok(d.into_py_any(py)?)
}

fn number_text_to_py(py: Python<'_>, bytes: &[u8]) -> PyResult<Option<PyObjectOwned>> {
    let Ok(text) = std::str::from_utf8(bytes) else {
        return Ok(None);
    };
    if let Ok(i) = text.parse::<i64>() {
        return Ok(Some(i.into_py_any(py)?));
    }
    if let Ok(f) = text.parse::<f64>() {
        return Ok(Some(f.into_py_any(py)?));
    }
    Ok(None)
}

#[cfg(test)]
fn extract_inner_text(inner: &[u8], tag: &[u8]) -> Option<String> {
    let open = find_tag_open(inner, 0, tag)?;
    let (_, after_open, self_closing) = read_open_tag(inner, open).ok()?;
    if self_closing {
        return Some(String::new());
    }
    let close = find_close_tag(inner, after_open, tag)?;
    let raw = &inner[after_open..close];
    Some(unescape_xml(raw))
}

#[cfg(test)]
fn extract_is_text(inner: &[u8]) -> Option<String> {
    let is_open = find_tag_open(inner, 0, b"is")?;
    let (_, after_is, _) = read_open_tag(inner, is_open).ok()?;
    let is_close = find_close_tag(inner, after_is, b"is")?;
    let is_inner = &inner[after_is..is_close];
    let mut out = String::new();
    let mut p = 0;
    while p < is_inner.len() {
        let t_open = match find_tag_open(is_inner, p, b"t") {
            Some(i) => i,
            None => break,
        };
        let (_, after_t, t_self) = read_open_tag(is_inner, t_open).ok()?;
        if t_self {
            p = after_t;
            continue;
        }
        let t_close = find_close_tag(is_inner, after_t, b"t")?;
        out.push_str(&unescape_xml(&is_inner[after_t..t_close]));
        p = t_close + b"</t>".len();
    }
    Some(out)
}

#[cfg(test)]
fn unescape_xml(b: &[u8]) -> String {
    if !b.contains(&b'&') {
        return String::from_utf8_lossy(b).into_owned();
    }
    let s = String::from_utf8_lossy(b);
    s.replace("&amp;", "&")
        .replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", "\"")
        .replace("&apos;", "'")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn find_tag_basic() {
        let xml = b"<row r=\"1\"><c r=\"A1\"><v>3</v></c></row>";
        assert_eq!(find_tag_open(xml, 0, b"row"), Some(0));
        assert_eq!(find_tag_open(xml, 0, b"c"), Some(11));
        assert_eq!(find_tag_open(xml, 0, b"v"), Some(21));
    }

    #[test]
    fn open_tag_attrs() {
        let xml = b"<c r=\"A1\" s=\"3\" t=\"s\"><v>0</v></c>";
        let (attrs, _end, sc) = read_open_tag(xml, 0).unwrap();
        assert!(!sc);
        assert!(attrs.contains("r=\"A1\""));
        assert_eq!(read_attr(attrs, "r").as_deref(), Some("A1"));
        assert_eq!(read_attr(attrs, "s").as_deref(), Some("3"));
        assert_eq!(read_attr(attrs, "t").as_deref(), Some("s"));
    }

    #[test]
    fn open_tag_self_closing() {
        let xml = b"<c r=\"A1\" s=\"3\"/>";
        let (_attrs, _end, sc) = read_open_tag(xml, 0).unwrap();
        assert!(sc);
    }

    #[test]
    fn close_tag_position() {
        let xml = b"<c r=\"A1\"><v>3</v></c>more";
        let close = find_close_tag(xml, 0, b"c").unwrap();
        assert_eq!(&xml[close..close + 4], b"</c>");
    }

    #[test]
    fn extract_v_text() {
        let inner = b"<v>3.14</v>";
        assert_eq!(extract_inner_text(inner, b"v").as_deref(), Some("3.14"));
    }

    #[test]
    fn extract_is_inline() {
        let inner = b"<is><t>hello</t></is>";
        assert_eq!(extract_is_text(inner).as_deref(), Some("hello"));
    }

    #[test]
    fn parse_row_index_basic() {
        assert_eq!(parse_row_index(" r=\"12\"").unwrap(), 12);
        assert!(parse_row_index(" r=\"bad\"").is_err());
    }

    #[test]
    fn parse_a1_column_basic() {
        assert_eq!(parse_a1_column(b"A1").unwrap(), 1);
        assert_eq!(parse_a1_column(b"Z99").unwrap(), 26);
        assert_eq!(parse_a1_column(b"AA10").unwrap(), 27);
        assert_eq!(parse_a1_column(b"$XFD$1").unwrap(), 16384);
        assert!(parse_a1_column(b"123").is_err());
    }

    #[test]
    fn parse_cell_inner_number() {
        pyo3::Python::initialize();
        Python::attach(|py| {
            let (obj, kind) = parse_cell_inner(py, b"<v>42</v>", "n", &[]).unwrap();
            let value: i64 = obj.extract(py).unwrap();
            assert_eq!(kind, "n");
            assert_eq!(value, 42);
        });
    }

    #[test]
    fn number_text_to_py_basic() {
        pyo3::Python::initialize();
        Python::attach(|py| {
            let int_obj = number_text_to_py(py, b"42").unwrap().unwrap();
            let int_value: i64 = int_obj.extract(py).unwrap();
            assert_eq!(int_value, 42);

            let float_obj = number_text_to_py(py, b"3.5").unwrap().unwrap();
            let float_value: f64 = float_obj.extract(py).unwrap();
            assert_eq!(float_value, 3.5);

            assert!(number_text_to_py(py, b"bad").unwrap().is_none());
        });
    }

    #[test]
    fn unescape_basic() {
        assert_eq!(unescape_xml(b"a &amp; b"), "a & b");
        assert_eq!(unescape_xml(b"plain"), "plain");
    }

    #[test]
    fn streaming_split_references_are_complete_values() {
        pyo3::Python::initialize();
        Python::attach(|py| {
            for use_file in [false, true] {
                for (xml, kind, expected) in [
                    ("<c><v>1&#50;3</v></c>", CellTypeAttr::Number, "123"),
                    ("<c><v>&#49;0</v></c>", CellTypeAttr::Shared, "entry 10"),
                    (
                        "<x:c xmlns:x=\"urn:test\"><x:v>1&#x32;3</x:v></x:c>",
                        CellTypeAttr::Number,
                        "123",
                    ),
                    (
                        "<c><is><t> A &amp; &#x1F600;&#10; B </t></is></c>",
                        CellTypeAttr::InlineString,
                        " A & 😀\n B ",
                    ),
                ] {
                    let mut file = tempfile::tempfile().unwrap();
                    let mut reader = if use_file {
                        std::io::Write::write_all(&mut file, xml.as_bytes()).unwrap();
                        file.rewind().unwrap();
                        SheetXmlReader::File(XmlReader::from_reader(BufReader::new(file)))
                    } else {
                        SheetXmlReader::Memory(XmlReader::from_reader(Cursor::new(
                            xml.as_bytes().to_vec(),
                        )))
                    };
                    let mut buf = Vec::new();
                    assert!(matches!(
                        reader.read_event_into(&mut buf).unwrap(),
                        Event::Start(_)
                    ));
                    let sst: Vec<_> = (0..11).map(|i| format!("entry {i}")).collect();
                    let (value, _) =
                        read_cell_contents(&mut reader, &mut buf, py, &kind, &sst).unwrap();
                    if matches!(kind, CellTypeAttr::Number) {
                        assert_eq!(value.extract::<i64>(py).unwrap().to_string(), expected);
                    } else {
                        assert_eq!(value.extract::<String>(py).unwrap(), expected);
                    }
                }
            }
        });
    }

    #[test]
    fn streaming_rejects_invalid_value_references() {
        pyo3::Python::initialize();
        Python::attach(|py| {
            for body in [
                "<v>1&unknown;2</v>",
                "<is><t>before &unknown; after</t></is>",
            ] {
                let xml = format!("<c>{body}</c>");
                let mut reader =
                    SheetXmlReader::Memory(XmlReader::from_reader(Cursor::new(xml.into_bytes())));
                let mut buf = Vec::new();
                assert!(matches!(
                    reader.read_event_into(&mut buf).unwrap(),
                    Event::Start(_)
                ));
                assert!(
                    read_cell_contents(&mut reader, &mut buf, py, &CellTypeAttr::Number, &[])
                        .is_err()
                );
            }
        });
    }

    #[test]
    fn streaming_formula_and_cache_aggregate_references() {
        pyo3::Python::initialize();
        Python::attach(|py| {
            let xml = b"<c><f> A1 &amp; &quot; B &quot;&#10;</f><v>1&#50;3</v></c>";
            let mut reader =
                SheetXmlReader::Memory(XmlReader::from_reader(Cursor::new(xml.to_vec())));
            let mut buf = Vec::new();
            assert!(matches!(
                reader.read_event_into(&mut buf).unwrap(),
                Event::Start(_)
            ));
            let (value, kind) =
                read_cell_contents(&mut reader, &mut buf, py, &CellTypeAttr::Number, &[]).unwrap();
            assert_eq!(kind, "formula");
            let fields = value
                .extract::<std::collections::HashMap<String, String>>(py)
                .unwrap();
            assert_eq!(fields["formula"], "= A1 & \" B \"\n");
            assert_eq!(fields["cached"], "123");
        });
    }
}
