//! SAX-based streaming sheet reader (Sprint Ι Pod-β).
//!
//! Activated by `load_workbook(path, read_only=True)` (or auto-trigger when
//! a sheet has > 50k rows). Walks `xl/worksheets/sheetN.xml` one row at a
//! time using an incremental XML reader over a bounded in-memory sheet part, or
//! for larger worksheets over the sheet part inflated from its ZIP entry while
//! parsing, resolving shared-string-table (SST) references against a table
//! loaded once per workbook. Each SST entry becomes one Python `str` on first
//! reference, reused by every later reference. Style metadata is
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
//! - `reader.close()` — eagerly closes the XML reader and releases the sheet part.
//!
//! Memory profile: SST loaded once (typically <10MB even on huge
//! workbooks); sheet XML uses a bounded in-memory fast path for small/medium
//! sheets and inflates large sheets from the ZIP entry as rows are read.

use std::fs::File;
use std::io::{self, BufReader, Cursor, Read, Seek, SeekFrom};
use std::sync::Arc;

use flate2::read::DeflateDecoder;
use pyo3::exceptions::{PyIOError, PyMemoryError, PyStopIteration, PyValueError};
use pyo3::ffi;
use pyo3::prelude::*;
use pyo3::sync::PyOnceLock;
use pyo3::types::{PyBytes, PyDateTime, PyDict, PyFloat, PyInt, PyList, PyString, PyTime, PyTuple};
use pyo3::IntoPyObjectExt;

use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader as XmlReader;
use zip::{CompressionMethod, ZipArchive};

use crate::iso8601::{is_iso8601, iso8601_to_py};
use crate::ooxml_util;
use crate::streaming_scan::{HandoffReader, RowScanner};

type PyObjectOwned = Py<PyAny>;
// Keep read_only value scans below openpyxl's peak RSS on ordinary table-sized
// workbooks; larger sheet XML parts inflate from the ZIP entry while parsing.
const IN_MEMORY_SHEET_XML_LIMIT: u64 = 1 * 1024 * 1024;
// Read buffer between the entry inflater and the XML reader.
const SHEET_ENTRY_BUFFER: usize = 64 * 1024;
// Bound the value cells materialized per chunk so wide sheets (up to 16,384
// columns) stream a few rows at a time instead of 1,024 padded rows.
const CHUNK_CELL_BUDGET: usize = 65_536;

// pyo3's `IntoPyObject` for str, ints and floats, `PyList::empty`,
// `PyDict::new`, and `PyTuple::new` panic when CPython returns NULL, which
// drops the pending `MemoryError` and raises `PanicException` (a
// `BaseException`) instead. The constructors below return the pending error
// so an exhausted allocator surfaces as a catchable `MemoryError`.

/// Take ownership of a new reference, or fetch the pending Python error.
fn owned_or_err(py: Python<'_>, ptr: *mut ffi::PyObject) -> PyResult<Bound<'_, PyAny>> {
    // SAFETY: every caller passes the result of a CPython constructor, which is
    // either NULL or a new strong reference.
    unsafe { Bound::from_owned_ptr_or_err(py, ptr) }
}

fn str_to_py(py: Python<'_>, value: &str) -> PyResult<PyObjectOwned> {
    let len = ffi::Py_ssize_t::try_from(value.len())
        .map_err(|_| PyMemoryError::new_err("string too large for a Python str"))?;
    // SAFETY: `value` is valid UTF-8 for `len` bytes.
    let ptr = unsafe { ffi::PyUnicode_FromStringAndSize(value.as_ptr().cast(), len) };
    Ok(owned_or_err(py, ptr)?.unbind())
}

fn u32_to_py(py: Python<'_>, value: u32) -> PyResult<PyObjectOwned> {
    // SAFETY: plain value constructor.
    let ptr = unsafe { ffi::PyLong_FromUnsignedLong(std::os::raw::c_ulong::from(value)) };
    Ok(owned_or_err(py, ptr)?.unbind())
}

fn i64_to_py(py: Python<'_>, value: i64) -> PyResult<PyObjectOwned> {
    // SAFETY: plain value constructor.
    let ptr = unsafe { ffi::PyLong_FromLongLong(value) };
    Ok(owned_or_err(py, ptr)?.unbind())
}

fn f64_to_py(py: Python<'_>, value: f64) -> PyResult<PyObjectOwned> {
    // SAFETY: plain value constructor.
    let ptr = unsafe { ffi::PyFloat_FromDouble(value) };
    Ok(owned_or_err(py, ptr)?.unbind())
}

fn new_list(py: Python<'_>) -> PyResult<Bound<'_, PyList>> {
    // SAFETY: plain constructor.
    let ptr = unsafe { ffi::PyList_New(0) };
    Ok(owned_or_err(py, ptr)?.cast_into::<PyList>()?)
}

fn new_dict(py: Python<'_>) -> PyResult<Bound<'_, PyDict>> {
    // SAFETY: plain constructor.
    let ptr = unsafe { ffi::PyDict_New() };
    Ok(owned_or_err(py, ptr)?.cast_into::<PyDict>()?)
}

fn set_dict_item(dict: &Bound<'_, PyDict>, key: &str, value: PyObjectOwned) -> PyResult<()> {
    dict.set_item(str_to_py(dict.py(), key)?, value)
}

fn new_tuple<'py>(
    py: Python<'py>,
    items: impl ExactSizeIterator<Item = PyObjectOwned>,
) -> PyResult<Bound<'py, PyTuple>> {
    let len = ffi::Py_ssize_t::try_from(items.len())
        .map_err(|_| PyMemoryError::new_err("tuple too large for a Python tuple"))?;
    // SAFETY: plain constructor.
    let ptr = unsafe { ffi::PyTuple_New(len) };
    let tuple = owned_or_err(py, ptr)?.cast_into::<PyTuple>()?;
    for (index, item) in (0..len).zip(items) {
        // SAFETY: `index` is in range and `PyTuple_SetItem` steals `item`'s
        // reference, including on failure.
        if unsafe { ffi::PyTuple_SetItem(tuple.as_ptr(), index, item.into_ptr()) } != 0 {
            return Err(PyErr::fetch(py));
        }
    }
    Ok(tuple)
}

/// Parse `xl/sharedStrings.xml` into a flat `Vec<String>`. Each entry is
/// the plain-text concatenation of any nested `<r><t>...</t></r>` runs
/// (matches Excel/openpyxl's flattening for `Cell.value`). Phonetic runs
/// (`<rPh>`) annotate the base text and are excluded.
fn load_sst<R: Read + Seek>(zip: &mut ZipArchive<R>) -> PyResult<Vec<String>> {
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
    let mut in_rph = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.local_name();
                match name.as_ref().as_bytes() {
                    b"si" => {
                        in_si = true;
                        in_rph = false;
                        current.clear();
                    }
                    b"rPh" if in_si => in_rph = true,
                    b"t" => {
                        if in_si && !in_rph {
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
                    b"rPh" => in_rph = false,
                    _ => {}
                }
            }
            Ok(Event::Text(t)) => {
                if in_si && in_t {
                    current.push_str(&normalize_xml_eol(&t));
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

/// XML end-of-line handling (XML 1.0 section 2.11): a literal CR LF pair or
/// lone CR in character data reads as LF. Character references such as
/// `&#13;` arrive as separate reference events and keep their CR.
fn normalize_xml_eol(text: &str) -> std::borrow::Cow<'_, str> {
    if text.contains('\r') {
        std::borrow::Cow::Owned(text.replace("\r\n", "\n").replace('\r', "\n"))
    } else {
        std::borrow::Cow::Borrowed(text)
    }
}

/// Resolve `<sheet name=...>` → ZIP path (`xl/worksheets/sheetN.xml`).
fn resolve_sheet_xml_path<R: Read + Seek>(
    zip: &mut ZipArchive<R>,
    sheet: &str,
) -> PyResult<String> {
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

/// Shared-string entries of one workbook. An entry becomes a Python `str` on
/// its first reference and every later reference, from any reader sharing the
/// table, returns that same object, so a string repeated across many cells is
/// materialized once rather than once per cell. The native xlsx backend builds
/// its table from the native reader's entries, so its eager reads and the
/// streaming readers it hands the table to share one `str` per entry.
pub(crate) struct SharedStringTable {
    strings: Vec<Arc<str>>,
    py_strings: Box<[PyOnceLock<PyObjectOwned>]>,
}

impl SharedStringTable {
    fn new(strings: Vec<String>) -> Self {
        Self::from_shared(strings.into_iter().map(Arc::from).collect())
    }

    /// Table over entries already held as `Arc<str>`, without copying text.
    pub(crate) fn from_shared(strings: Vec<Arc<str>>) -> Self {
        let py_strings = strings.iter().map(|_| PyOnceLock::new()).collect();
        Self {
            strings,
            py_strings,
        }
    }

    fn len(&self) -> usize {
        self.strings.len()
    }

    /// Python `str` of entry `idx`; an index past the table reads as `""`.
    pub(crate) fn get_py(&self, py: Python<'_>, idx: usize) -> PyResult<PyObjectOwned> {
        let (Some(text), Some(slot)) = (self.strings.get(idx), self.py_strings.get(idx)) else {
            return str_to_py(py, "");
        };
        if let Some(cached) = slot.get(py) {
            return Ok(cached.clone_ref(py));
        }
        let created = str_to_py(py, text)?;
        match slot.set(py, created.clone_ref(py)) {
            Ok(()) => Ok(created),
            // Another thread stored this entry first; share its object.
            Err(_) => Ok(slot.get(py).map_or(created, |stored| stored.clone_ref(py))),
        }
    }
}

/// Parsed `xl/sharedStrings.xml`, loaded once per workbook and shared by
/// every [`StreamingSheetReader`] opened from the same source, together with
/// the Python `str` of each entry already referenced.
#[pyclass(frozen, module = "wolfxl._rust")]
pub struct StreamingSharedStrings {
    sst: Arc<SharedStringTable>,
}

#[pymethods]
impl StreamingSharedStrings {
    /// Parse the shared-strings table of the xlsx at `path`.
    #[staticmethod]
    pub fn load(path: &str) -> PyResult<Self> {
        let file = File::open(path)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to open xlsx: {e}")))?;
        let zip = ZipArchive::new(file)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to open zip: {e}")))?;
        Self::from_archive(zip)
    }

    /// [`Self::load`] for in-memory xlsx bytes (file-like sources).
    #[staticmethod]
    pub fn load_from_bytes(data: &Bound<'_, PyBytes>) -> PyResult<Self> {
        let zip = ZipArchive::new(PyBytesCursor::new(data))
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to open zip: {e}")))?;
        Self::from_archive(zip)
    }

    fn __len__(&self) -> usize {
        self.sst.len()
    }
}

impl StreamingSharedStrings {
    fn from_archive<R: Read + Seek>(mut zip: ZipArchive<R>) -> PyResult<Self> {
        ooxml_util::validate_zip_archive(&mut zip)?;
        Ok(Self {
            sst: Arc::new(SharedStringTable::new(load_sst(&mut zip)?)),
        })
    }

    /// Handle over a table another backend already holds.
    pub(crate) fn from_table(sst: Arc<SharedStringTable>) -> Self {
        Self { sst }
    }
}

/// Streaming sheet reader. See module docs.
#[pyclass(unsendable, module = "wolfxl._rust")]
pub struct StreamingSheetReader {
    /// Incremental XML reader over the worksheet XML.
    reader: Option<SheetXmlReader>,
    /// Scratch buffer reused for quick-xml events.
    event_buf: Vec<u8>,
    /// Shared-strings table: supplied by the caller or loaded from
    /// `xl/sharedStrings.xml` when the reader opens.
    sst: Arc<SharedStringTable>,
    /// Whether the reader has been exhausted.
    exhausted: bool,
    /// Optional `min_row` bound (1-based, inclusive). Rows below are skipped.
    min_row: Option<u32>,
    /// Optional `max_row` bound (1-based, inclusive). Iteration stops past this.
    max_row: Option<u32>,
    /// Whether iteration stopped at a `<row>` beyond `max_row`. openpyxl's
    /// read-only iterator pads trailing blank rows only in that case.
    stopped_past_max_row: bool,
    /// Optional `min_col` bound (1-based, inclusive).
    min_col: Option<u32>,
    /// Optional `max_col` bound (1-based, inclusive).
    max_col: Option<u32>,
    /// Index of the previous `<row>`; a `<row>` without `r` follows it.
    last_row: u32,
    /// Date conversion configured once per values-only streaming pass.
    date_styles: DateStyles,
    /// What values-only rows return for formula cells.
    value_output: CellOutput,
}

/// How `read_cell_contents` reports formula and error cells.
#[derive(Clone, Copy, PartialEq, Eq)]
enum CellOutput {
    /// `{"type": "formula"|"error", ...}` payloads for `StreamingCell`.
    Payload,
    /// Values-only rows with `data_only=False`: the `=`-prefixed formula text.
    FormulaText,
    /// Values-only rows with `data_only=True`: the cached result.
    CachedValue,
}

#[derive(Default)]
struct DateStyles {
    kinds: Vec<u8>,
    mac_epoch: bool,
    timedelta: Option<PyObjectOwned>,
}

impl DateStyles {
    fn kind(&self, style_id: Option<u32>) -> u8 {
        style_id
            .and_then(|id| self.kinds.get(id as usize))
            .copied()
            .unwrap_or(0)
    }

    fn convert(&self, py: Python<'_>, kind: u8, value: PyObjectOwned) -> PyResult<PyObjectOwned> {
        let bound = value.bind(py);
        if !(bound.is_exact_instance_of::<PyFloat>() || bound.is_exact_instance_of::<PyInt>()) {
            return Ok(value);
        }
        if kind == 2 {
            return match &self.timedelta {
                Some(convert) => convert.call1(py, (bound,)),
                None => Ok(value),
            };
        }
        let serial: f64 = bound.extract()?;
        Ok(excel_serial_to_py(py, serial, self.mac_epoch)?.unwrap_or(value))
    }

    fn apply(
        &self,
        py: Python<'_>,
        style_id: Option<u32>,
        value: PyObjectOwned,
        cell_type: &str,
        numeric: bool,
    ) -> PyResult<PyObjectOwned> {
        let kind = self.kind(style_id);
        if kind == 0 || !numeric || cell_type != "n" {
            return Ok(value);
        }
        self.convert(py, kind, value)
    }
}

/// Match openpyxl's from_excel: millisecond ties-to-even rounding, fractional
/// serials as time, and the Windows leap-day adjustment below serial 60.
/// Keep serials outside Python's year range numeric, as the Python fallback does.
fn excel_serial_to_py(
    py: Python<'_>,
    value: f64,
    mac_epoch: bool,
) -> PyResult<Option<PyObjectOwned>> {
    if !value.is_finite() {
        return Ok(None);
    }
    // Python's divmod uses a nonnegative fraction, including for negative days.
    let fraction = value.rem_euclid(1.0);
    let millis = (fraction * 86_400.0 * 1_000.0).round_ties_even() as i64;
    if (0.0..1.0).contains(&value) && millis < 86_400_000 {
        let secs = (millis / 1_000) as u32;
        return Ok(Some(
            PyTime::new(
                py,
                (secs / 3_600) as u8,
                (secs / 60 % 60) as u8,
                (secs % 60) as u8,
                (millis % 1_000) as u32 * 1_000,
                None,
            )?
            .into_any()
            .unbind(),
        ));
    }
    let mut day = value.floor();
    if !mac_epoch && value > 0.0 && value < 60.0 {
        day += 1.0;
    }
    if day.abs() > 1e9 {
        return Ok(None);
    }
    let epoch = if mac_epoch {
        chrono::NaiveDate::from_ymd_opt(1904, 1, 1)
    } else {
        chrono::NaiveDate::from_ymd_opt(1899, 12, 30)
    }
    .expect("valid Excel epoch")
    .and_hms_opt(0, 0, 0)
    .expect("valid midnight");
    let Some(datetime) = chrono::Duration::try_days(day as i64)
        .and_then(|days| epoch.checked_add_signed(days))
        .and_then(|date| date.checked_add_signed(chrono::Duration::milliseconds(millis)))
    else {
        return Ok(None);
    };
    use chrono::{Datelike, Timelike};
    if !(1..=9999).contains(&datetime.year()) {
        return Ok(None);
    }
    Ok(Some(
        PyDateTime::new(
            py,
            datetime.year(),
            datetime.month() as u8,
            datetime.day() as u8,
            datetime.hour() as u8,
            datetime.minute() as u8,
            datetime.second() as u8,
            datetime.nanosecond() / 1_000,
            None,
        )?
        .into_any()
        .unbind(),
    ))
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

/// Worksheet XML source. Values-only rows come from the byte-level row
/// scanner; the first event read (other modes, or a row the fast path leaves
/// to quick-xml) hands the rest of the stream to quick-xml for good.
enum SheetXmlReader {
    Scan(RowScanner),
    Events(XmlReader<BufReader<HandoffReader>>),
}

impl SheetXmlReader {
    fn new(initial: Vec<u8>, source: Option<Box<dyn Read>>) -> Self {
        Self::Scan(RowScanner::new(initial, source))
    }

    fn events(&mut self) -> &mut XmlReader<BufReader<HandoffReader>> {
        if let Self::Scan(scanner) = self {
            let scanner = std::mem::replace(scanner, RowScanner::new(Vec::new(), None));
            let mut reader = XmlReader::from_reader(BufReader::with_capacity(
                SHEET_ENTRY_BUFFER,
                scanner.handoff(),
            ));
            reader.config_mut().trim_text(false);
            *self = Self::Events(reader);
        }
        match self {
            Self::Events(reader) => reader,
            Self::Scan(_) => unreachable!("the scanner was just replaced"),
        }
    }

    fn read_event_into<'a>(&mut self, buf: &'a mut Vec<u8>) -> quick_xml::Result<Event<'a>> {
        self.events().read_event_into(buf)
    }

    fn read_text_into<'a>(
        &mut self,
        end: quick_xml::name::QName<'_>,
        buf: &'a mut Vec<u8>,
    ) -> quick_xml::Result<quick_xml::events::BytesText<'a>> {
        self.events().read_text_into(end, buf)
    }
}

/// Compressed bytes of one ZIP entry, owned so the sheet part can be inflated
/// while it is parsed after the archive itself has been released.
enum EntrySource {
    File(io::Take<File>),
    /// A range of the caller's immutable `bytes` object, shared rather than copied.
    Bytes {
        data: Py<PyBytes>,
        pos: usize,
        end: usize,
    },
}

impl Read for EntrySource {
    fn read(&mut self, buf: &mut [u8]) -> io::Result<usize> {
        match self {
            Self::File(source) => source.read(buf),
            Self::Bytes { data, pos, end } => Python::attach(|py| {
                let available = &data.as_bytes(py)[*pos..*end];
                let n = available.len().min(buf.len());
                buf[..n].copy_from_slice(&available[..n]);
                *pos += n;
                Ok(n)
            }),
        }
    }
}

/// Archive readers that can hand over one entry's compressed byte range.
trait IntoEntrySource: Read + Seek + Sized {
    fn into_entry_source(self, start: u64, len: u64) -> io::Result<EntrySource>;
}

impl IntoEntrySource for File {
    fn into_entry_source(mut self, start: u64, len: u64) -> io::Result<EntrySource> {
        self.seek(SeekFrom::Start(start))?;
        Ok(EntrySource::File(self.take(len)))
    }
}

/// Caller-owned xlsx bytes read through a cursor; the owner is kept so the
/// sheet entry can stream from the same buffer without a copy.
struct PyBytesCursor<'a, 'py> {
    owner: &'a Bound<'py, PyBytes>,
    cursor: Cursor<&'a [u8]>,
}

impl<'a, 'py> PyBytesCursor<'a, 'py> {
    fn new(owner: &'a Bound<'py, PyBytes>) -> Self {
        Self {
            owner,
            cursor: Cursor::new(owner.as_bytes()),
        }
    }
}

impl Read for PyBytesCursor<'_, '_> {
    fn read(&mut self, buf: &mut [u8]) -> io::Result<usize> {
        self.cursor.read(buf)
    }
}

impl Seek for PyBytesCursor<'_, '_> {
    fn seek(&mut self, pos: SeekFrom) -> io::Result<u64> {
        self.cursor.seek(pos)
    }
}

impl IntoEntrySource for PyBytesCursor<'_, '_> {
    fn into_entry_source(self, start: u64, len: u64) -> io::Result<EntrySource> {
        let total = self.cursor.get_ref().len();
        let range = usize::try_from(start)
            .ok()
            .zip(usize::try_from(len).ok())
            .and_then(|(start, len)| Some(start..start.checked_add(len)?))
            .filter(|range| range.end <= total)
            .ok_or_else(|| {
                io::Error::new(
                    io::ErrorKind::UnexpectedEof,
                    "sheet part extends past the end of the package",
                )
            })?;
        Ok(EntrySource::Bytes {
            data: self.owner.clone().unbind(),
            pos: range.start,
            end: range.end,
        })
    }
}

enum EntryDecoder {
    Stored(EntrySource),
    Deflated(DeflateDecoder<EntrySource>),
}

/// Inflates a sheet part from its ZIP entry as it is parsed, so large sheets
/// need neither a full in-memory copy nor a temp file. Like the zip crate's
/// entry reader, it fails once output exceeds the declared uncompressed size
/// and verifies the CRC-32 at the end of the entry.
struct SheetEntryReader {
    decoder: EntryDecoder,
    crc: crc32fast::Hasher,
    expected_crc: u32,
    declared_size: u64,
    produced: u64,
}

impl SheetEntryReader {
    fn new(
        source: EntrySource,
        method: CompressionMethod,
        declared_size: u64,
        expected_crc: u32,
    ) -> io::Result<Self> {
        let decoder = match method {
            CompressionMethod::Stored => EntryDecoder::Stored(source),
            CompressionMethod::Deflated => EntryDecoder::Deflated(DeflateDecoder::new(source)),
            other => {
                return Err(io::Error::new(
                    io::ErrorKind::Unsupported,
                    format!("unsupported sheet part compression method: {other}"),
                ))
            }
        };
        Ok(Self {
            decoder,
            crc: crc32fast::Hasher::new(),
            expected_crc,
            declared_size,
            produced: 0,
        })
    }
}

impl Read for SheetEntryReader {
    fn read(&mut self, buf: &mut [u8]) -> io::Result<usize> {
        if buf.is_empty() {
            return Ok(0);
        }
        // Inflate at most one byte past the declared size, so an oversized
        // part is detected without producing more of it.
        let room = self
            .declared_size
            .saturating_sub(self.produced)
            .saturating_add(1);
        let len = buf.len().min(usize::try_from(room).unwrap_or(usize::MAX));
        let buf = &mut buf[..len];
        let n = match &mut self.decoder {
            EntryDecoder::Stored(source) => source.read(buf)?,
            EntryDecoder::Deflated(source) => source.read(buf)?,
        };
        if n == 0 {
            if self.crc.clone().finalize() != self.expected_crc {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidData,
                    "Invalid checksum",
                ));
            }
            return Ok(0);
        }
        self.produced += n as u64;
        if self.produced > self.declared_size {
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                "File is larger than its declared uncompressed size",
            ));
        }
        self.crc.update(&buf[..n]);
        Ok(n)
    }
}

#[pymethods]
impl StreamingSheetReader {
    /// Open `path` and prepare to stream `sheet`. Pass `shared_strings`
    /// loaded from the same source to skip re-parsing the table.
    #[staticmethod]
    #[pyo3(signature = (path, sheet, min_row=None, max_row=None, min_col=None, max_col=None, shared_strings=None))]
    #[allow(clippy::too_many_arguments)]
    pub fn open(
        path: &str,
        sheet: &str,
        min_row: Option<u32>,
        max_row: Option<u32>,
        min_col: Option<u32>,
        max_col: Option<u32>,
        shared_strings: Option<&StreamingSharedStrings>,
    ) -> PyResult<Self> {
        let file = File::open(path)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to open xlsx: {e}")))?;
        let zip = ZipArchive::new(file)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to open zip: {e}")))?;
        Self::from_archive(
            zip,
            sheet,
            min_row,
            max_row,
            min_col,
            max_col,
            shared_strings,
        )
    }

    /// Prepare to stream `sheet` from in-memory xlsx bytes (file-like sources).
    #[staticmethod]
    #[pyo3(signature = (data, sheet, min_row=None, max_row=None, min_col=None, max_col=None, shared_strings=None))]
    #[allow(clippy::too_many_arguments)]
    pub fn open_from_bytes(
        data: &Bound<'_, PyBytes>,
        sheet: &str,
        min_row: Option<u32>,
        max_row: Option<u32>,
        min_col: Option<u32>,
        max_col: Option<u32>,
        shared_strings: Option<&StreamingSharedStrings>,
    ) -> PyResult<Self> {
        let zip = ZipArchive::new(PyBytesCursor::new(data))
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to open zip: {e}")))?;
        Self::from_archive(
            zip,
            sheet,
            min_row,
            max_row,
            min_col,
            max_col,
            shared_strings,
        )
    }

    /// `(max_row, max_col)` covered by `sheet`'s cells and merged ranges in the
    /// xlsx at `path`: the extent an eager openpyxl load reports, which ignores
    /// the declared dimension and row-only formatting records.
    #[staticmethod]
    pub fn cell_bounds(path: &str, sheet: &str) -> PyResult<(u32, u32)> {
        let file = File::open(path)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to open xlsx: {e}")))?;
        let zip = ZipArchive::new(file)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to open zip: {e}")))?;
        archive_cell_bounds(zip, sheet)
    }

    /// [`Self::cell_bounds`] for in-memory xlsx bytes (file-like sources).
    #[staticmethod]
    pub fn cell_bounds_from_bytes(data: &Bound<'_, PyBytes>, sheet: &str) -> PyResult<(u32, u32)> {
        let zip = ZipArchive::new(PyBytesCursor::new(data))
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to open zip: {e}")))?;
        archive_cell_bounds(zip, sheet)
    }

    /// Whether `sheet` in the xlsx at `path` holds more than `limit` `<row>`
    /// start tags. Stops reading at the first tag past `limit`.
    #[staticmethod]
    pub fn row_tags_exceed(path: &str, sheet: &str, limit: usize) -> PyResult<bool> {
        let file = File::open(path)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to open xlsx: {e}")))?;
        let zip = ZipArchive::new(file)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to open zip: {e}")))?;
        archive_row_tags_exceed(zip, sheet, limit)
    }

    /// [`Self::row_tags_exceed`] for in-memory xlsx bytes (file-like sources).
    #[staticmethod]
    pub fn row_tags_exceed_from_bytes(
        data: &Bound<'_, PyBytes>,
        sheet: &str,
        limit: usize,
    ) -> PyResult<bool> {
        let zip = ZipArchive::new(PyBytesCursor::new(data))
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to open zip: {e}")))?;
        archive_row_tags_exceed(zip, sheet, limit)
    }

    /// Install source-style kinds (0 plain, 1 datetime/time, 2 duration).
    /// Durations use the retained converter to preserve Python timedelta rounding.
    pub fn set_date_styles(
        &mut self,
        kinds: &Bound<'_, PyBytes>,
        timedelta_converter: PyObjectOwned,
        mac_epoch: bool,
    ) {
        self.date_styles = DateStyles {
            kinds: kinds.as_bytes().to_vec(),
            mac_epoch,
            timedelta: Some(timedelta_converter),
        };
    }

    /// Choose what values-only rows return for formula cells: the cached
    /// result (`data_only=True`) or the formula text.
    pub fn set_data_only(&mut self, data_only: bool) {
        self.value_output = if data_only {
            CellOutput::CachedValue
        } else {
            CellOutput::FormulaText
        };
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
                    let cell_list = new_list(py)?;
                    for c in cells {
                        let style_obj = match c.style_id {
                            Some(s) => u32_to_py(py, s)?,
                            None => py.None(),
                        };
                        let tup = new_tuple(
                            py,
                            [
                                u32_to_py(py, c.col)?,
                                c.value,
                                style_obj,
                                str_to_py(py, c.cell_type)?,
                            ]
                            .into_iter(),
                        )?;
                        cell_list.append(tup)?;
                    }
                    let outer = new_tuple(
                        py,
                        [u32_to_py(py, row_idx)?, cell_list.into_any().unbind()].into_iter(),
                    )?;
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
                ValueStepResult::Row(_row_idx, values) => return Ok(Some(values)),
                ValueStepResult::Skip => continue,
                ValueStepResult::Done => {
                    self.exhausted = true;
                    return Ok(None);
                }
            }
        }
    }

    /// Read the next row as `(row_index, value_tuple)` without boxing
    /// style ids into Python, including configured date-style conversion.
    pub fn read_next_values_row<'py>(
        &mut self,
        py: Python<'py>,
    ) -> PyResult<Option<Bound<'py, PyTuple>>> {
        if self.exhausted {
            return Ok(None);
        }
        loop {
            match self.parse_one_value_row(py)? {
                ValueStepResult::Row(row_idx, values) => {
                    let row_obj = u32_to_py(py, row_idx)?;
                    return Ok(Some(new_tuple(
                        py,
                        [row_obj, values.into_any().unbind()].into_iter(),
                    )?));
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
        if self.exhausted {
            return Ok(None);
        }
        let limit = max_rows.max(1);
        let chunk = new_list(py)?;
        let mut cells = 0usize;
        while chunk.len() < limit && cells < CHUNK_CELL_BUDGET {
            match self.parse_one_value_row(py)? {
                ValueStepResult::Row(row_idx, values) => {
                    cells += values.len();
                    let row_obj = u32_to_py(py, row_idx)?;
                    let item = new_tuple(py, [row_obj, values.into_any().unbind()].into_iter())?;
                    chunk.append(item)?;
                }
                ValueStepResult::Skip => continue,
                ValueStepResult::Done => {
                    self.exhausted = true;
                    break;
                }
            }
        }
        Ok((!chunk.is_empty()).then_some(chunk))
    }

    /// Read a value-row chunk as `(items, first_row, last_row, dense)`.
    ///
    /// Dense chunks (contiguous row indexes) return only value tuples.
    /// Sparse chunks return `(row_index, value_tuple)` records.
    #[pyo3(signature = (max_rows=1024))]
    pub fn read_next_values_compact_chunk<'py>(
        &mut self,
        py: Python<'py>,
        max_rows: usize,
    ) -> PyResult<Option<Bound<'py, PyTuple>>> {
        if self.exhausted {
            return Ok(None);
        }
        let limit = max_rows.max(1);
        let mut rows: Vec<(u32, Bound<'py, PyTuple>)> = Vec::with_capacity(limit);
        let mut first_row = 0;
        let mut last_row = 0;
        let mut contiguous = true;

        let mut cells = 0usize;
        while rows.len() < limit && cells < CHUNK_CELL_BUDGET {
            match self.parse_one_value_row(py)? {
                ValueStepResult::Row(row_idx, values) => {
                    if rows.is_empty() {
                        first_row = row_idx;
                    } else if row_idx != last_row + 1 {
                        contiguous = false;
                    }
                    last_row = row_idx;
                    cells += values.len();
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

        let dense = contiguous;
        let items = new_list(py)?;
        if dense {
            for (_row_idx, values) in rows {
                items.append(values)?;
            }
        } else {
            for (row_idx, values) in rows {
                let row_obj = u32_to_py(py, row_idx)?;
                items.append(new_tuple(
                    py,
                    [row_obj, values.into_any().unbind()].into_iter(),
                )?)?;
            }
        }

        Ok(Some(new_tuple(
            py,
            [
                items.into_any().unbind(),
                u32_to_py(py, first_row)?,
                u32_to_py(py, last_row)?,
                dense.into_py_any(py)?,
            ]
            .into_iter(),
        )?))
    }

    /// True once the stream has been fully consumed.
    pub fn is_exhausted(&self) -> bool {
        self.exhausted
    }

    /// True once iteration stopped at a `<row>` beyond the `max_row` bound.
    pub fn stopped_past_max_row(&self) -> bool {
        self.stopped_past_max_row
    }

    /// Eagerly drop the in-memory XML buffer (releases peak RSS).
    pub fn close(&mut self) {
        self.exhausted = true;
        self.reader = None;
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
    Row(u32, Bound<'py, PyTuple>),
    Skip,
    Done,
}

impl StreamingSheetReader {
    /// Resolve shared strings and stage `sheet`'s XML: in memory when small,
    /// else inflated from its ZIP entry while rows are read.
    fn from_archive<R: IntoEntrySource>(
        mut zip: ZipArchive<R>,
        sheet: &str,
        min_row: Option<u32>,
        max_row: Option<u32>,
        min_col: Option<u32>,
        max_col: Option<u32>,
        shared_strings: Option<&StreamingSharedStrings>,
    ) -> PyResult<Self> {
        ooxml_util::validate_zip_archive(&mut zip)?;

        let sst = match shared_strings {
            Some(shared) => Arc::clone(&shared.sst),
            None => Arc::new(SharedStringTable::new(load_sst(&mut zip)?)),
        };
        let sheet_path = resolve_sheet_xml_path(&mut zip, sheet)?;
        let mut sheet_entry = zip.by_name(&sheet_path).map_err(|e| {
            PyErr::new::<PyIOError, _>(format!("Failed to open sheet part '{sheet_path}': {e}"))
        })?;
        let reader = if sheet_entry.size() <= IN_MEMORY_SHEET_XML_LIMIT {
            let mut xml = Vec::with_capacity(sheet_entry.size() as usize);
            sheet_entry
                .read_to_end(&mut xml)
                .map_err(|e| PyErr::new::<PyIOError, _>(format!("read sheet XML: {e}")))?;
            SheetXmlReader::new(xml, None)
        } else {
            let data_start = sheet_entry.data_start().ok_or_else(|| {
                PyErr::new::<PyIOError, _>(format!("sheet part '{sheet_path}' has no data offset"))
            })?;
            let compressed_size = sheet_entry.compressed_size();
            let method = sheet_entry.compression();
            let size = sheet_entry.size();
            let crc32 = sheet_entry.crc32();
            drop(sheet_entry);
            let entry = zip
                .into_inner()
                .into_entry_source(data_start, compressed_size)
                .and_then(|source| SheetEntryReader::new(source, method, size, crc32))
                .map_err(|e| PyErr::new::<PyIOError, _>(format!("open sheet XML: {e}")))?;
            SheetXmlReader::new(Vec::new(), Some(Box::new(entry)))
        };

        Ok(Self {
            reader: Some(reader),
            event_buf: Vec::with_capacity(8192),
            sst,
            exhausted: false,
            min_row,
            max_row,
            stopped_past_max_row: false,
            min_col,
            max_col,
            last_row: 0,
            date_styles: DateStyles::default(),
            value_output: CellOutput::FormulaText,
        })
    }

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
                    let row_idx = row_index_from_start(&e, self.last_row)?;
                    self.last_row = row_idx;
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
                    let row_idx = row_index_from_start(&e, self.last_row)?;
                    self.last_row = row_idx;
                    return self.row_result(row_idx, Vec::new());
                }
                Event::Eof => return Ok(StepResult::Done),
                _ => {}
            }
        }
    }

    fn parse_one_value_row<'py>(&mut self, py: Python<'py>) -> PyResult<ValueStepResult<'py>> {
        if let Some(step) = self.scan_value_row(py)? {
            return Ok(step);
        }
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
                    let row_idx = row_index_from_start(&e, self.last_row)?;
                    self.last_row = row_idx;
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
                            self.stopped_past_max_row = true;
                            return Ok(ValueStepResult::Done);
                        }
                    }
                    let values = read_values_until_row_end(
                        reader,
                        &mut self.event_buf,
                        py,
                        &self.sst,
                        &self.date_styles,
                        self.value_output,
                        row_idx,
                        self.min_col,
                        self.max_col,
                    )?;
                    return Ok(ValueStepResult::Row(row_idx, values));
                }
                Event::Empty(e) if e.local_name().as_ref() == "row" => {
                    let row_idx = row_index_from_start(&e, self.last_row)?;
                    self.last_row = row_idx;
                    return self.empty_value_row_result(py, row_idx);
                }
                Event::Eof => return Ok(ValueStepResult::Done),
                _ => {}
            }
        }
    }

    /// Values-only fast path: `Some(step)` when the row scanner handled the
    /// next row, `None` when the event reader must take over from here. A row
    /// the scanner cannot parse is left unconsumed, so quick-xml re-reads it.
    fn scan_value_row<'py>(&mut self, py: Python<'py>) -> PyResult<Option<ValueStepResult<'py>>> {
        let Some(SheetXmlReader::Scan(scanner)) = self.reader.as_mut() else {
            return Ok(None);
        };
        let Some(window) = scanner.next_row() else {
            return Ok(None);
        };
        let text = scanner.text();
        let previous = self.last_row;
        let tag = BytesStart::from_content(&text[window.tag.clone()], "row".len());
        let row_idx = row_index_from_start(&tag, previous)?;
        self.last_row = row_idx;
        if self.min_row.is_some_and(|min| row_idx < min) {
            // The window ends at the first `</row>` bytes, which a comment,
            // CDATA section, or attribute value may contain; such rows, and
            // text quick-xml would reject, are skipped by the event reader.
            if !window.self_closing && !is_plain_row_body(&text.as_bytes()[window.body.clone()]) {
                self.last_row = previous;
                return Ok(None);
            }
            scanner.consume(&window);
            return Ok(Some(ValueStepResult::Skip));
        }
        if self.max_row.is_some_and(|max| row_idx > max) {
            self.exhausted = true;
            self.stopped_past_max_row = true;
            return Ok(Some(ValueStepResult::Done));
        }
        let values = if window.self_closing {
            empty_value_row(py, self.min_col, self.max_col)?
        } else {
            match scan_value_row_body(
                &text[window.body.clone()],
                py,
                &self.sst,
                &self.date_styles,
                self.value_output,
                self.min_col,
                self.max_col,
            )? {
                Some(values) => values,
                None => {
                    self.last_row = previous;
                    return Ok(None);
                }
            }
        };
        scanner.consume(&window);
        Ok(Some(ValueStepResult::Row(row_idx, values)))
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
                self.stopped_past_max_row = true;
                return Ok(ValueStepResult::Done);
            }
        }
        Ok(ValueStepResult::Row(
            row_idx,
            empty_value_row(py, self.min_col, self.max_col)?,
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
                self.stopped_past_max_row = true;
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

/// Whether a scanned row body is plain markup: complete tags (no comments,
/// CDATA, or processing instructions) and text whose every `&` ends with `;`
/// as quick-xml requires, so its first `</row>` is the row's real end tag.
fn is_plain_row_body(body: &[u8]) -> bool {
    let mut at = 0;
    while let Some(offset) = memchr::memchr2(b'<', b'&', &body[at..]) {
        let i = at + offset;
        if body[i] == b'&' {
            match memchr::memchr3(b';', b'&', b'<', &body[i + 1..]) {
                Some(end) if body[i + 1 + end] == b';' => at = i + 2 + end,
                _ => return false,
            }
            continue;
        }
        if matches!(body.get(i + 1), Some(b'!' | b'?') | None) {
            return false;
        }
        let Some(gt) = scan_tag_end(body, i + 1) else {
            return false;
        };
        at = gt + 1;
    }
    true
}

fn read_cells_until_row_end(
    reader: &mut SheetXmlReader,
    buf: &mut Vec<u8>,
    py: Python<'_>,
    sst: &SharedStringTable,
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
                let (value, cell_type) =
                    read_cell_contents(reader, buf, py, &t_attr, sst, CellOutput::Payload)?;
                cells.push(ParsedCell {
                    col,
                    value,
                    style_id,
                    cell_type,
                });
            }
            Event::Empty(e) if e.local_name().as_ref() == "c" => {
                let (col, style_id, t_attr) = cell_attrs_from_start(&e, &cells)?;
                cells.push(ParsedCell {
                    col,
                    value: py.None(),
                    style_id,
                    cell_type: valueless_cell_type(&t_attr),
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
    sst: &SharedStringTable,
    date_styles: &DateStyles,
    output: CellOutput,
    row_idx: u32,
    min_col: Option<u32>,
    max_col: Option<u32>,
) -> PyResult<Bound<'py, PyTuple>> {
    let cmin = min_col.unwrap_or(1);
    let mut out: Vec<PyObjectOwned> = Vec::new();
    let mut last_col = cmin.saturating_sub(1);
    let mut next_col = 1;

    loop {
        buf.clear();
        let event = reader
            .read_event_into(buf)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Streaming reader XML: {e}")))?;
        match event {
            Event::Start(e) if e.local_name().as_ref() == "c" => {
                let (col, style_id, t_attr) = cell_value_attrs_from_start(&e, next_col)?;
                next_col = col.saturating_add(1);
                drop(e);
                let (value, cell_type) = read_cell_contents(reader, buf, py, &t_attr, sst, output)?;
                let value = date_styles.apply(
                    py,
                    style_id,
                    value,
                    cell_type,
                    matches!(t_attr, CellTypeAttr::Number),
                )?;
                push_bounded_value(py, &mut out, &mut last_col, col, value, min_col, max_col);
            }
            Event::Empty(e) if e.local_name().as_ref() == "c" => {
                let (col, _style_id, _t_attr) = cell_value_attrs_from_start(&e, next_col)?;
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
                return finish_value_row(py, out, last_col, min_col, max_col);
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
) {
    let cmin = min_col.unwrap_or(1);
    if col < cmin || max_col.is_some_and(|cmax| col > cmax) || col <= *last_col {
        return;
    }
    for _ in last_col.saturating_add(1)..col {
        out.push(py.None());
    }
    out.push(value);
    *last_col = col;
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
    new_tuple(py, out.into_iter())
}

fn empty_value_row<'py>(
    py: Python<'py>,
    min_col: Option<u32>,
    max_col: Option<u32>,
) -> PyResult<Bound<'py, PyTuple>> {
    let cmin = min_col.unwrap_or(1);
    finish_value_row(py, Vec::new(), cmin.saturating_sub(1), min_col, max_col)
}

/// Values-only fast path over a scanned `<row>` body, producing what
/// [`read_values_until_row_end`] would. `Ok(None)` when the body holds markup
/// left to quick-xml: anything but unprefixed `<c>` cells with `<f>`/`<v>`
/// children, or entity references outside `<v>`.
fn scan_value_row_body<'py>(
    body: &str,
    py: Python<'py>,
    sst: &SharedStringTable,
    date_styles: &DateStyles,
    output: CellOutput,
    min_col: Option<u32>,
    max_col: Option<u32>,
) -> PyResult<Option<Bound<'py, PyTuple>>> {
    let bytes = body.as_bytes();
    let mut out: Vec<PyObjectOwned> = Vec::new();
    let mut last_col = min_col.unwrap_or(1).saturating_sub(1);
    let mut next_col = 1;
    let mut at = 0;
    while let Some(offset) = memchr::memchr(b'<', &bytes[at..]) {
        let lt = at + offset;
        if bytes[at..lt].contains(&b'&') {
            return Ok(None);
        }
        let Some((tag, gt, self_closing)) = scan_start_tag(body, lt, "c") else {
            return Ok(None);
        };
        let (col, style_id, t_attr) =
            cell_value_attrs_from_start(&BytesStart::from_content(tag, "c".len()), next_col)?;
        next_col = col.saturating_add(1);
        if self_closing {
            push_bounded_value(
                py,
                &mut out,
                &mut last_col,
                col,
                py.None(),
                min_col,
                max_col,
            );
            at = gt + 1;
            continue;
        }
        let Some((value, cell_type, end)) =
            scan_cell_value(body, gt + 1, py, &t_attr, sst, output)?
        else {
            return Ok(None);
        };
        let value = date_styles.apply(
            py,
            style_id,
            value,
            cell_type,
            matches!(t_attr, CellTypeAttr::Number),
        )?;
        push_bounded_value(py, &mut out, &mut last_col, col, value, min_col, max_col);
        at = end;
    }
    if bytes[at..].contains(&b'&') {
        return Ok(None);
    }
    finish_value_row(py, out, last_col, min_col, max_col).map(Some)
}

/// [`read_cell_contents`] for a scanned cell whose content starts at `start`:
/// the value, its cell type, and the offset past `</c>`. `Ok(None)` when the
/// cell needs the event reader (rich or phonetic inline strings, unknown
/// children, unterminated references).
fn scan_cell_value(
    text: &str,
    start: usize,
    py: Python<'_>,
    t_attr: &CellTypeAttr,
    sst: &SharedStringTable,
    output: CellOutput,
) -> PyResult<Option<(PyObjectOwned, &'static str, usize)>> {
    let bytes = text.as_bytes();
    let keep_formula = output != CellOutput::CachedValue;
    let mut v_text: Option<String> = None;
    let mut f_text: Option<String> = None;
    let mut inline_text: Option<String> = None;
    let mut fast_number: Option<PyObjectOwned> = None;
    let mut fast_shared: Option<usize> = None;
    let mut has_formula = false;
    let mut at = start;
    loop {
        let Some(offset) = memchr::memchr(b'<', &bytes[at..]) else {
            return Ok(None);
        };
        let lt = at + offset;
        if !scan_cell_text(&text[at..lt], None)? {
            return Ok(None);
        }
        let rest = &text[lt..];
        if rest.starts_with("</c>") {
            // openpyxl reads an inline string only from `<is>` and every
            // other type only from `<v>`.
            let raw_value = if matches!(t_attr, CellTypeAttr::InlineString) {
                inline_text
            } else {
                v_text
            };
            let (value, cell_type) = value_cell_object(
                py,
                t_attr,
                output,
                has_formula,
                f_text,
                fast_number,
                fast_shared,
                raw_value,
                sst,
            )?;
            return Ok(Some((value, cell_type, lt + "</c>".len())));
        }
        if let Some((_, gt, self_closing)) = scan_start_tag(text, lt, "f") {
            has_formula = true;
            if self_closing {
                if keep_formula {
                    f_text.get_or_insert_with(String::new);
                }
                at = gt + 1;
                continue;
            }
            let Some((formula, end)) = scan_text_until(text, gt + 1, "</f>") else {
                return Ok(None);
            };
            let sink = keep_formula.then_some(&mut f_text);
            if !scan_cell_text(formula, sink)? {
                return Ok(None);
            }
            at = end;
            continue;
        }
        if let Some((_, gt, self_closing)) = scan_start_tag(text, lt, "is") {
            // openpyxl reads a present but empty `<is>` as "".
            inline_text.get_or_insert_with(String::new);
            if self_closing {
                at = gt + 1;
                continue;
            }
            let Some(end) = scan_inline_string(text, gt + 1, &mut inline_text)? else {
                return Ok(None);
            };
            at = end;
            continue;
        }
        if rest.starts_with("<v/>") {
            v_text.get_or_insert_with(String::new);
            at = lt + "<v/>".len();
            continue;
        }
        if !rest.starts_with("<v>") {
            return Ok(None);
        }
        let Some((raw, end)) = scan_text_until(text, lt + "<v>".len(), "</v>") else {
            return Ok(None);
        };
        at = end;
        // Formula-text rows ignore the cached result.
        if has_formula && output == CellOutput::FormulaText {
            continue;
        }
        let raw = normalize_xml_eol(raw);
        let value = quick_xml::escape::unescape(&raw)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("cell text decode: {e}")))?;
        if matches!(t_attr, CellTypeAttr::Number)
            && f_text.is_none()
            && v_text.is_none()
            && fast_number.is_none()
        {
            if let Some(number) = number_text_to_py(py, value.as_bytes())? {
                fast_number = Some(number);
                continue;
            }
        }
        if matches!(t_attr, CellTypeAttr::Shared) && v_text.is_none() && fast_shared.is_none() {
            if let Some(idx) = parse_usize_bytes(value.as_bytes()) {
                fast_shared = Some(idx);
                continue;
            }
        }
        if !value.is_empty() {
            v_text.get_or_insert_with(String::new).push_str(&value);
        }
    }
}

/// Plain `<t>` runs of an `<is>` body starting at `start`, appended to
/// `inline_text`; the offset past `</is>`. `Ok(None)` for rich runs or any
/// other markup, which the event reader handles.
fn scan_inline_string(
    text: &str,
    start: usize,
    inline_text: &mut Option<String>,
) -> PyResult<Option<usize>> {
    let bytes = text.as_bytes();
    let mut at = start;
    loop {
        let Some(offset) = memchr::memchr(b'<', &bytes[at..]) else {
            return Ok(None);
        };
        let lt = at + offset;
        if !scan_cell_text(&text[at..lt], None)? {
            return Ok(None);
        }
        if text[lt..].starts_with("</is>") {
            return Ok(Some(lt + "</is>".len()));
        }
        let Some((_, gt, self_closing)) = scan_start_tag(text, lt, "t") else {
            return Ok(None);
        };
        if self_closing {
            at = gt + 1;
            continue;
        }
        let Some((run, end)) = scan_text_until(text, gt + 1, "</t>") else {
            return Ok(None);
        };
        if !scan_cell_text(run, Some(inline_text))? {
            return Ok(None);
        }
        at = end;
    }
}

/// The `<name ...>` or `<name .../>` tag at `lt`, with no namespace prefix:
/// its content between `<` and `>` (without a trailing `/`), the offset of
/// `>`, and whether it is self-closing.
fn scan_start_tag<'a>(text: &'a str, lt: usize, name: &str) -> Option<(&'a str, usize, bool)> {
    let bytes = text.as_bytes();
    if !text[lt + 1..].starts_with(name) {
        return None;
    }
    let gt = scan_tag_end(bytes, lt + 1 + name.len())?;
    let self_closing = bytes[gt - 1] == b'/';
    let content = &text[lt + 1..gt - usize::from(self_closing)];
    // quick-xml's tag name runs to the first whitespace.
    let name_end = content
        .bytes()
        .position(|b| b.is_ascii_whitespace())
        .unwrap_or(content.len());
    (content[..name_end] == *name).then_some((content, gt, self_closing))
}

/// Offset of the first `>` at or after `from` outside a quoted value.
fn scan_tag_end(bytes: &[u8], from: usize) -> Option<usize> {
    let mut at = from;
    loop {
        let i = at + memchr::memchr3(b'>', b'"', b'\'', &bytes[at..])?;
        match bytes[i] {
            b'>' => return Some(i),
            quote => at = i + 1 + memchr::memchr(quote, &bytes[i + 1..])? + 1,
        }
    }
}

/// Text from `start` to the next markup, which must be `close`: the text and
/// the offset past `close`.
fn scan_text_until<'a>(text: &'a str, start: usize, close: &str) -> Option<(&'a str, usize)> {
    let lt = start + memchr::memchr(b'<', &text.as_bytes()[start..])?;
    text[lt..]
        .starts_with(close)
        .then(|| (&text[start..lt], lt + close.len()))
}

/// Apply one cell text node the way [`read_cell_contents`] handles its
/// `Text` and `GeneralRef` events: every reference is resolved (an unknown
/// one is an error) and, with a `sink`, line-normalized text and resolved
/// references are appended. `Ok(false)` for a `&` quick-xml would reject as
/// unterminated, so the event reader reports it.
fn scan_cell_text(raw: &str, mut sink: Option<&mut Option<String>>) -> PyResult<bool> {
    let mut push = |piece: &str| {
        if let Some(sink) = sink.as_mut() {
            if !piece.is_empty() {
                sink.get_or_insert_with(String::new).push_str(piece);
            }
        }
    };
    let mut rest = raw;
    while let Some(amp) = rest.find('&') {
        push(&normalize_xml_eol(&rest[..amp]));
        let after = &rest[amp + 1..];
        let Some(end) = after.find([';', '&']) else {
            return Ok(false);
        };
        if after.as_bytes()[end] != b';' {
            return Ok(false);
        }
        let resolved =
            ooxml_util::resolve_general_ref(&quick_xml::events::BytesRef::new(&after[..end]))
                .map_err(|e| PyErr::new::<PyIOError, _>(format!("cell text decode: {e}")))?;
        push(&resolved);
        rest = &after[end + 1..];
    }
    push(&normalize_xml_eol(rest));
    Ok(true)
}

fn read_cell_contents(
    reader: &mut SheetXmlReader,
    buf: &mut Vec<u8>,
    py: Python<'_>,
    t_attr: &CellTypeAttr,
    sst: &SharedStringTable,
    output: CellOutput,
) -> PyResult<(PyObjectOwned, &'static str)> {
    let mut v_text: Option<String> = None;
    let mut f_text: Option<String> = None;
    let mut inline_text: Option<String> = None;
    let mut fast_number: Option<PyObjectOwned> = None;
    let mut fast_shared: Option<usize> = None;
    let mut has_formula = false;
    // Cached-value rows never return formula text, so skip collecting it.
    let keep_formula = output != CellOutput::CachedValue;
    let mut in_f = false;
    let mut in_t = false;
    let mut in_rph = false;

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
                    // Formula-text rows ignore the cached result.
                    if has_formula && output == CellOutput::FormulaText {
                        continue;
                    }
                    let raw = normalize_xml_eol(&raw);
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
                b"f" => {
                    has_formula = true;
                    in_f = keep_formula;
                }
                b"t" => in_t = true,
                b"rPh" => in_rph = true,
                // openpyxl reads a present but empty `<is>` as "".
                b"is" => {
                    inline_text.get_or_insert_with(String::new);
                }
                _ => {}
            },
            Event::Empty(e) => match e.local_name().as_ref().as_bytes() {
                b"v" => {
                    v_text.get_or_insert_with(String::new);
                }
                b"f" => {
                    has_formula = true;
                    if keep_formula {
                        f_text.get_or_insert_with(String::new);
                    }
                }
                b"t" if !in_rph => {
                    inline_text.get_or_insert_with(String::new);
                }
                b"is" => {
                    inline_text.get_or_insert_with(String::new);
                }
                _ => {}
            },
            Event::Text(t) => {
                if in_f {
                    f_text
                        .get_or_insert_with(String::new)
                        .push_str(&normalize_xml_eol(&t));
                } else if in_t && !in_rph {
                    inline_text
                        .get_or_insert_with(String::new)
                        .push_str(&normalize_xml_eol(&t));
                }
            }
            Event::GeneralRef(r) => {
                let text = ooxml_util::resolve_general_ref(&r)
                    .map_err(|e| PyErr::new::<PyIOError, _>(format!("cell text decode: {e}")))?;
                if in_f {
                    f_text.get_or_insert_with(String::new).push_str(&text);
                } else if in_t && !in_rph {
                    inline_text.get_or_insert_with(String::new).push_str(&text);
                }
            }
            Event::End(e) => match e.local_name().as_ref().as_bytes() {
                b"c" => {
                    // openpyxl reads an inline string only from `<is>` and
                    // every other type only from `<v>`.
                    let raw_value = if matches!(t_attr, CellTypeAttr::InlineString) {
                        inline_text
                    } else {
                        v_text
                    };
                    if output != CellOutput::Payload {
                        return value_cell_object(
                            py,
                            t_attr,
                            output,
                            has_formula,
                            f_text,
                            fast_number,
                            fast_shared,
                            raw_value,
                            sst,
                        );
                    }
                    if f_text.is_none() {
                        if let Some(value) = fast_number {
                            return Ok((value, "n"));
                        }
                        if let Some(idx) = fast_shared {
                            return Ok((sst.get_py(py, idx)?, "s"));
                        }
                    } else if let Some(idx) = fast_shared {
                        let formula = f_text.unwrap_or_default();
                        let cached = sst.get_py(py, idx)?;
                        return Ok((build_formula_dict(py, &formula, cached)?, "formula"));
                    }
                    // `<v>` before `<f>` already parsed the numeric cache.
                    if let Some(cached) = fast_number {
                        let formula = f_text.unwrap_or_default();
                        return Ok((build_formula_dict(py, &formula, cached)?, "formula"));
                    }
                    return build_cell_object(py, t_attr, f_text, raw_value, sst);
                }
                b"f" => in_f = false,
                b"t" => in_t = false,
                b"rPh" => in_rph = false,
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
            b"t" => cell_type = cell_type_from_bytes(attr.value.as_bytes()),
            b"s" => style_id = parse_u32_bytes(attr.value.as_bytes()),
            _ => {}
        }
    }
    Ok((col.unwrap_or(default_col), style_id, cell_type))
}

/// `row@r`, or the row after `previous` when `r` is absent (openpyxl
/// `WorkSheetParser.parse_row`). Integral floats such as "3.0" are accepted.
fn row_index_from_start(e: &BytesStart<'_>, previous: u32) -> PyResult<u32> {
    for attr in e.attributes().with_checks(false).flatten() {
        let key = attr.key.as_ref().as_bytes();
        let local = key.rsplit(|b| *b == b':').next().unwrap_or(key);
        if local != b"r" {
            continue;
        }
        let bytes = attr.value.as_bytes();
        if let Some(row) = parse_u32_bytes(bytes) {
            return Ok(row);
        }
        let text = String::from_utf8_lossy(bytes);
        return match text.trim().parse::<f64>() {
            Ok(value) if value.fract() == 0.0 && value >= 1.0 && value <= f64::from(u32::MAX) => {
                Ok(value as u32)
            }
            _ => Err(PyErr::new::<PyIOError, _>(format!(
                "Streaming reader: bad row index {text:?}"
            ))),
        };
    }
    previous.checked_add(1).ok_or_else(|| {
        PyErr::new::<PyIOError, _>("Streaming reader: implicit row index overflow".to_string())
    })
}

fn archive_cell_bounds<R: Read + Seek>(
    mut zip: ZipArchive<R>,
    sheet: &str,
) -> PyResult<(u32, u32)> {
    ooxml_util::validate_zip_archive(&mut zip)?;
    let sheet_path = resolve_sheet_xml_path(&mut zip, sheet)?;
    let entry = zip.by_name(&sheet_path).map_err(|e| {
        PyErr::new::<PyIOError, _>(format!("Failed to open sheet part '{sheet_path}': {e}"))
    })?;
    scan_cell_bounds(BufReader::with_capacity(SHEET_ENTRY_BUFFER, entry))
}

fn archive_row_tags_exceed<R: Read + Seek>(
    mut zip: ZipArchive<R>,
    sheet: &str,
    limit: usize,
) -> PyResult<bool> {
    ooxml_util::validate_zip_archive(&mut zip)?;
    let sheet_path = resolve_sheet_xml_path(&mut zip, sheet)?;
    let entry = zip.by_name(&sheet_path).map_err(|e| {
        PyErr::new::<PyIOError, _>(format!("Failed to open sheet part '{sheet_path}': {e}"))
    })?;
    row_tags_exceed(entry, limit).map_err(|e| {
        PyErr::new::<PyIOError, _>(format!("Failed to read sheet part '{sheet_path}': {e}"))
    })
}

/// Whether `xml` holds more than `limit` matches of
/// `<(?:[A-Za-z0-9_]+:)?row(?:\s|>)`, counting each tag once even when it
/// straddles two reads.
fn row_tags_exceed<R: Read>(mut xml: R, limit: usize) -> io::Result<bool> {
    let mut buf = vec![0u8; SHEET_ENTRY_BUFFER * 4];
    // `buf[..kept]` holds the undecided tail of the previous read.
    let mut kept = 0usize;
    let mut count = 0usize;
    loop {
        if kept == buf.len() {
            buf.resize(buf.len() * 2, 0);
        }
        let read = xml.read(&mut buf[kept..])?;
        let eof = read == 0;
        let end = kept + read;
        let mut resume = end;
        for at in memchr::memchr_iter(b'<', &buf[..end]) {
            match row_tag_at(&buf[at..end]) {
                Some(true) => {
                    count += 1;
                    if count > limit {
                        return Ok(true);
                    }
                }
                Some(false) => {}
                // At end of input an unfinished tag cannot match.
                None if eof => {}
                None => {
                    resume = at;
                    break;
                }
            }
        }
        if eof {
            return Ok(false);
        }
        buf.copy_within(resume..end, 0);
        kept = end - resume;
    }
}

/// Whether the `<` that starts `s` opens a `row` element, optionally
/// namespace-prefixed; `None` when `s` ends before that is decided.
fn row_tag_at(s: &[u8]) -> Option<bool> {
    let name_end = s[1..]
        .iter()
        .position(|b| !(b.is_ascii_alphanumeric() || *b == b'_'))
        .map(|offset| offset + 1)?;
    let local = if s[name_end] == b':' && name_end > 1 {
        name_end + 1
    } else {
        1
    };
    let tag = s.get(local..local + 4)?;
    Some(&tag[..3] == b"row" && matches!(tag[3], b' ' | b'\t' | b'\n' | b'\r' | 0x0b | 0x0c | b'>'))
}

/// Largest row and column of any `<c>` (at its reference, else one column
/// past the previous cell of the row) or `<mergeCell>` range, mirroring where
/// openpyxl's eager reader places cells. Cell values are never decoded.
fn scan_cell_bounds<B: io::BufRead>(xml: B) -> PyResult<(u32, u32)> {
    let mut reader = XmlReader::from_reader(xml);
    let mut buf = Vec::with_capacity(8192);
    let (mut max_row, mut max_col) = (0u32, 0u32);
    let (mut row, mut col) = (0u32, 0u32);
    loop {
        buf.clear();
        let event = reader
            .read_event_into(&mut buf)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Cell bounds scan XML: {e}")))?;
        let e = match event {
            Event::Start(e) | Event::Empty(e) => e,
            Event::Eof => return Ok((max_row, max_col)),
            _ => continue,
        };
        match e.local_name().as_ref() {
            "row" => {
                row = row_index_from_start(&e, row)?;
                col = 0;
            }
            "c" => {
                let (cell_row, cell_col) = match local_attr(&e, "r").filter(|r| !r.is_empty()) {
                    Some(reference) => parse_a1_cell(reference.as_bytes()).ok_or_else(|| {
                        PyErr::new::<PyValueError, _>(format!(
                            "Invalid A1 reference: {reference:?}"
                        ))
                    })?,
                    None => (row, col.saturating_add(1)),
                };
                col = cell_col;
                max_row = max_row.max(cell_row);
                max_col = max_col.max(cell_col);
            }
            "mergeCell" => {
                if let Some((merge_row, merge_col)) = local_attr(&e, "ref")
                    .and_then(|reference| merge_range_max(reference.as_bytes()))
                {
                    max_row = max_row.max(merge_row);
                    max_col = max_col.max(merge_col);
                }
            }
            _ => {}
        }
    }
}

/// Value of the attribute whose local name is `name`, with character and
/// entity references decoded as an XML parser reports it; a value that does
/// not decode is returned raw and fails reference parsing downstream.
fn local_attr<'a>(e: &'a BytesStart<'_>, name: &str) -> Option<std::borrow::Cow<'a, str>> {
    let value = e
        .attributes()
        .with_checks(false)
        .flatten()
        .find_map(|attr| {
            let key = attr.key.as_ref();
            let local = key.rsplit(':').next().unwrap_or(key);
            (local == name).then_some(attr.value)
        })?;
    let decoded = match quick_xml::escape::unescape(&value) {
        Ok(std::borrow::Cow::Owned(text)) => Some(text),
        _ => None,
    };
    Some(decoded.map_or(value, std::borrow::Cow::Owned))
}

/// `(row, col)` of a reference matching `[A-Za-z]+[0-9]+`, the form
/// `wolfxl._utils.a1_to_rowcol` accepts.
fn parse_a1_cell(bytes: &[u8]) -> Option<(u32, u32)> {
    let letters = bytes.iter().take_while(|b| b.is_ascii_alphabetic()).count();
    if letters == 0 {
        return None;
    }
    let (col_part, row_part) = bytes.split_at(letters);
    let mut col = 0u32;
    for &byte in col_part {
        col = col
            .checked_mul(26)?
            .checked_add(u32::from(byte.to_ascii_uppercase() - b'A' + 1))?;
    }
    Some((parse_u32_bytes(row_part)?, col))
}

/// Bottom-right corner of a merged range such as `A1:C3` or `$B$2`; `None`
/// for a malformed range, which openpyxl's bounds also skip.
fn merge_range_max(reference: &[u8]) -> Option<(u32, u32)> {
    let cleaned: Vec<u8> = reference.iter().copied().filter(|&b| b != b'$').collect();
    let (first, last) = match cleaned.iter().position(|&b| b == b':') {
        Some(split) => (&cleaned[..split], &cleaned[split + 1..]),
        None => (&cleaned[..], &cleaned[..]),
    };
    let (first_row, first_col) = parse_a1_cell(first)?;
    let (last_row, last_col) = parse_a1_cell(last)?;
    Some((first_row.max(last_row), first_col.max(last_col)))
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

/// Cell type tag for a cell without a value: openpyxl keeps the `t=`
/// attribute, and a number without a value is blank.
fn valueless_cell_type(t_attr: &CellTypeAttr) -> &'static str {
    match t_attr {
        CellTypeAttr::Number => "blank",
        CellTypeAttr::Shared => "s",
        CellTypeAttr::FormulaString => "str",
        CellTypeAttr::InlineString => "inlineStr",
        CellTypeAttr::Bool => "b",
        CellTypeAttr::Error => "e",
        CellTypeAttr::Date => "d",
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
    sst: &SharedStringTable,
) -> PyResult<(PyObjectOwned, &'static str)> {
    let v_text = extract_inner_text(inner, b"v");
    let f_text = extract_inner_text(inner, b"f");
    let is_text = extract_is_text(inner);
    let t_attr = cell_type_from_bytes(t_attr.as_bytes());

    let raw_value = if matches!(t_attr, CellTypeAttr::InlineString) {
        is_text
    } else {
        v_text
    };
    build_cell_object(py, &t_attr, f_text, raw_value, sst)
}

/// Resolve a values-only row cell: an error reads as its code, a formula as
/// its `=`-prefixed text or, for `data_only`, its cached result (an empty
/// cache reads as `None`).
#[allow(clippy::too_many_arguments)]
fn value_cell_object(
    py: Python<'_>,
    t_attr: &CellTypeAttr,
    output: CellOutput,
    has_formula: bool,
    formula: Option<String>,
    fast_number: Option<PyObjectOwned>,
    fast_shared: Option<usize>,
    raw_value: Option<String>,
    sst: &SharedStringTable,
) -> PyResult<(PyObjectOwned, &'static str)> {
    if has_formula && output == CellOutput::FormulaText {
        let formula = formula.unwrap_or_default();
        return Ok((str_to_py(py, &formula_with_eq(&formula))?, "formula"));
    }
    if matches!(t_attr, CellTypeAttr::Error) {
        let code = raw_value.filter(|v| !v.is_empty());
        let value = match code {
            Some(code) => str_to_py(py, &code)?,
            None => py.None(),
        };
        return Ok((value, "e"));
    }
    if let Some(value) = fast_number {
        return Ok((value, "n"));
    }
    if !has_formula {
        if let Some(idx) = fast_shared {
            return Ok((sst.get_py(py, idx)?, "s"));
        }
        return build_cell_object(py, t_attr, None, raw_value, sst);
    }
    let cached = match fast_shared {
        Some(idx) => sst.get_py(py, idx)?,
        None => {
            let raw_value =
                raw_value.filter(|v| !v.is_empty() || matches!(t_attr, CellTypeAttr::InlineString));
            cached_value_to_py(py, t_attr, raw_value, sst)?
        }
    };
    let bound = cached.bind(py);
    if bound.is_exact_instance_of::<PyString>() && bound.len()? == 0 {
        return Ok((py.None(), "formula"));
    }
    // A numeric cache keeps date-style conversion.
    let kind = if matches!(t_attr, CellTypeAttr::Number) {
        "n"
    } else {
        "formula"
    };
    Ok((cached, kind))
}

fn formula_with_eq(formula: &str) -> std::borrow::Cow<'_, str> {
    if formula.starts_with('=') {
        std::borrow::Cow::Borrowed(formula)
    } else {
        std::borrow::Cow::Owned(format!("={formula}"))
    }
}

fn build_cell_object(
    py: Python<'_>,
    t_attr: &CellTypeAttr,
    formula: Option<String>,
    raw_value: Option<String>,
    sst: &SharedStringTable,
) -> PyResult<(PyObjectOwned, &'static str)> {
    // Value text as openpyxl reads it: an empty `<v>` is no value.
    let raw_value =
        raw_value.filter(|v| !v.is_empty() || matches!(t_attr, CellTypeAttr::InlineString));
    if let Some(formula) = formula {
        let cached = cached_value_to_py(py, t_attr, raw_value, sst)?;
        let payload = build_formula_dict(py, &formula, cached)?;
        if !matches!(t_attr, CellTypeAttr::Number) {
            payload
                .bind(py)
                .cast::<PyDict>()?
                .set_item("cached_type", valueless_cell_type(t_attr))?;
        }
        return Ok((payload, "formula"));
    }
    match t_attr {
        CellTypeAttr::Error => {
            let Some(v) = raw_value else {
                return Ok((py.None(), "e"));
            };
            let d = new_dict(py)?;
            set_dict_item(&d, "type", str_to_py(py, "error")?)?;
            set_dict_item(&d, "value", str_to_py(py, &v)?)?;
            Ok((d.into_any().unbind(), "e"))
        }
        CellTypeAttr::Number => {
            let Some(v) = raw_value else {
                return Ok((py.None(), "blank"));
            };
            if let Ok(i) = v.parse::<i64>() {
                // Excel-stored ints: surface as int when round-trip safe.
                return Ok((i64_to_py(py, i)?, "n"));
            }
            let f: f64 = v
                .parse()
                .map_err(|_| PyErr::new::<PyValueError, _>(format!("Bad numeric value: {v:?}")))?;
            Ok((f64_to_py(py, f)?, "n"))
        }
        _ => {
            let tag = match t_attr {
                CellTypeAttr::Shared => "s",
                CellTypeAttr::FormulaString => "str",
                CellTypeAttr::InlineString => "inlineStr",
                CellTypeAttr::Bool => "b",
                // Text that reads as no date or time stays a string.
                _ if raw_value.as_deref().is_some_and(|v| !is_iso8601(v)) => "s",
                _ => "d",
            };
            Ok((cached_value_to_py(py, t_attr, raw_value, sst)?, tag))
        }
    }
}

/// Python value of a cell's `<v>` (or inline `<is>`) text by its type,
/// following openpyxl's worksheet reader: shared-string indexes may carry
/// surrounding whitespace, booleans are `bool(int(v))`, numbers are `int`
/// unless they hold a decimal point or exponent, and missing text is `None`.
fn cached_value_to_py(
    py: Python<'_>,
    t_attr: &CellTypeAttr,
    raw_value: Option<String>,
    sst: &SharedStringTable,
) -> PyResult<PyObjectOwned> {
    let Some(v) = raw_value else {
        return Ok(py.None());
    };
    match t_attr {
        CellTypeAttr::Shared => {
            let idx: usize = v
                .trim()
                .parse()
                .map_err(|_| PyErr::new::<PyValueError, _>(format!("Bad SST index: {v:?}")))?;
            sst.get_py(py, idx)
        }
        CellTypeAttr::Bool => match v.trim().parse::<i64>() {
            Ok(n) => Ok((n != 0).into_py_any(py)?),
            Err(_) => Ok(matches!(v.trim(), "true" | "TRUE").into_py_any(py)?),
        },
        CellTypeAttr::Number => match number_text_to_py(py, v.trim().as_bytes())? {
            Some(number) => Ok(number),
            None => str_to_py(py, &v),
        },
        CellTypeAttr::Date => match iso8601_to_py(py, &v)? {
            Some(value) => Ok(value),
            None => str_to_py(py, &v),
        },
        CellTypeAttr::FormulaString | CellTypeAttr::InlineString | CellTypeAttr::Error => {
            str_to_py(py, &v)
        }
    }
}

fn build_formula_dict(
    py: Python<'_>,
    formula: &str,
    cached: PyObjectOwned,
) -> PyResult<PyObjectOwned> {
    let d = new_dict(py)?;
    let f_with_eq = formula_with_eq(formula);
    set_dict_item(&d, "type", str_to_py(py, "formula")?)?;
    set_dict_item(&d, "formula", str_to_py(py, &f_with_eq)?)?;
    set_dict_item(&d, "cached", cached)?;
    Ok(d.into_any().unbind())
}

fn number_text_to_py(py: Python<'_>, bytes: &[u8]) -> PyResult<Option<PyObjectOwned>> {
    let Ok(text) = std::str::from_utf8(bytes) else {
        return Ok(None);
    };
    if let Ok(i) = text.parse::<i64>() {
        return Ok(Some(i64_to_py(py, i)?));
    }
    if let Ok(f) = text.parse::<f64>() {
        return Ok(Some(f64_to_py(py, f)?));
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
            let (obj, kind) =
                parse_cell_inner(py, b"<v>42</v>", "n", &SharedStringTable::new(Vec::new()))
                    .unwrap();
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

    fn bytes_source(data: &[u8]) -> EntrySource {
        pyo3::Python::initialize();
        Python::attach(|py| EntrySource::Bytes {
            data: PyBytes::new(py, data).unbind(),
            pos: 0,
            end: data.len(),
        })
    }

    fn deflated_entry(xml: &[u8]) -> SheetEntryReader {
        let mut encoder =
            flate2::write::DeflateEncoder::new(Vec::new(), flate2::Compression::default());
        std::io::Write::write_all(&mut encoder, xml).unwrap();
        SheetEntryReader::new(
            bytes_source(&encoder.finish().unwrap()),
            CompressionMethod::Deflated,
            xml.len() as u64,
            crc32fast::hash(xml),
        )
        .unwrap()
    }

    /// An event reader over `xml` (inflated from a ZIP entry when
    /// `use_entry`) that refills three bytes at a time.
    fn tiny_buffer_events(xml: &[u8], use_entry: bool) -> SheetXmlReader {
        let scanner = if use_entry {
            RowScanner::new(Vec::new(), Some(Box::new(deflated_entry(xml))))
        } else {
            RowScanner::new(xml.to_vec(), None)
        };
        SheetXmlReader::Events(XmlReader::from_reader(BufReader::with_capacity(
            3,
            scanner.handoff(),
        )))
    }

    #[test]
    fn sheet_entry_reader_streams_stored_and_deflated_parts() {
        let xml: Vec<u8> = (0..20_000)
            .flat_map(|i| format!("<row r=\"{i}\"/>").into_bytes())
            .collect();
        let stored = SheetEntryReader::new(
            bytes_source(&xml),
            CompressionMethod::Stored,
            xml.len() as u64,
            crc32fast::hash(&xml),
        )
        .unwrap();
        for entry in [stored, deflated_entry(&xml)] {
            let mut out = Vec::new();
            BufReader::with_capacity(7, entry)
                .read_to_end(&mut out)
                .unwrap();
            assert_eq!(out, xml);
        }
    }

    #[test]
    fn sheet_entry_reader_enforces_declared_size_and_crc() {
        let xml = b"<worksheet><sheetData/></worksheet>";
        let mut entry = deflated_entry(xml);
        entry.declared_size -= 1;
        let err = io::copy(&mut entry, &mut io::sink()).unwrap_err();
        assert!(
            err.to_string().contains("larger than its declared"),
            "{err}"
        );

        let mut entry = deflated_entry(xml);
        entry.expected_crc ^= 1;
        let err = io::copy(&mut entry, &mut io::sink()).unwrap_err();
        assert_eq!(err.to_string(), "Invalid checksum");
    }

    #[test]
    fn entry_source_rejects_ranges_past_the_package() {
        pyo3::Python::initialize();
        Python::attach(|py| {
            let data = PyBytes::new(py, &[0u8; 8]);
            let mut source = PyBytesCursor::new(&data).into_entry_source(4, 4).unwrap();
            let mut out = Vec::new();
            source.read_to_end(&mut out).unwrap();
            assert_eq!(out, [0u8; 4]);
            for (start, len) in [(4, 5), (9, 0), (u64::MAX, 1)] {
                assert!(PyBytesCursor::new(&data)
                    .into_entry_source(start, len)
                    .is_err());
            }
        });
    }

    #[test]
    fn streaming_split_references_are_complete_values() {
        pyo3::Python::initialize();
        Python::attach(|py| {
            for use_entry in [false, true] {
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
                    // A tiny read buffer splits references across refills.
                    let mut reader = tiny_buffer_events(xml.as_bytes(), use_entry);
                    let mut buf = Vec::new();
                    assert!(matches!(
                        reader.read_event_into(&mut buf).unwrap(),
                        Event::Start(_)
                    ));
                    let sst =
                        SharedStringTable::new((0..11).map(|i| format!("entry {i}")).collect());
                    let (value, _) = read_cell_contents(
                        &mut reader,
                        &mut buf,
                        py,
                        &kind,
                        &sst,
                        CellOutput::Payload,
                    )
                    .unwrap();
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
                let mut reader = SheetXmlReader::new(xml.into_bytes(), None);
                let mut buf = Vec::new();
                assert!(matches!(
                    reader.read_event_into(&mut buf).unwrap(),
                    Event::Start(_)
                ));
                assert!(read_cell_contents(
                    &mut reader,
                    &mut buf,
                    py,
                    &CellTypeAttr::Number,
                    &SharedStringTable::new(Vec::new()),
                    CellOutput::Payload,
                )
                .is_err());
            }
        });
    }

    #[test]
    fn streaming_formula_and_cache_aggregate_references() {
        pyo3::Python::initialize();
        Python::attach(|py| {
            let read = |output| {
                let xml = b"<c><f> A1 &amp; &quot; B &quot;&#10;</f><v>1&#50;3</v></c>";
                let mut reader = SheetXmlReader::new(xml.to_vec(), None);
                let mut buf = Vec::new();
                assert!(matches!(
                    reader.read_event_into(&mut buf).unwrap(),
                    Event::Start(_)
                ));
                read_cell_contents(
                    &mut reader,
                    &mut buf,
                    py,
                    &CellTypeAttr::Number,
                    &SharedStringTable::new(Vec::new()),
                    output,
                )
                .unwrap()
            };
            let (value, kind) = read(CellOutput::Payload);
            assert_eq!(kind, "formula");
            let fields = value.bind(py).cast::<PyDict>().unwrap();
            let formula: String = fields
                .get_item("formula")
                .unwrap()
                .unwrap()
                .extract()
                .unwrap();
            let cached: i64 = fields
                .get_item("cached")
                .unwrap()
                .unwrap()
                .extract()
                .unwrap();
            assert_eq!(formula, "= A1 & \" B \"\n");
            assert_eq!(cached, 123);

            let (text, _) = read(CellOutput::FormulaText);
            assert_eq!(text.extract::<String>(py).unwrap(), "= A1 & \" B \"\n");
            let (cached, kind) = read(CellOutput::CachedValue);
            assert_eq!(cached.extract::<i64>(py).unwrap(), 123);
            assert_eq!(kind, "n");
        });
    }

    #[test]
    fn cell_bounds_follow_cells_and_merges_not_dimension() {
        let xml = br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:Z900"/><sheetData><row r="1"><c r="A1"><v>1</v></c><c><v>2</v></c><c r=""/></row><row><c r="b4"/></row><row r="9" s="2" customFormat="1"/></sheetData><mergeCells><mergeCell ref="$B$5:B6"/><mergeCell ref="bad"/></mergeCells></worksheet>"#;
        assert_eq!(scan_cell_bounds(&xml[..]).unwrap(), (6, 3));
        assert_eq!(scan_cell_bounds(&b"<worksheet/>"[..]).unwrap(), (0, 0));
        assert!(
            scan_cell_bounds(&br#"<sheetData><row><c r="A1B"/></row></sheetData>"#[..]).is_err()
        );
        let encoded = br#"<sheetData><row r="1"><c r="A5000&#49;"/></row></sheetData><mergeCells><mergeCell ref="A1:C&#x37;"/></mergeCells>"#;
        assert_eq!(scan_cell_bounds(&encoded[..]).unwrap(), (50001, 3));
    }

    /// Reads at most `step` bytes per call, so tags straddle reads.
    struct Trickle<'a> {
        data: &'a [u8],
        step: usize,
    }

    impl Read for Trickle<'_> {
        fn read(&mut self, out: &mut [u8]) -> io::Result<usize> {
            let n = self.step.min(out.len()).min(self.data.len());
            out[..n].copy_from_slice(&self.data[..n]);
            self.data = &self.data[n..];
            Ok(n)
        }
    }

    #[test]
    fn row_tags_counted_once_across_reads() {
        // Three rows match; `<rows>`, `<rowBreaks`, `<:row>` and a trailing
        // unfinished `<row` do not.
        let xml = b"<sheetData><row r=\"1\"/><x:row>\n</x:row><row\t/><rows><rowBreaks/><:row></sheetData><row";
        for step in [1, 2, 3, 7, xml.len()] {
            let probe = |limit| row_tags_exceed(Trickle { data: xml, step }, limit).unwrap();
            assert!(probe(2), "step {step}");
            assert!(!probe(3), "step {step}");
        }
        let many = b"<row>".repeat(50_001);
        assert!(row_tags_exceed(&many[..], 50_000).unwrap());
        assert!(!row_tags_exceed(&many[5..], 50_000).unwrap());
    }

    /// Reader over `xml` resolving shared strings against `sst`, through the
    /// row scanner when `fast`, else through quick-xml events only.
    fn test_reader(
        xml: &str,
        sst: Arc<SharedStringTable>,
        output: CellOutput,
        bounds: [Option<u32>; 4],
        fast: bool,
    ) -> StreamingSheetReader {
        let reader = if fast {
            SheetXmlReader::new(xml.as_bytes().to_vec(), None)
        } else {
            let mut events = SheetXmlReader::new(xml.as_bytes().to_vec(), None);
            events.events();
            events
        };
        let [min_row, max_row, min_col, max_col] = bounds;
        StreamingSheetReader {
            reader: Some(reader),
            event_buf: Vec::new(),
            sst,
            exhausted: false,
            min_row,
            max_row,
            stopped_past_max_row: false,
            min_col,
            max_col,
            last_row: 0,
            date_styles: DateStyles::default(),
            value_output: output,
        }
    }

    /// Value rows as `repr` strings (or the error), read through the row
    /// scanner when `fast`, else through quick-xml events only.
    fn value_rows(
        py: Python<'_>,
        xml: &str,
        output: CellOutput,
        bounds: [Option<u32>; 4],
        fast: bool,
    ) -> Vec<String> {
        let sst = Arc::new(SharedStringTable::new(
            (0..3).map(|i| format!("s{i}")).collect(),
        ));
        let mut sheet = test_reader(xml, sst, output, bounds, fast);
        let mut rows = Vec::new();
        loop {
            match sheet.parse_one_value_row(py) {
                Ok(ValueStepResult::Row(row, values)) => {
                    rows.push(format!("{row}: {}", values.repr().unwrap()));
                }
                Ok(ValueStepResult::Skip) => rows.push("skip".into()),
                Ok(ValueStepResult::Done) => break,
                Err(err) => {
                    rows.push(format!("error: {err}"));
                    break;
                }
            }
        }
        rows.push(format!("past max: {}", sheet.stopped_past_max_row));
        rows
    }

    #[test]
    fn shared_string_references_share_one_python_str() {
        let xml = concat!(
            "<worksheet><sheetData><row r=\"1\">",
            "<c r=\"A1\" t=\"s\"><v>1</v></c><c r=\"B1\" t=\"s\"><v> 1 </v></c>",
            "<c r=\"C1\" t=\"s\"><f>X1</f><v>1</v></c></row>",
            "<row r=\"2\"><c r=\"A2\" t=\"s\"><v>1</v></c><c r=\"B2\" t=\"s\"><v>7</v></c></row>",
            "</sheetData></worksheet>",
        );
        pyo3::Python::initialize();
        Python::attach(|py| {
            let table = Arc::new(SharedStringTable::new(vec!["s0".into(), "x".repeat(4096)]));
            let expected = table.get_py(py, 1).unwrap();
            assert!(expected.bind(py).is(table.get_py(py, 1).unwrap()));
            assert_eq!(
                table.get_py(py, 7).unwrap().extract::<String>(py).unwrap(),
                ""
            );

            // Every reader sharing the table, on both values-only paths and
            // the cell path, returns the one cached object for entry 1.
            let mut matches = 0;
            for output in [CellOutput::FormulaText, CellOutput::CachedValue] {
                for fast in [true, false] {
                    let mut sheet = test_reader(xml, Arc::clone(&table), output, [None; 4], fast);
                    while let ValueStepResult::Row(_, values) =
                        sheet.parse_one_value_row(py).unwrap()
                    {
                        for value in values.iter() {
                            if value.eq(&expected).unwrap() {
                                assert!(value.is(&expected));
                                matches += 1;
                            }
                        }
                    }
                }
            }
            // A1, B1, A2 per formula-text pass; C1's cached result as well
            // per cached-value pass.
            assert_eq!(matches, 2 * 3 + 2 * 4);

            let mut sheet = test_reader(
                xml,
                Arc::clone(&table),
                CellOutput::Payload,
                [None; 4],
                false,
            );
            let mut payload_matches = 0;
            while let StepResult::Row(_, cells) = sheet.parse_one_row(py).unwrap() {
                for cell in cells {
                    let value = cell.value.bind(py);
                    let value = match value.cast::<PyDict>() {
                        Ok(payload) => payload.get_item("cached").unwrap().unwrap(),
                        Err(_) => value.clone(),
                    };
                    if value.eq(&expected).unwrap() {
                        assert!(value.is(&expected));
                        payload_matches += 1;
                    }
                }
            }
            assert_eq!(payload_matches, 4);
        });
    }

    #[test]
    fn scanned_value_rows_match_event_rows() {
        let rows = [
            r#"<row r="1" spans="1:4"><c r="A1"><v>1.5</v></c><c r="B1" t="s"><v>2</v></c><c r="C1" t="b"><v>1</v></c><c r="D1" s="3"/></row>"#,
            r#"<row r="2"><c r="A2"><f t="shared" ref="A2:A3" si="0">B2*2</f><v>4</v></c><c r="C2"><f t="shared" si="0"/><v>6</v></c><c r="D2" t="str"><f>"a"&amp;"b"</f><v>ab</v></c></row>"#,
            r#"<row r="3"><c r="A3" t="e"><f>1/0</f><v>#DIV/0!</v></c><c r="B3"><f>X1</f><v/></c><c r="C3"><f></f><v>7</v></c><c r="D3"><v>1&#50;3</v></c></row>"#,
            r#"<row><c><v>8</v></c><c><v>9</v></c>
  <c r="F4" t="s"><v>99</v></c><c r="G4"><f>A1&#13;
B1</f><v>0</v></c></row>"#,
            r#"<row r="6"/><row r="7"><c r="A7" t="inlineStr"><is><t>inline</t></is></c></row><row r="8"><c r="A8"><v>10</v></c></row>"#,
            r#"<row r="9"><!-- note --><c r="A9"><v>11</v></c></row><row r="10"><c r="B10"><v>bad</v></c></row>"#,
            r#"<row r="13"><c r="A13" t="str"><f>IF(B1&gt;0,&quot;x&#x41;&quot;,&apos;S&apos;!A1)&#xD;&#10;</f><v>xA</v></c><c r="B13"><f>&lt;&amp;</f>  <v>1</v> </c></row>"#,
            r#"<row r="17"><c r="A17" t="inlineStr"><is/></c><c r="B17" t="inlineStr"><is></is></c><c r="C17" t="inlineStr"><is><t/></is></c><c r="D17" t="inlineStr"><is> <t xml:space="preserve"> a&amp;b&#10;</t>
<t>c</t></is></c><c r="E17"><is><t>9</t></is><v>3</v></c><c r="F17" t="inlineStr"><v>1</v><is><t>x</t></is></c><c r="G17" t="inlineStr"/><c r="H17" t="inlineStr"><is><t>&bad;</t></is></c></row>"#,
            r#"<row r="18"><c r="A18" t="inlineStr"><is><t>a</t></is></c><c r="B18" t="inlineStr"><is><r><t>rich</t></r></is></c></row><row r="19"><c r="A19" t="inlineStr"><is><t>a</t><rPh sb="0" eb="1"><t>b</t></rPh></is></c></row>"#,
        ];
        let sheets = [
            format!("<worksheet><sheetData>{}</sheetData></worksheet>", rows.join("")),
            format!(
                "<?xml version=\"1.0\"?>\n<x:worksheet xmlns:x=\"urn:x\"><x:sheetData>{}</x:sheetData></x:worksheet>",
                rows[0].replace("<row", "<x:row").replace("</row>", "</x:row>")
            ),
            format!("<worksheet><sheetData>{}<row r=\"11\"><c r=\"1A\"/></row></sheetData></worksheet>", rows[0]),
            format!("<worksheet><sheetData>{}<row r=\"12\"><c r=\"A12\"><v>1</v></c>", rows[1]),
            format!("<worksheet><sheetData>{}<row r=\"x\"/></sheetData></worksheet>", rows[2]),
            format!("<worksheet><sheetData>{}<row r=\"14\"><c r=\"A14\"><f>A1&nbsp;</f><v>1</v></c></row></sheetData></worksheet>", rows[6]),
            format!("<worksheet><sheetData>{}<row r=\"15\"><c r=\"A15\"><f>A1&amp</f><v>1</v></c></row></sheetData></worksheet>", rows[6]),
            format!("<worksheet><sheetData>{}<row r=\"16\"><c r=\"A16\"><f>A1 & B1</f><v>1</v></c></row></sheetData></worksheet>", rows[6]),
            format!("<worksheet><sheetData>{}{}</sheetData></worksheet>", rows[7].replace("<c r=\"H17\" t=\"inlineStr\"><is><t>&bad;</t></is></c>", ""), rows[8]),
            // Skipped rows (min_row 3) whose first `</row>` bytes are not the end tag.
            "<worksheet><sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c><!-- </row> --></row><row r=\"2\"><c r=\"A2\" t=\"str\"><f>\"a\"&amp;\"b\"</f><v>ab</v></c></row><row r=\"3\"><c r=\"A3\"><v>3</v></c></row><row r=\"4\"><c r=\"D4\"><v>4</v></c></row></sheetData></worksheet>".to_string(),
            "<worksheet><sheetData><row r=\"1\"><c r=\"A1\" x=\"</row>\"><v>1</v></c></row><row r=\"3\"><c r=\"D3\"><v>3</v></c></row></sheetData></worksheet>".to_string(),
            "<worksheet><sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t><![CDATA[</row>]]></t></is></c></row><row r=\"3\"><c r=\"D3\"><v>3</v></c></row></sheetData></worksheet>".to_string(),
            "<worksheet><sheetData><row r=\"2\"><c r=\"A2\" t=\"str\"><f>A2 & B2</f></c></row><row r=\"3\"><c r=\"D3\"><v>3</v></c></row></sheetData></worksheet>".to_string(),
            "<worksheet><sheetData><row r=\"2\"><c r=\"A2\" t=\"str\"><f>&a&b;</f></c></row><row r=\"3\"><c r=\"D3\"><v>3</v></c></row></sheetData></worksheet>".to_string(),
            format!("<worksheet><sheetData>{}</sheetData></worksheet>", rows[7]),
        ];
        pyo3::Python::initialize();
        Python::attach(|py| {
            for xml in &sheets {
                for output in [CellOutput::FormulaText, CellOutput::CachedValue] {
                    for bounds in [
                        [None, None, None, None],
                        [Some(2), Some(7), Some(2), Some(5)],
                        [Some(3), None, Some(4), None],
                    ] {
                        let fast = value_rows(py, xml, output, bounds, true);
                        let events = value_rows(py, xml, output, bounds, false);
                        assert_eq!(fast, events, "{xml}");
                    }
                }
            }
        });
    }
}
