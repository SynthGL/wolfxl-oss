//! Streaming write-only bridge for `Workbook(write_only=True)` (G20).
//!
//! Three FFI entry points + an internal save-time hook glue the Rust
//! `StreamingSheet` temp-file machinery to the Python
//! `WriteOnlyWorksheet`. The flow per write-only workbook is:
//!
//! 1. Python creates a `WriteOnlyWorksheet`. `wb._backend.add_sheet(name)`
//!    runs first, then `wb._backend.enable_streaming_sheet(name)` swaps
//!    the freshly-created (and empty) `Worksheet` into streaming mode.
//! 2. Each `ws.append(row)` invokes
//!    `wb._backend.append_streaming_row(name, row_idx, cells)` with a
//!    Python list of cell payload dicts (or `None` to skip a column).
//!    Index in the list is `column - 1` (1-based on Excel side).
//! 3. `wb.save(path)` calls the native workbook save helper, which invokes
//!    `finalize_all_streaming` before `emit_xlsx` so each streaming
//!    `BufWriter` is flushed to disk and the splice phase reads valid
//!    bytes.
//!
//! Errors at every layer are mapped to `PyValueError` / `PyIOError` so
//! the Python side can convert them into `WorkbookAlreadySaved` /
//! `RuntimeError` as appropriate.

use pyo3::exceptions::{PyIOError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList};

use wolfxl_writer::model::cell::WriteCell;
use wolfxl_writer::model::worksheet::Row;
use wolfxl_writer::model::{FormatSpec, WriteCellValue};
use wolfxl_writer::streaming::StreamingSheet;
use wolfxl_writer::Workbook;

use crate::native_writer_cells::payload_to_write_cell_value;

/// Convert the named sheet into streaming mode. Idempotent.
///
/// The sheet must exist (created by an earlier `add_sheet`) and must
/// not already have any eager rows or merges installed — streaming and
/// eager modes are mutually exclusive.
pub(crate) fn enable_streaming(wb: &mut Workbook, name: &str) -> PyResult<()> {
    let idx = wb
        .sheets
        .iter()
        .position(|s| s.name == name)
        .ok_or_else(|| PyValueError::new_err(format!("Unknown sheet: {name}")))?;
    let ws = &mut wb.sheets[idx];

    if ws.streaming.is_some() {
        return Ok(());
    }
    if !ws.rows.is_empty() || ws.has_dense_rows() {
        return Err(PyValueError::new_err(
            "cannot enable streaming on a sheet with existing rows",
        ));
    }

    let stream = StreamingSheet::new(idx as u32)
        .map_err(|e| PyIOError::new_err(format!("streaming temp file open failed: {e}")))?;
    ws.streaming = Some(stream);
    Ok(())
}

/// Append one row to a streaming sheet's temp file.
///
/// `cells` is a Python list whose length is the row's column count.
/// Each item is either `None` (skip this column) or a payload dict
/// shaped like the eager `write_cell_value` payload. Optional
/// `style_id` integer key on the dict attaches a styles-builder index
/// resolved Python-side.
pub(crate) fn append_streaming_row(
    wb: &mut Workbook,
    sheet: &str,
    row_idx: u32,
    cells: &Bound<'_, PyList>,
) -> PyResult<()> {
    // Build the Row first without touching wb, then borrow `wb.sheets`
    // and `wb.sst` disjointly via direct field access (the borrow
    // checker allows this when sheet lookup is done by index).
    // Build the (col, value, style_id) list first so we can intern any
    // auto-default number formats (date/datetime) into wb.styles BEFORE
    // the disjoint-borrow split below.
    let mut staged: Vec<(u32, WriteCellValue, Option<u32>)> = Vec::with_capacity(cells.len());
    for (i, item) in cells.iter().enumerate() {
        if item.is_none() {
            continue;
        }
        let dict = item
            .cast::<PyDict>()
            .map_err(|_| PyValueError::new_err("each cell must be None or a dict"))?;
        let value = payload_to_write_cell_value(&item, wb.date1904)?;
        let mut style_id: Option<u32> = dict
            .get_item("style_id")?
            .map(|v| v.extract::<u32>())
            .transpose()?;
        // Mirror eager-path semantics: a `date`/`datetime` payload
        // auto-attaches a default number_format if the caller didn't
        // pass one. Skipping this would leave `WriteCellValue::DateSerial`
        // cells displaying as raw serial numbers on re-read.
        if style_id.is_none() {
            if let Some(nf) = infer_default_number_format(dict, &value)? {
                let spec = FormatSpec {
                    number_format: Some(nf.to_string()),
                    ..Default::default()
                };
                style_id = Some(wb.styles.intern_format(&spec));
            }
        }
        let col = (i as u32) + 1;
        staged.push((col, value, style_id));
    }

    let mut row = Row::default();
    for (col, value, style_id) in staged {
        let cell = match style_id {
            Some(s) => WriteCell::new(value).with_style(s),
            None => WriteCell::new(value),
        };
        row.cells.insert(col, cell);
    }

    let idx = wb
        .sheets
        .iter()
        .position(|s| s.name == sheet)
        .ok_or_else(|| PyValueError::new_err(format!("Unknown sheet: {sheet}")))?;
    let stream = wb.sheets[idx].streaming.as_mut().ok_or_else(|| {
        PyValueError::new_err(format!("sheet '{sheet}' is not in streaming mode"))
    })?;
    stream
        .append_row(row_idx, &row, &mut wb.sst)
        .map_err(|e| PyIOError::new_err(format!("streaming append failed: {e}")))
}

/// Pick a default `number_format` string for a date, datetime, duration or time
/// payload that didn't get an explicit `style_id`. Mirrors the eager
/// path in `native_writer_cells::write_cell_payload` so streaming and
/// eager date cells render identically when re-opened.
fn infer_default_number_format(
    dict: &Bound<'_, PyDict>,
    value: &WriteCellValue,
) -> PyResult<Option<&'static str>> {
    if !matches!(
        value,
        WriteCellValue::DateSerial(_) | WriteCellValue::Number(_)
    ) {
        return Ok(None);
    }
    let type_str: Option<String> = dict
        .get_item("type")?
        .map(|v| v.extract::<String>())
        .transpose()?;
    Ok(match (type_str.as_deref(), value) {
        (Some("date"), WriteCellValue::DateSerial(_)) => Some("yyyy-mm-dd"),
        (Some("datetime"), WriteCellValue::DateSerial(_)) => Some("yyyy-mm-dd hh:mm:ss"),
        (Some("timedelta"), WriteCellValue::Number(_)) => Some("[hh]:mm:ss"),
        (Some("time"), WriteCellValue::Number(_)) => Some("h:mm:ss"),
        _ => None,
    })
}

/// Flush every streaming sheet's `BufWriter` so the splice phase reads
/// bytes-on-disk consistently. Idempotent: a second call is a no-op.
pub(crate) fn finalize_all_streaming(wb: &mut Workbook) -> PyResult<()> {
    for sheet in wb.sheets.iter_mut() {
        if let Some(stream) = sheet.streaming.as_mut() {
            stream
                .finalize()
                .map_err(|e| PyIOError::new_err(format!("streaming finalize failed: {e}")))?;
        }
    }
    Ok(())
}
