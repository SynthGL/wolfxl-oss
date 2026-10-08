//! Cached merge-only metadata; independent of worksheet cell/index hydration.

use crate::native_reader_backend::NativeXlsxBook;
use crate::native_reader_dimensions::parse_range_1based;
use pyo3::exceptions::{PyIOError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::PyDict;
use std::collections::HashMap;
use wolfxl_reader::BorderInfo;

#[derive(Default)]
pub(crate) struct MergeMetadata {
    ranges: Vec<String>,
    bounds: Vec<(u32, u32, u32, u32)>,
    endpoint_styles: Option<HashMap<(u32, u32), u32>>,
    extra_endpoint_styles: HashMap<(u32, u32), u32>,
}

fn ensure_metadata(book: &mut NativeXlsxBook, sheet: &str, styles: bool) -> PyResult<()> {
    if !book.sheet_names.iter().any(|name| name == sheet) {
        return Err(PyErr::new::<PyValueError, _>(format!(
            "Unknown sheet: {sheet}"
        )));
    }
    if !book.sheet_merge_metadata.contains_key(sheet) {
        let ranges = book.book.worksheet_merged_ranges(sheet).map_err(|e| {
            PyErr::new::<PyIOError, _>(format!("native merge metadata read failed: {e}"))
        })?;
        let bounds = ranges
            .iter()
            .filter_map(|r| parse_range_1based(r))
            .collect();
        book.sheet_merge_metadata.insert(
            sheet.to_string(),
            MergeMetadata {
                ranges,
                bounds,
                endpoint_styles: None,
                extra_endpoint_styles: HashMap::new(),
            },
        );
    }
    if styles && book.sheet_merge_metadata[sheet].endpoint_styles.is_none() {
        let ranges = &book.sheet_merge_metadata[sheet].ranges;
        let endpoints = book
            .book
            .worksheet_merge_endpoint_styles(sheet, ranges)
            .map_err(|e| {
                PyErr::new::<PyIOError, _>(format!("native merge endpoint read failed: {e}"))
            })?;
        book.sheet_merge_metadata
            .get_mut(sheet)
            .unwrap()
            .endpoint_styles = Some(endpoints);
    }
    Ok(())
}

pub(crate) fn read_ranges(book: &mut NativeXlsxBook, sheet: &str) -> PyResult<Vec<String>> {
    ensure_metadata(book, sheet, false)?;
    Ok(book.sheet_merge_metadata[sheet].ranges.clone())
}

/// Preserve the prior Python negative-source shortcut without tokenizing cells.
/// The exact read_ranges API still validates XML even on negative-probe sources.
pub(crate) fn read_ranges_if_present(
    book: &mut NativeXlsxBook,
    sheet: &str,
) -> PyResult<Vec<String>> {
    if !book.sheet_names.iter().any(|name| name == sheet) {
        return Err(PyErr::new::<PyValueError, _>(format!(
            "Unknown sheet: {sheet}"
        )));
    }
    if book.sheet_merge_metadata.contains_key(sheet) {
        return read_ranges(book, sheet);
    }
    if !book
        .book
        .worksheet_may_have_merged_cells(sheet)
        .map_err(|e| {
            PyErr::new::<PyIOError, _>(format!("native merge presence probe failed: {e}"))
        })?
    {
        return Ok(Vec::new());
    }
    read_ranges(book, sheet)
}

pub(crate) fn read_endpoint_style_ids(
    book: &mut NativeXlsxBook,
    sheet: &str,
) -> PyResult<Vec<(String, u32, u32)>> {
    ensure_metadata(book, sheet, true)?;
    let metadata = &book.sheet_merge_metadata[sheet];
    let styles = metadata.endpoint_styles.as_ref().unwrap();
    Ok(metadata
        .ranges
        .iter()
        .filter_map(|r| {
            let (r1, c1, r2, c2) = parse_range_1based(r)?;
            Some((
                r.clone(),
                *styles.get(&(r1, c1)).unwrap_or(&0),
                *styles.get(&(r2, c2)).unwrap_or(&0),
            ))
        })
        .collect())
}

pub(crate) fn merged_border(
    book: &mut NativeXlsxBook,
    sheet: &str,
    row: u32,
    col: u32,
) -> PyResult<Option<BorderInfo>> {
    ensure_metadata(book, sheet, false)?;
    let bounds = book.sheet_merge_metadata[sheet]
        .bounds
        .iter()
        .copied()
        .find(|(r1, c1, r2, c2)| row >= *r1 && row <= *r2 && col >= *c1 && col <= *c2);
    let Some((r1, c1, r2, c2)) = bounds else {
        return Ok(None);
    };
    ensure_metadata(book, sheet, true)?;
    let styles = book.sheet_merge_metadata[sheet]
        .endpoint_styles
        .as_ref()
        .unwrap();
    let mut anchor = book
        .book
        .border_for_style_id(*styles.get(&(r1, c1)).unwrap_or(&0))
        .cloned()
        .unwrap_or_default();
    if let Some(end) = book
        .book
        .border_for_style_id(*styles.get(&(r2, c2)).unwrap_or(&0))
    {
        if end.right.is_some() {
            anchor.right = end.right.clone();
        }
        if end.bottom.is_some() {
            anchor.bottom = end.bottom.clone();
        }
    }
    if row == r1 && col == c1 {
        return Ok(Some(anchor));
    }
    Ok(Some(BorderInfo {
        left: (col == c1).then(|| anchor.left.clone()).flatten(),
        right: (col == c2).then(|| anchor.right.clone()).flatten(),
        top: (row == r1).then(|| anchor.top.clone()).flatten(),
        bottom: (row == r2).then(|| anchor.bottom.clone()).flatten(),
        ..BorderInfo::default()
    }))
}

pub(crate) fn read_merged_cell_border(
    book: &mut NativeXlsxBook,
    py: Python<'_>,
    sheet: &str,
    row: u32,
    col: u32,
) -> PyResult<Py<PyAny>> {
    let Some(border) = merged_border(book, sheet, row, col)? else {
        return Ok(py.None());
    };
    let payload = PyDict::new(py);
    crate::native_reader_styles::populate_border(py, &payload, &border)?;
    Ok(payload.into())
}

/// Resolve sparse arbitrary endpoints used by newly-added live merge ranges.
pub(crate) fn read_sparse_endpoint_style_ids(
    book: &mut NativeXlsxBook,
    sheet: &str,
    positions: Vec<(u32, u32)>,
) -> PyResult<Vec<u32>> {
    ensure_metadata(book, sheet, false)?;
    let metadata = &book.sheet_merge_metadata[sheet];
    let missing: Vec<(u32, u32)> = positions
        .iter()
        .copied()
        .filter(|(row, col)| {
            *row > 0
                && *col > 0
                && !metadata.extra_endpoint_styles.contains_key(&(*row, *col))
                && !metadata
                    .endpoint_styles
                    .as_ref()
                    .is_some_and(|styles| styles.contains_key(&(*row, *col)))
        })
        .collect();
    if !missing.is_empty() {
        let refs: Vec<String> = missing
            .iter()
            .map(|(row, col)| crate::native_reader_dimensions::row_col_to_a1_1based(*row, *col))
            .collect();
        let found = book
            .book
            .worksheet_merge_endpoint_styles(sheet, &refs)
            .map_err(|e| {
                PyErr::new::<PyIOError, _>(format!("native sparse endpoint read failed: {e}"))
            })?;
        let metadata = book.sheet_merge_metadata.get_mut(sheet).unwrap();
        for position in missing {
            metadata
                .extra_endpoint_styles
                .insert(position, *found.get(&position).unwrap_or(&0));
        }
    }
    let metadata = &book.sheet_merge_metadata[sheet];
    Ok(positions
        .iter()
        .map(|position| {
            metadata
                .endpoint_styles
                .as_ref()
                .and_then(|styles| styles.get(position))
                .or_else(|| metadata.extra_endpoint_styles.get(position))
                .copied()
                .unwrap_or(0)
        })
        .collect())
}
