//! Workbook-level helpers for the native writer backend.

use std::io::BufWriter;

use pyo3::exceptions::{PyIOError, PyValueError};
use pyo3::prelude::*;
use wolfxl_writer::model::Worksheet;
use wolfxl_writer::Workbook;

pub(crate) fn add_sheet_if_missing(wb: &mut Workbook, name: &str) {
    if wb.sheet_by_name(name).is_some() {
        return;
    }
    wb.add_sheet(Worksheet::new(name));
}

pub(crate) fn rename_sheet(wb: &mut Workbook, old_name: &str, new_name: &str) -> PyResult<()> {
    wb.rename_sheet(old_name, new_name.to_string())
        .map_err(PyValueError::new_err)
}

pub(crate) fn move_sheet(wb: &mut Workbook, name: &str, offset: isize) -> PyResult<()> {
    wb.move_sheet(name, offset).map_err(PyValueError::new_err)
}

pub(crate) fn remove_sheet(wb: &mut Workbook, name: &str) -> PyResult<()> {
    wb.remove_sheet(name).map_err(PyValueError::new_err)
}

pub(crate) fn save(wb: &mut Workbook, path: &str) -> PyResult<()> {
    // G20: flush per-sheet streaming BufWriters so the splice phase
    // inside `emit_xlsx → sheet_xml::emit` reads consistent bytes.
    crate::native_writer_streaming::finalize_all_streaming(wb)?;
    // RFC-073 v1.5: stream the ZIP container straight into a BufWriter<File>
    // instead of materialising the whole archive as `Vec<u8>` first. The
    // dominant memory cost during save is still the per-sheet emit `String`
    // (~150 MB for 1M-row × 5-col sheets) — `package` itself only cost the
    // size of the compressed archive (~25 MB at 1M rows). The win is small
    // but real on the disk-write peak; closing the larger sheet-body
    // materialisation requires plumbing `Write` all the way through
    // `sheet_xml::emit`, which is out of scope for v1.5.
    // Fresh writer saves produce a brand-new ZIP from the in-memory model.
    // Keep the atomic temp-file + rename behavior, but skip the heavier
    // fsync/reopen validation loop that the modify-mode patcher still uses.
    crate::atomic_save::write_zip_atomically_fast(path, |file| {
        let mut writer = BufWriter::new(file);
        wolfxl_writer::emit_xlsx_to(wb, &mut writer)
            .map_err(|e| PyIOError::new_err(format!("failed to write {path}: {e}")))?;
        use std::io::Write;
        writer
            .flush()
            .map_err(|e| PyIOError::new_err(format!("failed to flush {path}: {e}")))?;
        Ok(())
    })
}
