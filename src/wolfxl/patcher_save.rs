//! Shared save-orchestration state for `XlsxPatcher`.

use std::collections::{HashMap, HashSet};
use std::fs::File;

use pyo3::exceptions::PyIOError;
use pyo3::prelude::*;
use zip::ZipArchive;

use wolfxl_merger::SheetBlock;

use crate::ooxml_util;

/// Open the source workbook as a ZIP archive with consistent PyO3 errors.
pub(super) fn open_source_zip(file_path: &str) -> PyResult<ZipArchive<File>> {
    let file = File::open(file_path)
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("Cannot open '{file_path}': {e}")))?;
    let mut zip = ZipArchive::new(file)
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("ZIP read error: {e}")))?;
    ooxml_util::validate_zip_archive(&mut zip)?;
    Ok(zip)
}

/// Mutable workspace threaded through the ordered save phases.
///
/// Keeping this state in one struct makes `XlsxPatcher::do_save` easier to
/// split without changing phase order or the public PyO3 surface.
pub(super) struct SaveWorkspace {
    pub(super) file_patches: HashMap<String, Vec<u8>>,
    pub(super) part_id_allocator: wolfxl_rels::PartIdAllocator,
    pub(super) cloned_table_names: HashSet<String>,
    pub(super) local_blocks: HashMap<String, Vec<SheetBlock>>,
}

impl SaveWorkspace {
    pub(super) fn new(
        zip: &mut ZipArchive<File>,
        queued_blocks: &HashMap<String, Vec<SheetBlock>>,
        seed_part_id_allocator: bool,
    ) -> Self {
        let part_id_allocator = if seed_part_id_allocator {
            let names: Vec<String> = (0..zip.len())
                .filter_map(|i| {
                    let entry = zip.by_index(i).ok()?;
                    entry.name().ok().map(|name| name.into_owned())
                })
                .collect();
            wolfxl_rels::PartIdAllocator::from_zip_parts(names.iter().map(|s| s.as_str()))
        } else {
            wolfxl_rels::PartIdAllocator::new()
        };

        Self {
            file_patches: HashMap::new(),
            part_id_allocator,
            cloned_table_names: HashSet::new(),
            local_blocks: queued_blocks.clone(),
        }
    }
}
