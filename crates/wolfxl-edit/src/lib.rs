//! Byte-preserving edits of existing XLSX and XLSM workbooks.
//!
//! Open an [`EditSession`], apply edits, then atomically [`EditSession::save`].
//! Untouched ZIP entries are copied raw, including VBA and unknown extensions.

use std::collections::{BTreeMap, BTreeSet};
use std::io::{Cursor, Read, Write};
use std::path::Path;

use quick_xml::events::Event;
use wolfxl_rels::{RelsGraph, TargetMode};
use wolfxl_structural::{Axis, AxisShiftOp, SheetXmlInputs, ShiftPlan};
use zip::{write::SimpleFileOptions, ZipArchive, ZipWriter};

mod formulas;
pub mod ooxml_util;
pub mod rich_text;
pub mod sheet_patcher;
mod styles;
mod workbook;

use ooxml_util::attr_value;
use sheet_patcher::{CellPatch, CellValue};

/// A new cell value. Formulas never carry a stale cached result.
#[derive(Debug, Clone, PartialEq)]
pub enum EditValue {
    Number(f64),
    Text(String),
    Bool(bool),
    /// Formula text with or without the leading equals sign.
    Formula(String),
    /// Remove value and formula, preserving style.
    Clear,
}

/// One cell modification, using an A1 reference without a sheet prefix.
#[derive(Debug, Clone, PartialEq)]
pub struct CellEdit {
    pub cell: String,
    pub value: EditValue,
}

/// An editing error, with the offending sheet, reference, or package part.
#[derive(Debug)]
pub enum EditError {
    Io(std::io::Error),
    Zip(zip::result::ZipError),
    Xml(String),
    UnknownSheet(String),
    BadReference(String),
    Unsupported(String),
    Invalid(String),
}

impl std::fmt::Display for EditError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Io(e) => write!(f, "{e}"),
            Self::Zip(e) => write!(f, "{e}"),
            Self::Xml(e) => write!(f, "XML: {e}"),
            Self::UnknownSheet(s) => write!(f, "unknown sheet: {s}"),
            Self::BadReference(s) => write!(f, "invalid cell or range reference: {s}"),
            Self::Unsupported(s) => write!(f, "unsupported edit: {s}"),
            Self::Invalid(s) => f.write_str(s),
        }
    }
}
impl std::error::Error for EditError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Io(e) => Some(e),
            Self::Zip(e) => Some(e),
            _ => None,
        }
    }
}
impl From<std::io::Error> for EditError {
    fn from(e: std::io::Error) -> Self {
        Self::Io(e)
    }
}
impl From<zip::result::ZipError> for EditError {
    fn from(e: zip::result::ZipError) -> Self {
        Self::Zip(e)
    }
}
impl From<quick_xml::Error> for EditError {
    fn from(e: quick_xml::Error) -> Self {
        Self::Xml(e.to_string())
    }
}

/// A file-based workbook edit transaction. Failed edits leave the session unchanged.
pub struct EditSession {
    original: Vec<u8>,
    names: BTreeSet<String>,
    changed: BTreeMap<String, Vec<u8>>,
    removed: BTreeSet<String>,
    sheets: Vec<workbook::Sheet>,
    dirty: bool,
}

const WORKBOOK: &str = "xl/workbook.xml";
const WORKBOOK_RELS: &str = "xl/_rels/workbook.xml.rels";
const CONTENT_TYPES: &str = "[Content_Types].xml";

impl EditSession {
    /// Open an existing XLSX or XLSM package.
    pub fn open(path: &Path) -> Result<Self, EditError> {
        Self::from_bytes(std::fs::read(path)?)
    }

    /// Open an existing OOXML package held in memory.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self, EditError> {
        use wolfxl_rels::package_guard::{
            DEFAULT_MAX_ZIP_ENTRIES, DEFAULT_MAX_ZIP_ENTRY_BYTES, DEFAULT_MAX_ZIP_TOTAL_BYTES,
        };
        let mut archive = ZipArchive::new(Cursor::new(&bytes))?;
        if archive.len() > DEFAULT_MAX_ZIP_ENTRIES {
            return Err(EditError::Invalid("too many package parts".into()));
        }
        let mut names = BTreeSet::new();
        let mut total = 0u64;
        for i in 0..archive.len() {
            let entry = archive.by_index(i)?;
            let name = entry.name()?.into_owned();
            if name.starts_with('/')
                || name.contains('\\')
                || name.split('/').any(|p| p == ".." || p.contains(':'))
                || !names.insert(name.to_ascii_lowercase())
            {
                return Err(EditError::Invalid(format!(
                    "unsafe or duplicate package part: {name}"
                )));
            }
            total = total.saturating_add(entry.size());
            if entry.size() > DEFAULT_MAX_ZIP_ENTRY_BYTES || total > DEFAULT_MAX_ZIP_TOTAL_BYTES {
                return Err(EditError::Invalid(format!(
                    "package size limit exceeded: {name}"
                )));
            }
        }
        drop(archive);
        // Keep exact spelling for ZIP lookups and new-part collision checks.
        let mut archive = ZipArchive::new(Cursor::new(&bytes))?;
        let names = (0..archive.len())
            .map(|i| Ok(archive.by_index(i)?.name()?.into_owned()))
            .collect::<Result<_, EditError>>()?;
        drop(archive);
        let mut session = Self {
            original: bytes,
            names,
            changed: BTreeMap::new(),
            removed: BTreeSet::new(),
            sheets: Vec::new(),
            dirty: false,
        };
        session.sheets = workbook::sheets(&session.xml(WORKBOOK)?, &session.xml(WORKBOOK_RELS)?)?;
        session.xml(CONTENT_TYPES)?;
        for sheet in &session.sheets {
            if !session.names.contains(&sheet.path) {
                return Err(EditError::Invalid(format!(
                    "missing worksheet part: {}",
                    sheet.path
                )));
            }
        }
        Ok(session)
    }

    /// Sheet names in workbook order.
    pub fn sheet_names(&self) -> Vec<String> {
        self.sheets.iter().map(|s| s.name.clone()).collect()
    }

    /// Write values, formulas, or clears, retaining existing cell styles.
    pub fn write_cells(&mut self, sheet: &str, cells: &[CellEdit]) -> Result<(), EditError> {
        let path = self.sheet_path(sheet)?.to_owned();
        let mut patches = Vec::with_capacity(cells.len());
        let mut formula = false;
        for edit in cells {
            let (row, col) = cell_ref(&edit.cell)
                .map_err(|_| EditError::BadReference(format!("{sheet}!{}", edit.cell)))?;
            let value = match &edit.value {
                EditValue::Number(n) if n.is_finite() => CellValue::Number(*n),
                EditValue::Number(_) => {
                    return Err(EditError::Invalid(format!(
                        "non-finite number at {sheet}!{}",
                        edit.cell
                    )))
                }
                EditValue::Text(s) => {
                    valid_xml_text(s)?;
                    CellValue::String(s.clone())
                }
                EditValue::Bool(b) => CellValue::Boolean(*b),
                EditValue::Formula(s) => {
                    valid_xml_text(s)?;
                    let text = s.strip_prefix('=').unwrap_or(s);
                    if text.is_empty() {
                        return Err(EditError::Invalid(format!(
                            "empty formula at {sheet}!{}",
                            edit.cell
                        )));
                    }
                    formula = true;
                    CellValue::Formula(text.to_owned())
                }
                EditValue::Clear => CellValue::Blank,
            };
            patches.push(CellPatch {
                row,
                col,
                value: Some(value),
                style_index: None,
            });
        }
        if patches.is_empty() {
            return Ok(());
        }
        let new =
            sheet_patcher::patch_worksheet(&self.xml(&path)?, &patches).map_err(EditError::Xml)?;
        let mut changes = BTreeMap::from([(path, new.into_bytes())]);
        // Any edit can change a formula's precedent, so stale cached results
        // must be recalculated whenever the workbook holds formulas.
        let removed = if formula || self.has_formulas()? {
            self.recalculation(&mut changes)?
        } else {
            Vec::new()
        };
        self.apply(changes, removed);
        Ok(())
    }

    /// Insert rows at a one-based index.
    pub fn insert_rows(&mut self, sheet: &str, at: u32, count: u32) -> Result<(), EditError> {
        self.axis_shift(sheet, Axis::Row, at, count, false)
    }
    /// Delete rows at a one-based index.
    pub fn delete_rows(&mut self, sheet: &str, at: u32, count: u32) -> Result<(), EditError> {
        self.axis_shift(sheet, Axis::Row, at, count, true)
    }
    /// Insert columns at a one-based index.
    pub fn insert_columns(&mut self, sheet: &str, at: u32, count: u32) -> Result<(), EditError> {
        self.axis_shift(sheet, Axis::Col, at, count, false)
    }
    /// Delete columns at a one-based index.
    pub fn delete_columns(&mut self, sheet: &str, at: u32, count: u32) -> Result<(), EditError> {
        self.axis_shift(sheet, Axis::Col, at, count, true)
    }

    /// Append a blank worksheet or insert at a zero-based workbook position.
    pub fn add_sheet(&mut self, name: &str, position: Option<usize>) -> Result<(), EditError> {
        workbook::validate_name(name, &self.sheets, None)?;
        let position = position.unwrap_or(self.sheets.len());
        if position > self.sheets.len() {
            return Err(EditError::Invalid(format!(
                "sheet position out of range: {position}"
            )));
        }
        let path = (1u32..)
            .map(|n| format!("xl/worksheets/sheet{n}.xml"))
            .find(|p| !self.contains(p))
            .ok_or_else(|| EditError::Invalid("no free worksheet part name".into()))?;
        let app = self.optional_xml("docProps/app.xml")?;
        let (wb, rels, types, app) = workbook::add_sheet(
            &self.xml(WORKBOOK)?,
            &self.xml(WORKBOOK_RELS)?,
            &self.xml(CONTENT_TYPES)?,
            app.as_deref(),
            &self.sheets,
            name,
            position,
            &path,
        )?;
        let sheets = workbook::sheets(&wb, &rels)?;
        let mut changes = BTreeMap::from([(WORKBOOK.into(), wb.into_bytes()), (WORKBOOK_RELS.into(), rels.into_bytes()), (CONTENT_TYPES.into(), types.into_bytes()), (path, b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData/></worksheet>".to_vec())]);
        if let Some(app) = app {
            changes.insert("docProps/app.xml".into(), app.into_bytes());
        }
        self.apply(changes, Vec::new());
        self.sheets = sheets;
        Ok(())
    }

    /// Rename a worksheet and rewrite references throughout the workbook.
    pub fn rename_sheet(&mut self, old: &str, new: &str) -> Result<(), EditError> {
        self.sheet_path(old)?;
        workbook::validate_name(new, &self.sheets, Some(old))?;
        if old == new {
            return Ok(());
        }
        let app = self.optional_xml("docProps/app.xml")?;
        let (wb, app) = workbook::rename_sheet(&self.xml(WORKBOOK)?, app.as_deref(), old, new)?;
        let wb = formulas::rewrite_xml(&wb, |f| formulas::rename(f, old, new))?;
        let sheets = workbook::sheets(&wb, &self.xml(WORKBOOK_RELS)?)?;
        let mut changes = BTreeMap::from([(WORKBOOK.into(), wb.into_bytes())]);
        for path in self.formula_parts() {
            let xml = self.xml(&path)?;
            let changed = formulas::rewrite_xml(&xml, |f| formulas::rename(f, old, new))?;
            if changed != xml {
                changes.insert(path, changed.into_bytes());
            }
        }
        if let Some(app) = app {
            changes.insert("docProps/app.xml".into(), app.into_bytes());
        }
        let removed = self.recalculation(&mut changes)?;
        self.apply(changes, removed);
        self.sheets = sheets;
        Ok(())
    }

    /// Apply a number format, preserving each cell's other style components.
    pub fn set_number_format(
        &mut self,
        sheet: &str,
        range: &str,
        format_code: &str,
    ) -> Result<(), EditError> {
        let path = self.sheet_path(sheet)?.to_owned();
        let bounds =
            range_ref(range).map_err(|_| EditError::BadReference(format!("{sheet}!{range}")))?;
        valid_xml_text(format_code)?;
        if format_code.is_empty() {
            return Err(EditError::Invalid(format!(
                "empty number format for {sheet}!{range}"
            )));
        }
        let old_styles = self.optional_xml("xl/styles.xml")?;
        let (styles, xml) = styles::set_number_format(
            old_styles.as_deref(),
            &self.xml(&path)?,
            bounds,
            format_code,
        )?;
        let mut changes = BTreeMap::from([
            ("xl/styles.xml".into(), styles.into_bytes()),
            (path, xml.into_bytes()),
        ]);
        if old_styles.is_none() {
            self.add_styles_relationship(&mut changes)?;
        }
        self.apply(changes, Vec::new());
        Ok(())
    }

    /// Whether at least one edit has been applied.
    pub fn is_dirty(&self) -> bool {
        self.dirty
    }

    /// Serialize deterministically, copying untouched ZIP entries raw.
    pub fn to_bytes(&self) -> Result<Vec<u8>, EditError> {
        if !self.dirty {
            return Ok(self.original.clone());
        }
        let mut source = ZipArchive::new(Cursor::new(&self.original))?;
        let mut output = ZipWriter::new(Cursor::new(Vec::new()));
        let mut copied = BTreeSet::new();
        for index in 0..source.len() {
            let entry = source.by_index(index)?;
            let name = entry.name()?.into_owned();
            if self.removed.contains(&name) {
                continue;
            }
            if let Some(bytes) = self.changed.get(&name) {
                output.start_file(
                    &name,
                    SimpleFileOptions::default()
                        .compression_method(entry.compression())
                        .last_modified_time(entry.last_modified().unwrap_or_default()),
                )?;
                output.write_all(bytes)?;
            } else {
                output.raw_copy_file(entry)?;
            }
            copied.insert(name);
        }
        for (name, bytes) in &self.changed {
            if !copied.contains(name) && !self.removed.contains(name) {
                output.start_file(name, SimpleFileOptions::default())?;
                output.write_all(bytes)?;
            }
        }
        Ok(output.finish()?.into_inner())
    }

    /// Write and fsync a sibling temporary file, then atomically replace the destination.
    /// An existing destination keeps its permissions; a new one gets the usual
    /// mode for a created file rather than the temporary file's owner-only mode.
    pub fn save(&self, path: &Path) -> Result<(), EditError> {
        let parent = path
            .parent()
            .filter(|p| !p.as_os_str().is_empty())
            .unwrap_or_else(|| Path::new("."));
        let existing = std::fs::metadata(path).ok().map(|m| m.permissions());
        let mut builder = tempfile::Builder::new();
        #[cfg(unix)]
        {
            use std::os::unix::fs::PermissionsExt;
            builder.permissions(std::fs::Permissions::from_mode(0o666));
        }
        let mut temporary = builder.tempfile_in(parent)?;
        if let Some(permissions) = existing {
            temporary.as_file().set_permissions(permissions)?;
        }
        temporary.write_all(&self.to_bytes()?)?;
        temporary.as_file().sync_all()?;
        temporary
            .persist(path)
            .map_err(|e| EditError::Io(e.error))?;
        #[cfg(unix)]
        std::fs::File::open(parent)?.sync_all()?;
        Ok(())
    }

    fn sheet_path(&self, name: &str) -> Result<&str, EditError> {
        self.sheets
            .iter()
            .find(|s| s.name == name)
            .map(|s| s.path.as_str())
            .ok_or_else(|| EditError::UnknownSheet(name.into()))
    }
    fn contains(&self, path: &str) -> bool {
        !self.removed.contains(path)
            && (self.names.contains(path) || self.changed.contains_key(path))
    }
    fn part(&self, path: &str) -> Result<Option<Vec<u8>>, EditError> {
        if !self.contains(path) {
            return Ok(None);
        }
        if let Some(bytes) = self.changed.get(path) {
            return Ok(Some(bytes.clone()));
        }
        let mut source = ZipArchive::new(Cursor::new(&self.original))?;
        let mut bytes = Vec::new();
        source.by_name(path)?.read_to_end(&mut bytes)?;
        Ok(Some(bytes))
    }
    fn optional_xml(&self, path: &str) -> Result<Option<String>, EditError> {
        self.part(path)?
            .map(|b| String::from_utf8(b).map_err(|e| EditError::Xml(format!("{path}: {e}"))))
            .transpose()
    }
    fn xml(&self, path: &str) -> Result<String, EditError> {
        self.optional_xml(path)?
            .ok_or_else(|| EditError::Invalid(format!("missing package part: {path}")))
    }
    fn apply(&mut self, changes: BTreeMap<String, Vec<u8>>, removed: Vec<String>) {
        for (path, bytes) in changes {
            self.removed.remove(&path);
            self.changed.insert(path, bytes);
        }
        for path in removed {
            self.changed.remove(&path);
            self.removed.insert(path);
        }
        self.dirty = true;
    }
    fn recalculation(
        &self,
        changes: &mut BTreeMap<String, Vec<u8>>,
    ) -> Result<Vec<String>, EditError> {
        let current = changes
            .get(WORKBOOK)
            .map(|b| {
                std::str::from_utf8(b)
                    .map(str::to_owned)
                    .map_err(|e| EditError::Xml(e.to_string()))
            })
            .transpose()?
            .unwrap_or(self.xml(WORKBOOK)?);
        let (wb, rels, types, mut removed) = workbook::recalculate(
            &current,
            &self.xml(WORKBOOK_RELS)?,
            &self.xml(CONTENT_TYPES)?,
        )?;
        changes.insert(WORKBOOK.into(), wb.into_bytes());
        let old_rels = self.xml(WORKBOOK_RELS)?;
        if rels != old_rels {
            changes.insert(WORKBOOK_RELS.into(), rels.into_bytes());
        }
        let old_types = self.xml(CONTENT_TYPES)?;
        if types != old_types {
            changes.insert(CONTENT_TYPES.into(), types.into_bytes());
        }
        if self.contains("xl/calcChain.xml") {
            removed.push("xl/calcChain.xml".into());
        }
        Ok(removed)
    }
    /// Whether the package holds any cell formula whose cached result an edit could make stale.
    fn has_formulas(&self) -> Result<bool, EditError> {
        if self.contains("xl/calcChain.xml") {
            return Ok(true);
        }
        for sheet in &self.sheets {
            if let Some(bytes) = self.part(&sheet.path)? {
                if has_formula_element(&bytes)? {
                    return Ok(true);
                }
            }
        }
        Ok(false)
    }
    fn formula_parts(&self) -> Vec<String> {
        self.names
            .iter()
            .chain(self.changed.keys())
            .filter(|p| {
                p.ends_with(".xml")
                    && (p.starts_with("xl/worksheets/")
                        || p.starts_with("xl/charts/")
                        || p.starts_with("xl/tables/"))
            })
            .cloned()
            .collect::<BTreeSet<_>>()
            .into_iter()
            .collect()
    }
    fn add_styles_relationship(
        &self,
        changes: &mut BTreeMap<String, Vec<u8>>,
    ) -> Result<(), EditError> {
        let mut rels = RelsGraph::parse(self.xml(WORKBOOK_RELS)?.as_bytes())
            .map_err(|e| EditError::Xml(e.to_string()))?;
        rels.add(wolfxl_rels::rt::STYLES, "styles.xml", TargetMode::Internal);
        changes.insert(WORKBOOK_RELS.into(), rels.serialize());
        let types = self.xml(CONTENT_TYPES)?;
        let at = types
            .rfind("</")
            .ok_or_else(|| EditError::Xml("missing content-types root".into()))?;
        let mut new = types.clone();
        new.insert_str(at, "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
        changes.insert(CONTENT_TYPES.into(), new.into_bytes());
        Ok(())
    }

    fn axis_shift(
        &mut self,
        sheet: &str,
        axis: Axis,
        at: u32,
        count: u32,
        delete: bool,
    ) -> Result<(), EditError> {
        let path = self.sheet_path(sheet)?.to_owned();
        let max = if axis == Axis::Row {
            wolfxl_formula::MAX_ROW
        } else {
            wolfxl_formula::MAX_COL
        };
        if at == 0
            || at > max
            || count > max
            || (count > 0 && at.checked_add(count - 1).is_none_or(|end| end > max))
        {
            return Err(EditError::BadReference(format!(
                "{sheet}: {axis:?} {at}, count {count}"
            )));
        }
        if count == 0 {
            return Ok(());
        }
        let n = if delete {
            -(count as i32)
        } else {
            count as i32
        };
        let plan = ShiftPlan { axis, idx: at, n };
        let xml = self.xml(&path)?;
        if !delete {
            check_shift_bounds(&xml, axis, at, count, max)?;
        }
        let mut prefix = "_wolfxl_preserved_reference_".to_owned();
        while xml.contains(&prefix) {
            prefix.push('_');
        }
        let (protected, references) = formulas::protect_foreign(&xml, sheet, &prefix)?;
        let wb = self.xml(WORKBOOK)?;
        // The orchestrator handles cell coordinates, merged ranges, dimensions,
        // validations, tables, drawings, comments, and control anchors.
        let mut owned = BTreeMap::new();
        let mut relations = Vec::new();
        let relpath = relationship_part(&path);
        if let Some(rels) = self.optional_xml(&relpath)? {
            let graph =
                RelsGraph::parse(rels.as_bytes()).map_err(|e| EditError::Xml(e.to_string()))?;
            for rel in graph.iter() {
                if rel.mode == TargetMode::Internal {
                    let target = resolve_target(&path, &rel.target)?;
                    if let Some(bytes) = self.part(&target)? {
                        owned.insert(target.clone(), bytes);
                        relations.push((rel.rel_type.clone(), target));
                    }
                }
            }
        }
        let mut table_formulas = BTreeMap::new();
        for (kind, target) in &relations {
            if kind == wolfxl_rels::rt::TABLE {
                let xml = String::from_utf8(owned[target].clone())
                    .map_err(|e| EditError::Xml(e.to_string()))?;
                let mut table_prefix = "_wolfxl_preserved_table_formula_".to_owned();
                while xml.contains(&table_prefix) {
                    table_prefix.push('_');
                }
                let mut formulas = Vec::new();
                let protected = formulas::rewrite_xml(&xml, |formula| {
                    let index = formulas.len();
                    formulas.push(formula.to_owned());
                    Ok(format!("{table_prefix}{index}_"))
                })?;
                owned.insert(target.clone(), protected.into_bytes());
                table_formulas.insert(target.clone(), (table_prefix, formulas));
            }
        }
        let mut inputs = SheetXmlInputs::empty();
        inputs.sheets.insert(sheet.into(), protected.as_bytes());
        inputs.sheet_paths.insert(sheet.into(), path.clone());
        inputs.workbook_xml = Some(wb.as_bytes());
        inputs.sheet_positions = self
            .sheets
            .iter()
            .enumerate()
            .map(|(i, s)| (s.name.clone(), i as u32))
            .collect();
        for (kind, target) in &relations {
            let bytes = owned[target].as_slice();
            match kind.as_str() {
                wolfxl_rels::rt::TABLE => {
                    inputs
                        .tables
                        .entry(sheet.into())
                        .or_default()
                        .push((target.clone(), bytes));
                }
                wolfxl_rels::rt::COMMENTS => {
                    inputs
                        .comments
                        .insert(sheet.into(), (target.clone(), bytes));
                }
                wolfxl_rels::rt::VML_DRAWING => {
                    inputs.vml.insert(sheet.into(), (target.clone(), bytes));
                }
                wolfxl_rels::rt::DRAWING => {
                    inputs
                        .drawings
                        .insert(sheet.into(), (target.clone(), bytes));
                }
                wolfxl_rels::rt::CTRL_PROP => {
                    inputs
                        .control_props
                        .entry(sheet.into())
                        .or_default()
                        .push((target.clone(), bytes));
                }
                _ => {}
            }
        }
        let mut changes = wolfxl_structural::apply_workbook_shift(
            inputs,
            &[AxisShiftOp {
                sheet: sheet.into(),
                axis,
                idx: at,
                n,
            }],
        )
        .file_patches;
        if let Some(bytes) = changes.get_mut(&path) {
            let xml = String::from_utf8(std::mem::take(bytes))
                .map_err(|e| EditError::Xml(e.to_string()))?;
            *bytes = formulas::unprotect(xml, &prefix, &references).into_bytes();
        }
        // Rewrite formulas on all other sheets without moving their coordinates.
        for other in &self.sheets {
            if other.name == sheet {
                continue;
            }
            let xml = self.xml(&other.path)?;
            let updated =
                formulas::rewrite_xml(&xml, |f| formulas::shift(f, &other.name, sheet, &plan))?;
            if updated != xml {
                changes.insert(other.path.clone(), updated.into_bytes());
            }
        }
        // Charts and tables can hold references outside worksheet XML.
        for path in self
            .formula_parts()
            .into_iter()
            .filter(|p| p.starts_with("xl/charts/") || p.starts_with("xl/tables/"))
        {
            let original = self.xml(&path)?;
            let mut xml = changes
                .get(&path)
                .map(|b| String::from_utf8(b.clone()).map_err(|e| EditError::Xml(e.to_string())))
                .transpose()?
                .unwrap_or_else(|| original.clone());
            let source = if let Some((prefix, references)) = table_formulas.get(&path) {
                xml = formulas::unprotect(xml, prefix, references);
                sheet
            } else {
                ""
            };
            let updated =
                formulas::rewrite_xml(&xml, |f| formulas::shift(f, source, sheet, &plan))?;
            if updated != original {
                changes.insert(path, updated.into_bytes());
            } else {
                changes.remove(&path);
            }
        }
        // Rebuild defined-name rewrites with explicit per-name scope.
        let updated = formulas::shift_defined_names(&wb, &self.sheet_names(), sheet, &plan)?;
        changes.remove(WORKBOOK);
        if updated != wb {
            changes.insert(WORKBOOK.into(), updated.into_bytes());
        }
        let removed = self.recalculation(&mut changes)?;
        self.apply(changes, removed);
        Ok(())
    }
}

fn valid_xml_text(text: &str) -> Result<(), EditError> {
    if text.chars().any(|c| {
        !(matches!(c, '\t' | '\n' | '\r') || c >= ' ' && c != '\u{fffe}' && c != '\u{ffff}')
    }) {
        Err(EditError::Invalid(
            "text contains a character forbidden by XML".into(),
        ))
    } else {
        Ok(())
    }
}

/// Whether worksheet XML contains a cell formula element, including shared-formula children.
fn has_formula_element(xml: &[u8]) -> Result<bool, EditError> {
    let mut reader = quick_xml::Reader::from_reader(xml);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref().as_bytes() == b"f" => {
                return Ok(true)
            }
            Event::Eof => return Ok(false),
            _ => {}
        }
        buf.clear();
    }
}

fn cell_ref(reference: &str) -> Result<(u32, u32), EditError> {
    if reference.is_empty() || reference.contains(['!', ':', '$']) {
        return Err(EditError::BadReference(reference.into()));
    }
    let cell = wolfxl_formula::parse_cell(reference)
        .ok_or_else(|| EditError::BadReference(reference.into()))?;
    Ok((cell.row, cell.col))
}
fn range_ref(reference: &str) -> Result<(u32, u32, u32, u32), EditError> {
    let mut parts = reference.split(':');
    let (r1, c1) = cell_ref(parts.next().unwrap_or_default())?;
    let (r2, c2) = parts.next().map(cell_ref).transpose()?.unwrap_or((r1, c1));
    if parts.next().is_some() || r2 < r1 || c2 < c1 {
        return Err(EditError::BadReference(reference.into()));
    }
    Ok((r1, c1, r2, c2))
}
fn relationship_part(path: &str) -> String {
    let (dir, name) = path.rsplit_once('/').unwrap_or(("", path));
    format!("{dir}/_rels/{name}.rels")
}
fn resolve_target(source: &str, target: &str) -> Result<String, EditError> {
    let joined = if target.starts_with('/') {
        target.trim_start_matches('/').to_owned()
    } else {
        format!(
            "{}/{target}",
            source.rsplit_once('/').map_or("", |(dir, _)| dir)
        )
    };
    let mut parts = Vec::new();
    for part in joined.split('/') {
        match part {
            "" | "." => {}
            ".." => {
                if parts.pop().is_none() {
                    return Err(EditError::Invalid(format!(
                        "unsafe relationship target: {target}"
                    )));
                }
            }
            _ => parts.push(part),
        }
    }
    Ok(parts.join("/"))
}
fn check_shift_bounds(
    xml: &str,
    axis: Axis,
    at: u32,
    count: u32,
    max: u32,
) -> Result<(), EditError> {
    let mut reader = quick_xml::Reader::from_str(xml);
    loop {
        match reader.read_event()? {
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref().as_bytes() == b"c" => {
                if let Some(reference) = attr_value(&e, b"r") {
                    let (row, col) = cell_ref(&reference)?;
                    let index = if axis == Axis::Row { row } else { col };
                    if index >= at && index + count > max {
                        return Err(EditError::Invalid(format!(
                            "insertion would move cell {reference} beyond Excel's worksheet bounds"
                        )));
                    }
                }
            }
            Event::Eof => return Ok(()),
            _ => {}
        }
    }
}
