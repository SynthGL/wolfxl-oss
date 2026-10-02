//! Top-level workbook data: sheets, doc props, styles, SST, defined names.

use super::comment::CommentAuthorTable;
use super::defined_name::DefinedName;
use super::format::StylesBuilder;
use super::threaded_comment::PersonTable;
use super::worksheet::Worksheet;
use crate::intern::SstBuilder;
use crate::parse::workbook_security::WorkbookSecurity;

/// A workbook awaiting serialization.
///
/// Build one up by calling [`Workbook::add_sheet`] (and the various property
/// setters), then hand it to the emitter (Wave 4 — the emitter entrypoint
/// arrives with the `NativeWorkbook` pyclass).
#[derive(Debug, Default)]
pub struct Workbook {
    /// Sheets in display order. The first sheet is the default visible tab.
    pub sheets: Vec<Worksheet>,

    /// Workbook-level defined names (formula aliases, print areas at
    /// the workbook scope).
    pub defined_names: Vec<DefinedName>,

    /// Document properties surfaced in `docProps/core.xml` and `app.xml`.
    pub doc_props: DocProperties,

    /// The shared styles table. Styles are deduped on insertion.
    pub styles: StylesBuilder,

    /// The shared string table. Strings are interned on insertion.
    pub sst: SstBuilder,

    /// Workbook-scope comment authors. Every `commentsN.xml` emits this
    /// full `<authors>` block — `authorId` on a `<comment>` indexes into
    /// it. Insertion order is preserved by [`CommentAuthorTable`] so
    /// multi-author workbooks round-trip without the BTreeMap reordering
    /// bug that motivated this rewrite.
    pub comment_authors: CommentAuthorTable,

    /// Workbook-level security blocks (`<workbookProtection>` and
    /// `<fileSharing>`). Both are optional; when both are `None` the
    /// emitter writes neither element.
    pub security: WorkbookSecurity,

    /// Use Excel's 1904 date epoch for date serials and `<workbookPr>`.
    pub date1904: bool,

    /// Workbook-scope threaded-comment person registry (RFC-068 / G08).
    /// Empty by default — `xl/persons/personList.xml` is only emitted when
    /// at least one threaded comment exists.
    pub persons: PersonTable,
}

impl Workbook {
    pub fn new() -> Self {
        Self::default()
    }

    /// Add a sheet and return its index.
    pub fn add_sheet(&mut self, sheet: Worksheet) -> usize {
        let idx = self.sheets.len();
        self.sheets.push(sheet);
        idx
    }

    pub fn sheet_mut(&mut self, idx: usize) -> Option<&mut Worksheet> {
        self.sheets.get_mut(idx)
    }

    pub fn sheet_by_name(&self, name: &str) -> Option<&Worksheet> {
        self.sheets.iter().find(|s| s.name == name)
    }

    /// Look up a sheet by name and return a mutable reference. Used by the
    /// pyclass when callers identify sheets by name rather than by index.
    pub fn sheet_mut_by_name(&mut self, name: &str) -> Option<&mut Worksheet> {
        self.sheets.iter_mut().find(|s| s.name == name)
    }

    /// Look up a sheet by name and return its position in `self.sheets`.
    ///
    /// Used by `NativeWorkbook::add_named_range` to translate the oracle's
    /// `scope="sheet"` + sheet name into the native
    /// [`DefinedName::scope_sheet_index`]. Returns `None` when no sheet
    /// matches — callers surface this to Python as a `ValueError`.
    pub fn sheet_index_by_name(&self, name: &str) -> Option<usize> {
        self.sheets.iter().position(|s| s.name == name)
    }

    /// Rename a sheet by its current name. Errors when no sheet matches
    /// `old`, when the new name fails Excel validation, or when the new
    /// name would collide with another existing sheet.
    pub fn rename_sheet(&mut self, old: &str, new: String) -> Result<(), String> {
        if old == new {
            return Ok(());
        }
        if self.sheets.iter().any(|s| s.name == new) {
            return Err(format!("sheet name {new:?} already exists in workbook"));
        }
        let sheet = self
            .sheet_mut_by_name(old)
            .ok_or_else(|| format!("no sheet named {old:?}"))?;
        sheet.rename(new)
    }

    /// Move a sheet by an openpyxl-style offset, clamped to workbook bounds.
    ///
    /// Sheet-scoped defined names store their owner as the sheet's numeric
    /// position, so keep those indexes attached to the same logical sheet.
    pub fn move_sheet(&mut self, name: &str, offset: isize) -> Result<(), String> {
        let idx = self
            .sheet_index_by_name(name)
            .ok_or_else(|| format!("no sheet named {name:?}"))?;
        if self.sheets.len() <= 1 || offset == 0 {
            return Ok(());
        }

        let max_idx = self.sheets.len() - 1;
        let shifted = idx as isize + offset;
        let new_pos = shifted.clamp(0, max_idx as isize) as usize;
        if new_pos == idx {
            return Ok(());
        }

        let sheet = self.sheets.remove(idx);
        self.sheets.insert(new_pos, sheet);

        for defined_name in &mut self.defined_names {
            if let Some(scope_idx) = defined_name.scope_sheet_index {
                defined_name.scope_sheet_index = Some(match scope_idx {
                    s if s == idx => new_pos,
                    s if idx < new_pos && s > idx && s <= new_pos => s - 1,
                    s if new_pos < idx && s >= new_pos && s < idx => s + 1,
                    s => s,
                });
            }
        }
        Ok(())
    }

    /// Remove a sheet by name.
    ///
    /// Names scoped to the removed sheet go with it, and names scoped to
    /// later sheets shift down so `localSheetId` keeps pointing at the same
    /// logical sheet.
    pub fn remove_sheet(&mut self, name: &str) -> Result<(), String> {
        let idx = self
            .sheet_index_by_name(name)
            .ok_or_else(|| format!("no sheet named {name:?}"))?;
        self.sheets.remove(idx);
        self.defined_names
            .retain(|defined_name| defined_name.scope_sheet_index != Some(idx));
        for defined_name in &mut self.defined_names {
            if let Some(scope_idx) = defined_name.scope_sheet_index {
                if scope_idx > idx {
                    defined_name.scope_sheet_index = Some(scope_idx - 1);
                }
            }
        }
        Ok(())
    }

    /// Replace the workbook-level document properties block.
    pub fn set_doc_props(&mut self, props: DocProperties) {
        self.doc_props = props;
    }
}

/// Document properties surfaced in the two docProps parts.
///
/// Excel shows these in File → Info. They're optional — if everything
/// is `None`, the emitter writes minimal stub parts so the container
/// stays valid.
#[derive(Debug, Clone, Default)]
pub struct DocProperties {
    pub title: Option<String>,
    pub subject: Option<String>,
    pub creator: Option<String>,
    pub keywords: Option<String>,
    pub description: Option<String>,
    pub last_modified_by: Option<String>,
    pub category: Option<String>,
    /// OOXML `<cp:contentStatus>` — free-form workflow tag (e.g. "Draft",
    /// "Final", "Reviewed"). Mirrors oracle's `set_status()` and the
    /// `contentStatus` key Python's `set_properties` accepts.
    pub content_status: Option<String>,

    /// If `None`, the emitter uses the current wall-clock time (or the
    /// `WOLFXL_TEST_EPOCH` override for deterministic output).
    pub created: Option<chrono::NaiveDateTime>,
    pub modified: Option<chrono::NaiveDateTime>,

    /// Company / application metadata shown in `docProps/app.xml`.
    pub company: Option<String>,
    pub manager: Option<String>,
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::worksheet::Worksheet;

    fn wb_with(sheet_names: &[&str]) -> Workbook {
        let mut wb = Workbook::new();
        for name in sheet_names {
            wb.add_sheet(Worksheet::new(*name));
        }
        wb
    }

    #[test]
    fn rename_sheet_valid_updates_target_only() {
        let mut wb = wb_with(&["Data", "Summary"]);
        assert!(wb.rename_sheet("Data", "Inputs".to_string()).is_ok());
        assert_eq!(wb.sheets[0].name, "Inputs");
        assert_eq!(wb.sheets[1].name, "Summary");
    }

    #[test]
    fn rename_sheet_missing_old_errors() {
        let mut wb = wb_with(&["Data"]);
        let err = wb.rename_sheet("Nope", "Whatever".to_string()).unwrap_err();
        assert!(
            err.contains("Nope"),
            "msg should reference missing name: {err}"
        );
        assert_eq!(wb.sheets[0].name, "Data", "no sheet may change on Err");
    }

    #[test]
    fn rename_sheet_collision_errors() {
        let mut wb = wb_with(&["Data", "Summary"]);
        let err = wb.rename_sheet("Data", "Summary".to_string()).unwrap_err();
        assert!(err.contains("already exists"), "{err}");
        assert_eq!(wb.sheets[0].name, "Data");
        assert_eq!(wb.sheets[1].name, "Summary");
    }

    #[test]
    fn rename_sheet_same_name_is_noop() {
        let mut wb = wb_with(&["Data"]);
        assert!(wb.rename_sheet("Data", "Data".to_string()).is_ok());
        assert_eq!(wb.sheets[0].name, "Data");
    }

    #[test]
    fn rename_sheet_propagates_validation_error() {
        let mut wb = wb_with(&["Data"]);
        // 32 chars — Worksheet::rename must reject this.
        let too_long = "x".repeat(32);
        assert!(wb.rename_sheet("Data", too_long).is_err());
        assert_eq!(wb.sheets[0].name, "Data");
    }

    #[test]
    fn move_sheet_positive_offset_reorders_and_remaps_scoped_names() {
        let mut wb = wb_with(&["A", "B", "C", "D"]);
        wb.defined_names.push(DefinedName {
            name: "OnA".to_string(),
            formula: "A!$A$1".to_string(),
            scope_sheet_index: Some(0),
            builtin: None,
            hidden: false,
            ..Default::default()
        });
        wb.defined_names.push(DefinedName {
            name: "OnC".to_string(),
            formula: "C!$A$1".to_string(),
            scope_sheet_index: Some(2),
            builtin: None,
            hidden: false,
            ..Default::default()
        });

        wb.move_sheet("A", 2).unwrap();

        let names: Vec<_> = wb.sheets.iter().map(|s| s.name.as_str()).collect();
        assert_eq!(names, ["B", "C", "A", "D"]);
        assert_eq!(wb.defined_names[0].scope_sheet_index, Some(2));
        assert_eq!(wb.defined_names[1].scope_sheet_index, Some(1));
    }

    #[test]
    fn move_sheet_negative_offset_reorders_and_clamps() {
        let mut wb = wb_with(&["A", "B", "C", "D"]);
        wb.move_sheet("D", -99).unwrap();

        let names: Vec<_> = wb.sheets.iter().map(|s| s.name.as_str()).collect();
        assert_eq!(names, ["D", "A", "B", "C"]);
    }

    #[test]
    fn move_sheet_missing_name_errors_without_mutating() {
        let mut wb = wb_with(&["A", "B"]);
        let err = wb.move_sheet("Nope", 1).unwrap_err();
        assert!(err.contains("Nope"), "{err}");
        let names: Vec<_> = wb.sheets.iter().map(|s| s.name.as_str()).collect();
        assert_eq!(names, ["A", "B"]);
    }

    #[test]
    fn remove_sheet_drops_owned_names_and_shifts_later_scopes() {
        let mut wb = wb_with(&["A", "B", "C"]);
        for (name, scope) in [
            ("OnA", Some(0)),
            ("OnB", Some(1)),
            ("OnC", Some(2)),
            ("Global", None),
        ] {
            wb.defined_names.push(DefinedName {
                name: name.to_string(),
                formula: "$A$1".to_string(),
                scope_sheet_index: scope,
                ..Default::default()
            });
        }

        wb.remove_sheet("B").unwrap();

        let names: Vec<_> = wb.sheets.iter().map(|s| s.name.as_str()).collect();
        assert_eq!(names, ["A", "C"]);
        let scopes: Vec<_> = wb
            .defined_names
            .iter()
            .map(|dn| (dn.name.as_str(), dn.scope_sheet_index))
            .collect();
        assert_eq!(
            scopes,
            [("OnA", Some(0)), ("OnC", Some(1)), ("Global", None)]
        );
    }

    #[test]
    fn remove_sheet_missing_name_errors_without_mutating() {
        let mut wb = wb_with(&["A"]);
        let err = wb.remove_sheet("Nope").unwrap_err();
        assert!(err.contains("Nope"), "{err}");
        assert_eq!(wb.sheets.len(), 1);
    }

    #[test]
    fn sheet_mut_by_name_finds_and_misses() {
        let mut wb = wb_with(&["Data", "Summary"]);
        assert!(wb.sheet_mut_by_name("Summary").is_some());
        assert!(wb.sheet_mut_by_name("Nope").is_none());
        // Mutability check — we can write through the returned ref.
        wb.sheet_mut_by_name("Data")
            .unwrap()
            .set_column_width(1, 25.0);
        assert_eq!(wb.sheets[0].columns[&1].width, Some(25.0));
    }

    #[test]
    fn sheet_index_by_name_returns_position() {
        let wb = wb_with(&["Data", "Summary", "Notes"]);
        assert_eq!(wb.sheet_index_by_name("Data"), Some(0));
        assert_eq!(wb.sheet_index_by_name("Summary"), Some(1));
        assert_eq!(wb.sheet_index_by_name("Notes"), Some(2));
    }

    #[test]
    fn sheet_index_by_name_missing_returns_none() {
        let wb = wb_with(&["Data"]);
        assert_eq!(wb.sheet_index_by_name("Nope"), None);
        assert_eq!(wb.sheet_index_by_name(""), None);
    }

    #[test]
    fn set_doc_props_replaces_in_full() {
        let mut wb = Workbook::new();
        wb.set_doc_props(DocProperties {
            title: Some("Q4 Report".to_string()),
            creator: Some("Wolfgang".to_string()),
            ..Default::default()
        });
        assert_eq!(wb.doc_props.title.as_deref(), Some("Q4 Report"));
        assert_eq!(wb.doc_props.creator.as_deref(), Some("Wolfgang"));
        // Replacement (not merge) — old fields go away.
        wb.set_doc_props(DocProperties {
            subject: Some("New".to_string()),
            ..Default::default()
        });
        assert_eq!(wb.doc_props.title, None);
        assert_eq!(wb.doc_props.subject.as_deref(), Some("New"));
    }
}
