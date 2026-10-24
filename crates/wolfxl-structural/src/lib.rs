//! `wolfxl-structural` — pure-Rust XML rewrites for `Worksheet.insert_rows`
//! / `delete_rows` (RFC-030) and `insert_cols` / `delete_cols` (RFC-031).
//!
//! This crate is consumed by the wolfxl patcher (modify-mode). Per the
//! workspace invariant from `CLAUDE.md`, it has **no PyO3 dependency** —
//! pure Rust only. Heavy lifting (formula re-emission) is delegated to
//! `wolfxl-formula`.
//!
//! See `Plans/rfcs/030-insert-delete-rows.md` for the authoritative spec.
//!
//! # Modules
//!
//! - [`axis`] — `Axis::{Row, Col}` enum + the unified [`ShiftPlan`] type.
//! - [`shift_cells`] — rewrite `<sheetData>` cell coordinates +
//!   `<row r="">` / `<dimension ref="">`.
//! - [`shift_anchors`] — rewrite `ref` / `sqref` attribute strings (single
//!   cell, range, multi-range space-separated).
//! - [`shift_formulas`] — thin wrapper over `wolfxl_formula::shift`.
//! - [`shift_workbook`] — top-level orchestrator that walks workbook.xml
//!   + every sheet.xml + every tableN.xml + every commentsN.xml + every
//!   vmlDrawingN.vml part touched.
//!
//! # `respect_dollar` decision (RFC-030 §10)
//!
//! The `respect_dollar` flag on `wolfxl_formula::shift` is required-no-
//! default until the BLOCKER doc at
//! `Plans/rfcs/notes/excel-respect-dollar-check.md` is resolved.
//! RFC-030 picks `respect_dollar=false` (insert/delete is a coordinate-
//! space remap; `$A$1` shifts) — see [`shift_formulas`] for the call
//! site and the TODO comment.

#![deny(rust_2018_idioms)]

pub mod axis;
pub mod cols;
pub(crate) mod control_props_shift;
pub(crate) mod drawing_shift;
pub mod range_move;
pub mod sheet_copy;
pub mod shift_anchors;
pub mod shift_cells;
pub mod shift_formulas;
pub mod shift_workbook;
pub(crate) mod table_shift;
pub(crate) mod vml_shift;
pub(crate) mod xml_text;

#[cfg(test)]
mod tests;

pub use axis::{Axis, ShiftPlan};
pub use cols::{shift_sheet_cols_block, split_col_spans, ColSpan};
pub use range_move::{apply_range_move, RangeMovePlan};
pub use sheet_copy::{
    plan_sheet_copy, DefinedNameClone, SheetCopyError, SheetCopyInputs, SheetCopyMutations,
};
pub use shift_anchors::{shift_anchor, shift_sqref};
pub use shift_cells::shift_sheet_cells;
pub use shift_formulas::{shift_formula, shift_formula_with_meta};
pub use shift_workbook::{apply_workbook_shift, AxisShiftOp, SheetXmlInputs, WorkbookMutations};

/// Convenience re-export so callers don't have to depend on
/// `wolfxl-formula` directly just to construct a `Range`.
pub use wolfxl_formula::{TranslateResult, MAX_COL, MAX_ROW};
