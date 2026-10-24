//! wolfxl-core: pure-Rust spreadsheet reader with Excel number-format-aware rendering.
//!
//! Read XLSX, XLS, XLSB and ODS with the independently authored MIT native
//! readers. Use the built-in delimited reader for CSV. The Python-independent
//! `Workbook → Sheet → Cell` API includes common Excel number-format rendering
//! for currency, percentage, scientific, date and time values.
//!
//! ```no_run
//! use wolfxl_core::Workbook;
//!
//! let mut wb = Workbook::open("examples/sample-financials.xlsx")?;
//! let sheet = wb.first_sheet()?;
//! let (rows, cols) = sheet.dimensions();
//! println!("{} rows × {} columns", rows, cols);
//! # Ok::<_, wolfxl_core::Error>(())
//! ```
//!
//! ## Scope
//!
//! - **In scope today:** native spreadsheet values, best-effort number formats,
//!   [`FormatCategory`], [`format_cell`], workbook maps and column summaries.
//!   Retain sparse cell origins, date epochs, scoped-name entries, XLSX table
//!   inventory and canonical error literals. For OOXML formats, use the
//!   `xl/styles.xml` fallback when reader metadata is unavailable. For CSV,
//!   infer numeric-looking strings while retaining their original values.
//! - **Not yet:** write side.
//!
//! Python support remains edition-specific: Community XLS values are supported;
//! Community Python ODS, binary-format editing, streaming and styles stay refused.

pub mod cell;
mod csv_reader;
pub mod error;
pub mod format;
pub mod map;
pub mod ooxml;
pub mod schema;
pub mod sheet;
pub mod styles;
pub mod workbook;
pub mod worksheet_xml;

pub use cell::{Cell, CellValue};
pub use error::{Error, Result};
pub use format::{classify_format, format_cell, FormatCategory};
pub use map::{classify_sheet, SheetClass, SheetMap, WorkbookMap};
pub use schema::{infer_sheet_schema, Cardinality, ColumnSchema, InferredType, SheetSchema};
pub use sheet::Sheet;
pub use styles::{builtin_num_fmt, resolve_num_fmt, XfEntry, BUILTIN_NUM_FMTS};
pub use workbook::{SourceFormat, Workbook, WorkbookStyles};
