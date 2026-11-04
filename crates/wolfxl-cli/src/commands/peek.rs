//! `wolfxl peek FILE|-`: preview one sheet.
//!
//! `--where COLUMN<op>VALUE` (repeatable, AND) keeps the header row plus
//! the data rows every
//! condition accepts; `--select A,B` keeps the named columns in the order
//! given. Both run before rendering, so `--max-rows` and every export
//! format see the reshaped sheet. Without them the sheet renders exactly as
//! loaded.
//!
//! `--interactive` browses the reshaped sheet in a terminal grid instead of
//! printing it (see `crate::tui`).
//!
//! Exit codes: 0 success, 1 input or rendering failure, 2 a malformed or
//! unresolvable `--where`/`--select`, standard input without `--from`, or
//! `--interactive` without a terminal on standard output or with standard
//! input as the data source.

use std::path::Path;
use std::process::ExitCode;

use anyhow::{anyhow, bail, Context, Result};
use wolfxl_core::{Cell, CellValue, Sheet};

use crate::filter::{resolve_column, Condition};
use crate::input;
use crate::render::{self, RenderOptions};
use crate::tui::{self, App, Grid};
use crate::{ExportFormat, PeekArgs};

enum Failure {
    /// Exit 2: the request itself is unusable.
    Usage(anyhow::Error),
    /// Exit 1: reading or rendering failed.
    Other(anyhow::Error),
}

impl From<anyhow::Error> for Failure {
    fn from(error: anyhow::Error) -> Self {
        Failure::Other(error)
    }
}

impl From<std::io::Error> for Failure {
    fn from(error: std::io::Error) -> Self {
        Failure::Other(error.into())
    }
}

pub fn run(args: PeekArgs) -> ExitCode {
    let (code, error) = match peek(&args) {
        Ok(()) => return ExitCode::SUCCESS,
        Err(Failure::Usage(error)) => (2, error),
        Err(Failure::Other(error)) => (1, error),
    };
    eprintln!("error: {error:#}");
    ExitCode::from(code)
}

fn peek(args: &PeekArgs) -> Result<(), Failure> {
    if args.interactive {
        tui::preflight(&args.file).map_err(Failure::Usage)?;
    }
    let mut input = input::open(Some(args.file.as_path()), args.from).map_err(|error| {
        // Without `--from`, `open` refuses standard input before reading it.
        if args.from.is_none() && args.file == Path::new("-") {
            Failure::Usage(error)
        } else {
            Failure::Other(error)
        }
    })?;

    let sheet_names: Vec<String> = input.sheet_names().to_vec();
    let target = match &args.sheet {
        Some(name) => {
            if !sheet_names.iter().any(|n| n == name) {
                return Err(Failure::Other(anyhow!(
                    "sheet {name:?} not found; available: {}",
                    sheet_names.join(", ")
                )));
            }
            name.clone()
        }
        None => sheet_names
            .first()
            .cloned()
            .ok_or_else(|| anyhow!("workbook has no sheets"))?,
    };

    let sheet = input.raw_sheet(&target)?;
    let sheet = if args.filters.is_empty() && args.select.is_empty() {
        sheet
    } else {
        reshape(&sheet, &args.filters, &args.select).map_err(Failure::Usage)?
    };
    if args.interactive {
        return interactive(args, input, sheet_names, &target, &sheet);
    }

    let opts = RenderOptions {
        max_rows: if args.max_rows == 0 {
            None
        } else {
            Some(args.max_rows)
        },
        max_columns: (args.max_columns != 0).then_some(args.max_columns),
        max_width: args.max_width.max(3),
        all_sheets: &sheet_names,
    };

    let mut out = std::io::stdout().lock();
    match args.export {
        Some(ExportFormat::Csv) => render::csv(&mut out, &sheet, &opts)?,
        Some(ExportFormat::Json) => render::json(&mut out, &sheet, &opts)?,
        Some(ExportFormat::Markdown) => render::markdown(&mut out, &sheet, &opts)?,
        Some(ExportFormat::Text) => render::text(&mut out, &sheet, &opts)?,
        Some(ExportFormat::Box) | None => render::boxed(&mut out, &sheet, &opts)?,
    }
    Ok(())
}

/// The header row plus the data rows every `--where` condition accepts,
/// narrowed to the `--select` columns in the order named. Conditions may
/// name columns the selection drops.
fn reshape(sheet: &Sheet, filters: &[String], select: &[String]) -> Result<Sheet> {
    let headers = sheet.headers();
    let conditions = filters
        .iter()
        .map(|expression| {
            let condition = Condition::parse(expression)?;
            let column = resolve_column(&headers, condition.column())
                .with_context(|| format!("--where {expression:?}"))?;
            Ok((column, condition))
        })
        .collect::<Result<Vec<_>>>()?;
    let columns = select
        .iter()
        .map(|name| {
            let name = name.trim();
            if name.is_empty() {
                bail!("--select: empty column name; expected HEADER[,HEADER...]");
            }
            resolve_column(&headers, name).context("--select")
        })
        .collect::<Result<Vec<_>>>()?;
    let rows = sheet
        .rows()
        .iter()
        .enumerate()
        .filter(|(index, row)| *index == 0 || accepts(row, &conditions))
        .map(|(_, row)| project(sheet, row, &columns))
        .collect();
    Ok(Sheet::from_rows(sheet.name.clone(), rows))
}

fn accepts(row: &[Cell], conditions: &[(usize, Condition)]) -> bool {
    conditions
        .iter()
        .all(|(column, condition)| match row.get(*column) {
            Some(cell) => condition.matches(&cell.value),
            None => condition.matches(&CellValue::Empty),
        })
}

/// Copy `row`, keeping `columns` in order, or every column when none are
/// selected.
fn project(sheet: &Sheet, row: &[Cell], columns: &[usize]) -> Vec<Cell> {
    if columns.is_empty() {
        return row.iter().map(|cell| detached(sheet, cell)).collect();
    }
    columns
        .iter()
        .map(|&column| {
            row.get(column)
                .map_or_else(Cell::empty, |cell| detached(sheet, cell))
        })
        .collect()
}

/// A copy of `cell` that renders the same in a sheet built by
/// `Sheet::from_rows`, which drops the source sheet's error spelling: error
/// literals carry that spelling instead.
pub(crate) fn detached(sheet: &Sheet, cell: &Cell) -> Cell {
    let value = match &cell.value {
        CellValue::Error(literal) => CellValue::Error(sheet.error_text(literal).to_owned()),
        value => value.clone(),
    };
    Cell {
        value,
        number_format: cell.number_format.clone(),
    }
}

/// `--interactive`: browse `sheet`, the reshaped `target`. Tab loads each
/// other sheet on its first visit; `--where` and `--select` apply to every
/// sheet that has the columns they name, and a sheet without them opens
/// unfiltered with a note on the status line.
fn interactive(
    args: &PeekArgs,
    mut input: input::Input,
    names: Vec<String>,
    target: &str,
    sheet: &Sheet,
) -> Result<(), Failure> {
    let width = args.max_width.max(3);
    let reshaping = !(args.filters.is_empty() && args.select.is_empty());
    let current = names.iter().position(|name| name == target).unwrap_or(0);
    let app = App::new(
        args.file.display().to_string(),
        names,
        current,
        Grid::new(sheet, width, None),
    );
    tui::run(app, |name| match input.raw_sheet(name) {
        Err(error) => Grid::unavailable(name, format!("not loaded: {error:#}")),
        Ok(raw) if reshaping => match reshape(&raw, &args.filters, &args.select) {
            Ok(shaped) => Grid::new(&shaped, width, None),
            Err(_) => Grid::new(
                &raw,
                width,
                Some("unfiltered: this sheet lacks a --where/--select column".into()),
            ),
        },
        Ok(raw) => Grid::new(&raw, width, None),
    })
    .map_err(Failure::Other)
}
