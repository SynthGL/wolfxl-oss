//! `wolfxl peek FILE|-`: preview one sheet.
//!
//! `--add NAME=FORMULA` (repeatable) first appends formula columns computed
//! over every data row (see `crate::derive`). `--where COLUMN<op>VALUE`
//! (repeatable, AND) then keeps the header row plus the data rows every
//! condition accepts; `--select A,B` keeps the named columns in the order
//! given. All three run before rendering, so `--max-rows`, `--max-columns`
//! and every export format see the reshaped sheet. Without them the sheet
//! renders exactly as loaded.
//!
//! `--interactive` browses the reshaped sheet in a terminal grid instead of
//! printing it (see `crate::tui`).
//!
//! Exit codes: 0 success, 1 input or rendering failure, 2 a malformed or
//! unresolvable `--where`/`--select`, a malformed or unsupported `--add`
//! formula or a taken `--add` name, standard input without `--from`, input
//! over its `--from` spool limit, or `--interactive` without a terminal on
//! standard output or with standard input as the data source.

use std::path::Path;
use std::process::ExitCode;

use anyhow::{anyhow, bail, Context, Result};
use wolfxl_core::{Cell, CellValue, Sheet};

use crate::derive;
use crate::filter::{resolve_column, Condition};
use crate::input::{self, Provenance};
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

impl From<derive::Failure> for Failure {
    fn from(failure: derive::Failure) -> Self {
        match failure {
            derive::Failure::Usage(error) => Failure::Usage(error),
            derive::Failure::Input(error) => Failure::Other(error),
        }
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
    if args.addresses && matches!(args.export, Some(ExportFormat::Csv | ExportFormat::Json)) {
        return Err(Failure::Usage(anyhow!(
            "--addresses is available for box, text and markdown previews, not CSV or JSON"
        )));
    }
    if args.interactive {
        tui::preflight(&args.file).map_err(Failure::Usage)?;
    }
    let mut input = input::open(Some(args.file.as_path()), args.from).map_err(|error| {
        // Without `--from`, `open` refuses standard input before reading it;
        // input over its `--from` spool limit is refused the same way.
        if (args.from.is_none() && args.file == Path::new("-"))
            || error.downcast_ref::<input::SpoolLimitExceeded>().is_some()
        {
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

    let (sheet, mut provenance) =
        input::header_sheet(input.raw_sheet(&target)?, args.header_row).map_err(Failure::Usage)?;
    let columns = if args.add.is_empty() {
        None
    } else {
        Some(derive::Columns::parse(&args.add).map_err(Failure::Usage)?)
    };
    let sheet = match &columns {
        Some(columns) => columns.append(&mut input, &target, &sheet, &provenance)?,
        None => sheet,
    };
    provenance.columns.resize(sheet.headers().len(), None);
    let sheet = if args.filters.is_empty() && args.select.is_empty() {
        sheet
    } else {
        let (sheet, shaped_provenance) =
            reshape(&sheet, &provenance, &args.filters, &args.select).map_err(Failure::Usage)?;
        provenance = shaped_provenance;
        sheet
    };
    if args.interactive {
        return interactive(
            args,
            input,
            sheet_names,
            &target,
            &sheet,
            &provenance,
            columns.as_ref(),
        );
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
        addresses: args.addresses.then_some(&provenance),
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
fn reshape(
    sheet: &Sheet,
    provenance: &Provenance,
    filters: &[String],
    select: &[String],
) -> Result<(Sheet, Provenance)> {
    let headers = sheet.headers();
    let conditions = filters
        .iter()
        .map(|expression| {
            let condition = Condition::parse(expression)?;
            let column = resolve_column(&headers, provenance, condition.column())
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
            resolve_column(&headers, provenance, name).context("--select")
        })
        .collect::<Result<Vec<_>>>()?;
    let mut source_rows = Vec::new();
    let rows = sheet
        .rows()
        .iter()
        .enumerate()
        .filter(|(index, row)| *index == 0 || accepts(row, &conditions))
        .map(|(index, row)| {
            source_rows.push(provenance.rows[index]);
            project(sheet, row, &columns)
        })
        .collect();
    let source_columns = if columns.is_empty() {
        provenance.columns.clone()
    } else {
        columns
            .iter()
            .map(|&column| provenance.columns[column])
            .collect()
    };
    Ok((
        Sheet::from_rows(sheet.name.clone(), rows),
        Provenance {
            rows: source_rows,
            columns: source_columns,
        },
    ))
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
/// other sheet on its first visit and reshapes it the same way (see
/// [`tab_sheet`]).
fn interactive(
    args: &PeekArgs,
    mut input: input::Input,
    names: Vec<String>,
    target: &str,
    sheet: &Sheet,
    provenance: &Provenance,
    columns: Option<&derive::Columns>,
) -> Result<(), Failure> {
    let width = args.max_width.max(3);
    let current = names.iter().position(|name| name == target).unwrap_or(0);
    let app = App::new(
        args.file.display().to_string(),
        names,
        current,
        Grid::with_provenance(sheet, width, None, provenance),
    );
    tui::run(app, |name| {
        match tab_sheet(
            &mut input,
            name,
            columns,
            args.header_row,
            &args.filters,
            &args.select,
        ) {
            Ok((sheet, provenance, note)) => {
                Grid::with_provenance(&sheet, width, note, &provenance)
            }
            Err(error) => Grid::unavailable(name, format!("not loaded: {error:#}")),
        }
    })
    .map_err(Failure::Other)
}

/// Sheet `name` as Tab first shows it, with its status-line note. The
/// `--add` columns are appended unless this sheet already has a column of
/// that name; `--where` and `--select` then apply when this sheet has the
/// columns they name. Whatever this sheet cannot take is skipped and named
/// in the note.
fn tab_sheet(
    input: &mut input::Input,
    name: &str,
    columns: Option<&derive::Columns>,
    header_row: Option<u32>,
    filters: &[String],
    select: &[String],
) -> Result<(Sheet, Provenance, Option<String>)> {
    let (raw, mut provenance) = input::header_sheet(input.raw_sheet(name)?, header_row)?;
    let mut notes = Vec::new();
    let sheet = match columns {
        None => raw,
        Some(columns) => match columns.append(input, name, &raw, &provenance) {
            Ok(derived) => derived,
            Err(derive::Failure::Usage(error)) => {
                notes.push(format!("no added columns: {error:#}"));
                raw
            }
            Err(derive::Failure::Input(error)) => return Err(error),
        },
    };
    provenance.columns.resize(sheet.headers().len(), None);
    let sheet = if filters.is_empty() && select.is_empty() {
        sheet
    } else {
        match reshape(&sheet, &provenance, filters, select) {
            Ok((shaped, shaped_provenance)) => {
                provenance = shaped_provenance;
                shaped
            }
            Err(_) => {
                notes.push("unfiltered: this sheet lacks a --where/--select column".to_owned());
                sheet
            }
        }
    };
    let note = (!notes.is_empty()).then(|| notes.join("; "));
    Ok((sheet, provenance, note))
}

#[cfg(test)]
mod tests {
    use super::*;

    /// `sample-minimal.csv` and the name of its one sheet.
    fn minimal() -> (input::Input, String) {
        let path = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("tests/fixtures")
            .join("sample-minimal.csv");
        let input = input::open(Some(path.as_path()), None).expect("open sample-minimal.csv");
        let name = input.sheet_names()[0].clone();
        (input, name)
    }

    fn strings(values: &[&str]) -> Vec<String> {
        values.iter().map(|value| (*value).to_owned()).collect()
    }

    fn shown(sheet: &Sheet) -> Vec<Vec<String>> {
        sheet
            .rows()
            .iter()
            .map(|row| {
                row.iter()
                    .map(|cell| match &cell.value {
                        CellValue::String(text) => text.clone(),
                        CellValue::Empty => String::new(),
                        other => format!("{other:?}"),
                    })
                    .collect()
            })
            .collect()
    }

    #[test]
    fn tab_appends_then_filters_a_sheet_that_takes_both() {
        let (mut input, name) = minimal();
        let columns = derive::Columns::parse(&strings(&["Growth=C2-B2"])).expect("parses");
        let (sheet, _, note) = tab_sheet(
            &mut input,
            &name,
            Some(&columns),
            None,
            &strings(&["Growth>10000"]),
            &strings(&["Account", "Growth"]),
        )
        .expect("loads");
        assert_eq!(note, None);
        let expected = [
            ["Account", "Growth"],
            ["Revenue", "25000"],
            ["COGS", "18000"],
        ];
        assert_eq!(shown(&sheet), expected);
    }

    #[test]
    fn tab_notes_an_added_name_the_sheet_already_has() {
        let (mut input, name) = minimal();
        let columns = derive::Columns::parse(&strings(&["jan=B2*2"])).expect("parses");
        let (sheet, _, note) =
            tab_sheet(&mut input, &name, Some(&columns), None, &[], &[]).expect("loads");
        let note = note.expect("a note");
        assert!(note.starts_with("no added columns: "), "{note}");
        assert!(note.contains("already exists"), "{note}");
        assert_eq!(sheet.headers(), ["Account", "Jan", "Feb", "Mar", "Q1"]);
    }

    #[test]
    fn tab_keeps_added_columns_when_a_filter_does_not_fit() {
        let (mut input, name) = minimal();
        let columns = derive::Columns::parse(&strings(&["Growth=C2-B2"])).expect("parses");
        let filters = strings(&["Profit>0"]);
        let (sheet, _, note) =
            tab_sheet(&mut input, &name, Some(&columns), None, &filters, &[]).expect("loads");
        assert_eq!(
            note.as_deref(),
            Some("unfiltered: this sheet lacks a --where/--select column")
        );
        assert_eq!(
            sheet.headers(),
            ["Account", "Jan", "Feb", "Mar", "Q1", "Growth"]
        );
        assert_eq!(sheet.rows().len(), 4);
    }

    #[test]
    fn tab_notes_both_when_neither_fits() {
        let (mut input, name) = minimal();
        let columns = derive::Columns::parse(&strings(&["Q1=B2"])).expect("parses");
        let (_, _, note) = tab_sheet(
            &mut input,
            &name,
            Some(&columns),
            None,
            &strings(&["Q1>0"]),
            &strings(&["Q1", "Missing"]),
        )
        .expect("loads");
        let note = note.expect("a note");
        assert!(note.starts_with("no added columns: "), "{note}");
        assert!(
            note.ends_with("; unfiltered: this sheet lacks a --where/--select column"),
            "{note}"
        );
    }
}
