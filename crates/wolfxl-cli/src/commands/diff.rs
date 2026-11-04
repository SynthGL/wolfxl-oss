//! `wolfxl diff LEFT RIGHT --key K`: reconcile two tables row by row on a key.
//!
//! Each side loads with its source spelling intact, so a key such as `007`
//! stays `007` and `$1,200.00` stays text until a comparison reads it as
//! money. Row 0 of each sheet is the header row. Rows match on the key
//! column after `--key-transform`; a key that is empty or repeated on one
//! side is an error rather than a guess. Rows whose every cell is blank are
//! spacing, not records, and are ignored. Cells compare as numbers when both
//! sides read as numbers (money spellings included), within `--tolerance`;
//! otherwise their text must match exactly.
//!
//! Exit codes: 0 when reconciled, 1 for unmatched rows or differences beyond
//! tolerance (any difference with `--strict-exit`), 2 for usage, input,
//! column, empty-key, and duplicate-key failures.

use std::collections::BTreeMap;
use std::io::Write;
use std::path::Path;
use std::process::ExitCode;

use anyhow::{anyhow, bail, Context, Result};
use serde_json::{json, Map, Value};
use wolfxl_core::{CellValue, Sheet};

use crate::filter::cell_number;
use crate::input::{self, InputFormat};
use crate::{DiffArgs, DiffFormat, KeyTransform};

/// Version of the JSON report shape.
const CONTRACT_VERSION: u32 = 1;

/// Stand-in for cells past the end of a short row.
static EMPTY: CellValue = CellValue::Empty;

pub(crate) fn run(args: DiffArgs) -> ExitCode {
    let report = match reconcile(&args) {
        Ok(report) => report,
        Err(error) => {
            eprintln!("error: {error:#}");
            return ExitCode::from(2);
        }
    };
    if !args.quiet {
        let document = match args.format {
            DiffFormat::Json => format!("{}\n", report.json()),
            DiffFormat::Text => report.text(),
        };
        let mut stdout = std::io::stdout().lock();
        if let Err(error) = stdout
            .write_all(document.as_bytes())
            .and_then(|()| stdout.flush())
        {
            if error.kind() != std::io::ErrorKind::BrokenPipe {
                eprintln!("error: writing report: {error}");
                return ExitCode::from(2);
            }
        }
    }
    let summary = &report.summary;
    let material = summary.only_left + summary.only_right + summary.diff_outside_tolerance > 0
        || (args.strict_exit && summary.diff > 0);
    if material {
        ExitCode::from(1)
    } else {
        ExitCode::SUCCESS
    }
}

/// One loaded side of the comparison.
struct Side {
    label: &'static str,
    sheet: Sheet,
    headers: Vec<String>,
}

impl Side {
    fn load(
        label: &'static str,
        file: &Path,
        from: Option<InputFormat>,
        sheet: Option<&str>,
    ) -> Result<Self> {
        let mut input =
            input::open(Some(file), from).with_context(|| format!("reading the {label} input"))?;
        let name = input
            .sheet_name(sheet)
            .with_context(|| format!("selecting the {label} sheet"))?;
        let sheet = input.raw_sheet(&name)?;
        let headers = sheet
            .rows()
            .first()
            .ok_or_else(|| anyhow!("{label} sheet {name:?} is empty; diff needs a header row"))?
            .iter()
            .map(|cell| text(&cell.value))
            .collect();
        Ok(Self {
            label,
            sheet,
            headers,
        })
    }

    /// Index of header `name`: the exact match, else the unique
    /// case-insensitive one. More than one candidate is an error.
    fn find(&self, name: &str) -> Result<Option<usize>> {
        let exact = self.positions(|header| header == name);
        let candidates = if exact.is_empty() {
            let folded = name.to_lowercase();
            self.positions(|header| header.to_lowercase() == folded)
        } else {
            exact
        };
        match candidates.as_slice() {
            [] => Ok(None),
            [index] => Ok(Some(*index)),
            _ => bail!(
                "column {name:?} matches {} headers on {}; headers must be distinct",
                candidates.len(),
                self.label
            ),
        }
    }

    fn require(&self, name: &str) -> Result<usize> {
        self.find(name)?.ok_or_else(|| {
            anyhow!(
                "no column {name:?} on {}; headers: {}",
                self.label,
                self.headers.join(", ")
            )
        })
    }

    fn positions(&self, accept: impl Fn(&str) -> bool) -> Vec<usize> {
        self.headers
            .iter()
            .enumerate()
            .filter(|(_, header)| accept(header.as_str()))
            .map(|(index, _)| index)
            .collect()
    }

    /// 1-based sheet row number of internal row `index`.
    fn row_number(&self, index: usize) -> u64 {
        u64::from(self.sheet.origin().0) + index as u64 + 1
    }

    fn value(&self, row: usize, column: usize) -> &CellValue {
        self.sheet.rows()[row]
            .get(column)
            .map_or(&EMPTY, |cell| &cell.value)
    }
}

/// A compared column: its left-side header and its index on each side.
struct Column {
    name: String,
    left: usize,
    right: usize,
}

#[derive(Copy, Clone)]
enum Status {
    Match,
    Diff,
    OnlyLeft,
    OnlyRight,
}

impl Status {
    fn name(self) -> &'static str {
        match self {
            Status::Match => "match",
            Status::Diff => "diff",
            Status::OnlyLeft => "only_left",
            Status::OnlyRight => "only_right",
        }
    }
}

/// A cell that differs between two matched rows.
struct Difference {
    /// Index into [`Report::columns`].
    column: usize,
    left: String,
    right: String,
    /// `right - left`, present when both cells read as numbers.
    delta: Option<f64>,
    within_tolerance: bool,
}

struct Entry {
    key: String,
    status: Status,
    left_row: Option<u64>,
    right_row: Option<u64>,
    differences: Vec<Difference>,
}

#[derive(Default)]
struct Summary {
    matched: usize,
    only_left: usize,
    only_right: usize,
    diff: usize,
    diff_outside_tolerance: usize,
}

struct Report {
    key: String,
    transform: KeyTransform,
    tolerance: f64,
    columns: Vec<Column>,
    left_sheet: String,
    right_sheet: String,
    summary: Summary,
    results: Vec<Entry>,
}

fn reconcile(args: &DiffArgs) -> Result<Report> {
    let (left_path, right_path) = (args.left.as_path(), args.right.as_path());
    let is_stdin = |path: &Path| path == Path::new("-");
    if is_stdin(left_path) && is_stdin(right_path) {
        bail!("only one side may read standard input");
    }
    if !args.tolerance.is_finite() || args.tolerance < 0.0 {
        bail!(
            "--tolerance must be a non-negative number, got {}",
            args.tolerance
        );
    }
    // With one side piped, --from names that side's format and the path side
    // keeps its extension; with two paths it overrides both extensions.
    let piped = is_stdin(left_path) || is_stdin(right_path);
    let from = |path: &Path| {
        if !piped || is_stdin(path) {
            args.from
        } else {
            None
        }
    };
    let sheet = args.sheet.as_deref();
    let left = Side::load("left", left_path, from(left_path), sheet)?;
    let right = Side::load("right", right_path, from(right_path), sheet)?;
    let key = (left.require(&args.key)?, right.require(&args.key)?);
    let columns = compared_columns(args, &left, &right, key)?;
    let left_keys = index_keys(&left, key.0, args.key_transform)?;
    let right_keys = index_keys(&right, key.1, args.key_transform)?;

    let mut pairs: BTreeMap<String, (Option<usize>, Option<usize>)> = BTreeMap::new();
    for (text, row) in left_keys {
        pairs.entry(text).or_default().0 = Some(row);
    }
    for (text, row) in right_keys {
        pairs.entry(text).or_default().1 = Some(row);
    }

    let mut summary = Summary::default();
    let results = pairs
        .into_iter()
        .map(|(text, rows)| {
            let (status, differences) = match rows {
                (Some(left_row), Some(right_row)) => {
                    summary.matched += 1;
                    let differences: Vec<Difference> = columns
                        .iter()
                        .enumerate()
                        .filter_map(|(index, column)| {
                            difference(
                                index,
                                left.value(left_row, column.left),
                                right.value(right_row, column.right),
                                args.tolerance,
                            )
                        })
                        .collect();
                    if differences.is_empty() {
                        (Status::Match, differences)
                    } else {
                        summary.diff += 1;
                        if differences.iter().any(|found| !found.within_tolerance) {
                            summary.diff_outside_tolerance += 1;
                        }
                        (Status::Diff, differences)
                    }
                }
                (Some(_), None) => {
                    summary.only_left += 1;
                    (Status::OnlyLeft, Vec::new())
                }
                (None, _) => {
                    summary.only_right += 1;
                    (Status::OnlyRight, Vec::new())
                }
            };
            Entry {
                key: text,
                status,
                left_row: rows.0.map(|index| left.row_number(index)),
                right_row: rows.1.map(|index| right.row_number(index)),
                differences,
            }
        })
        .collect();

    Ok(Report {
        key: left.headers[key.0].clone(),
        transform: args.key_transform,
        tolerance: args.tolerance,
        columns,
        left_sheet: left.sheet.name.clone(),
        right_sheet: right.sheet.name.clone(),
        summary,
        results,
    })
}

/// `--compare` columns, or every non-key header the two sides share.
fn compared_columns(
    args: &DiffArgs,
    left: &Side,
    right: &Side,
    key: (usize, usize),
) -> Result<Vec<Column>> {
    let explicit = !args.compare.is_empty();
    let names: Vec<&str> = if explicit {
        args.compare
            .iter()
            .map(|name| name.trim())
            .filter(|name| !name.is_empty())
            .collect()
    } else {
        left.headers
            .iter()
            .enumerate()
            .filter(|&(index, header)| index != key.0 && !header.trim().is_empty())
            .map(|(_, header)| header.as_str())
            .collect()
    };
    let mut columns: Vec<Column> = Vec::new();
    for name in names {
        let (left_index, right_index) = if explicit {
            (left.require(name)?, right.require(name)?)
        } else {
            match right.find(name)? {
                Some(right_index) => (left.require(name)?, right_index),
                None => continue,
            }
        };
        if left_index == key.0 || right_index == key.1 {
            if explicit {
                bail!("--compare column {name:?} is the key column");
            }
            continue;
        }
        if columns
            .iter()
            .all(|column| (column.left, column.right) != (left_index, right_index))
        {
            columns.push(Column {
                name: left.headers[left_index].clone(),
                left: left_index,
                right: right_index,
            });
        }
    }
    Ok(columns)
}

/// Map each transformed key to its internal row index. Empty and repeated
/// keys are errors that name the 1-based sheet rows involved.
fn index_keys(
    side: &Side,
    column: usize,
    transform: KeyTransform,
) -> Result<BTreeMap<String, usize>> {
    let mut rows: BTreeMap<String, Vec<usize>> = BTreeMap::new();
    let mut empty = Vec::new();
    for (index, row) in side.sheet.rows().iter().enumerate().skip(1) {
        if row.iter().all(|cell| is_blank(&cell.value)) {
            continue;
        }
        let key = transformed(text(side.value(index, column)), transform);
        if key.is_empty() {
            empty.push(side.row_number(index));
        } else {
            rows.entry(key).or_default().push(index);
        }
    }
    if !empty.is_empty() {
        bail!(
            "empty key in column {:?} on {} at rows {}",
            side.headers[column],
            side.label,
            join_rows(&empty)
        );
    }
    let duplicates: Vec<String> = rows
        .iter()
        .filter(|(_, found)| found.len() > 1)
        .map(|(key, found)| {
            let numbers: Vec<u64> = found.iter().map(|&index| side.row_number(index)).collect();
            format!(
                "duplicate key {key:?} on {} at rows {}",
                side.label,
                join_rows(&numbers)
            )
        })
        .collect();
    if !duplicates.is_empty() {
        bail!("{}", duplicates.join("; "));
    }
    Ok(rows
        .into_iter()
        .map(|(key, found)| (key, found[0]))
        .collect())
}

fn transformed(key: String, transform: KeyTransform) -> String {
    match transform {
        KeyTransform::None => key,
        KeyTransform::Trim => key.trim().to_string(),
        KeyTransform::Lower => key.trim().to_lowercase(),
        KeyTransform::Digits => key.chars().filter(char::is_ascii_digit).collect(),
    }
}

fn transform_name(transform: KeyTransform) -> &'static str {
    match transform {
        KeyTransform::None => "none",
        KeyTransform::Trim => "trim",
        KeyTransform::Lower => "lower",
        KeyTransform::Digits => "digits",
    }
}

/// Compare one cell pair. Numbers (money spellings included) compare by
/// value against `tolerance`; anything else must match as text. Two blanks
/// are equal.
fn difference(
    column: usize,
    left: &CellValue,
    right: &CellValue,
    tolerance: f64,
) -> Option<Difference> {
    if is_blank(left) && is_blank(right) {
        return None;
    }
    let delta = match (cell_number(left), cell_number(right)) {
        (Some(left), Some(right)) => {
            let delta = delta(left, right);
            if delta == 0.0 {
                return None;
            }
            Some(delta)
        }
        _ => None,
    };
    let (left, right) = (text(left), text(right));
    if delta.is_none() && left == right {
        return None;
    }
    Some(Difference {
        column,
        left,
        right,
        delta,
        within_tolerance: delta.is_some_and(|delta| delta.abs() <= tolerance),
    })
}

/// `right - left` kept to the 15 significant digits of the larger operand,
/// so `100.004 - 100` is `0.004` rather than its binary residue and values
/// equal at Excel's display precision have no difference.
fn delta(left: f64, right: f64) -> f64 {
    let raw = right - left;
    if raw == 0.0 || !raw.is_finite() {
        return raw;
    }
    let scale = left.abs().max(right.abs()).log10().floor();
    let digits = raw.abs().log10().floor() - scale + 15.0;
    if digits < 1.0 {
        return 0.0;
    }
    let precision = digits as usize - 1;
    format!("{raw:.precision$e}").parse().unwrap_or(raw)
}

/// Blank cells and whitespace-only text.
fn is_blank(value: &CellValue) -> bool {
    match value {
        CellValue::Empty => true,
        CellValue::String(text) => text.trim().is_empty(),
        _ => false,
    }
}

/// Source text of a cell: strings as stored, integers as digits, floats in
/// shortest round-trip form, booleans as `TRUE`/`FALSE`, dates and times in
/// ISO 8601, errors as their literal.
fn text(value: &CellValue) -> String {
    match value {
        CellValue::Empty => String::new(),
        CellValue::String(text) => text.clone(),
        CellValue::Int(integer) => integer.to_string(),
        CellValue::Float(float) => float.to_string(),
        CellValue::Bool(true) => "TRUE".to_string(),
        CellValue::Bool(false) => "FALSE".to_string(),
        CellValue::Date(date) => date.format("%Y-%m-%d").to_string(),
        CellValue::DateTime(datetime) => datetime.format("%Y-%m-%dT%H:%M:%S%.f").to_string(),
        CellValue::Time(time) => time.format("%H:%M:%S%.f").to_string(),
        CellValue::Error(literal) => literal.clone(),
    }
}

fn join_rows(numbers: &[u64]) -> String {
    numbers
        .iter()
        .map(u64::to_string)
        .collect::<Vec<_>>()
        .join(", ")
}

/// Integral values within f64's exact-integer range print without a fraction.
fn json_number(number: f64) -> Value {
    if number.fract() == 0.0 && number.abs() < 9_007_199_254_740_992.0 {
        Value::from(number as i64)
    } else {
        serde_json::Number::from_f64(number).map_or(Value::Null, Value::Number)
    }
}

/// Blank cells are `null`; everything else is its source text.
fn cell_json(text: &str) -> Value {
    if text.is_empty() {
        Value::Null
    } else {
        Value::from(text)
    }
}

/// Text-format cell: quoted source text, or `blank`.
fn cell_text(text: &str) -> String {
    if text.is_empty() {
        "blank".to_string()
    } else {
        format!("{text:?}")
    }
}

impl Report {
    fn json(&self) -> Value {
        let summary = &self.summary;
        let results: Vec<Value> = self
            .results
            .iter()
            .map(|entry| {
                let differences: Vec<Value> = entry
                    .differences
                    .iter()
                    .map(|difference| self.difference_json(difference))
                    .collect();
                json!({
                    "key": entry.key,
                    "status": entry.status.name(),
                    "left_row": entry.left_row,
                    "right_row": entry.right_row,
                    "differences": differences,
                })
            })
            .collect();
        let compare: Vec<&str> = self
            .columns
            .iter()
            .map(|column| column.name.as_str())
            .collect();
        json!({
            "contract_version": CONTRACT_VERSION,
            "key": self.key,
            "key_transform": transform_name(self.transform),
            "compare": compare,
            "tolerance": json_number(self.tolerance),
            "left": {"sheet": self.left_sheet},
            "right": {"sheet": self.right_sheet},
            "summary": {
                "matched": summary.matched,
                "only_left": summary.only_left,
                "only_right": summary.only_right,
                "diff": summary.diff,
                "diff_outside_tolerance": summary.diff_outside_tolerance,
            },
            "results": results,
        })
    }

    fn difference_json(&self, difference: &Difference) -> Value {
        let mut object = Map::new();
        object.insert(
            "column".to_string(),
            Value::from(self.columns[difference.column].name.as_str()),
        );
        object.insert("left".to_string(), cell_json(&difference.left));
        object.insert("right".to_string(), cell_json(&difference.right));
        if let Some(delta) = difference.delta {
            object.insert("delta".to_string(), json_number(delta));
        }
        object.insert(
            "within_tolerance".to_string(),
            Value::Bool(difference.within_tolerance),
        );
        Value::Object(object)
    }

    /// A summary line, then one line per key that is not a clean match.
    fn text(&self) -> String {
        let summary = &self.summary;
        let mut lines = vec![format!(
            "matched {}, only_left {}, only_right {}, diff {}, diff_outside_tolerance {}",
            summary.matched,
            summary.only_left,
            summary.only_right,
            summary.diff,
            summary.diff_outside_tolerance
        )];
        for entry in &self.results {
            let left_row = entry.left_row.unwrap_or_default();
            let right_row = entry.right_row.unwrap_or_default();
            match entry.status {
                Status::Match => {}
                Status::OnlyLeft => {
                    lines.push(format!("only_left {:?} (left row {left_row})", entry.key));
                }
                Status::OnlyRight => {
                    lines.push(format!(
                        "only_right {:?} (right row {right_row})",
                        entry.key
                    ));
                }
                Status::Diff => {
                    let cells: Vec<String> = entry
                        .differences
                        .iter()
                        .map(|difference| self.difference_text(difference))
                        .collect();
                    lines.push(format!(
                        "diff {:?} (left row {left_row}, right row {right_row}): {}",
                        entry.key,
                        cells.join("; ")
                    ));
                }
            }
        }
        let mut document = lines.join("\n");
        document.push('\n');
        document
    }

    fn difference_text(&self, difference: &Difference) -> String {
        let cells = format!(
            "{} {} -> {}",
            self.columns[difference.column].name,
            cell_text(&difference.left),
            cell_text(&difference.right)
        );
        match difference.delta {
            Some(delta) => {
                let side = if difference.within_tolerance {
                    "within"
                } else {
                    "outside"
                };
                format!("{cells} (delta {delta}, {side} tolerance)")
            }
            None => cells,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn deltas_keep_the_operands_significant_digits() {
        assert_eq!(delta(100.0, 100.004), 0.004);
        assert_eq!(delta(0.1 + 0.2, 0.3), 0.0);
        assert_eq!(delta(101.0, 100.0), -1.0);
        assert_eq!(delta(5.0, 5.0), 0.0);
    }

    #[test]
    fn key_transforms_normalize_source_text() {
        let key = || "  INV-0042 Acme ".to_string();
        assert_eq!(transformed(key(), KeyTransform::None), "  INV-0042 Acme ");
        assert_eq!(transformed(key(), KeyTransform::Trim), "INV-0042 Acme");
        assert_eq!(transformed(key(), KeyTransform::Lower), "inv-0042 acme");
        assert_eq!(transformed(key(), KeyTransform::Digits), "0042");
    }
}
