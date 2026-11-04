//! `wolfxl pivot FILE --rows R[,R2] [--column C] --values AGG:FIELD[,...]`:
//! group a sheet's data rows by header-named fields and aggregate.
//!
//! Row 0 of the selected sheet is the header row. Fields resolve by exact
//! header text in source spelling (`001` and `1` are distinct), then by a
//! unique case-insensitive match. Each `--rows` field
//! contributes one key column; `--column` spreads its distinct values across
//! columns and adds a grand-total column. A grand-total row closes the table.
//!
//! Aggregation follows the stored cell types: `Int` and `Float` cells are
//! numeric, everything else (text, booleans, dates, errors) is skipped by
//! `sum`, `avg`, `min`, and `max`, and every non-empty cell is counted by
//! `count`. Totals recompute the aggregation over all contributing rows.
//!
//! Keys sort numerically when every non-blank key of a field is numeric,
//! otherwise by display text in code point order. Blank keys render as
//! `(blank)` and sort last.
//!
//! Exit codes: 0 for a table, 2 for usage and input failures.

use std::cmp::Ordering;
use std::collections::HashMap;
use std::io::Write;
use std::process::ExitCode;

use anyhow::{anyhow, bail, Result};
use serde_json::{json, Value};
use wolfxl_core::{Cell, CellValue, Sheet};

use crate::render::{self, RenderOptions};
use crate::{input, PivotArgs, PivotFormat};

const CONTRACT_VERSION: u32 = 1;
const BLANK_KEY: &str = "(blank)";
const TOTAL: &str = "Total";
/// 2^53: integral values below it are exact and print without a fraction.
const MAX_EXACT_INTEGER: f64 = 9_007_199_254_740_992.0;
/// Top-level `--export json` keys, in output order.
const JSON_KEYS: [&str; 9] = [
    "contract_version",
    "sheet",
    "rows",
    "column",
    "values",
    "column_keys",
    "headers",
    "data",
    "totals",
];

pub(crate) fn run(args: PivotArgs) -> ExitCode {
    let (table, sheet_names) = match build(&args) {
        Ok(built) => built,
        Err(error) => {
            eprintln!("error: {error:#}");
            return ExitCode::from(2);
        }
    };
    let options = RenderOptions {
        max_rows: None,
        max_columns: None,
        max_width: usize::MAX,
        all_sheets: &sheet_names,
    };
    let mut document = Vec::new();
    let rendered = match args.export {
        PivotFormat::Box => render::boxed(&mut document, &table.to_sheet(), &options),
        PivotFormat::Csv => render::csv(&mut document, &table.to_sheet(), &options),
        PivotFormat::Json => table.write_json(&mut document),
    };
    if let Err(error) = rendered {
        eprintln!("error: rendering pivot: {error}");
        return ExitCode::from(2);
    }
    let mut stdout = std::io::stdout().lock();
    if let Err(error) = stdout.write_all(&document).and_then(|()| stdout.flush()) {
        if error.kind() != std::io::ErrorKind::BrokenPipe {
            eprintln!("error: writing result: {error}");
            return ExitCode::from(2);
        }
    }
    ExitCode::SUCCESS
}

/// The `--export json` table, built without writing anything.
pub(crate) fn structured(args: &PivotArgs) -> Result<Value> {
    Ok(build(args)?.0.json())
}

#[derive(Copy, Clone, Debug)]
enum Agg {
    Sum,
    Count,
    Avg,
    Min,
    Max,
}

impl Agg {
    fn parse(name: &str) -> Option<Self> {
        Some(match name.to_ascii_lowercase().as_str() {
            "sum" => Agg::Sum,
            "count" => Agg::Count,
            "avg" => Agg::Avg,
            "min" => Agg::Min,
            "max" => Agg::Max,
            _ => return None,
        })
    }

    fn name(self) -> &'static str {
        match self {
            Agg::Sum => "sum",
            Agg::Count => "count",
            Agg::Avg => "avg",
            Agg::Min => "min",
            Agg::Max => "max",
        }
    }
}

/// Running aggregate of one value field over a set of rows.
#[derive(Copy, Clone, Debug, Default)]
struct Acc {
    /// Non-empty cells.
    count: u64,
    /// Numeric cells.
    numbers: u64,
    sum: f64,
    min: f64,
    max: f64,
}

impl Acc {
    fn add(&mut self, value: &CellValue) {
        if is_blank(value) {
            return;
        }
        self.count += 1;
        let number = match value {
            CellValue::Int(integer) => *integer as f64,
            CellValue::Float(float) => *float,
            _ => return,
        };
        if self.numbers == 0 {
            self.min = number;
            self.max = number;
        } else {
            self.min = self.min.min(number);
            self.max = self.max.max(number);
        }
        self.sum += number;
        self.numbers += 1;
    }

    /// The aggregate, or `None` when a numeric aggregation saw no numbers.
    fn result(&self, agg: Agg) -> Option<f64> {
        let numeric = self.numbers > 0;
        match agg {
            Agg::Count => Some(self.count as f64),
            Agg::Sum => numeric.then(|| rounded(self.sum)),
            Agg::Avg => numeric.then(|| rounded(self.sum / self.numbers as f64)),
            Agg::Min => numeric.then_some(self.min),
            Agg::Max => numeric.then_some(self.max),
        }
    }
}

/// One group key. Numbers are identified by their 15-significant-digit
/// value, so `0.1+0.2` and `0.3` group together.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
enum Key {
    Blank,
    Number { bits: u64, text: String },
    Text(String),
}

impl Key {
    fn of(cell: Option<&Cell>, sheet: &Sheet) -> Self {
        let Some(cell) = cell else {
            return Key::Blank;
        };
        match &cell.value {
            CellValue::Empty => Key::Blank,
            CellValue::String(text) if text.trim().is_empty() => Key::Blank,
            CellValue::String(text) => Key::Text(text.clone()),
            CellValue::Int(integer) => Key::number(*integer as f64),
            CellValue::Float(float) => Key::number(*float),
            CellValue::Bool(boolean) => Key::Text(boolean.to_string()),
            CellValue::Date(date) => Key::Text(date.format("%Y-%m-%d").to_string()),
            CellValue::DateTime(datetime) => {
                Key::Text(datetime.format("%Y-%m-%d %H:%M:%S").to_string())
            }
            CellValue::Time(time) => Key::Text(time.format("%H:%M:%S").to_string()),
            CellValue::Error(literal) => Key::Text(sheet.error_text(literal).to_owned()),
        }
    }

    fn number(number: f64) -> Self {
        let number = rounded(number);
        // Fold -0 into 0 so both group together.
        let number = if number == 0.0 { 0.0 } else { number };
        Key::Number {
            bits: number.to_bits(),
            text: number.to_string(),
        }
    }

    fn value(&self) -> Option<f64> {
        match self {
            Key::Number { bits, .. } => Some(f64::from_bits(*bits)),
            _ => None,
        }
    }

    fn text(&self) -> &str {
        match self {
            Key::Blank => BLANK_KEY,
            Key::Number { text, .. } | Key::Text(text) => text.as_str(),
        }
    }

    fn json(&self) -> Value {
        match self.value() {
            Some(number) => json_number(number),
            None => Value::String(self.text().to_string()),
        }
    }

    /// Tie-break rank for keys whose display text is equal.
    fn rank(&self) -> u8 {
        match self {
            Key::Number { .. } => 0,
            Key::Text(_) => 1,
            Key::Blank => 2,
        }
    }
}

/// Blanks last; then numeric or lexical order. Total over distinct keys.
fn compare_keys(left: &Key, right: &Key, numeric: bool) -> Ordering {
    match (left, right) {
        (Key::Blank, Key::Blank) => Ordering::Equal,
        (Key::Blank, _) => Ordering::Greater,
        (_, Key::Blank) => Ordering::Less,
        _ => match (numeric, left.value(), right.value()) {
            (true, Some(left), Some(right)) => left.total_cmp(&right),
            _ => left
                .text()
                .cmp(right.text())
                .then_with(|| left.rank().cmp(&right.rank()))
                .then_with(|| {
                    let (left, right) = (left.value(), right.value());
                    left.unwrap_or_default()
                        .total_cmp(&right.unwrap_or_default())
                }),
        },
    }
}

/// Whether every non-blank key is numeric.
fn all_numeric<'a>(mut keys: impl Iterator<Item = &'a Key>) -> bool {
    keys.all(|key| !matches!(key, Key::Text(_)))
}

struct Table {
    sheet: String,
    row_fields: Vec<String>,
    column_field: Option<String>,
    value_labels: Vec<String>,
    column_keys: Vec<Key>,
    headers: Vec<String>,
    rows: Vec<(Vec<Key>, Vec<Option<f64>>)>,
    totals: Vec<Option<f64>>,
}

impl Table {
    /// Header row, data rows, then the grand-total row.
    fn to_sheet(&self) -> Sheet {
        let mut grid: Vec<Vec<Cell>> = Vec::with_capacity(self.rows.len() + 2);
        grid.push(
            self.headers
                .iter()
                .map(|header| text_cell(header.as_str()))
                .collect(),
        );
        for (keys, values) in &self.rows {
            grid.push(
                keys.iter()
                    .map(|key| text_cell(key.text()))
                    .chain(values.iter().map(|value| value_cell(*value)))
                    .collect(),
            );
        }
        grid.push(
            std::iter::once(text_cell(TOTAL))
                .chain(
                    std::iter::repeat_with(Cell::empty)
                        .take(self.row_fields.len().saturating_sub(1)),
                )
                .chain(self.totals.iter().map(|value| value_cell(*value)))
                .collect(),
        );
        Sheet::from_rows(format!("{} (pivot)", self.sheet), grid)
    }

    /// The JSON table object.
    fn json(&self) -> Value {
        let column_keys: Vec<Value> = self.column_keys.iter().map(Key::json).collect();
        let data: Vec<Value> = self
            .rows
            .iter()
            .map(|(keys, cells)| {
                Value::Array(
                    keys.iter()
                        .map(Key::json)
                        .chain(cells.iter().map(json_cell))
                        .collect(),
                )
            })
            .collect();
        let totals: Vec<Value> = self.totals.iter().map(json_cell).collect();
        json!({
            "contract_version": CONTRACT_VERSION,
            "sheet": self.sheet,
            "rows": self.row_fields,
            "column": self.column_field,
            "values": self.value_labels,
            "column_keys": column_keys,
            "headers": self.headers,
            "data": data,
            "totals": totals,
        })
    }

    /// Pretty top-level object with one inline array per line, like `peek -e json`.
    fn write_json<W: Write>(&self, w: &mut W) -> std::io::Result<()> {
        let document = self.json();
        writeln!(w, "{{")?;
        for (index, key) in JSON_KEYS.iter().enumerate() {
            let comma = if index + 1 == JSON_KEYS.len() {
                ""
            } else {
                ","
            };
            let name = Value::from(*key);
            match &document[*key] {
                Value::Array(rows) if *key == "data" && !rows.is_empty() => {
                    writeln!(w, "  {name}: [")?;
                    for (index, row) in rows.iter().enumerate() {
                        let separator = if index + 1 == rows.len() { "" } else { "," };
                        writeln!(w, "    {}{separator}", inline(row))?;
                    }
                    writeln!(w, "  ]{comma}")?;
                }
                value => writeln!(w, "  {name}: {}{comma}", inline(value))?,
            }
        }
        writeln!(w, "}}")
    }
}

/// Inline JSON text with `", "` between array items. Pivot documents hold no
/// nested objects.
fn inline(value: &Value) -> String {
    match value {
        Value::Array(items) => {
            let items: Vec<String> = items.iter().map(inline).collect();
            format!("[{}]", items.join(", "))
        }
        scalar => scalar.to_string(),
    }
}

fn json_cell(value: &Option<f64>) -> Value {
    value.map_or(Value::Null, json_number)
}

fn build(args: &PivotArgs) -> Result<(Table, Vec<String>)> {
    let specs = args
        .values
        .iter()
        .map(|entry| parse_value(entry.as_str()))
        .collect::<Result<Vec<_>>>()?;
    let mut input = input::open(Some(args.file.as_path()), args.from)?;
    let sheet_name = input.sheet_name(args.sheet.as_deref())?;
    let sheet_names = input.sheet_names().to_vec();
    // Header names keep their source spelling (`001` stays distinct from
    // `1`); the data rows below are typed for grouping and aggregation.
    let headers = input.raw_sheet(&sheet_name)?.headers();
    let sheet = input.sheet(&sheet_name)?;

    let row_columns = args
        .rows
        .iter()
        .map(|field| resolve(&headers, field))
        .collect::<Result<Vec<_>>>()?;
    let column = args
        .column
        .as_deref()
        .map(|field| resolve(&headers, field))
        .transpose()?;
    let values = specs
        .iter()
        .map(|&(agg, field)| resolve(&headers, field).map(|index| (agg, index)))
        .collect::<Result<Vec<_>>>()?;
    let width = values.len();

    let mut group_index: HashMap<Vec<Key>, usize> = HashMap::new();
    let mut column_index: HashMap<Key, usize> = HashMap::new();
    let mut group_totals: Vec<Vec<Acc>> = Vec::new();
    let mut column_totals: Vec<Vec<Acc>> = Vec::new();
    let mut cross: HashMap<(usize, usize), Vec<Acc>> = HashMap::new();
    let mut grand = vec![Acc::default(); width];

    for row in sheet.rows().iter().skip(1) {
        if row.iter().all(|cell| is_blank(&cell.value)) {
            continue;
        }
        let key: Vec<Key> = row_columns
            .iter()
            .map(|&index| Key::of(row.get(index), &sheet))
            .collect();
        let next = group_index.len();
        let group = *group_index.entry(key).or_insert(next);
        if group == group_totals.len() {
            group_totals.push(vec![Acc::default(); width]);
        }
        let column_key = column.map(|index| {
            let next = column_index.len();
            let slot = *column_index
                .entry(Key::of(row.get(index), &sheet))
                .or_insert(next);
            if slot == column_totals.len() {
                column_totals.push(vec![Acc::default(); width]);
            }
            slot
        });
        let mut intersection = column_key.map(|slot| {
            cross
                .entry((group, slot))
                .or_insert_with(|| vec![Acc::default(); width])
        });
        for (index, &(_, field)) in values.iter().enumerate() {
            let Some(cell) = row.get(field) else {
                continue;
            };
            grand[index].add(&cell.value);
            group_totals[group][index].add(&cell.value);
            if let Some(slot) = column_key {
                column_totals[slot][index].add(&cell.value);
            }
            if let Some(accs) = intersection.as_deref_mut() {
                accs[index].add(&cell.value);
            }
        }
    }

    let mut groups: Vec<(Vec<Key>, usize)> = group_index.into_iter().collect();
    let numeric: Vec<bool> = (0..row_columns.len())
        .map(|field| all_numeric(groups.iter().map(|(key, _)| &key[field])))
        .collect();
    groups.sort_by(|(left, _), (right, _)| {
        left.iter()
            .zip(right)
            .zip(&numeric)
            .map(|((left, right), &numeric)| compare_keys(left, right, numeric))
            .find(|ordering| ordering.is_ne())
            .unwrap_or(Ordering::Equal)
    });
    let mut columns: Vec<(Key, usize)> = column_index.into_iter().collect();
    let columns_numeric = all_numeric(columns.iter().map(|(key, _)| key));
    columns.sort_by(|(left, _), (right, _)| compare_keys(left, right, columns_numeric));

    let empty = vec![Acc::default(); width];
    let rows = groups
        .into_iter()
        .map(|(key, group)| {
            let mut cells = Vec::with_capacity(width * (columns.len() + 1));
            for &(_, slot) in &columns {
                let accs = cross.get(&(group, slot)).unwrap_or(&empty);
                cells.extend(aggregates(&values, accs));
            }
            cells.extend(aggregates(&values, &group_totals[group]));
            (key, cells)
        })
        .collect();
    let mut totals = Vec::with_capacity(width * (columns.len() + 1));
    for &(_, slot) in &columns {
        totals.extend(aggregates(&values, &column_totals[slot]));
    }
    totals.extend(aggregates(&values, &grand));

    let value_labels: Vec<String> = values
        .iter()
        .map(|&(agg, field)| format!("{}:{}", agg.name(), headers[field]))
        .collect();
    let row_fields: Vec<String> = row_columns
        .iter()
        .map(|&index| headers[index].clone())
        .collect();
    let mut table_headers = row_fields.clone();
    if column.is_some() {
        let single = value_labels.len() == 1;
        for (key, _) in &columns {
            for label in &value_labels {
                table_headers.push(if single {
                    key.text().to_string()
                } else {
                    format!("{} {label}", key.text())
                });
            }
        }
        for label in &value_labels {
            table_headers.push(if single {
                TOTAL.to_string()
            } else {
                format!("{TOTAL} {label}")
            });
        }
    } else {
        table_headers.extend(value_labels.iter().cloned());
    }

    let table = Table {
        column_field: column.map(|index| headers[index].clone()),
        sheet: sheet_name,
        row_fields,
        value_labels,
        column_keys: columns.into_iter().map(|(key, _)| key).collect(),
        headers: table_headers,
        rows,
        totals,
    };
    Ok((table, sheet_names))
}

/// Split one `AGG:FIELD` entry at its first colon.
fn parse_value(entry: &str) -> Result<(Agg, &str)> {
    let (agg, field) = entry
        .split_once(':')
        .filter(|(agg, field)| !agg.trim().is_empty() && !field.is_empty())
        .ok_or_else(|| {
            anyhow!(
                "malformed --values entry {entry:?}; expected AGG:FIELD, for example sum:Amount"
            )
        })?;
    let agg = Agg::parse(agg.trim()).ok_or_else(|| {
        anyhow!("unknown aggregation {agg:?} in {entry:?}; expected sum, count, avg, min, or max")
    })?;
    Ok((agg, field))
}

/// Exact header match, then a unique case-insensitive one.
fn resolve(headers: &[String], name: &str) -> Result<usize> {
    match positions(headers, |header| header == name).as_slice() {
        [only] => return Ok(*only),
        [] => {}
        many => bail!(
            "header {name:?} appears {} times; rename the duplicate columns",
            many.len()
        ),
    }
    let folded = name.to_lowercase();
    match positions(headers, |header| header.to_lowercase() == folded).as_slice() {
        [only] => Ok(*only),
        [] => bail!("unknown header {name:?}; headers: {}", headers.join(", ")),
        many => {
            let names: Vec<String> = many
                .iter()
                .map(|&index| format!("{:?}", headers[index]))
                .collect();
            bail!(
                "header {name:?} matches {} without case; use the exact header",
                names.join(", ")
            )
        }
    }
}

fn positions(headers: &[String], matches: impl Fn(&str) -> bool) -> Vec<usize> {
    headers
        .iter()
        .enumerate()
        .filter(|(_, header)| matches(header.as_str()))
        .map(|(index, _)| index)
        .collect()
}

fn aggregates<'a>(
    values: &'a [(Agg, usize)],
    accs: &'a [Acc],
) -> impl Iterator<Item = Option<f64>> + 'a {
    values
        .iter()
        .zip(accs)
        .map(|(&(agg, _), acc)| acc.result(agg))
}

fn is_blank(value: &CellValue) -> bool {
    match value {
        CellValue::Empty => true,
        CellValue::String(text) => text.trim().is_empty(),
        _ => false,
    }
}

/// Excel displays at most 15 significant digits.
fn rounded(number: f64) -> f64 {
    format!("{number:.14e}").parse().unwrap_or(number)
}

fn text_cell(text: &str) -> Cell {
    Cell {
        value: CellValue::String(text.to_string()),
        number_format: None,
    }
}

/// Integral results become `Int` so the renderers print them without a fraction.
fn value_cell(value: Option<f64>) -> Cell {
    let value = match value {
        None => CellValue::Empty,
        Some(number) if number.fract() == 0.0 && number.abs() < MAX_EXACT_INTEGER => {
            CellValue::Int(number as i64)
        }
        Some(number) => CellValue::Float(number),
    };
    Cell {
        value,
        number_format: None,
    }
}

fn json_number(number: f64) -> Value {
    if number.fract() == 0.0 && number.abs() < MAX_EXACT_INTEGER {
        Value::from(number as i64)
    } else {
        serde_json::Number::from_f64(number).map_or(Value::Null, Value::Number)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn numeric_keys_sort_by_value_and_blanks_sort_last() {
        let mut keys = [
            Key::number(100.0),
            Key::Blank,
            Key::number(9.0),
            Key::number(10.0),
        ];
        let numeric = all_numeric(keys.iter());
        keys.sort_by(|left, right| compare_keys(left, right, numeric));
        let texts: Vec<&str> = keys.iter().map(Key::text).collect();
        assert_eq!(texts, ["9", "10", "100", "(blank)"]);
    }

    #[test]
    fn mixed_keys_sort_by_display_text() {
        let mut keys = [Key::number(9.0), Key::Text("b".into()), Key::number(10.0)];
        let numeric = all_numeric(keys.iter());
        keys.sort_by(|left, right| compare_keys(left, right, numeric));
        let texts: Vec<&str> = keys.iter().map(Key::text).collect();
        assert_eq!(texts, ["10", "9", "b"]);
    }

    #[test]
    fn values_entries_parse_or_refuse() {
        assert!(matches!(parse_value("SUM:a:b"), Ok((Agg::Sum, "a:b"))));
        assert!(parse_value("Amount").is_err());
        assert!(parse_value("sum:").is_err());
        assert!(parse_value("median:Amount").is_err());
    }
}
