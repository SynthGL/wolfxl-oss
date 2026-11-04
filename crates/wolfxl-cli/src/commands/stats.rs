//! `wolfxl stats FILE|-`: profile every column of one sheet.
//!
//! Row 0 of the selected sheet is the header row and names the columns;
//! every later row is a data row. `--where` and `--select` follow `wolfxl
//! peek`: conditions test the source text of each data row, and the
//! selection profiles the named columns in the order given. Without
//! `--select`, every header-row column is profiled in sheet order.
//!
//! Per column the profile counts blank cells (empty, or text of only
//! whitespace) and non-empty cells, and splits the non-empty cells by type:
//!
//! - type counts read the typed sheet (`Input::sheet`), so a delimited
//!   `1200` is a number and `$1,200` stays text;
//! - the numeric summary covers every cell that `filter::cell_number`
//!   reads, which adds money-style text (`$1,200.00`, `(500)`) to stored
//!   numbers; its standard deviation is the population one;
//! - distinct and top values compare value text: strings in source
//!   spelling, numbers in plain notation with up to 15 significant digits,
//!   booleans as `TRUE`/`FALSE`, dates in ISO form, errors as their Excel
//!   literal. Top values order by count, ties by first appearance.
//!
//! The sheet loads twice, once as source text (headers, filters, value
//! text) and once typed (type counts); for workbooks both loads agree.
//!
//! Exit codes: 0 for a profile, 2 for usage and input failures.

use std::borrow::Cow;
use std::cmp::Reverse;
use std::collections::HashMap;
use std::io::Write;
use std::process::ExitCode;

use anyhow::{bail, Context, Result};
use serde_json::{json, Map, Value};
use wolfxl_core::{Cell, CellValue, Sheet};

use crate::filter::{cell_number, resolve_column, Condition};
use crate::render::{self, RenderOptions};
use crate::{input, StatsArgs, StatsFormat};

const CONTRACT_VERSION: u32 = 1;
/// 2^53: integral values below it are exact and print without a fraction.
const MAX_EXACT_INTEGER: f64 = 9_007_199_254_740_992.0;
/// Type buckets, in output order.
const KINDS: [&str; 5] = ["number", "text", "bool", "date", "error"];
/// Numeric summary fields, in output order.
const SUMMARY: [&str; 5] = ["sum", "mean", "min", "max", "stddev"];
/// Top-level `--export json` keys, in output order.
const JSON_KEYS: [&str; 4] = ["contract_version", "sheet", "rows", "columns"];
/// Column profile keys in output order. Each nested object (`types`,
/// `numeric`, and the `top` entries) lists its keys in this relative order.
const PROFILE_KEYS: [&str; 19] = [
    "header",
    "non_empty",
    "blank",
    "types",
    "number",
    "text",
    "bool",
    "date",
    "error",
    "distinct",
    "numeric",
    "value",
    "count",
    "sum",
    "mean",
    "min",
    "max",
    "stddev",
    "top",
];
/// Box and CSV layout: one row per profiled column.
const TABLE_HEADERS: [&str; 16] = [
    "column",
    "non_empty",
    "blank",
    "number",
    "text",
    "bool",
    "date",
    "error",
    "distinct",
    "numeric",
    "sum",
    "mean",
    "min",
    "max",
    "stddev",
    "top",
];

/// Stand-in for cells past the end of a short row.
static EMPTY: CellValue = CellValue::Empty;

pub(crate) fn run(args: StatsArgs) -> ExitCode {
    let report = match build(&args) {
        Ok(report) => report,
        Err(error) => {
            eprintln!("error: {error:#}");
            return ExitCode::from(2);
        }
    };
    let options = RenderOptions {
        max_rows: None,
        max_columns: None,
        max_width: usize::MAX,
        all_sheets: &report.sheet_names,
    };
    let mut document = Vec::new();
    let rendered = match args.export {
        StatsFormat::Box => render::boxed(&mut document, &report.to_sheet(), &options),
        StatsFormat::Csv => render::csv(&mut document, &report.to_sheet(), &options),
        StatsFormat::Json => report.write_json(&mut document),
    };
    if let Err(error) = rendered {
        eprintln!("error: rendering stats: {error}");
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

/// The `--export json` report for one sheet, built without writing anything.
pub(crate) fn structured(args: &StatsArgs) -> Result<Value> {
    Ok(build(args)?.json())
}

/// Type bucket of a non-blank cell, indexing [`KINDS`].
#[derive(Copy, Clone, Debug)]
enum Kind {
    Number,
    Text,
    Bool,
    Date,
    Error,
}

impl Kind {
    fn of(value: &CellValue) -> Self {
        match value {
            CellValue::Int(_) | CellValue::Float(_) => Kind::Number,
            CellValue::Bool(_) => Kind::Bool,
            CellValue::Date(_) | CellValue::DateTime(_) | CellValue::Time(_) => Kind::Date,
            CellValue::Error(_) => Kind::Error,
            CellValue::String(_) | CellValue::Empty => Kind::Text,
        }
    }
}

/// Running count, sum, extremes, and Welford mean and squared deviations.
#[derive(Copy, Clone, Debug, Default)]
struct Moments {
    count: u64,
    sum: f64,
    min: f64,
    max: f64,
    mean: f64,
    squares: f64,
}

impl Moments {
    fn add(&mut self, number: f64) {
        if self.count == 0 {
            self.min = number;
            self.max = number;
        } else {
            self.min = self.min.min(number);
            self.max = self.max.max(number);
        }
        self.count += 1;
        self.sum += number;
        let delta = number - self.mean;
        self.mean += delta / self.count as f64;
        self.squares += delta * (number - self.mean);
    }

    /// Sum, mean, min, max, and population standard deviation, in
    /// [`SUMMARY`] order, each at 15 significant digits; `None` throughout
    /// without numbers.
    fn summary(&self) -> [Option<f64>; 5] {
        if self.count == 0 {
            return [None; 5];
        }
        let count = self.count as f64;
        [
            self.sum,
            self.sum / count,
            self.min,
            self.max,
            (self.squares / count).max(0.0).sqrt(),
        ]
        .map(|number| Some(rounded(number)))
    }
}

/// Running profile of one column.
#[derive(Debug, Default)]
struct Tally {
    blank: u64,
    kinds: [u64; 5],
    moments: Moments,
    /// Value text to its slot in `counts`; slots follow first appearance.
    slots: HashMap<String, usize>,
    counts: Vec<u64>,
}

impl Tally {
    /// Count one cell from its source-text and typed readings.
    fn add(&mut self, raw: &CellValue, typed: &CellValue) {
        let Some(text) = value_text(raw) else {
            self.blank += 1;
            return;
        };
        self.kinds[Kind::of(typed) as usize] += 1;
        if let Some(number) = cell_number(typed) {
            self.moments.add(number);
        }
        match self.slots.get(&*text).copied() {
            Some(slot) => self.counts[slot] += 1,
            None => {
                self.slots.insert(text.into_owned(), self.counts.len());
                self.counts.push(1);
            }
        }
    }

    /// The finished profile, keeping the `top` most frequent values, or
    /// every value when `top` is 0.
    fn finish(self, header: String, top: usize) -> Profile {
        let distinct = self.counts.len();
        let mut texts: Vec<(usize, String)> = self
            .slots
            .into_iter()
            .map(|(text, slot)| (slot, text))
            .collect();
        texts.sort_unstable_by_key(|&(slot, _)| slot);
        let mut values: Vec<(String, u64)> = texts
            .into_iter()
            .map(|(_, text)| text)
            .zip(self.counts)
            .collect();
        // Stable: equal counts keep first-appearance order.
        values.sort_by_key(|&(_, count)| Reverse(count));
        if top > 0 {
            values.truncate(top);
        }
        Profile {
            header,
            blank: self.blank,
            kinds: self.kinds,
            distinct: distinct as u64,
            numeric: self.moments.count,
            summary: self.moments.summary(),
            top: values,
        }
    }
}

/// The finished profile of one column.
#[derive(Debug)]
struct Profile {
    header: String,
    blank: u64,
    kinds: [u64; 5],
    distinct: u64,
    numeric: u64,
    summary: [Option<f64>; 5],
    top: Vec<(String, u64)>,
}

impl Profile {
    fn non_empty(&self) -> u64 {
        self.kinds.iter().sum()
    }

    /// One table row in [`TABLE_HEADERS`] order.
    fn table_row(&self) -> Vec<Cell> {
        let mut row = Vec::with_capacity(TABLE_HEADERS.len());
        row.push(text_cell(self.header.clone()));
        row.extend(
            [self.non_empty(), self.blank]
                .into_iter()
                .chain(self.kinds)
                .chain([self.distinct, self.numeric])
                .map(|count| text_cell(count.to_string())),
        );
        row.extend(self.summary.iter().map(|number| match number {
            Some(number) => text_cell(number_text(*number)),
            None => Cell::empty(),
        }));
        let top: Vec<String> = self
            .top
            .iter()
            .map(|(text, count)| format!("{text} ({count})"))
            .collect();
        row.push(text_cell(top.join("; ")));
        row
    }

    /// The JSON profile object.
    fn json(&self) -> Value {
        let types: Map<String, Value> = KINDS
            .iter()
            .zip(self.kinds)
            .map(|(kind, count)| ((*kind).to_string(), Value::from(count)))
            .collect();
        let mut numeric = Map::new();
        numeric.insert("count".to_string(), Value::from(self.numeric));
        numeric.extend(SUMMARY.iter().zip(self.summary).map(|(field, number)| {
            (
                (*field).to_string(),
                number.map_or(Value::Null, json_number),
            )
        }));
        let top: Vec<Value> = self
            .top
            .iter()
            .map(|(text, count)| json!({"value": text, "count": count}))
            .collect();
        json!({
            "header": self.header,
            "non_empty": self.non_empty(),
            "blank": self.blank,
            "types": types,
            "distinct": self.distinct,
            "numeric": numeric,
            "top": top,
        })
    }
}

struct Report {
    sheet: String,
    sheet_names: Vec<String>,
    rows: u64,
    columns: Vec<Profile>,
}

impl Report {
    /// Header row, then one row per profiled column.
    fn to_sheet(&self) -> Sheet {
        let mut grid: Vec<Vec<Cell>> = Vec::with_capacity(self.columns.len() + 1);
        grid.push(
            TABLE_HEADERS
                .iter()
                .map(|header| text_cell((*header).to_string()))
                .collect(),
        );
        grid.extend(self.columns.iter().map(Profile::table_row));
        Sheet::from_rows(format!("{} (stats)", self.sheet), grid)
    }

    /// The JSON report object.
    fn json(&self) -> Value {
        let columns: Vec<Value> = self.columns.iter().map(Profile::json).collect();
        json!({
            "contract_version": CONTRACT_VERSION,
            "sheet": self.sheet,
            "rows": self.rows,
            "columns": columns,
        })
    }

    /// Pretty top-level object with one inline column object per line.
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
                Value::Array(columns) if *key == "columns" && !columns.is_empty() => {
                    writeln!(w, "  {name}: [")?;
                    for (index, column) in columns.iter().enumerate() {
                        let separator = if index + 1 == columns.len() { "" } else { "," };
                        writeln!(w, "    {}{separator}", inline(column))?;
                    }
                    writeln!(w, "  ]{comma}")?;
                }
                value => writeln!(w, "  {name}: {}{comma}", inline(value))?,
            }
        }
        writeln!(w, "}}")
    }
}

/// Inline JSON text with `", "` and `": "` separators and object keys in
/// [`PROFILE_KEYS`] order.
fn inline(value: &Value) -> String {
    match value {
        Value::Array(items) => {
            let items: Vec<String> = items.iter().map(inline).collect();
            format!("[{}]", items.join(", "))
        }
        Value::Object(object) => {
            let mut entries: Vec<(&String, &Value)> = object.iter().collect();
            entries.sort_by_key(|(key, _)| {
                PROFILE_KEYS
                    .iter()
                    .position(|known| *known == key.as_str())
                    .unwrap_or(PROFILE_KEYS.len())
            });
            let entries: Vec<String> = entries
                .into_iter()
                .map(|(key, value)| format!("{}: {}", Value::from(key.as_str()), inline(value)))
                .collect();
            format!("{{{}}}", entries.join(", "))
        }
        scalar => scalar.to_string(),
    }
}

fn build(args: &StatsArgs) -> Result<Report> {
    let mut input = input::open(Some(args.file.as_path()), args.from)?;
    let sheet_name = input.sheet_name(args.sheet.as_deref())?;
    let sheet_names = input.sheet_names().to_vec();
    let raw = input.raw_sheet(&sheet_name)?;
    let typed = input.sheet(&sheet_name)?;
    let headers = raw.headers();

    let conditions = args
        .filters
        .iter()
        .map(|expression| {
            let condition = Condition::parse(expression)?;
            let column = resolve_column(&headers, condition.column())
                .with_context(|| format!("--where {expression:?}"))?;
            Ok((column, condition))
        })
        .collect::<Result<Vec<_>>>()?;
    let columns = if args.select.is_empty() {
        (0..headers.len()).collect()
    } else {
        args.select
            .iter()
            .map(|name| {
                let name = name.trim();
                if name.is_empty() {
                    bail!("--select: empty column name; expected HEADER[,HEADER...]");
                }
                resolve_column(&headers, name).context("--select")
            })
            .collect::<Result<Vec<usize>>>()?
    };

    let mut tallies: Vec<Tally> = columns.iter().map(|_| Tally::default()).collect();
    let mut rows = 0;
    for (index, row) in raw.rows().iter().enumerate().skip(1) {
        if !accepts(row, &conditions) {
            continue;
        }
        rows += 1;
        let typed_row = typed.rows().get(index);
        for (tally, &column) in tallies.iter_mut().zip(&columns) {
            let raw_value = row.get(column).map_or(&EMPTY, |cell| &cell.value);
            let typed_value = typed_row
                .and_then(|cells| cells.get(column))
                .map_or(raw_value, |cell| &cell.value);
            tally.add(raw_value, typed_value);
        }
    }

    let columns = tallies
        .into_iter()
        .zip(&columns)
        .map(|(tally, &column)| tally.finish(headers[column].clone(), args.top))
        .collect();
    Ok(Report {
        sheet: sheet_name,
        sheet_names,
        rows,
        columns,
    })
}

/// Whether every `--where` condition accepts the row, as in `peek`.
fn accepts(row: &[Cell], conditions: &[(usize, Condition)]) -> bool {
    conditions.iter().all(|(column, condition)| {
        condition.matches(row.get(*column).map_or(&EMPTY, |cell| &cell.value))
    })
}

/// Value text of a non-blank cell; `None` for empty or whitespace-only cells.
fn value_text(value: &CellValue) -> Option<Cow<'_, str>> {
    let text = match value {
        CellValue::Empty => return None,
        CellValue::String(text) if text.trim().is_empty() => return None,
        CellValue::String(text) | CellValue::Error(text) => Cow::Borrowed(text.as_str()),
        CellValue::Int(integer) => Cow::Owned(integer.to_string()),
        CellValue::Float(float) => Cow::Owned(number_text(*float)),
        CellValue::Bool(true) => Cow::Borrowed("TRUE"),
        CellValue::Bool(false) => Cow::Borrowed("FALSE"),
        CellValue::Date(date) => Cow::Owned(date.format("%Y-%m-%d").to_string()),
        CellValue::DateTime(datetime) => {
            Cow::Owned(datetime.format("%Y-%m-%d %H:%M:%S").to_string())
        }
        CellValue::Time(time) => Cow::Owned(time.format("%H:%M:%S").to_string()),
    };
    Some(text)
}

/// Excel displays at most 15 significant digits.
fn rounded(number: f64) -> f64 {
    format!("{number:.14e}").parse().unwrap_or(number)
}

/// Plain notation at 15 significant digits, with `-0` as `0`.
fn number_text(number: f64) -> String {
    let number = rounded(number);
    if number == 0.0 {
        "0".to_string()
    } else {
        number.to_string()
    }
}

fn json_number(number: f64) -> Value {
    if number.fract() == 0.0 && number.abs() < MAX_EXACT_INTEGER {
        Value::from(number as i64)
    } else {
        serde_json::Number::from_f64(number).map_or(Value::Null, Value::Number)
    }
}

fn text_cell(text: String) -> Cell {
    Cell {
        value: CellValue::String(text),
        number_format: None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn population_deviation_and_mean_print_without_float_noise() {
        let mut moments = Moments::default();
        for number in [2.0, 4.0, 4.0, 4.0, 5.0, 5.0, 7.0, 9.0] {
            moments.add(number);
        }
        assert_eq!(
            moments.summary(),
            [Some(40.0), Some(5.0), Some(2.0), Some(9.0), Some(2.0)]
        );
        let mut tenths = Moments::default();
        for number in [0.1, 0.2] {
            tenths.add(number);
        }
        assert_eq!(tenths.summary()[0], Some(0.3));
        assert_eq!(number_text(-0.0), "0");
    }

    #[test]
    fn top_values_order_by_count_then_first_appearance() {
        let mut tally = Tally::default();
        for text in ["b", "a", "c", "a", "b", " ", "d"] {
            let value = CellValue::String(text.to_string());
            tally.add(&value, &value);
        }
        let profile = tally.finish("h".to_string(), 3);
        assert_eq!(profile.blank, 1);
        assert_eq!(profile.distinct, 4);
        assert_eq!(
            profile.top,
            [
                ("b".to_string(), 2),
                ("a".to_string(), 2),
                ("c".to_string(), 1)
            ]
        );
    }

    #[test]
    fn json_profiles_print_inline_with_keys_in_contract_order() {
        let mut tally = Tally::default();
        for text in ["b", "a", "b", " "] {
            let value = CellValue::String(text.to_string());
            tally.add(&value, &value);
        }
        let profile = tally.finish("h".to_string(), 5);
        assert_eq!(
            inline(&profile.json()),
            concat!(
                r#"{"header": "h", "non_empty": 3, "blank": 1, "#,
                r#""types": {"number": 0, "text": 3, "bool": 0, "date": 0, "error": 0}, "#,
                r#""distinct": 2, "#,
                r#""numeric": {"count": 0, "sum": null, "mean": null, "min": null, "max": null, "stddev": null}, "#,
                r#""top": [{"value": "b", "count": 2}, {"value": "a", "count": 1}]}"#,
            )
        );
    }
}
