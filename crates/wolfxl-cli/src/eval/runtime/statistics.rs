//! `MEDIAN`, `MODE`, `STDEV`, `VAR`, `PERCENTILE`, `QUARTILE`, `RANK`,
//! `LARGE`, `SMALL`, `COUNTBLANK`, `SUMPRODUCT`, and `SUBTOTAL`.
//!
//! The statistical functions read their data the way `SUM` does: ranges and
//! arrays contribute their numbers and skip text, booleans, and blanks, while
//! a direct argument counts booleans and numeric text and makes other text
//! `#VALUE!`. Errors propagate.

use std::collections::HashMap;

use super::{finite, number, Context, Grid, Operand, Outcome, Range};
use crate::eval::{DataSource, ExcelError, Expr, Function, Value};

/// Keep a range or array value when it is a number.
fn referenced(numbers: &mut Vec<f64>, value: Value) -> Result<(), ExcelError> {
    match value {
        Value::Number(number) => numbers.push(number),
        Value::Error(error) => return Err(error),
        _ => {}
    }
    Ok(())
}

/// The variance of `values` about their mean, dividing by `n - 1` for a
/// sample and by `n` for a population; `#DIV/0!` without enough values.
fn variance(values: &[f64], sample: bool) -> Result<f64, ExcelError> {
    let count = values.len() as f64;
    let divisor = if sample { count - 1.0 } else { count };
    if divisor <= 0.0 {
        return Err(ExcelError::Div0);
    }
    let mean = values.iter().sum::<f64>() / count;
    let squares: f64 = values.iter().map(|value| (value - mean).powi(2)).sum();
    Ok(squares / divisor)
}

fn sorted(mut values: Vec<f64>) -> Vec<f64> {
    values.sort_unstable_by(f64::total_cmp);
    values
}

/// The value at a zero-based fractional `rank` of sorted `values`,
/// interpolating linearly between neighbours.
fn interpolate(values: &[f64], rank: f64) -> f64 {
    let low = rank.floor() as usize;
    let fraction = rank - rank.floor();
    match values.get(low + 1) {
        Some(high) if fraction > 0.0 => values[low] + fraction * (high - values[low]),
        _ => values[low],
    }
}

/// `PERCENTILE.INC`: `k` from 0 through 1 over every value.
fn percentile_inclusive(values: &[f64], k: f64) -> Result<f64, ExcelError> {
    if values.is_empty() || !(0.0..=1.0).contains(&k) {
        return Err(ExcelError::Num);
    }
    Ok(interpolate(values, k * (values.len() - 1) as f64))
}

/// `PERCENTILE.EXC`: `k` strictly between 0 and 1, and close enough to the
/// middle that its rank among `n + 1` slots falls within the values.
fn percentile_exclusive(values: &[f64], k: f64) -> Result<f64, ExcelError> {
    let count = values.len() as f64;
    let rank = k * (count + 1.0);
    if values.is_empty() || k <= 0.0 || k >= 1.0 || rank < 1.0 || rank > count {
        return Err(ExcelError::Num);
    }
    Ok(interpolate(values, rank - 1.0))
}

/// The most frequent value, the earliest in data order on a tie, or `#N/A`
/// when no value repeats.
fn mode(values: &[f64]) -> Result<f64, ExcelError> {
    // Adding zero turns -0 into 0, so both count as one value.
    let mut counts: HashMap<u64, (usize, usize)> = HashMap::new();
    for (position, value) in values.iter().enumerate() {
        counts
            .entry((value + 0.0).to_bits())
            .or_insert((0, position))
            .0 += 1;
    }
    counts
        .into_values()
        .filter(|&(count, _)| count > 1)
        .max_by(|left, right| left.0.cmp(&right.0).then(right.1.cmp(&left.1)))
        .map(|(_, position)| values[position])
        .ok_or(ExcelError::Na)
}

impl<'a, S: DataSource> Context<'a, S> {
    /// Every number the statistical functions read from `arguments`.
    pub(super) fn sample(&self, arguments: &[Expr]) -> Result<Vec<f64>, ExcelError> {
        let mut numbers = Vec::new();
        for argument in arguments {
            match self.operand(argument) {
                Operand::Range(range) => {
                    self.values(&Grid::Range(range), |value| referenced(&mut numbers, value))?
                }
                Operand::Value(Value::Array(rows)) => {
                    for value in rows.into_iter().flatten() {
                        referenced(&mut numbers, value)?;
                    }
                }
                Operand::Value(Value::Blank) => {}
                Operand::Value(value) => numbers.push(number(value)?),
            }
        }
        Ok(numbers)
    }

    pub(super) fn median(&self, arguments: &[Expr]) -> Outcome {
        let values = sorted(self.sample(arguments)?);
        if values.is_empty() {
            return Err(ExcelError::Num);
        }
        Ok(finite(interpolate(
            &values,
            (values.len() - 1) as f64 / 2.0,
        )))
    }

    pub(super) fn mode(&self, arguments: &[Expr]) -> Outcome {
        mode(&self.sample(arguments)?).map(Value::Number)
    }

    /// `STDEV`, `STDEV.S`, `STDEV.P`, `STDEVP`, `VAR`, `VAR.S`, `VAR.P`, and
    /// `VARP`.
    pub(super) fn dispersion(&self, function: Function, arguments: &[Expr]) -> Outcome {
        use Function::*;
        let values = self.sample(arguments)?;
        let (sample, root) = match function {
            StDev | StDevDotS => (true, true),
            StDevP | StDevDotP => (false, true),
            Var | VarDotS => (true, false),
            VarP | VarDotP => (false, false),
            _ => unreachable!("not a dispersion function"),
        };
        let variance = variance(&values, sample)?;
        Ok(finite(if root { variance.sqrt() } else { variance }))
    }

    /// `PERCENTILE`, `PERCENTILE.INC`, and `PERCENTILE.EXC`.
    pub(super) fn percentile(&self, arguments: &[Expr], exclusive: bool) -> Outcome {
        let values = sorted(self.sample(&arguments[..1])?);
        let k = self.number(&arguments[1])?;
        let result = if exclusive {
            percentile_exclusive(&values, k)
        } else {
            percentile_inclusive(&values, k)
        };
        result.map(finite)
    }

    /// `QUARTILE` and `QUARTILE.INC` take quarters 0 through 4, and
    /// `QUARTILE.EXC` 1 through 3; the quarter is truncated.
    pub(super) fn quartile(&self, arguments: &[Expr], exclusive: bool) -> Outcome {
        let values = sorted(self.sample(&arguments[..1])?);
        let quarter = self.number(&arguments[1])?.trunc();
        let result = if exclusive {
            if !(1.0..=3.0).contains(&quarter) {
                return Err(ExcelError::Num);
            }
            percentile_exclusive(&values, quarter / 4.0)
        } else {
            if !(0.0..=4.0).contains(&quarter) {
                return Err(ExcelError::Num);
            }
            percentile_inclusive(&values, quarter / 4.0)
        };
        result.map(finite)
    }

    /// `LARGE`, or `SMALL` when not `largest`. A fractional `k` rounds up.
    pub(super) fn kth(&self, arguments: &[Expr], largest: bool) -> Outcome {
        let values = sorted(self.sample(&arguments[..1])?);
        let k = self.number(&arguments[1])?.ceil();
        if k < 1.0 || k > values.len() as f64 {
            return Err(ExcelError::Num);
        }
        let index = k as usize - 1;
        Ok(Value::Number(if largest {
            values[values.len() - 1 - index]
        } else {
            values[index]
        }))
    }

    /// `RANK` and `RANK.EQ`, or `RANK.AVG` when `average`: the position of
    /// the number among the reference's numbers, largest first unless the
    /// order is nonzero, and `#N/A` when the number is not among them.
    pub(super) fn rank(&self, arguments: &[Expr], average: bool) -> Outcome {
        let target = self.number(&arguments[0])?;
        let values = self.sample(&arguments[1..2])?;
        let ascending = self.number_or(arguments, 2, 0.0)? != 0.0;
        let mut ahead = 0_u64;
        let mut ties = 0_u64;
        for value in values {
            if value == target {
                ties += 1;
            } else if (value < target) == ascending {
                ahead += 1;
            }
        }
        if ties == 0 {
            return Err(ExcelError::Na);
        }
        let first = (ahead + 1) as f64;
        Ok(Value::Number(if average {
            first + (ties - 1) as f64 / 2.0
        } else {
            first
        }))
    }

    /// Blank cells and empty text in a range, cells past the used extent
    /// included.
    pub(super) fn count_blank(&self, arguments: &[Expr]) -> Outcome {
        let grid = self.grid(&arguments[0])?;
        let (rows, columns) = grid.size();
        let mut filled = 0_u64;
        self.values(&grid, |value| {
            if !matches!(&value, Value::Blank) && value != Value::Text(String::new()) {
                filled += 1;
            }
            Ok(())
        })?;
        Ok(Value::Number(
            (u64::from(rows) * u64::from(columns) - filled) as f64,
        ))
    }

    /// The sum of the products of same-position values. Every argument must
    /// have the same size, and values other than numbers count as zero.
    pub(super) fn sum_product(&self, arguments: &[Expr]) -> Outcome {
        let grids = arguments
            .iter()
            .map(|argument| self.grid(argument))
            .collect::<Result<Vec<_>, _>>()?;
        let size = grids[0].size();
        if grids.iter().any(|grid| grid.size() != size) {
            return Err(ExcelError::Value);
        }
        // Past every grid's used extent all values are blank, so zero.
        let (rows, columns) = grids.iter().fold((0, 0), |(rows, columns), grid| {
            let (live_rows, live_columns) = self.live(grid);
            (rows.max(live_rows), columns.max(live_columns))
        });
        let mut total = 0.0;
        for row in 0..rows {
            for column in 0..columns {
                let mut product = 1.0;
                for grid in &grids {
                    match self.at(grid, row, column) {
                        Value::Number(number) => product *= number,
                        Value::Error(error) => return Err(error),
                        _ => product = 0.0,
                    }
                }
                total += product;
            }
        }
        Ok(finite(total))
    }

    /// `SUBTOTAL(function_num, ref, ...)`: function numbers 1 through 11 are
    /// `AVERAGE`, `COUNT`, `COUNTA`, `MAX`, `MIN`, `PRODUCT`, `STDEV`,
    /// `STDEVP`, `SUM`, `VAR`, and `VARP`; 101 through 111 are the same but
    /// skip hidden rows. Cells that hold `SUBTOTAL` formulas are skipped so
    /// nested subtotals do not count twice. Every other argument must be a
    /// reference.
    pub(super) fn subtotal(&self, arguments: &[Expr]) -> Outcome {
        let code = self.number(&arguments[0])?.trunc();
        let (code, skip_hidden) = if (1.0..=11.0).contains(&code) {
            (code as u8, false)
        } else if (101.0..=111.0).contains(&code) {
            (code as u8 - 100, true)
        } else {
            return Err(ExcelError::Value);
        };
        let mut numbers = Vec::new();
        let mut filled = 0_u64;
        let mut error = None;
        for argument in &arguments[1..] {
            let range = match self.operand(argument) {
                Operand::Range(range) => range,
                Operand::Value(Value::Error(error)) => return Err(error),
                Operand::Value(_) => return Err(ExcelError::Value),
            };
            self.subtotal_cells(&range, skip_hidden, |value| {
                if matches!(value, Value::Blank) {
                    return;
                }
                filled += 1;
                match value {
                    Value::Number(number) => numbers.push(number),
                    Value::Error(found) => {
                        error.get_or_insert(found);
                    }
                    _ => {}
                }
            });
        }
        match (code, error) {
            (2, _) => return Ok(Value::Number(numbers.len() as f64)),
            (3, _) => return Ok(Value::Number(filled as f64)),
            (_, Some(error)) => return Err(error),
            _ => {}
        }
        let count = numbers.len() as f64;
        let result = match code {
            1 if numbers.is_empty() => return Err(ExcelError::Div0),
            1 => numbers.iter().sum::<f64>() / count,
            4..=6 if numbers.is_empty() => 0.0,
            4 => numbers.iter().copied().fold(f64::NEG_INFINITY, f64::max),
            5 => numbers.iter().copied().fold(f64::INFINITY, f64::min),
            6 => numbers.iter().product(),
            7 => variance(&numbers, true)?.sqrt(),
            8 => variance(&numbers, false)?.sqrt(),
            9 => numbers.iter().sum(),
            10 => variance(&numbers, true)?,
            _ => variance(&numbers, false)?,
        };
        Ok(finite(result))
    }

    /// Visit the cells `SUBTOTAL` reads in `range`.
    fn subtotal_cells(&self, range: &Range<'_>, skip_hidden: bool, mut visit: impl FnMut(Value)) {
        let (rows, columns) = self.live_range(range);
        for row in range.first_row..range.first_row + rows {
            if skip_hidden && self.source.row_hidden(range.sheet, row) {
                continue;
            }
            for column in range.first_column..range.first_column + columns {
                if !self.source.is_subtotal(range.sheet, row, column) {
                    visit(self.cell(range.sheet, row, column));
                }
            }
        }
    }
}
