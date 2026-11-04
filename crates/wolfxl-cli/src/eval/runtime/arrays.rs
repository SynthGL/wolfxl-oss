//! The dynamic array functions `UNIQUE`, `FILTER`, and `SORT`. Each returns
//! an array; blank cells stay blank in it. Where Excel returns `#CALC!` for
//! an empty result, these return `#VALUE!`, which this error set has.

use std::cmp::Ordering;
use std::collections::HashMap;

use super::{condition, supplied, Context, Grid, Operand, Outcome};
use crate::eval::{DataSource, ExcelError, Expr, Value};

/// A value as `UNIQUE` tells values apart: text ignores case.
#[derive(Clone, PartialEq, Eq, Hash)]
enum Key {
    Number(u64),
    Text(String),
    Boolean(bool),
    Error(&'static str),
    Blank,
}

fn key(value: &Value) -> Key {
    match value {
        // Adding zero turns -0 into 0.
        Value::Number(number) => Key::Number((number + 0.0).to_bits()),
        Value::Text(text) => Key::Text(text.to_lowercase()),
        Value::Boolean(value) => Key::Boolean(*value),
        Value::Error(error) => Key::Error(error.as_str()),
        Value::Blank | Value::Array(_) => Key::Blank,
    }
}

fn transpose(rows: Vec<Vec<Value>>) -> Vec<Vec<Value>> {
    let width = rows.first().map_or(0, Vec::len);
    let mut columns: Vec<Vec<Value>> = (0..width).map(|_| Vec::with_capacity(rows.len())).collect();
    for row in rows {
        for (column, value) in columns.iter_mut().zip(row) {
            column.push(value);
        }
    }
    columns
}

/// Sort position of a value's kind: numbers, text, booleans, errors.
fn kind(value: &Value) -> u8 {
    match value {
        Value::Number(_) => 0,
        Value::Text(_) => 1,
        Value::Boolean(_) => 2,
        _ => 3,
    }
}

/// `SORT`'s order for two non-blank values.
fn sort_order(left: &Value, right: &Value) -> Ordering {
    kind(left)
        .cmp(&kind(right))
        .then_with(|| match (left, right) {
            (Value::Number(left), Value::Number(right)) => left.total_cmp(right),
            (Value::Text(left), Value::Text(right)) => {
                left.to_lowercase().cmp(&right.to_lowercase())
            }
            (Value::Boolean(left), Value::Boolean(right)) => left.cmp(right),
            _ => Ordering::Equal,
        })
}

impl<'a, S: DataSource> Context<'a, S> {
    /// The argument's values as rows. A range reads like a bare reference,
    /// whole rows and columns cut at the used extent; a scalar is one value.
    fn matrix(&self, expression: &Expr) -> Result<Vec<Vec<Value>>, ExcelError> {
        let value = match self.operand(expression) {
            Operand::Range(range) => self.read(Operand::Range(range)),
            Operand::Value(value) => value,
        };
        match value {
            Value::Error(error) => Err(error),
            Value::Array(rows) if rows.iter().all(Vec::is_empty) => Ok(vec![vec![Value::Blank]]),
            Value::Array(rows) => Ok(rows),
            value => Ok(vec![vec![value]]),
        }
    }

    /// An optional boolean argument, false when omitted or empty.
    fn flag(&self, arguments: &[Expr], index: usize) -> Result<bool, ExcelError> {
        supplied(arguments, index).map_or(Ok(false), |argument| condition(self.scalar(argument)))
    }

    /// A grid's rows and columns as an array of its values has them: whole
    /// rows and columns stop at the used extent, as when read bare or
    /// through an operator.
    fn shape(&self, grid: &Grid<'_>) -> (u32, u32) {
        match grid {
            Grid::Range(range) if range.whole => self.live_range(range),
            grid => grid.size(),
        }
    }

    /// The distinct rows, or columns with `by_col`, in first-seen order;
    /// with `exactly_once`, only those that occur once.
    pub(super) fn unique(&self, arguments: &[Expr]) -> Outcome {
        let rows = self.matrix(&arguments[0])?;
        let by_column = self.flag(arguments, 1)?;
        let exactly_once = self.flag(arguments, 2)?;
        let lines = if by_column { transpose(rows) } else { rows };
        let mut seen: HashMap<Vec<Key>, usize> = HashMap::new();
        let mut distinct: Vec<(Vec<Value>, usize)> = Vec::new();
        for line in lines {
            let keys: Vec<Key> = line.iter().map(key).collect();
            match seen.get(&keys) {
                Some(&index) => distinct[index].1 += 1,
                None => {
                    seen.insert(keys, distinct.len());
                    distinct.push((line, 1));
                }
            }
        }
        let kept: Vec<Vec<Value>> = distinct
            .into_iter()
            .filter(|(_, count)| !exactly_once || *count == 1)
            .map(|(line, _)| line)
            .collect();
        if kept.is_empty() {
            return Err(ExcelError::Value);
        }
        Ok(Value::Array(if by_column { transpose(kept) } else { kept }))
    }

    /// The rows (or, for a one-row `include` as wide as the array, the
    /// columns) whose `include` value is true; `if_empty` when none is.
    pub(super) fn filter(&self, arguments: &[Expr]) -> Operand<'a> {
        match self.filtered(arguments) {
            Ok(Some(rows)) => Operand::Value(Value::Array(rows)),
            Ok(None) => match supplied(arguments, 2) {
                Some(argument) => self.result(self.operand(argument)),
                None => Operand::Value(Value::Error(ExcelError::Value)),
            },
            Err(error) => Operand::Value(Value::Error(error)),
        }
    }

    fn filtered(&self, arguments: &[Expr]) -> Result<Option<Vec<Vec<Value>>>, ExcelError> {
        let array = self.grid(&arguments[0])?;
        let include = self.grid(&arguments[1])?;
        let (rows, columns) = self.shape(&array);
        let (include_rows, include_columns) = self.shape(&include);
        let vertical = if include_columns == 1 && include_rows == rows {
            true
        } else if include_rows == 1 && include_columns == columns {
            false
        } else {
            return Err(ExcelError::Value);
        };
        // Past the include grid's used extent every value is blank, so false.
        let (live_rows, live_columns) = self.live(&include);
        let count = if vertical { live_rows } else { live_columns };
        let mut kept = Vec::new();
        for position in 0..count {
            let value = if vertical {
                self.at(&include, position, 0)
            } else {
                self.at(&include, 0, position)
            };
            if !condition(value)? {
                continue;
            }
            kept.push(position);
        }
        if kept.is_empty() {
            return Ok(None);
        }
        Ok(Some(if vertical {
            kept.into_iter()
                .map(|row| {
                    (0..columns)
                        .map(|column| self.at(&array, row, column))
                        .collect()
                })
                .collect()
        } else {
            (0..rows)
                .map(|row| {
                    kept.iter()
                        .map(|&column| self.at(&array, row, column))
                        .collect()
                })
                .collect()
        }))
    }

    /// Rows (or columns with `by_col`) ordered by the values in position
    /// `sort_index`: ascending for 1, descending for -1. Equal keys keep
    /// their order, and blanks sort last either way.
    pub(super) fn sort(&self, arguments: &[Expr]) -> Outcome {
        let rows = self.matrix(&arguments[0])?;
        let index = self.number_or_default(arguments, 1, 1.0)?.trunc();
        let descending = match self.number_or_default(arguments, 2, 1.0)? {
            1.0 => false,
            -1.0 => true,
            _ => return Err(ExcelError::Value),
        };
        let by_column = self.flag(arguments, 3)?;
        let mut lines = if by_column { transpose(rows) } else { rows };
        let width = lines.first().map_or(0, Vec::len);
        if index < 1.0 || index > width as f64 {
            return Err(ExcelError::Value);
        }
        let index = index as usize - 1;
        lines.sort_by(|left, right| match (&left[index], &right[index]) {
            (Value::Blank, Value::Blank) => Ordering::Equal,
            (Value::Blank, _) => Ordering::Greater,
            (_, Value::Blank) => Ordering::Less,
            (left, right) if descending => sort_order(right, left),
            (left, right) => sort_order(left, right),
        });
        Ok(Value::Array(if by_column {
            transpose(lines)
        } else {
            lines
        }))
    }
}

#[cfg(test)]
mod tests {
    use super::transpose;
    use crate::eval::Value;

    #[test]
    fn transpose_swaps_rows_and_columns() {
        let number = Value::Number;
        assert_eq!(
            transpose(vec![
                vec![number(1.0), number(2.0), number(3.0)],
                vec![number(4.0), number(5.0), number(6.0)],
            ]),
            vec![
                vec![number(1.0), number(4.0)],
                vec![number(2.0), number(5.0)],
                vec![number(3.0), number(6.0)],
            ]
        );
    }
}
