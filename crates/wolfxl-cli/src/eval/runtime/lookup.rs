//! `MATCH`, `VLOOKUP`, `HLOOKUP`, `INDEX`, `OFFSET`, and `XLOOKUP`.

use std::cmp::Ordering;

use super::criteria::Pattern;
use super::{condition, supplied, Context, Grid, Operand, Outcome, Range};
use crate::eval::{DataSource, ExcelError, Expr, Value, MAX_COLUMN, MAX_ROW};

/// What a lookup accepts as a match.
enum Mode {
    /// Equal values; text ignores case and may use wildcards.
    Exact { wildcards: bool },
    /// The last value not above the lookup value, in ascending data.
    Ascending,
    /// The last value not below the lookup value, in descending data.
    Descending,
    /// An exact match, else the largest smaller value.
    NextSmaller,
    /// An exact match, else the smallest larger value.
    NextLarger,
}

/// A one-row or one-column view of a grid.
struct Line<'g, 'a> {
    grid: &'g Grid<'a>,
    vertical: bool,
    /// Positions that can hold values.
    live: u32,
}

/// Compares values of the same kind; values of different kinds never match.
/// Blank only reaches here from XLOOKUP, where it matches blank cells.
fn same_kind(candidate: &Value, lookup: &Value) -> Option<Ordering> {
    match (candidate, lookup) {
        (Value::Number(candidate), Value::Number(lookup)) => candidate.partial_cmp(lookup),
        (Value::Text(candidate), Value::Text(lookup)) => {
            Some(candidate.to_lowercase().cmp(&lookup.to_lowercase()))
        }
        (Value::Boolean(candidate), Value::Boolean(lookup)) => Some(candidate.cmp(lookup)),
        (Value::Blank, Value::Blank) => Some(Ordering::Equal),
        _ => None,
    }
}

/// A lookup value: errors propagate and blank never matches.
fn lookup_value(value: Value) -> Result<Value, ExcelError> {
    match value {
        Value::Error(error) => Err(error),
        Value::Blank => Err(ExcelError::Na),
        value => Ok(value),
    }
}

fn integer(value: f64) -> f64 {
    value.trunc()
}

impl<'a, S: DataSource> Context<'a, S> {
    fn line<'g>(&self, grid: &'g Grid<'a>) -> Option<Line<'g, 'a>> {
        let (rows, columns) = grid.size();
        let (live_rows, live_columns) = self.live(grid);
        if columns == 1 {
            Some(Line {
                grid,
                vertical: true,
                live: live_rows,
            })
        } else if rows == 1 {
            Some(Line {
                grid,
                vertical: false,
                live: live_columns,
            })
        } else {
            None
        }
    }

    fn line_at(&self, line: &Line<'_, '_>, position: u32) -> Value {
        if line.vertical {
            self.at(line.grid, position, 0)
        } else {
            self.at(line.grid, 0, position)
        }
    }

    /// The zero-based position of `lookup` among `count` values read by
    /// `get`, searched from the last value first when `reverse` is set.
    fn search(
        &self,
        lookup: &Value,
        count: u32,
        get: impl Fn(u32) -> Value,
        mode: &Mode,
        reverse: bool,
    ) -> Option<u32> {
        let positions: Box<dyn Iterator<Item = u32>> = if reverse {
            Box::new((0..count).rev())
        } else {
            Box::new(0..count)
        };
        let pattern = match (mode, lookup) {
            (Mode::Exact { wildcards: true }, Value::Text(text)) => Some(Pattern::new(text)),
            _ => None,
        };
        let mut found: Option<(u32, Value)> = None;
        for position in positions {
            let candidate = get(position);
            if let (Some(pattern), Value::Text(text)) = (&pattern, &candidate) {
                if pattern.matches(text) {
                    return Some(position);
                }
                continue;
            }
            let Some(ordering) = same_kind(&candidate, lookup) else {
                continue;
            };
            match mode {
                Mode::Exact { .. } if ordering == Ordering::Equal => return Some(position),
                Mode::Exact { .. } => {}
                Mode::Ascending if ordering == Ordering::Greater => break,
                Mode::Descending if ordering == Ordering::Less => break,
                Mode::Ascending | Mode::Descending => found = Some((position, candidate)),
                Mode::NextSmaller | Mode::NextLarger if ordering == Ordering::Equal => {
                    return Some(position)
                }
                Mode::NextSmaller | Mode::NextLarger => {
                    let wanted = if matches!(mode, Mode::NextSmaller) {
                        Ordering::Less
                    } else {
                        Ordering::Greater
                    };
                    let better = match &found {
                        None => true,
                        // Closer than the best so far; ties keep the first found.
                        Some((_, best)) => same_kind(&candidate, best) == Some(wanted.reverse()),
                    };
                    if ordering == wanted && better {
                        found = Some((position, candidate));
                    }
                }
            }
        }
        found.map(|(position, _)| position)
    }

    pub(super) fn match_position(&self, arguments: &[Expr]) -> Outcome {
        let lookup = lookup_value(self.scalar(&arguments[0]))?;
        let grid = self.grid(&arguments[1])?;
        let kind = self.number_or(arguments, 2, 1.0)?;
        let mode = if kind > 0.0 {
            Mode::Ascending
        } else if kind < 0.0 {
            Mode::Descending
        } else {
            Mode::Exact { wildcards: true }
        };
        let line = self.line(&grid).ok_or(ExcelError::Na)?;
        self.search(
            &lookup,
            line.live,
            |position| self.line_at(&line, position),
            &mode,
            false,
        )
        .map(|position| Value::Number(f64::from(position + 1)))
        .ok_or(ExcelError::Na)
    }

    /// `VLOOKUP`, or `HLOOKUP` when not `vertical`.
    pub(super) fn table_lookup(&self, arguments: &[Expr], vertical: bool) -> Outcome {
        let lookup = lookup_value(self.scalar(&arguments[0]))?;
        let table = self.grid(&arguments[1])?;
        let index = integer(self.number(&arguments[2])?);
        let approximate = match arguments.get(3) {
            Some(argument) => condition(self.scalar(argument))?,
            None => true,
        };
        let (rows, columns) = table.size();
        let (live_rows, live_columns) = self.live(&table);
        let (extent, count) = if vertical {
            (columns, live_rows)
        } else {
            (rows, live_columns)
        };
        if index < 1.0 {
            return Err(ExcelError::Value);
        }
        if index > f64::from(extent) {
            return Err(ExcelError::Ref);
        }
        let mode = if approximate {
            Mode::Ascending
        } else {
            Mode::Exact { wildcards: true }
        };
        let get = |position| {
            if vertical {
                self.at(&table, position, 0)
            } else {
                self.at(&table, 0, position)
            }
        };
        let position = self
            .search(&lookup, count, get, &mode, false)
            .ok_or(ExcelError::Na)?;
        let offset = index as u32 - 1;
        let value = if vertical {
            self.at(&table, position, offset)
        } else {
            self.at(&table, offset, position)
        };
        Ok(match value {
            Value::Blank => Value::Number(0.0),
            value => value,
        })
    }

    pub(super) fn index(&self, arguments: &[Expr]) -> Operand<'a> {
        match self.index_operand(arguments) {
            Ok(operand) => self.result(operand),
            Err(error) => Operand::Value(Value::Error(error)),
        }
    }

    fn index_operand(&self, arguments: &[Expr]) -> Result<Operand<'a>, ExcelError> {
        let grid = self.grid(&arguments[0])?;
        let (rows, columns) = grid.size();
        let first = integer(self.number(&arguments[1])?);
        let (row, column) = match arguments.get(2) {
            Some(argument) => (first, integer(self.number(argument)?)),
            // One number indexes along a single row or column, or selects a row.
            None if rows == 1 => (1.0, first),
            None if columns == 1 => (first, 1.0),
            None => (first, 0.0),
        };
        if row < 0.0 || column < 0.0 {
            return Err(ExcelError::Value);
        }
        if row > f64::from(rows) || column > f64::from(columns) {
            return Err(ExcelError::Ref);
        }
        let (row, column) = (row as u32, column as u32);
        // Zero selects every row or every column.
        let (first_row, height) = if row == 0 { (0, rows) } else { (row - 1, 1) };
        let (first_column, width) = if column == 0 {
            (0, columns)
        } else {
            (column - 1, 1)
        };
        Ok(match grid {
            Grid::Range(range) => Operand::Range(Range {
                first_row: range.first_row + first_row,
                first_column: range.first_column + first_column,
                rows: height,
                columns: width,
                single: height == 1 && width == 1,
                whole: range.whole && (height > 1 || width > 1),
                ..range
            }),
            Grid::Array(values) => Operand::Value(Value::Array(
                values
                    .into_iter()
                    .skip(first_row as usize)
                    .take(height as usize)
                    .map(|values| {
                        values
                            .into_iter()
                            .skip(first_column as usize)
                            .take(width as usize)
                            .collect()
                    })
                    .collect(),
            )),
        })
    }

    pub(super) fn offset(&self, arguments: &[Expr]) -> Operand<'a> {
        match self.offset_range(arguments) {
            Ok(range) => self.result(Operand::Range(range)),
            Err(error) => Operand::Value(Value::Error(error)),
        }
    }

    fn offset_range(&self, arguments: &[Expr]) -> Result<Range<'a>, ExcelError> {
        let base = match self.operand(&arguments[0]) {
            Operand::Range(range) => range,
            Operand::Value(Value::Error(error)) => return Err(error),
            Operand::Value(_) => return Err(ExcelError::Value),
        };
        let rows = integer(self.number(&arguments[1])?);
        let columns = integer(self.number(&arguments[2])?);
        let height = integer(self.number_or_default(arguments, 3, f64::from(base.rows))?);
        let width = integer(self.number_or_default(arguments, 4, f64::from(base.columns))?);
        if height < 1.0 || width < 1.0 {
            return Err(ExcelError::Ref);
        }
        let first_row = f64::from(base.first_row) + rows;
        let first_column = f64::from(base.first_column) + columns;
        if first_row < 1.0
            || first_column < 1.0
            || first_row + height - 1.0 > f64::from(MAX_ROW)
            || first_column + width - 1.0 > f64::from(MAX_COLUMN)
        {
            return Err(ExcelError::Ref);
        }
        let (height, width) = (height as u32, width as u32);
        Ok(Range {
            first_row: first_row as u32,
            first_column: first_column as u32,
            rows: height,
            columns: width,
            single: height == 1 && width == 1,
            whole: base.whole && height == base.rows && width == base.columns,
            ..base
        })
    }

    pub(super) fn xlookup(&self, arguments: &[Expr]) -> Operand<'a> {
        match self.xlookup_operand(arguments) {
            Ok(Some(operand)) => self.result(operand),
            Ok(None) => match supplied(arguments, 3) {
                Some(argument) => self.result(self.operand(argument)),
                None => Operand::Value(Value::Error(ExcelError::Na)),
            },
            Err(error) => Operand::Value(Value::Error(error)),
        }
    }

    /// The matching slice of the return array, or `None` when nothing matches.
    fn xlookup_operand(&self, arguments: &[Expr]) -> Result<Option<Operand<'a>>, ExcelError> {
        let lookup = self.scalar(&arguments[0]);
        if let Value::Error(error) = lookup {
            return Err(error);
        }
        let lookup_grid = self.grid(&arguments[1])?;
        let returned = self.grid(&arguments[2])?;
        let mode = match self.number_or_default(arguments, 4, 0.0)? {
            0.0 => Mode::Exact { wildcards: false },
            -1.0 => Mode::NextSmaller,
            1.0 => Mode::NextLarger,
            2.0 => Mode::Exact { wildcards: true },
            _ => return Err(ExcelError::Value),
        };
        // Binary search modes give the same answer as a linear scan on the
        // sorted data they require.
        let reverse = match self.number_or_default(arguments, 5, 1.0)? {
            1.0 | 2.0 => false,
            -1.0 | -2.0 => true,
            _ => return Err(ExcelError::Value),
        };
        let line = self.line(&lookup_grid).ok_or(ExcelError::Value)?;
        let (lookup_rows, lookup_columns) = lookup_grid.size();
        let (return_rows, return_columns) = returned.size();
        if (line.vertical && return_rows != lookup_rows)
            || (!line.vertical && return_columns != lookup_columns)
        {
            return Err(ExcelError::Value);
        }
        // Cells past the used area are blank, so a blank lookup value also
        // matches the first of them in search order.
        let total = if line.vertical {
            lookup_rows
        } else {
            lookup_columns
        };
        let trailing_blank = matches!(lookup, Value::Blank) && total > line.live;
        let found = if trailing_blank && reverse {
            Some(total - 1)
        } else {
            self.search(
                &lookup,
                line.live,
                |position| self.line_at(&line, position),
                &mode,
                reverse,
            )
            .or(trailing_blank.then_some(line.live))
        };
        let Some(position) = found else {
            return Ok(None);
        };
        Ok(Some(match returned {
            Grid::Range(range) if line.vertical => Operand::Range(Range {
                first_row: range.first_row + position,
                rows: 1,
                single: range.columns == 1,
                whole: false,
                ..range
            }),
            Grid::Range(range) => Operand::Range(Range {
                first_column: range.first_column + position,
                columns: 1,
                single: range.rows == 1,
                whole: false,
                ..range
            }),
            Grid::Array(values) if line.vertical => Operand::Value(Value::Array(vec![values
                .into_iter()
                .nth(position as usize)
                .unwrap_or_default()])),
            Grid::Array(values) => Operand::Value(Value::Array(
                values
                    .into_iter()
                    .map(|row| {
                        vec![row
                            .into_iter()
                            .nth(position as usize)
                            .unwrap_or(Value::Blank)]
                    })
                    .collect(),
            )),
        }))
    }
}
