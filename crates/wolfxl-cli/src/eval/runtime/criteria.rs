//! `COUNTIF`, `SUMIF`, `AVERAGEIF`, their multi-criteria forms, `MINIFS`,
//! and `MAXIFS`, plus the wildcard patterns that lookups share.

use std::cmp::Ordering;

use super::{finite, lenient_number, Context, Grid, Outcome, Range, Statistic};
use crate::eval::{DataSource, ExcelError, Expr, Function, Value, MAX_COLUMN, MAX_ROW};

/// A case-insensitive pattern where `*` matches any run of characters, `?`
/// any one character, and `~` makes the next character literal.
pub(super) struct Pattern {
    tokens: Vec<Glob>,
}

#[derive(Clone, Copy, PartialEq)]
enum Glob {
    Literal(char),
    One,
    Any,
}

impl Pattern {
    pub(super) fn new(pattern: &str) -> Self {
        let mut tokens = Vec::new();
        let mut characters = pattern.chars().flat_map(char::to_lowercase);
        while let Some(character) = characters.next() {
            tokens.push(match character {
                '*' => Glob::Any,
                '?' => Glob::One,
                '~' => match characters.next() {
                    Some(next) => Glob::Literal(next),
                    None => Glob::Literal('~'),
                },
                other => Glob::Literal(other),
            });
        }
        Self { tokens }
    }

    pub(super) fn matches(&self, text: &str) -> bool {
        let text: Vec<char> = text.chars().flat_map(char::to_lowercase).collect();
        let (mut position, mut token) = (0, 0);
        // After a `*`, the token following it and the text position it last
        // tried, so a failed match can let the `*` absorb one more character.
        let mut retry: Option<(usize, usize)> = None;
        while position < text.len() {
            match self.tokens.get(token) {
                Some(Glob::Any) => {
                    token += 1;
                    retry = Some((token, position));
                }
                Some(Glob::One) => {
                    token += 1;
                    position += 1;
                }
                Some(Glob::Literal(literal)) if *literal == text[position] => {
                    token += 1;
                    position += 1;
                }
                _ => match retry {
                    Some((after, tried)) => {
                        token = after;
                        position = tried + 1;
                        retry = Some((after, tried + 1));
                    }
                    None => return false,
                },
            }
        }
        self.tokens[token..].iter().all(|token| *token == Glob::Any)
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum Comparison {
    Equal,
    NotEqual,
    Less,
    Greater,
    LessEqual,
    GreaterEqual,
}

enum Target {
    Number(f64),
    Boolean(bool),
    Error(ExcelError),
    /// Matched by blank cells and empty text.
    Empty,
    Text(String, Pattern),
}

/// One worksheet criterion such as `5`, `">=10"`, `"<>"`, or `"North*"`.
struct Criterion {
    comparison: Comparison,
    target: Target,
}

impl Criterion {
    fn new(value: Value) -> Self {
        let text = match value {
            Value::Number(number) => return Self::equal(Target::Number(number)),
            Value::Boolean(value) => return Self::equal(Target::Boolean(value)),
            Value::Error(error) => return Self::equal(Target::Error(error)),
            // An empty criteria cell counts as zero.
            Value::Blank => return Self::equal(Target::Number(0.0)),
            Value::Array(_) => return Self::equal(Target::Error(ExcelError::Value)),
            Value::Text(text) => text,
        };
        let (comparison, operand) = [
            ("<=", Comparison::LessEqual),
            (">=", Comparison::GreaterEqual),
            ("<>", Comparison::NotEqual),
            ("<", Comparison::Less),
            (">", Comparison::Greater),
            ("=", Comparison::Equal),
        ]
        .into_iter()
        .find_map(|(prefix, comparison)| {
            text.strip_prefix(prefix)
                .map(|operand| (comparison, operand))
        })
        .unwrap_or((Comparison::Equal, text.as_str()));
        let target = if operand.is_empty() {
            Target::Empty
        } else if let Some(number) = lenient_number(operand) {
            Target::Number(number)
        } else if operand.eq_ignore_ascii_case("TRUE") {
            Target::Boolean(true)
        } else if operand.eq_ignore_ascii_case("FALSE") {
            Target::Boolean(false)
        } else if let Some(error) = ExcelError::parse(operand) {
            Target::Error(error)
        } else {
            Target::Text(operand.to_lowercase(), Pattern::new(operand))
        };
        Self { comparison, target }
    }

    fn equal(target: Target) -> Self {
        Self {
            comparison: Comparison::Equal,
            target,
        }
    }

    fn matches(&self, value: &Value) -> bool {
        let ordering = match self.comparison {
            Comparison::Equal => return self.equals(value),
            Comparison::NotEqual => return !self.equals(value),
            _ => match (&self.target, value) {
                (Target::Number(target), Value::Number(value)) => value.partial_cmp(target),
                (Target::Boolean(target), Value::Boolean(value)) => Some(value.cmp(target)),
                (Target::Text(target, _), Value::Text(value)) => {
                    Some(value.to_lowercase().cmp(target))
                }
                (Target::Empty, Value::Text(value)) => Some(value.as_str().cmp("")),
                _ => None,
            },
        };
        let Some(ordering) = ordering else {
            return false;
        };
        match self.comparison {
            Comparison::Less => ordering == Ordering::Less,
            Comparison::Greater => ordering == Ordering::Greater,
            Comparison::LessEqual => ordering != Ordering::Greater,
            Comparison::GreaterEqual => ordering != Ordering::Less,
            Comparison::Equal | Comparison::NotEqual => unreachable!(),
        }
    }

    fn equals(&self, value: &Value) -> bool {
        match (&self.target, value) {
            (Target::Number(target), Value::Number(value)) => value == target,
            (Target::Number(target), Value::Text(value)) => lenient_number(value) == Some(*target),
            (Target::Boolean(target), Value::Boolean(value)) => value == target,
            (Target::Error(target), Value::Error(value)) => value == target,
            (Target::Empty, Value::Blank) => true,
            (Target::Empty, Value::Text(value)) => value.is_empty(),
            (Target::Text(_, pattern), Value::Text(value)) => pattern.matches(value),
            _ => false,
        }
    }
}

/// A grid of the given size anchored at the same top-left cell, as `SUMIF`
/// reads its sum range.
fn resize(grid: Grid<'_>, rows: u32, columns: u32) -> Grid<'_> {
    match grid {
        Grid::Range(range) => Grid::Range(Range {
            rows: rows.min(MAX_ROW - range.first_row + 1),
            columns: columns.min(MAX_COLUMN - range.first_column + 1),
            single: false,
            ..range
        }),
        array => array,
    }
}

impl<'a, S: DataSource> Context<'a, S> {
    pub(super) fn count_if(&self, arguments: &[Expr]) -> Outcome {
        let range = self.grid(&arguments[0])?;
        let criterion = Criterion::new(self.scalar(&arguments[1]));
        self.tally(Statistic::Count, None, vec![(range, criterion)])
    }

    /// `SUMIF` or, with `average`, `AVERAGEIF`.
    pub(super) fn sum_if(&self, arguments: &[Expr], average: bool) -> Outcome {
        let range = self.grid(&arguments[0])?;
        let criterion = Criterion::new(self.scalar(&arguments[1]));
        let target = match arguments.get(2) {
            Some(argument) => {
                let (rows, columns) = range.size();
                resize(self.grid(argument)?, rows, columns)
            }
            None => range.clone(),
        };
        let statistic = if average {
            Statistic::Average
        } else {
            Statistic::Sum
        };
        self.tally(statistic, Some(target), vec![(range, criterion)])
    }

    /// `COUNTIFS`, `SUMIFS`, `AVERAGEIFS`, `MINIFS`, and `MAXIFS`.
    pub(super) fn conditional(&self, function: Function, arguments: &[Expr]) -> Outcome {
        let (statistic, target, pairs) = match function {
            Function::CountIfs => (Statistic::Count, None, arguments),
            Function::SumIfs => (Statistic::Sum, Some(&arguments[0]), &arguments[1..]),
            Function::AverageIfs => (Statistic::Average, Some(&arguments[0]), &arguments[1..]),
            Function::MinIfs => (Statistic::Min, Some(&arguments[0]), &arguments[1..]),
            Function::MaxIfs => (Statistic::Max, Some(&arguments[0]), &arguments[1..]),
            _ => unreachable!("not a conditional aggregate"),
        };
        let target = target.map(|target| self.grid(target)).transpose()?;
        let mut conditions = Vec::with_capacity(pairs.len() / 2);
        for [range, criteria] in pairs.as_chunks::<2>().0 {
            conditions.push((self.grid(range)?, Criterion::new(self.scalar(criteria))));
        }
        self.tally(statistic, target, conditions)
    }

    /// Aggregate `target` (or count) over the positions where every criterion
    /// matches its grid. All grids must have the same size.
    fn tally(
        &self,
        statistic: Statistic,
        target: Option<Grid<'_>>,
        conditions: Vec<(Grid<'_>, Criterion)>,
    ) -> Outcome {
        let size = conditions[0].0.size();
        let grids = || target.iter().chain(conditions.iter().map(|(grid, _)| grid));
        if grids().any(|grid| grid.size() != size) {
            return Err(ExcelError::Value);
        }
        // Positions past every grid's used extent are blank in all of them.
        let (rows, columns) = grids().fold((0, 0), |(rows, columns), grid| {
            let (live_rows, live_columns) = self.live(grid);
            (rows.max(live_rows), columns.max(live_columns))
        });
        let mut count = 0_u64;
        let mut sum = 0.0;
        let mut minimum = f64::INFINITY;
        let mut maximum = f64::NEG_INFINITY;
        for row in 0..rows {
            for column in 0..columns {
                let matched = conditions
                    .iter()
                    .all(|(grid, criterion)| criterion.matches(&self.at(grid, row, column)));
                if !matched {
                    continue;
                }
                let Some(target) = &target else {
                    count += 1;
                    continue;
                };
                match self.at(target, row, column) {
                    Value::Number(number) => {
                        count += 1;
                        sum += number;
                        minimum = minimum.min(number);
                        maximum = maximum.max(number);
                    }
                    Value::Error(error) => return Err(error),
                    _ => {}
                }
            }
        }
        if target.is_none()
            && conditions
                .iter()
                .all(|(_, criterion)| criterion.matches(&Value::Blank))
        {
            let declared = u64::from(size.0) * u64::from(size.1);
            count += declared - u64::from(rows) * u64::from(columns);
        }
        Ok(match statistic {
            Statistic::Count | Statistic::CountA => Value::Number(count as f64),
            Statistic::Sum => finite(sum),
            Statistic::Average if count == 0 => return Err(ExcelError::Div0),
            Statistic::Average => finite(sum / count as f64),
            Statistic::Min | Statistic::Max if count == 0 => Value::Number(0.0),
            Statistic::Min => finite(minimum),
            Statistic::Max => finite(maximum),
        })
    }
}
