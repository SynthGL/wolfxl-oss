use std::cmp::Ordering;

use crate::filter::lenient_number;
use crate::formula_data::rounded;

use super::{
    Area, DataSource, ExcelError, Expr, Function, Operator, Reference, Value, MAX_COLUMN, MAX_ROW,
};

struct Context<'a, S> {
    source: &'a S,
    current_sheet: &'a str,
}

pub(super) fn evaluate(expression: &Expr, source: &impl DataSource, current_sheet: &str) -> Value {
    Context {
        source,
        current_sheet,
    }
    .evaluate(expression)
}

impl<S: DataSource> Context<'_, S> {
    fn evaluate(&self, expression: &Expr) -> Value {
        match expression {
            Expr::Literal(value) => value.clone(),
            Expr::Reference(reference) => match self.resolve(reference) {
                Ok(area) if area.single => self.cell(area.sheet, area.first_row, area.first_column),
                Ok(area) => {
                    let mut rows = Vec::new();
                    if !area.is_empty() {
                        for row in area.first_row..=area.last_row {
                            let values = (area.first_column..=area.last_column)
                                .map(|column| self.cell(area.sheet, row, column))
                                .collect();
                            rows.push(values);
                        }
                    }
                    Value::Array(rows)
                }
                Err(error) => Value::Error(error),
            },
            Expr::Unary { negative, operand } => match number(self.evaluate(operand)) {
                Ok(value) => finite(if *negative { -value } else { value }),
                Err(error) => Value::Error(error),
            },
            Expr::Percent(operand) => match number(self.evaluate(operand)) {
                Ok(value) => finite(value / 100.0),
                Err(error) => Value::Error(error),
            },
            Expr::Binary {
                operator,
                left,
                right,
            } => {
                let left = self.evaluate(left);
                if matches!(left, Value::Error(_)) {
                    return left;
                }
                let right = self.evaluate(right);
                if matches!(right, Value::Error(_)) {
                    return right;
                }
                binary(*operator, left, right)
            }
            Expr::Call {
                function: Function::If,
                arguments,
            } => match condition(self.evaluate(&arguments[0])) {
                Ok(selected) => {
                    let index = if selected { 1 } else { 2 };
                    let Some(branch) = arguments.get(index) else {
                        return Value::Boolean(false);
                    };
                    match self.evaluate(branch) {
                        Value::Blank => Value::Number(0.0),
                        value => value,
                    }
                }
                Err(error) => Value::Error(error),
            },
            Expr::Call {
                function,
                arguments,
            } => {
                let mut aggregate = Aggregate::new(*function);
                for argument in arguments {
                    // Read reference arguments directly rather than materializing large ranges.
                    let result = if let Expr::Reference(reference) = argument {
                        match self.resolve(reference) {
                            Ok(area) => self.add_area(&mut aggregate, &area),
                            Err(error) => aggregate.add(Value::Error(error), true),
                        }
                    } else {
                        aggregate.add(self.evaluate(argument), false)
                    };
                    if let Err(error) = result {
                        return Value::Error(error);
                    }
                }
                aggregate.finish()
            }
        }
    }

    fn add_area(&self, aggregate: &mut Aggregate, area: &Resolved<'_>) -> Result<(), ExcelError> {
        if !area.is_empty() {
            for row in area.first_row..=area.last_row {
                for column in area.first_column..=area.last_column {
                    aggregate.add(self.cell(area.sheet, row, column), true)?;
                }
            }
        }
        Ok(())
    }

    fn cell(&self, sheet: &str, row: u32, column: u32) -> Value {
        match self.source.cell_value(sheet, row, column) {
            Value::Number(number) => finite(number),
            Value::Array(_) => Value::Error(ExcelError::Value),
            value => value,
        }
    }

    fn resolve(&self, reference: &Reference) -> Result<Resolved<'_>, ExcelError> {
        let requested = reference.sheet.as_deref().unwrap_or(self.current_sheet);
        let sheet = self
            .source
            .sheet_names()
            .find(|name| name.eq_ignore_ascii_case(requested))
            .ok_or(ExcelError::Ref)?;
        let mut area = Resolved {
            sheet,
            first_row: 1,
            last_row: 0,
            first_column: 1,
            last_column: 0,
            single: matches!(reference.area, Area::Cell { .. }),
        };
        match reference.area {
            Area::Cell { row, column } => {
                area.first_row = row;
                area.last_row = row;
                area.first_column = column;
                area.last_column = column;
            }
            Area::Cells {
                first_row,
                last_row,
                first_column,
                last_column,
            } => {
                area.first_row = first_row;
                area.last_row = last_row;
                area.first_column = first_column;
                area.last_column = last_column;
            }
            Area::Columns { first, last } => {
                let extent = self.source.used_extent(sheet);
                area.last_row = extent.rows.min(MAX_ROW);
                area.first_column = first;
                area.last_column = last.min(extent.columns).min(MAX_COLUMN);
            }
            Area::Rows { first, last } => {
                let extent = self.source.used_extent(sheet);
                area.first_row = first;
                area.last_row = last.min(extent.rows).min(MAX_ROW);
                area.last_column = extent.columns.min(MAX_COLUMN);
            }
        }
        Ok(area)
    }
}

struct Resolved<'a> {
    sheet: &'a str,
    first_row: u32,
    last_row: u32,
    first_column: u32,
    last_column: u32,
    single: bool,
}

impl Resolved<'_> {
    fn is_empty(&self) -> bool {
        self.first_row > self.last_row || self.first_column > self.last_column
    }
}

fn finite(number: f64) -> Value {
    if number.is_finite() {
        Value::Number(number)
    } else {
        Value::Error(ExcelError::Num)
    }
}

fn number(value: Value) -> Result<f64, ExcelError> {
    match value {
        Value::Number(number) => Ok(number),
        Value::Boolean(value) => Ok(if value { 1.0 } else { 0.0 }),
        Value::Blank => Ok(0.0),
        Value::Text(text) => lenient_number(&text).ok_or(ExcelError::Value),
        Value::Error(error) => Err(error),
        Value::Array(_) => Err(ExcelError::Value),
    }
}

fn condition(value: Value) -> Result<bool, ExcelError> {
    match value {
        Value::Boolean(value) => Ok(value),
        Value::Number(value) => Ok(value != 0.0),
        Value::Blank => Ok(false),
        Value::Error(error) => Err(error),
        Value::Text(_) | Value::Array(_) => Err(ExcelError::Value),
    }
}

fn text(value: Value) -> Result<String, ExcelError> {
    match value {
        Value::Text(text) => Ok(text),
        Value::Number(number) => Ok({
            let number = rounded(number);
            if number == 0.0 {
                "0".to_owned()
            } else {
                number.to_string()
            }
        }),
        Value::Boolean(value) => Ok(if value { "TRUE" } else { "FALSE" }.to_owned()),
        Value::Blank => Ok(String::new()),
        Value::Error(error) => Err(error),
        Value::Array(_) => Err(ExcelError::Value),
    }
}

fn binary(operator: Operator, left: Value, right: Value) -> Value {
    use Operator::*;
    match operator {
        Equal | NotEqual | Less | Greater | LessEqual | GreaterEqual => {
            match compare(&left, &right) {
                Ok(ordering) => Value::Boolean(match operator {
                    Equal => ordering == Ordering::Equal,
                    NotEqual => ordering != Ordering::Equal,
                    Less => ordering == Ordering::Less,
                    Greater => ordering == Ordering::Greater,
                    LessEqual => ordering != Ordering::Greater,
                    GreaterEqual => ordering != Ordering::Less,
                    _ => unreachable!(),
                }),
                Err(error) => Value::Error(error),
            }
        }
        Concat => match text(left).and_then(|mut left| {
            left.push_str(&text(right)?);
            Ok(left)
        }) {
            Ok(text) => Value::Text(text),
            Err(error) => Value::Error(error),
        },
        _ => match number(left).and_then(|left| number(right).map(|right| (left, right))) {
            Ok((left, right)) => match operator {
                Add => finite(left + right),
                Subtract => finite(left - right),
                Multiply => finite(left * right),
                Divide if right == 0.0 => Value::Error(ExcelError::Div0),
                Divide => finite(left / right),
                Power if left == 0.0 && right < 0.0 => Value::Error(ExcelError::Div0),
                Power if left == 0.0 && right == 0.0 => Value::Error(ExcelError::Num),
                Power => finite(left.powf(right)),
                _ => unreachable!(),
            },
            Err(error) => Value::Error(error),
        },
    }
}

fn compare(left: &Value, right: &Value) -> Result<Ordering, ExcelError> {
    use Value::*;
    let ordering = match (left, right) {
        (Error(error), _) | (_, Error(error)) => return Err(*error),
        (Array(_), _) | (_, Array(_)) => return Err(ExcelError::Value),
        (Blank, Blank) => Ordering::Equal,
        (Blank, Number(right)) => 0.0_f64.partial_cmp(right).ok_or(ExcelError::Num)?,
        (Number(left), Blank) => left.partial_cmp(&0.0).ok_or(ExcelError::Num)?,
        (Blank, Text(right)) => "".cmp(right.as_str()),
        (Text(left), Blank) => left.as_str().cmp(""),
        (Blank, Boolean(right)) => false.cmp(right),
        (Boolean(left), Blank) => left.cmp(&false),
        (Number(left), Number(right)) => left.partial_cmp(right).ok_or(ExcelError::Num)?,
        (Text(left), Text(right)) => left.to_lowercase().cmp(&right.to_lowercase()),
        (Boolean(left), Boolean(right)) => left.cmp(right),
        (Number(_), Text(_) | Boolean(_)) | (Text(_), Boolean(_)) => Ordering::Less,
        (Text(_) | Boolean(_), Number(_)) | (Boolean(_), Text(_)) => Ordering::Greater,
    };
    Ok(ordering)
}

struct Aggregate {
    function: Function,
    count: u64,
    sum: f64,
    minimum: f64,
    maximum: f64,
}

impl Aggregate {
    fn new(function: Function) -> Self {
        Self {
            function,
            count: 0,
            sum: 0.0,
            minimum: f64::INFINITY,
            maximum: f64::NEG_INFINITY,
        }
    }

    fn add(&mut self, value: Value, referenced: bool) -> Result<(), ExcelError> {
        let value = match value {
            Value::Error(_) if matches!(self.function, Function::Count) => return Ok(()),
            Value::Error(error) => return Err(error),
            Value::Array(rows) => {
                for row in rows {
                    for value in row {
                        self.add(value, true)?;
                    }
                }
                return Ok(());
            }
            Value::Number(value) => value,
            Value::Blank => return Ok(()),
            Value::Text(_) | Value::Boolean(_) if referenced => return Ok(()),
            Value::Text(text) if matches!(self.function, Function::Count) => {
                let Some(value) = lenient_number(&text) else {
                    return Ok(());
                };
                value
            }
            value => number(value)?,
        };
        self.count += 1;
        // Do not compute unrelated statistics for COUNT, MIN, or MAX.
        match self.function {
            Function::Sum | Function::Average => self.sum += value,
            Function::Min => self.minimum = self.minimum.min(value),
            Function::Max => self.maximum = self.maximum.max(value),
            Function::Count => {}
            Function::If => unreachable!(),
        }
        Ok(())
    }

    fn finish(self) -> Value {
        match self.function {
            Function::Count => Value::Number(self.count as f64),
            Function::Sum => finite(self.sum),
            Function::Average if self.count == 0 => Value::Error(ExcelError::Div0),
            Function::Average => finite(self.sum / self.count as f64),
            Function::Min | Function::Max if self.count == 0 => Value::Number(0.0),
            Function::Min => finite(self.minimum),
            Function::Max => finite(self.maximum),
            Function::If => unreachable!(),
        }
    }
}
