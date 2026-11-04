//! `NA` and the `IS` functions `ISBLANK`, `ISERROR`, `ISNA`, `ISNUMBER`, and
//! `ISTEXT`. Each `IS` function tests the value's own type without coercing
//! it, so `ISNUMBER("1")` and `ISBLANK("")` are false, and an error argument
//! is a value to test rather than a result to propagate. A range or array
//! argument yields an array of the same shape, tested element by element.

use super::{Context, Operand, Outcome};
use crate::eval::{DataSource, ExcelError, Expr, Function, Value};

impl<'a, S: DataSource> Context<'a, S> {
    pub(super) fn information(&self, function: Function, arguments: &[Expr]) -> Outcome {
        let test: fn(&Value) -> bool = match function {
            Function::Na => return Err(ExcelError::Na),
            Function::IsBlank => |value| matches!(value, Value::Blank),
            Function::IsError => |value| matches!(value, Value::Error(_)),
            Function::IsNa => |value| matches!(value, Value::Error(ExcelError::Na)),
            Function::IsNumber => |value| matches!(value, Value::Number(_)),
            Function::IsText => |value| matches!(value, Value::Text(_)),
            _ => unreachable!("{function:?} is not an information function"),
        };
        let flag = |value: &Value| Value::Boolean(test(value));
        // Test cells as they are read instead of materializing the range, with
        // the shape `read` gives it: a single cell is a scalar and whole rows or
        // columns are cut at the used extent.
        Ok(match self.operand(&arguments[0]) {
            Operand::Range(range) if range.single => {
                flag(&self.cell(range.sheet, range.first_row, range.first_column))
            }
            Operand::Range(range) => {
                let (rows, columns) = if range.whole {
                    self.live_range(&range)
                } else {
                    (range.rows, range.columns)
                };
                if rows == 0 || columns == 0 {
                    return Ok(Value::Array(Vec::new()));
                }
                Value::Array(
                    (0..rows)
                        .map(|row| {
                            (0..columns)
                                .map(|column| {
                                    flag(&self.cell(
                                        range.sheet,
                                        range.first_row + row,
                                        range.first_column + column,
                                    ))
                                })
                                .collect()
                        })
                        .collect(),
                )
            }
            Operand::Value(Value::Array(mut rows)) => {
                for row in &mut rows {
                    for value in row {
                        *value = flag(value);
                    }
                }
                Value::Array(rows)
            }
            Operand::Value(value) => flag(&value),
        })
    }
}
