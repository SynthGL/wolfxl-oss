//! Trigonometric, logarithmic, and rounding functions, and `PRODUCT`.
//!
//! Like `ABS` and `SQRT`, every function except `PRODUCT` reads scalar
//! arguments: a direct boolean or numeric text converts, other text is
//! `#VALUE!`, errors propagate, and a result that overflows is `#NUM!`.

use std::f64::consts::PI;

use super::{finite, number, round_decimal, rounded, Context, Grid, Operand, Outcome, Rounding};
use crate::eval::{DataSource, ExcelError, Expr, Function, Value};

/// `SIN`, `COS`, and `TAN` return `#NUM!` from this magnitude on, as in
/// Excel.
const TRIGONOMETRIC_LIMIT: f64 = 134_217_728.0;

/// `+ 0.0` turns a negative zero into zero, so `CEILING(-0.5, 1)` is `0`.
fn unsigned_zero(number: f64) -> f64 {
    number + 0.0
}

/// The count of `significance` steps in `number`, taken as an integer when
/// its 15-significant-digit form is one: `0.3 / 0.1` is
/// `2.9999999999999996` in binary but three steps for Excel.
fn steps(number: f64, significance: f64) -> f64 {
    let quotient = number / significance;
    let nearest = quotient.round();
    if rounded(quotient) == nearest {
        nearest
    } else {
        quotient
    }
}

impl<S: DataSource> Context<'_, S> {
    pub(super) fn math(&self, function: Function, arguments: &[Expr]) -> Outcome {
        use Function::*;
        match function {
            Acos => self.unary(arguments, |value| {
                if value.abs() > 1.0 {
                    Value::Error(ExcelError::Num)
                } else {
                    finite(value.acos())
                }
            }),
            Asin => self.unary(arguments, |value| {
                if value.abs() > 1.0 {
                    Value::Error(ExcelError::Num)
                } else {
                    finite(value.asin())
                }
            }),
            Atan => self.unary(arguments, |value| finite(value.atan())),
            Atan2 => self.atan2(arguments),
            Cos => self.trigonometric(arguments, f64::cos),
            Sin => self.trigonometric(arguments, f64::sin),
            Tan => self.trigonometric(arguments, f64::tan),
            Degrees => self.unary(arguments, |value| finite(value.to_degrees())),
            Radians => self.unary(arguments, |value| finite(value.to_radians())),
            Pi => Ok(Value::Number(PI)),
            Exp => self.unary(arguments, |value| finite(value.exp())),
            Ln => self.unary(arguments, |value| {
                if value <= 0.0 {
                    Value::Error(ExcelError::Num)
                } else {
                    finite(value.ln())
                }
            }),
            Log => self.log(arguments),
            Log10 => self.unary(arguments, |value| {
                if value <= 0.0 {
                    Value::Error(ExcelError::Num)
                } else {
                    finite(value.log10())
                }
            }),
            Ceiling => self.ceiling(arguments),
            Floor => self.floor(arguments),
            Trunc => self.trunc(arguments),
            Product => self.product(arguments),
            _ => unreachable!("{} is not a math function", function.name()),
        }
    }

    /// A function of the one numeric argument.
    fn unary(&self, arguments: &[Expr], apply: impl FnOnce(f64) -> Value) -> Outcome {
        self.number(&arguments[0]).map(apply)
    }

    /// `SIN`, `COS`, or `TAN` of an angle in radians.
    fn trigonometric(&self, arguments: &[Expr], apply: fn(f64) -> f64) -> Outcome {
        self.unary(arguments, |angle| {
            if angle.abs() >= TRIGONOMETRIC_LIMIT {
                Value::Error(ExcelError::Num)
            } else {
                finite(apply(angle))
            }
        })
    }

    /// `ATAN2(x_num, y_num)`: the angle of the point `(x, y)` from the
    /// x-axis, in `(-pi, pi]`; `#DIV/0!` at the origin.
    fn atan2(&self, arguments: &[Expr]) -> Outcome {
        let x = self.number(&arguments[0])?;
        let y = self.number(&arguments[1])?;
        if x == 0.0 && y == 0.0 {
            return Err(ExcelError::Div0);
        }
        // Excel normalizes a literal -0 to 0. A nonzero negative y may
        // still round to -PI, so preserve that native boundary result.
        Ok(finite(unsigned_zero(y).atan2(x)))
    }

    /// `LOG(number, [base])`, base 10 when omitted. A number or base that
    /// is not positive is `#NUM!`, and base 1 is `#DIV/0!`.
    fn log(&self, arguments: &[Expr]) -> Outcome {
        let number = self.number(&arguments[0])?;
        let base = self.number_or(arguments, 1, 10.0)?;
        if number <= 0.0 || base <= 0.0 {
            return Err(ExcelError::Num);
        }
        Ok(if base == 1.0 {
            Value::Error(ExcelError::Div0)
        } else if base == 10.0 {
            finite(number.log10())
        } else if base == 2.0 {
            finite(number.log2())
        } else {
            finite(number.ln() / base.ln())
        })
    }

    /// Legacy `CEILING(number, significance)`: the nearest multiple of
    /// `significance` toward positive infinity, or away from zero when both
    /// are negative. Significance 0 gives 0, and a positive number with a
    /// negative significance is `#NUM!`.
    fn ceiling(&self, arguments: &[Expr]) -> Outcome {
        let number = self.number(&arguments[0])?;
        let significance = self.number(&arguments[1])?;
        if significance == 0.0 {
            return Ok(Value::Number(0.0));
        }
        if number > 0.0 && significance < 0.0 {
            return Err(ExcelError::Num);
        }
        let multiple = steps(number, significance).ceil() * significance;
        Ok(finite(unsigned_zero(rounded(multiple))))
    }

    /// Legacy `FLOOR(number, significance)`: the nearest multiple of
    /// `significance` toward negative infinity, or toward zero when both
    /// are negative. Significance 0 is `#DIV/0!`, and a positive number with
    /// a negative significance is `#NUM!`.
    fn floor(&self, arguments: &[Expr]) -> Outcome {
        let number = self.number(&arguments[0])?;
        let significance = self.number(&arguments[1])?;
        if number == 0.0 {
            return Ok(Value::Number(0.0));
        }
        if significance == 0.0 {
            return Err(ExcelError::Div0);
        }
        if number > 0.0 && significance < 0.0 {
            return Err(ExcelError::Num);
        }
        let multiple = steps(number, significance).floor() * significance;
        Ok(finite(unsigned_zero(rounded(multiple))))
    }

    /// `TRUNC(number, [digits])`: `number` cut toward zero at `digits`
    /// decimal places, 0 when omitted; the same as `ROUNDDOWN`.
    fn trunc(&self, arguments: &[Expr]) -> Outcome {
        let number = self.number(&arguments[0])?;
        let digits = self.number_or(arguments, 1, 0.0)?;
        Ok(finite(round_decimal(number, digits, Rounding::Down)))
    }

    /// `PRODUCT`: the product of the numbers, read the way `SUM` reads
    /// them. A direct argument counts booleans and numeric text and makes
    /// other text `#VALUE!`; ranges and arrays contribute only their
    /// numbers. Errors propagate, and without any number the result is 0.
    fn product(&self, arguments: &[Expr]) -> Outcome {
        let mut product = 1.0;
        let mut seen = false;
        let mut multiply = |value: Value, referenced: bool| -> Result<(), ExcelError> {
            let factor = match value {
                Value::Number(number) => number,
                Value::Error(error) => return Err(error),
                Value::Blank => return Ok(()),
                Value::Text(_) | Value::Boolean(_) if referenced => return Ok(()),
                value => number(value)?,
            };
            seen = true;
            product *= factor;
            Ok(())
        };
        for argument in arguments {
            match self.operand(argument) {
                // Read references cell by cell rather than materializing them.
                Operand::Range(range) => {
                    self.values(&Grid::Range(range), |value| multiply(value, true))?
                }
                Operand::Value(Value::Array(rows)) => {
                    for value in rows.into_iter().flatten() {
                        multiply(value, true)?;
                    }
                }
                Operand::Value(value) => multiply(value, false)?,
            }
        }
        Ok(if seen {
            finite(product)
        } else {
            Value::Number(0.0)
        })
    }
}
