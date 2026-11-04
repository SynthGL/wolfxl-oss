//! `PMT`, `PV`, `FV`, `NPV`, `IRR`, `SLN`, and `DB`.

use super::{finite, number, round_decimal, Context, Grid, Operand, Outcome, Rounding};
use crate::eval::{DataSource, ExcelError, Expr, Value};

/// Payments at the start of each period when `type` is nonzero.
fn due(value: f64) -> f64 {
    if value == 0.0 {
        0.0
    } else {
        1.0
    }
}

/// The present value of `values` discounted at `rate`, the first value one
/// period out when `first` is 1 and immediate when it is 0.
fn discounted(rate: f64, values: &[f64], first: i32) -> f64 {
    values
        .iter()
        .zip(first..)
        .map(|(value, period)| value / (1.0 + rate).powi(period))
        .sum()
}

impl<'a, S: DataSource> Context<'a, S> {
    /// Cash flows: every number in ranges and arrays, which skip text,
    /// booleans, and blanks, and every direct argument as a number.
    fn cash_flows(&self, arguments: &[Expr]) -> Result<Vec<f64>, ExcelError> {
        let mut values = Vec::new();
        let mut keep = |value: Value| match value {
            Value::Number(number) => {
                values.push(number);
                Ok(())
            }
            Value::Error(error) => Err(error),
            _ => Ok(()),
        };
        for argument in arguments {
            match self.operand(argument) {
                Operand::Range(range) => self.values(&Grid::Range(range), &mut keep)?,
                Operand::Value(Value::Array(rows)) => {
                    rows.into_iter().flatten().try_for_each(&mut keep)?
                }
                Operand::Value(value) => keep(Value::Number(number(value)?))?,
            }
        }
        Ok(values)
    }

    /// Rate, periods, and the payment or present value, then the optional
    /// future or present value and payment timing.
    fn annuity(&self, arguments: &[Expr]) -> Result<(f64, f64, f64, f64, f64), ExcelError> {
        Ok((
            self.number(&arguments[0])?,
            self.number(&arguments[1])?,
            self.number(&arguments[2])?,
            self.number_or(arguments, 3, 0.0)?,
            due(self.number_or(arguments, 4, 0.0)?),
        ))
    }

    pub(super) fn pmt(&self, arguments: &[Expr]) -> Outcome {
        let (rate, periods, present, future, timing) = self.annuity(arguments)?;
        if periods == 0.0 {
            return Err(ExcelError::Num);
        }
        if rate == 0.0 {
            return Ok(finite(-(present + future) / periods));
        }
        let growth = (1.0 + rate).powf(periods);
        Ok(finite(
            -(rate * (future + present * growth)) / ((1.0 + rate * timing) * (growth - 1.0)),
        ))
    }

    pub(super) fn pv(&self, arguments: &[Expr]) -> Outcome {
        let (rate, periods, payment, future, timing) = self.annuity(arguments)?;
        if rate == 0.0 {
            return Ok(finite(-(future + payment * periods)));
        }
        let growth = (1.0 + rate).powf(periods);
        Ok(finite(
            -(future + payment * (1.0 + rate * timing) * (growth - 1.0) / rate) / growth,
        ))
    }

    pub(super) fn fv(&self, arguments: &[Expr]) -> Outcome {
        let (rate, periods, payment, present, timing) = self.annuity(arguments)?;
        if rate == 0.0 {
            return Ok(finite(-(present + payment * periods)));
        }
        let growth = (1.0 + rate).powf(periods);
        Ok(finite(
            -(present * growth + payment * (1.0 + rate * timing) * (growth - 1.0) / rate),
        ))
    }

    pub(super) fn npv(&self, arguments: &[Expr]) -> Outcome {
        let rate = self.number(&arguments[0])?;
        if rate == -1.0 {
            return Err(ExcelError::Div0);
        }
        let values = self.cash_flows(&arguments[1..])?;
        Ok(finite(discounted(rate, &values, 1)))
    }

    /// Newton's method from the guess, then bisection on [-0.999, 10].
    pub(super) fn irr(&self, arguments: &[Expr]) -> Outcome {
        let values = self.cash_flows(&arguments[..1])?;
        let guess = self.number_or(arguments, 1, 0.1)?;
        if values.len() < 2
            || !values.iter().any(|&value| value > 0.0)
            || !values.iter().any(|&value| value < 0.0)
        {
            return Err(ExcelError::Num);
        }
        let npv = |rate: f64| discounted(rate, &values, 0);
        let slope = |rate: f64| -> f64 {
            values
                .iter()
                .zip(0..)
                .map(|(value, period)| -f64::from(period) * value / (1.0 + rate).powi(period + 1))
                .sum()
        };
        let mut rate = guess;
        for _ in 0..100 {
            let value = npv(rate);
            if value.abs() < 1e-10 {
                return Ok(finite(rate));
            }
            let derivative = slope(rate);
            if derivative.abs() < 1e-14 || !derivative.is_finite() {
                break;
            }
            let next = rate - value / derivative;
            if (next - rate).abs() < 1e-10 {
                return Ok(finite(next));
            }
            rate = next;
        }
        let (mut low, mut high) = (-0.999, 10.0);
        if npv(low) * npv(high) > 0.0 {
            return Err(ExcelError::Num);
        }
        for _ in 0..200 {
            let middle = (low + high) / 2.0;
            if npv(middle).abs() < 1e-10 || high - low < 1e-12 {
                return Ok(finite(middle));
            }
            if npv(low) * npv(middle) < 0.0 {
                high = middle;
            } else {
                low = middle;
            }
        }
        Err(ExcelError::Num)
    }

    pub(super) fn sln(&self, arguments: &[Expr]) -> Outcome {
        let cost = self.number(&arguments[0])?;
        let salvage = self.number(&arguments[1])?;
        let life = self.number(&arguments[2])?;
        if life == 0.0 {
            return Err(ExcelError::Div0);
        }
        Ok(finite((cost - salvage) / life))
    }

    /// Fixed-declining balance with the rate rounded to three places; the
    /// first year covers `month` months and a final partial year the rest.
    pub(super) fn db(&self, arguments: &[Expr]) -> Outcome {
        let cost = self.number(&arguments[0])?;
        let salvage = self.number(&arguments[1])?;
        let life = self.number(&arguments[2])?;
        let period = self.number(&arguments[3])?.trunc();
        let month = self.number_or(arguments, 4, 12.0)?.trunc();
        if cost < 0.0
            || salvage < 0.0
            || life <= 0.0
            || period <= 0.0
            || !(1.0..=12.0).contains(&month)
            || period > life + 1.0
            || (month == 12.0 && period > life)
        {
            return Err(ExcelError::Num);
        }
        if cost == 0.0 {
            return Ok(Value::Number(0.0));
        }
        let rate = round_decimal(
            1.0 - (salvage / cost).powf(1.0 / life),
            3.0,
            Rounding::Nearest,
        );
        let mut total = cost * rate * month / 12.0;
        if period == 1.0 {
            return Ok(finite(total));
        }
        let mut current = 2.0;
        loop {
            let depreciation = if current > life {
                (cost - total) * rate * (12.0 - month) / 12.0
            } else {
                (cost - total) * rate
            };
            if current == period {
                return Ok(finite(depreciation));
            }
            total += depreciation;
            current += 1.0;
        }
    }
}
