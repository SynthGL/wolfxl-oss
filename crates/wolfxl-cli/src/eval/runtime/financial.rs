//! `PMT`, `PV`, `FV`, `NPV`, `IRR`, `XNPV`, `XIRR`, `RATE`, `NPER`, `SLN`,
//! and `DB`.

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

/// The present value of flows `years` after the first date at `rate`.
fn dated_value(rate: f64, values: &[f64], years: &[f64]) -> f64 {
    values
        .iter()
        .zip(years)
        .map(|(value, years)| value / (1.0 + rate).powf(*years))
        .sum()
}

/// A root of `function` above -1: Newton's method from the guess, and when
/// that does not settle, bisection on the sign change between -0.99 and 10
/// nearest the guess.
fn rate_root(function: impl Fn(f64) -> f64, guess: f64) -> Option<f64> {
    let mut rate = guess;
    for _ in 0..100 {
        let value = function(rate);
        if !value.is_finite() {
            break;
        }
        if value.abs() < 1e-10 {
            return Some(rate);
        }
        let step = 1e-7 * rate.abs().max(1.0);
        let slope = (function(rate + step) - function(rate - step)) / (2.0 * step);
        if slope.abs() < 1e-14 || !slope.is_finite() {
            break;
        }
        let next = rate - value / slope;
        if next <= -1.0 || !next.is_finite() {
            break;
        }
        if (next - rate).abs() < 1e-12 {
            return Some(next);
        }
        rate = next;
    }
    // Sample the range and bisect the sign change closest to the guess.
    let points: Vec<f64> = (0..=1099)
        .map(|step| -0.99 + f64::from(step) / 100.0)
        .collect();
    let (mut low, mut high) = points
        .windows(2)
        .filter(|pair| {
            let (left, right) = (function(pair[0]), function(pair[1]));
            left.is_finite() && right.is_finite() && left * right <= 0.0
        })
        .map(|pair| (pair[0], pair[1]))
        .min_by(|left, right| (left.0 - guess).abs().total_cmp(&(right.0 - guess).abs()))?;
    for _ in 0..200 {
        let middle = (low + high) / 2.0;
        if function(middle).abs() < 1e-10 || high - low < 1e-14 {
            return Some(middle);
        }
        if function(low) * function(middle) <= 0.0 {
            high = middle;
        } else {
            low = middle;
        }
    }
    Some((low + high) / 2.0)
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

    /// Every value of an argument as a number for `XNPV` and `XIRR`: any
    /// value that is not a number is `#VALUE!`.
    fn dated_numbers(&self, expression: &Expr) -> Result<Vec<f64>, ExcelError> {
        let value = match self.operand(expression) {
            Operand::Range(range) => self.read(Operand::Range(range)),
            Operand::Value(value) => value,
        };
        let values = match value {
            Value::Array(rows) => rows.into_iter().flatten().collect(),
            value => vec![value],
        };
        values
            .into_iter()
            .map(|value| match value {
                Value::Number(number) => Ok(number),
                Value::Error(error) => Err(error),
                _ => Err(ExcelError::Value),
            })
            .collect()
    }

    /// Cash flows and their dates as years since the first date: the same
    /// count of each, else `#NUM!`, and no date before the first.
    fn dated_flows(&self, arguments: &[Expr]) -> Result<(Vec<f64>, Vec<f64>), ExcelError> {
        let values = self.dated_numbers(&arguments[0])?;
        let dates = self.dated_numbers(&arguments[1])?;
        if values.len() != dates.len() || values.is_empty() {
            return Err(ExcelError::Num);
        }
        let first = dates[0].trunc();
        let mut years = Vec::with_capacity(dates.len());
        for date in dates {
            let date = date.trunc();
            if date < 0.0 || date < first {
                return Err(ExcelError::Num);
            }
            years.push((date - first) / 365.0);
        }
        Ok((values, years))
    }

    /// Present value of cash flows on the given dates, discounted over
    /// 365-day years from the first date.
    pub(super) fn xnpv(&self, arguments: &[Expr]) -> Outcome {
        let rate = self.number(&arguments[0])?;
        if rate <= -1.0 {
            return Err(ExcelError::Num);
        }
        let (values, years) = self.dated_flows(&arguments[1..])?;
        Ok(finite(dated_value(rate, &values, &years)))
    }

    /// The rate at which `XNPV` of the flows is zero; at least one flow must
    /// be positive and one negative.
    pub(super) fn xirr(&self, arguments: &[Expr]) -> Outcome {
        let (values, years) = self.dated_flows(arguments)?;
        let guess = self.number_or_default(arguments, 2, 0.1)?;
        if guess <= -1.0
            || !values.iter().any(|&value| value > 0.0)
            || !values.iter().any(|&value| value < 0.0)
        {
            return Err(ExcelError::Num);
        }
        rate_root(|rate| dated_value(rate, &values, &years), guess)
            .map(finite)
            .ok_or(ExcelError::Num)
    }

    /// The periodic rate at which the annuity balances: periods, payment,
    /// present value, then optional future value, timing, and guess.
    pub(super) fn rate(&self, arguments: &[Expr]) -> Outcome {
        let (periods, payment, present, future, timing) = self.annuity(arguments)?;
        let guess = self.number_or_default(arguments, 5, 0.1)?;
        if periods <= 0.0 || guess <= -1.0 {
            return Err(ExcelError::Num);
        }
        let balance = |rate: f64| {
            if rate == 0.0 {
                return present + payment * periods + future;
            }
            let growth = (1.0 + rate).powf(periods);
            present * growth + payment * (1.0 + rate * timing) * (growth - 1.0) / rate + future
        };
        rate_root(balance, guess).map(finite).ok_or(ExcelError::Num)
    }

    /// The number of periods for the annuity to reach the future value.
    pub(super) fn nper(&self, arguments: &[Expr]) -> Outcome {
        let (rate, payment, present, future, timing) = self.annuity(arguments)?;
        if rate == 0.0 {
            if payment == 0.0 {
                return Err(ExcelError::Num);
            }
            return Ok(finite(-(present + future) / payment));
        }
        if rate <= -1.0 {
            return Err(ExcelError::Num);
        }
        let flow = payment * (1.0 + rate * timing);
        let ratio = (flow - future * rate) / (flow + present * rate);
        if ratio.is_nan() || ratio <= 0.0 || ratio.is_infinite() {
            return Err(ExcelError::Num);
        }
        Ok(finite(ratio.ln() / rate.ln_1p()))
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
