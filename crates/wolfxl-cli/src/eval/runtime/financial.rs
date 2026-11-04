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

/// Whether `left` and `right` lie on opposite sides of zero. Signs are
/// compared rather than multiplied, because the product of two tiny values
/// underflows to zero, and a value that underflowed to `0.0` counts as
/// positive so it is not mistaken for a root.
fn opposite_signs(left: f64, right: f64) -> bool {
    left.is_sign_negative() != right.is_sign_negative()
}

/// Lowest and highest rates the bracket search in [`rate_root`] samples.
const BRACKET: (f64, f64) = (-0.99, 10.0);

/// A root of `function` above -1: Newton's method from the guess, and when
/// that does not settle, bisection on the sign change nearest the guess.
/// Convergence never compares the balance with a fixed tolerance, so the
/// result does not change when every cash flow is scaled by a constant, and
/// a balance that only approaches zero, such as `100*(1+rate)^360`, has no
/// root. `function` runs at most about 420 times: 300 for Newton, 23 to find
/// a sign change, and 100 to bisect it.
fn rate_root(function: impl Fn(f64) -> f64, guess: f64) -> Option<f64> {
    let mut rate = guess;
    for _ in 0..100 {
        let value = function(rate);
        if !value.is_finite() {
            break;
        }
        let step = 1e-7 * rate.abs().max(1.0);
        if value == 0.0 {
            // An exact zero is a root only where the balance changes sign.
            if opposite_signs(function(rate - step), function(rate + step)) {
                return Some(rate);
            }
            break;
        }
        let slope = (function(rate + step) - function(rate - step)) / (2.0 * step);
        if slope == 0.0 || !slope.is_finite() {
            break;
        }
        let next = rate - value / slope;
        if next <= -1.0 || !next.is_finite() {
            break;
        }
        if (next - rate).abs() < 1e-12 * rate.abs().max(1.0) {
            return Some(next);
        }
        rate = next;
    }
    let (mut low, mut high) = bracket(&function, guess)?;
    let low_negative = function(low).is_sign_negative();
    for _ in 0..100 {
        let middle = (low + high) / 2.0;
        if middle <= low || middle >= high {
            break;
        }
        if function(middle).is_sign_negative() == low_negative {
            low = middle;
        } else {
            high = middle;
        }
    }
    Some((low + high) / 2.0)
}

/// The interval nearest the guess, within [`BRACKET`], whose ends lie on
/// opposite sides of zero. The search steps outward on both sides,
/// doubling the step from 0.01, so it samples finely near the guess and
/// reaches either end of the range within 11 steps. Each point is sampled
/// once.
fn bracket(function: &impl Fn(f64) -> f64, guess: f64) -> Option<(f64, f64)> {
    let (bottom, top) = BRACKET;
    let start = guess.clamp(bottom, top);
    let first = function(start);
    let (mut left, mut right) = ((start, first), (start, first));
    let mut step = 0.01;
    while left.0 > bottom || right.0 < top {
        if right.0 < top {
            let rate = (right.0 + step).min(top);
            let value = function(rate);
            if changes_sign(right.1, value) {
                return Some((right.0, rate));
            }
            right = (rate, value);
        }
        if left.0 > bottom {
            let rate = (left.0 - step).max(bottom);
            let value = function(rate);
            if changes_sign(value, left.1) {
                return Some((rate, left.0));
            }
            left = (rate, value);
        }
        step *= 2.0;
    }
    None
}

/// Whether two finite sampled values straddle zero.
fn changes_sign(left: f64, right: f64) -> bool {
    left.is_finite() && right.is_finite() && opposite_signs(left, right)
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

    /// The rate at which the cash flows' net present value is zero.
    pub(super) fn irr(&self, arguments: &[Expr]) -> Outcome {
        let values = self.cash_flows(&arguments[..1])?;
        let guess = self.number_or(arguments, 1, 0.1)?;
        if values.len() < 2
            || !values.iter().any(|&value| value > 0.0)
            || !values.iter().any(|&value| value < 0.0)
        {
            return Err(ExcelError::Num);
        }
        rate_root(|rate| discounted(rate, &values, 0), guess)
            .map(finite)
            .ok_or(ExcelError::Num)
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

#[cfg(test)]
mod tests {
    use std::cell::Cell;

    use super::rate_root;

    #[test]
    fn bisection_finds_a_root_newton_cannot_reach() {
        // Each Newton step on a cube root doubles the distance to the root.
        let root = rate_root(|rate: f64| (rate - 3.0).cbrt(), 0.1).expect("a root");
        assert!((root - 3.0).abs() < 1e-12, "{root}");
    }

    #[test]
    fn rootless_balances_stop_within_the_evaluation_bound() {
        // Workbook-sized cash flows make every evaluation expensive, so the
        // count stays bounded whatever the balance looks like.
        let calls = Cell::new(0_u32);
        let balance = |rate: f64| {
            calls.set(calls.get() + 1);
            1.0 + (rate - 5.0).powi(2)
        };
        assert_eq!(rate_root(balance, 0.1), None);
        assert!(calls.get() <= 430, "{} evaluations", calls.get());
    }
}
