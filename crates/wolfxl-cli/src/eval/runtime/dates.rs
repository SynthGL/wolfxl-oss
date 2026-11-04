//! Date serials and the date and time functions.
//!
//! In the 1900 date system serial 1 is 1900-01-01 and serial 60 is the
//! fictional 1900-02-29 that Excel keeps for compatibility; serial 0 reads as
//! 1900-01-00. In the 1904 date system serial 0 is 1904-01-01.

use super::{finite, Context, Outcome};
use crate::eval::{DataSource, ExcelError, Expr, Function, Value};

/// Days from 1970-01-01 to a proleptic Gregorian date.
fn days_from_civil(year: i64, month: i64, day: i64) -> i64 {
    let year = if month <= 2 { year - 1 } else { year };
    let era = year.div_euclid(400);
    let year_of_era = year.rem_euclid(400);
    let shifted_month = if month > 2 { month - 3 } else { month + 9 };
    let day_of_year = (153 * shifted_month + 2) / 5 + day - 1;
    let day_of_era = year_of_era * 365 + year_of_era / 4 - year_of_era / 100 + day_of_year;
    era * 146_097 + day_of_era - 719_468
}

/// The proleptic Gregorian date `days` after 1970-01-01.
fn civil_from_days(days: i64) -> (i64, u32, u32) {
    let days = days + 719_468;
    let era = days.div_euclid(146_097);
    let day_of_era = days.rem_euclid(146_097);
    let year_of_era =
        (day_of_era - day_of_era / 1460 + day_of_era / 36_524 - day_of_era / 146_096) / 365;
    let day_of_year = day_of_era - (365 * year_of_era + year_of_era / 4 - year_of_era / 100);
    let shifted_month = (5 * day_of_year + 2) / 153;
    let day = day_of_year - (153 * shifted_month + 2) / 5 + 1;
    let month = if shifted_month < 10 {
        shifted_month + 3
    } else {
        shifted_month - 9
    };
    let year = year_of_era + era * 400 + i64::from(month <= 2);
    (year, month as u32, day as u32)
}

/// The serial of 9999-12-31, the last date Excel represents.
pub(super) fn max_serial(date1904: bool) -> i64 {
    if date1904 {
        2_957_003
    } else {
        2_958_465
    }
}

/// The serial of the first day of a month, months past December rolling
/// into later years. May be out of range; callers check.
fn month_start(year: i64, month: i64, date1904: bool) -> i64 {
    let months = year * 12 + month - 1;
    let (year, month) = (months.div_euclid(12), months.rem_euclid(12) + 1);
    let days = days_from_civil(year, month, 1);
    if date1904 {
        days - days_from_civil(1904, 1, 1)
    } else if (year, month) < (1900, 3) {
        days - days_from_civil(1899, 12, 31)
    } else {
        // One more for the fictional 1900-02-29.
        days - days_from_civil(1899, 12, 30)
    }
}

/// Year, month, and day of a whole, in-range serial.
pub(super) fn civil(serial: i64, date1904: bool) -> (i64, u32, u32) {
    if date1904 {
        civil_from_days(serial + days_from_civil(1904, 1, 1))
    } else if serial == 0 {
        (1900, 1, 0)
    } else if serial == 60 {
        (1900, 2, 29)
    } else if serial < 60 {
        civil_from_days(serial + days_from_civil(1899, 12, 31))
    } else {
        civil_from_days(serial + days_from_civil(1899, 12, 30))
    }
}

/// Day of the week for a whole serial, with Sunday as 0.
pub(super) fn weekday(serial: i64, date1904: bool) -> u32 {
    // Serial 1 of the 1900 system reads as a Sunday; 1904-01-01 was a Friday.
    (serial + if date1904 { 5 } else { 6 }).rem_euclid(7) as u32
}

fn days_in_month(year: i64, month: u32, date1904: bool) -> u32 {
    match month {
        2 if !date1904 && year == 1900 => 29,
        2 if year % 4 == 0 && (year % 100 != 0 || year % 400 == 0) => 29,
        2 => 28,
        4 | 6 | 9 | 11 => 30,
        _ => 31,
    }
}

impl<'a, S: DataSource> Context<'a, S> {
    /// A computed serial, or `#NUM!` when it falls outside the date system.
    fn in_range(&self, serial: i64) -> Outcome {
        if (0..=max_serial(self.source.date1904())).contains(&serial) {
            Ok(Value::Number(serial as f64))
        } else {
            Err(ExcelError::Num)
        }
    }

    /// A date argument as a whole serial in range.
    fn serial(&self, expression: &Expr) -> Result<i64, ExcelError> {
        let value = self.number(expression)?;
        if value < 0.0 || value >= (max_serial(self.source.date1904()) + 1) as f64 {
            return Err(ExcelError::Num);
        }
        Ok(value.floor() as i64)
    }

    pub(super) fn date(&self, arguments: &[Expr]) -> Outcome {
        let date1904 = self.source.date1904();
        let year = self.number(&arguments[0])?.trunc();
        // Bound months and days so the arithmetic below cannot overflow;
        // anything this far out is past 9999 either way.
        let month = self.number(&arguments[1])?.trunc().clamp(-1e7, 1e7) as i64;
        let day = self.number(&arguments[2])?.trunc().clamp(-1e9, 1e9) as i64;
        if !(0.0..10_000.0).contains(&year) {
            return Err(ExcelError::Num);
        }
        let year = year as i64;
        // Two-digit and other early years count from 1900.
        let year = if year < 1900 { year + 1900 } else { year };
        self.in_range(month_start(year, month, date1904) + day - 1)
    }

    pub(super) fn date_part(&self, function: Function, arguments: &[Expr]) -> Outcome {
        let (year, month, day) = civil(self.serial(&arguments[0])?, self.source.date1904());
        Ok(Value::Number(match function {
            Function::Year => year as f64,
            Function::Month => f64::from(month),
            _ => f64::from(day),
        }))
    }

    pub(super) fn time_part(&self, function: Function, arguments: &[Expr]) -> Outcome {
        let value = self.number(&arguments[0])?;
        if value < 0.0 {
            return Err(ExcelError::Num);
        }
        // Times resolve to the nearest second.
        let seconds = ((value - value.floor()) * 86_400.0).round() as u32 % 86_400;
        Ok(Value::Number(f64::from(match function {
            Function::Hour => seconds / 3600,
            Function::Minute => seconds / 60 % 60,
            _ => seconds % 60,
        })))
    }

    pub(super) fn days(&self, arguments: &[Expr]) -> Outcome {
        let end = self.serial(&arguments[0])?;
        let start = self.serial(&arguments[1])?;
        Ok(Value::Number((end - start) as f64))
    }

    /// `EDATE`, or `EOMONTH` when `end_of_month` is set.
    pub(super) fn edate(&self, arguments: &[Expr], end_of_month: bool) -> Outcome {
        let date1904 = self.source.date1904();
        let (year, month, day) = civil(self.serial(&arguments[0])?, date1904);
        let months = self.number(&arguments[1])?.trunc().clamp(-1e6, 1e6) as i64;
        let target = year * 12 + i64::from(month) - 1 + months;
        let (year, month) = (target.div_euclid(12), target.rem_euclid(12) as u32 + 1);
        if !(0..10_000).contains(&year) {
            return Err(ExcelError::Num);
        }
        let start = month_start(year, i64::from(month), date1904);
        let last = days_in_month(year, month, date1904);
        let day = if end_of_month { last } else { day.min(last) };
        self.in_range(start + i64::from(day) - 1)
    }

    /// Today's serial in local time, with the time of day for `NOW`.
    pub(super) fn today(&self, with_time: bool) -> Value {
        use chrono::{Datelike, Timelike};
        let now = chrono::Local::now().naive_local();
        let date = month_start(
            i64::from(now.year()),
            i64::from(now.month()),
            self.source.date1904(),
        ) + i64::from(now.day())
            - 1;
        if !with_time {
            return Value::Number(date as f64);
        }
        let seconds = f64::from(now.num_seconds_from_midnight())
            + f64::from(now.nanosecond().min(999_999_999)) / 1e9;
        finite(date as f64 + seconds / 86_400.0)
    }
}

/// A number as a whole, in-range serial for `TEXT`.
pub(super) fn text_serial(value: f64, date1904: bool) -> Option<(i64, f64)> {
    if value < 0.0 || value >= (max_serial(date1904) + 1) as f64 {
        return None;
    }
    Some((value.floor() as i64, value - value.floor()))
}

#[cfg(test)]
mod tests {
    use super::{civil, month_start, weekday};

    #[test]
    fn serials_follow_both_date_systems() {
        assert_eq!(month_start(1900, 1, false), 1);
        assert_eq!(month_start(1900, 3, false), 61);
        assert_eq!(month_start(2024, 1, false), 45292);
        assert_eq!(month_start(1904, 1, true), 0);
        assert_eq!(month_start(2024, 1, true), 45292 - 1462);
        assert_eq!(civil(60, false), (1900, 2, 29));
        assert_eq!(civil(59, false), (1900, 2, 28));
        assert_eq!(civil(61, false), (1900, 3, 1));
        assert_eq!(civil(0, false), (1900, 1, 0));
        assert_eq!(civil(2_958_465, false), (9999, 12, 31));
        assert_eq!(civil(2_957_003, true), (9999, 12, 31));
        assert_eq!(weekday(45292, false), 1);
        assert_eq!(weekday(0, true), 5);
    }
}
