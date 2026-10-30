//! Python value of a `t="d"` cell's text, following openpyxl's `from_ISO8601`.
//!
//! openpyxl matches the text from its start against
//! `(YYYY-MM-DD)?T?(hh:mm(:ss(.f{1,3})?)?)?Z?` and ignores anything after the
//! match, so `2024-01-15T10:30:00+05:00` reads as a naive datetime. A date alone
//! is a `datetime.date`, a time alone a `datetime.time`, both a
//! `datetime.datetime`. Failing that, it matches
//! `PT(nH)?(nM)?(n(.f{1,3})?S)?` as a `datetime.timedelta`. Fraction digits
//! past the third are ignored. Text that matches neither, or names an invalid
//! date or time, fails openpyxl's load; here it stays a string.

use chrono::NaiveDate;
use pyo3::exceptions::PyOverflowError;
use pyo3::prelude::*;
use pyo3::types::{PyDate, PyDateTime, PyDict, PyTime};

enum IsoValue {
    Date(NaiveDate),
    Time(ClockTime),
    DateTime(NaiveDate, ClockTime),
    Duration {
        hours: Option<f64>,
        minutes: Option<f64>,
        seconds: Option<f64>,
    },
}

struct ClockTime {
    hour: u8,
    minute: u8,
    second: u8,
    microsecond: u32,
}

/// Whether `text` reads as a date, time, datetime or duration.
pub(crate) fn is_iso8601(text: &str) -> bool {
    parse(text.as_bytes()).is_some()
}

/// The Python date, time, datetime or timedelta `text` reads as, or `None`
/// when it reads as none of them.
pub(crate) fn iso8601_to_py(py: Python<'_>, text: &str) -> PyResult<Option<Py<PyAny>>> {
    use chrono::Datelike;

    let Some(value) = parse(text.as_bytes()) else {
        return Ok(None);
    };
    let value = match value {
        IsoValue::Date(date) => {
            PyDate::new(py, date.year(), date.month() as u8, date.day() as u8)?.into_any()
        }
        IsoValue::Time(t) => {
            PyTime::new(py, t.hour, t.minute, t.second, t.microsecond, None)?.into_any()
        }
        IsoValue::DateTime(date, t) => PyDateTime::new(
            py,
            date.year(),
            date.month() as u8,
            date.day() as u8,
            t.hour,
            t.minute,
            t.second,
            t.microsecond,
            None,
        )?
        .into_any(),
        IsoValue::Duration {
            hours,
            minutes,
            seconds,
        } => {
            // `timedelta` itself rounds the float parts, as openpyxl's call does.
            let kwargs = PyDict::new(py);
            for (name, part) in [("hours", hours), ("minutes", minutes), ("seconds", seconds)] {
                if let Some(part) = part {
                    kwargs.set_item(name, part)?;
                }
            }
            let timedelta = py.import("datetime")?.getattr("timedelta")?;
            match timedelta.call((), Some(&kwargs)) {
                Ok(value) => value,
                Err(err) if err.is_instance_of::<PyOverflowError>(py) => return Ok(None),
                Err(err) => return Err(err),
            }
        }
    };
    Ok(Some(value.unbind()))
}

fn parse(text: &[u8]) -> Option<IsoValue> {
    let date = parse_date(text);
    let mut pos = if date.is_some() { 10 } else { 0 };
    if text.get(pos) == Some(&b'T') {
        pos += 1;
    }
    match (date, parse_time(text, pos)) {
        (Some(date), Some(time)) => Some(IsoValue::DateTime(valid_date(date)?, valid_time(time)?)),
        (Some(date), None) => Some(IsoValue::Date(valid_date(date)?)),
        (None, Some(time)) => Some(IsoValue::Time(valid_time(time)?)),
        (None, None) => parse_duration(text),
    }
}

/// `YYYY-MM-DD` at the start of `text`.
fn parse_date(text: &[u8]) -> Option<(u32, u32, u32)> {
    let year = digits(text, 0, 4)?;
    (text.get(4) == Some(&b'-')).then_some(())?;
    let month = digits(text, 5, 2)?;
    (text.get(7) == Some(&b'-')).then_some(())?;
    Some((year, month, digits(text, 8, 2)?))
}

/// `hh:mm`, optionally `:ss` and a `.f{1,3}` fraction, at `pos`.
fn parse_time(text: &[u8], pos: usize) -> Option<(u32, u32, u32, u32)> {
    let hour = digits(text, pos, 2)?;
    (text.get(pos + 2) == Some(&b':')).then_some(())?;
    let minute = digits(text, pos + 3, 2)?;
    let pos = pos + 5;
    let Some(second) = (text.get(pos) == Some(&b':'))
        .then(|| digits(text, pos + 1, 2))
        .flatten()
    else {
        return Some((hour, minute, 0, 0));
    };
    let pos = pos + 3;
    let fraction_len = if text.get(pos) == Some(&b'.') {
        digit_run(text, pos + 1).min(3)
    } else {
        0
    };
    let microsecond = if fraction_len == 0 {
        0
    } else {
        // `int(float(".fff") * 1_000_000)`, truncating as openpyxl does.
        let fraction = std::str::from_utf8(&text[pos..pos + 1 + fraction_len]).ok()?;
        (format!("0{fraction}").parse::<f64>().ok()? * 1_000_000.0) as u32
    };
    Some((hour, minute, second, microsecond))
}

/// `PT`, then at least one of `nH`, `nM` and `n(.f{1,3})?S`, in that order.
fn parse_duration(text: &[u8]) -> Option<IsoValue> {
    let mut pos = text.starts_with(b"PT").then_some(2)?;
    let mut part = |suffix: u8, fraction: bool| -> Option<f64> {
        let whole = digit_run(text, pos);
        if whole == 0 {
            return None;
        }
        let mut end = pos + whole;
        if fraction && text.get(end) == Some(&b'.') {
            let len = digit_run(text, end + 1);
            if (1..=3).contains(&len) {
                end += 1 + len;
            }
        }
        if text.get(end) != Some(&suffix) {
            return None;
        }
        let value = std::str::from_utf8(&text[pos..end]).ok()?.parse().ok()?;
        pos = end + 1;
        Some(value)
    };
    let hours = part(b'H', false);
    let minutes = part(b'M', false);
    let seconds = part(b'S', true);
    if hours.is_none() && minutes.is_none() && seconds.is_none() {
        return None;
    }
    // `timedelta` holds under 1,000,000,000 days.
    let total =
        hours.unwrap_or(0.0) * 3600.0 + minutes.unwrap_or(0.0) * 60.0 + seconds.unwrap_or(0.0);
    (total < 86_400.0 * 1e9).then_some(IsoValue::Duration {
        hours,
        minutes,
        seconds,
    })
}

/// A date `datetime.date` accepts: years 1 through 9999 of the Gregorian calendar.
fn valid_date((year, month, day): (u32, u32, u32)) -> Option<NaiveDate> {
    (year >= 1).then_some(())?;
    NaiveDate::from_ymd_opt(year as i32, month, day)
}

/// A time `datetime.time` accepts, with no leap second.
fn valid_time((hour, minute, second, microsecond): (u32, u32, u32, u32)) -> Option<ClockTime> {
    (hour < 24 && minute < 60 && second < 60).then_some(ClockTime {
        hour: hour as u8,
        minute: minute as u8,
        second: second as u8,
        microsecond,
    })
}

/// The value of exactly `len` ASCII digits at `pos`.
fn digits(text: &[u8], pos: usize, len: usize) -> Option<u32> {
    let run = text.get(pos..pos + len)?;
    run.iter()
        .all(u8::is_ascii_digit)
        .then(|| run.iter().fold(0, |n, d| n * 10 + u32::from(d - b'0')))
}

/// The number of ASCII digits starting at `pos`.
fn digit_run(text: &[u8], pos: usize) -> usize {
    text.get(pos..).map_or(0, |rest| {
        rest.iter().take_while(|b| b.is_ascii_digit()).count()
    })
}
