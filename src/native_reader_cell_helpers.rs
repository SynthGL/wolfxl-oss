//! Cell-value serialization primitives shared across sheet-data, records, and
//! styles modules. These keep no per-backend state — they translate
//! `wolfxl_reader::Cell` into Python objects honouring date detection and the
//! data_only/formula split.

use chrono::{Datelike, Duration, NaiveDate, Timelike};
use pyo3::prelude::*;
use pyo3::types::{PyDateTime, PyDict, PyTime};
use pyo3::IntoPyObjectExt;

use wolfxl_reader::{Cell, CellDataType, CellValue};

use crate::util::{cell_blank, cell_with_value};

type PyObject = Py<PyAny>;

pub(crate) fn cell_to_dict(
    py: Python<'_>,
    cell: &Cell,
    data_only: bool,
    number_format: Option<&str>,
    date1904: bool,
) -> PyResult<PyObject> {
    if !data_only {
        if let Some(formula) = &cell.formula {
            return formula_to_py(py, formula);
        }
    }
    match &cell.value {
        CellValue::Empty => cell_blank(py),
        CellValue::String(s) => cell_with_value(py, "string", s),
        CellValue::Number(n) if is_date_format(number_format) => {
            let d = PyDict::new(py);
            d.set_item("type", "datetime")?;
            d.set_item("value", date_serial_to_py(py, *n, date1904, number_format)?)?;
            Ok(d.into())
        }
        CellValue::Number(n) => {
            let d = PyDict::new(py);
            d.set_item("type", "number")?;
            d.set_item("value", number_to_py(py, *n, is_float_number(cell))?)?;
            Ok(d.into())
        }
        CellValue::Bool(b) => cell_with_value(py, "boolean", *b),
        CellValue::Error(e) => cell_with_value(py, "error", e),
    }
}

pub(crate) fn cell_to_plain(
    py: Python<'_>,
    cell: &Cell,
    data_only: bool,
    number_format: Option<&str>,
    date1904: bool,
) -> PyResult<PyObject> {
    if !data_only {
        if let Some(formula) = &cell.formula {
            return Ok(ensure_formula_prefix(formula).into_py_any(py)?);
        }
    }
    match &cell.value {
        CellValue::Empty => Ok(py.None()),
        CellValue::String(s) => Ok(s.clone().into_py_any(py)?),
        CellValue::Number(n) if is_date_format(number_format) => {
            date_serial_to_py(py, *n, date1904, number_format)
        }
        CellValue::Number(n) => number_to_py(py, *n, is_float_number(cell)),
        CellValue::Bool(b) => Ok((*b).into_py_any(py)?),
        CellValue::Error(e) => Ok(e.clone().into_py_any(py)?),
    }
}

pub(crate) fn cell_to_plain_unstyled_data_only(py: Python<'_>, cell: &Cell) -> PyResult<PyObject> {
    match &cell.value {
        CellValue::Empty => Ok(py.None()),
        CellValue::String(s) => Ok(s.clone().into_py_any(py)?),
        CellValue::Number(n) => number_to_py(py, *n, is_float_number(cell)),
        CellValue::Bool(b) => Ok((*b).into_py_any(py)?),
        CellValue::Error(e) => Ok(e.clone().into_py_any(py)?),
    }
}

pub(crate) fn native_data_type(
    cell: &Cell,
    data_only: bool,
    number_format: Option<&str>,
) -> &'static str {
    if !data_only && cell.formula.is_some() {
        return "formula";
    }
    match cell.data_type {
        CellDataType::Bool => "boolean",
        CellDataType::Error => "error",
        CellDataType::FormulaString if matches!(cell.value, CellValue::Empty) => "str",
        CellDataType::DateIso
        | CellDataType::InlineString
        | CellDataType::SharedString
        | CellDataType::FormulaString => "string",
        CellDataType::Number | CellDataType::FloatNumber => {
            if matches!(cell.value, CellValue::Empty) {
                "blank"
            } else if is_date_format(number_format) {
                "datetime"
            } else {
                "number"
            }
        }
    }
}

pub(crate) fn ensure_formula_prefix(formula: &str) -> String {
    if formula.starts_with('=') {
        formula.to_string()
    } else {
        format!("={formula}")
    }
}

fn is_float_number(cell: &Cell) -> bool {
    cell.data_type == CellDataType::FloatNumber
}

/// Integral numbers read as `int` unless the cell text held a decimal point or
/// exponent (`float_text`, from `CellDataType::FloatNumber` or `ValueCell`), which
/// openpyxl's `_cast_number` reads as `float`.
pub(crate) fn number_to_py(py: Python<'_>, value: f64, float_text: bool) -> PyResult<PyObject> {
    if !float_text && value.is_finite() && value.fract() == 0.0 {
        let int_value = value as i64;
        if int_value as f64 == value {
            return Ok(int_value.into_py_any(py)?);
        }
    }
    Ok(value.into_py_any(py)?)
}

pub(crate) fn formula_to_py(py: Python<'_>, formula: &str) -> PyResult<PyObject> {
    let formula = ensure_formula_prefix(formula);
    let d = PyDict::new(py);
    d.set_item("type", "formula")?;
    d.set_item("formula", &formula)?;
    d.set_item("value", &formula)?;
    Ok(d.into())
}

pub(crate) fn is_date_format(format: Option<&str>) -> bool {
    let Some(format) = format else {
        return false;
    };
    let first = format.split(';').next().unwrap_or(format);
    let mut in_quote = false;
    let chars: Vec<char> = first.chars().collect();
    let mut i = 0;
    while i < chars.len() {
        let ch = chars[i];
        if ch == '"' {
            in_quote = !in_quote;
            i += 1;
            continue;
        }
        if in_quote {
            i += 1;
            continue;
        }
        if ch == '[' {
            let mut j = i + 1;
            while j < chars.len() && chars[j] != ']' {
                j += 1;
            }
            if j < chars.len() {
                let bracket: String = chars[i + 1..j].iter().collect();
                let lower = bracket.to_ascii_lowercase();
                if lower != "h"
                    && lower != "hh"
                    && lower != "m"
                    && lower != "mm"
                    && lower != "s"
                    && lower != "ss"
                {
                    i = j + 1;
                    continue;
                }
            }
        }
        if matches!(
            ch,
            'd' | 'D' | 'm' | 'M' | 'h' | 'H' | 'y' | 'Y' | 's' | 'S'
        ) {
            let prev = i.checked_sub(1).and_then(|idx| chars.get(idx)).copied();
            if prev != Some('_') && prev != Some('\\') {
                return true;
            }
        }
        i += 1;
    }
    false
}

/// Whether `format` is a duration format such as `[h]:mm:ss`, by the rule of
/// `wolfxl.styles.numbers.is_timedelta_format`: its first section holds an
/// elapsed-time token `[h]`, `[hh]`, `[m]`, `[mm]`, `[s]` or `[ss]`.
fn is_timedelta_format(format: &str) -> bool {
    let first = format.split(';').next().unwrap_or(format);
    first.split('[').skip(1).any(|rest| {
        matches!(
            rest.split_once(']'),
            Some((token, _)) if matches!(
                token.to_ascii_lowercase().as_str(),
                "h" | "hh" | "m" | "mm" | "s" | "ss"
            )
        )
    })
}

/// Python value of a date-formatted serial, matching openpyxl's `from_excel`:
/// a serial in `[0, 1)` is a `datetime.time` (serial 0 is midnight), any
/// other serial a `datetime.datetime`. The day fraction rounds to the
/// millisecond and is kept as microseconds. A duration format always yields a
/// `datetime.datetime`, which `Cell.value` converts to `datetime.timedelta`.
fn date_serial_to_py(
    py: Python<'_>,
    serial: f64,
    date1904: bool,
    number_format: Option<&str>,
) -> PyResult<PyObject> {
    if (0.0..1.0).contains(&serial) && !number_format.is_some_and(is_timedelta_format) {
        let millis = (serial * 86_400_000.0).round() as u32;
        if millis < 86_400_000 {
            let secs = millis / 1000;
            let time = PyTime::new(
                py,
                (secs / 3600) as u8,
                (secs % 3600 / 60) as u8,
                (secs % 60) as u8,
                millis % 1000 * 1000,
                None,
            )?;
            return Ok(time.into_any().unbind());
        }
    }
    let dt = excel_serial_to_datetime(serial, date1904);
    let py_dt = PyDateTime::new(
        py,
        dt.year(),
        dt.month() as u8,
        dt.day() as u8,
        dt.hour() as u8,
        dt.minute() as u8,
        dt.second() as u8,
        dt.nanosecond() / 1000,
        None,
    )?;
    Ok(py_dt.into_any().unbind())
}

pub(crate) fn excel_serial_to_datetime(serial: f64, date1904: bool) -> chrono::NaiveDateTime {
    let epoch = if date1904 {
        NaiveDate::from_ymd_opt(1904, 1, 1).unwrap()
    } else {
        NaiveDate::from_ymd_opt(1899, 12, 30).unwrap()
    };
    let mut days = serial.trunc() as i64;
    if !date1904 && serial > 0.0 && serial < 60.0 {
        days += 1;
    }
    let fraction = serial - serial.trunc();
    let millis = (fraction * 86_400_000.0).round() as i64;
    epoch.and_hms_opt(0, 0, 0).unwrap() + Duration::days(days) + Duration::milliseconds(millis)
}
