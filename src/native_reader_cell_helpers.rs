//! Cell-value serialization primitives shared across sheet-data, records, and
//! styles modules. They translate `wolfxl_reader::Cell` into Python objects
//! honouring date detection and the data_only/formula split. The only
//! per-backend state they touch is the caller's `SharedStringCache`.

use std::cell::{OnceCell, RefCell};
use std::collections::HashMap;
use std::sync::Arc;

use chrono::{Datelike, Duration, NaiveDate, Timelike};
use pyo3::exceptions::{PyOverflowError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyDateTime, PyDict, PyString, PyTime};
use pyo3::IntoPyObjectExt;

use wolfxl_reader::{Cell, CellDataType, CellValue};

use crate::iso8601::{is_iso8601, iso8601_to_py};
use crate::streaming::SharedStringTable;
use crate::util::{cell_blank, cell_with_value};

type PyObject = Py<PyAny>;

/// Workbook-scoped Python `str` objects for shared-string cell values, so
/// every native read of one shared-string table entry returns the same `str`.
///
/// Values are looked up by the address of the reader's shared `Arc<str>`, so
/// a lookup never hashes the text. Once linked to the workbook's table, an
/// entry resolves through that table's per-index slot, which streaming readers
/// given the same table also use, so eager and streaming reads of one entry
/// return one `str`. Values outside a linked table are cached per address;
/// each such entry keeps its `Arc` alive, so the address cannot be reused by
/// another string while the cache holds it.
#[derive(Default)]
pub(crate) struct SharedStringCache {
    linked: OnceCell<LinkedTable>,
    unlinked: RefCell<HashMap<usize, (Arc<str>, Py<PyString>)>>,
    rich_runs: RefCell<HashMap<usize, (Arc<[wolfxl_reader::RichTextRun]>, Vec<Py<PyString>>)>>,
}

struct LinkedTable {
    table: Arc<SharedStringTable>,
    index_by_address: HashMap<usize, usize>,
}

fn arc_address(text: &Arc<str>) -> usize {
    Arc::as_ptr(text) as *const u8 as usize
}

impl SharedStringCache {
    /// Link the cache to the workbook's shared-string table on first call and
    /// return that table. Later calls return the first table, so `entries`
    /// must be the same workbook table every time.
    pub(crate) fn link(&self, entries: &[Arc<str>]) -> Arc<SharedStringTable> {
        let linked = self.linked.get_or_init(|| LinkedTable {
            table: Arc::new(SharedStringTable::from_shared(entries.to_vec())),
            index_by_address: entries
                .iter()
                .enumerate()
                .map(|(idx, text)| (arc_address(text), idx))
                .collect(),
        });
        Arc::clone(&linked.table)
    }

    pub(crate) fn rich_run_text(
        &self,
        py: Python<'_>,
        runs: &Arc<[wolfxl_reader::RichTextRun]>,
        index: usize,
    ) -> PyObject {
        let key = Arc::as_ptr(runs) as *const wolfxl_reader::RichTextRun as usize;
        let mut entries = self.rich_runs.borrow_mut();
        let (_, values) = entries.entry(key).or_insert_with(|| {
            (
                Arc::clone(runs),
                runs.iter().map(|run| PyString::new(py, &run.text).unbind()).collect(),
            )
        });
        values[index].clone_ref(py).into_any()
    }

    fn get_or_insert(&self, py: Python<'_>, text: &Arc<str>) -> PyResult<PyObject> {
        let key = arc_address(text);
        if let Some(linked) = self.linked.get() {
            if let Some(&idx) = linked.index_by_address.get(&key) {
                return linked.table.get_py(py, idx);
            }
        }
        let mut entries = self.unlinked.borrow_mut();
        let (_, value) = entries
            .entry(key)
            .or_insert_with(|| (Arc::clone(text), PyString::new(py, text).unbind()));
        Ok(value.clone_ref(py).into_any())
    }
}

/// How a date-formatted serial reads: the workbook's epoch, and whether a
/// duration format such as `[h]:mm:ss` reads as `datetime.timedelta`, as
/// openpyxl reads xlsx. The xls and xlsb readers keep `datetime.datetime`.
#[derive(Clone, Copy, Default)]
pub(crate) struct SerialDates {
    pub(crate) date1904: bool,
    pub(crate) durations: bool,
}

pub(crate) fn cell_to_dict(
    py: Python<'_>,
    cell: &Cell,
    data_only: bool,
    number_format: Option<&str>,
    dates: SerialDates,
    strings: &SharedStringCache,
) -> PyResult<PyObject> {
    if !data_only {
        if let Some(formula) = &cell.formula {
            return formula_to_py(py, formula);
        }
    }
    match &cell.value {
        CellValue::Empty => cell_blank(py),
        CellValue::String(s) if cell.data_type == CellDataType::DateIso => {
            match iso8601_to_py(py, s)? {
                Some(value) => cell_with_value(py, "datetime", value),
                None => cell_with_value(py, "string", s),
            }
        }
        CellValue::String(s) => cell_with_value(py, "string", s),
        CellValue::SharedString(s) => {
            cell_with_value(py, "string", strings.get_or_insert(py, s)?)
        }
        CellValue::Number(n) if is_date_format(number_format) => {
            let d = PyDict::new(py);
            d.set_item("type", "datetime")?;
            d.set_item(
                "value",
                date_serial_to_py(py, cell, *n, dates, number_format)?,
            )?;
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
    dates: SerialDates,
    strings: &SharedStringCache,
) -> PyResult<PyObject> {
    if !data_only {
        if let Some(formula) = &cell.formula {
            return Ok(ensure_formula_prefix(formula).into_py_any(py)?);
        }
    }
    match &cell.value {
        CellValue::Empty => Ok(py.None()),
        CellValue::String(s) => string_to_py(py, cell, s),
        CellValue::SharedString(s) => strings.get_or_insert(py, s),
        CellValue::Number(n) if is_date_format(number_format) => {
            date_serial_to_py(py, cell, *n, dates, number_format)
        }
        CellValue::Number(n) => number_to_py(py, *n, is_float_number(cell)),
        CellValue::Bool(b) => Ok((*b).into_py_any(py)?),
        CellValue::Error(e) => Ok(e.clone().into_py_any(py)?),
    }
}

pub(crate) fn cell_to_plain_unstyled_data_only(
    py: Python<'_>,
    cell: &Cell,
    strings: &SharedStringCache,
) -> PyResult<PyObject> {
    match &cell.value {
        CellValue::Empty => Ok(py.None()),
        CellValue::String(s) => string_to_py(py, cell, s),
        CellValue::SharedString(s) => strings.get_or_insert(py, s),
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
        // openpyxl keeps `inlineStr` when the cell has no `<is>` child.
        CellDataType::InlineString if matches!(cell.value, CellValue::Empty) => "inlineStr",
        CellDataType::DateIso if matches!(cell.value, CellValue::Empty) => "datetime",
        CellDataType::DateIso if matches!(&cell.value, CellValue::String(s) if is_iso8601(s)) => {
            "datetime"
        }
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

/// A string cell value, or the date or time a `t="d"` cell's text reads as.
fn string_to_py(py: Python<'_>, cell: &Cell, text: &str) -> PyResult<PyObject> {
    if cell.data_type == CellDataType::DateIso {
        if let Some(value) = iso8601_to_py(py, text)? {
            return Ok(value);
        }
    }
    Ok(text.into_py_any(py)?)
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
/// millisecond and is kept as microseconds. A duration format reads as a
/// `datetime.timedelta` when `dates.durations` is set and otherwise always as
/// a `datetime.datetime`. A serial outside Python's year range stays a number,
/// as it does on streamed sheets.
fn date_serial_to_py(
    py: Python<'_>,
    cell: &Cell,
    serial: f64,
    dates: SerialDates,
    number_format: Option<&str>,
) -> PyResult<PyObject> {
    let duration = number_format.is_some_and(is_timedelta_format);
    if duration && dates.durations {
        return serial_to_timedelta(py, cell, serial);
    }
    if (0.0..1.0).contains(&serial) && !duration {
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
    let Some(dt) = excel_serial_to_datetime(serial, dates.date1904)
        .filter(|dt| (1..=9999).contains(&dt.year()))
    else {
        return number_to_py(py, serial, is_float_number(cell));
    };
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

/// `from_excel(serial, timedelta=True)`, through the same function the
/// streaming reader uses; a serial too large for `timedelta` stays a number.
fn serial_to_timedelta(py: Python<'_>, cell: &Cell, serial: f64) -> PyResult<PyObject> {
    let from_excel = py.import("wolfxl.utils.datetime")?.getattr("from_excel")?;
    let kwargs = PyDict::new(py);
    kwargs.set_item("timedelta", true)?;
    match from_excel.call((serial,), Some(&kwargs)) {
        Ok(value) => Ok(value.unbind()),
        Err(err)
            if err.is_instance_of::<PyOverflowError>(py)
                || err.is_instance_of::<PyValueError>(py) =>
        {
            number_to_py(py, serial, is_float_number(cell))
        }
        Err(err) => Err(err),
    }
}

pub(crate) fn excel_serial_to_datetime(
    serial: f64,
    date1904: bool,
) -> Option<chrono::NaiveDateTime> {
    let epoch = if date1904 {
        NaiveDate::from_ymd_opt(1904, 1, 1).unwrap()
    } else {
        NaiveDate::from_ymd_opt(1899, 12, 30).unwrap()
    };
    if !serial.is_finite() {
        return None;
    }
    let mut days = serial.trunc() as i64;
    if !date1904 && serial > 0.0 && serial < 60.0 {
        days += 1;
    }
    let fraction = serial - serial.trunc();
    let millis = (fraction * 86_400_000.0).round() as i64;
    epoch
        .and_hms_opt(0, 0, 0)
        .unwrap()
        .checked_add_signed(Duration::try_days(days)?)?
        .checked_add_signed(Duration::milliseconds(millis))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn linked_cache_shares_table_strs_with_streaming() {
        pyo3::Python::initialize();
        Python::attach(|py| {
            let entries: Vec<Arc<str>> = vec![Arc::from("a"), Arc::from("x".repeat(4096))];
            let cache = SharedStringCache::default();
            let table = cache.link(&entries);
            assert!(Arc::ptr_eq(&table, &cache.link(&entries)));

            // Eager reads of an entry and the table handed to streaming
            // readers return one object.
            let eager = cache.get_or_insert(py, &entries[1]).unwrap();
            assert!(eager.bind(py).is(table.get_py(py, 1).unwrap()));
            assert!(eager
                .bind(py)
                .is(cache.get_or_insert(py, &Arc::clone(&entries[1])).unwrap()));

            // Equal text at another address is a different entry, found
            // without hashing the text.
            let copy: Arc<str> = Arc::from("x".repeat(4096));
            let other = cache.get_or_insert(py, &copy).unwrap();
            assert!(!other.bind(py).is(&eager));
            assert!(other.bind(py).is(cache.get_or_insert(py, &copy).unwrap()));
            assert_eq!(other.extract::<String>(py).unwrap(), "x".repeat(4096));
        });
    }
}
