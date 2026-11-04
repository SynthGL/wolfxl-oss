//! Row filters and the lenient numeric reading they share with table
//! comparisons.
//!
//! Exported ledgers spell numbers as money: `$1,200.00`, `1,200`, and
//! accounting negatives such as `(500.00)`. [`lenient_number`] reads those
//! spellings; anything else that is not a plain decimal stays non-numeric.
//!
//! A [`Condition`] is one `COLUMN<op>VALUE` row filter. A right side that
//! [`lenient_number`] reads compares numerically against [`cell_number`];
//! any other right side compares case-insensitively against the cell's text.
//! [`resolve_column`] maps a header name to its column.

use std::borrow::Cow;
use std::cmp::Ordering;

use anyhow::{anyhow, bail, Result};
use wolfxl_core::CellValue;

use crate::input::{delimited_value, Provenance};

/// Currency symbols accepted before the digits.
const CURRENCY: [char; 4] = ['$', '€', '£', '¥'];

/// Read money-style text as a number: surrounding spaces, one sign or
/// accounting parentheses, one leading currency symbol, comma thousands
/// separators in groups of three, and an optional trailing percent (÷100).
/// Returns `None` for any other text, retaining leading-zero identifiers.
pub(crate) fn lenient_number(text: &str) -> Option<f64> {
    let mut body = text.trim();
    let mut negative = false;
    if let Some(inner) = body
        .strip_prefix('(')
        .and_then(|rest| rest.strip_suffix(')'))
    {
        negative = true;
        body = inner.trim();
    }
    if let Some(rest) = body.strip_prefix('-') {
        if negative {
            return None;
        }
        negative = true;
        body = rest.trim_start();
    } else if let Some(rest) = body.strip_prefix('+') {
        body = rest.trim_start();
    }
    body = body.strip_prefix(CURRENCY).unwrap_or(body).trim_start();
    let percent = body.ends_with('%');
    if percent {
        body = body.strip_suffix('%')?.trim_end();
    }
    if body.starts_with(['+', '-']) {
        return None;
    }
    let plain = if body.contains(',') {
        let integer = body.split(['.', 'e', 'E']).next().unwrap_or(body);
        let mut groups = integer.split(',');
        let first = groups.next().unwrap_or_default();
        let grouped = (1..=3).contains(&first.len())
            && groups.all(|group| group.len() == 3)
            && integer
                .bytes()
                .all(|byte| byte.is_ascii_digit() || byte == b',');
        if !grouped || body[integer.len()..].contains(',') {
            return None;
        }
        body.replace(',', "")
    } else {
        body.to_string()
    };
    let signed = if negative { format!("-{plain}") } else { plain };
    let number = match delimited_value(&signed) {
        CellValue::Int(integer) => integer as f64,
        CellValue::Float(float) => float,
        _ => return None,
    };
    Some(if percent { number / 100.0 } else { number })
}

/// Numeric view of a cell: numbers as stored, text through [`lenient_number`].
/// Booleans, dates, errors, and blanks are not numeric.
pub(crate) fn cell_number(value: &CellValue) -> Option<f64> {
    match value {
        CellValue::Int(integer) => Some(*integer as f64),
        CellValue::Float(float) => Some(*float),
        CellValue::String(text) => lenient_number(text),
        _ => None,
    }
}

/// Comparison named in a `COLUMN<op>VALUE` condition.
#[derive(Copy, Clone, Debug, PartialEq, Eq)]
pub(crate) enum Operator {
    Eq,
    Ne,
    Lt,
    Gt,
    Le,
    Ge,
    /// Case-insensitive substring test on the cell's text.
    Contains,
}

impl Operator {
    fn symbol(self) -> &'static str {
        match self {
            Operator::Eq => "=",
            Operator::Ne => "!=",
            Operator::Lt => "<",
            Operator::Gt => ">",
            Operator::Le => "<=",
            Operator::Ge => ">=",
            Operator::Contains => "~",
        }
    }

    /// Whether a cell that compares with the right side as `ordering`
    /// satisfies the operator. `None` (incomparable) satisfies only `!=`.
    fn holds(self, ordering: Option<Ordering>) -> bool {
        match self {
            Operator::Eq => ordering.is_some_and(Ordering::is_eq),
            Operator::Ne => !ordering.is_some_and(Ordering::is_eq),
            Operator::Lt => ordering.is_some_and(Ordering::is_lt),
            Operator::Gt => ordering.is_some_and(Ordering::is_gt),
            Operator::Le => ordering.is_some_and(Ordering::is_le),
            Operator::Ge => ordering.is_some_and(Ordering::is_ge),
            // Containment is a text test, decided before any ordering.
            Operator::Contains => false,
        }
    }
}

/// One parsed `COLUMN<op>VALUE` row filter.
#[derive(Debug)]
pub(crate) struct Condition {
    column: String,
    operator: Operator,
    /// The right side in lower case, for text comparison.
    folded: String,
    /// The right side read by [`lenient_number`]; `None` selects text mode.
    number: Option<f64>,
}

impl Condition {
    /// Parse `COLUMN<op>VALUE`. The earliest operator character splits the
    /// column from the value, reading `<=`, `>=`, or `!=` when the next
    /// character completes one. Spaces around the column name are ignored;
    /// the value is taken verbatim.
    pub(crate) fn parse(expression: &str) -> Result<Self> {
        let (start, operator) = find_operator(expression).ok_or_else(|| {
            anyhow!(
                "--where {expression:?}: expected COLUMN<op>VALUE with op one of = != < > <= >= ~"
            )
        })?;
        let column = expression[..start].trim();
        if column.is_empty() {
            bail!(
                "--where {expression:?}: missing the column name before {:?}",
                operator.symbol()
            );
        }
        let value = &expression[start + operator.symbol().len()..];
        let number = match operator {
            Operator::Contains => None,
            _ => lenient_number(value),
        };
        Ok(Self {
            column: column.to_owned(),
            operator,
            folded: value.to_lowercase(),
            number,
        })
    }

    /// The header name on the left side.
    pub(crate) fn column(&self) -> &str {
        &self.column
    }

    /// Whether a cell holding `value` passes the condition. In numeric mode
    /// a cell without a number passes only `!=`.
    pub(crate) fn matches(&self, value: &CellValue) -> bool {
        if let Some(target) = self.number {
            let ordering = cell_number(value).and_then(|number| number.partial_cmp(&target));
            return self.operator.holds(ordering);
        }
        let text = cell_text(value);
        if self.operator == Operator::Contains {
            return text.to_lowercase().contains(&self.folded);
        }
        let ordering = text
            .chars()
            .flat_map(char::to_lowercase)
            .cmp(self.folded.chars());
        self.operator.holds(Some(ordering))
    }
}

/// Byte offset and operator of the earliest operator in `expression`. A `!`
/// not followed by `=` belongs to the column name.
fn find_operator(expression: &str) -> Option<(usize, Operator)> {
    let bytes = expression.as_bytes();
    bytes.iter().enumerate().find_map(|(index, &byte)| {
        let completes = bytes.get(index + 1) == Some(&b'=');
        let operator = match byte {
            b'<' if completes => Operator::Le,
            b'>' if completes => Operator::Ge,
            b'!' if completes => Operator::Ne,
            b'=' => Operator::Eq,
            b'<' => Operator::Lt,
            b'>' => Operator::Gt,
            b'~' => Operator::Contains,
            _ => return None,
        };
        Some((index, operator))
    })
}

/// Text a condition compares against: strings and error literals as stored,
/// numbers in plain notation, booleans as `TRUE`/`FALSE`, dates in ISO form.
fn cell_text(value: &CellValue) -> Cow<'_, str> {
    match value {
        CellValue::Empty => Cow::Borrowed(""),
        CellValue::String(text) | CellValue::Error(text) => Cow::Borrowed(text.as_str()),
        CellValue::Int(integer) => Cow::Owned(integer.to_string()),
        CellValue::Float(float) => Cow::Owned(float.to_string()),
        CellValue::Bool(true) => Cow::Borrowed("TRUE"),
        CellValue::Bool(false) => Cow::Borrowed("FALSE"),
        CellValue::Date(date) => Cow::Owned(date.format("%Y-%m-%d").to_string()),
        CellValue::DateTime(datetime) => {
            Cow::Owned(datetime.format("%Y-%m-%d %H:%M:%S").to_string())
        }
        CellValue::Time(time) => Cow::Owned(time.format("%H:%M:%S").to_string()),
    }
}

/// Resolve a unique exact header, a sheet column letter, then a unique
/// case-insensitive header. Duplicate names identify their source header cells.
pub(crate) fn resolve_column(
    headers: &[String],
    provenance: &Provenance,
    name: &str,
) -> Result<usize> {
    let positions = |matches: &dyn Fn(&str) -> bool| {
        headers
            .iter()
            .enumerate()
            .filter(|(_, header)| matches(header))
            .map(|(index, _)| index)
            .collect::<Vec<_>>()
    };
    let ambiguous = |matches: &[usize]| {
        let cells = matches
            .iter()
            .map(|&index| {
                provenance
                    .address(0, index)
                    .unwrap_or_else(|| format!("computed column {}", index + 1))
            })
            .collect::<Vec<_>>();
        anyhow!(
            "column {name:?} is ambiguous: {}; use a column letter (several headers match)",
            cells.join(" and ")
        )
    };
    match positions(&|header| header == name).as_slice() {
        [only] => return Ok(*only),
        [] => {}
        many => return Err(ambiguous(many)),
    }
    if !name.is_empty() && name.bytes().all(|byte| byte.is_ascii_alphabetic()) {
        let column = name.bytes().try_fold(0_u32, |column, byte| {
            column
                .checked_mul(26)?
                .checked_add(u32::from(byte.to_ascii_uppercase() - b'A' + 1))
        });
        if let Some(index) = column.and_then(|column| {
            provenance
                .columns
                .iter()
                .position(|candidate| *candidate == Some(column))
        }) {
            return Ok(index);
        }
    }
    let wanted = name.to_lowercase();
    match positions(&|header| header.to_lowercase() == wanted).as_slice() {
        [only] => Ok(*only),
        [] => bail!(
            "unknown column {name:?}; available headers: {}",
            headers.join(", ")
        ),
        many => Err(ambiguous(many)),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn money_spellings_read_as_numbers() {
        for (text, expected) in [
            ("1200", 1200.0),
            (" 1,200 ", 1200.0),
            ("$1,200.00", 1200.0),
            ("(500.00)", -500.0),
            ("($1,234.50)", -1234.5),
            ("-$3.5", -3.5),
            ("€12", 12.0),
            ("1,234,567.8", 1_234_567.8),
            ("+7", 7.0),
            ("12%", 0.12),
            ("-3.5%", -0.035),
            ("(12%)", -0.12),
            ("1,200%", 12.0),
        ] {
            assert_eq!(lenient_number(text), Some(expected), "{text}");
        }
    }

    #[test]
    fn malformed_spellings_stay_text() {
        for text in [
            "",
            "abc",
            "1,20",
            "12,3456",
            ",100",
            "1.000,5",
            "(-5)",
            "$-5",
            "TRUE",
            "1e",
            "$",
            "()",
            "1,23,4",
            "1,23,4%",
            "12%%",
            "12% growth",
            "01",
            "01%",
        ] {
            assert_eq!(lenient_number(text), None, "{text}");
        }
    }

    fn headers_of(names: &[&str]) -> Vec<String> {
        names.iter().map(|name| name.to_string()).collect()
    }

    fn string(value: &str) -> CellValue {
        CellValue::String(value.to_string())
    }

    #[test]
    fn earliest_operator_splits_and_prefers_two_character_operators() {
        for (expression, column, operator, folded) in [
            ("Amount<=5", "Amount", Operator::Le, "5"),
            ("Amount>=5", "Amount", Operator::Ge, "5"),
            ("Amount!=5", "Amount", Operator::Ne, "5"),
            ("Amount=<5", "Amount", Operator::Eq, "<5"),
            ("Note~a=b", "Note", Operator::Contains, "a=b"),
            ("Yes!No=x", "Yes!No", Operator::Eq, "x"),
            (" Region =West", "Region", Operator::Eq, "west"),
            ("Region=", "Region", Operator::Eq, ""),
        ] {
            let condition = Condition::parse(expression).expect(expression);
            assert_eq!(condition.column(), column, "{expression}");
            assert_eq!(condition.operator, operator, "{expression}");
            assert_eq!(condition.folded, folded, "{expression}");
        }
    }

    #[test]
    fn expressions_without_operator_or_column_are_refused() {
        for expression in ["Region", "", "=5", "  <3", "Yes!No"] {
            assert!(Condition::parse(expression).is_err(), "{expression:?}");
        }
    }

    #[test]
    fn numeric_right_side_compares_lenient_cell_numbers() {
        let above = Condition::parse("Amount>1,000").unwrap();
        assert!(above.matches(&string("$1,200.00")));
        assert!(above.matches(&CellValue::Int(1001)));
        assert!(!above.matches(&string("(500.00)")));
        assert!(!above.matches(&string("n/a")));
        let negative = Condition::parse("Amount=(500)").unwrap();
        assert!(negative.matches(&CellValue::Float(-500.0)));
        let different = Condition::parse("Amount!=5").unwrap();
        assert!(different.matches(&string("n/a")));
        assert!(different.matches(&CellValue::Empty));
        assert!(!different.matches(&string("5.0")));
    }

    #[test]
    fn text_right_side_compares_case_insensitively() {
        let equal = Condition::parse("Region=WEST").unwrap();
        assert!(equal.matches(&string("west")));
        assert!(!equal.matches(&string("Westerly")));
        let before = Condition::parse("Region<north").unwrap();
        assert!(before.matches(&string("East")));
        assert!(!before.matches(&string("NORTH")));
        assert!(!before.matches(&string("Northwest")));
        let truth = Condition::parse("Flag=true").unwrap();
        assert!(truth.matches(&CellValue::Bool(true)));
        let contains = Condition::parse("Code~12").unwrap();
        assert!(contains.matches(&CellValue::Int(3124)));
        assert!(!contains.matches(&CellValue::Empty));
    }

    #[test]
    fn headers_resolve_exactly_then_by_unique_case_insensitive_match() {
        let headers = headers_of(&["id", "ID", "Region"]);
        let provenance = Provenance {
            rows: vec![1],
            columns: vec![Some(1), Some(2), Some(3)],
        };
        assert_eq!(resolve_column(&headers, &provenance, "ID").unwrap(), 1);
        assert_eq!(resolve_column(&headers, &provenance, "region").unwrap(), 2);
        let ambiguous = resolve_column(&headers, &provenance, "Id")
            .unwrap_err()
            .to_string();
        assert!(ambiguous.contains("several headers"), "{ambiguous}");
        let unknown = resolve_column(&headers, &provenance, "Rep")
            .unwrap_err()
            .to_string();
        assert!(
            unknown.contains("available headers: id, ID, Region"),
            "{unknown}"
        );
    }
}
