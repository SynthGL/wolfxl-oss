//! `TEXT` and the number format codes it applies.
//!
//! Supported: up to four `;`-separated sections (positive, negative, zero,
//! text); digit placeholders `0`, `#`, and `?`; the decimal point; `,` as a
//! thousands separator or, after the last digit, a scale of one thousand;
//! `%`; scientific `E+` and `E-`; `General`; `@`; quoted text, `\` escapes,
//! `_` spacing, and `*` fills (shown without the fill); and the date and time
//! codes `y`, `m`, `d`, `h`, `s`, `AM/PM`, and `A/P`. Bracketed colors,
//! conditions, and elapsed-time codes are ignored.

use super::dates::{civil, text_serial, weekday};
use super::{general, lenient_number, round_decimal, Context, Outcome, Rounding};
use crate::eval::{DataSource, ExcelError, Expr, Value};

const MONTHS: [&str; 12] = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
];

const WEEKDAYS: [&str; 7] = [
    "Sunday",
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
];

#[derive(Clone, Debug, PartialEq)]
enum Part {
    Literal(String),
    /// `0`, `#`, or `?`.
    Digit(char),
    Point,
    Comma,
    Percent,
    /// `E+` when `true`, `E-` when `false`.
    Exponent(bool),
    General,
    Text,
    Year(usize),
    /// `m` runs, which mean minutes after hours or before seconds.
    Month(usize),
    Minute(usize),
    Day(usize),
    Hour(usize),
    Second(usize),
    /// `AM/PM` when `true`, `A/P` when `false`, with the case written.
    Meridiem(bool, bool),
}

impl Part {
    fn is_date(&self) -> bool {
        matches!(
            self,
            Self::Year(_)
                | Self::Month(_)
                | Self::Minute(_)
                | Self::Day(_)
                | Self::Hour(_)
                | Self::Second(_)
                | Self::Meridiem(..)
        )
    }
}

/// Split a format code into sections at `;` outside quotes and escapes.
fn sections(code: &str) -> Vec<&str> {
    let mut sections = Vec::new();
    let (mut start, mut quoted, mut escaped) = (0, false, false);
    for (offset, character) in code.char_indices() {
        if escaped {
            escaped = false;
        } else if character == '"' {
            quoted = !quoted;
        } else if !quoted && matches!(character, '\\' | '_' | '*') {
            escaped = true;
        } else if !quoted && character == ';' {
            sections.push(&code[start..offset]);
            start = offset + 1;
        }
    }
    sections.push(&code[start..]);
    sections
}

fn parse(section: &str) -> Vec<Part> {
    let characters: Vec<char> = section.chars().collect();
    let mut parts = Vec::new();
    let mut index = 0;
    let starts_with = |index: usize, word: &str| {
        let word: Vec<char> = word.chars().collect();
        characters.len() >= index + word.len()
            && characters[index..index + word.len()]
                .iter()
                .zip(&word)
                .all(|(a, b)| a.eq_ignore_ascii_case(b))
    };
    while index < characters.len() {
        let character = characters[index];
        let run = characters[index..]
            .iter()
            .take_while(|next| next.eq_ignore_ascii_case(&character))
            .count();
        let (part, length) = match character {
            '"' => {
                let text: String = characters[index + 1..]
                    .iter()
                    .take_while(|&&next| next != '"')
                    .collect();
                let length = text.chars().count() + 2;
                (Part::Literal(text), length)
            }
            '\\' => (
                Part::Literal(
                    characters
                        .get(index + 1)
                        .map(char::to_string)
                        .unwrap_or_default(),
                ),
                2,
            ),
            '_' => (Part::Literal(" ".into()), 2),
            '*' => (Part::Literal(String::new()), 2),
            '[' => {
                let length = characters[index..]
                    .iter()
                    .position(|&next| next == ']')
                    .map_or(characters.len() - index, |end| end + 1);
                (Part::Literal(String::new()), length)
            }
            '0' | '#' | '?' => (Part::Digit(character), 1),
            '.' => (Part::Point, 1),
            ',' => (Part::Comma, 1),
            '%' => (Part::Percent, 1),
            '@' => (Part::Text, 1),
            'E' | 'e' if matches!(characters.get(index + 1), Some('+' | '-')) => {
                (Part::Exponent(characters[index + 1] == '+'), 2)
            }
            _ if starts_with(index, "General") => (Part::General, 7),
            _ if starts_with(index, "AM/PM") => {
                (Part::Meridiem(true, character.is_ascii_uppercase()), 5)
            }
            _ if starts_with(index, "A/P") => {
                (Part::Meridiem(false, character.is_ascii_uppercase()), 3)
            }
            'y' | 'Y' => (Part::Year(run), run),
            'm' | 'M' => (Part::Month(run), run),
            'd' | 'D' => (Part::Day(run), run),
            'h' | 'H' => (Part::Hour(run), run),
            's' | 'S' => (Part::Second(run), run),
            other => (Part::Literal(other.to_string()), 1),
        };
        parts.push(part);
        index += length;
    }
    // Short `m` runs after hours or before seconds are minutes.
    let dates: Vec<usize> = (0..parts.len()).filter(|&i| parts[i].is_date()).collect();
    for (position, &index) in dates.iter().enumerate() {
        if let Part::Month(length @ 1..=2) = parts[index] {
            let after_hour = position > 0 && matches!(parts[dates[position - 1]], Part::Hour(_));
            let before_second = dates
                .get(position + 1)
                .is_some_and(|&next| matches!(parts[next], Part::Second(_)));
            if after_hour || before_second {
                parts[index] = Part::Minute(length);
            }
        }
    }
    parts
}

/// Format a number with the section chosen by its sign.
fn format_number(value: f64, code: &str, date1904: bool) -> Result<String, ExcelError> {
    let sections = sections(code);
    let (section, value, sign) = match sections.len() {
        1 => (sections[0], value, value < 0.0),
        2 if value < 0.0 => (sections[1], -value, false),
        2 => (sections[0], value, false),
        _ if value > 0.0 => (sections[0], value, false),
        _ if value < 0.0 => (sections[1], -value, false),
        _ => (sections[2], value, false),
    };
    let parts = parse(section);
    if parts.iter().any(Part::is_date) {
        return format_date(value, &parts, date1904).ok_or(ExcelError::Value);
    }
    let mut formatted = format_digits(value.abs(), &parts);
    if sign {
        formatted.insert(0, '-');
    }
    Ok(formatted)
}

fn format_date(value: f64, parts: &[Part], date1904: bool) -> Option<String> {
    let (mut serial, fraction) = text_serial(value, date1904)?;
    let mut seconds = (fraction * 86_400.0).round() as u32;
    if seconds == 86_400 {
        serial += 1;
        seconds = 0;
    }
    let (year, month, day) = civil(serial, date1904);
    let twelve_hour = parts.iter().any(|part| matches!(part, Part::Meridiem(..)));
    let hour = seconds / 3600;
    let padded = |number: u32, length: usize| {
        if length >= 2 {
            format!("{number:02}")
        } else {
            number.to_string()
        }
    };
    let mut formatted = String::new();
    for part in parts {
        match part {
            Part::Year(length) if *length <= 2 => {
                formatted.push_str(&format!("{:02}", year.rem_euclid(100)))
            }
            Part::Year(_) => formatted.push_str(&format!("{year:04}")),
            Part::Month(length) => {
                let name = MONTHS[month as usize - 1];
                match length {
                    1 | 2 => formatted.push_str(&padded(month, *length)),
                    3 => formatted.push_str(&name[..3]),
                    4 => formatted.push_str(name),
                    _ => formatted.push_str(&name[..1]),
                }
            }
            Part::Day(length @ 1..=2) => formatted.push_str(&padded(day, *length)),
            Part::Day(length) => {
                let name = WEEKDAYS[weekday(serial, date1904) as usize];
                formatted.push_str(if *length == 3 { &name[..3] } else { name });
            }
            Part::Hour(length) => {
                let hour = match (twelve_hour, hour % 12) {
                    (true, 0) => 12,
                    (true, hour) => hour,
                    (false, _) => hour,
                };
                formatted.push_str(&padded(hour, *length));
            }
            Part::Minute(length) => formatted.push_str(&padded(seconds / 60 % 60, *length)),
            Part::Second(length) => formatted.push_str(&padded(seconds % 60, *length)),
            Part::Meridiem(long, upper) => {
                let text = match (hour < 12, long) {
                    (true, true) => "AM",
                    (false, true) => "PM",
                    (true, false) => "A",
                    (false, false) => "P",
                };
                if *upper {
                    formatted.push_str(text);
                } else {
                    formatted.push_str(&text.to_lowercase());
                }
            }
            Part::Literal(text) => formatted.push_str(text),
            Part::Digit(digit) => formatted.push(*digit),
            Part::Point => formatted.push('.'),
            Part::Comma => formatted.push(','),
            Part::Percent => formatted.push('%'),
            Part::Exponent(plus) => formatted.push_str(if *plus { "E+" } else { "E-" }),
            Part::General | Part::Text => formatted.push_str(&general(value)),
        }
    }
    Some(formatted)
}

/// The placeholders among `parts`.
fn placeholders(parts: &[Part]) -> Vec<char> {
    parts
        .iter()
        .filter_map(|part| match part {
            Part::Digit(digit) => Some(*digit),
            _ => None,
        })
        .collect()
}

/// Lay out a non-negative number through digit placeholders.
fn format_digits(value: f64, parts: &[Part]) -> String {
    let exponent_at = parts
        .iter()
        .position(|part| matches!(part, Part::Exponent(_)));
    let mantissa_end = exponent_at.unwrap_or(parts.len());
    let integer_end = parts[..mantissa_end]
        .iter()
        .position(|part| *part == Part::Point)
        .unwrap_or(mantissa_end);
    let integer_places = placeholders(&parts[..integer_end]);
    let decimals = placeholders(&parts[integer_end..mantissa_end]).len();
    let exponent_places = placeholders(&parts[mantissa_end..]);
    if integer_places.is_empty() && decimals == 0 && parts.contains(&Part::General) {
        return parts
            .iter()
            .map(|part| match part {
                Part::General => general(value),
                Part::Literal(text) => text.clone(),
                _ => String::new(),
            })
            .collect();
    }

    // A comma between integer placeholders groups thousands, a comma run
    // right after the last placeholder scales by a thousand per comma, and
    // any other comma is literal.
    let is_digit = |part: &Part| matches!(part, Part::Digit(_));
    let mut grouping = false;
    let mut scale = 0;
    let mut literal_commas = vec![false; parts.len()];
    for index in 0..parts.len() {
        if parts[index] != Part::Comma {
            continue;
        }
        if index >= mantissa_end {
            literal_commas[index] = true;
            continue;
        }
        let after_digit = parts[..index]
            .iter()
            .rposition(is_digit)
            .is_some_and(|digit| parts[digit + 1..index].iter().all(|p| *p == Part::Comma));
        let before_integer_digit = index < integer_end
            && parts[index + 1..integer_end]
                .iter()
                .find(|part| **part != Part::Comma)
                .is_some_and(is_digit);
        if before_integer_digit && parts[..index].iter().any(is_digit) {
            grouping = true;
        } else if after_digit {
            scale += 1;
        } else {
            literal_commas[index] = true;
        }
    }

    let percents = parts.iter().filter(|part| **part == Part::Percent).count();
    let mut value = value * 100_f64.powi(percents as i32) / 1000_f64.powi(scale);
    let mut exponent = 0_i32;
    if exponent_at.is_some() && value != 0.0 {
        let places = integer_places.len().max(1) as i32;
        let magnitude = value.log10().floor() as i32;
        // With several integer places, exponents step by that many.
        exponent = magnitude - magnitude.rem_euclid(places);
        let mut mantissa = round_decimal(
            value / 10_f64.powi(exponent),
            decimals as f64,
            Rounding::Nearest,
        );
        if mantissa >= 10_f64.powi(places) {
            exponent += places;
            mantissa = round_decimal(
                value / 10_f64.powi(exponent),
                decimals as f64,
                Rounding::Nearest,
            );
        }
        value = mantissa;
    } else {
        value = round_decimal(value, decimals as f64, Rounding::Nearest);
    }
    let fixed = format!("{value:.decimals$}");
    let (integer, fraction) = fixed.split_once('.').unwrap_or((&fixed, ""));
    let integer: Vec<char> = if integer == "0" && !integer_places.contains(&'0') {
        Vec::new()
    } else {
        integer.chars().collect()
    };
    let fraction: Vec<char> = fraction.chars().collect();
    let exponent_width = exponent_places
        .iter()
        .filter(|&&digit| digit == '0')
        .count();
    let exponent_digits = format!("{:0>exponent_width$}", exponent.unsigned_abs());

    let places = integer_places.len();
    let mut formatted = String::new();
    // Emit one integer digit, grouping by its place value.
    let push_digit = |formatted: &mut String, digit: char, place: usize| {
        formatted.push(digit);
        if grouping && place > 0 && place.is_multiple_of(3) {
            formatted.push(',');
        }
    };
    let mut integer_index = 0;
    let mut decimal_index = 0;
    let mut exponent_shown = exponent_places.is_empty();
    for (index, part) in parts.iter().enumerate() {
        match part {
            // The first exponent placeholder shows every exponent digit.
            Part::Digit(_) if index > mantissa_end => {
                if !std::mem::replace(&mut exponent_shown, true) {
                    formatted.push_str(&exponent_digits);
                }
            }
            Part::Digit(placeholder) if index < integer_end => {
                let place = places - 1 - integer_index;
                // The first placeholder also shows digits beyond the others.
                if integer_index == 0 {
                    let extra = integer.len().saturating_sub(places);
                    for (offset, digit) in integer[..extra].iter().enumerate() {
                        push_digit(&mut formatted, *digit, integer.len() - 1 - offset);
                    }
                }
                match integer.len().checked_sub(place + 1) {
                    Some(at) => push_digit(&mut formatted, integer[at], place),
                    None if *placeholder == '0' => push_digit(&mut formatted, '0', place),
                    None if *placeholder == '?' => formatted.push(' '),
                    None => {}
                }
                integer_index += 1;
            }
            Part::Digit(placeholder) => {
                let rest = &fraction[decimal_index.min(fraction.len())..];
                let trailing_zero = rest.iter().all(|&digit| digit == '0');
                match placeholder {
                    '#' if trailing_zero => {}
                    '?' if trailing_zero => formatted.push(' '),
                    _ => formatted.push(rest.first().copied().unwrap_or('0')),
                }
                decimal_index += 1;
            }
            Part::Point => formatted.push('.'),
            Part::Comma if literal_commas[index] => formatted.push(','),
            Part::Comma => {}
            Part::Percent => formatted.push('%'),
            Part::Exponent(plus) => {
                formatted.push('E');
                if exponent < 0 {
                    formatted.push('-');
                } else if *plus {
                    formatted.push('+');
                }
                if exponent_places.is_empty() {
                    formatted.push_str(&exponent_digits);
                }
            }
            Part::General | Part::Text => formatted.push_str(&general(value)),
            Part::Literal(text) => formatted.push_str(text),
            // Date codes never reach number layout.
            _ => {}
        }
    }
    formatted
}

/// Text through the text section, or unchanged when there is none.
fn format_string(text: &str, code: &str) -> String {
    let sections = sections(code);
    let section = match sections.get(3) {
        Some(section) => section,
        None if sections.len() == 1 && sections[0].contains('@') => sections[0],
        None => return text.to_owned(),
    };
    parse(section)
        .into_iter()
        .map(|part| match part {
            Part::Text => text.to_owned(),
            Part::Literal(literal) => literal,
            _ => String::new(),
        })
        .collect()
}

impl<'a, S: DataSource> Context<'a, S> {
    pub(super) fn format_text(&self, arguments: &[Expr]) -> Outcome {
        let value = self.scalar(&arguments[0]);
        let code = self.text(&arguments[1])?;
        let number = match value {
            Value::Number(number) => number,
            Value::Blank => 0.0,
            Value::Boolean(value) => {
                return Ok(Value::Text(if value { "TRUE" } else { "FALSE" }.into()))
            }
            Value::Text(text) => match lenient_number(&text) {
                Some(number) => number,
                None => return Ok(Value::Text(format_string(&text, &code))),
            },
            Value::Error(error) => return Err(error),
            Value::Array(_) => return Err(ExcelError::Value),
        };
        format_number(number, &code, self.source.date1904()).map(Value::Text)
    }
}

#[cfg(test)]
mod tests {
    use super::{format_number, format_string};

    fn text(value: f64, code: &str) -> String {
        format_number(value, code, false).unwrap()
    }

    #[test]
    fn number_codes_follow_excel() {
        let cases = [
            (0.256, "0%", "26%"),
            (0.256, "0.0%", "25.6%"),
            (1234567.891, "#,##0", "1,234,568"),
            (1234567.891, "#,##0.00", "1,234,567.89"),
            (1234.5, "$#,##0.00", "$1,234.50"),
            (1234.5, "\"$\"#,##0", "$1,235"),
            (1234.0, "#,##0_);(#,##0)", "1,234 "),
            (-1234.0, "#,##0_);(#,##0)", "(1,234)"),
            (12345.0, "0.00E+00", "1.23E+04"),
            (0.00012, "0.0E+0", "1.2E-4"),
            (2.71928, "0.000", "2.719"),
            (2.5, "0", "3"),
            (-2.5, "0", "-3"),
            (0.5, "#.##", ".5"),
            (0.0, "#,###", ""),
            (5.0, "000", "005"),
            (5.0, "0,000", "0,005"),
            (1234567.0, "#,##0,", "1,235"),
            (1234567.0, "0.0,,", "1.2"),
            (7.0, "0;(0);\"zero\"", "7"),
            (0.0, "0;(0);\"zero\"", "zero"),
            (-7.0, "0;(0)", "(7)"),
            (42.0, "General", "42"),
            (0.1, "General", "0.1"),
            (2.675, "0.00", "2.68"),
            (5.0, "0 \"units\"", "5 units"),
            (123.0, "000-000", "000-123"),
        ];
        for (value, code, expected) in cases {
            assert_eq!(text(value, code), expected, "TEXT({value}, {code:?})");
        }
    }

    #[test]
    fn date_codes_follow_excel() {
        // 2024-03-05 18:07:09 is serial 45356 plus that fraction of a day.
        let serial = 45356.0 + (18.0 * 3600.0 + 7.0 * 60.0 + 9.0) / 86_400.0;
        let cases = [
            ("yyyy-mm-dd", "2024-03-05"),
            ("yyyy/mm/dd", "2024/03/05"),
            ("mm/dd/yyyy", "03/05/2024"),
            ("d-mmm-yy", "5-Mar-24"),
            ("d-mmm-yyyy", "5-Mar-2024"),
            ("mmm-yy", "Mar-24"),
            ("mmmm d, yyyy", "March 5, 2024"),
            ("dddd", "Tuesday"),
            ("ddd", "Tue"),
            ("hh:mm:ss", "18:07:09"),
            ("h:mm AM/PM", "6:07 PM"),
            ("h:mm a/p", "6:07 p"),
            ("m/d/yy h:mm", "3/5/24 18:07"),
        ];
        for (code, expected) in cases {
            assert_eq!(text(serial, code), expected, "TEXT(serial, {code:?})");
        }
        assert_eq!(
            format_number(0.0, "yyyy-mm-dd", true).unwrap(),
            "1904-01-01"
        );
        assert!(format_number(-1.0, "yyyy-mm-dd", false).is_err());
    }

    #[test]
    fn text_values_use_the_text_section() {
        assert_eq!(format_string("abc", "0.00"), "abc");
        assert_eq!(format_string("abc", "\"<\"@\">\""), "<abc>");
        assert_eq!(format_string("abc", "0;-0;0;\"text: \"@"), "text: abc");
    }
}
