//! Text functions other than `TEXT`. Positions and lengths count characters.

use super::dates::parse_date_time;
use super::{condition, lenient_number, text, Context, Operand, Outcome};
use crate::eval::{DataSource, ExcelError, Expr, Value};

/// Longest text a cell holds.
const MAX_TEXT: usize = 32_767;

/// A character count or position: truncated, and `#VALUE!` when negative.
fn count(value: f64) -> Result<usize, ExcelError> {
    if value < 0.0 {
        Err(ExcelError::Value)
    } else {
        Ok(value.trunc().min(MAX_TEXT as f64 + 1.0) as usize)
    }
}

/// Text no longer than a cell holds, else `#VALUE!`.
fn bounded(text: String) -> Outcome {
    if text.chars().count() > MAX_TEXT {
        return Err(ExcelError::Value);
    }
    Ok(Value::Text(text))
}

impl<'a, S: DataSource> Context<'a, S> {
    pub(super) fn concatenate(&self, arguments: &[Expr]) -> Outcome {
        let mut joined = String::new();
        for argument in arguments {
            joined.push_str(&self.text(argument)?);
        }
        bounded(joined)
    }

    /// Every value of an argument in row-major order: a range or array
    /// gives each of its values, blanks included, and a scalar itself.
    fn flattened(&self, expression: &Expr) -> Vec<Value> {
        let value = match self.operand(expression) {
            Operand::Range(range) => self.read(Operand::Range(range)),
            Operand::Value(value) => value,
        };
        match value {
            Value::Array(rows) => rows.into_iter().flatten().collect(),
            value => vec![value],
        }
    }

    /// `CONCAT`: every value of every argument joined, ranges included.
    pub(super) fn concat(&self, arguments: &[Expr]) -> Outcome {
        let mut joined = String::new();
        for argument in arguments {
            for value in self.flattened(argument) {
                joined.push_str(&text(value)?);
            }
        }
        bounded(joined)
    }

    /// `TEXTJOIN`: the values of the text arguments joined by the
    /// delimiter, or by each value of a delimiter range in turn. With
    /// `ignore_empty`, blanks and empty text are left out.
    pub(super) fn text_join(&self, arguments: &[Expr]) -> Outcome {
        let mut delimiters = self
            .flattened(&arguments[0])
            .into_iter()
            .map(text)
            .collect::<Result<Vec<_>, _>>()?;
        if delimiters.is_empty() {
            delimiters.push(String::new());
        }
        let ignore_empty = condition(self.scalar(&arguments[1]))?;
        let mut joined = String::new();
        let mut count = 0_usize;
        for argument in &arguments[2..] {
            for value in self.flattened(argument) {
                let value = text(value)?;
                if ignore_empty && value.is_empty() {
                    continue;
                }
                if count > 0 {
                    joined.push_str(&delimiters[(count - 1) % delimiters.len()]);
                }
                joined.push_str(&value);
                count += 1;
                if joined.len() > MAX_TEXT * 4 {
                    return Err(ExcelError::Value);
                }
            }
        }
        bounded(joined)
    }

    /// `VALUE`: text read as a number, a percentage, a date, or a time.
    pub(super) fn number_value(&self, arguments: &[Expr]) -> Outcome {
        match self.scalar(&arguments[0]) {
            Value::Number(number) => Ok(Value::Number(number)),
            Value::Blank => Ok(Value::Number(0.0)),
            Value::Error(error) => Err(error),
            Value::Boolean(_) | Value::Array(_) => Err(ExcelError::Value),
            Value::Text(text) => {
                let trimmed = text.trim();
                lenient_number(trimmed)
                    .or_else(|| {
                        trimmed
                            .strip_suffix('%')
                            .and_then(lenient_number)
                            .map(|number| number / 100.0)
                    })
                    .or_else(|| parse_date_time(trimmed, self.source.date1904()))
                    .map(Value::Number)
                    .ok_or(ExcelError::Value)
            }
        }
    }

    pub(super) fn exact(&self, arguments: &[Expr]) -> Outcome {
        Ok(Value::Boolean(
            self.text(&arguments[0])? == self.text(&arguments[1])?,
        ))
    }

    /// Case-sensitive position of the first occurrence at or after `start`.
    pub(super) fn find(&self, arguments: &[Expr]) -> Outcome {
        let needle = self.text(&arguments[0])?;
        let haystack = self.text(&arguments[1])?;
        let start = self.number_or(arguments, 2, 1.0)?.trunc();
        let length = haystack.chars().count();
        if start < 1.0 || start > length as f64 + 1.0 {
            return Err(ExcelError::Value);
        }
        let skip = start as usize - 1;
        let offset = haystack
            .char_indices()
            .nth(skip)
            .map_or(haystack.len(), |(offset, _)| offset);
        let found = haystack[offset..].find(&needle).ok_or(ExcelError::Value)?;
        let position = skip + haystack[offset..offset + found].chars().count() + 1;
        Ok(Value::Number(position as f64))
    }

    pub(super) fn left(&self, arguments: &[Expr]) -> Outcome {
        let text = self.text(&arguments[0])?;
        let take = count(self.number_or(arguments, 1, 1.0)?)?;
        Ok(Value::Text(text.chars().take(take).collect()))
    }

    pub(super) fn right(&self, arguments: &[Expr]) -> Outcome {
        let text = self.text(&arguments[0])?;
        let take = count(self.number_or(arguments, 1, 1.0)?)?;
        let skip = text.chars().count().saturating_sub(take);
        Ok(Value::Text(text.chars().skip(skip).collect()))
    }

    pub(super) fn mid(&self, arguments: &[Expr]) -> Outcome {
        let text = self.text(&arguments[0])?;
        let start = self.number(&arguments[1])?.trunc();
        let take = count(self.number(&arguments[2])?)?;
        if start < 1.0 {
            return Err(ExcelError::Value);
        }
        let skip = start.min(MAX_TEXT as f64 + 1.0) as usize - 1;
        Ok(Value::Text(text.chars().skip(skip).take(take).collect()))
    }

    pub(super) fn len(&self, arguments: &[Expr]) -> Outcome {
        let text = self.text(&arguments[0])?;
        Ok(Value::Number(text.chars().count() as f64))
    }

    pub(super) fn rept(&self, arguments: &[Expr]) -> Outcome {
        let text = self.text(&arguments[0])?;
        let times = count(self.number(&arguments[1])?)?;
        if text.chars().count().saturating_mul(times) > MAX_TEXT {
            return Err(ExcelError::Value);
        }
        Ok(Value::Text(text.repeat(times)))
    }

    /// Replace every occurrence of the old text, or only the given one.
    pub(super) fn substitute(&self, arguments: &[Expr]) -> Outcome {
        let text = self.text(&arguments[0])?;
        let old = self.text(&arguments[1])?;
        let new = self.text(&arguments[2])?;
        let instance = match arguments.get(3) {
            Some(argument) => {
                let instance = self.number(argument)?.trunc();
                if instance < 1.0 {
                    return Err(ExcelError::Value);
                }
                Some(instance)
            }
            None => None,
        };
        if old.is_empty() {
            return Ok(Value::Text(text));
        }
        let replaced = match instance {
            None => text.replace(&old, &new),
            Some(instance) => match text.match_indices(&old).nth(instance as usize - 1) {
                Some((offset, _)) => {
                    let mut replaced = String::with_capacity(text.len() + new.len());
                    replaced.push_str(&text[..offset]);
                    replaced.push_str(&new);
                    replaced.push_str(&text[offset + old.len()..]);
                    replaced
                }
                None => text,
            },
        };
        if replaced.chars().count() > MAX_TEXT {
            return Err(ExcelError::Value);
        }
        Ok(Value::Text(replaced))
    }

    /// Drop leading and trailing spaces and collapse inner runs to one.
    pub(super) fn trim(&self, arguments: &[Expr]) -> Outcome {
        let text = self.text(&arguments[0])?;
        let mut trimmed = String::with_capacity(text.len());
        for word in text.split(' ').filter(|word| !word.is_empty()) {
            if !trimmed.is_empty() {
                trimmed.push(' ');
            }
            trimmed.push_str(word);
        }
        Ok(Value::Text(trimmed))
    }
}
