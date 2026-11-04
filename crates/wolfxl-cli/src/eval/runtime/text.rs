//! Text functions other than `TEXT`. Positions and lengths count characters.

use super::{Context, Outcome};
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

impl<'a, S: DataSource> Context<'a, S> {
    pub(super) fn concatenate(&self, arguments: &[Expr]) -> Outcome {
        let mut joined = String::new();
        for argument in arguments {
            joined.push_str(&self.text(argument)?);
        }
        if joined.chars().count() > MAX_TEXT {
            return Err(ExcelError::Value);
        }
        Ok(Value::Text(joined))
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
