//! `IFS`, `SWITCH`, and `IFNA`. Like `IF`, each evaluates only the argument
//! it returns.

use std::cmp::Ordering;

use super::{compare, condition, Context, Operand};
use crate::eval::{DataSource, ExcelError, Expr, Value};

impl<'a, S: DataSource> Context<'a, S> {
    /// The value after the first true condition, or `#N/A` when none is.
    pub(super) fn ifs(&self, arguments: &[Expr]) -> Operand<'a> {
        for [test, value] in arguments.as_chunks::<2>().0 {
            match condition(self.scalar(test)) {
                Ok(true) => return self.result(self.operand(value)),
                Ok(false) => {}
                Err(error) => return Operand::Value(Value::Error(error)),
            }
        }
        Operand::Value(Value::Error(ExcelError::Na))
    }

    /// The result paired with the first value equal to the expression, the
    /// trailing default when no value matches, or `#N/A` without one.
    pub(super) fn switch(&self, arguments: &[Expr]) -> Operand<'a> {
        let expression = self.scalar(&arguments[0]);
        if let Value::Error(error) = expression {
            return Operand::Value(Value::Error(error));
        }
        let (pairs, default) = arguments[1..].as_chunks::<2>();
        for [candidate, result] in pairs {
            match compare(&expression, &self.scalar(candidate)) {
                Ok(Ordering::Equal) => return self.result(self.operand(result)),
                Ok(_) => {}
                Err(error) => return Operand::Value(Value::Error(error)),
            }
        }
        match default {
            [default] => self.result(self.operand(default)),
            _ => Operand::Value(Value::Error(ExcelError::Na)),
        }
    }

    /// The first argument unless it is `#N/A`, then the second.
    pub(super) fn if_na(&self, arguments: &[Expr]) -> Operand<'a> {
        let first = self.result(self.operand(&arguments[0]));
        if matches!(first, Operand::Value(Value::Error(ExcelError::Na))) {
            self.result(self.operand(&arguments[1]))
        } else {
            first
        }
    }
}
