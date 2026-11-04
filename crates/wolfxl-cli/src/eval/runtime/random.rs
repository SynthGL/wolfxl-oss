//! Non-cryptographic worksheet randomness, independent of stored workbook values.

use std::collections::hash_map::RandomState;
use std::hash::{BuildHasher, Hasher};
use std::sync::atomic::{AtomicU64, Ordering};

use super::{finite, Context, Outcome};
use crate::eval::{DataSource, ExcelError, Expr, Function, Value};

thread_local! {
    // RandomState supplies process-local random keys through the standard
    // library on every supported platform; no clock seed or fixed sequence.
    static STATE: RandomState = RandomState::new();
}

static NEXT: AtomicU64 = AtomicU64::new(0);

/// A 53-bit fraction in [0, 1). Worksheet randomness is not a security API.
fn fraction() -> f64 {
    STATE.with(|state| {
        let mut hasher = state.build_hasher();
        hasher.write_u64(NEXT.fetch_add(1, Ordering::Relaxed));
        (hasher.finish() >> 11) as f64 * (1.0 / 9_007_199_254_740_992.0)
    })
}

impl<'a, S: DataSource> Context<'a, S> {
    pub(super) fn random(&self, function: Function, arguments: &[Expr]) -> Outcome {
        if function == Function::Rand {
            return Ok(Value::Number(fraction()));
        }
        let bottom = self.number(&arguments[0])?;
        let top = self.number(&arguments[1])?;
        if !bottom.is_finite() || !top.is_finite() || bottom > top {
            return Err(ExcelError::Num);
        }
        // Excel returns ceil(bottom) when ordered fractional bounds contain
        // no integer, as captured for RANDBETWEEN(1.1, 1.9).
        let bottom = bottom.ceil();
        let top = top.floor().max(bottom);
        if bottom == top {
            return Ok(Value::Number(bottom));
        }
        let width = top - bottom + 1.0;
        if !width.is_finite() {
            return Err(ExcelError::Num);
        }
        // Clamp only rounding at the upper endpoint of a very wide interval.
        Ok(finite((bottom + (fraction() * width).floor()).min(top)))
    }
}
