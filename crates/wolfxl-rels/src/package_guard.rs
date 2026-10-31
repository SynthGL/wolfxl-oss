//! Shared resource policy for OOXML package admission.

pub const DEFAULT_MAX_ZIP_ENTRIES: usize = 200_000;
pub const DEFAULT_MAX_ZIP_ENTRY_BYTES: u64 = 512 * 1024 * 1024;
pub const DEFAULT_MAX_ZIP_TOTAL_BYTES: u64 = 4 * 1024 * 1024 * 1024;
pub const DEFAULT_MAX_COMPRESSION_RATIO: u64 = 1_000;
pub const DEFAULT_MAX_CELLS: u64 = 8_000_000;
pub const DEFAULT_MAX_SHARED_STRING_BYTES: u64 = 256 * 1024 * 1024;
pub const DEFAULT_MAX_MATERIALIZED_STRING_BYTES: u64 = 512 * 1024 * 1024;

/// Limits are workbook-wide, not per worksheet. Bytes mean decoded UTF-8 bytes.
#[derive(Debug, Clone, Copy)]
pub struct Limits {
    pub max_package_bytes: u64,
    pub max_part_bytes: u64,
    pub max_cells: u64,
    pub max_shared_string_bytes: u64,
    pub max_materialized_string_bytes: u64,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_package_bytes: DEFAULT_MAX_ZIP_TOTAL_BYTES,
            max_part_bytes: DEFAULT_MAX_ZIP_ENTRY_BYTES,
            max_cells: DEFAULT_MAX_CELLS,
            max_shared_string_bytes: DEFAULT_MAX_SHARED_STRING_BYTES,
            max_materialized_string_bytes: DEFAULT_MAX_MATERIALIZED_STRING_BYTES,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LimitExceeded {
    pub name: &'static str,
    pub observed: u64,
    pub configured: u64,
}

impl std::fmt::Display for LimitExceeded {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{} exceeded: observed {}, configured {}; adjust load_workbook(..., limits=Limits({}=...))", self.name, self.observed, self.configured, self.name)
    }
}

impl std::error::Error for LimitExceeded {}

pub fn check(name: &'static str, observed: u64, configured: u64) -> Result<(), LimitExceeded> {
    if observed > configured {
        Err(LimitExceeded { name, observed, configured })
    } else {
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn budgets_accept_boundary_and_reject_next_byte() {
        assert!(check("max_cells", 0, 0).is_ok());
        assert!(check("max_cells", 12, 12).is_ok());
        let error = check("max_cells", 13, 12).unwrap_err();
        assert_eq!(error.observed, 13);
        assert_eq!(error.configured, 12);
        assert_eq!(
            error.to_string(),
            "max_cells exceeded: observed 13, configured 12; adjust load_workbook(..., limits=Limits(max_cells=...))"
        );
    }
}
