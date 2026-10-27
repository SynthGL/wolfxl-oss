//! `PivotCache` — the cache-definition half of an OOXML pivot.
//!
//! Mirrors the §10 contract of RFC-047. One `PivotCache` represents
//! both `xl/pivotCache/pivotCacheDefinition{N}.xml` and the
//! companion `xl/pivotCache/pivotCacheRecords{N}.xml`; the records
//! data lives in `model::records::CacheRecord` and is held in the
//! `records` field below.

use crate::{DEFAULT_CACHE_VERSION, DEFAULT_MIN_REFRESHABLE_VERSION};

/// Source range pointer. Either `(sheet, ref)` or a defined `name`,
/// not both. Validated at the §10 boundary (`PivotCache::new` enforces).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WorksheetSource {
    /// Sheet name; non-empty unless `name` is set.
    pub sheet: String,
    /// A1-style range, e.g. `"A1:D100"`. Non-empty unless `name` is set.
    pub range: String,
    /// Defined-name reference. When set, `sheet` and `range` are empty.
    pub name: Option<String>,
}

/// One source-data column's schema. Mirrors RFC-047 §10.3.
#[derive(Debug, Clone, PartialEq)]
pub struct CacheField {
    pub name: String,
    pub num_fmt_id: u32,
    pub data_type: DataType,
    pub shared_items: SharedItems,
    /// Calculated-field formula. Pinned to `None` for v2.0 (out of
    /// scope; see RFC-047 §11). Reserved for v2.1.
    pub formula: Option<String>,
    /// OLAP hierarchy index. Pinned to `None` for v2.0; out
    /// permanently per RFC-047 §11.
    pub hierarchy: Option<i32>,
}

/// Per-field unique-value enumeration + type flags. Mirrors RFC-047
/// §10.4 / §10.5.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct SharedItems {
    /// Number of `items` (mirrors the `count` attr). `None` →
    /// suppress the attribute entirely (numeric-only no-enumeration
    /// form).
    pub count: Option<u32>,
    /// Enumerated unique values. `None` → no `<sharedItems>`
    /// children, only attribute flags. RFC-047 §10.4 numeric-only
    /// no-enumeration path.
    pub items: Option<Vec<CacheValue>>,
    pub contains_blank: bool,
    pub contains_mixed_types: bool,
    pub contains_semi_mixed_types: bool,
    pub contains_string: bool,
    pub contains_number: bool,
    pub contains_integer: bool,
    pub contains_date: bool,
    pub contains_non_date: bool,
    pub min_value: Option<f64>,
    pub max_value: Option<f64>,
    pub min_date: Option<String>,
    pub max_date: Option<String>,
    pub long_text: bool,
}

/// Inferred type of a source-data column. RFC-047 §10.3.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DataType {
    String,
    Number,
    Date,
    Bool,
    /// Mixed types in a single column. Sets `contains_semi_mixed_types`.
    Mixed,
}

/// A single shared-or-inline value in a SharedItems list or a record
/// cell. RFC-047 §10.5.
#[derive(Debug, Clone, PartialEq)]
pub enum CacheValue {
    String(String),
    Number(f64),
    Boolean(bool),
    /// ISO 8601 date-time string, e.g. `"2026-01-15T00:00:00"`.
    Date(String),
    /// `<m/>` — missing.
    Missing,
    /// `<e v="..."/>` — error string like `"#REF!"`, `"#NAME?"`.
    Error(String),
}

/// RFC-061 §10.3 — calculated cache field.
///
/// Emitted as an extra `<cacheField formula="…" databaseField="0"/>`
/// after the source fields, so its cache-field index is
/// `fields.len() + position`. Records carry no cells for it. Excel
/// evaluates the formula itself.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CalculatedField {
    pub name: String,
    pub formula: String,
    /// Data type hint: "string" | "number" | "boolean" | "date".
    pub data_type: String,
}

/// Calculated item on a source field (cache-scoped; every pivot
/// table on the cache shows it).
///
/// The item itself is a trailing shared item of `field_index`
/// at `item_index`; this entry carries its formula and is emitted
/// in the cache definition's `<calculatedItems>` block.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CalculatedItem {
    pub field_index: u32,
    pub item_index: u32,
    pub formula: String,
}

/// RFC-061 §10.5 — typed cache-side field group (date / range /
/// discrete grouping).
#[derive(Debug, Clone, PartialEq)]
pub struct FieldGroup {
    /// 0-based cache field index this group applies to.
    pub field_index: u32,
    /// Parent field index (for recursive grouping). v2.0 caps depth
    /// at 4.
    pub parent_index: Option<u32>,
    pub kind: FieldGroupKind,
    pub date: Option<DateGroup>,
    pub range: Option<RangeGroup>,
    /// Synthesized item names (the labels Excel shows in the field
    /// header for each grouped bucket).
    pub items: Vec<String>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FieldGroupKind {
    Date,
    Range,
    Discrete,
}

impl FieldGroupKind {
    pub fn xml_value(&self) -> &'static str {
        match self {
            FieldGroupKind::Date => "date",
            FieldGroupKind::Range => "range",
            FieldGroupKind::Discrete => "discrete",
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DateGroup {
    /// One of "years"|"quarters"|"months"|"days"|"hours"|"minutes"|"seconds".
    pub group_by: String,
    pub start_date: String,
    pub end_date: String,
}

#[derive(Debug, Clone, PartialEq)]
pub struct RangeGroup {
    pub start: f64,
    pub end: f64,
    pub interval: f64,
}

/// Top-level pivot cache. Holds both the definition (fields) and the
/// records snapshot.
#[derive(Debug, Clone, PartialEq)]
pub struct PivotCache {
    /// 0-based cache id, allocated by `Workbook.add_pivot_cache(...)`.
    /// Foreign key referenced by `<pivotTableDefinition cacheId="…"/>`.
    pub cache_id: u32,
    pub source: WorksheetSource,
    pub fields: Vec<CacheField>,
    /// Records snapshot. May be empty when refresh-on-load is `true`
    /// (deferred Option-B path), but Sprint Ν Option A pins this
    /// non-empty.
    pub records: Vec<crate::model::records::CacheRecord>,
    pub refresh_on_load: bool,
    pub refreshed_version: u8,
    pub created_version: u8,
    pub min_refreshable_version: u8,
    pub refreshed_by: String,
    /// Path to the records part inside the ZIP (e.g.
    /// `"pivotCacheRecords1.xml"`). Set by the patcher when it
    /// allocates the part-id; emit uses this as the `<r:id>` target
    /// resolution for the cache → records rel.
    pub records_part_path: Option<String>,
    /// RFC-061 §2.2 — calculated cache fields (cache-scoped).
    pub calculated_fields: Vec<CalculatedField>,
    /// Calculated items on source fields (cache-scoped).
    pub calculated_items: Vec<CalculatedItem>,
    /// RFC-061 §2.4 — cache-side field groups (date / range).
    pub field_groups: Vec<FieldGroup>,
}

impl PivotCache {
    /// Build a new cache with sensible defaults from RFC-047 §10.1.
    pub fn new(cache_id: u32, source: WorksheetSource, fields: Vec<CacheField>) -> Self {
        Self {
            cache_id,
            source,
            fields,
            records: Vec::new(),
            refresh_on_load: false,
            refreshed_version: DEFAULT_CACHE_VERSION,
            created_version: DEFAULT_CACHE_VERSION,
            min_refreshable_version: DEFAULT_MIN_REFRESHABLE_VERSION,
            refreshed_by: "wolfxl".to_string(),
            records_part_path: None,
            calculated_fields: Vec::new(),
            calculated_items: Vec::new(),
            field_groups: Vec::new(),
        }
    }

    /// RFC-047 §10.8 validation. Called from the §10 dict→model
    /// converter in `src/wolfxl/`. Returns `Err(message)` on the
    /// first violation.
    pub fn validate(&self) -> Result<(), String> {
        if self.fields.is_empty() {
            return Err("PivotCache requires ≥1 source field".into());
        }
        let has_sheet_ref = !self.source.sheet.is_empty() || !self.source.range.is_empty();
        let has_name = self.source.name.is_some();
        match (has_sheet_ref, has_name) {
            (false, false) => {
                return Err("PivotCache.source requires sheet+ref or name".into());
            }
            (true, true) => {
                return Err("sheet+ref and name are mutually exclusive".into());
            }
            _ => {}
        }
        // Unique field names, calculated fields included.
        let mut seen: Vec<&str> =
            Vec::with_capacity(self.fields.len() + self.calculated_fields.len());
        let names = self
            .fields
            .iter()
            .map(|f| f.name.as_str())
            .chain(self.calculated_fields.iter().map(|f| f.name.as_str()));
        for name in names {
            if seen.contains(&name) {
                return Err(format!("duplicate cache field name: {name}"));
            }
            seen.push(name);
        }
        for ci in &self.calculated_items {
            let items = self
                .fields
                .get(ci.field_index as usize)
                .and_then(|f| f.shared_items.items.as_ref());
            if items.is_none_or(|items| ci.item_index as usize >= items.len()) {
                return Err(format!(
                    "calculated item {} of field {} has no shared item",
                    ci.item_index, ci.field_index
                ));
            }
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn validate_empty_fields_rejects() {
        let pc = PivotCache::new(
            0,
            WorksheetSource {
                sheet: "Sheet1".into(),
                range: "A1:D100".into(),
                name: None,
            },
            vec![],
        );
        assert!(pc.validate().is_err());
    }

    #[test]
    fn validate_no_source_rejects() {
        let pc = PivotCache::new(
            0,
            WorksheetSource {
                sheet: "".into(),
                range: "".into(),
                name: None,
            },
            vec![dummy_field("a")],
        );
        assert!(pc.validate().is_err());
    }

    #[test]
    fn validate_dual_source_rejects() {
        let pc = PivotCache::new(
            0,
            WorksheetSource {
                sheet: "Sheet1".into(),
                range: "A1:D100".into(),
                name: Some("MyRange".into()),
            },
            vec![dummy_field("a")],
        );
        assert!(pc.validate().is_err());
    }

    #[test]
    fn validate_duplicate_field_rejects() {
        let pc = PivotCache::new(
            0,
            WorksheetSource {
                sheet: "Sheet1".into(),
                range: "A1:D100".into(),
                name: None,
            },
            vec![dummy_field("a"), dummy_field("a")],
        );
        assert!(pc.validate().is_err());
    }

    #[test]
    fn validate_happy_path() {
        let pc = PivotCache::new(
            0,
            WorksheetSource {
                sheet: "Sheet1".into(),
                range: "A1:D100".into(),
                name: None,
            },
            vec![dummy_field("a"), dummy_field("b")],
        );
        assert!(pc.validate().is_ok());
    }

    fn dummy_field(name: &str) -> CacheField {
        CacheField {
            name: name.into(),
            num_fmt_id: 0,
            data_type: DataType::String,
            shared_items: SharedItems::default(),
            formula: None,
            hierarchy: None,
        }
    }
}
