//! Emit `xl/pivotCache/pivotCacheDefinition{N}.xml`.
//!
//! See RFC-047 §2.2 for the target XML skeleton and §10.1-§10.5 for
//! the in-memory model.

use super::{esc_attr, fmt_num, push_attr, push_attr_if, xml_decl};
use crate::model::cache::{
    CacheField, CacheValue, CalculatedField, CalculatedItem, DateGroup, FieldGroup, FieldGroupKind,
    PivotCache, RangeGroup, SharedItems, WorksheetSource,
};

/// Emit the pivotCacheDefinition XML. The `r:id` for the records part
/// must be supplied by the patcher (rels graph allocates it before
/// emit). When `records_rid` is `None`, the `<r:id>` attribute is
/// omitted — used in tests and by callers that wire the rel later.
pub fn pivot_cache_definition_xml(pc: &PivotCache, records_rid: Option<&str>) -> Vec<u8> {
    let mut out = String::with_capacity(2048);
    xml_decl(&mut out);

    // <pivotCacheDefinition …>
    out.push_str("<pivotCacheDefinition");
    push_attr(&mut out, "xmlns", crate::ns::SPREADSHEETML);
    push_attr(&mut out, "xmlns:r", crate::ns::RELATIONSHIPS);
    if let Some(rid) = records_rid {
        push_attr(&mut out, "r:id", rid);
    }
    push_attr(
        &mut out,
        "refreshOnLoad",
        if pc.refresh_on_load { "1" } else { "0" },
    );
    push_attr(&mut out, "refreshedBy", &pc.refreshed_by);
    push_attr(&mut out, "createdVersion", &pc.created_version.to_string());
    push_attr(
        &mut out,
        "refreshedVersion",
        &pc.refreshed_version.to_string(),
    );
    push_attr(
        &mut out,
        "minRefreshableVersion",
        &pc.min_refreshable_version.to_string(),
    );
    push_attr(&mut out, "recordCount", &pc.records.len().to_string());
    out.push('>');

    emit_cache_source(&mut out, &pc.source);
    emit_cache_fields(&mut out, pc);
    if !pc.calculated_items.is_empty() {
        emit_calculated_items(&mut out, &pc.calculated_items);
    }

    out.push_str("</pivotCacheDefinition>");
    out.into_bytes()
}

fn emit_cache_source(out: &mut String, src: &WorksheetSource) {
    out.push_str("<cacheSource");
    push_attr(out, "type", "worksheet");
    out.push('>');

    out.push_str("<worksheetSource");
    if let Some(name) = &src.name {
        push_attr(out, "name", name);
    } else {
        push_attr(out, "ref", &src.range);
        push_attr(out, "sheet", &src.sheet);
    }
    out.push_str("/>");
    out.push_str("</cacheSource>");
}

/// `<cacheFields>`: the source fields, then one formula-only field per
/// calculated field (`databaseField="0"`, no shared items).
fn emit_cache_fields(out: &mut String, pc: &PivotCache) {
    out.push_str("<cacheFields");
    push_attr(
        out,
        "count",
        &(pc.fields.len() + pc.calculated_fields.len()).to_string(),
    );
    out.push('>');
    for (idx, f) in pc.fields.iter().enumerate() {
        // Collect ALL groups keyed off this cache field index. For
        // recursive grouping (year → quarter → month) Excel emits
        // multiple <fieldGroup> elements nested in the same
        // cacheField, with `par=` pointing at the previous one.
        let field_groups: Vec<&FieldGroup> = pc
            .field_groups
            .iter()
            .filter(|g| g.field_index == idx as u32)
            .collect();
        let calculated: Vec<u32> = pc
            .calculated_items
            .iter()
            .filter(|ci| ci.field_index == idx as u32)
            .map(|ci| ci.item_index)
            .collect();
        emit_cache_field(out, f, &field_groups, &calculated);
    }
    for cf in &pc.calculated_fields {
        emit_calculated_field(out, cf);
    }
    out.push_str("</cacheFields>");
}

/// `calculated` lists the shared-item indices that are calculated items.
fn emit_cache_field(out: &mut String, f: &CacheField, groups: &[&FieldGroup], calculated: &[u32]) {
    out.push_str("<cacheField");
    push_attr(out, "name", &f.name);
    push_attr(out, "numFmtId", &f.num_fmt_id.to_string());
    out.push('>');
    emit_shared_items(out, &f.shared_items, calculated);
    for g in groups {
        emit_field_group(out, g);
    }
    out.push_str("</cacheField>");
}

fn emit_field_group(out: &mut String, g: &FieldGroup) {
    out.push_str("<fieldGroup");
    if let Some(p) = g.parent_index {
        push_attr(out, "par", &p.to_string());
    }
    push_attr(out, "base", &g.field_index.to_string());
    out.push('>');
    match g.kind {
        FieldGroupKind::Date => {
            if let Some(d) = &g.date {
                emit_date_range_pr(out, d);
            }
        }
        FieldGroupKind::Range => {
            if let Some(r) = &g.range {
                emit_numeric_range_pr(out, r);
            }
        }
        FieldGroupKind::Discrete => { /* no rangePr */ }
    }
    if !g.items.is_empty() {
        out.push_str("<groupItems");
        push_attr(out, "count", &g.items.len().to_string());
        out.push('>');
        for name in &g.items {
            out.push_str("<s");
            out.push_str(" v=\"");
            esc_attr(name, out);
            out.push_str("\"/>");
        }
        out.push_str("</groupItems>");
    }
    out.push_str("</fieldGroup>");
}

fn emit_date_range_pr(out: &mut String, d: &DateGroup) {
    out.push_str("<rangePr");
    push_attr(out, "groupBy", &d.group_by);
    push_attr(out, "startDate", &d.start_date);
    push_attr(out, "endDate", &d.end_date);
    out.push_str("/>");
}

fn emit_numeric_range_pr(out: &mut String, r: &RangeGroup) {
    out.push_str("<rangePr");
    push_attr(out, "autoStart", "0");
    push_attr(out, "autoEnd", "0");
    push_attr(out, "startNum", &fmt_num(r.start));
    push_attr(out, "endNum", &fmt_num(r.end));
    push_attr(out, "groupInterval", &fmt_num(r.interval));
    out.push_str("/>");
}

fn emit_calculated_field(out: &mut String, cf: &CalculatedField) {
    out.push_str("<cacheField");
    push_attr(out, "name", &cf.name);
    push_attr(out, "numFmtId", "0");
    push_attr(out, "formula", &cf.formula);
    push_attr(out, "databaseField", "0");
    out.push_str("/>");
}

/// `<calculatedItems>`: each item's formula plus a pivot area naming
/// the shared item it computes.
fn emit_calculated_items(out: &mut String, items: &[CalculatedItem]) {
    out.push_str("<calculatedItems");
    push_attr(out, "count", &items.len().to_string());
    out.push('>');
    for ci in items {
        out.push_str("<calculatedItem");
        push_attr(out, "formula", &ci.formula);
        out.push_str("><pivotArea");
        push_attr(out, "cacheIndex", "1");
        push_attr(out, "outline", "0");
        push_attr(out, "fieldPosition", "0");
        out.push_str("><references count=\"1\"><reference");
        push_attr(out, "field", &ci.field_index.to_string());
        push_attr(out, "count", "1");
        out.push_str("><x");
        push_attr(out, "v", &ci.item_index.to_string());
        out.push_str("/></reference></references></pivotArea></calculatedItem>");
    }
    out.push_str("</calculatedItems>");
}

fn emit_shared_items(out: &mut String, si: &SharedItems, calculated: &[u32]) {
    out.push_str("<sharedItems");
    // OOXML omits these attrs when they're at default value. We emit
    // explicitly for clarity in the supported flag set; openpyxl's
    // serializer behaves the same.
    push_attr_if(out, si.contains_blank, "containsBlank", "1");
    push_attr_if(out, si.contains_mixed_types, "containsMixedTypes", "1");
    // contains_semi_mixed_types defaults TRUE in OOXML; we emit `0` only
    // when explicitly false.
    push_attr_if(
        out,
        !si.contains_semi_mixed_types,
        "containsSemiMixedTypes",
        "0",
    );
    // contains_string defaults TRUE; emit `0` only when false.
    push_attr_if(out, !si.contains_string, "containsString", "0");
    push_attr_if(out, si.contains_number, "containsNumber", "1");
    push_attr_if(out, si.contains_integer, "containsInteger", "1");
    push_attr_if(out, si.contains_date, "containsDate", "1");
    // contains_non_date defaults TRUE; emit `0` only when false.
    push_attr_if(out, !si.contains_non_date, "containsNonDate", "0");
    if let Some(min) = si.min_value {
        push_attr(out, "minValue", &fmt_num(min));
    }
    if let Some(max) = si.max_value {
        push_attr(out, "maxValue", &fmt_num(max));
    }
    if let Some(min) = &si.min_date {
        push_attr(out, "minDate", min);
    }
    if let Some(max) = &si.max_date {
        push_attr(out, "maxDate", max);
    }
    push_attr_if(out, si.long_text, "longText", "1");

    if let Some(items) = &si.items {
        if let Some(c) = si.count {
            push_attr(out, "count", &c.to_string());
        } else {
            push_attr(out, "count", &items.len().to_string());
        }
        out.push('>');
        for (i, v) in items.iter().enumerate() {
            emit_shared_value(out, v, calculated.contains(&(i as u32)));
        }
        out.push_str("</sharedItems>");
    } else {
        if let Some(c) = si.count {
            push_attr(out, "count", &c.to_string());
        }
        out.push_str("/>");
    }
}

/// `calculated` marks a calculated item (`f="1"`).
fn emit_shared_value(out: &mut String, v: &CacheValue, calculated: bool) {
    match v {
        CacheValue::String(s) => {
            out.push_str("<s v=\"");
            esc_attr(s, out);
            out.push('"');
        }
        CacheValue::Number(n) => {
            out.push_str("<n v=\"");
            out.push_str(&fmt_num(*n));
            out.push('"');
        }
        CacheValue::Boolean(b) => {
            out.push_str("<b v=\"");
            out.push_str(if *b { "1" } else { "0" });
            out.push('"');
        }
        CacheValue::Date(d) => {
            out.push_str("<d v=\"");
            esc_attr(d, out);
            out.push('"');
        }
        CacheValue::Missing => {
            out.push_str("<m");
        }
        CacheValue::Error(s) => {
            out.push_str("<e v=\"");
            esc_attr(s, out);
            out.push('"');
        }
    }
    push_attr_if(out, calculated, "f", "1");
    out.push_str("/>");
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::cache::{DataType, SharedItems};

    fn cache_string_field(name: &str, items: Vec<&str>) -> CacheField {
        CacheField {
            name: name.into(),
            num_fmt_id: 0,
            data_type: DataType::String,
            shared_items: SharedItems {
                count: Some(items.len() as u32),
                items: Some(
                    items
                        .into_iter()
                        .map(|s| CacheValue::String(s.into()))
                        .collect(),
                ),
                contains_semi_mixed_types: true,
                contains_string: true,
                contains_non_date: true,
                ..Default::default()
            },
            formula: None,
            hierarchy: None,
        }
    }

    fn cache_number_field(name: &str, min: f64, max: f64) -> CacheField {
        CacheField {
            name: name.into(),
            num_fmt_id: 0,
            data_type: DataType::Number,
            shared_items: SharedItems {
                count: None,
                items: None,
                contains_semi_mixed_types: false,
                contains_string: false,
                contains_number: true,
                contains_integer: false,
                min_value: Some(min),
                max_value: Some(max),
                contains_non_date: true,
                ..Default::default()
            },
            formula: None,
            hierarchy: None,
        }
    }

    #[test]
    fn emit_minimal_cache_definition() {
        let pc = PivotCache::new(
            0,
            WorksheetSource {
                sheet: "Sheet1".into(),
                range: "A1:B3".into(),
                name: None,
            },
            vec![
                cache_string_field("region", vec!["North", "South"]),
                cache_number_field("revenue", 100.0, 999.0),
            ],
        );

        let xml = pivot_cache_definition_xml(&pc, Some("rId1"));
        let s = std::str::from_utf8(&xml).unwrap();
        assert!(s.starts_with("<?xml"));
        assert!(s.contains("<pivotCacheDefinition"));
        assert!(s.contains("r:id=\"rId1\""));
        assert!(s.contains("recordCount=\"0\""));
        assert!(s.contains("<worksheetSource ref=\"A1:B3\" sheet=\"Sheet1\"/>"));
        assert!(s.contains("<cacheField name=\"region\""));
        assert!(s.contains("<s v=\"North\"/>"));
        assert!(s.contains("<s v=\"South\"/>"));
        assert!(s.contains("<cacheField name=\"revenue\""));
        assert!(s.contains("containsNumber=\"1\""));
        assert!(s.contains("minValue=\"100\""));
        assert!(s.contains("maxValue=\"999\""));
    }

    #[test]
    fn emit_named_range_source() {
        let pc = PivotCache::new(
            0,
            WorksheetSource {
                sheet: String::new(),
                range: String::new(),
                name: Some("MyRange".into()),
            },
            vec![cache_string_field("a", vec!["x"])],
        );
        let xml = pivot_cache_definition_xml(&pc, None);
        let s = std::str::from_utf8(&xml).unwrap();
        assert!(s.contains("<worksheetSource name=\"MyRange\"/>"));
        assert!(!s.contains("ref="));
        assert!(!s.contains("r:id="));
    }

    #[test]
    fn emit_is_byte_stable() {
        // Determinism: same input → same output bytes.
        let pc = PivotCache::new(
            5,
            WorksheetSource {
                sheet: "Data".into(),
                range: "A1:C100".into(),
                name: None,
            },
            vec![
                cache_string_field("region", vec!["N", "S", "E", "W"]),
                cache_string_field("quarter", vec!["Q1", "Q2", "Q3", "Q4"]),
                cache_number_field("revenue", 0.0, 10000.0),
            ],
        );
        let a = pivot_cache_definition_xml(&pc, Some("rId1"));
        let b = pivot_cache_definition_xml(&pc, Some("rId1"));
        assert_eq!(a, b);
    }

    #[test]
    fn emit_with_xml_special_chars_escapes() {
        let mut field = cache_string_field("a", vec!["a&b", "c<d", "e\"f"]);
        // Restore num_fmt_id default to 0 (already there).
        field.shared_items.contains_semi_mixed_types = true;
        let pc = PivotCache::new(
            0,
            WorksheetSource {
                sheet: "S".into(),
                range: "A1:A4".into(),
                name: None,
            },
            vec![field],
        );
        let xml = pivot_cache_definition_xml(&pc, None);
        let s = std::str::from_utf8(&xml).unwrap();
        assert!(s.contains("a&amp;b"));
        assert!(s.contains("c&lt;d"));
        assert!(s.contains("e&quot;f"));
    }
}
