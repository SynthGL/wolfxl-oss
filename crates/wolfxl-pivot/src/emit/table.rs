//! Emit `xl/pivotTables/pivotTable{N}.xml`.
//!
//! See RFC-048 §2.2 + §10. Implements the v2.0 supported attr set
//! (RFC-048 §2.3). Unsupported attrs are omitted (Excel uses defaults).

use super::{push_attr, push_attr_if, xml_decl};
use crate::model::cache::PivotCache;
use crate::model::table::{
    AxisItem, CalculatedItem, DataField, Format, Location, PageField, PivotArea,
    PivotConditionalFormat, PivotField, PivotItem, PivotTable, PivotTableStyleInfo,
};

/// Emit pivotTable XML. The `cache` is needed to resolve cache-field
/// items into the pivot's `<items>` enumeration when the pivot field
/// has no explicit override (RFC-048 §10.4).
pub fn pivot_table_xml(pt: &PivotTable, cache: &PivotCache) -> Vec<u8> {
    let mut out = String::with_capacity(4096);
    xml_decl(&mut out);

    out.push_str("<pivotTableDefinition");
    push_attr(&mut out, "xmlns", crate::ns::SPREADSHEETML);
    push_attr(&mut out, "name", &pt.name);
    push_attr(&mut out, "cacheId", &pt.cache_id.to_string());

    push_attr(
        &mut out,
        "dataOnRows",
        if pt.data_on_rows { "1" } else { "0" },
    );
    push_attr(&mut out, "dataCaption", &pt.data_caption);
    if let Some(c) = &pt.grand_total_caption {
        push_attr(&mut out, "grandTotalCaption", c);
    }
    if let Some(c) = &pt.error_caption {
        push_attr(&mut out, "errorCaption", c);
    }
    if let Some(c) = &pt.missing_caption {
        push_attr(&mut out, "missingCaption", c);
    }

    push_attr(
        &mut out,
        "applyNumberFormats",
        if pt.apply_number_formats { "1" } else { "0" },
    );
    push_attr(
        &mut out,
        "applyBorderFormats",
        if pt.apply_border_formats { "1" } else { "0" },
    );
    push_attr(
        &mut out,
        "applyFontFormats",
        if pt.apply_font_formats { "1" } else { "0" },
    );
    push_attr(
        &mut out,
        "applyPatternFormats",
        if pt.apply_pattern_formats { "1" } else { "0" },
    );
    push_attr(
        &mut out,
        "applyAlignmentFormats",
        if pt.apply_alignment_formats { "1" } else { "0" },
    );
    push_attr(
        &mut out,
        "applyWidthHeightFormats",
        if pt.apply_width_height_formats {
            "1"
        } else {
            "0"
        },
    );

    push_attr(&mut out, "useAutoFormatting", "1");
    push_attr(&mut out, "itemPrintTitles", "1");
    push_attr(&mut out, "createdVersion", &pt.created_version.to_string());
    push_attr(&mut out, "updatedVersion", &pt.updated_version.to_string());
    push_attr(
        &mut out,
        "minRefreshableVersion",
        &pt.min_refreshable_version.to_string(),
    );
    push_attr(&mut out, "indent", "0");
    push_attr(&mut out, "outline", if pt.outline { "1" } else { "0" });
    push_attr(&mut out, "outlineData", if pt.outline { "1" } else { "0" });
    push_attr(&mut out, "compact", if pt.compact { "1" } else { "0" });
    push_attr(&mut out, "compactData", if pt.compact { "1" } else { "0" });
    push_attr(
        &mut out,
        "rowGrandTotals",
        if pt.row_grand_totals { "1" } else { "0" },
    );
    push_attr(
        &mut out,
        "colGrandTotals",
        if pt.col_grand_totals { "1" } else { "0" },
    );
    push_attr(&mut out, "multipleFieldFilters", "0");

    out.push('>');

    emit_location(&mut out, &pt.location);
    emit_pivot_fields(&mut out, &pt.pivot_fields, cache);

    // Two or more data fields add the `Values` pseudo-field (x="-2") as the
    // innermost field on the data axis; Excel reports the table as corrupt
    // without it.
    let multi_data = pt.data_fields.len() > 1;
    let values_on_rows = multi_data && pt.data_on_rows;
    let values_on_cols = multi_data && !pt.data_on_rows;
    if !pt.row_field_indices.is_empty() || values_on_rows {
        emit_axis_fields(&mut out, "rowFields", &pt.row_field_indices, values_on_rows);
    }
    if !pt.row_items.is_empty() {
        emit_axis_items(&mut out, "rowItems", &pt.row_items);
    }
    if !pt.col_field_indices.is_empty() || values_on_cols {
        emit_axis_fields(&mut out, "colFields", &pt.col_field_indices, values_on_cols);
    }
    if !pt.col_items.is_empty() {
        emit_axis_items(&mut out, "colItems", &pt.col_items);
    }
    if !pt.page_fields.is_empty() {
        emit_page_fields(&mut out, &pt.page_fields);
    }
    if !pt.data_fields.is_empty() {
        emit_data_fields(&mut out, &pt.data_fields);
    }
    if !pt.formats.is_empty() {
        emit_formats(&mut out, &pt.formats);
    }
    if !pt.conditional_formats.is_empty() {
        emit_conditional_formats(&mut out, &pt.conditional_formats);
    }
    if let Some(si) = &pt.style_info {
        emit_style_info(&mut out, si);
    }
    if !pt.calculated_items.is_empty() {
        emit_calculated_items(&mut out, &pt.calculated_items);
    }

    out.push_str("</pivotTableDefinition>");
    out.into_bytes()
}

fn emit_calculated_items(out: &mut String, items: &[CalculatedItem]) {
    // Pivot table XML carries calc items as a separate group; the
    // `<calculatedItems>` block lives at the same level as
    // `<pivotTableStyleInfo>` etc. Item-side schema mirrors the
    // calc-fields schema in the cache.
    out.push_str("<calculatedItems");
    push_attr(out, "count", &items.len().to_string());
    out.push('>');
    for ci in items {
        out.push_str("<calculatedItem");
        push_attr(out, "name", &ci.item_name);
        push_attr(out, "formula", &ci.formula);
        out.push_str("><pivotArea");
        push_attr(out, "type", "data");
        push_attr(out, "outline", "0");
        push_attr(out, "fieldPosition", "0");
        out.push_str("/></calculatedItem>");
    }
    out.push_str("</calculatedItems>");
}

fn emit_formats(out: &mut String, formats: &[Format]) {
    out.push_str("<formats");
    push_attr(out, "count", &formats.len().to_string());
    out.push('>');
    for f in formats {
        out.push_str("<format");
        push_attr(out, "dxfId", &f.dxf_id.to_string());
        push_attr(out, "action", &f.action);
        out.push('>');
        emit_pivot_area(out, &f.pivot_area);
        out.push_str("</format>");
    }
    out.push_str("</formats>");
}

fn emit_pivot_area(out: &mut String, a: &PivotArea) {
    out.push_str("<pivotArea");
    if let Some(f) = a.field {
        push_attr(out, "field", &f.to_string());
    }
    push_attr(out, "type", &a.area_type);
    push_attr_if(out, a.data_only, "dataOnly", "1");
    push_attr_if(out, a.label_only, "labelOnly", "1");
    push_attr_if(out, a.grand_row, "grandRow", "1");
    push_attr_if(out, a.grand_col, "grandCol", "1");
    if let Some(ci) = a.cache_index {
        push_attr(out, "cacheIndex", &ci.to_string());
    }
    if let Some(ax) = &a.axis {
        push_attr(out, "axis", ax);
    }
    if let Some(fp) = a.field_position {
        push_attr(out, "fieldPosition", &fp.to_string());
    }
    out.push_str("/>");
}

fn emit_conditional_formats(out: &mut String, cfs: &[PivotConditionalFormat]) {
    out.push_str("<conditionalFormats");
    push_attr(out, "count", &cfs.len().to_string());
    out.push('>');
    for cf in cfs {
        out.push_str("<conditionalFormat");
        push_attr(out, "scope", &cf.scope);
        push_attr(out, "type", &cf.cf_type);
        push_attr(out, "priority", &cf.priority.to_string());
        out.push('>');
        out.push_str("<pivotAreas");
        push_attr(out, "count", &cf.pivot_areas.len().to_string());
        out.push('>');
        for pa in &cf.pivot_areas {
            emit_pivot_area(out, pa);
        }
        out.push_str("</pivotAreas>");
        // Reference into workbook-scoped <dxfs> via dxfId (when set).
        if cf.dxf_id >= 0 {
            out.push_str("<extLst><ext uri=\"{B025F937-C7B1-47D3-B67F-A62EFF666E3E}\"><x14:dxfId xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\">");
            out.push_str(&cf.dxf_id.to_string());
            out.push_str("</x14:dxfId></ext></extLst>");
        }
        out.push_str("</conditionalFormat>");
    }
    out.push_str("</conditionalFormats>");
}

fn emit_location(out: &mut String, loc: &Location) {
    out.push_str("<location");
    push_attr(out, "ref", &loc.range);
    push_attr(out, "firstHeaderRow", &loc.first_header_row.to_string());
    push_attr(out, "firstDataRow", &loc.first_data_row.to_string());
    push_attr(out, "firstDataCol", &loc.first_data_col.to_string());
    if let Some(n) = loc.row_page_count {
        push_attr(out, "rowPageCount", &n.to_string());
    }
    if let Some(n) = loc.col_page_count {
        push_attr(out, "colPageCount", &n.to_string());
    }
    out.push_str("/>");
}

fn emit_pivot_fields(out: &mut String, fields: &[PivotField], cache: &PivotCache) {
    out.push_str("<pivotFields");
    push_attr(out, "count", &fields.len().to_string());
    out.push('>');
    for (i, f) in fields.iter().enumerate() {
        emit_pivot_field(out, f, cache, i);
    }
    out.push_str("</pivotFields>");
}

fn emit_pivot_field(out: &mut String, f: &PivotField, cache: &PivotCache, idx: usize) {
    out.push_str("<pivotField");
    if let Some(name) = &f.name {
        push_attr(out, "name", name);
    }
    if let Some(axis) = f.axis {
        push_attr(out, "axis", axis.xml_attr());
    }
    push_attr_if(out, f.data_field, "dataField", "1");
    push_attr(out, "showAll", if f.show_all { "1" } else { "0" });
    push_attr_if(out, !f.default_subtotal, "defaultSubtotal", "0");
    push_attr_if(out, f.sum_subtotal, "sumSubtotal", "1");
    push_attr_if(out, f.count_subtotal, "countSubtotal", "1");
    push_attr_if(out, f.avg_subtotal, "avgSubtotal", "1");
    push_attr_if(out, f.max_subtotal, "maxSubtotal", "1");
    push_attr_if(out, f.min_subtotal, "minSubtotal", "1");

    // Items: if explicit, emit those; if axis is set and field has
    // shared items in the cache, derive items from cache.
    let items: Option<Vec<PivotItem>> = if let Some(items) = &f.items {
        Some(items.clone())
    } else if f.axis.is_some() {
        derive_items_from_cache(cache, idx)
    } else {
        None
    };

    if let Some(items) = items {
        out.push('>');
        out.push_str("<items");
        push_attr(out, "count", &items.len().to_string());
        out.push('>');
        for it in &items {
            emit_pivot_item(out, it);
        }
        out.push_str("</items>");
        out.push_str("</pivotField>");
    } else {
        out.push_str("/>");
    }
}

/// Derive `<items>` for an axis-bearing pivot field from the cache's
/// sharedItems. RFC-048 §10.4: emits one `<item x="N"/>` per shared
/// item, then a trailing `<item t="default"/>` (the "(blank)" / total
/// catch-all).
fn derive_items_from_cache(cache: &PivotCache, idx: usize) -> Option<Vec<PivotItem>> {
    let cache_field = cache.fields.get(idx)?;
    let shared = cache_field.shared_items.items.as_ref()?;
    let mut out: Vec<PivotItem> = shared
        .iter()
        .enumerate()
        .map(|(i, _)| PivotItem {
            x: Some(i as u32),
            t: None,
            h: false,
            s: false,
            n: None,
        })
        .collect();
    out.push(PivotItem {
        x: None,
        t: Some(crate::model::table::PivotItemType::Default),
        h: false,
        s: false,
        n: None,
    });
    Some(out)
}

fn emit_pivot_item(out: &mut String, it: &PivotItem) {
    out.push_str("<item");
    if let Some(x) = it.x {
        push_attr(out, "x", &x.to_string());
    }
    if let Some(t) = it.t {
        push_attr(out, "t", t.xml_value());
    }
    push_attr_if(out, it.h, "h", "1");
    push_attr_if(out, it.s, "s", "1");
    if let Some(n) = &it.n {
        push_attr(out, "n", n);
    }
    out.push_str("/>");
}

/// `<rowFields>`/`<colFields>`. `values` appends the `Values` pseudo-field.
fn emit_axis_fields(out: &mut String, tag: &str, indices: &[u32], values: bool) {
    out.push('<');
    out.push_str(tag);
    let count = indices.len() + usize::from(values);
    push_attr(out, "count", &count.to_string());
    out.push('>');
    for &i in indices {
        out.push_str("<field");
        push_attr(out, "x", &i.to_string());
        out.push_str("/>");
    }
    if values {
        out.push_str("<field x=\"-2\"/>");
    }
    out.push_str("</");
    out.push_str(tag);
    out.push('>');
}

/// `<rowItems>`/`<colItems>`.
fn emit_axis_items(out: &mut String, tag: &str, items: &[AxisItem]) {
    out.push('<');
    out.push_str(tag);
    push_attr(out, "count", &items.len().to_string());
    out.push('>');
    for it in items {
        emit_axis_item(out, it);
    }
    out.push_str("</");
    out.push_str(tag);
    out.push('>');
}

fn emit_axis_item(out: &mut String, it: &AxisItem) {
    out.push_str("<i");
    if let Some(t) = it.t {
        push_attr(out, "t", t.xml_value());
    }
    if let Some(r) = it.r {
        push_attr(out, "r", &r.to_string());
    }
    if let Some(i) = it.i {
        push_attr(out, "i", &i.to_string());
    }
    if it.indices.is_empty() {
        out.push_str("/>");
        return;
    }
    out.push('>');
    for &x in &it.indices {
        out.push_str("<x");
        if x > 0 {
            push_attr(out, "v", &x.to_string());
        }
        out.push_str("/>");
    }
    out.push_str("</i>");
}

fn emit_page_fields(out: &mut String, fields: &[PageField]) {
    out.push_str("<pageFields");
    push_attr(out, "count", &fields.len().to_string());
    out.push('>');
    for f in fields {
        out.push_str("<pageField");
        push_attr(out, "fld", &f.field_index.to_string());
        if let Some(n) = &f.name {
            push_attr(out, "name", n);
        }
        push_attr(out, "item", &f.item_index.to_string());
        push_attr(out, "hier", &f.hier.to_string());
        if let Some(c) = &f.cap {
            push_attr(out, "cap", c);
        }
        out.push_str("/>");
    }
    out.push_str("</pageFields>");
}

fn emit_data_fields(out: &mut String, fields: &[DataField]) {
    out.push_str("<dataFields");
    push_attr(out, "count", &fields.len().to_string());
    out.push('>');
    for f in fields {
        out.push_str("<dataField");
        push_attr(out, "name", &f.name);
        push_attr(out, "fld", &f.field_index.to_string());
        push_attr(out, "subtotal", f.function.xml_value());
        if let Some(sda) = f.show_data_as {
            let s = match sda {
                crate::model::table::ShowDataAs::Normal => "normal",
                crate::model::table::ShowDataAs::Difference => "difference",
                crate::model::table::ShowDataAs::Percent => "percent",
                crate::model::table::ShowDataAs::PercentDiff => "percentDiff",
                crate::model::table::ShowDataAs::RunTotal => "runTotal",
                crate::model::table::ShowDataAs::PercentOfRow => "percentOfRow",
                crate::model::table::ShowDataAs::PercentOfCol => "percentOfCol",
                crate::model::table::ShowDataAs::PercentOfTotal => "percentOfTotal",
                crate::model::table::ShowDataAs::Index => "index",
            };
            push_attr(out, "showDataAs", s);
        }
        push_attr(out, "baseField", &f.base_field.to_string());
        push_attr(out, "baseItem", &f.base_item.to_string());
        if let Some(nfid) = f.num_fmt_id {
            push_attr(out, "numFmtId", &nfid.to_string());
        }
        out.push_str("/>");
    }
    out.push_str("</dataFields>");
}

fn emit_style_info(out: &mut String, si: &PivotTableStyleInfo) {
    out.push_str("<pivotTableStyleInfo");
    push_attr(out, "name", &si.name);
    push_attr(
        out,
        "showRowHeaders",
        if si.show_row_headers { "1" } else { "0" },
    );
    push_attr(
        out,
        "showColHeaders",
        if si.show_col_headers { "1" } else { "0" },
    );
    push_attr(
        out,
        "showRowStripes",
        if si.show_row_stripes { "1" } else { "0" },
    );
    push_attr(
        out,
        "showColStripes",
        if si.show_col_stripes { "1" } else { "0" },
    );
    push_attr(
        out,
        "showLastColumn",
        if si.show_last_column { "1" } else { "0" },
    );
    out.push_str("/>");
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::cache::{
        CacheField, CacheValue, DataType, PivotCache, SharedItems, WorksheetSource,
    };
    use crate::model::table::{AxisType, DataField, DataFunction, PivotField};

    fn dummy_cache() -> PivotCache {
        let region = CacheField {
            name: "region".into(),
            num_fmt_id: 0,
            data_type: DataType::String,
            shared_items: SharedItems {
                count: Some(2),
                items: Some(vec![
                    CacheValue::String("North".into()),
                    CacheValue::String("South".into()),
                ]),
                contains_semi_mixed_types: true,
                contains_string: true,
                contains_non_date: true,
                ..Default::default()
            },
            formula: None,
            hierarchy: None,
        };
        let revenue = CacheField {
            name: "revenue".into(),
            num_fmt_id: 0,
            data_type: DataType::Number,
            shared_items: SharedItems {
                count: None,
                items: None,
                contains_semi_mixed_types: false,
                contains_string: false,
                contains_number: true,
                min_value: Some(0.0),
                max_value: Some(1000.0),
                contains_non_date: true,
                ..Default::default()
            },
            formula: None,
            hierarchy: None,
        };
        PivotCache::new(
            0,
            WorksheetSource {
                sheet: "Sheet1".into(),
                range: "A1:B100".into(),
                name: None,
            },
            vec![region, revenue],
        )
    }

    fn dummy_table() -> PivotTable {
        PivotTable {
            name: "MyPivot".into(),
            cache_id: 0,
            location: Location {
                range: "F2:G5".into(),
                first_header_row: 0,
                first_data_row: 1,
                first_data_col: 1,
                row_page_count: None,
                col_page_count: None,
            },
            pivot_fields: vec![
                PivotField {
                    axis: Some(AxisType::Row),
                    ..Default::default()
                },
                PivotField {
                    data_field: true,
                    ..Default::default()
                },
            ],
            row_field_indices: vec![0],
            col_field_indices: vec![],
            page_fields: vec![],
            data_fields: vec![DataField {
                name: "Sum of revenue".into(),
                field_index: 1,
                function: DataFunction::Sum,
                show_data_as: None,
                base_field: 0,
                base_item: 0,
                num_fmt_id: None,
            }],
            row_items: vec![
                AxisItem {
                    indices: vec![0],
                    t: None,
                    r: None,
                    i: None,
                },
                AxisItem {
                    indices: vec![1],
                    t: None,
                    r: None,
                    i: None,
                },
                AxisItem {
                    indices: vec![0],
                    t: Some(crate::model::table::AxisItemType::Grand),
                    r: None,
                    i: None,
                },
            ],
            col_items: vec![],
            data_on_rows: false,
            outline: true,
            compact: true,
            row_grand_totals: true,
            col_grand_totals: true,
            data_caption: "Values".into(),
            grand_total_caption: None,
            error_caption: None,
            missing_caption: None,
            apply_number_formats: false,
            apply_border_formats: false,
            apply_font_formats: false,
            apply_pattern_formats: false,
            apply_alignment_formats: false,
            apply_width_height_formats: true,
            style_info: Some(PivotTableStyleInfo::default()),
            created_version: 6,
            updated_version: 6,
            min_refreshable_version: 3,
            calculated_items: vec![],
            formats: vec![],
            conditional_formats: vec![],
        }
    }

    #[test]
    fn emit_basic_pivot_table() {
        let cache = dummy_cache();
        let pt = dummy_table();
        let xml = pivot_table_xml(&pt, &cache);
        let s = std::str::from_utf8(&xml).unwrap();
        assert!(s.starts_with("<?xml"));
        assert!(s.contains("<pivotTableDefinition"));
        assert!(s.contains("name=\"MyPivot\""));
        assert!(s.contains("cacheId=\"0\""));
        assert!(s.contains("<location ref=\"F2:G5\""));
        assert!(s.contains("<pivotFields count=\"2\">"));
        assert!(s.contains("<pivotField axis=\"axisRow\""));
        assert!(s.contains("<pivotField dataField=\"1\""));
        assert!(s.contains("<rowFields count=\"1\">"));
        assert!(s.contains("<field x=\"0\"/>"));
        assert!(s.contains("<rowItems count=\"3\">"));
        assert!(s.contains("<dataFields count=\"1\">"));
        assert!(s.contains("name=\"Sum of revenue\""));
        assert!(s.contains("subtotal=\"sum\""));
        assert!(s.contains("fld=\"1\""));
        assert!(s.contains("<pivotTableStyleInfo"));
        assert!(s.contains("name=\"PivotStyleLight16\""));
    }

    #[test]
    fn emit_no_cols_omits_col_section() {
        let cache = dummy_cache();
        let pt = dummy_table();
        let xml = pivot_table_xml(&pt, &cache);
        let s = std::str::from_utf8(&xml).unwrap();
        assert!(!s.contains("<colFields"));
        assert!(!s.contains("<colItems"));
    }

    #[test]
    fn axis_grand_emits_grand_attr() {
        let cache = dummy_cache();
        let pt = dummy_table();
        let xml = pivot_table_xml(&pt, &cache);
        let s = std::str::from_utf8(&xml).unwrap();
        assert!(s.contains("<i t=\"grand\">"));
    }

    #[test]
    fn multiple_data_fields_add_values_col_field() {
        let cache = dummy_cache();
        let mut pt = dummy_table();
        let mut count = pt.data_fields[0].clone();
        count.name = "Count of revenue".into();
        count.function = DataFunction::Count;
        pt.data_fields.push(count);
        pt.col_items = vec![
            AxisItem {
                indices: vec![0],
                t: None,
                r: None,
                i: None,
            },
            AxisItem {
                indices: vec![1],
                t: None,
                r: None,
                i: Some(1),
            },
        ];
        let xml = pivot_table_xml(&pt, &cache);
        let s = std::str::from_utf8(&xml).unwrap();
        assert!(
            s.contains(r#"<colFields count="1"><field x="-2"/></colFields><colItems count="2"><i><x/></i><i i="1"><x v="1"/></i></colItems>"#),
            "{s}"
        );
        assert!(!s.contains(r#"<rowFields count="2""#), "{s}");
    }

    #[test]
    fn empty_axis_item_self_closes() {
        let cache = dummy_cache();
        let mut pt = dummy_table();
        pt.row_field_indices.clear();
        pt.row_items = vec![AxisItem {
            indices: vec![],
            t: None,
            r: None,
            i: None,
        }];
        let xml = pivot_table_xml(&pt, &cache);
        let s = std::str::from_utf8(&xml).unwrap();
        assert!(!s.contains("<rowFields"), "{s}");
        assert!(s.contains(r#"<rowItems count="1"><i/></rowItems>"#), "{s}");
    }

    #[test]
    fn determinism() {
        let cache = dummy_cache();
        let pt = dummy_table();
        assert_eq!(pivot_table_xml(&pt, &cache), pivot_table_xml(&pt, &cache));
    }
}
