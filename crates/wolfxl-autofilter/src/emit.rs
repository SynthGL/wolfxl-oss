//! Deterministic XML emit for `<autoFilter>` and `<sortState>`.
//!
//! Output is byte-stable: same input ⇒ same bytes, regardless of
//! HashMap iteration order. We emit in the source `Vec` order
//! (caller-controlled) and never use any unordered collection on the
//! hot path.
//!
//! Per RFC-056 §3.1, the canonical shape is:
//!
//! ```xml
//! <autoFilter ref="A1:D100">
//!   <filterColumn colId="0">
//!     <filters>
//!       <filter val="100"/>
//!     </filters>
//!   </filterColumn>
//!   <sortState ref="A2:A100">
//!     <sortCondition ref="A2:A100" descending="1"/>
//!   </sortState>
//! </autoFilter>
//! ```
//!
//! Note that `<sortState>` can appear standalone at slot 12 of a
//! `<worksheet>` OR nested inside `<autoFilter>`. Our writer ALWAYS
//! nests it inside `<autoFilter>` when an autofilter exists; if no
//! `<autoFilter>` is desired but a sort-state is, `emit_sort_state`
//! exists to write the standalone slot.

use crate::model::{
    AutoFilter, ColorFilter, CustomFilter, CustomFilters, DateGroupItem, DynamicFilter,
    FilterColumn, FilterKind, IconFilter, NumberFilter, SortCondition, SortState, StringFilter,
    Top10,
};

/// XML-escape into a `String` buffer, conservatively replacing the five
/// canonical XML predefined entities. Used for filter values, ref
/// strings, and any string attribute that may contain user input.
pub(crate) fn xml_escape(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for c in s.chars() {
        match c {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&apos;"),
            other => out.push(other),
        }
    }
    out
}

/// Format an `f64` as Excel does: integer-valued floats lose the
/// trailing `.0`, others use the minimum-digit representation. Pinned
/// here so the emitter is byte-stable across platforms (Rust's default
/// `{f64}` already prints `123` for integers, but we want explicit
/// guard rails against future stdlib changes).
pub(crate) fn format_number(n: f64) -> String {
    if n.is_finite() && n.fract() == 0.0 && n.abs() < 1e15 {
        format!("{}", n as i64)
    } else {
        format!("{n}")
    }
}

// ---------------------------------------------------------------------------
// Public emit functions
// ---------------------------------------------------------------------------

/// Emit the full `<autoFilter ...>...</autoFilter>` block. Returns
/// the bytes ready for `wolfxl_merger::SheetBlock::AutoFilter`. When
/// `auto_filter.ref_` is `None` AND there are no filter columns AND
/// no sort state, returns an empty vec (caller should skip the splice).
pub fn emit(auto_filter: &AutoFilter) -> Vec<u8> {
    if auto_filter.ref_.is_none()
        && auto_filter.filter_columns.is_empty()
        && auto_filter.sort_state.is_none()
    {
        return Vec::new();
    }

    let mut out = String::with_capacity(256);
    out.push_str("<autoFilter");
    if let Some(r) = &auto_filter.ref_ {
        out.push_str(" ref=\"");
        out.push_str(&xml_escape(r));
        out.push('"');
    }

    // Empty-element form: no children.
    if auto_filter.filter_columns.is_empty() && auto_filter.sort_state.is_none() {
        out.push_str("/>");
        return out.into_bytes();
    }
    out.push('>');

    for fc in &auto_filter.filter_columns {
        emit_filter_column(&mut out, fc);
    }
    if let Some(s) = &auto_filter.sort_state {
        emit_sort_state_inner(&mut out, s);
    }
    out.push_str("</autoFilter>");
    out.into_bytes()
}

/// Emit a standalone `<sortState>` block (slot 12). Used when the
/// caller wants sort but no autoFilter — rare but legal.
pub fn emit_sort_state(state: &SortState) -> Vec<u8> {
    let mut out = String::with_capacity(128);
    emit_sort_state_inner(&mut out, state);
    out.into_bytes()
}

fn emit_filter_column(out: &mut String, fc: &FilterColumn) {
    out.push_str("<filterColumn colId=\"");
    out.push_str(&fc.col_id.to_string());
    out.push('"');
    if fc.hidden_button {
        out.push_str(" hiddenButton=\"1\"");
    }
    if !fc.show_button {
        // Default true; emit only when false.
        out.push_str(" showButton=\"0\"");
    }

    let has_inner = fc.filter.is_some() || !fc.date_group_items.is_empty();
    if !has_inner {
        out.push_str("/>");
        return;
    }
    out.push('>');

    if let Some(f) = &fc.filter {
        match f {
            FilterKind::Blank(_) => emit_blank(out),
            FilterKind::Color(c) => emit_color(out, c),
            FilterKind::Custom(c) => emit_custom_filters(out, c),
            FilterKind::Dynamic(d) => emit_dynamic(out, d),
            FilterKind::Icon(i) => emit_icon(out, i),
            FilterKind::Number(n) => emit_number_filter(out, n, &fc.date_group_items),
            FilterKind::String(s) => emit_string_filter(out, s),
            FilterKind::Top10(t) => emit_top10(out, t),
        }
    } else if !fc.date_group_items.is_empty() {
        // dateGroupItems can appear standalone inside <filters> with
        // no `<filter>` siblings.
        emit_date_group_items_only(out, &fc.date_group_items);
    }

    out.push_str("</filterColumn>");
}

fn emit_blank(out: &mut String) {
    // openpyxl emits `<filters blank="1"/>`. Per RFC §2.1 BlankFilter
    // is the dedicated empty-cell marker; openpyxl's serialised form
    // is `<filters blank="1"/>`.
    out.push_str("<filters blank=\"1\"/>");
}

fn emit_color(out: &mut String, c: &ColorFilter) {
    out.push_str("<colorFilter dxfId=\"");
    out.push_str(&c.dxf_id.to_string());
    out.push('"');
    if !c.cell_color {
        out.push_str(" cellColor=\"0\"");
    }
    out.push_str("/>");
}

fn emit_custom_filters(out: &mut String, c: &CustomFilters) {
    out.push_str("<customFilters");
    if c.and_ {
        out.push_str(" and=\"1\"");
    }
    if c.filters.is_empty() {
        out.push_str("/>");
        return;
    }
    out.push('>');
    for cf in &c.filters {
        emit_custom_filter(out, cf);
    }
    out.push_str("</customFilters>");
}

fn emit_custom_filter(out: &mut String, cf: &CustomFilter) {
    out.push_str("<customFilter");
    // Default operator is "equal"; we always emit explicitly for
    // byte-stability + roundtrip clarity.
    out.push_str(" operator=\"");
    out.push_str(cf.operator.as_xml());
    out.push_str("\" val=\"");
    out.push_str(&xml_escape(&cf.val));
    out.push_str("\"/>");
}

fn emit_dynamic(out: &mut String, d: &DynamicFilter) {
    out.push_str("<dynamicFilter type=\"");
    out.push_str(&xml_escape(&d.type_.as_xml()));
    out.push('"');
    if let Some(v) = d.val {
        out.push_str(" val=\"");
        out.push_str(&format_number(v));
        out.push('"');
    }
    if let Some(s) = &d.val_iso {
        out.push_str(" valIso=\"");
        out.push_str(&xml_escape(s));
        out.push('"');
    }
    if let Some(s) = &d.max_val_iso {
        out.push_str(" maxValIso=\"");
        out.push_str(&xml_escape(s));
        out.push('"');
    }
    out.push_str("/>");
}

fn emit_icon(out: &mut String, i: &IconFilter) {
    out.push_str("<iconFilter iconSet=\"");
    out.push_str(&xml_escape(&i.icon_set));
    out.push_str("\" iconId=\"");
    out.push_str(&i.icon_id.to_string());
    out.push_str("\"/>");
}

fn emit_number_filter(out: &mut String, n: &NumberFilter, date_items: &[DateGroupItem]) {
    out.push_str("<filters");
    if n.blank {
        out.push_str(" blank=\"1\"");
    }
    if let Some(ct) = &n.calendar_type {
        out.push_str(" calendarType=\"");
        out.push_str(&xml_escape(ct));
        out.push('"');
    }
    if n.filters.is_empty() && date_items.is_empty() {
        out.push_str("/>");
        return;
    }
    out.push('>');
    for v in &n.filters {
        out.push_str("<filter val=\"");
        out.push_str(&format_number(*v));
        out.push_str("\"/>");
    }
    for dgi in date_items {
        emit_date_group_item(out, dgi);
    }
    out.push_str("</filters>");
}

fn emit_string_filter(out: &mut String, s: &StringFilter) {
    out.push_str("<filters");
    if s.values.is_empty() {
        out.push_str("/>");
        return;
    }
    out.push('>');
    for v in &s.values {
        out.push_str("<filter val=\"");
        out.push_str(&xml_escape(v));
        out.push_str("\"/>");
    }
    out.push_str("</filters>");
}

fn emit_top10(out: &mut String, t: &Top10) {
    out.push_str("<top10");
    // top default true; emit only when false
    if !t.top {
        out.push_str(" top=\"0\"");
    }
    if t.percent {
        out.push_str(" percent=\"1\"");
    }
    out.push_str(" val=\"");
    out.push_str(&format_number(t.val));
    out.push('"');
    if let Some(fv) = t.filter_val {
        out.push_str(" filterVal=\"");
        out.push_str(&format_number(fv));
        out.push('"');
    }
    out.push_str("/>");
}

fn emit_date_group_items_only(out: &mut String, items: &[DateGroupItem]) {
    out.push_str("<filters>");
    for dgi in items {
        emit_date_group_item(out, dgi);
    }
    out.push_str("</filters>");
}

fn emit_date_group_item(out: &mut String, d: &DateGroupItem) {
    out.push_str("<dateGroupItem year=\"");
    out.push_str(&d.year.to_string());
    out.push('"');
    if let Some(m) = d.month {
        out.push_str(" month=\"");
        out.push_str(&m.to_string());
        out.push('"');
    }
    if let Some(day) = d.day {
        out.push_str(" day=\"");
        out.push_str(&day.to_string());
        out.push('"');
    }
    if let Some(h) = d.hour {
        out.push_str(" hour=\"");
        out.push_str(&h.to_string());
        out.push('"');
    }
    if let Some(min) = d.minute {
        out.push_str(" minute=\"");
        out.push_str(&min.to_string());
        out.push('"');
    }
    if let Some(s) = d.second {
        out.push_str(" second=\"");
        out.push_str(&s.to_string());
        out.push('"');
    }
    out.push_str(" dateTimeGrouping=\"");
    out.push_str(d.date_time_grouping.as_xml());
    out.push_str("\"/>");
}

fn emit_sort_state_inner(out: &mut String, state: &SortState) {
    out.push_str("<sortState");
    if state.column_sort {
        out.push_str(" columnSort=\"1\"");
    }
    if state.case_sensitive {
        out.push_str(" caseSensitive=\"1\"");
    }
    if let Some(r) = &state.ref_ {
        out.push_str(" ref=\"");
        out.push_str(&xml_escape(r));
        out.push('"');
    }
    if state.sort_conditions.is_empty() {
        out.push_str("/>");
        return;
    }
    out.push('>');
    for sc in &state.sort_conditions {
        emit_sort_condition(out, sc);
    }
    out.push_str("</sortState>");
}

fn emit_sort_condition(out: &mut String, sc: &SortCondition) {
    out.push_str("<sortCondition ref=\"");
    out.push_str(&xml_escape(&sc.ref_));
    out.push('"');
    if sc.descending {
        out.push_str(" descending=\"1\"");
    }
    // sort_by default "value"; emit only when not the default.
    if !matches!(sc.sort_by, crate::model::SortBy::Value) {
        out.push_str(" sortBy=\"");
        out.push_str(sc.sort_by.as_xml());
        out.push('"');
    }
    if let Some(cl) = &sc.custom_list {
        out.push_str(" customList=\"");
        out.push_str(&xml_escape(cl));
        out.push('"');
    }
    if let Some(d) = sc.dxf_id {
        out.push_str(" dxfId=\"");
        out.push_str(&d.to_string());
        out.push('"');
    }
    if let Some(s) = &sc.icon_set {
        out.push_str(" iconSet=\"");
        out.push_str(&xml_escape(s));
        out.push('"');
    }
    if let Some(i) = sc.icon_id {
        out.push_str(" iconId=\"");
        out.push_str(&i.to_string());
        out.push('"');
    }
    out.push_str("/>");
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::{
        BlankFilter, ColorFilter, CustomFilter, CustomFilterOp, CustomFilters, DynamicFilterType,
        SortBy,
    };

    fn af_with_ref() -> AutoFilter {
        AutoFilter {
            ref_: Some("A1:D100".into()),
            filter_columns: Vec::new(),
            sort_state: None,
        }
    }

    #[test]
    fn empty_returns_empty_vec() {
        let af = AutoFilter::default();
        assert!(emit(&af).is_empty());
    }

    #[test]
    fn ref_only_emits_self_closing() {
        let af = af_with_ref();
        let bytes = emit(&af);
        assert_eq!(
            std::str::from_utf8(&bytes).unwrap(),
            r#"<autoFilter ref="A1:D100"/>"#
        );
    }

    #[test]
    fn number_filter_emits() {
        let mut af = af_with_ref();
        af.filter_columns.push(FilterColumn {
            col_id: 0,
            hidden_button: false,
            show_button: true,
            filter: Some(FilterKind::Number(NumberFilter {
                filters: vec![100.0, 200.0, 300.0],
                blank: false,
                calendar_type: None,
            })),
            date_group_items: Vec::new(),
        });
        let s = String::from_utf8(emit(&af)).unwrap();
        assert_eq!(
            s,
            r#"<autoFilter ref="A1:D100"><filterColumn colId="0"><filters><filter val="100"/><filter val="200"/><filter val="300"/></filters></filterColumn></autoFilter>"#
        );
    }

    #[test]
    fn string_filter_escapes() {
        let mut af = af_with_ref();
        af.filter_columns.push(FilterColumn {
            col_id: 1,
            hidden_button: false,
            show_button: true,
            filter: Some(FilterKind::String(StringFilter {
                values: vec!["red".into(), "<a>".into(), "&b".into()],
            })),
            date_group_items: Vec::new(),
        });
        let s = String::from_utf8(emit(&af)).unwrap();
        assert!(s.contains(r#"<filter val="red"/>"#));
        assert!(s.contains(r#"<filter val="&lt;a&gt;"/>"#));
        assert!(s.contains(r#"<filter val="&amp;b"/>"#));
    }

    #[test]
    fn blank_filter_emits() {
        let mut af = af_with_ref();
        af.filter_columns.push(FilterColumn {
            col_id: 0,
            hidden_button: false,
            show_button: true,
            filter: Some(FilterKind::Blank(BlankFilter)),
            date_group_items: Vec::new(),
        });
        let s = String::from_utf8(emit(&af)).unwrap();
        assert!(s.contains(r#"<filters blank="1"/>"#));
    }

    #[test]
    fn color_filter_default_cell_color() {
        let mut af = af_with_ref();
        af.filter_columns.push(FilterColumn {
            col_id: 0,
            hidden_button: false,
            show_button: true,
            filter: Some(FilterKind::Color(ColorFilter {
                dxf_id: 7,
                cell_color: true,
            })),
            date_group_items: Vec::new(),
        });
        let s = String::from_utf8(emit(&af)).unwrap();
        assert!(s.contains(r#"<colorFilter dxfId="7"/>"#));
    }

    #[test]
    fn color_filter_font_color_explicit() {
        let mut af = af_with_ref();
        af.filter_columns.push(FilterColumn {
            col_id: 0,
            hidden_button: false,
            show_button: true,
            filter: Some(FilterKind::Color(ColorFilter {
                dxf_id: 3,
                cell_color: false,
            })),
            date_group_items: Vec::new(),
        });
        let s = String::from_utf8(emit(&af)).unwrap();
        assert!(s.contains(r#"<colorFilter dxfId="3" cellColor="0"/>"#));
    }

    #[test]
    fn custom_filters_or() {
        let mut af = af_with_ref();
        af.filter_columns.push(FilterColumn {
            col_id: 2,
            hidden_button: false,
            show_button: true,
            filter: Some(FilterKind::Custom(CustomFilters {
                filters: vec![
                    CustomFilter {
                        operator: CustomFilterOp::GreaterThan,
                        val: "5".into(),
                    },
                    CustomFilter {
                        operator: CustomFilterOp::LessThan,
                        val: "100".into(),
                    },
                ],
                and_: false,
            })),
            date_group_items: Vec::new(),
        });
        let s = String::from_utf8(emit(&af)).unwrap();
        assert!(s.contains(r#"<customFilters>"#));
        assert!(s.contains(r#"<customFilter operator="greaterThan" val="5"/>"#));
        assert!(s.contains(r#"<customFilter operator="lessThan" val="100"/>"#));
    }

    #[test]
    fn custom_filters_and() {
        let mut af = af_with_ref();
        af.filter_columns.push(FilterColumn {
            col_id: 0,
            hidden_button: false,
            show_button: true,
            filter: Some(FilterKind::Custom(CustomFilters {
                filters: vec![CustomFilter {
                    operator: CustomFilterOp::Equal,
                    val: "x".into(),
                }],
                and_: true,
            })),
            date_group_items: Vec::new(),
        });
        let s = String::from_utf8(emit(&af)).unwrap();
        assert!(s.contains(r#"<customFilters and="1">"#));
    }

    #[test]
    fn dynamic_filter_today() {
        let mut af = af_with_ref();
        af.filter_columns.push(FilterColumn {
            col_id: 0,
            hidden_button: false,
            show_button: true,
            filter: Some(FilterKind::Dynamic(DynamicFilter {
                type_: DynamicFilterType::Today,
                val: None,
                val_iso: None,
                max_val_iso: None,
            })),
            date_group_items: Vec::new(),
        });
        let s = String::from_utf8(emit(&af)).unwrap();
        assert!(s.contains(r#"<dynamicFilter type="today"/>"#));
    }

    #[test]
    fn dynamic_filter_above_average_with_val() {
        let mut af = af_with_ref();
        af.filter_columns.push(FilterColumn {
            col_id: 0,
            hidden_button: false,
            show_button: true,
            filter: Some(FilterKind::Dynamic(DynamicFilter {
                type_: DynamicFilterType::AboveAverage,
                val: Some(42.5),
                val_iso: None,
                max_val_iso: None,
            })),
            date_group_items: Vec::new(),
        });
        let s = String::from_utf8(emit(&af)).unwrap();
        assert!(s.contains(r#"type="aboveAverage""#));
        assert!(s.contains(r#"val="42.5""#));
    }

    #[test]
    fn dynamic_filter_month() {
        assert_eq!(DynamicFilterType::Month(7).as_xml(), "M7");
        assert_eq!(
            DynamicFilterType::parse("M12"),
            Some(DynamicFilterType::Month(12))
        );
        assert_eq!(DynamicFilterType::parse("M0"), None);
        assert_eq!(DynamicFilterType::parse("M13"), None);
    }

    #[test]
    fn icon_filter_emits() {
        let mut af = af_with_ref();
        af.filter_columns.push(FilterColumn {
            col_id: 0,
            hidden_button: false,
            show_button: true,
            filter: Some(FilterKind::Icon(IconFilter {
                icon_set: "5Quarters".into(),
                icon_id: 2,
            })),
            date_group_items: Vec::new(),
        });
        let s = String::from_utf8(emit(&af)).unwrap();
        assert!(s.contains(r#"<iconFilter iconSet="5Quarters" iconId="2"/>"#));
    }

    #[test]
    fn top10_default_top_n() {
        let mut af = af_with_ref();
        af.filter_columns.push(FilterColumn {
            col_id: 0,
            hidden_button: false,
            show_button: true,
            filter: Some(FilterKind::Top10(Top10 {
                top: true,
                percent: false,
                val: 5.0,
                filter_val: None,
            })),
            date_group_items: Vec::new(),
        });
        let s = String::from_utf8(emit(&af)).unwrap();
        assert!(s.contains(r#"<top10 val="5"/>"#));
    }

    #[test]
    fn top10_bottom_percent() {
        let mut af = af_with_ref();
        af.filter_columns.push(FilterColumn {
            col_id: 0,
            hidden_button: false,
            show_button: true,
            filter: Some(FilterKind::Top10(Top10 {
                top: false,
                percent: true,
                val: 25.0,
                filter_val: Some(99.5),
            })),
            date_group_items: Vec::new(),
        });
        let s = String::from_utf8(emit(&af)).unwrap();
        assert!(s.contains(r#"<top10 top="0" percent="1" val="25" filterVal="99.5"/>"#));
    }

    #[test]
    fn sort_state_emits_inside_autofilter() {
        let mut af = af_with_ref();
        af.sort_state = Some(SortState {
            sort_conditions: vec![SortCondition {
                ref_: "A2:A100".into(),
                descending: true,
                sort_by: SortBy::Value,
                custom_list: None,
                dxf_id: None,
                icon_set: None,
                icon_id: None,
            }],
            column_sort: false,
            case_sensitive: false,
            ref_: Some("A2:A100".into()),
        });
        let s = String::from_utf8(emit(&af)).unwrap();
        assert!(s.contains(r#"<sortState ref="A2:A100">"#));
        assert!(s.contains(r#"<sortCondition ref="A2:A100" descending="1"/>"#));
    }

    #[test]
    fn non_value_sort_conditions_use_sort_by() {
        for (sort_by, key) in [
            (SortBy::CellColor, "cellColor"),
            (SortBy::FontColor, "fontColor"),
            (SortBy::Icon, "icon"),
        ] {
            let condition = SortCondition {
                ref_: "A2:A100".into(),
                descending: false,
                sort_by,
                custom_list: None,
                dxf_id: None,
                icon_set: None,
                icon_id: None,
            };
            let mut xml = String::new();
            emit_sort_condition(&mut xml, &condition);
            assert_eq!(
                xml,
                format!(r#"<sortCondition ref="A2:A100" sortBy="{key}"/>"#)
            );
        }
    }

    #[test]
    fn deterministic_two_runs_same_bytes() {
        let mut af = af_with_ref();
        af.filter_columns.push(FilterColumn {
            col_id: 0,
            hidden_button: false,
            show_button: true,
            filter: Some(FilterKind::Number(NumberFilter {
                filters: vec![1.0, 2.0, 3.0],
                blank: true,
                calendar_type: None,
            })),
            date_group_items: Vec::new(),
        });
        assert_eq!(emit(&af), emit(&af));
    }

    #[test]
    fn date_group_item_year_only() {
        let mut af = af_with_ref();
        af.filter_columns.push(FilterColumn {
            col_id: 0,
            hidden_button: false,
            show_button: true,
            filter: None,
            date_group_items: vec![DateGroupItem {
                year: 2024,
                month: None,
                day: None,
                hour: None,
                minute: None,
                second: None,
                date_time_grouping: crate::model::DateTimeGrouping::Year,
            }],
        });
        let s = String::from_utf8(emit(&af)).unwrap();
        assert!(s.contains(r#"<dateGroupItem year="2024" dateTimeGrouping="year"/>"#));
    }

    #[test]
    fn show_button_false_emits() {
        let mut af = af_with_ref();
        af.filter_columns.push(FilterColumn {
            col_id: 1,
            hidden_button: true,
            show_button: false,
            filter: None,
            date_group_items: Vec::new(),
        });
        let s = String::from_utf8(emit(&af)).unwrap();
        assert!(s.contains(r#"hiddenButton="1""#));
        assert!(s.contains(r#"showButton="0""#));
    }

    #[test]
    fn format_number_integer() {
        assert_eq!(format_number(100.0), "100");
        assert_eq!(format_number(-3.0), "-3");
        assert_eq!(format_number(1.5), "1.5");
        assert_eq!(format_number(0.0), "0");
    }

    #[test]
    fn xml_escape_all_five() {
        assert_eq!(xml_escape("<>&\"'"), "&lt;&gt;&amp;&quot;&apos;");
    }
}
