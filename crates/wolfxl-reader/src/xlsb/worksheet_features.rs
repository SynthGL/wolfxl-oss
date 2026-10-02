use crate::{AutoFilterInfo, ConditionalFormatRule, DataValidation, FilterColumnInfo, Table};

use super::{formula, le_i32, le_u16, le_u32, wide_string, worksheet_meta, Records, Result};

pub(super) fn parse_auto_filter_begin(payload: &[u8]) -> Option<AutoFilterInfo> {
    worksheet_meta::parse_rfx(payload).map(|ref_range| AutoFilterInfo {
        ref_range,
        filter_columns: Vec::new(),
        sort_state: None,
    })
}

pub(super) fn parse_filter_column(payload: &[u8]) -> Option<FilterColumnInfo> {
    if payload.len() < 6 {
        return None;
    }
    let flags = le_u16(&payload[4..6]);
    Some(FilterColumnInfo {
        col_id: le_u32(&payload[0..4]),
        hidden_button: flags & 0x0001 != 0,
        show_button: flags & 0x0002 == 0,
        filter: None,
        date_group_items: Vec::new(),
    })
}

pub(super) fn parse_filter_value(payload: &[u8]) -> Option<String> {
    let mut consumed = 0;
    wide_string(payload, &mut consumed).ok()
}

pub(super) fn parse_data_validation(
    payload: &[u8],
    list_formula: Option<String>,
) -> Option<DataValidation> {
    if payload.len() < 8 {
        return None;
    }
    let flags = le_u32(&payload[0..4]);
    let (range, consumed) = parse_sqref_with_consumed(payload.get(4..)?)?;
    let mut offset = 4 + consumed;
    let error_title = parse_optional_wide_string(payload, &mut offset);
    let error = parse_optional_wide_string(payload, &mut offset);
    let prompt_title = parse_optional_wide_string(payload, &mut offset);
    let prompt = parse_optional_wide_string(payload, &mut offset);
    let validation_type = data_validation_type(flags & 0x0f);
    let operator = if matches!(validation_type, "any" | "list" | "custom") {
        None
    } else {
        data_validation_operator((flags >> 20) & 0x0f).map(str::to_string)
    };
    Some(DataValidation {
        range,
        validation_type: validation_type.to_string(),
        operator,
        formula1: list_formula,
        formula2: None,
        allow_blank: flags & (1 << 8) != 0,
        show_input_message: flags & (1 << 18) != 0,
        show_error_message: flags & (1 << 19) != 0,
        error_title: error_title.filter(|value| !value.is_empty()),
        error: error.filter(|value| !value.is_empty()),
        prompt_title: prompt_title.filter(|value| !value.is_empty()),
        prompt: prompt.filter(|value| !value.is_empty()),
    })
}

pub(super) fn parse_conditional_formatting_begin(payload: &[u8]) -> Option<String> {
    if payload.len() < 8 {
        return None;
    }
    parse_sqref_with_consumed(payload.get(8..)?).map(|(range, _)| range)
}

pub(super) fn parse_conditional_format_rule(
    payload: &[u8],
    current_range: Option<&str>,
) -> Option<ConditionalFormatRule> {
    if payload.len() < 42 {
        return None;
    }
    let rule_type = conditional_format_rule_type(le_u32(&payload[0..4]), le_u32(&payload[4..8]));
    let operator = conditional_format_operator(le_i32(&payload[16..20]));
    let priority = Some(le_i32(&payload[12..16]) as i64).filter(|priority| *priority > 0);
    let flags = le_u16(&payload[28..30]);
    let formula1_len = le_u32(&payload[30..34]) as usize;
    let formula2_len = le_u32(&payload[34..38]) as usize;
    let formula3_len = le_u32(&payload[38..42]) as usize;
    let mut offset = 42;
    let str_param = read_nullable_string_value(payload, &mut offset);
    let formula1 = read_conditional_formula(payload, &mut offset, formula1_len);
    let formula2 = read_conditional_formula(payload, &mut offset, formula2_len);
    let _formula3 = read_conditional_formula(payload, &mut offset, formula3_len);
    let formula = formula1
        .or(formula2)
        .or_else(|| str_param.filter(|value| !value.is_empty()))
        .map(|value| ensure_formula_prefix_bin(&value));
    Some(ConditionalFormatRule {
        range: current_range.unwrap_or_default().to_string(),
        rule_type: rule_type.to_string(),
        operator: operator.map(str::to_string),
        formula,
        dxf_id: None,
        priority,
        stop_if_true: Some(flags & 0x0002 != 0),
        color_scale: None,
    })
}

fn read_conditional_formula(payload: &[u8], offset: &mut usize, len: usize) -> Option<String> {
    if len == 0 {
        return None;
    }
    let rgce = payload.get(*offset..offset.checked_add(len)?)?;
    *offset += len;
    formula::parse_formula_rgce_with_context(rgce, None)
}

fn conditional_format_rule_type(raw_type: u32, raw_template: u32) -> &'static str {
    match (raw_type, raw_template) {
        (1, _) => "cellIs",
        (2, _) => "expression",
        (3, _) => "colorScale",
        (4, _) => "dataBar",
        (5, _) => "top10",
        (6, _) => "iconSet",
        _ => "expression",
    }
}

fn conditional_format_operator(raw: i32) -> Option<&'static str> {
    match raw {
        1 => Some("between"),
        2 => Some("notBetween"),
        3 => Some("equal"),
        4 => Some("notEqual"),
        5 => Some("greaterThan"),
        6 => Some("lessThan"),
        7 => Some("greaterThanOrEqual"),
        8 => Some("lessThanOrEqual"),
        _ => None,
    }
}

fn ensure_formula_prefix_bin(formula: &str) -> String {
    if formula.starts_with('=') {
        formula.to_string()
    } else {
        format!("={formula}")
    }
}

#[derive(Debug, Default)]
struct TableBuilderBin {
    name: String,
    ref_range: String,
    header_row: bool,
    totals_row: bool,
    comment: Option<String>,
    table_type: Option<String>,
    totals_row_shown: Option<bool>,
    style: Option<String>,
    show_first_column: bool,
    show_last_column: bool,
    show_row_stripes: bool,
    show_column_stripes: bool,
    columns: Vec<String>,
    auto_filter: Option<AutoFilterInfo>,
}

impl TableBuilderBin {
    fn finish(self) -> Option<Table> {
        if self.name.is_empty() || self.ref_range.is_empty() {
            return None;
        }
        Some(Table {
            name: self.name,
            ref_range: self.ref_range,
            header_row: self.header_row,
            totals_row: self.totals_row,
            comment: self.comment,
            table_type: self.table_type,
            totals_row_shown: self.totals_row_shown,
            style: self.style,
            show_first_column: self.show_first_column,
            show_last_column: self.show_last_column,
            show_row_stripes: self.show_row_stripes,
            show_column_stripes: self.show_column_stripes,
            columns: self.columns,
            auto_filter: self.auto_filter,
        })
    }
}

pub(super) fn parse_table_bin(data: &[u8]) -> Option<Table> {
    let mut table = TableBuilderBin::default();
    for record in Records::new(data).filter_map(Result::ok) {
        match record.typ {
            0x0157 => {
                table = parse_table_begin(record.payload)?;
            }
            0x015b => {
                if let Some(column) = parse_table_column(record.payload) {
                    table.columns.push(column);
                }
            }
            0x00a1 => {
                table.auto_filter = parse_auto_filter_begin(record.payload);
            }
            0x0201 => {
                apply_table_style(record.payload, &mut table);
            }
            _ => {}
        }
    }
    table.finish()
}

fn parse_table_begin(payload: &[u8]) -> Option<TableBuilderBin> {
    if payload.len() < 64 {
        return None;
    }
    let ref_range = worksheet_meta::parse_rfx(&payload[0..16])?;
    let list_type = le_u32(&payload[16..20]);
    let header_row = le_u32(&payload[24..28]) != 0;
    let totals_row = le_u32(&payload[28..32]) != 0;
    let flags = le_u32(&payload[32..36]);
    let mut offset = 64;
    let name = read_nullable_string_value(payload, &mut offset);
    let display_name = read_nullable_string_value(payload, &mut offset);
    let comment =
        read_nullable_string_value(payload, &mut offset).filter(|value| !value.is_empty());
    let _style_header = read_nullable_string_value(payload, &mut offset);
    let _style_data = read_nullable_string_value(payload, &mut offset);
    let _style_agg = read_nullable_string_value(payload, &mut offset);
    let table_name = name
        .filter(|value| !value.is_empty())
        .or_else(|| display_name.filter(|value| !value.is_empty()))?;
    Some(TableBuilderBin {
        name: table_name,
        ref_range,
        header_row,
        totals_row,
        comment,
        table_type: (list_type != 0).then_some("worksheet".to_string()),
        totals_row_shown: Some(flags & 0x0001 != 0),
        ..TableBuilderBin::default()
    })
}

fn parse_table_column(payload: &[u8]) -> Option<String> {
    if payload.len() < 24 {
        return None;
    }
    let mut offset = 24;
    let name = read_nullable_string_value(payload, &mut offset);
    let caption = read_nullable_string_value(payload, &mut offset);
    caption
        .filter(|value| !value.is_empty())
        .or_else(|| name.filter(|value| !value.is_empty()))
}

fn apply_table_style(payload: &[u8], table: &mut TableBuilderBin) {
    if payload.len() < 2 {
        return;
    }
    let flags = le_u16(&payload[0..2]);
    table.show_first_column = flags & 0x0001 != 0;
    table.show_last_column = flags & 0x0002 != 0;
    table.show_row_stripes = flags & 0x0004 != 0;
    table.show_column_stripes = flags & 0x0008 != 0;
    let mut offset = 2;
    table.style =
        read_nullable_string_value(payload, &mut offset).filter(|value| !value.is_empty());
}

fn read_nullable_string_value(payload: &[u8], offset: &mut usize) -> Option<String> {
    worksheet_meta::read_nullable_wide_string_at(payload, offset).flatten()
}

fn parse_sqref_with_consumed(payload: &[u8]) -> Option<(String, usize)> {
    let count = payload.get(0..4).map(le_u32)? as usize;
    if count == 0 {
        return None;
    }
    let range_bytes = payload.get(4..4 + count.checked_mul(16)?)?;
    let ranges: Vec<String> = range_bytes
        .chunks_exact(16)
        .filter_map(worksheet_meta::parse_rfx)
        .collect();
    (!ranges.is_empty()).then(|| (ranges.join(" "), 4 + count * 16))
}

fn parse_optional_wide_string(payload: &[u8], offset: &mut usize) -> Option<String> {
    let start = *offset;
    let mut consumed = 0;
    let value = wide_string(payload.get(start..)?, &mut consumed).ok()?;
    *offset = start + consumed;
    Some(value)
}

fn data_validation_type(raw: u32) -> &'static str {
    match raw {
        1 => "whole",
        2 => "decimal",
        3 => "list",
        4 => "date",
        5 => "time",
        6 => "textLength",
        7 => "custom",
        _ => "any",
    }
}

fn data_validation_operator(raw: u32) -> Option<&'static str> {
    match raw {
        0 => Some("between"),
        1 => Some("notBetween"),
        2 => Some("equal"),
        3 => Some("notEqual"),
        4 => Some("greaterThan"),
        5 => Some("lessThan"),
        6 => Some("greaterThanOrEqual"),
        7 => Some("lessThanOrEqual"),
        _ => None,
    }
}
