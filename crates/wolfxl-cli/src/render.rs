//! Output renderers for `wolfxl peek`.
//!
//! Four shapes, all driven from the same `Sheet` snapshot:
//! - `boxed`: full TTY preview with banner, sheet metadata, box-drawn table
//! - `text`: tab-separated rows, no banner (drop-in for `awk`/`cut`)
//! - `csv`:  RFC 4180 CSV, integer thousand-grouping retained
//! - `json`: machine-shaped `{sheet, rows, columns, headers, data}`
//! - `markdown`: GitHub-flavored table for LLM/context converters
//!
//! Format choices (integer grouping, two-decimal floats, ISO dates, JSON shape)
//! are locked in by golden tests in `tests/cli.rs` so token-cost benchmarks in
//! the `spreadsheet-peek` skill stay reproducible.

use std::io::Write;

use serde_json::Value;
use unicode_width::UnicodeWidthStr;
use wolfxl_core::{classify_format, format_cell, Cell, CellValue, FormatCategory, Sheet};

pub struct RenderOptions<'a> {
    pub max_rows: Option<usize>,
    pub max_columns: Option<usize>,
    pub max_width: usize,
    pub all_sheets: &'a [String],
}

pub fn text<W: Write>(w: &mut W, sheet: &Sheet, _opts: &RenderOptions) -> std::io::Result<()> {
    // `-n` caps box-mode only; text/csv/json emit the full sheet so the
    // benchmark can pipe through `head` reproducibly. Header row is emitted
    // as raw strings (no int-grouping) — header text should round-trip even
    // when a header literal happens to be numeric.
    writeln!(w, "{}", sheet.headers().join("\t"))?;
    for row in sheet.rows().iter().skip(1) {
        let cells: Vec<String> = row
            .iter()
            .map(|c| display_sheet_cell(sheet, c, true))
            .collect();
        writeln!(w, "{}", cells.join("\t"))?;
    }
    Ok(())
}

pub fn csv<W: Write>(w: &mut W, sheet: &Sheet, _opts: &RenderOptions) -> std::io::Result<()> {
    // Quote every header cell individually so commas, quotes, and embedded
    // newlines in a header literal don't shred the output row.
    let headers: Vec<String> = sheet.headers().iter().map(|h| csv_quote(h)).collect();
    writeln!(w, "{}", headers.join(","))?;
    for row in sheet.rows().iter().skip(1) {
        let cells: Vec<String> = row
            .iter()
            .map(|c| csv_quote(&display_sheet_cell(sheet, c, true)))
            .collect();
        writeln!(w, "{}", cells.join(","))?;
    }
    Ok(())
}

pub fn markdown<W: Write>(w: &mut W, sheet: &Sheet, _opts: &RenderOptions) -> std::io::Result<()> {
    let headers: Vec<String> = sheet.headers().iter().map(|h| markdown_escape(h)).collect();
    if headers.is_empty() {
        writeln!(w, "(empty sheet)")?;
        return Ok(());
    }
    writeln!(w, "| {} |", headers.join(" | "))?;
    writeln!(w, "|{}|", vec!["---"; headers.len()].join("|"))?;
    for row in sheet.rows().iter().skip(1) {
        let cells: Vec<String> = row
            .iter()
            .map(|c| markdown_escape(&display_sheet_cell(sheet, c, true)))
            .collect();
        writeln!(w, "| {} |", cells.join(" | "))?;
    }
    Ok(())
}

/// Emit JSON in the agent-friendly shape:
/// - top-level object pretty-printed with 2-space indent
/// - `headers` array: one element per line
/// - `data` array: one row per line, each row inline `[v, v, v]` with a
///   space after every comma. The shape is a contract — token-cost figures
///   in `spreadsheet-peek/benchmarks/measure_tokens.py` are computed against
///   it — so changes here move the published "tokens per row" numbers.
pub fn json<W: Write>(w: &mut W, sheet: &Sheet, opts: &RenderOptions) -> std::io::Result<()> {
    let (total_rows, cols) = sheet.dimensions();
    let data_rows = total_rows.saturating_sub(1);
    let returned_rows = opts.max_rows.unwrap_or(data_rows).min(data_rows);
    let returned_columns = opts.max_columns.unwrap_or(cols).min(cols);
    let headers: Vec<String> = sheet.headers().into_iter().take(returned_columns).collect();
    let sheet_name_json = serde_json::to_string(&sheet.name).expect("string is JSON-safe");

    writeln!(w, "{{")?;
    writeln!(w, "  \"sheet\": {sheet_name_json},")?;
    writeln!(w, "  \"rows\": {data_rows},")?;
    writeln!(w, "  \"columns\": {cols},")?;

    if headers.is_empty() {
        writeln!(w, "  \"headers\": [],")?;
    } else {
        writeln!(w, "  \"headers\": [")?;
        for (i, h) in headers.iter().enumerate() {
            let h_json = serde_json::to_string(h).expect("string is JSON-safe");
            let comma = if i + 1 == headers.len() { "" } else { "," };
            writeln!(w, "    {h_json}{comma}")?;
        }
        writeln!(w, "  ],")?;
    }

    let body: Vec<&Vec<Cell>> = sheet.rows().iter().skip(1).take(returned_rows).collect();
    if body.is_empty() {
        writeln!(w, "  \"data\": []")?;
    } else {
        writeln!(w, "  \"data\": [")?;
        let n = body.len();
        for (i, row) in body.iter().enumerate() {
            let inline = format_row_inline(sheet, &row[..row.len().min(returned_columns)]);
            let comma = if i + 1 == n { "" } else { "," };
            writeln!(w, "    {inline}{comma}")?;
        }
        writeln!(w, "  ]")?;
    }
    writeln!(w, "}}")
}

fn format_row_inline(sheet: &Sheet, row: &[Cell]) -> String {
    let parts: Vec<String> = row
        .iter()
        .map(|c| {
            let value = if let CellValue::Error(error) = &c.value {
                Value::String(sheet.error_text(error).to_owned())
            } else {
                json_cell(c)
            };
            serde_json::to_string(&value).unwrap_or_else(|_| "null".to_string())
        })
        .collect();
    format!("[{}]", parts.join(", "))
}

pub fn boxed<W: Write>(w: &mut W, sheet: &Sheet, opts: &RenderOptions) -> std::io::Result<()> {
    let (total_rows, total_cols) = sheet.dimensions();
    let data_rows = total_rows.saturating_sub(1);
    write_banner(w)?;
    writeln!(
        w,
        "Sheet: {} ({} rows × {} columns)",
        sheet.name, data_rows, total_cols
    )?;
    writeln!(w, "Available sheets: {}", opts.all_sheets.join(", "))?;
    writeln!(w)?;

    if total_cols == 0 {
        writeln!(w, "(empty sheet)")?;
        return Ok(());
    }

    let cap_total = effective_row_cap(sheet, opts);
    let display_rows: Vec<Vec<String>> = sheet
        .rows()
        .iter()
        .take(cap_total)
        .map(|row| {
            row.iter()
                .map(|c| truncate(&display_sheet_cell(sheet, c, true), opts.max_width))
                .collect()
        })
        .collect();

    let widths = column_widths(&display_rows, total_cols);

    write_box_border(w, &widths, BorderStyle::Top)?;
    if let Some(header) = display_rows.first() {
        write_box_row(w, header, &widths, true)?;
        write_box_border(w, &widths, BorderStyle::Mid)?;
    }
    let body = display_rows.iter().skip(1);
    let body_count = display_rows.len().saturating_sub(1);
    for (i, row) in body.enumerate() {
        write_box_row(w, row, &widths, false)?;
        if i + 1 < body_count {
            write_box_border(w, &widths, BorderStyle::Mid)?;
        }
    }
    write_box_border(w, &widths, BorderStyle::Bottom)?;

    let shown_data = cap_total.saturating_sub(1);
    if data_rows > shown_data {
        writeln!(w)?;
        writeln!(
            w,
            "⚠️  Showing {shown_data} of {data_rows} rows (use -n 0 to show all)"
        )?;
    }
    Ok(())
}

fn write_banner<W: Write>(w: &mut W) -> std::io::Result<()> {
    let title = "wolfxl peek - Excel preview";
    let inner = title.len() + 4;
    let line = "═".repeat(inner);
    writeln!(w, "╔{line}╗")?;
    writeln!(w, "║  {title}  ║")?;
    writeln!(w, "╚{line}╝")?;
    writeln!(w)
}

#[derive(Copy, Clone)]
enum BorderStyle {
    Top,
    Mid,
    Bottom,
}

fn write_box_border<W: Write>(
    w: &mut W,
    widths: &[usize],
    style: BorderStyle,
) -> std::io::Result<()> {
    let (left, mid, right) = match style {
        BorderStyle::Top => ('┌', '┬', '┐'),
        BorderStyle::Mid => ('├', '┼', '┤'),
        BorderStyle::Bottom => ('└', '┴', '┘'),
    };
    write!(w, "{left}")?;
    for (i, width) in widths.iter().enumerate() {
        write!(w, "{}", "─".repeat(width + 2))?;
        let sep = if i + 1 == widths.len() { right } else { mid };
        write!(w, "{sep}")?;
    }
    writeln!(w)
}

fn write_box_row<W: Write>(
    w: &mut W,
    row: &[String],
    widths: &[usize],
    center: bool,
) -> std::io::Result<()> {
    write!(w, "│")?;
    for (i, width) in widths.iter().enumerate() {
        let cell = row.get(i).map(String::as_str).unwrap_or("");
        let cell_width = UnicodeWidthStr::width(cell);
        let pad = width.saturating_sub(cell_width);
        if center {
            let left = pad / 2;
            let right = pad - left;
            write!(w, " {}{}{} │", " ".repeat(left), cell, " ".repeat(right))?;
        } else {
            write!(w, " {}{} │", cell, " ".repeat(pad))?;
        }
    }
    writeln!(w)
}

fn column_widths(rows: &[Vec<String>], cols: usize) -> Vec<usize> {
    let mut widths = vec![0usize; cols];
    for row in rows {
        for (i, cell) in row.iter().enumerate() {
            if i >= cols {
                break;
            }
            let w = UnicodeWidthStr::width(cell.as_str());
            if w > widths[i] {
                widths[i] = w;
            }
        }
    }
    for w in widths.iter_mut() {
        if *w == 0 {
            *w = 1;
        }
    }
    widths
}

fn effective_row_cap(sheet: &Sheet, opts: &RenderOptions) -> usize {
    let total = sheet.rows().len();
    match opts.max_rows {
        None => total,
        // box renderer counts the header row inside the cap, so `-n 5`
        // shows 1 header + 5 data rows.
        Some(n) => (n + 1).min(total),
    }
}

fn truncate(s: &str, max: usize) -> String {
    let width = UnicodeWidthStr::width(s);
    if width <= max {
        return s.to_string();
    }
    let mut out = String::new();
    let mut acc = 0usize;
    let cap = max.saturating_sub(1);
    for ch in s.chars() {
        let cw = UnicodeWidthStr::width(ch.to_string().as_str());
        if acc + cw > cap {
            break;
        }
        out.push(ch);
        acc += cw;
    }
    out.push('…');
    out
}

fn csv_quote(s: &str) -> String {
    // RFC 4180 trigger set: `,`, `"`, `\r`, `\n`. The carriage return matters
    // because Excel cells with line breaks store `\r\n` — quoting only on `\n`
    // would let a stray `\r` slip through and shred downstream parsers.
    if s.contains(',') || s.contains('"') || s.contains('\n') || s.contains('\r') {
        let escaped = s.replace('"', "\"\"");
        format!("\"{escaped}\"")
    } else {
        s.to_string()
    }
}

fn markdown_escape(s: &str) -> String {
    s.replace("\r\n", "\n")
        .replace('\\', "\\\\")
        .replace('|', "\\|")
        .replace('\t', "\\t")
        .replace('\r', "\\n")
        .replace('\n', "\\n")
}

fn display_sheet_cell(sheet: &Sheet, cell: &Cell, group_ints: bool) -> String {
    if let CellValue::Error(error) = &cell.value {
        format!("ERROR: {}", sheet.error_text(error))
    } else {
        display_cell(cell, group_ints)
    }
}

fn display_cell(cell: &Cell, group_ints: bool) -> String {
    if let Some(rendered) = display_number_format(cell) {
        return rendered;
    }
    match &cell.value {
        CellValue::Empty => String::new(),
        CellValue::String(s) => s.clone(),
        // Lowercase `true`/`false` so JSON and text exports agree.
        CellValue::Bool(b) => if *b { "true" } else { "false" }.to_string(),
        CellValue::Int(n) => {
            if group_ints {
                group_thousands_signed(*n)
            } else {
                n.to_string()
            }
        }
        CellValue::Float(n) => trim_float(*n),
        CellValue::Date(d) => d.format("%Y-%m-%d").to_string(),
        CellValue::DateTime(dt) => dt.format("%Y-%m-%d %H:%M:%S").to_string(),
        CellValue::Time(t) => t.format("%H:%M:%S").to_string(),
        // Prefix Excel error sentinels with "ERROR: " so a `#REF!` cell is
        // visually distinct from a literal string `"#REF!"` in the output.
        CellValue::Error(e) => format!("ERROR: {e}"),
    }
}

fn display_number_format(cell: &Cell) -> Option<String> {
    let category = cell.number_format.as_deref().map(classify_format)?;
    match category {
        // Keep the long-standing plain-number contract for General /
        // Integer / Float. Route formats where the symbol or scale is
        // essential through the core formatter.
        FormatCategory::Currency | FormatCategory::Percentage | FormatCategory::Scientific => {
            Some(format_cell(cell))
        }
        _ => None,
    }
}

fn json_cell(cell: &Cell) -> Value {
    match &cell.value {
        CellValue::Empty => Value::Null,
        CellValue::String(s) => Value::String(s.clone()),
        CellValue::Bool(b) => Value::Bool(*b),
        CellValue::Int(n) => Value::Number((*n).into()),
        CellValue::Float(n) => serde_json::Number::from_f64(*n)
            .map(Value::Number)
            .unwrap_or(Value::Null),
        CellValue::Date(d) => Value::String(d.format("%Y-%m-%d").to_string()),
        CellValue::DateTime(dt) => Value::String(dt.format("%Y-%m-%dT%H:%M:%S").to_string()),
        CellValue::Time(t) => Value::String(t.format("%H:%M:%S").to_string()),
        CellValue::Error(e) => Value::String(e.clone()),
    }
}

fn group_thousands_signed(n: i64) -> String {
    let sign = if n < 0 { "-" } else { "" };
    let abs = if n == i64::MIN {
        // Avoid overflow on negation of i64::MIN.
        (i64::MAX as u64) + 1
    } else {
        n.unsigned_abs()
    };
    format!("{sign}{}", group_thousands(abs))
}

fn group_thousands(mut n: u64) -> String {
    if n == 0 {
        return "0".to_string();
    }
    let mut parts: Vec<String> = Vec::new();
    while n > 0 {
        let chunk = n % 1000;
        n /= 1000;
        if n > 0 {
            parts.push(format!("{chunk:03}"));
        } else {
            parts.push(chunk.to_string());
        }
    }
    parts.reverse();
    parts.join(",")
}

fn trim_float(n: f64) -> String {
    // Float rendering contract for text/csv/box: integer-valued floats render
    // with no decimals (`{n:.0}`), everything else is forced to two
    // (`{n:.2}`), and the integer part gets thousand separators. This is
    // destructive precision-wise; JSON keeps full precision via serde_json.
    if !n.is_finite() {
        return n.to_string();
    }
    let formatted = if n.fract() == 0.0 && n.abs() < 1e15 {
        format!("{n:.0}")
    } else {
        format!("{n:.2}")
    };
    // Split on the decimal point (if any) and group the integer part.
    let (int_part, frac_part) = match formatted.split_once('.') {
        Some((i, f)) => (i, Some(f)),
        None => (formatted.as_str(), None),
    };
    let negative = int_part.starts_with('-');
    let digits = int_part.trim_start_matches('-');
    let grouped = match digits.parse::<u64>() {
        Ok(n) => group_thousands(n),
        // f64 outside u64 range (>= 2^64) — fall back to ungrouped digits.
        Err(_) => digits.to_string(),
    };
    let grouped = if negative {
        format!("-{grouped}")
    } else {
        grouped
    };
    match frac_part {
        Some(f) => format!("{grouped}.{f}"),
        None => grouped,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn csv_quote_only_when_needed() {
        assert_eq!(csv_quote("plain"), "plain");
        assert_eq!(csv_quote("420,000"), "\"420,000\"");
        assert_eq!(csv_quote("a\"b"), "\"a\"\"b\"");
    }

    #[test]
    fn integer_grouping() {
        assert_eq!(group_thousands_signed(0), "0");
        assert_eq!(group_thousands_signed(1234567), "1,234,567");
        assert_eq!(group_thousands_signed(-12345), "-12,345");
    }

    #[test]
    fn truncate_respects_unicode_width() {
        assert_eq!(truncate("hello", 10), "hello");
        assert_eq!(truncate("supercalifragilistic", 8), "superca…");
    }

    #[test]
    fn trim_float_matches_format_contract() {
        // Integer-valued -> {:.0}, otherwise -> {:.2}, then thousand-separate
        // the integer part. Locked in by golden CSV/text fixtures.
        assert_eq!(trim_float(3.0), "3");
        assert_eq!(trim_float(3.14), "3.14");
        assert_eq!(trim_float(3.14159), "3.14");
        assert_eq!(trim_float(-2.0), "-2");
        assert_eq!(trim_float(2505.15), "2,505.15");
        assert_eq!(trim_float(1234567.5), "1,234,567.50");
        assert_eq!(trim_float(-1184.73), "-1,184.73");
    }

    #[test]
    fn csv_quotes_on_carriage_return() {
        assert_eq!(csv_quote("a\rb"), "\"a\rb\"");
        assert_eq!(csv_quote("a\nb"), "\"a\nb\"");
    }

    #[test]
    fn markdown_escape_keeps_table_shape() {
        assert_eq!(markdown_escape("plain"), "plain");
        assert_eq!(markdown_escape("A|B"), "A\\|B");
        assert_eq!(markdown_escape("back\\slash"), "back\\\\slash");
        assert_eq!(markdown_escape("quote \"ok\""), "quote \"ok\"");
        assert_eq!(markdown_escape("a\tb"), "a\\tb");
        assert_eq!(markdown_escape("line one\nline two"), "line one\\nline two");
        assert_eq!(
            markdown_escape("line one\r\nline two"),
            "line one\\nline two"
        );
        assert_eq!(markdown_escape("a\rb"), "a\\nb");
    }

    #[test]
    fn markdown_empty_sheet_is_explicit() {
        let sheet = Sheet::from_rows("Blank".to_string(), Vec::new());
        let mut out = Vec::new();
        markdown(
            &mut out,
            &sheet,
            &RenderOptions {
                max_rows: Some(5),
                max_columns: None,
                max_width: 20,
                all_sheets: &[],
            },
        )
        .unwrap();
        assert_eq!(String::from_utf8(out).unwrap(), "(empty sheet)\n");
    }

    #[test]
    fn display_cell_bool_lowercase_and_error_prefix() {
        let true_cell = Cell {
            value: CellValue::Bool(true),
            number_format: None,
        };
        let false_cell = Cell {
            value: CellValue::Bool(false),
            number_format: None,
        };
        let err_cell = Cell {
            value: CellValue::Error("#REF!".to_string()),
            number_format: None,
        };
        assert_eq!(display_cell(&true_cell, true), "true");
        assert_eq!(display_cell(&false_cell, true), "false");
        assert_eq!(display_cell(&err_cell, true), "ERROR: #REF!");
    }

    #[test]
    fn display_cell_respects_currency_and_percentage_formats() {
        let currency = Cell {
            value: CellValue::Float(1234.5),
            number_format: Some("$#,##0.00".to_string()),
        };
        let percentage = Cell {
            value: CellValue::Float(0.234),
            number_format: Some("0.0%".to_string()),
        };
        assert_eq!(display_cell(&currency, true), "$1,234.50");
        assert_eq!(display_cell(&percentage, true), "23.4%");
    }

    #[test]
    fn display_cell_preserves_non_usd_currency_symbol() {
        let euro = Cell {
            value: CellValue::Float(1234.5),
            number_format: Some("€#,##0.00".to_string()),
        };
        assert_eq!(display_cell(&euro, true), "€1,234.50");
    }
}
