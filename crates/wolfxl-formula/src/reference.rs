//! Parse and re-emit A1-syntax references.
//!
//! This module operates on the *value string* of a tokenizer
//! [`crate::Token`] whose subkind is [`crate::TokenSubKind::Range`].
//! It classifies the value into one of the [`RefKind`] variants and
//! provides round-trip emission via [`RefKind::render`].

use crate::{MAX_COL, MAX_ROW};

/// One A1 cell coordinate (row + col), each with an optional `$` marker.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct A1Cell {
    /// 1-based row (1..=1_048_576).
    pub row: u32,
    /// 1-based column (1..=16_384).
    pub col: u32,
    /// True if the source had `$` before the column letters.
    pub col_abs: bool,
    /// True if the source had `$` before the row digits.
    pub row_abs: bool,
}

/// A whole-row endpoint: just `$?5`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct A1Row {
    /// 1-based row.
    pub row: u32,
    /// True if `$` prefixed.
    pub abs: bool,
}

/// A whole-col endpoint: just `$?C`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct A1Col {
    /// 1-based col.
    pub col: u32,
    /// True if `$` prefixed.
    pub abs: bool,
}

/// One end of a range — kept distinct from a free-standing cell so that
/// range translation can be expressed cleanly in [`crate::translate`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct A1RangeEndpoint {
    /// The cell coordinate.
    pub cell: A1Cell,
}

/// Structured table reference parts. Pass-through only in this RFC; the
/// future table-rename RFC will rewrite [`TableRef::table`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TableRef {
    /// Table name (the identifier preceding `[`).
    pub table: String,
    /// The full bracket suffix, including outer `[` and `]`. Preserved
    /// verbatim (e.g. `[[#This Row], [Col1]]`).
    pub specifier: String,
}

/// What kind of reference a `Range`-subkind operand decomposed into.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum RefKind {
    /// Single cell with optional sheet prefix.
    Cell {
        /// Optional sheet prefix.
        sheet: Option<SheetPrefix>,
        /// The cell coordinate.
        cell: A1Cell,
    },
    /// Cell-to-cell range with optional sheet prefix on the left endpoint.
    Range {
        /// Optional sheet prefix.
        sheet: Option<SheetPrefix>,
        /// Left-hand endpoint.
        lhs: A1Cell,
        /// Right-hand endpoint.
        rhs: A1Cell,
    },
    /// Whole-row range like `1:5`.
    RowRange {
        /// Optional sheet prefix.
        sheet: Option<SheetPrefix>,
        /// Left-hand row.
        lhs: A1Row,
        /// Right-hand row.
        rhs: A1Row,
    },
    /// Whole-col range like `A:C`.
    ColRange {
        /// Optional sheet prefix.
        sheet: Option<SheetPrefix>,
        /// Left-hand col.
        lhs: A1Col,
        /// Right-hand col.
        rhs: A1Col,
    },
    /// Structured table reference (`Table1[Col1]`).
    Table(TableRef),
    /// External-workbook reference (`[Book2.xlsx]Sheet1!A1`) — pass-through.
    ExternalBook {
        /// Original raw value, preserved verbatim.
        raw: String,
    },
    /// Defined name or any other identifier we cannot classify.
    Name(String),
    /// `#REF!` (or other error) appearing where a ref would.
    Error(String),
}

/// A sheet prefix: name + whether it was originally quoted in the source.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SheetPrefix {
    /// Decoded sheet name (with `''` collapsed to `'` if it was quoted).
    pub name: String,
    /// True if the source had single-quotes around the name.
    pub quoted: bool,
}

impl SheetPrefix {
    /// Render the prefix exactly as it should appear in formula text,
    /// including the trailing `!`. Preserve source quotes when present,
    /// and add quotes when the name requires them.
    /// We deliberately drop unnecessary quotes after a rename: when the
    /// source was `'Old Name'!A1` and `Old Name` was renamed to `New`,
    /// we emit `New!A1` rather than `'New'!A1`.
    pub fn render(&self) -> String {
        let needs_quote = sheet_name_needs_quoting(&self.name);
        if self.quoted || needs_quote {
            let escaped = self.name.replace('\'', "''");
            format!("'{}'!", escaped)
        } else {
            format!("{}!", self.name)
        }
    }
}

/// Returns true if a sheet name can NOT be written without single-quote
/// wrapping in a formula. Conservatively quotes if the name has anything
/// other than `[A-Za-z_][A-Za-z0-9_]*`. This matches Excel's behavior.
pub fn sheet_name_needs_quoting(name: &str) -> bool {
    if name.is_empty() {
        return true;
    }
    let mut chars = name.chars();
    let first = chars.next().unwrap();
    if !first.is_ascii_alphabetic() && first != '_' {
        return true;
    }
    for c in chars {
        if !(c.is_ascii_alphanumeric() || c == '_') {
            return true;
        }
    }
    false
}

/// Try to parse an A1-syntax reference value (the `value` field of a
/// `Range`-subkind operand token).
///
/// Returns one of the [`RefKind`] variants, or [`RefKind::Name`] if the
/// value doesn't match any known reference shape (e.g. it's a defined
/// name like `MyTotal`).
pub fn parse_ref(value: &str) -> RefKind {
    if value.is_empty() {
        return RefKind::Name(String::new());
    }
    if let Some(stripped) = is_error_literal(value) {
        return RefKind::Error(stripped.to_string());
    }
    if value.starts_with('[') {
        return RefKind::ExternalBook {
            raw: value.to_string(),
        };
    }
    if value.starts_with('\'') && value.contains("[") && value.contains(".xlsx") {
        return RefKind::ExternalBook {
            raw: value.to_string(),
        };
    }

    let (sheet, rest) = strip_sheet_prefix(value);

    if let Some(bracket_pos) = rest.find('[') {
        let table = &rest[..bracket_pos];
        if !table.is_empty() && is_valid_table_name(table) {
            return RefKind::Table(TableRef {
                table: table.to_string(),
                specifier: rest[bracket_pos..].to_string(),
            });
        }
    }

    if let Some((l, r)) = split_top_level_colon(rest) {
        if let (Some(lr), Some(rr)) = (parse_row_only(l), parse_row_only(r)) {
            return RefKind::RowRange {
                sheet,
                lhs: lr,
                rhs: rr,
            };
        }
        if let (Some(lc), Some(rc)) = (parse_col_only(l), parse_col_only(r)) {
            return RefKind::ColRange {
                sheet,
                lhs: lc,
                rhs: rc,
            };
        }
        if let (Some(lhs), Some(rhs)) = (parse_cell(l), parse_cell(r)) {
            return RefKind::Range { sheet, lhs, rhs };
        }
        return RefKind::Name(value.to_string());
    }

    if let Some(cell) = parse_cell(rest) {
        return RefKind::Cell { sheet, cell };
    }

    RefKind::Name(value.to_string())
}

impl RefKind {
    /// Render this ref back to formula text.
    pub fn render(&self) -> String {
        match self {
            RefKind::Cell { sheet, cell } => {
                let mut out = String::new();
                if let Some(s) = sheet {
                    out.push_str(&s.render());
                }
                out.push_str(&render_cell(cell));
                out
            }
            RefKind::Range { sheet, lhs, rhs } => {
                let mut out = String::new();
                if let Some(s) = sheet {
                    out.push_str(&s.render());
                }
                out.push_str(&render_cell(lhs));
                out.push(':');
                out.push_str(&render_cell(rhs));
                out
            }
            RefKind::RowRange { sheet, lhs, rhs } => {
                let mut out = String::new();
                if let Some(s) = sheet {
                    out.push_str(&s.render());
                }
                out.push_str(&render_row(lhs));
                out.push(':');
                out.push_str(&render_row(rhs));
                out
            }
            RefKind::ColRange { sheet, lhs, rhs } => {
                let mut out = String::new();
                if let Some(s) = sheet {
                    out.push_str(&s.render());
                }
                out.push_str(&render_col(lhs));
                out.push(':');
                out.push_str(&render_col(rhs));
                out
            }
            RefKind::Table(t) => format!("{}{}", t.table, t.specifier),
            RefKind::ExternalBook { raw } => raw.clone(),
            RefKind::Name(s) => s.clone(),
            RefKind::Error(s) => s.clone(),
        }
    }
}

/// Render a cell with `$` markers preserved.
pub fn render_cell(c: &A1Cell) -> String {
    let mut out = String::with_capacity(8);
    if c.col_abs {
        out.push('$');
    }
    out.push_str(&col_letter(c.col));
    if c.row_abs {
        out.push('$');
    }
    out.push_str(&c.row.to_string());
    out
}

fn render_row(r: &A1Row) -> String {
    if r.abs {
        format!("${}", r.row)
    } else {
        r.row.to_string()
    }
}

fn render_col(c: &A1Col) -> String {
    if c.abs {
        format!("${}", col_letter(c.col))
    } else {
        col_letter(c.col)
    }
}

fn is_error_literal(s: &str) -> Option<&str> {
    const ERRS: &[&str] = &[
        "#NULL!",
        "#DIV/0!",
        "#VALUE!",
        "#REF!",
        "#NAME?",
        "#NUM!",
        "#N/A",
        "#GETTING_DATA",
    ];
    ERRS.iter().copied().find(|&error| s == error)
}

fn is_valid_table_name(s: &str) -> bool {
    if s.is_empty() {
        return false;
    }
    let mut it = s.chars();
    let first = it.next().unwrap();
    if !(first.is_ascii_alphabetic() || first == '_' || first == '\\') {
        return false;
    }
    it.all(|c| c.is_ascii_alphanumeric() || c == '_' || c == '.')
}

/// Split off a sheet prefix `Sheet!` or `'Sheet'!`. The `!` is consumed.
/// Returns `(prefix?, rest)`.
pub fn strip_sheet_prefix(s: &str) -> (Option<SheetPrefix>, &str) {
    let bytes = s.as_bytes();
    if bytes.is_empty() {
        return (None, s);
    }

    if bytes[0] == b'\'' {
        let mut i = 1;
        while i < bytes.len() {
            if bytes[i] == b'\'' {
                if i + 1 < bytes.len() && bytes[i + 1] == b'\'' {
                    i += 2;
                    continue;
                }
                if i + 1 < bytes.len() && bytes[i + 1] == b'!' {
                    let inner = &s[1..i];
                    let decoded = inner.replace("''", "'");
                    let rest = &s[i + 2..];
                    return (
                        Some(SheetPrefix {
                            name: decoded,
                            quoted: true,
                        }),
                        rest,
                    );
                }
                return (None, s);
            }
            i += 1;
        }
        return (None, s);
    }

    let mut depth: i32 = 0;
    for (i, &b) in bytes.iter().enumerate() {
        match b {
            b'[' => depth += 1,
            b']' => depth -= 1,
            b'!' if depth == 0 => {
                let name = &s[..i];
                if !name.is_empty() && is_valid_unquoted_sheet_name(name) {
                    return (
                        Some(SheetPrefix {
                            name: name.to_string(),
                            quoted: false,
                        }),
                        &s[i + 1..],
                    );
                }
                return (None, s);
            }
            _ => {}
        }
    }
    (None, s)
}

fn is_valid_unquoted_sheet_name(s: &str) -> bool {
    !sheet_name_needs_quoting(s)
}

/// Parse a single cell coordinate, e.g. `A1`, `$A$1`, `$B5`.
pub fn parse_cell(s: &str) -> Option<A1Cell> {
    let bytes = s.as_bytes();
    if bytes.is_empty() {
        return None;
    }
    let mut i = 0;
    let col_abs = if bytes[i] == b'$' {
        i += 1;
        true
    } else {
        false
    };
    let col_start = i;
    while i < bytes.len() && bytes[i].is_ascii_alphabetic() {
        i += 1;
    }
    if i == col_start {
        return None;
    }
    let col_str = &s[col_start..i];
    if col_str.len() > 3 {
        return None;
    }
    let col = col_letters_to_num(col_str)?;
    if col == 0 || col > MAX_COL {
        return None;
    }
    let row_abs = if i < bytes.len() && bytes[i] == b'$' {
        i += 1;
        true
    } else {
        false
    };
    let row_start = i;
    while i < bytes.len() && bytes[i].is_ascii_digit() {
        i += 1;
    }
    if i == row_start || i != bytes.len() {
        return None;
    }
    let row_str = &s[row_start..i];
    let row: u32 = row_str.parse().ok()?;
    if row == 0 || row > MAX_ROW {
        return None;
    }
    Some(A1Cell {
        row,
        col,
        col_abs,
        row_abs,
    })
}

fn parse_row_only(s: &str) -> Option<A1Row> {
    let bytes = s.as_bytes();
    if bytes.is_empty() {
        return None;
    }
    let mut i = 0;
    let abs = if bytes[i] == b'$' {
        i += 1;
        true
    } else {
        false
    };
    if i == bytes.len() {
        return None;
    }
    let n: u32 = s[i..].parse().ok()?;
    if n == 0 || n > MAX_ROW {
        return None;
    }
    Some(A1Row { row: n, abs })
}

fn parse_col_only(s: &str) -> Option<A1Col> {
    let bytes = s.as_bytes();
    if bytes.is_empty() {
        return None;
    }
    let mut i = 0;
    let abs = if bytes[i] == b'$' {
        i += 1;
        true
    } else {
        false
    };
    let rest = &s[i..];
    if rest.is_empty() || rest.len() > 3 {
        return None;
    }
    if !rest.bytes().all(|b| b.is_ascii_alphabetic()) {
        return None;
    }
    let col = col_letters_to_num(rest)?;
    if col == 0 || col > MAX_COL {
        return None;
    }
    Some(A1Col { col, abs })
}

/// Parse one half of a range string into either a [`A1Cell`], [`A1Row`],
/// or [`A1Col`] — useful for the translator to handle each endpoint
/// independently. Returned variants are wrapped as a [`RefKind`] for
/// consistency.
pub fn parse_range_part(s: &str) -> Option<RefKind> {
    if let Some(c) = parse_cell(s) {
        return Some(RefKind::Cell {
            sheet: None,
            cell: c,
        });
    }
    None
}

/// Convert column letters (`"A"`, `"AA"`, …) to 1-based index.
pub fn col_letters_to_num(s: &str) -> Option<u32> {
    if s.is_empty() {
        return None;
    }
    let mut n: u32 = 0;
    for c in s.chars() {
        if !c.is_ascii_alphabetic() {
            return None;
        }
        let v = (c.to_ascii_uppercase() as u32) - ('A' as u32) + 1;
        n = n.checked_mul(26)?.checked_add(v)?;
    }
    Some(n)
}

/// Convert 1-based column index back to letters (`1` → `"A"`, `27` → `"AA"`).
pub fn col_letter(mut n: u32) -> String {
    let mut out = Vec::new();
    while n > 0 {
        let rem = (n - 1) % 26;
        out.push((b'A' + rem as u8) as char);
        n = (n - 1) / 26;
    }
    out.iter().rev().collect()
}

/// Find a `:` that's at the top level (not inside `[...]`).
fn split_top_level_colon(s: &str) -> Option<(&str, &str)> {
    let bytes = s.as_bytes();
    let mut depth: i32 = 0;
    for (i, &b) in bytes.iter().enumerate() {
        match b {
            b'[' => depth += 1,
            b']' => depth -= 1,
            b':' if depth == 0 => {
                return Some((&s[..i], &s[i + 1..]));
            }
            _ => {}
        }
    }
    None
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn col_letters_roundtrip() {
        for n in [1, 2, 26, 27, 52, 53, 702, 703, 16384] {
            let s = col_letter(n);
            assert_eq!(col_letters_to_num(&s).unwrap(), n, "{}", s);
        }
        assert_eq!(col_letter(1), "A");
        assert_eq!(col_letter(26), "Z");
        assert_eq!(col_letter(27), "AA");
        assert_eq!(col_letter(703), "AAA");
        assert_eq!(col_letter(16384), "XFD");
    }

    #[test]
    fn parse_simple_cell() {
        let c = parse_cell("A1").unwrap();
        assert_eq!(
            c,
            A1Cell {
                row: 1,
                col: 1,
                col_abs: false,
                row_abs: false
            }
        );
        assert_eq!(render_cell(&c), "A1");
    }

    #[test]
    fn parse_full_absolute() {
        let c = parse_cell("$A$1").unwrap();
        assert!(c.col_abs && c.row_abs);
        assert_eq!(render_cell(&c), "$A$1");
    }

    #[test]
    fn parse_mixed() {
        let c = parse_cell("$B5").unwrap();
        assert!(c.col_abs && !c.row_abs);
        assert_eq!(render_cell(&c), "$B5");
        let c = parse_cell("B$5").unwrap();
        assert!(!c.col_abs && c.row_abs);
        assert_eq!(render_cell(&c), "B$5");
    }

    #[test]
    fn parse_range_basic() {
        let r = parse_ref("A1:B5");
        match r {
            RefKind::Range { sheet, lhs, rhs } => {
                assert!(sheet.is_none());
                assert_eq!(lhs.col, 1);
                assert_eq!(rhs.col, 2);
                assert_eq!(rhs.row, 5);
            }
            _ => panic!("not range"),
        }
        assert_eq!(parse_ref("A1:B5").render(), "A1:B5");
    }

    #[test]
    fn parse_whole_row_range() {
        let r = parse_ref("2:5");
        match r {
            RefKind::RowRange { lhs, rhs, .. } => {
                assert_eq!(lhs.row, 2);
                assert_eq!(rhs.row, 5);
            }
            _ => panic!("not row-range"),
        }
        assert_eq!(parse_ref("2:5").render(), "2:5");
    }

    #[test]
    fn parse_whole_col_range() {
        let r = parse_ref("A:C");
        match r {
            RefKind::ColRange { lhs, rhs, .. } => {
                assert_eq!(lhs.col, 1);
                assert_eq!(rhs.col, 3);
            }
            _ => panic!("not col-range"),
        }
        assert_eq!(parse_ref("A:C").render(), "A:C");
    }

    #[test]
    fn parse_3d_unquoted() {
        let r = parse_ref("Sheet2!A1");
        match r {
            RefKind::Cell { sheet, cell } => {
                let s = sheet.unwrap();
                assert_eq!(s.name, "Sheet2");
                assert!(!s.quoted);
                assert_eq!(cell.col, 1);
            }
            _ => panic!(),
        }
        assert_eq!(parse_ref("Sheet2!A1").render(), "Sheet2!A1");
    }

    #[test]
    fn parse_3d_quoted_with_apostrophe() {
        let r = parse_ref("'O''Brien'!A1");
        match r {
            RefKind::Cell { sheet, cell: _ } => {
                let s = sheet.unwrap();
                assert_eq!(s.name, "O'Brien");
                assert!(s.quoted);
            }
            _ => panic!("not cell"),
        }
        assert_eq!(parse_ref("'O''Brien'!A1").render(), "'O''Brien'!A1");
    }

    #[test]
    fn parse_3d_quoted_with_space() {
        let r = parse_ref("'2024 Data'!$A$1");
        let rendered = r.render();
        assert_eq!(rendered, "'2024 Data'!$A$1");
    }

    #[test]
    fn parse_table_ref_simple() {
        let r = parse_ref("Table1[Col1]");
        match r {
            RefKind::Table(t) => {
                assert_eq!(t.table, "Table1");
                assert_eq!(t.specifier, "[Col1]");
            }
            _ => panic!(),
        }
        assert_eq!(parse_ref("Table1[Col1]").render(), "Table1[Col1]");
    }

    #[test]
    fn parse_table_ref_nested() {
        let r = parse_ref("Table1[[#This Row], [Col1]]");
        match r {
            RefKind::Table(t) => assert_eq!(t.specifier, "[[#This Row], [Col1]]"),
            _ => panic!(),
        }
    }

    #[test]
    fn parse_external_book() {
        let r = parse_ref("[Book2.xlsx]Sheet1!A1");
        assert!(matches!(r, RefKind::ExternalBook { .. }));
    }

    #[test]
    fn parse_defined_name() {
        let r = parse_ref("MyTotal");
        assert!(matches!(r, RefKind::Name(_)));
    }

    #[test]
    fn quoted_unnecessary_quotes_get_normalized() {
        let s = SheetPrefix {
            name: "Sheet1".into(),
            quoted: false,
        };
        assert_eq!(s.render(), "Sheet1!");
    }

    #[test]
    fn sheet_with_space_force_quote_even_if_unquoted_in_struct() {
        let s = SheetPrefix {
            name: "My Sheet".into(),
            quoted: false,
        };
        assert_eq!(s.render(), "'My Sheet'!");
    }
}
