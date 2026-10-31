// SPDX-License-Identifier: MIT
//! Independent RPN rendering, MS-XLS §2.5.198; legacy layouts from
//! OpenOffice Excel format §§3.1–3.11. No evaluation or cache substitution.
use super::commands::command;
use super::functions::function;
use super::{bad, error_text, text, u16_at, Result};

#[derive(Clone, Copy)]
pub(super) struct Context<'a> {
    pub version: u8,
    pub page: u16,
    pub sheets: &'a [super::Sheet],
    pub names: &'a [String],
    pub xti: &'a [(i16, i16)],
    pub old_extern: &'a [String],
    pub extern_names: &'a [Vec<String>],
    pub xti_books: &'a [usize],
    pub books: &'a [Vec<String>],
    pub deleted_labels: &'a [String],
    pub labels: Option<&'a std::collections::BTreeMap<(u32, u32), super::Cell>>,
}
#[derive(Clone)]
struct Expr {
    text: String,
    precedence: u8,
}
impl Expr {
    fn atom(text: String) -> Self {
        Self {
            text,
            precedence: 100,
        }
    }
}
fn pop(stack: &mut Vec<Expr>) -> Result<Expr> {
    stack
        .pop()
        .ok_or_else(|| bad("XLS formula stack underflow"))
}
fn wrapped(e: Expr, p: u8) -> String {
    if e.precedence < p {
        format!("({})", e.text)
    } else {
        e.text
    }
}

pub(super) fn decode(
    tokens: &[u8],
    extra: &[u8],
    ctx: &Context<'_>,
    row: u32,
    col: u32,
    relative: bool,
) -> Result<String> {
    if tokens.is_empty() {
        return Ok(String::new());
    }
    let mut input = Input {
        data: tokens,
        pos: 0,
    };
    let mut arrays = Input {
        data: extra,
        pos: 0,
    };
    let mut stack: Vec<Expr> = Vec::new();
    while input.pos < tokens.len() {
        let raw = input.byte()?;
        let t = if raw >= 0x20 {
            (raw & 0x1f) | 0x20
        } else {
            raw
        };
        match t {
            0x03..=0x11 => {
                let (op, p) = match t {
                    3 => ("+", 30),
                    4 => ("-", 30),
                    5 => ("*", 40),
                    6 => ("/", 40),
                    7 => ("^", 50),
                    8 => ("&", 20),
                    9 => ("<", 10),
                    10 => ("<=", 10),
                    11 => ("=", 10),
                    12 => (">=", 10),
                    13 => (">", 10),
                    14 => ("<>", 10),
                    15 => (" ", 70),
                    16 => (",", 60),
                    _ => (":", 80),
                };
                let rhs = pop(&mut stack)?;
                let lhs = pop(&mut stack)?;
                let r = wrapped(rhs, p + u8::from(matches!(t, 4 | 6 | 7 | 9..=14)));
                stack.push(Expr {
                    text: format!("{}{op}{r}", wrapped(lhs, p)),
                    precedence: p,
                });
            }
            0x12 | 0x13 => {
                let a = pop(&mut stack)?;
                stack.push(Expr {
                    text: format!("{}{}", if t == 0x12 { "+" } else { "-" }, wrapped(a, 90)),
                    precedence: 90,
                });
            }
            0x14 => {
                let a = pop(&mut stack)?;
                stack.push(Expr {
                    text: format!("{}%", wrapped(a, 95)),
                    precedence: 95,
                });
            }
            0x15 => {
                let a = pop(&mut stack)?;
                stack.push(Expr::atom(format!("({})", a.text)));
            }
            0x16 => stack.push(Expr::atom(String::new())),
            0x17 => {
                let n = input.byte()? as usize;
                let s = if ctx.version == 8 {
                    let f = input.byte()?;
                    if f > 1 {
                        return Err(bad("invalid formula string flags"));
                    }
                    let d = input.take(n * if f == 1 { 2 } else { 1 })?;
                    if f == 1 {
                        text::utf16(d)?
                    } else {
                        d.iter().map(|&v| char::from(v)).collect()
                    }
                } else {
                    text::legacy(input.take(n)?, ctx.page)?
                };
                stack.push(Expr::atom(format!("\"{}\"", s.replace('"', "\"\""))));
            }
            0x18 => {
                if ctx.version != 8 {
                    return Err(bad("natural-language token outside BIFF8"));
                }
                stack.push(Expr::atom(natural_label(
                    &mut input,
                    &mut arrays,
                    ctx,
                    row,
                    col,
                )?));
            }
            0x19 => {
                let flags = input.byte()?;
                let n = if ctx.version == 2 {
                    input.byte()? as usize
                } else {
                    input.word()? as usize
                };
                if flags & 4 != 0 {
                    input.take((n + 1) * if ctx.version == 2 { 1 } else { 2 })?;
                }
                if flags & 0x10 != 0 {
                    let a = pop(&mut stack)?;
                    stack.push(Expr::atom(format!("SUM({})", a.text)));
                }
            }
            0x1c => stack.push(Expr::atom(error_text(input.byte()?)?.to_string())),
            0x1d => {
                let v = input.byte()?;
                if v > 1 {
                    return Err(bad("invalid formula Boolean"));
                }
                stack.push(Expr::atom(if v == 1 { "TRUE" } else { "FALSE" }.into()));
            }
            0x1e => stack.push(Expr::atom(input.word()?.to_string())),
            0x1f => {
                let v = f64::from_le_bytes(input.take(8)?.try_into().unwrap());
                if !v.is_finite() {
                    return Err(bad("nonfinite formula literal"));
                }
                stack.push(Expr::atom(v.to_string()));
            }
            0x20 => {
                input.take(if ctx.version == 2 { 6 } else { 7 })?;
                let cols = arrays.byte()? as usize + if ctx.version == 8 { 1 } else { 0 };
                let cols = if cols == 0 { 256 } else { cols };
                let rows = arrays.word()? as usize + if ctx.version == 8 { 1 } else { 0 };
                if cols * rows > 1_000_000 {
                    return Err(bad("XLS formula array too large"));
                }
                let mut s = String::from("{");
                for r in 0..rows {
                    if r > 0 {
                        s.push(';');
                    }
                    for c in 0..cols {
                        if c > 0 {
                            s.push(',');
                        }
                        let t = arrays.byte()?;
                        match t {
                            0 => {
                                arrays.take(8)?;
                            }
                            1 => {
                                let v = f64::from_le_bytes(arrays.take(8)?.try_into().unwrap());
                                s.push_str(&v.to_string());
                            }
                            2 => {
                                let n = if ctx.version == 8 {
                                    arrays.word()? as usize
                                } else {
                                    arrays.byte()? as usize
                                };
                                let text = if ctx.version == 8 {
                                    let f = arrays.byte()?;
                                    if f > 1 {
                                        return Err(bad("invalid array string flags"));
                                    }
                                    let d = arrays.take(n * if f == 1 { 2 } else { 1 })?;
                                    if f == 1 {
                                        text::utf16(d)?
                                    } else {
                                        d.iter().map(|&v| char::from(v)).collect()
                                    }
                                } else {
                                    text::legacy(arrays.take(n)?, ctx.page)?
                                };
                                s.push('"');
                                s.push_str(&text.replace('"', "\"\""));
                                s.push('"');
                            }
                            4 => {
                                let b = arrays.byte()?;
                                arrays.take(7)?;
                                s.push_str(if b == 0 { "FALSE" } else { "TRUE" });
                            }
                            16 => {
                                let b = arrays.byte()?;
                                arrays.take(7)?;
                                s.push_str(error_text(b)?);
                            }
                            _ => return Err(bad("invalid formula array constant")),
                        }
                    }
                }
                s.push('}');
                stack.push(Expr::atom(s));
            }
            0x21 | 0x22 | 0x38 => {
                let count = if t != 0x21 {
                    Some((input.byte()? & 0x7f) as usize)
                } else {
                    None
                };
                let id = if ctx.version <= 3 {
                    input.byte()? as u16
                } else {
                    input.word()?
                };
                let is_command = t == 0x38 || id & 0x8000 != 0;
                if t == 0x38 && ctx.version > 3 {
                    return Err(bad("legacy command token in newer BIFF"));
                }
                let (name, fixed) = if is_command {
                    (
                        command(id & 0x7fff).ok_or_else(|| bad("unknown XLS command id"))?,
                        None,
                    )
                } else {
                    function(id, ctx.version)
                        .ok_or_else(|| bad(&format!("unknown XLS function id {id}")))?
                };
                let n = count
                    .or(fixed)
                    .ok_or_else(|| bad("variable XLS function without argument count"))?;
                if n > stack.len() {
                    return Err(bad("XLS function stack underflow"));
                }
                let args = stack.split_off(stack.len() - n);
                let mut args: Vec<String> = args.into_iter().map(|e| e.text).collect();
                let name = if !is_command && id == 255 {
                    if args.is_empty() {
                        return Err(bad("missing user function name"));
                    }
                    args.remove(0)
                } else {
                    name.to_string()
                };
                stack.push(Expr::atom(format!("{name}({})", args.join(","))));
            }
            0x23 => {
                let n = input.word()? as usize;
                input.take(match ctx.version {
                    2 => 5,
                    3 | 4 => 8,
                    5 => 12,
                    _ => 2,
                })?;
                stack.push(Expr::atom(
                    ctx.names
                        .get(n.wrapping_sub(1))
                        .ok_or_else(|| bad("invalid XLS defined-name index"))?
                        .clone(),
                ));
            }
            0x24 | 0x2c => {
                let d = input.take(if ctx.version == 8 { 4 } else { 3 })?;
                stack.push(Expr::atom(reference(
                    d,
                    ctx.version,
                    row,
                    col,
                    relative || t == 0x2c,
                )?));
            }
            0x25 | 0x2d => {
                let d = input.take(if ctx.version == 8 { 8 } else { 6 })?;
                stack.push(Expr::atom(area(
                    d,
                    ctx.version,
                    row,
                    col,
                    relative || t == 0x2d,
                )?));
            }
            0x26 | 0x27 | 0x28 => {
                input.take(if ctx.version == 2 { 4 } else { 6 })?;
                if t == 0x26 && ctx.version == 8 {
                    let n = arrays.word()? as usize;
                    arrays.take(n * 8)?;
                }
            }
            0x29 | 0x2e | 0x2f => {
                input.take(if ctx.version == 2 { 1 } else { 2 })?;
            }
            0x2a | 0x2b => {
                input.take(if t == 0x2a {
                    if ctx.version == 8 {
                        4
                    } else {
                        3
                    }
                } else if ctx.version == 8 {
                    8
                } else {
                    6
                })?;
                stack.push(Expr::atom("#REF!".into()));
            }
            0x39 => {
                let x = input.word()?;
                let n = if ctx.version == 8 {
                    let n = input.word()? as usize;
                    input.take(2)?;
                    n
                } else {
                    input.take(8)?;
                    let n = input.word()? as usize;
                    input.take(12)?;
                    n
                };
                let name = if ctx.version == 5 && (x as i16) < 0 {
                    ctx.names.get(n.wrapping_sub(1))
                } else {
                    let book = ctx.xti_books.get(x as usize).copied().unwrap_or(x as usize);
                    ctx.extern_names
                        .get(book)
                        .and_then(|ns| ns.get(n.wrapping_sub(1)))
                }
                .ok_or_else(|| bad("invalid external name index"))?;
                stack.push(Expr::atom(name.clone()));
            }
            0x3a..=0x3d => {
                let x = input.word()?;
                let prefix = if ctx.version == 8 {
                    prefix(ctx, x as usize)?
                } else {
                    input.take(8)?;
                    let first = input.word()? as i16;
                    let last = input.word()? as i16;
                    if (x as i16) < 0 {
                        own_sheet_span(ctx.sheets, first, last)?
                    } else {
                        ctx.old_extern
                            .get(x.wrapping_sub(1) as usize)
                            .map(|s| quote_sheet(s))
                            .ok_or_else(|| bad("invalid BIFF5 external sheet"))?
                    }
                };
                let is_area = t == 0x3b || t == 0x3d;
                let d = input.take(if is_area {
                    if ctx.version == 8 {
                        8
                    } else {
                        6
                    }
                } else if ctx.version == 8 {
                    4
                } else {
                    3
                })?;
                let r = if t >= 0x3c {
                    "#REF!".into()
                } else if is_area {
                    area(d, ctx.version, row, col, relative)?
                } else {
                    reference(d, ctx.version, row, col, relative)?
                };
                stack.push(Expr::atom(format!("{prefix}!{r}")));
            }
            _ => return Err(bad(&format!("unsupported XLS formula token 0x{raw:02x}"))),
        }
        // Rendered text can expand faster than its input for repeated unions.
        if stack.len() > 8192 || stack.last().is_some_and(|e| e.text.len() > 1_048_576) {
            return Err(bad("XLS formula exceeds admission limit"));
        }
    }
    if stack.len() != 1 {
        return Err(bad("invalid final XLS formula stack"));
    }
    Ok(stack.pop().unwrap().text)
}

pub(super) fn quote_sheet(s: &str) -> String {
    if !s.is_empty()
        && s.chars()
            .all(|c| c.is_ascii_alphanumeric() || c == '_' || c == '.')
        && !s.chars().next().unwrap().is_ascii_digit()
    {
        s.into()
    } else {
        format!("'{}'", s.replace('\'', "''"))
    }
}
fn sheet_span(sheets: &[String], first: i16, last: i16) -> Result<String> {
    if first < 0 || last < 0 {
        return Ok("#REF".into());
    }
    let a = sheets
        .get(first as usize)
        .ok_or_else(|| bad("invalid XLS sheet reference"))?;
    let b = sheets
        .get(last as usize)
        .ok_or_else(|| bad("invalid XLS sheet reference"))?;
    Ok(if first == last {
        quote_sheet(a)
    } else {
        quote_sheet(&format!("{a}:{b}"))
    })
}
fn own_sheet_span(sheets: &[super::Sheet], first: i16, last: i16) -> Result<String> {
    if first < 0 || last < 0 {
        return Ok("#REF".into());
    }
    let a = &sheets
        .get(first as usize)
        .ok_or_else(|| bad("invalid XLS sheet reference"))?
        .name;
    let b = &sheets
        .get(last as usize)
        .ok_or_else(|| bad("invalid XLS sheet reference"))?
        .name;
    Ok(if first == last {
        quote_sheet(a)
    } else {
        quote_sheet(&format!("{a}:{b}"))
    })
}
fn prefix(ctx: &Context<'_>, index: usize) -> Result<String> {
    let &(a, b) = ctx
        .xti
        .get(index)
        .ok_or_else(|| bad("invalid XLS XTI index"))?;
    let book = ctx.xti_books.get(index).copied().unwrap_or(0);
    if let Some(sheets) = ctx.books.get(book).filter(|s| !s.is_empty()) {
        sheet_span(sheets, a, b)
    } else {
        own_sheet_span(ctx.sheets, a, b)
    }
}
fn reference(d: &[u8], version: u8, row: u32, col: u32, relative: bool) -> Result<String> {
    let rawrow = u16_at(d, 0)?;
    let (r, c, rr, cr) = if version == 8 {
        let c = u16_at(d, 2)?;
        (
            rawrow as i32,
            (c & 0xff) as i32,
            c & 0x8000 != 0,
            c & 0x4000 != 0,
        )
    } else {
        (
            (rawrow & 0x3fff) as i32,
            d[2] as i32,
            rawrow & 0x8000 != 0,
            rawrow & 0x4000 != 0,
        )
    };
    let r = if relative && rr {
        let delta = if version == 8 {
            r as i16 as i32
        } else {
            (r << 18) >> 18
        };
        (row as i32 + delta).rem_euclid(if version == 8 { 65536 } else { 16384 })
    } else {
        r
    };
    let c = if relative && cr {
        (col as i32 + c as i8 as i32).rem_euclid(256)
    } else {
        c
    };
    let a = crate::row_col_to_a1(r as u32 + 1, c as u32 + 1);
    let split = a.find(|c: char| c.is_ascii_digit()).unwrap();
    Ok(format!(
        "{}{}{}{}",
        if cr { "" } else { "$" },
        &a[..split],
        if rr { "" } else { "$" },
        r + 1
    ))
}
fn area(d: &[u8], version: u8, row: u32, col: u32, relative: bool) -> Result<String> {
    let (a, b) = if version == 8 {
        ([d[0], d[1], d[4], d[5]], [d[2], d[3], d[6], d[7]])
    } else {
        ([d[0], d[1], d[4], 0], [d[2], d[3], d[5], 0])
    };
    Ok(format!(
        "{}:{}",
        reference(&a, version, row, col, relative)?,
        reference(&b, version, row, col, relative)?
    ))
}
struct Input<'a> {
    data: &'a [u8],
    pos: usize,
}
impl<'a> Input<'a> {
    fn take(&mut self, n: usize) -> Result<&'a [u8]> {
        let end = self
            .pos
            .checked_add(n)
            .ok_or_else(|| bad("formula token length overflow"))?;
        let s = self.data.get(self.pos..end).ok_or_else(|| {
            super::ReaderError::Io(std::io::Error::new(
                std::io::ErrorKind::UnexpectedEof,
                "truncated XLS formula token or extra data",
            ))
        })?;
        self.pos = end;
        Ok(s)
    }
    fn byte(&mut self) -> Result<u8> {
        Ok(self.take(1)?[0])
    }
    fn word(&mut self) -> Result<u16> {
        Ok(u16::from_le_bytes(self.take(2)?.try_into().unwrap()))
    }
    fn dword(&mut self) -> Result<u32> {
        Ok(u32::from_le_bytes(self.take(4)?.try_into().unwrap()))
    }
}

// MS-XLS §2.5.198: PtgElf*, PtgExtraElf, RgceElfLoc and RgceElfLocExtra.
// Resolve the actual label cells, not an invented name or a cached result.
fn natural_label(
    input: &mut Input<'_>,
    extra: &mut Input<'_>,
    ctx: &Context<'_>,
    row: u32,
    col: u32,
) -> Result<String> {
    let subtype = input.byte()?;
    let data = input.take(4)?;
    let label = match subtype {
        0x01 | 0x10 => {
            let index = u16_at(data, 0)? as usize;
            if index < 2 {
                "#NAME?".to_string()
            } else {
                let name = ctx
                    .deleted_labels
                    .get(index - 2)
                    .ok_or_else(|| bad("deleted label index outside Lel collection"))?;
                label_text(name, u16_at(data, 2)? & 1 != 0)
            }
        }
        0x02 | 0x03 | 0x06 | 0x07 | 0x0a => {
            let r = u16_at(data, 0)?;
            let c = u16_at(data, 2)?;
            let quoted = c & 0x4000 != 0;
            let (r, c) = label_position(r, c & 0x3fff, c & 0x8000 != 0, row, col)?;
            cell_label(ctx, r, c, quoted)?
        }
        0x0b | 0x0d | 0x0f => {
            let flags = extra.dword()?;
            let count = (flags & 0x3fff_ffff) as usize;
            if count == 0 || count > 65536 {
                return Err(bad("invalid natural-language label count"));
            }
            let mut names = Vec::with_capacity(count);
            for _ in 0..count {
                let r = extra.word()?;
                let c = extra.word()? & 0x3fff;
                let (r, c) = label_position(r, c, flags & 0x8000_0000 != 0, row, col)?;
                names.push(cell_label(ctx, r, c, true)?);
            }
            names.join(" ")
        }
        _ => return Err(bad("invalid natural-language token subtype")),
    };
    if matches!(subtype, 0x0a | 0x0b | 0x10) {
        let t = input.byte()?;
        if !matches!(t, 0x25 | 0x45 | 0x65 | 0x2b | 0x4b | 0x6b) {
            return Err(bad("natural-language radical missing area"));
        }
        input.take(8)?;
    }
    Ok(label)
}
fn label_position(r: u16, c: u16, relative: bool, row: u32, col: u32) -> Result<(u32, u32)> {
    if c > 255 {
        return Err(bad("natural-language label column outside BIFF grid"));
    }
    Ok(if relative {
        (
            ((row as i32 + r as i16 as i32).rem_euclid(65536)) as u32,
            ((col as i32 + c as u8 as i8 as i32).rem_euclid(256)) as u32,
        )
    } else {
        (r as u32, c as u32)
    })
}
fn label_text(name: &str, quoted: bool) -> String {
    if !quoted {
        return name.to_string();
    }
    let mut text = String::with_capacity(name.len() + 2);
    text.push('\'');
    for ch in name.chars() {
        if ch == '\'' {
            text.push('\'');
        }
        text.push(ch);
    }
    text.push('\'');
    text
}
fn cell_label(ctx: &Context<'_>, row: u32, col: u32, quoted: bool) -> Result<String> {
    let cell = ctx
        .labels
        .and_then(|cells| cells.get(&(row, col)))
        .ok_or_else(|| bad("natural-language label references missing cell"))?;
    match &cell.value {
        super::CellValue::String(s) => Ok(label_text(s, quoted)),
        super::CellValue::SharedString(s) => Ok(label_text(s, quoted)),
        super::CellValue::Number(n) => Ok(if quoted {
            format!("'{n}'")
        } else {
            n.to_string()
        }),
        super::CellValue::Bool(b) => Ok(label_text(if *b { "TRUE" } else { "FALSE" }, quoted)),
        _ => Err(bad("natural-language label cell has no label text")),
    }
}
