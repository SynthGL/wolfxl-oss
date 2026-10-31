// SPDX-License-Identifier: MIT
//! Independently authored Community BIFF reader. Format provenance: MS-XLS
//! (2025-08-19), §§2.1, 2.4 and 2.5.198; MS-CFB §§2.2–2.6. Legacy BIFF2–5
//! layouts: OpenOffice.org Excel File Format rev.1.42 §§2–5 (format facts).
//! Uses only Community models plus character codecs, not another file parser.
mod cfb;
mod commands;
mod formula;
mod functions;
mod text;

use crate::{
    Cell, CellDataType, CellValue, NamedRange, ReaderError, Result, SheetState, WorksheetData,
};
use std::collections::{BTreeMap, HashMap, HashSet};
use std::fs::File;
use std::io::Read;
use std::path::Path;
use std::sync::Arc;

const MAX_BYTES: usize = 512 * 1024 * 1024;
const MAX_CELLS: usize = 8_000_000;
fn bad(message: &str) -> ReaderError {
    ReaderError::Unsupported(format!("XLS format error: {message}"))
}
fn bytes(data: &[u8], offset: usize, n: usize) -> Result<&[u8]> {
    data.get(
        offset
            ..offset
                .checked_add(n)
                .ok_or_else(|| bad("length overflow"))?,
    )
    .ok_or_else(|| bad("truncated record/stream"))
}
fn u16_at(data: &[u8], p: usize) -> Result<u16> {
    Ok(u16::from_le_bytes(bytes(data, p, 2)?.try_into().unwrap()))
}
fn u32_at(data: &[u8], p: usize) -> Result<u32> {
    Ok(u32::from_le_bytes(bytes(data, p, 4)?.try_into().unwrap()))
}
fn error_text(code: u8) -> Result<&'static str> {
    match code {
        0 => Ok("#NULL!"),
        7 => Ok("#DIV/0!"),
        15 => Ok("#VALUE!"),
        23 => Ok("#REF!"),
        29 => Ok("#NAME?"),
        36 => Ok("#NUM!"),
        42 => Ok("#N/A"),
        43 => Ok("#GETTING_DATA"),
        _ => Err(bad("invalid BIFF error code")),
    }
}

#[derive(Debug, Clone)]
pub struct NativeXlsBook {
    stream: Vec<u8>,
    sheets: Vec<Sheet>,
    names: Vec<NamedRange>,
    formula_names: Vec<String>,
    strings: Vec<Arc<str>>,
    formats: HashMap<u16, String>,
    xfs: Vec<u16>,
    font_pages: Vec<Option<u16>>,
    xf_fonts: Vec<u16>,
    deleted_labels: Vec<String>,
    version: u8,
    page: u16,
    epoch1904: bool,
    xti: Vec<(i16, i16)>,
    xti_books: Vec<usize>,
    books: Vec<Vec<String>>,
    old_extern: Vec<String>,
    extern_names: Vec<Vec<String>>,
}
#[derive(Debug, Clone)]
struct Sheet {
    name: String,
    offset: usize,
    state: SheetState,
}
#[derive(Clone, Copy)]
struct Record<'a> {
    id: u16,
    data: &'a [u8],
    offset: usize,
}
fn records(data: &[u8], start: usize) -> Result<Vec<Record<'_>>> {
    let mut pos = start;
    let mut depth = 0usize;
    let mut out = Vec::new();
    loop {
        let id = u16_at(data, pos)?;
        let n = u16_at(data, pos + 2)? as usize;
        if n > 8224 {
            return Err(bad("BIFF record exceeds maximum length"));
        }
        let payload = bytes(data, pos + 4, n)?;
        if matches!(id, 0x0009 | 0x0209 | 0x0409 | 0x0809) {
            depth += 1;
        }
        if out.is_empty() && depth == 0 {
            return Err(bad("substream does not start with BOF"));
        }
        out.push(Record {
            id,
            data: payload,
            offset: pos,
        });
        pos += 4 + n;
        if id == 0x0a {
            if depth == 0 {
                return Err(bad("unmatched EOF"));
            }
            depth -= 1;
            if depth == 0 {
                break;
            }
        }
        if out.len() > MAX_CELLS {
            return Err(bad("too many BIFF records"));
        }
    }
    Ok(out)
}
fn version(bof: Record<'_>) -> Result<u8> {
    match bof.id {
        9 => Ok(2),
        0x209 => Ok(3),
        0x409 => Ok(4),
        0x809 => match u16_at(bof.data, 0)? {
            0 => Ok(if u16_at(bof.data, 2)? == 0x1000 { 5 } else { 8 }),
            2 | 7 | 0x200 => Ok(2),
            0x300 => Ok(3),
            0x400 => Ok(4),
            0x500 => Ok(5),
            0x600 => Ok(8),
            _ => Err(bad("unknown BIFF version")),
        },
        _ => Err(bad("missing BIFF BOF")),
    }
}
fn continued<'a>(rs: &[Record<'a>], index: usize) -> Vec<&'a [u8]> {
    let mut parts = vec![rs[index].data];
    for r in &rs[index + 1..] {
        if r.id != 0x3c {
            break;
        }
        parts.push(r.data);
    }
    parts
}
fn joined_formula_records(rs: &[Record<'_>]) -> HashMap<usize, Vec<u8>> {
    let mut joined = HashMap::new();
    for (i, r) in rs.iter().enumerate() {
        if matches!(
            r.id,
            6 | 0x206 | 0x406 | 0x18 | 0x218 | 0x4bc | 0x221 | 0x21
        ) && rs.get(i + 1).is_some_and(|r| r.id == 0x3c)
        {
            let parts = continued(rs, i);
            let mut data = Vec::with_capacity(parts.iter().map(|p| p.len()).sum());
            for part in parts {
                data.extend_from_slice(part);
            }
            joined.insert(i, data);
        }
    }
    joined
}
impl NativeXlsBook {
    pub fn open_path(path: impl AsRef<Path>) -> Result<Self> {
        let mut f = File::open(path)?;
        if f.metadata()?.len() > MAX_BYTES as u64 {
            return Err(bad("source exceeds admission limit"));
        }
        let mut data = Vec::new();
        Read::by_ref(&mut f)
            .take(MAX_BYTES as u64 + 1)
            .read_to_end(&mut data)?;
        Self::open_bytes(data)
    }
    pub fn open_bytes(data: Vec<u8>) -> Result<Self> {
        if data.len() > MAX_BYTES {
            return Err(bad("source exceeds admission limit"));
        }
        let stream = if data.starts_with(&[0xd0, 0xcf, 0x11, 0xe0]) {
            cfb::workbook(&data)?
        } else {
            data
        };
        let rs = records(&stream, 0)?;
        let v = version(rs[0])?;
        let joined = joined_formula_records(&rs);
        let mut book = Self {
            stream: Vec::new(),
            sheets: Vec::new(),
            names: Vec::new(),
            formula_names: Vec::new(),
            strings: Vec::new(),
            formats: HashMap::new(),
            xfs: Vec::new(),
            font_pages: Vec::new(),
            xf_fonts: Vec::new(),
            deleted_labels: Vec::new(),
            version: v,
            page: if v == 8 { 1200 } else { 1252 },
            epoch1904: false,
            xti: Vec::new(),
            xti_books: Vec::new(),
            books: Vec::new(),
            old_extern: Vec::new(),
            extern_names: Vec::new(),
        };
        // CODEPAGE is global and can occur after strings in non-Excel writers.
        for r in &rs {
            if r.id == 0x42 {
                book.page = u16_at(r.data, 0)?;
            }
            if r.id == 0x2f {
                return Err(ReaderError::Unsupported(
                    "encrypted XLS workbooks require decryption before reading".into(),
                ));
            }
        }
        let mut raw_names = Vec::new();
        let mut old_format_id = 0u16;
        for (i, r) in rs.iter().enumerate() {
            match r.id {
                0x85 => {
                    let offset = u32_at(r.data, 0)? as usize;
                    let state = match *r.data.get(4).ok_or_else(|| bad("short BoundSheet"))? {
                        0 => SheetState::Visible,
                        1 => SheetState::Hidden,
                        2 => SheetState::VeryHidden,
                        _ => return Err(bad("invalid sheet state")),
                    };
                    let name = text::short(r.data, 6, v == 8, book.page)?;
                    if name.is_empty() || book.sheets.iter().any(|s| s.name == name) {
                        return Err(bad("empty/duplicate XLS sheet name"));
                    }
                    if offset >= stream.len() {
                        return Err(bad("sheet offset outside Workbook"));
                    }
                    book.sheets.push(Sheet {
                        name,
                        offset,
                        state,
                    });
                }
                0x8f if v == 4 => {
                    let name = text::short(r.data, 4, false, book.page)?;
                    book.sheets.push(Sheet {
                        name,
                        offset: r.offset + 4 + r.data.len(),
                        state: SheetState::Visible,
                    });
                }
                0x22 => {
                    book.epoch1904 = match u16_at(r.data, 0)? {
                        0 => false,
                        1 => true,
                        _ => return Err(bad("invalid DATEMODE")),
                    }
                }
                0xfc => {
                    if !book.strings.is_empty() {
                        return Err(bad("duplicate SST"));
                    }
                    let mut s = text::Segments::new(continued(&rs, i));
                    let total = s.dword()?;
                    let n = s.dword()? as usize;
                    if n > total as usize || n > MAX_CELLS || n > stream.len() / 3 {
                        return Err(bad("invalid SST counts"));
                    }
                    for _ in 0..n {
                        book.strings.push(Arc::from(s.unicode(None)?));
                    }
                }
                0x041e | 0x001e => {
                    let (id, offset) = if v >= 5 {
                        (u16_at(r.data, 0)?, 2)
                    } else {
                        let id = old_format_id;
                        old_format_id = old_format_id
                            .checked_add(1)
                            .ok_or_else(|| bad("too many formats"))?;
                        (id, if v == 4 { 2 } else { 0 })
                    };
                    let mut s = text::Segments::new(continued(&rs, i));
                    s.skip(offset)?;
                    let fmt = if v == 8 {
                        s.unicode(None)?
                    } else {
                        s.legacy_string(true, book.page)?
                    };
                    book.formats.insert(id, fmt);
                }
                0xe0 | 0x43 | 0x243 | 0x443 => {
                    let id = if v == 2 {
                        *r.data.get(2).ok_or_else(|| bad("truncated BIFF2 XF"))? as u16 & 0x3f
                    } else if v <= 4 {
                        *r.data.get(1).ok_or_else(|| bad("truncated legacy XF"))? as u16
                    } else {
                        u16_at(r.data, 2)?
                    };
                    book.xfs.push(id);
                    book.xf_fonts.push(if v <= 4 {
                        *r.data.first().ok_or_else(|| bad("short XF font"))? as u16
                    } else {
                        u16_at(r.data, 0)?
                    });
                }
                0x31 | 0x231 => {
                    if book.font_pages.len() == 4 {
                        book.font_pages.push(None);
                    }
                    let charset = if v >= 5 {
                        *r.data.get(12).ok_or_else(|| bad("short FONT charset"))?
                    } else {
                        1
                    };
                    book.font_pages.push(match charset {
                        0 => Some(1252),
                        1 | 2 => None,
                        77 => Some(10000),
                        128 => Some(932),
                        129 => Some(949),
                        130 => Some(1361),
                        134 => Some(936),
                        136 => Some(950),
                        161 => Some(1253),
                        162 => Some(1254),
                        163 => Some(1258),
                        177 => Some(1255),
                        178 => Some(1256),
                        186 => Some(1257),
                        204 => Some(1251),
                        222 => Some(874),
                        238 => Some(1250),
                        255 => Some(850),
                        _ => None,
                    });
                }
                0x18 | 0x218 => raw_names.push(i),
                0x1b9 if v == 8 => {
                    let mut s = text::Segments::new(continued(&rs, i));
                    let label = s.unicode(None)?;
                    if label.encode_utf16().count() >= 252 || book.deleted_labels.len() >= 2047 {
                        return Err(bad("invalid deleted natural-language label"));
                    }
                    book.deleted_labels.push(label);
                }
                0x1ae => {
                    let n = u16_at(r.data, 0)? as usize;
                    let tag = u16_at(r.data, 2)?;
                    let mut sheets = Vec::new();
                    if tag != 0x0401 && tag != 0x3a01 {
                        let mut s = text::Segments::new(vec![bytes(r.data, 2, r.data.len() - 2)?]);
                        let path = text::encoded_url(&s.unicode(None)?)?;
                        if n > 16384 {
                            return Err(bad("too many external sheets"));
                        }
                        for _ in 0..n {
                            let name = s.unicode(None)?;
                            sheets.push(format!("[{path}]{name}"));
                        }
                    }
                    book.books.push(sheets);
                    book.extern_names.push(Vec::new());
                }
                0x23 | 0x223 => {
                    let offset = if v >= 5 {
                        6
                    } else {
                        if v == 2 {
                            3
                        } else {
                            2
                        }
                    };
                    let name = text::short(r.data, offset, v == 8, book.page)?;
                    if let Some(ns) = book.extern_names.last_mut() {
                        ns.push(name);
                    } else {
                        book.extern_names.push(vec![name]);
                    }
                }
                0x17 => {
                    if v == 8 {
                        let n = u16_at(r.data, 0)? as usize;
                        if r.data.len() != 2 + n * 6 {
                            return Err(bad("invalid EXTERNSHEET size"));
                        }
                        for j in 0..n {
                            book.xti_books.push(u16_at(r.data, 2 + 6 * j)? as usize);
                            book.xti.push((
                                u16_at(r.data, 4 + 6 * j)? as i16,
                                u16_at(r.data, 6 + 6 * j)? as i16,
                            ));
                        }
                    } else {
                        book.old_extern
                            .push(text::legacy_external(r.data, book.page)?);
                        book.extern_names.push(Vec::new());
                    }
                }
                _ => {}
            }
        }
        if book.sheets.is_empty() {
            if u16_at(rs[0].data, 2)? == 0x10 || v < 5 {
                book.sheets.push(Sheet {
                    name: "Sheet1".into(),
                    offset: 0,
                    state: SheetState::Visible,
                });
            } else {
                return Err(bad("Workbook has no sheets"));
            }
        }
        let mut offsets = HashSet::new();
        for sheet in &book.sheets {
            if !offsets.insert(sheet.offset) {
                return Err(bad("duplicate sheet offset"));
            }
            let id = u16_at(&stream, sheet.offset)?;
            if !matches!(id, 9 | 0x209 | 0x409 | 0x809) {
                return Err(bad("BoundSheet does not point to BOF"));
            }
        }
        // Build name inventory first so forward/name-to-name tokens resolve.
        let mut pending = Vec::new();
        for index in raw_names {
            let raw = joined
                .get(&index)
                .map(Vec::as_slice)
                .unwrap_or(rs[index].data);
            let (name, tokens, extra, scope, hidden) = parse_name(raw, v, book.page)?;
            book.formula_names.push(name.clone());
            pending.push((name, tokens, extra, scope, hidden));
        }
        for (name, tokens, extra, scope, hidden) in pending {
            let local = scope.checked_sub(1).map(usize::from);
            if local.is_some_and(|n| n >= book.sheets.len()) {
                return Err(bad("invalid defined-name sheet scope"));
            }
            let refers_to = if tokens.is_empty() {
                "empty rgce".to_string()
            } else {
                rendered_formula(
                    formula::decode(tokens, extra, &book.context(), 0, 0, true),
                    None,
                )?
            };
            book.names.push(NamedRange {
                name,
                scope: if local.is_some() { "sheet" } else { "workbook" }.into(),
                local_sheet_id: local,
                refers_to,
                hidden,
                ..Default::default()
            });
        }
        book.stream = stream;
        Ok(book)
    }
    fn context(&self) -> formula::Context<'_> {
        formula::Context {
            version: self.version,
            page: self.page,
            sheets: &self.sheets,
            names: &self.formula_names,
            xti: &self.xti,
            old_extern: &self.old_extern,
            extern_names: &self.extern_names,
            xti_books: &self.xti_books,
            books: &self.books,
            deleted_labels: &self.deleted_labels,
            labels: None,
        }
    }
    pub fn sheet_names(&self) -> Vec<&str> {
        self.sheets.iter().map(|s| s.name.as_str()).collect()
    }
    pub fn sheet_state(&self, name: &str) -> Result<SheetState> {
        self.sheets
            .iter()
            .find(|s| s.name == name)
            .map(|s| s.state)
            .ok_or_else(|| ReaderError::SheetNotFound(name.into()))
    }
    pub fn named_ranges(&self) -> &[NamedRange] {
        &self.names
    }
    pub fn date1904(&self) -> bool {
        self.epoch1904
    }
    pub fn number_format_for_style_id(&self, id: u32) -> Option<&str> {
        let n = *self.xfs.get(id as usize)?;
        self.formats
            .get(&n)
            .map(String::as_str)
            .or_else(|| crate::builtin_num_fmt(n as u32))
    }
    fn cell_codepage(&self, style: u32) -> u16 {
        self.xf_fonts
            .get(style as usize)
            .and_then(|n| self.font_pages.get(*n as usize))
            .copied()
            .flatten()
            .unwrap_or(self.page)
    }
    pub fn worksheet(&self, name: &str) -> Result<WorksheetData> {
        let sheet = self
            .sheets
            .iter()
            .find(|s| s.name == name)
            .ok_or_else(|| ReaderError::SheetNotFound(name.into()))?;
        let rs = records(&self.stream, sheet.offset)?;
        let joined = joined_formula_records(&rs);
        let mut ctx = self.context();
        let mut local_extern = Vec::new();
        let mut local_names: Vec<Vec<String>> = Vec::new();
        if self.version < 8 {
            for r in &rs {
                if r.id == 0x17 {
                    local_extern.push(text::legacy_external(r.data, self.page)?);
                    local_names.push(Vec::new());
                }
                if matches!(r.id, 0x23 | 0x223) {
                    let offset = if self.version == 5 {
                        6
                    } else if self.version == 2 {
                        3
                    } else {
                        2
                    };
                    let name = text::short(r.data, offset, false, self.page)?;
                    local_names
                        .last_mut()
                        .ok_or_else(|| bad("external name without EXTERNSHEET"))?
                        .push(name);
                }
            }
            if !local_extern.is_empty() {
                ctx.old_extern = &local_extern;
                ctx.extern_names = &local_names;
            }
        }
        let mut cells = BTreeMap::new();
        let mut formulas = Vec::new();
        let mut shared: HashMap<(u32, u32), (&[u8], &[u8], bool)> = HashMap::new();
        let mut tables = HashMap::new();
        let mut pending_string = None;
        let mut result = WorksheetData::default();
        let mut depth = 0;
        let mut extended_xf = None;
        for (i, r) in rs.iter().enumerate() {
            if matches!(r.id, 9 | 0x209 | 0x409 | 0x809) {
                depth += 1;
                continue;
            }
            if r.id == 0xa {
                depth -= 1;
                continue;
            }
            if depth != 1 {
                continue;
            }
            if r.id == 0x2f {
                return Err(bad("encrypted worksheet"));
            }
            match r.id {
                0x44 if self.version == 2 => extended_xf = Some(u16_at(r.data, 0)? as u32),
                0x0002 | 0x0203 | 0x0003 | 0x027e | 0x007e | 0x0204 | 0x0004 | 0x00fd | 0x0205
                | 0x0005 | 0x0006 | 0x0206 | 0x0406 | 0x00d6 => {
                    let row = u16_at(r.data, 0)? as u32;
                    let col = u16_at(r.data, 2)? as u32;
                    check_coord(row, col, self.version)?;
                    let h = if self.version == 2 { 7 } else { 6 };
                    let style = if self.version == 2 {
                        let n = (r
                            .data
                            .get(4)
                            .copied()
                            .ok_or_else(|| bad("short BIFF2 attributes"))?
                            & 0x3f) as u32;
                        if n == 63 {
                            extended_xf
                                .take()
                                .ok_or_else(|| bad("BIFF2 cell missing IXFE"))?
                        } else {
                            n
                        }
                    } else {
                        u16_at(r.data, 4)? as u32
                    };
                    let mut cell = make_cell(row, col, style, CellValue::Empty);
                    match r.id {
                        2 => cell.value = CellValue::Number(u16_at(r.data, h)? as f64),
                        3 | 0x203 => {
                            cell.value = CellValue::Number(f64::from_le_bytes(
                                bytes(r.data, h, 8)?.try_into().unwrap(),
                            ))
                        }
                        0x27e | 0x7e => cell.value = CellValue::Number(rk(u32_at(r.data, h)?)),
                        4 | 0x204 | 0xd6 => {
                            let page = self.cell_codepage(style);
                            let mut s = text::Segments::new(continued(&rs, i));
                            s.skip(h)?;
                            let value = if self.version == 8 {
                                s.unicode(None)?
                            } else {
                                s.legacy_string(self.version == 2, page)?
                            };
                            cell.value = CellValue::String(value);
                        }
                        0xfd => {
                            let n = u32_at(r.data, h)? as usize;
                            cell.value = CellValue::SharedString(
                                self.strings
                                    .get(n)
                                    .ok_or_else(|| bad("LABELSST index outside SST"))?
                                    .clone(),
                            );
                            cell.data_type = CellDataType::SharedString;
                        }
                        5 | 0x205 => {
                            let b = bytes(r.data, h, 2)?;
                            cell.value = if b[1] == 0 {
                                if b[0] > 1 {
                                    return Err(bad("invalid BOOLERR Boolean"));
                                }
                                CellValue::Bool(b[0] == 1)
                            } else if b[1] == 1 {
                                CellValue::Error(error_text(b[0])?.into())
                            } else {
                                return Err(bad("invalid BOOLERR discriminator"));
                            };
                        }
                        6 | 0x206 | 0x406 => {
                            if pending_string.is_some() {
                                return Err(bad("new FORMULA before pending STRING cache"));
                            }
                            let cached = bytes(r.data, h, 8)?;
                            cell.value = if self.version > 2 && cached[6..] == [255, 255] {
                                match cached[0] {
                                    0 => {
                                        pending_string = Some((row, col));
                                        CellValue::Empty
                                    }
                                    1 => {
                                        if cached[2] > 1 {
                                            return Err(bad("invalid formula Boolean cache"));
                                        }
                                        CellValue::Bool(cached[2] == 1)
                                    }
                                    2 => CellValue::Error(error_text(cached[2])?.into()),
                                    3 => CellValue::String(String::new()),
                                    _ => return Err(bad("invalid formula cache type")),
                                }
                            } else {
                                CellValue::Number(f64::from_le_bytes(cached.try_into().unwrap()))
                            };
                            let offset = if self.version >= 5 { 20 } else { 16 };
                            let data = joined.get(&i).map(Vec::as_slice).unwrap_or(r.data);
                            let (tokens, extra) = parsed_formula(data, offset, self.version)?;
                            formulas.push(((row, col), tokens, extra));
                        }
                        _ => {}
                    }
                    cell.data_type = match cell.value {
                        CellValue::String(_) | CellValue::SharedString(_) => {
                            if cell.data_type == CellDataType::SharedString {
                                CellDataType::SharedString
                            } else {
                                CellDataType::InlineString
                            }
                        }
                        CellValue::Bool(_) => CellDataType::Bool,
                        CellValue::Error(_) => CellDataType::Error,
                        _ => CellDataType::Number,
                    };
                    insert_cell(&mut cells, (row, col), cell)?;
                }
                0x00bd => {
                    let row = u16_at(r.data, 0)? as u32;
                    let first = u16_at(r.data, 2)? as u32;
                    if r.data.len() < 12 || (r.data.len() - 6) % 6 != 0 {
                        return Err(bad("invalid MULRK length"));
                    }
                    let n = (r.data.len() - 6) / 6;
                    if u16_at(r.data, r.data.len() - 2)? as u32 != first + n as u32 - 1 {
                        return Err(bad("invalid MULRK column range"));
                    }
                    for j in 0..n {
                        check_coord(row, first + j as u32, self.version)?;
                        let style = u16_at(r.data, 4 + 6 * j)? as u32;
                        let value = CellValue::Number(rk(u32_at(r.data, 6 + 6 * j)?));
                        let c = make_cell(row, first + j as u32, style, value);
                        insert_cell(&mut cells, (row, first + j as u32), c)?;
                    }
                }
                0x207 | 0x7 => {
                    let key = pending_string
                        .take()
                        .ok_or_else(|| bad("STRING without pending formula cache"))?;
                    let style = cells.get(&key).and_then(|c: &Cell| c.style_id).unwrap_or(0);
                    let page = self.cell_codepage(style);
                    let mut s = text::Segments::new(continued(&rs, i));
                    let value = if self.version == 8 {
                        s.unicode(None)?
                    } else {
                        s.legacy_string(self.version == 2, page)?
                    };
                    let c = cells
                        .get_mut(&key)
                        .ok_or_else(|| bad("missing formula cache cell"))?;
                    c.value = CellValue::String(value);
                    c.data_type = CellDataType::FormulaString;
                }
                0x4bc | 0x221 | 0x21 => {
                    let row = u16_at(r.data, 0)? as u32;
                    let col = *r
                        .data
                        .get(4)
                        .ok_or_else(|| bad("short shared/array range"))?
                        as u32;
                    let offset = if r.id == 0x4bc {
                        8
                    } else if self.version >= 5 {
                        12
                    } else if self.version >= 3 {
                        8
                    } else {
                        7
                    };
                    let data = joined.get(&i).map(Vec::as_slice).unwrap_or(r.data);
                    let (tokens, extra) = parsed_formula(data, offset, self.version)?;
                    if shared
                        .insert((row, col), (tokens, extra, r.id == 0x4bc))
                        .is_some()
                    {
                        return Err(bad("duplicate shared/array formula anchor"));
                    }
                }
                0x36 | 0x236 | 0x37 => {
                    let row = u16_at(r.data, 0)? as u32;
                    let col = *r.data.get(4).ok_or_else(|| bad("short DATATABLE range"))? as u32;
                    tables.insert(
                        (row, col),
                        table_formula(r.data, self.version, r.id == 0x37)?,
                    );
                }
                0xe5 => {
                    let n = u16_at(r.data, 0)? as usize;
                    if r.data.len() != 2 + 8 * n {
                        return Err(bad("invalid MERGEDCELLS length"));
                    }
                    for j in 0..n {
                        let p = 2 + 8 * j;
                        let a = crate::row_col_to_a1(
                            u16_at(r.data, p)? as u32 + 1,
                            u16_at(r.data, p + 4)? as u32 + 1,
                        );
                        let b = crate::row_col_to_a1(
                            u16_at(r.data, p + 2)? as u32 + 1,
                            u16_at(r.data, p + 6)? as u32 + 1,
                        );
                        result.merged_ranges.push(format!("{a}:{b}"));
                    }
                }
                _ => {}
            }
        }
        if pending_string.is_some() {
            return Err(bad("missing formula STRING cache"));
        }
        for (key, tokens, extra) in formulas {
            let indirect = tokens.first() == Some(&1);
            if tokens.first() == Some(&2) {
                let row = u16_at(tokens, 1)? as u32;
                let col = if self.version == 2 {
                    *tokens.get(3).ok_or_else(|| bad("short PtgTbl"))? as u32
                } else {
                    u16_at(tokens, 3)? as u32
                };
                let text = tables
                    .get(&(row, col))
                    .ok_or_else(|| bad("missing DATATABLE definition"))?;
                let cell = cells
                    .get_mut(&key)
                    .ok_or_else(|| bad("missing data-table cell"))?;
                cell.formula = Some(text.clone());
                cell.formula_kind = Some("dataTable".into());
                continue;
            }
            let (tokens, extra, relative) = if indirect {
                let row = u16_at(tokens, 1)? as u32;
                let col = if self.version == 2 {
                    *tokens.get(3).ok_or_else(|| bad("short PtgExp"))? as u32
                } else {
                    u16_at(tokens, 3)? as u32
                };
                *shared
                    .get(&(row, col))
                    .ok_or_else(|| bad("missing shared/array formula definition"))?
            } else {
                (tokens, extra, false)
            };
            let formula = {
                let label_ctx = formula::Context {
                    labels: Some(&cells),
                    ..ctx
                };
                rendered_formula(
                    formula::decode(tokens, extra, &label_ctx, key.0, key.1, relative),
                    Some(key),
                )?
            };
            let cell = cells
                .get_mut(&key)
                .ok_or_else(|| bad("missing formula cell"))?;
            cell.formula = Some(formula);
            if indirect {
                cell.formula_kind = Some(if relative { "shared" } else { "array" }.into());
            }
        }
        if !cells.is_empty() {
            let min_row = cells.keys().map(|k| k.0).min().unwrap() + 1;
            let max_row = cells.keys().map(|k| k.0).max().unwrap() + 1;
            let min_col = cells.keys().map(|k| k.1).min().unwrap() + 1;
            let max_col = cells.keys().map(|k| k.1).max().unwrap() + 1;
            let a = crate::row_col_to_a1(min_row, min_col);
            let b = crate::row_col_to_a1(max_row, max_col);
            result.dimension = Some(format!("{a}:{b}"));
        }
        result.cells = cells.into_values().collect();
        Ok(result)
    }
}
fn check_coord(row: u32, col: u32, v: u8) -> Result<()> {
    if row >= if v == 8 { 65536 } else { 16384 } || col >= 256 {
        Err(bad("cell outside BIFF grid"))
    } else {
        Ok(())
    }
}
fn insert_cell(cells: &mut BTreeMap<(u32, u32), Cell>, key: (u32, u32), cell: Cell) -> Result<()> {
    if cells.len() >= MAX_CELLS {
        return Err(bad("too many XLS value cells"));
    }
    if cells.insert(key, cell).is_some() {
        return Err(bad("duplicate XLS cell record"));
    }
    Ok(())
}
fn make_cell(row: u32, col: u32, style: u32, value: CellValue) -> Cell {
    Cell {
        row: row + 1,
        col: col + 1,
        coordinate: crate::row_col_to_a1(row + 1, col + 1),
        style_id: Some(style),
        data_type: CellDataType::Number,
        value,
        formula: None,
        formula_kind: None,
        formula_shared_index: None,
        array_formula: None,
        rich_text: None,
    }
}
fn rk(n: u32) -> f64 {
    let v = if n & 2 != 0 {
        (n as i32 >> 2) as f64
    } else {
        f64::from_bits(((n & !3) as u64) << 32)
    };
    if n & 1 != 0 {
        v / 100.0
    } else {
        v
    }
}
fn parsed_formula(data: &[u8], offset: usize, v: u8) -> Result<(&[u8], &[u8])> {
    let (n, start) = if v == 2 {
        (
            *data.get(offset).ok_or_else(|| bad("short BIFF2 formula"))? as usize,
            offset + 1,
        )
    } else {
        (u16_at(data, offset)? as usize, offset + 2)
    };
    Ok((
        bytes(data, start, n)?,
        bytes(
            data,
            start + n,
            data.len()
                .checked_sub(start + n)
                .ok_or_else(|| bad("formula length overflow"))?,
        )?,
    ))
}
fn parse_name(data: &[u8], v: u8, page: u16) -> Result<(String, &[u8], &[u8], u16, bool)> {
    let flags = if v == 2 {
        *data.first().ok_or_else(|| bad("short NAME"))? as u16
    } else {
        u16_at(data, 0)?
    };
    let n = *data.get(3).ok_or_else(|| bad("short NAME"))? as usize;
    let sz = if v == 2 {
        *data.get(4).ok_or_else(|| bad("short NAME"))? as usize
    } else {
        u16_at(data, 4)? as usize
    };
    let start = if v >= 5 {
        14
    } else if v == 2 {
        5
    } else {
        6
    };
    let scope = if v >= 5 { u16_at(data, 8)? } else { 0 };
    let (name, end) = if v == 8 {
        let f = *data.get(start).ok_or_else(|| bad("short NAME text"))?;
        if f > 1 {
            return Err(bad("invalid NAME string flags"));
        }
        let d = bytes(data, start + 1, n * if f == 1 { 2 } else { 1 })?;
        let s = if f == 1 {
            text::utf16(d)?
        } else {
            d.iter().map(|&b| char::from(b)).collect()
        };
        (s, start + 1 + d.len())
    } else {
        (text::legacy(bytes(data, start, n)?, page)?, start + n)
    };
    let name = if flags & 0x20 != 0 {
        let id = name
            .chars()
            .next()
            .ok_or_else(|| bad("empty built-in name"))? as u32;
        match id {
            0 => "_xlnm.Consolidate_Area",
            1 => "_xlnm.Auto_Open",
            2 => "_xlnm.Auto_Close",
            3 => "_xlnm.Extract",
            4 => "_xlnm.Database",
            5 => "_xlnm.Criteria",
            6 => "_xlnm.Print_Area",
            7 => "_xlnm.Print_Titles",
            8 => "_xlnm.Recorder",
            9 => "_xlnm.Data_Form",
            10 => "_xlnm.Auto_Activate",
            11 => "_xlnm.Auto_Deactivate",
            12 => "_xlnm.Sheet_Title",
            13 => "_xlnm._FilterDatabase",
            _ => return Err(bad("invalid built-in name")),
        }
        .to_string()
    } else {
        name
    };
    let tokens = bytes(data, end, sz)?;
    Ok((
        name,
        tokens,
        bytes(data, end + sz, data.len() - (end + sz))?,
        scope,
        flags & 1 != 0,
    ))
}

fn table_formula(data: &[u8], v: u8, two: bool) -> Result<String> {
    let flags = if v == 2 {
        (*data.get(7).ok_or_else(|| bad("short DATATABLE flags"))? as u16) * 4
            | if two { 8 } else { 0 }
    } else {
        u16_at(data, 6)?
    };
    let input = |offset: usize, deleted: bool| -> Result<String> {
        if deleted {
            return Ok("#REF!".into());
        }
        let row = u16_at(data, offset)? as u32;
        let col = u16_at(data, offset + 2)? as u32;
        check_coord(row, col, v)?;
        let a = crate::row_col_to_a1(row + 1, col + 1);
        let split = a.find(|c: char| c.is_ascii_digit()).unwrap();
        Ok(format!("${}${}", &a[..split], row + 1))
    };
    let first = input(8, flags & 0x10 != 0)?;
    let (row, col) = if flags & 8 != 0 {
        (first, input(12, flags & 0x20 != 0)?)
    } else if flags & 4 != 0 {
        (first, String::new())
    } else {
        (String::new(), first)
    };
    Ok(format!("TABLE({row},{col})"))
}

// Keep the historical distinction between admitted cached data and optional
// formula rendering. Truncated token payloads remain structural read errors.
fn rendered_formula(result: Result<String>, position: Option<(u32, u32)>) -> Result<String> {
    match result {
        Ok(text) => Ok(text),
        Err(error @ ReaderError::Io(_)) => Err(error),
        Err(error) => Ok(match position {
            Some((row, col)) => format!("Unrecognised formula for cell ({row}, {col}): {error}"),
            None => format!("Unsupported ptg: {error}"),
        }),
    }
}
