// SPDX-License-Identifier: MIT
//! MS-XLS §2.5.293 XLUnicodeRichExtendedString and Continue records;
//! legacy byte strings per OpenOffice Excel format §2.5.2 (format facts only).
use super::{bad, bytes, Result};

pub(super) fn legacy(data: &[u8], page: u16) -> Result<String> {
    let enc = codepage::to_encoding(page)
        .ok_or_else(|| bad(&format!("unrecognized XLS codepage {page}")))?;
    let (s, _, _) = enc.decode(data);
    Ok(s.into_owned())
}
pub(super) fn utf16(data: &[u8]) -> Result<String> {
    if data.len() % 2 != 0 {
        return Err(bad("odd UTF-16 byte count"));
    }
    let chars: Vec<u16> = data
        .chunks_exact(2)
        .map(|s| u16::from_le_bytes(s.try_into().unwrap()))
        .collect();
    String::from_utf16(&chars).map_err(|_| bad("invalid XLS UTF-16"))
}

// Segment boundaries matter: only continuations *inside character data* carry
// a new compression flag. Header/rich-run/phonetic continuations do not.
pub(super) struct Segments<'a> {
    parts: Vec<&'a [u8]>,
    part: usize,
    pos: usize,
}
impl<'a> Segments<'a> {
    pub(super) fn new(parts: Vec<&'a [u8]>) -> Self {
        Self {
            parts,
            part: 0,
            pos: 0,
        }
    }
    fn boundary(&mut self) -> Result<bool> {
        let mut crossed = false;
        loop {
            let part = self
                .parts
                .get(self.part)
                .ok_or_else(|| bad("truncated continued XLS string"))?;
            if self.pos < part.len() {
                return Ok(crossed);
            }
            self.part += 1;
            self.pos = 0;
            crossed = true;
        }
    }
    pub(super) fn byte(&mut self) -> Result<u8> {
        self.boundary()?;
        let v = self.parts[self.part][self.pos];
        self.pos += 1;
        Ok(v)
    }
    pub(super) fn word(&mut self) -> Result<u16> {
        Ok(u16::from_le_bytes([self.byte()?, self.byte()?]))
    }
    pub(super) fn dword(&mut self) -> Result<u32> {
        Ok(u32::from_le_bytes([
            self.byte()?,
            self.byte()?,
            self.byte()?,
            self.byte()?,
        ]))
    }
    pub(super) fn skip(&mut self, n: usize) -> Result<()> {
        for _ in 0..n {
            self.byte()?;
        }
        Ok(())
    }
    pub(super) fn legacy_string(&mut self, short: bool, page: u16) -> Result<String> {
        let n = if short {
            self.byte()? as usize
        } else {
            self.word()? as usize
        };
        let mut data = Vec::with_capacity(n);
        for _ in 0..n {
            data.push(self.byte()?);
        }
        legacy(&data, page)
    }
    pub(super) fn unicode(&mut self, count: Option<usize>) -> Result<String> {
        let count = match count {
            Some(n) => n,
            None => self.word()? as usize,
        };
        let flags = self.byte()?;
        if flags & !0x0d != 0 {
            return Err(bad("invalid XLS Unicode string flags"));
        }
        let rich = if flags & 8 != 0 {
            self.word()? as usize
        } else {
            0
        };
        let ext = if flags & 4 != 0 {
            self.dword()? as usize
        } else {
            0
        };
        let mut wide = flags & 1 != 0;
        let mut chars = Vec::with_capacity(count);
        for _ in 0..count {
            if self.boundary()? {
                let f = self.byte()?;
                if f > 1 {
                    return Err(bad("invalid Continue compression flag"));
                }
                wide = f == 1;
            }
            let lo = self.byte()?;
            let ch = if wide {
                // A producer cannot split a UTF-16 code unit across records.
                if self.pos == self.parts[self.part].len() {
                    return Err(bad("split XLS UTF-16 code unit"));
                }
                u16::from_le_bytes([lo, self.byte()?])
            } else {
                lo as u16
            };
            chars.push(ch);
        }
        self.skip(
            rich.checked_mul(4)
                .and_then(|n| n.checked_add(ext))
                .ok_or_else(|| bad("XLS rich string length overflow"))?,
        )?;
        String::from_utf16(&chars).map_err(|_| bad("invalid continued XLS UTF-16"))
    }
}

pub(super) fn short(data: &[u8], offset: usize, wide: bool, page: u16) -> Result<String> {
    let count = *data
        .get(offset)
        .ok_or_else(|| bad("truncated short XLS string"))? as usize;
    if wide {
        Segments::new(vec![bytes(data, offset + 1, data.len() - offset - 1)?]).unicode(Some(count))
    } else {
        legacy(bytes(data, offset + 1, count)?, page)
    }
}

pub(super) fn legacy_external(data: &[u8], page: u16) -> Result<String> {
    let n = *data.first().ok_or_else(|| bad("short EXTERNSHEET"))? as usize;
    let internal = data.get(1) == Some(&3);
    let length = n + usize::from(internal);
    let raw = legacy(bytes(data, 1, length)?, page)?;
    encoded_url(&raw)
}

pub(super) fn encoded_url(raw: &str) -> Result<String> {
    let mut chars = raw.chars();
    match chars.next() {
        Some('\u{0}' | '\u{2}' | '\u{3}' | '\u{4}') => return Ok(chars.collect()),
        Some('\u{1}') => {}
        _ => return Ok(raw.to_string()),
    }
    let mut out = String::new();
    while let Some(c) = chars.next() {
        match c {
            '\u{1}' => {
                let drive = chars
                    .next()
                    .ok_or_else(|| bad("truncated external URL drive"))?;
                if drive == '@' {
                    out.push_str("\\\\");
                } else {
                    out.push(drive);
                    out.push_str(":\\");
                }
            }
            '\u{2}' | '\u{3}' => out.push('\\'),
            '\u{4}' => out.push_str("..\\"),
            '\u{5}' => {
                let n = chars
                    .next()
                    .ok_or_else(|| bad("truncated external URL length"))?
                    as usize;
                for _ in 0..n {
                    out.push(chars.next().ok_or_else(|| bad("truncated external URL"))?);
                }
            }
            // Preserve the specification's symbolic installation-root marker.
            // Resolving it to a host path would invent workbook information.
            '\u{6}' | '\u{8}' => out.push(c),
            '\u{9}' => {}
            c => out.push(c),
        }
    }
    Ok(out)
}
