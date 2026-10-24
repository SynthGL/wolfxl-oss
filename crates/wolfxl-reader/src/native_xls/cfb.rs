// SPDX-License-Identifier: MIT
//! Independently authored from MS-CFB §§2.2–2.6; no external parser code.
use super::{bad, bytes, u16_at, u32_at, Result, MAX_BYTES};
use std::collections::HashSet;

const FREE: u32 = 0xffff_ffff;
const END: u32 = 0xffff_fffe;
const FAT: u32 = 0xffff_fffd;
const DIF: u32 = 0xffff_fffc;

pub(super) fn workbook(data: &[u8]) -> Result<Vec<u8>> {
    if data.len() < 512 || data.get(..8) != Some(&[0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])
    {
        return Err(bad("invalid CFB signature/header"));
    }
    let version = u16_at(data, 26)?;
    let shift = u16_at(data, 30)?;
    if !matches!((version, shift), (3, 9) | (4, 12))
        || u16_at(data, 28)? != 0xfffe
        || u16_at(data, 32)? != 6
        || u32_at(data, 56)? != 4096
    {
        return Err(bad("invalid CFB version, byte order or sector geometry"));
    }
    let size = 1usize << shift;
    if data.len() % size != 0 || data.len() < size * 2 {
        return Err(bad("truncated CFB sectors"));
    }
    let count = data.len() / size - 1;
    let sector = |id: u32| -> Result<&[u8]> {
        if id as usize >= count {
            return Err(bad("CFB sector outside file"));
        }
        bytes(data, (id as usize + 1) * size, size)
    };
    let nf = u32_at(data, 44)? as usize;
    let nd = u32_at(data, 72)? as usize;
    if nf == 0 || nf > count || nd > count {
        return Err(bad("invalid CFB FAT count"));
    }
    let mut fats = Vec::with_capacity(nf);
    for p in (76..512).step_by(4) {
        let id = u32_at(data, p)?;
        if id != FREE {
            fats.push(id);
        }
    }
    let mut difats = HashSet::new();
    let mut next = u32_at(data, 68)?;
    for _ in 0..nd {
        if !difats.insert(next) {
            return Err(bad("cyclic CFB DIFAT"));
        }
        let s = sector(next)?;
        for p in (0..size - 4).step_by(4) {
            let id = u32_at(s, p)?;
            if id != FREE {
                fats.push(id);
            }
        }
        if fats.len() > nf {
            return Err(bad("too many CFB FAT sectors"));
        }
        next = u32_at(s, size - 4)?;
    }
    if (nd > 0 && next != END) || (nd == 0 && next != END && next != FREE) || fats.len() != nf {
        return Err(bad("CFB DIFAT count/termination mismatch"));
    }
    let mut reserved = difats.clone();
    let mut fat = Vec::with_capacity(nf * size / 4);
    for &id in &fats {
        if !reserved.insert(id) {
            return Err(bad("duplicate CFB allocation sector"));
        }
        for s in sector(id)?.chunks_exact(4) {
            fat.push(u32::from_le_bytes(s.try_into().unwrap()));
        }
    }
    if fat.len() < count {
        return Err(bad("CFB FAT does not cover file"));
    }
    for &id in &fats {
        if fat[id as usize] != FAT {
            return Err(bad("CFB FAT marker mismatch"));
        }
    }
    for &id in &difats {
        if fat[id as usize] != DIF {
            return Err(bad("CFB DIFAT marker mismatch"));
        }
    }
    let mut used = reserved;
    let mut read_chain = |first: u32, length: Option<usize>| -> Result<Vec<u8>> {
        let ids = chain(&fat, first, count, length.map(|n| n.div_ceil(size)))?;
        let n = length.unwrap_or(ids.len() * size);
        if n > MAX_BYTES {
            return Err(bad("CFB stream exceeds admission limit"));
        }
        let mut out = Vec::with_capacity(n);
        for id in ids {
            if !used.insert(id) {
                return Err(bad("overlapping CFB chains"));
            }
            let s = sector(id)?;
            out.extend_from_slice(&s[..(n - out.len()).min(size)]);
        }
        Ok(out)
    };
    let dir = read_chain(u32_at(data, 48)?, None)?;
    if version == 4 && dir.len() / size != u32_at(data, 40)? as usize {
        return Err(bad("CFB directory count mismatch"));
    }
    let root = bytes(&dir, 0, 128)?;
    if root[66] != 5 {
        return Err(bad("missing CFB root storage"));
    }
    let stream_len = |entry: &[u8]| -> Result<usize> {
        let lo = u32_at(entry, 120)? as u64;
        let hi = if version == 3 {
            0
        } else {
            u32_at(entry, 124)? as u64
        };
        let n = lo | (hi << 32);
        if n > MAX_BYTES as u64 {
            return Err(bad("CFB stream too large"));
        }
        Ok(n as usize)
    };
    let mini_len = stream_len(root)?;
    let mini = read_chain(u32_at(root, 116)?, Some(mini_len))?;
    let nm = u32_at(data, 64)? as usize;
    if nm > count {
        return Err(bad("invalid CFB mini FAT count"));
    }
    let mini_table = read_chain(u32_at(data, 60)?, Some(nm * size))?;
    let mini_fat: Vec<u32> = mini_table
        .chunks_exact(4)
        .map(|s| u32::from_le_bytes(s.try_into().unwrap()))
        .collect();
    // Only the root's sibling tree contains root-level streams; nested VBA or
    // embedded workbooks are not eligible Workbook streams.
    let mut todo = vec![u32_at(root, 76)?];
    let mut seen = HashSet::new();
    let mut found: Option<(u32, usize, bool)> = None;
    while let Some(id) = todo.pop() {
        if id == FREE {
            continue;
        }
        if !seen.insert(id) {
            return Err(bad("cyclic CFB directory tree"));
        }
        let e = bytes(&dir, id as usize * 128, 128)?;
        if !matches!(e[66], 1 | 2) || e[67] > 1 {
            return Err(bad("invalid CFB directory entry"));
        }
        todo.push(u32_at(e, 68)?);
        todo.push(u32_at(e, 72)?);
        let n = u16_at(e, 64)? as usize;
        if !(2..=64).contains(&n) || n % 2 != 0 || u16_at(e, n - 2)? != 0 {
            return Err(bad("invalid CFB directory name"));
        }
        let units: Vec<u16> = e[..n - 2]
            .chunks_exact(2)
            .map(|s| u16::from_le_bytes(s.try_into().unwrap()))
            .collect();
        let name = String::from_utf16(&units).map_err(|_| bad("invalid CFB UTF-16 name"))?;
        if e[66] == 2
            && (name.eq_ignore_ascii_case("Workbook") || name.eq_ignore_ascii_case("Book"))
        {
            let modern = name.eq_ignore_ascii_case("Workbook");
            if found.is_some_and(|(_, _, m)| m == modern) {
                return Err(bad("duplicate CFB Workbook stream"));
            }
            // DSF workbooks can contain both the BIFF8 Workbook and an older
            // Book stream. Select the modern stream deterministically.
            if found.is_none() || modern {
                found = Some((u32_at(e, 116)?, stream_len(e)?, modern));
            }
        }
    }
    let (start, n, _) = found.ok_or_else(|| bad("CFB has no Workbook or Book stream"))?;
    if n >= 4096 {
        return read_chain(start, Some(n));
    }
    let ids = chain(
        &mini_fat,
        start,
        mini.len().div_ceil(64),
        Some(n.div_ceil(64)),
    )?;
    let mut out = Vec::with_capacity(n);
    for id in ids {
        let s = bytes(&mini, id as usize * 64, (n - out.len()).min(64))?;
        out.extend_from_slice(s);
    }
    Ok(out)
}

fn chain(table: &[u32], mut next: u32, limit: usize, expected: Option<usize>) -> Result<Vec<u32>> {
    let mut out = Vec::new();
    let mut seen = HashSet::new();
    while next != END {
        if next as usize >= limit || !seen.insert(next) {
            return Err(bad("invalid/cyclic CFB allocation chain"));
        }
        if out.len() >= expected.unwrap_or(limit) {
            return Err(bad("overlong CFB allocation chain"));
        }
        out.push(next);
        next = *table
            .get(next as usize)
            .ok_or_else(|| bad("CFB chain outside FAT"))?;
    }
    if expected.is_some_and(|n| n != out.len()) {
        return Err(bad("truncated CFB allocation chain"));
    }
    Ok(out)
}
