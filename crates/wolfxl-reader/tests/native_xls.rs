// SPDX-License-Identifier: MIT
//! Independent, hand-constructed format fixtures. Expected values/formulas are
//! literal constants, never output from a spreadsheet parser. Record layouts:
//! MS-XLS §§2.4/2.5.198; MS-CFB §§2.2–2.6; legacy public Excel format rev1.42.
use wolfxl_reader::{CellValue, NativeXlsBook, ReaderError, SheetState};

fn word(out: &mut Vec<u8>, v: u16) {
    out.extend_from_slice(&v.to_le_bytes());
}
fn dword(out: &mut Vec<u8>, v: u32) {
    out.extend_from_slice(&v.to_le_bytes());
}
fn rec(out: &mut Vec<u8>, id: u16, p: &[u8]) {
    word(out, id);
    word(out, p.len() as u16);
    out.extend_from_slice(p);
}
fn bof(out: &mut Vec<u8>, v: u8, kind: u16) {
    let id = match v {
        2 => 9,
        3 => 0x209,
        4 => 0x409,
        _ => 0x809,
    };
    let mut p = Vec::new();
    word(
        &mut p,
        if v == 8 {
            0x600
        } else if v == 5 {
            0x500
        } else {
            0
        },
    );
    word(&mut p, kind);
    if v >= 5 {
        word(&mut p, 0x0dbb);
        word(&mut p, 1997);
    }
    if v == 8 {
        dword(&mut p, 0);
        dword(&mut p, 6);
    }
    rec(out, id, &p);
}
fn eof(out: &mut Vec<u8>) {
    rec(out, 0xa, &[]);
}
fn header(row: u16, col: u16, style: u16, v: u8) -> Vec<u8> {
    let mut p = Vec::new();
    word(&mut p, row);
    word(&mut p, col);
    word(&mut p, style);
    if v == 2 {
        p.push(0);
    }
    p
}
fn number(out: &mut Vec<u8>, r: u16, c: u16, v: u8, n: f64) {
    let mut p = header(r, c, 0, v);
    p.extend_from_slice(&n.to_le_bytes());
    rec(out, if v == 2 { 3 } else { 0x203 }, &p);
}
fn label(out: &mut Vec<u8>, r: u16, c: u16, v: u8, s: &[u8]) {
    let mut p = header(r, c, 0, v);
    if v == 2 {
        p.push(s.len() as u8);
    } else {
        word(&mut p, s.len() as u16);
    }
    if v == 8 {
        p.push(0);
    }
    p.extend_from_slice(s);
    rec(out, if v == 2 { 4 } else { 0x204 }, &p);
}
fn formula(out: &mut Vec<u8>, r: u16, c: u16, v: u8, cache: [u8; 8], tokens: &[u8]) {
    let mut p = header(r, c, 0, v);
    p.extend_from_slice(&cache);
    if v == 2 {
        p.push(0);
    } else {
        word(&mut p, 0);
    }
    if v >= 5 {
        dword(&mut p, 0);
    }
    if v == 2 {
        p.push(tokens.len() as u8);
    } else {
        word(&mut p, tokens.len() as u16);
    }
    p.extend_from_slice(tokens);
    rec(
        out,
        match v {
            3 => 0x206,
            4 => 0x406,
            _ => 6,
        },
        &p,
    );
}
fn boundsheet(out: &mut Vec<u8>, name: &[u8], state: u8, v: u8) -> usize {
    let patch = out.len() + 4;
    let mut p = Vec::new();
    dword(&mut p, 0);
    p.extend_from_slice(&[state, 0, name.len() as u8]);
    if v == 8 {
        p.push(0);
    }
    p.extend_from_slice(name);
    rec(out, 0x85, &p);
    patch
}
fn set32(data: &mut [u8], p: usize, n: u32) {
    data[p..p + 4].copy_from_slice(&n.to_le_bytes());
}
fn set16(data: &mut [u8], p: usize, n: u16) {
    data[p..p + 2].copy_from_slice(&n.to_le_bytes());
}
fn workbook(
    v: u8,
    globals: impl FnOnce(&mut Vec<u8>),
    sheet: impl FnOnce(&mut Vec<u8>),
) -> Vec<u8> {
    let mut b = Vec::new();
    bof(&mut b, v, 5);
    globals(&mut b);
    let patch = boundsheet(&mut b, b"Data", 0, v);
    eof(&mut b);
    let pos = b.len();
    set32(&mut b, patch, pos as u32);
    bof(&mut b, v, 0x10);
    sheet(&mut b);
    eof(&mut b);
    b
}
fn entry(dir: &mut [u8], offset: usize, name: &str, kind: u8, start: u32, size: u64, child: u32) {
    let e = &mut dir[offset..offset + 128];
    let chars: Vec<u16> = name.encode_utf16().chain(Some(0)).collect();
    for (i, ch) in chars.iter().enumerate() {
        set16(e, 2 * i, *ch);
    }
    set16(e, 64, (chars.len() * 2) as u16);
    e[66] = kind;
    e[67] = 1;
    set32(e, 68, u32::MAX);
    set32(e, 72, u32::MAX);
    set32(e, 76, child);
    set32(e, 116, start);
    e[120..128].copy_from_slice(&size.to_le_bytes());
}
// Compact producer for CFB v3 and v4, regular or mini streams. FAT placement
// is independent of production allocation traversal, including a reversed
// mini-sector chain to exercise noncontiguous logical stream order.
fn compound(stream: &[u8], mini: bool, version: u16, name: &str) -> Vec<u8> {
    let sector = if version == 3 { 512 } else { 4096 };
    let logical = if mini {
        stream.len()
    } else {
        stream.len().max(4096)
    };
    let slots = if mini {
        logical.div_ceil(64)
    } else {
        logical.div_ceil(sector)
    };
    let data_sectors = if mini {
        (slots * 64).div_ceil(sector)
    } else {
        slots
    };
    assert!(data_sectors + 3 < sector / 4);
    let mini_id = 2 + data_sectors;
    let count = 2 + data_sectors + usize::from(mini);
    let mut b = vec![0u8; (count + 1) * sector];
    b[..8].copy_from_slice(&[0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
    set16(&mut b, 24, 0x3e);
    set16(&mut b, 26, version);
    set16(&mut b, 28, 0xfffe);
    set16(&mut b, 30, if version == 3 { 9 } else { 12 });
    set16(&mut b, 32, 6);
    set32(&mut b, 40, u32::from(version == 4));
    set32(&mut b, 44, 1);
    set32(&mut b, 48, 1);
    set32(&mut b, 56, 4096);
    set32(&mut b, 60, if mini { mini_id as u32 } else { 0xffff_fffe });
    set32(&mut b, 64, u32::from(mini));
    set32(&mut b, 68, 0xffff_fffe);
    for p in (76..512).step_by(4) {
        set32(&mut b, p, u32::MAX);
    }
    set32(&mut b, 76, 0);
    let fat = &mut b[sector..2 * sector];
    fat.fill(255);
    set32(fat, 0, 0xffff_fffd);
    set32(fat, 4, 0xffff_fffe);
    for j in 0..data_sectors {
        set32(
            fat,
            (j + 2) * 4,
            if j + 1 == data_sectors {
                0xffff_fffe
            } else {
                j as u32 + 3
            },
        );
    }
    if mini {
        set32(fat, mini_id * 4, 0xffff_fffe);
    }
    let dir = &mut b[2 * sector..3 * sector];
    entry(
        dir,
        0,
        "Root Entry",
        5,
        if mini { 2 } else { 0xffff_fffe },
        if mini { (slots * 64) as u64 } else { 0 },
        1,
    );
    entry(
        dir,
        128,
        name,
        2,
        if mini { (slots - 1) as u32 } else { 2 },
        logical as u64,
        u32::MAX,
    );
    if mini {
        for (j, chunk) in stream.chunks(64).enumerate() {
            let p = 3 * sector + (slots - 1 - j) * 64;
            b[p..p + chunk.len()].copy_from_slice(chunk);
        }
        let table = &mut b[(mini_id + 1) * sector..(mini_id + 2) * sector];
        table.fill(255);
        for j in 0..slots {
            set32(
                table,
                j * 4,
                if j == 0 { 0xffff_fffe } else { j as u32 - 1 },
            );
        }
    } else {
        b[3 * sector..3 * sector + stream.len()].copy_from_slice(stream);
    }
    b
}

#[test]
fn legacy_versions_keep_numbers_strings_formula_and_sparse_origin() {
    for version in [2, 3, 4, 5, 8] {
        let mut raw = Vec::new();
        bof(&mut raw, version, 0x10);
        rec(&mut raw, 0x42, &1252u16.to_le_bytes());
        label(&mut raw, 4, 2, version, b"caf\xe9");
        number(&mut raw, 4, 3, version, -12.5);
        formula(
            &mut raw,
            5,
            2,
            version,
            3f64.to_le_bytes(),
            &[0x1e, 1, 0, 0x1e, 2, 0, 3],
        );
        eof(&mut raw);
        for source in [
            raw.clone(),
            compound(&raw, true, 3, "Book"),
            compound(&raw, false, 3, "WORKBOOK"),
        ] {
            let book = NativeXlsBook::open_bytes(source).unwrap();
            assert_eq!(book.sheet_names(), ["Sheet1"]);
            let s = book.worksheet("Sheet1").unwrap();
            assert_eq!(s.dimension.as_deref(), Some("C5:D6"));
            assert_eq!(s.cells.len(), 3);
            assert_eq!(s.cells[0].value, CellValue::String("café".into()));
            assert_eq!(s.cells[1].value, CellValue::Number(-12.5));
            assert_eq!(s.cells[2].formula.as_deref(), Some("1+2"));
            assert_eq!(s.cells[2].value, CellValue::Number(3.0));
        }
    }
}
#[test]
fn cfb_mini_regular_and_v4_preserve_sheet_inventory() {
    let mut raw = Vec::new();
    bof(&mut raw, 8, 5);
    let a = boundsheet(&mut raw, b"First", 2, 8);
    let b = boundsheet(&mut raw, b"Other", 1, 8);
    eof(&mut raw);
    let p = raw.len();
    set32(&mut raw, a, p as u32);
    bof(&mut raw, 8, 0x10);
    number(&mut raw, 2, 1, 8, 42.5);
    eof(&mut raw);
    let p = raw.len();
    set32(&mut raw, b, p as u32);
    bof(&mut raw, 8, 0x10);
    label(&mut raw, 0, 0, 8, b"last");
    eof(&mut raw);
    for (mini, v, name) in [
        (true, 3, "Book"),
        (false, 3, "Workbook"),
        (true, 4, "WORKBOOK"),
        (false, 4, "BOOK"),
    ] {
        let book = NativeXlsBook::open_bytes(compound(&raw, mini, v, name)).unwrap();
        assert_eq!(book.sheet_names(), ["First", "Other"]);
        assert_eq!(book.sheet_state("First").unwrap(), SheetState::VeryHidden);
        assert_eq!(book.sheet_state("Other").unwrap(), SheetState::Hidden);
        assert!(matches!(
            book.worksheet("missing"),
            Err(ReaderError::SheetNotFound(_))
        ));
        let s = book.worksheet("First").unwrap();
        assert_eq!(s.dimension.as_deref(), Some("B3:B3"));
        assert_eq!(s.cells[0].value, CellValue::Number(42.5));
    }
}
#[test]
fn legacy_codepages_and_cjk_multibyte_strings() {
    for (page, data, expected) in [
        (1251u16, b"\xcf\xf0\xe8\xe2\xe5\xf2".as_slice(), "Привет"),
        (932, b"\x93\xfa\x96\x7b".as_slice(), "日本"),
        (1252, b"\x80\x93\x94".as_slice(), "€“”"),
    ] {
        let raw = workbook(
            5,
            |g| rec(g, 0x42, &page.to_le_bytes()),
            |s| label(s, 9, 7, 5, data),
        );
        let b = NativeXlsBook::open_bytes(compound(&raw, true, 3, "Book")).unwrap();
        let s = b.worksheet("Data").unwrap();
        assert_eq!(s.cells[0].value, CellValue::String(expected.into()));
        assert_eq!(s.dimension.as_deref(), Some("H10:H10"));
    }
}
#[test]
fn sst_continuations_change_compression_and_skip_rich_phonetic_payloads() {
    let raw = workbook(
        8,
        |g| {
            let mut a = Vec::new();
            dword(&mut a, 2);
            dword(&mut a, 2);
            word(&mut a, 4);
            a.push(0x0c);
            word(&mut a, 1);
            dword(&mut a, 3);
            a.extend_from_slice(b"ab");
            rec(g, 0xfc, &a);
            // Continuation in text has a new compression byte; rich-run continuation
            // below has none. First string is abΩ中, second is Latin-1 é.
            rec(g, 0x3c, &[1, 0xa9, 3, 0x2d, 0x4e, 0, 0]);
            rec(g, 0x3c, &[0, 0, 9, 8, 7, 1, 0, 0, 0xe9]);
        },
        |s| {
            let mut p = header(2, 1, 0, 8);
            dword(&mut p, 0);
            rec(s, 0xfd, &p);
            let mut p = header(2, 2, 0, 8);
            dword(&mut p, 1);
            rec(s, 0xfd, &p);
        },
    );
    let b = NativeXlsBook::open_bytes(raw).unwrap();
    let s = b.worksheet("Data").unwrap();
    assert_eq!(s.cells[0].value, CellValue::String("abΩ中".into()));
    assert_eq!(s.cells[1].value, CellValue::String("é".into()));
    assert_eq!(s.dimension.as_deref(), Some("B3:C3"));
}
#[test]
fn typed_errors_formula_caches_and_precedence_are_separate() {
    let raw = workbook(
        8,
        |_| {},
        |s| {
            let mut p = header(6, 3, 0, 8);
            p.extend_from_slice(&[7, 1]);
            rec(s, 0x205, &p);
            formula(s, 6, 4, 8, [1, 0, 1, 0, 0, 0, 255, 255], &[0x1d, 1]);
            formula(s, 6, 5, 8, [2, 0, 42, 0, 0, 0, 255, 255], &[0x1c, 42]);
            formula(
                s,
                6,
                6,
                8,
                [0, 0, 0, 0, 0, 0, 255, 255],
                &[0x17, 3, 0, b'a', b'b', b'c'],
            );
            rec(s, 0x207, &[3, 0, 0, b'x', b'y', b'z']);
            formula(
                s,
                6,
                7,
                8,
                9f64.to_le_bytes(),
                &[0x1e, 1, 0, 0x1e, 2, 0, 3, 0x1e, 3, 0, 5],
            );
            formula(s, 6, 8, 8, [3, 0, 0, 0, 0, 0, 255, 255], &[0x17, 0, 0]);
        },
    );
    let b = NativeXlsBook::open_bytes(raw).unwrap();
    let s = b.worksheet("Data").unwrap();
    assert_eq!(s.cells[0].value, CellValue::Error("#DIV/0!".into()));
    assert_eq!(s.cells[1].value, CellValue::Bool(true));
    assert_eq!(s.cells[1].formula.as_deref(), Some("TRUE"));
    assert_eq!(s.cells[2].value, CellValue::Error("#N/A".into()));
    assert_eq!(s.cells[2].formula.as_deref(), Some("#N/A"));
    assert_eq!(s.cells[3].value, CellValue::String("xyz".into()));
    assert_eq!(s.cells[3].formula.as_deref(), Some("\"abc\""));
    assert_eq!(s.cells[4].formula.as_deref(), Some("(1+2)*3"));
    assert_eq!(s.cells[4].value, CellValue::Number(9.0));
    assert_eq!(s.cells[5].formula.as_deref(), Some("\"\""));
    assert_eq!(s.cells[5].value, CellValue::String(String::new()));
}
fn name(g: &mut Vec<u8>, v: u8, text: &str, scope: u16, tokens: &[u8]) {
    let mut p = vec![0, 0, 0, text.len() as u8];
    word(&mut p, tokens.len() as u16);
    word(&mut p, 0);
    word(&mut p, scope);
    p.extend_from_slice(&[0; 4]);
    if v == 8 {
        p.push(0);
    }
    p.extend_from_slice(text.as_bytes());
    p.extend_from_slice(tokens);
    rec(g, 0x18, &p);
}
#[test]
fn biff8_defined_ranges_scopes_names_and_formula_functions() {
    let raw = workbook(
        8,
        |g| {
            rec(g, 0x1ae, &[1, 0, 1, 4]);
            rec(g, 0x17, &[1, 0, 0, 0, 0, 0, 0, 0]);
            name(g, 8, "Region", 0, &[0x3b, 0, 0, 2, 0, 4, 0, 1, 0, 3, 0]);
            name(g, 8, "Local", 1, &[0x24, 2, 0, 1, 0]);
            name(g, 8, "Forward", 0, &[0x23, 4, 0, 0, 0]);
            name(g, 8, "Constant", 0, &[0x1e, 17, 0]);
        },
        |s| {
            formula(
                s,
                2,
                4,
                8,
                42f64.to_le_bytes(),
                &[0x23, 1, 0, 0, 0, 0x42, 1, 4, 0],
            );
            formula(
                s,
                3,
                4,
                8,
                3f64.to_le_bytes(),
                &[0x1e, 1, 0, 0x1e, 2, 0, 0x41, 27, 0],
            );
        },
    );
    let b = NativeXlsBook::open_bytes(raw).unwrap();
    let ns = b.named_ranges();
    assert_eq!(ns.len(), 4);
    assert_eq!(ns[0].name, "Region");
    assert_eq!(ns[0].refers_to, "Data!$B$3:$D$5");
    assert_eq!(ns[0].scope, "workbook");
    assert_eq!(ns[1].local_sheet_id, Some(0));
    assert_eq!(ns[1].scope, "sheet");
    assert_eq!(ns[1].refers_to, "$B$3");
    assert_eq!(ns[2].refers_to, "Constant");
    assert_eq!(ns[3].refers_to, "17");
    let s = b.worksheet("Data").unwrap();
    assert_eq!(s.cells[0].formula.as_deref(), Some("SUM(Region)"));
    assert_eq!(s.cells[1].formula.as_deref(), Some("ROUND(1,2)"));
}
#[test]
fn biff5_three_dimensional_names_and_number_formats_epochs() {
    for epoch in [0u16, 1] {
        let raw = workbook(
            5,
            |g| {
                rec(g, 0x22, &epoch.to_le_bytes());
                rec(g, 0xe0, &[0, 0, 14, 0]);
                let mut custom = Vec::new();
                word(&mut custom, 164);
                custom.push(10);
                custom.extend_from_slice(b"[h]:mm:ss;");
                rec(g, 0x41e, &custom);
                rec(g, 0xe0, &[0, 0, 164, 0]);
                let mut tokens = vec![0x3b, 255, 255];
                tokens.extend_from_slice(&[0; 8]);
                tokens.extend_from_slice(&[0, 0, 0, 0, 2, 0, 4, 0, 1, 3]);
                name(g, 5, "OldRange", 1, &tokens);
            },
            |s| {
                number(s, 2, 1, 5, 1462.0);
                let mut p = header(2, 2, 1, 5);
                p.extend_from_slice(&1.5f64.to_le_bytes());
                rec(s, 0x203, &p);
            },
        );
        let b = NativeXlsBook::open_bytes(compound(&raw, true, 3, "Book")).unwrap();
        assert_eq!(b.date1904(), epoch == 1);
        assert_eq!(b.number_format_for_style_id(0), Some("mm-dd-yy"));
        assert_eq!(b.number_format_for_style_id(1), Some("[h]:mm:ss;"));
        assert_eq!(b.named_ranges()[0].refers_to, "Data!$B$3:$D$5");
        assert_eq!(b.named_ranges()[0].local_sheet_id, Some(0));
        let s = b.worksheet("Data").unwrap();
        assert_eq!(s.cells[0].value, CellValue::Number(1462.0));
        assert_eq!(s.cells[1].style_id, Some(1));
    }
}
#[test]
fn shared_formulas_use_signed_relative_offsets_not_master_text() {
    let raw = workbook(
        8,
        |_| {},
        |s| {
            formula(s, 2, 2, 8, 8f64.to_le_bytes(), &[1, 2, 0, 2, 0]);
            let mut p = vec![2, 0, 3, 0, 2, 2, 0, 2];
            word(&mut p, 9);
            p.extend_from_slice(&[0x2c, 0, 0, 255, 0xc0, 0x1e, 2, 0, 5]);
            rec(s, 0x4bc, &p);
            formula(s, 3, 2, 8, 10f64.to_le_bytes(), &[1, 2, 0, 2, 0]);
        },
    );
    let b = NativeXlsBook::open_bytes(raw).unwrap();
    let s = b.worksheet("Data").unwrap();
    assert_eq!(s.cells[0].formula.as_deref(), Some("B3*2"));
    assert_eq!(s.cells[1].formula.as_deref(), Some("B4*2"));
    assert_eq!(s.cells[1].value, CellValue::Number(10.0));
}
#[test]
fn rk_and_mulrk_values_do_not_expand_sparse_grid() {
    let raw = workbook(
        8,
        |_| {},
        |s| {
            let mut p = header(65535, 255, 0, 8);
            dword(&mut p, ((-150i32 << 2) as u32) | 3);
            rec(s, 0x27e, &p);
            let mut p = Vec::new();
            word(&mut p, 100);
            word(&mut p, 10);
            word(&mut p, 0);
            dword(&mut p, 42 * 4 + 2);
            word(&mut p, 0);
            dword(&mut p, (1.5f64.to_bits() >> 32) as u32);
            word(&mut p, 11);
            rec(s, 0xbd, &p);
        },
    );
    let b = NativeXlsBook::open_bytes(raw).unwrap();
    let s = b.worksheet("Data").unwrap();
    assert_eq!(s.cells.len(), 3);
    assert_eq!(s.dimension.as_deref(), Some("K101:IV65536"));
    assert_eq!(s.cells[0].value, CellValue::Number(42.0));
    assert_eq!(s.cells[1].value, CellValue::Number(1.5));
    assert_eq!(s.cells[2].value, CellValue::Number(-1.5));
}
#[test]
fn path_input_and_bytes_have_identical_fixed_results() {
    let raw = workbook(8, |_| {}, |s| number(s, 3, 2, 8, 123.0));
    let source = compound(&raw, true, 3, "Workbook");
    let path =
        std::env::temp_dir().join(format!("wolfxl-native-xls-{}-path.xls", std::process::id()));
    std::fs::write(&path, &source).unwrap();
    let b = NativeXlsBook::open_path(&path).unwrap();
    std::fs::remove_file(&path).unwrap();
    let s = b.worksheet("Data").unwrap();
    assert_eq!(s.dimension.as_deref(), Some("C4:C4"));
    assert_eq!(s.cells[0].value, CellValue::Number(123.0));
}
#[test]
fn malformed_container_admission_is_bounded() {
    let raw = workbook(8, |_| {}, |s| number(s, 2, 1, 8, 7.0));
    let source = compound(&raw, true, 3, "Workbook");
    for cut in [0, 1, 7, 64, 511, 512, source.len() - 1] {
        assert!(NativeXlsBook::open_bytes(source[..cut].to_vec()).is_err());
    }
    let mut bad = source.clone();
    set16(&mut bad, 30, 12);
    assert!(NativeXlsBook::open_bytes(bad).is_err());
    let mut bad = source.clone();
    set32(&mut bad, 512 + 4, 1);
    assert!(
        NativeXlsBook::open_bytes(bad).is_err(),
        "directory FAT cycle"
    );
    let mut bad = source.clone();
    set32(&mut bad, 1024 + 128 + 116, u32::MAX - 10);
    assert!(
        NativeXlsBook::open_bytes(bad).is_err(),
        "out of range mini sector"
    );
    let mut bad = source.clone();
    let mini_fat = u32::from_le_bytes(bad[60..64].try_into().unwrap()) as usize;
    set32(&mut bad, (mini_fat + 1) * 512, 0);
    assert!(NativeXlsBook::open_bytes(bad).is_err(), "mini FAT cycle");
    let mut bad = source.clone();
    set32(&mut bad, 44, u32::MAX);
    assert!(
        NativeXlsBook::open_bytes(bad).is_err(),
        "unbounded FAT declaration"
    );
    let mut bad = source;
    set32(&mut bad, 1024 + 128 + 72, 1);
    assert!(
        NativeXlsBook::open_bytes(bad).is_err(),
        "directory sibling cycle"
    );
}
#[test]
fn malformed_biff_boundaries_are_not_empty_workbooks() {
    let raw = workbook(8, |_| {}, |s| number(s, 2, 1, 8, 7.0));
    for cut in 0..raw.len() {
        let b = NativeXlsBook::open_bytes(raw[..cut].to_vec());
        if let Ok(b) = b {
            assert!(b.worksheet("Data").is_err());
        }
    }
    let duplicate = workbook(
        8,
        |_| {},
        |s| {
            number(s, 2, 1, 8, 1.0);
            number(s, 2, 1, 8, 2.0);
        },
    );
    assert!(NativeXlsBook::open_bytes(duplicate)
        .unwrap()
        .worksheet("Data")
        .is_err());
    let invalid_sst = workbook(
        8,
        |g| rec(g, 0xfc, &[1, 0, 0, 0, 255, 255, 255, 255]),
        |_| {},
    );
    assert!(NativeXlsBook::open_bytes(invalid_sst).is_err());
    let missing_cache = workbook(
        8,
        |_| {},
        |s| {
            formula(
                s,
                0,
                0,
                8,
                [0, 0, 0, 0, 0, 0, 255, 255],
                &[0x17, 1, 0, b'a'],
            )
        },
    );
    assert!(NativeXlsBook::open_bytes(missing_cache)
        .unwrap()
        .worksheet("Data")
        .is_err());
    let invalid_bool = workbook(
        8,
        |_| {},
        |s| {
            let mut p = header(0, 0, 0, 8);
            p.extend_from_slice(&[2, 0]);
            rec(s, 0x205, &p);
        },
    );
    assert!(NativeXlsBook::open_bytes(invalid_bool)
        .unwrap()
        .worksheet("Data")
        .is_err());
    let encrypted = workbook(8, |g| rec(g, 0x2f, &[0, 0]), |_| {});
    assert!(matches!(
        NativeXlsBook::open_bytes(encrypted),
        Err(ReaderError::Unsupported(_))
    ));
    let invalid_formula = workbook(8, |_| {}, |s| formula(s, 0, 0, 8, 0f64.to_le_bytes(), &[3]));
    let b = NativeXlsBook::open_bytes(invalid_formula).unwrap();
    let s = b.worksheet("Data").unwrap();
    assert_eq!(s.cells[0].value, CellValue::Number(0.0));
    assert!(s.cells[0]
        .formula
        .as_deref()
        .unwrap()
        .contains("stack underflow"));
}

#[test]
fn normative_function_and_command_registries_remain_distinct() {
    let raw = workbook(
        8,
        |_| {},
        |s| {
            formula(s, 0, 0, 8, 0f64.to_le_bytes(), &[0x1e, 1, 0, 0x21, 0x70, 1]);
            formula(
                s,
                0,
                1,
                8,
                0f64.to_le_bytes(),
                &[0x1e, 5, 0, 0x22, 1, 0, 0x80],
            );
            formula(s, 0, 2, 8, 0f64.to_le_bytes(), &[0x22, 0, 31, 0x80]);
            formula(
                s,
                0,
                3,
                8,
                0f64.to_le_bytes(),
                &[0x1e, 1, 0, 0x22, 1, 106, 0],
            );
        },
    );
    let s = NativeXlsBook::open_bytes(raw)
        .unwrap()
        .worksheet("Data")
        .unwrap();
    assert_eq!(s.cells[0].formula.as_deref(), Some("BAHTTEXT(1)"));
    assert_eq!(s.cells[1].formula.as_deref(), Some("BEEP(5)"));
    assert_eq!(s.cells[2].formula.as_deref(), Some("CALCULATE.NOW()"));
    assert_eq!(s.cells[3].formula.as_deref(), Some("GET.FORMULA(1)"));
    for v in [2, 3] {
        let raw = workbook(
            v,
            |_| {},
            |s| formula(s, 0, 0, v, 0f64.to_le_bytes(), &[0x1e, 5, 0, 0x38, 1, 0]),
        );
        assert_eq!(
            NativeXlsBook::open_bytes(raw)
                .unwrap()
                .worksheet("Data")
                .unwrap()
                .cells[0]
                .formula
                .as_deref(),
            Some("BEEP(5)")
        );
    }
}

#[test]
fn empty_and_truncated_string_continuations_are_errors_not_panics() {
    let raw = workbook(
        8,
        |g| {
            rec(g, 0xfc, &[1, 0, 0, 0, 1, 0, 0, 0, 2, 0, 0, b'a']);
            rec(g, 0x3c, &[]);
            rec(g, 0x3c, &[]);
        },
        |_| {},
    );
    assert!(NativeXlsBook::open_bytes(raw).is_err());
}

fn formula_extra(out: &mut Vec<u8>, row: u16, col: u16, tokens: &[u8], extra: &[u8]) {
    let start = out.len();
    formula(out, row, col, 8, 42f64.to_le_bytes(), tokens);
    out.extend_from_slice(extra);
    let len = out.len() - start - 4;
    set16(out, start + 2, len as u16);
}

#[test]
fn natural_language_labels_resolve_cells_deleted_labels_and_extra_locations() {
    let raw = workbook(
        8,
        |g| {
            let mut p = vec![9, 0, 0];
            p.extend_from_slice(b"LostLabel");
            rec(g, 0x1b9, &p);
        },
        |s| {
            label(s, 0, 0, 8, b"Revenue");
            label(s, 0, 1, 8, b"Expenses");
            let area = [0x25, 1, 0, 2, 0, 1, 0, 1, 0];
            for (i, subtype) in [2, 3, 6, 7, 10].into_iter().enumerate() {
                let mut t = vec![0x18, subtype, 0, 0, 0, 0x40];
                if subtype == 10 {
                    t.extend_from_slice(&area);
                }
                formula_extra(s, 2, i as u16, &t, &[]);
            }
            for (i, subtype) in [1, 16].into_iter().enumerate() {
                let mut t = vec![0x18, subtype, 2, 0, 1, 0];
                if subtype == 16 {
                    t.extend_from_slice(&area);
                }
                formula_extra(s, 3, i as u16, &t, &[]);
            }
            for (i, subtype) in [11, 13, 15].into_iter().enumerate() {
                let mut t = vec![0x18, subtype, 0, 0, 0, 0];
                if subtype == 11 {
                    t.extend_from_slice(&area);
                }
                formula_extra(s, 4, i as u16, &t, &[2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0]);
            }
            formula_extra(s, 5, 2, &[0x18, 2, 251, 255, 255, 0xc0], &[]);
        },
    );
    let book = NativeXlsBook::open_bytes(raw).unwrap();
    let s = book.worksheet("Data").unwrap();
    for cell in s.cells.iter().filter(|c| c.row == 3) {
        assert_eq!(cell.formula.as_deref(), Some("'Revenue'"));
        assert_eq!(cell.value, CellValue::Number(42.0));
    }
    for cell in s.cells.iter().filter(|c| c.row == 4) {
        assert_eq!(cell.formula.as_deref(), Some("'LostLabel'"));
    }
    for cell in s.cells.iter().filter(|c| c.row == 5) {
        assert_eq!(cell.formula.as_deref(), Some("'Revenue' 'Expenses'"));
    }
    assert_eq!(
        s.cells.last().unwrap().formula.as_deref(),
        Some("'Expenses'")
    );
}

#[test]
fn malformed_natural_labels_keep_diagnostics_without_overwriting_caches() {
    for (token, truncated) in [
        (&[0x18, 2, 0, 0, 0][..], true),
        (&[0x18, 4, 0, 0, 0, 0][..], false),
        (&[0x18, 1, 3, 0, 0, 0][..], false),
        (&[0x18, 10, 0, 0, 0, 0, 0x1e, 1, 0][..], false),
    ] {
        let raw = workbook(8, |_| {}, |s| formula_extra(s, 0, 0, token, &[]));
        let result = NativeXlsBook::open_bytes(raw).unwrap().worksheet("Data");
        if truncated {
            assert!(result.is_err());
        } else {
            let s = result.unwrap();
            assert_eq!(s.cells[0].value, CellValue::Number(42.0));
            assert!(s.cells[0]
                .formula
                .as_deref()
                .unwrap()
                .starts_with("Unrecognised formula for cell"));
        }
    }
    let raw = workbook(
        8,
        |_| {},
        |s| {
            formula(
                s,
                0,
                0,
                8,
                [0, 0, 0, 0, 0, 0, 255, 255],
                &[0x17, 1, 0, b'a'],
            );
            formula(
                s,
                0,
                1,
                8,
                [0, 0, 0, 0, 0, 0, 255, 255],
                &[0x17, 1, 0, b'b'],
            );
            rec(s, 0x207, &[1, 0, 0, b'b']);
        },
    );
    assert!(NativeXlsBook::open_bytes(raw)
        .unwrap()
        .worksheet("Data")
        .is_err());
    let raw = workbook(
        8,
        |_| {},
        |s| formula(s, 0, 0, 8, [1, 0, 2, 0, 0, 0, 255, 255], &[0x1d, 1]),
    );
    assert!(NativeXlsBook::open_bytes(raw)
        .unwrap()
        .worksheet("Data")
        .is_err());
}

#[test]
fn installation_relative_external_urls_retain_specification_markers() {
    for marker in [6u8, 8] {
        let raw = workbook(
            8,
            |g| {
                let mut sup = vec![1, 0, 7, 0, 0, 1, marker];
                sup.extend_from_slice(b"a.xla");
                sup.extend_from_slice(&[4, 0, 0]);
                sup.extend_from_slice(b"Data");
                rec(g, 0x1ae, &sup);
                rec(g, 0x17, &[1, 0, 0, 0, 0, 0, 0, 0]);
            },
            |s| formula(s, 0, 0, 8, 42f64.to_le_bytes(), &[0x3a, 0, 0, 0, 0, 0, 0]),
        );
        let book = NativeXlsBook::open_bytes(raw).unwrap();
        let sheet = book.worksheet("Data").unwrap();
        assert_eq!(
            sheet.cells[0].formula,
            Some(format!("'[{}a.xla]Data'!$A$1", char::from(marker)))
        );
        assert_eq!(sheet.cells[0].value, CellValue::Number(42.0));
    }
}

#[test]
fn unsupported_formula_text_preserves_cache_and_name_inventory() {
    let raw = workbook(
        8,
        |g| {
            name(g, 8, "Unsupported", 0, &[0x1a]);
            name(g, 8, "Empty", 0, &[]);
        },
        |s| {
            formula(s, 3, 2, 8, 123.0f64.to_le_bytes(), &[0x1a]);
            formula(s, 4, 2, 8, 456.0f64.to_le_bytes(), &[0x22, 0, 0xff, 0x7f]);
        },
    );
    let book = NativeXlsBook::open_bytes(raw).unwrap();
    assert_eq!(book.named_ranges().len(), 2);
    assert!(book.named_ranges()[0]
        .refers_to
        .starts_with("Unsupported ptg: "));
    assert_eq!(book.named_ranges()[1].refers_to, "empty rgce");
    let s = book.worksheet("Data").unwrap();
    assert_eq!(s.dimension.as_deref(), Some("C4:C5"));
    assert_eq!(s.cells[0].value, CellValue::Number(123.0));
    assert_eq!(s.cells[1].value, CellValue::Number(456.0));
    assert!(s.cells[0]
        .formula
        .as_deref()
        .unwrap()
        .contains("Unrecognised formula for cell (3, 2)"));
    assert!(s.cells[0].formula.as_deref().unwrap().contains("0x1a"));
    assert!(s.cells[1]
        .formula
        .as_deref()
        .unwrap()
        .contains("unknown XLS function id"));
    let truncated = workbook(
        8,
        |_| {},
        |s| formula(s, 0, 0, 8, 7f64.to_le_bytes(), &[0x1f, 0, 0]),
    );
    assert!(NativeXlsBook::open_bytes(truncated)
        .unwrap()
        .worksheet("Data")
        .is_err());
}

#[test]
fn admitted_legacy_codecs_apply_bom_detection_and_replacement() {
    for (page, data, expected) in [
        (932u16, b"\x82".as_slice(), "\u{fffd}"),
        (1252, b"\xef\xbb\xbfhi".as_slice(), "hi"),
        (1200, b"\xff\xfeA\0".as_slice(), "A"),
    ] {
        let mut raw = Vec::new();
        bof(&mut raw, 5, 0x10);
        rec(&mut raw, 0x42, &page.to_le_bytes());
        label(&mut raw, 0, 0, 5, data);
        eof(&mut raw);
        let s = NativeXlsBook::open_bytes(raw)
            .unwrap()
            .worksheet("Sheet1")
            .unwrap();
        assert_eq!(s.cells[0].value, CellValue::String(expected.into()));
    }
    let mut raw = Vec::new();
    bof(&mut raw, 5, 0x10);
    rec(&mut raw, 0x42, &65535u16.to_le_bytes());
    label(&mut raw, 0, 0, 5, b"a");
    eof(&mut raw);
    assert!(NativeXlsBook::open_bytes(raw)
        .unwrap()
        .worksheet("Sheet1")
        .is_err());
}

#[test]
fn documented_bof_version_aliases_keep_known_legacy_cell_layouts() {
    for (version, tag, kind) in [
        (2, 2, 5),
        (2, 7, 5),
        (2, 0x200, 5),
        (3, 0x300, 5),
        (4, 0x400, 5),
        (5, 0x500, 5),
        (8, 0x600, 5),
        (5, 0, 0x1000),
        (8, 0, 5),
    ] {
        let mut raw = workbook(
            version,
            |g| rec(g, 0x42, &1252u16.to_le_bytes()),
            |s| {
                number(s, 4, 2, version, 123.0);
                label(s, 4, 3, version, b"alias");
            },
        );
        set16(&mut raw, 0, 0x809);
        set16(&mut raw, 4, tag);
        set16(&mut raw, 6, kind);
        let b = NativeXlsBook::open_bytes(compound(&raw, true, 3, "Book")).unwrap();
        let s = b.worksheet("Data").unwrap();
        assert_eq!(s.dimension.as_deref(), Some("C5:D5"));
        assert_eq!(s.cells[0].value, CellValue::Number(123.0));
        assert_eq!(s.cells[1].value, CellValue::String("alias".into()));
    }
}
