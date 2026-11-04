//! One-sheet binary XLSB packages built record by record, so tests control
//! the stored number, its cell format, and the workbook date system.

use std::io::{Cursor, Write};

use zip::write::SimpleFileOptions;
use zip::ZipWriter;

/// Cell format 0: General.
pub const GENERAL: u32 = 0;
/// Cell format 1: the built-in `m/d/yy h:mm` date-time format (id 22).
pub const DATE_TIME: u32 = 1;

const WORKBOOK_RELS: &str = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.bin"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.bin"/></Relationships>"#;

/// Package bytes for a workbook whose single sheet `Data` holds `cells`, each
/// `(row, column, cell format, stored number)` with one-based coordinates,
/// listed in row-major order.
pub fn book(date1904: bool, cells: &[(u32, u32, u32, f64)]) -> Vec<u8> {
    let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
    let options = SimpleFileOptions::default();
    for (name, bytes) in [
        (
            "xl/_rels/workbook.bin.rels",
            WORKBOOK_RELS.as_bytes().to_vec(),
        ),
        ("xl/workbook.bin", workbook(date1904)),
        ("xl/styles.bin", styles()),
        ("xl/worksheets/sheet1.bin", sheet(cells)),
    ] {
        writer.start_file(name, options).expect("start member");
        writer.write_all(&bytes).expect("write member");
    }
    writer.finish().expect("finish package").into_inner()
}

fn workbook(date1904: bool) -> Vec<u8> {
    let mut out = Vec::new();
    record(&mut out, 0x0083, &[]); // BrtBeginBook
    let mut properties = Vec::new();
    put_u32(&mut properties, u32::from(date1904)); // f1904 is bit 0
    put_u32(&mut properties, 0);
    put_wide(&mut properties, "");
    record(&mut out, 0x0099, &properties); // BrtWbProp
    record(&mut out, 0x008f, &[]); // BrtBeginBundleShs
    let mut sheet = Vec::new();
    put_u32(&mut sheet, 0); // visible
    put_u32(&mut sheet, 1);
    put_wide(&mut sheet, "rId1");
    put_wide(&mut sheet, "Data");
    record(&mut out, 0x009c, &sheet); // BrtBundleSh
    record(&mut out, 0x0090, &[]); // BrtEndBundleShs
    record(&mut out, 0x0084, &[]); // BrtEndBook
    out
}

fn styles() -> Vec<u8> {
    let mut out = Vec::new();
    record(&mut out, 0x0278, &[]); // BrtBeginStyleSheet
    record(&mut out, 0x0269, &[]); // BrtBeginCellXFs
    for number_format in [0_u16, 22] {
        let mut xf = Vec::new();
        put_u16(&mut xf, 0); // parent style
        put_u16(&mut xf, number_format);
        xf.extend_from_slice(&[0; 12]); // font, fill, border, alignment
        record(&mut out, 0x002f, &xf); // BrtXF
    }
    record(&mut out, 0x026a, &[]); // BrtEndCellXFs
    record(&mut out, 0x0279, &[]); // BrtEndStyleSheet
    out
}

fn sheet(cells: &[(u32, u32, u32, f64)]) -> Vec<u8> {
    let mut out = Vec::new();
    record(&mut out, 0x0081, &[]); // BrtBeginSheet
    record(&mut out, 0x0091, &[]); // BrtBeginSheetData
    let mut current = None;
    for &(row, column, format, value) in cells {
        if current != Some(row) {
            let mut header = Vec::new();
            put_u32(&mut header, row - 1);
            record(&mut out, 0x0000, &header); // BrtRowHdr
            current = Some(row);
        }
        let mut cell = Vec::new();
        put_u32(&mut cell, column - 1);
        put_u32(&mut cell, format);
        cell.extend_from_slice(&value.to_le_bytes());
        record(&mut out, 0x0005, &cell); // BrtCellReal
    }
    record(&mut out, 0x0092, &[]); // BrtEndSheetData
    record(&mut out, 0x0082, &[]); // BrtEndSheet
    out
}

/// One BIFF12 record: a one- or two-byte type, a 7-bit varint length, and
/// the payload.
fn record(out: &mut Vec<u8>, kind: u16, payload: &[u8]) {
    if kind < 0x80 {
        out.push(kind as u8);
    } else {
        out.push((kind & 0x7f) as u8 | 0x80);
        out.push((kind >> 7) as u8);
    }
    let mut length = payload.len();
    loop {
        let byte = (length & 0x7f) as u8;
        length >>= 7;
        out.push(if length == 0 { byte } else { byte | 0x80 });
        if length == 0 {
            break;
        }
    }
    out.extend_from_slice(payload);
}

fn put_u16(out: &mut Vec<u8>, value: u16) {
    out.extend_from_slice(&value.to_le_bytes());
}

fn put_u32(out: &mut Vec<u8>, value: u32) {
    out.extend_from_slice(&value.to_le_bytes());
}

fn put_wide(out: &mut Vec<u8>, text: &str) {
    let units: Vec<u16> = text.encode_utf16().collect();
    put_u32(out, units.len() as u32);
    for unit in units {
        put_u16(out, unit);
    }
}
