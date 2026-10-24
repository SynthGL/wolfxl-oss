//! Sprint Κ Pod-α — file-format magic-byte sniffer + `Read + Seek`
//! source enums shared by the `.xlsb` / `.xls` / `.xlsx` read
//! backends.
//!
//! This crate intentionally contains zero `pyo3` symbols so its
//! tests can be exercised under `cargo test --workspace --exclude
//! wolfxl` (the cdylib's own tests cannot link the Python framework
//! standalone on macOS).

use std::fs::File;
use std::io::{BufReader, Cursor, Read, Seek, SeekFrom};

// ---------------------------------------------------------------------------
// Source enums
// ---------------------------------------------------------------------------

/// Owned reader source for an `Xlsx` workbook.  Either a buffered
/// file handle or an in-memory bytes cursor.
pub enum XlsxSource {
    File(BufReader<File>),
    Bytes(Cursor<Vec<u8>>),
}

impl Read for XlsxSource {
    #[inline]
    fn read(&mut self, buf: &mut [u8]) -> std::io::Result<usize> {
        match self {
            XlsxSource::File(r) => r.read(buf),
            XlsxSource::Bytes(c) => c.read(buf),
        }
    }
}

impl Seek for XlsxSource {
    #[inline]
    fn seek(&mut self, pos: SeekFrom) -> std::io::Result<u64> {
        match self {
            XlsxSource::File(r) => r.seek(pos),
            XlsxSource::Bytes(c) => c.seek(pos),
        }
    }
}

/// Owned reader source for an `Xlsb` workbook.
pub enum XlsbSource {
    File(BufReader<File>),
    Bytes(Cursor<Vec<u8>>),
}

impl Read for XlsbSource {
    #[inline]
    fn read(&mut self, buf: &mut [u8]) -> std::io::Result<usize> {
        match self {
            XlsbSource::File(r) => r.read(buf),
            XlsbSource::Bytes(c) => c.read(buf),
        }
    }
}

impl Seek for XlsbSource {
    #[inline]
    fn seek(&mut self, pos: SeekFrom) -> std::io::Result<u64> {
        match self {
            XlsbSource::File(r) => r.seek(pos),
            XlsbSource::Bytes(c) => c.seek(pos),
        }
    }
}

/// Owned reader source for an `Xls` workbook (BIFF8 / OLE2).
pub enum XlsSource {
    File(BufReader<File>),
    Bytes(Cursor<Vec<u8>>),
}

impl Read for XlsSource {
    #[inline]
    fn read(&mut self, buf: &mut [u8]) -> std::io::Result<usize> {
        match self {
            XlsSource::File(r) => r.read(buf),
            XlsSource::Bytes(c) => c.read(buf),
        }
    }
}

impl Seek for XlsSource {
    #[inline]
    fn seek(&mut self, pos: SeekFrom) -> std::io::Result<u64> {
        match self {
            XlsSource::File(r) => r.seek(pos),
            XlsSource::Bytes(c) => c.seek(pos),
        }
    }
}

// ---------------------------------------------------------------------------
// File-format classification
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FileFormat {
    Xlsx,
    Xlsb,
    Xls,
    Ods,
    Unknown,
}

impl FileFormat {
    pub fn as_str(self) -> &'static str {
        match self {
            FileFormat::Xlsx => "xlsx",
            FileFormat::Xlsb => "xlsb",
            FileFormat::Xls => "xls",
            FileFormat::Ods => "ods",
            FileFormat::Unknown => "unknown",
        }
    }
}

/// OLE2 (Compound Document) magic — used by legacy `.xls`.
pub const OLE2_MAGIC: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

/// Standard ZIP local-file-header magic — shared by `.xlsx`,
/// `.xlsb`, and `.ods` containers.
pub const ZIP_MAGIC: [u8; 4] = [0x50, 0x4B, 0x03, 0x04];

/// Classify a file by inspecting its raw bytes.  When the buffer is
/// a ZIP container we crack it open to distinguish `.xlsx`
/// (`xl/workbook.xml`) from `.xlsb` (`xl/workbook.bin`) from `.ods`
/// (`mimetype` payload).
pub fn classify_file_format_bytes(data: &[u8]) -> FileFormat {
    if data.len() >= 8 && data[..8] == OLE2_MAGIC {
        return FileFormat::Xls;
    }
    if data.len() >= 4 && data[..4] == ZIP_MAGIC {
        return classify_zip_bytes(data);
    }
    FileFormat::Unknown
}

/// Classify a file path.  Streams the magic-byte prefix from disk
/// then re-opens the file as a zip to inspect the central directory
/// when needed.
pub fn classify_file_format_path(path: &str) -> FileFormat {
    let mut head = [0u8; 8];
    let read_n = match File::open(path) {
        Ok(mut f) => match f.read(&mut head) {
            Ok(n) => n,
            Err(_) => return FileFormat::Unknown,
        },
        Err(_) => return FileFormat::Unknown,
    };
    if read_n >= 8 && head == OLE2_MAGIC {
        return FileFormat::Xls;
    }
    if read_n >= 4 && head[..4] == ZIP_MAGIC {
        match File::open(path) {
            Ok(f) => match zip::ZipArchive::new(BufReader::new(f)) {
                Ok(mut zip) => classify_zip_archive(&mut zip),
                Err(_) => FileFormat::Unknown,
            },
            Err(_) => FileFormat::Unknown,
        }
    } else {
        FileFormat::Unknown
    }
}

fn classify_zip_bytes(data: &[u8]) -> FileFormat {
    let cursor = Cursor::new(data);
    match zip::ZipArchive::new(cursor) {
        Ok(mut zip) => classify_zip_archive(&mut zip),
        Err(_) => FileFormat::Unknown,
    }
}

fn classify_zip_archive<R: Read + Seek>(zip: &mut zip::ZipArchive<R>) -> FileFormat {
    let mut saw_workbook_xml = false;
    let mut saw_workbook_bin = false;
    let mut saw_mimetype = false;
    let mut saw_content_xml = false;
    for i in 0..zip.len() {
        let entry = match zip.by_index_raw(i) {
            Ok(entry) => entry,
            Err(_) => continue,
        };
        let name = match entry.name() {
            Ok(name) => name,
            Err(_) => continue,
        };
        match name.as_ref() {
            "xl/workbook.xml" => saw_workbook_xml = true,
            "xl/workbook.bin" => saw_workbook_bin = true,
            "mimetype" => saw_mimetype = true,
            "content.xml" => saw_content_xml = true,
            _ => {}
        }
    }
    if saw_workbook_bin {
        return FileFormat::Xlsb;
    }
    if saw_workbook_xml {
        return FileFormat::Xlsx;
    }
    if saw_mimetype && saw_content_xml {
        return FileFormat::Ods;
    }
    FileFormat::Unknown
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;

    const XLSB_BYTES: &[u8] = include_bytes!("../../../tests/fixtures/sprint_kappa_smoke.xlsb");
    const XLS_BYTES: &[u8] = include_bytes!("../../../tests/fixtures/sprint_kappa_smoke.xls");
    const XLSX_BYTES: &[u8] = include_bytes!("../../../tests/fixtures/sprint_kappa_smoke.xlsx");
    const ODS_BYTES: &[u8] = include_bytes!("../../../tests/fixtures/sprint_kappa_smoke.ods");

    #[test]
    fn classify_xlsx_bytes_smoke() {
        assert_eq!(classify_file_format_bytes(XLSX_BYTES), FileFormat::Xlsx);
    }

    #[test]
    fn classify_legacy_encoded_zip_filename() {
        let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let options = zip::write::SimpleFileOptions::default();
        writer.start_file("bad.txt", options).unwrap();
        writer.start_file("xl/workbook.xml", options).unwrap();
        let mut bytes = writer.finish().unwrap().into_inner();
        let central = bytes.windows(4).position(|w| w == b"PK\x01\x02").unwrap();
        // ZIP 9 falls back to CP437 for non-UTF-8 bytes, even with this flag.
        bytes[7] |= 0x08;
        bytes[30] = 0xff;
        bytes[central + 9] |= 0x08;
        bytes[central + 46] = 0xff;
        let mut archive = zip::ZipArchive::new(Cursor::new(&bytes)).unwrap();
        assert_eq!(
            archive.by_index_raw(0).unwrap().name().unwrap(),
            "\u{a0}ad.txt"
        );
        assert_eq!(classify_file_format_bytes(&bytes), FileFormat::Xlsx);
    }

    #[test]
    fn classify_xlsb_bytes_smoke() {
        assert_eq!(classify_file_format_bytes(XLSB_BYTES), FileFormat::Xlsb);
    }

    #[test]
    fn classify_xls_bytes_smoke() {
        assert_eq!(classify_file_format_bytes(XLS_BYTES), FileFormat::Xls);
    }

    #[test]
    fn classify_ods_bytes_smoke() {
        assert_eq!(classify_file_format_bytes(ODS_BYTES), FileFormat::Ods);
    }

    #[test]
    fn classify_unknown_bytes_smoke() {
        assert_eq!(
            classify_file_format_bytes(b"not a spreadsheet"),
            FileFormat::Unknown
        );
        assert_eq!(classify_file_format_bytes(&[]), FileFormat::Unknown);
    }

    #[test]
    fn fileformat_as_str_round_trip() {
        for (variant, expected) in [
            (FileFormat::Xlsx, "xlsx"),
            (FileFormat::Xlsb, "xlsb"),
            (FileFormat::Xls, "xls"),
            (FileFormat::Ods, "ods"),
            (FileFormat::Unknown, "unknown"),
        ] {
            assert_eq!(variant.as_str(), expected);
        }
    }

    #[test]
    fn classify_xlsx_via_path() {
        let path = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../tests/fixtures/sprint_kappa_smoke.xlsx"
        );
        assert_eq!(classify_file_format_path(path), FileFormat::Xlsx);
    }

    #[test]
    fn classify_xlsb_via_path() {
        let path = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../tests/fixtures/sprint_kappa_smoke.xlsb"
        );
        assert_eq!(classify_file_format_path(path), FileFormat::Xlsb);
    }

    #[test]
    fn classify_xls_via_path() {
        let path = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../tests/fixtures/sprint_kappa_smoke.xls"
        );
        assert_eq!(classify_file_format_path(path), FileFormat::Xls);
    }

    #[test]
    fn classify_ods_via_path() {
        let path = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../tests/fixtures/sprint_kappa_smoke.ods"
        );
        assert_eq!(classify_file_format_path(path), FileFormat::Ods);
    }

    #[test]
    fn classify_missing_path_is_unknown() {
        assert_eq!(
            classify_file_format_path("/tmp/wolfxl_does_not_exist_for_sprint_kappa.xlsb"),
            FileFormat::Unknown
        );
    }

    #[test]
    fn xlsb_source_delegates_read_and_seek() {
        let mut source = XlsbSource::Bytes(Cursor::new(XLSB_BYTES.to_vec()));
        let mut header = [0_u8; 4];
        source.read_exact(&mut header).unwrap();
        assert_eq!(&header, b"PK\x03\x04");
        source.seek(SeekFrom::End(-4)).unwrap();
        let mut trailer = [0_u8; 4];
        source.read_exact(&mut trailer).unwrap();
        assert_eq!(trailer, XLSB_BYTES[XLSB_BYTES.len() - 4..]);
    }

    #[test]
    fn xls_source_delegates_read_and_seek() {
        let mut source = XlsSource::Bytes(Cursor::new(XLS_BYTES.to_vec()));
        let mut header = [0_u8; 8];
        source.read_exact(&mut header).unwrap();
        assert_eq!(header, [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);
        source.seek(SeekFrom::Start(512)).unwrap();
        let mut sector = [0_u8; 16];
        source.read_exact(&mut sector).unwrap();
        assert_eq!(sector, XLS_BYTES[512..528]);
    }
}
