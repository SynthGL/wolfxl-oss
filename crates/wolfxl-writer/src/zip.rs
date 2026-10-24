//! Deterministic xlsx ZIP packager.
//!
//! # Behavior
//!
//! - Writes xlsx entries in the order supplied by the caller. The top-level
//!   facade (Wave 4) is responsible for canonical ordering — we never
//!   re-sort. Excel's canonical order is:
//!
//!   1. `[Content_Types].xml`
//!   2. `_rels/.rels`
//!   3. `xl/workbook.xml`
//!   4. `xl/_rels/workbook.xml.rels`
//!   5. `xl/worksheets/sheet1.xml`, sheet2, ...
//!   6. `xl/worksheets/_rels/sheet*.xml.rels`
//!   7. `xl/theme/theme1.xml`
//!   8. `xl/styles.xml`
//!   9. `xl/sharedStrings.xml`
//!   10. `xl/tables/table*.xml`
//!   11. `xl/comments/comments*.xml` + `xl/drawings/vmlDrawing*.vml`
//!   12. `docProps/core.xml`, `docProps/app.xml`
//!
//! - Stamps each entry's mtime from `WOLFXL_TEST_EPOCH` if set (for diff
//!   harness byte parity), otherwise lets the `zip` crate use its default.
//!
//! - DEFLATE level 1 for entries ≥ 128 bytes, STORE
//!   for anything smaller — tiny XML stubs don't benefit from DEFLATE and
//!   STORE keeps the hot path simple.
//!
//! # Determinism
//!
//! Setting `WOLFXL_TEST_EPOCH=0` (or any integer) forces every entry's
//! mtime to the same fixed Unix timestamp. Two calls to `package` with
//! the same input produce byte-identical output in that mode.

use std::io::{Cursor, Seek, Write};

use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, DateTime as ZipDateTime, ZipWriter};

/// Entries below this size skip DEFLATE and use STORE. Small XML stubs
/// (`_rels/.rels`, `docProps/app.xml`) gain nothing from compression.
const DEFLATE_MIN_BYTES: usize = 128;
const FAST_SAVE_DEFLATE_LEVEL: i64 = 1;

/// A single (path, bytes) pair awaiting packaging. Construct one per
/// emitted OOXML part; hand a `Vec<ZipEntry>` to the packager.
#[derive(Debug, Clone)]
pub struct ZipEntry {
    /// The full path inside the xlsx, e.g. `"xl/worksheets/sheet1.xml"`.
    pub path: String,
    pub bytes: Vec<u8>,
}

/// Package a sequence of entries into a complete xlsx. Returns the
/// serialized container bytes.
///
/// Thin wrapper around [`package_to`]: allocates a `Vec<u8>` and writes
/// the archive into it. Prefer [`package_to`] when you have a destination
/// `Write + Seek` sink (e.g. `BufWriter<File>`) — it skips the final
/// in-memory materialisation.
pub fn package(entries: &[ZipEntry]) -> Result<Vec<u8>, std::io::Error> {
    let mut cursor = Cursor::new(Vec::<u8>::new());
    package_to(entries, &mut cursor)?;
    Ok(cursor.into_inner())
}

/// Stream-package a sequence of entries directly into `dest`.
///
/// Entries are written in the order provided. Each entry uses fast DEFLATE
/// unless its body is under [`DEFLATE_MIN_BYTES`] bytes, in
/// which case it uses STORE. When `WOLFXL_TEST_EPOCH` is set, every
/// entry receives that epoch as its mtime.
///
/// `dest` must be `Write + Seek` because the underlying `ZipWriter`
/// patches local-file-header sizes after each entry. Production callers
/// pass a `BufWriter<File>` (which `seek`s into the kernel page cache);
/// the in-memory wrapper [`package`] passes a `Cursor<Vec<u8>>`.
pub fn package_to<W: Write + Seek>(
    entries: &[ZipEntry],
    dest: &mut W,
) -> Result<(), std::io::Error> {
    let mut writer = ZipWriter::new(dest);
    let epoch_override = test_epoch_override().and_then(epoch_to_zip_datetime);

    for entry in entries {
        let method = if entry.bytes.len() < DEFLATE_MIN_BYTES {
            CompressionMethod::Stored
        } else {
            CompressionMethod::Deflated
        };
        let mut opts = SimpleFileOptions::default().compression_method(method);
        if method == CompressionMethod::Deflated {
            opts = opts.compression_level(Some(FAST_SAVE_DEFLATE_LEVEL));
        }
        if let Some(dt) = epoch_override {
            opts = opts.last_modified_time(dt);
        }
        writer
            .start_file(entry.path.clone(), opts)
            .map_err(zip_to_io)?;
        writer.write_all(&entry.bytes)?;
    }
    writer.finish().map_err(zip_to_io)?;
    Ok(())
}

/// Read the `WOLFXL_TEST_EPOCH` env var; if set (to any value including
/// "0"), return it as the mtime to stamp on every entry. Otherwise
/// return `None` and the packager uses the `zip` crate's default.
pub fn test_epoch_override() -> Option<i64> {
    std::env::var("WOLFXL_TEST_EPOCH")
        .ok()
        .and_then(|s| s.parse().ok())
}

/// Convert a Unix epoch second count into a ZIP-compatible DOS datetime.
///
/// ZIP DOS datetimes can only represent years 1980–2107. Unix epoch (0) is
/// 1970-01-01 which is below that range, so we clamp to the minimum
/// representable ZIP datetime (1980-01-01 00:00:00) for values earlier
/// than 1980. Values after 2107 get clamped to the maximum.
fn epoch_to_zip_datetime(epoch_secs: i64) -> Option<ZipDateTime> {
    let dt = chrono::DateTime::<chrono::Utc>::from_timestamp(epoch_secs, 0)?;
    // Clamp to the ZIP-representable range [1980-01-01, 2107-12-31].
    let year: u16 = dt
        .naive_utc()
        .date()
        .format("%Y")
        .to_string()
        .parse()
        .ok()?;
    if year < 1980 {
        return ZipDateTime::from_date_and_time(1980, 1, 1, 0, 0, 0).ok();
    }
    if year > 2107 {
        return ZipDateTime::from_date_and_time(2107, 12, 31, 23, 59, 58).ok();
    }
    use chrono::{Datelike, Timelike};
    let naive = dt.naive_utc();
    ZipDateTime::from_date_and_time(
        naive.year() as u16,
        naive.month() as u8,
        naive.day() as u8,
        naive.hour() as u8,
        naive.minute() as u8,
        naive.second() as u8,
    )
    .ok()
}

fn zip_to_io(err: zip::result::ZipError) -> std::io::Error {
    std::io::Error::other(err.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::test_utils::EpochGuard;
    use std::io::Read;

    #[test]
    fn empty_input_produces_valid_empty_zip() {
        let bytes = package(&[]).expect("package empty");
        // The end-of-central-directory record is the last 22 bytes for a
        // zip with no entries and no comment. Its signature is PK\x05\x06.
        assert!(bytes.len() >= 22, "zip must have at least EOCD");
        let sig = &bytes[bytes.len() - 22..bytes.len() - 22 + 4];
        assert_eq!(sig, b"PK\x05\x06", "empty zip must end with EOCD sig");

        // And zip should be able to read it as an archive.
        let archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("read empty zip");
        assert_eq!(archive.len(), 0);
    }

    #[test]
    fn single_entry_round_trips() {
        let entry = ZipEntry {
            path: "hello.txt".to_string(),
            bytes: b"world".to_vec(),
        };
        let bytes = package(std::slice::from_ref(&entry)).expect("package");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("open archive");
        assert_eq!(archive.len(), 1);
        let mut file = archive.by_name("hello.txt").expect("by_name");
        let mut out = String::new();
        file.read_to_string(&mut out).expect("read");
        assert_eq!(out, "world");
    }

    #[test]
    fn large_entry_is_deflated_small_entry_is_stored() {
        // Tiny entry uses STORE.
        let small = ZipEntry {
            path: "s".to_string(),
            bytes: b"abc".to_vec(),
        };
        // Larger-than-threshold entry uses DEFLATE.
        let large = ZipEntry {
            path: "l".to_string(),
            bytes: b"abcdefghij".repeat(32), // 320 bytes
        };
        let bytes = package(&[small, large]).expect("package");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("open archive");
        {
            let s = archive.by_name("s").expect("small");
            assert_eq!(s.compression(), CompressionMethod::Stored);
        }
        {
            let l = archive.by_name("l").expect("large");
            assert_eq!(l.compression(), CompressionMethod::Deflated);
        }
    }

    #[test]
    fn order_is_preserved() {
        let a = ZipEntry {
            path: "z.txt".to_string(),
            bytes: b"z".to_vec(),
        };
        let b = ZipEntry {
            path: "a.txt".to_string(),
            bytes: b"a".to_vec(),
        };
        let c = ZipEntry {
            path: "m.txt".to_string(),
            bytes: b"m".to_vec(),
        };
        let bytes = package(&[a, b, c]).expect("package");
        let archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("open");
        let names: Vec<String> = archive
            .file_names()
            .map(|s| s.expect("valid fixture entry name").into_owned())
            .collect();
        assert_eq!(names, vec!["z.txt", "a.txt", "m.txt"]);
    }

    #[test]
    fn identical_input_with_epoch_produces_identical_bytes() {
        let _guard = EpochGuard::set("0");
        let entries = vec![
            ZipEntry {
                path: "a.xml".to_string(),
                bytes: b"<?xml version=\"1.0\"?><root/>".to_vec(),
            },
            ZipEntry {
                path: "sub/b.xml".to_string(),
                bytes: b"<b>".repeat(64),
            },
        ];
        let bytes1 = package(&entries).expect("package 1");
        let bytes2 = package(&entries).expect("package 2");
        assert_eq!(
            bytes1, bytes2,
            "WOLFXL_TEST_EPOCH=0 should yield byte-identical output"
        );
    }

    #[test]
    fn test_epoch_override_parses() {
        let _guard = EpochGuard::set("1700000000");
        assert_eq!(test_epoch_override(), Some(1_700_000_000));
    }

    #[test]
    fn epoch_to_zip_datetime_clamps_pre_1980() {
        // Unix epoch 0 (1970-01-01) is before ZIP's minimum (1980-01-01).
        let dt = epoch_to_zip_datetime(0).expect("clamp to 1980");
        assert!(dt.is_valid());
    }

    #[test]
    fn epoch_to_zip_datetime_representable() {
        // 2024-01-01T00:00:00Z = 1704067200
        let dt = epoch_to_zip_datetime(1_704_067_200).expect("2024-01-01");
        assert!(dt.is_valid());
    }

    #[test]
    fn package_to_matches_package_byte_for_byte() {
        // The streaming `package_to` path and the buffered `package` path
        // must produce identical archives. Test under WOLFXL_TEST_EPOCH so
        // mtimes are pinned and bytes are deterministic.
        let _guard = EpochGuard::set("0");
        let entries = vec![
            ZipEntry {
                path: "a.xml".to_string(),
                bytes: b"<?xml version=\"1.0\"?><root/>".to_vec(),
            },
            ZipEntry {
                path: "xl/worksheets/sheet1.xml".to_string(),
                bytes: b"<sheet>x</sheet>".repeat(50),
            },
        ];
        let buffered = package(&entries).expect("package");
        let mut streamed = Cursor::new(Vec::<u8>::new());
        package_to(&entries, &mut streamed).expect("package_to");
        assert_eq!(
            buffered,
            streamed.into_inner(),
            "package_to must produce byte-identical output to package"
        );
    }
}
