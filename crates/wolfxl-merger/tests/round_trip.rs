//! Integration test — for every `xl/worksheets/sheet*.xml` ZIP entry under
//! `tests/fixtures/`, call `merge_blocks(xml, vec![])` and assert the output
//! is byte-identical to the input. This is the empty-blocks fast path
//! contract from RFC-011 §5.6, validated against pathological real-world
//! XML (BOMs, x14ac extensions, third-party compat tags, unusual
//! whitespace).

use std::fs::File;
use std::io::Read;
use std::path::{Path, PathBuf};

use wolfxl_merger::merge_blocks;

fn fixtures_dir() -> PathBuf {
    let manifest = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    manifest
        .parent()
        .and_then(|p| p.parent())
        .expect("workspace root")
        .join("tests/fixtures")
}

fn collect_xlsx(dir: &Path, out: &mut Vec<PathBuf>) {
    let Ok(entries) = std::fs::read_dir(dir) else {
        return;
    };
    for entry in entries.flatten() {
        let path = entry.path();
        if path.is_dir() {
            collect_xlsx(&path, out);
        } else if path.extension().and_then(|s| s.to_str()) == Some("xlsx") {
            out.push(path);
        }
    }
}

fn sheet_xml_entries_in(xlsx: &Path) -> Vec<(String, Vec<u8>)> {
    let f = File::open(xlsx).expect("open xlsx");
    let mut zip = zip::ZipArchive::new(f).expect("read zip");
    let mut out = Vec::new();
    for i in 0..zip.len() {
        let mut entry = zip.by_index(i).expect("zip entry");
        let name = entry.name().expect("valid fixture entry name").into_owned();
        // Match xl/worksheets/sheetN.xml or xl/worksheets/chartsheetN.xml,
        // not xl/worksheets/_rels/...
        if !name.starts_with("xl/worksheets/")
            || !name.ends_with(".xml")
            || name.contains("/_rels/")
        {
            continue;
        }
        let mut buf = Vec::new();
        entry.read_to_end(&mut buf).expect("read entry");
        out.push((name, buf));
    }
    out
}

#[test]
fn empty_blocks_byte_identical_for_every_fixture_sheet() {
    let dir = fixtures_dir();
    if !dir.exists() {
        eprintln!("fixtures dir not found at {}, skipping", dir.display());
        return;
    }
    let mut xlsx_files = Vec::new();
    collect_xlsx(&dir, &mut xlsx_files);
    assert!(
        !xlsx_files.is_empty(),
        "no .xlsx fixtures found under {}",
        dir.display()
    );

    let mut total_sheets = 0;
    for xlsx in &xlsx_files {
        let entries = sheet_xml_entries_in(xlsx);
        for (name, bytes) in entries {
            total_sheets += 1;
            let out = merge_blocks(&bytes, vec![]).unwrap_or_else(|e| {
                panic!(
                    "merge_blocks(empty) failed for {} :: {} → {}",
                    xlsx.display(),
                    name,
                    e
                )
            });
            assert_eq!(
                out,
                bytes,
                "merge_blocks(empty) must be byte-identical for {} :: {}",
                xlsx.display(),
                name
            );
        }
    }
    assert!(total_sheets > 0, "no sheet XML entries found in fixtures");
    eprintln!(
        "round-tripped {} sheet XML entries across {} xlsx files",
        total_sheets,
        xlsx_files.len()
    );
}
