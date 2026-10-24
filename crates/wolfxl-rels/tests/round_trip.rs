//! Integration test — for every `.rels` ZIP entry under `tests/fixtures/`,
//! parse → serialize → parse and assert structural equality. Catches any
//! attribute we silently dropped during the round-trip.

use std::fs::File;
use std::io::Read;
use std::path::{Path, PathBuf};

use wolfxl_rels::RelsGraph;

fn fixtures_dir() -> PathBuf {
    // Run from the repo root via `cargo test --workspace`.
    // CARGO_MANIFEST_DIR is the wolfxl-rels crate dir; jump two levels up.
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

fn rels_entries_in(xlsx: &Path) -> Vec<(String, Vec<u8>)> {
    let f = File::open(xlsx).expect("open xlsx");
    let mut zip = zip::ZipArchive::new(f).expect("read zip");
    let mut out = Vec::new();
    for i in 0..zip.len() {
        let mut entry = zip.by_index(i).expect("zip entry");
        let name = entry.name().expect("valid fixture entry name").into_owned();
        if !name.ends_with(".rels") {
            continue;
        }
        let mut buf = Vec::new();
        entry.read_to_end(&mut buf).expect("read entry");
        out.push((name, buf));
    }
    out
}

#[test]
fn round_trip_every_rels_in_fixtures() {
    let dir = fixtures_dir();
    if !dir.exists() {
        // The crate is published / built outside the repo — fixtures aren't
        // available. This is a tests/ dir so cargo skips it cleanly.
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

    let mut total_rels = 0;
    for xlsx in &xlsx_files {
        let entries = rels_entries_in(xlsx);
        for (name, bytes) in entries {
            total_rels += 1;
            let g1 = RelsGraph::parse(&bytes).unwrap_or_else(|e| {
                panic!("parse failed for {} :: {} → {}", xlsx.display(), name, e)
            });
            let serialized = g1.serialize();
            let g2 = RelsGraph::parse(&serialized).unwrap_or_else(|e| {
                panic!("re-parse failed for {} :: {} → {}", xlsx.display(), name, e)
            });
            assert_eq!(
                g1,
                g2,
                "round-trip mismatch for {} :: {}",
                xlsx.display(),
                name
            );
            // serialize is deterministic on a fixed-point graph
            let serialized2 = g2.serialize();
            assert_eq!(
                serialized,
                serialized2,
                "serialize is not byte-stable for {} :: {}",
                xlsx.display(),
                name
            );
        }
    }
    assert!(total_rels > 0, "no .rels entries found in fixtures");
    eprintln!(
        "round-tripped {} .rels entries across {} xlsx files",
        total_rels,
        xlsx_files.len()
    );
}
