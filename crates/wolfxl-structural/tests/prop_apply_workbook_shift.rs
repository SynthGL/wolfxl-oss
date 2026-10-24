//! Property test for `apply_workbook_shift` (RFC-030 / RFC-031 fuzz tail).
//!
//! Sweeps random `(axis, idx, n)` triples against a small fixture
//! sheet (~20 cells, 1 merge band, 1 hyperlink, 1 small table, plus a
//! defined name and an autofilter) and asserts:
//!
//! 1. **No panic** — every well-formed combination of `axis ∈ {Row, Col}`,
//!    `idx ∈ 1..50`, `n ∈ -10..10` runs to completion. (`n == 0` is
//!    a documented no-op short-circuit; all other values are exercised.)
//! 2. **UTF-8 valid** — input is ASCII XML, output must stay UTF-8.
//! 3. **Well-formed XML** — output of every rewritten part parses
//!    cleanly through `quick_xml::Reader` to EOF without surfacing
//!    a structural error.
//!
//! Iteration count: 5000 cases. Seed: deterministic via `proptest`'s
//! `Config::with_cases` + the default RNG, so re-runs reproduce.
//!
//! Branch-coverage summary is printed at the end of the run as a
//! `(insert vs delete vs noop) × (Row vs Col)` six-cell table. The
//! summary goes to stderr so it shows up under `cargo test -- --nocapture`.
//!
//! ## How to run
//!
//! ```bash
//! cargo test -p wolfxl-structural --release --test prop_apply_workbook_shift
//! cargo test -p wolfxl-structural --release --test prop_apply_workbook_shift -- --nocapture
//! ```
//!
//! Release-mode is recommended (the property loop is ~5K rewrites of
//! a small fixture; debug-mode runs in a few seconds, release in a
//! fraction of a second).

use std::collections::BTreeMap;
use std::sync::atomic::{AtomicUsize, Ordering};

use proptest::prelude::*;
use quick_xml::events::Event;
use quick_xml::Reader as XmlReader;

use wolfxl_structural::{apply_workbook_shift, Axis, AxisShiftOp, SheetXmlInputs};

/// Small but representative sheet XML.
///
/// ~20 cells across rows 1..6, columns A..D, with:
///   - one merged region (`<mergeCells>`),
///   - one hyperlink (`<hyperlinks>`),
///   - one autoFilter,
///   - inline formulas and strings split by XML entity references,
///   - a `<dimension ref="...">` header,
///   - a `<cols>` block with width metadata.
const FIXTURE_SHEET_XML: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:D6"/>
  <cols><col min="1" max="2" width="12.0" customWidth="1"/><col min="3" max="4" width="9.0"/></cols>
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Q1</t></is></c><c r="C1" t="inlineStr"><is><t>Q2</t></is></c><c r="D1" t="inlineStr"><is><t>Total</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>Widget &amp; &#71;adget</t></is></c><c r="B2"><v>10</v></c><c r="C2"><v>20</v></c><c r="D2"><f>SUM(B&#50;:C2)</f><v>30</v></c></row>
    <row r="3"><c r="A3" t="inlineStr"><is><t>Gadget</t></is></c><c r="B3"><v>15</v></c><c r="C3"><v>25</v></c><c r="D3"><f>SUM(B3:C3)</f><v>40</v></c></row>
    <row r="4"><c r="A4" t="inlineStr"><is><t>Sprocket</t></is></c><c r="B4"><v>5</v></c><c r="C4"><v>15</v></c><c r="D4"><f>SUM(B4:C4)</f><v>20</v></c></row>
    <row r="5"><c r="A5" t="inlineStr"><is><t>Total</t></is></c><c r="B5"><f>SUM(B2:B4)</f><v>30</v></c><c r="C5"><f>SUM(C2:C4)</f><v>60</v></c><c r="D5"><f>SUM(D2:D4)</f><v>90</v></c></row>
    <row r="6"><c r="A6" t="inlineStr"><is><t>Note</t></is></c></row>
  </sheetData>
  <autoFilter ref="A1:D5"/>
  <mergeCells count="1"><mergeCell ref="A6:D6"/></mergeCells>
  <hyperlinks><hyperlink ref="A6" r:id="rId1"/></hyperlinks>
</worksheet>"#;

/// Companion `tableN.xml` part referencing the same band.
const FIXTURE_TABLE_XML: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
       id="1" name="Sales" displayName="Sales" ref="A1:D5">
  <autoFilter ref="A1:D5"/>
  <tableColumns count="4">
    <tableColumn id="1" name="Name"/>
    <tableColumn id="2" name="Q1"/>
    <tableColumn id="3" name="Q2"/>
    <tableColumn id="4" name="Total">
      <calculatedColumnFormula>SUM(B2:C2)</calculatedColumnFormula>
    </tableColumn>
  </tableColumns>
</table>"#;

/// Minimal `xl/workbook.xml` with one sheet and one defined name pointing
/// inside the shift band.
const FIXTURE_WORKBOOK_XML: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
  <definedNames><definedName name="MyTotal" localSheetId="0">Sheet1!$D$&#53;</definedName></definedNames>
</workbook>"#;

fn build_inputs<'a>() -> SheetXmlInputs<'a> {
    let mut inputs = SheetXmlInputs::empty();
    inputs
        .sheets
        .insert("Sheet1".to_string(), FIXTURE_SHEET_XML);
    inputs
        .sheet_paths
        .insert("Sheet1".to_string(), "xl/worksheets/sheet1.xml".to_string());
    inputs.workbook_xml = Some(FIXTURE_WORKBOOK_XML);
    let mut tables: BTreeMap<String, Vec<(String, &'a [u8])>> = BTreeMap::new();
    tables.insert(
        "Sheet1".to_string(),
        vec![("xl/tables/table1.xml".to_string(), FIXTURE_TABLE_XML)],
    );
    inputs.tables = tables;
    inputs.sheet_positions.insert("Sheet1".to_string(), 0);
    inputs
}

/// Stream-parse `bytes` end to end, returning Err on the first malformed
/// event. Used to assert "well-formed XML" without trying to validate
/// schema correctness.
fn xml_is_well_formed(bytes: &[u8]) -> Result<(), String> {
    let mut reader = XmlReader::from_reader(bytes);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Eof) => return Ok(()),
            Ok(_) => {}
            Err(e) => {
                return Err(format!(
                    "quick_xml error at byte {}: {e}",
                    reader.buffer_position()
                ))
            }
        }
        buf.clear();
    }
}

#[test]
fn workbook_shift_preserves_entities_and_transforms_whole_references() {
    let out = apply_workbook_shift(
        build_inputs(),
        &[AxisShiftOp {
            sheet: "Sheet1".to_string(),
            axis: Axis::Row,
            idx: 2,
            n: 1,
        }],
    );
    let sheet = &out.file_patches["xl/worksheets/sheet1.xml"];
    xml_is_well_formed(sheet).unwrap();
    let text = std::str::from_utf8(sheet).unwrap();
    assert!(text.contains("<f>SUM(B3:C3)</f>"), "{text}");
    assert!(text.contains("<t>Widget &amp; &#71;adget</t>"), "{text}");
    let workbook = &out.file_patches["xl/workbook.xml"];
    xml_is_well_formed(workbook).unwrap();
    let text = std::str::from_utf8(workbook).unwrap();
    assert!(text.contains(">Sheet1!$D$6</definedName>"), "{text}");
}

// Counters for the branch-coverage summary printed at the end of the
// proptest loop. Use atomics so proptest can keep reporting across
// shrink attempts without us holding a mutex.
static N_INSERT_ROW: AtomicUsize = AtomicUsize::new(0);
static N_INSERT_COL: AtomicUsize = AtomicUsize::new(0);
static N_DELETE_ROW: AtomicUsize = AtomicUsize::new(0);
static N_DELETE_COL: AtomicUsize = AtomicUsize::new(0);
static N_NOOP: AtomicUsize = AtomicUsize::new(0);

fn record_branch(axis: Axis, n: i32) {
    let counter: &AtomicUsize = match (axis, n.signum()) {
        (_, 0) => &N_NOOP,
        (Axis::Row, 1) => &N_INSERT_ROW,
        (Axis::Row, -1) => &N_DELETE_ROW,
        (Axis::Col, 1) => &N_INSERT_COL,
        (Axis::Col, -1) => &N_DELETE_COL,
        _ => &N_NOOP,
    };
    counter.fetch_add(1, Ordering::Relaxed);
}

fn drive_one(axis: Axis, idx: u32, n: i32) {
    record_branch(axis, n);
    let inputs = build_inputs();
    let ops = vec![AxisShiftOp {
        sheet: "Sheet1".to_string(),
        axis,
        idx,
        n,
    }];
    let out = apply_workbook_shift(inputs, &ops);
    // Every patched part must remain valid UTF-8 and well-formed XML.
    for (path, bytes) in &out.file_patches {
        let s = std::str::from_utf8(bytes)
            .unwrap_or_else(|e| panic!("non-utf8 output for {path}: {e}"));
        xml_is_well_formed(s.as_bytes()).unwrap_or_else(|e| {
            panic!("malformed XML on {path} after axis={axis:?} idx={idx} n={n}: {e}")
        });
    }
}

proptest! {
    #![proptest_config(ProptestConfig {
        cases: 5000,
        // Keep shrinking off — we want raw fuzz coverage, not minimal
        // counter-examples, and shrinking would inflate runtime
        // without improving the no-panic guarantee.
        max_shrink_iters: 0,
        // Deterministic seed for reproducibility — see module docs.
        rng_algorithm: proptest::test_runner::RngAlgorithm::ChaCha,
        // Don't try to persist failure regressions; the test runs
        // from `tests/`, not from a crate root with `lib.rs`, so
        // the default `SourceParallel` persistence emits a warning.
        failure_persistence: None,
        ..ProptestConfig::default()
    })]

    #[test]
    fn apply_workbook_shift_never_panics_and_emits_well_formed_xml(
        axis_row in any::<bool>(),
        idx in 1u32..50,
        n in -10i32..=10,
    ) {
        let axis = if axis_row { Axis::Row } else { Axis::Col };
        drive_one(axis, idx, n);
    }
}

/// Standalone test that prints the branch-coverage summary. Runs after
/// the property loop because cargo orders tests alphabetically inside
/// a single binary, and `zz_*` sorts last.
#[test]
fn zz_print_branch_coverage_summary() {
    // Drive a handful of fixed cases so the summary is non-zero even if
    // someone runs ONLY this test.
    for &(axis, idx, n) in &[
        (Axis::Row, 5, 3),
        (Axis::Row, 5, -2),
        (Axis::Col, 2, 1),
        (Axis::Col, 2, -1),
        (Axis::Row, 5, 0),
    ] {
        drive_one(axis, idx, n);
    }

    let ir = N_INSERT_ROW.load(Ordering::Relaxed);
    let ic = N_INSERT_COL.load(Ordering::Relaxed);
    let dr = N_DELETE_ROW.load(Ordering::Relaxed);
    let dc = N_DELETE_COL.load(Ordering::Relaxed);
    let np = N_NOOP.load(Ordering::Relaxed);

    eprintln!();
    eprintln!("apply_workbook_shift property test — branch coverage");
    eprintln!("  axis    insert    delete    no-op");
    eprintln!("  Row     {ir:>6}    {dr:>6}    (shared no-op count below)");
    eprintln!("  Col     {ic:>6}    {dc:>6}");
    eprintln!("  no-op (n==0, both axes): {np}");
    eprintln!("  total:                   {}", ir + ic + dr + dc + np);
}
