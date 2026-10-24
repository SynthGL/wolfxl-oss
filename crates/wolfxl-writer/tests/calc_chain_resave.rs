//! Calculation-chain metadata must reflect the current logical cells each save.
use std::collections::BTreeMap;
use std::io::{Cursor, Read};
use wolfxl_writer::model::cell::{WriteCell, WriteCellValue};
use wolfxl_writer::{emit_xlsx, Workbook, Worksheet};

fn parts(workbook: &mut Workbook) -> BTreeMap<String, String> {
    let mut zip = zip::ZipArchive::new(Cursor::new(emit_xlsx(workbook))).unwrap();
    (0..zip.len())
        .map(|i| {
            let mut entry = zip.by_index(i).unwrap();
            let name = entry.name().expect("valid fixture entry name").into_owned();
            let mut xml = String::new();
            entry.read_to_string(&mut xml).unwrap();
            (name, xml)
        })
        .collect()
}

#[test]
fn repeated_save_tracks_dense_formula_overwrites_and_metadata_together() {
    let mut workbook = Workbook::new();
    let mut sheet = Worksheet::new("Data");
    sheet
        .append_dense_row(
            1,
            1,
            vec![WriteCell::new(WriteCellValue::Formula {
                expr: "IF(1<2,\"a&b\",\"c\")".into(),
                result: None,
            })],
        )
        .unwrap();
    workbook.add_sheet(sheet);
    let first = parts(&mut workbook);
    assert_eq!(first, parts(&mut workbook));
    assert!(first["xl/calcChain.xml"].contains("<c r=\"A1\" i=\"1\"/>"));
    assert!(first["xl/worksheets/sheet1.xml"].contains("IF(1&lt;2,\"a&amp;b\",\"c\")"));
    workbook.sheets[0].write_cell(1, 1, WriteCellValue::Number(3.0), None);
    let replaced = parts(&mut workbook);
    assert!(!replaced.contains_key("xl/calcChain.xml"));
    assert!(!replaced["[Content_Types].xml"].contains("calcChain"));
    assert!(!replaced["xl/_rels/workbook.xml.rels"].contains("calcChain"));
    workbook.sheets[0].write_cell(
        2,
        2,
        WriteCellValue::Formula {
            expr: "1+1".into(),
            result: None,
        },
        None,
    );
    let restored = parts(&mut workbook);
    assert!(restored["xl/calcChain.xml"].contains("<c r=\"B2\" i=\"1\"/>"));
    assert!(!restored["xl/calcChain.xml"].contains("r=\"A1\""));
    assert!(restored["[Content_Types].xml"].contains("calcChain"));
    assert!(restored["xl/_rels/workbook.xml.rels"].contains("calcChain"));
}
