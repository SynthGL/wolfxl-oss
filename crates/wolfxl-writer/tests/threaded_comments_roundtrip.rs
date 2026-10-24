//! RFC-068 / G08 — end-to-end emit gate for threaded comments.
//!
//! Builds a `Workbook` with one sheet, one top-level threaded comment + one
//! reply, and one Person in the registry. Runs `emit_xlsx`, opens the
//! resulting bytes with `zip::ZipArchive`, and asserts:
//!
//!   1. Both new parts are present at the canonical paths.
//!   2. `[Content_Types].xml` declares overrides for them.
//!   3. `xl/_rels/workbook.xml.rels` carries the workbook→personList relationship.
//!   4. `xl/worksheets/_rels/sheet1.xml.rels` carries the sheet→threadedComments
//!      relationship.
//!   5. The legacy `xl/comments/comments1.xml` carries the synthetic
//!      `tc={guid}` author and the `[Threaded comment]` placeholder body.

use std::collections::HashSet;
use std::io::{Cursor, Read};

use wolfxl_writer::emit_xlsx;
use wolfxl_writer::model::threaded_comment::{Person, ThreadedComment};
use wolfxl_writer::model::workbook::Workbook;
use wolfxl_writer::model::worksheet::Worksheet;

fn read_zip_part(bytes: &[u8], path: &str) -> Option<String> {
    let mut zip = zip::ZipArchive::new(Cursor::new(bytes)).ok()?;
    let mut file = zip.by_name(path).ok()?;
    let mut s = String::new();
    file.read_to_string(&mut s).ok()?;
    Some(s)
}

fn list_zip_paths(bytes: &[u8]) -> HashSet<String> {
    let zip = zip::ZipArchive::new(Cursor::new(bytes)).expect("open xlsx zip");
    zip.file_names()
        .map(|s| s.expect("valid fixture entry name").into_owned())
        .collect()
}

#[test]
fn threaded_comments_round_trip_through_emit() {
    let mut wb = Workbook::new();

    // Workbook-level person registry.
    wb.persons.push(Person {
        display_name: "Alice".to_string(),
        id: "{P-ALICE}".to_string(),
        user_id: "alice@example.com".to_string(),
        provider_id: "None".to_string(),
    });

    // One sheet with a top-level thread + one reply.
    let mut sheet = Worksheet::new("Notes");
    sheet.threaded_comments.push(ThreadedComment {
        id: "{T-PARENT}".to_string(),
        cell_ref: "A1".to_string(),
        person_id: "{P-ALICE}".to_string(),
        created: "2024-09-12T15:31:01.42".to_string(),
        parent_id: None,
        text: "Looks wrong".to_string(),
        done: false,
    });
    sheet.threaded_comments.push(ThreadedComment {
        id: "{T-REPLY}".to_string(),
        cell_ref: "A1".to_string(),
        person_id: "{P-ALICE}".to_string(),
        created: "2024-09-12T15:33:00.00".to_string(),
        parent_id: Some("{T-PARENT}".to_string()),
        text: "Agreed; investigating".to_string(),
        done: false,
    });
    wb.add_sheet(sheet);

    let bytes = emit_xlsx(&mut wb);

    // 1. Canonical part paths.
    let paths = list_zip_paths(&bytes);
    assert!(
        paths.contains("xl/threadedComments/threadedComments1.xml"),
        "threadedComments1.xml present: paths = {paths:?}"
    );
    assert!(
        paths.contains("xl/persons/personList.xml"),
        "personList.xml present: paths = {paths:?}"
    );
    assert!(
        paths.contains("xl/comments/comments1.xml"),
        "legacy placeholder comments1.xml present: paths = {paths:?}"
    );

    // 2. Content type overrides.
    let ct = read_zip_part(&bytes, "[Content_Types].xml").expect("content types present");
    assert!(
        ct.contains("/xl/threadedComments/threadedComments1.xml"),
        "ct override for threadedComments: {ct}"
    );
    assert!(
        ct.contains("/xl/persons/personList.xml"),
        "ct override for personList: {ct}"
    );
    assert!(
        ct.contains("application/vnd.ms-excel.threadedcomments+xml"),
        "ct contains threadedcomments mime: {ct}"
    );
    assert!(
        ct.contains("application/vnd.ms-excel.person+xml"),
        "ct contains person mime: {ct}"
    );

    // 3. Workbook-level rels — personList relationship.
    let wb_rels = read_zip_part(&bytes, "xl/_rels/workbook.xml.rels").expect("wb rels present");
    assert!(
        wb_rels.contains("persons/personList.xml"),
        "workbook rels target personList: {wb_rels}"
    );
    assert!(
        wb_rels.contains("relationships/person"),
        "workbook rels rel-type for person: {wb_rels}"
    );

    // 4. Sheet rels — threadedComments relationship.
    let sheet_rels =
        read_zip_part(&bytes, "xl/worksheets/_rels/sheet1.xml.rels").expect("sheet1 rels present");
    assert!(
        sheet_rels.contains("../threadedComments/threadedComments1.xml"),
        "sheet rels target threadedComments: {sheet_rels}"
    );
    assert!(
        sheet_rels.contains("relationships/threadedComment"),
        "sheet rels rel-type for threadedComment: {sheet_rels}"
    );
    // Legacy comments + vmlDrawing are still emitted (the synthesized
    // placeholder forces them).
    assert!(
        sheet_rels.contains("../comments/comments1.xml"),
        "sheet rels target comments1: {sheet_rels}"
    );
    assert!(
        sheet_rels.contains("../drawings/vmlDrawing1.vml"),
        "sheet rels target vmlDrawing1: {sheet_rels}"
    );

    // 5a. threadedComments payload.
    let tc = read_zip_part(&bytes, "xl/threadedComments/threadedComments1.xml")
        .expect("tc payload present");
    assert!(tc.contains("<text>Looks wrong</text>"), "{tc}");
    assert!(tc.contains("<text>Agreed; investigating</text>"), "{tc}");
    assert!(tc.contains("id=\"{T-PARENT}\""), "{tc}");
    assert!(tc.contains("id=\"{T-REPLY}\""), "{tc}");
    assert!(tc.contains("parentId=\"{T-PARENT}\""), "{tc}");
    assert!(tc.contains("personId=\"{P-ALICE}\""), "{tc}");

    // 5b. personList payload.
    let pl = read_zip_part(&bytes, "xl/persons/personList.xml").expect("personList present");
    assert!(pl.contains("displayName=\"Alice\""), "{pl}");
    assert!(pl.contains("id=\"{P-ALICE}\""), "{pl}");
    assert!(pl.contains("userId=\"alice@example.com\""), "{pl}");

    // 5c. Legacy placeholder + synthetic author.
    let legacy =
        read_zip_part(&bytes, "xl/comments/comments1.xml").expect("legacy placeholder present");
    assert!(
        legacy.contains("<author>tc={T-PARENT}</author>"),
        "synthetic tc= author: {legacy}"
    );
    assert!(
        legacy.contains("<t>[Threaded comment]</t>"),
        "placeholder body: {legacy}"
    );
    // No second placeholder for the reply — the reply shares the parent's
    // anchor.
    let placeholder_count = legacy.matches("[Threaded comment]").count();
    assert_eq!(placeholder_count, 1, "exactly one placeholder: {legacy}");
}

#[test]
fn workbook_with_no_threaded_comments_omits_new_parts() {
    let mut wb = Workbook::new();
    wb.add_sheet(Worksheet::new("Plain"));

    let bytes = emit_xlsx(&mut wb);
    let paths = list_zip_paths(&bytes);

    assert!(!paths.contains("xl/persons/personList.xml"));
    assert!(
        !paths.iter().any(|p| p.starts_with("xl/threadedComments/")),
        "no threadedComments: {paths:?}"
    );

    let ct = read_zip_part(&bytes, "[Content_Types].xml").unwrap();
    assert!(!ct.contains("/xl/persons/personList.xml"));
    assert!(!ct.contains("/xl/threadedComments/"));
    assert!(!ct.contains("threadedcomments+xml"));
}
