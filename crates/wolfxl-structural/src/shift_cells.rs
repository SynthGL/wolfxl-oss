//! Rewrite a worksheet XML's cell coordinates and row indices.
//!
//! Streams `xl/worksheets/sheet*.xml` via `quick-xml`, copying every
//! byte verbatim except for:
//!   - `<row r="">` — shift the row attr (or drop the row entirely if
//!     it falls inside a delete band).
//!   - `<c r="">` — shift the cell coordinate (or drop the cell if it
//!     falls inside a delete band).
//!   - `<dimension ref="">` — rewrite via `shift_anchor`.
//!   - `<mergeCell ref="">` — rewrite via `shift_anchor`. Tombstoned
//!     merges are dropped (the surrounding `<mergeCells count>` is
//!     re-counted).
//!   - `<f>` text content — rewrite via `shift_formula`.
//!   - `<dataValidation sqref>` — rewrite via `shift_sqref`. Empty-sqref
//!     dataValidations are dropped.
//!   - `<dataValidation>/<formula1>` and `<formula2>` text — via
//!     `shift_formula`.
//!   - `<conditionalFormatting sqref>` — rewrite via `shift_sqref`.
//!     Empty-sqref CF blocks are dropped.
//!   - `<cfRule>/<formula>` text — via `shift_formula`.
//!   - `<xm:sqref>` extension text — rewrite via `shift_sqref`.
//!   - `<hyperlink ref>` — via `shift_anchor`.

use std::io::Cursor;

use quick_xml::events::{BytesStart, BytesText, Event};
use quick_xml::Reader as XmlReader;
use quick_xml::Writer as XmlWriter;

use crate::axis::ShiftPlan;
use crate::shift_anchors::{shift_anchor, shift_sqref};
use crate::shift_formulas::shift_formula;
use crate::xml_text::TextNode;

/// Push a resolved attribute value; quick-xml escapes the string tuple.
fn push_attr<'a>(e: &mut BytesStart<'a>, key: &str, val: &str) {
    e.push_attribute((key, val));
}

/// Streaming rewrite of a sheet XML. Returns the new bytes.
pub fn shift_sheet_cells(xml: &[u8], plan: &ShiftPlan) -> Vec<u8> {
    if plan.is_noop() {
        return xml.to_vec();
    }

    let xml_str = match std::str::from_utf8(xml) {
        Ok(s) => s,
        Err(_) => return xml.to_vec(),
    };
    let mut reader = XmlReader::from_str(xml_str);
    let mut writer = XmlWriter::new(Cursor::new(Vec::new()));
    let mut buf: Vec<u8> = Vec::new();

    // State: whether we're skipping the current element (drop tombstoned).
    let mut skip_depth: u32 = 0;
    // Track whether we're inside a tag whose text content is a formula.
    let mut in_f: bool = false;
    let mut in_formula1: bool = false;
    let mut in_formula2: bool = false;
    let mut in_sqref_text: bool = false;
    // Entity references split a logical formula/sqref into several events.
    let mut text = TextNode::default();

    loop {
        let event = reader.read_event_into(&mut buf);
        if !text.is_empty()
            && !matches!(event, Ok(Event::Text(_)) | Ok(Event::GeneralRef(_)))
        {
            if let Ok(value) = text.unescaped() {
                let shifted = if in_sqref_text {
                    shift_sqref(value, plan)
                } else {
                    shift_formula(value, plan)
                };
                let _ = writer.write_event(Event::Text(BytesText::new(&shifted)));
            } else {
                let _ = writer.write_event(Event::Text(BytesText::from_escaped(text.raw())));
            }
            text.clear();
        }
        match event {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                if skip_depth > 0 {
                    skip_depth += 1;
                    buf.clear();
                    continue;
                }

                match local.as_ref().as_bytes() {
                    b"row" => match handle_row_start(e, plan) {
                        RowAction::Keep(start) => {
                            let _ = writer.write_event(Event::Start(start));
                        }
                        RowAction::Drop => {
                            skip_depth = 1;
                        }
                    },
                    b"c" => match handle_cell_start(e, plan) {
                        CellAction::Keep(start) => {
                            let _ = writer.write_event(Event::Start(start));
                        }
                        CellAction::Drop => {
                            skip_depth = 1;
                        }
                    },
                    b"f" => {
                        in_f = true;
                        let _ = writer.write_event(Event::Start(e.to_owned()));
                    }
                    b"formula1" => {
                        in_formula1 = true;
                        let _ = writer.write_event(Event::Start(e.to_owned()));
                    }
                    b"formula2" => {
                        in_formula2 = true;
                        let _ = writer.write_event(Event::Start(e.to_owned()));
                    }
                    b"sqref" => {
                        in_sqref_text = true;
                        let _ = writer.write_event(Event::Start(e.to_owned()));
                    }
                    b"dimension" | b"mergeCell" | b"hyperlink" => {
                        let new_e = rewrite_ref_attr(e, plan, "ref", true);
                        match new_e {
                            Some(start) => {
                                let _ = writer.write_event(Event::Start(start));
                            }
                            None => {
                                skip_depth = 1;
                            }
                        }
                    }
                    b"dataValidation" => {
                        let new_e = rewrite_ref_attr(e, plan, "sqref", false);
                        match new_e {
                            Some(start) => {
                                let _ = writer.write_event(Event::Start(start));
                            }
                            None => {
                                skip_depth = 1;
                            }
                        }
                    }
                    b"conditionalFormatting" => {
                        let new_e = rewrite_ref_attr(e, plan, "sqref", false);
                        match new_e {
                            Some(start) => {
                                let _ = writer.write_event(Event::Start(start));
                            }
                            None => {
                                skip_depth = 1;
                            }
                        }
                    }
                    _ => {
                        let _ = writer.write_event(Event::Start(e.to_owned()));
                    }
                }
            }
            Ok(Event::Empty(ref e)) => {
                let local = e.local_name();
                if skip_depth > 0 {
                    buf.clear();
                    continue;
                }
                match local.as_ref().as_bytes() {
                    b"row" => match handle_row_start(e, plan) {
                        RowAction::Keep(start) => {
                            let _ = writer.write_event(Event::Empty(start));
                        }
                        RowAction::Drop => {}
                    },
                    b"c" => match handle_cell_start(e, plan) {
                        CellAction::Keep(start) => {
                            let _ = writer.write_event(Event::Empty(start));
                        }
                        CellAction::Drop => {}
                    },
                    b"dimension" | b"mergeCell" | b"hyperlink" => {
                        let new_e = rewrite_ref_attr(e, plan, "ref", true);
                        if let Some(start) = new_e {
                            let _ = writer.write_event(Event::Empty(start));
                        }
                    }
                    b"dataValidation" => {
                        let new_e = rewrite_ref_attr(e, plan, "sqref", false);
                        if let Some(start) = new_e {
                            let _ = writer.write_event(Event::Empty(start));
                        }
                    }
                    b"conditionalFormatting" => {
                        let new_e = rewrite_ref_attr(e, plan, "sqref", false);
                        if let Some(start) = new_e {
                            let _ = writer.write_event(Event::Empty(start));
                        }
                    }
                    _ => {
                        let _ = writer.write_event(Event::Empty(e.to_owned()));
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                let local = e.local_name();
                if skip_depth > 0 {
                    skip_depth -= 1;
                    buf.clear();
                    continue;
                }
                match local.as_ref().as_bytes() {
                    b"f" => in_f = false,
                    b"formula1" => in_formula1 = false,
                    b"formula2" => in_formula2 = false,
                    b"sqref" => in_sqref_text = false,
                    _ => {}
                }
                let _ = writer.write_event(Event::End(e.to_owned()));
            }
            Ok(Event::Text(t))
                if skip_depth == 0 && (in_f || in_formula1 || in_formula2 || in_sqref_text) =>
            {
                text.push_text(&t);
            }
            Ok(Event::GeneralRef(r))
                if skip_depth == 0 && (in_f || in_formula1 || in_formula2 || in_sqref_text) =>
            {
                text.push_ref(&r);
            }
            Ok(Event::Eof) => break,
            Ok(other) => {
                if skip_depth > 0 {
                    buf.clear();
                    continue;
                }
                let _ = writer.write_event(other);
            }
            Err(_) => break,
        }
        buf.clear();
    }

    cleanup_empty_sheet_containers(writer.into_inner().into_inner())
}

fn cleanup_empty_sheet_containers(bytes: Vec<u8>) -> Vec<u8> {
    let Ok(s) = String::from_utf8(bytes) else {
        unreachable!("quick-xml writer emitted non-UTF-8 worksheet XML");
    };
    remove_empty_elements_by_local_name(&s, "mergeCells").into_bytes()
}

fn remove_empty_elements_by_local_name(s: &str, local_name: &str) -> String {
    let mut out = String::with_capacity(s.len());
    let mut pos = 0;
    while let Some(rel_start) = s[pos..].find('<') {
        let start = pos + rel_start;
        out.push_str(&s[pos..start]);

        let Some(rel_end) = s[start..].find('>') else {
            out.push_str(&s[start..]);
            return out;
        };
        let open_end = start + rel_end;
        let tag_body = s[start + 1..open_end].trim_start();
        if tag_body.starts_with(['/', '?', '!']) {
            out.push_str(&s[start..=open_end]);
            pos = open_end + 1;
            continue;
        }

        let tag_name_end = tag_body
            .find(|c: char| c.is_whitespace() || c == '/')
            .unwrap_or(tag_body.len());
        let tag_name = &tag_body[..tag_name_end];
        let tag_local = tag_name.rsplit(':').next().unwrap_or(tag_name);
        if tag_local != local_name {
            out.push_str(&s[start..=open_end]);
            pos = open_end + 1;
            continue;
        }

        if tag_body.trim_end().ends_with('/') {
            pos = open_end + 1;
            continue;
        }

        let close_tag = format!("</{tag_name}>");
        let Some(close_rel) = s[open_end + 1..].find(&close_tag) else {
            out.push_str(&s[start..=open_end]);
            pos = open_end + 1;
            continue;
        };
        let close_start = open_end + 1 + close_rel;
        if s[open_end + 1..close_start].trim().is_empty() {
            pos = close_start + close_tag.len();
        } else {
            out.push_str(&s[start..=open_end]);
            pos = open_end + 1;
        }
    }
    out.push_str(&s[pos..]);
    out
}

enum RowAction<'a> {
    Keep(BytesStart<'a>),
    Drop,
}

enum CellAction<'a> {
    Keep(BytesStart<'a>),
    Drop,
}

/// Decide what to do with a `<row r="N">` element. If the row falls
/// inside a delete band, drop it. If on a row shift, rewrite `r`. If
/// on a col shift, leave `r` alone (but children's cell coords still
/// need col-shifting, handled by `<c>`).
fn handle_row_start<'a>(e: &BytesStart<'a>, plan: &ShiftPlan) -> RowAction<'a> {
    let mut new_e = BytesStart::new(e.name().as_ref().to_owned());
    let mut row_n: Option<u32> = None;
    let mut row_attr_kept = false;

    for attr_res in e.attributes().with_checks(false) {
        let Ok(attr) = attr_res else { continue };
        let key = attr.key.as_ref();
        let val = match quick_xml::escape::unescape(&attr.value) {
            Ok(v) => v.into_owned(),
            Err(_) => continue,
        };
        if key == "r" && plan.axis.is_row() {
            let n: u32 = match val.parse() {
                Ok(n) => n,
                Err(_) => {
                    push_attr(&mut new_e, key, &val);
                    row_attr_kept = true;
                    continue;
                }
            };
            row_n = Some(n);
            // Decide.
            if plan.is_insert() {
                let new_n = if n >= plan.idx {
                    n as i64 + plan.n as i64
                } else {
                    n as i64
                };
                if new_n < 1 || new_n > crate::MAX_ROW as i64 {
                    return RowAction::Drop;
                }
                push_attr(&mut new_e, key, &(new_n as u32).to_string());
                row_attr_kept = true;
            } else {
                // delete
                let abs = plan.abs_n();
                if n >= plan.idx && n < plan.idx + abs {
                    return RowAction::Drop;
                }
                let new_n = if n >= plan.idx + abs {
                    n as i64 + plan.n as i64
                } else {
                    n as i64
                };
                if new_n < 1 {
                    return RowAction::Drop;
                }
                push_attr(&mut new_e, key, &(new_n as u32).to_string());
                row_attr_kept = true;
            }
        } else if key == "spans" && plan.axis.is_col() {
            // spans is "min:max" col indices for the row. Not strictly
            // required for correctness — Excel re-derives — but keep
            // it consistent. We pass through unchanged for now; the
            // spans format isn't critical.
            push_attr(&mut new_e, key, &val);
        } else {
            push_attr(&mut new_e, key, &val);
            if key == "r" {
                row_attr_kept = true;
            }
        }
    }
    let _ = (row_n, row_attr_kept);
    RowAction::Keep(new_e)
}

/// Rewrite a `<c r="...">` cell start. If row/col is inside a delete
/// band, drop. Else shift the coordinate.
fn handle_cell_start<'a>(e: &BytesStart<'a>, plan: &ShiftPlan) -> CellAction<'a> {
    let mut new_e = BytesStart::new(e.name().as_ref().to_owned());
    for attr_res in e.attributes().with_checks(false) {
        let Ok(attr) = attr_res else { continue };
        let key = attr.key.as_ref();
        let val = match quick_xml::escape::unescape(&attr.value) {
            Ok(v) => v.into_owned(),
            Err(_) => continue,
        };
        if key == "r" {
            let new_ref = shift_anchor(&val, plan);
            if new_ref == "#REF!" {
                return CellAction::Drop;
            }
            push_attr(&mut new_e, key, &new_ref);
        } else {
            push_attr(&mut new_e, key, &val);
        }
    }
    CellAction::Keep(new_e)
}

/// Rewrite an attribute (`ref` or `sqref`) on a generic element.
/// `is_ref` chooses between `shift_anchor` (single ref) and
/// `shift_sqref` (multi-range). Returns None when the resulting
/// attribute would be `#REF!` (for `ref`) or empty (for `sqref`).
fn rewrite_ref_attr<'a>(
    e: &BytesStart<'a>,
    plan: &ShiftPlan,
    attr_name: &str,
    is_ref: bool,
) -> Option<BytesStart<'a>> {
    let mut new_e = BytesStart::new(e.name().as_ref().to_owned());
    let mut found_match = false;
    let mut keep = true;
    for attr_res in e.attributes().with_checks(false) {
        let Ok(attr) = attr_res else { continue };
        let key = attr.key.as_ref();
        let val = match quick_xml::escape::unescape(&attr.value) {
            Ok(v) => v.into_owned(),
            Err(_) => continue,
        };
        if key == attr_name {
            found_match = true;
            let new_val = if is_ref {
                shift_anchor(&val, plan)
            } else {
                shift_sqref(&val, plan)
            };
            if is_ref {
                if new_val == "#REF!" {
                    keep = false;
                }
            } else if new_val.is_empty() {
                keep = false;
            }
            push_attr(&mut new_e, key, &new_val);
        } else {
            push_attr(&mut new_e, key, &val);
        }
    }
    if !found_match || !keep {
        if !found_match {
            return Some(new_e);
        }
        return None;
    }
    Some(new_e)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn apply(xml: &str, plan: &ShiftPlan) -> String {
        String::from_utf8(shift_sheet_cells(xml.as_bytes(), plan)).unwrap()
    }

    #[test]
    fn shifts_entity_split_formulas_and_sqrefs_without_touching_literal_text() {
        let xml = r#"<worksheet><sheetData><row r="1"><c r="A1" note="A&amp;B"><f>A&#49;&amp;&quot;A1&quot;</f><is><t> &#65;&amp;&lt; </t></is></c></row></sheetData><dataValidations><dataValidation sqref="A1"><formula1>A&#49;</formula1><formula2>A&#x31;&lt;B1</formula2></dataValidation></dataValidations><ext><sqref>A&#49; B&#x31;</sqref></ext></worksheet>"#;
        let out = apply(xml, &ShiftPlan::insert(crate::Axis::Row, 1, 1));
        assert!(out.contains("<f>A2&amp;&quot;A1&quot;</f>"), "{out}");
        assert!(out.contains("<formula1>A2</formula1>"), "{out}");
        assert!(out.contains("<formula2>A2&lt;B2</formula2>"), "{out}");
        assert!(out.contains("<sqref>A2 B2</sqref>"), "{out}");
        assert!(out.contains(r#"note="A&amp;B""#), "{out}");
        assert!(out.contains("<t> &#65;&amp;&lt; </t>"), "{out}");
    }

    #[test]
    fn deleted_rows_do_not_leak_entity_references() {
        let xml = r#"<sheetData><row r="1"><c r="A1"><f>&quot;gone&quot;&amp;A1</f></c></row><row r="2"><c r="A2"><f>A&#50;</f></c></row></sheetData>"#;
        let out = apply(xml, &ShiftPlan::delete(crate::Axis::Row, 1, 1));
        assert_eq!(out, r#"<sheetData><row r="1"><c r="A1"><f>A1</f></c></row></sheetData>"#);
    }

    #[test]
    fn shifts_cell_coordinates() {
        let xml = r#"<sheetData><row r="5"><c r="A5"><v>1</v></c></row></sheetData>"#;
        let p = ShiftPlan::insert(crate::Axis::Row, 5, 3);
        let out = apply(xml, &p);
        assert!(out.contains(r#"<row r="8">"#));
        assert!(out.contains(r#"<c r="A8">"#));
    }

    #[test]
    fn shifts_dimension() {
        let xml = r#"<dimension ref="A1:B10"/><sheetData/>"#;
        let p = ShiftPlan::insert(crate::Axis::Row, 5, 3);
        let out = apply(xml, &p);
        assert!(out.contains(r#"ref="A1:B13""#));
    }

    #[test]
    fn shifts_merge_cell() {
        let xml = r#"<mergeCells count="1"><mergeCell ref="A5:B7"/></mergeCells>"#;
        let p = ShiftPlan::insert(crate::Axis::Row, 5, 3);
        let out = apply(xml, &p);
        assert!(out.contains(r#"ref="A8:B10""#));
    }

    #[test]
    fn drops_tombstoned_merge_cell() {
        let xml = r#"<mergeCells count="1"><mergeCell ref="A5:B7"/></mergeCells>"#;
        let p = ShiftPlan::delete(crate::Axis::Row, 5, 3);
        let out = apply(xml, &p);
        assert!(!out.contains("mergeCell ref"));
        assert!(!out.contains("mergeCells"));
    }

    #[test]
    fn drops_empty_prefixed_merge_cells_container_with_stale_count() {
        let xml = r#"<x:mergeCells count="1" xmlns:x="urn:test"><x:mergeCell ref="A1:B1"/></x:mergeCells>"#;
        let p = ShiftPlan::delete(crate::Axis::Row, 1, 1);
        let out = apply(xml, &p);
        assert!(!out.contains("mergeCell ref"), "{out}");
        assert!(!out.contains("mergeCells"), "{out}");
    }

    #[test]
    fn shifts_formula_text() {
        let xml = r#"<c r="A5"><f>SUM(A1:A4)</f><v>10</v></c>"#;
        let p = ShiftPlan::insert(crate::Axis::Row, 5, 3);
        let out = apply(xml, &p);
        assert!(out.contains(r#"<c r="A8">"#));
        assert!(out.contains("<f>SUM(A1:A4)</f>"));
    }

    #[test]
    fn shifts_formula_into_band() {
        let xml = r#"<c r="A10"><f>A5</f></c>"#;
        let p = ShiftPlan::insert(crate::Axis::Row, 5, 3);
        let out = apply(xml, &p);
        assert!(out.contains("<f>A8</f>"));
    }

    #[test]
    fn drops_cell_in_delete_band() {
        // Row 5 (in delete band) holds value "x"; row 8 (outside band)
        // holds value "y". After delete(5, 3) the band-row drops
        // entirely, and row 8 → row 5 with cell value "y".
        let xml =
            r#"<row r="5"><c r="A5"><v>x</v></c></row><row r="8"><c r="A8"><v>y</v></c></row>"#;
        let p = ShiftPlan::delete(crate::Axis::Row, 5, 3);
        let out = apply(xml, &p);
        // Original band-row content "x" is gone; surviving content "y"
        // remains. (`r="A5"` itself appears in the output because
        // shifted A8 → A5 — that's correct, not a regression.)
        assert!(!out.contains("<v>x</v>"));
        assert!(out.contains("<v>y</v>"));
        // Row 8 → row 5 after delete.
        assert!(out.contains(r#"<row r="5">"#));
        assert!(out.contains(r#"<c r="A5">"#));
    }

    #[test]
    fn shifts_hyperlink_ref() {
        let xml = r#"<hyperlinks><hyperlink ref="B5" r:id="rId1"/></hyperlinks>"#;
        let p = ShiftPlan::insert(crate::Axis::Row, 5, 3);
        let out = apply(xml, &p);
        assert!(out.contains(r#"ref="B8""#));
    }

    #[test]
    fn hyperlink_attrs_remain_xml_escaped_after_axis_shift() {
        let xml = r#"<hyperlinks><hyperlink ref="B5" location="'Sec. 1 &amp; 2 Notes'!A1" display="A &amp; B"/></hyperlinks>"#;
        let p = ShiftPlan::insert(crate::Axis::Row, 5, 3);
        let out = apply(xml, &p);
        assert!(out.contains(r#"ref="B8""#));
        assert!(
            out.contains(r#"location="&apos;Sec. 1 &amp; 2 Notes&apos;!A1""#),
            "got: {out}"
        );
        assert!(out.contains(r#"display="A &amp; B""#), "got: {out}");
    }

    #[test]
    fn shifts_dv_sqref_and_formula() {
        let xml = r#"<dataValidations><dataValidation type="list" sqref="A5:A10"><formula1>$Z$5:$Z$10</formula1></dataValidation></dataValidations>"#;
        let p = ShiftPlan::insert(crate::Axis::Row, 5, 3);
        let out = apply(xml, &p);
        assert!(out.contains(r#"sqref="A8:A13""#));
        assert!(out.contains("<formula1>$Z$8:$Z$13</formula1>"));
    }

    #[test]
    fn shifts_cf_sqref_and_formula() {
        let xml = r#"<conditionalFormatting sqref="A5:A10"><cfRule type="cellIs"><formula>5</formula></cfRule></conditionalFormatting>"#;
        let p = ShiftPlan::insert(crate::Axis::Row, 5, 3);
        let out = apply(xml, &p);
        assert!(out.contains(r#"sqref="A8:A13""#));
    }

    #[test]
    fn shifts_extension_sqref_text() {
        let xml = r#"<worksheet><extLst><ext><x14:conditionalFormattings xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><xm:sqref>C2:C4</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst></worksheet>"#;
        let p = ShiftPlan::delete(crate::Axis::Row, 1, 1);
        let out = apply(xml, &p);
        assert!(out.contains("<xm:sqref>C1:C3</xm:sqref>"), "{out}");
    }

    #[test]
    fn drops_dv_when_sqref_empty() {
        let xml = r#"<dataValidations><dataValidation type="list" sqref="A5"><formula1>1</formula1></dataValidation></dataValidations>"#;
        let p = ShiftPlan::delete(crate::Axis::Row, 5, 1);
        let out = apply(xml, &p);
        // The inner <dataValidation> (singular, with sqref) is dropped
        // because the only sqref tombstoned to empty. The outer
        // <dataValidations> wrapper remains as an empty container —
        // OOXML accepts that.
        assert!(!out.contains("<dataValidation "));
        assert!(!out.contains("formula1"));
    }

    #[test]
    fn passes_through_unrelated_elements() {
        let xml = r#"<sheetPr><pageSetup orientation="portrait"/></sheetPr>"#;
        let p = ShiftPlan::insert(crate::Axis::Row, 5, 3);
        let out = apply(xml, &p);
        assert!(out.contains("sheetPr"));
        assert!(out.contains("pageSetup"));
    }

    #[test]
    fn col_shift_passes_row_attr_unchanged() {
        let xml = r#"<row r="5"><c r="B5"><v>1</v></c></row>"#;
        let p = ShiftPlan::insert(crate::Axis::Col, 2, 1);
        let out = apply(xml, &p);
        assert!(out.contains(r#"<row r="5">"#)); // row attr unchanged
        assert!(out.contains(r#"<c r="C5">"#)); // cell shifted
    }

    #[test]
    fn noop_returns_input_bytes() {
        let xml = r#"<sheetData><row r="5"><c r="A5"/></row></sheetData>"#;
        let p = ShiftPlan {
            axis: crate::Axis::Row,
            idx: 1,
            n: 0,
        };
        let out = apply(xml, &p);
        assert_eq!(out, xml);
    }
}
