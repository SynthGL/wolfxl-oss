//! Top-level orchestrator that walks every workbook part touched by
//! a structural shift and emits the rewritten bytes.
//!
//! Designed to be called once per `(sheet, axis, idx, n)` op with all
//! the workbook's relevant XML parts loaded into the inputs struct.
//! The orchestrator returns a `WorkbookMutations` map of `path → new
//! bytes` that the patcher can fold into its `file_patches` map.
//!
//! Multi-op sequencing is the caller's responsibility — apply each
//! op individually, re-read the rewritten bytes for the next op, then
//! call again.

use std::collections::BTreeMap;

use quick_xml::events::{BytesStart, BytesText, Event};
use quick_xml::Reader as XmlReader;
use quick_xml::Writer as XmlWriter;

use crate::axis::{Axis, ShiftPlan};
use crate::control_props_shift::shift_control_props_xml;
use crate::drawing_shift::shift_drawing_xml;
use crate::shift_anchors::shift_anchor;
use crate::shift_cells::shift_sheet_cells;
use crate::shift_formulas::shift_formula_on_sheet;
use crate::table_shift::{repair_deleted_table_header_row, shift_table_xml};
use crate::vml_shift::shift_vml_xml;
use crate::xml_text::TextNode;

fn push_attr<'a>(e: &mut BytesStart<'a>, key: &str, val: &str) {
    e.push_attribute((key, val));
}

/// One queued structural-shift op, as supplied by the patcher's
/// `queued_axis_shifts` queue.
#[derive(Debug, Clone)]
pub struct AxisShiftOp {
    /// Sheet name (NOT path).
    pub sheet: String,
    /// Row or col axis.
    pub axis: Axis,
    /// 1-based index where shifting begins.
    pub idx: u32,
    /// Signed shift count (positive = insert; negative = delete).
    pub n: i32,
}

impl AxisShiftOp {
    fn plan(&self) -> ShiftPlan {
        ShiftPlan {
            axis: self.axis,
            idx: self.idx,
            n: self.n,
        }
    }
}

/// All XML/VML parts the orchestrator may need to rewrite, plus a
/// `sheet_positions` map for defined-name `localSheetId` resolution.
pub struct SheetXmlInputs<'a> {
    /// Sheet name → sheet XML bytes (e.g. `xl/worksheets/sheet1.xml`).
    pub sheets: BTreeMap<String, &'a [u8]>,
    /// Sheet name → sheet XML path (so output keys match source paths).
    pub sheet_paths: BTreeMap<String, String>,
    /// Optional `xl/workbook.xml` bytes (for defined-name shift).
    pub workbook_xml: Option<&'a [u8]>,
    /// Parsed `xl/sharedStrings.xml` values, used when row deletes promote a
    /// table data row into the header row.
    pub shared_strings: Option<&'a [String]>,
    /// Per-sheet table parts: sheet name → vec of (path, bytes).
    pub tables: BTreeMap<String, Vec<(String, &'a [u8])>>,
    /// Per-sheet comments part: sheet name → (path, bytes).
    pub comments: BTreeMap<String, (String, &'a [u8])>,
    /// Per-sheet vmlDrawing part: sheet name → (path, bytes).
    pub vml: BTreeMap<String, (String, &'a [u8])>,
    /// Per-sheet DrawingML part: sheet name → (path, bytes).
    pub drawings: BTreeMap<String, (String, &'a [u8])>,
    /// Per-sheet form-control property parts: sheet name → vec of (path, bytes).
    pub control_props: BTreeMap<String, Vec<(String, &'a [u8])>>,
    /// Sheet name → 0-based position (for definedName localSheetId).
    pub sheet_positions: BTreeMap<String, u32>,
}

impl<'a> SheetXmlInputs<'a> {
    /// Empty inputs. Useful as a starting point for callers that
    /// build the map field-by-field.
    pub fn empty() -> Self {
        Self {
            sheets: BTreeMap::new(),
            sheet_paths: BTreeMap::new(),
            workbook_xml: None,
            shared_strings: None,
            tables: BTreeMap::new(),
            comments: BTreeMap::new(),
            vml: BTreeMap::new(),
            drawings: BTreeMap::new(),
            control_props: BTreeMap::new(),
            sheet_positions: BTreeMap::new(),
        }
    }
}

/// Output of `apply_workbook_shift`: `path → new bytes`.
#[derive(Debug, Default)]
pub struct WorkbookMutations {
    /// Path → new bytes.
    pub file_patches: BTreeMap<String, Vec<u8>>,
}

/// Apply one structural shift op across every part in `inputs`.
///
/// No-op invariant: an empty `ops` slice or an op with `n == 0`
/// returns an empty `file_patches` map. Callers MUST short-circuit
/// before calling if they want byte-identical output for an empty
/// queue.
pub fn apply_workbook_shift(inputs: SheetXmlInputs<'_>, ops: &[AxisShiftOp]) -> WorkbookMutations {
    let mut out = WorkbookMutations::default();

    if ops.is_empty() {
        return out;
    }

    // Cache rewritten sheet bytes across ops in this call so a single
    // multi-op call sees its own rewrites. (In practice the patcher
    // calls one op at a time, but this future-proofs the API.)
    let mut sheet_bytes: BTreeMap<String, Vec<u8>> = inputs
        .sheets
        .iter()
        .map(|(k, v)| (k.clone(), v.to_vec()))
        .collect();
    let mut workbook_bytes: Option<Vec<u8>> = inputs.workbook_xml.map(|b| b.to_vec());
    let mut table_bytes: BTreeMap<String, BTreeMap<String, Vec<u8>>> = inputs
        .tables
        .iter()
        .map(|(sheet, parts)| {
            (
                sheet.clone(),
                parts.iter().map(|(p, b)| (p.clone(), b.to_vec())).collect(),
            )
        })
        .collect();
    let mut comments_bytes: BTreeMap<String, (String, Vec<u8>)> = inputs
        .comments
        .iter()
        .map(|(sheet, (path, b))| (sheet.clone(), (path.clone(), b.to_vec())))
        .collect();
    let mut vml_bytes: BTreeMap<String, (String, Vec<u8>)> = inputs
        .vml
        .iter()
        .map(|(sheet, (path, b))| (sheet.clone(), (path.clone(), b.to_vec())))
        .collect();
    let mut drawing_bytes: BTreeMap<String, (String, Vec<u8>)> = inputs
        .drawings
        .iter()
        .map(|(sheet, (path, b))| (sheet.clone(), (path.clone(), b.to_vec())))
        .collect();
    let mut control_prop_bytes: BTreeMap<String, BTreeMap<String, Vec<u8>>> = inputs
        .control_props
        .iter()
        .map(|(sheet, parts)| {
            (
                sheet.clone(),
                parts.iter().map(|(p, b)| (p.clone(), b.to_vec())).collect(),
            )
        })
        .collect();

    for op in ops {
        let plan = op.plan();
        if plan.is_noop() {
            continue;
        }

        // 1. Sheet XML.
        if let Some(bytes) = sheet_bytes.get(&op.sheet) {
            let new_bytes = shift_sheet_cells(bytes, &plan);
            // 1b. RFC-031: rewrite the `<cols>` block on Col-axis shifts.
            let new_bytes = crate::cols::shift_sheet_cols_block(&new_bytes, plan);
            sheet_bytes.insert(op.sheet.clone(), new_bytes);
        }

        // 2. Tables on this sheet.
        if let Some(parts) = table_bytes.get_mut(&op.sheet) {
            let mut updated: BTreeMap<String, Vec<u8>> = BTreeMap::new();
            for (path, bytes) in parts.iter() {
                let new_bytes = shift_table_xml(bytes, &plan);
                updated.insert(path.clone(), new_bytes);
            }
            *parts = updated;
        }

        // 2b. If a row delete removed a structured-table header row, keep the
        // table metadata stable and mark the table headerless. Promoting the
        // first data row into header metadata leaves Excel's PDF export path
        // rejecting the workbook.
        if plan.axis == Axis::Row && plan.is_delete() {
            if let (Some(sheet), Some(parts)) = (
                sheet_bytes.get_mut(&op.sheet),
                table_bytes.get_mut(&op.sheet),
            ) {
                for (_path, table) in parts.iter_mut() {
                    if let Some(new_sheet) = repair_deleted_table_header_row(
                        sheet,
                        table,
                        inputs.shared_strings.unwrap_or(&[]),
                        &plan,
                    ) {
                        *sheet = new_sheet;
                    }
                }
            }
        }

        // 3. Comments on this sheet.
        if let Some((path, bytes)) = comments_bytes.get(&op.sheet) {
            let new_bytes = shift_comments_xml(bytes, &plan);
            comments_bytes.insert(op.sheet.clone(), (path.clone(), new_bytes));
        }

        // 4. VML on this sheet.
        if let Some((path, bytes)) = vml_bytes.get(&op.sheet) {
            let new_bytes = shift_vml_xml(bytes, &plan);
            vml_bytes.insert(op.sheet.clone(), (path.clone(), new_bytes));
        }

        // 5. DrawingML anchors on this sheet.
        if let Some((path, bytes)) = drawing_bytes.get(&op.sheet) {
            let new_bytes = shift_drawing_xml(bytes, &plan);
            drawing_bytes.insert(op.sheet.clone(), (path.clone(), new_bytes));
        }

        // 6. Form-control properties on this sheet.
        if let Some(parts) = control_prop_bytes.get_mut(&op.sheet) {
            let mut updated: BTreeMap<String, Vec<u8>> = BTreeMap::new();
            for (path, bytes) in parts.iter() {
                let new_bytes = shift_control_props_xml(bytes, &plan);
                updated.insert(path.clone(), new_bytes);
            }
            *parts = updated;
        }

        // 7. Workbook defined names.
        if let Some(ref wb) = workbook_bytes {
            let new_bytes = shift_defined_names(
                wb,
                &plan,
                &op.sheet,
                inputs.sheet_positions.get(&op.sheet).copied(),
            );
            workbook_bytes = Some(new_bytes);
        }
    }

    // Emit file_patches.
    for (sheet, bytes) in &sheet_bytes {
        if let Some(orig) = inputs.sheets.get(sheet) {
            if bytes.as_slice() != *orig {
                if let Some(path) = inputs.sheet_paths.get(sheet) {
                    out.file_patches.insert(path.clone(), bytes.clone());
                }
            }
        }
    }
    if let Some(bytes) = &workbook_bytes {
        if let Some(orig) = inputs.workbook_xml {
            if bytes.as_slice() != orig {
                out.file_patches
                    .insert("xl/workbook.xml".to_string(), bytes.clone());
            }
        }
    }
    for (_sheet, parts) in &table_bytes {
        for (path, bytes) in parts {
            // Compare against original.
            let mut matched_orig = false;
            for (_s, parts_orig) in &inputs.tables {
                for (p_o, b_o) in parts_orig {
                    if p_o == path {
                        if bytes.as_slice() != *b_o {
                            out.file_patches.insert(path.clone(), bytes.clone());
                        }
                        matched_orig = true;
                        break;
                    }
                }
                if matched_orig {
                    break;
                }
            }
        }
    }
    for (sheet, (path, bytes)) in &comments_bytes {
        if let Some((_, orig)) = inputs.comments.get(sheet) {
            if bytes.as_slice() != *orig {
                out.file_patches.insert(path.clone(), bytes.clone());
            }
        }
    }
    for (sheet, (path, bytes)) in &vml_bytes {
        if let Some((_, orig)) = inputs.vml.get(sheet) {
            if bytes.as_slice() != *orig {
                out.file_patches.insert(path.clone(), bytes.clone());
            }
        }
    }
    for (sheet, (path, bytes)) in &drawing_bytes {
        if let Some((_, orig)) = inputs.drawings.get(sheet) {
            if bytes.as_slice() != *orig {
                out.file_patches.insert(path.clone(), bytes.clone());
            }
        }
    }
    for (_sheet, parts) in &control_prop_bytes {
        for (path, bytes) in parts {
            let mut matched_orig = false;
            for (_s, parts_orig) in &inputs.control_props {
                for (p_o, b_o) in parts_orig {
                    if p_o == path {
                        if bytes.as_slice() != *b_o {
                            out.file_patches.insert(path.clone(), bytes.clone());
                        }
                        matched_orig = true;
                        break;
                    }
                }
                if matched_orig {
                    break;
                }
            }
        }
    }

    out
}

/// Rewrite a `xl/comments*.xml` part: `<comment ref>`.
pub fn shift_comments_xml(xml: &[u8], plan: &ShiftPlan) -> Vec<u8> {
    if plan.is_noop() {
        return xml.to_vec();
    }
    let xml_str = match std::str::from_utf8(xml) {
        Ok(s) => s,
        Err(_) => return xml.to_vec(),
    };
    let mut reader = XmlReader::from_str(xml_str);
    let mut writer = XmlWriter::new(std::io::Cursor::new(Vec::new()));
    let mut buf: Vec<u8> = Vec::new();
    let mut skip_depth: u32 = 0;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                if skip_depth > 0 {
                    skip_depth += 1;
                    buf.clear();
                    continue;
                }
                if local.as_ref() == "comment" {
                    let mut keep = true;
                    let mut new_e = BytesStart::new(e.name().as_ref().to_owned());
                    for attr_res in e.attributes().with_checks(false) {
                        let Ok(attr) = attr_res else { continue };
                        let key = attr.key.as_ref();
                        let val = match quick_xml::escape::unescape(&attr.value) {
                            Ok(v) => v.into_owned(),
                            Err(_) => continue,
                        };
                        if key == "ref" {
                            let new_val = shift_anchor(&val, plan);
                            if new_val == "#REF!" {
                                keep = false;
                            }
                            push_attr(&mut new_e, key, &new_val);
                        } else {
                            push_attr(&mut new_e, key, &val);
                        }
                    }
                    if keep {
                        let _ = writer.write_event(Event::Start(new_e));
                    } else {
                        skip_depth = 1;
                    }
                } else {
                    let _ = writer.write_event(Event::Start(e.to_owned()));
                }
            }
            Ok(Event::Empty(ref e)) => {
                let local = e.local_name();
                if skip_depth > 0 {
                    buf.clear();
                    continue;
                }
                if local.as_ref() == "comment" {
                    let mut keep = true;
                    let mut new_e = BytesStart::new(e.name().as_ref().to_owned());
                    for attr_res in e.attributes().with_checks(false) {
                        let Ok(attr) = attr_res else { continue };
                        let key = attr.key.as_ref();
                        let val = match quick_xml::escape::unescape(&attr.value) {
                            Ok(v) => v.into_owned(),
                            Err(_) => continue,
                        };
                        if key == "ref" {
                            let new_val = shift_anchor(&val, plan);
                            if new_val == "#REF!" {
                                keep = false;
                            }
                            push_attr(&mut new_e, key, &new_val);
                        } else {
                            push_attr(&mut new_e, key, &val);
                        }
                    }
                    if keep {
                        let _ = writer.write_event(Event::Empty(new_e));
                    }
                } else {
                    let _ = writer.write_event(Event::Empty(e.to_owned()));
                }
            }
            Ok(Event::End(ref e)) => {
                if skip_depth > 0 {
                    skip_depth -= 1;
                    buf.clear();
                    continue;
                }
                let _ = writer.write_event(Event::End(e.to_owned()));
            }
            Ok(Event::Text(ref t)) => {
                if skip_depth > 0 {
                    buf.clear();
                    continue;
                }
                let _ = writer.write_event(Event::Text(t.to_owned()));
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

    writer.into_inner().into_inner()
}

/// Rewrite `<definedName>` text content inside `xl/workbook.xml`.
///
/// Workbook-scope names (no `localSheetId` attr) shift only refs that
/// target `op_sheet`. Per-sheet names (with `localSheetId="N"`) shift
/// only when `N == sheet_position[op_sheet]`.
pub fn shift_defined_names(
    xml: &[u8],
    plan: &ShiftPlan,
    op_sheet: &str,
    op_sheet_position: Option<u32>,
) -> Vec<u8> {
    if plan.is_noop() {
        return xml.to_vec();
    }
    let xml_str = match std::str::from_utf8(xml) {
        Ok(s) => s,
        Err(_) => return xml.to_vec(),
    };
    let mut reader = XmlReader::from_str(xml_str);
    let mut writer = XmlWriter::new(std::io::Cursor::new(Vec::new()));
    let mut buf: Vec<u8> = Vec::new();

    let mut in_dn = false;
    let mut current_local_sheet_id: Option<u32> = None;
    // Resolve split Text/GeneralRef events before translating a whole formula.
    let mut formula = TextNode::default();

    loop {
        let event = reader.read_event_into(&mut buf);
        if !formula.is_empty()
            && !matches!(event, Ok(Event::Text(_)) | Ok(Event::GeneralRef(_)))
        {
            let applies = match (current_local_sheet_id, op_sheet_position) {
                (Some(sid), Some(pos)) => sid == pos,
                (None, _) => true,
                (Some(_), None) => false,
            };
            if let Some(text) = formula.unescaped().ok().filter(|_| applies) {
                let shifted = shift_formula_on_sheet(text, plan, op_sheet);
                let _ = writer.write_event(Event::Text(BytesText::new(&shifted)));
            } else {
                let _ = writer.write_event(Event::Text(BytesText::from_escaped(formula.raw())));
            }
            formula.clear();
        }
        match event {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                if local.as_ref() == "definedName" {
                    in_dn = true;
                    current_local_sheet_id = None;
                    for attr_res in e.attributes().with_checks(false) {
                        let Ok(attr) = attr_res else { continue };
                        if attr.key.as_ref() == "localSheetId" {
                            if let Ok(v) = quick_xml::escape::unescape(&attr.value) {
                                if let Ok(n) = v.parse::<u32>() {
                                    current_local_sheet_id = Some(n);
                                }
                            }
                        }
                    }
                }
                let _ = writer.write_event(Event::Start(e.to_owned()));
            }
            Ok(Event::Empty(ref e)) => {
                let _ = writer.write_event(Event::Empty(e.to_owned()));
            }
            Ok(Event::End(ref e)) => {
                let local = e.local_name();
                if local.as_ref() == "definedName" {
                    in_dn = false;
                    current_local_sheet_id = None;
                }
                let _ = writer.write_event(Event::End(e.to_owned()));
            }
            Ok(Event::Text(t)) if in_dn => formula.push_text(&t),
            Ok(Event::GeneralRef(r)) if in_dn => formula.push_ref(&r),
            Ok(Event::Eof) => break,
            Ok(other) => {
                let _ = writer.write_event(other);
            }
            Err(_) => break,
        }
        buf.clear();
    }

    writer.into_inner().into_inner()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn defined_names_shift_whole_entity_split_formulas() {
        let xml = br#"<workbook><definedNames><definedName name="global">'A&amp;B'!A&#49;&amp;&quot;A1&quot;</definedName><definedName name="other" localSheetId="1">'A&amp;B'!A&#49;</definedName></definedNames></workbook>"#;
        let out = shift_defined_names(xml, &ShiftPlan::insert(Axis::Row, 1, 1), "A&B", Some(0));
        let text = String::from_utf8(out).unwrap();
        assert!(text.contains("&apos;A&amp;B&apos;!A2&amp;&quot;A1&quot;"), "{text}");
        assert!(text.contains(r#"localSheetId="1">'A&amp;B'!A&#49;</definedName>"#), "{text}");
    }

    #[test]
    fn comment_text_keeps_raw_entities_when_refs_shift() {
        let xml = br#"<comments><commentList><comment ref="A1" author="A&amp;B"><text><t> &#65;&amp;&lt; </t></text></comment></commentList></comments>"#;
        let out = shift_comments_xml(xml, &ShiftPlan::insert(Axis::Row, 1, 1));
        let text = String::from_utf8(out).unwrap();
        assert!(text.contains(r#"ref="A2" author="A&amp;B""#), "{text}");
        assert!(text.contains("<t> &#65;&amp;&lt; </t>"), "{text}");
    }

    #[test]
    fn empty_ops_returns_empty_mutations() {
        let inputs = SheetXmlInputs::empty();
        let out = apply_workbook_shift(inputs, &[]);
        assert!(out.file_patches.is_empty());
    }

    #[test]
    fn shifts_sheet_xml_via_orchestrator() {
        let sheet_xml = r#"<sheetData><row r="5"><c r="A5"><v>1</v></c></row></sheetData>"#;
        let mut inputs = SheetXmlInputs::empty();
        inputs
            .sheets
            .insert("Sheet1".to_string(), sheet_xml.as_bytes());
        inputs
            .sheet_paths
            .insert("Sheet1".to_string(), "xl/worksheets/sheet1.xml".to_string());
        let ops = vec![AxisShiftOp {
            sheet: "Sheet1".to_string(),
            axis: Axis::Row,
            idx: 5,
            n: 3,
        }];
        let out = apply_workbook_shift(inputs, &ops);
        let new_sheet = out
            .file_patches
            .get("xl/worksheets/sheet1.xml")
            .expect("sheet was rewritten");
        let s = String::from_utf8_lossy(new_sheet);
        assert!(s.contains(r#"<row r="8">"#));
    }

    #[test]
    fn shifts_comments_ref() {
        let xml = r#"<comments><commentList><comment ref="A5" authorId="0"><text><t>hi</t></text></comment></commentList></comments>"#;
        let p = ShiftPlan::insert(Axis::Row, 5, 3);
        let out = shift_comments_xml(xml.as_bytes(), &p);
        let s = String::from_utf8_lossy(&out);
        assert!(s.contains(r#"ref="A8""#));
    }

    #[test]
    fn drops_tombstoned_comment() {
        let xml = r#"<comments><commentList><comment ref="A5" authorId="0"><text><t>hi</t></text></comment></commentList></comments>"#;
        let p = ShiftPlan::delete(Axis::Row, 5, 1);
        let out = shift_comments_xml(xml.as_bytes(), &p);
        let s = String::from_utf8_lossy(&out);
        assert!(!s.contains("<comment "));
    }

    #[test]
    fn shifts_vml_anchor_row() {
        // 8 ints: 0, 0, 4, 0, 2, 0, 6, 0
        // (col-left=0, col-left-off=0, row-top=4, row-top-off=0,
        //  col-right=2, col-right-off=0, row-bottom=6, row-bottom-off=0)
        // Insert 3 rows at idx=5 (0-based row=4 is 1-based row=5).
        // Both 2 (row-top) and 6 (row-bottom) are >= 5 in 1-based, so
        // both shift by +3 → row-top=7, row-bottom=9.
        let vml = r#"<v:shape><x:ClientData><x:Anchor>0, 0, 4, 0, 2, 0, 6, 0</x:Anchor></x:ClientData></v:shape>"#;
        let p = ShiftPlan::insert(Axis::Row, 5, 3);
        let out = shift_vml_xml(vml.as_bytes(), &p);
        let s = String::from_utf8_lossy(&out);
        assert!(s.contains("0, 0, 7, 0, 2, 0, 9, 0"));
    }

    #[test]
    fn vml_shift_does_not_match_shapetype_as_shape() {
        let vml = r#"<xml><v:shapetype id="_x0000_t202"/><v:shape><x:ClientData><x:Anchor>0, 0, 4, 0, 2, 0, 6, 0</x:Anchor></x:ClientData></v:shape></xml>"#;
        let p = ShiftPlan::insert(Axis::Row, 5, 1);
        let out = shift_vml_xml(vml.as_bytes(), &p);
        let s = String::from_utf8_lossy(&out);
        assert!(s.contains("<v:shapetype id=\"_x0000_t202\"/>"));
        assert!(s.contains("<v:shape>"));
        assert!(s.contains("0, 0, 5, 0, 2, 0, 7, 0"));
    }

    #[test]
    fn shifts_unprefixed_vml_shape_anchor() {
        let vml = r#"<xml><shapetype id="_x0000_t202"/><shape><ClientData><Anchor>0, 0, 4, 0, 2, 0, 6, 0</Anchor></ClientData></shape></xml>"#;
        let p = ShiftPlan::insert(Axis::Row, 5, 1);
        let out = shift_vml_xml(vml.as_bytes(), &p);
        let s = String::from_utf8_lossy(&out);
        assert!(s.contains("<shapetype id=\"_x0000_t202\"/>"));
        assert!(s.contains("<shape>"));
        assert!(s.contains("<Anchor>0, 0, 5, 0, 2, 0, 7, 0</Anchor>"));
    }

    #[test]
    fn shifts_arbitrary_prefixed_vml_shape_anchor() {
        let vml = r#"<xml><ns1:shapetype id="_x0000_t202"/><ns1:shape><ns2:ClientData><ns2:Anchor>0, 0, 4, 0, 2, 0, 6, 0</ns2:Anchor></ns2:ClientData></ns1:shape></xml>"#;
        let p = ShiftPlan::insert(Axis::Row, 5, 1);
        let out = shift_vml_xml(vml.as_bytes(), &p);
        let s = String::from_utf8_lossy(&out);
        assert!(s.contains("<ns1:shapetype id=\"_x0000_t202\"/>"));
        assert!(s.contains("<ns1:shape>"));
        assert!(s.contains("<ns2:Anchor>0, 0, 5, 0, 2, 0, 7, 0</ns2:Anchor>"));
    }

    #[test]
    fn shifts_prefixed_vml_comment_row_column_markers() {
        let vml = r#"<xml><ns1:shape><ns2:ClientData ObjectType="Note"><ns2:Row>1</ns2:Row><ns2:Column>2</ns2:Column></ns2:ClientData></ns1:shape></xml>"#;

        let row_plan = ShiftPlan::delete(Axis::Row, 1, 1);
        let row_out = shift_vml_xml(vml.as_bytes(), &row_plan);
        let row_s = String::from_utf8_lossy(&row_out);
        assert!(row_s.contains("<ns2:Row>0</ns2:Row>"), "{row_s}");
        assert!(row_s.contains("<ns2:Column>2</ns2:Column>"), "{row_s}");

        let col_plan = ShiftPlan::delete(Axis::Col, 1, 1);
        let col_out = shift_vml_xml(vml.as_bytes(), &col_plan);
        let col_s = String::from_utf8_lossy(&col_out);
        assert!(col_s.contains("<ns2:Row>1</ns2:Row>"), "{col_s}");
        assert!(col_s.contains("<ns2:Column>1</ns2:Column>"), "{col_s}");
    }

    #[test]
    fn drops_vml_shape_when_anchor_tombstoned() {
        let vml = r#"<v:shape><x:ClientData><x:Anchor>0, 0, 4, 0, 2, 0, 4, 0</x:Anchor></x:ClientData></v:shape>"#;
        let p = ShiftPlan::delete(Axis::Row, 5, 1);
        let out = shift_vml_xml(vml.as_bytes(), &p);
        let s = String::from_utf8_lossy(&out);
        assert!(!s.contains("v:shape"));
    }

    #[test]
    fn shrinks_vml_row_anchor_when_bottom_edge_deleted() {
        let vml = r#"<v:shape><x:ClientData><x:Anchor>0, 0, 3, 0, 2, 0, 6, 0</x:Anchor></x:ClientData></v:shape>"#;
        let p = ShiftPlan::delete(Axis::Row, 7, 1);
        let out = shift_vml_xml(vml.as_bytes(), &p);
        let s = String::from_utf8_lossy(&out);
        assert!(s.contains("<v:shape>"));
        assert!(s.contains("0, 0, 3, 0, 2, 0, 5, 0"));
    }

    #[test]
    fn shrinks_vml_row_anchor_when_top_edge_deleted() {
        let vml = r#"<v:shape><x:ClientData><x:Anchor>0, 0, 4, 0, 2, 0, 6, 0</x:Anchor></x:ClientData></v:shape>"#;
        let p = ShiftPlan::delete(Axis::Row, 5, 1);
        let out = shift_vml_xml(vml.as_bytes(), &p);
        let s = String::from_utf8_lossy(&out);
        assert!(s.contains("<v:shape>"));
        assert!(s.contains("0, 0, 4, 0, 2, 0, 5, 0"));
    }

    #[test]
    fn shrinks_vml_col_anchor_when_right_edge_deleted() {
        let vml = r#"<v:shape><x:ClientData><x:Anchor>1, 0, 3, 0, 4, 0, 6, 0</x:Anchor></x:ClientData></v:shape>"#;
        let p = ShiftPlan::delete(Axis::Col, 5, 1);
        let out = shift_vml_xml(vml.as_bytes(), &p);
        let s = String::from_utf8_lossy(&out);
        assert!(s.contains("<v:shape>"));
        assert!(s.contains("1, 0, 3, 0, 3, 0, 6, 0"));
    }

    #[test]
    fn shrinks_vml_col_anchor_when_left_edge_deleted() {
        let vml = r#"<v:shape><x:ClientData><x:Anchor>1, 0, 3, 0, 4, 0, 6, 0</x:Anchor></x:ClientData></v:shape>"#;
        let p = ShiftPlan::delete(Axis::Col, 2, 1);
        let out = shift_vml_xml(vml.as_bytes(), &p);
        let s = String::from_utf8_lossy(&out);
        assert!(s.contains("<v:shape>"));
        assert!(s.contains("1, 0, 3, 0, 3, 0, 6, 0"));
    }

    #[test]
    fn shifts_workbook_scope_defined_name() {
        let xml = r#"<workbook><definedNames><definedName name="Total">Sheet1!$A$5</definedName></definedNames></workbook>"#;
        let p = ShiftPlan::insert(Axis::Row, 5, 3);
        let out = shift_defined_names(xml.as_bytes(), &p, "Sheet1", Some(0));
        let s = String::from_utf8_lossy(&out);
        assert!(s.contains("Sheet1!$A$8"));
    }

    #[test]
    fn preserves_full_sheet_row_defined_name_at_tail_boundary() {
        let xml = r#"<workbook><definedNames><definedName name="_bdm.hidden" hidden="1">'HB ROI'!$1:$1048576</definedName></definedNames></workbook>"#;
        let insert_tail = ShiftPlan::insert(Axis::Row, crate::MAX_ROW, 1);
        let inserted = shift_defined_names(xml.as_bytes(), &insert_tail, "HB ROI", Some(0));
        let inserted_text = String::from_utf8_lossy(&inserted);
        assert!(inserted_text.contains("$1:$1048576"));
        assert!(!inserted_text.contains("#REF!"));

        let delete_tail = ShiftPlan::delete(Axis::Row, crate::MAX_ROW, 1);
        let deleted = shift_defined_names(xml.as_bytes(), &delete_tail, "HB ROI", Some(0));
        let deleted_text = String::from_utf8_lossy(&deleted);
        assert!(deleted_text.contains("$1:$1048576"));
        assert!(!deleted_text.contains("$1:$1048575"));
    }

    #[test]
    fn preserves_unnecessary_source_quotes_in_print_area_defined_name() {
        let xml = r#"<workbook><definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">'Cover'!$A$1:$L$28</definedName></definedNames></workbook>"#;
        let plan = ShiftPlan::insert(Axis::Row, crate::MAX_ROW, 1);
        let shifted = shift_defined_names(xml.as_bytes(), &plan, "Cover", Some(0));
        let shifted_text = String::from_utf8_lossy(&shifted);
        assert!(
            shifted_text.contains("'Cover'!$A$1:$L$28")
                || shifted_text.contains("&apos;Cover&apos;!$A$1:$L$28")
                || shifted_text.contains("&#39;Cover&#39;!$A$1:$L$28")
        );
        assert!(!shifted_text.contains(">Cover!$A$1:$L$28<"));
    }

    #[test]
    fn skips_per_sheet_dn_for_other_sheet() {
        let xml = r#"<workbook><definedNames><definedName name="Total" localSheetId="1">A5</definedName></definedNames></workbook>"#;
        let p = ShiftPlan::insert(Axis::Row, 5, 3);
        // op_sheet is at position 0; DN's localSheetId is 1 → skip.
        let out = shift_defined_names(xml.as_bytes(), &p, "Sheet1", Some(0));
        let s = String::from_utf8_lossy(&out);
        assert!(s.contains(">A5<"));
    }

    #[test]
    fn shifts_per_sheet_dn_for_matching_sheet() {
        let xml = r#"<workbook><definedNames><definedName name="Total" localSheetId="0">A5</definedName></definedNames></workbook>"#;
        let p = ShiftPlan::insert(Axis::Row, 5, 3);
        let out = shift_defined_names(xml.as_bytes(), &p, "Sheet1", Some(0));
        let s = String::from_utf8_lossy(&out);
        assert!(s.contains(">A8<"));
    }

    #[test]
    fn passes_through_unrelated_workbook_xml() {
        let xml = r#"<workbook><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#;
        let p = ShiftPlan::insert(Axis::Row, 5, 3);
        let out = shift_defined_names(xml.as_bytes(), &p, "Sheet1", Some(0));
        let s = String::from_utf8_lossy(&out);
        assert!(s.contains(r#"sheet name="Sheet1""#));
    }
}
