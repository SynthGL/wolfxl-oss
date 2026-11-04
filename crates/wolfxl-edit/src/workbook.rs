//! Workbook-level package edits: listing worksheets, adding and renaming
//! sheets, and flagging a workbook for full recalculation.
//!
//! Every function takes part XML as text and returns rewritten text. Changes
//! are spliced into the original source at byte spans reported by the XML
//! parser, so markup outside the touched elements keeps its exact bytes.

use std::collections::{HashMap, HashSet};

use quick_xml::escape::{escape, unescape};
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;

use crate::EditError;

const WORKBOOK_PART: &str = "xl/workbook.xml";
/// Directory that relative targets in the workbook relationships resolve from.
const WORKBOOK_DIR: &str = "xl";
const WORKBOOK_RELS_PART: &str = "xl/_rels/workbook.xml.rels";
const CONTENT_TYPES_PART: &str = "[Content_Types].xml";
const APP_PART: &str = "docProps/app.xml";

const REL_NS_TRANSITIONAL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL_NS_STRICT: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const MAIN_NS_STRICT: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const WORKSHEET_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const CALC_CHAIN_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml";
/// HeadingPairs group that counts worksheet titles in `docProps/app.xml`.
const WORKSHEETS_HEADING: &str = "Worksheets";
/// Sheet names longer than this many UTF-16 code units are rejected.
const MAX_SHEET_NAME_UNITS: usize = 31;
/// Workbook children that follow `calcPr` in the CT_Workbook sequence
/// (ECMA-376 Part 1, 18.2.27).
const AFTER_CALC_PR: &[&str] = &[
    "oleSize",
    "customWorkbookViews",
    "pivotCaches",
    "smartTagPr",
    "smartTagTypes",
    "webPublishing",
    "fileRecoveryPr",
    "webPublishObjects",
    "extLst",
];

/// A worksheet listed in `xl/workbook.xml`, with its resolved package part.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Sheet {
    pub name: String,
    /// Package part path without a leading slash, e.g. `xl/worksheets/sheet1.xml`.
    pub path: String,
    pub id: u32,
    pub rel_id: String,
}

/// Worksheets in workbook order. Chartsheets, dialogsheets, and macrosheets
/// are skipped; every sheet entry is still validated.
pub(crate) fn sheets(workbook: &str, rels: &str) -> Result<Vec<Sheet>, EditError> {
    let wb = Doc::parse(workbook, WORKBOOK_PART)?;
    let (_, entries) = workbook_sheets(&wb)?;
    let rel_doc = Doc::parse(rels, WORKBOOK_RELS_PART)?;
    let rel_list = parse_rels(&rel_doc)?;
    let by_id: HashMap<&str, &Rel> = rel_list.iter().map(|r| (r.id.as_str(), r)).collect();

    let mut out: Vec<Sheet> = Vec::new();
    for entry in entries {
        let rel = by_id.get(entry.rel_id.as_str()).ok_or_else(|| {
            xml_error(
                WORKBOOK_PART,
                format!(
                    "sheet '{}' refers to missing relationship '{}'",
                    entry.name, entry.rel_id
                ),
            )
        })?;
        if rel.external {
            return Err(EditError::Unsupported(format!(
                "sheet '{}' refers to an external target '{}'",
                entry.name, rel.target
            )));
        }
        if !is_rel_type(&rel.ty, "worksheet") {
            continue;
        }
        let path = resolve_target(&rel.target).map_err(|reason| {
            EditError::Invalid(format!(
                "{WORKBOOK_RELS_PART}: sheet '{}': {reason}",
                entry.name
            ))
        })?;
        if let Some(other) = out.iter().find(|s| s.path.eq_ignore_ascii_case(&path)) {
            return Err(xml_error(
                WORKBOOK_PART,
                format!(
                    "sheets '{}' and '{}' share the part '{path}'",
                    other.name, entry.name
                ),
            ));
        }
        out.push(Sheet {
            name: entry.name,
            path,
            id: entry.sheet_id,
            rel_id: entry.rel_id,
        });
    }
    Ok(out)
}

/// Check a sheet name against the spreadsheet naming rules and against the
/// names in `existing`, ignoring the sheet named `excluding` (a rename source).
pub(crate) fn validate_name(
    name: &str,
    existing: &[Sheet],
    excluding: Option<&str>,
) -> Result<(), EditError> {
    check_name_rules(name)?;
    check_unique(name, existing.iter().map(|s| s.name.as_str()), excluding)
}

/// Add a worksheet entry named `name` whose part lives at `path`.
///
/// `position` indexes `existing` (worksheets in workbook order): the new sheet
/// is inserted before `existing[position]`, or appended after every sheet when
/// `position == existing.len()`. Returns the rewritten workbook, workbook
/// relationships, content types, and (when given) extended properties. The
/// caller writes the worksheet part itself.
#[allow(clippy::too_many_arguments)]
pub(crate) fn add_sheet(
    workbook: &str,
    rels: &str,
    content_types: &str,
    app: Option<&str>,
    existing: &[Sheet],
    name: &str,
    position: usize,
    path: &str,
) -> Result<(String, String, String, Option<String>), EditError> {
    check_name_rules(name)?;
    validate_part_path(path)?;
    if position > existing.len() {
        return Err(EditError::Invalid(format!(
            "sheet position {position} is out of range (the workbook has {} worksheets)",
            existing.len()
        )));
    }

    let wb = Doc::parse(workbook, WORKBOOK_PART)?;
    let (sheets_idx, entries) = workbook_sheets(&wb)?;
    check_unique(
        name,
        entries
            .iter()
            .map(|s| s.name.as_str())
            .chain(existing.iter().map(|s| s.name.as_str())),
        None,
    )?;
    let insert_at = match existing.get(position) {
        Some(before) => entries
            .iter()
            .position(|s| s.name == before.name)
            .ok_or_else(|| EditError::UnknownSheet(before.name.clone()))?,
        None => entries.len(),
    };

    let rel_doc = Doc::parse(rels, WORKBOOK_RELS_PART)?;
    let rel_list = parse_rels(&rel_doc)?;
    let ct = Doc::parse(content_types, CONTENT_TYPES_PART)?;
    let overrides = parse_overrides(&ct)?;

    let in_use = existing.iter().any(|s| s.path.eq_ignore_ascii_case(path))
        || rel_list.iter().any(|r| {
            !r.external && resolve_target(&r.target).is_ok_and(|p| p.eq_ignore_ascii_case(path))
        })
        || overrides
            .iter()
            .any(|o| resolve_target(&o.part).is_ok_and(|p| p.eq_ignore_ascii_case(path)));
    if in_use {
        return Err(EditError::Invalid(format!(
            "package part '{path}' is already in use"
        )));
    }

    let sheet_id = allocate_sheet_id(entries.iter().map(|s| s.sheet_id));
    let rel_id = allocate_rel_id(&rel_list);
    let strict = wb.namespace(sheets_idx)?.as_deref() == Some(MAIN_NS_STRICT);

    // Workbook: the new <sheet> entry plus index shifts for an insertion.
    let sheets_node = &wb.nodes[sheets_idx];
    let mut element = format!(
        "<{}sheet name=\"{}\" sheetId=\"{sheet_id}\"",
        sheets_node.prefix(),
        escape(name)
    );
    match wb.rel_prefix_in_scope(sheets_idx)? {
        Some(prefix) => element.push_str(&format!(" {prefix}:id=\"{rel_id}\"")),
        None => {
            let ns = if strict {
                REL_NS_STRICT
            } else {
                REL_NS_TRANSITIONAL
            };
            element.push_str(&format!(" xmlns:r=\"{ns}\" r:id=\"{rel_id}\""));
        }
    }
    element.push_str("/>");

    let mut wb_edits: Vec<Splice> = Vec::new();
    match entries.get(insert_at) {
        Some(next) => {
            let at = wb.nodes[next.node].start;
            wb_edits.push((at, at, element));
        }
        None => wb_edits.push(wb.append_child(sheets_idx, &element)),
    }
    if insert_at < entries.len() {
        shift_sheet_indexes(&wb, insert_at as u32, &mut wb_edits)?;
    }
    let wb_out = apply(workbook, wb_edits)?;

    // Relationships: reuse the package's worksheet relationship type spelling.
    let rel_type = rel_list
        .iter()
        .find(|r| is_rel_type(&r.ty, "worksheet"))
        .map(|r| r.ty.clone())
        .unwrap_or_else(|| {
            let ns = if strict {
                REL_NS_STRICT
            } else {
                REL_NS_TRANSITIONAL
            };
            format!("{ns}/worksheet")
        });
    let target = match path.strip_prefix("xl/") {
        Some(rest) => percent_encode(rest),
        None => format!("/{}", percent_encode(path)),
    };
    let relationship = format!(
        "<{}Relationship Id=\"{rel_id}\" Type=\"{}\" Target=\"{}\"/>",
        rel_doc.nodes[ROOT].prefix(),
        escape(rel_type.as_str()),
        escape(target.as_str())
    );
    let rels_out = apply(rels, vec![rel_doc.append_child(ROOT, &relationship)])?;

    let override_elem = format!(
        "<{}Override PartName=\"/{}\" ContentType=\"{WORKSHEET_CONTENT_TYPE}\"/>",
        ct.nodes[ROOT].prefix(),
        escape(percent_encode(path).as_str())
    );
    let ct_out = apply(content_types, vec![ct.append_child(ROOT, &override_elem)])?;

    let app_out = match app {
        Some(app) => Some(app_add_title(app, name, position)?),
        None => None,
    };
    Ok((wb_out, rels_out, ct_out, app_out))
}

/// Rename the sheet entry `old` to `new` in the workbook and in the
/// extended-properties title list. Formula and defined-name text is left to
/// the caller.
pub(crate) fn rename_sheet(
    workbook: &str,
    app: Option<&str>,
    old: &str,
    new: &str,
) -> Result<(String, Option<String>), EditError> {
    let wb = Doc::parse(workbook, WORKBOOK_PART)?;
    let (_, entries) = workbook_sheets(&wb)?;
    let entry = entries
        .iter()
        .find(|s| s.name == old)
        .ok_or_else(|| EditError::UnknownSheet(old.to_string()))?;
    check_name_rules(new)?;
    check_unique(new, entries.iter().map(|s| s.name.as_str()), Some(old))?;
    if old == new {
        return Ok((workbook.to_string(), app.map(str::to_string)));
    }

    let mut attrs = wb.nodes[entry.node].attrs.clone();
    set_attr(&mut attrs, "name", escape(new).into_owned());
    let wb_out = apply(workbook, vec![wb.retag(entry.node, &attrs)])?;
    let app_out = match app {
        Some(app) => Some(app_rename_title(app, old, new)?),
        None => None,
    };
    Ok((wb_out, app_out))
}

/// Mark the workbook for full recalculation on open and drop the calculation
/// chain. Returns the rewritten workbook, relationships, content types, and
/// the calcChain part paths the caller must remove from the package.
pub(crate) fn recalculate(
    workbook: &str,
    rels: &str,
    content_types: &str,
) -> Result<(String, String, String, Vec<String>), EditError> {
    let wb = Doc::parse(workbook, WORKBOOK_PART)?;
    wb.expect_root("workbook")?;
    let root = &wb.nodes[ROOT];
    let mut wb_edits: Vec<Splice> = Vec::new();
    match wb.child(ROOT, "calcPr") {
        Some(idx) => {
            if wb.attr(idx, "fullCalcOnLoad")?.as_deref() != Some("1") {
                let mut attrs = wb.nodes[idx].attrs.clone();
                set_attr(&mut attrs, "fullCalcOnLoad", "1".to_string());
                wb_edits.push(wb.retag(idx, &attrs));
            }
        }
        None => {
            if root.empty {
                return Err(xml_error(WORKBOOK_PART, "the <workbook> element is empty"));
            }
            let element = format!("<{}calcPr fullCalcOnLoad=\"1\"/>", root.prefix());
            let at = wb.children[ROOT]
                .iter()
                .find(|&&c| AFTER_CALC_PR.contains(&wb.nodes[c].local()))
                .map_or(root.close_start, |&c| wb.nodes[c].start);
            wb_edits.push((at, at, element));
        }
    }
    let wb_out = apply(workbook, wb_edits)?;

    let rel_doc = Doc::parse(rels, WORKBOOK_RELS_PART)?;
    let rel_list = parse_rels(&rel_doc)?;
    let mut removed: Vec<String> = Vec::new();
    let mut rel_edits: Vec<Splice> = Vec::new();
    for rel in rel_list.iter().filter(|r| is_rel_type(&r.ty, "calcChain")) {
        rel_edits.push(rel_doc.remove(rel.node));
        if rel.external {
            continue;
        }
        let path = resolve_target(&rel.target).map_err(|reason| {
            EditError::Invalid(format!(
                "{WORKBOOK_RELS_PART}: relationship '{}': {reason}",
                rel.id
            ))
        })?;
        push_unique_path(&mut removed, path);
    }
    let rels_out = apply(rels, rel_edits)?;

    let ct = Doc::parse(content_types, CONTENT_TYPES_PART)?;
    let overrides = parse_overrides(&ct)?;
    let mut ct_edits: Vec<Splice> = Vec::new();
    for o in &overrides {
        let part = resolve_target(&o.part).ok();
        let names_removed = part
            .as_deref()
            .is_some_and(|p| removed.iter().any(|r| r.eq_ignore_ascii_case(p)));
        if o.content_type == CALC_CHAIN_CONTENT_TYPE || names_removed {
            ct_edits.push(ct.remove(o.node));
            if let Some(part) = part {
                push_unique_path(&mut removed, part);
            }
        }
    }
    let ct_out = apply(content_types, ct_edits)?;
    Ok((wb_out, rels_out, ct_out, removed))
}

// Sheet entries ---------------------------------------------------------------

struct WorkbookSheet {
    node: usize,
    name: String,
    sheet_id: u32,
    rel_id: String,
}

/// The `<sheets>` element and every `<sheet>` entry, validated.
fn workbook_sheets(doc: &Doc) -> Result<(usize, Vec<WorkbookSheet>), EditError> {
    doc.expect_root("workbook")?;
    let sheets_idx = doc
        .child(ROOT, "sheets")
        .ok_or_else(|| xml_error(doc.part, "missing <sheets> element"))?;
    let mut out: Vec<WorkbookSheet> = Vec::new();
    let mut names: HashSet<String> = HashSet::new();
    let mut ids: HashSet<u32> = HashSet::new();
    let mut rel_ids: HashSet<String> = HashSet::new();
    for idx in doc.children_named(sheets_idx, "sheet") {
        let name = doc
            .attr(idx, "name")?
            .ok_or_else(|| xml_error(doc.part, "a <sheet> element has no name attribute"))?;
        if name.is_empty() {
            return Err(xml_error(doc.part, "a <sheet> element has an empty name"));
        }
        let raw_id = doc
            .attr(idx, "sheetId")?
            .ok_or_else(|| xml_error(doc.part, format!("sheet '{name}' has no sheetId")))?;
        let sheet_id: u32 = raw_id.trim().parse().map_err(|_| {
            xml_error(
                doc.part,
                format!("sheet '{name}' has invalid sheetId '{raw_id}'"),
            )
        })?;
        let rel_id = doc
            .rel_id(idx)?
            .ok_or_else(|| xml_error(doc.part, format!("sheet '{name}' has no relationship id")))?;
        if !names.insert(name_key(&name)) {
            return Err(xml_error(
                doc.part,
                format!("duplicate sheet name '{name}'"),
            ));
        }
        if !ids.insert(sheet_id) {
            return Err(xml_error(doc.part, format!("duplicate sheetId {sheet_id}")));
        }
        if !rel_ids.insert(rel_id.clone()) {
            return Err(xml_error(
                doc.part,
                format!("relationship id '{rel_id}' is used by more than one sheet"),
            ));
        }
        out.push(WorkbookSheet {
            node: idx,
            name,
            sheet_id,
            rel_id,
        });
    }
    Ok((sheets_idx, out))
}

/// Shift sheet-index attributes for a sheet inserted at workbook index `at`:
/// `definedName/@localSheetId`, `workbookView/@activeTab`, and
/// `workbookView/@firstSheet`, so each keeps pointing at the same sheet.
fn shift_sheet_indexes(doc: &Doc, at: u32, edits: &mut Vec<Splice>) -> Result<(), EditError> {
    if let Some(names) = doc.child(ROOT, "definedNames") {
        for idx in doc.children_named(names, "definedName") {
            let Some(raw) = doc.attr(idx, "localSheetId")? else {
                continue;
            };
            let index = parse_index(doc, &raw, "localSheetId")?;
            if index >= at {
                let mut attrs = doc.nodes[idx].attrs.clone();
                set_attr(
                    &mut attrs,
                    "localSheetId",
                    increment(doc, index)?.to_string(),
                );
                edits.push(doc.retag(idx, &attrs));
            }
        }
    }
    if let Some(views) = doc.child(ROOT, "bookViews") {
        for idx in doc.children_named(views, "workbookView") {
            let mut attrs = doc.nodes[idx].attrs.clone();
            let mut changed = false;
            // An absent activeTab means the first sheet is active.
            let active = match doc.attr(idx, "activeTab")? {
                Some(raw) => parse_index(doc, &raw, "activeTab")?,
                None => 0,
            };
            if active >= at {
                set_attr(&mut attrs, "activeTab", increment(doc, active)?.to_string());
                changed = true;
            }
            if let Some(raw) = doc.attr(idx, "firstSheet")? {
                let first = parse_index(doc, &raw, "firstSheet")?;
                if first > at {
                    set_attr(&mut attrs, "firstSheet", increment(doc, first)?.to_string());
                    changed = true;
                }
            }
            if changed {
                edits.push(doc.retag(idx, &attrs));
            }
        }
    }
    Ok(())
}

fn parse_index(doc: &Doc, raw: &str, attr: &str) -> Result<u32, EditError> {
    raw.trim()
        .parse()
        .map_err(|_| xml_error(doc.part, format!("invalid {attr} value '{raw}'")))
}

fn increment(doc: &Doc, value: u32) -> Result<u32, EditError> {
    value
        .checked_add(1)
        .ok_or_else(|| xml_error(doc.part, format!("sheet index {value} is too large")))
}

fn allocate_sheet_id(ids: impl Iterator<Item = u32> + Clone) -> u32 {
    match ids.clone().max() {
        None => 1,
        Some(max) => max.checked_add(1).unwrap_or_else(|| {
            let used: HashSet<u32> = ids.collect();
            (1..=u32::MAX)
                .find(|id| !used.contains(id))
                .unwrap_or(u32::MAX)
        }),
    }
}

fn allocate_rel_id(rels: &[Rel]) -> String {
    let used: HashSet<&str> = rels.iter().map(|r| r.id.as_str()).collect();
    let mut n = rels
        .iter()
        .filter_map(|r| r.id.strip_prefix("rId")?.parse::<u64>().ok())
        .max()
        .map_or(1, |max| max.saturating_add(1));
    loop {
        let id = format!("rId{n}");
        if !used.contains(id.as_str()) {
            return id;
        }
        n = n.wrapping_add(1);
    }
}

// Sheet names -----------------------------------------------------------------

fn check_name_rules(name: &str) -> Result<(), EditError> {
    let invalid =
        |reason: &str| EditError::Invalid(format!("invalid sheet name '{name}': {reason}"));
    if name.is_empty() {
        return Err(EditError::Invalid(
            "sheet name must not be empty".to_string(),
        ));
    }
    if name.encode_utf16().count() > MAX_SHEET_NAME_UNITS {
        return Err(invalid("longer than 31 characters"));
    }
    if let Some(c) = name
        .chars()
        .find(|c| matches!(c, '[' | ']' | ':' | '*' | '?' | '/' | '\\'))
    {
        return Err(invalid(&format!("contains the character '{c}'")));
    }
    if name.starts_with('\'') || name.ends_with('\'') {
        return Err(invalid("starts or ends with an apostrophe"));
    }
    if name
        .chars()
        .any(|c| c.is_control() || c == '\u{FFFE}' || c == '\u{FFFF}')
    {
        return Err(invalid("contains a control character"));
    }
    if name.eq_ignore_ascii_case("History") {
        return Err(invalid("'History' is reserved"));
    }
    Ok(())
}

fn check_unique<'a>(
    name: &str,
    others: impl Iterator<Item = &'a str>,
    excluding: Option<&str>,
) -> Result<(), EditError> {
    let key = name_key(name);
    for other in others {
        if Some(other) != excluding && name_key(other) == key {
            return Err(EditError::Invalid(format!(
                "a sheet named '{other}' already exists"
            )));
        }
    }
    Ok(())
}

/// Sheet names compare case-insensitively.
fn name_key(name: &str) -> String {
    name.to_lowercase()
}

// Relationships and content types ---------------------------------------------

struct Rel {
    node: usize,
    id: String,
    ty: String,
    target: String,
    external: bool,
}

fn parse_rels(doc: &Doc) -> Result<Vec<Rel>, EditError> {
    doc.expect_root("Relationships")?;
    let mut out: Vec<Rel> = Vec::new();
    let mut ids: HashSet<String> = HashSet::new();
    for idx in doc.children_named(ROOT, "Relationship") {
        let required = |key: &str| -> Result<String, EditError> {
            doc.attr(idx, key)?.ok_or_else(|| {
                xml_error(
                    doc.part,
                    format!("a <Relationship> element has no {key} attribute"),
                )
            })
        };
        let id = required("Id")?;
        let ty = required("Type")?;
        let target = required("Target")?;
        let external = match doc.attr(idx, "TargetMode")?.as_deref() {
            None | Some("Internal") => false,
            Some("External") => true,
            Some(other) => {
                return Err(xml_error(
                    doc.part,
                    format!("relationship '{id}' has invalid TargetMode '{other}'"),
                ))
            }
        };
        if !ids.insert(id.clone()) {
            return Err(xml_error(
                doc.part,
                format!("duplicate relationship id '{id}'"),
            ));
        }
        out.push(Rel {
            node: idx,
            id,
            ty,
            target,
            external,
        });
    }
    Ok(out)
}

/// True when `ty` is the transitional or strict relationship type `kind`.
fn is_rel_type(ty: &str, kind: &str) -> bool {
    ty.strip_prefix(REL_NS_TRANSITIONAL)
        .or_else(|| ty.strip_prefix(REL_NS_STRICT))
        .and_then(|rest| rest.strip_prefix('/'))
        == Some(kind)
}

fn is_rel_ns(uri: &str) -> bool {
    uri == REL_NS_TRANSITIONAL || uri == REL_NS_STRICT
}

struct Override {
    node: usize,
    part: String,
    content_type: String,
}

fn parse_overrides(doc: &Doc) -> Result<Vec<Override>, EditError> {
    doc.expect_root("Types")?;
    let mut out: Vec<Override> = Vec::new();
    for idx in doc.children_named(ROOT, "Override") {
        let required = |key: &str| -> Result<String, EditError> {
            doc.attr(idx, key)?.ok_or_else(|| {
                xml_error(
                    doc.part,
                    format!("an <Override> element has no {key} attribute"),
                )
            })
        };
        out.push(Override {
            node: idx,
            part: required("PartName")?,
            content_type: required("ContentType")?,
        });
    }
    Ok(out)
}

fn push_unique_path(paths: &mut Vec<String>, path: String) {
    if !paths.iter().any(|p| p.eq_ignore_ascii_case(&path)) {
        paths.push(path);
    }
}

// Part paths ------------------------------------------------------------------

/// Resolve a relationship target (or content-type part name) against
/// `xl/workbook.xml` into a package path without a leading slash. Absolute
/// URIs, query or fragment parts, backslashes, empty segments, and `..`
/// segments that climb above the package root are refused.
fn resolve_target(target: &str) -> Result<String, String> {
    if target.is_empty() {
        return Err("relationship target is empty".to_string());
    }
    if target.contains(['\\', '?', '#']) {
        return Err(format!("target '{target}' is not a plain part path"));
    }
    let (mut segments, rest): (Vec<String>, &str) = match target.strip_prefix('/') {
        Some(rest) => (Vec::new(), rest),
        None => {
            let first = target.split('/').next().unwrap_or_default();
            if first.contains(':') {
                return Err(format!(
                    "target '{target}' is an absolute URI, not a package part"
                ));
            }
            (vec![WORKBOOK_DIR.to_string()], target)
        }
    };
    for raw in rest.split('/') {
        match raw {
            "" => return Err(format!("target '{target}' has an empty path segment")),
            "." => {}
            ".." => {
                if segments.pop().is_none() {
                    return Err(format!("target '{target}' escapes the package root"));
                }
            }
            _ => {
                let segment = percent_decode(raw)
                    .filter(|s| !s.is_empty() && s != "." && s != ".." && !s.contains(['/', '\\']))
                    .ok_or_else(|| format!("target '{target}' has a malformed path segment"))?;
                segments.push(segment);
            }
        }
    }
    if segments.is_empty() {
        return Err(format!("target '{target}' names no part"));
    }
    Ok(segments.join("/"))
}

/// Refuse part paths that are absolute, empty, or not normalized.
fn validate_part_path(path: &str) -> Result<(), EditError> {
    let valid = !path.is_empty()
        && !path.starts_with('/')
        && !path.contains('\\')
        && path
            .split('/')
            .all(|s| !s.is_empty() && s != "." && s != "..");
    if valid {
        Ok(())
    } else {
        Err(EditError::Invalid(format!(
            "invalid package part path '{path}'"
        )))
    }
}

fn percent_decode(segment: &str) -> Option<String> {
    let bytes = segment.as_bytes();
    let mut out: Vec<u8> = Vec::with_capacity(bytes.len());
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] == b'%' {
            let hex = segment.get(i + 1..i + 3)?;
            out.push(u8::from_str_radix(hex, 16).ok()?);
            i += 3;
        } else {
            out.push(bytes[i]);
            i += 1;
        }
    }
    String::from_utf8(out).ok()
}

/// Percent-encode a part path for use as a URI reference. Colons are encoded
/// so a first segment can never read as a URI scheme.
fn percent_encode(path: &str) -> String {
    let mut out = String::with_capacity(path.len());
    for &b in path.as_bytes() {
        let keep = b.is_ascii_alphanumeric()
            || matches!(
                b,
                b'-' | b'.'
                    | b'_'
                    | b'~'
                    | b'!'
                    | b'$'
                    | b'&'
                    | b'\''
                    | b'('
                    | b')'
                    | b'*'
                    | b'+'
                    | b','
                    | b';'
                    | b'='
                    | b'@'
                    | b'/'
            );
        if keep {
            out.push(b as char);
        } else {
            out.push_str(&format!("%{b:02X}"));
        }
    }
    out
}

// Extended properties ---------------------------------------------------------

struct HeadingGroup {
    name: String,
    count: usize,
    count_node: usize,
}

struct AppLayout {
    vector: usize,
    titles: Vec<usize>,
    /// HeadingPairs groups, present only when their counts cover every title.
    groups: Option<Vec<HeadingGroup>>,
}

fn app_layout(doc: &Doc) -> Result<Option<AppLayout>, EditError> {
    doc.expect_root("Properties")?;
    let Some(titles_of_parts) = doc.child(ROOT, "TitlesOfParts") else {
        return Ok(None);
    };
    let Some(vector) = doc.child(titles_of_parts, "vector") else {
        return Ok(None);
    };
    let titles = doc.children[vector].clone();
    let groups = heading_groups(doc)?
        .filter(|groups| groups.iter().map(|g| g.count).sum::<usize>() == titles.len());
    Ok(Some(AppLayout {
        vector,
        titles,
        groups,
    }))
}

/// HeadingPairs as (name, count) groups; `None` when the vector is missing or
/// not a well-formed list of pairs.
fn heading_groups(doc: &Doc) -> Result<Option<Vec<HeadingGroup>>, EditError> {
    let Some(pairs) = doc.child(ROOT, "HeadingPairs") else {
        return Ok(None);
    };
    let Some(vector) = doc.child(pairs, "vector") else {
        return Ok(None);
    };
    let variants: Vec<usize> = doc.children_named(vector, "variant").collect();
    if !variants.len().is_multiple_of(2) {
        return Ok(None);
    }
    let mut groups: Vec<HeadingGroup> = Vec::new();
    for pair in variants.chunks(2) {
        let (Some(&name_node), Some(&count_node)) =
            (doc.children[pair[0]].first(), doc.children[pair[1]].first())
        else {
            return Ok(None);
        };
        let Ok(count) = doc.text(count_node).trim().parse::<usize>() else {
            return Ok(None);
        };
        groups.push(HeadingGroup {
            name: doc.text_value(name_node)?,
            count,
            count_node,
        });
    }
    Ok(Some(groups))
}

/// Insert `name` into the Worksheets titles at worksheet index `position`.
/// Properties without a consistent Worksheets group are returned unchanged.
fn app_add_title(app: &str, name: &str, position: usize) -> Result<String, EditError> {
    let doc = Doc::parse(app, APP_PART)?;
    let Some(layout) = app_layout(&doc)? else {
        return Ok(app.to_string());
    };
    let Some(groups) = &layout.groups else {
        return Ok(app.to_string());
    };
    let Some(gi) = groups
        .iter()
        .position(|g| g.name.eq_ignore_ascii_case(WORKSHEETS_HEADING))
    else {
        return Ok(app.to_string());
    };
    let group = &groups[gi];
    let offset: usize = groups[..gi].iter().map(|g| g.count).sum();
    let at = offset + position.min(group.count);

    let count_node = &doc.nodes[group.count_node];
    let vector = &doc.nodes[layout.vector];
    let mut vector_attrs = vector.attrs.clone();
    set_attr(
        &mut vector_attrs,
        "size",
        (layout.titles.len() + 1).to_string(),
    );
    let title_name = layout
        .titles
        .first()
        .map(|&t| doc.nodes[t].name.clone())
        .unwrap_or_else(|| format!("{}lpstr", vector.prefix()));
    let title = format!("<{title_name}>{}</{title_name}>", escape(name));

    let mut edits: Vec<Splice> = vec![(
        count_node.tag_end,
        count_node.close_start,
        (group.count + 1).to_string(),
    )];
    match layout.titles.get(at) {
        Some(&next) => {
            let start = doc.nodes[next].start;
            edits.push(doc.retag(layout.vector, &vector_attrs));
            edits.push((start, start, title));
        }
        None if vector.empty => edits.push((
            vector.start,
            vector.end,
            format!(
                "{}{title}</{}>",
                render_tag(&vector.name, &vector_attrs, false),
                vector.name
            ),
        )),
        None => {
            edits.push(doc.retag(layout.vector, &vector_attrs));
            edits.push((vector.close_start, vector.close_start, title));
        }
    }
    apply(app, edits)
}

/// Replace the title exactly equal to `old`, searching the Worksheets group
/// first when its bounds are known.
fn app_rename_title(app: &str, old: &str, new: &str) -> Result<String, EditError> {
    let doc = Doc::parse(app, APP_PART)?;
    let Some(layout) = app_layout(&doc)? else {
        return Ok(app.to_string());
    };
    let range = layout
        .groups
        .as_ref()
        .and_then(|groups| {
            let gi = groups
                .iter()
                .position(|g| g.name.eq_ignore_ascii_case(WORKSHEETS_HEADING))?;
            let offset: usize = groups[..gi].iter().map(|g| g.count).sum();
            Some(offset..offset + groups[gi].count)
        })
        .unwrap_or(0..layout.titles.len());
    for &idx in &layout.titles[range] {
        let node = &doc.nodes[idx];
        if node.empty || doc.text_value(idx)? != old {
            continue;
        }
        let edit = (node.tag_end, node.close_start, escape(new).into_owned());
        return apply(app, vec![edit]);
    }
    Ok(app.to_string())
}

// XML spans -------------------------------------------------------------------

/// Replace `src[start..end]` with the string.
type Splice = (usize, usize, String);

/// Index of the root element in [`Doc::nodes`].
const ROOT: usize = 0;

fn xml_error(part: &str, msg: impl std::fmt::Display) -> EditError {
    EditError::Xml(format!("{part}: {msg}"))
}

/// One element with the byte spans of its tags. Attribute values stay raw
/// (still escaped) so rewritten tags reproduce them exactly.
struct Node {
    name: String,
    attrs: Vec<(String, String)>,
    /// Offset of `<` in the start tag.
    start: usize,
    /// Offset just past the start tag.
    tag_end: usize,
    /// Offset of `</` in the end tag; equals `tag_end` for empty elements.
    close_start: usize,
    /// Offset just past the element.
    end: usize,
    empty: bool,
    parent: Option<usize>,
}

impl Node {
    fn local(&self) -> &str {
        match self.name.rfind(':') {
            Some(i) => &self.name[i + 1..],
            None => &self.name,
        }
    }

    /// Qualified-name prefix including the colon, or "".
    fn prefix(&self) -> &str {
        match self.name.rfind(':') {
            Some(i) => &self.name[..=i],
            None => "",
        }
    }
}

struct Doc<'a> {
    src: &'a str,
    part: &'static str,
    nodes: Vec<Node>,
    children: Vec<Vec<usize>>,
}

impl<'a> Doc<'a> {
    fn parse(src: &'a str, part: &'static str) -> Result<Self, EditError> {
        let mut reader = Reader::from_str(src);
        let mut nodes: Vec<Node> = Vec::new();
        let mut children: Vec<Vec<usize>> = Vec::new();
        let mut stack: Vec<usize> = Vec::new();
        loop {
            let before = reader.buffer_position() as usize;
            let event = reader.read_event().map_err(|e| xml_error(part, e))?;
            let after = reader.buffer_position() as usize;
            let (element, empty) = match event {
                Event::Start(e) => (e, false),
                Event::Empty(e) => (e, true),
                Event::End(_) => {
                    let idx = stack
                        .pop()
                        .ok_or_else(|| xml_error(part, "unexpected closing tag"))?;
                    nodes[idx].close_start = markup_start(src, before, after, part)?;
                    nodes[idx].end = after;
                    continue;
                }
                Event::Eof => break,
                _ => continue,
            };
            let parent = stack.last().copied();
            if parent.is_none() && !nodes.is_empty() {
                return Err(xml_error(part, "more than one root element"));
            }
            let idx = nodes.len();
            nodes.push(Node {
                name: element.name().into_inner().to_string(),
                attrs: collect_attrs(&element, part)?,
                start: markup_start(src, before, after, part)?,
                tag_end: after,
                close_start: after,
                end: after,
                empty,
                parent,
            });
            children.push(Vec::new());
            if let Some(p) = parent {
                children[p].push(idx);
            }
            if !empty {
                stack.push(idx);
            }
        }
        if let Some(&idx) = stack.last() {
            return Err(xml_error(
                part,
                format!("unclosed element <{}>", nodes[idx].name),
            ));
        }
        if nodes.is_empty() {
            return Err(xml_error(part, "no root element"));
        }
        Ok(Self {
            src,
            part,
            nodes,
            children,
        })
    }

    fn expect_root(&self, local: &str) -> Result<(), EditError> {
        let root = &self.nodes[ROOT];
        if root.local() == local {
            Ok(())
        } else {
            Err(xml_error(
                self.part,
                format!("expected root element <{local}>, found <{}>", root.name),
            ))
        }
    }

    /// Unescaped attribute value by qualified name.
    fn attr(&self, idx: usize, key: &str) -> Result<Option<String>, EditError> {
        let Some((_, raw)) = self.nodes[idx].attrs.iter().find(|(k, _)| k == key) else {
            return Ok(None);
        };
        unescape(raw)
            .map(|v| Some(v.into_owned()))
            .map_err(|e| xml_error(self.part, format!("attribute {key}: {e}")))
    }

    fn child(&self, idx: usize, local: &str) -> Option<usize> {
        self.children_named(idx, local).next()
    }

    fn children_named<'s>(
        &'s self,
        idx: usize,
        local: &'s str,
    ) -> impl Iterator<Item = usize> + 's {
        self.children[idx]
            .iter()
            .copied()
            .filter(move |&c| self.nodes[c].local() == local)
    }

    /// Raw element content between the start and end tags.
    fn text(&self, idx: usize) -> &'a str {
        let node = &self.nodes[idx];
        &self.src[node.tag_end..node.close_start]
    }

    fn text_value(&self, idx: usize) -> Result<String, EditError> {
        unescape(self.text(idx))
            .map(|v| v.into_owned())
            .map_err(|e| xml_error(self.part, format!("<{}>: {e}", self.nodes[idx].name)))
    }

    /// Namespace URI bound to `prefix` (no colon; "" for the default) in
    /// scope at `idx`.
    fn prefix_uri(&self, idx: usize, prefix: &str) -> Result<Option<String>, EditError> {
        let key = if prefix.is_empty() {
            "xmlns".to_string()
        } else {
            format!("xmlns:{prefix}")
        };
        let mut current = Some(idx);
        while let Some(i) = current {
            if let Some(uri) = self.attr(i, &key)? {
                return Ok(Some(uri));
            }
            current = self.nodes[i].parent;
        }
        Ok(None)
    }

    /// Namespace URI of the element at `idx`.
    fn namespace(&self, idx: usize) -> Result<Option<String>, EditError> {
        let prefix = self.nodes[idx].prefix();
        self.prefix_uri(idx, prefix.strip_suffix(':').unwrap_or(prefix))
    }

    /// Prefix bound to the relationships namespace in scope at `idx`, honoring
    /// nearer redeclarations.
    fn rel_prefix_in_scope(&self, idx: usize) -> Result<Option<String>, EditError> {
        let mut seen: HashSet<&str> = HashSet::new();
        let mut current = Some(idx);
        while let Some(i) = current {
            for (key, _) in &self.nodes[i].attrs {
                let Some(prefix) = key.strip_prefix("xmlns:") else {
                    continue;
                };
                if seen.insert(prefix) && self.attr(i, key)?.is_some_and(|uri| is_rel_ns(&uri)) {
                    return Ok(Some(prefix.to_string()));
                }
            }
            current = self.nodes[i].parent;
        }
        Ok(None)
    }

    /// The relationship-namespace `id` attribute (usually `r:id`).
    fn rel_id(&self, idx: usize) -> Result<Option<String>, EditError> {
        for (key, _) in &self.nodes[idx].attrs {
            let Some(prefix) = key.strip_suffix(":id") else {
                continue;
            };
            if prefix == "xmlns" || prefix.contains(':') {
                continue;
            }
            if self
                .prefix_uri(idx, prefix)?
                .is_some_and(|uri| is_rel_ns(&uri))
            {
                return self.attr(idx, key);
            }
        }
        Ok(None)
    }

    /// Rewrite the start tag of `idx` with `attrs`.
    fn retag(&self, idx: usize, attrs: &[(String, String)]) -> Splice {
        let node = &self.nodes[idx];
        (
            node.start,
            node.tag_end,
            render_tag(&node.name, attrs, node.empty),
        )
    }

    /// Append `child` as the last child of `idx`, expanding an empty element.
    fn append_child(&self, idx: usize, child: &str) -> Splice {
        let node = &self.nodes[idx];
        if node.empty {
            let open = render_tag(&node.name, &node.attrs, false);
            (
                node.start,
                node.end,
                format!("{open}{child}</{}>", node.name),
            )
        } else {
            (node.close_start, node.close_start, child.to_string())
        }
    }

    fn remove(&self, idx: usize) -> Splice {
        let node = &self.nodes[idx];
        (node.start, node.end, String::new())
    }
}

fn collect_attrs(element: &BytesStart<'_>, part: &str) -> Result<Vec<(String, String)>, EditError> {
    element
        .attributes()
        .map(|attr| -> Result<(String, String), EditError> {
            let attr = attr.map_err(|e| xml_error(part, e))?;
            Ok((attr.key.into_inner().to_string(), attr.value.into_owned()))
        })
        .collect()
}

/// Offset of the `<` opening the markup the reader just consumed.
fn markup_start(src: &str, before: usize, after: usize, part: &str) -> Result<usize, EditError> {
    src.get(before..after)
        .and_then(|s| s.find('<'))
        .map(|offset| before + offset)
        .ok_or_else(|| xml_error(part, format!("cannot locate markup at byte {before}")))
}

/// Render a tag from raw attribute values, keeping them escaped.
fn render_tag(name: &str, attrs: &[(String, String)], empty: bool) -> String {
    let mut out = String::with_capacity(name.len() + 2 + attrs.len() * 16);
    out.push('<');
    out.push_str(name);
    for (key, value) in attrs {
        out.push(' ');
        out.push_str(key);
        out.push_str("=\"");
        // Values from single-quoted attributes may hold a bare double quote.
        out.push_str(&value.replace('"', "&quot;"));
        out.push('"');
    }
    out.push_str(if empty { "/>" } else { ">" });
    out
}

/// Set a raw (already escaped) attribute value, keeping its position.
fn set_attr(attrs: &mut Vec<(String, String)>, key: &str, value: String) {
    match attrs.iter_mut().find(|(k, _)| k == key) {
        Some(slot) => slot.1 = value,
        None => attrs.push((key.to_string(), value)),
    }
}

/// Apply non-overlapping splices. Insertions at the same offset keep their
/// order in `edits`.
fn apply(src: &str, mut edits: Vec<Splice>) -> Result<String, EditError> {
    edits.sort_by_key(|e| (e.0, e.1));
    let extra: usize = edits.iter().map(|e| e.2.len()).sum();
    let mut out = String::with_capacity(src.len() + extra);
    let mut cursor = 0;
    for (start, end, text) in &edits {
        if *start < cursor {
            return Err(EditError::Invalid(
                "internal error: overlapping XML edits".to_string(),
            ));
        }
        out.push_str(&src[cursor..*start]);
        out.push_str(text);
        cursor = *end;
    }
    out.push_str(&src[cursor..]);
    Ok(out)
}

#[cfg(test)]
mod tests {
    use super::*;

    macro_rules! rel_type {
        ($kind:literal) => {
            concat!(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/",
                $kind
            )
        };
    }

    const WB: &str = concat!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
        r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" "#,
        r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
        r#"<workbookPr/><bookViews><workbookView xWindow="0" activeTab="1" firstSheet="2"/></bookViews>"#,
        r#"<sheets><sheet name="Data" sheetId="1" r:id="rId1"/>"#,
        r#"<sheet name="Chart" sheetId="3" r:id="rId2"/>"#,
        r#"<sheet name="Summary" sheetId="2" r:id="rId3"/></sheets>"#,
        r#"<definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">Data!$A$1:$B$2</definedName>"#,
        r#"<definedName name="Total" localSheetId="2">Summary!$A$1</definedName>"#,
        r#"<definedName name="Global">Data!$A$1</definedName></definedNames>"#,
        r#"<calcPr calcId="191029"/><extLst><ext uri="{0}"/></extLst></workbook>"#,
    );

    const RELS: &str = concat!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
        r#"<Relationship Id="rId1" Type=""#,
        rel_type!("worksheet"),
        r#"" Target="worksheets/sheet1.xml"/>"#,
        r#"<Relationship Id="rId2" Type=""#,
        rel_type!("chartsheet"),
        r#"" Target="chartsheets/sheet1.xml"/>"#,
        r#"<Relationship Id="rId3" Type=""#,
        rel_type!("worksheet"),
        r#"" Target="/xl/worksheets/sheet2.xml"/>"#,
        r#"<Relationship Id="rId4" Type=""#,
        rel_type!("styles"),
        r#"" Target="styles.xml"/>"#,
        r#"<Relationship Id="rId7" Type=""#,
        rel_type!("calcChain"),
        r#"" Target="calcChain.xml"/>"#,
        r#"</Relationships>"#,
    );

    const CT: &str = concat!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#,
        r#"<Default Extension="xml" ContentType="application/xml"/>"#,
        r#"<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>"#,
        r#"<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#,
        r#"<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#,
        r#"<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>"#,
        r#"</Types>"#,
    );

    const APP: &str = concat!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
        r#"<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" "#,
        r#"xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">"#,
        r#"<Application>WolfXL</Application><HeadingPairs><vt:vector size="4" baseType="variant">"#,
        r#"<vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>2</vt:i4></vt:variant>"#,
        r#"<vt:variant><vt:lpstr>Charts</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant>"#,
        r#"</vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="3" baseType="lpstr">"#,
        r#"<vt:lpstr>Data</vt:lpstr><vt:lpstr>Summary</vt:lpstr><vt:lpstr>Chart</vt:lpstr>"#,
        r#"</vt:vector></TitlesOfParts></Properties>"#,
    );

    fn names(list: &[Sheet]) -> Vec<&str> {
        list.iter().map(|s| s.name.as_str()).collect()
    }

    fn sheet(name: &str) -> Sheet {
        Sheet {
            name: name.to_string(),
            path: format!("xl/worksheets/{name}.xml"),
            id: 1,
            rel_id: "rId1".to_string(),
        }
    }

    fn minimal_wb(sheets: &str) -> String {
        format!(
            concat!(
                r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" "#,
                r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
                "<sheets>{}</sheets></workbook>"
            ),
            sheets
        )
    }

    fn minimal_rels(rels: &str) -> String {
        format!(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{rels}</Relationships>"#
        )
    }

    fn worksheet_rel(id: &str, target: &str) -> String {
        format!(
            r#"<Relationship Id="{id}" Type="{}" Target="{target}"/>"#,
            rel_type!("worksheet")
        )
    }

    #[test]
    fn sheets_lists_worksheets_in_order_and_resolves_targets() {
        let list = sheets(WB, RELS).unwrap();
        assert_eq!(
            list,
            vec![
                Sheet {
                    name: "Data".into(),
                    path: "xl/worksheets/sheet1.xml".into(),
                    id: 1,
                    rel_id: "rId1".into(),
                },
                Sheet {
                    name: "Summary".into(),
                    path: "xl/worksheets/sheet2.xml".into(),
                    id: 2,
                    rel_id: "rId3".into(),
                },
            ]
        );
    }

    #[test]
    fn sheets_honors_custom_prefixes() {
        let wb = concat!(
            r#"<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" "#,
            r#"xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
            r#"<x:sheets><x:sheet name="A &amp; B" sheetId="5" rel:id="rId9"/></x:sheets></x:workbook>"#,
        );
        let rels = minimal_rels(&worksheet_rel("rId9", "worksheets/sheet%201.xml"));
        let list = sheets(wb, &rels).unwrap();
        assert_eq!(list[0].name, "A & B");
        assert_eq!(list[0].path, "xl/worksheets/sheet 1.xml");
        assert_eq!(list[0].id, 5);
    }

    #[test]
    fn sheets_refuses_external_and_malformed_entries() {
        let external = minimal_rels(&format!(
            r#"<Relationship Id="rId1" Type="{}" Target="http://example.com/a.xml" TargetMode="External"/>"#,
            rel_type!("worksheet")
        ));
        let wb = minimal_wb(r#"<sheet name="A" sheetId="1" r:id="rId1"/>"#);
        assert!(matches!(
            sheets(&wb, &external),
            Err(EditError::Unsupported(_))
        ));

        let rels = minimal_rels(&worksheet_rel("rId1", "worksheets/sheet1.xml"));
        let cases = [
            minimal_wb(r#"<sheet name="A" sheetId="1" r:id="rId2"/>"#),
            minimal_wb(r#"<sheet sheetId="1" r:id="rId1"/>"#),
            minimal_wb(r#"<sheet name="" sheetId="1" r:id="rId1"/>"#),
            minimal_wb(r#"<sheet name="A" sheetId="x" r:id="rId1"/>"#),
            minimal_wb(r#"<sheet name="A" sheetId="1"/>"#),
            minimal_wb(r#"<sheet name="A" sheetId="1" r:id="rId1"/><sheet name="a" sheetId="2" r:id="rId1"/>"#),
            minimal_wb(r#"<sheet name="A" sheetId="1" r:id="rId1"/><sheet name="B" sheetId="1" r:id="rId2"/>"#),
            r#"<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>"#.to_string(),
            "<worksheet/>".to_string(),
            minimal_wb(r#"<sheet name="A" sheetId="1" r:id="rId1">"#),
        ];
        for wb in &cases {
            assert!(matches!(sheets(wb, &rels), Err(EditError::Xml(_))), "{wb}");
        }

        let escaping = minimal_rels(&worksheet_rel("rId1", "../../outside.xml"));
        let wb = minimal_wb(r#"<sheet name="A" sheetId="1" r:id="rId1"/>"#);
        assert!(matches!(sheets(&wb, &escaping), Err(EditError::Invalid(_))));
    }

    #[test]
    fn resolve_target_normalizes_and_refuses_unsafe_paths() {
        assert_eq!(
            resolve_target("worksheets/./sheet1.xml").unwrap(),
            "xl/worksheets/sheet1.xml"
        );
        assert_eq!(
            resolve_target("../customXml/item1.xml").unwrap(),
            "customXml/item1.xml"
        );
        assert_eq!(resolve_target("/xl/a%20b.xml").unwrap(), "xl/a b.xml");
        for bad in [
            "",
            "../../x.xml",
            "http://x/y.xml",
            "a//b.xml",
            "a\\b.xml",
            "a.xml#x",
            "a%2Fb.xml",
            "a%zz.xml",
            "//host/x.xml",
        ] {
            assert!(resolve_target(bad).is_err(), "{bad}");
        }
    }

    #[test]
    fn validate_name_enforces_naming_rules() {
        let existing = [sheet("Data")];
        for bad in [
            "", "a[b", "a]", "a:b", "a*b", "a?b", "a/b", "a\\b", "'a", "a'", "a\u{1}b", "History",
            "history", "DATA",
        ] {
            assert!(
                matches!(
                    validate_name(bad, &existing, None),
                    Err(EditError::Invalid(_))
                ),
                "{bad}"
            );
        }
        assert!(validate_name("it's", &existing, None).is_ok());
        assert!(validate_name(&"x".repeat(31), &existing, None).is_ok());
        assert!(validate_name(&"x".repeat(32), &existing, None).is_err());
        // Characters outside the BMP take two UTF-16 units.
        let emoji = "\u{1F600}".repeat(15);
        assert!(validate_name(&format!("{emoji}x"), &existing, None).is_ok());
        assert!(validate_name(&format!("{emoji}xy"), &existing, None).is_err());
        assert!(validate_name("DATA", &existing, Some("Data")).is_ok());
    }

    #[test]
    fn add_sheet_appends_entry_relationship_override_and_title() {
        let existing = sheets(WB, RELS).unwrap();
        let (wb, rels, ct, app) = add_sheet(
            WB,
            RELS,
            CT,
            Some(APP),
            &existing,
            "New",
            existing.len(),
            "xl/worksheets/sheet3.xml",
        )
        .unwrap();
        assert!(wb.contains(r#"<sheet name="New" sheetId="4" r:id="rId8"/></sheets>"#));
        // Appending shifts nothing.
        assert!(wb.contains(r#"localSheetId="0""#) && wb.contains(r#"localSheetId="2""#));
        assert!(wb.contains(r#"activeTab="1" firstSheet="2""#));
        assert_eq!(
            &wb[..WB.find("<sheets>").unwrap()],
            &WB[..WB.find("<sheets>").unwrap()]
        );
        assert!(rels.ends_with(&format!(
            r#"<Relationship Id="rId8" Type="{}" Target="worksheets/sheet3.xml"/></Relationships>"#,
            rel_type!("worksheet")
        )));
        assert!(ct.ends_with(&format!(
            r#"<Override PartName="/xl/worksheets/sheet3.xml" ContentType="{WORKSHEET_CONTENT_TYPE}"/></Types>"#
        )));
        let app = app.unwrap();
        assert!(app.contains("<vt:i4>3</vt:i4>"));
        assert!(app.contains(r#"<vt:vector size="4" baseType="lpstr">"#));
        assert!(app.contains(
            "<vt:lpstr>Summary</vt:lpstr><vt:lpstr>New</vt:lpstr><vt:lpstr>Chart</vt:lpstr>"
        ));
        let after = sheets(&wb, &rels).unwrap();
        assert_eq!(names(&after), ["Data", "Summary", "New"]);
        assert_eq!(after[2].path, "xl/worksheets/sheet3.xml");
    }

    #[test]
    fn add_sheet_insertion_shifts_sheet_indexes() {
        let existing = sheets(WB, RELS).unwrap();
        let (wb, rels, _, app) = add_sheet(
            WB,
            RELS,
            CT,
            Some(APP),
            &existing,
            "First",
            0,
            "xl/worksheets/sheet3.xml",
        )
        .unwrap();
        assert!(wb.contains(
            r#"<sheets><sheet name="First" sheetId="4" r:id="rId8"/><sheet name="Data""#
        ));
        assert!(wb.contains(r#"localSheetId="1">Data!"#));
        assert!(wb.contains(r#"localSheetId="3">Summary!"#));
        assert!(wb.contains(r#"<definedName name="Global">"#));
        assert!(wb.contains(r#"<workbookView xWindow="0" activeTab="2" firstSheet="3"/>"#));
        assert!(app
            .unwrap()
            .contains("<vt:lpstr>First</vt:lpstr><vt:lpstr>Data</vt:lpstr>"));
        assert_eq!(
            names(&sheets(&wb, &rels).unwrap()),
            ["First", "Data", "Summary"]
        );

        // Before "Summary", which is third in the workbook after the chartsheet.
        let (wb, rels, _, app) = add_sheet(
            WB,
            RELS,
            CT,
            Some(APP),
            &existing,
            "Mid",
            1,
            "xl/worksheets/sheet3.xml",
        )
        .unwrap();
        assert!(wb.contains(r#"localSheetId="0">Data!"#));
        assert!(wb.contains(r#"localSheetId="3">Summary!"#));
        assert!(wb.contains(r#"activeTab="1" firstSheet="2""#));
        assert!(app.unwrap().contains(
            "<vt:lpstr>Data</vt:lpstr><vt:lpstr>Mid</vt:lpstr><vt:lpstr>Summary</vt:lpstr>"
        ));
        assert_eq!(
            names(&sheets(&wb, &rels).unwrap()),
            ["Data", "Mid", "Summary"]
        );
    }

    #[test]
    fn add_sheet_sets_active_tab_when_inserting_before_default_active_sheet() {
        let wb = format!(
            concat!(
                r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" "#,
                r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
                r#"<bookViews><workbookView xWindow="0"/></bookViews><sheets>{}</sheets></workbook>"#
            ),
            r#"<sheet name="A" sheetId="1" r:id="rId1"/>"#
        );
        let rels = minimal_rels(&worksheet_rel("rId1", "worksheets/sheet1.xml"));
        let ct = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>"#;
        let existing = sheets(&wb, &rels).unwrap();
        let (wb, _, ct, app) = add_sheet(
            &wb,
            &rels,
            ct,
            None,
            &existing,
            "B",
            0,
            "xl/worksheets/sheet2.xml",
        )
        .unwrap();
        assert!(wb.contains(r#"<workbookView xWindow="0" activeTab="1"/>"#));
        assert!(ct.contains(r#"<Override PartName="/xl/worksheets/sheet2.xml""#));
        assert!(ct.ends_with("</Types>"));
        assert!(app.is_none());
    }

    #[test]
    fn add_sheet_declares_relationship_namespace_when_missing_in_scope() {
        let wb = concat!(
            r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets>"#,
            r#"<sheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" name="A" sheetId="1" r:id="rId1"/>"#,
            r#"</sheets></workbook>"#,
        );
        let rels = minimal_rels(&worksheet_rel("rId1", "worksheets/sheet1.xml"));
        let existing = sheets(wb, &rels).unwrap();
        let ct = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>"#;
        let (wb, rels, _, _) = add_sheet(
            wb,
            &rels,
            ct,
            None,
            &existing,
            "B",
            1,
            "xl/worksheets/sheet2.xml",
        )
        .unwrap();
        assert!(wb.contains(
            r#"<sheet name="B" sheetId="2" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId2"/>"#
        ));
        assert_eq!(names(&sheets(&wb, &rels).unwrap()), ["A", "B"]);
    }

    #[test]
    fn add_sheet_uses_existing_prefixes_and_expands_empty_containers() {
        let wb = concat!(
            r#"<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" "#,
            r#"xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><x:sheets/></x:workbook>"#,
        );
        let rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#;
        let ct = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>"#;
        let (wb, rels, ct, _) = add_sheet(
            wb,
            rels,
            ct,
            None,
            &[],
            "One",
            0,
            "xl/worksheets/sheet1.xml",
        )
        .unwrap();
        assert!(
            wb.contains(r#"<x:sheets><x:sheet name="One" sheetId="1" rel:id="rId1"/></x:sheets>"#)
        );
        assert!(rels.contains(r#"<Relationship Id="rId1""#) && rels.ends_with("</Relationships>"));
        assert!(ct.ends_with("/></Types>"));
        assert_eq!(names(&sheets(&wb, &rels).unwrap()), ["One"]);
    }

    #[test]
    fn add_sheet_rejects_conflicts() {
        let existing = sheets(WB, RELS).unwrap();
        let add = |name: &str, position: usize, path: &str| {
            add_sheet(WB, RELS, CT, Some(APP), &existing, name, position, path)
        };
        assert!(matches!(
            add("data", 2, "xl/worksheets/sheet3.xml"),
            Err(EditError::Invalid(_))
        ));
        assert!(matches!(
            add("chart", 2, "xl/worksheets/sheet3.xml"),
            Err(EditError::Invalid(_))
        ));
        assert!(matches!(
            add("Bad?", 2, "xl/worksheets/sheet3.xml"),
            Err(EditError::Invalid(_))
        ));
        assert!(matches!(
            add("New", 3, "xl/worksheets/sheet3.xml"),
            Err(EditError::Invalid(_))
        ));
        assert!(matches!(
            add("New", 2, "XL/Worksheets/Sheet2.xml"),
            Err(EditError::Invalid(_))
        ));
        assert!(matches!(
            add("New", 2, "xl/calcChain.xml"),
            Err(EditError::Invalid(_))
        ));
        assert!(matches!(
            add("New", 2, "xl/../sheet3.xml"),
            Err(EditError::Invalid(_))
        ));
        assert!(matches!(
            add("New", 2, "/xl/sheet3.xml"),
            Err(EditError::Invalid(_))
        ));
    }

    #[test]
    fn rename_sheet_rewrites_entry_and_title() {
        let (wb, app) = rename_sheet(WB, Some(APP), "Data", "Raw & Cooked").unwrap();
        assert!(wb.contains(r#"<sheet name="Raw &amp; Cooked" sheetId="1" r:id="rId1"/>"#));
        assert!(
            wb.contains("Data!$A$1:$B$2"),
            "formula text is left to the caller"
        );
        assert_eq!(
            wb.replace(r#"name="Raw &amp; Cooked""#, r#"name="Data""#),
            WB
        );
        let app = app.unwrap();
        assert!(app.contains("<vt:lpstr>Raw &amp; Cooked</vt:lpstr><vt:lpstr>Summary</vt:lpstr>"));
        assert_eq!(app.replace("Raw &amp; Cooked", "Data"), APP);

        let (wb, app) = rename_sheet(WB, None, "Data", "DATA").unwrap();
        assert!(wb.contains(r#"name="DATA""#));
        assert!(app.is_none());
        let (wb, app) = rename_sheet(WB, Some(APP), "Data", "Data").unwrap();
        assert_eq!((wb.as_str(), app.as_deref()), (WB, Some(APP)));
    }

    #[test]
    fn rename_sheet_prefers_worksheet_titles() {
        // A named range titled "Total" precedes the worksheets group.
        let app = concat!(
            r#"<Properties xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">"#,
            r#"<HeadingPairs><vt:vector size="4" baseType="variant">"#,
            r#"<vt:variant><vt:lpstr>Named Ranges</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant>"#,
            r#"<vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant>"#,
            r#"</vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="2" baseType="lpstr">"#,
            r#"<vt:lpstr>Total</vt:lpstr><vt:lpstr>Total</vt:lpstr></vt:vector></TitlesOfParts></Properties>"#,
        );
        let wb = minimal_wb(r#"<sheet name="Total" sheetId="1" r:id="rId1"/>"#);
        let (_, app_out) = rename_sheet(&wb, Some(app), "Total", "Sum").unwrap();
        assert!(app_out
            .unwrap()
            .contains("<vt:lpstr>Total</vt:lpstr><vt:lpstr>Sum</vt:lpstr>"));
    }

    #[test]
    fn rename_sheet_rejects_unknown_and_duplicate_names() {
        assert!(matches!(
            rename_sheet(WB, None, "Nope", "X"),
            Err(EditError::UnknownSheet(_))
        ));
        assert!(matches!(
            rename_sheet(WB, None, "data", "X"),
            Err(EditError::UnknownSheet(_))
        ));
        assert!(matches!(
            rename_sheet(WB, None, "Data", "summary"),
            Err(EditError::Invalid(_))
        ));
        assert!(matches!(
            rename_sheet(WB, None, "Data", "Chart"),
            Err(EditError::Invalid(_))
        ));
        assert!(matches!(
            rename_sheet(WB, None, "Data", "a/b"),
            Err(EditError::Invalid(_))
        ));
    }

    #[test]
    fn recalculate_sets_flag_and_drops_calc_chain() {
        let (wb, rels, ct, removed) = recalculate(WB, RELS, CT).unwrap();
        assert!(wb.contains(r#"<calcPr calcId="191029" fullCalcOnLoad="1"/>"#));
        assert_eq!(wb.replace(r#" fullCalcOnLoad="1""#, ""), WB);
        assert!(!rels.contains("calcChain"));
        assert!(rels.contains(r#"Id="rId4""#));
        assert!(!ct.contains("calcChain"));
        assert!(ct.contains("/xl/worksheets/sheet2.xml"));
        assert_eq!(removed, ["xl/calcChain.xml"]);

        // Already flagged and nothing to drop: every part is unchanged.
        let (wb2, rels2, ct2, removed2) = recalculate(&wb, &rels, &ct).unwrap();
        assert_eq!((wb2, rels2, ct2), (wb, rels, ct));
        assert!(removed2.is_empty());
    }

    #[test]
    fn recalculate_overrides_existing_flag_and_inserts_calc_pr_in_schema_order() {
        let flagged = WB.replace(
            r#"calcId="191029""#,
            r#"fullCalcOnLoad="0" calcId="191029""#,
        );
        let (wb, ..) = recalculate(&flagged, RELS, CT).unwrap();
        assert!(wb.contains(r#"<calcPr fullCalcOnLoad="1" calcId="191029"/>"#));

        let without = WB.replace(r#"<calcPr calcId="191029"/>"#, "");
        let (wb, ..) = recalculate(&without, RELS, CT).unwrap();
        assert!(wb.contains(r#"</definedNames><calcPr fullCalcOnLoad="1"/><extLst>"#));

        let bare = minimal_wb(r#"<sheet name="A" sheetId="1" r:id="rId1"/>"#);
        let (wb, ..) = recalculate(&bare, RELS, CT).unwrap();
        assert!(wb.ends_with(r#"</sheets><calcPr fullCalcOnLoad="1"/></workbook>"#));
    }

    #[test]
    fn recalculate_requires_valid_parts() {
        assert!(matches!(
            recalculate("<workbook/>", RELS, CT),
            Err(EditError::Xml(_))
        ));
        assert!(matches!(
            recalculate(WB, "<Types/>", CT),
            Err(EditError::Xml(_))
        ));
        assert!(matches!(
            recalculate(WB, RELS, "<Relationships/>"),
            Err(EditError::Xml(_))
        ));
        let bad = RELS.replace(
            r#"Target="calcChain.xml""#,
            r#"Target="../../calcChain.xml""#,
        );
        assert!(matches!(
            recalculate(WB, &bad, CT),
            Err(EditError::Invalid(_))
        ));
    }
}
