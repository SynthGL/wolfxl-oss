//! Table builder + ZIP scanner for the modify-mode patcher (RFC-024).
//!
//! Used by `XlsxPatcher::do_save`'s Phase 2.5f to:
//! 1. Scan the source ZIP for existing `xl/tables/tableN.xml` parts to
//!    learn (a) how many tables already exist (the next part-index
//!    starting point), (b) which `id` attributes are already in use
//!    (so newly-allocated ids stay workbook-unique), and (c) which
//!    `name` attributes are taken (collision detection).
//! 2. Serialize each [`TablePatch`] into the bytes of a new
//!    `xl/tables/tableN.xml` part.
//! 3. Allocate one new rId per table in the sheet's rels graph.
//! 4. Emit the `<tableParts>` block whose bytes feed
//!    `wolfxl_merger::SheetBlock::TableParts` (slot 37).
//! 5. Queue `[Content_Types].xml` `Override` entries for the new parts.
//!
//! The native writer's `crates/wolfxl-writer/src/emit/tables_xml.rs`
//! is reused as the XML serializer — the writer's `Table` model is
//! the input shape — so write-mode and modify-mode produce
//! byte-equivalent table parts for the same input.

use std::collections::HashSet;

use quick_xml::events::Event;
use quick_xml::Reader as XmlReader;
use zip::ZipArchive;

use wolfxl_rels::{rt, PartIdAllocator, RelId, RelsGraph, TargetMode};
use wolfxl_writer::emit::tables_xml as writer_tables_xml;
use wolfxl_writer::model::table::{Table, TableColumn, TableStyle};

// ---------------------------------------------------------------------------
// Public types
// ---------------------------------------------------------------------------

/// One queued table addition. Mirrors the openpyxl `Table` shape that
/// the patcher's Python wrapper receives (`name`, `displayName`, `ref`,
/// columns, optional style, header/totals flags).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct TablePatch {
    pub name: String,
    pub display_name: String,
    pub ref_range: String,
    pub columns: Vec<String>,
    pub style: Option<TableStylePatch>,
    /// Default `1`. `0` means no header strip — alters the `autoFilter`
    /// shape (handled inside the writer's emitter).
    pub header_row_count: u32,
    pub totals_row_shown: bool,
    pub autofilter: bool,
}

/// Growth of an existing table part, resolved by table name through the
/// worksheet's table relationships. The top-left anchor and the existing
/// columns stay fixed: `ref_range` replaces the table range,
/// `auto_filter_ref` replaces the range of the table's own `<autoFilter>`
/// when it has one, and `appended_columns` become new `<tableColumn>`
/// entries after the existing ones, with ids allocated from the part.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TableGrowthPatch {
    pub table_name: String,
    pub ref_range: String,
    pub auto_filter_ref: String,
    pub appended_columns: Vec<String>,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct TableStylePatch {
    pub name: String,
    pub show_first_column: bool,
    pub show_last_column: bool,
    pub show_row_stripes: bool,
    pub show_column_stripes: bool,
}

/// Output of [`build_tables`] — fresh ZIP parts, the merged
/// `<tableParts>` block bytes, allocated rels entries, and CT ops.
#[derive(Debug, Clone, Default)]
pub struct TablesResult {
    /// New `xl/tables/tableN.xml` entries: `(zip_path, xml_bytes)`.
    pub table_parts: Vec<(String, Vec<u8>)>,
    /// `<tableParts count="N">…</tableParts>` block bytes — feeds
    /// `SheetBlock::TableParts`. Empty when there were no patches AND
    /// no pre-existing rIds were merged in (the caller decides whether
    /// to skip pushing a SheetBlock based on this).
    pub table_parts_block: Vec<u8>,
    /// `(rId, abs_zip_path_of_part)` pairs — caller folds these into
    /// the sheet's rels graph at the rels level *and* into the
    /// `<tablePart>` block at the sheet-XML level.
    pub new_rels: Vec<(RelId, String)>,
    /// `Override` entries to inject into `[Content_Types].xml`.
    /// `(part_name, content_type)` — `part_name` is leading-slashed
    /// per OPC convention (`"/xl/tables/tableN.xml"`).
    pub new_content_types: Vec<(String, String)>,
}

const TABLE_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";

// ---------------------------------------------------------------------------
// ZIP scan: existing-table inventory
// ---------------------------------------------------------------------------

/// Workbook-level inventory derived from scanning every
/// `xl/tables/table*.xml` part in the source ZIP.
///
/// `ids` is the set of `id` attributes already in use (so a new
/// allocation can find the lowest free integer and stay
/// workbook-unique even when the source ZIP has gaps — e.g. tables
/// numbered 1 and 3 with no 2). `names` is the set of `name`
/// attributes already in use (for collision detection). `count` is
/// the total number of `tableN.xml` parts (used to compute the next
/// file-index, which is independent of `id` because the spec allows
/// the two to drift).
#[derive(Debug, Clone, Default)]
pub struct ExistingTablesInventory {
    pub ids: HashSet<u32>,
    pub names: HashSet<String>,
    pub count: u32,
    /// All `xl/tables/table*.xml` paths seen, sorted lexicographically.
    /// Reserved for callers that need the source-side list (none in
    /// the current slice; tests use it for assertions).
    pub paths: Vec<String>,
}

/// Walk every `xl/tables/table*.xml` entry in the source ZIP, parse
/// the `id` and `name` attributes off the root `<table>` element, and
/// return the aggregated inventory.
///
/// Tolerant of missing or malformed parts: anything that fails to
/// parse contributes its path to `paths` (and bumps `count`) but is
/// otherwise skipped. The caller's `next_available_id` allocation
/// then ignores those malformed slots — which is fine because the
/// failing parts will be copied through verbatim by the patcher's
/// main rewrite loop.
pub fn scan_existing_tables<R: std::io::Read + std::io::Seek>(
    zip: &mut ZipArchive<R>,
) -> Result<ExistingTablesInventory, String> {
    let mut paths: Vec<String> = Vec::new();
    for i in 0..zip.len() {
        let entry = zip.by_index(i).map_err(|e| format!("zip index {i}: {e}"))?;
        let name = entry.name().to_string();
        if is_table_part_path(&name) {
            paths.push(name);
        }
    }
    paths.sort();

    let mut ids: HashSet<u32> = HashSet::new();
    let mut names: HashSet<String> = HashSet::new();
    for p in &paths {
        let mut entry = match zip.by_name(p) {
            Ok(e) => e,
            Err(_) => continue,
        };
        let mut buf = Vec::with_capacity(entry.size() as usize);
        if std::io::Read::read_to_end(&mut entry, &mut buf).is_err() {
            continue;
        }
        let (id_opt, name_opt) = parse_table_root_attrs(&buf);
        if let Some(id) = id_opt {
            ids.insert(id);
        }
        if let Some(n) = name_opt {
            names.insert(n);
        }
    }

    let count = paths.len() as u32;
    Ok(ExistingTablesInventory {
        ids,
        names,
        count,
        paths,
    })
}

/// Match `xl/tables/table{n}.xml` (decimal index, no extra path segments).
fn is_table_part_path(p: &str) -> bool {
    let prefix = "xl/tables/table";
    let suffix = ".xml";
    if !p.starts_with(prefix) || !p.ends_with(suffix) {
        return false;
    }
    let middle = &p[prefix.len()..p.len() - suffix.len()];
    !middle.is_empty() && middle.chars().all(|c| c.is_ascii_digit())
}

/// Parse the `id` and `name` attributes off the root `<table>` element.
/// Tolerant: returns `(None, None)` on any malformation. Public so
/// the patcher can re-derive the id of a freshly-emitted part without
/// duplicating the parser.
pub fn parse_table_root_attrs(xml: &[u8]) -> (Option<u32>, Option<String>) {
    let mut reader = XmlReader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut buf: Vec<u8> = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e))
                if e.local_name().as_ref() == b"table" =>
            {
                let mut id: Option<u32> = None;
                let mut name: Option<String> = None;
                for a in e.attributes().with_checks(false).flatten() {
                    let value = a
                        .unescape_value()
                        .map(|v| v.into_owned())
                        .unwrap_or_else(|_| String::from_utf8_lossy(a.value.as_ref()).into_owned());
                    match a.key.as_ref() {
                        b"id" => id = value.parse::<u32>().ok(),
                        b"name" => name = Some(value),
                        _ => {}
                    }
                }
                return (id, name);
            }
            Ok(Event::Eof) => return (None, None),
            Err(_) => return (None, None),
            _ => {}
        }
        buf.clear();
    }
}

/// Resolve the table part named `table_name` among the table
/// relationships of the worksheet at `sheet_path` in the source package.
pub fn find_table_part_by_name<R: std::io::Read + std::io::Seek>(
    zip: &mut ZipArchive<R>,
    sheet_path: &str,
    table_name: &str,
) -> Result<Option<String>, String> {
    let rels_path = wolfxl_rels::rels_path_for(sheet_path)
        .ok_or_else(|| format!("no relationship part for '{sheet_path}'"))?;
    let rels_xml = match read_zip_entry(zip, &rels_path)? {
        Some(bytes) => bytes,
        None => return Ok(None),
    };
    let sheet_dir = sheet_path.rsplit_once('/').map_or("", |(dir, _)| dir);
    let graph =
        RelsGraph::parse(&rels_xml).map_err(|error| format!("parse '{rels_path}': {error}"))?;
    for rel in graph.find_by_type(rt::TABLE) {
        let part_path = wolfxl_rels::resolve_target(sheet_dir, &rel.target);
        let Some(table_xml) = read_zip_entry(zip, &part_path)? else {
            continue;
        };
        if parse_table_root_attrs(&table_xml).1.as_deref() == Some(table_name) {
            return Ok(Some(part_path));
        }
    }
    Ok(None)
}

fn read_zip_entry<R: std::io::Read + std::io::Seek>(
    zip: &mut ZipArchive<R>,
    path: &str,
) -> Result<Option<Vec<u8>>, String> {
    let mut entry = match zip.by_name(path) {
        Ok(entry) => entry,
        Err(_) => return Ok(None),
    };
    let mut bytes = Vec::with_capacity(entry.size() as usize);
    std::io::Read::read_to_end(&mut entry, &mut bytes)
        .map_err(|error| format!("read '{path}': {error}"))?;
    Ok(Some(bytes))
}

/// Apply a [`TableGrowthPatch`] to table part XML. Only the table `ref`,
/// the direct `<autoFilter ref>`, the `<tableColumns count>`, and the
/// appended `<tableColumn>` elements change; every other byte is kept.
pub fn patch_table_growth(xml: &[u8], patch: &TableGrowthPatch) -> Result<Vec<u8>, String> {
    let mut reader = XmlReader::from_reader(xml);
    let mut buf = Vec::new();
    let mut depth = 0usize;
    let mut saw_table = false;
    let mut in_columns = false;
    // Start-tag span and element prefix (e.g. `x:`) of `<tableColumns>`.
    let mut columns_tag: Option<(usize, usize, Vec<u8>)> = None;
    let mut columns_close: Option<usize> = None;
    let mut existing_names: Vec<String> = Vec::new();
    let mut max_id = 0u32;
    let mut replacements: Vec<(usize, usize, Vec<u8>)> = Vec::new();

    loop {
        let pre = reader.buffer_position() as usize;
        let event = reader
            .read_event_into(&mut buf)
            .map_err(|error| format!("parse table XML: {error}"))?;
        let post = reader.buffer_position() as usize;
        match event {
            Event::Start(element) if depth == 0 => {
                if element.local_name().as_ref() != b"table" {
                    return Err("table root not found".to_string());
                }
                let (start, end) = attribute_value_range(&xml[pre..post], b"ref")?
                    .ok_or_else(|| "table root has no ref attribute".to_string())?;
                replacements.push((
                    pre + start,
                    pre + end,
                    escape_xml_attribute(&patch.ref_range),
                ));
                saw_table = true;
                depth = 1;
            }
            Event::Empty(_) if depth == 0 => {
                return Err("table part has no columns".to_string());
            }
            Event::Start(element) => {
                let local = element.local_name();
                if depth == 1 && local.as_ref() == b"autoFilter" {
                    replace_ref(&mut replacements, xml, pre, post, &patch.auto_filter_ref)?;
                } else if depth == 1 && local.as_ref() == b"tableColumns" {
                    let qualified = element.name();
                    let prefix_len = qualified.as_ref().len() - local.as_ref().len();
                    columns_tag = Some((pre, post, qualified.as_ref()[..prefix_len].to_vec()));
                    in_columns = true;
                } else if depth == 2 && in_columns && local.as_ref() == b"tableColumn" {
                    record_table_column(&element, &mut existing_names, &mut max_id)?;
                }
                depth += 1;
            }
            Event::Empty(element) => {
                let local = element.local_name();
                if depth == 1 && local.as_ref() == b"autoFilter" {
                    replace_ref(&mut replacements, xml, pre, post, &patch.auto_filter_ref)?;
                } else if depth == 2 && in_columns && local.as_ref() == b"tableColumn" {
                    record_table_column(&element, &mut existing_names, &mut max_id)?;
                }
            }
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if depth == 1 && in_columns {
                    in_columns = false;
                    columns_close = Some(pre);
                }
                if depth == 0 {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    if !saw_table {
        return Err("table root not found".to_string());
    }
    if !patch.appended_columns.is_empty() {
        let (Some((tag_start, tag_end, prefix)), Some(close)) = (columns_tag, columns_close) else {
            return Err("table has no <tableColumns> element".to_string());
        };
        let mut seen: HashSet<String> = existing_names.iter().map(|n| n.to_lowercase()).collect();
        let mut inserted = Vec::new();
        for (offset, name) in patch.appended_columns.iter().enumerate() {
            if name.is_empty() {
                return Err("table column names must not be empty".to_string());
            }
            if !seen.insert(name.to_lowercase()) {
                return Err(format!("table already has a column named '{name}'"));
            }
            let id = max_id + 1 + offset as u32;
            inserted.extend_from_slice(b"<");
            inserted.extend_from_slice(&prefix);
            inserted.extend_from_slice(format!("tableColumn id=\"{id}\" name=\"").as_bytes());
            inserted.extend(escape_xml_attribute(name));
            inserted.extend_from_slice(b"\"/>");
        }
        let count = existing_names.len() + patch.appended_columns.len();
        if let Some((start, end)) = attribute_value_range(&xml[tag_start..tag_end], b"count")? {
            replacements.push((
                tag_start + start,
                tag_start + end,
                count.to_string().into_bytes(),
            ));
        }
        replacements.push((close, close, inserted));
    }

    replacements.sort_unstable_by_key(|(start, _, _)| *start);
    let mut output = xml.to_vec();
    for (start, end, replacement) in replacements.into_iter().rev() {
        output.splice(start..end, replacement);
    }
    Ok(output)
}

fn replace_ref(
    replacements: &mut Vec<(usize, usize, Vec<u8>)>,
    xml: &[u8],
    pre: usize,
    post: usize,
    value: &str,
) -> Result<(), String> {
    let (start, end) = attribute_value_range(&xml[pre..post], b"ref")?
        .ok_or_else(|| "table autoFilter has no ref attribute".to_string())?;
    replacements.push((pre + start, pre + end, escape_xml_attribute(value)));
    Ok(())
}

fn record_table_column(
    element: &quick_xml::events::BytesStart<'_>,
    names: &mut Vec<String>,
    max_id: &mut u32,
) -> Result<(), String> {
    for attr in element.attributes().with_checks(false).flatten() {
        match attr.key.as_ref() {
            b"id" => {
                let id = std::str::from_utf8(attr.value.as_ref())
                    .ok()
                    .and_then(|value| value.trim().parse::<u32>().ok())
                    .ok_or_else(|| "tableColumn has a non-numeric id".to_string())?;
                *max_id = (*max_id).max(id);
            }
            b"name" => names.push(
                attr.unescape_value()
                    .map_err(|error| format!("tableColumn name: {error}"))?
                    .into_owned(),
            ),
            _ => {}
        }
    }
    Ok(())
}

/// Byte range of the value of attribute `name` within one start tag.
fn attribute_value_range(tag: &[u8], name: &[u8]) -> Result<Option<(usize, usize)>, String> {
    let mut cursor = 1usize;
    while cursor < tag.len() && !tag[cursor].is_ascii_whitespace() && tag[cursor] != b'>' {
        cursor += 1;
    }
    while cursor < tag.len() {
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= tag.len() || tag[cursor] == b'/' || tag[cursor] == b'>' {
            return Ok(None);
        }
        let name_start = cursor;
        while cursor < tag.len()
            && !tag[cursor].is_ascii_whitespace()
            && !matches!(tag[cursor], b'=' | b'/' | b'>')
        {
            cursor += 1;
        }
        let name_end = cursor;
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= tag.len() || tag[cursor] != b'=' {
            return Err("malformed XML attribute".to_string());
        }
        cursor += 1;
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        let quote = *tag
            .get(cursor)
            .ok_or_else(|| "unterminated XML attribute".to_string())?;
        if quote != b'\'' && quote != b'"' {
            return Err("unquoted XML attribute".to_string());
        }
        cursor += 1;
        let value_start = cursor;
        while cursor < tag.len() && tag[cursor] != quote {
            cursor += 1;
        }
        if cursor == tag.len() {
            return Err("unterminated XML attribute".to_string());
        }
        let value_end = cursor;
        cursor += 1;
        if &tag[name_start..name_end] == name {
            return Ok(Some((value_start, value_end)));
        }
    }
    Ok(None)
}

fn escape_xml_attribute(value: &str) -> Vec<u8> {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('"', "&quot;")
        .into_bytes()
}

/// Lowest positive integer `i` such that `i` is not in `ids`.
fn next_available_id(ids: &HashSet<u32>) -> u32 {
    let mut i: u32 = 1;
    while ids.contains(&i) {
        i += 1;
    }
    i
}

// ---------------------------------------------------------------------------
// build_tables
// ---------------------------------------------------------------------------

/// Build the full `TablesResult` for one worksheet's queued table
/// patches.
///
/// `existing_ids` and `existing_names` MUST cover every table already
/// present in the workbook (across all sheets), not just this sheet —
/// table `id` and `name` are workbook-unique, not sheet-unique.
///
/// `existing_global_count` is the total number of `xl/tables/tableN.xml`
/// parts in the source ZIP. It seeds the next part index so the new
/// part filenames don't collide with existing files. Note: the part
/// index (filename) and the `id` attribute can drift (e.g. if the
/// source has `table1.xml` with `id=3` and `table2.xml` with `id=7`,
/// adding a new table produces `table3.xml` with `id=1` — first
/// available id, next sequential filename).
///
/// Returns `Err` on the first patch whose `name` collides with an
/// already-taken name.
///
/// `rels` is mutated as a side effect: one new TABLE rel is allocated
/// per patch, with `Target` pointing at the new part path relative to
/// the sheet's rels parent (e.g. `"../tables/table5.xml"`).
///
/// When `allocator` is `Some`, `tableN.xml` part filenames are minted
/// via [`PartIdAllocator::alloc_table`] (centralized; shared with
/// RFC-035 sheet copy and other subsystems that mint part suffixes in
/// the same save). When `None`, falls back to the legacy
/// `inventory.count + 1` allocation — used by tests and any caller
/// that has not yet adopted the centralized allocator.
pub fn build_tables(
    patches: &[TablePatch],
    inventory: &ExistingTablesInventory,
    rels: &mut RelsGraph,
    allocator: Option<&mut PartIdAllocator>,
) -> Result<TablesResult, String> {
    let mut result = TablesResult::default();
    if patches.is_empty() {
        return Ok(result);
    }

    // Local clones we update as we go so allocations across the same
    // batch don't collide with each other.
    let mut ids = inventory.ids.clone();
    let mut names = inventory.names.clone();
    // Fallback path: when no allocator is supplied, advance a local
    // counter from `inventory.count + 1` (legacy behaviour preserved
    // for tests and callers not yet migrated). Otherwise, every
    // `part_idx` comes from the shared allocator.
    let mut legacy_next_part_idx = inventory.count + 1;
    let mut allocator = allocator;

    for patch in patches {
        if names.contains(&patch.name) {
            return Err(format!("Table with name {:?} already exists", patch.name));
        }
        // Allocate workbook-unique id (lowest free) and a part-filename
        // index (centralized via PartIdAllocator if supplied; legacy
        // count-based fallback otherwise).
        let new_id = next_available_id(&ids);
        ids.insert(new_id);
        names.insert(patch.name.clone());
        let part_idx = match allocator.as_deref_mut() {
            Some(alloc) => alloc.alloc_table(),
            None => {
                let n = legacy_next_part_idx;
                legacy_next_part_idx += 1;
                n
            }
        };

        // Build the writer-model Table and serialize via the writer's
        // emitter (so write-mode and modify-mode produce byte-identical
        // parts for the same input).
        let table = patch_to_writer_table(patch);
        // The writer's emit signature treats `table_idx` as `id - 1`.
        // We pass `(new_id - 1)` so the emitted XML carries our
        // workbook-unique id.
        let xml_bytes = writer_tables_xml::emit(&table, 0, (new_id - 1) as usize);

        let part_path = format!("xl/tables/table{}.xml", part_idx);
        result.table_parts.push((part_path.clone(), xml_bytes));

        // Allocate a fresh rId for this rel. RelsGraph::add picks the
        // next monotonic rId based on the graph's existing entries.
        let target = format!("../tables/table{}.xml", part_idx);
        let rid = rels.add(rt::TABLE, &target, TargetMode::Internal);
        result.new_rels.push((rid, part_path.clone()));

        result
            .new_content_types
            .push((format!("/{}", part_path), TABLE_CONTENT_TYPE.to_string()));
    }

    // Build the <tableParts> block. We include rIds for ALL TABLE
    // relationships in the rels graph so any pre-existing tables on
    // this sheet remain referenced. The patcher's merger uses
    // replace-style insertion at slot 37 — see RFC-011 / the merger
    // tests — so any existing block in the source XML is dropped and
    // replaced with whatever we emit here.
    let table_rids: Vec<RelId> = rels
        .iter()
        .filter(|r| r.rel_type == rt::TABLE)
        .map(|r| r.id.clone())
        .collect();
    if !table_rids.is_empty() {
        let mut block = String::with_capacity(48 + table_rids.len() * 32);
        block.push_str(&format!("<tableParts count=\"{}\">", table_rids.len()));
        for rid in &table_rids {
            block.push_str(&format!("<tablePart r:id=\"{}\"/>", rid.0));
        }
        block.push_str("</tableParts>");
        result.table_parts_block = block.into_bytes();
    }

    Ok(result)
}

/// Convert a `TablePatch` into the writer's `Table` model so we can
/// reuse `wolfxl_writer::emit::tables_xml::emit` directly.
fn patch_to_writer_table(p: &TablePatch) -> Table {
    let columns: Vec<TableColumn> = p
        .columns
        .iter()
        .map(|n| TableColumn {
            name: n.clone(),
            totals_function: None,
            totals_label: None,
        })
        .collect();
    let style = p.style.as_ref().map(|s| TableStyle {
        name: s.name.clone(),
        show_first_column: s.show_first_column,
        show_last_column: s.show_last_column,
        show_row_stripes: s.show_row_stripes,
        show_column_stripes: s.show_column_stripes,
    });
    Table {
        name: p.name.clone(),
        // The writer's emitter falls back to `name` when `display_name`
        // is `None`. We always carry an explicit `display_name` (even
        // when equal to `name`) so the output XML matches openpyxl's
        // behavior of emitting both attributes.
        display_name: Some(if p.display_name.is_empty() {
            p.name.clone()
        } else {
            p.display_name.clone()
        }),
        range: p.ref_range.clone(),
        columns,
        header_row: p.header_row_count > 0,
        totals_row: p.totals_row_shown,
        style,
        autofilter: p.autofilter,
    }
}

// ---------------------------------------------------------------------------
// Tests
//
// Pure-Rust tests that compile under `cargo build` but not under
// `cargo test -p wolfxl --lib` (the cdylib links Python). End-to-end
// patcher behavior is exercised via pytest in `tests/test_tables_modify.py`.
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;

    fn simple_patch(name: &str, range: &str, cols: &[&str]) -> TablePatch {
        TablePatch {
            name: name.to_string(),
            display_name: name.to_string(),
            ref_range: range.to_string(),
            columns: cols.iter().map(|s| s.to_string()).collect(),
            style: None,
            header_row_count: 1,
            totals_row_shown: false,
            autofilter: true,
        }
    }

    #[test]
    fn next_available_id_handles_gaps() {
        let mut s = HashSet::new();
        s.insert(1);
        s.insert(3);
        assert_eq!(next_available_id(&s), 2);
        s.insert(2);
        assert_eq!(next_available_id(&s), 4);
    }

    #[test]
    fn next_available_id_starts_at_1_for_empty() {
        let s = HashSet::new();
        assert_eq!(next_available_id(&s), 1);
    }

    #[test]
    fn is_table_part_path_recognizes_canonical_paths() {
        assert!(is_table_part_path("xl/tables/table1.xml"));
        assert!(is_table_part_path("xl/tables/table42.xml"));
        // No extra path segments
        assert!(!is_table_part_path("xl/tables/sub/table1.xml"));
        // Must end in digits
        assert!(!is_table_part_path("xl/tables/tableA.xml"));
        // Other parts of the workbook
        assert!(!is_table_part_path("xl/tables/_rels/table1.xml.rels"));
        assert!(!is_table_part_path("xl/worksheets/sheet1.xml"));
    }

    #[test]
    fn parse_table_root_attrs_extracts_id_and_name() {
        let xml = br#"<?xml version="1.0"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
       id="7" name="MyTable" displayName="MyTable" ref="A1:B5"/>"#;
        let (id, name) = parse_table_root_attrs(xml);
        assert_eq!(id, Some(7));
        assert_eq!(name.as_deref(), Some("MyTable"));
    }

    #[test]
    fn build_tables_empty_patches_yields_empty_result() {
        let inventory = ExistingTablesInventory::default();
        let mut rels = RelsGraph::new();
        let r = build_tables(&[], &inventory, &mut rels, None).unwrap();
        assert!(r.table_parts.is_empty());
        assert!(r.table_parts_block.is_empty());
        assert!(r.new_rels.is_empty());
        assert!(r.new_content_types.is_empty());
    }

    #[test]
    fn build_tables_clean_file_allocates_id_1_and_part_1() {
        let inventory = ExistingTablesInventory::default();
        let mut rels = RelsGraph::new();
        let patches = vec![simple_patch("Sales", "A1:C10", &["Region", "Q1", "Q2"])];
        let r = build_tables(&patches, &inventory, &mut rels, None).unwrap();
        assert_eq!(r.table_parts.len(), 1);
        assert_eq!(r.table_parts[0].0, "xl/tables/table1.xml");
        let xml = std::str::from_utf8(&r.table_parts[0].1).unwrap();
        assert!(xml.contains("id=\"1\""), "{xml}");
        assert!(xml.contains("name=\"Sales\""), "{xml}");
        assert!(xml.contains("ref=\"A1:C10\""), "{xml}");
        // CT override
        assert_eq!(
            r.new_content_types[0],
            (
                "/xl/tables/table1.xml".to_string(),
                TABLE_CONTENT_TYPE.to_string(),
            )
        );
        // Rels graph mutated
        assert_eq!(rels.len(), 1);
        // Block bytes carry exactly one tablePart
        let block = std::str::from_utf8(&r.table_parts_block).unwrap();
        assert!(block.starts_with("<tableParts count=\"1\">"), "{block}");
        assert!(block.contains("<tablePart r:id=\"rId1\"/>"), "{block}");
    }

    #[test]
    fn build_tables_uses_lowest_free_id_with_gap() {
        // Workbook already has tables with ids 1 and 3 — new table must
        // pick id 2, not id 4.
        let mut inventory = ExistingTablesInventory::default();
        inventory.ids.insert(1);
        inventory.ids.insert(3);
        inventory.count = 2;
        let mut rels = RelsGraph::new();
        let patches = vec![simple_patch("New", "D1:D5", &["X"])];
        let r = build_tables(&patches, &inventory, &mut rels, None).unwrap();
        let xml = std::str::from_utf8(&r.table_parts[0].1).unwrap();
        assert!(xml.contains("id=\"2\""), "expected id=2, got: {xml}");
        // Part filename is sequential, not id-mapped
        assert_eq!(r.table_parts[0].0, "xl/tables/table3.xml");
    }

    #[test]
    fn build_tables_name_collision_errors() {
        let mut inventory = ExistingTablesInventory::default();
        inventory.names.insert("Existing".into());
        let mut rels = RelsGraph::new();
        let patches = vec![simple_patch("Existing", "A1:B2", &["A", "B"])];
        let err = build_tables(&patches, &inventory, &mut rels, None).unwrap_err();
        assert!(err.contains("Existing"), "{err}");
    }

    #[test]
    fn build_tables_two_patches_get_distinct_ids_and_paths() {
        let inventory = ExistingTablesInventory::default();
        let mut rels = RelsGraph::new();
        let patches = vec![
            simple_patch("T1", "A1:A5", &["A"]),
            simple_patch("T2", "B1:B5", &["B"]),
        ];
        let r = build_tables(&patches, &inventory, &mut rels, None).unwrap();
        assert_eq!(r.table_parts.len(), 2);
        assert_eq!(r.table_parts[0].0, "xl/tables/table1.xml");
        assert_eq!(r.table_parts[1].0, "xl/tables/table2.xml");
        let xml1 = std::str::from_utf8(&r.table_parts[0].1).unwrap();
        let xml2 = std::str::from_utf8(&r.table_parts[1].1).unwrap();
        assert!(xml1.contains("id=\"1\""));
        assert!(xml2.contains("id=\"2\""));
        // Two distinct rIds in the graph
        assert_eq!(rels.len(), 2);
        let block = std::str::from_utf8(&r.table_parts_block).unwrap();
        assert!(block.contains("<tableParts count=\"2\">"), "{block}");
        assert_eq!(block.matches("<tablePart ").count(), 2);
    }

    #[test]
    fn build_tables_preserves_pre_existing_table_rids_in_block() {
        // A sheet already has one table (rId1 → ../tables/table1.xml).
        // We add a second table on the same sheet — the merged block
        // must carry both rIds so the existing table doesn't disappear.
        let mut inventory = ExistingTablesInventory::default();
        inventory.ids.insert(1);
        inventory.names.insert("Existing".into());
        inventory.count = 1;

        let mut rels = RelsGraph::new();
        rels.add_with_id(
            RelId("rId1".into()),
            rt::TABLE,
            "../tables/table1.xml",
            TargetMode::Internal,
        );

        let patches = vec![simple_patch("Added", "C1:C5", &["X"])];
        let r = build_tables(&patches, &inventory, &mut rels, None).unwrap();

        // Two TABLE rels in the graph now.
        let table_rels: Vec<_> = rels.iter().filter(|r| r.rel_type == rt::TABLE).collect();
        assert_eq!(table_rels.len(), 2);

        // Block has two <tablePart/> entries — one for the pre-existing
        // table, one for the new one.
        let block = std::str::from_utf8(&r.table_parts_block).unwrap();
        assert!(block.contains("<tableParts count=\"2\">"), "{block}");
        assert!(block.contains("rId1"));
        // The new rel is rId2 (the next monotonic id RelsGraph::add picked)
        assert!(block.contains("rId2"));
    }

    #[test]
    fn build_tables_with_style_propagates_to_xml() {
        let inventory = ExistingTablesInventory::default();
        let mut rels = RelsGraph::new();
        let mut patch = simple_patch("Styled", "A1:C5", &["A", "B", "C"]);
        patch.style = Some(TableStylePatch {
            name: "TableStyleLight1".to_string(),
            show_first_column: true,
            show_last_column: false,
            show_row_stripes: false,
            show_column_stripes: true,
        });
        let r = build_tables(&[patch], &inventory, &mut rels, None).unwrap();
        let xml = std::str::from_utf8(&r.table_parts[0].1).unwrap();
        assert!(xml.contains("name=\"TableStyleLight1\""), "{xml}");
        assert!(xml.contains("showFirstColumn=\"1\""), "{xml}");
        assert!(xml.contains("showColumnStripes=\"1\""), "{xml}");
    }

    #[test]
    fn build_tables_displayname_defaults_to_name() {
        let inventory = ExistingTablesInventory::default();
        let mut rels = RelsGraph::new();
        let mut patch = simple_patch("Foo", "A1:B5", &["A", "B"]);
        patch.display_name = String::new(); // simulate "not provided"
        let r = build_tables(&[patch], &inventory, &mut rels, None).unwrap();
        let xml = std::str::from_utf8(&r.table_parts[0].1).unwrap();
        assert!(xml.contains("displayName=\"Foo\""), "{xml}");
    }

    #[test]
    fn build_tables_autofilter_false_omits_element() {
        let inventory = ExistingTablesInventory::default();
        let mut rels = RelsGraph::new();
        let mut patch = simple_patch("NoFilter", "A1:B5", &["A", "B"]);
        patch.autofilter = false;
        let r = build_tables(&[patch], &inventory, &mut rels, None).unwrap();
        let xml = std::str::from_utf8(&r.table_parts[0].1).unwrap();
        assert!(!xml.contains("<autoFilter"), "{xml}");
    }

    #[test]
    fn build_tables_totals_row_sets_attr() {
        let inventory = ExistingTablesInventory::default();
        let mut rels = RelsGraph::new();
        let mut patch = simple_patch("Totals", "A1:B10", &["A", "B"]);
        patch.totals_row_shown = true;
        let r = build_tables(&[patch], &inventory, &mut rels, None).unwrap();
        let xml = std::str::from_utf8(&r.table_parts[0].1).unwrap();
        assert!(xml.contains("totalsRowShown=\"1\""), "{xml}");
    }

    // -----------------------------------------------------------------
    // ZIP scanner tests — synthesize a small in-memory ZIP and verify
    // `scan_existing_tables` extracts the right inventory.
    // -----------------------------------------------------------------

    use std::io::{Cursor, Write};
    use zip::write::SimpleFileOptions;
    use zip::ZipWriter;

    fn make_zip(entries: &[(&str, &[u8])]) -> Vec<u8> {
        let mut buf = Cursor::new(Vec::new());
        {
            let mut writer = ZipWriter::new(&mut buf);
            let opts = SimpleFileOptions::default();
            for (path, bytes) in entries {
                writer.start_file(*path, opts).unwrap();
                writer.write_all(bytes).unwrap();
            }
            writer.finish().unwrap();
        }
        buf.into_inner()
    }

    fn open_zip(bytes: Vec<u8>) -> ZipArchive<Cursor<Vec<u8>>> {
        ZipArchive::new(Cursor::new(bytes)).unwrap()
    }

    #[test]
    fn scan_no_tables_yields_empty_inventory() {
        let z = make_zip(&[("xl/workbook.xml", b"<wb/>")]);
        let mut archive = open_zip(z);
        let inv = scan_existing_tables(&mut archive).unwrap();
        assert_eq!(inv.count, 0);
        assert!(inv.ids.is_empty());
        assert!(inv.names.is_empty());
        assert!(inv.paths.is_empty());
    }

    #[test]
    fn scan_two_tables_extracts_ids_and_names() {
        let t1 = br#"<?xml version="1.0"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
       id="1" name="Alpha" displayName="Alpha" ref="A1:A5"/>"#;
        let t2 = br#"<?xml version="1.0"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
       id="3" name="Bravo" displayName="Bravo" ref="B1:B5"/>"#;
        let z = make_zip(&[
            ("xl/tables/table1.xml", t1),
            ("xl/tables/table2.xml", t2),
            ("xl/workbook.xml", b"<wb/>"),
        ]);
        let mut archive = open_zip(z);
        let inv = scan_existing_tables(&mut archive).unwrap();
        assert_eq!(inv.count, 2);
        assert!(inv.ids.contains(&1));
        assert!(inv.ids.contains(&3));
        assert!(inv.names.contains("Alpha"));
        assert!(inv.names.contains("Bravo"));
        // Ids 2 should be free for next allocation
        assert_eq!(next_available_id(&inv.ids), 2);
    }

    #[test]
    fn scan_skips_non_table_paths() {
        let z = make_zip(&[
            ("xl/tables/_rels/table1.xml.rels", b"<rels/>"),
            ("xl/worksheets/sheet1.xml", b"<sheet/>"),
        ]);
        let mut archive = open_zip(z);
        let inv = scan_existing_tables(&mut archive).unwrap();
        assert_eq!(inv.count, 0);
    }
    fn growth(ref_range: &str, filter: &str, cols: &[&str]) -> TableGrowthPatch {
        TableGrowthPatch {
            table_name: "T".to_string(),
            ref_range: ref_range.to_string(),
            auto_filter_ref: filter.to_string(),
            appended_columns: cols.iter().map(|c| c.to_string()).collect(),
        }
    }

    #[test]
    fn table_growth_keeps_prefix_and_column_children() {
        let xml = br#"<?xml version="1.0"?><x:table xmlns:x="urn:x" id="3" ref="A1:B3" name="T"><x:autoFilter ref="A1:B3"/><x:tableColumns count="2"><x:tableColumn id="1" name="A"/><x:tableColumn id="5" name="B"><x:calculatedColumnFormula>A2*2</x:calculatedColumnFormula></x:tableColumn></x:tableColumns><x:tableStyleInfo name="S"/></x:table>"#;
        let out = patch_table_growth(xml, &growth("A1:C4", "A1:C4", &["C&D"])).unwrap();
        let out = String::from_utf8(out).unwrap();
        assert_eq!(
            out,
            r#"<?xml version="1.0"?><x:table xmlns:x="urn:x" id="3" ref="A1:C4" name="T"><x:autoFilter ref="A1:C4"/><x:tableColumns count="3"><x:tableColumn id="1" name="A"/><x:tableColumn id="5" name="B"><x:calculatedColumnFormula>A2*2</x:calculatedColumnFormula></x:tableColumn><x:tableColumn id="6" name="C&amp;D"/></x:tableColumns><x:tableStyleInfo name="S"/></x:table>"#
        );
    }

    #[test]
    fn table_growth_without_autofilter_adds_none() {
        let xml = br#"<table id="1" ref="A1:A2" name="T"><tableColumns count="1"><tableColumn id="1" name="A"/></tableColumns></table>"#;
        let out = patch_table_growth(xml, &growth("A1:A9", "A1:A9", &[])).unwrap();
        assert_eq!(
            out,
            br#"<table id="1" ref="A1:A9" name="T"><tableColumns count="1"><tableColumn id="1" name="A"/></tableColumns></table>"#
        );
    }

    #[test]
    fn table_growth_rejects_duplicate_column_name() {
        let xml = br#"<table id="1" ref="A1:A2" name="T"><tableColumns count="1"><tableColumn id="1" name="Sales"/></tableColumns></table>"#;
        let err = patch_table_growth(xml, &growth("A1:B2", "A1:B2", &["sales"])).unwrap_err();
        assert!(err.contains("already has a column named"), "{err}");
    }
}
