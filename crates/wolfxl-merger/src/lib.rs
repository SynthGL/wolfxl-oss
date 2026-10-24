//! Generic insertion/replacement of sibling blocks inside a `CT_Worksheet`
//! XML stream, preserving ECMA-376 §18.3.1.99 child-element order and passing
//! unknown elements (`extLst`, `x14ac`, future MS extensions, third-party
//! compat tags) through verbatim.
//!
//! See `Plans/rfcs/011-xml-block-merger.md` for the full design rationale.
//!
//! # Invariants
//!
//! - **Streaming and memory-bounded.** [`merge_blocks`] does exactly one
//!   pass through the source via `quick_xml::Reader`, writing every event
//!   verbatim except where it skips a block being replaced or injects a
//!   new one. Peak extra allocation is `O(input bytes)`, dominated by the
//!   output buffer; the merger never builds a DOM. A 50 MB sheet produces
//!   well under 4 MB peak heap above the input/output buffers (RFC-011
//!   §6 test #11).
//!
//! - **Replace-all semantics for `<conditionalFormatting>`.** Slot 17 is
//!   the only 0..N slot in CT_Worksheet. If any
//!   [`SheetBlock::ConditionalFormatting`] is in the supplied `blocks`,
//!   every existing `<conditionalFormatting>` element in the source is
//!   removed before the supplied list is inserted. Callers that want to
//!   preserve existing CF rules MUST read them out of the source first
//!   and re-include them in `blocks`. RFC-011 §5.5 / INDEX Q4
//!   (locked 2026-04-25).
//!
//! - **Verbatim pass-through for unknown elements.** Anything whose
//!   local-name is not in [`ct_worksheet_order::ECMA_ORDER`] flows through
//!   at its source position with attributes, prefix bindings, namespace
//!   declarations, and entity escaping byte-identical. This includes
//!   `extLst`-internal extensions (`x14:sparklineGroups`, icon-set
//!   extensions, mc:Ignorable compat tags), x14ac extension attributes,
//!   third-party tags, and any future MS schema extensions. The
//!   byte-preservation property is guarded by the
//!   `extlst_is_byte_preserved` test.
//!
//! - **Block payloads are opaque bytes.** The merger does NOT validate
//!   that supplied bytes are well-formed XML, that they start with the
//!   expected root element, or that any `r:id` references they contain
//!   correspond to relationships that exist in the rels graph. Each is
//!   the caller's responsibility — modify-mode RFCs (RFC-022/024/025/026)
//!   own block content; the merger only owns ordering. Garbage in →
//!   garbage out.
//!
//! - **Determinism.** With `WOLFXL_TEST_EPOCH=0` set, two saves of the
//!   same workbook with the same queued blocks produce byte-identical
//!   output. The merger has no time-dependent state; its determinism
//!   contract is "identical inputs ⇒ identical bytes". `BTreeMap`-keyed
//!   pending blocks ensure CF order matches caller insertion order.
//!
//! - **Single rewrite point.** The only place the merger ever rewrites
//!   a non-block element is the `<worksheet>` open tag, and only when
//!   needed to inject `xmlns:r=…` for a payload that uses the
//!   `r:` prefix without the source declaring it (RFC-011 §8 risk #4).
//!   Every other source element is passed through with its byte-slice
//!   intact.
//!
//! - **Idempotent on empty `blocks`.** `merge_blocks(xml, vec![])` short-
//!   circuits without invoking the parser; returns a `Vec` clone of the
//!   input bytes byte-identical to the source. Even malformed XML
//!   passes through.

use std::collections::BTreeMap;

use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader as XmlReader;
use quick_xml::Writer as XmlWriter;

const REL_NS: &[u8] = b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";

// ---------------------------------------------------------------------------
// SheetBlock — one sibling-block insertion.
// ---------------------------------------------------------------------------

/// One sibling-block insertion. The bytes are pre-serialized including the
/// wrapping element (e.g. `b"<hyperlinks><hyperlink ref=\"A1\" .../></hyperlinks>"`).
/// They MUST be UTF-8 and MUST be a valid XML fragment with one root element
/// at the top — the merger does NOT validate this; malformed input produces
/// malformed output.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum SheetBlock {
    /// `<sheetViews>…</sheetViews>` — slot 3 (RFC-055 §3.1).
    SheetViews(Vec<u8>),
    /// `<sheetFormatPr .../>` — slot 4 (RFC-062 §4).
    /// Sprint Π Pod Π-α. Replaces any existing `<sheetFormatPr>`.
    SheetFormatPr(Vec<u8>),
    /// `<sheetProtection .../>` — slot 8 (RFC-055 §3.1).
    SheetProtection(Vec<u8>),
    /// `<autoFilter ref="…">…</autoFilter>` — slot 11 in §18.3.1.99.
    /// Sprint Ο Pod 1B (RFC-056). Replaces any existing `<autoFilter>`.
    AutoFilter(Vec<u8>),
    /// `<mergeCells count="…">…</mergeCells>` — slot 15 in §18.3.1.99.
    MergeCells(Vec<u8>),
    /// `<conditionalFormatting sqref="…">…</conditionalFormatting>` (one per
    /// range; supply multiple `SheetBlock::ConditionalFormatting` to insert
    /// several). Slot 17.
    ///
    /// **Replace-all semantics:** if any `SheetBlock::ConditionalFormatting`
    /// is present in the `blocks` argument to [`merge_blocks`], every
    /// existing `<conditionalFormatting>` element in the source is removed
    /// before the supplied blocks are inserted. Callers that want to
    /// preserve existing CF rules MUST read them out of the source first
    /// and re-include them in the supplied set. See RFC-011 §5.5.
    ConditionalFormatting(Vec<u8>),
    /// `<dataValidations count="…">…</dataValidations>` — slot 18.
    DataValidations(Vec<u8>),
    /// `<hyperlinks>…</hyperlinks>` — slot 19.
    Hyperlinks(Vec<u8>),
    /// `<printOptions .../>` — slot 20 (RFC-055 / G24).
    PrintOptions(Vec<u8>),
    /// `<pageMargins .../>` — slot 21 (RFC-055 §3.1).
    PageMargins(Vec<u8>),
    /// `<pageSetup .../>` — slot 22 (RFC-055 §3.1).
    PageSetup(Vec<u8>),
    /// `<headerFooter>…</headerFooter>` — slot 23 (RFC-055 §3.1).
    HeaderFooter(Vec<u8>),
    /// `<rowBreaks count="…">…</rowBreaks>` — slot 24 (RFC-062 §4).
    /// Sprint Π Pod Π-α. Replaces any existing `<rowBreaks>`.
    RowBreaks(Vec<u8>),
    /// `<colBreaks count="…">…</colBreaks>` — slot 25 (RFC-062 §4).
    /// Sprint Π Pod Π-α. Replaces any existing `<colBreaks>`.
    ColBreaks(Vec<u8>),
    /// `<legacyDrawing r:id="…"/>` — slot 31. Empty-element form is canonical.
    LegacyDrawing(Vec<u8>),
    /// `<tableParts count="…">…</tableParts>` — slot 37.
    TableParts(Vec<u8>),
}

impl SheetBlock {
    /// The ECMA §18.3.1.99 ordinal (1..=38). Used to decide where to insert
    /// the block when not already present in the source.
    pub fn ecma_position(&self) -> u32 {
        match self {
            SheetBlock::SheetViews(_) => 3,
            SheetBlock::SheetFormatPr(_) => 4,
            SheetBlock::SheetProtection(_) => 8,
            SheetBlock::AutoFilter(_) => 11,
            SheetBlock::MergeCells(_) => 15,
            SheetBlock::ConditionalFormatting(_) => 17,
            SheetBlock::DataValidations(_) => 18,
            SheetBlock::Hyperlinks(_) => 19,
            SheetBlock::PrintOptions(_) => 20,
            SheetBlock::PageMargins(_) => 21,
            SheetBlock::PageSetup(_) => 22,
            SheetBlock::HeaderFooter(_) => 23,
            SheetBlock::RowBreaks(_) => 24,
            SheetBlock::ColBreaks(_) => 25,
            SheetBlock::LegacyDrawing(_) => 31,
            SheetBlock::TableParts(_) => 37,
        }
    }

    /// The local-name of the block's root element. Used to detect existing
    /// blocks for replacement (matched against `BytesStart::local_name()`).
    pub fn root_local_name(&self) -> &'static [u8] {
        match self {
            SheetBlock::SheetViews(_) => b"sheetViews",
            SheetBlock::SheetFormatPr(_) => b"sheetFormatPr",
            SheetBlock::SheetProtection(_) => b"sheetProtection",
            SheetBlock::AutoFilter(_) => b"autoFilter",
            SheetBlock::MergeCells(_) => b"mergeCells",
            SheetBlock::ConditionalFormatting(_) => b"conditionalFormatting",
            SheetBlock::DataValidations(_) => b"dataValidations",
            SheetBlock::Hyperlinks(_) => b"hyperlinks",
            SheetBlock::PrintOptions(_) => b"printOptions",
            SheetBlock::PageMargins(_) => b"pageMargins",
            SheetBlock::PageSetup(_) => b"pageSetup",
            SheetBlock::HeaderFooter(_) => b"headerFooter",
            SheetBlock::RowBreaks(_) => b"rowBreaks",
            SheetBlock::ColBreaks(_) => b"colBreaks",
            SheetBlock::LegacyDrawing(_) => b"legacyDrawing",
            SheetBlock::TableParts(_) => b"tableParts",
        }
    }

    /// The pre-serialized payload bytes the merger will emit verbatim at the
    /// chosen insertion point.
    pub fn bytes(&self) -> &[u8] {
        match self {
            SheetBlock::SheetViews(b)
            | SheetBlock::SheetFormatPr(b)
            | SheetBlock::SheetProtection(b)
            | SheetBlock::AutoFilter(b)
            | SheetBlock::MergeCells(b)
            | SheetBlock::ConditionalFormatting(b)
            | SheetBlock::DataValidations(b)
            | SheetBlock::Hyperlinks(b)
            | SheetBlock::PrintOptions(b)
            | SheetBlock::PageMargins(b)
            | SheetBlock::PageSetup(b)
            | SheetBlock::HeaderFooter(b)
            | SheetBlock::RowBreaks(b)
            | SheetBlock::ColBreaks(b)
            | SheetBlock::LegacyDrawing(b)
            | SheetBlock::TableParts(b) => b,
        }
    }
}

// ---------------------------------------------------------------------------
// ECMA-376 §18.3.1.99 child-element ordering table.
// ---------------------------------------------------------------------------

/// CT_Worksheet child element order per ECMA-376 Part 1 §18.3.1.99. Single
/// source of truth for both the streaming merger ([`merge_blocks`]) and the
/// native writer (`crates/wolfxl-writer/src/emit/sheet_xml.rs`). Adding a
/// slot for a future MS schema extension means updating one constant; both
/// consumers pick it up.
///
/// The table is exhaustive: every `<xsd:element>` child of CT_Worksheet
/// appears here. Unknown elements (extLst-internal extensions, x14ac compat
/// attributes, third-party tags) are NOT in this table and pass through
/// verbatim — they are not part of the §18.3.1.99 sequence.
pub mod ct_worksheet_order {
    /// `(local_name, ordinal)` pairs in ascending order. Lookup is a linear
    /// scan over 38 entries — slower than a `phf::Map` by a constant factor,
    /// but avoids the dependency and is fast enough for once-per-sheet use.
    pub const ECMA_ORDER: &[(&[u8], u32)] = &[
        (b"sheetPr", 1),
        (b"dimension", 2),
        (b"sheetViews", 3),
        (b"sheetFormatPr", 4),
        (b"cols", 5),
        (b"sheetData", 6),
        (b"sheetCalcPr", 7),
        (b"sheetProtection", 8),
        (b"protectedRanges", 9),
        (b"scenarios", 10),
        (b"autoFilter", 11),
        (b"sortState", 12),
        (b"dataConsolidate", 13),
        (b"customSheetViews", 14),
        (b"mergeCells", 15),
        (b"phoneticPr", 16),
        (b"conditionalFormatting", 17),
        (b"dataValidations", 18),
        (b"hyperlinks", 19),
        (b"printOptions", 20),
        (b"pageMargins", 21),
        (b"pageSetup", 22),
        (b"headerFooter", 23),
        (b"rowBreaks", 24),
        (b"colBreaks", 25),
        (b"customProperties", 26),
        (b"cellWatches", 27),
        (b"ignoredErrors", 28),
        (b"smartTags", 29),
        (b"drawing", 30),
        (b"legacyDrawing", 31),
        (b"legacyDrawingHF", 32),
        (b"picture", 33),
        (b"oleObjects", 34),
        (b"controls", 35),
        (b"webPublishItems", 36),
        (b"tableParts", 37),
        (b"extLst", 38),
    ];

    /// Look up the §18.3.1.99 ordinal for a local-name. `None` for unknown
    /// elements (extensions, third-party tags, etc.) — those flow through
    /// the merger verbatim at their source position.
    pub fn ordinal_of(local_name: &[u8]) -> Option<u32> {
        // Linear scan — 38 entries, called at most once per source element.
        // Compared first by length to short-circuit before byte-equality.
        ECMA_ORDER.iter().find_map(|(name, ord)| {
            if *name == local_name {
                Some(*ord)
            } else {
                None
            }
        })
    }
}

// ---------------------------------------------------------------------------
// merge_blocks — public API.
// ---------------------------------------------------------------------------

/// Merge a list of sibling blocks into a worksheet XML stream.
///
/// - If a block of the same root-element-name already exists in `sheet_xml`,
///   the existing block is replaced by the supplied bytes (semantics: "set
///   the entire block to this").
/// - If not present, the block is inserted at the position dictated by
///   ECMA §18.3.1.99 — strictly after every present sibling with a lower
///   ordinal and strictly before every present sibling with a higher ordinal.
/// - Unknown elements (extensions, x14ac, Microsoft-future, etc.) flow
///   through verbatim. The merger never re-serializes attributes or
///   reorders attribute lists on elements it does not own.
/// - `<conditionalFormatting>` is special: multiple
///   `SheetBlock::ConditionalFormatting` in `blocks` produce multiple
///   sibling elements in the output, and existing CF blocks in the source
///   are **all** removed before insertion. See RFC-011 §5.5.
///
/// Errors: `Err(...)` on malformed input XML (delegated from quick_xml).
/// Empty `blocks` returns `sheet_xml` unchanged with no allocation beyond
/// the `Vec::from(...)` for the return type.
pub fn merge_blocks(sheet_xml: &[u8], blocks: Vec<SheetBlock>) -> Result<Vec<u8>, String> {
    if blocks.is_empty() {
        // Empty-blocks fast path: byte-identical no-op. RFC-011 §5.6
        // idempotency contract.
        return Ok(sheet_xml.to_vec());
    }

    // -------- 1. Bucket the supplied blocks by ECMA position. --------
    //
    // Slot 17 (conditionalFormatting) is 0..N — multiple blocks per slot are
    // legal; every other slot is 0..1 by spec. Modeling the value as
    // `Vec<SheetBlock>` lets us preserve the caller's supplied order on
    // CF blocks without a second pass and tolerates a degenerate caller
    // who supplies more than one of a single-occurrence slot (we just emit
    // them in supplied order — garbage in, garbage out, RFC-011 §8 risk #6).
    let mut pending: BTreeMap<u32, Vec<SheetBlock>> = BTreeMap::new();
    let mut cf_replace = false;
    // The set of root local-names we are inserting — every existing source
    // block whose local-name appears here is dropped before its replacement
    // is written.
    let mut replace_names: Vec<&'static [u8]> = Vec::with_capacity(6);
    // Whether the caller supplied any block whose payload uses the `r:`
    // prefix (hyperlinks, tableParts, legacyDrawing). If yes, and the source
    // worksheet open tag does not declare xmlns:r, we inject the
    // declaration on the output's worksheet open tag (RFC-011 §8 risk #4).
    let mut needs_rel_ns = false;
    for block in blocks {
        let pos = block.ecma_position();
        let name = block.root_local_name();
        if !replace_names.contains(&name) {
            replace_names.push(name);
        }
        if matches!(block, SheetBlock::ConditionalFormatting(_)) {
            cf_replace = true;
        }
        if matches!(
            block,
            SheetBlock::Hyperlinks(_) | SheetBlock::TableParts(_) | SheetBlock::LegacyDrawing(_)
        ) {
            needs_rel_ns = true;
        }
        pending.entry(pos).or_default().push(block);
    }

    // -------- 2. Stream the source. --------
    let mut reader = XmlReader::from_reader(sheet_xml);
    // trim_text(false) — preserve all source whitespace (the modify-mode
    // minimal-diff promise from CLAUDE.md).
    reader.config_mut().trim_text(false);
    let mut writer = XmlWriter::new(Vec::with_capacity(sheet_xml.len() + 128));
    let mut buf: Vec<u8> = Vec::new();

    // Tracks whether we've already emitted the worksheet open tag (so we
    // know whether xmlns:r needs injection).
    let mut emitted_root = false;
    let mut depth: usize = 0;

    loop {
        let event = reader
            .read_event_into(&mut buf)
            .map_err(|e| format!("wolfxl-merger: XML parse error: {e}"))?;

        match event {
            Event::Start(e) => {
                let local = e.local_name().as_ref().as_bytes().to_vec();
                let current_depth = depth;

                // (a) Worksheet root: the only place we ever rewrite a
                // non-block element. We may need to inject xmlns:r.
                if current_depth == 0 && local == b"worksheet" && !emitted_root {
                    let to_emit =
                        ensure_rel_namespace(&e, needs_rel_ns).unwrap_or_else(|| e.borrow());
                    writer
                        .write_event(Event::Start(to_emit))
                        .map_err(|e| format!("wolfxl-merger: XML write error: {e}"))?;
                    emitted_root = true;
                    depth += 1;
                    buf.clear();
                    continue;
                }

                // (b) cf_replace: skip every <conditionalFormatting>
                // element wholesale; the caller's CF blocks land at the
                // ECMA-17 slot below.
                if current_depth == 1 && cf_replace && local == b"conditionalFormatting" {
                    consume_until_matching_end(&mut reader, &local)?;
                    buf.clear();
                    continue;
                }

                // (c) ECMA-known element: flush any pending blocks that
                // come strictly before this slot, then either emit
                // verbatim or skip-to-replace.
                if current_depth == 1 {
                    if let Some(ord) = ct_worksheet_order::ordinal_of(&local) {
                        flush_pending_before(&mut writer, &mut pending, ord)?;

                        if replace_names.contains(&local.as_slice()) {
                            // Drop the source block; its replacement will be
                            // emitted when we drain `pending` at slot `ord`
                            // (next call to `flush_pending_at` below).
                            consume_until_matching_end(&mut reader, &local)?;
                            flush_pending_at(&mut writer, &mut pending, ord)?;
                            buf.clear();
                            continue;
                        }

                        writer
                            .write_event(Event::Start(e.borrow()))
                            .map_err(|e| format!("wolfxl-merger: XML write error: {e}"))?;
                        depth += 1;
                        buf.clear();
                        continue;
                    }
                }

                // (d) Unknown element (extLst, x14ac:something, third-party
                // compat). Pass through verbatim — RFC-011 §3 / §5.3.
                writer
                    .write_event(Event::Start(e.borrow()))
                    .map_err(|e| format!("wolfxl-merger: XML write error: {e}"))?;
                depth += 1;
            }

            Event::Empty(e) => {
                let local = e.local_name().as_ref().as_bytes().to_vec();
                let current_depth = depth;

                // Self-closing <worksheet/> — RFC-011 §8 risk #3. Expand
                // to explicit <worksheet>...</worksheet> and flush every
                // pending block in between.
                if current_depth == 0 && local == b"worksheet" && !emitted_root {
                    let opened =
                        ensure_rel_namespace(&e, needs_rel_ns).unwrap_or_else(|| e.borrow());
                    writer
                        .write_event(Event::Start(opened))
                        .map_err(|e| format!("wolfxl-merger: XML write error: {e}"))?;
                    flush_all_pending(&mut writer, &mut pending)?;
                    writer
                        .write_event(Event::End(quick_xml::events::BytesEnd::new("worksheet")))
                        .map_err(|e| format!("wolfxl-merger: XML write error: {e}"))?;
                    emitted_root = true;
                    buf.clear();
                    continue;
                }

                if current_depth == 1 && cf_replace && local == b"conditionalFormatting" {
                    // Empty-element <conditionalFormatting/> — drop without
                    // recursing; emission of replacements happens at slot 17.
                    buf.clear();
                    continue;
                }

                if current_depth == 1 {
                    if let Some(ord) = ct_worksheet_order::ordinal_of(&local) {
                        flush_pending_before(&mut writer, &mut pending, ord)?;

                        if replace_names.contains(&local.as_slice()) {
                            // Empty source block — nothing to consume; the
                            // replacement still lands at the slot.
                            flush_pending_at(&mut writer, &mut pending, ord)?;
                            buf.clear();
                            continue;
                        }

                        writer
                            .write_event(Event::Empty(e.borrow()))
                            .map_err(|e| format!("wolfxl-merger: XML write error: {e}"))?;
                        buf.clear();
                        continue;
                    }
                }

                // Unknown empty element — verbatim.
                writer
                    .write_event(Event::Empty(e.borrow()))
                    .map_err(|e| format!("wolfxl-merger: XML write error: {e}"))?;
            }

            Event::End(e) => {
                if depth == 1 && e.local_name().as_ref() == "worksheet" {
                    flush_all_pending(&mut writer, &mut pending)?;
                }
                writer
                    .write_event(Event::End(e))
                    .map_err(|e| format!("wolfxl-merger: XML write error: {e}"))?;
                depth = depth.saturating_sub(1);
            }

            Event::Eof => break,

            // Decl, Text, CData, Comment, PI, DocType, GeneralRef — all
            // flow through at the source byte position. RFC-011 §8 risk #2:
            // a comment between two ECMA-ordered elements stays attached
            // to the preceding source element when we insert a block.
            other => {
                writer
                    .write_event(other)
                    .map_err(|e| format!("wolfxl-merger: XML write error: {e}"))?;
            }
        }
        buf.clear();
    }

    Ok(writer.into_inner())
}

/// Drain every pending block at slot `< ord` into the writer, in slot order.
fn flush_pending_before(
    writer: &mut XmlWriter<Vec<u8>>,
    pending: &mut BTreeMap<u32, Vec<SheetBlock>>,
    ord: u32,
) -> Result<(), String> {
    while let Some((&first_ord, _)) = pending.iter().next() {
        if first_ord >= ord {
            break;
        }
        if let Some(blocks) = pending.remove(&first_ord) {
            for b in blocks {
                writer.get_mut().extend_from_slice(b.bytes());
            }
        }
    }
    Ok(())
}

/// Drain every pending block at slot `== ord` into the writer, in supplied
/// order. Used when an existing source block at the same slot was just
/// dropped (replace path).
fn flush_pending_at(
    writer: &mut XmlWriter<Vec<u8>>,
    pending: &mut BTreeMap<u32, Vec<SheetBlock>>,
    ord: u32,
) -> Result<(), String> {
    if let Some(blocks) = pending.remove(&ord) {
        for b in blocks {
            writer.get_mut().extend_from_slice(b.bytes());
        }
    }
    Ok(())
}

/// Drain every remaining pending block into the writer, in slot order.
/// Called when the source `</worksheet>` is reached (or on a self-closing
/// `<worksheet/>`).
fn flush_all_pending(
    writer: &mut XmlWriter<Vec<u8>>,
    pending: &mut BTreeMap<u32, Vec<SheetBlock>>,
) -> Result<(), String> {
    while let Some((_, blocks)) = pending.pop_first() {
        for b in blocks {
            writer.get_mut().extend_from_slice(b.bytes());
        }
    }
    Ok(())
}

/// Consume reader events until the matching End event for `local` is
/// observed. Used to "skip" a source block we are replacing. Tracks nesting
/// depth in case the block contains nested elements with the same local
/// name (rare but legal — e.g. an `<extLst>` inside an `<extLst>` is not
/// possible, but conservative depth-tracking is cheap and correct).
fn consume_until_matching_end<R: std::io::BufRead>(
    reader: &mut XmlReader<R>,
    local: &[u8],
) -> Result<(), String> {
    let mut depth: i32 = 1;
    let mut buf: Vec<u8> = Vec::new();
    while depth > 0 {
        let event = reader
            .read_event_into(&mut buf)
            .map_err(|e| format!("wolfxl-merger: XML parse error during skip: {e}"))?;
        match event {
            Event::Start(e) if e.local_name().as_ref().as_bytes() == local => {
                depth += 1;
            }
            Event::End(e) if e.local_name().as_ref().as_bytes() == local => {
                depth -= 1;
            }
            Event::Eof => {
                return Err(format!(
                    "wolfxl-merger: unexpected EOF while skipping <{}>",
                    String::from_utf8_lossy(local)
                ));
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(())
}

/// If `needs_rel_ns` is true and the worksheet open tag does not already
/// declare `xmlns:r` (or any other prefix bound to the relationships
/// namespace), return a rewritten `BytesStart` with the declaration
/// appended. Otherwise return `None` (caller falls back to the borrowed
/// original — no allocation).
///
/// This is the *only* place the merger ever rewrites a non-block element
/// (RFC-011 §8 risk #4). The rewrite preserves the original attribute slice
/// byte-for-byte and appends a single new attribute at the tail.
fn ensure_rel_namespace(start: &BytesStart<'_>, needs_rel_ns: bool) -> Option<BytesStart<'static>> {
    if !needs_rel_ns {
        return None;
    }
    // Already declares the relationships namespace under some prefix — no-op.
    for attr in start.attributes().with_checks(false).flatten() {
        if attr.value.as_bytes() == REL_NS {
            return None;
        }
    }
    // Build an owned BytesStart with the same name, copy each existing
    // attribute through verbatim (its value stays in the source's escaped
    // form; `push_attribute` copies it into the owned buffer), then append
    // `xmlns:r="…"`.
    let mut new_start = BytesStart::new(start.name().as_ref().to_owned());
    for attr in start.attributes().with_checks(false).flatten() {
        new_start.push_attribute(attr);
    }
    new_start.push_attribute(("xmlns:r", std::str::from_utf8(REL_NS).unwrap()));
    Some(new_start)
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::ct_worksheet_order::{ordinal_of, ECMA_ORDER};
    use super::*;

    #[test]
    fn ecma_position_for_each_variant() {
        assert_eq!(SheetBlock::SheetViews(vec![]).ecma_position(), 3);
        assert_eq!(SheetBlock::SheetProtection(vec![]).ecma_position(), 8);
        assert_eq!(SheetBlock::AutoFilter(vec![]).ecma_position(), 11);
        assert_eq!(SheetBlock::MergeCells(vec![]).ecma_position(), 15);
        assert_eq!(
            SheetBlock::ConditionalFormatting(vec![]).ecma_position(),
            17
        );
        assert_eq!(SheetBlock::DataValidations(vec![]).ecma_position(), 18);
        assert_eq!(SheetBlock::Hyperlinks(vec![]).ecma_position(), 19);
        assert_eq!(SheetBlock::PrintOptions(vec![]).ecma_position(), 20);
        assert_eq!(SheetBlock::PageMargins(vec![]).ecma_position(), 21);
        assert_eq!(SheetBlock::PageSetup(vec![]).ecma_position(), 22);
        assert_eq!(SheetBlock::HeaderFooter(vec![]).ecma_position(), 23);
        assert_eq!(SheetBlock::LegacyDrawing(vec![]).ecma_position(), 31);
        assert_eq!(SheetBlock::TableParts(vec![]).ecma_position(), 37);
    }

    #[test]
    fn root_local_name_for_each_variant() {
        assert_eq!(
            SheetBlock::SheetViews(vec![]).root_local_name(),
            b"sheetViews"
        );
        assert_eq!(
            SheetBlock::SheetProtection(vec![]).root_local_name(),
            b"sheetProtection"
        );
        assert_eq!(
            SheetBlock::AutoFilter(vec![]).root_local_name(),
            b"autoFilter"
        );
        assert_eq!(
            SheetBlock::MergeCells(vec![]).root_local_name(),
            b"mergeCells"
        );
        assert_eq!(
            SheetBlock::ConditionalFormatting(vec![]).root_local_name(),
            b"conditionalFormatting"
        );
        assert_eq!(
            SheetBlock::DataValidations(vec![]).root_local_name(),
            b"dataValidations"
        );
        assert_eq!(
            SheetBlock::Hyperlinks(vec![]).root_local_name(),
            b"hyperlinks"
        );
        assert_eq!(
            SheetBlock::PrintOptions(vec![]).root_local_name(),
            b"printOptions"
        );
        assert_eq!(
            SheetBlock::PageMargins(vec![]).root_local_name(),
            b"pageMargins"
        );
        assert_eq!(
            SheetBlock::PageSetup(vec![]).root_local_name(),
            b"pageSetup"
        );
        assert_eq!(
            SheetBlock::HeaderFooter(vec![]).root_local_name(),
            b"headerFooter"
        );
        assert_eq!(
            SheetBlock::LegacyDrawing(vec![]).root_local_name(),
            b"legacyDrawing"
        );
        assert_eq!(
            SheetBlock::TableParts(vec![]).root_local_name(),
            b"tableParts"
        );
    }

    #[test]
    fn autofilter_replaces_existing_block() {
        // Source has <autoFilter ref="A1:B5"/> at slot 11.  Supplying a
        // new block must replace it in place.
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/><autoFilter ref="A1:B5"/><pageMargins/></worksheet>"#;
        let block = SheetBlock::AutoFilter(
            br#"<autoFilter ref="A1:D100"><filterColumn colId="0"/></autoFilter>"#.to_vec(),
        );
        let out = merge_blocks(xml, vec![block]).expect("merge");
        let s = std::str::from_utf8(&out).expect("utf8");
        assert!(s.contains(r#"<autoFilter ref="A1:D100">"#));
        assert!(!s.contains(r#"<autoFilter ref="A1:B5""#));
        // Exactly one autoFilter open in output.
        assert_eq!(s.matches("<autoFilter").count(), 1);
    }

    #[test]
    fn autofilter_inserts_before_mergecells() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/><mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells><pageMargins/></worksheet>"#;
        let block = SheetBlock::AutoFilter(br#"<autoFilter ref="A1:D10"/>"#.to_vec());
        let out = merge_blocks(xml, vec![block]).expect("merge");
        let s = std::str::from_utf8(&out).expect("utf8");
        let af = s.find("<autoFilter").unwrap();
        let mc = s.find("<mergeCells").unwrap();
        assert!(
            af < mc,
            "autoFilter (slot 11) must precede mergeCells (slot 15)"
        );
    }

    #[test]
    fn ecma_order_table_is_complete_38_slots() {
        // ECMA-376 §18.3.1.99 declares exactly 38 children of CT_Worksheet.
        // If this assertion ever fails, either the spec was extended or we
        // accidentally dropped/duplicated a slot — both are correctness bugs.
        assert_eq!(ECMA_ORDER.len(), 38);

        // Ordinals are strictly increasing 1..=38.
        for (i, (_, ord)) in ECMA_ORDER.iter().enumerate() {
            assert_eq!(*ord, (i as u32) + 1, "slot {} has wrong ordinal", i);
        }

        // No duplicate local-names.
        let mut names: Vec<&[u8]> = ECMA_ORDER.iter().map(|(n, _)| *n).collect();
        names.sort();
        let len = names.len();
        names.dedup();
        assert_eq!(names.len(), len, "ECMA_ORDER has duplicate local-names");
    }

    #[test]
    fn ecma_order_lookup_finds_each_known_local_name() {
        // For each SheetBlock variant, the lookup returns the
        // same ordinal as `ecma_position` — i.e. the block helpers and the
        // ordering table cannot drift independently.
        let variants = [
            SheetBlock::SheetViews(vec![]),
            SheetBlock::SheetProtection(vec![]),
            SheetBlock::MergeCells(vec![]),
            SheetBlock::ConditionalFormatting(vec![]),
            SheetBlock::DataValidations(vec![]),
            SheetBlock::Hyperlinks(vec![]),
            SheetBlock::PrintOptions(vec![]),
            SheetBlock::PageMargins(vec![]),
            SheetBlock::PageSetup(vec![]),
            SheetBlock::HeaderFooter(vec![]),
            SheetBlock::LegacyDrawing(vec![]),
            SheetBlock::TableParts(vec![]),
        ];
        for v in &variants {
            assert_eq!(
                ordinal_of(v.root_local_name()),
                Some(v.ecma_position()),
                "lookup mismatch for {:?}",
                v.root_local_name()
            );
        }
        // Unknown name returns None.
        assert_eq!(ordinal_of(b"x14ac:unknown"), None);
        assert_eq!(ordinal_of(b""), None);
    }

    #[test]
    fn merge_empty_blocks_is_noop() {
        // The empty-blocks fast path is the idempotency contract from
        // RFC-011 §5.6. Any input bytes — even malformed XML — pass through
        // because we never invoke the parser.
        let xml = b"<worksheet><sheetData/></worksheet>";
        let out = merge_blocks(xml, vec![]).expect("empty blocks must succeed");
        assert_eq!(out, xml);

        let garbage = b"not actually xml at all <<>>";
        let out = merge_blocks(garbage, vec![]).expect("empty blocks bypass parser");
        assert_eq!(out, garbage);
    }

    // -----------------------------------------------------------------------
    // RFC-011 §6 tests #2-#12. Test #1 (merge_empty_blocks_is_noop) is above.
    // -----------------------------------------------------------------------

    /// Helper: a minimal worksheet with the four "almost always present"
    /// children (dimension, sheetViews, sheetFormatPr, sheetData) plus
    /// pageMargins. Useful as a baseline for "where does block X land?"
    /// tests because pageMargins is slot 21, well past most insertion
    /// targets.
    fn minimal_with_pagemargins() -> &'static [u8] {
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>"#
    }

    fn pos_of(haystack: &[u8], needle: &[u8]) -> Option<usize> {
        haystack.windows(needle.len()).position(|w| w == needle)
    }

    #[test]
    fn insert_hyperlinks_into_minimal_sheet() {
        // Test #2: insert <hyperlinks> into a minimal sheet that has no
        // existing hyperlinks; must land at slot 19 — strictly before
        // <pageMargins> (slot 21) and strictly after <sheetData> (slot 6).
        let xml = minimal_with_pagemargins();
        let block = SheetBlock::Hyperlinks(
            br#"<hyperlinks><hyperlink ref="A1" r:id="rId1"/></hyperlinks>"#.to_vec(),
        );
        let out = merge_blocks(xml, vec![block]).expect("merge");
        let s = std::str::from_utf8(&out).expect("utf8");

        let hyper_pos = pos_of(&out, b"<hyperlinks>").expect("hyperlinks present");
        let pm_pos = pos_of(&out, b"<pageMargins").expect("pageMargins present");
        let sd_pos = pos_of(&out, b"<sheetData/>").expect("sheetData present");

        assert!(sd_pos < hyper_pos, "<sheetData> must precede <hyperlinks>");
        assert!(
            hyper_pos < pm_pos,
            "<hyperlinks> must precede <pageMargins>; got {s}"
        );
    }

    #[test]
    fn replace_existing_hyperlinks() {
        // Test #3: source already has a <hyperlinks> block; supplying a new
        // one drops the old.
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/><hyperlinks><hyperlink ref="OLD1" r:id="rIdOld"/></hyperlinks><pageMargins/></worksheet>"#;
        let block = SheetBlock::Hyperlinks(
            br#"<hyperlinks><hyperlink ref="NEW1" r:id="rIdNew"/></hyperlinks>"#.to_vec(),
        );
        let out = merge_blocks(xml, vec![block]).expect("merge");
        let s = std::str::from_utf8(&out).expect("utf8");
        assert!(s.contains("NEW1"), "new hyperlink ref present: {s}");
        assert!(
            !s.contains("OLD1"),
            "old hyperlink ref must be dropped: {s}"
        );
        assert!(!s.contains("rIdOld"), "old rId must be dropped: {s}");
        // Exactly one <hyperlinks> block in output.
        assert_eq!(s.matches("<hyperlinks>").count(), 1);
    }

    #[test]
    fn insert_into_correct_ecma_position() {
        // Test #4: for each of the 6 SheetBlock variants, build a sheet
        // with one earlier-slot element and one later-slot element, then
        // assert the inserted block lands strictly between them.
        struct Case {
            block: SheetBlock,
            earlier: &'static [u8],
            later: &'static [u8],
        }
        let cases = [
            Case {
                // mergeCells (15) lands between sheetData (6) and pageMargins (21).
                block: SheetBlock::MergeCells(
                    br#"<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>"#.to_vec(),
                ),
                earlier: b"<sheetData/>",
                later: b"<pageMargins",
            },
            Case {
                block: SheetBlock::ConditionalFormatting(
                    br#"<conditionalFormatting sqref="A1"><cfRule type="cellIs"/></conditionalFormatting>"#
                        .to_vec(),
                ),
                earlier: b"<sheetData/>",
                later: b"<pageMargins",
            },
            Case {
                block: SheetBlock::DataValidations(
                    br#"<dataValidations count="1"><dataValidation/></dataValidations>"#.to_vec(),
                ),
                earlier: b"<sheetData/>",
                later: b"<pageMargins",
            },
            Case {
                block: SheetBlock::Hyperlinks(
                    br#"<hyperlinks><hyperlink ref="A1" r:id="rId1"/></hyperlinks>"#.to_vec(),
                ),
                earlier: b"<sheetData/>",
                later: b"<pageMargins",
            },
            Case {
                block: SheetBlock::LegacyDrawing(br#"<legacyDrawing r:id="rId1"/>"#.to_vec()),
                earlier: b"<pageMargins",
                later: b"</worksheet>",
            },
            Case {
                block: SheetBlock::TableParts(
                    br#"<tableParts count="1"><tablePart r:id="rId1"/></tableParts>"#.to_vec(),
                ),
                earlier: b"<pageMargins",
                later: b"</worksheet>",
            },
        ];

        for case in cases {
            let block_root = case.block.root_local_name().to_vec();
            let out =
                merge_blocks(minimal_with_pagemargins(), vec![case.block.clone()]).expect("merge");
            let earlier_pos = pos_of(&out, case.earlier)
                .unwrap_or_else(|| panic!("earlier marker not found for {:?}", block_root));
            // The block we look for in output is the open tag of its root.
            let mut block_open = b"<".to_vec();
            block_open.extend_from_slice(&block_root);
            let block_pos = pos_of(&out, &block_open).unwrap_or_else(|| {
                panic!(
                    "inserted block <{}> not found in output: {}",
                    String::from_utf8_lossy(&block_root),
                    String::from_utf8_lossy(&out)
                )
            });
            let later_pos = pos_of(&out, case.later)
                .unwrap_or_else(|| panic!("later marker not found for {:?}", block_root));

            assert!(
                earlier_pos < block_pos,
                "earlier marker precedes block for <{}>",
                String::from_utf8_lossy(&block_root)
            );
            assert!(
                block_pos < later_pos,
                "block precedes later marker for <{}>",
                String::from_utf8_lossy(&block_root)
            );
        }
    }

    #[test]
    fn extlst_is_byte_preserved() {
        // Test #5 (RFC §8 risk #1, HEADLINE).
        //
        // Input has a top-level <extLst> with a `uri` attribute, an x14ac
        // namespace declaration, and an embedded extension element. After
        // merging in a Hyperlinks block, the extLst byte slice in the
        // output must be byte-identical to the slice in the input —
        // attribute order, prefix bindings, and entity escaping all
        // preserved.
        let extlst_bytes = br#"<extLst><ext uri="{0CCD9C8C-1C75-4C90-9DC1-3DA9F3D52A6F}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:sparklineGroup type="line" displayEmptyCellsAs="gap"><x14:colorSeries rgb="FF376092"/><x14:colorNegative rgb="FFFF0000"/><x14:colorAxis rgb="FF000000"/></x14:sparklineGroup></x14:sparklineGroups></ext></extLst>"#;
        let mut xml: Vec<u8> = b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheetData/><pageMargins/>".to_vec();
        xml.extend_from_slice(extlst_bytes);
        xml.extend_from_slice(b"</worksheet>");

        let block = SheetBlock::Hyperlinks(
            br#"<hyperlinks><hyperlink ref="A1" r:id="rId1"/></hyperlinks>"#.to_vec(),
        );
        let out = merge_blocks(&xml, vec![block]).expect("merge");

        assert!(
            pos_of(&out, extlst_bytes).is_some(),
            "<extLst> bytes must round-trip byte-identically; output was: {}",
            String::from_utf8_lossy(&out)
        );
    }

    #[test]
    fn unknown_element_passthrough() {
        // Test #6: an invented namespace-prefixed element between
        // <sheetData> and <pageMargins> survives the merge with attributes
        // intact, in the same relative position.
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wolfxl="urn:test"><sheetData/><wolfxl:custom value="42"/><pageMargins/></worksheet>"#;
        let block = SheetBlock::Hyperlinks(
            br#"<hyperlinks><hyperlink ref="A1" r:id="rId1"/></hyperlinks>"#.to_vec(),
        );
        let out = merge_blocks(xml, vec![block]).expect("merge");
        let s = std::str::from_utf8(&out).expect("utf8");

        // The unknown element appears verbatim with its attribute.
        assert!(
            s.contains(r#"<wolfxl:custom value="42"/>"#),
            "unknown element preserved with attribute: {s}"
        );
        // Hyperlinks landed at slot 19, between unknown element and
        // pageMargins. (Unknown element sticks to source position.)
        let custom = pos_of(&out, b"<wolfxl:custom").unwrap();
        let hyper = pos_of(&out, b"<hyperlinks>").unwrap();
        let pm = pos_of(&out, b"<pageMargins").unwrap();
        assert!(custom < hyper && hyper < pm);
    }

    #[test]
    fn multiple_conditionalformatting_blocks() {
        // Test #7: 3 supplied CF blocks all land contiguously at slot 17,
        // in supplied order.
        let xml = minimal_with_pagemargins();
        let blocks = vec![
            SheetBlock::ConditionalFormatting(
                br#"<conditionalFormatting sqref="A1:A10"><cfRule type="first"/></conditionalFormatting>"#.to_vec(),
            ),
            SheetBlock::ConditionalFormatting(
                br#"<conditionalFormatting sqref="B1:B10"><cfRule type="second"/></conditionalFormatting>"#.to_vec(),
            ),
            SheetBlock::ConditionalFormatting(
                br#"<conditionalFormatting sqref="C1:C10"><cfRule type="third"/></conditionalFormatting>"#.to_vec(),
            ),
        ];
        let out = merge_blocks(xml, blocks).expect("merge");
        let s = std::str::from_utf8(&out).expect("utf8");

        let first = s.find(r#"<cfRule type="first""#).unwrap();
        let second = s.find(r#"<cfRule type="second""#).unwrap();
        let third = s.find(r#"<cfRule type="third""#).unwrap();
        let pm = s.find("<pageMargins").unwrap();
        assert!(first < second && second < third);
        assert!(third < pm, "all CF before pageMargins");
        assert_eq!(s.matches("<conditionalFormatting").count(), 3);
    }

    #[test]
    fn conditionalformatting_replaces_all_existing() {
        // Test #8: source has 2 CF blocks; supply 1; output has only the
        // supplied one. RFC §5.5 replace-all semantics.
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/><conditionalFormatting sqref="A1"><cfRule type="OLD_A"/></conditionalFormatting><conditionalFormatting sqref="B1"><cfRule type="OLD_B"/></conditionalFormatting><pageMargins/></worksheet>"#;
        let block = SheetBlock::ConditionalFormatting(
            br#"<conditionalFormatting sqref="C1"><cfRule type="NEW_C"/></conditionalFormatting>"#
                .to_vec(),
        );
        let out = merge_blocks(xml, vec![block]).expect("merge");
        let s = std::str::from_utf8(&out).expect("utf8");

        assert!(s.contains("NEW_C"));
        assert!(!s.contains("OLD_A"));
        assert!(!s.contains("OLD_B"));
        assert_eq!(s.matches("<conditionalFormatting").count(), 1);
    }

    #[test]
    fn conditionalformatting_does_not_replace_nested_extension_elements() {
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><sheetData/><conditionalFormatting sqref="A1"><cfRule type="OLD_A"/></conditionalFormatting><extLst><ext><x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type="dataBar"/></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst></worksheet>"#;
        let block = SheetBlock::ConditionalFormatting(
            br#"<conditionalFormatting sqref="B1"><cfRule type="NEW_B"/></conditionalFormatting>"#
                .to_vec(),
        );
        let out = merge_blocks(xml, vec![block]).expect("merge");
        let s = std::str::from_utf8(&out).expect("utf8");

        assert!(s.contains("NEW_B"));
        assert!(!s.contains("OLD_A"));
        assert!(s.contains("x14:conditionalFormattings"));
        assert!(s.contains("x14:conditionalFormatting"));
        assert!(s.contains("x14:cfRule type=\"dataBar\""));
    }

    #[test]
    fn block_inserted_when_no_neighbors() {
        // Test #9: sheet has only <sheetData>. Insert TableParts(...).
        // Output: <sheetData/> then <tableParts> then </worksheet>.
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/></worksheet>"#;
        let block = SheetBlock::TableParts(
            br#"<tableParts count="1"><tablePart r:id="rId1"/></tableParts>"#.to_vec(),
        );
        let out = merge_blocks(xml, vec![block]).expect("merge");
        let s = std::str::from_utf8(&out).expect("utf8");

        let sd = s.find("<sheetData/>").unwrap();
        let tp = s.find("<tableParts").unwrap();
        let close = s.find("</worksheet>").unwrap();
        assert!(sd < tp && tp < close);
    }

    #[test]
    fn tableparts_after_extlst_is_wrong_and_we_fix_it() {
        // Test #10: pathological input has <extLst> before <tableParts>
        // (some third-party libs emit this). Source order is preserved on
        // pass-through (we don't reorder existing source elements), BUT
        // when we INSERT a fresh <tableParts>, it must land at slot 37,
        // i.e. before the existing <extLst> (slot 38).
        //
        // In other words: the merger doesn't rewrite source order on
        // pass-through, but it does place inserted blocks at the correct
        // slot — and slot 37 < 38 means the new tableParts must precede
        // the existing extLst in the output.
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/><extLst><ext uri="X"/></extLst></worksheet>"#;
        let block = SheetBlock::TableParts(
            br#"<tableParts count="1"><tablePart r:id="rId1"/></tableParts>"#.to_vec(),
        );
        let out = merge_blocks(xml, vec![block]).expect("merge");
        let s = std::str::from_utf8(&out).expect("utf8");

        let tp = s.find("<tableParts").unwrap();
        let ext = s.find("<extLst>").unwrap();
        assert!(
            tp < ext,
            "<tableParts> (slot 37) must precede <extLst> (slot 38) in output: {s}"
        );
    }

    #[test]
    fn large_sheet_streaming_memory_bounded() {
        // Test #11: a synthetic ~5 MB sheet (1k rows × 10 cells) merges in
        // bounded extra memory. Per RFC #11 the bound is < 4 MB peak on a
        // 50 MB sheet — we don't directly measure peak here (would require
        // a heap allocator probe), but the merger's contract is streaming:
        // O(input bytes) total work, no DOM build. This test guards
        // against a regression where someone refactors merge_blocks to
        // pre-buffer the whole input.
        //
        // Sanity property: output size is roughly input size + supplied
        // block size. If a future regression copies the input N times,
        // this test grows to N+1× the source and fails the assertion.
        let mut xml: Vec<u8> = b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheetData>".to_vec();
        for r in 1..=1000u32 {
            xml.extend_from_slice(format!("<row r=\"{r}\">").as_bytes());
            for c in 0..10u32 {
                let col_letter = (b'A' + c as u8) as char;
                xml.extend_from_slice(
                    format!("<c r=\"{col_letter}{r}\" t=\"n\"><v>{r}</v></c>").as_bytes(),
                );
            }
            xml.extend_from_slice(b"</row>");
        }
        xml.extend_from_slice(b"</sheetData></worksheet>");

        let input_size = xml.len();
        let block = SheetBlock::Hyperlinks(
            br#"<hyperlinks><hyperlink ref="A1" r:id="rId1"/></hyperlinks>"#.to_vec(),
        );
        let out = merge_blocks(&xml, vec![block]).expect("merge");
        let output_size = out.len();

        // Output should be input + ~60 bytes for the hyperlinks block,
        // not 2× input (which would indicate pre-buffering).
        assert!(
            output_size < input_size + 4096,
            "output {} far exceeds input {} + small block; possible buffering regression",
            output_size,
            input_size
        );
        assert!(
            output_size > input_size,
            "output must contain the inserted block"
        );
    }

    #[test]
    fn byte_identical_when_block_already_present_and_unchanged() {
        // Test #12: if the supplied bytes for a block exactly match what's
        // in the source, the output is byte-identical to the input.
        // Lets future RFC-022 etc. cheaply detect no-op patches.
        let block_bytes: &[u8] = br#"<hyperlinks><hyperlink ref="A1" r:id="rId1"/></hyperlinks>"#;
        let mut xml: Vec<u8> = b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheetData/>".to_vec();
        xml.extend_from_slice(block_bytes);
        xml.extend_from_slice(b"<pageMargins/></worksheet>");

        let out =
            merge_blocks(&xml, vec![SheetBlock::Hyperlinks(block_bytes.to_vec())]).expect("merge");
        // The output's <hyperlinks> block bytes equal the supplied bytes.
        assert!(
            pos_of(&out, block_bytes).is_some(),
            "supplied block bytes appear verbatim in output"
        );
        // Output contains exactly one <hyperlinks> block.
        assert_eq!(
            out.windows(b"<hyperlinks>".len())
                .filter(|w| *w == b"<hyperlinks>")
                .count(),
            1
        );
    }

    // -----------------------------------------------------------------------
    // Risk-fallback tests for §8 issues #2-#4 (comments, self-closing root,
    // namespace injection). Belong to commit 2 alongside the §6 tests.
    // -----------------------------------------------------------------------

    #[test]
    fn comments_pass_through_at_source_position() {
        // RFC §8 risk #2 — XML comments at worksheet level stay attached
        // to their preceding source element when a block is inserted.
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/><!-- created by foo --><pageMargins/></worksheet>"#;
        let block = SheetBlock::Hyperlinks(
            br#"<hyperlinks><hyperlink ref="A1" r:id="rId1"/></hyperlinks>"#.to_vec(),
        );
        let out = merge_blocks(xml, vec![block]).expect("merge");
        let s = std::str::from_utf8(&out).expect("utf8");
        assert!(s.contains("<!-- created by foo -->"));
    }

    #[test]
    fn self_closing_root_expands_and_flushes() {
        // RFC §8 risk #3. Degenerate <worksheet/> in the source — the
        // merger must expand to <worksheet>...</worksheet> with the block
        // inside.
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>"#;
        let block = SheetBlock::Hyperlinks(
            br#"<hyperlinks><hyperlink ref="A1" r:id="rId1"/></hyperlinks>"#.to_vec(),
        );
        let out = merge_blocks(xml, vec![block]).expect("merge");
        let s = std::str::from_utf8(&out).expect("utf8");

        assert!(s.contains("<hyperlinks>"));
        assert!(s.contains("</worksheet>"));
        assert!(!s.contains("<worksheet/>") && !s.contains("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>"),
                "self-closing form must be expanded");
    }

    #[test]
    fn rel_namespace_injected_when_missing() {
        // RFC §8 risk #4. Source <worksheet> declares the default ns but
        // not xmlns:r. When we insert a block whose payload uses r:id,
        // the merger appends xmlns:r on the output's worksheet open tag.
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>"#;
        let block = SheetBlock::Hyperlinks(
            br#"<hyperlinks><hyperlink ref="A1" r:id="rId1"/></hyperlinks>"#.to_vec(),
        );
        let out = merge_blocks(xml, vec![block]).expect("merge");
        let s = std::str::from_utf8(&out).expect("utf8");
        assert!(
            s.contains(
                r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#
            ),
            "xmlns:r must be injected when missing: {s}"
        );
    }

    #[test]
    fn rel_namespace_not_duplicated_when_already_present() {
        // The detection must accept any existing binding of the rels URI
        // to a prefix — not blindly match on `r:` — and not double-inject.
        let xml = minimal_with_pagemargins();
        let block = SheetBlock::Hyperlinks(
            br#"<hyperlinks><hyperlink ref="A1" r:id="rId1"/></hyperlinks>"#.to_vec(),
        );
        let out = merge_blocks(xml, vec![block]).expect("merge");
        let s = std::str::from_utf8(&out).expect("utf8");
        assert_eq!(
            s.matches("xmlns:r=").count(),
            1,
            "xmlns:r must not be duplicated when already present"
        );
    }
}
