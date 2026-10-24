//! `wolfxl agent <file> --max-tokens N` — token-budgeted workbook briefing.
//!
//! The brief is composed greedily for an LLM context window. Priority order:
//! 1. Workbook overview (every sheet, dims + class + first-column header)
//! 2. Focused sheet header (chosen sheet, dims, class)
//! 3. Column headers for the focused sheet
//! 4. Named ranges (capped, budgeted — drops first if budget is tight)
//! 5. Head rows (first 3) — ordering tends to be authored, so head matters
//! 6. Tail rows (last 2) — totals/EPS/footnotes typically live at the bottom
//! 7. Stratified middle samples — fills remaining budget with diverse rows
//! 8. Footer line: `# wolfxl agent: USED/LIMIT tokens (cl100k_base)`
//!
//! Block 1+2+3 is the orientation core and is emitted even if the row blocks
//! get dropped. If even the orientation core overflows, we still emit it
//! (truncating the budget would defeat the purpose of telling the agent what
//! sheets exist) and the footer reports the overage so the caller knows.
//!
//! Footer tokens count toward the budget. We reserve a worst-case footer
//! cost up-front (so body composition stops early enough), and the final
//! printed `USED` value re-encodes `body + footer` so the reported number
//! matches what `tiktoken.cl100k_base.encode(stdout)` returns.
//!
//! Token counts use `tiktoken-rs::cl100k_base` to match the tokenizer used by
//! `spreadsheet-peek/benchmarks/measure_tokens.py`.

use std::io::{self, Write};
use std::path::PathBuf;

use anyhow::{Context, Result};
use tiktoken_rs::{cl100k_base_singleton, CoreBPE};
use wolfxl_core::{Cell, CellValue, Sheet, SheetClass, Workbook, WorkbookMap};

pub fn run(file: PathBuf, max_tokens: usize, target_sheet: Option<String>) -> Result<()> {
    let mut wb =
        Workbook::open(&file).with_context(|| format!("opening workbook: {}", file.display()))?;
    let map = wb.map().context("building workbook map")?;
    let target = pick_target(&map, target_sheet.as_deref())?;
    let sheet = wb
        .sheet(&target)
        .with_context(|| format!("loading sheet {target:?}"))?;

    let bpe = cl100k_base_singleton();
    let budget = Budget::new(bpe, max_tokens);
    let mut buf = String::new();

    // Orientation core (overview + sheet header + columns) always lands,
    // even if it overflows — we'd rather report overage in the footer than
    // hide the workbook structure from the agent.
    write_overview(&mut buf, &map);
    write_sheet_header(&mut buf, &target, &sheet, &map);

    // Named ranges are best-effort: emitted via try_append so they drop
    // first when the budget is tight. They live between the orientation
    // core and the row blocks so they get priority over samples but never
    // crowd out columns.
    write_named_ranges(&mut buf, &budget, &map);

    write_rows(&mut buf, &sheet, &budget);

    // The footer reports `body+footer` tokens. Because the printed `used`
    // value is itself part of the footer, we iterate to a fixed point: in
    // practice this converges in 1-2 passes since cl100k_base token counts
    // are stable across small changes in the digit length of the numerator.
    let mut reported = budget.used(&buf);
    for _ in 0..3 {
        let probe = format_footer(reported, budget.limit);
        let total = budget.used_with(&buf, &probe);
        if total == reported {
            break;
        }
        reported = total;
    }
    buf.push_str(&format_footer(reported, budget.limit));

    io::stdout().lock().write_all(buf.as_bytes())?;
    Ok(())
}

fn format_footer(used: usize, limit: usize) -> String {
    format!("\n# wolfxl agent: {used}/{limit} tokens (cl100k_base)\n")
}

/// Pick the focus sheet. Explicit `--sheet` wins; otherwise prefer the
/// largest `Data` sheet by row count, falling back to the first sheet so
/// workbooks of all-summary or all-empty shape still produce output.
fn pick_target(map: &WorkbookMap, requested: Option<&str>) -> Result<String> {
    if let Some(name) = requested {
        if !map.sheets.iter().any(|s| s.name == name) {
            let names: Vec<&str> = map.sheets.iter().map(|s| s.name.as_str()).collect();
            anyhow::bail!("sheet {name:?} not found; available: {}", names.join(", "));
        }
        return Ok(name.to_string());
    }
    let best_data = map
        .sheets
        .iter()
        .filter(|s| s.class == SheetClass::Data)
        .max_by_key(|s| s.rows);
    if let Some(s) = best_data {
        return Ok(s.name.clone());
    }
    map.sheets
        .first()
        .map(|s| s.name.clone())
        .ok_or_else(|| anyhow::anyhow!("workbook has no sheets"))
}

fn write_overview(buf: &mut String, map: &WorkbookMap) {
    let path = std::path::Path::new(&map.path)
        .file_name()
        .and_then(|s| s.to_str())
        .unwrap_or(&map.path);
    buf.push_str(&format!("WORKBOOK: {path}\n"));
    buf.push_str("SHEETS:\n");
    for s in &map.sheets {
        let first_col = s.headers.first().map(String::as_str).unwrap_or("");
        if first_col.is_empty() {
            buf.push_str(&format!(
                "  - {name} ({rows}r×{cols}c) [{class}]\n",
                name = s.name,
                rows = s.rows,
                cols = s.cols,
                class = s.class.as_str()
            ));
        } else {
            buf.push_str(&format!(
                "  - {name} ({rows}r×{cols}c) [{class}] col1={first_col:?}\n",
                name = s.name,
                rows = s.rows,
                cols = s.cols,
                class = s.class.as_str()
            ));
        }
    }
}

/// Best-effort named-ranges section. Capped at 8 entries so a workbook
/// with 200 named ranges can't single-handedly drain the budget; longer
/// lists get a `… (+N more)` overflow marker. The whole section is also
/// gated through `try_append`, so under tight budgets it drops cleanly
/// and never partially emits.
fn write_named_ranges(buf: &mut String, budget: &Budget, map: &WorkbookMap) {
    if map.named_ranges.is_empty() {
        return;
    }
    const MAX: usize = 8;
    let total = map.named_ranges.len();
    let mut section = String::from("\nNAMED_RANGES:\n");
    for (name, formula) in map.named_ranges.iter().take(MAX) {
        section.push_str(&format!("  - {name} = {formula}\n"));
    }
    if total > MAX {
        section.push_str(&format!("  … (+{} more)\n", total - MAX));
    }
    try_append(buf, budget, &section);
}

fn write_sheet_header(buf: &mut String, target: &str, sheet: &Sheet, map: &WorkbookMap) {
    let (rows, cols) = sheet.dimensions();
    let class = map
        .sheets
        .iter()
        .find(|s| s.name == target)
        .map(|s| s.class.as_str())
        .unwrap_or("data");
    buf.push('\n');
    buf.push_str(&format!(
        "SHEET: {target}  [{class}]  {rows} rows × {cols} cols\n"
    ));
    // Headers may come back as a vector of empty strings when the first
    // row of a sheet has no header values (e.g., a sparse summary). In
    // that case a `HEADERS:\t\t\t\t` line burns tokens for no signal —
    // skip it. We also trim trailing empties so wide tables with mostly
    // unlabelled trailing columns don't pad the line with noise tabs.
    let headers = sheet.headers();
    if let Some(end) = headers.iter().rposition(|h| !h.is_empty()) {
        let trimmed = &headers[..=end];
        buf.push_str("HEADERS: ");
        buf.push_str(&trimmed.join("\t"));
        buf.push('\n');
    }
}

/// Emit head + tail + middle stratified samples within the remaining budget.
/// The body is `sheet.rows()[1..]` (header is row 0).
fn write_rows(buf: &mut String, sheet: &Sheet, budget: &Budget) {
    let body_start = 1usize;
    let total = sheet.rows().len();
    if total <= body_start {
        return;
    }
    let body_count = total - body_start;
    let head_n = 3.min(body_count);
    let tail_n = if body_count > head_n {
        2.min(body_count - head_n)
    } else {
        0
    };

    // Head: emit all-or-nothing as a labelled block. If the section overflows
    // we'd rather skip than half-emit (truncated rows lie about row count).
    let head_section = render_section(
        sheet,
        body_start,
        body_start + head_n,
        &format!("ROWS (head {head_n} of {body_count}):"),
    );
    try_append(buf, budget, &head_section);

    if tail_n > 0 {
        let start = total - tail_n;
        let tail_section = render_section(
            sheet,
            start,
            total,
            &format!("ROWS (tail {tail_n} of {body_count}):"),
        );
        try_append(buf, budget, &tail_section);
    }

    let middle_lo = body_start + head_n;
    let middle_hi = total.saturating_sub(tail_n);
    if middle_hi > middle_lo {
        let middle_count = middle_hi - middle_lo;
        emit_stratified_middle(buf, sheet, budget, middle_lo, middle_count);
    }
}

/// Stratification = uniform stride. For a 10-row middle and 4 picks we'd
/// take indices [0, 3, 6, 9]. Caller supplies `lo` (absolute row index) and
/// `middle_count` (the size of the middle window).
///
/// We try up to 8 picks (cap on visual noise + budget reasonableness) and
/// emit them one at a time. The header line lands first; if even the
/// header overflows the row block is skipped entirely.
fn emit_stratified_middle(
    buf: &mut String,
    sheet: &Sheet,
    budget: &Budget,
    lo: usize,
    middle_count: usize,
) {
    let target_picks = 8usize.min(middle_count);
    if target_picks == 0 {
        return;
    }

    let mut picks: Vec<usize> = Vec::with_capacity(target_picks);
    if target_picks == 1 {
        picks.push(lo + middle_count / 2);
    } else {
        for i in 0..target_picks {
            let off = (i as f64) * ((middle_count - 1) as f64) / ((target_picks - 1) as f64);
            let row_idx = lo + (off.round() as usize);
            if !picks.contains(&row_idx) {
                picks.push(row_idx);
            }
        }
    }

    let header = format!(
        "ROWS (middle stratified, up to {} of {}):\n",
        picks.len(),
        middle_count
    );
    if !try_append(buf, budget, &header) {
        return;
    }
    for idx in picks {
        let line = format!("  {}\n", row_as_tsv(sheet, &sheet.rows()[idx]));
        if !try_append(buf, budget, &line) {
            break;
        }
    }
}

fn render_section(sheet: &Sheet, lo: usize, hi: usize, label: &str) -> String {
    let mut out = String::new();
    out.push_str(label);
    out.push('\n');
    for row in &sheet.rows()[lo..hi] {
        out.push_str("  ");
        out.push_str(&row_as_tsv(sheet, row));
        out.push('\n');
    }
    out
}

/// Append `section` to `buf` only if doing so keeps the running token count
/// under budget. Returns whether the section landed.
///
/// We probe by counting `buf + section` rather than `used(buf) + count(section)`
/// because cl100k_base BPE merges across boundaries — a token boundary that
/// sits at the seam can collapse two pieces into one. Sub-additive merge
/// would let an additive check reject sections that would actually fit.
///
/// The check is against `body_limit`, not `limit`: `body_limit` reserves
/// space for the (yet-to-be-appended) footer line so the total emission
/// honors `--max-tokens N`.
fn try_append(buf: &mut String, budget: &Budget, section: &str) -> bool {
    let candidate_used = budget.used_with(buf, section);
    if candidate_used > budget.body_limit {
        return false;
    }
    buf.push_str(section);
    true
}

fn row_as_tsv(sheet: &Sheet, row: &[Cell]) -> String {
    row.iter()
        .map(|cell| {
            if let CellValue::Error(error) = &cell.value {
                format!("ERR:{}", sheet.error_text(error))
            } else {
                display_cell(cell)
            }
        })
        .collect::<Vec<_>>()
        .join("\t")
}

/// Compact text rendering for the agent brief. We deliberately do NOT
/// thousand-group integers here: every comma is a token boundary in
/// cl100k_base, so `1234567` is two tokens but `1,234,567` is five. The
/// agent doesn't need pretty output; it needs cheap output.
fn display_cell(cell: &Cell) -> String {
    match &cell.value {
        CellValue::Empty => String::new(),
        CellValue::String(s) => s.clone(),
        CellValue::Bool(b) => if *b { "true" } else { "false" }.to_string(),
        CellValue::Int(n) => n.to_string(),
        CellValue::Float(n) => {
            if n.is_finite() && n.fract() == 0.0 && n.abs() < 1e15 {
                format!("{n:.0}")
            } else if n.is_finite() {
                format!("{n:.2}")
            } else {
                n.to_string()
            }
        }
        CellValue::Date(d) => d.format("%Y-%m-%d").to_string(),
        CellValue::DateTime(dt) => dt.format("%Y-%m-%dT%H:%M:%S").to_string(),
        CellValue::Time(t) => t.format("%H:%M:%S").to_string(),
        CellValue::Error(e) => format!("ERR:{e}"),
    }
}

/// Stateless token-budget projection. Wraps `cl100k_base` and exposes
/// two queries: `used(buf)` re-encodes the buffer end-to-end, and
/// `used_with(buf, section)` re-encodes `buf + section` (used by
/// `try_append` to predict the cost of a candidate section).
///
/// `limit` is the user-supplied `--max-tokens N`. `body_limit` is
/// `limit - footer_reserve`, where `footer_reserve` is the worst-case
/// token cost of the trailing `# wolfxl agent: USED/LIMIT ...` line.
/// `try_append` checks against `body_limit` so the final emission
/// (body + footer) honors the user's limit.
///
/// We re-encode the whole accumulated buffer rather than incrementing
/// per-add because BPE is *not* additive — `tokens(a + b) <= tokens(a)
/// + tokens(b)` since adjacent pieces can merge into a single token
/// across the boundary. Recomputing on the final buffer is the only
/// way to get a count that matches what `tiktoken.cl100k_base.encode(
/// full_output)` produces in Python.
struct Budget<'a> {
    bpe: &'a CoreBPE,
    limit: usize,
    body_limit: usize,
}

impl<'a> Budget<'a> {
    fn new(bpe: &'a CoreBPE, limit: usize) -> Self {
        // Worst-case footer width: digits scaled to 10x the limit so an
        // overage report like `8500/800` still fits the reserve. The
        // reserve is intentionally a couple tokens looser than strictly
        // necessary; over-reserving costs at most a row of samples.
        let max_used = limit.saturating_mul(10).max(limit);
        let worst = format_footer(max_used, limit);
        let footer_reserve = bpe.encode_ordinary(&worst).len();
        let body_limit = limit.saturating_sub(footer_reserve);
        Self {
            bpe,
            limit,
            body_limit,
        }
    }

    /// Total tokens an accumulated buffer would cost (matches what
    /// `tiktoken.cl100k_base.encode(buf)` returns in Python).
    fn used(&self, buf: &str) -> usize {
        self.bpe.encode_ordinary(buf).len()
    }

    /// Tokens of `buf + section`, encoded as one piece so cross-boundary
    /// BPE merges are counted correctly. Cheaper allocations would split
    /// the encode but lose merge accuracy at the seam.
    fn used_with(&self, buf: &str, section: &str) -> usize {
        // Re-encoding the full concatenation is the only way to get the
        // exact final count; BPE is non-additive at piece boundaries.
        let mut probe = String::with_capacity(buf.len() + section.len());
        probe.push_str(buf);
        probe.push_str(section);
        self.bpe.encode_ordinary(&probe).len()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use wolfxl_core::SheetMap;

    #[test]
    fn budget_recomputes_against_full_buffer() {
        // BPE is non-additive across boundaries; `used_with` re-encodes
        // the full concatenation so cross-seam merges are counted. An
        // additive `count(a) + count(b)` would over-count and reject
        // sections that actually fit.
        let bpe = cl100k_base_singleton();
        let b = Budget::new(bpe, 100);
        let combined = b.used("hello world");
        let with = b.used_with("hello", " world");
        assert_eq!(combined, with);
    }

    #[test]
    fn budget_reserves_room_for_footer() {
        // body_limit must leave at least enough headroom that appending
        // the actual footer never crosses `limit`. We use a worst-case
        // reserve of 10x the limit's digit width as the numerator, so
        // even when the orientation core overflows and the footer reads
        // e.g. "8500/800", body_limit + footer fits the projection.
        let bpe = cl100k_base_singleton();
        let b = Budget::new(bpe, 800);
        assert!(b.body_limit < b.limit, "must reserve nonzero footer space");
        let realistic_footer = format_footer(b.limit, b.limit);
        let footer_tokens = b.used(&realistic_footer);
        assert!(
            b.body_limit + footer_tokens <= b.limit,
            "body_limit ({}) + realistic footer ({}) must fit limit ({})",
            b.body_limit,
            footer_tokens,
            b.limit
        );
    }

    #[test]
    fn pick_target_prefers_largest_data_sheet() {
        // Build a synthetic WorkbookMap with three sheets of different
        // sizes and classes; the largest-by-rows Data sheet wins even if
        // a Summary sheet has more rows.
        let map = WorkbookMap {
            path: "test.xlsx".to_string(),
            sheets: vec![
                SheetMap {
                    name: "Notes".to_string(),
                    rows: 50,
                    cols: 1,
                    class: SheetClass::Readme,
                    headers: vec!["note".to_string()],
                    tables: vec![],
                },
                SheetMap {
                    name: "P&L".to_string(),
                    rows: 21,
                    cols: 7,
                    class: SheetClass::Data,
                    headers: vec![],
                    tables: vec![],
                },
                SheetMap {
                    name: "Detail".to_string(),
                    rows: 200,
                    cols: 12,
                    class: SheetClass::Data,
                    headers: vec![],
                    tables: vec![],
                },
            ],
            named_ranges: vec![],
        };
        assert_eq!(pick_target(&map, None).unwrap(), "Detail");
    }

    #[test]
    fn pick_target_falls_back_to_first_sheet_when_no_data() {
        let map = WorkbookMap {
            path: "test.xlsx".to_string(),
            sheets: vec![
                SheetMap {
                    name: "Cover".to_string(),
                    rows: 5,
                    cols: 1,
                    class: SheetClass::Readme,
                    headers: vec![],
                    tables: vec![],
                },
                SheetMap {
                    name: "Summary".to_string(),
                    rows: 10,
                    cols: 4,
                    class: SheetClass::Summary,
                    headers: vec![],
                    tables: vec![],
                },
            ],
            named_ranges: vec![],
        };
        assert_eq!(pick_target(&map, None).unwrap(), "Cover");
    }

    #[test]
    fn pick_target_explicit_wins_over_heuristic() {
        let map = WorkbookMap {
            path: "test.xlsx".to_string(),
            sheets: vec![SheetMap {
                name: "P&L".to_string(),
                rows: 21,
                cols: 7,
                class: SheetClass::Data,
                headers: vec![],
                tables: vec![],
            }],
            named_ranges: vec![],
        };
        assert_eq!(pick_target(&map, Some("P&L")).unwrap(), "P&L");
        assert!(pick_target(&map, Some("Nope")).is_err());
    }
}
