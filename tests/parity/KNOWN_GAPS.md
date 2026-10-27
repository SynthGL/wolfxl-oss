# Known parity gaps — WolfXL vs openpyxl

This file is the historical parity-gap ledger. The active openpyxl-supported
gap inventory is currently empty; new gaps must be added here only after they
are also encoded in `openpyxl_surface.py` or `docs/migration/_compat_spec.py`.

Gaps are also encoded in `openpyxl_surface.py` via `wolfxl_supported=False`
— the parity smoke test keeps the two in sync.

## Roadmap status (overview)

- **Phase 1 — T1 DefinedName WRITE** — closed in 1.3 (Sprint Ι Pod-δ D4).
- **Phase 2 — T0 Password-protected reads** — closed in 1.3 (Sprint Ι Pod-γ).
- **Phase 3 — T2 Rich-text reads + writes** — closed in 1.3 (Sprint Ι Pod-α).
- **Phase 4 — T2 Streaming reads** — closed in 1.3 (Sprint Ι Pod-β).
- **Phase 5 — T1 `.xls` / `.xlsb`** — closed in 1.4 (Sprint Κ, RFC-043).
- **1.5 — Encryption writes + image construction + streaming-datetime fix** — closed in 1.5 (Sprint Λ Pod-α/β/γ; RFC-044, RFC-045, plus Pod-γ's streaming-datetime correctness fix). Lifts "writing encrypted xlsx" and "image construction" from out-of-scope.
- **1.6 — Chart construction (8 types, full depth) + RFC-035 chart-deep-clone + modify-mode `add_chart`** — closed in 1.6 (Sprint Μ Pod-α/β/γ/δ; RFC-046). Lifts "chart construction" from out-of-scope for the 2D type families (Bar / Line / Pie / Doughnut / Area / Scatter / Bubble / Radar). 3D / Stock / Surface / ProjectedPie deferred to v1.6.1; pivot-chart linkage deferred to v2.0.0 (Sprint Ν).
- **1.6.1 — Chart-dict contract reconciliation + 3D / Stock / Surface / ProjectedPie families + modify-mode high-level `add_chart`** — closed in 1.6.1 (Sprint Μ-prime Pod-α′/β′/γ′/δ′; RFC-046 §10/§11). Lifts the v1.6.0 chart-dict gap (37 xfailed advanced sub-feature tests in `tests/test_charts_write.py` flip to pass) and ships the eight deferred 3D / Stock / Surface / ProjectedPie chart classes as real implementations. Also replaces the warn-and-drop fallback in `Workbook._flush_pending_charts_to_patcher` with a real dict→bytes bridge via the new `serialize_chart_dict` PyO3 export.
- **1.7 — Public-launch slice (no pivot tables)** — closed in 1.7 (Sprint Ξ Pod-α/β/γ/δ; RFC-050/051/052/053). Lifts v1.6.1's lone `xfail` (`chart.title = RichText(...)`) and ships `Worksheet.remove_chart` / `replace_chart`. Bumps `pyproject.toml` and `Cargo.toml` to `1.7.0` (was drifted at `0.5.0`); promotes PyPI classifier to `Production/Stable`. Refreshes `docs/migration/` and `docs/performance/` with v1.7 status. Materialises `Plans/launch-posts.md`. Documents RFC-046 §13 legacy chart-dict key sunset (deprecated in 1.7; removed in 2.0).
- **2.0 — Pivot tables + pivot caches + pivot charts** — closed in 2.0 (Sprint Ν Pods α/β/γ/δ/ε; RFC-047 / RFC-048 / RFC-049 / RFC-054). Lifts the last construction-side row from "Out of scope" — pivot tables, pivot caches, and pivot-chart linkage are all constructible from Python. WolfXL emits `pivotCacheRecords{N}.xml` from scratch (Option A); pivots open in Excel / LibreOffice / openpyxl with data populated, no refresh-on-open required (true in Excel since 2.0.20, #63: earlier 2.0.x pivots triggered Excel's repair prompt and left the pivot range blank until refresh). Lifts the RFC-035 §10 limit on `copy_worksheet` of pivot-bearing sheets. Removes the RFC-046 §13 legacy chart-dict key shim (deprecated 1.7, removed 2.0). README rewrite now uses openpyxl-compatible positioning while v2.0 benchmark claims remain gated on the release audit.

The openpyxl-parity roadmap is now exhausted at the read level
(1.0–1.4) AND at the full construction level (1.5 encryption +
images, 1.6–1.6.1 charts, 1.7 chart-stack debt + RichText titles,
**2.0 pivot tables**, and the post-PR #23 audit slice for slicers,
pivot calculated fields/items, GroupItems, and the remaining chart
display-unit / data-point overrides). The "Out of scope" section
below now lists only surfaces openpyxl itself does not expose or project
caveats that are not openpyxl advantages.

## Gate

- Every gap must have a phase owner.
- Closing a gap: flip `wolfxl_supported=True`, remove the entry here, expect
  `test_known_gap_still_gaps` to fail red (which is the signal to also
  commit the ratchet-baseline update).

## Gaps by category

### Sheet access (type-hint imports — SHIPPED)

`wolfxl.Worksheet` and `wolfxl.Cell` are now re-exported at the top level
(see `python/wolfxl/__init__.py`). SynthGL's type-hint imports work as a
drop-in.

### Range / layout API shape (Phase 0 cleanup — SHIPPED)

`Worksheet.max_row`, `Worksheet.max_column`, and `Worksheet.merged_cells`
are now public properties. `merged_cells` returns a `_MergedCellsProxy`
backed by the Rust `read_merged_ranges` call in read mode (closes the
"merged_cells empty on read" per-fixture gap below as a side-effect).

### Utils (Phase 0 cleanup — SHIPPED)

All seven utility symbols ship through `python/wolfxl/utils/`:

- `wolfxl.utils.cell.get_column_letter`, `column_index_from_string`,
  `range_boundaries`, `coordinate_to_tuple`
- `wolfxl.utils.numbers.is_date_format`
- `wolfxl.utils.datetime.from_excel`, `CALENDAR_WINDOWS_1900`

Behavior is bug-for-bug compatible with openpyxl 3.1.x and pinned by
`test_utils_parity.py`. Bound checks (`get_column_letter` capped at 18278
= ZZZ) and the 1900 leap-year correction (`from_excel`) match openpyxl
verbatim.

### Phase 1 — T1 DefinedName WRITE

| openpyxl path | phase | note |
|---|---|---|
| `Workbook.defined_names["X"] = DefinedName(...)` | Phase 1 | Rust side (`add_named_range`) already exists; just expose `__setitem__` in the Python proxy. |

### Phase 2 — T0 Password-protected reads (SHIPPED — Sprint Ι Pod-γ)

`load_workbook(path, password=...)` decrypts OOXML-encrypted xlsx files via
the optional `msoffcrypto-tool` dependency (install with
`pip install wolfxl[encrypted]`). Modify mode + password also works in
this release; the on-save output is plaintext (write-side encryption is
documented T3 out-of-scope). See `tests/test_password_reads.py` for the
full coverage matrix.

### Phase 3 — T2 Rich-text reads (✅ SHIPPED in 1.3, Sprint Ι Pod-α)

Both reads and writes round-trip via the new
``wolfxl.cell.rich_text.{CellRichText, TextBlock, InlineFont}``
shims (matching openpyxl's iteration / equality protocol).

* ``Cell.rich_text`` always returns the structured runs (or ``None``
  for plain cells), regardless of how the workbook was opened.
* ``Cell.value`` keeps its prior contract by default — flattens
  rich-text to plain ``str`` so existing call sites are
  unaffected.  Pass ``load_workbook(..., rich_text=True)`` to
  flip ``Cell.value`` to return ``CellRichText`` for cells whose
  backing string carries `<r>` runs (matches openpyxl 3.x's own
  ``rich_text=True`` flag).
* Setting ``cell.value = CellRichText([...])`` round-trips in both
  write mode (native writer emits inline-string runs) and modify mode
  (patcher emits inline-string runs, SST left untouched).

### Phase 4 — T2 Streaming reads (SHIPPED — Sprint Ι Pod-β)

`load_workbook(path, read_only=True)` now activates a true SAX fast
path on `Worksheet.iter_rows`. The path is also auto-engaged for
sheets with > 50k rows even when the caller didn't opt in (see
`wolfxl._streaming.AUTO_STREAM_ROW_THRESHOLD`). Cells yielded in
streaming mode are
`wolfxl._streaming.StreamingCell` proxies that surface
`value`, `coordinate`, `row`, `column`, `font`, `fill`, `border`,
`alignment`, and `number_format` — every setter raises
`RuntimeError("read_only=True: ...")` immediately. Style attributes defer to the
eager native reader style table for O(1) lookups; the streaming layer parses
sheet XML directly via a hand-rolled byte scanner driven by
`quick-xml`-style lookahead, plus `xl/sharedStrings.xml` once at construction
time.

Implementation: `src/streaming.rs` (Rust SAX scanner), exposed via
`wolfxl._rust.StreamingSheetReader`; `python/wolfxl/_streaming.py`
(Python generator + `StreamingCell`). Tests:
`tests/test_streaming_reads.py` (16 cases), `tests/parity/
test_streaming_parity.py` (5 cases vs openpyxl `read_only=True`),
`tests/test_streaming_perf.py` (slow — 100k-row benchmark).

Documented divergences (out of scope for Pod-β, tracked elsewhere):

- ✅ FIXED in 1.5 (Sprint Λ Pod-γ) — Datetime cells previously surfaced as
  Excel serial floats in `values_only` mode and via `StreamingCell.value`.
  The streaming reader now consults the cell's style table on the fly
  (cached per `style_id` so the lookup is O(unique-styles)) and converts
  date-typed numeric cells via `wolfxl.utils.datetime.from_excel` —
  matching openpyxl's `read_only=True` semantics. Coverage:
  `tests/parity/test_streaming_parity.py::test_streaming_values_only_datetime_matches_openpyxl`,
  `::test_streaming_cell_value_datetime_matches_openpyxl`, and
  `tests/test_streaming_reads.py::test_streaming_datetime_yields_datetime_*`.
- Rich-text cells flatten to plain strings (matches the existing
  Phase 3 row).

### Phase 5 — T1 .xls / .xlsb (✅ SHIPPED in 1.4, Sprint Κ)

`load_workbook("foo.xlsb")` and `load_workbook("foo.xls")` ship in 1.4.
Use `NativeXlsbBook` for `.xlsb` and the independently authored MIT
`NativeXlsBook` for `.xls`. Read values and saved formula results; access
native `.xlsb` styles through `cell.font`, `.fill`, `.border`, `.alignment`
and `.number_format`. Expect `NotImplementedError` for `.xls` style access.
Binary-format `modify=True`, `read_only=True` and `password=` remain refused.
Keep Community Python ODS unsupported; use native ODS through Rust core/CLI.
For XLSB parity, use the committed JSON sidecars in `tests/parity/fixtures/xlsb/`.
For XLS parity, use `tests/parity/test_xls_reads.py` against the independent
xlwt fixture and literal expectations in `crates/wolfxl-reader/qualification/`.
See `Plans/rfcs/043-xlsb-xls-reads.md`.

## Per-fixture read gaps (surfaced by Phase 0 baseline run)

Phase 0's read-parity xfail list is now empty. Any newly discovered
fixture-specific drift should be added to `test_read_parity.py::KNOWN_FIXTURE_GAPS`
and documented here before the ratchet baseline is updated.

## Out of scope (documented, planned)

- ~~Writing encrypted xlsx~~ — ✅ SHIPPED in 1.5 (Sprint Λ Pod-α, RFC-044).
  `Workbook.save(path, password="...")` now emits an Agile (AES-256)
  encrypted file via `msoffcrypto-tool`'s high-level
  `OOXMLFile.encrypt()`. Standard (AES-128) and XOR remain
  decrypt-only in the upstream library and are deferred (no current
  customer ask). Install via `pip install wolfxl[encrypted]`.
- ~~Image construction~~ — ✅ SHIPPED in 1.5 (Sprint Λ Pod-β, RFC-045).
  `wolfxl.drawing.image.Image(...)` and `Worksheet.add_image(...)`
  are real, replacing the previous `_make_stub`. PNG / JPEG / GIF /
  BMP supported; one-cell, two-cell, and absolute anchors. Both
  write mode (native writer emits drawingN.xml + media + rels) and
  modify mode (patcher routes new images through `file_adds`).
- ~~Chart construction~~ — ✅ SHIPPED in 1.6 (Sprint Μ Pod-α/β/γ/δ,
  RFC-046) for the eight 2D chart families:
  `BarChart`, `LineChart`, `PieChart`, `DoughnutChart`, `AreaChart`,
  `ScatterChart`, `BubbleChart`, `RadarChart`, plus `Reference` and
  `Series`. `Worksheet.add_chart(chart, anchor)` accepts both string
  anchors and the RFC-045 anchor helper classes. Both write mode
  (native writer emits `xl/charts/chartN.xml` + drawing rels) and
  modify mode (patcher Phase 2.5l routes new charts through
  `file_adds`). RFC-035 `copy_worksheet` chart-aliasing limit is
  lifted — charts in copied sheets now deep-clone with cell-range
  re-pointing. See "Closed in 1.6 (Sprint Μ)" below.
- ~~**Pivot table construction**~~ — ✅ SHIPPED in 2.0 (Sprint Ν).
  See "Closed in 2.0 (Sprint Ν)" below.
- ~~**Pivot-chart linkage**~~ — ✅ SHIPPED in 2.0 (Sprint Ν Pod-δ,
  RFC-049). `chart.pivot_source = pt` on every chart family.
  See "Closed in 2.0 (Sprint Ν)" below.
- ~~**Slicers** (`xl/slicers/` + `xl/slicerCaches/`)~~ — ✅ SHIPPED
  in the post-PR #23 audit slice; covered by slicer copy,
  advanced-pivot bytes, and parity tests.
- ~~**Pivot calculated fields** (`<calculatedField>`)~~ — ✅ SHIPPED
  in the post-PR #23 audit slice.
- ~~**Pivot calculated items** (`<calculatedItem>`)~~ — ✅ SHIPPED
  in the post-PR #23 audit slice.
- ~~**Pivot GroupItems** (date / range grouping —
  `<fieldGroup base="N"><rangePr><groupItems/></rangePr></fieldGroup>`)~~
  — ✅ SHIPPED in the post-PR #23 audit slice.
- **OLAP / external pivot caches** (`xl/model/`, PowerPivot
  data-model) — out of scope permanently.
- **Pivot cache record regeneration after layout edits** — WolfXL
  mutates existing pivot source ranges, row/column/page field placement,
  page-field selection, and data-field aggregation. Layout edits stamp
  `refreshOnLoad="1"` and let Excel regenerate derived cache records on
  open rather than recalculating `pivotCacheRecords` inside WolfXL.
  openpyxl also does not provide a public cache-record regeneration
  engine, so this is not tracked as an openpyxl-parity gap.
- **OpenDocument (`.ods`)** — out of scope; not on the roadmap.
  Detected and rejected by `_rust.classify_format` with a friendly
  pointer.

## Closed in 2.0 (Sprint Ν)

The construction-side parity roadmap closes here. Every openpyxl
3.1.x construction idiom now has a wolfxl equivalent.

- ✅ **Pivot cache construction** (Sprint Ν Pod-α / β / γ,
  RFC-047). New crate `crates/wolfxl-pivot` ships
  `PivotCache`, `CacheField`, `SharedItems`, `CacheValue`,
  `CacheRecord`, `RecordCell`, `WorksheetSource` typed Rust
  models plus deterministic
  `emit::pivot_cache_definition_xml` and
  `emit::pivot_cache_records_xml`. Python
  `wolfxl.pivot.PivotCache(source=Reference(...))` walks the
  source range on `wb.add_pivot_cache(cache)`, infers per-column
  type (string / number / date / boolean / mixed) per RFC-047
  §10.9, builds `SharedItems`, and emits both the cache
  definition and the cache records part. **WolfXL is the first
  Python OOXML library to construct
  `pivotCacheRecords{N}.xml` from scratch** — pivots open in
  Excel / LibreOffice / openpyxl with data populated, no
  refresh-on-open required. (openpyxl preserves records on
  round-trip but doesn't construct them; XlsxWriter doesn't
  support pivots at all.)

  - ✅ Rust `wolfxl-pivot` crate model + emit (Pod-α)
  - ✅ Python `wolfxl.pivot.PivotCache` + `to_rust_dict()` /
    `to_rust_records_dict()` (Pod-β)
  - ✅ Patcher Phase 2.5m + `Workbook.add_pivot_cache` public
    API + PyO3 bindings (Pod-γ)
  - ✅ Workbook splice (`<pivotCaches>` + rels-of-type
    `pivotCacheDefinition`) via `PartIdAllocator` (Pod-γ)

- ✅ **Pivot table construction** (Sprint Ν Pod-α / β / γ,
  RFC-048). Replaces the v0.5+
  `_make_stub` in `python/wolfxl/pivot/__init__.py` with real
  classes: `PivotTable`, `PivotField`, `DataField`, `RowField`,
  `ColumnField`, `PageField`, `PivotItem`, `Location`,
  `PivotTableStyleInfo`, `PivotSource`. Bare-string axis specs
  (`rows=["region"]`) and explicit-builder axis specs
  (`rows=[RowField("region", custom_caption="Region")]`) both
  work. 11 aggregator functions on `DataField`: sum, count,
  average, max, min, product, count_nums, std_dev, std_dev_p,
  var, var_p. The Option A core — pre-computed `<rowItems>` /
  `<colItems>` aggregation in
  `python/wolfxl/pivot/_table.py` — is what makes the pivot
  open with data populated.
  `Worksheet.add_pivot_table(pt, anchor)` and the patcher's
  Phase 2.5m wire it all together.
  `tests/parity/openpyxl_surface.py`:
  `wolfxl.pivot.PivotTable` flips to
  `wolfxl_supported=True`. New `tests/test_pivot_construction.py`
  pins the construction surface (40+ tests). The
  `test_compat_shims.py` ratchet is updated:
  `wolfxl.pivot.PivotTable` no longer raises
  `NotImplementedError`.

- ✅ **Pivot-chart linkage** (Sprint Ν Pod-δ, RFC-049).
  `chart.pivot_source = pt` on every one of the 16 chart
  families (the attribute lives on `ChartBase`, so `BarChart`,
  `LineChart`, `PieChart`, `BarChart3D`, `StockChart`,
  `ProjectedPieChart`, etc. all gain the same setter).
  Validates `name` against
  `^([A-Za-z_][A-Za-z0-9_]*!)?[A-Za-z_][A-Za-z0-9_ ]*$`;
  validates `fmt_id` in `[0, 65535]`. Rust emit
  (`crates/wolfxl-writer/src/emit/charts.rs`) inserts
  `<c:pivotSource>` between `<c:chart>` and `<c:plotArea>` per
  ECMA-376 §21.2.2.158, plus `<c:fmtId val="0"/>` on every
  `<c:ser>`. When a pivot-bearing chart is `copy_worksheet`'d,
  the cloned chart's `pivot_source.name` is rewritten to match
  the cloned pivot table's `displayName`.

- ✅ **RFC-035 deep-clone of pivot-bearing sheets** (Sprint Ν
  Pod-γ, RFC-047 §6 + RFC-048 §6). Lifts the v1.6 limit at
  `Plans/rfcs/035-copy-worksheet.md` §10 ("sheets with pivot
  tables raise a clear error on copy"). Cloned pivot tables get
  fresh `pivotTable{N}.xml` IDs via `PartIdAllocator`; the
  source's pivot cache is **aliased** (one cache serves the
  source and the clone, mirroring the existing image-media
  alias pattern); the `Reference`-encoded source-range hint on
  the cache is re-pointed to the clone's sheet name (mirrors
  the chart deep-clone pattern from v1.6 / Sprint Μ). Existing
  RFC-035 cross-RFC composition tests (#1-#6) remain green.

- ✅ **RFC-046 §13 legacy chart-dict key removal** (Sprint Ν
  Pod-α / Pod-ε). Deprecated in v1.7 (documentation-only); the
  Rust parser's accept-also for the legacy
  `fill_color` / `line_color` / `line_dash` / `line_width_emu`
  keys is **removed in v2.0**. Only the §10.9
  `solid_fill` + nested `ln: {solid_fill, w_emu, prst_dash}`
  form is accepted. The Python emitter has used the §10.9 form
  exclusively since v1.6.1, so only out-of-tree callers that
  bypassed `Worksheet.add_chart` and built chart dicts by hand
  are affected (and those callers were warned in the v1.7
  release notes).

- ✅ **README rewrite** (Sprint Ν Pod-ε, RFC-054; post-PR #23 truth
  pass refreshed wording). Uses openpyxl-compatible positioning with
  pivot construction, adds "Pivot tables in 6 lines", and withholds
  benchmark headline claims until the v2.0 ExcelBench refresh is
  complete. Refreshes feature matrix (pivots ✅, charts ✅) and
  ecosystem comparison.

- ✅ **Compatibility Matrix v2.0** (Sprint Ν Pod-ε, RFC-054).
  Pivot row flips ❌ → ✅ with full sub-table covering import
  paths, public APIs, aggregator functions, axis-spec forms,
  chart linkage, deep-clone, and v2.1+ deferred items.
  Ecosystem comparison gains a Pivots column showing wolfxl is
  the only library that constructs pivots with pre-aggregated
  records (caveat: openpyxl preserves on round-trip).

- ✅ **Migration guide** (Sprint Ν Pod-ε, RFC-054).
  `docs/migration/openpyxl-migration.md` gains a "Pivot tables
  (Sprint Ν / v2.0)" section covering the 6-line snippet,
  chart linkage, openpyxl→wolfxl import-path mapping, two-step
  cache + table API explanation, 11 aggregator functions, v2.1+
  limits, and the empty-cache caveat.

- ✅ **Launch posts finalized** (Sprint Ν Pod-ε, RFC-054).
  `Plans/launch-posts.md` updated for v2.0 with pivot snippets
  in HN / Twitter / r/Python / dev.to / GH Discussions drafts;
  pre-launch and post-launch checklists.

## Closed in 1.6 (Sprint Μ)

- ✅ **Chart construction — eight 2D chart families** (Sprint Μ
  Pod-α/β, RFC-046). The `_make_stub` definitions at
  `python/wolfxl/chart/__init__.py` for `BarChart`, `LineChart`,
  `PieChart`, `ScatterChart`, `AreaChart`, `Reference`, and `Series`
  are replaced by real classes. Two new types ship in 1.6 alongside
  the original five: `DoughnutChart` (subclasses PieChart) and
  `BubbleChart` and `RadarChart`. Per-type unique features land at
  full openpyxl 3.1.x depth: bar `gap_width`/`overlap`/`grouping`/
  `bar_dir`; line `smooth`/`up_down_bars`/`drop_lines`/`hi_low_lines`/
  per-series `marker`; pie `vary_colors`/`first_slice_ang`; doughnut
  `hole_size`; area `grouping`/`drop_lines`; scatter `scatter_style`;
  bubble `bubble_3d`/`bubble_scale`/`show_neg_bubbles`/`size_represents`;
  radar `radar_style`. Both write mode (native writer emits
  `xl/charts/chartN.xml` via `crates/wolfxl-writer/src/emit/charts.rs`)
  and modify mode (patcher Phase 2.5l routes new charts through
  `file_adds`). New `RT_CHART` const in `crates/wolfxl-rels/src/lib.rs`.
  `crates/wolfxl-writer/src/emit/drawings.rs` extended for
  `<xdr:graphicFrame>` alongside the RFC-045 `<xdr:pic>` block (a
  worksheet with both an image and a chart shares a single drawing
  part).
- ✅ **`Worksheet.add_chart(chart, anchor)`** (Sprint Μ Pod-β).
  Accepts a `ChartBase` subclass and either a coordinate string
  (one-cell anchor pinned to the top-left of that cell, sized via
  `chart.width` / `chart.height`) or one of the RFC-045 anchor
  helper classes (`OneCellAnchor`, `TwoCellAnchor`, `AbsoluteAnchor`).
- ✅ **RFC-035 chart-deep-clone with cell-range re-pointing**
  (Sprint Μ Pod-γ). The deferred limit at
  `Plans/rfcs/035-copy-worksheet.md` lines 924-929 is lifted.
  `copy_worksheet` now deep-clones chart parts and re-points the
  cell-range references on every series. Self-references
  (`SourceSheet!$B$2:$B$10`) are rewritten to the copy's sheet name;
  cross-sheet references (`Other!$A$1:$A$5`) are preserved verbatim.
  Cached values (`<c:strCache>` / `<c:numCache>`) are preserved as-is
  and Excel rebuilds them on next open if stale. The new behaviour is
  the default; there is no opt-in flag (the old aliasing was a known
  limit, not a deliberate contract).
- ✅ **Modify-mode `add_chart`** (Sprint Μ Pod-γ). XlsxPatcher Phase
  2.5l drains the queued chart adds per sheet, allocates fresh
  `chartN.xml` and `drawingN.xml` numbers via `PartIdAllocator`,
  emits chart bytes through `file_adds`, and splices new
  `<xdr:graphicFrame>` blocks into the sheet's existing drawing part
  (or creates one if the sheet had no drawing yet). Composes cleanly
  with RFC-045 `add_image` and RFC-035 `copy_worksheet`: a single
  drawing part can contain image `<xdr:pic>` blocks, chart
  `<xdr:graphicFrame>` blocks, and any mix of the two.

### Deferred to 1.6.1 — ✅ Closed (see "Closed in 1.6.1" below)

The four families flagged here in 1.6.0 (3D variants, Stock,
Surface 2D + 3D, ProjectedPieChart) shipped in 1.6.1 as real
classes. See "Closed in 1.6.1 (Sprint Μ-prime)" immediately
below for the receipts.

## Closed in 1.7 (Sprint Ξ)

- ✅ **`Worksheet.remove_chart(chart)` + `Worksheet.replace_chart(old, new)`** (Sprint Ξ Pod-α, RFC-046 §14). New on
  `Worksheet`; mirror openpyxl's `ws._charts.remove(...)` and the swap-in-place idiom.
  Pending-list scope only in v1.7; modify-mode removal of charts that
  survive from the source workbook is a v1.8 follow-up.
  Verified by `tests/test_charts_remove.py` (7 tests).
- ✅ **`chart.title = RichText(...)` accepted** (Sprint Ξ Pod-α,
  RFC-046 §15). `TitleDescriptor.__set__` extended to accept
  wolfxl-typed and openpyxl-typed `RichText`. Coerces openpyxl's
  `ColorChoice`-typed `solidFill` to the hex string the Rust
  emitter expects. Closes the lone v1.6.1 `xfail` in
  `tests/test_charts_write.py::test_line_chart_title_rich_text`.
- ✅ **Version drift** (Sprint Ξ Pod-α). `pyproject.toml` and
  `Cargo.toml` synced from `0.5.0` to `1.7.0`. `wolfxl.__version__`
  now reports `1.7.0` (re-exported from `CARGO_PKG_VERSION` via
  `src/lib.rs:41`). PyPI classifier promoted to
  `Development Status :: 5 - Production/Stable`.
- ✅ **Docs refresh — `docs/migration/` and `docs/performance/`**
  (Sprint Ξ Pod-β + Pod-γ, RFC-051 + RFC-052). Compatibility
  Matrix and openpyxl-migration guide rewritten for v1.7;
  benchmark-results refreshed with read / write / modify-mode /
  chart-construction speedup tables on a 1k / 10k / 100k row
  matrix.
- ✅ **Public launch posts** (Sprint Ξ Pod-δ, RFC-053).
  `Plans/launch-posts.md` materialised with HN, Twitter/X,
  r/Python, dev.to, GitHub Discussions drafts + pre/post-launch
  checklist.

## Closed in 1.6.1 (Sprint Μ-prime)

- ✅ **Chart 3D / Stock / Surface / ProjectedPie variants** (Sprint
  Μ-prime Pod-β′, RFC-046 §11). The 8 deferred classes ship as real
  implementations: `BarChart3D`, `LineChart3D`, `PieChart3D` (alias
  `Pie3D`), `AreaChart3D`, `SurfaceChart`, `SurfaceChart3D`,
  `StockChart`, `ProjectedPieChart`. Each carries the appropriate
  `view_3d` defaults per §11.1; `StockChart` validates its 4-series
  OHLC ordering at construction; `ProjectedPieChart` exposes
  `of_pie_type`, `split_type`, `split_pos`, `second_pie_size`.
  Verified by `tests/test_charts_3d.py` (Pod-β′/Pod-γ′).
- ✅ **Sprint Μ chart-dict contract gap (Pod-α ↔ Pod-β)** (Sprint
  Μ-prime Pod-α′ + Pod-β′, RFC-046 §10). Pod-α's flat-key
  `parse_chart_dict` and Pod-β's `to_rust_dict()` now emit/consume
  the canonical §10 shape verbatim. The 37 advanced sub-feature
  tests in `tests/test_charts_write.py` (gridlines, error bars,
  trendlines, vary_colors, non-default grouping, scatter style,
  manual layout, marker symbol, fill color, title runs,
  invalid-input rejection) flip from xfail → pass. Pod-α′ also
  added 8 new `ChartKind` variants and the `serialize_chart_dict`
  PyO3 helper that Pod-γ′ uses to bridge modify-mode high-level
  `Worksheet.add_chart()` to the patcher.
- ✅ **Modify-mode high-level `Worksheet.add_chart()`** (Sprint
  Μ-prime Pod-γ′, RFC-046 §10.12). The v1.6.0 warn-and-drop
  fallback in `Workbook._flush_pending_charts_to_patcher` is
  replaced with a real dict→bytes bridge via
  `serialize_chart_dict`. The bytes-level escape hatch
  `Workbook.add_chart_modify_mode` continues to work unchanged.

## Closed in 1.5 (Sprint Λ)

- ✅ **Write-side OOXML encryption** (Sprint Λ Pod-α, RFC-044).
  `Workbook.save(path, password="...")` now encrypts on the way out
  via `msoffcrypto-tool`'s high-level `OOXMLFile.encrypt()`. Agile
  (AES-256 / SHA-512) is the only algorithm shipped; Standard
  (AES-128) and XOR remain decrypt-only in the upstream library.
  Lifts the `NotImplementedError` at `python/wolfxl/_workbook.py:1032`.
  `wolfxl[encrypted]` extra now covers writes too (was already on
  reads from Sprint Ι Pod-γ). Closes the long-standing T3
  out-of-scope row.
- ✅ **Image construction** (Sprint Λ Pod-β, RFC-045).
  `wolfxl.drawing.image.Image(...)` and `Worksheet.add_image(...)`
  are real, replacing the `_make_stub` at
  `python/wolfxl/drawing/image.py`. Supports PNG / JPEG / GIF / BMP;
  one-cell, two-cell, and absolute anchors. Both write mode (native
  writer emits drawingN.xml + media + rels) and modify mode
  (patcher routes new images through `file_adds`). Targets full
  openpyxl `Image()` parity.
- ✅ **Streaming-datetime fix** (Sprint Λ Pod-γ). The Phase 4
  divergence note in the streaming-reads section is closed —
  `iter_rows(values_only=True)` now returns Python `datetime`
  objects for date-formatted cells, matching openpyxl's
  `read_only=True` contract. The streaming reader consults the
  styles table for the cell's number format and converts Excel
  serial floats inline. (See the streaming-reads divergence note
  immediately above for the original behavior.)

## Closed in 1.3 (Sprint Ι Pod-α)

- ✅ **Rich-text read** — Phase 3 row above is now SHIPPED.
- ✅ **Rich-text write** — was previously listed as out-of-scope T3.
  Sprint Ι Pod-α shipped both write-mode (native writer) and
  modify-mode (patcher) inline-string emit paths, so user code that
  builds a workbook with rich-text cells round-trips end-to-end via
  wolfxl's writer.  The SST is intentionally left untouched — runs
  are emitted as inline strings (`t="inlineStr"` + `<is>`), matching
  openpyxl's own write path verbatim.

## Modify mode — T1.5 audit (now closed) and structural extensions

Modify mode (`load_workbook(path, modify=True)`) is served by `XlsxPatcher`,
which surgically rewrites changed parts and copies everything else verbatim.
The W4F audit originally enumerated seven mutation paths that were deferred
to a post-Wave-5 T1.5 slice. **All seven shipped in WolfXL 1.1's Phase 3**
(per `tests/test_modify_mode_independence.py` lines 14-21 and the per-RFC
modify-mode test files). Structural mutations and `copy_worksheet` followed.
This table is now a status snapshot, not a deferred-work list.

| Modify-mode mutation | Status |
|---|---|
| `wb.properties.title = ...` (any property mutation) on existing file | ✅ Shipped — RFC-020 (`tests/test_modify_properties.py`, `tests/test_workbook_properties_t1.py`) |
| `wb.defined_names[name] = DefinedName(...)` on existing file | ✅ Shipped — RFC-021 (`tests/test_defined_names_modify.py`). The `__setitem__` proxy is exposed; round-trip verified end-to-end. |
| `cell.comment = Comment(...)` | ✅ Shipped — RFC-023 (`tests/test_comments_modify.py`) |
| `cell.hyperlink = Hyperlink(...)` | ✅ Shipped — RFC-022 (`tests/test_modify_hyperlinks.py`, `tests/test_hyperlink_internal_flag.py`) |
| `ws.add_table(Table(...))` | ✅ Shipped — RFC-024 (`tests/test_tables_modify.py`) |
| `ws.data_validations.append(...)` | ✅ Shipped — RFC-025 (`tests/test_modify_data_validations.py`) |
| `ws.conditional_formatting.add(...)` | ✅ Shipped — RFC-026 (`tests/test_modify_conditional_formatting.py`) |
| Sheet/column/row structural mutations | ✅ Shipped — `insert_rows`/`delete_rows` (RFC-030), `insert_cols`/`delete_cols` (RFC-031), `Worksheet.move_range` (RFC-034), `Workbook.move_sheet` (RFC-036). |
| `wb.copy_worksheet(...)` (modify mode) | ✅ Shipped — RFC-035 in 1.1 (Sprint Ζ Pod-δ closed four of six composition gaps; two remain xfail per "RFC-035 cross-RFC composition gaps" below). See divergence section below. |
| `wb.copy_worksheet(...)` (write mode) | ✅ Shipped — Sprint Θ (1.2) Pod-C1 lifts the §3 OQ-a `NotImplementedError`. |

Supported in modify mode (round-trips cleanly via `_flush_to_patcher`):

- Cell values: string, number, boolean, formula, blank
- Font (bold/italic/underline/strikethrough/size/name/color)
- Fill (solid pattern bg color)
- Alignment (horizontal/vertical/wrap/indent/rotation)
- Number format
- Borders (left/right/top/bottom — style + color)

`tests/test_modify_mode_independence.py` encodes these contracts as
pre-rip-out invariants: any future change that breaks the patcher's
independence from the writer backend, or that silently falls through to
the writer for a T1.5-deferred feature, fails CI immediately.

## RFC-035 — `copy_worksheet` divergences from openpyxl (SHIPPED 1.1)

`Workbook.copy_worksheet` ships in modify mode in WolfXL 1.1
(RFC-035). The behaviour deliberately diverges from openpyxl's
`WorksheetCopy` in five places. Each divergence is asserted by
`tests/parity/test_copy_worksheet_parity.py`; this section is the
ratchet-tracked record of WHY wolfxl preserves what openpyxl drops.

| Feature | WolfXL behaviour | openpyxl behaviour | Rationale |
|---|---|---|---|
| Tables | Cloned with auto-renamed `name`/`displayName` (`{base}_{N}`, N starts at 2 per RFC-035 §3 OQ-b). New `<table id>`, new content-type `<Override>`, new rels entry. | `WorksheetCopy._copy_cells` walks the in-memory cell dict only; `<tableParts>` is silently dropped. | wolfxl operates on ZIP bytes, not an in-memory model — clone preserves the source's full feature surface. |
| Data validations | Cloned in-place inside the cloned sheet's XML. | Dropped — `WorksheetCopy` has no DV-handling branch. | DV is part of "what makes the sheet work"; dropping it silently degrades cloned templates. |
| Conditional formatting | Cloned in-place inside the cloned sheet's XML (with cross-sheet `dxfId` allocation). | Dropped — same reason as DV. | Same as DV. |
| Sheet-scoped defined names | Fresh entries emitted with `localSheetId == new_idx` (post-copy tab position) per §3 OQ-c. Source's sheet-scope names retained. | Dropped — `WorksheetCopy` does not touch `xl/workbook.xml`'s `<definedNames>`. | `_xlnm.Print_Area` is the canonical "make-this-sheet-printable" hint; dropping it silently breaks print-preview on the clone. |
| Image media | Aliased — cloned drawing rels point at the same `xl/media/imageN.png` as the source. | Deep-copied — pillow re-encodes the image binary on the clone. | Avoids 50× bloat on workbooks with logo images and many sheet copies. RFC-035 §5.3 documents the contract; future "modify a copy's image" RFC will deep-clone. |
| Calc chain (`xl/calcChain.xml`) | Not mutated — Excel rebuilds it on next open. | Same. | calcChain is a perf optimization, not a correctness contract. |

### RFC-035 cross-RFC composition gaps

Surfaced by Pod-γ's full harness (`tests/test_copy_worksheet_modify.py`,
originally six xfail cases). Pod-δ closed four of the six in
Sprint Ζ; Sprint Θ Pod-A closed bug #4 via the new
``permissive=True`` loader flag. One case (#6) remains as a
documented 1.2 follow-up.

#### Fixed in 1.1 (Sprint Ζ Pod-δ)

- ✅ **#1 `test_i_copy_and_edit_copy_in_same_save`** —
  fixed by Pod-δ commit `fix(rfc-035): patch cloned-sheet bytes through
  file_adds, not zip`. Phase 3 now reads cloned-sheet bytes from
  `file_adds` / `file_patches` first, falling back to the source ZIP
  only for genuine source-side sheets, and routes the rewrite back to
  `file_adds` for cloned paths so Phase 4's new-entry pass picks up the
  patched bytes. Test flips xfail (strict, OSError) → PASS.
- ✅ **#2 `test_j_copy_then_move_sheet_in_same_save`** —
  fixed by Pod-δ commit `fix(rfc-035): seed Phase 2.5h workbook.xml
  read from file_patches`. Phase 2.5h's reorder pass now prefers
  `file_patches["xl/workbook.xml"]` over the source-ZIP read so the
  Phase 2.7 → Phase 2.5h handoff happens through the shared
  `file_patches` map (the intended composition per RFC-035 §5.4).
  Test flips xfail (strict, AssertionError) → PASS.
- ✅ **#3 `test_k_copy_then_add_table_to_copy`** —
  fixed by the same commit as #1 (shared root cause). The Phase 2.5f
  rels-graph load also probes `file_adds` / `file_patches` first so a
  user `add_table` on a cloned sheet sees the cloned rels graph
  rather than an empty fallback. Test flips xfail (strict, OSError)
  → PASS.
- ✅ **#5 `test_q_defined_names_upsert_collision`** —
  fixed by Pod-δ commit `fix(rfc-035): route cloned defined names
  through RFC-021 merger`. Phase 2.7's `defined_names_to_add` push
  now scans `queued_defined_names` for a matching
  `(name, local_sheet_id)` key and skips on hit so the user's
  explicit upsert wins over the planner's default (per RFC-035 §5.4
  and Pod-β's last-write-wins-on-the-USER invariant). Test flips
  xfail (strict, AssertionError) → PASS.

#### Fixed in 1.2 (Sprint Θ Pods A + B)

- ✅ **#4 `test_p_self_closing_sheets_block`** —
  fixed by Sprint Θ Pod-A commit `fix(rfc-035): add permissive=True
  loader mode, close bug #4 self-closing <sheets/>`.
  `wolfxl.load_workbook(..., permissive=True)` now falls back to the
  workbook rels graph when `xl/workbook.xml`'s `<sheets>` block is
  empty / self-closing: each worksheet relationship target is
  registered under a synthesized title (`Sheet1`, `Sheet2`, ...) and
  the in-memory workbook.xml is rewritten to expose the synthesized
  `<sheet>` entries so downstream phases (Phase 2.7 splice, defined-
  names merger) see a well-formed workbook. The flag defaults to
  `False`; well-formed inputs are unaffected. Test flips xfail →
  PASS end-to-end (load → copy_worksheet → save → reload via
  openpyxl).

- ✅ **#6 `test_r_cdata_pi_fuzz_fakeout`** —
  fixed by Sprint Θ Pod-B commit `fix(rfc-035): replace naive splice
  with quick-xml SAX scan, close bug #6 CDATA fakeout`. The Phase 2.7
  `splice_into_sheets_block` helper now drives a `quick_xml::Reader`
  over `xl/workbook.xml` and locates the real `<sheets>` open/close
  by event-stream nesting depth rather than byte-substring search.
  Comments, CDATA sections, and processing instructions surface as
  separate quick-xml events and are ignored, so a workbook.xml
  comment containing the literal `</sheets>` token no longer
  perturbs the splice point. Five new Rust unit tests pin the
  invariant (normal, self-closing, comment fakeout, CDATA fakeout,
  malformed). Test flips xfail → PASS.

  As a side-fix, the test fixture helper
  `_inject_comment_with_sheets_token` was anchoring on `?>` (XML
  declaration close) — but openpyxl-saved workbooks omit the XML
  decl, so the injection was a silent no-op that masked the bug.
  The helper now anchors on the `<workbook ...>` opening tag via
  regex.

#### Deferred queue (post-1.2)

All RFC-035 cross-RFC composition gaps surfaced in Sprint Ζ have been
closed by Sprints Ζ (Pod-δ: #1, #2, #3, #5) and Θ (Pods A+B: #4, #6).
No further deferred items remain in this category.
