# Changelog

## 2.0.11 - 2026-10-24

### Security

- Remove `calamine-styles` after the native reader cutover and move direct XML
  parsing and the isolated reader qualifier to `quick-xml` 0.42.
- Upgrade ZIP handling to `zip` 9.0.5, clearing RUSTSEC-2026-0367,
  RUSTSEC-2026-0368 and RUSTSEC-2026-0369 (patched in 9.0.0).
- Exclude Python 3.9.0 and 3.9.1, which cannot install the fixed cryptography
  resolution for the encrypted extra. Python 3.9.2 and later remain supported.
- Keep Community Python ODS refusal and existing binary-format read, style, and
  modification boundaries.
- The optional `test` extra on Python 3.9 and 3.10 still resolves Pillow releases
  affected by CVE-2026-96639 and GHSA-78g4-xqwv-q3g8, because fixed Pillow 13
  requires Python 3.11. Runtime installs do not depend on Pillow.

### Fixed

- Calling `calculate()` on a workbook opened with `data_only=True` preserves
  its existing formula caches when saving, and writes caches only for formulas
  actually computed by that call.
- Preserve indexed streaming merge lookup and merged endpoint borders without
  materializing worksheet cells.
- Match openpyxl's inline-font boolean defaults and XML serialization.
- Build supported CPython 3.9 wheels explicitly and require the complete
  CPython 3.9 through 3.14 wheel matrix without dropping published platforms.
- Keep existing public benchmark test helpers available to the Community
  source-export regression tests.
- Complete the isolated native-reader qualification dependency closure.
- Pin historical evidence and contribution links to the public source snapshot.

## 2.0.10 - 2026-10-20

### Fixed

- `calculate()` followed by `save()` now writes supported formula results to
  the saved workbook's cached values, in both write and modify mode, so
  `load_workbook(path, data_only=True)` and other cache readers see them.
  Results that cannot be computed stay uncached, and editing inputs, formulas,
  sheet names, or defined names after `calculate()` requires calculating again
  before saving. After `insert_rows`, `delete_rows`, `insert_cols`,
  `delete_cols`, or `move_range` in modify mode, that workbook object writes
  no calculated results; reopen the saved file and calculate it to cache them.
  In 2.0.9 and earlier the saved caches were left unchanged.
- Shallow copies of implicit blank cells in streaming reads are preserved.

### Changed

- Faster reads, saves, and streaming: lazy merge metadata, shared styles in
  bounded read windows, spliced in-place cell patches, and a content-types fast
  path for unchanged templates. Measured receipts are in `benchmarks/`.

## 2.0.9 - 2026-10-03

### Added

- A reproducible one-cell workbook editing study in `study/agent-edit/`, with
  public fixtures, verified runtime downloads, offline regression tests, and
  separate Community plus LibreOffice and Commercial native calculation workflows.

### Fixed

- Changes to `Workbook.calculation` are saved in modify mode. Setting
  `fullCalcOnLoad = True` writes the recalculation-on-open flag while
  preserving untouched package parts.
- Scalar and style edits preserve the existing calculation chain and its
  relationships and content types. Calculation metadata is rebuilt only
  when formula membership or sheet structure changes.

### Changed

- README, PyPI copy, and the package summary now cite the Community 2.0.8
  versus openpyxl 3.1.5 benchmark: 3.0-17.3x faster in measured reads and
  writes on Apple M5 Pro, median of five rounds. Old cross-library read
  speed claims are removed.
- Documented the free recalculation-on-open recipe and the stale-cache
  boundary in the README and workbook-editing skill.

## 2.0.8

### Fixed

- In modify mode, a value written to a cell that does not exist yet takes
  the style Excel gives a cell typed there: the row's style when the row is
  formatted (`customFormat="1"`), else the column's `<col style>`, else the
  workbook default. New cells previously always got the workbook default
  style. Cells that already exist keep their own style. Expanding an empty
  `<row/>` to hold a new cell also keeps the row's height, style, and other
  attributes.
- Sheet-scoped defined names load with their `localSheetId` and their
  formula text exactly as stored, and names added through
  `ws.defined_names.add(...)` are saved with that worksheet's scope.
- `Workbook.remove()` in write mode re-indexes sheet-scoped names, so print
  titles, print areas, and local names set after removing a sheet point at
  the right sheet.
- Loaded tables expose their `autoFilter`, as openpyxl does.
- `ws.add_chart()` accepts `OneCellAnchor`, `TwoCellAnchor`, and
  `AbsoluteAnchor` objects (set on `chart.anchor` or passed as the anchor)
  in write mode, and saves the ranged placement.

## 2.0.7

### Fixed

- `calculate()` evaluates the `^` operator. Formulas using it, and every
  formula that depended on them, previously returned no value. `^` binds
  tighter than `*` and `/`, associates left to right, and returns `#DIV/0!`
  for zero to a negative power and `#NUM!` for a negative base with a
  fractional exponent or an overflowing result, as Excel does.
- `YEAR`, `MONTH`, `DAY`, `EDATE`, `EOMONTH`, `DAYS`, `HOUR`, `MINUTE`, and
  `SECOND` accept cells that hold dates. They previously returned no value
  for date-formatted inputs.
- Date serials follow the workbook's date system. In workbooks that use the
  1904 date system, `DATE`, `TODAY`, `NOW`, and functions that read date
  cells return the serials Excel returns instead of 1900-system serials.

## 2.0.6

### Changed

- The README, PyPI description, and package summary lead with preservation
  and speed (7-14x faster than openpyxl on most reads and writes) and describe
  `calculate()` as covering common functions only. Metadata only; the
  package code is unchanged from 2.0.5.

## 2.0.5

### Added

- Prebuilt musllinux wheels for x86_64 and aarch64, so `pip install wolfxl`
  on Alpine and other musl-based distributions no longer builds from source.

### Changed

- The README, PyPI description, and package summary now lead with editing
  existing Excel files without losing formatting, with an edit example and an
  openpyxl comparison.

## 2.0.4

### Fixed

- Tables loaded in modify mode now grow on save. Changing `table.ref` and
  appending `TableColumn` entries rewrites the existing table part: its range,
  its autoFilter range, and its column list. Previously these edits were
  silently dropped. Moving a table's top-left cell, or a range whose width
  does not match the column list, raises `ValueError` on save.

## 2.0.3

### Fixed

- `calculate()` and `recalculate()` no longer raise `TypeError` on loaded
  workbooks that contain defined names. The evaluator now reads each name's
  reference text, so formulas such as `=SUM(Sales)*Rate` calculate and
  propagate changes through the name.

## 2.0.2

### Changed

- Corrected PyPI release copy: removed obsolete pre-launch language and pinned the
  installation command to `wolfxl==2.0.2`.
- Canonicalized public GitHub links to `SynthGL/wolfxl-oss` after the repository
  rename.
- Updated package metadata and documentation only; this release contains no code
  changes.
