# Changelog

## Unreleased

### Performance

- Read-only `iter_rows(values_only=True)` reads ordinary worksheet rows with
  a byte-level row scanner and hands any row it does not recognize to the XML
  event reader, so output is unchanged. Sheets of one workbook now share a
  single parsed shared-string table instead of parsing it per sheet, and the
  ZIP reader no longer scans the whole file to locate the central directory.
  Reading the formulas of a 4-sheet, 200,000-row workbook drops from 0.61 s
  to 0.30 s.

## 2.0.26 - 2026-10-29

### Performance

- Read-only `iter_rows(values_only=True)` and `iter_value_chunks()` now
  resolve formula cells in the native reader. With `data_only=True`, cached
  formula results no longer cost more than the same constant values; formula
  text reads (`data_only=False`) are also faster.

### Fixed

- A formula cell whose cached `<v>` comes before its `<f>` now reads its
  numeric cached result in `data_only=True` mode instead of `None`, as
  openpyxl does.

## 2.0.25 - 2026-10-28

### Fixed

- On a read-only sheet with no `<dimension>` element, or after
  `ws.reset_dimensions()`, `ws.iter_value_chunks()` now yields the same rows
  as `ws.iter_rows(values_only=True)` and openpyxl. With `data_only=True` and
  an explicit `max_row`, rows keep their own lengths instead of being padded
  to the widest column. A missing row is now `[]`, as in `ws.values`, instead
  of `()`.

## 2.0.24 - 2026-10-28

### Fixed

- `ws.iter_value_chunks()` in `read_only=True, data_only=True` mode now
  yields the same rows as `ws.iter_rows(values_only=True)` for sheets
  without value cells. With an explicit `max_row`, an empty, formatting-only,
  or merge-only sheet yields the padded blank rows instead of none. Without
  bounds, a merge-only sheet yields its blank rows, as openpyxl does.
- In the same mode, a merged range that reaches the declared `<dimension>`
  with no stored rows for it no longer pads blank rows past the last stored
  row. The decision to pad now uses value cells only, and no longer parses
  the sheet a second time to get its bounds.

## 2.0.23 - 2026-10-28

### Fixed

- Iterating a worksheet that has no cells now yields no rows, matching
  openpyxl. Before, `ws.values`, `ws.iter_rows()`, and `ws.iter_cols()`
  without bounds yielded one row of `None` for an empty sheet, in normal and
  `read_only=True` mode.
- A sheet whose declared `<dimension>` is larger than its stored cells no
  longer reports or iterates the declared range in normal mode: `max_row`,
  `max_column`, `dimensions`, and unbounded iteration follow the stored cells,
  as in openpyxl. In `read_only=True` mode, an unbounded scan stops at the
  last stored row instead of padding blank rows up to the declared dimension.
  Explicit `min_row`/`max_row` bounds still yield every row in the range.

## 2.0.22 - 2026-10-27

### Changed

- The package summary and PyPI page now cite the WolfXL 2.0.21 versus
  openpyxl 3.1.5 benchmark on Linux (6.7-12.0x faster on reads and writes)
  instead of the 2.0.8 results. No library code changed.

## 2.0.21 - 2026-10-27

### Fixed

- Pivot tables with calculated fields or calculated items now open in
  Microsoft Excel without a repair prompt and show their values on first open
  (#64). A calculated field is written as a formula cache field, and a
  calculated item is written on the pivot cache as Excel stores it, with its
  formula naming items as `Field[Item]`. Before, a calculated field made Excel
  repair the workbook and a calculated item was dropped.
- Calculated items now render after the field's regular items and count
  toward subtotals and grand totals, on row and column fields. Calculated
  fields render their formula over each cell's field sums, including formulas
  that read other calculated fields. Formulas WolfXL does not evaluate, such as
  ones that call functions, are left blank and the cache is marked to refresh
  on load so Excel fills them in.
- Calculated items now belong to the pivot cache, so every pivot table on the
  cache shows them. `PivotCache.add_calculated_item()` adds one;
  `PivotTable.add_calculated_item()` remains and adds it to the table's cache;
  read the items from `PivotCache.calculated_items`, which replaces
  `PivotTable.calculated_items`. Placing a calculated field on the row, column
  or page axis raises `ValueError`, as Excel only allows calculated fields as
  data fields.

## 2.0.20 - 2026-10-26

### Fixed

- Pivot tables created with `wolfxl.pivot` now open in Microsoft Excel
  without a repair prompt, in both write mode and modify mode (#63). Before,
  every such workbook failed to open because `<pivotCaches>` was written
  before `<calcPr>` in `workbook.xml`, which the workbook schema forbids.
- The pivot range now shows its labels, values, subtotals and grand totals
  when the file is first opened, laid out as Excel's compact form lays them
  out, so a refresh in Excel leaves the range unchanged. Before, the cells were
  blank until the user clicked Refresh. This also covers nested row and column
  fields, two or more data fields (Excel's "Values" field), blank source
  values, and page fields, including a page field set to one item.
  Data fields shown as a percentage of the total, row or column, or as an
  index, render the values Excel shows. Running totals, differences and other
  options measured against a base field or item are left blank and the cache
  is marked to refresh on load, so Excel fills them in on refresh. Rendering
  replaces anything already in the pivot range, so earlier cell values cannot
  pass for pivot results.
- `location` is now the anchor of the pivot: WolfXL sizes the range from its
  top-left cell, and an explicit range such as `"F2:G9"` is resized to fit.
  Page fields sit above the table with one spacer row, as Excel places them;
  when there is no room above the anchor, the table moves down.
- Placing a field on the row, column or page axis now raises `ValueError`
  when its values are too many to enumerate as shared items, instead of
  writing a pivot that Excel cannot open. Such a field can still be a data
  field.

## 2.0.19 - 2026-10-26

### Documentation

- Lead the README and PyPI page with the jobs WolfXL is for: speeding up
  openpyxl code with one import change and editing styled templates without
  losing conditional formatting or data validation. Add a table that also
  names the read-only and write-only libraries suited to other jobs. No
  library code changed.

## 2.0.18 - 2026-10-26

### Fixed

- Save `date` and `datetime` values written in modify mode
  (`load_workbook(path, modify=True)`). 2.0.17 raised
  `ValueError: Unknown cell type: 'datetime'` on save. Dates are now stored as
  Excel serials in the workbook's epoch (1900 or 1904) with a `yyyy-mm-dd` or
  `yyyy-mm-dd h:mm:ss` number format, matching openpyxl, and a
  timezone-aware `datetime` raises the same `TypeError` as openpyxl.
- Add new number formats in modify mode to workbooks whose styles contain an
  empty `<numFmts count="0"/>` collection, as openpyxl writes them. Before,
  setting two or more new number formats in one save left the cells pointing
  at number formats that were never written, so openpyxl failed to read the
  saved file.

## 2.0.17 - 2026-10-26

### Fixed

- Restore the speed of `iter_rows()` without bounds in normal loads of large
  sheets, which 2.0.16 slowed down when it stopped trusting a sheet's
  `<dimension>`. WolfXL now finds the last cell and merged range with a Rust
  scan of the sheet XML instead of parsing the whole sheet in Python before the
  first row. On a 200,000-row by 8-column sheet (Linux x86_64, Python 3.12,
  median of 7 runs), the first row arrives in 0.28 s instead of 2.84 s and a
  full `iter_rows(values_only=True)` pass takes 0.60 s instead of 3.16 s. Rows
  and values are unchanged, and explicit `max_row` and `max_col` bounds still
  skip the scan.

## 2.0.16 - 2026-10-25

### Fixed

- Return the same cells as openpyxl from streamed reads: every cell from
  `load_workbook` (normal and `read_only`, path or file object) has
  `.hyperlink`, formula cells loaded with `data_only=True` report the data type
  of their cached value (`n`, `s`, `b` or `e`), and a formula whose cached
  string is empty reads as `None` with data type `str`.
- Yield the same rows as openpyxl when a sheet's `<dimension>` or empty row
  records reach past its last cell: `iter_rows()` without bounds now stops where
  openpyxl stops in normal and `read_only` loads, while explicit `max_row` and
  `max_col` bounds still pad with empty cells.
- Open workbooks whose `workbook.xml` lists a sheet with an empty relationship
  id (`<sheet r:id="">`), skipping that entry as openpyxl does instead of
  failing with `missing workbook part`. A sheet whose relationship id is set
  but missing from the workbook relationships still raises.
- Raise `MemoryError` when memory runs out while streaming rows, for example on
  a hostile shared-strings table. This previously raised `PanicException`,
  which derives from `BaseException` and escaped `except Exception` handlers.
- Save in place again after an in-place save that added an image: editing
  cells, adding an image or removing one no longer fails with a Rust panic, and
  the image added by the earlier save can be removed.
- Rename a worksheet to a name the workbook already uses the way openpyxl does:
  assigning `ws.title` an existing name (matched case-insensitively) picks the
  next free name (`Data` becomes `Data1`) instead of raising `ValueError`, and
  renaming a sheet to its own title stays a no-op. pandas'
  `ExcelWriter(path, engine="openpyxl", mode="a", if_sheet_exists="new")` now
  works after `wolfxl.install_as_openpyxl()`.
- Support Docling's Excel converter after `wolfxl.install_as_openpyxl()`:
  `openpyxl.reader.excel.ExcelReader` rewinds file-like inputs and supports the
  staged `read_workbook()` and `read_worksheets()` calls, exposing stored cells
  without materializing the sheet's declared dimension, so Docling extracts the
  same cell values as with openpyxl.

### Changed

- Import no openpyxl module during `import wolfxl` when openpyxl is installed.
  The compatibility shims read openpyxl's tag names and attributes on first
  use instead.
- Keep one copy of the input for `read_only` loads from bytes or file objects,
  and parse `sharedStrings.xml` only when a non-streaming API needs it.
- Inflate streamed sheet XML from the ZIP entry while parsing instead of
  spooling sheets over 1 MiB to a temporary file. ZIP safety checks and size and
  ratio limits still apply.
- Convert date, time and duration cells in Rust for streamed
  `iter_rows(values_only=True)` instead of in a per-cell Python loop, with the
  same values and types as openpyxl for the 1900 and 1904 date systems;
  out-of-range serials stay numeric.

## 2.0.15 - 2026-10-25

### Fixed

- Return the same values as openpyxl from normal (`read_only=False`) loads:
  numeric cell text with a decimal point or exponent (`100.0`, `1E2`) reads as
  `float` even when the value is whole, and line breaks stored as raw CR or CR
  LF in shared strings, inline strings and rich-text runs read as `\n`, while a
  `&#13;` reference keeps its carriage return.
- Attach a hyperlink on a cell range (`<hyperlink ref="A3:A5">`) to every
  covered cell on load, so `cell.hyperlink.target` and `cell.hyperlink.ref`
  match openpyxl for each cell instead of returning `None`. Open-ended ranges
  and ranges over 65,536 cells are limited to the sheet's used area, and the
  range hyperlinks on one sheet expand to at most 65,536 cells in total; a
  range that does not fit is attached to its first cell only, with a warning.
- Skip worksheet images whose data is unsupported or unreadable (for example
  WMF) with a format-specific warning, as openpyxl does, instead of failing
  when `ws._images` is read. Modify-mode saves keep the omitted media, and
  removing or replacing later images still targets the right drawing entries,
  including across repeated in-place saves. Invalid images passed to
  `Image(...)` still raise.
- Support pandas' `ExcelWriter(path, engine="openpyxl", mode="a")` after
  `wolfxl.install_as_openpyxl()`: it now adds, replaces
  (`if_sheet_exists="replace"`) and overlays (`"overlay"`) sheets in an
  existing workbook as openpyxl does. 2.0.14 raised `RuntimeError: save
  requires write or modify mode` on pandas 2 and 3.

## 2.0.14 - 2026-10-25

### Fixed

- Convert date, datetime, time and duration input cells to Excel serial values
  during `calculate()`, respecting the workbook's 1900 or 1904 date system.
- Apply the same input conversion to single-cell references and ranges without
  changing the values stored in the workbook.
- Preserve saved formula caches after date and time calculations while keeping
  formula text, input values and number formats unchanged in normal and modify
  modes.
- Skip an invalid `<row r="0">` in `read_only=True` loads, as openpyxl and normal loads do. It was returned as an extra first row.
- Keep pandas' `openpyxl` engine working after `wolfxl.install_as_openpyxl()`.
  pandas read the WolfXL version from the aliased module and refused it as an
  openpyxl older than 3.1; its version check now sees openpyxl 3.1.5, the API
  release the alias provides. `wolfxl.__version__` is unchanged.
- Accept pandas 3 restoring `_max_row` and `_max_column` on a `read_only=True`
  worksheet after `reset_dimensions()`, as openpyxl's read-only worksheet does.
  `read_excel` raised `AttributeError` after reading each sheet.

### Compatibility

- The built-in formula catalog is unchanged. `TIME` remains available through
  the existing optional `calc` extra, not through the default installation.

## 2.0.13 - 2026-10-24

### Fixed

- Read rows and cells that omit their `r` reference at the position Excel
  gives them: such a row follows the previous row, and such a cell follows the
  previous cell in its row. Normal loads placed these cells at the wrong
  coordinates, and `read_only=True` loads from a path raised `OSError`.
- Return only the base text of shared strings that carry phonetic (`rPh`)
  runs, as openpyxl does. The phonetic reading was appended to the cell value.
- Return `datetime.time` for cells with a time-only number format in normal
  loads, including serial 0 (midnight), and keep millisecond precision in
  dates and times, as openpyxl and `read_only=True` loads already did.
- Stream `read_only=True` workbooks opened from file objects as well as paths,
  and cap each streamed chunk at 65,536 cells, so wide sheets read from a file
  object no longer load whole into memory.
- Return the same values as openpyxl from `read_only=True` loads: cached
  formula results keep their stored type (booleans as `bool`, strings as `str`
  even when they look numeric), line breaks stored as raw CR or CR LF read as
  `\n`, empty shared-string and value elements read as `None`, and inline
  strings ignore a stray `<v>` element.

## 2.0.12 - 2026-10-24

### Fixed

- Restore Python 3.9 calculation and worksheet iteration by omitting the
  unsupported dataclass `slots` argument on Python 3.9. Python 3.10 and later
  retain frozen, slotted merged-range records. The failure also affected 2.0.10
  and 2.0.11.
- Restore Python 3.9 drawing text imports and keyword construction of bullet
  styles by omitting the unsupported dataclass `kw_only` argument and using a
  compatible runtime type union. Repair the same runtime union incompatibility
  in worksheet cell watches. Python 3.10 and later retain the same
  constructor signatures.
- Import every installed package module and check bullet construction in
  clean-installed CPython 3.9 wheels on every supported platform:
  calculate the README result 2450, save and reload cached values, and preserve
  formula text, edited inputs, formatting and merged headings in normal and
  modify modes.

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
- Keep shared-string reference parsing compatible with the XML dependency
  update in the Rust round-trip regression fixture.
- Preserve existing-cell XML byte splicing and its regression coverage in
  modify mode while updating the XML dependency.

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
