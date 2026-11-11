# Changelog

## [Unreleased]

### Fixed

- Modify-mode saves (`load_workbook(..., modify=True)`) that add cells or
  restyle existing cells no longer read the whole sheet a second time. 2.0.37
  added that pass to keep new cells in column order; the saver now reads ahead
  only the rows being changed, with the same output (#121).
- The worksheet patcher no longer writes malformed XML when it changes a cell
  value in a sheet part that starts with a UTF-8 byte-order mark (#121).

## 2.0.37 - 2026-11-11

### Added

- `calc` and `peek --add` gain 26 native-Excel-backed math and information
  functions (133 in total): `ACOS`, `ASIN`, `ATAN`, `ATAN2`, `COS`, `SIN`,
  `TAN`, `DEGREES`, `RADIANS`, `PI`, `EXP`, `LN`, `LOG`, `LOG10`,
  `CEILING`, `FLOOR`, `TRUNC`, `PRODUCT`, `RAND`, `RANDBETWEEN`, `ISBLANK`,
  `ISERROR`, `ISNA`, `ISNUMBER`, `ISTEXT`, and `NA`. Information predicates
  preserve array shape. Function discovery JSON includes volatility metadata;
  fingerprint and verify continue using stored workbook values.
- Generate a text-redacted diagnostic ZIP with `python -m wolfxl.report FILE`
  and attach it using the new "Workbook fails to load" issue form. Reports include
  bounded package inventories, part sizes, standard element counts, XML skeletons,
  sanitized stack traces and versions, without cell values, shared strings,
  formulas, attribute values or binary contents. Review structure and sizes before
  publishing. Existing output files are never overwritten.
- `wolfxl-cli` 0.10.0 adds analyst commands backed by `wolfxl-core` 0.9.0
  and the new MIT `wolfxl-reader`, `wolfxl-formula`, and `wolfxl-rels` crates.
  `wolfxl-cli` 0.11.0 adds the full formula function set, the `wolfxl mcp`
  server with workbook editing, and prebuilt binaries. `wolfxl-cli` 0.11.1
  adds table references, 40 more functions, and header-row selection for
  sheets with a title above the table.
- `calc`, `peek --add`, and the MCP `calc` tool evaluate Excel structured
  references: `Table[Column]`, `Table[]`, `[#All]`, `[#Data]`, `[#Headers]`,
  and `[#Totals]` and their combinations, and column ranges such as
  `Table[[#Data],[Q1]:[Q4]]`, with escaped column names. They resolve
  against the workbook's tables from any sheet. Unknown tables or columns
  evaluate to `#REF!`; `[@Column]` and `[#This Row]` evaluate to `#VALUE!`.
- `calc` and `peek --add` gain 40 functions (107 in total): `MEDIAN`,
  `MODE`/`MODE.SNGL`, `STDEV`/`STDEV.S`/`STDEV.P`/`STDEVP`,
  `VAR`/`VAR.S`/`VAR.P`/`VARP`, `PERCENTILE`/`.INC`/`.EXC`,
  `QUARTILE`/`.INC`/`.EXC`, `RANK`/`RANK.EQ`/`RANK.AVG`, `LARGE`, `SMALL`,
  `COUNTBLANK`, `SUMPRODUCT`, `SUBTOTAL`, `IFS`, `SWITCH`, `IFNA`, `CONCAT`,
  `TEXTJOIN`, `VALUE`, `NETWORKDAYS`, `WORKDAY`, `XNPV`, `XIRR`, `RATE`,
  `NPER`, `UNIQUE`, `FILTER`, and `SORT`.
- `SUBTOTAL` skips cells that themselves hold `SUBTOTAL` formulas, such as a
  table's totals row, and function numbers 101 to 111 skip hidden rows in
  `.xlsx`, `.xlsm`, and `.xlsb` workbooks.
- Operators work element by element over ranges and arrays
  (`=B2:B10*C2:C10`, `=SUMPRODUCT((A2:A10="North")*C2:C10)`). A single row
  or column repeats to match the other operand, and positions outside both
  shapes are `#N/A`.
- `peek`, `stats`, and `pivot` take `--header-row N` for sheets whose header
  is not row 1, and accept sheet column letters wherever they accept a
  header name. A header name shared by several columns exits 2 naming the
  cells and asking for a letter.
- `peek --addresses` labels box, text, and Markdown previews with sheet
  column letters and row numbers. The interactive viewer shows the cursor's
  source cell address through sorting, and keeps sort notices whole on
  narrow terminals.
- Numeric filters and formula imports accept trailing percentages (`12%` is
  0.12); `peek` still prints the source spelling.
- `wolfxl peek --where COLUMN<op>VALUE` keeps the data rows that satisfy every
  condition (`=`, `!=`, `<`, `>`, `<=`, `>=`, `~` contains), comparing money
  spellings such as `$1,200.00` and `(500.00)` numerically. `--select A,B`
  prints the named columns in the order given.
- `wolfxl peek --interactive` (`-i`) browses the sheet in a full-screen
  terminal grid with a frozen header row, search, column sort, value
  frequencies, and sheet switching.
- `wolfxl peek --max-columns N` bounds the columns in JSON output, alongside
  `--max-rows`.
- `wolfxl pivot` groups rows by header fields and aggregates value fields with
  `sum`, `count`, `avg`, `min`, and `max`, with an optional `--column` field
  and grand totals, as a box table, CSV, or JSON.
- `wolfxl stats` profiles each column: blank and non-empty counts, counts by
  type, distinct values, a numeric summary, and the most frequent values, with
  the same `--where` and `--select` options as `peek`.
- `wolfxl diff LEFT RIGHT --key K` reconciles two tables row by row on a key
  column, with numeric tolerance, key normalization, JSON or text reports, and
  exit codes for scripted checks.
- `wolfxl fingerprint` prints a style-independent content fingerprint
  (`wxf1:<sha256 hex>`), and `wolfxl verify --fingerprint F` checks a file
  against one.
- `wolfxl calc FORMULA [FILE|-]` evaluates one formula against a sheet's
  stored values, using cached results for formula cells, and prints the
  result as text, CSV rows for an array, or JSON. It exits 0 for a value, 1
  for an Excel error value, and 2 for usage, input, parse, or
  unsupported-function failures.
- A built-in formula evaluator behind `calc` and `peek --add` covers
  arithmetic, comparison, `&` concatenation, cell and sheet-qualified
  references, and ranges. A formula naming a function it does not implement
  is rejected with a pointer to `wolfxl list-functions`.
- Formula references import delimited money spellings such as `$1,200.00`
  and `(75.00)` as numbers, using the same parser as `--where`. Arithmetic
  text coercion accepts those spellings too. Stored Excel text cells stay
  text, and range aggregations continue to ignore them.
- `COUNT` skips error values in references and direct arguments, while the
  other aggregate functions propagate them. Formula concatenation formats
  numbers with the same 15-significant-digit rounding as scalar output.
- Formula references and fingerprints read raw stored numbers in XLS, XLSB,
  and ODS books, preserving fractional serials and the stored date system
  independently of number formats.
- `wolfxl peek --add NAME=FORMULA` appends formula columns filled down from
  the first data row with `$` anchoring, computed before `--where` and
  `--select` and available in `--interactive`. A name that collides with a
  header or an earlier `--add`, or a formula that fails to parse, exits 2
  naming the argument.
- `wolfxl list-functions` lists the evaluator's functions by name and
  category, with `--category` and `--search` filters, and `--check NAME,...`
  reports each name as supported or unsupported, exiting 1 when any is
  unsupported.
- The built-in evaluator behind `calc` and `peek --add` now implements the
  same 67 functions as the Python formula evaluator, in the `math`,
  `statistical`, `logical`, `lookup`, `text`, `date`, and `financial`
  categories, including `SUMIFS`, `COUNTIFS`, `XLOOKUP`, `INDEX`/`MATCH`,
  `OFFSET`, `TEXT`, `EOMONTH`, `NPV`, `IRR`, and `PMT`. Formulas accept
  array constants such as `{1,2;3,4}` and empty arguments such as
  `XLOOKUP(x,A:A,B:B,,1)`, report wrong argument counts as parse failures,
  and date functions follow the workbook's 1904 date system.
- `peek`, `pivot`, `stats`, `diff`, and `calc` read standard input when the
  file is `-`, with `--from csv|tsv|xlsx|xlsm|xls|xlsb|ods` naming the format;
  `calc` also reads it when no file is given.
- `wolfxl mcp` runs a stdio MCP server for AI agents. Its tools read
  workbooks (`get_workbook`, `read_range`, `calc`, `stats`, `pivot`, `diff`,
  `fingerprint`) and edit `.xlsx`/`.xlsm` files in place (`write_cells`,
  `set_number_format`, `insert_rows`, `delete_rows`, `insert_columns`,
  `delete_columns`, `add_sheet`, `rename_sheet`, `undo`, `redo`). Mutating
  tools accept `dry_run` and `expected_revision`; `--root DIR` confines paths
  and `--read-only` hides the editing tools.
- New `wolfxl-edit` crate: byte-preserving edit sessions for existing
  workbooks. Parts an edit does not touch keep their exact bytes; formula
  writes, and any write to a workbook that holds formulas, mark the workbook
  for recalculation on open. Renaming a sheet also renames sheet-qualified
  defined names such as `Data!Rate`. The Python package's
  modify mode now shares its sheet patcher.
- Prebuilt `wolfxl` binaries for Linux (x86_64, arm64, static musl), macOS
  (x86_64, arm64), and Windows (x86_64) on GitHub Releases under
  `wolfxl-cli-v*` tags, with checksummed `install.sh` and `install.ps1`
  installers and `cargo binstall` metadata.

### Changed

- Bumped `wolfxl-cli` to `0.11.1`, then `0.11.2`, and `wolfxl-edit` to `0.1.1`.
- Updated the CLI to use `wolfxl-core` 0.9.0 and the MIT reader crates.
- Preserve exact CSV/TSV header names when resolving pivot fields, matching
  the existing filters, selectors, keyed diffs, and column profiles.
- Keep stored formula input cells sparse, including sheets with widely
  separated cells, rather than allocating their full bounding rectangle.
- Stream stdin and extension-overriding `--from` copies into atomically
  owner-only temporary files, with named 512 MiB XLS and 4 GiB other-format
  spool limits reported as exit-2 input failures.
- Sanitize and display-width-bound interactive status metadata and load
  errors while keeping cursor coordinates visible on narrow terminals.

### Fixed

- Cells added to an existing row in modify mode
  (`load_workbook(..., modify=True)`) are saved in column order. Writing `B1`
  into a row that held `A1` and `C1` used to save the cells as `A1`, `C1`,
  `B1`; Excel reported the workbook as damaged and its repair discarded the
  new value (#114, #115). Cells that an earlier version already saved out of
  order are updated in place and are not reordered.
- `RATE`, `IRR`, and `XIRR` converge on the rate itself rather than on a
  near-zero balance, so scaling every cash flow by a constant keeps the rate
  (`RATE(1,0,1E-12,-2E-12)` is 1) and balances that only approach zero are
  `#NUM!` (`RATE(360,0,100)`). When Newton's method does not settle, the
  fallback searches outward from the guess, so a column of cash flows with no
  rate is `#NUM!` after a few hundred passes over the flows.
- `wolfxl-formula` 0.1.1 tokenizes escaped brackets in structured references
  (`Sales[Cost '[USD']]`) as one reference, so `wolfxl peek --add` accepts them.
- `wolfxl pivot` aggregates money-spelled text (`$12,500.00`, `(500.00)`)
  as numbers, the same reading `stats` and `--where` use, instead of leaving
  `sum`, `avg`, `min`, and `max` blank.
- `calc` and `peek --add` read ISO date fields of delimited input
  (`2008-01-01`) as date serials, as Excel's text import does, so
  `XIRR(B2:B6,A2:A6)` works on a CSV of dated cash flows.

### Removed

- Drop Python 3.9 and 3.10 support; Python 3.11 is now the minimum supported
  version. Version 2.0.36 is the last release supporting Python 3.9 and 3.10;
  pip on those versions continues to resolve compatible 2.0.x releases.

## 2.0.36 - 2026-10-31

### Added

- `Limits` and `LimitExceeded` let XLSX callers bound uncompressed package and
  part bytes, stored cells, shared-string bytes, and materialized text bytes
  with `load_workbook(..., limits=...)`.

### Changed

- XLSX loads now apply default workbook-wide budgets in every mode: 64 million
  stored cell elements, 256 MiB of shared text, and 512 MiB of decoded text and
  formulas. Package and part defaults remain 4 GiB and 512 MiB respectively.
  Oversized packages and parts raise `LimitExceeded` (a `ValueError`) rather
  than `OSError`. Environment ZIP-size overrides do not raise these new
  per-load budgets; pass an explicit `Limits` as well. Existing ZIP admission
  guards remain in force.
- After format detection, path-backed XLSX loads open the file once for
  admission, loading and later reads, never reopening the original path. On
  Linux and macOS the source is a read-only memory map where the filesystem
  supports it; otherwise, and on Windows, the compressed package is read into
  memory. Feature readers use the same source; modify mode uses a private copy.
  Saves back to a mapped source path write a new file and rename it over the
  path instead of rewriting it in place. On Linux and macOS, writing to or
  truncating that loaded file in place changes what the open workbook reads
  and can terminate the process with `SIGBUS`.
- Closing releases a mapped source unless a live reader still uses it; the
  mapping is then released when the last reader is collected.

### Fixed

- Bound memory for shared strings referenced by many cells. Eager and
  streaming readers reuse shared text and Python string objects per workbook.
- Admit and load one opened file for path-backed XLSX inputs, keeping
  streaming strings bound to that source after the original path is replaced
  by rename. Compressed path sources are also bounded by `max_package_bytes`.
- Apply worksheet cell and text budgets when package relationships select a
  nonstandard workbook part, before normalizing that part for loading.

### Security

- Untrusted XLSX files can no longer exhaust memory through a shared string
  referenced by many cells. In 2.0.35, a 110 KB workbook that references one
  1 MiB shared string from 12,000 cells used more than 2 GiB on a full load.
  Upgrade to 2.0.36. A GitHub security advisory with the affected version
  range will follow.
- Update anyhow to 1.0.104 (RUSTSEC-2026-0190).
- Require pytest 9.0.3 or newer in the test extra on Python 3.10 and newer,
  while retaining the compatible pytest range on Python 3.9.

## 2.0.35 - 2026-10-30

### Changed

- Saving a new workbook that holds a timezone-aware `datetime` or
  `datetime.time` now raises the same `TypeError` as openpyxl, from `save()`
  for `append()`, `write_rows()` and cell assignments, and from `append()` on
  a write-only sheet. Previously these values saved without their timezone.
  Modified workbooks already raised this error. Unlike openpyxl, WolfXL leaves
  no partial file behind.

### Fixed

- A new workbook no longer drops a date, datetime, rich text or other
  non-plain value assigned to a cell when the sheet holds the same number of
  plain values assigned with `ws["A1"] = ...` or `ws.cell(...)`. For example,
  `ws["A1"] = 5` followed by `ws["A2"] = date(2024, 1, 2)` saved only `A1`.
  This affected every release since 2.0.1.

## 2.0.34 - 2026-10-30

### Fixed

- A `datetime.time` written with `append()`, `write_rows()`, a cell
  assignment, a write-only sheet or a modified workbook now saves as its
  fraction of a day with the `h:mm:ss` format, as openpyxl saves it, and reads
  back as a `datetime.time`. Previously it saved as text such as `10:30:00`.

## 2.0.33 - 2026-10-30

### Fixed

- A `t="d"` cell now reads as the date, time, datetime or timedelta its ISO
  8601 text names, as openpyxl reads it, instead of as the raw text. For
  example, `2024-01-15` reads as `datetime.date(2024, 1, 15)` and `PT1H30M` as
  `datetime.timedelta(hours=1, minutes=30)`. The cell reports
  `data_type == "d"` in every load mode, and its number format no longer
  changes the value. Text that names no valid date or time stays a string;
  openpyxl refuses to load such a workbook.
- With a full load, `iter_rows(values_only=True)` and `iter_cell_records()`
  now read a serial under a duration format such as `[h]:mm:ss` as a
  `datetime.timedelta`, as `Cell.value` and the streaming reader already did.
- A full load no longer raises `PanicException` for a date- or
  duration-formatted serial outside Python's date range. The value stays a
  number with `data_type == "n"`, as it does on streamed loads.
- On streamed sheets, a cell that reads as a `datetime.timedelta` now reports
  `data_type == "d"` instead of `"n"`.
- A `datetime.timedelta` written with `append()`, `write_rows()`, a cell
  assignment or a write-only sheet now saves as its serial day count with the
  `[hh]:mm:ss` format, as openpyxl saves it. Previously it saved as text such
  as `1 day, 3:30:00`, so copying rows read with `values_only=True` kept their
  durations only as text.

## 2.0.32 - 2026-10-30

### Fixed

These fixes apply to workbooks loaded without streaming, as 2.0.31 did for
streamed sheets:

- A `t="b"`, `t="e"` or `t="d"` cell with a missing or empty `<v>` now reads
  as `None`, as openpyxl does, instead of `False` or `""`. An empty `t="d"`
  cell reports `data_type == "d"` instead of `"s"`.
- With `data_only=True`, a formula cell with no cached result keeps its `t=`
  type, such as `"b"`, `"d"`, `"e"` or `"s"`, and reads as `None`.

## 2.0.31 - 2026-10-30

### Fixed

These fixes apply to streamed sheets, including `read_only=True`:

- A self-closing typed cell such as `<c t="s"/>`, `<c t="str"/>`,
  `<c t="b"/>`, `<c t="e"/>` or `<c t="d"/>` now keeps its `data_type`, as
  openpyxl does, instead of reporting `"n"`.
- An error cell with no code now reads as `None` instead of `"#ERROR!"`.
- A formula cell whose cached result is an error now reads as its formula
  when `data_only=False`, instead of as the error code.
- With `data_only=True`, a formula cell with no cached result keeps its
  `t=` type, such as `"b"`, `"d"`, `"e"` or `"s"`, instead of `"n"`.

## 2.0.30 - 2026-10-30

### Fixed

- A `t="inlineStr"` cell without an `<is>` child now reports
  `data_type == "inlineStr"`, as openpyxl does, instead of `"s"` when the
  sheet is loaded eagerly or `"n"` when it is streamed or opened with
  `read_only=True`. This covers self-closing cells, cells with only a `<v>`,
  and, with `data_only=True`, formula cells. Cells with an `<is>` child still
  report `"s"`.

## 2.0.29 - 2026-10-30

### Fixed

- A worksheet without a `<dimension>` element and with up to 50,000 rows is
  read eagerly again outside `read_only=True`, as documented. The row count
  that decides auto-streaming counted some `<row>` tags twice, so sheets just
  under the threshold were streamed and missed the faster eager parser.

### Performance

- That row count now runs natively and stops at the first row past the
  threshold. Default-mode `load_workbook()` followed by
  `iter_rows(values_only=True)` on 50,000-row sheets without a dimension
  drops from 0.197 s to 0.167 s (mixed types) and from 0.297 s to 0.245 s
  (dates) in a local interleaved comparison.

## 2.0.28 - 2026-10-29

### Performance

- Loading a workbook without `read_only=True` parses each worksheet faster:
  sheets with no `<autoFilter>` skip a second XML pass, cell attributes are
  read in one pass, worksheet events borrow the in-memory XML instead of
  copying it, and the cell list is sized once from a byte count of `<c>` tags.
  Default-mode `load_workbook()` followed by `iter_rows(values_only=True)`
  over every sheet drops from 1.11 s to 0.69 s on the 4-sheet,
  200,000-row formula workbook and from 0.20 s to 0.13 s on a 50,000-row,
  8-column plain sheet.

## 2.0.27 - 2026-10-29

### Performance

- Read-only `iter_rows(values_only=True)` and `iter_value_chunks()` read
  ordinary worksheet rows, including formula cells and plain inline strings,
  with a byte-level row scanner and hand any row it does not recognize to the
  XML event reader, so output is unchanged. Sheets of one workbook now share a
  single parsed shared-string table instead of parsing it per sheet, and the
  ZIP reader no longer scans the whole file to locate the central directory.
  Reading the formulas of a 4-sheet, 200,000-row workbook drops from 0.61 s
  to 0.30 s.
- `iter_value_chunks()` with `data_only=True` no longer loads sheets of up to
  50,000 rows into memory before chunking them; every sheet streams, which is
  faster on all measured shapes. The 4-sheet formula workbook reads in 0.30 s
  instead of 0.49 s.

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
