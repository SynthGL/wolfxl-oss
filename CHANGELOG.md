# Changelog

## [Unreleased]

### Added

- A reproducible one-cell workbook editing study in `study/agent-edit/`, with
  public fixtures, verified runtime downloads, offline regression tests, and
  separate Community plus LibreOffice and Commercial native calculation workflows.

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
