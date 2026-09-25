# Changelog

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
