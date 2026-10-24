# Sprint Κ Pod-α Smoke Fixtures

This folder holds **placeholder** fixtures that exercise the `.xlsx` /
`.xlsb` / `.xls` / `.ods` read paths.  They are deliberately minimal —
just enough to guarantee that:

1. `_rust.classify_file_format(path)` returns the right format string.
2. `NativeXlsbBook.open` / `NativeXlsBook.open` (and their
   `open_from_bytes` siblings) accept real, parseable input.
3. Legacy `.xls` strict-raise style accessors fire `NotImplementedError`.

Curated parity fixtures live under `tests/parity/fixtures/{xls,xlsb}/`.
This directory remains the minimal format-dispatch smoke corpus.

For independently authored Community XLS/ODS scalar expectations, use
`crates/wolfxl-reader/qualification/fixtures.*` with `expected.json`; produce
the fixtures with the pinned xlwt/odfpy generator in that directory.

## Files

| Path                       | Format | Source                                         |
|----------------------------|--------|------------------------------------------------|
| `sprint_kappa_smoke.xlsx`  | xlsx   | Generated locally via `openpyxl` (4-row sheet) |
| `sprint_kappa_smoke.xlsb`  | xlsb   | Borrowed from `wolfxl-core` 0.8 sample-date    |
| `sprint_kappa_smoke.xls`   | xls    | LibreOffice headless `--convert-to xls`        |
| `sprint_kappa_smoke.ods`   | ods    | LibreOffice headless `--convert-to ods`        |

## Regenerating

LibreOffice's bundled xlsb export filter is broken at the time of
writing (`libreoffice 25` returns `Error: 0x81a` on the official
filter).  We therefore **borrow** an existing valid xlsb from the
`wolfxl-core` published crate's tests/fixtures.  Replace it with
anything legitimate: the Pod-α tests only check that the file
classifies correctly and opens through the native `.xlsb` reader.

```bash
# xlsx (with openpyxl)
python3 -c "
from openpyxl import Workbook
wb = Workbook()
ws = wb.active
ws['A1'] = 'name'; ws['B1'] = 'value'
ws['A2'] = 'alpha'; ws['B2'] = 1
ws['A3'] = 'beta';  ws['B3'] = 2
ws['C2'] = '=B2*2'; ws['C3'] = '=B3*2'
wb.save('sprint_kappa_smoke.xlsx')"

# xls + ods (LibreOffice)
soffice --headless --convert-to xls sprint_kappa_smoke.xlsx
soffice --headless --convert-to ods sprint_kappa_smoke.xlsx

# xlsb (manual: use Microsoft Excel or borrow from upstream)
cp ~/.cargo/registry/src/index.crates.io-*/wolfxl-core-*/tests/fixtures/sample-date.xlsb \
   sprint_kappa_smoke.xlsb
```
