# WolfXL Community

**Edit existing Excel files from Python without losing formatting, 7-14x faster than openpyxl on most reads and writes, with an openpyxl-compatible API. MIT licensed.**

WolfXL Community reads, writes, and edits Excel `.xlsx` and `.xlsm` workbooks
through the openpyxl API, with parsing, serialization, and cell storage
implemented in Rust. Modify mode saves the cells you change and keeps the rest
of the file, and `calculate()` covers common Excel functions in Python.

WolfXL Community is the maintained, MIT-licensed 2.0 release line for supported workbook creation, reading, writing, streaming exports, and existing-workbook edits. It is a free product, not a trial.

## Install

```bash
python -m pip install wolfxl==2.0.6
```

## Edit an existing workbook

```python
from wolfxl import load_workbook

wb = load_workbook("report.xlsx", modify=True)  # edit the existing file
wb["Summary"]["B2"] = 1500
print(wb.calculate()["Summary!B4"])  # 2450.0 from =SUM(B2:B3)
wb.save("report-updated.xlsx")  # cells you did not touch keep their formatting
wb.close()
```

Modify mode saves the cells you change and preserves unchanged styles, charts,
and package parts within the documented boundaries; add `keep_vba=True` for
`.xlsm` macros. `calculate()` returns the computed values and leaves the cached
results in the saved file unchanged.
[Edit Excel in Python without losing formatting](https://wolfxl.com/openpyxl-preservation?utm_source=pypi&utm_medium=registry&utm_campaign=community_commercial_2026_09).

## Compared with openpyxl

- **Same API for the covered surface.** Most supported openpyxl-shaped code
  starts with one import change:

  ```diff
  - from openpyxl import Workbook, load_workbook
  + from wolfxl import Workbook, load_workbook
  ```

  openpyxl implements more of its own API, so review the
  [compatibility matrix](https://wolfxl.com/docs/migration/compatibility-matrix/?utm_source=pypi&utm_medium=registry&utm_campaign=community_commercial_2026_09),
  the [known limitations](https://wolfxl.com/docs/trust/limitations/?utm_source=pypi&utm_medium=registry&utm_campaign=community_commercial_2026_09),
  and the [openpyxl migration guide](https://wolfxl.com/openpyxl-migration?utm_source=pypi&utm_medium=registry&utm_campaign=community_commercial_2026_09),
  then prove one representative workbook before replacing a production path.
- **Existing templates.** On a sheet with an extension data validation and a
  sparkline, one cell edit saved by openpyxl 3.1.5 removed both;
  `load_workbook(path, modify=True)` in WolfXL 2.0.2 kept both.
- **Formulas.** openpyxl stores formula text and never computes it; with
  `data_only=True` it returns the value Excel last cached, or `None` for a
  file Excel never opened. WolfXL Community `calculate()` covers common
  functions only, in Python. Need results that match Excel? Commercial
  includes a native engine verified on 704 Excel-calculated cases
  ([pricing](https://wolfxl.com/pricing?utm_source=pypi&utm_medium=registry&utm_campaign=community_commercial_2026_09)).
- **Large files.** The
  [large-file receipts](https://wolfxl.com/openpyxl-large-files?utm_source=pypi&utm_medium=registry&utm_campaign=community_commercial_2026_09)
  record seconds and peak memory for 1.6-million-cell reads, writes, and edits.
- **Where openpyxl fits.** openpyxl is pure Python and installs anywhere
  Python runs. WolfXL needs a published wheel for your platform or a Rust
  toolchain to build from source.

## Compare openpyxl alternatives, measured

WolfXL Community is benchmarked against the maintained open-source Python Excel
libraries on a committed run. On 2026-08-18 (AMD EPYC 9655, Python 3.13.15,
median of 5 rounds, 200,000-row x 8-column workbook), Community 2.0.1 wrote 1.6
million cells in 0.73 s (11.10x openpyxl 3.1.5, 6.46x XlsxWriter 3.2.9, 5.01x
PyExcelerate 0.13.0) and read them back in 0.39 s (11.85x openpyxl, 1.49x
python-calamine 0.8.2) at a 153 MiB read peak, below openpyxl (169 MiB) and
python-calamine (286 MiB). Raw results and the harness are published
under [`benchmarks/`](https://github.com/SynthGL/wolfxl-oss/tree/main/benchmarks).

Full tables, capability scope, and reproduction commands:
[openpyxl alternatives, measured](https://wolfxl.com/openpyxl-alternatives?utm_source=pypi&utm_medium=registry&utm_campaign=community_commercial_2026_09).

## Choose the right edition

Community is the free option when supported workbook I/O is sufficient.
Commercial 2.1+ is a separate option for the following workflow requirements:

| Requirement | Community 2.0 line | Commercial 2.1+ |
| --- | --- | --- |
| Supported workbook I/O | Included within the documented surface | Included |
| Native formula recalculation | Not included | [Calculate a supported formula set](https://wolfxl.com/calculate-excel-formulas-python?utm_source=pypi&utm_medium=registry&utm_campaign=community_commercial_2026_09) |
| PDF or image rendering | Not included | [Render supported output](https://wolfxl.com/render-excel-python?utm_source=pypi&utm_medium=registry&utm_campaign=community_commercial_2026_09) |
| Format conversion | Not included | [Scope an evaluation](https://wolfxl.com/pilot?utm_source=pypi&utm_medium=registry&utm_campaign=community_commercial_2026_09) |
| VBA or Power Query operations | Not included | [Scope an evaluation](https://wolfxl.com/pilot?utm_source=pypi&utm_medium=registry&utm_campaign=community_commercial_2026_09) |
| Production operations | Not included | [Scope an evaluation](https://wolfxl.com/pilot?utm_source=pypi&utm_medium=registry&utm_campaign=community_commercial_2026_09) |
| Support | Public Community documentation and issues | [Scope an evaluation](https://wolfxl.com/pilot?utm_source=pypi&utm_medium=registry&utm_campaign=community_commercial_2026_09) |

For existing-template work, review the [bounded preservation approach](https://wolfxl.com/openpyxl-preservation?utm_source=pypi&utm_medium=registry&utm_campaign=community_commercial_2026_09).
Coming from Aspose.Cells for Python? See [when WolfXL fits instead](https://wolfxl.com/aspose-cells-python-alternative?utm_source=pypi&utm_medium=registry&utm_campaign=community_commercial_2026_09).
Use the [local fit check](https://wolfxl.com/fit-check?utm_source=pypi&utm_medium=registry&utm_campaign=community_commercial_2026_09)
to scope a representative-workbook test. Commercial self-service pricing is
[$30/month or $299/year for one seat](https://wolfxl.com/pricing?utm_source=pypi&utm_medium=registry&utm_campaign=community_commercial_2026_09).
