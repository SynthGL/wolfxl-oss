# WolfXL Community

**An openpyxl alternative for Python with a Rust backend. Edit existing Excel `.xlsx` and `.xlsm` files in place without losing formatting, charts, conditional formatting, data validation, or VBA macros, and recalculate common formulas without Excel. MIT licensed. In the committed read/write benchmark, WolfXL 2.0.21 measured 6.7-12.0x faster than openpyxl 3.1.5 (small-file gains vary).**

WolfXL Community reads, writes, and edits Excel `.xlsx` and `.xlsm` workbooks
through the openpyxl API, with parsing, serialization, and cell storage
implemented in Rust. Modify mode saves the cells you change and keeps the rest
of the file, and `calculate()` covers common Excel functions in Python.
Most openpyxl code runs after a one-line import change, so it is the first
thing to try when an openpyxl script is too slow. No Excel or LibreOffice
install is needed: wheels cover Linux (glibc and musl), macOS, and Windows.

WolfXL Community is the maintained, MIT-licensed 2.0 release line for supported workbook creation, reading, writing, streaming exports, and existing-workbook edits. It is a free product, not a trial.

## Install

```bash
python -m pip install wolfxl==2.0.30
```

## When to use WolfXL instead of openpyxl, XlsxWriter, or pandas

| Job | Use |
| --- | --- |
| An openpyxl script is too slow | WolfXL: change `from openpyxl import ...` to `from wolfxl import ...` |
| Fill or edit a styled template, and openpyxl drops its charts, sparklines, conditional formats, or data validation on save | WolfXL modify mode: `load_workbook(path, modify=True)` |
| Compute formula results in Python, where openpyxl `data_only=True` returns `None` for files Excel never saved | WolfXL `calculate()` for common functions; Commercial for Excel-matched results |
| Write a pandas DataFrame into an existing styled workbook (`to_excel` writes a new file) | WolfXL modify mode ([recipe](https://wolfxl.com/pandas-write-existing-excel-sheet?utm_source=pypi&utm_medium=registry&utm_campaign=community_commercial_2026_09)) |
| Only write new files, never read or edit them | XlsxWriter, or WolfXL if you also read or edit (XlsxWriter cannot open an existing file) |
| Only read cell values, as fast as possible, into a DataFrame | python-calamine, fastexcel, or polars |
| Render a sheet or chart to PDF or PNG on Linux without LibreOffice | WolfXL Commercial ([render](https://wolfxl.com/render-excel-python?utm_source=pypi&utm_medium=registry&utm_campaign=community_commercial_2026_09)); not in Community |
| Call Python code from Excel as worksheet functions or macros | [PyXLL](https://www.pyxll.com) |

On the 1.6-million-cell plain write in the
[committed ecosystem benchmark](https://github.com/SynthGL/wolfxl-oss/blob/58498c5f533353338514a84d5889501b51ed0b7d/benchmarks/results/2026-10-27-ecosystem-full-linux-epyc.md),
WolfXL 2.0.21 took 0.74 s, XlsxWriter 4.66 s, and pandas with the openpyxl
engine 10.62 s. DuckDB wrote the mixed-type file faster than WolfXL.

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
`.xlsm` macros. Since 2.0.10, `calculate()` followed by `save()` writes
supported formula results into the saved file's cached values; calculate again
after further edits, before saving.
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
- **Large files.** The [current receipts](https://github.com/SynthGL/wolfxl-oss/blob/58498c5f533353338514a84d5889501b51ed0b7d/benchmarks/results/2026-10-27-ecosystem-full-linux-epyc.md)
  record seconds and peak memory for large-workbook reads and writes, and the
  [edit receipts](https://github.com/SynthGL/wolfxl-oss/blob/58498c5f533353338514a84d5889501b51ed0b7d/benchmarks/results/2026-10-27-community-2.0.21-modify-linux-epyc.md)
  record existing-workbook edits.
- **Where openpyxl fits.** openpyxl is pure Python and installs anywhere
  Python runs. WolfXL needs a published wheel for your platform or a Rust
  toolchain to build from source.

## Performance, measured

WolfXL 2.0.21 versus openpyxl 3.1.5 on an AMD EPYC 9655 (16 virtual CPUs),
Linux, Python 3.13.15, median of five rounds: read and write speedups range
from 6.7x on mixed-type writes to 12.0x on unique-string writes, with 11.1x on
a 1.6-million-cell streaming read (`read_only=True` for both libraries) and
10.9x on a 1.6-million-cell plain write. WolfXL writes use the bulk
`write_rows()` call; openpyxl writes use `append()` per row.

Editing two cells in an existing workbook and saving it measured 27.4x faster
for a 10,000-row file and 35.8x for a 200,000-row file. Including the
harness's reload with openpyxl to verify the edit, the end-to-end ratios are
2.5x for both files; they are not included in the read/write headline.

All samples, workload boundaries, and reproduction commands are in the
[ecosystem results](https://github.com/SynthGL/wolfxl-oss/blob/58498c5f533353338514a84d5889501b51ed0b7d/benchmarks/results/2026-10-27-ecosystem-full-linux-epyc.md)
and the [edit results](https://github.com/SynthGL/wolfxl-oss/blob/58498c5f533353338514a84d5889501b51ed0b7d/benchmarks/results/2026-10-27-community-2.0.21-modify-linux-epyc.md).
These measurements are not a guarantee for every workbook. In the same run,
fastexcel and Polars (which wraps fastexcel) read the large file slightly
faster, and DuckDB wrote the mixed-type file faster.
Small-file gains vary: the separate [officelibs small-feature benchmark](https://officelibs.com/excel/)
reports about 1.7x for reads and 2.4x for writes, outside this range.

The [October 29 cached-formula read comparison](https://github.com/SynthGL/wolfxl-oss/blob/main/benchmarks/results/2026-10-29-ecosystem-formulas-wolfxl-2.0.26-linux-epyc.md)
measures WolfXL 2.0.26 on four sheets containing 1.6 million cells in total:
0.747 s for WolfXL, 5.970 s for openpyxl 3.1.5, and 0.560 s for fastexcel
0.20.2. WolfXL is 8.0x faster than openpyxl here and takes 1.33x fastexcel's
time. Polars also beats WolfXL on this fixture.[^formula-read]

[^formula-read]: AMD EPYC 9655, Linux, Python 3.13.15; five measured rounds
    after one warmup. Cached values are requested, with cell count and a sum
    of the final numeric formula column checked. Output representations differ.
    This is not full value/type parity or formula evaluation. See the report
    and its raw JSON for samples, identities, and limitations.

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
