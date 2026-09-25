# WolfXL

**openpyxl-compatible Excel read/write for Python, backed by Rust. MIT licensed.**

WolfXL Community reads, writes, and edits Excel `.xlsx` and `.xlsm` workbooks
through the openpyxl API, with parsing, serialization, and cell storage
implemented in Rust. It is for Python developers whose openpyxl jobs are slow
or run out of memory on large workbooks, and for teams that edit existing
Excel templates and need the untouched parts of the file kept intact. Most
openpyxl code runs after a one-line import change.

```bash
python -m pip install wolfxl
```

[![PyPI](https://img.shields.io/pypi/v/wolfxl)](https://pypi.org/project/wolfxl/)
[![Python 3.9+](https://img.shields.io/badge/python-3.9%2B-blue)](https://pypi.org/project/wolfxl/)
[![License: MIT](https://img.shields.io/github/license/SynthGL/wolfxl-oss)](LICENSE)

[Switch from openpyxl](#switch-from-openpyxl) ·
[When to use WolfXL](#when-to-use-wolfxl) ·
[Quick start](#quick-start) ·
[Benchmarks](#performance) ·
[Fidelity](#fidelity) ·
[Community vs Commercial](#community-and-commercial) ·
[wolfxl.com](https://wolfxl.com)

## Switch from openpyxl

Most openpyxl-shaped code needs only an import change:

```diff
- from openpyxl import Workbook, load_workbook
+ from wolfxl import Workbook, load_workbook
```

The rest of the code stays the same:

```python
from wolfxl import Workbook, load_workbook

workbook = Workbook()
sheet = workbook.active
sheet.append(["region", "revenue"])
for row in [("North", 1200), ("South", 950)]:
    sheet.append(row)
workbook.save("sales.xlsx")

workbook = load_workbook("sales.xlsx", read_only=True)
for row in workbook.active.iter_rows(min_row=2, values_only=True):
    print(row)
workbook.close()
```

For applications that cannot change every import, install the runtime alias
once at process startup:

```python
import wolfxl

wolfxl.install_as_openpyxl()

import openpyxl
```

Step-by-step guide: [openpyxl migration](https://wolfxl.com/openpyxl-migration).

## When to use WolfXL

- **openpyxl is slow or runs out of memory on a large file.** On a
  200,000-row by 8-column workbook (1.6 million cells), a full read with
  WolfXL took 0.60 s against 6.43 s for openpyxl 3.1.5 at 0.36x the peak
  memory, and the edit-two-cells-and-save phase took 0.25 s against 18.43 s.
  [Large-file receipts](https://wolfxl.com/openpyxl-large-files).
- **You are comparing openpyxl alternatives.** A 13-library run on one machine
  covers PyExcelerate, XlsxWriter, python-calamine, fastexcel, pandas, Polars,
  DuckDB, and others. Writing 1.6 million cells took 0.73 s with WolfXL
  (PyExcelerate 3.64 s, XlsxWriter 4.69 s), and reading them back took 0.39 s
  (python-calamine 0.58 s). python-calamine and fastexcel only read files;
  WolfXL reads, writes, and edits them.
  [openpyxl alternatives, measured](https://wolfxl.com/openpyxl-alternatives).
- **openpyxl drops parts of your template when it saves.** openpyxl warns
  that it will remove data validations, conditional formats, and sparklines
  it does not support, and its documentation says shapes are lost. On a sheet
  with an extension data validation and a sparkline, one cell edit saved by
  openpyxl 3.1.5 lost both; `load_workbook(path, modify=True)` in WolfXL 2.0.2
  kept both. Modify mode saves the cells you change and preserves unchanged
  parts within the documented boundaries; add `keep_vba=True` for `.xlsm`
  macros. [Why openpyxl loses template parts](https://wolfxl.com/openpyxl-preservation?utm_source=github&utm_medium=readme&utm_campaign=problem_pages_2026_09).
- **You are moving existing openpyxl code.** Check the
  [compatibility matrix](https://wolfxl.com/docs/migration/compatibility-matrix/)
  for the API you use and the
  [known limitations](https://wolfxl.com/docs/trust/limitations/) before
  switching a production path.

Community does not include native formula recalculation, PDF or image
rendering, format conversion, or VBA and Power Query operations. Those ship in
[WolfXL Commercial](https://wolfxl.com/pricing?utm_source=github&utm_medium=readme&utm_campaign=problem_pages_2026_09):

- [Recalculate formulas openpyxl leaves stale](https://wolfxl.com/calculate-excel-formulas-python?utm_source=github&utm_medium=readme&utm_campaign=problem_pages_2026_09)
- [Render sheets and charts to PDF or PNG without LibreOffice](https://wolfxl.com/render-excel-python?utm_source=github&utm_medium=readme&utm_campaign=problem_pages_2026_09)
- [Coming from Aspose.Cells for Python](https://wolfxl.com/aspose-cells-python-alternative?utm_source=github&utm_medium=readme&utm_campaign=problem_pages_2026_09)

See [Community and Commercial](#community-and-commercial).

![Median speedup over openpyxl 3.1.5 by benchmark case, from the committed results file](https://raw.githubusercontent.com/SynthGL/wolfxl-oss/main/assets/benchmarks/speedup-vs-openpyxl.svg)

Median speedups over openpyxl 3.1.5 range from 2.6x on small in-place edits to
27x on styled row writes through the bulk `write_styled_rows` API, with most
reads and writes between 7x and 14x (wolfxl 2.0.1 PyPI wheel, Apple M4 Pro,
Python 3.13.9, median of 5 rounds). Every chart in this README is generated
from a [committed raw results file](benchmarks/results/2026-08-18-community-2.0.1-vs-openpyxl-3.1.5.md),
never edited by hand.

## Quick start

Install the current Community release:

```bash
python -m pip install wolfxl==2.0.3
```

WolfXL Community supports Python 3.9 and newer CPython versions for which a
wheel is published.

```python
from wolfxl import Alignment, Font, PatternFill, Workbook, load_workbook

workbook = Workbook()
sheet = workbook.active
sheet.title = "Summary"
sheet["A1"] = "Revenue"
sheet["A1"].font = Font(bold=True, color="FFFFFF")
sheet["A1"].fill = PatternFill(fill_type="solid", fgColor="336699")
sheet["B1"] = 125000
sheet["B1"].alignment = Alignment(horizontal="right")
workbook.save("report.xlsx")

loaded = load_workbook("report.xlsx")
print(loaded["Summary"]["B1"].value)
loaded.close()
```

### For AI coding agents

[`skills/wolfxl-xlsx`](skills/wolfxl-xlsx/SKILL.md) is an agent skill for
editing existing workbooks without losing the parts the edit did not touch.
It tells the agent to edit in modify mode, recalculate with WolfXL instead of
a LibreOffice round trip, and run `verify` to confirm that every package part
and sheet feature from the source is still present. Copy the directory into
your agent's skills folder, for example `~/.claude/skills/wolfxl-xlsx`.

## Performance

Full openpyxl comparison from the committed benchmark run (wolfxl 2.0.1 PyPI
wheel, Apple M4 Pro, Python 3.13.9, median of 5 rounds):

![1.6 million cells: wall-clock seconds for wolfxl and openpyxl](https://raw.githubusercontent.com/SynthGL/wolfxl-oss/main/assets/benchmarks/large-file-seconds.svg)

![1.6 million cells: peak memory for wolfxl and openpyxl](https://raw.githubusercontent.com/SynthGL/wolfxl-oss/main/assets/benchmarks/large-file-memory.svg)

### Against other open-source Python Excel libraries

Cross-library comparison on a separate machine (AMD EPYC 9655, x86_64 Linux,
Python 3.13, median of 5 rounds) against twelve other libraries: openpyxl,
XlsxWriter, PyExcelerate, pylightxl, pandas, Polars, DuckDB, Tablib, pyexcel,
python-calamine, fastexcel, and xlsx2csv. The bar for inclusion is xlsx
support, no external application, and roughly one million PyPI downloads per
month. To keep the baselines honest, the large plain write and the memory
pass also measure openpyxl in write_only mode, XlsxWriter in constant_memory
mode, and pandas with the xlsxwriter engine. Each library is measured only
inside its supported scope; write-only, read-only, DataFrame, and SQL
specialists are labeled:

![Write 200,000 x 8 plain values across thirteen libraries](https://raw.githubusercontent.com/SynthGL/wolfxl-oss/main/assets/benchmarks/ecosystem-write-large.svg)

![Write 10,000 x 5 mixed types across thirteen libraries](https://raw.githubusercontent.com/SynthGL/wolfxl-oss/main/assets/benchmarks/ecosystem-write-mixed.svg)

![Write 100,000 x 5 unique strings across thirteen libraries](https://raw.githubusercontent.com/SynthGL/wolfxl-oss/main/assets/benchmarks/ecosystem-write-strings.svg)

![Read 200,000 x 8, all values, across thirteen libraries](https://raw.githubusercontent.com/SynthGL/wolfxl-oss/main/assets/benchmarks/ecosystem-read-large.svg)

![Peak memory: write 200,000 x 8 across thirteen libraries](https://raw.githubusercontent.com/SynthGL/wolfxl-oss/main/assets/benchmarks/ecosystem-memory-write.svg)

![Peak memory: read 200,000 x 8 across thirteen libraries](https://raw.githubusercontent.com/SynthGL/wolfxl-oss/main/assets/benchmarks/ecosystem-memory-read.svg)

wolfxl leads every case in this run, including reads (387 ms vs 394 ms for
Polars and 403 ms for fastexcel, which return Arrow-backed tables rather than
Python cell values). The closest overall rival is DuckDB's excel extension,
which wins the small mixed-type write outright (25 ms vs 35 ms, timed from a
registered DataFrame) and stays within 1.4x elsewhere. The streaming modes
own write memory: openpyxl write_only and XlsxWriter constant_memory peak at
234 MiB, effectively the cost of the input grid itself, where wolfxl's fully
materialized workbook peaks at 610 MiB while writing 5-8x faster than
either. pylightxl's pure-Python writer scales quadratically (241 s on the
large plain write, 1,438 s on unique strings) and its bars are clipped to
keep the charts readable.

Speedups vary by workload, and small workbooks see smaller wins. Raw results,
the benchmark harnesses, and reproduction instructions are in
[`benchmarks/`](benchmarks/README.md).

## Fidelity

The [round-trip fidelity harness](fidelity-harness/README.md) compares workbook
packages before and after a no-edit save. Run it on your own files, inspect the
typed part and relationship differences, and add another engine through the
documented adapter protocol.

## Community and Commercial

| | WolfXL Community | WolfXL Commercial |
|---|---|---|
| License | MIT | Commercial |
| Release line | Maintained 2.0 generation | Current 2.1+ generation |
| Workbook I/O | Included | Included |
| Existing 2.0 modify and pivot APIs | Included | Current implementations and fixes |
| Native recalculation | Not included | Included |
| Render, PDF, and image output | Not included | Included |
| Format conversion | Not included | Included |
| VBA and Power Query operations | Not included | Included |
| Production operations SDK | Not included | Included |
| Direct support | Community issues | Included with paid plans |

Community receives critical correctness and security fixes. New engines,
expanded compatibility work, production operations, and direct support ship in
[WolfXL Commercial](https://wolfxl.com).

This split keeps the useful Excel I/O layer open while funding the
compatibility, fidelity, and support work required by production workbook
pipelines.

Use [wolfxl.com](https://wolfxl.com) for the current Commercial package,
evaluation access, pricing, compatibility information, and support. Commercial
source and releases are maintained separately and are not part of this
repository.

## Development

Prerequisites: a supported CPython, Rust, and `maturin`.

```bash
python -m pip install maturin pytest defusedxml openpyxl Pillow
maturin develop
pytest tests/test_community_distribution.py -q
```

The distribution-boundary test verifies the Community version, compiled
backends, and absence of Commercial-only Python modules.

See [CONTRIBUTING.md](CONTRIBUTING.md) for contribution guidelines and
[SECURITY.md](SECURITY.md) for how to report a vulnerability.

## License

WolfXL Community is available under the [MIT License](LICENSE).
