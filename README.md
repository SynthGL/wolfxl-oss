# WolfXL

**Edit existing Excel files from Python without losing formatting, 3.0-17.3x faster than openpyxl on the committed read/write benchmark (small-file gains vary), with an openpyxl-compatible API. MIT licensed.**

WolfXL Community reads, writes, and edits Excel `.xlsx` and `.xlsm` workbooks
through the openpyxl API, with parsing, serialization, and cell storage
implemented in Rust. Modify mode saves the cells you change and keeps the rest
of the file, and `calculate()` covers common Excel functions in Python.
Most openpyxl code runs after a one-line import change.

```bash
python -m pip install wolfxl
```

[![PyPI](https://img.shields.io/pypi/v/wolfxl)](https://pypi.org/project/wolfxl/)
[![Python 3.9+](https://img.shields.io/badge/python-3.9%2B-blue)](https://pypi.org/project/wolfxl/)
[![License: MIT](https://img.shields.io/github/license/SynthGL/wolfxl-oss)](LICENSE)

[Edit a workbook](#edit-an-existing-workbook) ·
[Compared with openpyxl](#compared-with-openpyxl) ·
[Switch from openpyxl](#switch-from-openpyxl) ·
[Quick start](#quick-start) ·
[Benchmarks](#performance) ·
[Fidelity](#fidelity) ·
[Community vs Commercial](#community-and-commercial) ·
[wolfxl.com](https://wolfxl.com)

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
[Edit Excel in Python without losing formatting](https://wolfxl.com/openpyxl-preservation?utm_source=github&utm_medium=readme&utm_campaign=problem_pages_2026_09).

The following free recalculation-on-open recipe requires the fix listed
under [Unreleased](CHANGELOG.md#unreleased); published Community 2.0.8 does
not persist changes to this flag. In that fixed version, keep the workbook
open in modify mode as above, then:

```python
wb["Summary"]["B2"] = 1500
wb.calculation.fullCalcOnLoad = True
wb.save("report-updated.xlsx")
```

Excel recalculates when it opens the saved workbook. This flag does not
calculate formulas in Python: cached values remain stale until Excel or
another calculation engine recalculates. Close the workbook after saving.

## Compared with openpyxl

- **Same API for the covered surface.** Change one import
  ([Switch from openpyxl](#switch-from-openpyxl)). openpyxl implements more of its own
  API, so check the
  [compatibility matrix](https://wolfxl.com/docs/migration/compatibility-matrix/)
  and the [known limitations](https://wolfxl.com/docs/trust/limitations/)
  before switching a production path.
- **Existing templates.** openpyxl warns that it will remove data
  validations, conditional formats, and sparklines it does not support, and
  its documentation says shapes are lost. On a sheet with an extension data
  validation and a sparkline, one cell edit saved by openpyxl 3.1.5 lost both;
  `load_workbook(path, modify=True)` in WolfXL 2.0.2 kept both.
- **Formulas.** openpyxl stores formula text and never computes it; with
  `data_only=True` it returns the value Excel last cached, or `None` for a
  file Excel never opened. WolfXL Community `calculate()` covers common
  functions only, in Python. Need results that match Excel? Commercial
  includes a native engine verified on 704 Excel-calculated cases
  ([pricing](https://wolfxl.com/pricing?utm_source=github&utm_medium=readme&utm_campaign=problem_pages_2026_09)).
- **Large files.** In the current benchmark's 200,000-row plain workload,
  a full read with WolfXL took 0.305 s against 4.380 s for openpyxl 3.1.5.
  The edit-two-cells-and-save phase took 0.159 s against 9.202 s; the full
  edit benchmark, including verification with openpyxl, took 5.458 s
  against 14.509 s. See the [current receipts](benchmarks/results/2026-10-02-community-2.0.8-vs-openpyxl-3.1.5.md).
- **Other alternatives.** python-calamine and fastexcel specialize in reading;
  XlsxWriter and PyExcelerate specialize in writing. WolfXL reads, writes,
  and edits through an openpyxl-compatible API. The current speed claims
  compare only WolfXL and openpyxl, not these specialists.
- **Where openpyxl fits.** openpyxl is pure Python and installs anywhere
  Python runs. WolfXL needs a published wheel for your platform or a Rust
  toolchain to build from source.

Community does not include native formula recalculation, PDF or image
rendering, format conversion, or VBA and Power Query operations. Those ship in
[WolfXL Commercial](https://wolfxl.com/pricing?utm_source=github&utm_medium=readme&utm_campaign=problem_pages_2026_09):

- [Recalculate formulas openpyxl leaves stale](https://wolfxl.com/calculate-excel-formulas-python?utm_source=github&utm_medium=readme&utm_campaign=problem_pages_2026_09)
- [Render sheets and charts to PDF or PNG without LibreOffice](https://wolfxl.com/render-excel-python?utm_source=github&utm_medium=readme&utm_campaign=problem_pages_2026_09)
- [Coming from Aspose.Cells for Python](https://wolfxl.com/aspose-cells-python-alternative?utm_source=github&utm_medium=readme&utm_campaign=problem_pages_2026_09)

See [Community and Commercial](#community-and-commercial).

Median read and write speedups over openpyxl 3.1.5 range from 3.0x on styled
reads to 17.3x on multi-sheet bulk writes (WolfXL 2.0.8 PyPI wheel,
Apple M5 Pro, Python 3.13.9, median of 5 rounds). Bulk writes use a
different API shape from openpyxl's per-row or per-cell calls; workload
details and all samples are in the [current results](benchmarks/results/2026-10-02-community-2.0.8-vs-openpyxl-3.1.5.md).
This is a measured range on this machine, not a guarantee for every workbook.
Small-file gains vary: the separate [officelibs small-feature benchmark](https://officelibs.com/excel/)
reports about 1.7x for reads and 2.4x for writes, outside this range.

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

## Quick start

Install the current Community release:

```bash
python -m pip install wolfxl==2.0.8
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

The current comparison uses the published WolfXL 2.0.8 and openpyxl 3.1.5
packages on Apple M5 Pro, macOS 26.5.1, Python 3.13.9. Every timed case,
including the large workloads, reports the median of five measured rounds.

| Workload | openpyxl seconds | WolfXL seconds | Speedup |
| --- | ---: | ---: | ---: |
| Styled cell read | 0.051807 | 0.017246 | 3.0x |
| Multi-sheet bulk write | 0.083623 | 0.004842 | 17.3x |
| Large plain full read | 4.380110 | 0.304859 | 14.4x |
| Large plain bulk write | 3.399683 | 0.228258 | 14.9x |
| Large two-cell edit, including verification | 14.509416 | 5.458442 | 2.7x |

Read and write ratios span **3.0-17.3x**. End-to-end edit ratios span
2.4-2.7x and include openpyxl verification; edit-only phases are reported
separately, not folded into the headline range. The large plain case is
configured for 200,000 rows and eight columns, but the committed harness
generates only five populated columns (one million cells). See the result
notes rather than interpreting its nominal units-per-second as populated
cell throughput.

The [results and reproduction notes](benchmarks/results/2026-10-02-community-2.0.8-vs-openpyxl-3.1.5.md)
include raw samples, machine and package identities, warmup policy, and memory
measurements. Earlier runs remain under [`benchmarks/results/`](benchmarks/results/)
as historical evidence, not current headline claims.

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
