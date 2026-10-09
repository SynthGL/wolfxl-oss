# WolfXL

**Edit existing Excel files from Python without losing formatting, with an openpyxl-compatible API. MIT licensed.**

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

Modify mode patches changed cells and required owned package metadata while
preserving unchanged styles, charts, and parts within the documented boundaries;
add `keep_vba=True` for
`.xlsm` macros. `calculate()` returns the computed values and leaves the cached
results in the saved file unchanged.
[Edit Excel in Python without losing formatting](https://wolfxl.com/openpyxl-preservation?utm_source=github&utm_medium=readme&utm_campaign=problem_pages_2026_09).

The following free recalculation-on-open recipe requires Community 2.0.9
or newer; earlier versions do not persist changes to this flag.
Keep the workbook open in modify mode as above, then:

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
- **Large files.** The [proposed-source comparison](#performance) measures
  styled Cell reads and two-cell edits at top, middle, and bottom positions.
  It reports the operation separately from the same independent verification
  used for both engines; runtime depends on workbook shape and read mode.
- **Other alternatives.** python-calamine and fastexcel specialize in reading;
  XlsxWriter and PyExcelerate specialize in writing. WolfXL reads, writes,
  and edits through an openpyxl-compatible API. The performance comparison
  below measures WolfXL and openpyxl, not these specialists.
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
python -m pip install wolfxl==2.0.9
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

### Sparse read-only follow-up: 2026-10-09

Missing XML cells now use a lightweight coordinate-aware proxy. The sparse
fixture has 800,000 positions but only 7,816 stored values. Full proxy
constructions fall from 800,000 to 7,816, and style-resolution calls from
3,200,000 to 31,264. Coordinates, default styles, border resolution, close
behavior, and mutation refusal are retained; authored empty cells and gradients
at source style zero keep the original style-aware path.

Five fresh-process medians, one warmup, GitHub-hosted Ubuntu 24.04, CPython
3.12.15, openpyxl 3.1.5, standard release Community wheel. Load + identical public
fill/font/number-format scan + close; all benchmark signatures match.

| Read-only styled workload | Before (s) | After (s) | openpyxl (s) | Before/after |
| --- | ---: | ---: | ---: | ---: |
| Sparse 25,000 × 32, 7,816 stored values | 3.279210 | 0.840991 | 0.294885 | 3.899× |
| Dense 20,000 × 5, 100,000 stored values | 0.538393 | 0.534949 | 1.360873 | 1.006× |

Sparse elapsed time falls 74.4%, but remains **2.85× slower than openpyxl**.
WolfXL still constructs 792,184 coordinate-aware blank proxies; openpyxl reuses
an empty-cell singleton. Dense timing is effectively unchanged. All 70 targeted
streaming compatibility tests pass in the source-built lane.

Baseline main `7b796d4dcc30` → candidate code `1d0927e623a3` (merged in PR #48).
The [raw receipt](benchmarks/sparse-styled-source-built-20261009.json) records all
samples, source/module/native hashes, signatures, and untimed profiles; the
[workflow](https://github.com/SynthGL/wolfxl-oss/actions/runs/37885386977) builds
the standard candidate wheel and confirms native sources are unchanged from the
base. Both Python overlays use that same binary and differ only in
`_streaming.py`. These measurements use a separate runner from the October 8
matrix below; their ratios are not multiplied into its cumulative figures.

### Pinned source comparison: 2026-10-08

This comparison uses unreleased Community source builds and a pinned
source baseline on Linux, CPython 3.12, with openpyxl 3.1.5. Both WolfXL variants
use standard release builds with Community features. These results are separate
from published-wheel measurements and apply to the fixed synthetic fixtures.

Measured source baseline `93e48b23605c` → final `02884309b62d`.

**Variance note:** some operation ranges exceed 20% of their medians. Use the linked core and guard raw ranges to assess these results.

| Workload / mode | Baseline operation (s) | Final operation (s) | Baseline/final operation | openpyxl/final operation | Baseline/final total |
| --- | ---: | ---: | ---: | ---: | ---: |
| Two-cell edit: top (modify) | 0.4718 | 0.3299 | 1.43× | 46.28× | 1.42× |
| Two-cell edit: middle (modify) | 0.8365 | 0.3834 | 2.18× | 38.35× | 1.18× |
| Two-cell edit: bottom (modify) | 1.1251 | 0.4812 | 2.34× | 31.61× | 1.11× |
| Two-cell edit: merged fixture (modify) | 5.5876 | 0.6350 | 8.80× | Not measured | 8.74× |
| Styled 2,000 × 5 (eager Cells) | 0.0522 | 0.0463 | 1.13× | 2.25× | 1.13× |
| Styled 20,000 × 5 (read-only Cells) | 0.8446 | 0.3369 | 2.51× | 2.25× | 2.51× |
| Merged 20,000 × 5 (eager Cells) | 0.9529 | 0.7725 | 1.23× | 1.39× | 1.23× |

Operation covers load, assignment, save, and close for edits; reads cover load,
the identical public Cell/font/fill/number-format loop, and close. Edit total adds
fixture copy and the same independent bounded openpyxl verification. Edit
outputs receive independent full value/type/style-meaning/merge and ZIP/XML
package checks outside both timings. Styled reads match rows, first-column
checksum, and styled-cell count against openpyxl; their input ZIP integrity is
checked separately. This read signature does not certify every cell's value,
type, or style meaning. Late-row verification streams the worksheet XML prefix
and costs more than verification near the top.

Separate untimed [styled-read validation](docs/performance/2026-10-08/cumulative/styled-read-validation.md)
checks every coordinate, including sparse and merged placeholders. Coordinates,
values, exact Python value types, Excel data types, geometry, merges, alignment,
and number formats match openpyxl on these fixtures. Four baseline merged-border
color mismatches are fixed. Remaining font metadata differences are identical to
the baseline; Community also retains fill ARGB `FFFFD966` versus openpyxl
`00FFD966`. These getter checks do not certify visual or rendered equivalence.

Plain edit fixtures contain 200,000 rows × five populated columns: 1,000,000
value cells, despite the historical eight-column setting. The merged edit
fixture contains 999,997 values. Small styled reads contain 2,000 × five cells;
large styled core reads contain 20,000 × five, with 99,997 values in the merged
case. Eager and `read_only=True` modes remain separate workloads. Read-only
limits Python Cell allocation; the native worksheet model is still eager.

| Workload / mode | Baseline operation (s) | Final operation (s) | Baseline/final operation | openpyxl/final operation | Baseline/final total |
| --- | ---: | ---: | ---: | ---: | ---: |
| Styled 25,000 × 5 (eager Cells) | 0.7980 | 0.5483 | 1.46× | 2.66× | 1.46× |
| Sparse 25,000 × 32 (eager Cells) | 3.4357 | 1.7656 | 1.95× | 1.84× | 1.95× |
| 1,024 format variants (eager Cells) | 0.9805 | 0.6669 | 1.47× | 2.13× | 1.47× |
| Sparse 25,000 × 32 (read-only Cells) | 4.3397 | 1.9321 | 2.25× | 0.08× | 2.25× |
| 1,024 format variants (read-only Cells) | 1.2953 | 0.5705 | 2.27× | 1.85× | 2.27× |

**Pinned 2026-10-08 sparse read-only guard:** that source build was 12.01× slower than openpyxl on this fixture (1.9321 s vs 0.1609 s), with a direct 2.25× baseline/final operation ratio. Eager and read-only modes have different tradeoffs; this result is not covered by a blanket speed claim.

The guard fixtures use 25,000 rows: dense and high-cardinality sheets contain
125,000 values; sparse sheets have 32-column dimensions and 7,816 stored values.
The high-cardinality guard configures 1,024 number-format variants. Sparse
read-only behavior is reported explicitly rather than inferred from dense reads.

Engine ratios divide matched WolfXL baseline/final medians; openpyxl comparisons
divide openpyxl/final medians. Ratios below 1 mean the final variant is slower.
Isolated PR gains are never multiplied, and the separate values/style-ID record
API choice is excluded from cumulative engine gains. Five fresh-process trials
follow one warmup, with serial, reproducibly shuffled variant order. Sample ranges and high-variance cases are
in the [core evidence](docs/performance/2026-10-08/cumulative/core.md) and
[guard evidence](docs/performance/2026-10-08/cumulative/guards.md), alongside raw
JSON samples, fixture hashes, source/native/build identities, and validation
boundaries. The [source bindings](docs/performance/2026-10-08/cumulative/source-bindings.json)
map measured build inputs to the proposed PR trees. Sample ranges are descriptive
rather than confidence intervals. The fixed-fixture headless checks are not
Microsoft Excel certification. [Build and functional gates](docs/performance/2026-10-08/cumulative/build-gates.md)
are reported separately; their overlapping per-run test counts must not be added
to imply distinct coverage.

For ingestion that needs values and workbook-local style IDs, see the existing
[bounded record recipe](docs/performance/style-id-records.md). It uses a different
API and has explicit date, elapsed-duration, merge, and pending-style limits.
The retained [intermediate top-regression experiment](docs/performance/2026-10-08/cumulative/intermediate-core.md)
record the earlier source stage separately and do not supply the final ratios.

### Historical published-wheel comparison: 2026-10-02

The released WolfXL 2.0.8 PyPI wheel was measured against openpyxl 3.1.5 on
Apple M5 Pro, macOS 26.5.1, Python 3.13.9. Each workload reports five-round
medians. This historical protocol differs from the proposed Linux source run
above and must not be combined with it.

| Workload | openpyxl seconds | WolfXL seconds | Speedup |
| --- | ---: | ---: | ---: |
| Styled cell read | 0.051807 | 0.017246 | 3.0x |
| Multi-sheet bulk write | 0.083623 | 0.004842 | 17.3x |
| Large plain full read | 4.380110 | 0.304859 | 14.4x |
| Large plain bulk write | 3.399683 | 0.228258 | 14.9x |
| Large two-cell edit, including verification | 14.509416 | 5.458442 | 2.7x |

Bulk writes use a different API shape from openpyxl's per-row/per-cell calls.
The historical large fixture also has five populated columns, despite its
nominal eight-column configuration. Its edit total includes the historical
openpyxl verification, so it is not an operation-only ratio. Read the
[historical receipts and reproduction notes](benchmarks/results/2026-10-02-community-2.0.8-vs-openpyxl-3.1.5.md)
for raw samples, environment identities, warmup policy, and memory measurements.
Older runs remain under [benchmarks/results/](benchmarks/results/).

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
