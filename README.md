# WolfXL

**An openpyxl alternative for Python with a Rust backend. Edit existing Excel `.xlsx` and `.xlsm` files in place without losing formatting, charts, conditional formatting, data validation, or VBA macros, and recalculate common formulas without Excel. WolfXL 2.0.21 measured 6.7-12.0x faster than openpyxl 3.1.5 on the [committed read/write benchmark](benchmarks/results/2026-10-27-ecosystem-full-linux-epyc.md) (small-file gains vary). MIT licensed.**

WolfXL Community reads, writes, and edits Excel `.xlsx` and `.xlsm` workbooks
through the openpyxl API, with parsing, serialization, and cell storage
implemented in Rust. Modify mode saves the cells you change and keeps the rest
of the file, and `calculate()` covers common Excel functions in Python.
Most openpyxl code runs after a one-line import change, so it is the first
thing to try when an openpyxl script is too slow. No Excel or LibreOffice
install is needed: wheels cover Linux (glibc and musl), macOS, and Windows.

```bash
python -m pip install wolfxl
```

[![PyPI](https://img.shields.io/pypi/v/wolfxl)](https://pypi.org/project/wolfxl/)
[![Python 3.11+](https://img.shields.io/badge/python-3.11%2B-blue)](https://pypi.org/project/wolfxl/)
[![License: MIT](https://img.shields.io/github/license/SynthGL/wolfxl-oss)](LICENSE)

[When to use WolfXL](#when-to-use-wolfxl-instead-of-openpyxl-xlsxwriter-or-pandas) ·
[Edit a workbook](#edit-an-existing-workbook) ·
[Compared with openpyxl](#compared-with-openpyxl) ·
[Switch from openpyxl](#switch-from-openpyxl) ·
[Quick start](#quick-start) ·
[Benchmarks](#performance) ·
[Fidelity](#fidelity) ·
[Community vs Commercial](#community-and-commercial) ·
[wolfxl.com](https://wolfxl.com)

## When to use WolfXL instead of openpyxl, XlsxWriter, or pandas

| Job | Use |
| --- | --- |
| An openpyxl script is too slow | WolfXL: change the import ([Switch from openpyxl](#switch-from-openpyxl)) |
| Fill or edit a styled template, and openpyxl drops its charts, sparklines, conditional formats, or data validation on save | WolfXL modify mode ([Edit an existing workbook](#edit-an-existing-workbook)) |
| Compute formula results in Python, where openpyxl `data_only=True` returns `None` for files Excel never saved | WolfXL `calculate()` for common functions; Commercial for Excel-matched results |
| Write a pandas DataFrame into an existing styled workbook (`to_excel` writes a new file) | WolfXL modify mode ([recipe](https://wolfxl.com/pandas-write-existing-excel-sheet?utm_source=github&utm_medium=readme&utm_campaign=problem_pages_2026_09)) |
| Only write new files, never read or edit them | XlsxWriter, or WolfXL if you also read or edit (XlsxWriter cannot open an existing file) |
| Only read cell values, as fast as possible, into a DataFrame | python-calamine, fastexcel, or polars |
| Render a sheet or chart to PDF or PNG on Linux without LibreOffice | WolfXL Commercial ([render](https://wolfxl.com/render-excel-python?utm_source=github&utm_medium=readme&utm_campaign=problem_pages_2026_09)); not in Community |
| Call Python code from Excel as worksheet functions or macros | [PyXLL](https://www.pyxll.com) |

On the 1.6-million-cell plain write in the
[committed ecosystem benchmark](benchmarks/results/2026-10-27-ecosystem-full-linux-epyc.md),
WolfXL 2.0.21 took 0.74 s, XlsxWriter 4.66 s, and pandas with the openpyxl
engine 10.62 s. DuckDB wrote the mixed-type file faster than WolfXL.

The [ecosystem read comparison](#ecosystem-reads-2026-10-29) includes
read-only specialists. WolfXL is for code that reads, writes, and edits through
one openpyxl-compatible API.

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
add `keep_vba=True` for `.xlsm` macros.

Since 2.0.10, `calculate()` followed by `save()` writes supported formula
results into the saved file, so readers that use cached values (including
`load_workbook(path, data_only=True)`) see them. After changing inputs,
formulas, sheet names, or defined names, calculate again before saving.
Results Community cannot compute are left uncached. In 2.0.9 and earlier,
`calculate()` returned the values without saving them.
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

If an openpyxl script is too slow, try WolfXL before rewriting it for another
library. Most openpyxl-shaped code needs only an import change:

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
python -m pip install wolfxl==2.0.36
```

WolfXL Community supports Python 3.11 and newer CPython versions for which a
wheel is published. Version 2.0.36 is the last release supporting Python 3.9
and 3.10; pip on those versions continues to resolve compatible 2.0.x releases.

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

### Limit workbook input

`load_workbook` bounds input before materializing XLSX cells in full,
`read_only=True`, and `modify=True` modes. Shared-string references reuse one
string object per workbook rather than copying the text into every cell.

```python
from wolfxl import LimitExceeded, Limits, load_workbook

try:
    workbook = load_workbook(
        "incoming.xlsx",
        limits=Limits(max_cells=500_000, max_shared_string_bytes=32 * 1024**2),
    )
except LimitExceeded as error:
    print(error)  # limit name, observed value, configured value, and limits= knob
```

| Field | Default | Counts |
| --- | --- | --- |
| `max_package_bytes` | 4 GiB | Sum of uncompressed ZIP member sizes; also bounds the compressed file size of path loads |
| `max_part_bytes` | 512 MiB | Uncompressed bytes in one member |
| `max_cells` | 64,000,000 | Stored cell elements across all worksheets |
| `max_shared_string_bytes` | 256 MiB | Decoded UTF-8 text in the shared-string table |
| `max_materialized_string_bytes` | 512 MiB | Decoded shared/inline text, string cached values, error/date text, and formulas |

`Limits` is immutable. Omit `limits=` or pass `None` for these defaults. Shared
strings count once, not once per reference. Worksheet aliases count once per
sheet reference; shared-formula template bytes count once per formula cell,
including children before their master. Formula translation can change the
final text length. A value of zero permits no content for that counter.
Existing ZIP entry-count, compression-ratio, and environment admission guards
still apply; `limits=` does not weaken those guards. These are input budgets,
not a process-memory or execution-time sandbox. They apply to XLSX-family
packages, including decrypted inputs, not legacy XLS or XLSB.

Cell and text budgets use a streaming admission pass, including for read-only
loads. This reads worksheet XML before iteration and can add startup work for
large inputs.

After format detection, path-backed loads open the file once and admit, load
and later read only that opened file, never reopening the caller's path. On
Linux and macOS the source is a read-only memory map where supported: pages load
on demand without copying the compressed package into a private in-memory
buffer. On unsupported filesystems and Windows the compressed package is read
into memory instead; Windows cannot replace or delete a mapped file. The
compressed file also must fit `max_package_bytes`.
Nonstandard workbook parts selected by package relationships receive the same
cell and text budgets before normalization.
Replacing the path afterwards, including an atomic rename over it, does not
change what the open workbook reads. The map is not an immutable copy: writing
to or truncating that same file in place while the source map remains alive,
including when a reader outlives `close()`, changes what it reads. Reading past
a truncation can terminate the process with `SIGBUS`. Replace workbooks by
writing a new file and renaming it over the old one. Feature readers consume the
same source without reopening the caller's path. Modify mode uses a private
temporary copy, removed on close or collection.
Saving back to the original path advances the private source and then publishes
it atomically through a new file renamed over the path, adding a second write of
the compressed package.
Atomic publication requires write permission on the target directory. The
publication step keeps existing permission bits and resolves symlinks;
ownership, ACLs and extended attributes are not copied, and replacing one
hard-link name breaks its sharing. Later package rewrites performed by some
saves, such as VBA package-shape normalization, keep their existing replacement
behavior.

On Linux and macOS, `close()` releases the map unless a live reader or row
iterator still uses it; it is then released when the last reader is collected.
Until release, a replaced source file can keep using disk space. On Python
versions before 3.13 each map also holds an open file descriptor. Save workbooks
before closing them.

### For AI coding agents

[`skills/wolfxl-xlsx`](skills/wolfxl-xlsx/SKILL.md) is an agent skill for
editing existing workbooks without losing the parts the edit did not touch.
It tells the agent to edit in modify mode, recalculate with WolfXL instead of
a LibreOffice round trip, and run `verify` to confirm that every package part
and sheet feature from the source is still present. Copy the directory into
your agent's skills folder, for example `~/.claude/skills/wolfxl-xlsx`.

## Performance

### Ecosystem reads: 2026-10-29

The [plain and multi-sheet formula comparison](benchmarks/results/2026-10-29-ecosystem-formulas-wolfxl-2.0.26-linux-epyc.md)
includes WolfXL 2.0.26, openpyxl 3.1.5, fastexcel 0.20.2, python-calamine,
Polars, and other readers. Both fixtures contain 1.6 million cells. For the
formula case, every engine reads four sheets with numeric and string cached
formula results; the harness requests those cached values.

Five measured rounds follow one discarded warmup per engine and case on an
AMD EPYC 9655, Linux, Python 3.13.15. See the
[raw samples and identities](benchmarks/results/2026-10-29-ecosystem-formulas-wolfxl-2.0.26-linux-epyc.json)
for package versions, fixture hashes, and measurement boundaries.

| Median read time | Plain values | Multi-sheet formula fixture |
| --- | ---: | ---: |
| WolfXL 2.0.26 | 0.475 s | 0.747 s |
| openpyxl 3.1.5 | 5.386 s | 5.970 s |
| fastexcel 0.20.2 | 0.466 s | 0.560 s |

WolfXL takes 1.02x fastexcel's time on plain values and 1.33x on the formula
fixture, where Polars also beats it. WolfXL is 11.3x faster than openpyxl on
plain values and 8.0x on formulas. On the same machine, 2.0.26 reads the
formula fixture 1.89x faster than 2.0.25.[^ecosystem-read]

[^ecosystem-read]: These are warmed read-and-consume workloads with different
    native output representations. Formula-case validation covers cell count
    and the final numeric formula column's sum, not every value or type.
    No memory pass was run. The raw report records identity limitations,
    including the uncaptured DuckDB extension version. This is not a
    formula-evaluation benchmark or a general speed guarantee.

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
The [raw receipt](https://github.com/SynthGL/wolfxl-oss/blob/206bf4e1cdc62c8121569066362c1521d9b94a72/benchmarks/sparse-styled-source-built-20261009.json) records all
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

Separate untimed [styled-read validation](https://github.com/SynthGL/wolfxl-oss/blob/206bf4e1cdc62c8121569066362c1521d9b94a72/docs/performance/2026-10-08/cumulative/styled-read-validation.md)
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
in the [core evidence](https://github.com/SynthGL/wolfxl-oss/blob/206bf4e1cdc62c8121569066362c1521d9b94a72/docs/performance/2026-10-08/cumulative/core.md) and
[guard evidence](https://github.com/SynthGL/wolfxl-oss/blob/206bf4e1cdc62c8121569066362c1521d9b94a72/docs/performance/2026-10-08/cumulative/guards.md), alongside raw
JSON samples, fixture hashes, source/native/build identities, and validation
boundaries. The [source bindings](https://github.com/SynthGL/wolfxl-oss/blob/206bf4e1cdc62c8121569066362c1521d9b94a72/docs/performance/2026-10-08/cumulative/source-bindings.json)
map measured build inputs to the proposed PR trees. Sample ranges are descriptive
rather than confidence intervals. The fixed-fixture headless checks are not
Microsoft Excel certification. [Build and functional gates](https://github.com/SynthGL/wolfxl-oss/blob/206bf4e1cdc62c8121569066362c1521d9b94a72/docs/performance/2026-10-08/cumulative/build-gates.md)
are reported separately; their overlapping per-run test counts must not be added
to imply distinct coverage.

For ingestion that needs values and workbook-local style IDs, see the existing
[bounded record recipe](https://github.com/SynthGL/wolfxl-oss/blob/206bf4e1cdc62c8121569066362c1521d9b94a72/docs/performance/style-id-records.md). It uses a different
API and has explicit date, elapsed-duration, merge, and pending-style limits.
The retained [intermediate top-regression experiment](https://github.com/SynthGL/wolfxl-oss/blob/206bf4e1cdc62c8121569066362c1521d9b94a72/docs/performance/2026-10-08/cumulative/intermediate-core.md)
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
[historical receipts and reproduction notes](https://github.com/SynthGL/wolfxl-oss/blob/206bf4e1cdc62c8121569066362c1521d9b94a72/benchmarks/results/2026-10-02-community-2.0.8-vs-openpyxl-3.1.5.md)
for raw samples, environment identities, warmup policy, and memory measurements.
Older runs remain under [benchmarks/results/](https://github.com/SynthGL/wolfxl-oss/tree/206bf4e1cdc62c8121569066362c1521d9b94a72/benchmarks/results/).

## Fidelity

The [round-trip fidelity harness](https://github.com/SynthGL/wolfxl-oss/blob/206bf4e1cdc62c8121569066362c1521d9b94a72/fidelity-harness/README.md) compares workbook
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

## Related libraries

The same MIT-licensed family covers the other Office formats:

- [WolfPPT](https://github.com/wolfiesch/wolfppt-oss): edit PowerPoint `.pptx`
  decks, a python-pptx alternative (`pip install wolfppt`).
- [WolfDocx](https://github.com/wolfiesch/wolfdocx-oss): edit Word `.docx`
  documents, a python-docx alternative (`pip install wolfdocx`).

## Development

Prerequisites: a supported CPython, Rust, and `maturin`.

```bash
python -m pip install maturin pytest defusedxml openpyxl Pillow
maturin develop
pytest tests/test_community_distribution.py -q
```

The distribution-boundary test verifies the Community version, compiled
backends, and absence of Commercial-only Python modules.

See the [contribution guidelines](https://github.com/SynthGL/wolfxl-oss/blob/206bf4e1cdc62c8121569066362c1521d9b94a72/CONTRIBUTING.md) and
[security policy](https://github.com/SynthGL/wolfxl-oss/blob/206bf4e1cdc62c8121569066362c1521d9b94a72/SECURITY.md) for how to contribute or report a vulnerability.

## License

WolfXL Community is available under the [MIT License](LICENSE).
