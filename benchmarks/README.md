# Benchmarks

Reproducible benchmarks for WolfXL Community. Every number shown in the
repository README charts comes from a committed raw results file in
`results/`, and the charts are regenerated from that file, never edited by hand.

Two harnesses:

- `benchmark_openpyxl_vs_wolfxl.py`: deep openpyxl comparison (writes, reads,
  modify-and-save, styled rows, multi-sheet, peak memory).
- `benchmark_python_excel_ecosystem.py`: cross-library comparison against
  other open-source Python Excel libraries, each measured only inside its
  supported scope.

## What is measured

`benchmark_openpyxl_vs_wolfxl.py` runs matched workloads through the public
APIs of both libraries in the same process environment:

- plain, mixed-type, wide, formula, and styled row writes
- multi-sheet workbook writes
- full and read-only value reads, formula-text reads
- modify-two-cells-and-save on an existing workbook
- large-file variants (200,000 rows x 8 columns, about 1.6 million cells)
- a separate peak-RSS pass that runs each large case in a fresh process

Each case reports the median of the configured rounds. The results JSON embeds
the machine, Python, and package versions used for the run.

## Reproduce

```bash
python -m venv .bench && .bench/bin/pip install wolfxl==2.0.1 openpyxl==3.1.5
.bench/bin/python benchmarks/benchmark_openpyxl_vs_wolfxl.py \
    --rounds 5 \
    --output-dir /tmp/wolfxl-bench-results \
    --prefix my-run
```

Useful options: `--case <name>[,<name>...]` runs a subset, `--no-large` skips
the 200,000-row workloads, `--no-memory` skips the peak-RSS pass. Run with
`--help` for the full list.

## Ecosystem comparison

`benchmark_python_excel_ecosystem.py` compares, in one environment:

| Library | Scope |
|---|---|
| wolfxl | full read/write |
| openpyxl | full read/write (baseline) |
| XlsxWriter | write-only |
| PyExcelerate | write-only |
| pylightxl | pure-Python read/write |
| pandas | DataFrame I/O (openpyxl write engine, openpyxl and calamine read engines) |
| Polars | DataFrame I/O (wraps XlsxWriter for writes, fastexcel for reads) |
| DuckDB | SQL engine (excel extension: COPY xlsx / read_xlsx to Arrow) |
| Tablib | Dataset wrapper (openpyxl xlsx backend) |
| pyexcel | wrapper (openpyxl backend via pyexcel-xlsx) |
| python-calamine | read-only, returns Python values |
| fastexcel | read-only, returns Arrow tables |
| xlsx2csv | read-only, transcodes to CSV text |

Inclusion bar: xlsx-capable, no external application required, and broad real
adoption (roughly one million PyPI downloads per month). Excluded for scope:
xlrd/xlwt (.xls only), pyxlsb (.xlsb only), xlwings (drives a running Excel
installation), and commercial SDKs.

Excluded pending adoption: excelize-py (the Python binding of Go's Excelize;
pre-1.0 and far below the download bar as of this run).

To steelman the baselines, three mode variants are measured on the large
plain write and the memory pass only: openpyxl in write_only mode, XlsxWriter
in constant_memory mode, and pandas with the xlsxwriter engine. They exist
because mode choice matters on exactly that workload; adding them to every
chart would only duplicate bars.

Cases: large plain write, mixed-type write, unique-strings write
(shared-strings stress), large value read, multi-sheet cached-formula read,
and a peak-RSS pass for the large plain cases. Write-only libraries appear
only in write cases and read-only libraries only in read cases. pandas,
Polars, and DuckDB are timed from a prebuilt DataFrame (DuckDB reads a
registered DataFrame view) and Tablib from a prebuilt Dataset, so their write
numbers include frame-to-worksheet conversion but not frame construction.
The large value-read fixture is written by openpyxl. pylightxl has no datetime
support and is excluded from the mixed-type write case.

The formula-read fixture is written by XlsxWriter: four sheets of 50,000 rows
and eight columns each, with numeric and string cached results in five formula
columns. All readers load every sheet. Each timed round includes a cell count
and a sum of the final numeric formula column. Arrow/DataFrame engines use
column reductions; row readers sum in Python. This is cached-value reading,
not formula evaluation or full value/type parity validation. Use
`--formula-sheets` and `--formula-rows` to change this workload.

Every run (including the discarded warmup) must finish one round within
`--round-budget` seconds (default 240). An engine that exceeds the budget is
recorded as DNF in the raw results and listed in the chart footnotes instead
of being silently dropped.

```bash
.bench/bin/pip install wolfxl==2.0.26 openpyxl==3.1.5 xlsxwriter pyexcelerate \
    python-calamine fastexcel pyarrow pylightxl pandas polars duckdb \
    tablib pyexcel pyexcel-xlsx xlsx2csv
.bench/bin/python benchmarks/benchmark_python_excel_ecosystem.py \
    --rounds 5 --output-dir /tmp/ecosystem-results --prefix my-run
```

## Current measured results

- [2026-10-29 plain and multi-sheet formula reads, WolfXL 2.0.26](results/2026-10-29-ecosystem-formulas-wolfxl-2.0.26-linux-epyc.md)
  ([raw JSON](results/2026-10-29-ecosystem-formulas-wolfxl-2.0.26-linux-epyc.json)):
  the same cases, pins, and fixture shapes as the 2.0.25 run below, with five
  measured rounds after one warmup. No write or memory measurements.
- [Same-node WolfXL 2.0.25 control](results/2026-10-29-ecosystem-formulas-wolfxl-2.0.25-control-linux-epyc.md)
  ([raw JSON](results/2026-10-29-ecosystem-formulas-wolfxl-2.0.25-control-linux-epyc.json)):
  the same two read cases, restricted to WolfXL, on the 2.0.26 machine.
- [2026-10-29 plain and multi-sheet formula reads, WolfXL 2.0.25](results/2026-10-29-ecosystem-formulas-linux-epyc.md)
  ([raw JSON](results/2026-10-29-ecosystem-formulas-linux-epyc.json)): WolfXL
  2.0.25 with the October 27 competitor pins, including fastexcel 0.20.2,
  on a different VM of the same type.
- [2026-10-27 ecosystem comparison](results/2026-10-27-ecosystem-full-linux-epyc.md)
  ([raw JSON](results/2026-10-27-ecosystem-full-linux-epyc.json)): WolfXL
  2.0.21 and the same direct competitor pins and engine set as the
  [2026-08-18 run](results/2026-08-18-ecosystem-full-linux-epyc.json).
- [Same-node WolfXL 2.0.1 control](results/2026-10-27-ecosystem-wolfxl-2.0.1-control-linux-epyc.md)
  ([raw JSON](results/2026-10-27-ecosystem-wolfxl-2.0.1-control-linux-epyc.json)):
  the same ecosystem cases, restricted to WolfXL, on the rerun machine.
- [WolfXL 2.0.21 modify-and-save comparison](results/2026-10-27-community-2.0.21-modify-linux-epyc.md)
  ([raw JSON](results/2026-10-27-community-2.0.21-modify-linux-epyc.json)):
  the existing deep harness, with five measured rounds for both sizes.
  Use its `modify_save` phase for operation claims; total time also includes
  copying the fixture and reopening the result with openpyxl for verification.

The October 27 ecosystem rerun uses the default 240-second per-round budget,
not the August run's 43,200-second budget. Every original engine is included;
cases that exceed the new limit are recorded as DNF rather than substituted
with older measurements. Ecosystem timings are warmed measurements after one
discarded round, not cold-start latency. In runs with a peak-RSS pass, memory
is measured once per engine in a fresh process and includes the harness's input
data. Hardware, package identities, commands, timeouts, and cache conditions
are recorded alongside the raw samples. Historical charts are unchanged.

## Regenerate the charts

```bash
python benchmarks/render_charts.py \
    benchmarks/results/2026-08-18-community-2.0.1-vs-openpyxl-3.1.5.json \
    assets/benchmarks
python benchmarks/render_ecosystem_charts.py \
    benchmarks/results/2026-08-19-ecosystem-full-linux-epyc-merged.json \
    assets/benchmarks
```

Both renderers are deterministic and stdlib-only: the same input JSON always
produces byte-identical SVGs.

## Incremental runs

Engines are measured in isolation, so a new engine or variant does not
require rerunning the whole suite. On the *same machine, Python, and pinned
library versions* as the committed run:

```bash
.bench/bin/python benchmarks/benchmark_python_excel_ecosystem.py \
    --rounds 5 --engines new_engine --output-dir /tmp/results --prefix incr
python benchmarks/merge_ecosystem_results.py \
    benchmarks/results/<base>.json /tmp/results/incr.json \
    benchmarks/results/<merged>.json
```

`merge_ecosystem_results.py` refuses to merge unless CPU, machine, platform,
Python, round count, round budget, and every shared library version match
exactly, and it records each increment's source and timestamp in the merged
metadata. If the environment changed, rerun the full suite instead; with
pylightxl's numbers already on record, a full rerun without it takes well
under an hour.

## Reading the results honestly

- Speedups vary by workload. Large streaming reads and writes are where the
  Rust backend helps most; small workbooks see smaller wins.
- The styled-rows case compares `ws.write_styled_rows()` against per-cell
  styling in openpyxl, which is an API-shape difference as well as an engine
  difference. The plain write cases use the same `append()`/row-iteration
  shapes in both libraries.
- On small workloads, peak memory is roughly equal between the two libraries;
  the memory chart shows the large-file cases where they diverge.
- In the ecosystem read case, the Arrow-native engines (fastexcel, Polars,
  DuckDB) avoid materializing Python cell objects; wolfxl and python-calamine
  return Python cell values, and only wolfxl, openpyxl, pylightxl, Tablib, and
  pyexcel can also write.
- In both runs, DuckDB's excel extension wins the small mixed-type write
  outright. It is timed from a registered DataFrame, so DataFrame-construction
  cost is excluded from its number.
- pylightxl's writer scales quadratically with cell count. In the August run,
  the 12-hour budget allowed every round to finish (241 s plain, 1,438 s
  unique strings), and the chart axis was clipped at the second-slowest engine
  so the remaining bars stayed readable. The printed value is always the
  true measurement; the October rerun applies the shorter budget above.
- PyExcelerate, Tablib, and pyexcel serialize floats with fewer significant
  digits than the other writers (a value like `8/7` does not round-trip
  bit-exactly), so their write numbers come with a precision tradeoff.
- Engines that exceed the per-round budget are recorded as DNF, not omitted.
- Numbers are from the pinned versions, hardware, and OS recorded in the
  results JSON. Different machines will produce different absolute times.

## Bounded reader admission

`reader_limits.py` runs each workbook in a fresh Linux interpreter with GNU
`/usr/bin/time`, a 5-second timeout and a 2 GiB address-space ceiling. It records
peak RSS in bytes, wall seconds, fixture SHA-256, normalized per-sheet value
hashes and errors. String normalization hashes each Python string identity
once so repeated shared strings do not dominate the measurement.

Use the `.xlsx` files from Apache POI's `test-data/spreadsheet` directory at
revision `6327f056ba3cfdae326a0261a236460592b80f5c` in a `corpus` directory.
That revision has 352 XLSX files and no `zip-bomb.xlsx`. Generate the substitute
1 GiB high-ratio part and 200,002-member ZIP locally; do not commit binaries:

```bash
python benchmarks/reader_limits.py --corpus-dir corpus --generate-cases
python benchmarks/reader_limits.py --corpus-dir corpus \
    --source HEAD --output reader-fixtures.json
python benchmarks/reader_limits.py --corpus-dir corpus --corpus \
    --source HEAD --output reader-corpus.json
```

Run the same committed harness in an isolated environment with `wolfxl==2.0.35`
for the released baseline (`--source pypi-wolfxl-2.0.35`). Install the current
checkout with `maturin develop --release` for the candidate. Keep the exact
source SHA, build configuration, Python version, native-library hash, fixture
hashes, iteration count and cache state with the raw results. One resource
sample is a bounded-load observation, not a throughput or speedup benchmark.
