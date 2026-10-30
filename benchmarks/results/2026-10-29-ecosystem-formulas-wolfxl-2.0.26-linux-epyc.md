# Python Excel ecosystem read benchmark: WolfXL 2.0.26

AMD EPYC 9655, 16 virtual CPUs | Linux x86_64 | Python 3.13.15 | five measured rounds after one warmup

[Raw samples and provenance](2026-10-29-ecosystem-formulas-wolfxl-2.0.26-linux-epyc.json)

WolfXL 2.0.26 reads the multi-sheet formula fixture in 0.747 s, 1.89x faster
than WolfXL 2.0.25 on the same machine (1.413 s). It now takes 1.33x
fastexcel's time on that fixture, down from 2.38x in the
[2.0.25 run](2026-10-29-ecosystem-formulas-linux-epyc.md), and ranks third
behind Polars and fastexcel, ahead of python-calamine and DuckDB. Against
openpyxl 3.1.5, WolfXL is 8.0x faster on formulas and 11.3x faster on plain
values, where it takes 1.02x fastexcel's time.[^scope][^identities]

| WolfXL on this machine | Plain values | Multi-sheet formula fixture |
|---|---:|---:|
| 2.0.25 (same-node control) | 0.4905 s | 1.4126 s |
| 2.0.26 | 0.4751 s | 0.7465 s |

The plain-value case is unchanged within run-to-run variation; the
formula-case gain comes from resolving cached formula results in the native
reader. The control is in
[its own report](2026-10-29-ecosystem-formulas-wolfxl-2.0.25-control-linux-epyc.md).[^control]

## read_values_large

| engine | scope | median s | units/s |
|---|---|---|---|
| wolfxl | full read/write | 0.4751 | 3,368,060 |
| openpyxl | full read/write | 5.3864 | 297,044 |
| python_calamine | read-only, Python values | 0.6669 | 2,399,094 |
| fastexcel | read-only, Arrow tables | 0.4662 | 3,431,896 |
| pylightxl | pure-Python read/write | 10.3699 | 154,292 |
| pandas_openpyxl | DataFrame I/O, openpyxl engine | 6.9965 | 228,686 |
| pandas_calamine | DataFrame read, calamine engine | 1.3938 | 1,147,918 |
| polars | DataFrame I/O, wraps XlsxWriter/fastexcel | 0.4768 | 3,355,838 |
| duckdb | SQL engine; DataFrame in, Arrow out | 0.5654 | 2,829,753 |
| tablib | Dataset wrapper, openpyxl backend | 7.2526 | 220,610 |
| pyexcel | wrapper, openpyxl backend | 17.3718 | 92,104 |
| xlsx2csv | read-only, transcodes to CSV text | 3.8811 | 412,253 |

## read_formulas_multisheet

| engine | scope | median s | units/s |
|---|---|---|---|
| wolfxl | full read/write | 0.7465 | 2,143,307 |
| openpyxl | full read/write | 5.9699 | 268,009 |
| python_calamine | read-only, Python values | 0.8547 | 1,872,009 |
| fastexcel | read-only, Arrow tables | 0.5605 | 2,854,720 |
| pylightxl | pure-Python read/write | 15.0993 | 105,965 |
| pandas_openpyxl | DataFrame I/O, openpyxl engine | 7.3522 | 217,622 |
| pandas_calamine | DataFrame read, calamine engine | 1.6207 | 987,225 |
| polars | DataFrame I/O, wraps XlsxWriter/fastexcel | 0.5568 | 2,873,568 |
| duckdb | SQL engine; DataFrame in, Arrow out | 0.8907 | 1,796,249 |
| tablib | Dataset wrapper, openpyxl backend | 7.9225 | 201,958 |
| pyexcel | wrapper, openpyxl backend | 17.9994 | 88,892 |
| xlsx2csv | read-only, transcodes to CSV text | 4.5544 | 351,308 |

## Workloads and measurement boundary

The workloads, fixture definitions, validation, and engine options are the same as in the
[2.0.25 run](2026-10-29-ecosystem-formulas-linux-epyc.md#workloads-and-measurement-boundary):
a 200,000-row by 8-column openpyxl-written sheet, and an XlsxWriter-written
workbook of four sheets x 50,000 rows x 8 columns with 1,000,000 cached
formula cells among 1,600,000 cells. WolfXL and openpyxl use
`read_only=True, data_only=True`. Every round reopens the workbook.[^validation]

## Reproduce

Run from the repository root with Python 3.13.15:

```bash
uv venv -p 3.13.15 .bench
uv pip install -p .bench/bin/python \
  wolfxl==2.0.26 openpyxl==3.1.5 XlsxWriter==3.2.9 PyExcelerate==0.13.0 \
  pylightxl==1.61 pandas==3.0.5 polars==1.43.2 python-calamine==0.8.2 \
  fastexcel==0.20.2 pyarrow==25.0.1 duckdb==1.5.5 tablib==3.10.0 \
  pyexcel==0.7.6 pyexcel-xlsx==0.6.1 xlsx2csv==0.8.6
.bench/bin/python benchmarks/benchmark_python_excel_ecosystem.py \
  --case read_values_large,read_formulas_multisheet \
  --rounds 5 --round-budget 240 --no-memory \
  --output-dir benchmarks/results --prefix 2026-10-29-ecosystem-formulas-wolfxl-2.0.26-linux-epyc
```

## Provenance

- Harness source: `37c5416a67b6690cfc6c8e7872a69e28b57b310f`.
- Harness SHA-256: `249842dcb9839ffc2146c1d9c09b89ccb82143b28e2c44e287558fa0baf274eb`.
  The measured code matches the 2.0.25 run; only comments and the module
  docstring differ.
- WolfXL: published PyPI 2.0.26 wheel, CPython 3.13, manylinux x86_64.
  Wheel SHA-256: `223ff2e4b0c68279894179d8c138e81289e2c759d0eac9ee717abb4ce8036640`.
  The installed native library matches the downloaded wheel, whose digest
  matches PyPI. No working-tree library build was used.[^build]
- Competitor versions match the 2.0.25 run, including fastexcel 0.20.2.
- Operation measurements ran sequentially on a dedicated VM, in fixed engine
  order, with one discarded warmup per engine and case. The OS page cache was
  not flushed. Cold-start latency was not measured.
- Per-round budget: 240 seconds, including warmup. Overall process timeout:
  3,000 seconds. Every engine completed; no discarded failures or timeouts.
- No memory pass was run; `metadata.provenance.memory_pass` records that it
  was skipped.
- Run: October 29, 2026, 7:53 PM to 8:06 PM PDT, including the control.
  UTC timestamps, fixture hashes, package/native identities, Python executable
  hash, and raw-output hash are in `metadata.provenance`, added after timing.
  The original sample, median, and comparison fields are unchanged.
- DuckDB emitted a `fetch_arrow_table()` deprecation warning. No measurement
  failed because of it.

[^scope]: These are warmed read-and-consume workload measurements, not a
    universal ranking or a formula-calculation benchmark. The two fixtures
    differ in writer, sheet count, string mix, and checksum work, so their
    timing difference does not isolate formula parsing cost. All medians and
    native output representations are retained.
[^control]: The 2.0.25 control ran on the same VM right after this run, in a
    separate environment with the same pinned packages, measuring WolfXL only.
    The harness regenerated its fixtures, which differ from this run's only by
    document timestamps. The 2.0.25 run linked above used a different VM of the
    same type; openpyxl and fastexcel took 8-10% longer on this one, so compare
    WolfXL versions with the same-node control, not across reports.
[^validation]: Cell count plus the column H sum catches omitted sheets and
    missing numeric caches in that column on this fixture. It does not prove
    equality of all values, string formula results, coordinates, row order, or
    Python types, or distinguish cached results from recomputation.
[^build]: Published compiler flags were not independently verified. Native
    hashes identify the measured distributions; this run did not rebuild them.
[^identities]: DuckDB's dynamically installed `excel` extension version and
    digest were not captured, so its rows are not fully reproducible. pandas,
    pylightxl, Tablib, pyexcel, pyexcel-xlsx, xlsx2csv, and PyExcelerate have
    version records only. Fixture hashes identify the measured files;
    regenerated ZIP bytes may differ because of document timestamps.
