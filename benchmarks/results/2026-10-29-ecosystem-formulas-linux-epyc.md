# Python Excel ecosystem read benchmark

AMD EPYC 9655, 16 virtual CPUs | Linux x86_64 | Python 3.13.15 | five measured rounds after one warmup

[Raw samples and provenance](2026-10-29-ecosystem-formulas-linux-epyc.json)

WolfXL 2.0.25 takes 1.05x fastexcel's time on the plain-value fixture and
2.38x on the multi-sheet formula fixture. Against openpyxl 3.1.5, WolfXL
is 11.1x faster on plain values and 4.5x on the formula fixture. WolfXL ranks
fifth on the formula fixture, behind Polars, fastexcel, python-calamine, and
DuckDB.[^scope][^identities]

## read_values_large

| engine | scope | median s | units/s |
|---|---|---|---|
| wolfxl | full read/write | 0.4436 | 3,606,750 |
| openpyxl | full read/write | 4.9284 | 324,646 |
| python_calamine | read-only, Python values | 0.6066 | 2,637,497 |
| fastexcel | read-only, Arrow tables | 0.4234 | 3,779,196 |
| pylightxl | pure-Python read/write | 9.6574 | 165,677 |
| pandas_openpyxl | DataFrame I/O, openpyxl engine | 6.1713 | 259,265 |
| pandas_calamine | DataFrame read, calamine engine | 1.2808 | 1,249,228 |
| polars | DataFrame I/O, wraps XlsxWriter/fastexcel | 0.4313 | 3,709,919 |
| duckdb | SQL engine; DataFrame in, Arrow out | 0.5163 | 3,098,999 |
| tablib | Dataset wrapper, openpyxl backend | 6.6572 | 240,340 |
| pyexcel | wrapper, openpyxl backend | 16.5547 | 96,649 |
| xlsx2csv | read-only, transcodes to CSV text | 3.5985 | 444,627 |

## read_formulas_multisheet

| engine | scope | median s | units/s |
|---|---|---|---|
| wolfxl | full read/write | 1.2381 | 1,292,314 |
| openpyxl | full read/write | 5.5346 | 289,088 |
| python_calamine | read-only, Python values | 0.7866 | 2,034,083 |
| fastexcel | read-only, Arrow tables | 0.5199 | 3,077,364 |
| pylightxl | pure-Python read/write | 14.2230 | 112,494 |
| pandas_openpyxl | DataFrame I/O, openpyxl engine | 6.7539 | 236,899 |
| pandas_calamine | DataFrame read, calamine engine | 1.4645 | 1,092,553 |
| polars | DataFrame I/O, wraps XlsxWriter/fastexcel | 0.5152 | 3,105,490 |
| duckdb | SQL engine; DataFrame in, Arrow out | 0.8152 | 1,962,806 |
| tablib | Dataset wrapper, openpyxl backend | 7.3682 | 217,149 |
| pyexcel | wrapper, openpyxl backend | 17.3534 | 92,201 |
| xlsx2csv | read-only, transcodes to CSV text | 4.1146 | 388,859 |

## Workloads and measurement boundary

| Case | Fixture | Timed work |
|---|---|---|
| `read_values_large` | openpyxl-written, 200,000 rows x 8 columns, one sheet | Open workbook, read all values, count cells |
| `read_formulas_multisheet` | XlsxWriter-written, four sheets x 50,000 rows x 8 columns | Open workbook, read every sheet, count cells, sum numeric cached results in column H |

Five columns per formula-fixture row contain formulas: three numeric and two
string results, all explicitly cached by XlsxWriter. This is 1,000,000 formula
cells among 1,600,000 total cells. The harness requests cached formula values.[^validation]
The checksum covers column H only: for globally unique row index `r`, its
cached result is `6*r+1`; the expected sum is `3*n*(n+1)+n`.

WolfXL and openpyxl use `read_only=True, data_only=True`. Every engine reads
every sheet. DuckDB receives the fixture's sheet names because its reader
cannot enumerate them. Arrow/DataFrame readers retain their native return
representation; their checksums use column reductions. Row readers accumulate
one Python numeric addition per row. xlsx2csv writes temporary CSV files and
reads them back to count cells and sum the last field.[^validation]
The units/s column counts all cells, not only formulas. Every round reopens
the workbook; imports and DuckDB's module-global connection remain warm.

## Reproduce

Run from the repository root with Python 3.13.15:

```bash
uv venv -p 3.13.15 .bench
uv pip install -p .bench/bin/python \
  wolfxl==2.0.25 openpyxl==3.1.5 XlsxWriter==3.2.9 PyExcelerate==0.13.0 \
  pylightxl==1.61 pandas==3.0.5 polars==1.43.2 python-calamine==0.8.2 \
  fastexcel==0.20.2 pyarrow==25.0.1 duckdb==1.5.5 tablib==3.10.0 \
  pyexcel==0.7.6 pyexcel-xlsx==0.6.1 xlsx2csv==0.8.6
.bench/bin/python benchmarks/benchmark_python_excel_ecosystem.py \
  --case read_values_large,read_formulas_multisheet \
  --rounds 5 --round-budget 240 --no-memory \
  --output-dir benchmarks/results --prefix 2026-10-29-ecosystem-formulas-linux-epyc
```

Use `--formula-sheets` and `--formula-rows` to change the sheet count and rows
per sheet. The default is four sheets of 50,000 rows each. Fixture generation
happens before timing. Keep result files from changed scenarios separate.

## Provenance

- Harness source: `65a7a07c36d630095f34d52d88acc5bcf87fc2dd`.
- Harness SHA-256: `078459135d4fa2bab7d3bc28c36d95629b111e2d6cb8e16d5a345a1e90f74e84`.
- WolfXL: published PyPI 2.0.25 wheel, CPython 3.13, manylinux x86_64.
  Wheel SHA-256: `30813f5e01e992245141af85647e6e5b07cec9b871ef220baad09c899f4bbfa1`.
  The installed native library matches the downloaded wheel, whose digest
  matches PyPI. No working-tree library build was used.[^build]
- Direct competitor versions match the October 27 ecosystem run, including
  fastexcel 0.20.2. This is not a measurement of fastexcel 0.21.0.
- Operation measurements ran sequentially on a dedicated VM, in fixed engine
  order, with one discarded warmup per engine and case. The OS page cache was
  not flushed. Cold-start latency was not measured.
- Per-round budget: 240 seconds, including warmup. Overall process timeout:
  3,000 seconds. Every engine completed; no discarded failures or timeouts.
- No memory pass was run. The raw harness metadata describes the optional
  default memory pass; `metadata.provenance.memory_pass` records that it was
  skipped for this run.
- Run: October 29, 2026, 3:48 PM to 4:00 PM PDT. UTC timestamps, fixture
  hashes, package/native identities, Python executable hash, and raw-output
  hash are in `metadata.provenance`, added after timing. The original sample,
  median, and comparison fields are unchanged.
- DuckDB emitted a `fetch_arrow_table()` deprecation warning. No measurement
  failed because of it.
- Only comments and the module docstring changed in the harness after the
  recorded source revision; the measured implementation is unchanged.

[^scope]: These are warmed read-and-consume workload measurements, not a
    universal ranking, a before/after regression claim, or a formula-calculation
    benchmark. The two fixtures also differ in writer, sheet count, string mix,
    and checksum work, so their timing difference does not isolate formula
    parsing cost. All medians and native output representations are retained.
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
