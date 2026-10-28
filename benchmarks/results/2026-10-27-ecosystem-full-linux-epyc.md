# Python Excel ecosystem benchmark

AMD EPYC 9655 96-Core Processor | Python 3.13.15 | median of 5 rounds

## write_plain_large

| engine | scope | median s | units/s |
|---|---|---|---|
| wolfxl | full read/write | 0.7449 | 2,148,061 |
| openpyxl | full read/write | 8.1507 | 196,303 |
| xlsxwriter | write-only | 4.6574 | 343,536 |
| pyexcelerate | write-only | 3.6882 | 433,818 |
| pylightxl | pure-Python read/write | DNF (> 240 s/round) | - |
| pandas_openpyxl | DataFrame I/O, openpyxl engine | 10.6173 | 150,697 |
| polars | DataFrame I/O, wraps XlsxWriter/fastexcel | 6.0663 | 263,753 |
| duckdb | SQL engine; DataFrame in, Arrow out | 0.9808 | 1,631,393 |
| tablib | Dataset wrapper, openpyxl backend | 11.2580 | 142,121 |
| pyexcel | wrapper, openpyxl backend | 6.4342 | 248,672 |

## write_mixed

| engine | scope | median s | units/s |
|---|---|---|---|
| wolfxl | full read/write | 0.0371 | 1,346,393 |
| openpyxl | full read/write | 0.2471 | 202,379 |
| xlsxwriter | write-only | 0.1394 | 358,779 |
| pyexcelerate | write-only | 0.1212 | 412,645 |
| pandas_openpyxl | DataFrame I/O, openpyxl engine | 0.3389 | 147,542 |
| polars | DataFrame I/O, wraps XlsxWriter/fastexcel | 0.1702 | 293,820 |
| duckdb | SQL engine; DataFrame in, Arrow out | 0.0256 | 1,950,557 |
| tablib | Dataset wrapper, openpyxl backend | 0.3399 | 147,103 |
| pyexcel | wrapper, openpyxl backend | 0.2492 | 200,616 |

## write_unique_strings

| engine | scope | median s | units/s |
|---|---|---|---|
| wolfxl | full read/write | 0.2590 | 1,930,758 |
| openpyxl | full read/write | 3.0970 | 161,448 |
| xlsxwriter | write-only | 2.1996 | 227,313 |
| pyexcelerate | write-only | 1.3333 | 375,011 |
| pylightxl | pure-Python read/write | DNF (> 240 s/round) | - |
| pandas_openpyxl | DataFrame I/O, openpyxl engine | 4.0814 | 122,506 |
| polars | DataFrame I/O, wraps XlsxWriter/fastexcel | 2.4578 | 203,435 |
| duckdb | SQL engine; DataFrame in, Arrow out | 0.3721 | 1,343,693 |
| tablib | Dataset wrapper, openpyxl backend | 4.0204 | 124,364 |
| pyexcel | wrapper, openpyxl backend | 2.5162 | 198,715 |

## read_values_large

| engine | scope | median s | units/s |
|---|---|---|---|
| wolfxl | full read/write | 0.4322 | 3,702,014 |
| openpyxl | full read/write | 4.7791 | 334,790 |
| python_calamine | read-only, Python values | 0.5796 | 2,760,390 |
| fastexcel | read-only, Arrow tables | 0.4140 | 3,864,597 |
| pylightxl | pure-Python read/write | 9.5057 | 168,320 |
| pandas_openpyxl | DataFrame I/O, openpyxl engine | 6.0673 | 263,708 |
| pandas_calamine | DataFrame read, calamine engine | 1.3412 | 1,192,978 |
| polars | DataFrame I/O, wraps XlsxWriter/fastexcel | 0.4108 | 3,894,637 |
| duckdb | SQL engine; DataFrame in, Arrow out | 0.5119 | 3,125,831 |
| tablib | Dataset wrapper, openpyxl backend | 6.3268 | 252,893 |
| pyexcel | wrapper, openpyxl backend | 15.9365 | 100,398 |
| xlsx2csv | read-only, transcodes to CSV text | 3.4524 | 463,452 |

## memory_write_large

| engine | peak RSS MiB |
|---|---|
| grid_baseline | 266 |
| wolfxl | 449 |
| openpyxl | 775 |
| xlsxwriter | 463 |
| pyexcelerate | 296 |
| pylightxl | DNF (> 240 s) |
| pandas_openpyxl | 900 |
| polars | 753 |
| duckdb | 370 |
| tablib | 866 |
| pyexcel | 295 |

## memory_read_large

| engine | peak RSS MiB |
|---|---|
| wolfxl | 186 |
| openpyxl | 200 |
| python_calamine | 319 |
| fastexcel | 299 |
| pylightxl | 1,520 |
| pandas_openpyxl | 325 |
| pandas_calamine | 422 |
| polars | 299 |
| duckdb | 253 |
| tablib | 311 |
| pyexcel | 981 |
| xlsx2csv | 184 |


## Reproduction and interpretation

- Run label: 2026-10-27. The raw timestamp is UTC and falls on 2026-10-28.
- CPU from `lscpu`: AMD EPYC 9655 96-Core Processor, 16 virtual CPUs.
- Python 3.13.15, Linux x86_64. Full OS, dependency versions, installed-file
  hashes, Python executable hash, and DuckDB extension identity are in the JSON.
- Harness source: `405465b04e67f844ef1013e755e1f25f7d2d59bc`.
- WolfXL is the published PyPI 2.0.21 wheel, not a working-tree build. Its wheel
  SHA-256 is `ae3c8dc65b26b895fbc3621f021f0cb18814d81ffa170c21714973ce786aaeb9`.
  The installed native library matches the downloaded wheel; the download
  matches PyPI's digest. Published compiler flags were not independently checked.
- Five measured rounds after one discarded warmup per case. The OS page cache
  was not flushed. These are warmed-operation timings, not cold-start latency.
  Each peak-RSS row is one fresh-process run and includes input data and imports.
- Overall process timeout: 3,000 seconds. Each timing round and each memory
  process has a 240-second budget, including warmups. The August run used
  43,200 seconds. pylightxl timed out for large plain writes, unique-string
  writes, and the large-write memory pass. DNF is not an observed regression.
- The engine filter preserves the August 18 engine set, not the August 19
  mode-expanded matrix. The same direct competitor package versions are pinned.
  Historical transitive packages and the OS were not reproduced, so use the
  same-node 2.0.1 control for version comparisons instead of August timings.
- DataFrame and Dataset inputs are prebuilt outside timed write operations.
  The read fixture is written by openpyxl. Arrow readers do not return the same
  Python-object representation as WolfXL and python-calamine.
- DuckDB was faster on the mixed-type write. fastexcel and Polars were slightly
  faster on the large read. WolfXL was fastest on large plain and unique-string
  writes in this matrix. Several writers used less peak RSS than WolfXL;
  xlsx2csv also used slightly less read RSS while returning CSV text.
- A DuckDB `fetch_arrow_table()` deprecation warning was emitted. No case
  failed for that warning.
- The same-node control is
  [WolfXL 2.0.1](2026-10-27-ecosystem-wolfxl-2.0.1-control-linux-epyc.md).

```bash
.bench/bin/python benchmarks/benchmark_python_excel_ecosystem.py --rounds 5 --round-budget 240 --engines wolfxl,openpyxl,xlsxwriter,pyexcelerate,pylightxl,pandas_openpyxl,pandas_calamine,polars,python_calamine,fastexcel,duckdb,tablib,pyexcel,xlsx2csv,grid_baseline --output-dir benchmarks/results --prefix 2026-10-27-ecosystem-full-linux-epyc
```

The command uses repository-relative output paths. Timing and memory workloads
ran sequentially with no concurrent benchmark jobs on the measured machine.

`metadata.provenance` was added after timing, using identities captured on
the measured machine. All other JSON fields are unmodified harness output.
