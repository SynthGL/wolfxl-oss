# Python Excel ecosystem benchmark

AMD EPYC 9655 96-Core Processor | Python 3.13.15 | median of 5 rounds

## write_plain_large

| engine | scope | median s | units/s |
|---|---|---|---|
| wolfxl | full read/write | 0.6659 | 2,402,834 |

## write_mixed

| engine | scope | median s | units/s |
|---|---|---|---|
| wolfxl | full read/write | 0.0345 | 1,448,290 |

## write_unique_strings

| engine | scope | median s | units/s |
|---|---|---|---|
| wolfxl | full read/write | 0.2914 | 1,715,610 |

## read_values_large

| engine | scope | median s | units/s |
|---|---|---|---|
| wolfxl | full read/write | 0.3915 | 4,086,551 |

## memory_write_large

| engine | peak RSS MiB |
|---|---|
| wolfxl | 644 |

## memory_read_large

| engine | peak RSS MiB |
|---|---|
| wolfxl | 186 |


## Same-node WolfXL 2.0.1 control

This is a control, not the current-release ecosystem comparison. It ran after
the 2.0.21 ecosystem and modify measurements on the same machine. Only the
WolfXL distribution was replaced; every other installed package and native
file hash remained identical. The filter measures WolfXL only, so there is
no repeated openpyxl timing to quantify machine drift between the runs.

- Run label: 2026-10-27; the raw UTC timestamp is on 2026-10-28.
- AMD EPYC 9655 96-Core Processor, 16 virtual CPUs; Python 3.13.15.
- Harness source: `405465b04e67f844ef1013e755e1f25f7d2d59bc`.
- Published PyPI wheel SHA-256:
  `a2f790b2dcf8de66f063d3197827c86fcb57f3817a9408d7c1977b9d4745147c`.
  The installed native library matches the downloaded wheel, and the wheel
  matches PyPI's published digest. No local WolfXL build was used.
- Five measured rounds after one discarded warmup per case; OS page cache
  not flushed. These are warmed-operation timings, not cold-start latency.
- Overall timeout: 600 seconds. Per-round and memory-process budget:
  240 seconds. Peak RSS is one fresh-process run per case, including imports
  and input data.
- Use the paired current-release file for competitor comparisons:
  [WolfXL 2.0.21](2026-10-27-ecosystem-full-linux-epyc.md). No inference
  about version-to-version modify speed is supported by this control.

```bash
.bench/bin/python benchmarks/benchmark_python_excel_ecosystem.py --rounds 5 --round-budget 240 --engines wolfxl --output-dir benchmarks/results --prefix 2026-10-27-ecosystem-wolfxl-2.0.1-control-linux-epyc
```

The command uses repository-relative output paths. Full dependency and binary
identities, CPU details, cache state, and timeouts are embedded in the JSON.

`metadata.provenance` was added after timing, using identities captured on
the measured machine. All other JSON fields are unmodified harness output.
