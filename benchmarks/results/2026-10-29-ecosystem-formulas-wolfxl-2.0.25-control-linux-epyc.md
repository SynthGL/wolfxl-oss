# Python Excel ecosystem benchmark

AMD EPYC 9655 96-Core Processor | Python 3.13.15 | median of 5 rounds

## read_values_large

| engine | scope | median s | units/s |
|---|---|---|---|
| wolfxl | full read/write | 0.4905 | 3,261,785 |

## read_formulas_multisheet

| engine | scope | median s | units/s |
|---|---|---|---|
| wolfxl | full read/write | 1.4126 | 1,132,666 |


## Same-node WolfXL 2.0.25 control

This is a control, not a competitor comparison. It ran on the same machine
right after the [WolfXL 2.0.26 run](2026-10-29-ecosystem-formulas-wolfxl-2.0.26-linux-epyc.md),
in a separate environment with the same pinned packages; only the WolfXL
distribution differs. The filter measures WolfXL only, so there is no repeated
openpyxl timing to quantify machine drift between the two runs.

- Run label: 2026-10-29 (PDT); the raw UTC timestamp is on 2026-10-30.
- AMD EPYC 9655 96-Core Processor, 16 virtual CPUs; Python 3.13.15.
- Harness source: `37c5416a67b6690cfc6c8e7872a69e28b57b310f`.
- Published PyPI wheel SHA-256:
  `30813f5e01e992245141af85647e6e5b07cec9b871ef220baad09c899f4bbfa1`.
  The installed native library matches the downloaded wheel, and the wheel
  matches PyPI's published digest. No local WolfXL build was used.
- Five measured rounds after one discarded warmup per case; OS page cache
  not flushed. These are warmed-operation timings, not cold-start latency.
- Overall timeout: 900 seconds. Per-round budget: 240 seconds. No memory pass.
- The harness regenerated both fixtures with the same writers and sizes;
  their bytes differ from the paired run's only by document timestamps.

```bash
.control/bin/python benchmarks/benchmark_python_excel_ecosystem.py --case read_values_large,read_formulas_multisheet --engines wolfxl --rounds 5 --round-budget 240 --no-memory --output-dir benchmarks/results --prefix 2026-10-29-ecosystem-formulas-wolfxl-2.0.25-control-linux-epyc
```

The command uses repository-relative output paths. WolfXL wheel and native
library hashes, fixture hashes, cache state, and timeouts are embedded in the JSON.

`metadata.provenance` was added after timing, using identities captured on
the measured machine. All other JSON fields are unmodified harness output.
