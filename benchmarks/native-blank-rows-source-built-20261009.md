# Native blank-row batching — Community, 2026-10-09

Build implicit blank rows in bounded native batches while retaining the existing Python `StreamingBlankCell` type. Python object allocation and member descriptors preserve the same three initialized coordinate slots; the implementation uses safe PyO3 APIs without object-layout casts or global worksheet references.

The iterator batches at most 128 rows and 4,096 cells. Populated rows, authored empty cells, source-style-zero gradients, very wide ranges, and values-only reads keep their existing paths. Callers may retain cells and obtain independent, stable coordinate snapshots.

## Source-built evidence

Baseline `0a7056e713e2678585642aab7fb9f84de643d120` → candidate `7a2338f96d71e12bd8d33032c40bfc4c4ce50e6b`. This report is a documentation-only follow-up to the measured candidate; runtime and build inputs are unchanged.

[Workflow](https://github.com/SynthGL/wolfxl-oss/actions/runs/37941604086) builds separate baseline/candidate wheels with repository-default maturin features and the standard Cargo release profile. Ubuntu 24.04, CPython 3.12.15, openpyxl 3.1.5, rustc 1.99.0. Five fresh-process trials per engine/workload follow one separate warmup; serial engine order is deterministically shuffled.

| Workload | Baseline (s) | Candidate (s) | openpyxl (s) | Baseline / candidate |
|---|---:|---:|---:|---:|
| sparse styles | 0.830955 | 0.527006 | 0.290952 | 1.577× |
| sparse rows | 0.580419 | 0.278712 | 0.097891 | 2.083× |
| dense styles | 0.523456 | 0.526192 | 1.336864 | 0.995× |
| dense rows | 0.274978 | 0.275144 | 1.140120 | 0.999× |

Sparse styled scan time falls 36.6%, but the candidate remains **1.811× slower than openpyxl**. Dense styled elapsed time is 0.5% higher with overlapping sample ranges; no dense speedup is claimed. Row-only timings are diagnostic and are not substituted for the public styled scan.

The sparse fixture has 25,000 × 32 positions and 7,816 stored values; dense has 20,000 × 5 stored values. Both phases include load and close. Styled scans use the same public fill/font/number-format loop as the existing frozen benchmark, checking row count, first-column checksum, and styled-cell count. Row-only scans check row and position counts. These signatures do not prove every style or value; focused compatibility tests cover the changed contract separately.

## Validation and limits

79 focused streaming tests pass. New checks cover retained snapshots across batch boundaries, clipping, unique cell identities, bounded batch sizes, worksheet lifetime, invalid native requests, zero-width requests, and fallback. Existing checks cover authored empty styles, gradients, merged borders, shallow copies, mutation refusal, and close behavior. Normal Community Python/Rust/lint CI is green on the measured source.

Construction still allocates one coordinate-aware object per absent position. Openpyxl uses its empty-cell singleton and returns less default-style information. Further representation/accessor work is tracked in [Commercial issue #757](https://github.com/SynthGL/wolfxl/issues/757).

## Raw samples and reproduction

[Unmodified source-built JSON](native-blank-rows-source-built-20261009.json) records all samples, fixture hashes, Python/native hashes, source references, and signatures. Different runners and the prior October 8/9 receipts are not multiplied into a cumulative ratio. These are unreleased source builds, not new registry artifacts.

Extract each standard wheel into its own directory, then run:

```bash
python benchmarks/native_blank_rows_profile.py \
  --baseline-package /tmp/baseline-package \
  --candidate-package /tmp/candidate-package \
  --baseline-ref BASE_SHA --candidate-ref CANDIDATE_SHA \
  --build-description "separate standard release wheels" \
  --output /tmp/native-blank-rows/results.json
```
