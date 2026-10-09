# Sparse styled read-only investigation — 2026-10-09

The 15 performance-stack PRs are merged. Community main is
`7b796d4dcc30612c2c5fbc860ca53747f55316aa`; Commercial main is
`ab47f6cd0370c17c2c3c683f4dff81d42cde8452`. Their trees match the final,
previously tested stack heads exactly.

## Cause

The frozen v2 sparse fixture has 25,000 rows × 32 columns (800,000 iterator
positions) but only 7,816 stored values. WolfXL creates a full StreamingCell at
every missing position and resolves its default style repeatedly.
Openpyxl 3.1.5 pads absent positions with its EMPTY_CELL singleton.

A cProfile run over the merged Community Python source recorded 800,000 full
proxy constructions and 3,200,000 _resolved_style calls. Its row-reading native
calls took 0.014 seconds of the 8.402-second instrumented run. Instrumented
times are diagnostic and are excluded from the speedup calculation.

## Targeted change

Absent XML cells with a non-gradient source style zero use a StreamingCell
subclass that initializes only worksheet and coordinates. Value/type and
immutable default font/fill/alignment are shared. Number format still observes
workbook close, and border access still resolves current source/merge state.
Authored empty cells retain their full style-aware proxy. A gradient at source
style zero takes the original path, preserving fresh mutable gradient copies.
Values-only iteration is unchanged.

Retaining coordinates is intentional: existing WolfXL tests require missing
cells to have coordinates, a parent, default style objects, and mutation
refusal. Replacing these cells with an openpyxl-style singleton would change
that contract.

## Local diagnostic evidence

Five fresh processes per engine/workload, one warmup, Python 3.12.14 and
openpyxl 3.1.5. Identical public fill/font/number-format loop, load + scan +
close median. Fixture hashes, all samples, signatures, module/native hashes,
and separate profile counters are in
[the raw receipt](sparse-styled-local-overlay-20261009.json).

| Workload | Before | Candidate | openpyxl | Candidate gain | Candidate / openpyxl |
|---|---:|---:|---:|---:|---:|
| Sparse 800k positions / 7,816 values | 1.934145 s | 0.496109 s | 0.158776 s | 3.899× | 3.125× slower |
| Dense 100k stored cells | 0.415763 s | 0.418819 s | 0.766113 s | 0.993× | 1.829× faster |

All signatures match: sparse 25,000 rows, first-column checksum 19,532,811,
3,127 styled cells. The candidate's profile has 7,816 full proxies, 792,184
specialized blanks and 31,264 style-resolution calls. Sparse wall time falls
74.4%; the dense result is 0.7% slower and does not support a dense speedup claim.

**Boundary:** the baseline Python package matches all 267 tracked Community
main package blobs; its 266 .py files match exactly. Candidate changes only
_streaming.py. Both overlays use the same published Community 2.0.9 native
extension (SHA-256 recorded in the receipt), not a newly compiled merged-main
extension. This diagnoses the Python change; it is not an updated release
benchmark or a measured Commercial result.

Local focused validation: lint passes; 67 tests pass. Two existing merged-border
tests and the new pending-merge blank-border test fail equally on baseline and
candidate with this older extension, which lacks the current endpoint-style
methods. The source-built workflow runs all 70 targeted tests without those
deselections and produces a separate benchmark artifact for each edition.

## Reproduction and source-built evidence

Build/install the ordinary wheel, pin openpyxl 3.1.5, extract the base commit's
python/wolfxl/_streaming.py, then run:

```bash
python benchmarks/sparse_styled_profile.py \
  --baseline-streaming /tmp/baseline-streaming.py \
  --baseline-ref BASE_SHA --candidate-ref CANDIDATE_SHA \
  --native-origin "standard source-built wheel; native sources unchanged" \
  --output /tmp/sparse-evidence/results.json
```

The `Sparse styled Python overlay evidence` workflow builds the standard
candidate wheel and verifies that native sources are unchanged from the base.
It runs the targeted compatibility tests and the five-trial paired comparison,
then uploads JSON, profile files, source references and changed-file receipts.
Both overlays hold every module except _streaming.py and the native binary
constant. This isolates the Python change, rather than claiming two independent
native builds. These new results must stay separately labelled; the existing
README performance receipts are unchanged.

## Completed source-built validation

The Community source-built [workflow](https://github.com/SynthGL/wolfxl-oss/actions/runs/37885386977)
passed all 70 targeted tests, including the three tests unsupported by the older
published native extension used in the local diagnostic. Normal product CI is
also green. PR #48 is merged.

The separately labelled [source-built raw receipt](sparse-styled-source-built-20261009.json)
records sparse medians 3.279210 → 0.840991 seconds (3.899×), versus openpyxl
0.294885 seconds (candidate still 2.852× slower). Dense medians are 0.538393 →
0.534949 seconds (1.006×). All signatures match. CPython 3.12.15, openpyxl 3.1.5,
Ubuntu 24.04 GitHub-hosted runner; standard release wheel with unchanged native
sources. The baseline Python fingerprint matches the local exact-main source
fingerprint, and the candidate _streaming.py hash is identical in both receipts.
No cross-run or cross-machine cumulative ratio is inferred.
