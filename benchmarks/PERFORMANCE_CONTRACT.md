# Styled-read and two-cell-edit benchmark contract v2

`performance_contract.py` is a new, separately labelled workload contract.
The original `benchmark_openpyxl_vs_wolfxl.py`, its fixtures and published receipts
remain unchanged. Its historical edit totals include an eager openpyxl verification
reload; these v2 totals use a bounded independent verification pass and cannot be
substituted into the historical table without relabelling the workload.

The default fixture sizes are 2,000 rows for small styled reads, 20,000 rows for
larger styled reads, and 200,000 rows for edits. Plain rows populate **five**
columns, matching the actual historic row builder, whose nominal column setting
was eight. Metadata reports both counts and populated value cells. Sparse sheets
place five cells every 16 rows plus a styled cell in column 32. High-cardinality
styles use 1,024 custom number formats by default. All sizes are bounded and can
be reduced for diagnostics. Generated workbooks stay outside Git.

| Workloads | What is measured |
|---|---|
| `edit_top`, `edit_middle`, `edit_bottom` | Existing value replacements at different XML offsets |
| `edit_missing` | Insertion beyond the existing worksheet bounds |
| `edit_cross_sheet` | One replacement on each of two large worksheets |
| `edit_formula` | Formula-to-value replacement while retaining another formula |
| `edit_styled`, `edit_merged` | Value edits with independently verified retained styles and merges |
| `styled_{small,large,sparse,high_cardinality,merged}_{eager,read_only}` | Identical public Cell/font/fill/number-format scan for both engines |

No bulk API replaces the Cell scan. Empty read-only cells are treated as unstyled,
allowing openpyxl's `EmptyCell` to participate in the same loop. ZIP timestamps,
core properties and member ordering are normalized. Fixtures are reused by
configuration hash, and altered cached fixtures fail instead of silently changing
the experiment. The installed openpyxl builder version is part of their identity.

Each measured edit records copy, load, assignment, save, close, engine operation,
bounded verification and total separately. Operation excludes copy and verification.
Total includes both. The same independent openpyxl read-only verifier uses one
bounded values iterator per edited sheet for both engines. A late-row verification
still streams the XML prefix; it is not constant-time random access.

After the timed rounds, a full eager openpyxl reopen compares all stored values,
data types, style meanings, sheet bounds and merges against an independently
edited reference. ZIP CRC and all XML members are validated separately. WolfXL
outputs must preserve untouched package members byte-for-byte; legitimate edited
worksheet and formula-cache/workbook metadata changes are allowed and reported.
Openpyxl's full-rewrite package changes are reported. Deep-validation time is
outside operation and total. This synthetic check is not Excel certification.

Runs alternate engine order, collect garbage before each sample, disable cyclic
GC during measured operations, and retain raw samples, ranges and warmup samples.
Receipts record Python/environment details, Git HEAD and dirty status, actual
import paths, the complete imported Python source hash, and native-library hash.
Unchanged native libraries can be reused for Python-only changes without implying
a native rebuild. Run performance stages serially on an otherwise quiet machine.

```bash
# Set PYTHONPATH only if measuring a source tree with its native extension present.
PYTHONPATH=/path/to/stage/python /path/to/env/bin/python \
  benchmarks/performance_contract.py --label baseline \
  --source-root /path/to/stage --fixture-dir /tmp/wolfxl-perf-v2-fixtures \
  --output /tmp/wolfxl-baseline.json --rounds 5 --warmups 1

# Repeat the exact command for each integrated stage; change only stage/env/label/output.
python benchmarks/performance_contract.py \
  --compare /tmp/wolfxl-baseline.json /tmp/wolfxl-final.json \
  --output /tmp/wolfxl-cumulative.json

# Measure the verifier change independently using the same two assertions.
python benchmarks/performance_contract.py --label verifier-attribution \
  --verifier-comparison --fixture-dir /tmp/wolfxl-perf-v2-fixtures \
  --output /tmp/wolfxl-verifier.json --rounds 5 --warmups 1
```

For a shorter matched run, pass the same `--cases` list at every stage, for example
`edit_top,edit_middle,edit_bottom,edit_merged,styled_small_eager,styled_large_read_only,styled_merged_eager`.
`--engines wolfxl` is diagnostic only; cumulative comparisons require both engines.
`--edit-rows`, `--styled-rows`, `--small-rows`, `--style-cardinality`, `--rounds` and
`--warmups` must match between receipts. The comparator also rejects different
fixtures, harness source, Python/platform details and openpyxl source/version.
It computes **baseline WolfXL / final WolfXL** directly for each workload, and
reports **final openpyxl / final WolfXL** separately. Isolated gains are never
multiplied, and verifier gains are never described as engine gains.
