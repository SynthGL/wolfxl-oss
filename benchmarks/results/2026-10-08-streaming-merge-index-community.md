# Dense streamed merge lookup evidence

A styled streamed cell previously searched every merged rectangle when checking its source style. Sheets with many merges therefore paid for all ranges repeatedly. The cache now indexes row intervals and bisects disjoint column buckets. It retains one parsed range at one tree node, preserves the original tuple precedence for accepted overlaps across the whole query path, and keeps a short tuple scan for eight or fewer ranges. Empty sheets return before touching the index.

The index and precedence map belong to the existing source-style cache. Reader replacement, merge mutation, structural remapping in Commercial, and close retain their existing cache invalidation behavior. Commercial reuses its worksheet interval index; Community builds the same index inside the style cache.

Validation: 19 new/existing cache tests passed. Guards cover 1,024 vertical ranges, 1,024 disjoint same-row column ranges, duplicate and overlapping bounds in both precedence orders, inclusive boundaries, million-row rectangles without row expansion, no-merge bypass, streamed style parity against openpyxl, dense merge mutation and close cleanup. Candidate visits are bounded by the row-tree path for valid disjoint buckets. Overlapping buckets intentionally retain a conservative scan; this is not a universal logarithmic-query claim.

| Case | Before lookup (ms) | After lookup (ms) | Lookup ratio | Before preparation (ms) | After preparation (ms) |
| --- | ---: | ---: | ---: | ---: | ---: |
| vertical | 147.944 | 27.182 | 5.44× | 1.209 | 9.267 |
| horizontal | 225.788 | 6.849 | 32.96× | 2.323 | 5.432 |
| empty | 0.922 | 0.935 | 0.99× | 0.003 | 0.003 |

Each nonempty fixture has 1,024 merged rectangles and 8,192 point queries. Each engine/case receives one discarded fresh-process warmup and three alternating fresh-process samples. Semantic checksums, fixture hashes and native-extension hashes match within each before/after pair. Preparation is measured separately and paid once per cached worksheet; the index adds several milliseconds in these fixtures.

**Scope:** these ratios measure cached point membership only. They exclude workbook open, XML traversal, style conversion and index preparation, and must not be presented as end-to-end workbook-read or openpyxl speedups. Original shared-style and cumulative contract receipts are unchanged. The empty control is sub-millisecond and its small variation is not evidence of a workload regression or gain.

Raw receipt: [2026-10-08-streaming-merge-index-community.json](2026-10-08-streaming-merge-index-community.json). Producer: `scripts/benchmark_streaming_merge_index.py`. Runtime source before `de3732a4d0db933e9761f5ad671e51063d1126cc`, after `ff5770d88d81d334953a8198d08f92c90953863e`.
