# Lazy merge metadata evidence

The native merge reader now scans refs and endpoint style IDs without constructing a worksheet cell model. Workbook open performs no Python worksheet XML scan. This initial standalone P0 receipt predates the byte-negative probe follow-up and measures open/first-access diagnostics; final full styled-read and edit/save throughput come from separate end-to-end receipts.

Both fixtures contain 20,000 rows × 5 styled columns (100,000 stored cells). The paired variant adds exactly one merge, D19999:E20000. They have a default border table; dedicated semantic tests cover non-default endpoint borders, live edits, colors and repeated saves. Fixture generation is outside timing.

| Matched operation | Before median (ms) | After median (ms) | Before / after |
|---|---:|---:|---:|
| unmerged/open | 9.490 | 1.196 | 7.938× |
| unmerged/open + merge_ranges | 132.581 | 31.187 | 4.251× |
| unmerged/open + first_border | 144.064 | 167.331 | 0.861× |
| one_merge/open | 730.399 | 1.203 | 607.086× |
| one_merge/open + merge_ranges | 910.147 | 33.264 | 27.362× |
| one_merge/open + first_border | 1000.423 | 87.120 | 11.483× |

The unmerged first-border operation regresses by 16.15% (0.861×): the new bounded merge-ref scan precedes the existing coordinate border reader's full native worksheet hydration. This isolated first-access regression is retained explicitly. The merged first-border result includes the deferred metadata and endpoint scans, so its 11.483× speedup is broader than the 607.085× open-only result. Full styled-read throughput and cumulative speedups require separate matched end-to-end measurements.

Untimed diagnostics show Python worksheet XML opens drop from one to zero for the unmerged fixture and from two to zero for the merged fixture. The merged first-border path calls `read_merged_ranges` once and `read_merged_endpoint_style_ids` once, with no `read_cell_border` call. Only the requested Python cell is created. Raw counters and samples are in `before.json` and `after.json`.

Each operation has one warm-up and seven measured iterations, median shown above, with GC before timing and `close()` outside timing. Baseline and optimized runs were serialized in a root-reserved quiet window. No compilation overlapped the measured pair. Process maximum RSS for the entire multi-workload probe was 128,168 KiB before and 78,420 KiB after; these are whole-probe peaks, not per-operation memory claims.

Baseline source: `93e48b23605c798fcb98b2e4a86445a0b9c6679e`. Final Python source: `9ca6e2c2d799c1f448b3ca15ff93339f5eb3e832`. The measured overlay used native `c4f24001355b98ebc5a4ade34181b4c976c8ad9029e17bb98ef549aed00a1b42` built from `089ba52765860c60f8b3dac592978587a89a9034`; native source was unchanged by the final Python-only marker fix. The final wheel was then force-reinstalled from `9ca6e2c2d799c1f448b3ca15ff93339f5eb3e832` with the same native hash, source manifest `1c7e98a3d6736e0828f911cd485a860492b1e4a6c01ae2467f71b4451e7a666c`. Actual Python source file hashes are recorded in `after.json`; fixture hashes match across runs.

Reproduce after preparing fixtures with the benchmark generator:

```sh
python benchmarks/bench_lazy_merge_metadata.py --fixture-dir /path/to/fixtures --iterations 7 --source-ref SOURCE_SHA --output results.json
```

Run the same command from each source-matched baseline/optimized environment. Keep compilation and other performance jobs outside the timing slot.

Validation of the final source-matched installed package: 86 focused Python tests passed, one existing optional skip. Public and Commercial merge reader tests each passed two cases; native bridge cargo checks passed for both. The measured Python hashes match the final installed package.
