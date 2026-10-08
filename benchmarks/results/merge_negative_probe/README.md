# Conservative absent-merge probe evidence

This is a Community Rust reader diagnostic on the exact million-cell top-edit fixture (200,000 rows × 5 columns, sheet `Data`). It measures a fresh exact merge-ref XML scan versus a bounded native byte-negative probe, including ZIP open and worksheet decompression on each call. Workbook open occurs before timing. There is one warm-up per method and seven serialized pairs in a reserved quiet slot after compilation finishes.

| Reader scan | Median |
|---|---:|
| Exact XML merge-ref scan | 257.855 ms |
| Bounded byte-negative probe | 38.791 ms |

The measured reader scan improves **6.647×**, avoiding about 219 ms of XML event processing on this source. This does not measure Python assignment, full styled reads, edit/save throughput or speedup over openpyxl; final matched end-to-end receipts supply those claims. Earlier standalone P0 receipts retain their original before/after samples and regressions.

The probe searches `mergeCell` bytes in bounded chunks, retaining eight bytes across short reads. Namespace, comment, CDATA and text occurrences are conservative positives that use exact XML parsing. Negative results are used only by lazy Python merge hydration, preserving the prior Python source-negative shortcut; they do not populate the exact native metadata cache, so later explicit exact calls still validate XML. Six focused reader tests cover chunk boundaries, namespaced merges, false positives, malformed exact-parser behavior and source I/O failures.

`reader_scan.json` retains all fourteen timed samples. `manifest.json` pins the fixture ZIP and worksheet XML hashes, compressed/uncompressed sizes and CRC, source commit and file hashes, compiler/build settings and measured executable hash. The measured Community source is `211fe26c459f4d0759a445cb0cd19aefa34d5a71`.

Reproduce with the checked-in reader example and a matching unmerged fixture:

```sh
cargo build --release --locked -p wolfxl-reader --example merge_presence_probe
./target/release/examples/merge_presence_probe /path/to/fixture.xlsx Data 7
```
