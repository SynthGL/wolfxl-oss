# Existing-cell XML splice evidence (Community)

The isolated native change removes reserialization of untouched worksheet XML. Existing value-cell edits near the middle and bottom become faster; top edits show no measured improvement.

Source: `93e48b2` → `944b956`. Standard source-built release wheels, Rust 1.90, Python 3.12.14; identical WolfXL Python sources and openpyxl 3.1.5. The fixture is 200,000 rows × 5 populated columns (1,000,000 cells), not eight populated columns. One warmup and five measured operations per case ran in a reserved quiet lane.

| Edit position | Save before → after | Save speedup | Load/edit/save/close before → after | Operation speedup |
|---|---:|---:|---:|---:|
| top | 0.39093s → 0.40701s | 0.96× | 0.51353s → 0.52839s | 0.97× |
| middle | 0.67713s → 0.44065s | 1.54× | 0.78305s → 0.55303s | 1.42× |
| bottom | 0.94586s → 0.56104s | 1.69× | 1.05208s → 0.67837s | 1.55× |

The seven-sample paired patch-only profile shows 0.97× top, 3.06× middle, and 3.11× bottom. It times only `patch_worksheet`: inflate/read, ZIP compression, compilation, validation, and output deallocation are excluded. Save timing includes remaining ZIP work, which limits complete-operation improvement. Full operation runs were sequential; patch-only alternates variant order. The patch-only v1 ZIP fixture differs from the full-operation v2 ZIP, but both contain byte-identical worksheet XML (SHA256 `7d7b766636aa47c9b9fa31df79a2f6047a73fafcbbcc54d402a5107373e94000`, 49,394,096 bytes); both artifact hashes are recorded.

Each case passed a full independent eager openpyxl reopen with equal complete semantic signatures, ZIP CRC checks, XML parsing, and exact untouched-part bytes. The final source passed 28 Rust sheet tests and 10 source-built Python preservation smoke tests, including byte preservation and two successive saves. XML bytes outside changed cell spans are retained; changed cells follow the existing writer's value/formula/style behavior. Insertions and unsupported mutations still use the generic path.

Raw data: [before](before.json), [after](after.json), [paired native patch](native-patch.json). [Summary](summary.json) includes ranges, receipt hashes and limits. [Build inputs](build-inputs.json) binds clean source builds and native hashes; every runtime/build input was rechecked against its build manifest after measurement. The candidate harness marks the checkout dirty solely because these untracked evidence files were present.

Reproduce patch-only using `benchmarks/bench_sheet_splices.py` at `944b956` with `--baseline-root`, `--candidate-root`, the matching million-cell fixture, `--rows 200000 --rounds 7 --warmups 1`, an isolated target directory and `--output`. Reproduce full phases with the performance-contract harness at `de85cf7`: `--cases edit_top,edit_middle,edit_bottom --engines wolfxl --rounds 5 --warmups 1`, once with each source-matched interpreter. WolfXL-only phase receipts are an isolated diagnostic; paired openpyxl and cumulative comparisons are recorded separately.
