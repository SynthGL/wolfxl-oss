# Community performance evidence — 2026-10-08

The final paired batches measure the complete optimization stack directly against the source-matched Community baseline and openpyxl controls. Ratios are workload-specific divisions of paired medians. Individual slice gains are not multiplied, and the existing-record recipe is an API-choice diagnostic.

| Evidence | Generated report | Captured inputs/results |
| --- | --- | --- |
| Core styled reads and two-cell edits | [core.md](core.md) | [core.json](core.json) |
| Read guards | [guards.md](guards.md) | [guards.json](guards.json) |
| Exact build identities and focused-gate boundaries | [build-gates.md](build-gates.md) | [build records](builds/) |
| Measured source to published code mapping | — | [source-bindings.json](source-bindings.json) |
| Code-to-README packaging delta | — | [readme-publication-delta.json](readme-publication-delta.json) |
| Styled-read semantic validation | [styled-read-validation.md](styled-read-validation.md) | [styled-read-validation.json](styled-read-validation.json) |

The [intermediate core](intermediate-core.md) ([raw](intermediate-core.json)) and [intermediate guards](intermediate-guards.md) ([raw](intermediate-guards.json)) preserve the earlier `19fdf4a` trial that exposed a top-edit assignment regression. That intermediate endpoint was rejected; the final paired rerun includes the conservative byte-negative merge probe that fixes the regression.

## Ordered implementation PRs

| PR | Purpose | Individual evidence |
| --- | --- | --- |
| [#40](https://github.com/SynthGL/wolfxl-oss/pull/40) | Lazy merge metadata, source-border lookup and conservative absent-merge probe | [lazy merge receipt](../../../../benchmarks/results/lazy_merge_metadata/README.md); [probe diagnostic](../../../../benchmarks/results/merge_negative_probe/README.md) |
| [#41](https://github.com/SynthGL/wolfxl-oss/pull/41) | Shared streaming style components and indexed dense merge membership | [style/resource receipt](../../../../benchmarks/results/2026-10-08-shared-stream-styles-community.json); [membership diagnostic](../../../../benchmarks/results/2026-10-08-streaming-merge-index-community.md) |
| [#42](https://github.com/SynthGL/wolfxl-oss/pull/42) | Frozen workload, fresh-process paired stages and semantic/package verification | [controller protocol](../../../../benchmarks/PERFORMANCE_STAGES.md); [verifier receipt](../../../../benchmarks/results/2026-10-08-verifier-contract-v2.json) |
| [#43](https://github.com/SynthGL/wolfxl-oss/pull/43) | Skip ZIP payload reads for unchanged template content types | [template no-op receipt](../template-noop.json) |
| [#44](https://github.com/SynthGL/wolfxl-oss/pull/44) | Share eager styles in bounded windows and release retained readers on close | [eager-style receipt](../eager-styles.json); [rejected prehydration diagnostic](../rejected-eager-prehydration.json) |
| [#45](https://github.com/SynthGL/wolfxl-oss/pull/45) | Splice unchanged worksheet XML around existing-cell edits | [XML splice receipt](../xml-splices/README.md) |
| [#46](https://github.com/SynthGL/wolfxl-oss/pull/46) | Document the existing values/style-ID record API | [API-choice receipt](../style-id-api-choice.json) |

## Reproduction

[Source bindings](source-bindings.json) map the locally frozen final build to its published code tree. Use that mapping when obtaining source; historical local measured commits are not promised to be separately available on the remote. Baseline identities and full compiler/features/native hashes are recorded in the [build gates](build-gates.md).

The separate [README delta proof](readme-publication-delta.json) checks the frozen code tree against the README rewrite before these receipts were added. Only root `README.md` changes among recorded build inputs; native, Python and build configuration remain identical. Community's wheel description comes from unchanged `PYPI.md`. The measured wheel was not rebuilt after the documentation edit. [The producer](../../../../scripts/performance/verify_readme_publication_delta.py) reruns the unchanged strict code binding and rejects other recorded-input or runtime-source changes.

Regenerate the build/gate summary from the preserved JSON:

```bash
python scripts/performance/render_build_gates.py \
  --evidence-root docs/performance/2026-10-08/cumulative/builds \
  --edition community \
  --source-bindings docs/performance/2026-10-08/cumulative/source-bindings.json \
  --output docs/performance/2026-10-08/cumulative/build-gates.md \
  --plan docs/performance/2026-10-08/cumulative/package-plan.json
```

Follow [benchmarks/PERFORMANCE_STAGES.md](../../../../benchmarks/PERFORMANCE_STAGES.md) for a fresh matched run using clean checkouts, isolated source-matched wheels and a quiet machine. Keep captured receipts unchanged and write new measurements to new paths. The build/gate report preserves initial test-expectation failures and the corrected affected-file checks; verification does not include Microsoft Excel GUI certification.
