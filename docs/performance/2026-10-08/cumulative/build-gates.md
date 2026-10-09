Build and focused-gate provenance for the 2026-10-08 performance batch.

This document is generated from captured build metadata and gate reviews. Its producer reads JSON only; it does not compile, execute tests, hash native binaries, or change performance inputs/results. Timed benchmark receipts remain separate evidence.

| Edition | Variant | Frozen source | Native SHA-256 |
| --- | --- | --- | --- |
| Community | oss-baseline | `93e48b23605c798fcb98b2e4a86445a0b9c6679e` | `0af60f1ed0becebb0b7f3033a2ecaf219026284c132d518e0f36bd09dc989baa` |
| Community | oss-final-v2 | `02884309b62de4f98b92c837c4c9184fe64ed05c` | `7ea062b1c794044b6cd1d894c780c14decbe3a5c66177bc6874b9ecdefc752dc` |

| Variant | Build-input manifest SHA-256 | Wheel SHA-256 |
| --- | --- | --- |
| oss-baseline | `abff347b18076ac73e2f7e542a3617e47dc683568655464235b61b1b7c5de80f` | `a517920ba7b376467517c6e7d9e60789b4a4f04958eb9d3c7f3745335cb44e7e` |
| oss-final-v2 | `753dc6532996dff8ab200aedb25dd61ee1a886350486f7a442ac0990d75d5671` | `2d4f15405dae21f98b3e9f2307dce2d45e1fdb5b732e6adceaaf6d23af9669c4` |

Compiler: `rustc 1.90.0 (1159e78c4 2025-09-14)`; Cargo: `cargo 1.90.0 (840b83a10 2025-07-30)`; maturin: `maturin 1.15.0`. Final release builds used 2 Cargo jobs. Captured baseline/final Rust build flags and profile overrides are unset.

The Community baseline metadata records rustc and maturin but not a separate Cargo version/jobs field.

Community retained the standard pyproject features: `extension-module`. The source-matched build command was `maturin build --release --locked --interpreter <CPython-3.12-build-env> --out <variant-wheel-dir>` with a distinct `CARGO_TARGET_DIR` per variant and pinned Rust 1.90.0.

Its isolated final target was seeded with 601 hash-verified third-party artifacts; 46 WolfXL workspace artifact entries were excluded. Compiler, features, flags and complete locked third-party graph were checked. The added memchr edge changes only the workspace dependency graph. Cargo still validates normal fingerprints. The freshly built wheel was force-installed into a distinct environment, and installed native/Python entry hashes were verified against it.

| Edition | Gate boundary | Captured outcome |
| --- | --- | --- |
| Community | Initial integrated file selection | 313 passed; 1 skipped; 1 initial test-expectation failure |
| Community | Corrected probe-only selection | 12 passed (0.23s) |

Community's initial assertion incorrectly required an edited coordinate to be absent from the Cell map. Community intentionally stores a lightweight Cell and records compact dirty values. The corrected test checks that existing contract and retains probe-once assertions. Its test-only source is `e1180a37e9f8d572dabf51bceddcf83fac4ac04b`. The 12 corrected cases overlap the original selection; these counts are not additive.

Initial failing assertions, original nonzero gate receipts and collection-error logs are preserved. Corrected affected-file verification supplements those records; it does not rewrite them or imply a clean rerun of the entire original suite. Focused coverage includes merges/probes, index/cache invalidation, styled cells, window/reader lifecycle, preservation and record APIs. Prebuild touched Python Ruff, diff checks and Rust fmt gates passed; the Commercial production module-size and include-sharding gate passed.

Verification covers this Linux CPython 3.12 environment, captured source-matched wheels, focused Python/Rust contracts, ZIP payload checks and independent openpyxl reopens. Microsoft Excel GUI/app certification was not run. Registry releases, macOS/Windows certification and GUI equivalence are outside this build receipt.

The [source binding](source-bindings.json) maps the frozen measured build to publication tree `3542f211f9bac8f475c9209a4bb5f833aed91944`, at published code reference `0ae83649ab649c5e588d786e96bb1130401ebf93`. It independently records 547 identical inputs of 547 recorded inputs. This renderer checks the enumerated paths/hashes, counts, manifest and runtime-scope digests, compiler/features/native identities and any root-wheel exclusion proof against captured build JSON. It does not re-fetch a remote tree or promise that the local frozen measured commit is remotely available.

Reproduce this summary from preserved inputs:

```bash
python scripts/performance/render_build_gates.py --evidence-root docs/performance/2026-10-08/cumulative/builds --edition community --output docs/performance/2026-10-08/cumulative/build-gates.md --plan docs/performance/2026-10-08/cumulative/package-plan.json --source-bindings docs/performance/2026-10-08/cumulative/source-bindings.json
```

The package plan copies original metadata, logs and review receipts unchanged. Keep the original measured wheels/environments and benchmark input/result files immutable. Run future builds and gates outside the exclusive timed lane. For a public-only evidence package, generate with `--edition community`; generate Commercial documentation with `--edition commercial`.
