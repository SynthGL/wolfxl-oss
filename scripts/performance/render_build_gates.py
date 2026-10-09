#!/usr/bin/env python3
"""Render build/gate provenance from captured JSON, without rebuilding or hashing binaries."""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path


EDITIONS = {
    "community": ("Community", "oss-baseline", "oss-final-v2"),
    "commercial": ("Commercial", "commercial-baseline", "commercial-final-v2"),
}
EXCEPTION_PATH = "crates/wolfxl-data-python/Cargo.lock"
STANDALONE_DIRECTORY = "crates/wolfxl-data-python"


def read_json(path: Path) -> dict:
    return json.loads(path.read_text())


def manifest_digest(entries: list[dict]) -> str:
    raw = json.dumps(entries, separators=(",", ":"), sort_keys=True).encode()
    return hashlib.sha256(raw).hexdigest()


def validate_source_binding(path: Path, metadata_path: Path, metadata: dict) -> dict:
    """Check captured binding consistency using recorded JSON, never native binaries."""
    binding = read_json(path)
    manifest = metadata["source_manifest"]
    entries = {entry["path"]: entry for entry in manifest}
    if len(entries) != len(manifest) or manifest_digest(manifest) != metadata["source_manifest_sha256"]:
        raise ValueError("Captured build manifest has duplicate paths or an inconsistent digest")
    expected_fields = {
        "measured_source_ref": metadata["local_git_head"],
        "source_manifest_sha256": metadata["source_manifest_sha256"],
        "native_sha256": metadata["native_sha256"],
        "wheel_sha256": metadata["wheel_sha256"],
        "manifest_scope": metadata["manifest_scope"],
        "standard_maturin_features": metadata["maturin_features"],
        "build_flag_environment": metadata["build_flag_environment"],
        "rustc": metadata["rustc"],
        "python": metadata["runtime_python"],
        "build_metadata_sha256": hashlib.sha256(metadata_path.read_bytes()).hexdigest(),
        "recorded_input_count": len(manifest),
        "all_other_recorded_inputs_identical": True,
        "runtime_and_build_dependency_closure_inputs_identical": True,
    }
    for key, expected in expected_fields.items():
        if binding.get(key) != expected:
            raise ValueError(f"Source binding differs from captured build at {key}")
    for key in ("publication_resolved_local_ref", "publication_remote_ref", "publication_tree"):
        if not binding.get(key):
            raise ValueError(f"Source binding lacks {key}")
    exception = binding.get("excluded_standalone_lock_exception", {})
    applied = exception.get("applied") is True
    if exception.get("applied") not in (True, False):
        raise ValueError("Source binding must explicitly record whether its exception was applied")
    if binding.get("all_measured_inputs_identical") is not (not applied):
        raise ValueError("Broad-manifest equality flag contradicts the explicit exception")
    if applied:
        if exception.get("requested") is not True or exception.get("path") != EXCEPTION_PATH:
            raise ValueError("Source binding permits an unsupported input exception")
        if exception.get("measured_sha256") != entries[EXCEPTION_PATH]["sha256"]:
            raise ValueError("Excluded lock before hash differs from captured build")
        if not exception.get("publication_sha256") or exception["publication_sha256"] == exception["measured_sha256"]:
            raise ValueError("Excluded lock exception lacks a distinct publication hash")
        proof = exception.get("proof", {})
        proof_fields = {
            "standalone_package_name": "wolfxl-data",
            "root_workspace_explicit_exclusion": STANDALONE_DIRECTORY,
            "standalone_has_own_workspace": True,
            "root_resolved_lock_omits_standalone_package": True,
            "root_manifest_has_no_standalone_dependency_reference": True,
            "only_semantic_lock_change": "wolfxl-reader.dependencies adds memchr",
            "root_cargo_toml_sha256": entries["Cargo.toml"]["sha256"],
            "root_cargo_lock_sha256": entries["Cargo.lock"]["sha256"],
            "standalone_manifest_sha256": entries[STANDALONE_DIRECTORY + "/Cargo.toml"]["sha256"],
        }
        for key, expected in proof_fields.items():
            if proof.get(key) != expected:
                raise ValueError(f"Excluded lock root-wheel proof differs at {key}")
    verified = [entry for entry in manifest if not applied or entry["path"] != EXCEPTION_PATH]
    if binding.get("verified_inputs") != verified or binding.get("verified_input_count") != len(verified):
        raise ValueError("Source binding does not enumerate every identical recorded input")
    publication = [
        {"path": entry["path"], "sha256": exception["publication_sha256"]}
        if applied and entry["path"] == EXCEPTION_PATH else entry
        for entry in manifest
    ]
    scope_digest = manifest_digest(verified)
    for key in ("measured_runtime_build_scope_sha256", "publication_runtime_build_scope_sha256"):
        if binding.get(key) != scope_digest:
            raise ValueError(f"Source binding has an inconsistent {key}")
    if binding.get("publication_recorded_inputs_sha256") != manifest_digest(publication):
        raise ValueError("Source binding publication input digest is inconsistent")
    groups = {
        "root_cargo": sum(entry["path"] in {"Cargo.toml", "Cargo.lock"} for entry in verified),
        "root_build_configuration": sum(entry["path"] in {"build.rs", "pyproject.toml"} for entry in verified),
        "root_python": sum(entry["path"].startswith("python/") for entry in verified),
        "root_rust": sum(entry["path"].startswith("src/") for entry in verified),
        "native_crates_and_assets": sum(entry["path"].startswith(("crates/", "vendor/")) for entry in verified),
    }
    if binding.get("verified_input_group_counts") != groups or (applied and not all(groups.values())):
        raise ValueError("Source binding group counts do not cover the recorded input closure")
    return binding


def code(value: object) -> str:
    return f"`{value}`"


def table(headers: list[str], rows: list[list[str]]) -> str:
    return "\n".join([
        "| " + " | ".join(headers) + " |",
        "| " + " | ".join("---" for _ in headers) + " |",
        *("| " + " | ".join(row) + " |" for row in rows),
    ])


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--evidence-root", type=Path, required=True)
    parser.add_argument("--edition", choices=("all", *EDITIONS), default="all")
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--plan", type=Path, required=True)
    parser.add_argument("--unused-input-exception", type=Path)
    parser.add_argument("--source-bindings", type=Path)
    args = parser.parse_args()
    evidence = args.evidence_root.resolve()
    selected = list(EDITIONS) if args.edition == "all" else [args.edition]
    if args.source_bindings and args.edition == "all":
        parser.error("--source-bindings accepts one edition's final report; select --edition")
    lines = [
        "Build and focused-gate provenance for the 2026-10-08 performance batch.",
        "",
        "This document is generated from captured build metadata and gate reviews. "
        "Its producer reads JSON only; it does not compile, execute tests, hash native "
        "binaries, or change performance inputs/results. Timed benchmark receipts "
        "remain separate evidence.",
        "",
    ]
    identities: list[list[str]] = []
    hashes: list[list[str]] = []
    package_files: list[dict[str, str]] = []
    gates: list[list[str]] = []
    descriptions: list[str] = []
    all_metadata = {}

    for edition in selected:
        label, baseline_id, final_id = EDITIONS[edition]
        baseline = read_json(evidence / baseline_id / "build-metadata.json")
        final = read_json(evidence / final_id / "build-metadata.json")
        review = read_json(evidence / final_id / "focused-test-review.json")
        for variant, metadata in ((baseline_id, baseline), (final_id, final)):
            if metadata.get("status") != "ready":
                raise ValueError(f"{variant}: captured build is not ready")
            if metadata.get("local_git_head") != metadata["upstream_source_ref"]:
                raise ValueError(f"{variant}: recorded source refs differ")
            if metadata.get("local_git_status") != "":
                raise ValueError(f"{variant}: recorded build checkout is dirty")
            all_metadata[variant] = metadata
            identities.append([
                label, variant, code(metadata["upstream_source_ref"]),
                code(metadata["native_sha256"]),
            ])
            hashes.append([
                variant, code(metadata["source_manifest_sha256"]),
                code(metadata["wheel_sha256"]),
            ])
            for name in ("build-metadata.json", "source-manifest.json", "build.log"):
                path = evidence / variant / name
                if path.exists():
                    package_files.append({"source": str(path), "destination": f"builds/{variant}/{name}"})
        for key in ("rustc", "maturin_features", "build_flag_environment"):
            if baseline.get(key) != final.get(key):
                raise ValueError(f"{label}: baseline/final {key} differ")
        if review.get("runtime_changed") is not False:
            raise ValueError(f"{label}: correction changed measured runtime")
        for name in ("focused-test-receipt.json", "focused-test-review.json", "focused-tests.log"):
            package_files.append({
                "source": str(evidence / final_id / name),
                "destination": f"builds/{final_id}/{name}",
            })
        for name in ("focused-tests-collection-error.log", "focused-test-receipt-collection-error.json", "xml-failure-isolation.log"):
            path = evidence / final_id / name
            if path.exists():
                package_files.append({"source": str(path), "destination": f"builds/{final_id}/{name}"})
        initial = review["initial_result"]
        gates.append([
            label, "Initial integrated file selection",
            f"{initial['passed']} passed; {initial['skipped']} skipped; {initial['failed']} initial test-expectation failure",
        ])
        if edition == "community":
            corrected = review["corrected_probe_validation"]
            if corrected["native_sha256"] != final["native_sha256"]:
                raise ValueError("Community corrected gate native differs")
            gates.append([label, "Corrected probe-only selection", f"{corrected['passed']} passed ({corrected['seconds']:.2f}s)"])
            descriptions.append(
                "Community's initial assertion incorrectly required an edited coordinate "
                "to be absent from the Cell map. Community intentionally stores a lightweight "
                "Cell and records compact dirty values. The corrected test checks that existing "
                "contract and retains probe-once assertions. Its test-only source is "
                f"{code(corrected['test_only_source_ref'])}. The 12 corrected cases overlap "
                "the original selection; these counts are not additive."
            )
        else:
            corrected = review["corrected_xml_validation"]
            if corrected["final"]["native_sha256"] != final["native_sha256"]:
                raise ValueError("Commercial corrected gate native differs")
            gates.append([label, "Native/Python probe selection within initial gate", "12 passed; already included in the 283 passing cases"])
            gates.append([label, "Corrected XML/formula contracts on baseline", f"{corrected['baseline']['passed']} passed ({corrected['baseline']['seconds']:.2f}s)"])
            gates.append([label, "Corrected XML/formula contracts on final", f"{corrected['final']['passed']} passed ({corrected['final']['seconds']:.2f}s)"])
            descriptions.append(
                "Commercial's initial exact-preservation fixture contained a formula. "
                "Existing non-noop saving intentionally invalidates stale cached formula "
                "values and publishes calculation-chain metadata; the relevant save and "
                "formula-cache implementations match baseline. The test-only correction "
                f"{code(corrected['test_only_source_ref'])} uses formula-free input for exact "
                "XML/package preservation and a separate formula case for intended cache, "
                "relationship, content-type and recalculation changes. Both cases pass on "
                "the exact baseline/final wheels with independent reopen and two-save checks."
            )

    lines += [table(["Edition", "Variant", "Frozen source", "Native SHA-256"], identities), "",
              table(["Variant", "Build-input manifest SHA-256", "Wheel SHA-256"], hashes), ""]
    first = all_metadata[EDITIONS[selected[0]][2]]
    lines += [
        f"Compiler: {code(first['rustc'])}; Cargo: {code(first['cargo'])}; "
        f"maturin: {code(first['maturin'])}. Final release builds used "
        f"{first['cargo_build_jobs']} Cargo jobs. Captured baseline/final Rust build flags "
        "and profile overrides are unset.",
        "",
    ]
    if "community" in selected:
        lines += [
            "The Community baseline metadata records rustc and maturin but not "
            "a separate Cargo version/jobs field.", "",
        ]
    for edition in selected:
        label, _, final_id = EDITIONS[edition]
        metadata = all_metadata[final_id]
        features = ", ".join(code(feature) for feature in metadata["maturin_features"])
        seed = metadata["cache_state"]
        lines += [
            f"{label} retained the standard pyproject features: {features}. "
            "The source-matched build command was "
            "`maturin build --release --locked --interpreter <CPython-3.12-build-env> --out <variant-wheel-dir>` "
            "with a distinct `CARGO_TARGET_DIR` per variant and pinned Rust 1.90.0.",
            "",
            f"Its isolated final target was seeded with {seed['copied_artifact_count']} "
            f"hash-verified third-party artifacts; {seed['excluded_workspace_entries']} "
            "WolfXL workspace artifact entries were excluded. Compiler, features, flags "
            "and complete locked third-party graph were checked. The added memchr edge "
            "changes only the workspace dependency graph. Cargo still validates normal "
            "fingerprints. The freshly built wheel was force-installed into a distinct "
            "environment, and installed native/Python entry hashes were verified against it.",
            "",
        ]
    lines += [table(["Edition", "Gate boundary", "Captured outcome"], gates), ""]
    for text in descriptions:
        lines += [text, ""]
    lines += [
        "Initial failing assertions, original nonzero gate receipts and collection-error "
        "logs are preserved. Corrected affected-file verification supplements those records; "
        "it does not rewrite them or imply a clean rerun of the entire original suite. "
        "Focused coverage includes merges/probes, index/cache invalidation, styled cells, "
        "window/reader lifecycle, preservation and record APIs. Prebuild touched Python "
        "Ruff, diff checks and Rust fmt gates passed; the Commercial production module-size "
        "and include-sharding gate passed.",
        "",
        "Verification covers this Linux CPython 3.12 environment, captured source-matched "
        "wheels, focused Python/Rust contracts, ZIP payload checks and independent openpyxl "
        "reopens. Microsoft Excel GUI/app certification was not run. Registry releases, "
        "macOS/Windows certification and GUI equivalence are outside this build receipt.",
        "",
    ]
    source_binding = None
    if args.source_bindings:
        final_id = EDITIONS[selected[0]][2]
        source_binding = validate_source_binding(
            args.source_bindings, evidence / final_id / "build-metadata.json", all_metadata[final_id]
        )
        lines += [
            f"The [source binding](source-bindings.json) maps the frozen measured build to "
            f"publication tree {code(source_binding['publication_tree'])}, at published code "
            f"reference {code(source_binding['publication_remote_ref'])}. It independently "
            f"records {source_binding['verified_input_count']} identical inputs of "
            f"{source_binding['recorded_input_count']} recorded inputs. This renderer checks "
            "the enumerated paths/hashes, counts, manifest and runtime-scope digests, "
            "compiler/features/native identities and any root-wheel exclusion proof "
            "against captured build JSON. It does not re-fetch a remote tree or promise "
            "that the local frozen measured commit is remotely available.", "",
        ]
    exception_state = {"status": "not applied to measured frozen inputs"}
    if "commercial" in selected:
        exception_state["path"] = EXCEPTION_PATH
        if source_binding:
            exception_state = source_binding["excluded_standalone_lock_exception"]
            if exception_state["applied"]:
                if args.unused_input_exception:
                    static_receipt = read_json(args.unused_input_exception)
                    if (
                        static_receipt.get("path") != EXCEPTION_PATH
                        or static_receipt.get("before_sha256") != exception_state["measured_sha256"]
                        or static_receipt.get("after_sha256") != exception_state["publication_sha256"]
                    ):
                        raise ValueError("Original lock correction receipt differs from publication binding")
                    package_files.append({"source": str(args.unused_input_exception.resolve()), "destination": "unused-input-exception.json"})
                lines += [
                    f"Exactly one conservative manifest input differs in publication: "
                    f"{code(EXCEPTION_PATH)}, measured {code(exception_state['measured_sha256'])}, "
                    f"published {code(exception_state['publication_sha256'])}. The generated "
                    "binding verifies every other recorded input is identical. Its root "
                    "Cargo/lock/config hashes and explicit workspace exclusion prove that "
                    "the standalone adapter lock is outside the measured root wheel's "
                    "dependency closure; its only semantic change adds the reader's direct "
                    "memchr edge. The literal broad source manifests differ, while the "
                    "recorded runtime/build closure remains identical. The original static "
                    "correction receipt and measured build metadata remain unchanged.", "",
                ]
            else:
                lines += ["The publication binding records no input exception; every recorded measured input is identical.", ""]
        elif args.unused_input_exception:
            exception_state = read_json(args.unused_input_exception)
            if exception_state.get("path") != EXCEPTION_PATH:
                raise ValueError("Unused-input exception must identify the exact standalone lockfile")
            for key in ("before_sha256", "after_sha256"):
                if key not in exception_state:
                    raise ValueError(f"Unused-input exception missing {key}")
            measured = all_metadata[EDITIONS["commercial"][2]]
            manifest = {entry["path"]: entry["sha256"] for entry in measured["source_manifest"]}
            if exception_state.get("measured_source_commit", measured["upstream_source_ref"]) != measured["upstream_source_ref"]:
                raise ValueError("Unused-input receipt is bound to another measured source")
            for key in ("runtime_changes", "native_changes"):
                if key in exception_state and exception_state[key] is not False:
                    raise ValueError("Unused-input correction changes measured runtime/native source")
            if manifest.get(EXCEPTION_PATH) != exception_state["before_sha256"]:
                raise ValueError("Unused-input before hash differs from frozen measured manifest")
            proof = exception_state.get("primary_wheel_proof", {})
            if proof:
                if not all(proof.get(key) is True for key in (
                    "workspace_excludes_standalone_adapter", "root_lock_has_no_data_package",
                    "root_package_has_no_data_dependency",
                )):
                    raise ValueError("Standalone adapter root-wheel exclusion proof is incomplete")
                if proof["root_cargo_lock_sha256"] != manifest["Cargo.lock"]:
                    raise ValueError("Standalone adapter proof uses a different root lock")
            elif exception_state.get("unused_by_measured_root_wheel") is not True:
                raise ValueError("Unused-input root-wheel exclusion proof is missing")
            commit = exception_state.get("commit", exception_state.get("isolated_fix_commit"))
            if not commit:
                raise ValueError("Unused-input exception has no isolated commit")
            complete = (
                exception_state.get("all_other_manifest_inputs_identical") is True
                or exception_state.get("all_other_recorded_inputs_identical") is True
            )
            status = (
                "The receipt verifies every other manifest input remains identical."
                if complete else
                "Static root-wheel exclusion is proved; the all-other-manifest-inputs-identical "
                "comparison is still pending before applying a publication-manifest exception."
            )
            lines += [
                f"The isolated publication consistency correction covers only {code(EXCEPTION_PATH)}: "
                f"before {code(exception_state['before_sha256'])}, after {code(exception_state['after_sha256'])}, "
                f"commit {code(commit)}. Its before hash matches the measured conservative "
                "source manifest. The separate adapter is excluded from the measured root "
                f"workspace and the root wheel does not use this lockfile. {status} "
                "The literal all-source manifest hash must remain distinct from a claim "
                "that the effective native build inputs are unchanged.", "",
            ]
            package_files.append({"source": str(args.unused_input_exception.resolve()), "destination": "unused-input-exception.json"})
        else:
            lines += [
                f"The proposed standalone-adapter consistency fix at {code(EXCEPTION_PATH)} "
                "is separate from frozen measured inputs. That adapter is excluded from the "
                "root workspace; its lock adds an already-locked memchr dependency edge. "
                "An exact before/after hash, isolated commit and all-other-inputs-identical "
                "receipt are still required before recording a publication-manifest exception. "
                "No broader lockfile exception is inferred.", "",
            ]
    lines += [
        "Reproduce this summary from preserved inputs:", "",
        "```bash",
        "python scripts/performance/render_build_gates.py "
        "--evidence-root docs/performance/2026-10-08/cumulative/builds "
        f"--edition {args.edition} --output docs/performance/2026-10-08/cumulative/build-gates.md "
        "--plan docs/performance/2026-10-08/cumulative/package-plan.json"
        + (" --source-bindings docs/performance/2026-10-08/cumulative/source-bindings.json"
           if args.source_bindings else "")
        + (" --unused-input-exception unused-input-exception.json"
           if args.unused_input_exception and "commercial" in selected else ""),
        "```", "",
        "The package plan copies original metadata, logs and review receipts unchanged. "
        "Keep the original measured wheels/environments and benchmark input/result files "
        "immutable. Run future builds and gates outside the exclusive timed lane. "
        "For a public-only evidence package, generate with `--edition community`; "
        "generate Commercial documentation with `--edition commercial`.", "",
    ]
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text("\n".join(lines))
    args.plan.parent.mkdir(parents=True, exist_ok=True)
    args.plan.write_text(json.dumps({
        "kind": "build/gate evidence packaging plan; no copies performed",
        "edition": args.edition,
        "producer": str(Path(__file__).resolve()),
        "generated_document": str(args.output.resolve()),
        "raw_files_to_copy_unchanged": package_files,
        "unused_input_exception": exception_state,
        "source_bindings": str(args.source_bindings.resolve()) if args.source_bindings else None,
        "measured_receipts": "Keep performance controller inputs/raw JSON/rendered evidence unchanged and package separately.",
        "verification": "Do not rerun builds/tests/native hashing while benchmark timing owns the lane.",
    }, indent=2) + "\n")
    print(json.dumps({"document": str(args.output.resolve()), "plan": str(args.plan.resolve()), "editions": selected}))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
