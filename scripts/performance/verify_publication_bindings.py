"""Bind a published Git tree to every input of a measured source build.

Requires Python 3.11+ for standard-library TOML parsing.
"""

import argparse
import copy
import hashlib
import json
from pathlib import Path
import posixpath
import subprocess
import tomllib


EXCLUDED_STANDALONE_DIRECTORY = "crates/wolfxl-data-python"
EXCLUDED_STANDALONE_LOCK = EXCLUDED_STANDALONE_DIRECTORY + "/Cargo.lock"
EXCLUDED_STANDALONE_MANIFEST = EXCLUDED_STANDALONE_DIRECTORY + "/Cargo.toml"
PROOF_INPUTS = {
    "Cargo.toml",
    "Cargo.lock",
    "build.rs",
    "pyproject.toml",
    EXCLUDED_STANDALONE_MANIFEST,
}


def digest(raw):
    return hashlib.sha256(raw).hexdigest()


def manifest_digest(entries):
    return digest(json.dumps(entries, separators=(",", ":"), sort_keys=True).encode())


def git_blob(repo, ref, path):
    return subprocess.check_output(["git", "-C", repo, "show", f"{ref}:{path}"])


def prove_excluded_lock_correction(root_inputs, measured_lock, published_lock):
    """Permit this one excluded lock's missing reader->memchr edge only."""
    root_manifest = tomllib.loads(root_inputs["Cargo.toml"].decode())
    root_lock = tomllib.loads(root_inputs["Cargo.lock"].decode())
    standalone = tomllib.loads(root_inputs[EXCLUDED_STANDALONE_MANIFEST].decode())
    package_name = standalone["package"]["name"]
    if EXCLUDED_STANDALONE_DIRECTORY not in root_manifest["workspace"].get("exclude", []):
        raise RuntimeError("Standalone adapter is not explicitly excluded from the root workspace")
    if "workspace" not in standalone:
        raise RuntimeError("Excluded adapter is not a standalone Cargo workspace")
    if any(p["name"] in {package_name, "wolfxl-data-python"} for p in root_lock["package"]):
        raise RuntimeError("Standalone adapter occurs in the root resolved Cargo dependency graph")

    def reject_dependency_reference(table):
        for key, value in table.items():
            if key in {
                "dependencies",
                "dev-dependencies",
                "build-dependencies",
            } and isinstance(value, dict):
                for alias, dependency in value.items():
                    actual_name = (
                        dependency.get("package", alias) if isinstance(dependency, dict) else alias
                    )
                    path = dependency.get("path", "") if isinstance(dependency, dict) else ""
                    if actual_name in {package_name, "wolfxl-data-python"} or (
                        path and posixpath.normpath(path) == EXCLUDED_STANDALONE_DIRECTORY
                    ):
                        raise RuntimeError("Root Cargo manifest references the standalone adapter")
            if isinstance(value, dict):
                reject_dependency_reference(value)

    reject_dependency_reference(root_manifest)
    before = tomllib.loads(measured_lock.decode())
    after = tomllib.loads(published_lock.decode())
    readers = [i for i, p in enumerate(before["package"]) if p["name"] == "wolfxl-reader"]
    if len(readers) != 1 or not any(p["name"] == "memchr" for p in before["package"]):
        raise RuntimeError("Excluded lock does not contain the expected reader/memchr packages")
    expected = copy.deepcopy(before)
    dependencies = expected["package"][readers[0]]["dependencies"]
    if "memchr" in dependencies:
        raise RuntimeError("Measured standalone lock already has the memchr edge")
    dependencies.append("memchr")
    dependencies.sort()
    if after != expected:
        raise RuntimeError("Excluded lock has changes beyond the reader->memchr dependency edge")
    return {
        "standalone_package_name": package_name,
        "root_workspace_explicit_exclusion": EXCLUDED_STANDALONE_DIRECTORY,
        "standalone_has_own_workspace": True,
        "root_resolved_lock_omits_standalone_package": True,
        "root_manifest_has_no_standalone_dependency_reference": True,
        "only_semantic_lock_change": "wolfxl-reader.dependencies adds memchr",
        "root_cargo_toml_sha256": digest(root_inputs["Cargo.toml"]),
        "root_cargo_lock_sha256": digest(root_inputs["Cargo.lock"]),
        "standalone_manifest_sha256": digest(root_inputs[EXCLUDED_STANDALONE_MANIFEST]),
    }


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--repo", required=True)
    parser.add_argument("--head", required=True)
    parser.add_argument("--remote-head", required=True)
    parser.add_argument("--build-metadata", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument(
        "--allow-excluded-standalone-lock",
        action="store_true",
        help="Permit only crates/wolfxl-data-python/Cargo.lock's missing reader->memchr edge, with proof and explicit differing hashes",
    )
    args = parser.parse_args()
    metadata_raw = args.build_metadata.read_bytes()
    metadata = json.loads(metadata_raw)
    if metadata["status"] != "ready" or metadata["local_git_status"].strip():
        raise RuntimeError("Build source was not clean and ready")
    manifest = metadata["source_manifest"]
    if manifest_digest(manifest) != metadata["source_manifest_sha256"]:
        raise RuntimeError("Recorded source manifest digest is inconsistent")
    entries = {entry["path"]: entry for entry in manifest}
    if len(entries) != len(manifest):
        raise RuntimeError("Recorded source manifest has duplicate paths")
    if (
        args.allow_excluded_standalone_lock
        and not (PROOF_INPUTS | {EXCLUDED_STANDALONE_LOCK}) <= entries.keys()
    ):
        raise RuntimeError("Recorded source manifest lacks required excluded-lock proof inputs")
    resolved = subprocess.check_output(
        ["git", "-C", args.repo, "rev-parse", args.head + "^{commit}"], text=True
    ).strip()
    inputs, publication_inputs, differences, proof_inputs = [], [], [], {}
    published_lock = None
    for entry in manifest:
        raw = git_blob(args.repo, resolved, entry["path"])
        actual = digest(raw)
        publication_inputs.append({"path": entry["path"], "sha256": actual})
        if actual != entry["sha256"]:
            if not args.allow_excluded_standalone_lock or entry["path"] != EXCLUDED_STANDALONE_LOCK:
                raise RuntimeError(f"Published input differs from measured build: {entry['path']}")
            differences.append(
                {
                    "path": entry["path"],
                    "measured_sha256": entry["sha256"],
                    "publication_sha256": actual,
                }
            )
            published_lock = raw
        else:
            inputs.append(entry)
            if entry["path"] in PROOF_INPUTS:
                proof_inputs[entry["path"]] = raw
    exception = {
        "requested": args.allow_excluded_standalone_lock,
        "applied": bool(differences),
    }
    if differences:
        # All root Cargo/config/Python/Rust/native-asset inputs were compared above.
        # The only differing path still requires exclusion and semantic-edge proof.
        measured_lock = git_blob(args.repo, metadata["local_git_head"], EXCLUDED_STANDALONE_LOCK)
        if digest(measured_lock) != entries[EXCLUDED_STANDALONE_LOCK]["sha256"]:
            raise RuntimeError("Measured standalone Git blob does not match build metadata")
        exception.update(differences[0])
        exception["proof"] = prove_excluded_lock_correction(
            proof_inputs, measured_lock, published_lock
        )
    build_scope_before = [
        e for e in manifest if not differences or e["path"] != EXCLUDED_STANDALONE_LOCK
    ]
    build_scope_after = [
        e for e in publication_inputs if not differences or e["path"] != EXCLUDED_STANDALONE_LOCK
    ]
    if build_scope_before != build_scope_after:
        raise RuntimeError(
            "Recorded runtime/build inputs differ outside the excluded standalone lock"
        )
    identical_groups = {
        "root_cargo": [e for e in inputs if e["path"] in {"Cargo.toml", "Cargo.lock"}],
        "root_build_configuration": [
            e for e in inputs if e["path"] in {"build.rs", "pyproject.toml"}
        ],
        "root_python": [e for e in inputs if e["path"].startswith("python/")],
        "root_rust": [e for e in inputs if e["path"].startswith("src/")],
        "native_crates_and_assets": [
            e for e in inputs if e["path"].startswith(("crates/", "vendor/"))
        ],
    }
    if differences and any(not group for group in identical_groups.values()):
        raise RuntimeError(
            "Recorded manifest lacks complete runtime/build input groups for the exception"
        )
    tree = subprocess.check_output(
        ["git", "-C", args.repo, "rev-parse", resolved + "^{tree}"], text=True
    ).strip()
    note = "Publication includes additional tests, diagnostics and documentation; every recorded build input is byte-identical. The remote tree is independently checked at publication."
    if differences:
        note = "Exactly one broad-manifest input differs: the explicitly excluded standalone adapter lock's reader->memchr edge. Every other recorded measured input is byte-identical, including root Cargo/build configuration, Python/Rust source and native assets. Workspace/lock proofs exclude that standalone lock from the measured root wheel's Cargo dependency closure. The remote tree is independently checked at publication."
    report = {
        "kind": "Published Git tree versus measured source-build inputs",
        "all_measured_inputs_identical": not differences,
        "all_other_recorded_inputs_identical": True,
        "runtime_and_build_dependency_closure_inputs_identical": True,
        "excluded_standalone_lock_exception": exception,
        "measured_source_ref": metadata["local_git_head"],
        "publication_repo": str(Path(args.repo).resolve()),
        "publication_local_ref": args.head,
        "publication_resolved_local_ref": resolved,
        "publication_remote_ref": args.remote_head,
        "publication_tree": tree,
        "verified_input_count": len(inputs),
        "recorded_input_count": len(manifest),
        "verified_input_group_counts": {
            name: len(group) for name, group in identical_groups.items()
        },
        "manifest_scope": metadata["manifest_scope"],
        "source_manifest_sha256": metadata["source_manifest_sha256"],
        "publication_recorded_inputs_sha256": manifest_digest(publication_inputs),
        "measured_runtime_build_scope_sha256": manifest_digest(build_scope_before),
        "publication_runtime_build_scope_sha256": manifest_digest(build_scope_after),
        "runtime_build_scope": "Every recorded measured input; only the proved excluded standalone lock is removed when the explicit exception is applied",
        "native_sha256": metadata["native_sha256"],
        "wheel_sha256": metadata["wheel_sha256"],
        "build_metadata_sha256": digest(metadata_raw),
        "standard_maturin_features": metadata["maturin_features"],
        "build_flag_environment": metadata["build_flag_environment"],
        "rustc": metadata["rustc"],
        "python": metadata["runtime_python"],
        "note": note,
        "verified_inputs": inputs,
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(report, indent=2, sort_keys=True) + "\n")
    print(
        json.dumps(
            {
                key: report[key]
                for key in (
                    "all_measured_inputs_identical",
                    "runtime_and_build_dependency_closure_inputs_identical",
                    "verified_input_count",
                    "measured_source_ref",
                    "publication_tree",
                )
            }
        )
    )


if __name__ == "__main__":
    main()
