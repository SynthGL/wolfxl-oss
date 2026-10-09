"""Verify a root README publication delta without relaxing source bindings.

Requires Python 3.11+. The measured wheel is inspected, never rebuilt.
"""

import argparse
import hashlib
import json
from pathlib import Path, PurePosixPath
import posixpath
import re
import subprocess
import sys
import tempfile
import tomllib
import zipfile


RUNTIME_ROOTS = ("src/", "python/", "crates/", "vendor/")
ROOT_CONFIGURATION = {
    "Cargo.toml",
    "Cargo.lock",
    "build.rs",
    "pyproject.toml",
    "rust-toolchain",
    "rust-toolchain.toml",
    "setup.py",
    "setup.cfg",
    "MANIFEST.in",
}
SOURCE_EXTENSIONS = {".rs", ".py", ".toml", ".c", ".cc", ".cpp", ".h", ".sh", ".cmake"}


def sha(raw):
    return hashlib.sha256(raw).hexdigest()


def git(repo, *args):
    return subprocess.check_output(["git", "-C", str(repo), *args])


def blob(repo, ref, path):
    return git(repo, "show", f"{ref}:{path}")


def resolve(repo, ref):
    return git(repo, "rev-parse", ref + "^{commit}").decode().strip()


def verify_code_binding(args, code, metadata, recorded):
    """Rerun the unchanged strict verifier, including its original lock proof."""
    with tempfile.TemporaryDirectory(prefix="readme-code-binding-") as directory:
        output = Path(directory) / "binding.json"
        command = [
            sys.executable,
            str(args.binding_verifier),
            "--repo",
            str(args.repo),
            "--head",
            code,
            "--remote-head",
            recorded["publication_remote_ref"],
            "--build-metadata",
            str(args.build_metadata),
            "--output",
            str(output),
        ]
        if recorded["excluded_standalone_lock_exception"]["applied"]:
            command.append("--allow-excluded-standalone-lock")
        subprocess.run(command, check=True, capture_output=True, text=True)
        fresh = json.loads(output.read_text())
    for key in (
        "measured_source_ref",
        "publication_resolved_local_ref",
        "publication_tree",
        "source_manifest_sha256",
        "build_metadata_sha256",
        "native_sha256",
        "wheel_sha256",
        "verified_inputs",
        "excluded_standalone_lock_exception",
        "all_measured_inputs_identical",
        "runtime_and_build_dependency_closure_inputs_identical",
    ):
        if fresh[key] != recorded[key]:
            raise RuntimeError(f"Recorded measured-source-to-code binding differs: {key}")
    if fresh["measured_source_ref"] != metadata["local_git_head"]:
        raise RuntimeError("Code binding does not reference the measured build source")
    return fresh


def audit_readme_source_uses(repo, ref, manifest):
    """Reject README filename uses in native/runtime source; enumerate metadata uses."""
    scanned, references = [], []
    for entry in manifest:
        path = entry["path"]
        if path not in ROOT_CONFIGURATION and not path.startswith(RUNTIME_ROOTS):
            continue
        if PurePosixPath(path).suffix not in SOURCE_EXTENSIONS:
            continue
        raw = blob(repo, ref, path)
        text = raw.decode("utf-8")
        scanned.append({"path": path, "sha256": sha(raw)})
        if path.endswith(".toml"):
            document = tomllib.loads(text)

            def visit(value, keys=()):
                if isinstance(value, dict):
                    for key, item in value.items():
                        visit(item, keys + (key,))
                elif isinstance(value, list):
                    for item in value:
                        visit(item, keys)
                elif isinstance(value, str) and re.search(r"(?i)readme\.md", value):
                    allowed = keys in {
                        ("package", "readme"),
                        ("package", "include"),
                        ("package", "exclude"),
                        ("project", "readme"),
                        ("project", "readme", "file"),
                    }
                    if not allowed:
                        raise RuntimeError(
                            f"README use outside packaging metadata: {path}:{'.'.join(keys)}"
                        )
                    target = posixpath.normpath(str(PurePosixPath(path).parent / value.lstrip("/")))
                    references.append(
                        {
                            "path": path,
                            "key": ".".join(keys),
                            "value": value,
                            "resolved_package_input": target,
                            "classification": "packaging metadata",
                        }
                    )

            visit(document)
            continue
        # Reject even a suspicious source comment containing README.md; no blanket
        # exception is made for tests, include macros or build-script declarations.
        if re.search(r"(?i)readme\.md", text) or re.search(
            r"(?is)include_(?:str|bytes)\s*!\s*\([^)]*readme", text
        ):
            raise RuntimeError(
                f"README native/runtime/build source reference requires a rebuild audit: {path}"
            )
        for line, value in enumerate(text.splitlines(), 1):
            if "readme" in value.lower():
                references.append(
                    {
                        "path": path,
                        "line": line,
                        "text": value.strip(),
                        "classification": "non-root-README filename identifier, comment or label",
                    }
                )
    return {
        "scanned_source_file_count": len(scanned),
        "scanned_source_inputs_sha256": sha(
            json.dumps(scanned, sort_keys=True, separators=(",", ":")).encode()
        ),
        "scanned_source_extensions": sorted(SOURCE_EXTENSIONS),
        "readme_references": references,
        "explicit_root_readme_native_source_references": [],
        "scope": "Recorded root/native/Python/build sources and TOML configuration at the immutable docs ref; filename and include-macro audit, with TOML packaging fields distinguished",
    }


def wheel_metadata_delta(repo, code, docs, metadata):
    project = tomllib.loads(blob(repo, code, "pyproject.toml").decode())["project"]
    readme = project["readme"]
    path = readme if isinstance(readme, str) else readme["file"]
    before, after = blob(repo, code, path), blob(repo, docs, path)
    wheel = Path(metadata["wheel"])
    with wheel.open("rb") as handle:
        wheel_sha = hashlib.file_digest(handle, "sha256").hexdigest()
    if wheel_sha != metadata["wheel_sha256"]:
        raise RuntimeError("Measured wheel bytes no longer match the recorded build")
    with zipfile.ZipFile(wheel) as archive:
        members = [name for name in archive.namelist() if name.endswith(".dist-info/METADATA")]
        if len(members) != 1:
            raise RuntimeError("Measured wheel has an ambiguous METADATA member")
        raw = archive.read(members[0])
    body = raw.replace(b"\r\n", b"\n").partition(b"\n\n")[2]

    def normalize(value):
        return value.replace(b"\r\n", b"\n").rstrip(b"\n")

    if normalize(body) != normalize(before):
        raise RuntimeError(
            "Measured wheel long description differs from the code-ref packaging input"
        )
    changed = before != after
    note = "A wheel built from the docs ref would contain the rewritten README description. The measured wheel still contains the pre-rewrite description and was not rebuilt."
    if not changed:
        note = f"Python wheel long-description input {path} is unchanged; root README.md is a separate recorded documentation/packaging input. The measured wheel was not rebuilt."
    return {
        "project_readme_input": path,
        "code_input_sha256": sha(before),
        "docs_input_sha256": sha(after),
        "code_input_bytes": len(before),
        "docs_input_bytes": len(after),
        "long_description_input_changed": changed,
        "measured_wheel": str(wheel),
        "measured_wheel_sha256_verified": wheel_sha,
        "metadata_member": members[0],
        "measured_metadata_sha256": sha(raw),
        "measured_description_matches_code_input": True,
        "measured_description_matches_docs_input": normalize(body) == normalize(after),
        "measured_wheel_rebuilt_for_docs_ref": False,
        "note": note,
    }


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--repo", required=True, type=Path)
    parser.add_argument("--code-ref", required=True)
    parser.add_argument("--docs-ref", required=True)
    parser.add_argument("--build-metadata", required=True, type=Path)
    parser.add_argument("--source-bindings", required=True, type=Path)
    parser.add_argument(
        "--binding-verifier",
        type=Path,
        default=Path(__file__).with_name("verify_publication_bindings.py"),
    )
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args()
    code, docs = resolve(args.repo, args.code_ref), resolve(args.repo, args.docs_ref)
    metadata_raw, binding_raw = args.build_metadata.read_bytes(), args.source_bindings.read_bytes()
    metadata, recorded = json.loads(metadata_raw), json.loads(binding_raw)
    fresh = verify_code_binding(args, code, metadata, recorded)
    manifest = metadata["source_manifest"]
    paths = {entry["path"] for entry in manifest}
    changes = []
    for line in (
        git(args.repo, "diff", "--name-status", "--no-renames", code, docs).decode().splitlines()
    ):
        status, path = line.split("\t", 1)
        changes.append({"status": status, "path": path})
        if path != "README.md" and (
            path in paths
            or path in ROOT_CONFIGURATION
            or path.startswith(RUNTIME_ROOTS)
            or path.startswith(".cargo/")
        ):
            raise RuntimeError(
                f"Docs delta changes a recorded/runtime/build input beyond README.md: {path}"
            )
    recorded_changes = [change for change in changes if change["path"] in paths]
    if recorded_changes != [{"status": "M", "path": "README.md"}]:
        raise RuntimeError("Only README.md must change among the recorded measured inputs")
    before, after = blob(args.repo, code, "README.md"), blob(args.repo, docs, "README.md")
    source_audit = audit_readme_source_uses(args.repo, docs, manifest)
    wheel = wheel_metadata_delta(args.repo, code, docs, metadata)
    measured_differences = [
        {"path": "README.md", "measured_sha256": sha(before), "docs_sha256": sha(after)}
    ]
    lock_exception = fresh["excluded_standalone_lock_exception"]
    if lock_exception["applied"]:
        measured_differences.append(
            {
                "path": lock_exception["path"],
                "measured_sha256": lock_exception["measured_sha256"],
                "docs_sha256": lock_exception["publication_sha256"],
            }
        )
    report = {
        "kind": "Measured code publication to final README documentation delta",
        "repo": str(args.repo.resolve()),
        "code_ref": code,
        "docs_ref": docs,
        "code_tree": fresh["publication_tree"],
        "docs_tree": git(args.repo, "rev-parse", docs + "^{tree}").decode().strip(),
        "measured_source_ref": metadata["local_git_head"],
        "measured_source_to_code_binding_valid": True,
        "measured_source_to_final_docs_all_recorded_inputs_identical": False,
        "code_to_docs_all_recorded_inputs_identical": False,
        "only_changed_recorded_input": "README.md",
        "all_other_recorded_inputs_identical_between_code_and_docs": True,
        "runtime_and_native_build_inputs_unchanged_between_code_and_docs": True,
        "recorded_input_count": len(manifest),
        "code_to_docs_identical_recorded_input_count": len(manifest) - 1,
        "measured_source_to_docs_recorded_input_differences": measured_differences,
        "readme_delta": {
            "path": "README.md",
            "code_sha256": sha(before),
            "docs_sha256": sha(after),
        },
        "source_bindings": {
            "path": str(args.source_bindings),
            "sha256": sha(binding_raw),
            "all_measured_inputs_identical_at_code": fresh["all_measured_inputs_identical"],
            "runtime_build_dependency_closure_identical_at_code": fresh[
                "runtime_and_build_dependency_closure_inputs_identical"
            ],
            "excluded_standalone_lock_exception": fresh["excluded_standalone_lock_exception"],
        },
        "binding_verifier_sha256": sha(args.binding_verifier.read_bytes()),
        "readme_delta_verifier_sha256": sha(Path(__file__).read_bytes()),
        "build_metadata_sha256": sha(metadata_raw),
        "source_audit": source_audit,
        "wheel_metadata": wheel,
        "all_code_to_docs_git_changes": changes,
        "note": "Strict measured-source-to-code bindings remain intact and were rerun unchanged. Final README publication changes a recorded packaging input; this separate proof does not claim all measured build inputs match the final docs tree. No native build or performance timing was run.",
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(report, indent=2, sort_keys=True) + "\n")
    print(
        json.dumps(
            {
                key: report[key]
                for key in (
                    "code_ref",
                    "docs_ref",
                    "only_changed_recorded_input",
                    "runtime_and_native_build_inputs_unchanged_between_code_and_docs",
                )
            }
        )
    )


if __name__ == "__main__":
    main()
