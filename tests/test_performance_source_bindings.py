"""Small synthetic Git cases; does not scan WolfXL or hash native artifacts."""

import hashlib
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest


SCRIPT = Path(__file__).resolve().parents[1] / "scripts/performance/verify_publication_bindings.py"
LOCK_PATH = "crates/wolfxl-data-python/Cargo.lock"
BEFORE_LOCK = """version = 4
[[package]]
name = "memchr"
version = "2.8.3"
[[package]]
name = "wolfxl-reader"
version = "0.1.0"
dependencies = ["zip"]
"""
AFTER_LOCK = BEFORE_LOCK.replace('["zip"]', '["memchr", "zip"]')


@unittest.skipIf(sys.version_info < (3, 11), "Source-binding tooling requires Python 3.11+")
class PublicationBindingTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory(prefix="binding-verifier-")
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)
        self.repo = self.root / "repo"
        self.repo.mkdir()
        self.git("init", "--quiet")
        self.git("config", "user.name", "Binding fixture")
        self.git("config", "user.email", "binding@example.invalid")
        self.files = {
            "Cargo.toml": '[workspace]\nexclude = ["crates/wolfxl-data-python"]\n[package]\nname = "wolfxl"\n',
            "Cargo.lock": 'version = 4\n[[package]]\nname = "wolfxl"\nversion = "2.3.0"\n',
            "build.rs": "fn main() {}\n",
            "pyproject.toml": "[build-system]\nrequires = []\n",
            "python/wolfxl/__init__.py": "value = 1\n",
            "src/lib.rs": "pub fn value() -> i32 { 1 }\n",
            "vendor/asset.txt": "native asset\n",
            "crates/wolfxl-data-python/Cargo.toml": '[package]\nname = "wolfxl-data"\n[workspace]\n',
            LOCK_PATH: BEFORE_LOCK,
        }
        self.freeze_measured()

    def git(self, *args):
        return subprocess.check_output(["git", "-C", str(self.repo), *args], text=True).strip()

    def write(self, path, value):
        target = self.repo / path
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(value)

    def commit(self):
        self.git("add", ".")
        self.git("commit", "--quiet", "-m", "synthetic input fixture")
        return self.git("rev-parse", "HEAD")

    def freeze_measured(self):
        for path, value in self.files.items():
            self.write(path, value)
        source_ref = self.commit()
        manifest = [
            {
                "path": path,
                "sha256": hashlib.sha256((self.repo / path).read_bytes()).hexdigest(),
            }
            for path in sorted(self.files)
        ]
        self.metadata = self.root / "metadata.json"
        self.metadata.write_text(
            json.dumps(
                {
                    "status": "ready",
                    "local_git_status": "",
                    "local_git_head": source_ref,
                    "source_manifest": manifest,
                    "source_manifest_sha256": hashlib.sha256(
                        json.dumps(manifest, separators=(",", ":"), sort_keys=True).encode()
                    ).hexdigest(),
                    "manifest_scope": "Synthetic root Cargo/Python/Rust/native-asset inputs",
                    "native_sha256": "synthetic-native",
                    "wheel_sha256": "synthetic-wheel",
                    "maturin_features": ["extension-module"],
                    "build_flag_environment": {},
                    "rustc": "synthetic",
                    "runtime_python": "synthetic",
                }
            )
        )

    def verify(self, allow=False):
        output = self.root / "result.json"
        command = [
            sys.executable,
            str(SCRIPT),
            "--repo",
            str(self.repo),
            "--head",
            "HEAD",
            "--remote-head",
            "synthetic-remote",
            "--build-metadata",
            str(self.metadata),
            "--output",
            str(output),
        ]
        if allow:
            command.append("--allow-excluded-standalone-lock")
        result = subprocess.run(command, text=True, capture_output=True)
        return result, json.loads(output.read_text()) if output.exists() else None

    def correct_lock(self):
        self.write(LOCK_PATH, AFTER_LOCK)
        self.commit()

    def test_default_identical_inputs_pass(self):
        result, report = self.verify()
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertTrue(report["all_measured_inputs_identical"])
        self.assertFalse(report["excluded_standalone_lock_exception"]["applied"])

    def test_default_rejects_even_the_excluded_lock(self):
        self.correct_lock()
        result, report = self.verify()
        self.assertNotEqual(result.returncode, 0)
        self.assertIsNone(report)
        self.assertIn("Published input differs", result.stderr)

    def test_explicit_single_edge_has_separate_truthful_hashes(self):
        self.correct_lock()
        result, report = self.verify(allow=True)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertFalse(report["all_measured_inputs_identical"])
        self.assertTrue(report["all_other_recorded_inputs_identical"])
        self.assertTrue(report["runtime_and_build_dependency_closure_inputs_identical"])
        self.assertEqual(report["recorded_input_count"], report["verified_input_count"] + 1)
        exception = report["excluded_standalone_lock_exception"]
        self.assertEqual(exception["path"], LOCK_PATH)
        self.assertEqual(
            exception["measured_sha256"],
            hashlib.sha256(BEFORE_LOCK.encode()).hexdigest(),
        )
        self.assertEqual(
            exception["publication_sha256"],
            hashlib.sha256(AFTER_LOCK.encode()).hexdigest(),
        )
        self.assertEqual(
            report["measured_runtime_build_scope_sha256"],
            report["publication_runtime_build_scope_sha256"],
        )

    def test_explicit_flag_rejects_another_runtime_change(self):
        self.correct_lock()
        self.write("src/lib.rs", "pub fn value() -> i32 { 2 }\n")
        self.commit()
        result, _ = self.verify(allow=True)
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("src/lib.rs", result.stderr)

    def test_explicit_flag_rejects_another_lock_change(self):
        self.write(LOCK_PATH, AFTER_LOCK.replace('version = "2.8.3"', 'version = "2.9.0"'))
        self.commit()
        result, _ = self.verify(allow=True)
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("changes beyond", result.stderr)

    def test_missing_workspace_exclusion_rejects_exception(self):
        self.files["Cargo.toml"] = '[workspace]\nexclude = []\n[package]\nname = "wolfxl"\n'
        self.freeze_measured()
        self.correct_lock()
        result, _ = self.verify(allow=True)
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("not explicitly excluded", result.stderr)

    def test_root_resolved_graph_reference_rejects_exception(self):
        self.files["Cargo.lock"] += '[[package]]\nname = "wolfxl-data"\nversion = "0.1.0"\n'
        self.freeze_measured()
        self.correct_lock()
        result, _ = self.verify(allow=True)
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("root resolved Cargo", result.stderr)


if __name__ == "__main__":
    unittest.main()
