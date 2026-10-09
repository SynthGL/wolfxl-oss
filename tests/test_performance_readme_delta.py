"""Synthetic README-delta checks; creates no native build or workbook fixtures."""

import importlib.util
import json
from pathlib import Path
import subprocess
import sys
import unittest
import zipfile


SCRIPT = Path(__file__).resolve().parents[1] / "scripts/performance/verify_readme_publication_delta.py"
BASE_TESTS = Path(__file__).with_name("test_performance_source_bindings.py")
spec = importlib.util.spec_from_file_location("binding_fixture", BASE_TESTS)
fixture_module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(fixture_module)


@unittest.skipIf(sys.version_info < (3, 11), "README-delta tooling requires Python 3.11+")
class ReadmeDeltaTests(unittest.TestCase):
    def setUp(self):
        self.fixture = fixture_module.PublicationBindingTests("test_default_identical_inputs_pass")
        self.fixture.setUp()
        self.addCleanup(self.fixture.doCleanups)
        self.fixture.files["README.md"] = "# Code README\n"
        self.fixture.files["PYPI.md"] = "# Package description\n"
        self.freeze_code("README.md")

    def freeze_code(self, readme):
        self.fixture.files["pyproject.toml"] = (
            '[build-system]\nrequires = []\n[project]\nreadme = "' + readme + '"\n'
        )
        self.fixture.freeze_measured()
        self.code = self.fixture.git("rev-parse", "HEAD")
        wheel = self.fixture.root / "measured.whl"
        with zipfile.ZipFile(wheel, "w") as archive:
            archive.writestr(
                "wolfxl.dist-info/METADATA",
                "Metadata-Version: 2.3\n\n" + self.fixture.files[readme],
            )
        metadata = json.loads(self.fixture.metadata.read_text())
        metadata["wheel"] = str(wheel)
        metadata["wheel_sha256"] = fixture_module.hashlib.sha256(wheel.read_bytes()).hexdigest()
        self.fixture.metadata.write_text(json.dumps(metadata))
        result, _ = self.fixture.verify()
        self.assertEqual(result.returncode, 0, result.stderr)
        self.binding = self.fixture.root / "result.json"

    def make_docs(self, extra=None):
        self.fixture.write("README.md", "# Rewritten performance README\n")
        self.fixture.write("docs/proof.md", "Documentation evidence\n")
        if extra:
            self.fixture.write(*extra)
        return self.fixture.commit()

    def verify(self, docs):
        output = self.fixture.root / "delta.json"
        result = subprocess.run(
            [
                sys.executable,
                str(SCRIPT),
                "--repo",
                str(self.fixture.repo),
                "--code-ref",
                self.code,
                "--docs-ref",
                docs,
                "--build-metadata",
                str(self.fixture.metadata),
                "--source-bindings",
                str(self.binding),
                "--output",
                str(output),
            ],
            text=True,
            capture_output=True,
        )
        return result, json.loads(output.read_text()) if output.exists() else None

    def test_readme_only_retains_code_binding_and_discloses_old_wheel_description(self):
        result, report = self.verify(self.make_docs())
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertTrue(report["measured_source_to_code_binding_valid"])
        self.assertFalse(report["measured_source_to_final_docs_all_recorded_inputs_identical"])
        self.assertTrue(report["runtime_and_native_build_inputs_unchanged_between_code_and_docs"])
        self.assertTrue(report["wheel_metadata"]["long_description_input_changed"])
        self.assertFalse(report["wheel_metadata"]["measured_wheel_rebuilt_for_docs_ref"])

    def test_separate_pypi_description_is_reported_as_unchanged(self):
        self.freeze_code("PYPI.md")
        result, report = self.verify(self.make_docs())
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(report["wheel_metadata"]["project_readme_input"], "PYPI.md")
        self.assertFalse(report["wheel_metadata"]["long_description_input_changed"])

    def test_another_recorded_input_change_is_rejected(self):
        result, report = self.verify(self.make_docs(("PYPI.md", "changed\n")))
        self.assertNotEqual(result.returncode, 0)
        self.assertIsNone(report)
        self.assertIn("beyond README.md: PYPI.md", result.stderr)

    def test_added_runtime_source_outside_manifest_is_rejected(self):
        result, _ = self.verify(self.make_docs(("src/new.rs", "pub fn new() {}\n")))
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("beyond README.md: src/new.rs", result.stderr)

    def test_embedded_readme_and_native_fs_or_build_dependency_are_rejected(self):
        cases = [
            ("src/lib.rs", 'pub const DOC: &str = include_str!("../README.md");\n'),
            ("src/lib.rs", 'pub const DOC: &[u8] = include_bytes!("../README.md");\n'),
            ("src/lib.rs", 'pub fn read() { std::fs::read_to_string("README.md").unwrap(); }\n'),
            ("build.rs", 'fn main() { println!("cargo:rerun-if-changed=README.md"); }\n'),
        ]
        original = {path: self.fixture.files[path] for path in ("src/lib.rs", "build.rs")}
        for path, value in cases:
            with self.subTest(path=path, source=value):
                self.fixture.files.update(original)
                self.fixture.files[path] = value
                self.freeze_code("README.md")
                result, _ = self.verify(self.make_docs())
                self.assertNotEqual(result.returncode, 0)
                self.assertIn("requires a rebuild audit", result.stderr)

    def test_forged_original_binding_is_rejected(self):
        report = json.loads(self.binding.read_text())
        report["native_sha256"] = "forged"
        self.binding.write_text(json.dumps(report))
        result, _ = self.verify(self.make_docs())
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("binding differs: native_sha256", result.stderr)


if __name__ == "__main__":
    unittest.main()
