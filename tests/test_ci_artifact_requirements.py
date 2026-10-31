"""Tests for .github/scripts/hashed_artifact_requirements.py.

The helper is stdlib only, so these tests never import the native extension.
"""

from __future__ import annotations

import hashlib
import importlib.util
import subprocess
import sys
from pathlib import Path
from types import ModuleType

import pytest

ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / ".github" / "scripts" / "hashed_artifact_requirements.py"

# sha256(b"abc"), FIPS 180-2 test vector.
ABC_SHA256 = "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad"


def _load_module() -> ModuleType:
    spec = importlib.util.spec_from_file_location("hashed_artifact_requirements", SCRIPT)
    assert spec is not None
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


helper = _load_module()


def _write(path: Path, data: bytes) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(data)
    return path


def _line(path: Path) -> str:
    digest = hashlib.sha256(path.read_bytes()).hexdigest()
    return f"{path.resolve().as_uri()} --hash=sha256:{digest}"


def _read_lines(path: Path) -> list:
    return path.read_text(encoding="utf-8").splitlines()


def test_exact_hash_and_space_safe_uri(tmp_path: Path) -> None:
    artifact = _write(tmp_path / "my dist" / "pkg one-1.0-py3-none-any.whl", b"abc")
    output = tmp_path / "requirements.txt"

    assert helper.main(["--output", str(output), str(artifact)]) == 0

    uri = artifact.resolve().as_uri()
    assert uri.startswith("file:///")
    assert " " not in uri
    assert "my%20dist/pkg%20one-1.0-py3-none-any.whl" in uri
    assert output.read_text(encoding="utf-8") == f"{uri} --hash=sha256:{ABC_SHA256}\n"


def test_quoted_relative_globs_are_expanded(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    wheel = _write(tmp_path / "dist" / "pkg-1.0-cp39-abi3-any.whl", b"wheel")
    sdist = _write(tmp_path / "dist" / "pkg-1.0.tar.gz", b"sdist")
    _write(tmp_path / "dist" / "notes.txt", b"not an artifact")
    monkeypatch.chdir(tmp_path)

    rc = helper.main(["--output", "req.txt", "dist/*.whl", "dist/*.tar.gz"])

    assert rc == 0
    assert _read_lines(tmp_path / "req.txt") == sorted([_line(wheel), _line(sdist)])


def test_order_is_sorted_and_duplicates_removed(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    dist = tmp_path / "dist"
    c = _write(dist / "c-1.0-py3-none-any.whl", b"c")
    a = _write(dist / "a-1.0-py3-none-any.whl", b"a")
    b = _write(dist / "b-1.0-py3-none-any.whl", b"b")
    monkeypatch.chdir(tmp_path)
    patterns = [
        "dist/c-1.0-py3-none-any.whl",
        "dist/*.whl",
        str(a.resolve()),
        "dist/./b-1.0-py3-none-any.whl",
    ]

    assert helper.main(["--output", "first.txt", *patterns]) == 0
    assert helper.main(["--output", "second.txt", *reversed(patterns)]) == 0

    expected = [_line(a), _line(b), _line(c)]
    assert expected == sorted(expected)
    assert _read_lines(tmp_path / "first.txt") == expected
    assert (tmp_path / "first.txt").read_bytes() == (tmp_path / "second.txt").read_bytes()


def test_unmatched_pattern_fails_without_writing(
    tmp_path: Path, capsys: pytest.CaptureFixture
) -> None:
    artifact = _write(tmp_path / "dist" / "pkg-1.0-py3-none-any.whl", b"abc")
    missing = str(tmp_path / "dist" / "*.tar.gz")
    output = tmp_path / "requirements.txt"

    rc = helper.main(["--output", str(output), str(artifact), missing])

    assert rc == 1
    assert not output.exists()
    assert f"pattern matched no files: {missing}" in capsys.readouterr().err


def test_failure_leaves_existing_output_untouched(tmp_path: Path) -> None:
    output = tmp_path / "requirements.txt"
    output.write_text("previous\n", encoding="utf-8")

    rc = helper.main(["--output", str(output), str(tmp_path / "nothing-*.whl")])

    assert rc == 1
    assert output.read_text(encoding="utf-8") == "previous\n"


def test_non_regular_file_match_fails(
    tmp_path: Path, capsys: pytest.CaptureFixture
) -> None:
    _write(tmp_path / "dist" / "pkg-1.0-py3-none-any.whl", b"abc")
    (tmp_path / "dist" / "bogus.whl").mkdir()
    output = tmp_path / "requirements.txt"

    rc = helper.main(["--output", str(output), str(tmp_path / "dist" / "*.whl")])

    assert rc == 1
    assert not output.exists()
    assert "not a regular file" in capsys.readouterr().err


def test_literal_path_with_glob_metacharacters(tmp_path: Path) -> None:
    artifact = _write(tmp_path / "pkg[1].whl", b"abc")
    output = tmp_path / "requirements.txt"

    assert helper.main(["--output", str(output), str(artifact)]) == 0
    assert _read_lines(output) == [f"{artifact.resolve().as_uri()} --hash=sha256:{ABC_SHA256}"]


def test_sha256_streams_in_bounded_chunks(tmp_path: Path) -> None:
    data = bytes(range(256)) * 9 + b"tail"
    artifact = _write(tmp_path / "big.whl", data)

    assert helper.sha256_file(artifact, chunk_size=7) == hashlib.sha256(data).hexdigest()
    assert helper.sha256_file(artifact) == hashlib.sha256(data).hexdigest()


def test_cli_entry_point_expands_quoted_glob(tmp_path: Path) -> None:
    artifact = _write(tmp_path / "dist dir" / "pkg-1.0-py3-none-any.whl", b"abc")
    output = tmp_path / "requirements.txt"

    ok = subprocess.run(
        [sys.executable, str(SCRIPT), "--output", str(output), "dist dir/*.whl"],
        cwd=str(tmp_path),
        capture_output=True,
        text=True,
        check=False,
    )
    assert ok.returncode == 0, ok.stderr
    assert _read_lines(output) == [f"{artifact.resolve().as_uri()} --hash=sha256:{ABC_SHA256}"]

    missing = subprocess.run(
        [sys.executable, str(SCRIPT), "--output", str(tmp_path / "x.txt"), "none/*.whl"],
        cwd=str(tmp_path),
        capture_output=True,
        text=True,
        check=False,
    )
    assert missing.returncode != 0
    assert "pattern matched no files: none/*.whl" in missing.stderr
    assert not (tmp_path / "x.txt").exists()
