from __future__ import annotations

import re
import tomllib
from pathlib import Path
from urllib.parse import parse_qs, urlsplit


ROOT = Path(__file__).resolve().parents[1]
PYPI_DESCRIPTION = ROOT / "PYPI.md"
EXPECTED_ATTRIBUTION = {
    "utm_source": ["pypi"],
    "utm_medium": ["registry"],
    "utm_campaign": ["community_commercial_2026_09"],
}


def _project_field(name: str) -> str:
    pyproject = (ROOT / "pyproject.toml").read_text(encoding="utf-8")
    match = re.search(
        rf"^{re.escape(name)}\s*=\s*\"([^\"]+)\"$", pyproject, re.MULTILINE
    )
    assert match is not None, f"missing [project] field: {name}"
    return match.group(1)


def test_package_uses_dedicated_pypi_description() -> None:
    assert _project_field("readme") == PYPI_DESCRIPTION.name
    assert PYPI_DESCRIPTION.is_file()


def test_supported_python_floor_and_test_dependencies() -> None:
    project = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))
    assert project["project"]["requires-python"] == ">=3.11"
    classifiers = project["project"]["classifiers"]
    for minor in range(11, 16):
        assert f"Programming Language :: Python :: 3.{minor}" in classifiers
    for minor in (9, 10):
        assert f"Programming Language :: Python :: 3.{minor}" not in classifiers
    dependencies = project["project"]["optional-dependencies"]["test"]
    assert "pytest>=9.0.3" in dependencies
    assert "Pillow>=13.0.0" in dependencies
    assert all("python_version" not in dependency for dependency in dependencies)
    assert project["tool"]["ruff"]["target-version"] == "py311"
    assert project["tool"]["mypy"]["python_version"] == "3.11"

    lock = tomllib.loads((ROOT / "uv.lock").read_text(encoding="utf-8"))
    assert lock["requires-python"] == ">=3.11"
    for package in lock["package"]:
        if package["name"] == "pillow":
            assert package["version"] not in {"11.3.0", "12.3.0"}
        if package["name"] == "pytest":
            assert package["version"] != "8.4.2"


def test_package_summary_fits_the_pypi_summary_field() -> None:
    # PyPI rejects an upload whose Summary exceeds 512 characters.
    summary = _project_field("description")
    assert summary.strip() == summary
    assert 0 < len(summary) <= 512


def test_pypi_description_pins_the_current_release_and_uses_only_pypi_attribution() -> (
    None
):
    description = PYPI_DESCRIPTION.read_text(encoding="utf-8")
    urls = re.findall(r"\]\((https://wolfxl\.com[^)]+)\)", description)

    assert f"python -m pip install wolfxl=={_project_field('version')}" in description
    assert "After the Community release is published:" not in description
    assert "free product, not a trial" in description
    assert "Community 2.0 line" in description
    assert "Commercial 2.1+" in description
    assert urls
    for url in urls:
        assert parse_qs(urlsplit(url).query) == EXPECTED_ATTRIBUTION
