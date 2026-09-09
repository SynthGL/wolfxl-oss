from __future__ import annotations

import re
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
    assert _project_field("description") == (
        "MIT-licensed Excel I/O for Python with openpyxl-shaped APIs and a Rust backend"
    )
    assert _project_field("readme") == PYPI_DESCRIPTION.name
    assert PYPI_DESCRIPTION.is_file()


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
