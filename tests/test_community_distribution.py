from __future__ import annotations

from importlib.util import find_spec

import pytest

import wolfxl
from wolfxl import _rust


COMMERCIAL_MODULES = (
    "wolfxl.operations",
    "wolfxl.render",
    "wolfxl.integrations",
    "wolfxl._conversion",
    "wolfxl._power_query",
    "wolfxl._vba_runtime",
)


def test_community_distribution_reports_its_release_and_backends() -> None:
    assert wolfxl.__version__ == "2.0.3"

    build = _rust.build_info()
    assert build["package"] == "wolfxl"
    assert build["package_version"] == "2.0.3"
    assert set(build["enabled_backends"]) == {
        "native-xlsx",
        "native-xlsb",
        "calamine-xls",
        "wolfxl",
    }


@pytest.mark.parametrize("module_name", COMMERCIAL_MODULES)
def test_community_distribution_omits_commercial_modules(module_name: str) -> None:
    assert find_spec(module_name) is None
