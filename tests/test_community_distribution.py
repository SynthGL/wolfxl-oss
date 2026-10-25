from __future__ import annotations

from importlib.util import find_spec

import pytest

from wolfxl import _rust


COMMERCIAL_MODULES = (
    "wolfxl.operations",
    "wolfxl.render",
    "wolfxl.integrations",
    "wolfxl._conversion",
    "wolfxl._power_query",
    "wolfxl._vba_runtime",
)


def test_community_distribution_reports_its_backends() -> None:
    build = _rust.build_info()
    assert build["package"] == "wolfxl"
    assert set(build["enabled_backends"]) == {
        "native-xlsx",
        "native-xlsb",
        "native-xls",
        "wolfxl",
    }


@pytest.mark.parametrize("module_name", COMMERCIAL_MODULES)
def test_community_distribution_omits_commercial_modules(module_name: str) -> None:
    assert find_spec(module_name) is None
