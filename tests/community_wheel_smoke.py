from __future__ import annotations

from importlib.metadata import version
from importlib.util import find_spec
from pathlib import Path
from tempfile import TemporaryDirectory

import wolfxl
from wolfxl import Workbook, load_workbook


assert version("wolfxl") == "2.0.4"
assert wolfxl.__version__ == "2.0.4"

for module_name in (
    "wolfxl.operations",
    "wolfxl.render",
    "wolfxl.integrations",
    "wolfxl._conversion",
    "wolfxl._power_query",
    "wolfxl._vba_runtime",
):
    assert find_spec(module_name) is None, module_name

with TemporaryDirectory() as directory:
    path = Path(directory) / "community-smoke.xlsx"
    workbook = Workbook()
    workbook.active["A1"] = "community"
    workbook.save(path)

    loaded = load_workbook(path)
    assert loaded.active["A1"].value == "community"
    loaded.close()

print("community artifact smoke: ok")
