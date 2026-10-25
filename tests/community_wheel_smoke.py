from __future__ import annotations

import os
import subprocess
from importlib.metadata import version
from importlib.util import find_spec
from pathlib import Path
from tempfile import TemporaryDirectory

import wolfxl
from wolfxl import Workbook, load_workbook


def calculate_roundtrip(directory: Path) -> None:
    source = directory / "report.xlsx"
    original = Workbook()
    sheet = original.active
    sheet.title = "Summary"
    sheet["B2"] = 1000
    sheet["B3"] = 950
    sheet["B4"] = "=SUM(B2:B3)"
    sheet["B2"].number_format = "#,##0.00"
    sheet.merge_cells("D1:E1")
    sheet["D1"] = "Merged heading"
    original.save(source)
    original.close()

    for modify in (False, True):
        workbook = load_workbook(source, modify=modify)
        workbook["Summary"]["B2"] = 1500
        assert workbook.calculate()["Summary!B4"] == 2450
        output = directory / f"calculated-{modify}.xlsx"
        workbook.save(output)
        workbook.close()

        for data_only in (False, True):
            saved = load_workbook(output, data_only=data_only)
            sheet = saved["Summary"]
            assert sheet["B2"].value == 1500
            assert sheet["B3"].value == 950
            assert sheet["B4"].value == (2450 if data_only else "=SUM(B2:B3)")
            assert sheet["B2"].number_format == "#,##0.00"
            assert sheet["D1"].value == "Merged heading"
            assert sheet["E1"].value is None
            saved.close()


def test_clean_installed_python39_wheel(tmp_path: Path) -> None:
    """Explicit remote qualification node; excluded from default test discovery."""
    project = Path(__file__).resolve().parents[1]
    dist = tmp_path / "dist"
    # The pinned gate image's interpreter directory is read-only.
    environment_variables = {
        **os.environ,
        "UV_PYTHON_INSTALL_DIR": str(tmp_path / "interpreters"),
    }
    subprocess.run(
        ["uv", "build", "--wheel", "--python", "3.9", "--out-dir", str(dist)],
        cwd=project,
        env=environment_variables,
        check=True,
    )
    wheels = list(dist.glob("wolfxl-2.0.12-cp39-cp39-*.whl"))
    assert len(wheels) == 1, wheels
    environment = tmp_path / "clean-python39"
    subprocess.run(
        ["uv", "venv", "--python", "3.9", str(environment)],
        env=environment_variables,
        check=True,
    )
    python = environment / "bin" / "python"
    subprocess.run(["uv", "pip", "install", "--python", str(python), str(wheels[0])], check=True)
    subprocess.run(
        [str(python), "-I", str(Path(__file__).resolve())],
        cwd=tmp_path,
        check=True,
    )


def main() -> None:
    assert version("wolfxl") == "2.0.12"
    assert wolfxl.__version__ == "2.0.12"
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
        calculate_roundtrip(Path(directory))
    print("community artifact smoke: calculate=2450; normal/modify caches and formulas: ok")


if __name__ == "__main__":
    main()
