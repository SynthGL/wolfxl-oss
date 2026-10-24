"""Pytest fixtures for the differential writer harness.

The function-scoped autouse fixture below forces ``WOLFXL_TEST_EPOCH=0`` so
ZIP entry mtimes are deterministic — without this, byte-level fixture
comparison is impossible (mtimes drift second-to-second).

Post-W5 the harness is single-backend: ``rust_xlsxwriter`` was ripped
out, so the writer the cases exercise is always ``NativeWorkbook``. The
``native_workbook`` fixture replaces the old ``dual_workbook``: it
constructs a plain ``Workbook()`` and yields it together with the output
path stem; the test populates the workbook, calls ``save()``, and the
written xlsx feeds into single-file structural and semantic checks
(openpyxl serves as the soft secondary oracle).
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Iterator

import pytest


@pytest.fixture(autouse=True)
def _force_test_epoch() -> Iterator[None]:
    """Force ``WOLFXL_TEST_EPOCH=0`` so ZIP mtimes are deterministic.

    Restore the previous value (or unset it) after each harness test so
    sibling calculation tests retain their own clock.
    """
    prev = os.environ.get("WOLFXL_TEST_EPOCH")
    os.environ["WOLFXL_TEST_EPOCH"] = "0"
    try:
        yield
    finally:
        if prev is None:
            os.environ.pop("WOLFXL_TEST_EPOCH", None)
        else:
            os.environ["WOLFXL_TEST_EPOCH"] = prev


@pytest.fixture
def native_workbook(tmp_path: Path):
    """Yield ``(Workbook, output_path)``.

    The build closure populates the workbook, then the test calls
    ``wb.save(str(output_path))`` to write the native xlsx.
    """
    import wolfxl

    wb = wolfxl.Workbook()
    yield wb, tmp_path / "out.xlsx"
