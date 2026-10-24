"""Community XLS reads against independently produced, committed xlwt inputs."""

from __future__ import annotations

import datetime
import json
from pathlib import Path

import pytest

import wolfxl

QUALIFICATION = Path(__file__).resolve().parents[2] / "crates/wolfxl-reader/qualification"
FIXTURE = QUALIFICATION / "fixtures.xls"
_FIXTURES = [FIXTURE]


def test_xls_values_match_independent_xlwt_fixture() -> None:
    authored = json.loads((QUALIFICATION / "expected.json").read_text())
    with wolfxl.load_workbook(FIXTURE, data_only=True) as book:
        assert book.sheetnames == authored["sheet_names"]
        values = book["Values"]
        for address in ("B3", "C3", "D3"):
            assert values[address].value == authored["values"][address]
        assert values["E3"].value == datetime.datetime(2024, 2, 29)
        assert values["A1"].value is None
        assert book["Second sheet"]["A1"].value == authored["second_sheet"]["A1"]
    with wolfxl.load_workbook(FIXTURE, data_only=False) as book:
        assert book["Values"]["F3"].value == "=" + authored["values"]["F3"]["xls_formula"]


def test_xls_modify_raises() -> None:
    fixture = _FIXTURES[0]
    with pytest.raises(NotImplementedError, match="transcribe"):
        wolfxl.load_workbook(str(fixture), modify=True)


def test_xls_read_only_raises() -> None:
    fixture = _FIXTURES[0]
    with pytest.raises(NotImplementedError, match="xlsx-only"):
        wolfxl.load_workbook(str(fixture), read_only=True)


def test_xls_password_raises() -> None:
    fixture = _FIXTURES[0]
    with pytest.raises(NotImplementedError, match="xlsx-only"):
        wolfxl.load_workbook(str(fixture), password="anything")


def test_xls_cell_font_raises() -> None:
    fixture = _FIXTURES[0]
    wb = wolfxl.load_workbook(str(fixture))
    ws = wb.active
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None:
                with pytest.raises(NotImplementedError, match="xlsx-only"):
                    _ = cell.font
                return
    pytest.fail("no non-empty cells in fixture")


def test_xls_from_bytes() -> None:
    fixture = _FIXTURES[0]
    data = fixture.read_bytes()
    wb_bytes = wolfxl.load_workbook(data)
    wb_path = wolfxl.load_workbook(str(fixture))
    assert wb_bytes.sheetnames == wb_path.sheetnames


def test_xls_classify_format() -> None:
    """``wolfxl.classify_file_format`` reports 'xls' for this fixture
    both as a path and as bytes.

    Note: ``wolfxl.classify_format`` (without ``_file_``) is the
    SynthGL number-format archetype classifier; the file-format
    detector is the separate ``classify_file_format`` symbol.
    """
    fixture = _FIXTURES[0]
    fmt_path = wolfxl.classify_file_format(str(fixture))
    assert fmt_path == "xls", f"path -> {fmt_path!r}"
    fmt_bytes = wolfxl.classify_file_format(fixture.read_bytes())
    assert fmt_bytes == "xls", f"bytes -> {fmt_bytes!r}"
