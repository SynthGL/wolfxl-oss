"""Excel rejects out-of-order cells even when openpyxl can read their values."""

from __future__ import annotations

from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

import openpyxl
import pytest
from openpyxl.styles import Font
from openpyxl.utils.cell import coordinate_to_tuple

import wolfxl
from wolfxl.styles import Font as WolfFont


@pytest.mark.parametrize(
    ("existing", "writes"),
    [
        (("A1", "C1"), ("B1",)),
        (("C1",), ("A1", "B1")),
        (("A1", "F1"), ("D1", "B1", "G1")),
        (("Z1", "AB1"), ("AA1", "Y1", "XFD1")),
        (("A1", "B1", "C1"), ("B1",)),
        (("A1", "C1"), ("D1",)),
        (("A3", "A10"), ("A1048576", "A11", "A9", "A2", "A1")),
        (("A3", "C3", "A10", "C10"), ("B10", "AA2", "A1", "B3", "A9", "A11")),
        ((), ("AB10", "AA2", "A10", "A1", "A1048576", "A2")),
    ],
    ids=[
        "middle",
        "left",
        "multiple",
        "numeric-columns",
        "existing",
        "right",
        "new-rows",
        "mixed-rows-and-cells",
        "empty",
    ],
)
@pytest.mark.parametrize("styled_empty", [False, True], ids=["values", "styled-empty"])
@pytest.mark.parametrize("style_only", [False, True], ids=["write-value", "write-style"])
def test_modify_saves_rows_and_cells_in_numeric_order(
    tmp_path: Path,
    existing: tuple[str, ...],
    writes: tuple[str, ...],
    styled_empty: bool,
    style_only: bool,
) -> None:
    source = tmp_path / "source.xlsx"
    output = tmp_path / "output.xlsx"
    original = openpyxl.Workbook()
    sheet = original.active
    for index, ref in enumerate(existing, 1):
        if styled_empty:
            sheet[ref].font = Font(bold=True)
        else:
            sheet[ref] = index
    original.create_sheet("Untouched")["A1"] = "keep me"
    original.save(source)
    original.close()
    source_bytes = source.read_bytes()

    workbook = wolfxl.load_workbook(source, modify=True)
    for index, ref in enumerate(writes, 100):
        if style_only:
            workbook.active[ref].font = WolfFont(italic=True)
        else:
            workbook.active[ref] = index
    # Repeated saves must not duplicate inserted cells or disturb their order.
    for _ in range(2):
        workbook.save(output)
        _assert_order(output, set(existing) | set(writes))
    workbook.close()

    reopened = openpyxl.load_workbook(output)
    for index, ref in enumerate(writes, 100):
        if style_only:
            expected = existing.index(ref) + 1 if ref in existing and not styled_empty else None
            assert reopened.active[ref].value == expected
            assert reopened.active[ref].font.italic
        else:
            assert reopened.active[ref].value == index
    for index, ref in enumerate(existing, 1):
        if ref not in writes:
            assert reopened.active[ref].value == (None if styled_empty else index)
        if styled_empty and not (style_only and ref in writes):
            assert reopened.active[ref].font.bold
    assert reopened["Untouched"]["A1"].value == "keep me"
    reopened.close()
    assert source.read_bytes() == source_bytes
    with ZipFile(source) as before, ZipFile(output) as after:
        assert before.read("xl/worksheets/sheet2.xml") == after.read("xl/worksheets/sheet2.xml")


@pytest.mark.parametrize("edit", ["value-new-row", "value-new-cell", "style-only"])
def test_reediting_out_of_order_cells_does_not_duplicate_them(tmp_path: Path, edit: str) -> None:
    source = tmp_path / "damaged.xlsx"
    output = tmp_path / "reedited.xlsx"
    original = openpyxl.Workbook()
    for ref, value in (("A1", 1), ("B1", 2), ("C1", 3)):
        original.active[ref] = value
    original.active["B1"].font = Font(bold=True)
    original.save(source)
    original.close()

    # Simulate the A1,C1,B1 output produced by older WolfXL versions (#114).
    with ZipFile(source) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    sheet_path = "xl/worksheets/sheet1.xml"
    root = ET.fromstring(parts[sheet_path])
    row = root.find("s:sheetData/s:row", ns)
    assert row is not None
    row[:] = [row[0], row[2], row[1]]
    parts[sheet_path] = ET.tostring(root)
    with ZipFile(source, "w", ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    source_bytes = source.read_bytes()

    workbook = wolfxl.load_workbook(source, modify=True)
    if edit == "style-only":
        workbook.active["B1"].font = WolfFont(italic=True)
    else:
        workbook.active["B1"] = 20
        workbook.active["A2" if edit == "value-new-row" else "E1"] = 5
    workbook.save(output)
    workbook.close()

    # Existing malformed order is not repaired here, but must not gain duplicates.
    with ZipFile(output) as archive:
        root = ET.fromstring(archive.read(sheet_path))
    refs = [c.attrib["r"] for c in root.findall("s:sheetData/s:row/s:c", ns)]
    expected = ["A1", "C1", "B1"]
    if edit != "style-only":
        expected.append("A2" if edit == "value-new-row" else "E1")
    assert refs == expected
    reopened = openpyxl.load_workbook(output)
    assert reopened.active["A1"].value == 1
    assert reopened.active["C1"].value == 3
    assert reopened.active["B1"].value == (2 if edit == "style-only" else 20)
    assert getattr(reopened.active["B1"].font, "italic" if edit == "style-only" else "bold")
    reopened.close()
    assert source.read_bytes() == source_bytes


def _assert_order(path: Path, expected_refs: set[str]) -> None:
    with ZipFile(path) as archive:
        root = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
    ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    rows = root.findall("s:sheetData/s:row", ns)
    row_numbers = [int(row.attrib["r"]) for row in rows]
    assert row_numbers == sorted(set(row_numbers))
    actual = []
    for row in rows:
        refs = [cell.attrib["r"] for cell in row.findall("s:c", ns)]
        assert all(coordinate_to_tuple(ref)[0] == int(row.attrib["r"]) for ref in refs)
        assert refs == sorted(set(refs), key=coordinate_to_tuple)
        actual.extend(refs)
    assert actual == sorted(expected_refs, key=coordinate_to_tuple)
