"""Community cell metadata parity on the native streaming path."""

from io import BytesIO
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

import openpyxl
import pytest

import wolfxl
from wolfxl import _streaming


@pytest.fixture
def cached_workbook(tmp_path):
    book = openpyxl.Workbook()
    sheet = book.active
    sheet.append(["number", "text", "boolean", "error", "missing", "error text"])
    sheet.append(["=1+29", '="hello"', "=TRUE()", "=1/0", "=1", '="#N/A"'])
    sheet["A1"].hyperlink = "https://example.com/table"
    sheet["C1"].hyperlink = openpyxl.worksheet.hyperlink.Hyperlink(
        ref="C1", location="Sheet!A2", tooltip="Internal link", display="Jump"
    )
    sheet["A4"] = "tail"
    original = BytesIO()
    book.save(original)
    output = BytesIO()
    namespace = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    ET.register_namespace("", namespace["s"])
    ET.register_namespace(
        "r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    )
    with ZipFile(original) as source, ZipFile(output, "w", ZIP_DEFLATED) as target:
        for entry in source.infolist():
            data = source.read(entry.filename)
            if entry.filename == "xl/worksheets/sheet1.xml":
                root = ET.fromstring(data)
                caches = [
                    ("n", "30.0"),
                    ("str", "hello"),
                    ("b", "1"),
                    ("e", "#DIV/0!"),
                    ("n", None),
                    ("str", "#N/A"),
                ]
                for cell, (kind, value) in zip(
                    root.findall("s:sheetData/s:row[@r='2']/s:c", namespace), caches
                ):
                    cell.set("t", kind)
                    cached = cell.find("s:v", namespace)
                    if cached is None:
                        cached = ET.SubElement(cell, "{" + namespace["s"] + "}v")
                    cached.text = value
                data = ET.tostring(root)
            target.writestr(entry, data)
    path = tmp_path / "cached.xlsx"
    path.write_bytes(output.getvalue())
    return path


def _source(path, file_object):
    return BytesIO(path.read_bytes()) if file_object else path


@pytest.mark.parametrize("read_only", [False, True])
@pytest.mark.parametrize("file_object", [False, True])
@pytest.mark.parametrize("data_only", [False, True])
def test_streaming_hyperlinks(cached_workbook, monkeypatch, read_only, file_object, data_only):
    # Force the same auto-stream branch used by large eager workbooks.
    monkeypatch.setattr(_streaming, "AUTO_STREAM_ROW_THRESHOLD", 0)
    expected = openpyxl.load_workbook(cached_workbook)
    actual = wolfxl.load_workbook(
        _source(cached_workbook, file_object), read_only=read_only, data_only=data_only
    )
    try:
        rows = actual.active.iter_rows()
        for actual_row, expected_row in zip(rows, expected.active.iter_rows(), strict=True):
            for actual_cell, expected_cell in zip(actual_row, expected_row, strict=True):
                link = actual_cell.hyperlink
                reference = expected_cell.hyperlink
                if reference is None:
                    assert link is None
                else:
                    assert (link.ref, link.target, link.location, link.tooltip, link.display) == (
                        reference.ref,
                        reference.target,
                        reference.location,
                        reference.tooltip,
                        reference.display,
                    )
    finally:
        actual.close()
        expected.close()


@pytest.mark.parametrize("read_only", [False, True])
@pytest.mark.parametrize("file_object", [False, True])
def test_streaming_cached_formula_types(cached_workbook, monkeypatch, read_only, file_object):
    monkeypatch.setattr(_streaming, "AUTO_STREAM_ROW_THRESHOLD", 0)
    expected = openpyxl.load_workbook(cached_workbook, data_only=True)
    actual = wolfxl.load_workbook(
        _source(cached_workbook, file_object), read_only=read_only, data_only=True
    )
    try:
        actual_row = next(actual.active.iter_rows(min_row=2, max_row=2))
        expected_row = next(expected.active.iter_rows(min_row=2, max_row=2))
        assert [(cell.value, cell.data_type) for cell in actual_row] == [
            (cell.value, cell.data_type) for cell in expected_row
        ]
    finally:
        actual.close()
        expected.close()
