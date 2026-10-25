"""A ``<sheet r:id="">`` stub is skipped and reported; a dangling non-empty r:id still fails.

Mirrors Apache POI 60825.xlsx: the workbook lists ``Sheet1`` plus a ``Module``
entry with an empty ``r:id`` that names no worksheet part.
"""

from __future__ import annotations

import io
import zipfile

import pytest

import wolfxl
from wolfxl._rust import NativeXlsxBook

_WORKSHEET_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"


def _package(sheets_xml: str) -> bytes:
    workbook = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f"<sheets>{sheets_xml}</sheets></workbook>"
    )
    rels = (
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" Type="{_WORKSHEET_REL}" Target="worksheets/sheet1.xml"/>'
        "</Relationships>"
    )
    sheet = (
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>hello</t></is></c>'
        '<c r="B1"><v>42</v></c></row></sheetData></worksheet>'
    )
    content_types = (
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        "</Types>"
    )
    root_rels = (
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        "</Relationships>"
    )
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", root_rels)
        zf.writestr("xl/workbook.xml", workbook)
        zf.writestr("xl/_rels/workbook.xml.rels", rels)
        zf.writestr("xl/worksheets/sheet1.xml", sheet)
    return buf.getvalue()


_STUB = '<sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Module" sheetId="2" r:id=""/>'


@pytest.mark.parametrize("read_only", [False, True])
def test_empty_rid_sheet_is_skipped_on_load(read_only: bool) -> None:
    wb = wolfxl.load_workbook(io.BytesIO(_package(_STUB)), read_only=read_only)
    assert wb.sheetnames == ["Sheet1"]
    rows = [[c.value for c in row] for row in wb["Sheet1"].iter_rows()]
    assert rows == [["hello", 42]]


def test_empty_rid_sheet_is_reported_as_omission() -> None:
    book = NativeXlsxBook.open_from_bytes(_package(_STUB))
    assert book.sheet_names() == ["Sheet1"]
    assert book.skipped_sheets() == [
        {"name": "Module", "sheet_id": "2", "reason": "empty_relationship_id"}
    ]


@pytest.mark.parametrize("read_only", [False, True])
def test_non_empty_missing_rid_still_fails_with_missing_part(read_only: bool) -> None:
    stub = '<sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Ghost" sheetId="2" r:id="rId9"/>'
    with pytest.raises(OSError, match="missing workbook part: workbook relationship rId9"):
        wolfxl.load_workbook(io.BytesIO(_package(stub)), read_only=read_only)
