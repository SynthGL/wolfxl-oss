"""Eager and streamed loads keep openpyxl 3.0.10 value types and XML line ends.

openpyxl reads numeric cell text with a decimal point or exponent as ``float``
and any other numeric text as ``int`` (``_cast_number``), and an XML parser
reads a literal CR or CR LF in text as LF (XML 1.0 section 2.11) while a
``&#13;`` character reference keeps its CR. Each workbook is built by hand so
the XML shape under test is exact.
"""

from __future__ import annotations

import io
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

import pytest

import wolfxl

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
SHEET_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
SST_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
BOOK_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"

SOURCES = ("path", "fileobj")
MODES = (False, True)


def _package(sheet_data: str, shared_strings: str | None = None) -> bytes:
    """Return a minimal one-sheet XLSX whose ``<sheetData>`` body is given."""
    overrides = [("/xl/workbook.xml", BOOK_CT), ("/xl/worksheets/sheet1.xml", SHEET_CT)]
    rels = [("rId1", "worksheet", "worksheets/sheet1.xml")]
    if shared_strings is not None:
        overrides.append(("/xl/sharedStrings.xml", SST_CT))
        rels.append(("rId2", "sharedStrings", "sharedStrings.xml"))
    parts = {
        "[Content_Types].xml": (
            f'<Types xmlns="{CT}">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            + "".join(f'<Override PartName="{name}" ContentType="{ct}"/>' for name, ct in overrides)
            + "</Types>"
        ),
        "_rels/.rels": (
            f'<Relationships xmlns="{PKG_REL}">'
            f'<Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/>'
            "</Relationships>"
        ),
        "xl/workbook.xml": (
            f'<workbook xmlns="{NS}" xmlns:r="{REL}"><sheets>'
            '<sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>'
        ),
        "xl/_rels/workbook.xml.rels": (
            f'<Relationships xmlns="{PKG_REL}">'
            + "".join(
                f'<Relationship Id="{rid}" Type="{REL}/{kind}" Target="{target}"/>'
                for rid, kind, target in rels
            )
            + "</Relationships>"
        ),
        "xl/worksheets/sheet1.xml": (
            f'<worksheet xmlns="{NS}"><sheetData>{sheet_data}</sheetData></worksheet>'
        ),
    }
    if shared_strings is not None:
        parts["xl/sharedStrings.xml"] = shared_strings
    buffer = io.BytesIO()
    with ZipFile(buffer, "w", ZIP_DEFLATED) as archive:
        for name, payload in parts.items():
            archive.writestr(name, payload)
    return buffer.getvalue()


def _open(data: bytes, tmp_path: Path, source: str, read_only: bool, data_only: bool):
    if source == "path":
        target = tmp_path / "book.xlsx"
        target.write_bytes(data)
        src: object = str(target)
    else:
        src = io.BytesIO(data)
    return wolfxl.load_workbook(src, read_only=read_only, data_only=data_only)


def _value_rows(
    data: bytes, tmp_path: Path, source: str, read_only: bool, data_only: bool
) -> list[list[str]]:
    """Rows of ``repr(value)``, so ``100`` and ``100.0`` stay distinct."""
    wb = _open(data, tmp_path, source, read_only, data_only)
    try:
        return [[repr(v) for v in row] for row in wb.active.iter_rows(values_only=True)]
    finally:
        wb.close()


def _cell_rows(
    data: bytes, tmp_path: Path, source: str, read_only: bool, data_only: bool
) -> list[list[str]]:
    wb = _open(data, tmp_path, source, read_only, data_only)
    try:
        return [[repr(c.value) for c in row] for row in wb.active.iter_rows()]
    finally:
        wb.close()


NUMBER_ROWS = (
    '<row r="1"><c r="A1"><v>100</v></c><c r="B1"><v>100.0</v></c>'
    '<c r="C1"><v>1E2</v></c><c r="D1"><v>1e2</v></c></row>'
    '<row r="2"><c r="A2"><v>+7</v></c><c r="B2"><v>-0.0</v></c>'
    '<c r="C2"><v>1.5</v></c><c r="D2"><v>0</v></c></row>'
    # A formula whose cached number carries a decimal point.
    '<row r="3"><c r="A3"><f>50*2</f><v>100.0</v></c><c r="B3"><v>3.0</v></c>'
    '<c r="C3"><v>-12</v></c><c r="D3"><v>2.5E+0</v></c></row>'
)
NUMBER_EXPECTED = [
    ["100", "100.0", "100.0", "100.0"],
    ["7", "-0.0", "1.5", "0"],
    ["100.0", "3.0", "-12", "2.5"],
]


@pytest.mark.parametrize("data_only", MODES)
@pytest.mark.parametrize("read_only", MODES)
@pytest.mark.parametrize("source", SOURCES)
def test_numeric_text_with_point_or_exponent_reads_as_float(
    tmp_path: Path, source: str, read_only: bool, data_only: bool
) -> None:
    data = _package(NUMBER_ROWS)
    if not data_only:
        # Formula cells read as their formula text unless cached values are asked for.
        expected = [list(row) for row in NUMBER_EXPECTED]
        expected[2][0] = "'=50*2'"
    else:
        expected = NUMBER_EXPECTED
    assert _value_rows(data, tmp_path, source, read_only, data_only) == expected
    assert _cell_rows(data, tmp_path, source, read_only, data_only) == expected


@pytest.mark.parametrize("source", SOURCES)
def test_sparse_sheet_numbers_read_as_float(tmp_path: Path, source: str) -> None:
    # A gap in the grid takes the sparse conversion path rather than the dense one.
    data = _package(
        '<row r="1"><c r="A1"><v>1.0</v></c><c r="C1"><v>2</v></c></row>'
        '<row r="3"><c r="B3"><v>4E0</v></c></row>'
    )
    assert _value_rows(data, tmp_path, source, False, True) == [
        ["1.0", "None", "2"],
        ["None", "None", "None"],
        ["None", "4.0", "None"],
    ]


FORMULA_STRING_ROWS = (
    '<row r="1"><c r="A1" t="str"><f>"x"</f><v>a\r\nb</v></c>'
    '<c r="B1" t="str"><f>"y"</f><v>a\rb</v></c>'
    '<c r="C1" t="str"><f>"z"</f><v>a&#13;&#10;b&#13;c</v></c></row>'
    '<row r="2"><c r="A2" t="inlineStr"><is><t>p\r\nq&#13;r</t></is></c>'
    '<c r="B2" t="s"><v>0</v></c><c r="C2" t="s"><v>1</v></c></row>'
)
FORMULA_STRING_SST = (
    f'<sst xmlns="{NS}" count="2" uniqueCount="2">'
    "<si><t>s\r\nt\ru</t></si>"
    "<si><t>s&#13;&#10;t&#13;u</t></si>"
    "</sst>"
)


@pytest.mark.parametrize("read_only", MODES)
@pytest.mark.parametrize("source", SOURCES)
def test_literal_cr_reads_as_lf_and_character_reference_keeps_cr(
    tmp_path: Path, source: str, read_only: bool
) -> None:
    data = _package(FORMULA_STRING_ROWS, shared_strings=FORMULA_STRING_SST)
    assert _value_rows(data, tmp_path, source, read_only, True) == [
        [repr("a\nb"), repr("a\nb"), repr("a\r\nb\rc")],
        [repr("p\nq\rr"), repr("s\nt\nu"), repr("s\r\nt\ru")],
    ]
