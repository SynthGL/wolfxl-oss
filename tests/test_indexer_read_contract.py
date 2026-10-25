"""Read contract a document indexer relies on, checked against openpyxl 3.0.10.

Each workbook is built by hand so the XML shape under test is exact: rows
and cells that omit ``r=``, shared strings with phonetic runs, time-only
number formats, cell and cached formula value types, and a wide sheet read
through ``read_only=True``.
"""

from __future__ import annotations

import io
import tracemalloc
from datetime import datetime, time
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

import pytest

import wolfxl

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
SHEET_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"

STYLES = (
    f'<styleSheet xmlns="{NS}">'
    '<numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd hh:mm:ss.000"/></numFmts>'
    '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
    '<fills count="2"><fill><patternFill patternType="none"/></fill>'
    '<fill><patternFill patternType="gray125"/></fill></fills>'
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    '<cellXfs count="3">'
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    '<xf numFmtId="21" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
    '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
    "</cellXfs>"
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
    "</styleSheet>"
)


def _package(
    sheet_data: str, shared_strings: str | None = None, dimension: str | None = None
) -> bytes:
    """Return a minimal one-sheet XLSX whose ``<sheetData>`` body is given."""
    dimension_xml = f'<dimension ref="{dimension}"/>' if dimension else ""
    overrides = [
        ("/xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"),
        ("/xl/worksheets/sheet1.xml", SHEET_CT),
        ("/xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"),
    ]
    workbook_rels = [
        ("rId1", "worksheet", "worksheets/sheet1.xml"),
        ("rId2", "styles", "styles.xml"),
    ]
    if shared_strings is not None:
        overrides.append(
            (
                "/xl/sharedStrings.xml",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
            )
        )
        workbook_rels.append(("rId3", "sharedStrings", "sharedStrings.xml"))
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
                for rid, kind, target in workbook_rels
            )
            + "</Relationships>"
        ),
        "xl/styles.xml": STYLES,
        "xl/worksheets/sheet1.xml": (
            f'<worksheet xmlns="{NS}">{dimension_xml}<sheetData>{sheet_data}</sheetData></worksheet>'
        ),
    }
    if shared_strings is not None:
        parts["xl/sharedStrings.xml"] = shared_strings
    buffer = io.BytesIO()
    with ZipFile(buffer, "w", ZIP_DEFLATED) as archive:
        for name, payload in parts.items():
            archive.writestr(name, payload)
    return buffer.getvalue()


SOURCES = ("path", "fileobj")
MODES = (False, True)


def _rows(
    data: bytes, tmp_path: Path, source: str, read_only: bool, data_only: bool = False
) -> list[tuple]:
    if source == "path":
        target = tmp_path / "book.xlsx"
        target.write_bytes(data)
        src: object = str(target)
    else:
        src = io.BytesIO(data)
    wb = wolfxl.load_workbook(src, read_only=read_only, data_only=data_only)
    try:
        return [tuple(row) for row in wb.active.iter_rows(values_only=True)]
    finally:
        wb.close()


@pytest.mark.parametrize("read_only", MODES)
@pytest.mark.parametrize("source", SOURCES)
def test_rows_and_cells_without_r_read_at_implicit_coordinates(
    tmp_path: Path, source: str, read_only: bool
) -> None:
    # A row without r= follows the previous row; a cell without r= follows
    # the previous cell in its row (ECMA-376 Part 1, 18.3.1.73 and 18.3.1.4).
    data = _package(
        "<row><c><v>1</v></c><c><v>2</v></c></row>"
        '<row r="3"><c r="B3"><v>3</v></c><c><v>4</v></c></row>'
        "<row><c><v>5</v></c></row>",
        dimension="A1:C4",
    )
    assert _rows(data, tmp_path, source, read_only) == [
        (1, 2, None),
        (None, None, None),
        (None, 3, 4),
        (5, None, None),
    ]


@pytest.mark.parametrize("read_only", MODES)
@pytest.mark.parametrize("source", SOURCES)
def test_shared_string_phonetic_runs_are_not_cell_text(
    tmp_path: Path, source: str, read_only: bool
) -> None:
    sst = (
        f'<sst xmlns="{NS}" count="2" uniqueCount="2">'
        '<si><t>東京都</t><rPh sb="0" eb="3"><t>トウキョウト</t></rPh>'
        '<phoneticPr fontId="0"/></si>'
        '<si><r><t>大阪</t></r><r><t>府</t></r>'
        '<rPh sb="0" eb="2"><t>オオサカ</t></rPh><phoneticPr fontId="0"/></si>'
        "</sst>"
    )
    data = _package(
        '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>',
        shared_strings=sst,
    )
    assert _rows(data, tmp_path, source, read_only) == [("東京都", "大阪府")]


@pytest.mark.parametrize("read_only", MODES)
@pytest.mark.parametrize("source", SOURCES)
def test_time_formats_and_sub_second_precision_match_openpyxl(
    tmp_path: Path, source: str, read_only: bool
) -> None:
    # Style 1 is built-in format 21 (h:mm:ss); style 2 a datetime with
    # milliseconds. openpyxl returns datetime.time for serials in [0, 1)
    # and rounds the day fraction to the millisecond.
    data = _package(
        '<row r="1"><c r="A1" s="1"><v>0</v></c></row>'
        '<row r="2"><c r="A2" s="1"><v>0.250005787037037</v></c></row>'
        '<row r="3"><c r="A3" s="2"><v>41224.99998836806</v></c></row>'
        '<row r="4"><c r="A4" s="2"><v>41224.5</v></c></row>'
    )
    assert _rows(data, tmp_path, source, read_only) == [
        (time(0, 0),),
        (time(6, 0, 0, 500000),),
        (datetime(2012, 11, 11, 23, 59, 58, 995000),),
        (datetime(2012, 11, 11, 12, 0),),
    ]


@pytest.mark.parametrize("read_only", MODES)
@pytest.mark.parametrize("source", SOURCES)
def test_cell_values_keep_openpyxl_types(tmp_path: Path, source: str, read_only: bool) -> None:
    # Cached formula results keep their stored type, raw CR and CR LF in
    # text read as LF (XML end-of-line handling), an empty <is> is "", an
    # empty shared-string <v/> is None, and an inline string ignores <v>.
    sst = f'<sst xmlns="{NS}" count="2" uniqueCount="2"><si><t>s0</t></si><si><t>c\rd</t></si></sst>'
    data = _package(
        '<row r="1"><c r="A1" t="b"><f>TRUE()</f><v>1</v></c>'
        '<c r="B1" t="str"><f>"007"</f><v>007</v></c>'
        '<c r="C1"><f>5/2</f><v>2.5</v></c><c r="D1" t="s"><v>0 </v></c></row>'
        '<row r="2"><c r="A2" t="inlineStr"><is><t></t></is></c>'
        '<c r="B2" t="inlineStr"><v>9</v><is><t>x</t></is></c><c r="C2" t="s"><v/></c></row>'
        '<row r="3"><c r="A3" t="inlineStr"><is><t>a\r\nb</t></is></c>'
        '<c r="B3" t="s"><v>1</v></c></row>',
        shared_strings=sst,
        dimension="A1:D3",
    )
    assert _rows(data, tmp_path, source, read_only, data_only=True) == [
        (True, "007", 2.5, "s0"),
        ("", "x", None, None),
        ("a\nb", "c\nd", None, None),
    ]


WIDE_ROWS = 2000
WIDE_EMPTY_CELLS = 16383
READ_ROWS = 61
PEAK_LIMIT = 32 * 1024 * 1024


@pytest.fixture(scope="module")
def wide_package(tmp_path_factory: pytest.TempPathFactory) -> Path:
    """2000 rows, each one value followed by 16383 empty cells."""
    target = tmp_path_factory.mktemp("wide") / "wide.xlsx"
    base = _package("@ROWS@")
    empties = "<c/>" * WIDE_EMPTY_CELLS
    with ZipFile(io.BytesIO(base)) as src, ZipFile(target, "w", ZIP_DEFLATED) as dst:
        for info in src.infolist():
            payload = src.read(info.filename)
            if info.filename != "xl/worksheets/sheet1.xml":
                dst.writestr(info.filename, payload)
                continue
            head, tail = payload.decode().split("@ROWS@")
            with dst.open(info.filename, "w", force_zip64=True) as sheet:
                sheet.write(head.encode())
                for row in range(1, WIDE_ROWS + 1):
                    sheet.write(f'<row r="{row}"><c r="A{row}"><v>{row}</v></c>{empties}</row>'.encode())
                sheet.write(tail.encode())
    return target


@pytest.mark.parametrize("source", SOURCES)
def test_read_only_streams_wide_sheet_for_path_and_file_object(
    wide_package: Path, source: str
) -> None:
    src: object = str(wide_package) if source == "path" else io.BytesIO(wide_package.read_bytes())
    tracemalloc.start()
    try:
        wb = wolfxl.load_workbook(src, read_only=True)
        first: list[object] = []
        for index, row in enumerate(wb.active.iter_rows(values_only=True)):
            first.append(row[0])
            if index + 1 == READ_ROWS:
                break
        wb.close()
        _, peak = tracemalloc.get_traced_memory()
    finally:
        tracemalloc.stop()
    assert first == list(range(1, READ_ROWS + 1))
    assert peak < PEAK_LIMIT, f"peak {peak / 2**20:.1f} MiB for {source} source"
