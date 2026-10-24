"""Sprint Κ Pod-α smoke tests for .xlsb / .xls / classify_file_format
read backends.

Pod-γ owns the curated parity matrix; this file only proves the new
pyclasses load, expose the right surface, and raise the right
exception when unsupported style accessors are touched.
"""

from __future__ import annotations

import pathlib

import pytest

FIXTURES = pathlib.Path(__file__).parent / "fixtures"

XLSX_FIXTURE = FIXTURES / "sprint_kappa_smoke.xlsx"
XLSB_FIXTURE = FIXTURES / "sprint_kappa_smoke.xlsb"
XLS_FIXTURE = FIXTURES / "sprint_kappa_smoke.xls"
ODS_FIXTURE = FIXTURES / "sprint_kappa_smoke.ods"


# ---------------------------------------------------------------------------
# classify_file_format
# ---------------------------------------------------------------------------


def test_classify_file_format_path_xlsx() -> None:
    from wolfxl._rust import classify_file_format

    assert classify_file_format(str(XLSX_FIXTURE)) == "xlsx"


def test_classify_file_format_path_xlsb() -> None:
    from wolfxl._rust import classify_file_format

    assert classify_file_format(str(XLSB_FIXTURE)) == "xlsb"


def test_classify_file_format_path_xls() -> None:
    from wolfxl._rust import classify_file_format

    assert classify_file_format(str(XLS_FIXTURE)) == "xls"


def test_classify_file_format_path_ods() -> None:
    from wolfxl._rust import classify_file_format

    assert classify_file_format(str(ODS_FIXTURE)) == "ods"


def test_classify_file_format_bytes_xlsx() -> None:
    from wolfxl._rust import classify_file_format

    assert classify_file_format(XLSX_FIXTURE.read_bytes()) == "xlsx"


def test_classify_file_format_bytes_xlsb() -> None:
    from wolfxl._rust import classify_file_format

    assert classify_file_format(XLSB_FIXTURE.read_bytes()) == "xlsb"


def test_classify_file_format_bytes_unknown() -> None:
    from wolfxl._rust import classify_file_format

    assert classify_file_format(b"not a spreadsheet") == "unknown"


# ---------------------------------------------------------------------------
# NativeXlsbBook
# ---------------------------------------------------------------------------


def test_native_xlsb_book_open_path() -> None:
    from wolfxl._rust import NativeXlsbBook

    book = NativeXlsbBook.open(str(XLSB_FIXTURE))
    names = book.sheet_names()
    assert names, "fixture must expose at least one sheet"
    values = book.read_sheet_values(names[0], None, True)
    assert isinstance(values, list)


def test_native_xlsb_book_open_from_bytes() -> None:
    from wolfxl._rust import NativeXlsbBook

    book = NativeXlsbBook.open_from_bytes(XLSB_FIXTURE.read_bytes())
    names = book.sheet_names()
    assert names
    values = book.read_sheet_values(names[0], None, True)
    assert isinstance(values, list)


@pytest.mark.parametrize("name", ("CalamineXlsBook", "CalamineXlsbBook", "CalamineOdsBook"))
def test_legacy_reader_exports_are_absent(name: str) -> None:
    import wolfxl._rust as rust

    assert not hasattr(rust, name)


# ---------------------------------------------------------------------------
# NativeXlsBook
# ---------------------------------------------------------------------------


def test_native_xls_book_open_path() -> None:
    from wolfxl._rust import NativeXlsBook

    book = NativeXlsBook.open(str(XLS_FIXTURE))
    names = book.sheet_names()
    assert names
    values = book.read_sheet_values(names[0])
    assert isinstance(values, list)
    assert book.source_path() == str(XLS_FIXTURE)
    assert book.opened_from_bytes() is False
    assert book.read_cell_value("Sheet1", "A1") == {"type": "string", "value": "name"}
    assert book.read_cell_value("Sheet1", "B2", True) == {"type": "number", "value": 1}
    with pytest.raises(ValueError, match="Unknown sheet: Missing"):
        book.read_sheet_bounds("Missing")


def test_native_xls_book_open_from_bytes() -> None:
    from wolfxl._rust import NativeXlsBook

    book = NativeXlsBook.open_from_bytes(XLS_FIXTURE.read_bytes())
    names = book.sheet_names()
    assert names
    values = book.read_sheet_values(names[0])
    assert isinstance(values, list)
    assert book.opened_from_bytes() is True
    assert book.source_path() is None
    assert book.read_cell_value("Sheet1", "C2") == {
        "type": "formula",
        "formula": "=B2*2",
        "value": "=B2*2",
    }


def test_native_xls_book_styles_raise() -> None:
    from wolfxl._rust import NativeXlsBook

    book = NativeXlsBook.open(str(XLS_FIXTURE))
    with pytest.raises(NotImplementedError, match="styles not supported"):
        book.read_cell_font(0, 0, "Sheet1")
    with pytest.raises(NotImplementedError, match="\\.xls"):
        book.read_cell_alignment(0, 0, "Sheet1")


# ---------------------------------------------------------------------------
# NativeXlsxBook.open_from_bytes (xlsx bytes path)
# ---------------------------------------------------------------------------


def test_native_xlsx_book_open_from_bytes() -> None:
    from wolfxl._rust import NativeXlsxBook

    book = NativeXlsxBook.open_from_bytes(XLSX_FIXTURE.read_bytes())
    names = book.sheet_names()
    assert names
    assert book.opened_from_bytes() is True


@pytest.mark.parametrize("value", (None, 42, bytearray(b"not a spreadsheet")))
def test_binding_classifier_rejects_unsupported_source_types(value) -> None:
    from wolfxl._rust import classify_file_format

    with pytest.raises(ValueError, match="^classify_file_format: expected str path or bytes$"):
        classify_file_format(value)
