"""Sprint Κ Pod-β — load_workbook format dispatcher tests.

Covers the unified entry point wired up in
``python/wolfxl/_loader.py`` + ``python/wolfxl/__init__.py``:

* path / bytes / bytearray / memoryview / BytesIO inputs all reach the
  same backend
* file-format sniffing distinguishes xlsx / xlsb / xls / encrypted /
  unknown
* xlsb / xls + (modify | read_only | password) raise
  NotImplementedError with the documented message
* cell.font / .fill / .border / .alignment / .number_format raise
  NotImplementedError on a non-xlsx workbook

xlsb / xls fixtures are produced by Pod-γ; tests gated on those skip
when the fixture isn't yet committed.
"""

from __future__ import annotations

import io
import tempfile
from pathlib import Path

import pytest

import wolfxl
from wolfxl import _workbook_sources
from wolfxl._loader import (
    _classify_bytes_python,
    _classify_path_python,
    classify_input,
)


FIXTURE_DIR = Path(__file__).parent / "parity" / "fixtures" / "synthgl_snapshot"
KAPPA_FIXTURES = Path(__file__).parent / "fixtures"


def test_from_bytes_modify_cleans_tempfile_on_open_failure(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    def boom(*args, **kwargs):  # noqa: ANN002, ANN003
        raise RuntimeError("patcher open failed")

    monkeypatch.setattr(tempfile, "tempdir", str(tmp_path))
    monkeypatch.setattr(_workbook_sources, "_xlsx_reader_class", lambda *args, **kwargs: object)
    monkeypatch.setattr(_workbook_sources, "from_patcher", boom)

    with pytest.raises(RuntimeError, match="patcher open failed"):
        _workbook_sources.from_bytes(wolfxl.Workbook, b"not an xlsx", modify=True)

    assert list(tmp_path.iterdir()) == []


def _first_xlsx() -> Path:
    """Pick a real xlsx fixture from the parity tree."""
    matches = list(FIXTURE_DIR.rglob("*.xlsx"))
    if not matches:
        pytest.skip("no xlsx fixtures available")
    return matches[0]


# ---------------------------------------------------------------------------
# classify_input + classify helpers
# ---------------------------------------------------------------------------


def test_classify_path_xlsx() -> None:
    fix = _first_xlsx()
    fmt, data, path = classify_input(str(fix))
    assert fmt == "xlsx"
    assert data is None
    assert path == str(fix)


def test_classify_pathlib_xlsx() -> None:
    fix = _first_xlsx()
    fmt, data, path = classify_input(fix)
    assert fmt == "xlsx"
    assert path == str(fix)


def test_classify_bytes_xlsx() -> None:
    fix = _first_xlsx()
    blob = fix.read_bytes()
    fmt, data, path = classify_input(blob)
    assert fmt == "xlsx"
    assert path is None
    assert data == blob


def test_classify_bytesio_xlsx() -> None:
    fix = _first_xlsx()
    bio = io.BytesIO(fix.read_bytes())
    fmt, data, path = classify_input(bio)
    assert fmt == "xlsx"
    assert path is None
    assert data is not None
    assert len(data) > 0


def test_classify_unknown_bytes_returns_unknown() -> None:
    assert _classify_bytes_python(b"hello world this is not a spreadsheet") == "unknown"
    assert _classify_bytes_python(b"") == "unknown"


def test_classify_xls_signature() -> None:
    """OLE CFB signature with no EncryptedPackage → xls."""
    head = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 4096
    assert _classify_bytes_python(head) == "xls"


def test_classify_encrypted_signature() -> None:
    """OLE CFB signature with EncryptedPackage stream → encrypted."""
    head = (
        b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
        + b"\x00" * 64
        + b"E\x00n\x00c\x00r\x00y\x00p\x00t\x00e\x00d\x00P\x00a\x00c\x00k\x00a\x00g\x00e\x00"
        + b"\x00" * 64
    )
    assert _classify_bytes_python(head) == "encrypted"


def test_classify_input_rejects_other_types() -> None:
    with pytest.raises(TypeError, match="load_workbook source must be"):
        classify_input(12345)
    with pytest.raises(TypeError, match="load_workbook source must be"):
        classify_input([1, 2, 3])


def test_classify_input_rejects_non_bytes_filelike() -> None:
    bio = io.StringIO("hello")
    with pytest.raises(TypeError, match="expected bytes"):
        classify_input(bio)


# ---------------------------------------------------------------------------
# load_workbook unified entry point
# ---------------------------------------------------------------------------


def test_load_workbook_path_round_trips() -> None:
    fix = _first_xlsx()
    wb = wolfxl.load_workbook(str(fix))
    assert wb._format == "xlsx"
    assert len(wb.sheetnames) >= 1


def test_load_workbook_bytes_matches_path() -> None:
    fix = _first_xlsx()
    wb_path = wolfxl.load_workbook(str(fix))
    wb_bytes = wolfxl.load_workbook(fix.read_bytes())
    assert wb_bytes.sheetnames == wb_path.sheetnames
    assert wb_bytes._format == "xlsx"


def test_load_workbook_bytearray_matches_path() -> None:
    fix = _first_xlsx()
    wb = wolfxl.load_workbook(bytearray(fix.read_bytes()))
    assert wb.sheetnames == wolfxl.load_workbook(str(fix)).sheetnames


def test_load_workbook_memoryview_matches_path() -> None:
    fix = _first_xlsx()
    blob = fix.read_bytes()
    wb = wolfxl.load_workbook(memoryview(blob))
    assert wb.sheetnames == wolfxl.load_workbook(str(fix)).sheetnames


def test_load_workbook_bytesio_matches_path() -> None:
    fix = _first_xlsx()
    bio = io.BytesIO(fix.read_bytes())
    wb = wolfxl.load_workbook(bio)
    assert wb.sheetnames == wolfxl.load_workbook(str(fix)).sheetnames


def test_load_workbook_pathlib_matches_str() -> None:
    fix = _first_xlsx()
    wb1 = wolfxl.load_workbook(fix)
    wb2 = wolfxl.load_workbook(str(fix))
    assert wb1.sheetnames == wb2.sheetnames


# ---------------------------------------------------------------------------
# Format-specific guards
# ---------------------------------------------------------------------------


def test_load_unknown_bytes_raises_value_error() -> None:
    # RFC-059 (Sprint Ο Pod-1E): unknown-format inputs now raise
    # ``InvalidFileException`` mirroring openpyxl's typed exception
    # hierarchy.  ``InvalidFileException`` is a plain ``Exception``
    # (not a ``ValueError``) to match openpyxl's contract.
    from wolfxl.utils.exceptions import InvalidFileException

    with pytest.raises(InvalidFileException, match="not determine file format"):
        wolfxl.load_workbook(b"this is definitely not a spreadsheet payload")


def test_load_xlsb_modify_raises() -> None:
    fix = KAPPA_FIXTURES / "sprint_kappa_smoke.xlsb"
    if not fix.exists():
        pytest.skip("xlsb fixture not yet committed (Pod-γ)")
    with pytest.raises(NotImplementedError, match="transcribe"):
        wolfxl.load_workbook(str(fix), modify=True)


def test_load_xlsb_read_only_raises() -> None:
    fix = KAPPA_FIXTURES / "sprint_kappa_smoke.xlsb"
    if not fix.exists():
        pytest.skip("xlsb fixture not yet committed (Pod-γ)")
    with pytest.raises(NotImplementedError, match="xlsx-only"):
        wolfxl.load_workbook(str(fix), read_only=True)


def test_load_xlsb_password_raises() -> None:
    fix = KAPPA_FIXTURES / "sprint_kappa_smoke.xlsb"
    if not fix.exists():
        pytest.skip("xlsb fixture not yet committed (Pod-γ)")
    with pytest.raises(NotImplementedError, match="xlsx-only"):
        wolfxl.load_workbook(str(fix), password="anything")


def test_load_xls_modify_raises_via_synthetic_bytes() -> None:
    """Even without a real .xls fixture we can drive the guard via raw bytes."""
    fake_xls = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 1024
    with pytest.raises(NotImplementedError, match="transcribe"):
        wolfxl.load_workbook(fake_xls, modify=True)


def test_load_xls_read_only_raises_via_synthetic_bytes() -> None:
    fake_xls = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 1024
    with pytest.raises(NotImplementedError, match="xlsx-only"):
        wolfxl.load_workbook(fake_xls, read_only=True)


def test_load_ods_raises_not_implemented() -> None:
    """An ODS-shaped zip is rejected with a clear pointer to openpyxl."""
    # Minimal ODS-like header: ZIP magic + 'opendocument.spreadsheet' near
    # the front, just enough to trip the sniffer.  We don't need a real
    # ODS file because the dispatcher never tries to read it.
    fake_ods = b"PK\x03\x04" + b"opendocument.spreadsheet" + b"\x00" * 1024
    with pytest.raises(NotImplementedError, match="OpenDocument"):
        wolfxl.load_workbook(fake_ods)


def test_load_encrypted_without_password_raises() -> None:
    """OLE CFB with EncryptedPackage stream but no password kwarg → ValueError."""
    fake_encrypted = (
        b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
        + b"\x00" * 64
        + b"E\x00n\x00c\x00r\x00y\x00p\x00t\x00e\x00d\x00P\x00a\x00c\x00k\x00a\x00g\x00e\x00"
        + b"\x00" * 1024
    )
    with pytest.raises(ValueError, match="OOXML-encrypted"):
        wolfxl.load_workbook(fake_encrypted)


# ---------------------------------------------------------------------------
# Workbook._format attribute
# ---------------------------------------------------------------------------


def test_workbook_format_write_mode_is_xlsx() -> None:
    wb = wolfxl.Workbook()
    assert wb._format == "xlsx"


def test_workbook_format_read_mode_is_xlsx() -> None:
    fix = _first_xlsx()
    wb = wolfxl.load_workbook(str(fix))
    assert wb._format == "xlsx"


def test_workbook_format_modify_mode_is_xlsx() -> None:
    fix = _first_xlsx()
    wb = wolfxl.load_workbook(str(fix), modify=True)
    assert wb._format == "xlsx"


def test_workbook_format_bytes_path_is_xlsx() -> None:
    fix = _first_xlsx()
    wb = wolfxl.load_workbook(fix.read_bytes())
    assert wb._format == "xlsx"


# ---------------------------------------------------------------------------
# Cell style guards on non-xlsx workbooks
#
# We synthesise the non-xlsx state by flipping ``wb._format`` on a real
# xlsx workbook.  This validates the Python-layer guard without
# requiring Pod-α's binary backends to be wired up.  When Pod-α + Pod-γ
# land, the smoke fixture-driven tests above also exercise this path
# end-to-end.
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("fmt", ["xls"])
@pytest.mark.parametrize("attr", ["font", "fill", "border", "alignment", "number_format"])
def test_cell_style_accessor_raises_on_non_xlsx(fmt: str, attr: str) -> None:
    fix = _first_xlsx()
    wb = wolfxl.load_workbook(str(fix))
    wb._format = fmt
    ws = wb.active
    cell = ws["A1"]
    with pytest.raises(NotImplementedError, match="xlsx-only"):
        getattr(cell, attr)


def test_xlsb_style_access_is_allowed_when_native_backend_supports_it() -> None:
    """Native .xlsb reads expose style getters instead of the old guard."""
    fixture = Path("tests/parity/fixtures/xlsb/dates.xlsb")
    wb = wolfxl.load_workbook(str(fixture))
    cell = wb.active["A1"]
    assert wb._rust_reader.__class__.__name__ == "NativeXlsbBook"  # noqa: SLF001
    assert cell.number_format is not None
    assert cell.font is not None


@pytest.mark.parametrize("fmt", ["xlsb", "xls"])
def test_cell_value_still_works_on_non_xlsx(fmt: str) -> None:
    """cell.value must keep working — only style accessors are gated."""
    fix = _first_xlsx()
    wb = wolfxl.load_workbook(str(fix))
    wb._format = fmt
    ws = wb.active
    # Just exercise the property — we don't care what value comes back.
    _ = ws["A1"].value


def test_is_date_returns_false_on_non_xlsx() -> None:
    """is_date defers to the value-type heuristic on non-xlsx workbooks."""
    fix = _first_xlsx()
    wb = wolfxl.load_workbook(str(fix))
    wb._format = "xlsb"
    ws = wb.active
    cell = ws["A1"]
    # is_date must not raise — it falls back to the value-type check.
    assert cell.is_date in (True, False)


# ---------------------------------------------------------------------------
# _classify_path_python sanity
# ---------------------------------------------------------------------------


def test_classify_path_python_handles_real_xlsx() -> None:
    fix = _first_xlsx()
    assert _classify_path_python(str(fix)) == "xlsx"


@pytest.mark.parametrize("kind", ("str", "path", "bytes", "bytearray", "memoryview", "filelike"))
def test_independent_native_xls_dispatch_source_forms(kind: str) -> None:
    fixture = KAPPA_FIXTURES / "sprint_kappa_smoke.xls"
    blob = fixture.read_bytes()
    source = {
        "str": str(fixture),
        "path": fixture,
        "bytes": blob,
        "bytearray": bytearray(blob),
        "memoryview": memoryview(blob),
        "filelike": io.BytesIO(blob),
    }[kind]
    # Literal values belong to the committed smoke corpus, not a second call to
    # the product parser masquerading as an independent oracle.
    with wolfxl.load_workbook(source, data_only=True) as workbook:
        assert type(workbook._rust_reader).__name__ == "NativeXlsBook"
        assert workbook._format == "xls"
        assert workbook.sheetnames == ["Sheet1"]
        assert workbook.active["A1"].value == "name"
        assert workbook.active["A2"].value == "alpha"
        assert workbook.active["B2"].value == 1
        with pytest.raises(NotImplementedError, match=r"\.xls"):
            _ = workbook.active["A1"].font


@pytest.mark.parametrize("kind", ("path", "bytes", "filelike"))
def test_real_ods_remains_refused_by_community_python(kind: str) -> None:
    fixture = KAPPA_FIXTURES / "sprint_kappa_smoke.ods"
    source = fixture if kind == "path" else fixture.read_bytes()
    if kind == "filelike":
        source = io.BytesIO(source)
    with pytest.raises(NotImplementedError, match="OpenDocument"):
        wolfxl.load_workbook(source)
