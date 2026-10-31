"""WolfXL — fast, openpyxl-compatible Excel I/O backed by Rust.

Usage::

    from wolfxl import load_workbook, Workbook, Font, PatternFill

    # Read
    wb = load_workbook("data.xlsx")
    ws = wb["Sheet1"]
    print(ws["A1"].value, ws["A1"].font.bold)

    # Write
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "Hello"
    ws["A1"].font = Font(bold=True)
    wb.save("out.xlsx")
"""

from __future__ import annotations

from io import BytesIO
import os
import importlib
from types import ModuleType
from typing import IO, Dict, Union, cast
from zipfile import BadZipFile, ZIP_DEFLATED, ZipFile

from wolfxl._cell import Cell
from wolfxl.chartsheet import Chartsheet
from wolfxl._external_links import ExternalFileLink, ExternalLink
from wolfxl._rust import __version__, classify_format
from wolfxl._limits import LimitExceeded, Limits
from wolfxl.xml import DEFUSEDXML, LXML

# File-format detector (xlsx / xlsb / xls / ods / unknown), distinct from the
# existing ``classify_format`` cell-format classifier. Re-exported here so
# callers can use a stable ``wolfxl.classify_file_format(...)`` import without
# depending on the private ``wolfxl._rust`` module.
try:
    from wolfxl._rust import classify_file_format  # type: ignore[attr-defined]
except ImportError:  # pragma: no cover - wheels should expose this
    classify_file_format = None  # type: ignore[assignment]

try:
    from wolfxl._rust import (  # type: ignore[attr-defined]
        read_first_sheet_formula_summary as _read_first_sheet_formula_summary,
        read_first_sheet_value_summary as _read_first_sheet_value_summary,
    )
except ImportError:  # pragma: no cover - wheels should expose this
    _read_first_sheet_formula_summary = None  # type: ignore[assignment]
    _read_first_sheet_value_summary = None  # type: ignore[assignment]
from wolfxl._styles import Alignment, Border, Color, Font, PatternFill, Side
from wolfxl._workbook import CopyOptions, Workbook
from wolfxl._worksheet import Worksheet
from wolfxl.compat.numbers import NUMPY
from wolfxl.utils.exceptions import InvalidFileException

DEBUG = False


def read_first_sheet_value_summary(
    filename: str | os.PathLike[str],
) -> dict[str, float | int]:
    """Return row count and numeric checksum for the first worksheet."""
    if _read_first_sheet_value_summary is None:  # pragma: no cover
        raise ImportError("wolfxl._rust does not expose read_first_sheet_value_summary")
    return cast(
        Dict[str, Union[float, int]],
        _read_first_sheet_value_summary(os.fspath(filename)),
    )


def read_first_sheet_formula_summary(
    filename: str | os.PathLike[str],
) -> dict[str, float | int]:
    """Return row count, first-column checksum, and formula-cell count."""
    if _read_first_sheet_formula_summary is None:  # pragma: no cover
        raise ImportError("wolfxl._rust does not expose read_first_sheet_formula_summary")
    return cast(
        Dict[str, Union[float, int]],
        _read_first_sheet_formula_summary(os.fspath(filename)),
    )


__all__ = [
    "__version__",
    "Alignment",
    "Border",
    "Cell",
    "Chartsheet",
    "Color",
    "CopyOptions",
    "DEBUG",
    "DEFUSEDXML",
    "ExternalFileLink",
    "ExternalLink",
    "Font",
    "LXML",
    "LimitExceeded",
    "Limits",
    "NUMPY",
    "PatternFill",
    "Side",
    "Workbook",
    "Worksheet",
    "classify_file_format",
    "classify_format",
    "install_as_openpyxl",
    "load_workbook",
    "read_first_sheet_formula_summary",
    "read_first_sheet_value_summary",
    "uninstall_as_openpyxl",
]

_OPENPYXL_ROOT_MODULE_ALIASES = {
    "cell",
    "chart",
    "chartsheet",
    "comments",
    "compat",
    "constants",
    "descriptors",
    "drawing",
    "formatting",
    "formula",
    "packaging",
    "pivot",
    "reader",
    "styles",
    "utils",
    "workbook",
    "worksheet",
    "writer",
    "xml",
}

_SUPPORTED_OOXML_FORMATS = (".xlsx", ".xlsm", ".xltx", ".xltm")


def __getattr__(name: str) -> ModuleType:
    if name == "constants":
        module = importlib.import_module("wolfxl._constants")
    elif name in _OPENPYXL_ROOT_MODULE_ALIASES:
        module = importlib.import_module(f"wolfxl.{name}")
    else:
        raise AttributeError(name)
    globals()[name] = module
    return module


def load_workbook(
    filename: (
        str
        | os.PathLike[str]
        | bytes
        | bytearray
        | memoryview
        | IO[bytes]
    ),
    read_only: bool = False,
    data_only: bool = False,
    keep_vba: bool = False,
    keep_links: bool = True,
    modify: bool = False,
    permissive: bool = False,
    rich_text: bool = False,
    password: str | bytes | None = None,
    limits: Limits | None = None,
) -> Workbook:
    """Open a workbook for reading, streaming, or modify-mode saves.

    Args:
        filename: Path, bytes-like object, or binary file-like object
            containing an ``.xlsx``, ``.xlsb``, or ``.xls`` workbook.
        read_only: Enable the streaming row reader for ``.xlsx`` files.
            Streaming cells are immutable.
        data_only: Return cached formula results when present.
        keep_vba: Preserve VBA/macro package parts on save. This opens OOXML
            workbooks through the modify-mode patcher, matching openpyxl's
            ``keep_vba=True`` save contract.
        keep_links: Compatibility shim accepted for openpyxl-shaped call sites.
        modify: Enable read-modify-write mode for ``.xlsx`` files. Modified
            cells and supported metadata are saved while preserving unchanged
            workbook parts where possible.
        permissive: Fall back to worksheet relationships when workbook sheet
            metadata is malformed but recoverable.
        rich_text: Return structured rich-text values for cells that carry
            shared-string runs.
        password: Decrypt OOXML-encrypted ``.xlsx`` inputs with the optional
            ``wolfxl[encrypted]`` dependency.
        limits: Workbook-wide XLSX admission budgets. ``None`` uses
            :class:`Limits` defaults. Applies to full, read-only and modify loads,
            including decrypted XLSX inputs. Not supported for XLS or XLSB.

    Returns:
        A :class:`Workbook` in read, streaming, or modify mode.

    Raises:
        InvalidFileException: If the input format cannot be identified.
        LimitExceeded: If an OOXML input exceeds a configured resource budget.
        NotImplementedError: If the requested mode is unsupported for the
            detected format, such as modify mode for ``.xlsb``.
        ValueError: If an encrypted workbook needs a password or the supplied
            password cannot decrypt it.
    """
    from wolfxl._loader import classify_input
    from wolfxl._workbook_sources import open_workbook_source

    fmt, data, path = classify_input(filename)
    if limits is not None and fmt in ("xls", "xlsb"):
        raise NotImplementedError("limits= is supported for XLSX-family workbooks only")
    if limits is None:
        limits = Limits()
    elif not isinstance(limits, Limits):
        raise TypeError("limits= must be a wolfxl.Limits instance or None")

    # OOXML-encrypted .xlsx files arrive as OLE CFB envelopes that the
    # sniffer reports as ``"encrypted"``.  Once decrypted via
    # msoffcrypto they're plain xlsx, so we route them through the
    # encrypted constructor unconditionally; if the caller forgot to
    # pass ``password=``, surface a clear error before _from_encrypted
    # raises something more cryptic.
    if fmt == "encrypted":
        if password is None:
            raise ValueError(
                "this workbook is OOXML-encrypted; pass password= to "
                "load_workbook() (install with pip install wolfxl[encrypted])"
            )
        # _from_encrypted will produce a plain xlsx workbook.
        fmt = "xlsx"

    # Format-specific guards — surface clear errors *before* we try to
    # materialise a backend that doesn't exist for the requested mode.
    if keep_vba and not modify:
        modify = True

    if fmt in ("xlsb", "xls"):
        if modify:
            raise NotImplementedError(
                f".{fmt} files are read-only in wolfxl; load + transcribe to "
                ".xlsx then save: load via load_workbook(path), reconstruct "
                "as a fresh Workbook(), wb.save('out.xlsx')"
            )
        if read_only:
            raise NotImplementedError(
                f"streaming read_only mode is xlsx-only; .{fmt} files load "
                "whole-sheet"
            )
        if password is not None:
            raise NotImplementedError(
                f"password reads are xlsx-only (msoffcrypto-tool); .{fmt} "
                "encryption is out of scope"
            )

    if fmt == "ods":
        raise NotImplementedError(
            ".ods files are not supported by wolfxl; use openpyxl/odfpy "
            "for OpenDocument reads"
        )

    if fmt == "unknown":
        if _should_surface_bad_zip(filename):
            raise BadZipFile("File is not a zip file")
        source_kind = "path" if path else "bytes"
        raise InvalidFileException(
            f"could not determine file format from {source_kind}; "
            "expected xlsx/xlsb/xls"
        )

    wb = cast(
        Workbook,
        open_workbook_source(
            Workbook,
            fmt=fmt,
            path=path,
            data=data,
            password=password,
            data_only=data_only,
        keep_links=keep_links,
        keep_vba=keep_vba,
        permissive=permissive,
        modify=modify,
        read_only=read_only,
        limits=limits,
        ),
    )

    if keep_vba:
        try:
            wb.vba_archive = _copy_archive(filename)  # type: ignore[attr-defined]
        except Exception:
            pass

    wb._rich_text = rich_text  # noqa: SLF001
    wb._limits = limits  # noqa: SLF001
    return wb


open = load_workbook


def _should_surface_bad_zip(filename: object) -> bool:
    if hasattr(filename, "read") and not isinstance(filename, (str, bytes, os.PathLike)):
        return True
    if isinstance(filename, (str, os.PathLike)):
        return str(filename).lower().endswith(_SUPPORTED_OOXML_FORMATS)
    return False


def _copy_archive(filename: object) -> ZipFile:
    original_position = None
    if hasattr(filename, "tell") and hasattr(filename, "seek"):
        try:
            original_position = filename.tell()  # type: ignore[attr-defined]
            filename.seek(0)  # type: ignore[attr-defined]
        except Exception:
            original_position = None
    try:
        with ZipFile(filename, "r") as archive:  # type: ignore[arg-type]
            target_bytes = BytesIO()
            target = ZipFile(target_bytes, "a", ZIP_DEFLATED)
            for name in archive.namelist():
                target.writestr(name, archive.read(name))
            return target
    finally:
        if original_position is not None:
            try:
                filename.seek(original_position)  # type: ignore[attr-defined]
            except Exception:
                pass


def install_as_openpyxl(*, force: bool = True) -> ModuleType:
    """Install WolfXL as the runtime ``openpyxl`` import target."""
    from wolfxl.openpyxl_compat import install_as_openpyxl as install

    return install(force=force)


def uninstall_as_openpyxl() -> None:
    """Remove WolfXL's runtime ``openpyxl`` import alias."""
    from wolfxl.openpyxl_compat import uninstall_as_openpyxl as uninstall

    uninstall()
