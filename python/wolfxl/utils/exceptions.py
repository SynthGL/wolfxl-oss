"""Public exception types — openpyxl-shaped hierarchy.

User code that catches openpyxl's typed exceptions (notably
``IllegalCharacterError`` and ``CellCoordinatesException``) needs
the same names from wolfxl when migrating.  We mirror openpyxl's
type names but **do not** subclass openpyxl's hierarchy — the
classes are independent.  Both ``IllegalCharacterError`` and
``CellCoordinatesException`` subclass :class:`ValueError` so
existing ``except ValueError`` callsites keep working unchanged.

The internal raise sites (``utils/cell.py``,
``wolfxl.load_workbook``, the cell value setter) chain via
``raise X(msg) from e`` so the original cause is preserved on
``__cause__`` for callers that introspect it.

Reference: ``openpyxl.utils.exceptions`` (openpyxl 3.1.x).
"""

from __future__ import annotations

import builtins
import importlib.abc
import sys
from typing import Any

from wolfxl import _compat

_UPSTREAM_MODULE = "openpyxl.utils.exceptions"
_DEFERRED = _compat._real_openpyxl_pending() or (
    _compat._real_openpyxl_if_imported() is not None
    and _UPSTREAM_MODULE not in sys.modules
)


def _openpyxl_exception_base(name: str, fallback: type[BaseException]) -> type[BaseException]:
    if _DEFERRED:
        # Match the upstream heap-type layout so existing class references
        # can acquire the real base after openpyxl is imported.
        return type(f"_Upstream{name}Placeholder", (Exception,), {"__module__": __name__})
    exceptions_module: Any = sys.modules.get(_UPSTREAM_MODULE)
    candidate = getattr(exceptions_module, name, None)
    return candidate if isinstance(candidate, type) else fallback


def _coordinate_exception_bases() -> tuple[type[BaseException], ...]:
    upstream = _openpyxl_exception_base("CellCoordinatesException", ValueError)
    if upstream is ValueError:
        return (ValueError,)
    return (ValueError, upstream)


class InvalidFileException(Exception):
    """Raised when :func:`wolfxl.load_workbook` is asked to open
    something that isn't a valid xlsx / xlsb / xls / encrypted
    OOXML file.

    The format detector looks at leading magic bytes; anything
    that doesn't match a known archetype surfaces as this
    exception.  Mirrors openpyxl's
    ``openpyxl.utils.exceptions.InvalidFileException``.

    Note: this is a plain :class:`Exception`, not a
    :class:`ValueError` — openpyxl's contract is the same.
    Existing wolfxl callsites that previously raised generic
    ``ValueError`` for unknown file formats are rewrapped to this
    type with ``from`` chaining for backward compatibility.
    """


class IllegalCharacterError(ValueError):
    """Raised when a cell value contains characters illegal in
    OOXML strings.

    The OOXML spec rejects the C0 control characters
    ``\\x00`` – ``\\x08``, ``\\x0B``, ``\\x0C``, ``\\x0E`` –
    ``\\x1F``, plus ``\\x7F``.  Tab (``\\x09``), newline
    (``\\x0A``), and carriage return (``\\x0D``) are allowed.
    Surrogates are also rejected because Excel's serializer
    cannot round-trip them.

    Subclasses :class:`ValueError` so callers using ``except
    ValueError`` continue to function — code that wants the more
    specific type can switch to ``except IllegalCharacterError``.
    Mirrors openpyxl's
    ``openpyxl.utils.exceptions.IllegalCharacterError``.
    """


class CellCoordinatesException(*_coordinate_exception_bases()):
    """Raised when an invalid coordinate or range string is
    parsed (e.g. ``"AAAA1"``, ``"A0"``, ``"A1:"``).

    Triggered by the
    :func:`wolfxl.utils.cell.coordinate_to_tuple`,
    :func:`wolfxl.utils.cell.range_boundaries`, and
    :func:`wolfxl.utils.cell.column_index_from_string` parsers.

    Subclasses :class:`ValueError` so existing callsites that
    catch :class:`ValueError` are unaffected.  Mirrors openpyxl's
    ``openpyxl.utils.exceptions.CellCoordinatesException``.
    """


_cached_cell_coordinates_exception = getattr(
    builtins,
    "_wolfxl_CellCoordinatesException",
    None,
)
if isinstance(_cached_cell_coordinates_exception, type):
    CellCoordinatesException = _cached_cell_coordinates_exception
else:
    builtins._wolfxl_CellCoordinatesException = CellCoordinatesException


class _BindingLoader:
    """Bind the existing exception class when its upstream module finishes."""

    def __init__(self, loader: Any, binder: _UpstreamBinder) -> None:
        self._loader = loader
        self._binder = binder

    def create_module(self, spec: Any) -> Any:
        return self._loader.create_module(spec)

    def exec_module(self, module: Any) -> None:
        self._loader.exec_module(module)
        module.__spec__.loader = module.__loader__ = self._loader
        self._binder.bind(module)

    def __getattr__(self, name: str) -> Any:
        return getattr(self._loader, name)


class _UpstreamBinder(importlib.abc.MetaPathFinder):
    """One-shot hook preserving exception identity across either import order."""

    def find_spec(self, fullname: str, path: Any = None, target: Any = None) -> Any:
        if fullname != _UPSTREAM_MODULE:
            return None
        for finder in sys.meta_path:
            find = getattr(finder, "find_spec", None)
            if finder is self or find is None:
                continue
            spec = find(fullname, path, target)
            if spec is not None and spec.loader is not None:
                spec.loader = _BindingLoader(spec.loader, self)
                return spec
        return None

    def bind(self, module: Any) -> None:
        upstream = module.CellCoordinatesException
        CellCoordinatesException.__bases__ = (ValueError, upstream)
        sys.meta_path.remove(self)


if _DEFERRED:
    sys.meta_path.insert(0, _UpstreamBinder())


class ReadOnlyWorkbookException(Exception):
    """Raised when user code tries to mutate a workbook that was
    opened with ``read_only=True``.

    Wolfxl's read-only mode (Sprint Ι Pod-β) returns immutable
    cell proxies — assigning to ``cell.value`` raises this type.
    Mirrors openpyxl's
    ``openpyxl.utils.exceptions.ReadOnlyWorkbookException``.

    Note: kept as a plain :class:`Exception` (not
    :class:`RuntimeError`) for openpyxl name parity even though
    wolfxl's existing read-only setter raises ``RuntimeError``
    today.  Future patches may transition the read-only setter
    to raise this type directly.
    """


class WorkbookAlreadySaved(Exception):
    """Raised when :meth:`Workbook.save` is called twice on a
    write-only workbook.

    Wolfxl doesn't expose a separate write-only mode (the native
    writer streams internally), but the type exists for
    drop-in-import parity with openpyxl's
    ``openpyxl.utils.exceptions.WorkbookAlreadySaved``.  User
    code that catches this exception can be migrated unchanged.
    """


class NamedRangeException(Exception):
    """Compatibility name retained from openpyxl's exception surface."""


class SheetTitleException(Exception):
    """Raised for invalid worksheet title operations."""


__all__ = [
    "CellCoordinatesException",
    "IllegalCharacterError",
    "InvalidFileException",
    "NamedRangeException",
    "ReadOnlyWorkbookException",
    "SheetTitleException",
    "WorkbookAlreadySaved",
]
