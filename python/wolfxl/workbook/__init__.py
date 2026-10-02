"""openpyxl.workbook compatibility.

Exposes :class:`DefinedNameDict` — the openpyxl-shape container for
``wb.defined_names``.
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING

from wolfxl._compat import _openpyxl_name_fallback
from wolfxl._workbook import Workbook
from wolfxl.workbook.defined_name import DefinedName, _DeletedDefinedName
from wolfxl.workbook.protection import FileSharing, WorkbookProtection
from wolfxl.workbook.views import BookView

if TYPE_CHECKING:
    from wolfxl._workbook import Workbook
    from wolfxl._worksheet import Worksheet


# A1-style cell reference: at least one letter (col) followed by digits (row).
_CELL_REF_RE = re.compile(r"^[A-Za-z]{1,3}[0-9]+$")


def _validate_defined_name(name: str) -> None:
    """Validate an Excel defined-name token.

    Mirrors Excel's name rules (subset enforced here):
        * Must be a non-empty string.
        * First character must be a letter or underscore (no digits, no
          leading whitespace).
        * Must not contain spaces or other whitespace.
        * Must not look like an A1-style cell reference (e.g. ``A1``,
          ``XFD1048576``).
        * Must not be the single letters ``R`` or ``C`` (reserved for
          R1C1 addressing).

    Raises ``ValueError`` on any violation. The message names the rule
    that was violated so users can fix the input.
    """
    if not isinstance(name, str) or not name:
        raise ValueError("DefinedName name must be a non-empty string")
    first = name[0]
    if not (first.isalpha() or first == "_"):
        raise ValueError(
            f"DefinedName name {name!r} must start with a letter or underscore"
        )
    if any(ch.isspace() for ch in name):
        raise ValueError(
            f"DefinedName name {name!r} must not contain whitespace"
        )
    if name in ("R", "C", "r", "c"):
        raise ValueError(
            f"DefinedName name {name!r} is reserved (R/C used in R1C1 addressing)"
        )
    if _CELL_REF_RE.match(name):
        raise ValueError(
            f"DefinedName name {name!r} looks like an A1-style cell reference"
        )


class DefinedNameDict(dict):
    """``dict``-subclass container for workbook defined names.

    Looks like a dict, but enforces that values are
    :class:`DefinedName` objects and that the key matches the DN's
    ``.name``. Having a subclass (rather than a plain dict) also lets
    write-mode callers queue a Rust-side flush on ``__setitem__`` in
    PR6 — the ``_wb`` back-ref is set by ``Workbook.defined_names``.
    """

    def __init__(self, *args: object, **kwargs: object) -> None:
        super().__init__(*args, **kwargs)
        # The workbook attaches itself after construction so the dict's
        # superclass init can stay plain. ``_wb`` is used by the PR6
        # write path; PR3 reads only need the container behavior.
        self._wb: Workbook | None = None
        # Set on ``ws.defined_names``. Sheet-scoped names get their
        # ``localSheetId`` from the worksheet position at save time.
        self._ws: Worksheet | None = None

    def __setitem__(self, key: str, value: DefinedName) -> None:
        if not isinstance(value, DefinedName):
            raise TypeError(
                f"value must be DefinedName, got {type(value).__name__}"
            )
        if value.value is not None:
            _validate_defined_name(key)
            _validate_defined_name(value.name)
        if key != value.name:
            raise ValueError(
                f"key '{key}' does not match DefinedName.name '{value.name}'"
            )
        super().__setitem__(key, value)
        # Queue a flush in both write and modify mode. Write mode ships
        # via ``_rust_writer.add_named_range`` during ``Workbook.save()``;
        # modify mode raises in ``save()`` with a T1.5 pointer — but we
        # still need to know a write was attempted, hence the queue.
        ws = self._ws
        if ws is not None:
            ws._pending_defined_names[key] = value  # noqa: SLF001
            return
        wb = self._wb
        if wb is not None:
            wb._pending_defined_names[key] = value  # noqa: SLF001

    def __delitem__(self, key: str) -> None:
        value = self[key]
        super().__delitem__(key)
        self._queue_delete(key, value)

    def pop(self, key: str, *default: object) -> DefinedName | object:
        if len(default) > 1:
            raise TypeError(f"pop expected at most 2 arguments, got {len(default) + 1}")
        if key not in self:
            if default:
                return default[0]
            raise KeyError(key)
        value = self[key]
        super().__delitem__(key)
        self._queue_delete(key, value)
        return value

    def add(self, value: DefinedName) -> None:
        """openpyxl helper: ``dn.add(DefinedName(...))`` is common."""
        self[value.name] = value

    def _queue_delete(self, key: str, value: DefinedName) -> None:
        ws = self._ws
        if ws is not None:
            ws._pending_defined_names.pop(key, None)  # noqa: SLF001
            if value.localSheetId is None:
                # Added since load and never saved: nothing to delete.
                return
            wb = ws._workbook  # noqa: SLF001
        else:
            wb = self._wb
            if wb is None:
                return
        local_sheet_id = value.localSheetId
        _drop_pending_defined_name(wb, key, local_sheet_id)
        if wb._rust_patcher is not None:  # noqa: SLF001
            wb._pending_defined_names[(key, local_sheet_id, "delete")] = (  # noqa: SLF001
                _DeletedDefinedName(key, local_sheet_id)
            )


def _drop_pending_defined_name(
    wb: Workbook,
    name: str,
    local_sheet_id: int | None,
) -> None:
    for pending_key, pending in list(wb._pending_defined_names.items()):  # noqa: SLF001
        if (
            getattr(pending, "name", None) == name
            and getattr(pending, "localSheetId", None) == local_sheet_id
        ):
            del wb._pending_defined_names[pending_key]  # noqa: SLF001


__all__ = [
    "BookView",
    "Workbook",
    "DefinedNameDict",
    "DefinedName",
    "FileSharing",
    "WorkbookProtection",
]

__getattr__ = _openpyxl_name_fallback(globals())
