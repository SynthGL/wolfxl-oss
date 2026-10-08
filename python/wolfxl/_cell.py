"""Cell proxy — dispatches property access to the Rust backend."""

from __future__ import annotations

import datetime as _dt
import numbers as _numbers
import re
import inspect as _inspect
from copy import copy as _copy
from dataclasses import dataclass
from typing import TYPE_CHECKING, Any

from wolfxl._cell_annotations import (
    get_comment,
    get_hyperlink,
    set_comment,
    set_hyperlink,
)
from wolfxl._cell_payloads import (
    alignment_to_format_dict as alignment_to_format_dict,
    border_payload_to_border as _border_payload_to_border,
    border_to_rust_dict as border_to_rust_dict,
    fill_to_format_dict as fill_to_format_dict,
    font_to_format_dict as font_to_format_dict,
    format_to_alignment as _format_to_alignment,
    format_to_fill as _format_to_fill,
    format_to_font as _format_to_font,
    format_to_protection as _format_to_protection,
    payload_to_python as _payload_to_python,
    protection_to_format_dict as protection_to_format_dict,
    python_value_to_payload as python_value_to_payload,
)
from wolfxl._styles import Alignment, Border, Font, PatternFill
from wolfxl.styles.fills import GradientFill
from wolfxl.styles.protection import Protection
from wolfxl._utils import column_letter as _column_letter
from wolfxl._utils import rowcol_to_a1
from wolfxl._worksheet_rich_text import runs_payload_to_cellrichtext
from wolfxl.utils.cell import column_index_from_string, range_boundaries
from wolfxl.utils.datetime import from_excel, to_excel
from wolfxl.utils.exceptions import IllegalCharacterError
from wolfxl.utils.numbers import is_date_format
from wolfxl.styles.numbers import is_timedelta_format

if TYPE_CHECKING:
    from wolfxl._worksheet import Worksheet


# RFC-059 (Sprint Ο Pod-1E): OOXML-illegal control characters.
# The C0 controls 0x00–0x08, 0x0B, 0x0C, 0x0E–0x1F plus 0x7F are
# rejected by Excel's serializer.  Tab (0x09), newline (0x0A), and
# carriage return (0x0D) are allowed and pass through unchanged.
# Mirrors openpyxl's ``ILLEGAL_CHARACTERS_RE``.
ILLEGAL_CHARACTERS_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")
_STYLE_PAYLOAD_CACHE_DISABLED = object()
_STYLE_PAYLOAD_CACHE_CELL_LIMIT = 100_000
_EMPTY_STYLE_PAYLOAD: dict[str, Any] = {}
_EMPTY_STYLE_ENTRY = (
    _EMPTY_STYLE_PAYLOAD, _format_to_font(_EMPTY_STYLE_PAYLOAD),
    _format_to_fill(_EMPTY_STYLE_PAYLOAD), "General",
)
_PLAIN_VALUE_RETURN_TYPES = frozenset((type(None), bool, int, float, str))
_STYLE_PAYLOAD_KEYS = frozenset(
    {
        "bold",
        "italic",
        "underline",
        "strikethrough",
        "font_name",
        "font_size",
        "font_color",
        "bg_color",
        "gradient",
        "h_align",
        "v_align",
        "wrap",
        "rotation",
        "indent",
        "locked",
        "hidden",
        "number_format",
        "named_style",
    }
)


def _illegal_string_message(value: str) -> str:
    match = ILLEGAL_CHARACTERS_RE.search(value)
    codepoint = ord(match.group(0)) if match is not None else 0
    preview = repr(value[:80])
    if len(value) > 80:
        preview = f"{preview}... (length={len(value)})"
    return (
        f"Cell value {preview} contains character 0x{codepoint:02X}, which is "
        "not allowed in OOXML strings (control chars 0x00-0x08, 0x0B, 0x0C, "
        "0x0E-0x1F, 0x7F)"
    )


@dataclass
class _CellStyleSnapshot:
    font: Any
    fill: Any
    border: Any
    alignment: Any
    number_format: str | None
    protection: Any
    named_style: str | None
    pivotButton: bool = False  # noqa: N815 - openpyxl private style field
    quotePrefix: bool = False  # noqa: N815 - openpyxl private style field


class Cell:
    """Lightweight proxy for a single cell.

    In read mode, properties call into the Rust backend on first access.
    In write mode, assignments queue pending state flushed on ``save()``.
    """

    __slots__ = (
        "_ws",
        "_row",
        "_col",
        "_value",
        "_font",
        "_fill",
        "_border",
        "_border_authored",
        "_alignment",
        "_number_format",
        "_protection",
        "_format_payload",
        # Workbook-level named-style binding ("Highlight", "Heading 1", ...).
        # ``_UNSET`` means the cell has never been touched; ``None`` means
        # the cell was explicitly cleared. The flush layer threads the name
        # through `_named_style` in the format dict so the native writer can
        # look up the cellStyleXfs slot.
        "_named_style",
        "_value_dirty",
        "_format_dirty",
        "_style_slots_ready",
        "_explicit_data_type",
        "_value_is_plain",
        # Array / data-table formula metadata. Populated when ``cell.value`` is
        # assigned an
        # :class:`ArrayFormula` / :class:`DataTableFormula` instance,
        # or when an existing cell parses back as one of those types.
        # ``_formula_type`` is one of: ``None`` (plain), ``"array"``,
        # ``"dataTable"``.
        "_formula_type",
        "_array_ref",
        "_formula_text",
        "_dt_ca",
        "_dt_2d",
        "_dt_r",
        "_dt_r1",
        "_dt_r2",
        "_dt_del1",
        "_dt_del2",
        "_style_snapshot",
        "_comment",
        "_hyperlink",
    )

    def __init__(
        self,
        worksheet: Worksheet,
        row: int | None = None,
        column: int | None = None,
        value: Any = None,
        style_array: Any = None,
        *,
        col: int | None = None,
    ) -> None:
        if column is None:
            column = col
        if isinstance(column, str):
            column = column_index_from_string(column)
        self._ws = worksheet
        self._row = row
        self._col = column
        # Sentinel — None is a valid value so we use a special marker.
        self._value: Any = _UNSET
        self._font: Font | None | _Sentinel = _UNSET
        self._fill: PatternFill | GradientFill | None | _Sentinel = _UNSET
        self._border: Border | None | _Sentinel = _UNSET
        self._border_authored = False
        self._alignment: Alignment | None | _Sentinel = _UNSET
        self._number_format: str | None | _Sentinel = _UNSET
        self._protection: Protection | None | _Sentinel = _UNSET
        self._format_payload: Any = _UNSET
        self._named_style: str | None | _Sentinel = _UNSET
        self._value_dirty = False
        self._format_dirty = False
        self._style_slots_ready = True
        self._explicit_data_type: str | _Sentinel = _UNSET
        self._value_is_plain = False
        # None until the cell is identified as array / data-table either via
        # setter or on read-back.
        self._formula_type: str | None = None
        self._array_ref: str | None = None
        self._formula_text: str | None = None
        self._dt_ca: bool = False
        self._dt_2d: bool = False
        self._dt_r: bool = False
        self._dt_r1: str | None = None
        self._dt_r2: str | None = None
        self._dt_del1: bool = False
        self._dt_del2: bool = False
        self._style_snapshot: Any = _UNSET
        self._comment: Any = None
        self._hyperlink: Any = None
        if style_array is not None:
            self._style = style_array
        if value is not None:
            self.value = value

    def _workbook_like(self) -> Any:
        """Return the real WolfXL workbook or an openpyxl-style parent."""
        return getattr(self._ws, "_workbook", getattr(self._ws, "parent", None))

    def _rust_workbook(self) -> Any:
        """Return the real WolfXL workbook when this cell is backed by one."""
        return getattr(self._ws, "_workbook", None)

    def _mark_dirty(self) -> None:
        """Mark the cell dirty when the worksheet supports write tracking."""
        marker = getattr(self._ws, "_mark_dirty", None)
        if marker is not None:
            marker(self._row, self._col)
            return
        dirty = getattr(self._ws, "_dirty", None)
        if dirty is not None:
            dirty.add((self._row, self._col))

    def _mark_format_dirty(self) -> None:
        """Mark this cell as needing a style/format flush."""
        self._ensure_style_slots()
        self._format_dirty = True
        format_dirty_cells = getattr(self._ws, "_format_dirty_cells", None)
        if format_dirty_cells is not None:
            format_dirty_cells.add((self._row, self._col))

    def _ensure_style_slots(self) -> None:
        """Populate lazy style slots on fast-path cells before format flush."""
        if getattr(self, "_style_slots_ready", True):
            return
        self._font = _UNSET
        self._fill = _UNSET
        self._border = _UNSET
        self._alignment = _UNSET
        self._number_format = _UNSET
        self._protection = _UNSET
        self._named_style = _UNSET
        self._style_slots_ready = True

    def _dummy_collection_value(self, collection_name: str, default: Any) -> Any:
        """Read style slot 0 from an openpyxl dummy workbook, when present."""
        wb = self._workbook_like()
        collection = getattr(wb, collection_name, None)
        if collection is not None:
            try:
                if len(collection):
                    return collection[0]
            except TypeError:
                pass
        return default()

    @property
    def coordinate(self) -> str:
        """Return this cell's A1-style coordinate."""
        return rowcol_to_a1(self._row, self._col)

    @property
    def row(self) -> int:
        """Return this cell's 1-based row index."""
        return self._row

    @property
    def column(self) -> int:
        """Return this cell's 1-based column index."""
        return self._col

    @property
    def col_idx(self) -> int:
        """Return this cell's 1-based column index."""
        return self._col

    @property
    def column_letter(self) -> str:
        """Column letter (e.g. ``"A"``, ``"AA"``) — openpyxl alias."""
        return _column_letter(self._col)

    @property
    def parent(self) -> Worksheet:
        """The containing Worksheet — openpyxl alias."""
        return self._ws

    @property
    def base_date(self) -> Any:
        """Workbook epoch used for Excel serial date conversion."""
        wb = self._workbook_like()
        return getattr(wb, "excel_base_date", getattr(wb, "epoch", None))

    @property
    def encoding(self) -> str:
        """Cell text encoding marker for openpyxl compatibility."""
        wb = self._workbook_like()
        return getattr(wb, "encoding", getattr(self._ws, "encoding", "utf-8"))

    @property
    def internal_value(self) -> Any:
        """Internal cell value; WolfXL stores the openpyxl-facing value."""
        return self.value

    @property
    def pivotButton(self) -> bool:  # noqa: N802 - openpyxl public alias
        """Whether the cell displays a pivot-table field button."""
        snapshot = getattr(self, "_style_snapshot", _UNSET)
        if snapshot is not _UNSET:
            return bool(getattr(snapshot, "pivotButton", False))
        return False

    @property
    def quotePrefix(self) -> bool:  # noqa: N802 - openpyxl public alias
        """Whether the cell has Excel's quote-prefix flag."""
        snapshot = getattr(self, "_style_snapshot", _UNSET)
        if snapshot is not _UNSET:
            return bool(getattr(snapshot, "quotePrefix", False))
        return False

    @property
    def style_id(self) -> int:
        """Return the workbook style identifier for this cell."""
        if self._format_dirty:
            wb = self._workbook_like()
            styles = getattr(wb, "_cell_styles", None)
            if styles is not None:
                return styles.add(self._style_array_from_components(wb))
            return 1
        return self._read_style_id()

    @property
    def _style(self) -> Any:
        """Openpyxl private-style copy shim.

        Many existing openpyxl programs copy styles with
        ``dst._style = copy(src._style)``. WolfXL stores style components
        directly on cells, so this property exposes a lightweight snapshot
        that preserves the visible style payload across that idiom.
        """
        self._ensure_style_slots()
        snapshot = getattr(self, "_style_snapshot", _UNSET)
        if snapshot is not _UNSET:
            return snapshot
        self._style_snapshot = _CellStyleSnapshot(
            font=_copy(self.font),
            fill=_copy(self.fill),
            border=_copy(self.border),
            alignment=_copy(self.alignment),
            number_format=self.number_format,
            protection=_copy(self.protection),
            named_style=self.style,
        )
        return self._style_snapshot

    @_style.setter
    def _style(self, value: Any) -> None:
        self._require_xlsx_for_style("_style")
        self._ensure_style_slots()
        if isinstance(value, _CellStyleSnapshot):
            self._style_snapshot = value
            self._font = _copy(value.font)
            self._fill = _copy(value.fill)
            self._border = _copy(value.border)
            self._border_authored = True
            self._alignment = _copy(value.alignment)
            self._number_format = value.number_format
            self._protection = _copy(value.protection)
            self._named_style = value.named_style
            self._mark_format_dirty()
            self._mark_dirty()
            return
        from wolfxl.styles.styleable import StyleArray

        if isinstance(value, StyleArray):
            self._style_snapshot = value
            self._mark_format_dirty()
            self._mark_dirty()
            return
        raise TypeError(f"cell._style must be a WolfXL style snapshot, got {type(value).__name__}")

    def offset(self, row: int = 0, column: int = 0) -> Cell:
        """Return the cell ``row`` rows down and ``column`` columns right.

        Matches openpyxl's ``Cell.offset(row=0, column=0)`` signature. Negative
        offsets are allowed as long as the target row/col stays within Excel's
        1-based address space.
        """
        getter = getattr(self._ws, "_get_or_create_cell", None)
        if getter is not None:
            return getter(self._row + row, self._col + column)
        return self._ws.cell(row=self._row + row, column=self._col + column)

    def check_string(self, value: Any) -> str | None:
        """Validate a worksheet string using openpyxl's public helper rules."""
        if value is None:
            return None
        if not isinstance(value, str):
            value = str(value, self.encoding)
        value = str(value)
        if len(value) > 32767:
            value = value[:32767]
        if ILLEGAL_CHARACTERS_RE.search(value):
            raise IllegalCharacterError(f"{value} cannot be used in worksheets.")
        return value

    def check_error(self, value: Any) -> str:
        """Convert an error-like value to openpyxl's string representation."""
        try:
            return str(value)
        except UnicodeDecodeError:
            return "#N/A"

    @property
    def data_type(self) -> str:
        """openpyxl-compatible single-char type code.

        Maps to openpyxl's tags:
        - ``"s"``: string
        - ``"n"``: number (openpyxl also uses this for blank cells)
        - ``"b"``: boolean
        - ``"d"``: date / datetime
        - ``"f"``: formula
        - ``"e"``: error
        """
        if self._explicit_data_type is not _UNSET:
            return self._explicit_data_type
        stored_type = self._read_stored_data_type()
        if stored_type is not None:
            return stored_type
        from wolfxl._worksheet import _canonical_data_type

        canon = _canonical_data_type(self.value)
        mapping = {
            "string": "s",
            "number": "n",
            "boolean": "b",
            "datetime": "d",
            "date": "d",
            "formula": "f",
            "error": "e",
            "blank": "n",
        }
        return mapping.get(canon, "n")

    @data_type.setter
    def data_type(self, value: str) -> None:
        """Override the cell's storage type, matching openpyxl's mutable field."""
        if not isinstance(value, str):
            raise TypeError("Cell.data_type must be a string")
        self._explicit_data_type = value
        self._value_dirty = True
        dirty_values = getattr(self._ws, "_dirty_values", None)
        if dirty_values is not None:
            dirty_values.pop((self._row, self._col), None)
        self._mark_dirty()

    @property
    def has_style(self) -> bool:
        """True if any style attribute has been explicitly set on this cell.

        In read mode, checks whether the on-disk format carries any non-default
        style. In write mode, checks the dirty-flag sentinels so an unset cell
        reads as False and a cell with ``font = Font(bold=True)`` reads as True.
        """
        if self._format_dirty:
            return True
        font_value = getattr(self, "_font", _UNSET)
        fill_value = getattr(self, "_fill", _UNSET)
        border_value = getattr(self, "_border", _UNSET)
        alignment_value = getattr(self, "_alignment", _UNSET)
        number_format_value = getattr(self, "_number_format", _UNSET)
        protection_value = getattr(self, "_protection", _UNSET)
        font = font_value if font_value is not _UNSET else None
        fill = fill_value if fill_value is not _UNSET else None
        border = border_value if border_value is not _UNSET else None
        align = alignment_value if alignment_value is not _UNSET else None
        nfmt = number_format_value if number_format_value is not _UNSET else None
        prot = protection_value if protection_value is not _UNSET else None
        if font and font != Font():
            return True
        if fill and fill != PatternFill():
            return True
        if border and border != Border():
            return True
        if align and align != Alignment():
            return True
        if nfmt and nfmt != "General":
            return True
        if prot and prot != Protection():
            return True
        return False

    @property
    def is_date(self) -> bool:
        """True if the value is a date/datetime or the number format is a date."""
        value = self.value
        if hasattr(value, "year") and hasattr(value, "month"):
            return True
        if self.data_type == "d":
            return True
        if self.data_type != "n":
            return False
        # Binary formats may not expose style metadata. Fall back to the
        # value-type check above rather than raise from an introspection
        # accessor.
        wb_format = getattr(self._rust_workbook(), "_format", "xlsx")
        if wb_format != "xlsx":
            return False
        return is_date_format(self.number_format)

    @property
    def style(self) -> str | None:
        """Return the workbook-level named style bound to this cell.

        Pending in-memory writes win over the on-disk binding so a cell
        whose style has just been assigned reads back the same name even
        before the workbook is saved. ``None`` is returned for cells using
        the implicit Normal style.
        """
        self._ensure_style_slots()
        named_style = getattr(self, "_named_style", _UNSET)
        if named_style is not _UNSET:
            return named_style  # type: ignore[return-value]
        return self._read_named_style()

    @style.setter
    def style(self, value: str | None) -> None:
        """Bind this cell to a workbook-level named cell style.

        Args:
            value: Name of a previously-registered named style, or ``None``
                to clear an existing binding.
        """
        if value is None:
            self._ensure_style_slots()
            self._named_style = None
            self._mark_format_dirty()
            self._mark_dirty()
            return
        if not isinstance(value, str):
            raise TypeError(
                f"cell.style must be a string or None, got {type(value).__name__}"
            )
        wb = self._rust_workbook()
        self._ensure_style_slots()
        if wb is None:
            self._named_style = value
            self._mark_format_dirty()
            self._mark_dirty()
            return
        if not wb._has_named_style(value):  # noqa: SLF001
            raise ValueError(
                f"Named style {value!r} is not registered on the workbook. "
                "Call workbook.add_named_style(...) first or use the explicit "
                "font/fill/border accessors."
            )
        self._named_style = value
        style = wb._named_style_registry()[value]  # noqa: SLF001
        for style_attr, cell_attr in (
            ("font", "_font"),
            ("fill", "_fill"),
            ("border", "_border"),
            ("alignment", "_alignment"),
            ("protection", "_protection"),
        ):
            style_value = getattr(style, style_attr, None)
            if getattr(self, cell_attr, _UNSET) is _UNSET and style_value is not None:
                setattr(self, cell_attr, _copy(style_value))
                if cell_attr == "_border":
                    self._border_authored = True
        number_format = getattr(style, "number_format", "General")
        if (
            getattr(self, "_number_format", _UNSET) is _UNSET
            and number_format
            and number_format != "General"
        ):
            self._number_format = number_format
        self._mark_format_dirty()
        self._mark_dirty()

    # ------------------------------------------------------------------
    # T1 PR1: hyperlink / comment read access (write-mode setters land in PR4)
    #
    # Reads pull from per-worksheet lazy maps populated on first access.
    # Cells without a hyperlink/comment return None (matches openpyxl).
    # Setters raise NotImplementedError with a T1.5 pointer when the file
    # was opened via load_workbook(...) (no rust writer); write-mode
    # implementations land in PR4.
    # ------------------------------------------------------------------

    @property
    def hyperlink(self) -> Any:
        """Return the cell hyperlink, including pending unsaved edits."""
        if not hasattr(self._ws, "_pending_hyperlinks"):
            return getattr(self, "_hyperlink", None)
        return get_hyperlink(self, _UNSET)

    @hyperlink.setter
    def hyperlink(self, value: Any) -> None:
        """Set or clear the cell hyperlink.

        Args:
            value: ``Hyperlink`` instance, URL string, or ``None`` to delete
                the hyperlink on the next save.
        """
        if not hasattr(self._ws, "_pending_hyperlinks"):
            if isinstance(value, str):
                from wolfxl.worksheet.hyperlink import Hyperlink

                value = Hyperlink(ref=self.coordinate, target=value)
            elif value is not None:
                try:
                    value.ref = self.coordinate
                except AttributeError:
                    pass
            self._hyperlink = value
            return
        set_hyperlink(self, value)

    @property
    def comment(self) -> Any:
        """Return the cell comment, including pending unsaved edits."""
        if not hasattr(self._ws, "_pending_comments"):
            return getattr(self, "_comment", None)
        return get_comment(self, _UNSET)

    @comment.setter
    def comment(self, value: Any) -> None:
        """Set or clear the cell comment.

        Args:
            value: ``Comment`` instance, or ``None`` to delete the comment on
                the next save.
        """
        if value is not None:
            ws = self._ws
            pending_t = getattr(ws, "_pending_threaded_comments", {}).get(
                self.coordinate
            )
            if pending_t is not None:
                raise ValueError(
                    f"cell {self.coordinate} already has a threaded comment; "
                    "remove it before adding a legacy comment"
                )
        if not hasattr(self._ws, "_pending_comments"):
            old = getattr(self, "_comment", None)
            if value is None:
                if old is not None and getattr(old, "parent", None) is self:
                    old.unbind()
                self._comment = None
                return
            if getattr(value, "parent", None) is not None and value.parent is not self:
                value = _copy(value)
            if hasattr(value, "bind"):
                value.bind(self)
            else:
                value.parent = self
            self._comment = value
            return
        set_comment(self, value)

    @property
    def threaded_comment(self) -> Any:
        """Return the cell's top-level threaded comment, or ``None`` (G08).

        Pending unsaved edits take precedence over reader state, matching
        the legacy ``comment`` property's contract. When no pending edit
        exists, the workbook's reader-side cache is consulted; that cache
        is populated lazily on first access from
        ``xl/threadedComments/threadedCommentsN.xml``.
        """
        ws = self._ws
        pending = ws._pending_threaded_comments.get(self.coordinate, _UNSET)  # noqa: SLF001
        if pending is None:
            return None
        if pending is not _UNSET:
            return pending
        from wolfxl._worksheet_features import get_threaded_comments_map

        return get_threaded_comments_map(ws).get(self.coordinate)

    @threaded_comment.setter
    def threaded_comment(self, value: Any) -> None:
        """Attach or clear the cell's top-level threaded comment.

        Setting to ``None`` deletes the threaded comment (and replies) on
        the next save. Assigning a reply (one whose ``parent`` is non-None)
        raises — replies belong on the top-level's ``replies`` list.
        Assigning to a cell that already carries a legacy ``Comment``
        raises with a precise message; remove the legacy comment first.
        """
        from wolfxl.comments import ThreadedComment

        ws = self._ws
        wb = ws._workbook  # noqa: SLF001
        if wb._rust_writer is None and wb._rust_patcher is None:  # noqa: SLF001
            raise RuntimeError("cell.threaded_comment requires write or modify mode")

        coord = self.coordinate
        if value is None:
            ws._pending_threaded_comments[coord] = None  # noqa: SLF001
            return
        if not isinstance(value, ThreadedComment):
            raise TypeError(
                f"threaded_comment must be a ThreadedComment, got {type(value).__name__}"
            )
        if value.parent is not None:
            raise ValueError(
                "threaded_comment must be a top-level comment; "
                "reply via threaded_comment.replies.append(...)"
            )
        legacy = ws._pending_comments.get(coord)  # noqa: SLF001
        if legacy is not None and legacy is not _UNSET:
            raise ValueError(
                f"cell {coord} already has a legacy comment; "
                "remove it before adding a threaded comment"
            )
        ws._pending_threaded_comments[coord] = value  # noqa: SLF001

    @property
    def protection(self) -> Protection:
        """Return the resolved cell protection flags.

        Falls back to ``Protection()`` (Excel's default of
        ``locked=True, hidden=False``) when the cell carries no
        explicit ``<protection>`` override.
        """
        self._ensure_style_slots()
        self._require_xlsx_for_style("protection")
        if self._protection is _UNSET:
            self._protection = self._read_protection()
        return self._protection if self._protection is not None else Protection()  # type: ignore[return-value]

    @protection.setter
    def protection(self, val: Protection) -> None:
        """Set the cell protection flags.

        Args:
            val: Protection object with ``locked``/``hidden`` flags.
        """
        self._set_style_value("protection", "_protection", val)

    # ------------------------------------------------------------------
    # Value
    # ------------------------------------------------------------------

    @property
    def value(self) -> Any:
        """Return the cell value using openpyxl-compatible Python types."""
        if (
            self._value_is_plain
            and self._value is not _UNSET
            and getattr(self, "_formula_type", None) is None
            and type(self._value) in _PLAIN_VALUE_RETURN_TYPES
        ):
            return self._value
        # RFC-057: surface array / data-table formulas as the typed
        # instance regardless of what's been cached in ``_value``.
        # The metadata is populated either by the setter or by the
        # read-back path (parse_cell in the calamine backend tags the
        # cell post-read; pending-array map carries write-side state).
        pending_value = self._value_from_pending_formula()
        if pending_value is not _UNSET:
            return pending_value

        formula_value = self._value_from_formula_metadata()
        if formula_value is not _UNSET:
            return formula_value

        if self._value is _UNSET:
            self._value = self._read_value()
            # _read_value may have populated the formula metadata —
            # re-check after the read.
            formula_value = self._value_from_formula_metadata()
            if formula_value is not _UNSET:
                return formula_value
        # Sprint Ι Pod-α: when the workbook was opened with
        # ``rich_text=True``, surface ``CellRichText`` for cells whose
        # backing string carries `<r>` runs.  Default load mode mirrors
        # openpyxl 3.x, which flattens to plain ``str`` unless the user
        # opts in via the same flag.
        if isinstance(self._value, str):
            wb = self._rust_workbook()
            if getattr(wb, "_rich_text", False):
                rt = self.rich_text
                if rt is not None:
                    return rt
        wb_format = getattr(self._rust_workbook(), "_format", "xlsx")
        if (
            wb_format == "xlsx"
            and isinstance(self._value, _dt.datetime)
            and is_timedelta_format(self.number_format)
        ):
            epoch = self.base_date
            serial = to_excel(self._value, epoch) if epoch is not None else to_excel(self._value)
            if epoch is not None:
                return from_excel(serial, epoch, timedelta=True)
            return from_excel(serial, timedelta=True)
        return self._value

    @value.setter
    def value(self, val: Any) -> None:
        """Set the cell value and queue it for the next workbook save.

        Args:
            val: Scalar value, formula string, rich text object, array formula,
                data-table formula, or ``None``.
        """
        # Accept CellRichText pass-through: if the user assigns a
        # CellRichText, defer rich-text serialization to the writer.
        # Plain strings keep the existing fast path.
        ws = self._ws

        if val is None or isinstance(val, (bool, int, float, str)):
            self._clear_formula_metadata()
            if isinstance(val, str):
                if len(val) > 32767:
                    val = val[:32767]
                if ILLEGAL_CHARACTERS_RE.search(val):
                    raise IllegalCharacterError(_illegal_string_message(val))
            self._explicit_data_type = _UNSET
            self._value = val
            self._value_is_plain = True
            self._value_dirty = True
            try:
                ws._dirty.add((self._row, self._col))  # noqa: SLF001
                if self._row >= ws._next_append_row:  # noqa: SLF001
                    ws._next_append_row = self._row + 1  # noqa: SLF001
                if self._col > ws._max_col_idx:  # noqa: SLF001
                    ws._max_col_idx = self._col  # noqa: SLF001
                dirty_values = getattr(ws, "_dirty_values", None)
                if dirty_values is not None:
                    dirty_values[(self._row, self._col)] = val
            except (AttributeError, TypeError):
                self._mark_dirty()
            pending_rich_text = getattr(ws, "_pending_rich_text", None)
            if pending_rich_text is not None:
                pending_rich_text.pop((self._row, self._col), None)
            return

        from wolfxl.cell.cell import ArrayFormula, DataTableFormula
        from wolfxl.cell.rich_text import CellRichText  # local import — avoids cycles

        # RFC-057 — array / data-table formula assignment.
        if isinstance(val, ArrayFormula) or _is_array_formula_like(val):
            self._queue_array_formula(val)
            return

        if isinstance(val, DataTableFormula) or _is_data_table_formula_like(val):
            self._queue_data_table_formula(val)
            return

        # Plain assignment — clear any previous array / data-table
        # state so a former master cell can be replaced cleanly.
        self._clear_formula_metadata()
        val = self._bind_value(val)

        self._value = val
        self._value_dirty = True
        dirty_values = getattr(ws, "_dirty_values", None)
        if dirty_values is not None:
            dirty_values.pop((self._row, self._col), None)
        self._mark_dirty()

        # Pod-α: when a CellRichText is assigned, also stash it on the
        # worksheet's pending-rich-text map so the flush layer can pick
        # it up (write-mode and modify-mode both consume the same map).
        pending_rich_text = getattr(ws, "_pending_rich_text", None)
        if pending_rich_text is not None:
            if isinstance(val, CellRichText):
                pending_rich_text[(self._row, self._col)] = val
            else:
                # Clearing or replacing with plain — drop any prior rich entry.
                pending_rich_text.pop((self._row, self._col), None)

    def _bind_value(self, value: Any) -> Any:
        """Normalize a Python assignment and set openpyxl's data-type code."""
        from wolfxl.cell.rich_text import CellRichText

        self._explicit_data_type = _UNSET
        if value is None:
            return None

        if isinstance(value, CellRichText) or _is_cell_rich_text_like(value):
            return value

        if isinstance(value, bool):
            return value

        if isinstance(value, _numbers.Number):
            return value

        if isinstance(value, (_dt.datetime, _dt.date, _dt.time, _dt.timedelta)):
            from wolfxl.cell.cell import get_time_format

            current_format = self.number_format
            if not current_format or not is_date_format(current_format):
                self._number_format = get_time_format(type(value))
                self._mark_format_dirty()
            return value

        if isinstance(value, (str, bytes, bytearray)):
            if isinstance(value, bytearray):
                value = bytes(value)
            value = self.check_string(value)
            return value

        if hasattr(value, "year") and hasattr(value, "month"):
            from wolfxl.cell.cell import get_time_format

            current_format = self.number_format
            if not current_format or not is_date_format(current_format):
                self._number_format = get_time_format(type(value))
                self._mark_format_dirty()
            return value

        raise ValueError(
            f"Cannot convert object of type {type(value).__name__} to Excel"
        )

    def _value_from_pending_formula(self) -> Any:
        """Return pending array/data-table formula value or ``_UNSET``."""
        from wolfxl.cell.cell import ArrayFormula, DataTableFormula

        pending_map = getattr(self._ws, "_pending_array_formulas", None)
        if pending_map is None:
            return _UNSET
        pending = pending_map.get((self._row, self._col))
        if pending is None:
            return _UNSET
        kind, payload = pending
        if kind == "spill_child":
            return None
        if kind == "array":
            return ArrayFormula(payload["ref"], payload["text"])
        if kind == "data_table":
            return DataTableFormula(
                ref=payload["ref"],
                ca=payload.get("ca", False),
                dt2D=payload.get("dt2D", False),
                dtr=payload.get("dtr", False),
                r1=payload.get("r1"),
                r2=payload.get("r2"),
                del1=payload.get("del1", False),
                del2=payload.get("del2", False),
            )
        return _UNSET

    def _value_from_formula_metadata(self) -> Any:
        """Return read-side formula metadata value or ``_UNSET``."""
        from wolfxl.cell.cell import ArrayFormula, DataTableFormula

        formula_type = getattr(self, "_formula_type", None)
        if formula_type == "array":
            return ArrayFormula(
                getattr(self, "_array_ref", None) or "",
                getattr(self, "_formula_text", None) or "",
            )
        if formula_type == "dataTable":
            return DataTableFormula(
                ref=getattr(self, "_array_ref", None) or "",
                ca=getattr(self, "_dt_ca", False),
                dt2D=getattr(self, "_dt_2d", False),
                dtr=getattr(self, "_dt_r", False),
                r1=getattr(self, "_dt_r1", None),
                r2=getattr(self, "_dt_r2", None),
                del1=getattr(self, "_dt_del1", False),
                del2=getattr(self, "_dt_del2", False),
            )
        if formula_type == "array_child":
            return None
        return _UNSET

    def _clear_formula_metadata(self) -> None:
        """Clear array/data-table metadata and pending formula state."""
        self._formula_type = None
        self._array_ref = None
        self._formula_text = None
        self._dt_ca = False
        self._dt_2d = False
        self._dt_r = False
        self._dt_r1 = None
        self._dt_r2 = None
        self._dt_del1 = False
        self._dt_del2 = False
        pending_map = getattr(self._ws, "_pending_array_formulas", None)
        if pending_map is not None:
            pending_map.pop((self._row, self._col), None)

    def _queue_array_formula(self, val: Any) -> None:
        """Queue an array formula assignment for save."""
        ws = self._ws
        self._formula_type = "array"
        self._explicit_data_type = "f"
        self._array_ref = val.ref
        self._formula_text = val.text
        self._value = val
        self._value_dirty = True
        ws._dirty_values.pop((self._row, self._col), None)  # noqa: SLF001
        self._mark_dirty()
        ws._pending_array_formulas[(self._row, self._col)] = (  # noqa: SLF001
            "array",
            {"ref": val.ref, "text": val.text},
        )
        ws._plain_cell_fast_path = False  # noqa: SLF001
        # Populate placeholder entries for cells inside the spill range
        # (excluding the master). These show up as ``<c r="..."/>``
        # placeholders so Excel sees the spill area pre-populated.
        self._populate_spill_placeholders(val.ref)
        ws._pending_rich_text.pop((self._row, self._col), None)  # noqa: SLF001

    def _queue_data_table_formula(self, val: Any) -> None:
        """Queue a data-table formula assignment for save."""
        ws = self._ws
        self._formula_type = "dataTable"
        self._explicit_data_type = "f"
        self._array_ref = val.ref
        self._dt_ca = val.ca
        self._dt_2d = val.dt2D
        self._dt_r = val.dtr
        self._dt_r1 = val.r1
        self._dt_r2 = val.r2
        self._dt_del1 = val.del1
        self._dt_del2 = val.del2
        self._value = val
        self._value_dirty = True
        ws._dirty_values.pop((self._row, self._col), None)  # noqa: SLF001
        self._mark_dirty()
        ws._pending_array_formulas[(self._row, self._col)] = (  # noqa: SLF001
            "data_table",
            {
                "ref": val.ref,
                "ca": val.ca,
                "dt2D": val.dt2D,
                "dtr": val.dtr,
                "r1": val.r1,
                "r2": val.r2,
                "del1": val.del1,
                "del2": val.del2,
            },
        )
        ws._plain_cell_fast_path = False  # noqa: SLF001
        ws._pending_rich_text.pop((self._row, self._col), None)  # noqa: SLF001

    def _populate_spill_placeholders(self, ref: str) -> None:
        """Mark every non-master cell in ``ref`` as a spill child.

        RFC-057: when the user assigns ``cell.value = ArrayFormula(...)``,
        every cell inside the spill range becomes a placeholder so the
        ``cell.value`` getter on those cells returns ``None`` (mirroring
        openpyxl/Excel).  Only the master cell carries the actual
        formula text.
        """
        from wolfxl._utils import a1_to_rowcol  # noqa: SLF001

        ws = self._ws
        # Parse the ref ("A1:A10") into a 2-tuple of cells.
        if ":" not in ref:
            return  # single-cell array — nothing else to mark
        try:
            top_left, bottom_right = ref.split(":", 1)
            r1, c1 = a1_to_rowcol(top_left)
            r2, c2 = a1_to_rowcol(bottom_right)
        except Exception:  # noqa: BLE001
            return
        top, bottom = sorted((r1, r2))
        left, right = sorted((c1, c2))
        master_key = (self._row, self._col)
        for r in range(top, bottom + 1):
            for c in range(left, right + 1):
                if (r, c) == master_key:
                    continue
                ws._pending_array_formulas[(r, c)] = ("spill_child", {})  # noqa: SLF001

    # ------------------------------------------------------------------
    # Rich text
    # ------------------------------------------------------------------

    @property
    def rich_text(self) -> Any:
        """Structured rich-text runs for this cell, or ``None``.

        Returns a :class:`wolfxl.cell.rich_text.CellRichText` object
        when the on-disk cell carries `<r>` runs (either via the SST
        or as inline-string runs).  Returns ``None`` for plain-text
        cells, non-string types, and brand-new cells with no on-disk
        backing.

        Parity with openpyxl: openpyxl exposes the same data via
        ``Cell.value`` *only* when the workbook is loaded with
        ``rich_text=True``.  WolfXL goes one step further and always
        surfaces the structured representation through this side
        channel — defaulting ``Cell.value`` to flattened ``str`` so
        existing user code keeps working unchanged.
        """
        ws = self._ws
        # Pre-save visibility for write/modify-mode setters.
        pending_map = getattr(ws, "_pending_rich_text", None)
        if pending_map is not None:
            pending = pending_map.get((self._row, self._col))
            if pending is not None:
                return pending

        wb = self._rust_workbook()
        if wb is None:
            return None
        reader = getattr(wb, "_rust_reader", None)
        if reader is None:
            return None
        payload = reader.read_cell_rich_text(ws.title, self.coordinate)
        return runs_payload_to_cellrichtext(payload)

    @rich_text.setter
    def rich_text(self, val: Any) -> None:
        """Setter alias for ``cell.value = CellRichText(...)``.

        Lets users round-trip via ``cell.rich_text = ...`` even if they
        never touch ``cell.value`` directly — handy in code that wants
        to add/edit runs without disturbing other state.
        """
        self.value = val

    # ------------------------------------------------------------------
    # Style guard (Sprint Κ Pod-β)
    # ------------------------------------------------------------------

    def _require_xlsx_for_style(self, attr: str) -> None:
        """Raise NotImplementedError if this format cannot expose styles."""
        wb_format = getattr(self._rust_workbook(), "_format", "xlsx")
        if wb_format == "xlsb" and attr in {
            "font",
            "fill",
            "border",
            "alignment",
            "number_format",
        }:
            return
        if wb_format != "xlsx":
            raise NotImplementedError(
                f"cell.{attr} is xlsx-only; this workbook is .{wb_format}. "
                "Use .xlsx for style-aware reads."
            )

    # ------------------------------------------------------------------
    # Font
    # ------------------------------------------------------------------

    @property
    def font(self) -> Font:
        """Return the resolved cell font."""
        self._ensure_style_slots()
        if self._font is _UNSET:
            if not self._hydrate_basic_style_components():
                self._require_xlsx_for_style("font")
                self._font = self._read_font()
        return self._font  # type: ignore[return-value]

    @font.setter
    def font(self, val: Font) -> None:
        """Set the cell font.

        Args:
            val: Font object to apply to the cell.
        """
        self._set_style_value("font", "_font", val)

    # ------------------------------------------------------------------
    # Fill
    # ------------------------------------------------------------------

    @property
    def fill(self) -> PatternFill:
        """Return the resolved cell fill."""
        self._ensure_style_slots()
        if self._fill is _UNSET:
            if not self._hydrate_basic_style_components():
                self._require_xlsx_for_style("fill")
                self._fill = self._read_fill()
        return self._fill  # type: ignore[return-value]

    @fill.setter
    def fill(self, val: PatternFill) -> None:
        """Set the cell fill.

        Args:
            val: Pattern fill object to apply to the cell.
        """
        self._set_style_value("fill", "_fill", val)

    # ------------------------------------------------------------------
    # Border
    # ------------------------------------------------------------------

    @property
    def border(self) -> Border:
        """Return the resolved cell border."""
        self._ensure_style_slots()
        if self._border is _UNSET:
            self._require_xlsx_for_style("border")
            self._border = self._read_border()
        return self._border  # type: ignore[return-value]

    @border.setter
    def border(self, val: Border) -> None:
        """Set the cell border.

        Args:
            val: Border object to apply to the cell.
        """
        self._set_style_value("border", "_border", val)

    # ------------------------------------------------------------------
    # Alignment
    # ------------------------------------------------------------------

    @property
    def alignment(self) -> Alignment:
        """Return the resolved cell alignment."""
        self._ensure_style_slots()
        if self._alignment is _UNSET:
            self._require_xlsx_for_style("alignment")
            self._alignment = self._read_alignment()
        return self._alignment  # type: ignore[return-value]

    @alignment.setter
    def alignment(self, val: Alignment) -> None:
        """Set the cell alignment.

        Args:
            val: Alignment object to apply to the cell.
        """
        self._set_style_value("alignment", "_alignment", val)

    # ------------------------------------------------------------------
    # Number format
    # ------------------------------------------------------------------

    @property
    def number_format(self) -> str | None:
        """Return the resolved number format string."""
        self._ensure_style_slots()
        if self._number_format is _UNSET:
            self._require_xlsx_for_style("number_format")
            self._number_format = self._read_number_format()
        return self._number_format  # type: ignore[return-value]

    @number_format.setter
    def number_format(self, val: str | None) -> None:
        """Set the cell number format.

        Args:
            val: Number format code, or ``None`` to clear the cached format.
        """
        self._set_style_value("number_format", "_number_format", val)

    def _set_style_value(self, public_attr: str, storage_attr: str, value: Any) -> None:
        """Set a cached style value and mark the cell dirty."""
        ws = self._ws
        wb = getattr(ws, "_workbook", None)
        if getattr(wb, "_format", "xlsx") == "xlsb":
            raise NotImplementedError(
                f"cell.{public_attr} assignment is xlsx-only; .xlsb workbooks "
                "are read-only in WolfXL. Transcribe to .xlsx before editing styles."
            )
        self._ensure_style_slots()
        setattr(self, storage_attr, value)
        if storage_attr == "_border":
            self._border_authored = True
        key = (self._row, self._col)
        self._format_dirty = True
        format_dirty_cells = getattr(ws, "_format_dirty_cells", None)
        if format_dirty_cells is not None:
            format_dirty_cells.add(key)
        if wb is not None and getattr(wb, "_rust_reader", None) is not None:
            cache = getattr(
                ws,
                "_style_payload_cache",
                _STYLE_PAYLOAD_CACHE_DISABLED,
            )
            if cache is not _STYLE_PAYLOAD_CACHE_DISABLED:
                ws._style_payload_cache = _STYLE_PAYLOAD_CACHE_DISABLED  # noqa: SLF001
        dirty = getattr(ws, "_dirty", None)
        if dirty is None:
            self._mark_dirty()
        elif key not in dirty and not (
            wb is not None
            and getattr(wb, "_rust_reader", None) is None
            and key in getattr(ws, "_dirty_values", {})
        ):
            dirty.add(key)

    def _style_array_from_components(self, wb: Any) -> Any:
        from wolfxl.styles.cell_style import StyleArray
        from wolfxl.styles.numbers import BUILTIN_FORMATS_MAX_SIZE, BUILTIN_FORMATS_REVERSE

        style = StyleArray()
        font = getattr(self, "_font", _UNSET)
        fill = getattr(self, "_fill", _UNSET)
        border = getattr(self, "_border", _UNSET)
        alignment = getattr(self, "_alignment", _UNSET)
        protection = getattr(self, "_protection", _UNSET)
        number_format = getattr(self, "_number_format", _UNSET)
        if font is not _UNSET and font is not None:
            style.fontId = wb._fonts.add(font)  # noqa: SLF001
        if fill is not _UNSET and fill is not None:
            style.fillId = wb._fills.add(fill)  # noqa: SLF001
        if border is not _UNSET and border is not None:
            style.borderId = wb._borders.add(border)  # noqa: SLF001
        if alignment is not _UNSET and alignment is not None:
            style.alignmentId = wb._alignments.add(alignment)  # noqa: SLF001
        if protection is not _UNSET and protection is not None:
            style.protectionId = wb._protections.add(protection)  # noqa: SLF001
        if number_format is not _UNSET and number_format is not None:
            if number_format in BUILTIN_FORMATS_REVERSE:
                style.numFmtId = BUILTIN_FORMATS_REVERSE[number_format]
            else:
                style.numFmtId = (
                    wb._number_formats.add(number_format)  # noqa: SLF001
                    + BUILTIN_FORMATS_MAX_SIZE
                )
        return style

    # ------------------------------------------------------------------
    # Read helpers (dispatch to Rust via workbook)
    # ------------------------------------------------------------------

    def _read_value(self) -> Any:
        wb = self._rust_workbook()
        if wb is None:
            return None
        if getattr(wb, "_rust_reader", None) is None:
            return None
        # RFC-057: tag the cell with array / data-table metadata
        # before falling through to the regular payload read.  The
        # reader returns ``None`` for plain cells so the cost is one
        # extra dict-lookup at most.
        try:
            af_payload = wb._rust_reader.read_cell_array_formula(  # noqa: SLF001
                self._ws.title, self.coordinate,
            )
        except AttributeError:
            af_payload = None
        if af_payload is not None:
            kind = af_payload.get("kind")
            if kind == "array":
                self._formula_type = "array"
                self._array_ref = af_payload.get("ref")
                self._formula_text = af_payload.get("text", "")
            elif kind == "data_table":
                self._formula_type = "dataTable"
                self._array_ref = af_payload.get("ref")
                self._dt_ca = bool(af_payload.get("ca", False))
                self._dt_2d = bool(af_payload.get("dt2D", False))
                self._dt_r = bool(af_payload.get("dtr", False))
                self._dt_r1 = af_payload.get("r1")
                self._dt_r2 = af_payload.get("r2")
                self._dt_del1 = bool(af_payload.get("del1", False))
                self._dt_del2 = bool(af_payload.get("del2", False))
            elif kind == "spill_child":
                self._formula_type = "array_child"
        payload = wb._rust_reader.read_cell_value(  # noqa: SLF001
            self._ws.title, self.coordinate, getattr(wb, "_data_only", False),
        )
        return _payload_to_python(payload)

    def _read_stored_data_type(self) -> str | None:
        wb = self._rust_workbook()
        if wb is None:
            return None
        reader = getattr(wb, "_rust_reader", None)
        if reader is None:
            return None
        try:
            records = reader.read_sheet_records(
                self._ws.title,
                f"{self.coordinate}:{self.coordinate}",
                getattr(wb, "_data_only", False),
                False,
                True,
                True,
                True,
                False,
                False,
                False,
            )
        except AttributeError:
            return None
        mapping = {
            "string": "s",
            "number": "n",
            "boolean": "b",
            "datetime": "d",
            "date": "d",
            "formula": "f",
            "error": "e",
            "blank": "n",
        }
        for record in records:
            return mapping.get(record.get("data_type"))
        return None

    def _read_format_payload(self) -> Any:
        payload = getattr(self, "_format_payload", _UNSET)
        if payload is not _UNSET:
            return payload
        wb = self._rust_workbook()
        reader = getattr(wb, "_rust_reader", None) if wb is not None else None
        if reader is None:
            self._format_payload = None
            return None
        cache = _style_payload_cache_for(self._ws)
        if cache is not _STYLE_PAYLOAD_CACHE_DISABLED:
            entry = cache.get((self._row, self._col), _EMPTY_STYLE_ENTRY)
            if entry is not _STYLE_PAYLOAD_CACHE_DISABLED:
                self._format_payload = entry[0]
                return self._format_payload
            _disable_style_payload_cache(self._ws)
        read_cell_format_rc = getattr(reader, "read_cell_format_rc", None)
        if read_cell_format_rc is not None:
            self._format_payload = read_cell_format_rc(
                self._ws.title,
                self._row,
                self._col,
            )
            return self._format_payload
        self._format_payload = reader.read_cell_format(self._ws.title, self.coordinate)
        return self._format_payload

    def _read_font(self) -> Font:
        wb = self._rust_workbook()
        if wb is None:
            return self._dummy_collection_value("_fonts", Font)
        reader = getattr(wb, "_rust_reader", None)
        if reader is None:
            return Font()
        payload = self._read_format_payload()
        return _format_to_font(payload)

    def _read_fill(self) -> PatternFill | GradientFill:
        wb = self._rust_workbook()
        if wb is None:
            return self._dummy_collection_value("_fills", PatternFill)
        reader = getattr(wb, "_rust_reader", None)
        if reader is None:
            return PatternFill()
        payload = self._read_format_payload()
        return _format_to_fill(payload)

    def _hydrate_basic_style_components(self) -> bool:
        """Cache common style fields together from one native payload."""
        ws = self._ws
        cache = getattr(ws, "_style_payload_cache", None)
        if cache is None or cache is _STYLE_PAYLOAD_CACHE_DISABLED:
            wb = self._rust_workbook()
            if wb is None or getattr(wb, "_format", "xlsx") != "xlsx":
                return False
            if getattr(wb, "_rust_reader", None) is None:
                return False
            if cache is None:
                cache = _style_payload_cache_for(ws)
        entry = (cache.get((self._row, self._col), _EMPTY_STYLE_ENTRY)
                 if cache is not _STYLE_PAYLOAD_CACHE_DISABLED
                 else _STYLE_PAYLOAD_CACHE_DISABLED)
        if entry is _STYLE_PAYLOAD_CACHE_DISABLED:
            _disable_style_payload_cache(ws)
            payload = self._read_format_payload()
            font, fill, number_format = _basic_style_components_for_payload(ws, payload)
        else:
            self._format_payload, font, fill, number_format = entry
        if self._font is _UNSET:
            self._font = font
        if self._fill is _UNSET:
            self._fill = fill
        if self._number_format is _UNSET and _worksheet_is_known_unmerged(self._ws):
            self._number_format = number_format
        return True

    def _read_border(self) -> Border:
        wb = self._rust_workbook()
        if wb is None:
            return self._dummy_collection_value("_borders", Border)
        reader = getattr(wb, "_rust_reader", None)
        if reader is None:
            return Border()
        from wolfxl._worksheet_collections import _merged_border_for_cell

        merged_border = _merged_border_for_cell(self._ws, self.row, self.column)
        if merged_border is not None:
            return merged_border
        payload = reader.read_cell_border(
            self._ws.title, self.coordinate,
        )
        return _border_payload_to_border(payload)

    def _read_alignment(self) -> Alignment:
        wb = self._rust_workbook()
        if wb is None:
            return self._dummy_collection_value("_alignments", Alignment)
        reader = getattr(wb, "_rust_reader", None)
        if reader is None:
            return Alignment()
        payload = self._read_format_payload()
        return _format_to_alignment(payload)

    def _read_number_format(self) -> str | None:
        wb = self._rust_workbook()
        if wb is None:
            return "General"
        reader = getattr(wb, "_rust_reader", None)
        if reader is None:
            return "General"
        ws = self._ws  # noqa: SLF001
        may_have_merges = (
            not getattr(ws, "_merged_ranges_loaded", False)
            or bool(getattr(ws, "_merged_ranges", ()))
            or bool(getattr(ws, "_collection_merged_ranges", ()))
        )
        if may_have_merges and _is_merged_subordinate(self):
            return None
        payload = self._read_format_payload()
        if isinstance(payload, dict):
            return payload.get("number_format") or "General"
        return "General"

    def _read_protection(self) -> Protection | None:
        wb = self._rust_workbook()
        if wb is None:
            return self._dummy_collection_value("_protections", Protection)
        reader = getattr(wb, "_rust_reader", None)
        if reader is None:
            return None
        payload = self._read_format_payload()
        return _format_to_protection(payload)

    def _read_named_style(self) -> str | None:
        wb = self._rust_workbook()
        if wb is None:
            return None
        reader = getattr(wb, "_rust_reader", None)
        if reader is None:
            return None
        payload = self._read_format_payload()
        if isinstance(payload, dict):
            name = payload.get("named_style")
            if isinstance(name, str) and name:
                return name
        return None

    def _read_style_id(self) -> int:
        wb = self._rust_workbook()
        if wb is None:
            return 0
        reader = getattr(wb, "_rust_reader", None)
        if reader is None:
            return 0
        try:
            records = reader.read_sheet_records(
                self._ws.title,
                f"{self.coordinate}:{self.coordinate}",
                getattr(wb, "_data_only", False),
                True,
                True,
                True,
                False,
                True,
                False,
                False,
            )
        except AttributeError:
            return 0
        for record in records:
            style_id = record.get("style_id")
            return int(style_id or 0)
        return 0

    def __repr__(self) -> str:
        """Return a compact debug representation for this cell."""
        title = getattr(self._ws, "title", None)
        if title:
            return f"<Cell {title!r}.{self.coordinate}>"
        return f"<Cell {self.coordinate}>"


def _style_payload_cache_for(ws: Worksheet) -> Any:
    """Return a worksheet-level cache of native style payloads when safe.

    Styled reads normally cross from Python into Rust once per cell. For clean
    native-reader worksheets, we can build each distinct style payload and its
    common components once, then reuse that tuple for ``cell.font`` /
    ``cell.fill`` / ``cell.number_format`` access. Larger coordinate rectangles
    use bounded windows rather than disabling sharing for the whole sheet.
    """
    cache = getattr(ws, "_style_payload_cache", None)
    if cache is not None:
        return cache
    if not _can_build_style_payload_cache(ws):
        return _disable_style_payload_cache(ws)

    try:
        max_row = int(ws._max_row())  # noqa: SLF001
        max_col = int(ws._max_col())  # noqa: SLF001
    except Exception:
        return _disable_style_payload_cache(ws)
    if max_row <= 0 or max_col <= 0:
        ws._style_payload_cache = {}  # noqa: SLF001
        return ws._style_payload_cache  # noqa: SLF001
    wb = ws._workbook  # noqa: SLF001
    reader = wb._rust_reader  # noqa: SLF001
    if max_row * max_col > _STYLE_PAYLOAD_CACHE_CELL_LIMIT:
        from wolfxl._cell_style_window import StylePayloadWindow

        ws._style_payload_cache = StylePayloadWindow(  # noqa: SLF001
            reader, ws.title, max_row, max_col, _STYLE_PAYLOAD_CACHE_CELL_LIMIT,
            lambda record: _style_entry_from_record(ws, record), _STYLE_PAYLOAD_CACHE_DISABLED,
            workbook=wb,
        )
        return ws._style_payload_cache  # noqa: SLF001

    try:
        range_str = f"A1:{rowcol_to_a1(max_row, max_col)}"
        style_ids = reader.read_sheet_style_ids(ws.title, range_str)
        read_style = reader.read_format_for_style_id
        entries_by_id: dict[int, tuple[Any, ...]] = {}
        payloads: dict[tuple[int, int], tuple[Any, ...]] = {}
        for row, col, style_id in style_ids:
            style_id = int(style_id)
            entry = entries_by_id.get(style_id)
            if entry is None:
                entry = _style_entry_from_record(ws, read_style(style_id))
                entries_by_id[style_id] = entry
            payloads[(int(row), int(col))] = entry
    except Exception:
        return _disable_style_payload_cache(ws)

    ws._style_payload_cache = payloads  # noqa: SLF001
    return payloads


def _can_build_style_payload_cache(ws: Worksheet) -> bool:
    """Return True when the worksheet can reuse native style-id payloads."""
    wb = getattr(ws, "_workbook", None)
    reader = getattr(wb, "_rust_reader", None) if wb is not None else None
    if reader is None:
        return False
    if getattr(wb, "_format", "xlsx") != "xlsx":
        return False
    if not hasattr(reader, "read_sheet_style_ids"):
        return False
    if not hasattr(reader, "read_format_for_style_id"):
        return False
    if getattr(ws, "_dirty", None) or getattr(ws, "_format_dirty_cells", None):
        return False
    return True


def _disable_style_payload_cache(ws: Worksheet) -> object:
    """Mark the style cache disabled for this worksheet."""
    try:
        ws._style_payload_cache = _STYLE_PAYLOAD_CACHE_DISABLED  # noqa: SLF001
    except AttributeError:
        pass
    return _STYLE_PAYLOAD_CACHE_DISABLED


def _style_payload_from_record(record: dict[str, Any]) -> dict[str, Any]:
    """Keep only fields consumed by the Python style conversion helpers."""
    return {key: record[key] for key in _STYLE_PAYLOAD_KEYS if key in record}


def _style_entry_from_record(ws: Worksheet, record: dict[str, Any]) -> tuple[Any, ...]:
    """Build one payload/font/fill/number-format tuple per distinct style ID."""
    payload = _style_payload_from_record(record)
    if not payload:
        return _EMPTY_STYLE_ENTRY
    font, fill, number_format = _basic_style_components_for_payload(ws, payload)
    return payload, font, fill, number_format


def _basic_style_components_for_payload(
    ws: Worksheet,
    payload: Any,
) -> tuple[Font, PatternFill | GradientFill, str]:
    """Return cached font/fill/number-format objects for a read style payload."""
    cache = getattr(ws, "_style_component_cache", None)
    if cache is None:
        cache = {}
        try:
            ws._style_component_cache = cache  # noqa: SLF001
        except AttributeError:
            return (
                _format_to_font(payload),
                _format_to_fill(payload),
                _number_format_from_payload(payload),
            )

    key = id(payload)
    cached = cache.get(key)
    if cached is None:
        cached = (
            _format_to_font(payload),
            _format_to_fill(payload),
            _number_format_from_payload(payload),
        )
        cache[key] = cached
    return cached


Cell.__signature__ = _inspect.Signature(
    [
        _inspect.Parameter("worksheet", _inspect.Parameter.POSITIONAL_OR_KEYWORD),
        _inspect.Parameter(
            "row",
            _inspect.Parameter.POSITIONAL_OR_KEYWORD,
            default=None,
        ),
        _inspect.Parameter(
            "column",
            _inspect.Parameter.POSITIONAL_OR_KEYWORD,
            default=None,
        ),
        _inspect.Parameter(
            "value",
            _inspect.Parameter.POSITIONAL_OR_KEYWORD,
            default=None,
        ),
        _inspect.Parameter(
            "style_array",
            _inspect.Parameter.POSITIONAL_OR_KEYWORD,
            default=None,
        ),
    ]
)


# ======================================================================
# Sentinel type for lazy-init detection
# ======================================================================

class _Sentinel:
    """Marker to distinguish 'not yet loaded' from None."""

    _instance: _Sentinel | None = None

    def __new__(cls) -> _Sentinel:
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance

    def __repr__(self) -> str:
        """Return the sentinel's debug label."""
        return "<UNSET>"

    def __bool__(self) -> bool:
        return False


_UNSET = _Sentinel()


def _is_array_formula_like(value: Any) -> bool:
    """Accept openpyxl-shaped ArrayFormula objects across alias modules."""
    return (
        getattr(value, "t", None) == "array"
        and hasattr(value, "ref")
        and hasattr(value, "text")
    )


def _is_data_table_formula_like(value: Any) -> bool:
    """Accept openpyxl-shaped DataTableFormula objects across alias modules."""
    return getattr(value, "t", None) == "dataTable" and hasattr(value, "ref")


def _is_cell_rich_text_like(value: Any) -> bool:
    """Accept CellRichText objects imported through the openpyxl alias."""
    return type(value).__name__ == "CellRichText"


def _is_merged_subordinate(cell: Cell) -> bool:
    """Return True for cells inside a merged range but not its anchor."""
    ws = cell._ws  # noqa: SLF001
    try:
        ranges = _merged_range_refs(ws)
    except Exception:
        return False
    for ref in ranges:
        try:
            min_col, min_row, max_col, max_row = range_boundaries(str(ref))
        except Exception:
            continue
        if None in (min_col, min_row, max_col, max_row):
            continue
        if (
            int(min_row) <= cell.row <= int(max_row)
            and int(min_col) <= cell.column <= int(max_col)
        ):
            return not (cell.row == int(min_row) and cell.column == int(min_col))
    return False


def _worksheet_is_known_unmerged(ws: Any) -> bool:
    """Return True after merge metadata has been loaded and found empty."""
    return (
        bool(getattr(ws, "_merged_ranges_loaded", False))
        and not getattr(ws, "_merged_ranges", ())
        and not getattr(ws, "_collection_merged_ranges", ())
    )


def _number_format_from_payload(payload: Any) -> str:
    """Return the visible number format from a native style payload."""
    if isinstance(payload, dict):
        return payload.get("number_format") or "General"
    return "General"


def _merged_range_refs(ws: Any) -> set[str]:
    """Return raw merged-range refs without constructing public range objects."""
    stored = getattr(ws, "_merged_ranges", set())
    collection = getattr(ws, "_collection_merged_ranges", set())
    if stored or collection:
        return set(stored) | set(collection)
    if getattr(ws, "_merged_ranges_loaded", False):
        return set()
    reader = getattr(getattr(ws, "_workbook", None), "_rust_reader", None)
    if reader is None:
        return set()
    try:
        refs = {str(ref) for ref in reader.read_merged_ranges(ws._title)}  # noqa: SLF001
    except Exception:
        return set(getattr(ws, "_merged_ranges", set()))
    try:
        ws._merged_ranges.update(refs)  # noqa: SLF001
        ws._merged_ranges_loaded = True  # noqa: SLF001
    except AttributeError:
        pass
    return refs
