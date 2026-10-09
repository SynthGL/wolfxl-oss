"""MergedCell — placeholder for non-anchor cells in a merged range.

openpyxl exposes a ``MergedCell`` type whose ``.value`` is always
``None`` and whose value setter raises :class:`AttributeError`.
The anchor (top-left) cell of the merged range remains a regular
:class:`Cell`; every other coordinate inside the range becomes a
``MergedCell``.

Wolfxl manages merges differently (the rows/columns continue to
hold real :class:`Cell` proxies and the merge is tracked via the
worksheet-level ``merged_cells`` collection), so this class is
a thin compatibility shim — user code that constructs or
``isinstance``-checks ``MergedCell`` to detect non-anchor
positions can migrate unchanged.

Reference: ``openpyxl.cell.cell.MergedCell`` (openpyxl 3.1.x).
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from wolfxl._styles import Alignment, Border, Font, PatternFill
from wolfxl._utils import rowcol_to_a1
from wolfxl.styles.protection import Protection

if TYPE_CHECKING:
    from wolfxl._worksheet import Worksheet


class MergedCell:
    """Placeholder for cells inside a merged range that aren't
    the top-left anchor.

    ``value`` is always ``None``; assignment raises
    :class:`AttributeError` to match openpyxl's contract.
    """

    __slots__ = ("_parent", "_row", "_col", "_value")

    def __init__(
        self,
        parent: Worksheet | None,
        row: int | None = None,
        column: int | None = None,
    ) -> None:
        self._parent = parent
        self._row = row
        self._col = column
        self._value = None

    @property
    def parent(self) -> Worksheet | None:
        return self._parent

    @property
    def row(self) -> int:
        return self._row

    @row.setter
    def row(self, value: int) -> None:
        self._row = value

    @property
    def column(self) -> int:
        return self._col

    @column.setter
    def column(self, value: int) -> None:
        self._col = value

    @property
    def coordinate(self) -> str:
        return rowcol_to_a1(self._row, self._col)

    @property
    def value(self) -> Any:
        return None

    @value.setter
    def value(self, _: Any) -> None:
        raise AttributeError("Cell is part of a merged range")

    @property
    def data_type(self) -> str:
        return "n"

    @property
    def comment(self) -> None:
        return None

    @property
    def hyperlink(self) -> None:
        return None

    @property
    def number_format(self) -> str:
        return "General"

    @number_format.setter
    def number_format(self, _: str) -> None:
        return

    @property
    def style_id(self) -> int:
        return 0

    @property
    def style(self) -> str:
        return "Normal"

    @style.setter
    def style(self, _: str) -> None:
        return

    @property
    def font(self) -> Font:
        return Font(name="Calibri", size=11, family=2.0, scheme="minor")

    @property
    def fill(self) -> PatternFill:
        return PatternFill()

    @property
    def border(self) -> Border:
        if self._parent is None:
            return Border()
        from wolfxl._worksheet_collections import _merged_border_for_cell

        return _merged_border_for_cell(self._parent, self._row, self._col) or Border()

    @property
    def alignment(self) -> Alignment:
        return Alignment()

    @property
    def protection(self) -> Protection:
        return Protection()

    def __repr__(self) -> str:
        title = getattr(self._parent, "title", None)
        if title:
            return f"<MergedCell {title!r}.{self.coordinate}>"
        return f"<MergedCell {self.coordinate}>"


__all__ = ["MergedCell"]
