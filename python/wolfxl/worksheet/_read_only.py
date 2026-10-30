"""Read-only worksheet compatibility path.

WolfXL's normal ``load_workbook(..., read_only=True)`` uses the Rust streaming
reader. This module exists for openpyxl-shaped imports and for tests/tools that
construct a read-only worksheet directly around an OOXML worksheet part.
"""

from __future__ import annotations

import re
from collections.abc import Iterator
from typing import Any

from wolfxl.cell.read_only import EMPTY_CELL, ReadOnlyCell
from wolfxl._worksheet import Worksheet
from wolfxl.workbook import DefinedNameDict
from wolfxl.utils import get_column_letter
from wolfxl.utils.cell import coordinate_to_tuple, range_boundaries
from wolfxl.utils.datetime import from_excel
from wolfxl.worksheet._reader import WorkSheetParser
from wolfxl.xml.constants import SHEET_MAIN_NS
from wolfxl.xml.functions import fromstring

_NS = {"main": SHEET_MAIN_NS}
_COORD_RE = re.compile(r"([A-Za-z]{1,3})([0-9]+)")


def read_dimension(source: Any) -> tuple[int, int, int | None, int | None] | None:
    """Read the worksheet dimension as ``(min_col, min_row, max_col, max_row)``."""
    root = fromstring(source.read())
    dimension = root.find("main:dimension", _NS)
    if dimension is None:
        return None
    ref = dimension.get("ref")
    if not ref:
        return None
    min_col, min_row, max_col, max_row = range_boundaries(ref)
    return min_col or 1, min_row or 1, max_col, max_row


class ReadOnlyWorksheet:
    """Openpyxl-compatible direct reader for a worksheet XML member."""

    _min_column = 1
    _min_row = 1
    _max_column: int | None = None
    _max_row: int | None = None

    def __init__(
        self,
        parent_workbook: Any,
        title: str,
        worksheet_path: str,
        shared_strings: list[Any],
    ) -> None:
        self.parent = parent_workbook
        self.title = title
        self.sheet_state = "visible"
        self._current_row = None
        self._worksheet_path = worksheet_path
        self._shared_strings = shared_strings
        self.defined_names = {}
        self._get_size()

    def _get_source(self) -> Any:
        return self.parent._archive.open(self._worksheet_path)

    def _get_size(self) -> None:
        src = self._get_source()
        try:
            dimensions = read_dimension(src)
        finally:
            src.close()
        if dimensions is not None:
            self._min_column, self._min_row, self._max_column, self._max_row = dimensions

    @property
    def min_row(self) -> int:
        return self._min_row

    @property
    def max_row(self) -> int | None:
        return self._max_row

    @property
    def min_column(self) -> int:
        return self._min_column

    @property
    def max_column(self) -> int | None:
        return self._max_column

    @property
    def rows(self) -> Iterator[tuple[Any, ...]]:
        return self.iter_rows()

    @property
    def values(self) -> Iterator[tuple[Any, ...]]:
        return self.iter_rows(values_only=True)

    def __iter__(self) -> Iterator[tuple[Any, ...]]:
        return self.rows

    def __getitem__(self, key: str) -> Any:
        if ":" in key:
            min_col, min_row, max_col, max_row = range_boundaries(key)
            return tuple(
                self.iter_rows(
                    min_row=min_row,
                    max_row=max_row,
                    min_col=min_col,
                    max_col=max_col,
                )
            )
        row, column = coordinate_to_tuple(key)
        return self._get_cell(row, column)

    def cell(self, row: int, column: int, value: Any = None) -> Any:
        if value is not None:
            raise AttributeError("Cell is read only")
        return self._get_cell(row, column)

    def iter_rows(
        self,
        min_row: int | None = None,
        max_row: int | None = None,
        min_col: int | None = None,
        max_col: int | None = None,
        values_only: bool = False,
    ) -> Iterator[tuple[Any, ...]]:
        yield from self._cells_by_row(
            min_col or 1,
            min_row or 1,
            max_col,
            max_row,
            values_only=values_only,
        )

    def _cells_by_row(
        self,
        min_col: int,
        min_row: int,
        max_col: int | None,
        max_row: int | None,
        values_only: bool = False,
    ) -> Iterator[tuple[Any, ...]]:
        filler = None if values_only else EMPTY_CELL
        max_col = max_col or self.max_column
        if max_row is None:
            max_row = self.max_row
        elif self.max_row is not None:
            max_row = min(max_row, self.max_row)
        empty_row: tuple[Any, ...] = ()
        if max_col is not None:
            empty_row = (filler,) * (max_col + 1 - min_col)

        counter = min_row
        stopped_at_overbound = False
        for idx, row in self._parse_rows():
            if idx < min_row:
                continue
            if max_row is not None and idx > max_row:
                stopped_at_overbound = True
                break
            for _ in range(counter, idx):
                counter += 1
                yield empty_row
            if counter <= idx:
                counter += 1
                yield self._get_row(row, min_col, max_col, values_only)

        if max_row is not None and (stopped_at_overbound or self.max_row is not None):
            for _ in range(counter, max_row + 1):
                yield empty_row

    def _get_row(
        self,
        row: list[dict[str, Any]],
        min_col: int = 1,
        max_col: int | None = None,
        values_only: bool = False,
    ) -> tuple[Any, ...]:
        if not row and max_col is None:
            return ()

        max_col = max_col or row[-1]["column"]
        row_width = max_col + 1 - min_col
        new_row: list[Any] = [None if values_only else EMPTY_CELL] * row_width

        for cell in row:
            column = cell["column"]
            if min_col <= column <= max_col:
                idx = column - min_col
                new_row[idx] = cell["value"] if values_only else ReadOnlyCell(self, **cell)
        return tuple(new_row)

    def _get_cell(self, row: int, column: int) -> Any:
        for row_values in self._cells_by_row(column, row, column, row):
            if row_values:
                return row_values[0]
        return EMPTY_CELL

    def calculate_dimension(self, force: bool = False) -> str:
        if not all([self.max_column, self.max_row]):
            if not force:
                raise ValueError("Worksheet is unsized, use calculate_dimension(force=True)")
            self._calculate_dimension()
        return (
            f"{get_column_letter(self.min_column)}{self.min_row}:"
            f"{get_column_letter(self.max_column or 1)}{self.max_row or 1}"
        )

    def _calculate_dimension(self) -> None:
        max_col = 0
        max_row = 0
        for row in self.rows:
            if not row:
                continue
            cell = row[-1]
            max_col = max(max_col, getattr(cell, "column", 0))
            max_row = max(max_row, getattr(cell, "row", 0))
        self._max_row = max_row or None
        self._max_column = max_col or None

    def reset_dimensions(self) -> None:
        self._max_row = self._max_column = None

    def _parse_rows(self) -> Iterator[tuple[int, list[dict[str, Any]]]]:
        src = self._get_source()
        try:
            root = fromstring(src.read())
        finally:
            src.close()
        for row_el in root.findall(".//main:sheetData/main:row", _NS):
            row_idx = int(row_el.get("r", "0") or 0)
            row_cells: list[dict[str, Any]] = []
            for cell_el in row_el.findall("main:c", _NS):
                ref = cell_el.get("r")
                if ref:
                    row, column = coordinate_to_tuple(ref)
                else:
                    row = row_idx
                    column = len(row_cells) + 1
                style_id = int(cell_el.get("s", "0") or 0)
                value, data_type = self._cell_value(cell_el, style_id)
                row_cells.append(
                    {
                        "row": row,
                        "column": column,
                        "value": value,
                        "data_type": data_type,
                        "style_id": style_id,
                    }
                )
            yield row_idx, row_cells

    def _cell_value(self, cell_el: Any, style_id: int = 0) -> tuple[Any, str]:
        cell_type = cell_el.get("t", "n")
        if cell_type == "inlineStr" and cell_el.find("main:is", _NS) is None:
            # openpyxl keeps ``inlineStr`` for a cell without an ``<is>`` child.
            return None, "inlineStr"
        value_el = cell_el.find("main:v", _NS)
        if value_el is None:
            inline = cell_el.find("main:is/main:t", _NS)
            return (inline.text if inline is not None else None), "s"
        text = value_el.text
        if text is None:
            return None, cell_type
        if cell_type == "s":
            try:
                return self._shared_strings[int(text)], "s"
            except Exception:
                return text, "s"
        if cell_type == "b":
            return text == "1", "b"
        if cell_type in {"str", "inlineStr"}:
            return text, "s"
        if cell_type == "n":
            try:
                number = float(text)
            except ValueError:
                return text, "n"
            if style_id in getattr(self.parent, "_date_formats", set()):
                value = from_excel(
                    number,
                    getattr(self.parent, "epoch", None),
                    timedelta=style_id in getattr(self.parent, "_timedelta_formats", set()),
                )
                return value, "d"
            if number.is_integer() and "E" not in text.upper() and "." not in text:
                return int(number), "n"
            return number, "n"
        return text, cell_type


__all__ = [
    "DefinedNameDict",
    "ReadOnlyWorksheet",
    "WorkSheetParser",
    "Worksheet",
    "read_dimension",
]
