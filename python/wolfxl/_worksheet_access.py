"""Worksheet cell and range access helpers."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from wolfxl._cell import Cell
from wolfxl._utils import a1_to_rowcol
from wolfxl.utils.cell import column_index_from_string, range_boundaries

if TYPE_CHECKING:
    from wolfxl._worksheet import Worksheet


def get_item(ws: Worksheet, key: Any) -> Any:
    """Resolve an openpyxl-compatible worksheet index."""
    if isinstance(key, int):
        return get_row_tuple(ws, key, key)[0]

    if isinstance(key, slice):
        if key.start is None or key.stop is None:
            raise ValueError("Row slice bounds must be specified")
        if isinstance(key.start, str) or isinstance(key.stop, str):
            if not isinstance(key.start, str) or not isinstance(key.stop, str):
                raise TypeError("A1 slice bounds must both be strings")
            min_row, min_col = a1_to_rowcol(key.start)
            max_row, max_col = a1_to_rowcol(key.stop)
            return get_rect(ws, min_row, min_col, max_row, max_col)
        return get_row_tuple(ws, key.start, key.stop)

    if isinstance(key, str):
        return resolve_string_key(ws, key)

    raise TypeError(
        f"Worksheet indices must be str, int, or slice, not {type(key).__name__}"
    )


def resolve_string_key(ws: Worksheet, key: str) -> Any:
    """Resolve a string key to a Cell, row tuple, column tuple, or 2D range."""
    try:
        row, col = a1_to_rowcol(key)
    except ValueError:
        pass
    else:
        return ws._get_or_create_cell(row, col)  # noqa: SLF001

    if key.isdigit():
        row = int(key)
        return get_row_tuple(ws, row, row)[0]

    try:
        col_idx = column_index_from_string(key)
    except ValueError:
        col_idx = None
    if col_idx is not None and not any(ch.isdigit() for ch in key):
        return tuple(row for row in get_col_tuple(ws, col_idx, col_idx))[0]

    min_col, min_row, max_col, max_row = range_boundaries(key)

    if min_row is None and max_row is None:
        bounded_max_row = ws._bounds_max_row()  # noqa: SLF001
        if min_col is None or max_col is None:
            raise ValueError(f"Invalid range: {key!r}")
        return get_col_tuple(ws, min_col, max_col, 1, bounded_max_row)

    if min_col is None and max_col is None:
        bounded_max_col = ws._bounds_max_col()  # noqa: SLF001
        if min_row is None or max_row is None:
            raise ValueError(f"Invalid range: {key!r}")
        if min_row == max_row:
            return get_row_tuple(ws, min_row, max_row)[0]
        return get_rect(ws, min_row, 1, max_row, bounded_max_col)

    if min_row is None or max_row is None or min_col is None or max_col is None:
        raise ValueError(f"Invalid range: {key!r}")

    return get_rect(ws, min_row, min_col, max_row, max_col)


def get_rect(
    ws: Worksheet,
    min_row: int,
    min_col: int,
    max_row: int,
    max_col: int,
) -> tuple[tuple[Cell, ...], ...]:
    """Return a 2D tuple of cells for an inclusive rectangle."""
    return tuple(
        tuple(ws._get_or_create_cell(row, col) for col in range(min_col, max_col + 1))  # noqa: SLF001
        for row in range(min_row, max_row + 1)
    )


def get_row_tuple(
    ws: Worksheet,
    min_row: int,
    max_row: int,
) -> tuple[tuple[Cell, ...], ...]:
    """Return row-major cell tuples for an inclusive row span."""
    max_col = ws._bounds_max_col()  # noqa: SLF001
    return tuple(
        tuple(ws._get_or_create_cell(row, col) for col in range(1, max_col + 1))  # noqa: SLF001
        for row in range(min_row, max_row + 1)
    )


def get_col_tuple(
    ws: Worksheet,
    min_col: int,
    max_col: int,
    min_row: int | None = None,
    max_row: int | None = None,
) -> tuple[tuple[Cell, ...], ...]:
    """Return column-major cell tuples for an inclusive column span."""
    row_min = min_row if min_row is not None else 1
    row_max = max_row if max_row is not None else ws._bounds_max_row()  # noqa: SLF001
    return tuple(
        tuple(ws._get_or_create_cell(row, col) for row in range(row_min, row_max + 1))  # noqa: SLF001
        for col in range(min_col, max_col + 1)
    )
