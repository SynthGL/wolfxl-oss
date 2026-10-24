"""openpyxl-compatible utility surface.

Mirrors the symbols SynthGL imports from ``openpyxl.utils.*`` and
``openpyxl.styles.numbers.is_date_format``. The behavior is deliberately
*bug-for-bug* identical to openpyxl 3.1+ - pinned by
``tests/parity/test_utils_parity.py``. Any divergence is a parity bug.

Helpers below (``absolute_coordinate``, ``quote_sheetname``, ``range_to_tuple``,
``rows_from_range``, ``cols_from_range``, ``get_column_interval``) mirror
openpyxl's ``openpyxl/utils/cell.py`` under its MIT license.
"""

from __future__ import annotations

from collections.abc import Iterator

from wolfxl._compat import _openpyxl_name_fallback
from wolfxl.utils.cell import (
    ABSOLUTE_RE,
    column_index_from_string,
    coordinate_to_tuple,
    get_column_letter,
    range_boundaries,
)
from wolfxl.utils.datetime import (
    CALENDAR_WINDOWS_1900,
    from_excel,
)
from wolfxl.utils.exceptions import CellCoordinatesException
from wolfxl.utils.numbers import is_date_format

__all__ = [
    "CALENDAR_WINDOWS_1900",
    "FORMULAE",
    "absolute_coordinate",
    "cols_from_range",
    "column_index_from_string",
    "coordinate_to_tuple",
    "from_excel",
    "get_column_interval",
    "get_column_letter",
    "is_date_format",
    "quote_sheetname",
    "range_boundaries",
    "range_to_tuple",
    "rows_from_range",
]

_name_fallback = _openpyxl_name_fallback(globals())


def __getattr__(name: str):
    if name == "FORMULAE":
        from wolfxl.utils import formulas

        value = formulas.FORMULAE
        globals()[name] = value
        return value
    return _name_fallback(name)


def absolute_coordinate(coord_string: str) -> str:
    """Convert ``"A1"`` / ``"A1:B2"`` to absolute form with ``$`` prefixes.

    ``"A1"`` -> ``"$A$1"``. ``"A1:B2"`` -> ``"$A$1:$B$2"``. Mirrors
    openpyxl's ``absolute_coordinate``.
    """
    m = ABSOLUTE_RE.match(coord_string)
    if not m:
        raise ValueError(f"{coord_string} is not a valid coordinate range")

    parts = m.groupdict("")
    for key, value in parts.items():
        if value:
            parts[key] = f"${value}"

    if parts["max_col"] or parts["max_row"]:
        fmt = "{min_col}{min_row}:{max_col}{max_row}"
    else:
        fmt = "{min_col}{min_row}"
    return fmt.format(**parts)


def quote_sheetname(sheetname: str) -> str:
    """Wrap a sheet name in single quotes.

    openpyxl's contract (verified empirically): the sheet name is *always*
    quoted, with any embedded single quotes doubled per Excel's escape
    convention. We keep parity even though it looks over-eager for plain
    names like ``"Sheet1"``, since consumers concatenating ``f"{quote(s)}!{ref}"``
    depend on the shape.
    """
    return "'{}'".format(sheetname.replace("'", "''"))


def range_to_tuple(range_string: str) -> tuple[str, tuple[int, int, int, int]]:
    """Parse ``"'Sheet 1'!A1:B2"`` into ``(sheet, (min_col, min_row, max_col, max_row))``.

    A sheet qualifier (``'Sheet'!``) is required - openpyxl's contract.
    Single quotes around the sheet name are stripped if present.
    """
    if "!" not in range_string:
        raise CellCoordinatesException(
            f"Range must have a sheet qualifier: {range_string!r}"
        )
    sheet_part, _, range_part = range_string.rpartition("!")
    sheet = sheet_part
    if sheet.startswith("'") and sheet.endswith("'"):
        sheet = sheet[1:-1].replace("''", "'")
    bounds = range_boundaries(range_part)
    if any(b is None for b in bounds):
        raise CellCoordinatesException(
            f"Range must have explicit bounds: {range_string!r}"
        )
    # mypy: bounds are all ints after the None check.
    return sheet, (int(bounds[0]), int(bounds[1]), int(bounds[2]), int(bounds[3]))  # type: ignore[arg-type]


def rows_from_range(range_string: str) -> Iterator[tuple[str, ...]]:
    """Yield each row of a range as a tuple of A1 coord strings.

    ``"A1:C2"`` yields ``("A1", "B1", "C1")`` then ``("A2", "B2", "C2")``.
    Matches openpyxl's ``rows_from_range``.
    """
    min_col, min_row, max_col, max_row = range_boundaries(range_string)
    if min_col is None or min_row is None or max_col is None or max_row is None:
        raise CellCoordinatesException(
            f"Range must have explicit bounds: {range_string!r}"
        )
    for row in range(min_row, max_row + 1):
        yield tuple(
            f"{get_column_letter(col)}{row}" for col in range(min_col, max_col + 1)
        )


def cols_from_range(range_string: str) -> Iterator[tuple[str, ...]]:
    """Yield each column of a range as a tuple of A1 coord strings.

    ``"A1:C2"`` yields ``("A1", "A2")`` then ``("B1", "B2")`` then
    ``("C1", "C2")``. Matches openpyxl's ``cols_from_range``.
    """
    min_col, min_row, max_col, max_row = range_boundaries(range_string)
    if min_col is None or min_row is None or max_col is None or max_row is None:
        raise CellCoordinatesException(
            f"Range must have explicit bounds: {range_string!r}"
        )
    for col in range(min_col, max_col + 1):
        letter = get_column_letter(col)
        yield tuple(f"{letter}{row}" for row in range(min_row, max_row + 1))


def get_column_interval(start: int | str, end: int | str) -> list[str]:
    """Return the column letters from ``start`` to ``end`` inclusive.

    Either argument may be a 1-based int or an A1 column letter string.
    ``get_column_interval("A", "C")`` -> ``["A", "B", "C"]``. Matches
    openpyxl's behavior of returning an empty list when ``end < start``
    (no automatic swap).
    """
    start_idx = start if isinstance(start, int) else column_index_from_string(start)
    end_idx = end if isinstance(end, int) else column_index_from_string(end)
    return [get_column_letter(i) for i in range(start_idx, end_idx + 1)]
