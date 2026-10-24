"""Coordinate conversion helpers for A1-style Excel references."""

from __future__ import annotations

import re
from functools import lru_cache

_A1_RE = re.compile(r"^([A-Z]+)(\d+)$")


def column_letter(col_idx: int) -> str:
    """Convert a 1-based column index to a letter (1 -> 'A', 27 -> 'AA')."""
    if col_idx < 1:
        raise ValueError(f"Column index must be >= 1, got {col_idx}")
    if type(col_idx) is int and col_idx <= 16384:
        return _cached_column_letter(col_idx)
    return _uncached_column_letter(col_idx)


@lru_cache(maxsize=16384)
def _cached_column_letter(col_idx: int) -> str:
    return _uncached_column_letter(col_idx)


def _uncached_column_letter(col_idx: int) -> str:
    result = []
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result.append(chr(65 + remainder))
    return "".join(reversed(result))


def column_index(col_letter: str) -> int:
    """Convert a column letter to a 1-based index ('A' -> 1, 'AA' -> 27)."""
    result = 0
    for ch in col_letter.upper():
        result = result * 26 + (ord(ch) - 64)
    return result


def a1_to_rowcol(a1: str) -> tuple[int, int]:
    """Convert an A1-style reference to (row, col) — both 1-based.

    Example: 'B3' -> (3, 2)
    """
    m = _A1_RE.match(a1.upper())
    if not m:
        raise ValueError(f"Invalid A1 reference: {a1!r}")
    return int(m.group(2)), column_index(m.group(1))


def rowcol_to_a1(row: int, col: int) -> str:
    """Convert (row, col) — both 1-based — to an A1-style reference.

    Example: (3, 2) -> 'B3'
    """
    return f"{column_letter(col)}{row}"
