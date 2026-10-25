"""Worksheet access helpers use bounds methods, not pandas dimension attributes."""

from __future__ import annotations

import pytest

import wolfxl
from wolfxl._worksheet_access import get_item


@pytest.mark.parametrize(
    ("key", "expected"),
    [
        ("A:B", (("a", "b", "c"), (1, 2, 3))),
        ("1:2", (("a", 1), ("b", 2))),
        (1, ("a", 1)),
        ("A", ("a", "b", "c")),
    ],
)
def test_access_helpers_use_bounds_methods(key, expected):
    workbook = wolfxl.Workbook()
    sheet = workbook.active
    for row in (("a", 1), ("b", 2), ("c", 3)):
        sheet.append(row)

    def values(result):
        if isinstance(result, tuple):
            return tuple(values(cell) for cell in result)
        return result.value

    assert values(get_item(sheet, key)) == expected
    workbook.close()
