"""Table style model compatibility."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from wolfxl import _compat
from wolfxl._compat import _OpenpyxlSerialisable, _resolve_openpyxl_class


def _upstream_class(name: str) -> type | None:
    """Borrow upstream classes only when openpyxl has already been imported."""
    if _compat._real_openpyxl_if_imported() is None:
        return None
    return _resolve_openpyxl_class(__name__, name)


TableStyleElement = _upstream_class("TableStyleElement") or type(
    "TableStyleElement",
    (_OpenpyxlSerialisable,),
    {"__attrs__": ("type", "size", "dxfId"), "tagname": "tableStyleElement"},
)
TableStyle = _upstream_class("TableStyle")
if TableStyle is None:

    @dataclass
    class TableStyle(_OpenpyxlSerialisable):
        tagname = "tableStyle"
        __attrs__ = ("name", "pivot", "table", "count")
        __elements__ = ("tableStyleElement",)

        name: str | None = None
        pivot: bool | None = None
        table: bool | None = None
        count: int | None = None
        tableStyleElement: list[TableStyleElement] = field(default_factory=list)  # noqa: N815


_OpenpyxlTableStyleList = _upstream_class("TableStyleList")


if _OpenpyxlTableStyleList is None:

    @dataclass
    class TableStyleList:
        count: int | None = None
        defaultTableStyle: str | None = "TableStyleMedium9"  # noqa: N815
        defaultPivotStyle: str | None = "PivotStyleLight16"  # noqa: N815
        tableStyle: list[TableStyle] = field(default_factory=list)  # noqa: N815

        def __post_init__(self) -> None:
            if self.count is None:
                self.count = len(self.tableStyle)

        def __iter__(self):
            return iter(self.tableStyle)

        def append(self, value: TableStyle) -> None:
            self.tableStyle.append(value)
            self.count = len(self.tableStyle)

else:

    class TableStyleList(_OpenpyxlTableStyleList):  # type: ignore[misc, valid-type]
        __attrs__ = _OpenpyxlTableStyleList.__attrs__
        __elements__ = _OpenpyxlTableStyleList.__elements__

        def append(self, value: TableStyle) -> None:
            self.tableStyle.append(value)
            self.count = len(self.tableStyle)


Bool = Color = Float = Integer = NoneSet = Sequence = Serialisable = Set = String = Typed = _OpenpyxlSerialisable

__all__ = [
    "Bool",
    "Color",
    "Float",
    "Integer",
    "NoneSet",
    "Sequence",
    "Serialisable",
    "Set",
    "String",
    "TableStyle",
    "TableStyleElement",
    "TableStyleList",
    "Typed",
]
