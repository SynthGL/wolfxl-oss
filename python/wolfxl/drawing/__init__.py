"""Shim for ``openpyxl.drawing``."""

from __future__ import annotations

from importlib import import_module as _import_module
from typing import Any

from wolfxl._compat import _openpyxl_name_fallback

_EXPORTS = {
    "Drawing": ("spreadsheet_drawing", "SpreadsheetDrawing"),
    "Image": ("image", "Image"),
    "AbsoluteAnchor": ("spreadsheet_drawing", "AbsoluteAnchor"),
    "AnchorMarker": ("spreadsheet_drawing", "AnchorMarker"),
    "OneCellAnchor": ("spreadsheet_drawing", "OneCellAnchor"),
    "TwoCellAnchor": ("spreadsheet_drawing", "TwoCellAnchor"),
    "XDRPoint2D": ("spreadsheet_drawing", "XDRPoint2D"),
    "XDRPositiveSize2D": ("spreadsheet_drawing", "XDRPositiveSize2D"),
}

__all__ = [
    "AbsoluteAnchor",
    "AnchorMarker",
    "Drawing",
    "Image",
    "OneCellAnchor",
    "TwoCellAnchor",
    "XDRPoint2D",
    "XDRPositiveSize2D",
]

_name_fallback = _openpyxl_name_fallback(globals())


def __getattr__(name: str) -> Any:
    target = _EXPORTS.get(name)
    if target is None:
        return _name_fallback(name)
    value = getattr(_import_module(f"{__name__}.{target[0]}"), target[1])
    globals()[name] = value
    return value


def __dir__() -> list[str]:
    return sorted(set(globals()) | set(__all__))
