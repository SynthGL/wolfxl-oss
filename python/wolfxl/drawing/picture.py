"""Drawing picture containers and atomic original-raster geometry operations."""

from __future__ import annotations
from copy import copy as _copy
from fractions import Fraction as _Fraction
import math as _math
import struct as _struct
from typing import Any as _Any


from wolfxl._compat import _OpenpyxlSerialisable, _make_serialisable, _resolve_openpyxl_class
from wolfxl.chart.shapes import GraphicalProperties
from wolfxl.drawing.fill import BlipFillProperties
from wolfxl.xml.constants import DRAWING_NS


def _openpyxl_class(name: str) -> type:
    return _resolve_openpyxl_class(__name__, name) or _make_serialisable(name)


for _name in (
    "NonVisualDrawingProps",
    "NonVisualPictureProperties",
    "OfficeArtExtensionList",
    "PictureFrame",
    "PictureLocking",
    "PictureNonVisual",
):
    globals()[_name] = _openpyxl_class(_name)

ShapeStyle = _openpyxl_class("ShapeStyle")
Alias = Bool = Serialisable = String = Typed = _OpenpyxlSerialisable


def _source_dimensions(image: _Any) -> tuple[int, int]:
    from wolfxl.drawing.image import _dimensions, _sniff_format

    try:
        data = image._data
        width, height = _dimensions(_sniff_format(data), data)
        if width <= 0 or height <= 0:
            raise ValueError("nonpositive source dimensions")
    except (AttributeError, TypeError, ValueError, _struct.error) as exc:
        raise ValueError("picture_source_dimensions_unavailable") from exc
    return width, height


def _positive_integer(value: _Any) -> bool:
    return isinstance(value, int) and not isinstance(value, bool) and value > 0


def _fill_for_crop(image: _Any, *, fitting: bool = False) -> _Any:
    fill = image.blipFill
    code = "picture_fit_fill_unsupported" if fitting else "picture_crop_fill_unsupported"
    if fill is not None:
        if getattr(fill, "tile", None) is not None:
            raise ValueError(code)
        if fitting:
            stretch = getattr(fill, "stretch", None)
            rect = getattr(stretch, "fillRect", None)
            if rect is not None and any(
                getattr(rect, side, None) not in (None, 0) for side in ("l", "t", "r", "b")
            ):
                raise ValueError(code)
    from wolfxl.drawing.fill import Blip

    result = _copy(fill) if fill is not None else BlipFillProperties()
    blip = getattr(result, "blip", None)
    if blip is None:
        result.blip = Blip(embed="picture-source")
    elif not getattr(blip, "embed", None):
        result.blip = _copy(blip)
        result.blip.embed = "picture-source"
    return result


def _nearest(value: _Fraction) -> int:
    """Round a nonnegative rational to nearest integer, ties upward."""
    return (2 * value.numerator + value.denominator) // (2 * value.denominator)


def _set_source_rect(fill: _Any, sides: tuple[int, int, int, int]) -> None:
    from wolfxl.drawing.fill import RelativeRect, StretchInfoProperties

    fill.srcRect = RelativeRect(**dict(zip(("l", "t", "r", "b"), sides))) if any(sides) else None
    if getattr(fill, "stretch", None) is None:
        fill.stretch = StretchInfoProperties()


def _crop_source_pixels(image: _Any, x: int, y: int, width: int, height: int) -> None:
    if (
        not _positive_integer(width)
        or not _positive_integer(height)
        or any(
            not isinstance(value, int) or isinstance(value, bool) or value < 0 for value in (x, y)
        )
    ):
        raise ValueError("picture_crop_rectangle_invalid")
    source_width, source_height = _source_dimensions(image)
    if x + width > source_width or y + height > source_height:
        raise ValueError("picture_crop_rectangle_invalid")
    fill = _fill_for_crop(image)
    sides = (
        _nearest(_Fraction(x * 100000, source_width)),
        _nearest(_Fraction(y * 100000, source_height)),
        _nearest(_Fraction((source_width - x - width) * 100000, source_width)),
        _nearest(_Fraction((source_height - y - height) * 100000, source_height)),
    )
    # A valid pixel region can be too narrow to represent in DrawingML.
    if sides[0] + sides[2] >= 100000 or sides[1] + sides[3] >= 100000:
        raise ValueError("picture_crop_rectangle_invalid")
    _set_source_rect(fill, sides)
    image.blipFill = fill


def _source_crop_pixels(image: _Any) -> tuple[float, float, float, float]:
    width, height = _source_dimensions(image)
    if getattr(image.blipFill, "tile", None) is not None:
        raise ValueError("picture_crop_fill_unsupported")
    rect = getattr(image.blipFill, "srcRect", None)
    sides = tuple(
        0 if getattr(rect, side, None) is None else getattr(rect, side)
        for side in ("l", "t", "r", "b")
    )
    if (
        any(
            isinstance(value, bool)
            or not isinstance(value, (int, float))
            or not 0 <= value <= 100000
            or int(value) != value
            for value in sides
        )
        or sides[0] + sides[2] >= 100000
        or sides[1] + sides[3] >= 100000
    ):
        raise ValueError("picture_crop_source_invalid")
    left, top, right, bottom = sides
    return (
        width * left / 100000,
        height * top / 100000,
        width * (100000 - left - right) / 100000,
        height * (100000 - top - bottom) / 100000,
    )


def _fit_box(
    image: _Any,
    width: int,
    height: int,
    *,
    mode: str,
    alignment: tuple[float, float],
) -> None:
    from wolfxl.drawing.spreadsheet_drawing import (
        AbsoluteAnchor,
        OneCellAnchor,
        XDRPositiveSize2D,
    )

    if mode not in ("contain", "cover", "stretch"):
        raise ValueError("picture_fit_mode_unsupported")
    if not _positive_integer(width) or not _positive_integer(height):
        raise ValueError("picture_fit_box_invalid")
    if (
        not isinstance(alignment, (tuple, list))
        or len(alignment) != 2
        or any(
            isinstance(value, bool)
            or not isinstance(value, (int, float))
            or not 0 <= value <= 1
            or not _math.isfinite(value)
            for value in alignment
        )
    ):
        raise ValueError("picture_fit_alignment_invalid")
    anchor = image.anchor
    if anchor is not None and not isinstance(anchor, (str, OneCellAnchor, AbsoluteAnchor)):
        raise ValueError("picture_fit_anchor_unsupported")
    source_width, source_height = _source_dimensions(image)
    fill = _fill_for_crop(image, fitting=True)
    sides = (0, 0, 0, 0)
    display_width, display_height = width, height
    scale_x = _Fraction(width, source_width)
    scale_y = _Fraction(height, source_height)
    if mode == "contain":
        scale = min(scale_x, scale_y)
        display_width = max(1, _nearest(source_width * scale))
        display_height = max(1, _nearest(source_height * scale))
    elif mode == "cover":
        scale = max(scale_x, scale_y)
        horizontal = 100000 * (1 - scale_x / scale)
        vertical = 100000 * (1 - scale_y / scale)
        left = _nearest(horizontal * _Fraction(alignment[0]))
        top = _nearest(vertical * _Fraction(alignment[1]))
        right = _nearest(horizontal * (1 - _Fraction(alignment[0])))
        bottom = _nearest(vertical * (1 - _Fraction(alignment[1])))
        sides = (left, top, right, bottom)
        if left + right >= 100000 or top + bottom >= 100000:
            raise ValueError("picture_fit_box_invalid")
    _set_source_rect(fill, sides)
    extent = XDRPositiveSize2D(cx=display_width * 9525, cy=display_height * 9525)
    image.blipFill = fill
    image.width, image.height = display_width, display_height
    if isinstance(anchor, (OneCellAnchor, AbsoluteAnchor)):
        anchor.ext = extent


__all__ = [name for name in list(globals()) if not name.startswith("_")]
