"""Utility descriptors shared across the chart module.

Mirrors :mod:`openpyxl.chart.descriptors` — these are the bounded numeric
helpers that openpyxl uses for gap amounts, overlap, and free-form numeric
formats. We keep the names identical so that openpyxl-shaped client code
that imports them by name keeps working.
"""

from __future__ import annotations

import math
import re
from typing import Any

from wolfxl._compat import _openpyxl_name_fallback


class _BoundedNumber:
    """Validate a numeric attribute against an inclusive [min, max] range.

    Used by chart-side descriptors that need range-checking (gap width,
    overlap, etc.). Attribute access goes through ``__set_name__`` so
    each descriptor instance manages its own per-class slot name.
    """

    _min: float | None = None
    _max: float | None = None
    _allow_none: bool = True

    def __init__(
        self,
        min: float | None = None,
        max: float | None = None,
        allow_none: bool = True,
    ) -> None:
        if min is not None:
            self._min = min
        if max is not None:
            self._max = max
        self._allow_none = allow_none

    def __set_name__(self, owner: type, name: str) -> None:
        self._attr = "_" + name

    def __get__(self, instance: Any, owner: type | None = None) -> Any:
        if instance is None:
            return self
        return getattr(instance, self._attr, None)

    def __set__(self, instance: Any, value: Any) -> None:
        if value is None:
            if not self._allow_none:
                raise ValueError(f"{self._attr[1:]} cannot be None")
            setattr(instance, self._attr, None)
            return
        v = float(value)
        if self._min is not None and v < self._min:
            raise ValueError(f"{self._attr[1:]}={v} below min={self._min}")
        if self._max is not None and v > self._max:
            raise ValueError(f"{self._attr[1:]}={v} above max={self._max}")
        setattr(instance, self._attr, v)


class NestedGapAmount(_BoundedNumber):
    """Gap amount for bar / pie variants — 0..500 (percentage units)."""

    def __init__(self, allow_none: bool = True) -> None:
        super().__init__(min=0, max=500, allow_none=allow_none)


class NestedOverlap(_BoundedNumber):
    """Bar-chart overlap percentage — -100..100."""

    def __init__(self, allow_none: bool = True) -> None:
        super().__init__(min=-100, max=100, allow_none=allow_none)


# DrawingML chart scalar contracts. Keep these native: openpyxl is optional
# and must not determine whether a shipped save accepts invalid state.
_ENUM_MEMBERS = {
    "barDir": ("bar", "col"),
    "grouping": ("standard", "stacked", "percentStacked"),
    "shape": ("box", "cone", "coneToMax", "cylinder", "pyramid", "pyramidToMax"),
    "ofPieType": ("bar", "pie"),
    "splitType": ("auto", "pos", "percent", "val", "cust"),
    "scatterStyle": ("line", "lineMarker", "marker", "smooth", "smoothMarker"),
    "radarStyle": ("standard", "marker", "filled"),
    "sizeRepresents": ("area", "w"),
}
_BOOL_MEMBERS = frozenset(
    ("varyColors", "marker", "bubble3D", "showNegBubbles", "wireframe", "rAngAx")
)
_NUMBER_MEMBERS = {
    "gapWidth": (0, 500),
    "overlap": (-100, 100),
    "gapDepth": (0, 500),
    "firstSliceAng": (0, 360),
    "holeSize": (1, 90),
    "splitPos": (None, None),
    "secondPieSize": (5, 200),
    "bubbleScale": (0, 300),
    "rotX": (-90, 90),
    "rotY": (0, 360),
    "hPercent": (5, 500),
    "depthPercent": (20, 2000),
    "perspective": (0, 240),
    "thickness": (0, None),
    "order": (0, None),
}
_INTEGER_MEMBERS = frozenset(("rotY", "depthPercent", "perspective", "thickness", "order"))


def _validate_members(model: Any, names: tuple[str | tuple[str, str], ...]) -> None:
    """Check live members without coercion or mutation before serialization."""
    for member in names:
        name, attribute = (member, member) if isinstance(member, str) else member
        value = getattr(model, attribute, None)
        if value is None:
            continue
        choices = _ENUM_MEMBERS.get(name)
        if (
            name == "grouping"
            and value == "clustered"
            and getattr(model, "tagname", None) in ("barChart", "bar3DChart")
        ):
            continue
        if choices is not None:
            if not isinstance(value, str) or value not in choices:
                raise ValueError(f"{type(model).__name__}.{name}={value!r} is not allowed")
        elif name in _BOOL_MEMBERS:
            if not isinstance(value, bool):
                raise TypeError(f"{type(model).__name__}.{name} must be a bool or None")
        elif name in _NUMBER_MEMBERS:
            if isinstance(value, bool) or not isinstance(value, (int, float)):
                raise TypeError(f"{type(model).__name__}.{name} must be numeric or None")
            if isinstance(value, float) and not math.isfinite(value):
                raise ValueError(f"{type(model).__name__}.{name} must be finite")
            if name in _INTEGER_MEMBERS and value != int(value):
                raise ValueError(f"{type(model).__name__}.{name} must be integral")
            low, high = _NUMBER_MEMBERS[name]
            if (low is not None and value < low) or (high is not None and value > high):
                raise ValueError(f"{type(model).__name__}.{name}={value!r} outside [{low}, {high}]")


def _validate_graphical_colors(properties: Any) -> None:
    if properties is None:
        return
    from .text import SCHEME_COLOR_NAMES

    for owner in (properties, getattr(properties, "ln", None)):
        color = getattr(owner, "solidFill", None)
        if color is None:
            continue
        # ColorChoice retains its typed carriers; raw strings follow the
        # existing RGB and scheme color compatibility surface.
        if not isinstance(color, str):
            color = getattr(color, "srgbClr", None)
            if color is None:
                continue
        if (
            re.fullmatch(r"(?:[0-9a-fA-F]{6}|[0-9a-fA-F]{8})", color) is None
            and color not in SCHEME_COLOR_NAMES
        ):
            raise ValueError(f"solidFill={color!r} must be RGB, ARGB or a scheme color")


def _validate_chart_members(chart: Any) -> None:
    """Validate every chart family before reindexing or publishing save bytes."""
    for name in ("roundedCorners", "visible_cells_only", "vary_colors"):
        value = getattr(chart, name, None)
        if value is not None and not isinstance(value, bool):
            raise TypeError(f"{name} must be a bool or None")
    _validate_members(
        chart,
        (
            "barDir",
            "grouping",
            "varyColors",
            "gapWidth",
            "overlap",
            "gapDepth",
            "shape",
            "marker",
            "firstSliceAng",
            "holeSize",
            "ofPieType",
            "splitType",
            "splitPos",
            "secondPieSize",
            "scatterStyle",
            "radarStyle",
            "bubble3D",
            "bubbleScale",
            "showNegBubbles",
            "sizeRepresents",
            "wireframe",
        ),
    )
    # Projected pie exposes both spellings as separate mutable attributes.
    # Validate the snake-case emitter values as well as the XML-facing ones.
    if getattr(chart, "tagname", None) == "ofPieChart":
        _validate_members(
            chart,
            (
                ("ofPieType", "of_pie_type"),
                ("splitType", "split_type"),
                ("splitPos", "split_pos"),
                ("secondPieSize", "second_pie_size"),
            ),
        )
    view = getattr(chart, "view3D", None)
    if view is not None:
        _validate_members(
            view, ("rotX", "rotY", "hPercent", "depthPercent", "rAngAx", "perspective")
        )
    for name in ("floor", "sideWall", "backWall"):
        surface = getattr(chart, name, None)
        if surface is not None:
            _validate_members(surface, ("thickness",))
            _validate_graphical_colors(getattr(surface, "spPr", surface))
    _validate_graphical_colors(getattr(chart, "graphical_properties", None))
    plot_area = getattr(chart, "plot_area", None)
    _validate_graphical_colors(getattr(plot_area, "graphicalProperties", None))
    for series in chart.ser:
        _validate_members(series, ("order",))


__all__ = ["NestedGapAmount", "NestedOverlap"]

__getattr__ = _openpyxl_name_fallback(globals())
