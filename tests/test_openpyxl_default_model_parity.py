"""Focused default XML model parity against openpyxl 3.1.x."""

from __future__ import annotations

from collections.abc import Callable
from typing import Any
from xml.etree.ElementTree import tostring

import pytest


def _snapshot(factory: Callable[[], Any]) -> dict[str, Any]:
    obj = factory()
    attrs = tuple(getattr(obj, "__attrs__", ()))
    elements = tuple(getattr(obj, "__elements__", ()))
    return {
        "defaults": {name: getattr(obj, name) for name in (*attrs, *elements)},
        "attrs": attrs,
        "elements": elements,
        "iter": list(obj),
    }


@pytest.mark.parametrize(
    ("label", "openpyxl_factory", "wolfxl_factory"),
    [
        pytest.param(
            "cell.text.InlineFont",
            lambda: __import__("openpyxl.cell.text", fromlist=["InlineFont"]).InlineFont(),
            lambda: __import__("wolfxl.cell.text", fromlist=["InlineFont"]).InlineFont(),
            id="inline-font",
        ),
        pytest.param(
            "cell.text.Text",
            lambda: __import__("openpyxl.cell.text", fromlist=["Text"]).Text(),
            lambda: __import__("wolfxl.cell.text", fromlist=["Text"]).Text(),
            id="cell-text",
        ),
        pytest.param(
            "chart._3d.View3D",
            lambda: __import__("openpyxl.chart._3d", fromlist=["View3D"]).View3D(),
            lambda: __import__("wolfxl.chart._3d", fromlist=["View3D"]).View3D(),
            id="view-3d",
        ),
        pytest.param(
            "drawing.colors.ColorMapping",
            lambda: __import__("openpyxl.drawing.colors", fromlist=["ColorMapping"]).ColorMapping(),
            lambda: __import__("wolfxl.drawing.colors", fromlist=["ColorMapping"]).ColorMapping(),
            id="drawing-color-mapping",
        ),
        pytest.param(
            "chart.chartspace.ColorMapping",
            lambda: __import__(
                "openpyxl.chart.chartspace", fromlist=["ColorMapping"]
            ).ColorMapping(),
            lambda: __import__("wolfxl.chart.chartspace", fromlist=["ColorMapping"]).ColorMapping(),
            id="chartspace-color-mapping",
        ),
        pytest.param(
            "chart.print_settings.PageMargins",
            lambda: __import__(
                "openpyxl.chart.print_settings", fromlist=["PageMargins"]
            ).PageMargins(),
            lambda: __import__(
                "wolfxl.chart.print_settings", fromlist=["PageMargins"]
            ).PageMargins(),
            id="chart-page-margins",
        ),
        pytest.param(
            "chart.print_settings.PrintSettings",
            lambda: __import__(
                "openpyxl.chart.print_settings", fromlist=["PrintSettings"]
            ).PrintSettings(),
            lambda: __import__(
                "wolfxl.chart.print_settings", fromlist=["PrintSettings"]
            ).PrintSettings(),
            id="chart-print-settings",
        ),
        pytest.param(
            "chart.pivot.PivotFormat",
            lambda: __import__("openpyxl.chart.pivot", fromlist=["PivotFormat"]).PivotFormat(),
            lambda: __import__("wolfxl.chart.pivot", fromlist=["PivotFormat"]).PivotFormat(),
            id="pivot-format",
        ),
    ],
)
def test_default_xml_model_snapshots_match_openpyxl(
    label: str,
    openpyxl_factory: Callable[[], Any],
    wolfxl_factory: Callable[[], Any],
) -> None:
    assert _snapshot(wolfxl_factory) == _snapshot(openpyxl_factory), label


@pytest.mark.parametrize("flags", [{}, {"b": None, "i": None}, {"b": True}, {"i": True}])
def test_inline_font_serialization_matches_openpyxl(flags: dict[str, Any]) -> None:
    from openpyxl.cell.text import InlineFont as OpenpyxlInlineFont
    from wolfxl.cell.text import InlineFont

    reference = OpenpyxlInlineFont(**flags)
    font = InlineFont(**flags)
    assert (font.b, font.i) == (reference.b, reference.i)
    assert tostring(font.to_tree()) == tostring(reference.to_tree())
