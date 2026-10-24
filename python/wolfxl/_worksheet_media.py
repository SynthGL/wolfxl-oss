"""Worksheet chart, pivot, slicer, and image queue helpers."""

from __future__ import annotations

import re
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from wolfxl._worksheet import Worksheet

_PENDING_IMAGE_DELETIONS: dict[int, list[int]] = {}


def _is_alias_instance(obj: Any, class_name: str, module_suffix: str) -> bool:
    return any(
        cls.__name__ == class_name and cls.__module__.endswith(module_suffix)
        for cls in type(obj).__mro__
    )


def add_chart(ws: Worksheet, chart: Any, anchor: Any = None) -> None:
    """Queue a chart for write-mode or modify-mode save processing."""
    from wolfxl.chart._chart import ChartBase as _ChartBase

    if not isinstance(chart, _ChartBase) and not _is_alias_instance(
        chart, "ChartBase", "chart._chart"
    ):
        raise TypeError(
            f"add_chart expected wolfxl.chart.ChartBase, got "
            f"{type(chart).__name__}"
        )

    if anchor is None:
        anchor = chart.anchor if chart.anchor is not None else "E15"

    if isinstance(anchor, str):
        validate_a1_anchor(anchor)

    chart.anchor = anchor
    chart._anchor = anchor  # noqa: SLF001
    ws._pending_charts.append(chart)  # noqa: SLF001
    if ws._charts_cache is not None:  # noqa: SLF001
        ws._charts_cache.append(chart)  # noqa: SLF001


def add_pivot_table(ws: Worksheet, pivot_table: Any) -> None:
    """Queue a pivot table for modify-mode save processing."""
    from wolfxl.pivot import PivotTable as _PivotTable

    if not isinstance(pivot_table, _PivotTable):
        raise TypeError(
            f"add_pivot_table expected wolfxl.pivot.PivotTable, "
            f"got {type(pivot_table).__name__}"
        )
    wb_obj = ws._workbook  # noqa: SLF001
    if wb_obj._rust_patcher is None and wb_obj._rust_writer is None:  # noqa: SLF001
        raise RuntimeError(
            "add_pivot_table requires a Workbook in write or modify mode"
        )
    if pivot_table.cache._cache_id is None:  # noqa: SLF001
        raise ValueError(
            "PivotTable.cache has not been registered with the "
            "workbook yet. Call Workbook.add_pivot_cache(cache) "
            "before Worksheet.add_pivot_table(pt)."
        )
    if hasattr(pivot_table, "_compute_layout"):
        pivot_table._compute_layout()
    ws._pending_pivot_tables.append(pivot_table)  # noqa: SLF001


def add_slicer(ws: Worksheet, slicer: Any, anchor: str) -> None:
    """Queue a slicer presentation for modify-mode save processing."""
    from wolfxl.pivot import Slicer as _Slicer

    if not isinstance(slicer, _Slicer):
        raise TypeError(
            f"add_slicer expected wolfxl.pivot.Slicer, got "
            f"{type(slicer).__name__}"
        )
    if ws._workbook._rust_patcher is None:  # noqa: SLF001
        raise RuntimeError(
            "add_slicer requires modify mode — open the workbook "
            "with load_workbook(..., modify=True)."
        )
    if slicer.cache._slicer_cache_id is None:  # noqa: SLF001
        raise ValueError(
            "Slicer.cache has not been registered with the "
            "workbook yet. Call Workbook.add_slicer_cache(cache) "
            "before Worksheet.add_slicer(slicer, anchor)."
        )
    if not isinstance(anchor, str) or not anchor:
        raise ValueError("Worksheet.add_slicer: anchor must be a non-empty A1 string")
    validate_a1_anchor(anchor)
    slicer.anchor = anchor
    ws._pending_slicers.append(slicer)  # noqa: SLF001


def validate_a1_anchor(anchor: str) -> None:
    """Raise ValueError when *anchor* is not a valid single A1 cell ref."""
    if not anchor:
        raise ValueError("anchor must not be empty")
    match = re.match(r"^([A-Z]+)([0-9]+)$", anchor)
    if not match:
        raise ValueError(
            f"anchor={anchor!r} must be a single A1 cell ref like 'E15' "
            f"(regex ^[A-Z]+[0-9]+$); for ranged or absolute placement "
            f"pass an OneCellAnchor / TwoCellAnchor / AbsoluteAnchor"
        )
    col_letters, row_str = match.group(1), match.group(2)
    col_idx = 0
    for char in col_letters:
        col_idx = col_idx * 26 + (ord(char) - ord("A") + 1)
    if col_idx > 16_384:
        raise ValueError(
            f"anchor={anchor!r}: column {col_letters!r} exceeds Excel max XFD (16384)"
        )
    row_idx = int(row_str)
    if row_idx < 1 or row_idx > 1_048_576:
        raise ValueError(
            f"anchor={anchor!r}: row {row_idx} out of Excel range [1, 1048576]"
        )


def remove_chart(ws: Worksheet, chart: Any) -> None:
    """Remove a chart from this worksheet."""
    try:
        ws._pending_charts.remove(chart)  # noqa: SLF001
        return
    except ValueError:
        pass
    meta = getattr(chart, "_wolfxl_source_chart", None)
    if meta is None:
        raise ValueError(
            "chart was not added to this worksheet via add_chart() "
            "(or has already been removed)."
        ) from None
    _pending_chart_deletions(ws).append(meta)
    if ws._charts_cache is not None:  # noqa: SLF001
        idx = _find_identity_index(ws._charts_cache, chart)  # noqa: SLF001
        if idx is not None:
            del ws._charts_cache[idx]  # noqa: SLF001


def replace_chart(ws: Worksheet, old: Any, new: Any) -> None:
    """Replace one not-yet-flushed chart with another."""
    from wolfxl.chart._chart import ChartBase as _ChartBase

    if not isinstance(new, _ChartBase):
        raise TypeError(
            f"replace_chart expected wolfxl.chart.ChartBase for new, got "
            f"{type(new).__name__}"
        )
    try:
        index = ws._pending_charts.index(old)  # noqa: SLF001
    except ValueError:
        meta = getattr(old, "_wolfxl_source_chart", None)
        if meta is None:
            raise ValueError(
                "old chart was not added to this worksheet via add_chart()"
            ) from None
        anchor = new._anchor if new._anchor is not None else _a1_from_anchor(old._anchor)  # noqa: SLF001
        if anchor is None:
            anchor = old._anchor  # noqa: SLF001
        if anchor is None:
            anchor = "E15"
        if isinstance(anchor, str):
            validate_a1_anchor(anchor)
        new._anchor = anchor  # noqa: SLF001
        new._wolfxl_source_chart = meta  # noqa: SLF001
        new._wolfxl_source_title = _title_signature(new)  # noqa: SLF001
        if ws._charts_cache is not None:  # noqa: SLF001
            old_index = _find_identity_index(ws._charts_cache, old)  # noqa: SLF001
            if old_index is not None:
                ws._charts_cache[old_index] = new  # noqa: SLF001
        _queue_source_chart_op(ws, {"op": "replace", "meta": meta, "chart": new})
        return
    anchor = new._anchor if new._anchor is not None else old._anchor  # noqa: SLF001
    if anchor is None:
        anchor = "E15"
    if isinstance(anchor, str):
        validate_a1_anchor(anchor)
    new._anchor = anchor  # noqa: SLF001
    ws._pending_charts[index] = new  # noqa: SLF001


def pop_pending_chart_deletions(ws: Worksheet) -> list[dict[str, str]]:
    """Drain queued source-chart deletions for *ws* in append order."""
    pending = list(ws._pending_chart_deletions)  # noqa: SLF001
    ws._pending_chart_deletions.clear()  # noqa: SLF001
    return list(pending)


def get_charts(ws: Worksheet) -> list[Any]:
    """Return charts attached to this worksheet, hydrating read-mode drawings."""
    workbook = ws._workbook  # noqa: SLF001
    reader = getattr(workbook, "_rust_reader", None)
    if reader is None or not hasattr(reader, "read_charts"):
        return ws._pending_charts  # noqa: SLF001

    if ws._charts_cache is None:  # noqa: SLF001
        charts = []
        source_refs = _source_chart_refs_for_ws(ws)
        for idx, payload in enumerate(reader.read_charts(ws._title)):  # noqa: SLF001
            if isinstance(payload, dict):
                chart = _chart_from_payload(payload)
                if chart is not None:
                    if idx < len(source_refs):
                        chart._wolfxl_source_chart = source_refs[idx]  # noqa: SLF001
                        chart._wolfxl_source_title = _title_signature(chart)  # noqa: SLF001
                    charts.append(chart)
        charts.extend(ws._pending_charts)  # noqa: SLF001
        ws._charts_cache = charts  # noqa: SLF001
    return ws._charts_cache  # noqa: SLF001


def _source_chart_refs_for_ws(ws: Worksheet) -> list[dict[str, str]]:
    workbook = ws._workbook  # noqa: SLF001
    if getattr(workbook, "_rust_patcher", None) is None:
        return []
    from wolfxl import _source_charts

    return _source_charts.source_chart_refs(
        getattr(workbook, "_source_path", None), ws._title  # noqa: SLF001
    )


def _queue_source_chart_op(ws: Worksheet, op: dict[str, Any]) -> None:
    workbook = ws._workbook  # noqa: SLF001
    if getattr(workbook, "_rust_patcher", None) is None:
        raise RuntimeError(
            "source chart mutation requires modify mode — open the workbook "
            "with load_workbook(..., modify=True)."
        )
    workbook._pending_source_chart_ops.append(op)  # noqa: SLF001


def _title_signature(chart: Any) -> Any:
    title = getattr(chart, "title", None)
    if title is None:
        return None
    to_dict = getattr(title, "to_dict", None)
    if to_dict is not None:
        try:
            return to_dict()
        except Exception:
            pass
    return str(title)


def _chart_from_payload(payload: dict[str, Any]) -> Any:
    from wolfxl.chart import (
        AreaChart,
        AreaChart3D,
        BarChart,
        BarChart3D,
        BubbleChart,
        DoughnutChart,
        LineChart,
        LineChart3D,
        PieChart,
        PieChart3D,
        ProjectedPieChart,
        RadarChart,
        ScatterChart,
        StockChart,
        SurfaceChart,
        SurfaceChart3D,
    )

    classes = {
        "area": AreaChart,
        "area3d": AreaChart3D,
        "bar": BarChart,
        "bar3d": BarChart3D,
        "bubble": BubbleChart,
        "doughnut": DoughnutChart,
        "line": LineChart,
        "line3d": LineChart3D,
        "pie": PieChart,
        "pie3d": PieChart3D,
        "of_pie": ProjectedPieChart,
        "radar": RadarChart,
        "scatter": ScatterChart,
        "stock": StockChart,
        "surface": SurfaceChart,
        "surface3d": SurfaceChart3D,
    }
    chart_cls = classes.get(str(payload.get("kind") or ""))
    if chart_cls is None:
        return None

    chart = chart_cls()
    if payload.get("title") is not None:
        chart.title = str(payload["title"])
    if payload.get("x_axis_title") is not None and hasattr(chart, "x_axis"):
        chart.x_axis.title = str(payload["x_axis_title"])
    if payload.get("y_axis_title") is not None and hasattr(chart, "y_axis"):
        chart.y_axis.title = str(payload["y_axis_title"])
    if isinstance(payload.get("x_axis"), dict) and hasattr(chart, "x_axis"):
        _hydrate_axis_from_payload(chart.x_axis, payload["x_axis"])
    if isinstance(payload.get("y_axis"), dict) and hasattr(chart, "y_axis"):
        _hydrate_axis_from_payload(chart.y_axis, payload["y_axis"])
    if isinstance(payload.get("data_labels"), dict):
        chart.dLbls = _data_labels_from_payload(payload["data_labels"])
    if payload.get("legend_position") is not None and chart.legend is not None:
        chart.legend.position = str(payload["legend_position"])
    if payload.get("bar_dir") is not None and hasattr(chart, "barDir"):
        chart.barDir = str(payload["bar_dir"])
    if payload.get("grouping") is not None and hasattr(chart, "grouping"):
        chart.grouping = str(payload["grouping"])
    if payload.get("scatter_style") is not None and hasattr(chart, "scatterStyle"):
        chart.scatterStyle = str(payload["scatter_style"])
    if payload.get("vary_colors") is not None and hasattr(chart, "varyColors"):
        chart.varyColors = bool(payload["vary_colors"])
    if payload.get("style") is not None:
        chart.style = int(payload["style"])
    chart._anchor = _anchor_from_payload(payload.get("anchor"))  # noqa: SLF001
    chart._source_drawing_path = payload.get("source_drawing_path")  # noqa: SLF001
    chart._source_chart_rid = payload.get("source_chart_rid")  # noqa: SLF001
    chart._source_chart_path = payload.get("source_chart_path")  # noqa: SLF001
    chart.ser = [
        _series_from_payload(str(payload.get("kind") or ""), series_payload)
        for series_payload in payload.get("series", [])
        if isinstance(series_payload, dict)
    ]
    subcharts = payload.get("subcharts")
    if isinstance(subcharts, list) and len(subcharts) > 1:
        for subchart_payload in subcharts[1:]:
            if not isinstance(subchart_payload, dict):
                continue
            secondary = _chart_from_payload(subchart_payload)
            if secondary is not None:
                chart += secondary
    return chart


def _hydrate_axis_from_payload(axis: Any, payload: dict[str, Any]) -> None:
    from wolfxl.chart.axis import DisplayUnitsLabelList
    from wolfxl.chart.data_source import NumFmt

    scalar_attrs = {
        "ax_id": "axId",
        "cross_ax": "crossAx",
        "axis_position": "axPos",
        "major_unit": "majorUnit",
        "minor_unit": "minorUnit",
        "tick_lbl_pos": "tickLblPos",
        "major_tick_mark": "majorTickMark",
        "minor_tick_mark": "minorTickMark",
        "crosses": "crosses",
        "crosses_at": "crossesAt",
        "cross_between": "crossBetween",
    }
    for key, attr in scalar_attrs.items():
        if payload.get(key) is not None and hasattr(axis, attr):
            setattr(axis, attr, payload[key])

    if payload.get("num_format_code") is not None:
        axis.numFmt = NumFmt(
            formatCode=str(payload["num_format_code"]),
            sourceLinked=bool(payload.get("num_format_source_linked") or False),
        )

    scaling = getattr(axis, "scaling", None)
    if scaling is not None:
        scaling_attrs = {
            "scaling_min": "min",
            "scaling_max": "max",
            "scaling_orientation": "orientation",
            "scaling_log_base": "logBase",
        }
        for key, attr in scaling_attrs.items():
            if payload.get(key) is not None:
                setattr(scaling, attr, payload[key])

    if payload.get("display_unit") is not None and hasattr(axis, "dispUnits"):
        axis.dispUnits = DisplayUnitsLabelList(builtInUnit=str(payload["display_unit"]))


def _series_from_payload(kind: str, payload: dict[str, Any]) -> Any:
    from wolfxl.chart.data_source import AxDataSource, NumDataSource, NumRef, StrRef
    from wolfxl.chart.series import Series, SeriesLabel, XYSeries

    is_xy = kind in {"scatter", "bubble"}
    series = XYSeries() if is_xy else Series()
    if payload.get("idx") is not None:
        series.idx = int(payload["idx"])
    if payload.get("order") is not None:
        series.order = int(payload["order"])
    if payload.get("title_ref"):
        series.tx = SeriesLabel(strRef=StrRef(str(payload["title_ref"])))
    elif payload.get("title_value"):
        series.tx = SeriesLabel(v=str(payload["title_value"]))

    if is_xy:
        if payload.get("x_ref"):
            series.xVal = AxDataSource(numRef=NumRef(f=str(payload["x_ref"])))
        if payload.get("y_ref"):
            series.yVal = NumDataSource(numRef=NumRef(f=str(payload["y_ref"])))
        if payload.get("bubble_size_ref"):
            series.bubbleSize = NumDataSource(
                numRef=NumRef(f=str(payload["bubble_size_ref"]))
            )
    else:
        if payload.get("cat_ref"):
            series.cat = AxDataSource(strRef=StrRef(str(payload["cat_ref"])))
        if payload.get("val_ref"):
            series.val = NumDataSource(numRef=NumRef(f=str(payload["val_ref"])))
    if isinstance(payload.get("graphical_properties"), dict):
        series.spPr = _graphical_properties_from_payload(payload["graphical_properties"])
    if isinstance(payload.get("data_labels"), dict):
        series.dLbls = _data_labels_from_payload(payload["data_labels"])
    if isinstance(payload.get("trendline"), dict):
        series.trendline = _trendline_from_payload(payload["trendline"])
    if isinstance(payload.get("error_bars"), dict):
        series.errBars = _error_bars_from_payload(payload["error_bars"])
    return series


def _graphical_properties_from_payload(payload: dict[str, Any]) -> Any:
    from wolfxl.chart.shapes import GraphicalProperties, LineProperties

    line_payload = payload.get("ln")
    line = None
    if isinstance(line_payload, dict) and any(
        line_payload.get(key) is not None
        for key in ("no_fill", "solid_fill", "prst_dash", "w_emu")
    ):
        line = LineProperties(
            noFill=line_payload.get("no_fill"),
            solidFill=line_payload.get("solid_fill"),
            prstDash=line_payload.get("prst_dash"),
            w=line_payload.get("w_emu"),
        )
    return GraphicalProperties(
        noFill=payload.get("no_fill"),
        solidFill=payload.get("solid_fill"),
        ln=line,
    )


def _data_labels_from_payload(payload: dict[str, Any]) -> Any:
    from wolfxl.chart.label import DataLabelList

    kwargs = {
        "position": payload.get("position"),
        "showLegendKey": payload.get("show_legend_key"),
        "showVal": payload.get("show_val"),
        "showCatName": payload.get("show_cat_name"),
        "showSerName": payload.get("show_ser_name"),
        "showPercent": payload.get("show_percent"),
        "showBubbleSize": payload.get("show_bubble_size"),
        "showLeaderLines": payload.get("show_leader_lines"),
    }
    return DataLabelList(**{key: value for key, value in kwargs.items() if value is not None})


def _trendline_from_payload(payload: dict[str, Any]) -> Any:
    from wolfxl.chart.trendline import Trendline

    return Trendline(
        trendlineType=str(payload.get("trendline_type") or "linear"),
        order=payload.get("order"),
        period=payload.get("period"),
        forward=payload.get("forward"),
        backward=payload.get("backward"),
        intercept=payload.get("intercept"),
        dispEq=payload.get("display_equation"),
        dispRSqr=payload.get("display_r_squared"),
    )


def _error_bars_from_payload(payload: dict[str, Any]) -> Any:
    from wolfxl.chart.error_bar import ErrorBars

    return ErrorBars(
        errDir=payload.get("direction"),
        errBarType=str(payload.get("bar_type") or "both"),
        errValType=str(payload.get("val_type") or "fixedVal"),
        noEndCap=payload.get("no_end_cap"),
        val=payload.get("val"),
    )


def add_image(ws: Worksheet, image: Any, anchor: Any = None) -> None:
    """Queue an image for write-mode or modify-mode save processing."""
    from wolfxl.drawing.image import Image as _Image

    if not isinstance(image, _Image) and not _is_alias_instance(
        image, "Image", "drawing.image"
    ):
        raise TypeError(
            f"add_image expected wolfxl.drawing.image.Image, got {type(image).__name__}"
        )

    if anchor is None:
        anchor = "A1"

    image.anchor = anchor
    ws._pending_images.append(image)  # noqa: SLF001
    if ws._images_cache is not None:  # noqa: SLF001
        ws._images_cache.append(image)  # noqa: SLF001


def remove_image(ws: Worksheet, index_or_image: int | Any) -> None:
    """Remove one image from this worksheet by index or object identity."""
    images = get_images(ws)
    index = _resolve_image_index(images, index_or_image)
    target = images[index]

    pending_index = _find_identity_index(ws._pending_images, target)  # noqa: SLF001
    if pending_index is not None:
        del ws._pending_images[pending_index]  # noqa: SLF001
    else:
        source_index = _source_image_index_for_cache_position(ws, index, images)
        _pending_image_deletions(ws).append(source_index)

    if ws._images_cache is not None:  # noqa: SLF001
        del ws._images_cache[index]  # noqa: SLF001


def replace_image(ws: Worksheet, index_or_image: int | Any, new_image: Any) -> None:
    """Replace one attached image using remove + add semantics."""
    from wolfxl.drawing.image import Image as _Image

    if not isinstance(new_image, _Image):
        raise TypeError(
            "replace_image expected wolfxl.drawing.image.Image for new_image, "
            f"got {type(new_image).__name__}"
        )

    images = get_images(ws)
    index = _resolve_image_index(images, index_or_image)
    old_image = images[index]

    anchor = new_image.anchor if new_image.anchor is not None else old_image.anchor
    if anchor is None:
        anchor = "A1"
    if isinstance(anchor, str):
        validate_a1_anchor(anchor)
    new_image.anchor = anchor

    remove_image(ws, index)

    ws._pending_images.append(new_image)  # noqa: SLF001
    if ws._images_cache is not None:  # noqa: SLF001
        ws._images_cache.insert(index, new_image)  # noqa: SLF001


def pop_pending_image_deletions(ws: Worksheet) -> list[int]:
    """Drain queued source-image deletions for *ws* in append order."""
    pending = _PENDING_IMAGE_DELETIONS.pop(id(ws), [])
    return list(pending)


def has_pending_image_deletions(ws: Worksheet) -> bool:
    """Return whether source-image deletions are queued for *ws*."""
    return bool(_PENDING_IMAGE_DELETIONS.get(id(ws)))


def get_images(ws: Worksheet) -> list[Any]:
    """Return images attached to this worksheet, hydrating read-mode drawings."""
    workbook = ws._workbook  # noqa: SLF001
    reader = getattr(workbook, "_rust_reader", None)
    if reader is None or not hasattr(reader, "read_images"):
        return ws._pending_images  # noqa: SLF001

    if ws._images_cache is None:  # noqa: SLF001
        images = [
            _image_from_payload(payload)
            for payload in reader.read_images(ws._title)  # noqa: SLF001
            if isinstance(payload, dict)
        ]
        images.extend(ws._pending_images)  # noqa: SLF001
        ws._images_cache = images  # noqa: SLF001
    return ws._images_cache  # noqa: SLF001


def _pending_image_deletions(ws: Worksheet) -> list[int]:
    return _PENDING_IMAGE_DELETIONS.setdefault(id(ws), [])


def _pending_chart_deletions(ws: Worksheet) -> list[dict[str, str]]:
    return ws._pending_chart_deletions  # noqa: SLF001


def _source_chart_meta(chart: Any) -> dict[str, str] | None:
    drawing_path = getattr(chart, "_source_drawing_path", None)
    chart_rid = getattr(chart, "_source_chart_rid", None)
    chart_path = getattr(chart, "_source_chart_path", None)
    if not drawing_path or not chart_rid or not chart_path:
        return None
    return {
        "drawing_path": str(drawing_path),
        "chart_rid": str(chart_rid),
        "chart_path": str(chart_path),
    }


def _a1_from_anchor(anchor: Any) -> str | None:
    if isinstance(anchor, str):
        return anchor
    marker = getattr(anchor, "_from", None)
    if marker is None:
        return None
    col = int(getattr(marker, "col", 0)) + 1
    row = int(getattr(marker, "row", 0)) + 1
    letters = ""
    while col:
        col, rem = divmod(col - 1, 26)
        letters = chr(ord("A") + rem) + letters
    return f"{letters}{row}"


def _resolve_image_index(images: list[Any], index_or_image: int | Any) -> int:
    if isinstance(index_or_image, int):
        index = index_or_image
        if index < 0:
            index += len(images)
        if index < 0 or index >= len(images):
            raise ValueError(
                f"image index {index_or_image} out of range for {len(images)} images"
            )
        return index

    index = _find_identity_index(images, index_or_image)
    if index is None:
        raise ValueError("image is not attached to this worksheet")
    return index


def _find_identity_index(items: list[Any], needle: Any) -> int | None:
    for idx, item in enumerate(items):
        if item is needle:
            return idx
    return None


def _source_image_index_for_cache_position(
    ws: Worksheet,
    cache_index: int,
    images: list[Any],
) -> int:
    pending_ids = {id(image) for image in ws._pending_images}  # noqa: SLF001
    source_index = 0
    for idx, image in enumerate(images):
        if id(image) in pending_ids:
            continue
        if idx == cache_index:
            return source_index
        source_index += 1
    raise ValueError("image selection does not resolve to a source workbook image")


def _image_from_payload(payload: dict[str, Any]) -> Any:
    from wolfxl.drawing.image import Image

    image = Image(payload["data"])
    image.anchor = _anchor_from_payload(payload.get("anchor"))
    return image


def _anchor_from_payload(payload: Any) -> Any:
    if not isinstance(payload, dict):
        return None

    from wolfxl.drawing.spreadsheet_drawing import (
        AbsoluteAnchor,
        OneCellAnchor,
        TwoCellAnchor,
        XDRPoint2D,
        XDRPositiveSize2D,
    )

    kind = payload.get("type")
    if kind == "one_cell":
        ext = _extent_from_payload(payload)
        anchor = OneCellAnchor(
            _from=_marker_from_payload(payload, "from"),
            ext=ext,
        )
        if ext is None:
            anchor.__dict__["ext"] = None
        return anchor
    if kind == "two_cell":
        return TwoCellAnchor(
            _from=_marker_from_payload(payload, "from"),
            to=_marker_from_payload(payload, "to"),
            editAs=str(payload.get("edit_as") or "oneCell"),
        )
    if kind == "absolute":
        return AbsoluteAnchor(
            pos=XDRPoint2D(
                x=int(payload.get("x_emu") or 0),
                y=int(payload.get("y_emu") or 0),
            ),
            ext=_extent_from_payload(payload) or XDRPositiveSize2D(cx=0, cy=0),
        )
    return None


def _marker_from_payload(payload: dict[str, Any], prefix: str) -> Any:
    from wolfxl.drawing.spreadsheet_drawing import AnchorMarker

    return AnchorMarker(
        col=int(payload.get(f"{prefix}_col") or 0),
        row=int(payload.get(f"{prefix}_row") or 0),
        colOff=int(payload.get(f"{prefix}_col_off") or 0),
        rowOff=int(payload.get(f"{prefix}_row_off") or 0),
    )


def _extent_from_payload(payload: dict[str, Any]) -> Any:
    from wolfxl.drawing.spreadsheet_drawing import XDRPositiveSize2D

    cx = payload.get("cx_emu")
    cy = payload.get("cy_emu")
    if cx is None or cy is None:
        return None
    if int(cx) == 0 and int(cy) == 0:
        return None
    return XDRPositiveSize2D(cx=int(cx), cy=int(cy))
