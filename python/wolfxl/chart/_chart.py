"""`ChartBase` — common ancestor for every chart type.

Mirrors :class:`openpyxl.chart._chart.ChartBase`. Every chart subclass
inherits ``title``, ``legend``, ``layout``, ``style``, ``display_blanks``,
``visible_cells_only``, ``roundedCorners``, ``graphical_properties``, plus
the ``add_data``/``set_categories``/``append`` helpers that drive the
fluent construction pattern.

The :meth:`to_rust_dict` method is the contract surface the Rust
emitter consumes. It returns a flat-shape dict of snake_case keys.
"""

from __future__ import annotations

import re
from collections import OrderedDict
from operator import attrgetter
from typing import Any

from wolfxl._compat import _openpyxl_name_fallback, _resolve_openpyxl_class

from .data_source import AxDataSource, NumRef
from .layout import Layout
from .legend import Legend
from .reference import Reference
from .series import Series, SeriesFactory, attribute_mapping
from .shapes import GraphicalProperties
from .title import TitleDescriptor


_VALID_DISPLAY_BLANKS = ("span", "gap", "zero")

# `pivot_source.name` regex. Optional sheet prefix + table name. Sheet
# prefix only allows the conservative identifier set; table-name segment
# additionally allows spaces (Excel pivot names like "PivotTable 1" are
# commonplace in the wild).
_PIVOT_SOURCE_NAME_RE = re.compile(
    r"^([A-Za-z_][A-Za-z0-9_]*!)?[A-Za-z_][A-Za-z0-9_ ]*$"
)
_PIVOT_FMT_ID_MAX = 65535


# Map openpyxl tagname (`barChart`, `bar3DChart`, …) to the short
# kind string consumed by the Rust emitter. Both 2D and 3D variants
# are covered.
_TAGNAME_TO_KIND = {
    "barChart": "bar",
    "bar3DChart": "bar3d",
    "lineChart": "line",
    "line3DChart": "line3d",
    "pieChart": "pie",
    "pie3DChart": "pie3d",
    "ofPieChart": "of_pie",
    "doughnutChart": "doughnut",
    "areaChart": "area",
    "area3DChart": "area3d",
    "scatterChart": "scatter",
    "bubbleChart": "bubble",
    "radarChart": "radar",
    "surfaceChart": "surface",
    "surface3DChart": "surface3d",
    "stockChart": "stock",
}


def _anchor_payload(anchor: Any) -> Any:
    """Return the emitter shape for a chart anchor.

    A1 strings and ``None`` pass through; ``OneCellAnchor`` /
    ``TwoCellAnchor`` / ``AbsoluteAnchor`` objects become the flat anchor
    dict shared with images, so ranged placement survives the save.
    """
    if anchor is None or isinstance(anchor, str):
        return anchor
    from wolfxl._images import _resolve_anchor_dict

    return _resolve_anchor_dict(anchor, 0, 0)


class _ChartMeta(type):
    def __instancecheck__(cls, instance: Any) -> bool:
        if type.__instancecheck__(cls, instance):
            return True
        instance_cls = getattr(instance, "__class__", None)
        return (
            getattr(cls, "__name__", None) == getattr(instance_cls, "__name__", None)
            and ".chart." in str(getattr(cls, "__module__", ""))
            and ".chart." in str(getattr(instance_cls, "__module__", ""))
        )


class ChartBase(metaclass=_ChartMeta):
    """Base class for all chart kinds.

    Subclasses provide:
    - ``tagname`` — XML root for this chart's plot block (e.g. ``"barChart"``).
    - ``_series_type`` — key into :data:`series.attribute_mapping`.
    - ``_chart_type_specific_keys()`` — flat dict of per-type keys merged
      into the top-level shape.
    """

    title = TitleDescriptor()

    # Default tagname / series-type — subclasses override.
    tagname: str = "chart"
    _series_type: str = ""

    # Anchor + dimension defaults match openpyxl
    anchor: str | Any = "E15"
    width: float = 15.0  # cm
    height: float = 7.5  # cm
    _path: str = "/xl/charts/chart{0}.xml"
    mime_type: str = (
        "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
    )

    def __init__(self, axId: tuple[int, ...] = (), **kw: Any) -> None:
        # Per-instance state (kept off __slots__ so subclasses can add).
        self._charts: list[ChartBase] = [self]
        self.title = None  # via TitleDescriptor; user can later assign str/Title
        self.layout: Layout | None = None
        self.roundedCorners: bool | None = None
        self.legend: Legend | None = Legend()
        self.graphical_properties: GraphicalProperties | None = None
        self._style: int | None = None
        from .chartspace import PlotArea

        self.plot_area = PlotArea()
        self.axId = list(axId) if axId else []
        self.display_blanks: str = "gap"
        # Internal storage for the snake-case ``pivot_source`` attribute
        # is a dict shaped like the chart-pivot-source contract (or
        # ``None``). The legacy ``pivotSource`` openpyxl alias (typed as
        # ``Any``) is preserved for back-compat with callers that
        # imported the openpyxl PivotSource class directly; it does not
        # flow through ``to_rust_dict``.
        self._pivot_source: dict[str, Any] | None = None
        self.pivotSource: Any | None = None
        self.pivotFormats: tuple[Any, ...] = ()
        self.visible_cells_only: bool = True
        self.idx_base: int = 0

        # ser is the storage; ``series`` is the openpyxl alias.
        self.ser: list[Series] = []

        # Per-type defaults can be set in subclass __init__ before super().__init__.
        # We swallow remaining kwargs that subclasses already bound on themselves.
        for key, val in kw.items():
            if not hasattr(self, key):
                setattr(self, key, val)

        # Anchor for placement on the worksheet — set by ``ws.add_chart``.
        self._anchor: Any = None

    # ------------------------------------------------------------------
    # ``style`` is bounded 1..48 in the spec.
    # ------------------------------------------------------------------
    @property
    def path(self) -> str:
        return self._path.format(getattr(self, "_id", 1))

    def to_tree(
        self,
        namespace: str | None = None,
        tagname: str | None = None,
        idx: int | None = None,
    ):
        if self.__class__ is ChartBase and self.tagname == "chart":
            raise NotImplementedError
        self.axId = [axis_id for axis_id in self._axes]
        if self.ser is not None:
            for series in self.ser:
                series.__elements__ = attribute_mapping.get(self._series_type, ())
        from openpyxl.xml.functions import Element

        attrs = {key: value for key, value in self}
        return Element(tagname or self.tagname, attrs)

    def _write(self):
        from ._3d import _3DBase
        from openpyxl.xml.constants import CHART_NS
        from openpyxl.xml.functions import Element, SubElement

        self.plot_area.layout = self.layout

        idx_base = self.idx_base
        for chart in self._charts:
            chart._reindex()
            chart.axId = [axis_id for axis_id in chart._axes]
            for series in chart.ser:
                series.__elements__ = attribute_mapping.get(chart._series_type, ())
            if chart not in self.plot_area._charts:
                chart.idx_base = idx_base
                idx_base += len(chart.series)
        self.plot_area._charts = self._charts
        self.plot_area._axes = list(self._axes.values())

        last_chart = self._charts[-1]
        chartspace = Element("chartSpace", {"xmlns": CHART_NS})
        if self.pivotSource is not None:
            chartspace.append(self.pivotSource.to_tree())
        if self.style is not None:
            SubElement(chartspace, "style", {"val": str(self.style)})
        if self.roundedCorners is not None:
            SubElement(
                chartspace,
                "roundedCorners",
                {"val": "1" if self.roundedCorners else "0"},
            )

        chart_node = SubElement(chartspace, "chart")
        if self.title is not None:
            chart_node.append(self.title.to_tree())

        if self.pivotFormats:
            pivot_fmts = SubElement(chart_node, "pivotFmts")
            for fmt in self.pivotFormats:
                pivot_fmts.append(fmt.to_tree("pivotFmt"))

        plot_area = SubElement(chart_node, "plotArea")
        if self.layout is not None:
            plot_area.append(self.layout.to_tree())
        if isinstance(last_chart, _3DBase):
            if last_chart.view3D is not None:
                chart_node.append(last_chart.view3D.to_tree())
            for surface_name in ("floor", "sideWall", "backWall"):
                surface = getattr(last_chart, surface_name, None)
                if surface is not None:
                    chart_node.append(surface.to_tree(surface_name))
        for chart in self._charts:
            plot_area.append(chart.to_tree())
        for axis in self._axes.values():
            plot_area.append(axis.to_tree())

        if self.legend is not None:
            chart_node.append(self.legend.to_tree())
        SubElement(
            chart_node,
            "plotVisOnly",
            {"val": "1" if self.visible_cells_only else "0"},
        )
        SubElement(chart_node, "dispBlanksAs", {"val": self.display_blanks})
        if self.graphical_properties is not None:
            chartspace.append(self.graphical_properties.to_tree())
        return chartspace

    def __iter__(self):
        if self.__class__ is not ChartBase:
            return
        yield "visible_cells_only", "1" if self.visible_cells_only else "0"
        yield "display_blanks", self.display_blanks

    @property
    def style(self) -> int | None:
        """Return the chart style index.

        Returns:
            An Excel chart style number in the inclusive range ``1..48``, or
            ``None`` when no explicit style is set.
        """
        return self._style

    @style.setter
    def style(self, value: int | None) -> None:
        """Set the chart style index.

        Args:
            value: Excel chart style number in ``1..48``, or ``None`` to clear
                the explicit style.

        Raises:
            ValueError: If ``value`` falls outside Excel's supported style
                range.
        """
        if value is None:
            self._style = None
            return
        v = int(value)
        if not (1 <= v <= 48):
            raise ValueError(f"style={v} must be in [1, 48]")
        self._style = v

    # ------------------------------------------------------------------
    # ``display_blanks`` is bounded to ('span', 'gap', 'zero').
    # ------------------------------------------------------------------
    @property
    def display_blanks(self) -> str:
        return self._display_blanks

    @display_blanks.setter
    def display_blanks(self, value: str) -> None:
        if value not in _VALID_DISPLAY_BLANKS:
            raise ValueError(f"display_blanks={value!r} not in {_VALID_DISPLAY_BLANKS}")
        self._display_blanks = value

    # ------------------------------------------------------------------
    # ``pivot_source``.
    #
    # Linking a chart to a pivot table is what makes Excel render it as
    # a "pivot chart" (right-click → Refresh, pivot-aware toolbar, etc.).
    # The OOXML serialization is a top-of-`<c:chart>` ``<c:pivotSource>``
    # block + an extra ``<c:fmtId val="0"/>`` on every series. The Rust
    # emitter handles both; this attribute is the Python surface.
    # ------------------------------------------------------------------
    @property
    def pivot_source(self) -> dict[str, Any] | None:
        """The chart's pivot-source linkage as a dict, or
        ``None`` if unlinked.
        """
        return self._pivot_source

    @pivot_source.setter
    def pivot_source(self, value: Any) -> None:
        if value is None:
            self._pivot_source = None
            return

        # Tuple form: (name, fmt_id).
        if isinstance(value, tuple):
            if len(value) != 2:
                raise ValueError(
                    "Chart.pivot_source tuple must be (name, fmt_id); "
                    f"got tuple of length {len(value)}"
                )
            name, fmt_id = value
            self._pivot_source = self._validate_pivot_source(name, fmt_id)
            return

        # Dict form (round-tripped from to_rust_dict).
        if isinstance(value, dict):
            if "name" not in value:
                raise ValueError(
                    "Chart.pivot_source dict must include 'name'"
                )
            name = value["name"]
            fmt_id = value.get("fmt_id", 0)
            self._pivot_source = self._validate_pivot_source(name, fmt_id)
            return

        # Duck-typed PivotTable. We avoid a hard import to dodge the
        # circular wolfxl.pivot._table → wolfxl.chart import cycle. Any
        # object that quacks with a ``.name`` string attribute and isn't
        # a primitive is treated as a pivot table.
        name_attr = getattr(value, "name", None)
        if isinstance(name_attr, str):
            self._pivot_source = self._validate_pivot_source(name_attr, 0)
            return

        raise TypeError(
            "Chart.pivot_source accepts a PivotTable, (name, fmt_id) "
            "tuple, dict {'name': str, 'fmt_id': int}, or None; "
            f"got {type(value).__name__}"
        )

    @staticmethod
    def _validate_pivot_source(name: Any, fmt_id: Any) -> dict[str, Any]:
        """Validate and normalise a pivot-source linkage.

        Checks that ``name`` is a non-empty string matching the OOXML
        pivot-source name regex (optional sheet prefix plus a table-name
        segment that may contain spaces) and that ``fmt_id`` is an
        integer in ``[0, 65535]``. Returns the canonical
        ``{"name": str, "fmt_id": int}`` dict consumed by
        :meth:`to_rust_dict`. Called from the ``pivot_source`` setter
        whenever a user assigns a tuple, dict, or duck-typed pivot
        table; raises :class:`ValueError` on any out-of-range or
        malformed input.
        """
        if not isinstance(name, str) or not name:
            raise ValueError(
                "pivot_source.name must be a non-empty string"
            )
        if not _PIVOT_SOURCE_NAME_RE.match(name):
            raise ValueError(
                f"pivot_source.name={name!r} does not match the OOXML "
                f"pivot-source name regex"
            )
        try:
            fmt_id_int = int(fmt_id)
        except (TypeError, ValueError) as exc:
            raise ValueError(
                f"pivot_source.fmt_id must be an int, got {fmt_id!r}"
            ) from exc
        if not (0 <= fmt_id_int <= _PIVOT_FMT_ID_MAX):
            raise ValueError(
                f"pivot_source.fmt_id={fmt_id_int} must be in "
                f"[0, {_PIVOT_FMT_ID_MAX}]"
            )
        return {"name": name, "fmt_id": fmt_id_int}

    # ------------------------------------------------------------------
    # openpyxl alias: ``series`` is read/write on top of ``ser``.
    # ------------------------------------------------------------------
    @property
    def series(self) -> list[Series]:
        return self.ser

    @series.setter
    def series(self, value: list[Series]) -> None:
        self.ser = list(value)

    # openpyxl alias used by some legacy code paths
    @property
    def dataLabels(self) -> Any:
        return getattr(self, "dLbls", None)

    @dataLabels.setter
    def dataLabels(self, value: Any) -> None:
        self.dLbls = value

    # ------------------------------------------------------------------
    # Composition helpers
    # ------------------------------------------------------------------
    def __hash__(self) -> int:  # noqa: D401
        return id(self)

    def __iadd__(self, other: "ChartBase") -> "ChartBase":
        if not isinstance(other, ChartBase):
            raise TypeError("Only other charts can be added")
        self._charts.append(other)
        return self

    def append(self, value: Series) -> None:
        """Append a single :class:`Series` to ``self.series``."""
        if not isinstance(value, Series):
            raise TypeError(f"expected {Series}")
        self.ser = list(self.ser) + [value]

    def add_data(
        self,
        data: Any,
        from_rows: bool = False,
        titles_from_data: bool = False,
    ) -> None:
        """Add a range of data as one or more series.

        If ``from_rows`` is True, each row of *data* becomes a series;
        otherwise (default) each column becomes a series. ``titles_from_data``
        consumes the first cell of each row/column as the legend label
        (matching openpyxl exactly).
        """
        if not isinstance(data, Reference):
            data = Reference(range_string=data)
        values = data.rows if from_rows else data.cols
        for ref in values:
            series = SeriesFactory(ref, title_from_data=titles_from_data)
            self.ser.append(series)

    def set_categories(self, labels: Any) -> None:
        """Set the categories (x-axis labels) for every series."""
        if not isinstance(labels, Reference):
            labels = Reference(range_string=labels)
        cat = AxDataSource(numRef=NumRef(f=labels))
        for s in self.ser:
            s.cat = cat

    def _reindex(self) -> None:
        """Sort series by ``order`` and rebase indexes (matches openpyxl)."""
        ds = sorted(self.ser, key=attrgetter("order"))
        for idx, s in enumerate(ds):
            s.idx = idx
            s.order = idx
        self.ser = ds

    @property
    def _axes(self) -> "OrderedDict[int, Any]":
        x = getattr(self, "x_axis", None)
        y = getattr(self, "y_axis", None)
        z = getattr(self, "z_axis", None)
        return OrderedDict(
            [(axis.axId, axis) for axis in (x, y, z) if axis is not None]
        )

    # ------------------------------------------------------------------
    # Validation
    # ------------------------------------------------------------------
    def _validate_at_emit(self) -> None:
        """Raise at ``to_rust_dict`` time on bad state.

        For combination charts, additionally walk ``self._charts[1:]``
        and reject:
        - empty series in any secondary chart family,
        - Pie / Doughnut secondaries (out-of-scope),
        - a secondary whose ``(x_axis.axId, y_axis.axId)`` exactly equals
          the primary's *and* whose ``kind`` matches the primary's kind
          (likely a copy-paste bug - fail loudly).
        """
        if not self.ser:
            raise ValueError(
                f"{type(self).__name__} requires at least one series "
                f"(call chart.add_data(...) before saving)."
            )

        # Secondary-chart validation. The first entry of ``self._charts``
        # is ``self`` (primary); siblings live at ``[1:]``.
        secondaries = list(self._charts[1:])
        if not secondaries:
            return

        primary_kind = _TAGNAME_TO_KIND.get(self.tagname)
        primary_x_id = getattr(getattr(self, "x_axis", None), "axId", None)
        primary_y_id = getattr(getattr(self, "y_axis", None), "axId", None)

        for secondary in secondaries:
            sec_kind = _TAGNAME_TO_KIND.get(secondary.tagname)
            # (a) Empty series in a secondary.
            if not getattr(secondary, "ser", None):
                raise ValueError(
                    f"combination chart secondary {type(secondary).__name__} "
                    f"requires at least one series "
                    f"(call chart.add_data(...) on the secondary before saving)."
                )
            # (b) Pie/Doughnut secondaries — out of scope.
            if sec_kind in {"pie", "pie3d", "doughnut", "of_pie"}:
                raise ValueError(
                    f"combination chart cannot include a "
                    f"{type(secondary).__name__} secondary "
                    f"(Pie/Doughnut combos are out of scope)."
                )
            # (c) Same-kind, same-axId secondary: copy-paste smell.
            sec_x_id = getattr(getattr(secondary, "x_axis", None), "axId", None)
            sec_y_id = getattr(getattr(secondary, "y_axis", None), "axId", None)
            if (
                sec_kind == primary_kind
                and sec_x_id == primary_x_id
                and sec_y_id == primary_y_id
            ):
                raise ValueError(
                    f"combination chart secondary {type(secondary).__name__} "
                    f"has the same kind and axIds as the primary; this is "
                    f"likely a copy-paste bug. Set a distinct y_axis.axId "
                    f"on the secondary (or use a different chart kind)."
                )

    # ------------------------------------------------------------------
    # Rust-side serialisation
    # ------------------------------------------------------------------
    def _chart_type_specific_keys(self) -> dict[str, Any]:
        """Return per-type flat keys to merge into the §10.1 top-level dict.

        Replaces the v1.6.0 ``_chart_dict_extras`` envelope. Keys are
        snake_case, flat at top-level (no nesting). Subclasses override.
        """
        return {}

    # v1.6.0 hook kept as a no-op default for any subclass that hasn't
    # been migrated yet — never called by ``to_rust_dict`` itself.
    def _chart_dict_extras(self) -> dict[str, Any]:  # pragma: no cover
        return {}

    def to_rust_dict(self) -> dict[str, Any]:
        """Produce a typed dict describing this chart for the Rust emitter.

        Shape: flat top-level keys, snake_case throughout; no
        ``axes`` list, no ``extras`` envelope.
        """
        self._validate_at_emit()
        self._reindex()

        kind = _TAGNAME_TO_KIND.get(self.tagname)
        if kind is None:
            raise ValueError(
                f"Unknown chart tagname={self.tagname!r}; expected one of "
                f"{sorted(_TAGNAME_TO_KIND)}"
            )

        x_axis = getattr(self, "x_axis", None)
        y_axis = getattr(self, "y_axis", None)
        z_axis = getattr(self, "z_axis", None)

        series_list = [s.to_rust_dict(self._series_type) for s in self.ser]

        d: dict[str, Any] = {
            # Required
            "kind": kind,
            "series_type": self._series_type,
            "series": series_list,

            # Optional shared
            "style": self._style,
            "display_blanks_as": self._display_blanks,
            "plot_visible_only": self.visible_cells_only,
            "rounded_corners": self.roundedCorners,

            # Decorations
            "title": self.title.to_dict() if self.title is not None else None,
            "legend": self.legend.to_dict() if self.legend is not None else None,
            "layout": self.layout.to_dict() if self.layout is not None else None,
            "graphical_properties": (
                self.graphical_properties.to_dict()
                if self.graphical_properties is not None
                else None
            ),

            # Axes — flat keys (NOT a list)
            "x_axis": x_axis.to_dict() if x_axis is not None else None,
            "y_axis": y_axis.to_dict() if y_axis is not None else None,
            "z_axis": z_axis.to_dict() if z_axis is not None else None,

            # Anchor + dimensions
            "anchor": _anchor_payload(self._anchor),
            "width_emu": int(self.width * 360_000) if self.width is not None else None,
            "height_emu": int(self.height * 360_000) if self.height is not None else None,

            # ``None`` → no ``<c:pivotSource>`` block emitted; chart
            # is a standard chart. Dict shape
            # `{"name": str, "fmt_id": int}`.
            "pivot_source": self._pivot_source,
        }

        # Default vary_colors (None → omit; subclasses may set)
        vc = getattr(self, "vary_colors", None)
        if vc is not None:
            d["vary_colors"] = vc

        # Merge in per-type flat keys (snake_case at top level).
        d.update(self._chart_type_specific_keys())

        # Combination charts. Each secondary is fully serialised so
        # the Rust side does not need a half-shape; only per-family
        # fields (`kind`, `series_type`, `series`, type-specific keys,
        # `y_axis`) are consumed by the emitter. The outer-frame
        # fields on a secondary (anchor, dimensions, title, legend,
        # layout) are intentionally ignored downstream.
        secondary_dicts = [
            secondary.to_rust_dict()
            for secondary in self._charts[1:]
        ]
        if secondary_dicts:
            d["secondary_charts"] = secondary_dicts

        return d


_CHART_XML_MODEL_NAMES = (
    "AreaChart",
    "AreaChart3D",
    "BarChart",
    "BarChart3D",
    "BubbleChart",
    "LineChart",
    "LineChart3D",
    "DoughnutChart",
    "PieChart",
    "PieChart3D",
    "ProjectedPieChart",
    "RadarChart",
    "ScatterChart",
    "StockChart",
    "SurfaceChart",
    "SurfaceChart3D",
)


def _xml_model_names(cls: type) -> tuple[str, ...]:
    return tuple(getattr(cls, "__attrs__", ())) + tuple(getattr(cls, "__elements__", ()))


def _chart_to_openpyxl_model(value: Any) -> Any:
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, list):
        return [_chart_to_openpyxl_model(item) for item in value]
    if isinstance(value, tuple):
        return tuple(_chart_to_openpyxl_model(item) for item in value)

    cls = value.__class__
    module_name = str(getattr(cls, "__module__", ""))
    class_name = str(getattr(cls, "__name__", ""))

    if module_name.endswith(".series"):
        from .series import _to_openpyxl_model as _series_to_openpyxl_model

        return _series_to_openpyxl_model(value)
    if module_name.endswith(".label") or class_name in {"DataLabel", "DataLabelList"}:
        from .label import _to_openpyxl_model as _label_to_openpyxl_model

        return _label_to_openpyxl_model(value)
    if module_name.endswith(".updown_bars") or class_name in {"ChartLines", "UpDownBars"}:
        from .updown_bars import _to_openpyxl_model as _updown_to_openpyxl_model

        return _updown_to_openpyxl_model(value)

    upstream_cls = _resolve_openpyxl_class(cls.__module__, cls.__name__)
    if upstream_cls is None or cls is upstream_cls:
        return value

    names = _xml_model_names(upstream_cls)
    if not names:
        return value

    kwargs: dict[str, Any] = {}
    for name in names:
        if not hasattr(value, name):
            continue
        attr = getattr(value, name)
        if attr is None or attr == []:
            continue
        if name == "ser":
            for series in attr:
                series.__elements__ = attribute_mapping.get(value._series_type, ())
        if class_name == "ProjectedPieChart" and name == "serLines":
            continue
        if (
            class_name in {"AreaChart3D", "LineChart3D"}
            and name == "gapDepth"
            and attr == 150
        ):
            continue
        kwargs[name] = _chart_to_openpyxl_model(attr)
    return upstream_cls(**kwargs)


def _chart_to_native_model(value: Any) -> Any:
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, list):
        return [_chart_to_native_model(item) for item in value]
    if isinstance(value, tuple):
        return tuple(_chart_to_native_model(item) for item in value)

    module_name = str(getattr(value.__class__, "__module__", ""))
    class_name = str(getattr(value.__class__, "__name__", ""))
    if module_name.endswith(".series"):
        from .series import _to_native_series_model

        return _to_native_series_model(value)
    if module_name.endswith(".label") or class_name in {"DataLabel", "DataLabelList"}:
        from .label import _to_native_label_model

        return _to_native_label_model(value)
    if module_name.endswith(".updown_bars") or class_name in {"ChartLines", "UpDownBars"}:
        from .updown_bars import _to_native_updown_model

        return _to_native_updown_model(value)
    return value


def _from_openpyxl_chart_model(native_cls: type, value: Any) -> Any:
    names = _xml_model_names(value.__class__)
    kwargs: dict[str, Any] = {}
    for name in names:
        if not hasattr(value, name):
            continue
        attr = getattr(value, name)
        if attr is None:
            continue
        if name != "axId" and attr == []:
            continue
        kwargs[name] = _chart_to_native_model(attr)
    return native_cls(**kwargs)


def _chart_type_to_tree(
    self: Any,
    tagname: str | None = None,
    idx: int | None = None,
    namespace: str | None = None,
):
    upstream = _chart_to_openpyxl_model(self)
    return upstream.to_tree(tagname=tagname, idx=idx, namespace=namespace)


def _chart_type_from_tree(cls: type, node: Any) -> Any:
    upstream_cls = _resolve_openpyxl_class(cls.__module__, cls.__name__)
    if upstream_cls is None:
        raise AttributeError("from_tree")
    return _from_openpyxl_chart_model(cls, upstream_cls.from_tree(node))


def _chart_type_eq(self: Any, other: Any) -> bool:
    if type(self) is not type(other):
        return False
    return _chart_to_openpyxl_model(self) == _chart_to_openpyxl_model(other)


def _install_chart_type_xml_methods(*classes: type) -> None:
    for cls in classes:
        if cls.__name__ not in _CHART_XML_MODEL_NAMES:
            continue
        if "to_tree" not in cls.__dict__:
            cls.to_tree = _chart_type_to_tree  # type: ignore[attr-defined]
        if "from_tree" not in cls.__dict__:
            cls.from_tree = classmethod(_chart_type_from_tree)  # type: ignore[attr-defined]
        cls.__eq__ = _chart_type_eq  # type: ignore[attr-defined]


__all__ = ["ChartBase", "_install_chart_type_xml_methods"]

__getattr__ = _openpyxl_name_fallback(globals())
