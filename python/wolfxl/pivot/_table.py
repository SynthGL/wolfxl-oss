"""``PivotTable`` — pivot-table layout pointing at a :class:`PivotCache`.

Mirrors :class:`openpyxl.pivot.table.TableDefinition`.

:meth:`PivotTable._compute_layout` walks the cache's records and
renders the compact-form grid Excel would produce (see
:mod:`wolfxl.pivot._layout`): ``<rowItems>``/``<colItems>``, the
``<location>`` footprint, and the cell values written to the host
worksheet, so the table displays without a refresh.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import TYPE_CHECKING, Any, Iterable, Sequence

from wolfxl._compat import _install_openpyxl_iter

if TYPE_CHECKING:
    from ._cache import CacheValue, PivotCache


# ---------------------------------------------------------------------------
# Enums
# ---------------------------------------------------------------------------


class DataFunction:
    """Aggregator enum. Use the string constants verbatim."""

    SUM = "sum"
    COUNT = "count"
    AVERAGE = "average"
    MAX = "max"
    MIN = "min"
    PRODUCT = "product"
    COUNT_NUMS = "countNums"
    STD_DEV = "stdDev"
    STD_DEVP = "stdDevp"
    VAR = "var"
    VARP = "varp"

    ALL = (SUM, COUNT, AVERAGE, MAX, MIN, PRODUCT, COUNT_NUMS,
           STD_DEV, STD_DEVP, VAR, VARP)


# ---------------------------------------------------------------------------
# Field-axis builders
# ---------------------------------------------------------------------------


@dataclass
class RowField:
    """Explicit row-axis field. Name resolves to a cache field index
    at PivotTable construction time."""

    name: str


@dataclass
class ColumnField:
    """Explicit column-axis field."""

    name: str


@dataclass
class PageField:
    """Page-axis (filter) field. Default selection is "(All)" via
    ``item_index=-1``."""

    name: str
    item_index: int = -1
    hier: int = -1
    cap: str | None = None


@dataclass
class DataField:
    """A data-axis field — aggregator + display name + data field index.

    The ``name`` arg is the cache field name (e.g. ``"revenue"``) or
    a fully-spelled display name (``"Sum of revenue"``). When
    ``display_name`` is None we auto-format as ``"<Function> of
    <name>"`` per Excel convention.
    """

    name: str
    function: str = DataFunction.SUM
    display_name: str | None = None
    show_data_as: str | None = None
    base_field: int = 0
    base_item: int = 0
    num_fmt_id: int | None = None

    def __post_init__(self) -> None:
        if self.function not in DataFunction.ALL:
            raise ValueError(
                f"Unknown DataField.function={self.function!r}; "
                f"valid: {DataFunction.ALL}"
            )

    def resolved_display_name(self) -> str:
        if self.display_name is not None:
            return self.display_name
        # Excel-canonical name.
        verb = {
            DataFunction.SUM: "Sum",
            DataFunction.COUNT: "Count",
            DataFunction.AVERAGE: "Average",
            DataFunction.MAX: "Max",
            DataFunction.MIN: "Min",
            DataFunction.PRODUCT: "Product",
            DataFunction.COUNT_NUMS: "Count",
            DataFunction.STD_DEV: "StdDev",
            DataFunction.STD_DEVP: "StdDevp",
            DataFunction.VAR: "Var",
            DataFunction.VARP: "Varp",
        }.get(self.function, "Sum")
        return f"{verb} of {self.name}"


# ---------------------------------------------------------------------------
# Pivot field (per cache field appearance)
# ---------------------------------------------------------------------------


@dataclass
class PivotItem:
    """A single ``<item>`` child of a pivot field."""

    x: int | None = None
    t: str | None = None
    h: bool = False
    s: bool = False
    n: str | None = None

    def to_rust_dict(self) -> dict:
        return {"x": self.x, "t": self.t, "h": self.h, "s": self.s, "n": self.n}


@dataclass(init=False)
class PivotField:
    """One entry per cache field, in cache-field order.

    Constructed by :meth:`PivotTable._build_pivot_fields`; users don't
    normally instantiate this directly.
    """

    name: str | None = None
    axis: str | None = None
    data_field: bool = False
    show_all: bool = False
    default_subtotal: bool = True
    sum_subtotal: bool = False
    count_subtotal: bool = False
    avg_subtotal: bool = False
    max_subtotal: bool = False
    min_subtotal: bool = False
    items: list[PivotItem] | None = None
    outline: bool = True
    compact: bool = True
    subtotal_top: bool = True

    def __init__(
        self,
        name: str | None = None,
        axis: str | None = None,
        data_field: bool = False,
        show_all: bool = False,
        default_subtotal: bool = True,
        sum_subtotal: bool = False,
        count_subtotal: bool = False,
        avg_subtotal: bool = False,
        max_subtotal: bool = False,
        min_subtotal: bool = False,
        items: list[PivotItem] | None = None,
        outline: bool = True,
        compact: bool = True,
        subtotal_top: bool = True,
        **kw: Any,
    ) -> None:
        self.name = name
        self.axis = axis
        self.data_field = bool(kw.pop("dataField", data_field))
        self.show_all = bool(kw.pop("showAll", show_all))
        self.default_subtotal = bool(kw.pop("defaultSubtotal", default_subtotal))
        self.sum_subtotal = bool(kw.pop("sumSubtotal", sum_subtotal))
        self.count_subtotal = bool(kw.pop("countSubtotal", count_subtotal))
        self.avg_subtotal = bool(kw.pop("avgSubtotal", avg_subtotal))
        self.max_subtotal = bool(kw.pop("maxSubtotal", max_subtotal))
        self.min_subtotal = bool(kw.pop("minSubtotal", min_subtotal))
        self.items = items
        self.outline = bool(outline)
        self.compact = bool(compact)
        self.subtotal_top = bool(kw.pop("subtotalTop", subtotal_top))
        for extra_name, value in kw.items():
            setattr(self, extra_name, value)

    def to_rust_dict(self) -> dict:
        return {
            "name": self.name,
            "axis": self.axis,
            "data_field": self.data_field,
            "show_all": self.show_all,
            "default_subtotal": self.default_subtotal,
            "sum_subtotal": self.sum_subtotal,
            "count_subtotal": self.count_subtotal,
            "avg_subtotal": self.avg_subtotal,
            "max_subtotal": self.max_subtotal,
            "min_subtotal": self.min_subtotal,
            "items": [it.to_rust_dict() for it in self.items] if self.items else None,
            "outline": self.outline,
            "compact": self.compact,
            "subtotal_top": self.subtotal_top,
        }


# ---------------------------------------------------------------------------
# Location + style info
# ---------------------------------------------------------------------------


@dataclass
class Location:
    """Pivot-table anchor location."""

    ref: str
    first_header_row: int = 0
    first_data_row: int = 1
    first_data_col: int = 1
    row_page_count: int | None = None
    col_page_count: int | None = None

    def to_rust_dict(self) -> dict:
        return {
            "ref": self.ref,
            "first_header_row": self.first_header_row,
            "first_data_row": self.first_data_row,
            "first_data_col": self.first_data_col,
            "row_page_count": self.row_page_count,
            "col_page_count": self.col_page_count,
        }


@dataclass
class PivotTableStyleInfo:
    """Pivot-table style info (table style + header/stripe toggles)."""

    name: str = "PivotStyleLight16"
    show_row_headers: bool = True
    show_col_headers: bool = True
    show_row_stripes: bool = False
    show_col_stripes: bool = False
    show_last_column: bool = True

    def to_rust_dict(self) -> dict:
        return {
            "name": self.name,
            "show_row_headers": self.show_row_headers,
            "show_col_headers": self.show_col_headers,
            "show_row_stripes": self.show_row_stripes,
            "show_col_stripes": self.show_col_stripes,
            "show_last_column": self.show_last_column,
        }


# ---------------------------------------------------------------------------
# PivotSource — chart-side linkage marker
# ---------------------------------------------------------------------------


@dataclass
class PivotSource:
    """The ``<c:pivotSource>`` block on a pivot chart.

    Set on a chart via ``chart.pivot_source = pt`` (where ``pt`` is a
    :class:`PivotTable`) or ``chart.pivot_source = ("name", fmt_id)``.
    """

    name: str
    fmt_id: int = 0

    def __post_init__(self) -> None:
        if self.fmt_id < 0 or self.fmt_id > 65535:
            raise ValueError(
                f"PivotSource.fmt_id={self.fmt_id} out of [0, 65535]"
            )

    def to_rust_dict(self) -> dict:
        return {"name": self.name, "fmt_id": self.fmt_id}


# ---------------------------------------------------------------------------
# PivotTable (top level)
# ---------------------------------------------------------------------------


# Internal: a tuple of cache-field shared-items indices used as a
# row/col axis-key. Reserved for type-hints clarity.
_AxisKey = tuple


class PivotTable:
    """Pivot-table layout pointing at a :class:`PivotCache`.

    The bare-string convenience form resolves a string field name
    (``rows=["region"]``) to a :class:`RowField`. Order of fields in
    ``rows``/``cols`` controls nesting (first → outermost).
    """

    def __init__(
        self,
        *,
        cache: "PivotCache",
        location: str | tuple[str, str],
        rows: Sequence[str | RowField] | None = None,
        cols: Sequence[str | ColumnField] | None = None,
        data: Sequence[str | DataField] | None = None,
        page: Sequence[str | PageField] | None = None,
        name: str = "PivotTable1",
        style_name: str = "PivotStyleLight16",
        row_grand_totals: bool = True,
        col_grand_totals: bool = True,
        outline: bool = True,
        compact: bool = True,
        data_caption: str = "Values",
        grand_total_caption: str | None = None,
        error_caption: str | None = None,
        missing_caption: str | None = None,
    ) -> None:
        from ._cache import PivotCache  # avoid cycle at import time

        if not isinstance(cache, PivotCache):
            raise TypeError(
                f"PivotTable.cache must be a PivotCache, got {type(cache).__name__}"
            )

        self.cache = cache
        self.name = name
        self.style_info = PivotTableStyleInfo(name=style_name)
        self.row_grand_totals = row_grand_totals
        self.col_grand_totals = col_grand_totals
        self.outline = outline
        self.compact = compact
        self.data_caption = data_caption
        self.grand_total_caption = grand_total_caption
        self.error_caption = error_caption
        self.missing_caption = missing_caption

        # Defer field validation/resolution until cache is materialized.
        self._row_field_specs = self._normalize_row_specs(rows or [])
        self._col_field_specs = self._normalize_col_specs(cols or [])
        self._page_field_specs = self._normalize_page_specs(page or [])
        self._data_field_specs = self._normalize_data_specs(data or [])

        if not self._data_field_specs:
            raise ValueError("PivotTable requires ≥1 data field")

        # Location is normalized eagerly; defaults computed when we
        # know the row/col counts (in _compute_layout).
        self.location = self._normalize_location(location)

        # These are populated by _compute_layout(), called by
        # Worksheet.add_pivot_table() (or eagerly by .to_rust_dict()).
        self._pivot_fields: list[PivotField] | None = None
        self._row_field_indices: list[int] | None = None
        self._col_field_indices: list[int] | None = None
        self._page_field_dicts: list[dict] | None = None
        self._data_field_dicts: list[dict] | None = None
        self._row_items: list[dict] | None = None
        self._col_items: list[dict] | None = None
        # Rendered cells as ((row, col) 1-based absolute) -> value, and the
        # (first row, first col, last row, last col) rectangle they cover,
        # page filters included.
        self._rendered_cells: dict[tuple[int, int], Any] = {}
        self._footprint: tuple[int, int, int, int] | None = None
        # Calculated items (table-scoped).
        self.calculated_items: list[Any] = []
        # Pivot-area Format / CF directives (table-scoped).
        self.formats: list[Any] = []
        self.conditional_formats: list[Any] = []
        self.chart_formats: list[Any] = []

    # ------------------------------------------------------------------
    # Public attribute proxies expected by openpyxl-shaped users.
    # ------------------------------------------------------------------

    @property
    def rows(self) -> list[RowField]:
        """Return row-field specifications for this pivot table.

        Returns:
            The normalized row field list, preserving the caller's field order.
        """
        return self._row_field_specs

    @property
    def cols(self) -> list[ColumnField]:
        """Return column-field specifications for this pivot table.

        Returns:
            The normalized column field list.
        """
        return self._col_field_specs

    @property
    def page(self) -> list[PageField]:
        """Return report-filter page fields for this pivot table.

        Returns:
            The normalized page field list.
        """
        return self._page_field_specs

    @property
    def data(self) -> list[DataField]:
        """Return data-value fields for this pivot table.

        Returns:
            The normalized data field list.
        """
        return self._data_field_specs

    # ------------------------------------------------------------------
    # Spec normalization
    # ------------------------------------------------------------------

    @staticmethod
    def _normalize_row_specs(items: Iterable) -> list[RowField]:
        out: list[RowField] = []
        for it in items:
            if isinstance(it, RowField):
                out.append(it)
            elif isinstance(it, str):
                out.append(RowField(name=it))
            else:
                raise TypeError(
                    f"rows[*] must be str or RowField, got {type(it).__name__}"
                )
        return out

    @staticmethod
    def _normalize_col_specs(items: Iterable) -> list[ColumnField]:
        out: list[ColumnField] = []
        for it in items:
            if isinstance(it, ColumnField):
                out.append(it)
            elif isinstance(it, str):
                out.append(ColumnField(name=it))
            else:
                raise TypeError(
                    f"cols[*] must be str or ColumnField, got {type(it).__name__}"
                )
        return out

    @staticmethod
    def _normalize_page_specs(items: Iterable) -> list[PageField]:
        out: list[PageField] = []
        for it in items:
            if isinstance(it, PageField):
                out.append(it)
            elif isinstance(it, str):
                out.append(PageField(name=it))
            else:
                raise TypeError(
                    f"page[*] must be str or PageField, got {type(it).__name__}"
                )
        return out

    @staticmethod
    def _normalize_data_specs(items: Iterable) -> list[DataField]:
        out: list[DataField] = []
        for it in items:
            if isinstance(it, DataField):
                out.append(it)
            elif isinstance(it, str):
                out.append(DataField(name=it))
            elif isinstance(it, tuple) and len(it) == 2 and all(
                isinstance(x, str) for x in it
            ):
                fname, fn = it
                out.append(DataField(name=fname, function=fn))
            else:
                raise TypeError(
                    f"data[*] must be str | DataField | (str, str), "
                    f"got {type(it).__name__}"
                )
        return out

    @staticmethod
    def _normalize_location(loc: str | tuple[str, str] | Location) -> Location:
        if isinstance(loc, Location):
            return loc
        if isinstance(loc, str):
            return Location(ref=loc)  # caller passes "F2"; we widen during layout
        if isinstance(loc, tuple) and len(loc) == 2:
            return Location(ref=f"{loc[0]}:{loc[1]}")
        raise TypeError(
            "PivotTable.location must be str (e.g. 'F2'), "
            "tuple ('F2', 'I20'), or Location"
        )

    # ------------------------------------------------------------------
    # Layout pre-computation
    # ------------------------------------------------------------------

    def _compute_layout(self) -> None:
        """Resolve fields, render the compact layout, and size the location.

        Must be called after :attr:`cache` has been materialized via
        :meth:`Workbook.add_pivot_cache`. Worksheet.add_pivot_table()
        invokes this before serializing.
        """
        from ._layout import (
            ALL_ITEMS_CAPTION,
            DataSpec,
            compute_layout,
            item_label,
            renders_show_data_as,
        )

        # 1. Resolve cache-field indices for each axis.
        cache_fields = self.cache.fields
        cache_field_count = len(cache_fields)
        row_indices = [self.cache.field_index(rf.name) for rf in self._row_field_specs]
        col_indices = [self.cache.field_index(cf.name) for cf in self._col_field_specs]
        page_indices = [
            self.cache.field_index(pf.name) for pf in self._page_field_specs
        ]
        data_field_indices = [
            self.cache.field_index(df.name) for df in self._data_field_specs
        ]

        # 2. Validate no field on multiple axes.
        all_axis_indices = (
            list(row_indices) + list(col_indices) + list(page_indices)
        )
        if len(all_axis_indices) != len(set(all_axis_indices)):
            raise ValueError(
                "PivotTable: a cache field appears on multiple axes"
            )
        for i in all_axis_indices:
            if cache_fields[i].shared_items.items is None:
                raise ValueError(
                    f"PivotTable: field {cache_fields[i].name!r} has too many "
                    "unique values to enumerate as shared items, so it cannot "
                    "be placed on the row, column, or page axis; use it as a "
                    "data field instead"
                )

        # 3. Build pivot_fields (one per cache field).
        pivot_fields: list[PivotField] = []
        for i in range(cache_field_count):
            axis: str | None = None
            if i in row_indices:
                axis = "axisRow"
            elif i in col_indices:
                axis = "axisCol"
            elif i in page_indices:
                axis = "axisPage"

            data_field = i in data_field_indices
            pivot_fields.append(
                PivotField(
                    axis=axis,
                    data_field=data_field,
                    show_all=False,
                )
            )

        # 4. Render the compact layout from the cache records, keeping only
        #    records that match each page field's selected item.
        lookups: dict[int, dict[tuple, int]] = {
            i: {
                (v.kind, v.value): n
                for n, v in enumerate(cache_fields[i].shared_items.items)
            }
            for i in set(row_indices + col_indices + page_indices)
        }

        def item_index(cell: "CacheValue", field_index: int) -> int:
            return lookups[field_index][(cell.kind, cell.value)]

        page_filters = [
            (field_index, pf.item_index)
            for field_index, pf in zip(page_indices, self._page_field_specs)
            if self._selected_page_item(pf, cache_fields[field_index]) is not None
        ]
        records = [
            record
            for record in self.cache._records or []  # noqa: SLF001
            if all(
                item_index(record[f], f) == selected for f, selected in page_filters
            )
        ]

        layout = compute_layout(
            records=records,
            item_index=item_index,
            shared_items=[f.shared_items.items for f in cache_fields],
            row_fields=row_indices,
            col_fields=col_indices,
            data=[
                DataSpec(
                    field_index=data_field_indices[i],
                    function=df.function,
                    caption=df.resolved_display_name(),
                    show_data_as=df.show_data_as,
                )
                for i, df in enumerate(self._data_field_specs)
            ],
            row_grand_totals=self.row_grand_totals,
            col_grand_totals=self.col_grand_totals,
            grand_total_caption=self.grand_total_caption,
            data_caption=self.data_caption,
        )
        # Excel computes base-relative "show values as" options (difference,
        # running total, ...) itself; refresh the cache on open so they
        # appear without a manual refresh.
        if any(
            not renders_show_data_as(df.show_data_as) for df in self._data_field_specs
        ):
            self.cache.refresh_on_load = True

        # 5. Build data-field dicts.
        data_field_dicts = [
            {
                "name": df.resolved_display_name(),
                "field_index": data_field_indices[i],
                "function": df.function,
                "show_data_as": df.show_data_as,
                "base_field": df.base_field,
                "base_item": df.base_item,
                "num_fmt_id": df.num_fmt_id,
            }
            for i, df in enumerate(self._data_field_specs)
        ]

        # 6. Build page-field dicts.
        page_field_dicts = [
            {
                "field_index": page_indices[i],
                "name": pf.name if pf.name else None,
                "item_index": pf.item_index,
                "hier": pf.hier,
                "cap": pf.cap,
            }
            for i, pf in enumerate(self._page_field_specs)
        ]

        # 7. Place the table. The location anchor is the table's top-left
        #    cell; page filters sit above it with one spacer row, and the
        #    table moves down when the anchor leaves no room for them.
        top, left = self._anchor_cell()
        n_pages = len(page_indices)
        if n_pages:
            top = max(top, n_pages + 2)
        from wolfxl.chart.reference import _index_to_col

        self.location.ref = (
            f"{_index_to_col(left)}{top}:"
            f"{_index_to_col(left + layout.width - 1)}{top + layout.height - 1}"
        )
        # Compact layouts use firstHeaderRow="1" regardless of axis shape.
        self.location.first_header_row = 1
        self.location.first_data_row = layout.first_data_row
        self.location.first_data_col = layout.first_data_col
        self.location.row_page_count = n_pages or None
        self.location.col_page_count = 1 if n_pages else None

        rendered = {
            (top + r, left + c): v for (r, c), v in layout.cells.items()
        }
        for i, (field_index, pf) in enumerate(
            zip(page_indices, self._page_field_specs)
        ):
            row = top - n_pages - 1 + i
            selected = self._selected_page_item(pf, cache_fields[field_index])
            rendered[(row, left)] = cache_fields[field_index].name
            rendered[(row, left + 1)] = (
                ALL_ITEMS_CAPTION
                if selected is None
                else item_label(cache_fields[field_index].shared_items.items[selected])
            )

        # 8. Persist.
        self._pivot_fields = pivot_fields
        self._row_field_indices = row_indices
        self._col_field_indices = col_indices
        self._page_field_dicts = page_field_dicts
        self._data_field_dicts = data_field_dicts
        self._row_items = layout.row_items
        self._col_items = layout.col_items
        self._rendered_cells = rendered
        self._footprint = (
            top - n_pages - 1 if n_pages else top,
            left,
            top + layout.height - 1,
            left + max(layout.width, 2 if n_pages else 1) - 1,
        )

    def _anchor_cell(self) -> tuple[int, int]:
        """1-based (row, column) of the location's top-left cell."""
        from wolfxl.chart.reference import _col_to_index

        anchor = self.location.ref.split(":", 1)[0].replace("$", "").upper()
        m = re.match(r"^([A-Z]+)(\d+)$", anchor)
        if not m:
            raise ValueError(
                f"Cannot parse PivotTable.location ref={self.location.ref!r}"
            )
        return int(m.group(2)), _col_to_index(m.group(1))

    @staticmethod
    def _selected_page_item(pf: Any, cache_field: Any) -> int | None:
        """Shared-item index a page field filters on, or None for "(All)"."""
        items = cache_field.shared_items.items
        if pf.item_index is not None and 0 <= pf.item_index < len(items):
            return pf.item_index
        return None

    def _write_rendered_cells(self, ws: Any) -> None:
        """Write the rendered pivot into the host worksheet.

        Every cell in the footprint is assigned, so values already in the
        range cannot pass for pivot results.
        """
        if self._footprint is None:
            return
        first_row, first_col, last_row, last_col = self._footprint
        for row in range(first_row, last_row + 1):
            for col in range(first_col, last_col + 1):
                value = self._rendered_cells.get((row, col))
                cell = ws.cell(row=row, column=col)
                if value is not None or cell.value is not None:
                    cell.value = value

    # ------------------------------------------------------------------
    # to_rust_dict
    # ------------------------------------------------------------------

    # ------------------------------------------------------------------
    # Calculated items
    # ------------------------------------------------------------------

    def add_calculated_item(
        self,
        field: str,
        item_name: str,
        formula: str,
    ):
        """Add a calculated item to a specific field of this pivot table.

        Calc items live inside pivot table XML, NOT cache XML.
        Excel evaluates the formula on open — wolfxl does not
        pre-evaluate.

        Returns the registered :class:`CalculatedItem`.
        """
        from ._calc import CalculatedItem

        ci = CalculatedItem(
            field_name=field, item_name=item_name, formula=formula
        )
        self.calculated_items.append(ci)
        return ci

    # ------------------------------------------------------------------
    # Pivot-area Format / CF directives
    # ------------------------------------------------------------------

    def add_format(
        self,
        pivot_area,
        dxf=None,
        action: str = "formatting",
        dxf_id: int | None = None,
    ):
        """Add a Format directive (pivot-area + dxf + action).

        Either pass an explicit ``dxf_id`` (resolved to an existing
        workbook dxf table entry) or pass a ``dxf=`` instance for
        the patcher to allocate one via the workbook dxf allocator.

        Returns the registered :class:`Format`.
        """
        from ._styling import Format, PivotArea

        if not isinstance(pivot_area, PivotArea):
            raise TypeError(
                "PivotTable.add_format: pivot_area must be a PivotArea"
            )
        if dxf_id is None:
            # The patcher will assign a workbook-scoped dxf id from
            # the dxf= payload at flush time. Until then, we tag a
            # sentinel `dxf_id=-1` and stash the dxf payload.
            dxf_id = -1
        f = Format(pivot_area=pivot_area, dxf_id=dxf_id, action=action)
        f._dxf_payload = dxf  # type: ignore[attr-defined]
        self.formats.append(f)
        return f

    def add_conditional_format(
        self,
        rule,
        pivot_area,
        priority: int = 1,
    ):
        """Add a pivot-scoped CF rule.

        ``rule`` is a CF rule object (e.g.
        :class:`wolfxl.formatting.Rule` /
        :class:`wolfxl.formatting.ColorScale` / etc.). The rule
        references a workbook-scoped dxf entry through the workbook
        ``dxfId`` allocator.

        Returns the registered :class:`PivotConditionalFormat`.
        """
        from ._styling import PivotArea, PivotConditionalFormat

        if isinstance(pivot_area, PivotArea):
            areas = [pivot_area]
        else:
            areas = list(pivot_area)
            if not all(isinstance(a, PivotArea) for a in areas):
                raise TypeError(
                    "PivotTable.add_conditional_format: pivot_area must "
                    "be a PivotArea (or list of PivotArea)"
                )
        pcf = PivotConditionalFormat(
            rule=rule, pivot_areas=areas, priority=priority
        )
        self.conditional_formats.append(pcf)
        return pcf

    def to_rust_dict(self) -> dict:
        """Pivot-table dict for the Rust emitter.

        Calls :meth:`_compute_layout` if not already invoked.
        """
        if self._pivot_fields is None:
            self._compute_layout()

        return {
            "name": self.name,
            "cache_id": self.cache.cache_id,
            "location": self.location.to_rust_dict(),
            "pivot_fields": [pf.to_rust_dict() for pf in self._pivot_fields],
            "row_field_indices": list(self._row_field_indices or []),
            "col_field_indices": list(self._col_field_indices or []),
            "page_fields": list(self._page_field_dicts or []),
            "data_fields": list(self._data_field_dicts or []),
            "row_items": list(self._row_items or []),
            "col_items": list(self._col_items or []),
            "data_on_rows": False,
            "outline": self.outline,
            "compact": self.compact,
            "row_grand_totals": self.row_grand_totals,
            "col_grand_totals": self.col_grand_totals,
            "data_caption": self.data_caption,
            "grand_total_caption": self.grand_total_caption,
            "error_caption": self.error_caption,
            "missing_caption": self.missing_caption,
            "apply_number_formats": False,
            "apply_border_formats": False,
            "apply_font_formats": False,
            "apply_pattern_formats": False,
            "apply_alignment_formats": False,
            "apply_width_height_formats": True,
            "style_info": self.style_info.to_rust_dict() if self.style_info else None,
            "created_version": 6,
            "updated_version": 6,
            "min_refreshable_version": 3,
            # Extensions: calc items + pivot-scoped formats / CF.
            "calculated_items": [
                ci.to_rust_dict() for ci in self.calculated_items
            ],
            "formats": [f.to_rust_dict() for f in self.formats],
            "conditional_formats": [
                cf.to_rust_dict() for cf in self.conditional_formats
            ],
            "chart_formats": [
                cf.to_rust_dict() for cf in self.chart_formats
            ],
        }


_install_openpyxl_iter(PivotField, PivotTable)
