"""Worksheet collection proxy objects."""

from __future__ import annotations

import inspect as _inspect
from typing import TYPE_CHECKING, Any
from xml.etree import ElementTree as ET

from wolfxl._worksheet_merge_metadata import _merged_border_for_cell as _merged_border_for_cell
from wolfxl.worksheet.cell_range import CellRange, MultiCellRange as _BaseMultiCellRange

if TYPE_CHECKING:
    from wolfxl._worksheet import Worksheet


def _loaded_merged_range_refs(ws: Worksheet) -> set[str]:
    """Load native merge metadata once, retaining pending additions/removals."""
    if not ws._merged_ranges_loaded:  # noqa: SLF001
        reader = getattr(ws._workbook, "_rust_reader", None)  # noqa: SLF001
        if reader is not None:
            try:
                read_ranges = getattr(
                    reader, "read_merged_ranges_if_present", reader.read_merged_ranges
                )
                ws._merged_ranges.update(str(ref) for ref in read_ranges(ws._title))  # noqa: SLF001
            except Exception:
                pass
        ws._merged_ranges_loaded = True  # noqa: SLF001
    return (ws._merged_ranges | ws._collection_merged_ranges) - ws._pending_unmerged_ranges  # noqa: SLF001


def _invalidate_streaming_merge_cache(ws: Worksheet) -> None:
    wb = ws._workbook  # noqa: SLF001
    mutated = getattr(wb, "_merge_mutated_sheets", None)
    if mutated is None:
        mutated = set()
        wb._merge_mutated_sheets = mutated  # noqa: SLF001
    mutated.add(ws.title)
    if hasattr(ws._workbook, "_streaming_style_cache"):  # noqa: SLF001
        ws._workbook._streaming_style_cache = None  # noqa: SLF001


def _invalidate_clean_merge_anchor_border(ws: Worksheet, range_string: str) -> None:
    """Recompose a source border that was cached before this merge existed."""
    from wolfxl._cell import _UNSET
    from wolfxl.utils.cell import range_boundaries

    col, row, _, _ = range_boundaries(range_string)
    anchor = ws._cells.get((row, col))  # noqa: SLF001
    if anchor is not None and not getattr(anchor, "_border_authored", False):
        anchor._border = _UNSET  # noqa: SLF001


class AutoFilter:
    """Proxy for ``ws.auto_filter``.

    The object mirrors openpyxl's worksheet filter surface and keeps filter
    column plus sort-state serialization in one place.
    """

    __slots__ = ("_ref", "filter_columns", "sort_state")

    def __init__(self) -> None:
        from wolfxl.worksheet.filters import FilterColumn  # noqa: F401

        self._ref: str | None = None
        self.filter_columns: list[Any] = []
        self.sort_state: Any = None

    @property
    def ref(self) -> str | None:
        return self._ref

    @ref.setter
    def ref(self, value: str | None) -> None:
        self._ref = value.upper() if isinstance(value, str) else value

    @property
    def filterColumn(self) -> list[Any]:  # noqa: N802
        return self.filter_columns

    @filterColumn.setter
    def filterColumn(self, value: list[Any]) -> None:  # noqa: N802
        self.filter_columns = value

    @property
    def sortState(self) -> Any:  # noqa: N802
        return self.sort_state

    @sortState.setter
    def sortState(self, value: Any) -> None:  # noqa: N802
        self.sort_state = value

    def __iter__(self):  # type: ignore[no-untyped-def]
        if self._ref is not None:
            yield "ref", str(self._ref)

    def to_tree(
        self,
        tagname: str | None = None,
        idx: int | None = None,  # noqa: ARG002 - openpyxl signature
        namespace: str | None = None,  # noqa: ARG002 - openpyxl signature
    ) -> ET.Element:
        """Serialize this proxy using the same shape as the dataclass model."""
        from wolfxl.worksheet.filters import AutoFilter as _AF

        return _AF(
            ref=self._ref,
            filter_columns=list(self.filter_columns),
            sort_state=self.sort_state,
        ).to_tree(tagname=tagname, idx=idx, namespace=namespace)

    def add_filter_column(
        self,
        col_id: int,
        vals: Any = None,
        blank: bool = False,
        *,
        hidden_button: bool = False,
        show_button: bool = True,
        date_group_items: Any = None,
    ) -> None:
        """Append a ``FilterColumn`` entry."""
        from wolfxl.worksheet.filters import FilterColumn, Filters

        filter = (
            Filters(blank=blank, filter=[str(value) for value in vals])
            if isinstance(vals, (list, tuple))
            else vals
        )
        fc = FilterColumn(
            col_id=col_id,
            hidden_button=hidden_button,
            show_button=show_button,
            filter=filter,
            date_group_items=list(date_group_items) if date_group_items else [],
        )
        self.filter_columns.append(fc)
        return None

    def add_sort_condition(
        self,
        ref: str,
        descending: bool = False,
        sort_by: str | None = None,
        *,
        custom_list: str | None = None,
        dxf_id: int | None = None,
        icon_set: str | None = None,
        icon_id: int | None = None,
    ) -> None:
        """Append a ``SortCondition`` entry."""
        from wolfxl.worksheet.filters import SortCondition, SortState

        if self.sort_state is None:
            self.sort_state = SortState(ref=self._ref)
        sc = SortCondition(
            ref=ref,
            descending=descending,
            sort_by=sort_by,
            custom_list=custom_list,
            dxf_id=dxf_id,
            icon_set=icon_set,
            icon_id=icon_id,
        )
        self.sort_state.sort_conditions.append(sc)
        return None

    def to_rust_dict(self) -> dict[str, Any]:
        """Serialize the filter state for the native save pipeline."""
        from wolfxl.worksheet.filters import AutoFilter as _AF

        af = _AF(
            ref=self._ref,
            filter_columns=list(self.filter_columns),
            sort_state=self.sort_state,
        )
        return af.to_rust_dict()


class MultiCellRange(_BaseMultiCellRange):
    """Openpyxl-shaped proxy for ``Worksheet.merged_cells``."""

    __slots__ = ("_ws",)

    def __init__(self, ws: Worksheet) -> None:
        self._ws = ws

    @property
    def ranges(self) -> set[Any]:
        ws = self._ws
        refs = _loaded_merged_range_refs(ws)
        from wolfxl.worksheet.merge import MergedCellRange

        return {MergedCellRange(ws, str(ref)) for ref in refs}

    def __iter__(self):  # type: ignore[no-untyped-def]
        return iter(self.ranges)

    def __len__(self) -> int:
        return len(self.ranges)

    def __contains__(self, item: object) -> bool:
        if isinstance(item, str):
            try:
                target = CellRange(item)
            except Exception:
                return False
        elif isinstance(item, CellRange):
            target = item
        else:
            return False
        return any(target.issubset(CellRange(str(rng))) for rng in self.ranges)

    def add(self, coord: Any) -> None:
        range_string = _coerce_range_string(coord)
        _loaded_merged_range_refs(self._ws)
        _invalidate_streaming_merge_cache(self._ws)
        self._ws._pending_unmerged_ranges.discard(range_string)  # noqa: SLF001
        self._ws._plain_cell_fast_path = False  # noqa: SLF001
        self._ws._collection_merged_ranges.add(range_string)  # noqa: SLF001
        _invalidate_clean_merge_anchor_border(self._ws, range_string)
        self._ws._merged_ranges_loaded = True  # noqa: SLF001
        self._ws._pending_collection_merged_ranges.add(range_string)  # noqa: SLF001

    def remove(self, coord: Any) -> None:
        range_string = _coerce_range_string(coord)
        if range_string not in {str(rng) for rng in self.ranges}:
            raise KeyError(CellRange(range_string))
        _invalidate_streaming_merge_cache(self._ws)
        self._ws._pending_unmerged_ranges.add(range_string)  # noqa: SLF001
        self._ws._merged_ranges.discard(range_string)  # noqa: SLF001
        self._ws._collection_merged_ranges.discard(range_string)  # noqa: SLF001
        self._ws._merged_ranges_loaded = True  # noqa: SLF001
        self._ws._pending_collection_merged_ranges.discard(range_string)  # noqa: SLF001

    def __repr__(self) -> str:
        return f"<MultiCellRange {[str(rng) for rng in self.ranges]}>"

    def __eq__(self, other: object) -> bool:
        if isinstance(other, str):
            return {str(rng) for rng in self.ranges} == set(other.split())
        return super().__eq__(other)

    def __str__(self) -> str:
        return " ".join(sorted(str(rng) for rng in self.ranges))


def _coerce_range_string(coord: Any) -> str:
    if isinstance(coord, CellRange):
        return coord.coord
    return str(coord)


MergedCellsProxy = MultiCellRange


AutoFilter.add_filter_column.__signature__ = _inspect.Signature(
    [
        _inspect.Parameter("self", _inspect.Parameter.POSITIONAL_OR_KEYWORD),
        _inspect.Parameter("col_id", _inspect.Parameter.POSITIONAL_OR_KEYWORD),
        _inspect.Parameter("vals", _inspect.Parameter.POSITIONAL_OR_KEYWORD),
        _inspect.Parameter(
            "blank",
            _inspect.Parameter.POSITIONAL_OR_KEYWORD,
            default=False,
        ),
    ]
)
AutoFilter.add_sort_condition.__signature__ = _inspect.Signature(
    [
        _inspect.Parameter("self", _inspect.Parameter.POSITIONAL_OR_KEYWORD),
        _inspect.Parameter("ref", _inspect.Parameter.POSITIONAL_OR_KEYWORD),
        _inspect.Parameter(
            "descending",
            _inspect.Parameter.POSITIONAL_OR_KEYWORD,
            default=False,
        ),
    ]
)


def merge_cells(ws: Worksheet, range_string: str) -> None:
    """Merge a cell range through the write-mode Rust backend."""
    _loaded_merged_range_refs(ws)
    _invalidate_streaming_merge_cache(ws)
    ws._pending_unmerged_ranges.discard(range_string)  # noqa: SLF001
    ws._plain_cell_fast_path = False  # noqa: SLF001
    wb = ws._workbook  # noqa: SLF001
    rust_writer = getattr(wb, "_rust_writer", None)
    if rust_writer is not None:
        try:
            rust_writer.merge_cells(ws._title, range_string)
        except ValueError as exc:
            if "Unknown sheet" not in str(exc):
                raise
    ws._merged_ranges.add(range_string)  # noqa: SLF001
    _invalidate_clean_merge_anchor_border(ws, range_string)
    ws._merged_ranges_loaded = True  # noqa: SLF001
    _discard_merged_subordinate_cells(ws, range_string)


def unmerge_cells(ws: Worksheet, range_string: str) -> None:
    """Forget a merged range from the worksheet's pending merge set."""
    _loaded_merged_range_refs(ws)
    _invalidate_streaming_merge_cache(ws)
    ws._merged_ranges.discard(range_string)  # noqa: SLF001
    ws._collection_merged_ranges.discard(range_string)  # noqa: SLF001
    ws._pending_collection_merged_ranges.discard(range_string)  # noqa: SLF001
    ws._merged_ranges_loaded = True  # noqa: SLF001
    ws._pending_unmerged_ranges.add(range_string)  # noqa: SLF001


def _discard_merged_subordinate_cells(ws: Worksheet, range_string: str) -> None:
    from wolfxl.utils.cell import range_boundaries

    min_col, min_row, max_col, max_row = range_boundaries(range_string)
    if None in (min_col, min_row, max_col, max_row):
        return
    for row in range(int(min_row), int(max_row) + 1):
        for col in range(int(min_col), int(max_col) + 1):
            if row == int(min_row) and col == int(min_col):
                continue
            ws._cells.pop((row, col), None)  # noqa: SLF001
            ws._dirty.discard((row, col))  # noqa: SLF001
