"""Workbook-scoped style components for the XML streaming reader.

Cell style IDs refer to the immutable source stylesheet. Reusing their
converted components keeps styled streaming bounded by distinct styles rather
than cells and never asks the eager reader to materialize worksheet cells.
"""

from __future__ import annotations

from copy import deepcopy
from typing import Any, NamedTuple

from wolfxl._cell_payloads import (
    border_payload_to_border,
    format_to_alignment,
    format_to_fill,
    format_to_font,
)
from wolfxl._styles import Alignment, Border, Font, PatternFill, Side
from wolfxl.styles.fills import GradientFill
from wolfxl.styles.numbers import BUILTIN_FORMATS, BUILTIN_FORMATS_MAX_SIZE, is_timedelta_format
from wolfxl.utils.cell import get_column_letter, range_boundaries
from wolfxl.utils.numbers import is_date_format


class StyleComponents(NamedTuple):
    font: Font
    fill: PatternFill | GradientFill
    border: Border
    alignment: Alignment
    number_format: str
    is_date: bool
    is_timedelta: bool


_DEFAULT_STYLE = StyleComponents(
    format_to_font({}), format_to_fill({}), Border(), Alignment(), "General", False, False
)


class StreamingStyleCache:
    """Cache source styles and merge metadata without retaining a native reader."""

    __slots__ = (
        "reader_identity",
        "styles",
        "merged_bounds",
        "merged_state",
        "source_borders",
        "source_gradients",
        "date_ids",
    )

    def __init__(self, reader: Any) -> None:
        self.reader_identity = id(reader)
        self.styles: dict[int, StyleComponents] = {}
        self.merged_bounds: dict[str, tuple[tuple[int, int, int, int], ...]] = {}
        self.merged_state: dict[str, tuple[set[str], ...]] = {}
        self.source_borders: dict[int, Border] = {}
        self.source_gradients: dict[int, GradientFill] = {}
        self.date_ids: tuple[Any, frozenset[int]] | None = None

    def date_style_ids(self, workbook: Any) -> tuple[Any, frozenset[int]]:
        if self.date_ids is not None:
            return self.date_ids
        date_styles = getattr(workbook, "_date_formats", ())
        styles = getattr(workbook, "_cell_styles", ())
        number_formats = getattr(workbook, "_number_formats", ())
        timedeltas = set()
        # The stylesheet's timedelta set historically also includes locale
        # date formats. Filter its candidates with the strict public classifier
        # so [$-409]m/d/yy remains a datetime rather than a timedelta.
        for style_id in getattr(workbook, "_timedelta_formats", ()):
            if style_id not in date_styles or not 0 <= style_id < len(styles):
                continue
            number_id = styles[style_id].numFmtId
            fmt = BUILTIN_FORMATS.get(number_id)
            custom_id = number_id - BUILTIN_FORMATS_MAX_SIZE
            if fmt is None and 0 <= custom_id < len(number_formats):
                fmt = number_formats[custom_id]
            if is_timedelta_format(fmt):
                timedeltas.add(style_id)
        self.date_ids = date_styles, frozenset(timedeltas)
        return self.date_ids

    def components(self, workbook: Any, style_id: int | None) -> StyleComponents:
        style_id = int(style_id or 0)
        cached = self.styles.get(style_id)
        if cached is not None:
            return cached
        reader = workbook._rust_reader
        if reader is None:
            return _DEFAULT_STYLE
        payload = reader.read_format_for_style_id(style_id) if style_id else {}
        if not isinstance(payload, dict):
            payload = {}
        number_format = payload.get("number_format") or "General"
        # The Python stylesheet is already hydrated at workbook open. Borders
        # are absent from native format dictionaries, so use its indexed table
        # rather than the coordinate API that parses the entire worksheet.
        border = Border()
        cell_styles = getattr(workbook, "_cell_styles", ())
        borders = getattr(workbook, "_borders", ())
        fill = format_to_fill(payload)
        if 0 <= style_id < len(cell_styles):
            source_style = cell_styles[style_id]
            border_id = source_style.borderId
            if 0 <= border_id < len(borders):
                border = self.source_borders.get(border_id)
                if border is None:
                    border = deepcopy(borders[border_id])
                    self.source_borders[border_id] = border
            fills = getattr(workbook, "_fills", ())
            if 0 <= source_style.fillId < len(fills):
                source_fill = fills[source_style.fillId]
                if isinstance(source_fill, GradientFill):
                    # Native gradient color payloads flatten alpha/theme
                    # details. The hydrated stylesheet preserves those.
                    fill = self.source_gradients.get(source_style.fillId)
                    if fill is None:
                        fill = deepcopy(source_fill)
                        self.source_gradients[source_style.fillId] = fill
        cached = StyleComponents(
            format_to_font(payload),
            fill,
            border,
            format_to_alignment(payload),
            number_format,
            is_date_format(number_format),
            is_timedelta_format(number_format),
        )
        self.styles[style_id] = cached
        return cached

    def merge_bounds(self, worksheet: Any) -> tuple[tuple[int, int, int, int], ...]:
        state = (
            worksheet._merged_ranges,
            worksheet._collection_merged_ranges,
            worksheet._pending_unmerged_ranges,
        )
        cached = self.merged_bounds.get(worksheet.title)
        if cached is not None and self.merged_state[worksheet.title] == state:
            return cached
        # Collection and worksheet merge mutations must both invalidate retained
        # cells' view, including remove/add operations that keep the same size.
        reader = worksheet._workbook._rust_reader
        refs = set(reader.read_merged_ranges(worksheet.title)) if reader is not None else set()
        refs.update(worksheet._merged_ranges)
        refs.update(worksheet._collection_merged_ranges)
        refs.difference_update(getattr(worksheet, "_pending_unmerged_ranges", ()))
        bounds = []
        for ref in refs:
            try:
                min_col, min_row, max_col, max_row = range_boundaries(str(ref))
            except (TypeError, ValueError):
                continue
            if None not in (min_col, min_row, max_col, max_row):
                bounds.append((min_row, min_col, max_row, max_col))
        cached = tuple(bounds)
        self.merged_bounds[worksheet.title] = cached
        self.merged_state[worksheet.title] = tuple(part.copy() for part in state)
        return cached

    def merge_at(self, worksheet: Any, row: int, col: int) -> tuple[int, int, int, int] | None:
        merged_bounds = self.merge_bounds(worksheet)
        if not merged_bounds:
            return None
        for bounds in merged_bounds:
            min_row, min_col, max_row, max_col = bounds
            if min_row <= row <= max_row and min_col <= col <= max_col:
                return bounds
        return None


def workbook_style_cache(workbook: Any) -> StreamingStyleCache:
    """Return a fresh cache after native-reader replacement or merge mutation."""
    reader = workbook._rust_reader
    cache = getattr(workbook, "_streaming_style_cache", None)
    if cache is None or cache.reader_identity != id(reader):
        cache = StreamingStyleCache(reader)
        workbook._streaming_style_cache = cache
    return cache


def streaming_date_style_ids(workbook: Any) -> tuple[Any, frozenset[int]]:
    """Reuse stylesheet classifications without resolving or retaining formats."""
    return workbook_style_cache(workbook).date_style_ids(workbook)


def cell_style_components(cell: Any) -> StyleComponents:
    workbook = cell._ws._workbook
    if workbook._rust_reader is None:
        return _DEFAULT_STYLE
    cache = workbook_style_cache(workbook)
    merged = cache.merge_at(cell._ws, cell._row, cell._col)
    if merged is not None and (cell._row, cell._col) != merged[:2]:
        return _DEFAULT_STYLE
    return cache.components(workbook, cell._style_id)


def cell_border(cell: Any) -> Border:
    workbook = cell._ws._workbook
    reader = workbook._rust_reader
    if reader is None:
        return _DEFAULT_STYLE.border
    cache = workbook_style_cache(workbook)
    merged = cache.merge_at(cell._ws, cell._row, cell._col)
    if merged is not None:
        min_row, min_col, max_row, max_col = merged
        start = border_payload_to_border(
            reader.read_cell_border(cell._ws.title, f"{get_column_letter(min_col)}{min_row}")
        )
        end = border_payload_to_border(
            reader.read_cell_border(cell._ws.title, f"{get_column_letter(max_col)}{max_row}")
        )
        anchor = cell._ws._cells.get((min_row, min_col))
        local_border = getattr(anchor, "_border", None)
        if isinstance(local_border, Border):
            start = local_border
        right = end.right if end.right is not None and end.right != Side() else start.right
        bottom = end.bottom if end.bottom is not None and end.bottom != Side() else start.bottom
        if (cell._row, cell._col) == (min_row, min_col):
            return deepcopy(
                Border(
                    left=start.left,
                    right=right,
                    top=start.top,
                    bottom=bottom,
                    diagonal=start.diagonal,
                    diagonalUp=start.diagonalUp,
                    diagonalDown=start.diagonalDown,
                    outline=start.outline,
                    vertical=start.vertical,
                    horizontal=start.horizontal,
                    start=start.start,
                    end=start.end,
                )
            )
        return deepcopy(
            Border(
                left=start.left if cell._col == min_col else None,
                right=right if cell._col == max_col else None,
                top=start.top if cell._row == min_row else None,
                bottom=bottom if cell._row == max_row else None,
            )
        )
    return cache.components(workbook, cell._style_id).border


def cell_fill(cell: Any) -> PatternFill | GradientFill:
    """Share immutable pattern fills; preserve fresh mutable gradient values."""
    fill = cell._resolved_style().fill
    return deepcopy(fill) if isinstance(fill, GradientFill) else fill
