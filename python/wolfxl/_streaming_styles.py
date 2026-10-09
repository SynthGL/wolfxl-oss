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
from wolfxl._styles import Alignment, Border, Font, PatternFill
from wolfxl.styles.fills import GradientFill
from wolfxl.styles.numbers import BUILTIN_FORMATS, BUILTIN_FORMATS_MAX_SIZE, is_timedelta_format
from wolfxl._worksheet_collections import _loaded_merged_range_refs, _merged_border_for_cell
from wolfxl._worksheet_merged_index import MergedRangeIndex
from wolfxl.utils.cell import range_boundaries
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

# A short tuple scan is cheaper for ordinary headers with a few merges.
_MERGED_INDEX_THRESHOLD = 8


class StreamingStyleCache:
    """Cache source styles and merge metadata without retaining a native reader."""

    __slots__ = (
        "reader_identity", "styles", "merged_bounds", "merged_indexes", "merged_priorities",
        "source_borders", "source_gradients", "date_ids",
    )

    def __init__(self, reader: Any) -> None:
        self.reader_identity = id(reader)
        self.styles: dict[int, StyleComponents] = {}
        self.merged_bounds: dict[str, tuple[tuple[int, int, int, int], ...]] = {}
        self.merged_indexes: dict[str, MergedRangeIndex] = {}
        self.merged_priorities: dict[str, dict[tuple[int, int, int, int], int]] = {}
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
        cached = self.merged_bounds.get(worksheet.title)
        if cached is not None:
            return cached
        # Worksheet merge bookkeeping includes pending merge/unmerge changes.
        # Its source lookup is a native metadata-only read, cached per sheet.
        refs = _loaded_merged_range_refs(worksheet)
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
        if len(cached) > _MERGED_INDEX_THRESHOLD:
            self.merged_indexes[worksheet.title] = MergedRangeIndex(refs)
            priorities = {}
            for rank, bound in enumerate(cached):
                priorities.setdefault(bound, rank)
            self.merged_priorities[worksheet.title] = priorities
        return cached

    def merge_at(self, worksheet: Any, row: int, col: int) -> tuple[int, int, int, int] | None:
        merged_bounds = self.merge_bounds(worksheet)
        if not merged_bounds:
            return None
        index = self.merged_indexes.get(worksheet.title)
        if index is not None:
            return index.bounds_for_cell(row, col, self.merged_priorities[worksheet.title])
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
    if cache.merge_at(cell._ws, cell._row, cell._col) is not None:
        border = _merged_border_for_cell(cell._ws, cell._row, cell._col)
        if border is not None:
            # The merge helper already returns detached nested Color values.
            return border
        read_merged_border = getattr(reader, "read_merged_cell_border", None)
        if read_merged_border is not None:
            payload = read_merged_border(cell._ws.title, cell._row, cell._col)
        else:
            # Older extension modules preserve their coordinate border path.
            payload = reader.read_cell_border(cell._ws.title, cell.coordinate)
        return border_payload_to_border(payload)
    return cache.components(workbook, cell._style_id).border


def cell_fill(cell: Any) -> PatternFill | GradientFill:
    """Share immutable pattern fills; preserve fresh mutable gradient values."""
    fill = cell._resolved_style().fill
    return deepcopy(fill) if isinstance(fill, GradientFill) else fill
