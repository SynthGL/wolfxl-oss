"""Worksheet append and bulk-write buffer helpers."""

from __future__ import annotations

from collections.abc import Iterable
from typing import TYPE_CHECKING, Any, Optional

from wolfxl._cell_payloads import (
    alignment_to_format_dict,
    border_payload_to_border,
    border_to_rust_dict,
    fill_to_format_dict,
    format_to_alignment,
    format_to_fill,
    format_to_font,
    format_to_protection,
    font_to_format_dict,
    protection_to_format_dict,
)
from wolfxl.utils.cell import column_index_from_string

if TYPE_CHECKING:
    from wolfxl._worksheet import Worksheet

_BATCHABLE_TYPES = (bool, int, float, str)
_BATCHABLE_EXACT_TYPES = frozenset(_BATCHABLE_TYPES)
BulkWrite = tuple[list[list[Any]], int, int, bool, Optional[list[list[Optional[dict[str, Any]]]]]]


def append_row(ws: Worksheet, iterable: Iterable[Any]) -> None:
    """Append a row of values to the worksheet's write buffer."""
    row = _coerce_append_row(iterable)
    ws._plain_cell_fast_path = False  # noqa: SLF001
    if not ws._append_buffer:  # noqa: SLF001
        if getattr(ws._workbook, "_rust_reader", None) is not None:
            ws._next_append_row = max(ws._next_append_row, ws._bounds_max_row() + 1)  # noqa: SLF001
        ws._append_buffer_start = ws._next_append_row  # noqa: SLF001
    ws._append_buffer.append(row)  # noqa: SLF001
    ncols = len(row)
    if ncols > ws._max_col_idx:  # noqa: SLF001
        ws._max_col_idx = ncols  # noqa: SLF001
    ws._current_row = ws._next_append_row  # noqa: SLF001
    ws._next_append_row += 1  # noqa: SLF001


def _coerce_append_row(iterable: Iterable[Any]) -> list[Any]:
    if isinstance(iterable, (str, bytes, bytearray)):
        raise TypeError("Value must be a list, tuple, range or generator, or a dict")
    if not isinstance(iterable, dict):
        return list(iterable)
    values: dict[int, Any] = {}
    for key, value in iterable.items():
        if isinstance(key, str):
            col_idx = column_index_from_string(key)
        elif isinstance(key, int) and not isinstance(key, bool):
            col_idx = key
        else:
            raise TypeError("append dict keys must be column letters or 1-based integers")
        if col_idx < 1:
            raise ValueError("append dict column indexes must be >= 1")
        values[col_idx] = value
    if not values:
        return []
    row = [None] * max(values)
    for col_idx, value in values.items():
        row[col_idx - 1] = value
    return row


def write_rows(
    ws: Worksheet,
    rows: list[list[Any]],
    start_row: int = 1,
    start_col: int = 1,
    *,
    copy: bool = True,
    plain_values_only: bool = False,
    style_grid: list[list[dict[str, Any] | None]] | None = None,
) -> None:
    """Queue a 2D value grid for a later batch write."""
    if not rows:
        return
    ws._plain_cell_fast_path = False  # noqa: SLF001
    grid = [list(row) for row in rows] if copy else rows
    ws._bulk_writes.append(  # noqa: SLF001
        (grid, start_row, start_col, plain_values_only, style_grid)
    )


def styled_rows_style_grid(
    rows: list[list[Any]],
    styles: Any,
    start_row: int,
    start_col: int,
) -> list[list[dict[str, Any] | None]]:
    """Return a normalized same-shape format grid for ``write_styled_rows``."""
    if callable(styles):
        cache: dict[tuple[tuple[str, Any], ...], dict[str, Any]] = {}
        return [
            [
                _normalize_cached_by_content(
                    styles(start_row + row_offset, start_col + col_offset, value),
                    cache,
                )
                for col_offset, value in enumerate(row)
            ]
            for row_offset, row in enumerate(rows)
        ]
    if isinstance(styles, dict) and not _looks_like_style_spec(styles):
        cache: dict[int, dict[str, Any] | None] = {}
        return [
            [
                _normalize_cached_by_id(
                    _style_from_column_map(styles, col_offset + 1, start_col + col_offset),
                    cache,
                )
                for col_offset, _value in enumerate(row)
            ]
            for row in rows
        ]
    if isinstance(styles, (list, tuple)):
        if _looks_like_style_grid(styles):
            return _normalize_style_grid(rows, styles)
        cache: dict[int, dict[str, Any] | None] = {}
        return [
            [
                _normalize_cached_by_id(
                    styles[col_offset] if col_offset < len(styles) else None,
                    cache,
                )
                for col_offset, _value in enumerate(row)
            ]
            for row in rows
        ]
    return [[normalize_style_spec(styles) for _value in row] for row in rows]


def normalize_style_spec(spec: Any) -> dict[str, Any] | None:
    """Convert a public style spec into the Rust writer's flat format dict."""
    if spec is None:
        return None
    if isinstance(spec, str):
        return {"number_format": spec} if spec else None
    if isinstance(spec, dict):
        fmt: dict[str, Any] = {}
        nested_keys = {"font", "fill", "border", "alignment", "number_format", "protection"}
        if nested_keys.intersection(spec):
            if spec.get("font") is not None:
                fmt.update(font_to_format_dict(spec["font"]))
            if spec.get("fill") is not None:
                fmt.update(fill_to_format_dict(spec["fill"]))
            if spec.get("border") is not None:
                border = border_to_rust_dict(spec["border"])
                if border:
                    fmt["border"] = border
            if spec.get("alignment") is not None:
                fmt.update(alignment_to_format_dict(spec["alignment"]))
            if spec.get("protection") is not None:
                fmt.update(protection_to_format_dict(spec["protection"]))
            number_format = spec.get("number_format")
            if number_format is not None:
                fmt["number_format"] = number_format
            for key, value in spec.items():
                if key not in nested_keys and value is not None:
                    fmt[key] = value
            return fmt or None
        return {key: value for key, value in spec.items() if value is not None} or None

    cls_name = type(spec).__name__
    if cls_name == "Font":
        return font_to_format_dict(spec) or None
    if cls_name in {"PatternFill", "GradientFill"}:
        return fill_to_format_dict(spec) or None
    if cls_name == "Border":
        border = border_to_rust_dict(spec)
        return {"border": border} if border else None
    if cls_name == "Alignment":
        return alignment_to_format_dict(spec) or None
    if cls_name == "Protection":
        return protection_to_format_dict(spec) or None
    raise TypeError(
        "style specs must be None, a number-format string, a dict, or a "
        "Font/Fill/Border/Alignment/Protection object"
    )


def materialize_append_buffer(ws: Worksheet) -> None:
    """Convert a worksheet append buffer into dirty Cell objects."""
    start = ws._append_buffer_start  # noqa: SLF001
    buffer = ws._append_buffer  # noqa: SLF001
    if not buffer:
        return
    ws._append_buffer = []  # noqa: SLF001
    for row_offset, row_values in enumerate(buffer):
        row_index = start + row_offset
        for col_index, value in enumerate(row_values, start=1):
            if hasattr(value, "value") and hasattr(value, "coordinate"):
                value = value.value
            ws.cell(row=row_index, column=col_index, value=value)


def materialize_bulk_writes(ws: Worksheet) -> None:
    """Convert bulk write buffers into dirty Cell objects."""
    writes = ws._bulk_writes  # noqa: SLF001
    if not writes:
        return
    ws._bulk_writes = []  # noqa: SLF001
    for grid, start_row, start_col, _plain_values_only, style_grid in writes:
        for row_offset, row_values in enumerate(grid):
            for col_offset, value in enumerate(row_values):
                if value is not None:
                    cell = ws.cell(
                        row=start_row + row_offset,
                        column=start_col + col_offset,
                        value=value,
                    )
                    if style_grid is not None:
                        try:
                            fmt = style_grid[row_offset][col_offset]
                        except IndexError:
                            fmt = None
                        if fmt:
                            apply_format_dict_to_cell(cell, fmt)


def extract_non_batchable(
    grid: list[list[Any]],
    start_row: int,
    start_col: int,
) -> list[tuple[int, int, Any]]:
    """Extract values that still require per-cell writes from a batch grid.

    The Rust bulk writer already handles plain strings, formulas expressed as
    strings, booleans, numbers, error literals, and blanks. Values such as
    dates and datetimes still use the per-cell path so the existing default
    number-format behavior is preserved.
    """
    individual: list[tuple[int, int, Any]] = []
    for row_offset, row_values in enumerate(grid):
        for col_offset, value in enumerate(row_values):
            if value is None:
                continue
            value_type = type(value)
            if (
                value_type not in _BATCHABLE_EXACT_TYPES
                and not isinstance(value, _BATCHABLE_TYPES)
            ):
                individual.append(
                    (start_row + row_offset, start_col + col_offset, value)
                )
                row_values[col_offset] = None
    return individual


def batch_write_dicts(
    ws: Worksheet,
    batch_fn: Any,
    entries: list[tuple[int, int, dict[str, Any]]],
) -> None:
    """Build a bounding-box grid of dicts and call a batch Rust method."""
    _batch_write_grid(ws, batch_fn, entries)


def batch_write_style_ids(
    ws: Worksheet,
    batch_fn: Any,
    entries: list[tuple[int, int, int]],
) -> None:
    """Build a bounding-box grid of style IDs and call a batch Rust method."""
    _batch_write_grid(ws, batch_fn, entries)


def _batch_write_grid(
    ws: Worksheet,
    batch_fn: Any,
    entries: list[tuple[int, int, Any]],
) -> None:
    """Build a bounding-box grid for sparse row/column payload entries."""
    min_row = entries[0][0]
    min_col = entries[0][1]
    max_row = min_row
    max_col = min_col
    for row, col, _payload in entries:
        if row < min_row:
            min_row = row
        if row > max_row:
            max_row = row
        if col < min_col:
            min_col = col
        if col > max_col:
            max_col = col

    num_rows = max_row - min_row + 1
    num_cols = max_col - min_col + 1
    grid: list[list[Any]] = [[None] * num_cols for _ in range(num_rows)]
    for row, col, payload in entries:
        grid[row - min_row][col - min_col] = payload

    from wolfxl._utils import rowcol_to_a1

    start = rowcol_to_a1(min_row, min_col)
    batch_fn(ws._title, start, grid)  # noqa: SLF001


def intern_style_grid(
    writer: Any,
    style_grid: list[list[dict[str, Any] | None]],
) -> list[list[int | None]]:
    """Intern repeated format dicts and return a same-shape style-id grid."""
    cache: dict[tuple[tuple[str, Any], ...], int] = {}
    id_cache: dict[int, int] = {}
    out: list[list[int | None]] = []
    for row in style_grid:
        out_row: list[int | None] = []
        for fmt in row:
            if not fmt:
                out_row.append(None)
                continue
            fmt_id = id(fmt)
            style_id = id_cache.get(fmt_id)
            if style_id is not None:
                out_row.append(style_id)
                continue
            key = _format_cache_key(fmt)
            style_id = cache.get(key)
            if style_id is None:
                style_id = int(writer.intern_format(fmt))
                cache[key] = style_id
            id_cache[fmt_id] = style_id
            out_row.append(style_id)
        out.append(out_row)
    return out


def apply_format_dict_to_cell(cell: Any, fmt: dict[str, Any]) -> None:
    """Apply a normalized format dict when a queued bulk write is materialized."""
    if _font_keys(fmt):
        cell.font = format_to_font(fmt)
    if "bg_color" in fmt or "gradient" in fmt:
        cell.fill = format_to_fill(fmt)
    if "border" in fmt:
        cell.border = border_payload_to_border(fmt["border"])
    if _alignment_keys(fmt):
        cell.alignment = format_to_alignment(fmt)
    if "number_format" in fmt:
        cell.number_format = fmt["number_format"]
    protection = format_to_protection(fmt)
    if protection is not None:
        cell.protection = protection


def _style_from_column_map(styles: dict[Any, Any], local_col: int, absolute_col: int) -> Any:
    from wolfxl.utils.cell import get_column_letter

    for key in (
        local_col,
        absolute_col,
        get_column_letter(local_col),
        get_column_letter(absolute_col),
    ):
        if key in styles:
            return styles[key]
    return None


def _looks_like_style_grid(styles: Any) -> bool:
    if not styles:
        return False
    first = styles[0]
    return isinstance(first, (list, tuple))


def _normalize_style_grid(
    rows: list[list[Any]],
    styles: Any,
) -> list[list[dict[str, Any] | None]]:
    out: list[list[dict[str, Any] | None]] = []
    cache: dict[int, dict[str, Any] | None] = {}
    for row_offset, row in enumerate(rows):
        style_row = styles[row_offset] if row_offset < len(styles) else []
        out.append(
            [
                _normalize_cached_by_id(
                    style_row[col_offset] if col_offset < len(style_row) else None,
                    cache,
                )
                for col_offset, _value in enumerate(row)
            ]
        )
    return out


def _normalize_cached_by_id(
    spec: Any,
    cache: dict[int, dict[str, Any] | None],
) -> dict[str, Any] | None:
    if spec is None:
        return None
    key = id(spec)
    if key not in cache:
        cache[key] = normalize_style_spec(spec)
    return cache[key]


def _normalize_cached_by_content(
    spec: Any,
    cache: dict[tuple[tuple[str, Any], ...], dict[str, Any]],
) -> dict[str, Any] | None:
    fmt = normalize_style_spec(spec)
    if fmt is None:
        return None
    key = _format_cache_key(fmt)
    cached = cache.get(key)
    if cached is None:
        cache[key] = fmt
        return fmt
    return cached


def _looks_like_style_spec(value: dict[Any, Any]) -> bool:
    style_keys = {
        "font",
        "fill",
        "border",
        "alignment",
        "number_format",
        "protection",
        "bold",
        "italic",
        "underline",
        "strikethrough",
        "font_name",
        "font_size",
        "font_color",
        "bg_color",
        "gradient",
        "h_align",
        "v_align",
        "wrap",
        "rotation",
        "indent",
        "locked",
        "hidden",
    }
    return bool(style_keys.intersection(value))


def _format_cache_key(fmt: dict[str, Any]) -> tuple[tuple[str, Any], ...]:
    return tuple(sorted((key, _freeze_style_value(value)) for key, value in fmt.items()))


def _freeze_style_value(value: Any) -> Any:
    if isinstance(value, dict):
        return tuple(sorted((key, _freeze_style_value(val)) for key, val in value.items()))
    if isinstance(value, (list, tuple)):
        return tuple(_freeze_style_value(item) for item in value)
    return value


def _font_keys(fmt: dict[str, Any]) -> bool:
    return any(
        key in fmt
        for key in (
            "bold",
            "italic",
            "underline",
            "strikethrough",
            "font_name",
            "font_size",
            "font_color",
        )
    )


def _alignment_keys(fmt: dict[str, Any]) -> bool:
    return any(key in fmt for key in ("h_align", "v_align", "wrap", "rotation", "indent"))
