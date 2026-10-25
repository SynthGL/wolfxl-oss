"""Worksheet record iteration, visibility, and schema helpers."""

from __future__ import annotations

import datetime as _dt
from collections.abc import Iterator
from decimal import Decimal
from typing import TYPE_CHECKING, Any

from wolfxl._utils import rowcol_to_a1

if TYPE_CHECKING:
    from wolfxl._worksheet import Worksheet

_ERROR_CODES = {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"}


def canonical_data_type(value: Any) -> str:
    """Map a Python value to the canonical Rust reader data-type label."""
    if value is None:
        return "blank"
    if isinstance(value, bool):
        return "boolean"
    if isinstance(value, (int, float, Decimal)):
        return "number"
    if isinstance(value, str):
        if value in _ERROR_CODES:
            return "error"
        return "formula" if len(value) > 1 and value.startswith("=") else "string"
    if isinstance(value, (_dt.datetime, _dt.date, _dt.time, _dt.timedelta)):
        return "datetime"
    return "string"


def iter_cell_records(
    ws: Worksheet,
    min_row: int | None = None,
    max_row: int | None = None,
    min_col: int | None = None,
    max_col: int | None = None,
    *,
    data_only: bool | None = None,
    include_format: bool = True,
    include_empty: bool = False,
    include_formula_blanks: bool = True,
    include_coordinate: bool = True,
    include_style_id: bool = True,
    include_extended_format: bool = True,
    include_cached_formula_value: bool = False,
) -> Iterator[dict[str, Any]]:
    """Iterate worksheet cells as compact dictionaries."""
    if ws._workbook._rust_reader is None:  # noqa: SLF001
        yield from iter_cell_records_python(
            ws,
            min_row=min_row,
            max_row=max_row,
            min_col=min_col,
            max_col=max_col,
            include_empty=include_empty,
            include_coordinate=include_coordinate,
        )
        return

    reader = ws._workbook._rust_reader  # noqa: SLF001
    effective_data_only = (
        ws._workbook._data_only if data_only is None else data_only  # noqa: SLF001
    )
    overlay = ws._collect_pending_overlay()  # noqa: SLF001
    row_min, row_max, col_min, col_max, range_str = _record_scan_range(
        ws,
        min_row=min_row,
        max_row=max_row,
        min_col=min_col,
        max_col=max_col,
        include_empty=include_empty,
        has_overlay=bool(overlay),
    )
    records = reader.read_sheet_records(
        ws._title,  # noqa: SLF001
        range_str,
        effective_data_only,
        include_format,
        include_empty,
        include_formula_blanks,
        include_coordinate,
        include_style_id,
        include_extended_format,
        include_cached_formula_value,
    )

    if not overlay:
        yield from records
        return

    seen: set[tuple[int, int]] = set()
    for record in records:
        row, col = int(record["row"]), int(record["column"])
        key = (row, col)
        if key not in overlay:
            yield record
            continue
        seen.add(key)
        patched = _patched_overlay_record(
            record,
            overlay[key],
            include_empty=include_empty,
        )
        if patched is not None:
            yield patched

    for (row, col), value in overlay.items():
        extra = _extra_overlay_record(
            row,
            col,
            value,
            seen=seen,
            row_min=row_min,
            row_max=row_max,
            col_min=col_min,
            col_max=col_max,
            include_empty=include_empty,
            include_coordinate=include_coordinate,
        )
        if extra is not None:
            yield extra


def _record_scan_range(
    ws: Worksheet,
    *,
    min_row: int | None,
    max_row: int | None,
    min_col: int | None,
    max_col: int | None,
    include_empty: bool,
    has_overlay: bool,
) -> tuple[int, int | None, int, int | None, str | None]:
    """Return scan bounds and optional A1 range for Rust record reads."""
    unbounded_sparse_read = (
        min_row is None
        and max_row is None
        and min_col is None
        and max_col is None
        and not include_empty
        and not has_overlay
    )
    if unbounded_sparse_read:
        return 1, None, 1, None, None
    row_min = min_row or 1
    row_max = max_row or ws._bounds_max_row()  # noqa: SLF001
    col_min = min_col or 1
    col_max = max_col or ws._bounds_max_col()  # noqa: SLF001
    range_str = f"{rowcol_to_a1(row_min, col_min)}:{rowcol_to_a1(row_max, col_max)}"
    return row_min, row_max, col_min, col_max, range_str


def _patched_overlay_record(
    record: dict[str, Any],
    value: Any,
    *,
    include_empty: bool,
) -> dict[str, Any] | None:
    """Return a Rust record patched with a pending overlay value."""
    if value is None and not include_empty:
        return None
    patched = dict(record)
    patched["value"] = value
    patched["data_type"] = canonical_data_type(value)
    if isinstance(value, str) and value.startswith("="):
        patched["formula"] = value[1:]
    else:
        patched.pop("formula", None)
    patched.pop("cached_value", None)
    return patched


def _extra_overlay_record(
    row: int,
    col: int,
    value: Any,
    *,
    seen: set[tuple[int, int]],
    row_min: int,
    row_max: int | None,
    col_min: int,
    col_max: int | None,
    include_empty: bool,
    include_coordinate: bool,
) -> dict[str, Any] | None:
    """Build a record for a pending edit outside the Rust record stream."""
    if (row, col) in seen:
        return None
    if row_max is None or col_max is None:
        return None
    if not (row_min <= row <= row_max and col_min <= col <= col_max):
        return None
    if value is None and not include_empty:
        return None
    record: dict[str, Any] = {
        "row": row,
        "column": col,
        "value": value,
        "data_type": canonical_data_type(value),
    }
    if isinstance(value, str) and value.startswith("="):
        record["formula"] = value[1:]
    if include_coordinate:
        record["coordinate"] = rowcol_to_a1(row, col)
    return record


def iter_cell_records_python(
    ws: Worksheet,
    *,
    min_row: int | None,
    max_row: int | None,
    min_col: int | None,
    max_col: int | None,
    include_empty: bool,
    include_coordinate: bool = True,
) -> Iterator[dict[str, Any]]:
    """Iterate cell records from Python-side cells in pure write mode."""
    row_min = min_row or 1
    row_max = max_row or ws._bounds_max_row()  # noqa: SLF001
    col_min = min_col or 1
    col_max = max_col or ws._bounds_max_col()  # noqa: SLF001
    for row in range(row_min, row_max + 1):
        for col in range(col_min, col_max + 1):
            cell = ws._get_or_create_cell(row, col)  # noqa: SLF001
            value = cell.value
            if value is None and not include_empty:
                continue
            record: dict[str, Any] = {
                "row": row,
                "column": col,
                "value": value,
                "data_type": canonical_data_type(value),
            }
            if include_coordinate:
                record["coordinate"] = rowcol_to_a1(row, col)
            yield record


def cached_formula_values(ws: Worksheet, *, qualified: bool = False) -> dict[str, Any]:
    """Return saved formula cache values for a worksheet."""
    workbook = ws._workbook  # noqa: SLF001
    if workbook._rust_reader is None:  # noqa: SLF001
        return {}
    values = dict(workbook._rust_reader.read_cached_formula_values(ws._title))  # noqa: SLF001
    if not qualified:
        return values
    return {f"{ws._title}!{cell_ref}": value for cell_ref, value in values.items()}  # noqa: SLF001


def sheet_visibility(ws: Worksheet) -> dict[str, Any]:
    """Return cached hidden row/column and outline metadata for a worksheet."""
    if ws._sheet_visibility_cache is not None:  # noqa: SLF001
        return ws._sheet_visibility_cache  # noqa: SLF001

    workbook = ws._workbook  # noqa: SLF001
    if workbook._rust_reader is None:  # noqa: SLF001
        ws._sheet_visibility_cache = {  # noqa: SLF001
            "hidden_rows": [],
            "hidden_columns": [],
            "row_outline_levels": {},
            "column_outline_levels": {},
        }
        return ws._sheet_visibility_cache  # noqa: SLF001
    ws._sheet_visibility_cache = dict(  # noqa: SLF001
        workbook._rust_reader.read_sheet_visibility(ws._title)  # noqa: SLF001
    )
    return ws._sheet_visibility_cache  # noqa: SLF001


def classify_format(fmt: str) -> str:
    """Classify an Excel number-format string through the Rust classifier."""
    from wolfxl._rust import classify_format as _classify_format

    return _classify_format(fmt)


def schema(ws: Worksheet) -> dict[str, Any]:
    """Infer a worksheet schema through the shared Rust schema engine."""
    from wolfxl._cell import _UNSET
    from wolfxl._rust import infer_sheet_schema as _infer_sheet_schema

    max_row = ws._bounds_max_row()  # noqa: SLF001
    max_col = ws._bounds_max_col()  # noqa: SLF001
    values: list[list[Any]] = [[None] * max_col for _ in range(max_row)]
    formats: list[list[str | None]] = [[None] * max_col for _ in range(max_row)]
    for record in ws.iter_cell_records(
        include_format=True,
        include_empty=False,
        include_coordinate=False,
    ):
        row = int(record["row"]) - 1
        col = int(record["column"]) - 1
        if row >= max_row or col >= max_col:
            continue
        values[row][col] = record.get("value")
        number_format = record.get("number_format")
        if number_format:
            formats[row][col] = number_format
    for (row, col), cell in ws._cells.items():  # noqa: SLF001
        if row > max_row or col > max_col:
            continue
        number_format = getattr(cell, "_number_format", _UNSET)
        if number_format is not _UNSET and number_format:
            formats[row - 1][col - 1] = number_format
    return _infer_sheet_schema(values, ws._title, formats)  # noqa: SLF001
