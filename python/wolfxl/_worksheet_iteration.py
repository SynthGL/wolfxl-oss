"""Worksheet row and column iteration helpers."""

from __future__ import annotations

from collections.abc import Iterable, Iterator
from operator import itemgetter
from typing import TYPE_CHECKING, Any

from wolfxl._utils import rowcol_to_a1

if TYPE_CHECKING:
    from wolfxl._worksheet import Worksheet


def iter_rows(
    ws: Worksheet,
    min_row: int | None = None,
    max_row: int | None = None,
    min_col: int | None = None,
    max_col: int | None = None,
    values_only: bool = False,
) -> Iterator[tuple[Any, ...]]:
    """Iterate worksheet rows with streaming and bulk-read fast paths."""
    workbook = ws._workbook  # noqa: SLF001
    rust_reader = getattr(workbook, "_rust_reader", None)
    if rust_reader is not None and getattr(workbook, "_source_path", None):
        from wolfxl._streaming import should_auto_stream, stream_iter_rows

        stream_now = bool(getattr(workbook, "_read_only", False)) or should_auto_stream(ws)
        if stream_now:
            return stream_iter_rows(
                ws,
                min_row,
                max_row,
                min_col,
                max_col,
                values_only=values_only,
            )

    if values_only and rust_reader is not None:
        return iter_rows_bulk(ws, min_row, max_row, min_col, max_col)

    if _can_prefill_cell_values(ws):
        return iter_rows_cells_with_bulk_values(
            ws,
            min_row,
            max_row,
            min_col,
            max_col,
            data_only=getattr(workbook, "_data_only", False),
        )

    if (
        min_row is None
        and max_row is None
        and min_col is None
        and max_col is None
        and rust_reader is None
        and not ws._cells  # noqa: SLF001
        and not ws._append_buffer  # noqa: SLF001
        and not ws._bulk_writes  # noqa: SLF001
    ):
        return iter(())

    row_min = min_row or 1
    row_max = max_row or ws._max_row()  # noqa: SLF001
    col_min = min_col or 1
    col_max = max_col or ws._max_col()  # noqa: SLF001

    return _iter_materialized_rows(ws, row_min, row_max, col_min, col_max, values_only)


def _iter_materialized_rows(
    ws: Worksheet,
    row_min: int,
    row_max: int,
    col_min: int,
    col_max: int,
    values_only: bool,
) -> Iterator[tuple[Any, ...]]:
    """Yield rows from materialized worksheet cells."""
    for row in range(row_min, row_max + 1):
        if values_only:
            yield tuple(
                ws._get_or_create_cell(row, col).value  # noqa: SLF001
                for col in range(col_min, col_max + 1)
            )
        else:
            yield tuple(
                ws._get_or_create_cell(row, col)  # noqa: SLF001
                for col in range(col_min, col_max + 1)
            )


def iter_rows_cells_with_bulk_values(
    ws: Worksheet,
    min_row: int | None,
    max_row: int | None,
    min_col: int | None,
    max_col: int | None,
    *,
    data_only: bool,
) -> Iterator[tuple[Any, ...]]:
    """Yield Cell rows while filling values from one bulk native read."""
    reader = ws._workbook._rust_reader  # noqa: SLF001
    if not hasattr(reader, "read_sheet_values_plain"):
        row_min = min_row or 1
        row_max = max_row or ws._max_row()  # noqa: SLF001
        col_min = min_col or 1
        col_max = max_col or ws._max_col()  # noqa: SLF001
        for row in range(row_min, row_max + 1):
            yield tuple(
                ws._get_or_create_cell(row, col)  # noqa: SLF001
                for col in range(col_min, col_max + 1)
            )
        return

    row_min = min_row or 1
    row_max = max_row or ws._max_row()  # noqa: SLF001
    col_min = min_col or 1
    col_max = max_col or ws._max_col()  # noqa: SLF001
    expected_cols = col_max - col_min + 1
    range_str = f"{rowcol_to_a1(row_min, col_min)}:{rowcol_to_a1(row_max, col_max)}"
    rows = reader.read_sheet_values_plain(ws._title, range_str, data_only)  # noqa: SLF001

    get_cell = _plain_cell_getter(ws)
    row_count = row_max - row_min + 1
    for row_offset in range(row_count):
        row_idx = row_min + row_offset
        values = rows[row_offset] if row_offset < len(rows) else ()
        width = len(values)
        cells = []
        for col_offset in range(expected_cols):
            cell = get_cell(row_idx, col_min + col_offset)
            if hasattr(cell, "_value_dirty") and not cell._value_dirty:  # noqa: SLF001
                value = values[col_offset] if col_offset < width else None
                cell._value = value  # noqa: SLF001
                cell._value_is_plain = _bulk_value_is_plain(ws, value)  # noqa: SLF001
            cells.append(cell)
        yield tuple(cells)


def iter_cols(
    ws: Worksheet,
    min_col: int | None = None,
    max_col: int | None = None,
    min_row: int | None = None,
    max_row: int | None = None,
    values_only: bool = False,
) -> Iterator[tuple[Any, ...]]:
    """Iterate worksheet columns with a bulk-read fast path when possible."""
    if values_only and getattr(ws._workbook, "_rust_reader", None) is not None:  # noqa: SLF001
        yield from iter_cols_bulk(ws, min_col, max_col, min_row, max_row)
        return

    if (
        _can_prefill_cell_values(ws)
        and not getattr(ws._workbook, "_read_only", False)  # noqa: SLF001
    ):
        yield from iter_cols_cells_with_bulk_values(
            ws,
            min_col,
            max_col,
            min_row,
            max_row,
            data_only=getattr(ws._workbook, "_data_only", False),  # noqa: SLF001
        )
        return

    if (
        min_row is None
        and max_row is None
        and min_col is None
        and max_col is None
        and getattr(ws._workbook, "_rust_reader", None) is None  # noqa: SLF001
        and not ws._cells  # noqa: SLF001
        and not ws._append_buffer  # noqa: SLF001
        and not ws._bulk_writes  # noqa: SLF001
    ):
        return

    row_min = min_row or 1
    row_max = max_row or ws._max_row()  # noqa: SLF001
    col_min = min_col or 1
    col_max = max_col or ws._max_col()  # noqa: SLF001

    for col in range(col_min, col_max + 1):
        if values_only:
            yield tuple(
                ws._get_or_create_cell(row, col).value  # noqa: SLF001
                for row in range(row_min, row_max + 1)
            )
        else:
            yield tuple(
                ws._get_or_create_cell(row, col)  # noqa: SLF001
                for row in range(row_min, row_max + 1)
            )


def iter_value_chunks(
    ws: Worksheet,
    min_row: int | None = None,
    max_row: int | None = None,
    min_col: int | None = None,
    max_col: int | None = None,
    *,
    chunk_size: int = 1024,
) -> Iterator[tuple[tuple[Any, ...], ...]]:
    """Yield ``iter_rows(values_only=True)`` rows grouped into chunks."""
    if chunk_size < 1:
        raise ValueError("chunk_size must be >= 1")

    workbook = ws._workbook  # noqa: SLF001
    if getattr(workbook, "_read_only", False) and getattr(workbook, "_source_path", None):
        from wolfxl._streaming import stream_value_chunks

        return stream_value_chunks(
            ws,
            min_row=min_row,
            max_row=max_row,
            min_col=min_col,
            max_col=max_col,
            chunk_size=chunk_size,
        )

    return _chunk_rows(
        iter_rows(
            ws,
            min_row=min_row,
            max_row=max_row,
            min_col=min_col,
            max_col=max_col,
            values_only=True,
        ),
        chunk_size,
    )


def _chunk_rows(
    rows: Iterable[tuple[Any, ...]],
    chunk_size: int,
) -> Iterator[tuple[tuple[Any, ...], ...]]:
    chunk: list[tuple[Any, ...]] = []
    for row in rows:
        chunk.append(row)
        if len(chunk) >= chunk_size:
            yield tuple(chunk)
            chunk = []
    if chunk:
        yield tuple(chunk)


def iter_cols_cells_with_bulk_values(
    ws: Worksheet,
    min_col: int | None,
    max_col: int | None,
    min_row: int | None,
    max_row: int | None,
    *,
    data_only: bool,
) -> Iterator[tuple[Any, ...]]:
    """Yield Cell columns while filling values from one bulk native read."""
    reader = ws._workbook._rust_reader  # noqa: SLF001
    if not hasattr(reader, "read_sheet_values_plain"):
        row_min = min_row or 1
        row_max = max_row or ws._max_row()  # noqa: SLF001
        col_min = min_col or 1
        col_max = max_col or ws._max_col()  # noqa: SLF001
        for col in range(col_min, col_max + 1):
            yield tuple(
                ws._get_or_create_cell(row, col)  # noqa: SLF001
                for row in range(row_min, row_max + 1)
            )
        return

    row_min = min_row or 1
    row_max = max_row or ws._max_row()  # noqa: SLF001
    col_min = min_col or 1
    col_max = max_col or ws._max_col()  # noqa: SLF001
    expected_rows = row_max - row_min + 1
    range_str = f"{rowcol_to_a1(row_min, col_min)}:{rowcol_to_a1(row_max, col_max)}"
    rows = reader.read_sheet_values_plain(ws._title, range_str, data_only)  # noqa: SLF001

    get_cell = _plain_cell_getter(ws)
    for col_offset, col_idx in enumerate(range(col_min, col_max + 1)):
        cells = []
        for row_offset in range(expected_rows):
            row_idx = row_min + row_offset
            values = rows[row_offset] if row_offset < len(rows) else ()
            cell = get_cell(row_idx, col_idx)
            if hasattr(cell, "_value_dirty") and not cell._value_dirty:  # noqa: SLF001
                value = values[col_offset] if col_offset < len(values) else None
                cell._value = value  # noqa: SLF001
                cell._value_is_plain = _bulk_value_is_plain(ws, value)  # noqa: SLF001
            cells.append(cell)
        yield tuple(cells)


def _can_prefill_cell_values(ws: Worksheet) -> bool:
    """Return True when Cell rows can share one bulk value read."""
    workbook = ws._workbook  # noqa: SLF001
    reader = getattr(workbook, "_rust_reader", None)
    if reader is None or not hasattr(reader, "read_sheet_values_plain"):
        return False
    if getattr(workbook, "_read_only", False):
        return False
    if getattr(workbook, "_data_only", False):
        return True
    if getattr(workbook, "_rich_text", False):
        return False
    if ws._pending_array_formulas:  # noqa: SLF001
        return False
    return not _reader_has_array_formulas(ws)


def _reader_has_array_formulas(ws: Worksheet) -> bool:
    """Check whether disk-backed array/data-table formulas need Cell metadata."""
    cached = ws._reader_has_array_formulas_cache  # noqa: SLF001
    if cached is not None:
        return cached
    reader = ws._workbook._rust_reader  # noqa: SLF001
    if reader is None or not hasattr(reader, "read_sheet_array_formulas"):
        return True
    has_array_formula_flag = getattr(reader, "has_sheet_array_formulas", None)
    if has_array_formula_flag is not None:
        has_array_formulas = bool(has_array_formula_flag(ws._title))  # noqa: SLF001
    else:
        has_array_formulas = bool(reader.read_sheet_array_formulas(ws._title))  # noqa: SLF001
    ws._reader_has_array_formulas_cache = has_array_formulas  # noqa: SLF001
    return has_array_formulas


def _bulk_value_is_plain(ws: Worksheet, value: Any) -> bool:
    """Return whether the prefetched value can bypass Cell formula checks."""
    return not (
        isinstance(value, str)
        and getattr(ws._workbook, "_rich_text", False)  # noqa: SLF001
    )


def _plain_cell_getter(ws: Worksheet):
    """Return the cheapest safe Cell getter for bulk value prefill."""
    if (  # noqa: SLF001
        ws._append_buffer
        or ws._merged_ranges
        or ws._collection_merged_ranges
        or ws._pending_collection_merged_ranges
    ):
        return ws._get_or_create_cell  # noqa: SLF001

    from wolfxl._cell import Cell, _merged_range_refs

    # Loading merge refs once also protects the fresh-cell fast path after
    # workbook-open metadata hydration becomes lazy.
    if _merged_range_refs(ws):
        return ws._get_or_create_cell  # noqa: SLF001

    cells = ws._cells  # noqa: SLF001
    if not cells:
        def create_plain_cell(row: int, col: int) -> Cell:
            key = (row, col)
            cell = cells.get(key)
            if cell is None:
                cell = Cell(ws, row, col)
                cells[key] = cell
            return cell

        return create_plain_cell

    def get_plain_cell(row: int, col: int) -> Cell:
        key = (row, col)
        cell = cells.get(key)
        if cell is None:
            cell = Cell(ws, row, col)
            cells[key] = cell
        return cell

    return get_plain_cell


def iter_cols_bulk(
    ws: Worksheet,
    min_col: int | None,
    max_col: int | None,
    min_row: int | None,
    max_row: int | None,
) -> Iterator[tuple[Any, ...]]:
    """Bulk-read column values through one Rust FFI call, then transpose."""
    from wolfxl._cell import _payload_to_python

    reader = ws._workbook._rust_reader  # noqa: SLF001
    sheet = ws._title  # noqa: SLF001
    data_only = getattr(ws._workbook, "_data_only", False)  # noqa: SLF001

    row_min = min_row or 1
    row_max = max_row or ws._max_row()  # noqa: SLF001
    col_min = min_col or 1
    col_max = max_col or ws._max_col()  # noqa: SLF001
    range_str = f"{rowcol_to_a1(row_min, col_min)}:{rowcol_to_a1(row_max, col_max)}"

    use_plain = hasattr(reader, "read_sheet_values_plain")
    if use_plain:
        rows = reader.read_sheet_values_plain(sheet, range_str, data_only)
    else:
        rows = reader.read_sheet_values(sheet, range_str, data_only)

    if not rows:
        return

    expected_cols = col_max - col_min + 1
    expected_rows = row_max - row_min + 1
    if (
        use_plain
        and len(rows) == expected_rows
        and all(isinstance(row, tuple) and len(row) == expected_cols for row in rows)
    ):
        for col_offset in range(expected_cols):
            yield tuple(map(itemgetter(col_offset), rows))
        return

    normalized: list[list[Any]] = []
    for row in rows:
        if use_plain:
            values = list(row)
        else:
            values = [_payload_to_python(cell) for cell in row]
        width = len(values)
        if width >= expected_cols:
            normalized.append(values[:expected_cols])
        else:
            normalized.append(values + [None] * (expected_cols - width))

    while len(normalized) < expected_rows:
        normalized.append([None] * expected_cols)

    for col_offset in range(expected_cols):
        yield tuple(
            normalized[row_offset][col_offset]
            for row_offset in range(expected_rows)
        )


def iter_rows_bulk(
    ws: Worksheet,
    min_row: int | None,
    max_row: int | None,
    min_col: int | None,
    max_col: int | None,
) -> Iterator[tuple[Any, ...]]:
    """Bulk-read row values through one Rust FFI call."""
    from wolfxl._cell import _payload_to_python

    reader = ws._workbook._rust_reader  # noqa: SLF001
    sheet = ws._title  # noqa: SLF001
    data_only = getattr(ws._workbook, "_data_only", False)  # noqa: SLF001

    is_unbounded = (
        min_row is None
        and max_row is None
        and min_col is None
        and max_col is None
    )
    range_str = None
    expected_cols = None
    if not is_unbounded:
        row_min = min_row or 1
        row_max = max_row or ws._max_row()  # noqa: SLF001
        col_min = min_col or 1
        col_max = max_col or ws._max_col()  # noqa: SLF001
        expected_cols = col_max - col_min + 1
        range_str = f"{rowcol_to_a1(row_min, col_min)}:{rowcol_to_a1(row_max, col_max)}"

    use_plain = hasattr(reader, "read_sheet_values_plain")
    if use_plain:
        rows = reader.read_sheet_values_plain(sheet, range_str, data_only)
    else:
        rows = reader.read_sheet_values(sheet, range_str, data_only)

    if not rows:
        return

    if use_plain and expected_cols is None:
        yield from rows
        return

    for row in rows:
        if use_plain:
            width = len(row)
            if width == expected_cols and isinstance(row, tuple):
                yield row
                continue
            values = list(row)
        else:
            values = [_payload_to_python(cell) for cell in row]
            width = len(values)
        if width >= expected_cols:
            yield tuple(values[:expected_cols])
        else:
            yield tuple(values) + (None,) * (expected_cols - width)
