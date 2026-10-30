"""Write-mode worksheet cell flush helpers for the native Rust writer."""

from __future__ import annotations

from operator import itemgetter
from typing import TYPE_CHECKING, Any

from wolfxl._utils import rowcol_to_a1
from wolfxl._worksheet_write_buffers import intern_style_grid

if TYPE_CHECKING:
    from wolfxl._worksheet import Worksheet


def flush_to_writer(
    ws: Worksheet,
    writer: Any,
    python_value_to_payload: Any,
    font_to_format_dict: Any,
    fill_to_format_dict: Any,
    alignment_to_format_dict: Any,
    border_to_rust_dict: Any,
    rich_text_to_runs_payload: Any,
    protection_to_format_dict: Any,
) -> None:
    """Flush dirty worksheet cells to the native write-mode backend.

    Args:
        ws: Worksheet whose pending write-mode state should be drained.
        writer: Native workbook writer exposed by the Rust extension.
        python_value_to_payload: Converter for per-cell value payloads.
        font_to_format_dict: Converter for font payload fragments.
        fill_to_format_dict: Converter for fill payload fragments.
        alignment_to_format_dict: Converter for alignment payload fragments.
        border_to_rust_dict: Converter for border payloads.
        rich_text_to_runs_payload: Converter for CellRichText runs.
    """
    _flush_append_buffer(ws, writer, python_value_to_payload)
    _flush_bulk_writes(ws, writer, python_value_to_payload)

    dirty_values = ws._dirty_values  # noqa: SLF001
    dirty = ws._dirty  # noqa: SLF001
    format_keys = ws._format_dirty_cells  # noqa: SLF001
    if dirty_values and dirty <= dirty_values.keys():
        _write_dirty_values(ws, writer, dirty_values)
        format_cells = _format_cell_entries(ws, format_keys) if format_keys else []
    else:
        batch_values, individual_values, format_cells = _partition_dirty_cells(ws)
        _write_batch_values(ws, writer, batch_values)
        _write_individual_values(
            ws,
            writer,
            individual_values,
            python_value_to_payload,
            rich_text_to_runs_payload,
        )
    _flush_spill_child_placeholders(ws, writer)
    _flush_format_cells(
        ws,
        writer,
        format_cells,
        font_to_format_dict,
        fill_to_format_dict,
        alignment_to_format_dict,
        border_to_rust_dict,
        protection_to_format_dict,
    )


def _flush_append_buffer(
    ws: Worksheet,
    writer: Any,
    python_value_to_payload: Any,
) -> None:
    """Flush rows queued via ``append()`` using batch writes where possible."""
    if not ws._append_buffer:  # noqa: SLF001
        return

    buffer = ws._append_buffer  # noqa: SLF001
    start_row = ws._append_buffer_start  # noqa: SLF001
    start_a1 = rowcol_to_a1(start_row, 1)
    individual_values = ws._extract_non_batchable(buffer, start_row, 1)  # noqa: SLF001

    writer.write_sheet_values(ws._title, start_a1, buffer)  # noqa: SLF001

    for row, col, value in individual_values:
        coord = rowcol_to_a1(row, col)
        payload = python_value_to_payload(value)
        writer.write_cell_value(ws._title, coord, payload)  # noqa: SLF001

    ws._append_buffer = []  # noqa: SLF001


def _flush_bulk_writes(
    ws: Worksheet,
    writer: Any,
    python_value_to_payload: Any,
) -> None:
    """Flush rows queued via ``write_rows()`` using batch writes where possible."""
    for grid, start_row, start_col, plain_values_only, style_grid in ws._bulk_writes:  # noqa: SLF001
        start_a1 = rowcol_to_a1(start_row, start_col)
        individual_values = (
            []
            if plain_values_only
            else ws._extract_non_batchable(  # noqa: SLF001
                grid, start_row, start_col
            )
        )
        if style_grid:
            if not (hasattr(writer, "intern_format") and hasattr(writer, "write_sheet_style_ids")):
                raise RuntimeError("write_styled_rows requires native style-id batch support")
            style_id_grid = intern_style_grid(writer, style_grid)
            if not individual_values and hasattr(writer, "write_sheet_values_with_style_ids"):
                writer.write_sheet_values_with_style_ids(  # noqa: SLF001
                    ws._title,
                    start_a1,
                    grid,
                    style_id_grid,
                )
                continue
        else:
            style_id_grid = None
        writer.write_sheet_values(ws._title, start_a1, grid)  # noqa: SLF001
        for row, col, value in individual_values:
            coord = rowcol_to_a1(row, col)
            payload = python_value_to_payload(value)
            writer.write_cell_value(ws._title, coord, payload)  # noqa: SLF001
        if style_id_grid:
            writer.write_sheet_style_ids(ws._title, start_a1, style_id_grid)  # noqa: SLF001

    ws._bulk_writes = []  # noqa: SLF001


def _partition_dirty_cells(
    ws: Worksheet,
) -> tuple[list[tuple[int, int, Any]], list[tuple[int, int, Any]], list[tuple[int, int, Any]]]:
    """Split dirty cells into batchable values, individual values, and formats."""
    from wolfxl._cell import _UNSET

    dirty = ws._dirty  # noqa: SLF001
    dirty_values = ws._dirty_values  # noqa: SLF001
    format_keys = ws._format_dirty_cells  # noqa: SLF001

    # ``cell()``/``ws[...] = v`` record plain values only in ``_dirty_values``,
    # so the value map covers every edit only when ``_dirty`` is a subset of it.
    values_cover_dirty = bool(dirty_values) and dirty <= dirty_values.keys()
    if values_cover_dirty and not format_keys:
        return (
            [(row, col, value) for (row, col), value in dirty_values.items()],
            [],
            [],
        )
    if values_cover_dirty:
        return (
            [(row, col, value) for (row, col), value in dirty_values.items()],
            [],
            _format_cell_entries(ws, format_keys),
        )

    dirty_value_keys = dirty_values.keys()
    batch_values: list[tuple[int, int, Any]] = []
    individual_values: list[tuple[int, int, Any]] = []
    format_cells: list[tuple[int, int, Any]] = []

    for (row, col), value in dirty_values.items():
        batch_values.append((row, col, value))

    for row, col in dirty - dirty_value_keys:
        cell = ws._cells.get((row, col))  # noqa: SLF001
        if cell is None:
            continue

        if cell._value_dirty:  # noqa: SLF001
            value = cell._value  # noqa: SLF001
            if getattr(cell, "_explicit_data_type", None) is not _UNSET:
                individual_values.append((row, col, cell))
            elif value is None or isinstance(value, (bool, int, float, str)):
                batch_values.append((row, col, value))
            else:
                individual_values.append((row, col, cell))

    format_cells.extend(_format_cell_entries(ws, format_keys))

    return batch_values, individual_values, format_cells


def _format_cell_entries(
    ws: Worksheet,
    keys: set[tuple[int, int]],
) -> list[tuple[int, int, Any]]:
    """Return format-dirty cell entries for the supplied worksheet keys."""
    format_cells: list[tuple[int, int, Any]] = []
    cells = ws._cells  # noqa: SLF001
    for key in keys:
        cell = cells.get(key)
        if cell is None:
            continue
        row, col = key
        format_cells.append((row, col, cell))
    return format_cells


def _write_batch_values(
    ws: Worksheet,
    writer: Any,
    batch_values: list[tuple[int, int, Any]],
) -> None:
    """Write simple dirty values as compact rectangular batches."""
    for start_row, start_col, grid in _plan_batch_value_grids(batch_values):
        start = rowcol_to_a1(start_row, start_col)
        writer.write_sheet_values(ws._title, start, grid)  # noqa: SLF001


def _write_dirty_values(
    ws: Worksheet,
    writer: Any,
    dirty_values: dict[tuple[int, int], Any],
) -> None:
    """Write simple dirty-value maps without first building triple records."""
    for start_row, start_col, grid in _plan_dirty_value_grids(dirty_values):
        start = rowcol_to_a1(start_row, start_col)
        writer.write_sheet_values(ws._title, start, grid)  # noqa: SLF001


def _plan_dirty_value_grids(
    dirty_values: dict[tuple[int, int], Any],
) -> list[tuple[int, int, list[list[Any]]]]:
    """Group a dirty-value map into compact rectangular writer grids."""
    if not dirty_values:
        return []

    dense = _plan_dense_dirty_values(dirty_values)
    if dense is not None:
        return [dense]

    batch_values = [
        (row, col, value)
        for (row, col), value in dirty_values.items()
    ]
    return _plan_batch_value_grids(batch_values)


def _plan_dense_dirty_values(
    dirty_values: dict[tuple[int, int], Any],
) -> tuple[int, int, list[list[Any]]] | None:
    """Return one grid when a dirty-value map is already row-major dense."""
    iterator = iter(dirty_values.items())
    try:
        (first_row, first_col), first_value = next(iterator)
    except StopIteration:
        return None

    current_row = first_row
    current_values: list[Any] = [first_value]
    next_col = first_col + 1
    expected_width: int | None = None
    grid: list[list[Any]] = []

    for (row, col), value in iterator:
        if row == current_row:
            if col != next_col:
                return None
            current_values.append(value)
            next_col += 1
            continue

        if row != current_row + 1 or col != first_col:
            return None

        current_width = next_col - first_col
        if expected_width is None:
            expected_width = current_width
        elif current_width != expected_width:
            return None

        grid.append(current_values)
        current_row = row
        current_values = [value]
        next_col = first_col + 1

    if expected_width is None:
        expected_width = next_col - first_col
    if next_col - first_col != expected_width:
        return None

    grid.append(current_values)
    return first_row, first_col, grid


def _plan_batch_value_grids(
    batch_values: list[tuple[int, int, Any]],
) -> list[tuple[int, int, list[list[Any]]]]:
    """Group dirty values into dense row-contiguous grids.

    A single bounding box can turn two far-apart cells into a huge mostly empty
    grid. This planner keeps only contiguous column runs, then stacks adjacent
    rows when their runs have the same start column and width.
    """
    if not batch_values:
        return []

    dense = _plan_dense_rectangle(batch_values)
    if dense is not None:
        return [dense]

    ordered = sorted(batch_values, key=itemgetter(0, 1))
    dense = _plan_dense_rectangle(ordered)
    if dense is not None:
        return [dense]

    return _plan_sparse_value_grids(ordered)


def _plan_dense_rectangle(
    ordered_values: list[tuple[int, int, Any]],
) -> tuple[int, int, list[list[Any]]] | None:
    """Return one grid when dirty values form a compact rectangle."""
    first_row, first_col, _value = ordered_values[0]
    current_row = first_row
    current_values: list[Any] = []
    expected_width: int | None = None
    grid: list[list[Any]] = []

    for row, col, value in ordered_values:
        if row == current_row:
            if col != first_col + len(current_values):
                return None
            current_values.append(value)
            continue

        if row != current_row + 1 or col != first_col:
            return None

        if expected_width is None:
            expected_width = len(current_values)
        elif len(current_values) != expected_width:
            return None

        grid.append(current_values)
        current_row = row
        current_values = [value]

    if expected_width is None:
        expected_width = len(current_values)
    if len(current_values) != expected_width:
        return None

    grid.append(current_values)
    return first_row, first_col, grid


def _plan_sparse_value_grids(
    ordered_values: list[tuple[int, int, Any]],
) -> list[tuple[int, int, list[list[Any]]]]:
    """Group sorted dirty values into compact sparse grids."""
    row_runs: list[tuple[int, int, list[Any]]] = []
    current_row: int | None = None
    current_start_col: int | None = None
    current_next_col: int | None = None
    current_values: list[Any] = []

    for row, col, value in ordered_values:
        if (
            current_row == row
            and current_next_col is not None
            and col == current_next_col
        ):
            current_values.append(value)
            current_next_col = col + 1
            continue

        if current_row is not None and current_start_col is not None:
            row_runs.append((current_row, current_start_col, current_values))

        current_row = row
        current_start_col = col
        current_next_col = col + 1
        current_values = [value]

    if current_row is not None and current_start_col is not None:
        row_runs.append((current_row, current_start_col, current_values))

    runs_by_signature: dict[tuple[int, int], list[tuple[int, list[Any]]]] = {}
    for row, start_col, values in row_runs:
        runs_by_signature.setdefault((start_col, len(values)), []).append((row, values))

    grids: list[tuple[int, int, list[list[Any]]]] = []
    for (start_col, _width), runs in runs_by_signature.items():
        active_start_row: int | None = None
        active_row: int | None = None
        active_grid: list[list[Any]] = []

        for row, values in sorted(runs, key=itemgetter(0)):
            if active_row is not None and row == active_row + 1:
                active_grid.append(values)
                active_row = row
                continue

            if active_start_row is not None:
                grids.append((active_start_row, start_col, active_grid))

            active_start_row = row
            active_row = row
            active_grid = [values]

        if active_start_row is not None:
            grids.append((active_start_row, start_col, active_grid))

    return sorted(grids, key=itemgetter(0, 1))


def _write_individual_values(
    ws: Worksheet,
    writer: Any,
    individual_values: list[tuple[int, int, Any]],
    python_value_to_payload: Any,
    rich_text_to_runs_payload: Any,
) -> None:
    """Write non-batchable dirty values with type-preserving writer calls."""
    from wolfxl.cell.cell import ArrayFormula, DataTableFormula
    from wolfxl.cell.rich_text import CellRichText

    for _row, _col, cell in individual_values:
        coord = rowcol_to_a1(cell._row, cell._col)  # noqa: SLF001
        value = cell._value  # noqa: SLF001
        if isinstance(value, ArrayFormula):
            if hasattr(writer, "write_cell_array_formula"):
                writer.write_cell_array_formula(
                    ws._title,  # noqa: SLF001
                    coord,
                    {"kind": "array", "ref": value.ref, "text": value.text},
                )
            else:
                payload = python_value_to_payload(f"={value.text}")
                writer.write_cell_value(ws._title, coord, payload)  # noqa: SLF001
            continue
        if isinstance(value, DataTableFormula):
            if hasattr(writer, "write_cell_array_formula"):
                writer.write_cell_array_formula(
                    ws._title,  # noqa: SLF001
                    coord,
                    {
                        "kind": "data_table",
                        "ref": value.ref,
                        "ca": value.ca,
                        "dt2D": value.dt2D,
                        "dtr": value.dtr,
                        "r1": value.r1,
                        "r2": value.r2,
                        "del1": value.del1,
                        "del2": value.del2,
                    },
                )
            continue
        if isinstance(value, CellRichText):
            if hasattr(writer, "write_cell_rich_text"):
                runs_payload = rich_text_to_runs_payload(value)
                writer.write_cell_rich_text(ws._title, coord, runs_payload)  # noqa: SLF001
            else:
                payload = python_value_to_payload(str(value))
                writer.write_cell_value(ws._title, coord, payload)  # noqa: SLF001
            continue

        payload = _payload_for_cell(cell, value, python_value_to_payload)
        writer.write_cell_value(ws._title, coord, payload)  # noqa: SLF001


def _payload_for_cell(
    cell: Any,
    value: Any,
    python_value_to_payload: Any,
) -> dict[str, Any]:
    """Build a value payload, honoring explicit openpyxl ``Cell.data_type``."""
    from wolfxl._cell import _UNSET

    explicit = getattr(cell, "_explicit_data_type", _UNSET)
    if explicit == "s":
        return {"type": "string", "value": "" if value is None else str(value)}
    if explicit == "e":
        return {"type": "error", "value": "" if value is None else str(value)}
    if explicit == "f":
        formula = "" if value is None else str(value)
        if formula and not formula.startswith("="):
            formula = f"={formula}"
        return {"type": "formula", "formula": formula, "value": formula}
    if explicit == "b":
        return {"type": "boolean", "value": bool(value)}
    if explicit == "n":
        return {"type": "number", "value": value}
    return python_value_to_payload(value)


def _flush_spill_child_placeholders(ws: Worksheet, writer: Any) -> None:
    """Flush placeholder cells for array-formula spill ranges."""
    for (row, col), (kind, _payload) in ws._pending_array_formulas.items():  # noqa: SLF001
        if kind != "spill_child":
            continue
        if (row, col) in ws._dirty:  # noqa: SLF001
            continue
        coord = rowcol_to_a1(row, col)
        if hasattr(writer, "write_cell_array_formula"):
            writer.write_cell_array_formula(
                ws._title,  # noqa: SLF001
                coord,
                {"kind": "spill_child"},
            )


def _flush_format_cells(
    ws: Worksheet,
    writer: Any,
    format_cells: list[tuple[int, int, Any]],
    font_to_format_dict: Any,
    fill_to_format_dict: Any,
    alignment_to_format_dict: Any,
    border_to_rust_dict: Any,
    protection_to_format_dict: Any,
) -> None:
    """Flush dirty cell format and border payloads.

    Borders are merged into the same dict as font / fill / alignment so
    each cell receives exactly one ``write_cell_format`` call. Splitting
    border into a separate ``write_cell_border`` call would mint a new
    style_id that overwrites the format-only one (RFC-064 follow-up:
    native writer interns eagerly per call, so two calls = two ids).
    """
    from wolfxl._cell import _UNSET

    if not format_cells:
        return

    format_entries: list[tuple[int, int, dict[str, Any]]] = []
    style_id_entries: list[tuple[int, int, int]] = []
    font_cache: dict[Any, dict[str, Any]] = {}
    fill_cache: dict[Any, dict[str, Any]] = {}
    alignment_cache: dict[Any, dict[str, Any]] = {}
    protection_cache: dict[Any, dict[str, Any]] = {}
    border_cache: dict[Any, dict[str, Any]] = {}
    single_style_id_cache: dict[tuple[str, Any], int] = {}
    can_write_style_ids = hasattr(writer, "intern_format") and hasattr(
        writer,
        "write_sheet_style_ids",
    )

    for row, col, cell in format_cells:
        font = cell._font  # noqa: SLF001
        fill = cell._fill  # noqa: SLF001
        alignment = cell._alignment  # noqa: SLF001
        number_format = cell._number_format  # noqa: SLF001
        protection = cell._protection  # noqa: SLF001
        cell_border = cell._border  # noqa: SLF001
        named_style = cell._named_style  # noqa: SLF001

        has_named_style = (
            named_style is not _UNSET
            and named_style is not None
            and named_style != "Normal"
        )
        no_uncommon_styles = (
            (alignment is _UNSET or alignment is None)
            and (protection is _UNSET or protection is None)
            and (cell_border is _UNSET or cell_border is None)
            and not has_named_style
        )
        has_font = font is not _UNSET and font is not None
        has_fill = fill is not _UNSET and fill is not None
        has_number_format = number_format is not _UNSET and number_format is not None

        if no_uncommon_styles:
            if has_font and not has_fill and not has_number_format:
                style_key = _style_cache_key(font)
                fmt = _cached_format_payload(
                    font_cache, font, font_to_format_dict, style_key
                )
                if fmt:
                    style_id = _cached_single_style_id(
                        single_style_id_cache,
                        writer,
                        "font",
                        font,
                        fmt,
                        can_write_style_ids,
                        style_key,
                    )
                    if style_id is None:
                        format_entries.append((row, col, fmt.copy()))
                    else:
                        style_id_entries.append((row, col, style_id))
                continue
            if has_fill and not has_font and not has_number_format:
                style_key = _style_cache_key(fill)
                fmt = _cached_format_payload(
                    fill_cache, fill, fill_to_format_dict, style_key
                )
                if fmt:
                    style_id = _cached_single_style_id(
                        single_style_id_cache,
                        writer,
                        "fill",
                        fill,
                        fmt,
                        can_write_style_ids,
                        style_key,
                    )
                    if style_id is None:
                        format_entries.append((row, col, fmt.copy()))
                    else:
                        style_id_entries.append((row, col, style_id))
                continue
            if has_number_format and not has_font and not has_fill:
                fmt = {"number_format": number_format}
                style_id = _cached_single_style_id(
                    single_style_id_cache,
                    writer,
                    "number_format",
                    number_format,
                    fmt,
                    can_write_style_ids,
                )
                if style_id is None:
                    format_entries.append((row, col, fmt))
                else:
                    style_id_entries.append((row, col, style_id))
                continue

        fmt: dict[str, Any] = {}

        if has_font:
            fmt.update(_cached_format_payload(font_cache, font, font_to_format_dict))
        if has_fill:
            fmt.update(_cached_format_payload(fill_cache, fill, fill_to_format_dict))
        if alignment is not _UNSET and alignment is not None:
            fmt.update(
                _cached_format_payload(
                    alignment_cache,
                    alignment,
                    alignment_to_format_dict,
                )
            )
        if has_number_format:
            fmt["number_format"] = number_format
        if protection is not _UNSET and protection is not None:
            fmt.update(
                _cached_format_payload(
                    protection_cache,
                    protection,
                    protection_to_format_dict,
                )
            )
        if cell_border is not _UNSET and cell_border is not None:
            border = _cached_format_payload(
                border_cache,
                cell_border,
                border_to_rust_dict,
            )
            if border:
                fmt.update(border)
        if has_named_style:
            # Threaded into the same dict so the native writer can stamp
            # the resulting xf record with the correct cellStyleXfs slot
            # (xfId attr). Cells using Normal/None stay on slot 0 implicitly.
            fmt["_named_style"] = named_style

        if fmt:
            format_entries.append((row, col, fmt))

    if style_id_entries:
        ws._batch_write_style_ids(writer.write_sheet_style_ids, style_id_entries)  # noqa: SLF001
    if len(format_entries) > 1:
        ws._batch_write_dicts(writer.write_sheet_formats, format_entries)  # noqa: SLF001
    else:
        for row, col, fmt in format_entries:
            writer.write_cell_format(ws._title, rowcol_to_a1(row, col), fmt)  # noqa: SLF001


def _cached_format_payload(
    cache: dict[Any, dict[str, Any]],
    value: Any,
    converter: Any,
    key: Any | None = None,
) -> dict[str, Any]:
    """Convert a style object once per worksheet flush when it is hashable."""
    if key is None:
        key = _style_cache_key(value)
    try:
        cached = cache.get(key)
    except TypeError:
        return converter(value)
    if cached is None:
        cached = converter(value)
        cache[key] = cached
    return cached


def _cached_single_style_id(
    cache: dict[tuple[str, Any], int],
    writer: Any,
    kind: str,
    value: Any,
    format_payload: dict[str, Any],
    enabled: bool,
    value_key: Any | None = None,
) -> int | None:
    """Intern one-field style payloads once per flush when the backend supports it."""
    if not enabled:
        return None
    try:
        key = (kind, _style_cache_key(value) if value_key is None else value_key)
        cached = cache.get(key)
    except TypeError:
        return None
    if cached is None:
        cached = int(writer.intern_format(format_payload))
        cache[key] = cached
    return cached


def _style_cache_key(value: Any) -> Any:
    """Return a cheap equality key for common style objects."""
    cls_name = type(value).__name__
    if cls_name == "Font":
        return (
            "Font",
            value.name,
            value.size,
            value.bold,
            value.italic,
            value.underline,
            value.strike,
            _style_color_key(value.color),
            value.family,
            value.charset,
            value.scheme,
            value.vertAlign,
            value.outline,
            value.shadow,
            value.condense,
            value.extend,
        )
    if cls_name == "PatternFill":
        return (
            "PatternFill",
            value.patternType,
            _style_color_key(value.fgColor),
            _style_color_key(value.bgColor),
        )
    return value


def _style_color_key(value: Any) -> Any:
    """Return a hashable key for style color objects."""
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    return (
        type(value).__name__,
        getattr(value, "type", None),
        getattr(value, "rgb", None),
        getattr(value, "indexed", None),
        getattr(value, "theme", None),
        getattr(value, "tint", None),
        getattr(value, "auto", None),
    )
