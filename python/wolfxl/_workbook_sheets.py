"""Workbook sheet creation, copying, removal, and ordering helpers."""

from __future__ import annotations

import re
from typing import Any

from wolfxl._worksheet import Worksheet


def remove_sheet(wb: Any, worksheet: Worksheet) -> None:
    """Remove ``worksheet`` from a write or modify-mode workbook.

    Args:
        wb: Workbook-like object carrying writer and sheet state.
        worksheet: Worksheet object to remove.

    Raises:
        RuntimeError: If ``wb`` is not in write mode.
        ValueError: If ``worksheet`` is not registered on ``wb``.
    """
    if wb._rust_writer is None and wb._rust_patcher is None:  # noqa: SLF001
        raise RuntimeError("remove requires write mode or modify mode")
    if not isinstance(worksheet, Worksheet):
        raise ValueError(f"{worksheet!r} is not in list")
    if worksheet.title not in wb._sheets:  # noqa: SLF001
        raise ValueError(f"Worksheet '{worksheet.title}' is not in this workbook")
    title = worksheet.title
    wb._sheet_names.remove(title)  # noqa: SLF001
    wb._sheets.pop(title)  # noqa: SLF001
    if wb._rust_patcher is not None:  # noqa: SLF001
        wb._rust_patcher.queue_sheet_delete(title)  # noqa: SLF001
        return
    wb._active_sheet_dirty = True  # noqa: SLF001
    wb._rust_writer.remove_sheet(title)  # noqa: SLF001


def create_sheet(wb: Any, title: str | None, index: int | None = None) -> Worksheet:
    """Create and append a worksheet in write or modify mode.

    Args:
        wb: Workbook-like object carrying writer and sheet state.
        title: Worksheet title. ``None``/``""`` and duplicate names are
            resolved with openpyxl-style numeric suffixes.

    Returns:
        The newly created worksheet.

    Raises:
        RuntimeError: If ``wb`` is not in write mode.
    """
    title = _unique_sheet_title(wb, title)
    insert_at = _normalize_insert_index(wb, index)
    if wb._rust_patcher is None and wb._rust_writer is None:  # noqa: SLF001
        _promote_path_backed_read_workbook(wb)
    if wb._rust_patcher is not None:  # noqa: SLF001
        wb._rust_patcher.queue_sheet_create(title, insert_at)  # noqa: SLF001
        wb._sheet_names.insert(insert_at, title)  # noqa: SLF001
        ws = Worksheet(wb, title)
        wb._sheets[title] = ws  # noqa: SLF001
        return ws
    if wb._rust_writer is None:  # noqa: SLF001
        raise RuntimeError("create_sheet requires write or modify mode")
    # G20: streaming write-only mode dispatches to a different sheet
    # type — the eager Worksheet is materialisation-heavy, while
    # WriteOnlyWorksheet streams rows straight to a temp file.
    if getattr(wb, "_write_only", False):
        from wolfxl._worksheet_write_only import WriteOnlyWorksheet

        wb._rust_writer.add_sheet(title)  # noqa: SLF001
        wb._sheet_names.insert(insert_at, title)  # noqa: SLF001
        ws = WriteOnlyWorksheet(wb, title)
        wb._sheets[title] = ws  # noqa: SLF001
        _move_new_writer_sheet_to_index(wb, title, insert_at)
        return ws  # type: ignore[return-value]
    wb._rust_writer.add_sheet(title)  # noqa: SLF001
    wb._sheet_names.insert(insert_at, title)  # noqa: SLF001
    ws = Worksheet(wb, title)
    wb._sheets[title] = ws  # noqa: SLF001
    _move_new_writer_sheet_to_index(wb, title, insert_at)
    return ws


def _promote_path_backed_read_workbook(wb: Any) -> None:
    """Enable lazy openpyxl-style editing for a loaded XLSX workbook."""
    source_path = getattr(wb, "_source_path", None)
    if source_path is None or getattr(wb, "_read_only", False):
        return
    from wolfxl import _rust

    wb._rust_patcher = _rust.XlsxPatcher.open(source_path, False)  # noqa: SLF001


def _unique_sheet_title(wb: Any, title: str | None) -> str:
    candidate = title or "Sheet"
    names = [str(name) for name in wb._sheet_names]  # noqa: SLF001
    if not any(name.lower() == candidate.lower() for name in names):
        return candidate

    joined_names = ",".join(names)
    title_regex = re.compile(f"(?P<title>{re.escape(candidate)})(?P<count>\\d*),?", re.I)
    counts = [
        int(match.group("count"))
        for match in title_regex.finditer(joined_names)
        if match.group("count").isdigit()
    ]
    return f"{candidate}{max(counts) + 1 if counts else 1}"


def _normalize_insert_index(wb: Any, index: int | None) -> int:
    if index is None:
        return len(wb._sheet_names)  # noqa: SLF001
    if not isinstance(index, int):
        raise TypeError("create_sheet index must be an int")
    if index < 0:
        return max(len(wb._sheet_names) + index, 0)  # noqa: SLF001
    return min(index, len(wb._sheet_names))  # noqa: SLF001


def _move_new_writer_sheet_to_index(wb: Any, title: str, index: int) -> None:
    current = len(wb._sheet_names) - 1  # noqa: SLF001
    if index == current:
        return
    move = getattr(wb._rust_writer, "move_sheet", None)  # noqa: SLF001
    if move is not None:
        move(title, index - current)


def copy_worksheet(
    wb: Any,
    source: Worksheet,
    *,
    name: str | None = None,
) -> Worksheet:
    """Duplicate ``source`` into a new workbook sheet.

    Args:
        wb: Workbook-like object carrying writer, patcher, and sheet state.
        source: Source worksheet to copy.
        name: Optional explicit destination title.

    Returns:
        The newly created worksheet proxy.

    Raises:
        TypeError: If ``source`` is not a worksheet.
        ValueError: If the source belongs to another workbook or the title
            collides.
        RuntimeError: If the workbook cannot copy in its current mode.
    """
    if not isinstance(source, Worksheet):
        raise TypeError(
            f"copy_worksheet: source must be a Worksheet, got {type(source).__name__}"
        )
    if source._workbook is not wb:  # noqa: SLF001
        raise ValueError("copy_worksheet: source must belong to this workbook")
    if wb._rust_patcher is None and wb._rust_writer is None:  # noqa: SLF001
        raise RuntimeError("copy_worksheet requires write or modify mode")

    new_title = _new_copy_title(wb, source, name=name)
    if wb._rust_patcher is not None:  # noqa: SLF001
        return _copy_worksheet_modify_mode(wb, source, new_title)
    return copy_worksheet_write_mode(wb, source, new_title)


def copy_worksheet_write_mode(
    wb: Any,
    source: Worksheet,
    new_title: str,
) -> Worksheet:
    """Clone an in-memory worksheet into a fresh write-mode sheet.

    Args:
        wb: Workbook-like object carrying writer and sheet state.
        source: Source worksheet to clone.
        new_title: Destination worksheet title, already validated.

    Returns:
        The newly created destination worksheet.
    """
    _materialize_copy_source(source)
    dst = create_sheet(wb, new_title)
    _copy_cells_write_mode(source, dst)
    _copy_layout_state(source, dst)
    _copy_autofilter_state(source, dst)
    _copy_sheet_setup_state(source, dst)
    _copy_print_titles(source, dst)
    return dst


def _materialize_copy_source(source: Worksheet) -> None:
    """Materialize pending write buffers before copying a worksheet."""
    if source._append_buffer:  # noqa: SLF001
        source._materialize_append_buffer()  # noqa: SLF001
    if source._bulk_writes:  # noqa: SLF001
        source._materialize_bulk_writes()  # noqa: SLF001


def _copy_cells_write_mode(source: Worksheet, dst: Worksheet) -> None:
    """Replay dirty write-mode cells from ``source`` onto ``dst``."""
    from wolfxl._cell import _UNSET

    for (row, col), src_cell in source._cells.items():  # noqa: SLF001
        value = src_cell._value  # noqa: SLF001
        has_value = value is not _UNSET and src_cell._value_dirty  # noqa: SLF001
        font = getattr(src_cell, "_font", _UNSET)
        fill = getattr(src_cell, "_fill", _UNSET)
        border = getattr(src_cell, "_border", _UNSET)
        alignment = getattr(src_cell, "_alignment", _UNSET)
        number_format = getattr(src_cell, "_number_format", _UNSET)
        has_format = src_cell._format_dirty  # noqa: SLF001

        if not has_value and not has_format:
            continue

        dst_cell = dst.cell(row=row, column=col)
        if has_value:
            dst_cell.value = value
        if font is not _UNSET:
            dst_cell.font = font  # type: ignore[assignment]
        if fill is not _UNSET:
            dst_cell.fill = fill  # type: ignore[assignment]
        if border is not _UNSET:
            dst_cell.border = border  # type: ignore[assignment]
        if alignment is not _UNSET:
            dst_cell.alignment = alignment  # type: ignore[assignment]
        if number_format is not _UNSET:
            dst_cell.number_format = number_format  # type: ignore[assignment]


def _copy_layout_state(source: Worksheet, dst: Worksheet) -> None:
    """Copy row, column, merge, freeze-pane, and print-area state."""
    for r, h in source._row_heights.items():  # noqa: SLF001
        dst._row_heights[r] = h  # noqa: SLF001
    dst._row_dimensions.update(source._row_dimensions)  # noqa: SLF001
    dst._row_hidden.update(source._row_hidden)  # noqa: SLF001
    dst._row_outline_levels.update(source._row_outline_levels)  # noqa: SLF001
    for letter, w in source._col_widths.items():  # noqa: SLF001
        dst._col_widths[letter] = w  # noqa: SLF001
    dst._col_dimensions.update(source._col_dimensions)  # noqa: SLF001
    dst._col_hidden.update(source._col_hidden)  # noqa: SLF001
    dst._col_outline_levels.update(source._col_outline_levels)  # noqa: SLF001
    for rng in source._merged_ranges:  # noqa: SLF001
        dst.merge_cells(rng)
    if source._freeze_panes is not None:  # noqa: SLF001
        dst._freeze_panes = source._freeze_panes  # noqa: SLF001
    if source._print_area is not None:  # noqa: SLF001
        dst._print_area = source._print_area  # noqa: SLF001


def move_sheet(wb: Any, sheet: Worksheet | str, offset: int = 0) -> None:
    """Move a sheet by ``offset`` positions within workbook tab order.

    Args:
        wb: Workbook-like object carrying writer, patcher, and sheet state.
        sheet: Worksheet instance or sheet title.
        offset: Integer count of positions to shift.

    Raises:
        TypeError: If ``sheet`` or ``offset`` has the wrong type.
        KeyError: If the sheet is not present.
    """
    name = _resolve_sheet_name(sheet)
    _validate_sheet_offset(offset)
    _move_sheet_in_memory(wb, name, offset)

    if wb._rust_patcher is not None:  # noqa: SLF001
        wb._flush_pending_sheet_moves_to_patcher(name, offset)  # noqa: SLF001
    if wb._rust_writer is not None:  # noqa: SLF001
        wb._rust_writer.move_sheet(name, offset)  # noqa: SLF001


def _resolve_sheet_name(sheet: Worksheet | str) -> str:
    """Return a sheet title from a worksheet object or title string."""
    if isinstance(sheet, Worksheet):
        return sheet.title
    if isinstance(sheet, str):
        return sheet
    raise TypeError(
        f"move_sheet: 'sheet' must be a Worksheet or str, got {type(sheet).__name__}"
    )


def _validate_sheet_offset(offset: int) -> None:
    """Validate the ``Workbook.move_sheet`` offset argument."""
    if isinstance(offset, bool) or not isinstance(offset, int):
        raise TypeError(
            f"move_sheet: 'offset' must be an int, got {type(offset).__name__}"
        )


def _move_sheet_in_memory(wb: Any, name: str, offset: int) -> None:
    """Move one sheet title within ``wb``'s in-memory tab order."""
    if name not in wb._sheet_names:  # noqa: SLF001
        raise KeyError(name)

    n = len(wb._sheet_names)  # noqa: SLF001
    idx = wb._sheet_names.index(name)  # noqa: SLF001
    new_pos = max(0, min(idx + offset, n - 1))

    del wb._sheet_names[idx]  # noqa: SLF001
    wb._sheet_names.insert(new_pos, name)  # noqa: SLF001


def _new_copy_title(
    wb: Any,
    source: Worksheet,
    *,
    name: str | None,
) -> str:
    """Return a collision-free copy title, or validate the explicit name."""
    if name is not None:
        if not isinstance(name, str) or not name:
            raise ValueError("copy_worksheet: name must be a non-empty string")
        if name in wb._sheets:  # noqa: SLF001
            raise ValueError(f"copy_worksheet: sheet '{name}' already exists")
        return name

    base = f"{source.title} Copy"
    new_title = base
    suffix = 2
    while new_title in wb._sheets:  # noqa: SLF001
        new_title = f"{base} {suffix}"
        suffix += 1
    return new_title


def _copy_worksheet_modify_mode(
    wb: Any,
    source: Worksheet,
    new_title: str,
) -> Worksheet:
    """Queue a modify-mode worksheet copy and return the destination proxy."""
    wb._pending_sheet_copies.append(  # noqa: SLF001
        (source.title, new_title, bool(wb.copy_options.deep_copy_images))
    )
    wb._sheet_names.append(new_title)  # noqa: SLF001
    ws = Worksheet(wb, new_title)
    wb._sheets[new_title] = ws  # noqa: SLF001
    _copy_sheet_setup_state(source, ws)
    _copy_print_titles(source, ws)
    _copy_pending_charts(source, ws)
    return ws


def _copy_autofilter_state(source: Worksheet, dst: Worksheet) -> None:
    """Deep-copy source autofilter state onto ``dst`` when configured."""
    src_af = source._auto_filter  # noqa: SLF001
    if (
        src_af.ref is None
        and not src_af.filter_columns
        and src_af.sort_state is None
    ):
        return

    import copy as _copy

    dst._auto_filter._ref = src_af.ref  # noqa: SLF001
    dst._auto_filter.filter_columns = _copy.deepcopy(src_af.filter_columns)  # noqa: SLF001
    dst._auto_filter.sort_state = _copy.deepcopy(src_af.sort_state)  # noqa: SLF001


def _copy_sheet_setup_state(source: Worksheet, dst: Worksheet) -> None:
    """Deep-copy mutable sheet setup slots from ``source`` to ``dst``."""
    import copy as _copy

    for slot in (
        "_page_setup",
        "_page_margins",
        "_print_options",
        "_header_footer",
        "_sheet_view",
        "_protection",
        "_row_breaks",
        "_col_breaks",
        "_sheet_format",
    ):
        src_v = getattr(source, slot, None)
        if src_v is not None:
            setattr(dst, slot, _copy.deepcopy(src_v))


def _copy_print_titles(source: Worksheet, dst: Worksheet) -> None:
    """Copy print-title selectors from ``source`` to ``dst``."""
    if getattr(source, "_print_title_rows", None) is not None:
        dst._print_title_rows = source._print_title_rows  # noqa: SLF001
    if getattr(source, "_print_title_cols", None) is not None:
        dst._print_title_cols = source._print_title_cols  # noqa: SLF001


def _copy_pending_charts(source: Worksheet, dst: Worksheet) -> None:
    """Clone high-level charts queued on the source in the same save."""
    pending = getattr(source, "_pending_charts", None)
    if not pending:
        return
    import copy as _copy

    dst._pending_charts.extend(_copy.deepcopy(pending))  # noqa: SLF001
