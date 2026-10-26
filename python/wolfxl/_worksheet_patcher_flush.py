"""Modify-mode worksheet cell flush helpers for the Rust patcher."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from wolfxl._utils import rowcol_to_a1
from wolfxl.utils.datetime import to_excel

if TYPE_CHECKING:
    from wolfxl._worksheet import Worksheet


def flush_to_patcher(
    ws: Worksheet,
    patcher: Any,
    python_value_to_payload: Any,
    font_to_format_dict: Any,
    fill_to_format_dict: Any,
    alignment_to_format_dict: Any,
    border_to_rust_dict: Any,
    rich_text_to_runs_payload: Any,
    protection_to_format_dict: Any,
) -> None:
    """Flush dirty worksheet cells to the ``XlsxPatcher`` backend."""
    from wolfxl._cell import _UNSET
    from wolfxl.cell.cell import ArrayFormula, DataTableFormula
    from wolfxl.cell.rich_text import CellRichText

    spill_children: set[tuple[int, int]] = {
        key
        for key, (kind, _payload) in ws._pending_array_formulas.items()  # noqa: SLF001
        if kind == "spill_child" and key not in ws._dirty  # noqa: SLF001
    }

    dirty_values = ws._dirty_values  # noqa: SLF001
    dirty_value_keys = set(dirty_values)
    for row, col in sorted(dirty_value_keys):
        coord = rowcol_to_a1(row, col)
        value = dirty_values[(row, col)]
        patcher.queue_value(  # noqa: SLF001
            ws._title,  # noqa: SLF001
            coord,
            _normalize_patcher_value_payload(ws, value, python_value_to_payload(value)),
        )

    for row, col in ws._dirty:  # noqa: SLF001
        cell = ws._cells.get((row, col))  # noqa: SLF001
        if cell is None:
            continue
        coord = rowcol_to_a1(row, col)

        if cell._value_dirty and (row, col) not in dirty_value_keys:  # noqa: SLF001
            value = cell._value  # noqa: SLF001
            if isinstance(value, ArrayFormula):
                patcher.queue_array_formula(
                    ws._title,  # noqa: SLF001
                    coord,
                    {"kind": "array", "ref": value.ref, "text": value.text},
                )
            elif isinstance(value, DataTableFormula):
                patcher.queue_array_formula(
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
            elif isinstance(value, CellRichText):
                runs_payload = rich_text_to_runs_payload(value)
                patcher.queue_rich_text_value(ws._title, coord, runs_payload)  # noqa: SLF001
            else:
                payload = _payload_for_cell(cell, value, python_value_to_payload)
                patcher.queue_value(  # noqa: SLF001
                    ws._title,  # noqa: SLF001
                    coord,
                    _normalize_patcher_value_payload(ws, value, payload),
                )

        if cell._format_dirty:  # noqa: SLF001
            fmt: dict[str, Any] = {}

            font = getattr(cell, "_font", _UNSET)
            fill = getattr(cell, "_fill", _UNSET)
            alignment = getattr(cell, "_alignment", _UNSET)
            number_format = getattr(cell, "_number_format", _UNSET)
            protection = getattr(cell, "_protection", _UNSET)
            if font is not _UNSET and font is not None:
                fmt.update(font_to_format_dict(font))
            if fill is not _UNSET and fill is not None:
                fmt.update(fill_to_format_dict(fill))
            if alignment is not _UNSET and alignment is not None:
                fmt.update(alignment_to_format_dict(alignment))
            if number_format is not _UNSET and number_format is not None:
                fmt["number_format"] = number_format
            if protection is not _UNSET and protection is not None:
                fmt.update(protection_to_format_dict(protection))

            if fmt:
                patcher.queue_format(ws._title, coord, fmt)  # noqa: SLF001

            cell_border = getattr(cell, "_border", _UNSET)
            if cell_border is not _UNSET and cell_border is not None:
                border = border_to_rust_dict(cell_border)
                if border:
                    patcher.queue_border(ws._title, coord, border)  # noqa: SLF001

    for row, col in spill_children:
        coord = rowcol_to_a1(row, col)
        patcher.queue_array_formula(
            ws._title,  # noqa: SLF001
            coord,
            {"kind": "spill_child"},
        )


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


def _normalize_patcher_value_payload(
    ws: Worksheet,
    value: Any,
    payload: dict[str, Any],
) -> dict[str, Any]:
    """Convert naive date/time payloads to Excel serials for ``XlsxPatcher``."""
    if payload.get("type") in {"date", "datetime", "time"}:
        if getattr(value, "tzinfo", None) is not None:
            raise TypeError(
                "Excel does not support timezones in datetimes. "
                "The tzinfo in the datetime/time object must be set to None."
            )
        return {"type": "number", "value": to_excel(value, ws._workbook.epoch)}  # noqa: SLF001
    return payload
