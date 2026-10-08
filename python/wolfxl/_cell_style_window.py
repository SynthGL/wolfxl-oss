"""Bounded coordinate windows for clean eager worksheet style reads."""

from __future__ import annotations

from typing import Any, Callable
from weakref import WeakSet

from wolfxl._utils import rowcol_to_a1


class StylePayloadWindow:
    """Keep one bounded coordinate window while sharing payloads by style ID.

    Worksheet dimensions select window boundaries, never the cache eligibility.
    Sparse sheets with far-away cells therefore keep the same bounded storage
    instead of falling back to one native format conversion per requested cell.
    """

    def __init__(
        self,
        reader: Any,
        title: str,
        max_row: int,
        max_col: int,
        cell_limit: int,
        decode: Callable[[dict[str, Any]], tuple[Any, ...]],
        disabled: object,
        workbook: Any = None,
    ) -> None:
        if workbook is not None:
            windows = getattr(workbook, "_style_payload_windows", None)
            if windows is None:
                workbook._style_payload_windows = windows = WeakSet()
            windows.add(self)
        self._reader = reader
        self._title = title
        self._max_row = max_row
        self._max_col = max_col
        self._cols_per_window = min(max_col, cell_limit)
        self._rows_per_window = max(1, cell_limit // self._cols_per_window)
        self._decode = decode
        self._disabled = disabled
        self._failed = False
        self._start = 0
        self._end = 0
        self._left = 0
        self._right = 0
        self._entries: dict[tuple[int, int], tuple[Any, ...]] = {}
        self._payloads_by_id: dict[int, tuple[Any, ...]] = {}

    def close(self) -> None:
        """Detach the native reader, including from externally retained windows."""
        self._reader = None
        self._failed = True
        self._entries.clear()
        self._payloads_by_id.clear()

    def get(self, key: tuple[int, int], default: Any = None) -> Any:
        """Return a payload, filling the requested row window only when needed."""
        if self._failed:
            return self._disabled
        row, col = key
        if row < 1 or col < 1 or row > self._max_row or col > self._max_col:
            return default
        if not (self._start <= row <= self._end and self._left <= col <= self._right):
            start = ((row - 1) // self._rows_per_window) * self._rows_per_window + 1
            end = min(self._max_row, start + self._rows_per_window - 1)
            left = ((col - 1) // self._cols_per_window) * self._cols_per_window + 1
            right = min(self._max_col, left + self._cols_per_window - 1)
            range_str = f"{rowcol_to_a1(start, left)}:{rowcol_to_a1(end, right)}"
            entries: dict[tuple[int, int], tuple[Any, ...]] = {}
            try:
                for item_row, item_col, style_id in self._reader.read_sheet_style_ids(self._title, range_str):
                    style_id = int(style_id)
                    payload = self._payloads_by_id.get(style_id)
                    if payload is None:
                        payload = self._decode(self._reader.read_format_for_style_id(style_id))
                        self._payloads_by_id[style_id] = payload
                    if payload:
                        entries[(int(item_row), int(item_col))] = payload
            except Exception:
                # A partial native failure must fall back to the authoritative
                # per-cell reader instead of treating missing styles as default.
                self._failed = True
                self._entries.clear()
                return self._disabled
            self._entries = entries
            self._start, self._end = start, end
            self._left, self._right = left, right
        return self._entries.get(key, default)
