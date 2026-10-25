"""RFC-059 §2.3 (Sprint Ο Pod-1E): MergedCell + WriteOnlyCell.

Pins the cell-class compatibility shims so user code that does
``isinstance(cell, MergedCell)`` to detect non-anchor positions
inside merged ranges, or constructs ``WriteOnlyCell(value=42)``
to hand to ``ws.append([...])``, can migrate to wolfxl with a
one-line import swap.
"""

from __future__ import annotations

import pytest

from wolfxl._cell import Cell, _is_merged_subordinate
from wolfxl import Workbook
from wolfxl.cell import MergedCell, WriteOnlyCell
from wolfxl.cell._merged import MergedCell as MergedCellDirect
from wolfxl.cell._write_only import WriteOnlyCell as WriteOnlyCellDirect
from wolfxl.styles import PatternFill


def test_plain_prefilled_value_skips_formula_metadata(monkeypatch: pytest.MonkeyPatch) -> None:
    """Bulk-prefilled scalar values should return without formula lookups."""

    class Worksheet:
        pass

    def fail_formula_lookup(self: Cell) -> object:
        raise AssertionError("plain values should not check formula metadata")

    cell = Cell(Worksheet(), 1, 1)  # type: ignore[arg-type]
    cell._value = 42  # noqa: SLF001
    cell._value_is_plain = True  # noqa: SLF001

    monkeypatch.setattr(Cell, "_value_from_pending_formula", fail_formula_lookup)
    monkeypatch.setattr(Cell, "_value_from_formula_metadata", fail_formula_lookup)

    assert cell.value == 42


def test_plain_scalar_fast_cell_initializes_style_slots_on_read() -> None:
    """Simple fast-path writes should still expose normal style objects."""

    wb = Workbook()
    ws = wb.active
    ws["A1"] = 4

    cell = ws["A1"]

    assert cell.font is not None
    assert cell.fill is not None
    assert cell.border is not None
    assert cell.alignment is not None
    assert cell.number_format == "General"
    assert cell.protection is not None
    assert cell._style is not None  # noqa: SLF001


# ---------------------------------------------------------------------------
# MergedCell
# ---------------------------------------------------------------------------


def test_merged_cell_value_is_none() -> None:
    mc = MergedCell(parent=None, row=2, column=3)
    assert mc.value is None


def test_merged_cell_setter_raises_attribute_error() -> None:
    mc = MergedCell(parent=None, row=2, column=3)
    with pytest.raises(AttributeError, match="merged range"):
        mc.value = "anything"


def test_merged_cell_coordinate() -> None:
    mc = MergedCell(parent=None, row=3, column=2)
    assert mc.coordinate == "B3"
    assert mc.row == 3
    assert mc.column == 2


def test_merged_cell_reexport_path_matches_direct() -> None:
    """``wolfxl.cell.MergedCell`` and ``wolfxl.cell._merged.MergedCell``
    must be the same class object (re-export, not a copy)."""
    assert MergedCell is MergedCellDirect


def test_merged_subordinate_check_caches_raw_reader_refs() -> None:
    class Reader:
        def __init__(self) -> None:
            self.calls = 0

        def read_merged_ranges(self, _sheet: str) -> list[str]:
            self.calls += 1
            return ["A1:B2"]

    class Workbook:
        def __init__(self) -> None:
            self._rust_reader = Reader()

    class Worksheet:
        title = "Sheet"
        _title = "Sheet"

        def __init__(self) -> None:
            self._workbook = Workbook()
            self._merged_ranges: set[str] = set()
            self._merged_ranges_loaded = False
            self._collection_merged_ranges: set[str] = set()

    ws = Worksheet()
    anchor = Cell(ws, 1, 1)  # type: ignore[arg-type]
    subordinate = Cell(ws, 2, 2)  # type: ignore[arg-type]

    assert _is_merged_subordinate(anchor) is False
    assert _is_merged_subordinate(subordinate) is True
    assert ws._workbook._rust_reader.calls == 1


def test_empty_merged_subordinate_check_caches_reader_miss() -> None:
    class Reader:
        def __init__(self) -> None:
            self.calls = 0

        def read_merged_ranges(self, _sheet: str) -> list[str]:
            self.calls += 1
            return []

    class Workbook:
        def __init__(self) -> None:
            self._rust_reader = Reader()

    class Worksheet:
        title = "Sheet"
        _title = "Sheet"

        def __init__(self) -> None:
            self._workbook = Workbook()
            self._merged_ranges: set[str] = set()
            self._merged_ranges_loaded = False
            self._collection_merged_ranges: set[str] = set()

    ws = Worksheet()
    cell = Cell(ws, 2, 2)  # type: ignore[arg-type]

    assert _is_merged_subordinate(cell) is False
    assert _is_merged_subordinate(cell) is False
    assert ws._workbook._rust_reader.calls == 1


# ---------------------------------------------------------------------------
# WriteOnlyCell
# ---------------------------------------------------------------------------


def test_write_only_cell_default_construction() -> None:
    wc = WriteOnlyCell()
    assert wc.parent is None
    assert wc.value is None
    assert wc.font is None
    assert wc.fill is None
    assert wc.number_format is None


def test_write_only_cell_value_and_style_passthrough() -> None:
    """Construction-time fields stick on the instance."""
    sentinel_font = object()
    sentinel_fill = object()
    wc = WriteOnlyCell(
        ws=None,
        value=42,
        font=sentinel_font,
        fill=sentinel_fill,
        number_format="0.00",
    )
    assert wc.value == 42
    assert wc.font is sentinel_font
    assert wc.fill is sentinel_fill
    assert wc.number_format == "0.00"


def test_write_only_cell_value_is_settable() -> None:
    """Unlike MergedCell, WriteOnlyCell.value is mutable."""
    wc = WriteOnlyCell(value="initial")
    assert wc.value == "initial"
    wc.value = "updated"
    assert wc.value == "updated"


def test_write_only_cell_reexport_path_matches_direct() -> None:
    assert WriteOnlyCell is WriteOnlyCellDirect


def test_cell_style_accessors_share_format_payload() -> None:
    """font/fill/number_format should reuse one reader format lookup."""

    class Reader:
        def __init__(self) -> None:
            self.calls = 0

        def read_cell_format(self, sheet: str, coordinate: str) -> dict[str, object]:
            self.calls += 1
            assert sheet == "Sheet"
            assert coordinate == "B2"
            return {
                "bold": True,
                "bg_color": "FFD966",
                "number_format": "#,##0.00",
            }

    class Workbook:
        _format = "xlsx"

        def __init__(self) -> None:
            self._rust_reader = Reader()

    class Worksheet:
        title = "Sheet"

        def __init__(self) -> None:
            self._workbook = Workbook()

    ws = Worksheet()
    cell = Cell(ws, 2, 2)  # type: ignore[arg-type]

    assert cell.font.bold is True
    assert cell.fill.fill_type == "solid"
    assert cell.number_format == "#,##0.00"
    assert ws._workbook._rust_reader.calls == 1


def test_cached_style_accessor_skips_repeated_format_guard(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """A cached style object should return without rechecking workbook format."""
    cell = Cell(object(), 1, 1)  # type: ignore[arg-type]
    cached_fill = PatternFill(patternType="solid", fgColor="FFD966")
    cell._fill = cached_fill  # noqa: SLF001

    def fail_guard(self: Cell, attr: str) -> None:
        raise AssertionError(f"unexpected repeated style guard for {attr}")

    monkeypatch.setattr(Cell, "_require_xlsx_for_style", fail_guard)

    assert cell.fill is cached_fill


def test_cell_style_reader_prefers_row_column_format_lookup() -> None:
    """XLSX readers can avoid repeated A1 coordinate conversion."""

    class Reader:
        def __init__(self) -> None:
            self.calls: list[tuple[str, int, int]] = []

        def read_cell_format_rc(
            self,
            sheet: str,
            row: int,
            col: int,
        ) -> dict[str, object]:
            self.calls.append((sheet, row, col))
            return {"bold": True}

        def read_cell_format(self, sheet: str, coordinate: str) -> dict[str, object]:
            raise AssertionError(f"unexpected A1 lookup for {sheet} {coordinate}")

    class Workbook:
        _format = "xlsx"

        def __init__(self) -> None:
            self._rust_reader = Reader()

    class Worksheet:
        title = "Sheet"

        def __init__(self) -> None:
            self._workbook = Workbook()

    ws = Worksheet()
    cell = Cell(ws, 2, 3)  # type: ignore[arg-type]

    assert cell.font.bold is True
    assert ws._workbook._rust_reader.calls == [("Sheet", 2, 3)]


def test_cell_style_reader_batches_payloads_for_clean_worksheets() -> None:
    """First style access can cache native style-id payloads for nearby cells."""

    class Reader:
        def __init__(self) -> None:
            self.record_calls = 0

        def read_sheet_style_ids(
            self,
            sheet: str,
            range_str: str,
        ) -> list[tuple[int, int, int]]:
            self.record_calls += 1
            assert sheet == "Sheet"
            assert range_str == "A1:A2"
            return [(1, 1, 7), (2, 1, 8)]

        def read_format_for_style_id(self, style_id: int) -> dict[str, object]:
            if style_id == 7:
                return {"bold": True, "number_format": "0.00"}
            if style_id == 8:
                return {"bg_color": "FFD966"}
            raise AssertionError(f"unexpected style id {style_id}")

        def read_cell_format_rc(
            self,
            sheet: str,
            row: int,
            col: int,
        ) -> dict[str, object]:
            raise AssertionError(f"unexpected per-cell style lookup for {(sheet, row, col)}")

    class Workbook:
        _format = "xlsx"
        _data_only = False

        def __init__(self) -> None:
            self._rust_reader = Reader()

    class Worksheet:
        title = "Sheet"
        _dirty: set[tuple[int, int]] = set()
        _format_dirty_cells: set[tuple[int, int]] = set()
        _merged_bounds_cache: list[tuple[int, int, int, int]] = []

        def __init__(self) -> None:
            self._workbook = Workbook()
            self._style_payload_cache = None

        def _bounds_max_row(self) -> int:
            return 2

        def _bounds_max_col(self) -> int:
            return 1

    ws = Worksheet()
    first = Cell(ws, 1, 1)  # type: ignore[arg-type]
    second = Cell(ws, 2, 1)  # type: ignore[arg-type]

    assert first.font.bold is True
    assert first.number_format == "0.00"
    assert second.fill.fill_type == "solid"
    assert ws._workbook._rust_reader.record_calls == 1


def test_basic_style_access_caches_sibling_components() -> None:
    """One native style payload should satisfy repeated font/fill reads."""

    class Reader:
        def __init__(self) -> None:
            self.record_calls = 0

        def read_sheet_style_ids(
            self,
            sheet: str,
            range_str: str,
        ) -> list[tuple[int, int, int]]:
            self.record_calls += 1
            return [(1, 1, 7)]

        def read_format_for_style_id(self, style_id: int) -> dict[str, object]:
            assert style_id == 7
            return {"bold": True, "bg_color": "FFD966"}

    class Workbook:
        _format = "xlsx"

        def __init__(self) -> None:
            self._rust_reader = Reader()

    class Worksheet:
        title = "Sheet"
        _dirty: set[tuple[int, int]] = set()
        _format_dirty_cells: set[tuple[int, int]] = set()

        def __init__(self) -> None:
            self._workbook = Workbook()
            self._style_payload_cache = None

        def _bounds_max_row(self) -> int:
            return 1

        def _bounds_max_col(self) -> int:
            return 1

    ws = Worksheet()
    cell = Cell(ws, 1, 1)  # type: ignore[arg-type]

    assert cell.fill.fill_type == "solid"
    assert cell.font.bold is True
    assert ws._workbook._rust_reader.record_calls == 1


def test_basic_style_hydration_caches_number_format_after_empty_merge_probe(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Known-unmerged sheets can reuse the font/fill payload for number format."""

    class Reader:
        def read_sheet_style_ids(
            self,
            sheet: str,
            range_str: str,
        ) -> list[tuple[int, int, int]]:
            return [(1, 1, 7)]

        def read_format_for_style_id(self, style_id: int) -> dict[str, object]:
            assert style_id == 7
            return {"bg_color": "FFD966", "number_format": "0.00"}

    class Workbook:
        _format = "xlsx"

        def __init__(self) -> None:
            self._rust_reader = Reader()

    class Worksheet:
        title = "Sheet"
        _dirty: set[tuple[int, int]] = set()
        _format_dirty_cells: set[tuple[int, int]] = set()
        _merged_ranges_loaded = True
        _merged_ranges: set[str] = set()
        _collection_merged_ranges: set[str] = set()

        def __init__(self) -> None:
            self._workbook = Workbook()
            self._style_payload_cache = None

        def _bounds_max_row(self) -> int:
            return 1

        def _bounds_max_col(self) -> int:
            return 1

    def fail_number_format_read(self: Cell) -> str:
        raise AssertionError("number format should come from the hydrated style payload")

    monkeypatch.setattr(Cell, "_read_number_format", fail_number_format_read)

    cell = Cell(Worksheet(), 1, 1)  # type: ignore[arg-type]

    assert cell.fill.fill_type == "solid"
    assert cell.number_format == "0.00"


def test_repeated_read_style_payloads_reuse_python_style_objects() -> None:
    """Cells with the same read-side payload should reuse style value objects."""

    class Reader:
        def read_sheet_style_ids(
            self,
            sheet: str,
            range_str: str,
        ) -> list[tuple[int, int, int]]:
            return [(1, 1, 7), (1, 2, 7)]

        def read_format_for_style_id(self, style_id: int) -> dict[str, object]:
            assert style_id == 7
            return {"bold": True, "bg_color": "FFD966", "number_format": "0.00"}

    class Workbook:
        _format = "xlsx"

        def __init__(self) -> None:
            self._rust_reader = Reader()

    class Worksheet:
        title = "Sheet"
        _dirty: set[tuple[int, int]] = set()
        _format_dirty_cells: set[tuple[int, int]] = set()

        def __init__(self) -> None:
            self._workbook = Workbook()
            self._style_payload_cache = None

        def _bounds_max_row(self) -> int:
            return 1

        def _bounds_max_col(self) -> int:
            return 2

    ws = Worksheet()
    left = Cell(ws, 1, 1)  # type: ignore[arg-type]
    right = Cell(ws, 1, 2)  # type: ignore[arg-type]

    assert left.font is right.font
    assert left.fill is right.fill


def test_repeated_read_style_payloads_convert_components_once(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Cells with the same style payload should not rerun style converters."""
    import wolfxl._cell as cell_module

    class Reader:
        def read_sheet_style_ids(
            self,
            sheet: str,
            range_str: str,
        ) -> list[tuple[int, int, int]]:
            return [(1, 1, 7), (1, 2, 7)]

        def read_format_for_style_id(self, style_id: int) -> dict[str, object]:
            assert style_id == 7
            return {"bold": True, "bg_color": "FFD966", "number_format": "0.00"}

    class Workbook:
        _format = "xlsx"

        def __init__(self) -> None:
            self._rust_reader = Reader()

    class Worksheet:
        title = "Sheet"
        _dirty: set[tuple[int, int]] = set()
        _format_dirty_cells: set[tuple[int, int]] = set()
        _merged_ranges_loaded = True
        _merged_ranges: set[str] = set()
        _collection_merged_ranges: set[str] = set()

        def __init__(self) -> None:
            self._workbook = Workbook()
            self._style_payload_cache = None
            self._style_component_cache = None

        def _bounds_max_row(self) -> int:
            return 1

        def _bounds_max_col(self) -> int:
            return 2

    font_calls = 0
    fill_calls = 0

    original_font = cell_module._format_to_font
    original_fill = cell_module._format_to_fill

    def tracked_font(payload: object) -> object:
        nonlocal font_calls
        font_calls += 1
        return original_font(payload)

    def tracked_fill(payload: object) -> object:
        nonlocal fill_calls
        fill_calls += 1
        return original_fill(payload)

    monkeypatch.setattr(cell_module, "_format_to_font", tracked_font)
    monkeypatch.setattr(cell_module, "_format_to_fill", tracked_fill)

    ws = Worksheet()
    left = Cell(ws, 1, 1)  # type: ignore[arg-type]
    right = Cell(ws, 1, 2)  # type: ignore[arg-type]

    assert left.fill.fill_type == "solid"
    assert right.font.bold is True
    assert right.number_format == "0.00"
    assert font_calls == 1
    assert fill_calls == 1
