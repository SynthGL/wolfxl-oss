"""Absent XML cells retain the coordinate-aware streaming contract."""
from pathlib import Path
from copy import copy

import openpyxl
import pytest
from openpyxl.styles import Border, Font, GradientFill, Side

import wolfxl
from wolfxl._streaming import StreamingCell


def test_missing_cells_keep_coordinates_and_authored_empty_styles(tmp_path: Path) -> None:
    path = tmp_path / "sparse.xlsx"
    source = openpyxl.Workbook()
    source.active["A2"] = 2
    source.active["B2"].font = Font(bold=True)
    source.active["D4"] = 4
    source.save(path)
    wb = wolfxl.load_workbook(path, read_only=True)
    try:
        rows = list(wb.active.iter_rows(min_row=1, max_row=6, min_col=1, max_col=5))
        assert len(rows) == 6
        assert all(len(row) == 5 for row in rows)
        assert rows[1][1].value is None
        assert rows[1][1].font.bold
        for row_index, row in enumerate(rows, 1):
            for column, cell in enumerate(row, 1):
                assert isinstance(cell, StreamingCell)
                assert (cell.row, cell.column) == (row_index, column)
                assert cell.coordinate == f"{openpyxl.utils.get_column_letter(column)}{row_index}"
                assert cell.parent is wb.active
                if (row_index, column) not in ((2, 1), (2, 2), (4, 4)):
                    assert cell.value is None
                    assert cell.data_type == "n"
                    assert not cell.font.bold
                    assert cell.fill.patternType is None
                    assert cell.number_format == "General"
                    assert cell.alignment.horizontal is None
    finally:
        wb.close()


@pytest.mark.parametrize("attribute", ["value", "font", "fill", "border", "alignment", "number_format", "typo"])
def test_missing_cells_refuse_mutations_and_observe_close(tmp_path: Path, attribute: str) -> None:
    path = tmp_path / "blank.xlsx"
    source = openpyxl.Workbook()
    source.active["A1"] = 1
    source.save(path)
    wb = wolfxl.load_workbook(path, read_only=True)
    cell = next(wb.active.iter_rows(min_row=2, max_row=2, max_col=1))[0]
    assert cell.coordinate == "A2"
    assert "A2" in repr(cell)
    assert cell.number_format == "General"
    with pytest.raises(RuntimeError, match="read_only=True.*A2"):
        setattr(cell, attribute, "changed")
    wb.close()
    assert cell.number_format is None
    assert cell.value is None
    assert not cell.font.bold
    assert cell.fill.patternType is None
    assert cell.border.left.style is None


def test_missing_cell_shallow_copy_keeps_coordinate_snapshot(tmp_path: Path) -> None:
    path = tmp_path / "copy-blank.xlsx"
    source = openpyxl.Workbook()
    source.active["A1"] = 1
    source.save(path)
    wb = wolfxl.load_workbook(path, read_only=True)
    try:
        cell = next(wb.active.iter_rows(min_row=2, max_row=2, max_col=1))[0]
        cloned = copy(cell)
        assert cloned is not cell
        assert cloned == cell
        assert cloned.parent is cell.parent
        assert cloned.coordinate == "A2"
        assert cloned.font == cell.font
        assert cloned.fill == cell.fill
        assert cloned.number_format == cell.number_format
        with pytest.raises(RuntimeError, match="read_only=True"):
            cloned.value = "changed"
    finally:
        wb.close()


def test_missing_cells_keep_default_source_border(tmp_path: Path) -> None:
    path = tmp_path / "border-zero.xlsx"
    source = openpyxl.Workbook()
    source._cell_styles[0].borderId = source._borders.add(
        Border(left=Side(style="thin", color="FF123456")))
    source.active["A1"] = 1
    source.save(path)
    wb = wolfxl.load_workbook(path, read_only=True)
    try:
        cell = next(wb.active.iter_rows(min_row=2, max_row=2, max_col=1))[0]
        assert cell.border.left.style == "thin"
        assert cell.border.left.color.rgb == "FF123456"
    finally:
        wb.close()


def test_gradient_at_source_style_zero_keeps_fresh_mutable_fill(tmp_path: Path) -> None:
    path = tmp_path / "gradient-zero.xlsx"
    source = openpyxl.Workbook()
    source._cell_styles[0].fillId = source._fills.add(
        GradientFill(degree=30, stop=("FF000000", "FFFFFFFF")))
    source.active["A1"] = 1
    source.save(path)
    wb = wolfxl.load_workbook(path, read_only=True)
    try:
        cell = next(wb.active.iter_rows(min_row=2, max_row=2, max_col=1))[0]
        assert cell.value is None
        assert cell.fill.degree == 30
        changed = cell.fill
        changed.degree = 90
        changed.stop[0].color.rgb = "FF123456"
        assert cell.fill.degree == 30
        assert cell.fill.stop[0].color.rgb == "FF000000"
    finally:
        wb.close()


def test_missing_cell_border_tracks_reader_replacement_and_merge_state(tmp_path: Path) -> None:
    from wolfxl._streaming import stream_iter_rows

    path = tmp_path / "merge-state.xlsx"
    source = openpyxl.Workbook()
    source.active["A1"] = "anchor"
    source.active["A1"].border = Border(bottom=Side(style="double"))
    source.save(path)
    wb = wolfxl.load_workbook(path, modify=True)
    try:
        cell = next(stream_iter_rows(wb.active, min_row=2, max_row=2, min_col=2, max_col=2))[0]
        assert cell.border.bottom.style is None
        class ReaderReplacement:
            def __init__(self, inner):
                self.inner = inner

            def __getattr__(self, name):
                return getattr(self.inner, name)

        wb._rust_reader = ReaderReplacement(wb._rust_reader)
        assert cell.number_format == "General"
        assert cell.border.bottom.style is None
        wb.active.merge_cells("A1:B2")
        assert cell.border.bottom.style == "double"
        assert not cell.font.bold
        wb.active.unmerge_cells("A1:B2")
        assert cell.border.bottom.style is None
        assert cell == StreamingCell(wb.active, 2, 2, None, None, "blank")
    finally:
        wb.close()
