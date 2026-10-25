"""A second in-place save after an in-place save that added an image.

The patcher used to keep the first save's added parts (``file_adds``) after the
source file was rewritten in place, so the next save hit the "file_adds collision
with source entry" assertion and surfaced as a ``PanicException``.
"""

from __future__ import annotations

import zipfile
from collections import Counter
from pathlib import Path

import pytest

import wolfxl
from wolfxl import load_workbook
from wolfxl.drawing import Image

FIXTURES = Path(__file__).parent / "fixtures" / "images"
PNG_PATH = FIXTURES / "tiny_red_dot.png"
JPG_PATH = FIXTURES / "tiny_blue_dot.jpg"


def _workbook(tmp_path: Path, *, with_image: bool) -> Path:
    wb = wolfxl.Workbook()
    ws = wb.active
    ws["A1"] = "baseline"
    if with_image:
        ws.add_image(Image(PNG_PATH), "B2")
    path = tmp_path / "book.xlsx"
    wb.save(path)
    return path


def _anchors(path: Path) -> list[tuple[int, int]]:
    wb = load_workbook(path)
    out = []
    for image in wb.active._images:
        marker = image.anchor._from
        out.append((int(marker.col), int(marker.row)))
    return out


def _assert_clean_package(path: Path) -> None:
    with zipfile.ZipFile(path) as z:
        assert z.testzip() is None
        names = Counter(z.namelist())
    assert [n for n, c in names.items() if c > 1] == []


@pytest.mark.parametrize("with_source_image", [False, True])
def test_cell_edit_after_in_place_image_add(tmp_path: Path, with_source_image: bool) -> None:
    path = _workbook(tmp_path, with_image=with_source_image)
    wb = load_workbook(path, modify=True)
    ws = wb.active
    ws.add_image(Image(JPG_PATH), "D4")
    wb.save(path)

    ws["A1"] = "edited"
    wb.save(path)

    expected = [(3, 3)]
    if with_source_image:
        expected.insert(0, (1, 1))
    assert _anchors(path) == expected
    assert load_workbook(path).active["A1"].value == "edited"
    _assert_clean_package(path)


@pytest.mark.parametrize("with_source_image", [False, True])
def test_second_image_add_after_in_place_image_add(tmp_path: Path, with_source_image: bool) -> None:
    path = _workbook(tmp_path, with_image=with_source_image)
    wb = load_workbook(path, modify=True)
    ws = wb.active
    ws.add_image(Image(JPG_PATH), "D4")
    wb.save(path)
    ws.add_image(Image(PNG_PATH), "F6")
    wb.save(path)

    expected = [(3, 3), (5, 5)]
    if with_source_image:
        expected.insert(0, (1, 1))
    assert _anchors(path) == expected
    _assert_clean_package(path)


@pytest.mark.parametrize("with_source_image", [False, True])
def test_remove_image_added_by_first_in_place_save(tmp_path: Path, with_source_image: bool) -> None:
    path = _workbook(tmp_path, with_image=with_source_image)
    wb = load_workbook(path, modify=True)
    ws = wb.active
    added = Image(JPG_PATH)
    ws.add_image(added, "D4")
    wb.save(path)
    ws.remove_image(added)
    wb.save(path)

    assert _anchors(path) == ([(1, 1)] if with_source_image else [])
    _assert_clean_package(path)


def test_remove_added_image_after_hydrated_images_and_third_save(tmp_path: Path) -> None:
    path = _workbook(tmp_path, with_image=True)
    wb = load_workbook(path, modify=True)
    ws = wb.active
    assert len(ws.images) == 1
    first, second = Image(JPG_PATH), Image(PNG_PATH)
    ws.add_image(first, "D4")
    ws.add_image(second, "F6")
    wb.save(path)
    ws.remove_image(first)
    wb.save(path)
    assert _anchors(path) == [(1, 1), (5, 5)]
    ws.remove_image(second)
    wb.save(path)

    assert _anchors(path) == [(1, 1)]
    _assert_clean_package(path)


def test_reopened_workbook_after_resave_accepts_further_edits(tmp_path: Path) -> None:
    path = _workbook(tmp_path, with_image=False)
    wb = load_workbook(path, modify=True)
    wb.active.add_image(Image(JPG_PATH), "D4")
    wb.save(path)
    wb.active["B1"] = 7
    wb.save(path)

    reopened = load_workbook(path, modify=True)
    reopened.active.add_image(Image(PNG_PATH), "F6")
    reopened.save(path)
    assert _anchors(path) == [(3, 3), (5, 5)]
    _assert_clean_package(path)
