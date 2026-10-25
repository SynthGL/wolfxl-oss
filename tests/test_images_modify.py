"""Sprint Λ Pod-β (RFC-045) — modify-mode image tests.

Counterpart to ``test_images_write.py``. Verifies that
``ws.add_image(Image(...))`` works in modify mode (load → mutate →
save) and produces a valid xlsx with all the same parts. Also
exercises the v1.5 limit: appending to a sheet that already has a
drawing part now merges into the existing drawing part.
"""

from __future__ import annotations

import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import pytest

import wolfxl
from wolfxl import load_workbook
from wolfxl.drawing import Image

FIXTURES = Path(__file__).parent / "fixtures" / "images"
PNG_PATH = FIXTURES / "tiny_red_dot.png"
JPG_PATH = FIXTURES / "tiny_blue_dot.jpg"


def _baseline(tmp_path: Path) -> Path:
    """Write a baseline xlsx with no images."""
    wb = wolfxl.Workbook()
    ws = wb.active
    ws["A1"] = "baseline"
    out = tmp_path / "baseline.xlsx"
    wb.save(out)
    return out


def _image_anchor_col_row(image: object) -> tuple[int, int]:
    marker = getattr(getattr(image, "anchor", None), "_from", None)
    assert marker is not None
    return int(marker.col), int(marker.row)


def test_modify_add_first_image(tmp_path: Path) -> None:
    base = _baseline(tmp_path)
    wb = load_workbook(base, modify=True)
    wb.active.add_image(Image(PNG_PATH), "B5")
    out = tmp_path / "with_image.xlsx"
    wb.save(out)

    with zipfile.ZipFile(out) as z:
        names = set(z.namelist())
        assert "xl/drawings/drawing1.xml" in names
        assert "xl/drawings/_rels/drawing1.xml.rels" in names
        assert "xl/media/image1.png" in names

        # Original cell preserved.
        s1 = z.read("xl/worksheets/sheet1.xml").decode()
        assert "<drawing r:id=" in s1

        rels = z.read("xl/worksheets/_rels/sheet1.xml.rels").decode()
        assert "/relationships/drawing" in rels

        ct = z.read("[Content_Types].xml").decode()
        assert 'Extension="png"' in ct
        assert "/xl/drawings/drawing1.xml" in ct


def test_modify_preserves_existing_data(tmp_path: Path) -> None:
    """Adding an image must not corrupt other cells."""
    openpyxl = pytest.importorskip("openpyxl")

    base = _baseline(tmp_path)
    wb = load_workbook(base, modify=True)
    wb.active.add_image(Image(PNG_PATH), "B5")
    out = tmp_path / "with_image.xlsx"
    wb.save(out)

    wb2 = openpyxl.load_workbook(out)
    assert wb2.active["A1"].value == "baseline"
    assert len(wb2.active._images) == 1


def test_modify_multiple_images(tmp_path: Path) -> None:
    base = _baseline(tmp_path)
    wb = load_workbook(base, modify=True)
    wb.active.add_image(Image(PNG_PATH), "B2")
    wb.active.add_image(Image(JPG_PATH), "D4")
    out = tmp_path / "multi.xlsx"
    wb.save(out)

    with zipfile.ZipFile(out) as z:
        names = set(z.namelist())
        assert "xl/media/image1.png" in names
        assert "xl/media/image2.jpeg" in names
        d_rels = z.read("xl/drawings/_rels/drawing1.xml.rels").decode()
        assert d_rels.count("/relationships/image") == 2


def test_modify_append_to_existing_drawing_round_trip(tmp_path: Path) -> None:
    """Appending to a sheet that already has a drawing part preserves both images."""
    openpyxl = pytest.importorskip("openpyxl")

    # Stage 1: create a workbook WITH an image (so it has a drawing rel).
    wb = wolfxl.Workbook()
    wb.active.add_image(Image(PNG_PATH), "B5")
    base = tmp_path / "with_drawing.xlsx"
    wb.save(base)

    # Stage 2: open in modify mode, add another image into the existing drawing.
    wb2 = load_workbook(base, modify=True)
    wb2.active.add_image(Image(JPG_PATH), "D4")
    out = tmp_path / "merged_drawing.xlsx"
    wb2.save(out)

    wb3 = openpyxl.load_workbook(out)
    images = wb3.active._images
    assert len(images) == 2
    assert _image_anchor_col_row(images[0]) == (1, 4)  # B5
    assert _image_anchor_col_row(images[1]) == (3, 3)  # D4

    with zipfile.ZipFile(out) as z:
        names = set(z.namelist())
        drawing_files = sorted(n for n in names if n.startswith("xl/drawings/drawing"))
        assert drawing_files == ["xl/drawings/drawing1.xml"]
        drawing = z.read("xl/drawings/drawing1.xml").decode()
        assert drawing.count("<xdr:pic") == 2
        rels = z.read("xl/drawings/_rels/drawing1.xml.rels").decode()
        assert rels.count("/relationships/image") == 2


def test_modify_remove_image_by_index_round_trip(tmp_path: Path) -> None:
    openpyxl = pytest.importorskip("openpyxl")

    wb = wolfxl.Workbook()
    ws = wb.active
    ws.add_image(Image(PNG_PATH), "B2")
    ws.add_image(Image(JPG_PATH), "D4")
    base = tmp_path / "two_images.xlsx"
    wb.save(base)

    wb2 = load_workbook(base, modify=True)
    ws2 = wb2.active
    assert len(ws2.images) == 2
    ws2.remove_image(0)
    out = tmp_path / "after_remove.xlsx"
    wb2.save(out)

    wb3 = openpyxl.load_workbook(out)
    images = wb3.active._images
    assert len(images) == 1
    assert _image_anchor_col_row(images[0]) == (3, 3)  # D4


def test_modify_replace_image_by_index_round_trip(tmp_path: Path) -> None:
    openpyxl = pytest.importorskip("openpyxl")

    wb = wolfxl.Workbook()
    ws = wb.active
    ws.add_image(Image(PNG_PATH), "B2")
    base = tmp_path / "one_image.xlsx"
    wb.save(base)

    wb2 = load_workbook(base, modify=True)
    ws2 = wb2.active
    replacement = Image(JPG_PATH)
    replacement.anchor = "F6"
    ws2.replace_image(0, replacement)
    out = tmp_path / "after_replace.xlsx"
    wb2.save(out)

    wb3 = openpyxl.load_workbook(out)
    images = wb3.active._images
    assert len(images) == 1
    assert _image_anchor_col_row(images[0]) == (5, 5)  # F6


def _unreadable_first_image(tmp_path: Path, ext: str = "wmf") -> tuple[Path, str, bytes]:
    """Keep real drawing relationships but replace the first image's bytes."""
    wb = wolfxl.Workbook()
    ws = wb.active
    ws.add_image(Image(PNG_PATH), "B2")
    ws.add_image(Image(PNG_PATH), "D4")
    ws.add_image(Image(JPG_PATH), "F6")
    original = tmp_path / "original_images.xlsx"
    wb.save(original)

    part = f"xl/media/image1.{ext}"
    unreadable = b"unreadable source image"
    source = tmp_path / "unreadable_image.xlsx"
    with zipfile.ZipFile(original) as zin, zipfile.ZipFile(source, "w") as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "xl/media/image1.png":
                zout.writestr(part, unreadable)
                continue
            if info.filename == "xl/drawings/_rels/drawing1.xml.rels":
                data = data.replace(b"image1.png", f"image1.{ext}".encode())
            elif info.filename == "[Content_Types].xml" and ext != "png":
                root = ET.fromstring(data)
                ET.SubElement(
                    root,
                    "{http://schemas.openxmlformats.org/package/2006/content-types}Default",
                    Extension=ext,
                    ContentType=f"image/{ext}",
                )
                data = ET.tostring(root)
            zout.writestr(info, data)
    return source, part, unreadable


@pytest.mark.parametrize("ext", ["wmf", "emf", "unknown", "png"])
@pytest.mark.parametrize("modify", [False, True])
def test_unreadable_source_image_is_omitted(
    tmp_path: Path, ext: str, modify: bool,
) -> None:
    source, part, unreadable = _unreadable_first_image(tmp_path, ext)
    wb = load_workbook(source, modify=modify)
    with pytest.warns(UserWarning, match=rf"{ext} image format.*being dropped") as caught:
        images = wb.active._images
    assert len(caught) == 1
    assert [image.format for image in images] == ["png", "jpeg"]
    assert images[0]._data() == PNG_PATH.read_bytes()
    assert images[1]._data() == JPG_PATH.read_bytes()
    assert wb.active._images is images

    if modify:
        output = tmp_path / "preserved.xlsx"
        wb.save(output)
        with zipfile.ZipFile(source) as before, zipfile.ZipFile(output) as after:
            assert after.read(part) == unreadable
            for name in (
                "xl/drawings/drawing1.xml",
                "xl/drawings/_rels/drawing1.xml.rels",
                "xl/media/image2.png",
                "xl/media/image3.jpeg",
            ):
                assert after.read(name) == before.read(name)


@pytest.mark.parametrize("operation", ["remove", "replace", "remove_both"])
def test_unreadable_source_image_keeps_mutation_indices(
    tmp_path: Path, operation: str,
) -> None:
    source, part, unreadable = _unreadable_first_image(tmp_path)
    wb = load_workbook(source, modify=True)
    ws = wb.active
    with pytest.warns(UserWarning, match="wmf image format"):
        png, jpeg = ws._images
    if operation == "replace":
        replacement = Image(JPG_PATH)
        replacement.anchor = "H8"
        ws.replace_image(png, replacement)
    else:
        ws.remove_image(png)
        if operation == "remove_both":
            ws.remove_image(jpeg)
    output = tmp_path / "mutated.xlsx"
    wb.save(output)

    with zipfile.ZipFile(output) as archive:
        assert archive.read(part) == unreadable
        relationships = ET.fromstring(
            archive.read("xl/drawings/_rels/drawing1.xml.rels")
        )
        targets = [rel.attrib["Target"] for rel in relationships]
        assert any(target.endswith("image1.wmf") for target in targets)
        assert not any(target.endswith("image2.png") for target in targets)
        drawing = ET.fromstring(archive.read("xl/drawings/drawing1.xml"))
        ns = {"xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"}
        markers = [
            (int(marker.findtext("xdr:col", namespaces=ns)),
             int(marker.findtext("xdr:row", namespaces=ns)))
            for marker in drawing.findall("*/xdr:from", ns)
        ]
        expected = [(1, 1)]
        if operation != "remove_both":
            assert archive.read("xl/media/image3.jpeg") == JPG_PATH.read_bytes()
            expected.append((5, 5))
        else:
            assert not any(target.endswith("image3.jpeg") for target in targets)
        if operation == "replace":
            expected.append((7, 7))
        assert sorted(markers) == expected


def test_unreadable_source_warning_names_payload_part() -> None:
    from wolfxl._worksheet_media import _image_from_payload

    with pytest.warns(UserWarning, match=r"wmf image format.*xl/media/image1\.wmf"):
        assert _image_from_payload({
            "data": b"unreadable source image",
            "ext": "wmf",
            "media_path": "xl/media/image1.wmf",
        }) is None


def test_unreadable_user_image_still_raises() -> None:
    with pytest.raises(ValueError, match="could not parse image dimensions"):
        Image(b"unreadable source image")


def _three_image_workbook(tmp_path: Path) -> Path:
    wb = wolfxl.Workbook()
    ws = wb.active
    ws.add_image(Image(PNG_PATH), "B2")
    ws.add_image(Image(PNG_PATH), "D4")
    ws.add_image(Image(JPG_PATH), "F6")
    path = tmp_path / "three_images.xlsx"
    wb.save(path)
    return path


def _saved_anchors(path: Path) -> list[tuple[int, int]]:
    return [_image_anchor_col_row(image) for image in load_workbook(path).active._images]


def test_repeated_in_place_saves_remove_the_selected_images(tmp_path: Path) -> None:
    path = _three_image_workbook(tmp_path)
    wb = load_workbook(path, modify=True)
    ws = wb.active
    ws.remove_image(0)
    wb.save(path)
    ws.remove_image(0)
    wb.save(path)

    assert _saved_anchors(path) == [(5, 5)]
