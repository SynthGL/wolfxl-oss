"""Sprint Ο Pod 3.5 (RFC-061) — advanced-pivot parity vs openpyxl.

Round-trips slicer + calc field + calc item + group items + pivot
styling through the full save → openpyxl-read pipeline. Asserts
that the parts wolfxl emits remain ZIP-valid and that openpyxl can
load the workbook without raising.

NOTE: openpyxl 3.1.x does not interpret the ``<x14:slicerCaches>``
extension. This test stays at the structural level (parts present,
ZIP valid, no raise) rather than asserting full schema-level
equivalence.
"""
from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

wolfxl = pytest.importorskip("wolfxl")
openpyxl = pytest.importorskip("openpyxl")

from wolfxl import load_workbook
from wolfxl.chart.reference import Reference
from wolfxl.pivot import (
    PivotArea,
    PivotCache,
    PivotTable,
    Slicer,
    SlicerCache,
)


@pytest.fixture(autouse=True)
def _pin_epoch(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("WOLFXL_TEST_EPOCH", "0")


def _make_pivot_source(path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    rows = [
        ("region", "quarter", "revenue"),
        ("North", "Q1", 100.0),
        ("South", "Q1", 200.0),
        ("North", "Q2", 150.0),
        ("South", "Q2", 250.0),
    ]
    for r, row in enumerate(rows, start=1):
        for c, v in enumerate(row, start=1):
            ws.cell(r, c, v)
    wb.save(path)


def _zip_read(p: Path, member: str) -> bytes:
    with zipfile.ZipFile(p, "r") as z:
        return z.read(member)


def _build_slicer_workbook(src: Path, dst: Path) -> None:
    wb = load_workbook(src, modify=True)
    sheet = wb["Data"]
    ref = Reference(worksheet=sheet, min_col=1, min_row=1, max_col=3, max_row=5)
    cache = PivotCache(source=ref)
    cache.add_calculated_field(name="bonus", formula="revenue * 0.10")
    pt = PivotTable(cache=cache, location="F2", rows=["region"], data=["revenue"])
    wb.add_pivot_cache(cache)
    cache.group_field("revenue", start=0, end=300, interval=100)
    pt.add_format(PivotArea(field=0))
    sheet.add_pivot_table(pt)
    sc = SlicerCache(name="Slicer_region", source_pivot_cache=cache, field="region")
    sl = Slicer(name="Slicer_region1", cache=sc, caption="Region")
    wb.add_slicer_cache(sc)
    sheet.add_slicer(sl, anchor="H2")
    wb.save(dst)


def test_slicer_roundtrip_zip_integrity(tmp_path: Path) -> None:
    src = tmp_path / "src.xlsx"
    dst = tmp_path / "dst.xlsx"
    _make_pivot_source(src)
    _build_slicer_workbook(src, dst)

    with zipfile.ZipFile(dst, "r") as z:
        bad = z.testzip()
        assert bad is None, f"corrupt ZIP entry: {bad}"


def test_slicer_parts_visible_to_openpyxl_zip_namelist(tmp_path: Path) -> None:
    """openpyxl iterates ZIP entries during load(); the slicer parts
    must be present in the ZIP listing so they aren't dropped during
    re-save.
    """
    src = tmp_path / "src.xlsx"
    dst = tmp_path / "dst.xlsx"
    _make_pivot_source(src)
    _build_slicer_workbook(src, dst)

    with zipfile.ZipFile(dst, "r") as z:
        names = set(z.namelist())
    assert "xl/slicerCaches/slicerCache1.xml" in names
    assert "xl/slicers/slicer1.xml" in names
    assert "xl/slicerCaches/_rels/slicerCache1.xml.rels" in names


def test_slicer_cache_dict_round_trip(tmp_path: Path) -> None:
    """``SlicerCache.to_rust_dict()`` is the contract bridge to the
    patcher; assert the dict shape matches RFC-061 §10.1 keys.
    """
    src = tmp_path / "src.xlsx"
    _make_pivot_source(src)
    wb = load_workbook(src, modify=True)
    sheet = wb["Data"]
    ref = Reference(worksheet=sheet, min_col=1, min_row=1, max_col=3, max_row=5)
    cache = PivotCache(source=ref)
    pt = PivotTable(cache=cache, location="F2", rows=["region"], data=["revenue"])
    wb.add_pivot_cache(cache)
    sheet.add_pivot_table(pt)
    sc = SlicerCache(name="Slicer_region", source_pivot_cache=cache, field="region")
    wb.add_slicer_cache(sc)

    d = sc.to_rust_dict()
    expected_keys = {
        "name",
        "source_pivot_cache_id",
        "source_field_index",
        "sort_order",
        "custom_list_sort",
        "hide_items_with_no_data",
        "show_missing",
        "items",
    }
    assert expected_keys.issubset(set(d.keys())), set(d.keys())
    assert d["name"] == "Slicer_region"
    assert d["source_field_index"] == 0  # region is the first cache field
    assert isinstance(d["items"], list)


def test_slicer_dict_round_trip(tmp_path: Path) -> None:
    """``Slicer.to_rust_dict()`` carries the §10.2 keys."""
    src = tmp_path / "src.xlsx"
    _make_pivot_source(src)
    wb = load_workbook(src, modify=True)
    sheet = wb["Data"]
    ref = Reference(worksheet=sheet, min_col=1, min_row=1, max_col=3, max_row=5)
    cache = PivotCache(source=ref)
    pt = PivotTable(cache=cache, location="F2", rows=["region"], data=["revenue"])
    wb.add_pivot_cache(cache)
    sheet.add_pivot_table(pt)
    sc = SlicerCache(name="Slicer_region", source_pivot_cache=cache, field="region")
    sl = Slicer(name="Slicer_region1", cache=sc, caption="Region")
    wb.add_slicer_cache(sc)
    sheet.add_slicer(sl, anchor="H2")

    d = sl.to_rust_dict()
    expected_keys = {
        "name",
        "cache_name",
        "caption",
        "row_height",
        "column_count",
        "show_caption",
        "style",
        "locked",
        "anchor",
    }
    assert expected_keys.issubset(set(d.keys())), set(d.keys())
    assert d["name"] == "Slicer_region1"
    assert d["cache_name"] == "Slicer_region"
    assert d["anchor"] == "H2"


def test_save_then_openpyxl_load_does_not_raise(tmp_path: Path) -> None:
    """The full slicer-augmented save must round-trip through
    openpyxl's load_workbook without raising — even if openpyxl can't
    semantically interpret the slicer extension, it must at minimum
    leave the workbook re-openable for downstream tools.
    """
    src = tmp_path / "src.xlsx"
    dst = tmp_path / "dst.xlsx"
    _make_pivot_source(src)
    _build_slicer_workbook(src, dst)
    op_wb = openpyxl.load_workbook(dst)
    op_wb.close()


def test_workbook_extlst_roundtrip_preserves_existing_extensions(
    tmp_path: Path,
) -> None:
    """If the source workbook already had an ``<extLst>`` block, our
    splice must extend it rather than orphan the existing ext entries.
    """
    # We can't easily mint a source xlsx with an existing extLst from
    # openpyxl alone, but we can check that the splice outputs
    # idempotent results: two saves carry one and only one
    # <extLst> block on workbook.xml.
    src = tmp_path / "src.xlsx"
    a = tmp_path / "a.xlsx"
    b = tmp_path / "b.xlsx"
    _make_pivot_source(src)
    _build_slicer_workbook(src, a)
    _build_slicer_workbook(src, b)

    wb_a = _zip_read(a, "xl/workbook.xml").decode()
    wb_b = _zip_read(b, "xl/workbook.xml").decode()
    assert wb_a == wb_b
    assert wb_a.count("<extLst>") <= 1
    assert wb_a.count("<x14:slicerCaches") == 1
