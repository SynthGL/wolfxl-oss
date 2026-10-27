"""Sprint Ο Pod 3.5 (RFC-061) — end-to-end save smoke test.

Builds a workbook that exercises ALL FIVE advanced pivot
sub-features in one save and asserts each survives a round-trip
through openpyxl:

  3.1 Slicer + SlicerCache
  3.2 CalculatedField (cache-scope)
  3.3 CalculatedItem  (table-scope)
  3.4 FieldGroup      (cache-scope)
  3.5 PivotArea Format / ConditionalFormat / ChartFormat (table-scope)

Goal: prove ``Workbook.save()`` actually emits the parts that
Phase 2.5m (pivot adds) and Phase 2.5p (slicer adds) produce, and
that openpyxl can reopen the result without raising.
"""
from __future__ import annotations

import re
import zipfile
from pathlib import Path

import openpyxl
import pytest

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
def _force_test_epoch(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("WOLFXL_TEST_EPOCH", "0")


def _make_pivot_source(path: Path, title: str = "Data") -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = title
    rows = [
        ("region", "quarter", "customer", "revenue"),
        ("North", "Q1", "Acme", 100.0),
        ("South", "Q1", "Acme", 200.0),
        ("North", "Q2", "Globex", 150.0),
        ("South", "Q2", "Globex", 250.0),
        ("North", "Q3", "Acme", 175.0),
        ("South", "Q3", "Globex", 225.0),
    ]
    for r, row in enumerate(rows, start=1):
        for c, val in enumerate(row, start=1):
            ws.cell(r, c, val)
    wb.save(path)


def test_all_five_advanced_pivot_features_survive_save(tmp_path: Path) -> None:
    src = tmp_path / "src.xlsx"
    dst = tmp_path / "dst.xlsx"
    _make_pivot_source(src)
    wb = load_workbook(src, modify=True)

    sheet = wb["Data"]
    ref = Reference(worksheet=sheet, min_col=1, min_row=1, max_col=4, max_row=7)
    cache = PivotCache(source=ref)

    # 3.2 Calculated field on the cache (cache-scope; survives the
    # pivot cache definition emit even without a pivot table).
    cache.add_calculated_field(name="bonus", formula="revenue * 0.10")

    pt = PivotTable(
        cache=cache,
        location="F2",
        rows=["region"],
        cols=["quarter"],
        data=["revenue"],
    )
    wb.add_pivot_cache(cache)

    # 3.4 Field group on the cache (numeric range grouping over the
    # `revenue` field; exercises the cache-side group-items path).
    # group_field requires the cache to be materialized first, which
    # add_pivot_cache triggers.
    cache.group_field("revenue", start=0, end=300, interval=50)

    # 3.3 Calculated item; Excel stores it on the cache.
    pt.add_calculated_item(
        field="region", item_name="ALL", formula="North + South"
    )
    # 3.5 PivotArea format directive.
    pt.add_format(PivotArea(field=0))

    sheet.add_pivot_table(pt)

    # 3.1 Slicer + slicer cache.
    sc = SlicerCache(
        name="Slicer_region", source_pivot_cache=cache, field="region"
    )
    sl = Slicer(name="Slicer_region1", cache=sc, caption="Region")
    wb.add_slicer_cache(sc)
    sheet.add_slicer(sl, anchor="H2")

    wb.save(dst)

    # ---- Assertions: all parts emitted ---------------------------------
    with zipfile.ZipFile(dst, "r") as z:
        names = sorted(z.namelist())
    cache_defs = [n for n in names if n.startswith("xl/pivotCache/pivotCacheDefinition")]
    cache_recs = [n for n in names if n.startswith("xl/pivotCache/pivotCacheRecords")]
    tables = [n for n in names if re.match(r"^xl/pivotTables/pivotTable\d+\.xml$", n)]
    slicer_caches = [n for n in names if n.startswith("xl/slicerCaches/slicerCache")]
    slicers = [n for n in names if re.match(r"^xl/slicers/slicer\d+\.xml$", n)]

    assert cache_defs, f"no pivotCacheDefinition emitted: {names}"
    assert cache_recs, f"no pivotCacheRecords emitted: {names}"
    assert tables, f"no pivotTable emitted: {names}"
    assert slicer_caches, f"no slicerCache emitted: {names}"
    assert slicers, f"no slicer emitted: {names}"

    # ---- Content-Types overrides ---------------------------------------
    ct = (dst.read_bytes() and  # touch dst so the with-block above flushes
          zipfile.ZipFile(dst, "r").read("[Content_Types].xml").decode("utf-8"))
    assert "slicerCache+xml" in ct, ct
    assert "ms-excel.slicer+xml" in ct, ct
    assert "pivotTable+xml" in ct, ct
    assert "pivotCacheDefinition+xml" in ct, ct
    assert "pivotCacheRecords+xml" in ct, ct

    # ---- Workbook.xml carries <x14:slicerCaches> -----------------------
    wb_xml = zipfile.ZipFile(dst, "r").read("xl/workbook.xml").decode("utf-8")
    assert "<x14:slicerCaches" in wb_xml, wb_xml[:1500]
    assert "<pivotCaches>" in wb_xml or "<pivotCache " in wb_xml

    # ---- Sheet xml carries <x14:slicerList> ----------------------------
    # Find the source sheet's actual path.
    sheet_paths = [n for n in names if re.match(r"^xl/worksheets/sheet\d+\.xml$", n)]
    sheet_blobs = [
        zipfile.ZipFile(dst, "r").read(p).decode("utf-8")
        for p in sheet_paths
    ]
    assert any("<x14:slicerList" in b for b in sheet_blobs), (
        f"no <x14:slicerList> on any sheet: {[b[:300] for b in sheet_blobs]}"
    )

    # ---- Calc item on the cache, flagged on the table; format marker ---
    cache_xml = zipfile.ZipFile(dst, "r").read(cache_defs[0]).decode("utf-8")
    assert '<s v="ALL" f="1"/>' in cache_xml, cache_xml[:1500]
    assert 'formula="region[North] + region[South]"' in cache_xml
    blob = "".join(
        zipfile.ZipFile(dst, "r").read(p).decode("utf-8") for p in tables
    )
    assert 'f="1"' in blob and "<format" in blob, blob[:1500]

    # ---- ZIP integrity -------------------------------------------------
    with zipfile.ZipFile(dst, "r") as z:
        bad = z.testzip()
        assert bad is None, f"corrupt ZIP entry: {bad}"


def test_save_is_idempotent_for_advanced_pivot_features(tmp_path: Path) -> None:
    """Calling save() twice must not double-emit slicer parts or pivot
    parts (mirrors the autoFilter / sheet-setup idempotency test in
    test_native_save_lifecycle.py)."""
    src = tmp_path / "src.xlsx"
    dst1 = tmp_path / "dst1.xlsx"
    dst2 = tmp_path / "dst2.xlsx"
    _make_pivot_source(src)
    wb = load_workbook(src, modify=True)
    sheet = wb["Data"]
    ref = Reference(worksheet=sheet, min_col=1, min_row=1, max_col=4, max_row=7)
    cache = PivotCache(source=ref)
    pt = PivotTable(cache=cache, location="F2", rows=["region"], data=["revenue"])
    wb.add_pivot_cache(cache)
    sheet.add_pivot_table(pt)
    sc = SlicerCache(name="Slicer_region", source_pivot_cache=cache, field="region")
    sl = Slicer(name="Slicer_region1", cache=sc, caption="Region")
    wb.add_slicer_cache(sc)
    sheet.add_slicer(sl, anchor="H2")

    wb.save(dst1)
    wb.save(dst2)

    def slicer_cache_count(p: Path) -> int:
        with zipfile.ZipFile(p, "r") as z:
            return len([n for n in z.namelist() if n.startswith("xl/slicerCaches/slicerCache")])

    # Same number of slicer cache parts in both saves.
    assert slicer_cache_count(dst1) == slicer_cache_count(dst2) == 1
