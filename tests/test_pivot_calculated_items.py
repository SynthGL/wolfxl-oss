"""Calculated items: cache-scoped emit and rendered values (issue #64)."""

from __future__ import annotations

import zipfile

import pytest

import wolfxl
from wolfxl import Workbook
from wolfxl.chart.reference import Reference
from wolfxl.pivot import CalculatedItem, PivotCache, PivotTable
from wolfxl.pivot._calc import qualify_item_formula

ROWS = [
    ("region", "product", "revenue"),
    ("North", "A", 100.0),
    ("South", "A", 200.0),
    ("North", "B", 150.0),
    ("East", "B", 50.0),
    ("South", "B", 25.0),
]


def _save(tmp_path, items, **table):
    wb = Workbook()
    ws = wb.active
    for row in ROWS:
        ws.append(row)
    pc = PivotCache(
        source=Reference(worksheet=ws, min_col=1, min_row=1, max_col=3, max_row=len(ROWS))
    )
    cache = wb.add_pivot_cache(pc)
    for item in items:
        cache.add_calculated_item(*item)
    table.setdefault("data", ["revenue"])
    ws.add_pivot_table(PivotTable(cache=cache, location="F2", **table))
    path = tmp_path / "calc_items.xlsx"
    wb.save(path)
    return path


def _parts(path):
    with zipfile.ZipFile(path) as z:
        cache = z.read("xl/pivotCache/pivotCacheDefinition1.xml").decode()
        table = z.read("xl/pivotTables/pivotTable1.xml").decode()
    return cache, table


def _grid(path, ref):
    ws = wolfxl.load_workbook(path).active
    return [[c.value for c in row] for row in ws[ref]]


# ---------------------------------------------------------------------------
# CalculatedItem construction
# ---------------------------------------------------------------------------


@pytest.mark.parametrize(
    "kwargs",
    [
        {"field_name": "", "item_name": "x", "formula": "a + b"},
        {"field_name": "region", "item_name": "", "formula": "a + b"},
        {"field_name": "region", "item_name": "x", "formula": ""},
        {"field_name": "region", "item_name": "x", "formula": "(a + b"},
        {"field_name": "region", "item_name": "x", "formula": '= "hi'},
    ],
)
def test_calc_item_rejects_malformed_definition(kwargs):
    with pytest.raises(ValueError):
        CalculatedItem(**kwargs)


# ---------------------------------------------------------------------------
# Emit: Excel's cache-scoped shape
# ---------------------------------------------------------------------------


def test_calc_item_is_emitted_on_the_cache(tmp_path):
    path = _save(tmp_path, [("region", "Both", "= North + South")], rows=["region"])
    cache, table = _parts(path)

    assert '<s v="Both" f="1"/>' in cache
    assert (
        '<calculatedItems count="1"><calculatedItem formula="region[North] + region[South]">'
        '<pivotArea cacheIndex="1" outline="0" fieldPosition="0"><references count="1">'
        '<reference field="0" count="1"><x v="3"/></reference></references></pivotArea>'
        "</calculatedItem></calculatedItems>"
    ) in cache
    assert '<item x="3" f="1"/>' in table
    assert "<calculatedItems" not in table


def test_pivot_table_add_calculated_item_registers_on_cache(tmp_path):
    wb = Workbook()
    ws = wb.active
    for row in ROWS:
        ws.append(row)
    pc = PivotCache(
        source=Reference(worksheet=ws, min_col=1, min_row=1, max_col=3, max_row=len(ROWS))
    )
    cache = wb.add_pivot_cache(pc)
    pt = PivotTable(cache=cache, location="F2", rows=["region"], data=["revenue"])
    ci = pt.add_calculated_item("region", "Both", "North + South")
    assert cache.calculated_items == [ci]


@pytest.mark.parametrize(
    ("field", "item"),
    [("nope", "X"), ("region", "north"), ("region", "North")],
)
def test_calc_item_rejects_unknown_field_and_duplicate_item(tmp_path, field, item):
    with pytest.raises(ValueError):
        _save(tmp_path, [(field, item, "North + South")], rows=["region"])


@pytest.mark.parametrize(
    ("expr", "expected"),
    [
        ("North + South", "region[North] + region[South]"),
        ("north*2", "region[North]*2"),
        ("'New York' - North", "region['New York'] - region[North]"),
        ("region[North] + North", "region[North] + region[North]"),
        ("SUM(North) + Other", "SUM(region[North]) + Other"),
        ('IF(North > 0, "North", 0)', 'IF(region[North] > 0, "North", 0)'),
        ("1e2 + North", "1e2 + region[North]"),
    ],
)
def test_qualify_item_formula(expr, expected):
    items = {"north": "North", "south": "South", "new york": "New York"}
    assert qualify_item_formula(expr, "region", items) == expected


def test_qualify_item_formula_quotes_field_names():
    assert (
        qualify_item_formula("Shipped", "Order Status", {"shipped": "Shipped"})
        == "'Order Status'[Shipped]"
    )


# ---------------------------------------------------------------------------
# Rendered values (matched against Excel refresh output)
# ---------------------------------------------------------------------------


def test_calc_item_renders_after_regular_items_and_in_grand_total(tmp_path):
    path = _save(tmp_path, [("region", "Both", "North + South")], rows=["region"])
    assert _grid(path, "F2:G7") == [
        ["Row Labels", "Sum of revenue"],
        ["North", 250],
        ["South", 225],
        ["East", 50],
        ["Both", 475],
        ["Grand Total", 1000],
    ]
    cache, _ = _parts(path)
    assert 'refreshOnLoad="0"' in cache


def test_nonlinear_calc_item_applies_formula_to_item_sums(tmp_path):
    path = _save(tmp_path, [("region", "Ratio", "North / South")], rows=["region"])
    grid = _grid(path, "F3:G7")
    assert grid[3][0] == "Ratio"
    assert grid[3][1] == pytest.approx(250 / 225)
    assert grid[4][1] == pytest.approx(525 + 250 / 225)


def test_calc_item_on_outer_row_field_expands_children(tmp_path):
    path = _save(
        tmp_path, [("region", "Both", "North + South")], rows=["region", "product"]
    )
    grid = _grid(path, "F3:G14")
    both = [row[0] for row in grid].index("Both")
    assert grid[both : both + 3] == [["Both", 475], ["A", 300], ["B", 175]]
    assert grid[-1] == ["Grand Total", 1000]


def test_calc_item_on_column_field(tmp_path):
    path = _save(
        tmp_path,
        [("region", "Both", "North + South")],
        rows=["product"],
        cols=["region"],
    )
    grid = _grid(path, "F3:K6")
    assert grid[0] == ["Row Labels", "North", "South", "East", "Both", "Grand Total"]
    assert grid[1] == ["A", 100, 200, None, 300, 600]
    assert grid[3] == ["Grand Total", 250, 225, 50, 475, 1000]


def test_calc_item_on_field_outside_layout_refreshes_on_load(tmp_path):
    path = _save(tmp_path, [("region", "Both", "North + South")], rows=["product"])
    cache, _ = _parts(path)
    assert 'refreshOnLoad="1"' in cache


def test_unsupported_calc_item_formula_is_left_to_excel(tmp_path):
    path = _save(
        tmp_path, [("region", "X", "SUM(North, South)")], rows=["region"]
    )
    cache, _ = _parts(path)
    assert 'refreshOnLoad="1"' in cache
    assert 'formula="SUM(region[North], region[South])"' in cache
    grid = _grid(path, "F3:G7")
    assert grid[3] == ["X", None]
