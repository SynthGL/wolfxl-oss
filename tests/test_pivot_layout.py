"""Rendered pivot layouts must match what Excel shows after a refresh.

Each expected grid below was captured from Excel 16 by refreshing the
pivot table and reading the sheet back. Excel validates a pivot table's
``rowItems``/``colItems``/``location`` against the stored cells on open,
and it does not re-render until the user refreshes, so a mismatch shows
up as a repair prompt or a stale table.
"""

from __future__ import annotations

import re
import zipfile

import pytest

from wolfxl import Workbook, load_workbook
from wolfxl.chart.reference import Reference
from wolfxl.pivot import DataField, PageField, PivotCache, PivotTable

ROWS = [
    ("region", "quarter", "customer", "revenue"),
    ("North", "Q1", "Acme", 100.0),
    ("South", "Q1", "Acme", 200.0),
    ("North", "Q2", "Globex", 150.0),
    ("South", "Q2", "Globex", 250.0),
    ("North", "Q3", "Acme", 175.0),
    ("South", "Q3", "Globex", 225.0),
    ("East", "Q1", "Globex", 50.0),
]


def _populate(ws, rows):
    for row in rows:
        ws.append(row)
    return Reference(worksheet=ws, min_col=1, min_row=1, max_col=4, max_row=len(rows))


def _pivot(source=ROWS, location="F2", **kw):
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    cache = wb.add_pivot_cache(PivotCache(source=_populate(ws, source)))
    pt = PivotTable(cache=cache, location=location, **kw)
    ws.add_pivot_table(pt)
    return wb, ws, pt


def _grid(ws, ref):
    """Values in ``ref`` as rows, with empty cells as ``None``."""
    return [[cell.value for cell in row] for row in ws[ref]]


def _flat(ws, ref):
    """Values in ``ref`` in row-major order."""
    return [value for row in _grid(ws, ref) for value in row]


def _part(path, name):
    with zipfile.ZipFile(path) as z:
        return z.read(name).decode()


def test_nested_rows_and_columns_render_compact():
    _, ws, pt = _pivot(rows=["region", "customer"], cols=["quarter"], data=["revenue"])
    assert pt.location.ref == "F2:J12"
    assert _grid(ws, "F2:J12") == [
        ["Sum of revenue", "Column Labels", None, None, None],
        ["Row Labels", "Q1", "Q2", "Q3", "Grand Total"],
        ["North", 100.0, 150.0, 175.0, 425.0],
        ["Acme", 100.0, None, 175.0, 275.0],
        ["Globex", None, 150.0, None, 150.0],
        ["South", 200.0, 250.0, 225.0, 675.0],
        ["Acme", 200.0, None, None, 200.0],
        ["Globex", None, 250.0, 225.0, 475.0],
        ["East", 50.0, None, None, 50.0],
        ["Globex", 50.0, None, None, 50.0],
        ["Grand Total", 350.0, 400.0, 400.0, 1150.0],
    ]


def test_nested_rows_emit_repeat_counts(tmp_path):
    wb, _, _ = _pivot(rows=["region", "customer"], data=["revenue"])
    path = tmp_path / "pivot.xlsx"
    wb.save(path)
    xml = _part(path, "xl/pivotTables/pivotTable1.xml")
    row_items = re.search(r"<rowItems.*?</rowItems>", xml).group(0)
    assert row_items == (
        '<rowItems count="9">'
        "<i><x/></i>"
        '<i r="1"><x/></i>'
        '<i r="1"><x v="1"/></i>'
        '<i><x v="1"/></i>'
        '<i r="1"><x/></i>'
        '<i r="1"><x v="1"/></i>'
        '<i><x v="2"/></i>'
        '<i r="1"><x v="1"/></i>'
        '<i t="grand"><x/></i>'
        "</rowItems>"
    )
    assert '<colItems count="1"><i/></colItems>' in xml


def test_nested_columns_render_outer_subtotals():
    _, ws, pt = _pivot(rows=["region"], cols=["quarter", "customer"], data=["revenue"])
    assert pt.location.ref == "F2:O8"
    assert _grid(ws, "F2:O8") == [
        ["Sum of revenue", "Column Labels"] + [None] * 8,
        [None, "Q1", None, "Q1 Total", "Q2", "Q2 Total", "Q3", None, "Q3 Total", "Grand Total"],
        ["Row Labels", "Acme", "Globex", None, "Globex", None, "Acme", "Globex", None, None],
        ["North", 100.0, None, 100.0, 150.0, 150.0, 175.0, None, 175.0, 425.0],
        ["South", 200.0, None, 200.0, 250.0, 250.0, None, 225.0, 225.0, 675.0],
        ["East", None, 50.0, 50.0, None, None, None, None, None, 50.0],
        ["Grand Total", 300.0, 50.0, 350.0, 400.0, 400.0, 175.0, 225.0, 400.0, 1150.0],
    ]


def test_multiple_data_fields_add_values_column_field(tmp_path):
    wb, ws, pt = _pivot(rows=["region"], data=["revenue", ("revenue", "count")])
    assert pt.location.ref == "F2:H7"
    assert _grid(ws, "F2:H7") == [
        [None, "Values", None],
        ["Row Labels", "Sum of revenue", "Count of revenue"],
        ["North", 425.0, 3.0],
        ["South", 675.0, 3.0],
        ["East", 50.0, 1.0],
        ["Grand Total", 1150.0, 7.0],
    ]
    path = tmp_path / "pivot.xlsx"
    wb.save(path)
    xml = _part(path, "xl/pivotTables/pivotTable1.xml")
    assert (
        '<colFields count="1"><field x="-2"/></colFields>'
        '<colItems count="2"><i><x/></i><i i="1"><x v="1"/></i></colItems>'
    ) in xml


def test_multiple_data_fields_without_row_fields_drop_label_column():
    _, ws, pt = _pivot(cols=["quarter"], data=["revenue", ("revenue", "count")])
    assert pt.location.ref == "F2:M5"
    assert _grid(ws, "F2:M5") == [
        ["Column Labels"] + [None] * 7,
        ["Q1", None, "Q2", None, "Q3", None, "Total Sum of revenue", "Total Count of revenue"],
        ["Sum of revenue", "Count of revenue"] * 3 + [None, None],
        [350.0, 3.0, 400.0, 2.0, 400.0, 2.0, 1150.0, 7.0],
    ]


def test_page_field_sits_above_table_and_shifts_it_down():
    _, ws, pt = _pivot(rows=["region"], page=["customer"], data=["revenue"])
    # Excel reports TableRange1 F3:G7 and TableRange2 F1:G7.
    assert pt.location.ref == "F3:G7"
    assert (pt.location.row_page_count, pt.location.col_page_count) == (1, 1)
    assert _grid(ws, "F1:G7") == [
        ["customer", "(All)"],
        [None, None],
        ["Row Labels", "Sum of revenue"],
        ["North", 425.0],
        ["South", 675.0],
        ["East", 50.0],
        ["Grand Total", 1150.0],
    ]


def test_selected_page_item_filters_rendered_values():
    _, ws, pt = _pivot(
        rows=["region"], page=[PageField("customer", item_index=0)], data=["revenue"]
    )
    # Excel reports TableRange1 F3:G6 and TableRange2 F1:G6; East has no
    # Acme records, so the filter hides it.
    assert pt.location.ref == "F3:G6"
    assert _grid(ws, "F1:G6") == [
        ["customer", "Acme"],
        [None, None],
        ["Row Labels", "Sum of revenue"],
        ["North", 275.0],
        ["South", 200.0],
        ["Grand Total", 475.0],
    ]


def test_percent_of_total_renders_shares_and_zero_for_empty_cells():
    _, ws, _ = _pivot(
        rows=["region"],
        cols=["quarter"],
        data=[DataField("revenue", show_data_as="percentOfTotal")],
    )
    assert _flat(ws, "G4:J7") == pytest.approx(
        [
            *(0.086956521739, 0.130434782609, 0.152173913043, 0.369565217391),
            *(0.173913043478, 0.217391304348, 0.195652173913, 0.586956521739),
            *(0.04347826087, 0.0, 0.0, 0.04347826087),
            *(0.304347826087, 0.347826086957, 0.347826086957, 1.0),
        ]
    )


def test_index_matches_excel():
    _, ws, _ = _pivot(
        rows=["region"], cols=["quarter"], data=[DataField("revenue", show_data_as="index")]
    )
    assert _flat(ws, "G4:J7") == pytest.approx(
        [
            *(0.773109243697, 1.014705882353, 1.183823529412, 1.0),
            *(0.973544973545, 1.064814814815, 0.958333333333, 1.0),
            *(3.285714285714, 0.0, 0.0, 1.0),
            *(1.0, 1.0, 1.0, 1.0),
        ]
    )


def test_base_relative_show_data_as_is_left_for_excel_to_refresh():
    _, ws, pt = _pivot(
        rows=["region"],
        cols=["quarter"],
        data=[DataField("revenue", show_data_as="runTotal", base_field=1)],
    )
    assert _grid(ws, "F4:J4") == [["North", None, None, None, None]]
    assert pt.cache.refresh_on_load is True


def test_rendering_clears_existing_values_inside_the_pivot_range():
    wb = Workbook()
    ws = wb.active
    ref = _populate(ws, ROWS)
    ws["G6"] = 999  # North / Globex / Q1 has no records.
    ws["F13"] = "below the pivot"
    cache = wb.add_pivot_cache(PivotCache(source=ref))
    ws.add_pivot_table(
        PivotTable(
            cache=cache,
            location="F2",
            rows=["region", "customer"],
            cols=["quarter"],
            data=["revenue"],
        )
    )
    assert ws["G6"].value is None
    assert ws["F13"].value == "below the pivot"



def test_blank_axis_value_renders_blank_item(tmp_path):
    source = ROWS + [(None, "Q2", "Acme", 30.0)]
    wb, ws, _ = _pivot(source=source, rows=["region"], cols=["quarter"], data=["revenue"])
    assert _grid(ws, "F7:J8") == [
        ["(blank)", None, 30.0, None, 30.0],
        ["Grand Total", 350.0, 430.0, 400.0, 1180.0],
    ]
    path = tmp_path / "pivot.xlsx"
    wb.save(path)
    cache_xml = _part(path, "xl/pivotCache/pivotCacheDefinition1.xml")
    assert re.search(
        r'<cacheField name="region"[^>]*><sharedItems containsBlank="1"[^>]* count="4">'
        r'<s v="North"/><s v="South"/><s v="East"/><m/></sharedItems>',
        cache_xml,
    )


def test_axis_field_without_shared_items_is_rejected():
    rows = [("id", "revenue")] + [(float(i), 1.0) for i in range(300)]
    wb = Workbook()
    ws = wb.active
    for row in rows:
        ws.append(row)
    ref = Reference(worksheet=ws, min_col=1, min_row=1, max_col=2, max_row=len(rows))
    cache = wb.add_pivot_cache(PivotCache(source=ref))
    with pytest.raises(ValueError, match="'id' has too many unique values"):
        ws.add_pivot_table(PivotTable(cache=cache, location="D2", rows=["id"], data=["revenue"]))


@pytest.mark.parametrize("mode", ["write", "modify"])
def test_pivot_caches_follow_calc_pr_in_workbook_xml(tmp_path, mode):
    path = tmp_path / "pivot.xlsx"
    if mode == "write":
        wb, _, _ = _pivot(rows=["region"], data=["revenue"])
    else:
        src = Workbook()
        src.active.title = "Data"
        for row in ROWS:
            src.active.append(row)
        src.save(path)
        wb = load_workbook(path, modify=True)
        ws = wb["Data"]
        ref = Reference(worksheet=ws, min_col=1, min_row=1, max_col=4, max_row=len(ROWS))
        cache = wb.add_pivot_cache(PivotCache(source=ref))
        ws.add_pivot_table(PivotTable(cache=cache, location="F2", rows=["region"], data=["revenue"]))
    wb.save(path)

    workbook_xml = _part(path, "xl/workbook.xml")
    assert workbook_xml.index("<calcPr") < workbook_xml.index("<pivotCaches>")
    reloaded = load_workbook(path)["Data"]
    assert _grid(reloaded, "F2:G6") == [
        ["Row Labels", "Sum of revenue"],
        ["North", 425.0],
        ["South", 675.0],
        ["East", 50.0],
        ["Grand Total", 1150.0],
    ]
