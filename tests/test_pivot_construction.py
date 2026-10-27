"""Sprint Ν Pod-β tests — Python pivot construction surface.

Validates the public API of ``wolfxl.pivot.PivotCache``,
``wolfxl.pivot.PivotTable``, and the ``to_rust_dict()`` shape per
RFC-047 §10 / RFC-048 §10.

These tests do NOT require the full Rust→Python end-to-end emit
path to work — they only validate the Python construction surface
and the §10 dict shape Pod-α's Rust parser will consume. End-to-end
emit tests live in ``tests/diffwriter/test_pivot_*.py`` (Pod-γ
work, post-merge).
"""

from __future__ import annotations

import pytest

from wolfxl.chart.reference import Reference
from wolfxl.pivot import (
    CacheValue,
    ColumnField,
    DataField,
    PageField,
    PivotCache,
    PivotSource,
    PivotTable,
    RowField,
)


# ---------------------------------------------------------------------------
# Test fixtures — minimal worksheet stub
# ---------------------------------------------------------------------------


class _StubCell:
    def __init__(self, value):
        self.value = value


class _StubWorksheet:
    """Minimal worksheet that supports ``ws[addr]`` cell lookup.

    Used to avoid pulling the full Workbook stack into pivot
    construction tests.
    """

    def __init__(self, title, data):
        self.title = title
        self._data = data  # dict[address_str, value]

    def __getitem__(self, addr):
        return _StubCell(self._data.get(addr))


def _build_sample_worksheet():
    """4 columns × 5 rows (1 header + 4 data) — region/quarter/customer/revenue."""
    data = {
        "A1": "region", "B1": "quarter", "C1": "customer", "D1": "revenue",
        "A2": "North",  "B2": "Q1",      "C2": "Acme",     "D2": 100.0,
        "A3": "South",  "B3": "Q1",      "C3": "Acme",     "D3": 200.0,
        "A4": "North",  "B4": "Q2",      "C4": "Globex",   "D4": 150.0,
        "A5": "South",  "B5": "Q2",      "C5": "Globex",   "D5": 250.0,
    }
    return _StubWorksheet("Sheet1", data)


# ---------------------------------------------------------------------------
# CacheValue — RFC-047 §10.5
# ---------------------------------------------------------------------------


def test_cache_value_constructors():
    assert CacheValue.string("x").to_rust_dict() == {"kind": "string", "value": "x"}
    assert CacheValue.number(1.5).to_rust_dict() == {"kind": "number", "value": 1.5}
    assert CacheValue.boolean(True).to_rust_dict() == {"kind": "boolean", "value": True}
    assert CacheValue.missing().to_rust_dict() == {"kind": "missing"}
    assert CacheValue.error("#REF!").to_rust_dict() == {"kind": "error", "value": "#REF!"}


@pytest.mark.parametrize("value", [float("nan"), float("inf"), -float("inf")])
def test_cache_value_number_rejects_non_finite(value):
    with pytest.raises(ValueError, match="non-finite floats"):
        CacheValue.number(value)


@pytest.mark.parametrize("value", [float("nan"), float("inf"), -float("inf")])
def test_pivot_cache_materialize_rejects_non_finite_numbers(value):
    data = {
        "A1": "region", "B1": "revenue",
        "A2": "North", "B2": value,
    }
    ws = _StubWorksheet("Sheet1", data)
    src = Reference(worksheet=ws, min_col=1, min_row=1, max_col=2, max_row=2)
    pc = PivotCache(source=src)
    pc._cache_id = 0

    with pytest.raises(ValueError, match="non-finite floats"):
        pc._materialize(ws)


def test_cache_value_date_normalizes_to_iso():
    from datetime import date
    cv = CacheValue.date(date(2026, 1, 15))
    assert cv.value == "2026-01-15T00:00:00"


def test_cache_value_eq_and_hash():
    a = CacheValue.string("x")
    b = CacheValue.string("x")
    c = CacheValue.string("y")
    assert a == b
    assert a != c
    assert hash(a) == hash(b)


# ---------------------------------------------------------------------------
# PivotCache — RFC-047 §10
# ---------------------------------------------------------------------------


def test_pivot_cache_requires_reference_source():
    with pytest.raises(TypeError):
        PivotCache(source="A1:D5")


def test_pivot_cache_unmaterialized_raises_on_to_rust_dict():
    ws = _build_sample_worksheet()
    src = Reference(worksheet=ws, min_col=1, min_row=1, max_col=4, max_row=5)
    pc = PivotCache(source=src)
    with pytest.raises(RuntimeError):
        pc.to_rust_dict()
    with pytest.raises(RuntimeError):
        pc.fields  # property access raises until materialized


def test_pivot_cache_materializes_string_field_with_enumeration():
    ws = _build_sample_worksheet()
    src = Reference(worksheet=ws, min_col=1, min_row=1, max_col=4, max_row=5)
    pc = PivotCache(source=src)
    pc._cache_id = 0
    pc._materialize(ws)

    assert len(pc.fields) == 4
    region = pc.fields[0]
    assert region.name == "region"
    assert region.data_type == "string"
    assert region.shared_items.items is not None
    item_values = [v.value for v in region.shared_items.items]
    # Insertion-ordered unique values.
    assert set(item_values) == {"North", "South"}


def test_pivot_cache_materializes_numeric_field_enumerates():
    ws = _build_sample_worksheet()
    src = Reference(worksheet=ws, min_col=1, min_row=1, max_col=4, max_row=5)
    pc = PivotCache(source=src)
    pc._cache_id = 0
    pc._materialize(ws)

    revenue = pc.fields[3]
    assert revenue.name == "revenue"
    assert revenue.data_type == "number"
    si = revenue.shared_items
    # 4 unique values < 200 threshold → enumerate.
    assert si.items is not None
    assert si.contains_number is True
    assert si.min_value == 100.0
    assert si.max_value == 250.0


def test_pivot_cache_to_rust_dict_shape():
    ws = _build_sample_worksheet()
    src = Reference(worksheet=ws, min_col=1, min_row=1, max_col=4, max_row=5)
    pc = PivotCache(source=src)
    pc._cache_id = 7
    pc._materialize(ws)

    d = pc.to_rust_dict()
    # RFC-047 §10.1 required keys.
    assert d["cache_id"] == 7
    assert d["refresh_on_load"] is False
    assert d["refreshed_by"] == "wolfxl"
    assert d["created_version"] == 6
    assert d["min_refreshable_version"] == 3
    assert "source" in d
    assert d["source"]["sheet"] == "Sheet1"
    assert d["source"]["ref"] == "A1:D5"
    assert d["source"]["name"] is None
    assert "fields" in d
    assert len(d["fields"]) == 4
    # First field shape (RFC-047 §10.3).
    region = d["fields"][0]
    assert region["name"] == "region"
    assert region["num_fmt_id"] == 0
    assert region["data_type"] == "string"
    assert region["formula"] is None
    assert region["hierarchy"] is None
    assert "shared_items" in region


def test_pivot_cache_records_dict_indexes_when_enumerated():
    ws = _build_sample_worksheet()
    src = Reference(worksheet=ws, min_col=1, min_row=1, max_col=4, max_row=5)
    pc = PivotCache(source=src)
    pc._cache_id = 0
    pc._materialize(ws)

    d = pc.to_rust_records_dict()
    assert d["field_count"] == 4
    assert d["record_count"] == 4
    assert len(d["records"]) == 4

    # Each record cell for region/quarter/customer should be `index`
    # form (those are enumerated string fields).
    for row in d["records"]:
        # 4 cells per row (one per cache field).
        assert len(row) == 4
        # region / quarter / customer all index-form.
        for cell in row[:3]:
            assert cell["kind"] == "index"
            assert isinstance(cell["value"], int)
        # revenue is a number — could be index OR inline (depends on
        # whether the 4 unique numeric values were enumerated).
        assert row[3]["kind"] in ("index", "number")


def test_pivot_cache_unique_field_check_via_validate():
    """Cache field uniqueness is enforced server-side (Rust crate), but the
    Python materializer should never produce duplicates from a normal
    worksheet."""
    ws = _build_sample_worksheet()
    src = Reference(worksheet=ws, min_col=1, min_row=1, max_col=4, max_row=5)
    pc = PivotCache(source=src)
    pc._cache_id = 0
    pc._materialize(ws)
    names = [f["name"] for f in pc.to_rust_dict()["fields"]]
    assert len(names) == len(set(names))


def test_pivot_cache_empty_source_rejects():
    ws = _StubWorksheet("Empty", {})
    src = Reference(worksheet=ws, min_col=1, min_row=1, max_col=2, max_row=1)
    pc = PivotCache(source=src)
    pc._cache_id = 0
    with pytest.raises(ValueError):
        pc._materialize(ws)


def test_pivot_cache_only_header_rejects():
    ws = _StubWorksheet("OnlyHeader", {"A1": "region", "B1": "qtr"})
    src = Reference(worksheet=ws, min_col=1, min_row=1, max_col=2, max_row=1)
    pc = PivotCache(source=src)
    pc._cache_id = 0
    with pytest.raises(ValueError):
        pc._materialize(ws)


# ---------------------------------------------------------------------------
# DataField — RFC-048 §10.5
# ---------------------------------------------------------------------------


def test_data_field_validates_function():
    with pytest.raises(ValueError):
        DataField(name="x", function="bogus_aggregator")


def test_data_field_resolved_display_name_default():
    df = DataField(name="revenue", function="sum")
    assert df.resolved_display_name() == "Sum of revenue"


def test_data_field_resolved_display_name_explicit():
    df = DataField(name="revenue", function="sum", display_name="Total $")
    assert df.resolved_display_name() == "Total $"


def test_data_field_average_display_name():
    df = DataField(name="revenue", function="average")
    assert df.resolved_display_name() == "Average of revenue"


# ---------------------------------------------------------------------------
# PivotTable — RFC-048 §10
# ---------------------------------------------------------------------------


def _build_sample_pivot_cache():
    ws = _build_sample_worksheet()
    src = Reference(worksheet=ws, min_col=1, min_row=1, max_col=4, max_row=5)
    pc = PivotCache(source=src)
    pc._cache_id = 0
    pc._materialize(ws)
    return pc


def test_pivot_table_requires_data_field():
    pc = _build_sample_pivot_cache()
    with pytest.raises(ValueError, match="≥1 data field"):
        PivotTable(cache=pc, location="F2", rows=["region"])


def test_pivot_table_rejects_non_pivot_cache_for_cache_arg():
    with pytest.raises(TypeError):
        PivotTable(cache="not-a-cache", location="F2", data=["revenue"])


def test_pivot_table_basic_construction():
    pc = _build_sample_pivot_cache()
    pt = PivotTable(
        cache=pc,
        location="F2",
        rows=["region"],
        cols=["quarter"],
        data=[DataField(name="revenue", function="sum")],
    )
    assert pt.name == "PivotTable1"
    assert len(pt.rows) == 1
    assert isinstance(pt.rows[0], RowField)
    assert pt.rows[0].name == "region"
    assert len(pt.data) == 1


def test_pivot_table_string_field_specs_normalized():
    pc = _build_sample_pivot_cache()
    pt = PivotTable(
        cache=pc,
        location="F2",
        rows=["region"],
        cols=["quarter"],
        data=["revenue"],
    )
    assert isinstance(pt.rows[0], RowField)
    assert isinstance(pt.cols[0], ColumnField)
    assert isinstance(pt.data[0], DataField)
    assert pt.data[0].function == "sum"  # default


def test_pivot_table_tuple_data_spec_normalized():
    pc = _build_sample_pivot_cache()
    pt = PivotTable(
        cache=pc,
        location="F2",
        rows=["region"],
        data=[("revenue", "average")],
    )
    assert pt.data[0].name == "revenue"
    assert pt.data[0].function == "average"


def test_pivot_table_unknown_field_raises():
    pc = _build_sample_pivot_cache()
    pt = PivotTable(
        cache=pc,
        location="F2",
        rows=["this_field_does_not_exist"],
        data=["revenue"],
    )
    # Layout computation resolves field names — error surfaces there.
    with pytest.raises(KeyError):
        pt.to_rust_dict()


def test_pivot_table_dual_axis_rejects():
    pc = _build_sample_pivot_cache()
    pt = PivotTable(
        cache=pc,
        location="F2",
        rows=["region"],
        cols=["region"],  # same field on rows and cols
        data=["revenue"],
    )
    with pytest.raises(ValueError, match="multiple axes"):
        pt.to_rust_dict()


def test_pivot_table_to_rust_dict_top_level_keys():
    pc = _build_sample_pivot_cache()
    pt = PivotTable(
        cache=pc,
        location="F2",
        rows=["region"],
        cols=["quarter"],
        data=[DataField(name="revenue", function="sum")],
    )
    d = pt.to_rust_dict()

    # RFC-048 §10.1 required keys.
    required_keys = {
        "name", "cache_id", "location", "pivot_fields",
        "row_field_indices", "col_field_indices", "page_fields",
        "data_fields", "row_items", "col_items",
        "data_on_rows", "outline", "compact",
        "row_grand_totals", "col_grand_totals",
        "data_caption", "grand_total_caption", "error_caption",
        "missing_caption",
        "apply_number_formats", "apply_border_formats",
        "apply_font_formats", "apply_pattern_formats",
        "apply_alignment_formats", "apply_width_height_formats",
        "style_info",
        "created_version", "updated_version", "min_refreshable_version",
    }
    assert required_keys.issubset(d.keys())

    assert d["name"] == "PivotTable1"
    assert d["cache_id"] == 0
    assert d["data_caption"] == "Values"
    assert d["created_version"] == 6
    assert d["data_on_rows"] is False


def test_pivot_table_pivot_fields_count_matches_cache():
    pc = _build_sample_pivot_cache()
    pt = PivotTable(
        cache=pc, location="F2", rows=["region"], data=["revenue"]
    )
    d = pt.to_rust_dict()
    assert len(d["pivot_fields"]) == len(pc.fields)


def test_pivot_table_row_indices_resolve_correctly():
    pc = _build_sample_pivot_cache()
    pt = PivotTable(
        cache=pc, location="F2", rows=["region"], data=["revenue"]
    )
    d = pt.to_rust_dict()
    # `region` is the 0th cache field.
    assert d["row_field_indices"] == [0]
    # Pivot field 0 should have axisRow.
    assert d["pivot_fields"][0]["axis"] == "axisRow"
    # Pivot field 3 (revenue) should have data_field=True.
    assert d["pivot_fields"][3]["data_field"] is True


def test_pivot_table_data_field_dict_shape():
    pc = _build_sample_pivot_cache()
    pt = PivotTable(
        cache=pc,
        location="F2",
        rows=["region"],
        data=[DataField(name="revenue", function="sum")],
    )
    d = pt.to_rust_dict()
    assert len(d["data_fields"]) == 1
    df = d["data_fields"][0]
    assert df["name"] == "Sum of revenue"
    assert df["field_index"] == 3
    assert df["function"] == "sum"
    assert df["base_field"] == 0
    assert df["base_item"] == 0


def test_pivot_table_row_items_grand_total_appended():
    pc = _build_sample_pivot_cache()
    pt = PivotTable(
        cache=pc, location="F2", rows=["region"], data=["revenue"],
        row_grand_totals=True,
    )
    d = pt.to_rust_dict()
    # 2 unique regions + 1 grand total.
    grand = [it for it in d["row_items"] if it.get("t") == "grand"]
    assert len(grand) == 1


def test_pivot_table_no_grand_totals_omits_grand_row():
    pc = _build_sample_pivot_cache()
    pt = PivotTable(
        cache=pc, location="F2", rows=["region"], data=["revenue"],
        row_grand_totals=False,
    )
    d = pt.to_rust_dict()
    grand = [it for it in d["row_items"] if it.get("t") == "grand"]
    assert len(grand) == 0


def test_pivot_table_location_widens_from_anchor():
    pc = _build_sample_pivot_cache()
    pt = PivotTable(
        cache=pc, location="F2", rows=["region"], cols=["quarter"],
        data=["revenue"],
    )
    d = pt.to_rust_dict()
    # Caption row + quarter row + North/South + grand; label + Q1/Q2 + grand.
    assert d["location"]["ref"] == "F2:I6"
    assert d["location"]["first_data_row"] == 2
    assert d["location"]["first_data_col"] == 1


def test_pivot_table_location_range_resized_from_top_left():
    pc = _build_sample_pivot_cache()
    pt = PivotTable(
        cache=pc, location="F2:Z100", rows=["region"], data=["revenue"]
    )
    d = pt.to_rust_dict()
    assert d["location"]["ref"] == "F2:G5"


def test_pivot_table_aggregation_sum():
    pc = _build_sample_pivot_cache()
    pt = PivotTable(
        cache=pc, location="F2", rows=["region"], data=[("revenue", "sum")]
    )
    pt._compute_layout()
    cells = pt._rendered_cells
    assert [cells[(r, 6)] for r in range(2, 6)] == [
        "Row Labels", "North", "South", "Grand Total",
    ]
    assert [cells[(r, 7)] for r in range(3, 6)] == [250.0, 450.0, 700.0]


def test_pivot_table_multi_aggregator():
    pc = _build_sample_pivot_cache()
    pt = PivotTable(
        cache=pc,
        location="F2",
        rows=["region"],
        data=[
            DataField(name="revenue", function="sum"),
            DataField(name="revenue", function="average"),
        ],
    )
    d = pt.to_rust_dict()
    assert len(d["data_fields"]) == 2
    assert d["data_fields"][0]["function"] == "sum"
    assert d["data_fields"][1]["function"] == "average"


def test_pivot_table_page_field_construction():
    pc = _build_sample_pivot_cache()
    pt = PivotTable(
        cache=pc,
        location="F2",
        rows=["region"],
        data=["revenue"],
        page=[PageField(name="customer")],
    )
    d = pt.to_rust_dict()
    assert len(d["page_fields"]) == 1
    pf = d["page_fields"][0]
    assert pf["field_index"] == 2  # customer is 0-indexed col 2
    assert pf["item_index"] == -1  # default "(All)"
    assert pf["hier"] == -1


def test_pivot_table_style_name_propagated():
    pc = _build_sample_pivot_cache()
    pt = PivotTable(
        cache=pc, location="F2", rows=["region"], data=["revenue"],
        style_name="PivotStyleMedium9",
    )
    d = pt.to_rust_dict()
    assert d["style_info"]["name"] == "PivotStyleMedium9"


# ---------------------------------------------------------------------------
# PivotSource (RFC-049 §10)
# ---------------------------------------------------------------------------


def test_pivot_source_construction():
    ps = PivotSource(name="MyPivot")
    assert ps.fmt_id == 0
    assert ps.to_rust_dict() == {"name": "MyPivot", "fmt_id": 0}


def test_pivot_source_fmt_id_validation():
    with pytest.raises(ValueError):
        PivotSource(name="X", fmt_id=-1)
    with pytest.raises(ValueError):
        PivotSource(name="X", fmt_id=70000)


def test_pivot_source_qualified_name():
    ps = PivotSource(name="Sheet1!MyPivot", fmt_id=42)
    assert ps.to_rust_dict() == {"name": "Sheet1!MyPivot", "fmt_id": 42}


# ---------------------------------------------------------------------------
# Stubs are dropped — sanity check the module no longer raises on import.
# ---------------------------------------------------------------------------


def test_pivot_module_no_longer_stub():
    """Sprint Ν acceptance gate: PivotTable construction does NOT raise
    NotImplementedError (which the v0.5+ _make_stub did)."""
    pc = _build_sample_pivot_cache()
    # Must not raise NotImplementedError.
    pt = PivotTable(
        cache=pc, location="F2", rows=["region"], data=["revenue"]
    )
    assert pt is not None


def test_pivot_compat_re_exports():
    """Sanity: ``from wolfxl.pivot import PivotTable`` is the v0.5+
    compat path; v2.0 swaps it from a stub to a real class."""
    from wolfxl.pivot import PivotTable as P
    pc = _build_sample_pivot_cache()
    instance = P(cache=pc, location="F2", rows=["region"], data=["revenue"])
    # Has the openpyxl-shaped attrs.
    assert hasattr(instance, "name")
    assert hasattr(instance, "cache")
    assert hasattr(instance, "to_rust_dict")
