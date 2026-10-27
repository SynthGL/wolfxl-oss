"""RFC-061 Sub-feature 3.5 — Pivot styling beyond the named-style picker.

Covers PivotArea selectors, table-scoped Format directives, pivot CFs,
and the ChartFormat stub for pivot charts.
"""

from __future__ import annotations

import pytest

from wolfxl.chart.reference import Reference
from wolfxl.pivot import (
    ChartFormat,
    Format,
    PivotArea,
    PivotCache,
    PivotConditionalFormat,
    PivotTable,
)


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------


class _StubCell:
    def __init__(self, v):
        self.value = v


class _StubWorksheet:
    def __init__(self, title, data):
        self.title = title
        self._data = data

    def __getitem__(self, addr):
        return _StubCell(self._data.get(addr))


def _materialized_cache():
    data = {
        "A1": "region", "B1": "revenue",
        "A2": "east",   "B2": 100.0,
        "A3": "west",   "B3": 200.0,
    }
    ws = _StubWorksheet("Sheet1", data)
    src = Reference(worksheet=ws, min_col=1, min_row=1, max_col=2, max_row=3)
    pc = PivotCache(source=src)
    pc._cache_id = 0
    pc._materialize(ws)
    return pc


def _pivot_table(cache):
    return PivotTable(
        cache=cache,
        location="F2",
        rows=["region"],
        data=[("revenue", "sum")],
    )


# ---------------------------------------------------------------------------
# PivotArea
# ---------------------------------------------------------------------------


def test_pivot_area_default_is_data():
    pa = PivotArea()
    assert pa.type == "data"
    assert pa.data_only is True
    assert pa.field is None


def test_pivot_area_invalid_type():
    with pytest.raises(ValueError):
        PivotArea(type="bogus")


def test_pivot_area_targeting_specific_field():
    pa = PivotArea(field=2, type="data", data_only=True)
    d = pa.to_rust_dict()
    assert d["field"] == 2
    assert d["type"] == "data"


def test_pivot_area_grand_row():
    pa = PivotArea(type="all", grand_row=True)
    d = pa.to_rust_dict()
    assert d["grand_row"] is True


def test_pivot_area_button_label():
    pa = PivotArea(type="button", field=0)
    d = pa.to_rust_dict()
    assert d["type"] == "button"


def test_pivot_area_to_rust_dict_full_shape():
    pa = PivotArea(
        field=1,
        type="data",
        data_only=True,
        label_only=False,
        grand_row=False,
        grand_col=True,
        cache_index=0,
        axis="axisRow",
        field_position=2,
    )
    d = pa.to_rust_dict()
    assert d["field"] == 1
    assert d["type"] == "data"
    assert d["axis"] == "axisRow"
    assert d["field_position"] == 2
    assert d["grand_col"] is True


# ---------------------------------------------------------------------------
# Format
# ---------------------------------------------------------------------------


def test_format_invalid_action():
    pa = PivotArea()
    with pytest.raises(ValueError):
        Format(pivot_area=pa, dxf_id=0, action="bogus")


def test_format_negative_dxf_id_rejected():
    pa = PivotArea()
    with pytest.raises(ValueError):
        Format(pivot_area=pa, dxf_id=-2)


def test_format_default_action_formatting():
    pa = PivotArea()
    f = Format(pivot_area=pa, dxf_id=3)
    assert f.action == "formatting"


def test_format_blank_action():
    pa = PivotArea()
    f = Format(pivot_area=pa, dxf_id=0, action="blank")
    assert f.action == "blank"


def test_format_to_rust_dict():
    pa = PivotArea(field=0, type="data")
    f = Format(pivot_area=pa, dxf_id=5, action="formatting")
    d = f.to_rust_dict()
    assert d == {
        "action": "formatting",
        "dxf_id": 5,
        "pivot_area": pa.to_rust_dict(),
    }


# ---------------------------------------------------------------------------
# PivotConditionalFormat
# ---------------------------------------------------------------------------


def test_pcf_requires_at_least_one_pivot_area():
    with pytest.raises(ValueError):
        PivotConditionalFormat(rule=object(), pivot_areas=[])


def test_pcf_to_rust_dict_with_object_rule():
    class _R:
        def to_rust_dict(self):
            return {"type": "colorScale", "min_color": "FF0000"}

    pa = PivotArea(field=0, type="data")
    cf = PivotConditionalFormat(rule=_R(), pivot_areas=[pa], priority=2)
    d = cf.to_rust_dict()
    assert d["rule"]["type"] == "colorScale"
    assert d["priority"] == 2
    assert d["scope"] == "data"
    assert len(d["pivot_areas"]) == 1


def test_pcf_to_rust_dict_with_dataclass_like_rule():
    class _R:
        def __init__(self):
            self.type = "cellIs"
            self.operator = "greaterThan"
            self.value = 100

    pa = PivotArea()
    cf = PivotConditionalFormat(rule=_R(), pivot_areas=[pa])
    d = cf.to_rust_dict()
    assert d["rule"]["operator"] == "greaterThan"


def test_pcf_default_scope_data():
    pa = PivotArea()
    cf = PivotConditionalFormat(rule=object(), pivot_areas=[pa])
    assert cf.scope == "data"


# ---------------------------------------------------------------------------
# ChartFormat stub
# ---------------------------------------------------------------------------


def test_chart_format_constructs():
    cf = ChartFormat(chart_index=0, series_index=1, formatting={"line": "thick"})
    d = cf.to_rust_dict()
    assert d["chart_index"] == 0
    assert d["series_index"] == 1
    assert d["formatting"] == {"line": "thick"}


def test_chart_format_default_empty_formatting():
    cf = ChartFormat(chart_index=2, series_index=0)
    assert cf.to_rust_dict()["formatting"] == {}


# ---------------------------------------------------------------------------
# PivotTable.add_format
# ---------------------------------------------------------------------------


def test_pt_add_format_requires_pivot_area():
    pc = _materialized_cache()
    pt = _pivot_table(pc)
    with pytest.raises(TypeError):
        pt.add_format("not-a-pivot-area", dxf_id=0)


def test_pt_add_format_appends_with_explicit_dxf_id():
    pc = _materialized_cache()
    pt = _pivot_table(pc)
    pa = PivotArea(field=0, type="data")
    f = pt.add_format(pa, dxf_id=4)
    assert f.dxf_id == 4
    assert pt.formats == [f]


def test_pt_add_format_uses_sentinel_when_no_dxf_id():
    pc = _materialized_cache()
    pt = _pivot_table(pc)
    pa = PivotArea(field=0)
    f = pt.add_format(pa, dxf=object())
    # Sentinel value; the patcher will allocate at flush time.
    assert f.dxf_id == -1
    assert hasattr(f, "_dxf_payload")


def test_pt_add_format_blank_action():
    pc = _materialized_cache()
    pt = _pivot_table(pc)
    pa = PivotArea(type="button", field=0)
    f = pt.add_format(pa, dxf_id=0, action="blank")
    assert f.action == "blank"


def test_pt_add_format_in_rust_dict():
    pc = _materialized_cache()
    pt = _pivot_table(pc)
    pa = PivotArea(field=0, type="data")
    pt.add_format(pa, dxf_id=2)
    d = pt.to_rust_dict()
    assert "formats" in d
    assert len(d["formats"]) == 1
    assert d["formats"][0]["dxf_id"] == 2


# ---------------------------------------------------------------------------
# PivotTable.add_conditional_format
# ---------------------------------------------------------------------------


def test_pt_add_cf_appends():
    class _R:
        def to_rust_dict(self):
            return {"type": "colorScale"}

    pc = _materialized_cache()
    pt = _pivot_table(pc)
    pa = PivotArea(field=0, type="data")
    pcf = pt.add_conditional_format(_R(), pa)
    assert pt.conditional_formats == [pcf]


def test_pt_add_cf_accepts_list_of_areas():
    class _R:
        def to_rust_dict(self):
            return {"type": "cellIs"}

    pc = _materialized_cache()
    pt = _pivot_table(pc)
    pa1 = PivotArea(field=0, type="data")
    pa2 = PivotArea(type="all", grand_row=True)
    pcf = pt.add_conditional_format(_R(), [pa1, pa2])
    assert len(pcf.pivot_areas) == 2


def test_pt_add_cf_rejects_non_pivot_area_in_list():
    pc = _materialized_cache()
    pt = _pivot_table(pc)
    with pytest.raises(TypeError):
        pt.add_conditional_format(object(), [PivotArea(), "not-a-pivot-area"])


def test_pt_add_cf_in_rust_dict():
    class _R:
        def to_rust_dict(self):
            return {"type": "dataBar"}

    pc = _materialized_cache()
    pt = _pivot_table(pc)
    pa = PivotArea(field=0, type="data")
    pt.add_conditional_format(_R(), pa, priority=3)
    d = pt.to_rust_dict()
    assert "conditional_formats" in d
    assert d["conditional_formats"][0]["priority"] == 3
    assert d["conditional_formats"][0]["rule"]["type"] == "dataBar"


# ---------------------------------------------------------------------------
# Rust round-trip
# ---------------------------------------------------------------------------


def test_serialize_pivot_table_with_format():
    from wolfxl._rust import serialize_pivot_table_dict

    pc = _materialized_cache()
    pt = _pivot_table(pc)
    pa = PivotArea(field=0, type="data")
    pt.add_format(pa, dxf_id=7, action="formatting")
    cache_d = pc.to_rust_dict()
    table_d = pt.to_rust_dict()
    xml = serialize_pivot_table_dict(cache_d, table_d)
    s = xml.decode()
    assert "<formats" in s
    assert 'dxfId="7"' in s


def test_serialize_pivot_table_with_cf():
    from wolfxl._rust import serialize_pivot_table_dict

    class _R:
        def to_rust_dict(self):
            return {"type": "colorScale", "priority": 1}

    pc = _materialized_cache()
    pt = _pivot_table(pc)
    pa = PivotArea(field=0, type="data")
    pt.add_conditional_format(_R(), pa, priority=5)
    cache_d = pc.to_rust_dict()
    table_d = pt.to_rust_dict()
    xml = serialize_pivot_table_dict(cache_d, table_d)
    s = xml.decode()
    assert "<conditionalFormats" in s
    assert 'priority="5"' in s


def test_serialize_pivot_table_no_formats_no_block():
    from wolfxl._rust import serialize_pivot_table_dict

    pc = _materialized_cache()
    pt = _pivot_table(pc)
    cache_d = pc.to_rust_dict()
    table_d = pt.to_rust_dict()
    xml = serialize_pivot_table_dict(cache_d, table_d)
    s = xml.decode()
    assert "<formats" not in s
    assert "<conditionalFormats" not in s


def test_serialize_pivot_table_format_byte_stable():
    from wolfxl._rust import serialize_pivot_table_dict

    pc = _materialized_cache()
    pt = _pivot_table(pc)
    pa = PivotArea(field=0, type="data")
    pt.add_format(pa, dxf_id=3)
    cache_d = pc.to_rust_dict()
    table_d = pt.to_rust_dict()
    a = serialize_pivot_table_dict(cache_d, table_d)
    b = serialize_pivot_table_dict(cache_d, table_d)
    assert a == b


def test_serialize_pivot_table_multiple_formats_count():
    from wolfxl._rust import serialize_pivot_table_dict

    pc = _materialized_cache()
    pt = _pivot_table(pc)
    pt.add_format(PivotArea(field=0, type="data"), dxf_id=1)
    pt.add_format(PivotArea(field=0, type="all", grand_row=True), dxf_id=2)
    pt.add_format(PivotArea(type="button", field=0), dxf_id=3, action="blank")
    cache_d = pc.to_rust_dict()
    table_d = pt.to_rust_dict()
    xml = serialize_pivot_table_dict(cache_d, table_d)
    s = xml.decode()
    assert s.count("<format ") == 3


def test_pt_styling_combined_with_calc_items():
    """Calc items + Format + CF can coexist on the same pivot table."""
    from wolfxl._rust import serialize_pivot_table_dict

    class _R:
        def to_rust_dict(self):
            return {"type": "iconSet"}

    pc = _materialized_cache()
    pt = _pivot_table(pc)
    pt.add_calculated_item("region", "ew", "= east + west")
    pt.add_format(PivotArea(field=0, type="data"), dxf_id=1)
    pt.add_conditional_format(_R(), PivotArea(field=0, type="data"))
    cache_d = pc.to_rust_dict()
    table_d = pt.to_rust_dict()
    xml = serialize_pivot_table_dict(cache_d, table_d)
    s = xml.decode()
    assert '<item x="2" f="1"/>' in s
    assert "<formats" in s
    assert "<conditionalFormats" in s
