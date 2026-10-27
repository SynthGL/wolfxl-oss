"""RFC-061 Sub-feature 3.2 — Calculated fields (~30 tests)."""

from __future__ import annotations

import pytest

from wolfxl.chart.reference import Reference
from wolfxl.pivot import CalculatedField, PivotCache


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
        "A1": "region", "B1": "revenue", "C1": "cost",
        "A2": "North",  "B2": 100.0,     "C2": 60.0,
        "A3": "South",  "B3": 200.0,     "C3": 90.0,
    }
    ws = _StubWorksheet("Sheet1", data)
    src = Reference(worksheet=ws, min_col=1, min_row=1, max_col=3, max_row=3)
    pc = PivotCache(source=src)
    pc._cache_id = 0
    pc._materialize(ws)
    return pc


# ---------------------------------------------------------------------------
# CalculatedField construction validation
# ---------------------------------------------------------------------------


def test_calculated_field_requires_name():
    with pytest.raises(ValueError):
        CalculatedField(name="", formula="= a + b")


def test_calculated_field_requires_formula():
    with pytest.raises(ValueError):
        CalculatedField(name="profit", formula="")


def test_calculated_field_invalid_data_type():
    with pytest.raises(ValueError):
        CalculatedField(name="x", formula="= a + b", data_type="float")


def test_calculated_field_default_data_type_is_number():
    cf = CalculatedField(name="profit", formula="= revenue - cost")
    assert cf.data_type == "number"


def test_calculated_field_supports_string_type():
    cf = CalculatedField(name="x", formula='= "hello"', data_type="string")
    assert cf.data_type == "string"


def test_calculated_field_supports_boolean_type():
    cf = CalculatedField(name="x", formula="= a > b", data_type="boolean")
    assert cf.data_type == "boolean"


def test_calculated_field_supports_date_type():
    cf = CalculatedField(name="x", formula="= today()", data_type="date")
    assert cf.data_type == "date"


# ---------------------------------------------------------------------------
# Formula structural validation
# ---------------------------------------------------------------------------


def test_formula_unbalanced_paren_raises():
    with pytest.raises(ValueError):
        CalculatedField(name="x", formula="= (a + b")


def test_formula_unbalanced_close_paren_raises():
    with pytest.raises(ValueError):
        CalculatedField(name="x", formula="= a + b)")


def test_formula_unterminated_string_raises():
    with pytest.raises(ValueError):
        CalculatedField(name="x", formula='= "hello')


def test_formula_balanced_parens_accepted():
    cf = CalculatedField(name="x", formula="= sum((a + b) * c)")
    assert cf.formula == "= sum((a + b) * c)"


def test_formula_string_literal_accepted():
    cf = CalculatedField(name="x", formula='= "hello world"')
    assert cf.formula == '= "hello world"'


def test_formula_leading_equals_optional():
    cf = CalculatedField(name="x", formula="a + b")
    # Stored as-is, but to_rust_dict strips leading '='.
    assert cf.to_rust_dict()["formula"] == "a + b"


def test_formula_leading_equals_stripped_in_dict():
    cf = CalculatedField(name="x", formula="= a + b")
    assert cf.to_rust_dict()["formula"] == " a + b"


# ---------------------------------------------------------------------------
# to_rust_dict shape
# ---------------------------------------------------------------------------


def test_calculated_field_to_rust_dict_shape():
    cf = CalculatedField(name="profit", formula="= revenue - cost")
    d = cf.to_rust_dict()
    assert d["name"] == "profit"
    assert "formula" in d
    assert d["data_type"] == "number"


# ---------------------------------------------------------------------------
# PivotCache.add_calculated_field()
# ---------------------------------------------------------------------------


def test_pivot_cache_add_calculated_field():
    pc = _materialized_cache()
    cf = pc.add_calculated_field("profit", "= revenue - cost")
    assert isinstance(cf, CalculatedField)
    assert cf.name == "profit"
    assert pc.calculated_fields == [cf]


def test_pivot_cache_add_multiple_calculated_fields():
    pc = _materialized_cache()
    pc.add_calculated_field("profit", "= revenue - cost")
    pc.add_calculated_field("margin", "= profit / revenue")
    assert len(pc.calculated_fields) == 2


def test_pivot_cache_calculated_fields_in_to_rust_dict():
    pc = _materialized_cache()
    pc.add_calculated_field("profit", "= revenue - cost")
    d = pc.to_rust_dict()
    assert "calculated_fields" in d
    assert len(d["calculated_fields"]) == 1
    assert d["calculated_fields"][0]["name"] == "profit"


def test_pivot_cache_with_no_calc_fields_emits_empty_list():
    pc = _materialized_cache()
    d = pc.to_rust_dict()
    assert d["calculated_fields"] == []


# ---------------------------------------------------------------------------
# PyO3 round-trip via serialize_pivot_cache_dict
# ---------------------------------------------------------------------------


def test_serialize_calc_field_as_formula_cache_field():
    from wolfxl._rust import serialize_pivot_cache_dict

    pc = _materialized_cache()
    pc.add_calculated_field("profit", "= revenue - cost")
    s = serialize_pivot_cache_dict(pc.to_rust_dict()).decode()
    assert '<cacheFields count="4">' in s
    assert (
        '<cacheField name="profit" numFmtId="0" formula=" revenue - cost" '
        'databaseField="0"/></cacheFields>'
    ) in s
    assert "<calculatedItems" not in s


def test_calc_field_name_must_be_unique():
    pc = _materialized_cache()
    pc.add_calculated_field("profit", "revenue - cost")
    for name in ("profit", "revenue"):
        with pytest.raises(ValueError):
            pc.add_calculated_field(name, "revenue * 2")


def _save_workbook(tmp_path, fields, **table):
    from wolfxl import Workbook
    from wolfxl.pivot import PivotTable

    wb = Workbook()
    ws = wb.active
    for row in [
        ("region", "revenue", "cost"),
        ("North", 100.0, 40.0),
        ("South", 200.0, 90.0),
        ("North", 150.0, 60.0),
        ("East", 50.0, 10.0),
    ]:
        ws.append(row)
    cache = wb.add_pivot_cache(
        PivotCache(source=Reference(worksheet=ws, min_col=1, min_row=1, max_col=3, max_row=5))
    )
    for name, formula in fields:
        cache.add_calculated_field(name, formula)
    ws.add_pivot_table(PivotTable(cache=cache, location="F2", **table))
    path = tmp_path / "calc_fields.xlsx"
    wb.save(path)
    return path


def test_calc_field_renders_formula_over_field_sums(tmp_path):
    import wolfxl

    path = _save_workbook(
        tmp_path,
        [("profit", "revenue - cost"), ("margin", "profit / revenue")],
        rows=["region"],
        data=["revenue", "profit", "margin"],
    )
    ws = wolfxl.load_workbook(path).active
    grid = [[c.value for c in row] for row in ws["F3:I7"]]
    assert grid[0] == ["Row Labels", "Sum of revenue", "Sum of profit", "Sum of margin"]
    assert grid[1][:3] == ["North", 250, 150]
    assert grid[1][3] == pytest.approx(150 / 250)
    # Totals apply the formula to the summed inputs, not sum the rows.
    assert grid[4][:3] == ["Grand Total", 500, 300]
    assert grid[4][3] == pytest.approx(300 / 500)


def test_calc_field_cannot_be_an_axis_field(tmp_path):
    with pytest.raises(ValueError, match="data field"):
        _save_workbook(
            tmp_path, [("profit", "revenue - cost")], rows=["profit"], data=["revenue"]
        )


@pytest.mark.parametrize("bad", [float("nan"), float("inf"), float("-inf")])
def test_serialize_pivot_cache_rejects_non_finite_shared_item_bounds(bad):
    from wolfxl._rust import serialize_pivot_cache_dict

    pc = _materialized_cache()
    d = pc.to_rust_dict()
    d["fields"][1]["shared_items"]["min_value"] = bad
    with pytest.raises(ValueError, match="non-finite floats"):
        serialize_pivot_cache_dict(d)


@pytest.mark.parametrize("bad", [float("nan"), float("inf"), float("-inf")])
def test_serialize_pivot_records_rejects_non_finite_numeric_values(bad):
    from wolfxl._rust import serialize_pivot_records_dict

    pc = _materialized_cache()
    cache_d = pc.to_rust_dict()
    records_d = pc.to_rust_records_dict()
    records_d["records"][0][1] = {"kind": "number", "value": bad}
    with pytest.raises(ValueError, match="non-finite floats"):
        serialize_pivot_records_dict(cache_d, records_d)


# ---------------------------------------------------------------------------
# Edge cases
# ---------------------------------------------------------------------------


def test_calc_field_with_complex_formula():
    cf = CalculatedField(name="x", formula="= IF(a > 0, b * 0.1, 0)")
    assert cf.name == "x"


def test_calc_field_formula_with_spaces():
    cf = CalculatedField(name="x", formula="=    a    +    b   ")
    assert cf.formula == "=    a    +    b   "


def test_calc_field_formula_with_quotes():
    cf = CalculatedField(name="x", formula='= IF(a > 0, "yes", "no")')
    assert "yes" in cf.formula


def test_calc_field_unicode_in_name():
    cf = CalculatedField(name="πrofit", formula="= a + b")
    assert cf.name == "πrofit"


def test_pivot_cache_calc_field_chained_to_rust_dict():
    pc = _materialized_cache()
    pc.add_calculated_field("a", "x + 1")
    pc.add_calculated_field("b", "y * 2")
    pc.add_calculated_field("c", "a - b")
    d = pc.to_rust_dict()
    names = [c["name"] for c in d["calculated_fields"]]
    assert names == ["a", "b", "c"]


def test_pivot_cache_calc_field_keeps_data_type():
    pc = _materialized_cache()
    pc.add_calculated_field("flag", "= a > b", data_type="boolean")
    d = pc.to_rust_dict()
    assert d["calculated_fields"][0]["data_type"] == "boolean"


def test_pivot_cache_calc_field_returns_field_for_chaining():
    pc = _materialized_cache()
    cf = pc.add_calculated_field("profit", "= a - b")
    assert cf in pc.calculated_fields


def test_serialize_pivot_cache_calc_field_escapes_xml():
    from wolfxl._rust import serialize_pivot_cache_dict

    pc = _materialized_cache()
    pc.add_calculated_field("x", "a < b & c")
    d = pc.to_rust_dict()
    xml = serialize_pivot_cache_dict(d)
    s = xml.decode()
    assert "&amp;" in s
    assert "&lt;" in s


def test_calc_field_dict_preserves_formula_after_strip():
    cf = CalculatedField(name="x", formula="= a + b")
    d = cf.to_rust_dict()
    # leading '=' stripped
    assert not d["formula"].startswith("=")


def test_calc_field_no_strip_when_no_equals():
    cf = CalculatedField(name="x", formula="a+b")
    assert cf.to_rust_dict()["formula"] == "a+b"
