"""Tests for wolfxl.calc function registry and builtins."""

from __future__ import annotations

from typing import Any

import pytest
from wolfxl.calc._functions import (
    _BUILTINS,
    FUNCTION_WHITELIST_V1,
    ExcelError,
    FunctionRegistry,
    RangeValue,
    first_error,
    is_error,
    is_supported,
)


class TestWhitelist:
    def test_whitelist_has_67_functions(self) -> None:
        assert len(FUNCTION_WHITELIST_V1) == 67

    def test_all_categories_represented(self) -> None:
        categories = set(FUNCTION_WHITELIST_V1.values())
        assert categories == {"math", "logic", "lookup", "statistical", "financial", "text", "date"}

    def test_is_supported_case_insensitive(self) -> None:
        assert is_supported("sum")
        assert is_supported("SUM")
        assert is_supported("Sum")
        assert not is_supported("WEBSERVICE")
        assert not is_supported("RAND")


class TestFunctionRegistry:
    def test_builtins_registered(self) -> None:
        reg = FunctionRegistry()
        assert reg.has("SUM")
        assert reg.has("IF")
        assert reg.has("AVERAGE")

    def test_custom_registration(self) -> None:
        reg = FunctionRegistry()
        reg.register("MYFUNC", lambda args: 42)
        assert reg.has("MYFUNC")
        assert reg.get("MYFUNC")([]) == 42

    def test_case_insensitive_lookup(self) -> None:
        reg = FunctionRegistry()
        assert reg.get("sum") is reg.get("SUM")

    def test_supported_functions_property(self) -> None:
        reg = FunctionRegistry()
        funcs = reg.supported_functions
        assert isinstance(funcs, frozenset)
        assert "SUM" in funcs


class TestBuiltinSUM:
    def test_basic(self) -> None:
        fn = _BUILTINS["SUM"]
        assert fn([1, 2, 3]) == 6.0

    def test_nested_lists(self) -> None:
        fn = _BUILTINS["SUM"]
        assert fn([[1, 2], [3, 4]]) == 10.0

    def test_skip_none_and_strings(self) -> None:
        fn = _BUILTINS["SUM"]
        assert fn([1, None, "text", 3]) == 4.0

    def test_empty(self) -> None:
        fn = _BUILTINS["SUM"]
        assert fn([]) == 0.0

    def test_booleans_coerced(self) -> None:
        fn = _BUILTINS["SUM"]
        assert fn([True, False, 1]) == 2.0


class TestBuiltinABS:
    def test_positive(self) -> None:
        assert _BUILTINS["ABS"]([-5]) == 5.0

    def test_zero(self) -> None:
        assert _BUILTINS["ABS"]([0]) == 0.0

    def test_already_positive(self) -> None:
        assert _BUILTINS["ABS"]([3.14]) == 3.14

    def test_wrong_arity(self) -> None:
        with pytest.raises(ValueError, match="exactly 1"):
            _BUILTINS["ABS"]([1, 2])


class TestBuiltinROUND:
    def test_round_default_digits(self) -> None:
        assert _BUILTINS["ROUND"]([3.14159]) == 3.0

    def test_round_2_digits(self) -> None:
        assert _BUILTINS["ROUND"]([3.14159, 2]) == 3.14

    def test_round_negative_digits(self) -> None:
        assert _BUILTINS["ROUND"]([1234, -2]) == 1200.0


class TestBuiltinROUNDUP:
    def test_roundup_basic(self) -> None:
        assert _BUILTINS["ROUNDUP"]([3.2]) == 4.0

    def test_roundup_2_digits(self) -> None:
        assert _BUILTINS["ROUNDUP"]([3.141, 2]) == 3.15


class TestBuiltinINT:
    def test_positive(self) -> None:
        assert _BUILTINS["INT"]([3.7]) == 3.0

    def test_negative(self) -> None:
        # Excel INT floors toward negative infinity
        assert _BUILTINS["INT"]([-3.2]) == -4.0


class TestBuiltinIF:
    def test_true_branch(self) -> None:
        assert _BUILTINS["IF"]([True, "yes", "no"]) == "yes"

    def test_false_branch(self) -> None:
        assert _BUILTINS["IF"]([False, "yes", "no"]) == "no"

    def test_numeric_condition(self) -> None:
        assert _BUILTINS["IF"]([1, "yes", "no"]) == "yes"
        assert _BUILTINS["IF"]([0, "yes", "no"]) == "no"

    def test_missing_false_branch(self) -> None:
        assert _BUILTINS["IF"]([False, "yes"]) is False


class TestBuiltinIFERROR:
    def test_no_error(self) -> None:
        assert _BUILTINS["IFERROR"]([42, 0]) == 42

    def test_error_string(self) -> None:
        assert _BUILTINS["IFERROR"](["#DIV/0!", 0]) == 0

    def test_ref_error(self) -> None:
        assert _BUILTINS["IFERROR"](["#REF!", "fallback"]) == "fallback"


class TestBuiltinLogic:
    def test_and_all_true(self) -> None:
        assert _BUILTINS["AND"]([True, True, 1]) is True

    def test_and_one_false(self) -> None:
        assert _BUILTINS["AND"]([True, False]) is False

    def test_or_one_true(self) -> None:
        assert _BUILTINS["OR"]([False, True]) is True

    def test_or_all_false(self) -> None:
        assert _BUILTINS["OR"]([False, 0, None]) is False

    def test_not(self) -> None:
        assert _BUILTINS["NOT"]([True]) is False
        assert _BUILTINS["NOT"]([False]) is True


class TestBuiltinCounting:
    def test_count_numeric(self) -> None:
        assert _BUILTINS["COUNT"]([1, "text", None, 3.5, True]) == 3.0

    def test_counta_non_empty(self) -> None:
        assert _BUILTINS["COUNTA"]([1, "text", None, 3.5]) == 3.0

    def test_count_empty(self) -> None:
        assert _BUILTINS["COUNT"]([]) == 0.0


class TestBuiltinMinMax:
    def test_min(self) -> None:
        assert _BUILTINS["MIN"]([3, 1, 4, 1, 5]) == 1.0

    def test_max(self) -> None:
        assert _BUILTINS["MAX"]([3, 1, 4, 1, 5]) == 5.0

    def test_min_empty(self) -> None:
        assert _BUILTINS["MIN"]([]) == 0.0

    def test_max_nested(self) -> None:
        assert _BUILTINS["MAX"]([[1, 2], [3, 4]]) == 4.0


class TestBuiltinAVERAGE:
    def test_basic(self) -> None:
        assert _BUILTINS["AVERAGE"]([2, 4, 6]) == 4.0

    def test_empty_raises(self) -> None:
        with pytest.raises(ValueError, match="no numeric"):
            _BUILTINS["AVERAGE"]([])

    def test_skip_non_numeric(self) -> None:
        assert _BUILTINS["AVERAGE"]([10, None, "text", 20]) == 15.0


class TestBuiltinDivisionByZero:
    """Edge case: ensure no unhandled ZeroDivisionError from builtins."""

    def test_average_single(self) -> None:
        assert _BUILTINS["AVERAGE"]([0]) == 0.0


# ---------------------------------------------------------------------------
# ExcelError infrastructure
# ---------------------------------------------------------------------------


class TestExcelError:
    def test_singleton_identity(self) -> None:
        """ExcelError.of() returns cached singletons."""
        assert ExcelError.of("#N/A") is ExcelError.NA
        assert ExcelError.of("#VALUE!") is ExcelError.VALUE
        assert ExcelError.of("#REF!") is ExcelError.REF
        assert ExcelError.of("#DIV/0!") is ExcelError.DIV0
        assert ExcelError.of("#NUM!") is ExcelError.NUM
        assert ExcelError.of("#NAME?") is ExcelError.NAME

    def test_equality_with_strings(self) -> None:
        """ExcelError compares equal to its string code."""
        assert ExcelError.NA == "#N/A"
        assert ExcelError.DIV0 == "#DIV/0!"
        assert "#REF!" == ExcelError.REF  # reverse direction too

    def test_inequality(self) -> None:
        assert ExcelError.NA != ExcelError.REF
        assert ExcelError.NA != "#VALUE!"
        assert ExcelError.NA != 42

    def test_repr_and_str(self) -> None:
        assert repr(ExcelError.NA) == "#N/A"
        assert str(ExcelError.DIV0) == "#DIV/0!"

    def test_hashable(self) -> None:
        """ExcelError can be used in sets and as dict keys."""
        s = {ExcelError.NA, ExcelError.REF}
        assert len(s) == 2
        assert ExcelError.NA in s

    def test_case_insensitive_of(self) -> None:
        assert ExcelError.of("#n/a") is ExcelError.NA
        assert ExcelError.of("#div/0!") is ExcelError.DIV0


class TestIsError:
    def test_detects_excel_error(self) -> None:
        assert is_error(ExcelError.NA)
        assert is_error(ExcelError.DIV0)

    def test_rejects_non_errors(self) -> None:
        assert not is_error("#N/A")  # plain string is NOT an error object
        assert not is_error(42)
        assert not is_error(None)
        assert not is_error("")


class TestFirstError:
    def test_finds_first(self) -> None:
        assert first_error(1, ExcelError.NA, ExcelError.REF) is ExcelError.NA

    def test_none_when_no_errors(self) -> None:
        assert first_error(1, "hello", None) is None

    def test_single_error(self) -> None:
        assert first_error(ExcelError.DIV0) is ExcelError.DIV0


# ---------------------------------------------------------------------------
# Error propagation through builtins
# ---------------------------------------------------------------------------


class TestErrorPropagation:
    """Verify errors propagate correctly through builtin functions."""

    def test_sum_skips_errors_in_range(self) -> None:
        """SUM over a range silently skips errors (Excel behavior)."""
        assert _BUILTINS["SUM"]([1, ExcelError.NA, 3]) == 4.0

    def test_average_skips_errors_in_range(self) -> None:
        """AVERAGE skips error cells in aggregation."""
        assert _BUILTINS["AVERAGE"]([10, ExcelError.DIV0, 20]) == 15.0

    def test_count_skips_errors(self) -> None:
        """COUNT only counts numeric values, errors are skipped."""
        assert _BUILTINS["COUNT"]([1, ExcelError.NA, 3]) == 2.0

    def test_min_skips_errors(self) -> None:
        assert _BUILTINS["MIN"]([5, ExcelError.REF, 2]) == 2.0

    def test_max_skips_errors(self) -> None:
        assert _BUILTINS["MAX"]([5, ExcelError.REF, 2]) == 5.0

    def test_iferror_catches_excel_error(self) -> None:
        """IFERROR returns fallback for ExcelError instances."""
        assert _BUILTINS["IFERROR"]([ExcelError.NA, 0]) == 0
        assert _BUILTINS["IFERROR"]([ExcelError.DIV0, "safe"]) == "safe"

    def test_iferror_passes_non_error(self) -> None:
        assert _BUILTINS["IFERROR"]([42, 0]) == 42

    def test_iferror_catches_legacy_string_errors(self) -> None:
        """IFERROR still catches plain error strings for backward compat."""
        assert _BUILTINS["IFERROR"](["#DIV/0!", 0]) == 0

    def test_power_returns_excel_error(self) -> None:
        """POWER returns ExcelError.NUM for invalid operations."""
        result = _BUILTINS["POWER"]([-4, 0.5])
        assert isinstance(result, ExcelError)
        assert result == ExcelError.NUM

    def test_match_returns_excel_error_on_miss(self) -> None:
        """MATCH returns ExcelError.NA when no match found."""
        rv = RangeValue(values=["a", "b", "c"], n_rows=3, n_cols=1)
        result = _BUILTINS["MATCH"](["z", rv, 0])
        assert isinstance(result, ExcelError)
        assert result == ExcelError.NA

    def test_index_returns_ref_error(self) -> None:
        """INDEX returns ExcelError.REF for out-of-bounds."""
        rv = RangeValue(values=[1, 2, 3], n_rows=3, n_cols=1)
        result = _BUILTINS["INDEX"]([rv, 99])
        assert isinstance(result, ExcelError)
        assert result == ExcelError.REF

    def test_vlookup_returns_na_on_miss(self) -> None:
        """VLOOKUP returns ExcelError.NA when lookup value not found."""
        rv = RangeValue(values=["a", 1, "b", 2], n_rows=2, n_cols=2)
        result = _BUILTINS["VLOOKUP"](["z", rv, 2, False])
        assert isinstance(result, ExcelError)
        assert result == ExcelError.NA

    def test_choose_returns_value_error(self) -> None:
        """CHOOSE returns ExcelError.VALUE for out-of-range index."""
        result = _BUILTINS["CHOOSE"]([5, "a", "b"])
        assert isinstance(result, ExcelError)
        assert result == ExcelError.VALUE


# ---------------------------------------------------------------------------
# Financial builtins
# ---------------------------------------------------------------------------


class TestBuiltinPV:
    def test_basic_loan(self) -> None:
        """PV of $1000/month for 30 years at 5% annual."""
        result = _BUILTINS["PV"]([0.05 / 12, 360, -1000])
        assert abs(result - 186281.62) < 0.01

    def test_zero_rate(self) -> None:
        result = _BUILTINS["PV"]([0, 12, -100])
        assert result == 1200.0

    def test_with_fv(self) -> None:
        result = _BUILTINS["PV"]([0.08 / 12, 60, -200, -5000])
        assert abs(result - 13219.74) < 1.0

    def test_annuity_due(self) -> None:
        """type=1 (beginning of period) shifts payments."""
        result_ordinary = _BUILTINS["PV"]([0.06 / 12, 12, -100, 0, 0])
        result_due = _BUILTINS["PV"]([0.06 / 12, 12, -100, 0, 1])
        # Annuity due is worth more (payments earlier)
        assert result_due > result_ordinary


class TestBuiltinFV:
    def test_basic_savings(self) -> None:
        """FV of $100/month for 10 years at 7% annual."""
        result = _BUILTINS["FV"]([0.07 / 12, 120, -100])
        assert abs(result - 17308.48) < 0.01

    def test_zero_rate(self) -> None:
        result = _BUILTINS["FV"]([0, 12, -100])
        assert result == 1200.0

    def test_with_pv(self) -> None:
        result = _BUILTINS["FV"]([0.06 / 12, 120, -100, -1000])
        assert result > 0  # should accumulate


class TestBuiltinPMT:
    def test_mortgage_payment(self) -> None:
        """$200,000 mortgage at 5% for 30 years."""
        result = _BUILTINS["PMT"]([0.05 / 12, 360, 200000])
        assert abs(result - (-1073.64)) < 0.01

    def test_zero_rate(self) -> None:
        result = _BUILTINS["PMT"]([0, 12, 1200])
        assert result == -100.0

    def test_with_fv(self) -> None:
        """Save towards a $10k goal."""
        result = _BUILTINS["PMT"]([0.06 / 12, 60, 0, -10000])
        assert result > 0  # positive because we receive money

    def test_roundtrip_pv_pmt(self) -> None:
        """PV(rate, nper, PMT(rate, nper, pv)) == pv (roundtrip identity)."""
        rate, nper, pv = 0.05 / 12, 360, 200000
        pmt = _BUILTINS["PMT"]([rate, nper, pv])
        computed_pv = _BUILTINS["PV"]([rate, nper, pmt])
        assert abs(computed_pv - pv) < 0.01


class TestBuiltinNPV:
    def test_basic(self) -> None:
        """NPV at 10% for [-10000, 3000, 4200, 6800]."""
        result = _BUILTINS["NPV"]([0.10, -10000, 3000, 4200, 6800])
        assert abs(result - 1188.44) < 1.0

    def test_all_positive(self) -> None:
        result = _BUILTINS["NPV"]([0.05, 100, 100, 100])
        assert result > 0

    def test_with_range(self) -> None:
        """NPV accepts RangeValue for cash flows."""
        rv = RangeValue(values=[-1000.0, 500.0, 500.0, 500.0], n_rows=4, n_cols=1)
        result = _BUILTINS["NPV"]([0.10, rv])
        assert abs(result - 221.30) < 1.0


class TestBuiltinIRR:
    def test_basic(self) -> None:
        """IRR of [-10000, 3000, 4200, 6800] should be ~16.3%."""
        result = _BUILTINS["IRR"]([[-10000, 3000, 4200, 6800]])
        assert abs(result - 0.1634) < 0.001

    def test_zero_npv_at_irr(self) -> None:
        """Verify NPV at the IRR rate is approximately zero."""
        flows = [-10000, 3000, 4200, 6800]
        irr = _BUILTINS["IRR"]([flows])
        npv_at_irr = sum(v / (1 + irr) ** i for i, v in enumerate(flows))
        assert abs(npv_at_irr) < 0.01

    def test_no_sign_change_returns_num(self) -> None:
        """All positive cash flows -> #NUM! (no IRR exists)."""
        result = _BUILTINS["IRR"]([[100, 200, 300]])
        assert isinstance(result, ExcelError)
        assert result == ExcelError.NUM

    def test_with_range_value(self) -> None:
        rv = RangeValue(values=[-5000.0, 2000.0, 2000.0, 2000.0], n_rows=4, n_cols=1)
        result = _BUILTINS["IRR"]([rv])
        assert isinstance(result, float)
        assert abs(result - 0.0970) < 0.001


class TestBuiltinSLN:
    def test_basic(self) -> None:
        """SLN(30000, 7500, 10) = 2250."""
        assert _BUILTINS["SLN"]([30000, 7500, 10]) == 2250.0

    def test_zero_salvage(self) -> None:
        assert _BUILTINS["SLN"]([10000, 0, 5]) == 2000.0


class TestBuiltinDB:
    def test_basic_first_year(self) -> None:
        """DB(1000000, 100000, 6, 1) - first year depreciation."""
        result = _BUILTINS["DB"]([1000000, 100000, 6, 1])
        assert abs(result - 319000.0) < 1.0

    def test_second_year(self) -> None:
        result = _BUILTINS["DB"]([1000000, 100000, 6, 2])
        assert abs(result - 217239.0) < 1.0

    def test_partial_first_year(self) -> None:
        """DB with month=7 (7 months in first year)."""
        result = _BUILTINS["DB"]([1000000, 100000, 6, 1, 7])
        assert abs(result - 186083.33) < 1.0

    def test_invalid_life_returns_num(self) -> None:
        result = _BUILTINS["DB"]([1000, 100, 0, 1])
        assert isinstance(result, ExcelError)
        assert result == ExcelError.NUM


# ---------------------------------------------------------------------------
# Date builtins
# ---------------------------------------------------------------------------


class TestDateSerialRoundtrip:
    """Verify date <-> serial conversion with the Lotus 1-2-3 bug."""

    def test_jan_1_1900(self) -> None:
        from wolfxl.calc._functions import _date_to_serial, _serial_to_date

        assert _date_to_serial(1900, 1, 1) == 1
        assert _serial_to_date(1) == (1900, 1, 1)

    def test_feb_28_1900(self) -> None:
        from wolfxl.calc._functions import _date_to_serial, _serial_to_date

        assert _date_to_serial(1900, 2, 28) == 59
        assert _serial_to_date(59) == (1900, 2, 28)

    def test_lotus_bug_serial_60(self) -> None:
        """Serial 60 = Feb 29, 1900 (phantom date - Lotus 1-2-3 bug)."""
        from wolfxl.calc._functions import _serial_to_date

        assert _serial_to_date(60) == (1900, 2, 29)

    def test_mar_1_1900(self) -> None:
        """Mar 1, 1900 is serial 61 (one more than expected due to Lotus bug)."""
        from wolfxl.calc._functions import _date_to_serial, _serial_to_date

        assert _date_to_serial(1900, 3, 1) == 61
        assert _serial_to_date(61) == (1900, 3, 1)

    def test_jan_1_2000(self) -> None:
        from wolfxl.calc._functions import _date_to_serial, _serial_to_date

        serial = _date_to_serial(2000, 1, 1)
        assert serial == 36526
        assert _serial_to_date(serial) == (2000, 1, 1)

    def test_roundtrip_modern_dates(self) -> None:
        from wolfxl.calc._functions import _date_to_serial, _serial_to_date

        for y, m, d in [(2024, 6, 15), (2025, 12, 31), (2020, 2, 29)]:
            serial = _date_to_serial(y, m, d)
            assert _serial_to_date(serial) == (y, m, d)


class TestBuiltinTODAY:
    def test_returns_serial(self) -> None:

        result = _BUILTINS["TODAY"]([])
        # Should be a reasonable serial number for today
        assert isinstance(result, int)
        assert result > 40000  # After ~2009

    def test_matches_current_date(self) -> None:
        import datetime

        from wolfxl.calc._functions import _serial_to_date

        serial = _BUILTINS["TODAY"]([])
        y, m, d = _serial_to_date(serial)
        today = datetime.date.today()
        assert (y, m, d) == (today.year, today.month, today.day)


class TestBuiltinDATE:
    def test_basic(self) -> None:
        """DATE(2024, 1, 15) = serial for Jan 15, 2024."""
        result = _BUILTINS["DATE"]([2024, 1, 15])
        assert result == 45306

    def test_month_overflow(self) -> None:
        """DATE(2020, 14, 1) wraps to Feb 1, 2021."""
        result = _BUILTINS["DATE"]([2020, 14, 1])
        expected = _BUILTINS["DATE"]([2021, 2, 1])
        assert result == expected

    def test_month_underflow(self) -> None:
        """DATE(2020, -1, 1) goes back to Nov 1, 2019."""
        result = _BUILTINS["DATE"]([2020, -1, 1])
        expected = _BUILTINS["DATE"]([2019, 11, 1])
        assert result == expected

    def test_two_digit_year(self) -> None:
        """Excel treats 0-99 as 1900-1999."""
        result = _BUILTINS["DATE"]([20, 1, 1])
        expected = _BUILTINS["DATE"]([1920, 1, 1])
        assert result == expected


class TestBuiltinYearMonthDay:
    def test_year(self) -> None:
        serial = _BUILTINS["DATE"]([2024, 6, 15])
        assert _BUILTINS["YEAR"]([serial]) == 2024

    def test_month(self) -> None:
        serial = _BUILTINS["DATE"]([2024, 6, 15])
        assert _BUILTINS["MONTH"]([serial]) == 6

    def test_day(self) -> None:
        serial = _BUILTINS["DATE"]([2024, 6, 15])
        assert _BUILTINS["DAY"]([serial]) == 15


class TestBuiltinEDATE:
    def test_forward(self) -> None:
        """EDATE(Jan 15, +3) = Apr 15."""
        start = _BUILTINS["DATE"]([2024, 1, 15])
        result = _BUILTINS["EDATE"]([start, 3])
        expected = _BUILTINS["DATE"]([2024, 4, 15])
        assert result == expected

    def test_backward(self) -> None:
        """EDATE(Mar 15, -2) = Jan 15."""
        start = _BUILTINS["DATE"]([2024, 3, 15])
        result = _BUILTINS["EDATE"]([start, -2])
        expected = _BUILTINS["DATE"]([2024, 1, 15])
        assert result == expected

    def test_leap_year_clamp(self) -> None:
        """EDATE from Jan 31 + 1 month clamps to Feb 29 in leap year."""
        start = _BUILTINS["DATE"]([2024, 1, 31])
        result = _BUILTINS["EDATE"]([start, 1])
        # Should be Feb 29, 2024 (leap year) - day clamped
        y = _BUILTINS["YEAR"]([result])
        m = _BUILTINS["MONTH"]([result])
        d = _BUILTINS["DAY"]([result])
        assert (y, m) == (2024, 2)
        assert d == 29


class TestBuiltinEOMONTH:
    def test_same_month(self) -> None:
        """EOMONTH(Jan 15, 0) = Jan 31."""
        start = _BUILTINS["DATE"]([2024, 1, 15])
        result = _BUILTINS["EOMONTH"]([start, 0])
        expected = _BUILTINS["DATE"]([2024, 1, 31])
        assert result == expected

    def test_feb_leap(self) -> None:
        """EOMONTH(Jan 15, 1) in 2024 = Feb 29."""
        start = _BUILTINS["DATE"]([2024, 1, 15])
        result = _BUILTINS["EOMONTH"]([start, 1])
        assert _BUILTINS["DAY"]([result]) == 29

    def test_backward(self) -> None:
        start = _BUILTINS["DATE"]([2024, 3, 15])
        result = _BUILTINS["EOMONTH"]([start, -1])
        expected = _BUILTINS["DATE"]([2024, 2, 29])
        assert result == expected


class TestBuiltinDAYS:
    def test_basic(self) -> None:
        start = _BUILTINS["DATE"]([2024, 1, 1])
        end = _BUILTINS["DATE"]([2024, 1, 31])
        assert _BUILTINS["DAYS"]([end, start]) == 30

    def test_negative(self) -> None:
        start = _BUILTINS["DATE"]([2024, 3, 1])
        end = _BUILTINS["DATE"]([2024, 1, 1])
        assert _BUILTINS["DAYS"]([end, start]) < 0


# ---------------------------------------------------------------------------
# Phase 5: Conditional stats
# ---------------------------------------------------------------------------


class TestBuiltinAVERAGEIF:
    def test_basic(self) -> None:
        vals = RangeValue(values=[10, 20, 30, 40, 50], n_rows=5, n_cols=1)
        result = _BUILTINS["AVERAGEIF"]([vals, ">20"])
        assert result == pytest.approx(40.0)  # (30+40+50)/3

    def test_with_avg_range(self) -> None:
        criteria_range = RangeValue(values=["A", "B", "A", "B"], n_rows=4, n_cols=1)
        avg_range = RangeValue(values=[10, 20, 30, 40], n_rows=4, n_cols=1)
        result = _BUILTINS["AVERAGEIF"]([criteria_range, "A", avg_range])
        assert result == pytest.approx(20.0)  # (10+30)/2

    def test_no_match_returns_div0(self) -> None:
        vals = RangeValue(values=[1, 2, 3], n_rows=3, n_cols=1)
        result = _BUILTINS["AVERAGEIF"]([vals, ">100"])
        assert result == ExcelError.DIV0


class TestBuiltinAVERAGEIFS:
    def test_basic(self) -> None:
        avg_range = RangeValue(values=[10, 20, 30, 40], n_rows=4, n_cols=1)
        crit_range1 = RangeValue(values=["A", "B", "A", "B"], n_rows=4, n_cols=1)
        crit_range2 = RangeValue(values=[1, 2, 3, 4], n_rows=4, n_cols=1)
        result = _BUILTINS["AVERAGEIFS"]([avg_range, crit_range1, "A", crit_range2, ">1"])
        assert result == pytest.approx(30.0)  # only index 2 matches: A and >1

    def test_no_match_returns_div0(self) -> None:
        avg_range = RangeValue(values=[10, 20], n_rows=2, n_cols=1)
        crit_range = RangeValue(values=["A", "B"], n_rows=2, n_cols=1)
        result = _BUILTINS["AVERAGEIFS"]([avg_range, crit_range, "C"])
        assert result == ExcelError.DIV0


class TestBuiltinMINIFS:
    def test_basic(self) -> None:
        min_range = RangeValue(values=[10, 20, 30, 5, 15], n_rows=5, n_cols=1)
        crit_range = RangeValue(values=["A", "B", "A", "B", "A"], n_rows=5, n_cols=1)
        result = _BUILTINS["MINIFS"]([min_range, crit_range, "A"])
        assert result == pytest.approx(10.0)  # min of 10, 30, 15 where A

    def test_no_match_returns_zero(self) -> None:
        min_range = RangeValue(values=[10, 20], n_rows=2, n_cols=1)
        crit_range = RangeValue(values=["A", "B"], n_rows=2, n_cols=1)
        result = _BUILTINS["MINIFS"]([min_range, crit_range, "C"])
        assert result == 0.0


class TestBuiltinMAXIFS:
    def test_basic(self) -> None:
        max_range = RangeValue(values=[10, 20, 30, 5, 15], n_rows=5, n_cols=1)
        crit_range = RangeValue(values=["A", "B", "A", "B", "A"], n_rows=5, n_cols=1)
        result = _BUILTINS["MAXIFS"]([max_range, crit_range, "A"])
        assert result == pytest.approx(30.0)  # max of 10, 30, 15 where A

    def test_numeric_criteria(self) -> None:
        max_range = RangeValue(values=[100, 200, 300, 400], n_rows=4, n_cols=1)
        crit_range = RangeValue(values=[1, 2, 3, 4], n_rows=4, n_cols=1)
        result = _BUILTINS["MAXIFS"]([max_range, crit_range, ">=3"])
        assert result == pytest.approx(400.0)


# ---------------------------------------------------------------------------
# Phase 5: Text functions
# ---------------------------------------------------------------------------


class TestBuiltinUPPER:
    def test_basic(self) -> None:
        assert _BUILTINS["UPPER"](["hello"]) == "HELLO"

    def test_mixed(self) -> None:
        assert _BUILTINS["UPPER"](["Hello World"]) == "HELLO WORLD"

    def test_number_coerced(self) -> None:
        assert _BUILTINS["UPPER"]([123]) == "123"


class TestBuiltinLOWER:
    def test_basic(self) -> None:
        assert _BUILTINS["LOWER"](["HELLO"]) == "hello"

    def test_mixed(self) -> None:
        assert _BUILTINS["LOWER"](["Hello World"]) == "hello world"


class TestBuiltinTRIM:
    def test_leading_trailing(self) -> None:
        assert _BUILTINS["TRIM"](["  hello  "]) == "hello"

    def test_internal_spaces(self) -> None:
        assert _BUILTINS["TRIM"](["  hello   world  "]) == "hello world"


class TestBuiltinSUBSTITUTE:
    def test_replace_all(self) -> None:
        result = _BUILTINS["SUBSTITUTE"](["aaa", "a", "b"])
        assert result == "bbb"

    def test_replace_nth(self) -> None:
        result = _BUILTINS["SUBSTITUTE"](["aaa", "a", "b", 2])
        assert result == "aba"

    def test_not_found(self) -> None:
        result = _BUILTINS["SUBSTITUTE"](["hello", "x", "y"])
        assert result == "hello"


class TestBuiltinTEXT:
    def test_percent(self) -> None:
        assert _BUILTINS["TEXT"]([0.75, "0%"]) == "75%"

    def test_percent_decimal(self) -> None:
        assert _BUILTINS["TEXT"]([0.756, "0.0%"]) == "75.6%"

    def test_number_format(self) -> None:
        assert _BUILTINS["TEXT"]([1234567.89, "#,##0.00"]) == "1,234,567.89"

    def test_integer_format(self) -> None:
        assert _BUILTINS["TEXT"]([1234567, "#,##0"]) == "1,234,567"

    def test_date_format(self) -> None:
        serial = _BUILTINS["DATE"]([2024, 6, 15])
        result = _BUILTINS["TEXT"]([serial, "yyyy-mm-dd"])
        assert result == "2024-06-15"

    def test_non_numeric_returns_string(self) -> None:
        assert _BUILTINS["TEXT"](["hello", "0%"]) == "hello"


class TestBuiltinREPT:
    def test_basic(self) -> None:
        assert _BUILTINS["REPT"](["ab", 3]) == "ababab"

    def test_zero(self) -> None:
        assert _BUILTINS["REPT"](["x", 0]) == ""

    def test_negative(self) -> None:
        assert _BUILTINS["REPT"](["x", -1]) == ""


class TestBuiltinEXACT:
    def test_match(self) -> None:
        assert _BUILTINS["EXACT"](["hello", "hello"]) is True

    def test_case_sensitive(self) -> None:
        assert _BUILTINS["EXACT"](["Hello", "hello"]) is False

    def test_numbers(self) -> None:
        assert _BUILTINS["EXACT"]([123, "123"]) is True


class TestBuiltinFIND:
    def test_basic(self) -> None:
        assert _BUILTINS["FIND"](["lo", "hello"]) == 4

    def test_not_found(self) -> None:
        result = _BUILTINS["FIND"](["xyz", "hello"])
        assert result == ExcelError.VALUE

    def test_start_pos(self) -> None:
        assert _BUILTINS["FIND"](["l", "hello", 4]) == 4

    def test_case_sensitive(self) -> None:
        result = _BUILTINS["FIND"](["H", "hello"])
        assert result == ExcelError.VALUE

    def test_invalid_start(self) -> None:
        result = _BUILTINS["FIND"](["h", "hello", 0])
        assert result == ExcelError.VALUE


# ---------------------------------------------------------------------------
# XLOOKUP enhanced tests
# ---------------------------------------------------------------------------


class TestBuiltinXLOOKUP:
    """Tests for XLOOKUP match modes and search modes."""

    def _range(self, vals: list[Any]) -> RangeValue:
        return RangeValue(values=vals, n_rows=len(vals), n_cols=1)

    def test_exact_match_basic(self) -> None:
        lookup = self._range(["A", "B", "C"])
        returns = self._range([10, 20, 30])
        assert _BUILTINS["XLOOKUP"](["B", lookup, returns]) == 20

    def test_exact_match_not_found(self) -> None:
        lookup = self._range(["A", "B", "C"])
        returns = self._range([10, 20, 30])
        result = _BUILTINS["XLOOKUP"](["Z", lookup, returns, "missing"])
        assert result == "missing"

    def test_exact_match_default_not_found(self) -> None:
        lookup = self._range(["A", "B"])
        returns = self._range([10, 20])
        result = _BUILTINS["XLOOKUP"](["Z", lookup, returns])
        assert result == ExcelError.NA

    def test_next_smaller(self) -> None:
        """match_mode=-1: find largest value <= lookup."""
        lookup = self._range([10, 20, 30, 40, 50])
        returns = self._range(["a", "b", "c", "d", "e"])
        # 35 -> largest <= 35 is 30 -> "c"
        result = _BUILTINS["XLOOKUP"]([35, lookup, returns, "nope", -1])
        assert result == "c"

    def test_next_smaller_exact(self) -> None:
        """match_mode=-1 with exact value present."""
        lookup = self._range([10, 20, 30])
        returns = self._range(["a", "b", "c"])
        result = _BUILTINS["XLOOKUP"]([20, lookup, returns, "nope", -1])
        assert result == "b"

    def test_next_smaller_below_all(self) -> None:
        """match_mode=-1 when lookup is below all values."""
        lookup = self._range([10, 20, 30])
        returns = self._range(["a", "b", "c"])
        result = _BUILTINS["XLOOKUP"]([5, lookup, returns, "nope", -1])
        assert result == "nope"

    def test_next_larger(self) -> None:
        """match_mode=1: find smallest value >= lookup."""
        lookup = self._range([10, 20, 30, 40, 50])
        returns = self._range(["a", "b", "c", "d", "e"])
        # 25 -> smallest >= 25 is 30 -> "c"
        result = _BUILTINS["XLOOKUP"]([25, lookup, returns, "nope", 1])
        assert result == "c"

    def test_next_larger_above_all(self) -> None:
        """match_mode=1 when lookup is above all values."""
        lookup = self._range([10, 20, 30])
        returns = self._range(["a", "b", "c"])
        result = _BUILTINS["XLOOKUP"]([99, lookup, returns, "nope", 1])
        assert result == "nope"

    def test_wildcard_star(self) -> None:
        """match_mode=2: wildcard match with *."""
        lookup = self._range(["Apple", "Banana", "Cherry"])
        returns = self._range([1, 2, 3])
        result = _BUILTINS["XLOOKUP"](["Ban*", lookup, returns, "nope", 2])
        assert result == 2

    def test_wildcard_question(self) -> None:
        """match_mode=2: wildcard match with ?."""
        lookup = self._range(["Cat", "Car", "Cup"])
        returns = self._range([1, 2, 3])
        result = _BUILTINS["XLOOKUP"](["Ca?", lookup, returns, "nope", 2])
        assert result == 1  # first match

    def test_reverse_search(self) -> None:
        """search_mode=-1: last-to-first search."""
        lookup = self._range(["A", "B", "A", "C"])
        returns = self._range([10, 20, 30, 40])
        # search_mode=-1 finds last "A" at index 2 -> 30
        result = _BUILTINS["XLOOKUP"](["A", lookup, returns, "nope", 0, -1])
        assert result == 30

    def test_case_insensitive(self) -> None:
        lookup = self._range(["apple", "BANANA"])
        returns = self._range([1, 2])
        assert _BUILTINS["XLOOKUP"](["Apple", lookup, returns]) == 1
        assert _BUILTINS["XLOOKUP"](["banana", lookup, returns]) == 2


# ---------------------------------------------------------------------------
# TEXT format expansion tests
# ---------------------------------------------------------------------------


class TestBuiltinTEXTExpanded:
    """Tests for expanded TEXT format patterns."""

    def test_currency(self) -> None:
        assert _BUILTINS["TEXT"]([1234.5, "$#,##0.00"]) == "$1,234.50"

    def test_currency_integer(self) -> None:
        assert _BUILTINS["TEXT"]([1234, "$#,##0"]) == "$1,234"

    def test_currency_quoted(self) -> None:
        assert _BUILTINS["TEXT"]([1234.5, '"$"#,##0.00']) == "$1,234.50"

    def test_accounting_positive(self) -> None:
        result = _BUILTINS["TEXT"]([1234, "#,##0_);(#,##0)"])
        assert result == "1,234 "

    def test_accounting_negative(self) -> None:
        result = _BUILTINS["TEXT"]([-1234, "#,##0_);(#,##0)"])
        assert result == "(1,234)"

    def test_scientific(self) -> None:
        result = _BUILTINS["TEXT"]([1234.5, "0.00E+00"])
        assert "1.23" in result and "E" in result

    def test_date_us_format(self) -> None:
        serial = _BUILTINS["DATE"]([2024, 6, 15])
        result = _BUILTINS["TEXT"]([serial, "mm/dd/yyyy"])
        assert result == "06/15/2024"

    def test_date_abbreviated(self) -> None:
        serial = _BUILTINS["DATE"]([2024, 1, 5])
        result = _BUILTINS["TEXT"]([serial, "d-mmm-yyyy"])
        assert result == "5-Jan-2024"

    def test_date_abbreviated_short_year(self) -> None:
        serial = _BUILTINS["DATE"]([2024, 12, 25])
        result = _BUILTINS["TEXT"]([serial, "d-mmm-yy"])
        assert result == "25-Dec-24"

    def test_month_year(self) -> None:
        serial = _BUILTINS["DATE"]([2024, 3, 1])
        result = _BUILTINS["TEXT"]([serial, "mmm-yy"])
        assert result == "Mar-24"

    def test_one_decimal_comma(self) -> None:
        assert _BUILTINS["TEXT"]([1234.56, "#,##0.0"]) == "1,234.6"

    def test_three_decimal_places(self) -> None:
        assert _BUILTINS["TEXT"]([3.14159, "0.000"]) == "3.142"

    def test_general(self) -> None:
        assert _BUILTINS["TEXT"]([42, "General"]) == "42"
        assert _BUILTINS["TEXT"]([3.14, "General"]) == "3.14"


# ---------------------------------------------------------------------------
# OFFSET raw-arg protocol
# ---------------------------------------------------------------------------


class TestBuiltinOFFSET:
    """Verify OFFSET is in _BUILTINS with _raw_args attribute."""

    def test_in_builtins(self) -> None:
        assert "OFFSET" in _BUILTINS

    def test_raw_args_attribute(self) -> None:
        assert getattr(_BUILTINS["OFFSET"], '_raw_args', False) is True

    def test_callable(self) -> None:
        assert callable(_BUILTINS["OFFSET"])


# ---------------------------------------------------------------------------
# Time builtins
# ---------------------------------------------------------------------------


class TestSerialToTime:
    """Verify _serial_to_time fractional day extraction."""

    def test_midnight(self) -> None:
        from wolfxl.calc._functions import _serial_to_time

        assert _serial_to_time(0.0) == (0, 0, 0)

    def test_noon(self) -> None:
        from wolfxl.calc._functions import _serial_to_time

        assert _serial_to_time(0.5) == (12, 0, 0)

    def test_6pm(self) -> None:
        from wolfxl.calc._functions import _serial_to_time

        assert _serial_to_time(0.75) == (18, 0, 0)

    def test_with_integer_part(self) -> None:
        """Integer part is ignored - only fractional portion matters."""
        from wolfxl.calc._functions import _serial_to_time

        assert _serial_to_time(45000.5) == (12, 0, 0)

    def test_specific_time(self) -> None:
        """6:30:00 AM = 0.270833..."""
        from wolfxl.calc._functions import _serial_to_time

        h, m, s = _serial_to_time(6.5 / 24)
        assert h == 6
        assert m == 30
        assert s == 0


class TestBuiltinNOW:
    def test_returns_float(self) -> None:
        result = _BUILTINS["NOW"]([])
        assert isinstance(result, float)

    def test_greater_than_today(self) -> None:
        """NOW() should be >= TODAY() (has fractional time part)."""
        today_serial = _BUILTINS["TODAY"]([])
        now_serial = _BUILTINS["NOW"]([])
        assert now_serial >= today_serial

    def test_integer_part_matches_today(self) -> None:
        today_serial = _BUILTINS["TODAY"]([])
        now_serial = _BUILTINS["NOW"]([])
        assert int(now_serial) == today_serial


class TestBuiltinHOUR:
    def test_noon(self) -> None:
        """HOUR(0.5) = 12 (noon)."""
        assert _BUILTINS["HOUR"]([0.5]) == 12

    def test_6pm(self) -> None:
        """HOUR(0.75) = 18 (6 PM)."""
        assert _BUILTINS["HOUR"]([0.75]) == 18

    def test_midnight(self) -> None:
        assert _BUILTINS["HOUR"]([0.0]) == 0

    def test_wrong_arity(self) -> None:
        with pytest.raises(ValueError, match="exactly 1"):
            _BUILTINS["HOUR"]([0.5, 0.5])


class TestBuiltinMINUTE:
    def test_half_hour(self) -> None:
        """MINUTE(0.5 + 30/1440) = 30 (12:30 PM)."""
        serial = 0.5 + 30 / 1440  # 12:30:00
        assert _BUILTINS["MINUTE"]([serial]) == 30

    def test_noon_exact(self) -> None:
        assert _BUILTINS["MINUTE"]([0.5]) == 0


class TestBuiltinSECOND:
    def test_noon_exact(self) -> None:
        assert _BUILTINS["SECOND"]([0.5]) == 0

    def test_30_seconds(self) -> None:
        """SECOND at 12:00:30 = 30."""
        serial = 0.5 + 30 / 86400  # noon + 30 seconds
        assert _BUILTINS["SECOND"]([serial]) == 30

    def test_roundtrip_hour_now(self) -> None:
        """HOUR(NOW()) should return the current hour."""
        import datetime

        now_serial = _BUILTINS["NOW"]([])
        hour = _BUILTINS["HOUR"]([now_serial])
        assert hour == datetime.datetime.now().hour
