"""Function whitelist and builtin implementations for formula evaluation."""

from __future__ import annotations

import calendar
import contextlib
import datetime
import fnmatch
import math
import re
from collections.abc import Iterator
from contextvars import ContextVar
from dataclasses import dataclass
from typing import Any, Callable

# ---------------------------------------------------------------------------
# ExcelError: typed error values that propagate through formula chains
# ---------------------------------------------------------------------------


class ExcelError:
    """Excel error value that propagates through formula chains.

    Use ``ExcelError.of(code)`` to get a cached singleton for each error code.
    Errors compare equal to their string code (e.g., ``ExcelError.NA == ExcelError.NA``).
    """

    __slots__ = ("code",)
    _cache: dict[str, ExcelError] = {}

    NA: ExcelError
    VALUE: ExcelError
    REF: ExcelError
    DIV0: ExcelError
    NUM: ExcelError
    NAME: ExcelError

    def __init__(self, code: str) -> None:
        self.code = code

    @classmethod
    def of(cls, code: str) -> ExcelError:
        canon = code.upper()
        if canon not in cls._cache:
            cls._cache[canon] = cls(canon)
        return cls._cache[canon]

    def __repr__(self) -> str:
        return self.code

    def __str__(self) -> str:
        return self.code

    def __eq__(self, other: object) -> bool:
        if isinstance(other, ExcelError):
            return self.code == other.code
        if isinstance(other, str):
            return self.code == other.upper()
        return NotImplemented

    def __hash__(self) -> int:
        return hash(self.code)


# Singletons
ExcelError.NA = ExcelError.of("#N/A")
ExcelError.VALUE = ExcelError.of("#VALUE!")
ExcelError.REF = ExcelError.of("#REF!")
ExcelError.DIV0 = ExcelError.of("#DIV/0!")
ExcelError.NUM = ExcelError.of("#NUM!")
ExcelError.NAME = ExcelError.of("#NAME?")


def is_error(val: Any) -> bool:
    """Return True if *val* is an ExcelError instance."""
    return isinstance(val, ExcelError)


def first_error(*values: Any) -> ExcelError | None:
    """Return the first ExcelError found in *values*, or None."""
    for v in values:
        if isinstance(v, ExcelError):
            return v
    return None


# ---------------------------------------------------------------------------
# RangeValue: shape-aware 2D range container
# ---------------------------------------------------------------------------


@dataclass
class RangeValue:
    """A resolved cell range that preserves 2D shape metadata.

    Iterable and sized for backward compat with functions that expect lists.
    """

    values: list[Any]
    n_rows: int
    n_cols: int

    def get(self, row: int, col: int) -> Any:
        """Get value at 1-based (row, col) position."""
        if row < 1 or row > self.n_rows or col < 1 or col > self.n_cols:
            return None
        idx = (row - 1) * self.n_cols + (col - 1)
        return self.values[idx] if idx < len(self.values) else None

    def column(self, col: int) -> list[Any]:
        """Extract a 1-based column as a list."""
        if col < 1 or col > self.n_cols:
            return []
        return [self.values[(r * self.n_cols) + (col - 1)]
                for r in range(self.n_rows)
                if (r * self.n_cols) + (col - 1) < len(self.values)]

    def row(self, row: int) -> list[Any]:
        """Extract a 1-based row as a list."""
        if row < 1 or row > self.n_rows:
            return []
        start = (row - 1) * self.n_cols
        return self.values[start:start + self.n_cols]

    def as_flat(self) -> list[Any]:
        """Return values as a flat list."""
        return list(self.values)

    def __iter__(self):
        return iter(self.values)

    def __len__(self):
        return len(self.values)

# ---------------------------------------------------------------------------
# Whitelist: functions the calc engine will attempt to evaluate.
# Organized by category for readability.
# ---------------------------------------------------------------------------

FUNCTION_WHITELIST_V1: dict[str, str] = {
    # Math (10)
    "SUM": "math",
    "ABS": "math",
    "ROUND": "math",
    "ROUNDUP": "math",
    "ROUNDDOWN": "math",
    "INT": "math",
    "MOD": "math",
    "POWER": "math",
    "SQRT": "math",
    "SIGN": "math",
    # Logic (5)
    "IF": "logic",
    "AND": "logic",
    "OR": "logic",
    "NOT": "logic",
    "IFERROR": "logic",
    # Lookup (7)
    "VLOOKUP": "lookup",
    "HLOOKUP": "lookup",
    "INDEX": "lookup",
    "MATCH": "lookup",
    "OFFSET": "lookup",
    "CHOOSE": "lookup",
    "XLOOKUP": "lookup",
    # Statistical (13)
    "AVERAGE": "statistical",
    "AVERAGEIF": "statistical",
    "AVERAGEIFS": "statistical",
    "COUNT": "statistical",
    "COUNTA": "statistical",
    "COUNTIF": "statistical",
    "COUNTIFS": "statistical",
    "MIN": "statistical",
    "MINIFS": "statistical",
    "MAX": "statistical",
    "MAXIFS": "statistical",
    "SUMIF": "statistical",
    "SUMIFS": "statistical",
    # Financial (7)
    "PV": "financial",
    "FV": "financial",
    "PMT": "financial",
    "NPV": "financial",
    "IRR": "financial",
    "SLN": "financial",
    "DB": "financial",
    # Text (13)
    "LEFT": "text",
    "RIGHT": "text",
    "MID": "text",
    "LEN": "text",
    "CONCATENATE": "text",
    "UPPER": "text",
    "LOWER": "text",
    "TRIM": "text",
    "SUBSTITUTE": "text",
    "TEXT": "text",
    "REPT": "text",
    "EXACT": "text",
    "FIND": "text",
    # Date (8)
    "TODAY": "date",
    "DATE": "date",
    "YEAR": "date",
    "MONTH": "date",
    "DAY": "date",
    "EDATE": "date",
    "EOMONTH": "date",
    "DAYS": "date",
    # Time (4)
    "NOW": "date",
    "HOUR": "date",
    "MINUTE": "date",
    "SECOND": "date",
}


def is_supported(func_name: str) -> bool:
    """Check if a function name is in the evaluation whitelist."""
    return func_name.upper() in FUNCTION_WHITELIST_V1


# ---------------------------------------------------------------------------
# Builtin implementations - pure Python, no external deps.
# Each takes a list of resolved argument values.
# ---------------------------------------------------------------------------


def _coerce_numeric(values: list[Any]) -> list[float]:
    """Flatten and coerce values to floats, skipping None/str/bool.

    ExcelError values inside ranges are silently skipped (Excel SUM/AVERAGE
    behavior).  Direct scalar errors are handled by callers.
    """
    result: list[float] = []
    for v in values:
        if isinstance(v, ExcelError):
            continue  # skip errors in aggregation context
        if isinstance(v, RangeValue):
            result.extend(_coerce_numeric(v.values))
        elif isinstance(v, (list, tuple)):
            result.extend(_coerce_numeric(list(v)))
        elif isinstance(v, bool):
            # In Excel, TRUE=1, FALSE=0 in numeric context
            result.append(float(v))
        elif isinstance(v, (int, float)):
            result.append(float(v))
        # Skip None, str
    return result


def _builtin_sum(args: list[Any]) -> float:
    return sum(_coerce_numeric(args))


def _builtin_abs(args: list[Any]) -> float:
    if len(args) != 1:
        raise ValueError("ABS requires exactly 1 argument")
    nums = _coerce_numeric(args)
    if not nums:
        raise ValueError("ABS: non-numeric argument")
    return abs(nums[0])


def _builtin_round(args: list[Any]) -> float:
    if len(args) < 1 or len(args) > 2:
        raise ValueError("ROUND requires 1 or 2 arguments")
    nums = _coerce_numeric([args[0]])
    if not nums:
        raise ValueError("ROUND: non-numeric argument")
    digits = int(_coerce_numeric([args[1]])[0]) if len(args) > 1 else 0
    return round(nums[0], digits)


def _builtin_roundup(args: list[Any]) -> float:
    if len(args) < 1 or len(args) > 2:
        raise ValueError("ROUNDUP requires 1 or 2 arguments")
    nums = _coerce_numeric([args[0]])
    if not nums:
        raise ValueError("ROUNDUP: non-numeric argument")
    digits = int(_coerce_numeric([args[1]])[0]) if len(args) > 1 else 0
    if digits == 0:
        return float(math.ceil(nums[0]))
    factor = 10 ** digits
    return math.ceil(nums[0] * factor) / factor


def _builtin_int(args: list[Any]) -> float:
    if len(args) != 1:
        raise ValueError("INT requires exactly 1 argument")
    nums = _coerce_numeric(args)
    if not nums:
        raise ValueError("INT: non-numeric argument")
    return float(math.floor(nums[0]))


def _builtin_if(args: list[Any]) -> Any:
    if len(args) < 2 or len(args) > 3:
        raise ValueError("IF requires 2 or 3 arguments")
    condition = args[0]
    # Excel truthy: 0/False/None/"" are falsy
    truthy = bool(condition) if not isinstance(condition, (int, float)) else condition != 0
    if truthy:
        return args[1]
    return args[2] if len(args) > 2 else False


def _builtin_iferror(args: list[Any]) -> Any:
    if len(args) != 2:
        raise ValueError("IFERROR requires exactly 2 arguments")
    value = args[0]
    # Detect ExcelError instances or legacy error strings
    if isinstance(value, ExcelError):
        return args[1]
    if isinstance(value, str) and value.startswith("#"):
        return args[1]
    return value


def _builtin_and(args: list[Any]) -> bool:
    if not args:
        raise ValueError("AND requires at least 1 argument")
    for a in args:
        if isinstance(a, (RangeValue, list, tuple)):
            if not all(bool(x) for x in a if x is not None):
                return False
        elif not a:
            return False
    return True


def _builtin_or(args: list[Any]) -> bool:
    if not args:
        raise ValueError("OR requires at least 1 argument")
    for a in args:
        if isinstance(a, (RangeValue, list, tuple)):
            if any(bool(x) for x in a if x is not None):
                return True
        elif a:
            return True
    return False


def _builtin_not(args: list[Any]) -> bool:
    if len(args) != 1:
        raise ValueError("NOT requires exactly 1 argument")
    return not bool(args[0])


def _builtin_count(args: list[Any]) -> float:
    """COUNT - counts numeric values only."""
    return float(len(_coerce_numeric(args)))


def _builtin_counta(args: list[Any]) -> float:
    """COUNTA - counts non-empty values."""
    count = 0
    for v in args:
        if isinstance(v, (RangeValue, list, tuple)):
            count += sum(1 for x in v if x is not None)
        elif v is not None:
            count += 1
    return float(count)


def _builtin_min(args: list[Any]) -> float:
    nums = _coerce_numeric(args)
    if not nums:
        return 0.0
    return min(nums)


def _builtin_max(args: list[Any]) -> float:
    nums = _coerce_numeric(args)
    if not nums:
        return 0.0
    return max(nums)


def _builtin_average(args: list[Any]) -> float:
    nums = _coerce_numeric(args)
    if not nums:
        raise ValueError("AVERAGE: no numeric values")
    return sum(nums) / len(nums)


# ---------------------------------------------------------------------------
# Additional math builtins
# ---------------------------------------------------------------------------


def _builtin_rounddown(args: list[Any]) -> float:
    if len(args) < 1 or len(args) > 2:
        raise ValueError("ROUNDDOWN requires 1 or 2 arguments")
    nums = _coerce_numeric([args[0]])
    if not nums:
        raise ValueError("ROUNDDOWN: non-numeric argument")
    digits = int(_coerce_numeric([args[1]])[0]) if len(args) > 1 else 0
    if digits == 0:
        return float(math.trunc(nums[0]))
    factor = 10 ** digits
    return math.trunc(nums[0] * factor) / factor


def _builtin_mod(args: list[Any]) -> float:
    if len(args) != 2:
        raise ValueError("MOD requires exactly 2 arguments")
    nums = _coerce_numeric(args)
    if len(nums) != 2:
        raise ValueError("MOD: non-numeric argument")
    if nums[1] == 0:
        raise ValueError("MOD: division by zero")
    # Excel MOD: result has the sign of the divisor
    return nums[0] - nums[1] * math.floor(nums[0] / nums[1])


def _builtin_power(args: list[Any]) -> float | ExcelError:
    if len(args) != 2:
        raise ValueError("POWER requires exactly 2 arguments")
    nums = _coerce_numeric(args)
    if len(nums) != 2:
        raise ValueError("POWER: non-numeric argument")
    # Excel returns #NUM! for negative base with fractional exponent
    if nums[0] < 0 and not float(nums[1]).is_integer():
        return ExcelError.NUM
    return nums[0] ** nums[1]


def _builtin_sqrt(args: list[Any]) -> float:
    if len(args) != 1:
        raise ValueError("SQRT requires exactly 1 argument")
    nums = _coerce_numeric(args)
    if not nums:
        raise ValueError("SQRT: non-numeric argument")
    if nums[0] < 0:
        raise ValueError("SQRT: negative argument")
    return math.sqrt(nums[0])


def _builtin_sign(args: list[Any]) -> float:
    if len(args) != 1:
        raise ValueError("SIGN requires exactly 1 argument")
    nums = _coerce_numeric(args)
    if not nums:
        raise ValueError("SIGN: non-numeric argument")
    if nums[0] > 0:
        return 1.0
    if nums[0] < 0:
        return -1.0
    return 0.0


# ---------------------------------------------------------------------------
# Text builtins
# ---------------------------------------------------------------------------


def _coerce_string(val: Any) -> str:
    if val is None:
        return ""
    return str(val)


def _builtin_left(args: list[Any]) -> str | ExcelError:
    if len(args) < 1 or len(args) > 2:
        raise ValueError("LEFT requires 1 or 2 arguments")
    text = _coerce_string(args[0])
    num_chars = int(_coerce_numeric([args[1]])[0]) if len(args) > 1 else 1
    if num_chars < 0:
        return ExcelError.VALUE
    return text[:num_chars]


def _builtin_right(args: list[Any]) -> str:
    if len(args) < 1 or len(args) > 2:
        raise ValueError("RIGHT requires 1 or 2 arguments")
    text = _coerce_string(args[0])
    num_chars = int(_coerce_numeric([args[1]])[0]) if len(args) > 1 else 1
    return text[-num_chars:] if num_chars > 0 else ""


def _builtin_mid(args: list[Any]) -> str | ExcelError:
    if len(args) != 3:
        raise ValueError("MID requires exactly 3 arguments")
    text = _coerce_string(args[0])
    start = int(_coerce_numeric([args[1]])[0])
    num_chars = int(_coerce_numeric([args[2]])[0])
    if start < 1 or num_chars < 0:
        return ExcelError.VALUE
    # Excel MID is 1-indexed
    return text[start - 1 : start - 1 + num_chars]


def _builtin_len(args: list[Any]) -> float:
    if len(args) != 1:
        raise ValueError("LEN requires exactly 1 argument")
    return float(len(_coerce_string(args[0])))


def _builtin_concatenate(args: list[Any]) -> str:
    if not args:
        raise ValueError("CONCATENATE requires at least 1 argument")
    return "".join(_coerce_string(a) for a in args)


# ---------------------------------------------------------------------------
# Criteria matching engine (shared by SUMIF, SUMIFS, COUNTIF, COUNTIFS)
# ---------------------------------------------------------------------------

_CRITERIA_OP_RE = re.compile(r"^(>=|<=|<>|>|<|=)(.*)$")


def _parse_criteria(criteria: Any) -> Callable[[Any], bool]:
    """Parse an Excel criteria value into a predicate function.

    Supports:
    - Numeric exact match: ``100`` matches cells equal to 100
    - String exact match (case-insensitive): ``"Sales"``
    - Operator prefix: ``">100"``, ``"<=50"``, ``"<>0"``
    - Wildcards: ``"apple*"``, ``"?pple"`` (via fnmatch)
    """
    if isinstance(criteria, (int, float)):
        target = float(criteria)
        return lambda v: isinstance(v, (int, float)) and float(v) == target

    crit_str = str(criteria)

    # Check for operator prefix
    m = _CRITERIA_OP_RE.match(crit_str)
    if m:
        op, val_str = m.group(1), m.group(2).strip()
        try:
            threshold = float(val_str)
        except (ValueError, TypeError):
            # String comparison with operator
            val_lower = val_str.lower()
            if op == ">":
                return lambda v: str(v).lower() > val_lower if v is not None else False
            if op == "<":
                return lambda v: str(v).lower() < val_lower if v is not None else False
            if op == ">=":
                return lambda v: str(v).lower() >= val_lower if v is not None else False
            if op == "<=":
                return lambda v: str(v).lower() <= val_lower if v is not None else False
            if op == "<>":
                return lambda v: str(v).lower() != val_lower if v is not None else True
            if op == "=":
                return lambda v: str(v).lower() == val_lower if v is not None else False
            return lambda v: False

        if op == ">":
            return lambda v, t=threshold: isinstance(v, (int, float)) and float(v) > t
        if op == "<":
            return lambda v, t=threshold: isinstance(v, (int, float)) and float(v) < t
        if op == ">=":
            return lambda v, t=threshold: isinstance(v, (int, float)) and float(v) >= t
        if op == "<=":
            return lambda v, t=threshold: isinstance(v, (int, float)) and float(v) <= t
        if op == "<>":
            return lambda v, t=threshold: not (isinstance(v, (int, float)) and float(v) == t)
        if op == "=":
            return lambda v, t=threshold: isinstance(v, (int, float)) and float(v) == t

    # Wildcard check (contains * or ? not escaped)
    if "*" in crit_str or "?" in crit_str:
        pattern = crit_str.lower()
        return lambda v, p=pattern: fnmatch.fnmatch(str(v).lower(), p) if v is not None else False

    # Plain string exact match (case-insensitive)
    lower = crit_str.lower()
    return lambda v, lo=lower: str(v).lower() == lo if v is not None else False


def _match_criteria(criteria: Any, value: Any) -> bool:
    """Convenience: check whether *value* satisfies *criteria*."""
    return _parse_criteria(criteria)(value)


# ---------------------------------------------------------------------------
# Lookup builtins (INDEX, MATCH, XLOOKUP, CHOOSE)
# ---------------------------------------------------------------------------


def _builtin_index(args: list[Any]) -> Any:
    """INDEX(array, row_num [, col_num])."""
    if len(args) < 2 or len(args) > 3:
        raise ValueError("INDEX requires 2 or 3 arguments")
    array = args[0]
    row_num = args[1]
    col_num = args[2] if len(args) > 2 else None

    # Safety net: if row_num is None (e.g. from unsupported nested func), bail
    if row_num is None:
        return None

    row_num = int(float(row_num))

    if isinstance(array, RangeValue):
        if col_num is not None:
            col_num = int(float(col_num))
            if row_num < 1 or row_num > array.n_rows or col_num < 1 or col_num > array.n_cols:
                return ExcelError.REF
            return array.get(row_num, col_num)
        # 1D horizontal range: row_num acts as column index
        if array.n_rows == 1:
            if row_num < 1 or row_num > array.n_cols:
                return ExcelError.REF
            return array.get(1, row_num)
        # 1D column or multi-col: row_num selects row, return first col
        if row_num < 1 or row_num > array.n_rows:
            return ExcelError.REF
        if array.n_cols == 1:
            return array.get(row_num, 1)
        # Multi-col without col_num: return first column value
        return array.get(row_num, 1)

    # Plain list fallback
    if isinstance(array, (list, tuple)):
        if row_num < 1 or row_num > len(array):
            return ExcelError.REF
        return array[row_num - 1]

    return None


def _builtin_match(args: list[Any]) -> Any:
    """MATCH(lookup_value, lookup_array, [match_type]).

    match_type: 0=exact, 1=largest<=, -1=smallest>=. Default 0.
    """
    if len(args) < 2 or len(args) > 3:
        raise ValueError("MATCH requires 2 or 3 arguments")
    lookup_value = args[0]
    lookup_array = args[1]
    match_type = int(float(args[2])) if len(args) > 2 and args[2] is not None else 0

    # Flatten to list
    if isinstance(lookup_array, RangeValue):
        values = lookup_array.values
    elif isinstance(lookup_array, (list, tuple)):
        values = list(lookup_array)
    else:
        return ExcelError.NA

    if match_type == 0:
        # Exact match - case-insensitive for strings
        for i, v in enumerate(values):
            if v is None:
                continue
            if isinstance(lookup_value, str) and isinstance(v, str):
                if lookup_value.lower() == v.lower():
                    return i + 1  # 1-based
            elif isinstance(lookup_value, (int, float)) and isinstance(v, (int, float)):
                if float(lookup_value) == float(v):
                    return i + 1
            elif lookup_value == v:
                return i + 1
        return ExcelError.NA

    if match_type == 1:
        # Largest value <= lookup (assumes sorted ascending)
        best_idx = None
        for i, v in enumerate(values):
            if isinstance(v, (int, float)) and isinstance(lookup_value, (int, float)):
                if float(v) <= float(lookup_value):
                    best_idx = i + 1
        return best_idx if best_idx is not None else ExcelError.NA

    if match_type == -1:
        # Smallest value >= lookup (assumes sorted descending)
        best_idx = None
        for i, v in enumerate(values):
            if isinstance(v, (int, float)) and isinstance(lookup_value, (int, float)):
                if float(v) >= float(lookup_value):
                    best_idx = i + 1
        return best_idx if best_idx is not None else ExcelError.NA

    return ExcelError.NA


def _xlookup_wildcard_match(pattern: str, text: str) -> bool:
    """Match Excel wildcard pattern (*, ?) against text. Case-insensitive."""
    import re as _re

    regex = ""
    i = 0
    pat = pattern.lower()
    while i < len(pat):
        c = pat[i]
        if c == "~" and i + 1 < len(pat):
            regex += _re.escape(pat[i + 1])
            i += 2
        elif c == "*":
            regex += ".*"
            i += 1
        elif c == "?":
            regex += "."
            i += 1
        else:
            regex += _re.escape(c)
            i += 1
    return bool(_re.fullmatch(regex, text.lower()))


def _lookup_values(seq: Any) -> list[Any] | None:
    """Normalize a lookup vector to a Python list."""
    if isinstance(seq, RangeValue):
        return seq.values
    if isinstance(seq, (list, tuple)):
        return list(seq)
    return None


def _lookup_exact_index(
    lookup_value: Any,
    values: list[Any],
    indexes: range | None = None,
    *,
    wildcard: bool = False,
) -> int | None:
    """Return the index of the first exact lookup match."""
    for i in indexes if indexes is not None else range(len(values)):
        value = values[i]
        if value is None:
            continue
        if wildcard and isinstance(lookup_value, str) and isinstance(value, str):
            if _xlookup_wildcard_match(lookup_value, value):
                return i
        elif isinstance(lookup_value, str) and isinstance(value, str):
            if lookup_value.lower() == value.lower():
                return i
        elif isinstance(lookup_value, (int, float)) and isinstance(value, (int, float)):
            if float(lookup_value) == float(value):
                return i
        elif lookup_value == value:
            return i
    return None


def _lookup_last_sorted_lte_index(lookup_value: Any, values: list[Any]) -> int | None:
    """Return the last sorted-lookup index whose value is <= lookup."""
    best_idx = None
    for i, value in enumerate(values):
        if value is None:
            continue
        if isinstance(lookup_value, (int, float)) and isinstance(value, (int, float)):
            if float(value) <= float(lookup_value):
                best_idx = i
        elif isinstance(lookup_value, str) and isinstance(value, str):
            if value.lower() <= lookup_value.lower():
                best_idx = i
    return best_idx


def _coerce_range_lookup(value: Any) -> bool:
    """Normalize VLOOKUP/HLOOKUP's optional range_lookup argument."""
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return bool(value)
    if isinstance(value, str):
        return value.upper() != "FALSE"
    return True


def _builtin_xlookup(args: list[Any]) -> Any:
    """XLOOKUP(lookup_value, lookup_array, return_array, ...).

    match_mode: 0=exact (default), -1=next smaller, 1=next larger, 2=wildcard.
    search_mode: 1=first-to-last (default), -1=last-to-first.
    """
    if len(args) < 3 or len(args) > 6:
        raise ValueError("XLOOKUP requires 3 to 6 arguments")
    lookup_value = args[0]
    lookup_array = args[1]
    return_array = args[2]
    if_not_found = args[3] if len(args) > 3 else ExcelError.NA
    match_mode = int(float(args[4])) if len(args) > 4 and args[4] is not None else 0
    search_mode = int(float(args[5])) if len(args) > 5 and args[5] is not None else 1

    if match_mode not in (0, -1, 1, 2) or search_mode not in (1, -1):
        return None  # fall through to formulas lib for unsupported modes

    lookup_vals = _lookup_values(lookup_array)
    if lookup_vals is None:
        return if_not_found

    return_vals = _lookup_values(return_array)
    if return_vals is None:
        return if_not_found

    if search_mode == 1:
        search_range = range(len(lookup_vals))
    else:
        search_range = range(len(lookup_vals) - 1, -1, -1)

    def _safe_return(idx: int) -> Any:
        return return_vals[idx] if idx < len(return_vals) else if_not_found

    # --- Exact match (0) or wildcard match (2) ---
    if match_mode in (0, 2):
        idx = _lookup_exact_index(
            lookup_value,
            lookup_vals,
            search_range,
            wildcard=match_mode == 2,
        )
        if idx is not None:
            return _safe_return(idx)
        return if_not_found

    # --- Approximate match: -1 (next smaller) or 1 (next larger) ---
    if not isinstance(lookup_value, (int, float)):
        return if_not_found

    lv = float(lookup_value)
    best_idx: int | None = None
    best_val: float | None = None

    for i in search_range:
        v = lookup_vals[i]
        if not isinstance(v, (int, float)):
            continue
        fv = float(v)

        if match_mode == -1:  # next smaller: largest value <= lookup
            if fv <= lv:
                if best_val is None or fv > best_val:
                    best_val = fv
                    best_idx = i
        else:  # match_mode == 1: next larger: smallest value >= lookup
            if fv >= lv:
                if best_val is None or fv < best_val:
                    best_val = fv
                    best_idx = i

    if best_idx is not None:
        return _safe_return(best_idx)
    return if_not_found


def _builtin_vlookup(args: list[Any]) -> Any:
    """VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup]).

    range_lookup: FALSE (or 0) = exact match, TRUE (or 1, default) = approximate.
    Approximate match assumes the first column is sorted ascending and finds
    the largest value <= lookup_value.
    """
    if len(args) < 3 or len(args) > 4:
        raise ValueError("VLOOKUP requires 3 or 4 arguments")
    lookup_value = args[0]
    table_array = args[1]
    col_index_num = int(float(args[2]))
    range_lookup = True
    if len(args) > 3 and args[3] is not None:
        range_lookup = _coerce_range_lookup(args[3])

    if col_index_num < 1:
        return ExcelError.VALUE

    if isinstance(table_array, RangeValue):
        if col_index_num > table_array.n_cols:
            return ExcelError.REF
        first_col = table_array.column(1)
        return_col = table_array.column(col_index_num)
    elif isinstance(table_array, (list, tuple)):
        # Flat list treated as single column
        if col_index_num > 1:
            return ExcelError.REF
        first_col = list(table_array)
        return_col = first_col
    else:
        return ExcelError.NA

    if range_lookup:
        # Approximate match: largest value <= lookup_value (sorted ascending)
        best_idx = _lookup_last_sorted_lte_index(lookup_value, first_col)
        if best_idx is None:
            return ExcelError.NA
        return return_col[best_idx] if best_idx < len(return_col) else ExcelError.NA
    else:
        # Exact match (case-insensitive for strings)
        idx = _lookup_exact_index(lookup_value, first_col)
        if idx is not None:
            return return_col[idx] if idx < len(return_col) else ExcelError.NA
        return ExcelError.NA


def _builtin_hlookup(args: list[Any]) -> Any:
    """HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup]).

    Searches the first row of a table and returns a value from the specified row.
    range_lookup: FALSE (or 0) = exact match, TRUE (or 1, default) = approximate.
    """
    if len(args) < 3 or len(args) > 4:
        raise ValueError("HLOOKUP requires 3 or 4 arguments")
    lookup_value = args[0]
    table_array = args[1]
    row_index_num = int(float(args[2]))
    range_lookup = True
    if len(args) > 3 and args[3] is not None:
        range_lookup = _coerce_range_lookup(args[3])

    if row_index_num < 1:
        return ExcelError.VALUE

    if isinstance(table_array, RangeValue):
        if row_index_num > table_array.n_rows:
            return ExcelError.REF
        first_row = table_array.row(1)
        return_row = table_array.row(row_index_num)
    elif isinstance(table_array, (list, tuple)):
        # Flat list treated as single row
        if row_index_num > 1:
            return ExcelError.REF
        first_row = list(table_array)
        return_row = first_row
    else:
        return ExcelError.NA

    if range_lookup:
        # Approximate match: largest value <= lookup_value (sorted ascending)
        best_idx = _lookup_last_sorted_lte_index(lookup_value, first_row)
        if best_idx is None:
            return ExcelError.NA
        return return_row[best_idx] if best_idx < len(return_row) else ExcelError.NA
    else:
        # Exact match (case-insensitive for strings)
        idx = _lookup_exact_index(lookup_value, first_row)
        if idx is not None:
            return return_row[idx] if idx < len(return_row) else ExcelError.NA
        return ExcelError.NA


def _builtin_choose(args: list[Any]) -> Any:
    """CHOOSE(index_num, value1, value2, ...)."""
    if len(args) < 2:
        raise ValueError("CHOOSE requires at least 2 arguments")
    index_num = int(float(args[0]))
    if index_num < 1 or index_num > len(args) - 1:
        return ExcelError.VALUE
    return args[index_num]


# ---------------------------------------------------------------------------
# Conditional aggregation builtins (SUMIF, SUMIFS, COUNTIF, COUNTIFS)
# ---------------------------------------------------------------------------


def _builtin_sumif(args: list[Any]) -> float:
    """SUMIF(criteria_range, criteria, [sum_range])."""
    if len(args) < 2 or len(args) > 3:
        raise ValueError("SUMIF requires 2 or 3 arguments")
    criteria_range = args[0]
    criteria = args[1]
    sum_range = args[2] if len(args) > 2 else None

    # Flatten ranges
    if isinstance(criteria_range, RangeValue):
        crit_vals = criteria_range.values
    elif isinstance(criteria_range, (list, tuple)):
        crit_vals = list(criteria_range)
    else:
        crit_vals = [criteria_range]

    if sum_range is None:
        sum_vals = crit_vals
    elif isinstance(sum_range, RangeValue):
        sum_vals = sum_range.values
    elif isinstance(sum_range, (list, tuple)):
        sum_vals = list(sum_range)
    else:
        sum_vals = [sum_range]

    predicate = _parse_criteria(criteria)
    total = 0.0
    for i, cv in enumerate(crit_vals):
        if predicate(cv):
            sv = sum_vals[i] if i < len(sum_vals) else 0
            if isinstance(sv, (int, float)):
                total += float(sv)
    return total


def _builtin_sumifs(args: list[Any]) -> float:
    """SUMIFS(sum_range, criteria_range1, criteria1, ...).

    Note: sum_range is FIRST (unlike SUMIF where it's last).
    """
    if len(args) < 3 or len(args) % 2 == 0:
        raise ValueError("SUMIFS requires sum_range + pairs of (criteria_range, criteria)")
    sum_range = args[0]

    # Flatten sum_range
    if isinstance(sum_range, RangeValue):
        sum_vals = sum_range.values
    elif isinstance(sum_range, (list, tuple)):
        sum_vals = list(sum_range)
    else:
        sum_vals = [sum_range]

    # Build predicate pairs
    predicates: list[tuple[list[Any], Callable[[Any], bool]]] = []
    for j in range(1, len(args), 2):
        crit_range = args[j]
        criteria = args[j + 1]
        if isinstance(crit_range, RangeValue):
            cv = crit_range.values
        elif isinstance(crit_range, (list, tuple)):
            cv = list(crit_range)
        else:
            cv = [crit_range]
        predicates.append((cv, _parse_criteria(criteria)))

    total = 0.0
    for i in range(len(sum_vals)):
        if all(pred(cv[i]) if i < len(cv) else False for cv, pred in predicates):
            sv = sum_vals[i]
            if isinstance(sv, (int, float)):
                total += float(sv)
    return total


def _builtin_countif(args: list[Any]) -> float:
    """COUNTIF(range, criteria)."""
    if len(args) != 2:
        raise ValueError("COUNTIF requires exactly 2 arguments")
    count_range = args[0]
    criteria = args[1]

    if isinstance(count_range, RangeValue):
        values = count_range.values
    elif isinstance(count_range, (list, tuple)):
        values = list(count_range)
    else:
        values = [count_range]

    predicate = _parse_criteria(criteria)
    return float(sum(1 for v in values if predicate(v)))


def _builtin_countifs(args: list[Any]) -> float:
    """COUNTIFS(criteria_range1, criteria1, [criteria_range2, criteria2, ...])."""
    if len(args) < 2 or len(args) % 2 != 0:
        raise ValueError("COUNTIFS requires pairs of (criteria_range, criteria)")

    # Build predicate pairs
    predicates: list[tuple[list[Any], Callable[[Any], bool]]] = []
    for j in range(0, len(args), 2):
        crit_range = args[j]
        criteria = args[j + 1]
        if isinstance(crit_range, RangeValue):
            cv = crit_range.values
        elif isinstance(crit_range, (list, tuple)):
            cv = list(crit_range)
        else:
            cv = [crit_range]
        predicates.append((cv, _parse_criteria(criteria)))

    # Length of first criteria range determines row count
    n = len(predicates[0][0]) if predicates else 0
    count = 0
    for i in range(n):
        if all(pred(cv[i]) if i < len(cv) else False for cv, pred in predicates):
            count += 1
    return float(count)


# ---------------------------------------------------------------------------
# Conditional stats: AVERAGEIF, AVERAGEIFS, MINIFS, MAXIFS
# ---------------------------------------------------------------------------


def _flatten_range(arg: Any) -> list[Any]:
    """Extract values from a RangeValue, list, or single value."""
    if isinstance(arg, RangeValue):
        return arg.values
    if isinstance(arg, (list, tuple)):
        return list(arg)
    return [arg]


def _builtin_averageif(args: list[Any]) -> float | ExcelError:
    """AVERAGEIF(criteria_range, criteria, [average_range])."""
    if len(args) < 2 or len(args) > 3:
        raise ValueError("AVERAGEIF requires 2 or 3 arguments")
    crit_vals = _flatten_range(args[0])
    criteria = args[1]
    avg_vals = _flatten_range(args[2]) if len(args) > 2 else crit_vals

    predicate = _parse_criteria(criteria)
    total = 0.0
    count = 0
    for i, cv in enumerate(crit_vals):
        if predicate(cv):
            av = avg_vals[i] if i < len(avg_vals) else 0
            if isinstance(av, (int, float)):
                total += float(av)
                count += 1
    if count == 0:
        return ExcelError.DIV0
    return total / count


def _builtin_averageifs(args: list[Any]) -> float | ExcelError:
    """AVERAGEIFS(average_range, criteria_range1, criteria1, ...).

    Note: average_range is FIRST (like SUMIFS).
    """
    if len(args) < 3 or len(args) % 2 == 0:
        raise ValueError("AVERAGEIFS requires average_range + pairs of (criteria_range, criteria)")
    avg_vals = _flatten_range(args[0])

    predicates: list[tuple[list[Any], Callable[[Any], bool]]] = []
    for j in range(1, len(args), 2):
        cv = _flatten_range(args[j])
        predicates.append((cv, _parse_criteria(args[j + 1])))

    total = 0.0
    count = 0
    for i in range(len(avg_vals)):
        if all(pred(cv[i]) if i < len(cv) else False for cv, pred in predicates):
            av = avg_vals[i]
            if isinstance(av, (int, float)):
                total += float(av)
                count += 1
    if count == 0:
        return ExcelError.DIV0
    return total / count


def _builtin_minifs(args: list[Any]) -> float | ExcelError:
    """MINIFS(min_range, criteria_range1, criteria1, ...).

    Returns the minimum value among cells meeting all criteria.
    """
    if len(args) < 3 or len(args) % 2 == 0:
        raise ValueError("MINIFS requires min_range + pairs of (criteria_range, criteria)")
    min_vals = _flatten_range(args[0])

    predicates: list[tuple[list[Any], Callable[[Any], bool]]] = []
    for j in range(1, len(args), 2):
        cv = _flatten_range(args[j])
        predicates.append((cv, _parse_criteria(args[j + 1])))

    candidates: list[float] = []
    for i in range(len(min_vals)):
        if all(pred(cv[i]) if i < len(cv) else False for cv, pred in predicates):
            mv = min_vals[i]
            if isinstance(mv, (int, float)):
                candidates.append(float(mv))
    return min(candidates) if candidates else 0.0


def _builtin_maxifs(args: list[Any]) -> float | ExcelError:
    """MAXIFS(max_range, criteria_range1, criteria1, ...).

    Returns the maximum value among cells meeting all criteria.
    """
    if len(args) < 3 or len(args) % 2 == 0:
        raise ValueError("MAXIFS requires max_range + pairs of (criteria_range, criteria)")
    max_vals = _flatten_range(args[0])

    predicates: list[tuple[list[Any], Callable[[Any], bool]]] = []
    for j in range(1, len(args), 2):
        cv = _flatten_range(args[j])
        predicates.append((cv, _parse_criteria(args[j + 1])))

    candidates: list[float] = []
    for i in range(len(max_vals)):
        if all(pred(cv[i]) if i < len(cv) else False for cv, pred in predicates):
            mv = max_vals[i]
            if isinstance(mv, (int, float)):
                candidates.append(float(mv))
    return max(candidates) if candidates else 0.0


# ---------------------------------------------------------------------------
# Text builtins (extended): UPPER, LOWER, TRIM, SUBSTITUTE, TEXT, REPT, EXACT, FIND
# ---------------------------------------------------------------------------


def _builtin_upper(args: list[Any]) -> str:
    if len(args) != 1:
        raise ValueError("UPPER requires exactly 1 argument")
    return _coerce_string(args[0]).upper()


def _builtin_lower(args: list[Any]) -> str:
    if len(args) != 1:
        raise ValueError("LOWER requires exactly 1 argument")
    return _coerce_string(args[0]).lower()


def _builtin_trim(args: list[Any]) -> str:
    """TRIM: remove leading/trailing spaces and collapse internal spaces."""
    if len(args) != 1:
        raise ValueError("TRIM requires exactly 1 argument")
    return " ".join(_coerce_string(args[0]).split())


def _builtin_substitute(args: list[Any]) -> str:
    """SUBSTITUTE(text, old_text, new_text, [instance_num])."""
    if len(args) < 3 or len(args) > 4:
        raise ValueError("SUBSTITUTE requires 3 or 4 arguments")
    text = _coerce_string(args[0])
    old_text = _coerce_string(args[1])
    new_text = _coerce_string(args[2])

    if len(args) > 3 and args[3] is not None:
        instance = int(float(args[3]))
        # Replace only the Nth occurrence
        count = 0
        start = 0
        while True:
            idx = text.find(old_text, start)
            if idx == -1:
                break
            count += 1
            if count == instance:
                return text[:idx] + new_text + text[idx + len(old_text):]
            start = idx + 1
        return text  # instance not found, return unchanged

    return text.replace(old_text, new_text)


_MONTH_ABBRS = [
    "", "Jan", "Feb", "Mar", "Apr", "May", "Jun",
    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
]


def _builtin_text(args: list[Any]) -> str:
    """TEXT(value, format_text). Supports common Excel format patterns."""
    if len(args) != 2:
        raise ValueError("TEXT requires exactly 2 arguments")
    value = args[0]
    fmt = _coerce_string(args[1])

    if not isinstance(value, (int, float)):
        return _coerce_string(value)

    val = float(value)
    # Strip leading/trailing quotes for matching (e.g., "$"#,##0 -> $#,##0)
    fmt_clean = fmt.replace('"', '')
    fmt_lower = fmt_clean.lower()

    # --- Percentage formats ---
    if fmt_lower == "0%":
        return f"{val * 100:.0f}%"
    if fmt_lower == "0.0%":
        return f"{val * 100:.1f}%"
    if fmt_lower == "0.00%":
        return f"{val * 100:.2f}%"

    # --- Number formats with commas ---
    if fmt_lower in ("#,##0", "#,##0.0", "#,##0.00"):
        decimals = len(fmt_lower.split(".")[-1]) if "." in fmt_lower else 0
        return f"{val:,.{decimals}f}"

    # --- Currency: $#,##0 and $#,##0.00 ---
    if fmt_lower.startswith("$"):
        num_part = fmt_lower[1:]
        if num_part in ("#,##0", "#,##0.0", "#,##0.00"):
            decimals = len(num_part.split(".")[-1]) if "." in num_part else 0
            return f"${val:,.{decimals}f}"

    # --- Accounting format with parentheses for negatives ---
    if fmt_lower in ("#,##0_);(#,##0)", "#,##0.00_);(#,##0.00)"):
        decimals = 2 if ".00" in fmt_lower else 0
        if val < 0:
            return f"({abs(val):,.{decimals}f})"
        return f"{val:,.{decimals}f} "  # trailing space aligns with paren width

    # --- Scientific notation ---
    if fmt_lower == "0.00e+00":
        return f"{val:.2E}"

    # --- Date serial formats ---
    if fmt_lower in ("yyyy-mm-dd", "yyyy/mm/dd"):
        y, m, d = _serial_to_date(int(val))
        sep = "-" if "-" in fmt_clean else "/"
        return f"{y:04d}{sep}{m:02d}{sep}{d:02d}"
    if fmt_lower == "mm/dd/yyyy":
        y, m, d = _serial_to_date(int(val))
        return f"{m:02d}/{d:02d}/{y:04d}"
    if fmt_lower in ("d-mmm-yy", "d-mmm-yyyy"):
        y, m, d = _serial_to_date(int(val))
        abbr = _MONTH_ABBRS[m] if 1 <= m <= 12 else f"{m:02d}"
        yr = y % 100 if fmt_lower == "d-mmm-yy" else y
        return f"{d}-{abbr}-{yr:02d}" if fmt_lower == "d-mmm-yy" else f"{d}-{abbr}-{yr:04d}"
    if fmt_lower == "mmm-yy":
        y, m, d = _serial_to_date(int(val))
        abbr = _MONTH_ABBRS[m] if 1 <= m <= 12 else f"{m:02d}"
        return f"{abbr}-{y % 100:02d}"

    # --- Plain numeric 0, 0.0, 0.00, 0.000, etc. ---
    if fmt_lower.replace("0", "").replace(".", "") == "" and fmt_lower.startswith("0"):
        decimals = len(fmt_lower.split(".")[-1]) if "." in fmt_lower else 0
        return f"{val:.{decimals}f}"

    # --- General ---
    if fmt_lower == "general":
        return str(int(val)) if val == int(val) else str(val)

    # Fallback
    return str(value)


def _builtin_rept(args: list[Any]) -> str:
    """REPT(text, number_times)."""
    if len(args) != 2:
        raise ValueError("REPT requires exactly 2 arguments")
    text = _coerce_string(args[0])
    n = int(float(args[1]))
    if n < 0:
        return ""
    return text * n


def _builtin_exact(args: list[Any]) -> bool:
    """EXACT(text1, text2). Case-sensitive comparison."""
    if len(args) != 2:
        raise ValueError("EXACT requires exactly 2 arguments")
    return _coerce_string(args[0]) == _coerce_string(args[1])


def _builtin_find(args: list[Any]) -> int | ExcelError:
    """FIND(find_text, within_text, [start_num]). Case-sensitive, 1-based."""
    if len(args) < 2 or len(args) > 3:
        raise ValueError("FIND requires 2 or 3 arguments")
    find_text = _coerce_string(args[0])
    within_text = _coerce_string(args[1])
    start_num = int(float(args[2])) if len(args) > 2 and args[2] is not None else 1

    if start_num < 1:
        return ExcelError.VALUE

    # Convert to 0-based for Python's str.find
    idx = within_text.find(find_text, start_num - 1)
    if idx == -1:
        return ExcelError.VALUE
    return idx + 1  # Back to 1-based


# ---------------------------------------------------------------------------
# Financial builtins (PV, FV, PMT, NPV, IRR, SLN, DB)
# ---------------------------------------------------------------------------


def _builtin_pv(args: list[Any]) -> float | ExcelError:
    """PV(rate, nper, pmt, [fv], [type]).

    Present value of an investment: the total amount that a series of future
    payments is worth right now.
    """
    if len(args) < 3 or len(args) > 5:
        raise ValueError("PV requires 3 to 5 arguments")
    rate = float(args[0])
    nper = float(args[1])
    pmt = float(args[2])
    fv = float(args[3]) if len(args) > 3 and args[3] is not None else 0.0
    pmt_type = int(float(args[4])) if len(args) > 4 and args[4] is not None else 0

    if rate == 0:
        return -(fv + pmt * nper)
    pv_annuity = pmt * (1 + rate * pmt_type) * (1 - (1 + rate) ** (-nper)) / rate
    pv_fv = fv / (1 + rate) ** nper
    return -(pv_annuity + pv_fv)


def _builtin_fv(args: list[Any]) -> float | ExcelError:
    """FV(rate, nper, pmt, [pv], [type]).

    Future value of an investment based on periodic, constant payments
    and a constant interest rate.
    """
    if len(args) < 3 or len(args) > 5:
        raise ValueError("FV requires 3 to 5 arguments")
    rate = float(args[0])
    nper = float(args[1])
    pmt = float(args[2])
    pv = float(args[3]) if len(args) > 3 and args[3] is not None else 0.0
    pmt_type = int(float(args[4])) if len(args) > 4 and args[4] is not None else 0

    if rate == 0:
        return -(pv + pmt * nper)
    fv_pv = pv * (1 + rate) ** nper
    fv_annuity = pmt * (1 + rate * pmt_type) * ((1 + rate) ** nper - 1) / rate
    return -(fv_pv + fv_annuity)


def _builtin_pmt(args: list[Any]) -> float | ExcelError:
    """PMT(rate, nper, pv, [fv], [type]).

    Payment for a loan based on constant payments and constant interest rate.
    """
    if len(args) < 3 or len(args) > 5:
        raise ValueError("PMT requires 3 to 5 arguments")
    rate = float(args[0])
    nper = float(args[1])
    pv = float(args[2])
    fv = float(args[3]) if len(args) > 3 and args[3] is not None else 0.0
    pmt_type = int(float(args[4])) if len(args) > 4 and args[4] is not None else 0

    if rate == 0:
        return -(pv + fv) / nper
    pvif = (1 + rate) ** nper
    return -(rate * (pv * pvif + fv)) / (pvif - 1) / (1 + rate * pmt_type)


def _builtin_npv(args: list[Any]) -> float:
    """NPV(rate, value1, [value2], ...).

    Net present value of a series of cash flows. Note: Excel NPV
    excludes time-0 cash flow (first value is at period 1).
    """
    if len(args) < 2:
        raise ValueError("NPV requires at least 2 arguments (rate + values)")
    rate = float(args[0])

    # Flatten remaining args (could be scalars or ranges)
    values: list[float] = []
    for a in args[1:]:
        if isinstance(a, RangeValue):
            for v in a.values:
                if isinstance(v, (int, float)):
                    values.append(float(v))
        elif isinstance(a, (list, tuple)):
            for v in a:
                if isinstance(v, (int, float)):
                    values.append(float(v))
        elif isinstance(a, (int, float)):
            values.append(float(a))

    return sum(v / (1 + rate) ** (i + 1) for i, v in enumerate(values))


def _builtin_irr(args: list[Any]) -> float | ExcelError:
    """IRR(values, [guess]).

    Internal rate of return for a series of cash flows.
    Uses Newton-Raphson with bisection fallback.
    """
    if len(args) < 1 or len(args) > 2:
        raise ValueError("IRR requires 1 or 2 arguments")

    # Flatten values
    raw = args[0]
    values: list[float] = []
    if isinstance(raw, RangeValue):
        for v in raw.values:
            if isinstance(v, (int, float)):
                values.append(float(v))
    elif isinstance(raw, (list, tuple)):
        for v in raw:
            if isinstance(v, (int, float)):
                values.append(float(v))
    else:
        return ExcelError.NUM

    if len(values) < 2:
        return ExcelError.NUM

    # Must have both positive and negative cash flows
    has_pos = any(v > 0 for v in values)
    has_neg = any(v < 0 for v in values)
    if not (has_pos and has_neg):
        return ExcelError.NUM

    guess = float(args[1]) if len(args) > 1 and args[1] is not None else 0.1

    def _npv(rate: float) -> float:
        return sum(v / (1 + rate) ** i for i, v in enumerate(values))

    def _npv_deriv(rate: float) -> float:
        return sum(-i * v / (1 + rate) ** (i + 1) for i, v in enumerate(values))

    # Newton-Raphson
    rate = guess
    for _ in range(100):
        npv_val = _npv(rate)
        if abs(npv_val) < 1e-10:
            return rate
        deriv = _npv_deriv(rate)
        if abs(deriv) < 1e-14:
            break
        new_rate = rate - npv_val / deriv
        if abs(new_rate - rate) < 1e-10:
            return new_rate
        rate = new_rate

    # Bisection fallback: search [-0.999, 10.0]
    lo, hi = -0.999, 10.0
    if _npv(lo) * _npv(hi) > 0:
        return ExcelError.NUM
    for _ in range(200):
        mid = (lo + hi) / 2
        if abs(_npv(mid)) < 1e-10 or (hi - lo) < 1e-12:
            return mid
        if _npv(lo) * _npv(mid) < 0:
            hi = mid
        else:
            lo = mid
    return ExcelError.NUM


def _builtin_sln(args: list[Any]) -> float:
    """SLN(cost, salvage, life).

    Straight-line depreciation for one period.
    """
    if len(args) != 3:
        raise ValueError("SLN requires exactly 3 arguments")
    cost = float(args[0])
    salvage = float(args[1])
    life = float(args[2])
    if life == 0:
        raise ValueError("SLN: life cannot be zero")
    return (cost - salvage) / life


def _builtin_db(args: list[Any]) -> float | ExcelError:
    """DB(cost, salvage, life, period, [month]).

    Fixed-declining balance depreciation. *month* is the number of months
    in the first year (default 12).
    """
    if len(args) < 4 or len(args) > 5:
        raise ValueError("DB requires 4 or 5 arguments")
    cost = float(args[0])
    salvage = float(args[1])
    life = int(float(args[2]))
    period = int(float(args[3]))
    month = int(float(args[4])) if len(args) > 4 and args[4] is not None else 12

    if life <= 0 or period <= 0:
        return ExcelError.NUM
    if cost <= 0:
        return 0.0

    # Excel rounds rate to 3 decimal places
    rate = round(1 - (salvage / cost) ** (1 / life), 3)
    book_value = cost

    for yr in range(1, period + 1):
        if yr == 1:
            dep = cost * rate * month / 12
        elif yr == life + 1:
            # Final partial year
            dep = book_value * rate * (12 - month) / 12
        else:
            dep = book_value * rate
        book_value -= dep

    return dep  # type: ignore[possibly-unbound]


# ---------------------------------------------------------------------------
# Date serial number helpers (Excel epoch: serial 1 = Jan 1, 1900)
# ---------------------------------------------------------------------------

# The Lotus 1-2-3 bug: serial 60 = Feb 29, 1900 (doesn't exist).
# Serials >= 61 are off by one vs a correct calendar.
_LOTUS_BUG_SERIAL = 60


# Workbooks on the 1904 date system number days from Jan 1, 1904, which is
# serial 1462 in the 1900 system. The evaluator sets this for each calculation.
_DATE1904: ContextVar[bool] = ContextVar("wolfxl_calc_date1904", default=False)
_DATE1904_OFFSET = 1462


@contextlib.contextmanager
def date_system(date1904: bool) -> Iterator[None]:
    """Evaluate date serials in the 1904 date system while the context is active."""
    token = _DATE1904.set(date1904)
    try:
        yield
    finally:
        _DATE1904.reset(token)


def _date_to_serial(y: int, m: int, d: int) -> int:
    """Convert (year, month, day) to an Excel serial number.

    Handles month overflow/underflow (e.g., month 14 wraps to Feb next year).
    Reproduces the Lotus 1-2-3 bug for dates before March 1, 1900.
    """
    # Normalize month overflow/underflow
    m -= 1  # 0-based
    y += m // 12
    m = m % 12 + 1

    # Build a Python date
    # Clamp day to max for the month
    max_day = calendar.monthrange(y, m)[1]
    d = min(d, max_day)
    dt = datetime.date(y, m, d)

    # Days from Jan 1, 1900
    epoch = datetime.date(1899, 12, 31)  # serial 0 is Dec 31, 1899
    serial = (dt - epoch).days

    # Lotus bug: for dates >= Mar 1, 1900, add 1 to account for
    # the phantom Feb 29, 1900
    if serial >= _LOTUS_BUG_SERIAL:
        serial += 1

    if _DATE1904.get():
        serial -= _DATE1904_OFFSET
    return serial


def _serial_to_date(serial: int) -> tuple[int, int, int]:
    """Convert an Excel serial number to (year, month, day).

    Handles the Lotus 1-2-3 bug: serial 60 = Feb 29, 1900.
    """
    if _DATE1904.get():
        serial += _DATE1904_OFFSET
    if serial == _LOTUS_BUG_SERIAL:
        return (1900, 2, 29)  # The phantom date

    # For serials > 60, subtract 1 to undo the Lotus bug offset
    adjusted = serial - 1 if serial > _LOTUS_BUG_SERIAL else serial

    epoch = datetime.date(1899, 12, 31)
    dt = epoch + datetime.timedelta(days=adjusted)
    return (dt.year, dt.month, dt.day)


def _date_serial(value: Any) -> float:
    """Read a date argument as an Excel serial number.

    Cells formatted as dates load as ``datetime.date`` or ``datetime.datetime``;
    Excel stores them as serial numbers, so date and time functions accept both.
    """
    if isinstance(value, datetime.datetime):
        seconds = value.hour * 3600 + value.minute * 60 + value.second + value.microsecond / 1e6
        return _date_to_serial(value.year, value.month, value.day) + seconds / 86400
    if isinstance(value, datetime.date):
        return _date_to_serial(value.year, value.month, value.day)
    return float(value)


def _serial_to_time(serial: int | float) -> tuple[int, int, int]:
    """Extract (hour, minute, second) from the fractional portion of a serial number."""
    frac = abs(float(serial)) - int(abs(float(serial)))
    total_seconds = round(frac * 86400)
    hour = total_seconds // 3600
    minute = (total_seconds % 3600) // 60
    second = total_seconds % 60
    return (hour, minute, second)


# ---------------------------------------------------------------------------
# Date builtins (TODAY, DATE, YEAR, MONTH, DAY, EDATE, EOMONTH, DAYS)
# ---------------------------------------------------------------------------


def _builtin_today(args: list[Any]) -> int:
    """TODAY(). Returns the current date as an Excel serial number."""
    today = datetime.date.today()
    return _date_to_serial(today.year, today.month, today.day)


def _builtin_date(args: list[Any]) -> int | ExcelError:
    """DATE(year, month, day). Returns an Excel serial number.

    Handles month overflow: DATE(2020,14,1) = DATE(2021,2,1).
    """
    if len(args) != 3:
        raise ValueError("DATE requires exactly 3 arguments")
    y = int(float(args[0]))
    m = int(float(args[1]))
    d = int(float(args[2]))
    # Excel interprets 0-29 as 1900-1929, 30-99 as 1930-1999
    if 0 <= y <= 29:
        y += 1900
    elif 30 <= y <= 99:
        y += 1900
    result = _date_to_serial(y, m, d)
    # Serial 0 is a real date (Jan 1, 1904) only in the 1904 system.
    if result < (0 if _DATE1904.get() else 1):
        return ExcelError.NUM
    return result


def _builtin_year(args: list[Any]) -> int:
    """YEAR(serial). Extract year from a serial number."""
    if len(args) != 1:
        raise ValueError("YEAR requires exactly 1 argument")
    serial = int(_date_serial(args[0]))
    y, _m, _d = _serial_to_date(serial)
    return y


def _builtin_month(args: list[Any]) -> int:
    """MONTH(serial). Extract month (1-12) from a serial number."""
    if len(args) != 1:
        raise ValueError("MONTH requires exactly 1 argument")
    serial = int(_date_serial(args[0]))
    _y, m, _d = _serial_to_date(serial)
    return m


def _builtin_day(args: list[Any]) -> int:
    """DAY(serial). Extract day (1-31) from a serial number."""
    if len(args) != 1:
        raise ValueError("DAY requires exactly 1 argument")
    serial = int(_date_serial(args[0]))
    _y, _m, d = _serial_to_date(serial)
    return d


def _builtin_edate(args: list[Any]) -> int | ExcelError:
    """EDATE(start_date, months). Date N months from start."""
    if len(args) != 2:
        raise ValueError("EDATE requires exactly 2 arguments")
    start_serial = int(_date_serial(args[0]))
    months = int(float(args[1]))
    y, m, d = _serial_to_date(start_serial)
    return _date_to_serial(y, m + months, d)


def _builtin_eomonth(args: list[Any]) -> int | ExcelError:
    """EOMONTH(start_date, months). End of month N months from start."""
    if len(args) != 2:
        raise ValueError("EOMONTH requires exactly 2 arguments")
    start_serial = int(_date_serial(args[0]))
    months = int(float(args[1]))
    y, m, _d = _serial_to_date(start_serial)
    # Move to target month
    m += months
    # Normalize
    m -= 1
    y += m // 12
    m = m % 12 + 1
    last_day = calendar.monthrange(y, m)[1]
    return _date_to_serial(y, m, last_day)


def _builtin_days(args: list[Any]) -> int:
    """DAYS(end_date, start_date). Simple subtraction."""
    if len(args) != 2:
        raise ValueError("DAYS requires exactly 2 arguments")
    end = int(_date_serial(args[0]))
    start = int(_date_serial(args[1]))
    return end - start


# ---------------------------------------------------------------------------
# Time builtins (NOW, HOUR, MINUTE, SECOND)
# ---------------------------------------------------------------------------


def _builtin_now(args: list[Any]) -> float:
    """NOW(). Returns the current date and time as an Excel serial number."""
    now = datetime.datetime.now()
    date_serial = _date_to_serial(now.year, now.month, now.day)
    time_frac = (now.hour * 3600 + now.minute * 60 + now.second) / 86400
    return date_serial + time_frac


def _builtin_hour(args: list[Any]) -> int:
    """HOUR(serial). Extract hour (0-23) from a serial number."""
    if len(args) != 1:
        raise ValueError("HOUR requires exactly 1 argument")
    serial = _date_serial(args[0])
    h, _m, _s = _serial_to_time(serial)
    return h


def _builtin_minute(args: list[Any]) -> int:
    """MINUTE(serial). Extract minute (0-59) from a serial number."""
    if len(args) != 1:
        raise ValueError("MINUTE requires exactly 1 argument")
    serial = _date_serial(args[0])
    _h, m, _s = _serial_to_time(serial)
    return m


def _builtin_second(args: list[Any]) -> int:
    """SECOND(serial). Extract second (0-59) from a serial number."""
    if len(args) != 1:
        raise ValueError("SECOND requires exactly 1 argument")
    serial = _date_serial(args[0])
    _h, _m, s = _serial_to_time(serial)
    return s


# ---------------------------------------------------------------------------
# Raw-arg builtins (receive raw arg strings, not resolved values)
# ---------------------------------------------------------------------------


def _builtin_offset(raw_args: list[str], eval_fn: Callable, sheet: str) -> Any:
    """OFFSET(reference, rows, cols, [height], [width]).

    Raw-arg function: first arg is a cell reference token, not a resolved value.
    Returns a RangeValue for multi-cell results or a scalar for single-cell.
    """
    from wolfxl._utils import a1_to_rowcol, rowcol_to_a1
    from wolfxl.calc._parser import expand_range

    if len(raw_args) < 3 or len(raw_args) > 5:
        return ExcelError.REF

    ref_str = raw_args[0].strip().replace('$', '')
    rows_offset = eval_fn(raw_args[1].strip(), sheet)
    cols_offset = eval_fn(raw_args[2].strip(), sheet)
    height = eval_fn(raw_args[3].strip(), sheet) if len(raw_args) > 3 else None
    width = eval_fn(raw_args[4].strip(), sheet) if len(raw_args) > 4 else None

    try:
        rows_offset = int(float(rows_offset))
        cols_offset = int(float(cols_offset))
    except (ValueError, TypeError):
        return ExcelError.REF

    # Parse the base reference
    if '!' in ref_str:
        parts = ref_str.split('!', 1)
        ref_sheet = parts[0].strip("'")
        cell_a1 = parts[1].upper()
    else:
        ref_sheet = sheet
        cell_a1 = ref_str.upper()

    # Handle range references (e.g. OFFSET(A1:A5, ...))
    if ':' in cell_a1:
        cell_a1 = cell_a1.split(':')[0]

    try:
        base_row, base_col = a1_to_rowcol(cell_a1)
    except ValueError:
        return ExcelError.REF

    target_row = base_row + rows_offset
    target_col = base_col + cols_offset

    h = int(float(height)) if height is not None else 1
    w = int(float(width)) if width is not None else 1

    if h < 1 or w < 1 or target_row < 1 or target_col < 1:
        return ExcelError.REF

    if h == 1 and w == 1:
        cell_ref = f"{ref_sheet}!{rowcol_to_a1(target_row, target_col)}"
        return eval_fn(cell_ref.split('!')[1], ref_sheet)

    end_row = target_row + h - 1
    end_col = target_col + w - 1
    start_a1 = rowcol_to_a1(target_row, target_col)
    end_a1 = rowcol_to_a1(end_row, end_col)
    range_ref = f"{ref_sheet}!{start_a1}:{end_a1}"

    cells = expand_range(range_ref)
    values = [eval_fn(c.split('!')[1], c.split('!')[0]) for c in cells]
    return RangeValue(values=values, n_rows=h, n_cols=w)


_builtin_offset._raw_args = True  # type: ignore[attr-defined]


# ---------------------------------------------------------------------------
# Registry
# ---------------------------------------------------------------------------

_BUILTINS: dict[str, Callable[..., Any]] = {
    "SUM": _builtin_sum,
    "ABS": _builtin_abs,
    "ROUND": _builtin_round,
    "ROUNDUP": _builtin_roundup,
    "ROUNDDOWN": _builtin_rounddown,
    "INT": _builtin_int,
    "MOD": _builtin_mod,
    "POWER": _builtin_power,
    "SQRT": _builtin_sqrt,
    "SIGN": _builtin_sign,
    "IF": _builtin_if,
    "IFERROR": _builtin_iferror,
    "AND": _builtin_and,
    "OR": _builtin_or,
    "NOT": _builtin_not,
    "COUNT": _builtin_count,
    "COUNTA": _builtin_counta,
    "MIN": _builtin_min,
    "MAX": _builtin_max,
    "AVERAGE": _builtin_average,
    "LEFT": _builtin_left,
    "RIGHT": _builtin_right,
    "MID": _builtin_mid,
    "LEN": _builtin_len,
    "CONCATENATE": _builtin_concatenate,
    "INDEX": _builtin_index,
    "MATCH": _builtin_match,
    "VLOOKUP": _builtin_vlookup,
    "HLOOKUP": _builtin_hlookup,
    "XLOOKUP": _builtin_xlookup,
    "CHOOSE": _builtin_choose,
    "SUMIF": _builtin_sumif,
    "SUMIFS": _builtin_sumifs,
    "COUNTIF": _builtin_countif,
    "COUNTIFS": _builtin_countifs,
    "PV": _builtin_pv,
    "FV": _builtin_fv,
    "PMT": _builtin_pmt,
    "NPV": _builtin_npv,
    "IRR": _builtin_irr,
    "SLN": _builtin_sln,
    "DB": _builtin_db,
    "TODAY": _builtin_today,
    "DATE": _builtin_date,
    "YEAR": _builtin_year,
    "MONTH": _builtin_month,
    "DAY": _builtin_day,
    "EDATE": _builtin_edate,
    "EOMONTH": _builtin_eomonth,
    "DAYS": _builtin_days,
    # Time (Phase 7)
    "NOW": _builtin_now,
    "HOUR": _builtin_hour,
    "MINUTE": _builtin_minute,
    "SECOND": _builtin_second,
    # Conditional stats (Phase 5)
    "AVERAGEIF": _builtin_averageif,
    "AVERAGEIFS": _builtin_averageifs,
    "MINIFS": _builtin_minifs,
    "MAXIFS": _builtin_maxifs,
    # Text (Phase 5)
    "UPPER": _builtin_upper,
    "LOWER": _builtin_lower,
    "TRIM": _builtin_trim,
    "SUBSTITUTE": _builtin_substitute,
    "TEXT": _builtin_text,
    "REPT": _builtin_rept,
    "EXACT": _builtin_exact,
    "FIND": _builtin_find,
    # Raw-arg functions (receive raw strings, dispatched via _raw_args protocol)
    "OFFSET": _builtin_offset,
}


class FunctionRegistry:
    """Registry of callable function implementations.

    Starts with builtins and can be extended with custom functions.
    """

    def __init__(self) -> None:
        self._functions: dict[str, Callable[..., Any]] = dict(_BUILTINS)

    def register(self, name: str, func: Callable[..., Any]) -> None:
        self._functions[name.upper()] = func

    def get(self, name: str) -> Callable[..., Any] | None:
        return self._functions.get(name.upper())

    def has(self, name: str) -> bool:
        return name.upper() in self._functions

    @property
    def supported_functions(self) -> frozenset[str]:
        return frozenset(self._functions.keys())
