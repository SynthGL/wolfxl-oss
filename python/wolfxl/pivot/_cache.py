"""``PivotCache`` — schema + records snapshot for a pivot's source range.

Mirrors :class:`openpyxl.pivot.cache.CacheDefinition` plus the
records part.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
import math
from typing import Any

from wolfxl._compat import _install_openpyxl_iter
from wolfxl.chart.reference import Reference

# Per-column type-inference thresholds. Tunable but documented.
_INFER_THRESHOLDS = {
    "max_string_unique_for_enumeration": 2000,
    "max_number_unique_for_enumeration": 200,
}


def _finite_float(value: Any, context: str) -> float:
    n = float(value)
    if not math.isfinite(n):
        raise ValueError(
            f"{context}: non-finite floats are not representable in xlsx pivot caches"
        )
    return n


# ---------------------------------------------------------------------------
# CacheValue — the variant emitted into each pivot record.
# ---------------------------------------------------------------------------


class CacheValue:
    """Tagged variant for a single pivot-cache value.

    Use the classmethod constructors (``CacheValue.string("North")``)
    rather than the dataclass init for clarity at call sites.
    """

    __slots__ = ("kind", "value")

    def __init__(self, kind: str, value: Any = None) -> None:
        self.kind = kind
        self.value = value

    @classmethod
    def string(cls, s: str) -> "CacheValue":
        """Create a string pivot-cache value.

        Args:
            s: Text value to store in the cache record.

        Returns:
            A ``CacheValue`` tagged as ``"string"``.
        """
        return cls("string", s)

    @classmethod
    def number(cls, n: float) -> "CacheValue":
        """Create a numeric pivot-cache value.

        Args:
            n: Numeric value to store. Values are normalized to ``float``.

        Returns:
            A ``CacheValue`` tagged as ``"number"``.
        """
        return cls("number", _finite_float(n, "CacheValue.number"))

    @classmethod
    def boolean(cls, b: bool) -> "CacheValue":
        """Create a boolean pivot-cache value.

        Args:
            b: Truth value to store.

        Returns:
            A ``CacheValue`` tagged as ``"boolean"``.
        """
        return cls("boolean", bool(b))

    @classmethod
    def date(cls, d: str | date | datetime) -> "CacheValue":
        """Create a datetime-like pivot-cache value.

        Args:
            d: ISO-like string, ``date``, or ``datetime`` value. ``date``
                inputs are normalized to midnight ISO datetimes.

        Returns:
            A ``CacheValue`` tagged as ``"date"``.
        """
        if isinstance(d, datetime):
            iso = d.isoformat()
        elif isinstance(d, date):
            iso = f"{d.isoformat()}T00:00:00"
        else:
            iso = str(d)
        return cls("date", iso)

    @classmethod
    def missing(cls) -> "CacheValue":
        """Create a missing-value pivot-cache marker.

        Returns:
            A ``CacheValue`` tagged as ``"missing"`` with no payload.
        """
        return cls("missing", None)

    @classmethod
    def error(cls, s: str) -> "CacheValue":
        """Create an error pivot-cache value.

        Args:
            s: Excel-style error token, such as ``"#N/A"``.

        Returns:
            A ``CacheValue`` tagged as ``"error"``.
        """
        return cls("error", s)

    def to_rust_dict(self) -> dict:
        if self.value is None:
            return {"kind": self.kind}
        return {"kind": self.kind, "value": self.value}

    def __repr__(self) -> str:
        if self.value is None:
            return f"CacheValue.{self.kind}()"
        return f"CacheValue.{self.kind}({self.value!r})"

    def __eq__(self, other: object) -> bool:
        if not isinstance(other, CacheValue):
            return NotImplemented
        return self.kind == other.kind and self.value == other.value

    def __hash__(self) -> int:
        return hash((self.kind, self.value))


# ---------------------------------------------------------------------------
# WorksheetSource
# ---------------------------------------------------------------------------


@dataclass
class WorksheetSource:
    """A worksheet source for a pivot cache.

    Either ``(sheet, ref)`` or ``name``, mutually exclusive.
    """

    sheet: str = ""
    ref: str = ""
    name: str | None = None

    def __post_init__(self) -> None:
        has_sheet_ref = bool(self.sheet) or bool(self.ref)
        has_name = self.name is not None and self.name != ""
        if has_sheet_ref and has_name:
            raise ValueError(
                "WorksheetSource: sheet+ref and name are mutually exclusive"
            )

    def to_rust_dict(self) -> dict:
        return {"sheet": self.sheet, "ref": self.ref, "name": self.name}


# ---------------------------------------------------------------------------
# SharedItems
# ---------------------------------------------------------------------------


@dataclass(init=False)
class SharedItems:
    """The ``<sharedItems>`` payload on a cache field.

    ``items=None`` + numeric attrs (min/max) → numeric-only no-enumeration
    form (``<sharedItems containsNumber="1" minValue=… maxValue=…/>``).

    ``items=[…]`` → enumeration form (``<s>``/``<n>``/``<d>``/etc.
    children).
    """

    count: int | None = None
    items: list[CacheValue] | None = None
    contains_blank: bool = False
    contains_mixed_types: bool = False
    contains_semi_mixed_types: bool = True
    contains_string: bool = True
    contains_number: bool = False
    contains_integer: bool = False
    contains_date: bool = False
    contains_non_date: bool = True
    min_value: float | None = None
    max_value: float | None = None
    min_date: str | None = None
    max_date: str | None = None
    long_text: bool = False

    def __init__(
        self,
        count: int | None = None,
        items: list[CacheValue] | None = None,
        contains_blank: bool = False,
        contains_mixed_types: bool = False,
        contains_semi_mixed_types: bool = True,
        contains_string: bool = True,
        contains_number: bool = False,
        contains_integer: bool = False,
        contains_date: bool = False,
        contains_non_date: bool = True,
        min_value: float | None = None,
        max_value: float | None = None,
        min_date: str | None = None,
        max_date: str | None = None,
        long_text: bool = False,
        **kw: Any,
    ) -> None:
        self.items = items if items is not None else kw.pop("_fields", None)
        self.count = len(self.items or []) if count is None else int(count)
        self.contains_blank = bool(kw.pop("containsBlank", contains_blank))
        self.contains_mixed_types = bool(kw.pop("containsMixedTypes", contains_mixed_types))
        self.contains_semi_mixed_types = bool(
            kw.pop("containsSemiMixedTypes", contains_semi_mixed_types)
        )
        self.contains_string = bool(kw.pop("containsString", contains_string))
        self.contains_number = bool(kw.pop("containsNumber", contains_number))
        self.contains_integer = bool(kw.pop("containsInteger", contains_integer))
        self.contains_date = bool(kw.pop("containsDate", contains_date))
        self.contains_non_date = bool(kw.pop("containsNonDate", contains_non_date))
        self.min_value = kw.pop("minValue", min_value)
        self.max_value = kw.pop("maxValue", max_value)
        self.min_date = kw.pop("minDate", min_date)
        self.max_date = kw.pop("maxDate", max_date)
        self.long_text = bool(kw.pop("longText", long_text))
        for extra_name, value in kw.items():
            setattr(self, extra_name, value)

    def to_rust_dict(self) -> dict:
        return {
            "count": self.count,
            "items": (
                [v.to_rust_dict() for v in self.items]
                if self.items is not None
                else None
            ),
            "contains_blank": self.contains_blank,
            "contains_mixed_types": self.contains_mixed_types,
            "contains_semi_mixed_types": self.contains_semi_mixed_types,
            "contains_string": self.contains_string,
            "contains_number": self.contains_number,
            "contains_integer": self.contains_integer,
            "contains_date": self.contains_date,
            "contains_non_date": self.contains_non_date,
            "min_value": self.min_value,
            "max_value": self.max_value,
            "min_date": self.min_date,
            "max_date": self.max_date,
            "long_text": self.long_text,
        }


# ---------------------------------------------------------------------------
# CacheField
# ---------------------------------------------------------------------------


@dataclass
class CacheField:
    """A single pivot-cache field — name + dtype + shared-items.

    `data_type` is one of ``"string" | "number" | "date" | "bool" |
    "mixed"``. ``formula`` and ``hierarchy`` are reserved for future
    use and pinned to ``None`` here.
    """

    name: str
    num_fmt_id: int = 0
    data_type: str = "string"
    shared_items: SharedItems = field(default_factory=SharedItems)
    formula: str | None = None
    hierarchy: int | None = None

    def to_rust_dict(self) -> dict:
        return {
            "name": self.name,
            "num_fmt_id": self.num_fmt_id,
            "data_type": self.data_type,
            "shared_items": self.shared_items.to_rust_dict(),
            "formula": self.formula,
            "hierarchy": self.hierarchy,
        }


# ---------------------------------------------------------------------------
# PivotCache (top-level)
# ---------------------------------------------------------------------------


class PivotCache:
    """The pivot cache — schema + records snapshot.

    Constructed by the user with a :class:`Reference` source range, then
    registered via :meth:`Workbook.add_pivot_cache(cache)` which:

    1. allocates ``cache_id`` (0-based, returned to the caller);
    2. walks the source range to materialize :attr:`fields` and
       :attr:`records` (per-column type inference);
    3. enqueues the cache for emit at save time.

    Until step 2 runs, :attr:`fields` and :attr:`records` are
    ``None`` and :meth:`to_rust_dict` raises.
    """

    def __init__(
        self,
        source: Reference,
        *,
        refresh_on_load: bool = False,
        refreshed_by: str = "wolfxl",
    ) -> None:
        if not isinstance(source, Reference):
            raise TypeError(
                f"PivotCache(source=...) must be a Reference, got {type(source).__name__}"
            )
        self.source = source
        self.refresh_on_load = refresh_on_load
        self.refreshed_by = refreshed_by
        # Set by Workbook.add_pivot_cache().
        self._cache_id: int | None = None
        # Set by _materialize() (called by add_pivot_cache).
        self._fields: list[CacheField] | None = None
        self._records: list[list[CacheValue]] | None = None
        # Calculated fields + field groups. Both are cache-scoped
        # (deep-clone aliases via the parent cache).
        self.calculated_fields: list[Any] = []
        self.field_groups: list[Any] = []

    @property
    def cache_id(self) -> int:
        if self._cache_id is None:
            raise RuntimeError(
                "PivotCache has not been registered yet — call "
                "Workbook.add_pivot_cache(cache) first"
            )
        return self._cache_id

    @property
    def fields(self) -> list[CacheField]:
        if self._fields is None:
            raise RuntimeError(
                "PivotCache has not been materialized yet — call "
                "Workbook.add_pivot_cache(cache) first"
            )
        return self._fields

    @property
    def field_names(self) -> list[str]:
        return [f.name for f in self.fields]

    def field_index(self, name: str) -> int:
        """Look up a field by name; raises ``KeyError`` if not found.

        Used by :class:`PivotTable` to resolve string field names
        (``rows=["region"]``) to indices.
        """
        for i, f in enumerate(self.fields):
            if f.name == name:
                return i
        raise KeyError(
            f"PivotCache has no field named {name!r}; "
            f"available: {self.field_names}"
        )

    # ------------------------------------------------------------------
    # Materialization — walk the source range, infer types, build
    # SharedItems and records. Called by Workbook.add_pivot_cache().
    # ------------------------------------------------------------------

    def _materialize(self, ws: Any) -> None:
        """Walk ``ws[self.source.range]``; build fields + records.

        Per-column type inference:

        =======================  =============  =====================================
        Column observation       data_type      shared_items strategy
        =======================  =============  =====================================
        all numeric, ≤200 unique  number         enumerate items
        all numeric, >200 unique  number         min/max attrs only (no items)
        all string, ≤2000 unique  string         enumerate items
        all string, >2000 unique  string         attrs only
        all ISO date              date           enumerate dates as <d> items
        mixed types               mixed          contains_semi_mixed_types=True
        all None                  string         contains_blank=True, count=0
        =======================  =============  =====================================
        """
        rows = self._collect_source_rows(ws)
        if not rows:
            raise ValueError(
                f"PivotCache.source ({self.source}) is empty — "
                "needs ≥1 header row + ≥1 data row"
            )
        header = rows[0]
        data_rows = rows[1:]
        if not data_rows:
            raise ValueError(
                "PivotCache.source has only a header row; "
                "needs ≥1 data row"
            )

        n_cols = len(header)
        cache_fields: list[CacheField] = []
        records: list[list[CacheValue]] = [[] for _ in data_rows]

        for col_idx in range(n_cols):
            field_name = self._cell_to_field_name(header[col_idx], col_idx)
            col_values = [row[col_idx] if col_idx < len(row) else None
                          for row in data_rows]
            cf = self._infer_cache_field(field_name, col_values)
            cache_fields.append(cf)
            # Materialize per-row record cell using inferred field.
            for ri, raw in enumerate(col_values):
                cv = self._raw_to_cache_value(raw)
                records[ri].append(cv)

        self._fields = cache_fields
        self._records = records

    def _collect_source_rows(self, ws: Any) -> list[list[Any]]:
        """Iterate worksheet rows in :attr:`source` range; return values."""
        from wolfxl.chart.reference import _index_to_col

        if self.source.range_string:
            # Reference parsed from a string is portable across sheets.
            ws_target = self._resolve_worksheet(ws)
        else:
            ws_target = self.source.worksheet

        rows: list[list[Any]] = []
        for r in range(self.source.min_row, self.source.max_row + 1):
            row_vals: list[Any] = []
            for c in range(self.source.min_col, self.source.max_col + 1):
                col_letter = _index_to_col(c)
                addr = f"{col_letter}{r}"
                cell = ws_target[addr]
                row_vals.append(cell.value if hasattr(cell, "value") else cell)
            rows.append(row_vals)
        return rows

    def _resolve_worksheet(self, wb_or_ws: Any) -> Any:
        """If a workbook is passed, resolve via source's sheet name;
        if a worksheet is passed, return it directly."""
        title = self.source.worksheet.title if self.source.worksheet else None
        if hasattr(wb_or_ws, "sheetnames") and title is not None:
            return wb_or_ws[title]
        return wb_or_ws  # assume it's already a worksheet

    @staticmethod
    def _cell_to_field_name(cell: Any, col_idx: int) -> str:
        if cell is None or cell == "":
            return f"Field{col_idx + 1}"
        return str(cell)

    def _infer_cache_field(
        self, name: str, values: list[Any]
    ) -> CacheField:
        field = self._infer_cache_field_items(name, values)
        si = field.shared_items
        if si.items is not None and any(v is None for v in values):
            # Excel enumerates blanks as a trailing ``<m/>`` shared item so
            # records index it like any other item and pivot axes can show
            # a "(blank)" row.
            si.items.append(CacheValue.missing())
            si.count = len(si.items)
            si.contains_blank = True
        return field

    def _infer_cache_field_items(
        self, name: str, values: list[Any]
    ) -> CacheField:
        # Classify per-value type; pick dominant.
        kinds = [self._classify(v) for v in values]
        non_missing = [k for k in kinds if k != "missing"]
        unique_kinds = set(non_missing)

        if not unique_kinds:
            # All None.
            return CacheField(
                name=name,
                data_type="string",
                shared_items=SharedItems(
                    count=0,
                    items=[],
                    contains_blank=True,
                    contains_semi_mixed_types=False,
                    contains_string=False,
                    contains_non_date=False,
                ),
            )

        if unique_kinds == {"number"}:
            return self._infer_numeric_field(name, values)
        if unique_kinds == {"string"}:
            return self._infer_string_field(name, values)
        if unique_kinds == {"date"}:
            return self._infer_date_field(name, values)
        if unique_kinds == {"boolean"}:
            return self._infer_boolean_field(name, values)
        # Mixed.
        return self._infer_mixed_field(name, values)

    @staticmethod
    def _classify(v: Any) -> str:
        if v is None:
            return "missing"
        if isinstance(v, bool):
            return "boolean"
        if isinstance(v, (int, float)):
            return "number"
        if isinstance(v, (date, datetime)):
            return "date"
        return "string"

    def _infer_numeric_field(
        self, name: str, values: list[Any]
    ) -> CacheField:
        nums = [
            _finite_float(v, f"PivotCache field {name!r}")
            for v in values
            if v is not None
        ]
        unique_nums = set(nums)
        threshold = _INFER_THRESHOLDS["max_number_unique_for_enumeration"]
        if len(unique_nums) <= threshold:
            # Enumerate.
            ordered = sorted(unique_nums)
            items = [CacheValue.number(n) for n in ordered]
            si = SharedItems(
                count=len(items),
                items=items,
                contains_semi_mixed_types=False,
                contains_string=False,
                contains_number=True,
                contains_integer=all(float(n).is_integer() for n in ordered),
                min_value=min(ordered) if ordered else None,
                max_value=max(ordered) if ordered else None,
            )
        else:
            # Attrs only.
            si = SharedItems(
                count=None,
                items=None,
                contains_semi_mixed_types=False,
                contains_string=False,
                contains_number=True,
                contains_integer=all(float(n).is_integer() for n in nums),
                min_value=min(nums) if nums else None,
                max_value=max(nums) if nums else None,
            )
        return CacheField(name=name, data_type="number", shared_items=si)

    def _infer_string_field(
        self, name: str, values: list[Any]
    ) -> CacheField:
        strs = [str(v) for v in values if v is not None]
        unique = list(dict.fromkeys(strs))  # preserve insertion order
        threshold = _INFER_THRESHOLDS["max_string_unique_for_enumeration"]
        if len(unique) <= threshold:
            items = [CacheValue.string(s) for s in unique]
            si = SharedItems(
                count=len(items),
                items=items,
                contains_blank=any(v is None for v in values),
                contains_semi_mixed_types=True,
                contains_string=True,
            )
        else:
            si = SharedItems(
                count=None,
                items=None,
                contains_blank=any(v is None for v in values),
                contains_semi_mixed_types=True,
                contains_string=True,
                long_text=any(len(s) > 256 for s in strs),
            )
        return CacheField(name=name, data_type="string", shared_items=si)

    def _infer_date_field(
        self, name: str, values: list[Any]
    ) -> CacheField:
        dates = [self._date_to_iso(v) for v in values if v is not None]
        unique = list(dict.fromkeys(dates))
        items = [CacheValue.date(d) for d in unique]
        si = SharedItems(
            count=len(items),
            items=items,
            contains_blank=any(v is None for v in values),
            contains_semi_mixed_types=False,
            contains_string=False,
            contains_date=True,
            contains_non_date=False,
            min_date=min(unique) if unique else None,
            max_date=max(unique) if unique else None,
        )
        return CacheField(name=name, data_type="date", shared_items=si)

    def _infer_boolean_field(
        self, name: str, values: list[Any]
    ) -> CacheField:
        bools = [bool(v) for v in values if v is not None]
        unique = list(dict.fromkeys(bools))
        items = [CacheValue.boolean(b) for b in unique]
        si = SharedItems(
            count=len(items),
            items=items,
            contains_blank=any(v is None for v in values),
            contains_semi_mixed_types=False,
            contains_string=False,
        )
        return CacheField(name=name, data_type="bool", shared_items=si)

    def _infer_mixed_field(
        self, name: str, values: list[Any]
    ) -> CacheField:
        items: list[CacheValue] = []
        seen: set = set()
        for v in values:
            if v is None:
                continue
            cv = self._raw_to_cache_value(v)
            key = (cv.kind, cv.value)
            if key not in seen:
                seen.add(key)
                items.append(cv)
        si = SharedItems(
            count=len(items),
            items=items,
            contains_blank=any(v is None for v in values),
            contains_mixed_types=True,
            contains_semi_mixed_types=True,
            contains_string=any(c.kind == "string" for c in items),
            contains_number=any(c.kind == "number" for c in items),
            contains_date=any(c.kind == "date" for c in items),
            contains_non_date=not all(c.kind == "date" for c in items),
        )
        return CacheField(name=name, data_type="mixed", shared_items=si)

    @staticmethod
    def _date_to_iso(v: Any) -> str:
        if isinstance(v, datetime):
            return v.isoformat()
        if isinstance(v, date):
            return f"{v.isoformat()}T00:00:00"
        return str(v)

    def _raw_to_cache_value(self, v: Any) -> CacheValue:
        if v is None:
            return CacheValue.missing()
        if isinstance(v, bool):
            return CacheValue.boolean(v)
        if isinstance(v, (int, float)):
            return CacheValue.number(_finite_float(v, "PivotCache record value"))
        if isinstance(v, (date, datetime)):
            return CacheValue.date(v)
        return CacheValue.string(str(v))

    # ------------------------------------------------------------------
    # to_rust_dict — cache definition + records contracts.
    # ------------------------------------------------------------------

    # ------------------------------------------------------------------
    # Calculated fields
    # ------------------------------------------------------------------

    def add_calculated_field(
        self,
        name: str,
        formula: str,
        data_type: str = "number",
    ):
        """Add a calculated field to this cache.

        Excel evaluates the formula on open — wolfxl does not
        pre-evaluate values. The formula uses pivot field names as
        bare identifiers (``"= revenue - cost"``).

        Returns the registered :class:`CalculatedField` for chaining.
        """
        from ._calc import CalculatedField

        cf = CalculatedField(name=name, formula=formula, data_type=data_type)
        self.calculated_fields.append(cf)
        return cf

    # ------------------------------------------------------------------
    # Field grouping (date / range)
    # ------------------------------------------------------------------

    def group_field(
        self,
        field: str,
        *,
        by: str | None = None,
        start: float | None = None,
        end: float | None = None,
        interval: float | None = None,
        parent: str | None = None,
    ):
        """Group a cache field by date precision or numeric range.

        Date grouping::

            cache.group_field("order_date", by="months")
            # by ∈ {"years", "quarters", "months", "days",
            #       "hours", "minutes", "seconds"}

        Range grouping::

            cache.group_field("age", start=0, end=100, interval=10)

        Recursive grouping::

            cache.group_field("order_date", by="years")
            cache.group_field("order_date", by="quarters",
                              parent="order_date")  # nested

        Returns the registered :class:`FieldGroup`.
        """
        from ._group import (
            FieldGroup,
            FieldGroupDate,
            FieldGroupRange,
            synthesize_date_group_items,
            synthesize_range_group_items,
        )

        # Recursion-depth cap.
        if len(self.field_groups) >= 4:
            raise ValueError(
                "PivotCache.group_field: recursion depth exceeds 4"
            )

        field_index = self.field_index(field)
        parent_index = self.field_index(parent) if parent else None

        if by is not None:
            # Date grouping path.
            if start is None or end is None:
                # Pull from the source field's min_date/max_date.
                cf = self.fields[field_index]
                si = cf.shared_items
                if si.min_date is None or si.max_date is None:
                    raise ValueError(
                        f"PivotCache.group_field({field!r}): cannot infer "
                        "date range; pass start= and end= explicitly"
                    )
                from datetime import datetime as _dt

                start_dt = _dt.fromisoformat(si.min_date)
                end_dt = _dt.fromisoformat(si.max_date)
            else:
                from datetime import datetime as _dt

                start_dt = (
                    start if isinstance(start, _dt) else _dt.fromisoformat(str(start))
                )
                end_dt = (
                    end if isinstance(end, _dt) else _dt.fromisoformat(str(end))
                )
            items, start_iso, end_iso = synthesize_date_group_items(
                by, start_dt, end_dt
            )
            fg = FieldGroup(
                field_index=field_index,
                parent_index=parent_index,
                kind="date",
                date=FieldGroupDate(
                    group_by=by, start_date=start_iso, end_date=end_iso
                ),
                items=items,
            )
        else:
            # Numeric range grouping path.
            if start is None or end is None or interval is None:
                raise ValueError(
                    "PivotCache.group_field: range grouping requires "
                    "start=, end=, and interval=  (or by= for date grouping)"
                )
            start_f = _finite_float(start, "PivotCache.group_field start")
            end_f = _finite_float(end, "PivotCache.group_field end")
            interval_f = _finite_float(interval, "PivotCache.group_field interval")
            items = synthesize_range_group_items(start_f, end_f, interval_f)
            fg = FieldGroup(
                field_index=field_index,
                parent_index=parent_index,
                kind="range",
                range=FieldGroupRange(
                    start=start_f, end=end_f, interval=interval_f
                ),
                items=items,
            )

        self.field_groups.append(fg)
        return fg

    def to_rust_dict(self) -> dict:
        """Cache-definition dict for the Rust emitter."""
        if self._cache_id is None:
            raise RuntimeError(
                "PivotCache.cache_id not set — call "
                "Workbook.add_pivot_cache() first"
            )
        if self._fields is None:
            raise RuntimeError(
                "PivotCache._materialize not yet called"
            )
        return {
            "cache_id": self._cache_id,
            "source": self.source_to_rust_dict(),
            "fields": [f.to_rust_dict() for f in self._fields],
            "refresh_on_load": self.refresh_on_load,
            "refreshed_version": 6,
            "created_version": 6,
            "min_refreshable_version": 3,
            "refreshed_by": self.refreshed_by,
            "records_part_path": None,  # set by patcher
            # Extensions: calculated fields + field groups.
            "calculated_fields": [
                cf.to_rust_dict() for cf in self.calculated_fields
            ],
            "field_groups": [
                fg.to_rust_dict() for fg in self.field_groups
            ],
        }

    def to_rust_records_dict(self) -> dict:
        """Records dict for the Rust emitter.

        Each record cell is converted to inline form unless the field
        has an enumerable ``shared_items.items`` — in which case the
        cell is rewritten as ``{"kind": "index", "value": N}`` where
        N is the item's index.
        """
        if self._records is None or self._fields is None:
            raise RuntimeError(
                "PivotCache._materialize not yet called"
            )
        index_lookups = []
        for f in self._fields:
            si = f.shared_items
            if si.items is not None:
                # Map (kind, value) → 0-based index.
                lookup = {(v.kind, v.value): i for i, v in enumerate(si.items)}
                index_lookups.append(lookup)
            else:
                index_lookups.append(None)

        out_records = []
        for row in self._records:
            cells = []
            for col_i, cv in enumerate(row):
                lookup = index_lookups[col_i]
                if lookup is not None:
                    key = (cv.kind, cv.value)
                    if key in lookup:
                        cells.append({"kind": "index", "value": lookup[key]})
                        continue
                # Inline form.
                cells.append(cv.to_rust_dict())
            out_records.append(cells)

        return {
            "field_count": len(self._fields),
            "record_count": len(out_records),
            "records": out_records,
        }

    def source_to_rust_dict(self) -> dict:
        sheet = (
            self.source.worksheet.title
            if self.source.worksheet is not None
            else ""
        )
        ref = self._reference_to_a1()
        return {"sheet": sheet, "ref": ref, "name": None}

    def _reference_to_a1(self) -> str:
        from wolfxl.chart.reference import _index_to_col

        if self.source.min_col == self.source.max_col and self.source.min_row == self.source.max_row:
            return f"{_index_to_col(self.source.min_col)}{self.source.min_row}"
        return (
            f"{_index_to_col(self.source.min_col)}{self.source.min_row}:"
            f"{_index_to_col(self.source.max_col)}{self.source.max_row}"
        )


_install_openpyxl_iter(WorksheetSource, SharedItems, CacheField)
