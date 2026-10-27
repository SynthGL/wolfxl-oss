"""Compact-form pivot layout: axis items, location, and rendered cells.

Excel stores a pivot table twice: the ``pivotTableDefinition`` part
describes the layout (``rowItems``/``colItems``/``location``) and the
worksheet holds the rendered values. Excel validates the two against
each other on open and does not re-render a pivot until the user
refreshes it, so both must describe the same compact-form grid that
Excel itself would produce.

Grid conventions (matched against Excel 16 refresh output):

* Row fields render compact: every level gets its own row, the parent
  row carries the subtotal, and the grand total row comes last.
* Column fields render one header row per field; a field with a real
  field nested under it gets ``"<item> Total"`` subtotal columns.
* Two or more data fields add the ``Values`` pseudo-field (index -2) as
  the innermost column field.
* Item order follows the cache's shared-items order, and items without
  records are hidden.
* Calculated items follow the regular items; totals include them.
  Calculated fields apply their formula to the field sums of each cell.
"""

from __future__ import annotations

import statistics
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Callable, Sequence

GRAND_TOTAL_CAPTION = "Grand Total"
ROW_HEADER_CAPTION = "Row Labels"
COL_HEADER_CAPTION = "Column Labels"
BLANK_ITEM_CAPTION = "(blank)"
ALL_ITEMS_CAPTION = "(All)"


# "Show values as" options that depend only on the rendered grid's own
# totals. The base-relative ones (difference, percent, percentDiff,
# runTotal) are left for Excel to compute on open.
GRID_SHOW_DATA_AS = frozenset({"percentOfTotal", "percentOfRow", "percentOfCol", "index"})


def renders_show_data_as(show_data_as: str | None) -> bool:
    """Whether WolfXL renders values for this ``show_data_as`` option."""
    return show_data_as in (None, "normal") or show_data_as in GRID_SHOW_DATA_AS


@dataclass
class DataSpec:
    """One data field: cache field index, aggregator, caption, display mode.

    A calculated field sets ``calculated``; ``inputs`` lists the source
    fields its formula reads, and ``formula(resolve)`` evaluates it with
    ``resolve(field_index)`` giving that field's sum in the cell.
    ``formula`` is ``None`` when WolfXL cannot evaluate the formula.
    """

    field_index: int
    function: str
    caption: str
    show_data_as: str | None = None
    calculated: bool = False
    inputs: tuple[int, ...] = ()
    formula: Callable[[Callable[[int], Any]], Any] | None = None


@dataclass
class CalcItemSpec:
    """A calculated item of an axis field.

    ``refs`` are the item indices its formula reads; ``evaluate(resolve)``
    computes it with ``resolve(item_index)`` giving that item's value.
    ``evaluate`` is ``None`` when WolfXL cannot evaluate the formula.
    """

    field_index: int
    item_index: int
    refs: frozenset[int]
    evaluate: Callable[[Callable[[int], Any]], Any] | None


class _Unrenderable(Exception):
    """A cell value WolfXL leaves for Excel to compute."""

@dataclass
class PivotLayout:
    """Excel-consistent layout for one pivot table.

    Cell offsets are relative to the table's top-left cell (the
    ``location`` anchor). Page-field cells are relative to the same
    anchor and therefore have negative row offsets.
    """

    row_items: list[dict]
    col_items: list[dict]
    height: int
    width: int
    first_data_row: int
    first_data_col: int
    cells: dict[tuple[int, int], Any] = field(default_factory=dict)
    # True when some cells were left for Excel to compute on refresh.
    unrendered: bool = False


class _Accumulator:
    __slots__ = ("numbers", "non_blank")

    def __init__(self) -> None:
        self.numbers: list[float] = []
        self.non_blank = 0


def aggregate(acc: _Accumulator, function: str) -> float | None:
    """Apply an Excel pivot aggregator; ``None`` renders as an empty cell."""
    nums = acc.numbers
    if function == "count":
        return float(acc.non_blank)
    if function == "countNums":
        return float(len(nums))
    if not nums:
        return None
    if function == "sum":
        return float(sum(nums))
    if function == "average":
        return float(sum(nums) / len(nums))
    if function == "max":
        return float(max(nums))
    if function == "min":
        return float(min(nums))
    if function == "product":
        product = 1.0
        for n in nums:
            product *= n
        return product
    if function == "stdDev":
        return statistics.stdev(nums) if len(nums) > 1 else None
    if function == "stdDevp":
        return statistics.pstdev(nums)
    if function == "var":
        return statistics.variance(nums) if len(nums) > 1 else None
    if function == "varp":
        return statistics.pvariance(nums)
    raise ValueError(f"Unknown aggregator function: {function!r}")


def item_label(value: Any) -> Any:
    """Cell value Excel shows for a shared item."""
    if value is None or value.kind == "missing":
        return BLANK_ITEM_CAPTION
    if value.kind == "date":
        try:
            return datetime.fromisoformat(value.value)
        except ValueError:
            return value.value
    return value.value


def _repeat_count(path: tuple, previous: tuple | None) -> int:
    """Leading entries shared with the previous axis item (``<i r>``).

    At least one entry is always written, so a child row that follows its
    parent row repeats the parent's full path.
    """
    if previous is None:
        return 0
    limit = min(len(path) - 1, len(previous))
    r = 0
    while r < limit and path[r] == previous[r]:
        r += 1
    return r


def _axis_item(
    path: tuple, previous: tuple | None, *, t: str | None = None, i: int | None = None
) -> dict:
    r = _repeat_count(path, previous)
    return {
        "indices": [int(x) for x in path[r:]],
        "t": t,
        "r": r or None,
        "i": i or None,
    }


def _children(keys: Sequence[tuple], depth: int) -> dict[tuple, list[int]]:
    """Map each key prefix to its sorted child indices."""
    tree: dict[tuple, set[int]] = {}
    for key in keys:
        for level in range(depth):
            tree.setdefault(key[:level], set()).add(key[level])
    return {prefix: sorted(kids) for prefix, kids in tree.items()}


def compute_layout(
    *,
    records: Sequence[Sequence[Any]],
    item_index: Any,
    shared_items: Sequence[Sequence[Any] | None],
    row_fields: Sequence[int],
    col_fields: Sequence[int],
    data: Sequence[DataSpec],
    row_grand_totals: bool,
    col_grand_totals: bool,
    grand_total_caption: str | None,
    data_caption: str,
    calculated_items: Sequence[CalcItemSpec] = (),
) -> PivotLayout:
    """Compute axis items, location metadata, and rendered cells.

    ``item_index(record_cell, field_index)`` maps a record cell to its
    shared-items index for an axis field. ``calculated_items`` lists the
    calculated items of row and column fields.
    """
    n_rows, n_cols, n_data = len(row_fields), len(col_fields), len(data)
    multi_data = n_data > 1
    grand_caption = grand_total_caption or GRAND_TOTAL_CAPTION

    # 1. Accumulate every (row prefix, col prefix) bucket in one pass, one
    #    accumulator per measured source field.
    measured = sorted(
        {f for d in data for f in (d.inputs if d.calculated else (d.field_index,))}
    )
    slot = {f: n for n, f in enumerate(measured)}
    buckets: dict[tuple[tuple, tuple], list[_Accumulator]] = {}
    row_keys: set[tuple] = set()
    col_keys: set[tuple] = set()
    for record in records:
        rk = tuple(item_index(record[f], f) for f in row_fields)
        ck = tuple(item_index(record[f], f) for f in col_fields)
        row_keys.add(rk)
        col_keys.add(ck)
        cells = [record[f] if f < len(record) else None for f in measured]
        for rl in range(n_rows + 1):
            for cl in range(n_cols + 1):
                accs = buckets.get((rk[:rl], ck[:cl]))
                if accs is None:
                    accs = [_Accumulator() for _ in measured]
                    buckets[(rk[:rl], ck[:cl])] = accs
                for acc, cell in zip(accs, cells):
                    if cell is None or cell.kind == "missing":
                        continue
                    acc.non_blank += 1
                    if cell.kind == "number":
                        acc.numbers.append(cell.value)

    # Calculated items add axis keys next to the items they read, keyed by
    # axis level: {level: {item_index: spec}}.
    def calc_levels(fields: Sequence[int], keys: set[tuple]) -> dict[int, dict]:
        levels: dict[int, dict[int, CalcItemSpec]] = {}
        for spec in calculated_items:
            if spec.field_index not in fields:
                continue
            level = list(fields).index(spec.field_index)
            levels.setdefault(level, {})[spec.item_index] = spec
            for key in list(keys):
                if not spec.refs or key[level] in spec.refs:
                    keys.add(key[:level] + (spec.item_index,) + key[level + 1 :])
        return levels

    row_calc = calc_levels(row_fields, row_keys)
    col_calc = calc_levels(col_fields, col_keys)
    row_tree = _children(sorted(row_keys), n_rows)
    col_tree = _children(sorted(col_keys), n_cols)

    def field_value(row_path: tuple, col_path: tuple, f: int, function: str) -> Any:
        """Aggregate of source field ``f`` in a cell, calculated items included."""
        for on_rows, path, levels in (
            (True, row_path, row_calc),
            (False, col_path, col_calc),
        ):
            for level, specs in levels.items():
                if level < len(path) and path[level] in specs:
                    spec = specs[path[level]]
                    if spec.evaluate is None or function != "sum":
                        raise _Unrenderable

                    def resolve(item: int, path=path, level=level, on_rows=on_rows) -> Any:
                        sub = path[:level] + (item,) + path[level + 1 :]
                        v = field_value(
                            sub if on_rows else row_path,
                            col_path if on_rows else sub,
                            f,
                            function,
                        )
                        return 0.0 if v is None else v

                    return spec.evaluate(resolve)
        # A total over an axis level with calculated items sums the items
        # it shows, calculated ones included.
        for on_rows, path, levels, tree in (
            (True, row_path, row_calc, row_tree),
            (False, col_path, col_calc, col_tree),
        ):
            if levels and len(path) <= max(levels):
                if function != "sum":
                    raise _Unrenderable
                total: Any = None
                for kid in tree.get(path, []):
                    sub = path + (kid,)
                    v = field_value(
                        sub if on_rows else row_path,
                        col_path if on_rows else sub,
                        f,
                        function,
                    )
                    if isinstance(v, str):
                        return v
                    if v is not None:
                        total = v if total is None else total + v
                return total
        accs = buckets.get((row_path, col_path))
        if accs is None:
            return None
        return aggregate(accs[slot[f]], function)

    def value(row_path: tuple, col_path: tuple, di: int) -> Any:
        d = data[di]
        if not d.calculated:
            return field_value(row_path, col_path, d.field_index, d.function)
        if d.formula is None or d.function != "sum":
            raise _Unrenderable
        sums = {f: field_value(row_path, col_path, f, "sum") for f in d.inputs}
        if d.inputs and all(v is None for v in sums.values()):
            return None
        return d.formula(lambda f: 0.0 if sums[f] is None else sums[f])

    def shown(row_path: tuple, col_path: tuple, di: int) -> Any:
        """``value`` after the data field's "show values as" option.

        Excel shows 0 rather than an empty cell for an empty intersection
        under the grid-relative options.
        """
        v = value(row_path, col_path, di)
        how = data[di].show_data_as
        if how in (None, "normal"):
            return v
        if how not in GRID_SHOW_DATA_AS:
            return None
        if v is None:
            v = 0.0
        total = value((), (), di)
        if how == "percentOfTotal":
            numerator, denominator = v, total
        elif how == "percentOfRow":
            numerator, denominator = v, value(row_path, (), di)
        elif how == "percentOfCol":
            numerator, denominator = v, value((), col_path, di)
        else:  # index
            row_total = value(row_path, (), di)
            col_total = value((), col_path, di)
            if total is None or row_total is None or col_total is None:
                return None
            for part in (v, total, row_total, col_total):
                if isinstance(part, str):
                    return part
            numerator, denominator = v * total, row_total * col_total
        for part in (numerator, denominator):
            if isinstance(part, str):
                return part
        if not denominator:
            return None
        return numerator / denominator

    def label(field_index: int, item: int) -> Any:
        items = shared_items[field_index]
        if items is None or item >= len(items):
            return BLANK_ITEM_CAPTION
        return item_label(items[item])

    # 2. Row lines: compact form, one line per tree node, parents first.
    row_lines: list[tuple[tuple, str | None]] = []

    def walk_rows(prefix: tuple) -> None:
        for kid in row_tree.get(prefix, []):
            path = prefix + (kid,)
            row_lines.append((path, None))
            if len(path) < n_rows:
                walk_rows(path)

    if n_rows:
        walk_rows(())
        if row_grand_totals:
            row_lines.append(((), "grand"))

    row_items: list[dict] = []
    previous: tuple | None = None
    for path, kind in row_lines:
        if kind == "grand":
            row_items.append({"indices": [0], "t": "grand", "r": None, "i": None})
            previous = None
        else:
            row_items.append(_axis_item(path, previous))
            previous = path
    if not row_items:
        row_items.append({"indices": [], "t": None, "r": None, "i": None})

    # 3. Column lines: leaves, outer-field subtotals, then grand totals.
    #    Each line is (col path, data index, kind, subtotal level).
    col_lines: list[tuple[tuple, int, str | None]] = []
    data_range = range(n_data)

    def walk_cols(prefix: tuple) -> None:
        if len(prefix) == n_cols:
            col_lines.extend((prefix, di, None) for di in data_range)
            return
        for kid in col_tree.get(prefix, []):
            path = prefix + (kid,)
            walk_cols(path)
            if len(path) < n_cols:
                col_lines.extend((path, di, "default") for di in data_range)

    if n_cols or multi_data:
        walk_cols(())
        if n_cols and col_grand_totals:
            col_lines.extend(((), di, "grand") for di in data_range)
    else:
        col_lines.append(((), 0, None))

    col_items: list[dict] = []
    col_repeats: list[int] = []
    previous = None
    for path, di, kind in col_lines:
        full = path + (di,) if multi_data and kind is None else path
        if kind == "grand":
            col_items.append({"indices": [0], "t": "grand", "r": None, "i": di or None})
            col_repeats.append(0)
            previous = None
            continue
        if not full:
            col_items.append({"indices": [], "t": None, "r": None, "i": None})
            col_repeats.append(0)
            continue
        item = _axis_item(full, previous, t=kind, i=di if multi_data else None)
        col_items.append(item)
        col_repeats.append(item["r"] or 0)
        previous = full

    # 4. Geometry. Every column-axis field (including Values) adds a header
    #    row below the caption row. The label column holds row items, or
    #    the single data caption when only column fields are present.
    n_col_levels = n_cols + (1 if multi_data else 0)
    has_label_col = bool(n_rows) or (bool(n_cols) and not multi_data)
    header_rows = 1 + n_col_levels
    first_data_col = 1 if has_label_col else 0
    data_rows = len(row_lines) if n_rows else 1
    layout = PivotLayout(
        row_items=row_items,
        col_items=col_items,
        height=header_rows + data_rows,
        width=first_data_col + len(col_lines),
        first_data_row=header_rows,
        first_data_col=first_data_col,
    )
    cells = layout.cells

    # 5. Header block.
    def col_caption(di: int) -> str:
        return data[di].caption

    if n_col_levels:
        if n_rows and not multi_data:
            cells[(0, 0)] = col_caption(0)
        cells[(0, first_data_col)] = COL_HEADER_CAPTION if n_cols else data_caption
        if n_rows:
            cells[(header_rows - 1, 0)] = ROW_HEADER_CAPTION
        for j, (path, di, kind) in enumerate(col_lines):
            col = first_data_col + j
            if kind == "grand":
                cells[(1, col)] = (
                    f"Total {col_caption(di)}" if multi_data else grand_caption
                )
                continue
            start = col_repeats[j]
            for level in range(start, len(path)):
                item = label(col_fields[level], path[level])
                if kind == "default" and level == len(path) - 1:
                    item = (
                        f"{item} {col_caption(di)}" if multi_data else f"{item} Total"
                    )
                cells[(1 + level, col)] = item
            if multi_data and kind is None:
                cells[(header_rows - 1, col)] = col_caption(di)
    else:
        if n_rows:
            cells[(0, 0)] = ROW_HEADER_CAPTION
        for j, (_path, di, _kind) in enumerate(col_lines):
            cells[(0, first_data_col + j)] = col_caption(di)

    # 6. Data block.
    data_lines: list[tuple[tuple, str | None]] = row_lines if n_rows else [((), None)]
    for i, (row_path, kind) in enumerate(data_lines):
        row = header_rows + i
        if n_rows:
            cells[(row, 0)] = (
                grand_caption
                if kind == "grand"
                else label(row_fields[len(row_path) - 1], row_path[-1])
            )
        elif has_label_col and not multi_data:
            cells[(row, 0)] = col_caption(0)
        for j, (col_path, di, _kind) in enumerate(col_lines):
            try:
                v = shown(row_path, col_path, di)
            except _Unrenderable:
                layout.unrendered = True
                continue
            if v is not None:
                cells[(row, first_data_col + j)] = v

    return layout
