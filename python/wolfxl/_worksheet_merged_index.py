"""Bounded point index for worksheet merged ranges."""

from __future__ import annotations

from bisect import bisect_right
from dataclasses import dataclass
from typing import Any, Iterable

from wolfxl.utils.cell import range_boundaries


@dataclass(frozen=True, slots=True)
class _MergedRangeBounds:
    min_row: int
    min_col: int
    max_row: int
    max_col: int


@dataclass(frozen=True, slots=True)
class _RowIntervalNode:
    center: int
    by_start: tuple[_MergedRangeBounds, ...]
    starts: tuple[int, ...]
    by_end: tuple[_MergedRangeBounds, ...]
    negative_ends: tuple[int, ...]
    by_column: tuple[_MergedRangeBounds, ...]
    column_starts: tuple[int, ...]
    disjoint_columns: bool
    left: _RowIntervalNode | None
    right: _RowIntervalNode | None


class MergedRangeIndex:
    """Static interval tree that retains one parsed entry per merged range.

    Rows choose a narrow interval-tree path; matching candidates are then
    checked against their column bounds. A range is stored at exactly one node,
    so large rectangles never expand into per-cell or per-row storage.
    """

    __slots__ = ("_entries", "_root")

    def __init__(self, refs: Iterable[str]) -> None:
        self._entries = tuple(
            bounds
            for ref in sorted({str(ref) for ref in refs})
            if (bounds := _parse_range_bounds(ref)) is not None
        )
        self._root = _build_row_interval_tree(self._entries)

    def is_subordinate(self, row: int, col: int) -> bool:
        """Return whether ``(row, col)`` is covered by a non-anchor cell."""
        node = self._root
        while node is not None:
            if row < node.center:
                candidates = node.by_start
                candidate_count = bisect_right(node.starts, row)
                node = node.left
            elif row > node.center:
                candidates = node.by_end
                candidate_count = bisect_right(node.negative_ends, -row)
                node = node.right
            else:
                candidates = node.by_start
                candidate_count = len(candidates)
                node = None
            for candidate_index in range(candidate_count):
                bounds = candidates[candidate_index]
                if bounds.min_col <= col <= bounds.max_col and not (
                    row == bounds.min_row and col == bounds.min_col
                ):
                    return True
        return False

    def bounds_for_cell(
        self, row: int, col: int,
        priorities: dict[tuple[int, int, int, int], int] | None = None,
    ) -> tuple[int, int, int, int] | None:
        """Find a containing range, retaining first-match overlap precedence.

        Every bucket spans its row center. Valid nonoverlapping rectangles
        therefore have disjoint columns in that bucket, allowing one column
        bisect rather than scanning all ranges on wide merged headers.
        Overlapping inputs retain the conservative candidate scan and choose
        the lowest caller rank across the entire row-tree path.
        """
        node = self._root
        found = None
        found_rank = len(self._entries)
        while node is not None:
            current = node
            if row < node.center:
                candidates = node.by_start
                candidate_count = bisect_right(node.starts, row)
            elif row > node.center:
                candidates = node.by_end
                candidate_count = bisect_right(node.negative_ends, -row)
            else:
                candidates = node.by_start
                candidate_count = len(candidates)
            # Row navigation must continue after a match: an accepted overlap
            # may live in an ancestor and another in a descendant bucket.
            node = current.left if row < current.center else (
                current.right if row > current.center else None
            )
            if current.disjoint_columns:
                column_index = bisect_right(current.column_starts, col) - 1
                candidates = current.by_column
                first_candidate = max(0, column_index)
                candidate_count = column_index + 1
            else:
                first_candidate = 0
            for candidate_index in range(first_candidate, candidate_count):
                bounds = candidates[candidate_index]
                if not (
                    bounds.min_row <= row <= bounds.max_row
                    and bounds.min_col <= col <= bounds.max_col
                ):
                    continue
                value = (bounds.min_row, bounds.min_col, bounds.max_row, bounds.max_col)
                rank = priorities[value] if priorities is not None else 0
                if found is None or rank < found_rank:
                    found, found_rank = value, rank
        return found


def merged_range_index(ws: Any, refs: Iterable[str]) -> MergedRangeIndex:
    """Return the worksheet's lazily-built index for its current merge refs."""
    cached = ws._merged_range_index  # noqa: SLF001
    if cached is None:
        cached = MergedRangeIndex(refs)
        ws._merged_range_index = cached  # noqa: SLF001
    return cached


def invalidate_merged_range_index(ws: Any) -> None:
    """Drop stale parsed bounds after merge metadata or sheet coordinates change."""
    ws._merged_range_index = None  # noqa: SLF001


def _parse_range_bounds(ref: str) -> _MergedRangeBounds | None:
    try:
        min_col, min_row, max_col, max_row = range_boundaries(ref)
    except Exception:
        return None
    if min_col is None or min_row is None or max_col is None or max_row is None:
        return None
    if min_row > max_row or min_col > max_col:
        return None
    return _MergedRangeBounds(
        min_row=int(min_row),
        min_col=int(min_col),
        max_row=int(max_row),
        max_col=int(max_col),
    )


def _build_row_interval_tree(
    entries: tuple[_MergedRangeBounds, ...],
) -> _RowIntervalNode | None:
    if not entries:
        return None
    endpoints = sorted(
        endpoint
        for bounds in entries
        for endpoint in (bounds.min_row, bounds.max_row)
    )
    center = endpoints[len(endpoints) // 2]
    left: list[_MergedRangeBounds] = []
    right: list[_MergedRangeBounds] = []
    spanning: list[_MergedRangeBounds] = []
    for bounds in entries:
        if bounds.max_row < center:
            left.append(bounds)
        elif bounds.min_row > center:
            right.append(bounds)
        else:
            spanning.append(bounds)
    by_start = tuple(sorted(spanning, key=lambda bounds: (bounds.min_row, bounds.min_col)))
    by_end = tuple(
        sorted(spanning, key=lambda bounds: (-bounds.max_row, bounds.min_col))
    )
    by_column = tuple(sorted(spanning, key=lambda bounds: (bounds.min_col, bounds.max_col)))
    disjoint_columns = all(
        left.max_col < right.min_col
        for left, right in zip(by_column, by_column[1:])
    )
    return _RowIntervalNode(
        center=center,
        by_start=by_start,
        starts=tuple(bounds.min_row for bounds in by_start),
        by_end=by_end,
        negative_ends=tuple(-bounds.max_row for bounds in by_end),
        by_column=by_column,
        column_starts=tuple(bounds.min_col for bounds in by_column),
        disjoint_columns=disjoint_columns,
        left=_build_row_interval_tree(tuple(left)),
        right=_build_row_interval_tree(tuple(right)),
    )
