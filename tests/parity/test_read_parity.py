"""Per-fixture read parity: open the same xlsx with openpyxl and wolfxl, diff.

Each xlsx fixture produces one test case. The test walks every sheet, every
populated cell, and compares:

* HARD dimensions — any mismatch fails the test.
* SOFT dimensions — recorded in ``ratchet.json``; never allowed to increase.
* INFO dimensions — reported but not gated.

Defined names, merged cells, and sheet dimensions are also diffed once per
workbook.

Large fixtures (>10k populated cells) are capped at the first 10k cells
visited to keep CI wall-clock bounded — the goal is per-row determinism, not
exhaustive byte-for-byte verification on megabyte-scale files.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

# Guard the wolfxl import so this file doesn't abort collection when the
# Rust wheel is missing (e.g. a fresh checkout that hasn't run ``maturin
# develop`` yet).
wolfxl = pytest.importorskip("wolfxl")
openpyxl = pytest.importorskip("openpyxl")

from ._scoring import ParityReport  # noqa: E402

MAX_CELLS_PER_SHEET = 10_000

RATCHET_PATH = Path(__file__).parent / "ratchet.json"


# Fixtures with known wolfxl read-side gaps. The harness still runs them
# (so we get measurement), but the test is xfail to keep CI green during
# Phase 0. Each entry must reference a documented gap in KNOWN_GAPS.md.
KNOWN_FIXTURE_GAPS: dict[str, str] = {}


def _load_ratchet() -> dict[str, dict[str, int]]:
    if not RATCHET_PATH.exists():
        return {}
    try:
        return json.loads(RATCHET_PATH.read_text())
    except json.JSONDecodeError:
        return {}


def _ratchet_key(fixture_path: Path) -> str:
    return "/".join(fixture_path.parts[-2:])


def _diff_workbooks(
    report: ParityReport,
    op_wb: Any,
    wx_wb: Any,
) -> None:
    op_names = list(op_wb.sheetnames)
    wx_names = list(wx_wb.sheetnames)
    report.record("sheet_names", "workbook", op_names, wx_names)

    # Intersect so mismatched names don't cascade into per-sheet noise;
    # name-diff is already captured above.
    common = [n for n in op_names if n in wx_names]
    for sheet in common:
        _diff_sheet(report, op_wb[sheet], wx_wb[sheet], sheet)

    op_named = dict(op_wb.defined_names)
    # Both wolfxl (T1) and openpyxl now return DefinedName objects with
    # a ``.value`` attribute — pull the refers_to string off both sides
    # before comparing. Strip the leading ``=`` to normalize.
    def _resolve(defn: Any) -> str | None:
        refers = getattr(defn, "value", defn)
        if isinstance(refers, str) and refers.startswith("="):
            refers = refers[1:]
        return refers

    op_named_resolved: dict[str, str | None] = {
        name: _resolve(defn) for name, defn in op_named.items()
    }
    wx_named_raw = wx_wb.defined_names
    wx_named_resolved: dict[str, str | None] = {
        name: _resolve(defn) for name, defn in wx_named_raw.items()
    }
    for name in set(op_named_resolved) | set(wx_named_resolved):
        report.record(
            "defined_name.refers_to",
            f"defined_name:{name}",
            op_named_resolved.get(name),
            wx_named_resolved.get(name),
        )


def _diff_sheet(
    report: ParityReport, op_ws: Any, wx_ws: Any, sheet_name: str,
) -> None:
    # Dimensions — HARD tier.
    op_max_row = op_ws.max_row
    op_max_col = op_ws.max_column

    wx_max_row = wx_ws.max_row
    wx_max_col = wx_ws.max_column

    report.record("max_row", f"{sheet_name}:dim", op_max_row, wx_max_row)
    report.record("max_col", f"{sheet_name}:dim", op_max_col, wx_max_col)

    # Freeze panes.
    report.record(
        "freeze_panes",
        f"{sheet_name}:freeze",
        op_ws.freeze_panes,
        getattr(wx_ws, "freeze_panes", None),
    )

    # Merged cells — compare as sets of range strings.
    op_merged = {str(r) for r in op_ws.merged_cells.ranges}
    wx_merged: set[str] = set()
    if hasattr(wx_ws, "merged_cells") and hasattr(wx_ws.merged_cells, "ranges"):
        wx_merged = {str(r) for r in wx_ws.merged_cells.ranges}
    elif hasattr(wx_ws, "_merged_ranges"):
        wx_merged = set(wx_ws._merged_ranges)  # noqa: SLF001
    report.record("merged_cells", f"{sheet_name}:merged", op_merged, wx_merged)

    # Cell-level diff — HARD on value and number_format.
    cells_seen = 0
    for row in op_ws.iter_rows():
        for op_cell in row:
            if op_cell.value is None and getattr(op_cell, "style_id", 0) == 0:
                continue
            if cells_seen >= MAX_CELLS_PER_SHEET:
                return
            cells_seen += 1
            coord = op_cell.coordinate
            wx_cell = wx_ws[coord]

            report.record(
                "value",
                f"{sheet_name}!{coord}",
                _normalize_value(op_cell.value),
                _normalize_value(wx_cell.value),
            )
            report.record(
                "number_format",
                f"{sheet_name}!{coord}",
                _normalize_number_format(op_cell.number_format),
                _normalize_number_format(wx_cell.number_format),
            )


def _normalize_value(v: Any) -> Any:
    """Normalize values for comparison.

    openpyxl sometimes yields ``None`` where wolfxl yields empty string for
    blank cells — ``_scoring._values_equal`` already treats those as equal.
    This hook is here for future per-type normalization (e.g. trimming).
    """
    return v


def _normalize_number_format(v: Any) -> Any:
    """Normalize archived pre-fix ``None`` values to openpyxl's ``'General'``.

    Current wolfxl matches openpyxl here; the coercion keeps the harness useful
    for old baseline wheels without hiding other number-format mismatches.
    """
    if v is None:
        return "General"
    return v


def test_read_parity(xlsx_fixture: Path) -> None:
    """Every HARD dimension must match; SOFT mismatches ratcheted."""
    fixture_id = _ratchet_key(xlsx_fixture)
    if fixture_id in KNOWN_FIXTURE_GAPS:
        pytest.xfail(KNOWN_FIXTURE_GAPS[fixture_id])
    report = ParityReport(fixture_id=fixture_id)

    op_wb = openpyxl.load_workbook(xlsx_fixture, data_only=True)
    wx_wb = wolfxl.load_workbook(str(xlsx_fixture), data_only=True)
    try:
        _diff_workbooks(report, op_wb, wx_wb)
    finally:
        wx_wb.close()
        op_wb.close()

    hard = report.hard_failures()
    if hard:
        lines = "\n".join(str(m) for m in hard[:20])
        remaining = len(hard) - 20
        extra = f"\n... +{remaining} more" if remaining > 0 else ""
        pytest.fail(
            f"{len(hard)} HARD parity mismatches on {report.fixture_id}:\n"
            f"{lines}{extra}",
        )

    # SOFT failures are ratcheted — don't fail, but the ratchet-enforcement
    # test below catches any increase.
    _update_ratchet_observation(report)


_OBSERVATIONS: dict[str, dict[str, int]] = {}


def _update_ratchet_observation(report: ParityReport) -> None:
    summary = report.summary()
    _OBSERVATIONS[report.fixture_id] = summary


def test_zz_ratchet_soft_failures_nondecreasing() -> None:
    """SOFT mismatch count must never rise above the committed baseline.

    This depends on observations collected by every ``test_read_parity``
    case, so it must run *after* all of them. pytest collects tests in
    file declaration order by default, but third-party plugins (e.g.
    pytest-xdist, pytest-randomly) can reorder collection — the ``zz_``
    prefix sorts this test last under any plain alphabetical ordering and
    keeps the dependency robust without a session finalizer.
    """
    ratchet = _load_ratchet()
    if not _OBSERVATIONS:
        pytest.skip("No fixtures observed — cannot check ratchet")

    violations: list[str] = []
    for fixture_id, observed in _OBSERVATIONS.items():
        baseline = ratchet.get(fixture_id, {"hard": 0, "soft": 0, "info": 0})
        if observed["soft"] > baseline["soft"]:
            violations.append(
                f"{fixture_id}: SOFT {observed['soft']} > baseline {baseline['soft']}",
            )

    if violations:
        pytest.fail(
            "Soft-tier parity regressed vs ratchet.json:\n" + "\n".join(violations),
        )
