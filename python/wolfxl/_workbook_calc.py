"""Workbook formula evaluation helpers."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from wolfxl.calc._protocol import RecalcResult


def calculate_workbook(wb: Any) -> dict[str, Any]:
    """Evaluate all formulas, cache the evaluator, and remember the results.

    The results are persisted as typed formula caches by the next save while
    the calculation inputs are unchanged.
    """
    from wolfxl._formula_cache import record_calculation
    from wolfxl.calc._evaluator import WorkbookEvaluator

    ev = WorkbookEvaluator()
    ev.load(wb)
    # Taken before evaluation replaces formula text with results, and copied
    # so what-if recalculate() perturbations cannot alter the saved snapshot.
    inputs = ev._cell_values.copy()  # noqa: SLF001
    result = ev.calculate()
    wb._evaluator = ev  # noqa: SLF001
    record_calculation(wb, inputs, result)
    return result


def cached_formula_values(wb: Any) -> dict[str, Any]:
    """Return Excel-saved cached formula results for every readable sheet."""
    if wb._rust_reader is None:  # noqa: SLF001
        return {}
    values: dict[str, Any] = {}
    for sheet_name in wb._sheet_names:  # noqa: SLF001
        sheet = wb._sheets[sheet_name]  # noqa: SLF001
        values.update(sheet.cached_formula_values(qualified=True))
    return values


def recalculate_workbook(
    wb: Any,
    perturbations: dict[str, float | int],
    tolerance: float = 1e-10,
) -> "RecalcResult":
    """Recompute affected formulas, loading and caching an evaluator if needed."""
    ev = wb._evaluator  # noqa: SLF001
    if ev is None:
        from wolfxl.calc._evaluator import WorkbookEvaluator

        ev = WorkbookEvaluator()
        ev.load(wb)
        ev.calculate()
        wb._evaluator = ev  # noqa: SLF001
    return ev.recalculate(perturbations, tolerance)
