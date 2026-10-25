"""Cross-surface parity test for the wolfxl-core classifier bridge.

The goal of this test is to catch drift between the two Python
surfaces that go through `wolfxl_core` classifiers:

1. **CLI**: `wolfxl schema <file> --format json` (wolfxl-cli 0.8.0+,
   built from this workspace via `cargo run -p wolfxl-cli`).
2. **Python**: `wolfxl.load_workbook(...)` + `worksheet.schema()`
   (the PyO3 bridge, task #22a + #22b).

Both paths ultimately call into `wolfxl_core::infer_sheet_schema`, so
they *should* agree on every field. Where they don't (see "Known
divergences" below), those gaps are pre-existing cdylib limitations
that the sprint-3 "Option A" engine-collapse work is meant to close —
the test documents them rather than failing, so real drift is still
surfaced on every run.

## Known divergences (pre-Option-A)

- **Numeric type (`int` vs `float`)**: the cdylib's reader widens
  calamine's `Data::Int(N)` to a Python `float`. `wolfxl_core`'s
  inference sees `CellValue::Float` on the Python path and
  `CellValue::Int` on the CLI path. Fix: Option-A (Python delegates
  reading to `wolfxl_core::Workbook::open`). Until then, numeric
  columns may differ between `int` / `float` / `mixed`.
- **Format category** for openpyxl-generated workbooks: the cdylib's
  `resolved_number_format` doesn't yet route through the task-#9
  styles walker that `wolfxl-core` has. Workbooks authored by real
  Excel go through calamine's fast path and agree on both surfaces;
  openpyxl-styled workbooks may report `general` on Python while the
  CLI reports `currency` / `percentage` / `date`.
- **Trailing empty columns**: Python's `worksheet._bounds_max_col()` honors
  the sheet's `<dimension ref="A1:X">` tag (openpyxl compat), which
  can include trailing columns with formatting but no values.
  `wolfxl_core`'s `Sheet::from_calamine` trims those. We strip
  trailing "all-empty, unnamed" columns from both sides before
  comparing so the test doesn't flag a reader-layer disagreement
  that's unrelated to the classifier.

Both divergences narrow the parity bar, but the surface-agnostic
fields (sheet names, row counts, column names, null counts, unique
counts, cardinality, sample values) must match exactly — the test
fails if any of them drift.
"""

from __future__ import annotations

import json
import os
import subprocess
from pathlib import Path

import pytest

import wolfxl

REPO_ROOT = Path(__file__).parent.parent
FIXTURE = (
    REPO_ROOT
    / "crates"
    / "wolfxl-core"
    / "tests"
    / "fixtures"
    / "sample-financials.xlsx"
)


def _cli_schema(path: Path) -> dict:
    """Run `cargo run -p wolfxl-cli -- schema <file> --format json`.

    Uses the workspace build (not whatever's on PATH) so the test is
    always comparing the same wolfxl-core version Python's bridge
    links against.
    """
    result = subprocess.run(
        [
            "cargo",
            "run",
            "--quiet",
            "--release",
            "-p",
            "wolfxl-cli",
            "--",
            "schema",
            str(path),
            "--format",
            "json",
        ],
        capture_output=True,
        text=True,
        cwd=REPO_ROOT,
        timeout=120,
        env={**dict(os.environ), "CARGO_TERM_COLOR": "never"},
    )
    assert result.returncode == 0, (
        f"wolfxl schema failed (returncode={result.returncode}):\n"
        f"stderr: {result.stderr}"
    )
    return json.loads(result.stdout)


def _python_schemas(path: Path) -> list[dict]:
    """Mirror `wolfxl schema`'s structure from the Python bridge."""
    wb = wolfxl.load_workbook(str(path))
    return [wb[name].schema() for name in wb.sheetnames]


def _trim_trailing_empty(schema: dict) -> dict:
    """Drop trailing columns that are empty-named and all-null.

    Normalizes away the Python-side "trailing empty columns kept for
    openpyxl dimension-tag parity" divergence documented at the top
    of this file. Does not touch columns in the middle of the sheet
    (those would be a real data shape difference, and we want the
    test to flag them).
    """
    cols = list(schema["columns"])
    while cols and cols[-1]["name"] == "" and cols[-1]["null_count"] == schema["rows"]:
        cols.pop()
    return {**schema, "columns": cols}


@pytest.mark.skipif(
    not FIXTURE.exists(),
    reason=f"fixture {FIXTURE} not present (lives alongside wolfxl-core tests)",
)
@pytest.mark.slow
def test_schema_parity_structural() -> None:
    """Sheet count, row counts, column names / metadata must match."""
    cli = _cli_schema(FIXTURE)
    py = _python_schemas(FIXTURE)

    cli_sheets = cli["sheets"]
    assert len(cli_sheets) == len(py), (
        f"sheet count mismatch: cli={len(cli_sheets)}, py={len(py)}"
    )

    for cli_sheet, py_sheet_raw in zip(cli_sheets, py):
        py_sheet = _trim_trailing_empty(py_sheet_raw)
        assert cli_sheet["name"] == py_sheet["name"], (
            f"sheet name mismatch: cli={cli_sheet['name']!r}, "
            f"py={py_sheet['name']!r}"
        )
        assert cli_sheet["rows"] == py_sheet["rows"], (
            f"sheet {cli_sheet['name']!r} row count mismatch: "
            f"cli={cli_sheet['rows']}, py={py_sheet['rows']}"
        )
        cli_cols = cli_sheet["columns"]
        py_cols = py_sheet["columns"]
        assert len(cli_cols) == len(py_cols), (
            f"sheet {cli_sheet['name']!r} column count mismatch: "
            f"cli={len(cli_cols)}, py={len(py_cols)}"
        )

        for cli_col, py_col in zip(cli_cols, py_cols):
            assert cli_col["name"] == py_col["name"], (
                f"column name mismatch in {cli_sheet['name']!r}: "
                f"cli={cli_col['name']!r}, py={py_col['name']!r}"
            )
            # Structural fields that don't depend on int/float widening
            # or on the styles-walker fallback. These must be exact.
            for field in (
                "null_count",
                "unique_count",
                "unique_capped",
                "cardinality",
            ):
                assert cli_col[field] == py_col[field], (
                    f"column {cli_col['name']!r} in sheet "
                    f"{cli_sheet['name']!r}: {field} mismatch "
                    f"(cli={cli_col[field]!r}, py={py_col[field]!r})"
                )


@pytest.mark.skipif(
    not FIXTURE.exists(),
    reason=f"fixture {FIXTURE} not present (lives alongside wolfxl-core tests)",
)
@pytest.mark.slow
def test_schema_parity_samples() -> None:
    """Sample-value lists must match as sets.

    `samples` is the small per-column preview (≤3 rendered values).
    Order can diverge if the two surfaces walk the grid differently,
    so compare as sets rather than as lists — drift in membership is
    what matters, ordering churn is noise.
    """
    cli = _cli_schema(FIXTURE)
    py = _python_schemas(FIXTURE)

    for cli_sheet, py_sheet_raw in zip(cli["sheets"], py):
        py_sheet = _trim_trailing_empty(py_sheet_raw)
        for cli_col, py_col in zip(
            cli_sheet["columns"],
            py_sheet["columns"],
        ):
            # Normalize to strings before comparing — int/float widening
            # would otherwise cause "420000" vs "420000.0" drift. Set
            # compare because order isn't semantically meaningful.
            cli_samples = {_normalize_sample(s) for s in cli_col["samples"]}
            py_samples = {_normalize_sample(s) for s in py_col["samples"]}
            assert cli_samples == py_samples, (
                f"column {cli_col['name']!r} in sheet "
                f"{cli_sheet['name']!r}: samples differ "
                f"(cli={cli_col['samples']!r}, py={py_col['samples']!r})"
            )


def _normalize_sample(sample: str) -> str:
    """Collapse int/float rendering difference (420000 vs 420000.0).

    Parity today is blocked by the cdylib widening `Data::Int(N)` to
    Python float — so CLI emits ``"420000"`` and Python emits
    ``"420000.0"`` for the same cell. Once Option-A lands and Python
    reads through `wolfxl_core::Workbook`, we can remove this shim and
    compare raw strings.
    """
    try:
        f = float(sample)
    except (ValueError, TypeError):
        return sample
    if f.is_integer():
        return str(int(f))
    return str(f)


def test_classify_format_direct() -> None:
    """The thin `classify_format` wrapper round-trips every category."""
    # Representative format strings for each FormatCategory variant the
    # Rust classifier emits. Not exhaustive (classify_format has its own
    # unit tests in wolfxl-core) — the point here is proving the bridge
    # doesn't drop or mangle any variant.
    cases = {
        "$#,##0.00": "currency",
        "0.00%": "percentage",
        "0.00E+00": "scientific",
        "yyyy-mm-dd": "date",
        "hh:mm:ss": "time",
        "yyyy-mm-dd hh:mm:ss": "datetime",
        "#,##0": "integer",
        "0.00": "float",
        "@": "text",
        "General": "general",
        "": "general",
    }
    for fmt, expected in cases.items():
        got = wolfxl.classify_format(fmt)
        assert got == expected, (
            f"classify_format({fmt!r}): expected {expected!r}, got {got!r}"
        )


def test_worksheet_classify_format_proxies() -> None:
    """`ws.classify_format(fmt)` must equal module-level `classify_format`."""
    wb = wolfxl.Workbook()
    ws = wb.active
    for fmt in ("$#,##0", "0.00%", "yyyy-mm-dd", "General"):
        assert ws.classify_format(fmt) == wolfxl.classify_format(fmt)


def test_classify_sheet_direct_bridge() -> None:
    """The Python-visible bridge exposes direct sheet classification."""
    from wolfxl._rust import classify_sheet

    assert classify_sheet([], "Empty") == "empty"
    assert classify_sheet([["account"], ["cash"]], "Readme") == "readme"


def test_infer_sheet_schema_pads_ragged_rows() -> None:
    """Ragged Python rows keep later wider cells visible to schema inference."""
    from wolfxl._rust import infer_sheet_schema

    schema = infer_sheet_schema([["name"], ["Alice", "extra"]], "Ragged")
    assert schema["rows"] == 1
    assert len(schema["columns"]) == 2
    assert schema["columns"][0]["name"] == "name"
    assert schema["columns"][1]["name"] == ""
    assert schema["columns"][1]["samples"] == ["extra"]


def test_infer_sheet_schema_preserves_oversized_python_int_text() -> None:
    """Huge Python ints should not be rounded through a lossy f64 fallback."""
    from wolfxl._rust import infer_sheet_schema

    huge = 2**80 + 12345
    schema = infer_sheet_schema([["value"], [huge]], "Huge")
    col = schema["columns"][0]
    assert col["samples"] == [str(huge)]


def test_worksheet_schema_includes_pending_number_format_edits() -> None:
    """Unsaved cell.number_format edits should influence schema categories."""
    wb = wolfxl.Workbook()
    ws = wb.active
    ws["A1"] = "price"
    ws["A2"] = 12.5
    ws["A2"].number_format = "$#,##0.00"

    schema = ws.schema()
    assert schema["columns"][0]["format"] == "currency"
