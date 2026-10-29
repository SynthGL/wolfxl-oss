"""Render the ecosystem comparison charts from a committed results JSON.

Usage:

    python benchmarks/render_ecosystem_charts.py \
        benchmarks/results/2026-08-18-ecosystem-linux-epyc.json \
        assets/benchmarks

Deterministic, stdlib-only. Every number is read from the results JSON
produced by ``benchmarks/benchmark_python_excel_ecosystem.py``, including
row/column counts, so chart titles can never drift from the measured data.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any

from render_charts import (
    FONT,
    MARGIN_X,
    OPENPYXL_FILL,
    TEXT_MUTED,
    TEXT_PRIMARY,
    WIDTH,
    WOLFXL_A,
    esc,
    fmt_seconds,
    footer,
    header,
)

OTHER_FILL = "#64748b"

# engine key -> (display name, version-dict key)
ENGINE_META = {
    "wolfxl": ("WolfXL Community", "wolfxl"),
    "openpyxl": ("openpyxl", "openpyxl"),
    "openpyxl_wo": ("openpyxl, write_only mode", "openpyxl"),
    "xlsxwriter": ("XlsxWriter", "XlsxWriter"),
    "xlsxwriter_cm": ("XlsxWriter, constant_memory", "XlsxWriter"),
    "pyexcelerate": ("PyExcelerate", "PyExcelerate"),
    "pylightxl": ("pylightxl", "pylightxl"),
    "pandas_openpyxl": ("pandas", "pandas"),
    "pandas_xlsxwriter": ("pandas + XlsxWriter", "pandas"),
    "pandas_calamine": ("pandas", "pandas"),
    "polars": ("Polars", "polars"),
    "python_calamine": ("python-calamine", "python-calamine"),
    "fastexcel": ("fastexcel", "fastexcel"),
    "duckdb": ("DuckDB", "duckdb"),
    "tablib": ("Tablib", "tablib"),
    "pyexcel": ("pyexcel", "pyexcel"),
    "xlsx2csv": ("xlsx2csv", "xlsx2csv"),
    "grid_baseline": ("input grid only", ""),
}

CASE_TITLES = {
    "write_plain_large": "Write {rows:,} x {cols} plain values",
    "write_mixed": "Write {rows:,} x {cols} mixed types (str/int/float/bool/datetime)",
    "write_unique_strings": "Write {rows:,} x {cols} unique strings (shared-strings stress)",
    "read_values_large": "Read {rows:,} x {cols}, all values",
    "read_formulas_multisheet": "Read {sheets} sheets, {rows:,} x {cols} total, cached formula results",
    "memory_write_large": "Peak memory: write {rows:,} x {cols} plain values",
    "memory_read_large": "Peak memory: read {rows:,} x {cols}, all values",
}

OUTPUTS = {
    "write_plain_large": "ecosystem-write-large.svg",
    "write_mixed": "ecosystem-write-mixed.svg",
    "write_unique_strings": "ecosystem-write-strings.svg",
    "read_values_large": "ecosystem-read-large.svg",
    "read_formulas_multisheet": "ecosystem-read-formulas.svg",
    "memory_write_large": "ecosystem-memory-write.svg",
    "memory_read_large": "ecosystem-memory-read.svg",
}

EXTRA_NOTES = {
    "read_values_large": [
        "fastexcel/Polars/DuckDB return Arrow-backed tables, not Python cell "
        "objects; xlsx2csv emits CSV text; read fixture written by openpyxl.",
    ],
    "read_formulas_multisheet": [
        "Fixture written by XlsxWriter with a cached result for every formula; "
        "five of eight columns are formulas. Every engine reads every sheet.",
    ],
    "memory_write_large": [
        "One fresh process per engine, single run. 'input grid only' is the cost "
        "of the Python data before any Excel library is imported.",
    ],
    "memory_read_large": [
        "One fresh process per engine, single run, peak RSS including the "
        "returned values.",
    ],
}


def engine_label(engine: str, versions: dict[str, str]) -> tuple[str, str]:
    name, version_key = ENGINE_META[engine]
    version = versions.get(version_key, "") if version_key else ""
    return f"{name} {version}".strip(), name


def render_bars(
    title: str,
    subtitle: str,
    notes: list[str],
    rows: list[tuple[str, str, float, str]],
    max_value: float,
) -> str:
    """rows: (label_line, scope_line, value, value_text) sorted ascending."""
    bar_h, gap, top, label_w = 24, 34, 96, 300
    height = top + gap * len(rows) + 76
    chart_w = WIDTH - MARGIN_X * 2 - label_w - 110
    parts = header(title, subtitle, height)
    for idx, (label, scope, value, value_text) in enumerate(rows):
        y = top + idx * gap
        is_wolfxl = label.startswith("WolfXL")
        fill = "url(#wolfxl)" if is_wolfxl else (
            OPENPYXL_FILL if label.startswith("openpyxl") else OTHER_FILL
        )
        label_fill = TEXT_PRIMARY if is_wolfxl else TEXT_MUTED
        weight = "700" if is_wolfxl else "600"
        parts.append(
            f'<text x="{MARGIN_X}" y="{y + 12}" font-family="{FONT}" font-size="13" '
            f'font-weight="{weight}" fill="{label_fill}">{esc(label)}</text>'
        )
        parts.append(
            f'<text x="{MARGIN_X}" y="{y + 26}" font-family="{FONT}" font-size="10.5" '
            f'fill="{TEXT_MUTED}">{esc(scope)}</text>'
        )
        bar_w = max(chart_w * value / max_value, 3)
        parts.append(
            f'<rect x="{MARGIN_X + label_w}" y="{y}" width="{bar_w:.1f}" '
            f'height="{bar_h}" rx="{bar_h / 2}" fill="{fill}"/>'
        )
        value_fill = WOLFXL_A if is_wolfxl else TEXT_MUTED
        parts.append(
            f'<text x="{MARGIN_X + label_w + bar_w + 10:.1f}" y="{y + 17}" '
            f'font-family="{FONT}" font-size="13.5" font-weight="700" '
            f'fill="{value_fill}">{esc(value_text)}</text>'
        )
    return footer(parts, notes, height)


def main() -> None:
    if len(sys.argv) != 3:
        raise SystemExit(__doc__)
    results_path = Path(sys.argv[1])
    output_dir = Path(sys.argv[2])
    output_dir.mkdir(parents=True, exist_ok=True)
    payload = json.loads(results_path.read_text(encoding="utf-8"))
    meta = payload["metadata"]
    versions = meta["versions"]

    os_name = meta["platform"].split("-", 1)[0]
    env = f"{meta['cpu_brand']} | {meta['machine']} {os_name} | Python {meta['python']}"
    timing_subtitle = (
        f"{env} | median of {meta['rounds']} rounds | lower is better"
    )
    memory_subtitle = (
        f"{env} | peak RSS, fresh process per engine | lower is better"
    )
    provenance = [
        f"Raw data: benchmarks/results/{results_path.name} (run {meta['timestamp_utc'][:10]})",
        "Reproduce: python benchmarks/benchmark_python_excel_ecosystem.py "
        "(see benchmarks/README.md)",
    ]

    by_case: dict[str, list[dict[str, Any]]] = {}
    for row in payload["results"]:
        by_case.setdefault(row["case"], []).append(row)
    for row in payload.get("memory", []):
        by_case.setdefault(row["case"], []).append(row)

    written = 0
    for case, filename in OUTPUTS.items():
        case_rows = by_case.get(case)
        if not case_rows:
            continue
        title = CASE_TITLES[case].format(
            rows=case_rows[0]["rows"],
            cols=case_rows[0]["cols"],
            sheets=case_rows[0].get("sheets"),
        )
        finished = [r for r in case_rows if not r.get("timed_out")]
        dnf_notes = [
            f"{engine_label(r['engine'], versions)[0]}: did not finish one round "
            f"within {r['round_budget_seconds']} s; excluded from the chart."
            for r in case_rows
            if r.get("timed_out")
        ]
        is_memory = case.startswith("memory_")
        value_key = "peak_rss_bytes" if is_memory else "median_seconds"
        ranked = sorted(finished, key=lambda r: r[value_key])
        values = [r[value_key] for r in ranked]
        max_value = values[-1]
        clip_notes: list[str] = []
        # A single extreme outlier (e.g. a quadratic engine hundreds of times
        # slower) would flatten every other bar. Clip the axis and say so; the
        # printed value is always the true measurement.
        if len(values) >= 3 and values[-1] > 8 * values[-2]:
            max_value = values[-2] * 1.25
            clipped = [r for r in ranked if r[value_key] > max_value]
            names = ", ".join(
                engine_label(r["engine"], versions)[0] for r in clipped
            )
            clip_notes.append(
                f"Axis clipped: the {names} bar extends far beyond the chart; "
                "the printed value is the true measurement."
            )
        bars = []
        for row in ranked:
            label, _ = engine_label(row["engine"], versions)
            value_text = (
                f"{row['peak_rss_bytes'] / 1048576:,.0f} MiB"
                if is_memory
                else fmt_seconds(row["median_seconds"])
            )
            bars.append(
                (label, row["scope"], min(row[value_key], max_value), value_text)
            )
        subtitle = memory_subtitle if is_memory else timing_subtitle
        notes = EXTRA_NOTES.get(case, []) + clip_notes + dnf_notes + provenance
        chart = render_bars(title, subtitle, notes, bars, max_value)
        (output_dir / filename).write_text(chart, encoding="utf-8")
        written += 1
    print(f"Wrote {written} charts to {output_dir}")


if __name__ == "__main__":
    main()
