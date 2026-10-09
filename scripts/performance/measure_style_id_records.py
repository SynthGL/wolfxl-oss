"""Measure an existing record API separately from the same-result Cell scan.

This measures a different API contract, not a cumulative engine optimization.
Both scans verify the styled-fixture rows, first-column checksum and styled-cell
count. Record windows bound Python output; the native sheet remains eager.
"""

from __future__ import annotations

import argparse
import gc
import hashlib
import json
from pathlib import Path
import platform
import statistics
import sys
import time

import wolfxl


def cells_scan(path: Path) -> tuple[int, float, int]:
    workbook = wolfxl.load_workbook(path, data_only=True)
    rows = 0
    checksum = 0.0
    styled = 0
    try:
        for row in workbook.active.iter_rows():
            rows += 1
            if row and isinstance(row[0].value, (int, float)):
                checksum += float(row[0].value)
            for cell in row:
                fill = getattr(cell.fill, "fill_type", None) or getattr(cell.fill, "patternType", None)
                styled += bool(cell.font.bold or fill or cell.number_format != "General")
    finally:
        workbook.close()
    return rows, checksum, styled


def records_scan(path: Path) -> tuple[int, float, int]:
    workbook = wolfxl.load_workbook(path, data_only=True)
    checksum = 0.0
    styled = 0
    try:
        worksheet = workbook.active
        rows = worksheet.max_row
        columns = max(1, worksheet.max_column)
        rows_per_batch = max(1, 50_000 // columns)
        styles = {0: False}
        for first in range(1, rows + 1, rows_per_batch):
            for record in worksheet.iter_cell_records(
                min_row=first, max_row=min(rows, first + rows_per_batch - 1),
                min_col=1, max_col=columns, include_format=True,
                include_extended_format=False, include_coordinate=False,
                include_style_id=True, include_empty=False,
            ):
                value = record["value"]
                if record["column"] == 1 and isinstance(value, (int, float)):
                    checksum += float(value)
                style_id = int(record.get("style_id", 0))
                if style_id not in styles:
                    # Diagnostic code resolves each raw ID through the native
                    # reader once; this private hook is not a new public API.
                    payload = workbook._rust_reader.read_format_for_style_id(style_id)
                    styles[style_id] = bool(payload.get("bold") or payload.get("bg_color")
                                            or payload.get("gradient")
                                            or (payload.get("number_format") or "General") != "General")
                styled += styles[style_id]
        assert not worksheet._cells
    finally:
        workbook.close()
    return rows, checksum, styled


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--fixture", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--samples", type=int, default=7)
    args = parser.parse_args()
    expected = cells_scan(args.fixture)
    samples = {"cell_api": [], "existing_record_api": []}
    readers = {"cell_api": cells_scan, "existing_record_api": records_scan}
    for index in range(args.samples + 1):
        for name in (("cell_api", "existing_record_api") if index % 2 == 0 else ("existing_record_api", "cell_api")):
            gc.collect()
            gc.disable()
            try:
                start = time.perf_counter()
                actual = readers[name](args.fixture)
                elapsed = time.perf_counter() - start
            finally:
                gc.enable()
            assert actual == expected, (name, actual, expected)
            if index:
                samples[name].append(elapsed)
    medians = {name: statistics.median(values) for name, values in samples.items()}
    result = dict(python=sys.version, platform=platform.platform(),
                  fixture_sha256=hashlib.sha256(args.fixture.read_bytes()).hexdigest(),
                  result=expected, measurement=dict(warmups=1, samples=args.samples,
                                                    alternating_order=True, gc_disabled_during_timing=True),
                  samples=samples, medians=medians,
                  api_choice_ratio=medians["cell_api"] / medians["existing_record_api"],
                  notes=["Different API; not included in cumulative engine speedups.",
                         "Native worksheet remains eager; Python record windows are bounded.",
                         "Raw style IDs are source-derived and resolved through a private native hook for this diagnostic."])
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(result, indent=2) + "\n")
    print(json.dumps(dict(medians=medians, api_choice_ratio=result["api_choice_ratio"]), indent=2))


if __name__ == "__main__":
    main()
