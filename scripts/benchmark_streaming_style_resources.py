"""Fresh-process native lookup counts and peak RSS for styled streaming.

Use the same fixture for both source-matched environments. Wall times here are
instrumented diagnostics; use benchmarks/performance_contract.py for timings.

python scripts/benchmark_streaming_style_resources.py --fixture workbook.xlsx \
    --source-sha <commit> --output /tmp/styled-resources.json
"""

from __future__ import annotations

import argparse
from collections import Counter
import hashlib
import json
from pathlib import Path
import platform
import subprocess
import sys
from time import perf_counter


class NativeLookupCounter:
    def __init__(self, inner: object) -> None:
        self.inner = inner
        self.calls: Counter[str] = Counter()

    def __getattr__(self, name: str) -> object:
        value = getattr(self.inner, name)
        if name not in {
            "read_format_for_style_id", "read_cell_format", "read_cell_border",
            "read_merged_ranges", "read_merged_cell_border", "read_merged_endpoint_style_ids",
        }:
            return value

        def counted(*args: object, **kwargs: object) -> object:
            self.calls[name] += 1
            return value(*args, **kwargs)

        return counted


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def measure(fixture: Path) -> dict:
    import resource
    import wolfxl
    from wolfxl import _rust

    started = perf_counter()
    wb = wolfxl.load_workbook(str(fixture), read_only=True, data_only=True)
    counter = NativeLookupCounter(wb._rust_reader)
    wb._rust_reader = counter
    rows, checksum, styled_cells = 0, 0.0, 0
    try:
        # Match the v2 public Cell/font/fill/number-format scan exactly.
        for row in wb.active.iter_rows():
            rows += 1
            first = row[0].value if row else None
            if type(first) in (int, float):
                checksum += float(first)
            for cell in row:
                fill = getattr(cell.fill, "fill_type", None) or getattr(cell.fill, "patternType", None)
                if getattr(cell.font, "bold", False) or fill or (cell.number_format or "General") != "General":
                    styled_cells += 1
        cache = getattr(wb, "_streaming_style_cache", None)
        cached_styles = len(cache.styles) if cache is not None else None
    finally:
        wb.close()
    peak = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    return {
        "signature": {"rows": rows, "first_column_checksum": checksum, "styled_cells": styled_cells},
        "native_lookup_counts": dict(counter.calls),
        "cached_style_entries": cached_styles,
        "peak_rss_bytes": peak if sys.platform == "darwin" else peak * 1024,
        "instrumented_operation_seconds": perf_counter() - started,
        "package_path": wolfxl.__file__,
        "native_path": _rust.__file__,
        "native_sha256": sha256(Path(_rust.__file__)),
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--fixture", required=True, type=Path)
    parser.add_argument("--source-sha", default="unrecorded")
    parser.add_argument("--repeats", type=int, default=3)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--worker", action="store_true", help=argparse.SUPPRESS)
    args = parser.parse_args()
    if args.repeats < 1:
        parser.error("--repeats must be >= 1")
    if args.worker:
        print(json.dumps(measure(args.fixture)))
        return
    command = [sys.executable, str(Path(__file__).resolve()), "--worker", "--fixture", str(args.fixture)]
    samples = [json.loads(subprocess.check_output(command, text=True)) for _ in range(args.repeats)]
    if any(sample["signature"] != samples[0]["signature"] for sample in samples[1:]):
        raise AssertionError("Fresh-process styled signatures differ")
    receipt = {
        "contract": "wolfxl-styled-edit-v2-resource-diagnostic",
        "source_sha": args.source_sha,
        "fixture_sha256": sha256(args.fixture),
        "producer_sha256": sha256(Path(__file__)),
        "python": platform.python_version(),
        "platform": platform.platform(),
        "samples": samples,
        "timing_note": "Instrumented resource diagnostic; use performance_contract.py for speedup claims.",
        "memory_note": "Absolute peak RSS measured independently in fresh subprocesses; includes interpreter/imports.",
    }
    encoded = json.dumps(receipt, indent=2, sort_keys=True) + "\n"
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(encoded)
    else:
        print(encoded, end="")


if __name__ == "__main__":
    main()
