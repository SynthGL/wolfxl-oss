"""Compare identical eager Cell scans in alternating fresh processes.

Use an installed source-matched baseline and point --optimized-pythonpath at the
modified source's python directory. The fixture remains unchanged; all samples
verify the same rows, checksum, and styled-cell count against openpyxl.
"""

from __future__ import annotations

import argparse
import gc
import hashlib
import json
import os
from pathlib import Path
import platform
import statistics
import subprocess
import sys
import time


def scan(path: Path, styled: bool, module: object) -> tuple[int, float, int]:
    workbook = module.load_workbook(path, data_only=True)
    rows = 0
    checksum = 0.0
    styled_cells = 0
    try:
        for row in workbook.active.iter_rows():
            rows += 1
            if row and isinstance(row[0].value, (int, float)):
                checksum += float(row[0].value)
            if styled:
                for cell in row:
                    fill_kind = getattr(cell.fill, "fill_type", None) or getattr(cell.fill, "patternType", None)
                    if cell.font.bold or fill_kind or cell.number_format != "General":
                        styled_cells += 1
    finally:
        workbook.close()
    return rows, checksum, styled_cells


def measure(path: Path) -> dict[str, object]:
    import openpyxl
    import wolfxl

    observations = {}
    for name, styled in (("styled_cells", True), ("value_cells", False)):
        expected = scan(path, styled, openpyxl)
        gc.collect()
        gc.disable()
        try:
            start = time.perf_counter()
            actual = scan(path, styled, wolfxl)
            elapsed = time.perf_counter() - start
        finally:
            gc.enable()
        assert actual == expected, (actual, expected)
        observations[name] = dict(seconds=elapsed, result=actual)
    return observations


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--fixture", required=True, type=Path)
    parser.add_argument("--baseline-python", default=sys.executable)
    parser.add_argument("--optimized-pythonpath", type=Path)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--samples", type=int, default=7)
    parser.add_argument("--worker", action="store_true", help=argparse.SUPPRESS)
    args = parser.parse_args()
    if args.worker:
        print(json.dumps(measure(args.fixture)))
        return
    if args.optimized_pythonpath is None or args.output is None:
        parser.error("--optimized-pythonpath and --output are required")
    observations = {"baseline": [], "optimized": []}
    for index in range(args.samples + 1):
        for name in (("baseline", "optimized") if index % 2 == 0 else ("optimized", "baseline")):
            environment = os.environ.copy()
            environment.pop("PYTHONPATH", None)
            if name == "optimized":
                environment["PYTHONPATH"] = str(args.optimized_pythonpath.resolve())
            process = subprocess.run(
                [args.baseline_python, str(Path(__file__).resolve()), "--worker", "--fixture", str(args.fixture.resolve())],
                env=environment, text=True, check=True, capture_output=True,
            )
            if index:
                observations[name].append(json.loads(process.stdout))
    medians = {name: {case: statistics.median(sample[case]["seconds"] for sample in samples)
                      for case in ("styled_cells", "value_cells")}
               for name, samples in observations.items()}
    receipt = dict(python=sys.version, platform=platform.platform(),
                   fixture_sha256=hashlib.sha256(args.fixture.read_bytes()).hexdigest(),
                   measurement=dict(warmups=1, samples=args.samples, alternating_order=True,
                                    fresh_process_per_sample=True, gc_disabled_during_timing=True,
                                    expected_result_from_openpyxl=True),
                   samples=observations, medians=medians,
                   speedup={case: medians["baseline"][case] / medians["optimized"][case]
                            for case in ("styled_cells", "value_cells")})
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(receipt, indent=2) + "\n")
    print(json.dumps(dict(medians=medians, speedup=receipt["speedup"]), indent=2))


if __name__ == "__main__":
    main()
