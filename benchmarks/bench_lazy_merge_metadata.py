"""Before/after lazy-merge probe; fixture generation is excluded from timing."""

from pathlib import Path
import argparse
import gc
import hashlib
import json
import platform
import resource
import statistics
import sys
import time
import zipfile
from collections import Counter
from unittest.mock import patch
import wolfxl

parser = argparse.ArgumentParser()
parser.add_argument("--output", required=True)
parser.add_argument("--source-ref")
parser.add_argument("--native-source-ref")
parser.add_argument("--iterations", type=int, default=7)
parser.add_argument("--fixture-dir", type=Path, required=True)
parser.add_argument("--prepare", action="store_true")
args = parser.parse_args()
base = args.fixture_dir
if args.prepare:
    from benchmark_openpyxl_vs_wolfxl import build_openpyxl_styled

    base.mkdir(parents=True, exist_ok=True)
    unmerged = base / "styled-20000.xlsx"
    merged = base / "styled-20000-one-merge.xlsx"
    if not unmerged.exists():
        build_openpyxl_styled(unmerged, 20000, 5)
    if not merged.exists():
        with (
            zipfile.ZipFile(unmerged) as source,
            zipfile.ZipFile(merged, "w", zipfile.ZIP_DEFLATED) as output,
        ):
            for info in source.infolist():
                data = source.read(info)
                if info.filename == "xl/worksheets/sheet1.xml":
                    data = data.replace(
                        b"</sheetData>",
                        b'</sheetData><mergeCells count="1"><mergeCell ref="D19999:E20000"/></mergeCells>',
                    )
                output.writestr(info, data)
fixtures = {
    "unmerged_100000_cells": base / "styled-20000.xlsx",
    "one_merge_100000_cells": base / "styled-20000-one-merge.xlsx",
}
results = {
    "source_ref": args.source_ref,
    "native_source_ref": args.native_source_ref or args.source_ref,
    "python": sys.version,
    "platform": platform.platform(),
    "wolfxl": wolfxl.__version__,
    "native_path": wolfxl._rust.__file__,
    "native_sha256": hashlib.sha256(Path(wolfxl._rust.__file__).read_bytes()).hexdigest(),
    "python_source_sha256": {
        name: hashlib.sha256((Path(wolfxl.__file__).parent / name).read_bytes()).hexdigest()
        for name in ("_workbook_state.py", "_worksheet_collections.py", "_cell.py")
    },
    "iterations": args.iterations,
    "workloads": {},
    "untimed_diagnostics": {},
    "fixture_sha256": {k: hashlib.sha256(v.read_bytes()).hexdigest() for k, v in fixtures.items()},
}


def consume(wb, workload, merged):
    if workload == "load_and_merge_ranges":
        ranges = wb._rust_reader.read_merged_ranges(wb.active.title)
        assert len(ranges) == (1 if merged else 0)
    elif workload == "load_and_first_border":
        # Select a late-sheet endpoint: a sparse scan must not decode all cells.
        border = wb.active["D19999"].border
        assert border is not None


class ReaderCounter:
    def __init__(self, native):
        self.native = native
        self.calls = Counter()

    def __getattr__(self, name):
        value = getattr(self.native, name)
        if not callable(value):
            return value

        def counted(*args, **kwargs):
            self.calls[name] += 1
            return value(*args, **kwargs)

        return counted


for name, path in fixtures.items():
    for workload in ("load_only", "load_and_merge_ranges", "load_and_first_border"):
        samples = []
        # One warm-up, then matched full operation from load until result.
        for i in range(args.iterations + 1):
            gc.collect()
            start = time.perf_counter()
            wb = wolfxl.load_workbook(path)
            consume(wb, workload, name.startswith("one_merge"))
            elapsed = time.perf_counter() - start
            wb.close()
            if i:
                samples.append(elapsed)
        results["workloads"][f"{name}/{workload}"] = {
            "samples_seconds": samples,
            "median_seconds": statistics.median(samples),
            "min_seconds": min(samples),
            "max_seconds": max(samples),
        }
        # Keep counters out of the timed loop. Python ZIP opens directly prove
        # elimination of eager worksheet parsing, independently of wall time.
        worksheet_opens = []
        original_open = zipfile.ZipFile.open

        def tracked_open(archive, entry, *positional, **kwargs):
            part = entry.filename if isinstance(entry, zipfile.ZipInfo) else str(entry)
            if part.startswith("xl/worksheets/"):
                worksheet_opens.append(part)
            return original_open(archive, entry, *positional, **kwargs)

        with patch.object(zipfile.ZipFile, "open", tracked_open):
            wb = wolfxl.load_workbook(path)
            native = ReaderCounter(wb._rust_reader)
            wb._rust_reader = native
            consume(wb, workload, name.startswith("one_merge"))
            results["untimed_diagnostics"][f"{name}/{workload}"] = {
                "python_worksheet_xml_opens": worksheet_opens,
                "native_public_calls_after_load": dict(native.calls),
                "python_cells_materialized": len(wb.active._cells),
                "python_border_table_entries": len(wb._borders),
            }
            wb.close()
            del native, wb
results["maxrss_kib"] = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
Path(args.output).write_text(json.dumps(results, indent=2) + "\n")
print(json.dumps({k: v["median_seconds"] for k, v in results["workloads"].items()}, indent=2))
