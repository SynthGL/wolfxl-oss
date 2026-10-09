"""Paired, fresh-process evidence for the Python-only sparse blank-cell fix.

Build/install the ordinary wheel first. Supply the baseline _streaming.py from
the pinned base commit; all other Python files and the native binary stay fixed.
This is a labelled overlay comparison, not a second native source build.
"""
from __future__ import annotations

import argparse
import cProfile
import hashlib
import inspect
import json
import os
import platform
import pstats
import shutil
import statistics
import subprocess
import sys
import tempfile
from pathlib import Path
from time import perf_counter

from performance_fixtures import prepare_fixture, sha256


def operation(module, path):
    # Keep the public style loop identical to performance_contract.styled_signature.
    started = perf_counter()
    wb = module.load_workbook(str(path), read_only=True, data_only=True)
    loaded = perf_counter()
    try:
        rows, checksum, styled_cells = 0, 0.0, 0
        for row in wb.active.iter_rows():
            rows += 1
            first = row[0].value if row else None
            if type(first) in (int, float):
                checksum += float(first)
            for cell in row:
                fill = getattr(cell.fill, "fill_type", None) or getattr(cell.fill, "patternType", None)
                if getattr(cell.font, "bold", False) or fill or (cell.number_format or "General") != "General":
                    styled_cells += 1
    finally:
        wb.close()
    ended = perf_counter()
    return {"total_seconds": ended - started, "iterate_and_close_seconds": ended - loaded,
            "signature": {"rows": rows, "first_column_checksum": checksum, "styled_cells": styled_cells}}


def worker(args):
    module = __import__(args.engine)
    if args.profile:
        profiler = cProfile.Profile()
        sample = profiler.runcall(operation, module, args.fixture)
        profiler.dump_stats(args.profile)
        stats = pstats.Stats(profiler)
        sample["profile_top_self_time"] = [
            {"file": Path(key[0]).name, "line": key[1], "function": key[2],
             "calls": value[1], "self_seconds": value[2], "cumulative_seconds": value[3]}
            for key, value in sorted(stats.stats.items(), key=lambda item: item[1][2], reverse=True)[:25]
        ]
        if args.engine == "wolfxl":
            import wolfxl._streaming as streaming
            counts = {}
            for name in ("StreamingCell", "StreamingBlankCell"):
                cls = getattr(streaming, name, None)
                if cls is not None:
                    method = cls.__init__
                    key = (inspect.getfile(method), inspect.getsourcelines(method)[1], "__init__")
                    counts[name] = stats.stats.get(key, (0, 0))[1]
            sample["proxy_constructions"] = counts
    else:
        sample = operation(module, args.fixture)
    print(json.dumps(sample))


def fingerprint(directory):
    digest = hashlib.sha256()
    files = sorted(directory.rglob("*.py"))
    for path in files:
        digest.update(str(path.relative_to(directory)).encode())
        digest.update(path.read_bytes())
    return {"sha256": digest.hexdigest(), "files": len(files)}


def compare(args):
    import wolfxl
    import wolfxl._rust as native
    import openpyxl

    package = Path(wolfxl.__file__).resolve().parent
    report = {"schema": "wolfxl-sparse-python-overlay-v1", "baseline_ref": args.baseline_ref,
              "candidate_ref": args.candidate_ref, "python": platform.python_version(),
              "platform": platform.platform(), "openpyxl": openpyxl.__version__,
              "native_origin": args.native_origin, "native_sha256": sha256(Path(native.__file__)),
              "candidate_streaming_sha256": sha256(package / "_streaming.py"),
              "baseline_streaming_sha256": sha256(args.baseline_streaming),
              "warmups_per_engine_per_case": 1, "fresh_process_trials": args.trials,
              "timing_statistic": "median load + identical public scan + close",
              "workloads": {}}
    script = Path(__file__).resolve()
    with tempfile.TemporaryDirectory(prefix="wolfxl-sparse-pair-") as temporary:
        root = Path(temporary)
        overlays = {}
        for label in ("baseline", "candidate"):
            directory = root / label
            shutil.copytree(package, directory / "wolfxl", ignore=shutil.ignore_patterns("__pycache__"))
            if label == "baseline":
                shutil.copyfile(args.baseline_streaming, directory / "wolfxl" / "_streaming.py")
            overlays[label] = directory
            report[label + "_python"] = fingerprint(directory / "wolfxl")

        def run(label, fixture, profile=None):
            env = dict(os.environ)
            # Exclude inherited source overlays. Each worker imports only its
            # own copy of the installed wheel (or the pinned openpyxl package).
            env["PYTHONPATH"] = str(overlays[label]) if label in overlays else ""
            command = [sys.executable, str(script), "--worker", "--engine",
                       "openpyxl" if label == "openpyxl" else "wolfxl", "--fixture", str(fixture)]
            if profile:
                command += ["--profile", str(profile)]
            return json.loads(subprocess.check_output(command, env=env, text=True))

        for label, kind, rows in (("sparse", "styled_sparse", 25000), ("dense", "styled", 20000)):
            fixture, manifest = prepare_fixture(args.fixture_dir, rows=rows, kind=kind)
            expected = run("openpyxl", fixture)["signature"]
            samples = {engine: [] for engine in ("baseline", "candidate", "openpyxl")}
            for engine in samples:
                assert run(engine, fixture)["signature"] == expected
            for _ in range(args.trials):
                for engine in samples:
                    sample = run(engine, fixture)
                    assert sample["signature"] == expected, (engine, sample, expected)
                    samples[engine].append(sample)
            medians = {engine: statistics.median(s["total_seconds"] for s in values)
                       for engine, values in samples.items()}
            case = {"fixture": manifest, "expected_signature": expected, "samples": samples,
                    "median_seconds": medians, "candidate_speedup": medians["baseline"] / medians["candidate"],
                    "baseline_over_openpyxl": medians["baseline"] / medians["openpyxl"],
                    "candidate_over_openpyxl": medians["candidate"] / medians["openpyxl"]}
            if label == "sparse":
                args.output.parent.mkdir(parents=True, exist_ok=True)
                case["profiles"] = {engine: run(engine, fixture, args.output.with_name(engine + ".prof"))
                                    for engine in ("baseline", "candidate")}
            report["workloads"][label] = case
            print(label, json.dumps({"median_seconds": medians, "candidate_speedup": case["candidate_speedup"]}), flush=True)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(report, indent=2, sort_keys=True) + "\n")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--worker", action="store_true")
    parser.add_argument("--engine", choices=("wolfxl", "openpyxl"), default="wolfxl")
    parser.add_argument("--fixture", type=Path)
    parser.add_argument("--profile", type=Path)
    parser.add_argument("--baseline-streaming", type=Path)
    parser.add_argument("--baseline-ref", default="unspecified")
    parser.add_argument("--candidate-ref", default="unspecified")
    parser.add_argument("--native-origin", default="unspecified; do not treat as source-matched")
    parser.add_argument("--fixture-dir", type=Path, default=Path("/tmp/wolfxl-sparse-fixtures"))
    parser.add_argument("--output", type=Path, default=Path("/tmp/wolfxl-sparse-evidence/results.json"))
    parser.add_argument("--trials", type=int, default=5)
    args = parser.parse_args()
    if args.worker:
        parser.error("--worker requires --fixture") if args.fixture is None else worker(args)
    else:
        if args.baseline_streaming is None or args.trials < 1:
            parser.error("comparison requires --baseline-streaming and positive --trials")
        compare(args)


if __name__ == "__main__":
    main()
