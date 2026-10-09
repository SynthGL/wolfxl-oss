"""Compare complete source-built wheels using the frozen styled-read loop.

Extract each wheel into a separate directory, then pass those directories.
The existing sparse overlay receipts and producer are left unchanged.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import platform
import random
import statistics
import subprocess
import sys
import time

from performance_fixtures import prepare_fixture, sha256


def fingerprint(package):
    digest = hashlib.sha256()
    for path in sorted(package.rglob("*.py")):
        digest.update(str(path.relative_to(package)).encode())
        digest.update(path.read_bytes())
    return digest.hexdigest()


def worker(args):
    module = __import__(args.engine)
    started = time.perf_counter()
    wb = module.load_workbook(args.fixture, read_only=True, data_only=True)
    loaded = time.perf_counter()
    rows, checksum, styled_cells, positions = 0, 0.0, 0, 0
    try:
        for row in wb.active.iter_rows():
            rows += 1
            if args.phase == "rows":
                positions += len(row)
                continue
            first = row[0].value if row else None
            if type(first) in (int, float):
                checksum += float(first)
            for cell in row:
                fill = getattr(cell.fill, "fill_type", None) or getattr(cell.fill, "patternType", None)
                if getattr(cell.font, "bold", False) or fill or (cell.number_format or "General") != "General":
                    styled_cells += 1
    finally:
        wb.close()
    ended = time.perf_counter()
    signature = ({"rows": rows, "positions": positions} if args.phase == "rows" else
                 {"rows": rows, "first_column_checksum": checksum, "styled_cells": styled_cells})
    result = {"total_seconds": ended - started, "load_seconds": loaded - started,
              "iterate_and_close_seconds": ended - loaded, "signature": signature}
    if args.engine == "wolfxl":
        import wolfxl._rust as native
        result["native_batch_available"] = hasattr(native, "streaming_blank_rows")
    print(json.dumps(result))


def compare(args):
    import openpyxl

    roots = {name: Path(getattr(args, name + "_package")).resolve()
             for name in ("baseline", "candidate")}
    for name, root in roots.items():
        if not (root / "wolfxl" / "__init__.py").is_file():
            raise ValueError(f"{name} must contain an extracted wolfxl package")
    report = {"schema": "wolfxl-native-blank-rows-v1", "python": sys.version,
              "platform": platform.platform(), "openpyxl": openpyxl.__version__,
              "baseline_ref": args.baseline_ref, "candidate_ref": args.candidate_ref,
              "build_description": args.build_description,
              "fresh_process_trials": args.trials, "warmups_per_engine_per_case": 1,
              "timing_statistic": "median load + public scan + close",
              "packages": {}, "workloads": {}}
    for name, root in roots.items():
        package = root / "wolfxl"
        extensions = list(package.glob("_rust*.so")) + list(package.glob("_rust*.pyd"))
        if len(extensions) != 1:
            raise ValueError(f"Expected one native extension in {package}")
        report["packages"][name] = {
            "python_sha256": fingerprint(package), "native_sha256": sha256(extensions[0])}
    output = Path(args.output).resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    rng = random.Random(20261009)
    script = str(Path(__file__).resolve())

    def run(name, fixture, phase):
        env = os.environ.copy()
        env.pop("PYTHONPATH", None)
        if name != "openpyxl":
            env["PYTHONPATH"] = str(roots[name])
        command = [sys.executable, script, "--worker", "--engine",
                   "openpyxl" if name == "openpyxl" else "wolfxl",
                   "--fixture", str(fixture), "--phase", phase]
        result = subprocess.run(command, env=env, text=True, capture_output=True, check=True)
        sample = json.loads(result.stdout)
        if name == "candidate" and not sample.get("native_batch_available"):
            raise ValueError("Candidate wheel lacks the native blank-row implementation")
        return sample

    for name, kind, rows in [("sparse", "styled_sparse", 25000), ("dense", "styled", 20000)]:
        fixture, metadata = prepare_fixture(output.parent / "fixtures", kind=kind, rows=rows)
        for phase in ("styles", "rows"):
            labels = ["baseline", "candidate", "openpyxl"]
            rng.shuffle(labels)
            for label in labels:
                run(label, fixture, phase)
            samples = {label: [] for label in labels}
            order = []
            for _ in range(args.trials):
                rng.shuffle(labels)
                for label in labels:
                    samples[label].append(run(label, fixture, phase))
                    order.append(label)
            expected = samples["openpyxl"][0]["signature"]
            for label, values in samples.items():
                if any(sample["signature"] != expected for sample in values):
                    raise ValueError(f"Signature mismatch: {name}/{phase}/{label}")
            medians = {label: statistics.median(s["total_seconds"] for s in values)
                       for label, values in samples.items()}
            report["workloads"][name + "_" + phase] = {
                "fixture": metadata, "signature": expected, "samples": samples,
                "trial_order": order, "median_seconds": medians,
                "baseline_over_candidate": medians["baseline"] / medians["candidate"],
                "openpyxl_over_candidate": medians["openpyxl"] / medians["candidate"]}
            print(name, phase, medians, flush=True)
    output.write_text(json.dumps(report, indent=2, sort_keys=True) + "\n")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--worker", action="store_true")
    parser.add_argument("--engine", choices=("wolfxl", "openpyxl"))
    parser.add_argument("--fixture")
    parser.add_argument("--phase", choices=("styles", "rows"), default="styles")
    parser.add_argument("--baseline-package")
    parser.add_argument("--candidate-package")
    parser.add_argument("--baseline-ref")
    parser.add_argument("--candidate-ref")
    parser.add_argument("--build-description")
    parser.add_argument("--trials", type=int, default=5)
    parser.add_argument("--output")
    args = parser.parse_args()
    if args.worker:
        worker(args)
    else:
        if args.trials < 1:
            parser.error("--trials must be positive")
        for key in ("baseline_package", "candidate_package", "baseline_ref", "candidate_ref",
                    "build_description", "output"):
            if not getattr(args, key):
                parser.error(f"--{key.replace('_', '-')} is required")
        compare(args)


if __name__ == "__main__":
    main()
