"""Measure bounded XLSX reads and emit deterministic per-sheet value hashes.

Run with the interpreter containing the build to measure. Inputs must already be
in --corpus-dir. Use --corpus for all .xlsx files or omit it for the three POI
cases and two generated ZIP cases from tests/test_load_limits.py. Each case runs
in a fresh interpreter, with a five-second timeout and a 2 GiB address-space cap.
String values are represented by their length and SHA-256; other values retain
their type and content. The JSON includes every normalized row and load error.
Linux only: /usr/bin/time records peak RSS even when the child aborts.
"""
from __future__ import annotations

import argparse
import datetime
import hashlib
import importlib
import json
import os
from pathlib import Path
import platform
import resource
import runpy
import signal
import subprocess
import sys
import time

CASES = (
    "poc-shared-strings.xlsx", "poc-xmlbomb.xlsx", "poc-xmlbomb-empty.xlsx",
    "generated-zip-bomb.xlsx", "generated-many-members.xlsx",
)


def normalized(value, strings):
    if isinstance(value, str):
        key = id(value)
        if key not in strings:
            digest = hashlib.sha256(value.encode("utf-8", errors="surrogatepass")).hexdigest()
            strings[key] = (value, ["str", len(value), digest])
        return strings[key][1]
    if value is None or isinstance(value, (bool, int)):
        return value
    if isinstance(value, float):
        return ["float", repr(value)]
    if isinstance(value, (datetime.datetime, datetime.date, datetime.time)):
        return [type(value).__name__, value.isoformat()]
    if isinstance(value, datetime.timedelta):
        return ["timedelta", value.total_seconds()]
    return [type(value).__name__, str(value)]


def child(args):
    resource.setrlimit(resource.RLIMIT_AS, (2 * 1024**3, 2 * 1024**3))
    start = time.perf_counter()
    sheets, strings = [], {}
    result = {"file": args.file, "read_only": args.read_only, "status": "ok", "error": None}
    try:
        library = importlib.import_module(args.library)
        workbook = library.load_workbook(args.corpus_dir / args.file, read_only=args.read_only)
        for sheet in workbook.worksheets:
            digest = hashlib.sha256()
            rows, nonempty = [], 0
            for row in sheet.iter_rows(values_only=True):
                values = [normalized(value, strings) for value in row]
                rows.append(values)
                digest.update(json.dumps(values, sort_keys=True, separators=(",", ":"), ensure_ascii=True).encode())
                digest.update(b"\n")
                nonempty += sum(value is not None for value in row)
            sheets.append({"name": sheet.title, "rows": rows, "nonempty": nonempty, "values_sha256": digest.hexdigest()})
        workbook.close()
    except BaseException as error:
        result.update(status="error", error={"type": type(error).__name__, "message": str(error).replace(str(args.corpus_dir.resolve().parent), "$SCRATCH")})
    result.update(sheets=sheets, nonempty=sum(sheet["nonempty"] for sheet in sheets), load_scan_s=time.perf_counter() - start)
    print(json.dumps(result, sort_keys=True))


def measure(args):
    library = importlib.import_module(args.library)
    binary = Path(importlib.import_module("wolfxl._rust").__file__) if args.library == "wolfxl" else Path(library.__file__)
    identity = {"sha256": hashlib.sha256(binary.read_bytes()).hexdigest(), "mtime_ns": binary.stat().st_mtime_ns, "bytes": binary.stat().st_size, "version": library.__version__, "file": binary.name}
    files = sorted(path.name for path in args.corpus_dir.glob("*.xlsx") if not path.name.startswith("generated-")) if args.corpus else CASES
    args.output.parent.mkdir(parents=True, exist_ok=True)
    results = []
    timing = args.output.with_suffix(".timing.json")
    for filename in files:
        if not (args.corpus_dir / filename).is_file():
            raise FileNotFoundError(filename)
        for read_only in ([False] if args.corpus else [False, True]):
            command = [sys.executable, str(Path(__file__).resolve()), "--child", "--library", args.library, "--file", filename, "--corpus-dir", str(args.corpus_dir)]
            if read_only:
                command.append("--read-only")
            command = ["/usr/bin/time", "-f", '{"peak_rss_kib":%M,"wall_s":%e}', "-o", str(timing), *command]
            process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, start_new_session=True)
            start = time.perf_counter()
            try:
                stdout, stderr = process.communicate(timeout=5)
                try:
                    result = json.loads(stdout)
                except ValueError:
                    result = {"file": filename, "read_only": read_only, "status": "crash", "returncode": process.returncode, "stderr": stderr[-2000:]}
            except subprocess.TimeoutExpired:
                os.killpg(process.pid, signal.SIGKILL)
                process.communicate()
                result = {"file": filename, "read_only": read_only, "status": "timeout", "wall_s": time.perf_counter() - start}
            if timing.exists():
                for line in timing.read_text().splitlines():
                    if line.startswith("{"):
                        measured = json.loads(line)
                        result.update(peak_rss_bytes=measured["peak_rss_kib"] * 1024, wall_s=measured["wall_s"])
                timing.unlink()
            result["fixture_sha256"] = hashlib.sha256((args.corpus_dir / filename).read_bytes()).hexdigest()
            results.append(result)
    if args.source == "HEAD":
        args.source = subprocess.check_output(
            ["git", "rev-parse", "HEAD"], cwd=Path(__file__).resolve().parent.parent, text=True
        ).strip()
    metadata = {
        "source": args.source,
        "library": args.library,
        "python": platform.python_version(),
        "platform": platform.platform(),
        "architecture": platform.machine(),
        "zip_environment": {
            key: os.environ.get(key)
            for key in (
                "WOLFXL_MAX_ZIP_ENTRIES", "WOLFXL_MAX_ZIP_ENTRY_BYTES",
                "WOLFXL_MAX_ZIP_TOTAL_BYTES", "WOLFXL_MAX_ZIP_COMPRESSION_RATIO",
            )
        },
        "iterations": 1,
        "warmups": 0,
        "cache_state": "fresh interpreter; OS cache unspecified",
        "build_mode": "release" if args.library == "wolfxl" else "pure Python",
        "timeout_s": 5,
        "memory_address_space_limit_bytes": 2 * 1024**3,
        "binary": identity,
        "results": results,
    }
    args.output.write_text(json.dumps(metadata, indent=2))
    print(args.output.name, len(results))


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--corpus-dir", type=Path, required=True)
    parser.add_argument("--generate-cases", action="store_true")
    parser.add_argument("--output", type=Path)
    parser.add_argument("--source")
    parser.add_argument("--library", default="wolfxl")
    parser.add_argument("--corpus", action="store_true")
    parser.add_argument("--child", action="store_true")
    parser.add_argument("--file")
    parser.add_argument("--read-only", action="store_true")
    args = parser.parse_args()
    if args.generate_cases:
        args.corpus_dir.mkdir(parents=True, exist_ok=True)
        helper = Path(__file__).resolve().parent.parent / "tests" / "test_load_limits.py"
        generate = runpy.run_path(str(helper))["generate_zip_bomb"]
        generate(args.corpus_dir / "generated-zip-bomb.xlsx")
        generate(args.corpus_dir / "generated-many-members.xlsx", members=200_001)
        return
    if sys.platform != "linux":
        parser.error("requires Linux /usr/bin/time")
    if args.child:
        child(args)
    elif args.output is None or args.source is None:
        parser.error("--output and --source are required")
    else:
        measure(args)


if __name__ == "__main__":
    main()
