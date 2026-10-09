#!/usr/bin/env python3
"""Fresh-process dense-merge point diagnostic; not a workbook speed claim."""
from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import statistics
import subprocess
import sys
import time
from types import SimpleNamespace


def child(case: str, merges: int, queries: int) -> dict:
    import wolfxl._streaming_styles as styles
    from wolfxl.utils.cell import get_column_letter
    import wolfxl._rust as native

    if case == "vertical":
        refs = [f"A{4*i+1}:B{4*i+2}" for i in range(merges)]
        points = [(4*(i % merges) + 1 + i % 3, 1 + i % 3) for i in range(queries)]
    elif case == "horizontal":
        refs = [f"{get_column_letter(3*i+1)}1:{get_column_letter(3*i+2)}2" for i in range(merges)]
        points = [(1 + i % 3, 1 + (i*29) % (3*merges)) for i in range(queries)]
    else:
        refs = []
        points = [(i+1, 1) for i in range(queries)]
    fixture_hash = hashlib.sha256(json.dumps([refs, points], separators=(",", ":")).encode()).hexdigest()
    styles._loaded_merged_range_refs = lambda ws: refs
    worksheet = SimpleNamespace(title="Density", _merged_range_index=None)
    cache = styles.StreamingStyleCache(object())
    start = time.perf_counter()
    cache.merge_bounds(worksheet)
    preparation_seconds = time.perf_counter() - start
    hits = checksum = 0
    start = time.perf_counter()
    for row, col in points:
        bounds = cache.merge_at(worksheet, row, col)
        if bounds is not None:
            hits += 1
            checksum += sum(bounds)
    elapsed = time.perf_counter() - start
    return {
        "seconds": elapsed, "preparation_seconds": preparation_seconds,
        "signature": {"queries": queries, "hits": hits, "checksum": checksum},
        "fixture_sha256": fixture_hash,
        "native_sha256": hashlib.sha256(Path(native.__file__).read_bytes()).hexdigest(),
        "source_module_sha256": hashlib.sha256(Path(styles.__file__).read_bytes()).hexdigest(),
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--child", choices=("vertical", "horizontal", "empty"))
    parser.add_argument("--before", type=Path)
    parser.add_argument("--after", type=Path)
    parser.add_argument("--before-sha")
    parser.add_argument("--after-sha")
    parser.add_argument("--python", default=sys.executable)
    parser.add_argument("--edition")
    parser.add_argument("--output", type=Path)
    parser.add_argument("--merges", type=int, default=1024)
    parser.add_argument("--queries", type=int, default=8192)
    parser.add_argument("--repeats", type=int, default=3)
    args = parser.parse_args()
    if args.child:
        print(json.dumps(child(args.child, args.merges, args.queries)))
        return
    receipt = {
        "scope": "Isolated cached merge_at point membership, excluding workbook open, XML traversal, style conversion and index preparation; not an end-to-end read or openpyxl speedup.",
        "edition": args.edition,
        "source_before": args.before_sha, "source_after": args.after_sha,
        "merged_ranges": args.merges, "queries": args.queries,
        "sampling": "One discarded fresh-process warmup per engine/case, then three alternating fresh-process samples; same source-matched native overlay within edition.",
        "bounds": "O(merges) retained storage without row/cell expansion; disjoint column buckets bisect one candidate per row-tree node. Overlapping buckets preserve first-match order with a conservative scan.",
        "cases": {},
    }
    for case in ("vertical", "horizontal", "empty"):
        samples = {"before": [], "after": []}
        for iteration in range(-1, args.repeats):
            order = ("before", "after") if iteration % 2 == 0 else ("after", "before")
            for engine in order:
                source = getattr(args, engine)
                env = dict(os.environ, PYTHONPATH=str(source / "python"))
                command = [args.python, str(Path(__file__).resolve()), "--child", case,
                           "--merges", str(args.merges), "--queries", str(args.queries)]
                sample = json.loads(subprocess.check_output(command, env=env, text=True))
                if iteration >= 0:
                    samples[engine].append(sample)
        signatures = {json.dumps(sample["signature"], sort_keys=True) for engine in samples.values() for sample in engine}
        fixtures = {sample["fixture_sha256"] for engine in samples.values() for sample in engine}
        native_hashes = {sample["native_sha256"] for engine in samples.values() for sample in engine}
        if len(signatures) != 1 or len(fixtures) != 1 or len(native_hashes) != 1:
            raise RuntimeError("Mismatched semantic signature, fixture or native overlay")
        medians = {engine: statistics.median(sample["seconds"] for sample in values) for engine, values in samples.items()}
        receipt["cases"][case] = {"samples": samples, "median_seconds": medians,
                                     "isolated_lookup_speedup": medians["before"] / medians["after"]}
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(receipt, indent=2) + "\n")
    print(json.dumps({case: data["isolated_lookup_speedup"] for case, data in receipt["cases"].items()}))


if __name__ == "__main__":
    main()
