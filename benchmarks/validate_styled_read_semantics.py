"""Independent, untimed Cell semantics checks for complete v2 performance receipts.

Run after the timing lane is released. Each case/variant gets a fresh process
using the interpreter and installed Python/native identities from its receipt.
The reference uses independent openpyxl in exactly the same read mode. This
producer does not modify the frozen performance contract or its receipts.
"""
from __future__ import annotations

import argparse
import hashlib
import importlib
import itertools
import json
import math
import os
import re
import subprocess
import sys
import tempfile
import traceback
from collections import Counter
from datetime import date, datetime, time, timedelta, timezone
from pathlib import Path
from xml.etree import ElementTree
from zipfile import ZipFile


CASES = (
    "styled_small_eager", "styled_large_read_only", "styled_merged_eager",
    "styled_large_eager", "styled_sparse_eager", "styled_high_cardinality_eager",
    "styled_sparse_read_only", "styled_high_cardinality_read_only",
)
MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
FIELDS = ("value", "python_type", "data_type", "font", "fill", "border", "alignment", "number_format")
MISSING = object()


def encoded(value) -> bytes:
    return json.dumps(value, sort_keys=True, separators=(",", ":"), ensure_ascii=False).encode()


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def scalar(value):
    if isinstance(value, float):
        if not math.isfinite(value):
            return {"float": repr(value)}
        return int(value) if value.is_integer() else value
    if value is None or isinstance(value, (str, int, bool)):
        return value
    if isinstance(value, (datetime, date, time)):
        return {type(value).__name__: value.isoformat()}
    if isinstance(value, timedelta):
        return {"timedelta_microseconds": (value.days * 86400 + value.seconds) * 1000000 + value.microseconds}
    return {"python_type": type(value).__name__, "repr": repr(value)}


def color(value):
    if value is None:
        return None
    kind = getattr(value, "type", None)
    active = getattr(value, kind, None) if kind in {"rgb", "indexed", "theme", "auto"} else None
    if kind == "rgb" and isinstance(active, str):
        active = active.lstrip("#").upper()
    return {"selector": kind, "value": scalar(active), "tint": scalar(getattr(value, "tint", 0) or 0)}


def font(value):
    result = {key: scalar(getattr(value, key, None)) for key in (
        "name", "sz", "u", "vertAlign", "scheme", "family", "charset",
    )}
    result.update({key: bool(getattr(value, key, False)) for key in (
        "b", "i", "strike", "outline", "shadow", "condense", "extend",
    )})
    result["color"] = color(getattr(value, "color", None))
    return result


def fill(value):
    if value is not None and type(value).__name__ == "GradientFill":
        return {"kind": "gradient", "type": getattr(value, "type", None),
                **{key: scalar(getattr(value, key, 0) or 0) for key in ("degree", "left", "right", "top", "bottom")},
                "stops": [[scalar(stop.position), color(stop.color)] for stop in value.stop]}
    # Read-only EmptyCell has no component objects. Its absent pattern and
    # implicit zero colors mean the same unstyled fill as a default object.
    zero = {"selector": "rgb", "value": "00000000", "tint": 0}
    return {"kind": "pattern", "patternType": getattr(value, "patternType", None),
            "fgColor": color(getattr(value, "fgColor", None)) or zero,
            "bgColor": color(getattr(value, "bgColor", None)) or zero}


def border(value):
    result = {"outline": True if getattr(value, "outline", None) is None else bool(value.outline),
              "diagonalUp": bool(getattr(value, "diagonalUp", False)),
              "diagonalDown": bool(getattr(value, "diagonalDown", False))}
    for key in ("left", "right", "top", "bottom", "diagonal", "vertical", "horizontal", "start", "end"):
        side = getattr(value, key, None)
        result[key] = {"style": getattr(side, "style", None), "color": color(getattr(side, "color", None))}
    return result


def alignment(value):
    return {"horizontal": getattr(value, "horizontal", None) or "general",
            "vertical": getattr(value, "vertical", None) or "bottom",
            **{key: scalar(getattr(value, key, 0) or 0) for key in ("textRotation", "indent", "relativeIndent", "readingOrder")},
            **{key: bool(getattr(value, key, False)) for key in ("wrapText", "shrinkToFit", "justifyLastLine")}}


class StyleCache:
    """Bounded canonicalization of returned style objects, never style IDs."""

    def __init__(self):
        self.cache = {}

    def get(self, field: str, component):
        # Openpyxl's per-cell StyleProxy delegates to an immutable source-table
        # object. Retain a strong reference so object-ID reuse cannot hit cache.
        target = getattr(component, "_StyleProxy__target", component)
        key = field, id(target)
        old = self.cache.get(key)
        if old is not None and old[0] is target:
            return old[1], old[2]
        meaning = {"font": font, "fill": fill, "border": border, "alignment": alignment}[field](target)
        if len(self.cache) >= 4096:
            self.cache.clear()
        fingerprint = hashlib.sha256(encoded(meaning)).hexdigest()
        self.cache[key] = target, meaning, fingerprint
        return meaning, fingerprint


def public_fields(cell, cache: StyleCache):
    result, digest_fields = {}, {}
    for field in FIELDS:
        attribute = "value" if field == "python_type" else field
        try:
            value = getattr(cell, attribute)
            if field == "python_type":
                result[field] = type(value).__name__
            elif field == "number_format":
                result[field] = value or "General"
            elif field in {"font", "fill", "border", "alignment"}:
                result[field], digest_fields[field] = cache.get(field, value)
            else:
                result[field] = scalar(value)
        except Exception as error:
            result[field] = {"unavailable": type(error).__name__, "message": str(error)}
        if field not in digest_fields:
            digest_fields[field] = result[field]
    return result, digest_fields


class Mismatches:
    def __init__(self):
        self.counts = Counter()
        self.digests = {}
        self.examples = {}
        self.scopes = {}

    def compare(self, field, coordinate, expected, actual, scope="metadata", digest_values=None):
        if expected == actual:
            return
        self.counts[field] += 1
        self.scopes.setdefault(field, Counter())[scope] += 1
        expected_digest, actual_digest = digest_values if digest_values is not None else (expected, actual)
        self.digests.setdefault(field, hashlib.sha256()).update(encoded([coordinate, scope, expected_digest, actual_digest]))
        examples = self.examples.setdefault(field, [])
        if len(examples) < 12:
            examples.append({"coordinate": coordinate, "source_scope": scope, "expected": expected, "actual": actual})

    def report(self):
        return {field: {"count": count, "all_mismatches_sha256": self.digests[field].hexdigest(),
                        "source_scope_counts": dict(self.scopes[field]),
                        "first_examples": self.examples[field]} for field, count in sorted(self.counts.items())}


def package_fingerprint(module):
    # Same ordered Python-source algorithm as the frozen timing producer.
    directory = Path(module.__file__).resolve().parent
    source = hashlib.sha256()
    count = 0
    for path in sorted(directory.rglob("*.py")):
        source.update(str(path.relative_to(directory)).encode())
        source.update(path.read_bytes())
        count += 1
    return {"version": module.__version__, "import_path": str(module.__file__),
            "python_source_sha256": source.hexdigest(), "python_source_files": count}


def source_coordinate(value):
    match = re.fullmatch(r"([A-Z]+)([1-9][0-9]*)", value)
    assert match, f"Unexpected source coordinate: {value}"
    column = 0
    for letter in match[1]:
        column = column * 26 + ord(letter) - ord("A") + 1
    return int(match[2]), column


def worksheet_source(path: Path):
    with ZipFile(path) as archive:
        root = ElementTree.fromstring(archive.read("xl/worksheets/sheet1.xml"))
    merges = sorted(item.attrib["ref"] for item in root.findall(f"{{{MAIN_NS}}}mergeCells/{{{MAIN_NS}}}mergeCell"))
    cells = {}
    for cell in root.findall(f"{{{MAIN_NS}}}sheetData/{{{MAIN_NS}}}row/{{{MAIN_NS}}}c"):
        has_value = any(cell.find(f"{{{MAIN_NS}}}{tag}") is not None for tag in ("v", "is", "f"))
        cells[source_coordinate(cell.attrib["r"])] = "stored_value" if has_value else "stored_blank"
    placeholders = set()
    for merge in merges:
        start, end = (source_coordinate(coord) for coord in merge.split(":"))
        for row in range(start[0], end[0] + 1):
            for column in range(start[1], end[1] + 1):
                if (row, column) != start:
                    placeholders.add((row, column))
    return merges, cells, placeholders


def worker(plan):
    # This process runs inside the frozen variant interpreter, without a local
    # source path override; openpyxl remains an independent installed module.
    reference_module = importlib.import_module("openpyxl")
    actual_module = importlib.import_module("wolfxl")
    native_module = importlib.import_module("wolfxl._rust")
    expected_identity = plan["identity"]
    actual_identity = {"python": sys.version, "packages": {
        "openpyxl": package_fingerprint(reference_module), "wolfxl": package_fingerprint(actual_module)},
        "native": {"path": str(native_module.__file__), "sha256": sha256(Path(native_module.__file__))}}
    assert actual_identity["python"] == expected_identity["environment"]["python"]
    assert actual_identity["packages"] == expected_identity["packages"], "Installed Python package identity changed"
    assert actual_identity["native"] == expected_identity["native"], "Installed native identity changed"
    fixture = Path(plan["fixture_path"])
    assert sha256(fixture) == plan["fixture"]["sha256"], "Fixture changed"
    source_merges, source_cells, source_placeholders = worksheet_source(fixture)
    read_only = plan["case"].endswith("read_only")
    mismatches = Mismatches()
    reference = reference_module.load_workbook(fixture, read_only=read_only, data_only=True)
    actual = actual_module.load_workbook(fixture, read_only=read_only, data_only=True)
    checked = 0
    reference_digest, actual_digest = hashlib.sha256(), hashlib.sha256()
    counts = {"reference_python_types": Counter(), "actual_python_types": Counter()}
    scope_counts = Counter()
    reference_cache, actual_cache = StyleCache(), StyleCache()
    metadata = {}
    try:
        mismatches.compare("sheetnames", "workbook", reference.sheetnames, actual.sheetnames)
        assert reference.sheetnames == ["Data"], "This fixture supplement expects one Data worksheet"
        for name in reference.sheetnames:
            if name not in actual.sheetnames:
                continue
            expected_sheet, actual_sheet = reference[name], actual[name]
            expected_geometry = [expected_sheet.max_row, expected_sheet.max_column]
            actual_geometry = [actual_sheet.max_row, actual_sheet.max_column]
            mismatches.compare("geometry", name, expected_geometry, actual_geometry)
            # Openpyxl RO does not expose merged_cells; independent package XML
            # supplies its merge metadata without switching its read mode.
            expected_merges = source_merges if read_only else sorted(str(r) for r in expected_sheet.merged_cells.ranges)
            actual_merges = sorted(str(r) for r in actual_sheet.merged_cells.ranges)
            mismatches.compare("merges", name, expected_merges, actual_merges)
            metadata[name] = {"reference_geometry": expected_geometry, "actual_geometry": actual_geometry,
                              "reference_merges": expected_merges, "actual_merges": actual_merges}
            for row_index, pair in enumerate(itertools.zip_longest(expected_sheet.iter_rows(), actual_sheet.iter_rows(), fillvalue=MISSING), 1):
                expected_row, actual_row = pair
                if expected_row is MISSING or actual_row is MISSING:
                    mismatches.compare("row_presence", [name, row_index], expected_row is not MISSING, actual_row is not MISSING)
                    continue
                mismatches.compare("row_width", [name, row_index], len(expected_row), len(actual_row))
                for col_index, cells in enumerate(itertools.zip_longest(expected_row, actual_row, fillvalue=MISSING), 1):
                    expected_cell, actual_cell = cells
                    coordinate = [name, row_index, col_index]
                    scope = source_cells.get((row_index, col_index), "implicit_blank")
                    if (row_index, col_index) in source_placeholders:
                        scope += "/merged_placeholder"
                    scope_counts[scope] += 1
                    if expected_cell is MISSING or actual_cell is MISSING:
                        mismatches.compare("cell_presence", coordinate, expected_cell is not MISSING, actual_cell is not MISSING)
                        continue
                    # RO EmptyCell has no row/column metadata. Where present,
                    # validate public coordinates against iterator position.
                    for field, index in (("row", row_index), ("column", col_index)):
                        reported = getattr(actual_cell, field, index)
                        mismatches.compare("coordinate_" + field, coordinate, index, reported, scope)
                    left, left_digest = public_fields(expected_cell, reference_cache)
                    right, right_digest = public_fields(actual_cell, actual_cache)
                    counts["reference_python_types"][left["python_type"] if isinstance(left["python_type"], str) else "unavailable"] += 1
                    counts["actual_python_types"][right["python_type"] if isinstance(right["python_type"], str) else "unavailable"] += 1
                    reference_digest.update(encoded([coordinate, left_digest]))
                    actual_digest.update(encoded([coordinate, right_digest]))
                    for field in FIELDS:
                        mismatches.compare(field, coordinate, left[field], right[field], scope,
                                           (left_digest[field], right_digest[field]))
                    checked += 1
    finally:
        reference.close()
        actual.close()
    assert sha256(fixture) == plan["fixture"]["sha256"], "Read changed original input bytes"
    return {"case": plan["case"], "variant": plan["variant"], "read_only": read_only,
            "data_only": True, "status": "pass" if not mismatches.counts else "mismatch",
            "cells_compared": checked, "metadata": metadata, "mismatches": mismatches.report(),
            "source_scope_counts": dict(scope_counts),
            "reference_full_digest": reference_digest.hexdigest(), "actual_full_digest": actual_digest.hexdigest(),
            "type_counts": dict(counts), "observed_identity": actual_identity,
            "bound_source": expected_identity["source"], "bound_native": expected_identity["native"],
            "input_unchanged": True}


def render(report):
    lines = [f"# {report['edition']}: independent styled-read semantic validation", "",
             "This is an untimed validation supplement, not a replacement timing receipt or Microsoft Excel certification.", "",
             "Every iterator coordinate is checked for value, exact Python value type, Excel data_type, font, fill, border, alignment and number format. Geometry and merge metadata are checked separately. Both engines use the same read_only and data_only=True mode. No read-only worksheet is converted to eager Cells.", "",
             "Style meanings use component fields rather than style IDs. Schema-default booleans/numbers, implicit alignment defaults and absent read-only EmptyCell fills are normalized; active color selector/value/tint is retained. Returned style objects are canonicalized with a bounded identity cache. Full-cell digests bind component-meaning SHA256 hashes, so identical style dictionaries are not serialized again at every coordinate. Protection, conditional formatting, dates not present in these fixtures, calculation and rendering are outside this supplement.", "",
             "Mismatch categories retain source-XML scope: stored_value and stored_blank are authored cells; implicit_blank is an un-authored iterator placeholder. The merged_placeholder suffix identifies a non-anchor coordinate in an authored merge. Coordinate/value/type checks apply to every scope. Style API differences for implicit blanks are reported as representation boundaries separately from authored-cell style differences; no missing authored style is normalized away.", "",
             "Numeric value equality normalizes integral float/int values, while exact Python numeric subtype differences are reported separately, without treating them as silently equivalent API types.", "",
             "| Case | Mode | Cells/variant | Baseline | Final | Classification |", "|---|---|---:|---|---|---|"]
    for item in report["comparison"]:
        baseline, final = item["baseline"], item["final"]
        lines.append(f"| {item['case']} | {'read-only' if final.get('read_only') else 'eager'} | {final.get('cells_compared', 0):,} | {baseline['status']} | {final['status']} | {item['classification']} |")
    lines.extend(["", "All mismatch counts and ordered mismatch digests are retained in JSON; each category includes its first twelve examples. A mismatch present on the original baseline is a baseline boundary, not evidence that the final changes introduced it. Changed baseline/final mismatch sets remain explicitly classified as changed rather than assumed fixed.", "",
                  "## Mismatch categories", ""])
    for item in report["comparison"]:
        for role in ("baseline", "final"):
            result = item[role]
            if result.get("mismatches"):
                lines.append(f"- {item['case']}/{role}: " + ", ".join(f"{field}={details['count']:,}" for field, details in result["mismatches"].items()))
                for field, details in result["mismatches"].items():
                    lines.append(f"  - {field}: " + ", ".join(f"{scope}={count:,}" for scope, count in details["source_scope_counts"].items()))
            elif result["status"] == "error":
                lines.append(f"- {item['case']}/{role}: ERROR {result['error']}")
    if not any(r.get("mismatches") or r["status"] == "error" for c in report["comparison"] for r in (c["baseline"], c["final"])):
        lines.append("All compared fields matched on every coordinate.")
    lines.extend(["", "## Binding", "", f"- Producer SHA256: `{report['producer_sha256']}`"])
    for receipt in report["receipts"]:
        lines.append(f"- {receipt['name']} receipt SHA256: `{receipt['sha256']}`")
    for role, identity in report["identities"].items():
        lines.extend([f"- {role} source: `{identity['source']['commit']}`", f"- {role} native SHA256: `{identity['native']['sha256']}`"])
    lines.extend(["", "Installed package Python/native hashes were checked against the frozen receipts in every worker; source fixture hashes were checked before and after each scan. This report never modifies the timing receipts or the input workbooks.", ""])
    return "\n".join(lines)


def run(args):
    receipts = [json.loads(path.read_text()) for path in (args.core, args.guards)]
    assert all(r["status"] == "complete" for r in receipts), "Timing receipts must be complete"
    assert receipts[0]["identities"] == receipts[1]["identities"], "Core/guard variant identities differ"
    identities = receipts[0]["identities"]
    fixtures = {case: fixture for receipt in receipts for case, fixture in receipt["fixtures"].items() if case in CASES}
    assert set(fixtures) == set(CASES), "All eight styled-read cases are required"
    index = {}
    for manifest in args.fixture_dir.glob("*.json"):
        metadata = json.loads(manifest.read_text())
        index[metadata["sha256"]] = manifest.with_suffix(".xlsx")
    results = []
    environment = dict(os.environ)
    environment.pop("PYTHONPATH", None)
    environment["PYTHONNOUSERSITE"] = "1"
    with tempfile.TemporaryDirectory(prefix="wolfxl-styled-semantics-") as directory:
        temporary = Path(directory)
        for case in CASES:
            for role in ("baseline", "final"):
                identity = identities[role]
                plan = {"case": case, "variant": role, "fixture": fixtures[case], "identity": identity,
                        "fixture_path": str(index[fixtures[case]["sha256"]])}
                plan_path, result_path = temporary / "plan.json", temporary / "result.json"
                plan_path.write_bytes(encoded(plan))
                completed = subprocess.run([identity["environment"]["executable"], str(Path(__file__).resolve()),
                                            "--worker", str(plan_path), "--worker-output", str(result_path)],
                                           env=environment, cwd=temporary, capture_output=True, text=True)
                if result_path.exists():
                    result = json.loads(result_path.read_text())
                    result_path.unlink()
                else:
                    result = {"case": case, "variant": role, "status": "error", "error": completed.stderr[-8000:]}
                results.append(result)
                print(f"{args.edition}/{case}/{role}: {result['status']}; cells={result.get('cells_compared', 0)}", flush=True)
    comparison = []
    for case in CASES:
        pair = {r["variant"]: r for r in results if r["case"] == case}
        baseline, final = pair["baseline"], pair["final"]
        if "error" in {baseline["status"], final["status"]}:
            classification = "validation error; not certified"
        elif baseline["status"] == final["status"] == "pass":
            classification = "pass on both"
        elif baseline["status"] == "pass":
            classification = "final-only mismatch"
        elif final["status"] == "pass":
            classification = "baseline mismatch resolved"
        elif {k: (v["count"], v["all_mismatches_sha256"]) for k, v in baseline["mismatches"].items()} == {k: (v["count"], v["all_mismatches_sha256"]) for k, v in final["mismatches"].items()}:
            classification = "identical baseline boundary"
        else:
            classification = "baseline and final mismatches differ; inspect categories"
        comparison.append({"case": case, "classification": classification, **pair})
    report = {"contract": "wolfxl-independent-styled-read-semantics-v1", "edition": args.edition,
              "created_utc": datetime.now(timezone.utc).isoformat(), "status": "complete",
              "timing_scope": "No performance timing; executed after the paired timing lane was released",
              "producer_sha256": sha256(Path(__file__)), "identities": identities,
              "receipts": [{"name": path.name, "path": str(path.resolve()), "sha256": sha256(path)} for path in (args.core, args.guards)],
              "fixtures": fixtures, "normalization_and_limits": "See generated Markdown; all mismatch counts/digests remain explicit",
              "comparison": comparison}
    args.output_dir.mkdir(parents=True, exist_ok=True)
    (args.output_dir / "styled-read-validation.json").write_text(json.dumps(report, indent=2, sort_keys=True) + "\n")
    (args.output_dir / "styled-read-validation.md").write_text(render(report))
    return all(r["status"] == "pass" for r in results)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--worker", type=Path, help=argparse.SUPPRESS)
    parser.add_argument("--worker-output", type=Path, help=argparse.SUPPRESS)
    parser.add_argument("--edition")
    parser.add_argument("--core", type=Path)
    parser.add_argument("--guards", type=Path)
    parser.add_argument("--fixture-dir", type=Path)
    parser.add_argument("--output-dir", type=Path)
    args = parser.parse_args()
    if args.worker:
        try:
            result = worker(json.loads(args.worker.read_text()))
        except Exception:
            plan = json.loads(args.worker.read_text())
            result = {"case": plan["case"], "variant": plan["variant"], "status": "error", "error": traceback.format_exc()}
        args.worker_output.write_text(json.dumps(result, sort_keys=True))
    else:
        if not all((args.edition, args.core, args.guards, args.fixture_dir, args.output_dir)):
            parser.error("edition, core, guards, fixture-dir and output-dir are required")
        sys.exit(0 if run(args) else 1)


if __name__ == "__main__":
    main()
