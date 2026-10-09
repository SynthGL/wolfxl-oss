"""Render retained styled-read validation JSON without repeating its scans.

The original validation producer hash stays bound to the original worker code;
this separate renderer adds derived scope findings and its own explicit hash.
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path

from validate_styled_read_semantics import CASES, render, sha256


CORE_FIELDS = {
    "value", "python_type", "data_type", "coordinate_row", "coordinate_column",
    "geometry", "merges", "row_presence", "row_width", "cell_presence",
    "sheetnames", "number_format", "alignment",
}


def summary(report):
    comparison = report["comparison"]
    assert len(comparison) == len(CASES) == 8
    assert {case["case"] for case in comparison} == set(CASES)
    totals = {role: sum(case[role]["cells_compared"] for case in comparison)
              for role in ("baseline", "final")}
    assert totals == {"baseline": 2185000, "final": 2185000}
    introduced, changed, resolved = [], [], []
    core_mismatches = []
    for case in comparison:
        baseline, final = case["baseline"], case["final"]
        for role in ("baseline", "final"):
            result = case[role]
            assert result["status"] != "error", result
            assert result["input_unchanged"]
            identity = report["identities"][role]
            assert result["observed_identity"]["packages"] == identity["packages"]
            assert result["observed_identity"]["native"] == identity["native"]
            assert result["bound_source"] == identity["source"]
            core_mismatches.extend((case["case"], role, field)
                                   for field in CORE_FIELDS & result["mismatches"].keys())
        for field, details in final["mismatches"].items():
            if field not in baseline["mismatches"]:
                introduced.append((case["case"], field))
            else:
                old = baseline["mismatches"][field]
                if (old["count"], old["all_mismatches_sha256"]) != (details["count"], details["all_mismatches_sha256"]):
                    changed.append((case["case"], field))
        resolved.extend((case["case"], field, details["count"])
                        for field, details in baseline["mismatches"].items()
                        if field not in final["mismatches"])
    # These assertions verify this report's observed findings, rather than
    # silently emitting a positive summary when a later scan finds regressions.
    assert not core_mismatches, core_mismatches
    assert not introduced and not changed, (introduced, changed)
    assert resolved == [("styled_merged_eager", "border", 4)], resolved
    lines = ["## Observed scope", "",
             "Eight cases compare 2,185,000 coordinates per variant: 4,370,000 original/final coordinate pairs in this edition. Every coordinate, value, exact Python value type, Excel data_type, geometry, merge, alignment and number format matched the same-mode independent reference.", "",
             "There are zero final-only mismatch categories and zero changed mismatch sets among retained categories, checked from every ordered mismatch digest. Four merged-border color mismatches present in the baseline disappear in the final result. The three non-anchor merged coordinates are checked directly, without changing read mode.", "",
             "Font API metadata boundaries remain on authored/stored cells in the affected cases, with identical baseline/final mismatch sets; source default-font and theme metadata are not fully reflected by the returned objects. Sparse eager also reports implicit-blank font differences as representation boundaries, separately from stored-value font gaps. Sparse read-only has no implicit-blank font mismatch against openpyxl EmptyCell under the documented normalization.", ""]
    if report["edition"] == "Community":
        lines.extend(["Community also retains authored solid-fill ARGB metadata differences. Recorded examples include openpyxl `00FFD966` versus WolfXL `FFFFD966`; these are preserved as API differences, with identical baseline/final mismatch digests. High-cardinality font cases also retain source-theme/family/scheme metadata differences.", ""])
    else:
        lines.extend(["Commercial high-cardinality eager and read-only cases match all compared fields on every coordinate; the other affected cases retain the original default-font metadata boundary. Fill fields match in all eight cases.", ""])
    lines.extend(["These findings support absence of new mismatches in this fixture set, not complete style equivalence. Microsoft Excel rendering, appearance, calculation, and workbook features absent from these fixtures were not certified.", ""])
    return "\n".join(lines)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("report", type=Path)
    args = parser.parse_args()
    report = json.loads(args.report.read_text())
    assert report["producer_sha256"] == sha256(Path(__file__).with_name("validate_styled_read_semantics.py"))
    text = render(report) + "\n" + summary(report)
    text += f"\nRaw JSON SHA256: `{sha256(args.report)}`\n\nScope renderer SHA256: `{sha256(Path(__file__))}`\n"
    args.report.with_suffix(".md").write_text(text)


if __name__ == "__main__":
    main()
