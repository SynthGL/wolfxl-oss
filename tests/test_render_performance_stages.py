from __future__ import annotations

import copy
import importlib
import json
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "benchmarks"))
renderer = importlib.import_module("render_performance_stages")


def example():
    variants = [{"label": label, "role": role, "engine": engine, "source_ref": letter * 40}
                for label, role, engine, letter in (("baseline", "baseline", "wolfxl", "a"),
                                                    ("final", "final", "wolfxl", "b"),
                                                    ("openpyxl", "control", "openpyxl", "c"))]
    results, identities = [], {}
    def sample(value):
        return {"phase_seconds": {"operation": value, "total": value + 1}}
    for variant, seconds in zip(variants, (12, 3, 24)):
        label = variant["label"]
        results.append({"label": label, "role": variant["role"], "engine": variant["engine"],
                        "case": "edit_top", "raw_samples": [sample(seconds), sample(seconds + 1)],
                        "warmup_samples": [sample(seconds)],
                        "phase_medians_seconds": {"operation": seconds + 0.5, "total": seconds + 1.5},
                        "deep_validation": {"status": "pass", "full_reopen_seconds": 10, "package_seconds": 1}})
        identities[label] = {"source": {"commit": variant["source_ref"], "dirty": False},
                             "controller_sha256": "controller", "harness_sha256": "harness",
                             "native": {"sha256": "native"},
                             "packages": {"wolfxl": {"version": "2.0.9", "python_source_sha256": "python"},
                                          "openpyxl": {"version": "3.1.5", "python_source_sha256": "openpyxl"}},
                             "environment": {"executable": "/exact/python", "python": "3.12",
                                             "platform": "Linux", "machine": "x86_64", "cpu_count": 8},
                             "build_metadata": {"sha256": "build", "data": {"native_sha256": "native"}}}
    return {"status": "complete", "contract": "wolfxl-styled-edit-v2", "label": "Fixture evidence",
            "controller_sha256": "controller", "variants": variants, "results": results, "identities": identities,
            "workload_config": {"rounds": 2, "warmups": 1, "cases": ["edit_top"]},
            "fixtures": {"edit_top": {"rows_per_sheet": 12, "populated_value_cells": 60,
                                      "max_column": 5, "sha256": "fixture"}},
            "summary": [{"cumulative_engine_speedup": 999}]}


def test_renderer_computes_ratios_from_raw_samples_and_discloses_variance(tmp_path):
    receipt = example()
    path = tmp_path / "receipt.json"
    path.write_text(json.dumps(receipt))
    markdown = renderer.render(receipt, path)
    assert "3.571×" in markdown
    assert "999" not in markdown
    assert "**High variance:**" in markdown
    assert "outside operation and total" in markdown
    assert "not Microsoft Excel certification" in markdown
    assert "Fixture evidence" in markdown


@pytest.mark.parametrize("mutation,message", [
    (lambda receipt: receipt.update(status="failed"), "complete"),
    (lambda receipt: receipt["results"][0]["phase_medians_seconds"].update(operation=999), "raw samples"),
    (lambda receipt: receipt["identities"]["final"]["source"].update(dirty=True), "dirty"),
    (lambda receipt: receipt["identities"]["final"]["native"].update(sha256="changed"), "build metadata"),
    (lambda receipt: receipt["results"][0]["deep_validation"].update(status="failed"), "validation"),
])
def test_renderer_refuses_invalid_or_unverified_receipts(mutation, message):
    receipt = copy.deepcopy(example())
    mutation(receipt)
    with pytest.raises(ValueError, match=message):
        renderer.validate(receipt)
