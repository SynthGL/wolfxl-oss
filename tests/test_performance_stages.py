from __future__ import annotations

import importlib
import json
import sys
from pathlib import Path
from types import SimpleNamespace

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "benchmarks"))
stages = importlib.import_module("performance_stages")


def definitions(tmp_path):
    return [{"label": label, "role": role, "python": sys.executable,
             "source_root": str(tmp_path), "source_ref": "a" * 40}
            for label, role in (("before", "baseline"), ("middle", "stage"), ("after", "final"), ("openpyxl", "control"))]


def test_variant_spec_and_shuffle_are_reproducible(tmp_path):
    path = tmp_path / "variants.json"
    path.write_text(json.dumps({"variants": definitions(tmp_path), "control_cases": ["edit_top"]}))
    variants, cases = stages.load_variants(path, ["edit_top", "edit_bottom"])
    assert cases == ["edit_top"]
    assert variants[0]["python"] == str(Path(sys.executable).absolute())
    first = stages.shuffled_order(variants, seed=123, case_index=0, round_index=0)
    assert first == stages.shuffled_order(variants, seed=123, case_index=0, round_index=0)
    assert {item["label"] for item in first} == {item["label"] for item in variants}
    assert first != stages.shuffled_order(variants, seed=123, case_index=0, round_index=1)
    invalid = definitions(tmp_path)
    invalid[1]["label"] = "before"
    path.write_text(json.dumps(invalid))
    with pytest.raises(ValueError, match="unique"):
        stages.load_variants(path, ["edit_top"])


def test_trial_removes_inherited_pythonpath_and_uses_exact_interpreter(tmp_path, monkeypatch):
    monkeypatch.setenv("PYTHONPATH", "/wrong-stage")
    calls = []
    def fake_run(argv, **kwargs):
        calls.append((argv, kwargs))
        job = json.loads(Path(argv[-1]).read_text())
        stages.write_json(Path(job["result_path"]), {"ok": True})
        return SimpleNamespace(returncode=0)
    monkeypatch.setattr(stages.subprocess, "run", fake_run)
    job = {"variant": {"label": "before", "python": "/exact/python"}, "case": "edit_top"}
    assert stages.subprocess_trial(job, tmp_path, 0) == {"ok": True}
    argv, kwargs = calls[-1]
    assert argv[0] == "/exact/python"
    assert "PYTHONPATH" not in kwargs["env"]
    assert kwargs["env"]["PYTHONNOUSERSITE"] == "1"
    job["variant"]["pythonpath"] = "/chosen-stage"
    stages.subprocess_trial(job, tmp_path, 1)
    assert calls[-1][1]["env"]["PYTHONPATH"] == "/chosen-stage"


def test_direct_cumulative_and_incremental_ratios():
    records = [{"case": "one", "label": label, "role": role, "engine": engine,
                "phase_medians_seconds": {"operation": seconds, "total": seconds + 1}}
               for label, role, engine, seconds in (("before", "baseline", "wolfxl", 12),
                                                    ("middle", "stage", "wolfxl", 6),
                                                    ("after", "final", "wolfxl", 3),
                                                    ("openpyxl", "control", "openpyxl", 24))]
    summary = stages.stage_summary({"results": records})[0]
    assert summary["cumulative_engine_speedup"] == 4
    assert summary["final_speedup_vs_openpyxl"] == 8
    assert [row["incremental_engine_speedup"] for row in summary["stages"]] == [1, 2, 2]


def test_environment_mismatch_is_rejected():
    identity = {"environment": dict.fromkeys(("python", "platform", "machine", "processor", "cpu_count"), "same"),
                "packages": {"openpyxl": {"version": "3.1.5", "python_source_sha256": "one"}},
                "harness_sha256": "frozen"}
    different = json.loads(json.dumps(identity))
    different["packages"]["openpyxl"]["version"] = "different"
    with pytest.raises(ValueError, match="openpyxl version"):
        stages.comparable_environment(identity, different)


def test_controller_schedules_fresh_trials_and_one_deep_validation(tmp_path, monkeypatch):
    config = stages.frozen.parse_args(["--output", str(tmp_path / "receipt.json"),
                                      "--fixture-dir", str(tmp_path / "fixtures"), "--cases", "edit_top,styled_small_eager",
                                      "--edit-rows", "12", "--small-rows", "12", "--rounds", "2", "--warmups", "1"])
    variants = definitions(tmp_path)
    for variant in variants:
        variant["engine"] = "openpyxl" if variant["role"] == "control" else "wolfxl"
    calls = []
    openpyxl_identity = stages.frozen.package_fingerprint(stages.frozen.openpyxl)
    def fake_trial(job, directory, index):
        calls.append(job)
        sample = {"phase_seconds": {"operation": index + 1, "total": index + 2}}
        if job["deep_validate"]:
            sample["deep_validation"] = {"status": "pass"}
        identity = {"harness_sha256": "frozen", "controller_sha256": "controller",
                    "environment": dict.fromkeys(("python", "platform", "machine", "processor", "cpu_count"), "same"),
                    "packages": {"openpyxl": openpyxl_identity},
                    "native": {}, "source": {}, "build_metadata": {}}
        return {"identity": identity, "sample": sample}
    monkeypatch.setattr(stages, "subprocess_trial", fake_trial)
    receipt = stages.run_stages(config, variants, ["edit_top"], seed=123)
    assert receipt["status"] == "complete"
    assert len(calls) == 21  # (4 edit variants + 3 styled variants) * (2 rounds + 1 warmup)
    assert sum(job["deep_validate"] for job in calls) == 7
    assert all(len(result["raw_samples"]) == 2 and len(result["warmup_samples"]) == 1
               for result in receipt["results"])
    assert all(result["deep_validation"]["status"] == "pass" for result in receipt["results"])
    assert len(receipt["scheduling"]["orders"]) == 6
