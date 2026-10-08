"""Serial, reproducibly shuffled fresh-process measurements of frozen v2 workloads.

python benchmarks/performance_stages.py --variants variants.json --seed 20261008 \
    --fixture-dir /tmp/perf-fixtures --output /tmp/stages.json --rounds 5 --warmups 1

All remaining CLI options are passed to the frozen performance_contract parser.
This controller does not alter the harness, fixtures, or measured operations.
"""
from __future__ import annotations

import argparse
import gc
import hashlib
import importlib
import json
import os
import random
import subprocess
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from types import SimpleNamespace

import performance_contract as frozen
from performance_validation import full_verify, package_verify

CONTROLLER_VERSION = 1


def write_json(path: Path, value: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n")
    temporary.replace(path)


def load_variants(path: Path, cases: list[str]) -> tuple[list[dict], list[str]]:
    document = json.loads(path.read_text())
    variants = document if isinstance(document, list) else document["variants"]
    control_cases = cases if isinstance(document, list) else document.get("control_cases", cases)
    if set(control_cases) - set(cases):
        raise ValueError("control_cases must be a subset of measured cases")
    labels = set()
    result = []
    for definition in variants:
        variant = dict(definition)
        label = variant["label"]
        if not isinstance(label, str) or not label or label in labels:
            raise ValueError("Variant labels must be nonempty and unique")
        labels.add(label)
        role = variant.get("role", "stage")
        engine = variant.get("engine", "openpyxl" if role == "control" else "wolfxl")
        if role not in ("baseline", "stage", "final", "control"):
            raise ValueError(f"Unknown role: {role}")
        if engine != ("openpyxl" if role == "control" else "wolfxl"):
            raise ValueError("Only the control variant may measure openpyxl")
        variant.update(role=role, engine=engine)
        for field in ("python", "source_root"):
            value = Path(variant[field])
            absolute = path.parent / value if not value.is_absolute() else value
            # Resolving a venv's interpreter symlink bypasses its pyvenv.cfg.
            variant[field] = os.path.abspath(absolute) if field == "python" else str(absolute.resolve())
        if not Path(variant["python"]).is_file() or not Path(variant["source_root"]).is_dir():
            raise ValueError(f"Missing interpreter/source root for {label}")
        if len(variant["source_ref"]) != 40:
            raise ValueError("source_ref must be an exact 40-character source commit")
        pythonpath = variant.get("pythonpath")
        if pythonpath is not None:
            entries = pythonpath if isinstance(pythonpath, list) else pythonpath.split(os.pathsep)
            variant["pythonpath"] = os.pathsep.join(str((path.parent / Path(entry)).resolve()) for entry in entries)
        build = variant.get("build_metadata")
        if isinstance(build, str):
            variant["build_metadata"] = str((path.parent / build).resolve())
        result.append(variant)
    roles = [variant["role"] for variant in result]
    if roles.count("baseline") != 1 or roles.count("final") != 1 or roles.count("control") > 1:
        raise ValueError("Require exactly one baseline and final, and at most one control")
    wolf_roles = [variant["role"] for variant in result if variant["engine"] == "wolfxl"]
    if wolf_roles[0] != "baseline" or wolf_roles[-1] != "final":
        raise ValueError("List WolfXL variants in baseline, intermediate stages, final order")
    return result, control_cases


def shuffled_order(variants: list[dict], *, seed: int, case_index: int, round_index: int) -> list[dict]:
    order = list(variants)
    random.Random(f"{seed}:{case_index}:{round_index}").shuffle(order)
    return order


def worker(job: dict) -> dict:
    variant = job["variant"]
    module = importlib.import_module("wolfxl")
    identity = frozen.metadata(SimpleNamespace(label=variant["label"],
                                               source_root=Path(variant["source_root"])), module)
    if identity["source"]["commit"] != variant["source_ref"] or identity["source"]["dirty"] is not False:
        raise ValueError("Variant source HEAD changed or is dirty")
    identity["controller_sha256"] = frozen.sha256(Path(__file__))
    identity["environment"].update(engine_order="controller-scheduled-shuffle",
                                   process_model="fresh subprocess per trial")
    build = variant.get("build_metadata")
    if isinstance(build, str):
        build_path = Path(build)
        identity["build_metadata"] = {"path": build, "sha256": frozen.sha256(build_path),
                                      "data": json.loads(build_path.read_text())}
    elif isinstance(build, dict):
        identity["build_metadata"] = {"sha256": hashlib.sha256(json.dumps(build, sort_keys=True).encode()).hexdigest(),
                                      "data": build}
    else:
        identity["build_metadata"] = {"status": "not_provided"}
    fixture = Path(job["fixture"])
    if frozen.sha256(fixture) != job["fixture_sha256"]:
        raise ValueError("Cached fixture changed")
    engine = variant["engine"]
    measured_module = frozen.openpyxl if engine == "openpyxl" else module
    enabled = gc.isenabled()
    gc.collect()
    gc.disable()
    try:
        with tempfile.TemporaryDirectory(prefix="wolfxl-stage-trial-") as temporary:
            output = Path(temporary) / "edited.xlsx"
            if job["edits"]:
                sample = frozen.measure_edit(measured_module, engine, fixture, output, job["edits"])
            else:
                sample = frozen.measure_styled(measured_module, fixture, job["read_only"], job["expected"])
            sample["process_pid"] = os.getpid()
            # Deep checks are explicitly outside the frozen measured call.
            if enabled:
                gc.enable()
            if job["deep_validate"]:
                if job["edits"]:
                    seconds, semantic = frozen.timed(lambda: full_verify(output, job["expected"]))
                    indexes = {1 if sheet == "Data" else 2 for sheet, *_ in job["edits"]}
                    package_seconds, package = frozen.timed(lambda: package_verify(fixture, output, indexes, engine))
                    sample["deep_validation"] = {"status": "pass", "full_reopen_seconds": seconds,
                                                  "package_seconds": package_seconds, **semantic, **package}
                else:
                    package_seconds, package = frozen.timed(lambda: package_verify(fixture, fixture, set(), engine))
                    sample["deep_validation"] = {"status": "pass", "scope": "independent-styled-signature-and-input-integrity",
                                                  "fixture_sha256": frozen.sha256(fixture),
                                                  "package_seconds": package_seconds, **package,
                                                  "note": "Read workload emits no modified package; every sample matches independent openpyxl signature"}
    finally:
        if enabled:
            gc.enable()
    return {"identity": identity, "sample": sample}


def stable_identity(identity: dict) -> dict:
    return {key: identity[key] for key in ("harness_sha256", "controller_sha256", "environment", "packages",
                                           "native", "source", "build_metadata")}


def comparable_environment(left: dict, right: dict) -> None:
    for key in ("python", "platform", "machine", "processor", "cpu_count"):
        if left["environment"][key] != right["environment"][key]:
            raise ValueError(f"Variants have different environment.{key}")
    for key in ("version", "python_source_sha256"):
        if left["packages"]["openpyxl"][key] != right["packages"]["openpyxl"][key]:
            raise ValueError(f"Variants have different openpyxl {key}")
    if left["harness_sha256"] != right["harness_sha256"]:
        raise ValueError("Variants use different frozen harness source")


def subprocess_trial(job: dict, directory: Path, index: int) -> dict:
    job_path = directory / f"job-{index}.json"
    result_path = directory / f"result-{index}.json"
    write_json(job_path, {**job, "result_path": str(result_path)})
    environment = dict(os.environ)
    # Never let the controller's source override leak into another stage.
    environment.pop("PYTHONPATH", None)
    environment["PYTHONNOUSERSITE"] = "1"
    if job["variant"].get("pythonpath") is not None:
        environment["PYTHONPATH"] = job["variant"]["pythonpath"]
    completed = subprocess.run([job["variant"]["python"], str(Path(__file__).resolve()),
                                "--worker-job", str(job_path)], env=environment, text=True,
                               capture_output=True)
    if completed.returncode:
        raise RuntimeError(f"{job['variant']['label']}/{job['case']} failed: {completed.stderr[-6000:]}")
    return json.loads(result_path.read_text())


def stage_summary(receipt: dict) -> list[dict]:
    by_case = {}
    for result in receipt["results"]:
        results = by_case.setdefault(result["case"], {})
        if result["role"] in ("baseline", "final", "control"):
            results[result["role"]] = result
    summaries = []
    for case, results in by_case.items():
        baseline = results["baseline"]["phase_medians_seconds"]
        final = results["final"]["phase_medians_seconds"]
        stages = []
        previous = baseline
        for result in receipt["results"]:
            if result["case"] == case and result["engine"] == "wolfxl":
                phases = result["phase_medians_seconds"]
                stages.append({"label": result["label"], "role": result["role"],
                               "operation_seconds": phases["operation"], "total_seconds": phases["total"],
                               "engine_speedup_vs_baseline": baseline["operation"] / phases["operation"],
                               "total_speedup_vs_baseline": baseline["total"] / phases["total"],
                               "incremental_engine_speedup": previous["operation"] / phases["operation"],
                               "incremental_total_speedup": previous["total"] / phases["total"]})
                previous = phases
        summary = {"case": case, "cumulative_engine_speedup": baseline["operation"] / final["operation"],
                   "cumulative_total_speedup": baseline["total"] / final["total"], "stages": stages}
        if "control" in results:
            control = results["control"]["phase_medians_seconds"]
            summary.update(final_speedup_vs_openpyxl=control["operation"] / final["operation"],
                           final_total_speedup_vs_openpyxl=control["total"] / final["total"])
        summaries.append(summary)
    return summaries


def run_stages(config, variants: list[dict], control_cases: list[str], seed: int) -> dict:
    receipt = {"contract": frozen.CONTRACT, "controller_version": CONTROLLER_VERSION,
               "controller_sha256": frozen.sha256(Path(__file__)), "status": "in_progress",
               "created_utc": datetime.now(timezone.utc).isoformat(),
               "label": config.label, "variants": variants, "fixtures": {}, "identities": {},
               "reference_identity": {"openpyxl": frozen.package_fingerprint(frozen.openpyxl),
                                      "python": sys.version, "executable": sys.executable},
               "workload_config": {key: getattr(config, key) for key in
                                   ("edit_rows", "styled_rows", "small_rows", "style_cardinality", "cases", "rounds", "warmups")},
               "scheduling": {"seed": seed, "strategy": "serial shuffled stage/control order per case/round; fresh process per trial",
                              "warmup_note": "Discarded fresh-process warmups prime filesystem caches; process caches are fresh for every trial",
                              "orders": []}, "results": []}
    records = {}
    try:
        with tempfile.TemporaryDirectory(prefix="wolfxl-stage-controller-") as temporary:
            directory = Path(temporary)
            trial_index = 0
            for case_index, case in enumerate(config.cases):
                kind, rows, cardinality, read_only = frozen.workload_spec(case, config)
                fixture, manifest = frozen.prepare_fixture(config.fixture_dir, rows=rows, kind=kind, cardinality=cardinality)
                receipt["fixtures"][case] = manifest
                edits = frozen.edits_for(case[5:], rows) if case.startswith("edit_") else []
                reference_seconds, expected = frozen.timed(lambda: frozen.expected_digest(fixture, edits)) if edits else frozen.timed(
                    lambda: frozen.styled_signature(frozen.openpyxl, fixture, read_only=read_only)[1])
                active = [variant for variant in variants if variant["role"] != "control" or case in control_cases]
                for variant in active:
                    records[case, variant["label"]] = {"case": case, "label": variant["label"], "role": variant["role"],
                                                       "engine": variant["engine"], "samples": [], "warmup_samples": [],
                                                       "reference_validation_seconds": reference_seconds}
                for round_index in range(config.warmups + config.rounds):
                    order = shuffled_order(active, seed=seed, case_index=case_index, round_index=round_index)
                    receipt["scheduling"]["orders"].append({"case": case, "round": round_index - config.warmups,
                                                          "warmup": round_index < config.warmups,
                                                          "labels": [variant["label"] for variant in order]})
                    for order_index, variant in enumerate(order):
                        job = {"variant": variant, "case": case, "fixture": str(fixture.resolve()),
                               "fixture_sha256": manifest["sha256"], "edits": edits, "expected": expected,
                               "read_only": read_only, "deep_validate": round_index == config.warmups + config.rounds - 1}
                        result = subprocess_trial(job, directory, trial_index)
                        trial_index += 1
                        identity = stable_identity(result["identity"])
                        label = variant["label"]
                        for field in ("version", "python_source_sha256"):
                            if receipt["reference_identity"]["openpyxl"][field] != identity["packages"]["openpyxl"][field]:
                                raise ValueError(f"Reference and variant openpyxl {field} differ")
                        if receipt["identities"]:
                            comparable_environment(next(iter(receipt["identities"].values())), identity)
                        if label in receipt["identities"] and receipt["identities"][label] != identity:
                            raise ValueError(f"Variant identity changed during measurement: {label}")
                        receipt["identities"][label] = identity
                        record = records[case, label]
                        sample = result["sample"]
                        sample.update(round=round_index - config.warmups, order_index=order_index)
                        if "deep_validation" in sample:
                            record["deep_validation"] = sample.pop("deep_validation")
                        record["warmup_samples" if round_index < config.warmups else "samples"].append(sample)
                        receipt["results"] = list(records.values())
                        write_json(config.output, receipt)
                print(f"Completed paired stages: {case}", flush=True)
        for record in records.values():
            measured = record.pop("samples")
            record.update(frozen.phase_summary(measured))
        receipt["results"] = list(records.values())
        receipt["summary"] = stage_summary(receipt)
        receipt["status"] = "complete"
        receipt["completed_utc"] = datetime.now(timezone.utc).isoformat()
        write_json(config.output, receipt)
        return receipt
    except Exception as error:
        receipt.update(status="failed", error=f"{type(error).__name__}: {error}")
        write_json(config.output, receipt)
        raise


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__, add_help=False)
    parser.add_argument("--variants", type=Path)
    parser.add_argument("--seed", type=int, default=20261008)
    parser.add_argument("--worker-job", type=Path)
    known, remaining = parser.parse_known_args(argv)
    if known.worker_job:
        job = json.loads(known.worker_job.read_text())
        write_json(Path(job["result_path"]), worker(job))
    else:
        config = frozen.parse_args(remaining)
        if known.variants is None:
            parser.error("--variants is required")
        variants, control_cases = load_variants(known.variants.resolve(), config.cases)
        run_stages(config, variants, control_cases, known.seed)


if __name__ == "__main__":
    main()
