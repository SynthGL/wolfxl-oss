"""Render a complete matched-stage receipt as source-backed Markdown evidence.

python benchmarks/render_performance_stages.py paired-stages.json --output evidence.md
"""
from __future__ import annotations

import argparse
import hashlib
import json
import math
import statistics
from pathlib import Path


def medians(record: dict) -> dict:
    samples = record["raw_samples"]
    if not samples:
        raise ValueError("No measured samples")
    phases = set(samples[0]["phase_seconds"])
    for sample in samples:
        if set(sample["phase_seconds"]) != phases:
            raise ValueError("Raw phase coverage differs")
        if any(not math.isfinite(value) or value < 0 for value in sample["phase_seconds"].values()):
            raise ValueError("Invalid phase duration")
    return {phase: statistics.median(sample["phase_seconds"][phase] for sample in samples) for phase in phases}


def validate(receipt: dict) -> None:
    if receipt.get("status") != "complete":
        raise ValueError("Only a complete receipt may be rendered")
    variants = {variant["label"]: variant for variant in receipt["variants"]}
    roles = [variant["role"] for variant in variants.values()]
    if roles.count("baseline") != 1 or roles.count("final") != 1:
        raise ValueError("Missing unique baseline/final")
    seen = set()
    for record in receipt["results"]:
        key = record["case"], record["label"]
        if key in seen:
            raise ValueError("Duplicate result")
        seen.add(key)
        if len(record["raw_samples"]) != receipt["workload_config"]["rounds"]:
            raise ValueError("Incomplete measured rounds")
        if len(record["warmup_samples"]) != receipt["workload_config"]["warmups"]:
            raise ValueError("Incomplete warmup rounds")
        if record["deep_validation"].get("status") != "pass":
            raise ValueError("Deep validation did not pass")
        computed = medians(record)
        for phase, value in computed.items():
            if not math.isclose(record["phase_medians_seconds"][phase], value, rel_tol=1e-12, abs_tol=1e-12):
                raise ValueError("Stored medians disagree with raw samples")
        variant = variants[record["label"]]
        identity = receipt["identities"][record["label"]]
        if identity["source"]["dirty"] is not False or identity["source"]["commit"] != variant["source_ref"]:
            raise ValueError("Unpinned or dirty measured source")
        if identity["controller_sha256"] != receipt["controller_sha256"]:
            raise ValueError("Controller identity differs")
        native = identity.get("build_metadata", {}).get("data", {}).get("native_sha256")
        if native is not None and native != identity["native"]["sha256"]:
            raise ValueError("Measured native library differs from build metadata")
    for case in receipt["workload_config"]["cases"]:
        for variant in variants.values():
            if variant["role"] in ("baseline", "final") and (case, variant["label"]) not in seen:
                raise ValueError("Missing baseline/final workload")
    hashes = {identity["harness_sha256"] for identity in receipt["identities"].values()}
    if len(hashes) != 1:
        raise ValueError("Frozen harness identities differ")


def timing(record: dict, phase: str) -> str:
    values = [sample["phase_seconds"][phase] for sample in record["raw_samples"]]
    return f"{statistics.median(values):.6f} [{min(values):.6f}, {max(values):.6f}]"


def table(headers: list[str], rows: list[list[object]]) -> list[str]:
    return ["| " + " | ".join(headers) + " |",
            "| " + " | ".join("---" for _ in headers) + " |"] + [
                "| " + " | ".join(str(value).replace("|", "\\|").replace("\n", " ") for value in row) + " |"
                for row in rows]


def render(receipt: dict, receipt_path: Path) -> str:
    validate(receipt)
    variants = {variant["role"]: variant for variant in receipt["variants"] if variant["role"] in ("baseline", "final", "control")}
    records = {(record["case"], record["label"]): record for record in receipt["results"]}
    config = receipt["workload_config"]
    lines = [f"# {receipt['label']}: matched performance evidence", "",
             f"Contract: {receipt['contract']}; {config['rounds']} measured trials and {config['warmups']} discarded warmup(s) per variant/case.",
             "Every trial uses a fresh process. Stage/control order is reproducibly shuffled and executed serially. Times below are seconds.",
             "Cumulative ratios divide matched baseline/final medians directly; isolated optimization gains are not multiplied.", "",
             f"Receipt SHA256: {hashlib.sha256(receipt_path.read_bytes()).hexdigest()}",
             f"Controller SHA256: {receipt['controller_sha256']}",
             f"Frozen harness SHA256: {next(iter(receipt['identities'].values()))['harness_sha256']}", "",
             "## Operation and total", ""]
    rows = []
    variance = []
    for case in config["cases"]:
        baseline = records[case, variants["baseline"]["label"]]
        final = records[case, variants["final"]["label"]]
        before, after = medians(baseline), medians(final)
        control = records.get((case, variants.get("control", {}).get("label")))
        rows.append([case, timing(baseline, "operation"), timing(final, "operation"),
                     f"{before['operation'] / after['operation']:.3f}×",
                     f"{medians(control)['operation'] / after['operation']:.3f}×" if control else "Not measured",
                     f"{before['total'] / after['total']:.3f}×"])
        for record in (baseline, final, control):
            if record is None:
                continue
            values = [sample["phase_seconds"]["operation"] for sample in record["raw_samples"]]
            if max(values) - min(values) > statistics.median(values) * 0.2:
                variance.append(f"{case}/{record['label']}")
    lines += table(["Workload", "Baseline operation median [min,max]", "Final operation median [min,max]",
                    "Cumulative engine gain", "Final vs openpyxl operation", "Cumulative total gain"], rows)
    lines += ["", "Operation = load + assignment + save + close for edits; load + identical public Cell/style loop + close for reads.",
              "Edit total adds fixture copy and the same independent bounded openpyxl verification. Read total equals operation.",
              "Full semantic/package validation is outside operation and total and is reported below.", ""]
    if variance:
        lines += ["**High variance:** operation range exceeds 20% of its median for " + ", ".join(variance) + ".",
                  "These five-trial ranges are descriptive, not confidence intervals.", ""]
    else:
        lines += ["Operation ranges do not exceed 20% of their medians. Five trials still do not establish a confidence interval.", ""]
    lines += ["## Phase attribution and raw ranges", ""]
    for case in config["cases"]:
        lines += [f"### {case}", ""]
        rows = []
        for variant in receipt["variants"]:
            record = records.get((case, variant["label"]))
            if record is None:
                continue
            for phase in ("copy", "load", "assignment", "save", "close", "iterate_styled_cells", "bounded_verification", "operation", "total"):
                if phase in record["phase_medians_seconds"]:
                    rows.append([variant["label"], phase, timing(record, phase)])
        lines += table(["Variant", "Phase", "Median [min,max] seconds"], rows) + [""]
    lines += ["## Independent validation", ""]
    rows = []
    for record in receipt["results"]:
        check = record["deep_validation"]
        default_scope = "full output values/types/styles/merges, ZIP/XML"
        default_scope += ", untouched-part bytes" if record["engine"] == "wolfxl" else ", package rewrite reported"
        rows.append([record["case"], record["label"], check["status"],
                     check.get("scope", default_scope),
                     f"{check.get('full_reopen_seconds', 0):.6f}", f"{check.get('package_seconds', 0):.6f}"])
    lines += table(["Workload", "Variant", "Result", "Scope", "Full reopen seconds", "Package check seconds"], rows)
    lines += ["", "Every edit trial verifies both edited values through one bounded values iterator per edited sheet.",
              "Full edit validation independently reopens the last output with eager openpyxl and checks all stored values, data types, style meanings, sheet bounds and merges.",
              "ZIP CRC/XML checks and WolfXL untouched-part byte checks run once per variant/case. Openpyxl's full rewrite is reported rather than required to retain package bytes.",
              "Styled reads match an independent openpyxl signature on every trial; their final check validates the original input package. It does not certify every reader feature.",
              "These synthetic headless checks are not Microsoft Excel certification.", "", "## Provenance", ""]
    for variant in receipt["variants"]:
        identity = receipt["identities"].get(variant["label"])
        if identity is None:
            continue
        build = identity.get("build_metadata", {})
        build_data = build.get("data", {})
        env = identity["environment"]
        lines += [f"### {variant['label']}", "",
                  f"- Measured engine: {variant['engine']}",
                  f"- Clean source HEAD: {identity['source']['commit']}",
                  f"- WolfXL version: {identity['packages']['wolfxl']['version']}",
                  f"- Imported Python source SHA256: {identity['packages']['wolfxl']['python_source_sha256']}",
                  f"- Native SHA256: {identity['native']['sha256']}",
                  f"- Python executable: {env['executable']}",
                  f"- Python: {env['python']}",
                  f"- Platform: {env['platform']}; machine: {env['machine']}; CPUs: {env['cpu_count']}",
                  f"- openpyxl: {identity['packages']['openpyxl']['version']}; source SHA256: {identity['packages']['openpyxl']['python_source_sha256']}",
                  f"- Build metadata SHA256: {build.get('sha256', 'Not provided')}",
                  f"- Rust compiler: {build_data.get('rustc', 'Not provided')}",
                  f"- Build target/features: {build_data.get('target', 'Not provided')} / {build_data.get('maturin_features', build_data.get('release_flags', 'Not provided'))}", ""]
    lines += table(["Workload", "Rows/sheet", "Populated value cells", "Max column", "Fixture SHA256"], [
        [case, metadata["rows_per_sheet"], metadata["populated_value_cells"], metadata["max_column"], metadata["sha256"]]
        for case, metadata in receipt["fixtures"].items()])
    lines += ["", "## Claim boundaries", "",
              "- This fresh-process v2 protocol differs from the historical in-process Apple benchmark and its eager verification totals.",
              "- The numeric results apply to these fixed synthetic fixtures, engine versions, source/native builds and this machine.",
              "- Plain rows contain five populated columns even though the historic nominal column setting was eight.",
              "- Bounded late-row verification still streams the XML prefix; it is not constant-time random access.",
              "- Engine gains and verification changes are separate. No overall gain is synthesized across unlike workloads.",
              "- Raw trial phases, warmups, PIDs, order, identity hashes and validation results remain in the JSON receipt.", ""]
    return "\n".join(lines)


def main(argv=None) -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("receipt", type=Path)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args(argv)
    receipt = json.loads(args.receipt.read_text())
    markdown = render(receipt, args.receipt)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(markdown)


if __name__ == "__main__":
    main()
