"""Paired, source-matched native patch-only diagnostic; ZIP IO is outside timing.

python benchmarks/bench_sheet_splices.py --baseline-root /tmp/baseline \
    --candidate-root . --fixture /tmp/plain-200000.xlsx --rows 200000 \
    --target-dir /tmp/patch-profile-target --output /tmp/patch-profile.json

Requires Cargo and the repository's Rust toolchain. The producer compiles both
worksheet modules into one release executable, alternates variant order, and
records every sample. Complete load/save/verification timings use the separately
labelled performance_contract.py harness instead.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import platform
import shutil
import statistics
import subprocess
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from zipfile import ZipFile


def digest(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def source_info(root: Path, module: Path) -> dict:
    return {
        "root": str(root),
        "commit": subprocess.check_output(["git", "-C", str(root), "rev-parse", "HEAD"], text=True).strip(),
        "dirty": bool(subprocess.check_output(["git", "-C", str(root), "status", "--porcelain"], text=True).strip()),
        "module": str(module.relative_to(root)),
        "module_sha256": digest(module.read_bytes()),
        "cargo_lock_sha256": digest((root / "Cargo.lock").read_bytes()),
    }


RUST_BENCH = r'''
#![allow(dead_code)]
use std::{hint::black_box, time::Instant};
__SUPPORT__
mod baseline;
mod candidate;
fn main() {
    let xml = std::fs::read_to_string("sheet.xml").unwrap();
    let cases = [("top",2), ("middle",__ROWS__/2), ("bottom",__ROWS__-1)];
    for (label,row) in cases {
        let before = [
            baseline::CellPatch { row, col:2, value:Some(baseline::CellValue::String("changed".into())), style_index:None },
            baseline::CellPatch { row:row+1, col:3, value:Some(baseline::CellValue::Number(12345.0)), style_index:None },
        ];
        let after = [
            candidate::CellPatch { row, col:2, value:Some(candidate::CellValue::String("changed".into())), style_index:None },
            candidate::CellPatch { row:row+1, col:3, value:Some(candidate::CellValue::Number(12345.0)), style_index:None },
        ];
        // This diagnostic fixture has ordinary existing cells, so outputs must
        // match before measuring. Opaque-XML preservation has separate tests.
        assert_eq!(baseline::patch_worksheet(&xml, &before).unwrap(),
                   candidate::patch_worksheet(&xml, &after).unwrap());
        for sample in 0..(__ROUNDS__+__WARMUPS__) {
            let order = if sample % 2 == 0 { [false,true] } else { [true,false] };
            for updated in order {
                let started = Instant::now();
                let output = if updated {
                    candidate::patch_worksheet(black_box(&xml), black_box(&after)).unwrap()
                } else {
                    baseline::patch_worksheet(black_box(&xml), black_box(&before)).unwrap()
                };
                black_box(&output);
                let seconds = started.elapsed().as_secs_f64();
                if sample >= __WARMUPS__ {
                    let variant = if updated { "candidate" } else { "baseline" };
                    println!("{{\"position\":\"{label}\",\"variant\":\"{variant}\",\"round\":{},\"seconds\":{seconds},\"output_bytes\":{}}}", sample-__WARMUPS__, output.len());
                }
            }
        }
    }
}
'''


def run(args) -> dict:
    baseline, candidate = args.baseline_root.resolve(), args.candidate_root.resolve()
    relative = Path("src/wolfxl/sheet_patcher.rs")
    commercial = not (candidate / relative).exists()
    if commercial:
        relative = Path("crates/wolfxl-patch/src/sheet.rs")
        support = "pub use wolfxl_patch::PatchError;"
        dependencies = {
            name: {"path": str(candidate / "crates" / name)}
            for name in ("wolfxl-core", "wolfxl-package", "wolfxl-patch")
        }
    else:
        utilities = (candidate / "src/ooxml_util.rs").read_text()
        start = utilities.index("pub fn attr_value(")
        end = utilities.index("\nfn relationship_id_attr(", start)
        function = utilities[start:end]
        assert function in (baseline / "src/ooxml_util.rs").read_text()
        support = "mod ooxml_util { use quick_xml::events::BytesStart;\n" + function + "\n}"
        dependencies = {"wolfxl-writer": {"path": str(candidate / "crates/wolfxl-writer")}}
    modules = {"baseline": baseline / relative, "candidate": candidate / relative}
    sources = {name: source_info(root, modules[name])
               for name, root in (("baseline", baseline), ("candidate", candidate))}
    with ZipFile(args.fixture) as archive:
        xml = archive.read("xl/worksheets/sheet1.xml")
    assert args.rows >= 4 and args.rounds >= 1 and args.warmups >= 0
    program = RUST_BENCH.replace("__SUPPORT__", support)
    for key, value in (("ROWS", args.rows), ("ROUNDS", args.rounds), ("WARMUPS", args.warmups)):
        program = program.replace(f"__{key}__", str(value))
    with tempfile.TemporaryDirectory(prefix="wolfxl-patch-profile-") as temporary:
        directory = Path(temporary)
        (directory / "src").mkdir()
        for name, module in modules.items():
            shutil.copy2(module, directory / "src" / f"{name}.rs")
        (directory / "src/main.rs").write_text(program)
        (directory / "sheet.xml").write_bytes(xml)
        manifest = '[package]\nname="wolfxl-patch-profile"\nversion="0.1.0"\nedition="2021"\n\n[dependencies]\nquick-xml="=0.37.5"\n'
        for name, dependency in dependencies.items():
            manifest += f'{name}={{path={json.dumps(dependency["path"])}}}\n'
        manifest += '\n[profile.release]\nlto="fat"\ncodegen-units=1\n'
        (directory / "Cargo.toml").write_text(manifest)
        shutil.copy2(candidate / "Cargo.lock", directory / "Cargo.lock")
        environment = dict(os.environ, CARGO_TARGET_DIR=str(args.target_dir.resolve()))
        measured = subprocess.run(["cargo", "run", "--release", "--manifest-path", str(directory / "Cargo.toml")],
                                  cwd=directory, env=environment, text=True, stdout=subprocess.PIPE, check=True)
        samples = [json.loads(line) for line in measured.stdout.splitlines()]
        compiled_lock = digest((directory / "Cargo.lock").read_bytes())
    for name, module in modules.items():
        assert digest(module.read_bytes()) == sources[name]["module_sha256"], "Source changed during measurement"
    medians = {position: {variant: statistics.median(sample["seconds"] for sample in samples
                if sample["position"] == position and sample["variant"] == variant)
                for variant in modules} for position in ("top", "middle", "bottom")}
    return {
        "contract": "wolfxl-xml-splice-patch-v1", "created_utc": datetime.now(timezone.utc).isoformat(),
        "timing_scope": "native patch_worksheet only; XML inflate/read, Cargo compilation, validation, output deallocation, and ZIP compression excluded",
        "environment": {"platform": platform.platform(), "rustc": subprocess.check_output(["rustc", "--version"], text=True).strip()},
        "producer_sha256": digest(Path(__file__).read_bytes()), "compiled_cargo_lock_sha256": compiled_lock,
        "sources": sources,
        "fixture": {"sha256": digest(args.fixture.read_bytes()), "worksheet_sha256": digest(xml), "worksheet_bytes": len(xml), "rows": args.rows},
        "rounds": args.rounds, "warmups": args.warmups, "order": "alternates baseline/candidate each round",
        "raw_samples": samples, "medians_seconds": medians,
        "speedups": {position: values["baseline"] / values["candidate"] for position, values in medians.items()},
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    for name in ("baseline-root", "candidate-root", "fixture", "target-dir", "output"):
        parser.add_argument(f"--{name}", type=Path, required=True)
    parser.add_argument("--rows", type=int, default=200000)
    parser.add_argument("--rounds", type=int, default=7)
    parser.add_argument("--warmups", type=int, default=1)
    args = parser.parse_args()
    result = run(args)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n")
    print(json.dumps({"output": str(args.output), "speedups": result["speedups"]}))


if __name__ == "__main__":
    main()
