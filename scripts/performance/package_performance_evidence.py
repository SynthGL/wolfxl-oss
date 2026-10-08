"""Copy captured performance evidence without modifying its original bytes."""
from __future__ import annotations

import argparse
import json
from pathlib import Path
import shutil


def copy(source: Path, target: Path) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(source, target)
    if source.read_bytes() != target.read_bytes():
        raise RuntimeError(f"Evidence copy differs: {target}")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--receipts", required=True, type=Path)
    parser.add_argument("--destination", required=True, type=Path)
    parser.add_argument("--build-plan", required=True, type=Path)
    parser.add_argument("--source-bindings", required=True, type=Path)
    parser.add_argument("--intermediate", type=Path)
    args = parser.parse_args()
    destination = args.destination.resolve()
    for name in ("core", "guards"):
        source = args.receipts / f"{name}.json"
        if json.loads(source.read_text())["status"] != "complete":
            raise RuntimeError(f"Incomplete receipt: {source}")
        for suffix in ("json", "md"):
            copy(args.receipts / f"{name}.{suffix}", destination / f"{name}.{suffix}")
    plan = json.loads(args.build_plan.read_text())
    copy(Path(plan["generated_document"]), destination / "build-gates.md")
    for item in plan["raw_files_to_copy_unchanged"]:
        target = (destination / item["destination"]).resolve()
        if not target.is_relative_to(destination):
            raise RuntimeError("Build plan destination escapes evidence directory")
        copy(Path(item["source"]), target)
    copy(args.source_bindings, destination / "source-bindings.json")
    if args.intermediate:
        for name in ("core", "guards"):
            for suffix in ("json", "md"):
                copy(args.intermediate / f"{name}.{suffix}", destination / f"intermediate-{name}.{suffix}")
    print(json.dumps({"destination": str(destination), "edition": plan["edition"]}))


if __name__ == "__main__":
    main()
