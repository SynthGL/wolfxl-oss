"""Generate ordinary large XLSX inputs and measure path-backed row reads.

Run the measurements on Linux with GNU time. Generated workbooks are not fixtures
for source control. Each case uses a fresh interpreter and consumes every value.
"""

from __future__ import annotations

import argparse
import base64
import hashlib
import json
import os
from pathlib import Path
import platform
import resource
import signal
import subprocess
import sysconfig
import time
import zipfile

ADDRESS_SPACE_LIMIT = 8 * 1024**3
CASE_TIMEOUT = 600
SOURCE_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def sha256(path: Path) -> str:
    with path.open("rb") as stream:
        return hashlib.file_digest(stream, "sha256").hexdigest()


def generate(directory: Path) -> None:
    directory.mkdir(parents=True, exist_ok=True)
    inputs = []
    for label, total_rows, sheet_count in (("50mb", 91_000, 1), ("250mb", 455_000, 3)):
        path = directory / f"{label}.xlsx"
        sheet_rows = [
            total_rows // sheet_count + (index < total_rows % sheet_count)
            for index in range(sheet_count)
        ]
        started = time.perf_counter()
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as archive:
            overrides = "".join(
                f'<Override PartName="/xl/worksheets/sheet{index + 1}.xml" '
                'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                for index in range(sheet_count)
            )
            archive.writestr(
                "[Content_Types].xml",
                '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                '<Default Extension="xml" ContentType="application/xml"/>'
                '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                + overrides + "</Types>",
            )
            archive.writestr(
                "_rels/.rels",
                f'<Relationships xmlns="{REL_NS}"><Relationship Id="rId1" '
                f'Type="{DOC_REL_NS}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
            )
            archive.writestr(
                "xl/workbook.xml",
                f'<workbook xmlns="{SOURCE_NS}" xmlns:r="{DOC_REL_NS}"><sheets>'
                + "".join(
                    f'<sheet name="Sheet{index + 1}" sheetId="{index + 1}" r:id="rId{index + 1}"/>'
                    for index in range(sheet_count)
                ) + "</sheets></workbook>",
            )
            archive.writestr(
                "xl/_rels/workbook.xml.rels",
                f'<Relationships xmlns="{REL_NS}">'
                + "".join(
                    f'<Relationship Id="rId{index + 1}" Type="{DOC_REL_NS}/worksheet" '
                    f'Target="worksheets/sheet{index + 1}.xml"/>'
                    for index in range(sheet_count)
                ) + "</Relationships>",
            )
            for sheet, count in enumerate(sheet_rows):
                with archive.open(f"xl/worksheets/sheet{sheet + 1}.xml", "w", force_zip64=True) as stream:
                    stream.write(
                        f'<worksheet xmlns="{SOURCE_NS}"><dimension ref="A1:D{count}"/><sheetData>'.encode()
                    )
                    for row in range(1, count + 1):
                        seed = f"wolfxl-large-path-v1:{sheet}:{row}:".encode()
                        first = base64.b64encode(hashlib.shake_256(seed + b"a").digest(256)).decode()
                        second = base64.b64encode(hashlib.shake_256(seed + b"c").digest(256)).decode()
                        stream.write(
                            (f'<row r="{row}"><c r="A{row}" t="inlineStr"><is><t>{first}</t></is></c>'
                             f'<c r="B{row}"><v>{row}</v></c>'
                             f'<c r="C{row}" t="inlineStr"><is><t>{second}</t></is></c>'
                             f'<c r="D{row}"><v>{row * 0.125}</v></c></row>').encode()
                        )
                    stream.write(b"</sheetData></worksheet>")
        inputs.append({
            "label": label, "path": path.name, "bytes": path.stat().st_size,
            "sha256": sha256(path), "rows": total_rows, "cells": total_rows * 4,
            "sheet_rows": sheet_rows, "generation_wall_s": time.perf_counter() - started,
        })
    manifest = {
        "recipe": "SHAKE256 + base64 unique 256-byte inline string payloads; two string and two numeric columns; DEFLATE level 6; ordinary worksheet dimensions; no padding or opaque members",
        "inputs": inputs,
    }
    (directory / "inputs.json").write_text(json.dumps(manifest, indent=2) + "\n")
    print(json.dumps(manifest, indent=2), flush=True)


def consume(path: Path, mode: str) -> None:
    resource.setrlimit(resource.RLIMIT_AS, (ADDRESS_SPACE_LIMIT, ADDRESS_SPACE_LIMIT))
    import wolfxl

    started = time.perf_counter()
    workbook = wolfxl.load_workbook(path, read_only=mode == "read_only", data_only=True)
    opened = time.perf_counter()
    digest = hashlib.sha256()
    rows = cells = 0
    for sheet in workbook.worksheets:
        digest.update(sheet.title.encode() + b"\n")
        for row in sheet.iter_rows(values_only=True):
            rows += 1
            for value in row:
                cells += 1
                digest.update(json.dumps(value, ensure_ascii=True, separators=(",", ":")).encode() + b"\n")
    workbook.close()
    print(json.dumps({
        "mode": mode, "rows": rows, "cells": cells, "value_sha256": digest.hexdigest(),
        "load_s": opened - started, "iterate_and_close_s": time.perf_counter() - opened,
        "library_s": time.perf_counter() - started, "python": platform.python_version(),
        "version": getattr(wolfxl, "__version__", None),
    }), flush=True)


def measure(args: argparse.Namespace) -> None:
    binary = Path(subprocess.check_output(
        [str(args.python), "-c", "import wolfxl._rust; print(wolfxl._rust.__file__)"], text=True
    ).strip())
    binary_identity = {
        "sha256": sha256(binary), "bytes": binary.stat().st_size,
        "mtime_ns": binary.stat().st_mtime_ns,
    }
    results = []
    args.output.parent.mkdir(parents=True, exist_ok=True)
    for label, mode in (("50mb", "read_only"), ("250mb", "read_only"), ("50mb", "full")):
        rss_path = args.output.with_suffix(f".{label}.{mode}.rss")
        command = [
            "/usr/bin/time", "-f", "%M", "-o", str(rss_path), str(args.python),
            str(Path(__file__).resolve()), "consume", str(args.directory / f"{label}.xlsx"), mode,
        ]
        started = time.perf_counter()
        with subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                              text=True, start_new_session=True) as child:
            timed_out = False
            try:
                stdout, stderr = child.communicate(timeout=CASE_TIMEOUT)
            except subprocess.TimeoutExpired:
                timed_out = True
                os.killpg(child.pid, signal.SIGKILL)
                stdout, stderr = child.communicate()
        item = {
            "input": label, "mode": mode, "returncode": child.returncode,
            "timed_out": timed_out, "stdout": stdout, "stderr": stderr,
            "wall_s": time.perf_counter() - started,
        }
        if rss_path.exists() and rss_path.read_text().strip():
            last_line = rss_path.read_text().strip().splitlines()[-1]
            if last_line.isdecimal():
                item["peak_rss_bytes"] = int(last_line) * 1024
        if child.returncode == 0:
            item["result"] = json.loads(stdout)
        results.append(item)
        payload = {
            "label": args.label, "source": args.source, "binary": binary_identity,
            "iterations": 1, "warmups": 0, "per_run_timeout_s": CASE_TIMEOUT,
            "address_space_limit_bytes": ADDRESS_SPACE_LIMIT,
            "os_page_cache": "unspecified; generation and prior runs may warm it",
            "fresh_interpreter": True, "build_mode": "release",
            "build_reuse": args.build_reuse,
            "scenario": "path input, data_only=True; iterate every sheet with iter_rows(values_only=True); full means read_only=False and does not disable the library's automatic large-sheet streaming",
            "inputs": json.loads((args.directory / "inputs.json").read_text()),
            "results": results, "harness_sha256": sha256(Path(__file__)),
            "platform": platform.platform(),
        }
        args.output.write_text(json.dumps(payload, indent=2) + "\n")
        print(json.dumps(item), flush=True)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="action", required=True)
    generate_parser = subparsers.add_parser("generate")
    generate_parser.add_argument("directory", type=Path)
    child_parser = subparsers.add_parser("consume")
    child_parser.add_argument("path", type=Path)
    child_parser.add_argument("mode", choices=("read_only", "full"))
    measure_parser = subparsers.add_parser("measure")
    measure_parser.add_argument("directory", type=Path)
    measure_parser.add_argument("--python", type=Path, required=True)
    measure_parser.add_argument("--label", required=True)
    measure_parser.add_argument("--source", required=True)
    measure_parser.add_argument("--output", type=Path, required=True)
    measure_parser.add_argument("--build-reuse", required=True)
    args = parser.parse_args()
    if args.action == "generate":
        generate(args.directory)
    elif args.action == "consume":
        consume(args.path, args.mode)
    else:
        measure(args)


if __name__ == "__main__":
    # Native extension discovery and measured children use the same interpreter
    # distribution's shared libpython directory, without assuming a home path.
    os.environ["LD_LIBRARY_PATH"] = str(sysconfig.get_config_var("LIBDIR")) + os.pathsep + os.environ.get("LD_LIBRARY_PATH", "")
    main()
