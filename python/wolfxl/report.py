"""Create a text-redacted diagnostic ZIP: ``python -m wolfxl.report FILE``.

Only package shape is retained. Text, attribute values, custom names, binary
contents, exception messages, source lines and local paths are never included.
The result is diagnostic data, not a workbook that can be loaded again.
"""

from __future__ import annotations

import argparse
from collections import Counter
from contextlib import redirect_stderr, redirect_stdout
import io
import json
from pathlib import Path
import platform
import re
import sys
import traceback
from xml.parsers import expat
from zipfile import BadZipFile, ZIP_DEFLATED, ZipFile

MAX_PARTS = 4096
MAX_PART_BYTES = 8 * 1024 * 1024
MAX_TOTAL_BYTES = 64 * 1024 * 1024
MAX_ELEMENTS = 100_000
MAX_DEPTH = 128

# Preserve only fixed OOXML vocabulary. Extension names and namespace URIs may
# contain user data, just like text and attributes, so discard those too.
_TAGS = frozenset(
    "Types Default Override Relationships Relationship workbook bookViews workbookView "
    "sheets sheet definedNames definedName calcPr workbookPr worksheet dimension "
    "sheetViews sheetView selection sheetFormatPr cols col sheetData row c v f is t "
    "sst si r rPr phoneticPr rPh mergeCells mergeCell hyperlinks hyperlink "
    "styleSheet numFmts numFmt fonts font fills fill borders border cellStyleXfs "
    "cellXfs xf alignment protection cellStyles cellStyle dxfs dxf tableStyles "
    "conditionalFormatting cfRule formula dataValidations dataValidation formula1 "
    "formula2 autoFilter filterColumn filters filter table tableColumns tableColumn "
    "tableParts tablePart drawing legacyDrawing extLst ext pageMargins pageSetup "
    "printOptions headerFooter oddHeader oddFooter evenHeader evenFooter "
    "Properties creator title subject description keywords modified created "
    "comments authors author commentList comment text".split()
)
_PART_NAME = re.compile(
    r"(?:\[Content_Types\]\.xml|_rels/\.rels|docProps/(?:app|core|custom)\.xml|"
    r"xl/(?:workbook|sharedStrings|styles)\.xml|xl/_rels/workbook\.xml\.rels|"
    r"xl/(?:worksheets/sheet|charts/chart|tables/table|theme/theme|drawings/drawing)"
    r"[0-9]{1,6}\.xml|xl/worksheets/_rels/sheet[0-9]{1,6}\.xml\.rels)\Z"
)


class _ReportLimit(ValueError):
    pass


def _error(exc: Exception) -> dict:
    """Keep stack locations, never str(exc), source snippets or locals."""
    return {
        "type": type(exc).__name__,
        "trace": [
            {"file": Path(frame.filename).name, "line": frame.lineno, "function": frame.name}
            for frame in traceback.extract_tb(exc.__traceback__)
        ],
    }


def _skeleton(stream) -> tuple[str, dict[str, int]]:
    parser = expat.ParserCreate(namespace_separator="}")
    output = io.StringIO()
    counts: Counter[str] = Counter()
    depth = 0
    elements = 0

    def start(name, attrs):
        nonlocal depth, elements
        depth += 1
        elements += 1
        if depth > MAX_DEPTH or elements > MAX_ELEMENTS:
            raise _ReportLimit("XML structure limit")
        local = name.rsplit("}", 1)[-1]
        tag = local if local in _TAGS else "element"
        counts[tag] += 1
        output.write(f"<{tag}>")

    def end(name):
        nonlocal depth
        local = name.rsplit("}", 1)[-1]
        tag = local if local in _TAGS else "element"
        output.write(f"</{tag}>")
        depth -= 1

    def reject(*args):
        raise _ReportLimit("DTD and entities are forbidden")

    parser.StartElementHandler = start
    parser.EndElementHandler = end
    parser.StartDoctypeDeclHandler = reject
    parser.EntityDeclHandler = reject
    parser.ExternalEntityRefHandler = reject
    # No character, comment, PI or namespace handlers: no source data is copied.
    size = 0
    while chunk := stream.read(65536):
        size += len(chunk)
        if size > MAX_PART_BYTES:
            raise _ReportLimit("XML part byte limit")
        parser.Parse(chunk, False)
    parser.Parse(b"", True)
    return output.getvalue(), dict(sorted(counts.items()))


def write_report(filename: str | Path, output: str | Path = "workbook-report.zip") -> Path:
    """Inspect FILE and exclusively create OUTPUT, even for malformed packages.

    Native loading is attempted only after bounded ZIP/XML inspection succeeds.
    Binary parts are inventoried but omitted. Unknown part and element names are
    anonymized; every attribute and text node is omitted, including formulas.
    """
    from wolfxl import __version__, load_workbook

    destination = Path(output)
    manifest = {
        "format": "wolfxl-redacted-report-1",
        "versions": {"wolfxl": __version__, "python": platform.python_version()},
        "redaction": "All text and attributes omitted; custom names anonymized; binaries omitted.",
        "limits": {
            "parts": MAX_PARTS,
            "part_bytes": MAX_PART_BYTES,
            "total_bytes": MAX_TOTAL_BYTES,
            "elements_per_part": MAX_ELEMENTS,
            "xml_depth": MAX_DEPTH,
        },
        "parts": [],
        "load": {"status": "not_attempted"},
    }
    skeletons = {}
    safe_to_load = True
    try:
        with ZipFile(filename) as archive:
            infos = archive.infolist()
            if len(infos) > MAX_PARTS:
                raise _ReportLimit("Package part count limit")
            total = 0
            for index, info in enumerate(infos, 1):
                is_xml = info.filename.lower().endswith((".xml", ".rels"))
                name = (
                    info.filename
                    if _PART_NAME.fullmatch(info.filename)
                    else f"part-{index:04d}.{'xml' if is_xml else 'bin'}"
                )
                part = {
                    "name": name,
                    "size": info.file_size,
                    "compressed_size": info.compress_size,
                    "elements": {},
                }
                manifest["parts"].append(part)
                total += info.file_size
                if info.file_size > MAX_PART_BYTES or total > MAX_TOTAL_BYTES:
                    safe_to_load = False
                    part["error"] = {"type": "ReportLimit", "trace": []}
                    continue
                if is_xml:
                    try:
                        with archive.open(info) as stream:
                            xml, counts = _skeleton(stream)
                        skeleton_name = f"skeleton/part-{index:04d}.xml"
                        skeletons[skeleton_name] = xml
                        part["skeleton"] = skeleton_name
                        part["elements"] = counts
                    except (expat.ExpatError, _ReportLimit, BadZipFile, RuntimeError) as exc:
                        safe_to_load = False
                        part["error"] = _error(exc)
                        if isinstance(exc, expat.ExpatError):
                            part["error"].update(line=exc.lineno, column=exc.offset)
    except (OSError, BadZipFile, _ReportLimit, RuntimeError, NotImplementedError) as exc:
        safe_to_load = False
        manifest["package_error"] = _error(exc)

    if safe_to_load:
        try:
            # A library warning or exception may echo workbook data. Discard it.
            with redirect_stdout(io.StringIO()), redirect_stderr(io.StringIO()):
                workbook = load_workbook(filename)
                workbook.close()
            manifest["load"] = {"status": "loaded"}
        except Exception as exc:
            manifest["load"] = {"status": "error", **_error(exc)}
    else:
        manifest["load"]["reason"] = "Package inspection failed or exceeded a stated limit."

    # Exclusive creation protects both an existing report and the source workbook.
    with ZipFile(destination, "x", compression=ZIP_DEFLATED) as bundle:
        bundle.writestr("report.json", json.dumps(manifest, indent=2, sort_keys=True) + "\n")
        for name, xml in skeletons.items():
            bundle.writestr(name, xml)
    return destination


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("file", type=Path, help="Workbook to diagnose without changing it")
    parser.add_argument("--output", type=Path, default=Path("workbook-report.zip"))
    args = parser.parse_args(argv)
    try:
        write_report(args.file, args.output)
    except (OSError, BadZipFile):
        print("Cannot create report. Choose a new writable --output path.", file=sys.stderr)
        return 1
    print("Redacted report written. Review the bundle before attaching it to an issue.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
