#!/usr/bin/env python3
"""Inspect, recalculate, and verify existing Excel workbooks without losing parts.

Subcommands print one JSON document to stdout:

  inspect PATH             what the workbook contains and which tools would damage it
  recalc PATH [--out OUT]  recalculate formulas with WolfXL; never uses LibreOffice
  verify SOURCE OUTPUT     compare an edited file with its source; exit 1 on loss

Requires the ``wolfxl`` package for ``recalc``. ``inspect`` and ``verify`` use
only the standard library, so they can check files written by any tool.
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

# Package parts grouped by feature. A family that shrinks between source and
# output means something the user can see in Excel is gone.
PART_FAMILIES = {
    "vba_macros": re.compile(r"vbaProject"),
    "pivot_tables": re.compile(r"pivotTables/"),
    "pivot_caches": re.compile(r"pivotCache/"),
    "charts": re.compile(r"charts/chart\d"),
    "drawings": re.compile(r"drawings/drawing\d"),
    "images": re.compile(r"media/"),
    "slicers": re.compile(r"slicers/"),
    "slicer_caches": re.compile(r"slicerCaches/"),
    "timelines": re.compile(r"timelines/"),
    "timeline_caches": re.compile(r"timelineCaches/"),
    "data_model": re.compile(r"xl/model/"),
    "external_links": re.compile(r"externalLinks/"),
    "tables": re.compile(r"tables/table\d"),
    "comments": re.compile(r"comments\d|threadedComments/"),
    "controls": re.compile(r"activeX/|ctrlProps/"),
    "custom_xml": re.compile(r"^customXml/"),
    "query_tables": re.compile(r"queryTables/"),
    "connections": re.compile(r"xl/connections\.xml"),
}

# Worksheet XML markers. Counted per sheet part.
SHEET_FEATURES = {
    "data_validations": b"<dataValidation ",
    "x14_data_validations": b"<x14:dataValidation ",
    "conditional_formats": b"<conditionalFormatting",
    "x14_conditional_formats": b"<x14:conditionalFormatting",
    "sparkline_groups": b"<x14:sparklineGroup ",
    "slicer_lists": b"<x14:slicerList",
    "timeline_refs": b"<x15:timelineRef",
    "table_parts": b"<tablePart ",
    "hyperlinks": b"<hyperlink ",
    "merged_ranges": b"<mergeCell ",
    "sheet_protection": b"<sheetProtection",
    "protected_ranges": b"<protectedRange ",
    "legacy_drawings": b"<legacyDrawing ",
    "drawing_refs": b"<drawing ",
}

# Parts Excel rebuilds on its own; their absence is not a loss.
REBUILT_PARTS = {"xl/calcChain.xml"}

# What openpyxl's load/save round trip drops, from its documentation, its
# reader warnings, and measured round trips of real Excel files.
OPENPYXL_DROPS = {
    "vba_macros": "removed unless load_workbook(keep_vba=True)",
    "controls": "kept only with keep_vba=True",
    "drawings": "shapes are lost (openpyxl documentation)",
    "charts": "rewritten from openpyxl's chart model; unsupported formatting is lost",
    "slicers": "removed",
    "slicer_caches": "removed",
    "timelines": "removed",
    "timeline_caches": "removed",
    "data_model": "Power Pivot data model and the pivots built on it removed",
    "x14_data_validations": "extension data validations removed",
    "x14_conditional_formats": "extension conditional formats removed",
    "sparkline_groups": "sparklines removed",
    "slicer_lists": "removed",
    "timeline_refs": "removed",
}

# What a LibreOffice headless open/save recalculation dropped when measured,
# even on files whose previous edit had preserved every part.
LIBREOFFICE_DROPS = {
    "slicers": "removed",
    "slicer_caches": "removed",
    "timelines": "removed",
    "timeline_caches": "removed",
    "data_model": "Power Pivot data model and the pivots built on it removed",
    "controls": "ActiveX controls removed",
}

FORMULA_CELL = re.compile(rb"<c\b[^>]*>(?:(?!</c>).)*?<f\b(?:(?!</c>).)*?</c>", re.S)
EXCEL_ERRORS = (
    "#NULL!",
    "#DIV/0!",
    "#VALUE!",
    "#REF!",
    "#NAME?",
    "#NUM!",
    "#N/A",
    "#GETTING_DATA",
    "#SPILL!",
    "#CALC!",
    "#FIELD!",
    "#BLOCKED!",
    "#UNKNOWN!",
)


def _sheet_parts(z: zipfile.ZipFile) -> list[str]:
    return sorted(
        n for n in z.namelist() if re.match(r"xl/worksheets/sheet\d+\.xml$", n)
    )


def _families(names) -> dict[str, int]:
    return {k: sum(1 for n in names if p.search(n)) for k, p in PART_FAMILIES.items()}


def _sheet_features(z: zipfile.ZipFile) -> dict[str, dict[str, int]]:
    out = {}
    for name in _sheet_parts(z):
        body = z.read(name)
        counts = {k: body.count(m) for k, m in SHEET_FEATURES.items()}
        out[name] = {k: v for k, v in counts.items() if v}
    return out


def _calc_properties(z: zipfile.ZipFile) -> dict[str, str]:
    match = re.search(rb"<calcPr\b([^>]*)/?>", z.read("xl/workbook.xml"))
    if not match:
        return {}
    return {
        k.decode(): v.decode()
        for k, v in re.findall(rb'(\w+)="([^"]*)"', match.group(1))
    }


def _formula_counts(z: zipfile.ZipFile) -> dict[str, int]:
    total = uncached = 0
    for name in _sheet_parts(z):
        for cell in FORMULA_CELL.findall(z.read(name)):
            total += 1
            if b"<v>" not in cell and b"<v " not in cell:
                uncached += 1
    return {"formula_cells": total, "formula_cells_without_cached_value": uncached}


def _edition() -> dict[str, str | None]:
    try:
        import wolfxl
    except ImportError:
        return {"wolfxl": None, "edition": None}
    commercial = hasattr(wolfxl, "CalculationOptions")
    return {
        "wolfxl": getattr(wolfxl, "__version__", "unknown"),
        "edition": "commercial" if commercial else "community",
    }


def inspect(path: Path) -> dict:
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        fam = {k: v for k, v in _families(names).items() if v}
        sheets = _sheet_features(z)
        calc = _calc_properties(z)
        formulas = _formula_counts(z)
    present = set(fam)
    for counts in sheets.values():
        present.update(counts)
    openpyxl_risk = {
        k: OPENPYXL_DROPS[k] for k in sorted(present) if k in OPENPYXL_DROPS
    }
    libreoffice_risk = {
        k: LIBREOFFICE_DROPS[k] for k in sorted(present) if k in LIBREOFFICE_DROPS
    }
    return {
        "file": str(path),
        **_edition(),
        "part_families": fam,
        "sheet_features": sheets,
        "calc_properties": calc,
        **formulas,
        "openpyxl_save_would_drop": openpyxl_risk,
        "libreoffice_recalc_would_drop": libreoffice_risk,
        "safe_edit": "wolfxl.load_workbook(path, modify=True"
        + (", keep_vba=True)" if path.suffix.lower() == ".xlsm" else ")"),
    }


def _set_full_calc_on_load(path: Path) -> None:
    """Mark the workbook so Excel recalculates every formula when it opens.

    Rewrites only xl/workbook.xml; every other part is copied unchanged.
    """
    with zipfile.ZipFile(path) as src:
        xml = src.read("xl/workbook.xml")
        match = re.search(rb"<calcPr\b[^>]*?(/?)>", xml)
        if match and b'fullCalcOnLoad="1"' in match.group(0):
            return
        if match:
            tag = re.sub(rb'\sfullCalcOnLoad="[^"]*"', b"", match.group(0))
            tag = re.sub(rb"\s*(/?)>$", rb' fullCalcOnLoad="1"\1>', tag)
            xml = xml[: match.start()] + tag + xml[match.end() :]
        else:
            # calcPr follows definedNames, externalReferences, or sheets, in that preference.
            anchor = None
            for pattern in (
                rb"</definedNames>",
                rb"</externalReferences>",
                rb"</sheets>",
            ):
                anchor = re.search(pattern, xml)
                if anchor:
                    break
            if anchor is None:
                raise SystemExit("workbook.xml has no <sheets> element")
            xml = (
                xml[: anchor.end()]
                + b'<calcPr fullCalcOnLoad="1"/>'
                + xml[anchor.end() :]
            )
        tmp = Path(tempfile.mkstemp(suffix=path.suffix, dir=path.parent)[1])
        with zipfile.ZipFile(tmp, "w") as dst:
            for info in src.infolist():
                data = xml if info.filename == "xl/workbook.xml" else src.read(info)
                dst.writestr(info, data, compress_type=info.compress_type)
    tmp.replace(path)


def recalc(path: Path, out: Path | None) -> tuple[dict, int]:
    try:
        import wolfxl
    except ImportError:
        raise SystemExit("wolfxl is not installed: pip install wolfxl")
    target = out or path
    if out is not None:
        shutil.copyfile(path, out)
    edition = _edition()
    commercial = edition["edition"] == "commercial"
    keep_vba = target.suffix.lower() == ".xlsm"
    wb = wolfxl.load_workbook(str(target), modify=True, keep_vba=keep_vba)
    failure = None
    values: dict = {}
    try:
        result = wb.calculate()
        values = dict(result if isinstance(result, dict) else result.values)
        if commercial:
            wb.save(str(target))
    except Exception as exc:  # noqa: BLE001 - reported to the caller as not_calculated
        failure = f"{type(exc).__name__}: {exc}"
    finally:
        wb.close()
    errors = {
        k: v
        for k, v in values.items()
        if isinstance(v, str) and v.startswith(EXCEL_ERRORS)
    }
    not_evaluated = sorted(k for k, v in values.items() if v is None)
    report = {"file": str(target), **edition}
    if failure is None:
        report.update(
            formulas_evaluated=len(values) - len(not_evaluated),
            formula_errors=dict(sorted(errors.items())),
            not_evaluated=not_evaluated,
        )
    if commercial and failure is None:
        report["cached_values"] = "written to the file"
    else:
        # Edits made the saved cached values stale; make Excel recompute on open.
        _set_full_calc_on_load(target)
        report["cached_values"] = (
            "not written. The file is marked fullCalcOnLoad so Excel recalculates on open; "
            "until then, readers such as pandas or data_only=True see the old cached values."
        )
        if failure is None:
            report["values"] = {
                k: v for k, v in sorted(values.items()) if v is not None
            }
    if failure is not None:
        report["status"] = "not_calculated"
        report["reason"] = failure
        return report, 2
    report["status"] = "errors_found" if errors else "success"
    return report, 1 if errors else 0


def verify(source: Path, output: Path, allow: set[str]) -> tuple[dict, int]:
    with zipfile.ZipFile(source) as a, zipfile.ZipFile(output) as b:
        src_names, out_names = set(a.namelist()), set(b.namelist())
        src_fam, out_fam = _families(src_names), _families(out_names)
        src_sheets, out_sheets = _sheet_features(a), _sheet_features(b)
        unchanged = sum(1 for n in src_names & out_names if a.read(n) == b.read(n))
    missing = sorted(src_names - out_names - REBUILT_PARTS)
    families_lost = {
        k: {"before": src_fam[k], "after": out_fam[k]}
        for k in src_fam
        if out_fam[k] < src_fam[k] and k not in allow
    }
    sheet_features_lost = {}
    for sheet, counts in src_sheets.items():
        after = out_sheets.get(sheet, {})
        lost = {
            k: {"before": v, "after": after.get(k, 0)}
            for k, v in counts.items()
            if after.get(k, 0) < v and k not in allow
        }
        if lost:
            sheet_features_lost[sheet] = lost
    ok = (
        not families_lost
        and not sheet_features_lost
        and not (missing and "parts" not in allow)
    )
    return {
        "source": str(source),
        "output": str(output),
        "status": "preserved" if ok else "lost",
        "parts_missing": missing,
        "parts_unchanged": unchanged,
        "parts_in_source": len(src_names),
        "families_lost": families_lost,
        "sheet_features_lost": sheet_features_lost,
        "allowed": sorted(allow),
    }, 0 if ok else 1


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    sub = parser.add_subparsers(dest="command", required=True)
    p = sub.add_parser(
        "inspect", help="list contents and which tools would damage them"
    )
    p.add_argument("path", type=Path)
    p = sub.add_parser("recalc", help="recalculate formulas with WolfXL")
    p.add_argument("path", type=Path)
    p.add_argument("--out", type=Path, help="write to this path instead of in place")
    p = sub.add_parser("verify", help="exit 1 if OUTPUT lost anything SOURCE had")
    p.add_argument("source", type=Path)
    p.add_argument("output", type=Path)
    p.add_argument(
        "--allow",
        action="append",
        default=[],
        metavar="NAME",
        help="a family or sheet feature the user asked to remove; 'parts' allows missing parts",
    )
    args = parser.parse_args(argv)
    code = 0
    if args.command == "inspect":
        report = inspect(args.path)
    elif args.command == "recalc":
        report, code = recalc(args.path, args.out)
    else:
        report, code = verify(args.source, args.output, set(args.allow))
    json.dump(report, sys.stdout, indent=2, default=str)
    sys.stdout.write("\n")
    return code


if __name__ == "__main__":
    sys.exit(main())
