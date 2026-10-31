"""Surgical rewrites for charts that came from the source workbook."""

from __future__ import annotations

import os
import posixpath
import tempfile
import zipfile
from typing import Any
from xml.etree import ElementTree as ET

REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
DRAWING_REL = f"{REL_NS}/drawing"
CHART_REL = f"{REL_NS}/chart"
CHART_CT = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"

ET.register_namespace("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")
ET.register_namespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")
ET.register_namespace("c", "http://schemas.openxmlformats.org/drawingml/2006/chart")
ET.register_namespace("r", REL_NS)


def source_chart_refs(source: str | bytes | None, sheet_title: str) -> list[dict[str, str]]:
    """Return source chart OOXML identities for ``sheet_title`` in reader order."""
    if not source:
        return []
    from wolfxl._streaming import _open_source_zip

    try:
        with _open_source_zip(source) as zf:
            sheet_path = _sheet_path_for_title(zf, sheet_title)
            if not sheet_path:
                return []
            return _chart_refs_for_sheet(zf, sheet_path)
    except (OSError, KeyError, zipfile.BadZipFile, ET.ParseError):
        return []


def apply_source_chart_authoring_to_xlsx(path: str, ops: list[dict[str, Any]]) -> None:
    """Apply source-chart remove/replace operations to ``path`` in-place."""
    if not ops:
        return

    generated: dict[str, bytes] = {}
    deletes: set[str] = set()
    remove_chart_paths: set[str] = set()

    with zipfile.ZipFile(path, "r") as src:
        content_types = ET.fromstring(src.read("[Content_Types].xml"))

        for op in ops:
            meta = op["meta"]
            chart_path = meta["chart_path"]
            if op["op"] in {"replace", "title"}:
                generated[chart_path] = op["chart_xml"]
                continue

            if op["op"] != "remove":
                continue

            drawing_path = meta["drawing_path"]
            drawing_rels_path = meta["drawing_rels_path"]
            chart_rid = meta["chart_rid"]

            drawing_xml = generated.get(drawing_path) or src.read(drawing_path)
            generated[drawing_path] = _remove_chart_anchor(drawing_xml, chart_rid)

            drawing_rels = generated.get(drawing_rels_path) or src.read(drawing_rels_path)
            generated[drawing_rels_path] = _remove_rel_by_id(drawing_rels, chart_rid)

            deletes.add(chart_path)
            remove_chart_paths.add(chart_path)

        if remove_chart_paths:
            _remove_content_type_overrides(content_types, remove_chart_paths)
            generated["[Content_Types].xml"] = ET.tostring(
                content_types, encoding="utf-8", xml_declaration=True
            )

    fd, tmp_name = tempfile.mkstemp(prefix="wolfxl-source-charts-", suffix=".xlsx")
    os.close(fd)
    try:
        with zipfile.ZipFile(path, "r") as src, zipfile.ZipFile(
            tmp_name, "w", zipfile.ZIP_DEFLATED
        ) as dst:
            for info in src.infolist():
                name = info.filename
                if name in deletes or name in generated:
                    continue
                with src.open(info, "r") as handle:
                    dst.writestr(info, handle.read())
            for name in sorted(generated):
                dst.writestr(name, generated[name])
        os.replace(tmp_name, path)
    finally:
        if os.path.exists(tmp_name):
            os.unlink(tmp_name)


def _sheet_path_for_title(zf: zipfile.ZipFile, sheet_title: str) -> str | None:
    workbook = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = _parse_rels(zf.read("xl/_rels/workbook.xml.rels"))
    for sheet in workbook.iter():
        if _local(sheet.tag) != "sheet" or sheet.attrib.get("name") != sheet_title:
            continue
        rid = _rel_id(sheet.attrib)
        if not rid or rid not in rels:
            return None
        return _resolve_target("xl/workbook.xml", rels[rid]["Target"])
    return None


def _chart_refs_for_sheet(zf: zipfile.ZipFile, sheet_path: str) -> list[dict[str, str]]:
    rels_path = _rels_path_for(sheet_path)
    try:
        sheet_rels = _parse_rels(zf.read(rels_path))
    except KeyError:
        return []

    refs: list[dict[str, str]] = []
    for rel in sheet_rels.values():
        if rel.get("Type") != DRAWING_REL:
            continue
        drawing_path = _resolve_target(sheet_path, rel["Target"])
        drawing_rels_path = _rels_path_for(drawing_path)
        try:
            drawing_xml = zf.read(drawing_path)
            drawing_rels = _parse_rels(zf.read(drawing_rels_path))
        except KeyError:
            continue
        chart_rids = _chart_rids_in_drawing_order(drawing_xml)
        for chart_rid in chart_rids:
            chart_rel = drawing_rels.get(chart_rid)
            if not chart_rel or chart_rel.get("Type") != CHART_REL:
                continue
            refs.append(
                {
                    "sheet_path": sheet_path,
                    "drawing_path": drawing_path,
                    "drawing_rels_path": drawing_rels_path,
                    "chart_path": _resolve_target(drawing_path, chart_rel["Target"]),
                    "chart_rid": chart_rid,
                }
            )
    return refs


def _chart_rids_in_drawing_order(drawing_xml: bytes) -> list[str]:
    root = ET.fromstring(drawing_xml)
    rids: list[str] = []
    for anchor in list(root):
        if _local(anchor.tag) not in {"oneCellAnchor", "twoCellAnchor", "absoluteAnchor"}:
            continue
        for child in anchor.iter():
            if _local(child.tag) == "chart":
                rid = _rel_id(child.attrib)
                if rid:
                    rids.append(rid)
                break
    return rids


def _remove_chart_anchor(drawing_xml: bytes, chart_rid: str) -> bytes:
    root = ET.fromstring(drawing_xml)
    for anchor in list(root):
        if _local(anchor.tag) not in {"oneCellAnchor", "twoCellAnchor", "absoluteAnchor"}:
            continue
        if any(_local(node.tag) == "chart" and _rel_id(node.attrib) == chart_rid for node in anchor.iter()):
            root.remove(anchor)
            break
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _remove_rel_by_id(rels_xml: bytes, rid: str) -> bytes:
    root = ET.fromstring(rels_xml)
    for rel in list(root):
        if _local(rel.tag) == "Relationship" and rel.attrib.get("Id") == rid:
            root.remove(rel)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _remove_content_type_overrides(root: ET.Element, chart_paths: set[str]) -> None:
    part_names = {f"/{path}" for path in chart_paths}
    for child in list(root):
        if (
            _local(child.tag) == "Override"
            and child.attrib.get("PartName") in part_names
            and child.attrib.get("ContentType") == CHART_CT
        ):
            root.remove(child)


def _parse_rels(xml: bytes) -> dict[str, dict[str, str]]:
    root = ET.fromstring(xml)
    out: dict[str, dict[str, str]] = {}
    for rel in root:
        if _local(rel.tag) == "Relationship" and rel.attrib.get("Id"):
            out[str(rel.attrib["Id"])] = dict(rel.attrib)
    return out


def _rel_id(attrs: dict[str, str]) -> str | None:
    for key, value in attrs.items():
        if key == f"{{{REL_NS}}}id" or key.endswith("}id") or key == "r:id":
            return value
    return None


def _rels_path_for(part_path: str) -> str:
    parent, name = posixpath.split(part_path)
    return posixpath.join(parent, "_rels", f"{name}.rels")


def _resolve_target(base_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(base_part), target))


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]
