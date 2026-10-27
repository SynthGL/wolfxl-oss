#!/usr/bin/env python3
"""Audit OOXML package fidelity between two workbook files.

This is intentionally package-level rather than API-level. It catches the
class of modify-save regressions where a workbook still opens, but an OOXML
dependency has been dropped, orphaned, or left pointing at a missing part.

The comparison is engine-neutral: it reads two zip archives and compares them.
It never opens a workbook through any spreadsheet library, so it can grade any
engine without favouring one. It imports only the Python standard library.

Public API:
    snapshot(path) -> Snapshot        Structural census of one workbook.
    audit(before, after) -> dict      Typed issue list for a save round trip.

CLI:
    python3 ooxml_fidelity.py before.xlsx after.xlsx [--json]
    Exit status is 1 when the after-file lost or broke something.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import posixpath
import re
import sys
import time
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any, Iterable, TypedDict, cast
from xml.etree import ElementTree

REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"
CT_NS = "{http://schemas.openxmlformats.org/package/2006/content-types}"
MAIN_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

FEATURE_PART_PREFIXES = {
    "calc_chain": ("xl/calcChain.xml",),
    "chart": ("xl/charts/",),
    "chart_sheet": ("xl/chartsheets/",),
    "chart_style": ("xl/charts/style", "xl/charts/colors"),
    "comment": ("xl/comments", "xl/threadedComments/", "xl/persons/"),
    "conditional_formatting": ("xl/worksheets/", "xl/styles.xml"),
    "connection": ("xl/connections.xml", "xl/queryTables/"),
    "custom_property": ("xl/customProperty",),
    "custom_xml": ("customXml/", "xl/customXml/"),
    "data_model": ("xl/model/",),
    "doc_metadata": ("docMetadata/",),
    "drawing": ("xl/drawings/",),
    "embedded_object": ("xl/embeddings/", "xl/ctrlProps/", "xl/activeX/"),
    "external_link": ("xl/externalLinks/",),
    "image_media": ("xl/media/",),
    "javascript_project": ("xl/jsaProject.bin",),
    "named_sheet_view": ("xl/namedSheetViews/",),
    "pivot": ("xl/pivotCache/", "xl/pivotTables/", "pivotCache/"),
    "printer_settings": ("xl/printerSettings/",),
    "python": ("xl/python.xml",),
    "sheet_metadata": ("xl/metadata.xml",),
    "slicer": ("xl/slicers/", "xl/slicerCaches/"),
    "table": ("xl/tables/",),
    "timeline": ("xl/timelines/", "xl/timelineCaches/"),
    "vba": ("xl/vbaProject.bin",),
}

CF_EXTENSION_NAMES = frozenset(
    {
        "conditionalFormatting",
        "conditionalFormattings",
        "cfRule",
        "colorScale",
        "dataBar",
        "iconSet",
        "pivotAreas",
    }
)

SLICER_EXTENSION_NAMES = frozenset({"slicerCaches", "slicerList"})
TIMELINE_EXTENSION_NAMES = frozenset({"timelineCacheRefs", "timelineRefs", "timelineList"})
CONDITIONAL_FORMATTING_MARKERS = (b"conditionalFormatting", b"cfRule", b"extLst")
DATA_VALIDATION_MARKERS = (b"dataValidation",)
EXTENSION_MARKERS = (b"<ext", b":ext", b"extLst")
FORMULA_MARKERS = (b"<f", b":f")
PAGE_SETUP_MARKERS = (b"pageMargins", b"pageSetup", b"printOptions", b"headerFooter")
MARKER_SENSITIVE_SEMANTIC_FEATURES = frozenset(
    {
        "conditional_formatting",
        "data_validations",
        "external_links",
        "worksheet_formulas",
    }
)


class FingerprintStats(TypedDict):
    dicts: int
    lists: int
    tuples: int
    scalars: int
    top_level_keys: list[str]


class _OptionalAuditIssueFields(TypedDict, total=False):
    before_empty: bool
    after_empty: bool
    before_summary: FingerprintStats
    after_summary: FingerprintStats


class AuditIssue(_OptionalAuditIssueFields):
    severity: str
    kind: str
    part: str
    message: str


@dataclass(frozen=True)
class Relationship:
    rels_part: str
    rel_id: str
    rel_type: str
    target: str
    target_mode: str | None
    resolved_target: str | None


@dataclass
class Snapshot:
    path: str
    parts: set[str]
    xml_parse_errors: list[tuple[str, str]]
    content_overrides: dict[str, str]
    relationships: list[Relationship]
    dxfs_count: int
    cf_dxf_refs: list[tuple[str, int]]
    table_integrity_issues: list[AuditIssue]
    worksheet_sheet_ref_issues: list[AuditIssue]
    workbook_sheet_ref_issues: list[AuditIssue]
    chart_sheet_ref_issues: list[AuditIssue]
    pivot_sheet_ref_issues: list[AuditIssue]
    feature_parts: dict[str, list[str]]
    semantic_fingerprints: dict[str, dict[str, object]]


def snapshot(
    path: Path,
    *,
    compact_semantic_drift: bool = False,
    progress_label: str | None = None,
    phase_prefix: str = "snapshot",
    phase_timings: dict[str, float] | None = None,
) -> Snapshot:
    with zipfile.ZipFile(path) as archive:
        parts = set(archive.namelist())
        _emit_audit_progress(progress_label, phase_prefix, "parts", "done")
        return Snapshot(
            path=str(path),
            parts=parts,
            xml_parse_errors=_run_audit_phase(
                progress_label,
                phase_timings,
                f"{phase_prefix}.xml_parse_errors",
                lambda: _read_xml_parse_errors(archive),
            ),
            content_overrides=_run_audit_phase(
                progress_label,
                phase_timings,
                f"{phase_prefix}.content_overrides",
                lambda: _read_content_overrides(archive),
            ),
            relationships=_run_audit_phase(
                progress_label,
                phase_timings,
                f"{phase_prefix}.relationships",
                lambda: _read_relationships(archive),
            ),
            dxfs_count=_run_audit_phase(
                progress_label,
                phase_timings,
                f"{phase_prefix}.dxfs_count",
                lambda: _read_dxfs_count(archive),
            ),
            cf_dxf_refs=_run_audit_phase(
                progress_label,
                phase_timings,
                f"{phase_prefix}.cf_dxf_refs",
                lambda: _read_cf_dxf_refs(archive),
            ),
            table_integrity_issues=_run_audit_phase(
                progress_label,
                phase_timings,
                f"{phase_prefix}.table_integrity",
                lambda: _read_table_integrity_issues(archive),
            ),
            worksheet_sheet_ref_issues=_run_audit_phase(
                progress_label,
                phase_timings,
                f"{phase_prefix}.worksheet_sheet_refs",
                lambda: _read_worksheet_sheet_ref_issues(archive, parts),
            ),
            workbook_sheet_ref_issues=_run_audit_phase(
                progress_label,
                phase_timings,
                f"{phase_prefix}.workbook_sheet_refs",
                lambda: _read_workbook_sheet_ref_issues(archive),
            ),
            chart_sheet_ref_issues=_run_audit_phase(
                progress_label,
                phase_timings,
                f"{phase_prefix}.chart_sheet_refs",
                lambda: _read_chart_sheet_ref_issues(archive),
            ),
            pivot_sheet_ref_issues=_run_audit_phase(
                progress_label,
                phase_timings,
                f"{phase_prefix}.pivot_sheet_refs",
                lambda: _read_pivot_sheet_ref_issues(archive, parts),
            ),
            feature_parts=_classify_feature_parts(parts),
            semantic_fingerprints=_run_audit_phase(
                progress_label,
                phase_timings,
                f"{phase_prefix}.semantic_fingerprints",
                lambda: _read_semantic_fingerprints(
                    archive,
                    compact=compact_semantic_drift,
                    progress_label=progress_label,
                    phase_prefix=f"{phase_prefix}.semantic_fingerprints",
                    phase_timings=phase_timings,
                ),
            ),
        )


def audit(
    before: Path,
    after: Path,
    *,
    compact_semantic_drift: bool = False,
    phase_timings: bool = False,
    progress_label: str | None = None,
) -> dict:
    timings: dict[str, float] | None = {} if phase_timings else None
    before_snapshot = _run_audit_phase(
        progress_label,
        timings,
        "before_snapshot",
        lambda: snapshot(
            before,
            compact_semantic_drift=compact_semantic_drift,
            progress_label=progress_label,
            phase_prefix="before_snapshot",
            phase_timings=timings,
        ),
    )
    after_snapshot = _run_audit_phase(
        progress_label,
        timings,
        "after_snapshot",
        lambda: snapshot(
            after,
            compact_semantic_drift=compact_semantic_drift,
            progress_label=progress_label,
            phase_prefix="after_snapshot",
            phase_timings=timings,
        ),
    )
    issues: list[AuditIssue] = []

    missing_parts = _run_audit_phase(
        progress_label,
        timings,
        "compare.missing_parts",
        lambda: sorted(before_snapshot.parts - after_snapshot.parts),
    )
    for part in missing_parts:
        issues.append(
            {
                "severity": "error",
                "kind": "missing_part",
                "part": part,
                "message": f"Part existed before save and is missing after save: {part}",
            }
        )

    _run_audit_phase(
        progress_label,
        timings,
        "compare.relationship_preservation",
        lambda: _audit_relationship_preservation(before_snapshot, after_snapshot, issues),
    )
    _run_audit_phase(
        progress_label,
        timings,
        "compare.xml_well_formed",
        lambda: _audit_xml_well_formed(before_snapshot, after_snapshot, issues),
    )
    _run_audit_phase(
        progress_label,
        timings,
        "compare.dangling_relationships",
        lambda: _audit_dangling_relationships(after_snapshot, issues),
    )
    _run_audit_phase(
        progress_label,
        timings,
        "compare.content_types",
        lambda: _audit_content_type_preservation(before_snapshot, after_snapshot, issues),
    )
    _run_audit_phase(
        progress_label,
        timings,
        "compare.conditional_formatting_refs",
        lambda: _audit_conditional_formatting_refs(after_snapshot, issues),
    )
    _run_audit_phase(
        progress_label,
        timings,
        "compare.table_integrity",
        lambda: _audit_table_integrity(before_snapshot, after_snapshot, issues),
    )
    _run_audit_phase(
        progress_label,
        timings,
        "compare.worksheet_sheet_refs",
        lambda: _audit_worksheet_sheet_refs(before_snapshot, after_snapshot, issues),
    )
    _run_audit_phase(
        progress_label,
        timings,
        "compare.workbook_sheet_refs",
        lambda: _audit_workbook_sheet_refs(before_snapshot, after_snapshot, issues),
    )
    _run_audit_phase(
        progress_label,
        timings,
        "compare.chart_sheet_refs",
        lambda: _audit_chart_sheet_refs(before_snapshot, after_snapshot, issues),
    )
    _run_audit_phase(
        progress_label,
        timings,
        "compare.pivot_sheet_refs",
        lambda: _audit_pivot_sheet_refs(before_snapshot, after_snapshot, issues),
    )
    _run_audit_phase(
        progress_label,
        timings,
        "compare.feature_hotspots",
        lambda: _audit_feature_hotspots(before_snapshot, after_snapshot, issues),
    )
    _run_audit_phase(
        progress_label,
        timings,
        "compare.semantic_fingerprints",
        lambda: _audit_semantic_fingerprints(
            before_snapshot,
            after_snapshot,
            issues,
            compact=compact_semantic_drift,
        ),
    )

    report = {
        "before": _snapshot_summary(before_snapshot),
        "after": _snapshot_summary(after_snapshot),
        "issue_count": len(issues),
        "issues": issues,
    }
    if timings is not None:
        report["phase_timings"] = timings
    return report


def _run_audit_phase(
    progress_label: str | None,
    timings: dict[str, float] | None,
    phase: str,
    func,
):
    _emit_audit_progress(progress_label, phase, "", "start")
    start = time.perf_counter()
    try:
        return func()
    finally:
        if timings is not None:
            timings[phase] = round(time.perf_counter() - start, 6)
        _emit_audit_progress(progress_label, phase, "", "done")


def _emit_audit_progress(
    progress_label: str | None,
    phase_prefix: str,
    phase: str,
    event: str,
) -> None:
    if progress_label is None:
        return
    label = phase_prefix if not phase else f"{phase_prefix}.{phase}"
    print(f"[{progress_label}] audit {label} {event}", file=sys.stderr, flush=True)


def _audit_relationship_preservation(
    before: Snapshot, after: Snapshot, issues: list[AuditIssue]
) -> None:
    before_rels = {_relationship_key(rel): rel for rel in before.relationships}
    after_rels = {_relationship_key(rel): rel for rel in after.relationships}
    for key, rel in sorted(before_rels.items()):
        if key not in after_rels:
            issues.append(
                {
                    "severity": "error",
                    "kind": "missing_relationship",
                    "part": rel.rels_part,
                    "message": (
                        "Relationship existed before save and is missing after save: "
                        f"{rel.rels_part} {rel.rel_id} {rel.rel_type} -> {rel.target}"
                    ),
                }
            )


def _audit_xml_well_formed(before: Snapshot, after: Snapshot, issues: list[AuditIssue]) -> None:
    before_errors = set(before.xml_parse_errors)
    for part, error in after.xml_parse_errors:
        if (part, error) in before_errors:
            continue
        issues.append(
            {
                "severity": "error",
                "kind": "malformed_xml_part",
                "part": part,
                "message": f"{part} is not well-formed XML after save: {error}",
            }
        )


def _audit_dangling_relationships(snapshot_: Snapshot, issues: list[AuditIssue]) -> None:
    for rel in snapshot_.relationships:
        if rel.resolved_target is None:
            continue
        if not _has_part_case_insensitive(snapshot_.parts, rel.resolved_target):
            issues.append(
                {
                    "severity": "error",
                    "kind": "dangling_relationship",
                    "part": rel.rels_part,
                    "message": (
                        f"{rel.rels_part} {rel.rel_id} points to missing {rel.resolved_target}"
                    ),
                }
            )


def _has_part_case_insensitive(parts: set[str], target: str) -> bool:
    target_lc = target.lower()
    return any(part.lower() == target_lc for part in parts)


def _audit_content_type_preservation(
    before: Snapshot, after: Snapshot, issues: list[AuditIssue]
) -> None:
    for part, content_type in sorted(before.content_overrides.items()):
        if part not in after.parts:
            continue
        after_content_type = after.content_overrides.get(part)
        if after_content_type != content_type:
            issues.append(
                {
                    "severity": "error",
                    "kind": "content_type_changed",
                    "part": part,
                    "message": (
                        f"Content type for {part} changed from {content_type!r} "
                        f"to {after_content_type!r}"
                    ),
                }
            )


def _audit_conditional_formatting_refs(snapshot_: Snapshot, issues: list[AuditIssue]) -> None:
    for sheet_part, dxf_id in snapshot_.cf_dxf_refs:
        if dxf_id >= snapshot_.dxfs_count:
            issues.append(
                {
                    "severity": "error",
                    "kind": "conditional_formatting_dxf_out_of_range",
                    "part": sheet_part,
                    "message": (
                        f"{sheet_part} references dxfId={dxf_id}, but styles.xml "
                        f"only has {snapshot_.dxfs_count} <dxf> entries"
                    ),
                }
            )


def _audit_table_integrity(before: Snapshot, after: Snapshot, issues: list[AuditIssue]) -> None:
    _extend_new_issues(before.table_integrity_issues, after.table_integrity_issues, issues)


def _audit_worksheet_sheet_refs(
    before: Snapshot, after: Snapshot, issues: list[AuditIssue]
) -> None:
    _extend_new_issues(before.worksheet_sheet_ref_issues, after.worksheet_sheet_ref_issues, issues)


def _audit_workbook_sheet_refs(before: Snapshot, after: Snapshot, issues: list[AuditIssue]) -> None:
    _extend_new_issues(before.workbook_sheet_ref_issues, after.workbook_sheet_ref_issues, issues)


def _audit_chart_sheet_refs(before: Snapshot, after: Snapshot, issues: list[AuditIssue]) -> None:
    _extend_new_issues(before.chart_sheet_ref_issues, after.chart_sheet_ref_issues, issues)


def _audit_pivot_sheet_refs(before: Snapshot, after: Snapshot, issues: list[AuditIssue]) -> None:
    _extend_new_issues(before.pivot_sheet_ref_issues, after.pivot_sheet_ref_issues, issues)


def _extend_new_issues(
    before_issues: list[AuditIssue],
    after_issues: list[AuditIssue],
    issues: list[AuditIssue],
) -> None:
    before_keys = {_issue_key(issue) for issue in before_issues}
    issues.extend(issue for issue in after_issues if _issue_key(issue) not in before_keys)


def _issue_key(issue: AuditIssue) -> tuple[str, str, str]:
    return (
        issue.get("kind", ""),
        issue.get("part", ""),
        issue.get("message", ""),
    )


def _audit_feature_hotspots(before: Snapshot, after: Snapshot, issues: list[AuditIssue]) -> None:
    for feature, before_parts in sorted(before.feature_parts.items()):
        if not before_parts:
            continue
        after_parts = set(after.feature_parts.get(feature, []))
        missing = sorted(set(before_parts) - after_parts)
        if missing:
            issues.append(
                {
                    "severity": "error",
                    "kind": "feature_part_loss",
                    "part": feature,
                    "message": f"{feature} parts disappeared after save: {missing}",
                }
            )


def _audit_semantic_fingerprints(
    before: Snapshot,
    after: Snapshot,
    issues: list[AuditIssue],
    *,
    compact: bool = False,
) -> None:
    for feature, before_fingerprint in sorted(before.semantic_fingerprints.items()):
        if not before_fingerprint:
            continue
        after_fingerprint = after.semantic_fingerprints.get(feature, {})
        if feature in {"drawing_objects", "extensions"}:
            after_fingerprint = {
                part: after_fingerprint.get(part)
                for part in before_fingerprint
                if part in after.parts
            }
        if after_fingerprint != before_fingerprint:
            issues.append(
                _semantic_drift_issue(
                    feature,
                    before_fingerprint,
                    after_fingerprint,
                    compact=compact,
                )
            )


def _semantic_drift_issue(
    feature: str,
    before_fingerprint: dict[str, object],
    after_fingerprint: dict[str, object],
    *,
    compact: bool,
) -> AuditIssue:
    if compact and feature not in MARKER_SENSITIVE_SEMANTIC_FEATURES:
        before_summary = _fingerprint_summary(before_fingerprint)
        after_summary = _fingerprint_summary(after_fingerprint)
        return {
            "severity": "error",
            "kind": f"{feature}_semantic_drift",
            "part": feature,
            "before_empty": not bool(before_fingerprint),
            "after_empty": not bool(after_fingerprint),
            "before_summary": before_summary,
            "after_summary": after_summary,
            "message": (
                f"{feature} semantic fingerprint changed after save "
                f"(compact summary): before={before_summary} "
                f"after={after_summary}"
            ),
        }
    return {
        "severity": "error",
        "kind": f"{feature}_semantic_drift",
        "part": feature,
        "before_empty": not bool(before_fingerprint),
        "after_empty": not bool(after_fingerprint),
        "message": (
            f"{feature} semantic fingerprint changed after save: "
            f"before={before_fingerprint!r} after={after_fingerprint!r}"
        ),
    }


def _fingerprint_summary(value: object) -> FingerprintStats:
    stats = FingerprintStats(
        dicts=0,
        lists=0,
        tuples=0,
        scalars=0,
        top_level_keys=[],
    )
    if isinstance(value, dict):
        stats["top_level_keys"] = sorted(str(key) for key in value)[:20]
    _accumulate_fingerprint_stats(value, stats)
    return stats


def _accumulate_fingerprint_stats(value: object, stats: FingerprintStats) -> None:
    if isinstance(value, dict):
        stats["dicts"] += 1
        for item in value.values():
            _accumulate_fingerprint_stats(item, stats)
    elif isinstance(value, list):
        stats["lists"] += 1
        for item in value:
            _accumulate_fingerprint_stats(item, stats)
    elif isinstance(value, tuple):
        stats["tuples"] += 1
        for item in value:
            _accumulate_fingerprint_stats(item, stats)
    else:
        stats["scalars"] += 1


def _read_content_overrides(archive: zipfile.ZipFile) -> dict[str, str]:
    try:
        xml = archive.read("[Content_Types].xml")
    except KeyError:
        return {}
    root = ElementTree.fromstring(xml)
    overrides: dict[str, str] = {}
    for node in root.findall(f"{CT_NS}Override"):
        part_name = node.attrib.get("PartName", "").lstrip("/")
        content_type = node.attrib.get("ContentType")
        if part_name and content_type:
            overrides[part_name] = content_type
    return overrides


def _read_content_defaults(archive: zipfile.ZipFile) -> dict[str, str]:
    try:
        xml = archive.read("[Content_Types].xml")
    except KeyError:
        return {}
    root = ElementTree.fromstring(xml)
    defaults: dict[str, str] = {}
    for node in root.findall(f"{CT_NS}Default"):
        extension = node.attrib.get("Extension")
        content_type = node.attrib.get("ContentType")
        if extension and content_type:
            defaults[extension] = content_type
    return defaults


def _read_xml_parse_errors(archive: zipfile.ZipFile) -> list[tuple[str, str]]:
    errors: list[tuple[str, str]] = []
    xml_cache = getattr(archive, "_wolfxl_ooxml_xml_cache", None)
    if xml_cache is None:
        xml_cache = {}
        setattr(archive, "_wolfxl_ooxml_xml_cache", xml_cache)
    for part in sorted(name for name in archive.namelist() if name.endswith((".xml", ".rels"))):
        if part in xml_cache:
            continue
        payload = _read_part_bytes_or_none(archive, part)
        if payload is None:
            xml_cache[part] = None
            continue
        try:
            xml_cache[part] = ElementTree.fromstring(payload)
        except ElementTree.ParseError as exc:
            xml_cache[part] = None
            errors.append((part, str(exc)))
    return errors


def _read_relationships(archive: zipfile.ZipFile) -> list[Relationship]:
    out: list[Relationship] = []
    for rels_part in sorted(p for p in archive.namelist() if p.endswith(".rels")):
        root = _read_xml_or_none(archive, rels_part)
        if root is None:
            continue
        seen_ids: set[str] = set()
        for node in root.findall(f"{REL_NS}Relationship"):
            rel_id = node.attrib.get("Id", "")
            rel_type = node.attrib.get("Type", "")
            target = node.attrib.get("Target", "")
            target_mode = node.attrib.get("TargetMode")
            resolved = _resolve_relationship_target(rels_part, target, target_mode)
            if rel_id in seen_ids:
                rel_id = f"{rel_id}#duplicate"
            seen_ids.add(rel_id)
            out.append(
                Relationship(
                    rels_part=rels_part,
                    rel_id=rel_id,
                    rel_type=rel_type,
                    target=target,
                    target_mode=target_mode,
                    resolved_target=resolved,
                )
            )
    return out


def _resolve_relationship_target(
    rels_part: str, target: str, target_mode: str | None
) -> str | None:
    if not target or target_mode == "External" or target.startswith("#"):
        return None
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))

    source_part = _source_part_for_rels(rels_part)
    source_dir = posixpath.dirname(source_part)
    return posixpath.normpath(posixpath.join(source_dir, target))


def _source_part_for_rels(rels_part: str) -> str:
    if rels_part == "_rels/.rels":
        return ""
    if "\\" in rels_part or rels_part.startswith("/") or "/_rels/" not in rels_part:
        raise ValueError(f"unsafe OOXML package part path: {rels_part}")
    prefix, name = rels_part.rsplit("/_rels/", 1)
    return posixpath.join(prefix, name.removesuffix(".rels"))


def _read_dxfs_count(archive: zipfile.ZipFile) -> int:
    root = _read_xml_or_none(archive, "xl/styles.xml")
    if root is None:
        return 0
    dxfs = root.find(f"{MAIN_NS}dxfs")
    if dxfs is None:
        return 0
    return len(dxfs.findall(f"{MAIN_NS}dxf"))


def _read_cf_dxf_refs(archive: zipfile.ZipFile) -> list[tuple[str, int]]:
    refs: list[tuple[str, int]] = []
    for part in sorted(_worksheet_parts(archive.namelist())):
        if not _part_contains_any(archive, part, CONDITIONAL_FORMATTING_MARKERS):
            continue
        root = _read_xml_or_none(archive, part)
        if root is None:
            continue
        for cf_rule in root.findall(f".//{MAIN_NS}cfRule"):
            raw = cf_rule.attrib.get("dxfId")
            if raw is not None and raw.isdigit():
                refs.append((part, int(raw)))
    return refs


def _read_table_integrity_issues(archive: zipfile.ZipFile) -> list[AuditIssue]:
    issues: list[AuditIssue] = []
    for part in sorted(_feature_xml_parts(set(archive.namelist()), "xl/tables/", ".xml")):
        root = _read_xml_or_none(archive, part)
        if root is None:
            continue
        ref = _attr(root, "ref")
        ref_width = _range_width(ref) if ref else None
        columns = _first_child_by_local(root, "tableColumns")
        child_count = len(_children_by_local(columns, "tableColumn")) if columns is not None else 0
        declared_count_raw = _attr(columns, "count") if columns is not None else None
        declared_count = (
            int(declared_count_raw)
            if declared_count_raw is not None and declared_count_raw.isdigit()
            else None
        )
        mismatches: list[str] = []
        if ref_width is not None and child_count != ref_width:
            mismatches.append(f"ref width={ref_width}, tableColumn children={child_count}")
        if declared_count is not None and declared_count != child_count:
            mismatches.append(f"count={declared_count}, tableColumn children={child_count}")
        if declared_count is not None and ref_width is not None and declared_count != ref_width:
            mismatches.append(f"count={declared_count}, ref width={ref_width}")
        if mismatches:
            mismatch_text = "; ".join(mismatches)
            issues.append(
                {
                    "severity": "error",
                    "kind": "table_column_count_mismatch",
                    "part": part,
                    "message": (
                        f"{part} table metadata is internally inconsistent: {mismatch_text}"
                    ),
                }
            )
    return issues


def _read_chart_sheet_ref_issues(
    archive: zipfile.ZipFile,
) -> list[AuditIssue]:
    sheet_names = _workbook_sheet_names(archive)
    if not sheet_names:
        return []

    parts = set(archive.namelist())
    issues: list[AuditIssue] = []
    for part in sorted(_feature_xml_parts(parts, "xl/charts/", ".xml")):
        if _is_chart_style_part(part):
            continue
        root = _read_xml_or_none(archive, part)
        if root is None:
            continue
        for formula in _texts_by_local(root, "f"):
            refs = _formula_sheet_reference_names(formula)
            missing = sorted(ref for ref in refs if ref not in sheet_names)
            if missing:
                issues.append(
                    {
                        "severity": "error",
                        "kind": "chart_formula_missing_sheet",
                        "part": part,
                        "message": (
                            f"{part} chart formula {formula!r} references missing "
                            f"sheet(s): {missing}"
                        ),
                    }
                )
        for name in _pivot_source_names(root):
            sheet_name = _pivot_source_sheet_reference_name(name)
            if sheet_name is not None and sheet_name not in sheet_names:
                issues.append(
                    {
                        "severity": "error",
                        "kind": "chart_pivot_source_missing_sheet",
                        "part": part,
                        "message": (
                            f"{part} pivot source {name!r} references missing sheet: {sheet_name!r}"
                        ),
                    }
                )
    return issues


def _read_workbook_sheet_ref_issues(
    archive: zipfile.ZipFile,
) -> list[AuditIssue]:
    sheet_names = _workbook_sheet_names(archive)
    if not sheet_names:
        return []

    root = _read_xml_or_none(archive, "xl/workbook.xml")
    if root is None:
        return []

    issues: list[AuditIssue] = []
    for node in _nodes_by_local(root, "definedName"):
        formula = _text(node)
        refs = _formula_sheet_reference_names(formula) if formula is not None else set()
        missing = sorted(ref for ref in refs if ref not in sheet_names)
        if missing:
            name = _attr(node, "name") or "<unnamed>"
            issues.append(
                {
                    "severity": "error",
                    "kind": "workbook_defined_name_missing_sheet",
                    "part": "xl/workbook.xml",
                    "message": (
                        f"defined name {name!r} formula {formula!r} references "
                        f"missing sheet(s): {missing}"
                    ),
                }
            )
    return issues


def _read_worksheet_sheet_ref_issues(archive: zipfile.ZipFile, parts: set[str]) -> list[AuditIssue]:
    sheet_names = _workbook_sheet_names(archive)
    if not sheet_names:
        return []

    issues: list[AuditIssue] = []
    for part in sorted(_worksheet_parts(parts)):
        if not _part_contains_any(archive, part, FORMULA_MARKERS):
            continue
        root = _read_xml_or_none(archive, part)
        if root is None:
            continue
        for element_name, formula in _worksheet_formula_reference_texts(root):
            refs = _formula_sheet_reference_names(formula)
            missing = sorted(ref for ref in refs if ref not in sheet_names)
            if missing:
                issues.append(
                    {
                        "severity": "error",
                        "kind": "worksheet_formula_missing_sheet",
                        "part": part,
                        "message": (
                            f"{part} {element_name} formula {formula!r} references "
                            f"missing sheet(s): {missing}"
                        ),
                    }
                )
    return issues


def _read_pivot_sheet_ref_issues(archive: zipfile.ZipFile, parts: set[str]) -> list[AuditIssue]:
    sheet_names = _workbook_sheet_names(archive)
    if not sheet_names:
        return []

    issues: list[AuditIssue] = []
    for part in sorted(_feature_xml_parts(parts, "xl/pivotCache/", ".xml")):
        root = _read_xml_or_none(archive, part)
        if root is None:
            continue
        for source in _nodes_by_local(root, "worksheetSource"):
            sheet_name = _attr(source, "sheet")
            if not sheet_name or sheet_name in sheet_names:
                continue
            issues.append(
                {
                    "severity": "error",
                    "kind": "pivot_cache_source_missing_sheet",
                    "part": part,
                    "message": (
                        f"{part} worksheetSource sheet {sheet_name!r} references a missing sheet"
                    ),
                }
            )
    return issues


def _worksheet_formula_reference_texts(
    root: ElementTree.Element,
) -> list[tuple[str, str]]:
    out: list[tuple[str, str]] = []
    formula_node_names = {"f", "formula", "formula1", "formula2"}
    for node in root.iter():
        element_name = _local_name(node.tag)
        if element_name not in formula_node_names:
            continue
        if (
            element_name in {"formula1", "formula2"}
            and _first_child_by_local(node, "f") is not None
        ):
            continue
        if formula := _text(node):
            out.append((element_name, formula))
    return out


def _workbook_sheet_names(archive: zipfile.ZipFile) -> set[str]:
    root = _read_xml_or_none(archive, "xl/workbook.xml")
    if root is None:
        return set()
    return {name for sheet in _nodes_by_local(root, "sheet") if (name := _attr(sheet, "name"))}


def _formula_sheet_reference_names(formula: str) -> set[str]:
    refs: set[str] = set()
    consumed_ranges: list[tuple[int, int]] = []

    for match in re.finditer(r'"(?:[^"]|"")*"', formula):
        consumed_ranges.append(match.span())

    for match in re.finditer(r"'((?:[^']|'')+)'!", formula):
        consumed_ranges.append(match.span())
        refs.update(_local_sheet_names_from_token(match.group(1).replace("''", "'")))

    for match in re.finditer(r"(?<![\]\w'])((?:[A-Za-z0-9_][A-Za-z0-9_ .:]*)!)", formula):
        if any(start <= match.start() < end for start, end in consumed_ranges):
            continue
        refs.update(_local_sheet_names_from_token(match.group(1)[:-1].strip()))

    return refs


def _local_sheet_names_from_token(token: str) -> set[str]:
    if "[" in token or "]" in token:
        return set()
    return {
        name for name in token.split(":") if name and not _is_formula_error_reference_token(name)
    }


def _is_formula_error_reference_token(token: str) -> bool:
    return token.lstrip("#").upper() in {
        "REF",
        "VALUE",
        "DIV/0",
        "NAME?",
        "N/A",
        "NULL",
        "NUM",
    }


def _pivot_source_sheet_reference_name(name: str) -> str | None:
    if "!" not in name:
        return None
    left, _, _ = name.partition("!")
    if "]" in left:
        left = left.rsplit("]", 1)[1]
    left = left.strip("'").replace("''", "'")
    return left or None


def _range_width(ref: str) -> int | None:
    first, _, last = ref.partition(":")
    last = last or first
    first_col = _cell_col_index(first)
    last_col = _cell_col_index(last)
    if first_col is None or last_col is None:
        return None
    return abs(last_col - first_col) + 1


def _cell_col_index(cell: str) -> int | None:
    match = re.match(r"^\$?([A-Za-z]+)\$?\d+$", cell)
    if match is None:
        return None
    value = 0
    for char in match.group(1).upper():
        value = value * 26 + (ord(char) - ord("A") + 1)
    return value


def _read_semantic_fingerprints(
    archive: zipfile.ZipFile,
    *,
    compact: bool = False,
    progress_label: str | None = None,
    phase_prefix: str = "semantic_fingerprints",
    phase_timings: dict[str, float] | None = None,
) -> dict[str, dict[str, object]]:
    parts = set(archive.namelist())
    fingerprints: dict[str, dict[str, object]] = {}

    def read_feature(name: str, func):
        value = _run_audit_phase(
            progress_label,
            phase_timings,
            f"{phase_prefix}.{name}",
            func,
        )
        fingerprints[name] = value
        return value

    read_feature("charts", lambda: _chart_fingerprint(archive, parts))
    read_feature("chart_sheets", lambda: _chart_sheet_fingerprint(archive, parts))
    read_feature("chart_styles", lambda: _chart_style_fingerprint(archive, parts))
    read_feature(
        "conditional_formatting",
        lambda: _conditional_formatting_fingerprint(archive, parts),
    )
    read_feature("connections", lambda: _connection_fingerprint(archive, parts))
    read_feature("data_model", lambda: _data_model_fingerprint(archive, parts))
    read_feature("data_validations", lambda: _data_validation_fingerprint(archive, parts))
    read_feature("drawing_objects", lambda: _drawing_object_fingerprint(archive, parts))
    read_feature("extensions", lambda: _extension_payload_fingerprint(archive, parts))
    worksheet_formulas = read_feature(
        "worksheet_formulas",
        lambda: _worksheet_formula_fingerprint(archive, parts),
    )
    read_feature(
        "external_links",
        lambda: _external_link_fingerprint(
            archive,
            parts,
            worksheet_formulas=worksheet_formulas,
        ),
    )
    read_feature(
        "named_sheet_views",
        lambda: _xml_part_fingerprint(
            archive,
            parts,
            tuple(sorted(part for part in parts if part.startswith("xl/namedSheetViews/"))),
        ),
    )
    read_feature("page_setup", lambda: _page_setup_fingerprint(archive, parts))
    read_feature("pivots", lambda: _pivot_fingerprint(archive, parts))
    read_feature("python", lambda: _xml_part_fingerprint(archive, parts, ("xl/python.xml",)))
    read_feature(
        "sheet_metadata",
        lambda: _xml_part_fingerprint(archive, parts, ("xl/metadata.xml",)),
    )
    read_feature("slicers", lambda: _slicer_fingerprint(archive, parts))
    read_feature(
        "style_theme",
        lambda: _style_theme_fingerprint(archive, parts, compact=compact),
    )
    read_feature(
        "structured_references",
        lambda: _structured_reference_fingerprint(
            archive,
            parts,
            worksheet_formulas=worksheet_formulas,
        ),
    )
    read_feature("timelines", lambda: _timeline_fingerprint(archive, parts))
    read_feature("workbook_globals", lambda: _workbook_global_fingerprint(archive, parts))
    return fingerprints


def _chart_fingerprint(archive: zipfile.ZipFile, parts: set[str]) -> dict[str, list[object]]:
    out: dict[str, list[object]] = {}
    rels_by_owner = _relationships_by_owner(archive)
    for part in sorted(_feature_xml_parts(parts, "xl/charts/", ".xml")):
        if _is_chart_style_part(part):
            continue
        root = _read_xml_or_none(archive, part)
        if root is None:
            continue
        out[part] = [
            ("formulas", _texts_by_local(root, "f")),
            ("pivot_sources", _pivot_source_names(root)),
            ("dPt_count", len(_nodes_by_local(root, "dPt"))),
            ("style_vals", _vals_by_path(root, ("style",))),
            ("chart_types", _chart_types(root)),
            ("axis_ids", _chart_axis_ids(root)),
            ("axes", _chart_axes(root)),
            ("manual_layouts", _manual_layouts(root)),
            ("series", _chart_series(root)),
            ("rels", rels_by_owner.get(part, [])),
        ]
    return out


def _chart_sheet_fingerprint(archive: zipfile.ZipFile, parts: set[str]) -> dict[str, list[object]]:
    out: dict[str, list[object]] = {}
    rels_by_owner = _relationships_by_owner(archive)
    for part in sorted(_feature_xml_parts(parts, "xl/chartsheets/", ".xml")):
        root = _read_xml_or_none(archive, part)
        if root is None:
            continue
        out[part] = [
            ("rels", rels_by_owner.get(part, [])),
            (
                "drawing_ids",
                [
                    _relationship_id(node)
                    for node in _nodes_by_local(root, "drawing")
                    + _nodes_by_local(root, "chartsheetDrawing")
                ],
            ),
            ("views", [_all_stable_attrs(node) for node in _nodes_by_local(root, "sheetView")]),
            (
                "protection",
                _stable_attrs(
                    _first_node_by_local(root, "sheetProtection"),
                    ("sheet", "objects", "scenarios"),
                ),
            ),
            ("extensions", _xml_extensions(root)),
        ]
    return out


def _chart_style_fingerprint(archive: zipfile.ZipFile, parts: set[str]) -> dict[str, object]:
    out: dict[str, object] = {}
    for part in sorted(_feature_xml_parts(parts, "xl/charts/", ".xml")):
        if not _is_chart_style_part(part):
            continue
        root = _read_xml_or_none(archive, part)
        if root is not None:
            out[part] = _xml_tree_fingerprint(root)
    return out


def _conditional_formatting_fingerprint(
    archive: zipfile.ZipFile, parts: set[str]
) -> dict[str, list[object]]:
    out: dict[str, list[object]] = {}
    for part in sorted(_worksheet_parts(parts)):
        if not _part_contains_any(archive, part, CONDITIONAL_FORMATTING_MARKERS):
            continue
        root = _read_xml_or_none(archive, part)
        if root is None:
            continue
        blocks: list[object] = []
        for block in _nodes_by_local(root, "conditionalFormatting"):
            rules: list[object] = []
            for rule in _children_by_local(block, "cfRule"):
                rules.append(
                    (
                        _stable_attrs(rule, ("type", "priority", "operator", "dxfId")),
                        _texts_by_local(rule, "formula"),
                    )
                )
            blocks.append(
                (
                    _attr(block, "sqref"),
                    rules,
                    _extension_fingerprints(block, CF_EXTENSION_NAMES),
                )
            )
        extensions = _extension_fingerprints(root, CF_EXTENSION_NAMES)
        if blocks or extensions:
            out[part] = [("blocks", blocks), ("extensions", extensions)]
    return out


def _data_validation_fingerprint(
    archive: zipfile.ZipFile, parts: set[str]
) -> dict[str, list[object]]:
    out: dict[str, list[object]] = {}
    for part in sorted(_worksheet_parts(parts)):
        if not _part_contains_any(archive, part, DATA_VALIDATION_MARKERS):
            continue
        root = _read_xml_or_none(archive, part)
        if root is None:
            continue
        validations: list[object] = []
        for validation in _nodes_by_local(root, "dataValidation"):
            validations.append(
                (
                    _stable_attrs(
                        validation,
                        (
                            "type",
                            "operator",
                            "allowBlank",
                            "showErrorMessage",
                            "showInputMessage",
                            "sqref",
                        ),
                    ),
                    _texts_by_local(validation, "formula1"),
                    _texts_by_local(validation, "formula2"),
                )
            )
        if validations:
            out[part] = validations
    return out


def _connection_fingerprint(archive: zipfile.ZipFile, parts: set[str]) -> dict[str, list[object]]:
    out: dict[str, list[object]] = {}
    rels_by_owner = _relationships_by_owner(archive)
    for part in sorted(_feature_xml_parts(parts, "xl/connections", ".xml")):
        root = _read_xml_or_none(archive, part)
        if root is None:
            continue
        connections: list[object] = []
        for connection in _nodes_by_local(root, "connection"):
            connections.append(
                (
                    _stable_attrs(
                        connection,
                        (
                            "id",
                            "name",
                            "description",
                            "type",
                            "refreshedVersion",
                            "background",
                            "saveData",
                            "deleted",
                        ),
                    ),
                    [
                        _stable_attrs(
                            node,
                            ("connection", "command", "commandType", "serverCommand"),
                        )
                        for node in _nodes_by_local(connection, "dbPr")
                    ],
                    [_xml_tree_fingerprint(node) for node in _nodes_by_local(connection, "extLst")],
                )
            )
        out[part] = [
            ("attrs", _stable_attrs(root, ("count",))),
            ("rels", rels_by_owner.get(part, [])),
            ("connections", connections),
        ]
    return out


def _data_model_fingerprint(archive: zipfile.ZipFile, parts: set[str]) -> dict[str, object]:
    model_parts = sorted(part for part in parts if part.startswith("xl/model/"))
    if not model_parts:
        return {}

    defaults = _read_content_defaults(archive)
    rels_by_owner = _relationships_by_owner(archive)
    out: dict[str, object] = {
        "xl/workbook.xml": [
            (
                "rels",
                [
                    rel
                    for rel in rels_by_owner.get("xl/workbook.xml", [])
                    if rel[1].endswith("/powerPivotData")
                    or rel[1].endswith("/model")
                    or str(rel[2]).startswith("model/")
                ],
            )
        ],
        "content_defaults": {
            ext: content_type
            for ext, content_type in sorted(defaults.items())
            if any(part.rsplit(".", 1)[-1] == ext for part in model_parts)
        },
        "parts": [
            (
                part,
                len(payload := archive.read(part)),
                hashlib.sha256(payload).hexdigest(),
            )
            for part in model_parts
        ],
    }
    return out


def _drawing_object_fingerprint(archive: zipfile.ZipFile, parts: set[str]) -> dict[str, object]:
    out: dict[str, object] = {}
    rels_by_owner = _relationships_by_owner(archive)
    for part in sorted(parts):
        if not part.startswith(
            (
                "xl/drawings/",
                "xl/comments",
                "xl/threadedComments/",
                "xl/persons/",
                "xl/media/",
                "xl/embeddings/",
                "xl/ctrlProps/",
                "xl/activeX/",
            )
        ):
            continue
        if part.endswith(".rels"):
            continue
        if part.endswith((".xml", ".vml")):
            root = _read_xml_or_none(archive, part)
            if root is None:
                continue
            out[part] = [
                ("xml", _xml_tree_fingerprint(root)),
                ("rels", rels_by_owner.get(part, [])),
            ]
        else:
            payload = archive.read(part)
            out[part] = [
                ("bytes", len(payload), hashlib.sha256(payload).hexdigest()),
                ("rels", rels_by_owner.get(part, [])),
            ]
    return out


def _external_link_fingerprint(
    archive: zipfile.ZipFile,
    parts: set[str],
    *,
    worksheet_formulas: dict[str, list[object]] | None = None,
) -> dict[str, object]:
    out: dict[str, object] = {}
    rel_targets = _rels_target_lookup(archive)
    for part in sorted(_feature_xml_parts(parts, "xl/externalLinks/", ".xml")):
        root = _read_xml_or_none(archive, part)
        if root is None:
            continue
        external_books: list[object] = []
        for book in _nodes_by_local(root, "externalBook"):
            rid = _relationship_id(book)
            external_books.append(
                (
                    rid,
                    rel_targets.get((part, rid)) if rid is not None else None,
                    _external_sheet_names(book),
                    _defined_name_refs(book),
                    _external_sheet_data(book),
                )
            )
        out[part] = external_books
    if worksheet_formulas is None:
        workbook_formulas = _worksheet_formulas(archive, parts, external_only=True)
    else:
        workbook_formulas = _worksheet_formula_texts(
            worksheet_formulas,
            external_only=True,
        )
    if workbook_formulas:
        out["worksheet_formulas"] = workbook_formulas
    return out


def _extension_payload_fingerprint(
    archive: zipfile.ZipFile, parts: set[str]
) -> dict[str, list[object]]:
    out: dict[str, list[object]] = {}
    for part in sorted(p for p in parts if p.endswith(".xml")):
        if part == "xl/workbook.xml":
            continue
        if not _part_contains_any(archive, part, EXTENSION_MARKERS):
            continue
        root = _read_xml_or_none(archive, part)
        if root is None:
            continue
        extensions = _xml_extensions(root)
        if extensions:
            out[part] = extensions
    return out


def _xml_part_fingerprint(
    archive: zipfile.ZipFile, parts: set[str], part_names: tuple[str, ...]
) -> dict[str, object]:
    rels_by_owner = _relationships_by_owner(archive)
    out: dict[str, object] = {}
    for part in part_names:
        if part not in parts:
            continue
        root = _read_xml_or_none(archive, part)
        if root is None:
            continue
        out[part] = [
            ("xml", _xml_tree_fingerprint(root)),
            ("rels", rels_by_owner.get(part, [])),
        ]
    return out


def _style_theme_fingerprint(
    archive: zipfile.ZipFile, parts: set[str], *, compact: bool = False
) -> dict[str, object]:
    out: dict[str, object] = {}
    rels_by_owner = _relationships_by_owner(archive)
    for part in sorted(p for p in parts if p == "xl/styles.xml" or p.startswith("xl/theme/theme")):
        if compact:
            payload = archive.read(part)
            out[part] = [
                ("bytes", len(payload), hashlib.sha256(payload).hexdigest()),
                ("rels", rels_by_owner.get(part, [])),
            ]
            continue
        root = _read_xml_or_none(archive, part)
        if root is None:
            continue
        out[part] = [
            ("xml", _xml_tree_fingerprint(root)),
            ("rels", rels_by_owner.get(part, [])),
        ]
    return out


def _pivot_fingerprint(archive: zipfile.ZipFile, parts: set[str]) -> dict[str, list[object]]:
    out: dict[str, list[object]] = {}
    rels_by_owner = _relationships_by_owner(archive)
    for part in sorted(_feature_xml_parts(parts, "xl/pivotTables/", ".xml")):
        root = _read_xml_or_none(archive, part)
        if root is None:
            continue
        out[part] = [
            ("attrs", _stable_attrs(root, ("name", "cacheId", "dataOnRows"))),
            ("rels", rels_by_owner.get(part, [])),
            ("data_fields", _pivot_data_fields(root)),
            ("row_fields", _pivot_field_indices(root, "rowFields", "field")),
            ("col_fields", _pivot_field_indices(root, "colFields", "field")),
            ("page_fields", _pivot_field_indices(root, "pageFields", "pageField")),
            ("calculated_items", _pivot_calculated_items(root)),
            ("formats", _pivot_formats(root)),
            ("conditional_formats", _pivot_conditional_formats(root)),
        ]
    for part in sorted(_feature_xml_parts(parts, "xl/pivotCache/", ".xml")):
        root = _read_xml_or_none(archive, part)
        if root is None:
            continue
        source = _first_node_by_local(root, "worksheetSource")
        out[part] = [
            ("cacheSource", _stable_attrs(source, ("ref", "sheet", "name"))),
            ("refreshOnLoad", _attr(root, "refreshOnLoad")),
            ("rels", rels_by_owner.get(part, [])),
            ("fields", _pivot_cache_fields(root)),
            ("calculated_fields", _pivot_calculated_fields(root)),
            ("calculated_items", _pivot_calculated_items(root)),
            ("field_groups", _pivot_field_groups(root)),
        ]
    return out


def _slicer_fingerprint(archive: zipfile.ZipFile, parts: set[str]) -> dict[str, list[object]]:
    out: dict[str, list[object]] = {}
    rels_by_owner = _relationships_by_owner(archive)
    workbook_root = _read_xml_or_none(archive, "xl/workbook.xml")
    if workbook_root is not None:
        extensions = _extension_fingerprints(workbook_root, SLICER_EXTENSION_NAMES)
        if extensions:
            out["xl/workbook.xml"] = [("extensions", extensions)]
    for part in sorted(_worksheet_parts(parts)):
        if not _part_contains_any(archive, part, EXTENSION_MARKERS):
            continue
        root = _read_xml_or_none(archive, part)
        if root is None:
            continue
        extensions = _extension_fingerprints(root, SLICER_EXTENSION_NAMES)
        if extensions:
            out[part] = [("extensions", extensions)]
    for part in sorted(_feature_xml_parts(parts, "xl/slicerCaches/", ".xml")):
        root = _read_xml_or_none(archive, part)
        if root is None:
            continue
        out[part] = [
            ("attrs", _stable_attrs(root, ("name", "pivotCacheId"))),
            ("rels", rels_by_owner.get(part, [])),
            ("data", _stable_attrs(_first_node_by_local(root, "data"), ("pivotCacheId",))),
            ("items", _slicer_items(root)),
        ]
    for part in sorted(_feature_xml_parts(parts, "xl/slicers/", ".xml")):
        root = _read_xml_or_none(archive, part)
        if root is None:
            continue
        out[part] = [
            ("attrs", _stable_attrs(root, ("name", "cache", "caption", "style"))),
            ("rels", rels_by_owner.get(part, [])),
            (
                "slicers",
                [
                    _stable_attrs(node, ("name", "cache", "caption"))
                    for node in _nodes_by_local(root, "slicer")
                ],
            ),
        ]
    return out


def _timeline_fingerprint(archive: zipfile.ZipFile, parts: set[str]) -> dict[str, list[object]]:
    out: dict[str, list[object]] = {}
    rels_by_owner = _relationships_by_owner(archive)
    workbook_root = _read_xml_or_none(archive, "xl/workbook.xml")
    if workbook_root is not None:
        extensions = _extension_fingerprints(workbook_root, TIMELINE_EXTENSION_NAMES)
        if extensions:
            out["xl/workbook.xml"] = [("extensions", extensions)]
    for part in sorted(_worksheet_parts(parts)):
        if not _part_contains_any(archive, part, EXTENSION_MARKERS):
            continue
        root = _read_xml_or_none(archive, part)
        if root is None:
            continue
        extensions = _extension_fingerprints(root, TIMELINE_EXTENSION_NAMES)
        if extensions:
            out[part] = [("extensions", extensions)]
    for prefix in ("xl/timelineCaches/", "xl/timelines/"):
        for part in sorted(_feature_xml_parts(parts, prefix, ".xml")):
            root = _read_xml_or_none(archive, part)
            if root is not None:
                out[part] = [
                    ("attrs", _stable_attrs(root, ("name", "pivotCacheId", "cache"))),
                    ("rels", rels_by_owner.get(part, [])),
                    ("xml", _xml_tree_fingerprint(root)),
                ]
    return out


def _worksheet_formula_fingerprint(
    archive: zipfile.ZipFile, parts: set[str]
) -> dict[str, list[object]]:
    out: dict[str, list[object]] = {}
    for part in sorted(_worksheet_parts(parts)):
        if not _part_contains_any(archive, part, FORMULA_MARKERS):
            continue
        root = _read_xml_or_none(archive, part)
        if root is None:
            continue
        formulas: list[object] = []
        for cell in root.iter():
            if _local_name(cell.tag) != "c":
                continue
            formula = _first_child_by_local(cell, "f")
            if formula is None:
                continue
            formulas.append(
                (
                    _stable_attrs(cell, ("r",)),
                    _stable_attrs(formula, ("t", "ref", "si", "ca", "bx")),
                    _text(formula),
                )
            )
        if formulas:
            out[part] = formulas
    return out


def _page_setup_fingerprint(archive: zipfile.ZipFile, parts: set[str]) -> dict[str, object]:
    out: dict[str, object] = {}
    for part in sorted(_worksheet_parts(parts)):
        if not _part_contains_any(archive, part, PAGE_SETUP_MARKERS):
            continue
        root = _read_xml_or_none(archive, part)
        if root is None:
            continue
        page_margins: list[tuple[tuple[str, str | None], ...]] = []
        page_setups: list[tuple[tuple[str, str | None], ...]] = []
        print_options: list[tuple[tuple[str, str | None], ...]] = []
        header_footers: list[object] = []
        for node in root.iter():
            local = _local_name(node.tag)
            if local == "pageMargins":
                page_margins.append(_all_stable_attrs(node))
            elif local == "pageSetup":
                page_setups.append(_all_stable_attrs(node))
            elif local == "printOptions":
                print_options.append(_all_stable_attrs(node))
            elif local == "headerFooter":
                header_footers.append(_xml_tree_fingerprint(node))
        entries = [
            ("page_margins", page_margins),
            ("page_setup", page_setups),
            ("print_options", print_options),
            ("header_footer", header_footers),
        ]
        entries = [(label, value) for label, value in entries if value]
        if entries:
            out[part] = entries
    return out


def _structured_reference_fingerprint(
    archive: zipfile.ZipFile,
    parts: set[str],
    *,
    worksheet_formulas: dict[str, list[object]] | None = None,
) -> dict[str, object]:
    out: dict[str, object] = {}
    formulas_by_part = worksheet_formulas
    if formulas_by_part is None:
        formulas_by_part = _worksheet_formula_fingerprint(archive, parts)
    for part, formulas in formulas_by_part.items():
        structured = [
            formula
            for formula in formulas
            if isinstance(formula, tuple)
            and len(formula) == 3
            and isinstance(formula[2], str)
            and _is_structured_reference_formula(formula[2])
        ]
        if structured:
            out[part] = structured
    return out


def _workbook_global_fingerprint(archive: zipfile.ZipFile, parts: set[str]) -> dict[str, object]:
    out: dict[str, object] = {}
    workbook_root = _read_xml_or_none(archive, "xl/workbook.xml")
    if workbook_root is not None:
        defined_names = [
            (
                _stable_attrs(
                    node,
                    ("name", "localSheetId", "hidden", "function", "vbProcedure"),
                ),
                _text(node),
            )
            for node in _nodes_by_local(workbook_root, "definedName")
        ]
        protection = _first_node_by_local(workbook_root, "workbookProtection")
        calc_pr = _first_node_by_local(workbook_root, "calcPr")
        extensions = _xml_extensions(workbook_root)
        workbook_entries = [
            ("defined_names", defined_names),
            ("workbook_protection", _all_stable_attrs(protection)),
            ("calc_pr", _all_stable_attrs(calc_pr)),
            ("extensions", extensions),
        ]
        if any(value for _, value in workbook_entries):
            out["xl/workbook.xml"] = workbook_entries
    global_parts = sorted(
        part
        for part in parts
        if part == "xl/vbaProject.bin"
        or part.startswith("customXml/")
        or part.startswith("xl/customXml/")
        or part.startswith("xl/customProperty")
        or part.startswith("xl/printerSettings/")
    )
    if global_parts:
        out["package_parts"] = global_parts
        out["package_payloads"] = {
            part: _global_package_part_fingerprint(archive, part)
            for part in global_parts
            if not part.endswith(".rels")
        }
    return out


def _global_package_part_fingerprint(archive: zipfile.ZipFile, part: str) -> object:
    if part.endswith(".xml"):
        root = _read_xml_or_none(archive, part)
        if root is not None:
            return ("xml", _xml_tree_fingerprint(root))
    payload = archive.read(part)
    return ("bytes", len(payload), hashlib.sha256(payload).hexdigest())


def _read_part_bytes_or_none(archive: zipfile.ZipFile, part: str) -> bytes | None:
    cache = getattr(archive, "_wolfxl_ooxml_bytes_cache", None)
    if cache is None:
        cache = {}
        setattr(archive, "_wolfxl_ooxml_bytes_cache", cache)
    if part in cache:
        return cache[part]
    try:
        payload = archive.read(part)
    except KeyError:
        cache[part] = None
        return None
    cache[part] = payload
    return payload


def _part_contains_any(
    archive: zipfile.ZipFile,
    part: str,
    markers: Iterable[bytes],
) -> bool:
    payload = _read_part_bytes_or_none(archive, part)
    if payload is None:
        return False
    return any(marker in payload for marker in markers)


def _read_xml_or_none(archive: zipfile.ZipFile, part: str) -> ElementTree.Element | None:
    cache = getattr(archive, "_wolfxl_ooxml_xml_cache", None)
    if cache is None:
        cache = {}
        setattr(archive, "_wolfxl_ooxml_xml_cache", cache)
    if part in cache:
        return cache[part]
    payload = _read_part_bytes_or_none(archive, part)
    if payload is None:
        cache[part] = None
        return None
    try:
        root = ElementTree.fromstring(payload)
    except ElementTree.ParseError:
        cache[part] = None
        return None
    cache[part] = root
    return root


def _feature_xml_parts(parts: set[str], prefix: str, suffix: str) -> Iterable[str]:
    return (part for part in parts if part.startswith(prefix) and part.endswith(suffix))


def _is_chart_style_part(part: str) -> bool:
    name = posixpath.basename(part)
    return name.startswith("style") or name.startswith("colors")


def _nodes_by_local(root: ElementTree.Element, name: str) -> list[ElementTree.Element]:
    return [node for node in root.iter() if _local_name(node.tag) == name]


def _first_node_by_local(root: ElementTree.Element, name: str) -> ElementTree.Element | None:
    for node in root.iter():
        if _local_name(node.tag) == name:
            return node
    return None


def _children_by_local(root: ElementTree.Element, name: str) -> list[ElementTree.Element]:
    return [node for node in root if _local_name(node.tag) == name]


def _first_child_by_local(root: ElementTree.Element, name: str) -> ElementTree.Element | None:
    for node in root:
        if _local_name(node.tag) == name:
            return node
    return None


def _texts_by_local(root: ElementTree.Element, name: str) -> list[str]:
    return [text for node in _nodes_by_local(root, name) if (text := _text(node))]


def _vals_by_path(root: ElementTree.Element, names: tuple[str, ...]) -> list[str]:
    return [
        val
        for node in root.iter()
        if _local_name(node.tag) in names and (val := _attr(node, "val")) is not None
    ]


def _pivot_source_names(root: ElementTree.Element) -> list[str]:
    names: list[str] = []
    for pivot_source in _nodes_by_local(root, "pivotSource"):
        names.extend(_texts_by_local(pivot_source, "name"))
    return names


def _chart_types(root: ElementTree.Element) -> list[str]:
    plot_area = _first_node_by_local(root, "plotArea")
    if plot_area is None:
        return []
    return [
        _local_name(node.tag) for node in list(plot_area) if _local_name(node.tag).endswith("Chart")
    ]


def _chart_axis_ids(root: ElementTree.Element) -> list[str | None]:
    return [_attr(node, "val") for node in _nodes_by_local(root, "axId")]


def _chart_axes(root: ElementTree.Element) -> list[object]:
    axes: list[object] = []
    for node in root.iter():
        local = _local_name(node.tag)
        if local not in {"catAx", "valAx", "dateAx", "serAx"}:
            continue
        axes.append(
            (
                local,
                _axis_child_val(node, "axId"),
                _axis_child_val(node, "crossAx"),
                _axis_child_val(node, "axPos"),
                _axis_child_val(node, "orientation"),
                _axis_child_val(node, "crosses"),
                _axis_child_val(node, "crossBetween"),
                _stable_attrs(_first_node_by_local(node, "numFmt"), ("formatCode", "sourceLinked")),
                _texts_by_local(node, "t"),
            )
        )
    return axes


def _manual_layouts(root: ElementTree.Element) -> list[object]:
    return [_xml_tree_fingerprint(node) for node in _nodes_by_local(root, "manualLayout")]


def _chart_series(root: ElementTree.Element) -> list[object]:
    return [
        (
            _axis_child_val(node, "idx"),
            _axis_child_val(node, "order"),
            _texts_by_local(node, "f"),
            len(_nodes_by_local(node, "dPt")),
        )
        for node in _nodes_by_local(root, "ser")
    ]


def _axis_child_val(root: ElementTree.Element, name: str) -> str | None:
    node = _first_node_by_local(root, name)
    return _attr(node, "val")


def _defined_name_refs(root: ElementTree.Element) -> list[tuple[str | None, str | None]]:
    return [
        (_attr(node, "name"), _attr(node, "refersTo"))
        for node in _nodes_by_local(root, "definedName")
    ]


def _external_sheet_names(root: ElementTree.Element) -> list[str | None]:
    return [_attr(node, "val") for node in _nodes_by_local(root, "sheetName")]


def _external_sheet_data(root: ElementTree.Element) -> list[object]:
    sheets: list[object] = []
    for sheet_data in _nodes_by_local(root, "sheetData"):
        rows: list[object] = []
        for row in _children_by_local(sheet_data, "row"):
            cells = [
                (_stable_attrs(cell, ("r", "t", "vm")), _texts_by_local(cell, "v"))
                for cell in _children_by_local(row, "cell")
            ]
            rows.append((_stable_attrs(row, ("r",)), cells))
        sheets.append(
            (
                _stable_attrs(sheet_data, ("sheetId", "refreshError")),
                rows,
            )
        )
    return sheets


def _worksheet_formulas(
    archive: zipfile.ZipFile, parts: set[str], *, external_only: bool = False
) -> dict[str, list[str]]:
    out: dict[str, list[str]] = {}
    for part in sorted(_worksheet_parts(parts)):
        if not _part_contains_any(archive, part, FORMULA_MARKERS):
            continue
        root = _read_xml_or_none(archive, part)
        if root is None:
            continue
        formulas = _texts_by_local(root, "f")
        if external_only:
            formulas = [formula for formula in formulas if _is_external_workbook_formula(formula)]
        if formulas:
            out[part] = formulas
    return out


def _worksheet_formula_texts(
    formulas_by_part: dict[str, list[object]],
    *,
    external_only: bool = False,
) -> dict[str, list[str]]:
    out: dict[str, list[str]] = {}
    for part, formulas in formulas_by_part.items():
        texts: list[str] = []
        for formula in formulas:
            if not (
                isinstance(formula, tuple) and len(formula) == 3 and isinstance(formula[2], str)
            ):
                continue
            text = formula[2]
            if external_only and not _is_external_workbook_formula(text):
                continue
            texts.append(text)
        if texts:
            out[part] = texts
    return out


def _is_external_workbook_formula(formula: str) -> bool:
    # External workbook refs look like [Book.xlsx]Sheet!A1 or
    # '[Book.xlsx]Sheet 1'!A1. Structured table refs also use brackets
    # (Table1[Column]) but do not carry a sheet bang after the closing bracket.
    return bool(re.search(r"\[[^\]]+\][^!]*!", formula))


def _is_structured_reference_formula(formula: str) -> bool:
    return "[" in formula and "]" in formula and not _is_external_workbook_formula(formula)


def _pivot_data_fields(root: ElementTree.Element) -> list[tuple[tuple[str, str | None], ...]]:
    return [
        _stable_attrs(node, ("name", "fld", "subtotal", "baseField", "baseItem"))
        for node in _nodes_by_local(root, "dataField")
    ]


def _pivot_field_indices(
    root: ElementTree.Element, container_name: str, child_name: str
) -> list[str | None]:
    container = _first_node_by_local(root, container_name)
    if container is None:
        return []
    return [
        _attr(child, "x") or _attr(child, "fld")
        for child in _children_by_local(container, child_name)
    ]


def _pivot_cache_fields(root: ElementTree.Element) -> list[tuple[tuple[str, str | None], ...]]:
    return [
        _stable_attrs(node, ("name", "numFmtId", "databaseField", "formula"))
        for node in _nodes_by_local(root, "cacheField")
    ]


def _pivot_calculated_fields(
    root: ElementTree.Element,
) -> list[tuple[tuple[str, str | None], ...]]:
    return [
        _stable_attrs(node, ("name", "formula", "hierarchy", "memberName", "mdx", "solveOrder"))
        for node in _nodes_by_local(root, "calculatedField")
    ]


def _pivot_calculated_items(root: ElementTree.Element) -> list[object]:
    return [
        (
            _stable_attrs(node, ("field", "formula")),
            _extension_fingerprints(node, CF_EXTENSION_NAMES),
        )
        for node in _nodes_by_local(root, "calculatedItem")
    ]


def _pivot_formats(root: ElementTree.Element) -> list[object]:
    return [
        (
            _stable_attrs(node, ("action", "dxfId")),
            [
                _stable_attrs(area, ("type", "field", "fieldPosition"))
                for area in _nodes_by_local(node, "pivotArea")
            ],
        )
        for node in _nodes_by_local(root, "format")
    ]


def _pivot_conditional_formats(root: ElementTree.Element) -> list[object]:
    return [
        (
            _stable_attrs(node, ("scope", "type", "priority")),
            [
                _stable_attrs(area, ("type", "field", "fieldPosition"))
                for area in _nodes_by_local(node, "pivotArea")
            ],
            _extension_fingerprints(node, CF_EXTENSION_NAMES),
        )
        for node in _nodes_by_local(root, "conditionalFormat")
    ]


def _pivot_field_groups(root: ElementTree.Element) -> list[object]:
    groups: list[object] = []
    for node in _nodes_by_local(root, "fieldGroup"):
        groups.append(_xml_tree_fingerprint(node))
    return groups


def _slicer_items(root: ElementTree.Element) -> list[tuple[tuple[str, str | None], ...]]:
    return [_stable_attrs(node, ("n", "c", "x", "s")) for node in _nodes_by_local(root, "i")]


def _extension_fingerprints(
    root: ElementTree.Element, interesting_names: frozenset[str]
) -> list[object]:
    out: list[object] = []
    for ext in _nodes_by_local(root, "ext"):
        if any(_local_name(node.tag) in interesting_names for node in ext.iter()):
            out.append(_xml_tree_fingerprint(ext))
    return out


def _xml_extensions(root: ElementTree.Element) -> list[object]:
    return [_xml_tree_fingerprint(node) for node in _nodes_by_local(root, "ext")]


def _xml_tree_fingerprint(node: ElementTree.Element) -> object:
    return (
        _local_name(node.tag),
        _all_stable_attrs(node),
        _text(node),
        [_xml_tree_fingerprint(child) for child in list(node)],
    )


def _all_stable_attrs(node: ElementTree.Element | None) -> tuple[tuple[str, str], ...]:
    if node is None:
        return tuple()
    return tuple(sorted((_local_name(key), value) for key, value in node.attrib.items()))


def _relationships_by_owner(
    archive: zipfile.ZipFile,
) -> dict[str, list[tuple[str, str, str, str | None]]]:
    lookup: dict[str, list[tuple[str, str, str, str | None]]] = {}
    for rel in _read_relationships(archive):
        owner = _source_part_for_rels(rel.rels_part)
        if owner:
            lookup.setdefault(owner, []).append(
                (rel.rel_id, rel.rel_type, rel.target, rel.target_mode)
            )
    return lookup


def _rels_target_lookup(archive: zipfile.ZipFile) -> dict[tuple[str, str], str]:
    lookup: dict[tuple[str, str], str] = {}
    for rel in _read_relationships(archive):
        owner = _source_part_for_rels(rel.rels_part)
        if owner:
            lookup[(owner, rel.rel_id)] = rel.target
    return lookup


def _relationship_id(node: ElementTree.Element) -> str | None:
    for key, value in node.attrib.items():
        if key.endswith("}id") or key == "id":
            return value
    return None


def _stable_attrs(
    node: ElementTree.Element | None, names: Iterable[str]
) -> tuple[tuple[str, str | None], ...]:
    if node is None:
        return tuple((name, None) for name in names)
    return tuple((name, _attr(node, name)) for name in names)


def _attr(node: ElementTree.Element | None, name: str) -> str | None:
    if node is None:
        return None
    for key, value in node.attrib.items():
        if _local_name(key) == name:
            return value
    return None


def _text(node: ElementTree.Element) -> str | None:
    value = (node.text or "").strip()
    return value or None


def _local_name(name: str) -> str:
    return name.rsplit("}", 1)[-1] if "}" in name else name


def _worksheet_parts(parts: Iterable[str]) -> Iterable[str]:
    pattern = re.compile(r"^xl/worksheets/sheet\d+\.xml$")
    return (part for part in parts if pattern.match(part))


def _classify_feature_parts(parts: set[str]) -> dict[str, list[str]]:
    classified: dict[str, list[str]] = {}
    for feature, prefixes in FEATURE_PART_PREFIXES.items():
        classified[feature] = sorted(
            part for part in parts if any(part.startswith(prefix) for prefix in prefixes)
        )
    return classified


def _relationship_key(rel: Relationship) -> tuple[str, str, str, str | None]:
    return (rel.rels_part, rel.rel_type, rel.target, rel.target_mode)


def _snapshot_summary(snapshot_: Snapshot) -> dict:
    return {
        "path": snapshot_.path,
        "part_count": len(snapshot_.parts),
        "xml_parse_error_count": len(snapshot_.xml_parse_errors),
        "relationship_count": len(snapshot_.relationships),
        "content_override_count": len(snapshot_.content_overrides),
        "dxfs_count": snapshot_.dxfs_count,
        "cf_dxf_ref_count": len(snapshot_.cf_dxf_refs),
        "table_integrity_issue_count": len(snapshot_.table_integrity_issues),
        "worksheet_sheet_ref_issue_count": len(snapshot_.worksheet_sheet_ref_issues),
        "workbook_sheet_ref_issue_count": len(snapshot_.workbook_sheet_ref_issues),
        "chart_sheet_ref_issue_count": len(snapshot_.chart_sheet_ref_issues),
        "pivot_sheet_ref_issue_count": len(snapshot_.pivot_sheet_ref_issues),
        "feature_part_counts": {
            feature: len(parts) for feature, parts in snapshot_.feature_parts.items()
        },
        "semantic_fingerprint_counts": {
            feature: len(fingerprint)
            for feature, fingerprint in snapshot_.semantic_fingerprints.items()
        },
    }


def _json_default(value: object) -> object:
    if isinstance(value, set):
        return sorted(value)
    if hasattr(value, "__dataclass_fields__"):
        return asdict(cast(Any, value))
    raise TypeError(f"Object of type {type(value).__name__} is not JSON serializable")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("before", type=Path, help="Workbook before modify-save")
    parser.add_argument("after", type=Path, help="Workbook after modify-save")
    parser.add_argument("--json", action="store_true", help="Emit machine-readable JSON")
    args = parser.parse_args(argv)

    report = audit(args.before, args.after)
    if args.json:
        print(json.dumps(report, indent=2, default=_json_default, sort_keys=True))
    else:
        _print_text_report(report)
    return 1 if report["issues"] else 0


def _print_text_report(report: dict) -> None:
    print(f"Before parts: {report['before']['part_count']}")
    print(f"After parts:  {report['after']['part_count']}")
    print(f"Issues:       {report['issue_count']}")
    for issue in report["issues"]:
        print(f"- [{issue['severity']}] {issue['kind']}: {issue['message']}")


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
