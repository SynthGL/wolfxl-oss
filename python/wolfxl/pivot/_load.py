"""Load-time materialisation of existing pivot tables in modify mode.

The patcher's source ZIP contains zero or more
``xl/pivotTables/pivotTable*.xml`` parts. Each is reachable from a
worksheet via the sheet's ``_rels/sheet{N}.xml.rels`` mapping (rel
type ``…/relationships/pivotTable``). The pivot table in turn
references its cache definition via its own per-table
``_rels/pivotTable{N}.xml.rels``.

This module performs the rels walk + parse on demand the first time
``Worksheet.pivot_tables`` is read, returning a list of
:class:`PivotTableHandle` instances. Construction-mode workbooks (no
source ZIP) and write-only workbooks return an empty list.

The Rust parser exposed via ``_rust.parse_pivot_table_meta`` and
``_rust.parse_pivot_cache_source`` does the heavy lifting; the
Python wrapper here is purely orchestration (rels graph walk + path
joining).
"""

from __future__ import annotations

import posixpath
import xml.etree.ElementTree as ET
import zipfile
from typing import TYPE_CHECKING

from wolfxl._zip_safety import read_entry_optional, validate_zipfile

from ._handle import PivotTableHandle

if TYPE_CHECKING:
    from wolfxl._workbook import Workbook
    from wolfxl._worksheet import Worksheet

_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_REL_PIVOT_TABLE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable"
)
_REL_PIVOT_CACHE_DEF = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition"
)
_REL_PIVOT_CACHE_RECORDS = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords"
)


def load_pivot_tables_for_sheet(
    workbook: Workbook,
    worksheet: Worksheet,
) -> list[PivotTableHandle]:
    """Walk the source ZIP and build pivot handles for one sheet.

    Returns an empty list on construction-mode workbooks (no source
    path), missing rels, malformed XML, or any other non-fatal load
    failure — accessing ``ws.pivot_tables`` must never raise on a
    well-formed workbook that simply has no pivots.
    """
    from wolfxl._streaming import _open_source_zip, stream_source

    source = stream_source(workbook)
    if not source:
        return []

    sheet_zip_path = _resolve_sheet_zip_path(workbook, worksheet)
    if sheet_zip_path is None:
        return []

    sheet_dir = posixpath.dirname(sheet_zip_path)
    sheet_name = posixpath.basename(sheet_zip_path)
    sheet_rels_path = posixpath.join(sheet_dir, "_rels", f"{sheet_name}.rels")

    try:
        from wolfxl import _rust
    except ImportError:
        return []

    parse_table = getattr(_rust, "parse_pivot_table_meta", None)
    parse_cache = getattr(_rust, "parse_pivot_cache_source", None)
    if parse_table is None or parse_cache is None:
        # Rust extension predates pivot-handle support — return empty rather than crash.
        return []

    handles: list[PivotTableHandle] = []
    try:
        with _open_source_zip(source) as zf:
            validate_zipfile(zf)
            sheet_rels_xml = _safe_read(zf, sheet_rels_path)
            if not sheet_rels_xml:
                return []
            pivot_table_targets = _rels_targets_by_type(
                sheet_rels_xml, _REL_PIVOT_TABLE
            )
            for rel_target in pivot_table_targets:
                table_zip_path = _join_zip_path(sheet_dir, rel_target)
                table_xml = _safe_read(zf, table_zip_path)
                if not table_xml:
                    continue
                try:
                    table_meta = parse_table(table_xml)
                except Exception:
                    continue
                # Resolve the cache definition path through the table's
                # own _rels.
                table_dir = posixpath.dirname(table_zip_path)
                table_name = posixpath.basename(table_zip_path)
                table_rels_path = posixpath.join(
                    table_dir, "_rels", f"{table_name}.rels"
                )
                table_rels_xml = _safe_read(zf, table_rels_path)
                if not table_rels_xml:
                    continue
                cache_def_targets = _rels_targets_by_type(
                    table_rels_xml, _REL_PIVOT_CACHE_DEF
                )
                if not cache_def_targets:
                    continue
                cache_zip_path = _join_zip_path(table_dir, cache_def_targets[0])
                cache_xml = _safe_read(zf, cache_zip_path)
                if not cache_xml:
                    continue
                try:
                    cache_meta = parse_cache(cache_xml)
                except Exception:
                    continue
                field_names = _cache_field_names(cache_xml)

                # Best-effort: resolve the records part path via the
                # cache's own _rels so the handle can carry it for
                # future record-regen RFCs. Records are not touched
                # in v1.0.
                cache_dir = posixpath.dirname(cache_zip_path)
                cache_name = posixpath.basename(cache_zip_path)
                cache_rels_path = posixpath.join(
                    cache_dir, "_rels", f"{cache_name}.rels"
                )
                cache_rels_xml = _safe_read(zf, cache_rels_path)
                records_zip_path = ""
                if cache_rels_xml:
                    records_targets = _rels_targets_by_type(
                        cache_rels_xml, _REL_PIVOT_CACHE_RECORDS
                    )
                    if records_targets:
                        records_zip_path = _join_zip_path(
                            cache_dir, records_targets[0]
                        )

                handle = PivotTableHandle(
                    workbook=workbook,
                    owner_sheet=worksheet,
                    name=str(table_meta["name"]),
                    location=str(table_meta["location_ref"]),
                    cache_id=int(table_meta["cache_id"]),
                    cache_part_path=cache_zip_path,
                    records_part_path=records_zip_path,
                    table_part_path=table_zip_path,
                    orig_source_range=str(cache_meta["range"]),
                    orig_source_sheet=str(cache_meta["sheet"]),
                    orig_field_count=int(cache_meta["field_count"]),
                    field_names=field_names,
                )
                handles.append(handle)
    except (zipfile.BadZipFile, OSError, ValueError):
        return []

    return handles


def _cache_field_names(cache_xml: bytes) -> list[str]:
    out: list[str] = []
    try:
        root = ET.fromstring(cache_xml)
    except ET.ParseError:
        return out
    for elem in root.iter():
        if elem.tag.endswith("cacheField"):
            name = elem.get("name")
            if name is not None:
                out.append(name)
    return out


def _resolve_sheet_zip_path(
    workbook: Workbook,
    worksheet: Worksheet,
) -> str | None:
    """Return the ZIP entry path for ``worksheet`` (e.g.
    ``xl/worksheets/sheet1.xml``)."""
    from wolfxl._streaming import _open_source_zip, stream_source

    source = stream_source(workbook)
    if not source:
        return None
    try:
        with _open_source_zip(source) as zf:
            validate_zipfile(zf)
            wb_rels = _safe_read(zf, "xl/_rels/workbook.xml.rels")
            wb_xml = _safe_read(zf, "xl/workbook.xml")
    except (zipfile.BadZipFile, OSError, ValueError):
        return None
    if not wb_rels or not wb_xml:
        return None
    rid_to_target = _rels_id_to_target(wb_rels)
    sheet_rids = _workbook_sheet_rids(wb_xml)
    target = rid_to_target.get(sheet_rids.get(worksheet.title, ""))
    if not target:
        return None
    return _join_zip_path("xl", target)


def _safe_read(zf: zipfile.ZipFile, path: str) -> bytes | None:
    """Read a ZIP entry, returning ``None`` on missing entries."""
    return read_entry_optional(zf, path)


def _rels_targets_by_type(rels_xml: bytes, rel_type: str) -> list[str]:
    """Return rel `Target` values whose `Type` matches."""
    out: list[str] = []
    try:
        root = ET.fromstring(rels_xml)
    except ET.ParseError:
        return out
    for child in root:
        if not child.tag.endswith("Relationship"):
            continue
        if child.get("Type") == rel_type:
            target = child.get("Target")
            if target:
                out.append(target)
    return out


def _rels_id_to_target(rels_xml: bytes) -> dict[str, str]:
    """Return a `{rId: Target}` map from a rels XML blob."""
    out: dict[str, str] = {}
    try:
        root = ET.fromstring(rels_xml)
    except ET.ParseError:
        return out
    for child in root:
        if not child.tag.endswith("Relationship"):
            continue
        rid = child.get("Id")
        target = child.get("Target")
        if rid and target:
            out[rid] = target
    return out


def _workbook_sheet_rids(wb_xml: bytes) -> dict[str, str]:
    """Return a `{sheet_title: rId}` map from a `xl/workbook.xml`."""
    out: dict[str, str] = {}
    try:
        root = ET.fromstring(wb_xml)
    except ET.ParseError:
        return out
    # workbook.xml uses the spreadsheetml namespace; <sheets><sheet
    # name="..." r:id="..."/></sheets>.
    for sheets in root:
        if not sheets.tag.endswith("sheets"):
            continue
        for sheet in sheets:
            if not sheet.tag.endswith("sheet"):
                continue
            name = sheet.get("name")
            rid = None
            for k, v in sheet.attrib.items():
                if k.endswith("}id") or k == "id":
                    rid = v
                    break
            if name and rid:
                out[name] = rid
    return out


def _join_zip_path(base_dir: str, target: str) -> str:
    """Resolve a rel `Target` against a `base_dir` using OOXML rules.

    Rels targets are usually relative; ``../charts/chart1.xml`` from
    ``xl/worksheets`` resolves to ``xl/charts/chart1.xml``. Absolute
    targets (leading ``/``) are taken as-is, minus the leading slash.
    """
    if not target:
        return base_dir
    if target.startswith("/"):
        return target.lstrip("/")
    joined = posixpath.normpath(posixpath.join(base_dir, target))
    return joined


__all__ = ["load_pivot_tables_for_sheet"]
