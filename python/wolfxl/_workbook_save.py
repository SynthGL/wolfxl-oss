"""Workbook save orchestration helpers.

This module keeps the public :class:`wolfxl.Workbook` methods thin while
preserving the exact writer/patcher flush order used by the save pipeline.
"""

from __future__ import annotations

import os
from typing import Any, BinaryIO

from wolfxl._workbook_state import same_existing_path
from wolfxl.xml.constants import (
    ARC_CONTENT_TYPES,
    ARC_CUSTOM,
    ARC_ROOT_RELS,
    CONTYPES_NS,
    CPROPS_TYPE,
    DOC_NS,
    PKG_REL_NS,
    XLSM,
    XLSX,
    XLTM,
    XLTX,
)


def normalize_openpyxl_package_shape(wb: Any, filename: str) -> None:
    """Apply source-backed openpyxl package-shape cleanup when relevant."""
    from wolfxl._openpyxl_package_shape import normalize_openpyxl_package_shape as _normalize

    keep_vba = bool(getattr(wb, "_keep_vba", False))
    if keep_vba or getattr(wb, "_rust_patcher", None) is None:
        _normalize(filename, keep_vba=keep_vba)


def save_workbook(
    wb: Any,
    filename: str | os.PathLike[str] | BinaryIO,
    *,
    password: str | bytes | None = None,
) -> None:
    """Flush workbook state and save it through the active backend."""
    if hasattr(filename, "write") and not isinstance(filename, (str, bytes, os.PathLike)):
        save_workbook_to_fileobj(wb, filename, password=password)
        return
    filename = str(filename)
    # G20: write-only mode is consumed-on-save. A second save raises
    # WorkbookAlreadySaved (matches openpyxl's `_write_only.py`).
    # Eager-mode workbooks remain re-savable.
    if getattr(wb, "_saved", False) and getattr(wb, "_write_only", False):
        from wolfxl.utils.exceptions import WorkbookAlreadySaved

        raise WorkbookAlreadySaved(
            "Workbook(write_only=True) is consumed-on-save; "
            "open a new workbook to write again"
        )

    if password is not None:
        # Validate password early so we don't write a plaintext tempfile that
        # we'd then have to throw away.
        from wolfxl._encryption import _coerce_password

        _coerce_password(password)  # raises ValueError on empty
        save_encrypted(wb, filename, password)
        return

    if wb._rust_patcher is not None:  # noqa: SLF001
        save_modify_mode(wb, filename)
    elif wb._rust_writer is not None:  # noqa: SLF001
        if getattr(wb, "_write_only", False):
            save_write_only_mode(wb, filename)
        else:
            save_write_mode(wb, filename)
    elif getattr(wb, "_rust_reader", None) is not None and getattr(wb, "_source_path", None):
        save_read_mode(wb, filename)
    else:
        raise RuntimeError("save requires write or modify mode")
    if wb._rust_writer is not None:  # noqa: SLF001
        apply_writer_sheet_deletes(wb, filename)
        apply_writer_unmerged_ranges(wb, filename)
        apply_sheet_state_authoring(wb, filename)
    apply_workbook_template_content_type(wb, filename)
    apply_custom_doc_props_authoring(wb, filename)
    # Mark consumed AFTER save succeeds so a write failure leaves the
    # workbook in a re-tryable state for the eager path. Write-only
    # mode flips this once and never re-saves.
    wb._saved = True  # noqa: SLF001
    # Mark every WriteOnlyWorksheet closed so subsequent appends raise
    # WorkbookAlreadySaved cleanly.
    if getattr(wb, "_write_only", False):
        for ws in wb._sheets.values():  # noqa: SLF001
            close = getattr(ws, "close", None)
            if close is not None:
                close()


def save_workbook_to_fileobj(
    wb: Any,
    fileobj: BinaryIO,
    *,
    password: str | bytes | None = None,
) -> None:
    """Save to a binary file-like object using the path-oriented backends."""
    import tempfile

    tmp = tempfile.NamedTemporaryFile(prefix="wolfxl-save-", suffix=".xlsx", delete=False)
    tmp_path = tmp.name
    tmp.close()
    try:
        save_workbook(wb, tmp_path, password=password)
        with open(tmp_path, "rb") as src:
            data = src.read()
        try:
            fileobj.seek(0)
            fileobj.truncate()
        except Exception:
            pass
        fileobj.write(data)
        try:
            fileobj.flush()
        except Exception:
            pass
    finally:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass


def save_read_mode(wb: Any, filename: str) -> None:
    """Save an unmodified path-backed read workbook by copying the source package."""
    import shutil

    source_path = getattr(wb, "_source_path", None)
    if source_path is None:
        raise RuntimeError("save requires write or modify mode")
    if _read_mode_has_pending_changes(wb):
        if getattr(wb, "_read_only", False):
            raise RuntimeError(
                "save() on a read_only=True workbook would discard pending changes; "
                "reopen with modify=True before editing"
            )
        _promote_read_mode_to_patcher(wb, source_path)
        save_modify_mode(wb, filename)
        return
    shutil.copyfile(source_path, filename)
    normalize_openpyxl_package_shape(wb, filename)


def apply_workbook_template_content_type(wb: Any, filename: str) -> None:
    """Mirror openpyxl's ``Workbook.template`` content-type switch."""
    import tempfile
    import zipfile
    from xml.etree import ElementTree as ET

    try:
        with zipfile.ZipFile(filename, "r") as src:
            infos = src.infolist()
            parts = {info.filename: src.read(info.filename) for info in infos}
    except (OSError, zipfile.BadZipFile):
        return

    content_types = parts.get(ARC_CONTENT_TYPES)
    if content_types is None:
        return
    try:
        root = ET.fromstring(content_types)
    except ET.ParseError:
        return

    template = bool(getattr(wb, "template", False))
    content_type = XLTX if template else XLSX
    workbook_override = None
    for child in root:
        if (
            child.tag.rsplit("}", 1)[-1] == "Override"
            and child.get("PartName") == "/xl/workbook.xml"
        ):
            workbook_override = child
            break
    if workbook_override is not None and workbook_override.get("ContentType") in {XLSM, XLTM}:
        content_type = XLTM if template else XLSM

    changed = False
    ET.register_namespace("", CONTYPES_NS)
    if workbook_override is not None:
        if workbook_override.get("ContentType") != content_type:
            workbook_override.set("ContentType", content_type)
            changed = True
    else:
        ET.SubElement(
            root,
            f"{{{CONTYPES_NS}}}Override",
            {"PartName": "/xl/workbook.xml", "ContentType": content_type},
        )
        changed = True

    if not changed:
        return

    parts[ARC_CONTENT_TYPES] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    fd, tmp_name = tempfile.mkstemp(prefix="wolfxl-template-shape-", suffix=".xlsx")
    os.close(fd)
    try:
        with zipfile.ZipFile(tmp_name, "w", zipfile.ZIP_DEFLATED) as dst:
            emitted: set[str] = set()
            for info in infos:
                data = parts.get(info.filename)
                if data is None:
                    continue
                dst.writestr(info, data)
                emitted.add(info.filename)
            for name, data in parts.items():
                if name not in emitted:
                    dst.writestr(name, data)
        os.replace(tmp_name, filename)
    finally:
        try:
            os.unlink(tmp_name)
        except OSError:
            pass


def _promote_read_mode_to_patcher(wb: Any, source_path: str) -> None:
    from wolfxl import _rust

    wb._rust_patcher = _rust.XlsxPatcher.open(source_path, False)


def _read_mode_has_pending_changes(wb: Any) -> bool:
    workbook_pending_attrs = (
        "_chartsheets_dirty",
        "_properties_dirty",
        "_custom_doc_props_dirty",
        "_pending_defined_names",
        "_pending_security_update",
        "_pending_axis_shifts",
        "_pending_range_moves",
        "_pending_sheet_copies",
        "_pending_chart_adds",
        "_pending_source_chart_ops",
        "_pending_pivot_caches",
        "_pending_slicer_caches",
        "_strip_external_links_on_save",
        "_active_sheet_dirty",
    )
    if any(bool(getattr(wb, attr, None)) for attr in workbook_pending_attrs):
        return True
    links = getattr(wb, "_external_links_cache", None)
    if links is not None and getattr(links, "dirty", False):
        return True

    worksheet_pending_attrs = (
        "_dirty",
        "_dirty_values",
        "_append_buffer",
        "_bulk_writes",
        "_pending_comments",
        "_pending_threaded_comments",
        "_pending_hyperlinks",
        "_pending_tables",
        "_pending_data_validations",
        "_pending_conditional_formats",
        "_pending_rich_text",
        "_pending_array_formulas",
        "_pending_images",
        "_pending_charts",
        "_pending_chart_deletions",
        "_pending_pivot_tables",
        "_pending_slicers",
        "_print_titles_dirty",
        "_sheet_state_dirty",
    )
    from wolfxl._worksheet_media import has_pending_image_deletions

    for ws in getattr(wb, "_sheets", {}).values():
        if any(bool(getattr(ws, attr, None)) for attr in worksheet_pending_attrs):
            return True
        if getattr(ws, "_header_footer", None) is not None:
            return True
        if has_pending_image_deletions(ws):
            return True
        for handle in getattr(ws, "_pivot_handles_cache", None) or []:
            if getattr(handle, "_dirty", False) or getattr(handle, "_layout_dirty", False):
                return True
    return False


def save_write_only_mode(wb: Any, filename: str) -> None:
    """Save path for ``Workbook(write_only=True)`` (G20).

    The streaming temp files have been accumulating row XML throughout
    the session. This flushes each per-sheet ``BufWriter`` then drives
    the standard ``emit_xlsx`` pipeline — the only difference from
    eager save is that ``sheet_xml::emit`` splices the temp file into
    the ``<sheetData>`` slot instead of walking ``Worksheet.rows``.
    """
    # Workbook-level metadata flush (defined names, properties) still
    # composes through the writer-side path; sheet-level flush is a
    # no-op for write-only sheets because the temp files have been
    # written incrementally.
    wb._flush_workbook_writes()  # noqa: SLF001
    wb._rust_writer.finalize_streaming_sheets()  # noqa: SLF001
    wb._rust_writer.save(filename)  # noqa: SLF001
    flush_chartsheets_authoring(wb, filename)


def save_modify_mode(wb: Any, filename: str) -> None:
    """Flush pending modify-mode queues and write through ``XlsxPatcher``."""
    if not _modify_mode_has_pending_changes(wb):
        if same_existing_path(filename, wb._source_path):  # noqa: SLF001
            wb._rust_patcher.save_in_place()  # noqa: SLF001
        else:
            wb._rust_patcher.save(filename)  # noqa: SLF001
        normalize_openpyxl_package_shape(wb, filename)
        return

    # Workbook-level metadata flushes before per-sheet drains so the patcher
    # composes workbook.xml once, with all pending workbook-scoped edits.
    if wb._properties_dirty:  # noqa: SLF001
        wb._flush_properties_to_patcher()  # noqa: SLF001
    wb._flush_defined_names_to_patcher()  # noqa: SLF001
    # Workbook-level security also targets workbook.xml and must precede
    # sheet-scoped patch queues.
    if wb._pending_security_update:  # noqa: SLF001
        wb._flush_security_to_patcher()  # noqa: SLF001
    for ws in wb._sheets.values():  # noqa: SLF001
        ws._flush()  # noqa: SLF001
    # Sheet copies must flush before every per-sheet phase so cloned sheets are
    # visible to downstream drains as if they had always been part of the
    # source workbook.
    wb._flush_pending_sheet_copies_to_patcher()  # noqa: SLF001
    # Hyperlinks share the sheet rels graph with tables and comments. Flush
    # them first so validations/conditional formats run afterward against an
    # already-stable rels graph.
    wb._flush_pending_hyperlinks_to_patcher()  # noqa: SLF001
    # Tables also touch the rels graph, add ZIP parts, and add content-type
    # overrides. Flush after hyperlinks so external-hyperlink rIds are stable.
    wb._flush_pending_tables_to_patcher()  # noqa: SLF001
    # Threaded comments + person list (RFC-068 G08). Drained BEFORE the
    # legacy comments flush so the Rust patcher's threaded-comments phase
    # (which synthesizes `tc={topId}` placeholders) can pre-populate
    # queued_comments before apply_comments_phase runs.
    wb._flush_pending_threaded_comments_to_patcher()  # noqa: SLF001
    wb._flush_pending_persons_to_patcher()  # noqa: SLF001
    # Comments and VML drawings.
    wb._flush_pending_comments_to_patcher()  # noqa: SLF001
    # Worksheet-level data validation setters.
    wb._flush_pending_data_validations_to_patcher()  # noqa: SLF001
    # Conditional formatting sibling blocks.
    wb._flush_pending_conditional_formats_to_patcher()  # noqa: SLF001
    # Structural axis shifts. Drained last among core sheet-data phases so they
    # see earlier per-cell and per-block rewrites.
    wb._flush_pending_axis_shifts_to_patcher()  # noqa: SLF001
    # Range moves. Drained after axis shifts so coordinate space is post-shift.
    wb._flush_pending_range_moves_to_patcher()  # noqa: SLF001
    # Images.
    wb._flush_pending_images_to_patcher()  # noqa: SLF001
    # Chart additions.
    wb._flush_pending_charts_to_patcher()  # noqa: SLF001
    # Pivot caches and tables.
    wb._flush_pending_pivots_to_patcher()  # noqa: SLF001
    # G17 / RFC-070 — pivot source-range edits on existing pivots.
    # Sequenced AFTER the adds flush so a session that both adds and
    # edits goes through the patcher's pivot phases in source order.
    wb._flush_pending_pivot_source_edits_to_patcher()  # noqa: SLF001
    # Sheet-setup blocks.
    wb._flush_pending_sheet_setup_to_patcher()  # noqa: SLF001
    # Page breaks and sheetFormatPr.
    wb._flush_pending_page_breaks_to_patcher()  # noqa: SLF001
    # Slicers.
    wb._flush_pending_slicers_to_patcher()  # noqa: SLF001
    # AutoFilter dicts.
    wb._flush_pending_autofilters_to_patcher()  # noqa: SLF001

    if same_existing_path(filename, wb._source_path):  # noqa: SLF001
        wb._rust_patcher.save_in_place()  # noqa: SLF001
    else:
        wb._rust_patcher.save(filename)  # noqa: SLF001
    flush_pivot_layout_authoring(wb, filename)
    flush_external_links_authoring(wb, filename)
    flush_source_chart_authoring(wb, filename)
    flush_chartsheets_authoring(wb, filename)
    normalize_openpyxl_package_shape(wb, filename)


def apply_custom_doc_props_authoring(wb: Any, filename: str) -> None:
    """Rewrite ``docProps/custom.xml`` after backend save when user-edited."""
    props = getattr(wb, "_custom_doc_props_cache", None)
    if props is None or not getattr(wb, "_custom_doc_props_dirty", False):
        return

    import os
    import tempfile
    import zipfile
    from xml.etree import ElementTree as ET

    from wolfxl.xml.functions import tostring

    try:
        with zipfile.ZipFile(filename, "r") as src:
            infos = src.infolist()
            parts = {info.filename: src.read(info.filename) for info in infos}
    except (OSError, zipfile.BadZipFile):
        return

    if len(props) >= 1:
        parts[ARC_CUSTOM] = tostring(props.to_tree())
        _ensure_custom_doc_props_relationship(parts, ET, DOC_NS, PKG_REL_NS)
        _ensure_custom_doc_props_content_type(parts, ET, CONTYPES_NS, CPROPS_TYPE)
    else:
        parts.pop(ARC_CUSTOM, None)
        _remove_custom_doc_props_relationship(parts, ET, DOC_NS)
        _remove_custom_doc_props_content_type(parts, ET)

    parent = os.path.dirname(os.path.abspath(filename)) or "."
    tmp = tempfile.NamedTemporaryFile(
        prefix=".wolfxl-custom-props-",
        suffix=".xlsx",
        dir=parent,
        delete=False,
    )
    tmp_name = tmp.name
    tmp.close()
    try:
        written: set[str] = set()
        with zipfile.ZipFile(tmp_name, "w", zipfile.ZIP_DEFLATED) as dst:
            for info in infos:
                if info.filename not in parts:
                    continue
                dst.writestr(info, parts[info.filename])
                written.add(info.filename)
            for name in sorted(set(parts) - written):
                dst.writestr(name, parts[name])
        os.replace(tmp_name, filename)
        wb._custom_doc_props_dirty = False  # noqa: SLF001
    finally:
        try:
            os.unlink(tmp_name)
        except OSError:
            pass


def _ensure_custom_doc_props_relationship(
    parts: dict[str, bytes],
    et: Any,
    doc_ns: str,
    pkg_rel_ns: str,
) -> None:
    rel_type = f"{doc_ns}relationships/custom-properties"
    rels_xml = parts.get(ARC_ROOT_RELS)
    et.register_namespace("", pkg_rel_ns)
    if rels_xml is None:
        root = et.Element(f"{{{pkg_rel_ns}}}Relationships")
    else:
        try:
            root = et.fromstring(rels_xml)
        except et.ParseError:
            return
    for rel in root:
        if rel.attrib.get("Type") == rel_type:
            rel.set("Target", "docProps/custom.xml")
            parts[ARC_ROOT_RELS] = et.tostring(
                root,
                encoding="utf-8",
                xml_declaration=True,
            )
            return
    existing_ids = {rel.attrib.get("Id", "") for rel in root}
    idx = 1
    while f"rId{idx}" in existing_ids:
        idx += 1
    et.SubElement(
        root,
        f"{{{pkg_rel_ns}}}Relationship",
        {
            "Id": f"rId{idx}",
            "Type": rel_type,
            "Target": "docProps/custom.xml",
        },
    )
    parts[ARC_ROOT_RELS] = et.tostring(root, encoding="utf-8", xml_declaration=True)


def _remove_custom_doc_props_relationship(
    parts: dict[str, bytes],
    et: Any,
    doc_ns: str,
) -> None:
    rels_xml = parts.get(ARC_ROOT_RELS)
    if rels_xml is None:
        return
    rel_type = f"{doc_ns}relationships/custom-properties"
    try:
        root = et.fromstring(rels_xml)
    except et.ParseError:
        return
    changed = False
    for rel in list(root):
        if rel.attrib.get("Type") == rel_type:
            root.remove(rel)
            changed = True
    if changed:
        parts[ARC_ROOT_RELS] = et.tostring(root, encoding="utf-8", xml_declaration=True)


def _ensure_custom_doc_props_content_type(
    parts: dict[str, bytes],
    et: Any,
    content_types_ns: str,
    custom_props_type: str,
) -> None:
    content_types = parts.get(ARC_CONTENT_TYPES)
    if content_types is None:
        return
    et.register_namespace("", content_types_ns)
    try:
        root = et.fromstring(content_types)
    except et.ParseError:
        return
    for node in root:
        if (
            node.tag.rsplit("}", 1)[-1] == "Override"
            and node.attrib.get("PartName") == "/docProps/custom.xml"
        ):
            node.set("ContentType", custom_props_type)
            parts[ARC_CONTENT_TYPES] = et.tostring(
                root,
                encoding="utf-8",
                xml_declaration=True,
            )
            return
    et.SubElement(
        root,
        f"{{{content_types_ns}}}Override",
        {
            "PartName": "/docProps/custom.xml",
            "ContentType": custom_props_type,
        },
    )
    parts[ARC_CONTENT_TYPES] = et.tostring(
        root,
        encoding="utf-8",
        xml_declaration=True,
    )


def _remove_custom_doc_props_content_type(parts: dict[str, bytes], et: Any) -> None:
    content_types = parts.get(ARC_CONTENT_TYPES)
    if content_types is None:
        return
    try:
        root = et.fromstring(content_types)
    except et.ParseError:
        return
    changed = False
    for node in list(root):
        if (
            node.tag.rsplit("}", 1)[-1] == "Override"
            and node.attrib.get("PartName") == "/docProps/custom.xml"
        ):
            root.remove(node)
            changed = True
    if changed:
        parts[ARC_CONTENT_TYPES] = et.tostring(
            root,
            encoding="utf-8",
            xml_declaration=True,
        )


def _modify_mode_has_pending_changes(wb: Any) -> bool:
    if _read_mode_has_pending_changes(wb):
        return True
    patcher = getattr(wb, "_rust_patcher", None)
    if patcher is not None:
        has_pending = getattr(patcher, "_has_pending_save_work", None)
        if has_pending is None:
            return True
        if bool(has_pending()):
            return True
    if _has_pending_source_chart_authoring(wb):
        return True
    if _has_pending_loaded_table_growth(wb):
        return True
    return False


def _has_pending_loaded_table_growth(wb: Any) -> bool:
    for ws in getattr(wb, "_sheets", {}).values():
        loaded = getattr(getattr(ws, "_tables_cache", None), "loaded", None)
        for table, saved_ref, saved_columns in (loaded or {}).values():
            if table.ref != saved_ref:
                return True
            if tuple(column.name for column in table.tableColumns) != saved_columns:
                return True
    return False


def _has_pending_source_chart_authoring(wb: Any) -> bool:
    if getattr(wb, "_pending_source_chart_ops", None):
        return True
    for ws in getattr(wb, "_sheets", {}).values():
        for chart in getattr(ws, "_charts_cache", None) or []:
            meta = getattr(chart, "_wolfxl_source_chart", None)
            if not meta:
                continue
            original_title = getattr(chart, "_wolfxl_source_title", None)
            if _source_chart_title_signature(chart) != original_title:
                return True
    return False


def apply_sheet_state_authoring(wb: Any, filename: str) -> None:
    """Apply user-authored workbook.xml state not handled by the native writer."""
    dirty_states = {
        title: ws.sheet_state
        for title, ws in getattr(wb, "_sheets", {}).items()
        if getattr(ws, "_sheet_state_dirty", False)
    }
    tab_colors = _pending_sheet_tab_colors(wb)
    outline_props = _pending_sheet_outline_properties(wb)
    page_setup_props = _pending_sheet_page_setup_properties(wb)
    workbook_view = _pending_workbook_view(wb)
    active_dirty = bool(getattr(wb, "_active_sheet_dirty", False))
    if (
        not dirty_states
        and not active_dirty
        and not tab_colors
        and not outline_props
        and not page_setup_props
        and workbook_view is None
    ):
        return

    import tempfile
    import zipfile
    from xml.etree import ElementTree as ET

    ns_main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    ns_rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    ET.register_namespace("", ns_main)
    ET.register_namespace("r", ns_rel)

    try:
        with zipfile.ZipFile(filename, "r") as src:
            infos = src.infolist()
            parts = {info.filename: src.read(info.filename) for info in infos}
    except (OSError, zipfile.BadZipFile):
        return
    workbook_xml = parts.get("xl/workbook.xml")
    if workbook_xml is None:
        return
    try:
        root = ET.fromstring(workbook_xml)
    except ET.ParseError:
        return

    changed = _apply_sheet_visibility_states(root, dirty_states)
    if workbook_view is not None:
        changed = _apply_workbook_view(root, ns_main, workbook_view) or changed
    if active_dirty:
        active_index = getattr(wb, "_active_index_for_save", wb._active_index)()
        changed = _apply_active_tab(root, ns_main, active_index) or changed
    if tab_colors:
        changed = _apply_sheet_tab_colors(parts, root, tab_colors, ns_main, ns_rel) or changed
    if outline_props:
        changed = _apply_sheet_outline_properties(
            parts,
            root,
            outline_props,
            ns_main,
            ns_rel,
        ) or changed
    if page_setup_props:
        changed = _apply_sheet_page_setup_properties(
            parts,
            root,
            page_setup_props,
            ns_main,
            ns_rel,
        ) or changed
    if not changed:
        _clear_workbook_xml_dirty_flags(wb)
        return

    parts["xl/workbook.xml"] = ET.tostring(
        root,
        encoding="utf-8",
        xml_declaration=True,
    )
    parent = os.path.dirname(os.path.abspath(filename)) or "."
    tmp = tempfile.NamedTemporaryFile(
        prefix=".wolfxl-sheet-state-",
        suffix=".xlsx",
        dir=parent,
        delete=False,
    )
    tmp_name = tmp.name
    tmp.close()
    try:
        with zipfile.ZipFile(tmp_name, "w", zipfile.ZIP_DEFLATED) as dst:
            for info in infos:
                dst.writestr(info, parts[info.filename])
        os.replace(tmp_name, filename)
    finally:
        try:
            os.unlink(tmp_name)
        except OSError:
            pass
    _clear_workbook_xml_dirty_flags(wb)


def _apply_sheet_visibility_states(root: Any, dirty_states: dict[str, str]) -> bool:
    changed = False
    for sheet in root.iter():
        if sheet.tag.rsplit("}", 1)[-1] != "sheet":
            continue
        name = sheet.attrib.get("name")
        if name not in dirty_states:
            continue
        state = dirty_states[name]
        if state == "visible":
            if "state" in sheet.attrib:
                del sheet.attrib["state"]
                changed = True
        elif sheet.attrib.get("state") != state:
            sheet.set("state", state)
            changed = True
    return changed


def _apply_active_tab(root: Any, ns_main: str, active_index: int | None) -> bool:
    from xml.etree import ElementTree as ET

    def local_name(node: Any) -> str:
        return node.tag.rsplit("}", 1)[-1]

    workbook_view = _first_workbook_view(root, ns_main, ET, local_name)
    if active_index is None:
        if "activeTab" not in workbook_view.attrib:
            return False
        del workbook_view.attrib["activeTab"]
        return True
    value = str(active_index)
    if workbook_view.attrib.get("activeTab") == value:
        return False
    workbook_view.set("activeTab", value)
    return True


def _first_workbook_view(root: Any, ns_main: str, et: Any, local_name: Any) -> Any:
    book_views = next(
        (child for child in list(root) if local_name(child) == "bookViews"),
        None,
    )
    if book_views is None:
        book_views = et.Element(f"{{{ns_main}}}bookViews")
        insert_at = 1 if list(root) and local_name(list(root)[0]) == "workbookPr" else 0
        root.insert(insert_at, book_views)
    workbook_view = next(
        (child for child in list(book_views) if local_name(child) == "workbookView"),
        None,
    )
    if workbook_view is None:
        workbook_view = et.SubElement(book_views, f"{{{ns_main}}}workbookView")
    return workbook_view


def _pending_workbook_view(wb: Any) -> dict[str, Any] | None:
    views = getattr(wb, "_views_cache", None)
    if not views:
        return None
    view = views[0]
    payload = {
        "visibility": getattr(view, "visibility", "visible"),
        "minimized": getattr(view, "minimized", False),
        "showHorizontalScroll": getattr(view, "showHorizontalScroll", True),
        "showVerticalScroll": getattr(view, "showVerticalScroll", True),
        "showSheetTabs": getattr(view, "showSheetTabs", True),
        "xWindow": getattr(view, "xWindow", None),
        "yWindow": getattr(view, "yWindow", None),
        "windowWidth": getattr(view, "windowWidth", None),
        "windowHeight": getattr(view, "windowHeight", None),
        "tabRatio": getattr(view, "tabRatio", 600),
        "firstSheet": getattr(view, "firstSheet", 0),
        "autoFilterDateGrouping": getattr(view, "autoFilterDateGrouping", True),
    }
    if payload == {
        "visibility": "visible",
        "minimized": False,
        "showHorizontalScroll": True,
        "showVerticalScroll": True,
        "showSheetTabs": True,
        "xWindow": None,
        "yWindow": None,
        "windowWidth": None,
        "windowHeight": None,
        "tabRatio": 600,
        "firstSheet": 0,
        "autoFilterDateGrouping": True,
    }:
        return None
    payload["activeTab"] = getattr(wb, "_active_index_for_save", wb._active_index)()
    return payload


def _apply_workbook_view(root: Any, ns_main: str, values: dict[str, Any]) -> bool:
    from xml.etree import ElementTree as ET

    def local_name(node: Any) -> str:
        return node.tag.rsplit("}", 1)[-1]

    workbook_view = _first_workbook_view(root, ns_main, ET, local_name)
    attr_specs = {
        "visibility": ("visibility", str),
        "minimized": ("minimized", _xml_bool),
        "showHorizontalScroll": ("showHorizontalScroll", _xml_bool),
        "showVerticalScroll": ("showVerticalScroll", _xml_bool),
        "showSheetTabs": ("showSheetTabs", _xml_bool),
        "xWindow": ("xWindow", str),
        "yWindow": ("yWindow", str),
        "windowWidth": ("windowWidth", str),
        "windowHeight": ("windowHeight", str),
        "tabRatio": ("tabRatio", str),
        "firstSheet": ("firstSheet", str),
        "activeTab": ("activeTab", str),
        "autoFilterDateGrouping": ("autoFilterDateGrouping", _xml_bool),
    }
    changed = False
    for key, (attr, formatter) in attr_specs.items():
        value = values.get(key)
        if value is None:
            if attr in workbook_view.attrib:
                del workbook_view.attrib[attr]
                changed = True
            continue
        attr_value = formatter(value)
        if workbook_view.attrib.get(attr) != attr_value:
            workbook_view.set(attr, attr_value)
            changed = True
    return changed


def _xml_bool(value: Any) -> str:
    return "1" if bool(value) else "0"


def _clear_workbook_xml_dirty_flags(wb: Any) -> None:
    for ws in getattr(wb, "_sheets", {}).values():
        ws._sheet_state_dirty = False  # noqa: SLF001
    wb._active_sheet_dirty = False  # noqa: SLF001


def apply_writer_sheet_deletes(wb: Any, filename: str) -> None:
    """Prune sheets removed from a write-mode workbook before save."""
    deleted = set(getattr(wb, "_pending_writer_sheet_deletes", []) or [])
    if not deleted:
        return

    import posixpath
    import tempfile
    import zipfile
    from xml.etree import ElementTree as ET

    ns_main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    ns_rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    ET.register_namespace("", ns_main)
    ET.register_namespace("r", ns_rel)

    try:
        with zipfile.ZipFile(filename, "r") as src:
            infos = src.infolist()
            parts = {info.filename: src.read(info.filename) for info in infos}
    except (OSError, zipfile.BadZipFile):
        return

    workbook_xml = parts.get("xl/workbook.xml")
    rels_xml = parts.get("xl/_rels/workbook.xml.rels")
    if workbook_xml is None or rels_xml is None:
        return
    try:
        workbook_root = ET.fromstring(workbook_xml)
        rels_root = ET.fromstring(rels_xml)
    except ET.ParseError:
        return

    rel_targets = _workbook_relationship_targets(parts, ns_rel)
    removed_rel_ids: set[str] = set()
    removed_paths: set[str] = set()
    for sheets_node in workbook_root.iter():
        if sheets_node.tag.rsplit("}", 1)[-1] != "sheets":
            continue
        for sheet in list(sheets_node):
            if sheet.tag.rsplit("}", 1)[-1] != "sheet":
                continue
            if sheet.attrib.get("name") not in deleted:
                continue
            rel_id = sheet.attrib.get(f"{{{ns_rel}}}id")
            if rel_id:
                removed_rel_ids.add(rel_id)
                target = rel_targets.get(rel_id)
                if target:
                    removed_paths.add(_workbook_relationship_target_to_part(target))
            sheets_node.remove(sheet)

    if not removed_rel_ids:
        wb._pending_writer_sheet_deletes = []  # noqa: SLF001
        return

    for rel in list(rels_root):
        if rel.attrib.get("Id") in removed_rel_ids:
            rels_root.remove(rel)

    for path in removed_paths:
        parts.pop(path, None)
        rels_path = posixpath.join(
            posixpath.dirname(path),
            "_rels",
            posixpath.basename(path) + ".rels",
        )
        parts.pop(rels_path, None)
    _remove_content_type_overrides(parts, removed_paths)

    parts["xl/workbook.xml"] = ET.tostring(workbook_root, encoding="utf-8", xml_declaration=True)
    parts["xl/_rels/workbook.xml.rels"] = ET.tostring(rels_root, encoding="utf-8", xml_declaration=True)

    parent = os.path.dirname(os.path.abspath(filename)) or "."
    tmp = tempfile.NamedTemporaryFile(
        prefix=".wolfxl-sheet-delete-",
        suffix=".xlsx",
        dir=parent,
        delete=False,
    )
    tmp_name = tmp.name
    tmp.close()
    try:
        with zipfile.ZipFile(tmp_name, "w", zipfile.ZIP_DEFLATED) as dst:
            for info in infos:
                if info.filename in parts:
                    dst.writestr(info, parts[info.filename])
        os.replace(tmp_name, filename)
    finally:
        try:
            os.unlink(tmp_name)
        except OSError:
            pass
    wb._pending_writer_sheet_deletes = []  # noqa: SLF001


def apply_writer_unmerged_ranges(wb: Any, filename: str) -> None:
    """Prune mergeCell entries removed from a write-mode workbook."""
    pending = {
        title: set(getattr(ws, "_pending_unmerged_ranges", set()) or set())
        for title, ws in getattr(wb, "_sheets", {}).items()
        if getattr(ws, "_pending_unmerged_ranges", None)
    }
    if not pending:
        return

    import tempfile
    import zipfile
    from xml.etree import ElementTree as ET

    ns_main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    ns_rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    ET.register_namespace("", ns_main)
    ET.register_namespace("r", ns_rel)

    try:
        with zipfile.ZipFile(filename, "r") as src:
            infos = src.infolist()
            parts = {info.filename: src.read(info.filename) for info in infos}
    except (OSError, zipfile.BadZipFile):
        return

    workbook_xml = parts.get("xl/workbook.xml")
    if workbook_xml is None:
        return
    try:
        workbook_root = ET.fromstring(workbook_xml)
    except ET.ParseError:
        return

    rel_targets = _workbook_relationship_targets(parts, ns_rel)
    sheet_paths: dict[str, str] = {}
    for sheet in workbook_root.iter():
        if sheet.tag.rsplit("}", 1)[-1] != "sheet":
            continue
        name = sheet.attrib.get("name")
        rel_id = sheet.attrib.get(f"{{{ns_rel}}}id")
        target = rel_targets.get(rel_id or "")
        if name and target:
            sheet_paths[name] = _workbook_relationship_target_to_part(target)

    changed = False
    for title, refs in pending.items():
        path = sheet_paths.get(title)
        if not path or path not in parts:
            continue
        try:
            sheet_root = ET.fromstring(parts[path])
        except ET.ParseError:
            continue
        if _remove_merge_cells(sheet_root, refs):
            parts[path] = ET.tostring(sheet_root, encoding="utf-8", xml_declaration=True)
            changed = True

    if not changed:
        _clear_pending_unmerged_ranges(wb)
        return

    parent = os.path.dirname(os.path.abspath(filename)) or "."
    tmp = tempfile.NamedTemporaryFile(
        prefix=".wolfxl-unmerge-",
        suffix=".xlsx",
        dir=parent,
        delete=False,
    )
    tmp_name = tmp.name
    tmp.close()
    try:
        with zipfile.ZipFile(tmp_name, "w", zipfile.ZIP_DEFLATED) as dst:
            for info in infos:
                dst.writestr(info, parts[info.filename])
        os.replace(tmp_name, filename)
    finally:
        try:
            os.unlink(tmp_name)
        except OSError:
            pass
    _clear_pending_unmerged_ranges(wb)


def _remove_merge_cells(sheet_root: Any, refs: set[str]) -> bool:
    changed = False
    for parent in list(sheet_root.iter()):
        if parent.tag.rsplit("}", 1)[-1] != "mergeCells":
            continue
        for child in list(parent):
            if (
                child.tag.rsplit("}", 1)[-1] == "mergeCell"
                and child.attrib.get("ref") in refs
            ):
                parent.remove(child)
                changed = True
        remaining = [child for child in list(parent) if child.tag.rsplit("}", 1)[-1] == "mergeCell"]
        if remaining:
            parent.set("count", str(len(remaining)))
        else:
            sheet_root.remove(parent)
    return changed


def _clear_pending_unmerged_ranges(wb: Any) -> None:
    for ws in getattr(wb, "_sheets", {}).values():
        if hasattr(ws, "_pending_unmerged_ranges"):
            ws._pending_unmerged_ranges.clear()  # noqa: SLF001


def _workbook_relationship_target_to_part(target: str) -> str:
    import posixpath

    normalized = posixpath.normpath(target.lstrip("/"))
    if normalized.startswith("xl/"):
        return normalized
    return posixpath.normpath(posixpath.join("xl", normalized))


def _remove_content_type_overrides(
    parts: dict[str, bytes],
    removed_paths: set[str],
) -> None:
    from xml.etree import ElementTree as ET

    content_types = parts.get("[Content_Types].xml")
    if content_types is None:
        return
    try:
        root = ET.fromstring(content_types)
    except ET.ParseError:
        return
    removed_part_names = {"/" + path for path in removed_paths}
    changed = False
    for node in list(root):
        if (
            node.tag.rsplit("}", 1)[-1] == "Override"
            and node.attrib.get("PartName") in removed_part_names
        ):
            root.remove(node)
            changed = True
    if changed:
        parts["[Content_Types].xml"] = ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _pending_sheet_tab_colors(wb: Any) -> dict[str, str]:
    colors: dict[str, str] = {}
    for title, ws in getattr(wb, "_sheets", {}).items():
        properties = getattr(ws, "_sheet_properties", None)
        color = getattr(properties, "tabColor", None)
        normalized = _normalize_tab_color(color)
        if normalized:
            colors[title] = normalized
    return colors


def _pending_sheet_outline_properties(wb: Any) -> dict[str, dict[str, bool | None]]:
    outlines: dict[str, dict[str, bool | None]] = {}
    for title, ws in getattr(wb, "_sheets", {}).items():
        properties = getattr(ws, "_sheet_properties", None)
        outline = getattr(properties, "outlinePr", None)
        if outline is None:
            continue
        values = {
            "summaryBelow": getattr(outline, "summaryBelow", True),
            "summaryRight": getattr(outline, "summaryRight", True),
            "applyStyles": getattr(outline, "applyStyles", None),
            "showOutlineSymbols": getattr(outline, "showOutlineSymbols", None),
        }
        if values != {
            "summaryBelow": True,
            "summaryRight": True,
            "applyStyles": None,
            "showOutlineSymbols": None,
        }:
            outlines[title] = values
    return outlines


def _pending_sheet_page_setup_properties(wb: Any) -> dict[str, dict[str, bool | None]]:
    page_setups: dict[str, dict[str, bool | None]] = {}
    for title, ws in getattr(wb, "_sheets", {}).items():
        properties = getattr(ws, "_sheet_properties", None)
        page_setup = getattr(properties, "pageSetUpPr", None)
        if page_setup is None:
            continue
        values = {
            "autoPageBreaks": getattr(page_setup, "autoPageBreaks", None),
            "fitToPage": getattr(page_setup, "fitToPage", None),
        }
        if values != {
            "autoPageBreaks": None,
            "fitToPage": None,
        }:
            page_setups[title] = values
    return page_setups


def _normalize_tab_color(color: Any) -> str | None:
    if color is None:
        return None
    rgb = getattr(color, "rgb", color)
    if rgb is None:
        return None
    value = str(rgb).removeprefix("#").upper()
    if len(value) == 6:
        return "00" + value
    if len(value) == 8:
        return value
    return None


def _sheet_paths_by_title(
    parts: dict[str, bytes],
    workbook_root: Any,
    ns_rel: str,
) -> dict[str, str]:
    rel_targets = _workbook_relationship_targets(parts, ns_rel)
    sheet_paths: dict[str, str] = {}
    for sheet in workbook_root.iter():
        if sheet.tag.rsplit("}", 1)[-1] != "sheet":
            continue
        name = sheet.attrib.get("name")
        rel_id = sheet.attrib.get(f"{{{ns_rel}}}id")
        target = rel_targets.get(rel_id or "")
        if name and target:
            sheet_paths[name] = _workbook_relationship_target_to_part(target)
    return sheet_paths


def _apply_sheet_tab_colors(
    parts: dict[str, bytes],
    workbook_root: Any,
    tab_colors: dict[str, str],
    ns_main: str,
    ns_rel: str,
) -> bool:
    from xml.etree import ElementTree as ET

    sheet_paths = _sheet_paths_by_title(parts, workbook_root, ns_rel)
    changed = False
    for title, color in tab_colors.items():
        path = sheet_paths.get(title)
        if not path or path not in parts:
            continue
        try:
            sheet_root = ET.fromstring(parts[path])
        except ET.ParseError:
            continue
        if _apply_sheet_tab_color(sheet_root, color, ns_main):
            parts[path] = ET.tostring(
                sheet_root,
                encoding="utf-8",
                xml_declaration=True,
            )
            changed = True
    return changed


def _apply_sheet_outline_properties(
    parts: dict[str, bytes],
    workbook_root: Any,
    outline_props: dict[str, dict[str, bool | None]],
    ns_main: str,
    ns_rel: str,
) -> bool:
    from xml.etree import ElementTree as ET

    sheet_paths = _sheet_paths_by_title(parts, workbook_root, ns_rel)
    changed = False
    for title, values in outline_props.items():
        path = sheet_paths.get(title)
        if not path or path not in parts:
            continue
        try:
            sheet_root = ET.fromstring(parts[path])
        except ET.ParseError:
            continue
        if _apply_sheet_outline_property(sheet_root, values, ns_main):
            parts[path] = ET.tostring(
                sheet_root,
                encoding="utf-8",
                xml_declaration=True,
            )
            changed = True
    return changed


def _apply_sheet_page_setup_properties(
    parts: dict[str, bytes],
    workbook_root: Any,
    page_setup_props: dict[str, dict[str, bool | None]],
    ns_main: str,
    ns_rel: str,
) -> bool:
    from xml.etree import ElementTree as ET

    sheet_paths = _sheet_paths_by_title(parts, workbook_root, ns_rel)
    changed = False
    for title, values in page_setup_props.items():
        path = sheet_paths.get(title)
        if not path or path not in parts:
            continue
        try:
            sheet_root = ET.fromstring(parts[path])
        except ET.ParseError:
            continue
        if _apply_sheet_page_setup_property(sheet_root, values, ns_main):
            parts[path] = ET.tostring(
                sheet_root,
                encoding="utf-8",
                xml_declaration=True,
            )
            changed = True
    return changed


def _workbook_relationship_targets(parts: dict[str, bytes], ns_rel: str) -> dict[str, str]:
    from xml.etree import ElementTree as ET

    rels_xml = parts.get("xl/_rels/workbook.xml.rels")
    if rels_xml is None:
        return {}
    try:
        rels_root = ET.fromstring(rels_xml)
    except ET.ParseError:
        return {}
    targets: dict[str, str] = {}
    for rel in rels_root.iter():
        if rel.tag.rsplit("}", 1)[-1] != "Relationship":
            continue
        rel_id = rel.attrib.get("Id")
        target = rel.attrib.get("Target")
        if rel_id and target:
            targets[rel_id] = target
    return targets


def _apply_sheet_tab_color(sheet_root: Any, color: str, ns_main: str) -> bool:
    from xml.etree import ElementTree as ET

    def local_name(node: Any) -> str:
        return node.tag.rsplit("}", 1)[-1]

    children = list(sheet_root)
    sheet_pr = next((child for child in children if local_name(child) == "sheetPr"), None)
    if sheet_pr is None:
        sheet_pr = ET.Element(f"{{{ns_main}}}sheetPr")
        sheet_root.insert(0, sheet_pr)
    tab_color = next(
        (child for child in list(sheet_pr) if local_name(child) == "tabColor"),
        None,
    )
    if tab_color is None:
        tab_color = ET.SubElement(sheet_pr, f"{{{ns_main}}}tabColor")
    if tab_color.attrib.get("rgb") == color:
        return False
    tab_color.attrib.clear()
    tab_color.set("rgb", color)
    return True


def _apply_sheet_outline_property(
    sheet_root: Any,
    values: dict[str, bool | None],
    ns_main: str,
) -> bool:
    from xml.etree import ElementTree as ET

    def local_name(node: Any) -> str:
        return node.tag.rsplit("}", 1)[-1]

    children = list(sheet_root)
    sheet_pr = next((child for child in children if local_name(child) == "sheetPr"), None)
    if sheet_pr is None:
        sheet_pr = ET.Element(f"{{{ns_main}}}sheetPr")
        sheet_root.insert(0, sheet_pr)
    outline_pr = next(
        (child for child in list(sheet_pr) if local_name(child) == "outlinePr"),
        None,
    )
    if outline_pr is None:
        outline_pr = ET.SubElement(sheet_pr, f"{{{ns_main}}}outlinePr")

    changed = False
    for key in ("summaryBelow", "summaryRight", "applyStyles", "showOutlineSymbols"):
        value = values.get(key)
        if value is None:
            if key in outline_pr.attrib:
                del outline_pr.attrib[key]
                changed = True
            continue
        attr_value = "1" if bool(value) else "0"
        if outline_pr.attrib.get(key) != attr_value:
            outline_pr.set(key, attr_value)
            changed = True
    return changed


def _apply_sheet_page_setup_property(
    sheet_root: Any,
    values: dict[str, bool | None],
    ns_main: str,
) -> bool:
    from xml.etree import ElementTree as ET

    def local_name(node: Any) -> str:
        return node.tag.rsplit("}", 1)[-1]

    children = list(sheet_root)
    sheet_pr = next((child for child in children if local_name(child) == "sheetPr"), None)
    if sheet_pr is None:
        sheet_pr = ET.Element(f"{{{ns_main}}}sheetPr")
        sheet_root.insert(0, sheet_pr)
    page_setup_pr = next(
        (child for child in list(sheet_pr) if local_name(child) == "pageSetUpPr"),
        None,
    )
    if page_setup_pr is None:
        page_setup_pr = ET.SubElement(sheet_pr, f"{{{ns_main}}}pageSetUpPr")

    changed = False
    for key in ("autoPageBreaks", "fitToPage"):
        value = values.get(key)
        if value is None:
            if key in page_setup_pr.attrib:
                del page_setup_pr.attrib[key]
                changed = True
            continue
        attr_value = "1" if bool(value) else "0"
        if page_setup_pr.attrib.get(key) != attr_value:
            page_setup_pr.set(key, attr_value)
            changed = True
    return changed


def save_write_mode(wb: Any, filename: str) -> None:
    """Flush pending write-mode queues and save through ``NativeWorkbook``.

    Write-mode pivot construction (G17 / RFC-070 §8.7 reach-extension):
    when pending pivot caches or pivot tables are queued, a two-phase
    save runs — the native writer emits cell data + sheet structure
    to a tempfile, then the same tempfile is reopened in modify mode
    and the patcher's ``apply_pivot_adds_phase`` stamps in the pivot
    parts. Final bytes copy onto ``filename``.
    """
    has_pending_pivots = bool(getattr(wb, "_pending_pivot_caches", None)) or any(
        getattr(ws, "_pending_pivot_tables", None) for ws in wb._sheets.values()  # noqa: SLF001
    )
    if not has_pending_pivots:
        wb._flush_workbook_writes()  # noqa: SLF001
        for ws in wb._sheets.values():  # noqa: SLF001
            ws._flush()  # noqa: SLF001
        wb._rust_writer.save(filename)  # noqa: SLF001
        flush_external_links_authoring(wb, filename)
        flush_chartsheets_authoring(wb, filename)
        return

    _save_write_mode_with_pivots(wb, filename)
    flush_external_links_authoring(wb, filename)
    flush_chartsheets_authoring(wb, filename)


def flush_external_links_authoring(wb: Any, filename: str) -> None:
    links = getattr(wb, "_external_links_cache", None)
    strip_links = bool(getattr(wb, "_strip_external_links_on_save", False))
    if links is None and strip_links:
        links = wb._external_links  # noqa: SLF001
    if links is None or (not getattr(links, "dirty", False) and not strip_links):
        return
    from wolfxl import _external_links as _el

    if strip_links:
        links._mark_dirty()  # noqa: SLF001
    _el.apply_authoring_to_xlsx(filename, links)
    wb._strip_external_links_on_save = False  # noqa: SLF001


def flush_chartsheets_authoring(wb: Any, filename: str) -> None:
    chartsheets = getattr(wb, "_chartsheets", None)
    if not chartsheets or not any(
        not getattr(cs, "_source_chartsheet", False) for cs in chartsheets.values()
    ):
        return
    from wolfxl import _chartsheets

    _chartsheets.apply_chartsheets_to_xlsx(filename, wb)


def flush_source_chart_authoring(wb: Any, filename: str) -> None:
    ops = list(getattr(wb, "_pending_source_chart_ops", []))
    touched = {
        op.get("meta", {}).get("chart_path")
        for op in ops
        if isinstance(op.get("meta"), dict)
    }
    for ws in getattr(wb, "_sheets", {}).values():
        for chart in getattr(ws, "_charts_cache", None) or []:
            meta = getattr(chart, "_wolfxl_source_chart", None)
            if not meta or meta.get("chart_path") in touched:
                continue
            original_title = getattr(chart, "_wolfxl_source_title", None)
            current_title = _source_chart_title_signature(chart)
            if current_title != original_title:
                ops.append({"op": "title", "meta": meta, "chart": chart})
                touched.add(meta.get("chart_path"))

    if not ops:
        return

    from wolfxl import _source_charts

    materialized = []
    for op in ops:
        if op["op"] in {"replace", "title"}:
            op = dict(op)
            op["chart_xml"] = _serialize_chart_xml(op["chart"])
        materialized.append(op)
    _source_charts.apply_source_chart_authoring_to_xlsx(filename, materialized)
    wb._pending_source_chart_ops.clear()  # noqa: SLF001


def _serialize_chart_xml(chart: Any) -> bytes:
    from wolfxl._rust import serialize_chart_dict  # type: ignore[attr-defined]

    original_anchor = getattr(chart, "_anchor", None)
    chart._anchor = "E15"  # noqa: SLF001
    try:
        return serialize_chart_dict(chart.to_rust_dict(), "E15")
    finally:
        chart._anchor = original_anchor  # noqa: SLF001


def _source_chart_title_signature(chart: Any) -> Any:
    title = getattr(chart, "title", None)
    if title is None:
        return None
    to_dict = getattr(title, "to_dict", None)
    if to_dict is not None:
        try:
            return to_dict()
        except Exception:
            pass
    return str(title)


def flush_pivot_layout_authoring(wb: Any, filename: str) -> None:
    from wolfxl.pivot._handle import apply_pivot_layout_authoring_to_xlsx

    apply_pivot_layout_authoring_to_xlsx(filename, wb)


def _save_write_mode_with_pivots(wb: Any, filename: str) -> None:
    """Two-phase save: writer → tempfile → patcher → final destination.

    This unblocks the openpyxl-shaped pattern used by the G17 oracle
    probe:

    .. code-block:: python

        wb = wolfxl.Workbook()
        wb.add_pivot_cache(cache)
        ws.add_pivot_table(pt, "A1")
        wb.save(path)

    by emitting cell data through the native writer, then surgically
    grafting the pivot parts on via the existing modify-mode patcher
    pipeline. No new pivot emit code is introduced.
    """
    import tempfile

    # Stash the pending pivot queues — the writer-side flush would
    # otherwise observe them and (in some future hardening) try to
    # double-emit. Re-attach them onto the modify-mode workbook below.
    pending_caches = list(getattr(wb, "_pending_pivot_caches", []))
    pending_tables_by_sheet: dict[str, list[Any]] = {}
    for ws in wb._sheets.values():  # noqa: SLF001
        title = ws.title
        pending_tables_by_sheet[title] = list(
            getattr(ws, "_pending_pivot_tables", [])
        )
        ws._pending_pivot_tables = []  # noqa: SLF001
    wb._pending_pivot_caches = []  # noqa: SLF001
    # Reset cache id allocator and clear cache._cache_id stamps so the
    # modify-mode add_pivot_cache call re-allocates cleanly.
    wb._next_pivot_cache_id = 0  # noqa: SLF001
    for cache in pending_caches:
        cache._cache_id = None  # noqa: SLF001

    # Stage 1 — writer emits the workbook (cells, sheets, formats, etc.).
    wb._flush_workbook_writes()  # noqa: SLF001
    for ws in wb._sheets.values():  # noqa: SLF001
        ws._flush()  # noqa: SLF001

    tmp_fd, tmp_name = tempfile.mkstemp(prefix=".wolfxl-pivot-", suffix=".xlsx")
    os.close(tmp_fd)
    try:
        wb._rust_writer.save(tmp_name)  # noqa: SLF001

        # Stage 2 — reopen the tempfile in modify mode and replay the
        # pivot adds.
        import wolfxl as _wolfxl

        modify_wb = _wolfxl.load_workbook(tmp_name, modify=True)
        for cache in pending_caches:
            modify_wb.add_pivot_cache(cache)
        for sheet_title, pts in pending_tables_by_sheet.items():
            if not pts:
                continue
            target = modify_wb._sheets.get(sheet_title)  # noqa: SLF001
            if target is None:
                continue
            for pt in pts:
                # add_pivot_table on Worksheet expects a single arg in
                # write/modify mode; the anchor is captured on the
                # PivotTable's `location` attr at construction time.
                target.add_pivot_table(pt)
        modify_wb.save(filename)
    finally:
        try:
            os.unlink(tmp_name)
        except OSError:
            pass


def save_encrypted(wb: Any, filename: str, password: str | bytes) -> None:
    """Save plaintext to a tempfile, then encrypt it atomically."""
    import tempfile

    from wolfxl._encryption import encrypt_xlsx_to_path

    tmp_fd, tmp_name = tempfile.mkstemp(
        prefix=".wolfxl-plain-",
        suffix=".xlsx",
    )
    os.close(tmp_fd)
    try:
        # Re-enter the normal plaintext save path so writer and patcher modes
        # exercise the same pipeline before encryption.
        save_workbook(wb, tmp_name)
        with open(tmp_name, "rb") as fp:
            plaintext_bytes = fp.read()
        encrypt_xlsx_to_path(plaintext_bytes, password, filename)
    finally:
        try:
            os.unlink(tmp_name)
        except OSError:
            pass
