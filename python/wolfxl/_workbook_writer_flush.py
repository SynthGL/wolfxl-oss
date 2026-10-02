"""Write-mode workbook flush helpers for the native Rust writer."""

from __future__ import annotations

from typing import Any


def flush_workbook_writes(wb: Any) -> None:
    """Push workbook metadata, names, and security into the native writer."""
    writer = wb._rust_writer  # noqa: SLF001
    if writer is None:
        return

    _flush_workbook_properties(wb, writer)
    _flush_properties(wb, writer)
    _flush_defined_names(wb, writer)
    _flush_named_styles(wb, writer)
    _flush_security(wb, writer)
    _flush_persons(wb, writer)


def _flush_properties(wb: Any, writer: Any) -> None:
    """Push dirty document properties to the native writer."""
    if wb._properties_dirty and wb._properties_cache is not None:  # noqa: SLF001
        props = wb._properties_cache  # noqa: SLF001
        payload = {
            "title": props.title,
            "subject": props.subject,
            "creator": props.creator,
            "keywords": props.keywords,
            "description": props.description,
            "lastModifiedBy": props.lastModifiedBy,
            "category": props.category,
            "contentStatus": props.contentStatus,
            "identifier": props.identifier,
            "language": props.language,
            "revision": props.revision,
            "version": props.version,
            "created": props.created.isoformat() if props.created else None,
            "modified": props.modified.isoformat() if props.modified else None,
        }
        writer.set_properties(payload)
        wb._properties_dirty = False  # noqa: SLF001


def _flush_workbook_properties(wb: Any, writer: Any) -> None:
    """Push workbookPr properties that affect writer serialization."""
    props = getattr(wb, "_workbook_properties_cache", None)  # noqa: SLF001
    if props is None or not hasattr(writer, "set_date1904"):
        return
    writer.set_date1904(bool(getattr(props, "date1904", False)))


def _flush_defined_names(wb: Any, writer: Any) -> None:
    """Push pending defined names to the native writer."""
    from wolfxl._workbook_metadata import queue_worksheet_defined_names

    queue_worksheet_defined_names(wb)
    _flush_print_titles(wb, writer)
    if wb._pending_defined_names:  # noqa: SLF001
        primary_sheet = wb._sheet_names[0] if wb._sheet_names else "Sheet"  # noqa: SLF001
        for defined_name in wb._pending_defined_names.values():  # noqa: SLF001
            if getattr(defined_name, "delete", False):
                continue
            if defined_name.localSheetId is not None:
                if 0 <= defined_name.localSheetId < len(wb._sheet_names):  # noqa: SLF001
                    sheet_hint = wb._sheet_names[defined_name.localSheetId]  # noqa: SLF001
                else:
                    sheet_hint = primary_sheet
                scope = "sheet"
            else:
                sheet_hint = primary_sheet
                scope = "workbook"
            writer.add_named_range(
                sheet_hint,
                {
                    "name": defined_name.name,
                    "refers_to": defined_name.value,
                    "scope": scope,
                    "comment": defined_name.comment,
                    "local_sheet_id": defined_name.localSheetId,
                    "hidden": defined_name.hidden,
                    # G22 — full ECMA-376 attribute surface.
                    "custom_menu": defined_name.custom_menu,
                    "description": defined_name.description,
                    "help": defined_name.help,
                    "status_bar": defined_name.status_bar,
                    "shortcut_key": defined_name.shortcut_key,
                    "function": defined_name.function,
                    "function_group_id": defined_name.function_group_id,
                    "vb_procedure": defined_name.vb_procedure,
                    "xlm": defined_name.xlm,
                    "publish_to_server": defined_name.publish_to_server,
                    "workbook_parameter": defined_name.workbook_parameter,
                },
            )
        wb._pending_defined_names.clear()  # noqa: SLF001


def _flush_print_titles(wb: Any, writer: Any) -> None:
    """Push worksheet repeat-row/column titles as reserved defined names."""
    from wolfxl.worksheet.print_settings import ColRange, PrintTitles, RowRange

    for sheet_idx, sheet_name in enumerate(wb._sheet_names):  # noqa: SLF001
        ws = wb._sheets.get(sheet_name)  # noqa: SLF001
        if ws is None:
            continue
        rows = getattr(ws, "_print_title_rows", None)
        cols = getattr(ws, "_print_title_cols", None)
        if rows is None and cols is None:
            continue
        if _has_pending_print_titles_for_sheet(wb, sheet_idx):
            continue
        titles = PrintTitles(
            rows=RowRange.from_string(rows) if rows is not None else None,
            cols=ColRange.from_string(cols) if cols is not None else None,
        )
        formula = titles.to_definedname_value(sheet_name)
        if formula is None:
            continue
        writer.add_named_range(
            sheet_name,
            {
                "name": "_xlnm.Print_Titles",
                "refers_to": formula,
                "scope": "sheet",
                "local_sheet_id": sheet_idx,
            },
        )


def _has_pending_print_titles_for_sheet(wb: Any, sheet_idx: int) -> bool:
    """Return whether user-defined names already carry this sheet's titles."""
    return any(
        defined_name.name == "_xlnm.Print_Titles"
        and defined_name.localSheetId == sheet_idx
        for defined_name in wb._pending_defined_names.values()  # noqa: SLF001
    )


def _flush_named_styles(wb: Any, writer: Any) -> None:
    """Push workbook named-style registrations to the native writer."""
    if not hasattr(writer, "add_named_style"):
        return
    registry = getattr(wb, "_named_styles_registry", None)  # noqa: SLF001
    if registry is None:
        return
    referenced_names = _referenced_named_styles(wb)
    for style in registry.user_styles():
        writer.add_named_style(style.name)
    for style in registry:
        if style.is_builtin and style.name in referenced_names and style.name != "Normal":
            writer.add_named_style(style.name)


def _referenced_named_styles(wb: Any) -> set[str]:
    """Return non-default named styles assigned to pending cell formats."""
    names: set[str] = set()
    for ws in getattr(wb, "_sheets", {}).values():
        cells = getattr(ws, "_cells", {})
        dirty = getattr(ws, "_dirty", set())
        for key in dirty:
            cell = cells.get(key)
            if cell is None:
                continue
            name = getattr(cell, "_named_style", None)
            if isinstance(name, str) and name and name != "Normal":
                names.add(name)
    return names


def _flush_security(wb: Any, writer: Any) -> None:
    """Push pending workbook security metadata to the native writer."""
    if wb._pending_security_update:  # noqa: SLF001
        payload = wb._build_security_dict()  # noqa: SLF001
        if hasattr(writer, "set_workbook_security"):
            writer.set_workbook_security(payload)
        wb._pending_security_update = False  # noqa: SLF001


def _flush_persons(wb: Any, writer: Any) -> None:
    """Push the workbook-scope threaded-comment person registry (RFC-068).

    The Rust ``add_person`` entry point is idempotent on Person ``id`` so
    repeated flush calls produce a stable personList. The registry is only
    seeded lazily; if the user never accessed ``wb.persons`` there is
    nothing to flush.
    """
    if not hasattr(writer, "add_person"):
        return
    registry = getattr(wb, "_persons_registry", None)  # noqa: SLF001
    if registry is None or len(registry) == 0:
        return
    for person in registry:
        if person.id is None:
            # PersonRegistry should always allocate a GUID, but guard the
            # contract so a misbehaving caller can't poison the save path.
            continue
        writer.add_person(
            {
                "id": person.id,
                "name": person.name,
                "user_id": person.user_id,
                "provider_id": person.provider_id,
            }
        )
