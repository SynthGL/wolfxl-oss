"""Workbook metadata, defined-name, and security helper APIs."""

from __future__ import annotations

from typing import Any


def get_properties(wb: Any) -> Any:
    """Return a lazy-loaded workbook document properties object.

    Args:
        wb: Workbook-like object carrying reader and properties cache state.

    Returns:
        A ``DocumentProperties`` instance attached back to ``wb`` so later
        attribute writes mark workbook metadata dirty.
    """
    if wb._properties_cache is not None:  # noqa: SLF001
        return wb._properties_cache  # noqa: SLF001
    from wolfxl.packaging.core import DocumentProperties, _doc_props_from_dict

    if wb._rust_reader is not None:  # noqa: SLF001
        try:
            raw = wb._rust_reader.read_doc_properties()  # noqa: SLF001
        except Exception:
            raw = {}
        props = _doc_props_from_dict(raw)
    else:
        props = DocumentProperties()
    props._attach_workbook(wb)  # noqa: SLF001
    wb._properties_cache = props  # noqa: SLF001
    return props


def set_properties(wb: Any, value: Any) -> None:
    """Replace the workbook document properties object.

    Args:
        wb: Workbook-like object carrying properties cache state.
        value: ``DocumentProperties`` instance to attach and mark dirty.

    Raises:
        TypeError: If ``value`` is not a ``DocumentProperties`` instance.
    """
    from wolfxl.packaging.core import DocumentProperties

    if not isinstance(value, DocumentProperties):
        raise TypeError(
            f"properties must be a DocumentProperties, got {type(value).__name__}"
        )
    value._attach_workbook(wb)  # noqa: SLF001
    wb._properties_cache = value  # noqa: SLF001
    wb._properties_dirty = True  # noqa: SLF001


def get_custom_doc_props(wb: Any) -> Any:
    """Return workbook custom document properties, lazily loaded."""
    if wb._custom_doc_props_cache is not None:  # noqa: SLF001
        return wb._custom_doc_props_cache  # noqa: SLF001
    payload = _reader_workbook_payload(wb, "read_custom_doc_properties")
    props = _custom_doc_props_from_payload(payload)
    props._attach_workbook(wb)  # noqa: SLF001
    wb._custom_doc_props_cache = props  # noqa: SLF001
    return wb._custom_doc_props_cache  # noqa: SLF001


def set_custom_doc_props(wb: Any, value: Any) -> None:
    """Replace workbook custom document properties."""
    from wolfxl.packaging.custom import CustomPropertyList

    if not isinstance(value, CustomPropertyList):
        raise TypeError(
            "custom_doc_props must be a CustomPropertyList, "
            f"got {type(value).__name__}"
        )
    value._attach_workbook(wb)  # noqa: SLF001
    wb._custom_doc_props_cache = value  # noqa: SLF001
    wb._custom_doc_props_dirty = True  # noqa: SLF001


def get_defined_names(wb: Any) -> Any:
    """Return a lazy-loaded workbook ``DefinedNameDict``.

    Args:
        wb: Workbook-like object carrying reader and defined-name cache state.

    Returns:
        A ``DefinedNameDict`` attached back to ``wb`` so user writes queue for
        save-time flushing.
    """
    if wb._defined_names_cache is not None:  # noqa: SLF001
        return wb._defined_names_cache  # noqa: SLF001
    from wolfxl.workbook import DefinedNameDict
    from wolfxl.workbook.defined_name import DefinedName

    dnd = DefinedNameDict()
    if wb._rust_reader is not None:  # noqa: SLF001
        seen: set[str] = set()
        for sheet_name in wb._sheet_names:  # noqa: SLF001
            try:
                entries = wb._rust_reader.read_named_ranges(sheet_name)  # noqa: SLF001
            except Exception:
                continue
            for entry in entries:
                name = entry["name"]
                scope = entry.get("scope", "workbook")
                if scope == "sheet":
                    continue
                if name in seen:
                    continue
                seen.add(name)
                refers_to = entry["refers_to"]
                if refers_to.startswith("="):
                    refers_to = refers_to[1:]
                dn = DefinedName(
                    name=name,
                    value=refers_to,
                    localSheetId=None,
                    comment=entry.get("comment"),
                    hidden=bool(entry.get("hidden", False)),
                    customMenu=entry.get("custom_menu"),
                    description=entry.get("description"),
                    help=entry.get("help"),
                    statusBar=entry.get("status_bar"),
                    shortcutKey=entry.get("shortcut_key"),
                    function=entry.get("function"),
                    functionGroupId=entry.get("function_group_id"),
                    vbProcedure=entry.get("vb_procedure"),
                    xlm=entry.get("xlm"),
                    publishToServer=entry.get("publish_to_server"),
                    workbookParameter=entry.get("workbook_parameter"),
                )
                dict.__setitem__(dnd, name, dn)
    dnd._wb = wb  # noqa: SLF001
    wb._defined_names_cache = dnd  # noqa: SLF001
    return dnd


def queue_worksheet_defined_names(wb: Any) -> None:
    """Move names added via ``ws.defined_names`` into the workbook save queue.

    Like openpyxl, each name's ``localSheetId`` is the worksheet's position
    at save time, so sheet moves and removals before save are honoured.
    """
    for sheet_idx, sheet_name in enumerate(wb._sheet_names):  # noqa: SLF001
        ws = wb._sheets.get(sheet_name)  # noqa: SLF001
        pending = getattr(ws, "_pending_defined_names", None)
        if not pending:
            continue
        for defined_name in pending.values():
            defined_name.localSheetId = sheet_idx
            wb._pending_defined_names[(defined_name.name, sheet_idx)] = defined_name  # noqa: SLF001
        pending.clear()


def get_workbook_properties(wb: Any) -> Any:
    """Return workbook-level properties, lazily loaded from the reader."""
    if wb._workbook_properties_cache is not None:  # noqa: SLF001
        return wb._workbook_properties_cache  # noqa: SLF001
    payload = _reader_workbook_payload(wb, "read_workbook_properties")
    wb._workbook_properties_cache = _workbook_properties_from_payload(payload)  # noqa: SLF001
    return wb._workbook_properties_cache  # noqa: SLF001


def set_workbook_properties(wb: Any, value: Any) -> None:
    """Replace workbook-level properties."""
    from wolfxl.workbook.properties import WorkbookProperties

    if not isinstance(value, WorkbookProperties):
        raise TypeError(
            "workbook_properties must be a WorkbookProperties, "
            f"got {type(value).__name__}"
        )
    wb._workbook_properties_cache = value  # noqa: SLF001


def get_calc_properties(wb: Any) -> Any:
    """Return workbook calculation properties, lazily loaded from the reader."""
    if wb._calc_properties_cache is not None:  # noqa: SLF001
        return wb._calc_properties_cache  # noqa: SLF001
    payload = _reader_workbook_payload(wb, "read_calc_properties")
    wb._calc_properties_cache = _calc_properties_from_payload(payload)  # noqa: SLF001
    return wb._calc_properties_cache  # noqa: SLF001


def set_calc_properties(wb: Any, value: Any) -> None:
    """Replace workbook calculation properties."""
    from wolfxl.workbook.properties import CalcProperties

    if not isinstance(value, CalcProperties):
        raise TypeError(
            f"calculation must be a CalcProperties, got {type(value).__name__}"
        )
    wb._calc_properties_cache = value  # noqa: SLF001


def get_views(wb: Any) -> list[Any]:
    """Return workbook window views, lazily loaded from the reader."""
    if wb._views_cache is not None:  # noqa: SLF001
        return wb._views_cache  # noqa: SLF001
    payload = _reader_workbook_payload(wb, "read_workbook_views")
    wb._views_cache = _book_views_from_payload(payload)  # noqa: SLF001
    return wb._views_cache  # noqa: SLF001


def set_views(wb: Any, value: Any) -> None:
    """Replace workbook window views."""
    wb._views_cache = list(value)  # noqa: SLF001


def set_security(wb: Any, value: Any) -> None:
    """Set workbook protection metadata.

    Args:
        wb: Workbook-like object carrying security state.
        value: ``WorkbookProtection`` instance, or ``None``.

    Raises:
        TypeError: If ``value`` has the wrong type.
    """
    from wolfxl.workbook.protection import WorkbookProtection

    if value is not None and not isinstance(value, WorkbookProtection):
        raise TypeError(
            "wb.security must be a WorkbookProtection or None, "
            f"got {type(value).__name__}"
        )
    wb._security = value  # noqa: SLF001
    wb._security_loaded = True  # noqa: SLF001
    wb._pending_security_update = True  # noqa: SLF001


def get_security(wb: Any) -> Any:
    """Return workbook protection metadata, lazily loaded from the reader."""
    _ensure_security_loaded(wb)
    return wb._security  # noqa: SLF001


def set_file_sharing(wb: Any, value: Any) -> None:
    """Set workbook file-sharing metadata.

    Args:
        wb: Workbook-like object carrying file-sharing state.
        value: ``FileSharing`` instance, or ``None``.

    Raises:
        TypeError: If ``value`` has the wrong type.
    """
    from wolfxl.workbook.protection import FileSharing

    if value is not None and not isinstance(value, FileSharing):
        raise TypeError(
            "wb.fileSharing must be a FileSharing or None, "
            f"got {type(value).__name__}"
        )
    wb._file_sharing = value  # noqa: SLF001
    wb._security_loaded = True  # noqa: SLF001
    wb._pending_security_update = True  # noqa: SLF001


def get_file_sharing(wb: Any) -> Any:
    """Return workbook file-sharing metadata, lazily loaded from the reader."""
    _ensure_security_loaded(wb)
    return wb._file_sharing  # noqa: SLF001


def _reader_workbook_payload(wb: Any, method_name: str) -> Any:
    reader = getattr(wb, "_rust_reader", None)
    if reader is None or not hasattr(reader, method_name):
        return None
    try:
        return getattr(reader, method_name)()
    except Exception:
        return None


def _custom_doc_props_from_payload(payload: Any) -> Any:
    from wolfxl.packaging.custom import CustomPropertyList

    props = CustomPropertyList()
    if not isinstance(payload, list):
        return props
    for item in payload:
        if isinstance(item, dict):
            prop = _custom_doc_prop_from_payload(item)
            if prop is not None:
                props.append(prop)
    return props


def _custom_doc_prop_from_payload(payload: dict[str, Any]) -> Any:
    from wolfxl.packaging.custom import (
        BoolProperty,
        DateTimeProperty,
        FloatProperty,
        IntProperty,
        LinkProperty,
        StringProperty,
    )

    name = str(payload.get("name") or "")
    if not name:
        return None
    kind = str(payload.get("kind") or "string")
    value = str(payload.get("value") or "")
    if kind == "int":
        return IntProperty(name=name, value=int(value))
    if kind == "float":
        return FloatProperty(name=name, value=float(value))
    if kind == "bool":
        return BoolProperty(name=name, value=value not in {"0", "false", "False"})
    if kind == "datetime":
        return DateTimeProperty(name=name, value=_parse_custom_datetime(value))
    if kind == "link":
        return LinkProperty(name=name, value=value)
    return StringProperty(name=name, value=value)


def _parse_custom_datetime(value: str) -> Any:
    from datetime import datetime

    try:
        return datetime.fromisoformat(value.replace("Z", "+00:00")).replace(tzinfo=None)
    except ValueError:
        return datetime.fromisoformat(value.removesuffix("Z"))


def _workbook_properties_from_payload(payload: Any) -> Any:
    from wolfxl.workbook.properties import WorkbookProperties

    if not isinstance(payload, dict):
        return WorkbookProperties()
    return WorkbookProperties(
        date1904=bool(payload.get("date1904", False)),
        dateCompatibility=_payload_value(payload, "date_compatibility", True),
        showObjects=_payload_value(payload, "show_objects", "all"),
        showBorderUnselectedTables=_payload_value(
            payload, "show_border_unselected_tables", True
        ),
        filterPrivacy=_payload_value(payload, "filter_privacy", False),
        promptedSolutions=_payload_value(payload, "prompted_solutions", False),
        showInkAnnotation=_payload_value(payload, "show_ink_annotation", True),
        backupFile=_payload_value(payload, "backup_file", False),
        saveExternalLinkValues=_payload_value(payload, "save_external_link_values", True),
        updateLinks=_payload_value(payload, "update_links", "userSet"),
        codeName=payload.get("code_name"),
        hidePivotFieldList=_payload_value(payload, "hide_pivot_field_list", False),
        showPivotChartFilter=_payload_value(payload, "show_pivot_chart_filter", False),
        allowRefreshQuery=_payload_value(payload, "allow_refresh_query", False),
        publishItems=_payload_value(payload, "publish_items", False),
        checkCompatibility=_payload_value(payload, "check_compatibility", False),
        autoCompressPictures=_payload_value(payload, "auto_compress_pictures", True),
        refreshAllConnections=_payload_value(payload, "refresh_all_connections", False),
        defaultThemeVersion=_payload_value(payload, "default_theme_version", 124226),
    )


def _calc_properties_from_payload(payload: Any) -> Any:
    from wolfxl.workbook.properties import CalcProperties

    if not isinstance(payload, dict):
        return CalcProperties()
    return CalcProperties(
        calcId=_payload_value(payload, "calc_id", 124519),
        calcMode=_payload_value(payload, "calc_mode", "auto"),
        fullCalcOnLoad=_payload_value(payload, "full_calc_on_load", False),
        refMode=_payload_value(payload, "ref_mode", "A1"),
        iterate=_payload_value(payload, "iterate", False),
        iterateCount=_payload_value(payload, "iterate_count", 100),
        iterateDelta=_payload_value(payload, "iterate_delta", 0.001),
        fullPrecision=_payload_value(payload, "full_precision", True),
        calcCompleted=_payload_value(payload, "calc_completed", True),
        calcOnSave=_payload_value(payload, "calc_on_save", True),
        concurrentCalc=_payload_value(payload, "concurrent_calc", True),
        concurrentManualCount=payload.get("concurrent_manual_count"),
        forceFullCalc=_payload_value(payload, "force_full_calc", False),
    )


def _book_views_from_payload(payload: Any) -> list[Any]:
    from wolfxl.workbook.views import BookView

    if not isinstance(payload, list):
        return [BookView()]
    views = [
        _book_view_from_payload(item)
        for item in payload
        if isinstance(item, dict)
    ]
    return views or [BookView()]


def _book_view_from_payload(payload: dict[str, Any]) -> Any:
    from wolfxl.workbook.views import BookView

    return BookView(
        visibility=str(payload.get("visibility") or "visible"),
        minimized=bool(payload.get("minimized", False)),
        showHorizontalScroll=bool(payload.get("show_horizontal_scroll", True)),
        showVerticalScroll=bool(payload.get("show_vertical_scroll", True)),
        showSheetTabs=bool(payload.get("show_sheet_tabs", True)),
        xWindow=payload.get("x_window"),
        yWindow=payload.get("y_window"),
        windowWidth=payload.get("window_width"),
        windowHeight=payload.get("window_height"),
        tabRatio=int(payload.get("tab_ratio", 600)),
        firstSheet=int(payload.get("first_sheet", 0)),
        activeTab=int(payload.get("active_tab", 0)),
        autoFilterDateGrouping=bool(
            payload.get("auto_filter_date_grouping", True)
        ),
    )


def _payload_value(payload: dict[str, Any], key: str, default: Any) -> Any:
    value = payload.get(key)
    return default if value is None else value


def _ensure_security_loaded(wb: Any) -> None:
    if getattr(wb, "_security_loaded", False):
        return
    wb._security_loaded = True  # noqa: SLF001
    reader = getattr(wb, "_rust_reader", None)
    if reader is None or not hasattr(reader, "read_workbook_security"):
        return
    try:
        payload = reader.read_workbook_security()
    except Exception:
        return
    if not isinstance(payload, dict):
        return
    wb._security = _workbook_protection_from_payload(  # noqa: SLF001
        payload.get("workbook_protection")
    )
    wb._file_sharing = _file_sharing_from_payload(payload.get("file_sharing"))  # noqa: SLF001


def _workbook_protection_from_payload(payload: Any) -> Any:
    if not isinstance(payload, dict):
        return None
    from wolfxl.workbook.protection import WorkbookProtection

    protection = WorkbookProtection(
        lock_structure=bool(payload.get("lock_structure", False)),
        lock_windows=bool(payload.get("lock_windows", False)),
        lock_revision=bool(payload.get("lock_revision", False)),
    )
    protection.workbook_password = payload.get("workbook_password")
    protection.workbook_password_character_set = payload.get("workbook_password_character_set")
    protection.workbook_algorithm_name = payload.get("workbook_algorithm_name")
    protection.workbook_hash_value = payload.get("workbook_hash_value")
    protection.workbook_salt_value = payload.get("workbook_salt_value")
    protection.workbook_spin_count = payload.get("workbook_spin_count")
    protection.revisions_password = payload.get("revisions_password")
    protection.revisions_password_character_set = payload.get("revisions_password_character_set")
    protection.revisions_algorithm_name = payload.get("revisions_algorithm_name")
    protection.revisions_hash_value = payload.get("revisions_hash_value")
    protection.revisions_salt_value = payload.get("revisions_salt_value")
    protection.revisions_spin_count = payload.get("revisions_spin_count")
    return protection


def _file_sharing_from_payload(payload: Any) -> Any:
    if not isinstance(payload, dict):
        return None
    from wolfxl.workbook.protection import FileSharing

    return FileSharing(
        read_only_recommended=bool(payload.get("read_only_recommended", False)),
        user_name=payload.get("user_name"),
        reservation_password=payload.get("reservation_password"),
        algorithm_name=payload.get("algorithm_name"),
        hash_value=payload.get("hash_value"),
        salt_value=payload.get("salt_value"),
        spin_count=payload.get("spin_count"),
    )
