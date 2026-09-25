"""Modify-mode workbook flush helpers for the Rust patcher."""

from __future__ import annotations

from typing import Any


def flush_pending_hyperlinks_to_patcher(wb: Any) -> None:
    """Drain pending worksheet hyperlinks into the Rust patcher."""
    patcher = wb._rust_patcher  # noqa: SLF001
    if patcher is None:
        return
    for ws in wb._sheets.values():  # noqa: SLF001
        pending = ws._pending_hyperlinks  # noqa: SLF001
        if not pending:
            continue
        for coord, hl in pending.items():
            if hl is None:
                patcher.queue_hyperlink_delete(ws.title, coord)
                continue
            patcher.queue_hyperlink(ws.title, coord, _hyperlink_payload(hl))
        pending.clear()


def _hyperlink_payload(hyperlink: Any) -> dict[str, Any]:
    """Build the Rust patcher payload for a worksheet hyperlink."""
    payload: dict[str, Any] = {}
    if hyperlink.target is not None:
        payload["target"] = hyperlink.target
    if hyperlink.location is not None:
        payload["location"] = hyperlink.location
    if hyperlink.tooltip is not None:
        payload["tooltip"] = hyperlink.tooltip
    if hyperlink.display is not None:
        payload["display"] = hyperlink.display
    return payload


def flush_pending_tables_to_patcher(wb: Any) -> None:
    """Drain pending worksheet tables into the Rust patcher."""
    patcher = wb._rust_patcher  # noqa: SLF001
    if patcher is None:
        return
    for ws in wb._sheets.values():  # noqa: SLF001
        pending = ws._pending_tables  # noqa: SLF001
        for table in pending:
            patcher.queue_table(ws.title, _table_payload(table))
        pending.clear()
        loaded_tables = getattr(ws._tables_cache, "loaded", None)  # noqa: SLF001
        if loaded_tables:
            _queue_loaded_table_growth(patcher, ws.title, loaded_tables)


def _queue_loaded_table_growth(
    patcher: Any,
    sheet: str,
    loaded_tables: dict[str, tuple[Any, str, tuple[str, ...]]],
) -> None:
    """Queue growth of tables read from the source file.

    A loaded table may grow down and to the right: its top-left cell and
    existing columns stay fixed, and each new column is appended to
    ``tableColumns``. The recorded state then advances so a later save
    queues only further growth.
    """
    from wolfxl.utils.cell import get_column_letter, range_boundaries

    for name, (table, saved_ref, saved_columns) in loaded_tables.items():
        columns = tuple(column.name for column in table.tableColumns)
        if table.ref == saved_ref and columns == saved_columns:
            continue
        min_col, min_row, max_col, max_row = range_boundaries(table.ref)
        saved_min_col, saved_min_row, _, _ = range_boundaries(saved_ref)
        if (min_col, min_row) != (saved_min_col, saved_min_row):
            raise ValueError(
                f"table '{name}' on '{sheet}': moving a loaded table's top-left "
                f"cell is not supported (was {saved_ref}, now {table.ref})"
            )
        if columns[: len(saved_columns)] != saved_columns:
            raise ValueError(
                f"table '{name}' on '{sheet}': existing columns must stay as loaded; "
                "append new TableColumn entries after them"
            )
        width = max_col - min_col + 1
        if width != len(columns):
            raise ValueError(
                f"table '{name}' on '{sheet}': range {table.ref} spans {width} "
                f"columns but the table lists {len(columns)}; append one "
                "TableColumn per new column"
            )
        filter_max_row = max_row - (table.totalsRowCount or 0)
        auto_filter_ref = (
            f"{get_column_letter(min_col)}{min_row}:"
            f"{get_column_letter(max_col)}{filter_max_row}"
        )
        patcher.queue_table_growth(
            sheet,
            name,
            table.ref,
            auto_filter_ref,
            [str(column) for column in columns[len(saved_columns) :]],
        )
        loaded_tables[name] = (table, table.ref, columns)


def _table_payload(table: Any) -> dict[str, Any]:
    """Build the Rust patcher payload for a worksheet table."""
    payload: dict[str, Any] = {
        "name": table.name,
        "ref": table.ref,
        "columns": (
            [column.name for column in table.tableColumns]
            if table.tableColumns
            else []
        ),
        "header_row_count": int(table.headerRowCount or 0),
        "totals_row_shown": bool(
            table.totalsRowCount and table.totalsRowCount > 0
        ),
        "autofilter": True,
    }
    if table.displayName and table.displayName != table.name:
        payload["display_name"] = table.displayName
    if table.tableStyleInfo is not None and table.tableStyleInfo.name:
        payload["style"] = {
            "name": table.tableStyleInfo.name,
            "show_first_column": bool(table.tableStyleInfo.showFirstColumn),
            "show_last_column": bool(table.tableStyleInfo.showLastColumn),
            "show_row_stripes": bool(table.tableStyleInfo.showRowStripes),
            "show_column_stripes": bool(table.tableStyleInfo.showColumnStripes),
        }
    return payload


def flush_pending_images_to_patcher(wb: Any) -> None:
    """Drain pending worksheet images into the Rust patcher."""
    from wolfxl._images import image_to_writer_payload
    from wolfxl._worksheet_media import pop_pending_image_deletions

    patcher = wb._rust_patcher  # noqa: SLF001
    if patcher is None:
        return
    for ws in wb._sheets.values():  # noqa: SLF001
        pending_deletions = pop_pending_image_deletions(ws)
        pending = ws._pending_images  # noqa: SLF001
        if not pending_deletions and not pending:
            continue
        for image_index in pending_deletions:
            patcher.queue_image_remove(ws.title, image_index)
        for image in pending:
            patcher.queue_image_add(ws.title, image_to_writer_payload(image))
        pending.clear()


def flush_pending_comments_to_patcher(wb: Any) -> None:
    """Drain pending worksheet comments into the Rust patcher."""
    patcher = wb._rust_patcher  # noqa: SLF001
    if patcher is None:
        return
    for ws in wb._sheets.values():  # noqa: SLF001
        pending = ws._pending_comments  # noqa: SLF001
        if not pending:
            continue
        for coord, comment in pending.items():
            if comment is None:
                patcher.queue_comment_delete(ws.title, coord)
                continue
            patcher.queue_comment(ws.title, coord, _comment_payload(comment))
        pending.clear()


def _comment_payload(comment: Any) -> dict[str, Any]:
    """Build the Rust patcher payload for a worksheet comment."""
    payload: dict[str, Any] = {
        "text": comment.text,
        "author": comment.author or "wolfxl",
    }
    if getattr(comment, "width", None) is not None:
        payload["width_pt"] = float(comment.width)
    if getattr(comment, "height", None) is not None:
        payload["height_pt"] = float(comment.height)
    return payload


def flush_pending_threaded_comments_to_patcher(wb: Any) -> None:
    """Drain pending worksheet threaded comments into the Rust patcher.

    Mirrors :func:`flush_pending_comments_to_patcher` but bundles the
    full top + replies payload per coordinate into one
    ``queue_threaded_comment`` call. ``None`` sentinel routes through
    ``queue_threaded_comment_delete``.

    Workbook-scope persons flush separately via
    :func:`flush_pending_persons_to_patcher` so the patcher can keep the
    personList registry in lockstep with the threaded payload.
    """
    patcher = wb._rust_patcher  # noqa: SLF001
    if patcher is None:
        return
    if not hasattr(patcher, "queue_threaded_comment"):
        return
    for ws in wb._sheets.values():  # noqa: SLF001
        pending = ws._pending_threaded_comments  # noqa: SLF001
        if not pending:
            continue
        for coord, top in pending.items():
            if top is None:
                patcher.queue_threaded_comment_delete(ws.title, coord)
                continue
            top.ensure_id()
            top.ensure_created()
            replies_payload: list[dict[str, Any]] = []
            for reply in top.replies:
                reply.ensure_id()
                reply.ensure_created()
                replies_payload.append(
                    _threaded_comment_payload(coord, reply, parent_id=top.id)
                )
            patcher.queue_threaded_comment(
                ws.title,
                coord,
                {
                    "top": _threaded_comment_payload(coord, top, parent_id=None),
                    "replies": replies_payload,
                },
            )
        pending.clear()


def flush_pending_persons_to_patcher(wb: Any) -> None:
    """Drain the workbook-scope person registry into the Rust patcher.

    The patcher's ``queue_person`` is idempotent on ``id`` so re-running
    the flush is safe; calling it lets the modify-mode personList stay
    in lockstep with any newly-added :class:`Person` instances.
    """
    patcher = wb._rust_patcher  # noqa: SLF001
    if patcher is None or not hasattr(patcher, "queue_person"):
        return
    registry = getattr(wb, "_persons_registry", None)  # noqa: SLF001
    if registry is None or len(registry) == 0:
        return
    for person in registry:
        if person.id is None:
            continue
        patcher.queue_person(
            {
                "id": person.id,
                "name": person.name,
                "user_id": person.user_id,
                "provider_id": person.provider_id,
            }
        )


def _threaded_comment_payload(
    coord: str, tc: Any, *, parent_id: str | None
) -> dict[str, Any]:
    """Build the patcher payload for one ThreadedComment.

    Mirrors :func:`wolfxl._worksheet_flush._threaded_comment_payload` so
    write-mode and modify-mode hand the Rust side identical shapes.
    """
    person_id = tc.person.id if tc.person is not None else ""
    return {
        "id": tc.id,
        "cell": coord,
        "person_id": person_id,
        "created": tc.created.isoformat(timespec="milliseconds"),
        "parent_id": parent_id,
        "text": tc.text,
        "done": tc.done,
    }


def flush_pending_data_validations_to_patcher(wb: Any) -> None:
    """Drain pending worksheet data validations into the Rust patcher."""
    from wolfxl.worksheet.datavalidation import _dv_to_patcher_dict

    patcher = wb._rust_patcher  # noqa: SLF001
    if patcher is None:
        return
    for ws in wb._sheets.values():  # noqa: SLF001
        pending = ws._pending_data_validations  # noqa: SLF001
        if not pending:
            continue
        for data_validation in pending:
            patcher.queue_data_validation(
                ws.title,
                _dv_to_patcher_dict(data_validation),
            )
        pending.clear()


def flush_pending_conditional_formats_to_patcher(wb: Any) -> None:
    """Drain pending worksheet conditional formatting into the Rust patcher."""
    from wolfxl.formatting import _cf_to_patcher_dict

    patcher = wb._rust_patcher  # noqa: SLF001
    if patcher is None:
        return
    for ws in wb._sheets.values():  # noqa: SLF001
        pending = ws._pending_conditional_formats  # noqa: SLF001
        if not pending:
            continue
        by_sqref: dict[str, list[Any]] = {}
        order: list[str] = []
        for sqref, rule in pending:
            if sqref not in by_sqref:
                by_sqref[sqref] = []
                order.append(sqref)
            by_sqref[sqref].append(rule)
        for sqref in order:
            patcher.queue_conditional_formatting(
                ws.title,
                _cf_to_patcher_dict(sqref, by_sqref[sqref]),
            )
        pending.clear()


def flush_pending_axis_shifts_to_patcher(wb: Any) -> None:
    """Drain pending row/column structural shifts into the Rust patcher."""
    patcher = wb._rust_patcher  # noqa: SLF001
    if patcher is None or not wb._pending_axis_shifts:  # noqa: SLF001
        return
    for sheet_title, axis, idx, n in wb._pending_axis_shifts:  # noqa: SLF001
        patcher.queue_axis_shift(sheet_title, axis, idx, n)
    wb._pending_axis_shifts.clear()  # noqa: SLF001


def flush_pending_range_moves_to_patcher(wb: Any) -> None:
    """Drain pending range moves into the Rust patcher."""
    patcher = wb._rust_patcher  # noqa: SLF001
    if patcher is None or not wb._pending_range_moves:  # noqa: SLF001
        return
    for (
        sheet_title,
        src_min_col,
        src_min_row,
        src_max_col,
        src_max_row,
        d_row,
        d_col,
        translate,
    ) in wb._pending_range_moves:  # noqa: SLF001
        patcher.queue_range_move(
            sheet_title,
            src_min_col,
            src_min_row,
            src_max_col,
            src_max_row,
            d_row,
            d_col,
            translate,
        )
    wb._pending_range_moves.clear()  # noqa: SLF001


def flush_pending_sheet_copies_to_patcher(wb: Any) -> None:
    """Drain pending sheet-copy operations into the Rust patcher."""
    patcher = wb._rust_patcher  # noqa: SLF001
    if patcher is None or not wb._pending_sheet_copies:  # noqa: SLF001
        return
    for src_title, dst_title, deep_copy_images in wb._pending_sheet_copies:  # noqa: SLF001
        patcher.queue_sheet_copy(src_title, dst_title, deep_copy_images)
    wb._pending_sheet_copies.clear()  # noqa: SLF001


def queue_sheet_move_to_patcher(wb: Any, name: str, offset: int) -> None:
    """Queue one sheet move into the Rust patcher."""
    patcher = wb._rust_patcher  # noqa: SLF001
    if patcher is None:
        return
    patcher.queue_sheet_move(name, offset)


def flush_defined_names_to_patcher(wb: Any) -> None:
    """Drain pending workbook defined-name writes into the Rust patcher."""
    patcher = wb._rust_patcher  # noqa: SLF001
    if patcher is None:
        return
    _queue_print_titles_to_patcher(wb, patcher)
    if not wb._pending_defined_names:  # noqa: SLF001
        return
    for defined_name in wb._pending_defined_names.values():  # noqa: SLF001
        patcher.queue_defined_name(_defined_name_payload(defined_name))
    wb._pending_defined_names.clear()  # noqa: SLF001


def _queue_print_titles_to_patcher(wb: Any, patcher: Any) -> None:
    """Queue worksheet repeat-row/column titles as reserved defined names."""
    from wolfxl.worksheet.print_settings import ColRange, PrintTitles, RowRange

    for sheet_idx, sheet_name in enumerate(wb._sheet_names):  # noqa: SLF001
        ws = wb._sheets.get(sheet_name)  # noqa: SLF001
        if ws is None:
            continue
        rows = getattr(ws, "_print_title_rows", None)
        cols = getattr(ws, "_print_title_cols", None)
        if rows is None and cols is None:
            if getattr(ws, "_print_titles_dirty", False):
                patcher.queue_defined_name(
                    {
                        "name": "_xlnm.Print_Titles",
                        "local_sheet_id": sheet_idx,
                        "delete": True,
                    }
                )
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
        patcher.queue_defined_name(
            {
                "name": "_xlnm.Print_Titles",
                "formula": formula,
                "local_sheet_id": sheet_idx,
            }
        )


def _has_pending_print_titles_for_sheet(wb: Any, sheet_idx: int) -> bool:
    """Return whether user-defined names already carry this sheet's titles."""
    return any(
        defined_name.name == "_xlnm.Print_Titles"
        and defined_name.localSheetId == sheet_idx
        for defined_name in wb._pending_defined_names.values()  # noqa: SLF001
    )


def _defined_name_payload(defined_name: Any) -> dict[str, Any]:
    """Build the Rust patcher payload for a workbook defined name.

    Phase 2 (G22) — full ECMA-376 §18.2.5 attribute surface. ``None``
    fields are omitted so the Rust patcher's ``serialize_upsert_over_existing``
    treats them as "preserve source"; an explicit ``False`` for a
    ``true``-emitted bool attr (e.g. ``vbProcedure``) drives a remove on
    upsert, matching openpyxl's "absent attr means default".
    """
    if getattr(defined_name, "delete", False):
        payload = {
            "name": defined_name.name,
            "delete": True,
        }
        if defined_name.localSheetId is not None:
            payload["local_sheet_id"] = defined_name.localSheetId
        return payload

    payload: dict[str, Any] = {
        "name": defined_name.name,
        "formula": defined_name.value,
    }
    if defined_name.localSheetId is not None:
        payload["local_sheet_id"] = defined_name.localSheetId
    # `hidden` is always a `bool` on the Python class (default False, never
    # None), so send it explicitly. The patcher upsert path treats
    # `Some(false)` as "remove the attr from the source XML" — without an
    # explicit False here, an existing `hidden="1"` would survive a
    # `dn.hidden = False` modify-mode write.
    payload["hidden"] = bool(defined_name.hidden)
    if defined_name.comment is not None:
        payload["comment"] = defined_name.comment
    for src_attr, payload_key in (
        ("custom_menu", "custom_menu"),
        ("description", "description"),
        ("help", "help"),
        ("status_bar", "status_bar"),
        ("shortcut_key", "shortcut_key"),
        ("function", "function"),
        ("function_group_id", "function_group_id"),
        ("vb_procedure", "vb_procedure"),
        ("xlm", "xlm"),
        ("publish_to_server", "publish_to_server"),
        ("workbook_parameter", "workbook_parameter"),
    ):
        v = getattr(defined_name, src_attr, None)
        if v is not None:
            payload[payload_key] = v
    return payload


def build_security_dict(wb: Any) -> dict[str, Any]:
    """Build the Rust patcher payload for workbook security blocks."""
    return {
        "workbook_protection": (
            wb._security.to_dict() if wb._security is not None else None  # noqa: SLF001
        ),
        "file_sharing": (
            wb._file_sharing.to_dict()  # noqa: SLF001
            if wb._file_sharing is not None  # noqa: SLF001
            else None
        ),
    }


def flush_security_to_patcher(wb: Any) -> None:
    """Drain workbook security metadata into the Rust patcher."""
    patcher = wb._rust_patcher  # noqa: SLF001
    if patcher is None or not wb._pending_security_update:  # noqa: SLF001
        return
    patcher.queue_workbook_security(build_security_dict(wb))
    wb._pending_security_update = False  # noqa: SLF001


def flush_properties_to_patcher(wb: Any) -> None:
    """Drain dirty workbook document properties into the Rust patcher."""
    patcher = wb._rust_patcher  # noqa: SLF001
    if patcher is None:
        return
    props = wb._properties_cache  # noqa: SLF001
    if props is None:
        wb._properties_dirty = False  # noqa: SLF001
        return

    patcher.queue_properties(_properties_payload(wb, props))
    wb._properties_dirty = False  # noqa: SLF001


def _properties_payload(wb: Any, props: Any) -> dict[str, Any]:
    """Build the Rust patcher payload for workbook document properties."""
    user_set: set[str] = getattr(props, "_user_set", set())
    modified_iso: str | None = None
    if "modified" in user_set and props.modified is not None:
        modified_iso = props.modified.isoformat()
    payload: dict[str, Any] = {
        "title": props.title,
        "subject": props.subject,
        "creator": props.creator,
        "keywords": props.keywords,
        "description": props.description,
        "last_modified_by": props.lastModifiedBy,
        "category": props.category,
        "content_status": props.contentStatus,
        "created_iso": props.created.isoformat() if props.created else None,
        "modified_iso": modified_iso,
        "sheet_names": list(wb._sheet_names),  # noqa: SLF001
    }
    return {key: value for key, value in payload.items() if value is not None}


def flush_pending_sheet_setup_to_patcher(wb: Any) -> None:
    """Drain pending worksheet setup metadata into the Rust patcher."""
    patcher = wb._rust_patcher  # noqa: SLF001
    if patcher is None:
        return
    for ws in wb._sheets.values():  # noqa: SLF001
        if not _has_sheet_setup_updates(ws):
            continue
        _queue_sheet_setup_update(patcher, ws)


def _has_sheet_setup_updates(ws: Any) -> bool:
    """Return whether a worksheet has pending sheet setup state."""
    return not (
        ws._page_setup is None  # noqa: SLF001
        and ws._page_margins is None  # noqa: SLF001
        and (
            ws._print_options is None  # noqa: SLF001
            or ws._print_options.is_default()  # noqa: SLF001
        )
        and ws._header_footer is None  # noqa: SLF001
        and ws._sheet_view is None  # noqa: SLF001
        and ws._protection is None  # noqa: SLF001
    )


def _sheet_setup_payload(ws: Any) -> dict[str, Any] | None:
    """Build the Rust patcher payload for worksheet setup metadata."""
    payload = ws.to_rust_setup_dict()
    if all(value is None for value in payload.values()):
        return None
    return payload


def _queue_sheet_setup_update(patcher: Any, ws: Any) -> None:
    """Queue one sheet-setup update into the Rust patcher."""
    payload = _sheet_setup_payload(ws)
    if payload is None:
        return
    patcher.queue_sheet_setup_update(ws.title, payload)


def flush_pending_page_breaks_to_patcher(wb: Any) -> None:
    """Drain pending page-break and sheet-format metadata into the Rust patcher."""
    patcher = wb._rust_patcher  # noqa: SLF001
    if patcher is None:
        return
    for ws in wb._sheets.values():  # noqa: SLF001
        if not _has_page_break_updates(ws):
            continue
        _queue_page_breaks_update(patcher, ws)


def _has_page_break_updates(ws: Any) -> bool:
    """Return whether a worksheet has pending page-break or format state."""
    row_breaks = getattr(ws, "_row_breaks", None)
    col_breaks = getattr(ws, "_col_breaks", None)
    sheet_format = getattr(ws, "_sheet_format", None)
    return row_breaks is not None or col_breaks is not None or sheet_format is not None


def _page_breaks_payload(ws: Any) -> dict[str, Any] | None:
    """Build the Rust patcher payload for page breaks and sheet format."""
    try:
        breaks_dict = ws.to_rust_page_breaks_dict()
        fmt_dict = ws.to_rust_sheet_format_dict()
    except Exception:
        return None
    payload = {
        "row_breaks": breaks_dict.get("row_breaks"),
        "col_breaks": breaks_dict.get("col_breaks"),
        "sheet_format": fmt_dict,
    }
    if all(value is None for value in payload.values()):
        return None
    return payload


def _queue_page_breaks_update(patcher: Any, ws: Any) -> None:
    """Queue one page-break update, preserving the historical skip-on-error path."""
    payload = _page_breaks_payload(ws)
    if payload is None:
        return
    try:
        patcher.queue_page_breaks_update(ws.title, payload)
    except Exception:
        return


def flush_pending_autofilters_to_patcher(wb: Any) -> None:
    """Drain pending worksheet autofilter state into the Rust patcher."""
    patcher = wb._rust_patcher  # noqa: SLF001
    if patcher is None:
        return
    for ws in wb._sheets.values():  # noqa: SLF001
        autofilter = ws._auto_filter  # noqa: SLF001
        if not _autofilter_has_state(autofilter):
            continue
        _queue_autofilter(patcher, ws.title, autofilter)


def _autofilter_has_state(autofilter: Any) -> bool:
    """Return whether an autofilter contains state worth queueing."""
    return (
        autofilter.ref is not None
        or bool(autofilter.filter_columns)
        or autofilter.sort_state is not None
    )


def _queue_autofilter(patcher: Any, sheet_title: str, autofilter: Any) -> None:
    """Queue one autofilter update, failing loudly on malformed state."""
    try:
        payload = autofilter.to_rust_dict()
    except Exception as exc:
        raise ValueError(
            f"failed to serialize autofilter for sheet {sheet_title!r}"
        ) from exc
    try:
        patcher.queue_autofilter(sheet_title, payload)
    except Exception as exc:
        raise RuntimeError(
            f"failed to queue autofilter for sheet {sheet_title!r}"
        ) from exc


def flush_pending_charts_to_patcher(wb: Any) -> None:
    """Drain pending chart additions into the Rust patcher."""
    patcher = wb._rust_patcher  # noqa: SLF001
    if patcher is None:
        return

    from wolfxl._worksheet_media import pop_pending_chart_deletions

    for ws in wb._sheets.values():  # noqa: SLF001
        for meta in pop_pending_chart_deletions(ws):
            patcher.queue_chart_remove(
                ws.title,
                meta["drawing_path"],
                meta["chart_rid"],
                meta["chart_path"],
            )
    _flush_pending_chart_bytes(wb, patcher)
    _flush_pending_chart_objects(wb, patcher)


def _flush_pending_chart_bytes(wb: Any, patcher: Any) -> None:
    """Drain pre-serialized chart XML payloads into the Rust patcher."""
    pending_bytes = getattr(wb, "_pending_chart_adds", None)
    if pending_bytes:
        for sheet_title, items in pending_bytes.items():
            if not items:
                continue
            for chart_xml, anchor_a1, width_emu, height_emu in items:
                patcher.queue_chart_add(
                    sheet_title,
                    chart_xml,
                    anchor_a1,
                    int(width_emu),
                    int(height_emu),
                )
        pending_bytes.clear()


def _flush_pending_chart_objects(wb: Any, patcher: Any) -> None:
    """Serialize queued chart objects and drain them into the Rust patcher."""
    any_pending = any(ws._pending_charts for ws in wb._sheets.values())  # noqa: SLF001
    if not any_pending:
        return
    try:
        from wolfxl._rust import serialize_chart_dict  # type: ignore[attr-defined]
    except ImportError as exc:  # pragma: no cover - defensive
        raise NotImplementedError(
            "Modify-mode high-level Worksheet.add_chart() requires "
            "Sprint Μ-prime Pod-α′'s serialize_chart_dict PyO3 export. "
            "Build the wolfxl wheel from a branch that includes the "
            "Pod-α′ commits, or fall back to "
            "Workbook.add_chart_modify_mode(sheet, chart_xml_bytes, anchor) "
            "with pre-serialised XML."
        ) from exc

    cm_to_emu = 360_000
    for ws in wb._sheets.values():  # noqa: SLF001
        pending_objs = ws._pending_charts  # noqa: SLF001
        if not pending_objs:
            continue
        for chart in pending_objs:
            chart_dict = chart.to_rust_dict()
            anchor = chart._anchor or "E15"  # noqa: SLF001
            chart_xml = serialize_chart_dict(chart_dict, anchor)
            width_emu, height_emu = _chart_size_emu(chart, cm_to_emu)
            patcher.queue_chart_add(
                ws.title,
                chart_xml,
                anchor,
                width_emu,
                height_emu,
            )
        pending_objs.clear()


def _chart_size_emu(chart: Any, cm_to_emu: int) -> tuple[int, int]:
    """Return chart width and height in EMUs for patcher queueing."""
    return int(chart.width * cm_to_emu), int(chart.height * cm_to_emu)


def flush_pending_slicers_to_patcher(wb: Any) -> None:
    """Drain pending slicer additions into the Rust patcher."""
    patcher = wb._rust_patcher  # noqa: SLF001
    if patcher is None:
        return
    for ws in wb._sheets.values():  # noqa: SLF001
        pending = getattr(ws, "_pending_slicers", None)
        if not pending:
            continue
        for slicer in pending:
            _queue_slicer_add(patcher, ws.title, slicer)
        ws._pending_slicers = []  # noqa: SLF001
    wb._pending_slicer_caches = []  # noqa: SLF001


def _queue_slicer_add(patcher: Any, sheet_title: str, slicer: Any) -> None:
    """Queue one slicer presentation, failing loudly on malformed state."""
    cache = slicer.cache
    try:
        cache_dict = cache.to_rust_dict()
    except Exception as exc:
        raise ValueError(
            f"failed to serialize slicer cache for sheet {sheet_title!r}"
        ) from exc
    try:
        slicer_dict = slicer.to_rust_dict()
    except Exception as exc:
        raise ValueError(
            f"failed to serialize slicer for sheet {sheet_title!r}"
        ) from exc
    try:
        patcher.queue_slicer_add(sheet_title, cache_dict, slicer_dict)
    except Exception as exc:
        raise RuntimeError(
            f"failed to queue slicer for sheet {sheet_title!r}"
        ) from exc


def flush_pending_pivots_to_patcher(wb: Any) -> None:
    """Drain pending pivot caches and tables into the Rust patcher."""
    patcher = wb._rust_patcher  # noqa: SLF001
    if patcher is None:
        return

    any_caches = bool(wb._pending_pivot_caches)  # noqa: SLF001
    any_tables = any(
        getattr(ws, "_pending_pivot_tables", None)
        for ws in wb._sheets.values()  # noqa: SLF001
    )
    if not any_caches and not any_tables:
        return

    serializers = _pivot_serializers()
    cache_dicts = _flush_pending_pivot_caches(wb, patcher, serializers)
    _flush_pending_pivot_tables(wb, patcher, serializers, cache_dicts)


def _pivot_serializers() -> tuple[Any, Any, Any]:
    """Load the Rust pivot serializer exports required for modify-mode saves."""
    try:
        from wolfxl._rust import (  # type: ignore[attr-defined]
            serialize_pivot_cache_dict,
            serialize_pivot_records_dict,
            serialize_pivot_table_dict,
        )
    except ImportError as exc:  # pragma: no cover - defensive
        raise NotImplementedError(
            "Modify-mode Workbook.add_pivot_cache() / "
            "Worksheet.add_pivot_table() require Sprint Ν Pod-γ's "
            "serialize_pivot_*_dict PyO3 exports. Build the wolfxl wheel "
            "from a branch that includes the Pod-γ commits."
        ) from exc
    return (
        serialize_pivot_cache_dict,
        serialize_pivot_records_dict,
        serialize_pivot_table_dict,
    )


def _flush_pending_pivot_caches(
    wb: Any,
    patcher: Any,
    serializers: tuple[Any, Any, Any],
) -> dict[int, dict[str, Any]]:
    """Drain pending pivot caches and return their definition dicts by cache id."""
    serialize_pivot_cache_dict, serialize_pivot_records_dict, _ = serializers
    cache_dicts: dict[int, dict[str, Any]] = {}
    for cache in wb._pending_pivot_caches:  # noqa: SLF001
        definition_dict = cache.to_rust_dict()
        records_dict = cache.to_rust_records_dict()
        definition_xml = serialize_pivot_cache_dict(definition_dict)
        records_xml = serialize_pivot_records_dict(definition_dict, records_dict)
        cache_dicts[int(cache._cache_id)] = definition_dict
        allocated = patcher.queue_pivot_cache_add(definition_xml, records_xml)
        if allocated != cache._cache_id:
            raise RuntimeError(
                f"Pivot cache id mismatch: python={cache._cache_id} "
                f"vs patcher={allocated}. This indicates a queue-ordering "
                f"bug in _flush_pending_pivots_to_patcher."
            )
    wb._pending_pivot_caches.clear()  # noqa: SLF001
    return cache_dicts


def _flush_pending_pivot_tables(
    wb: Any,
    patcher: Any,
    serializers: tuple[Any, Any, Any],
    cache_dicts: dict[int, dict[str, Any]],
) -> None:
    """Drain pending pivot tables into the Rust patcher."""
    _, _, serialize_pivot_table_dict = serializers
    for ws in wb._sheets.values():  # noqa: SLF001
        pending = getattr(ws, "_pending_pivot_tables", None)
        if not pending:
            continue
        for pivot_table in pending:
            if pivot_table.cache._cache_id is None:
                raise ValueError(
                    f"PivotTable on sheet {ws.title!r} references a PivotCache "
                    f"that was not registered via Workbook.add_pivot_cache() - "
                    f"register the cache before calling save()."
                )
            if hasattr(pivot_table, "_compute_layout"):
                pivot_table._compute_layout()
            table_dict = pivot_table.to_rust_dict()
            cache_id = int(pivot_table.cache._cache_id)
            cache_dict = cache_dicts.get(cache_id)
            if cache_dict is None:
                cache_dict = pivot_table.cache.to_rust_dict()
            table_xml = serialize_pivot_table_dict(cache_dict, table_dict)
            patcher.queue_pivot_table_add(
                ws.title,
                table_xml,
                cache_id,
            )
        pending.clear()


def flush_pending_pivot_source_edits_to_patcher(wb: Any) -> None:
    """G17 / RFC-070 — drain any dirty :class:`PivotTableHandle`s.

    For each worksheet, walk its lazily-materialised
    ``_pivot_handles_cache`` and register a source-range edit for
    every handle whose ``_dirty`` flag is set. The Rust patcher's
    ``apply_pivot_source_edits_phase`` (Phase 2.5m-edit) consumes the
    queue at save time and rewrites the linked
    ``pivotCacheDefinition*.xml`` parts.
    """
    patcher = wb._rust_patcher  # noqa: SLF001
    if patcher is None:
        return
    register = getattr(patcher, "register_pivot_source_edit", None)
    if register is None:
        # Older Rust extension predates G17 — nothing to do.
        return
    for ws in wb._sheets.values():  # noqa: SLF001
        cache = getattr(ws, "_pivot_handles_cache", None)
        if not cache:
            continue
        for handle in cache:
            if not getattr(handle, "_dirty", False):
                continue
            new_ref = handle._new_source_to_a1()  # noqa: SLF001
            new_sheet = handle._new_source_sheet_name()  # noqa: SLF001
            new_cols = handle._column_count()  # noqa: SLF001
            orig_cols = _orig_column_count(handle)
            force_refresh = orig_cols is not None and new_cols != orig_cols
            register(
                handle._cache_part_path,  # noqa: SLF001
                new_ref,
                new_sheet,
                bool(force_refresh),
            )
            # Reset dirty so a subsequent save() on the same workbook
            # doesn't double-register.
            handle._orig_source_range = new_ref  # noqa: SLF001
            handle._orig_source_sheet = new_sheet  # noqa: SLF001
            handle._orig_field_count = new_cols  # noqa: SLF001
            handle._new_source = None  # noqa: SLF001
            handle._dirty = False  # noqa: SLF001


def _orig_column_count(handle: Any) -> int | None:
    """Best-effort: derive the original source column count from the
    handle's parsed metadata. Returns ``None`` for unparseable
    ranges."""
    rng = getattr(handle, "_orig_source_range", "") or ""
    if ":" in rng:
        left, right = rng.split(":", 1)
    else:
        left = right = rng
    l_col = _leading_col(left)
    r_col = _leading_col(right)
    if l_col is None or r_col is None:
        return None
    return r_col - l_col + 1


def _leading_col(cell: str) -> int | None:
    s = cell.lstrip("$")
    col = 0
    seen = False
    for ch in s:
        if ch.isalpha():
            seen = True
            col = col * 26 + (ord(ch.upper()) - ord("A") + 1)
        else:
            break
    return col if seen else None
