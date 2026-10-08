"""Worksheet proxy — provides ``ws['A1']`` access and tracks dirty cells."""

from __future__ import annotations

import inspect as _inspect
import re
from collections.abc import Sequence
from collections.abc import Iterable, Iterator
from typing import TYPE_CHECKING, Any

from wolfxl._cell import Cell, ILLEGAL_CHARACTERS_RE, _UNSET, _illegal_string_message
from wolfxl.cell._merged import MergedCell
from wolfxl._worksheet_access import (
    get_col_tuple,
    get_item,
    get_rect,
    get_row_tuple,
    resolve_string_key,
)
from wolfxl._worksheet_bounds import (
    calculate_dimension as _calculate_dimension,
    max_col as _worksheet_max_col,
    max_row as _worksheet_max_row,
    read_dimension_bounds as _read_dimension_bounds,
    read_dimensions as _read_dimensions,
)
from wolfxl._worksheet_collections import AutoFilter as _AutoFilter
from wolfxl._worksheet_collections import MergedCellsProxy as _MergedCellsProxy
from wolfxl._worksheet_collections import merge_cells as _merge_cells
from wolfxl._worksheet_collections import unmerge_cells as _unmerge_cells
from wolfxl._worksheet_dimensions import ColumnDimensionProxy, RowDimensionProxy
from wolfxl._worksheet_features import (
    add_data_validation as _add_data_validation,
    add_table as _add_table,
    get_defined_names,
    get_comments_map,
    get_conditional_formatting,
    get_data_validations,
    get_hyperlinks_map,
    get_tables_map,
)
from wolfxl._worksheet_flush import (
    flush_autofilter_post_cells,
    flush_compat_properties,
    flush_worksheet,
)
from wolfxl._worksheet_iteration import (
    iter_cols as _iter_cols,
    iter_cols_bulk as _iter_cols_bulk,
    iter_rows as _iter_rows,
    iter_rows_bulk as _iter_rows_bulk,
)
from wolfxl._worksheet_media import (
    add_chart as _add_chart,
    add_image as _add_image,
    add_pivot_table as _add_pivot_table,
    get_charts as _get_charts,
    get_images as _get_images,
    add_slicer as _add_slicer,
    remove_chart as _remove_chart,
    remove_image as _remove_image,
    replace_chart as _replace_chart,
    replace_image as _replace_image,
    validate_a1_anchor as _validate_a1_anchor,
)
from wolfxl._worksheet_pending import collect_pending_overlay, pending_writes_bounds
from wolfxl._worksheet_patcher_flush import flush_to_patcher
from wolfxl._worksheet_records import (
    cached_formula_values as _cached_formula_values,
    classify_format as _classify_worksheet_format,
    canonical_data_type as _canonical_data_type,  # noqa: F401 - legacy import path
    iter_cell_records as _iter_cell_records,
    iter_cell_records_python as _iter_cell_records_python,
    schema as _infer_worksheet_schema,
    sheet_visibility as _sheet_visibility,
)
from wolfxl._worksheet_repr import worksheet_repr as _worksheet_repr
from wolfxl._worksheet_rich_text import cellrichtext_to_runs_payload
from wolfxl._worksheet_setup import (
    get_auto_filter,
    get_col_breaks,
    get_dimension_holder,
    get_freeze_panes,
    get_header_footer,
    get_page_margins,
    get_page_setup,
    get_print_options,
    get_protection,
    get_row_breaks,
    get_sheet_format,
    get_sheet_properties,
    get_sheet_view,
    set_freeze_panes,
    set_print_title_cols,
    set_print_title_rows,
    to_rust_page_breaks_dict as _to_rust_page_breaks_dict,
    to_rust_setup_dict as _to_rust_setup_dict,
    to_rust_sheet_format_dict as _to_rust_sheet_format_dict,
)
from wolfxl._worksheet_structural import (
    delete_cols as _delete_cols,
    delete_rows as _delete_rows,
    insert_cols as _insert_cols,
    insert_rows as _insert_rows,
    move_cell as _move_cell,
    move_range as _move_range,
)
from wolfxl._worksheet_state import initialize_worksheet_state
from wolfxl._worksheet_writer_flush import flush_to_writer
from wolfxl._worksheet_write_buffers import (
    append_row,
    batch_write_dicts,
    batch_write_style_ids,
    extract_non_batchable,
    materialize_append_buffer,
    materialize_bulk_writes,
    styled_rows_style_grid,
    write_rows as _write_rows,
)
from wolfxl._utils import a1_to_rowcol, rowcol_to_a1
from wolfxl.utils import quote_sheetname
from wolfxl.utils.cell import get_column_letter, range_boundaries
from wolfxl.utils.exceptions import IllegalCharacterError

if TYPE_CHECKING:
    from wolfxl._workbook import Workbook

_PLAIN_SCALAR_TYPES = (bool, int, float, str)
_PLAIN_SCALAR_EXACT_TYPES = frozenset(_PLAIN_SCALAR_TYPES)
_CELL_NEW = Cell.__new__
_ILLEGAL_CHARACTERS_SEARCH = ILLEGAL_CHARACTERS_RE.search


def _new_plain_scalar_cell(
    ws: Worksheet,
    row: int,
    col: int,
    value: Any,
    value_type: type,
    key: tuple[int, int],
    cells: dict[tuple[int, int], Cell],
) -> Cell:
    """Create a brand-new simple cell without default-then-overwrite work."""
    if value_type is str:
        if len(value) > 32767:
            value = value[:32767]
        if _ILLEGAL_CHARACTERS_SEARCH(value):
            raise IllegalCharacterError(_illegal_string_message(value))
    elif isinstance(value, str):
        if len(value) > 32767:
            value = value[:32767]
        if _ILLEGAL_CHARACTERS_SEARCH(value):
            raise IllegalCharacterError(_illegal_string_message(value))
    return _new_plain_scalar_cell_unchecked(ws, row, col, value, key, cells)


def _new_plain_scalar_cell_unchecked(
    ws: Worksheet,
    row: int,
    col: int,
    value: Any,
    key: tuple[int, int],
    cells: dict[tuple[int, int], Cell],
) -> Cell:
    """Create a brand-new simple cell after caller-side value validation."""
    cell = _CELL_NEW(Cell)
    cell._ws = ws  # noqa: SLF001
    cell._row = row  # noqa: SLF001
    cell._col = col  # noqa: SLF001
    cell._value = value  # noqa: SLF001
    cell._explicit_data_type = _UNSET  # noqa: SLF001
    cell._value_dirty = True  # noqa: SLF001
    cell._format_dirty = False  # noqa: SLF001
    cell._style_slots_ready = False  # noqa: SLF001
    cell._value_is_plain = True  # noqa: SLF001
    dirty_values = ws._dirty_values  # noqa: SLF001
    cells[key] = cell
    # Simple scalar writes only need the compact value map; complex edits
    # still use ``_dirty`` when Cell setters need type- or format-aware flushes.
    dirty_values[key] = value
    if row >= ws._next_append_row:  # noqa: SLF001
        ws._next_append_row = row + 1  # noqa: SLF001
    if col > ws._max_col_idx:  # noqa: SLF001
        ws._max_col_idx = col  # noqa: SLF001
    return cell


def _retarget_sheet_formula(refers_to: str, old: str, new: str) -> str:
    prefix = "=" if refers_to.startswith("=") else ""
    body = refers_to[1:] if prefix else refers_to
    quoted_old = f"{quote_sheetname(old)}!"
    quoted_new = f"{quote_sheetname(new)}!"
    updated = body.replace(quoted_old, quoted_new)
    if updated == body:
        old_token = re.escape(f"{old}!")
        updated = re.sub(rf"(?<![A-Za-z0-9_]){old_token}", quoted_new, body)
    return f"{prefix}{updated}"


def _merge_range_from_args(
    range_string: str | None,
    start_row: int | None,
    start_column: int | None,
    end_row: int | None,
    end_column: int | None,
) -> str:
    if range_string is not None:
        return range_string
    start = f"{get_column_letter(start_column)}{start_row}"
    end = f"{get_column_letter(end_column)}{end_row}"
    return f"{start}:{end}"


def _unique_detached_sheet_title(parent: Workbook, title: str | None) -> str:
    base = title or "Sheet"
    existing = set(getattr(parent, "_sheets", {}))
    if base not in existing:
        return base
    index = 1
    while f"{base}{index}" in existing:
        index += 1
    return f"{base}{index}"


def _pending_defined_name_for_sheet(
    wb: Workbook, name: str, sheet_idx: int
) -> object | None:
    for pending in wb._pending_defined_names.values():  # noqa: SLF001
        if (
            getattr(pending, "name", None) == name
            and getattr(pending, "localSheetId", None) == sheet_idx
        ):
            return pending
    return None


def _drop_pending_defined_name_aliases(
    wb: Workbook, name: str, sheet_idx: int, keep_key: object
) -> None:
    for key, pending in list(wb._pending_defined_names.items()):  # noqa: SLF001
        if key == keep_key:
            continue
        if (
            getattr(pending, "name", None) == name
            and getattr(pending, "localSheetId", None) == sheet_idx
        ):
            del wb._pending_defined_names[key]  # noqa: SLF001


def _queue_defined_name_sheet_rename(
    wb: Workbook, old: str, new: str, sheet_idx: int
) -> None:
    reader = getattr(wb, "_rust_reader", None)
    if reader is None or not hasattr(reader, "read_named_ranges"):
        return
    try:
        entries = reader.read_named_ranges(old)
    except Exception:
        return
    from wolfxl.workbook.defined_name import DefinedName

    queued_names: set[str] = set()
    for entry in entries:
        if entry.get("scope") != "sheet":
            continue
        name = entry["name"]
        pending = _pending_defined_name_for_sheet(wb, entry["name"], sheet_idx)
        refers_to = (
            getattr(pending, "value", None)
            if pending is not None
            else entry.get("refers_to")
        ) or ""
        retargeted = _retarget_sheet_formula(refers_to, old, new)
        if retargeted == refers_to:
            continue
        defined_name = DefinedName(
            name=name,
            value=retargeted[1:] if retargeted.startswith("=") else retargeted,
            localSheetId=sheet_idx,
            comment=(
                getattr(pending, "comment", None)
                if pending is not None
                else entry.get("comment")
            ),
            hidden=(
                bool(getattr(pending, "hidden", False))
                if pending is not None
                else bool(entry.get("hidden", False))
            ),
            customMenu=(
                getattr(pending, "custom_menu", None)
                if pending is not None
                else entry.get("custom_menu")
            ),
            description=(
                getattr(pending, "description", None)
                if pending is not None
                else entry.get("description")
            ),
            help=getattr(pending, "help", None) if pending is not None else entry.get("help"),
            statusBar=(
                getattr(pending, "status_bar", None)
                if pending is not None
                else entry.get("status_bar")
            ),
            shortcutKey=(
                getattr(pending, "shortcut_key", None)
                if pending is not None
                else entry.get("shortcut_key")
            ),
            function=(
                getattr(pending, "function", None)
                if pending is not None
                else entry.get("function")
            ),
            functionGroupId=(
                getattr(pending, "function_group_id", None)
                if pending is not None
                else entry.get("function_group_id")
            ),
            vbProcedure=(
                getattr(pending, "vb_procedure", None)
                if pending is not None
                else entry.get("vb_procedure")
            ),
            xlm=getattr(pending, "xlm", None) if pending is not None else entry.get("xlm"),
            publishToServer=(
                getattr(pending, "publish_to_server", None)
                if pending is not None
                else entry.get("publish_to_server")
            ),
            workbookParameter=(
                getattr(pending, "workbook_parameter", None)
                if pending is not None
                else entry.get("workbook_parameter")
            ),
        )
        keep_key = (defined_name.name, sheet_idx)
        wb._pending_defined_names[keep_key] = defined_name  # noqa: SLF001
        _drop_pending_defined_name_aliases(wb, defined_name.name, sheet_idx, keep_key)
        queued_names.add(name)

    for key, pending in list(wb._pending_defined_names.items()):  # noqa: SLF001
        name = getattr(pending, "name", None)
        if not name or name in queued_names:
            continue
        if getattr(pending, "localSheetId", None) != sheet_idx:
            continue
        refers_to = getattr(pending, "value", "") or ""
        retargeted = _retarget_sheet_formula(refers_to, old, new)
        if retargeted == refers_to:
            continue
        pending.value = retargeted[1:] if retargeted.startswith("=") else retargeted
        keep_key = (name, sheet_idx)
        wb._pending_defined_names[keep_key] = pending  # noqa: SLF001
        if key != keep_key:
            del wb._pending_defined_names[key]  # noqa: SLF001


def _store_print_area(value: Any) -> str | None:
    if value is None:
        return None
    if isinstance(value, Sequence) and not isinstance(value, str):
        return ",".join(
            area
            for area in (_store_single_print_area(item) for item in value)
            if area
        )
    return _store_single_print_area(value)


def _store_single_print_area(value: Any) -> str:
    area = str(value)
    if "!" in area:
        area = area.split("!", 1)[1]
    return area.replace("$", "").strip("'")


def _display_print_area(sheet_title: str, value: str) -> str:
    areas = []
    for raw_area in (_store_print_area(value) or "").split(","):
        area = raw_area.strip()
        if not area:
            continue
        if ":" in area:
            start, end = area.split(":", 1)
            area = f"{_absolute_cell_ref(start)}:{_absolute_cell_ref(end)}"
        else:
            area = _absolute_cell_ref(area)
        areas.append(area)
    quoted = "'" + sheet_title.replace("'", "''") + "'"
    return ",".join(f"{quoted}!{area}" for area in areas)


def _absolute_cell_ref(ref: str) -> str:
    match = re.fullmatch(r"\$?([A-Za-z]+)\$?(\d+)", ref.strip())
    if match is None:
        return ref
    col, row = match.groups()
    return f"${col.upper()}${row}"


class Worksheet:
    """Openpyxl-shaped proxy for a worksheet in a :class:`Workbook`.

    The proxy presents cell access, sheet metadata, and feature collections
    while keeping read, write, and modify-mode behavior aligned with the
    workbook that owns it.
    """

    BREAK_NONE = 0
    BREAK_ROW = 1
    BREAK_COLUMN = 2
    ORIENTATION_LANDSCAPE = "landscape"
    ORIENTATION_PORTRAIT = "portrait"
    PAPERSIZE_LETTER = "1"
    PAPERSIZE_LETTER_SMALL = "2"
    PAPERSIZE_TABLOID = "3"
    PAPERSIZE_LEDGER = "4"
    PAPERSIZE_LEGAL = "5"
    PAPERSIZE_STATEMENT = "6"
    PAPERSIZE_EXECUTIVE = "7"
    PAPERSIZE_A3 = "8"
    PAPERSIZE_A4 = "9"
    PAPERSIZE_A4_SMALL = "10"
    PAPERSIZE_A5 = "11"
    SHEETSTATE_VISIBLE = "visible"
    SHEETSTATE_HIDDEN = "hidden"
    SHEETSTATE_VERYHIDDEN = "veryHidden"

    __slots__ = (
        "_workbook", "_title", "_cells", "_dirty", "_dirty_values",
        "_format_dirty_cells", "_dimensions", "_rels",
        "_read_only_dimensions_reset", "_read_only_dimensions_forced",
        "_max_col_idx", "_next_append_row", "_current_row",
        "_append_buffer", "_append_buffer_start", "_bulk_writes",
        "_plain_cell_fast_path",
        "_freeze_panes", "_auto_filter", "_sort_state", "_scenarios",
        "_row_heights", "_row_dimensions", "_row_hidden", "_row_outline_levels",
        "_col_widths", "_col_dimensions", "_column_dimension_cache",
        "_col_hidden", "_col_outline_levels",
        "_sheet_state", "_sheet_state_dirty",
        "_merged_ranges", "_merged_ranges_loaded", "_collection_merged_ranges",
        "_pending_collection_merged_ranges", "_pending_unmerged_ranges",
        "_print_area", "_sheet_visibility_cache",
        # Read caches populated lazily on first access.
        "_style_payload_cache", "_style_component_cache",
        "_comments", "_comments_cache", "_threaded_comments_cache",
        "_hyperlinks",
        "_hyperlinks_cache", "_defined_names_cache",
        "_tables_cache", "_data_validations_cache",
        "_conditional_formatting_cache", "_images_cache", "_charts_cache",
        # Write-mode pending queues flushed in _flush() on save().
        "_pending_comments", "_pending_threaded_comments", "_pending_hyperlinks",
        "_pending_defined_names",
        "_pending_tables", "_pending_data_validations",
        "_pending_conditional_formats",
        # Pending rich-text values keyed by sparse (row, col) coordinate.
        "_pending_rich_text",
        # Pending array / data-table
        # formulas keyed by ``(row, col)``.  Each entry is a
        # ``(kind, payload)`` tuple where ``kind`` is one of
        # ``"array"``, ``"data_table"``, ``"spill_child"``.  Drained
        # at save time alongside the regular cell flush.
        "_pending_array_formulas", "_reader_has_array_formulas_cache",
        # Pending image queue.
        "_pending_images",
        # Pending chart queue.
        "_pending_charts", "_pending_chart_deletions",
        # Pending pivot table queue.
        "_pending_pivot_tables",
        # G17 / RFC-070 — lazy cache of PivotTableHandle objects
        # parsed off disk in modify mode. Set on first read.
        "_pivot_handles_cache",
        # Print/view/protection lazy slots.
        "_page_setup", "_page_margins", "_print_options", "_header_footer",
        "_sheet_properties", "_sheet_view", "_views", "_protection",
        "_legacy_drawing",
        "_print_title_rows", "_print_title_cols", "_print_titles_dirty",
        # Pending slicer presentations.
        "_pending_slicers",
        # Page breaks and sheetFormatPr lazy slots.
        "_row_breaks", "_col_breaks", "_sheet_format",
    )

    def __init__(self, parent: Workbook, title: str | None = None) -> None:
        initialize_worksheet_state(self, parent, _unique_detached_sheet_title(parent, title))

    @property
    def __dict__(self) -> dict[str, Any]:  # type: ignore[override]
        """Expose openpyxl-style instance attributes for introspection.

        WolfXL keeps worksheets slot-based to avoid per-sheet dictionaries, but
        upstream openpyxl tests and some user code inspect ``ws.__dict__``.
        Return the openpyxl-shaped attribute names without changing storage.
        """

        names = (
            "HeaderFooter",
            "_WorkbookChild__title",
            "_cells",
            "_charts",
            "_comments",
            "_current_row",
            "_drawing",
            "_hyperlinks",
            "_images",
            "_parent",
            "_pivots",
            "_print_area",
            "_print_cols",
            "_print_rows",
            "_rels",
            "_tables",
            "auto_filter",
            "col_breaks",
            "column_dimensions",
            "conditional_formatting",
            "data_validations",
            "defined_names",
            "legacy_drawing",
            "merged_cells",
            "page_margins",
            "page_setup",
            "print_options",
            "protection",
            "row_breaks",
            "row_dimensions",
            "scenarios",
            "sheet_format",
            "sheet_properties",
            "sheet_state",
            "views",
        )
        aliases = {
            "_WorkbookChild__title": "title",
            "_parent": "_workbook",
            "_pivots": "_pivot_handles_cache",
            "_print_cols": "_print_title_cols",
            "_print_rows": "_print_title_rows",
        }
        values: dict[str, Any] = {}
        for name in names:
            source = aliases.get(name, name)
            try:
                values[name] = getattr(self, source)
            except Exception:
                values[name] = None
        return values

    @property
    def title(self) -> str:
        """Worksheet title shown in the workbook tab list."""
        return self._title

    @title.setter
    def title(self, value: str) -> None:
        """Rename this worksheet (write mode only)."""
        wb = self._workbook
        old = self._title
        if old == value:
            return
        if value in wb._sheets or value in getattr(wb, "_chartsheets", {}):  # noqa: SLF001
            raise ValueError(f"Sheet '{value}' already exists")
        idx = wb._sheet_names.index(old)  # noqa: SLF001
        _queue_defined_name_sheet_rename(wb, old, value, idx)
        if wb._rust_patcher is not None:  # noqa: SLF001
            queue_rename = getattr(wb._rust_patcher, "queue_sheet_rename", None)  # noqa: SLF001
            if queue_rename is not None:
                queue_rename(old, value)
        # Update workbook bookkeeping.
        wb._sheet_names[idx] = value  # noqa: SLF001
        wb._sheets[value] = wb._sheets.pop(old)  # noqa: SLF001
        self._title = value
        self._defined_names_cache = None
        # Sync the Rust writer so ensure_sheet_exists() sees the new name.
        if wb._rust_writer is not None:  # noqa: SLF001
            wb._rust_writer.rename_sheet(old, value)  # noqa: SLF001

    # ------------------------------------------------------------------
    # openpyxl compat properties
    # ------------------------------------------------------------------

    @property
    def freeze_panes(self) -> str | None:
        """Get/set the freeze panes cell reference (e.g. ``'B2'``).

        In read mode, reads from the Rust backend.  In write mode,
        the value is stored and flushed to Rust on ``save()``.
        """
        return get_freeze_panes(self)

    @freeze_panes.setter
    def freeze_panes(self, value: str | None) -> None:
        """Set the freeze panes cell reference.

        Args:
            value: The top-left scrollable cell, such as ``"B2"``, or
                ``None`` to clear frozen panes.
        """
        set_freeze_panes(self, value)

    # ------------------------------------------------------------------
    # Print, view, and protection accessors
    # ------------------------------------------------------------------

    @property
    def page_setup(self) -> Any:
        """Return the worksheet page setup settings, creating them lazily."""
        return get_page_setup(self)

    @page_setup.setter
    def page_setup(self, value: Any) -> None:
        """Replace the worksheet page setup block.

        Args:
            value: Page setup object compatible with WolfXL's page setup
                serializer.
        """
        self._page_setup = value

    @property
    def page_margins(self) -> Any:
        """Return the worksheet page margins, creating them lazily."""
        return get_page_margins(self)

    @page_margins.setter
    def page_margins(self, value: Any) -> None:
        """Replace the worksheet page margins block.

        Args:
            value: Page margins object compatible with WolfXL's page margins
                serializer.
        """
        self._page_margins = value

    @property
    def HeaderFooter(self) -> Any:  # noqa: N802 - openpyxl alias
        """Openpyxl camel-case alias for :attr:`header_footer`."""
        return self.header_footer

    @property
    def oddHeader(self) -> Any:  # noqa: N802 - openpyxl alias
        """Odd-page header alias."""
        return self.header_footer.odd_header

    @property
    def oddFooter(self) -> Any:  # noqa: N802 - openpyxl alias
        """Odd-page footer alias."""
        return self.header_footer.odd_footer

    @property
    def evenHeader(self) -> Any:  # noqa: N802 - openpyxl alias
        """Even-page header alias."""
        return self.header_footer.even_header

    @property
    def evenFooter(self) -> Any:  # noqa: N802 - openpyxl alias
        """Even-page footer alias."""
        return self.header_footer.even_footer

    @property
    def firstHeader(self) -> Any:  # noqa: N802 - openpyxl alias
        """First-page header alias."""
        return self.header_footer.first_header

    @property
    def firstFooter(self) -> Any:  # noqa: N802 - openpyxl alias
        """First-page footer alias."""
        return self.header_footer.first_footer

    @property
    def header_footer(self) -> Any:
        """Return the worksheet header/footer settings, creating them lazily."""
        return get_header_footer(self)

    @header_footer.setter
    def header_footer(self, value: Any) -> None:
        """Replace the worksheet header/footer block.

        Args:
            value: Header/footer object compatible with WolfXL's
                header/footer serializer.
        """
        self._header_footer = value

    @property
    def sheet_view(self) -> Any:
        """Return the primary worksheet view settings.

        ``ws.freeze_panes`` mutations are mirrored into ``sheet_view.pane``
        on the setter side; on the getter side, if a sheet view has a
        non-None pane we surface that as ``freeze_panes`` for parity.
        """
        return get_sheet_view(self)

    @sheet_view.setter
    def sheet_view(self, value: Any) -> None:
        """Replace the worksheet view block.

        Args:
            value: Sheet view object compatible with WolfXL's sheet view
                serializer.
        """
        self._sheet_view = value

    @property
    def views(self) -> Any:
        """Return worksheet sheet views as a ``SheetViewList``."""
        from wolfxl.worksheet.views import SheetViewList

        if self._views is not None:
            return self._views
        return SheetViewList([self.sheet_view])

    @views.setter
    def views(self, value: Any) -> None:
        self._views = value

    @property
    def active_cell(self) -> str:
        """Active cell from the first sheet-view selection."""
        selections = self.sheet_view.selection
        return selections[0].activeCell if selections else "A1"

    @property
    def selected_cell(self) -> str:
        """Selected cell reference from the first sheet-view selection."""
        selections = self.sheet_view.selection
        return selections[0].sqref if selections else self.active_cell

    @property
    def show_gridlines(self) -> bool | None:
        """Whether worksheet gridlines are visible."""
        return self.sheet_view.showGridLines

    @show_gridlines.setter
    def show_gridlines(self, value: bool) -> None:
        """Set worksheet gridline visibility."""
        self.sheet_view.showGridLines = bool(value)

    @property
    def encoding(self) -> str:
        """Worksheet encoding marker for openpyxl compatibility."""
        return "utf-8"

    @property
    def mime_type(self) -> str:
        """Openpyxl-compatible worksheet MIME type."""
        return "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"

    @property
    def path(self) -> str:
        """Best-effort worksheet part path."""
        if self._title not in getattr(self._workbook, "_sheets", {}):  # noqa: SLF001
            return "/xl/worksheets/sheetNone.xml"
        try:
            index = self._workbook._sheet_names.index(self._title) + 1  # noqa: SLF001
        except ValueError:
            index = 1
        return f"/xl/worksheets/sheet{index}.xml"

    @property
    def legacy_drawing(self) -> Any:
        """Legacy VML drawing relationship, when exposed."""
        return self._legacy_drawing

    @legacy_drawing.setter
    def legacy_drawing(self, value: Any) -> None:
        self._legacy_drawing = value

    @property
    def array_formulae(self) -> dict[str, str]:
        """Return array formula ranges by master cell."""
        out: dict[str, str] = {}
        reader = getattr(self._workbook, "_rust_reader", None)  # noqa: SLF001
        if reader is not None and hasattr(reader, "read_sheet_array_formulas"):
            payloads = reader.read_sheet_array_formulas(self._title)  # noqa: SLF001
            if isinstance(payloads, dict):
                for coord, payload in payloads.items():
                    try:
                        row_col = a1_to_rowcol(coord)
                    except ValueError:
                        row_col = None
                    cell = self._cells.get(row_col) if row_col is not None else None  # noqa: SLF001
                    if cell is not None and getattr(cell, "_value_dirty", False):
                        continue
                    if isinstance(payload, dict) and payload.get("kind") == "array":
                        ref = payload.get("ref")
                        if ref:
                            out[coord] = ref
        for (row, col), (kind, payload) in self._pending_array_formulas.items():  # noqa: SLF001
            if kind == "array":
                ref = payload.get("ref")
                if ref:
                    out[rowcol_to_a1(row, col)] = ref
        for (row, col), cell in self._cells.items():  # noqa: SLF001
            if getattr(cell, "_formula_type", None) == "array" and cell._array_ref:  # noqa: SLF001
                out[rowcol_to_a1(row, col)] = cell._array_ref  # noqa: SLF001
        return out

    @property
    def column_groups(self) -> list[Any]:
        """Column group metadata placeholder."""
        groups = []
        for dimension in self.column_dimensions.values():
            if getattr(dimension, "min", None) and getattr(dimension, "max", None):
                if dimension.max > dimension.min:
                    groups.append(
                        f"{get_column_letter(dimension.min)}:{get_column_letter(dimension.max)}"
                    )
        return groups

    @property
    def defined_names(self) -> dict[str, Any]:
        """Worksheet-scoped defined names."""
        return get_defined_names(self)

    @property
    def scenarios(self) -> list[Any]:
        """Scenario metadata list."""
        return self._scenarios

    @property
    def sort_state(self) -> Any:
        """Worksheet-global sort state placeholder."""
        return self._sort_state

    @sort_state.setter
    def sort_state(self, value: Any) -> None:
        """Set worksheet-global sort state for openpyxl writer parity."""
        self._sort_state = value

    @property
    def print_options(self) -> Any:
        """Return worksheet print options."""
        return get_print_options(self)

    @property
    def print_titles(self) -> str:
        """Combined print-title range string."""
        parts = [
            f"{quote_sheetname(self.title)}!{part}"
            for part in [self.print_title_rows, self.print_title_cols]
            if part
        ]
        return ",".join(parts)

    def set_printer_settings(self, paper_size: int | str, orientation: str) -> None:
        """Set page paper size and orientation."""
        self.page_setup.paperSize = paper_size
        self.page_setup.orientation = orientation

    @property
    def protection(self) -> Any:
        """Return the worksheet protection settings, creating them lazily."""
        return get_protection(self)

    @protection.setter
    def protection(self, value: Any) -> None:
        """Replace the worksheet protection block.

        Args:
            value: Sheet protection object compatible with WolfXL's sheet
                protection serializer.
        """
        self._protection = value

    # ------------------------------------------------------------------
    # Page breaks and sheet format properties
    # ------------------------------------------------------------------

    @property
    def row_breaks(self) -> Any:
        """Return the horizontal page-break collection, creating it lazily."""
        return get_row_breaks(self)

    @row_breaks.setter
    def row_breaks(self, value: Any) -> None:
        """Replace the horizontal page-break collection.

        Args:
            value: Page break collection for row breaks.
        """
        self._row_breaks = value

    @property
    def col_breaks(self) -> Any:
        """Return the vertical page-break collection, creating it lazily."""
        return get_col_breaks(self)

    @col_breaks.setter
    def col_breaks(self, value: Any) -> None:
        """Replace the vertical page-break collection.

        Args:
            value: Page break collection for column breaks.
        """
        self._col_breaks = value

    @property
    def page_breaks(self) -> Any:
        """openpyxl alias — ``ws.page_breaks`` is a row-breaks alias."""
        return self.row_breaks

    @page_breaks.setter
    def page_breaks(self, value: Any) -> None:
        """Set the openpyxl row-break alias.

        Args:
            value: Page break collection assigned to ``row_breaks``.
        """
        self._row_breaks = value

    @property
    def sheet_format(self) -> Any:
        """Return sheet format properties, creating them lazily."""
        return get_sheet_format(self)

    @sheet_format.setter
    def sheet_format(self, value: Any) -> None:
        """Replace the sheet format properties block.

        Args:
            value: Sheet format properties object compatible with WolfXL's
                sheet format serializer.
        """
        self._sheet_format = value

    @property
    def sheet_properties(self) -> Any:
        """Return worksheet properties, creating them lazily."""
        return get_sheet_properties(self)

    @sheet_properties.setter
    def sheet_properties(self, value: Any) -> None:
        """Replace worksheet properties.

        Args:
            value: Worksheet properties object compatible with openpyxl's
                ``Worksheet.sheet_properties`` shape.
        """
        self._sheet_properties = value

    @property
    def dimension_holder(self) -> Any:
        """Return a fresh ``DimensionHolder`` view bound to this worksheet."""
        return get_dimension_holder(self)

    def to_rust_page_breaks_dict(self) -> dict[str, Any]:
        """Return the flat dict shape for ``<rowBreaks>`` / ``<colBreaks>``.

        Each side is ``None`` when the corresponding ``PageBreakList``
        is un-touched OR carries zero breaks — the patcher / writer
        then knows to skip emitting the corresponding XML block.
        """
        return _to_rust_page_breaks_dict(self)

    def to_rust_sheet_format_dict(self) -> dict[str, Any] | None:
        """Return the flat dict for ``<sheetFormatPr>`` or ``None``.

        Returns ``None`` when the wrapper is un-touched OR at all-default
        values — the writer then keeps the legacy hardcoded
        ``<sheetFormatPr defaultRowHeight="15"/>`` emit path.
        """
        return _to_rust_sheet_format_dict(self)

    @property
    def print_title_rows(self) -> str | None:
        """Row range repeated at the top of each printed page."""
        reader = getattr(self._workbook, "_rust_reader", None)
        if (
            self._print_title_rows is None
            and not self._print_titles_dirty
            and reader is not None
            and hasattr(reader, "read_print_titles")
        ):
            payload = reader.read_print_titles(self._title)
            if isinstance(payload, dict):
                rows = payload.get("rows")
                cols = payload.get("cols")
                if rows is not None:
                    set_print_title_rows(self, rows)
                    self._print_titles_dirty = False
                if cols is not None:
                    set_print_title_cols(self, cols)
                    self._print_titles_dirty = False
        return self._print_title_rows

    @print_title_rows.setter
    def print_title_rows(self, value: str | None) -> None:
        """Set repeat rows for printed pages.

        Args:
            value: Row range string such as ``"1:3"``, or ``None`` to clear
                repeat rows.
        """
        set_print_title_rows(self, value)

    @property
    def print_title_cols(self) -> str | None:
        """Column range repeated at the left of each printed page."""
        reader = getattr(self._workbook, "_rust_reader", None)
        if (
            self._print_title_cols is None
            and not self._print_titles_dirty
            and reader is not None
            and hasattr(reader, "read_print_titles")
        ):
            payload = reader.read_print_titles(self._title)
            if isinstance(payload, dict):
                rows = payload.get("rows")
                cols = payload.get("cols")
                if rows is not None:
                    set_print_title_rows(self, rows)
                    self._print_titles_dirty = False
                if cols is not None:
                    set_print_title_cols(self, cols)
                    self._print_titles_dirty = False
        return self._print_title_cols

    @print_title_cols.setter
    def print_title_cols(self, value: str | None) -> None:
        """Set repeat columns for printed pages.

        Args:
            value: Column range string such as ``"A:C"``, or ``None`` to clear
                repeat columns.
        """
        set_print_title_cols(self, value)

    def to_rust_setup_dict(self) -> dict[str, Any]:
        """Return the flat dict contract for the Rust patcher / writer.

        Returns ``None`` for any sub-block whose Python wrapper is at
        its construction defaults — the Rust side then knows to skip
        emitting the corresponding XML.
        """
        return _to_rust_setup_dict(self)

    @property
    def auto_filter(self) -> _AutoFilter:
        """Worksheet auto-filter proxy."""
        return get_auto_filter(self)

    @property
    def row_dimensions(self) -> RowDimensionProxy:
        """Dict-like row-dimension accessor keyed by 1-based row number."""
        return RowDimensionProxy(self)

    @property
    def column_dimensions(self) -> ColumnDimensionProxy:
        """Dict-like column-dimension accessor keyed by column letter."""
        return ColumnDimensionProxy(self)

    @property
    def print_area(self) -> str | None:
        """Get/set the print area range string (e.g. ``'A1:D10'``).

        Stored locally and flushed to the Rust writer on ``save()`` if the
        writer supports ``set_print_area()``.
        """
        reader = getattr(self._workbook, "_rust_reader", None)
        if (
            self._print_area is None
            and reader is not None
            and hasattr(reader, "read_print_area")
        ):
            self._print_area = reader.read_print_area(self._title)
        if self._print_area is None:
            return ""
        return _display_print_area(self._title, self._print_area)

    @print_area.setter
    def print_area(self, value: Any) -> None:
        """Set the worksheet print area.

        Args:
            value: A worksheet range string, or ``None`` to clear the print
                area.
        """
        self._print_area = _store_print_area(value)

    # ------------------------------------------------------------------
    # Cell access
    # ------------------------------------------------------------------

    def __getitem__(self, key: Any) -> Any:
        """openpyxl-compatible cell access.

        Supports:
        - ``ws["A1"]`` -> single Cell
        - ``ws["A1:B2"]`` -> tuple of tuples of Cell (2D range)
        - ``ws["A:B"]`` -> column range bounded by used range
        - ``ws["1:3"]`` -> row range
        - ``ws["A"]`` -> single column (tuple of Cell)
        - ``ws["1"]`` -> single row (str key; tuple of Cell)
        - ``ws[1]`` -> single row (int key; tuple of Cell)
        - ``ws[1:3]`` -> row slice (tuple of tuples of Cell)
        """
        return get_item(self, key)

    def _resolve_string_key(self, key: str) -> Any:
        """Resolve a string key to Cell / tuple / tuple-of-tuples."""
        return resolve_string_key(self, key)

    def _get_rect(
        self, min_row: int, min_col: int, max_row: int, max_col: int,
    ) -> tuple[tuple[Cell, ...], ...]:
        """Return a 2D tuple of Cells for the inclusive rectangle."""
        return get_rect(self, min_row, min_col, max_row, max_col)

    def _get_row_tuple(
        self, min_row: int, max_row: int,
    ) -> tuple[tuple[Cell, ...], ...]:
        """Return a tuple of row-tuples for rows min_row..max_row inclusive."""
        return get_row_tuple(self, min_row, max_row)

    def _get_col_tuple(
        self,
        min_col: int,
        max_col: int,
        min_row: int | None = None,
        max_row: int | None = None,
    ) -> tuple[tuple[Cell, ...], ...]:
        """Return a tuple of column-tuples for cols min_col..max_col inclusive."""
        return get_col_tuple(self, min_col, max_col, min_row, max_row)

    def __setitem__(self, key: str, value: Any) -> None:
        """``ws['A1'] = 42`` — shorthand for setting a cell's value."""
        value_type = type(value)
        if (
            isinstance(key, str)
            and value is not None
            and (
                value_type in _PLAIN_SCALAR_EXACT_TYPES
                or isinstance(value, _PLAIN_SCALAR_TYPES)
            )
        ):
            try:
                row, col = a1_to_rowcol(key)
            except ValueError:
                pass
            else:
                self.cell(row, col, value)
                return
        cell = self[key]
        cell.value = value

    def __delitem__(self, key: str) -> None:
        """Delete a materialized cell by A1 coordinate."""
        row, col = a1_to_rowcol(key)
        self._cells.pop((row, col), None)
        self._dirty.discard((row, col))
        self._dimensions = None

    def cell(self, row: int, column: int, value: Any = None) -> Cell:
        """Get or create a cell by 1-based (row, column). Matches openpyxl API."""
        if row < 1 or column < 1:
            raise ValueError("Row or column values must be at least 1")
        key = (row, column)
        cells = self._cells
        value_type = type(value)
        pending_array_formulas = self._pending_array_formulas
        if (
            value is not None
            and (
                value_type in _PLAIN_SCALAR_EXACT_TYPES
                or isinstance(value, _PLAIN_SCALAR_TYPES)
            )
            and key not in cells
            and self._plain_cell_fast_path
            and not (
                (self._merged_ranges or getattr(self._workbook, "_rust_reader", None) is not None)
                and self._is_merged_subordinate(row, column)
            )
            and (not pending_array_formulas or key not in pending_array_formulas)
        ):
            if value_type in _PLAIN_SCALAR_EXACT_TYPES:
                if value_type is str:
                    if len(value) > 32767:
                        value = value[:32767]
                    if _ILLEGAL_CHARACTERS_SEARCH(value):
                        raise IllegalCharacterError(_illegal_string_message(value))
                cell = _CELL_NEW(Cell)
                cell._ws = self  # noqa: SLF001
                cell._row = row  # noqa: SLF001
                cell._col = column  # noqa: SLF001
                cell._value = value  # noqa: SLF001
                cell._explicit_data_type = _UNSET  # noqa: SLF001
                cell._value_dirty = True  # noqa: SLF001
                cell._format_dirty = False  # noqa: SLF001
                cell._style_slots_ready = False  # noqa: SLF001
                cell._value_is_plain = True  # noqa: SLF001
                cells[key] = cell
                self._dirty_values[key] = value
                if row >= self._next_append_row:
                    self._next_append_row = row + 1
                if column > self._max_col_idx:
                    self._max_col_idx = column
                return cell
            return _new_plain_scalar_cell(self, row, column, value, value_type, key, cells)

        c = self._get_or_create_cell(row, column)
        if value is not None:
            c.value = value
        return c

    def _get_or_create_cell(self, row: int, col: int) -> Cell | MergedCell:
        if row < 1 or col < 1:
            raise ValueError("Row or column values must be at least 1")
        # Materialize the append buffer on first random cell access so that
        # Cell objects exist for previously-appended rows.
        if self._append_buffer:
            self._materialize_append_buffer()
        if (
            (self._merged_ranges or getattr(self._workbook, "_rust_reader", None) is not None)
            and self._is_merged_subordinate(row, col)
        ):
            return MergedCell(self, row, col)
        key = (row, col)
        cell = self._cells.get(key)
        if cell is None:
            cell = Cell(self, row, col)
            self._cells[key] = cell
        return cell

    def _is_merged_subordinate(self, row: int, col: int) -> bool:
        from wolfxl._worksheet_collections import _loaded_merged_range_refs

        refs = _loaded_merged_range_refs(self)
        for ref in refs:
            try:
                min_col, min_row, max_col, max_row = range_boundaries(str(ref))
            except Exception:
                continue
            if None in (min_col, min_row, max_col, max_row):
                continue
            if (
                int(min_row) <= row <= int(max_row)
                and int(min_col) <= col <= int(max_col)
            ):
                return not (row == int(min_row) and col == int(min_col))
        return False

    def _mark_dirty(self, row: int, col: int) -> None:
        self._dirty.add((row, col))
        if row >= self._next_append_row:
            self._next_append_row = row + 1
        if col > self._max_col_idx:
            self._max_col_idx = col

    # ------------------------------------------------------------------
    # Append (openpyxl-compatible row insertion)
    # ------------------------------------------------------------------

    def append(self, iterable: Iterable[Any]) -> None:
        """Append a row of values. Matches openpyxl's ``ws.append()`` API.

        Successive calls auto-increment the row index. Values are written to
        columns starting at 1 (A).

        Performance: rows are buffered as raw Python lists — no Cell objects
        are created. The buffer is flushed directly to ``write_sheet_values()``
        on save, bypassing per-cell FFI overhead entirely.  If cell-level
        access is needed later (e.g. ``ws.cell(1,1).font = ...``), the buffer
        is materialized into Cell objects on demand.
        """
        append_row(self, iterable)

    def _materialize_append_buffer(self) -> None:
        """Convert the append buffer into Cell objects.

        Called lazily on the first ``cell()`` / ``__getitem__`` access after
        appending.  After this call ``_append_buffer`` is empty and all values
        live in the normal ``_cells`` / ``_dirty`` tracking.
        """
        materialize_append_buffer(self)

    def write_rows(
        self,
        rows: list[list[Any]],
        start_row: int = 1,
        start_col: int = 1,
        *,
        copy: bool = True,
        plain_values_only: bool = False,
    ) -> None:
        """Bulk-write a 2D grid of values starting at (start_row, start_col).

        Unlike ``append()``, this writes to an arbitrary position. Values are
        buffered and flushed via a single ``write_sheet_values()`` Rust call
        at save time, avoiding per-cell FFI overhead.

        ``rows`` is a list of lists. Each inner list is one row of values.
        By default WolfXL snapshots those rows immediately, so later changes to
        the original list do not affect the saved workbook. Pass
        ``copy=False`` when the caller owns the grid and will not mutate it
        before saving.

        Pass ``plain_values_only=True`` when every value is a plain Excel value:
        ``None``, ``bool``, ``int``, ``float``, or ``str``. This skips the
        defensive scan for date/object values during save.
        """
        _write_rows(
            self,
            rows,
            start_row,
            start_col,
            copy=copy,
            plain_values_only=plain_values_only,
        )

    def write_styled_rows(
        self,
        rows: list[list[Any]],
        styles: Any,
        start_row: int = 1,
        start_col: int = 1,
        *,
        copy: bool = True,
        plain_values_only: bool = False,
        normalized_styles: bool = False,
    ) -> None:
        """Bulk-write values plus repeated styles in one public fast path.

        ``rows`` is the same 2D value grid accepted by :meth:`write_rows`.
        ``styles`` describes the cell formatting to apply:

        * a dict keyed by local column number or letter, for column styles;
        * a same-shaped 2D style grid, with ``None`` for default cells;
        * a callable ``(row, column, value) -> style`` for row-aware styles;
        * one style spec applied to every cell.

        A style spec can be a number-format string, a dict such as
        ``{"font": Font(bold=True)}``, a native flat format dict, or a
        Font/Fill/Border/Alignment/Protection object. Repeated styles are
        interned once and applied through the Rust writer's style-id grid.

        Pass ``normalized_styles=True`` when ``styles`` is already a same-shaped
        grid of flat native format dicts or ``None``. This skips the Python
        normalization pass for power users that prebuild reusable style grids.
        """
        if not rows:
            return
        grid = [list(row) for row in rows] if copy else rows
        style_grid = (
            styles
            if normalized_styles
            else styled_rows_style_grid(grid, styles, start_row, start_col)
        )
        _write_rows(
            self,
            grid,
            start_row,
            start_col,
            copy=False,
            plain_values_only=plain_values_only,
            style_grid=style_grid,
        )

    def _materialize_bulk_writes(self) -> None:
        """Convert bulk write buffers into Cell objects.

        Called before the patcher flush path which has no batch API and
        needs all values as individual dirty cells.
        """
        materialize_bulk_writes(self)

    def to_dataframe(
        self,
        *,
        header: bool = True,
        min_row: int | None = None,
        max_row: int | None = None,
        min_col: int | None = None,
        max_col: int | None = None,
    ) -> Any:
        """Return worksheet values as a pandas DataFrame."""
        from wolfxl.utils.dataframe import worksheet_to_dataframe

        return worksheet_to_dataframe(
            self,
            header=header,
            min_row=min_row,
            max_row=max_row,
            min_col=min_col,
            max_col=max_col,
        )

    @staticmethod
    def _extract_non_batchable(
        grid: list[list[Any]], start_row: int, start_col: int,
    ) -> list[tuple[int, int, Any]]:
        """Extract non-batchable values from grid, replacing them with None.

        The Rust batch writer handles primitive values, booleans, formula
        strings, and error literals. Non-primitive values such as dates and
        datetimes still use per-cell ``write_cell_value()`` calls so their
        default number-format behavior stays unchanged.
        """
        return extract_non_batchable(grid, start_row, start_col)

    # ------------------------------------------------------------------
    # Iteration
    # ------------------------------------------------------------------

    def iter_rows(
        self,
        min_row: int | None = None,
        max_row: int | None = None,
        min_col: int | None = None,
        max_col: int | None = None,
        values_only: bool = False,
    ) -> Iterator[tuple[Any, ...]]:
        """Iterate over rows in a range. Matches openpyxl's iter_rows API.

        When the workbook was opened with ``read_only=True`` (or when this
        sheet has more than
        :data:`wolfxl._streaming.AUTO_STREAM_ROW_THRESHOLD` rows) this
        method becomes a true SAX-streaming generator backed by the
        Rust ``StreamingSheetReader``. Cells yielded in that path are
        :class:`wolfxl._streaming.StreamingCell` instances —
        immutable read-only proxies whose ``font`` / ``fill`` /
        ``border`` / ``alignment`` / ``number_format`` properties
        delegate to the existing eager style table. ``values_only=True``
        yields plain value tuples and never instantiates a cell object.
        """
        return _iter_rows(
            self,
            min_row=min_row,
            max_row=max_row,
            min_col=min_col,
            max_col=max_col,
            values_only=values_only,
        )

    def iter_cols(
        self,
        min_col: int | None = None,
        max_col: int | None = None,
        min_row: int | None = None,
        max_row: int | None = None,
        values_only: bool = False,
    ) -> Iterator[tuple[Any, ...]]:
        """Iterate over columns in a range. Matches openpyxl's iter_cols API.

        Unlike ``iter_rows``, the first-position arguments are column bounds.
        Yields one tuple per column, each containing values (or Cells) for
        every row in range.
        """
        yield from _iter_cols(
            self,
            min_col=min_col,
            max_col=max_col,
            min_row=min_row,
            max_row=max_row,
            values_only=values_only,
        )

    def _cells_by_col(
        self,
        min_col: int,
        max_col: int,
        min_row: int | None = None,
        max_row: int | None = None,
    ) -> Iterator[tuple[Any, ...]]:
        """openpyxl private helper used by legacy worksheet callers."""
        yield from self.iter_cols(
            min_col=min_col,
            max_col=max_col,
            min_row=min_row,
            max_row=max_row,
        )

    def _iter_cols_bulk(
        self,
        min_col: int | None,
        max_col: int | None,
        min_row: int | None,
        max_row: int | None,
    ) -> Iterator[tuple[Any, ...]]:
        """Bulk-read column values via a single Rust FFI call, then transpose.

        Mirrors ``_iter_rows_bulk`` but yields column-major tuples. One
        ``read_sheet_values_plain`` call reads the whole rectangle; transposition
        happens in Python. This avoids per-cell Rust calls in the values_only
        fast path and keeps parity with ``iter_rows`` performance characteristics.
        """
        yield from _iter_cols_bulk(self, min_col, max_col, min_row, max_row)

    @property
    def rows(self) -> Iterator[tuple[Any, ...]]:
        """Iterator over rows (tuples of Cell) — openpyxl alias for ``iter_rows()``."""
        return self.iter_rows()

    @property
    def columns(self) -> Iterator[tuple[Any, ...]]:
        """Iterator over columns (tuples of Cell) — openpyxl alias for ``iter_cols()``."""
        return self.iter_cols()

    @property
    def values(self) -> Iterator[tuple[Any, ...]]:
        """Iterator over row values — openpyxl alias for ``iter_rows(values_only=True)``."""
        return self.iter_rows(values_only=True)

    def iter_value_chunks(
        self,
        min_row: int | None = None,
        max_row: int | None = None,
        min_col: int | None = None,
        max_col: int | None = None,
        *,
        chunk_size: int = 1024,
    ) -> Iterator[tuple[tuple[Any, ...], ...]]:
        """Iterate worksheet values in row chunks.

        This WolfXL extension returns the same value tuples as
        ``iter_rows(values_only=True)``, grouped into chunks so large table
        scans can use fewer Python iterator round trips.
        """
        from wolfxl._worksheet_iteration import iter_value_chunks

        return iter_value_chunks(
            self,
            min_row=min_row,
            max_row=max_row,
            min_col=min_col,
            max_col=max_col,
            chunk_size=chunk_size,
        )

    def _iter_rows_bulk(
        self,
        min_row: int | None,
        max_row: int | None,
        min_col: int | None,
        max_col: int | None,
    ) -> Iterator[tuple[Any, ...]]:
        """Bulk-read values via a single Rust FFI call (values_only fast path).

        Uses ``read_sheet_values_plain()`` when available (returns native
        Python objects), falling back to ``read_sheet_values()`` + per-cell
        ``_payload_to_python()`` conversion otherwise.
        """
        yield from _iter_rows_bulk(self, min_row, max_row, min_col, max_col)

    def iter_cell_records(
        self,
        min_row: int | None = None,
        max_row: int | None = None,
        min_col: int | None = None,
        max_col: int | None = None,
        *,
        data_only: bool | None = None,
        include_format: bool = True,
        include_empty: bool = False,
        include_formula_blanks: bool = True,
        include_coordinate: bool = True,
        include_style_id: bool = True,
        include_extended_format: bool = True,
        include_cached_formula_value: bool = False,
    ) -> Iterator[dict[str, Any]]:
        """Iterate populated cells as compact dictionaries.

        This is WolfXL's high-throughput read API for ingestion and dataframe
        workloads. In read mode it makes one Rust call for the requested range,
        returning native Python values plus optional formatting metadata such as
        ``number_format``, ``bold``, ``indent``, and border cues.

        Coordinates are openpyxl-style 1-based ``row`` / ``column`` integers.
        Empty cells are skipped by default; pass ``include_empty=True`` when a
        dense rectangular record stream is needed. Formula cells without a
        backing cached value are included by default; pass
        ``include_formula_blanks=False`` to skip those template-only formulas.
        Pass ``include_coordinate=False`` when row/column integers are enough
        and avoiding A1 string allocation matters. Pass
        ``include_style_id=False`` when semantic format fields are enough and
        callers do not need workbook-internal style identifiers. Pass
        ``include_extended_format=False`` to keep raw font flags and number
        formats while skipping expensive style-grid fields such as fill,
        alignment, and border cues. Pass
        ``include_cached_formula_value=True`` to include a ``cached_value`` key
        on formula records that have a saved cached result.

        Args:
            min_row: Optional 1-based first row.
            max_row: Optional 1-based last row.
            min_col: Optional 1-based first column.
            max_col: Optional 1-based last column.
            data_only: Override workbook formula mode for this scan.
            include_format: Include number-format and style summary fields.
            include_empty: Include empty cells inside the requested bounds.
            include_formula_blanks: Include formulas without cached values.
            include_coordinate: Include A1 coordinate strings.
            include_style_id: Include workbook-internal style identifiers.
            include_extended_format: Include fill, alignment, and border cues.
            include_cached_formula_value: Include saved cached formula values.

        Yields:
            Compact dictionaries for worksheet cells in row-major order.
        """
        yield from _iter_cell_records(
            self,
            min_row=min_row,
            max_row=max_row,
            min_col=min_col,
            max_col=max_col,
            data_only=data_only,
            include_format=include_format,
            include_empty=include_empty,
            include_formula_blanks=include_formula_blanks,
            include_coordinate=include_coordinate,
            include_style_id=include_style_id,
            include_extended_format=include_extended_format,
            include_cached_formula_value=include_cached_formula_value,
        )

    def _collect_pending_overlay(self) -> dict[tuple[int, int], Any]:
        """Return ``{(row, col): value}`` for cells modified since the last save.

        Includes compact scalar edits in ``_dirty_values``, explicit cell edits
        in ``_dirty``, the append buffer, and bulk-write grids. Returns an empty
        dict when nothing is pending — the Rust read path stays a hot,
        allocation-free loop.
        """
        return collect_pending_overlay(self)

    def cell_records(
        self,
        min_row: int | None = None,
        max_row: int | None = None,
        min_col: int | None = None,
        max_col: int | None = None,
        *,
        data_only: bool | None = None,
        include_format: bool = True,
        include_empty: bool = False,
        include_formula_blanks: bool = True,
        include_coordinate: bool = True,
        include_style_id: bool = True,
        include_extended_format: bool = True,
        include_cached_formula_value: bool = False,
    ) -> list[dict[str, Any]]:
        """Return ``iter_cell_records(...)`` as a list.

        Args:
            min_row: Optional 1-based first row.
            max_row: Optional 1-based last row.
            min_col: Optional 1-based first column.
            max_col: Optional 1-based last column.
            data_only: Override workbook formula mode for this scan.
            include_format: Include number-format and style summary fields.
            include_empty: Include empty cells inside the requested bounds.
            include_formula_blanks: Include formulas without cached values.
            include_coordinate: Include A1 coordinate strings.
            include_style_id: Include workbook-internal style identifiers.
            include_extended_format: Include fill, alignment, and border cues.
            include_cached_formula_value: Include saved cached formula values.

        Returns:
            A list containing the same dictionaries yielded by
            :meth:`iter_cell_records`.
        """
        return list(
            self.iter_cell_records(
                min_row=min_row,
                max_row=max_row,
                min_col=min_col,
                max_col=max_col,
                data_only=data_only,
                include_format=include_format,
                include_empty=include_empty,
                include_formula_blanks=include_formula_blanks,
                include_coordinate=include_coordinate,
                include_style_id=include_style_id,
                include_extended_format=include_extended_format,
                include_cached_formula_value=include_cached_formula_value,
            ),
        )

    def cached_formula_values(self, *, qualified: bool = False) -> dict[str, Any]:
        """Return cached formula results for this sheet.

        Keys are A1 coordinates by default. Pass ``qualified=True`` to return
        ``"Sheet!A1"`` keys, matching :meth:`Workbook.cached_formula_values`.
        Only formula cells with saved cached values are included; uncached
        template formulas are omitted.

        Args:
            qualified: Include the worksheet name in each key.

        Returns:
            Mapping of cell reference to saved cached formula value.
        """
        return _cached_formula_values(self, qualified=qualified)

    def sheet_visibility(self) -> dict[str, Any]:
        """Return hidden rows/columns and outline levels for this sheet.

        Row and column identifiers are 1-based to mirror openpyxl's dimension
        collections. The returned shape is:
        ``hidden_rows``, ``hidden_columns``, ``row_outline_levels``, and
        ``column_outline_levels``.
        """
        return _sheet_visibility(self)

    def _iter_cell_records_python(
        self,
        *,
        min_row: int | None,
        max_row: int | None,
        min_col: int | None,
        max_col: int | None,
        include_empty: bool,
        include_coordinate: bool = True,
    ) -> Iterator[dict[str, Any]]:
        yield from _iter_cell_records_python(
            self,
            min_row=min_row,
            max_row=max_row,
            min_col=min_col,
            max_col=max_col,
            include_empty=include_empty,
            include_coordinate=include_coordinate,
        )

    def _read_only_declared_dimension_bounds(self) -> tuple[int, int, int, int] | None:
        """Return read-only ``<dimension>`` bounds unless the user reset them."""
        if not getattr(self._workbook, "_read_only", False):  # noqa: SLF001
            return None
        if self._read_only_dimensions_reset:
            return None
        path = getattr(self._workbook, "_source_path", None)  # noqa: SLF001
        if not path:
            return None
        from wolfxl._streaming import _source_dimension_bounds

        return _source_dimension_bounds(path, self.title)

    def _calculate_read_only_dimension(self, force: bool = False) -> str:
        declared = self._read_only_declared_dimension_bounds()
        if declared is not None:
            min_row, min_col, max_row, max_col = declared
            return f"{rowcol_to_a1(min_row, min_col)}:{rowcol_to_a1(max_row, max_col)}"

        if not force:
            raise ValueError("Worksheet is unsized, use calculate_dimension(force=True)")

        max_row, max_col = self._read_dimensions()
        self._dimensions = (max_row, max_col)
        self._read_only_dimensions_forced = True
        return f"A1:{rowcol_to_a1(max_row or 1, max_col or 1)}"

    def calculate_dimension(self, force: Any = None) -> str:
        """Return the used worksheet range in openpyxl's ``A1:C10`` form."""
        if getattr(self._workbook, "_read_only", False):  # noqa: SLF001
            return self._calculate_read_only_dimension(force=bool(force))
        if force is not None:
            raise TypeError("calculate_dimension() got an unexpected keyword argument 'force'")
        return _calculate_dimension(self)

    def reset_dimensions(self) -> None:
        """Discard read-only worksheet dimensions so callers can force-scan them."""
        if not getattr(self._workbook, "_read_only", False):  # noqa: SLF001
            raise AttributeError("reset_dimensions is only available in read_only=True mode")
        self._dimensions = None
        self._read_only_dimensions_reset = True
        self._read_only_dimensions_forced = False

    def _read_dimension_bounds(self) -> tuple[int, int, int, int] | None:
        """Return 1-based ``(min_row, min_col, max_row, max_col)`` bounds.

        Modify mode has both a Rust reader (for the on-disk extents) and
        Python-side pending writes (cells/append buffer/bulk writes). The
        reported bounds must be the union, otherwise callers that derive
        ranges from ``calculate_dimension()`` miss unsaved edits.
        """
        return _read_dimension_bounds(self)

    def _pending_writes_bounds(self) -> tuple[int, int, int, int] | None:
        """Bounds of unsaved Python-side writes: ``(min_row, min_col, max_row, max_col)``.

        Used by ``_read_dimension_bounds`` to fold modify-mode edits into the
        on-disk Rust bounds, and by ``_max_row``/``_max_col`` so write-mode
        ``ws.max_row`` reflects ``append()``/``write_rows()`` that haven't
        materialized yet.

        Iterates pending write trackers (``_dirty_values`` and ``_dirty``)
        rather than ``_cells`` — the cell map is populated by mere read access
        (``ws['Z999']`` materializes a Cell without modifying it), so reading a
        far cell would otherwise inflate dimension bounds and trigger oversized
        scans in downstream callers.
        """
        return pending_writes_bounds(self)

    def _read_dimensions(self) -> tuple[int, int]:
        """Discover sheet dimensions from the Rust backend (read mode only)."""
        return _read_dimensions(self)

    def _max_row(self) -> int:
        return _worksheet_max_row(self)

    def _max_col(self) -> int:
        return _worksheet_max_col(self)

    # openpyxl exposes these as properties, not methods. Mirror that contract
    # so ``ws.max_row`` (no parens) works as a drop-in for openpyxl callers.
    # Pinned by ``tests/parity/test_read_parity.py`` (uses ``op_ws.max_row``).
    @property
    def max_row(self) -> int | None:
        """Largest row index visible to openpyxl-style callers."""
        if getattr(self._workbook, "_read_only", False):  # noqa: SLF001
            if self._read_only_dimensions_forced and self._dimensions is not None:
                return self._dimensions[0]
            declared = self._read_only_declared_dimension_bounds()
            if declared is not None:
                return declared[2]
            return None
        return self._max_row()

    @property
    def max_column(self) -> int | None:
        """Largest column index visible to openpyxl-style callers."""
        if getattr(self._workbook, "_read_only", False):  # noqa: SLF001
            if self._read_only_dimensions_forced and self._dimensions is not None:
                return self._dimensions[1]
            declared = self._read_only_declared_dimension_bounds()
            if declared is not None:
                return declared[3]
            return None
        return self._max_col()

    @property
    def min_row(self) -> int:
        """Smallest declared row index visible to read-only callers."""
        declared = self._read_only_declared_dimension_bounds()
        if declared is not None:
            return declared[0]
        return 1

    @property
    def min_column(self) -> int:
        """Smallest declared column index visible to read-only callers."""
        declared = self._read_only_declared_dimension_bounds()
        if declared is not None:
            return declared[1]
        return 1

    @property
    def dimensions(self) -> str:
        """Used worksheet range in A1 form, e.g. ``"A1:C10"``."""
        return self.calculate_dimension()

    @property
    def parent(self) -> Workbook:
        """The containing Workbook."""
        return self._workbook

    @property
    def sheet_state(self) -> str:
        """Visibility state: ``"visible"``, ``"hidden"``, or ``"veryHidden"``.

        Defaults to ``"visible"`` and lazily reflects the native reader's
        parsed ``<sheet state="hidden">`` workbook metadata when available.
        """
        reader = getattr(self._workbook, "_rust_reader", None)
        if (
            self._sheet_state is None
            and reader is not None
            and hasattr(reader, "read_sheet_state")
        ):
            self._sheet_state = reader.read_sheet_state(self._title)
        return self._sheet_state or "visible"

    @sheet_state.setter
    def sheet_state(self, value: str) -> None:
        """Set worksheet visibility state for openpyxl-shaped callers."""
        if value not in {"visible", "hidden", "veryHidden"}:
            raise ValueError(
                "sheet_state must be 'visible', 'hidden', or 'veryHidden'"
            )
        self._sheet_state = value
        self._sheet_state_dirty = True

    @property
    def _charts(self) -> list[Any]:
        """Return the charts attached to this worksheet.

        The returned list is live, matching openpyxl's private ``_charts``
        compatibility surface. Mutating it affects the next save.
        """
        return _get_charts(self)

    @_charts.setter
    def _charts(self, value: list[Any]) -> None:
        """Replace the openpyxl-compatible private chart list."""
        self._pending_charts = value
        self._charts_cache = value

    def add_chart(self, chart: Any, anchor: Any = None) -> None:
        """Attach a chart to this worksheet.

        Mirrors :meth:`openpyxl.worksheet.worksheet.Worksheet.add_chart`.

        Args:
            chart: Chart object to embed, such as a bar, line, or pie chart.
            anchor: Optional A1-style anchor cell. ``None`` uses the chart's
                default anchor, matching openpyxl behavior.

        Raises:
            TypeError: If ``chart`` is not a supported chart object.
            ValueError: If ``anchor`` is not a valid A1-style cell reference.
        """
        _add_chart(self, chart, anchor)

    def add_pivot_table(self, pivot_table: Any, anchor: str | None = None) -> None:
        """Attach a pivot table to this worksheet.

        The pivot table's cache must already be registered on the owning
        workbook with :meth:`Workbook.add_pivot_cache`.

        Args:
            pivot_table: Pivot table object to attach.
            anchor: Optional A1-style anchor that overrides the pivot
                table's pre-set location. Matches the openpyxl/G17
                pattern ``ws.add_pivot_table(pt, "A1")``.

        Raises:
            TypeError: If ``pivot_table`` is not a supported pivot table.
            ValueError: If the pivot table cache has not been registered.
            RuntimeError: If the workbook is not in modify mode.
        """
        if anchor is not None:
            from wolfxl.pivot._table import Location

            if isinstance(pivot_table.location, Location):
                pivot_table.location.ref = str(anchor)
            else:
                pivot_table.location = str(anchor)
        _add_pivot_table(self, pivot_table)

    def add_pivot(self, pivot_table: Any | None = None, *, pivot: Any | None = None) -> None:
        """Openpyxl-compatible alias for :meth:`add_pivot_table`."""
        pivot_table = pivot_table if pivot_table is not None else pivot
        if pivot_table is None:
            raise TypeError("add_pivot() missing required argument: 'pivot'")
        self.add_pivot_table(pivot_table)

    def add_slicer(self, slicer: Any, anchor: str) -> None:
        """Attach a slicer presentation to this worksheet.

        The slicer's cache must already be registered on the workbook with
        :meth:`Workbook.add_slicer_cache`.

        Args:
            slicer: Slicer object to attach.
            anchor: A1-style top-left anchor cell, such as ``"H2"``.

        Raises:
            TypeError: If ``slicer`` is not a supported slicer object.
            ValueError: If the slicer's cache is not registered or ``anchor``
                is not a valid A1-style cell reference.
            RuntimeError: If the workbook is not in modify mode.
        """
        _add_slicer(self, slicer, anchor)

    @staticmethod
    def _validate_a1_anchor(anchor: str) -> None:
        """Validate an A1-style single-cell anchor.

        Args:
            anchor: Cell reference such as ``"E15"`` or ``"AA200"``.

        Raises:
            ValueError: If ``anchor`` is a range, sheet-qualified reference,
                malformed cell reference, or outside Excel's worksheet bounds.
        """
        _validate_a1_anchor(anchor)

    def remove_chart(self, chart: Any) -> None:
        """Remove a chart that was previously attached to this worksheet.

        Mirrors the openpyxl idiom ``ws._charts.remove(chart)``.

        Args:
            chart: Chart instance previously passed to :meth:`add_chart`.

        Raises:
            ValueError: If ``chart`` was never attached to this worksheet or
                has already been removed.
        """
        _remove_chart(self, chart)

    def replace_chart(self, old: Any, new: Any) -> None:
        """Replace one attached chart with another.

        The replacement keeps the old chart's anchor and list position unless
        the new chart already has an explicit anchor.

        Args:
            old: Chart instance previously passed to :meth:`add_chart`.
            new: Replacement chart object.

        Raises:
            TypeError: If ``new`` is not a supported chart object.
            ValueError: If ``old`` was never attached to this worksheet.
        """
        _replace_chart(self, old, new)

    @property
    def _images(self) -> list[Any]:
        """Return the images attached to this worksheet.

        The returned list is live, matching openpyxl's private ``_images``
        compatibility surface. Mutating it affects the next save.
        """
        return _get_images(self)

    @_images.setter
    def _images(self, value: list[Any]) -> None:
        """Replace the openpyxl-compatible private image list."""
        self._pending_images = value
        self._images_cache = value

    @property
    def images(self) -> list[Any]:
        """Public alias for :attr:`_images`."""
        return _get_images(self)

    def add_image(self, img: Any, anchor: Any = None) -> None:
        """Attach an image to this worksheet.

        Mirrors :meth:`openpyxl.worksheet.worksheet.Worksheet.add_image`.

        Args:
            img: Image object to embed.
            anchor: Optional placement. Accepts an A1-style cell reference,
                ``TwoCellAnchor``, ``AbsoluteAnchor``, or ``None``. ``None``
                defaults to ``"A1"``.

        Raises:
            TypeError: If ``img`` or ``anchor`` is not supported.
            ValueError: If an A1-style ``anchor`` is malformed.
        """
        _add_image(self, img, anchor)

    def remove_image(self, index_or_image: int | Any) -> None:
        """Remove one image attached to this worksheet.

        Args:
            index_or_image: Zero-based index into ``ws.images`` (or ``ws._images``)
                or a concrete image object from that list.

        Raises:
            ValueError: If the index is out of range or the image is not attached.
        """
        _remove_image(self, index_or_image)

    def replace_image(self, index_or_image: int | Any, new_image: Any) -> None:
        """Replace one attached image with a new image object.

        Replacement is implemented as remove + add at flush time.

        Args:
            index_or_image: Zero-based index into ``ws.images`` (or ``ws._images``)
                or a concrete image object from that list.
            new_image: Replacement image object.

        Raises:
            TypeError: If ``new_image`` is not a supported image object.
            ValueError: If the selected image is not attached to this worksheet.
        """
        _replace_image(self, index_or_image, new_image)

    @property
    def merged_cells(self) -> _MergedCellsProxy:
        """openpyxl-shape merged-cells accessor.

        Lazy: on first access, pulls merged ranges from the Rust reader
        (read mode) or returns the in-memory write-mode set. Always exposes
        a ``.ranges`` iterable of range strings — matching openpyxl's
        ``MultiCellRange`` shape closely enough for SynthGL's needs.
        """
        return _MergedCellsProxy(self)

    # ------------------------------------------------------------------
    # Write-mode helpers
    # ------------------------------------------------------------------

    def merge_cells(
        self,
        range_string: str | None = None,
        start_row: int | None = None,
        start_column: int | None = None,
        end_row: int | None = None,
        end_column: int | None = None,
    ) -> None:
        """Merge cells (write mode only). Example: ``ws.merge_cells('A1:B2')``."""
        _merge_cells(
            self,
            _merge_range_from_args(
                range_string,
                start_row,
                start_column,
                end_row,
                end_column,
            ),
        )

    def unmerge_cells(
        self,
        range_string: str | None = None,
        start_row: int | None = None,
        start_column: int | None = None,
        end_row: int | None = None,
        end_column: int | None = None,
    ) -> None:
        """Unmerge a previously merged range.

        If *range_string* was not previously merged, silently does nothing
        (matches openpyxl behaviour).
        """
        _unmerge_cells(
            self,
            _merge_range_from_args(
                range_string,
                start_row,
                start_column,
                end_row,
                end_column,
            ),
        )

    # ------------------------------------------------------------------
    # Structural operations
    # ------------------------------------------------------------------

    def insert_rows(self, idx: int, amount: int = 1) -> None:
        """Shift rows down to insert *amount* empty rows starting at *idx*.

        The operation follows openpyxl's 1-based row indexing and is
        applied when the workbook is saved. In modify mode WolfXL preserves
        the existing workbook package and rewrites the affected worksheet
        parts in place.

        Args:
            idx: First row to insert before. Must be at least 1.
            amount: Number of rows to insert. Must be at least 1.

        Raises:
            ValueError: If ``idx`` or ``amount`` is outside the supported
                range.
        """
        _insert_rows(self, idx, amount)

    def delete_rows(self, idx: int, amount: int = 1) -> None:
        """Delete *amount* rows starting at *idx*, shifting subsequent rows up.

        Args:
            idx: First row to delete. Must be at least 1.
            amount: Number of rows to delete. Must be at least 1.

        Raises:
            ValueError: If ``idx`` or ``amount`` is outside the supported
                range.
        """
        _delete_rows(self, idx, amount)

    def insert_cols(self, idx: int | str, amount: int = 1) -> None:
        """Shift columns right to insert *amount* empty columns at *idx*.

        Args:
            idx: First column to insert before, either as a 1-based integer
                or an Excel column label such as ``"A"`` or ``"AB"``.
            amount: Number of columns to insert. ``0`` is a no-op.

        Raises:
            ValueError: If ``idx`` or ``amount`` is outside the supported
                range.
        """
        _insert_cols(self, idx, amount)

    def delete_cols(self, idx: int | str, amount: int = 1) -> None:
        """Delete *amount* columns starting at *idx*, shifting subsequent columns left.

        Args:
            idx: First column to delete, either as a 1-based integer or an
                Excel column label.
            amount: Number of columns to delete. ``0`` is a no-op.

        Raises:
            ValueError: If ``idx`` or ``amount`` is outside the supported
                range.
        """
        _delete_cols(self, idx, amount)

    def move_range(
        self,
        cell_range: Any,
        rows: int = 0,
        cols: int = 0,
        translate: bool = False,
    ) -> None:
        """Move a rectangular block of cells by *rows* / *cols*.

        This is a paste-style relocation: every cell inside ``cell_range``
        is moved to ``(row + rows, col + cols)``. Existing destination
        cells are overwritten, matching openpyxl.

        Formulas inside the moved block are paste-translated:
        relative refs shift by ``(rows, cols)``; absolute refs (``$A$1``)
        DO NOT shift. With ``translate=True``, formulas in cells
        outside the moved block that reference cells inside the source
        rectangle are also re-anchored to the new location.

        Args:
            cell_range: A1 range string such as ``"C3:E10"`` or a single
                cell reference such as ``"A1"``.
            rows: Signed row offset.
            cols: Signed column offset.
            translate: Whether formulas outside the moved block that point
                into the source range should be re-anchored.

        Raises:
            TypeError: If ``rows`` or ``cols`` is not an integer.
            ValueError: If the destination would fall outside Excel's
                coordinate limits.
        """
        _move_range(self, cell_range, rows=rows, cols=cols, translate=translate)

    def _move_cell(
        self,
        row: int,
        column: int,
        row_offset: int,
        col_offset: int,
        translate: bool = False,
    ) -> None:
        """openpyxl private helper for moving one cell in memory."""
        _move_cell(self, row, column, row_offset, col_offset, translate)

    def _move_cells(
        self,
        min_row: int | None = None,
        min_col: int | None = None,
        offset: int = 0,
        row_or_col: str = "row",
    ) -> None:
        """openpyxl private helper for shifting cells from a row/column boundary."""
        if row_or_col == "row":
            if min_row is None:
                raise TypeError("min_row is required when row_or_col='row'")
            if offset > 0:
                _insert_rows(self, min_row, offset)
            else:
                _delete_rows(self, min_row + offset, -offset)
            return
        if min_col is None:
            raise TypeError("min_col is required when row_or_col='column'")
        if offset > 0:
            _insert_cols(self, min_col, offset)
        else:
            _delete_cols(self, min_col + offset, -offset)

    # ------------------------------------------------------------------
    # Lazy per-sheet read maps.
    # ------------------------------------------------------------------
    #
    # Each method caches the fully-materialized result in a per-worksheet
    # slot on first access. That single FFI hop populates the cache;
    # subsequent calls (including per-cell lookups like ``cell.comment``)
    # are O(1) dict probes in Python. The Rust reader memoizes too, so
    # even a cold cache is cheap — but avoiding PyObject conversion on
    # every cell-level access is the dominant win here.

    def _get_comments_map(self) -> dict[str, Any]:
        """Return ``{cell_ref: Comment}`` for this sheet, cached per instance."""
        return get_comments_map(self)

    def _get_hyperlinks_map(self) -> dict[str, Any]:
        """Return ``{cell_ref: Hyperlink}`` for this sheet, cached per instance."""
        return get_hyperlinks_map(self)

    # ------------------------------------------------------------------
    # Worksheet collections share the same lazy-cache contract as
    # comments/hyperlinks: one Rust call per sheet, then dict/list-shaped
    # caches for subsequent reads.
    # ------------------------------------------------------------------

    @property
    def tables(self) -> Any:
        """Return the openpyxl-shaped table mapping for this sheet.

        Loaded from the Rust reader on first access. In write mode the
        mapping starts empty and is populated by ``add_table()``.
        """
        return get_tables_map(self)

    @property
    def _tables(self) -> Any:
        """openpyxl private alias for the worksheet table mapping."""
        return self.tables

    @property
    def pivot_tables(self) -> list[Any]:
        """Return the list of pivot tables on this sheet.

        In modify mode the list is populated lazily by walking the
        sheet's relationship graph and parsing the on-disk pivot
        table parts. Each entry is a
        :class:`~wolfxl.pivot.PivotTableHandle` whose ``.source`` and
        layout fields can be reassigned to mutate the underlying pivot
        table/cache XML on save.

        In construction mode the list is empty: pivots queued via
        :meth:`add_pivot_table` are tracked separately on
        ``_pending_pivot_tables`` until save.

        Layout edits that can change derived cache records mark the
        linked cache refresh-on-open; source edits with matching shape
        stay byte-stable.
        """
        if self._pivot_handles_cache is not None:
            return self._pivot_handles_cache
        from wolfxl.pivot._load import load_pivot_tables_for_sheet

        self._pivot_handles_cache = load_pivot_tables_for_sheet(
            self._workbook, self
        )
        return self._pivot_handles_cache

    def add_table(self, table: Any) -> None:
        """Attach a table to this worksheet.

        Tables are supported for both new workbooks and workbooks opened
        with ``load_workbook(..., modify=True)``. WolfXL assigns a
        workbook-unique table id during save, so any explicit id on the
        supplied table object is ignored.

        Args:
            table: Table object to attach to this worksheet.
        """
        _add_table(self, table)

    @property
    def data_validations(self) -> Any:
        """Return the ``DataValidationList`` for this sheet (lazy-loaded)."""
        return get_data_validations(self)

    def add_data_validation(
        self,
        dv: Any | None = None,
        *,
        data_validation: Any | None = None,
    ) -> None:
        """openpyxl-style alias for ``ws.data_validations.append(dv)``."""
        dv = dv if dv is not None else data_validation
        if dv is None:
            raise TypeError("add_data_validation() missing required argument: 'data_validation'")
        _add_data_validation(self, dv)

    @property
    def conditional_formatting(self) -> Any:
        """Return the ``ConditionalFormattingList`` for this sheet."""
        return get_conditional_formatting(self)

    # ------------------------------------------------------------------
    # Flush pending writes to Rust
    # ------------------------------------------------------------------

    def _flush(self) -> None:
        """Write all pending changes to the active Rust backend."""
        flush_worksheet(self)

    def _flush_to_writer(
        self, writer: Any, python_value_to_payload: Any,
        font_to_format_dict: Any, fill_to_format_dict: Any,
        alignment_to_format_dict: Any, border_to_rust_dict: Any,
        protection_to_format_dict: Any,
    ) -> None:
        """Flush dirty cells to the NativeWorkbook backend (write mode).

        Values are batched into a single ``write_sheet_values()`` call when
        possible (int/float/str/bool/None), eliminating per-cell FFI overhead.
        Dates, datetimes, and other non-primitive values fall through to
        per-cell ``write_cell_value()`` with type-preserving payload dicts.
        """
        flush_to_writer(
            self,
            writer,
            python_value_to_payload,
            font_to_format_dict,
            fill_to_format_dict,
            alignment_to_format_dict,
            border_to_rust_dict,
            cellrichtext_to_runs_payload,
            protection_to_format_dict,
        )

    def _batch_write_dicts(
        self,
        batch_fn: Any,
        entries: list[tuple[int, int, dict[str, Any]]],
    ) -> None:
        """Build a bounding-box grid of dicts and call a batch Rust method."""
        batch_write_dicts(self, batch_fn, entries)

    def _batch_write_style_ids(
        self,
        batch_fn: Any,
        entries: list[tuple[int, int, int]],
    ) -> None:
        """Build a bounding-box grid of style IDs and call a batch Rust method."""
        batch_write_style_ids(self, batch_fn, entries)

    def _flush_to_patcher(
        self, patcher: Any, python_value_to_payload: Any,
        font_to_format_dict: Any, fill_to_format_dict: Any,
        alignment_to_format_dict: Any, border_to_rust_dict: Any,
        protection_to_format_dict: Any,
    ) -> None:
        """Flush dirty cells to the XlsxPatcher backend (modify mode)."""
        flush_to_patcher(
            self,
            patcher,
            python_value_to_payload,
            font_to_format_dict,
            fill_to_format_dict,
            alignment_to_format_dict,
            border_to_rust_dict,
            cellrichtext_to_runs_payload,
            protection_to_format_dict,
        )

    def _flush_autofilter_post_cells(self, writer: Any) -> None:
        """Flush write-mode auto-filter metadata after cell values."""
        flush_autofilter_post_cells(self, writer)

    def _flush_compat_properties(self, writer: Any) -> None:
        """Flush openpyxl compatibility metadata to the write-mode backend."""
        flush_compat_properties(self, writer)

    # ------------------------------------------------------------------
    # wolfxl-core classifier bridge (delegates to the single Rust
    # classifier that `wolfxl schema --format json` also goes through —
    # so Python callers and the CLI can never drift in their answers).
    # ------------------------------------------------------------------

    def classify_format(self, fmt: str) -> str:
        """Classify an Excel number-format string (e.g. ``"$#,##0.00"``).

        Returns the same category string the CLI's ``schema`` subcommand
        emits in the per-column ``format`` field: ``"general"``,
        ``"currency"``, ``"percentage"``, ``"scientific"``, ``"date"``,
        ``"time"``, ``"datetime"``, ``"integer"``, ``"float"``, or
        ``"text"``. The method is an
        instance method for discoverability; it doesn't use any
        worksheet state.

        Args:
            fmt: Excel number-format string.

        Returns:
            Category string such as ``"general"``, ``"currency"``, or
            ``"date"``.
        """
        return _classify_worksheet_format(fmt)

    def schema(self) -> dict[str, Any]:
        """Infer this worksheet's schema via ``wolfxl_core::infer_sheet_schema``.

        Returns a dict shaped exactly like one entry of
        ``wolfxl schema <file> --format json``'s ``sheets`` array:

        .. code-block:: python

            {
                "name": "Sheet1",
                "rows": 50,
                "columns": [
                    {"name": "Account", "type": "string",
                     "format": "general", "null_count": 0,
                     "unique_count": 12, "unique_capped": false,
                     "cardinality": "categorical",
                     "samples": ["Revenue", "COGS", ...]},
                    ...
                ],
            }

        Builds two parallel grids — values and per-cell
        ``number_format`` strings — from ``cell_records()`` so the
        bridge sees the same format metadata the CLI does. Without the
        format grid, currency / percentage / date columns would
        classify as ``general`` and the Python answer would silently
        drift from the CLI's. Pending in-memory ``number_format`` edits
        are overlaid before inference so unsaved worksheet changes are
        included too.

        Returns:
            Dict shaped like a single ``wolfxl schema --format json`` sheet
            entry, with ``name``, ``rows``, and ``columns`` keys.
        """
        return _infer_worksheet_schema(self)

    def __repr__(self) -> str:
        """Return a compact debug representation for this worksheet.

        Returns:
            A string containing the worksheet title.
        """
        return _worksheet_repr(self._title)


Worksheet.add_data_validation.__signature__ = _inspect.Signature(
    [
        _inspect.Parameter("self", _inspect.Parameter.POSITIONAL_OR_KEYWORD),
        _inspect.Parameter(
            "data_validation", _inspect.Parameter.POSITIONAL_OR_KEYWORD
        ),
    ]
)
Worksheet.add_pivot.__signature__ = _inspect.Signature(
    [
        _inspect.Parameter("self", _inspect.Parameter.POSITIONAL_OR_KEYWORD),
        _inspect.Parameter("pivot", _inspect.Parameter.POSITIONAL_OR_KEYWORD),
    ]
)
Worksheet.__signature__ = _inspect.Signature(
    [
        _inspect.Parameter("parent", _inspect.Parameter.POSITIONAL_OR_KEYWORD),
        _inspect.Parameter(
            "title",
            _inspect.Parameter.POSITIONAL_OR_KEYWORD,
            default=None,
        ),
    ]
)
