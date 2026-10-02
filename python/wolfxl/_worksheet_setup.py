"""Worksheet setup, print, view, protection, and page-break helpers."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from wolfxl._worksheet import Worksheet


def get_freeze_panes(ws: Worksheet) -> str | None:
    """Return the worksheet freeze-pane top-left cell, if any."""
    workbook = ws._workbook  # noqa: SLF001
    if getattr(workbook, "_rust_reader", None) is not None and ws._freeze_panes is None:  # noqa: SLF001
        info = workbook._rust_reader.read_freeze_panes(ws._title)  # noqa: SLF001
        if info and info.get("mode"):
            return info.get("top_left_cell")
        return None
    return ws._freeze_panes  # noqa: SLF001


def set_freeze_panes(ws: Worksheet, value: str | None) -> None:
    """Set freeze panes and mirror the state into ``sheet_view.pane``."""
    if value is not None and hasattr(value, "coordinate"):
        value = value.coordinate
    if value == "":
        value = None
    if value == "A1":
        value = None
    ws._freeze_panes = value  # noqa: SLF001

    from wolfxl.worksheet.views import Pane, Selection

    sheet_view = get_sheet_view(ws)
    if value is None:
        sheet_view.pane = None
        sheet_view.selection = [Selection()]
        return

    from wolfxl._utils import a1_to_rowcol

    try:
        row, col = a1_to_rowcol(value)
    except Exception:
        return

    frozen_rows = row - 1
    frozen_cols = col - 1
    if frozen_rows and frozen_cols:
        active_pane = "bottomRight"
        selection = [
            Selection(pane="topRight", activeCell=None, sqref=None),
            Selection(pane="bottomLeft", activeCell=None, sqref=None),
            Selection(pane="bottomRight", activeCell="A1", sqref="A1"),
        ]
    elif frozen_rows:
        active_pane = "bottomLeft"
        selection = [Selection(pane="bottomLeft", activeCell="A1", sqref="A1")]
    elif frozen_cols:
        active_pane = "topRight"
        selection = [Selection(pane="topRight", activeCell="A1", sqref="A1")]
    else:
        sheet_view.pane = None
        sheet_view.selection = [Selection()]
        return

    sheet_view.pane = Pane(
        xSplit=frozen_cols if frozen_cols else None,
        ySplit=frozen_rows if frozen_rows else None,
        topLeftCell=value,
        activePane=active_pane,
        state="frozen",
    )
    sheet_view.selection = selection


def get_auto_filter(ws: Worksheet) -> Any:
    """Return the worksheet auto-filter proxy, loading reader metadata once."""
    workbook = ws._workbook  # noqa: SLF001
    auto_filter = ws._auto_filter  # noqa: SLF001
    if (
        auto_filter.ref is None
        and not auto_filter.filter_columns
        and auto_filter.sort_state is None
        and workbook._rust_reader is not None  # noqa: SLF001
        and hasattr(workbook._rust_reader, "read_auto_filter")  # noqa: SLF001
    ):
        loaded = auto_filter_from_payload(
            workbook._rust_reader.read_auto_filter(ws._title)  # noqa: SLF001
        )
        if loaded is not None:
            auto_filter.ref = loaded.ref
            auto_filter.filter_columns = loaded.filter_columns
            auto_filter.sort_state = loaded.sort_state
    return auto_filter


def auto_filter_from_payload(payload: Any) -> Any:
    """Build an ``AutoFilter`` from a reader payload, or ``None`` when absent."""
    if not isinstance(payload, dict):
        return None
    from wolfxl.worksheet.filters import AutoFilter

    return AutoFilter(
        ref=payload.get("ref"),
        filter_columns=[
            _filter_column_from_payload(column)
            for column in payload.get("filter_columns", [])
            if isinstance(column, dict)
        ],
        sort_state=_sort_state_from_payload(payload.get("sort_state")),
    )


def _filter_column_from_payload(payload: dict[str, Any]) -> Any:
    from wolfxl.worksheet.filters import FilterColumn

    return FilterColumn(
        col_id=int(payload.get("col_id", 0)),
        hidden_button=bool(payload.get("hidden_button", False)),
        show_button=bool(payload.get("show_button", True)),
        filter=_filter_from_payload(payload.get("filter")),
        date_group_items=[
            _date_group_item_from_payload(item)
            for item in payload.get("date_group_items", [])
            if isinstance(item, dict)
        ],
    )


def _filter_from_payload(payload: Any) -> Any:
    if not isinstance(payload, dict):
        return None
    from wolfxl.worksheet.filters import (
        BlankFilter,
        ColorFilter,
        CustomFilter,
        CustomFilters,
        DynamicFilter,
        Filters,
        IconFilter,
        Top10,
    )

    kind = payload.get("kind")
    if kind == "blank":
        return BlankFilter()
    if kind == "color":
        return ColorFilter(
            dxf_id=int(payload.get("dxf_id", 0)),
            cell_color=bool(payload.get("cell_color", True)),
        )
    if kind == "custom":
        return CustomFilters(
            customFilter=[
                CustomFilter(
                    operator=str(item.get("operator", "equal")),
                    val=str(item.get("val", "")),
                )
                for item in payload.get("filters", [])
                if isinstance(item, dict)
            ],
            and_=bool(payload.get("and_", False)),
        )
    if kind == "dynamic":
        return DynamicFilter(
            type=str(payload.get("type", "null")),
            val=payload.get("val"),
            val_iso=payload.get("val_iso"),
            max_val_iso=payload.get("max_val_iso"),
        )
    if kind == "icon":
        return IconFilter(
            icon_set=str(payload.get("icon_set", "3Arrows")),
            icon_id=int(payload.get("icon_id", 0)),
        )
    if kind == "string":
        return Filters(filter=[str(value) for value in payload.get("values", [])])
    if kind == "top10":
        return Top10(
            top=bool(payload.get("top", True)),
            percent=bool(payload.get("percent", False)),
            val=float(payload.get("val", 10.0)),
            filter_val=payload.get("filter_val"),
        )
    return None


def _date_group_item_from_payload(payload: dict[str, Any]) -> Any:
    from wolfxl.worksheet.filters import DateGroupItem

    return DateGroupItem(
        year=int(payload.get("year", 0)),
        month=payload.get("month"),
        day=payload.get("day"),
        hour=payload.get("hour"),
        minute=payload.get("minute"),
        second=payload.get("second"),
        date_time_grouping=str(payload.get("date_time_grouping", "year")),
    )


def _sort_state_from_payload(payload: Any) -> Any:
    if not isinstance(payload, dict):
        return None
    from wolfxl.worksheet.filters import SortState

    return SortState(
        sort_conditions=[
            _sort_condition_from_payload(condition)
            for condition in payload.get("sort_conditions", [])
            if isinstance(condition, dict)
        ],
        column_sort=bool(payload.get("column_sort", False)),
        case_sensitive=bool(payload.get("case_sensitive", False)),
        ref=payload.get("ref"),
    )


def _sort_condition_from_payload(payload: dict[str, Any]) -> Any:
    from wolfxl.worksheet.filters import SortCondition

    return SortCondition(
        ref=str(payload.get("ref", "")),
        descending=bool(payload.get("descending", False)),
        sort_by=str(payload.get("sort_by", "value")),
        custom_list=payload.get("custom_list"),
        dxf_id=payload.get("dxf_id"),
        icon_set=payload.get("icon_set"),
        icon_id=payload.get("icon_id"),
    )


def get_page_setup(ws: Worksheet) -> Any:
    """Return the lazy page setup object."""
    if ws._page_setup is None:  # noqa: SLF001
        payload = _reader_payload(ws, "read_page_setup")
        ws._page_setup = _page_setup_from_payload(payload)  # noqa: SLF001
    return ws._page_setup  # noqa: SLF001


def get_print_options(ws: Worksheet) -> Any:
    """Return the lazy print options object."""
    if ws._print_options is None:  # noqa: SLF001
        payload = _reader_payload(ws, "read_print_options")
        ws._print_options = _print_options_from_payload(payload)  # noqa: SLF001
    return ws._print_options  # noqa: SLF001


def get_page_margins(ws: Worksheet) -> Any:
    """Return the lazy page margins object."""
    if ws._page_margins is None:  # noqa: SLF001
        payload = _reader_payload(ws, "read_page_margins")
        ws._page_margins = _page_margins_from_payload(payload)  # noqa: SLF001
    return ws._page_margins  # noqa: SLF001


def get_header_footer(ws: Worksheet) -> Any:
    """Return the lazy header/footer object."""
    if ws._header_footer is None:  # noqa: SLF001
        payload = _reader_payload(ws, "read_header_footer")
        ws._header_footer = _header_footer_from_payload(payload)  # noqa: SLF001
    return ws._header_footer  # noqa: SLF001


def get_sheet_properties(ws: Worksheet) -> Any:
    """Return the lazy worksheet properties object."""
    if ws._sheet_properties is None:  # noqa: SLF001
        payload = _reader_payload(ws, "read_sheet_properties")
        ws._sheet_properties = _sheet_properties_from_payload(payload)  # noqa: SLF001
    return ws._sheet_properties  # noqa: SLF001


def _reader_payload(ws: Worksheet, method_name: str) -> Any:
    reader = getattr(ws._workbook, "_rust_reader", None)  # noqa: SLF001
    if reader is None or not hasattr(reader, method_name):
        return None
    return getattr(reader, method_name)(ws._title)  # noqa: SLF001


def _page_setup_from_payload(payload: Any) -> Any:
    from wolfxl.worksheet.page_setup import PageSetup

    if not isinstance(payload, dict):
        return PageSetup()
    return PageSetup(
        orientation=payload.get("orientation"),
        paperSize=payload.get("paper_size"),
        fitToWidth=payload.get("fit_to_width"),
        fitToHeight=payload.get("fit_to_height"),
        scale=payload.get("scale"),
        firstPageNumber=payload.get("first_page_number"),
        horizontalDpi=payload.get("horizontal_dpi"),
        verticalDpi=payload.get("vertical_dpi"),
        cellComments=payload.get("cell_comments"),
        errors=payload.get("errors"),
        useFirstPageNumber=payload.get("use_first_page_number"),
        paperHeight=payload.get("paper_height"),
        paperWidth=payload.get("paper_width"),
        pageOrder=payload.get("page_order"),
        usePrinterDefaults=payload.get("use_printer_defaults"),
        blackAndWhite=payload.get("black_and_white"),
        draft=payload.get("draft"),
        copies=payload.get("copies"),
    )


def _print_options_from_payload(payload: Any) -> Any:
    from wolfxl.worksheet.page import PrintOptions

    if not isinstance(payload, dict):
        return PrintOptions()
    return PrintOptions(
        horizontalCentered=payload.get("horizontal_centered"),
        verticalCentered=payload.get("vertical_centered"),
        headings=payload.get("headings"),
        gridLines=payload.get("grid_lines"),
        gridLinesSet=payload.get("grid_lines_set"),
    )


def _page_margins_from_payload(payload: Any) -> Any:
    from wolfxl.worksheet.page_setup import PageMargins

    if not isinstance(payload, dict):
        return PageMargins()
    return PageMargins(
        left=float(payload.get("left", 0.75)),
        right=float(payload.get("right", 0.75)),
        top=float(payload.get("top", 1.0)),
        bottom=float(payload.get("bottom", 1.0)),
        header=float(payload.get("header", 0.5)),
        footer=float(payload.get("footer", 0.5)),
    )


def _header_footer_from_payload(payload: Any) -> Any:
    from wolfxl.worksheet.header_footer import HeaderFooter

    if not isinstance(payload, dict):
        return HeaderFooter()
    return HeaderFooter(
        odd_header=_header_footer_item_from_payload(payload.get("odd_header")),
        odd_footer=_header_footer_item_from_payload(payload.get("odd_footer")),
        even_header=_header_footer_item_from_payload(payload.get("even_header")),
        even_footer=_header_footer_item_from_payload(payload.get("even_footer")),
        first_header=_header_footer_item_from_payload(payload.get("first_header")),
        first_footer=_header_footer_item_from_payload(payload.get("first_footer")),
        different_odd_even=bool(payload.get("different_odd_even", False)),
        different_first=bool(payload.get("different_first", False)),
        scale_with_doc=bool(payload.get("scale_with_doc", True)),
        align_with_margins=bool(payload.get("align_with_margins", True)),
    )


def _header_footer_item_from_payload(payload: Any) -> Any:
    from wolfxl.worksheet.header_footer import HeaderFooterItem

    if not isinstance(payload, dict):
        return HeaderFooterItem()
    return HeaderFooterItem(
        left=payload.get("left"),
        center=payload.get("center"),
        right=payload.get("right"),
    )


def _sheet_properties_from_payload(payload: Any) -> Any:
    from wolfxl.worksheet.properties import (
        Outline,
        PageSetupProperties,
        WorksheetProperties,
    )

    if not isinstance(payload, dict):
        return WorksheetProperties()
    outline_payload = payload.get("outline")
    page_setup_payload = payload.get("page_setup")
    return WorksheetProperties(
        codeName=payload.get("code_name"),
        enableFormatConditionsCalculation=payload.get(
            "enable_format_conditions_calculation"
        ),
        filterMode=payload.get("filter_mode"),
        published=payload.get("published"),
        syncHorizontal=payload.get("sync_horizontal"),
        syncRef=payload.get("sync_ref"),
        syncVertical=payload.get("sync_vertical"),
        transitionEvaluation=payload.get("transition_evaluation"),
        transitionEntry=payload.get("transition_entry"),
        tabColor=_normalize_sheet_property_color(payload.get("tab_color")),
        outlinePr=(
            Outline(
                summaryBelow=bool(outline_payload.get("summary_below", True)),
                summaryRight=bool(outline_payload.get("summary_right", True)),
                applyStyles=outline_payload.get("apply_styles"),
                showOutlineSymbols=outline_payload.get("show_outline_symbols"),
            )
            if isinstance(outline_payload, dict)
            else Outline()
        ),
        pageSetUpPr=(
            PageSetupProperties(
                autoPageBreaks=page_setup_payload.get("auto_page_breaks"),
                fitToPage=page_setup_payload.get("fit_to_page"),
            )
            if isinstance(page_setup_payload, dict)
            else PageSetupProperties()
        ),
    )


def _normalize_sheet_property_color(value: Any) -> str | None:
    if not isinstance(value, str):
        return None
    color = value.removeprefix("#")
    if len(color) == 6:
        return "00" + color
    return color


def get_sheet_view(ws: Worksheet) -> Any:
    """Return the lazy sheet view object, carrying pending freeze panes."""
    if ws._sheet_view is None:  # noqa: SLF001
        payload = _reader_payload(ws, "read_sheet_view")
        sheet_view = _sheet_view_from_payload(payload)
        if ws._freeze_panes is not None:  # noqa: SLF001
            from wolfxl._utils import a1_to_rowcol
            from wolfxl.worksheet.views import Pane

            try:
                row, col = a1_to_rowcol(ws._freeze_panes)  # noqa: SLF001
                sheet_view.pane = Pane(
                    xSplit=float(col - 1) if col > 1 else None,
                    ySplit=float(row - 1) if row > 1 else None,
                    topLeftCell=ws._freeze_panes,  # noqa: SLF001
                    activePane="bottomRight",
                    state="frozen",
                )
            except Exception:
                pass
        ws._sheet_view = sheet_view  # noqa: SLF001
    return ws._sheet_view  # noqa: SLF001


def _sheet_view_from_payload(payload: Any) -> Any:
    from wolfxl.worksheet.views import Pane, Selection, SheetView

    if not isinstance(payload, dict):
        return SheetView()
    pane_payload = payload.get("pane")
    return SheetView(
        zoomScale=int(payload.get("zoom_scale", 100)),
        zoomScaleNormal=int(payload.get("zoom_scale_normal", 100)),
        view=str(payload.get("view") or "normal"),
        showGridLines=payload.get("show_grid_lines"),
        showRowColHeaders=payload.get("show_row_col_headers"),
        showOutlineSymbols=payload.get("show_outline_symbols"),
        showZeros=payload.get("show_zeros"),
        rightToLeft=payload.get("right_to_left"),
        tabSelected=payload.get("tab_selected"),
        topLeftCell=payload.get("top_left_cell"),
        workbookViewId=int(payload.get("workbook_view_id", 0)),
        pane=_pane_from_payload(pane_payload, Pane),
        selection=[
            Selection(
                activeCell=str(item.get("active_cell") or "A1"),
                sqref=str(item.get("sqref") or item.get("active_cell") or "A1"),
                pane=item.get("pane"),
                activeCellId=item.get("active_cell_id"),
            )
            for item in payload.get("selection", [])
            if isinstance(item, dict)
        ]
        or [Selection()],
    )


def _pane_from_payload(payload: Any, pane_cls: Any) -> Any:
    if not isinstance(payload, dict):
        return None
    x_split = float(payload.get("x_split", 0.0))
    y_split = float(payload.get("y_split", 0.0))
    return pane_cls(
        xSplit=x_split if x_split else None,
        ySplit=y_split if y_split else None,
        topLeftCell=str(payload.get("top_left_cell") or "A1"),
        activePane=str(payload.get("active_pane") or "topLeft"),
        state=str(payload.get("state") or "frozen"),
    )


def get_protection(ws: Worksheet) -> Any:
    """Return the lazy sheet protection object."""
    if ws._protection is None:  # noqa: SLF001
        workbook = ws._workbook  # noqa: SLF001
        payload = None
        reader = getattr(workbook, "_rust_reader", None)
        if reader is not None and hasattr(reader, "read_sheet_protection"):
            payload = reader.read_sheet_protection(ws._title)  # noqa: SLF001
        ws._protection = _sheet_protection_from_payload(payload)  # noqa: SLF001
    return ws._protection  # noqa: SLF001


def _sheet_protection_from_payload(payload: Any) -> Any:
    from wolfxl.worksheet.protection import SheetProtection

    if not isinstance(payload, dict):
        return SheetProtection()
    protection = SheetProtection(
        sheet=bool(payload.get("sheet", False)),
        objects=bool(payload.get("objects", False)),
        scenarios=bool(payload.get("scenarios", False)),
        formatCells=bool(payload.get("format_cells", True)),
        formatColumns=bool(payload.get("format_columns", True)),
        formatRows=bool(payload.get("format_rows", True)),
        insertColumns=bool(payload.get("insert_columns", True)),
        insertRows=bool(payload.get("insert_rows", True)),
        insertHyperlinks=bool(payload.get("insert_hyperlinks", True)),
        deleteColumns=bool(payload.get("delete_columns", True)),
        deleteRows=bool(payload.get("delete_rows", True)),
        selectLockedCells=bool(payload.get("select_locked_cells", False)),
        sort=bool(payload.get("sort", True)),
        autoFilter=bool(payload.get("auto_filter", True)),
        pivotTables=bool(payload.get("pivot_tables", True)),
        selectUnlockedCells=bool(payload.get("select_unlocked_cells", False)),
    )
    protection.password = payload.get("password_hash")
    return protection


def get_row_breaks(ws: Worksheet) -> Any:
    """Return the lazy row page-break list."""
    if ws._row_breaks is None:  # noqa: SLF001
        payload = _reader_payload(ws, "read_row_breaks")
        ws._row_breaks = _page_breaks_from_payload(payload, "row")  # noqa: SLF001
    return ws._row_breaks  # noqa: SLF001


def get_col_breaks(ws: Worksheet) -> Any:
    """Return the lazy column page-break list."""
    if ws._col_breaks is None:  # noqa: SLF001
        payload = _reader_payload(ws, "read_column_breaks")
        ws._col_breaks = _page_breaks_from_payload(payload, "column")  # noqa: SLF001
    return ws._col_breaks  # noqa: SLF001


def get_sheet_format(ws: Worksheet) -> Any:
    """Return the lazy sheet format properties object."""
    if ws._sheet_format is None:  # noqa: SLF001
        payload = _reader_payload(ws, "read_sheet_format")
        ws._sheet_format = _sheet_format_from_payload(payload)  # noqa: SLF001
    return ws._sheet_format  # noqa: SLF001


def _page_breaks_from_payload(payload: Any, kind: str) -> Any:
    from wolfxl.worksheet.pagebreak import ColBreak, PageBreakList, RowBreak

    if not isinstance(payload, dict):
        return PageBreakList()
    cls = RowBreak if kind == "row" else ColBreak
    breaks = [
        cls(
            id=int(item.get("id", 0)),
            min=item.get("min"),
            max=item.get("max"),
            man=bool(item.get("man", True)),
            pt=bool(item.get("pt", False)),
        )
        for item in payload.get("breaks", [])
        if isinstance(item, dict)
    ]
    page_breaks = PageBreakList(breaks=breaks)
    page_breaks.count = int(payload.get("count", page_breaks.count))
    page_breaks.manualBreakCount = int(
        payload.get("manual_break_count", page_breaks.manualBreakCount)
    )
    return page_breaks


def _sheet_format_from_payload(payload: Any) -> Any:
    from wolfxl.worksheet.dimensions import SheetFormatProperties

    if not isinstance(payload, dict):
        return SheetFormatProperties()
    return SheetFormatProperties(
        baseColWidth=int(payload.get("base_col_width", 8)),
        defaultColWidth=payload.get("default_col_width"),
        defaultRowHeight=float(payload.get("default_row_height", 15.0)),
        customHeight=payload.get("custom_height"),
        zeroHeight=payload.get("zero_height"),
        thickTop=payload.get("thick_top"),
        thickBottom=payload.get("thick_bottom"),
        outlineLevelRow=int(payload.get("outline_level_row", 0)),
        outlineLevelCol=int(payload.get("outline_level_col", 0)),
    )


def get_dimension_holder(ws: Worksheet) -> Any:
    """Return a fresh dimension holder bound to a worksheet."""
    from wolfxl.worksheet.dimensions import DimensionHolder

    return DimensionHolder(ws)


def to_rust_page_breaks_dict(ws: Worksheet) -> dict[str, Any]:
    """Return the Rust payload shape for row and column page breaks."""
    return {
        "row_breaks": (
            ws._row_breaks.to_rust_dict()  # noqa: SLF001
            if ws._row_breaks is not None and len(ws._row_breaks) > 0  # noqa: SLF001
            else None
        ),
        "col_breaks": (
            ws._col_breaks.to_rust_dict()  # noqa: SLF001
            if ws._col_breaks is not None and len(ws._col_breaks) > 0  # noqa: SLF001
            else None
        ),
    }


def to_rust_sheet_format_dict(ws: Worksheet) -> dict[str, Any] | None:
    """Return the Rust payload shape for sheet format properties."""
    if ws._sheet_format is None or ws._sheet_format.is_default():  # noqa: SLF001
        return None
    return ws._sheet_format.to_rust_dict()  # noqa: SLF001


def set_print_title_rows(ws: Worksheet, value: str | None) -> None:
    """Set repeat rows for printed pages."""
    _ensure_print_titles_loaded(ws)
    if value is not None:
        from wolfxl.worksheet.print_settings import RowRange

        ws._print_title_rows = str(RowRange.from_string(value))  # noqa: SLF001
    else:
        ws._print_title_rows = None  # noqa: SLF001
    ws._print_titles_dirty = True  # noqa: SLF001


def set_print_title_cols(ws: Worksheet, value: str | None) -> None:
    """Set repeat columns for printed pages."""
    _ensure_print_titles_loaded(ws)
    if value is not None:
        from wolfxl.worksheet.print_settings import ColRange

        ws._print_title_cols = str(ColRange.from_string(value))  # noqa: SLF001
    else:
        ws._print_title_cols = None  # noqa: SLF001
    ws._print_titles_dirty = True  # noqa: SLF001


def _ensure_print_titles_loaded(ws: Worksheet) -> None:
    """Load both print-title selectors before mutating one side."""
    if (
        ws._print_titles_dirty  # noqa: SLF001
        or ws._print_title_rows is not None  # noqa: SLF001
        or ws._print_title_cols is not None  # noqa: SLF001
    ):
        return
    reader = getattr(ws._workbook, "_rust_reader", None)  # noqa: SLF001
    if reader is None or not hasattr(reader, "read_print_titles"):
        return
    payload = reader.read_print_titles(ws._title)  # noqa: SLF001
    if isinstance(payload, dict):
        ws._print_title_rows = payload.get("rows")  # noqa: SLF001
        ws._print_title_cols = payload.get("cols")  # noqa: SLF001


def to_rust_setup_dict(ws: Worksheet) -> dict[str, Any]:
    """Return the Rust payload shape for worksheet setup metadata."""
    payload: dict[str, Any] = {}
    payload["page_setup"] = (
        ws._page_setup.to_rust_dict()  # noqa: SLF001
        if ws._page_setup is not None and not ws._page_setup.is_default()  # noqa: SLF001
        else None
    )
    payload["page_margins"] = (
        ws._page_margins.to_rust_dict()  # noqa: SLF001
        if ws._page_margins is not None and not ws._page_margins.is_default()  # noqa: SLF001
        else None
    )
    payload["print_options"] = (
        ws._print_options.to_rust_dict()  # noqa: SLF001
        if ws._print_options is not None and not ws._print_options.is_default()  # noqa: SLF001
        else None
    )
    payload["header_footer"] = (
        ws._header_footer.to_rust_dict()  # noqa: SLF001
        if ws._header_footer is not None and not ws._header_footer.is_default()  # noqa: SLF001
        else None
    )
    payload["sheet_view"] = (
        ws._sheet_view.to_rust_dict()  # noqa: SLF001
        if ws._sheet_view is not None and not ws._sheet_view.is_default()  # noqa: SLF001
        else None
    )
    payload["sheet_protection"] = (
        ws._protection.to_rust_dict()  # noqa: SLF001
        if ws._protection is not None and not ws._protection.is_default()  # noqa: SLF001
        else None
    )
    if ws._print_title_rows is not None or ws._print_title_cols is not None:  # noqa: SLF001
        payload["print_titles"] = {
            "rows": ws._print_title_rows,  # noqa: SLF001
            "cols": ws._print_title_cols,  # noqa: SLF001
        }
    else:
        payload["print_titles"] = None
    return payload
