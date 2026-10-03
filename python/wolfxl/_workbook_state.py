"""Shared workbook state and constructor helpers."""

from __future__ import annotations

import os
import zipfile
from dataclasses import dataclass
from typing import Any
from xml.etree import ElementTree as ET

from wolfxl._worksheet import Worksheet
from wolfxl.xml.constants import XLTM, XLTX


@dataclass
class CopyOptions:
    """Per-workbook flags controlling :meth:`Workbook.copy_worksheet`.

    Attributes:
        deep_copy_images: When ``True``, drawings reachable from a cloned sheet
            have their referenced ``xl/media/imageN.<ext>`` targets cloned into
            freshly numbered media parts. When ``False`` (default), the cloned
            drawing relationships point at the same image bytes as the source.
    """

    deep_copy_images: bool = False


def same_existing_path(left: str, right: str | None) -> bool:
    """Return whether two paths identify the same existing filesystem entry."""
    if right is None:
        return False
    try:
        return os.path.samefile(left, right)
    except OSError:
        return os.path.abspath(left) == os.path.abspath(right)


def xlsb_xls_via_tempfile(
    rust_cls: Any,
    data: bytes | bytearray | memoryview,
    *,
    suffix: str,
    permissive: bool,
) -> tuple[Any, str]:
    """Materialize bytes to a tempfile and open them with a binary backend.

    Args:
        rust_cls: Rust-backed binary workbook class.
        data: Workbook bytes supplied to ``load_workbook``.
        suffix: File extension used for the temporary file.
        permissive: Whether to pass permissive parsing through to the backend.

    Returns:
        A ``(rust_book, tempfile_path)`` pair. The caller owns cleanup.
    """
    import tempfile

    with tempfile.NamedTemporaryFile(
        prefix="wolfxl-", suffix=suffix, delete=False
    ) as tmp:
        tmp.write(bytes(data))
        tmp_path = tmp.name

    opener = rust_cls.open
    try:
        rust_book = opener(tmp_path, permissive)
    except TypeError:
        rust_book = opener(tmp_path)
    return rust_book, tmp_path


def build_xlsx_wb(
    cls: type,
    *,
    rust_reader: Any,
    rust_patcher: Any | None,
    data_only: bool,
    read_only: bool,
    source_path: str | None,
    source_bytes: bytes | None = None,
    keep_links: bool = True,
    keep_vba: bool = False,
) -> Any:
    """Wire up read/modify-mode workbook fields shared by xlsx inputs.

    Args:
        cls: Workbook class to instantiate without calling ``__init__``.
        rust_reader: Open Rust reader with the workbook read API.
        rust_patcher: Optional open ``XlsxPatcher`` for modify mode.
        data_only: Whether formula cells should expose cached values.
        read_only: Whether streaming read mode was explicitly requested.
        source_path: Source path, or ``None`` for bytes-backed readers.

    Returns:
        A workbook instance with sheet proxies and pending queues initialized.
    """
    wb: Any = object.__new__(cls)
    wb._rust_writer = None
    wb._rust_patcher = rust_patcher
    wb._rust_reader = rust_reader
    wb._data_only = data_only
    wb._iso_dates = False
    wb.template = _source_is_template(
        source_path=source_path, source_bytes=source_bytes
    )
    wb._is_template = False
    wb.encoding = "utf-8"
    wb._rich_text = False
    wb._evaluator = None
    wb._read_only = read_only
    wb._source_path = source_path
    wb._source_bytes = source_bytes
    wb._keep_links = keep_links
    wb._keep_vba = keep_vba
    wb._strip_external_links_on_save = bool(rust_patcher is not None and not keep_links)
    wb._format = "xlsx"
    _initialize_sheet_proxies(wb, rust_reader)
    initialize_pending_state(wb)
    _hydrate_stylesheet(
        wb,
        source_path=source_path,
        source_bytes=source_bytes,
        read_only=read_only,
    )
    active_tab = _read_source_active_tab(
        source_path=source_path,
        source_bytes=source_bytes,
    )
    if active_tab is not None:
        wb._active_sheet_index = active_tab
    return wb


def _hydrate_stylesheet(
    wb: Any,
    *,
    source_path: str | None,
    source_bytes: bytes | None,
    read_only: bool = False,
) -> None:
    """Load workbook-level styles without walking worksheet XML.

    Older builds also parsed every worksheet XML document here to precompute
    merged-cell border variants. Cell border reads now resolve through the
    native reader on demand, so workbook open only needs `xl/styles.xml`.
    """
    try:
        from io import BytesIO

        from wolfxl.styles.stylesheet import apply_stylesheet

        if source_bytes is not None:
            with zipfile.ZipFile(BytesIO(source_bytes), "r") as archive:
                apply_stylesheet(archive, wb)
                if not read_only:
                    if _archive_has_merged_cells(archive):
                        _hydrate_merged_cell_borders(archive, wb)
                    else:
                        _mark_workbook_known_unmerged(wb)
        elif source_path is not None:
            with zipfile.ZipFile(source_path, "r") as archive:
                apply_stylesheet(archive, wb)
                if not read_only:
                    if _archive_has_merged_cells(archive):
                        _hydrate_merged_cell_borders(archive, wb)
                    else:
                        _mark_workbook_known_unmerged(wb)
    except Exception:
        return


def _archive_has_merged_cells(archive: Any) -> bool:
    """Return whether any worksheet XML contains merged-cell declarations."""
    for name in archive.namelist():
        if not (name.startswith("xl/worksheets/") and name.endswith(".xml")):
            continue
        try:
            if b"mergeCell" in archive.read(name):
                return True
        except Exception:
            continue
    return False


def _mark_workbook_known_unmerged(wb: Any) -> None:
    """Mark worksheet proxies when the source archive has no merged cells."""
    for ws in getattr(wb, "_sheets", {}).values():
        try:
            ws._merged_ranges_loaded = True  # noqa: SLF001
            ws._merged_ranges.clear()  # noqa: SLF001
            ws._collection_merged_ranges.clear()  # noqa: SLF001
        except Exception:
            continue


def _hydrate_merged_cell_borders(archive: Any, wb: Any) -> None:
    """Legacy eager merged-border helper kept for focused regression tests."""
    from wolfxl._styles import Border, Side
    from wolfxl.utils.cell import range_boundaries
    from wolfxl.xml.functions import localname

    def is_empty(side: Any) -> bool:
        return side is None or side == Side()

    def cell_key(ref: str) -> tuple[int, int]:
        min_col, min_row, _, _ = range_boundaries(ref)
        return min_row, min_col

    for name in archive.namelist():
        if not (name.startswith("xl/worksheets/") and name.endswith(".xml")):
            continue
        try:
            xml = archive.read(name)
        except Exception:
            continue
        if b"mergeCell" not in xml:
            continue
        try:
            root = ET.fromstring(xml)
        except Exception:
            continue
        style_by_cell: dict[tuple[int, int], int] = {}
        merge_refs: list[str] = []
        for node in root.iter():
            tag = localname(node)
            if tag == "c" and node.get("r") is not None:
                style_by_cell[cell_key(node.get("r"))] = int(node.get("s", "0") or 0)
            elif tag == "mergeCell" and node.get("ref"):
                merge_refs.append(node.get("ref"))
        for ref in merge_refs:
            min_col, min_row, max_col, max_row = range_boundaries(ref)
            start_style = style_by_cell.get((min_row, min_col))
            end_style = style_by_cell.get((max_row, max_col))
            if start_style is None or end_style is None:
                continue
            try:
                start_border = wb._borders[wb._cell_styles[start_style].borderId]  # noqa: SLF001
                end_border = wb._borders[wb._cell_styles[end_style].borderId]  # noqa: SLF001
            except Exception:
                continue
            merged_border = Border(
                left=start_border.left,
                right=end_border.right
                if not is_empty(end_border.right)
                else start_border.right,
                top=start_border.top,
                bottom=end_border.bottom
                if not is_empty(end_border.bottom)
                else start_border.bottom,
                diagonal=start_border.diagonal,
                diagonalUp=start_border.diagonalUp,
                diagonalDown=start_border.diagonalDown,
                outline=start_border.outline,
            )
            wb._borders.add(merged_border)  # noqa: SLF001
            for edge in ("top", "left", "right"):
                side = getattr(merged_border, edge)
                if is_empty(side):
                    continue
                wb._borders.add(Border(**{edge: side}))  # noqa: SLF001


def build_xlsb_xls_wb(
    cls: type,
    *,
    rust_book: Any,
    fmt: str,
    data_only: bool,
    source_path: str | None,
) -> Any:
    """Wire up read-mode workbook fields shared by xlsb and xls inputs."""
    wb: Any = object.__new__(cls)
    wb._rust_writer = None
    wb._rust_patcher = None
    wb._rust_reader = rust_book
    wb._data_only = data_only
    wb._iso_dates = False
    wb.template = False
    wb._is_template = False
    wb.encoding = "utf-8"
    wb._rich_text = False
    wb._evaluator = None
    wb._read_only = False
    wb._source_path = source_path
    wb._keep_links = True
    wb._format = fmt
    _initialize_sheet_proxies(wb, rust_book)
    initialize_pending_state(wb)
    return wb


def _source_is_template(
    *,
    source_path: str | None,
    source_bytes: bytes | None,
) -> bool:
    """Return whether the OOXML source declares a template workbook part."""
    data: bytes | None = None
    try:
        if source_bytes is not None:
            from io import BytesIO

            with zipfile.ZipFile(BytesIO(source_bytes), "r") as zf:
                data = zf.read("[Content_Types].xml")
        elif source_path is not None:
            with zipfile.ZipFile(source_path, "r") as zf:
                data = zf.read("[Content_Types].xml")
    except (OSError, KeyError, zipfile.BadZipFile):
        return False
    if not data:
        return False
    try:
        root = ET.fromstring(data)
    except ET.ParseError:
        return False
    for child in root:
        if (
            child.tag.rsplit("}", 1)[-1] == "Override"
            and child.get("PartName") == "/xl/workbook.xml"
            and child.get("ContentType") in {XLTX, XLTM}
        ):
            return True
    return False


def _read_source_active_tab(
    *,
    source_path: str | None,
    source_bytes: bytes | None,
) -> int | None:
    """Read workbook.xml activeTab exactly, including openpyxl's ``-1``."""
    from io import BytesIO

    try:
        if source_bytes is not None:
            archive_obj: str | BytesIO = BytesIO(source_bytes)
        elif source_path is not None:
            archive_obj = source_path
        else:
            return None
        with zipfile.ZipFile(archive_obj) as archive:
            workbook_xml = archive.read("xl/workbook.xml")
    except (KeyError, OSError, zipfile.BadZipFile):
        return None

    try:
        root = ET.fromstring(workbook_xml)
    except ET.ParseError:
        return None
    for node in root.iter():
        if node.tag.rsplit("}", 1)[-1] != "workbookView":
            continue
        raw = node.attrib.get("activeTab")
        if raw is None:
            return None
        try:
            return int(raw)
        except ValueError:
            return None
    return None


def _initialize_sheet_proxies(wb: Any, rust_book: Any) -> None:
    """Attach worksheet proxies from a Rust reader's tab list."""
    from wolfxl.chartsheet import Chartsheet

    names = [str(n) for n in rust_book.sheet_names()]
    chartsheet_names = set(_read_chartsheet_names(rust_book))
    wb._sheet_names = names
    wb._sheets = {
        name: Worksheet(wb, name) for name in names if name not in chartsheet_names
    }
    wb._chartsheets = {}
    for name in names:
        if name in chartsheet_names:
            cs = Chartsheet(wb, name)
            cs._source_chartsheet = True
            read_state = getattr(rust_book, "read_sheet_state", None)
            if read_state is not None:
                try:
                    cs.sheet_state = read_state(name)
                except Exception:
                    pass
            cs._charts = _read_chartsheet_charts(rust_book, name)
            wb._chartsheets[name] = cs
    wb._chartsheets_dirty = False


def _read_chartsheet_names(rust_book: Any) -> list[str]:
    reader = getattr(rust_book, "chartsheet_names", None)
    if reader is None:
        return []
    try:
        return [str(n) for n in reader()]
    except Exception:
        return []


def _read_chartsheet_charts(rust_book: Any, name: str) -> list[Any]:
    reader = getattr(rust_book, "read_chartsheet_charts", None)
    if reader is None:
        return []
    try:
        payloads = reader(name)
    except Exception:
        return []

    from wolfxl._worksheet_media import _chart_from_payload

    charts = []
    for payload in payloads:
        if isinstance(payload, dict):
            chart = _chart_from_payload(payload)
            if chart is not None:
                charts.append(chart)
    return charts


def initialize_pending_state(wb: Any) -> None:
    """Initialize workbook caches and pending mutation queues."""
    from wolfxl._styles import COLOR_INDEX, Alignment, Border, Color, Font
    from wolfxl.styles.cell_style import StyleArray
    from wolfxl.styles.differential import DifferentialStyleList
    from wolfxl.styles.fills import DEFAULT_EMPTY_FILL, DEFAULT_GRAY_FILL
    from wolfxl.styles.protection import Protection
    from wolfxl.styles.stylesheet import IndexedList
    from wolfxl.styles.table import TableStyleList

    if not hasattr(wb, "_chartsheets"):
        wb._chartsheets = {}
    wb._chartsheets_dirty = bool(getattr(wb, "_chartsheets_dirty", False))
    wb._properties_cache = None
    wb._custom_doc_props_cache = None
    wb._properties_dirty = False
    wb._custom_doc_props_dirty = False
    wb._defined_names_cache = None
    wb._named_styles_registry = None
    wb._style_names_cache = None
    wb._pending_defined_names = {}
    wb._security = None
    wb._file_sharing = None
    wb._security_loaded = False
    wb._pending_security_update = False
    wb._workbook_properties_cache = None
    wb._calc_properties_cache = None
    wb._calc_properties_original = None
    wb._views_cache = None
    wb._pending_axis_shifts = []
    wb._pending_range_moves = []
    wb._pending_sheet_copies = []
    wb._pending_chart_adds = {}
    wb._pending_source_chart_ops = []
    wb._active_sheet_index = None
    wb._active_sheet_dirty = False
    wb._pending_pivot_caches = []
    wb._next_pivot_cache_id = 0
    wb._pending_slicer_caches = []
    wb._next_slicer_cache_id = 0
    wb._fonts = IndexedList(
        [Font(name="Calibri", size=11, family=2, color=Color(theme=1), scheme="minor")]
    )
    wb._fills = IndexedList([DEFAULT_EMPTY_FILL, DEFAULT_GRAY_FILL])
    wb._borders = IndexedList([Border()])
    wb._alignments = IndexedList([Alignment()])
    wb._protections = IndexedList([Protection()])
    wb._number_formats = IndexedList()
    wb._cell_styles = IndexedList([StyleArray()])
    wb._differential_styles = DifferentialStyleList()
    wb._table_styles = TableStyleList()
    wb._colors = COLOR_INDEX[:64]
    wb._date_formats = set()
    wb._timedelta_formats = set()
    wb.copy_options = CopyOptions()
