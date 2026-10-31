"""Streaming read path — Sprint Ι Pod-β.

Public entry-point: :func:`stream_iter_rows`. Activated by
``load_workbook(path, read_only=True)`` or auto-trigger when a sheet has
more than ``AUTO_STREAM_ROW_THRESHOLD`` rows. Wraps the Rust
``StreamingSheetReader`` and converts its row tuples into either:

- ``StreamingCell`` instances (mutation-rejected proxies sharing converted
  workbook style-ID entries), or
- plain value tuples (when ``values_only=True``), padded by the configured
  column bounds.

The streaming path bypasses eager sheet materialization for the value scan but
still uses the eager workbook reader's style table for ``StreamingCell.font`` /
``.fill`` / etc. That table is loaded once on first style access and shared
across every cell in the iteration.
"""

from __future__ import annotations

from collections.abc import Iterator
from datetime import date, datetime, time, timedelta
import io
import os
import posixpath
import re
from typing import TYPE_CHECKING, Any
from xml.etree import ElementTree as ET
import zipfile

from wolfxl._cell_annotations import get_hyperlink
from wolfxl._utils import a1_to_rowcol
from wolfxl._utils import column_letter as _column_letter
from wolfxl._utils import rowcol_to_a1
from wolfxl._zip_safety import read_entry, validate_zipfile
from wolfxl.utils.datetime import from_excel
from wolfxl._streaming_styles import (
    _DEFAULT_STYLE,
    cell_border,
    cell_fill,
    cell_style_components,
    streaming_date_style_ids,
    workbook_style_cache,
)
from wolfxl.styles.fills import GradientFill

if TYPE_CHECKING:
    from wolfxl._styles import Alignment, Border, Font, PatternFill
    from wolfxl._worksheet import Worksheet
    from wolfxl._streaming_styles import StyleComponents


#: Row count above which ``iter_rows`` transparently uses the streaming
#: path even when the workbook was opened without ``read_only=True``.
#: Tuned to keep the eager path for small/medium workbooks (where the
#: bulk-read FFI is fastest) while still scaling to multi-million-cell
#: sheets without exhausting RSS.
AUTO_STREAM_ROW_THRESHOLD = 50_000

_DIMENSION_REF_RE = re.compile(rb"<(?:[A-Za-z0-9_]+:)?dimension\b[^>]*\bref=[\"']([^\"']+)[\"']")
_UNSET = object()


def stream_source(wb: Any) -> str | bytes | None:
    """Path, or in-memory bytes for file-like loads, the workbook came from."""
    data = getattr(wb, "_source_bytes", None)
    if data is not None:
        return data
    return getattr(wb, "_source_path", None)


def _open_source_zip(source: str | bytes) -> zipfile.ZipFile:
    if isinstance(source, (bytes, bytearray)):
        return zipfile.ZipFile(io.BytesIO(source))
    return zipfile.ZipFile(source)


def _stream_shared_strings(wb: Any, source: str | bytes) -> Any:
    """Shared-strings table for ``source``, parsed once per workbook.

    Native tables belong only to the workbook's admitted immutable byte source.
    Other sources must build their own table, even on the first streaming read.
    """
    from wolfxl import _rust

    cached = getattr(wb, "_stream_shared_strings", None)
    native = getattr(getattr(wb, "_rust_reader", None), "shared_strings", None)
    if isinstance(source, bytes):
        # Immutable bytes owned by the workbook: identity is the cache key.
        if cached is not None and cached[0] is source:
            return cached[1]
        key: Any = source
        table = (
            native()
            if cached is None and native is not None and source is getattr(wb, "_source_bytes", None)
            else _rust.StreamingSharedStrings.load_from_bytes(source)
        )
    elif isinstance(source, bytearray):
        return _rust.StreamingSharedStrings.load_from_bytes(bytes(source))
    else:
        st = os.stat(source)
        key = (source, st.st_ino, st.st_size, st.st_mtime_ns)
        if cached is not None and cached[0] == key:
            return cached[1]
        if cached is not None:
            limits = getattr(wb, "_limits", None)
            if limits is not None:
                limits._validate(source)
        table = _rust.StreamingSharedStrings.load(source)
    wb._stream_shared_strings = (key, table)
    return table


def _open_stream_reader(
    wb: Any, source: str | bytes, sheet_title: str, *bounds: int | None
) -> Any:
    from wolfxl import _rust

    shared_strings = _stream_shared_strings(wb, source)
    if isinstance(source, (bytes, bytearray)):
        return _rust.StreamingSheetReader.open_from_bytes(
            bytes(source), sheet_title, *bounds, shared_strings=shared_strings
        )
    return _rust.StreamingSheetReader.open(
        source, sheet_title, *bounds, shared_strings=shared_strings
    )


def _local_name(tag: str) -> str:
    if "}" in tag:
        return tag.rsplit("}", 1)[1]
    return tag.rsplit(":", 1)[-1]


def _resolve_package_target(base_dir: str, target: str) -> str:
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    return posixpath.normpath(posixpath.join(base_dir, target))


def _sheet_path_from_workbook(zf: zipfile.ZipFile, sheet_title: str) -> str | None:
    try:
        workbook_xml = read_entry(zf, "xl/workbook.xml")
        rels_xml = read_entry(zf, "xl/_rels/workbook.xml.rels")
    except (KeyError, OSError, zipfile.BadZipFile):
        return None

    rid_by_sheet: dict[str, str] = {}
    try:
        root = ET.fromstring(workbook_xml)
        for elem in root.iter():
            if _local_name(elem.tag) != "sheet":
                continue
            if elem.attrib.get("name") != sheet_title:
                continue
            for key, value in elem.attrib.items():
                if _local_name(key) == "id":
                    rid_by_sheet[sheet_title] = value
                    break
    except ET.ParseError:
        return None
    rid = rid_by_sheet.get(sheet_title)
    if rid is None:
        return None

    try:
        rels_root = ET.fromstring(rels_xml)
    except ET.ParseError:
        return None
    for rel in rels_root.iter():
        if _local_name(rel.tag) != "Relationship":
            continue
        if rel.attrib.get("Id") != rid:
            continue
        target = rel.attrib.get("Target")
        if not target:
            return None
        return _resolve_package_target("xl", target)
    return None


def _dimension_ref_from_sheet_head(zf: zipfile.ZipFile, sheet_path: str) -> str | None:
    try:
        with zf.open(sheet_path) as fh:
            head = bytearray()
            while len(head) < 131_072:
                chunk = fh.read(8192)
                if not chunk:
                    break
                head.extend(chunk)
                match = _DIMENSION_REF_RE.search(head)
                if match is not None:
                    return match.group(1).decode("ascii", errors="ignore")
                if b"<sheetData" in head or b"<sheetData>" in head:
                    break
    except (KeyError, OSError, zipfile.BadZipFile):
        return None
    return None


def _max_row_from_dimension_ref(ref: str) -> int | None:
    cell = ref.split(":", 1)[-1].replace("$", "")
    try:
        return a1_to_rowcol(cell)[0]
    except ValueError:
        return None


def _bounds_from_dimension_ref(ref: str) -> tuple[int, int, int, int] | None:
    parts = ref.replace("$", "").split(":", 1)
    if len(parts) == 1:
        parts = [parts[0], parts[0]]
    try:
        min_row, min_col = a1_to_rowcol(parts[0])
        max_row, max_col = a1_to_rowcol(parts[1])
    except ValueError:
        return None
    return (
        min(min_row, max_row),
        min(min_col, max_col),
        max(min_row, max_row),
        max(min_col, max_col),
    )


def _source_dimension_bounds(
    source: str | bytes, sheet_title: str
) -> tuple[int, int, int, int] | None:
    try:
        with _open_source_zip(source) as zf:
            validate_zipfile(zf)
            sheet_path = _sheet_path_from_workbook(zf, sheet_title)
            if sheet_path is None:
                return None
            ref = _dimension_ref_from_sheet_head(zf, sheet_path)
            return _bounds_from_dimension_ref(ref) if ref is not None else None
    except (OSError, ValueError, zipfile.BadZipFile):
        return None


def _source_cell_bounds(source: str | bytes, sheet_title: str) -> tuple[int, int, int, int]:
    """Eager cell/merge bounds, scanned natively without decoding values."""
    from wolfxl import _rust

    if isinstance(source, (bytes, bytearray)):
        max_row, max_col = _rust.StreamingSheetReader.cell_bounds_from_bytes(
            bytes(source), sheet_title
        )
    else:
        max_row, max_col = _rust.StreamingSheetReader.cell_bounds(source, sheet_title)
    return 1, 1, max_row, max_col


def _source_dimension_max_row(source: str | bytes, sheet_title: str) -> int | None:
    try:
        with _open_source_zip(source) as zf:
            validate_zipfile(zf)
            sheet_path = _sheet_path_from_workbook(zf, sheet_title)
            if sheet_path is None:
                return None
            ref = _dimension_ref_from_sheet_head(zf, sheet_path)
        bounds = _bounds_from_dimension_ref(ref) if ref is not None else None
        if bounds is not None:
            return bounds[2]
        if _row_tags_exceed(source, sheet_title, AUTO_STREAM_ROW_THRESHOLD):
            return AUTO_STREAM_ROW_THRESHOLD + 1
        return None
    except (OSError, ValueError, zipfile.BadZipFile):
        return None


def _row_tags_exceed(source: str | bytes, sheet_title: str, limit: int) -> bool:
    """Whether the sheet XML holds more than ``limit`` ``<row>`` tags, counted natively."""
    from wolfxl import _rust

    if isinstance(source, (bytes, bytearray)):
        return _rust.StreamingSheetReader.row_tags_exceed_from_bytes(
            bytes(source), sheet_title, limit
        )
    return _rust.StreamingSheetReader.row_tags_exceed(source, sheet_title, limit)


def _streaming_value(payload: Any, *, data_only: bool = False) -> Any:
    """Normalize a Rust streaming payload into the same Python types
    that the eager value-reader emits.

    The Rust streaming layer surfaces:
      - plain ``str``/``int``/``float``/``bool`` for the common types,
      - a ``dict`` with ``type=formula|error`` for formulas/errors,
      - the raw ISO string for ``t="d"`` cells (Excel rarely emits these).
    """
    if isinstance(payload, dict):
        t = payload.get("type")
        if t == "formula":
            if data_only:
                return _cached_formula_value(payload.get("cached"))
            # openpyxl returns the formula string from `Cell.value` in
            # read mode, so do the same here.
            return payload.get("formula")
        if t == "error":
            return payload.get("value")
    return payload


def _cached_formula_value(value: Any) -> Any:
    """Cached formula results arrive typed from Rust; an empty cache is None."""
    if value == "":
        return None
    return value


def _timedelta_from_serial(value: Any) -> Any:
    """Keep openpyxl's two-stage timedelta rounding and numeric overflow fallback."""
    try:
        return from_excel(value, timedelta=True)
    except (OverflowError, ValueError):
        return value


def _configure_value_reader(reader: Any, wb: Any) -> None:
    """Configure values-only rows once: formula output and date-style conversion."""
    reader.set_data_only(bool(getattr(wb, "_data_only", False)))
    date_styles, timedelta_styles = streaming_date_style_ids(wb)
    if not date_styles:
        return
    from wolfxl.utils.datetime import CALENDAR_MAC_1904

    kinds = bytearray(len(getattr(wb, "_cell_styles", ())))
    for style_id in date_styles:
        if 0 <= style_id < len(kinds):
            kinds[style_id] = 2 if style_id in timedelta_styles else 1
    reader.set_date_styles(bytes(kinds), _timedelta_from_serial, wb.epoch == CALENDAR_MAC_1904)


class StreamingCell:
    """Read-only cell proxy yielded by streaming ``iter_rows()``.

    Holds a snapshot of the value (parsed by the Rust SAX scanner) plus
    the originating row/column and a reference back to the Worksheet so
    style lookups can reuse the workbook stylesheet by source style ID.
    Style attributes are
    fully featured (font, fill, border, alignment, number_format) and
    behave identically to the eager ``Cell`` properties — the difference
    is mutation: every setter raises ``RuntimeError``.

    The class is ``__slots__``-based so a 1M-cell scan doesn't blow the
    Python heap with per-cell ``__dict__`` storage.
    """

    __slots__ = (
        "_ws",
        "_row",
        "_col",
        "_value",
        "_style_id",
        "_cell_type",
        "_style_cache",
        "_style_components",
    )

    def __init__(
        self,
        ws: Worksheet,
        row: int,
        col: int,
        value: Any,
        style_id: int | None,
        cell_type: str,
    ) -> None:
        object.__setattr__(self, "_ws", ws)
        object.__setattr__(self, "_row", row)
        object.__setattr__(self, "_col", col)
        data_only = bool(getattr(ws._workbook, "_data_only", False))  # noqa: SLF001
        if data_only and cell_type == "formula" and isinstance(value, dict):
            # Like openpyxl, an empty cache keeps the cell's `t=` type, and an
            # error cache stays an error.
            cached_type = value.get("cached_type")
            if cached_type is not None and (
                cached_type == "e" or _cached_formula_value(value.get("cached")) is None
            ):
                cell_type = cached_type
        object.__setattr__(self, "_value", _streaming_value(value, data_only=data_only))
        object.__setattr__(self, "_style_id", style_id)
        object.__setattr__(self, "_cell_type", cell_type)
        object.__setattr__(self, "_style_cache", None)
        object.__setattr__(self, "_style_components", None)

    # ------------------------------------------------------------------
    # Coordinate accessors — match openpyxl's read_only Cell API.
    # ------------------------------------------------------------------

    @property
    def value(self) -> Any:
        """Return the streaming cell value, converting date serials when needed."""
        # Sprint Λ Pod-γ: numeric cells whose number format is date-typed
        # surface as Python datetime, mirroring openpyxl's read_only path
        # (and the eager Cell.value path, which converts via calamine).
        if isinstance(self._value, (int, float)) and not isinstance(self._value, bool):
            style = self._resolved_style()
            if style.is_date:
                try:
                    return from_excel(
                        self._value,
                        epoch=self._ws._workbook.epoch,
                        timedelta=style.is_timedelta,
                    )
                except Exception:
                    return self._value
        return self._value

    @property
    def row(self) -> int:
        """Return this cell's 1-based row index."""
        return self._row

    @property
    def column(self) -> int:
        """Return this cell's 1-based column index."""
        return self._col

    @property
    def column_letter(self) -> str:
        """Return this cell's Excel column letter."""
        return _column_letter(self._col)

    @property
    def coordinate(self) -> str:
        """Return this cell's A1-style coordinate."""
        return rowcol_to_a1(self._row, self._col)

    @property
    def parent(self) -> Worksheet:
        """Return the containing worksheet."""
        return self._ws

    @property
    def data_type(self) -> str:
        """Map the SAX ``t=`` token onto openpyxl's ``data_type`` letters."""
        if self._cell_type == "formula" and self._ws._workbook._data_only:
            value = self.value
            if isinstance(value, bool):
                return "b"
            if isinstance(value, str):
                return "s"
            if isinstance(value, (date, datetime, time, timedelta)):
                return "d"
            return "n"
        if self._cell_type == "n" and isinstance(self.value, (date, datetime, time, timedelta)):
            return "d"
        return {
            "n": "n",
            "s": "s",
            "str": "str" if self._value is None else "s",
            "inlineStr": "inlineStr" if self._value is None else "s",
            "b": "b",
            "e": "e",
            "d": "d",
            "formula": "f",
            "blank": "n",
        }.get(self._cell_type, "n")

    @property
    def hyperlink(self) -> Any:
        """Return the worksheet hyperlink, including pending unsaved edits."""
        return get_hyperlink(self, _UNSET)

    # ------------------------------------------------------------------
    # Style lookups — share converted stylesheet entries by native style ID.
    # The streaming path does NOT re-implement xl/styles.xml parsing;
    # the styles table is small (~KB) regardless of sheet size and the
    # eager reader caches it after first access.
    # ------------------------------------------------------------------

    def _resolved_style(self) -> StyleComponents:
        workbook = self._ws._workbook
        cache = getattr(workbook, "_streaming_style_cache", None)
        if (
            self._style_components is None
            or cache is not self._style_cache
            or (cache is not None and cache.reader_identity != id(workbook._rust_reader))
        ):
            style = cell_style_components(self)
            object.__setattr__(self, "_style_components", style)
            object.__setattr__(
                self, "_style_cache", getattr(workbook, "_streaming_style_cache", None)
            )
        return self._style_components

    @property
    def font(self) -> Font:
        """Return the resolved cell font."""
        return self._resolved_style().font

    @property
    def fill(self) -> PatternFill:
        """Return the resolved cell fill."""
        return cell_fill(self)

    @property
    def border(self) -> Border:
        """Return the resolved cell border."""
        return cell_border(self)

    @property
    def alignment(self) -> Alignment:
        """Return the resolved cell alignment."""
        return self._resolved_style().alignment

    @property
    def number_format(self) -> str | None:
        """Return the resolved number format string."""
        if self._ws._workbook._rust_reader is None:  # noqa: SLF001
            return None
        return self._resolved_style().number_format

    # ------------------------------------------------------------------
    # Mutation — strictly rejected. Sprint Ι Pod-β contract.
    #
    # We don't bother defining `@value.setter` etc.; ``__setattr__``
    # below is the single chokepoint that rejects every assignment and
    # carries a uniform error message. The trade-off: a user who
    # registers a property setter via inheritance can't override the
    # rejection — by design, since StreamingCell is final-by-contract.
    # ------------------------------------------------------------------

    def _reject(self, what: str) -> None:
        raise RuntimeError(
            f"read_only=True: cannot set {what} on streaming cell "
            f"{self.coordinate}; reload without read_only or use modify=True"
        )

    def __setattr__(self, name: str, value: Any) -> None:
        # __slots__ permits assignment to declared slot names; reject
        # any others up-front so a typo'd attribute doesn't silently
        # succeed via setter resolution. The constructor is the only
        # legitimate caller for slot writes.
        if name in StreamingCell.__slots__:
            object.__setattr__(self, name, value)
            return
        self._reject(name)

    def __repr__(self) -> str:
        """Return a compact debug representation for this streaming cell."""
        return f"<StreamingCell {self.coordinate} value={self._value!r}>"

    def __eq__(self, other: object) -> bool:
        return (
            getattr(other, "coordinate", None) == self.coordinate
            and getattr(other, "value", None) == self.value
        )


class StreamingBlankCell(StreamingCell):
    """Coordinate-aware filler for an absent XML cell with a pattern default.

    Source style zero resolves font/alignment/number format from an empty
    payload. Pattern fills have the same empty default. Share those immutable
    components without initializing value and per-cell style-cache slots.
    Border access still resolves workbook/merge state through the normal path.
    A gradient at source style zero uses the full proxy instead.
    """

    __slots__ = ()
    _value = None
    _style_id = None
    _cell_type = "blank"
    value = None
    data_type = "n"
    font = _DEFAULT_STYLE.font
    fill = _DEFAULT_STYLE.fill
    alignment = _DEFAULT_STYLE.alignment

    def __init__(self, ws: Worksheet, row: int, col: int) -> None:
        object.__setattr__(self, "_ws", ws)
        object.__setattr__(self, "_row", row)
        object.__setattr__(self, "_col", col)

    @property
    def number_format(self) -> str | None:
        return "General" if self._ws._workbook._rust_reader is not None else None

    def _resolved_style(self) -> StyleComponents:
        return cell_style_components(self)

    def __copy__(self) -> StreamingBlankCell:
        # The inherited slot state also includes class-level blank constants.
        # Construct a fresh coordinate snapshot instead of assigning those.
        return type(self)(self._ws, self._row, self._col)


def _full_blank_cell(ws: Worksheet, row: int, col: int) -> StreamingCell:
    return StreamingCell(ws, row, col, None, None, "blank")


def _blank_rows(ws, blank_cell, first_row, stop_row, min_col, max_col):
    """Build bounded batches while keeping every cell's coordinate snapshot."""
    from wolfxl import _rust

    width = 0 if max_col is None else max(0, max_col - min_col + 1)
    native = getattr(_rust, "streaming_blank_rows", None)
    if blank_cell is StreamingBlankCell and native is not None and 0 < width <= 4096:
        batch_size = min(128, 4096 // width)
        while first_row < stop_row:
            count = min(batch_size, stop_row - first_row)
            yield from native(blank_cell, ws, first_row, count, min_col, max_col)
            first_row += count
        return
    for row in range(first_row, stop_row):
        yield (
            tuple(blank_cell(ws, row, col) for col in range(min_col, max_col + 1))
            if max_col is not None
            else []
        )


def _resolve_bounds(
    ws: Worksheet,
    min_row: int | None,
    max_row: int | None,
    min_col: int | None,
    max_col: int | None,
) -> tuple[int | None, int | None, int | None, int | None]:
    """Default ``min_row`` to 1 as openpyxl does, so an invalid
    ``<row r="0">`` is skipped; other ``None`` bounds stay unbounded for the
    Rust reader. Kept as a seam so future sheet-dimension probing can land
    here without churning callers.
    """
    return min_row or 1, max_row, min_col, max_col


def stream_iter_rows(
    ws: Worksheet,
    min_row: int | None = None,
    max_row: int | None = None,
    min_col: int | None = None,
    max_col: int | None = None,
    values_only: bool = False,
) -> Iterator[tuple[Any, ...]]:
    """SAX-streaming generator backing ``iter_rows`` in read-only / large-sheet mode.

    Yields:
        - When ``values_only=True``: plain tuples of cell values padded to
          ``(min_col..max_col)`` width. Missing cells become ``None``.
          When no column bound is supplied, the tuple is sized by the
          actual cells present in that row.
        - Otherwise: tuples of :class:`StreamingCell` instances covering
          the same range.
    """
    wb = ws._workbook  # noqa: SLF001
    path = stream_source(wb)
    if path is None:
        raise RuntimeError(
            "stream_iter_rows requires a workbook loaded from a path or file-like "
            "source (use load_workbook to obtain one)"
        )
    mn_r, mx_r, mn_c, mx_c = _resolve_bounds(ws, min_row, max_row, min_col, max_col)
    source_bounds = (
        None
        if bool(getattr(ws, "_read_only_dimensions_reset", False))
        else _source_dimension_bounds(path, ws.title)
    )
    forced_bounds = None
    if bool(getattr(ws, "_read_only_dimensions_forced", False)):
        dimensions = getattr(ws, "_dimensions", None)
        if dimensions is not None:
            forced_bounds = (1, 1, int(dimensions[0]), int(dimensions[1]))
    default_bounds = source_bounds or forced_bounds
    # Eager openpyxl ignores stale dimensions and row-only formatting.
    # Keep the read-only contract, including declared trailing blank rows.
    if not bool(getattr(wb, "_read_only", False)) and (mx_r is None or mx_c is None):
        default_bounds = _source_cell_bounds(path, ws.title)
    stream_mx_r = (
        mx_r if mx_r is not None else (default_bounds[2] if default_bounds is not None else None)
    )
    stream_mx_c = (
        mx_c if mx_c is not None else (default_bounds[3] if default_bounds is not None else None)
    )
    reader = _open_stream_reader(wb, path, ws.title, mn_r, stream_mx_r, mn_c, stream_mx_c)

    def _fixed_cmax() -> int | None:
        if mx_c is not None:
            return mx_c
        if default_bounds is not None:
            return default_bounds[3]
        return None

    def _resolved_cmax(cells: list[tuple[Any, ...]]) -> int:
        fixed = _fixed_cmax()
        if fixed is not None:
            return fixed
        return max((int(cell[0]) for cell in cells), default=(mn_c or 1) - 1)

    try:
        if values_only:
            _configure_value_reader(reader, wb)
            counter = mn_r if mn_r is not None else 1
            cmin = mn_c if mn_c is not None else 1
            fixed_cmax = _fixed_cmax()
            empty_row = (None,) * max(0, fixed_cmax + 1 - cmin) if fixed_cmax is not None else []
            while True:
                chunk_result = reader.read_next_values_compact_chunk(1024)
                if chunk_result is None:
                    break
                chunk, first_row, _last_row, dense = chunk_result
                if dense:
                    while counter < first_row:
                        yield empty_row
                        counter += 1
                    for offset, values in enumerate(chunk):
                        row_idx = first_row + offset
                        while counter < row_idx:
                            yield empty_row
                            counter += 1
                        yield values
                        counter = row_idx + 1
                    continue
                for row_idx, values in chunk:
                    while counter < row_idx:
                        yield empty_row
                        counter += 1
                    yield values
                    counter = row_idx + 1
            if stream_mx_r is not None and _pad_trailing_rows(wb, reader, mx_r):
                while counter <= stream_mx_r:
                    yield empty_row
                    counter += 1
        else:
            default_style = workbook_style_cache(ws._workbook).components(ws._workbook, None)
            blank_cell = (
                _full_blank_cell
                if isinstance(default_style.fill, GradientFill)
                else StreamingBlankCell
            )
            counter = mn_r if mn_r is not None else 1
            while True:
                row = reader.read_next_row()
                if row is None:
                    break
                row_idx, cells = row
                # Determine emitted column bounds: explicit min/max wins,
                # else span observed cells.
                cmin = mn_c if mn_c is not None else 1
                cmax = _resolved_cmax(cells)
                if counter < row_idx:
                    yield from _blank_rows(ws, blank_cell, counter, row_idx, cmin, _fixed_cmax())
                    counter = row_idx
                if cmax < cmin:
                    yield ()
                    counter = row_idx + 1
                    continue
                by_col = {c[0]: c for c in cells}
                row_out: list[Any] = []
                for col in range(cmin, cmax + 1):
                    rec = by_col.get(col)
                    if rec is None:
                        row_out.append(blank_cell(ws, row_idx, col))
                    else:
                        _, value, style_id, cell_type = rec
                        row_out.append(StreamingCell(ws, row_idx, col, value, style_id, cell_type))
                yield tuple(row_out)
                counter = row_idx + 1
            if stream_mx_r is not None and _pad_trailing_rows(wb, reader, mx_r):
                cmin = mn_c if mn_c is not None else 1
                fixed_cmax = _fixed_cmax()
                yield from _blank_rows(ws, blank_cell, counter, stream_mx_r + 1, cmin, fixed_cmax)
    finally:
        reader.close()


def _pad_trailing_rows(wb: Any, reader: Any, explicit_max_row: int | None) -> bool:
    """Whether to pad blank rows from the last streamed row up to ``max_row``.

    An explicit ``max_row`` always pads, as does the auto-streamed eager path
    (eager openpyxl materializes every cell in the range). When ``max_row``
    comes from the declared ``<dimension>``, follow openpyxl's read-only
    iterator: pad only if the sheet has a ``<row>`` past it. A sheet whose
    rows end early, including one with no rows at all, yields only its rows.
    """
    if explicit_max_row is not None or not bool(getattr(wb, "_read_only", False)):
        return True
    return bool(reader.stopped_past_max_row())


def stream_value_chunks(
    ws: Worksheet,
    min_row: int | None = None,
    max_row: int | None = None,
    min_col: int | None = None,
    max_col: int | None = None,
    *,
    chunk_size: int = 1024,
) -> Iterator[tuple[tuple[Any, ...], ...]]:
    """Stream value rows in chunks for high-throughput table scans."""
    if chunk_size < 1:
        raise ValueError("chunk_size must be >= 1")
    wb = ws._workbook  # noqa: SLF001

    path = stream_source(wb)
    if path is None:
        raise RuntimeError(
            "stream_value_chunks requires a workbook loaded from a path or file-like "
            "source (use load_workbook to obtain one)"
        )
    mn_r, mx_r, mn_c, mx_c = _resolve_bounds(ws, min_row, max_row, min_col, max_col)
    source_bounds = (
        None
        if bool(getattr(ws, "_read_only_dimensions_reset", False))
        else _source_dimension_bounds(path, ws.title)
    )
    forced_bounds = None
    if bool(getattr(ws, "_read_only_dimensions_forced", False)):
        dimensions = getattr(ws, "_dimensions", None)
        if dimensions is not None:
            forced_bounds = (1, 1, int(dimensions[0]), int(dimensions[1]))
    default_bounds = source_bounds or forced_bounds
    stream_mx_r = (
        mx_r if mx_r is not None else (default_bounds[2] if default_bounds is not None else None)
    )
    stream_mx_c = (
        mx_c if mx_c is not None else (default_bounds[3] if default_bounds is not None else None)
    )
    reader = _open_stream_reader(wb, path, ws.title, mn_r, stream_mx_r, mn_c, stream_mx_c)
    try:
        _configure_value_reader(reader, wb)
        counter = mn_r if mn_r is not None else 1
        cmin = mn_c if mn_c is not None else 1
        fixed_cmax = (
            mx_c
            if mx_c is not None
            else (default_bounds[3] if default_bounds is not None else None)
        )
        # Without a column bound, openpyxl yields ``[]`` for a missing row
        # (and ``stream_iter_rows`` mirrors it); keep chunks equal to ``ws.values``.
        empty_row = (None,) * max(0, fixed_cmax + 1 - cmin) if fixed_cmax is not None else []
        pending: list[tuple[Any, ...]] = []
        while True:
            chunk_result = reader.read_next_values_compact_chunk(chunk_size)
            if chunk_result is None:
                break
            chunk, first_row, last_row, dense = chunk_result
            if (
                dense
                and not pending
                and chunk
                and first_row == counter
                and last_row == counter + len(chunk) - 1
            ):
                yield tuple(chunk)
                counter = last_row + 1
                continue
            if dense:
                chunk = tuple((first_row + offset, values) for offset, values in enumerate(chunk))
            for row_idx, values in chunk:
                while counter < row_idx:
                    pending.append(empty_row)
                    if len(pending) >= chunk_size:
                        yield tuple(pending)
                        pending = []
                    counter += 1
                pending.append(values)
                if len(pending) >= chunk_size:
                    yield tuple(pending)
                    pending = []
                counter = row_idx + 1
        if stream_mx_r is not None and _pad_trailing_rows(wb, reader, mx_r):
            while counter <= stream_mx_r:
                pending.append(empty_row)
                if len(pending) >= chunk_size:
                    yield tuple(pending)
                    pending = []
                counter += 1
        if pending:
            yield tuple(pending)
    finally:
        reader.close()


def should_auto_stream(ws: Worksheet) -> bool:
    """Heuristic: auto-engage streaming for ``iter_rows`` on huge sheets.

    Triggers when the sheet's ``max_row`` exceeds
    :data:`AUTO_STREAM_ROW_THRESHOLD`. This deliberately consults the
    worksheet ``<dimension ref=...>`` tag directly from the ZIP head instead
    of ``Worksheet._bounds_max_row()``, because the normal dimensions API may parse
    and cache the full sheet before streaming can engage.
    """
    wb = ws._workbook  # noqa: SLF001
    if not getattr(wb, "_auto_stream", True):
        return False
    path = stream_source(wb)
    if not path:
        return False
    rows = _source_dimension_max_row(path, ws.title)
    return rows is not None and rows > AUTO_STREAM_ROW_THRESHOLD
