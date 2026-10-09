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
from datetime import date, datetime, time
import posixpath
import re
from typing import TYPE_CHECKING, Any
from xml.etree import ElementTree as ET
import zipfile

from wolfxl._utils import a1_to_rowcol
from wolfxl._utils import column_letter as _column_letter
from wolfxl._utils import rowcol_to_a1
from wolfxl._zip_safety import read_entry, validate_zipfile
from wolfxl.utils.datetime import from_excel
from wolfxl.styles.numbers import is_timedelta_format
from wolfxl.utils.numbers import is_date_format
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
_ROW_TAG_RE = re.compile(rb"<(?:[A-Za-z0-9_]+:)?row(?:\s|>)")


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


def _row_count_exceeds_threshold(zf: zipfile.ZipFile, sheet_path: str) -> bool:
    try:
        with zf.open(sheet_path) as fh:
            tail = b""
            count = 0
            while True:
                chunk = fh.read(65_536)
                if not chunk:
                    break
                buf = tail + chunk
                count += len(_ROW_TAG_RE.findall(buf))
                if count > AUTO_STREAM_ROW_THRESHOLD:
                    return True
                tail = buf[-32:]
    except (KeyError, OSError, zipfile.BadZipFile):
        return False
    return False


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


def _source_dimension_bounds(path: str, sheet_title: str) -> tuple[int, int, int, int] | None:
    try:
        with zipfile.ZipFile(path) as zf:
            validate_zipfile(zf)
            sheet_path = _sheet_path_from_workbook(zf, sheet_title)
            if sheet_path is None:
                return None
            ref = _dimension_ref_from_sheet_head(zf, sheet_path)
            return _bounds_from_dimension_ref(ref) if ref is not None else None
    except (OSError, zipfile.BadZipFile):
        return None


def _source_dimension_max_row(path: str, sheet_title: str) -> int | None:
    try:
        with zipfile.ZipFile(path) as zf:
            validate_zipfile(zf)
            sheet_path = _sheet_path_from_workbook(zf, sheet_title)
            if sheet_path is None:
                return None
            ref = _dimension_ref_from_sheet_head(zf, sheet_path)
            bounds = _bounds_from_dimension_ref(ref) if ref is not None else None
            max_row = bounds[2] if bounds is not None else None
            if max_row is not None:
                return max_row
            if _row_count_exceeds_threshold(zf, sheet_path):
                return AUTO_STREAM_ROW_THRESHOLD + 1
            return None
    except (OSError, zipfile.BadZipFile):
        return None


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
    if value in (None, ""):
        return None
    if not isinstance(value, str):
        return value
    try:
        return int(value)
    except ValueError:
        pass
    try:
        return float(value)
    except ValueError:
        return value


def _maybe_datetime_from_serial(value: Any, num_fmt: str | None) -> Any:
    """Convert an Excel serial to a ``datetime``/``date``/``time`` when
    ``num_fmt`` is a date-typed format string (Sprint Λ Pod-γ).

    Mirrors openpyxl's read-only path: numeric cells with a date number
    format surface as Python datetimes rather than raw serials. Non-date
    formats and non-numeric values pass through unchanged.
    """
    if not isinstance(value, (int, float)) or isinstance(value, bool):
        return value
    if not is_date_format(num_fmt):
        return value
    try:
        return from_excel(value, timedelta=is_timedelta_format(num_fmt))
    except Exception:
        # Out-of-range serial / unrepresentable timedelta — fall back to
        # the raw serial so the row keeps flowing rather than erroring
        # mid-iteration. Matches openpyxl's behavior on bad serials.
        return value


def _values_only_needs_style_dates(wb: Any) -> bool:
    """Return whether streaming value rows need style ids for date conversion."""
    return bool(getattr(wb, "_date_formats", ()))


def _normalise_streaming_value_tuple(row: tuple[Any, ...], *, data_only: bool) -> tuple[Any, ...]:
    """Convert formula/error payload dictionaries only when a row contains them."""
    for value in row:
        if isinstance(value, dict):
            return tuple(
                _streaming_value(item, data_only=data_only)
                if isinstance(item, dict)
                else item
                for item in row
            )
    return row


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
        "_ws", "_row", "_col", "_value", "_style_id", "_cell_type",
        "_style_cache", "_style_components",
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
        if self._cell_type == "n" and isinstance(self.value, (date, datetime, time)):
            return "d"
        return {
            "n": "n",
            "s": "s",
            "str": "s",
            "inlineStr": "s",
            "b": "b",
            "e": "e",
            "d": "d",
            "formula": "f",
            "blank": "n",
        }.get(self._cell_type, "n")

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
            object.__setattr__(self, "_style_cache", getattr(workbook, "_streaming_style_cache", None))
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


def _full_blank_cell(ws: Worksheet, row: int, col: int) -> StreamingCell:
    return StreamingCell(ws, row, col, None, None, "blank")


def _resolve_bounds(
    ws: Worksheet,
    min_row: int | None,
    max_row: int | None,
    min_col: int | None,
    max_col: int | None,
) -> tuple[int | None, int | None, int | None, int | None]:
    """Pass-through that just clamps user ``None``s to ``None`` (the Rust
    layer treats ``None`` as 'unbounded'). Kept as a seam so future
    sheet-dimension probing can land here without churning callers.
    """
    return min_row, max_row, min_col, max_col


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
    from wolfxl import _rust

    wb = ws._workbook  # noqa: SLF001
    path = getattr(wb, "_source_path", None)
    if path is None:
        raise RuntimeError(
            "stream_iter_rows requires a source-path-bearing workbook "
            "(use load_workbook to obtain one)"
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
    stream_mx_r = mx_r if mx_r is not None else (
        default_bounds[2] if default_bounds is not None else None
    )
    stream_mx_c = mx_c if mx_c is not None else (
        default_bounds[3] if default_bounds is not None else None
    )
    reader = _rust.StreamingSheetReader.open(
        path, ws.title, mn_r, stream_mx_r, mn_c, stream_mx_c
    )

    # The hydrated stylesheet has already classified source style IDs.
    date_styles, timedelta_styles = streaming_date_style_ids(wb)

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
            if not _values_only_needs_style_dates(wb) and hasattr(reader, "read_next_values_row"):
                data_only = bool(getattr(wb, "_data_only", False))
                counter = mn_r if mn_r is not None else 1
                cmin = mn_c if mn_c is not None else 1
                fixed_cmax = _fixed_cmax()
                empty_row = (
                    (None,) * max(0, fixed_cmax + 1 - cmin)
                    if fixed_cmax is not None
                    else []
                )
                read_chunk_with_flags = getattr(
                    reader, "read_next_values_chunk_with_flags", None
                )
                read_compact_chunk_with_flags = getattr(
                    reader, "read_next_values_compact_chunk_with_flags", None
                )
                read_chunk = getattr(reader, "read_next_values_chunk", None)
                if (
                    read_compact_chunk_with_flags is not None
                    or read_chunk_with_flags is not None
                ):
                    while True:
                        if read_compact_chunk_with_flags is not None:
                            chunk_result = read_compact_chunk_with_flags(1024)
                        else:
                            chunk_result = read_chunk_with_flags()
                        if chunk_result is None:
                            break
                        if len(chunk_result) == 5:
                            (
                                chunk,
                                needs_normalise,
                                first_row,
                                _last_row,
                                dense_plain,
                            ) = chunk_result
                            if dense_plain:
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
                        else:
                            chunk, needs_normalise = chunk_result
                        for row_idx, values in chunk:
                            while counter < row_idx:
                                yield empty_row
                                counter += 1
                            if needs_normalise:
                                values = _normalise_streaming_value_tuple(
                                    values, data_only=data_only
                                )
                            yield values
                            counter = row_idx + 1
                elif read_chunk is not None:
                    while True:
                        chunk = read_chunk()
                        if chunk is None:
                            break
                        for row_idx, values in chunk:
                            while counter < row_idx:
                                yield empty_row
                                counter += 1
                            yield _normalise_streaming_value_tuple(
                                values, data_only=data_only
                            )
                            counter = row_idx + 1
                else:
                    while True:
                        row = reader.read_next_values_row()
                        if row is None:
                            break
                        row_idx, values = row
                        while counter < row_idx:
                            yield empty_row
                            counter += 1
                        yield _normalise_streaming_value_tuple(values, data_only=data_only)
                        counter = row_idx + 1
                if stream_mx_r is not None:
                    while counter <= stream_mx_r:
                        yield empty_row
                        counter += 1
                return

            counter = mn_r if mn_r is not None else 1
            # Switch to ``read_next_row`` rather than ``read_next_values``
            # so we have access to each cell's style_id. The slight extra
            # work (vs Rust-side tuple padding) is offset by skipping a
            # second materialization of the tuple in Python.
            while True:
                row = reader.read_next_row()
                if row is None:
                    break
                row_idx, cells = row
                cmin = mn_c if mn_c is not None else 1
                cmax = _resolved_cmax(cells)
                fixed_cmax = _fixed_cmax()
                empty_row = (
                    (None,) * max(0, fixed_cmax + 1 - cmin)
                    if fixed_cmax is not None
                    else []
                )
                while counter < row_idx:
                    yield empty_row
                    counter += 1
                if cmax < cmin:
                    yield ()
                    counter = row_idx + 1
                    continue
                row_out: list[Any] = []
                cell_idx = 0
                cell_count = len(cells)
                for col in range(cmin, cmax + 1):
                    while cell_idx < cell_count and int(cells[cell_idx][0]) < col:
                        cell_idx += 1
                    rec = (
                        cells[cell_idx]
                        if cell_idx < cell_count and int(cells[cell_idx][0]) == col
                        else None
                    )
                    if rec is None:
                        row_out.append(None)
                        continue
                    cell_idx += 1
                    _, value, style_id, _cell_type = rec
                    py_val = (
                        _streaming_value(
                            value,
                            data_only=bool(getattr(wb, "_data_only", False)),
                        )
                        if isinstance(value, dict)
                        else value
                    )
                    if (
                        isinstance(py_val, (int, float))
                        and not isinstance(py_val, bool)
                        and style_id in date_styles
                    ):
                        try:
                            py_val = from_excel(
                                py_val, epoch=wb.epoch, timedelta=style_id in timedelta_styles
                            )
                        except Exception:
                            pass
                    row_out.append(py_val)
                yield tuple(row_out)
                counter = row_idx + 1
            if stream_mx_r is not None:
                cmin = mn_c if mn_c is not None else 1
                fixed_cmax = _fixed_cmax()
                empty_row = (
                    (None,) * max(0, fixed_cmax + 1 - cmin)
                    if fixed_cmax is not None
                    else []
                )
                while counter <= stream_mx_r:
                    yield empty_row
                    counter += 1
        else:
            default_style = workbook_style_cache(ws._workbook).components(ws._workbook, None)
            blank_cell = (
                _full_blank_cell if isinstance(default_style.fill, GradientFill)
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
                while counter < row_idx:
                    fixed_cmax = _fixed_cmax()
                    empty_row = (
                        tuple(
                            blank_cell(ws, counter, col)
                            for col in range(cmin, fixed_cmax + 1)
                        )
                        if fixed_cmax is not None
                        else []
                    )
                    yield empty_row
                    counter += 1
                if cmax < cmin:
                    yield ()
                    counter = row_idx + 1
                    continue
                by_col = {c[0]: c for c in cells}
                row_out: list[Any] = []
                for col in range(cmin, cmax + 1):
                    rec = by_col.get(col)
                    if rec is None:
                        row_out.append(
                            blank_cell(ws, row_idx, col)
                        )
                    else:
                        _, value, style_id, cell_type = rec
                        row_out.append(
                            StreamingCell(ws, row_idx, col, value, style_id, cell_type)
                        )
                yield tuple(row_out)
                counter = row_idx + 1
            if stream_mx_r is not None:
                cmin = mn_c if mn_c is not None else 1
                fixed_cmax = _fixed_cmax()
                while counter <= stream_mx_r:
                    yield (
                        tuple(
                            blank_cell(ws, counter, col)
                            for col in range(cmin, fixed_cmax + 1)
                        )
                        if fixed_cmax is not None
                        else []
                    )
                    counter += 1
    finally:
        reader.close()


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
    if _values_only_needs_style_dates(wb):
        yield from _chunk_value_rows(
            stream_iter_rows(
                ws,
                min_row=min_row,
                max_row=max_row,
                min_col=min_col,
                max_col=max_col,
                values_only=True,
            ),
            chunk_size,
        )
        return

    from wolfxl import _rust

    path = getattr(wb, "_source_path", None)
    if path is None:
        raise RuntimeError(
            "stream_value_chunks requires a source-path-bearing workbook "
            "(use load_workbook to obtain one)"
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
    stream_mx_r = mx_r if mx_r is not None else (
        default_bounds[2] if default_bounds is not None else None
    )
    stream_mx_c = mx_c if mx_c is not None else (
        default_bounds[3] if default_bounds is not None else None
    )
    direct_chunks = _direct_plain_value_chunks(
        wb,
        ws.title,
        chunk_size,
        mn_r,
        stream_mx_r,
        mn_c,
        stream_mx_c,
    )
    if direct_chunks is not None:
        yield from direct_chunks
        return

    reader = _rust.StreamingSheetReader.open(
        path, ws.title, mn_r, stream_mx_r, mn_c, stream_mx_c
    )
    try:
        read_chunk_with_flags = getattr(reader, "read_next_values_chunk_with_flags", None)
        if read_chunk_with_flags is None:
            yield from _chunk_value_rows(
                stream_iter_rows(
                    ws,
                    min_row=min_row,
                    max_row=max_row,
                    min_col=min_col,
                    max_col=max_col,
                    values_only=True,
                ),
                chunk_size,
            )
            return
        read_compact_chunk_with_flags = getattr(
            reader, "read_next_values_compact_chunk_with_flags", None
        )

        data_only = bool(getattr(wb, "_data_only", False))
        counter = mn_r if mn_r is not None else 1
        cmin = mn_c if mn_c is not None else 1
        fixed_cmax = mx_c if mx_c is not None else (
            default_bounds[3] if default_bounds is not None else None
        )
        empty_row = (
            (None,) * max(0, fixed_cmax + 1 - cmin)
            if fixed_cmax is not None
            else ()
        )
        pending: list[tuple[Any, ...]] = []
        while True:
            if read_compact_chunk_with_flags is not None:
                chunk_result = read_compact_chunk_with_flags(chunk_size)
            else:
                chunk_result = read_chunk_with_flags(chunk_size)
            if chunk_result is None:
                break
            if len(chunk_result) == 5:
                chunk, needs_normalise, first_row, last_row, dense_plain = chunk_result
            else:
                chunk, needs_normalise = chunk_result
                first_row = chunk[0][0] if chunk else counter
                last_row = chunk[-1][0] if chunk else counter - 1
                dense_plain = False
            if (
                dense_plain
                and not pending
                and chunk
                and first_row == counter
                and last_row == counter + len(chunk) - 1
            ):
                yield tuple(chunk)
                counter = last_row + 1
                continue
            if dense_plain:
                chunk = tuple((first_row + offset, values) for offset, values in enumerate(chunk))
            for row_idx, values in chunk:
                while counter < row_idx:
                    pending.append(empty_row)
                    if len(pending) >= chunk_size:
                        yield tuple(pending)
                        pending = []
                    counter += 1
                if needs_normalise:
                    values = _normalise_streaming_value_tuple(values, data_only=data_only)
                pending.append(values)
                if len(pending) >= chunk_size:
                    yield tuple(pending)
                    pending = []
                counter = row_idx + 1
        if stream_mx_r is not None:
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


def _direct_plain_value_chunks(
    wb: Any,
    sheet: str,
    chunk_size: int,
    min_row: int | None,
    max_row: int | None,
    min_col: int | None,
    max_col: int | None,
) -> Iterator[tuple[tuple[Any, ...], ...]] | None:
    """Use the value-only Rust parser for bounded plain read chunks.

    This is deliberately limited to data-only, style-date-free, small/medium
    scans. Larger or style-sensitive sheets keep the streaming path so
    ``read_only=True`` remains memory-friendly for huge workbooks.
    """
    if not bool(getattr(wb, "_data_only", False)):
        return None
    if max_row is None or max_row > AUTO_STREAM_ROW_THRESHOLD:
        return None
    reader = getattr(wb, "_rust_reader", None)
    read_chunks = getattr(reader, "read_sheet_value_chunks_plain", None)
    if read_chunks is None:
        return None

    chunks = read_chunks(
        sheet,
        chunk_size,
        min_row,
        max_row,
        min_col,
        max_col,
    )
    return iter(chunks)


def _chunk_value_rows(
    rows: Iterator[tuple[Any, ...]],
    chunk_size: int,
) -> Iterator[tuple[tuple[Any, ...], ...]]:
    pending: list[tuple[Any, ...]] = []
    for row in rows:
        pending.append(row)
        if len(pending) >= chunk_size:
            yield tuple(pending)
            pending = []
    if pending:
        yield tuple(pending)


def should_auto_stream(ws: Worksheet) -> bool:
    """Heuristic: auto-engage streaming for ``iter_rows`` on huge sheets.

    Triggers when the sheet's ``max_row`` exceeds
    :data:`AUTO_STREAM_ROW_THRESHOLD`. This deliberately consults the
    worksheet ``<dimension ref=...>`` tag directly from the ZIP head instead
    of ``Worksheet._max_row()``, because the normal dimensions API may parse
    and cache the full sheet before streaming can engage.
    """
    wb = ws._workbook  # noqa: SLF001
    path = getattr(wb, "_source_path", None)
    if not path:
        return False
    rows = _source_dimension_max_row(path, ws.title)
    return rows is not None and rows > AUTO_STREAM_ROW_THRESHOLD
