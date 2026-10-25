"""Worksheet cell watches: the entries of Excel's Watch Window.

A worksheet stores its watched cells in a ``<cellWatches>`` block of its
OOXML part, one ``<cellWatch r="B2"/>`` per entry. :class:`CellWatch` and
:class:`CellWatches` mirror the openpyxl model classes of the same names and
add editing helpers (``add``, ``edit``, ``remove``, ``clear``).

:func:`worksheet_cell_watches` backs ``ws.cell_watches``: it loads the
watches stored in the worksheet's source package once, and
:func:`apply_cell_watches_authoring` writes them into the package the
workbook saves: edited collections in modify mode, every non-empty
collection in write mode, whose backend regenerates each sheet part.
Unedited worksheets of a loaded workbook keep their source block byte for
byte.

Persistence also works directly on a saved package:

.. code-block:: python

    wb.save("book.xlsx")
    watches = CellWatches(["B2"])
    watches.add("C3")
    save_cell_watches("book.xlsx", {"Sheet1": watches})
    load_cell_watches("book.xlsx", "Sheet1").refs  # ("B2", "C3")

``save_cell_watches`` rewrites only the ``<cellWatches>`` block of each named
sheet part; every other package member is copied unchanged. An empty
collection removes the block, because the schema requires at least one entry.
"""

from __future__ import annotations

import os
import re
import sys
import tempfile
import zipfile
from collections.abc import Iterable, Iterator, Mapping
from io import BytesIO
from typing import Any, Union

from wolfxl._compat import _OpenpyxlSerialisable, _xml_element_factory

Sequence = Serialisable = String = _OpenpyxlSerialisable

_MAIN_NAMESPACES = frozenset(
    {
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "http://purl.oclc.org/ooxml/spreadsheetml/main",
    }
)
_REL_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_MAX_ROW = 1_048_576
_MAX_COLUMN = 16_384
_CELL_REF = re.compile(r"\$?([A-Za-z]{1,3})\$?([0-9]{1,7})")

# ECMA-376 CT_Worksheet children that precede ``cellWatches``.
_BEFORE_CELL_WATCHES = frozenset(
    {
        "sheetPr",
        "dimension",
        "sheetViews",
        "sheetFormatPr",
        "cols",
        "sheetData",
        "sheetCalcPr",
        "sheetProtection",
        "protectedRanges",
        "scenarios",
        "autoFilter",
        "sortState",
        "dataConsolidate",
        "customSheetViews",
        "mergeCells",
        "phoneticPr",
        "conditionalFormatting",
        "dataValidations",
        "hyperlinks",
        "printOptions",
        "pageMargins",
        "pageSetup",
        "headerFooter",
        "rowBreaks",
        "colBreaks",
        "customProperties",
    }
)

_NAME = rb"[A-Za-z_][\w.\-]*(?::[A-Za-z_][\w.\-]*)?"
_TOKEN = re.compile(
    rb"<!--.*?-->|<!\[CDATA\[.*?\]\]>|<\?.*?\?>|"
    rb"<(/?)(" + _NAME + rb")((?:\s+[^\s=/>]+\s*=\s*(?:\"[^\"]*\"|'[^']*'))*)\s*(/?)>",
    re.DOTALL,
)


def _normalize_ref(ref: Any) -> str:
    """Return the canonical relative A1 form of one in-grid cell reference."""
    if not isinstance(ref, str):
        raise TypeError(f"cell watch reference must be a str, got {type(ref).__name__}")
    match = _CELL_REF.fullmatch(ref.strip())
    if match is None:
        raise ValueError(f"cell watch reference must be one cell like 'B2', got {ref!r}")
    letters, digits = match.groups()
    column = 0
    for letter in letters.upper():
        column = column * 26 + ord(letter) - 64
    row = int(digits)
    if not 1 <= column <= _MAX_COLUMN or not 1 <= row <= _MAX_ROW:
        raise ValueError(f"cell watch reference {ref!r} is outside the worksheet grid")
    return f"{letters.upper()}{row}"


def _ref_from_position(row: Any, column: Any) -> str:
    from wolfxl.utils.cell import get_column_letter

    if isinstance(row, bool) or not isinstance(row, int):
        raise TypeError("row must be an int")
    if isinstance(column, bool) or not isinstance(column, int):
        raise TypeError("column must be an int")
    if not 1 <= column <= _MAX_COLUMN or not 1 <= row <= _MAX_ROW:
        raise ValueError(f"cell ({row}, {column}) is outside the worksheet grid")
    return f"{get_column_letter(column)}{row}"


def _split_ref(ref: str) -> tuple[int, int]:
    from wolfxl.utils.cell import coordinate_to_tuple

    return coordinate_to_tuple(ref)


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _namespace(tag: str) -> str:
    return tag[1:].split("}", 1)[0] if tag.startswith("{") else ""


class CellWatch:
    """One Watch Window entry: a single watched cell on its worksheet."""

    tagname = "cellWatch"
    __attrs__ = ("r",)

    def __init__(self, r: str | None = None) -> None:
        self.r = r

    @property
    def r(self) -> str | None:
        """Watched cell in relative A1 form, such as ``"B2"``."""
        return self._r

    @r.setter
    def r(self, value: str | None) -> None:
        self._r = None if value is None else _normalize_ref(value)

    @property
    def row(self) -> int | None:
        return None if self._r is None else _split_ref(self._r)[0]

    @row.setter
    def row(self, value: int) -> None:
        if self._r is None:
            raise ValueError("cannot set the row of a cell watch without a reference")
        self._r = _ref_from_position(value, _split_ref(self._r)[1])

    @property
    def column(self) -> int | None:
        return None if self._r is None else _split_ref(self._r)[1]

    @column.setter
    def column(self, value: int) -> None:
        if self._r is None:
            raise ValueError("cannot set the column of a cell watch without a reference")
        self._r = _ref_from_position(_split_ref(self._r)[0], value)

    def to_tree(
        self,
        tagname: str | None = None,
        idx: int | None = None,
        namespace: str | None = None,
    ) -> Any:
        name = tagname or self.tagname
        element = _xml_element_factory()(f"{{{namespace}}}{name}" if namespace else name)
        if self._r is not None:
            element.set("r", self._r)
        return element

    @classmethod
    def from_tree(cls, node: Any) -> CellWatch:
        ref = node.attrib.get("r")
        if ref is None:
            raise ValueError("cellWatch element has no r attribute")
        return cls(ref)

    def __eq__(self, other: object) -> bool:
        if not isinstance(other, CellWatch):
            return NotImplemented
        return self._r == other._r

    __hash__ = None  # type: ignore[assignment] - mutable

    def __repr__(self) -> str:
        return f"<CellWatch r={self._r!r}>"


if sys.version_info >= (3, 10):
    _WatchLike = CellWatch | str
else:
    _WatchLike = Union[CellWatch, str]


def _coerce_watch(value: Any) -> CellWatch:
    if isinstance(value, CellWatch):
        return value
    if isinstance(value, str):
        return CellWatch(value)
    raise TypeError(f"expected CellWatch or str, got {type(value).__name__}")


def _target_ref(ref: Any, column: Any = None) -> str:
    if column is not None:
        return _ref_from_position(ref, column)
    if isinstance(ref, CellWatch):
        if ref.r is None:
            raise ValueError("cell watch has no reference")
        return ref.r
    return _normalize_ref(ref)


class CellWatches:
    """Ordered, editable collection of a worksheet's Watch Window entries.

    ``cellWatch`` is the backing list, named as in openpyxl. Entries keep
    their insertion order, which is the order Excel lists them in.
    """

    tagname = "cellWatches"
    __elements__ = ("cellWatch",)

    def __init__(self, cellWatch: Iterable[_WatchLike] = ()) -> None:
        self.cellWatch: list[CellWatch] = [_coerce_watch(item) for item in cellWatch]

    # ------------------------------------------------------------------
    # Editing
    # ------------------------------------------------------------------

    def add(self, ref: _WatchLike | int, column: int | None = None) -> CellWatch:
        """Watch a cell given as ``"B2"``, ``(row, column)`` or a :class:`CellWatch`.

        Watching an already watched cell returns the existing entry, matching
        Excel's Watch Window, which lists each cell once.
        """
        target = _target_ref(ref, column)
        for watch in self.cellWatch:
            if watch.r == target:
                return watch
        watch = ref if isinstance(ref, CellWatch) else CellWatch(target)
        self.cellWatch.append(watch)
        return watch

    def edit(self, ref: _WatchLike, new_ref: _WatchLike) -> CellWatch:
        """Point the entry watching ``ref`` at ``new_ref`` and return it."""
        old = _target_ref(ref)
        new = _target_ref(new_ref)
        index = self.index(old)
        if new != old and new in self:
            raise ValueError(f"cell {new} is already watched")
        watch = self.cellWatch[index]
        watch.r = new
        return watch

    def remove(self, ref: _WatchLike | int, column: int | None = None) -> None:
        """Stop watching a cell; raises :class:`KeyError` when it is not watched."""
        del self.cellWatch[self.index(ref, column)]

    def clear(self) -> None:
        self.cellWatch.clear()

    def index(self, ref: _WatchLike | int, column: int | None = None) -> int:
        target = _target_ref(ref, column)
        for position, watch in enumerate(self.cellWatch):
            if watch.r == target:
                return position
        raise KeyError(target)

    @property
    def refs(self) -> tuple[str | None, ...]:
        """Watched cell references in Watch Window order."""
        return tuple(watch.r for watch in self.cellWatch)

    # ------------------------------------------------------------------
    # Container protocol
    # ------------------------------------------------------------------

    def __len__(self) -> int:
        return len(self.cellWatch)

    def __bool__(self) -> bool:
        return bool(self.cellWatch)

    def __iter__(self) -> Iterator[CellWatch]:
        return iter(self.cellWatch)

    def __getitem__(self, index: int) -> CellWatch:
        return self.cellWatch[index]

    def __contains__(self, item: object) -> bool:
        if isinstance(item, CellWatch):
            target = item.r
        elif isinstance(item, str):
            try:
                target = _normalize_ref(item)
            except ValueError:
                return False
        else:
            return False
        return target is not None and any(watch.r == target for watch in self.cellWatch)

    def __eq__(self, other: object) -> bool:
        if not isinstance(other, CellWatches):
            return NotImplemented
        return self.refs == other.refs

    __hash__ = None  # type: ignore[assignment] - mutable

    def __repr__(self) -> str:
        return f"<CellWatches {list(self.refs)!r}>"

    # ------------------------------------------------------------------
    # XML
    # ------------------------------------------------------------------

    def to_tree(
        self,
        tagname: str | None = None,
        idx: int | None = None,
        namespace: str | None = None,
    ) -> Any:
        name = tagname or self.tagname
        element = _xml_element_factory()(f"{{{namespace}}}{name}" if namespace else name)
        for watch in self.cellWatch:
            if not isinstance(watch, CellWatch):
                raise TypeError(f"expected CellWatch, got {type(watch).__name__}")
            element.append(watch.to_tree(namespace=namespace))
        return element

    @classmethod
    def from_tree(cls, node: Any) -> CellWatches:
        namespace = _namespace(node.tag)
        watches = []
        for child in node:
            if not isinstance(child.tag, str):
                continue  # comments and processing instructions
            if _local(child.tag) != "cellWatch" or _namespace(child.tag) != namespace:
                raise ValueError(f"unexpected {child.tag!r} element in cellWatches")
            watches.append(CellWatch.from_tree(child))
        return cls(watches)

    def _xml_bytes(self, prefix: bytes) -> bytes:
        refs = []
        for watch in self.cellWatch:
            if not isinstance(watch, CellWatch):
                raise TypeError(f"expected CellWatch, got {type(watch).__name__}")
            if watch.r is None:
                raise ValueError("cannot save a cell watch without a reference")
            refs.append(watch.r.encode("ascii"))
        if not refs:
            return b""
        body = b"".join(b"<" + prefix + b'cellWatch r="' + ref + b'"/>' for ref in refs)
        return b"<" + prefix + b"cellWatches>" + body + b"</" + prefix + b"cellWatches>"

    @classmethod
    def from_worksheet(cls, worksheet: Any) -> CellWatches:
        """Load the watches stored in a loaded worksheet's source package.

        A worksheet that was not loaded from a file has no stored watches.
        """
        if not getattr(worksheet, "_source_backed", False):
            return cls()
        workbook = getattr(worksheet, "_workbook", None) or worksheet.parent
        source_bytes = getattr(workbook, "_source_bytes", None)
        source_path = getattr(workbook, "_source_path", None)
        source: Any = source_bytes if source_bytes is not None else source_path
        if source is None:
            return cls()
        return load_cell_watches(source, worksheet._source_title)


# ----------------------------------------------------------------------
# Package persistence
# ----------------------------------------------------------------------


def _parse_xml(data: bytes) -> Any:
    from defusedxml import ElementTree as SafeET

    if b"<!DOCTYPE" in data.upper():
        raise ValueError("XML with a document type declaration is not accepted")
    return SafeET.fromstring(data)


def _read_parts(
    archive: zipfile.ZipFile,
) -> tuple[list[zipfile.ZipInfo], dict[str, bytes]]:
    infos = archive.infolist()
    names = [info.filename for info in infos]
    if len(set(names)) != len(names):
        raise ValueError("package has duplicate member names")
    return infos, {info.filename: archive.read(info.filename) for info in infos}


def _sheet_paths(parts: Mapping[str, bytes]) -> dict[str, str]:
    """Map worksheet titles to their part names through the workbook relationships."""
    import posixpath

    workbook_xml = parts.get("xl/workbook.xml")
    if workbook_xml is None:
        raise ValueError("package has no xl/workbook.xml part")
    rels_xml = parts.get("xl/_rels/workbook.xml.rels")
    targets: dict[str, str] = {}
    worksheet_targets: list[str] = []
    if rels_xml is not None:
        for rel in _parse_xml(rels_xml):
            if _local(rel.tag) != "Relationship" or rel.get("TargetMode") == "External":
                continue
            rel_id, target = rel.get("Id"), rel.get("Target")
            if rel_id and target:
                targets[rel_id] = (
                    target.lstrip("/")
                    if target.startswith("/")
                    else posixpath.normpath(posixpath.join("xl", target))
                )
                if (rel.get("Type") or "").endswith("/worksheet"):
                    worksheet_targets.append(targets[rel_id])
    paths: dict[str, str] = {}
    declared = False
    for sheet in _parse_xml(workbook_xml).iter():
        if _local(sheet.tag) != "sheet":
            continue
        declared = True
        name = sheet.get("name")
        target = targets.get(sheet.get(f"{{{_REL_NAMESPACE}}}id") or "")
        if name and target:
            paths[name] = target
    if not declared:
        # ``load_workbook(permissive=True)`` synthesizes Sheet1, Sheet2, ...
        # titles from the worksheet relationships when ``<sheets>`` is empty.
        for index, target in enumerate(worksheet_targets, start=1):
            paths[f"Sheet{index}"] = target
    return paths


def _sheet_part(parts: Mapping[str, bytes], paths: Mapping[str, str], title: str) -> str:
    path = paths.get(title)
    if path is None or path not in parts:
        raise KeyError(f"worksheet {title!r} not found in package")
    return path


def _watches_from_root(root: Any) -> CellWatches:
    if _local(root.tag) != "worksheet" or _namespace(root.tag) not in _MAIN_NAMESPACES:
        raise ValueError("part is not a SpreadsheetML worksheet")
    blocks = [
        child
        for child in root
        if isinstance(child.tag, str)
        and _local(child.tag) == "cellWatches"
        and _namespace(child.tag) == _namespace(root.tag)
    ]
    if len(blocks) > 1:
        raise ValueError("worksheet has more than one cellWatches element")
    return CellWatches.from_tree(blocks[0]) if blocks else CellWatches()


def load_cell_watches(source: Any, sheet: str) -> CellWatches:
    """Read the watches of worksheet ``sheet`` from a package.

    ``source`` is a path, the package bytes, or a binary file object.
    """
    if isinstance(source, (bytes, bytearray, memoryview)):
        source = BytesIO(bytes(source))
    with zipfile.ZipFile(source) as archive:
        _, parts = _read_parts(archive)
    path = _sheet_part(parts, _sheet_paths(parts), sheet)
    return _watches_from_root(_parse_xml(parts[path]))


def _scan_worksheet(
    data: bytes,
) -> tuple[bytes, int, list[tuple[str, int, int]]]:
    """Return the root prefix, root start-tag end, and direct child spans."""
    depth = 0
    prefix = b""
    root_end = -1
    children: list[tuple[str, int, int]] = []
    open_child: tuple[str, int] | None = None
    for match in _TOKEN.finditer(data):
        closing, name, _attrs, self_closing = match.groups()
        if name is None:
            continue
        if closing:
            depth -= 1
            if depth == 1 and open_child is not None:
                children.append((open_child[0], open_child[1], match.end()))
                open_child = None
            elif depth == 0:
                return prefix, root_end, children
            continue
        if depth == 0:
            if self_closing:
                break
            prefix = name[: name.index(b":") + 1] if b":" in name else b""
            root_end = match.end()
        elif depth == 1:
            local = name.rsplit(b":", 1)[-1].decode("ascii")
            if self_closing:
                children.append((local, match.start(), match.end()))
                continue
            open_child = (local, match.start())
        if not self_closing:
            depth += 1
    raise ValueError("could not locate the worksheet element boundaries")


def _splice_cell_watches(data: bytes, watches: CellWatches) -> bytes:
    """Replace, insert, or remove the worksheet's ``cellWatches`` block."""
    root = _parse_xml(data)
    if _watches_from_root(root) == watches:
        return data
    prefix, root_end, children = _scan_worksheet(data)
    parsed_children = [_local(child.tag) for child in root if isinstance(child.tag, str)]
    if [name for name, _, _ in children] != parsed_children:
        raise ValueError("could not locate the worksheet child elements")
    replacement = watches._xml_bytes(prefix)
    existing = [(start, end) for name, start, end in children if name == "cellWatches"]
    if existing:
        start, end = existing[0]
        return data[:start] + replacement + data[end:]
    insert_at = root_end
    for name, _, end in children:
        if name in _BEFORE_CELL_WATCHES:
            insert_at = end
    return data[:insert_at] + replacement + data[insert_at:]


def save_cell_watches(
    filename: str | os.PathLike[str],
    watches: Mapping[str, CellWatches | Iterable[_WatchLike]],
) -> tuple[str, ...]:
    """Store the watches of each named worksheet into a saved package in place.

    ``watches`` maps worksheet titles to their complete watch list; sheets not
    named keep their parts unchanged. Every title and entry is validated before
    the file is replaced, so a failure leaves it untouched. Returns the titles
    whose worksheet parts changed.
    """
    resolved = {
        title: value if isinstance(value, CellWatches) else CellWatches(value)
        for title, value in watches.items()
    }
    with zipfile.ZipFile(filename) as archive:
        infos, parts = _read_parts(archive)
    paths = _sheet_paths(parts)
    changed: list[str] = []
    for title, value in resolved.items():
        path = _sheet_part(parts, paths, title)
        updated = _splice_cell_watches(parts[path], value)
        if updated != parts[path]:
            parts[path] = updated
            changed.append(title)
    if not changed:
        return ()

    fd, tmp_name = tempfile.mkstemp(
        prefix=".wolfxl-cell-watches-",
        suffix=".xlsx",
        dir=os.path.dirname(os.path.abspath(filename)) or None,
    )
    os.close(fd)
    try:
        with zipfile.ZipFile(tmp_name, "w", zipfile.ZIP_DEFLATED) as dst:
            for info in infos:
                dst.writestr(info, parts[info.filename])
        os.replace(tmp_name, filename)
    finally:
        if os.path.exists(tmp_name):
            os.unlink(tmp_name)
    return tuple(changed)


# ----------------------------------------------------------------------
# Worksheet binding
# ----------------------------------------------------------------------


def _watch_states(
    workbook: Any,
) -> dict[Any, tuple[CellWatches, tuple[str | None, ...]]]:
    """Map each worksheet whose watches were read to ``(watches, saved refs)``."""
    states = getattr(workbook, "_cell_watch_states", None)
    if states is None:
        states = {}
        workbook._cell_watch_states = states
    return states


def worksheet_cell_watches(worksheet: Any) -> CellWatches:
    """Return the live, editable watch list behind ``ws.cell_watches``.

    The first call loads the watches stored in a loaded worksheet's source
    package; later calls return the same collection.
    """
    states = _watch_states(worksheet.parent)
    state = states.get(worksheet)
    if state is None:
        watches = CellWatches.from_worksheet(worksheet)
        state = states[worksheet] = (watches, watches.refs)
    return state[0]


def _live_watch_states(
    workbook: Any,
) -> Iterator[tuple[Any, CellWatches, tuple[str | None, ...]]]:
    sheets = set(map(id, getattr(workbook, "_sheets", {}).values()))
    for worksheet, (watches, saved) in _watch_states(workbook).items():
        if id(worksheet) in sheets:
            yield worksheet, watches, saved


def _pending_watches(workbook: Any) -> dict[str, CellWatches]:
    return {
        worksheet.title: watches
        for worksheet, watches, saved in _live_watch_states(workbook)
        if watches.refs != saved
    }


def _watches_to_write(workbook: Any) -> dict[str, CellWatches]:
    """Watch lists the next save must splice into its package.

    A modify-mode save copies untouched worksheet parts byte for byte, so only
    edited lists are written. The write-mode backend regenerates every part on
    each save, so every non-empty list is written again.
    """
    regenerates = getattr(workbook, "_rust_writer", None) is not None
    return {
        worksheet.title: watches
        for worksheet, watches, saved in _live_watch_states(workbook)
        if watches.refs != saved or (regenerates and watches)
    }


def has_pending_cell_watches(workbook: Any) -> bool:
    """Whether any worksheet's watches differ from the last saved package."""
    return bool(_pending_watches(workbook))


def has_cell_watch_authoring(workbook: Any) -> bool:
    """Whether the next save has to write worksheet watches into its package."""
    return bool(_watches_to_write(workbook))


def apply_cell_watches_authoring(workbook: Any, filename: str | os.PathLike[str]) -> None:
    """Write the worksheet watch lists the save owns into the package ``filename``."""
    pending = _watches_to_write(workbook)
    if pending:
        save_cell_watches(filename, pending)


def mark_cell_watches_saved(workbook: Any) -> None:
    """Record the current watch lists as saved once publication succeeded."""
    states = _watch_states(workbook)
    for worksheet, (watches, _saved) in states.items():
        states[worksheet] = (watches, watches.refs)


__all__ = [
    "CellWatch",
    "CellWatches",
    "Sequence",
    "Serialisable",
    "String",
    "load_cell_watches",
    "save_cell_watches",
]
