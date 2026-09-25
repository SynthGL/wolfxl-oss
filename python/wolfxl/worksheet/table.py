"""openpyxl.worksheet.table compatibility.

T1 makes ``Table``, ``TableStyleInfo``, and ``TableColumn`` real
dataclasses. Read access (``ws.tables["SalesTable"].ref``) works for any
file opened in read/modify mode. Write access (``ws.add_table(t)``) works
in write mode (T1 PR5).

Field naming follows openpyxl's camelCase convention (``displayName``,
``headerRowCount``, ``tableStyleInfo``) even though that's un-Pythonic —
the whole point of this shim is drop-in compatibility with openpyxl.
"""

from __future__ import annotations

from collections.abc import Iterator
from dataclasses import dataclass, field
from typing import Any

from wolfxl._compat import _install_openpyxl_iter, _openpyxl_name_fallback
from wolfxl.utils.cell import range_boundaries
from wolfxl.worksheet.filters import AutoFilter, SortState
from wolfxl.xml.constants import REL_NS, SHEET_MAIN_NS
from wolfxl.xml.functions import Element, SubElement, localname, tostring

PIVOTSTYLES = tuple(
    f"PivotStyle{kind}{idx}"
    for kind, max_idx in (("Medium", 28), ("Light", 28), ("Dark", 28))
    for idx in range(1, max_idx + 1)
)
TABLESTYLES = tuple(
    f"TableStyle{kind}{idx}"
    for kind, max_idx in (("Medium", 28), ("Light", 21), ("Dark", 11))
    for idx in range(1, max_idx + 1)
)


@dataclass
class TableColumn:
    """A single column within a table.

    openpyxl's ``TableColumn`` also carries calculated-column formulas,
    totals-row formulas, and style IDs. wolfxl preserves those on round-
    trip but does not expose them on the Python side yet; construction
    accepts just ``id`` + ``name``, which is what covers 99% of
    user-built tables.
    """

    id: int
    name: str | None = None
    uniqueName: str | None = None  # noqa: N815
    totalsRowFunction: str | None = None  # noqa: N815
    totalsRowLabel: str | None = None  # noqa: N815
    queryTableFieldId: int | None = None  # noqa: N815
    headerRowDxfId: int | None = None  # noqa: N815
    dataDxfId: int | None = None  # noqa: N815
    totalsRowDxfId: int | None = None  # noqa: N815
    headerRowCellStyle: str | None = None  # noqa: N815
    dataCellStyle: str | None = None  # noqa: N815
    totalsRowCellStyle: str | None = None  # noqa: N815
    calculatedColumnFormula: Any = None  # noqa: N815
    totalsRowFormula: Any = None  # noqa: N815
    xmlColumnPr: Any = None  # noqa: N815
    extLst: Any = None  # noqa: N815

    def to_tree(self, tagname: str | None = None, namespace: str | None = None) -> Any:
        node = Element(tagname or _tag("tableColumn", namespace))
        for attr in (
            "id",
            "uniqueName",
            "name",
            "totalsRowFunction",
            "totalsRowLabel",
            "queryTableFieldId",
            "headerRowDxfId",
            "dataDxfId",
            "totalsRowDxfId",
            "headerRowCellStyle",
            "dataCellStyle",
            "totalsRowCellStyle",
        ):
            _set_attr(node, attr, getattr(self, attr))
        for child_attr in (
            "calculatedColumnFormula",
            "totalsRowFormula",
            "xmlColumnPr",
            "extLst",
        ):
            child = getattr(self, child_attr)
            if child is not None and hasattr(child, "to_tree"):
                node.append(child.to_tree(namespace=namespace))
        return node

    @classmethod
    def from_tree(cls, node: Any) -> "TableColumn":
        kwargs: dict[str, Any] = dict(node.attrib)
        if "id" in kwargs:
            kwargs["id"] = int(kwargs["id"])
        for child in list(node):
            name = localname(child)
            if name in {"calculatedColumnFormula", "totalsRowFormula"}:
                kwargs[name] = TableFormula.from_tree(child)
            elif name == "xmlColumnPr":
                kwargs["xmlColumnPr"] = XMLColumnProps.from_tree(child)
        return cls(**kwargs)


@dataclass
class TableStyleInfo:
    """Table style reference (``name``) plus banded-row/column flags.

    Excel ships named styles like ``"TableStyleLight9"``; this object
    records which style and which banding options are active.
    """

    name: str | None = None
    showFirstColumn: bool = False  # noqa: N815 - openpyxl public API
    showLastColumn: bool = False  # noqa: N815
    showRowStripes: bool = False  # noqa: N815
    showColumnStripes: bool = False  # noqa: N815

    def to_tree(self, tagname: str | None = None, namespace: str | None = None) -> Any:
        node = Element(tagname or _tag("tableStyleInfo", namespace))
        _set_attr(node, "name", self.name)
        _set_attr(node, "showFirstColumn", self.showFirstColumn, skip_false=True)
        _set_attr(node, "showLastColumn", self.showLastColumn, skip_false=True)
        _set_attr(node, "showRowStripes", self.showRowStripes, skip_false=True)
        _set_attr(node, "showColumnStripes", self.showColumnStripes, skip_false=True)
        return node

    @classmethod
    def from_tree(cls, node: Any) -> "TableStyleInfo":
        return cls(
            name=node.get("name"),
            showFirstColumn=_bool_attr(node.get("showFirstColumn"), False),
            showLastColumn=_bool_attr(node.get("showLastColumn"), False),
            showRowStripes=_bool_attr(node.get("showRowStripes"), False),
            showColumnStripes=_bool_attr(node.get("showColumnStripes"), False),
        )


@dataclass(init=False)
class Table:
    """An Excel table (ListObject) — a named, styled range.

    ``name`` is the internal identifier; ``displayName`` is what users
    see in the Name Box. openpyxl allows them to differ but they usually
    match. ``ref`` is the A1 range string (e.g. ``"A1:D10"``).

    ``headerRowCount`` is 1 when the first row is a header, 0 otherwise.
    ``totalsRowCount`` is the number of totals rows at the bottom.

    When constructed from a Rust-side dict, boolean fields like
    ``header_row=True`` map to ``headerRowCount=1``.
    """

    mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"
    name: str = ""
    id: int = 1
    displayName: str = ""  # noqa: N815 - openpyxl public API
    ref: str = ""
    comment: str | None = None
    tableType: str | None = None  # noqa: N815
    headerRowCount: int = 1  # noqa: N815
    totalsRowCount: int = 0  # noqa: N815
    totalsRowShown: bool | None = None  # noqa: N815
    tableStyleInfo: TableStyleInfo | None = None  # noqa: N815
    tableColumns: list[TableColumn] = field(default_factory=list)  # noqa: N815
    autoFilter: AutoFilter | None = None  # noqa: N815
    sortState: SortState | None = None  # noqa: N815
    _rel_type = f"{REL_NS}/table"

    def __init__(
        self,
        id: int = 1,  # noqa: A002 - openpyxl public API
        name: str | None = None,
        displayName: str | None = None,  # noqa: N803
        ref: str = "",
        comment: str | None = None,
        tableType: str | None = None,  # noqa: N803
        headerRowCount: int = 1,  # noqa: N803
        totalsRowCount: int = 0,  # noqa: N803
        totalsRowShown: bool | None = None,  # noqa: N803
        autoFilter: AutoFilter | None = None,  # noqa: N803
        sortState: SortState | None = None,  # noqa: N803
        tableStyleInfo: TableStyleInfo | None = None,  # noqa: N803
        tableColumns: list[TableColumn] | None = None,  # noqa: N803
        extLst: Any = None,  # noqa: N803
    ) -> None:
        resolved_name = name or displayName
        if not resolved_name:
            raise TypeError("Table requires 'name' or openpyxl alias 'displayName'")
        self.id = int(id)
        self.name = resolved_name
        self.displayName = displayName or resolved_name
        self.ref = ref
        self.comment = comment
        self.tableType = tableType
        self.headerRowCount = headerRowCount
        self.totalsRowCount = totalsRowCount
        self.totalsRowShown = totalsRowShown
        self.autoFilter = autoFilter
        self.sortState = sortState
        self.tableStyleInfo = tableStyleInfo
        self.tableColumns = list(tableColumns or [])
        self.extLst = extLst

    @property
    def path(self) -> str:
        return f"/xl/tables/table{self.id}.xml"

    @property
    def column_names(self) -> list[str | None]:
        return [column.name for column in self.tableColumns]

    def _initialise_columns(self) -> None:
        min_col, _min_row, max_col, _max_row = range_boundaries(self.ref)
        if min_col is None or max_col is None:
            return
        if self.autoFilter is None:
            self.autoFilter = AutoFilter(ref=self.ref)
        if not self.tableColumns:
            self.tableColumns = [
                TableColumn(id=idx, name=f"Column{idx}")
                for idx in range(1, max_col - min_col + 2)
            ]

    def to_tree(self) -> Any:
        node = Element(f"{{{SHEET_MAIN_NS}}}table")
        for attr in (
            "id",
            "name",
            "displayName",
            "comment",
            "ref",
            "tableType",
            "headerRowCount",
            "totalsRowCount",
            "totalsRowShown",
        ):
            if attr == "totalsRowCount" and getattr(self, attr) == 0:
                continue
            _set_attr(node, attr, getattr(self, attr))
        if self.autoFilter is not None:
            node.append(
                self.autoFilter.to_tree(tagname=f"{{{SHEET_MAIN_NS}}}autoFilter")
            )
        if self.sortState is not None:
            node.append(self.sortState.to_tree(tagname=f"{{{SHEET_MAIN_NS}}}sortState"))
        if self.tableColumns:
            columns = SubElement(node, f"{{{SHEET_MAIN_NS}}}tableColumns")
            columns.set("count", str(len(self.tableColumns)))
            for column in self.tableColumns:
                columns.append(column.to_tree(namespace=SHEET_MAIN_NS))
        if self.tableStyleInfo is not None:
            node.append(self.tableStyleInfo.to_tree(namespace=SHEET_MAIN_NS))
        return node

    @classmethod
    def from_tree(cls, node: Any) -> "Table":
        kwargs: dict[str, Any] = dict(node.attrib)
        for attr in ("id", "headerRowCount", "totalsRowCount"):
            if attr in kwargs:
                kwargs[attr] = int(kwargs[attr])
        if "totalsRowShown" in kwargs:
            kwargs["totalsRowShown"] = _bool_attr(kwargs["totalsRowShown"])
        for child in list(node):
            name = localname(child)
            if name == "autoFilter":
                kwargs["autoFilter"] = AutoFilter.from_tree(child)
            elif name == "sortState":
                kwargs["sortState"] = SortState.from_tree(child)
            elif name == "tableColumns":
                kwargs["tableColumns"] = [
                    TableColumn.from_tree(grandchild)
                    for grandchild in list(child)
                    if localname(grandchild) == "tableColumn"
                ]
            elif name == "tableStyleInfo":
                kwargs["tableStyleInfo"] = TableStyleInfo.from_tree(child)
        return cls(**kwargs)

    def _write(self, archive: Any) -> None:
        archive.writestr(self.path[1:], tostring(self.to_tree()))


@dataclass
class TableFormula:
    array: bool | None = None
    attr_text: str | None = None
    text: str | None = None

    def to_tree(self, tagname: str | None = None, namespace: str | None = None) -> Any:
        node = Element(tagname or _tag("tableFormula", namespace))
        _set_attr(node, "array", self.array, skip_false=True)
        _set_attr(node, "attr_text", self.attr_text)
        node.text = self.text
        return node

    @classmethod
    def from_tree(cls, node: Any) -> "TableFormula":
        return cls(
            array=_bool_attr(node.get("array")),
            attr_text=node.get("attr_text"),
            text=node.text,
        )


# ---------------------------------------------------------------------------
# Pod 2 (RFC-060 §2.1) — extended re-exports.
# Sprint Π Pod-β (RFC-063) replaced the four construction stubs below
# (TableList / TablePartList / Related / XMLColumnProps) with real
# value types.  No save() pipeline changes — these are openpyxl-shaped
# wrappers over state that already plumbs through the patcher / native
# writer via the existing ``Table`` class and ``ws.tables`` dict.
# ---------------------------------------------------------------------------

@dataclass
class Related:
    """rels-pointer dataclass mirroring ``r:id="rId1"``.

    Used by openpyxl to point a ``<tablePart>`` element at the table's
    relationship entry.  Wolfxl tracks the rId allocation internally,
    but the dataclass is exposed here so user code that hand-builds a
    :class:`TablePartList` continues to work.
    """

    id: str = ""

    def to_tree(self, tagname: str | None = None, idx: Any = None) -> Any:  # noqa: ARG002
        node = Element(tagname or "tablePart")
        if self.id:
            node.set(f"{{{REL_NS}}}id", self.id)
        return node

    @classmethod
    def from_tree(cls, node: Any) -> "Related":
        return cls(id=_relationship_id(node))


@dataclass
class XMLColumnProps:
    """XML-column metadata for table-bound columns (CT_XmlColumnPr).

    Wolfxl preserves these properties on round-trip via the patcher.
    Construction is exposed here so user code that explicitly attaches
    a ``XMLColumnProps`` to a :class:`TableColumn` (openpyxl's
    ``column.xmlColumnPr``) ports mechanically.
    """

    mapId: int = 0  # noqa: N815 - openpyxl public API
    xpath: str = ""
    denormalized: bool = False
    xmlDataType: str = "string"  # noqa: N815

    def to_tree(self, tagname: str | None = None, namespace: str | None = None) -> Any:
        node = Element(tagname or _tag("xmlColumnPr", namespace))
        _set_attr(node, "mapId", self.mapId)
        _set_attr(node, "xpath", self.xpath)
        _set_attr(node, "denormalized", self.denormalized, skip_false=True)
        _set_attr(node, "xmlDataType", self.xmlDataType)
        return node

    @classmethod
    def from_tree(cls, node: Any) -> "XMLColumnProps":
        return cls(
            mapId=node.get("mapId"),
            xpath=node.get("xpath"),
            denormalized=_bool_attr(node.get("denormalized"), False),
            xmlDataType=node.get("xmlDataType"),
        )


class TableList(dict[str, Any]):
    """Openpyxl-shaped table mapping.

    Openpyxl implements ``Worksheet.tables`` as a ``dict`` subclass whose
    values are :class:`Table` objects, while ``items()`` reports
    ``(name, ref)`` pairs.  WolfXL keeps the same object shape and, when
    bound to a worksheet, mirrors ``add()`` calls onto the worksheet's
    pending-tables queue so save picks them up automatically.
    """

    __slots__ = ("loaded", "worksheet")

    def __init__(self, worksheet: Any = None, values: dict[str, Any] | None = None) -> None:
        super().__init__(values or {})
        self.worksheet = worksheet
        # Tables read from the source file in modify mode, keyed by their
        # source name: (table, saved ref, saved column names). Save compares
        # each table against this record to queue growth of the part.
        self.loaded: dict[str, tuple[Table, str, tuple[str, ...]]] = {}

    # ------------------------------------------------------------------
    # openpyxl-shape API
    # ------------------------------------------------------------------

    def add(self, table: Any) -> None:
        """Register *table* on the underlying worksheet (or local view)."""
        if not isinstance(table, Table):
            raise TypeError("You can only add tables")
        name = getattr(table, "name", None) or getattr(table, "displayName", None)
        if not name:
            raise ValueError("TableList.add: table must expose a non-empty `name`")
        self[name] = table
        ws = self.worksheet
        if ws is not None:
            if getattr(ws, "_tables_cache", None) is None:
                ws._tables_cache = self  # noqa: SLF001
            elif ws._tables_cache is not self:  # noqa: SLF001
                ws._tables_cache[name] = table  # noqa: SLF001
            pending = getattr(ws, "_pending_tables", None)
            if pending is not None and table not in pending:
                pending.append(table)

    def remove(self, table_name: str) -> None:
        """Remove a table by *name*.  Silently no-ops if the name is unknown."""
        self.pop(table_name, None)

    def get(
        self,
        name: str | None = None,
        default: Any = None,
        table_range: str | None = None,
    ) -> Any:
        if name is not None:
            return super().get(name, default)
        if table_range is not None:
            for table in super().values():
                if getattr(table, "ref", None) == table_range:
                    return table
        return default

    def items(self) -> list[tuple[str, str]]:
        return [(name, table.ref) for name, table in super().items()]

    def __repr__(self) -> str:  # pragma: no cover - trivial
        return f"TableList(count={len(self)})"


@dataclass
class TablePartList:
    """`<tableParts>` serialization helper (CT_TableParts §18.3.1.91).

    A simple holder for the count + list of :class:`Related` pointers
    Excel writes into ``<tableParts>`` underneath each ``<worksheet>``.
    Wolfxl regenerates this block from ``ws.tables`` at save time, so
    this dataclass is informational — but exposed for openpyxl source
    compatibility.
    """

    count: int = 0
    tablePart: list[Related] = field(default_factory=list)  # noqa: N815

    def __post_init__(self) -> None:
        if self.tablePart is None:  # pragma: no cover - defensive
            self.tablePart = []
        # Keep ``count`` and the list length in sync when the user only
        # supplied one of them.
        if self.count == 0 and self.tablePart:
            self.count = len(self.tablePart)

    def append(self, part: Related) -> None:
        """Add a :class:`Related` entry and bump :attr:`count`."""
        self.tablePart.append(part)
        self.count = len(self.tablePart)

    def __iter__(self) -> Iterator[Related]:
        return iter(self.tablePart)

    def __len__(self) -> int:
        return len(self.tablePart)

    def to_tree(self) -> Any:
        node = Element("tableParts")
        node.set("count", str(len(self.tablePart)))
        for part in self.tablePart:
            node.append(part.to_tree())
        return node

    @classmethod
    def from_tree(cls, node: Any) -> "TablePartList":
        parts = [
            Related.from_tree(child)
            for child in list(node)
            if localname(child) == "tablePart"
        ]
        return cls(count=len(parts), tablePart=parts)


def _bool_attr(value: Any, default: bool | None = None) -> bool | None:
    if value is None:
        return default
    if isinstance(value, bool):
        return value
    return str(value).lower() in {"1", "true"}


def _relationship_id(node: Any) -> str:
    return node.get(f"{{{REL_NS}}}id") or node.get("r:id") or node.get("id") or ""


def _tag(name: str, namespace: str | None = None) -> str:
    return f"{{{namespace}}}{name}" if namespace else name


def _set_attr(
    node: Any,
    name: str,
    value: Any,
    *,
    skip_false: bool = False,
) -> None:
    if value is None:
        return
    if skip_false and value is False:
        return
    if isinstance(value, bool):
        value = "1" if value else "0"
    node.set(name, str(value))


_install_openpyxl_iter(
    TableStyleInfo,
    Related,
    TableColumn,
    TableFormula,
    XMLColumnProps,
    Table,
)

__all__ = [
    "AutoFilter",
    "PIVOTSTYLES",
    "Related",
    "SortState",
    "Table",
    "TableColumn",
    "TableFormula",
    "TableList",
    "TablePartList",
    "TableStyleInfo",
    "TABLESTYLES",
    "XMLColumnProps",
]

__getattr__ = _openpyxl_name_fallback(globals())
