"""Chartsheet containers compatible with ``openpyxl.chartsheet``."""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING, Any

from wolfxl._compat import _iter_openpyxl_attrs, _openpyxl_name_fallback
from wolfxl.chartsheet.relation import DrawingHF, SheetBackgroundPicture
from wolfxl.chartsheet.views import ChartsheetViewList
from wolfxl.styles.colors import Color
from wolfxl.utils.protection import hash_password
from wolfxl.worksheet.header_footer import HeaderFooter
from wolfxl.worksheet.page import PageMargins, PrintPageSetup
from wolfxl.xml.constants import REL_NS, SHEET_MAIN_NS
from wolfxl.xml.functions import Element, localname

if TYPE_CHECKING:
    from wolfxl.drawing.spreadsheet_drawing import SpreadsheetDrawing


@dataclass
class ChartsheetProperties:
    """Small openpyxl-shaped chartsheet properties container."""

    tagname = "sheetPr"
    __attrs__ = ("published", "codeName")
    __elements__ = ("tabColor",)

    published: bool | None = None
    codeName: str | None = None  # noqa: N815 - openpyxl camelCase
    tabColor: Any = None  # noqa: N815

    def __iter__(self):
        yield from _iter_openpyxl_attrs(self, self.__attrs__)

    @classmethod
    def from_tree(cls, node: Any) -> "ChartsheetProperties":
        props = cls(
            published=_to_bool(node.get("published")),
            codeName=node.get("codeName"),
        )
        for child in list(node):
            if localname(child) == "tabColor":
                props.tabColor = Color.from_tree(child)
        return props

    def to_tree(self, tagname: str | None = None, idx: Any = None) -> Any:  # noqa: ARG002
        node = Element(tagname or self.tagname)
        _set_attrs(node, self)
        if self.tabColor is not None:
            node.append(self.tabColor.to_tree("tabColor"))
        return node

    def _is_empty(self) -> bool:
        return not any(value is not None for _, value in self) and self.tabColor is None


@dataclass(init=False)
class ChartsheetProtection:
    """Small openpyxl-shaped chartsheet protection container."""

    tagname = "sheetProtection"
    __attrs__ = (
        "content",
        "objects",
        "password",
        "hashValue",
        "spinCount",
        "saltValue",
        "algorithmName",
    )

    content: bool | None = None
    objects: bool | None = None
    algorithmName: str | None = None  # noqa: N815
    hashValue: str | None = None  # noqa: N815
    saltValue: str | None = None  # noqa: N815
    spinCount: int | None = None  # noqa: N815

    def __init__(
        self,
        content: bool | None = None,
        objects: bool | None = None,
        password: str | None = None,
        algorithmName: str | None = None,  # noqa: N803
        hashValue: str | None = None,  # noqa: N803
        saltValue: str | None = None,  # noqa: N803
        spinCount: int | None = None,  # noqa: N803
        **kw: Any,
    ) -> None:
        self.content = content
        self.objects = objects
        self._password = None
        self.password = password
        self.algorithmName = algorithmName
        self.hashValue = hashValue
        self.saltValue = saltValue
        self.spinCount = spinCount
        for key, value in kw.items():
            setattr(self, key, value)

    @property
    def password(self) -> str | None:
        return self._password

    @password.setter
    def password(self, value: str | None) -> None:
        self._password = None if value is None else hash_password(str(value))

    def set_password(
        self,
        value: str | None = "",
        already_hashed: bool = False,
    ) -> None:
        self._password = value if already_hashed else hash_password(value or "")

    def __iter__(self):
        yield from _iter_openpyxl_attrs(self, self.__attrs__)

    @classmethod
    def from_tree(cls, node: Any) -> "ChartsheetProtection":
        protection = cls(
            content=_to_bool(node.get("content")),
            objects=_to_bool(node.get("objects")),
            hashValue=node.get("hashValue"),
            spinCount=_to_int(node.get("spinCount")),
            saltValue=node.get("saltValue"),
            algorithmName=node.get("algorithmName"),
        )
        password = node.get("password")
        if password is not None:
            protection.set_password(password, already_hashed=True)
        return protection

    def to_tree(self, tagname: str | None = None, idx: Any = None) -> Any:  # noqa: ARG002
        node = Element(tagname or self.tagname)
        _set_attrs(node, self)
        return node

    def _is_empty(self) -> bool:
        return not any(value is not None for _, value in self)


class Chartsheet:
    """A workbook tab that contains a single full-sheet chart."""

    tagname = "chartsheet"
    mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"
    _path = "/xl/chartsheets/sheet{0}.xml"
    __elements__ = (
        "sheetPr",
        "sheetViews",
        "sheetProtection",
        "customSheetViews",
        "pageMargins",
        "pageSetup",
        "headerFooter",
        "drawing",
        "drawingHF",
        "picture",
        "webPublishItems",
    )

    def __init__(
        self,
        parent: Any = None,
        title: str = "",
        sheetPr: Any = None,  # noqa: N803
        sheetViews: Any = None,  # noqa: N803
        sheetProtection: Any = None,  # noqa: N803
        customSheetViews: Any = None,  # noqa: N803
        pageMargins: Any = None,  # noqa: N803
        pageSetup: Any = None,  # noqa: N803
        headerFooter: Any = None,  # noqa: N803
        drawing: Any = None,
        drawingHF: Any = None,  # noqa: N803
        picture: Any = None,
        webPublishItems: Any = None,  # noqa: N803
        extLst: Any = None,  # noqa: N803
        sheet_state: str = "visible",
        **kw: Any,
    ) -> None:
        self._parent = parent
        self.title = title or "Chart"
        self.sheet_state = sheet_state
        self.sheet_properties = sheetPr if sheetPr is not None else ChartsheetProperties()
        self.sheetPr = self.sheet_properties
        self.protection = (
            sheetProtection if sheetProtection is not None else ChartsheetProtection()
        )
        self.sheetProtection = self.protection
        self.sheetViews = sheetViews
        self.customSheetViews = customSheetViews
        self.pageMargins = pageMargins
        self.pageSetup = pageSetup
        self.headerFooter = headerFooter
        self.drawing = drawing
        self.drawingHF = drawingHF
        self.picture = picture
        self.webPublishItems = webPublishItems
        self.extLst = extLst
        self._charts: list[Any] = []
        self._source_chartsheet = False
        for key, value in kw.items():
            setattr(self, key, value)

    @property
    def path(self) -> str:
        return self._path.format(getattr(self, "_id", None))

    def add_chart(self, chart: Any) -> None:
        """Attach ``chart`` to this chartsheet."""
        if chart is None:
            raise TypeError("Chartsheet.add_chart expects a chart object")
        self._charts[:] = [chart]
        if self.drawing is None:
            self.drawing = _DrawingRef("rId1")
        try:
            from wolfxl.drawing.spreadsheet_drawing import AbsoluteAnchor

            chart.anchor = AbsoluteAnchor()
        except Exception:  # noqa: BLE001 - anchor is best-effort compatibility
            pass

    def __iter__(self):
        yield from _iter_openpyxl_attrs(self)

    @classmethod
    def from_tree(cls, node: Any) -> "Chartsheet":
        kwargs: dict[str, Any] = {}
        for child in list(node):
            child_name = localname(child)
            if child_name == "sheetPr":
                kwargs["sheetPr"] = ChartsheetProperties.from_tree(child)
            elif child_name == "sheetViews":
                kwargs["sheetViews"] = ChartsheetViewList.from_tree(child)
            elif child_name == "sheetProtection":
                kwargs["sheetProtection"] = ChartsheetProtection.from_tree(child)
            elif child_name == "customSheetViews":
                from wolfxl.chartsheet.custom import CustomChartsheetViews

                kwargs["customSheetViews"] = CustomChartsheetViews.from_tree(child)
            elif child_name == "pageMargins":
                kwargs["pageMargins"] = _page_margins_from_tree(child)
            elif child_name == "pageSetup":
                kwargs["pageSetup"] = PrintPageSetup(**dict(child.attrib))
            elif child_name == "headerFooter":
                kwargs["headerFooter"] = HeaderFooter.from_tree(child)
            elif child_name == "drawing":
                kwargs["drawing"] = _DrawingRef(_relationship_id(child))
        return cls(**kwargs)

    def to_tree(self, tagname: str | None = None, idx: Any = None) -> Any:  # noqa: ARG002
        node = _chartsheet_root(tagname or self.tagname)

        if self.sheetPr is not None and not _is_empty(self.sheetPr):
            node.append(self.sheetPr.to_tree())

        sheet_views = self.sheetViews
        if sheet_views is None and self._charts:
            sheet_views = ChartsheetViewList()
        if sheet_views is not None:
            node.append(sheet_views.to_tree())

        if self.sheetProtection is not None and not _is_empty(self.sheetProtection):
            node.append(self.sheetProtection.to_tree())
        if self.customSheetViews is not None:
            node.append(self.customSheetViews.to_tree())
        if self.pageMargins is not None:
            node.append(_to_tree(self.pageMargins, "pageMargins"))
        if self.pageSetup is not None:
            node.append(_to_tree(self.pageSetup, "pageSetup"))
        if self.headerFooter is not None:
            node.append(_to_tree(self.headerFooter, "headerFooter"))

        drawing = self.drawing
        if drawing is None and self._charts:
            drawing = _DrawingRef("rId1")
        if drawing is not None:
            node.append(_drawing_to_tree(drawing))

        return node

    @property
    def _drawing(self) -> Any | None:
        """openpyxl-shaped private drawing placeholder."""
        return self._charts[0] if self._charts else None


@dataclass
class _DrawingRef:
    id: str


def _chartsheet_root(tagname: str) -> Any:
    try:
        return Element(tagname, nsmap={None: SHEET_MAIN_NS, "r": REL_NS})
    except TypeError:  # pragma: no cover - stdlib ElementTree fallback
        return Element(tagname, {"xmlns": SHEET_MAIN_NS, "xmlns:r": REL_NS})


def _set_attrs(node: Any, obj: Any) -> None:
    for name, value in obj:
        node.set(name, str(value))


def _to_bool(value: Any) -> bool | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return value
    return str(value).lower() in {"1", "true"}


def _to_int(value: Any) -> int | None:
    return int(value) if value is not None else None


def _page_margins_from_tree(node: Any) -> PageMargins:
    return PageMargins(**{name: float(value) for name, value in node.attrib.items()})


def _to_tree(obj: Any, tagname: str) -> Any:
    if hasattr(obj, "to_tree"):
        return obj.to_tree(tagname)
    node = Element(tagname)
    for name, value in dict(obj).items():
        if value is not None:
            node.set(name, str(value))
    return node


def _relationship_id(node: Any) -> str:
    return node.get(f"{{{REL_NS}}}id") or node.get("r:id") or node.get("id") or ""


def _drawing_to_tree(drawing: Any) -> Any:
    node = Element("drawing")
    rel_id = getattr(drawing, "id", None) or _relationship_id(drawing)
    if rel_id:
        node.set(f"{{{REL_NS}}}id", str(rel_id))
    return node


def _is_empty(obj: Any) -> bool:
    method = getattr(obj, "_is_empty", None)
    if callable(method):
        return bool(method())
    return not any(True for _ in obj)


__all__ = [
    "Chartsheet",
    "ChartsheetProperties",
    "ChartsheetProtection",
    "DrawingHF",
    "SpreadsheetDrawing",
    "SheetBackgroundPicture",
]

_name_fallback = _openpyxl_name_fallback(globals())


def __getattr__(name: str) -> Any:
    # Keep this re-export available without loading upstream drawing models
    # during the ordinary WolfXL import.
    if name == "SpreadsheetDrawing":
        from wolfxl.drawing.spreadsheet_drawing import SpreadsheetDrawing

        globals()[name] = SpreadsheetDrawing
        return SpreadsheetDrawing
    return _name_fallback(name)
