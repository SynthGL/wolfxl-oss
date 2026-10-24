"""Print / page-setup classes for worksheets (RFC-055 §2.1 / §2.2).

These classes back ``ws.page_setup``, ``ws.page_margins``, and the
``PrintOptions`` accessor. They are openpyxl-shaped dataclasses with
``to_rust_dict()`` helpers that emit the §10 dict contract for the
PyO3 boundary.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from wolfxl._compat import _iter_openpyxl_attrs
from wolfxl.xml.functions import Element


_VALID_ORIENTATION = ("default", "portrait", "landscape")
_VALID_CELL_COMMENTS = ("asDisplayed", "atEnd", "none")
_VALID_ERRORS = ("displayed", "blank", "dash", "NA")
_VALID_PAGE_ORDER = ("downThenOver", "overThenDown")
_ENUM_FIELDS = {
    "orientation": _VALID_ORIENTATION,
    "cellComments": _VALID_CELL_COMMENTS,
    "errors": _VALID_ERRORS,
    "pageOrder": _VALID_PAGE_ORDER,
}


def _validate_page_setup_field(name: str, value: Any) -> None:
    """Refuse a CT_PageSetup enum or scale value the schema does not allow."""
    if value is None:
        return
    allowed = _ENUM_FIELDS.get(name)
    if allowed is not None and value not in allowed:
        raise ValueError(f"{name} must be one of {allowed}, got {value!r}")
    if name == "scale" and not (10 <= value <= 400):
        raise ValueError(f"scale must be between 10 and 400, got {value}")


@dataclass
class PageSetup:
    """Page setup (CT_PageSetup, ECMA-376 §18.3.1.51).

    Attributes correspond 1:1 to OOXML ``pageSetup`` element attributes.
    All attributes are optional — ``None`` means "let Excel default it".
    """

    orientation: str | None = None
    paperSize: int | None = None  # noqa: N815 - openpyxl public API
    fitToWidth: int | None = None  # noqa: N815
    fitToHeight: int | None = None  # noqa: N815
    scale: int | None = None
    firstPageNumber: int | None = None  # noqa: N815
    horizontalDpi: int | None = None  # noqa: N815
    verticalDpi: int | None = None  # noqa: N815
    cellComments: str | None = None  # noqa: N815
    errors: str | None = None
    useFirstPageNumber: bool | None = None  # noqa: N815
    paperHeight: str | None = None  # noqa: N815
    paperWidth: str | None = None  # noqa: N815
    pageOrder: str | None = None  # noqa: N815
    usePrinterDefaults: bool | None = None  # noqa: N815
    blackAndWhite: bool | None = None  # noqa: N815
    draft: bool | None = None
    copies: int | None = None
    _fitToPage: bool | None = field(default=None, init=False, repr=False)  # noqa: N815
    _autoPageBreaks: bool | None = field(default=None, init=False, repr=False)  # noqa: N815

    def __iter__(self):
        yield from _iter_openpyxl_attrs(self)

    # openpyxl aliases (snake_case alternatives for the camelCase OOXML names)
    @property
    def paper_size(self) -> int | None:
        return self.paperSize

    @paper_size.setter
    def paper_size(self, value: int | None) -> None:
        self.paperSize = value

    @property
    def fit_to_width(self) -> int | None:
        return self.fitToWidth

    @fit_to_width.setter
    def fit_to_width(self, value: int | None) -> None:
        self.fitToWidth = value

    @property
    def fit_to_height(self) -> int | None:
        return self.fitToHeight

    @fit_to_height.setter
    def fit_to_height(self, value: int | None) -> None:
        self.fitToHeight = value

    @property
    def paper_height(self) -> str | None:
        return self.paperHeight

    @paper_height.setter
    def paper_height(self, value: str | None) -> None:
        self.paperHeight = value

    @property
    def paper_width(self) -> str | None:
        return self.paperWidth

    @paper_width.setter
    def paper_width(self, value: str | None) -> None:
        self.paperWidth = value

    @property
    def page_order(self) -> str | None:
        return self.pageOrder

    @page_order.setter
    def page_order(self, value: str | None) -> None:
        self.pageOrder = value

    @property
    def fitToPage(self) -> bool | None:  # noqa: N802
        return self._fitToPage

    @fitToPage.setter
    def fitToPage(self, value: Any) -> None:  # noqa: N802
        self._fitToPage = None if value is None else bool(value)

    @property
    def autoPageBreaks(self) -> bool | None:  # noqa: N802
        return self._autoPageBreaks

    @autoPageBreaks.setter
    def autoPageBreaks(self, value: Any) -> None:  # noqa: N802
        self._autoPageBreaks = None if value is None else bool(value)

    def __setattr__(self, name: str, value: Any) -> None:
        # Validate on every assignment, not only at construction, so
        # ``ws.page_setup.orientation = "sideways"`` refuses at the setter
        # instead of reaching the saved package.
        _validate_page_setup_field(name, value)
        super().__setattr__(name, value)

    def to_rust_dict(self) -> dict[str, Any]:
        """Emit the §10 ``page_setup`` dict for the PyO3 boundary."""
        return {
            "orientation": self.orientation,
            "paper_size": self.paperSize,
            "fit_to_width": self.fitToWidth,
            "fit_to_height": self.fitToHeight,
            "scale": self.scale,
            "first_page_number": self.firstPageNumber,
            "horizontal_dpi": self.horizontalDpi,
            "vertical_dpi": self.verticalDpi,
            "cell_comments": self.cellComments,
            "errors": self.errors,
            "use_first_page_number": self.useFirstPageNumber,
            "paper_height": self.paperHeight,
            "paper_width": self.paperWidth,
            "page_order": self.pageOrder,
            "use_printer_defaults": self.usePrinterDefaults,
            "black_and_white": self.blackAndWhite,
            "draft": self.draft,
            "copies": self.copies,
        }

    def to_tree(
        self,
        tagname: str | None = None,
        idx: int | None = None,  # noqa: ARG002 - openpyxl signature
        namespace: str | None = None,  # noqa: ARG002 - openpyxl signature
    ) -> Any:
        return _attrs_to_tree(self, tagname or "pageSetup")

    def is_default(self) -> bool:
        """True iff this PageSetup is at its construction defaults."""
        return self == PageSetup()


@dataclass(init=False)
class PageMargins:
    """Page margins in inches (CT_PageMargins, ECMA-376 §18.3.1.49)."""

    left: float = 0.75
    right: float = 0.75
    top: float = 1.0
    bottom: float = 1.0
    header: float = 0.5
    footer: float = 0.5
    _extra: dict[str, Any] = field(default_factory=dict, repr=False)

    def __init__(
        self,
        left: float = 0.75,
        right: float = 0.75,
        top: float = 1.0,
        bottom: float = 1.0,
        header: float = 0.5,
        footer: float = 0.5,
        l: float | None = None,  # noqa: E741 - openpyxl constructor alias
        r: float | None = None,
        t: float | None = None,
        b: float | None = None,
        **kw: Any,
    ) -> None:
        self.left = left if l is None else l
        self.right = right if r is None else r
        self.top = top if t is None else t
        self.bottom = bottom if b is None else b
        self.header = header
        self.footer = footer
        self._extra = dict(kw)

    def __getattr__(self, name: str) -> Any:
        try:
            extra = object.__getattribute__(self, "_extra")
        except AttributeError as exc:
            raise AttributeError(name) from exc
        if name in extra:
            return extra[name]
        raise AttributeError(name)

    def __iter__(self):
        for name in ("left", "right", "top", "bottom", "header", "footer"):
            value = getattr(self, name)
            yield name, str(int(value)) if float(value).is_integer() else str(value)

    def to_rust_dict(self) -> dict[str, float]:
        return {
            "top": float(self.top),
            "bottom": float(self.bottom),
            "left": float(self.left),
            "right": float(self.right),
            "header": float(self.header),
            "footer": float(self.footer),
        }

    def to_tree(
        self,
        tagname: str | None = None,
        idx: int | None = None,  # noqa: ARG002 - openpyxl signature
        namespace: str | None = None,  # noqa: ARG002 - openpyxl signature
    ) -> Any:
        return _attrs_to_tree(self, tagname or "pageMargins")

    def is_default(self) -> bool:
        return self == PageMargins()


@dataclass
class PrintOptions:
    """`<printOptions>` toggles (CT_PrintOptions, ECMA-376 §18.3.1.70).

    Pod 2 re-exports this under ``wolfxl.worksheet.page.PrintOptions``.
    """

    horizontalCentered: bool | None = None  # noqa: N815
    verticalCentered: bool | None = None  # noqa: N815
    headings: bool | None = None
    gridLines: bool | None = None  # noqa: N815
    gridLinesSet: bool | None = None  # noqa: N815

    def __iter__(self):
        yield from _iter_openpyxl_attrs(self)

    @property
    def horizontal_centered(self) -> bool | None:
        return self.horizontalCentered

    @horizontal_centered.setter
    def horizontal_centered(self, value: bool | None) -> None:
        self.horizontalCentered = None if value is None else bool(value)

    @property
    def vertical_centered(self) -> bool | None:
        return self.verticalCentered

    @vertical_centered.setter
    def vertical_centered(self, value: bool | None) -> None:
        self.verticalCentered = None if value is None else bool(value)

    def to_rust_dict(self) -> dict[str, Any]:
        return {
            "horizontal_centered": self.horizontalCentered,
            "vertical_centered": self.verticalCentered,
            "headings": self.headings,
            "grid_lines": self.gridLines,
            "grid_lines_set": self.gridLinesSet,
        }

    def to_tree(
        self,
        tagname: str | None = None,
        idx: int | None = None,  # noqa: ARG002 - openpyxl signature
        namespace: str | None = None,  # noqa: ARG002 - openpyxl signature
    ) -> Any:
        return _attrs_to_tree(self, tagname or "printOptions")

    def is_default(self) -> bool:
        return self == PrintOptions()


class PrintPageSetup(PageSetup):
    """Compatibility alias for openpyxl's ``PrintPageSetup``.

    openpyxl exposes this name in some module paths; the underlying
    object is the same as ``PageSetup``. Re-export for source
    compatibility — Pod 2 wires the import shim.
    """

    def __init__(
        self,
        worksheet: Any | None = None,
        orientation: str | None = None,
        paperSize: int | None = None,  # noqa: N803
        scale: int | None = None,
        fitToHeight: int | None = None,  # noqa: N803
        fitToWidth: int | None = None,  # noqa: N803
        firstPageNumber: int | None = None,  # noqa: N803
        useFirstPageNumber: bool | None = None,  # noqa: N803
        paperHeight: str | None = None,  # noqa: N803
        paperWidth: str | None = None,  # noqa: N803
        pageOrder: str | None = None,  # noqa: N803
        usePrinterDefaults: bool | None = None,  # noqa: N803
        blackAndWhite: bool | None = None,  # noqa: N803
        draft: bool | None = None,
        cellComments: str | None = None,  # noqa: N803
        errors: str | None = None,
        horizontalDpi: int | None = None,  # noqa: N803
        verticalDpi: int | None = None,  # noqa: N803
        copies: int | None = None,
        id: str | None = None,  # noqa: A002
        **kw: Any,
    ) -> None:
        self.worksheet = worksheet
        self.id = id
        self._extra = dict(kw)
        self.orientation = orientation
        self.paperSize = paperSize
        self.fitToWidth = fitToWidth
        self.fitToHeight = fitToHeight
        self.scale = scale
        self.firstPageNumber = firstPageNumber
        self.horizontalDpi = horizontalDpi
        self.verticalDpi = verticalDpi
        self.cellComments = cellComments
        self.errors = errors
        self.useFirstPageNumber = useFirstPageNumber
        self.paperHeight = paperHeight
        self.paperWidth = paperWidth
        self.pageOrder = pageOrder
        self.usePrinterDefaults = usePrinterDefaults
        self.blackAndWhite = blackAndWhite
        self.draft = draft
        self.copies = copies

    def __getattr__(self, name: str) -> Any:
        try:
            extra = object.__getattribute__(self, "_extra")
        except AttributeError as exc:
            raise AttributeError(name) from exc
        if name in extra:
            return extra[name]
        raise AttributeError(name)

    @property
    def fitToPage(self) -> bool | None:  # noqa: N802
        page_setup = _worksheet_page_setup_properties(self.worksheet)
        return None if page_setup is None else page_setup.fitToPage

    @fitToPage.setter
    def fitToPage(self, value: Any) -> None:  # noqa: N802
        page_setup = _ensure_worksheet_page_setup_properties(self.worksheet)
        page_setup.fitToPage = None if value is None else bool(value)

    @property
    def autoPageBreaks(self) -> bool | None:  # noqa: N802
        page_setup = _worksheet_page_setup_properties(self.worksheet)
        return None if page_setup is None else page_setup.autoPageBreaks

    @autoPageBreaks.setter
    def autoPageBreaks(self, value: Any) -> None:  # noqa: N802
        page_setup = _ensure_worksheet_page_setup_properties(self.worksheet)
        page_setup.autoPageBreaks = None if value is None else bool(value)


def _attrs_to_tree(obj: Any, tagname: str) -> Any:
    node = Element(tagname)
    for name, value in obj:
        node.set(name, value)
    return node


def _worksheet_page_setup_properties(worksheet: Any) -> Any:
    if worksheet is None:
        return None
    properties = getattr(worksheet, "sheet_properties", None)
    return None if properties is None else getattr(properties, "pageSetUpPr", None)


def _ensure_worksheet_page_setup_properties(worksheet: Any) -> Any:
    from wolfxl.worksheet.properties import PageSetupProperties

    page_setup = _worksheet_page_setup_properties(worksheet)
    if page_setup is not None:
        return page_setup
    properties = worksheet.sheet_properties
    properties.pageSetUpPr = PageSetupProperties()
    return properties.pageSetUpPr


__all__ = [
    "PageSetup",
    "PageMargins",
    "PrintOptions",
    "PrintPageSetup",
]
