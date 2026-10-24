"""Rich-text cell value support.

Provides three classes that mirror openpyxl's
``openpyxl.cell.rich_text`` surface:

* :class:`InlineFont` — a subset of font properties that may decorate a
  rich-text *run* (bold, italic, underline, strike, font name, size,
  color).  Matches openpyxl's ``InlineFont`` constructor keyword
  contract for the fields wolfxl actually round-trips.
* :class:`TextBlock` — a single styled run: ``font`` + ``text``.
* :class:`CellRichText` — an iterable container of ``str`` and
  ``TextBlock`` items, modeling a cell value that carries multiple
  styled runs.

These shims intentionally match openpyxl's iteration / equality /
``__str__`` semantics so user code that walks ``cell.value`` items via
``isinstance(item, (str, TextBlock))`` Just Works regardless of which
library produced the value.

Sprint Ι Pod-α (RFC pending) — closes the Phase 3 rich-text-reads
gap and the implicit T3 rich-text-write deferral.
"""

from __future__ import annotations

from collections.abc import Iterable
import copy
from dataclasses import dataclass
from decimal import Decimal
from typing import Optional, Union

from wolfxl._compat import _OpenpyxlSerialisable
from wolfxl.descriptors import Strict, String, Typed
from wolfxl.styles.colors import Color
from wolfxl.xml.functions import Element, localname, whitespace

NUMERIC_TYPES = (int, float, Decimal)
_NULLABLE_BOOL_ELEMENTS = ("strike", "outline", "shadow", "condense", "extend")


def _to_bool(value: object) -> bool:
    """Coerce a value the way openpyxl's ``Bool`` descriptor does."""
    if isinstance(value, str) and value in ("false", "f", "0"):
        return False
    return bool(value)


class Text(_OpenpyxlSerialisable):
    """Passive OOXML rich-text text node with openpyxl's ``content`` alias."""

    tagname = "text"
    __attrs__ = ("t",)

    def __init__(self, t: str | None = None, **kw: object) -> None:
        super().__init__(**kw)
        self.t = "" if t is None else t

    @property
    def content(self) -> str:
        return self.t

    @content.setter
    def content(self, value: str | None) -> None:
        self.t = "" if value is None else value

    def __iter__(self):
        if self.t:
            yield "t", self.t


@dataclass(eq=True)
class InlineFont:
    """Font properties for a single rich-text run.

    Field names mirror openpyxl's ``InlineFont`` keyword arguments
    (single-letter for the boolean attributes — ``b`` for bold, ``i``
    for italic, ``u`` for underline style, etc.).  ``b`` and ``i`` are
    always booleans (``None`` coerces to ``False``) and are only
    serialized when true; the remaining boolean flags stay ``None`` until
    set and serialize ``val="1"``/``val="0"`` like openpyxl.

    Attributes:
        rFont: Font family name.
        charset: Font character set id.
        family: Font family id.
        b: Bold flag.
        i: Italic flag.
        strike: Strikethrough flag.
        color: ARGB hex string or theme/indexed color descriptor.
        sz: Font size in points.
        u: Underline style.
        vertAlign: Vertical alignment.
        scheme: Font scheme.
    """

    __attrs__ = ()
    tagname = "RPrElt"
    __elements__ = (
        "rFont",
        "charset",
        "family",
        "b",
        "i",
        "strike",
        "outline",
        "shadow",
        "condense",
        "extend",
        "color",
        "sz",
        "u",
        "vertAlign",
        "scheme",
    )

    rFont: Optional[str] = None
    """Font family name (openpyxl alias for ``Font.name``)."""

    charset: Optional[int] = None
    family: Optional[int] = None
    b: bool = False
    """Bold."""
    i: bool = False
    """Italic."""
    strike: Optional[bool] = None
    outline: Optional[bool] = None
    shadow: Optional[bool] = None
    condense: Optional[bool] = None
    extend: Optional[bool] = None
    color: Optional[str] = None
    """ARGB hex string (e.g. ``"FFFF0000"``) or theme/indexed color
    descriptor.  Stored verbatim from XML."""
    sz: Optional[float] = None
    """Font size in points."""
    u: Optional[str] = None
    """Underline style (e.g. ``"single"``, ``"double"``).  ``True``
    coerces to ``"single"`` for openpyxl parity."""
    vertAlign: Optional[str] = None
    scheme: Optional[str] = None

    def __post_init__(self) -> None:
        # openpyxl's ``b``/``i`` descriptors are non-nullable bools, while
        # the remaining flags are nullable bools; all share string coercion.
        self.b = _to_bool(self.b)
        self.i = _to_bool(self.i)
        for attr in _NULLABLE_BOOL_ELEMENTS:
            value = getattr(self, attr)
            if value is not None:
                setattr(self, attr, _to_bool(value))
        if self.sz is not None:
            try:
                self.sz = float(self.sz)
            except (TypeError, ValueError) as exc:
                raise TypeError("expected <class 'float'>") from exc
        # openpyxl coerces the underline boolean shorthand to the
        # canonical "single" style.  Mirror that so user code that does
        # ``InlineFont(u=True)`` ends up with ``u="single"``.
        if self.u is True:
            self.u = "single"
        elif self.u is False:
            self.u = None

    def __eq__(self, other: object) -> bool:
        if type(other).__name__ != "InlineFont":
            return NotImplemented
        return all(getattr(other, attr, None) == getattr(self, attr) for attr in self.__elements__)

    def __iter__(self):
        return iter(())

    @classmethod
    def from_tree(cls, node) -> "InlineFont":  # type: ignore[no-untyped-def]
        attrs = {}
        for child in node:
            name = localname(child)
            if name == "u":
                attrs[name] = child.get("val", "single")
            elif name in {"rFont", "charset", "family", "sz", "vertAlign", "scheme"}:
                attrs[name] = child.get("val")
            elif name in {"b", "i", *_NULLABLE_BOOL_ELEMENTS}:
                attrs[name] = child.get("val", True)
            elif name == "color":
                attrs[name] = Color.from_tree(child)
        return cls(**attrs)

    def to_tree(self, tagname: str | None = None):
        node = Element(tagname or self.tagname)
        for attr in self.__elements__:
            value = getattr(self, attr)
            if value is None or (attr in ("b", "i") and not value):
                continue
            if attr == "color":
                color = value if isinstance(value, Color) else Color(rgb=value)
                child = Element("color")
                if color.type == "rgb" and color.rgb is not None:
                    rgb = color.rgb
                    child.set("rgb", rgb if len(rgb) == 8 else f"00{rgb}")
                elif color.type == "indexed" and color.indexed is not None:
                    child.set("indexed", str(color.indexed))
                elif color.type == "theme" and color.theme is not None:
                    child.set("theme", str(color.theme))
                elif color.type == "auto" and color.auto is not None:
                    child.set("auto", "1" if color.auto else "0")
                if color.tint:
                    child.set("tint", str(color.tint))
                node.append(child)
                continue
            child = Element(attr)
            if isinstance(value, bool):
                child.set("val", "1" if value else "0")
            else:
                child.set("val", str(value).rstrip("0").rstrip(".") if attr == "sz" else str(value))
            node.append(child)
        return node


@dataclass
class TextBlock:
    """A styled run inside a :class:`CellRichText`."""

    font: InlineFont
    text: str

    def __init__(self, font: InlineFont, text: str) -> None:
        """Create a styled rich-text run.

        Args:
            font: Inline font applied to this run.
            text: Plain text for this run.
        """
        # openpyxl positions ``font`` first to match the ``<r><rPr>...</rPr><t>...</t></r>``
        # XML reading order.  Mirror that signature.
        if not isinstance(font, InlineFont):
            raise TypeError(
                f"{type(self)}.font should be {InlineFont} but value is {type(font)}"
            )
        if not isinstance(text, str):
            raise TypeError(
                f"{type(self)}.text should be {str} but value is {type(text)}"
            )
        self.font = font
        self.text = text

    def __eq__(self, other: object) -> bool:
        if type(other).__name__ != "TextBlock":
            return NotImplemented
        return self.font == getattr(other, "font", None) and self.text == getattr(other, "text", None)

    def __hash__(self) -> int:  # pragma: no cover - dataclasses default
        return hash((self.text,))

    def __str__(self) -> str:
        # openpyxl's ``str(TextBlock)`` returns the plain text — this
        # makes ``"".join(map(str, cell_rich_text))`` produce the
        # flattened representation.
        return self.text

    def __repr__(self) -> str:
        font = self.font if self.font != InlineFont() else "default"
        return f"{self.__class__.__name__} text={self.text}, font={font}"

    def to_tree(self):
        el = Element("r")
        el.append(self.font.to_tree(tagname="rPr"))
        t = Element("t")
        t.text = self.text
        whitespace(t)
        el.append(t)
        return el


class CellRichText(list):
    """Sequence of ``str`` and :class:`TextBlock` runs.

    Subclassing ``list`` so existing user code that iterates,
    indexes, slices, or appends to a rich-text value Just Works
    (matches openpyxl's design — ``CellRichText`` is also a list
    subclass there).
    """

    def __init__(self, *args) -> None:  # type: ignore[no-untyped-def]
        """Create a rich-text run list.

        Args:
            args: A single run, an iterable of runs, or variadic runs.
        """
        if len(args) == 1:
            items = args[0]
            if isinstance(items, (list, tuple)):
                self._check_rich_text(items)
            else:
                self._check_element(items)
                items = [items]
        else:
            items = args
            self._check_rich_text(items)
        super().__init__(items)

    @classmethod
    def _check_element(cls, value):  # type: ignore[no-untyped-def]
        if not isinstance(value, (str, TextBlock, NUMERIC_TYPES)):
            raise TypeError(f"Illegal CellRichText element {value}")

    @classmethod
    def _check_rich_text(cls, rich_text):  # type: ignore[no-untyped-def]
        for item in rich_text:
            cls._check_element(item)

    @classmethod
    def from_tree(cls, node):  # type: ignore[no-untyped-def]
        items = []
        for child in node:
            tag = localname(child)
            if tag == "t":
                if child.text:
                    items.append(child.text.replace("x005F_", ""))
            elif tag == "r":
                text = ""
                font = None
                for run_child in child:
                    name = localname(run_child)
                    if name == "t":
                        text = (run_child.text or "").replace("x005F_", "")
                    elif name == "rPr":
                        font = InlineFont.from_tree(run_child)
                items.append(TextBlock(font, text) if font else text)
        return cls(items)

    def _opt(self):
        last_t = None
        optimized = CellRichText(tuple())
        for item in self:
            if isinstance(item, str):
                if not item:
                    continue
            elif not item.text:
                continue
            if type(last_t) == type(item):
                if isinstance(item, str):
                    last_t += item
                    continue
                if last_t.font == item.font:
                    last_t.text += item.text
                    continue
            if last_t:
                optimized.append(last_t)
            last_t = item
        if last_t:
            optimized.append(last_t)
        super().__setitem__(slice(None), optimized)
        return self

    def append(self, value: Union[str, TextBlock]) -> None:  # type: ignore[override]
        """Append one rich-text run.

        Args:
            value: Plain string run or styled ``TextBlock``.

        Raises:
            TypeError: If ``value`` is not a string or ``TextBlock``.
        """
        self._check_element(value)
        super().append(value)

    def extend(self, values):  # type: ignore[no-untyped-def,override]
        self._check_rich_text(values)
        super().extend(values)

    def __iadd__(self, other: Iterable[Union[str, TextBlock]]) -> "CellRichText":  # type: ignore[override]
        self._check_rich_text(other)
        super().__iadd__([copy.copy(item) for item in list(other)])
        return self._opt()

    def __add__(self, other: Iterable[Union[str, TextBlock]]) -> "CellRichText":  # type: ignore[override]
        return CellRichText([copy.copy(item) for item in list(self) + list(other)])._opt()

    def __setitem__(self, index, value):  # type: ignore[no-untyped-def,override]
        self._check_element(value)
        super().__setitem__(index, value)
        self._opt()

    def __str__(self) -> str:
        # Flatten to plain text — same semantics as openpyxl's
        # ``CellRichText.__str__``.
        return "".join(str(item) for item in self)

    def __repr__(self) -> str:
        inner = ", ".join(repr(item) for item in self)
        return f"CellRichText([{inner}])"

    def as_list(self) -> list[Union[str, TextBlock]]:
        """Materialize a plain ``list`` copy of the runs.

        Convenience for callers that want to avoid handing the
        underlying mutable list to downstream code.

        Returns:
            A shallow list copy of the current rich-text runs.
        """
        return [str(item) for item in self]

    def to_tree(self):
        container = Element("is")
        for obj in self:
            if isinstance(obj, TextBlock):
                container.append(obj.to_tree())
            else:
                el = Element("r")
                t = Element("t")
                t.text = str(obj)
                whitespace(t)
                el.append(t)
                container.append(el)
        return container


__all__ = [
    "CellRichText",
    "Element",
    "InlineFont",
    "NUMERIC_TYPES",
    "Strict",
    "String",
    "Text",
    "TextBlock",
    "Typed",
    "copy",
    "whitespace",
]
