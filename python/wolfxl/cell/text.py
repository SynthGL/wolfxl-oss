"""``openpyxl.cell.text``-shaped re-exports.

openpyxl exposes the rich-text classes at *both* ``openpyxl.cell.rich_text``
and ``openpyxl.cell.text``.  Wolfxl mirrors the former at
:mod:`wolfxl.cell.rich_text`; this module is the openpyxl-shaped alias so
``from openpyxl.cell.text import CellRichText`` swaps to wolfxl mechanically.

Pod 2 (RFC-060) — re-export shim only.
"""

from __future__ import annotations

from wolfxl._compat import _install_openpyxl_iter
from wolfxl.cell.rich_text import CellRichText, InlineFont, TextBlock
from wolfxl.descriptors import Alias, Bool, Integer, NoneSet, Sequence, Set, String, Typed
from wolfxl.descriptors.nested import NestedBool, NestedInteger, NestedString, NestedText
from wolfxl.descriptors.serialisable import Serialisable
from wolfxl.styles.fonts import Font
from wolfxl.xml.functions import Element, localname, whitespace


def _to_int(value):  # type: ignore[no-untyped-def]
    return int(value) if value is not None else None


_PHONETIC_TYPES = ("noConversion", "halfwidthKatakana", "fullwidthKatakana", "Hiragana")
_PHONETIC_ALIGNMENTS = ("noControl", "left", "center", "distributed")
_MAX_FONT_ID = 2**32 - 1


class PhoneticProperties:
    """``<phoneticPr>``: ``fontId`` (xsd:unsignedInt), ``type`` and ``alignment``.

    Each attribute is validated against the schema when assigned, so a bad value is
    refused at the assignment instead of failing when the workbook is saved. Values
    read from a stored workbook bypass validation so a foreign file reads back as written.
    """

    tagname = "phoneticPr"

    def __init__(self, fontId=None, type=None, alignment=None):  # noqa: N803, ANN001
        self.fontId = fontId
        self.type = type
        self.alignment = alignment

    @classmethod
    def _from_stored(cls, fontId=None, type=None, alignment=None):  # type: ignore[no-untyped-def]  # noqa: N803, ANN001
        """Build from a stored workbook without validating its values."""
        properties = cls()
        properties._fontId = fontId
        properties._type = type
        properties._alignment = alignment
        return properties

    @staticmethod
    def _choice(name, value, choices):  # type: ignore[no-untyped-def]
        if value is not None:
            if not isinstance(value, str):
                raise TypeError(
                    f"PhoneticProperties.{name} must be a str or None, got {type(value).__name__}"
                )
            if value not in choices:
                raise ValueError(
                    f"PhoneticProperties.{name} {value!r} is not one of {', '.join(choices)}"
                )
        return value

    @property
    def fontId(self):  # type: ignore[no-untyped-def]  # noqa: N802
        return self._fontId

    @fontId.setter
    def fontId(self, value):  # type: ignore[no-untyped-def]  # noqa: N802, ANN001
        if value is not None:
            if type(value) is not int:
                raise TypeError(
                    f"PhoneticProperties.fontId must be an int or None, got {type(value).__name__}"
                )
            if not 0 <= value <= _MAX_FONT_ID:
                raise ValueError(
                    f"PhoneticProperties.fontId {value} is outside 0..{_MAX_FONT_ID}"
                )
        self._fontId = value

    @property
    def type(self):  # type: ignore[no-untyped-def]
        return self._type

    @type.setter
    def type(self, value):  # type: ignore[no-untyped-def]  # noqa: ANN001
        self._type = self._choice("type", value, _PHONETIC_TYPES)

    @property
    def alignment(self):  # type: ignore[no-untyped-def]
        return self._alignment

    @alignment.setter
    def alignment(self, value):  # type: ignore[no-untyped-def]  # noqa: ANN001
        self._alignment = self._choice("alignment", value, _PHONETIC_ALIGNMENTS)

    @classmethod
    def from_tree(cls, node):  # type: ignore[no-untyped-def]
        return cls._from_stored(
            fontId=_to_int(node.get("fontId")),
            type=node.get("type"),
            alignment=node.get("alignment"),
        )

    def to_tree(self):  # type: ignore[no-untyped-def]
        node = Element(self.tagname)
        if self.fontId is not None:
            node.set("fontId", str(self.fontId))
        if self.type is not None:
            node.set("type", self.type)
        if self.alignment is not None:
            node.set("alignment", self.alignment)
        return node

    def __eq__(self, other):  # type: ignore[no-untyped-def]
        return isinstance(other, self.__class__) and self.__dict__ == other.__dict__


class PhoneticText:
    tagname = "rPh"

    def __init__(self, sb=None, eb=None, t=None):  # noqa: ANN001
        self.sb = sb
        self.eb = eb
        self.t = t

    @property
    def text(self):  # type: ignore[no-untyped-def]
        return self.t

    @text.setter
    def text(self, value):  # type: ignore[no-untyped-def]
        self.t = value

    @classmethod
    def from_tree(cls, node):  # type: ignore[no-untyped-def]
        text = None
        for child in node:
            if localname(child) == "t":
                text = child.text
                break
        return cls(sb=_to_int(node.get("sb")), eb=_to_int(node.get("eb")), t=text)

    def to_tree(self):  # type: ignore[no-untyped-def]
        node = Element(self.tagname)
        if self.sb is not None:
            node.set("sb", str(self.sb))
        if self.eb is not None:
            node.set("eb", str(self.eb))
        if self.t is not None:
            child = Element("t")
            child.text = self.t
            whitespace(child)
            node.append(child)
        return node

    def __eq__(self, other):  # type: ignore[no-untyped-def]
        return isinstance(other, self.__class__) and self.__dict__ == other.__dict__

class Text:
    tagname = "text"
    __attrs__ = ()
    __elements__ = ("t", "r", "rPh", "phoneticPr")

    def __init__(self, t=None, r=(), rPh=(), phoneticPr=None):  # noqa: N803, ANN001
        self.t = t
        self.r = list(r)
        self.rPh = list(rPh)
        self.phoneticPr = phoneticPr

    def __iter__(self):
        return iter(())

    @property
    def plain(self):  # type: ignore[no-untyped-def]
        return self.t

    @plain.setter
    def plain(self, value):  # type: ignore[no-untyped-def]
        self.t = value

    @property
    def formatted(self):  # type: ignore[no-untyped-def]
        return self.r

    @property
    def phonetic(self):  # type: ignore[no-untyped-def]
        return self.rPh

    @property
    def PhoneticProperties(self):  # type: ignore[no-untyped-def]  # noqa: N802
        return self.phoneticPr

    @property
    def content(self):  # type: ignore[no-untyped-def]
        snippets = []
        if self.plain is not None:
            snippets.append(self.plain)
        for block in self.formatted:
            if block.t is not None:
                snippets.append(block.t)
        return "".join(snippets)

    @classmethod
    def from_tree(cls, node):  # type: ignore[no-untyped-def]
        plain = None
        rich = []
        phonetic = []
        phonetic_pr = None
        for child in node:
            name = localname(child)
            if name == "t":
                plain = child.text
            elif name == "r":
                rich.append(RichText.from_tree(child))
            elif name == "rPh":
                phonetic.append(PhoneticText.from_tree(child))
            elif name == "phoneticPr":
                phonetic_pr = PhoneticProperties.from_tree(child)
        return cls(t=plain, r=rich, rPh=phonetic, phoneticPr=phonetic_pr)

    def to_tree(self):  # type: ignore[no-untyped-def]
        node = Element(self.tagname)
        if self.t is not None:
            child = Element("t")
            child.text = self.t
            whitespace(child)
            node.append(child)
        for rich in self.r:
            node.append(rich.to_tree())
        for phonetic in self.rPh:
            node.append(phonetic.to_tree())
        if self.phoneticPr is not None:
            node.append(self.phoneticPr.to_tree())
        return node

    def __eq__(self, other):  # type: ignore[no-untyped-def]
        return isinstance(other, self.__class__) and self.__dict__ == other.__dict__


class RichText:
    def __init__(self, rPr=None, t=None):  # type: ignore[no-untyped-def]  # noqa: N803
        self.rPr = rPr
        self.t = t

    @property
    def font(self):  # type: ignore[no-untyped-def]
        return self.rPr

    @property
    def text(self):  # type: ignore[no-untyped-def]
        return self.t

    @classmethod
    def from_tree(cls, node):  # type: ignore[no-untyped-def]
        font = None
        text = None
        for child in node:
            name = localname(child)
            if name in {"rPr", "RPrElt"}:
                font = InlineFont.from_tree(child)
            elif name == "t":
                text = child.text
        return cls(rPr=font, t=text)

    def to_tree(self):  # type: ignore[no-untyped-def]
        node = Element("RElt")
        if self.rPr is not None:
            node.append(self.rPr.to_tree("rPr"))
        if self.t is not None:
            child = Element("t")
            child.text = self.t
            whitespace(child)
            node.append(child)
        return node

    def __eq__(self, other):  # type: ignore[no-untyped-def]
        return isinstance(other, self.__class__) and self.__dict__ == other.__dict__


_install_openpyxl_iter(RichText)

__all__ = [
    "Alias",
    "Bool",
    "CellRichText",
    "Font",
    "InlineFont",
    "Integer",
    "NestedBool",
    "NestedInteger",
    "NestedString",
    "NestedText",
    "NoneSet",
    "PhoneticProperties",
    "PhoneticText",
    "RichText",
    "Sequence",
    "Serialisable",
    "Set",
    "String",
    "Text",
    "TextBlock",
    "Typed",
]
