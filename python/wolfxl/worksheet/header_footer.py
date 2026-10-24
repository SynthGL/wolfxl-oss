"""Header / footer classes for worksheets (RFC-055 §2.3, G09).

Backs ``ws.header_footer``. Includes the OOXML format-code grammar
validator (``&L`` / ``&C`` / ``&R`` / ``&P`` / ``&N`` / ``&D`` /
``&T`` / ``&F`` / ``&A`` / ``&Z`` / ``&K{RRGGBB}`` / ``&"font,style"``
/ ``&NN`` / ``&B`` / ``&I`` / ``&U`` / ``&S`` / ``&X`` / ``&Y`` /
``&&``).

Also exposes openpyxl-shaped ``_HeaderFooterPart`` access on each
segment so ``ws.oddHeader.center.text``, ``.font``, ``.size``, and
``.color`` round-trip through the inline mini-format codes.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any

from wolfxl._compat import _iter_openpyxl_attrs, _openpyxl_name_fallback
from wolfxl.utils.escape import unescape
from wolfxl.xml.functions import Element


# Recognized single-letter codes (case-sensitive in OOXML).
_FORMAT_CODES_SINGLE = set("LCRPNDTFAZGBIUSXY")

# Recognized format-code grammar regex - used for validation only;
# unknown ``&X`` codes pass through unchanged for forward compat.
_FORMAT_CODE_RE = re.compile(
    r"""
    &&                                  # literal &
    | &"[^"]*"                          # &"font,style"
    | &K[0-9A-Fa-f]{6}                  # &Kffrrggbb (case-insensitive hex)
    | &K\{[0-9A-Fa-f]{6}\}              # &K{RRGGBB} (alternate form)
    | &\d+                              # &NN font size
    | &[A-Za-z]                         # single-letter code
    """,
    re.VERBOSE,
)


# Mini-format inline regex for prefix codes. Mirrors openpyxl's
# FORMAT_REGEX so we parse ``&"Arial,Bold"&14&KFF0000Title`` into
# (font, color, size, text).
_PREFIX_FORMAT_RE = re.compile(r'&"(?P<font>[^"]+)"|&K(?P<color>[A-Fa-f0-9]{6})|&(?P<size>\d+\s?)')
FORMAT_REGEX = _PREFIX_FORMAT_RE
FONT_PATTERN = r'&"(?P<font>[^"]+)"'
COLOR_PATTERN = r"&K(?P<color>[A-Fa-f0-9]{6})"
SIZE_REGEX = r"&(?P<size>\d+\s?)"

# Section split for ``&L...&C...&R...`` form.
_ITEM_SPLIT_RE = re.compile(
    r"""
    (&L(?P<left>.+?))?
    (&C(?P<center>.+?))?
    (&R(?P<right>.+?))?
    $
    """,
    re.VERBOSE | re.DOTALL,
)

_EXPANDED_CONTROL_CODES = {
    "&[Tab]": "&A",
    "&[Pages]": "&N",
    "&[Date]": "&D",
    "&[Path]": "&Z",
    "&[Page]": "&P",
    "&[Time]": "&T",
    "&[File]": "&F",
    "&[Picture]": "&G",
}
_EXPANDED_CONTROL_RE = re.compile(
    "|".join(f"({re.escape(key)})" for key in _EXPANDED_CONTROL_CODES)
)


def _split_string(text: str) -> dict[str, str | None]:
    """Split a combined ``&L...&C...&R...`` string into position parts."""
    match = _ITEM_SPLIT_RE.match(text)
    if match is None:
        return {"left": "", "right": "", "center": ""}
    return match.groupdict()


def validate_header_footer_format(text: str) -> bool:
    """Return True iff ``text`` parses as a valid OOXML header/footer string.

    "Valid" means: every ``&`` is followed by a recognized code or is
    a ``&&`` literal escape. Unknown single-letter codes (e.g. invented
    by a later spec revision) are treated as opaque - they pass through
    unchanged for forward compat.

    The function never raises; it returns False for clearly malformed
    inputs (a trailing bare ``&`` with no code character, an unclosed
    quoted font block).
    """
    if text is None:
        return True
    i = 0
    n = len(text)
    while i < n:
        ch = text[i]
        if ch != "&":
            i += 1
            continue
        if i + 1 >= n:
            return False
        nxt = text[i + 1]
        if nxt == "&":
            i += 2
            continue
        if nxt == '"':
            close = text.find('"', i + 2)
            if close == -1:
                return False
            i = close + 1
            continue
        if nxt in ("K", "k"):
            if i + 2 >= n:
                return False
            if text[i + 2] == "{":
                if i + 9 >= n or text[i + 9] != "}":
                    return False
                if not all(c in "0123456789abcdefABCDEF" for c in text[i + 3 : i + 9]):
                    return False
                i += 10
                continue
            if i + 8 > n:
                return False
            if not all(c in "0123456789abcdefABCDEF" for c in text[i + 2 : i + 8]):
                return False
            i += 8
            continue
        if nxt.isdigit():
            j = i + 1
            while j < n and text[j].isdigit():
                j += 1
            i = j
            continue
        if nxt.isalpha():
            i += 2
            continue
        return False
    return True


_RGB_RE = re.compile(r"^[A-Fa-f0-9]{6}$")


class _HeaderFooterPart:
    """Openpyxl-shaped left/center/right segment.

    Holds an inline mini-format payload for one alignment slot. The
    ``text`` field is the user-visible plain text; ``font`` /
    ``size`` / ``color`` are leading format prefixes. Toggle codes
    (``&B``, ``&I``, ``&U``, ``&S``, ``&X``, ``&Y``) are not lifted
    onto attributes - they survive opaquely inside ``text``.

    Comparison against a plain ``str`` matches the composed
    OOXML mini-format string (i.e. what the writer emits). ``bool``
    returns True iff there is any text content.
    """

    __slots__ = ("_text", "_font", "_size", "_color")

    def __init__(
        self,
        text: str | None = None,
        font: str | None = None,
        size: int | None = None,
        color: str | None = None,
    ) -> None:
        self._text: str | None = None
        self._font: str | None = None
        self._size: int | None = None
        self._color: str | None = None
        self.text = text
        self.font = font
        self.size = size
        self.color = color

    @property
    def text(self) -> str | None:
        return self._text

    @text.setter
    def text(self, value: Any) -> None:
        if value is None:
            self._text = None
            return
        if not isinstance(value, str):
            raise TypeError(
                f"_HeaderFooterPart.text must be str or None; got {type(value).__name__}"
            )
        self._text = value

    @property
    def font(self) -> str | None:
        return self._font

    @font.setter
    def font(self, value: Any) -> None:
        if value is None:
            self._font = None
            return
        if not isinstance(value, str):
            raise TypeError(
                f"_HeaderFooterPart.font must be str or None; got {type(value).__name__}"
            )
        self._font = value

    @property
    def size(self) -> int | None:
        return self._size

    @size.setter
    def size(self, value: Any) -> None:
        if value is None:
            self._size = None
            return
        try:
            self._size = int(value)
        except (TypeError, ValueError) as exc:
            raise TypeError(f"_HeaderFooterPart.size must be int or None; got {value!r}") from exc

    @property
    def color(self) -> str | None:
        return self._color

    @color.setter
    def color(self, value: Any) -> None:
        if value is None:
            self._color = None
            return
        if not isinstance(value, str) or not _RGB_RE.match(value):
            raise ValueError(f"_HeaderFooterPart.color must match RRGGBB hex; got {value!r}")
        self._color = value.upper()

    def is_empty(self) -> bool:
        return (
            self._text is None and self._font is None and self._size is None and self._color is None
        )

    def to_format_string(self) -> str | None:
        """Compose this part to its inline mini-format string.

        Returns None when the part is empty so the parent
        ``HeaderFooterItem`` can treat the slot as absent (matching
        openpyxl's ``&L``/``&C``/``&R`` suppression).
        """
        if self.is_empty():
            return None
        parts: list[str] = []
        if self._font is not None:
            parts.append(f'&"{self._font}"')
        if self._size is not None:
            parts.append(f"&{self._size} ")
        if self._color is not None:
            parts.append(f"&K{self._color}")
        if self._text is not None:
            parts.append(self._text)
        return "".join(parts)

    @classmethod
    def from_format_string(cls, text: str | None) -> "_HeaderFooterPart":
        """Parse an inline mini-format string into a part.

        Mirrors openpyxl's ``_HeaderFooterPart.from_str`` behavior:
        leading ``&"font"``, ``&KRRGGBB``, ``&NN`` prefixes are
        stripped onto attributes; the remaining text is the
        user-visible payload (which may still contain inline toggle
        codes like ``&B`` or page-number macros).
        """
        if text is None:
            return cls()
        font: str | None = None
        size: int | None = None
        color: str | None = None
        for match in _PREFIX_FORMAT_RE.finditer(text):
            if match.group("font") is not None and font is None:
                font = match.group("font")
            elif match.group("color") is not None and color is None:
                color = match.group("color").upper()
            elif match.group("size") is not None and size is None:
                size = int(match.group("size").strip())
        residual = _PREFIX_FORMAT_RE.sub("", text)
        return cls(
            text=residual if residual != "" else (None if (font or size or color) else ""),
            font=font,
            size=size,
            color=color,
        )

    from_str = from_format_string

    def __bool__(self) -> bool:
        return bool(self._text)

    def __eq__(self, other: object) -> bool:
        if isinstance(other, _HeaderFooterPart):
            return (
                self._text == other._text
                and self._font == other._font
                and self._size == other._size
                and self._color == other._color
            )
        if isinstance(other, str):
            return (self.to_format_string() or "") == other
        if other is None:
            return self.is_empty()
        return NotImplemented

    def __ne__(self, other: object) -> bool:
        result = self.__eq__(other)
        if result is NotImplemented:
            return result
        return not result

    def __hash__(self) -> int:
        return hash((self._text, self._font, self._size, self._color))

    def __str__(self) -> str:
        return self.to_format_string() or ""

    def __repr__(self) -> str:
        return (
            f"_HeaderFooterPart(text={self._text!r}, font={self._font!r}, "
            f"size={self._size!r}, color={self._color!r})"
        )


def _coerce_part(value: Any) -> _HeaderFooterPart:
    """Accept ``None`` / ``str`` / ``_HeaderFooterPart`` and return a part."""
    if value is None:
        return _HeaderFooterPart()
    if isinstance(value, _HeaderFooterPart):
        return value
    if isinstance(value, str):
        if not validate_header_footer_format(value):
            raise ValueError(f"invalid format code in {value!r}")
        return _HeaderFooterPart.from_format_string(value)
    raise TypeError(
        f"HeaderFooterItem segment must be None, str, or _HeaderFooterPart; got {type(value).__name__}"
    )


class HeaderFooterItem:
    """One header or footer (left / center / right segments).

    Each slot is a :class:`_HeaderFooterPart` instance (always non-None,
    matching openpyxl). Assigning a plain ``str`` parses the inline
    mini-format and replaces the slot in place; assigning ``None``
    clears the slot. Reading the slot returns the
    :class:`_HeaderFooterPart` object, which compares equal to its
    composed format string.
    """

    __slots__ = ("_left", "_center", "_right")

    def __init__(
        self,
        left: Any = None,
        center: Any = None,
        right: Any = None,
    ) -> None:
        self._left = _coerce_part(left)
        self._center = _coerce_part(center)
        self._right = _coerce_part(right)

    @property
    def left(self) -> _HeaderFooterPart:
        return self._left

    @left.setter
    def left(self, value: Any) -> None:
        self._left = _coerce_part(value)

    @property
    def center(self) -> _HeaderFooterPart:
        return self._center

    @center.setter
    def center(self, value: Any) -> None:
        self._center = _coerce_part(value)

    @property
    def centre(self) -> _HeaderFooterPart:
        return self._center

    @centre.setter
    def centre(self, value: Any) -> None:
        self._center = _coerce_part(value)

    @property
    def right(self) -> _HeaderFooterPart:
        return self._right

    @right.setter
    def right(self, value: Any) -> None:
        self._right = _coerce_part(value)

    @property
    def text(self) -> str:
        """Compose the segments back into the OOXML text form.

        Used by the emitter and by tests that compare against
        openpyxl's serialization.
        """
        return str(self)

    def __str__(self) -> str:
        parts: list[str] = []
        for marker, part in (("L", self._left), ("C", self._center), ("R", self._right)):
            if part.text is not None:
                payload = part.to_format_string() or ""
                parts.append("&" + marker + payload)
        text = "".join(parts)
        return _EXPANDED_CONTROL_RE.sub(lambda match: _EXPANDED_CONTROL_CODES[match.group(0)], text)

    def is_empty(self) -> bool:
        return self._left.is_empty() and self._center.is_empty() and self._right.is_empty()

    def to_rust_dict(self) -> dict[str, Any] | None:
        if self.is_empty():
            return None
        return {
            "left": self._left.to_format_string(),
            "center": self._center.to_format_string(),
            "right": self._right.to_format_string(),
        }

    def to_tree(self, tagname: str) -> Any:
        node = Element(tagname)
        node.text = str(self)
        return node

    @classmethod
    def from_str(cls, text: str | None) -> "HeaderFooterItem":
        """Parse a full ``&L...&C...&R...`` string into segments.

        Mirrors openpyxl's ``HeaderFooterItem`` reconstruction so
        incoming OOXML can be reflected back through the Python API
        with full attribute fidelity.
        """
        if not text:
            return cls()
        match = _ITEM_SPLIT_RE.match(text)
        if match is None:
            return cls()
        return cls(
            left=match.group("left"),
            center=match.group("center"),
            right=match.group("right"),
        )

    @classmethod
    def from_tree(cls, node: Any) -> "HeaderFooterItem":
        if not getattr(node, "text", None):
            return cls()
        parts = _split_string(unescape(node.text))
        return cls(
            left=(
                _HeaderFooterPart.from_str(parts["left"]) if parts.get("left") is not None else None
            ),
            center=(
                _HeaderFooterPart.from_str(parts["center"])
                if parts.get("center") is not None
                else None
            ),
            right=(
                _HeaderFooterPart.from_str(parts["right"])
                if parts.get("right") is not None
                else None
            ),
        )

    def __bool__(self) -> bool:
        return not self.is_empty()

    def __eq__(self, other: object) -> bool:
        if isinstance(other, HeaderFooterItem):
            return (
                self._left == other._left
                and self._center == other._center
                and self._right == other._right
            )
        return NotImplemented

    def __repr__(self) -> str:
        return (
            f"HeaderFooterItem(left={self._left!r}, center={self._center!r}, right={self._right!r})"
        )


@dataclass(init=False)
class HeaderFooter:
    """Complete CT_HeaderFooter (ECMA-376 §18.3.1.36)."""

    odd_header: HeaderFooterItem = field(default_factory=HeaderFooterItem)
    odd_footer: HeaderFooterItem = field(default_factory=HeaderFooterItem)
    even_header: HeaderFooterItem = field(default_factory=HeaderFooterItem)
    even_footer: HeaderFooterItem = field(default_factory=HeaderFooterItem)
    first_header: HeaderFooterItem = field(default_factory=HeaderFooterItem)
    first_footer: HeaderFooterItem = field(default_factory=HeaderFooterItem)
    different_odd_even: bool = False
    different_first: bool = False
    scale_with_doc: bool = True
    align_with_margins: bool = True

    def __init__(
        self,
        odd_header: HeaderFooterItem | str | None = None,
        odd_footer: HeaderFooterItem | str | None = None,
        even_header: HeaderFooterItem | str | None = None,
        even_footer: HeaderFooterItem | str | None = None,
        first_header: HeaderFooterItem | str | None = None,
        first_footer: HeaderFooterItem | str | None = None,
        different_odd_even: bool | None = None,
        different_first: bool | None = None,
        scale_with_doc: bool | None = None,
        align_with_margins: bool | None = None,
        differentOddEven: bool | None = None,  # noqa: N803
        differentFirst: bool | None = None,  # noqa: N803
        scaleWithDoc: bool | None = None,  # noqa: N803
        alignWithMargins: bool | None = None,  # noqa: N803
        oddHeader: HeaderFooterItem | str | None = None,  # noqa: N803
        oddFooter: HeaderFooterItem | str | None = None,  # noqa: N803
        evenHeader: HeaderFooterItem | str | None = None,  # noqa: N803
        evenFooter: HeaderFooterItem | str | None = None,  # noqa: N803
        firstHeader: HeaderFooterItem | str | None = None,  # noqa: N803
        firstFooter: HeaderFooterItem | str | None = None,  # noqa: N803
    ) -> None:
        self.odd_header = _coerce_item(oddHeader if oddHeader is not None else odd_header)
        self.odd_footer = _coerce_item(oddFooter if oddFooter is not None else odd_footer)
        self.even_header = _coerce_item(evenHeader if evenHeader is not None else even_header)
        self.even_footer = _coerce_item(evenFooter if evenFooter is not None else even_footer)
        self.first_header = _coerce_item(firstHeader if firstHeader is not None else first_header)
        self.first_footer = _coerce_item(firstFooter if firstFooter is not None else first_footer)
        self.different_odd_even = (
            differentOddEven if differentOddEven is not None else different_odd_even
        )
        self.different_first = differentFirst if differentFirst is not None else different_first
        self.scale_with_doc = scaleWithDoc if scaleWithDoc is not None else scale_with_doc
        self.align_with_margins = (
            alignWithMargins if alignWithMargins is not None else align_with_margins
        )

    def __setattr__(self, name: str, value: Any) -> None:
        if name in _FLAG_DEFAULTS:
            value = _coerce_flag(name, value)
        super().__setattr__(name, value)

    @property
    def oddHeader(self) -> HeaderFooterItem:  # noqa: N802
        return self.odd_header

    @oddHeader.setter
    def oddHeader(self, value: HeaderFooterItem) -> None:  # noqa: N802
        self.odd_header = _coerce_item(value)

    @property
    def oddFooter(self) -> HeaderFooterItem:  # noqa: N802
        return self.odd_footer

    @oddFooter.setter
    def oddFooter(self, value: HeaderFooterItem) -> None:  # noqa: N802
        self.odd_footer = _coerce_item(value)

    @property
    def evenHeader(self) -> HeaderFooterItem:  # noqa: N802
        return self.even_header

    @evenHeader.setter
    def evenHeader(self, value: HeaderFooterItem) -> None:  # noqa: N802
        self.even_header = _coerce_item(value)

    @property
    def evenFooter(self) -> HeaderFooterItem:  # noqa: N802
        return self.even_footer

    @evenFooter.setter
    def evenFooter(self, value: HeaderFooterItem) -> None:  # noqa: N802
        self.even_footer = _coerce_item(value)

    @property
    def firstHeader(self) -> HeaderFooterItem:  # noqa: N802
        return self.first_header

    @firstHeader.setter
    def firstHeader(self, value: HeaderFooterItem) -> None:  # noqa: N802
        self.first_header = _coerce_item(value)

    @property
    def firstFooter(self) -> HeaderFooterItem:  # noqa: N802
        return self.first_footer

    @firstFooter.setter
    def firstFooter(self, value: HeaderFooterItem) -> None:  # noqa: N802
        self.first_footer = _coerce_item(value)

    def is_default(self) -> bool:
        return (
            self.odd_header.is_empty()
            and self.odd_footer.is_empty()
            and self.even_header.is_empty()
            and self.even_footer.is_empty()
            and self.first_header.is_empty()
            and self.first_footer.is_empty()
            and not self.different_odd_even
            and not self.different_first
            and self.scale_with_doc
            and self.align_with_margins
        )

    def __iter__(self):
        yield from _iter_openpyxl_attrs(
            self,
            (
                "differentOddEven",
                "differentFirst",
                "scaleWithDoc",
                "alignWithMargins",
            ),
        )

    def to_rust_dict(self) -> dict[str, Any]:
        return {
            "odd_header": self.odd_header.to_rust_dict(),
            "odd_footer": self.odd_footer.to_rust_dict(),
            "even_header": self.even_header.to_rust_dict(),
            "even_footer": self.even_footer.to_rust_dict(),
            "first_header": self.first_header.to_rust_dict(),
            "first_footer": self.first_footer.to_rust_dict(),
            "different_odd_even": bool(self.different_odd_even),
            "different_first": bool(self.different_first),
            "scale_with_doc": bool(self.scale_with_doc),
            "align_with_margins": bool(self.align_with_margins),
        }

    def __bool__(self) -> bool:
        return any(
            [
                self.odd_header,
                self.odd_footer,
                self.even_header,
                self.even_footer,
                self.first_header,
                self.first_footer,
                self.different_odd_even,
                self.different_first,
                not self.scale_with_doc,
                not self.align_with_margins,
            ]
        )

    def to_tree(self, tagname: str | None = None) -> Any:
        attrs = {}
        if self.different_odd_even:
            attrs["differentOddEven"] = "1"
        if self.different_first:
            attrs["differentFirst"] = "1"
        if not self.scale_with_doc:
            attrs["scaleWithDoc"] = "0"
        if not self.align_with_margins:
            attrs["alignWithMargins"] = "0"
        node = Element(tagname or "headerFooter", attrs)
        for child_tag, item in (
            ("oddHeader", self.odd_header),
            ("oddFooter", self.odd_footer),
            ("evenHeader", self.even_header),
            ("evenFooter", self.even_footer),
            ("firstHeader", self.first_header),
            ("firstFooter", self.first_footer),
        ):
            node.append(item.to_tree(child_tag))
        return node

    @classmethod
    def from_tree(cls, node: Any) -> "HeaderFooter":
        values: dict[str, Any] = {}
        attr_map = {
            "differentOddEven": "differentOddEven",
            "differentFirst": "differentFirst",
            "scaleWithDoc": "scaleWithDoc",
            "alignWithMargins": "alignWithMargins",
        }
        for xml_name, init_name in attr_map.items():
            if xml_name in node.attrib:
                values[init_name] = _xml_bool(node.attrib[xml_name])
        child_map = {
            "oddHeader": "oddHeader",
            "oddFooter": "oddFooter",
            "evenHeader": "evenHeader",
            "evenFooter": "evenFooter",
            "firstHeader": "firstHeader",
            "firstFooter": "firstFooter",
        }
        for child in node:
            tag = getattr(child, "tag", "")
            local = tag.rsplit("}", 1)[-1]
            if local in child_map:
                values[child_map[local]] = HeaderFooterItem.from_tree(child)
        return cls(**values)


# CT_HeaderFooter boolean attributes and their schema defaults.
_FLAG_DEFAULTS = {
    "different_odd_even": False,
    "different_first": False,
    "scale_with_doc": True,
    "align_with_margins": True,
}


def _coerce_flag(name: str, value: Any) -> bool:
    """Return a header/footer flag as a bool; ``None`` restores the default."""
    if value is None:
        return _FLAG_DEFAULTS[name]
    if isinstance(value, bool):
        return value
    if isinstance(value, int) and value in (0, 1):
        return bool(value)
    raise TypeError(f"HeaderFooter.{name} must be a bool or None; got {value!r}")


def _coerce_item(value: Any) -> HeaderFooterItem:
    if value is None:
        return HeaderFooterItem()
    if isinstance(value, HeaderFooterItem):
        return value
    if isinstance(value, str):
        return HeaderFooterItem.from_str(value)
    raise TypeError(
        f"HeaderFooter slot must be None, str, or HeaderFooterItem; got {type(value).__name__}"
    )


def _xml_bool(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    return str(value).lower() in {"1", "true", "on"}


__all__ = [
    "COLOR_PATTERN",
    "FONT_PATTERN",
    "FORMAT_REGEX",
    "HeaderFooter",
    "HeaderFooterItem",
    "SIZE_REGEX",
    "_HeaderFooterPart",
    "_split_string",
    "validate_header_footer_format",
]

__getattr__ = _openpyxl_name_fallback(globals())
