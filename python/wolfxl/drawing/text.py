"""Drawing text compatibility containers and the shape paragraph model."""

from __future__ import annotations
import html as _html
import re as _re
from html.parser import HTMLParser as _HTMLParser
import math as _math
import sys as _sys
from collections.abc import Iterable as _Iterable
from dataclasses import dataclass as _dataclass
from dataclasses import field as _field
from typing import TYPE_CHECKING as _TYPE_CHECKING
from typing import Any as _Any
from typing import Union as _Union

from wolfxl._compat import (
    _OpenpyxlSerialisable,
    _make_serialisable,
    _resolve_openpyxl_class,
)
from wolfxl.drawing.colors import ColorChoiceDescriptor, OfficeArtExtensionList
from wolfxl.drawing.fill import (
    Blip,
    BlipFillProperties,
    GradientFillProperties,
    PatternFillProperties,
)
from wolfxl.drawing.geometry import GeomGuide, GeomGuideList, PresetTextShape, Scene3D
from wolfxl.drawing.line import LineProperties
from wolfxl.xml.constants import DRAWING_NS

if _TYPE_CHECKING:
    from wolfxl.drawing.drawing import TextRun as _TextRun


def _text_runs_payload(runs: _Iterable[_TextRun]) -> list[dict[str, _Any]] | None:
    payload = [
        {
            "text": run.text,
            "family": run.family,
            "size": None if run.size is None else round(run.size * 100),
            "bold": run.bold,
            "italic": run.italic,
            "color": run.color,
        }
        for run in runs
    ]
    return payload or None


def _openpyxl_class(name: str) -> type:
    return _resolve_openpyxl_class(__name__, name) or _make_serialisable(name)


for _name in (
    "AutonumberBullet",
    "CharacterProperties",
    "EmbeddedWAVAudioFile",
    "Font",
    "LineBreak",
    "ListStyle",
    "Paragraph",
    "ParagraphProperties",
    "RegularTextRun",
    "RichTextProperties",
    "Spacing",
    "TabStop",
    "TabStopList",
    "TextField",
    "TextNormalAutofit",
):
    globals()[_name] = _openpyxl_class(_name)

Run = RegularTextRun

Alias = Bool = Color = Coordinate = EffectContainer = EffectList = EmptyTag = HexBinary = (
    Integer
) = MinMax = NestedBool = NestedInteger = NestedText = NestedValue = NoneSet = Relation = (
    Sequence
) = Serialisable = Set = String = Typed = _OpenpyxlSerialisable
Hyperlink = _openpyxl_class("Hyperlink")


# --- Shape paragraph model -------------------------------------------------
#
# Public values use points and friendly names; ``*_payload`` methods produce
# the DrawingML tokens and integer units consumed by the native writer.

_EMU_PER_POINT = 12_700
_MAX_COORDINATE_EMU = 51_206_400

PARAGRAPH_ALIGNMENTS = {
    "left": "l",
    "center": "ctr",
    "right": "r",
    "justify": "just",
    "justify_low": "justLow",
    "distributed": "dist",
    "thai_distributed": "thaiDist",
}
FONT_ALIGNMENTS = {
    "auto": "auto",
    "top": "t",
    "center": "ctr",
    "baseline": "base",
    "bottom": "b",
}
TAB_ALIGNMENTS = {"left": "l", "center": "ctr", "right": "r", "decimal": "dec"}
VERTICAL_ANCHORS = {
    "top": "t",
    "middle": "ctr",
    "bottom": "b",
    "justified": "just",
    "distributed": "dist",
}
VERTICAL_TEXT_TYPES = {
    "horizontal": "horz",
    "vertical": "vert",
    "vertical270": "vert270",
    "word_art_vertical": "wordArtVert",
    "east_asian_vertical": "eaVert",
    "mongolian_vertical": "mongolianVert",
    "word_art_vertical_rtl": "wordArtVertRtl",
}
AUTOFIT_TYPES = ("none", "shrink", "shape")
HORIZONTAL_OVERFLOW_TYPES = ("overflow", "clip")
VERTICAL_OVERFLOW_TYPES = ("overflow", "ellipsis", "clip")
#: ``ST_TextAutonumberScheme`` tokens, in schema order.
AUTONUMBER_SCHEMES = (
    "alphaLcParenBoth",
    "alphaUcParenBoth",
    "alphaLcParenR",
    "alphaUcParenR",
    "alphaLcPeriod",
    "alphaUcPeriod",
    "arabicParenBoth",
    "arabicParenR",
    "arabicPeriod",
    "arabicPlain",
    "romanLcParenBoth",
    "romanUcParenBoth",
    "romanLcParenR",
    "romanUcParenR",
    "romanLcPeriod",
    "romanUcPeriod",
    "circleNumDbPlain",
    "circleNumWdBlackPlain",
    "circleNumWdWhitePlain",
    "arabicDbPeriod",
    "arabicDbPlain",
    "ea1ChsPeriod",
    "ea1ChsPlain",
    "ea1ChtPeriod",
    "ea1ChtPlain",
    "ea1JpnChsDbPeriod",
    "ea1JpnKorPlain",
    "ea1JpnKorPeriod",
    "arabic1Minus",
    "arabic2Minus",
    "hebrew2Minus",
    "thaiAlphaPeriod",
    "thaiAlphaParenR",
    "thaiAlphaParenBoth",
    "thaiNumPeriod",
    "thaiNumParenR",
    "thaiNumParenBoth",
    "hindiAlphaPeriod",
    "hindiNumPeriod",
    "hindiNumParenR",
    "hindiAlpha1Period",
)
PICTURE_BULLET_FORMATS = ("png", "jpeg", "gif")


def _choice(value: str, choices: _Any, label: str) -> str:
    if not isinstance(value, str) or value not in choices:
        allowed = ", ".join(sorted(choices))
        raise ValueError(f"{label} must be one of: {allowed}")
    return value


def _number(value: _Any, label: str) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise TypeError(f"{label} must be a number")
    number = float(value)
    if not _math.isfinite(number):
        raise ValueError(f"{label} must be finite")
    return number


def _points(value: float | None, label: str, *, low: float, high: float) -> float | None:
    if value is None:
        return None
    number = _number(value, label)
    if not low <= number <= high:
        raise ValueError(f"{label} must be from {low:g} through {high:g} points")
    return number


def _emu(value: float | None) -> int | None:
    return None if value is None else round(value * _EMU_PER_POINT)


def _optional_bool(value: bool | None, label: str) -> bool | None:
    if value is not None and not isinstance(value, bool):
        raise TypeError(f"{label} must be bool or None")
    return value


def _rgb(value: str | None, label: str) -> str | None:
    if value is None:
        return None
    if not isinstance(value, str):
        raise TypeError(f"{label} must be a six-digit RGB string")
    normalized = value.removeprefix("#").upper()
    if len(normalized) != 6 or any(c not in "0123456789ABCDEF" for c in normalized):
        raise ValueError(f"{label} must be a six-digit RGB string")
    return normalized


_MAX_POINTS = _MAX_COORDINATE_EMU / _EMU_PER_POINT


@_dataclass(frozen=True)
class TextSpacing:
    """Line or paragraph spacing: a percentage of the line or exact points.

    ``percent`` follows DrawingML ``a:spcPct`` (100 is single spacing for
    line spacing; for space before/after it is a percentage of the text
    size). ``points`` follows ``a:spcPts``.
    """

    value: float
    unit: str = "points"

    def __post_init__(self) -> None:
        _choice(self.unit, ("percent", "points"), "text spacing unit")
        value = _number(self.value, "text spacing value")
        high = 13_200 if self.unit == "percent" else 1_584
        if not 0 <= value <= high:
            raise ValueError(f"text spacing {self.unit} must be from 0 through {high}")
        object.__setattr__(self, "value", value)

    @classmethod
    def percent(cls, value: float) -> TextSpacing:
        return cls(value, "percent")

    @classmethod
    def points(cls, value: float) -> TextSpacing:
        return cls(value, "points")

    def to_payload(self) -> dict[str, _Any]:
        scale = 1_000 if self.unit == "percent" else 100
        return {"unit": self.unit, "value": round(self.value * scale)}


@_dataclass(frozen=True)
class TextTabStop:
    """One paragraph tab stop, positioned in points from the text box margin."""

    position: float
    alignment: str = "left"

    def __post_init__(self) -> None:
        position = _points(self.position, "tab stop position", low=0, high=_MAX_POINTS)
        object.__setattr__(self, "position", position)
        _choice(self.alignment, TAB_ALIGNMENTS, "tab stop alignment")

    def to_payload(self) -> dict[str, _Any]:
        return {
            "position_emu": _emu(self.position),
            "alignment": TAB_ALIGNMENTS[self.alignment],
        }


# Python 3.9 dataclass fields have no kw_only keyword.
_KW_ONLY_FIELD = {"kw_only": True} if _sys.version_info >= (3, 10) else {}


@_dataclass(frozen=True)
class _BulletStyle:
    """Bullet colour/font, and either percent-of-text or absolute point size."""

    color: str | None = _field(default=None, **_KW_ONLY_FIELD)
    size_percent: float | None = _field(default=None, **_KW_ONLY_FIELD)
    size_points: float | None = _field(default=None, **_KW_ONLY_FIELD)
    font: str | None = _field(default=None, **_KW_ONLY_FIELD)

    def __post_init__(self) -> None:
        object.__setattr__(self, "color", _rgb(self.color, "bullet color"))
        if self.size_percent is not None:
            size = _number(self.size_percent, "bullet size_percent")
            if not 25 <= size <= 400:
                raise ValueError("bullet size_percent must be from 25 through 400")
            object.__setattr__(self, "size_percent", size)
        if self.size_points is not None:
            if self.size_percent is not None:
                raise ValueError("bullet size_points and size_percent cannot both be set")
            object.__setattr__(
                self,
                "size_points",
                _points(self.size_points, "bullet size_points", low=1, high=4_000),
            )
        if self.font is not None and (not isinstance(self.font, str) or not self.font.strip()):
            raise ValueError("bullet font must be a non-empty typeface name")

    def _style_payload(self, kind: str) -> dict[str, _Any]:
        return {
            "kind": kind,
            "color": self.color,
            "size_percent": None if self.size_percent is None else round(self.size_percent * 1_000),
            "size_points": None if self.size_points is None else round(self.size_points * 100),
            "font": self.font,
        }


@_dataclass(frozen=True)
class NoBullet:
    """An explicit ``a:buNone`` that suppresses any inherited bullet."""

    def to_payload(self) -> dict[str, _Any]:
        return {"kind": "none"}


@_dataclass(frozen=True)
class CharacterBullet(_BulletStyle):
    """A character bullet such as ``"•"`` (``a:buChar``).

    ``size_percent`` scales against the text size; ``size_points`` specifies
    an absolute size. Supply at most one of these optional size values.
    """

    character: str = "\u2022"

    def __post_init__(self) -> None:
        super().__post_init__()
        if not isinstance(self.character, str) or len(self.character) != 1:
            raise ValueError("bullet character must be exactly one character")

    def to_payload(self) -> dict[str, _Any]:
        return {**self._style_payload("character"), "character": self.character}


@_dataclass(frozen=True)
class AutoNumberedBullet(_BulletStyle):
    """An automatic numbering scheme (``a:buAutoNum``).

    ``start_at=None`` keeps numbering inherited rather than writing ``startAt``.
    The authored default remains an explicit start at one.
    """

    scheme: str = "arabicPeriod"
    start_at: int | None = 1

    def __post_init__(self) -> None:
        super().__post_init__()
        _choice(self.scheme, AUTONUMBER_SCHEMES, "auto-number scheme")
        if self.start_at is None:
            return
        if isinstance(self.start_at, bool) or not isinstance(self.start_at, int):
            raise TypeError("auto-number start_at must be an integer")
        if not 1 <= self.start_at <= 32_767:
            raise ValueError("auto-number start_at must be from 1 through 32767")

    def to_payload(self) -> dict[str, _Any]:
        return {
            **self._style_payload("auto_number"),
            "scheme": self.scheme,
            "start_at": self.start_at,
        }


@_dataclass(frozen=True)
class PictureBullet(_BulletStyle):
    """A picture bullet (``a:buBlip``) carrying its image bytes."""

    image: bytes = b""
    image_format: str = "png"

    def __post_init__(self) -> None:
        super().__post_init__()
        if not isinstance(self.image, (bytes, bytearray)) or not self.image:
            raise ValueError("picture bullet image must be non-empty bytes")
        object.__setattr__(self, "image", bytes(self.image))
        _choice(self.image_format, PICTURE_BULLET_FORMATS, "picture bullet image_format")

    def to_payload(self) -> dict[str, _Any]:
        return {
            **self._style_payload("picture"),
            "image": self.image,
            "image_format": self.image_format,
        }


if _sys.version_info >= (3, 10):
    Bullet = NoBullet | CharacterBullet | AutoNumberedBullet | PictureBullet
else:
    # Python 3.9 cannot evaluate PEP 604 unions at runtime.
    Bullet = _Union[NoBullet, CharacterBullet, AutoNumberedBullet, PictureBullet]
_BULLET_TYPES = (NoBullet, CharacterBullet, AutoNumberedBullet, PictureBullet)


def _text_runs(runs: _Iterable[_Any] | None, text: str, label: str) -> tuple[_TextRun, ...]:
    from wolfxl.drawing.drawing import TextRun

    if runs is None:
        if not isinstance(text, str):
            raise TypeError(f"{label} text must be a string")
        return (TextRun(text),) if text else ()
    if text:
        raise ValueError(f"{label} text and runs cannot both provide text")
    values = tuple(runs)
    if not all(isinstance(run, TextRun) for run in values):
        raise TypeError(f"{label} runs must contain TextRun values")
    return values


@_dataclass(frozen=True, init=False)
class TextParagraph:
    """One paragraph of shape or text-box text (``a:p``).

    Lengths are points. ``indent`` is the first-line indent relative to
    ``margin_left`` and may be negative for a hanging bullet. ``level`` is the
    DrawingML outline level from 0 through 8.
    """

    runs: tuple[_TextRun, ...]
    alignment: str | None
    level: int | None
    indent: float | None
    margin_left: float | None
    margin_right: float | None
    line_spacing: TextSpacing | None
    space_before: TextSpacing | None
    space_after: TextSpacing | None
    bullet: Bullet | None
    tab_stops: tuple[TextTabStop, ...]
    default_tab_size: float | None
    font_alignment: str | None
    right_to_left: bool | None

    def __init__(
        self,
        runs: _Iterable[_TextRun] | None = None,
        *,
        text: str = "",
        alignment: str | None = None,
        level: int | None = None,
        indent: float | None = None,
        margin_left: float | None = None,
        margin_right: float | None = None,
        line_spacing: TextSpacing | None = None,
        space_before: TextSpacing | None = None,
        space_after: TextSpacing | None = None,
        bullet: Bullet | None = None,
        tab_stops: _Iterable[TextTabStop] = (),
        default_tab_size: float | None = None,
        font_alignment: str | None = None,
        right_to_left: bool | None = None,
    ) -> None:
        setter = object.__setattr__
        setter(self, "runs", _text_runs(runs, text, "paragraph"))
        if alignment is not None:
            _choice(alignment, PARAGRAPH_ALIGNMENTS, "paragraph alignment")
        setter(self, "alignment", alignment)
        if level is not None:
            if isinstance(level, bool) or not isinstance(level, int):
                raise TypeError("paragraph level must be an integer")
            if not 0 <= level <= 8:
                raise ValueError("paragraph level must be from 0 through 8")
        setter(self, "level", level)
        setter(
            self,
            "indent",
            _points(indent, "paragraph indent", low=-_MAX_POINTS, high=_MAX_POINTS),
        )
        setter(
            self,
            "margin_left",
            _points(margin_left, "margin_left", low=0, high=_MAX_POINTS),
        )
        setter(
            self,
            "margin_right",
            _points(margin_right, "margin_right", low=0, high=_MAX_POINTS),
        )
        for name, spacing in (
            ("line_spacing", line_spacing),
            ("space_before", space_before),
            ("space_after", space_after),
        ):
            if spacing is not None and not isinstance(spacing, TextSpacing):
                raise TypeError(f"{name} must be a TextSpacing or None")
            setter(self, name, spacing)
        if bullet is not None and not isinstance(bullet, _BULLET_TYPES):
            raise TypeError(
                "bullet must be NoBullet, CharacterBullet, AutoNumberedBullet, "
                "PictureBullet or None"
            )
        setter(self, "bullet", bullet)
        stops = tuple(tab_stops)
        if not all(isinstance(stop, TextTabStop) for stop in stops):
            raise TypeError("tab_stops must contain TextTabStop values")
        if len(stops) > 32:
            raise ValueError("a paragraph supports at most 32 tab stops")
        setter(self, "tab_stops", stops)
        setter(
            self,
            "default_tab_size",
            _points(default_tab_size, "default_tab_size", low=0, high=_MAX_POINTS),
        )
        if font_alignment is not None:
            _choice(font_alignment, FONT_ALIGNMENTS, "font alignment")
        setter(self, "font_alignment", font_alignment)
        setter(self, "right_to_left", _optional_bool(right_to_left, "right_to_left"))

    @property
    def text(self) -> str:
        return "".join(run.text for run in self.runs)

    def to_payload(self) -> dict[str, _Any]:
        return {
            "runs": _text_runs_payload(self.runs) or [],
            "alignment": None if self.alignment is None else PARAGRAPH_ALIGNMENTS[self.alignment],
            "level": self.level,
            "indent_emu": _emu(self.indent),
            "margin_left_emu": _emu(self.margin_left),
            "margin_right_emu": _emu(self.margin_right),
            "line_spacing": None if self.line_spacing is None else self.line_spacing.to_payload(),
            "space_before": None if self.space_before is None else self.space_before.to_payload(),
            "space_after": None if self.space_after is None else self.space_after.to_payload(),
            "bullet": None if self.bullet is None else self.bullet.to_payload(),
            "tab_stops": [stop.to_payload() for stop in self.tab_stops],
            "default_tab_size_emu": _emu(self.default_tab_size),
            "font_alignment": None
            if self.font_alignment is None
            else FONT_ALIGNMENTS[self.font_alignment],
            "rtl": self.right_to_left,
        }


@_dataclass(frozen=True)
class TextBodyOptions:
    """Shape text-body options (``a:bodyPr``). ``None`` keeps the default.

    Insets are points. ``autofit`` is ``"none"`` (``a:noAutofit``),
    ``"shrink"`` (``a:normAutofit`` with optional ``font_scale`` and
    ``line_space_reduction`` percentages) or ``"shape"`` (``a:spAutoFit``).
    """

    vertical_anchor: str | None = None
    anchor_center: bool | None = None
    vertical_text: str | None = None
    wrap: bool | None = None
    autofit: str | None = None
    font_scale: float | None = None
    line_space_reduction: float | None = None
    horizontal_overflow: str | None = None
    vertical_overflow: str | None = None
    left_inset: float | None = None
    top_inset: float | None = None
    right_inset: float | None = None
    bottom_inset: float | None = None

    def __post_init__(self) -> None:
        if self.vertical_anchor is not None:
            _choice(self.vertical_anchor, VERTICAL_ANCHORS, "vertical_anchor")
        if self.vertical_text is not None:
            _choice(self.vertical_text, VERTICAL_TEXT_TYPES, "vertical_text")
        _optional_bool(self.anchor_center, "anchor_center")
        _optional_bool(self.wrap, "wrap")
        if self.autofit is not None:
            _choice(self.autofit, AUTOFIT_TYPES, "autofit")
        for name, low, high in (
            ("font_scale", 1, 100),
            ("line_space_reduction", 0, 100),
        ):
            value = getattr(self, name)
            if value is None:
                continue
            if self.autofit != "shrink":
                raise ValueError(f"{name} requires autofit='shrink'")
            number = _number(value, name)
            if not low <= number <= high:
                raise ValueError(f"{name} must be from {low} through {high} percent")
            object.__setattr__(self, name, number)
        if self.horizontal_overflow is not None:
            _choice(
                self.horizontal_overflow,
                HORIZONTAL_OVERFLOW_TYPES,
                "horizontal_overflow",
            )
        if self.vertical_overflow is not None:
            _choice(self.vertical_overflow, VERTICAL_OVERFLOW_TYPES, "vertical_overflow")
        for name in ("left_inset", "top_inset", "right_inset", "bottom_inset"):
            object.__setattr__(
                self, name, _points(getattr(self, name), name, low=0, high=_MAX_POINTS)
            )

    def to_payload(self) -> dict[str, _Any]:
        def percent(value: float | None) -> int | None:
            return None if value is None else round(value * 1_000)

        return {
            "anchor": None
            if self.vertical_anchor is None
            else VERTICAL_ANCHORS[self.vertical_anchor],
            "anchor_center": self.anchor_center,
            "vertical": None
            if self.vertical_text is None
            else VERTICAL_TEXT_TYPES[self.vertical_text],
            "wrap": self.wrap,
            "autofit": self.autofit,
            "font_scale": percent(self.font_scale),
            "line_space_reduction": percent(self.line_space_reduction),
            "horizontal_overflow": self.horizontal_overflow,
            "vertical_overflow": self.vertical_overflow,
            "left_inset_emu": _emu(self.left_inset),
            "top_inset_emu": _emu(self.top_inset),
            "right_inset_emu": _emu(self.right_inset),
            "bottom_inset_emu": _emu(self.bottom_inset),
        }


class _ShapeText:
    """Text state shared by :class:`BasicShape` and :class:`TextBox`.

    ``text=`` and ``runs=`` are one-paragraph shorthands; ``paragraphs``
    holds the full paragraph collection. ``body`` holds text-body options.
    """

    _runs: tuple[_TextRun, ...]
    _text: str
    _paragraphs: tuple[TextParagraph, ...] | None
    _body: TextBodyOptions | None
    _runs_label = "shape"

    def _init_text(
        self,
        text: str,
        runs: _Iterable[_TextRun] | None,
        paragraphs: _Iterable[TextParagraph] | None,
        body: TextBodyOptions | None,
    ) -> None:
        self._runs = ()
        self._source_text_body_xml: str | None = None
        self._source_text_body_images: tuple[dict[str, _Any], ...] = ()
        self._text = ""
        self._paragraphs = None
        self.body = body
        if paragraphs is not None:
            if runs is not None or text:
                raise ValueError("paragraphs cannot be combined with text or runs")
            self.paragraphs = paragraphs
            return
        self.runs = runs
        if runs is None:
            self.text = str(text)
        elif text:
            raise ValueError("text and runs cannot both provide text")

    @property
    def text(self) -> str:
        """Plain-text view; paragraphs are joined with newlines."""

        if self._paragraphs is not None:
            return "\n".join(paragraph.text for paragraph in self._paragraphs)
        return "".join(run.text for run in self._runs) if self._runs else self._text

    @text.setter
    def text(self, value: str) -> None:
        self._source_text_body_xml = None
        self._source_text_body_images = ()
        self._text = str(value)
        self._runs = ()
        self._paragraphs = None

    @property
    def runs(self) -> tuple[_TextRun, ...]:
        """Rich runs; a paragraph collection is flattened with newline runs."""

        if self._paragraphs is None:
            return self._runs
        from wolfxl.drawing.drawing import TextRun

        flattened: list[_TextRun] = []
        for index, paragraph in enumerate(self._paragraphs):
            if index:
                flattened.append(TextRun("\n"))
            flattened.extend(paragraph.runs)
        return tuple(flattened)

    @runs.setter
    def runs(self, value: _Iterable[_TextRun] | None) -> None:
        from wolfxl.drawing.drawing import TextRun

        if value is None:
            self._source_text_body_xml = None
            self._source_text_body_images = ()
            self._runs = ()
            self._paragraphs = None
            return
        runs = tuple(value)
        if not runs:
            raise ValueError(f"{self._runs_label} runs must not be empty")
        if not all(isinstance(run, TextRun) for run in runs):
            raise TypeError(f"{self._runs_label} runs must contain TextRun values")
        self._source_text_body_xml = None
        self._source_text_body_images = ()
        self._runs = runs
        self._text = ""
        self._paragraphs = None

    @property
    def paragraphs(self) -> tuple[TextParagraph, ...]:
        """The paragraph collection; shorthand text is one plain paragraph."""

        if self._paragraphs is not None:
            return self._paragraphs
        if self._runs:
            return (TextParagraph(self._runs),)
        return (TextParagraph(text=self._text),) if self._text else ()

    @paragraphs.setter
    def paragraphs(self, value: _Iterable[TextParagraph] | None) -> None:
        if value is None:
            self._source_text_body_xml = None
            self._source_text_body_images = ()
            self._paragraphs = None
            self._runs = ()
            self._text = ""
            return
        paragraphs = tuple(value)
        if not paragraphs:
            raise ValueError("paragraphs must not be empty")
        if not all(isinstance(paragraph, TextParagraph) for paragraph in paragraphs):
            raise TypeError("paragraphs must contain TextParagraph values")
        self._source_text_body_xml = None
        self._source_text_body_images = ()
        self._paragraphs = paragraphs
        self._runs = ()
        self._text = ""

    @property
    def body(self) -> TextBodyOptions | None:
        return self._body

    @body.setter
    def body(self, value: TextBodyOptions | None) -> None:
        if value is not None and not isinstance(value, TextBodyOptions):
            raise TypeError("body must be TextBodyOptions or None")
        self._source_text_body_xml = None
        self._source_text_body_images = ()
        self._body = value

    def _text_body_payload(self) -> dict[str, _Any]:
        """Extra writer payload keys, present only when the model is used."""

        payload: dict[str, _Any] = {}
        if self._paragraphs is not None:
            payload["paragraphs"] = [p.to_payload() for p in self._paragraphs]
        if self._body is not None:
            payload["text_body"] = self._body.to_payload()
        if self._source_text_body_xml is not None:
            # Retain richer source markup on unrelated shape edits; native
            # package emission rebinds its embedded image relationships.
            payload["text_body_xml"] = self._source_text_body_xml
            payload["text_body_images"] = [dict(image) for image in self._source_text_body_images]
        return payload


__all__ = [name for name in list(globals()) if not name.startswith("_")]


def _markup_error(reason: str) -> ValueError:
    return ValueError(f"shape_inline_markup_unsupported: {reason}")


def _markup_characters(text: str) -> None:
    if any(
        not (
            char in "\t\n\r"
            or "\x20" <= char <= "\ud7ff"
            or "\ue000" <= char <= "\ufffd"
            or "\U00010000" <= char <= "\U0010ffff"
        )
        for char in text
    ):
        raise _markup_error("text must contain XML-compatible characters")


# Split only outside quoted font-family values. CSS escapes are deliberately
# limited to quoted backslashes and quotes; no resource or layout grammar exists.
_CSS_DECLARATION = _re.compile(r"""(?:[^;'"]|'(?:[^'\\]|\\['"\\])*'|"(?:[^"\\]|\\['"\\])*")*""")
_CSS_FAMILY = _re.compile(
    r"""'((?:[^'\\\n\r\f]|\\['"\\])*)'|"((?:[^"\\\n\r\f]|\\['"\\])*)"|"""
    r"((?:[^\W\d]|-+[^\W\d])[\w-]*(?: +(?:[^\W\d]|-+[^\W\d])[\w-]*)*)"
)


def _markup_style(value: str) -> dict:
    style = {}
    remaining = value.strip()
    while remaining:
        match = _CSS_DECLARATION.match(remaining)
        declaration = match[0].strip()
        if not declaration or ":" not in declaration:
            raise _markup_error("invalid CSS declaration")
        name, token = (part.strip() for part in declaration.split(":", 1))
        name = name.lower()
        if name == "font-family":
            family_match = _CSS_FAMILY.fullmatch(token)
            if family_match is None:
                raise _markup_error("font-family requires one literal family name")
            family = next(group for group in family_match.groups() if group is not None)
            family = _re.sub(r"""\\(['"\\])""", r"\1", family)
            if not family.strip():
                raise _markup_error("font-family must not be empty")
            _markup_characters(family)
            style["family"] = family
        elif name == "font-size":
            if not _re.fullmatch(r"(?:[0-9]+(?:\.[0-9]{1,2})?|\.[0-9]{1,2})pt", token, _re.I):
                raise _markup_error("font-size requires points with at most two decimal places")
            size = float(token[:-2])
            if not 0 < size <= 4000:
                raise _markup_error("font-size must be greater than 0 and at most 4000pt")
            style["size"] = size
        elif name == "color":
            if _re.fullmatch(r"#[0-9a-fA-F]{6}", token):
                color = token[1:]
            elif _re.fullmatch(r"#[0-9a-fA-F]{3}", token):
                color = "".join(char * 2 for char in token[1:])
            else:
                rgb = _re.fullmatch(
                    r"rgb\(\s*([0-9]{1,3})\s*,\s*([0-9]{1,3})\s*,\s*([0-9]{1,3})\s*\)",
                    token,
                    _re.I,
                )
                if rgb is None or any(int(channel) > 255 for channel in rgb.groups()):
                    raise _markup_error("color requires hexadecimal RGB or integer rgb()")
                color = "".join(f"{int(channel):02X}" for channel in rgb.groups())
            style["color"] = color.upper()
        else:
            raise _markup_error(f"unsupported CSS property {name!r}")
        remaining = remaining[match.end() :]
        if remaining and not remaining.startswith(";"):
            raise _markup_error("invalid CSS syntax")
        remaining = remaining[1:].strip()
    return style


class _InlineMarkupParser(_HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=False)
        self.stack: list[tuple[str, dict]] = [("", {})]
        self.segments: list[tuple[list[str], dict]] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag not in {"b", "strong", "i", "em", "span"}:
            raise _markup_error(f"unsupported tag {tag!r}")
        if len(self.stack) > 64:
            raise _markup_error("nesting limit is 64 elements")
        style = self.stack[-1][1].copy()
        if tag == "span":
            if attrs:
                if len(attrs) != 1 or attrs[0][0] != "style" or attrs[0][1] is None:
                    raise _markup_error("span accepts only one style attribute")
                style.update(_markup_style(attrs[0][1]))
        elif attrs:
            raise _markup_error(f"{tag} does not accept attributes")
        if tag in {"b", "strong"}:
            style["bold"] = True
        elif tag in {"i", "em"}:
            style["italic"] = True
        self.stack.append((tag, style))

    def parse_endtag(self, index: int) -> int:
        end = self.rawdata.find(">", index)
        if end != -1 and not _re.fullmatch(r"</[a-zA-Z]+\s*>", self.rawdata[index : end + 1]):
            raise _markup_error("malformed closing tag")
        return super().parse_endtag(index)

    def handle_endtag(self, tag: str) -> None:
        if len(self.stack) == 1 or self.stack[-1][0] != tag:
            raise _markup_error("tags must be explicitly closed and properly nested")
        self.stack.pop()

    def handle_startendtag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        raise _markup_error("self-closing elements are unsupported")

    def _append(self, text: str) -> None:
        _markup_characters(text)
        if not text:
            return
        style = self.stack[-1][1]
        if self.segments and self.segments[-1][1] == style:
            self.segments[-1][0].append(text)
        else:
            if len(self.segments) >= 4096:
                raise _markup_error("run limit is 4096")
            self.segments.append(([text], style))

    def handle_data(self, data: str) -> None:
        if "<" in data:
            raise _markup_error("literal '<' must be escaped")
        self._append(data)

    def handle_entityref(self, name: str) -> None:
        value = _html.entities.html5.get(name + ";")
        if value is None:
            raise _markup_error(f"unknown entity {name!r}")
        self._append(value)

    def handle_charref(self, name: str) -> None:
        try:
            value = int(name[1:], 16) if name[:1].lower() == "x" else int(name, 10)
            self._append(chr(value))
        except (ValueError, OverflowError) as error:
            raise _markup_error("invalid character reference") from error

    def handle_comment(self, data: str) -> None:
        raise _markup_error("comments are unsupported")

    def handle_decl(self, decl: str) -> None:
        raise _markup_error("declarations are unsupported")

    def unknown_decl(self, data: str) -> None:
        raise _markup_error("declarations are unsupported")

    def handle_pi(self, data: str) -> None:
        raise _markup_error("processing instructions are unsupported")


def _parse_inline_markup(value: str) -> tuple[_TextRun, ...]:
    """Parse a complete, bounded fragment without mutating a drawing object."""
    from wolfxl.drawing.drawing import TextRun

    if not isinstance(value, str):
        raise _markup_error("input must be a string")
    if len(value) > 262_144:
        raise _markup_error("input limit is 262144 characters")
    parser = _InlineMarkupParser()
    try:
        parser.feed(value)
        parser.close()
    except AssertionError as error:
        raise _markup_error("malformed markup") from error
    if len(parser.stack) != 1:
        raise _markup_error("unclosed tag")
    return tuple(TextRun("".join(parts), **style) for parts, style in parser.segments)


def _canonical_inline_markup(runs: tuple[_TextRun, ...], text: str) -> str:
    """Serialize the current run model; never retain input markup."""
    if not runs:
        return _html.escape(text, quote=False)
    fragments = []
    for run in runs:
        fragment = _html.escape(run.text, quote=False)
        if run.italic:
            fragment = f"<i>{fragment}</i>"
        if run.bold:
            fragment = f"<b>{fragment}</b>"
        styles = []
        if run.family is not None:
            family = run.family.replace("\\", "\\\\").replace("'", "\\'")
            styles.append(f"font-family: '{family}'")
        if run.size is not None:
            styles.append(f"font-size: {str(run.size).removesuffix('.0')}pt")
        if run.color is not None:
            styles.append(f"color: #{run.color}")
        if styles:
            style = _html.escape("; ".join(styles), quote=True)
            fragment = f'<span style="{style}">{fragment}</span>'
        fragments.append(fragment)
    return "".join(fragments)
