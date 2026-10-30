"""Cell payload conversion helpers."""

from __future__ import annotations

from datetime import date, datetime, time, timedelta
from functools import lru_cache
from typing import Any

from wolfxl._styles import Alignment, Border, Color, Font, PatternFill, Side
from wolfxl.styles.fills import GradientFill
from wolfxl.styles.protection import Protection
from wolfxl.utils.datetime import time_to_days

_ERROR_CODES = {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"}
_DEFAULT_FONT = Font()
_DEFAULT_FILL = PatternFill()


def payload_to_python(payload: Any) -> Any:
    """Convert a Rust cell-value payload dict to a plain Python value."""
    if not isinstance(payload, dict):
        return payload
    t = payload.get("type", "blank")
    v = payload.get("value")
    if t == "blank":
        return None
    if t == "string":
        return v
    if t == "number":
        return v
    if t == "boolean":
        return bool(v)
    if t == "error":
        return v
    if t == "formula":
        return payload.get("formula", v)
    if t == "date":
        if isinstance(v, str):
            return datetime.fromisoformat(v)
        if isinstance(v, date) and not isinstance(v, datetime):
            return datetime.combine(v, datetime.min.time())
        return v
    if t == "datetime":
        if isinstance(v, str):
            return datetime.fromisoformat(v)
        return v
    return v


def format_to_font(payload: Any) -> Font:
    """Extract Font fields from a Rust format dict."""
    if not isinstance(payload, dict) or not payload:
        return _DEFAULT_FONT
    color_raw = payload.get("font_color")
    color_key = str(color_raw) if color_raw else None
    return _cached_font(
        payload.get("font_name"),
        payload.get("font_size"),
        bool(payload.get("bold", False)),
        bool(payload.get("italic", False)),
        payload.get("underline"),
        bool(payload["strikethrough"]) if "strikethrough" in payload else None,
        color_key,
    )


@lru_cache(maxsize=512)
def _cached_font(
    name: Any,
    size: Any,
    bold: bool,
    italic: bool,
    underline: Any,
    strike: bool | None,
    color_key: str | None,
) -> Font:
    """Return a shared Font for repeated read-side style payloads."""
    return Font(
        name=name,
        size=size,
        bold=bold,
        italic=italic,
        underline=underline,
        strike=strike,
        color=_format_color_to_color(color_key) if color_key else None,
    )


def format_to_fill(payload: Any) -> PatternFill | GradientFill:
    """Extract Fill fields from a Rust format dict.

    Returns a ``GradientFill`` when the format dict carries a ``gradient``
    sub-dict (round-tripped from the writer), otherwise a ``PatternFill``.
    """
    if not isinstance(payload, dict) or not payload:
        return _DEFAULT_FILL
    grad = payload.get("gradient")
    if isinstance(grad, dict):
        stops_raw = grad.get("stops") or grad.get("stop") or []
        stops: list[Any] = []
        for entry in stops_raw:
            if isinstance(entry, dict):
                pos = entry.get("position", 0.0)
                color = entry.get("color")
                stops.append((float(pos), color))
        return GradientFill(
            type=str(grad.get("type", "linear")),
            degree=float(grad.get("degree", 0.0)),
            left=float(grad.get("left", 0.0)),
            right=float(grad.get("right", 0.0)),
            top=float(grad.get("top", 0.0)),
            bottom=float(grad.get("bottom", 0.0)),
            stop=stops,
        )
    bg = payload.get("bg_color")
    if bg:
        return _cached_solid_fill(str(bg))
    return _DEFAULT_FILL


@lru_cache(maxsize=512)
def _cached_solid_fill(bg: str) -> PatternFill:
    """Return a shared solid fill for repeated read-side style payloads."""
    return PatternFill(patternType="solid", fgColor=_format_color_to_color(bg))


def format_to_alignment(payload: Any) -> Alignment:
    """Extract Alignment fields from a Rust format dict."""
    if not isinstance(payload, dict) or not payload:
        return Alignment()
    return Alignment(
        horizontal=payload.get("h_align"),
        vertical=payload.get("v_align"),
        wrap_text=bool(payload.get("wrap", False)),
        text_rotation=int(payload.get("rotation", 0)),
        indent=int(payload.get("indent", 0)),
    )


def border_payload_to_border(payload: Any) -> Border:
    """Convert a Rust border dict to a Border dataclass.

    Diagonal direction is encoded as two optional Side dicts
    (``diagonal_up`` / ``diagonal_down``); the OOXML model has a
    single ``<diagonal>`` element gated by the ``diagonalUp`` /
    ``diagonalDown`` bool attrs on the parent ``<border>``. The Side
    is taken from whichever direction is present (downward wins on
    conflict, matching the writer's intern logic).
    """
    if not isinstance(payload, dict) or not payload:
        return Border()
    diag_up_payload = payload.get("diagonal_up")
    diag_down_payload = payload.get("diagonal_down")
    diagonal_up = isinstance(diag_up_payload, dict)
    diagonal_down = isinstance(diag_down_payload, dict)
    if diagonal_down:
        diagonal = _edge_to_side(diag_down_payload)
    elif diagonal_up:
        diagonal = _edge_to_side(diag_up_payload)
    else:
        diagonal = Side()
    return Border(
        left=_edge_to_side(payload.get("left")),
        right=_edge_to_side(payload.get("right")),
        top=_edge_to_side(payload.get("top")),
        bottom=_edge_to_side(payload.get("bottom")),
        diagonal=diagonal,
        diagonalUp=diagonal_up,
        diagonalDown=diagonal_down,
    )


def _edge_to_side(edge: Any) -> Side:
    if not isinstance(edge, dict):
        return Side()
    color_raw = edge.get("color")
    return Side(
        style=edge.get("style"),
        color=_format_color_to_color(color_raw) if color_raw else None,
    )


def _format_color_to_color(value: Any) -> Color:
    """Normalize Rust-side style color payloads to openpyxl-shaped Color objects."""
    return Color.from_hex(value if isinstance(value, str) else str(value))


def python_value_to_payload(value: Any) -> dict[str, Any]:
    """Convert a plain Python value to a Rust cell-value payload dict."""
    if value is None:
        return {"type": "blank"}
    if isinstance(value, bool):
        return {"type": "boolean", "value": value}
    if isinstance(value, (int, float)):
        return {"type": "number", "value": value}
    if isinstance(value, datetime):
        return {"type": "datetime", "value": value.replace(microsecond=0).isoformat()}
    if isinstance(value, date):
        return {"type": "date", "value": value.isoformat()}
    if isinstance(value, timedelta):
        return {"type": "timedelta", "value": value.total_seconds() / 86400}
    if isinstance(value, time):
        return {"type": "time", "value": time_to_days(value)}
    if isinstance(value, str) and value.startswith("="):
        return {"type": "formula", "formula": value, "value": value}
    if isinstance(value, str) and value in _ERROR_CODES:
        return {"type": "error", "value": value}
    return {"type": "string", "value": str(value)}


def font_to_format_dict(font: Font) -> dict[str, Any]:
    """Convert a Font to a Rust format dict."""
    d: dict[str, Any] = {}
    if font.bold:
        d["bold"] = True
    if font.italic:
        d["italic"] = True
    if font.underline:
        d["underline"] = font.underline
    if font.strike:
        d["strikethrough"] = True
    if font.name:
        d["font_name"] = font.name
    if font.size is not None:
        d["font_size"] = font.size
    color_hex = font._color_hex()  # noqa: SLF001
    if color_hex:
        d["font_color"] = color_hex
    return d


def fill_to_format_dict(fill: Any) -> dict[str, Any]:
    """Convert a Fill (PatternFill or GradientFill) to a Rust format dict.

    GradientFill round-trips as a nested ``gradient`` sub-dict so the Rust
    bridge can route it into ``GradientFillSpec`` without disturbing the
    pattern-fill path. Stops are normalized to ``[{position, color}, ...]``.
    """
    if isinstance(fill, GradientFill):
        stops: list[dict[str, Any]] = []
        # openpyxl's ``stop`` accepts heterogeneous entries; normalize.
        for idx, entry in enumerate(fill.stop):
            position: float = float(idx)
            color: str | None = None
            if isinstance(entry, tuple) and len(entry) == 2:
                position = float(entry[0])
                color = str(entry[1])
            elif isinstance(entry, dict):
                position = float(entry.get("position", idx))
                color = entry.get("color")
            elif isinstance(entry, str):
                color = entry
            else:
                # openpyxl ``Stop`` / ``Color`` objects — best-effort.
                pos_attr = getattr(entry, "position", None)
                if pos_attr is not None:
                    position = float(pos_attr)
                col_attr = (
                    getattr(entry, "color", None)
                    or getattr(entry, "rgb", None)
                    or getattr(entry, "value", None)
                )
                if col_attr is not None:
                    color = str(getattr(col_attr, "rgb", None) or col_attr)
            stops.append({"position": position, "color": color})
        return {
            "gradient": {
                "type": fill.type,
                "degree": float(fill.degree),
                "left": float(fill.left),
                "right": float(fill.right),
                "top": float(fill.top),
                "bottom": float(fill.bottom),
                "stops": stops,
            }
        }
    d: dict[str, Any] = {}
    fg_getter = getattr(fill, "_fg_hex", None)
    if callable(fg_getter):
        fg = fg_getter()
        if fg:
            d["bg_color"] = fg
    return d


def alignment_to_format_dict(alignment: Alignment) -> dict[str, Any]:
    """Convert an Alignment to a Rust format dict."""
    d: dict[str, Any] = {}
    if alignment.horizontal:
        d["h_align"] = alignment.horizontal
    if alignment.vertical:
        d["v_align"] = alignment.vertical
    if alignment.wrap_text:
        d["wrap"] = True
    if alignment.text_rotation:
        d["rotation"] = alignment.text_rotation
    if alignment.indent:
        d["indent"] = alignment.indent
    return d


def format_to_protection(payload: Any) -> Protection | None:
    """Extract Protection fields from a Rust format dict.

    Returns ``None`` when neither ``locked`` nor ``hidden`` are present so
    callers can distinguish "no protection emitted" from "default
    protection". Excel's default is ``locked=True, hidden=False``.
    """
    if not isinstance(payload, dict):
        return None
    if "locked" not in payload and "hidden" not in payload:
        return None
    return Protection(
        locked=bool(payload.get("locked", True)),
        hidden=bool(payload.get("hidden", False)),
    )


def protection_to_format_dict(protection: Protection) -> dict[str, Any]:
    """Convert a Protection to a Rust format dict.

    Always emits both keys so the writer's `applyProtection="1"` gate
    fires for any non-default Protection (including the default itself
    after explicit assignment, since the user opting in to the default
    shouldn't silently lose the override on round-trip).
    """
    return {
        "locked": bool(protection.locked),
        "hidden": bool(protection.hidden),
    }


def border_to_rust_dict(border: Border) -> dict[str, Any]:
    """Convert a Border to a Rust border dict.

    Diagonal direction is encoded as two optional Side dicts
    (``diagonal_up`` / ``diagonal_down``). Each is emitted only when
    the corresponding bool flag on the Border is True AND the shared
    ``diagonal`` Side has a style; this matches the OOXML semantics
    where a single ``<diagonal>`` element is gated by ``diagonalUp``
    / ``diagonalDown`` attrs on the parent ``<border>``.
    """
    d: dict[str, Any] = {}
    for edge_name in ("left", "right", "top", "bottom"):
        side: Side = getattr(border, edge_name)
        if side.style:
            edge: dict[str, str] = {"style": side.style}
            color = side._color_hex()  # noqa: SLF001
            if color:
                edge["color"] = color
            else:
                edge["color"] = "#000000"
            d[edge_name] = edge

    diagonal: Side = border.diagonal
    if diagonal.style and (border.diagonalUp or border.diagonalDown):
        diag_edge: dict[str, str] = {"style": diagonal.style}
        diag_color = diagonal._color_hex()  # noqa: SLF001
        diag_edge["color"] = diag_color if diag_color else "#000000"
        if border.diagonalUp:
            d["diagonal_up"] = dict(diag_edge)
        if border.diagonalDown:
            d["diagonal_down"] = dict(diag_edge)
    return d
