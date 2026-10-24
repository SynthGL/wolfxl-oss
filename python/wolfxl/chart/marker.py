"""`<c:marker>` and `<c:dPt>` — series markers + per-point overrides.

Mirrors :mod:`openpyxl.chart.marker`.
"""

from __future__ import annotations

from typing import Any

from wolfxl._compat import (
    _install_openpyxl_iter,
    _openpyxl_name_fallback,
    _resolve_openpyxl_class,
)

from .shapes import GraphicalProperties

PRESET_COLORS: list[str] = []


_VALID_SYMBOLS = (
    None,
    "none",
    "auto",
    "circle",
    "dash",
    "diamond",
    "dot",
    "picture",
    "plus",
    "square",
    "star",
    "triangle",
    "x",
)


class Marker:
    """`<c:marker>` — symbol, size, and per-marker shape properties."""

    __slots__ = ("symbol", "size", "spPr", "extLst")

    def __init__(
        self,
        symbol: str | None = None,
        size: int | None = None,
        spPr: GraphicalProperties | None = None,
        extLst: Any = None,  # noqa: N803
        **kw: Any,
    ) -> None:
        if symbol not in _VALID_SYMBOLS:
            raise ValueError(f"symbol={symbol!r} not in {_VALID_SYMBOLS}")
        if size is not None and not (2 <= size <= 72):
            raise ValueError(f"size={size} must be in [2, 72]")
        self.symbol = symbol
        self.size = size
        self.spPr = spPr
        self.extLst = extLst if extLst is not None else kw.get("extLst")

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {}
        if self.symbol is not None:
            d["symbol"] = self.symbol
        if self.size is not None:
            d["size"] = self.size
        if self.spPr is not None:
            d["spPr"] = self.spPr.to_dict()
        return d


class DataPoint:
    """`<c:dPt>` — per-data-point override (colour, marker, explosion …)."""

    __slots__ = (
        "idx",
        "invertIfNegative",
        "marker",
        "bubble3D",
        "explosion",
        "spPr",
    )

    def __init__(
        self,
        idx: int | None = None,
        invertIfNegative: bool | None = None,
        marker: Marker | None = None,
        bubble3D: bool | None = None,
        explosion: int | None = None,
        spPr: GraphicalProperties | None = None,
    ) -> None:
        if idx is not None:
            idx = _coerce_int(idx)
        if marker is not None and not isinstance(marker, Marker):
            raise TypeError(f"{type(self)}.marker should be {Marker} but value is {type(marker)}")
        if explosion is not None:
            explosion = _coerce_int(explosion)
        self.idx = idx
        self.invertIfNegative = invertIfNegative
        self.marker = marker
        self.bubble3D = bubble3D
        self.explosion = explosion
        self.spPr = spPr

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {}
        if self.idx is not None:
            d["idx"] = self.idx
        if self.invertIfNegative is not None:
            d["invertIfNegative"] = self.invertIfNegative
        if self.marker is not None:
            d["marker"] = self.marker.to_dict()
        if self.bubble3D is not None:
            d["bubble3D"] = self.bubble3D
        if self.explosion is not None:
            d["explosion"] = self.explosion
        if self.spPr is not None:
            d["spPr"] = self.spPr.to_dict()
        return d


_MARKER_XML_MODEL_NAMES = ("Marker", "DataPoint")


def _xml_model_names(cls: type) -> tuple[str, ...]:
    return tuple(getattr(cls, "__attrs__", ())) + tuple(getattr(cls, "__elements__", ()))


def _to_openpyxl_model(value: Any) -> Any:
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, list):
        return [_to_openpyxl_model(item) for item in value]
    if isinstance(value, tuple):
        return tuple(_to_openpyxl_model(item) for item in value)

    cls = value.__class__
    upstream_cls = _resolve_openpyxl_class(cls.__module__, cls.__name__)
    if upstream_cls is None or cls is upstream_cls:
        return value

    names = _xml_model_names(upstream_cls)
    if not names:
        return value

    kwargs: dict[str, Any] = {}
    for name in names:
        if hasattr(value, name):
            kwargs[name] = _to_openpyxl_model(getattr(value, name))
    return upstream_cls(**kwargs)


def _from_openpyxl_marker_model(native_cls: type, value: Any) -> Any:
    names = tuple(
        name
        for name in _xml_model_names(value.__class__)
        if name in tuple(getattr(native_cls, "__slots__", ()))
    )
    kwargs: dict[str, Any] = {}
    for name in names:
        if hasattr(value, name):
            kwargs[name] = _to_native_marker_model(getattr(value, name))
    return native_cls(**kwargs)


def _to_native_marker_model(value: Any) -> Any:
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, list):
        return [_to_native_marker_model(item) for item in value]
    if isinstance(value, tuple):
        return tuple(_to_native_marker_model(item) for item in value)

    native_cls = globals().get(value.__class__.__name__)
    if not isinstance(native_cls, type) or native_cls.__name__ not in _MARKER_XML_MODEL_NAMES:
        return value
    upstream_cls = _resolve_openpyxl_class(__name__, native_cls.__name__)
    if upstream_cls is None or not isinstance(value, upstream_cls):
        return value
    return _from_openpyxl_marker_model(native_cls, value)


def _marker_to_tree(
    self: Any,
    tagname: str | None = None,
    idx: int | None = None,
    namespace: str | None = None,
):
    upstream = _to_openpyxl_model(self)
    return upstream.to_tree(tagname=tagname, idx=idx, namespace=namespace)


def _marker_from_tree(cls: type, node: Any) -> Any:
    upstream_cls = _resolve_openpyxl_class(__name__, cls.__name__)
    if upstream_cls is None:
        return _native_marker_from_tree(cls, node)
    return _from_openpyxl_marker_model(cls, upstream_cls.from_tree(node))


def _nested_bool(node: Any) -> bool:
    """``CT_Boolean``: an omitted ``val`` means true."""
    return node.get("val", "1") in ("1", "true")


def _native_marker_from_tree(cls: type, node: Any) -> Any:
    """Hydrate ``c:marker``/``c:dPt`` without openpyxl's descriptors."""
    kwargs: dict[str, Any] = {}
    for child in node:
        if not isinstance(child.tag, str):
            continue
        name = child.tag.rsplit("}", 1)[-1]
        if name == "spPr":
            kwargs[name] = GraphicalProperties.from_tree(child)
        elif name == "marker" and cls is DataPoint:
            kwargs[name] = Marker.from_tree(child)
        elif name in ("invertIfNegative", "bubble3D") and cls is DataPoint:
            kwargs[name] = _nested_bool(child)
        elif name in ("idx", "explosion", "size"):
            kwargs[name] = int(child.get("val"))
        elif name == "symbol" and cls is Marker:
            kwargs[name] = child.get("val")
        elif name == "extLst" and cls is Marker:
            from .text import _TextXML

            kwargs[name] = _TextXML(child)
    return cls(**kwargs)


def _marker_eq(self: Any, other: Any) -> bool:
    if type(self) is not type(other):
        return False
    if _resolve_openpyxl_class(__name__, type(self).__name__) is None:
        # Without openpyxl the conversion below is the identity, and comparing
        # the results would re-enter this method forever.
        return all(getattr(self, n) == getattr(other, n) for n in type(self).__slots__)
    return _to_openpyxl_model(self) == _to_openpyxl_model(other)


def _install_marker_xml_methods(*classes: type) -> None:
    for cls in classes:
        if not hasattr(cls, "to_tree"):
            cls.to_tree = _marker_to_tree  # type: ignore[attr-defined]
        if "from_tree" not in cls.__dict__:
            cls.from_tree = classmethod(_marker_from_tree)  # type: ignore[attr-defined]
        if "__eq__" not in cls.__dict__:
            cls.__eq__ = _marker_eq  # type: ignore[attr-defined]


_install_marker_xml_methods(Marker, DataPoint)

_install_openpyxl_iter(Marker, DataPoint)


def _coerce_int(value: Any) -> int:
    try:
        return int(value)
    except (TypeError, ValueError) as exc:
        raise TypeError("expected <class 'int'>") from exc


__all__ = ["DataPoint", "Marker", "PRESET_COLORS"]

__getattr__ = _openpyxl_name_fallback(globals())
