"""`<c:errBars>` — series error bars.

Mirrors :class:`openpyxl.chart.error_bar.ErrorBars`.
"""

from __future__ import annotations

from typing import Any

from wolfxl._compat import (
    _install_openpyxl_iter,
    _openpyxl_name_fallback,
    _resolve_openpyxl_class,
)

from .shapes import GraphicalProperties

_VALID_DIR = (None, "x", "y")
_VALID_BAR_TYPE = ("both", "minus", "plus")
_VALID_VAL_TYPE = ("cust", "fixedVal", "percentage", "stdDev", "stdErr")


class ErrorBars:
    """`<c:errBars>` — direction, magnitude, and source data."""

    __slots__ = (
        "errDir",
        "errBarType",
        "errValType",
        "noEndCap",
        "plus",
        "minus",
        "val",
        "spPr",
        "extLst",
    )

    def __init__(
        self,
        errDir: str | None = None,
        errBarType: str = "both",
        errValType: str = "fixedVal",
        noEndCap: bool | None = None,
        plus: Any | None = None,
        minus: Any | None = None,
        val: float | None = None,
        spPr: GraphicalProperties | None = None,
        extLst: Any = None,  # noqa: N803
        **kw: Any,
    ) -> None:
        if errDir not in _VALID_DIR:
            raise ValueError(f"errDir={errDir!r} not in {_VALID_DIR}")
        if errBarType not in _VALID_BAR_TYPE:
            raise ValueError(f"errBarType={errBarType!r} not in {_VALID_BAR_TYPE}")
        if errValType not in _VALID_VAL_TYPE:
            raise ValueError(f"errValType={errValType!r} not in {_VALID_VAL_TYPE}")
        self.errDir = errDir
        self.errBarType = errBarType
        self.errValType = errValType
        self.noEndCap = noEndCap
        self.plus = plus
        self.minus = minus
        self.val = val
        self.spPr = spPr
        self.extLst = extLst if extLst is not None else kw.get("extLst")

    # openpyxl aliases
    @property
    def direction(self) -> str | None:
        return self.errDir

    @direction.setter
    def direction(self, value: str | None) -> None:
        self.errDir = value

    @property
    def style(self) -> str:
        return self.errBarType

    @style.setter
    def style(self, value: str) -> None:
        self.errBarType = value

    @property
    def size(self) -> str:
        return self.errValType

    @size.setter
    def size(self, value: str) -> None:
        self.errValType = value

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {
            "errBarType": self.errBarType,
            "errValType": self.errValType,
        }
        if self.errDir is not None:
            d["errDir"] = self.errDir
        if self.noEndCap is not None:
            d["noEndCap"] = self.noEndCap
        if self.val is not None:
            d["val"] = self.val
        if self.plus is not None:
            d["plus"] = self.plus.to_dict() if hasattr(self.plus, "to_dict") else self.plus
        if self.minus is not None:
            d["minus"] = self.minus.to_dict() if hasattr(self.minus, "to_dict") else self.minus
        if self.spPr is not None:
            d["spPr"] = self.spPr.to_dict()
        return d


_ERROR_BAR_XML_MODEL_NAMES = ("ErrorBars",)


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


def _from_openpyxl_error_bar_model(native_cls: type, value: Any) -> Any:
    names = _xml_model_names(value.__class__)
    kwargs: dict[str, Any] = {}
    for name in names:
        if hasattr(value, name):
            kwargs[name] = _to_native_error_bar_model(getattr(value, name))
    return native_cls(**kwargs)


def _to_native_error_bar_model(value: Any) -> Any:
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, list):
        return [_to_native_error_bar_model(item) for item in value]
    if isinstance(value, tuple):
        return tuple(_to_native_error_bar_model(item) for item in value)

    native_cls = globals().get(value.__class__.__name__)
    if not isinstance(native_cls, type) or native_cls.__name__ not in _ERROR_BAR_XML_MODEL_NAMES:
        return value
    upstream_cls = _resolve_openpyxl_class(__name__, native_cls.__name__)
    if upstream_cls is None or not isinstance(value, upstream_cls):
        return value
    return _from_openpyxl_error_bar_model(native_cls, value)


def _error_bar_to_tree(
    self: Any,
    tagname: str | None = None,
    idx: int | None = None,
    namespace: str | None = None,
):
    upstream = _to_openpyxl_model(self)
    return upstream.to_tree(tagname=tagname, idx=idx, namespace=namespace)


def _error_bar_from_tree(cls: type, node: Any) -> Any:
    upstream_cls = _resolve_openpyxl_class(__name__, cls.__name__)
    if upstream_cls is None:
        return _native_error_bar_from_tree(cls, node)
    return _from_openpyxl_error_bar_model(cls, upstream_cls.from_tree(node))


def _native_error_bar_from_tree(cls: type, node: Any) -> Any:
    """Hydrate ``c:errBars`` without openpyxl's descriptors."""
    from .data_source import NumDataSource

    kwargs: dict[str, Any] = {}
    for child in node:
        if not isinstance(child.tag, str):
            continue
        name = child.tag.rsplit("}", 1)[-1]
        if name in ("errDir", "errBarType", "errValType"):
            kwargs[name] = child.get("val")
        elif name == "noEndCap":
            kwargs[name] = child.get("val", "1") in ("1", "true")
        elif name == "val":
            kwargs[name] = float(child.get("val"))
        elif name in ("plus", "minus"):
            kwargs[name] = NumDataSource.from_tree(child)
        elif name == "spPr":
            kwargs[name] = GraphicalProperties.from_tree(child)
    return cls(**kwargs)


def _error_bar_eq(self: Any, other: Any) -> bool:
    if type(self) is not type(other):
        return False
    if _resolve_openpyxl_class(__name__, type(self).__name__) is None:
        # Without openpyxl the conversion below is the identity, and comparing
        # the results would re-enter this method forever.
        return all(getattr(self, n) == getattr(other, n) for n in type(self).__slots__)
    return _to_openpyxl_model(self) == _to_openpyxl_model(other)


def _install_error_bar_xml_methods(*classes: type) -> None:
    for cls in classes:
        if not hasattr(cls, "to_tree"):
            cls.to_tree = _error_bar_to_tree  # type: ignore[attr-defined]
        if "from_tree" not in cls.__dict__:
            cls.from_tree = classmethod(_error_bar_from_tree)  # type: ignore[attr-defined]
        if "__eq__" not in cls.__dict__:
            cls.__eq__ = _error_bar_eq  # type: ignore[attr-defined]


_install_error_bar_xml_methods(ErrorBars)

_install_openpyxl_iter(ErrorBars)

__all__ = ["ErrorBars"]

__getattr__ = _openpyxl_name_fallback(globals())
