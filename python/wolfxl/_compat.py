"""Shared infrastructure for openpyxl-API shims.

Many openpyxl modules expose classes wolfxl doesn't implement yet — charts,
images, pivots, data validations, conditional formatting rules, named styles.
Instead of silently missing the module (``ModuleNotFoundError``, which masks
the real story from users migrating from openpyxl), we expose the module
paths with stub classes that raise a clear ``NotImplementedError`` on
instantiation. The error message tells the user what's missing and points
at modify mode (which preserves most of these features on round-trip without
needing a Python-side class).
"""

from __future__ import annotations

import pkgutil
import sys
from importlib import import_module
from importlib.util import find_spec
from types import ModuleType
from typing import Any
from xml.etree import ElementTree as ET


def _snapshot_real_openpyxl() -> ModuleType | None:
    """Capture a reference to the upstream openpyxl package.

    When openpyxl is installed, the shim classes borrow ``tagname`` and
    ``__attrs__`` from the upstream definitions so vendored corpus tests
    that check XML serialization match openpyxl's conventions (lowercase
    tagnames such as ``phoneticPr``; attribute names like ``fontId``).

    ``wolfxl.openpyxl_compat.install_as_openpyxl()`` later deletes every
    ``openpyxl*`` entry from ``sys.modules``. We keep the root reference so
    an attribute walk (``_REAL_OPENPYXL.styles.fonts``) still resolves via
    Python object refs after that wipe. ``install_as_openpyxl`` therefore calls
    :func:`_prepare_for_openpyxl_alias` first: everything the walk would add
    to the root must be imported before the wipe.

    Returns ``None`` when openpyxl is not installed (production case) or
    when ``sys.modules['openpyxl']`` already points at wolfxl (a previous
    ``install_as_openpyxl`` call within the same process).
    """
    existing = sys.modules.get("openpyxl")
    if existing is not None:
        if getattr(existing, "__name__", "").startswith("wolfxl"):
            return None
        return existing
    try:
        root = import_module("openpyxl")
    except ImportError:
        return None
    if getattr(root, "__name__", "").startswith("wolfxl"):
        return None
    return root


_UNSET: Any = object()
_real_root: Any = _UNSET
_submodules_imported = False
_class_index: dict[str, type] | None = None


def _real_openpyxl_root() -> ModuleType | None:
    """Return the upstream openpyxl root, importing it on first use.

    The import is deferred from ``import wolfxl`` to the first consumer that
    needs an upstream class, so a process that never asks for one never loads
    openpyxl.
    """
    global _real_root
    if _real_root is _UNSET:
        _real_root = _snapshot_real_openpyxl()
    return _real_root


def _real_openpyxl_if_imported() -> ModuleType | None:
    """Return the upstream openpyxl root when it is already imported; never imports it.

    ``None`` means "not imported yet" as well as "not installed" and "aliased to
    wolfxl"; use :func:`_real_openpyxl_pending` to tell the first case apart.
    """
    if _real_root is not _UNSET:
        return _real_root
    existing = sys.modules.get("openpyxl")
    if existing is None or getattr(existing, "__name__", "").startswith("wolfxl"):
        return None
    return _real_openpyxl_root()


def _real_openpyxl_pending() -> bool:
    """True when upstream openpyxl is installed and nothing has imported it yet."""
    if _real_root is not _UNSET or "openpyxl" in sys.modules:
        return False
    try:
        return find_spec("openpyxl") is not None
    except (ImportError, ValueError):
        return False


def _import_real_openpyxl_submodules() -> None:
    """Import every upstream submodule so each is an attribute of its parent.

    ``pkgutil.walk_packages`` imports each submodule, which attaches it to its
    parent. The post-wipe attribute walk and the class-name index both need
    that, but only a lookup that misses the already-imported modules does.
    """
    global _submodules_imported
    if _submodules_imported:
        return
    root = _real_openpyxl_root()
    paths = getattr(root, "__path__", None)
    if paths:
        for info in pkgutil.walk_packages(list(paths), prefix=root.__name__ + "."):
            try:
                import_module(info.name)
            except Exception:  # noqa: BLE001 - any import failure is non-fatal
                pass
    _submodules_imported = True


def _prepare_for_openpyxl_alias() -> None:
    """Finish the upstream snapshot before ``openpyxl*`` leaves ``sys.modules``.

    Called by ``install_as_openpyxl``. After it installs the alias an upstream
    import resolves to wolfxl, so the root, every submodule and the class-name
    index (which scans ``sys.modules``) must already exist by then.
    """
    if _real_openpyxl_root() is not None:
        _import_real_openpyxl_submodules()
        _openpyxl_class_index()


class _RealOpenpyxl:
    """The upstream openpyxl root; imports the remaining submodules on a missing attribute."""

    __slots__ = ("_root",)

    def __init__(self, root: ModuleType) -> None:
        object.__setattr__(self, "_root", root)

    def __getattr__(self, name: str) -> Any:
        root = object.__getattribute__(self, "_root")
        try:
            return getattr(root, name)
        except AttributeError:
            _import_real_openpyxl_submodules()
            return getattr(root, name)


_real_view: _RealOpenpyxl | None = None


def __getattr__(name: str) -> Any:
    """Serve ``_REAL_OPENPYXL`` (``None`` or the upstream root view) on demand."""
    global _real_view
    if name != "_REAL_OPENPYXL":
        raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
    root = _real_openpyxl_root()
    if root is None:
        return None
    if _real_view is None:
        _real_view = _RealOpenpyxl(root)
    return _real_view


def _build_openpyxl_class_index() -> dict[str, type]:
    """Index openpyxl Serialisable subclasses by simple class name.

    Used as a fallback when precise wolfxl→openpyxl module-path resolution
    misses (wolfxl's submodule layout doesn't always match openpyxl's; e.g.
    wolfxl exposes ``PageMargins`` from ``worksheet.print_settings`` while
    openpyxl defines it in ``worksheet.page``). The index prefers classes
    declared in their owning module over re-exports.

    Built on the first fallback lookup, while openpyxl entries are still
    present in ``sys.modules``: ``install_as_openpyxl()`` builds it first
    through :func:`_prepare_for_openpyxl_alias`.
    """
    index: dict[str, type] = {}
    if _real_openpyxl_root() is None:
        return index
    _import_real_openpyxl_submodules()
    for mod_name, module in list(sys.modules.items()):
        if not isinstance(mod_name, str):
            continue
        if not (mod_name == "openpyxl" or mod_name.startswith("openpyxl.")):
            continue
        for attr_name in dir(module):
            if attr_name.startswith("_"):
                continue
            try:
                value = getattr(module, attr_name)
            except Exception:  # noqa: BLE001
                continue
            if not isinstance(value, type):
                continue
            if not hasattr(value, "__attrs__"):
                continue
            existing = index.get(attr_name)
            if existing is None:
                index[attr_name] = value
                continue
            defined_here = getattr(value, "__module__", "") == mod_name
            existing_defined = getattr(existing, "__module__", "") == getattr(
                sys.modules.get(getattr(existing, "__module__", "")), "__name__", ""
            )
            if defined_here and not existing_defined:
                index[attr_name] = value
    return index


def _openpyxl_class_index() -> dict[str, type]:
    global _class_index
    if _class_index is None:
        _class_index = _build_openpyxl_class_index()
    return _class_index


def _walk_attributes(obj: Any, parts: list[str]) -> Any:
    """Follow ``parts`` as attributes from ``obj``; ``None`` when one is missing."""
    for part in parts:
        try:
            obj = getattr(obj, part)
        except AttributeError:
            return None
    return obj


def _resolve_openpyxl_class(module_name: str | None, class_name: str) -> type | None:
    """Look up an upstream openpyxl class for tagname/__attrs__ borrowing.

    Tries precise wolfxl→openpyxl module-path mapping first; falls back to a
    name-only lookup against the global index so layout drift doesn't drop
    coverage. Returns ``None`` when the snapshot is unavailable or no
    matching upstream class exists.
    """
    root = _real_openpyxl_root()
    if root is None:
        return None
    if module_name and module_name.startswith("wolfxl"):
        rest = module_name[len("wolfxl"):].lstrip(".")
        parts = rest.split(".") if rest else []
        obj = _walk_attributes(root, parts)
        if obj is None and not _submodules_imported:
            # Import just the mirrored upstream module; only a miss on that
            # imports every submodule.
            try:
                import_module(f"{root.__name__}.{rest}")
            except Exception:  # noqa: BLE001 - any import failure is non-fatal
                pass
            obj = _walk_attributes(root, parts)
            if obj is None:
                _import_real_openpyxl_submodules()
                obj = _walk_attributes(root, parts)
        if obj is not None:
            candidate = getattr(obj, class_name, None)
            if isinstance(candidate, type):
                return candidate
    fallback = _openpyxl_class_index().get(class_name)
    return fallback if isinstance(fallback, type) else None


def _xml_element_factory():
    """Return the active ``Element`` factory (lxml when available).

    Resolved lazily to avoid a circular import: ``wolfxl.xml.__init__``
    imports ``_openpyxl_name_fallback`` from this module, so we can't
    pull from ``wolfxl.xml.functions`` at module load time.
    """
    from wolfxl.xml.functions import Element

    return Element


def _iter_openpyxl_attrs(obj: Any, names: tuple[str, ...] | None = None):
    attr_names = names or tuple(
        name for name in getattr(obj, "__dict__", {}) if not name.startswith("_")
    )
    for name in attr_names:
        if not hasattr(obj, name):
            continue
        value = getattr(obj, name)
        if value is None or isinstance(value, (list, tuple, dict)):
            continue
        if isinstance(value, bool):
            value = "1" if value else "0"
        else:
            value = str(value)
        yield name, value


def _openpyxl_attr_iter_method(self):
    yield from _iter_openpyxl_attrs(self)


def _install_openpyxl_iter(*classes: type) -> None:
    for cls in classes:
        if "__iter__" not in cls.__dict__:
            cls.__iter__ = _openpyxl_attr_iter_method  # type: ignore[attr-defined]


class _UnsupportedFeature:
    """Base for openpyxl classes that wolfxl exposes as shims.

    Subclasses carry a human-readable ``_feature_name`` and a ``_hint``
    pointing to the recommended path: modify-mode preservation, native
    WolfXL support when available, or an openpyxl fallback. Attempting to
    instantiate raises with the full message so users discover the gap at
    the construction site rather than at a downstream ``AttributeError``.
    """

    _feature_name: str = "<unknown>"
    _hint: str = ""

    def __init__(self, *args: Any, **kwargs: Any) -> None:  # noqa: ARG002
        raise NotImplementedError(
            f"wolfxl does not implement {self._feature_name}. {self._hint} "
            "See https://github.com/SynthGL/wolfxl#openpyxl-compatibility "
            "for compatibility notes."
        )


def _make_stub(name: str, hint: str) -> type:
    """Create a named subclass of ``_UnsupportedFeature`` for shim modules."""
    return type(name, (_UnsupportedFeature,), {"_feature_name": name, "_hint": hint})


class _OpenpyxlSerialisable:
    """Small attribute carrier for openpyxl internal XML model shims.

    openpyxl exposes a large set of ``Serialisable`` classes from internal
    modules. Downstream code often imports and lightly constructs those
    classes while preparing workbooks, even when the resulting object is never
    serialized by the replacement library. This base preserves that import and
    construction shape without promising full XML round-trip semantics.
    """

    tagname: str | None = None
    __attrs__: tuple[str, ...] = ()
    __elements__: tuple[str, ...] = ()

    def __init__(self, *args: Any, **kwargs: Any) -> None:
        attrs = tuple(getattr(self, "__attrs__", ()))
        for name, value in zip(attrs, args):
            setattr(self, name, value)
        if len(args) > len(attrs):
            self._args = args[len(attrs):]
        for name, value in kwargs.items():
            setattr(self, name, value)

    @property
    def count(self) -> int:
        if "_count" in self.__dict__:
            return self.__dict__["_count"]
        for value in self.__dict__.values():
            if isinstance(value, (list, tuple)):
                return len(value)
        return 0

    @count.setter
    def count(self, value: int | None) -> None:
        self.__dict__["_count"] = 0 if value is None else int(value)

    def __iter__(self):
        attrs = tuple(getattr(self, "__attrs__", ())) or None
        yield from _iter_openpyxl_attrs(self, attrs)

    @classmethod
    def from_tree(cls, node: ET.Element) -> "_OpenpyxlSerialisable":
        return cls(**dict(node.attrib))

    def to_tree(
        self,
        tagname: str | None = None,
        idx: int | None = None,  # noqa: ARG002 - openpyxl compatibility
        namespace: str | None = None,  # noqa: ARG002 - openpyxl compatibility
        value: Any = None,  # noqa: ARG002 - openpyxl descriptor compatibility
    ) -> ET.Element:
        if type(self).__name__ == "Serialisable" and type(self).__module__.endswith(
            ".descriptors.serialisable"
        ):
            raise NotImplementedError()
        element = _xml_element_factory()(
            tagname or self.tagname or self.__class__.__name__
        )
        attrs = tuple(getattr(self, "__attrs__", ())) or None
        for attr_name, attr_value in _iter_openpyxl_attrs(self, attrs):
            element.set(attr_name, attr_value)
        return element


def _make_serialisable(
    name: str,
    attrs: tuple[str, ...] = (),
    module_name: str | None = None,
) -> type:
    """Create a passive named openpyxl-style XML model class.

    When ``module_name`` is supplied and the upstream openpyxl snapshot is
    available, copy ``tagname`` and ``__attrs__`` from the equivalent
    upstream class. Falls back to today's behavior (tagname = class name,
    no attribute list) when openpyxl isn't installed.
    """
    tagname = name
    resolved_attrs: tuple[str, ...] = attrs
    upstream = _resolve_openpyxl_class(module_name, name)
    if upstream is not None:
        upstream_tag = getattr(upstream, "tagname", None)
        if isinstance(upstream_tag, str) and upstream_tag:
            tagname = upstream_tag
        upstream_attrs = getattr(upstream, "__attrs__", None)
        if isinstance(upstream_attrs, tuple) and all(
            isinstance(item, str) for item in upstream_attrs
        ):
            resolved_attrs = upstream_attrs
    return type(
        name,
        (_OpenpyxlSerialisable,),
        {"__attrs__": resolved_attrs, "tagname": tagname},
    )


def _openpyxl_name_fallback(module_globals: dict[str, Any]):
    """Return a module ``__getattr__`` for openpyxl helper-name imports.

    This is for openpyxl's descriptor-heavy internal names such as ``Alias``,
    ``NestedBool``, or ``ExtensionList``. Those classes are often imported by
    adapter code even though WolfXL's native implementation does not need the
    same descriptor framework. Returning a passive class keeps source imports
    compatible without pretending the full openpyxl XML model is implemented.
    """

    module_name = str(module_globals.get("__name__", ""))

    def __getattr__(name: str) -> Any:
        if name.startswith("_"):
            raise AttributeError(name)
        if name in _OPENPYXL_CONSTANTS:
            value = _OPENPYXL_CONSTANTS[name]()
        elif name.isupper():
            constants = import_module("wolfxl.xml.constants")
            if not hasattr(constants, name):
                raise AttributeError(name)
            value = getattr(constants, name)
        elif name and name[0].islower():
            dotted = f"{module_name}.{name}"
            try:
                spec = find_spec(dotted)
            except (ModuleNotFoundError, AttributeError, ValueError):
                spec = None
            value = import_module(dotted) if spec is not None else _make_helper(name)
        else:
            value = _make_serialisable(name, module_name=module_name)
        module_globals[name] = value
        return value

    return __getattr__


def _make_helper(name: str):
    def helper(value: Any = None, *args: Any, **kwargs: Any) -> Any:  # noqa: ARG001
        if name in {"safe_string", "escape", "unescape"}:
            return "" if value is None else str(value)
        return value

    helper.__name__ = name
    return helper


def _constant(module: str, attr: str):
    return lambda: getattr(import_module(module), attr)


_OPENPYXL_CONSTANTS = {
    "CHART_NS": _constant("wolfxl.xml.constants", "CHART_NS"),
    "COMMENTS_NS": _constant("wolfxl.xml.constants", "COMMENTS_NS"),
    "CUSTOMUI_NS": _constant("wolfxl.xml.constants", "CUSTOMUI_NS"),
    "DRAWING_NS": _constant("wolfxl.xml.constants", "DRAWING_NS"),
    "IMAGE_NS": _constant("wolfxl.xml.constants", "IMAGE_NS"),
    "REL_NS": _constant("wolfxl.xml.constants", "REL_NS"),
    "SHEET_DRAWING_NS": _constant("wolfxl.xml.constants", "SHEET_DRAWING_NS"),
    "SHEET_MAIN_NS": _constant("wolfxl.xml.constants", "SHEET_MAIN_NS"),
}
