"""Runtime import aliasing for openpyxl-shaped programs.

This module supports the smallest possible migration for existing openpyxl
programs:

    import wolfxl; wolfxl.install_as_openpyxl()

After that line, later ``import openpyxl`` and ``from openpyxl.styles import
Font`` imports resolve to WolfXL modules inside the current Python process.

pandas gates its ``openpyxl`` engine on ``openpyxl.__version__``. While the
alias is active, pandas' version probe reports :data:`OPENPYXL_API_VERSION`,
the openpyxl release whose API the alias provides, for the aliased module.
``wolfxl.__version__`` itself is never changed.
"""

from __future__ import annotations

import functools
import importlib
import importlib.abc
import importlib.util
import sys
from types import ModuleType
from typing import Any

__all__ = ["OPENPYXL_API_VERSION", "install_as_openpyxl", "uninstall_as_openpyxl"]

#: openpyxl release whose public API the alias provides. WolfXL vendors this
#: release's formula tokenizer.
OPENPYXL_API_VERSION = "3.1.5"

_PANDAS_OPTIONAL = "pandas.compat._optional"
_ALIAS_ORIGINAL = "_wolfxl_alias_original"


def install_as_openpyxl(*, force: bool = True) -> ModuleType:
    """Expose WolfXL under the ``openpyxl`` import name for this process.

    Args:
        force: Replace existing ``openpyxl`` entries in ``sys.modules``. This
            is the default because user programs often import optional helpers
            before their main workbook code. Already-bound local variables from
            an earlier real ``openpyxl`` import cannot be rewritten.

    Returns:
        The top-level :mod:`wolfxl` module, also installed as ``openpyxl``.

    Raises:
        RuntimeError: If ``openpyxl`` is already imported and ``force=False``.
    """
    import wolfxl

    existing = sys.modules.get("openpyxl")
    if existing is not None and existing is not wolfxl and not force:
        raise RuntimeError(
            "openpyxl is already imported; call install_as_openpyxl(force=True) "
            "before importing openpyxl-shaped modules"
        )

    if force:
        _remove_openpyxl_modules()

    _install_finder()
    sys.modules["openpyxl"] = wolfxl
    _alias_loaded_wolfxl_modules(force=True)
    _install_pandas_version_probe()
    return wolfxl


def uninstall_as_openpyxl() -> None:
    """Remove WolfXL's runtime ``openpyxl`` alias from this process."""
    _remove_finder()
    _remove_pandas_version_probe()
    for name, module in list(sys.modules.items()):
        if _is_openpyxl_name(name) and _module_is_wolfxl(module):
            del sys.modules[name]


class _OpenpyxlAliasFinder(importlib.abc.MetaPathFinder, importlib.abc.Loader):
    """Import hook that maps future ``openpyxl.*`` imports to ``wolfxl.*``."""

    def find_spec(
        self,
        fullname: str,
        path: object | None,
        target: ModuleType | None = None,
    ) -> importlib.machinery.ModuleSpec | None:
        if not _is_openpyxl_name(fullname):
            return None

        mapped_name = _wolfxl_name(fullname)
        mapped_spec = importlib.util.find_spec(mapped_name)
        if mapped_spec is None:
            return None

        is_package = mapped_spec.submodule_search_locations is not None
        spec = importlib.util.spec_from_loader(
            fullname,
            self,
            origin=mapped_spec.origin,
            is_package=is_package,
        )
        if spec is not None and is_package:
            spec.submodule_search_locations = list(
                mapped_spec.submodule_search_locations or []
            )
        return spec

    def create_module(self, spec: importlib.machinery.ModuleSpec) -> ModuleType:
        module = importlib.import_module(_wolfxl_name(spec.name))
        sys.modules[spec.name] = module
        _attach_to_parent(spec.name, module)
        return module

    def exec_module(self, module: ModuleType) -> None:
        return None


class _PandasVersionProbeFinder(importlib.abc.MetaPathFinder):
    """Import hook that adapts pandas' optional-dependency version probe.

    It locates ``pandas.compat._optional`` through the remaining finders and
    wraps only its loader, so the module executes unchanged before the probe
    is adapted.
    """

    def find_spec(
        self,
        fullname: str,
        path: object | None,
        target: ModuleType | None = None,
    ) -> importlib.machinery.ModuleSpec | None:
        if fullname != _PANDAS_OPTIONAL:
            return None
        for finder in sys.meta_path:
            if finder is self or isinstance(finder, _PandasVersionProbeFinder):
                continue
            find_spec = getattr(finder, "find_spec", None)
            if find_spec is None:
                continue
            spec = find_spec(fullname, path, target)
            if spec is not None:
                if spec.loader is not None:
                    spec.loader = _PandasVersionProbeLoader(spec.loader)
                return spec
        return None


class _PandasVersionProbeLoader(importlib.abc.Loader):
    """Loader wrapper that adapts the probe after the real module executes."""

    def __init__(self, loader: importlib.abc.Loader) -> None:
        self._loader = loader

    def create_module(self, spec: importlib.machinery.ModuleSpec) -> ModuleType | None:
        return self._loader.create_module(spec)

    def exec_module(self, module: ModuleType) -> None:
        self._loader.exec_module(module)
        _adapt_pandas_version_probe(module)

    def __getattr__(self, name: str) -> Any:
        return getattr(self._loader, name)


def _install_pandas_version_probe() -> None:
    loaded = sys.modules.get(_PANDAS_OPTIONAL)
    if loaded is not None:
        _adapt_pandas_version_probe(loaded)
    if not any(isinstance(finder, _PandasVersionProbeFinder) for finder in sys.meta_path):
        sys.meta_path.insert(0, _PandasVersionProbeFinder())


def _remove_pandas_version_probe() -> None:
    sys.meta_path[:] = [
        finder for finder in sys.meta_path if not isinstance(finder, _PandasVersionProbeFinder)
    ]
    loaded = sys.modules.get(_PANDAS_OPTIONAL)
    adapted = getattr(loaded, "get_version", None)
    original = getattr(adapted, _ALIAS_ORIGINAL, None)
    if original is not None:
        loaded.get_version = original


def _adapt_pandas_version_probe(module: ModuleType) -> None:
    """Report :data:`OPENPYXL_API_VERSION` when pandas probes the aliased module."""
    get_version = getattr(module, "get_version", None)
    if not callable(get_version) or hasattr(get_version, _ALIAS_ORIGINAL):
        return

    @functools.wraps(get_version)
    def get_version_with_openpyxl_alias(dependency: ModuleType) -> str:
        if dependency is sys.modules.get("openpyxl") and _module_is_wolfxl(dependency):
            return OPENPYXL_API_VERSION
        return get_version(dependency)

    setattr(get_version_with_openpyxl_alias, _ALIAS_ORIGINAL, get_version)
    module.get_version = get_version_with_openpyxl_alias


def _install_finder() -> None:
    if not any(isinstance(finder, _OpenpyxlAliasFinder) for finder in sys.meta_path):
        sys.meta_path.insert(0, _OpenpyxlAliasFinder())


def _remove_finder() -> None:
    sys.meta_path[:] = [
        finder for finder in sys.meta_path if not isinstance(finder, _OpenpyxlAliasFinder)
    ]


def _remove_openpyxl_modules() -> None:
    for name in list(sys.modules):
        if _is_openpyxl_name(name):
            del sys.modules[name]


def _alias_loaded_wolfxl_modules(*, force: bool) -> None:
    for name, module in list(sys.modules.items()):
        if not _is_wolfxl_name(name):
            continue
        alias = "openpyxl" + name[len("wolfxl") :]
        if force or alias not in sys.modules:
            sys.modules[alias] = module
            _attach_to_parent(alias, module)


def _attach_to_parent(fullname: str, module: ModuleType) -> None:
    if "." not in fullname:
        return
    parent_name, child_name = fullname.rsplit(".", 1)
    parent = sys.modules.get(parent_name)
    if parent is not None:
        setattr(parent, child_name, module)


def _is_openpyxl_name(name: str) -> bool:
    return name == "openpyxl" or name.startswith("openpyxl.")


def _is_wolfxl_name(name: str) -> bool:
    return name == "wolfxl" or name.startswith("wolfxl.")


def _wolfxl_name(openpyxl_name: str) -> str:
    return "wolfxl" + openpyxl_name[len("openpyxl") :]


def _module_is_wolfxl(module: Any) -> bool:
    module_name = getattr(module, "__name__", "")
    return isinstance(module_name, str) and _is_wolfxl_name(module_name)
