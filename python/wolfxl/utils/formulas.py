"""``openpyxl.utils.formulas`` — Excel function name catalog.

openpyxl exposes a frozen ``FORMULAE`` set of every Excel-recognised
function name (``"SUM"``, ``"VLOOKUP"``, ``"XLOOKUP"``, ...) so callers
can validate user-supplied formula strings against the canonical list.

Wolfxl's calc engine has its own function registry under
:mod:`wolfxl.calc._functions`; this module exposes that catalogue under
the openpyxl-shaped name.  Names are uppercased to match openpyxl.

Pod 2 (RFC-060).
"""

from __future__ import annotations

from typing import Any

from wolfxl._compat import _openpyxl_name_fallback


def _load_formulae() -> frozenset[str]:
    try:
        from wolfxl.calc._functions import _BUILTINS as _FUNCTIONS

        return frozenset(name.upper() for name in _FUNCTIONS)
    except Exception:  # pragma: no cover — defensive: calc engine optional at import.
        return frozenset()


def _formulae() -> frozenset[str]:
    cached = globals().get("FORMULAE")
    return cached if cached is not None else __getattr__("FORMULAE")


def validate(formula: str) -> bool:
    """Check that a formula uses known Excel function names."""
    assert formula.startswith("=")
    tokenized = Tokenizer(formula)
    for token in tokenized.items:
        if token.type == "FUNC" and token.subtype == "OPEN":
            name = token.value[:-1]
            if not token.value.startswith("_xlfn.") and name not in _formulae():
                raise ValueError(
                    f"Unknown function {token.value} in {formula}. "
                    "The function may need a prefix"
                )
    return True


class Tokenizer:
    """Lazy wrapper for openpyxl's utility-module tokenizer export."""

    def __new__(cls, formula: str):  # noqa: D102
        from wolfxl.formula.tokenizer import Tokenizer as FormulaTokenizer

        return FormulaTokenizer(formula)


__all__ = ["FORMULAE", "Tokenizer", "validate"]  # noqa: F822  # FORMULAE is bound lazily by __getattr__

_name_fallback = _openpyxl_name_fallback(globals())


def __getattr__(name: str) -> Any:
    # The catalogue is built from the calc engine on first use; ``import wolfxl``
    # and the write path never load the engine for it.
    if name == "FORMULAE":
        value = _load_formulae()
        globals()[name] = value
        return value
    return _name_fallback(name)
