"""``CalculatedField`` / ``CalculatedItem`` — pivot calc support.

Both live on the pivot cache, as Excel stores them. A calculated field
is an extra formula-only cache field, usable as a data field. A
calculated item is an extra item of a source field, computed from that
field's other items; every pivot table on the cache shows it.

WolfXL renders the values Excel shows for formulas built from field
or item references, numbers, ``+ - * / ^`` and parentheses (see
:func:`parse_formula`). Other formulas are written for Excel to
evaluate when it refreshes the cache on open.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Callable, Union


_VALID_DATA_TYPES = ("string", "number", "boolean", "date")


@dataclass
class CalculatedField:
    """Calculated field, scoped to the pivot cache.

    Lives in :attr:`PivotCache.calculated_fields`; emitted as an extra
    ``<cacheField formula="…" databaseField="0"/>``. Its formula
    references source fields by name (``"= revenue - cost"``).
    """

    name: str
    formula: str
    data_type: str = "number"

    def __post_init__(self) -> None:
        if not self.name:
            raise ValueError("CalculatedField requires a non-empty name")
        if not self.formula:
            raise ValueError("CalculatedField requires a non-empty formula")
        if self.data_type not in _VALID_DATA_TYPES:
            raise ValueError(
                f"CalculatedField.data_type must be one of "
                f"{_VALID_DATA_TYPES}, got {self.data_type!r}"
            )
        # Parse-validate the formula expression. wolfxl-formula's
        # parser tolerates an optional leading "=" — strip if present.
        expr = self.formula
        if expr.startswith("="):
            expr = expr[1:]
        # Best-effort validation. We don't depend on the formula crate
        # at runtime (no PyO3 binding for it from the Python side); we
        # do a structural sanity check.
        _shallow_validate_formula(expr)

    def to_rust_dict(self) -> dict:
        return {
            "name": self.name,
            "formula": self.formula.lstrip("="),
            "data_type": self.data_type,
        }


@dataclass
class CalculatedItem:
    """Calculated item of a source field, scoped to the pivot cache.

    Lives in :attr:`PivotCache.calculated_items`; emitted as a trailing
    shared item of the field plus a ``<calculatedItem>`` in the cache
    definition. Its formula references the field's items by name
    (``"= North + South"``).
    """

    field_name: str
    item_name: str
    formula: str

    def __post_init__(self) -> None:
        if not self.field_name:
            raise ValueError("CalculatedItem requires a non-empty field_name")
        if not self.item_name:
            raise ValueError("CalculatedItem requires a non-empty item_name")
        if not self.formula:
            raise ValueError("CalculatedItem requires a non-empty formula")
        expr = self.formula
        if expr.startswith("="):
            expr = expr[1:]
        _shallow_validate_formula(expr)

    @property
    def expression(self) -> str:
        """The formula without its leading ``=`` and surrounding spaces."""
        return self.formula.lstrip("=").strip()


def _shallow_validate_formula(expr: str) -> None:
    """Structural sanity check for a calc-field/item formula.

    The Rust ``wolfxl-formula`` crate does the authoritative parse;
    here we only catch the most obvious shape errors (mismatched
    parentheses, empty body) so users don't ship broken formulas
    silently.
    """
    s = expr.strip()
    if not s:
        raise ValueError("formula body is empty")
    depth = 0
    in_str = False
    quote = ""
    for ch in s:
        if in_str:
            if ch == quote:
                in_str = False
            continue
        if ch in ('"', "'"):
            in_str = True
            quote = ch
            continue
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
            if depth < 0:
                raise ValueError(
                    f"unbalanced parentheses in formula: {expr!r}"
                )
    if depth != 0:
        raise ValueError(f"unbalanced parentheses in formula: {expr!r}")
    if in_str:
        raise ValueError(f"unterminated string literal in formula: {expr!r}")


# ---------------------------------------------------------------------------
# Formula evaluation for rendering
# ---------------------------------------------------------------------------


class UnsupportedFormula(ValueError):
    """The formula uses syntax WolfXL does not evaluate (functions, text)."""


# A parsed formula: ("num", value), ("ref", field_or_None, name),
# ("neg", node) or ("bin", operator, left, right).
Node = tuple
_DIV0 = "#DIV/0!"


def parse_formula(expr: str) -> Node:
    """Parse a calculated field/item formula into a small expression tree.

    Supports numbers, bare or ``'quoted'`` names, ``Field[Item]``
    references, unary ``+``/``-``, ``+ - * / ^`` with Excel precedence,
    and parentheses. Anything else raises :class:`UnsupportedFormula`.
    """
    tokens = _tokenize(expr.lstrip("="))
    pos = 0

    def peek() -> tuple[str, object] | None:
        return tokens[pos] if pos < len(tokens) else None

    def take() -> tuple[str, object]:
        nonlocal pos
        tok = peek()
        if tok is None:
            raise UnsupportedFormula(f"unexpected end of formula: {expr!r}")
        pos += 1
        return tok

    def binary(next_level: Callable[[], Node], ops: str) -> Node:
        node = next_level()
        while (tok := peek()) is not None and tok[0] == "op" and tok[1] in ops:
            take()
            node = ("bin", tok[1], node, next_level())
        return node

    def additive() -> Node:
        return binary(multiplicative, "+-")

    def multiplicative() -> Node:
        return binary(power, "*/")

    def power() -> Node:
        # Excel applies negation before ``^`` and evaluates ``^`` left to right.
        return binary(unary, "^")

    def unary() -> Node:
        tok = peek()
        if tok is not None and tok[0] == "op" and tok[1] in "+-":
            take()
            operand = unary()
            return ("neg", operand) if tok[1] == "-" else operand
        return primary()

    def primary() -> Node:
        kind, value = take()
        if kind == "num":
            return ("num", value)
        if kind == "name":
            nxt = peek()
            if nxt is not None and nxt[0] == "op" and nxt[1] == "(":
                raise UnsupportedFormula(f"functions are not evaluated: {expr!r}")
            if nxt is not None and nxt[0] == "item":
                take()
                return ("ref", value, nxt[1])
            return ("ref", None, value)
        if kind == "op" and value == "(":
            node = additive()
            if take() != ("op", ")"):
                raise UnsupportedFormula(f"expected ')' in formula: {expr!r}")
            return node
        raise UnsupportedFormula(f"unexpected {value!r} in formula: {expr!r}")

    node = additive()
    if peek() is not None:
        raise UnsupportedFormula(f"unexpected {peek()[1]!r} in formula: {expr!r}")
    return node


def _tokenize(expr: str) -> list[tuple[str, object]]:
    tokens: list[tuple[str, object]] = []
    i, n = 0, len(expr)
    while i < n:
        ch = expr[i]
        if ch.isspace():
            i += 1
        elif ch in "+-*/^()":
            tokens.append(("op", ch))
            i += 1
        elif ch.isdigit() or (ch == "." and i + 1 < n and expr[i + 1].isdigit()):
            m = _NUMBER.match(expr, i)
            tokens.append(("num", float(m.group())))
            i = m.end()
        elif ch == "'":
            name, i = _quoted(expr, i)
            tokens.append(("name", name))
        elif ch == "[":
            end = expr.find("]", i)
            if end < 0:
                raise UnsupportedFormula(f"unterminated '[' in formula: {expr!r}")
            inner = expr[i + 1 : end].strip()
            if inner.startswith("'"):
                inner, _ = _quoted(inner, 0)
            tokens.append(("item", inner))
            i = end + 1
        elif ch.isalpha() or ch == "_":
            m = _NAME.match(expr, i)
            tokens.append(("name", m.group()))
            i = m.end()
        else:
            raise UnsupportedFormula(f"unexpected {ch!r} in formula: {expr!r}")
    return tokens


def _quoted(expr: str, start: int) -> tuple[str, int]:
    """Read a ``'…'`` name (``''`` escapes a quote); return (name, end)."""
    out = []
    i = start + 1
    while i < len(expr):
        if expr[i] == "'":
            if expr[i + 1 : i + 2] == "'":
                out.append("'")
                i += 2
                continue
            return "".join(out), i + 1
        out.append(expr[i])
        i += 1
    raise UnsupportedFormula(f"unterminated quoted name in formula: {expr!r}")


def _name_ref(name: str) -> str:
    """``name`` as written in a formula, quoted unless a plain name."""
    if _NAME.fullmatch(name):
        return name
    return "'" + name.replace("'", "''") + "'"


def qualify_item_formula(expr: str, field: str, items: dict[str, str]) -> str:
    """Rewrite bare item names in a calculated-item formula as ``Field[Item]``.

    Excel stores calculated-item formulas in that qualified form and
    asks to repair a workbook whose formula names items bare. ``items``
    maps each casefolded item name of ``field`` to its spelling. Text
    other than those bare names is kept verbatim.
    """
    out: list[str] = []
    i, n = 0, len(expr)
    while i < n:
        ch = expr[i]
        if ch == '"':
            end = expr.find('"', i + 1)
            end = n if end < 0 else end + 1
            out.append(expr[i:end])
            i = end
            continue
        if ch == "[":
            end = expr.find("]", i)
            end = n if end < 0 else end + 1
            out.append(expr[i:end])
            i = end
            continue
        if ch.isdigit() or (ch == "." and expr[i + 1 : i + 2].isdigit()):
            end = _NUMBER.match(expr, i).end()
            out.append(expr[i:end])
            i = end
            continue
        if ch == "'" or ch.isalpha() or ch == "_":
            if ch == "'":
                try:
                    name, end = _quoted(expr, i)
                except UnsupportedFormula:
                    out.append(expr[i:])
                    break
            else:
                m = _NAME.match(expr, i)
                name, end = m.group(), m.end()
            rest = expr[end:].lstrip()
            spelled = items.get(name.casefold())
            if spelled is None or rest[:1] in ("(", "["):
                out.append(expr[i:end])
            else:
                out.append(f"{_name_ref(field)}[{_name_ref(spelled)}]")
            i = end
            continue
        out.append(ch)
        i += 1
    return "".join(out)


_NUMBER = re.compile(r"\d*\.?\d+(?:[eE][+-]?\d+)?")
_NAME = re.compile(r"[^\W\d]\w*(?:\.\w+)*")


def references(node: Node) -> list[tuple[str | None, str]]:
    """Every ``(field, name)`` reference in a parsed formula."""
    if node[0] == "ref":
        return [(node[1], node[2])]
    if node[0] == "neg":
        return references(node[1])
    if node[0] == "bin":
        return references(node[2]) + references(node[3])
    return []


def substitute(
    node: Node, replace: Callable[[str | None, str], Node | None]
) -> Node:
    """Copy of ``node`` with each reference swapped for ``replace(field, name)``.

    References for which ``replace`` returns ``None`` are kept.
    """
    if node[0] == "ref":
        return replace(node[1], node[2]) or node
    if node[0] == "neg":
        return ("neg", substitute(node[1], replace))
    if node[0] == "bin":
        return ("bin", node[1], substitute(node[2], replace), substitute(node[3], replace))
    return node


Value = Union[float, str]


def evaluate(node: Node, resolve: Callable[[str | None, str], Value]) -> Value:
    """Evaluate a parsed formula; references resolve through ``resolve``.

    Returns a float, or an Excel error string such as ``"#DIV/0!"``.
    """
    kind = node[0]
    if kind == "num":
        return node[1]
    if kind == "ref":
        return resolve(node[1], node[2])
    if kind == "neg":
        v = evaluate(node[1], resolve)
        return v if isinstance(v, str) else -v
    left = evaluate(node[2], resolve)
    if isinstance(left, str):
        return left
    right = evaluate(node[3], resolve)
    if isinstance(right, str):
        return right
    op = node[1]
    if op == "+":
        return left + right
    if op == "-":
        return left - right
    if op == "*":
        return left * right
    if op == "/":
        return _DIV0 if right == 0 else left / right
    try:
        result = left**right
    except ZeroDivisionError:
        return _DIV0
    return "#NUM!" if isinstance(result, complex) else float(result)
