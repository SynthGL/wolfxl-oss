"""WorkbookEvaluator: recursive expression evaluator for Excel formulas.

Replaces fragile regex-based dispatch with a proper recursive descent
parser that handles balanced parentheses, operator precedence, and
arbitrarily nested expressions like ``=ROUND(SUM(A1:A5)*IF(B1>0,1.1,1.0),2)``.

When the ``formulas`` library is installed (via ``wolfxl[calc]``), unsupported
functions fall back to the library's Excel function implementations.
"""

from __future__ import annotations

import inspect
import logging
import re
from typing import TYPE_CHECKING, Any

from wolfxl.calc._functions import ExcelError, FunctionRegistry, RangeValue, first_error
from wolfxl.calc._graph import DependencyGraph
from wolfxl.calc._parser import expand_range, range_shape
from wolfxl.calc._protocol import CellDelta, RecalcResult

if TYPE_CHECKING:
    from wolfxl._workbook import Workbook

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# formulas library availability
# ---------------------------------------------------------------------------

_formulas_available: bool | None = None


def _check_formulas() -> bool:
    global _formulas_available
    if _formulas_available is None:
        try:
            import formulas  # noqa: F401

            _formulas_available = True
        except ImportError:
            _formulas_available = False
    return _formulas_available


# ---------------------------------------------------------------------------
# Expression parsing helpers
# ---------------------------------------------------------------------------


def _find_matching_paren(expr: str, start: int) -> int:
    """Index of the ``')'`` matching the ``'('`` at *expr[start]*, or -1."""
    depth = 1
    i = start + 1
    in_string = False
    while i < len(expr):
        ch = expr[i]
        if ch == '"':
            in_string = not in_string
        elif not in_string:
            if ch == '(':
                depth += 1
            elif ch == ')':
                depth -= 1
                if depth == 0:
                    return i
        i += 1
    return -1


def _match_function_call(expr: str) -> tuple[str, str] | None:
    """If *expr* is exactly ``FUNC(balanced_args)``, return ``(name, args_str)``.

    Uses balanced parenthesis matching so ``SUM(A1:A5)*2`` is NOT matched
    (there's trailing content after the close-paren).
    """
    stripped = expr.strip()
    m = re.match(r'^([A-Z][A-Z0-9_.]*)\s*\(', stripped, re.IGNORECASE)
    if not m:
        return None
    open_idx = m.end() - 1  # position of '('
    close_idx = _find_matching_paren(stripped, open_idx)
    # The close-paren must be the very last character
    if close_idx >= 0 and close_idx == len(stripped) - 1:
        return (m.group(1), stripped[open_idx + 1 : close_idx])
    return None


def _find_top_level_split(expr: str) -> tuple[str, str, str] | None:
    """Find the rightmost lowest-precedence binary operator at paren depth 0.

    Precedence (lowest to highest)::

        1. comparison   (>=, <=, <>, >, <, =)
        2. additive     (+, -)
        3. multiplicative (*, /)

    Right-to-left scan produces correct left-to-right associativity.
    Returns ``(left, op, right)`` or ``None``.
    """
    length = len(expr)

    for pass_type in ("cmp", "add", "mul"):
        depth = 0
        in_string = False
        i = length - 1
        while i > 0:
            ch = expr[i]

            # Skip string literals
            if ch == '"':
                in_string = not in_string
                i -= 1
                continue
            if in_string:
                i -= 1
                continue

            # Track parentheses (inverted for right-to-left)
            if ch == ')':
                depth += 1
                i -= 1
                continue
            if ch == '(':
                depth -= 1
                i -= 1
                continue

            if depth != 0:
                i -= 1
                continue

            matched_op: str | None = None
            op_start = i

            if pass_type == "cmp":
                # 2-char comparison operators checked first
                if i >= 1 and expr[i - 1 : i + 1] in (">=", "<=", "<>"):
                    matched_op = expr[i - 1 : i + 1]
                    op_start = i - 1
                elif ch in ('>', '<'):
                    matched_op = ch
                elif ch == '=' and not (i >= 1 and expr[i - 1] in ('>', '<', '!')):
                    matched_op = ch
            elif pass_type == "add" and ch in ('+', '-', '&'):
                matched_op = ch
            elif pass_type == "mul" and ch in ('*', '/'):
                matched_op = ch

            if matched_op is not None:
                # Verify it's a binary operator (not unary prefix)
                if op_start <= 0:
                    i -= 1
                    continue
                # Check preceding non-space character
                j = op_start - 1
                while j >= 0 and expr[j] == ' ':
                    j -= 1
                if j < 0 or expr[j] in ('(', ',', '+', '-', '*', '/', '>', '<', '='):
                    i -= 1
                    continue
                # Skip +/- that are part of scientific notation (e.g. 2.5e-1)
                if matched_op in ('+', '-') and j >= 1 and expr[j] in ('e', 'E'):
                    pre_e = j - 1
                    if pre_e >= 0 and expr[pre_e].isdigit():
                        i -= 1
                        continue

                left = expr[:op_start].strip()
                right = expr[op_start + len(matched_op) :].strip()
                if left and right:
                    return (left, matched_op, right)

            i -= 1

    return None


def _has_top_level_colon(expr: str) -> bool:
    """``True`` when *expr* contains ``:`` at paren depth 0 (range ref)."""
    depth = 0
    for ch in expr:
        if ch == '(':
            depth += 1
        elif ch == ')':
            depth -= 1
        elif ch == ':' and depth == 0:
            return True
    return False


def _binary_op(left: Any, op: str, right: Any) -> Any:
    """Evaluate an arithmetic or string binary operation."""
    # Error propagation: if either operand is an error, propagate it
    err = first_error(left, right)
    if err is not None:
        return err
    if op == '&':
        return str(left if left is not None else "") + str(right if right is not None else "")
    if not isinstance(left, (int, float)) or not isinstance(right, (int, float)):
        return None
    if op == '+':
        return left + right
    if op == '-':
        return left - right
    if op == '*':
        return left * right
    if op == '/':
        return ExcelError.DIV0 if right == 0 else left / right
    return None


def _compare(left: Any, right: Any, op: str) -> Any:
    """Evaluate a comparison operation.

    Handles both numeric and string comparisons. String comparisons are
    case-insensitive (matching Excel behavior).  Returns an ExcelError
    if either operand is an error.
    """
    # Error propagation: if either operand is an error, propagate it
    err = first_error(left, right)
    if err is not None:
        return err
    # Both numeric -> numeric comparison
    if isinstance(left, (int, float)) and isinstance(right, (int, float)):
        lf, rf = left, right
    else:
        # Try numeric coercion first
        try:
            lf = float(left) if not isinstance(left, (int, float)) else left
            rf = float(right) if not isinstance(right, (int, float)) else right
        except (ValueError, TypeError):
            # Fall back to string comparison (case-insensitive, like Excel)
            ls = str(left).lower() if left is not None else ""
            rs = str(right).lower() if right is not None else ""
            if op in ('=', '=='):
                return ls == rs
            if op in ('<>', '!='):
                return ls != rs
            if op == '>':
                return ls > rs
            if op == '<':
                return ls < rs
            if op == '>=':
                return ls >= rs
            if op == '<=':
                return ls <= rs
            return False
    if op == '>':
        return lf > rf
    if op == '<':
        return lf < rf
    if op == '>=':
        return lf >= rf
    if op == '<=':
        return lf <= rf
    if op in ('=', '=='):
        return lf == rf
    if op in ('<>', '!='):
        return lf != rf
    return False


def _values_differ(a: Any, b: Any, tolerance: float) -> bool:
    """Check if two values differ beyond tolerance."""
    if a is None and b is None:
        return False
    if a is None or b is None:
        return True
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(float(a) - float(b)) > tolerance
    return a != b


# ---------------------------------------------------------------------------
# Evaluator
# ---------------------------------------------------------------------------


class WorkbookEvaluator:
    """Evaluates Excel formulas in a wolfxl Workbook.

    Usage::

        evaluator = WorkbookEvaluator()
        evaluator.load(workbook)
        results = evaluator.calculate()
        recalc = evaluator.recalculate({"Sheet1!A1": 42.0})
    """

    def __init__(self) -> None:
        self._cell_values: dict[str, Any] = {}
        self._graph = DependencyGraph()
        self._functions = FunctionRegistry()
        self._named_ranges: dict[str, str] = {}  # NAME -> refers_to
        self._loaded = False
        self._use_formulas = _check_formulas()
        self._compiled_cache: dict[str, Any] = {}  # formula -> compiled callable

    def load(self, workbook: Workbook) -> None:
        """Scan workbook, store cell values, build dependency graph."""
        self._cell_values.clear()
        self._graph = DependencyGraph()
        self._named_ranges.clear()

        # Load named ranges first (needed for dependency graph). Loaded
        # workbooks map each name to a DefinedName; its ``value`` holds the
        # reference text.
        for name, defined_name in workbook.defined_names.items():
            refers_to = getattr(defined_name, "value", defined_name)
            if refers_to is not None:
                self._named_ranges[name.upper()] = str(refers_to)

        nr = self._named_ranges if self._named_ranges else None

        for sheet_name in workbook.sheetnames:
            ws = workbook[sheet_name]
            for row in ws.iter_rows(values_only=False):
                for cell in row:
                    val = cell.value
                    cell_ref = f"{sheet_name}!{cell.coordinate}"
                    if isinstance(val, str) and val.startswith("="):
                        # Formula cell: store formula string, register in graph
                        self._cell_values[cell_ref] = val
                        self._graph.add_formula(
                            cell_ref, val, sheet_name, named_ranges=nr,
                        )
                    elif val is not None:
                        # Value cell: store the value
                        self._cell_values[cell_ref] = val

        self._loaded = True

    def calculate(self) -> dict[str, Any]:
        """Evaluate all formulas in topological order.

        Returns dict of cell_ref -> computed value for formula cells.
        """
        if not self._loaded:
            raise RuntimeError("Call load() before calculate()")

        order = self._graph.topological_order()
        results: dict[str, Any] = {}

        for cell_ref in order:
            formula = self._graph.formulas[cell_ref]
            value = self._evaluate_formula(cell_ref, formula)
            self._cell_values[cell_ref] = value
            results[cell_ref] = value

        return results

    def recalculate(
        self,
        perturbations: dict[str, float | int],
        tolerance: float = 1e-10,
    ) -> RecalcResult:
        """Perturb input cells and recompute affected formulas."""
        if not self._loaded:
            raise RuntimeError("Call load() before recalculate()")

        # Snapshot old values for delta computation
        old_values: dict[str, Any] = {}
        for cell_ref in self._graph.formulas:
            old_values[cell_ref] = self._cell_values.get(cell_ref)

        # Apply perturbations
        for cell_ref, value in perturbations.items():
            self._cell_values[cell_ref] = value

        # Find and evaluate affected cells
        affected = self._graph.affected_cells(set(perturbations.keys()))
        for cell_ref in affected:
            formula = self._graph.formulas[cell_ref]
            value = self._evaluate_formula(cell_ref, formula)
            self._cell_values[cell_ref] = value

        # Build deltas
        deltas: list[CellDelta] = []
        propagated = 0
        for cell_ref in affected:
            old_val = old_values.get(cell_ref)
            new_val = self._cell_values.get(cell_ref)
            if _values_differ(old_val, new_val, tolerance):
                propagated += 1
                deltas.append(CellDelta(
                    cell_ref=cell_ref,
                    old_value=old_val,
                    new_value=new_val,
                    formula=self._graph.formulas.get(cell_ref),
                ))

        max_depth = self._graph.max_depth(set(perturbations.keys()))

        return RecalcResult(
            perturbations=dict(perturbations),
            deltas=tuple(deltas),
            total_formula_cells=len(self._graph.formulas),
            propagated_cells=propagated,
            max_chain_depth=max_depth,
        )

    # ------------------------------------------------------------------
    # Formula evaluation (recursive descent)
    # ------------------------------------------------------------------

    def _evaluate_formula(self, cell_ref: str, formula: str) -> Any:
        """Evaluate a single formula string (starting with ``=``).

        Tries the builtin recursive descent evaluator first. If that returns
        None (unsupported function), falls back to the ``formulas`` library.
        """
        body = formula.strip()
        if body.startswith('='):
            body = body[1:]
        sheet = self._sheet_from_ref(cell_ref)
        result = self._eval_expr(body.strip(), sheet)
        if result is not None:
            return result

        # Fallback: try the formulas library for unsupported functions
        if self._use_formulas:
            fb = self._formulas_fallback(formula, sheet)
            if fb is not None:
                return fb

        logger.debug("Cannot evaluate formula %r in %s", formula, cell_ref)
        return None

    def _eval_expr(self, expr: str, sheet: str) -> Any:
        """Recursively evaluate an expression (no leading ``=``).

        Dispatch order (first match wins):

        1. Binary/comparison split at top level (paren-aware, precedence-correct)
        2. Parenthesized sub-expression ``(...)``
        3. Function call ``FUNC(balanced_args)``
        4. Unary minus / plus
        5. Numeric literal
        6. String literal
        7. Boolean literal
        8. Cell reference
        """
        expr = expr.strip()
        if not expr:
            return None

        # 1. Binary split (comparison → additive → multiplicative)
        split = _find_top_level_split(expr)
        if split:
            left_str, op, right_str = split
            left_val = self._eval_expr(left_str, sheet)
            right_val = self._eval_expr(right_str, sheet)
            if op in ('+', '-', '*', '/', '&'):
                return _binary_op(left_val, op, right_val)
            return _compare(left_val, right_val, op)

        # 2. Parenthesized sub-expression: (expr)
        if expr.startswith('('):
            close = _find_matching_paren(expr, 0)
            if close == len(expr) - 1:
                return self._eval_expr(expr[1:close], sheet)

        # 3. Function call: FUNC(balanced_args)
        func = _match_function_call(expr)
        if func:
            return self._eval_function(func[0].upper(), func[1], sheet)

        # 4. Unary minus / plus
        if expr.startswith('-'):
            val = self._eval_expr(expr[1:], sheet)
            if isinstance(val, (int, float)):
                return -val
            return val
        if expr.startswith('+'):
            return self._eval_expr(expr[1:], sheet)

        # 5. Numeric literal (int, float, and scientific notation like 1E3)
        try:
            num = float(expr)
        except ValueError:
            pass
        else:
            # Preserve int for plain integer literals
            if re.fullmatch(r'[+-]?\d+', expr):
                return int(expr)
            return num

        # 6. String literal
        if len(expr) >= 2 and expr[0] == '"' and expr[-1] == '"':
            return expr[1:-1]

        # 7. Boolean
        upper = expr.upper()
        if upper == 'TRUE':
            return True
        if upper == 'FALSE':
            return False

        # 7b. Named range resolution
        if upper in self._named_ranges:
            refers_to = self._named_ranges[upper]
            if ':' in refers_to:
                return self._resolve_range_2d(refers_to, sheet)
            return self._resolve_cell_ref(refers_to, sheet)

        # 8. Cell reference
        return self._resolve_cell_ref(expr, sheet)

    # ------------------------------------------------------------------
    # Atom / argument resolution
    # ------------------------------------------------------------------

    def _resolve_cell_ref(self, expr: str, sheet: str) -> Any:
        """Resolve a cell reference string to its stored value."""
        clean = expr.strip().replace('$', '')
        if '!' in clean:
            parts = clean.split('!', 1)
            ref_sheet = parts[0].strip("'")
            ref = f"{ref_sheet}!{parts[1].upper()}"
        else:
            ref = f"{sheet}!{clean.upper()}"
        return self._cell_values.get(ref)

    def _resolve_range(self, arg: str, sheet: str) -> list[Any]:
        """Resolve a range like ``A1:A5`` to a flat list of cell values.

        Kept for the ``formulas`` library fallback which needs flat lists.
        """
        clean = arg.strip().replace('$', '')
        if '!' not in clean:
            range_ref = f"{sheet}!{clean.upper()}"
        else:
            parts = clean.split('!', 1)
            ref_sheet = parts[0].strip("'")
            range_ref = f"{ref_sheet}!{parts[1].upper()}"
        cells = expand_range(range_ref)
        return [self._cell_values.get(c) for c in cells]

    def _resolve_range_2d(self, arg: str, sheet: str) -> RangeValue:
        """Resolve a range to a :class:`RangeValue` preserving 2D shape."""
        clean = arg.strip().replace('$', '')
        if '!' not in clean:
            range_ref = f"{sheet}!{clean.upper()}"
        else:
            parts = clean.split('!', 1)
            ref_sheet = parts[0].strip("'")
            range_ref = f"{ref_sheet}!{parts[1].upper()}"
        cells = expand_range(range_ref)
        n_rows, n_cols = range_shape(range_ref)
        values = [self._cell_values.get(c) for c in cells]
        return RangeValue(values=values, n_rows=n_rows, n_cols=n_cols)

    # ------------------------------------------------------------------
    # Function dispatch
    # ------------------------------------------------------------------

    def _eval_function(self, func_name: str, args_str: str, sheet: str) -> Any:
        """Evaluate a function call with resolved arguments.

        Functions with ``_raw_args = True`` receive raw argument strings
        and the evaluator's expression evaluator, rather than resolved values.
        """
        func = self._functions.get(func_name)
        if func is None:
            logger.debug("Unsupported function: %s", func_name)
            return None
        if getattr(func, '_raw_args', False):
            raw_args = self._split_top_level_args(args_str)
            try:
                return func(raw_args, self._eval_expr, sheet)
            except Exception as e:
                logger.debug("Error evaluating %s: %s", func_name, e)
                return None
        args = self._parse_function_args(args_str, sheet)
        try:
            return func(args)
        except Exception as e:
            logger.debug("Error evaluating %s: %s", func_name, e)
            return None

    def _split_top_level_args(self, args_str: str) -> list[str]:
        """Split on commas at depth 0 WITHOUT resolving - returns raw strings."""
        args: list[str] = []
        depth = 0
        in_string = False
        current = ""
        for ch in args_str:
            if ch == '"':
                in_string = not in_string
                current += ch
            elif not in_string:
                if ch == '(':
                    depth += 1
                    current += ch
                elif ch == ')':
                    depth -= 1
                    current += ch
                elif ch == ',' and depth == 0:
                    args.append(current)
                    current = ""
                else:
                    current += ch
            else:
                current += ch
        if current:
            args.append(current)
        return args

    def _parse_function_args(self, args_str: str, sheet: str) -> list[Any]:
        """Split on commas at depth 0 (respecting strings), resolve each argument."""
        args: list[Any] = []
        depth = 0
        in_string = False
        current = ""
        i = 0
        length = len(args_str)

        while i < length:
            ch = args_str[i]

            if ch == '"':
                if in_string:
                    # Handle Excel escaped quote ("")
                    if i + 1 < length and args_str[i + 1] == '"':
                        current += '""'
                        i += 2
                        continue
                    in_string = False
                else:
                    in_string = True
                current += ch
            elif not in_string:
                if ch == '(':
                    depth += 1
                    current += ch
                elif ch == ')':
                    depth -= 1
                    current += ch
                elif ch == ',' and depth == 0:
                    args.append(self._resolve_arg(current.strip(), sheet))
                    current = ""
                else:
                    current += ch
            else:
                current += ch

            i += 1

        if current.strip():
            args.append(self._resolve_arg(current.strip(), sheet))

        return args

    def _resolve_arg(self, arg: str, sheet: str) -> Any:
        """Resolve a single function argument.

        Range references (containing ``:`` at depth 0) return a
        :class:`RangeValue` with 2D shape metadata.  Named ranges that
        refer to ranges also resolve to :class:`RangeValue`.  Everything
        else delegates to ``_eval_expr``.
        """
        if not arg:
            return None

        # Range reference at top level
        if _has_top_level_colon(arg) and not arg.startswith('"'):
            return self._resolve_range_2d(arg, sheet)

        # Named range that refers to a range (needed for SUM(MyRange) etc.)
        upper = arg.strip().upper()
        if upper in self._named_ranges:
            refers_to = self._named_ranges[upper]
            if ':' in refers_to:
                return self._resolve_range_2d(refers_to, sheet)
            return self._resolve_cell_ref(refers_to, sheet)

        return self._eval_expr(arg, sheet)

    # ------------------------------------------------------------------
    # formulas library fallback
    # ------------------------------------------------------------------

    def _formulas_fallback(self, formula: str, sheet: str) -> Any:
        """Evaluate a formula via the ``formulas`` library.

        Compiles the formula into a callable, resolves its cell reference
        parameters from ``_cell_values``, and returns the scalar result.
        """
        import formulas as fm
        import numpy as np

        # Compile (with caching)
        compiled = self._compiled_cache.get(formula)
        if compiled is None:
            try:
                result = fm.Parser().ast(formula)
                if result and len(result) > 1:
                    compiled = result[1].compile()
                    self._compiled_cache[formula] = compiled
            except Exception:
                logger.debug("formulas: cannot compile %r", formula)
                return None
        if compiled is None:
            return None

        # Resolve parameters: the compiled function's signature tells us
        # which cell references it needs (e.g., "A1:A5", "B1")
        try:
            params = list(inspect.signature(compiled).parameters.keys())
        except (ValueError, TypeError):
            params = []

        if not params:
            # No cell references - purely constant formula (e.g., =PMT(0.05/12,360,200000))
            try:
                raw = compiled()
                return self._normalize_formulas_result(raw)
            except Exception as e:
                logger.debug("formulas: error evaluating %r: %s", formula, e)
                return None

        # Map parameter names to cell values
        args: list[Any] = []
        for param in params:
            # Param names from formulas lib use the formula's raw ref tokens
            # like "A1:A5" or "A1" (no sheet prefix for same-sheet refs)
            if ':' in param:
                # Range parameter - resolve to numpy array
                # Qualify with sheet name for range_shape parsing
                qualified = param if '!' in param else f"{sheet}!{param}"
                values = self._resolve_range(param, sheet)
                flat = np.array([v if v is not None else 0 for v in values])
                n_rows, n_cols = range_shape(qualified)
                if n_cols > 1 and flat.size == n_rows * n_cols:
                    flat = flat.reshape(n_rows, n_cols)
                args.append(flat)
            else:
                # Single cell parameter
                val = self._resolve_cell_ref(param, sheet)
                if isinstance(val, (int, float)):
                    args.append(np.float64(val))
                elif isinstance(val, str):
                    args.append(val)
                else:
                    args.append(np.float64(0) if val is None else val)

        try:
            raw = compiled(*args)
            return self._normalize_formulas_result(raw)
        except Exception as e:
            logger.debug("formulas: error evaluating %r: %s", formula, e)
            return None

    @staticmethod
    def _normalize_formulas_result(raw: Any) -> Any:
        """Convert a ``formulas`` library result to a plain Python value."""
        if raw is None:
            return None
        # numpy scalar types
        if hasattr(raw, 'item'):
            try:
                val = raw.item()
                if isinstance(val, float) and val == int(val):
                    return int(val)
                return val
            except (ValueError, TypeError):
                pass
        # numpy array with single element
        if hasattr(raw, 'shape') and hasattr(raw, 'flat'):
            try:
                if raw.size == 1:
                    val = raw.flat[0]
                    if hasattr(val, 'item'):
                        val = val.item()
                    if isinstance(val, float) and val == int(val):
                        return int(val)
                    return val
            except (ValueError, TypeError, IndexError):
                pass
        # Already a plain Python type
        if isinstance(raw, (int, float, str, bool)):
            return raw
        return raw

    @staticmethod
    def _sheet_from_ref(cell_ref: str) -> str:
        """Extract sheet name from a canonical cell reference."""
        if '!' in cell_ref:
            return cell_ref.rsplit('!', 1)[0]
        return 'Sheet1'
