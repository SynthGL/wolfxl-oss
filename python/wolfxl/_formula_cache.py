"""Persist evaluator results as typed SpreadsheetML formula caches.

ECMA-376 Part 1 section 18.3.1.4 (``c``), 18.3.1.40 (``f``) and 18.3.1.96
(``v``) keep a formula separate from its typed cached value, so a formula cell
is ``<c r=".." t=".."><f>formula</f><v>value</v></c>``. XLOOKUP is an Excel
future function and is stored with the ``_xlfn.`` namespace prefix.

Only formula-cell XML is rewritten. Every other worksheet byte, every other
package part, ZIP entry metadata, namespace bindings and the archive comment
are carried over unchanged. Results the evaluator could not compute, or that
have no SpreadsheetML scalar form, are left uncached.
"""

from __future__ import annotations

import copy
import datetime
import math
import numbers
import os
import posixpath
import re
import shutil
import tempfile
import zipfile
from typing import Any

from defusedxml import ElementTree as ET

from wolfxl.formula.tokenizer import Token, Tokenizer, TokenizerError
from wolfxl.utils.datetime import to_excel

_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

_PREFIX = r"(?:[A-Za-z_][\w.-]*:)?"
# A cell is either self-closing (never a formula cell) or has a body. Matching
# the self-closing form explicitly keeps ``<c r="A1" s="1"/>`` from swallowing
# the following cell.
_CELL = re.compile(
    rf"<(?P<tag>{_PREFIX}c)(?P<attrs>(?:\s[^>]*?)?)(?:/>|>(?P<body>.*?)</(?P=tag)>)",
    re.S,
)
_FORMULA = re.compile(
    rf"<(?P<tag>{_PREFIX}f)(?P<attrs>(?:\s[^>]*?)?)(?:/>|>(?P<text>.*?)</(?P=tag)>)",
    re.S,
)
_VALUE = re.compile(rf"<{_PREFIX}v(?:\s[^>]*?)?(?:/>|>.*?</{_PREFIX}v>)", re.S)
_TYPE = re.compile(r"""\s+t\s*=\s*(?:"[^"]*"|'[^']*')""")
_REF = re.compile(r"""(?:^|\s)r\s*=\s*(["'])(.*?)\1""")
_ENTITY = re.compile(r"&(#x[0-9A-Fa-f]+|#[0-9]+|lt|gt|amp|quot|apos);")
_XLOOKUP_BYTES = re.compile(rb"xlookup", re.I)
_NAMED_ENTITIES = {"lt": "<", "gt": ">", "amp": "&", "quot": '"', "apos": "'"}
# XML 1.0 forbids these characters in any text node, escaped or not.
_ILLEGAL_XML = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f\ud800-\udfff\ufffe\uffff]")
_ERROR_CODES = frozenset({"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"})

CalculationState = tuple[dict[str, Any], dict[str, str], bool]


def defined_name_state(wb: Any) -> dict[str, str]:
    """Return defined names as the evaluator sees them: name -> reference text."""
    return {name: str(getattr(value, "value", value)) for name, value in wb.defined_names.items()}


def calculation_state(wb: Any) -> CalculationState:
    """Capture the evaluator inputs a calculation depends on.

    The state holds every non-empty cell (values and formula text) keyed like
    the evaluator's cell references, the defined names, and the date system.
    """
    values: dict[str, Any] = {}
    for name in wb.sheetnames:
        for row in wb[name].iter_rows(values_only=False):
            for cell in row:
                if cell.value is not None:
                    values[f"{name}!{cell.coordinate}"] = cell.value
    return values, defined_name_state(wb), bool(wb.workbook_properties.date1904)


def record_calculation(wb: Any, inputs: dict[str, Any], results: dict[str, Any]) -> None:
    """Remember a calculation so the next save can persist its results.

    ``inputs`` must be the evaluator's cell map as loaded, before evaluation
    replaced formula text with results.
    """
    wb._calculated_state = (  # noqa: SLF001
        dict(inputs),
        defined_name_state(wb),
        bool(wb.workbook_properties.date1904),
    )
    wb._calculated_formula_results = dict(results)  # noqa: SLF001


def saved_formula_results(wb: Any) -> tuple[dict[str, Any] | None, bool]:
    """Return ``(results, stale)`` for the next save.

    ``results`` is ``None`` when :meth:`Workbook.calculate` has not run. When
    the inputs, formulas, sheet names, defined names or date system changed
    since the calculation, ``stale`` is true: no calculated result is written
    and cached formula values are dropped instead of saved out of date. A
    what-if ``recalculate()`` never changes the workbook, so it never affects
    this decision.
    """
    if getattr(wb, "_pending_axis_shifts", None) or getattr(wb, "_pending_range_moves", None):
        wb._cells_relocated_on_save = True  # noqa: SLF001
    state = getattr(wb, "_calculated_state", None)
    if state is None:
        return None, False
    if getattr(wb, "_cells_relocated_on_save", False):
        # Row, column and range moves in modify mode relocate cells only in the
        # saved package; in-memory cells keep their old addresses. Results keyed
        # by those addresses would land on the wrong formulas, so they are
        # never written for this workbook object.
        return {}, True
    if calculation_state(wb) != state:
        return {}, True
    return wb._calculated_formula_results, False  # noqa: SLF001


def _xml_unescape(text: str) -> str:
    def replace(match: re.Match[str]) -> str:
        entity = match[1]
        if entity in _NAMED_ENTITIES:
            return _NAMED_ENTITIES[entity]
        number = int(entity[2:], 16) if entity[1] in "xX" else int(entity[1:])
        return chr(number)

    return _ENTITY.sub(replace, text)


def _xml_escape(text: str) -> str:
    return (
        text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\r", "&#13;")
    )


def _future_formula(text: str) -> str:
    """Return *text* with every ``XLOOKUP(`` call written as ``_xlfn.XLOOKUP(``."""
    if "XLOOKUP" not in text.upper():
        return text
    try:
        tokens = Tokenizer("=" + text).items
    except TokenizerError:
        return text
    changed = False
    for token in tokens:
        if token.type == Token.FUNC and token.subtype == Token.OPEN:
            if token.value.upper() == "XLOOKUP(":
                token.value = "_xlfn.XLOOKUP("
                changed = True
    return "".join(token.value for token in tokens) if changed else text


def _cached_value(value: Any, epoch: datetime.datetime) -> tuple[str | None, str] | None:
    """Return ``(t attribute, <v> text)`` for a supported scalar, else ``None``."""
    from wolfxl.calc._functions import ExcelError

    if isinstance(value, ExcelError):
        return ("e", value.code) if value.code in _ERROR_CODES else None
    if isinstance(value, bool):
        return "b", "1" if value else "0"
    if isinstance(value, str):
        return None if _ILLEGAL_XML.search(value) else ("str", value)
    if isinstance(value, (datetime.datetime, datetime.date, datetime.time, datetime.timedelta)):
        value = to_excel(value, epoch)
    if isinstance(value, numbers.Integral):
        return None, str(int(value))
    if isinstance(value, numbers.Real):
        number = float(value)
        if math.isfinite(number):
            # Excel writes whole numbers without a fractional part.
            if number.is_integer() and abs(number) < 2**53:
                return None, str(int(number))
            return None, repr(number)
    return None


def _sheet_xml(
    xml: bytes,
    name: str,
    results: dict[str, Any] | None,
    stale: bool,
    epoch: datetime.datetime,
) -> bytes:
    """Qualify XLOOKUP and, when *results* is given, rewrite formula caches."""
    if results is None and _XLOOKUP_BYTES.search(xml) is None:
        return xml

    def replace_cell(match: re.Match[str]) -> str:
        body = match["body"]
        if body is None:
            return match[0]
        formula = _FORMULA.search(body)
        if formula is None:
            return match[0]
        attrs = match["attrs"]
        formula_xml = formula[0]
        # Shared-formula followers are self-closing ``<f/>`` elements: they carry
        # a cache but no text to qualify.
        text = formula["text"]
        if text is not None:
            original = _xml_unescape(text)
            qualified = _future_formula(original)
            if qualified != original:
                start, end = formula.span("text")
                offset = formula.start()
                formula_xml = (
                    formula_xml[: start - offset]
                    + _xml_escape(qualified)
                    + formula_xml[end - offset :]
                )
        before, after = body[: formula.start()], body[formula.end() :]
        if results is not None:
            reference = _REF.search(attrs)
            key = f"{name}!{_xml_unescape(reference[2])}" if reference else None
            # A calculation supplies only results it could actually compute.
            # Never carry an old answer for an unsupported or nonscalar result:
            # inputs may have changed before calculate() was called.
            cached = None if stale or key not in results else _cached_value(results[key], epoch)
            before, after = _VALUE.sub("", before), _VALUE.sub("", after)
            attrs = _TYPE.sub("", attrs)
            if cached is not None:
                kind, cached_text = cached
                if kind is not None:
                    attrs += f' t="{kind}"'
                prefix = formula["tag"].removesuffix("f")
                formula_xml += f"<{prefix}v>{_xml_escape(cached_text)}</{prefix}v>"
        return f"<{match['tag']}{attrs}>{before}{formula_xml}{after}</{match['tag']}>"

    try:
        text = xml.decode("utf-8")
    except UnicodeDecodeError:
        return xml
    patched = _CELL.sub(replace_cell, text)
    return xml if patched == text else patched.encode("utf-8")


def _worksheet_parts(src: zipfile.ZipFile) -> list[tuple[str, str]]:
    """Return ``(sheet name, part path)`` for every worksheet in workbook order."""
    names = set(src.namelist())
    if "xl/workbook.xml" not in names or "xl/_rels/workbook.xml.rels" not in names:
        return []
    workbook = ET.fromstring(src.read("xl/workbook.xml"))
    rels = ET.fromstring(src.read("xl/_rels/workbook.xml.rels"))
    targets = {}
    for rel in rels:
        target = rel.get("Target", "")
        if rel.get("Type", "").endswith("/worksheet") and rel.get("TargetMode") != "External":
            joined = target.lstrip("/") if target.startswith("/") else posixpath.join("xl", target)
            targets[rel.get("Id")] = posixpath.normpath(joined)
    parts = []
    for sheet in workbook.findall(f"{{{_MAIN}}}sheets/{{{_MAIN}}}sheet"):
        path = targets.get(sheet.get(f"{{{_REL}}}id"))
        if path is not None and path in names:
            parts.append((sheet.get("name", ""), path))
    return parts


def _mentions_xlookup(src: zipfile.ZipFile, path: str) -> bool:
    """Scan a part in bounded chunks so plain saves never hold a whole sheet."""
    tail = b""
    with src.open(path) as part:
        while chunk := part.read(1 << 20):
            if _XLOOKUP_BYTES.search(tail + chunk):
                return True
            tail = chunk[-(len("xlookup") - 1) :]
    return False


def apply_formula_results(
    wb: Any,
    filename: str,
    results: dict[str, Any] | None,
    stale: bool = False,
) -> None:
    """Qualify XLOOKUP and patch typed formula caches into a saved workbook.

    The package is rewritten atomically and only when a worksheet changed.
    """
    with zipfile.ZipFile(filename) as src:
        updates: dict[str, bytes] = {}
        for name, path in _worksheet_parts(src):
            if results is None and not _mentions_xlookup(src, path):
                continue
            original = src.read(path)
            patched = _sheet_xml(original, name, results, stale, wb.epoch)
            if patched is not original:
                updates[path] = patched
        if not updates:
            return
        comment = src.comment
        parts = [
            (
                copy.copy(info),
                updates[info.filename]
                if info.filename in updates
                else src.read(info.filename),
            )
            for info in src.infolist()
        ]
    # The source must be closed before os.replace: Windows refuses to replace
    # a file that is still open.
    fd, temporary = tempfile.mkstemp(
        prefix=".wolfxl-formula-cache-",
        suffix=".xlsx",
        dir=os.path.dirname(os.path.abspath(filename)),
    )
    os.close(fd)
    try:
        with zipfile.ZipFile(temporary, "w") as dst:
            dst.comment = comment
            for info, data in parts:
                dst.writestr(info, data)
        shutil.copymode(filename, temporary)
        os.replace(temporary, filename)
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)
