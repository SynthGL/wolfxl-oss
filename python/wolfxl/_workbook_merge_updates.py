"""Persist live merge mutations for both writer and source-package save modes."""

from __future__ import annotations

import os
import tempfile
from typing import Any
from xml.etree import ElementTree as ET
import zipfile

from wolfxl._worksheet_collections import _loaded_merged_range_refs

_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_AFTER = {
    "phoneticPr",
    "conditionalFormatting",
    "dataValidations",
    "hyperlinks",
    "printOptions",
    "pageMargins",
    "pageSetup",
    "headerFooter",
    "rowBreaks",
    "colBreaks",
    "customProperties",
    "cellWatches",
    "ignoredErrors",
    "smartTags",
    "drawing",
    "legacyDrawing",
    "legacyDrawingHF",
    "picture",
    "oleObjects",
    "controls",
    "webPublishItems",
    "tableParts",
    "extLst",
}


def apply_merge_updates(wb: Any, filename: str) -> None:
    """Replay live merge refs idempotently; source backends retain their inputs."""
    titles = getattr(wb, "_merge_mutated_sheets", ())
    if not titles:
        return
    from wolfxl._workbook_save import (
        _workbook_relationship_target_to_part,
        _workbook_relationship_targets,
    )

    replacements: dict[str, bytes] = {}
    temporary: str | None = None
    with zipfile.ZipFile(filename) as source:
        package = {
            name: source.read(name) for name in ("xl/workbook.xml", "xl/_rels/workbook.xml.rels")
        }
        targets = _workbook_relationship_targets(package, _REL)
        root = ET.fromstring(package["xl/workbook.xml"])
        for node in root.iter():
            title = node.get("name")
            if node.tag.rsplit("}", 1)[-1] != "sheet" or title not in titles:
                continue
            ws = wb._sheets.get(title)  # noqa: SLF001
            target = targets.get(node.get(f"{{{_REL}}}id", ""))
            if ws is None or target is None:
                continue
            part = _workbook_relationship_target_to_part(target)
            updated = _replace_merge_refs(source.read(part), _loaded_merged_range_refs(ws))
            if updated is not None:
                replacements[part] = updated
        if not replacements:
            return
        descriptor, temporary = tempfile.mkstemp(
            prefix=".wolfxl-merges-", suffix=".xlsx", dir=os.path.dirname(os.path.abspath(filename))
        )
        os.close(descriptor)
        try:
            with zipfile.ZipFile(temporary, "w") as output:
                output.comment = source.comment
                for info in source.infolist():
                    output.writestr(
                        info,
                        replacements.get(info.filename)
                        if info.filename in replacements
                        else source.read(info.filename),
                    )
        except BaseException:
            os.unlink(temporary)
            raise
    try:
        os.replace(temporary, filename)
    finally:
        if temporary is not None and os.path.exists(temporary):
            os.unlink(temporary)


def _replace_merge_refs(xml: bytes, refs: set[str]) -> bytes | None:
    """Splice only the direct mergeCells span; preserve all surrounding bytes."""
    from xml.parsers import expat
    from xml.sax.saxutils import quoteattr

    parser = expat.ParserCreate(namespace_separator="}")
    depth = 0
    spans: list[tuple[int, int]] = []
    original: set[str] = set()
    active: tuple[int, int] | None = None
    insertion: int | None = None
    root_close = 0
    prefix = b""

    def tag_end(offset: int) -> int:
        quote = 0
        for index in range(offset, len(xml)):
            byte = xml[index]
            if quote:
                if byte == quote:
                    quote = 0
            elif byte in (34, 39):
                quote = byte
            elif byte == 62:
                return index + 1
        raise ValueError("unterminated worksheet tag")

    def start(name: str, attributes: dict[str, str]) -> None:
        nonlocal depth, active, insertion, prefix
        offset = parser.CurrentByteIndex
        if depth == 0:
            qname = xml[offset + 1 : tag_end(offset)].split(None, 1)[0].rstrip(b">")
            prefix = qname.split(b":", 1)[0] + b":" if b":" in qname else b""
        elif depth == 1:
            if name == f"{_NS}}}mergeCells":
                active = (offset, tag_end(offset))
            elif name.rsplit("}", 1)[-1] in _AFTER and insertion is None:
                insertion = offset
        elif depth == 2 and active is not None and name == f"{_NS}}}mergeCell":
            original.add(attributes.get("ref", ""))
        depth += 1

    def end(name: str) -> None:
        nonlocal depth, active, root_close
        depth -= 1
        if depth == 1 and name == f"{_NS}}}mergeCells" and active is not None:
            offset, start_end = active
            end_offset = (
                start_end
                if xml[start_end - 2 : start_end] == b"/>"
                else tag_end(parser.CurrentByteIndex)
            )
            spans.append((offset, end_offset))
            active = None
        elif depth == 0:
            root_close = parser.CurrentByteIndex

    parser.StartElementHandler = start
    parser.EndElementHandler = end
    parser.Parse(xml, True)
    if original == refs:
        return None
    tag = prefix + b"mergeCells"
    child = prefix + b"mergeCell"
    replacement = b""
    if refs:
        children = b"".join(
            b"<" + child + b" ref=" + quoteattr(ref).encode("utf-8") + b"/>" for ref in sorted(refs)
        )
        replacement = (
            b"<"
            + tag
            + b' count="'
            + str(len(refs)).encode("ascii")
            + b'">'
            + children
            + b"</"
            + tag
            + b">"
        )
    if spans:
        first, last = spans[0]
        pieces = [xml[:first], replacement]
        for begin, finish in spans[1:]:
            pieces.append(xml[last:begin])
            last = finish
        pieces.append(xml[last:])
        return b"".join(pieces)
    point = insertion if insertion is not None else root_close
    return xml[:point] + replacement + xml[point:]
