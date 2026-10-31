"""External links surface (RFC-071 / G18).

This module wraps the Rust parsers
(``wolfxl._rust.parse_external_link_part`` /
``parse_external_link_rels``) into Python dataclasses that match
openpyxl's ``ExternalLink`` / ``ExternalFileLink`` shape just enough for
the compat oracle and downstream introspection:

* ``ExternalLink.target`` — linked workbook filename, e.g. ``ext.xlsx``.
* ``ExternalLink.sheet_names`` — sheet names referenced by formulas.
* ``ExternalLink.cached_data`` — a loose dict (sheetId -> [{r, v}]) of
  any cached cell values present in ``<sheetDataSet>``.
* ``ExternalLink.rid`` — the rels id under which the part is wired into
  ``xl/workbook.xml.rels``.
* ``ExternalLink.file_link`` — small companion holding ``target`` /
  ``target_mode`` so callers can check ``link.file_link.target_mode ==
  "External"``.

The collection is mutable: callers can append/remove links and update
targets. Modify-mode saves rewrite the workbook relationship graph,
content types, external-link parts, and per-link rels when the
collection is dirty; unchanged collections preserve source bytes.
"""

from __future__ import annotations

import zipfile
import os
import mmap
import re
import tempfile
from dataclasses import dataclass, field
from typing import Any
from xml.etree import ElementTree as ET

from wolfxl._compat import _install_openpyxl_iter
from wolfxl._zip_safety import read_entry, validate_zipfile

_RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_REL_TYPE_EXTERNAL_LINK = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink"
)
_REL_TYPE_EXTERNAL_LINK_PATH = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath"
)
_REL_TYPE_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_EXTERNAL_LINK_CT = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"
)
_FORMULA_TAG_RE = re.compile(
    rb"(<(?:[A-Za-z_][\w.-]*:)?(?:f|formula|formula1|formula2|definedName)\b[^>]*>)"
    rb"(.*?)"
    rb"(</(?:[A-Za-z_][\w.-]*:)?(?:f|formula|formula1|formula2|definedName)>)",
    re.DOTALL,
)
_FORMULA_PART_PREFIXES = (
    "xl/worksheets/",
    "xl/charts/",
    "xl/chartsheets/",
    "xl/tables/",
    "xl/pivotCache/",
)


@dataclass
class ExternalFileLink:
    """Reference to the linked workbook file.

    ``target`` is the relationship Target verbatim — e.g. ``ext.xlsx`` for
    a sibling file or ``../foo/data.xlsx`` for one in a parent folder. We
    deliberately do NOT URL-decode it because Open Packaging Conventions
    treat the Target as opaque.
    """

    target: str
    target_mode: str = "External"


@dataclass
class ExternalLink:
    """One entry in :attr:`Workbook._external_links`.

    ``cached_data`` shape is loose by design — see RFC-071 §3.
    """

    file_link: ExternalFileLink | None = None
    rid: str = ""
    target: str = ""
    sheet_names: list[str] = field(default_factory=list)
    cached_data: dict[str, list[dict[str, str]]] = field(default_factory=dict)
    externalBook: Any = None  # noqa: N815
    ddeLink: Any = None  # noqa: N815
    oleLink: Any = None  # noqa: N815
    extLst: Any = None  # noqa: N815
    _collection: "ExternalLinkCollection | None" = field(
        default=None, repr=False, compare=False
    )
    _original_target: str = field(default="", repr=False, compare=False)

    def __post_init__(self) -> None:
        if self.file_link is None:
            self.file_link = ExternalFileLink(target=self.target)
        if not self.target:
            self.target = self.file_link.target
        if self.file_link.target != self.target:
            self.file_link.target = self.target
        if not self._original_target:
            self._original_target = self.target

    def update_target(self, target: str) -> None:
        self.target = target
        if self.file_link is None:
            self.file_link = ExternalFileLink(target=target)
        else:
            self.file_link.target = target
        self._mark_dirty()

    @property
    def path(self) -> str:
        return f"/xl/externalLinks/externalLink{getattr(self, '_id', None)}.xml"

    def _mark_dirty(self) -> None:
        if self._collection is not None:
            self._collection._mark_dirty()

    def _signature(self) -> tuple[Any, ...]:
        return (
            self.target,
            self.file_link.target_mode if self.file_link is not None else "External",
            tuple(self.sheet_names),
            _freeze_cached_data(self.cached_data),
        )


class ExternalLinkCollection(list[ExternalLink]):
    """Mutable workbook external-link collection with dirty tracking."""

    def __init__(self, links: list[ExternalLink] | None = None) -> None:
        super().__init__()
        self._dirty = False
        self._snapshot: tuple[Any, ...] = ()
        for link in links or []:
            self._attach(link)
            super().append(link)
        self.mark_clean()

    def append(self, item: ExternalLink) -> None:  # type: ignore[override]
        self._attach(item)
        super().append(item)
        self._mark_dirty()

    def extend(self, items: list[ExternalLink]) -> None:  # type: ignore[override]
        for item in items:
            self.append(item)

    def insert(self, index: int, item: ExternalLink) -> None:  # type: ignore[override]
        self._attach(item)
        super().insert(index, item)
        self._mark_dirty()

    def remove(self, item: ExternalLink) -> None:  # type: ignore[override]
        super().remove(item)
        item._collection = None
        self._mark_dirty()

    def pop(self, index: int = -1) -> ExternalLink:  # type: ignore[override]
        item = super().pop(index)
        item._collection = None
        self._mark_dirty()
        return item

    def clear(self) -> None:  # type: ignore[override]
        for item in self:
            item._collection = None
        super().clear()
        self._mark_dirty()

    def __setitem__(self, key: Any, value: Any) -> None:
        if isinstance(key, slice):
            values = list(value)
            for item in values:
                self._attach(item)
            for item in self[key]:
                item._collection = None
            super().__setitem__(key, values)
        else:
            self[key]._collection = None
            self._attach(value)
            super().__setitem__(key, value)
        self._mark_dirty()

    def __delitem__(self, key: Any) -> None:
        if isinstance(key, slice):
            for item in self[key]:
                item._collection = None
        else:
            self[key]._collection = None
        super().__delitem__(key)
        self._mark_dirty()

    @property
    def dirty(self) -> bool:
        return self._dirty or self._signature() != self._snapshot

    def mark_clean(self) -> None:
        self._snapshot = self._signature()
        for link in self:
            link._original_target = link.target
        self._dirty = False

    def _mark_dirty(self) -> None:
        self._dirty = True

    def _attach(self, item: ExternalLink) -> None:
        if not isinstance(item, ExternalLink):
            raise TypeError(
                f"external links collection requires ExternalLink, got {type(item).__name__}"
            )
        item._collection = self

    def _signature(self) -> tuple[Any, ...]:
        return tuple(link._signature() for link in self)


_install_openpyxl_iter(ExternalFileLink, ExternalLink)


def load_external_links(source_path: str | None) -> ExternalLinkCollection:
    """Load external-link parts from an xlsx at ``source_path``.

    Returns an empty list when:

    * ``source_path`` is ``None`` (workbook materialised from bytes with
      no tempfile path threaded through).
    * The xlsx has no ``xl/_rels/workbook.xml.rels`` (degenerate input).
    * No rel of type ``…/externalLink`` exists.

    Raises ``FileNotFoundError`` only on the hard case where
    ``source_path`` is set but the file is gone.
    """
    if not source_path:
        return ExternalLinkCollection()

    try:
        with zipfile.ZipFile(source_path, "r") as zf:
            validate_zipfile(zf)
            return ExternalLinkCollection(_load_from_zip(zf))
    except FileNotFoundError:
        raise
    except (zipfile.BadZipFile, KeyError, OSError, ValueError):
        # A malformed source ZIP shouldn't kill the whole load. The
        # patcher will surface a clearer error on save if the file is
        # actually unusable.
        return ExternalLinkCollection()


def load_external_links_from_bytes(data: bytes | mmap.mmap) -> ExternalLinkCollection:
    """Load external-link parts from a workbook's read-only source buffer."""
    from wolfxl._source_buffer import open_zip

    try:
        with open_zip(data) as zf:
            return ExternalLinkCollection(_load_from_zip(zf))
    except (zipfile.BadZipFile, KeyError, OSError):
        return ExternalLinkCollection()


def apply_authoring_to_xlsx(path: str, links: ExternalLinkCollection) -> None:
    """Rewrite workbook external-link parts to match ``links``."""
    if not links.dirty:
        return
    with zipfile.ZipFile(path, "r") as src:
        try:
            workbook_xml = src.read("xl/workbook.xml").decode("utf-8")
            workbook_rels_xml = src.read("xl/_rels/workbook.xml.rels")
            content_types_xml = src.read("[Content_Types].xml")
        except KeyError as exc:
            raise ValueError(f"workbook missing required OOXML part: {exc.args[0]}") from exc

    link_list = list(links)
    formula_replacements = _external_link_formula_replacements(link_list)
    rels_xml, rel_ids = _rewrite_workbook_rels(workbook_rels_xml, link_list)
    workbook_xml_bytes = _rewrite_workbook_external_references(workbook_xml, rel_ids)
    workbook_xml_bytes = _rewrite_external_formula_targets_xml(
        workbook_xml_bytes,
        formula_replacements,
    )
    content_types_bytes = _rewrite_content_types(content_types_xml, len(link_list))

    generated: dict[str, bytes] = {
        "xl/workbook.xml": workbook_xml_bytes,
        "xl/_rels/workbook.xml.rels": rels_xml,
        "[Content_Types].xml": content_types_bytes,
    }
    for idx, link in enumerate(link_list, start=1):
        generated[f"xl/externalLinks/externalLink{idx}.xml"] = _render_external_link_xml(link)
        generated[f"xl/externalLinks/_rels/externalLink{idx}.xml.rels"] = (
            _render_external_link_rels_xml(link)
        )

    fd, tmp_name = tempfile.mkstemp(prefix="wolfxl-extlinks-", suffix=".xlsx")
    os.close(fd)
    try:
        with zipfile.ZipFile(path, "r") as src, zipfile.ZipFile(
            tmp_name, "w", zipfile.ZIP_DEFLATED
        ) as dst:
            for info in src.infolist():
                name = info.filename
                if name.startswith("xl/externalLinks/") or name in generated:
                    continue
                with src.open(info, "r") as handle:
                    data = _rewrite_part_external_formula_targets(
                        name,
                        handle.read(),
                        formula_replacements,
                    )
                    dst.writestr(info, data)
            for name in sorted(generated):
                dst.writestr(name, generated[name])
        os.replace(tmp_name, path)
    finally:
        if os.path.exists(tmp_name):
            os.unlink(tmp_name)
    links.mark_clean()


def _load_from_zip(zf: zipfile.ZipFile) -> list[ExternalLink]:
    """Walk the workbook rels graph and parse every external link."""
    try:
        rels_bytes = read_entry(zf, "xl/_rels/workbook.xml.rels")
    except KeyError:
        return []

    rels = _parse_workbook_rels(rels_bytes)
    external_rels = [r for r in rels if r["type"] == _REL_TYPE_EXTERNAL_LINK]
    if not external_rels:
        return []

    out: list[ExternalLink] = []
    for rel in external_rels:
        # Target is relative to xl/_rels/workbook.xml.rels, so to xl/.
        target = rel["target"]
        part_path = _normalize_part_path("xl/", target)
        try:
            part_bytes = read_entry(zf, part_path)
        except KeyError:
            # Rel without a matching part — preserved on save by the
            # passthrough but not introspectable. Skip.
            continue

        parsed_part = _parse_part_xml(part_bytes)
        rels_path = _sibling_rels_path(part_path)
        try:
            part_rels_bytes = read_entry(zf, rels_path)
        except KeyError:
            part_rels_bytes = b""

        parsed_rels = _parse_part_rels_xml(part_rels_bytes)
        link_target = parsed_rels.get("target") or ""
        link_mode = parsed_rels.get("target_mode") or "External"

        file_link = ExternalFileLink(target=link_target, target_mode=link_mode)
        out.append(
            ExternalLink(
                file_link=file_link,
                rid=rel["id"],
                target=link_target,
                sheet_names=list(parsed_part.get("sheet_names", [])),
                cached_data=dict(parsed_part.get("cached_data", {})),
            )
        )

    return out


def _parse_workbook_rels(xml: bytes) -> list[dict[str, str]]:
    """Tiny rels parser specialised for the workbook rels graph.

    The Rust ``RelsGraph`` parser is overkill on the load path — we only
    need ``Id``, ``Type``, ``Target`` for one rel type. Using stdlib
    ElementTree keeps this hot path off the PyO3 boundary entirely for
    the empty case (no externalLink rels), which is the common case.
    """
    root = ET.fromstring(xml)
    out: list[dict[str, str]] = []
    for child in root:
        if not child.tag.endswith("}Relationship") and child.tag != "Relationship":
            continue
        out.append(
            {
                "id": child.attrib.get("Id", ""),
                "type": child.attrib.get("Type", ""),
                "target": child.attrib.get("Target", ""),
            }
        )
    return out


def _parse_part_xml(xml: bytes) -> dict[str, Any]:
    """Call the Rust parser, returning ``{}`` on import failure.

    The Rust extension is the production path; the import-failure
    fallback exists only so a half-built dev environment doesn't crash
    every load.
    """
    try:
        from wolfxl import _rust
    except ImportError:  # pragma: no cover - dev-only safety net
        return {"book_rid": None, "sheet_names": [], "cached_data": {}}
    return dict(_rust.parse_external_link_part(xml))


def _parse_part_rels_xml(xml: bytes) -> dict[str, Any]:
    """Call the Rust rels parser. Empty bytes -> empty dict."""
    if not xml:
        return {"target": None, "target_mode": None, "rid": None}
    try:
        from wolfxl import _rust
    except ImportError:  # pragma: no cover - dev-only safety net
        return {"target": None, "target_mode": None, "rid": None}
    return dict(_rust.parse_external_link_rels(xml))


def _normalize_part_path(base: str, target: str) -> str:
    """Resolve a rels target relative to ``base`` (`xl/`).

    Mirrors :func:`wolfxl._rust.ooxml_util::join_and_normalize` for the
    shapes external-link rels actually ship with: relative
    (``externalLinks/externalLink1.xml``) and parent-anchored
    (``../foo/x.xml``). Absolute targets (``/xl/...``) lose the leading
    slash and are taken as-is from the package root.
    """
    if target.startswith("/"):
        return target.lstrip("/")
    parts = (base + target).split("/")
    stack: list[str] = []
    for p in parts:
        if p == "" or p == ".":
            continue
        if p == "..":
            if stack:
                stack.pop()
            continue
        stack.append(p)
    return "/".join(stack)


def _sibling_rels_path(part_path: str) -> str:
    """For ``xl/externalLinks/externalLink1.xml`` return
    ``xl/externalLinks/_rels/externalLink1.xml.rels``."""
    if "/" not in part_path:
        return f"_rels/{part_path}.rels"
    parent, name = part_path.rsplit("/", 1)
    return f"{parent}/_rels/{name}.rels"


def _rewrite_workbook_rels(xml: bytes, links: list[ExternalLink]) -> tuple[bytes, list[str]]:
    ET.register_namespace("", _RELS_NS)
    root = ET.fromstring(xml)
    kept = [
        child
        for child in list(root)
        if child.attrib.get("Type") != _REL_TYPE_EXTERNAL_LINK
    ]
    for child in list(root):
        root.remove(child)
    for child in kept:
        root.append(child)

    max_rid = 0
    for child in root:
        rid = child.attrib.get("Id", "")
        if rid.startswith("rId") and rid[3:].isdigit():
            max_rid = max(max_rid, int(rid[3:]))

    rel_ids: list[str] = []
    for idx, _link in enumerate(links, start=1):
        max_rid += 1
        rid = f"rId{max_rid}"
        rel_ids.append(rid)
        ET.SubElement(
            root,
            f"{{{_RELS_NS}}}Relationship",
            {
                "Id": rid,
                "Type": _REL_TYPE_EXTERNAL_LINK,
                "Target": f"externalLinks/externalLink{idx}.xml",
            },
        )
    return _xml_bytes(root), rel_ids


def _rewrite_content_types(xml: bytes, count: int) -> bytes:
    ns = "http://schemas.openxmlformats.org/package/2006/content-types"
    ET.register_namespace("", ns)
    root = ET.fromstring(xml)
    for child in list(root):
        part = child.attrib.get("PartName", "")
        if part.startswith("/xl/externalLinks/externalLink") and part.endswith(".xml"):
            root.remove(child)
    for idx in range(1, count + 1):
        ET.SubElement(
            root,
            f"{{{ns}}}Override",
            {
                "PartName": f"/xl/externalLinks/externalLink{idx}.xml",
                "ContentType": _EXTERNAL_LINK_CT,
            },
        )
    return _xml_bytes(root)


def _rewrite_workbook_external_references(workbook_xml: str, rel_ids: list[str]) -> bytes:
    workbook_xml = re.sub(
        r"<(?:\w+:)?externalReferences\b[^>]*>.*?</(?:\w+:)?externalReferences>",
        "",
        workbook_xml,
        flags=re.DOTALL,
    )
    if rel_ids and "xmlns:r=" not in workbook_xml:
        workbook_xml = workbook_xml.replace(
            "<workbook ",
            f'<workbook xmlns:r="{_REL_TYPE_NS}" ',
            1,
        )
    if rel_ids:
        refs = "<externalReferences>" + "".join(
            f'<externalReference r:id="{_xml_attr_escape(rid)}"/>' for rid in rel_ids
        ) + "</externalReferences>"
        if "</sheets>" in workbook_xml:
            workbook_xml = workbook_xml.replace("</sheets>", f"</sheets>{refs}", 1)
        else:
            workbook_xml = workbook_xml.replace("</workbook>", f"{refs}</workbook>", 1)
    return workbook_xml.encode("utf-8")


def _render_external_link_xml(link: ExternalLink) -> bytes:
    sheet_names = "".join(
        f'<sheetName val="{_xml_attr_escape(name)}"/>' for name in link.sheet_names
    )
    sheet_names_block = f"<sheetNames>{sheet_names}</sheetNames>" if sheet_names else ""
    cached_block = _render_cached_data(link.cached_data)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<externalLink xmlns="{_MAIN_NS}" xmlns:r="{_REL_TYPE_NS}">'
        f'<externalBook r:id="rId1">{sheet_names_block}{cached_block}</externalBook>'
        "</externalLink>"
    ).encode("utf-8")


def _render_external_link_rels_xml(link: ExternalLink) -> bytes:
    target = link.target
    target_mode = link.file_link.target_mode if link.file_link is not None else "External"
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{_RELS_NS}">'
        f'<Relationship Id="rId1" Type="{_REL_TYPE_EXTERNAL_LINK_PATH}" '
        f'Target="{_xml_attr_escape(target)}" TargetMode="{_xml_attr_escape(target_mode)}"/>'
        "</Relationships>"
    ).encode("utf-8")


def _external_link_formula_replacements(
    links: list[ExternalLink],
) -> list[tuple[bytes, bytes]]:
    replacements: list[tuple[bytes, bytes]] = []
    seen: set[tuple[str, str]] = set()
    for link in links:
        old_target = link._original_target
        new_target = link.target
        if not old_target or old_target == new_target:
            continue
        old_book = _formula_workbook_name(old_target)
        new_book = _formula_workbook_name(new_target)
        if not old_book or not new_book or old_book == new_book:
            continue
        pair = (old_book, new_book)
        if pair in seen:
            continue
        seen.add(pair)
        replacements.append(
            (
                f"[{_xml_text_escape(old_book)}]".encode("utf-8"),
                f"[{_xml_text_escape(new_book)}]".encode("utf-8"),
            )
        )
    return replacements


def _formula_workbook_name(target: str) -> str:
    target = target.replace("\\", "/").rstrip("/")
    if not target:
        return ""
    return target.rsplit("/", 1)[-1]


def _rewrite_part_external_formula_targets(
    name: str,
    data: bytes,
    replacements: list[tuple[bytes, bytes]],
) -> bytes:
    if not replacements or not name.endswith(".xml"):
        return data
    if not name.startswith(_FORMULA_PART_PREFIXES):
        return data
    return _rewrite_external_formula_targets_xml(data, replacements)


def _rewrite_external_formula_targets_xml(
    xml: bytes,
    replacements: list[tuple[bytes, bytes]],
) -> bytes:
    if not replacements:
        return xml
    replacement_map = dict(replacements)
    token_re = re.compile(
        b"|".join(re.escape(old) for old in sorted(replacement_map, key=len, reverse=True))
    )

    def rewrite_match(match: re.Match[bytes]) -> bytes:
        text = match.group(2)
        text = token_re.sub(lambda token: replacement_map[token.group(0)], text)
        return match.group(1) + text + match.group(3)

    return _FORMULA_TAG_RE.sub(rewrite_match, xml)


def _render_cached_data(cached_data: dict[str, list[dict[str, str]]]) -> str:
    if not cached_data:
        return ""
    sheets: list[str] = []
    for sheet_id, cells in cached_data.items():
        rows: dict[str, list[dict[str, str]]] = {}
        for cell in cells:
            ref = str(cell.get("r", ""))
            row = "".join(ch for ch in ref if ch.isdigit()) or "1"
            rows.setdefault(row, []).append(cell)
        row_xml = ""
        for row_num, row_cells in rows.items():
            cell_xml = "".join(
                f'<cell r="{_xml_attr_escape(str(cell.get("r", "")))}"><v>{_xml_text_escape(str(cell.get("v", "")))}</v></cell>'
                for cell in row_cells
            )
            row_xml += f'<row r="{_xml_attr_escape(row_num)}">{cell_xml}</row>'
        sheets.append(f'<sheetData sheetId="{_xml_attr_escape(str(sheet_id))}">{row_xml}</sheetData>')
    return "<sheetDataSet>" + "".join(sheets) + "</sheetDataSet>"


def _freeze_cached_data(value: Any) -> Any:
    if isinstance(value, dict):
        return tuple(sorted((k, _freeze_cached_data(v)) for k, v in value.items()))
    if isinstance(value, list):
        return tuple(_freeze_cached_data(v) for v in value)
    return value


def _xml_bytes(root: ET.Element) -> bytes:
    return (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        + ET.tostring(root, encoding="utf-8", short_empty_elements=True)
    )


def _xml_attr_escape(value: str) -> str:
    return (
        value.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def _xml_text_escape(value: str) -> str:
    return value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
