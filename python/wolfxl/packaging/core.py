"""openpyxl.packaging.core.DocumentProperties compatibility.

Matches the openpyxl shape used for ``wb.properties.title`` etc.
Dates are ``datetime`` objects on the Python side; the Rust layer
delivers them as ISO 8601 strings. Missing fields are ``None``.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import Any

from wolfxl._compat import _OpenpyxlSerialisable, _iter_openpyxl_attrs, _openpyxl_name_fallback
from wolfxl.xml.constants import COREPROPS_NS, DCORE_NS, DCTERMS_NS, XSI_NS
from wolfxl.xml.functions import Element, QName, localname

_DC_FIELDS = {"creator", "title", "description", "subject", "identifier", "language"}
_DCTERMS_FIELDS = {"created", "modified"}


@dataclass
class DocumentProperties:
    """Workbook-level metadata (``docProps/core.xml`` + ``docProps/app.xml``).

    Construction with all defaults produces an "empty" properties object
    — every field is ``None``. This is the shape returned for a fresh
    ``Workbook()`` in write mode (no file to read from yet) and for a
    workbook whose metadata file was missing or malformed.

    In-place attribute assignments (``wb.properties.title = "X"``) flag
    the owning workbook's ``_properties_dirty`` when the
    :meth:`_attach_workbook` helper is called — that lets
    :meth:`Workbook.save` distinguish between "untouched" and "user
    mutated these" across both write and modify modes.
    """

    title: str | None = None
    subject: str | None = None
    creator: str | None = None
    keywords: str | None = None
    description: str | None = None
    lastModifiedBy: str | None = None  # noqa: N815 - openpyxl public API
    category: str | None = None
    contentStatus: str | None = None  # noqa: N815
    identifier: str | None = None
    language: str | None = None
    revision: str | None = None
    version: str | None = None
    created: datetime | None = None
    modified: datetime | None = None
    lastPrinted: datetime | None = None  # noqa: N815

    def __iter__(self):
        yield from _iter_openpyxl_attrs(self)

    def to_tree(self) -> Any:
        try:
            root = Element(
                f"{{{COREPROPS_NS}}}coreProperties",
                nsmap={
                    None: COREPROPS_NS,
                    "dc": DCORE_NS,
                    "dcterms": DCTERMS_NS,
                    "xsi": XSI_NS,
                },
            )
        except TypeError:
            root = Element(f"{{{COREPROPS_NS}}}coreProperties")
        for name in _ELEMENTS:
            value = getattr(self, name, None)
            if value is None:
                continue
            if name in _DCTERMS_FIELDS:
                root.append(QualifiedDateTime().to_tree(name, value, DCTERMS_NS))
                continue
            namespace = DCORE_NS if name in _DC_FIELDS else COREPROPS_NS
            child = Element(f"{{{namespace}}}{name}")
            if isinstance(value, datetime):
                child.text = _to_w3cdtf(value)
            else:
                child.text = str(value)
            root.append(child)
        return root

    @classmethod
    def from_tree(cls, node: Any) -> "DocumentProperties":
        kwargs = {}
        for child in node:
            name = localname(child)
            if name not in _ELEMENTS:
                continue
            if name in {"created", "modified", "lastPrinted"}:
                kwargs[name] = _parse_dt(child.text)
            else:
                kwargs[name] = child.text
        return cls(**kwargs)

    # Dirty-tracking support. Set via _attach_workbook after construction;
    # the dataclass's own __init__ uses __setattr__, so we only enable
    # the tracking hook once _wb is present.
    def __setattr__(self, name: str, value: Any) -> None:
        object.__setattr__(self, name, value)
        wb = self.__dict__.get("_wb")
        if wb is not None and name != "_wb":
            wb._properties_dirty = True  # noqa: SLF001
            # Track per-field user mutations so the modify-mode flush
            # can distinguish "user explicitly set X" from "X was
            # hydrated from the source on first read". Required for
            # ``modified``: a dirty save re-stamps it to save-time
            # unless the user supplied a specific datetime.
            user_set: set[str] = self.__dict__.setdefault("_user_set", set())
            user_set.add(name)

    def _attach_workbook(self, wb: Any) -> None:
        """Link this properties object to its owning Workbook.

        After this call, every subsequent attribute assignment flips
        ``wb._properties_dirty = True`` — transparent to the user.
        Also resets the per-field user-mutation set so cache hydration
        prior to attach is not counted as a user mutation.
        """
        object.__setattr__(self, "_wb", wb)
        object.__setattr__(self, "_user_set", set())


def _doc_props_from_dict(raw: dict[str, Any] | None) -> DocumentProperties:
    """Build a ``DocumentProperties`` from the Rust reader's dict output.

    The Rust side emits every field as a string (or omits it); we parse
    the two datetime fields here. A malformed ``created`` / ``modified``
    string collapses to ``None`` rather than raising, so a corrupt
    sidecar can't break opening the workbook.
    """
    raw = raw or {}

    def _parse_dt(value: Any) -> datetime | None:
        if value is None:
            return None
        if isinstance(value, datetime):
            return value
        if not isinstance(value, str) or not value:
            return None
        # Python 3.11+ accepts OOXML's ISO 8601 trailing Z directly.
        try:
            return datetime.fromisoformat(value)
        except ValueError:
            return None

    return DocumentProperties(
        title=raw.get("title"),
        subject=raw.get("subject"),
        creator=raw.get("creator"),
        keywords=raw.get("keywords"),
        description=raw.get("description"),
        lastModifiedBy=raw.get("lastModifiedBy"),
        category=raw.get("category"),
        contentStatus=raw.get("contentStatus"),
        identifier=raw.get("identifier"),
        language=raw.get("language"),
        revision=raw.get("revision"),
        version=raw.get("version"),
        created=_parse_dt(raw.get("created")),
        modified=_parse_dt(raw.get("modified")),
    )


class NestedDateTime(_OpenpyxlSerialisable):
    def to_tree(
        self,
        tagname: str | None = None,
        value: datetime | None = None,
        namespace: str | None = None,
    ) -> Any:
        tag = tagname or "date"
        if namespace is not None:
            tag = f"{{{namespace}}}{tag}"
        node = Element(tag)
        if value is not None:
            node.text = _to_w3cdtf(value)
        return node


class QualifiedDateTime(NestedDateTime):
    def to_tree(
        self,
        tagname: str | None = None,
        value: datetime | None = None,
        namespace: str | None = None,
    ) -> Any:
        node = super().to_tree(tagname, value, namespace)
        node.set(f"{{{XSI_NS}}}type", QName(DCTERMS_NS, "W3CDTF"))
        return node


NestedText = Serialisable = DateTime = Alias = _OpenpyxlSerialisable

_ELEMENTS = (
    "creator",
    "title",
    "description",
    "subject",
    "identifier",
    "language",
    "created",
    "modified",
    "lastModifiedBy",
    "category",
    "contentStatus",
    "version",
    "revision",
    "keywords",
    "lastPrinted",
)


def _to_w3cdtf(value: datetime) -> str:
    if value.tzinfo is not None:
        value = value.replace(tzinfo=None)
    return value.replace(microsecond=0).isoformat() + "Z"


def _parse_dt(value: Any) -> datetime | None:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value
    if not isinstance(value, str) or not value:
        return None
    try:
        return datetime.fromisoformat(value.removesuffix("Z"))
    except ValueError:
        return None


__all__ = [
    "Alias",
    "COREPROPS_NS",
    "DCORE_NS",
    "DCTERMS_NS",
    "DateTime",
    "DocumentProperties",
    "Element",
    "NestedDateTime",
    "NestedText",
    "QName",
    "QualifiedDateTime",
    "Serialisable",
    "XSI_NS",
    "_doc_props_from_dict",
    "datetime",
]

__getattr__ = _openpyxl_name_fallback(globals())
