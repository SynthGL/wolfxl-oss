"""Extended document properties compatibility."""

from __future__ import annotations

import base64
import binascii
from typing import Any

from wolfxl import __version__
from wolfxl._compat import _OpenpyxlSerialisable, _iter_openpyxl_attrs
from wolfxl.xml.constants import XPROPS_NS
from wolfxl.xml.functions import Element, SubElement, localname


class _TypedVector(list):
    def __init__(self, values=()):
        values = list(values)
        self._validate(values)
        super().__init__(values)

    def _validate(self, values) -> None:
        for value in values:
            if type(value) not in self._types:
                raise TypeError(f"invalid {type(self).__name__} entry")
            if type(value) is int and not -(2**31) <= value < 2**31:
                raise ValueError("vector integer must fit an OOXML signed 32-bit value")

    def append(self, value) -> None:
        self._validate((value,))
        super().append(value)

    def extend(self, values) -> None:
        values = list(values)
        self._validate(values)
        super().extend(values)

    def insert(self, index, value) -> None:
        self._validate((value,))
        super().insert(index, value)

    def __setitem__(self, index, value) -> None:
        values = list(value) if isinstance(index, slice) else [value]
        self._validate(values)
        super().__setitem__(index, values if isinstance(index, slice) else value)

    def __iadd__(self, values):
        self.extend(values)
        return self


class VectorLpstr(_TypedVector):
    _types = (str,)


class VectorVariant(_TypedVector):
    _types = (str, int, bool)


class DigSigBlob(_OpenpyxlSerialisable):
    """Base64 signature value stored in an extended-property ``vt:blob``."""

    def __init__(self, blob: str = "") -> None:
        self.blob = blob

    def __setattr__(self, name: str, value: Any) -> None:
        if name == "blob":
            if not isinstance(value, str):
                raise TypeError("DigSig blob must be a base64 string")
            try:
                base64.b64decode(value, validate=True)
            except (binascii.Error, ValueError) as exc:
                raise ValueError("DigSig blob must be valid base64") from exc
        object.__setattr__(self, name, value)

    def _validate(self) -> None:
        self.blob = self.blob


class ExtendedProperties:
    """Typed ``docProps/app.xml`` values; assign ``None`` to remove a member.

    Use ``VectorVariant`` for HeadingPairs and HLinks, ``VectorLpstr`` for
    TitlesOfParts, and ``DigSigBlob`` for the base64 document signature.
    Worksheet title vectors follow the workbook's sheet graph on save.
    """

    __elements__ = (
        "Template",
        "Manager",
        "Company",
        "Pages",
        "Words",
        "Characters",
        "PresentationFormat",
        "Lines",
        "Paragraphs",
        "Slides",
        "Notes",
        "TotalTime",
        "HiddenSlides",
        "MMClips",
        "ScaleCrop",
        "HeadingPairs",
        "TitlesOfParts",
        "LinksUpToDate",
        "CharactersWithSpaces",
        "SharedDoc",
        "HyperlinkBase",
        "HLinks",
        "HyperlinksChanged",
        "DigSig",
        "Application",
        "AppVersion",
        "DocSecurity",
    )
    _integers = frozenset(
        {
            "Pages",
            "Words",
            "Characters",
            "Lines",
            "Paragraphs",
            "Slides",
            "Notes",
            "TotalTime",
            "HiddenSlides",
            "MMClips",
            "CharactersWithSpaces",
            "DocSecurity",
        }
    )
    _booleans = frozenset(
        {
            "ScaleCrop",
            "LinksUpToDate",
            "SharedDoc",
            "HyperlinksChanged",
        }
    )

    def __init__(
        self,
        Application: str | None = None,  # noqa: N803
        AppVersion: str | None = None,  # noqa: N803
        Company: str | None = None,  # noqa: N803
        Manager: str | None = None,  # noqa: N803
        HyperlinksChanged: bool | None = None,  # noqa: N803
        SharedDoc: bool | None = None,  # noqa: N803
        LinksUpToDate: bool | None = None,  # noqa: N803
        ScaleCrop: bool | None = None,  # noqa: N803
        HeadingPairs: Any = None,  # noqa: N803
        TitlesOfParts: Any = None,  # noqa: N803
        DocSecurity: int | None = None,  # noqa: N803
        **kw: Any,
    ) -> None:
        for name in self.__elements__:
            object.__setattr__(self, name, None)
        self.Application = f"Microsoft Excel Compatible / Openpyxl {__version__}"
        self.AppVersion = ".".join(__version__.split(".")[:-1])
        self.Company = Company
        self.Manager = Manager
        self.HyperlinksChanged = HyperlinksChanged
        self.SharedDoc = SharedDoc
        self.LinksUpToDate = LinksUpToDate
        self.ScaleCrop = ScaleCrop
        self.HeadingPairs = HeadingPairs
        self.TitlesOfParts = TitlesOfParts
        self.DocSecurity = DocSecurity
        if Application is not None:
            self.Application = Application
        if AppVersion is not None:
            self.AppVersion = AppVersion
        for name, value in kw.items():
            setattr(self, name, value)

    def __iter__(self):
        yield from _iter_openpyxl_attrs(self)

    def __setattr__(self, name: str, value: Any) -> None:
        if name in self.__elements__:
            self._validate(name, value)
        elif not name.startswith("_"):
            raise AttributeError(f"unknown extended property {name!r}")
        object.__setattr__(self, name, value)
        if name in self.__elements__ and "_user_set" in self.__dict__:
            self._user_set.add(name)

    @classmethod
    def _validate(cls, name: str, value: Any) -> None:
        if value is None:
            return
        if name in cls._integers:
            valid = type(value) is int
            if valid and not -(2**31) <= value < 2**31:
                raise ValueError(f"{name} must fit an OOXML signed 32-bit integer")
        elif name in cls._booleans:
            valid = type(value) is bool
        elif name in ("HeadingPairs", "HLinks"):
            valid = isinstance(value, VectorVariant) and all(
                type(v) in (str, int, bool) for v in value
            )
            if valid and any(type(v) is int and not -(2**31) <= v < 2**31 for v in value):
                raise ValueError(f"{name} vector integers must fit signed 32-bit values")
        elif name == "TitlesOfParts":
            valid = isinstance(value, VectorLpstr) and all(isinstance(v, str) for v in value)
        elif name == "DigSig":
            valid = isinstance(value, DigSigBlob)
            if valid:
                value._validate()
        else:
            valid = isinstance(value, str)
        if not valid:
            raise TypeError(f"invalid value for extended property {name}")

    def _attach_workbook(self, wb: Any) -> None:
        object.__setattr__(self, "_wb", wb)
        object.__setattr__(self, "_user_set", set())
        object.__setattr__(self, "_baseline", self._member_values())

    def _member_values(self) -> dict[str, Any]:
        return {
            name: tuple(value)
            if isinstance(value, list)
            else value.blob
            if isinstance(value, DigSigBlob)
            else value
            for name in self.__elements__
            for value in (getattr(self, name),)
        }

    def _changed_fields(self) -> set[str]:
        current = self._member_values()
        baseline = getattr(self, "_baseline", {})
        return self._user_set | {
            name for name in self.__elements__ if current.get(name) != baseline.get(name)
        }

    def to_tree(self, namespace: str = XPROPS_NS, variant_namespace: str | None = None) -> Any:
        vt = variant_namespace or (
            "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
        )
        root = Element(f"{{{namespace}}}Properties")
        for name in self.__elements__:
            value = getattr(self, name, None)
            self._validate(name, value)
            if value is None:
                continue
            child = SubElement(root, f"{{{namespace}}}{name}")
            if name == "DigSig":
                SubElement(child, f"{{{vt}}}blob").text = value.blob
            elif name in ("HeadingPairs", "TitlesOfParts", "HLinks"):
                variants = name != "TitlesOfParts"
                vector = SubElement(
                    child,
                    f"{{{vt}}}vector",
                    {
                        "size": str(len(value)),
                        "baseType": "variant" if variants else "lpstr",
                    },
                )
                for entry in value:
                    parent = SubElement(vector, f"{{{vt}}}variant") if variants else vector
                    kind = (
                        "bool"
                        if type(entry) is bool
                        else "i4"
                        if type(entry) is int
                        else "lpwstr"
                        if name == "HLinks"
                        else "lpstr"
                    )
                    SubElement(parent, f"{{{vt}}}{kind}").text = (
                        str(entry).lower() if type(entry) is bool else str(entry)
                    )
            else:
                child.text = str(value).lower() if type(value) is bool else str(value)
        return root

    @classmethod
    def from_tree(cls, node: Any) -> "ExtendedProperties":
        props = cls()
        # Missing members remain absent, including Application and AppVersion.
        for name in cls.__elements__:
            object.__setattr__(props, name, None)
        for child in node:
            name = localname(child)
            if name not in cls.__elements__:
                continue
            if name in cls._integers:
                value = int(child.text or "")
            elif name in cls._booleans:
                if child.text not in ("true", "false", "1", "0"):
                    raise ValueError(f"invalid OOXML boolean for {name}")
                value = child.text in ("true", "1")
            elif name == "DigSig":
                blob = next(iter(child), None)
                value = DigSigBlob(blob.text or "") if blob is not None else None
            elif name in ("HeadingPairs", "TitlesOfParts", "HLinks"):
                vector = next(iter(child), None)
                entries = [] if vector is None else list(vector)
                values = []
                for entry in entries:
                    if localname(entry) == "variant":
                        entry = next(iter(entry), None)
                    if entry is None:
                        raise ValueError(f"empty variant in {name}")
                    kind = localname(entry)
                    if kind == "i4":
                        values.append(int(entry.text or ""))
                    elif kind == "bool":
                        if entry.text not in ("true", "false", "1", "0"):
                            raise ValueError(f"invalid OOXML boolean in {name}")
                        values.append(entry.text in ("true", "1"))
                    elif kind in ("lpstr", "lpwstr"):
                        values.append(entry.text or "")
                    else:
                        raise ValueError(f"unsupported extended property variant {kind}")
                value = VectorLpstr(values) if name == "TitlesOfParts" else VectorVariant(values)
            else:
                value = child.text or ""
            setattr(props, name, value)
        return props


NestedText = Serialisable = Typed = _OpenpyxlSerialisable

__all__ = [
    "DigSigBlob",
    "ExtendedProperties",
    "NestedText",
    "Serialisable",
    "Typed",
    "VectorLpstr",
    "VectorVariant",
    "XPROPS_NS",
]
