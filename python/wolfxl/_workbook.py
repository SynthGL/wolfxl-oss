"""Workbook — multi-mode openpyxl-compatible wrapper.

Write mode (``Workbook()``): creates a new workbook via NativeWorkbook.
Read mode (``Workbook._from_reader(path)``): opens an existing .xlsx via NativeXlsxBook by default.
Modify mode (``Workbook._from_patcher(path)``): read via NativeXlsxBook by default, save via XlsxPatcher.
"""

from __future__ import annotations

import os
import inspect as _inspect
import zipfile
from xml.etree import ElementTree as ET
from typing import TYPE_CHECKING, Any

from wolfxl._workbook_state import CopyOptions as CopyOptions
from wolfxl._workbook_state import initialize_pending_state
from wolfxl._zip_safety import read_entry, validate_zipfile
from wolfxl import _workbook_features
from wolfxl import _workbook_calc
from wolfxl import _workbook_metadata
from wolfxl import _workbook_lifecycle
from wolfxl import _workbook_patcher_flush
from wolfxl import _workbook_save
from wolfxl import _workbook_sheets
from wolfxl import _workbook_sources
from wolfxl import _workbook_writer_flush
from wolfxl._worksheet import Worksheet
from wolfxl.utils.exceptions import ReadOnlyWorkbookException
from wolfxl.xml.constants import XLSM, XLSX, XLTM, XLTX


if TYPE_CHECKING:
    from wolfxl.calc._protocol import RecalcResult


class Workbook:
    """Openpyxl-compatible workbook backed by Rust.

    A workbook operates in one of three modes:

    * write mode, created with ``Workbook()``;
    * read mode, created with ``load_workbook(path)``;
    * modify mode, created with ``load_workbook(path, modify=True)``.

    Public methods mirror openpyxl where practical while using WolfXL's native
    Excel I/O engine for fast reads, writes, and preserving modify-mode saves.
    """

    def __init__(self, write_only: bool = False, iso_dates: bool = False) -> None:
        """Create a new workbook in write mode.

        Args:
            write_only: When ``True``, opt into openpyxl-style streaming
                write-only mode (G20 / RFC-073). Each
                :meth:`create_sheet` call returns a
                :class:`WriteOnlyWorksheet` that streams rows directly
                to a per-sheet temp file; the eager BTreeMap row
                accumulation is bypassed, so 10M-row exports run with
                bounded peak memory. The default sheet ``"Sheet"`` is
                NOT created — callers must explicitly create sheets
                before appending. A workbook saved in write-only mode
                is consumed-on-save: a second :meth:`save` raises
                :class:`WorkbookAlreadySaved`.
        """
        from wolfxl import _backend, _rust  # noqa: F401  (_rust kept for typing parity)

        self._rust_writer: Any = _backend.make_writer()
        self._rust_reader: Any = None
        self._rust_patcher: Any = None
        self._data_only = False
        self._iso_dates = bool(iso_dates)
        self.template = False
        self._is_template = False
        self.encoding = "utf-8"
        # Flipped to True via load_workbook(rich_text=True).
        self._rich_text: bool = False
        self._evaluator: Any = None
        # Last calculate() snapshot: (inputs, defined names, date1904) and its
        # results, persisted as formula caches by the next save.
        self._calculated_state: Any = None
        self._calculated_formula_results: dict[str, Any] = {}
        # G20: streaming write-only mode. When set, create_sheet returns
        # WriteOnlyWorksheet instances and the default sheet is skipped.
        self._write_only: bool = bool(write_only)
        # Re-entry guard for openpyxl-shape write-only consumed-on-save semantic.
        self._saved: bool = False
        if self._write_only:
            self._sheet_names: list[str] = []
            self._sheets: dict[str, Worksheet] = {}
        else:
            self._sheet_names = ["Sheet"]
            self._sheets = {}
            self._sheets["Sheet"] = Worksheet(self, "Sheet")
            self._rust_writer.add_sheet("Sheet")
        initialize_pending_state(self)
        self._active_sheet_index = 0
        # Streaming read flag (write mode never streams).
        self._read_only: bool = False
        self._source_path: str | None = None
        self._keep_links: bool = True
        self._keep_vba: bool = False
        # File format the workbook came from. Write
        # mode is xlsx by definition; the read/modify constructors set
        # this to "xlsx" / "xlsb" / "xls" as appropriate.
        self._format: str = "xlsx"

    @classmethod
    def _from_reader(
        cls,
        path: str,
        *,
        data_only: bool = False,
        permissive: bool = False,
        read_only: bool = False,
    ) -> Workbook:
        """Open an existing .xlsx file in read mode.

        ``permissive`` plumbs through to the Rust reader and triggers a
        rels-graph fallback when ``<sheets>`` is empty/self-closing.

        ``read_only`` activates the SAX streaming fast path on ``iter_rows``.
        The workbook still keeps a Rust reader for style/format lookups, while
        the streaming reader bypasses eager materialization for the large-sheet
        scan path. See
        :func:`wolfxl.load_workbook` for details.
        """
        return _workbook_sources.from_reader(
            cls,
            path,
            data_only=data_only,
            permissive=permissive,
            read_only=read_only,
        )

    @classmethod
    def _from_encrypted(
        cls,
        path: str | None = None,
        *,
        data: bytes | bytearray | memoryview | None = None,
        password: str | bytes,
        data_only: bool = False,
        permissive: bool = False,
        modify: bool = False,
        read_only: bool = False,
    ) -> Workbook:
        """Open an OOXML-encrypted .xlsx via msoffcrypto-tool.

        Decrypts the source (path or in-memory blob) into an in-memory
        buffer, then dispatches through the bytes-aware reader path. On a
        non-encrypted file the password is silently ignored and the normal
        path is used (matches openpyxl).

        Wrong / missing passwords raise ``ValueError`` with a clear
        message; ``ImportError`` (with install hint) surfaces when
        ``msoffcrypto-tool`` isn't installed.

        Modify mode + password works because the decrypted bytes are
        rematerialised through ``_from_bytes``; on save the result is
        plaintext (write-side encryption is documented T3
        out-of-scope).

        Exactly one of ``path`` / ``data`` must be supplied — the
        ``load_workbook`` dispatcher threads whichever the caller passed in.
        """
        return _workbook_sources.from_encrypted(
            cls,
            path,
            data=data,
            password=password,
            data_only=data_only,
            permissive=permissive,
            modify=modify,
            read_only=read_only,
        )

    @classmethod
    def _from_bytes(
        cls,
        data: bytes | bytearray | memoryview,
        *,
        data_only: bool = False,
        permissive: bool = False,
        modify: bool = False,
        read_only: bool = False,
    ) -> Workbook:
        """Open an .xlsx blob from memory.

        When the underlying Rust reader exposes ``open_from_bytes``
        the blob is handed to the reader directly with no intermediate
        tempfile. Otherwise the bytes are materialised to a tempfile and
        the path-based reader / patcher is used; the tempfile is tracked on
        the workbook so :meth:`close` can clean it up. Either way the
        public surface (``Workbook._format``, ``Workbook._source_path``,
        etc.) is identical.

        ``read_only`` plumbs through to the streaming SAX path;
        ``modify=True`` always uses a tempfile because the XlsxPatcher is
        path-only by design (it reopens the source zip on save).
        """
        return _workbook_sources.from_bytes(
            cls,
            data,
            data_only=data_only,
            permissive=permissive,
            modify=modify,
            read_only=read_only,
        )

    @classmethod
    def _from_patcher(
        cls,
        path: str,
        *,
        data_only: bool = False,
        permissive: bool = False,
    ) -> Workbook:
        """Open an existing .xlsx file in modify mode (read + surgical save).

        ``permissive`` plumbs through to both the reader and the
        patcher; see :func:`wolfxl.load_workbook` for the user-facing
        contract.
        """
        return _workbook_sources.from_patcher(
            cls,
            path,
            data_only=data_only,
            permissive=permissive,
        )

    @classmethod
    def _from_xlsb(
        cls,
        *,
        path: str | None,
        data: bytes | None,
        data_only: bool = False,
        permissive: bool = False,
    ) -> Workbook:
        """Open an .xlsb workbook via the native BIFF12 reader.

        xlsb is a binary OOXML container. WolfXL surfaces values,
        cached formula results, and read-side cell styles; modify mode,
        streaming, password reads, and writes remain xlsx-only.
        """
        return _workbook_sources.from_xlsb(
            cls,
            path=path,
            data=data,
            data_only=data_only,
            permissive=permissive,
        )

    @classmethod
    def _from_xls(
        cls,
        *,
        path: str | None,
        data: bytes | None,
        data_only: bool = False,
        permissive: bool = False,
    ) -> Workbook:
        """Open a legacy .xls workbook via ``NativeXlsBook``.

        Same shape as :meth:`_from_xlsb` — values + cached formula
        results only.
        """
        return _workbook_sources.from_xls(
            cls,
            path=path,
            data=data,
            data_only=data_only,
            permissive=permissive,
        )

    # ------------------------------------------------------------------
    # Sheet access
    # ------------------------------------------------------------------

    @property
    def sheetnames(self) -> list[str]:
        """Return worksheet titles in tab order."""
        return list(self._sheet_names)

    @property
    def worksheets(self) -> list[Worksheet]:
        """List of Worksheet objects in sheet order — openpyxl alias."""
        return [self._sheets[name] for name in self._sheet_names if name in self._sheets]

    @property
    def active(self) -> Worksheet | None:
        """Return the active worksheet, or None if no worksheets exist."""
        index = self._active_index()
        try:
            name = self._sheet_names[index]
        except IndexError:
            return None
        if name in self._sheets:
            return self._sheets[name]
        chartsheets = getattr(self, "_chartsheets", {})
        if name in chartsheets:
            return chartsheets[name]
        return None

    @active.setter
    def active(self, value: int | Worksheet) -> None:
        """Set the active worksheet by index or worksheet object."""
        if isinstance(value, int):
            index = value
        elif isinstance(value, Worksheet):
            if (
                getattr(value, "_workbook", None) is not self
                or value.title not in self._sheets
                or self._sheets[value.title] is not value
            ):
                raise ValueError("Worksheet is not in the workbook")
            if value.sheet_state != "visible":
                raise ValueError("Only visible sheets can be made active")
            index = self.index(value)
        else:
            raise TypeError("Workbook.active must be a worksheet index or Worksheet")
        self._active_sheet_index = index
        self._active_sheet_dirty = True
        if self._views_cache is not None and self._views_cache:
            self._views_cache[0].activeTab = index

    def _active_index(self) -> int:
        """Return the raw workbook active-tab index."""
        index = getattr(self, "_active_sheet_index", None)
        if index is None:
            views = _workbook_metadata.get_views(self)
            if views:
                index = int(getattr(views[0], "activeTab", 0) or 0)
                self._active_sheet_index = index
            else:
                index = 0
        if not self._sheet_names:
            return 0
        return int(index)

    def _active_index_for_save(self) -> int | None:
        """Return openpyxl-compatible active-tab value for workbook.xml."""
        visible_indices = [
            idx
            for idx, name in enumerate(self._sheet_names)
            if getattr(self[name], "sheet_state", "visible") == "visible"
        ]
        if not visible_indices:
            raise IndexError("At least one sheet must be visible")

        index = self._active_index()
        active = self.active
        if active is not None and getattr(active, "sheet_state", "visible") == "visible":
            return index

        for visible_index in visible_indices[index:]:
            self.active = visible_index
            return visible_index
        return None

    @property
    def read_only(self) -> bool:
        """True if this workbook was opened with ``read_only=True``.

        The flag reflects the explicit ``read_only=True`` argument passed
        to :func:`wolfxl.load_workbook`. Workbooks opened with the default
        ``read_only=False`` keep the normal in-memory read behavior, while
        streaming row iteration is used for explicit read-only workbooks
        and very large sheets.
        """
        return bool(getattr(self, "_read_only", False))

    @property
    def write_only(self) -> bool:
        """Whether this workbook is in openpyxl write-only mode (G20)."""
        return bool(getattr(self, "_write_only", False))

    @property
    def data_only(self) -> bool:
        """Whether formulas expose cached values where available."""
        return bool(getattr(self, "_data_only", False))

    @property
    def path(self) -> str:
        """Openpyxl-compatible workbook part path."""
        return "/xl/workbook.xml"

    @property
    def rels(self) -> list[Any]:
        """Workbook relationship list placeholder."""
        return []

    @property
    def shared_strings(self) -> list[Any]:
        """Workbook shared-string table placeholder."""
        return []

    @property
    def loaded_theme(self) -> Any:
        """Loaded theme bytes, when exposed."""
        return getattr(self, "_loaded_theme_override", None)

    @loaded_theme.setter
    def loaded_theme(self, value: Any) -> None:
        self._loaded_theme_override = value

    @property
    def vba_archive(self) -> bytes | None:
        """Raw ``xl/vbaProject.bin`` bytes for ``.xlsm`` workbooks (G19, RFC-072).

        Returns the underlying VBA archive bytes when the workbook was
        loaded via ``load_workbook(path, modify=True)`` from a macro-
        bearing file (``.xlsm`` / ``.xlam``). Returns ``None`` for
        plain ``.xlsx`` inputs (no VBA part), for write-mode workbooks
        (``Workbook()``), and for non-modify reads (``load_workbook(path)``
        without ``modify=True``) — those code paths do not retain the
        bytes in v1.0.

        Read-only inspection — there is no setter. VBA authoring is
        tracked under G28 and remains decision-gated.
        """
        override = getattr(self, "_vba_archive_override", None)
        if override is not None:
            return override
        patcher = getattr(self, "_rust_patcher", None)
        if patcher is None:
            return None
        getter = getattr(patcher, "get_vba_archive_bytes", None)
        if getter is None:
            return None
        return getter()

    @vba_archive.setter
    def vba_archive(self, value: Any) -> None:
        """Set an openpyxl-compatible VBA marker/archive placeholder."""
        self._vba_archive_override = value

    @property
    def _external_links(self) -> list[Any]:
        """Workbook-level external-link collection (RFC-071 / G18).

        Lazily loaded the first time the property is read. For
        write-mode workbooks (``Workbook()``) and bytes-backed reads
        with no source path the list is empty by definition. Modify-mode
        and path-backed read mode parse ``xl/externalLinks/`` parts on
        first access.

        The returned collection is mutable. Unchanged source links are
        preserved byte-for-byte; dirty collections are rewritten with
        workbook rels, content-types, externalLink parts, and per-link
        rels kept in sync.
        """
        cached = getattr(self, "_external_links_cache", None)
        if cached is not None:
            return cached
        from wolfxl import _external_links as _el

        if not getattr(self, "_keep_links", True):
            cached = _el.ExternalLinkCollection()
            self._external_links_cache = cached
            return cached

        source_path = getattr(self, "_source_path", None)
        source_bytes = getattr(self, "_source_bytes", None)
        if source_path:
            cached = _el.load_external_links(source_path)
        elif source_bytes:
            cached = _el.load_external_links_from_bytes(source_bytes)
        # Write mode: no source ZIP exists yet, return an empty list.
        elif self._rust_writer is not None:
            cached = _el.ExternalLinkCollection()
        else:
            cached = _el.ExternalLinkCollection()
        self._external_links_cache = cached
        return cached

    @property
    def external_links(self) -> list[Any]:
        """Alias for :attr:`_external_links` (openpyxl-shape compat)."""
        return self._external_links

    @property
    def chartsheets(self) -> list[Any]:
        """Return chart sheets in this workbook.

        Write/modify-mode chartsheets created through
        :meth:`create_chartsheet` are returned in tab order.
        """
        chartsheets = getattr(self, "_chartsheets", {})
        return [chartsheets[name] for name in self._sheet_names if name in chartsheets]

    @property
    def epoch(self) -> Any:
        """Workbook date epoch, matching openpyxl's public property."""
        from wolfxl.utils.datetime import CALENDAR_MAC_1904, CALENDAR_WINDOWS_1900

        return CALENDAR_MAC_1904 if self.workbook_properties.date1904 else CALENDAR_WINDOWS_1900

    @epoch.setter
    def epoch(self, value: Any) -> None:
        """Set the workbook date epoch."""
        from wolfxl.utils.datetime import CALENDAR_MAC_1904, CALENDAR_WINDOWS_1900

        if value == CALENDAR_MAC_1904:
            self.workbook_properties.date1904 = True
        elif value == CALENDAR_WINDOWS_1900:
            self.workbook_properties.date1904 = False
        else:
            raise ValueError("epoch must be CALENDAR_WINDOWS_1900 or CALENDAR_MAC_1904")

    @property
    def excel_base_date(self) -> Any:
        """Compatibility alias for :attr:`epoch`."""
        return self.epoch

    @property
    def iso_dates(self) -> bool:
        """Whether datetime values should be stored as ISO strings."""
        return bool(getattr(self, "_iso_dates", False))

    @iso_dates.setter
    def iso_dates(self, value: bool) -> None:
        """Set ISO-date serialization preference."""
        self._iso_dates = bool(value)

    @property
    def mime_type(self) -> str:
        """Openpyxl-compatible workbook MIME type."""
        has_vba = bool(getattr(self, "_vba_archive_override", None)) or bool(self.vba_archive)
        if getattr(self, "template", False):
            return XLTM if has_vba else XLTX
        return XLSM if has_vba else XLSX

    @property
    def is_template(self) -> bool:
        """Openpyxl-compatible independent template marker."""
        return bool(getattr(self, "_is_template", False))

    @is_template.setter
    def is_template(self, value: bool) -> None:
        """Set the workbook ``is_template`` marker."""
        self._is_template = bool(value)

    @property
    def code_name(self) -> str | None:
        """Workbook code name from ``workbook_properties.codeName``."""
        return self.workbook_properties.codeName

    @code_name.setter
    def code_name(self, value: str | None) -> None:
        """Set workbook code name."""
        self.workbook_properties.codeName = value

    @property
    def named_styles(self) -> list[Any]:
        """Return workbook named-style names."""
        return self._named_style_names()

    @property
    def _named_styles(self) -> Any:
        """Openpyxl-compatible named-style registry object."""
        return self._named_style_registry()

    @_named_styles.setter
    def _named_styles(self, value: Any) -> None:
        """Set named styles during openpyxl-compatible stylesheet loading."""
        self._named_styles_registry = value

    @property
    def style_names(self) -> list[str]:
        """Return names for workbook named styles."""
        return self._named_style_names()

    def _named_style_names(self) -> list[str]:
        """Return openpyxl-shaped named-style names."""
        read_names = self._read_style_names()
        if read_names is not None:
            return read_names
        return ["Normal"] + [style.name for style in self._named_style_registry().user_styles()]

    def _read_style_names(self) -> list[str] | None:
        """Read named-style names from ``xl/styles.xml`` when available."""
        if getattr(self, "_rust_reader", None) is None:
            return None
        if getattr(self, "_style_names_cache", None) is not None:
            return list(self._style_names_cache)
        from wolfxl._streaming import _open_source_zip, stream_source

        source = stream_source(self)
        if not source:
            return None
        try:
            with _open_source_zip(source) as zf:
                validate_zipfile(zf)
                styles_xml = read_entry(zf, "xl/styles.xml")
        except (KeyError, OSError, zipfile.BadZipFile, ValueError):
            self._style_names_cache = ["Normal"]
            return list(self._style_names_cache)
        root = ET.fromstring(styles_xml)
        namespace = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        style_nodes = root.findall("main:cellStyles/main:cellStyle", namespace)
        names = [node.attrib["name"] for node in style_nodes if node.attrib.get("name")]
        self._style_names_cache = names or ["Normal"]
        return list(self._style_names_cache)

    def _named_style_registry(self) -> Any:
        """Return the lazily seeded named-style registry."""
        if getattr(self, "_named_styles_registry", None) is None:
            from wolfxl.styles._named_style import _NamedStyleList

            self._named_styles_registry = _NamedStyleList()
        return self._named_styles_registry

    @property
    def persons(self) -> Any:
        """Workbook-scoped threaded-comment author registry (G08).

        Lazily seeded on first access. ``wb.persons.add(name=..., user_id=...)``
        returns a ``Person`` whose GUID is auto-allocated and which is
        idempotent on ``(user_id, provider_id)`` when both are non-empty.
        When the workbook is read from a file with an existing
        ``personList.xml``, the registry is hydrated from the reader so
        the original GUIDs round-trip byte-stable.
        """
        if getattr(self, "_persons_registry", None) is None:
            from wolfxl.comments._person import Person, PersonRegistry

            registry = PersonRegistry()
            reader = getattr(self, "_rust_reader", None)
            if reader is not None and hasattr(reader, "read_persons"):
                try:
                    entries = reader.read_persons()
                except Exception:
                    entries = []
                for entry in entries:
                    person_id = entry.get("id")
                    if not person_id:
                        continue
                    registry._seed(  # noqa: SLF001
                        Person(
                            name=entry.get("display_name") or "",
                            user_id=entry.get("user_id") or "",
                            provider_id=entry.get("provider_id") or "None",
                            id=person_id,
                        )
                    )
            self._persons_registry: Any = registry
        return self._persons_registry

    def __getitem__(self, name: str) -> Worksheet:
        """Return a worksheet or chartsheet by title."""
        if name in self._sheets:
            return self._sheets[name]
        chartsheets = getattr(self, "_chartsheets", {})
        if name in chartsheets:
            return chartsheets[name]
        raise KeyError(f"Worksheet '{name}' does not exist")

    def __contains__(self, name: str) -> bool:
        """Return whether the workbook contains a sheet named ``name``."""
        return name in self._sheets or name in getattr(self, "_chartsheets", {})

    def __iter__(self):  # type: ignore[no-untyped-def]
        """Iterate worksheets in tab order."""
        return iter(self.worksheets)

    def get_sheet_by_name(self, name: str) -> Worksheet:
        """Look up a sheet by name. Deprecated in openpyxl but still widely used."""
        return self[name]

    def get_sheet_names(self) -> list[str]:
        """Deprecated openpyxl alias for :attr:`sheetnames`."""
        return self.sheetnames

    def index(self, worksheet: Worksheet) -> int:
        """Return the 0-based index of ``worksheet`` in sheet order."""
        return self._sheet_names.index(worksheet.title)

    def get_index(self, worksheet: Worksheet) -> int:
        """Deprecated openpyxl alias for :meth:`index`."""
        return self.index(worksheet)

    def remove(self, worksheet: Worksheet) -> None:
        """Remove a worksheet from the workbook (write mode only).

        In read and modify modes, WolfXL raises instead of pretending a
        destructive edit succeeded against the source workbook.

        Args:
            worksheet: Worksheet to remove.

        Raises:
            RuntimeError: If the workbook mode cannot remove sheets.
        """
        _workbook_sheets.remove_sheet(self, worksheet)

    def remove_sheet(self, worksheet: Worksheet) -> None:
        """openpyxl alias for :meth:`remove` (deprecated there, kept for parity)."""
        self.remove(worksheet)

    def __delitem__(self, name: str) -> None:
        """Delete a worksheet or chartsheet by title."""
        if name in self._sheets:
            self.remove(self._sheets[name])
            return
        chartsheets = getattr(self, "_chartsheets", {})
        if name in chartsheets:
            self._sheet_names.remove(name)
            chartsheets.pop(name)
            self._chartsheets_dirty = True
            return
        raise KeyError(name)

    def _duplicate_name(self, name: str) -> bool:
        """Return whether *name* collides with tables or defined names."""
        target = name.lower()
        for defined_name in self.defined_names:
            if str(defined_name).lower() == target:
                return True
        for ws in self.worksheets:
            tables = getattr(ws, "tables", None)
            if tables is None:
                continue
            try:
                values = tables.values()
            except AttributeError:
                values = []
            for table in values:
                table_name = getattr(table, "name", None) or getattr(table, "displayName", None)
                if table_name and str(table_name).lower() == target:
                    return True
        return False

    # ------------------------------------------------------------------
    # Workbook-level metadata queues.
    # ------------------------------------------------------------------

    @property
    def properties(self) -> Any:
        """Return the workbook's :class:`DocumentProperties` (lazy-loaded).

        In read/modify mode, parses ``docProps/core.xml`` once via the
        Rust reader and caches the result. In write mode, starts as an
        empty (all-fields-None) ``DocumentProperties`` whose attribute
        assignments flip ``self._properties_dirty`` so :meth:`save` knows
        to flush them.
        """
        return _workbook_metadata.get_properties(self)

    @properties.setter
    def properties(self, value: Any) -> None:
        """Replace the entire properties object wholesale.

        Used by callers that prefer to construct a fresh
        ``DocumentProperties`` rather than mutate fields one at a time.
        Sets the dirty flag unconditionally — replacing the object is by
        definition a write intent.
        """
        _workbook_metadata.set_properties(self, value)

    @property
    def custom_doc_props(self) -> Any:
        """Return workbook custom document properties."""
        return _workbook_metadata.get_custom_doc_props(self)

    @custom_doc_props.setter
    def custom_doc_props(self, value: Any) -> None:
        """Replace workbook custom document properties."""
        _workbook_metadata.set_custom_doc_props(self, value)

    # ------------------------------------------------------------------
    # Named ranges
    # ------------------------------------------------------------------

    @property
    def defined_names(self) -> Any:
        """Return the workbook's :class:`DefinedNameDict`.

        Lazy-loaded on first access. The container is a ``dict``
        subclass whose values are :class:`DefinedName` objects.
        Workbook-scoped names override sheet-scoped on collision.
        Mutations (``wb.defined_names["X"] = DefinedName(...)``) queue
        through to the Rust writer in write mode.
        """
        return _workbook_metadata.get_defined_names(self)

    def create_named_range(
        self,
        name: str,
        worksheet: Worksheet | None = None,
        value: str | None = None,
        scope: Any = None,  # noqa: ARG002 - openpyxl deprecated signature
    ) -> None:
        """Create a deprecated openpyxl-style defined name."""
        from wolfxl.utils import quote_sheetname
        from wolfxl.workbook.defined_name import DefinedName

        if worksheet is not None:
            value = f"{quote_sheetname(worksheet.title)}!{value}"
        self.defined_names[name] = DefinedName(name=name, value=value)

    def add_named_style(self, style: Any) -> None:
        """Register a named style on the workbook."""
        from wolfxl.styles import NamedStyle

        if not isinstance(style, NamedStyle) and hasattr(style, "name"):
            style = NamedStyle(name=str(style.name))
        self._named_style_registry().append(style)
        if hasattr(style, "bind"):
            style.bind(self)

    def _has_named_style(self, name: str) -> bool:
        """Return True if ``name`` is a registered named style.

        ``"Normal"`` always resolves to True (Excel's reserved default,
        always present in the cellStyles table). Other names must have
        been registered via :meth:`add_named_style` or read in from the
        source workbook on load.
        """
        if name in self._named_style_registry():
            return True
        if name in self._named_style_names():
            return True
        return any(
            getattr(style, "name", None) == name
            for style in self._named_style_registry().user_styles()
        )

    def create_chartsheet(self, title: str | None = None, index: int | None = None) -> Any:
        """Create a chartsheet tab."""
        from wolfxl import _chartsheets

        return _chartsheets.create_chartsheet(self, title=title, index=index)

    @property
    def workbook_properties(self) -> Any:
        """Return workbook-wide `<workbookPr>` properties."""
        return _workbook_metadata.get_workbook_properties(self)

    @workbook_properties.setter
    def workbook_properties(self, value: Any) -> None:
        """Replace workbook-wide properties."""
        _workbook_metadata.set_workbook_properties(self, value)

    @property
    def calculation(self) -> Any:
        """Return workbook calculation properties (`<calcPr>`)."""
        return _workbook_metadata.get_calc_properties(self)

    @calculation.setter
    def calculation(self, value: Any) -> None:
        """Replace workbook calculation properties."""
        _workbook_metadata.set_calc_properties(self, value)

    @property
    def calc_properties(self) -> Any:
        """Compatibility alias for :attr:`calculation`."""
        return self.calculation

    @calc_properties.setter
    def calc_properties(self, value: Any) -> None:
        """Compatibility alias for :attr:`calculation`."""
        self.calculation = value

    @property
    def views(self) -> list[Any]:
        """Return workbook window views."""
        return _workbook_metadata.get_views(self)

    @views.setter
    def views(self, value: Any) -> None:
        """Replace workbook window views."""
        _workbook_metadata.set_views(self, value)

    # ------------------------------------------------------------------
    # Workbook-level security
    # ------------------------------------------------------------------

    @property
    def security(self) -> Any:
        """Return the workbook's :class:`WorkbookProtection` block, if any.

        ``None`` when no protection is configured (the default). Assign a
        :class:`wolfxl.workbook.protection.WorkbookProtection` instance to
        enable structure / window / revision locks. Mutating an already-
        attached instance also queues the update — call
        ``wb.security = wb.security`` after mutating to force the flush
        if needed (the property write is what flips the dirty flag).
        """
        return _workbook_metadata.get_security(self)

    @security.setter
    def security(self, value: Any) -> None:
        """Set workbook protection metadata.

        Args:
            value: A ``WorkbookProtection`` instance, or ``None`` to clear the
                in-memory protection block before the next save.
        """
        _workbook_metadata.set_security(self, value)

    @property
    def fileSharing(self) -> Any:  # noqa: N802 — openpyxl-shape camelCase
        """Return the workbook's :class:`FileSharing` block, if any.

        ``None`` when no file-sharing block is configured (the default).
        Attribute name matches openpyxl's ``wb.fileSharing`` exactly so
        existing code continues to work.
        """
        return _workbook_metadata.get_file_sharing(self)

    @fileSharing.setter
    def fileSharing(self, value: Any) -> None:  # noqa: N802
        """Set workbook file-sharing metadata.

        Args:
            value: A ``FileSharing`` instance, or ``None`` to clear the
                in-memory file-sharing block before the next save.
        """
        _workbook_metadata.set_file_sharing(self, value)

    # ------------------------------------------------------------------
    # Write-mode operations
    # ------------------------------------------------------------------

    def create_sheet(self, title: str | None = None, index: int | None = None) -> Worksheet:
        """Create and append a worksheet.

        Args:
            title: Worksheet title. When omitted, empty, or already present,
                WolfXL mirrors openpyxl by generating the next unique title.

        Returns:
            The newly created :class:`Worksheet`.

        Raises:
            RuntimeError: If the workbook is not in write mode.
        """
        if getattr(self, "_read_only", False):
            raise ReadOnlyWorkbookException("Cannot create new sheet in a read-only workbook")
        return _workbook_sheets.create_sheet(self, title, index)

    def _add_sheet(self, worksheet: Worksheet, index: int | None = None) -> None:
        """Openpyxl-compatible private sheet insertion helper."""
        if not isinstance(worksheet, Worksheet):
            raise TypeError("Cannot be added to a workbook")
        if worksheet._workbook is not self:  # noqa: SLF001
            raise ValueError("You cannot add worksheets from another workbook.")
        insert_at = len(self._sheet_names) if index is None else index
        self._sheet_names.insert(insert_at, worksheet.title)

    def copy_worksheet(
        self,
        source: Worksheet | None = None,
        *,
        from_worksheet: Worksheet | None = None,
        name: str | None = None,
    ) -> Worksheet:
        """Duplicate *source* into a new sheet within this workbook.

        Supported for both new workbooks and workbooks opened with
        ``load_workbook(..., modify=True)``. Read-only workbooks raise
        ``RuntimeError``.

        The new sheet appends at the end of the tab list. The default
        title is ``f"{source.title} Copy"``; on collision an incrementing
        suffix (``Copy 2``, ``Copy 3``, ...) is appended until unique. An
        explicit ``name`` keyword argument overrides the default and
        must not collide with any existing sheet name.

        Args:
            source: Worksheet to duplicate.
            name: Optional explicit title for the copy.

        Returns:
            The newly created worksheet.

        Raises:
            RuntimeError: If the workbook is read-only.
            ValueError: If ``name`` collides with an existing sheet title.
        """
        source = source if source is not None else from_worksheet
        if source is None:
            raise TypeError("copy_worksheet() missing required argument: 'from_worksheet'")
        if getattr(self, "_read_only", False) or getattr(self, "_write_only", False):
            raise ValueError("Cannot copy worksheets in read-only or write-only mode")
        return _workbook_sheets.copy_worksheet(self, source, name=name)

    def _copy_worksheet_write_mode(self, source: Worksheet, new_title: str) -> Worksheet:
        """Clone an in-memory worksheet into a fresh sheet (write mode).

        Materialises any pending append/bulk-write buffers on the source
        so every cell lives in ``source._cells`` first, then walks that
        map and replays each cell's value + format/border onto the
        destination's lazily-allocated ``Cell`` objects. Sheet-scope
        attributes (row heights, column widths, merged ranges, freeze
        pane) are copied verbatim. The destination sheet is registered
        with the native writer via ``create_sheet`` so that downstream
        save/flush passes see it like any other sheet.
        """
        return _workbook_sheets.copy_worksheet_write_mode(self, source, new_title)

    def move_sheet(self, sheet: Worksheet | str, offset: int = 0) -> None:
        """Move *sheet* by *offset* positions within the workbook tab list.

        Mirrors openpyxl's ``Workbook.move_sheet``. The new position is
        ``current_index + offset`` and is clamped to the workbook's sheet
        bounds. The in-memory tab list is updated immediately, so
        subsequent reads of ``wb.sheetnames`` and ``wb.worksheets`` see
        the post-move order before save.

        Args:
            sheet: A ``Worksheet`` instance or sheet name string.
            offset: Integer count of positions to shift.

        Raises:
            TypeError: ``sheet`` is neither a ``Worksheet`` nor a
                ``str``; or ``offset`` is not an integer (``bool``
                is rejected explicitly).
            KeyError: the resolved sheet name is not in this workbook.
        """
        _workbook_sheets.move_sheet(self, sheet, offset)

    def save(
        self,
        filename: str | os.PathLike[str],
        *,
        password: str | bytes | None = None,
    ) -> None:
        """Flush all pending writes and save to disk.

        Args:
            filename: Destination path. In modify mode this may be the original
                source path; WolfXL writes that case through its safe in-place
                save path.
            password: Optional encryption password for the final ``.xlsx``
                payload. Install ``wolfxl[encrypted]`` to enable encryption.

        Raises:
            ValueError: If ``password`` is empty.
            RuntimeError: If the workbook mode cannot save the requested
                pending changes.
        """
        _workbook_save.save_workbook(self, filename, password=password)

    def _save_encrypted(
        self,
        filename: str,
        password: str | bytes,
    ) -> None:
        """Save plaintext to a tempfile then re-route through encryption.

        Write-side encryption stays Python-side; the Rust writer/patcher
        is unchanged. We materialise the unencrypted xlsx via the normal
        save path, slurp it back as bytes, hand it to
        :func:`wolfxl._encryption.encrypt_xlsx_to_path` for the
        in-place encryption + atomic rename onto ``filename``.

        The plaintext tempfile is always cleaned up — including on
        error paths — so we never leak unencrypted user data on disk.
        """
        _workbook_save.save_encrypted(self, filename, password)

    def _flush_pending_hyperlinks_to_patcher(self) -> None:
        """Drain each sheet's pending hyperlinks into the patcher.

        Modify-mode counterpart to the writer-side compatibility flush.
        Each ``Hyperlink`` is converted to the patcher's flat-dict shape
        and routed to ``queue_hyperlink``; the ``None`` sentinel routes
        to ``queue_hyperlink_delete``.

        Cleared after queueing so a subsequent ``save()`` on the same
        workbook doesn't double-emit.
        """
        _workbook_patcher_flush.flush_pending_hyperlinks_to_patcher(self)

    def _flush_pending_tables_to_patcher(self) -> None:
        """Drain each sheet's pending tables into the patcher.

        Modify-mode counterpart to the writer-side compatibility flush.
        Each ``Table`` is converted to the patcher's flat-dict shape and
        routed to ``queue_table``. The patcher allocates a
        workbook-unique table ``id`` at save time (any explicit ``id`` on
        the Python ``Table`` object is ignored), serializes
        ``xl/tables/tableN.xml``, splices a ``<tableParts>`` block into
        the sheet XML, mutates the sheet rels, and adds a
        ``[Content_Types].xml`` Override.

        Per-sheet drain happens in workbook tab order; within a sheet,
        append order wins (which matches openpyxl's first-add → first-
        slot semantics).

        Cleared after queueing so a subsequent ``save()`` on the same
        workbook doesn't double-emit.
        """
        _workbook_patcher_flush.flush_pending_tables_to_patcher(self)

    def _flush_pending_images_to_patcher(self) -> None:
        """Drain pending images into the patcher.

        Modify-mode counterpart to the writer-side flush in
        ``Worksheet._flush_compat_properties``. Each queued
        :class:`wolfxl.drawing.image.Image` is converted to the flat
        dict shape and routed to ``XlsxPatcher.queue_image_add``.

        Sheets that already have a drawing rel are merged into the existing
        drawing part so images and charts can share one drawing.
        """
        _workbook_patcher_flush.flush_pending_images_to_patcher(self)

    def _flush_pending_charts_to_patcher(self) -> None:
        """Drain pending chart adds in modify mode.

        Two queues are drained here:

        1. Workbook-level ``_pending_chart_adds``: a dict keyed by sheet
           title with values of ``(chart_xml: bytes, anchor_a1,
           width_emu, height_emu)`` tuples populated via
           :meth:`add_chart_modify_mode`. These are routed straight to
           ``XlsxPatcher.queue_chart_add`` and preserve the bytes-level
           escape hatch for callers that want to pass pre-serialised
           chart XML directly.
        2. Per-sheet ``Worksheet._pending_charts``: a list of
           high-level :class:`~wolfxl.chart._chart.ChartBase` instances
           queued via :meth:`Worksheet.add_chart`. In **write** mode
           these are drained inside
           ``_worksheet._flush_compat_properties`` via
           ``writer.add_chart_native``. In **modify** mode we bridge each
           chart through the ``serialize_chart_dict`` PyO3 export,
           producing chart XML bytes that are routed through the same
           ``patcher.queue_chart_add`` path as the bytes escape hatch.

        Sequenced after images / axis shifts but before the final
        ``patcher.save()`` so chart cell-range formulas can compose
        with cell rewrites in the same save.
        """
        _workbook_patcher_flush.flush_pending_charts_to_patcher(self)

    def add_pivot_cache(self, cache: Any) -> Any:
        """Register a pivot cache with this workbook.

        The cache should already be materialized, for example with
        ``wolfxl.pivot.PivotCache.from_worksheet``. A registered cache can be
        referenced by one or more pivot tables.

        Args:
            cache: Pivot cache object to register.

        Returns:
            The same cache object, with its workbook-scoped cache id set.

        Raises:
            RuntimeError: If the workbook is not open in modify mode.
            ValueError: If the cache has already been registered.
        """
        return _workbook_features.add_pivot_cache(self, cache)

    def _flush_pending_slicers_to_patcher(self) -> None:
        """Drain queued slicers into the patcher's slicer queue.

        For each worksheet, iterate over ``ws._pending_slicers`` and
        bridge the (cache, slicer) pair to the Rust patcher via
        ``queue_slicer_add(sheet_title, cache_dict, slicer_dict)``.
        Sequenced after pivots and sheet setup, and before autofilters, so the
        relationship graph is stable before filter rows are rewritten.
        """
        _workbook_patcher_flush.flush_pending_slicers_to_patcher(self)

    def add_slicer_cache(self, cache: Any) -> Any:
        """Register a slicer cache with this workbook.

        Slicer caches are workbook-scoped, and one cache can be referenced by
        multiple slicer presentations. The source pivot cache must already be
        registered with :meth:`add_pivot_cache`.

        Args:
            cache: Slicer cache object to register.

        Returns:
            The same cache object, with its workbook-scoped slicer cache id
            set.

        Raises:
            RuntimeError: If the workbook is not open in modify mode.
            ValueError: If the cache has already been registered or
                the source pivot cache is not registered.
        """
        return _workbook_features.add_slicer_cache(self, cache)

    def _flush_pending_sheet_setup_to_patcher(self) -> None:
        """Drain each sheet's queued sheet-setup mutations into the patcher.

        Sheets whose Worksheet has any of ``_page_setup``,
        ``_page_margins``, ``_print_options``, ``_header_footer``,
        ``_sheet_view``, or ``_protection`` non-default get their
        ``to_rust_setup_dict()`` queued. The
        Rust patcher then re-emits the sheet-scope XML blocks and splices
        them into the sheet via wolfxl_merger::merge_blocks.

        ``print_titles`` (workbook-scope ``_xlnm.Print_Titles``
        definedName) does not route through this sheet-setup patch; it
        composes through the existing defined-names queue.
        The dict still includes a ``print_titles`` slot for the writer-mode
        path.
        """
        _workbook_patcher_flush.flush_pending_sheet_setup_to_patcher(self)

    def _flush_pending_page_breaks_to_patcher(self) -> None:
        """Drain each sheet's queued page-break and sheet-format mutations.

        Sheets whose Worksheet has any of ``_row_breaks``,
        ``_col_breaks``, or ``_sheet_format`` non-default get their
        merged patch dict queued. The Rust patcher then re-emits the
        sheet-scope XML blocks
        (``<rowBreaks>`` / ``<colBreaks>`` / ``<sheetFormatPr>``) and
        splices them into the sheet via wolfxl_merger::merge_blocks.

        Sheets whose all three slots are at construction defaults
        (e.g. zero breaks AND default sheet-format) are skipped to
        keep the no-op save path byte-identical.
        """
        _workbook_patcher_flush.flush_pending_page_breaks_to_patcher(self)

    def _flush_pending_autofilters_to_patcher(self) -> None:
        """Drain each sheet's ``ws.auto_filter`` into the patcher.

        Only sheets where the user actually configured filter columns
        OR a sort state OR (legacy) just a ref are queued. The Rust
        patcher then re-emits the ``<autoFilter>`` block and computes the
        ``<row hidden="1">`` markers.
        """
        _workbook_patcher_flush.flush_pending_autofilters_to_patcher(self)

    def _flush_pending_pivots_to_patcher(self) -> None:
        """Drain pending pivot caches and tables in modify mode.

        Two queues are drained here:

        1. Workbook-level ``_pending_pivot_caches`` — a list of
           :class:`~wolfxl.pivot.PivotCache` instances queued via
           :meth:`add_pivot_cache`. Each cache is bridged through
           ``serialize_pivot_cache_dict`` (definition XML) +
           ``serialize_pivot_records_dict`` (records XML), then
           routed to ``patcher.queue_pivot_cache_add``.
        2. Per-sheet ``Worksheet._pending_pivot_tables`` — a list of
           :class:`~wolfxl.pivot.PivotTable` instances queued via
           :meth:`Worksheet.add_pivot_table`. Each table is bridged
           through ``serialize_pivot_table_dict`` and routed to
           ``patcher.queue_pivot_table_add(sheet, xml, cache_id)``.

        Sequenced after charts and before the final ``patcher.save()`` so
        pivot relationships are added against an already-stable rels graph.
        """
        _workbook_patcher_flush.flush_pending_pivots_to_patcher(self)

    def _flush_pending_pivot_source_edits_to_patcher(self) -> None:
        """G17 / RFC-070 — drain pivot source-range mutations.

        For each worksheet, scan the lazily-materialised
        ``_pivot_handles_cache`` for dirty
        :class:`~wolfxl.pivot.PivotTableHandle` instances and queue a
        source-range edit for each via ``patcher.register_pivot_source_edit``.
        Sequenced after :meth:`_flush_pending_pivots_to_patcher` so adds
        and edits compose in the same save.
        """
        _workbook_patcher_flush.flush_pending_pivot_source_edits_to_patcher(self)

    def add_chart_modify_mode(
        self,
        sheet_title: str,
        chart_xml: bytes,
        anchor_a1: str,
        width_emu: int = 4_572_000,
        height_emu: int = 2_743_200,
    ) -> None:
        """Queue a pre-serialized chart XML part for a worksheet.

        This lower-level API is for callers that already have chart XML bytes
        and want to attach them while modifying an existing workbook. Most
        callers should prefer :meth:`Worksheet.add_chart` with a chart object.

        Args:
            sheet_title: Title of the target worksheet.
            chart_xml: Serialized ``<chartSpace>`` XML bytes.
            anchor_a1: A1-style anchor cell, such as ``"D2"``.
            width_emu: Chart width in EMUs.
            height_emu: Chart height in EMUs.

        Raises:
            ValueError: If ``sheet_title`` or ``anchor_a1`` is invalid.
            RuntimeError: If the workbook is not open in modify mode.
        """
        _workbook_features.add_chart_modify_mode(
            self,
            sheet_title,
            chart_xml,
            anchor_a1,
            width_emu,
            height_emu,
        )

    def _flush_pending_comments_to_patcher(self) -> None:
        """Drain each sheet's pending comments into the patcher.

        Modify-mode counterpart to the writer-side compatibility flush.
        Each ``Comment`` is converted to the patcher's flat-dict shape and
        routed to ``queue_comment``; the ``None`` sentinel routes to
        ``queue_comment_delete``.
        """
        _workbook_patcher_flush.flush_pending_comments_to_patcher(self)

    def _flush_pending_threaded_comments_to_patcher(self) -> None:
        """Drain each sheet's pending threaded comments into the patcher.

        Modify-mode counterpart to the writer-side flush. Top-level threads
        are bundled with their replies into one ``queue_threaded_comment``
        call keyed by cell coordinate. The ``None`` sentinel routes to
        ``queue_threaded_comment_delete``.
        """
        _workbook_patcher_flush.flush_pending_threaded_comments_to_patcher(self)

    def _flush_pending_persons_to_patcher(self) -> None:
        """Drain ``wb.persons`` into the patcher's personList queue.

        Idempotent on Person ``id``: re-running the flush is safe and the
        patcher's ``queue_person`` skips duplicates against both the
        existing personList and the queued additions.
        """
        _workbook_patcher_flush.flush_pending_persons_to_patcher(self)

    def _flush_pending_data_validations_to_patcher(self) -> None:
        """Drain ``_pending_data_validations`` on every sheet into the patcher.

        Modify-mode counterpart to the writer-side compatibility flush.
        Each DV is converted to the patcher's flat-dict payload via
        ``_dv_to_patcher_dict``. Per-sheet drain happens in ``ws.title``
        order; within a sheet, append order wins so the final
        ``<dataValidations>`` block reflects the order the user appended
        them.

        Cleared after queueing so a subsequent ``save()`` on the same
        workbook doesn't double-emit.
        """
        _workbook_patcher_flush.flush_pending_data_validations_to_patcher(self)

    def _flush_pending_conditional_formats_to_patcher(self) -> None:
        """Drain ``_pending_conditional_formats`` on every sheet into the patcher.

        Modify-mode counterpart to the writer-side compatibility flush.
        Rules sharing a sqref are coalesced into a single
        ``ConditionalFormattingPatch`` (one wrapper per range) so
        priority ordering within a wrapper reflects insertion order.
        Multiple ``add()`` calls with different sqrefs produce multiple
        patches; the patcher then emits them in encounter order while
        threading the workbook-wide ``running_dxf_count`` through the Rust
        writer.

        Cleared after queueing so a subsequent ``save()`` on the same
        workbook doesn't double-emit.
        """
        _workbook_patcher_flush.flush_pending_conditional_formats_to_patcher(self)

    def _flush_pending_axis_shifts_to_patcher(self) -> None:
        """Drain queued row and column shifts into the patcher.

        Each tuple ``(sheet_title, axis, idx, n)`` is forwarded to
        ``_rust_patcher.queue_axis_shift(sheet, axis, idx, n)``. The patcher
        drains the queue in append order during ``save()``.

        Empty queue is the no-op identity path — patcher is not
        called, no FFI hop, no file mutation.

        Cleared after queueing so a subsequent ``save()`` on the same
        workbook doesn't double-emit.
        """
        _workbook_patcher_flush.flush_pending_axis_shifts_to_patcher(self)

    def _flush_pending_range_moves_to_patcher(self) -> None:
        """Drain queued range moves into the patcher.

        Each tuple ``(sheet_title, src_min_col, src_min_row,
        src_max_col, src_max_row, d_row, d_col, translate)`` is
        forwarded to ``_rust_patcher.queue_range_move(...)``. The patcher
        drains the queue in append order during ``save()``.

        Empty queue is the no-op identity path — patcher is not
        called, no FFI hop, no file mutation.

        Cleared after queueing so a subsequent ``save()`` on the same
        workbook doesn't double-emit.
        """
        _workbook_patcher_flush.flush_pending_range_moves_to_patcher(self)

    def _flush_pending_sheet_copies_to_patcher(self) -> None:
        """Drain queued sheet copies into the patcher.

        Each ``(src_title, dst_title)`` pair forwards to
        ``_rust_patcher.queue_sheet_copy(src, dst)``. The patcher drains the
        queue in append order before per-sheet mutations, so cloned sheets
        are visible to downstream drains.

        Empty queue is the no-op identity path — patcher is not
        called, no FFI hop, no file mutation. Cleared after queueing
        so a subsequent ``save()`` on the same workbook doesn't
        double-emit.
        """
        _workbook_patcher_flush.flush_pending_sheet_copies_to_patcher(self)

    def _flush_defined_names_to_patcher(self) -> None:
        """Drain pending defined-name updates into the patcher.

        Modify-mode counterpart to ``_flush_workbook_writes``'s
        defined-name branch. Each ``DefinedName`` is converted to the
        patcher's flat-dict shape and routed to
        ``_rust_patcher.queue_defined_name``. ``None``-valued optional
        fields are filtered out before crossing the PyO3 boundary so
        the Rust extractors see a clean "missing key" rather than a
        Python ``None`` (matches the convention in
        ``_flush_properties_to_patcher``).

        Cleared after queueing so a subsequent ``save()`` on the same
        workbook doesn't double-emit. Empty queue is a no-op (the Rust
        side's no-op guard is the second line of defence; workbook.xml
        is left untouched if no upserts arrive).
        """
        _workbook_patcher_flush.flush_defined_names_to_patcher(self)

    def _flush_security_to_patcher(self) -> None:
        """Drain workbook security and file-sharing state into the patcher.

        Builds the flat dict and forwards it to
        ``_rust_patcher.queue_workbook_security``. The Rust side merges the
        payload into ``xl/workbook.xml``.

        Empty (no setter ever ran) is a no-op; the patcher leaves
        workbook.xml byte-identical with the source.
        """
        _workbook_patcher_flush.flush_security_to_patcher(self)

    def _build_security_dict(self) -> dict[str, Any]:
        """Return the flat dict for the workbook's security blocks.

        Either branch may be ``None`` (the user only set one of the two
        slots). Always returns a dict — never ``None`` — so callers can
        unconditionally forward to the Rust side.
        """
        return _workbook_patcher_flush.build_security_dict(self)

    def _flush_pending_sheet_moves_to_patcher(self, name: str, offset: int) -> None:
        """Queue a single sheet reorder on the patcher.

        Called eagerly from ``move_sheet`` rather than batched at
        ``save()`` time: each ``move_sheet`` call queues exactly one
        entry, and the patcher composes them in queue order against
        its own running tab list (which is initialised from the
        source ZIP's ``xl/workbook.xml`` and updated in place on save).

        The empty-queue invariant lives on the Rust side: an unused
        ``move_sheet`` call (i.e. modify-mode workbook never touched)
        means ``queued_sheet_moves`` is empty, which in turn keeps
        ``xl/workbook.xml`` byte-identical with the source.
        """
        _workbook_patcher_flush.queue_sheet_move_to_patcher(self, name, offset)

    def _flush_properties_to_patcher(self) -> None:
        """Drain dirty document properties into the patcher.

        Modify-mode counterpart to ``_flush_workbook_writes``'s
        property branch. Builds a flat dict keyed with the patcher's
        snake_case schema (``last_modified_by``, ``content_status``,
        ``created_iso``, ``modified_iso``) and filters ``None`` before
        crossing the PyO3 boundary so ``extract_str`` sees a clean
        "missing key" rather than a Python ``None``.

        Resets ``_properties_dirty`` so a subsequent ``save()`` on the
        same workbook doesn't double-emit. ``modified_iso`` is left
        unset on this side — the Rust patcher stamps it via
        ``current_timestamp_iso8601`` (or ``WOLFXL_TEST_EPOCH=0`` →
        ``1970-01-01T00:00:00Z`` for byte-identical save tests). If the
        user explicitly set ``props.modified``, that value wins.
        """
        _workbook_patcher_flush.flush_properties_to_patcher(self)

    def _flush_workbook_writes(self) -> None:
        """Push workbook-level metadata + defined names into the Rust writer."""
        _workbook_writer_flush.flush_workbook_writes(self)

    # ------------------------------------------------------------------
    # Formula evaluation (requires wolfxl.calc)
    # ------------------------------------------------------------------

    def calculate(self) -> dict[str, Any]:
        """Evaluate all formulas in the workbook.

        Returns a dict of cell_ref -> computed value for all formula cells.
        Builtin functions need no extra dependencies; the optional
        ``wolfxl[calc]`` extra enables additional formula evaluation.

        The internal evaluator is cached so that a subsequent
        :meth:`recalculate` call can reuse it without rescanning.

        The next :meth:`save` writes the computed numbers, strings, booleans
        and errors as typed XLSX formula caches, readable with
        ``load_workbook(data_only=True)``; the formula text is kept. Formulas
        the evaluator cannot compute are not given a new cache. Editing an
        input, formula, sheet name, defined name or the date system after
        ``calculate()`` invalidates these caches and the save drops cached
        formula values rather than keep them out of date; call ``calculate()``
        again before saving to write fresh results. :meth:`recalculate` is a
        what-if and never changes what is saved. Saved XLOOKUP formulas use
        Excel's ``_xlfn.XLOOKUP`` spelling.
        """
        return _workbook_calc.calculate_workbook(self)

    def cached_formula_values(self) -> dict[str, Any]:
        """Return Excel-saved cached formula results for every sheet.

        Keys are workbook-qualified cell references like ``"Sheet1!B2"``.
        This is a fast read-only path for ingestion workloads that need
        Excel's last-calculated formula values without evaluating formulas in
        Python. Cells whose formulas have no cached value are omitted.
        """
        return _workbook_calc.cached_formula_values(self)

    def recalculate(
        self,
        perturbations: dict[str, float | int],
        tolerance: float = 1e-10,
    ) -> RecalcResult:
        """Perturb input cells and recompute affected formulas.

        Returns a ``RecalcResult`` describing which cells changed.
        Requires the ``wolfxl.calc`` module.

        If :meth:`calculate` was called first, the cached evaluator is
        reused (avoiding a full rescan + recalculate).
        """
        return _workbook_calc.recalculate_workbook(self, perturbations, tolerance)

    # ------------------------------------------------------------------
    # Context manager + cleanup
    # ------------------------------------------------------------------

    def close(self) -> None:
        """Release native handles and delete any temporary decrypted input."""
        _workbook_lifecycle.close_workbook(self)

    def __enter__(self) -> Workbook:
        """Return this workbook for ``with`` statement use."""
        return _workbook_lifecycle.enter_workbook(self)

    def __exit__(self, *args: object) -> None:
        """Close this workbook at the end of a ``with`` block."""
        _workbook_lifecycle.exit_workbook(self, *args)

    def __repr__(self) -> str:
        """Return a compact debug representation for this workbook.

        Returns:
            A string containing the workbook mode and sheet names.
        """
        return _workbook_lifecycle.repr_workbook(self)


Workbook.copy_worksheet.__signature__ = _inspect.Signature(
    [
        _inspect.Parameter("self", _inspect.Parameter.POSITIONAL_OR_KEYWORD),
        _inspect.Parameter("from_worksheet", _inspect.Parameter.POSITIONAL_OR_KEYWORD),
    ]
)
Workbook.__signature__ = _inspect.Signature(
    [
        _inspect.Parameter(
            "write_only",
            _inspect.Parameter.POSITIONAL_OR_KEYWORD,
            default=False,
        ),
        _inspect.Parameter(
            "iso_dates",
            _inspect.Parameter.POSITIONAL_OR_KEYWORD,
            default=False,
        ),
    ]
)
Workbook.save.__signature__ = _inspect.Signature(
    [
        _inspect.Parameter("self", _inspect.Parameter.POSITIONAL_OR_KEYWORD),
        _inspect.Parameter("filename", _inspect.Parameter.POSITIONAL_OR_KEYWORD),
    ]
)
