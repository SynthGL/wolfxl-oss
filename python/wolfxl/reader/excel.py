"""Excel reader entry point compatible with ``openpyxl.reader.excel``."""

from __future__ import annotations

from zipfile import BadZipFile, ZIP_DEFLATED, ZipFile

from wolfxl._compat import _openpyxl_name_fallback
from wolfxl import load_workbook
from wolfxl.packaging.manifest import Manifest, Override
from wolfxl.reader.strings import read_rich_text, read_string_table
from wolfxl.xml.functions import fromstring
from wolfxl.xml.constants import (
    ARC_CONTENT_TYPES,
    ARC_THEME,
    ARC_WORKBOOK,
    SHARED_STRINGS,
    XLSM,
    XLSX,
    XLTM,
    XLTX,
)

KEEP_VBA = False
SUPPORTED_FORMATS = (".xlsx", ".xlsm", ".xltx", ".xltm")


def _find_workbook_part(package):
    workbook_types = [XLTM, XLTX, XLSM, XLSX]
    for content_type in workbook_types:
        part = package.find(content_type)
        if part:
            return part

    defaults = {part.ContentType for part in package.Default}
    workbook_type = defaults & set(workbook_types)
    if workbook_type:
        return Override("/" + ARC_WORKBOOK, workbook_type.pop())

    raise OSError("File contains no valid workbook part")


class ExcelReader:
    """Openpyxl-shaped staged reader over the native workbook and sheet readers."""

    def __init__(
        self,
        fn,
        read_only: bool = False,
        keep_vba: bool = KEEP_VBA,
        data_only: bool = False,
        keep_links: bool = True,
        rich_text: bool = False,
    ) -> None:
        self.fn = fn
        self.archive = ZipFile(fn)
        self.valid_files = self.archive.namelist()
        self.read_only = read_only
        self.keep_vba = keep_vba
        self.data_only = data_only
        self.keep_links = keep_links
        self.rich_text = rich_text
        self.shared_strings = []
        self.package = None
        self.wb = None

    def read_manifest(self) -> None:
        src = self.archive.read(ARC_CONTENT_TYPES)
        self.package = Manifest.from_tree(fromstring(src))

    def read_strings(self) -> None:
        if self.package is None:
            self.read_manifest()
        content_type = self.package.find(SHARED_STRINGS)
        if content_type is None:
            return
        reader = read_rich_text if self.rich_text else read_string_table
        with self.archive.open(content_type.PartName[1:]) as src:
            self.shared_strings = reader(src)

    def read_workbook(self) -> None:
        # ZipFile's staged reads share a file-like input's cursor.
        if hasattr(self.fn, "seek"):
            self.fn.seek(0)
        self.wb = load_workbook(
            self.fn,
            read_only=self.read_only,
            keep_vba=self.keep_vba,
            data_only=self.data_only,
            keep_links=self.keep_links,
            rich_text=self.rich_text,
        )
        if not self.read_only:
            # Staged eager consumers require full Cell metadata from iter_rows.
            self.wb._auto_stream = False

    def read_worksheets(self) -> None:
        """Expose stored cells to consumers that inspect ``Worksheet._cells``.

        ``read_workbook`` creates native-backed sheet proxies. Ordinary WolfXL
        reads keep their cells lazy; ExcelReader's eager stage materializes only
        populated coordinates, never the dimension rectangle. Values, styles,
        merges and links still resolve through the native-backed cell objects.
        Read-only workbooks retain their streaming behavior.
        """
        if self.wb is None:
            self.read_workbook()
        if self.read_only:
            return
        for sheet in self.wb.worksheets:
            for record in sheet.iter_cell_records(
                include_format=False,
                include_coordinate=False,
                include_style_id=False,
                include_extended_format=False,
            ):
                sheet.cell(row=record["row"], column=record["column"])

    def read_theme(self) -> None:
        if self.wb is None:
            self.read_workbook()
        if ARC_THEME in self.valid_files:
            self.wb.loaded_theme = self.archive.read(ARC_THEME)

    def read(self) -> None:
        try:
            self.read_manifest()
            self.read_strings()
            self.read_workbook()
            self.read_theme()
            self.read_worksheets()
        finally:
            if not self.read_only:
                self.archive.close()

    def read_chartsheet(self, sheet, rel) -> None:  # noqa: ANN001
        if self.wb is None:
            self.read_workbook()
        if sheet.name not in self.wb.sheetnames:
            self.wb.create_sheet(sheet.name)


__all__ = [
    "BadZipFile",
    "ExcelReader",
    "KEEP_VBA",
    "SUPPORTED_FORMATS",
    "ZIP_DEFLATED",
    "ZipFile",
    "_find_workbook_part",
    "fromstring",
    "load_workbook",
]

__getattr__ = _openpyxl_name_fallback(globals())
