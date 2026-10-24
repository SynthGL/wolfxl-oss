"""Shim for ``openpyxl.worksheet`` - subpackage re-exports.

openpyxl organizes worksheet-adjacent classes (DataValidation, Table,
Hyperlink, PageSetup, HeaderFooter, SheetView, SheetProtection) under
this package. wolfxl exposes the same module paths so imports succeed.

Sprint Ο Pod 1A (RFC-055) added: page_setup, header_footer, views,
protection, properties, print_settings.
"""

from __future__ import annotations

from importlib import import_module as _import_module

from wolfxl._compat import _openpyxl_name_fallback

# ``from wolfxl.worksheet import X`` users keep every re-export, but a name loads
# its defining module on first access. Importing any ``wolfxl.worksheet.<module>``
# runs this package first, and an eager namespace made every write_only workbook
# load the page-setup, header/footer, view, protection, property and print-settings
# models it never touches.
_EXPORTS = {
    "PageMargins": "wolfxl.worksheet.page_setup",
    "PageSetup": "wolfxl.worksheet.page_setup",
    "PrintOptions": "wolfxl.worksheet.page_setup",
    "PrintPageSetup": "wolfxl.worksheet.page_setup",
    "HeaderFooter": "wolfxl.worksheet.header_footer",
    "HeaderFooterItem": "wolfxl.worksheet.header_footer",
    "Pane": "wolfxl.worksheet.views",
    "Selection": "wolfxl.worksheet.views",
    "SheetView": "wolfxl.worksheet.views",
    "SheetViewList": "wolfxl.worksheet.views",
    "SheetProtection": "wolfxl.worksheet.protection",
    "Outline": "wolfxl.worksheet.properties",
    "PageSetupProperties": "wolfxl.worksheet.properties",
    "WorksheetProperties": "wolfxl.worksheet.properties",
    "ColRange": "wolfxl.worksheet.print_settings",
    "PrintArea": "wolfxl.worksheet.print_settings",
    "PrintTitles": "wolfxl.worksheet.print_settings",
    "RowRange": "wolfxl.worksheet.print_settings",
}

__all__ = [
    "PageMargins",
    "PageSetup",
    "PrintOptions",
    "PrintPageSetup",
    "HeaderFooter",
    "HeaderFooterItem",
    "Pane",
    "Selection",
    "SheetView",
    "SheetViewList",
    "SheetProtection",
    "Outline",
    "PageSetupProperties",
    "WorksheetProperties",
    "ColRange",
    "PrintArea",
    "PrintTitles",
    "RowRange",
]

_fallback = _openpyxl_name_fallback(globals())


def __getattr__(name: str):
    module = _EXPORTS.get(name)
    if module is None:
        return _fallback(name)
    value = getattr(_import_module(module), name)
    globals()[name] = value
    return value
