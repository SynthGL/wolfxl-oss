# SPDX-License-Identifier: MIT
# Copyright (c) 2026 Wolfie

"""Reproduce fixed reader-only fixtures with independent spreadsheet producers.

Run: uv run --with xlwt==1.3.0 --with odfpy==1.4.1 python generate_fixtures.py
xlwt is BSD-licensed; odfpy is Apache-2.0/GPL-2.0-or-later/LGPL-2.1-or-later.
Neither producer is a runtime or test dependency of the native Rust readers.
"""

from datetime import datetime
from pathlib import Path

import xlwt
from odf.opendocument import OpenDocumentSpreadsheet
from odf.table import Table, TableCell, TableRow
from odf.text import P

ROOT = Path(__file__).resolve().parent


def main() -> None:
    xls = xlwt.Workbook()
    sheet = xls.add_sheet("Values")
    sheet.write(2, 1, "Independent café 中文")
    sheet.write(2, 2, 42.5)
    sheet.write(2, 3, True)
    sheet.write(2, 4, datetime(2024, 2, 29), xlwt.easyxf(num_format_str="yyyy-mm-dd"))
    sheet.write(2, 5, xlwt.Formula("SUM(C3,2.5)"))
    xls.add_sheet("Second sheet").write(0, 0, "last")
    xls.save(str(ROOT / "fixtures.xls"))

    ods = OpenDocumentSpreadsheet()
    table = Table(name="Values")
    table.addElement(TableRow(numberrowsrepeated=2))
    row = TableRow()
    row.addElement(TableCell())
    label = TableCell(valuetype="string")
    label.addElement(P(text="Independent café 中文"))
    row.addElement(label)
    row.addElement(TableCell(valuetype="float", value=42.5))
    row.addElement(TableCell(valuetype="boolean", booleanvalue="true"))
    row.addElement(TableCell(valuetype="date", datevalue="2024-02-29"))
    row.addElement(TableCell(valuetype="float", value=45, formula="of:=SUM([.C3];2.5)"))
    table.addElement(row)
    ods.spreadsheet.addElement(table)
    second = Table(name="Second sheet")
    last_row = TableRow()
    last = TableCell(valuetype="string")
    last.addElement(P(text="last"))
    last_row.addElement(last)
    second.addElement(last_row)
    ods.spreadsheet.addElement(second)
    ods.save(str(ROOT / "fixtures.ods"))


if __name__ == "__main__":
    main()
