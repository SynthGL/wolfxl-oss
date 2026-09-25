---
name: wolfxl-xlsx
description: Edit an existing Excel workbook (.xlsx or .xlsm) without losing the parts the edit did not touch, such as pivot tables, slicers, timelines, charts, shapes, sparklines, data validation, Power Pivot data models, and VBA macros. Use when changing cells, formulas, or styles in a workbook someone already built, when recalculating its formulas, or when checking that a saved file kept everything. Not needed to create a new workbook from scratch.
license: MIT
---

# Editing existing Excel workbooks with WolfXL

The usual Python workflow (load with openpyxl, save, then recalculate with
LibreOffice) rebuilds the whole file from what those tools understand. Anything
they do not model disappears on save, with no warning and no repair prompt
when Excel later opens the file. In measured runs this removed slicers,
timelines, sparklines, extension data validation, and a Power Pivot data model
together with the pivots built on it.

WolfXL modify mode patches only the cells you change and copies every other
package part through unchanged. It uses the openpyxl API, so the editing code
looks the same.

## Setup

```bash
python -m pip install wolfxl
```

This installs WolfXL Community (MIT). WolfXL Commercial adds native formula
calculation that writes fresh cached values into the saved file; it installs
from `https://packages.wolfxl.com/simple/` with a license credential. The
helper detects which one is installed.

Run the helper, `scripts/wolfxl_xlsx.py` in this skill's directory, with the
same Python interpreter that has `wolfxl` installed. If the environment
already provides one, use it instead of installing. Every subcommand prints
JSON.

## Workflow

1. **Inspect** the workbook before choosing tools:

   ```bash
   python scripts/wolfxl_xlsx.py inspect report.xlsx
   ```

   `openpyxl_save_would_drop` and `libreoffice_recalc_would_drop` list what
   each tool would remove from this particular file.

2. **Edit a copy** in modify mode. Never overwrite the original until step 4
   passes.

   ```python
   import shutil
   import wolfxl

   shutil.copyfile("report.xlsx", "report-edited.xlsx")
   wb = wolfxl.load_workbook("report-edited.xlsx", modify=True)  # add keep_vba=True for .xlsm
   ws = wb["Inputs"]
   ws["B4"] = 42
   ws["C10"] = "=SUM(C2:C9)"
   wb.save("report-edited.xlsx")
   wb.close()
   ```

   Write formulas as formulas, not values computed in Python, so the sheet
   still updates when its inputs change.

   If `save` raises, the copy may be partly written. Copy the original again
   before retrying.

3. **Recalculate** with WolfXL, never with LibreOffice:

   ```bash
   python scripts/wolfxl_xlsx.py recalc report-edited.xlsx
   ```

   | Exit | `status` | Meaning |
   |---|---|---|
   | 0 | `success` | Formulas evaluated with no Excel error values |
   | 1 | `errors_found` | `formula_errors` lists cells showing `#REF!`, `#DIV/0!`, and so on; fix them and rerun |
   | 2 | `not_calculated` | The evaluator could not load this workbook; `reason` says why |

   The saved file is also marked so Excel recalculates it on open. Even on
   exit 2, `recalc` rewrites `xl/workbook.xml` to set that flag.

   With Commercial, fresh values are written into the file. With Community,
   values are computed in memory and returned under `values`, but the file
   stores none of them: existing formulas keep their old cached values and new
   formulas have no cached value at all. Anything that reads stored values
   (pandas, `data_only=True`, file previews, spreadsheet imports) shows old
   numbers or blanks until someone opens and saves the file in Excel. Tell the
   user this when you hand the file back.

   `not_evaluated` lists formulas the installed evaluator does not support;
   report them rather than guessing their values. On exit 2, leave the
   workbook's names and formulas alone rather than rewriting them to satisfy
   the evaluator. Report the `reason` and that values are not stored.

4. **Verify** nothing was lost:

   ```bash
   python scripts/wolfxl_xlsx.py verify report.xlsx report-edited.xlsx
   ```

   Exit 0 with `"status": "preserved"` means every part and feature in the
   source is still present. Exit 1 lists `parts_missing`, `families_lost`, and
   `sheet_features_lost`. If the user asked you to remove one of those, pass
   `--allow NAME` for it; otherwise discard the output and redo the edit.
   `verify` needs only the standard library, so it also checks files written by
   other tools.

5. **Report** what changed, the recalc status, and the verify result. Replace
   the original only if the user wants that.

## Rules

- Do not save an existing workbook with openpyxl when `inspect` lists anything
  under `openpyxl_save_would_drop`. openpyxl is fine for reading and for new
  workbooks.
- Do not run a LibreOffice headless recalculation (including another skill's
  `recalc.py`) on a file whose `libreoffice_recalc_would_drop` is not empty. It
  drops those parts even when the edit before it preserved everything. Use
  `recalc` above, or tell the user which parts LibreOffice would remove and
  let them decide.
- Keep `.xlsm` files as `.xlsm` and load them with `keep_vba=True`.
- Excel opening a file without a repair prompt does not prove nothing was
  lost; the damaged files in the measurement all opened cleanly. Run `verify`.

## Limits to report

- **Pivot tables and pivot charts are not refreshed.** If the edit changed
  their source data, they still show the old numbers. Tell the user to use
  Refresh in Excel, and note any filters that could change what a refresh
  shows.
- **Excel tables do not grow on their own.** Writing a column or row next to
  a table leaves it outside the table's range. Extend the table in the same
  modify-mode session (Commercial, or Community 2.0.4 and later):

  ```python
  from wolfxl.worksheet.table import TableColumn

  table = ws.tables["Table1"]
  next_id = max(column.id for column in table.tableColumns) + 1
  table.tableColumns.append(TableColumn(id=next_id, name="PROFIT"))
  table.ref = "A1:H38"  # the table's new full range
  ```

  Community 2.0.3 and earlier ignore table range changes in modify mode
  (check `wolfxl.__version__`). There, tell the user the new cells sit
  outside the table. Do not hand-edit the table XML.

## Ambiguous requests

When a value to change matches more than one row, list every match. Ask which
one the user means, or change all of them and list the old values so the user
can undo any that were wrong.
