# Read values and style IDs without constructing Cells

WolfXL already exposes `Worksheet.iter_cell_records()`. Use it for ingestion or
style grouping when you need Python values and workbook-local style IDs, rather
than a complete `Cell` object for every coordinate. This is a different API from
`iter_rows()` and its timings must be reported separately from Cell API results.

```python
import wolfxl

workbook = wolfxl.load_workbook("input.xlsx", data_only=True)
try:
    worksheet = workbook.active
    columns = max(1, worksheet.max_column)
    rows_per_batch = max(1, 50_000 // columns)
    for first_row in range(1, worksheet.max_row + 1, rows_per_batch):
        for record in worksheet.iter_cell_records(
            min_row=first_row,
            max_row=min(worksheet.max_row, first_row + rows_per_batch - 1),
            min_col=1,
            max_col=columns,
            include_format=True,
            include_extended_format=False,
            include_coordinate=False,
            include_style_id=True,
            include_empty=False,
        ):
            value = record["value"]
            style_id = record.get("style_id", 0)
            # Consume (record["row"], record["column"], value, style_id).
finally:
    workbook.close()
```

`include_format=True` is required for style IDs. `include_extended_format=False`
skips font, fill, alignment, and border summary fields while retaining style IDs,
number formats, and named-style metadata. `include_coordinate=False` avoids A1
string allocation; numeric `row` and `column` remain 1-based.

The example bounds each Python record batch to at most 50,000 coordinates on a
valid XLSX sheet. Each native call returns a Python list for its selected range,
then `iter_cell_records()` yields its dictionaries. The native reader still
materializes the worksheet once: this is **bounded Python output, not XML
streaming**. Range calls can scan the native cell collection again; avoid tiny
batches. For sparse sheets with a few populated cells and huge dimensions, one
unbounded sparse scan can be cheaper than stepping through many empty windows.

| Field or condition | Existing behavior |
|---|---|
| Ordinary numbers, strings, booleans, errors | Native Python scalar values; no Cell allocation. |
| Excel serial date formats | Native date conversion uses the workbook's 1900 or 1904 epoch and returns `datetime`. These values are not raw Excel serials. |
| Elapsed formats such as `[h]:mm:ss` | The record API currently returns a native `datetime`, while `Cell.value` converts it to `timedelta`. Use `number_format` and explicit normalization if elapsed-duration parity is required. |
| Subsecond date/time precision | The current native record conversion returns whole seconds and does not expose the original date-formatted numeric serial. Validate another extraction path if retaining subsecond precision matters. |
| Formulas | `data_only=False` returns the `=` formula string. `data_only=True` returns its saved cached value, or `None` when no cached value exists. Formula records are included by default even without a cache. |
| Merged subordinate positions | Format metadata, including style IDs, is suppressed. Blank positions are omitted by default; `include_empty=True` includes blank records within the selected rectangle. This does not compose merged borders into records. |
| Missing `style_id` | Treat as default style 0 for grouping. IDs are workbook-local and are not comparable across files. |
| Pending value edits | Existing record overlays reflect pending values. |
| Pending style edits | Style IDs and style summary metadata remain source-derived until saving and reopening. Use this recipe on a clean read workbook when grouping by style. |
| Complete Font/Fill/Border/Protection objects | Use the Cell API. Record summaries and raw IDs do not replace its full style contract. |

No new API is introduced by this recipe. The tests in
`tests/test_style_id_record_recipe.py` pin the existing value, formula, merge,
style-ID, and pending-edit behavior and verify that clean reads leave the
worksheet's Cell map empty.
