# wolfxl-cli test fixtures

`data-table-cities.xlsx` models a typical analyst workbook. It is built by
`generate_data_table_cities.py` and the build is byte-for-byte reproducible:

```bash
uv run --with openpyxl==3.1.5 python generate_data_table_cities.py data-table-cities.xlsx
```

| Sheet | Content |
|---|---|
| `Table` | Title cells in row 1 (`C1`, `I1`), header rows in row 2, and two side-by-side Excel tables: `BigCity` at `C2:F15` and `SmallCity` at `I2:L15`. Each has a totals row (row 15) of `SUBTOTAL` formulas. Both tables have a `City` column, so the header row repeats `City` at `C2` and `I2`. Columns `G` and `H` are empty. |
| `Formula` | Structured references such as `=BigCity[[#Totals],[Latitude]]` and `COUNTA(BigCity[City])`. |
| `PivotTable` | Values laid out like a pivot table's output (no pivot cache). |

Every formula cell stores the cached value Excel writes, because `wolfxl`
reads stored values. The totals-row averages are the values Excel stores,
which differ in the last bits from a left-to-right float sum.

The other fixtures in this directory were authored for this crate.
