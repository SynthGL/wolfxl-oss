# wolfxl-core

[![crates.io](https://img.shields.io/crates/v/wolfxl-core.svg)](https://crates.io/crates/wolfxl-core)
[![docs.rs](https://docs.rs/wolfxl-core/badge.svg)](https://docs.rs/wolfxl-core)

Pure-Rust spreadsheet reader with Excel number-format-aware cell rendering. Backs the
[`wolfxl-cli`](https://crates.io/crates/wolfxl-cli) previewer.

```toml
[dependencies]
wolfxl-core = "0.8"
```

```rust
use wolfxl_core::Workbook;

let mut wb = Workbook::open("examples/sample-financials.xlsx")?;
let sheet = wb.first_sheet()?;
let (rows, cols) = sheet.dimensions();
println!("{} rows x {} columns", rows, cols);
# Ok::<_, wolfxl_core::Error>(())
```

## Scope

- **In scope today**: read xlsx / xls / xlsb / ods through the native readers,
  and csv through the delimited reader. Resolve best-effort number-format
  metadata with the `xl/styles.xml` cellXfs fallback, classify formats into
  `FormatCategory`, render via `format_cell`, map workbook structure, and infer
  per-column schema/cardinality summaries. Preserve sparse origins, date
  epochs, names, XLSX tables, and canonical Excel error values.
- **Not yet**: write side.

The PyO3 layer in the sibling [`wolfxl`](https://pypi.org/project/wolfxl/)
package consumes the same native readers through edition-specific adapters.

## Exports

- `Workbook`, `Sheet`, `Cell`, `CellValue`
- `format_cell`, `FormatCategory`
- `classify_sheet`, `WorkbookMap`, `SheetMap`, `SheetClass`
- `infer_sheet_schema`, `SheetSchema`, `ColumnSchema`, `InferredType`, `Cardinality`

## License

MIT
