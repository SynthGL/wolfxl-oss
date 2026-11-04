# wolfxl-edit

Pure Rust edits of existing XLSX and XLSM workbooks. Unchanged ZIP parts retain
identical bytes, including macros, embedded objects, and custom XML. Saves write
a sibling temporary file, sync it, and atomically replace the destination.

```rust,no_run
use std::path::Path;
use wolfxl_edit::{CellEdit, EditSession, EditValue};

let mut book = EditSession::open(Path::new("research.xlsx"))?;
book.write_cells("Data", &[
    CellEdit { cell: "B2".into(), value: EditValue::Number(125.0) },
    CellEdit { cell: "C2".into(), value: EditValue::Formula("=B2*1.1".into()) },
])?;
book.set_number_format("Data", "B2:C2", "#,##0.00")?;
book.save(Path::new("research-edited.xlsx"))?;
# Ok::<(), wolfxl_edit::EditError>(())
```

`EditSession` supports batched values, formulas, and clears; row and column
insertion and deletion; adding and renaming worksheets; and number formats.
References use A1 notation. Row and column positions are one-based; an optional
new-sheet position is zero-based. Clears retain existing cell styles. Number
formats retain other style components and reuse equivalent formats.

Formula writes and structural changes request full recalculation on open and
remove a stale calculation chain. New formulas have no cached result. This
crate edits package XML; it does not calculate formulas or render workbooks.

The `sheet_patcher` and `rich_text` modules provide the same shared XML helpers
used by the workbook writer and Python modify mode.

Licensed under MIT.
