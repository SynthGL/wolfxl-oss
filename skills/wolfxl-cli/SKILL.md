---
name: wolfxl-cli
description: Read, filter, pivot, profile, reconcile, or fingerprint spreadsheets and CSV/TSV exports from the terminal with the wolfxl CLI. Use when you need to understand a workbook or delimited file without opening Excel or writing Python, such as getting a token-budgeted briefing, filtering rows, building a pivot, profiling columns, comparing two files by key with a tolerance, checking a file against a fingerprint, or evaluating a simple formula over piped data. Connected wolfxl MCP tools can also edit existing XLSX/XLSM files; otherwise use the wolfxl-xlsx skill for workbook changes.
license: MIT
---

# Reading and checking spreadsheets with the wolfxl CLI

`wolfxl` reads `.xlsx`, `.xlsm`, `.xls`, `.xlsb`, `.ods`, `.csv`, `.tsv`, and
`.txt` files, and the commands below never write to them. Every command can
print JSON for you to parse and plain text to show the user.

## Setup

```bash
cargo install wolfxl-cli   # installs the `wolfxl` binary
wolfxl --version
```

If `wolfxl` is already on `PATH`, use it instead of installing.

## MCP alternative

When the wolfxl MCP server (`wolfxl mcp`) is connected, its tools return the
same JSON without a shell: `get_workbook` (sheets, headers, and a `revision`),
`read_range` (values and formula text, at most 10,000 cells per call),
`calc`, `stats`, `pivot`, `diff`, and `fingerprint`. It also offers edits to
existing `.xlsx`/`.xlsm` files (`write_cells`, `set_number_format`, row and
column inserts and deletes, `add_sheet`, `rename_sheet`, `undo`, `redo`).
Edit only when the user asks for a change: pass the `revision` from your last
read as `expected_revision`, preview with `dry_run: true` when unsure, and
tell the user that written formulas recalculate when Excel opens the file.
With `--read-only` the server offers only the reading tools.

## Pick the command

| Need | Command |
|---|---|
| Sheets, sizes, headers, named ranges | `wolfxl map FILE --format text` |
| A briefing that fits your context budget | `wolfxl agent FILE --max-tokens 800` |
| Column types, cardinality, formats | `wolfxl schema FILE` |
| Rows, filtered and narrowed | `wolfxl peek FILE --where 'COL>=100' --select A,B -e json` |
| Per-column profile | `wolfxl stats FILE -e json` |
| Grouped totals | `wolfxl pivot FILE --rows R --column C --values sum:F -e json` |
| One formula over a sheet or piped data | `wolfxl calc '=SUM(B:B)' FILE --json` |
| Which functions `calc` supports | `wolfxl list-functions --check NAME,NAME --json` |
| Two tables compared by key | `wolfxl diff LEFT RIGHT --key K --tolerance 0.01` |
| Content fingerprint, and a check against one | `wolfxl fingerprint FILE`, `wolfxl verify FILE --fingerprint wxf1:...` |

## Workflow

1. **Orient.** Run `map`, then `agent` with a budget that fits what you have
   left. Use `schema` when column types matter. Pick the sheet with `-s` or
   `--sheet`; commands default to the first sheet.
2. **Narrow.** Use `peek --where` and `--select` instead of dumping a whole
   sheet. Row 1 holds the headers. Repeat `--where` to require every
   condition. Operators are `=`, `!=`, `<`, `>`, `<=`, `>=`, and `~`
   (contains, ignoring case). Numeric comparisons accept money spellings such
   as `$1,200.00`, `1,200`, and `(500.00)`.
3. **Summarize.** Use `stats` for distributions and `pivot` for grouped
   totals (`sum`, `count`, `avg`, `min`, `max`).
4. **Compute.** Use `calc` for a single answer and `peek --add NAME=FORMULA`
   for a computed column. Write the `--add` formula for the first data row
   (row 2); each later row gets it filled down the way Excel fills formulas.
   The built-in engine covers arithmetic, comparison, `&`, references, and
   67 functions: math and rounding, `SUMIF(S)`, `COUNTIF(S)`,
   `AVERAGEIF(S)`, `MAXIFS`, `MINIFS`, `IF`, `IFERROR`, `AND`, `OR`,
   `VLOOKUP`, `HLOOKUP`, `XLOOKUP`, `INDEX`, `MATCH`, `OFFSET`, `CHOOSE`,
   text functions including `TEXT`, date functions including `EOMONTH`, and
   `PMT`, `PV`, `FV`, `NPV`, `IRR`, `SLN`, and `DB`. Run
   `list-functions --check` before relying on any function.
5. **Check.** Use `diff` to reconcile two files and `verify` to confirm a
   file still matches a recorded fingerprint. Report the exit code along with
   the output.

## Rules

- Use JSON output (`-e json`, `--json`, or `--format json`) when you will
  parse the result, and text output when you show it to the user.
- Standard input needs `-` as the file and `--from csv|tsv|xlsx|...`.
- Never run `peek -i` (`--interactive`). It needs a terminal and exits 2
  without one. Suggest it to the user for browsing instead.
- `calc` and `--add` read the values stored in the file. Formula cells count
  at their last saved results, which may be stale.
- When `calc` exits 2 for an unsupported function, say so. Do not
  approximate the result with a different formula without telling the user.
- `diff` keys keep their spelling: `007` and `7` differ. Use
  `--key-transform trim|lower|digits` when the sources format keys
  differently, and say which transform you used.

## Exit codes

| Command | 0 | 1 | 2 |
|---|---|---|---|
| `calc` | Value printed | Result is an Excel error such as `#DIV/0!` | Usage, input, parse, or unsupported function |
| `diff` | Reconciled | Unmatched key or difference outside tolerance (any difference with `--strict-exit`) | Usage or input failure, missing column, empty or duplicate key |
| `verify` | Matches | Does not match; both fingerprints printed | Unreadable input or malformed fingerprint |
| `list-functions --check` | All supported | At least one name unsupported | Usage failure |
| `peek`, `stats`, `pivot` | Printed | File could not be read (`peek`) | Usage or input failure, unknown column |
