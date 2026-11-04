# wolfxl-cli

[![crates.io](https://img.shields.io/crates/v/wolfxl-cli.svg)](https://crates.io/crates/wolfxl-cli)

Command-line previewer and analyst toolkit for spreadsheets. Installs the
`wolfxl` binary and opens `.xlsx`, `.xlsm`, `.xls`, `.xlsb`, `.ods`, `.csv`,
`.tsv`, and `.txt` inputs through the same `peek` / `map` / `agent` /
`schema` workflow, plus `calc`, `pivot`, `stats`, `diff`, `fingerprint`, and
`verify` for analysis and reconciliation. A built-in formula evaluator powers
`calc` and `peek --add` formula columns, and `list-functions` lists the
functions it implements.

```bash
cargo install wolfxl-cli
```

## Subcommands

- `wolfxl peek <file|->` - styled preview or text/csv/json/markdown export, with `--add` formula columns, `--where` row filters and `--select` columns; `--interactive` browses it in a terminal grid
- `wolfxl map <file>` - one-page workbook overview (sheets, dims, headers)
- `wolfxl agent <file> --max-tokens N` - token-budgeted workbook briefing
- `wolfxl schema <file>` - per-column type, cardinality, and format inference
- `wolfxl calc <formula> [file|-]` - evaluate one formula against a sheet's stored values or piped input
- `wolfxl list-functions [--check NAME,...]` - list the formula functions `calc` implements, or check names against them
- `wolfxl pivot <file|-> --rows R --values AGG:FIELD` - group rows by header fields and aggregate with grand totals
- `wolfxl stats <file|->` - per-column profile: counts by type, distinct values, numeric summary, and most frequent values
- `wolfxl diff <left> <right> --key K` - reconcile two tables row by row on a key column, with numeric tolerance
- `wolfxl fingerprint <file>` - style-independent content fingerprint (`wxf1:<sha256 hex>`)
- `wolfxl verify <file> --fingerprint <F>` - check a file against a content fingerprint

## Usage

```bash
wolfxl peek workbook.xlsx                   # styled box preview (default)
wolfxl peek workbook.xlsx -e text           # tab-separated for awk/cut
wolfxl peek workbook.xlsx -e csv            # RFC 4180 CSV
wolfxl peek workbook.xlsx -e json           # machine-readable JSON
wolfxl peek workbook.xlsx -e json -n 20 --max-columns 8 # bounded JSON
wolfxl peek workbook.xlsx -e markdown       # Markdown table
wolfxl peek workbook.xlsx -n 20 -w 30       # 20 rows, 30-char column cap
wolfxl peek workbook.xlsx -s "Balance Sheet"
wolfxl peek workbook.xlsb                   # binary Excel workbook
wolfxl peek ledger.csv --where 'Amount>=1,000' --where 'Status~open' -e csv
wolfxl peek workbook.xlsx --select Account,Q1  # chosen columns, in this order
wolfxl peek ledger.csv --add 'Margin=C2-B2' --where 'Margin<0' -e csv
wolfxl peek sales.xlsx --add 'Running=SUM($D$2:D2)'  # running total column
cat ledger.csv | wolfxl peek - --from csv --where 'Region=west'
wolfxl peek workbook.xlsx --interactive     # full-screen grid: search, sort, value counts
wolfxl map export.csv --format text         # delimited file as one sheet
wolfxl map workbook.xlsx                    # workbook inventory for agents
wolfxl map workbook.xlsx --format text      # terminal-friendly summary
wolfxl agent workbook.xlsx --max-tokens 400 # fit a briefing to a token budget
wolfxl schema workbook.xlsx                 # JSON schema-style column summary
wolfxl schema workbook.xlsx -s "P&L" -f text
wolfxl calc '=SUM(B:B)' export.csv         # formula over a delimited file
cat export.csv | wolfxl calc 'SUM(B:B)' --from csv
wolfxl calc "=SUM('P&L'!B3:B5)" workbook.xlsx --sheet "Balance Sheet" --json
wolfxl list-functions                      # NAME  category lines plus a count
wolfxl list-functions --category math --json
wolfxl list-functions --check SUM,IF,XLOOKUP  # exit 1: XLOOKUP is not built in
wolfxl pivot sales.csv --rows Region --values sum:Amount,count:Amount
wolfxl pivot sales.csv --rows Region --column Year --values avg:Amount -e csv
cat sales.csv | wolfxl pivot - --from csv --rows Region,Product --values max:Amount -e json
wolfxl stats ledger.csv                    # box table, one row per column
wolfxl stats workbook.xlsx --select Account,Q1 --top 3 -e json
cat ledger.csv | wolfxl stats - --from csv --where 'Region=west' -e csv
wolfxl diff ledger.xlsx bank.csv --key Invoice --tolerance 0.01
cat export.csv | wolfxl diff - books.xlsx --key ID --from csv --format text
wolfxl diff old.csv new.csv --key SKU --compare Price,Qty --strict-exit --quiet
wolfxl fingerprint workbook.xlsx           # wxf1:<64 hex digits>
wolfxl fingerprint export.csv --json       # {fingerprint, version, sheets, cells, formulas}
wolfxl verify workbook.xlsx --fingerprint wxf1:3f0c...
```

The `text`, `csv`, `json`, and `markdown` exporters are tuned for piping into
LLM / agent contexts and Unix tooling. Human-facing `box`, `text`, `csv`, and
`markdown` renders preserve common Excel display formats such as currency and
percentages, while JSON keeps machine-shaped values
(`{sheet, rows, columns, headers, data}`).

JSON output honors `--max-rows` and `--max-columns` before serialization. The
`rows` and `columns` fields continue to report the full sheet dimensions while
`headers` and `data` contain only the requested prefix.

The default `box` exporter is wolfxl-branded with `╔═╗` banner and `┌─┬─┐`
table borders.

## Standard input and delimited fields

`peek`, `pivot`, `stats`, `diff`, and `calc` read standard input when the
file is `-`; `calc` also reads it when no file is given. Standard input needs
`--from csv|tsv|xlsx|xlsm|xls|xlsb|ods`; for a path, `--from` overrides the
extension. A delimited input is one sheet named after the file stem, or
`stdin` when piped.

Delimited fields arrive as text. `pivot`, `stats`, `calc`, and the formulas
of `peek --add` type them the way a spreadsheet text import does. After
surrounding whitespace is trimmed, `TRUE` and `FALSE` in any letter case
become booleans, and a plain decimal number becomes a number: an optional
sign, digits with an optional decimal point (`42`, `-1.5`, `.5`), and an
optional exponent (`-1.5e3`). Every other field stays text, including
thousands separators (`1,234`), currency symbols (`$1,200`), hex (`0x10`),
`inf`, and `NaN`. `peek` output and filters, `diff`, and `fingerprint` work
from the source text, so `007` and `1,200` print and match as written.

## Formula evaluation

```text
wolfxl calc FORMULA [FILE|-] [--sheet S] [--from FMT] [--json]
```

`wolfxl calc` loads the stored values of every sheet and evaluates FORMULA
on the selected sheet, the first one unless `--sheet` names another. The
leading `=` is optional. Without `FILE`, or with `-`, it reads standard
input, which needs `--from` as described above.

- **Values.** References read stored values: a formula cell contributes its
  cached result, and `calc` never recalculates it. Delimited fields are typed
  as described in
  [Standard input and delimited fields](#standard-input-and-delimited-fields),
  and dates and times read as 1900-system serial numbers. References use the
  sheet's own addresses, so a table that starts at `C3` is read at `C3`.
- **Built-in engine.** The evaluator covers arithmetic (`+`, `-`, `*`, `/`,
  `^`, unary minus, and postfix `%`), comparison (`=`, `<>`, `<`, `>`, `<=`,
  `>=`), and `&` concatenation over numbers, strings, `TRUE`/`FALSE`, and
  error literals such as `#N/A`. It reads A1 references with or without `$`,
  sheet-qualified references such as `'P&L'!B3`, and ranges such as `B2:D10`,
  `B:B`, and `2:2`; whole columns and rows stop at the sheet's used extent.
- **Functions.** The only functions are `SUM`, `AVERAGE`, `COUNT`, `MIN`,
  `MAX`, and `IF`, which evaluates only the branch it returns. A formula
  that names any other function exits 2, naming it and pointing to
  [`wolfxl list-functions`](#function-registry).
- **Semantics.** Arithmetic uses `f64`. Text converts to a number with
  locale-neutral decimal and exponent rules. Text comparisons ignore case by
  Unicode lowercase ordering rather than locale-specific collation, and `&`
  joins numbers in plain decimal form rather than with cell number formats.
  A range is an array: operators do not intersect or broadcast it, and
  results never spill.

Scalars print as text with up to 15 significant digits, `TRUE` or `FALSE`
for booleans, an empty line for a blank, and the literal for an Excel error.
An array result, such as the bare range `=B2:C4`, prints as CSV rows.
`--json` prints `{formula, sheet, type, value}` on one line: `formula` with
its leading `=`, the evaluated `sheet`, a `type` of `number`, `text`,
`boolean`, `blank`, `error`, or `array`, and a `value` that is a number,
string, boolean, `null` for a blank, the error literal as a string, or a 2-D
array of those.

| Exit | Meaning |
|---|---|
| 0 | The formula produced a value, an array included |
| 1 | The formula produced an Excel error value, printed as `#DIV/0!` and so on |
| 2 | Usage or input failure, a formula that cannot be parsed, or an unsupported function, described on stderr |

```bash
wolfxl calc '=SUM(B:B)/COUNT(B:B)' sales.tsv || echo "exit $?"
wolfxl calc '=IF(MAX(C2:C20)>1000,"review","ok")' ledger.csv
```

## Function registry

`wolfxl list-functions` prints the functions of the built-in evaluator
behind `wolfxl calc` and `wolfxl peek --add`, one per line, sorted by name,
with the category of each.

```text
wolfxl list-functions [--category C] [--search TEXT] [--json]
wolfxl list-functions --check NAME[,NAME...] [--json]
```

```text
$ wolfxl list-functions
AVERAGE  statistical
COUNT    statistical
IF       logical
MAX      math
MIN      math
SUM      math
6 functions
```

- **Categories.** The registry's categories are `logical`, `math`, and
  `statistical`. `--category` keeps one of them, ignoring case and
  surrounding spaces, so `Math` works. Any other category exits 2 and lists
  the valid ones.
- **Search.** `--search TEXT` keeps the functions whose name contains
  `TEXT`, ignoring case: `--search m` finds `MAX`, `MIN`, and `SUM`. With
  both flags, a function must pass both.
- **Text output.** Each line is the name padded to the longest listed name,
  two spaces, and the category. The last line is the count: `N functions`
  for the full list, or `M of N functions` when a filter is active. A filter
  that matches nothing prints only the count line and still exits 0.

`--json` prints one document. `total` is the size of the whole registry and
`count` the number of listed entries; `category` and `search` echo the
filters (`null` when absent), with `category` in its registry spelling.

```json
{"contract_version": 1, "total": 6, "count": 3, "category": "math", "search": null,
 "functions": [{"name": "MAX", "category": "math"}, {"name": "MIN", "category": "math"},
               {"name": "SUM", "category": "math"}]}
```

`--check` takes function names, comma-separated or repeated, and reports
each in the order given. Names ignore case and surrounding whitespace and
print in upper case. A name is `supported` when the registry lists it, and
`unsupported` otherwise, which includes spreadsheet functions the built-in
engine does not implement, such as `XLOOKUP`. `--check` cannot be combined
with `--category` or `--search`.

```text
$ wolfxl list-functions --check sum,xlookup,nope
SUM      supported (math)
XLOOKUP  unsupported
NOPE     unsupported
1 of 3 supported
```

With `--json`, `total` is the size of the registry, `checked` counts the names,
`supported` the supported ones, and `category` is `null` for an unsupported name:

```json
{"contract_version": 1, "total": 6, "checked": 2, "supported": 1, "results": [
  {"name": "SUM", "status": "supported", "category": "math"},
  {"name": "XLOOKUP", "status": "unsupported", "category": null}
]}
```

JSON is printed on one line; it is wrapped above for reading.

| Exit | Meaning |
|---|---|
| 0 | Listed, or every checked name is supported |
| 1 | A checked name is unsupported |
| 2 | Usage failure: an unknown category, an empty `--check` name, or `--check` with `--category` or `--search` |

## Row filters and column selection

`wolfxl peek` reads a path, or standard input when the file is `-` (see
above). The first row of the sheet is the header row.

`--where COLUMN<op>VALUE` keeps the data rows that satisfy the condition;
the header row is always kept. Repeat `--where` to require every condition
(AND). The operators are `=`, `!=`, `<`, `>`, `<=`, `>=`, and `~`
(contains). The first operator character in the expression ends the column
name, and `<=`, `>=`, and `!=` are read as one operator. Spaces around the
column name are ignored; the value is used as written.

- When the value reads as a number, the comparison is numeric. Money
  spellings count on both sides: `$1,200.00`, `1,200`, and the accounting
  negative `(500.00)`. A cell that holds no number fails every numeric
  comparison except `!=`.
- Any other value compares as text, ignoring case, against the cell's
  displayed value: strings as stored, numbers in plain notation, booleans as
  `TRUE` or `FALSE`. `<` and `>` order text lexically.
- `~` always tests text: the cell contains the value, ignoring case.

`--select A,B` prints only the named columns, in the order given. A filter
may test a column that `--select` leaves out. Column names match a header
exactly first, then by a unique case-insensitive match.

Filtering and selection run before `--max-rows` and `--max-columns` and apply
to every export format. Without `--where` or `--select`, output is unchanged.

| Exit | Meaning |
|---|---|
| 0 | The sheet was printed, possibly with no data rows |
| 1 | The file, sheet, or output could not be read or written |
| 2 | A malformed `--where`, an unknown or ambiguous column in `--where` or `--select` (stderr lists the headers), standard input without `--from`, or `--interactive` without a terminal or with standard input as the data source |

```bash
wolfxl peek ledger.csv --where 'Amount<(100)' --select Rep,Amount -e json
```

## Formula columns

`wolfxl peek --add NAME=FORMULA` appends a column named NAME whose cells
are FORMULA evaluated on every data row. Repeat `--add` for more columns.
The argument splits at its first `=`, so `Margin=C2-B2` and `Margin==C2-B2`
are the same formula.

Write the formula for the first data row, the row under the header (row 2
for a sheet that starts in A1). Each later row gets it filled down the way a
spreadsheet fills a formula: relative row references move with the row,
`$`-anchored rows stay put, and columns never move. `C2-B2` on row 5 reads
`C5-B5`, and `SUM($B$2:B2)` is a running total. References to other sheets
move like local ones, so `Rates!B2` reads `Rates!B3` one row down. Whole
columns such as `B:B`, string literals, and function names are never
rewritten. Rows after the last non-blank row get no formula.

Formulas use the evaluator of `wolfxl calc`, described in
[Formula evaluation](#formula-evaluation), against the stored values of
every sheet, with delimited fields typed as for `calc`. Added columns take
the free column letters after the sheet's last column, in `--add` order, so
a later formula can use an earlier added column: with headers in A:D,
`--add 'Revenue=B2*C2' --add 'Double=E2*2'` doubles Revenue. A formula also
sees the rows above it in its own column, while later added columns and the
rows below in its own column read as blank.

Results print the way `wolfxl calc` prints a scalar: text with up to 15
significant digits, also in JSON, and `TRUE`/`FALSE` for booleans. A blank
result leaves the cell empty, and a range result such as `B2:C2` is
`#VALUE!`, since ranges are not intersected with the row. An Excel error
such as `#DIV/0!` stays in its cell, shown as `ERROR: #DIV/0!` in the text
views and as `"#DIV/0!"` in JSON, and `--where 'Unit=#DIV/0!'` finds it.

Formula columns are computed over the whole sheet before `--where` and
`--select`, so filters can test added columns, `--select` can show them,
and a filter never changes what a formula sees.

| Exit | Meaning |
|---|---|
| 0 | The sheet was printed with its added columns, error cells included |
| 1 | The file, sheet, or output could not be read or written |
| 2 | A `--add` without `NAME=` or a formula, a formula that cannot be parsed or names an unsupported function (stderr points to `wolfxl list-functions`), or a NAME already used by a header or an earlier `--add`, ignoring case; stderr names the `--add` argument |

```bash
wolfxl peek orders.csv --add 'Unit=D2/B2' --where 'Unit=#DIV/0!' --select Order,Unit
```

## Interactive grid

`wolfxl peek FILE --interactive` (or `-i`) opens the sheet in a full-screen
terminal grid instead of printing it. The header row stays frozen at the
top, and the last line is a status line showing the file, the sheet and its
position among the workbook's sheets, the cursor's row and column, the
current sort, the search and its match count, and notes.

| Key | Action |
|---|---|
| Arrow keys | Move one cell |
| PageUp / PageDown | Move one screen of rows |
| Home / End | First / last column |
| `g` / `G` | First / last row |
| `/` | Type a search, then Enter; Esc cancels. Matches cell text, ignoring case |
| `n` / `N` | Next / previous match, wrapping around the sheet |
| `s` | Sort by the cursor's column; press again to reverse |
| `S` | Restore the source row order |
| `f` | Show the value frequencies of the cursor's column; any key closes it |
| Tab / Shift-Tab | Next / previous sheet |
| `q`, Esc, Ctrl-C | Quit and restore the terminal |

A column sorts numerically when every non-blank value in it is a number,
money spellings such as `$1,200.00` and `(500.00)` included; otherwise it
sorts as text, ignoring case. Blank cells stay last in both directions, and
rows with equal values keep their source order. Frequencies count each
displayed value of the column over the rows shown, most frequent first.

`--add`, `--where`, and `--select` shape the sheet named by `--sheet` (or
the first sheet) before the grid opens. Tab loads each other sheet the first
time it is reached and applies the same `--add` columns, filters, and
selection to it. A sheet where an added NAME matches a header opens without
added columns, and a sheet that lacks a column the filters or selection name
opens unfiltered; the status line notes either case. Each sheet keeps its
cursor, scroll, and sort while you move between sheets.

Columns are as wide as their header or widest value among the first 1,000
rows, capped at `--max-width`, and longer values end in `…`. Numbers are
right-aligned. `--max-rows` and `--max-columns` do not apply: every row and
column is reachable.

The grid needs a terminal: when standard output is not one, or when the
data comes from standard input (`-`), which the grid needs for keys,
`--interactive` exits 2 with a message and draws nothing. It cannot be
combined with `-e`. The terminal is restored on exit, including after a
panic.

## Pivot tables

```text
wolfxl pivot FILE --rows R[,R2] [--column C] --values AGG:FIELD[,AGG:FIELD]
             [--sheet S] [--from FMT] [-e box|csv|json]
```

`wolfxl pivot` groups the data rows of one sheet (the first unless `--sheet`
names another) and aggregates value fields per group. Row 1 holds the
headers. `FILE` of `-` reads standard input, which needs `--from`;
delimited fields are typed as described in
[Standard input and delimited fields](#standard-input-and-delimited-fields).
Fields are header names, matched exactly and then case-insensitively when
exactly one header matches. `--rows` and `--values` take comma-separated
lists or repeat. Fully blank rows are skipped.

| Aggregation | Result |
|---|---|
| `sum` | Sum of the numeric cells |
| `count` | Number of non-empty cells, numeric or not |
| `avg` | Mean of the numeric cells |
| `min` | Smallest numeric cell |
| `max` | Largest numeric cell |

Numeric cells are stored numbers. Text, booleans, dates, and errors are
skipped by `sum`, `avg`, `min`, and `max`, and a group without numeric cells
leaves those aggregates blank (`null` in JSON). Each `--rows` field becomes a
key column; an empty key prints as `(blank)`. Keys of a field sort
numerically when every non-blank key is a number, otherwise by display text
in code point order, with `(blank)` last.

Value columns are labelled `AGG:FIELD`, for example `sum:Amount`. With
`--column`, each distinct value of that field becomes a column, sorted the
same way, followed by a `Total` column. With several `--values`, each column
key gets one column per value (`2024 sum:Amount`, `2024 count:Amount`) and the
totals become `Total sum:Amount` and so on. A final `Total` row closes every
table. Totals recompute the aggregation over all contributing rows, so an
`avg` total is the mean of every value rather than a mean of group means.

`-e box` (the default) and `-e csv` print numbers the way `peek` does:
integers with thousands separators and other values with two decimals.
`-e json` prints one object with full-precision numbers:

```json
{
  "contract_version": 1,
  "sheet": "sales",
  "rows": ["Region"],
  "column": "Year",
  "values": ["sum:Amount"],
  "column_keys": [2023, 2024],
  "headers": ["Region", "2023", "2024", "Total"],
  "data": [
    ["East", 20, 12.5, 32.5],
    ["West", 7, 5, 12]
  ],
  "totals": [27, 17.5, 44.5]
}
```

`column` is `null` and `column_keys` is empty without `--column`. Each `data`
row holds the key values (numbers stay numbers) and then the value cells in
header order. `totals` holds the value cells of the `Total` row.

| Exit | Meaning |
|---|---|
| 0 | The pivot table was printed |
| 2 | Usage or input failure, an unknown header, an unknown aggregation, or a malformed `--values` entry, described on stderr |

## Content fingerprints

`wolfxl fingerprint` prints `wxf1:` followed by the lowercase hex SHA-256 of
a canonical text stream built from the workbook's content: sheet order and
names, every non-empty cell's address and typed value, and formula text.
Styles, number formats, and package bytes never enter the stream, so the same
content saved with a different ZIP member order or compression keeps its
fingerprint, while a changed value or formula changes it. `--json` prints
`{fingerprint, version, sheets, cells, formulas}` with the record counts.

The `wxf1` stream is UTF-8 text with one record per line, each ending in LF.
Fields are separated by TAB; inside a field, `\` is written `\\`, TAB `\t`,
CR `\r`, and LF `\n`.

```text
wxf1
S<TAB>sheet name                  one per sheet, in workbook order
V<TAB>A1<TAB>type<TAB>value       one per non-empty cell, row-major
F<TAB>A1<TAB>formula              one per formula cell, row-major
```

Each sheet's `V` lines follow its `S` line, and its `F` lines follow its
values. `type` is `s` (text), `n` (number), `b` (`TRUE` or `FALSE`), or `e`
(an Excel error literal such as `#DIV/0!`). Numbers are written as integer
digits when integral and otherwise as the shortest decimal that reads back as
the same double; negative zero is written `0`. Formula text has no leading
`=`.

- XLSX, XLSM, and XLAM cells contribute their stored values (cached results
  for formula cells) exactly as written, and every formula adds an `F` line.
- XLS, XLSB, and ODS cells contribute their stored values only. Dates and
  times typed from a number format are written as 1900-system Excel serial
  numbers; a date outside that range is written as ISO 8601 text.
- CSV, TSV, and TXT fields contribute their source text with type `s`, so
  `7` and `7.0` fingerprint differently. The single sheet of a delimited file
  is named after the file, so its `S` line carries an empty name and a renamed
  copy keeps its fingerprint.

`wolfxl verify <file> --fingerprint <F>` recomputes the fingerprint and
compares it with `F`, which must be `wxf1:` followed by 64 hex digits in
either case. A match prints `ok wxf1:...`; a mismatch prints `mismatch`, then
`expected wxf1:...` and `actual wxf1:...` lines.

| Exit | Meaning |
|---|---|
| 0 | The file matches the fingerprint |
| 1 | The file does not match; both fingerprints are printed |
| 2 | Unreadable input, or an unknown fingerprint version or malformed digest, described on stderr |

`wolfxl fingerprint` exits 0 on success and 2 when the input cannot be read.

## Keyed diff

`wolfxl diff LEFT RIGHT --key K` matches the rows of two tables by the
value in column `K` and reports, for each key, whether it appears on both
sides and whether the compared cells agree. Either side may be any
supported input, so a workbook can be reconciled against a CSV export.

```text
wolfxl diff LEFT RIGHT --key K [--compare C[,C2...]] [--tolerance T]
            [--key-transform none|trim|lower|digits] [--sheet S]
            [--from FMT] [--format json|text] [--strict-exit] [--quiet]
```

- **Inputs.** One of `LEFT` or `RIGHT` may be `-` for standard input, which
  needs `--from csv|tsv|xlsx|xlsm|xls|xlsb|ods`. With one side piped,
  `--from` describes that side only and the path side keeps its extension.
  With two paths, `--from` overrides both extensions. `--sheet` names the
  sheet on both sides; without it each side uses its first sheet. A
  delimited file is one sheet named after the file stem (`stdin` when
  piped).
- **Headers.** Row 1 of each sheet holds the headers. A column name matches
  its header exactly, or else the one header equal to it ignoring case. The
  key column must exist on both sides. `--compare` lists the columns to
  check, each of which must exist on both sides; by default every header
  of the left side that also exists on the right, other than the key, is
  compared. Two headers matching one name is an error.
- **Keys.** Keys keep their source spelling: `007` and `7` are different
  keys. `--key-transform` normalizes them on both sides before matching:
  `none` (the default) compares as spelled, `trim` strips surrounding
  whitespace, `lower` trims and lowercases, and `digits` keeps only the
  ASCII digits (`INV-0042` becomes `0042`). A row whose cells are all blank
  is spacing and is ignored. Any other row with an empty key after the
  transform, or a key that occurs twice on one side, stops the run with
  exit 2 and names the rows, for example
  `duplicate key "2" on left at rows 3, 5`.
- **Values.** When both cells read as numbers, they compare by value.
  Money spellings such as `$1,200.00`, `1,200`, and `(500.00)` count as
  numbers. Values equal to 15 significant digits match, a difference no
  larger than `--tolerance` (default 0) is within tolerance, and a larger
  one is outside it. Other cells must have identical text, and a text
  mismatch is always outside tolerance. Two blank cells match.

`--format json` (the default) prints one document:

```json
{
  "contract_version": 1,
  "key": "id",
  "key_transform": "none",
  "compare": ["name", "amount"],
  "tolerance": 0.01,
  "left": {"sheet": "ledger"},
  "right": {"sheet": "bank"},
  "summary": {"matched": 3, "only_left": 1, "only_right": 0, "diff": 1, "diff_outside_tolerance": 0},
  "results": [
    {"key": "1", "status": "diff", "left_row": 2, "right_row": 2, "differences": [
      {"column": "amount", "left": "100", "right": "100.004", "delta": 0.004, "within_tolerance": true}
    ]},
    {"key": "2", "status": "match", "left_row": 3, "right_row": 3, "differences": []},
    {"key": "3", "status": "match", "left_row": 4, "right_row": 4, "differences": []},
    {"key": "4", "status": "only_left", "left_row": 5, "right_row": null, "differences": []}
  ]
}
```

`results` holds every key in byte order of its transformed text. `status` is
`match`, `diff`, `only_left`, or `only_right`; `left_row` and `right_row` are
1-based sheet row numbers, `null` on the side that lacks the key.
`differences` lists the compared columns that disagree, by left header,
with each cell's source text (`null` when blank). `delta` is `right - left`
and appears only when both cells are numeric. `matched` counts keys on both
sides, `diff` the matched keys with any difference, and
`diff_outside_tolerance` the matched keys with at least one difference
outside tolerance.

`--format text` prints the summary on one line, then one line per key that
is not a clean match:

```text
matched 3, only_left 1, only_right 0, diff 1, diff_outside_tolerance 0
diff "1" (left row 2, right row 2): amount "100" -> "100.004" (delta 0.004, within tolerance)
only_left "4" (left row 5)
```

`--quiet` prints no report and leaves the result to the exit code; errors
still go to stderr.

| Exit | Meaning |
|---|---|
| 0 | Reconciled: no unmatched keys and no difference outside tolerance |
| 1 | A key is on one side only or a difference is outside tolerance; with `--strict-exit`, any difference at all |
| 2 | Usage or input failure, missing or ambiguous column, empty key, or duplicate key, described on stderr |

## Column profiles

```text
wolfxl stats FILE [--sheet S] [--from FMT] [--where EXPR]... [--select A,B]
             [--top N] [-e box|csv|json]
```

`wolfxl stats` profiles each column of one sheet, the first unless
`--sheet` names another. Row 1 holds the headers, and every header cell is
one column. `FILE` of `-` reads standard input, which needs `--from`; for a
path, `--from` overrides the extension. `--where` and `--select` work as in
`wolfxl peek`: conditions keep the data rows they accept, `--select`
profiles the named columns in the order given, and a condition may test a
column the selection leaves out. Without `--select`, columns appear in
sheet order. Every data row counts, including fully blank ones.

Each profile reports:

- **Counts.** `blank` counts empty cells and cells holding only spaces;
  `non_empty` counts the rest, so the two add up to the row count.
- **Types.** Each non-empty cell counts once as `number`, `text`, `bool`,
  `date` (dates, times, and date-times), or `error`. Types follow the
  stored values. Delimited fields are typed as described in
  [Standard input and delimited fields](#standard-input-and-delimited-fields),
  so `1200` and `2.5` are numbers and `TRUE` is a boolean, while `$1,200`
  stays text.
- **Numeric summary.** `numeric` counts the cells that read as numbers:
  stored numbers plus text in money spellings such as `$1,200.00`, `1,200`,
  and `(500.00)`, the same reading `--where` uses. `sum`, `mean`, `min`,
  `max`, and `stddev` (population standard deviation) cover those cells,
  rounded to 15 significant digits, and stay blank (`null` in JSON) when a
  column has no numbers.
- **Values.** `distinct` counts the different values and `top` lists the
  `--top N` most frequent (default 5; `--top 0` lists them all) with their
  counts, most frequent first and ties in order of first appearance. Values
  compare by their text: strings as written in the source, so `007` and `7`
  are different values, numbers in plain notation with up to 15 significant
  digits, booleans as `TRUE` or `FALSE`, dates in ISO form, and errors as
  their Excel literal such as `#DIV/0!`. Blank cells are not values.

`-e box` (the default) and `-e csv` print one row per profiled column with
the fixed columns `column`, `non_empty`, `blank`, `number`, `text`, `bool`,
`date`, `error`, `distinct`, `numeric`, `sum`, `mean`, `min`, `max`,
`stddev`, and `top`. The `top` cell joins the values as
`value (count); value (count)`. `-e json` prints one document with one
object per column, one per line, with the same order and full value text:

```json
{
  "contract_version": 1,
  "sheet": "ledger",
  "rows": 4,
  "columns": [
    {"header": "Region", "non_empty": 4, "blank": 0, "types": {"number": 0, "text": 4, "bool": 0, "date": 0, "error": 0}, "distinct": 3, "numeric": {"count": 0, "sum": null, "mean": null, "min": null, "max": null, "stddev": null}, "top": [{"value": "West", "count": 2}, {"value": "East", "count": 1}, {"value": "North", "count": 1}]},
    {"header": "Amount", "non_empty": 4, "blank": 0, "types": {"number": 1, "text": 3, "bool": 0, "date": 0, "error": 0}, "distinct": 4, "numeric": {"count": 4, "sum": 2850, "mean": 712.5, "min": -500, "max": 1200, "stddev": 707.438159841551}, "top": [{"value": "$1,200.00", "count": 1}, {"value": "950", "count": 1}, {"value": "(500.00)", "count": 1}, {"value": "1,200", "count": 1}]}
  ]
}
```

A sheet with only a header row gives profiles with zero counts, and an
empty sheet gives an empty `columns` list.

| Exit | Meaning |
|---|---|
| 0 | The profile was printed |
| 2 | Usage or input failure, a malformed `--where`, or an unknown or ambiguous column in `--where` or `--select` (stderr lists the headers) |

## Format Support

`wolfxl-cli` routes `.xlsx`, `.xlsm`, `.xlam`, `.xls`, `.xla`, `.xlsb`, and
`.ods` through the workbook reader, and treats `.csv`, `.tsv`, and `.txt` as a
single synthetic sheet named after the file stem.

Read workbook values with the independent MIT native readers and resolve
available number formats, including the OOXML style-walker fallback. Retain
historical CLI error spelling separately from canonical core error values.
For delimited text, infer numeric-looking strings without changing their values.

`map` emits workbook-level metadata for planning a next step: sheet names,
dimensions, detected sheet class, headers, and named ranges. `agent` uses the
same core metadata plus a stratified row sample to compose a briefing that fits
within a `cl100k_base` token budget. To protect that budget, `agent` keeps
compact raw numeric rendering instead of spending tokens on display symbols.
`schema` emits per-column type, cardinality, null-count, format-category, and
sample-value inference for one sheet or the whole workbook.

## Built on

- [`wolfxl-core`](https://crates.io/crates/wolfxl-core): pure-Rust
  spreadsheet reader.
- [`wolfxl-reader`](../wolfxl-reader/) for native XLSX, XLS, XLSB and ODS reads.

## License

MIT
