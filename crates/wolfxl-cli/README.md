# wolfxl-cli

[![crates.io](https://img.shields.io/crates/v/wolfxl-cli.svg)](https://crates.io/crates/wolfxl-cli)

Command-line previewer and analyst toolkit for spreadsheets. Installs the
`wolfxl` binary and opens `.xlsx`, `.xlsm`, `.xls`, `.xlsb`, `.ods`, `.csv`,
`.tsv`, and `.txt` inputs through the same `peek` / `map` / `agent` /
`schema` workflow, plus `calc`, `pivot`, `stats`, `diff`, `fingerprint`, and
`verify` for analysis and reconciliation. A built-in formula evaluator powers
`calc` and `peek --add` formula columns, and `list-functions` lists the
functions it implements. `wolfxl mcp` serves the same reads, plus edits to
existing `.xlsx`/`.xlsm` workbooks, to Model Context Protocol clients.

## Install

Prebuilt binaries are available for Linux (x86_64 and ARM64, statically linked
with musl), macOS (Intel and Apple silicon), and Windows (x86_64).

Linux and macOS:

```sh
curl -fsSL https://raw.githubusercontent.com/SynthGL/wolfxl-oss/main/install.sh | sh
```

Windows PowerShell:

```powershell
irm https://raw.githubusercontent.com/SynthGL/wolfxl-oss/main/install.ps1 | iex
```

The installers verify the archive against the release's `SHA256SUMS` before
installing `wolfxl` to `$HOME/.local/bin` (or `$HOME\.local\bin` on Windows).
Add that directory to your `PATH` if prompted. The Unix installer requires
`curl`, `tar`, and `sha256sum` or `shasum`. Windows requires PowerShell 5.1 or
newer.

Set `WOLFXL_VERSION` to a CLI version, such as `0.11.1`, to pin the download
and skip version discovery. Set `WOLFXL_INSTALL_DIR` to choose a different
destination. For a release mirror, `WOLFXL_RELEASE_BASE_URL` overrides
`https://github.com/SynthGL/wolfxl-oss/releases/download`; the mirror must serve
`wolfxl-cli-v<version>/SHA256SUMS` and the matching archive. Without a pinned
version, discovery still queries the official GitHub releases API. CLI releases
are separate from the repository's Python package releases.

With [cargo-binstall](https://github.com/cargo-bins/cargo-binstall):

```sh
cargo binstall wolfxl-cli
```

Or compile from source with a Rust toolchain:

```sh
cargo install --locked wolfxl-cli
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
- `wolfxl mcp [--root DIR] [--read-only]` - serve workbook reads, analysis, and edits as MCP tools over standard input/output

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
wolfxl list-functions --check SUM,IF,LOOKUP  # exit 1: LOOKUP is not built in
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
wolfxl mcp --root ~/workbooks --read-only   # MCP server confined to one directory
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

Standard input and extension-overriding `--from` copies stream to a private
temporary file, with a 512 MiB source-byte limit for XLS and 4 GiB for the
other formats. These named CLI spool limits match the XLS reader's source
limit and the other readers' default uncompressed-package budget,
respectively; the 4 GiB limit also applies to CSV/TSV. Exceeding a spool
limit exits 2 and names the limit. The limits are not changed by
`WOLFXL_MAX_ZIP_TOTAL_BYTES`, and do not add a source-byte limit to files
opened directly without an extension override. On Unix, spool directories
are created owner-only (0700) and files owner-only (0600), and are removed
after use or a failed copy.

Delimited header names retain their source spelling: `001`, `01`, and `1`
are distinct fields for `--rows`, `--values`, `--where`, `--select`, and
`--key`. Data fields arrive as text. `pivot` and `stats` trim surrounding
whitespace and read `TRUE` and `FALSE` in any letter case as booleans, and
plain decimals as numbers: an optional sign, digits with an optional decimal
point (`42`, `-1.5`, `.5`), and an optional exponent (`-1.5e3`). `calc`, the
formulas of `peek --add`, and the aggregates of `pivot` and `stats` also read
whole-field money spellings as numbers, using the same parser as `--where`:
leading signs and currency symbols, grouped thousands commas (`$1,200.00`),
and accounting negatives (`(75.00)`). For `calc` and `peek --add`, a field in
ISO date form (`2008-01-01`, `2008-01-01T12:00:00`) is the date's serial
number, as Excel's text import reads it, so `XIRR` and the date functions take
a column of ISO dates. Other text, including hex (`0x10`), `inf`, and `NaN`,
stays text.
Text cells stored in an Excel workbook stay text, even when they look like
money; range functions such as `SUM` ignore them. `peek` output and filters,
`diff`, and `fingerprint` work from the source text, so `007` and `1,200`
print and match as written.

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
  and date-formatted numbers retain their stored numeric serials without
  converting between date systems or rounding to whole seconds. Date
  functions read and build serials in the workbook's own date system, 1904
  included. References use the sheet's own addresses, so a table that starts
  at `C3` is read at `C3`.
- **Built-in engine.** The evaluator covers arithmetic (`+`, `-`, `*`, `/`,
  `^`, unary minus, and postfix `%`), comparison (`=`, `<>`, `<`, `>`, `<=`,
  `>=`), and `&` concatenation over numbers, strings, `TRUE`/`FALSE`, and
  error literals such as `#N/A`. It reads A1 references with or without `$`,
  sheet-qualified references such as `'P&L'!B3`, and ranges such as `B2:D10`,
  `B:B`, and `2:2`; whole columns and rows stop at the sheet's used extent.
- **Functions.** 133 functions in eight categories:
  - `math`: `ABS`, `ACOS`, `ASIN`, `ATAN`, `ATAN2`, `CEILING`, `COS`,
    `DEGREES`, `EXP`, `FLOOR`, `INT`, `LN`, `LOG`, `LOG10`, `MOD`, `PI`,
    `POWER`, `PRODUCT`, `RADIANS`, `RAND`, `RANDBETWEEN`, `ROUND`,
    `ROUNDDOWN`, `ROUNDUP`, `SIGN`, `SIN`, `SQRT`, `SUBTOTAL`, `SUM`,
    `SUMIF`, `SUMIFS`, `SUMPRODUCT`, `TAN`, `TRUNC`
  - `information`: `ISBLANK`, `ISERROR`, `ISNA`, `ISNUMBER`, `ISTEXT`, `NA`
  - `statistical`: `AVERAGE`, `AVERAGEIF`, `AVERAGEIFS`, `COUNT`, `COUNTA`,
    `COUNTBLANK`, `COUNTIF`, `COUNTIFS`, `LARGE`, `MAX`, `MAXIFS`, `MEDIAN`,
    `MIN`, `MINIFS`, `MODE`, `MODE.SNGL`, `PERCENTILE`, `PERCENTILE.EXC`,
    `PERCENTILE.INC`, `QUARTILE`, `QUARTILE.EXC`, `QUARTILE.INC`, `RANK`,
    `RANK.AVG`, `RANK.EQ`, `SMALL`, `STDEV`, `STDEV.P`, `STDEV.S`, `STDEVP`,
    `VAR`, `VAR.P`, `VAR.S`, `VARP`
  - `logical`: `AND`, `IF`, `IFERROR`, `IFNA`, `IFS`, `NOT`, `OR`, `SWITCH`
  - `lookup`: `CHOOSE`, `FILTER`, `HLOOKUP`, `INDEX`, `MATCH`, `OFFSET`,
    `SORT`, `UNIQUE`, `VLOOKUP`, `XLOOKUP`
  - `text`: `CONCAT`, `CONCATENATE`, `EXACT`, `FIND`, `LEFT`, `LEN`, `LOWER`,
    `MID`, `REPT`, `RIGHT`, `SUBSTITUTE`, `TEXT`, `TEXTJOIN`, `TRIM`, `UPPER`,
    `VALUE`
  - `date`: `DATE`, `DAY`, `DAYS`, `EDATE`, `EOMONTH`, `HOUR`, `MINUTE`,
    `MONTH`, `NETWORKDAYS`, `NOW`, `SECOND`, `TODAY`, `WORKDAY`, `YEAR`
  - `financial`: `DB`, `FV`, `IRR`, `NPER`, `NPV`, `PMT`, `PV`, `RATE`,
    `SLN`, `XIRR`, `XNPV`

  `IF`, `IFERROR`, `IFNA`, `IFS`, `SWITCH`, and `CHOOSE` evaluate only the
  argument they return. `COUNT` ignores errors, including computed errors;
  the other aggregates propagate errors. The statistical functions read
  numbers the way `SUM` does: ranges skip text, booleans, and blanks.
  Criteria in `COUNTIF`, `SUMIF`, and the other `*IF` and `*IFS` functions
  take an optional comparison prefix (`">100"`, `"<>West"`)
  and the `*`, `?`, and `~` wildcards. `MATCH`, `VLOOKUP`, `HLOOKUP`, and
  `XLOOKUP` take Excel's match modes. `TEXT` covers number, percent,
  scientific, and date and time format codes. `TODAY` and `NOW` read the
  local clock. An empty argument counts as a blank, so `MATCH(x,A:A,)` is an
  exact match, while `XLOOKUP` and `OFFSET` give empty optional arguments
  their defaults. Names stored with the `_xlfn.` or `_xlws.` prefix, such as
  `_xlfn.STDEV.S`, resolve to the same functions. A formula naming any other
  function exits 2, naming it and pointing to
  [`wolfxl list-functions`](#function-registry).
- **Math and information.** Angles are in radians; `ATAN2(x, y)` follows
  Excel's x-first argument order. `LOG` defaults to base 10. `CEILING`
  and `FLOOR` use the legacy Excel sign rules; `TRUNC` cuts toward zero.
  `PRODUCT` ignores text, booleans, and blanks in ranges and arrays, but
  accepts numeric text and booleans supplied directly. `ISBLANK`,
  `ISERROR`, `ISNA`, `ISNUMBER`, and `ISTEXT` inspect types without
  coercion and preserve the shape of array arguments. `NA()` returns `#N/A`.
  Numeric expectations and array elements come from the committed
  [native Excel capture](tests/fixtures/math-information.oracle.json).
- **Randomness.** `RAND()` returns a fresh non-cryptographic number in
  [0, 1); `RANDBETWEEN(bottom, top)` returns an integer using Excel's
  endpoint rounding. Ordered fractional bounds containing no integer yield
  `CEILING(bottom, 1)`. Fingerprinting and verification always use the
  workbook's stored values and formulas, without recalculating randomness.
- **Google Sheets names.** `ARRAYFORMULA` (Sheets array evaluation),
  `SPARKLINE` (in-cell chart rendering), `SPLIT` (Sheets splitting rules),
  `REGEXMATCH` (Sheets regex dialect), `COUNTUNIQUE` (Sheets counting rules),
  and `AVG` (non-Excel alias) are deliberately unsupported. Use an Excel
  formula supported by the registry. A call to any of these names exits 2
  with an unsupported-function diagnostic.
- **Arrays.** Operators work element by element when an operand is a range
  or array, as in Excel: `=B2:B10*C2:C10` is an array of products,
  `=SUMPRODUCT((A2:A10="North")*C2:C10)` and `=SUMPRODUCT(--(C2:C10>100))`
  count and sum by condition, and `=FILTER(A2:A10,C2:C10>100)` keeps the
  matching rows. A scalar pairs with every element, and a single row or column
  repeats to match the other operand; positions outside both shapes are
  `#N/A`. Ranges are not implicitly intersected, and results never spill.
- **`SUBTOTAL`.** Function numbers 1 to 11 and 101 to 111 select `AVERAGE`,
  `COUNT`, `COUNTA`, `MAX`, `MIN`, `PRODUCT`, `STDEV`, `STDEVP`, `SUM`,
  `VAR`, and `VARP`. Cells that themselves hold a `SUBTOTAL` formula, such as
  a table's totals row, are skipped, so `=SUBTOTAL(109,F3:F15)` over a column
  ending in its own total does not count it twice.
- **Semantics.** Arithmetic uses `f64`. Arithmetic text-to-number coercion
  accepts the same decimal, exponent, and money spellings as `--where`.
  Text comparisons ignore case by Unicode lowercase ordering rather than
  locale-specific collation, and `&` joins numbers in plain decimal form
  with up to 15 significant digits rather than with cell number formats.
- **Structured references.** Formulas can name a workbook table from any
  sheet: `Sales[Amount]` is the column's data rows, without the header or
  totals row; `Sales[]` and `Sales[#Data]` are every data row;
  `Sales[#All]`, `Sales[#Headers]`, and `Sales[#Totals]` select those rows;
  `Sales[[#Headers],[#Data],[Amount]]` and `Sales[[#Data],[#Totals],[Amount]]`
  combine them; and `Sales[[#Data],[Q1]:[Q4]]` spans a column range. Each
  resolves to the cells it covers when the formula is evaluated, so it works
  wherever a range does: `SUM(Sales[Amount])`, `Sales[[#Totals],[Amount]]`,
  `INDEX(Sales[Amount],2)`. Table, column, and item names ignore case, and
  within a column name `'` escapes `[`, `]`, `#`, and `'`, as in
  `Sales[Q1 '[est']]`. Tables come from the table parts of `.xlsx`, `.xlsm`,
  and `.xlsb` workbooks and resolve by the name stored in the table part;
  other inputs have none. A table missing from the workbook's table parts,
  an unknown column, and `#Headers` or `#Totals` of a table without that row
  are `#REF!`. Unlike Excel, `Sales[@Amount]` and `[#This Row]` are
  `#VALUE!`, because a formula evaluated here has no row of its own. A
  bracketed column without a table name, such as `[Amount]`, and item
  combinations Excel rejects, such as `[#Headers],[#Totals]`, exit 2.
- **Differences from Excel.**
  - `FILTER` with no match and no `if_empty` argument, and `UNIQUE` with
    nothing to return, give `#VALUE!` where Excel gives `#CALC!`.
  - Blank cells stay blank inside array results (`FILTER`, `UNIQUE`, `SORT`,
    and bare ranges) instead of becoming 0.
  - `SUBTOTAL` 101 to 111 skip hidden rows only in `.xlsx`, `.xlsm`, and
    `.xlsb` workbooks; other formats record no hidden rows. Rows hidden by a
    filter and rows hidden by hand are not told apart, so 1 to 11 include
    both and 101 to 111 skip both, where Excel's 1 to 11 skip filtered rows.
  - `VALUE` reads numbers, percentages, `m/d/y`, `y-m-d`, `d-Mon-y`,
    `Mon d, y`, and `Mon y` dates, and `h:mm[:ss]` times with optional
    `AM`/`PM`, in US order only. A date without a year, such as `"3/14"`, is
    `#VALUE!` rather than a date in the current year.
  - `RATE` and `XIRR` start from the guess with Newton's method and, when
    that does not converge, search rates from -0.99 to 10 for the sign
    change nearest the guess, so they can return a root where Excel returns
    `#NUM!`.

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
wolfxl calc '=SUMIFS(C:C,A:A,"North",B:B,"<>Apple")' sales.csv
wolfxl calc '=XLOOKUP("ACME",A:A,D:D,"missing")' deals.csv
wolfxl calc '=PMT(B2/12,B3,B4)' loan.csv
wolfxl calc '=TEXT(EOMONTH(A2,1),"yyyy-mm-dd")' invoices.xlsx
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
ABS             math
ACOS            math
AND             logical
AVERAGE         statistical
...
XNPV            financial
YEAR            date
133 functions
```

- **Categories.** The registry's categories are `date`, `financial`,
  `information`, `logical`, `lookup`, `math`, `statistical`, and `text`. `--category` keeps
  one of them, ignoring case and surrounding spaces, so `Math` works. Any
  other category exits 2 and lists the valid ones.
- **Search.** `--search TEXT` keeps the functions whose name contains
  `TEXT`, ignoring case: `--search look` finds `HLOOKUP`, `VLOOKUP`, and
  `XLOOKUP`. With both flags, a function must pass both.
- **Text output.** Each line is the name padded to the longest listed name,
  two spaces, and the category. The last line is the count: `N functions`
  for the full list, or `M of N functions` when a filter is active. A filter
  that matches nothing prints only the count line and still exits 0.

`--json` prints one document. `total` is the size of the whole registry and
`count` the number of listed entries; `category` and `search` echo the
filters (`null` when absent), with `category` in its registry spelling. Each
listing entry also has a `volatile` boolean. `RAND`, `RANDBETWEEN`, `NOW`,
`TODAY`, and `OFFSET` are volatile.

```json
{"contract_version": 1, "total": 133, "count": 3, "category": "lookup", "search": "look",
 "functions": [{"name": "HLOOKUP", "category": "lookup", "volatile": false}, {"name": "VLOOKUP", "category": "lookup", "volatile": false},
               {"name": "XLOOKUP", "category": "lookup", "volatile": false}]}
```

`--check` takes function names, comma-separated or repeated, and reports
each in the order given. Names ignore case and surrounding whitespace and
print in upper case. A name is `supported` when the registry lists it, and
`unsupported` otherwise, which includes spreadsheet functions the built-in
engine does not implement, such as `LOOKUP`. `--check` cannot be combined
with `--category` or `--search`.

```text
$ wolfxl list-functions --check sum,lookup,nope
SUM     supported (math)
LOOKUP  unsupported
NOPE    unsupported
1 of 3 supported
```

With `--json`, `total` is the size of the registry, `checked` counts the names,
`supported` the supported ones, and `category` is `null` for an unsupported name:

```json
{"contract_version": 1, "total": 133, "checked": 2, "supported": 1, "results": [
  {"name": "SUM", "status": "supported", "category": "math"},
  {"name": "LOOKUP", "status": "unsupported", "category": null}
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
above). By default the first used row supplies the headers. `--header-row N`
instead uses sheet row N and discards preceding rows. N is an absolute,
one-based sheet row, not an offset into the used range; an out-of-range N
exits 2 and reports the used rows. This also applies to `stats` and `pivot`.

`--where COLUMN<op>VALUE` keeps the data rows that satisfy the condition;
the header row is always kept. Repeat `--where` to require every condition
(AND). The operators are `=`, `!=`, `<`, `>`, `<=`, `>=`, and `~`
(contains). The first operator character in the expression ends the column
name, and `<=`, `>=`, and `!=` are read as one operator. Spaces around the
column name are ignored; the value is used as written.

- When the value reads as a number, the comparison is numeric. Money
  spellings count on both sides: `$1,200.00`, `1,200`, and the accounting
  negative `(500.00)`. A trailing percent divides by 100: `12%` is `0.12`,
  `-3.5%` is `-0.035`, `(12%)` is `-0.12`, and `1,200%` is `12`.
  This same rule types delimited fields for `calc` and `--add`; malformed
  grouping (`1,23,4`) and mixed text remain text. Source spelling, including
  leading zeroes, is preserved in `peek` output.
  A cell that holds no number fails every numeric comparison except `!=`.
- Any other value compares as text, ignoring case, against the cell's
  displayed value: strings as stored, numbers in plain notation, booleans as
  `TRUE` or `FALSE`. `<` and `>` order text lexically.
- `~` always tests text: the cell contains the value, ignoring case.

`--select A,B` prints only the named columns, in the order given. A filter
may test a column that `--select` leaves out. Column resolution first accepts
a unique exact header, then a sheet column letter inside the original used
range (`C`, `F`, `AA`), then a unique case-insensitive header. Duplicate
header names are refused with their header-cell addresses: use a letter to
disambiguate, such as `--header-row 2 --select C,F` when both C2 and I2
contain `City`. Letters always refer to the original sheet, even after
selection reorders columns; computed columns are selected by their names.

`--addresses` labels box, text, and markdown previews with sheet column
letters above the headers and sheet row numbers on the left. Filtering
retains original row numbers and selection retains original letters, so
these labels are valid `wolfxl calc` coordinates. Computed `--add` columns
are marked `+`: they have no stored sheet address. CSV and JSON reject
`--addresses` with exit 2 rather than changing their machine-readable shape.

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

Write the formula for the first data row, the row under the chosen header
(row 2 for a sheet that starts in A1, or row N+1 with `--header-row N`).
Each later row gets it filled down the way a
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
result leaves the cell empty, and an array result, such as the range
`B2:C2` or a `FILTER`, `UNIQUE`, or `SORT` call, is `#VALUE!`, since
ranges are not intersected with the row and results do not spill. An
Excel error such as `#DIV/0!` stays in its cell, shown as `ERROR: #DIV/0!`
in the text views and as `"#DIV/0!"` in JSON, and `--where 'Unit=#DIV/0!'`
finds it.

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
The cursor also shows its original sheet address (for example `F3`), retained
through filtering, selection, and sorting; computed columns have no address.
Status notices and cursor position take priority over the path and sheet
description. Long paths elide leading directories first, and notices that
cannot fit end with `…`.

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
             [--sheet S] [--header-row N] [--from FMT] [-e box|csv|json]
```

`wolfxl pivot` groups the data rows of one sheet (the first unless `--sheet`
names another) and aggregates value fields per group. The first used row
holds headers unless `--header-row N` selects another absolute sheet row
and discards rows above it. Out-of-range N exits 2 with the used row range.
`FILE` of `-` reads standard input, which needs `--from`;
delimited fields are typed as described in
[Standard input and delimited fields](#standard-input-and-delimited-fields).
Fields in `--rows`, `--column`, and `--values` accept unique exact headers,
then original sheet column letters, then unique case-insensitive headers.
Duplicate names are refused with their header-cell addresses; use a letter.
`--rows` and `--values` take comma-separated lists or repeat. Fully blank
rows are skipped. A stored totals row is ordinary input data, not an Excel
table total to exclude automatically.

| Aggregation | Result |
|---|---|
| `sum` | Sum of the numeric cells |
| `count` | Number of non-empty cells, numeric or not |
| `avg` | Mean of the numeric cells |
| `min` | Smallest numeric cell |
| `max` | Largest numeric cell |

Numeric cells are stored numbers and text in the money spellings that
`--where` and `stats` read, such as `$1,200.00`, `1,200`, and `(500.00)`.
Other text, booleans, dates, and errors are
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
- XLS, XLSB, and ODS cells contribute their raw stored values only.
  Date-formatted numbers retain the exact stored numeric value, independent
  of the number format and workbook date system.
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
wolfxl stats FILE [--sheet S] [--header-row N] [--from FMT] [--where EXPR]... [--select A,B]
             [--top N] [-e box|csv|json]
```

`wolfxl stats` profiles each column of one sheet, the first unless
`--sheet` names another. The first used row holds headers, or `--header-row N`
selects absolute sheet row N and discards preceding rows. An N outside the
used range exits 2 naming the used rows. Every header cell is one column.
`FILE` of `-` reads standard input, which needs `--from`; for a
path, `--from` overrides the extension. `--where` and `--select` work as in
`wolfxl peek`: conditions keep the data rows they accept, `--select`
profiles the named columns in the order given, and a condition may test a
column the selection leaves out. Without `--select`, columns appear in
sheet order. Every data row counts, including fully blank ones.
Column names resolve as for `peek`: unique exact header, original sheet
letter, then unique case-insensitive header; duplicate names report their
header-cell addresses. Stored totals rows count as data, so profiling a
table including its total can count those values twice.

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

## MCP server

```text
wolfxl mcp [--root DIR] [--read-only]
```

`wolfxl mcp` is a Model Context Protocol server on standard input and output:
JSON-RPC 2.0, one message per line, with logs on stderr only. It negotiates
protocol version `2025-06-18` and also accepts `2025-03-26` and `2024-11-05`.
The server is file based: every tool takes a `path` to a workbook on disk and
reads it fresh on each call.

Register it with Claude Code:

```bash
claude mcp add wolfxl -- wolfxl mcp
```

Any client that reads an `mcpServers` configuration can start it the same way:

```json
{
  "mcpServers": {
    "wolfxl": {
      "command": "wolfxl",
      "args": ["mcp", "--root", "/path/to/workbooks"]
    }
  }
}
```

| Tool | Purpose |
|---|---|
| `get_workbook` | Sheets with dimensions and headers, named ranges, format, and `revision` |
| `read_range` | Stored values and formula text for an inclusive A1 range, at most 10,000 cells per call |
| `calc` | Evaluate one formula against a sheet, as `wolfxl calc --json` does |
| `stats`, `pivot`, `diff`, `fingerprint` | The JSON documents of the matching commands |
| `write_cells` | Batch numbers, text, booleans, formulas, and clears (`null`) |
| `set_number_format` | Apply an Excel number format code such as `$#,##0.00` to a range |
| `insert_rows`, `delete_rows`, `insert_columns`, `delete_columns` | Structural edits at a 1-based index |
| `add_sheet`, `rename_sheet` | Add a blank sheet (optionally at a 0-based position) or rename one |
| `undo`, `redo` | Restore the exact prior bytes of a file this server edited |

Reads accept every format the CLI reads. Edits require an existing `.xlsx` or
`.xlsm` file and never create one. Each tool returns its result as structured
content plus a text block holding the same JSON; a failed call returns
`isError: true` with an `error` message, and malformed arguments or unknown
tools are JSON-RPC invalid-params errors.

**Paths and `--root`.** With `--root DIR`, relative paths resolve under `DIR`
and any path whose canonical form lies outside it is refused, including paths
that leave through `..` or a symbolic link. Without `--root`, relative paths
resolve against the server's working directory.

**`--read-only`.** Lists only `get_workbook`, `read_range`, `calc`, `stats`,
`pivot`, `diff`, and `fingerprint`; calls to any editing tool are rejected.

**Revisions.** `revision` is the lowercase hex SHA-256 of the file's bytes.
Every editing tool accepts `expected_revision` and refuses to write when the
file no longer matches it, so pass the revision from your last read to avoid
overwriting a change made in the meantime. A successful edit returns the new
`revision`, `dry_run: false`, and the `change` it applied. The check is best
effort: the server rereads the file just before replacing it, but a process
that replaces the file in the instant between that read and the rename is not
detected. Do not rely on it as a lock between concurrent writers.

**Dry runs.** `dry_run: true` validates the edit and returns the planned
`change` and the revision the file would have, without writing the file or
recording history. Applying the same edit afterwards produces that revision.

**Undo and redo.** The server keeps, per path, the bytes each of its edits
replaced, in memory for the life of the process. `undo` and `redo` restore
those exact bytes, a new edit clears the redo history, and both refuse when
the file has changed on disk since the server last wrote it. Each path keeps
at most the 32 newest snapshots and at most 256 MiB of them; older snapshots
are dropped, so `undo` reports nothing to undo once it reaches them. The newest
snapshot is always kept, even when it alone exceeds 256 MiB.

**What an edit changes.** Edits rewrite only the package parts they touch;
every other part of the file, such as images, embedded objects, and VBA
projects, is carried over byte for byte. Written formulas are stored without a
cached result. When an edit writes a formula, or the workbook already holds
formulas whose cached results the edit could make stale, the workbook is
marked for a full recalculation, so Excel computes them when it opens the
file. Until then, `read_range` reports new formulas as `null`, existing
formulas keep their old cached values, and `calc` reads both as stored.

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
