# wolfxl-cli

[![crates.io](https://img.shields.io/crates/v/wolfxl-cli.svg)](https://crates.io/crates/wolfxl-cli)

Command-line previewer for spreadsheets. Installs the `wolfxl` binary and
opens `.xlsx`, `.xlsm`, `.xls`, `.xlsb`, `.ods`, `.csv`, `.tsv`, and `.txt`
inputs through the same `peek` / `map` / `agent` / `schema` workflow.

```bash
cargo install wolfxl-cli
```

## Subcommands

- `wolfxl peek <file>` - styled preview or text/csv/json/markdown export
- `wolfxl map <file>` - one-page workbook overview (sheets, dims, headers)
- `wolfxl agent <file> --max-tokens N` - token-budgeted workbook briefing
- `wolfxl schema <file>` - per-column type, cardinality, and format inference

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
wolfxl map export.csv --format text         # delimited file as one sheet
wolfxl map workbook.xlsx                    # workbook inventory for agents
wolfxl map workbook.xlsx --format text      # terminal-friendly summary
wolfxl agent workbook.xlsx --max-tokens 400 # fit a briefing to a token budget
wolfxl schema workbook.xlsx                 # JSON schema-style column summary
wolfxl schema workbook.xlsx -s "P&L" -f text
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

- [`wolfxl-core`](https://crates.io/crates/wolfxl-core) — pure-Rust
  spreadsheet reader.
- [`wolfxl-reader`](../wolfxl-reader/) for native XLSX, XLS, XLSB and ODS reads.

## License

MIT
