# `.xls` parity fixtures (Sprint Κ Pod-γ)

Legacy BIFF8 fixtures with retained authoring provenance. Current Community
XLS parity uses the independently authored xlwt fixture and literal
`crates/wolfxl-reader/qualification/expected.json` expectations instead.

| File              | Content                                                |
|-------------------|--------------------------------------------------------|
| `numbers.xls`     | A1=1, A2=2.5, A3=1e10, A4=0.5 (%), A5=$100             |
| `strings.xls`     | "hello", "münchen", multiline, empty                    |
| `dates.xls`       | date, datetime, time                                    |
| `formulas.xls`    | A1=1, A2=2, A3=A1+A2, B1=SUM(A1:A2)                     |
| `multisheet.xls`  | Sheet1 numbers, Sheet2 strings, Sheet3 empty            |

## Generation

These are produced by writing source `.xlsx` files via openpyxl and
converting with LibreOffice headless using the `MS Excel 97` filter.
LibreOffice's `.xls` export filter is fully functional (unlike its
broken `.xlsb` filter — see `xlsb/README.md`).

To regenerate:

```bash
python3 scripts/sprint_kappa_build_fixtures.py /tmp/calamine
```

(`.xlsb` fixtures are vendored from calamine's corpus; `.xls` fixtures
regenerate from the script's own `_build_*` functions.)

## Why committed, not generated at test time?

Generating `.xls` at test time would require LibreOffice on every CI
runner (~600 MB), and would re-introduce flakiness from LibreOffice
version drift (subtle changes to BIFF8 record ordering, default styles,
etc.). Committing the bytes (~30 KB total) keeps parity stable.
