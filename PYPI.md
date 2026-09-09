# WolfXL Community

**High-performance, openpyxl-compatible Excel library for Python, backed by Rust.**

WolfXL is the fastest Python library to edit existing Excel (`.xlsx`, `.xlsm`) workbooks in place without rewriting from scratch. It is a drop-in openpyxl alternative designed to modify workbooks while preserving formatting, formulas, charts, drawing objects, and VBA macros.

WolfXL Community is the maintained, MIT-licensed 2.0 release line for supported workbook creation, reading, writing, streaming exports, and existing-workbook edits. It is a free product, not a trial.
## Install

```bash
python -m pip install wolfxl==2.0.2
```

Most supported openpyxl-shaped code starts with one import change:

```diff
- from openpyxl import Workbook, load_workbook
+ from wolfxl import Workbook, load_workbook
```

Start with one representative workbook, define the output that must remain
correct, and prove that bounded operation before replacing a production path.
Review the [compatibility matrix](https://wolfxl.com/docs/migration/compatibility-matrix/?utm_source=pypi&utm_medium=registry&utm_campaign=community_commercial_2026_09)
and [known limitations](https://wolfxl.com/docs/trust/limitations/?utm_source=pypi&utm_medium=registry&utm_campaign=community_commercial_2026_09).

## Compare openpyxl alternatives, measured

WolfXL Community is benchmarked against the maintained open-source Python Excel
libraries on a committed run. On 2026-08-18 (AMD EPYC 9655, Python 3.13.15,
median of 5 rounds, 200,000-row x 8-column workbook), Community 2.0.1 wrote 1.6
million cells in 0.73 s (11.10x openpyxl 3.1.5, 6.46x XlsxWriter 3.2.9, 5.01x
PyExcelerate 0.13.0) and read them back in 0.39 s (11.85x openpyxl), with the
lowest read peak memory in the field. Raw results and the harness are published
under [`benchmarks/`](https://github.com/SynthGL/wolfxl-oss/tree/main/benchmarks).

Full tables, capability scope, and reproduction commands:
[openpyxl alternatives, measured](https://wolfxl.com/openpyxl-alternatives?utm_source=pypi&utm_medium=registry&utm_campaign=community_commercial_2026_09).

## Choose the right edition

Community is the free option when supported workbook I/O is sufficient.
Commercial 2.1+ is a separate option for the following workflow requirements:

| Requirement | Community 2.0 line | Commercial 2.1+ |
| --- | --- | --- |
| Supported workbook I/O | Included within the documented surface | Included |
| Native formula recalculation | Not included | [Calculate a supported formula set](https://wolfxl.com/calculate-excel-formulas-python?utm_source=pypi&utm_medium=registry&utm_campaign=community_commercial_2026_09) |
| PDF or image rendering | Not included | [Render supported output](https://wolfxl.com/render-excel-python?utm_source=pypi&utm_medium=registry&utm_campaign=community_commercial_2026_09) |
| Format conversion | Not included | [Scope an evaluation](https://wolfxl.com/pilot?utm_source=pypi&utm_medium=registry&utm_campaign=community_commercial_2026_09) |
| VBA or Power Query operations | Not included | [Scope an evaluation](https://wolfxl.com/pilot?utm_source=pypi&utm_medium=registry&utm_campaign=community_commercial_2026_09) |
| Production operations | Not included | [Scope an evaluation](https://wolfxl.com/pilot?utm_source=pypi&utm_medium=registry&utm_campaign=community_commercial_2026_09) |
| Support | Public Community documentation and issues | [Scope an evaluation](https://wolfxl.com/pilot?utm_source=pypi&utm_medium=registry&utm_campaign=community_commercial_2026_09) |

For existing-template work, review the [bounded preservation approach](https://wolfxl.com/openpyxl-preservation?utm_source=pypi&utm_medium=registry&utm_campaign=community_commercial_2026_09).
Use the [local fit check](https://wolfxl.com/fit-check?utm_source=pypi&utm_medium=registry&utm_campaign=community_commercial_2026_09)
to scope a representative-workbook test. Commercial self-service pricing is
[$30/month or $299/year for one seat](https://wolfxl.com/pricing?utm_source=pypi&utm_medium=registry&utm_campaign=community_commercial_2026_09).
