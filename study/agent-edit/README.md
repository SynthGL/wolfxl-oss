# One-cell workbook editing study

This harness measures what remains in an Excel workbook after the same edit:
write `=1+1` to `Z1` on the first worksheet, save, and recalculate. It compares
OOXML package parts, feature-family counts, worksheet features, and formula
caches against each unmodified source. It is a fixed workflow harness, not a
live interactive system: no prompts, task planning, or autonomous tool selection
are involved.

## Three workflows

| Command | Edit and save | Recalculation | Package required |
| --- | --- | --- | --- |
| `run.py skill` | openpyxl, preserving VBA for `.xlsm` | Pinned reference `recalc.py` using LibreOffice | Community is sufficient |
| `lo_path.py` | WolfXL modify mode | The same reference script and LibreOffice | Community is sufficient |
| `run.py wolfxl` | WolfXL modify mode | Native `Workbook.calculate()` before save | Commercial 2.2.0+ wheel |

With no arguments, `run.py` runs both `skill` and `wolfxl`. Do not use that
combined command with a Community wheel: the native calculation workflow needs
Commercial 2.2.0 or later, where `calculate()` persists formula caches on save.
Community evaluates formulas without persisting those caches, so the end-to-end
workflow rejects Community with an explicit error; use `lo_path.py` instead.
Commercial 2.1.1 also lacks the required cache persistence. The `skill` name is
retained as a stable command-line label for the
reference recipe, not a claim that the harness runs an interactive system.

## Reproduce

Prerequisites: `uv`, Python 3.12 or later, and LibreOffice with `soffice` on `PATH`.
The commands below are for POSIX shells (macOS or Linux), from a fresh clone.
The source workbooks are copied; originals are never edited.

```sh
git clone https://github.com/SynthGL/wolfxl-oss.git
cd wolfxl-oss/study/agent-edit
uv venv --python 3.13 .venv
uv pip install --python .venv/bin/python 'wolfxl==2.0.8' 'openpyxl==3.1.5' defusedxml
soffice --version
.venv/bin/python fetch.py
.venv/bin/python run.py skill
.venv/bin/python lo_path.py
```

The fetch command downloads two hash-verified reference script files at commit
[`33375500bcea98d610eb30ce10ac4e59b89c390d`](https://github.com/anthropics/skills/tree/33375500bcea98d610eb30ce10ac4e59b89c390d/skills/xlsx/scripts)
and the Microsoft sample described below. Both workflows also fetch any missing
inputs automatically. No additional Python packages beyond the three above are
needed on a normal macOS or Linux installation. The upstream LibreOffice helper
may need `gcc` if a Linux sandbox blocks Unix sockets. Internet access is required
on the first run; cached files are verified again on subsequent runs.

For Commercial, use a separate environment containing a released Commercial
2.2.0+ wheel whose `calculate()` persists formula caches on save, installed through
your existing authorized package mechanism, together with `openpyxl==3.1.5` and
`defusedxml`. The workflow deliberately keeps calculation before saving; it does
not replace formulas with their calculated values. Then run:

```sh
.venv/bin/python run.py wolfxl
# Or compare both workflows in that Commercial environment:
.venv/bin/python run.py skill wolfxl
```

No credentials or private package addresses are needed or recorded by this
harness. It records installed package versions rather than assuming an edition
from the command name.

## Input corpus and provenance

Six workbooks are the public fixtures under
[`tests/fixtures/external_oracle/`](../../tests/fixtures/external_oracle):

- `real-excel-pivot-chart-slicers.xlsx`
- `real-excel-timeline-slicer.xlsx`
- `real-excel-normalized-pivot-cf-table.xlsx`
- `real-excel-p1-comments-validation-protection.xlsx`
- `real-excel-chart-cf-basic.xlsx`
- `real-excel-macro-basic.xlsm`

`x14-validation-sparkline.xlsx` is the synthetic eighth-feature probe included
here, containing x14 validation and sparkline extensions.

The remaining workbook is Microsoft's Excel 2013 Contoso profit-and-loss
PowerPivot sample. It is **not rehosted** in this repository. Its official
[Download Center page](https://www.microsoft.com/en-us/download/details.aspx?id=38838)
provides
[`ContosoPnL_Excel2013.zip`](https://download.microsoft.com/download/b/e/c/becf5873-6b88-4920-9096-2c10ba98de60/ContosoPnL_Excel2013.zip).
`fetch.py` extracts its sole `.xlsx` member into the ignored `.cache/` directory
as `real-excel-powerpivot-contoso-pnl.xlsx`, preserving the original bytes.
The workbook must be 7,361,732 bytes with SHA-256:

```text
d9a77819ed43a93cda82ab7ba08dd6406981d9d3eae7cfb7cae4720a40f87ac5
```

A SHA-256 mismatch stops the run, including for a previously cached workbook.
If Microsoft's download is unavailable, the harness prints a clear note and
runs the seven available files; JSON marks `contoso_available` as false. Do not
compare a seven-file run to an eight-file total without noting the omission.

## Outputs and interpretation

`run.py` writes `results.json` and edited workbooks in `out/`. `lo_path.py`
writes `results_lo.json` and workbooks in `out_lo/`. These local outputs and the
cache are ignored. Each invocation overwrites its matching output files and
receipt; save copies outside those locations before another run if needed.
No historical receipts are bundled here.

Every receipt records Python, openpyxl, LibreOffice, the reference commit, source
SHA-256 values, and (where used) WolfXL's version. Per-file results include:

- `parts_missing`: source ZIP entries absent from the output, excluding
  `xl/calcChain.xml`, which Excel can rebuild. A renamed part can appear missing;
  this is a package census, not proof that its user-visible content disappeared.
- `parts_unchanged`: source entries whose bytes are unchanged in the output.
- `families_lost`: primary-part counts `[source, output]` for VBA, pivots, pivot
  caches, charts, slicers, slicer caches, timelines, timeline caches, external
  links, drawings, tables, comments, and PowerPivot, where the count decreased.
  Relationship parts are excluded from these counts.
- `sheet_features`: source/output counts for x14 elements, conditional-formatting
  blocks (`cf`), and data-validation blocks (`dv`) across all worksheet parts;
  `x14_digest` hashes sorted tags, attributes except `id`, and `xm` text.
  `sheet_features_changed` names differing fields. The digest is a change probe,
  not semantic equivalence or a rendering test.
- `formula_cells_checked`: source formula-cell count.
- `cached_value_regressions` and `cached_value_regression_samples`: previously
  non-error cached formula values that become errors, with at most ten examples.
  Missing caches or removed formulas are not counted by this field.
- `cached_value`: the output's cached `Z1` value, expected to be `2`.
- `edit_step` or `edit_step_parts_missing`: losses after editing but before
  LibreOffice, so edit losses can be distinguished from recalc/save losses.
- `recalc`: the reference script's status, error census, or failure details.
  `errors_found` includes existing formula errors and does not by itself mean
  the edit introduced them. Commercial instead reports the calculation report
  type under `calculate`.

Execution errors are recorded rather than silently discarded. The commands
exit nonzero if a workflow or recalculation fails, or the edited cell does not
have cached value `2`. Feature losses are observations, not execution failures.

## Limits

This study has **eight files, one edit per file, and one run per workflow**. The
corpus is intentionally feature-rich and not a representative random sample of
all spreadsheets. Results apply to the recorded tool versions and this exact
edit; they do not establish universal preservation, speed, formula coverage,
visual fidelity, or compatibility with Excel's complete behavior. No elapsed-time
benchmark or Excel UI validation is performed. A retained ZIP part is not proof
that it still functions. Conversely, a missing part can be renamed or rebuilt.
The x14 probe is synthetic, and recalculation itself can rewrite unsupported
features independently of the edit library.

## Focused checks

Repository CI lints `python/`, `tests/`, and this study directory. Its Python job
also runs the offline study tests, covering fetch failure boundaries, worksheet
relationship resolution, feature and cache scanning, preservation verdicts, and
workflow failure decisions. Network downloads and LibreOffice are not needed for
these tests. To run the same focused checks locally:

```sh
.venv/bin/python -m unittest discover -s . -p 'test_*.py' -v
uvx --from ruff==0.16.5 ruff check .
```

The two eight-file commands above are the end-to-end smoke proof. They require
LibreOffice and the Microsoft download and are intentionally not network tests
inside the product suite.
