# Community: independent styled-read semantic validation

This is an untimed validation supplement, not a replacement timing receipt or Microsoft Excel certification.

Every iterator coordinate is checked for value, exact Python value type, Excel data_type, font, fill, border, alignment and number format. Geometry and merge metadata are checked separately. Both engines use the same read_only and data_only=True mode. No read-only worksheet is converted to eager Cells.

Style meanings use component fields rather than style IDs. Schema-default booleans/numbers, implicit alignment defaults and absent read-only EmptyCell fills are normalized; active color selector/value/tint is retained. Returned style objects are canonicalized with a bounded identity cache. Full-cell digests bind component-meaning SHA256 hashes, so identical style dictionaries are not serialized again at every coordinate. Protection, conditional formatting, dates not present in these fixtures, calculation and rendering are outside this supplement.

Mismatch categories retain source-XML scope: stored_value and stored_blank are authored cells; implicit_blank is an un-authored iterator placeholder. The merged_placeholder suffix identifies a non-anchor coordinate in an authored merge. Coordinate/value/type checks apply to every scope. Style API differences for implicit blanks are reported as representation boundaries separately from authored-cell style differences; no missing authored style is normalized away.

Numeric value equality normalizes integral float/int values, while exact Python numeric subtype differences are reported separately, without treating them as silently equivalent API types.

| Case | Mode | Cells/variant | Baseline | Final | Classification |
|---|---|---:|---|---|---|
| styled_small_eager | eager | 10,000 | mismatch | mismatch | identical baseline boundary |
| styled_large_read_only | read-only | 100,000 | mismatch | mismatch | identical baseline boundary |
| styled_merged_eager | eager | 100,000 | mismatch | mismatch | baseline and final mismatches differ; inspect categories |
| styled_large_eager | eager | 125,000 | mismatch | mismatch | identical baseline boundary |
| styled_sparse_eager | eager | 800,000 | mismatch | mismatch | identical baseline boundary |
| styled_high_cardinality_eager | eager | 125,000 | mismatch | mismatch | identical baseline boundary |
| styled_sparse_read_only | read-only | 800,000 | mismatch | mismatch | identical baseline boundary |
| styled_high_cardinality_read_only | read-only | 125,000 | mismatch | mismatch | identical baseline boundary |

All mismatch counts and ordered mismatch digests are retained in JSON; each category includes its first twelve examples. A mismatch present on the original baseline is a baseline boundary, not evidence that the final changes introduced it. Changed baseline/final mismatch sets remain explicitly classified as changed rather than assumed fixed.

## Mismatch categories

- styled_small_eager/baseline: fill=1,000, font=8,000
  - fill: stored_value=1,000
  - font: stored_value=8,000
- styled_small_eager/final: fill=1,000, font=8,000
  - fill: stored_value=1,000
  - font: stored_value=8,000
- styled_large_read_only/baseline: fill=10,000, font=80,000
  - fill: stored_value=10,000
  - font: stored_value=80,000
- styled_large_read_only/final: fill=10,000, font=80,000
  - fill: stored_value=10,000
  - font: stored_value=80,000
- styled_merged_eager/baseline: border=4, fill=10,000, font=80,000
  - border: stored_blank/merged_placeholder=3, stored_value=1
  - fill: stored_value=10,000
  - font: stored_blank/merged_placeholder=3, stored_value=79,997
- styled_merged_eager/final: fill=10,000, font=80,000
  - fill: stored_value=10,000
  - font: stored_blank/merged_placeholder=3, stored_value=79,997
- styled_large_eager/baseline: fill=12,500, font=100,000
  - fill: stored_value=12,500
  - font: stored_value=100,000
- styled_large_eager/final: fill=12,500, font=100,000
  - fill: stored_value=12,500
  - font: stored_value=100,000
- styled_sparse_eager/baseline: font=798,437
  - font: implicit_blank=792,184, stored_value=6,253
- styled_sparse_eager/final: font=798,437
  - font: implicit_blank=792,184, stored_value=6,253
- styled_high_cardinality_eager/baseline: fill=12,500, font=100,000
  - fill: stored_value=12,500
  - font: stored_value=100,000
- styled_high_cardinality_eager/final: fill=12,500, font=100,000
  - fill: stored_value=12,500
  - font: stored_value=100,000
- styled_sparse_read_only/baseline: font=6,253
  - font: stored_value=6,253
- styled_sparse_read_only/final: font=6,253
  - font: stored_value=6,253
- styled_high_cardinality_read_only/baseline: fill=12,500, font=100,000
  - fill: stored_value=12,500
  - font: stored_value=100,000
- styled_high_cardinality_read_only/final: fill=12,500, font=100,000
  - fill: stored_value=12,500
  - font: stored_value=100,000

## Binding

- Producer SHA256: `f7a3c8a1bcc767160a1f90e285207fddb9673a06ce8caacc45f58056b90353e5`
- core.json receipt SHA256: `c49ce5f0ca3da44835657c39594169d3100715b91d630568b387e7bd52806701`
- guards.json receipt SHA256: `06cb40937240a19aff608fe15ccaed96a36df9ec8666a77073b3cd06a2d06657`
- baseline source: `93e48b23605c798fcb98b2e4a86445a0b9c6679e`
- baseline native SHA256: `0af60f1ed0becebb0b7f3033a2ecaf219026284c132d518e0f36bd09dc989baa`
- final source: `02884309b62de4f98b92c837c4c9184fe64ed05c`
- final native SHA256: `7ea062b1c794044b6cd1d894c780c14decbe3a5c66177bc6874b9ecdefc752dc`
- openpyxl source: `93e48b23605c798fcb98b2e4a86445a0b9c6679e`
- openpyxl native SHA256: `0af60f1ed0becebb0b7f3033a2ecaf219026284c132d518e0f36bd09dc989baa`

Installed package Python/native hashes were checked against the frozen receipts in every worker; source fixture hashes were checked before and after each scan. This report never modifies the timing receipts or the input workbooks.

## Observed scope

Eight cases compare 2,185,000 coordinates per variant: 4,370,000 original/final coordinate pairs in this edition. Every coordinate, value, exact Python value type, Excel data_type, geometry, merge, alignment and number format matched the same-mode independent reference.

There are zero final-only mismatch categories and zero changed mismatch sets among retained categories, checked from every ordered mismatch digest. Four merged-border color mismatches present in the baseline disappear in the final result. The three non-anchor merged coordinates are checked directly, without changing read mode.

Font API metadata boundaries remain on authored/stored cells in the affected cases, with identical baseline/final mismatch sets; source default-font and theme metadata are not fully reflected by the returned objects. Sparse eager also reports implicit-blank font differences as representation boundaries, separately from stored-value font gaps. Sparse read-only has no implicit-blank font mismatch against openpyxl EmptyCell under the documented normalization.

Community also retains authored solid-fill ARGB metadata differences. Recorded examples include openpyxl `00FFD966` versus WolfXL `FFFFD966`; these are preserved as API differences, with identical baseline/final mismatch digests. High-cardinality font cases also retain source-theme/family/scheme metadata differences.

These findings support absence of new mismatches in this fixture set, not complete style equivalence. Microsoft Excel rendering, appearance, calculation, and workbook features absent from these fixtures were not certified.

Raw JSON SHA256: `b91917a99ff381d663f8a99f79112024889c81107d638ce0eba40b8eb552b92c`

Scope renderer SHA256: `0630cf4d735baa30654955be654997498b0c126ee7f009628481707c195a30a7`
