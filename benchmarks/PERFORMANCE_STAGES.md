# Matched performance stages

**performance_stages.py** wraps the frozen performance_contract.py functions
without changing the workload, fixture definition, or harness hash. Its separate
controller hash identifies a fresh-process protocol: each discarded warmup and
each measured trial uses a new subprocess with the variant's exact interpreter,
imported Python source and native library.

List variants in priority order: one baseline, zero or more intermediate stage
entries, then one final. An optional control entry measures openpyxl once per
case/round rather than rerunning it for every intermediate stage. The control_cases
field may restrict control measurements to selected workloads.

Write a JSON file with this structure, using actual exact Git SHAs:

~~~json
{
  "variants": [
    {
      "label": "baseline",
      "role": "baseline",
      "python": "/absolute/baseline-env/bin/python",
      "source_root": "/absolute/baseline-source",
      "source_ref": "93e48b23605c798fcb98b2e4a86445a0b9c6679e",
      "build_metadata": "/absolute/baseline-build-metadata.json"
    },
    {
      "label": "final",
      "role": "final",
      "python": "/absolute/final-env/bin/python",
      "source_root": "/absolute/final-source",
      "source_ref": "0000000000000000000000000000000000000000",
      "pythonpath": "/absolute/final-source/python"
    },
    {
      "label": "openpyxl",
      "role": "control",
      "python": "/absolute/baseline-env/bin/python",
      "source_root": "/absolute/baseline-source",
      "source_ref": "93e48b23605c798fcb98b2e4a86445a0b9c6679e"
    }
  ],
  "control_cases": ["edit_top", "edit_bottom", "styled_small_eager"]
}
~~~

Replace the all-zero final SHA with its actual commit. Intermediate entries use
role "stage" and the same fields. Paths are interpreted relative to the JSON file.
Interpreter symlinks are preserved so virtual environments work correctly.
The optional pythonpath field may be a string or list of paths. Inherited
PYTHONPATH is cleared for each trial; only the variant's explicit override applies.
The declared source HEAD must match source_ref and be clean. The optional
build_metadata field accepts a JSON-file path or an inline JSON object.
Its contents and hash are recorded; missing metadata is explicitly labelled.

Run the controller using an environment containing the same openpyxl version as
the measured variants. Standard workload options go to the frozen parser:

~~~bash
/absolute/baseline-env/bin/python benchmarks/performance_stages.py \
  --variants /absolute/variants.json --seed 20261008 \
  --fixture-dir /tmp/wolfxl-performance-fixtures-v2 \
  --output /tmp/wolfxl-stages.json --label optimization-stack \
  --cases edit_top,edit_middle,edit_bottom,edit_merged,styled_small_eager,styled_large_read_only,styled_merged_eager \
  --rounds 5 --warmups 1
~~~

The controller serially shuffles active stage/control order with a random
generator seeded by the string "{seed}:{case_index}:{round_index}". Every order
is stored in the receipt. Fresh-process warmups prime filesystem caches; module
and process caches are fresh in every measured trial. This process protocol
differs from older in-process measurements and must be labelled separately in
published evidence. Keep the machine quiet throughout the paired run.

Every edit trial records the frozen copy/load/assignment/save/close/operation/
bounded-verification/total phases. The same bounded verifier runs for every edit
trial. The final measured trial for each variant/case performs one full eager
openpyxl value/style/merge check and one ZIP/XML/untouched-part check, outside the
timed operation and total. The independently edited reference is prepared once
per case. Styled reads match the independent openpyxl signature on every trial;
their final trial validates the original input package, because reads emit no
modified package. These scopes are explicit in the deep_validation field.

The single output JSON contains:

| Field | Meaning |
|---|---|
| contract, controller_sha256 | Frozen workload identity and controller source identity |
| workload_config, fixtures | Sizes, cases, rounds/warmups and stable fixture hashes |
| identities | Per-variant clean HEAD, import paths, Python/native/harness/environment/build hashes |
| scheduling.orders | Every warmup/measured stage order and reproducibility seed |
| results | Per-variant/case raw phase samples, ranges, medians, warmups, PIDs and one deep validation |
| summary | Direct baseline/final cumulative ratios, adjacent-stage ratios, optional final/openpyxl ratios |
| status | In progress, complete or failed; checkpoints retain completed samples and any error |

The controller rejects changed identities mid-run and different Python/platform/
openpyxl/harness identities between variants. Native libraries and WolfXL Python
source should differ between optimization stages. Summary ratios divide matched
medians directly; isolated improvements are never multiplied. There is no aggregate
speedup mixing different workloads. Do not publish incomplete or failed receipts.
