# community-final-v2-core-fresh-process: matched performance evidence

Contract: wolfxl-styled-edit-v2; 5 measured trials and 1 discarded warmup(s) per variant/case.
Every trial uses a fresh process. Stage/control order is reproducibly shuffled and executed serially. Times below are seconds.
Cumulative ratios divide matched baseline/final medians directly; isolated optimization gains are not multiplied.

Receipt SHA256: c49ce5f0ca3da44835657c39594169d3100715b91d630568b387e7bd52806701
Controller SHA256: c400130f81234d927d0987c5a7005a1331a07bfaf53ec635b82267b98d6b9537
Frozen harness SHA256: fadd85865a17424cc5642e2177fdffbcd1679f37715435dafb00e23006a72922

## Operation and total

| Workload | Baseline operation median [min,max] | Final operation median [min,max] | Cumulative engine gain | Final vs openpyxl operation | Cumulative total gain |
| --- | --- | --- | --- | --- | --- |
| edit_top | 0.471756 [0.467557, 0.506411] | 0.329874 [0.303201, 0.342099] | 1.430× | 46.281× | 1.419× |
| edit_middle | 0.836474 [0.757934, 0.882086] | 0.383352 [0.367436, 0.471613] | 2.182× | 38.353× | 1.182× |
| edit_bottom | 1.125115 [1.003846, 1.325068] | 0.481225 [0.459042, 0.493460] | 2.338× | 31.614× | 1.115× |
| edit_merged | 5.587626 [5.285182, 5.714922] | 0.635027 [0.617142, 0.648470] | 8.799× | Not measured | 8.742× |
| styled_small_eager | 0.052196 [0.047507, 0.061998] | 0.046342 [0.044374, 0.062849] | 1.126× | 2.250× | 1.126× |
| styled_large_read_only | 0.844555 [0.839360, 0.970056] | 0.336914 [0.328464, 0.348102] | 2.507× | 2.250× | 2.507× |
| styled_merged_eager | 0.952887 [0.889487, 0.975571] | 0.772520 [0.755913, 0.873404] | 1.233× | 1.390× | 1.233× |

Operation = load + assignment + save + close for edits; load + identical public Cell/style loop + close for reads.
Edit total adds fixture copy and the same independent bounded openpyxl verification. Read total equals operation.
Full semantic/package validation is outside operation and total and is reported below.

**High variance:** operation range exceeds 20% of its median for edit_middle/final, edit_bottom/baseline, styled_small_eager/baseline, styled_small_eager/final.
These five-trial ranges are descriptive, not confidence intervals.

## Phase attribution and raw ranges

### edit_top

| Variant | Phase | Median [min,max] seconds |
| --- | --- | --- |
| baseline | copy | 0.001680 [0.001533, 0.001879] |
| baseline | load | 0.111262 [0.108211, 0.122249] |
| baseline | assignment | 0.000034 [0.000034, 0.000035] |
| baseline | save | 0.360729 [0.355936, 0.384082] |
| baseline | close | 0.000031 [0.000029, 0.000033] |
| baseline | bounded_verification | 0.005296 [0.004777, 0.009739] |
| baseline | operation | 0.471756 [0.467557, 0.506411] |
| baseline | total | 0.478592 [0.474183, 0.518022] |
| final | copy | 0.001775 [0.001603, 0.002036] |
| final | load | 0.007061 [0.006256, 0.007319] |
| final | assignment | 0.025001 [0.024664, 0.026057] |
| final | save | 0.296444 [0.271438, 0.309788] |
| final | close | 0.000030 [0.000028, 0.000040] |
| final | bounded_verification | 0.005085 [0.004627, 0.005806] |
| final | operation | 0.329874 [0.303201, 0.342099] |
| final | total | 0.337299 [0.309869, 0.348597] |
| openpyxl | copy | 0.001718 [0.001586, 0.001762] |
| openpyxl | load | 8.008293 [7.731511, 8.886613] |
| openpyxl | assignment | 0.000078 [0.000075, 0.000081] |
| openpyxl | save | 7.137840 [6.832164, 7.258488] |
| openpyxl | close | 0.000001 [0.000001, 0.000002] |
| openpyxl | bounded_verification | 0.005803 [0.004938, 0.008654] |
| openpyxl | operation | 15.266873 [14.869444, 15.718865] |
| openpyxl | total | 15.274447 [14.876049, 15.726420] |

### edit_middle

| Variant | Phase | Median [min,max] seconds |
| --- | --- | --- |
| baseline | copy | 0.001726 [0.001695, 0.002136] |
| baseline | load | 0.124209 [0.111220, 0.130018] |
| baseline | assignment | 0.000035 [0.000033, 0.000077] |
| baseline | save | 0.706345 [0.646636, 0.757448] |
| baseline | close | 0.000031 [0.000029, 0.000032] |
| baseline | bounded_verification | 3.112166 [3.020431, 3.508528] |
| baseline | operation | 0.836474 [0.757934, 0.882086] |
| baseline | total | 3.951323 [3.780100, 4.392435] |
| final | copy | 0.001893 [0.001790, 0.003822] |
| final | load | 0.008390 [0.006407, 0.015506] |
| final | assignment | 0.024561 [0.023656, 0.051045] |
| final | save | 0.352424 [0.337043, 0.405024] |
| final | close | 0.000029 [0.000025, 0.000107] |
| final | bounded_verification | 2.972728 [2.855845, 3.162623] |
| final | operation | 0.383352 [0.367436, 0.471613] |
| final | total | 3.341976 [3.232776, 3.632249] |
| openpyxl | copy | 0.001650 [0.001556, 0.004309] |
| openpyxl | load | 7.657224 [7.434111, 7.996858] |
| openpyxl | assignment | 0.000077 [0.000073, 0.000085] |
| openpyxl | save | 6.966597 [6.681630, 7.268504] |
| openpyxl | close | 0.000002 [0.000001, 0.000002] |
| openpyxl | bounded_verification | 3.008112 [2.974308, 3.321149] |
| openpyxl | operation | 14.702704 [14.338946, 15.227611] |
| openpyxl | total | 17.678575 [17.349126, 18.295457] |

### edit_bottom

| Variant | Phase | Median [min,max] seconds |
| --- | --- | --- |
| baseline | copy | 0.001763 [0.001634, 0.002191] |
| baseline | load | 0.117069 [0.108623, 0.139448] |
| baseline | assignment | 0.000034 [0.000033, 0.000046] |
| baseline | save | 1.016114 [0.895145, 1.185526] |
| baseline | close | 0.000031 [0.000029, 0.000033] |
| baseline | bounded_verification | 6.017844 [5.502281, 7.008174] |
| baseline | operation | 1.125115 [1.003846, 1.325068] |
| baseline | total | 7.116279 [6.507811, 8.193293] |
| final | copy | 0.001746 [0.001558, 0.001893] |
| final | load | 0.006830 [0.006552, 0.008027] |
| final | assignment | 0.024114 [0.023839, 0.025954] |
| final | save | 0.450152 [0.428611, 0.461152] |
| final | close | 0.000027 [0.000025, 0.000030] |
| final | bounded_verification | 5.890953 [5.795881, 6.468069] |
| final | operation | 0.481225 [0.459042, 0.493460] |
| final | total | 6.382618 [6.276626, 6.951046] |
| openpyxl | copy | 0.001780 [0.001747, 0.002556] |
| openpyxl | load | 8.074691 [7.704038, 8.470309] |
| openpyxl | assignment | 0.000078 [0.000074, 0.000081] |
| openpyxl | save | 7.169886 [6.706908, 7.720586] |
| openpyxl | close | 0.000001 [0.000001, 0.000002] |
| openpyxl | bounded_verification | 5.965707 [5.901941, 6.343567] |
| openpyxl | operation | 15.213320 [14.411032, 16.190988] |
| openpyxl | total | 21.180782 [20.501835, 22.536330] |

### edit_merged

| Variant | Phase | Median [min,max] seconds |
| --- | --- | --- |
| baseline | copy | 0.001651 [0.001378, 0.002560] |
| baseline | load | 3.892522 [3.683399, 4.003248] |
| baseline | assignment | 1.304754 [1.186904, 1.336508] |
| baseline | save | 0.376523 [0.358704, 0.383187] |
| baseline | close | 0.030374 [0.026826, 0.031670] |
| baseline | bounded_verification | 0.018834 [0.017589, 0.022861] |
| baseline | operation | 5.587626 [5.285182, 5.714922] |
| baseline | total | 5.609206 [5.304160, 5.735417] |
| final | copy | 0.001450 [0.001372, 0.001704] |
| final | load | 0.007125 [0.006694, 0.009076] |
| final | assignment | 0.336934 [0.329251, 0.345530] |
| final | save | 0.296203 [0.278000, 0.301015] |
| final | close | 0.000032 [0.000031, 0.000034] |
| final | bounded_verification | 0.005307 [0.005039, 0.006421] |
| final | operation | 0.635027 [0.617142, 0.648470] |
| final | total | 0.641614 [0.624159, 0.656446] |

### styled_small_eager

| Variant | Phase | Median [min,max] seconds |
| --- | --- | --- |
| baseline | load | 0.008863 [0.008141, 0.010267] |
| baseline | close | 0.000891 [0.000661, 0.000996] |
| baseline | iterate_styled_cells | 0.042045 [0.038593, 0.050804] |
| baseline | operation | 0.052196 [0.047507, 0.061998] |
| baseline | total | 0.052196 [0.047507, 0.061998] |
| final | load | 0.006965 [0.006094, 0.008596] |
| final | close | 0.000707 [0.000692, 0.001041] |
| final | iterate_styled_cells | 0.038661 [0.037441, 0.053173] |
| final | operation | 0.046342 [0.044374, 0.062849] |
| final | total | 0.046342 [0.044374, 0.062849] |
| openpyxl | load | 0.075268 [0.070246, 0.083008] |
| openpyxl | close | 0.000003 [0.000002, 0.000005] |
| openpyxl | iterate_styled_cells | 0.027840 [0.026456, 0.029683] |
| openpyxl | operation | 0.104256 [0.096713, 0.110860] |
| openpyxl | total | 0.104256 [0.096713, 0.110860] |

### styled_large_read_only

| Variant | Phase | Median [min,max] seconds |
| --- | --- | --- |
| baseline | load | 0.007029 [0.006216, 0.008654] |
| baseline | close | 0.004855 [0.004563, 0.005008] |
| baseline | iterate_styled_cells | 0.833734 [0.827166, 0.958126] |
| baseline | operation | 0.844555 [0.839360, 0.970056] |
| baseline | total | 0.844555 [0.839360, 0.970056] |
| final | load | 0.006663 [0.005917, 0.007573] |
| final | close | 0.000046 [0.000041, 0.000079] |
| final | iterate_styled_cells | 0.330194 [0.321745, 0.342134] |
| final | operation | 0.336914 [0.328464, 0.348102] |
| final | total | 0.336914 [0.328464, 0.348102] |
| openpyxl | load | 0.003652 [0.003490, 0.004088] |
| openpyxl | close | 0.000025 [0.000024, 0.000028] |
| openpyxl | iterate_styled_cells | 0.753856 [0.720769, 0.792484] |
| openpyxl | operation | 0.757980 [0.724596, 0.796028] |
| openpyxl | total | 0.757980 [0.724596, 0.796028] |

### styled_merged_eager

| Variant | Phase | Median [min,max] seconds |
| --- | --- | --- |
| baseline | load | 0.360906 [0.333023, 0.376164] |
| baseline | close | 0.004385 [0.004189, 0.004845] |
| baseline | iterate_styled_cells | 0.587205 [0.550861, 0.615145] |
| baseline | operation | 0.952887 [0.889487, 0.975571] |
| baseline | total | 0.952887 [0.889487, 0.975571] |
| final | load | 0.006550 [0.006258, 0.006808] |
| final | close | 0.005540 [0.005204, 0.006903] |
| final | iterate_styled_cells | 0.760451 [0.743587, 0.859678] |
| final | operation | 0.772520 [0.755913, 0.873404] |
| final | total | 0.772520 [0.755913, 0.873404] |
| openpyxl | load | 0.728127 [0.705177, 0.847673] |
| openpyxl | close | 0.000003 [0.000002, 0.000004] |
| openpyxl | iterate_styled_cells | 0.345376 [0.305572, 0.400628] |
| openpyxl | operation | 1.073518 [1.018868, 1.231134] |
| openpyxl | total | 1.073518 [1.018868, 1.231134] |

## Independent validation

| Workload | Variant | Result | Scope | Full reopen seconds | Package check seconds |
| --- | --- | --- | --- | --- | --- |
| edit_top | baseline | pass | full output values/types/styles/merges, ZIP/XML, untouched-part bytes | 18.574360 | 5.587243 |
| edit_top | final | pass | full output values/types/styles/merges, ZIP/XML, untouched-part bytes | 16.696619 | 4.856130 |
| edit_top | openpyxl | pass | full output values/types/styles/merges, ZIP/XML, package rewrite reported | 18.630924 | 5.037557 |
| edit_middle | baseline | pass | full output values/types/styles/merges, ZIP/XML, untouched-part bytes | 16.202144 | 4.746464 |
| edit_middle | final | pass | full output values/types/styles/merges, ZIP/XML, untouched-part bytes | 16.193264 | 4.915634 |
| edit_middle | openpyxl | pass | full output values/types/styles/merges, ZIP/XML, package rewrite reported | 16.794833 | 5.394748 |
| edit_bottom | baseline | pass | full output values/types/styles/merges, ZIP/XML, untouched-part bytes | 16.273136 | 5.200776 |
| edit_bottom | final | pass | full output values/types/styles/merges, ZIP/XML, untouched-part bytes | 16.359842 | 5.033270 |
| edit_bottom | openpyxl | pass | full output values/types/styles/merges, ZIP/XML, package rewrite reported | 17.752207 | 5.037727 |
| edit_merged | baseline | pass | full output values/types/styles/merges, ZIP/XML, untouched-part bytes | 16.489566 | 5.043686 |
| edit_merged | final | pass | full output values/types/styles/merges, ZIP/XML, untouched-part bytes | 17.052857 | 4.841894 |
| styled_small_eager | baseline | pass | independent-styled-signature-and-input-integrity | 0.000000 | 0.027647 |
| styled_small_eager | final | pass | independent-styled-signature-and-input-integrity | 0.000000 | 0.024353 |
| styled_small_eager | openpyxl | pass | independent-styled-signature-and-input-integrity | 0.000000 | 0.024209 |
| styled_large_read_only | baseline | pass | independent-styled-signature-and-input-integrity | 0.000000 | 0.256187 |
| styled_large_read_only | final | pass | independent-styled-signature-and-input-integrity | 0.000000 | 0.251116 |
| styled_large_read_only | openpyxl | pass | independent-styled-signature-and-input-integrity | 0.000000 | 0.245773 |
| styled_merged_eager | baseline | pass | independent-styled-signature-and-input-integrity | 0.000000 | 0.242376 |
| styled_merged_eager | final | pass | independent-styled-signature-and-input-integrity | 0.000000 | 0.273730 |
| styled_merged_eager | openpyxl | pass | independent-styled-signature-and-input-integrity | 0.000000 | 0.254289 |

Every edit trial verifies both edited values through one bounded values iterator per edited sheet.
Full edit validation independently reopens the last output with eager openpyxl and checks all stored values, data types, style meanings, sheet bounds and merges.
ZIP CRC/XML checks and WolfXL untouched-part byte checks run once per variant/case. Openpyxl's full rewrite is reported rather than required to retain package bytes.
Styled reads match an independent openpyxl signature on every trial; their final check validates the original input package. It does not certify every reader feature.
These synthetic headless checks are not Microsoft Excel certification.

## Provenance

### baseline

- Measured engine: wolfxl
- Clean source HEAD: 93e48b23605c798fcb98b2e4a86445a0b9c6679e
- WolfXL version: 2.0.9
- Imported Python source SHA256: fa13b727a30551bfab994f48325b287a4c206215f277a9807dfc3fac02dbfa9d
- Native SHA256: 0af60f1ed0becebb0b7f3033a2ecaf219026284c132d518e0f36bd09dc989baa
- Python executable: /workspace/scratch/e6acfd2bd33d/baseline-envs/oss/bin/python
- Python: 3.12.14 (main, Aug 25 2026, 14:00:49) [Clang 22.1.3 ]
- Platform: Linux-6.18.44-x86_64-with-glibc2.39; machine: x86_64; CPUs: 9
- openpyxl: 3.1.5; source SHA256: fd9dcd8b157ca860b9076903c8038d7fbb83d38ff6a6e9b615f542b6e535ab56
- Build metadata SHA256: 323548c1c599405a91b1800a1b5ee43268646a310dacce9a05020e49adc1f2b7
- Rust compiler: rustc 1.90.0 (1159e78c4 2025-09-14)
- Build target/features: /workspace/scratch/e6acfd2bd33d/build-targets/oss-baseline / ['extension-module']

### final

- Measured engine: wolfxl
- Clean source HEAD: 02884309b62de4f98b92c837c4c9184fe64ed05c
- WolfXL version: 2.0.9
- Imported Python source SHA256: 229de3400f62024f5f07b34c1f1d9823ea22f93a4f1de185a922cac89980d885
- Native SHA256: 7ea062b1c794044b6cd1d894c780c14decbe3a5c66177bc6874b9ecdefc752dc
- Python executable: /workspace/scratch/e6acfd2bd33d/variant-envs/oss-final-v2/bin/python
- Python: 3.12.14 (main, Aug 25 2026, 14:00:49) [Clang 22.1.3 ]
- Platform: Linux-6.18.44-x86_64-with-glibc2.39; machine: x86_64; CPUs: 9
- openpyxl: 3.1.5; source SHA256: fd9dcd8b157ca860b9076903c8038d7fbb83d38ff6a6e9b615f542b6e535ab56
- Build metadata SHA256: 186cb3dbbb58aff988588db348905ddabac2e25d9d18cb8787f685c271810b0e
- Rust compiler: rustc 1.90.0 (1159e78c4 2025-09-14)
- Build target/features: /workspace/scratch/e6acfd2bd33d/build-targets/oss-final-v2 / ['extension-module']

### openpyxl

- Measured engine: openpyxl
- Clean source HEAD: 93e48b23605c798fcb98b2e4a86445a0b9c6679e
- WolfXL version: 2.0.9
- Imported Python source SHA256: fa13b727a30551bfab994f48325b287a4c206215f277a9807dfc3fac02dbfa9d
- Native SHA256: 0af60f1ed0becebb0b7f3033a2ecaf219026284c132d518e0f36bd09dc989baa
- Python executable: /workspace/scratch/e6acfd2bd33d/baseline-envs/oss/bin/python
- Python: 3.12.14 (main, Aug 25 2026, 14:00:49) [Clang 22.1.3 ]
- Platform: Linux-6.18.44-x86_64-with-glibc2.39; machine: x86_64; CPUs: 9
- openpyxl: 3.1.5; source SHA256: fd9dcd8b157ca860b9076903c8038d7fbb83d38ff6a6e9b615f542b6e535ab56
- Build metadata SHA256: 323548c1c599405a91b1800a1b5ee43268646a310dacce9a05020e49adc1f2b7
- Rust compiler: rustc 1.90.0 (1159e78c4 2025-09-14)
- Build target/features: /workspace/scratch/e6acfd2bd33d/build-targets/oss-baseline / ['extension-module']

| Workload | Rows/sheet | Populated value cells | Max column | Fixture SHA256 |
| --- | --- | --- | --- | --- |
| edit_bottom | 200000 | 1000000 | 5 | 0e41152df64d917eedda420f6ccb51f98323d5e3a49a88aa385e26aa73930341 |
| edit_merged | 200000 | 999997 | 5 | 229e85173cd77ffd95d68c931c4f93967036fa0883382c48fb081dbfc6cb9875 |
| edit_middle | 200000 | 1000000 | 5 | 0e41152df64d917eedda420f6ccb51f98323d5e3a49a88aa385e26aa73930341 |
| edit_top | 200000 | 1000000 | 5 | 0e41152df64d917eedda420f6ccb51f98323d5e3a49a88aa385e26aa73930341 |
| styled_large_read_only | 20000 | 100000 | 5 | 766f1a5c1f97378588d586a54665b7bd3a5ab0c0378e1f4ddaeb74fedeb88110 |
| styled_merged_eager | 20000 | 99997 | 5 | edd43d2a9fe792e53f9872c9e9016dcd43ff750b44fb81152ccd54ee3bc98e72 |
| styled_small_eager | 2000 | 10000 | 5 | e35622d88f991eae2ce09063c0ab39a522a52576cf37e1c47cfe3e0e19c70a79 |

## Claim boundaries

- This fresh-process v2 protocol differs from the historical in-process Apple benchmark and its eager verification totals.
- The numeric results apply to these fixed synthetic fixtures, engine versions, source/native builds and this machine.
- Plain rows contain five populated columns even though the historic nominal column setting was eight.
- Bounded late-row verification still streams the XML prefix; it is not constant-time random access.
- Engine gains and verification changes are separate. No overall gain is synthesized across unlike workloads.
- Raw trial phases, warmups, PIDs, order, identity hashes and validation results remain in the JSON receipt.
