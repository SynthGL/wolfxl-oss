# community-final-guards-fresh-process-v2: matched performance evidence

Contract: wolfxl-styled-edit-v2; 5 measured trials and 1 discarded warmup(s) per variant/case.
Every trial uses a fresh process. Stage/control order is reproducibly shuffled and executed serially. Times below are seconds.
Cumulative ratios divide matched baseline/final medians directly; isolated optimization gains are not multiplied.

Receipt SHA256: 805d319acc29b7aa00ca9efc5be379d46765575e7fd8cb283dfb1b5c59cd44e2
Controller SHA256: c400130f81234d927d0987c5a7005a1331a07bfaf53ec635b82267b98d6b9537
Frozen harness SHA256: fadd85865a17424cc5642e2177fdffbcd1679f37715435dafb00e23006a72922

## Operation and total

| Workload | Baseline operation median [min,max] | Final operation median [min,max] | Cumulative engine gain | Final vs openpyxl operation | Cumulative total gain |
| --- | --- | --- | --- | --- | --- |
| styled_large_eager | 0.876615 [0.827768, 0.947905] | 0.614065 [0.540804, 0.628436] | 1.428× | 2.270× | 1.428× |
| styled_sparse_eager | 3.403551 [3.340596, 3.713602] | 1.815314 [1.626295, 1.953815] | 1.875× | 1.745× | 1.875× |
| styled_high_cardinality_eager | 0.948151 [0.912700, 1.077674] | 0.681685 [0.662656, 0.725048] | 1.391× | 2.073× | 1.391× |
| styled_sparse_read_only | 4.295201 [4.085524, 4.611708] | 1.995125 [1.929984, 2.011292] | 2.153× | 0.084× | 2.153× |
| styled_high_cardinality_read_only | 1.330946 [1.306431, 1.555266] | 0.609583 [0.562129, 0.630725] | 2.183× | 1.788× | 2.183× |

Operation = load + assignment + save + close for edits; load + identical public Cell/style loop + close for reads.
Edit total adds fixture copy and the same independent bounded openpyxl verification. Read total equals operation.
Full semantic/package validation is outside operation and total and is reported below.

**High variance:** operation range exceeds 20% of its median for styled_high_cardinality_read_only/openpyxl.
These five-trial ranges are descriptive, not confidence intervals.

## Phase attribution and raw ranges

### styled_large_eager

| Variant | Phase | Median [min,max] seconds |
| --- | --- | --- |
| baseline | load | 0.022032 [0.021413, 0.024294] |
| baseline | close | 0.006029 [0.005720, 0.006319] |
| baseline | iterate_styled_cells | 0.848556 [0.800620, 0.917334] |
| baseline | operation | 0.876615 [0.827768, 0.947905] |
| baseline | total | 0.876615 [0.827768, 0.947905] |
| final | load | 0.006921 [0.006328, 0.008550] |
| final | close | 0.006902 [0.005978, 0.008362] |
| final | iterate_styled_cells | 0.600673 [0.526580, 0.612968] |
| final | operation | 0.614065 [0.540804, 0.628436] |
| final | total | 0.614065 [0.540804, 0.628436] |
| openpyxl | load | 0.994108 [0.938893, 1.030635] |
| openpyxl | close | 0.000003 [0.000003, 0.000003] |
| openpyxl | iterate_styled_cells | 0.438498 [0.395955, 0.463328] |
| openpyxl | operation | 1.393875 [1.378478, 1.457451] |
| openpyxl | total | 1.393875 [1.378478, 1.457451] |

### styled_sparse_eager

| Variant | Phase | Median [min,max] seconds |
| --- | --- | --- |
| baseline | load | 0.008430 [0.007204, 0.009164] |
| baseline | close | 0.000729 [0.000566, 0.001041] |
| baseline | iterate_styled_cells | 3.393806 [3.331113, 3.704529] |
| baseline | operation | 3.403551 [3.340596, 3.713602] |
| baseline | total | 3.403551 [3.340596, 3.713602] |
| final | load | 0.007876 [0.006426, 0.012908] |
| final | close | 0.000653 [0.000598, 0.000881] |
| final | iterate_styled_cells | 1.801638 [1.619210, 1.944976] |
| final | operation | 1.815314 [1.626295, 1.953815] |
| final | total | 1.815314 [1.626295, 1.953815] |
| openpyxl | load | 0.058502 [0.056162, 0.070470] |
| openpyxl | close | 0.000003 [0.000003, 0.000003] |
| openpyxl | iterate_styled_cells | 3.109195 [2.955397, 3.250610] |
| openpyxl | operation | 3.167710 [3.011573, 3.321094] |
| openpyxl | total | 3.167710 [3.011573, 3.321094] |

### styled_high_cardinality_eager

| Variant | Phase | Median [min,max] seconds |
| --- | --- | --- |
| baseline | load | 0.074216 [0.073534, 0.091447] |
| baseline | close | 0.006000 [0.005344, 0.008223] |
| baseline | iterate_styled_cells | 0.867916 [0.830382, 0.994128] |
| baseline | operation | 0.948151 [0.912700, 1.077674] |
| baseline | total | 0.948151 [0.912700, 1.077674] |
| final | load | 0.059723 [0.056944, 0.061594] |
| final | close | 0.007245 [0.006982, 0.007639] |
| final | iterate_styled_cells | 0.612432 [0.595293, 0.660641] |
| final | operation | 0.681685 [0.662656, 0.725048] |
| final | total | 0.681685 [0.662656, 0.725048] |
| openpyxl | load | 0.995647 [0.976506, 1.027193] |
| openpyxl | close | 0.000003 [0.000002, 0.000004] |
| openpyxl | iterate_styled_cells | 0.411485 [0.399669, 0.440091] |
| openpyxl | operation | 1.413241 [1.393602, 1.438697] |
| openpyxl | total | 1.413241 [1.393602, 1.438697] |

### styled_sparse_read_only

| Variant | Phase | Median [min,max] seconds |
| --- | --- | --- |
| baseline | load | 0.006548 [0.006209, 0.008851] |
| baseline | close | 0.000802 [0.000649, 0.001038] |
| baseline | iterate_styled_cells | 4.287399 [4.078264, 4.604560] |
| baseline | operation | 4.295201 [4.085524, 4.611708] |
| baseline | total | 4.295201 [4.085524, 4.611708] |
| final | load | 0.006723 [0.006367, 0.007779] |
| final | close | 0.000041 [0.000038, 0.000049] |
| final | iterate_styled_cells | 1.987551 [1.922125, 2.004865] |
| final | operation | 1.995125 [1.929984, 2.011292] |
| final | total | 1.995125 [1.929984, 2.011292] |
| openpyxl | load | 0.003449 [0.003348, 0.004218] |
| openpyxl | close | 0.000029 [0.000027, 0.000116] |
| openpyxl | iterate_styled_cells | 0.164163 [0.155343, 0.178500] |
| openpyxl | operation | 0.167566 [0.158917, 0.182438] |
| openpyxl | total | 0.167566 [0.158917, 0.182438] |

### styled_high_cardinality_read_only

| Variant | Phase | Median [min,max] seconds |
| --- | --- | --- |
| baseline | load | 0.058767 [0.056827, 0.085499] |
| baseline | close | 0.005781 [0.005560, 0.007258] |
| baseline | iterate_styled_cells | 1.268119 [1.236983, 1.462488] |
| baseline | operation | 1.330946 [1.306431, 1.555266] |
| baseline | total | 1.330946 [1.306431, 1.555266] |
| final | load | 0.065366 [0.060130, 0.077500] |
| final | close | 0.000400 [0.000361, 0.000580] |
| final | iterate_styled_cells | 0.540952 [0.499222, 0.564947] |
| final | operation | 0.609583 [0.562129, 0.630725] |
| final | total | 0.609583 [0.562129, 0.630725] |
| openpyxl | load | 0.047651 [0.045869, 0.054567] |
| openpyxl | close | 0.000031 [0.000026, 0.000054] |
| openpyxl | iterate_styled_cells | 1.043666 [0.938983, 1.176589] |
| openpyxl | operation | 1.089768 [0.984893, 1.224279] |
| openpyxl | total | 1.089768 [0.984893, 1.224279] |

## Independent validation

| Workload | Variant | Result | Scope | Full reopen seconds | Package check seconds |
| --- | --- | --- | --- | --- | --- |
| styled_large_eager | baseline | pass | independent-styled-signature-and-input-integrity | 0.000000 | 0.333978 |
| styled_large_eager | final | pass | independent-styled-signature-and-input-integrity | 0.000000 | 0.322093 |
| styled_large_eager | openpyxl | pass | independent-styled-signature-and-input-integrity | 0.000000 | 0.337393 |
| styled_sparse_eager | baseline | pass | independent-styled-signature-and-input-integrity | 0.000000 | 0.020179 |
| styled_sparse_eager | final | pass | independent-styled-signature-and-input-integrity | 0.000000 | 0.023449 |
| styled_sparse_eager | openpyxl | pass | independent-styled-signature-and-input-integrity | 0.000000 | 0.024036 |
| styled_high_cardinality_eager | baseline | pass | independent-styled-signature-and-input-integrity | 0.000000 | 0.358708 |
| styled_high_cardinality_eager | final | pass | independent-styled-signature-and-input-integrity | 0.000000 | 0.380848 |
| styled_high_cardinality_eager | openpyxl | pass | independent-styled-signature-and-input-integrity | 0.000000 | 0.374096 |
| styled_sparse_read_only | baseline | pass | independent-styled-signature-and-input-integrity | 0.000000 | 0.022169 |
| styled_sparse_read_only | final | pass | independent-styled-signature-and-input-integrity | 0.000000 | 0.018682 |
| styled_sparse_read_only | openpyxl | pass | independent-styled-signature-and-input-integrity | 0.000000 | 0.018753 |
| styled_high_cardinality_read_only | baseline | pass | independent-styled-signature-and-input-integrity | 0.000000 | 0.370904 |
| styled_high_cardinality_read_only | final | pass | independent-styled-signature-and-input-integrity | 0.000000 | 0.387776 |
| styled_high_cardinality_read_only | openpyxl | pass | independent-styled-signature-and-input-integrity | 0.000000 | 0.343268 |

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
- Clean source HEAD: 19fdf4a947ab067ac17aa2bca1cd2663e6590639
- WolfXL version: 2.0.9
- Imported Python source SHA256: 790dc0f4f26aa808db95f4d3aac99d451f4eb56f83702dfaeed08fc5fa1cbd70
- Native SHA256: 83b642037d8e7668a93f118efc87832d968ada1e20160f31726f034822be6ed2
- Python executable: /workspace/scratch/e6acfd2bd33d/variant-envs/oss-final/bin/python
- Python: 3.12.14 (main, Aug 25 2026, 14:00:49) [Clang 22.1.3 ]
- Platform: Linux-6.18.44-x86_64-with-glibc2.39; machine: x86_64; CPUs: 9
- openpyxl: 3.1.5; source SHA256: fd9dcd8b157ca860b9076903c8038d7fbb83d38ff6a6e9b615f542b6e535ab56
- Build metadata SHA256: 034dcda201adeb1f2cb831317a841413ea2d20d99e33983db9ebf441ce70de0c
- Rust compiler: rustc 1.90.0 (1159e78c4 2025-09-14)
- Build target/features: /workspace/scratch/e6acfd2bd33d/build-targets/oss-final / ['extension-module']

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
| styled_high_cardinality_eager | 25000 | 125000 | 5 | 16f0fcf0af96a2156b8720ef4a22e3be1d5dc46838470422b5a65e80362abd85 |
| styled_high_cardinality_read_only | 25000 | 125000 | 5 | 16f0fcf0af96a2156b8720ef4a22e3be1d5dc46838470422b5a65e80362abd85 |
| styled_large_eager | 25000 | 125000 | 5 | a3b9df76f47a519ee05b11c1b6e322265f195fe6f3736cae6c95322bd4daf1d8 |
| styled_sparse_eager | 25000 | 7816 | 32 | 5a562d80a9cd2da0db518a52171af98f4ffadeab80297e0a8a9392850914dea5 |
| styled_sparse_read_only | 25000 | 7816 | 32 | 5a562d80a9cd2da0db518a52171af98f4ffadeab80297e0a8a9392850914dea5 |

## Claim boundaries

- This fresh-process v2 protocol differs from the historical in-process Apple benchmark and its eager verification totals.
- The numeric results apply to these fixed synthetic fixtures, engine versions, source/native builds and this machine.
- Plain rows contain five populated columns even though the historic nominal column setting was eight.
- Bounded late-row verification still streams the XML prefix; it is not constant-time random access.
- Engine gains and verification changes are separate. No overall gain is synthesized across unlike workloads.
- Raw trial phases, warmups, PIDs, order, identity hashes and validation results remain in the JSON receipt.
