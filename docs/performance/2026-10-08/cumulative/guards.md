# community-final-v2-guards-fresh-process: matched performance evidence

Contract: wolfxl-styled-edit-v2; 5 measured trials and 1 discarded warmup(s) per variant/case.
Every trial uses a fresh process. Stage/control order is reproducibly shuffled and executed serially. Times below are seconds.
Cumulative ratios divide matched baseline/final medians directly; isolated optimization gains are not multiplied.

Receipt SHA256: 06cb40937240a19aff608fe15ccaed96a36df9ec8666a77073b3cd06a2d06657
Controller SHA256: c400130f81234d927d0987c5a7005a1331a07bfaf53ec635b82267b98d6b9537
Frozen harness SHA256: fadd85865a17424cc5642e2177fdffbcd1679f37715435dafb00e23006a72922

## Operation and total

| Workload | Baseline operation median [min,max] | Final operation median [min,max] | Cumulative engine gain | Final vs openpyxl operation | Cumulative total gain |
| --- | --- | --- | --- | --- | --- |
| styled_large_eager | 0.797995 [0.771152, 0.889157] | 0.548327 [0.525427, 0.580760] | 1.455× | 2.662× | 1.455× |
| styled_sparse_eager | 3.435690 [3.212537, 3.588880] | 1.765592 [1.618842, 1.883346] | 1.946× | 1.838× | 1.946× |
| styled_high_cardinality_eager | 0.980464 [0.943893, 1.116953] | 0.666920 [0.651396, 0.694713] | 1.470× | 2.134× | 1.470× |
| styled_sparse_read_only | 4.339737 [4.089201, 4.645300] | 1.932072 [1.906666, 1.980900] | 2.246× | 0.083× | 2.246× |
| styled_high_cardinality_read_only | 1.295286 [1.248127, 1.367927] | 0.570462 [0.516474, 0.614585] | 2.271× | 1.851× | 2.271× |

Operation = load + assignment + save + close for edits; load + identical public Cell/style loop + close for reads.
Edit total adds fixture copy and the same independent bounded openpyxl verification. Read total equals operation.
Full semantic/package validation is outside operation and total and is reported below.

**High variance:** operation range exceeds 20% of its median for styled_sparse_read_only/openpyxl.
These five-trial ranges are descriptive, not confidence intervals.

## Phase attribution and raw ranges

### styled_large_eager

| Variant | Phase | Median [min,max] seconds |
| --- | --- | --- |
| baseline | load | 0.020307 [0.019711, 0.021324] |
| baseline | close | 0.005797 [0.005432, 0.008220] |
| baseline | iterate_styled_cells | 0.771959 [0.745359, 0.859597] |
| baseline | operation | 0.797995 [0.771152, 0.889157] |
| baseline | total | 0.797995 [0.771152, 0.889157] |
| final | load | 0.007864 [0.006819, 0.008663] |
| final | close | 0.006261 [0.005910, 0.007682] |
| final | iterate_styled_cells | 0.532764 [0.512684, 0.565823] |
| final | operation | 0.548327 [0.525427, 0.580760] |
| final | total | 0.548327 [0.525427, 0.580760] |
| openpyxl | load | 1.001229 [0.935789, 1.066314] |
| openpyxl | close | 0.000003 [0.000003, 0.000004] |
| openpyxl | iterate_styled_cells | 0.428379 [0.401604, 0.458601] |
| openpyxl | operation | 1.459848 [1.337409, 1.512113] |
| openpyxl | total | 1.459848 [1.337409, 1.512113] |

### styled_sparse_eager

| Variant | Phase | Median [min,max] seconds |
| --- | --- | --- |
| baseline | load | 0.008003 [0.007230, 0.010867] |
| baseline | close | 0.000657 [0.000594, 0.000761] |
| baseline | iterate_styled_cells | 3.427432 [3.204298, 3.577365] |
| baseline | operation | 3.435690 [3.212537, 3.588880] |
| baseline | total | 3.435690 [3.212537, 3.588880] |
| final | load | 0.006438 [0.006045, 0.007717] |
| final | close | 0.000649 [0.000564, 0.000722] |
| final | iterate_styled_cells | 1.757284 [1.611528, 1.876597] |
| final | operation | 1.765592 [1.618842, 1.883346] |
| final | total | 1.765592 [1.618842, 1.883346] |
| openpyxl | load | 0.060793 [0.058927, 0.075553] |
| openpyxl | close | 0.000003 [0.000002, 0.000003] |
| openpyxl | iterate_styled_cells | 3.185085 [2.960177, 3.489456] |
| openpyxl | operation | 3.245891 [3.027282, 3.549441] |
| openpyxl | total | 3.245891 [3.027282, 3.549441] |

### styled_high_cardinality_eager

| Variant | Phase | Median [min,max] seconds |
| --- | --- | --- |
| baseline | load | 0.075756 [0.072780, 0.078986] |
| baseline | close | 0.005669 [0.005451, 0.005998] |
| baseline | iterate_styled_cells | 0.902217 [0.861010, 1.036833] |
| baseline | operation | 0.980464 [0.943893, 1.116953] |
| baseline | total | 0.980464 [0.943893, 1.116953] |
| final | load | 0.059014 [0.057645, 0.061338] |
| final | close | 0.007376 [0.006948, 0.007767] |
| final | iterate_styled_cells | 0.597798 [0.586323, 0.628067] |
| final | operation | 0.666920 [0.651396, 0.694713] |
| final | total | 0.666920 [0.651396, 0.694713] |
| openpyxl | load | 1.004005 [0.966555, 1.100565] |
| openpyxl | close | 0.000003 [0.000002, 0.000003] |
| openpyxl | iterate_styled_cells | 0.432890 [0.401167, 0.450499] |
| openpyxl | operation | 1.423463 [1.399459, 1.547096] |
| openpyxl | total | 1.423463 [1.399459, 1.547096] |

### styled_sparse_read_only

| Variant | Phase | Median [min,max] seconds |
| --- | --- | --- |
| baseline | load | 0.006493 [0.006109, 0.008410] |
| baseline | close | 0.000765 [0.000704, 0.000850] |
| baseline | iterate_styled_cells | 4.332896 [4.081699, 4.636079] |
| baseline | operation | 4.339737 [4.089201, 4.645300] |
| baseline | total | 4.339737 [4.089201, 4.645300] |
| final | load | 0.006564 [0.006073, 0.007056] |
| final | close | 0.000041 [0.000039, 0.000055] |
| final | iterate_styled_cells | 1.924949 [1.900084, 1.974036] |
| final | operation | 1.932072 [1.906666, 1.980900] |
| final | total | 1.932072 [1.906666, 1.980900] |
| openpyxl | load | 0.003741 [0.003615, 0.004108] |
| openpyxl | close | 0.000030 [0.000024, 0.000084] |
| openpyxl | iterate_styled_cells | 0.156749 [0.155076, 0.227541] |
| openpyxl | operation | 0.160888 [0.158766, 0.231319] |
| openpyxl | total | 0.160888 [0.158766, 0.231319] |

### styled_high_cardinality_read_only

| Variant | Phase | Median [min,max] seconds |
| --- | --- | --- |
| baseline | load | 0.058140 [0.056169, 0.069953] |
| baseline | close | 0.006012 [0.005597, 0.007790] |
| baseline | iterate_styled_cells | 1.230682 [1.186202, 1.292361] |
| baseline | operation | 1.295286 [1.248127, 1.367927] |
| baseline | total | 1.295286 [1.248127, 1.367927] |
| final | load | 0.065300 [0.062973, 0.083845] |
| final | close | 0.000394 [0.000379, 0.000450] |
| final | iterate_styled_cells | 0.504766 [0.453110, 0.530279] |
| final | operation | 0.570462 [0.516474, 0.614585] |
| final | total | 0.570462 [0.516474, 0.614585] |
| openpyxl | load | 0.045946 [0.043613, 0.049733] |
| openpyxl | close | 0.000027 [0.000025, 0.000029] |
| openpyxl | iterate_styled_cells | 1.010486 [0.976259, 1.066986] |
| openpyxl | operation | 1.056134 [1.022242, 1.115053] |
| openpyxl | total | 1.056134 [1.022242, 1.115053] |

## Independent validation

| Workload | Variant | Result | Scope | Full reopen seconds | Package check seconds |
| --- | --- | --- | --- | --- | --- |
| styled_large_eager | baseline | pass | independent-styled-signature-and-input-integrity | 0.000000 | 0.303509 |
| styled_large_eager | final | pass | independent-styled-signature-and-input-integrity | 0.000000 | 0.303118 |
| styled_large_eager | openpyxl | pass | independent-styled-signature-and-input-integrity | 0.000000 | 0.330377 |
| styled_sparse_eager | baseline | pass | independent-styled-signature-and-input-integrity | 0.000000 | 0.019558 |
| styled_sparse_eager | final | pass | independent-styled-signature-and-input-integrity | 0.000000 | 0.018736 |
| styled_sparse_eager | openpyxl | pass | independent-styled-signature-and-input-integrity | 0.000000 | 0.020218 |
| styled_high_cardinality_eager | baseline | pass | independent-styled-signature-and-input-integrity | 0.000000 | 0.325750 |
| styled_high_cardinality_eager | final | pass | independent-styled-signature-and-input-integrity | 0.000000 | 0.342832 |
| styled_high_cardinality_eager | openpyxl | pass | independent-styled-signature-and-input-integrity | 0.000000 | 0.394217 |
| styled_sparse_read_only | baseline | pass | independent-styled-signature-and-input-integrity | 0.000000 | 0.019229 |
| styled_sparse_read_only | final | pass | independent-styled-signature-and-input-integrity | 0.000000 | 0.027404 |
| styled_sparse_read_only | openpyxl | pass | independent-styled-signature-and-input-integrity | 0.000000 | 0.025011 |
| styled_high_cardinality_read_only | baseline | pass | independent-styled-signature-and-input-integrity | 0.000000 | 0.379919 |
| styled_high_cardinality_read_only | final | pass | independent-styled-signature-and-input-integrity | 0.000000 | 0.347203 |
| styled_high_cardinality_read_only | openpyxl | pass | independent-styled-signature-and-input-integrity | 0.000000 | 0.351343 |

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
