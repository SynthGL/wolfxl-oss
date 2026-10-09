# community-final-core-fresh-process-v2: matched performance evidence

Contract: wolfxl-styled-edit-v2; 5 measured trials and 1 discarded warmup(s) per variant/case.
Every trial uses a fresh process. Stage/control order is reproducibly shuffled and executed serially. Times below are seconds.
Cumulative ratios divide matched baseline/final medians directly; isolated optimization gains are not multiplied.

Receipt SHA256: 5189227adb805978c086601fdb32f60a5a87daa1cae9c9b0d531d71085f6da94
Controller SHA256: c400130f81234d927d0987c5a7005a1331a07bfaf53ec635b82267b98d6b9537
Frozen harness SHA256: fadd85865a17424cc5642e2177fdffbcd1679f37715435dafb00e23006a72922

## Operation and total

| Workload | Baseline operation median [min,max] | Final operation median [min,max] | Cumulative engine gain | Final vs openpyxl operation | Cumulative total gain |
| --- | --- | --- | --- | --- | --- |
| edit_top | 0.498799 [0.480502, 0.544156] | 0.590165 [0.584090, 0.663055] | 0.845× | 25.639× | 0.848× |
| edit_middle | 0.760970 [0.728832, 0.803887] | 0.695379 [0.672053, 0.709544] | 1.094× | 21.115× | 1.003× |
| edit_bottom | 1.063460 [1.052353, 1.181825] | 0.765193 [0.748088, 0.826172] | 1.390× | 19.684× | 1.003× |
| edit_merged | 5.537609 [5.359011, 5.766811] | 0.663033 [0.605751, 0.692191] | 8.352× | Not measured | 8.304× |
| styled_small_eager | 0.061669 [0.048874, 0.069127] | 0.048093 [0.043115, 0.063001] | 1.282× | 2.176× | 1.282× |
| styled_large_read_only | 1.024442 [0.812071, 1.163428] | 0.445728 [0.435581, 0.489925] | 2.298× | 2.030× | 2.298× |
| styled_merged_eager | 0.987439 [0.934179, 1.099716] | 0.805362 [0.751203, 0.887253] | 1.226× | 1.297× | 1.226× |

Operation = load + assignment + save + close for edits; load + identical public Cell/style loop + close for reads.
Edit total adds fixture copy and the same independent bounded openpyxl verification. Read total equals operation.
Full semantic/package validation is outside operation and total and is reported below.

**High variance:** operation range exceeds 20% of its median for styled_small_eager/baseline, styled_small_eager/final, styled_small_eager/openpyxl, styled_large_read_only/baseline, styled_large_read_only/openpyxl.
These five-trial ranges are descriptive, not confidence intervals.

## Phase attribution and raw ranges

### edit_top

| Variant | Phase | Median [min,max] seconds |
| --- | --- | --- |
| baseline | copy | 0.001405 [0.001363, 0.001777] |
| baseline | load | 0.120895 [0.113304, 0.133108] |
| baseline | assignment | 0.000034 [0.000031, 0.000058] |
| baseline | save | 0.369147 [0.363068, 0.410915] |
| baseline | close | 0.000030 [0.000030, 0.000060] |
| baseline | bounded_verification | 0.005626 [0.004870, 0.006826] |
| baseline | operation | 0.498799 [0.480502, 0.544156] |
| baseline | total | 0.506363 [0.486801, 0.552408] |
| final | copy | 0.001460 [0.001388, 0.001542] |
| final | load | 0.006845 [0.006298, 0.006922] |
| final | assignment | 0.299844 [0.298766, 0.333906] |
| final | save | 0.284339 [0.277959, 0.322184] |
| final | close | 0.000030 [0.000029, 0.000032] |
| final | bounded_verification | 0.005634 [0.005139, 0.006071] |
| final | operation | 0.590165 [0.584090, 0.663055] |
| final | total | 0.597444 [0.590752, 0.670592] |
| openpyxl | copy | 0.001455 [0.001415, 0.001568] |
| openpyxl | load | 7.933108 [7.727496, 8.375815] |
| openpyxl | assignment | 0.000078 [0.000074, 0.000079] |
| openpyxl | save | 7.197806 [6.880265, 7.822325] |
| openpyxl | close | 0.000001 [0.000001, 0.000002] |
| openpyxl | bounded_verification | 0.005695 [0.005193, 0.008590] |
| openpyxl | operation | 15.131005 [14.641858, 16.019546] |
| openpyxl | total | 15.138228 [14.650003, 16.026312] |

### edit_middle

| Variant | Phase | Median [min,max] seconds |
| --- | --- | --- |
| baseline | copy | 0.001627 [0.001539, 0.002057] |
| baseline | load | 0.113781 [0.111370, 0.120734] |
| baseline | assignment | 0.000035 [0.000033, 0.000037] |
| baseline | save | 0.647414 [0.617385, 0.686817] |
| baseline | close | 0.000030 [0.000028, 0.000031] |
| baseline | bounded_verification | 2.965935 [2.778319, 3.510901] |
| baseline | operation | 0.760970 [0.728832, 0.803887] |
| baseline | total | 3.744448 [3.536065, 4.316588] |
| final | copy | 0.001554 [0.001469, 0.001901] |
| final | load | 0.006343 [0.006096, 0.006861] |
| final | assignment | 0.308921 [0.298254, 0.325414] |
| final | save | 0.377812 [0.363164, 0.400403] |
| final | close | 0.000027 [0.000026, 0.000034] |
| final | bounded_verification | 3.023352 [2.773043, 3.122181] |
| final | operation | 0.695379 [0.672053, 0.709544] |
| final | total | 3.734372 [3.469963, 3.815772] |
| openpyxl | copy | 0.001487 [0.001443, 0.001637] |
| openpyxl | load | 7.690695 [7.383188, 8.079523] |
| openpyxl | assignment | 0.000079 [0.000076, 0.000080] |
| openpyxl | save | 6.955930 [6.663471, 7.157939] |
| openpyxl | close | 0.000001 [0.000001, 0.000001] |
| openpyxl | bounded_verification | 2.960623 [2.830578, 3.019858] |
| openpyxl | operation | 14.682810 [14.197851, 15.237554] |
| openpyxl | total | 17.603158 [17.102422, 18.255360] |

### edit_bottom

| Variant | Phase | Median [min,max] seconds |
| --- | --- | --- |
| baseline | copy | 0.001475 [0.001440, 0.001904] |
| baseline | load | 0.112249 [0.106966, 0.115507] |
| baseline | assignment | 0.000034 [0.000034, 0.000035] |
| baseline | save | 0.951132 [0.945306, 1.066239] |
| baseline | close | 0.000031 [0.000029, 0.000031] |
| baseline | bounded_verification | 5.850054 [5.787099, 6.651733] |
| baseline | operation | 1.063460 [1.052353, 1.181825] |
| baseline | total | 6.903887 [6.861785, 7.835007] |
| final | copy | 0.001803 [0.001580, 0.002207] |
| final | load | 0.006993 [0.006382, 0.012826] |
| final | assignment | 0.315348 [0.304066, 0.329530] |
| final | save | 0.447098 [0.434651, 0.490027] |
| final | close | 0.000029 [0.000026, 0.000033] |
| final | bounded_verification | 6.101383 [5.660333, 6.803358] |
| final | operation | 0.765193 [0.748088, 0.826172] |
| final | total | 6.883440 [6.427636, 7.553098] |
| openpyxl | copy | 0.001659 [0.001440, 0.001727] |
| openpyxl | load | 7.907920 [7.855239, 8.818959] |
| openpyxl | assignment | 0.000077 [0.000074, 0.000081] |
| openpyxl | save | 6.945418 [6.830807, 7.652686] |
| openpyxl | close | 0.000001 [0.000001, 0.000002] |
| openpyxl | bounded_verification | 5.979720 [5.844125, 6.281854] |
| openpyxl | operation | 15.061853 [14.728421, 15.978553] |
| openpyxl | total | 20.925630 [20.709876, 22.261855] |

### edit_merged

| Variant | Phase | Median [min,max] seconds |
| --- | --- | --- |
| baseline | copy | 0.001125 [0.001068, 0.002397] |
| baseline | load | 3.847101 [3.722077, 4.080224] |
| baseline | assignment | 1.220207 [1.209748, 1.269743] |
| baseline | save | 0.412453 [0.382775, 0.419374] |
| baseline | close | 0.033929 [0.028641, 0.045962] |
| baseline | bounded_verification | 0.020045 [0.018346, 0.021911] |
| baseline | operation | 5.537609 [5.359011, 5.766811] |
| baseline | total | 5.558327 [5.382002, 5.789264] |
| final | copy | 0.001345 [0.001141, 0.001354] |
| final | load | 0.007967 [0.006326, 0.011451] |
| final | assignment | 0.345663 [0.309393, 0.347475] |
| final | save | 0.310475 [0.289594, 0.336701] |
| final | close | 0.000032 [0.000031, 0.000038] |
| final | bounded_verification | 0.005101 [0.004737, 0.005822] |
| final | operation | 0.663033 [0.605751, 0.692191] |
| final | total | 0.669345 [0.612010, 0.699366] |

### styled_small_eager

| Variant | Phase | Median [min,max] seconds |
| --- | --- | --- |
| baseline | load | 0.008214 [0.007700, 0.010504] |
| baseline | close | 0.000780 [0.000683, 0.000928] |
| baseline | iterate_styled_cells | 0.050453 [0.040383, 0.060628] |
| baseline | operation | 0.061669 [0.048874, 0.069127] |
| baseline | total | 0.061669 [0.048874, 0.069127] |
| final | load | 0.006646 [0.006276, 0.009328] |
| final | close | 0.000794 [0.000744, 0.000987] |
| final | iterate_styled_cells | 0.040686 [0.036036, 0.055698] |
| final | operation | 0.048093 [0.043115, 0.063001] |
| final | total | 0.048093 [0.043115, 0.063001] |
| openpyxl | load | 0.076692 [0.072397, 0.102282] |
| openpyxl | close | 0.000003 [0.000002, 0.000004] |
| openpyxl | iterate_styled_cells | 0.027921 [0.026943, 0.043229] |
| openpyxl | operation | 0.104628 [0.100191, 0.145528] |
| openpyxl | total | 0.104628 [0.100191, 0.145528] |

### styled_large_read_only

| Variant | Phase | Median [min,max] seconds |
| --- | --- | --- |
| baseline | load | 0.006706 [0.006289, 0.008779] |
| baseline | close | 0.005243 [0.004817, 0.006255] |
| baseline | iterate_styled_cells | 1.009392 [0.799515, 1.150207] |
| baseline | operation | 1.024442 [0.812071, 1.163428] |
| baseline | total | 1.024442 [0.812071, 1.163428] |
| final | load | 0.007728 [0.006499, 0.014653] |
| final | close | 0.000041 [0.000040, 0.000046] |
| final | iterate_styled_cells | 0.437948 [0.429027, 0.481974] |
| final | operation | 0.445728 [0.435581, 0.489925] |
| final | total | 0.445728 [0.435581, 0.489925] |
| openpyxl | load | 0.004925 [0.004049, 0.012752] |
| openpyxl | close | 0.000026 [0.000023, 0.000030] |
| openpyxl | iterate_styled_cells | 0.899693 [0.791640, 1.144966] |
| openpyxl | operation | 0.904656 [0.795727, 1.149401] |
| openpyxl | total | 0.904656 [0.795727, 1.149401] |

### styled_merged_eager

| Variant | Phase | Median [min,max] seconds |
| --- | --- | --- |
| baseline | load | 0.377332 [0.353995, 0.443434] |
| baseline | close | 0.004671 [0.004571, 0.006658] |
| baseline | iterate_styled_cells | 0.614481 [0.575598, 0.651646] |
| baseline | operation | 0.987439 [0.934179, 1.099716] |
| baseline | total | 0.987439 [0.934179, 1.099716] |
| final | load | 0.006376 [0.006239, 0.007728] |
| final | close | 0.006033 [0.005657, 0.006804] |
| final | iterate_styled_cells | 0.793136 [0.738463, 0.873296] |
| final | operation | 0.805362 [0.751203, 0.887253] |
| final | total | 0.805362 [0.751203, 0.887253] |
| openpyxl | load | 0.733816 [0.716980, 0.748563] |
| openpyxl | close | 0.000002 [0.000002, 0.000003] |
| openpyxl | iterate_styled_cells | 0.310515 [0.300516, 0.322867] |
| openpyxl | operation | 1.044345 [1.017511, 1.060325] |
| openpyxl | total | 1.044345 [1.017511, 1.060325] |

## Independent validation

| Workload | Variant | Result | Scope | Full reopen seconds | Package check seconds |
| --- | --- | --- | --- | --- | --- |
| edit_top | baseline | pass | full output values/types/styles/merges, ZIP/XML, untouched-part bytes | 17.357894 | 4.814619 |
| edit_top | final | pass | full output values/types/styles/merges, ZIP/XML, untouched-part bytes | 16.783272 | 5.205582 |
| edit_top | openpyxl | pass | full output values/types/styles/merges, ZIP/XML, package rewrite reported | 16.025752 | 5.546793 |
| edit_middle | baseline | pass | full output values/types/styles/merges, ZIP/XML, untouched-part bytes | 16.624169 | 5.127169 |
| edit_middle | final | pass | full output values/types/styles/merges, ZIP/XML, untouched-part bytes | 17.122684 | 5.129904 |
| edit_middle | openpyxl | pass | full output values/types/styles/merges, ZIP/XML, package rewrite reported | 17.470880 | 5.322848 |
| edit_bottom | baseline | pass | full output values/types/styles/merges, ZIP/XML, untouched-part bytes | 16.939028 | 5.135812 |
| edit_bottom | final | pass | full output values/types/styles/merges, ZIP/XML, untouched-part bytes | 16.244140 | 5.773825 |
| edit_bottom | openpyxl | pass | full output values/types/styles/merges, ZIP/XML, package rewrite reported | 17.388863 | 5.079722 |
| edit_merged | baseline | pass | full output values/types/styles/merges, ZIP/XML, untouched-part bytes | 17.553567 | 6.060169 |
| edit_merged | final | pass | full output values/types/styles/merges, ZIP/XML, untouched-part bytes | 16.506148 | 5.254601 |
| styled_small_eager | baseline | pass | independent-styled-signature-and-input-integrity | 0.000000 | 0.034264 |
| styled_small_eager | final | pass | independent-styled-signature-and-input-integrity | 0.000000 | 0.033489 |
| styled_small_eager | openpyxl | pass | independent-styled-signature-and-input-integrity | 0.000000 | 0.024525 |
| styled_large_read_only | baseline | pass | independent-styled-signature-and-input-integrity | 0.000000 | 0.241538 |
| styled_large_read_only | final | pass | independent-styled-signature-and-input-integrity | 0.000000 | 0.258873 |
| styled_large_read_only | openpyxl | pass | independent-styled-signature-and-input-integrity | 0.000000 | 0.240909 |
| styled_merged_eager | baseline | pass | independent-styled-signature-and-input-integrity | 0.000000 | 0.228190 |
| styled_merged_eager | final | pass | independent-styled-signature-and-input-integrity | 0.000000 | 0.251767 |
| styled_merged_eager | openpyxl | pass | independent-styled-signature-and-input-integrity | 0.000000 | 0.249700 |

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
