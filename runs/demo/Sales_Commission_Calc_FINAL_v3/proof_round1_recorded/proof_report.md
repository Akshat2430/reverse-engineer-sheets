# Proof report

**99.9% of 3000 input sets matched on every output**: 301 recalculated fresh in Excel (100.0%) and 2699 probe runs replayed against Excel's recorded answers (99.9%, date pinned to the probe day 2026-10-28). (20/43 outputs never mismatched). Tolerance 0.005 on stored values. Engine: excel (via bridge) 16.114.2, locale en_IN, today pinned to 2026-10-28.

Case mix: {'saved inputs': 1, 'random': 300, 'replay': 2699}

## Outputs with mismatches

| Output | Pass | Fail |
|---|---|---|
| Summary!C5 | 2996 | 4 |
| Summary!C19 | 2996 | 4 |
| Summary!C24 | 2996 | 4 |
| Summary!C22 | 2997 | 3 |
| Summary!D5 | 2997 | 3 |
| Summary!C12 | 2998 | 2 |
| Summary!C15 | 2998 | 2 |
| Summary!C16 | 2998 | 2 |
| Summary!C21 | 2998 | 2 |
| Summary!D12 | 2998 | 2 |
| Summary!D16 | 2998 | 2 |
| Summary!C6 | 2999 | 1 |
| Summary!C7 | 2999 | 1 |
| Summary!C8 | 2999 | 1 |
| Summary!C10 | 2999 | 1 |
| Summary!C11 | 2999 | 1 |
| Summary!C13 | 2999 | 1 |
| Summary!C20 | 2999 | 1 |
| Summary!C26 | 2999 | 1 |
| Summary!D7 | 2999 | 1 |
| Summary!D10 | 2999 | 1 |
| Summary!D11 | 2999 | 1 |
| Summary!D13 | 2999 | 1 |

## Mismatches by input block changed

| Input block | Failing cases | Cases where it changed |
|---|---|---|
| Rates!A2:A10 | 4 | 229 |

## First failing cases

### Case 2491 (replay: blank)
Changed: Rates!A2=None
- Summary!C5: Excel '#N/A' vs model 307615.0
- Summary!C16: Excel '#N/A' vs model 36562.25
- Summary!C19: Excel '#N/A' vs model 447049.675
- Summary!C22: Excel '#N/A' vs model 155487.0
- Summary!C24: Excel '#N/A' vs model 1541655.4259000001
- Summary!C26: Excel '#N/A' vs model 0.10073545647543126
- Summary!D16: Excel '#N/A' vs model 6765.0

### Case 2493 (replay: blank)
Changed: Rates!A5=None
- Summary!C5: Excel 307615.0 vs model 309455.0 (diff 1840)
- Summary!C12: Excel 98333.56250000001 vs model 99353.56250000001 (diff 1020)
- Summary!C15: Excel 118924.75000000001 vs model 121864.75000000001 (diff 2940)
- Summary!C19: Excel 447049.675 vs model 448889.675 (diff 1840)
- Summary!C21: Excel 644374.1625 vs model 645394.1625 (diff 1020)
- Summary!C22: Excel 155487.0 vs model 158427.0 (diff 2940)
- Summary!C24: Excel 1541655.4259000001 vs model 1547455.4259000001 (diff 5800)
- Summary!D5: Excel 306910.0 vs model 308750.0 (diff 1840)
- Summary!D12: Excel 4207.5 vs model 5227.5 (diff 1020)

### Case 2497 (replay: blank)
Changed: Rates!A6=None
- Summary!C5: Excel 307615.0 vs model 309240.0 (diff 1625)
- Summary!C6: Excel 134158.1625 vs model 131796.3625 (diff 2361.8)
- Summary!C7: Excel 7638.3125 vs model 9073.3125 (diff 1435)
- Summary!C8: Excel 100189.23062500001 vs model 98156.23062500001 (diff 2033)
- Summary!C10: Excel 69575.827125 vs model 67577.60212499999 (diff 1998.23)
- Summary!C11: Excel 157688.5 vs model 158388.5 (diff 700)
- Summary!C12: Excel 98333.56250000001 vs model 99783.56250000001 (diff 1450)
- Summary!C13: Excel 389472.1 vs model 387152.1 (diff 2320)
- Summary!C15: Excel 117524.75000000001 vs model 120394.75000000001 (diff 2870)
- Summary!C16: Excel 36422.25 vs model 36562.25 (diff 140)
- Summary!C19: Excel 449411.475 vs model 450109.675 (diff 698.2)
- Summary!C20: Excel 295475.81340000004 vs model 291444.5884 (diff 4031.23)
- Summary!C21: Excel 646694.1625 vs model 646524.1625 (diff 170)
- Summary!C22: Excel 153947.0 vs model 156957.0 (diff 3010)
- Summary!C24: Excel 1548828.4508999998 vs model 1548335.4259000001 (diff 493.025)

### Case 2501 (replay: blank)
Changed: Rates!A10=None
- Summary!C5: Excel 328495.0 vs model 307615.0 (diff 20880)
- Summary!C19: Excel 467929.67500000005 vs model 447049.675 (diff 20880)
- Summary!C24: Excel 1562535.4259000001 vs model 1541655.4259000001 (diff 20880)
- Summary!D5: Excel 327790.0 vs model 306910.0 (diff 20880)

