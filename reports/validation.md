# Validation

Result: **ALL PASS**. 27 headline outputs, 86 data-table cells and 12 in-workbook checks compared; 0 Excel error cells.

The reference values come from `verify.py`, a separate plain-Python calculation that shares only the input values with the workbook.

## Headline outputs

| Output | Independent | Excel | Result |
|---|---:|---:|---|
| acc27 | -0.198201 | -0.198201 | pass |
| acc28 | -0.063041 | -0.063041 | pass |
| amort | 774.804697 | 774.804697 | pass |
| breakeven27 | 1,029.886830 | 1,029.886830 | pass |
| breakeven28 | 747.619255 | 747.619255 | pass |
| cash_acc27 | -0.007408 | -0.007408 | pass |
| cash_acc28 | 0.101076 | 0.101076 | pass |
| cash_total | 12,636.419100 | 12,636.419100 | pass |
| consid | 26,678.140896 | 26,678.140896 | pass |
| fd_shares | 16.055625 | 16.055625 | pass |
| goodwill | 14,547.345580 | 14,547.345580 | pass |
| implied_ev | 25,989.140896 | 25,989.140896 | pass |
| intangibles | 11,622.070448 | 11,622.070448 | pass |
| lev26 | 2.663528 | 2.663528 | pass |
| new_debt | 12,736.419100 | 12,736.419100 | pass |
| new_shares | 4.160510 | 4.160510 | pass |
| pf_eps27 | 64.117957 | 64.117957 | pass |
| pf_eps28 | 87.105193 | 87.105193 | pass |
| pf_pat27 | 2,436.515065 | 2,436.515065 | pass |
| pf_pat28 | 3,310.041784 | 3,310.041784 | pass |
| pf_shares | 38.000510 | 38.000510 | pass |
| sa_eps27 | 79.967660 | 79.967660 | pass |
| sa_eps28 | 92.965864 | 92.965864 | pass |
| stock_value | 14,041.721796 | 14,041.721796 | pass |
| summary_acc27 | -0.198201 | -0.198201 | pass |
| summary_acc28 | -0.063041 | -0.063041 | pass |
| summary_lev26 | 2.663528 | 2.663528 | pass |

## Data tables (each cell recomputed by a full independent run)

| Table | Cells | Largest difference | Result |
|---|---:|---:|---|
| t1 | 36 | 4.16e-16 | pass |
| t2 | 25 | 3.19e-16 | pass |
| t3 | 25 | 4.16e-16 | pass |

## In-workbook checks

| Check | Excel |
|---|---|
| Sources equal uses | True |
| Cash and stock stakes cover 100% of JB | True |
| Pro forma ownership sums to 100% | True |
| Purchase price allocation ties to consideration | True |
| Goodwill is positive | True |
| New debt is not negative | True |
| Open offer acceptance is between 0% and 100% | True |
| Torrent FY25A EPS ties to reported 56.47 | True |
| JB FY25A EPS ties to reported 42.36 | True |
| Breakeven synergies give zero dilution, FY27E | True |
| Sensitivity input cells are blank | True |
| All checks pass | True |
