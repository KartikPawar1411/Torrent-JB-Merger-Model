# Torrent Pharma / JB Chemicals | Merger model

**Accretion / dilution analysis of Torrent Pharmaceuticals' cash-and-stock acquisition of JB Chemicals & Pharmaceuticals, modelled at the 29 June 2025 announcement. INR crore unless stated.**

**Result: dilutive to reported EPS, accretive to cash EPS by year two.** On base assumptions the deal cuts Torrent's FY27E EPS by 19.8% and FY28E EPS by 6.3%. Excluding non-cash brand amortisation, FY28E EPS rises 10.1%. Reported EPS breaks even in FY28E only with ₹748 cr of pre-tax synergies, against management's ₹450 cr target.

| Open | What it is |
|---|---|
| [Excel model](output/Torrent_JB_Merger_Model.xlsx) | Seven sheets, live formulas, three Excel data tables. Start at Summary; edit the blue cells on Inputs and Standalone. |
| [Validation](reports/validation.md) | 27 outputs and all 86 data-table cells matched against an independent Python calculation |
| [Source register](data/sources.csv) | Where every non-assumption input came from |

## The deal

Torrent bought JB in two steps. It paid cash for KKR's 46.39% stake and a 2.80% employee stake at ₹1,600 per share, then merged JB into itself by issuing 51 Torrent shares for every 100 JB shares. A mandatory open offer for a further 26% at ₹1,639.18 sat between the two steps.

| | |
|---|---:|
| Total consideration for 100% of JB | ₹26,678 cr |
| Cash (49.19% of JB) | ₹12,636 cr |
| Torrent shares issued (50.81% of JB), at ₹3,375 | ₹14,042 cr |
| Implied enterprise value | ₹25,989 cr |
| EV / EBITDA, FY25A | 25.2x |
| P / E, FY25A (Torrent's own: 59.8x) | 40.4x |
| Blended price per JB share vs ₹1,799.35 pre-announcement close | ₹1,662 (7.7% discount) |
| New acquisition debt, including ₹100 cr of fees | ₹12,736 cr |
| New Torrent shares | 4.16 cr (JB public owns 10.9% of Torrent) |
| Net debt / EBITDA at closing, FY26E | 2.66x |

The buyer paid a discount, not a premium. KKR sold a control block below the market price, and the open offer was priced 8.9% below the last close. The base case therefore assumes no one tenders into the open offer.

## Earnings impact

| | FY27E | FY28E |
|---|---:|---:|
| Torrent standalone EPS | ₹79.97 | ₹92.97 |
| Pro forma EPS | ₹64.12 | ₹87.11 |
| **EPS accretion / (dilution)** | **(19.8%)** | **(6.3%)** |
| Cash EPS accretion / (dilution) | (0.7%) | 10.1% |
| Pre-tax synergies needed for breakeven | ₹1,030 cr | ₹748 cr |

FY27E bridge, pre-tax: ₹225 cr of synergies, less ₹976 cr of interest on the new debt, less ₹775 cr of brand amortisation. JB adds ₹872 cr of profit, and the share count rises 12.3%.

Three things drive the answer:

1. **Amortisation, not financing, decides the sign.** Allocating 30% instead of 50% of the excess purchase price to brands, over 25 years instead of 15, cuts FY27E dilution from 19.8% to 7.6%.
2. **Synergies help less than the headline suggests.** Each ₹150 cr of run-rate synergy adds about 1.8 points to FY27E EPS. Even ₹750 cr leaves FY27E 16.1% dilutive.
3. **Debt is worse than stock here.** Torrent trades at 59.8x earnings, so its shares are cheap currency. Full acceptance of the open offer would swap ₹6,843 cr of stock for debt and deepen FY28E dilution from 6.3% to 12.3%.

## How it compares with what happened

| | Model | Reported |
|---|---|---|
| FY27E EPS dilution | 19.8% | About 15% (PL Capital estimate at announcement) |
| Net debt / EBITDA, FY26E | 2.66x | 2.6x (PL Capital estimate) |
| Net debt / EBITDA, FY27E | 1.86x | 1.8x (PL Capital estimate) |
| Open offer acceptance | 0% assumed | 1,317 shares tendered |
| Acquisition debt | ₹12,736 cr | ₹12,500 cr approved: ₹11,000 cr NCDs and ₹1,500 cr commercial paper |
| New Torrent shares issued in the merger | 4.16 cr | 4.19 cr (4,19,22,416 shares allotted on 20 July 2026) |
| Torrent shares after the merger | 38.00 cr | 38.04 cr |

The acquisition closed on 21 January 2026 and the merger took effect in July 2026. The model issues slightly fewer shares than Torrent did because it uses the announced 49.19% cash stake; Torrent held about 48.8% before the merger, so more JB shares were swapped. The broker figures are cross-checks only and are not model inputs.

**Two figures that differ from press reports, and why.**

- *₹26,678 cr against the reported ₹25,689 cr.* The reported figure prices every JB share at the ₹1,600 cash price. The model's figure is the consideration Torrent gives up: ₹12,636 cr of cash plus 4.16 cr Torrent shares at ₹3,375, which values the swapped shares at ₹1,721 each.
- *2.66x leverage against management's 1.8x.* The model's ratio is total net debt, including Torrent's existing ₹2,250 cr, over FY26E EBITDA. Torrent's CFO guided 1.8x without the open offer and 2.8x with full acceptance, without stating the basis. Both figures are consistent with acquisition debt alone (₹12,636 cr and ₹19,479 cr) over about ₹7,000 cr of EBITDA; that reading is an inference.

## Rebuild

Requires Windows, desktop Excel and Python 3.12 or later.

```powershell
python -m pip install -r requirements.txt
python build.py
powershell -ExecutionPolicy Bypass -File finalize.ps1
python verify.py
```

`build.py` writes the workbook with formulas. `finalize.ps1` opens it in Excel, creates the data tables and recalculates. `verify.py` recomputes the deal in plain Python and fails if any output or data-table cell differs.

## Model conventions

- Pro forma earnings assume the merger is effective for the whole of each year. The interim period, when Torrent owned 49% and JB had minority shareholders, is not modelled.
- Sponsor cash consideration and fees are fully debt-funded at 7.66%, the weighted coupon on the NCDs Torrent later issued. ₹2,500 cr is assumed repaid before FY28E.
- 50% of the excess purchase price is allocated to brands and amortised over 15 years, with a deferred tax liability at 25.17%. The rest is goodwill, which is not amortised under Ind AS.
- The sensitivity tables are native Excel data tables. Their input cells sit on the Sensitivity sheet and must stay blank; the Inputs sheet reads them through the "Live case" block.

## Limitations

- **Inputs come from secondary sources** (a data aggregator and press reports), not from the filings themselves. Reconcile them to the annual reports and the scheme document before relying on any figure.
- FY26E to FY28E forecasts are my assumptions, not consensus. Torrent's standalone growth and margin drive the dilution percentage as much as the deal terms do.
- The purchase price allocation is an assumption. Torrent's actual allocation will be in its FY26 annual report.
- JB's employee stock options are captured only through the fully diluted share count. Exercise proceeds are ignored.
- No revenue synergies, integration costs, or change in JB's tax rate after the merger.

Independent analysis from public information. Not investment advice.
