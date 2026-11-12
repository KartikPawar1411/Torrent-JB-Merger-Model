"""Single source of assumptions for the Torrent / JB Chemicals merger model.

All money in INR crore, shares in crore, per-share values in INR.
Every figure carries the source key it came from (see data/sources.csv) or is
marked "Analyst assumption".
"""

DEAL = {
    # --- market data at last close before the 29 June 2025 announcement ---
    "acq_price": (3375.00, "INR", "Torrent close, 27 Jun 2025", "S5"),
    "acq_shares": (33.84, "cr", "FY25 PAT 1,911 / EPS 56.47", "S1"),
    "tgt_price": (1799.35, "INR", "JB close on BSE, 27 Jun 2025", "S3"),
    # --- announced terms ---
    "tgt_equity_fd": (25689.0, "INR cr", "Fully diluted equity value at the SPA price", "S3"),
    "cash_price": (1600.00, "INR", "SPA and employee purchase price per share", "S3"),
    "stake_spa": (0.4639, "%", "Stake bought from KKR (Tau Investment Holdings)", "S3"),
    "stake_emp": (0.0280, "%", "Stake bought from JB employees", "S3"),
    "oo_size": (0.26, "%", "Mandatory open offer size", "S3"),
    "oo_price": (1639.18, "INR", "Open offer price per share", "S3"),
    "oo_accept_base": (0.0, "% of offer", "Offer was below market; 1,317 shares were tendered", "S6"),
    "swap_base": (0.51, "x", "51 Torrent shares per 100 JB shares", "S3"),
    "fees": (100.0, "INR cr", "Advisory, stamp duty and financing fees", "Analyst assumption"),
    # --- financing ---
    "cod_base": (0.0766, "%", "Weighted coupon on INR 11,000 cr of NCDs (7.45%-7.80%)", "S7"),
    "cash_used": (0.0, "INR cr", "Balance-sheet cash applied to the purchase", "Analyst assumption"),
    "repay_fy28": (2500.0, "INR cr", "Acquisition debt repaid before FY28E", "Analyst assumption"),
    "tax_adj": (0.2517, "%", "Indian corporate tax rate applied to deal adjustments", "Analyst assumption"),
    # --- purchase price allocation ---
    "tgt_book": (3434.0, "INR cr", "JB equity capital 16 + reserves 3,418, Mar 2025", "S2"),
    "alloc_base": (0.50, "%", "Share of excess purchase price allocated to brands", "Analyst assumption"),
    "life_base": (15.0, "years", "Amortisation life of acquired brands", "Analyst assumption"),
    # --- synergies ---
    "syn_base": (450.0, "INR cr", "Annual run-rate savings targeted by Torrent", "S8"),
    "phase27": (0.50, "%", "Share of run-rate realised in FY27E", "Analyst assumption"),
    "phase28": (1.00, "%", "Share of run-rate realised in FY28E", "Analyst assumption"),
    # --- balance sheet ---
    "acq_net_debt": (2250.0, "INR cr", "Torrent net debt before the deal", "S10"),
    "tgt_net_cash": (689.0, "INR cr", "JB net cash, 31 Mar 2025", "S9"),
}

LABELS = {
    "acq_price": "Torrent share price",
    "acq_shares": "Torrent shares outstanding",
    "tgt_price": "JB share price before announcement",
    "tgt_equity_fd": "JB fully diluted equity value at SPA price",
    "cash_price": "Cash price per JB share (SPA)",
    "stake_spa": "Stake acquired from KKR",
    "stake_emp": "Stake acquired from employees",
    "oo_size": "Open offer size",
    "oo_price": "Open offer price",
    "oo_accept_base": "Open offer acceptance (base)",
    "swap_base": "Merger swap ratio (base)",
    "fees": "Transaction fees",
    "cod_base": "Cost of acquisition debt (base)",
    "cash_used": "Balance-sheet cash used",
    "repay_fy28": "Acquisition debt repaid before FY28E",
    "tax_adj": "Tax rate on adjustments",
    "tgt_book": "JB book equity",
    "alloc_base": "Excess allocated to intangibles (base)",
    "life_base": "Intangible amortisation life (base)",
    "syn_base": "Run-rate pre-tax synergies (base)",
    "phase27": "Synergy phasing FY27E",
    "phase28": "Synergy phasing FY28E",
    "acq_net_debt": "Torrent net debt",
    "tgt_net_cash": "JB net cash",
}

SECTIONS = [
    ("Market data (27 June 2025)", ["acq_price", "acq_shares", "tgt_price"]),
    ("Announced deal terms", ["tgt_equity_fd", "cash_price", "stake_spa", "stake_emp", "oo_size",
                              "oo_price", "oo_accept_base", "swap_base", "fees"]),
    ("Financing", ["cod_base", "cash_used", "repay_fy28", "tax_adj"]),
    ("Purchase price allocation", ["tgt_book", "alloc_base", "life_base"]),
    ("Synergies", ["syn_base", "phase27", "phase28"]),
    ("Balance sheet", ["acq_net_debt", "tgt_net_cash"]),
]

# The six inputs the sensitivity tables flex: live name -> base name
FLEX = {"syn": "syn_base", "cod": "cod_base", "oo_accept": "oo_accept_base",
        "swap": "swap_base", "alloc": "alloc_base", "life": "life_base"}

# FY25A reported (consolidated, INR cr) and FY26E-FY28E drivers (analyst assumptions)
COMPANIES = {
    "Torrent Pharmaceuticals": {
        "src": "S1",
        "fy25": {"rev": 11516.0, "ebitda": 3721.0, "da": 795.0, "oi": -1.0, "int": 252.0,
                 "tax": 762.0, "shares": 33.84},
        "growth": [0.12, 0.12, 0.11],
        "margin": [0.325, 0.330, 0.335],
        "da_pct": [0.066, 0.063, 0.060],
        "oi": [0.0, 0.0, 0.0],
        "int": [200.0, 150.0, 100.0],
        "tax_rate": [0.27, 0.27, 0.27],
    },
    "JB Chemicals & Pharmaceuticals": {
        "src": "S2",
        "fy25": {"rev": 3918.0, "ebitda": 1032.0, "da": 171.0, "oi": 38.0, "int": 12.0,
                 "tax": 227.0, "shares": 15.58},
        "growth": [0.12, 0.12, 0.12],
        "margin": [0.268, 0.275, 0.280],
        "da_pct": [0.044, 0.044, 0.044],
        "oi": [40.0, 40.0, 40.0],
        "int": [5.0, 5.0, 5.0],
        "tax_rate": [0.255, 0.255, 0.255],
    },
}

SENS = {
    "t1": {"title": "FY27E EPS accretion / (dilution): run-rate synergies (down) vs cost of debt (across)",
           "row_name": "cod", "row_vals": [0.0700, 0.0725, 0.0750, 0.0766, 0.0800, 0.0850],
           "col_name": "syn", "col_vals": [0.0, 150.0, 300.0, 450.0, 600.0, 750.0], "year": "FY27E"},
    "t2": {"title": "FY28E EPS accretion / (dilution): open offer acceptance (down) vs swap ratio (across)",
           "row_name": "swap", "row_vals": [0.45, 0.48, 0.51, 0.54, 0.57],
           "col_name": "oo_accept", "col_vals": [0.0, 0.25, 0.50, 0.75, 1.0], "year": "FY28E"},
    "t3": {"title": "FY27E EPS accretion / (dilution): intangible allocation (down) vs amortisation life (across)",
           "row_name": "life", "row_vals": [10.0, 12.0, 15.0, 20.0, 25.0],
           "col_name": "alloc", "col_vals": [0.30, 0.40, 0.50, 0.60, 0.70], "year": "FY27E"},
}
