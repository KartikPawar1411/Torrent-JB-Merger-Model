"""Build the Torrent / JB Chemicals merger model workbook (live Excel formulas).

Run:  python build.py
Then: powershell -File finalize.ps1   (creates the Excel data tables and recalculates)
"""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.workbook.defined_name import DefinedName

from inputs import COMPANIES, DEAL, FLEX, LABELS, SECTIONS, SENS

OUT = Path(__file__).parent / "output" / "Torrent_JB_Merger_Model.xlsx"

NAVY = "1F3864"
F_IN = Font(name="Calibri", size=10, color="0000FF")
F_FX = Font(name="Calibri", size=10, color="000000")
F_B = Font(name="Calibri", size=10, bold=True)
F_NOTE = Font(name="Calibri", size=9, italic=True, color="595959")
F_H = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
F_T = Font(name="Calibri", size=14, bold=True, color=NAVY)
FILL_H = PatternFill("solid", fgColor=NAVY)
FILL_IN = PatternFill("solid", fgColor="FFF2CC")
FILL_KEY = PatternFill("solid", fgColor="DDEBF7")
TOP = Border(top=Side(style="thin"))

NUM = "#,##0;(#,##0);-"
NUM2 = "#,##0.00;(#,##0.00);-"
PCT = "0.0%;(0.0%);-"
PCT2 = "0.00%;(0.00%);-"
MULT = '0.00"x"'
FMT = {"INR cr": NUM, "INR": NUM2, "cr": NUM2, "%": PCT2, "% of offer": PCT, "x": MULT, "years": "0"}

wb = Workbook()
YEARS = ["FY25A", "FY26E", "FY27E", "FY28E"]


def name(n, ws, cell):
    wb.defined_names[n] = DefinedName(n, attr_text=f"'{ws.title}'!${cell[0]}${cell[1:]}")


def sheet(title, widths, heading, sub):
    ws = wb.create_sheet(title)
    ws.sheet_view.showGridLines = False
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws["B2"] = heading
    ws["B2"].font = F_T
    ws["B3"] = sub
    ws["B3"].font = F_NOTE
    return ws


def header(ws, row, labels, start_col=2):
    for i, text in enumerate(labels):
        c = ws.cell(row=row, column=start_col + i, value=text)
        c.font, c.fill = F_H, FILL_H
        c.alignment = Alignment(horizontal="left" if i == 0 else "right")


def put(ws, cell, value, fmt=None, inp=False, bold=False, key=False, total=False):
    c = ws[cell]
    c.value = value
    c.font = F_IN if inp else (F_B if bold else F_FX)
    if inp:
        c.fill = FILL_IN
    if key:
        c.fill = FILL_KEY
    if fmt:
        c.number_format = fmt
    if total:
        c.border = TOP
    return c


def label(ws, cell, text, bold=False, note=False):
    c = ws[cell]
    c.value = text
    c.font = F_NOTE if note else (F_B if bold else F_FX)


# ---------------------------------------------------------------- Summary (filled last)
ws_sum = wb.active
ws_sum.title = "Summary"

# ---------------------------------------------------------------- Inputs
ws = sheet("Inputs", {"A": 2, "B": 44, "C": 14, "D": 11, "E": 58, "F": 10},
           "Inputs", "Blue cells on yellow are inputs. INR crore unless stated. Source keys refer to data/sources.csv.")
r = 5
for title, keys in SECTIONS:
    header(ws, r, [title, "Value", "Unit", "Basis", "Source"])
    r += 1
    for k in keys:
        val, unit, basis, src = DEAL[k]
        label(ws, f"B{r}", LABELS[k])
        put(ws, f"C{r}", val, FMT[unit], inp=True)
        label(ws, f"D{r}", unit)
        label(ws, f"E{r}", basis, note=True)
        label(ws, f"F{r}", src, note=True)
        name(k, ws, f"C{r}")
        r += 1
    r += 1

header(ws, r, ["Live case (base input unless a sensitivity table is flexing it)", "Value", "", "", ""])
r += 1
LIVE_LABEL = {"syn": "Run-rate pre-tax synergies", "cod": "Cost of acquisition debt",
              "oo_accept": "Open offer acceptance", "swap": "Merger swap ratio",
              "alloc": "Excess allocated to intangibles", "life": "Intangible amortisation life"}
for live, base in FLEX.items():
    label(ws, f"B{r}", LIVE_LABEL[live])
    put(ws, f"C{r}", f'=IF(ovr_{live}="",{base},ovr_{live})', FMT[DEAL[base][1]])
    name(live, ws, f"C{r}")
    r += 1

# ---------------------------------------------------------------- Standalone
ws = sheet("Standalone", {"A": 2, "B": 36, "C": 12, "D": 12, "E": 12, "F": 12},
           "Standalone forecasts", "FY25A is reported (consolidated). FY26E-FY28E drivers are analyst assumptions.")
ROWS = {}


def company(r0, cname, d):
    header(ws, r0 - 1, [cname] + YEARS)
    a = d["fy25"]
    rr = {k: r0 + i for i, k in enumerate(
        ["rev", "growth", "ebitda", "margin", "da", "da_pct", "oi", "int", "pbt", "tax", "tax_rate",
         "pat", "shares", "eps"])}
    labels = {"rev": "Revenue", "growth": "  growth", "ebitda": "EBITDA", "margin": "  margin",
              "da": "Depreciation & amortisation", "da_pct": "  % of revenue", "oi": "Other income",
              "int": "Interest expense", "pbt": "Profit before tax", "tax": "Tax", "tax_rate": "  tax rate",
              "pat": "Profit after tax", "shares": "Shares outstanding (cr)", "eps": "EPS (INR)"}
    for k, row in rr.items():
        label(ws, f"B{row}", labels[k], bold=k in ("rev", "ebitda", "pat", "eps"))
    # FY25A
    for k in ("rev", "ebitda", "da", "oi", "int", "tax", "shares"):
        put(ws, f"C{rr[k]}", a[k], NUM2 if k == "shares" else NUM, inp=True)
    put(ws, f"C{rr['margin']}", f"=C{rr['ebitda']}/C{rr['rev']}", PCT)
    put(ws, f"C{rr['da_pct']}", f"=C{rr['da']}/C{rr['rev']}", PCT)
    put(ws, f"C{rr['tax_rate']}", f"=C{rr['tax']}/C{rr['pbt']}", PCT)
    for j, col in enumerate("DEF"):
        prev = "CDE"[j]
        put(ws, f"{col}{rr['growth']}", d["growth"][j], PCT, inp=True)
        put(ws, f"{col}{rr['rev']}", f"={prev}{rr['rev']}*(1+{col}{rr['growth']})", NUM, bold=True)
        put(ws, f"{col}{rr['margin']}", d["margin"][j], PCT, inp=True)
        put(ws, f"{col}{rr['ebitda']}", f"={col}{rr['rev']}*{col}{rr['margin']}", NUM, bold=True)
        put(ws, f"{col}{rr['da_pct']}", d["da_pct"][j], PCT, inp=True)
        put(ws, f"{col}{rr['da']}", f"={col}{rr['rev']}*{col}{rr['da_pct']}", NUM)
        put(ws, f"{col}{rr['oi']}", d["oi"][j], NUM, inp=True)
        put(ws, f"{col}{rr['int']}", d["int"][j], NUM, inp=True)
        put(ws, f"{col}{rr['tax_rate']}", d["tax_rate"][j], PCT, inp=True)
        put(ws, f"{col}{rr['tax']}", f"={col}{rr['pbt']}*{col}{rr['tax_rate']}", NUM)
        put(ws, f"{col}{rr['shares']}", f"={prev}{rr['shares']}", NUM2)
    for col in "CDEF":
        put(ws, f"{col}{rr['pbt']}",
            f"={col}{rr['ebitda']}-{col}{rr['da']}+{col}{rr['oi']}-{col}{rr['int']}", NUM, total=True)
        put(ws, f"{col}{rr['pat']}", f"={col}{rr['pbt']}-{col}{rr['tax']}", NUM, bold=True, total=True)
        put(ws, f"{col}{rr['eps']}", f"={col}{rr['pat']}/{col}{rr['shares']}", NUM2, bold=True)
    return rr


names = list(COMPANIES)
A = company(6, names[0], COMPANIES[names[0]])      # Torrent rows 6-19
T = company(23, names[1], COMPANIES[names[1]])     # JB rows 23-36
ROWS["A"], ROWS["T"] = A, T
header(ws, 39, ["Combined before deal adjustments"] + YEARS)
for i, (k, lab) in enumerate([("rev", "Revenue"), ("ebitda", "EBITDA"), ("pat", "Profit after tax")]):
    row = 40 + i
    label(ws, f"B{row}", lab)
    for col in "CDEF":
        put(ws, f"{col}{row}", f"={col}{A[k]}+{col}{T[k]}", NUM)
COMB = {"rev": 40, "ebitda": 41, "pat": 42}

# ---------------------------------------------------------------- Deal
ws = sheet("Deal", {"A": 2, "B": 44, "C": 12, "D": 14, "E": 14, "F": 16},
           "Deal structure, consideration and Sources & Uses",
           "Cash for the KKR, employee and open-offer shares; Torrent shares for the rest.")
label(ws, "B5", "JB fully diluted equity value at SPA price")
put(ws, "F5", "=tgt_equity_fd", NUM)
label(ws, "B6", "JB fully diluted shares (cr)")
put(ws, "F6", "=tgt_equity_fd/cash_price", "#,##0.0000")
name("fd_shares", ws, "F6")

header(ws, 8, ["Consideration", "% of JB", "Shares (cr)", "INR / share", "INR cr"])
rows = [(9, "Cash: KKR stake (SPA)", "=stake_spa", "=cash_price"),
        (10, "Cash: employee shares", "=stake_emp", "=cash_price"),
        (11, "Cash: open offer accepted", "=oo_size*oo_accept", "=oo_price")]
for row, lab, pct, price in rows:
    label(ws, f"B{row}", lab)
    put(ws, f"C{row}", pct, PCT2)
    put(ws, f"D{row}", f"=fd_shares*C{row}", "#,##0.0000")
    put(ws, f"E{row}", price, NUM2)
    put(ws, f"F{row}", f"=D{row}*E{row}", NUM)
label(ws, "B12", "Total cash consideration", bold=True)
put(ws, "C12", "=SUM(C9:C11)", PCT2, bold=True, total=True)
put(ws, "D12", "=SUM(D9:D11)", "#,##0.0000", bold=True, total=True)
put(ws, "E12", "=F12/D12", NUM2, bold=True, total=True)
put(ws, "F12", "=SUM(F9:F11)", NUM, bold=True, total=True)
name("cash_total", ws, "F12")

label(ws, "B14", "Stock: JB shares swapped in the merger")
put(ws, "C14", "=1-C12", PCT2)
put(ws, "D14", "=fd_shares-D12", "#,##0.0000")
put(ws, "E14", "=swap*acq_price", NUM2)
put(ws, "F14", "=D14*E14", NUM)
name("stock_value", ws, "F14")
label(ws, "B15", "New Torrent shares issued (cr)")
put(ws, "D15", "=D14*swap", "#,##0.0000", key=True)
name("new_shares", ws, "D15")

label(ws, "B17", "Total consideration for 100% of JB", bold=True)
put(ws, "F17", "=cash_total+stock_value", NUM, bold=True, total=True, key=True)
name("consid", ws, "F17")
label(ws, "B18", "Blended price per JB share (INR)")
put(ws, "F18", "=consid/fd_shares", NUM2)
label(ws, "B19", "Premium / (discount) to pre-announcement price")
put(ws, "F19", "=F18/tgt_price-1", PCT)
label(ws, "B20", "Cash price premium / (discount)")
put(ws, "F20", "=cash_price/tgt_price-1", PCT)
label(ws, "B21", "Swap value premium / (discount)")
put(ws, "F21", "=E14/tgt_price-1", PCT)

header(ws, 23, ["Implied valuation", "", "", "", "INR cr"])
label(ws, "B24", "Less: JB net cash")
put(ws, "F24", "=-tgt_net_cash", NUM)
label(ws, "B25", "Implied enterprise value", bold=True)
put(ws, "F25", "=consid+F24", NUM, bold=True, total=True)
name("implied_ev", ws, "F25")
label(ws, "B26", "EV / EBITDA, FY25A")
put(ws, "F26", f"=implied_ev/Standalone!C{T['ebitda']}", MULT)
label(ws, "B27", "EV / EBITDA, FY26E")
put(ws, "F27", f"=implied_ev/Standalone!D{T['ebitda']}", MULT)
label(ws, "B28", "P / E, FY25A")
put(ws, "F28", f"=consid/Standalone!C{T['pat']}", MULT)
label(ws, "B29", "Torrent P / E, FY25A (for comparison)")
put(ws, "F29", f"=acq_price/Standalone!C{A['eps']}", MULT)

header(ws, 31, ["Uses", "", "", "", "INR cr"])
uses = [(32, "Cash to KKR", "=F9"), (33, "Cash to employees", "=F10"), (34, "Cash to open-offer holders", "=F11"),
        (35, "Torrent shares issued to JB shareholders", "=stock_value"), (36, "Transaction fees", "=fees")]
for row, lab, f in uses:
    label(ws, f"B{row}", lab)
    put(ws, f"F{row}", f, NUM)
label(ws, "B37", "Total uses", bold=True)
put(ws, "F37", "=SUM(F32:F36)", NUM, bold=True, total=True)
header(ws, 39, ["Sources", "", "", "", "INR cr"])
label(ws, "B40", "New acquisition debt")
put(ws, "F40", "=cash_total+fees-cash_used", NUM, key=True)
name("new_debt", ws, "F40")
label(ws, "B41", "Balance-sheet cash")
put(ws, "F41", "=cash_used", NUM)
label(ws, "B42", "Torrent shares issued")
put(ws, "F42", "=stock_value", NUM)
label(ws, "B43", "Total sources", bold=True)
put(ws, "F43", "=SUM(F40:F42)", NUM, bold=True, total=True)
label(ws, "B44", "Sources less uses (must be zero)", note=True)
put(ws, "F44", "=F43-F37", NUM2)

header(ws, 46, ["Pro forma ownership of Torrent", "", "", "Shares (cr)", "%"])
label(ws, "B47", "Existing Torrent shareholders")
put(ws, "E47", "=acq_shares", NUM2)
put(ws, "F47", "=E47/E49", PCT)
label(ws, "B48", "Former JB public shareholders")
put(ws, "E48", "=new_shares", NUM2)
put(ws, "F48", "=E48/E49", PCT)
label(ws, "B49", "Pro forma shares", bold=True)
put(ws, "E49", "=E47+E48", NUM2, bold=True, total=True)
put(ws, "F49", "=F47+F48", PCT, bold=True, total=True)
name("pf_shares", ws, "E49")

# ---------------------------------------------------------------- PPA
ws = sheet("PPA", {"A": 2, "B": 50, "C": 16},
           "Purchase price allocation", "Excess over book is split between amortising brands and goodwill.")
header(ws, 5, ["", "INR cr"])
ppa = [(6, "Total consideration", "=consid", NUM, False),
       (7, "Less: JB book equity", "=-tgt_book", NUM, False),
       (8, "Excess purchase price", "=C6+C7", NUM, True),
       (10, "Allocated to identifiable intangibles (%)", "=alloc", PCT, False),
       (11, "Identifiable intangibles (brands)", "=C8*C10", NUM, False),
       (12, "Deferred tax liability on write-up", "=C11*tax_adj", NUM, False),
       (13, "Goodwill", "=C8-C11+C12", NUM, True),
       (15, "Amortisation life (years)", "=life", "0", False),
       (16, "Annual pre-tax amortisation", "=C11/C15", NUM, True),
       (18, "Check: book + intangibles - DTL + goodwill - consideration", "=tgt_book+C11-C12+C13-consid", NUM2, False)]
for row, lab, f, fmt, tot in ppa:
    label(ws, f"B{row}", lab, bold=tot)
    put(ws, f"C{row}", f, fmt, bold=tot, total=tot)
name("intangibles", ws, "C11")
name("goodwill", ws, "C13")
name("amort", ws, "C16")

# ---------------------------------------------------------------- Accretion / dilution
ws = sheet("Accretion", {"A": 2, "B": 52, "C": 4, "D": 13, "E": 13},
           "Pro forma earnings and accretion / (dilution)",
           "As if the merger were effective for the full year. INR crore unless stated.")
header(ws, 5, ["", "", "FY27E", "FY28E"])
S = "Standalone"
for col, scol, ph in (("D", "E", "phase27"), ("E", "F", "phase28")):
    put(ws, f"{col}6", f"={S}!{scol}{A['pat']}", NUM)
    put(ws, f"{col}7", f"={S}!{scol}{T['pat']}", NUM)
    put(ws, f"{col}8", f"={col}6+{col}7", NUM, bold=True, total=True)
    put(ws, f"{col}10", f"={ph}", PCT)
    put(ws, f"{col}11", f"=syn*{col}10", NUM)
    put(ws, f"{col}12", "=new_debt" if col == "D" else "=new_debt-repay_fy28", NUM)
    put(ws, f"{col}13", f"=-{col}12*cod", NUM)
    put(ws, f"{col}14", "=-amort", NUM)
    put(ws, f"{col}15", f"={col}11+{col}13+{col}14", NUM, total=True)
    put(ws, f"{col}16", f"=-{col}15*tax_adj", NUM)
    put(ws, f"{col}17", f"={col}15+{col}16", NUM, bold=True, total=True)
    put(ws, f"{col}19", f"={col}8+{col}17", NUM, bold=True, key=True)
    put(ws, f"{col}20", f"={S}!{scol}{A['shares']}", NUM2)
    put(ws, f"{col}21", "=new_shares", NUM2)
    put(ws, f"{col}22", f"={col}20+{col}21", NUM2, bold=True, total=True)
    put(ws, f"{col}24", f"={col}19/{col}22", NUM2, bold=True, key=True)
    put(ws, f"{col}25", f"={S}!{scol}{A['eps']}", NUM2)
    put(ws, f"{col}26", f"={col}24-{col}25", NUM2)
    put(ws, f"{col}27", f"={col}26/{col}25", PCT, bold=True, key=True)
    put(ws, f"{col}29", f"=-{col}14*(1-tax_adj)", NUM)
    put(ws, f"{col}30", f"=({col}19+{col}29)/{col}22", NUM2)
    put(ws, f"{col}31", f"={col}30/{col}25-1", PCT, bold=True, key=True)
    put(ws, f"{col}33", f"={col}11-({col}19-{col}25*{col}22)/(1-tax_adj)", NUM, bold=True)
    put(ws, f"{col}34", f"={col}33/syn_base", PCT)
for row, lab, b in [(6, "Torrent standalone profit after tax", False), (7, "JB standalone profit after tax", False),
                    (8, "Combined profit before deal adjustments", True), (10, "Synergy phasing", False),
                    (11, "Pre-tax synergies", False), (12, "Acquisition debt outstanding", False),
                    (13, "Interest on acquisition debt", False), (14, "Amortisation of acquired intangibles", False),
                    (15, "Pre-tax deal adjustments", False), (16, "Tax on adjustments", False),
                    (17, "Post-tax deal adjustments", True), (19, "Pro forma profit after tax", True),
                    (20, "Torrent existing shares (cr)", False), (21, "New shares issued to JB holders (cr)", False),
                    (22, "Pro forma shares (cr)", True), (24, "Pro forma EPS (INR)", True),
                    (25, "Torrent standalone EPS (INR)", False), (26, "Accretion / (dilution), INR", False),
                    (27, "Accretion / (dilution), %", True), (29, "Add back: post-tax intangible amortisation", False),
                    (30, "Pro forma cash EPS (INR)", False), (31, "Cash EPS accretion / (dilution), %", True),
                    (33, "Pre-tax synergies needed for EPS breakeven", True),
                    (34, "  as % of management run-rate target", False)]:
    label(ws, f"B{row}", lab, bold=b)
name("acc_fy27", ws, "D27")
name("acc_fy28", ws, "E27")

header(ws, 37, ["Leverage at closing", "", "INR cr", ""])
lev = [(38, "Torrent net debt before the deal", "=acq_net_debt", NUM),
       (39, "Less: JB net cash", "=-tgt_net_cash", NUM),
       (40, "New acquisition debt", "=new_debt", NUM),
       (41, "Pro forma net debt", "=SUM(D38:D40)", NUM),
       (42, "Combined EBITDA, FY26E", f"={S}!D{COMB['ebitda']}", NUM),
       (43, "Net debt / EBITDA, FY26E", "=D41/D42", MULT),
       (44, "Net debt / EBITDA, FY27E incl. synergies", f"=(D41-repay_fy28)/({S}!E{COMB['ebitda']}+D11)", MULT)]
for row, lab, f, fmt in lev:
    label(ws, f"B{row}", lab, bold=row in (41, 43))
    put(ws, f"D{row}", f, fmt, bold=row in (41, 43), total=row == 41)
name("lev_fy26", ws, "D43")

# ---------------------------------------------------------------- Sensitivity
ws = sheet("Sensitivity", {"A": 2, "B": 30, "C": 12, "D": 12, "E": 12, "F": 12, "G": 12, "H": 12},
           "Sensitivity analysis",
           "Excel data tables: every cell recalculates the whole model. Leave the six input cells below blank.")
header(ws, 5, ["Data-table input cells (leave blank)", "Value"])
for i, live in enumerate(FLEX):
    row = 6 + i
    label(ws, f"B{row}", LIVE_LABEL[live])
    c = ws[f"C{row}"]
    c.fill = FILL_IN
    name(f"ovr_{live}", ws, f"C{row}")
OVR = {live: f"$C${6 + i}" for i, live in enumerate(FLEX)}
FLEX_FMT = {"syn": NUM, "cod": PCT2, "oo_accept": PCT, "swap": MULT, "alloc": PCT, "life": "0"}

TABLES = []
r = 14
for key, t in SENS.items():
    label(ws, f"B{r}", t["title"], bold=True)
    top = r + 1
    put(ws, f"B{top}", "=acc_fy27" if t["year"] == "FY27E" else "=acc_fy28", PCT, bold=True, key=True)
    for j, v in enumerate(t["row_vals"]):
        c = ws.cell(row=top, column=3 + j, value=v)
        c.font, c.fill, c.number_format = F_H, FILL_H, FLEX_FMT[t["row_name"]]
    for i, v in enumerate(t["col_vals"]):
        c = ws.cell(row=top + 1 + i, column=2, value=v)
        c.font, c.fill, c.number_format = F_H, FILL_H, FLEX_FMT[t["col_name"]]
        c.alignment = Alignment(horizontal="right")
        for j in range(len(t["row_vals"])):
            ws.cell(row=top + 1 + i, column=3 + j).number_format = PCT
    last_col = "ABCDEFGHIJ"[2 + len(t["row_vals"]) - 1]
    rng = f"B{top}:{last_col}{top + len(t['col_vals'])}"
    TABLES.append((key, rng, OVR[t["row_name"]], OVR[t["col_name"]]))
    r = top + len(t["col_vals"]) + 3

# ---------------------------------------------------------------- Checks
ws = sheet("Checks", {"A": 2, "B": 60, "C": 12}, "Model checks", "Every line must read TRUE.")
header(ws, 5, ["Check", "Pass"])
checks = [
    ("Sources equal uses", "=ABS(Deal!F44)<0.001"),
    ("Cash and stock stakes cover 100% of JB", "=ABS(Deal!C12+Deal!C14-1)<0.000001"),
    ("Pro forma ownership sums to 100%", "=ABS(Deal!F49-1)<0.000001"),
    ("Purchase price allocation ties to consideration", "=ABS(PPA!C18)<0.001"),
    ("Goodwill is positive", "=goodwill>0"),
    ("New debt is not negative", "=new_debt>=0"),
    ("Open offer acceptance is between 0% and 100%", "=AND(oo_accept>=0,oo_accept<=1)"),
    ("Torrent FY25A EPS ties to reported 56.47", f"=ABS(Standalone!C{A['eps']}-56.47)<0.01"),
    ("JB FY25A EPS ties to reported 42.36", f"=ABS(Standalone!C{T['eps']}-42.36)<0.01"),
    ("Breakeven synergies give zero dilution, FY27E",
     "=ABS((Accretion!D19+(Accretion!D33-Accretion!D11)*(1-tax_adj))/Accretion!D22-Accretion!D25)<0.0001"),
    ("Sensitivity input cells are blank", '=COUNTA(Sensitivity!C6:C11)=0'),
]
for i, (lab, f) in enumerate(checks):
    label(ws, f"B{6 + i}", lab)
    put(ws, f"C{6 + i}", f)
end = 6 + len(checks)
label(ws, f"B{end + 1}", "All checks pass", bold=True)
put(ws, f"C{end + 1}", f"=AND(C6:C{end - 1})", bold=True, key=True)
name("all_ok", ws, f"C{end + 1}")

# ---------------------------------------------------------------- Summary
ws = ws_sum
ws.sheet_view.showGridLines = False
for col, w in {"A": 2, "B": 50, "C": 16, "D": 16}.items():
    ws.column_dimensions[col].width = w
ws["B2"] = "Torrent Pharmaceuticals / JB Chemicals & Pharmaceuticals"
ws["B2"].font = F_T
ws["B3"] = ("Merger model at announcement (29 June 2025). Market data at 27 June 2025 close. INR crore unless stated. "
            "Independent analysis from public information; not investment advice.")
ws["B3"].font = F_NOTE
header(ws, 5, ["Transaction", "", ""])
summ = [(6, "Total consideration for 100% of JB", "=consid", NUM),
        (7, "  of which cash", "=cash_total", NUM),
        (8, "  of which Torrent shares", "=stock_value", NUM),
        (9, "Implied enterprise value", "=implied_ev", NUM),
        (10, "EV / EBITDA, FY25A", "=Deal!F26", MULT),
        (11, "Blended premium / (discount) to pre-announcement price", "=Deal!F19", PCT),
        (12, "New acquisition debt", "=new_debt", NUM),
        (13, "New Torrent shares issued (cr)", "=new_shares", NUM2),
        (14, "JB public shareholders' stake in Torrent", "=Deal!F48", PCT),
        (15, "Net debt / EBITDA at closing, FY26E", "=lev_fy26", MULT),
        (16, "Goodwill created", "=goodwill", NUM)]
for row, lab, f, fmt in summ:
    label(ws, f"B{row}", lab)
    put(ws, f"C{row}", f, fmt)
header(ws, 18, ["Earnings impact on Torrent", "FY27E", "FY28E"])
imp = [(19, "Torrent standalone EPS (INR)", 25, NUM2), (20, "Pro forma EPS (INR)", 24, NUM2),
       (21, "EPS accretion / (dilution)", 27, PCT), (22, "Cash EPS accretion / (dilution)", 31, PCT),
       (23, "Pre-tax synergies needed for EPS breakeven", 33, NUM)]
for row, lab, src, fmt in imp:
    label(ws, f"B{row}", lab, bold=row == 21)
    put(ws, f"C{row}", f"=Accretion!D{src}", fmt, bold=row == 21, key=row == 21)
    put(ws, f"D{row}", f"=Accretion!E{src}", fmt, bold=row == 21, key=row == 21)
label(ws, "B25", "Verdict", bold=True)
put(ws, "C25", '=IF(acc_fy27>=0,"Accretive in FY27E",IF(acc_fy28>=0,"Dilutive in FY27E, accretive in FY28E",'
               '"Dilutive in FY27E and FY28E"))', bold=True)
label(ws, "B26", "Model checks", bold=True)
put(ws, "C26", '=IF(all_ok,"All pass","CHECK FAILED")', bold=True)
label(ws, "B28", "Colour code: blue on yellow = input; black = formula; shaded blue = key output.", note=True)
label(ws, "B29", "Sheets: Inputs, Standalone, Deal, PPA, Accretion, Sensitivity, Checks.", note=True)

OUT.parent.mkdir(exist_ok=True)
wb.save(OUT)
(Path(__file__).parent / "output" / "tables.txt").write_text(
    "\n".join(f"{k}|{rng}|{ri}|{ci}" for k, rng, ri, ci in TABLES), encoding="utf-8")
print("built", OUT)
for tbl in TABLES:
    print(tbl)
