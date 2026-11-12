"""Independent check of the Excel model.

Recomputes the deal in plain Python (no spreadsheet formulas), then compares
against the values Excel calculated, including every data-table cell.
Run after build.py and finalize.ps1.
"""
import sys
from pathlib import Path

from openpyxl import load_workbook

from inputs import COMPANIES, DEAL, SENS

ROOT = Path(__file__).parent
BOOK = ROOT / "output" / "Torrent_JB_Merger_Model.xlsx"
TOL = 1e-6


def standalone(d):
    a = d["fy25"]
    rev, out = a["rev"], []
    pbt = a["ebitda"] - a["da"] + a["oi"] - a["int"]
    out.append({"rev": rev, "ebitda": a["ebitda"], "pat": pbt - a["tax"], "shares": a["shares"]})
    for j in range(3):
        rev *= 1 + d["growth"][j]
        ebitda = rev * d["margin"][j]
        pbt = ebitda - rev * d["da_pct"][j] + d["oi"][j] - d["int"][j]
        out.append({"rev": rev, "ebitda": ebitda, "pat": pbt * (1 - d["tax_rate"][j]), "shares": a["shares"]})
    for y in out:
        y["eps"] = y["pat"] / y["shares"]
    return out


def model(**over):
    i = {k: v[0] for k, v in DEAL.items()}
    live = {"syn": i["syn_base"], "cod": i["cod_base"], "oo_accept": i["oo_accept_base"],
            "swap": i["swap_base"], "alloc": i["alloc_base"], "life": i["life_base"]}
    live.update(over)
    acq, tgt = (standalone(d) for d in COMPANIES.values())
    fd = i["tgt_equity_fd"] / i["cash_price"]
    spa, emp, oo = fd * i["stake_spa"], fd * i["stake_emp"], fd * i["oo_size"] * live["oo_accept"]
    cash = (spa + emp) * i["cash_price"] + oo * i["oo_price"]
    new_shares = (fd - spa - emp - oo) * live["swap"]
    stock = new_shares * i["acq_price"]
    consid = cash + stock
    debt = cash + i["fees"] - i["cash_used"]
    excess = consid - i["tgt_book"]
    intang = excess * live["alloc"]
    goodwill = excess - intang + intang * i["tax_adj"]
    amort = intang / live["life"]
    shares = i["acq_shares"] + new_shares
    res = {"fd_shares": fd, "cash_total": cash, "stock_value": stock, "consid": consid, "new_debt": debt,
           "new_shares": new_shares, "pf_shares": shares, "intangibles": intang, "goodwill": goodwill,
           "amort": amort, "implied_ev": consid - i["tgt_net_cash"]}
    for yr, idx, phase, repay in (("27", 2, i["phase27"], 0.0), ("28", 3, i["phase28"], i["repay_fy28"])):
        pretax = live["syn"] * phase - (debt - repay) * live["cod"] - amort
        pat = acq[idx]["pat"] + tgt[idx]["pat"] + pretax * (1 - i["tax_adj"])
        eps = pat / shares
        res[f"pf_pat{yr}"] = pat
        res[f"pf_eps{yr}"] = eps
        res[f"sa_eps{yr}"] = acq[idx]["eps"]
        res[f"acc{yr}"] = eps / acq[idx]["eps"] - 1
        res[f"cash_acc{yr}"] = (pat + amort * (1 - i["tax_adj"])) / shares / acq[idx]["eps"] - 1
        res[f"breakeven{yr}"] = live["syn"] * phase - (pat - acq[idx]["eps"] * shares) / (1 - i["tax_adj"])
    net_debt = i["acq_net_debt"] - i["tgt_net_cash"] + debt
    res["lev26"] = net_debt / (acq[1]["ebitda"] + tgt[1]["ebitda"])
    return res


def main():
    wb = load_workbook(BOOK, data_only=True)

    def named(n):
        sheet, cell = wb.defined_names[n].attr_text.replace("'", "").replace("$", "").split("!")
        return wb[sheet][cell].value

    base = model()
    acc = wb["Accretion"]
    excel = {k: named(k) for k in ("fd_shares", "cash_total", "stock_value", "consid", "new_debt", "new_shares",
                                   "pf_shares", "intangibles", "goodwill", "amort", "implied_ev")}
    summary = wb["Summary"]
    excel.update({"acc27": named("acc_fy27"), "acc28": named("acc_fy28"), "lev26": named("lev_fy26"),
                  "summary_acc27": summary["C21"].value, "summary_acc28": summary["D21"].value,
                  "summary_lev26": summary["C15"].value})
    base.update({"summary_acc27": base["acc27"], "summary_acc28": base["acc28"], "summary_lev26": base["lev26"]})
    excel.update({"pf_pat27": acc["D19"].value, "pf_pat28": acc["E19"].value,
                  "pf_eps27": acc["D24"].value, "pf_eps28": acc["E24"].value,
                  "sa_eps27": acc["D25"].value, "sa_eps28": acc["E25"].value,
                  "cash_acc27": acc["D31"].value, "cash_acc28": acc["E31"].value,
                  "breakeven27": acc["D33"].value, "breakeven28": acc["E33"].value})

    lines, fails = [], 0
    for k in sorted(excel):
        ok = excel[k] is not None and abs(excel[k] - base[k]) <= TOL * max(1.0, abs(base[k]))
        fails += not ok
        lines.append(f"| {k} | {base[k]:,.6f} | {excel[k] if excel[k] is None else format(excel[k], ',.6f')} | "
                     f"{'pass' if ok else 'FAIL'} |")

    sens = wb["Sensitivity"]
    grid_lines, cells = [], 0
    for key, rng, _, _ in (l.split("|") for l in (ROOT / "output" / "tables.txt").read_text().splitlines()):
        t = SENS[key]
        top = sens[rng.split(":")[0]].row
        worst = 0.0
        for a, cv in enumerate(t["col_vals"]):
            for b, rv in enumerate(t["row_vals"]):
                want = model(**{t["col_name"]: cv, t["row_name"]: rv})["acc27" if t["year"] == "FY27E" else "acc28"]
                got = sens.cell(row=top + 1 + a, column=3 + b).value
                diff = abs((got if got is not None else 9e9) - want)
                worst = max(worst, diff)
                cells += 1
                fails += diff > TOL
        grid_lines.append(f"| {key} | {len(t['col_vals']) * len(t['row_vals'])} | {worst:.2e} | "
                          f"{'pass' if worst <= TOL else 'FAIL'} |")

    checks = wb["Checks"]
    check_rows = [(checks[f"B{r}"].value, checks[f"C{r}"].value) for r in range(6, 40) if checks[f"B{r}"].value]
    fails += sum(1 for _, v in check_rows if v is not True)

    errors = sum(1 for ws in wb for row in ws.iter_rows() for c in row
                 if isinstance(c.value, str) and c.value.startswith("#"))
    fails += errors

    report = ["# Validation", "",
              f"Result: **{'ALL PASS' if not fails else str(fails) + ' FAILURE(S)'}**. "
              f"{len(excel)} headline outputs, {cells} data-table cells and {len(check_rows)} in-workbook checks "
              f"compared; {errors} Excel error cells.", "",
              "The reference values come from `verify.py`, a separate plain-Python calculation that shares only the "
              "input values with the workbook.", "",
              "## Headline outputs", "", "| Output | Independent | Excel | Result |", "|---|---:|---:|---|", *lines, "",
              "## Data tables (each cell recomputed by a full independent run)", "",
              "| Table | Cells | Largest difference | Result |", "|---|---:|---:|---|", *grid_lines, "",
              "## In-workbook checks", "", "| Check | Excel |", "|---|---|",
              *[f"| {b} | {v} |" for b, v in check_rows], ""]
    (ROOT / "reports").mkdir(exist_ok=True)
    (ROOT / "reports" / "validation.md").write_text("\n".join(report), encoding="utf-8")
    print("\n".join(report))
    sys.exit(1 if fails else 0)


if __name__ == "__main__":
    main()
