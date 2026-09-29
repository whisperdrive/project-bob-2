"""A client model in two versions and an overlay inside a copy of last year's, as recurring valuations have them,
for the DCF facts (bench/dcffacts.py) and for finding last year's rows in this year's model (bench/rowfind.py).

Last year's client model (FY2024-FY2043, June years; FY2024-FY2025 actual, valued at 30 June 2025):
  Ops        Revenue, Operating costs, EBITDA
  CF         EBITDA, Capital expenditure, Tax paid, Free cash flow, Debt service, Distributions to equity
  Financing  Opening debt, Interest, Debt service, Net debt
  Summary    a banner in its first rows (the forecast distributions' total, net debt at the valuation date),
             repeated in CF's first row
The overlay (sheets Inputs, Val, Bridge, Report, in a copy of that workbook) reads the distributions from CF:
  Val     the cash flow to equity (from CF, other columns), mid-period factors at a low and a high rate, a
          present-value row for each, their sums, the mid as their average, and the scenario picked by INDEX
  Bridge  equity value cum-div = the selected scenario; ex-div = cum-div less the distribution payable
  Report  the report's figure
This year's client model (FY2025-FY2044; FY2025 actual as last year, FY2026 now actual, forecast revised):
  changes as models change between valuations: rows inserted in CF (Capital expenditure moves down), Tax paid
  renamed "Income tax paid", Distributions to equity renamed "Equity distributions" and moved below a new
  reserve row, Free cash flow restructured (working capital added, renamed "Free cash flow after working
  capital"), and the Financing sheet renamed "Debt"
Every formula's value is saved as Excel would, so the models read as saved workbooks do.
"""
from datetime import date
from pathlib import Path

import xlsxwriter
from xlsxwriter.utility import xl_col_to_name as COL

C0 = 4          # column E: the first period on the client sheets
N = 20
LOW, HIGH, DIST_PAYABLE = 0.08, 0.09, 12.5
OV0 = 3         # column D: the first period on the overlay's Val sheet


def serial(d: date) -> int:
    return (d - date(1899, 12, 30)).days


def client_numbers(fy0: int, revision: float = 0.0) -> dict:
    """Every client row by financial year (fy0 .. fy0+N-1). Actual years are the same in both versions (they're
    history); later years move by revision."""
    out = {k: {} for k in ("revenue", "opex", "ebitda", "capex", "tax", "fcf", "wc", "debt_open", "interest",
                           "debt_service", "net_debt", "dist", "reserve")}
    for i, fy in enumerate(range(fy0, fy0 + N)):
        k = fy - 2024
        grow = 1.0 if fy <= 2026 else 1 + revision
        rev = 100 * 1.03 ** k * grow
        opex = 0.35 * rev
        ebitda = rev - opex
        capex = -(10 + k)
        tax = -max(0.0, ebitda + capex) * 0.3
        wc = -0.5 * (1 + k % 3) if revision and fy > 2026 else 0.0  # history doesn't change between versions
        fcf = ebitda + capex + tax + wc
        debt_open = 300 - 15 * k
        interest = debt_open * 0.05
        ds = interest + 15
        net_debt = debt_open - 15
        reserve = -2.0 if revision and fy > 2026 else 0.0
        dist = fcf - ds + reserve
        for key, v in (("revenue", rev), ("opex", opex), ("ebitda", ebitda), ("capex", capex), ("tax", tax),
                       ("fcf", fcf), ("wc", wc), ("debt_open", debt_open), ("interest", interest),
                       ("debt_service", ds), ("net_debt", net_debt), ("dist", dist), ("reserve", reserve)):
            out[key][fy] = v
    return out


def write_client(wb, fy0: int, this_year: bool, banner: bool = True) -> dict:
    """The client sheets. Returns where things are: {"cf_dist_row", "cf_date_row", "years", "numbers"}."""
    n = client_numbers(fy0, 0.04 if this_year else 0.0)
    years = list(range(fy0, fy0 + N))
    num, dt = wb.add_format({"num_format": "#,##0.0"}), wb.add_format({"num_format": "dd-mmm-yy"})
    col = lambda i: COL(C0 + i)

    def timeline(ws):
        ws.write(2, 1, "Period ending")
        for i, fy in enumerate(years):
            ws.write_datetime(2, C0 + i, date(fy, 6, 30), dt)
            ws.write(3, C0 + i, "Actual" if fy <= fy0 + 1 else "Forecast")

    ops = wb.add_worksheet("Ops")
    ops.write(0, 1, "Operations")
    timeline(ops)
    for r, lab in ((5, "Revenue"), (6, "Operating costs"), (7, "EBITDA")):
        ops.write(r - 1, 1, lab)
    for i, fy in enumerate(years):
        c = col(i)
        ops.write_number(4, C0 + i, n["revenue"][fy], num)
        ops.write_formula(f"{c}6", f"={c}5*0.35", num, n["opex"][fy])
        ops.write_formula(f"{c}7", f"={c}5-{c}6", num, n["ebitda"][fy])

    fin_name = "Debt" if this_year else "Financing"
    cf = wb.add_worksheet("CF")
    timeline(cf)
    if this_year:  # two rows inserted above the capital expenditure; tax renamed; FCF restructured; distributions moved
        rows = {"ebitda": 5, "maint": 6, "growth": 7, "capex": 8, "tax": 9, "wc": 10, "fcf": 11, "ds": 12, "reserve": 13, "dist": 15}
        labels = {"ebitda": "EBITDA", "maint": "Maintenance programme", "growth": "Growth programme",
                  "capex": "Capital expenditure", "tax": "Income tax paid", "wc": "Working capital movement",
                  "fcf": "Free cash flow after working capital", "ds": "Debt service", "reserve": "Reserve top-up",
                  "dist": "Equity distributions"}
    else:
        rows = {"ebitda": 5, "capex": 6, "tax": 7, "fcf": 8, "ds": 9, "dist": 11}
        labels = {"ebitda": "EBITDA", "capex": "Capital expenditure", "tax": "Tax paid", "fcf": "Free cash flow",
                  "ds": "Debt service", "dist": "Distributions to equity"}
    for key, r in rows.items():
        cf.write(r - 1, 1, labels[key])
    R = rows
    for i, fy in enumerate(years):
        c = col(i)
        cf.write_formula(f"{c}{R['ebitda']}", f"=Ops!{c}7", num, n["ebitda"][fy])
        if this_year:
            cf.write_number(f"{c}{R['maint']}", n["capex"][fy] * 0.6, num)
            cf.write_number(f"{c}{R['growth']}", n["capex"][fy] * 0.4, num)
            cf.write_formula(f"{c}{R['capex']}", f"={c}{R['maint']}+{c}{R['growth']}", num, n["capex"][fy])
        else:
            cf.write_number(f"{c}{R['capex']}", n["capex"][fy], num)
        cf.write_formula(f"{c}{R['tax']}", f"=-MAX(0,{c}{R['ebitda']}+{c}{R['capex']})*0.3", num, n["tax"][fy])
        if this_year:
            cf.write_number(f"{c}{R['wc']}", n["wc"][fy], num)
            cf.write_formula(f"{c}{R['fcf']}", f"={c}{R['ebitda']}+{c}{R['capex']}+{c}{R['tax']}+{c}{R['wc']}", num, n["fcf"][fy])
            cf.write_number(f"{c}{R['reserve']}", n["reserve"][fy], num)
            cf.write_formula(f"{c}{R['dist']}", f"={c}{R['fcf']}-{c}{R['ds']}+{c}{R['reserve']}", num, n["dist"][fy])
        else:
            cf.write_formula(f"{c}{R['fcf']}", f"={c}{R['ebitda']}+{c}{R['capex']}+{c}{R['tax']}", num, n["fcf"][fy])
            cf.write_formula(f"{c}{R['dist']}", f"={c}{R['fcf']}-{c}{R['ds']}", num, n["dist"][fy])
        cf.write_formula(f"{c}{R['ds']}", f"={fin_name}!{c}7", num, n["debt_service"][fy])

    fin = wb.add_worksheet(fin_name)
    timeline(fin)
    for r, lab in ((5, "Opening debt"), (6, "Interest"), (7, "Debt service"), (9, "Net debt")):
        fin.write(r - 1, 1, lab)
    for i, fy in enumerate(years):
        c = col(i)
        fin.write_number(4, C0 + i, n["debt_open"][fy], num)
        fin.write_formula(f"{c}6", f"={c}5*0.05", num, n["interest"][fy])
        fin.write_formula(f"{c}7", f"={c}6+15", num, n["debt_service"][fy])
        fin.write_formula(f"{c}9", f"={c}5-15", num, n["net_debt"][fy])

    # the banner: the forecast distributions' total and net debt at the valuation date, on Summary and CF's row 1
    fc = [i for i, fy in enumerate(years) if fy > fy0 + 1]
    total = sum(n["dist"][years[i]] for i in fc)
    vd_col = col(1)
    if not banner:  # (an adviser who reused the name for the valuation's own summary)
        return {"years": years, "numbers": n, "rows": R, "fin": fin_name}
    sm = wb.add_worksheet("Summary")
    sm.write(0, 1, "Model summary")
    sm.write(2, 1, "Equity distributions (forecast total)")
    sm.write_formula("C3", f"=SUM(CF!{col(fc[0])}{R['dist']}:{col(fc[-1])}{R['dist']})", num, total)
    sm.write(3, 1, "Net debt at valuation date")
    sm.write_formula("C4", f"={fin_name}!{vd_col}9", num, n["net_debt"][years[1]])
    cf.write(0, 1, "Equity distributions (forecast total)")
    cf.write_formula("C1", "=Summary!C3", num, total)
    return {"years": years, "numbers": n, "rows": R, "fin": fin_name}


def write_overlay(wb, made: dict, vd: date, report_sheet: str = "Report") -> dict:
    """The overlay's sheets, reading the client sheets' distributions for the forecast years."""
    n, years, R = made["numbers"], made["years"], made["rows"]
    num, dt, pct = wb.add_format({"num_format": "#,##0.0"}), wb.add_format({"num_format": "dd-mmm-yy"}), \
        wb.add_format({"num_format": "0.00%"})
    fc = [i for i, fy in enumerate(years) if date(fy, 6, 30) > vd]
    inp = wb.add_worksheet("Inputs")
    for r, lab, v, f in ((4, "Valuation date", vd, dt), (5, "Discount rate (low)", LOW, pct), (6, "Discount rate (high)", HIGH, pct),
                         (8, "Scenario", "Mid", None), (9, "Distribution payable", DIST_PAYABLE, num)):
        inp.write(r - 1, 1, lab)
        (inp.write_datetime if isinstance(v, date) else inp.write)(r - 1, 2, v, f)
    val = wb.add_worksheet("Val")
    for r, lab in ((3, "Period ending"), (5, "Cash flow to equity"), (7, "Discount factor (low)"), (8, "Present value (low)"),
                   (10, "Discount factor (high)"), (11, "Present value (high)"), (13, "PV (low)"), (14, "PV (high)"),
                   (15, "Mid value"), (17, "Low"), (18, "Mid"), (19, "High"), (21, "Selected valuation")):
        val.write(r - 1, 1, lab)
    pv = {LOW: 0.0, HIGH: 0.0}
    for j, i in enumerate(fc):
        c, src = COL(OV0 + j), COL(C0 + i)
        end = date(years[i], 6, 30)
        t = (end - vd).days / 365 - 0.5
        flow = n["dist"][years[i]]
        val.write_formula(f"{c}3", f"=CF!{src}3", dt, serial(end))
        val.write_formula(f"{c}5", f"=CF!{src}{R['dist']}", num, flow)
        for rate, dr, pr in ((LOW, 7, 8), (HIGH, 10, 11)):
            cell = "$C$5" if rate == LOW else "$C$6"
            f = 1 / (1 + rate) ** t
            val.write_formula(f"{c}{dr}", f"=1/(1+Inputs!{cell})^(({c}$3-Inputs!$C$4)/365-0.5)", None, f)
            val.write_formula(f"{c}{pr}", f"={c}5*{c}{dr}", num, flow * f)
            pv[rate] += flow * f
    last = COL(OV0 + len(fc) - 1)
    mid = (pv[LOW] + pv[HIGH]) / 2
    val.write_formula("C13", f"=SUM(D8:{last}8)", num, pv[LOW])
    val.write_formula("C14", f"=SUM(D11:{last}11)", num, pv[HIGH])
    val.write_formula("C15", "=AVERAGE(C13:C14)", num, mid)
    val.write_formula("C17", "=C13", num, pv[LOW])
    val.write_formula("C18", "=C15", num, mid)
    val.write_formula("C19", "=C14", num, pv[HIGH])
    val.write_formula("C21", "=INDEX(C17:C19,MATCH(Inputs!C8,B17:B19,0))", num, mid)
    br = wb.add_worksheet("Bridge")
    for r, lab, f, v in ((5, "Equity value (cum-div)", "=Val!C21", mid),
                         (6, "Less: distribution payable", "=-Inputs!C9", -DIST_PAYABLE),
                         (7, "Equity value (ex-div)", "=C5+C6", mid - DIST_PAYABLE)):
        br.write(r - 1, 1, lab)
        br.write_formula(f"C{r}", f, num, v)
    rp = wb.add_worksheet(report_sheet)
    rp.write(4, 1, "Fair value of equity (ex-div)")
    rp.write_formula("C5", "=Bridge!C7", num, mid - DIST_PAYABLE)
    return {"low": pv[LOW], "high": pv[HIGH], "mid": mid, "ex_div": mid - DIST_PAYABLE}


def build(out: Path) -> dict:
    """The three workbooks: last year's client model, the overlay inside a copy of it, this year's client model."""
    out.mkdir(parents=True, exist_ok=True)
    paths = {"prior": out / "Client_model_FY25.xlsx", "overlay": out / "Client_model_FY25_with_valuation.xlsx",
             "current": out / "Client_model_FY26.xlsx"}
    wb = xlsxwriter.Workbook(paths["prior"])
    write_client(wb, 2024, False)
    wb.close()
    wb = xlsxwriter.Workbook(paths["overlay"])
    made = write_client(wb, 2024, False)
    figures = write_overlay(wb, made, date(2025, 6, 30))
    wb.close()
    wb = xlsxwriter.Workbook(paths["current"])
    now = write_client(wb, 2025, True)
    wb.close()
    return {"paths": paths, "figures": figures, "prior": made, "current": now}


if __name__ == "__main__":
    import sys
    res = build(Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).resolve().parent / "rollforward_pack")
    print({k: str(v) for k, v in res["paths"].items()}, {k: round(v, 3) for k, v in res["figures"].items()})
