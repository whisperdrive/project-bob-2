"""Write a small synthetic overlay for the valuation tracer (bench/dcftrace.py) to tests/engagement_pack/.

The shapes an overlay's valuation takes that the label-based DCF finder misses, with every formula's value saved
as Excel would:
  Flows      quarterly equity cash flows built from parts (trust distributions + company dividends, plus the
             franking credits used = gross credits x utilisation), a discount factor row and a PV row
  Valuation  PV at a low and a high rate by XNPV, the mid value as their average, equity value cum-div = the
             mid, ex-div = cum-div less the distribution payable; and checks computing the low-rate PV three
             other ways: SUMPRODUCT with the factor row, SUMPRODUCT with an inline factor expression, and the
             SUM of the PV row, plus an NPV
    uv run python tests/make_trace_workbook.py
"""
import sys
from datetime import date, timedelta
from pathlib import Path

import xlsxwriter
from xlsxwriter.utility import xl_col_to_name as COL

OUT = Path(__file__).resolve().parent / "engagement_pack"
N = 40          # quarters
FIRST = 3       # column D
VD = date(2025, 6, 30)
LOW, HIGH, UTIL, DIST = 0.077, 0.087, 0.8, 13.1


def quarter_ends(n: int) -> list[date]:
    out, y, m = [], 2025, 9
    for _ in range(n):
        nxt = date(y + (m == 12), m % 12 + 1, 1)
        out.append(nxt - timedelta(days=1))
        y, m = (y + 1, 3) if m == 12 else (y, m + 3)
    return out


def serial(d: date) -> int:
    return (d - date(1899, 12, 30)).days


def xnpv(rate: float, values: list[float], dates: list[date]) -> float:
    return sum(v / (1 + rate) ** ((d - dates[0]).days / 365) for v, d in zip(values, dates))


def build(path: Path) -> dict:
    ends = quarter_ends(N)
    trust = [180 + 3 * k for k in range(N)]
    company = [40 + k for k in range(N)]
    credits = [17 + 0.4 * k for k in range(N)]
    used = [c * UTIL for c in credits]
    excl = [t + c for t, c in zip(trust, company)]
    flow = [e + u for e, u in zip(excl, used)]
    df = [1 / (1 + LOW) ** ((d - VD).days / 365) for d in ends]
    pv_row = [f * x for f, x in zip(flow, df)]
    low = xnpv(LOW, [0.0] + flow, [VD] + ends)
    high = xnpv(HIGH, [0.0] + flow, [VD] + ends)
    mid = (low + high) / 2
    npv = sum(f / (1 + LOW) ** (k + 1) for k, f in enumerate(flow))

    wb = xlsxwriter.Workbook(str(path))
    num, dt = wb.add_format({"num_format": "#,##0.0"}), wb.add_format({"num_format": "yyyy-mm-dd"})
    pct = wb.add_format({"num_format": "0.00%"})
    last = COL(FIRST + N - 1)

    i = wb.add_worksheet("Inputs")
    i.write("B2", "Valuation inputs")
    for r, (lab, v, f) in enumerate([("Valuation date", VD, dt), ("Discount rate (low)", LOW, pct),
                                     ("Discount rate (high)", HIGH, pct), ("Franking credit utilisation", UTIL, pct),
                                     ("Distribution payable", DIST, num)], start=3):
        i.write(r, 1, lab)
        (i.write_datetime if isinstance(v, date) else i.write)(r, 2, v, f)

    fl = wb.add_worksheet("Flows")
    fl.write("B3", "Period ending")
    fl.write("B11", "XNPV dates")
    fl.write_formula("C11", "=Inputs!C4", dt, serial(VD))
    fl.write("C10", 0, num)
    labels = {5: "Trust distributions", 6: "Company dividends", 7: "Franking credits (gross)", 8: "Franking credits used",
              9: "Equity cash flow excluding credits", 10: "Equity cash flow", 13: "Discount factor (low)",
              14: "Present value (low)"}
    for r, lab in labels.items():
        fl.write(r - 1, 1, lab)
    for k in range(N):
        c = COL(FIRST + k)
        fl.write_datetime(2, FIRST + k, ends[k], dt)
        fl.write(4, FIRST + k, trust[k], num)
        fl.write(5, FIRST + k, company[k], num)
        fl.write(6, FIRST + k, credits[k], num)
        fl.write_formula(f"{c}8", f"={c}7*Inputs!$C$7", num, used[k])
        fl.write_formula(f"{c}9", f"={c}5+{c}6", num, excl[k])
        fl.write_formula(f"{c}10", f"={c}9+{c}8", num, flow[k])
        fl.write_formula(f"{c}11", f"={c}3", dt, serial(ends[k]))
        fl.write_formula(f"{c}13", f"=1/(1+Inputs!$C$5)^(({c}3-Inputs!$C$4)/365)", None, df[k])
        fl.write_formula(f"{c}14", f"={c}10*{c}13", num, pv_row[k])

    v = wb.add_worksheet("Valuation")
    rows = [(10, "PV at low rate", f"=XNPV(Inputs!C5,Flows!C10:{last}10,Flows!C11:{last}11)", low),
            (11, "PV at high rate", f"=XNPV(Inputs!C6,Flows!C10:{last}10,Flows!C11:{last}11)", high),
            (12, "Mid value", "=AVERAGE(F10:F11)", mid),
            (14, "Equity value (cum-div)", "=F12", mid),
            (15, "Distribution payable", "=Inputs!C8", DIST),
            (16, "Equity value (ex-div)", "=F14-F15", mid - DIST),
            (20, "Check: PV low, factor row", f"=SUMPRODUCT(Flows!D10:{last}10,Flows!D13:{last}13)", sum(pv_row)),
            (21, "Check: PV low, inline factors",
             f"=SUMPRODUCT(Flows!D10:{last}10,1/(1+Inputs!$C$5)^((Flows!D3:{last}3-Inputs!$C$4)/365))", sum(pv_row)),
            (22, "Check: PV low, PV row", f"=SUM(Flows!D14:{last}14)", sum(pv_row)),
            (23, "Check: NPV", f"=NPV(Inputs!C5,Flows!D10:{last}10)", npv)]
    for r, lab, f, val in rows:
        v.write(r - 1, 1, lab)
        v.write_formula(f"F{r}", f, num, val)
    wb.close()
    return {"low": low, "high": high, "mid": mid, "ex_div": mid - DIST, "pv_row": sum(pv_row), "npv": npv}


if __name__ == "__main__":
    OUT.mkdir(parents=True, exist_ok=True)
    out = build(OUT / "trace_overlay.xlsx")
    print({k: round(x, 4) for k, x in out.items()}, file=sys.stderr)
    print(OUT / "trace_overlay.xlsx")
