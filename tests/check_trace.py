"""The valuation tracer (bench/dcftrace.py) on a present-value row with quarters that have no cash flow.

  pv row    a figure = SUM of a present-value row (each column a cash flow times a factor), where some quarters have
            no cash flow: the factors are seen only where there is one (present value / cash flow), and the others
            aren't factors of 0. The rate, the valuation date and the convention are found from the ones seen, so
            the figure's discounting is known (the bridge's steps, the method selector, the DCF facts need it)

    uv run python tests/check_trace.py
"""
import sqlite3
import sys
import tempfile
from datetime import date, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "bench"))
import xlsxwriter  # noqa: E402
from xlsxwriter.utility import xl_col_to_name as COL  # noqa: E402

import build_map  # noqa: E402
import dcftrace  # noqa: E402


def main() -> None:
    out = Path(tempfile.mkdtemp(prefix="trace_"))
    vd, rate, n = date(2025, 6, 30), 0.087, 24
    ends, y, m = [], 2025, 9
    for _ in range(n):
        ends.append(date(y + (m == 12), m % 12 + 1, 1) - timedelta(days=1))
        y, m = (y + 1, 3) if m == 12 else (y, m + 3)
    flows = [0.0 if k % 5 == 2 else 100.0 + 3 * k for k in range(n)]  # every fifth quarter has no cash flow
    df = [1 / (1 + rate) ** ((e - vd).days / 365) for e in ends]
    wb = xlsxwriter.Workbook(out / "pv.xlsx")
    dt = wb.add_format({"num_format": "dd-mmm-yy"})
    inp = wb.add_worksheet("Inputs")
    inp.write(3, 0, "Valuation date")
    inp.write_datetime(3, 2, vd, dt)
    inp.write(4, 0, "Discount rate")
    inp.write_number(4, 2, rate)
    fl = wb.add_worksheet("Flows")
    for r, label in ((2, "Period ending"), (9, "Cash flow"), (12, "Discount factor"), (13, "Present value"),
                     (15, "Equity value")):
        fl.write(r, 1, label)
    for k in range(n):
        c = COL(3 + k)
        fl.write_datetime(2, 3 + k, ends[k], dt)
        fl.write_number(9, 3 + k, flows[k])
        fl.write_formula(f"{c}13", f"=1/(1+Inputs!$C$5)^(({c}3-Inputs!$C$4)/365)", None, df[k])
        fl.write_formula(f"{c}14", f"={c}10*{c}13", None, flows[k] * df[k])
    fl.write_formula("C16", f"=SUM(D14:{COL(3 + n - 1)}14)", None, sum(f * x for f, x in zip(flows, df)))
    wb.close()
    db = sqlite3.connect(build_map.main(str(out / "pv.xlsx"), str(out / "db"))["db"])
    core = next(c for c in dcftrace.cores(dcftrace.trace(db, "Flows!C16")) if c.get("kind") == "pv row")
    m = core.get("method") or {}
    assert (m.get("rate"), m.get("valuation_date"), m.get("terminal_date")) == ("Inputs!C5", "Inputs!C4", None), m
    assert core.get("inputs", {}).get("cashflow") == ["Flows!D10:AA10"], core.get("inputs")
    assert core["periods"] == n - 5, core["periods"]
    print(f"pv row: ok (the discounting read from the {core['periods']} of {n} quarters with a cash flow: the rate "
          f"{m['rate']}, the valuation date {m['valuation_date']}, {m['timing']}, {m['day_count']}, no cut-off)")


if __name__ == "__main__":
    main()
