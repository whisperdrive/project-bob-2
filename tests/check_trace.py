"""The valuation tracer (bench/dcftrace.py) on a present-value row with quarters that have no cash flow.

  pv row    a figure = SUM of a present-value row (each column a cash flow times a factor), where some quarters have
            no cash flow: the factors are seen only where there is one (present value / cash flow), and the others
            aren't factors of 0. The rate, the valuation date and the convention are found from the ones seen, so
            the figure's discounting is known (the bridge's steps, the method selector, the DCF facts need it)
  rate      the discount rate three ways (overlay._rate_check): nothing said when the lever moves the rate the
            discountings use; the report's, the lever's and the discountings' rates side by side when it doesn't

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


def rate_check() -> None:
    import overlay as ov

    class Book:
        def labels(self):
            return {("Inputs", 5): "Discount rate", ("Inputs", 11): "Cost of equity", ("Inputs", 4): "Valuation date"}

        def value(self, s, r, c):
            return {("Inputs", 5, 3): 0.07, ("Inputs", 11, 5): 0.0815, ("Inputs", 4, 3): 45838.0}.get((s, r, c))

    class Sess:  # the rate cell E11 reads C5 when wired so; otherwise it's its own input
        def __init__(self, wired):
            self.ov, self.wired, self.over = Book(), wired, {}

        def configure(self, mode, overrides=None, *a):
            self.over = dict(overrides or {})

        def values(self, cells):
            lever = self.over.get(("Inputs", 5, 3))
            return [0.0115 + lever if self.wired and lever is not None else self.ov.value(*k) for k in cells]
    rows = [{"kind": "assumption", "key": "discount_rate", "report": "7.50%", "basis": "post-tax nominal WACC",
             "lever": {"cell": "Inputs!C5", "label": "Discount rate"}}]
    traces = {"Report!C5": (None, [{"inputs": {"rate": "Inputs!E11"}, "rate_value": 0.0815}])}
    assert ov._rate_check(Sess(True), None, rows, traces) is None
    got = ov._rate_check(Sess(False), None, rows, traces)
    assert got["lever"] == {"cell": "Inputs!C5", "label": "Discount rate", "value": 0.07}, got
    assert got["dcf"] == [{"cell": "Inputs!E11", "label": "Cost of equity", "value": 0.0815}]
    assert got["report"] == {"value": "7.50%", "basis": "post-tax nominal WACC"}
    print("rate: ok (a lever that moves the discountings' rate is left alone; one that doesn't is shown beside it and "
          "the report's, with their bases)")


if __name__ == "__main__":
    main()
    rate_check()
