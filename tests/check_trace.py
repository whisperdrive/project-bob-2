"""The valuation tracer (bench/dcftrace.py) on a present-value row with quarters that have no cash flow.

  pv row    a figure = SUM of a present-value row (each column a cash flow times a factor), where some quarters have
            no cash flow: the factors are seen only where there is one (present value / cash flow), and the others
            aren't factors of 0. The rate, the valuation date and the convention are found from the ones seen, so
            the figure's discounting is known (the bridge's steps, the method selector, the DCF facts need it)
  mid-year  factors written (end - valuation date) / 365 - 0.5, the mid-year convention as models often have it,
            are read back as such
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
    assert got["moves"] == [] and "range" not in got, "on other rows: not a range, and the lever moves nothing"

    class Row:  # a low / mid / high on one row; the figure averages the values at the low and the high rate
        def labels(self):
            return {("Inputs", 11): "Discount rate"}

        def value(self, s, r, c):
            return {("Inputs", 11, 5): 0.091, ("Inputs", 11, 6): 0.086, ("Inputs", 11, 7): 0.081}.get((s, r, c))

    class RowSess(Sess):
        def values(self, cells):  # each rate cell is its own input: the lever moves only itself
            return [self.over.get(k, self.ov.value(*k)) for k in cells]
    sess = RowSess(False)
    sess.ov = Row()
    lever_on_end = [{**rows[0], "report": "8.60%", "lever": {"cell": "Inputs!G11", "label": "Discount rate"}}]
    two = {"Report!C5": (None, [{"inputs": {"rate": "Inputs!E11"}, "rate_value": 0.091},
                                {"inputs": {"rate": "Inputs!G11"}, "rate_value": 0.081}])}
    got = ov._rate_check(sess, None, lever_on_end, two)
    assert got["moves"] == ["Inputs!G11"] and got["report_at"] == "Inputs!F11", got
    assert [c["cell"] for c in got["range"]] == ["Inputs!E11", "Inputs!F11", "Inputs!G11"], got["range"]
    print("rate: ok (a lever that moves the discountings' rate is left alone; one that doesn't is shown beside it and "
          "the report's, with their bases; a low / mid / high on one row is a range, the lever one end of it)")


def mid_year_check() -> None:
    out = Path(tempfile.mkdtemp(prefix="trace_mid_"))
    vd, rate, years = date(2025, 6, 30), 0.08, list(range(2026, 2036))
    wb = xlsxwriter.Workbook(out / "mid.xlsx")
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
    total = 0.0
    for k, y in enumerate(years):
        c, end = COL(3 + k), date(y, 6, 30)
        f = 1 / (1 + rate) ** ((end - vd).days / 365 - 0.5)
        fl.write_datetime(2, 3 + k, end, dt)
        fl.write_number(9, 3 + k, 50.0 + k)
        fl.write_formula(f"{c}13", f"=1/(1+Inputs!$C$5)^(({c}3-Inputs!$C$4)/365-0.5)", None, f)
        fl.write_formula(f"{c}14", f"={c}10*{c}13", None, (50.0 + k) * f)
        total += (50.0 + k) * f
    fl.write_formula("C16", f"=SUM(D14:{COL(3 + len(years) - 1)}14)", None, total)
    wb.close()
    db = sqlite3.connect(build_map.main(str(out / "mid.xlsx"), str(out / "db"))["db"])
    core = next(c for c in dcftrace.cores(dcftrace.trace(db, "Flows!C16")) if c.get("kind") == "pv row")
    m = core.get("method") or {}
    assert (m.get("timing"), m.get("day_count"), m.get("rate")) == ("mid-year", "actual/365", "Inputs!C5"), m
    print("mid-year: ok (factors written (end - valuation date) / 365 - 0.5 read back as the mid-year convention)")


if __name__ == "__main__":
    main()
    mid_year_check()
    rate_check()
