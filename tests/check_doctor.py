"""The overlay doctor (bench/doctor.py) on two broken variants of the synthetic engagement pack. No model calls.

  A  net debt read through an Excel table reference (Python can't read tables; the cell reads nothing that
     changes between years, so holding it at Excel's value is safe and clears the error), last year's client
     model a later version than the one the overlay last read, and this year's model with the free cash flow
     label on the tax row (the row followed into this year is the wrong one).
  B  a macro-style function in the discount factors (it reads the discount rate, so it can't be held), and a
     union formula Python can't compile.

    uv run python tests/check_doctor.py
"""
import sqlite3
import sys
import tempfile
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "bench"))
sys.path.insert(0, str(ROOT / "tests"))
import xlsxwriter  # noqa: E402

import build_map  # noqa: E402
import doctor  # noqa: E402
import extlinks  # noqa: E402
import make_engagement_pack as pack  # noqa: E402
import overlay as ov  # noqa: E402

VD, RATE, G, ND = date(2025, 6, 30), 0.0725, 0.025, 850.0
PRIOR_IN = dict(traffic0=42.0, growth=0.02, toll0=6.50, cpi=0.025, opex_pct=0.18, capex=35.0, major=120.0)
LATER_IN = {**PRIOR_IN, "capex": 36.0}
CUR_IN = dict(traffic0=43.1, growth=0.021, toll0=6.70, cpi=0.03, opex_pct=0.18, capex=38.0, major=120.0)
FACTS = [
    {"id": 1, "category": "conclusion", "key": "equity_value", "label": "Equity value", "value_text": "A$2,296.7m",
     "value": 2296.7, "unit": "A$m"},
    {"id": 2, "category": "conclusion", "key": "enterprise_value", "label": "Enterprise value", "value_text": "A$3,146.7m",
     "value": 3146.7, "unit": "A$m"},
    {"id": 3, "category": "assumption", "key": "discount_rate", "label": "Discount rate", "value_text": "7.25%",
     "value": 7.25, "unit": "%"},
    {"id": 4, "category": "assumption", "key": "terminal_growth_rate", "label": "Terminal growth rate",
     "value_text": "2.50%", "value": 2.5, "unit": "%"},
    {"id": 5, "category": "assumption", "key": "net_debt", "label": "Net debt", "value_text": "A$850.0m", "value": 850.0,
     "unit": "A$m"},
    {"id": 6, "category": "identity", "key": "valuation_date", "label": "Valuation date", "value_text": "30 June 2025",
     "value": 20250630.0, "unit": "date"},
]


def client(out: Path, name: str, n: dict, insurance: bool, start: date, inputs: dict) -> Path:
    p = out / name
    wb = xlsxwriter.Workbook(p)
    pack.write_client(wb, n, pack.inputs_rows(start, **inputs), insurance=insurance)
    wb.close()
    return p


def overlay_file(out: Path, name: str, prior: dict, faults) -> Path:
    p = out / name
    wb = xlsxwriter.Workbook(p)
    v = pack.write_overlay(wb, prior, lambda c: f"=[1]CashFlow!{c}9", VD, RATE, G, ND)
    faults(wb.get_worksheet_by_name("DCF"), v, prior)
    wb.close()
    cached = {("CashFlow", f"{pack.COL(pack.FIRST_COL + k)}9"): prior["fcf"][k] for k in range(pack.YEARS)}
    pack.add_external_link(p, "client_BP25.xlsx", ["Inputs", "Operations", "CashFlow"], cached)
    return p


def table_reference(d, v, prior):
    d.write_formula("D13", "=-NetDebtTbl[Amount]", None, -ND)


def macro_function(d, v, prior):
    for k in range(pack.YEARS):
        c = pack.COL(pack.FIRST_COL + k)
        d.write_formula(f"{c}9", f"=1/(1+DiscRate(Val_Inputs!$C$5))^YEARFRAC(Val_Inputs!$C$4,{c}$3,1)", None, v["df"][k])
    d.write_formula("F16", "=SUM((D5,E5))", None, prior["fcf"][0] + prior["fcf"][1])


def examine(out: Path, tag: str, ov_path: Path, prior_path: Path, cur_path: Path, swap_rows: bool = False) -> dict:
    db = {p: build_map.main(str(p), str(out / f"{p.stem}_db"))["db"] for p in (ov_path, prior_path, cur_path)}
    extlinks.ensure(str(ov_path), db[ov_path])
    if swap_rows:  # this year's model: the free cash flow label on the tax row, and the other way round
        with sqlite3.connect(db[cur_path]) as c:
            c.execute("UPDATE rows SET label='~' WHERE sheet='CashFlow' AND row=9")
            c.execute("UPDATE rows SET label='Unlevered free cash flow' WHERE sheet='CashFlow' AND row=8")
            c.execute("UPDATE rows SET label='Tax paid' WHERE sheet='CashFlow' AND label='~'")
    w = {"overlay": {"db_path": db[ov_path], "filename": ov_path.name, "sheets": ["Val_Inputs", "DCF", "Summary"],
                     "source_path": str(ov_path)},
         "prior": {"db_path": db[prior_path], "filename": prior_path.name, "sheets": None, "valuation_date": "2025-06-30"},
         "current": {"db_path": db[cur_path], "filename": cur_path.name, "sheets": None, "valuation_date": "2026-06-30"},
         "client_link": 1, "prior_valuation_date": "2025-06-30", "same_file": False, "client_sheets": None}
    summary, sess = ov.build(out / f"ovl_{tag}", w["overlay"], w["prior"], w["current"], FACTS, tag, 1, "2025-06-30")
    summary["wiring"] = w
    ev = ov.deep(doctor.examine, sess, summary)
    print(f"---- {tag}\n" + doctor.report_text({"evidence": ev}) + "\n")
    return ev


def causes(ev: dict, feed: str) -> dict:
    return {g["cause"]: g for g in ev["roots"].get(feed) or []}


def main() -> None:
    out = Path(tempfile.mkdtemp(prefix="doctor_"))
    prior = pack.client_numbers(2026, **PRIOR_IN, insurance=None)
    cur = pack.client_numbers(2027, **CUR_IN, insurance=4.0)
    p_prior = client(out, "client_BP25.xlsx", prior, False, date(2025, 7, 1), PRIOR_IN)
    p_later = client(out, "client_BP25_v2.xlsx", pack.client_numbers(2026, **LATER_IN, insurance=None), False,
                     date(2025, 7, 1), LATER_IN)
    p_cur = client(out, "client_BP26.xlsx", cur, True, date(2026, 7, 1), {**CUR_IN, "insurance": 4.0})

    a = examine(out, "A", overlay_file(out, "overlay_A.xlsx", prior, table_reference), p_later, p_cur, swap_rows=True)
    lay = {r["cell"]: r for r in a["layers"]}
    assert lay["DCF!D14"]["breaks_at"] == "workbook" and lay["DCF!D12"]["breaks_at"] == "prior", lay
    name = causes(a, "workbook")["name"]
    assert name["cells"] == ["DCF!D13"] and "Excel table" in name["detail"], name
    differs = causes(a, "prior")["client_differs"]
    assert differs["n_cells"] == pack.YEARS, differs
    version = next(f for f in a["files"] if f["check"].startswith("Last year's client model is the version"))
    assert version["status"] == "bad", version
    assert [x["status"] for x in a["rows"]["flagged"]] == ["sign"], a["rows"]
    assert [x["cell"] for x in a["holds"]["safe"]] == ["DCF!D13"], a["holds"]
    assert all(c["errors_after"] == 0 for c in a["holds"]["check"].values()), a["holds"]["check"]

    b = examine(out, "B", overlay_file(out, "overlay_B.xlsx", prior, macro_function), p_prior, p_cur)
    fn = causes(b, "workbook")["function"]
    assert fn["n_cells"] == pack.YEARS and "DISCRATE" in fn["detail"], fn
    assert not b["holds"]["safe"] and all(("Discount rate" in x["why"] or "timeline" in x["why"])
                                          for x in b["holds"]["unsafe"]), b["holds"]
    assert b["whole_model"]["not_compiled"] == 1, b["whole_model"]
    version = next(f for f in b["files"] if f["check"].startswith("Last year's client model is the version"))
    assert version["status"] == "ok", version
    assert not b["rows"]["flagged"], b["rows"]
    print("doctor: all checks passed")


if __name__ == "__main__":
    main()
