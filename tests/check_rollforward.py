"""The facts behind a figure (bench/dcffacts.py) and last year's rows found in this year's changed model
(bench/rowfind.py), on tests/rollforward_pack.py: an overlay inside a copy of last year's client model, and
this year's model with rows inserted, renamed, restructured and a sheet renamed. No model calls.

  facts         from the report's figure down through the bridge, an INDEX-picked scenario and the mid of a
                low and a high rate to the two discountings (present-value rows), each tying exactly, with the
                rate, valuation date and period dates the factors use, and the client row the cash flows are
                read from
  finding       every row the label alone finds or misses, found by its history, words, neighbours or banner
  rolling       the figure on this year's model equals a calculation made here from the numbers
  standing in   a row this year's model doesn't have: last year's value stands in and is counted, not zero

    uv run python tests/check_rollforward.py
"""
import sys
import tempfile
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "bench"))
sys.path.insert(0, str(ROOT / "tests"))
import build_map  # noqa: E402
import dcffacts  # noqa: E402
import overlay as ov  # noqa: E402
import rollforward_pack as rp  # noqa: E402

FACTS = [{"id": 1, "category": "identity", "key": "valuation_date", "label": "Valuation date", "value_text": "30 June 2025",
          "value": 20250630.0, "unit": "date"},
         {"id": 2, "category": "assumption", "key": "discount_rate", "label": "Discount rate (low)", "value_text": "8.00%",
          "value": 8.0, "unit": "%"}]


def main() -> None:
    out = Path(tempfile.mkdtemp(prefix="rollforward_"))
    res = rp.build(out)
    db = {k: build_map.main(str(p), str(out / (p.stem + "_db")))["db"] for k, p in res["paths"].items()}
    w = {"overlay": {"db_path": db["overlay"], "filename": res["paths"]["overlay"].name,
                     "sheets": ["Inputs", "Val", "Bridge", "Report"], "source_path": str(res["paths"]["overlay"])},
         "prior": {"db_path": db["prior"], "filename": res["paths"]["prior"].name, "sheets": None,
                   "valuation_date": "2025-06-30"},
         "current": {"db_path": db["current"], "filename": res["paths"]["current"].name, "sheets": None,
                     "valuation_date": "2026-06-30"},
         "client_link": None, "prior_valuation_date": "2025-06-30", "same_file": False,
         "client_sheets": ["Ops", "CF", "Financing", "Summary"]}
    summary, sess = ov.build(out / "ovl", w["overlay"], w["prior"], w["current"], FACTS, "t", None, "2025-06-30",
                             client_sheets=w["client_sheets"])
    summary["wiring"] = w
    v = summary["validation"]
    assert v["matched"] == v["cells"], v["mismatches"][:3]

    # ---- facts
    fx = ov.deep(dcffacts.facts, sess, summary, "Report!C5")
    print(dcffacts.text(fx))
    f = res["figures"]
    assert abs(fx["value"] - f["ex_div"]) < 1e-9 and fx["python_equals_excel"]
    ds = fx["discountings"]
    assert sorted(round(c["pv"], 6) for c in ds) == sorted(round(x, 6) for x in (f["low"], f["high"])), [c["pv"] for c in ds]
    assert all(c["ties"] for c in ds)
    low = next(c for c in ds if abs(c["pv"] - f["low"]) < 1e-6)
    assert low["cashflow"]["row"] == "Val!r5" and low["cashflow"]["periods"] == 18
    assert low["cashflow"]["first_period"] == "2026-06-30" and low["cashflow"]["last_period"] == "2043-06-30"
    ins = {x["cell"]: x for x in low["factors"]["inputs"]}
    assert ins["Inputs!C5"]["what"] == "a rate" and ins["Inputs!C4"]["date"] == "2025-06-30", ins
    assert [o["row"] for o in low["origins"]] == ["CF!r11"], low["origins"]
    words = [p["words"] for p in low["path"]]
    assert any("INDEX" in (x or "") for x in words) and any("AVERAGE" in (x or "") for x in words), words
    print("facts: ok")

    # ---- finding last year's rows this year
    fnd = sess.rowmap
    want = {("CF", 5): ("CF", 5), ("CF", 6): ("CF", 8), ("CF", 7): ("CF", 9), ("CF", 8): ("CF", 11),
            ("CF", 9): ("CF", 12), ("CF", 11): ("CF", 15), ("Financing", 7): ("Debt", 7), ("Financing", 9): ("Debt", 9),
            ("Financing", 3): ("Debt", 3), ("Summary", 3): ("Summary", 3)}
    label_only = {k for k in want if fnd.rowmap.row(*k) == want[k][1] and k[0] == want[k][0]}
    for k, to in want.items():
        ex = fnd.explain(*k)
        assert ex["found"] == to, (k, ex)
    assert len(label_only) < len(want), "the fixture must need more than labels"
    assert "banner" in {n for n, _ in fnd.explain("CF", 11)["evidence"]}
    assert "words" in {n for n, _ in fnd.explain("CF", 8)["evidence"]}
    print(f"finding: ok ({len(want) - len(label_only)} of {len(want)} rows found only by history, words, neighbours or banner)")

    # ---- rolling forward: the figure on this year's model, against a calculation made here
    def this_year():
        defaults, _, months = ov._feed(summary, "current", None, None)
        sess.configure("current", defaults, months)
        return sess.values([("Report", 5, 3)])[0], dict(sess.stood_in)
    got, stood = ov.deep(this_year)
    n = res["current"]["numbers"]
    vd = date(2026, 6, 30)
    flows = [(date(fy, 6, 30), n["dist"][fy]) for fy in range(2027, 2045)]
    pv = lambda r: sum(x / (1 + r) ** ((d - vd).days / 365 - 0.5) for d, x in flows)
    expect = (pv(rp.LOW) + pv(rp.HIGH)) / 2 - rp.DIST_PAYABLE
    print(f"  this year: {got:.6f}, expected {expect:.6f}; stood in {len(stood)}")
    assert abs(got - expect) < 1e-6 and not stood
    print("rolling: ok")

    # ---- a row this year's model doesn't have: last year's value stands in, counted
    real = fnd.locate
    fnd.locate = lambda s, r: None if (s, r) == ("CF", 11) else real(s, r)
    try:
        got2, stood2 = ov.deep(this_year)
    finally:
        fnd.locate = real
    # 18 periods rolled on a year: 17 from last year's forecast for the same years; the last (FY2044) is past it,
    # so last year's own cell (the period it read) stands in: never a blank read as zero
    assert len(stood2) == 18 and all(k[:2] == ("CF", 11) for k in stood2), len(stood2)
    assert got2 > 0 and abs(got2 - expect) > 1e-3, got2  # last year's forecast, not zero
    print(f"standing in: ok (with the distributions missing: {got2:.3f}, from last year's forecast, not 0)")
    # ---- the roll: from last year's valuation date to this year's model's; periods move by whole periods
    def plan(ov_vd, prior_vd, cur_vd):
        r = ov.plan_roll(sess, w["prior"], w["overlay"], False, ov_vd, prior_vd, cur_vd)
        sess.shift = r["months"]
        return r, sess.period_shift(sess.prior, "CF"), sess.period_shift(sess.ov, "Val")
    r, cf, val = ov.deep(plan, "2025-09-30", "2025-06-30", "2025-12-31")  # the overlay dated after its client copy
    assert (r["months"], r["current_valuation_date"], cf, val) == (3, "2025-12-31", 0, 0), (r, cf, val)
    r, cf, val = ov.deep(plan, "2025-06-30", "2025-06-30", "2025-12-31")  # half a year: no year-end passed
    assert (r["months"], cf, val) == (6, 0, 0), (r, cf, val)
    r, cf, val = ov.deep(plan, "2025-06-30", "2025-06-30", "2026-06-30")  # a year: FY2026 has ended
    assert (r["months"], cf, val) == (12, 12, 12), (r, cf, val)
    summary["roll"].update(ov.deep(plan, "2025-09-30", "2025-06-30", "2025-12-31")[0])

    def three_months():
        defaults, _, months = ov._feed(summary, "current", None, None)
        sess.configure("current", defaults, months)
        sess.values([("Report", 5, 3)])
        return len(sess.client_reads), dict(sess.unmatched)
    reads, missed = ov.deep(three_months)
    assert reads and not missed, list(missed.items())[:3]
    print(f"roll: ok (Sep-25 overlay, Dec-25 model: 3 months, annual periods stay; {reads} client values read, none missed)")
    ov.deep(sess.configure, "workbook")
    roles_check(res, db)
    rebuilt_check(out)
    print("rollforward: all checks passed")


def rebuilt_check(out: Path) -> None:
    """A model rebuilt rather than revised: a sheet keeps its name but holds other line items, and last year's
    rows moved to other sheets. The same-named sheet isn't taken for the same sheet (no row found by its place
    there), and a moved row is found by its label and its history anywhere in the model."""
    import xlsxwriter
    import rowfind
    years = [date(2024 + k, 6, 30) for k in range(8)]

    def book(path, sheets):
        wb = xlsxwriter.Workbook(path)
        dt = wb.add_format({"num_format": "dd-mmm-yy"})
        for name, rows in sheets.items():
            ws = wb.add_worksheet(name)
            ws.write(2, 1, "Period ending")
            for k, d in enumerate(years):
                ws.write_datetime(2, 3 + k, d, dt)
            for i, (label, vals) in enumerate(rows):
                ws.write(4 + i, 1, label)
                for k, v in enumerate(vals):
                    ws.write_number(4 + i, 3 + k, v)
        wb.close()
        return build_map.main(str(path), str(out / (path.stem + "_db")))["db"]
    rev = [100.0 * 1.03 ** k for k in range(8)]
    fees = [7.0 + k for k in range(8)]
    prior = book(out / "rebuilt_prior.xlsx", {"Hub": [("Revenue", rev), ("Fees", fees)] + [(f"Driver {i}", [float(i)] * 8) for i in range(6)]})
    later = [v * (1.05 if k >= 2 else 1.0) for k, v in enumerate(rev)]  # history the same, forecast revised
    current = book(out / "rebuilt_current.xlsx", {"Hub": [(f"Scenario switch {i}", [1.0] * 8) for i in range(8)],
                                                  "Model": [("Opex", [30.0] * 8), ("Revenue", later), ("Fee income", fees)]})
    a, b = ov.Workbook(prior), ov.Workbook(current)
    f = rowfind.RowFinder(ov.RowMap(a, b), a, b)
    assert f.sheet_for("Hub") is None, "a sheet sharing no labels isn't the same sheet because of its name"
    ex = f.explain("Hub", 5)
    assert ex["found"] == ("Model", 6) and {"label", "history"} <= {n for n, _ in ex["evidence"]}, ex
    ex = f.explain("Hub", 6)
    assert ex["found"] == ("Model", 7), ex  # renamed as well as moved: its history finds it
    assert f.explain("Hub", 7)["found"] is None, "no row taken by its place on an unrelated sheet"
    print("rebuilt: ok (a same-named sheet with other contents isn't the same sheet; moved rows found by label and history)")


def roles_check(res: dict, db: dict) -> None:
    """Who's who, as the rules suggest it, where a report figure also sits on a client sheet (the net debt the
    client's model shows on its summary) and the client's model is dated before the overlay built on it."""
    import roles
    n = res["prior"]["numbers"]
    nd = n["net_debt"][2025]
    facts = FACTS + [{"id": 3, "category": "conclusion", "key": "equity_value", "label": "Fair value of equity (ex-div)",
                      "value_text": f"{res['figures']['ex_div']:.1f}", "value": round(res["figures"]["ex_div"], 1), "unit": "A$m"},
                     {"id": 4, "category": "assumption", "key": "net_debt", "label": "Net debt at valuation date",
                      "value_text": f"{nd:.1f}", "value": round(nd, 1), "unit": "A$m"}]
    wbs = [{"id": 1, "filename": res["paths"]["prior"].name, "db_path": db["prior"], "source_path": str(res["paths"]["prior"]),
            "valuation_date": "2025-06-30", "uploaded_at": 1},
           {"id": 2, "filename": res["paths"]["overlay"].name, "db_path": db["overlay"], "source_path": str(res["paths"]["overlay"]),
            "valuation_date": "2025-09-30", "uploaded_at": 2},
           {"id": 3, "filename": res["paths"]["current"].name, "db_path": db["current"], "source_path": str(res["paths"]["current"]),
            "valuation_date": "2026-06-30", "uploaded_at": 3}]
    got = roles.suggest([{"id": 9, "filename": "report.pdf", "n_facts": 4}], wbs, facts)
    r = {k: (v["id"], v.get("sheets")) for k, v in got["roles"].items()}
    print("  roles:", r)
    assert r["prior_overlay"] == (2, ["Inputs", "Val", "Bridge", "Report"]), r["prior_overlay"]
    assert r["prior_model"][0] == 1 and r["current_model"][0] == 3, r
    print("roles: ok (a report figure on a client sheet doesn't make it the overlay; different dates don't stop the "
          "client's own file being last year's model)")


if __name__ == "__main__":
    main()
