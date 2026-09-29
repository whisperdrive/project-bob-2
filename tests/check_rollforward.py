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
  blank         a row found this year but blank there: the same, and the gate stays shut
  actuals       an actuals sheet with last year's history and no forecast doesn't win the row
  per figure    each figure is held back on its own discounting's rows; flags and dates are timing, not asked for
  weak          a row found only weakly holds the figures until a person keeps it, picks another, or keeps last
                year's values on purpose
  snapshot      a pasted copy of last year's rows (typed values) doesn't win over this year's own formula row
  agents        the rows the gate waits on settled by their numbers without a person, as the agents' picks
  unlabelled    a row with no label found by its numbers

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
          "value": 8.0, "unit": "%"},
         {"id": 5, "category": "conclusion", "key": "equity_value", "label": "Fair value of equity (ex-div)",
          "value_text": "198.4", "value": 198.4, "unit": "A$m"}]


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
    # the rows the cash flows come from: the distributions (amounts, to find this year), and the forecast flag
    # they're multiplied by (timing: the discounting depends on it, this year's figures don't wait for it)
    assert [o["row"] for o in low["origins"] if o["kind"] == "amount"] == ["CF!r11"], low["origins"]
    assert [(o["row"], o["kind_why"]) for o in low["origins"] if o["kind"] == "timing"] == \
        [("CF!r2", "its label says it's a flag or a date")], low["origins"]
    fo = dcffacts.Facts(sess, summary)
    assert fo.row_kind("CF", 3, "Period ending") == ("timing", "its values are dates")  # rising dates, no telling label

    class OneFee:  # a single number in the date range is an amount (a charge of 45,000), not a date
        def sheet(self, s): return {(9, 5): 45000.0}
        def timeline(self, s): return {5: 45838.0, 6: 46203.0}
    fo.sess = type("S", (), {"prior": OneFee(), "ov": None})()
    assert fo.row_kind("X", 9, "Access charge")[0] == "amount"
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
    def plan(ov_vd, prior_vd, cur_vd, this_vd=None):
        r = ov.plan_roll(sess, w["prior"], w["overlay"], False, ov_vd, prior_vd, cur_vd, this_vd)
        sess.shift = r["months"]
        return r, sess.period_shift(sess.prior, "CF"), sess.period_shift(sess.ov, "Val")
    r, cf, val = ov.deep(plan, "2025-09-30", "2025-06-30", "2025-12-31")  # the overlay dated after its client copy
    assert (r["months"], r["current_valuation_date"], cf, val) == (3, "2025-12-31", 0, 0), (r, cf, val)
    r, cf, val = ov.deep(plan, "2025-06-30", "2025-06-30", "2025-12-31")  # half a year: no year-end passed
    assert (r["months"], cf, val) == (6, 0, 0), (r, cf, val)
    r, cf, val = ov.deep(plan, "2025-06-30", "2025-06-30", "2026-06-30")  # a year: FY2026 has ended
    assert (r["months"], cf, val) == (12, 12, 12), (r, cf, val)
    # this year's model dated on last year's valuation date (its own date, not this year's valuation date):
    # nothing is rolled, it's flagged, and the feed agrees; guessing a roll from the timelines put a 21-year roll
    # on a real engagement
    r, cf, val = ov.deep(plan, "2025-09-30", "2025-06-30", "2025-09-30")
    assert r["date_check"] and not r["months_assumed"] and r["months_basis"].startswith(
        "check: this year's model's date (2025-09-30) isn't after last year's valuation date (2025-09-30)"), r
    assert "isn't known" not in r["months_basis"] and (r["months"], r["current_valuation_date"]) == (0, "2025-09-30"), r
    assert ov._feed({**summary, "roll": {**summary["roll"], **r}}, "current", None, None)[2] == 0
    assert ov._feed({**summary, "roll": {**summary["roll"], **r}}, "current", "2026-03-31", None)[2] == 6  # a date chosen
    held = {**summary, "roll": {**summary["roll"], **r}}
    g = ov.deep(lambda: ov.summary_table(sess, held, FACTS)["this_year_gaps"])
    assert g["date_check"] and not g["reliable"] and not any(x["reliable"] for x in g["by_cell"].values()), g
    # this year's valuation date set for the engagement: the roll runs to it, whatever the model's own date
    r, cf, val = ov.deep(plan, "2025-09-30", "2025-06-30", "2025-09-30", "2025-12-31")
    assert (r["months"], r["current_valuation_date"], r["date_check"]) == (3, "2025-12-31", False), r
    assert r["months_basis"].endswith("set for the engagement"), r
    # no dates at all: the timelines of only three sheets don't make a roll (at least five), so 12 is assumed, flagged
    r = ov.deep(ov.plan_roll, sess, w["prior"], w["overlay"], False, None, None, None)
    assert r["months_assumed"] and r["months"] == 12 and "moves seen on 3 sheet(s)" in r["months_basis"], r
    timelines_check()
    summary["roll"].update(ov.deep(plan, "2025-09-30", "2025-06-30", "2025-12-31")[0])
    starts_check(out)

    def three_months():
        defaults, _, months = ov._feed(summary, "current", None, None)
        sess.configure("current", defaults, months)
        sess.values([("Report", 5, 3)])
        return len(sess.client_reads), dict(sess.unmatched)
    reads, missed = ov.deep(three_months)
    assert reads and not missed, list(missed.items())[:3]
    print(f"roll: ok (Sep-25 overlay, Dec-25 model: 3 months, annual periods stay; {reads} client values read, none missed)")
    # ---- the gate: this year's figures show once every row the DCF's cash flows come from is found (or picked)
    gate = lambda: ov.summary_table(sess, summary, FACTS)["this_year_gaps"]
    g = ov.deep(gate)
    assert g["basis"] == "dcf" and g["dcf_rows"] == 1 and g["reliable"], g
    assert [x["row"] for x in g["timing"]] == ["CF!r2"] and g["timing"][0]["found"] == "CF!r2", g["timing"]
    real = fnd.locate
    fnd.locate = lambda s_, r_: None if (s_, r_) == ("CF", 11) else real(s_, r_)
    try:
        g = ov.deep(gate)
        assert not g["reliable"] and [x["row"] for x in g["dcf_missing"]] == ["CF!r11"], g
    finally:
        fnd.locate = real
    ov.deep(fnd.pick, "CF", 11, ("CF", 15))  # a person's pick opens it again
    assert ov.deep(gate)["reliable"]
    ov.deep(fnd.pick, "CF", 11, None)
    print("gate: ok (shut while the DCF's cash-flow row is missing, open once it's found or picked)")
    # ---- a row found this year but blank there (as an actuals sheet matched on its history is): last year's
    # values stand in, counted, never zero; the gate stays shut and says where the row was found
    fnd.locate = lambda s_, r_: None if (s_, r_) == ("CF", 11) else real(s_, r_)
    try:
        missing_row, _ = ov.deep(this_year)  # on the roll as it is now: the value with the row not found at all
        fnd.locate = lambda s_, r_: ("CF", 14) if (s_, r_) == ("CF", 11) else real(s_, r_)  # an empty row
        got3, stood3 = ov.deep(this_year)
        assert len(stood3) == 18 and abs(got3 - missing_row) < 1e-9, (len(stood3), got3, missing_row)
        assert len(sess.blank) == 18 and all(v[:2] == ("CF", 14) for v in sess.blank.values())
        g = ov.deep(gate)
        assert not g["reliable"] and g["dcf_missing"][0]["why"].startswith("found at CF!r14, but blank"), g["dcf_missing"]
        # a row the figures read that isn't the DCF's cash flows (the period dates): found blank, it shuts the gate too
        fnd.locate = lambda s_, r_: ("CF", 14) if (s_, r_) == ("CF", 3) else real(s_, r_)
        g = ov.deep(gate)
        assert not g["reliable"] and not g["dcf_missing"] and [x["row"] for x in g["blank_rows"]] == ["CF!r3"], g
        assert g["blank_rows"][0]["found"] == "CF!r14" and g["blank_rows"][0]["blank"] == g["blank_rows"][0]["of"]
    finally:
        fnd.locate = real
    assert ov.deep(gate)["reliable"]
    print(f"blank: ok (a row found but blank this year stands in with last year's values ({got3:.3f}, not 0) and shuts the gate)")
    # ---- each figure on its own rows: a figure with no discounting under it isn't held back by another's
    summary["outputs"].append({"fact_id": 7, "cell": "Bridge!C6", "label": "Less: distribution payable",
                               "value": -rp.DIST_PAYABLE, "scale": 1.0, "sign": -1, "report": "12.5"})
    facts2 = FACTS + [{"id": 7, "category": "conclusion", "key": "distribution_payable", "label": "Distribution payable",
                       "value_text": "12.5", "value": 12.5, "unit": "A$m"}]
    fnd.locate = lambda s_, r_: None if (s_, r_) == ("CF", 11) else real(s_, r_)
    try:
        g = ov.deep(lambda: ov.summary_table(sess, summary, facts2)["this_year_gaps"])
        bc = g["by_cell"]
        assert bc["Report!C5"]["basis"] == "dcf" and not bc["Report!C5"]["reliable"] and bc["Report!C5"]["missing"] == ["CF!r11"], bc
        assert bc["Bridge!C6"]["basis"] == "share" and bc["Bridge!C6"]["reliable"] and not g["reliable"], bc
        b = ov.deep(ov.value_bridge, sess, summary, facts2)
        assert [x["cell"] for x in b["bridges"]] == ["Bridge!C6"] and [x["cell"] for x in b["withheld"]] == ["Report!C5"], b
    finally:
        fnd.locate = real
        summary["outputs"].pop()
    print("per figure: ok (the figure whose DCF row is missing is held back; one with no discounting under it isn't)")
    # ---- a row the figures read that was found only weakly holds them back until a person looks: keeping the
    # row found (a pick of it) or keeping last year's values on purpose both count
    import rowfind
    sure = fnd.confident
    fnd.confident = lambda s_, r_: sure(s_, r_) and ((s_, r_) != ("CF", 3) or (s_, r_) in fnd.picks)
    try:
        g = ov.deep(gate)
        assert not g["reliable"] and [x["row"] for x in g["weak_rows"]] == ["CF!r3"] and g["weak_rows"][0]["found"] == "CF!r3", g
        ov.deep(fnd.pick, "CF", 3, ("CF", 3))  # "keep this row"
        assert ov.deep(gate)["reliable"]
    finally:
        fnd.confident = sure
        ov.deep(fnd.pick, "CF", 3, None)
    ov.deep(fnd.pick, "CF", 11, rowfind.STAND_IN)  # last year's values for the distributions, on purpose
    try:
        kept, stood = ov.deep(this_year)
        g = ov.deep(gate)
        assert len(stood) == 18 and g["reliable"] and not g["dcf_missing"], g
        assert all(w == "last year's values kept on purpose (your pick)" for k, w in sess.unmatched.items() if k[:2] == ("CF", 11))
    finally:
        ov.deep(fnd.pick, "CF", 11, None)
    print("weak: ok (a row found weakly holds the figures until it's kept, picked, or its last year's values kept on purpose)")
    # ---- timing: the forecast flag found only weakly is worked out from the period dates (1 for periods ending
    # after the valuation date), giving the figure the row found gives; with no such relation it's a row to find
    # (on the roll this year's model is built for: June 2025 to June 2026, where FY2026 is this year's actual)
    rule = ov.timing_rule(sess.prior, "CF", 2, sess.base_vd)
    assert rule and rule["kind"] == "after", rule
    assert ov.timing_rule(sess.prior, "CF", 3, sess.base_vd)["kind"] in ("timeline", "end")
    assert ov.timing_rule(sess.prior, "CF", 11, sess.base_vd) is None  # an amount follows from no date
    saved_roll = dict(summary["roll"])
    summary["roll"].update(ov.deep(plan, "2025-06-30", "2025-06-30", "2026-06-30")[0])
    found_val, _ = ov.deep(this_year)
    fnd.confident = lambda s_, r_: sure(s_, r_) and (s_, r_) != ("CF", 2)
    real_rule = ov.timing_rule
    try:
        g = ov.deep(gate)
        t2 = next(x for x in g["timing"] if x["row"] == "CF!r2")
        assert t2["derived"] and not t2["open"] and g["reliable"], g
        derived_val, _ = ov.deep(this_year)
        assert sess.derived_used and abs(derived_val - found_val) < 1e-9, (derived_val, found_val)
        ov.timing_rule = lambda *a: None
        g = ov.deep(gate)
        assert not g["reliable"] and next(x for x in g["timing"] if x["row"] == "CF!r2")["open"], g
    finally:
        ov.timing_rule = real_rule
        fnd.confident = sure
        summary["roll"].clear()
        summary["roll"].update(saved_roll)
        ov.deep(plan, "2025-09-30", "2025-06-30", "2025-12-31")
    assert ov.deep(gate)["reliable"] and not sess.derived
    print(f"timing rows: ok (a flag found weakly is worked out from the period dates: {derived_val:.3f}, as with the row)")
    # ---- the zero-roll check: this year's model at last year's valuation date gives about last year's figure;
    # a wrong row, found confidently, doesn't
    zero_of = lambda: next(r for r in ov.summary_table(sess, summary, FACTS)["rows"] if r.get("cell") == "Report!C5")
    zr = ov.deep(zero_of)["zero_roll"]
    assert zr["ok"] and 0.9 < zr["ratio"] < 1.1 and zr["valuation_date"] == "2025-09-30", zr
    fnd.locate = lambda s_, r_: ("CF", 13) if (s_, r_) == ("CF", 11) else real(s_, r_)  # the reserve top-up row
    try:
        t = ov.deep(lambda: ov.summary_table(sess, summary, FACTS))
        g, row = t["this_year_gaps"], next(r for r in t["rows"] if r.get("cell") == "Report!C5")
        assert not row["zero_roll"]["ok"] and row["zero_roll"]["ratio"] < 0.75 and not g["reliable"], row["zero_roll"]
        assert [x["cell"] for x in g["zero_roll_off"]] == ["Report!C5"] and not g["by_cell"]["Report!C5"]["reliable"]
    finally:
        fnd.locate = real
    print(f"zero roll: ok (at last year's date this year's model gives {zr['ratio']:.3f}x last year's; a wrong row "
          f"gives {row['zero_roll']['ratio']:.3f}x and holds the figure)")
    agents_check(sess, summary, fnd, sure)
    ov.deep(sess.configure, "workbook")
    roles_check(res, db)
    rebuilt_check(out)
    actuals_check(out)
    catalogue_check(out)
    snapshot_check(out)
    unlabelled_check(out)
    print("rollforward: all checks passed")


def agents_check(sess, summary, fnd, sure) -> None:
    """The row agents' first stage, by the numbers: rows the gate waits on are settled without a person, as the
    agents' picks, where a row of this year's model carries last year's numbers. A person's pick stays; the agents
    may not keep last year's values for a DCF cash-flow row."""
    import rowagent
    import rowfind
    doubt = {("CF", 3), ("CF", 11)}
    fnd.confident = lambda s_, r_: (s_, r_) in fnd.picks or (sure(s_, r_) and (s_, r_) not in doubt)
    try:
        assert not ov.deep(lambda: ov.summary_table(sess, summary, FACTS)["this_year_gaps"])["reliable"]
        ov.deep(fnd.pick, "CF", 3, ("CF", 3))  # a person's pick
        res = rowagent.run(sess, summary, FACTS)
        got = {d["row"]: (d["decision"], d["how"]) for d in res["decisions"]}
        assert got == {"CF!r11": ("CF!r15", "numbers")}, res["decisions"]
        assert fnd.pick_by[("CF", 11)] == "agent" and fnd.pick_by[("CF", 3)] == "you"
        assert fnd.explain("CF", 11)["how"] == "the agents' pick" and res["reliable_after"], res
        ov.deep(fnd.pick, "CF", 3, ("Ops", 5), "agent")  # the agents never replace a person's pick
        assert fnd.picks[("CF", 3)] == ("CF", 3)
        # the agents keeping last year's values for the DCF's cash-flow row: the figure stays held, and says why
        ov.deep(fnd.pick, "CF", 11, None)
        ov.deep(fnd.pick, "CF", 11, rowfind.STAND_IN, "agent")
        g = ov.deep(lambda: ov.summary_table(sess, summary, FACTS)["this_year_gaps"])
        assert not g["reliable"] and g["dcf_missing"][0]["why"].startswith("the agents couldn't find it"), g["dcf_missing"]
    finally:
        fnd.confident = sure
        for k in doubt:
            ov.deep(fnd.pick, *k, None)
    print(f"agents: ok (a row the gate waited on settled by its numbers, as the agents' pick; a person's pick stays; "
          f"a cash-flow row they can't find keeps its figure held)")


def unlabelled_check(out: Path) -> None:
    """A row with no label (a block of figures under a heading, as valuation outputs often are), moved to another
    sheet this year with its forecast revised: found by its numbers."""
    import xlsxwriter
    import rowagent
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
                if label:
                    ws.write(4 + i, 1, label)
                for k, v in enumerate(vals):
                    ws.write_number(4 + i, 3 + k, v)
        wb.close()
        return build_map.main(str(path), str(out / (path.stem + "_db")))["db"]
    flows = [120.0 + 7 * k for k in range(8)]
    prior = book(out / "unl_prior.xlsx", {"Output": [("", flows), ("", [5.0] * 8)]})
    later = [v * (1.03 if k >= 2 else 1.0) for k, v in enumerate(flows)]
    current = book(out / "unl_current.xlsx", {"Output": [("", [60.0 + k for k in range(8)])],
                                              "Valuation": [("", [9.0] * 8), ("", [300.0] * 8), ("", later)]})
    a, b = ov.Workbook(prior), ov.Workbook(current)
    f = rowfind.RowFinder(ov.RowMap(a, b), a, b)
    got = rowagent.by_numbers(f, "Output", 5)
    assert got and got["to"] == ("Valuation", 7) and got["check"]["ok"], got
    print(f"unlabelled: ok (a row with no label found by its numbers: {got['check']['text']})")


def timelines_check() -> None:
    """The roll from the timelines alone: the most common move, the smaller on a tie; a move beyond two years
    isn't a roll-forward (a sheet matched to another that starts decades later)."""
    class TL:
        def __init__(self, firsts):
            self.f = firsts

        def timeline(self, s):
            return {4: self.f[s]} if s in self.f else {}
    d = lambda y, m: ov.serial(date(y, m, 30 if m in (6, 9) else 31))
    last = {f"S{i}": d(2024, 6) for i in range(6)}
    now = {"S0": d(2024, 9), "S1": d(2024, 9), "S2": d(2045, 6), "S3": d(2045, 6), "S4": d(2025, 6), "S5": d(2025, 6)}
    fake = type("Sess", (), {"prior": TL(last), "ov": None, "current": TL(now), "rowmap": None,
                             "client_sheets": set(last), "ext_cached": {}})()
    m, basis, _ = ov.roll_months(fake, None, {}, False, (None, None, None, None))
    assert m == 3 and "on 2 of 6 sheet(s)" in basis, (m, basis)  # 3 and 12 twice each: the smaller; 252 isn't a roll
    print("timelines: ok (the most common move, the smaller on a tie, never beyond two years, from five sheets or more)")


def starts_check(out: Path) -> None:
    """A timeline of period starts (1 July each year, as many models date their columns): periods move only by
    the years that have ended, counted by their ends, not by the starts that fall in the window."""
    import xlsxwriter
    path = out / "starts.xlsx"
    wb = xlsxwriter.Workbook(path)
    dt = wb.add_format({"num_format": "dd-mmm-yy"})
    ws = wb.add_worksheet("Annual")
    ws.write(2, 1, "Year starting")
    for k in range(8):
        ws.write_datetime(2, 3 + k, date(2024 + k, 7, 1), dt)
    ws.write(4, 1, "Revenue")
    for k in range(8):
        ws.write_number(4, 3 + k, 100.0 + k)
    wb.close()
    wbk = ov.Workbook(build_map.main(str(path), str(out / "starts_db"))["db"])

    class Roll:  # the parts of a session period_shift uses
        _pshift = {}

    def shift(frm, to):
        r = Roll()
        r._pshift, r.base_vd = {}, ov.serial(date.fromisoformat(frm))
        r.shift = ov.months_between(frm, to)
        return ov.Session.period_shift(r, wbk, "Annual")
    got = (shift("2025-06-30", "2025-12-31"), shift("2025-06-30", "2026-06-30"), shift("2025-09-30", "2025-12-31"))
    assert got == (0, 12, 0), got  # the year starting 1 July 2025 hasn't ended by December
    print("starts: ok (a 1-July-start timeline moves a year only once the year has ended)")


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


def snapshot_check(out: Path) -> None:
    """A reconciliation sheet of pasted values in this year's model ("LINKED EBITDA": typed copies of last year's
    rows) has last year's history exactly and a full series; this year's own EBITDA, a formula, has restated
    history. The copy mustn't win: last year's row is a formula, the copy typed values."""
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
            for i, (label, vals, formula) in enumerate(rows):
                ws.write(4 + i, 1, label)
                for k, v in enumerate(vals):
                    c = rp.COL(3 + k)
                    if formula:
                        ws.write_formula(4 + i, 3 + k, formula.format(c=c), None, v)
                    else:
                        ws.write_number(4 + i, 3 + k, v)
        wb.close()
        return build_map.main(str(path), str(out / (path.stem + "_db")))["db"]
    rev = [100.0 * 1.03 ** k for k in range(8)]
    cost = [0.4 * v for v in rev]
    ebitda = [a - b for a, b in zip(rev, cost)]
    prior = book(out / "snap_prior.xlsx", {"Hub": [("Revenue", rev, None), ("Costs", cost, None),
                                                   ("EBITDA", ebitda, "={c}5-{c}6")]})
    rev2 = [v * 1.01 for v in rev]  # restated, history included
    cost2 = [0.4 * v for v in rev2]
    current = book(out / "snap_current.xlsx", {
        "Model": [("Revenue", rev2, None), ("Costs", cost2, None), ("EBITDA", [a - b for a, b in zip(rev2, cost2)], "={c}5-{c}6")],
        "Rec": [("LINKED EBITDA", ebitda, None)]})
    a, b = ov.Workbook(prior), ov.Workbook(current)
    f = rowfind.RowFinder(ov.RowMap(a, b), a, b)
    ex = f.explain("Hub", 7)
    assert ex["found"] == ("Model", 7), ex
    copy = next(x for x in ex["alternatives"] if x["row"] == "Rec!r5")
    assert any(e.startswith("shape: last year's row is formulas") for e in copy["evidence"]), copy
    print("snapshot: ok (a pasted copy with last year's history doesn't win over this year's own formula row)")


def catalogue_check(out: Path) -> None:
    """Every DCF in a workbook (valuation.catalogue), saved beside its model.db: a restarted server reads it back
    instead of redoing it (minutes on a large model), and redoes it when the model.db changes."""
    import os
    import xlsxwriter
    import valuation
    wb = xlsxwriter.Workbook(out / "sumproduct_dcf.xlsx")
    ws, dt = wb.add_worksheet("DCF"), wb.add_format({"num_format": "dd-mmm-yy"})
    vd, rate = date(2025, 6, 30), 0.08
    ws.write(0, 1, "Valuation date"), ws.write_datetime(0, 2, vd, dt)
    ws.write(1, 1, "Discount rate"), ws.write(1, 2, rate)
    ws.write(2, 1, "Period ending"), ws.write(4, 1, "Free cash flow"), ws.write(5, 1, "Discount factor")
    ws.write(7, 1, "Enterprise value")
    total = 0.0
    for k in range(6):
        end, c = date(2026 + k, 6, 30), rp.COL(3 + k)
        f = 1 / (1 + rate) ** ((end - vd).days / 365)
        ws.write_datetime(2, 3 + k, end, dt)
        ws.write_number(4, 3 + k, 50.0 + 5 * k)
        ws.write_formula(f"{c}6", f"=1/(1+$C$2)^(({c}3-$C$1)/365)", None, f)
        total += (50.0 + 5 * k) * f
    ws.write_formula("C8", "=SUMPRODUCT(D5:I5,D6:I6)", None, total)
    wb.close()
    path = build_map.main(str(out / "sumproduct_dcf.xlsx"), str(out / "sumproduct_dcf_db"))["db"]
    side = Path(path).parent / "catalogue.json"
    valuation._CACHE.clear()
    if side.exists():
        side.unlink()
    first = valuation.catalogue(path)
    assert side.exists() and first, "saved beside the model.db"
    valuation._CACHE.clear()
    real, valuation.find = valuation.find, lambda db: (_ for _ in ()).throw(AssertionError("recomputed"))
    try:
        assert valuation.catalogue(path) == first  # read back, not recomputed
        valuation._CACHE.clear()
        os.utime(path, (os.path.getatime(path), os.path.getmtime(path) + 5))  # the model.db changed
        try:
            valuation.catalogue(path)
            raise RuntimeError("a changed model.db must be catalogued again")
        except AssertionError:
            pass
    finally:
        valuation.find = real
    print("catalogue: ok (saved beside the model.db, read back after a restart, redone when the model.db changes)")


def actuals_check(out: Path) -> None:
    """An actuals input sheet this year carries last year's history exactly and nothing after it; the row that
    carries the history and the forecast is found, not the actuals row (which would read blank for every forecast
    year)."""
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
                    if v is not None:
                        ws.write_number(4 + i, 3 + k, v)
        wb.close()
        return build_map.main(str(path), str(out / (path.stem + "_db")))["db"]
    dist = [40.0 + 3 * k for k in range(8)]
    prior = book(out / "actuals_prior.xlsx", {"Hub": [("Distributions", dist), ("Opex", [30.0] * 8)]})
    revised = [v * (1.04 if k >= 2 else 1.0) for k, v in enumerate(dist)]
    current = book(out / "actuals_current.xlsx", {
        "Inputs actual": [("Distributions - actual", dist[:2] + [None] * 6), ("Opex - actual", [30.0] * 2 + [None] * 6)],
        "Model": [("Equity distributions", revised), ("Operating costs", [30.0] * 8)]})
    a, b = ov.Workbook(prior), ov.Workbook(current)
    f = rowfind.RowFinder(ov.RowMap(a, b), a, b)
    ex = f.explain("Hub", 5)
    assert ex["found"] == ("Model", 5), ex
    blank = next(x for x in ex["alternatives"] if x["row"] == "Inputs actual!r5")
    assert any("blank in 6 of the 8 periods" in e for e in blank["evidence"]), blank
    print("actuals: ok (the row with the forecast wins over an actuals sheet that has the history and blanks after it)")


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
    # this year's model has a sheet named like one of the overlay's, with other contents; and a report figure
    # (opening debt) whose value is only on a client sheet: neither moves a sheet across
    import xlsxwriter
    out = res["paths"]["prior"].parent
    path = out / "Client_model_FY26_named_alike.xlsx"
    wb = xlsxwriter.Workbook(path)
    rp.write_client(wb, 2025, True)
    ws = wb.add_worksheet("Val")
    for i, lab in enumerate(["Asset register", "Depreciation", "Written-down value", "Disposals"]):
        ws.write(4 + i, 1, lab)
        for k in range(6):
            ws.write_number(4 + i, 4 + k, 50.0 * (i + 1) + k)
    wb.close()
    orig3 = {k: wbs[2][k] for k in ("filename", "db_path", "source_path")}
    wbs[2] = {**wbs[2], "filename": path.name, "db_path": build_map.main(str(path), str(out / "named_alike_db"))["db"],
              "source_path": str(path)}
    od = n["debt_open"][2025]
    facts2 = facts + [{"id": 6, "category": "assumption", "key": "opening_debt", "label": "Opening debt",
                       "value_text": f"{od:.1f}", "value": od, "unit": "A$m"}]
    got = roles.suggest([{"id": 9, "filename": "report.pdf", "n_facts": 5}], wbs, facts2)
    r = {k: (v["id"], v.get("sheets")) for k, v in got["roles"].items()}
    assert r["prior_overlay"] == (2, ["Inputs", "Val", "Bridge", "Report"]), r["prior_overlay"]
    assert r["prior_model"][0] == 1 and r["current_model"][0] == 3, r
    # the adviser reused a client sheet's name: the client's "Summary" replaced by the valuation's own summary,
    # which holds the report's figure. Same name as the client's, other contents: it's the overlay's
    path = out / "Client_model_FY25_valuation_summary.xlsx"
    wb = xlsxwriter.Workbook(path)
    made = rp.write_client(wb, 2024, False, banner=False)
    rp.write_overlay(wb, made, date(2025, 6, 30), report_sheet="Summary")
    wb.close()
    wbs3 = [wbs[0], {**wbs[1], "filename": path.name, "db_path": build_map.main(str(path), str(out / "val_summary_db"))["db"],
                     "source_path": str(path)}, wbs[2]]
    got = roles.suggest([{"id": 9, "filename": "report.pdf", "n_facts": 5}], wbs3, facts2)
    r = {k: (v["id"], v.get("sheets")) for k, v in got["roles"].items()}
    assert r["prior_overlay"] == (2, ["Inputs", "Val", "Bridge", "Summary"]), r["prior_overlay"]
    # this year's model rebuilt (unlike last year's) and, like any big model, full of valuation words and charts:
    # the structure alone would call it a standalone overlay. The engagement has one overlay, the one with the
    # report's figures, so this year's model stays this year's model
    real = roles.structure

    def rebuilt(workbooks):
        st = real(workbooks)
        for j, sh in st["shape"].items():
            sh["family"] = [x for x in sh["family"] if 3 not in (j, x)]
        st["shape"][3].update(family=[], extra={}, points=9, standalone=True)
        return st
    roles.structure = rebuilt
    try:
        got = roles.suggest([{"id": 9, "filename": "report.pdf", "n_facts": 4}], wbs[:2] + [{**wbs[2], **orig3}], facts)
    finally:
        roles.structure = real
    r = {k: v["id"] for k, v in got["roles"].items()}
    assert r["prior_overlay"] == 2 and r["prior_model"] == 1 and r["current_model"] == 3, r
    assert got["workbooks"][3]["mode"] == "client model" and "typed as a client model" in got["workbooks"][3]["structure"][-1]
    print("roles: ok (a report figure on a client sheet doesn't make it the overlay; different dates don't stop the "
          "client's own file being last year's model; a rebuilt model that looks like valuation work stays a model)")


if __name__ == "__main__":
    main()
