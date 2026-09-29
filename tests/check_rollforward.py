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
    # 18 periods rolled on a year: 17 are in last year's forecast; the last (FY2044) is past it and stays unmatched
    assert len(stood2) == 17 and all(k[:2] == ("CF", 11) for k in stood2), len(stood2)
    assert got2 > 0 and abs(got2 - expect) > 1e-3, got2  # last year's forecast, not zero
    print(f"standing in: ok (with the distributions missing: {got2:.3f}, from last year's forecast, not 0)")
    ov.deep(sess.configure, "workbook")
    print("rollforward: all checks passed")


if __name__ == "__main__":
    main()
