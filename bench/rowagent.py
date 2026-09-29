"""The row agents: this year's row for each of last year's client rows the figures need, found and checked without
a person stopping the work.

The Summary's gate lists the rows it can't vouch for (overlay.summary_table: the DCF's cash-flow rows not found,
rows found but blank, rows found only weakly, timing rows that don't follow from the period dates). For each:
  1. by the numbers, with no model: the rows rowfind has for it and rows whose values are close to last year's
     in several periods (rowfind.near, which finds a row with no label at all). A candidate that carries last
     year's numbers (rowfind.check: over the periods both have, a median difference within 15%, the same kind of
     row, not blank where last year's has values) is taken, the numbers as the reason: forecasts are revised
     between valuations, not replaced.
Each choice is the agents' pick (rowfind.pick(by="agent")): it counts as settled, shows as the agents', and a
person's pick always wins over it. The agents never keep last year's values for a row the DCF's cash flows come
from: that figure stays held, saying the agents couldn't find the row.
"""
import overlay as ovmod

KINDS = ("dcf_missing", "blank_rows", "weak_rows")


def _ref(text: str) -> tuple[str, int]:
    s, r = text.rsplit("!r", 1)
    return (s.split("]", 1)[1] if s.startswith("[") else s), int(r)


def open_rows(gaps: dict) -> list[dict]:
    """The rows the gate is waiting on, each once: {"row", "sheet", "r", "label", "kind", "origin"}."""
    out, seen = [], set()
    for kind in KINDS:
        for x in gaps.get(kind) or []:
            k = _ref(x["row"])
            if k not in seen:
                seen.add(k)
                out.append({"row": x["row"], "sheet": k[0], "r": k[1], "label": x.get("label", ""), "kind": kind,
                            "origin": kind == "dcf_missing"})
    for x in gaps.get("timing") or []:
        k = _ref(x["row"])
        if x.get("open") and k not in seen:
            seen.add(k)
            out.append({"row": x["row"], "sheet": k[0], "r": k[1], "label": x.get("label", ""), "kind": "timing",
                        "origin": False})
    return out


def candidates(finder, s: str, r: int) -> list[tuple]:
    """This year's rows worth checking for last year's row: what rowfind found and its alternatives, then rows
    close to it by value, each once."""
    ex = finder.explain(s, r)
    out = [ex["found"]] if ex.get("found") else []
    for a in ex.get("alternatives") or []:
        out.append(_ref(a["row"]))
    out += finder.near(s, r)
    seen, uniq = set(), []
    for k in out:
        if k and k not in seen:
            seen.add(k)
            uniq.append(k)
    return uniq


def by_numbers(finder, s: str, r: int) -> dict | None:
    """The candidate that carries last year's numbers best, if one does: {"to", "check", "label"}."""
    best = None
    for k in candidates(finder, s, r):
        c = finder.check(s, r, k)
        if c["ok"] and (best is None or (c["median_gap"], -c["periods"]) < (best["check"]["median_gap"], -best["check"]["periods"])):
            best = {"to": k, "check": c, "label": finder.current.labels().get(k, "")}
    return best


def run(sess, summary: dict, facts: list[dict], step=None) -> dict:
    """One pass over the rows the gate is waiting on. Session work goes through overlay.deep (one thread for
    every session), a step at a time, so the page stays responsive while the agents work."""
    step = step or (lambda msg: None)
    finder = sess.rowmap
    if not finder:
        return {"decisions": [], "why": "no current client model"}
    step("Finding the rows the figures need")
    gaps = ovmod.deep(lambda: ovmod.summary_table(sess, summary, facts)["this_year_gaps"]) or {}
    rows = open_rows(gaps)
    decisions = []
    for i, x in enumerate(rows, 1):
        step(f"Checking row {i} of {len(rows)} by its numbers: {x['label'] or x['row']}")
        s, r = x["sheet"], x["r"]
        if finder.pick_by.get((s, r)) == "you":
            continue
        got = ovmod.deep(by_numbers, finder, s, r)
        d = {**x, "decision": None, "how": None, "why": None}
        if got:
            ovmod.deep(finder.pick, s, r, got["to"], "agent")
            d.update(decision=f"{got['to'][0]}!r{got['to'][1]}", to_label=got["label"], how="numbers",
                     why=got["check"]["text"], check=got["check"])
        else:
            d["why"] = "no row of this year's model carries last year's numbers"
        decisions.append(d)
    step("Checking the figures again")
    after = ovmod.deep(lambda: ovmod.summary_table(sess, summary, facts)["this_year_gaps"]) or {}
    return {"decisions": decisions, "open_before": len(rows), "open_after": len(open_rows(after)),
            "reliable_after": after.get("reliable")}
