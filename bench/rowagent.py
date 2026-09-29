"""The row agents: this year's row for each of last year's client rows the figures need, found and checked without
a person stopping the work.

The Summary's gate lists the rows it can't vouch for (overlay.summary_table: the DCF's cash-flow rows not found,
rows found but blank, rows found only weakly, timing rows that don't follow from the period dates). For each:
  1. by the numbers, with no model: the rows rowfind has for it and rows whose values are close to last year's
     in several periods (rowfind.near, which finds a row with no label at all). A candidate that carries last
     year's numbers (rowfind.check: over the periods both have, a median difference within 15%, the same kind of
     row, not blank where last year's has values) is taken, the numbers as the reason: forecasts are revised
     between valuations, not replaced.
  2. the rest, by the models (a Reader: the engagement's model and its reviewer): luna gets a dossier on last year's
     row that doesn't lean on its label (sheet, section, the labelled rows around it, a formula, the kind of row,
     its values by period, the rows it reads and that read it) and the candidates with their numbers checked, and
     may search this year's model and inspect rows before proposing one (or saying the model has no such line
     item). Sol checks each proposal, with the numbers, and accepts or rejects it with a reason; a rejection goes
     back to luna. At most MAX_TURNS actions and MAX_PROPOSALS proposals a row.
  3. the figures, again: the zero-roll check (this year's model at last year's valuation date should give about
     last year's figure). Where a figure is still off, the rows the figures read are ranked by how far their
     numbers are from last year's; sol, as advisor, sees the figures, those rows and the decisions so far, and
     names the rows to look at again and what to look for; luna looks again with that. At most ROUNDS rounds.
Each choice is the agents' pick (rowfind.pick(by="agent")): it counts as settled, shows as the agents', and a
person's pick always wins over it. The agents never keep last year's values for a row the DCF's cash flows come
from: that figure stays held, saying the agents couldn't find the row.
"""
import re
from concurrent.futures import ThreadPoolExecutor

import overlay as ovmod
import rowfind

MAX_TURNS = 6        # luna's actions for one row (searches, inspections, proposals)
MAX_PROPOSALS = 2    # proposals sol may reject before the row is left open
WORKERS = 3          # rows worked on at once (the model calls; the session's work is one thread anyway)
ROUNDS = 3           # rounds against the zero-roll check
ADVISED = 5          # rows the advisor may send back a round

KINDS = ("dcf_missing", "blank_rows", "weak_rows")


def _ref(text: str) -> tuple[str, int]:
    s, r = text.rsplit("!r", 1)
    return (s.split("]", 1)[1] if s.startswith("[") else s), int(r)


def open_rows(gaps: dict) -> list[dict]:
    """The rows the gate is waiting on, each once: {"row", "sheet", "r", "label", "kind", "origin"}."""
    out, seen = [], set()
    origins = {_ref(x) for x in gaps.get("dcf_origins") or []}  # the rows the DCF's cash flows come from
    for kind in KINDS:
        for x in gaps.get(kind) or []:
            k = _ref(x["row"])
            if k not in seen:
                seen.add(k)
                out.append({"row": x["row"], "sheet": k[0], "r": k[1], "label": x.get("label", ""), "kind": kind,
                            "origin": kind == "dcf_missing" or k in origins})
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


def run(sess, summary: dict, facts: list[dict], step=None, reader=None) -> dict:
    """One pass over the rows the gate is waiting on: by the numbers, then (with a reader) by the models. Session
    work goes through overlay.deep (one thread for every session), a step at a time, and the model calls happen
    outside it, so the page stays responsive while the agents work."""
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
    todo = [d for d in decisions if not d["decision"]]
    if reader is not None and todo:
        done = [0]

        def work(d):
            try:
                got = agent_row(reader, finder, d)
            except Exception as e:  # a model that isn't there: the row stays open, the rest go on
                got = {"decision": None, "why": f"the models couldn't be asked: {type(e).__name__}: {e}"[:300]}
            done[0] += 1
            step(f"The models checked {done[0]} of {len(todo)} row(s) the numbers didn't settle")
            return d, got
        with ThreadPoolExecutor(WORKERS) as pool:
            for d, got in pool.map(work, todo):
                d.update({k: v for k, v in got.items() if k != "decision"})
                if got.get("decision") == rowfind.STAND_IN:
                    if d["origin"]:  # never for the DCF's cash flows: the figure stays held
                        d["why"] = "the models found no such row: " + (got.get("why") or "")
                        continue
                    ovmod.deep(finder.pick, d["sheet"], d["r"], rowfind.STAND_IN, "agent")
                    d.update(decision="-", how="agents", to_label="last year's values, kept")
                elif got.get("decision"):
                    k = got["decision"]
                    ovmod.deep(finder.pick, d["sheet"], d["r"], k, "agent")
                    d.update(decision=f"{k[0]}!r{k[1]}", how="agents", to_label=finder.current.labels().get(k, ""))
    step("Checking the figures again")
    after = ovmod.deep(lambda: ovmod.summary_table(sess, summary, facts)["this_year_gaps"]) or {}
    rounds = []
    while reader is not None and after.get("zero_roll_off") and len(rounds) < ROUNDS:
        step(f"Round {len(rounds) + 1}: {len(after['zero_roll_off'])} figure(s) far from last year's at last year's date")
        r = one_round(reader, finder, after, decisions, rows, step)
        rounds.append(r)
        if not r["revisited"]:
            break
        step("Checking the figures again")
        after = ovmod.deep(lambda: ovmod.summary_table(sess, summary, facts)["this_year_gaps"]) or {}
        r["still_off"] = [x["label"] for x in after.get("zero_roll_off") or []]
    return {"decisions": decisions, "open_before": len(rows), "open_after": len(open_rows(after)),
            "reliable_after": after.get("reliable"), "rounds": rounds,
            "zero_roll_off": after.get("zero_roll_off") or []}


def suspects(finder, gaps: dict, limit: int = 12) -> list[dict]:
    """The rows the figures read that look least like last year's: not found (last year's values stand in), or
    found with numbers far from last year's, a person's picks left out. The DCF's rows first, then the widest gap."""
    origins = {_ref(x) for x in gaps.get("dcf_origins") or []}
    out = []
    for text in gaps.get("read_rows") or []:
        k = _ref(text)
        if finder.pick_by.get(k) == "you":
            continue
        ex = finder.explain(*k)
        if ex.get("stand_in") or ex.get("found") is None:
            out.append({"row": text, "label": finder.prior.labels().get(k, ""), "now": "not found: last year's values stand in",
                        "gap": 9.9, "origin": k in origins})
            continue
        c = finder.check(k[0], k[1], ex["found"])
        if c["periods"] and (c["median_gap"] or 0) > rowfind.CHECK_GAP:
            out.append({"row": text, "label": finder.prior.labels().get(k, ""),
                        "now": f"{ex['found'][0]}!r{ex['found'][1]} ({ex['how']}): {c['text']}", "gap": c["median_gap"],
                        "origin": k in origins})
    return sorted(out, key=lambda x: (not x["origin"], -x["gap"]))[:limit]


ADVISE = """You advise on rolling a valuation forward onto this year's version of a client's model. Each of last
year's client rows the valuation reads was matched to a row of this year's model. A check: this year's model at
last year's valuation date, rolled by nothing, should give about last year's figures (forecasts are revised, not
replaced). These figures don't:
{figures}

The client rows the figures read that look least like last year's (what they're matched to now, and how their
numbers compare with last year's over the periods both models have):
{suspects}

What was decided so far:
{decisions}

Name at most {n} rows to look at again ("Sheet!rN" as above), the most likely cause first, each with why and what
to look for in this year's model. Say done if none of them could explain the gap."""
_ADVICE = {"type": "json_schema", "name": "row_advice", "strict": True, "schema": {
    "type": "object", "additionalProperties": False, "required": ["rows", "done"],
    "properties": {"done": {"type": "boolean"}, "rows": {"type": "array", "items": {
        "type": "object", "additionalProperties": False, "required": ["row", "why", "look_for"],
        "properties": {"row": {"type": "string"}, "why": {"type": "string"}, "look_for": {"type": "string"}}}}}}}


def one_round(reader, finder, gaps: dict, decisions: list[dict], rows: list[dict], step) -> dict:
    """Sol advises which rows to look at again for the figures still off; luna looks again at each with the advice."""
    sus = ovmod.deep(suspects, finder, gaps)
    figs = "\n".join(f"- {x['label']}: {x['value']} at last year's date, {x['ratio']}x last year's" for x in gaps["zero_roll_off"])
    dec = "\n".join(f"- {d['row']} {d.get('label', '')}: " + (f"{d['decision']} ({d.get('how')}: {d.get('why')})"
                                                               if d.get("decision") else f"open ({d.get('why')})")
                     for d in decisions) or "- nothing: every row was found by rowfind"
    sus_text = "\n".join(f"- {x['row']} {x['label']}{' (a DCF cash-flow row)' if x['origin'] else ''}: {x['now']}"
                         for x in sus) or "- none stands out"
    advice = reader._call(reader.reviewer_model, ADVISE.format(figures=figs, suspects=sus_text, decisions=dec, n=ADVISED),
                          None, _ADVICE, "row-advice")
    known = {x["row"]: x for x in sus} | {d["row"]: d for d in rows}
    todo = [a for a in advice.get("rows") or [] if a.get("row") in known][:ADVISED]
    out = {"figures_off": [x["label"] for x in gaps["zero_roll_off"]], "advice": advice.get("rows") or [],
           "done": advice.get("done"), "revisited": []}
    if advice.get("done") or not todo:
        return out
    for a in todo:
        k = _ref(a["row"])
        if finder.pick_by.get(k) == "you":
            continue
        step(f"Looking again at {a['row']}: {a['why'][:80]}")
        before = finder.explain(*k).get("found")
        ovmod.deep(finder.pick, k[0], k[1], None, "agent")  # the agents' own earlier choice is set aside
        x = {"row": a["row"], "sheet": k[0], "r": k[1], "label": finder.prior.labels().get(k, ""),
             "origin": any(s["row"] == a["row"] and s["origin"] for s in sus) or known[a["row"]].get("origin", False)}
        note = (f"A first pass didn't reproduce last year's figures at last year's date ({figs.strip()}). "
                f"A reviewer advises looking again at this row: {a['why']} Look for: {a['look_for']}"
                + (f" It was matched to {before[0]}!r{before[1]}." if before else ""))
        try:
            got = agent_row(reader, finder, x, note)
        except Exception as e:
            got = {"decision": None, "why": f"the models couldn't be asked: {type(e).__name__}: {e}"[:300]}
        rec = {"row": a["row"], "advice": a["why"], "why": got.get("why"), "review": got.get("review"),
               "calls": got.get("calls"), "decision": None}
        dec_ = got.get("decision")
        if dec_ == rowfind.STAND_IN and not x["origin"]:
            ovmod.deep(finder.pick, k[0], k[1], rowfind.STAND_IN, "agent")
            rec["decision"] = "-"
        elif dec_ and dec_ != rowfind.STAND_IN:
            ovmod.deep(finder.pick, k[0], k[1], dec_, "agent")
            rec["decision"] = f"{dec_[0]}!r{dec_[1]}"
        out["revisited"].append(rec)
        d = next((d for d in decisions if d["row"] == a["row"]), None)
        if d is None:
            d = {**x, "kind": "advised", "decision": None}
            decisions.append(d)
        d.update(decision=rec["decision"], how="agents" if rec["decision"] else None, why=got.get("why"),
                 review=got.get("review"), advice=a["why"],
                 to_label=(finder.current.labels().get(dec_, "") if isinstance(dec_, tuple) else
                           "last year's values, kept" if rec["decision"] == "-" else None))
    return out


# ---- the models ----------------------------------------------------------------------------------------------

LUNA = """You find, in this year's version of a client's financial model, the row that is the same line item as a
row of last year's model. The versions can differ a lot: rows moved, renamed, split, restructured, sheets renamed
or rebuilt, and a row may have no label (a block of figures under a heading). Forecasts are revised between
valuations, not replaced: the same line item carries numbers close to last year's in the periods both models have,
and its actual years usually equal. Beware of lookalikes: a reconciliation sheet of pasted values (rows like
"LINKED ...", typed values where last year's row was formulas) copies last year's numbers but isn't the model's own
row; an actuals sheet has the history and nothing after it; a total or subtotal is not its parts.

Last year's row:
{row}

This year's candidates so far, each checked against last year's numbers:
{candidates}
{context}{history}
Reply with one action:
- "search": rows of this year's model with these words in their label or section (query: a few words)
- "inspect": one row of this year's model in full, checked against last year's (row: "Sheet!rN")
- "propose": this is the row (row: "Sheet!rN"): why, and how confident you are
- "not_in_this_model": this year's model has no such line item (why)
{left} action(s) left{must}."""

SOL = """You review a proposed match between a row of last year's client model and a row of this year's version of
the model, before the valuation is rolled forward on it. Accept only if it is the same line item: the same thing,
measured the same way, on the same basis (not a total of it, a part of it, a pasted copy, or an actuals-only row).
The numbers check compares the two over the periods both models have, with no roll: forecasts are revised, not
replaced, so a large median difference needs a reason you can see (a sign convention, units), or it's a reject.

Last year's row:
{row}

Proposed this year: {proposed}
The numbers check: {check}
Why it was proposed: {why}

The other candidates:
{candidates}

Reply with your verdict (accept or reject), why in one or two sentences, and better_row ("Sheet!rN") if one of the
other candidates is clearly the right one, else null."""

_S, _N = {"type": "string"}, {"type": ["string", "null"]}
_ACTION = {"type": "json_schema", "name": "row_action", "strict": True, "schema": {
    "type": "object", "additionalProperties": False, "required": ["action", "query", "row", "why", "confidence"],
    "properties": {"action": {"type": "string", "enum": ["search", "inspect", "propose", "not_in_this_model"]},
                   "query": _N, "row": _N, "why": _S, "confidence": {"type": "string", "enum": ["high", "medium", "low"]}}}}
_VERDICT = {"type": "json_schema", "name": "row_verdict", "strict": True, "schema": {
    "type": "object", "additionalProperties": False, "required": ["verdict", "why", "better_row"],
    "properties": {"verdict": {"type": "string", "enum": ["accept", "reject"]}, "why": _S, "better_row": _N}}}


def _num(v) -> str:
    return f"{v:,.4g}" if isinstance(v, float) and abs(v) < 1e4 else f"{v:,.0f}" if isinstance(v, float) else str(v)


def describe(finder, wb, s: str, r: int, which: int) -> str:
    """A row in words that don't lean on its label: where it sits, what's around it, what kind of row it is, its
    values by period, and the rows it reads and that read it. which: 0 last year's model, 1 this year's."""
    labels = wb.labels()
    meta = wb.db.execute("SELECT section, units, n_formula, n_const FROM rows WHERE sheet=? AND row=?", (s, r)).fetchone()
    section, units, nf, nc = meta if meta else (None, None, None, None)
    around = [f"r{rr} {labels[(s, rr)]}" for rr in range(r - 4, r + 5) if rr != r and labels.get((s, rr))]
    f = wb.db.execute("SELECT addr, formula FROM cells WHERE sheet=? AND row=? AND formula IS NOT NULL LIMIT 1",
                      (s, r)).fetchone()
    series = [(w, v) for w, v in finder._series(wb, s, r).items()]
    step = max(1, len(series) // 12)
    vals = ", ".join(f"{ovmod.to_date(w).isoformat()[:7]}: {_num(v)}" for w, v in series[::step][:12])
    reads, read_by = finder._index()["edges"][which]
    name = lambda k: f"{k[0]}!r{k[1]} {labels.get(k, '')}".strip()
    L = [f"{s}!r{r} label: '{labels.get((s, r), '')}'" + (f"; section: {section}" if section else "")
         + (f"; units: {units}" if units else "")]
    if nf is not None:
        L.append(f"kind: {nf} formula(s), {nc} typed value(s)")
    if around:
        L.append("rows around it: " + "; ".join(around))
    if f:
        L.append(f"a formula: {f[0]} ={f[1]}"[:240])
    L.append("values: " + (vals or "none by period"))
    if reads.get((s, r)):
        L.append("reads: " + "; ".join(name(k) for k in sorted(reads[(s, r)])[:8]))
    if read_by.get((s, r)):
        L.append("read by: " + "; ".join(name(k) for k in sorted(read_by[(s, r)])[:8]))
    return "\n".join(L)


def view(finder, s: str, r: int, k: tuple) -> str:
    c = finder.check(s, r, k)
    return describe(finder, finder.current, *k, 1) + f"\nagainst last year's: {c['text']}" + (" (carries them)" if c["ok"] else "")


def search(finder, query: str, limit: int = 15) -> list[tuple]:
    """Rows of this year's model with the query's words in their label or section, most words first."""
    words = [w for w in re.findall(r"[A-Za-z0-9]{3,}", query or "")][:6]
    hits = {}
    for w in words:
        like = f"%{w}%"
        for sh, rr in finder.current.db.execute("SELECT sheet, row FROM rows WHERE label LIKE ? OR section LIKE ? LIMIT 200",
                                                (like, like)):
            hits[(sh, rr)] = hits.get((sh, rr), 0) + 1
    return sorted(hits, key=lambda k: (-hits[k], k))[:limit]


def _row(text: str | None, finder) -> tuple | None:
    m = re.match(r"^\s*(?:\[\d+\])?(.+?)!r(\d+)\s*$", text or "")
    if not m:
        return None
    k = (m[1], int(m[2]))
    return k if finder.current.db.execute("SELECT 1 FROM rows WHERE sheet=? AND row=?", k).fetchone() else None


def agent_row(reader, finder, x: dict, note: str | None = None) -> dict:
    """Luna finds this year's row for one of last year's, sol checks it: {"decision": (sheet, row) | STAND_IN |
    None, "why", "review", "check", "calls"}. The session's work goes through overlay.deep."""
    s, r = x["sheet"], x["r"]
    deep = ovmod.deep
    me = deep(describe, finder, finder.prior, s, r, 0)
    cands = deep(lambda: [(k, view(finder, s, r, k)) for k in candidates(finder, s, r)[:8]])
    cand_text = "\n\n".join(v for _, v in cands) or "none found yet"
    history, proposals, calls = [], 0, {"luna": 0, "sol": 0}
    for turn in range(MAX_TURNS):
        left = MAX_TURNS - turn
        prompt = LUNA.format(row=me, candidates=cand_text, left=left, context=f"\nContext: {note}\n" if note else "",
                             history=("\nWhat you've done so far:\n" + "\n\n".join(history) + "\n") if history else "",
                             must=": propose a row or say it isn't in this model" if left == 1 else "")
        a = reader._call(reader.model, prompt, None, _ACTION, "row-agent")
        calls["luna"] += 1
        act = a.get("action")
        if act == "search":
            found = deep(lambda: [view(finder, s, r, k) for k in search(finder, a.get("query") or "")[:6]])
            history.append(f"You searched for '{a.get('query')}':\n" + ("\n\n".join(found) or "nothing found"))
        elif act == "inspect":
            k = _row(a.get("row"), finder)
            history.append(f"You inspected {a.get('row')}:\n" + (deep(view, finder, s, r, k) if k else "no such row"))
        elif act == "not_in_this_model":
            return {"decision": rowfind.STAND_IN, "why": a.get("why"), "confidence": a.get("confidence"), "calls": calls}
        elif act == "propose":
            k = _row(a.get("row"), finder)
            if not k:
                history.append(f"You proposed {a.get('row')}, which isn't a row of this year's model.")
                continue
            chk = deep(finder.check, s, r, k)
            others = "\n\n".join(v for kk, v in cands if kk != k) or "none"
            v = reader._call(reader.reviewer_model, SOL.format(row=me, proposed=deep(view, finder, s, r, k),
                                                               check=chk["text"], why=a.get("why"), candidates=others)
                             + (f"\n\nContext: {note}" if note else ""),
                             None, _VERDICT, "row-review")
            calls["sol"] += 1
            if v.get("verdict") == "accept":
                return {"decision": k, "why": a.get("why"), "confidence": a.get("confidence"), "review": v.get("why"),
                        "check": chk, "calls": calls}
            proposals += 1
            history.append(f"You proposed {a.get('row')}; the reviewer rejected it: {v.get('why')}"
                           + (f" It suggests {v['better_row']}." if v.get("better_row") else ""))
            if proposals >= MAX_PROPOSALS:
                return {"decision": None, "why": f"the reviewer rejected {proposals} proposals: {v.get('why')}",
                        "calls": calls}
    return {"decision": None, "why": "the models ran out of actions without a row the reviewer accepted", "calls": calls}
