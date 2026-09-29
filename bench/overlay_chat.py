"""Ask about an engagement: the Model Desk's chat agent (agent.py) on the Valuation Desk's Python overlay step.

The Model Desk's workbook tools (find, rows, trace, cells, sql, chart, dcf) work on any of the engagement's
workbooks (a `workbook` argument picks the prior overlay, the prior client model or the current client model),
and five more tools run the Python overlay (overlay.py):
  overlay_run      every output on a feed (as saved / prior client model / current model rolled forward), with
                   changed inputs: the what-ifs that recompute the cash flows, which the dcf tool can't
  overlay_chart    a row across the timeline: as Excel saved it, recomputed on a feed, and with changes
  overlay_dcf      a DCF on the module's own cash flows under another discounting method, checked against the
                   module's anchor value
  overlay_formula  the Python compiled for a cell and the client-model values it reads
  overlay_inputs   the overlay's inputs (what overlay_run can change), by label or cell
The instructions carry the engagement: its files, the report's key facts, the levers and outputs, the feeds.
"""
import copy
import json

import agent

WORKBOOK_TOOLS = ("find", "rows", "trace", "cells", "sql", "chart", "dcf")
ROLE_KEYS = ("prior_overlay", "prior_model", "current_model")
_CHANGES = {"type": "array", "description": "inputs to change, e.g. [{\"cell\": \"Val_Inputs!C5\", \"value\": 0.075}]: "
                                             "raw cell values (0.075 for 7.5%), dates as YYYY-MM-DD",
            "items": {"type": "object", "properties": {"cell": {"type": "string"},
                                                       "value": {"type": ["number", "string"]}},
                      "required": ["cell", "value"]}}
_FEED = {"type": "string", "enum": ["workbook", "prior", "current"],
         "description": "where client-model values come from: workbook = as saved in the overlay (default), "
                        "prior = the prior client model, current = this year's client model, rolled forward"}

OVERLAY_TOOLS = [
    {"type": "function", "name": "overlay_run",
     "description": "Run the Python overlay on a feed with changed inputs and get every output (enterprise value, "
                    "equity value, ...) against the feed's base and the workbook, plus the report's figure where "
                    "there is one. Use for any what-if that changes cash flows (growth, CPI, volumes, costs), the "
                    "discount rate or growth levers, the valuation date, or rolling forward onto this year's model.",
     "parameters": {"type": "object", "properties": {
         "changes": _CHANGES, "feed": _FEED,
         "valuation_date": {"type": "string", "description": "feed=current only: the new valuation date, YYYY-MM-DD "
                                                             "(default: the roll-forward's)"},
         "months": {"type": "integer", "description": "feed=current only: months to roll forward (default: the "
                                                     "gap between the models' timelines)"}},
         "required": []}},
    {"type": "function", "name": "overlay_chart",
     "description": "Draw one chart of an overlay row across its timeline as the Python overlay computes it: the "
                    "values Excel saved, then one line per feed listed (e.g. [\"prior\", \"current\"] to compare "
                    "last year's client model with this year's, rolled forward), and with changes on the last feed. "
                    "Add components (the rows that add up to it) to draw them as stacked columns under those lines, "
                    "in one chart. Pass the row range. Returns a short summary of what was drawn.",
     "parameters": {"type": "object", "properties": {
         "title": {"type": "string"}, "range": {"type": "string", "description": "one overlay row, e.g. DCF!C7:V7"},
         "feeds": {"type": "array", "items": {"type": "string", "enum": ["workbook", "prior", "current"]},
                   "description": "feeds to draw (default none: only the saved values, plus changes if given)"},
         "components": {"type": "array", "maxItems": 6, "items": {"type": "string"},
                        "description": "overlay rows that add up to the charted row (e.g. its trust, company and "
                                       "credit components), drawn as stacked columns over the same periods"},
         "component_feed": {"type": "string", "enum": ["workbook", "prior", "current"],
                            "description": "which feed the components come from (default: the last feed listed)"},
         "changes": _CHANGES, "kind": {"type": "string", "enum": ["line", "bar", "stacked", "area", "combo"],
                                       "description": "default line; combo when components are given"}},
         "required": ["title", "range"]}},
    {"type": "function", "name": "overlay_dcf",
     "description": "Recompute one of the overlay's DCFs from the cash flows the Python overlay computes on a feed "
                    "with changes, check it against the module's own anchor value, and rerun it under another "
                    "discounting method: rate, mid-period, day count, cut-off, bridge items left out. Returns the "
                    "workbook, live and scenario values and a low / high rate sensitivity. Call with no arguments "
                    "for the overlay's EV / equity DCF as saved.",
     "parameters": {"type": "object", "properties": {
         "cell": {"type": "string", "description": "the DCF's anchor cell, e.g. DCF!D12 (default: the main one)"},
         "changes": _CHANGES, "feed": _FEED,
         "rate": {"type": "number", "description": "discount rate for the scenario method, e.g. 0.08"},
         "timing": {"type": "string", "enum": ["end", "mid"]},
         "day_count": {"type": "string", "enum": ["actual/actual", "actual/365"]},
         "cutoff": {"type": "string", "description": "'model' (default), 'none', or YYYY-MM-DD"},
         "drop_bridge": {"type": "array", "items": {"type": "string"}, "description": "labels of bridge items to leave out"}},
         "required": []}},
    {"type": "function", "name": "overlay_value",
     "description": "How a valuation figure is built in the overlay: from the cell a report conclusion was matched "
                    "to (default: the conclusions that tie to the report) down through every cell it reads to the "
                    "discounting (SUMPRODUCT, XNPV, NPV or a sum of present values). Gives each formula in line-item "
                    "words, Excel's and the Python overlay's values, each discounting recomputed with its rate, "
                    "valuation date and convention, and the rows its cash flows add up. Use it first for how the "
                    "value is derived, which cash flows sit behind it, or before charting a valuation's cash flows.",
     "parameters": {"type": "object", "properties": {
         "cell": {"type": "string", "description": "the figure's cell, e.g. Valuation!F12 (default: the report's)"}},
         "required": []}},
    {"type": "function", "name": "overlay_formula",
     "description": "The Python compiled for an overlay cell's line item (one function per row, the Excel formula "
                    "beside each branch), its value, and the client-model values it reads.",
     "parameters": {"type": "object", "properties": {"cell": {"type": "string", "description": "e.g. DCF!D12"}},
                    "required": ["cell"]}},
    {"type": "function", "name": "overlay_inputs",
     "description": "Search the overlay's inputs (hard-coded cells overlay_run can change) by label or cell.",
     "parameters": {"type": "object", "properties": {"text": {"type": "string"}}, "required": ["text"]}},
]

SYSTEM = """You answer questions about a recurring valuation engagement. Its files: last year's valuation report
(its key facts are listed below), last year's client model, the overlay (our workings that take the client model's
cash flows to the valuation: a separate workbook, or sheets inside the client model), and this year's client
model. The overlay has been rebuilt as a Python module from its own formulas, "the Python overlay": it reproduces
the workbook cell by cell and can be rerun with changed inputs, fed from the prior client model, or rolled forward
onto this year's client model.

Two kinds of tools:
- Workbook tools (find, rows, trace, cells, sql, chart, dcf) read a workbook's saved formulas and values. Pass
  workbook="prior_model" or "current_model" to read a client model; the default is the overlay workbook. The
  guidance further down for these tools applies to every workbook.
- Overlay tools run the Python overlay: overlay_run, overlay_chart, overlay_dcf, overlay_value, overlay_formula,
  overlay_inputs. For how the value is derived or which cash flows sit behind it, start with overlay_value: it
  traces the report's figure down to the discounting in one call.
Use overlay_run or overlay_chart for any what-if that changes the cash flows or rolls the valuation forward
(growth, CPI, volumes, the discount rate or growth levers, the valuation date, this year's model); the dcf tool
only re-discounts saved cash flows. Use overlay_dcf for discounting-method questions on the live numbers
(mid-period, day count, cut-off, bridge items). For a what-if, give the outputs from overlay_run (enterprise and
equity value) as well as any chart. To show how a row is made up and how it compares with last year in one
chart, call overlay_chart with the row, its feeds and its components (the rows that add up to it): the parts are
stacked as columns and the feeds drawn as lines. Say which feed and which changes each number comes from, compare
with the workbook and the report where it helps, and cite cells as Sheet!A1. Change inputs with raw cell values:
0.075 for 7.5%, dates as YYYY-MM-DD. Levers are the report's assumptions found in the overlay; any other overlay
input can be changed too (find it with overlay_inputs).

Guidance for the workbook tools:
"""


def _changes(items) -> dict:
    out = {}
    for x in items or []:
        if isinstance(x, dict) and x.get("cell") is not None:
            out[str(x["cell"]).strip()] = x.get("value")
    return out


def _n(v, d: int = 2) -> str:
    return f"{v:,.{d}f}" if isinstance(v, (int, float)) and not isinstance(v, bool) else str(v)


def _run_text(res: dict) -> str:
    lines = []
    r = res.get("roll")
    feed = {"workbook": "as saved in the overlay", "prior": "fed from the prior client model",
            "current": "fed from the current client model, rolled forward"}[res["mode"]]
    lines.append(f"Feed: {feed}" + (f" {r['months']} months to {r['valuation_date']} (periods {r.get('first_period')} "
                                     f"to {r.get('last_period')})" if r else ""))
    if res.get("changes"):
        lines.append("Changes: " + ", ".join(f"{k} = {v}" for k, v in res["changes"].items()))
    for o in res["outputs"]:
        bits = [f"workbook {_n(o['workbook'])}", f"this feed {_n(o['base'])}"]
        if res.get("changes"):
            bits.append(f"with changes {_n(o['value'])}" + (f" (change {o['change']:+,.2f})" if o.get("change") is not None else ""))
        if o.get("report"):
            bits.append(f"report {o['report']}" + (" (ties)" if (o.get("tie") or {}).get("ok") else ""))
        lines.append(f"- {o['cell']} {o['label']}: " + "; ".join(bits))
    if res.get("n_unmatched"):
        lines.append(f"{res['n_unmatched']} client value(s) couldn't be matched in this year's model and are blank, e.g. "
                     + "; ".join(f"{u['cell']} ({u['why']})" for u in res["unmatched"][:4]))
    return "\n".join(lines)


def _dcf_text(r: dict) -> str:
    if not r.get("selected"):
        return "No DCF found on the overlay sheets (a SUMPRODUCT of a cash-flow row and a discount-factor row)."
    L, S, W = r["live"], r["scenario"], r["workbook"]
    lines = [f"{r['label']} ({r['selected']}), {r['feed']['words']}"
             + (f", changes {r['feed']['changes']}" if r["feed"]["changes"] else "") + ":",
             f"- as saved in Excel: {_n(W['total'])} (the workbook's cell: {_n(W['anchor'])})",
             f"- Python overlay: {_n(L['total'])} recomputed from the module's cash flows; the module's own cell gives "
             f"{_n(L['anchor'])} ({'they agree' if r['agrees'] else 'THEY DIFFER'})",
             f"  method: {L['rate']:.2%}, {L['timing']}-of-period, {L['day_count']}, valuation date {L['valuation_date']}, "
             f"cut-off {L['terminal_date'] or 'none'}, {L['periods']} periods {L['first_period']} to {L['last_period']}"]
    if L["bridge"]:
        lines.append("  bridge: " + "; ".join(f"{b['label']} {_n(b['value'])}" for b in L["bridge"]))
    if r["method_changes"]:
        lines.append(f"- scenario method ({'; '.join(r['method_changes'])}): {_n(S['total'])} "
                     f"({S['total'] - L['total']:+,.2f} vs the Python overlay)")
    lines.append("- sensitivity: " + "; ".join(f"{x['rate']:.2%} -> {_n(x['total'])}" for x in r["sensitivity"]))
    if r["units"]:
        lines.append("units: " + ", ".join(r["units"]))
    return "\n".join(lines)


def _formula_text(t: dict) -> str:
    out = [f"{t['cell']} {t.get('label') or ''} = {_n(t['value'], 6)}", t.get("source") or "(an input: no formula)"]
    if t.get("n_client_reads"):
        out.append(f"Reads {t['n_client_reads']} client-model value(s), e.g. "
                   + "; ".join(f"{x['cell']} = {_n(x['value'], 4)} ({x['source']})" for x in t["client_reads"][:12]))
    return "\n".join(out)[:12000]


def row_chart(eid: int, rng: str, title: str, feeds: list[str] | None = None, changes: dict | None = None,
              kind: str | None = None, components: list[str] | None = None, component_feed: str | None = None) -> dict:
    """A chart spec for one overlay row: the values Excel saved, the row on each feed, and with changes (on the
    last feed); with components, the rows that add up to it as stacked columns under those lines (a combo)."""
    import chartdata
    import engagement
    import overlay as ovmod
    import rodb
    import tools
    feeds = [f for f in dict.fromkeys(feeds or []) if f != "workbook"] or ["workbook"]
    for f in feeds:
        sess, summary, clean = engagement._live(eid, f, changes)
    path = summary["wiring"]["overlay"]["db_path"]
    comps = [c for c in (components or []) if c][:6]
    kind = kind or ("combo" if comps else "line")
    with tools.using(path):
        spec = tools.chart(title, [{"range": rng, "name": "As saved in Excel"}], kind=kind)
    cols = spec.get("columns")
    if not cols:
        raise ValueError(f"{rng}: give one row of the overlay, e.g. DCF!C7:V7")
    sheet, row, _ = ovmod.parse_a1(rng.split(":")[0])
    keys = [(sheet, row, c) for c in cols]

    def run():
        out = []
        for f in feeds:
            defaults, _, months = ovmod._feed(summary, f, None, None)
            if f != "workbook":
                sess.configure(f, defaults, months)
                out.append((ovmod.FEED_WORDS[f].capitalize(), sess.values(keys)))
            if clean and f == feeds[-1]:
                sess.configure(f, {**defaults, **{ovmod.parse_a1(k): v for k, v in clean.items()}}, months)
                out.append(("With changes" + ("" if f == "workbook" else f" ({ovmod.FEED_WORDS[f]})"), sess.values(keys)))
        parts = []
        if comps:  # the rows that add up to it, on one feed, over the same periods
            cf = component_feed if component_feed in ("workbook", "prior", "current") else feeds[-1]
            defaults, _, months = ovmod._feed(summary, cf, None, None)
            sess.configure(cf, defaults, months)
            for c in comps:
                s_, r_, _ = ovmod.parse_a1(c.split(":")[0])
                parts.append((s_, r_, sess.values([(s_, r_, col) for col in cols]), cf))
        sess.configure("workbook")
        return out, parts

    lines, parts = ovmod.deep(run)
    for name, vals in lines:
        spec["series"].append({"name": name, "range": f"{rng} (Python overlay)", "label": spec["series"][0].get("label"),
                               "units": spec["series"][0].get("units"),
                               "data": [v if isinstance(v, float) else None for v in vals]})
    if parts:
        db = rodb.connect(path)
        for x in spec["series"]:
            x["as"] = "line"
        for s_, r_, vals, cf in parts:
            lab = (db.execute("SELECT label FROM rows WHERE sheet=? AND row=?", (s_, r_)).fetchone() or [None])[0]
            spec["series"].append({"name": lab or f"{s_}!r{r_}", "range": f"{s_}!r{r_} (Python overlay, "
                                   f"{ovmod.FEED_WORDS[cf]})", "label": lab, "units": spec["series"][0].get("units"),
                                   "as": "bar", "data": [v if isinstance(v, float) else None for v in vals]})
    notes = []
    # "As saved in Excel" and the prior client model are the same numbers when the overlay's saved link values are
    # current and the Python reproduces the workbook (the overlay's own check): then draw them once. Where they
    # differ, keep both and say where, because that is the check failing.
    saved = spec["series"][0]
    prior = next((x for x in spec["series"][1:] if x["name"] == ovmod.FEED_WORDS["prior"].capitalize()), None)
    if prior:
        same = lambda a, b: abs((a or 0.0) - (b or 0.0)) <= 1e-6 * max(1.0, abs(b or 0.0))
        off = [i for i, (a, b) in enumerate(zip(prior["data"], saved["data"])) if not same(a, b)]
        lab = spec.get("period_labels") or spec["labels"]
        if not off:
            spec["series"].remove(saved)
            prior["name"] = "The prior client model (= as saved in Excel)"
            notes.append("The prior client model, recomputed in Python, gives the values Excel saved in every period, "
                         "so they are drawn as one line.")
        else:
            notes.append(f"The prior client model, recomputed in Python, differs from the values Excel saved in "
                         f"{len(off)} of {len(lab)} periods (first {lab[off[0]]}): the overlay's saved link values may be "
                         f"out of date, or the Python differs there. Both lines are drawn.")
    # chart rule F1: a last period more than 5x the next largest (a terminal value) flattens the rest, so leave it
    # out of the default view and say so ("Full range" shows it)
    size = [max((abs(x["data"][i]) for x in spec["series"] if isinstance(x["data"][i], float)), default=0.0)
            for i in range(len(cols))]
    order = sorted(range(len(size)), key=lambda i: -size[i])
    if len(order) > 2 and size[order[1]] and size[order[0]] > 5 * size[order[1]] and order[0] >= len(size) - 2 and order[0] > 0:
        spec["view"] = {"x_start": 0, "x_end": order[0] - 1}
        notes.append(f"The last period ({size[order[0]] / size[order[1]]:.0f}x the next largest, typically the terminal "
                     f"value) is left out of this view so the rest is readable; \u201cFull range\u201d shows it.")
    if "current" in feeds:
        notes.append("Periods are labelled with last year's dates; the rolled-forward line's periods each move on by the roll.")
    if notes:
        spec["note"] = " ".join(notes)
    with engagement._fy_hint(eid):  # a quarterly model's financial years: the engagement's, not the calendar's
        return chartdata.enrich(spec, rodb.connect(path))


def extra_tool(eid: int):
    """The overlay tools for agent.ask(extra=...): (text for the model, [UI events]) or None for other tools."""
    import engagement
    import tools

    def run(name: str, a: dict):
        if name == "overlay_run":
            res = engagement.overlay_run(eid, a.get("feed") or "workbook", _changes(a.get("changes")),
                                         a.get("valuation_date"), a.get("months"))
            return _run_text(res), [{"type": "overlay_run", "result": res}]
        if name == "overlay_dcf":
            cut = a.get("cutoff") or "model"
            r = engagement.overlay_dcf(eid, a.get("feed") or "workbook", _changes(a.get("changes")), None, None,
                                       cell=a.get("cell"), rate=a.get("rate"), timing=a.get("timing"),
                                       day_count=a.get("day_count"), cutoff="" if cut == "none" else cut)
            drop = [s.lower() for s in a.get("drop_bridge") or []]
            if drop and r.get("selected"):
                keep = [not any(d in b["label"].lower() for d in drop) for b in r["live"]["bridge"]]
                r = engagement.overlay_dcf(eid, a.get("feed") or "workbook", _changes(a.get("changes")), None, None,
                                           cell=r["selected"], rate=a.get("rate"), timing=a.get("timing"),
                                           day_count=a.get("day_count"), cutoff="" if cut == "none" else cut, include=keep)
            return _dcf_text(r), ([{"type": "overlay_dcf", "result": {k: v for k, v in r.items() if k != "chart"}},
                                   {"type": "chart", "spec": r["chart"]}] if r.get("selected") else [])
        if name == "overlay_chart":
            spec = row_chart(eid, a["range"], a.get("title") or a["range"], a.get("feeds") or [],
                             _changes(a.get("changes")), a.get("kind"), a.get("components"), a.get("component_feed"))
            return tools.chart_note(spec), [{"type": "chart", "spec": spec}]
        if name == "overlay_formula":
            return _formula_text(engagement.overlay_trace(eid, a["cell"])), []
        if name == "overlay_value":
            r = engagement.overlay_value_trace(eid, a.get("cell"))
            if not r.get("selected"):
                return "No report conclusion is matched to an overlay cell that ties to the report; pass a cell.", []
            others = [f"{x['cell']} {x.get('label') or ''} (report {x.get('report') or '-'})" for x in r["starts"]
                      if x["cell"] != r["selected"]]
            return (f"Traced from {r['selected']}" + (f", the cell the report's {r['start'].get('report')} was matched to"
                                                      if r["start"].get("report") else "") + ":\n" + r["text"]
                    + (f"\nOther report figures to trace: {'; '.join(others)}" if others else ""))[:14000], []
        if name == "overlay_inputs":
            found = engagement.overlay_inputs(eid, a.get("text") or "")
            return ("\n".join(f"{x['cell']} {x['label']} = {_n(x['value'], 6)}" for x in found[:60])
                    or "No numeric inputs match."), []
        return None
    return run


def context(eid: int) -> tuple[str, dict[str, str]]:
    """(the engagement for the instructions, {workbook key: model.db path} for the workbook tools)."""
    import engagement
    e = engagement.get(eid)
    _, summary = engagement.overlay_session(eid)
    dbs, lines = {}, [f"Engagement: {e['name']}", "Workbooks (workbook= argument of the workbook tools):"]
    for key in ROLE_KEYS:
        w = engagement._role_wb(eid, key)
        if w:
            dbs[key] = w["db_path"]
            sheets = f", sheets {', '.join(sorted(w['sheets']))}" if w["sheets"] else ""
            lines.append(f"- {key}: {w['filename']}{sheets}" + (" (the default)" if key == "prior_overlay" else ""))
    ref = engagement.reference(eid)
    if ref:
        lines.append("Report key facts" + ("" if all(f["approved"] for f in ref) else " (not all approved yet)") + ":")
        for f in ref[:60]:
            rng = f" (range {f['low_text']} - {f['high_text']})" if f.get("low_text") or f.get("high_text") else ""
            lines.append(f"- {f['label'] or f['key']}: {f['value_text']}{rng}" + (f", {f['basis']}" if f.get("basis") else "")
                         + (f", p.{f['page']}" if f.get("page") else ""))
    v = summary["validation"]
    lines.append(f"The Python overlay: sheets {', '.join(summary['sheets'])}; {v['matched']:,} of {v['cells']:,} formula "
                 f"cells reproduce Excel's saved values.")
    lines.append("Outputs: " + "; ".join(f"{o['cell']} {o['label']} = {_n(o['value'])}" + (f" (report {o['report']})" if o.get("report") else "")
                                         for o in summary["outputs"][:20]))
    if summary["levers"]:
        lines.append("Levers: " + "; ".join(f"{l['cell']} {l['label']} = {l['value']}" + (f" (report {l['report']})" if l.get("report") else "")
                                            for l in summary["levers"]))
    w = summary["wiring"]
    feeds = ["workbook (as saved)"] + (["prior (the prior client model)"] if w.get("prior") else [])
    roll = summary.get("roll")
    if w.get("current") and roll:
        feeds.append(f"current (this year's model, rolled forward {roll['months']} months to {roll['current_valuation_date']}"
                     + (f"; valuation date cell {roll['valuation_date_cell']}" if roll.get("valuation_date_cell") else "") + ")")
    lines.append("Feeds: " + "; ".join(feeds))
    m = e.get("map") or {}
    if m.get("chain"):
        lines.append("Overlay rows reading the client model (overlay <- prior client -> this year's row): " + "; ".join(
            f"{c['overlay']} <- {c['client']} -> {c.get('current') or 'not found'}" for c in m["chain"][:20]))
    return "\n".join(lines), dbs


def ask(eid: int, question: str, model: str, history: list | None = None, session: str | None = None,
        on_usage=None, interactive: bool = False):
    """agent.ask on the overlay workbook, with the other workbooks and the overlay tools."""
    ctx, dbs = context(eid)
    keys = list(dbs)
    specs = copy.deepcopy(agent.TOOLS)
    for t in specs:
        if t["name"] in WORKBOOK_TOOLS and len(keys) > 1:
            t["parameters"]["properties"]["workbook"] = {"type": "string", "enum": keys,
                                                         "description": "which workbook to read (default prior_overlay)"}
    default = dbs.get("prior_overlay") or next(iter(dbs.values()))

    def db_for(args):
        return dbs.get(args.pop("workbook", None) or "prior_overlay", default)

    yield from agent.ask(question, default, model, history, interactive=interactive, context=ctx, on_usage=on_usage,
                         file_id=None, session=session, system=SYSTEM + agent.SYSTEM, tool_specs=specs + OVERLAY_TOOLS,
                         extra=extra_tool(eid), db_for=db_for, log={"engagement": eid, "step": "chat"})


if __name__ == "__main__":
    import sys
    from pathlib import Path
    sys.path.insert(0, str(Path(__file__).parent))
    for ev in ask(int(sys.argv[1]), sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else "gpt-6-luna"):
        if ev["type"] in ("tool_call", "answer", "error", "usage"):
            print(json.dumps(ev, default=str)[:600])
