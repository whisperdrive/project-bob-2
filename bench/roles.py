"""Which uploaded file plays which part in a recurring valuation: prior report, prior client model, prior overlay,
current client model. Suggested from the files themselves, with the evidence, for a person to confirm.

The overlay is found at sheet level, because it is sometimes a separate workbook and sometimes a set of sheets
added to a copy of the client model. An overlay sheet:
  - holds the report's conclusions (enterprise / equity value) or valuation-only assumptions (discount rate,
    terminal growth, exit or RAB multiple, valuation date), or a DCF that valuation.py reproduces, or
  - reads such a sheet (summaries and outputs built on the valuation)
Every other sheet with formulas is client model. A workbook with both is "overlay inside the client model";
one with only overlay sheets is a standalone overlay, and its external links name the client model it read.
Prior vs current client model: the file the overlay reads (link name or cached values), then the later timeline.
"""
import json
import re
import sqlite3

import linkmap
import rodb

ROLES = {"prior_report": "Prior report", "prior_model": "Prior client model", "prior_overlay": "Prior overlay",
         "current_model": "Current client model"}
VALUATION_KEYS = re.compile(r"discount|wacc|terminal|exit_multiple|rab|valuation_date|enterprise_value|equity_value|"
                            r"sensitivity|cost_of_equity", re.I)
NAME_HINT = re.compile(r"val|dcf|wacc|overlay|sensitiv", re.I)


def profile(wb: dict, fact_matches: list[dict]) -> dict:
    """Per-sheet evidence for one workbook: {"sheets": {sheet: {...}}, "timeline_start"}."""
    import valuation
    db = rodb.connect(wb["db_path"])
    sheets = {}
    for s, layout in db.execute("SELECT sheet, layout FROM sheets"):
        lay = json.loads(layout or "{}")
        sheets[s] = {"rows": 0, "formulas": 0, "reads": set(), "read_by": set(), "ext_rows": 0, "anchors": [],
                     "facts": [], "start": lay.get("period_start"), "why": []}
    for s, n, f in db.execute("SELECT sheet, COUNT(*), SUM(n_formula) FROM rows GROUP BY sheet"):
        if s in sheets:
            sheets[s].update(rows=n, formulas=f or 0)
    live = ",".join(f"'{k}'" for k in linkmap.LIVE)
    for a, b in db.execute(f"SELECT DISTINCT src_sheet, dst_sheet FROM edges WHERE src_sheet<>dst_sheet AND kind IN ({live})"):
        if a in sheets and b in sheets:
            sheets[a]["reads"].add(b)
            sheets[b]["read_by"].add(a)
    have = {r[0] for r in db.execute("SELECT name FROM sqlite_master WHERE type='table'")}
    if "extrefs" in have:
        for s, n in db.execute("SELECT sheet, COUNT(DISTINCT row) FROM extrefs GROUP BY sheet"):
            if s in sheets:
                sheets[s]["ext_rows"] = n
    db.close()
    try:
        for a in valuation.catalogue(wb["db_path"]):
            if a.get("ok"):
                s = a["cell"].split("!")[0]
                if s in sheets:
                    sheets[s]["anchors"].append(a)
    except Exception:
        pass
    for fm in fact_matches:
        if not VALUATION_KEYS.search(fm["key"] or ""):
            continue
        # A bare number is weak evidence (CPI 2.5% equals a 2.5% terminal growth rate): the cell's label must
        # agree with the fact, or the cell must be a DCF result valuation.py reproduces.
        best = next((m for m in fm["matches"] if m["label_match"] or m["anchor"]), None)
        if best:
            sheets[best["sheet"]]["facts"].append(f"{fm['label'] or fm['key']} {fm['value_text']} ({best['sheet']}!{best['addr']})")
    starts = [v["start"] for v in sheets.values() if v["start"]]
    return {"sheets": sheets, "timeline_start": min(starts) if starts else None}


def split(prof: dict) -> tuple[list[str], list[str]]:
    """(overlay sheets, client sheets), with the reasons written into each overlay sheet's "why"."""
    sh = prof["sheets"]
    overlay = set()
    for s, v in sh.items():
        if v["facts"]:
            v["why"].append("holds " + "; ".join(v["facts"][:3]))
        if v["anchors"]:
            v["why"].append(f"DCF reproduced in Python: {', '.join(a['cell'] for a in v['anchors'][:2])}")
        if v["facts"] or v["anchors"]:
            overlay.add(s)
    grew = True
    while grew:  # sheets built on the valuation (summaries, outputs) belong to it too
        grew = False
        for s, v in sh.items():
            if s not in overlay and v["reads"] & overlay and v["formulas"]:
                v["why"].append("reads " + ", ".join(sorted(v["reads"] & overlay)))
                overlay.add(s)
                grew = True
    for s in overlay:
        if sh[s]["ext_rows"]:
            sh[s]["why"].append(f"{sh[s]['ext_rows']} line items read other workbooks (external links)")
        if NAME_HINT.search(s):
            sh[s]["why"].append("sheet name")
    client = [s for s, v in sh.items() if s not in overlay and v["rows"]]
    return sorted(overlay, key=list(sh).index), sorted(client, key=list(sh).index)


def suggest(reports: list[dict], workbooks: list[dict], facts: list[dict]) -> dict:
    """reports: [{"id", "filename", "n_facts"}]; workbooks: [{"id", "filename", "db_path", "source_path",
    "valuation_date", "uploaded_at"}]; facts: approved (or all) report facts.
    Returns {"roles": {role: {"kind", "id", "sheets", "why"}}, "workbooks": {id: {"mode", "overlay", "client"}}}."""
    roles, info = {}, {}
    if reports:
        r = max(reports, key=lambda r: (r.get("n_facts") or 0, r["id"]))
        roles["prior_report"] = {"kind": "document", "id": r["id"], "sheets": None,
                                 "why": ["the report" if len(reports) == 1 else f"has {r.get('n_facts') or 0} extracted facts"]}
    profs = {}
    for wb in workbooks:
        fm = linkmap.match_facts(wb["db_path"], facts) if facts else []
        prof = profile(wb, fm)
        ov, cl = split(prof)
        total = sum(v["formulas"] for v in prof["sheets"].values()) or 1
        client_share = sum(prof["sheets"][s]["formulas"] for s in cl) / total
        mode = "client model" if not ov else ("overlay inside the client model" if cl and client_share >= 0.2
                                              else "standalone overlay")
        score = sum(len(prof["sheets"][s]["facts"]) * 2 + len(prof["sheets"][s]["anchors"]) for s in ov)
        info[wb["id"]] = {"mode": mode, "overlay": ov, "client": cl if mode != "standalone overlay" else [],
                          "score": score, "timeline_start": prof["timeline_start"], "client_share": round(client_share, 2),
                          "why": {s: prof["sheets"][s]["why"] for s in ov}}
        profs[wb["id"]] = prof
    by_id = {wb["id"]: wb for wb in workbooks}
    ov_id = max((i for i in info if info[i]["overlay"]), key=lambda i: info[i]["score"], default=None)
    prior_id = None
    if ov_id is not None:
        o = info[ov_id]
        why = [f"{s}: {'; '.join(o['why'][s])}" for s in o["overlay"] if o["why"][s]]
        roles["prior_overlay"] = {"kind": "workbook", "id": ov_id, "sheets": o["overlay"], "why": why}
        if o["mode"] == "overlay inside the client model":
            prior_id = ov_id
            roles["prior_model"] = {"kind": "workbook", "id": ov_id, "sheets": o["client"],
                                    "why": [f"same workbook as the overlay; client sheets: {', '.join(o['client'])}"]}
        else:
            # The workbook the overlay's external links read: same file name, or the cached values match.
            best, why = None, []
            others = [w for w in workbooks if w["id"] != ov_id]
            for w in others:
                res = linkmap.overlay_to_client({**by_id[ov_id], "sheets": None}, {**w, "sheets": None})
                hit = next((b for b in res["books"] if b.get("is_client")), None)
                if hit:
                    chk = hit.get("cached_check") or {}
                    rank = (linkmap._norm_file(hit["filename"]) == linkmap._norm_file(w["filename"]), chk.get("matched", 0))
                    if best is None or rank > best[0]:
                        best = (rank, w["id"])
                        why = [f"the overlay's external link [{hit['idx']}] points to {hit['filename']}"]
                        if chk.get("cells"):
                            why.append(f"{chk['matched']} of {chk['cells']} values the overlay last read from it match this file")
            if best:
                prior_id = best[1]
                roles["prior_model"] = {"kind": "workbook", "id": prior_id, "sheets": None, "why": why}
    rest = [w for w in workbooks if w["id"] not in (ov_id, prior_id) or (w["id"] == ov_id and info[ov_id]["mode"] == "client model")]
    rest = [w for w in rest if info[w["id"]]["mode"] == "client model"]
    order = lambda w: (info[w["id"]]["timeline_start"] or "", w.get("valuation_date") or "", w.get("uploaded_at") or 0)
    rest.sort(key=order)
    if prior_id is None and len(rest) >= 2:
        prior_id = rest[0]["id"]
        roles["prior_model"] = {"kind": "workbook", "id": prior_id, "sheets": None,
                                "why": [f"earlier timeline (starts {info[prior_id]['timeline_start'] or 'unknown'})"]}
    current = [w for w in rest if w["id"] != prior_id]
    if current:
        c = current[-1]
        why = [f"client model with the latest timeline (starts {info[c['id']]['timeline_start'] or 'unknown'})"]
        if prior_id is not None and info[prior_id]["timeline_start"] and info[c["id"]]["timeline_start"]:
            why.append(f"prior model's timeline starts {info[prior_id]['timeline_start']}")
        roles["current_model"] = {"kind": "workbook", "id": c["id"], "sheets": None, "why": why}
    return {"roles": roles, "workbooks": {i: {k: v for k, v in x.items() if k != "score"} for i, x in info.items()}}
