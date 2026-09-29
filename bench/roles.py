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

The structure comes first (likeness.py), so there is a suggestion as soon as the files are read, before the
report: workbooks that are mostly the same are one client model twice; one that contains another plus extra
sheets carrying valuation work is a client model with the overlay added (the extra sheets are the overlay); one
unlike the others, with valuation vocabulary, charts, external links or the adviser's name, is a standalone
overlay. The report's facts, once extracted and approved, add the stronger evidence above, and the suggestion is
redone each time they change, until a person confirms.
"""
import json
import re
import sqlite3

import likeness
import linkmap
import rodb

FAMILY = 50          # % alike: the same model (last year's and this year's, or one with the overlay added)
CONTAINS = 80        # % of one workbook's line items found in another: the other is it plus more
OVERLAY_POINTS = 4   # structure points for a workbook unlike the others to count as a standalone overlay

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


def structure(workbooks: list[dict]) -> dict:
    """The likeness evidence: {"sigs", "pairs": {(a, b): compare}, "shape": {id: {...}}} where shape says, per
    workbook, its family (the same model), the extra sheets it has over a family member it contains, the points
    that make it look like a standalone overlay, and why."""
    sigs = {w["id"]: likeness.signature(w) for w in workbooks}
    ids = list(sigs)
    pairs = {}
    for i, a in enumerate(ids):
        for b in ids[i + 1:]:
            pairs[(a, b)] = likeness.compare(sigs[a], sigs[b])

    def pair(a, b):  # a's view: a_in_b = share of a's line items found in b
        if (a, b) in pairs:
            return pairs[(a, b)]
        c = pairs[(b, a)]
        return {**c, "a_in_b": c["b_in_a"], "b_in_a": c["a_in_b"], "only_a": c["only_b"], "only_b": c["only_a"]}

    # versions of one model: alike overall (last year's and this year's client model, or one with the overlay added)
    parent = {i: i for i in ids}
    find = lambda i: i if parent[i] == i else find(parent[i])
    for (a, b), c in pairs.items():
        if c["similarity"] >= FAMILY:
            parent[find(a)] = find(b)
    names = {w["id"]: w["filename"] for w in workbooks}
    shape = {}
    for i in ids:
        sig = sigs[i]
        fam = [j for j in ids if j != i and find(j) == find(i)]
        why, extra, points = [], {}, 0
        for j in sorted(fam, key=lambda j: -pair(i, j)["similarity"]):
            c = pair(i, j)
            why.append(f"{c['similarity']}% alike to {names[j]} (sheets {c['sheets']}%, line items {c['line_items']}%, "
                       f"formulas {c['formulas']}%)")
        # a workbook holding all of a smaller one plus more sheets: the overlay added to a copy of the client model.
        # The host is the workbook it holds most completely (this year's model is held nearly as well, being so
        # alike; a standalone overlay's sheets may be held too), the bigger one on a tie.
        held = [j for j in ids if j != i and pair(j, i)["a_in_b"] >= CONTAINS and len(sigs[j]["sheets"]) < len(sig["sheets"])]
        host = max(held, key=lambda j: (pair(j, i)["a_in_b"], sigs[j]["line_items"]), default=None)  # the fullest copy
        if host is not None:
            others = set(sigs[host]["sheet_keys"])
            for k, s in sig["sheet_keys"].items():
                ps = sig["per_sheet"].get(s, {})
                if k in others or not ps.get("rows"):
                    continue
                reasons = []
                if len(ps.get("terms", [])) >= 2:
                    reasons.append("valuation terms: " + ", ".join(ps["terms"][:5]))
                if sig["charts"].get(s):
                    reasons.append(f"{sig['charts'][s]} chart(s)")
                extra[s] = reasons
            if extra:
                why.append(f"contains {pair(host, i)['a_in_b']}% of {names[host]}'s line items, plus sheet(s) it doesn't "
                           f"have: {', '.join(extra)}")
        if not fam and not extra:  # unlike every other workbook: how much does it look like valuation work?
            n_terms = len(sig["terms"])
            points = n_terms // 2 + (2 if sig["charts"] else 0) + (2 if sig["external_links"] else 0) + \
                (4 if sig["markers"] else 0)
            others = [sigs[j] for j in ids if j != i]
            if others and len(sig["sheets"]) < min(len(o["sheets"]) for o in others):
                points += 1
            best = max((pair(i, j)["similarity"] for j in ids if j != i), default=0)
            why.append(f"unlike the other workbooks (at most {best}% alike)")
            if n_terms:
                why.append(f"valuation terms ({n_terms}): " + ", ".join(list(sig["terms"])[:6]))
            if sig["charts"]:
                why.append(f"{sum(sig['charts'].values())} chart(s) on " + ", ".join(sig["charts"]))
            if sig["external_links"]:
                why.append("links to " + ", ".join(sig["external_links"][:3]))
        if sig["markers"]:
            why.append("mentions " + ", ".join(f"{m} ({w})" for m, w in list(sig["markers"].items())[:3]))
        shape[i] = {"family": sorted(fam), "extra": extra, "points": points, "why": why,
                    "standalone": not fam and not extra and points >= OVERLAY_POINTS,
                    "host": host if extra else None, "host_share": pair(host, i)["a_in_b"] if extra else None}
    return {"sigs": sigs, "pairs": pairs, "shape": shape}


def split(prof: dict, seed: dict[str, list[str]] | None = None, shared: set[str] | None = None) -> tuple[list[str], list[str]]:
    """(overlay sheets, client sheets), with the reasons written into each overlay sheet's "why". seed: sheets the
    structure already marks as overlay (the extra sheets over a copy of the model), with reasons. shared: sheets
    another version of the same model also has, which can't be last year's overlay unless the report's figures
    sit there (a client model's own DCF is not the overlay)."""
    sh = prof["sheets"]
    overlay = set()
    for s, reasons in (seed or {}).items():
        if s in sh:
            sh[s]["why"].append("not in the other version of the model" + (": " + "; ".join(reasons) if reasons else ""))
            overlay.add(s)
    shared = shared or set()
    for s, v in sh.items():
        if v["facts"]:
            v["why"].append("holds " + "; ".join(v["facts"][:3]))
        if v["anchors"]:
            v["why"].append(f"DCF reproduced in Python: {', '.join(a['cell'] for a in v['anchors'][:2])}")
        if (v["facts"] or v["anchors"]) and s not in shared:
            overlay.add(s)
        elif v["facts"] or v["anchors"]:
            # a report figure can match the client's own inputs (a cost of equity on its assumptions sheet): a
            # sheet the other version of the model has too is the client's, whatever it holds
            v["why"].append("the other version of the model has this sheet too, so it's the client's own")
    seeds = set(overlay)
    grew = True
    while grew:  # sheets built on the valuation (summaries, outputs) belong to it too, but not the client's own
        grew = False
        for s, v in sh.items():
            if s not in overlay and s not in shared and v["reads"] & overlay and v["formulas"]:
                v["why"].append("reads " + ", ".join(sorted(v["reads"] & overlay)))
                overlay.add(s)
                grew = True
    total = sum(v["formulas"] for v in sh.values()) or 1
    if len(overlay) > len(seeds) and sum(sh[s]["formulas"] for s in overlay) > 0.6 * total:
        # the spread took in most of the workbook: a hub the whole model reads was taken for the valuation
        for s in overlay - seeds:
            sh[s]["why"].append("left out: following what reads the valuation took in most of the workbook")
        overlay = seeds
    for s in overlay:
        if sh[s]["ext_rows"]:
            sh[s]["why"].append(f"{sh[s]['ext_rows']} line items read other workbooks (external links)")
        if NAME_HINT.search(s):
            sh[s]["why"].append("sheet name")
    client = [s for s, v in sh.items() if s not in overlay and v["rows"]]
    return sorted(overlay, key=list(sh).index), sorted(client, key=list(sh).index)


SAME_SHEET = 0.35  # a sheet of the same name in another version is the same sheet if it shares this much of its labels


def _shared(sig: dict, family: list[dict], sheet: str) -> bool:
    """Another version of the model has this sheet: one of the same name with much the same line items. A name
    alone isn't enough: an overlay's "Inputs" or "Summary" isn't the client's sheet of that name."""
    ns = likeness.norm(sheet) or sheet.lower()
    mine = {lab for (s, lab) in sig["items"] if s == ns}
    for other in family:
        if ns not in other["sheet_keys"]:
            continue
        theirs = {lab for (s, lab) in other["items"] if s == ns}
        if not mine and not theirs:
            return True
        if mine and theirs and len(mine & theirs) / len(mine | theirs) >= SAME_SHEET:
            return True
    return False


def suggest(reports: list[dict], workbooks: list[dict], facts: list[dict]) -> dict:
    """reports: [{"id", "filename", "n_facts"}]; workbooks: [{"id", "filename", "db_path", "source_path",
    "valuation_date", "uploaded_at"}]; facts: approved (or all) report facts.
    Returns {"roles": {role: {"kind", "id", "sheets", "why"}}, "workbooks": {id: {"mode", "overlay", "client"}}}."""
    roles, info = {}, {}
    st = structure(workbooks)
    if reports:
        r = max(reports, key=lambda r: (r.get("n_facts") or 0, r["id"]))
        roles["prior_report"] = {"kind": "document", "id": r["id"], "sheets": None,
                                 "why": ["the report" if len(reports) == 1 else f"has {r.get('n_facts') or 0} extracted facts"]}
    profs = {}
    for wb in workbooks:
        fm = linkmap.match_facts(wb["db_path"], facts) if facts else []
        prof = profile(wb, fm)
        sp = st["shape"][wb["id"]]
        sig = st["sigs"][wb["id"]]
        shared = {s for s in prof["sheets"] if _shared(sig, [st["sigs"][j] for j in sp["family"]], s)}
        # extra sheets count as the overlay only with some sign of valuation work (a new client sheet isn't)
        seed = {s: r for s, r in sp["extra"].items() if r or prof["sheets"].get(s, {}).get("anchors")}
        ov, cl = split(prof, seed, shared)
        if sp["standalone"] and len(workbooks) > 1:  # a separate valuation workbook: every sheet is overlay work
            for s, v in prof["sheets"].items():
                if v["rows"] and s not in ov:
                    v["why"].append("in the separate valuation workbook")
            ov = [s for s, v in prof["sheets"].items() if v["rows"] or s in ov]
            cl = []
        total = sum(v["formulas"] for v in prof["sheets"].values()) or 1
        client_share = sum(prof["sheets"][s]["formulas"] for s in cl) / total
        mode = "client model" if not ov else ("overlay inside the client model" if cl and client_share >= 0.2
                                              else "standalone overlay")
        facts_on = sum(len(prof["sheets"][s]["facts"]) for s in ov)
        score = sum(len(prof["sheets"][s]["facts"]) * 2 + len(prof["sheets"][s]["anchors"]) for s in ov) \
            + 2 * len(seed) + (4 if sig["markers"] else 0)
        info[wb["id"]] = {"mode": mode, "overlay": ov, "client": cl if mode != "standalone overlay" else [],
                          "score": score, "points": sp["points"] if mode == "standalone overlay" else 0,
                          "facts_on": facts_on, "timeline_start": prof["timeline_start"],
                          "client_share": round(client_share, 2),
                          "why": {s: prof["sheets"][s]["why"] for s in ov}, "structure": sp["why"],
                          "family": sp["family"]}
        profs[wb["id"]] = prof
    # A standalone overlay's structure points (valuation words, charts, fewest sheets) count only when no other
    # workbook has the report's figures on its overlay sheets: a large client model has the words and the charts
    # too, and report figures are the stronger evidence.
    figures_elsewhere = lambda i: any(j != i and x["mode"] != "standalone overlay" and x["facts_on"] for j, x in info.items())
    for i, x in info.items():
        if x["points"] and not figures_elsewhere(i):
            x["score"] += x["points"]
    by_id = {wb["id"]: wb for wb in workbooks}
    ov_id = max((i for i in info if info[i]["overlay"]), key=lambda i: info[i]["score"], default=None)
    names = {w["id"]: w["filename"] for w in workbooks}
    # one overlay per engagement: another workbook typed a standalone overlay only by its structure (unlike the
    # others, with valuation words and charts, as a rebuilt client model is) is a client model
    for i, x in info.items():
        if ov_id is not None and i != ov_id and x["mode"] == "standalone overlay":
            x.update(mode="client model", overlay=[], why={},
                     client=[s for s, v in profs[i]["sheets"].items() if v["rows"]])
            x["structure"] = x["structure"] + [f"typed as a client model: the engagement's overlay is {names[ov_id]}"]
    when = {w["id"]: _when(w, info[w["id"]]["timeline_start"]) for w in workbooks}
    prior_id = None
    if ov_id is not None:
        o = info[ov_id]
        why = [f"{s}: {'; '.join(o['why'][s])}" for s in o["overlay"] if o["why"][s]]
        if o["mode"] == "standalone overlay":
            why = o["structure"] + why
        if FILE_OVERLAY.search(names[ov_id]):
            why.append("its file name says so")
        roles["prior_overlay"] = {"kind": "workbook", "id": ov_id, "sheets": o["overlay"], "why": why}
        host = st["shape"][ov_id].get("host")
        if o["mode"] == "overlay inside the client model" and host is not None and info[host]["mode"] == "client model":
            # the overlay was added to a copy of a client model that's here too: that file is the client's own. Its
            # valuation date can differ (an adviser builds on the model the client sent, then values at a later
            # date), so the date explains, it doesn't decide
            prior_id = host
            dates = ([f"the client's model is dated {when[host]['valuation_date']}, the overlay "
                      f"{when[ov_id]['valuation_date']}: the valuation was built on it later"]
                     if when[host]["valuation_date"] and when[ov_id]["valuation_date"]
                     and when[host]["valuation_date"] != when[ov_id]["valuation_date"] else [])
            roles["prior_model"] = {"kind": "workbook", "id": host, "sheets": None, "why": [
                f"the client's model as sent: {names[ov_id]} is a copy of it ({st['shape'][ov_id]['host_share']}% of "
                f"its line items) with the overlay added ({', '.join(o['overlay'])})", *_date_why(when[host]), *dates]}
        elif o["mode"] == "overlay inside the client model":
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
    rest.sort(key=lambda w: when[w["id"]]["key"])  # valuation date, then the date in the file name, then the timeline
    if prior_id is not None:  # the current model is another version of the prior one, if there is such a version
        fam = [w for w in rest if w["id"] in info[prior_id]["family"] and w["id"] != prior_id
               and not _same_period(when[w["id"]], when[prior_id])]
        rest = fam or [w for w in rest if w["id"] != prior_id]
    if prior_id is None and len(rest) >= 2:
        prior_id = rest[0]["id"]
        roles["prior_model"] = {"kind": "workbook", "id": prior_id, "sheets": None,
                                "why": ["the earlier of the client models", *_date_why(when[prior_id])]
                                + info[prior_id]["structure"][:1]}
    current = [w for w in rest if w["id"] != prior_id]
    if current:
        c = current[-1]
        why = ["the latest of the client models", *_date_why(when[c["id"]], when.get(prior_id))]
        roles["current_model"] = {"kind": "workbook", "id": c["id"], "sheets": None, "why": why + info[c["id"]]["structure"][:1]}
    checks = _checks(roles, info, st, when, names, facts)
    pairs = [{"a": a, "b": b, **{k: v for k, v in c.items()}} for (a, b), c in st["pairs"].items()]
    return {"roles": roles, "workbooks": {i: {k: v for k, v in x.items() if k not in ("score", "points", "facts_on")}
                                          for i, x in info.items()},
            "likeness": {"pairs": pairs, "profiles": {i: likeness.public(sg) for i, sg in st["sigs"].items()}},
            "dates": {i: {k: v for k, v in w.items() if k != "key"} for i, w in when.items()}, "checks": checks,
            "evidence": f"structure and {len(facts)} report fact(s)" if facts else "structure only"}


# ---- dates: which version is earlier ------------------------------------------------------------------------------
# The identified valuation date says it best; the date in the file name ("20250523 ...", "... Jun 25 ...", "FY26")
# next; the timeline's first period last (a model with actuals from years back starts on the same date every year).
FILE_OVERLAY = re.compile(r"overlay|valuation|\bval\b|\bdcf\b", re.I)
_MONTHS = {m: i for i, m in enumerate("jan feb mar apr may jun jul aug sep oct nov dec".split(), 1)}


def file_date(name: str) -> str | None:
    """The date a file name carries, as YYYY-MM-DD (the day is 1 when only a month is given)."""
    n = name or ""
    m = re.search(r"(?<!\d)(20\d{2})[-_. ]?(0[1-9]|1[0-2])[-_. ]?(0[1-9]|[12]\d|3[01])(?!\d)", n)
    if m:
        return f"{m[1]}-{m[2]}-{m[3]}"
    m = re.search(r"(?<![A-Za-z])(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*[-_ ']*(20)?(\d{2})(?!\d)", n, re.I)
    if m:
        return f"20{m[3]}-{_MONTHS[m[1].lower()]:02d}-01"
    m = re.search(r"(?<![A-Za-z])(?:FY|BP)[-_ ]?(20)?(\d{2})(?!\d)", n, re.I)  # financial year, business plan
    return f"20{m[2]}-06-30" if m else None


def file_version(name: str) -> str | None:
    m = re.search(r"\bv(\d+(?:\.\d+)*)\b", name or "", re.I)
    return m[1] if m else None


def _when(w: dict, timeline_start: str | None) -> dict:
    vd, fd = w.get("valuation_date"), file_date(w.get("filename"))
    return {"valuation_date": vd, "file_date": fd, "version": file_version(w.get("filename")),
            "timeline_start": timeline_start, "key": (vd or "", fd or "", timeline_start or "", w.get("uploaded_at") or 0)}


def _same_period(a: dict, b: dict) -> bool:
    """Same valuation (same valuation date, or file dates within four months when there's none)."""
    if a["valuation_date"] and b["valuation_date"]:
        return a["valuation_date"] == b["valuation_date"]
    if a["file_date"] and b["file_date"]:
        from datetime import date
        da, db_ = date.fromisoformat(a["file_date"]), date.fromisoformat(b["file_date"])
        return abs((da - db_).days) <= 120
    return a["timeline_start"] == b["timeline_start"]


def _date_why(w: dict, other: dict | None = None) -> list[str]:
    out = []
    if w["valuation_date"]:
        out.append(f"valuation date {w['valuation_date']}" + (f" (the prior's is {other['valuation_date']})"
                                                              if other and other.get("valuation_date") else ""))
    if w["file_date"]:
        out.append(f"its file name is dated {w['file_date']}" + (f" (the prior's {other['file_date']})"
                                                                 if other and other.get("file_date") else ""))
    if w["version"]:
        out.append(f"version {w['version']}")
    return out


# ---- checks on a suggestion ------------------------------------------------------------------------------------------

def _checks(roles: dict, info: dict, st: dict, when: dict, names: dict, facts: list[dict]) -> list[dict]:
    """Plain checks on the assignment, each ok (True), a warning (False) or can't tell yet (None)."""
    out = []
    add = lambda ok, text: out.append({"ok": ok, "text": text})
    wid = lambda role: (roles.get(role) or {}).get("id") if (roles.get(role) or {}).get("kind") == "workbook" else None
    ov, pr, cu = wid("prior_overlay"), wid("prior_model"), wid("current_model")
    pair = lambda a, b: st["pairs"].get((a, b)) or st["pairs"].get((b, a))
    if pr and cu:
        c = pair(pr, cu)
        if c:
            add(c["similarity"] >= FAMILY, f"the prior and current client models are {c['similarity']}% alike"
                + (" (versions of one model)" if c["similarity"] >= FAMILY else ": are they the same model?"))
        a, b = when[pr], when[cu]
        if a["valuation_date"] and b["valuation_date"]:
            add(b["valuation_date"] > a["valuation_date"], f"the current model's valuation date ({b['valuation_date']}) "
                f"is {'after' if b['valuation_date'] > a['valuation_date'] else 'not after'} the prior's ({a['valuation_date']})")
        elif a["file_date"] and b["file_date"]:
            add(b["file_date"] > a["file_date"], f"the current model's file date ({b['file_date']}) is "
                f"{'after' if b['file_date'] > a['file_date'] else 'not after'} the prior's ({a['file_date']})")
        else:
            add(None, "no valuation date or file date to tell the prior and current models apart; check the order")
    elif not cu:
        add(False, "no current client model found")
    if ov:
        o = info[ov]
        if o["mode"] == "standalone overlay" and pr:
            add(True if roles["prior_model"]["why"] and "external link" in roles["prior_model"]["why"][0] else None,
                "the overlay reads the prior client model through an external link"
                if roles["prior_model"]["why"] and "external link" in roles["prior_model"]["why"][0]
                else "no external link from the overlay to the prior client model found")
        elif pr and pr != ov:
            add(True, f"the overlay sits in a copy of the prior client model ({names[ov]})")
        elif pr == ov:
            add(True, "the overlay sits inside the prior client model")
        vo = when[ov]["valuation_date"]
        rep_vd = next((f for f in facts if f.get("key") == "valuation_date" and f.get("value")), None)
        if rep_vd and vo:
            v = int(rep_vd["value"])
            iso = f"{v // 10000:04d}-{v // 100 % 100:02d}-{v % 100:02d}"
            add(iso == vo, f"the report's valuation date ({iso}) {'matches' if iso == vo else 'differs from'} the overlay "
                f"workbook's ({vo})")
        elif not facts:
            add(None, "the report's facts aren't in yet: they will confirm the overlay by where its figures sit")
        n_fig = sum(1 for s in o["overlay"] for w in (o["why"].get(s) or []) if w.startswith("holds "))
        if facts:
            add(n_fig > 0 or None, f"the report's figures sit on {n_fig} of the overlay's sheets" if n_fig
                else "none of the report's figures were found on the overlay sheets")
    else:
        add(False, "no overlay found: pick the overlay workbook and its sheets")
    return out


# ---- a second opinion ------------------------------------------------------------------------------------------------
# The rules above are explicit but can miss what a person sees at a glance (a file name, a sheet called "Overlay",
# a cover sheet's version note). A model looks at the same evidence and the suggestion, and agrees or says what it
# would assign instead, with its reasons. It never assigns anything itself: the page shows it beside the rules'
# suggestion for the person to decide.
SECOND_PROMPT = """You check which file plays which part in a recurring valuation engagement:
- prior_report: last year's final valuation report (PDF / PPTX)
- prior_model: the client's model behind last year's valuation, as the client sent it
- prior_overlay: our valuation workings (the "overlay") that take the client model to the report's conclusions:
  a separate workbook, or sheets added to a copy of the client model (then name the workbook and those sheets)
- current_model: this year's client model, onto which the valuation is rolled forward

Use the evidence: file names (dates, versions, words like overlay or valuation), identified valuation dates,
how alike the workbooks are (versions of one model are mostly alike; a copy with extra valuation sheets holds
the overlay), where the report's figures sit, external links, charts and valuation vocabulary. Say whether you
agree with the suggested assignment. If not, give yours (file names exactly as listed, "" for none) and why.

Evidence:
{evidence}

Suggested assignment:
{suggestion}"""
_S = {"type": "string"}
SECOND_SCHEMA = {"type": "json_schema", "name": "roles_second_opinion", "strict": True, "schema": {
    "type": "object", "additionalProperties": False,
    "required": ["agree", "prior_report", "prior_model", "prior_overlay", "overlay_sheets", "current_model", "confidence", "reasons"],
    "properties": {"agree": {"type": "boolean"}, "prior_report": _S, "prior_model": _S, "prior_overlay": _S,
                   "overlay_sheets": {"type": "array", "items": _S}, "current_model": _S,
                   "confidence": {"type": "string", "enum": ["high", "medium", "low"]},
                   "reasons": {"type": "array", "items": _S}}}}


def evidence(reports: list[dict], workbooks: list[dict], facts: list[dict], res: dict) -> dict:
    """What the second opinion sees: each file with what's known about it, how alike the workbooks are, the report."""
    names = {w["id"]: w["filename"] for w in workbooks}
    prof = res["likeness"]["profiles"]
    out = {"reports": [{"file": r["filename"], "facts_extracted": r.get("n_facts") or 0} for r in reports], "workbooks": []}
    for w in workbooks:
        p, i, d = prof.get(w["id"]) or {}, res["workbooks"].get(w["id"]) or {}, res["dates"].get(w["id"]) or {}
        out["workbooks"].append({
            "file": w["filename"], "sheets": (p.get("sheets") or [])[:80], "formula_cells": p.get("formulas"),
            "identified_target": w.get("target_name"), "identified_valuation_date": w.get("valuation_date"),
            "date_in_file_name": d.get("file_date"), "version_in_file_name": d.get("version"),
            "timeline_starts": d.get("timeline_start"), "valuation_terms": list((p.get("terms") or {}).keys()),
            "charts": p.get("charts"), "external_links_to": p.get("external_links"),
            "author": (p.get("props") or {}).get("creator"), "company": (p.get("props") or {}).get("company"),
            "rules_say": i.get("mode"), "overlay_sheets_by_rules": i.get("overlay"),
            "evidence_by_rules": {s: why for s, why in (i.get("why") or {}).items()}})
    out["alike"] = [{"a": names[p["a"]], "b": names[p["b"]], "similarity_pct": p["similarity"],
                     "share_of_a_in_b_pct": p["a_in_b"], "share_of_b_in_a_pct": p["b_in_a"],
                     "sheets_only_in_a": p["only_a"][:12], "sheets_only_in_b": p["only_b"][:12]}
                    for p in res["likeness"]["pairs"] if p["a"] in names and p["b"] in names]
    ident = {f["key"]: f["value_text"] for f in facts if f.get("category") in ("identity", "conclusion")}
    out["report_facts"] = ident
    return out


def second_opinion(reports: list[dict], workbooks: list[dict], facts: list[dict], res: dict, model: str,
                   on_usage=None) -> dict:
    from llm import client, create
    names = {w["id"]: w["filename"] for w in workbooks}
    docs = {r["id"]: r["filename"] for r in reports}
    sug = {}
    for role, r in res["roles"].items():
        sug[role] = (docs if r["kind"] == "document" else names).get(r["id"], "")
        if role == "prior_overlay":
            sug["overlay_sheets"] = r.get("sheets") or []
    prompt = SECOND_PROMPT.format(evidence=json.dumps(evidence(reports, workbooks, facts, res), indent=1, default=str)[:60000],
                                  suggestion=json.dumps(sug, indent=1))
    r = create(client(interactive=False), model, input=prompt, text={"format": SECOND_SCHEMA}, max_output_tokens=3000,
               purpose="roles-review")
    if r.usage and on_usage:
        on_usage(model, r.usage, "roles-review")
    out = json.loads(r.output_text)
    out["model"] = model
    # compare file by file, whatever the model says about agreeing
    out["differs"] = [k for k in ("prior_report", "prior_model", "prior_overlay", "current_model")
                      if (out.get(k) or "") != (sug.get(k) or "")]
    if not out["differs"] and sorted(out.get("overlay_sheets") or []) != sorted(sug.get("overlay_sheets") or []) \
            and out.get("overlay_sheets"):
        out["differs"].append("overlay_sheets")
    return out
