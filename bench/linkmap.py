"""The map between the four files of a recurring valuation, built from cell values and formulas (no model calls).

  report  -> overlay         each approved report figure -> overlay cells holding that value (unit scaling
                             and rounding to the printed precision allowed), ranked by label and cell kind
  overlay -> prior client    overlay rows -> the client rows they read: external links ([n]Sheet!A1, with the
                             link's cached values compared to the client workbook) when the overlay is a separate
                             file, or ordinary cross-sheet references when it sits inside the client model
  prior   -> current client  those client rows -> the same line items in the current model (sheet + label,
                             n-th occurrence), with the values side by side
"""
import bisect
import re
import sqlite3
from collections import defaultdict

import extlinks
import rodb

LIVE = ("direct", "offset", "active")  # edge kinds the current scenario uses (edges.py)
STOP = {"the", "of", "and", "a", "at", "to", "in", "on", "value", "rate", "less", "total", "text"}


def _ro(path: str) -> sqlite3.Connection:
    return rodb.connect(path)


def _tokens(s: str) -> set[str]:
    return {w for w in re.findall(r"[a-z]+", (s or "").lower().replace("_", " ")) if w not in STOP and len(w) > 1}


def _labels(db) -> dict[tuple[str, int], str]:
    return {(s, r): lab or "" for s, r, lab in db.execute("SELECT sheet, row, label FROM rows")}


# ---- report -> overlay --------------------------------------------------------------------------------------

def _decimals(text: str) -> int:
    m = re.search(r"\d[\d,]*\.(\d+)", text or "")
    return len(m.group(1)) if m else 0


def _scales(unit: str) -> list[float]:
    u = (unit or "").lower()
    if "%" in u:
        return [0.01, 1.0]
    if re.search(r"bn|billion", u):
        return [1.0, 1e3, 1e6, 1e9]
    if re.search(r"(^|[^a-z])m$|million|\$m|mn", u):
        return [1.0, 1e3, 1e6]
    if re.search(r"'000|000s|\bk\b|thousand", u):
        return [1.0, 1e3]
    return [1.0]


class Numbers:
    """Every numeric cell of a model.db, sorted by value, for fast "which cells hold about X" lookups."""

    def __init__(self, db, sheets: set[str] | None = None):
        rows = db.execute("SELECT sheet, row, col, addr, value, formula IS NOT NULL FROM cells "
                          "WHERE typeof(value) IN ('real','integer') AND value <> 0")
        self.cells = sorted((float(v), s, r, c, a, bool(f)) for s, r, c, a, v, f in rows if not sheets or s in sheets)
        self.keys = [c[0] for c in self.cells]

    def between(self, lo: float, hi: float) -> list[tuple]:
        return self.cells[bisect.bisect_left(self.keys, lo):bisect.bisect_right(self.keys, hi)]


def match_fact(f: dict, nums: Numbers, labels: dict, anchors: set[str], limit: int = 5) -> dict:
    """Cells holding a report figure (preferred value, or low / high), best first."""
    out, n = [], 0
    want = _tokens(f.get("label")) | _tokens(f.get("key"))
    for part in ("value_text", "low_text", "high_text"):
        text = f.get(part) or ""
        m = re.search(r"\d[\d,]*(?:\.\d+)?", text)
        if not m:
            continue
        x = float(m.group(0).replace(",", ""))
        half = 0.5 * 10 ** -_decimals(text)
        for scale in _scales(f.get("unit") or text):
            for sign in (1, -1):
                lo, hi = sorted((sign * (x - half) * scale, sign * (x + half) * scale))
                for v, s, r, c, a, is_f in nums.between(lo, hi):
                    n += 1
                    lab = labels.get((s, r), "")
                    overlap = len(want & _tokens(lab)) / max(1, len(want))
                    score = 3 * overlap + (1.5 if f"{s}!{a}" in anchors else 0) \
                        + (0.5 if is_f == (f.get("category") in ("conclusion", "sensitivity")) else 0) \
                        + (0.5 if sign == 1 else 0) + (0.3 if part == "value_text" else 0)
                    out.append({"sheet": s, "addr": a, "row": r, "label": lab, "value": v, "formula": is_f,
                                "part": part.removesuffix("_text"), "scale": scale, "sign": sign,
                                "label_match": overlap > 0, "anchor": f"{s}!{a}" in anchors, "score": round(score, 2),
                                # where the figure is: its label says so, or it's a DCF. The value alone, in a
                                # sheet of thousands of numbers, is as likely a coincidence
                                "located": overlap > 0 or f"{s}!{a}" in anchors})
    out.sort(key=lambda m: -m["score"])
    return {"matches": out[:limit], "n": n}


def match_facts(db_path: str, facts: list[dict], sheets: set[str] | None = None) -> list[dict]:
    """[{fact_id, key, label, value_text, matches, n}] for the numeric report facts."""
    import valuation
    with _ro(db_path) as db:
        nums, labels = Numbers(db, sheets), _labels(db)
        dates = defaultdict(list)
        for s, r, a, v in db.execute("SELECT sheet, row, addr, value FROM cells WHERE typeof(value)='text' "
                                     "AND value GLOB '[12][0-9][0-9][0-9]-[01][0-9]-[0-3][0-9]*'"):
            if not sheets or s in sheets:
                dates[v[:10]].append((s, r, a))
    try:
        anchors = {a["cell"] for a in valuation.catalogue(db_path) if a.get("ok")}
    except Exception:
        anchors = set()
    out = []
    for f in facts:
        if f.get("unit") == "date" or f.get("key") == "valuation_date":
            v = f.get("value")
            iso = f"{int(v) // 10000:04d}-{int(v) // 100 % 100:02d}-{int(v) % 100:02d}" if v else None
            hits = dates.get(iso, []) if iso else []
            ms = [{"sheet": s, "addr": a, "row": r, "label": labels.get((s, r), ""), "value": iso, "formula": None,
                   "part": "value", "label_match": bool(re.search(r"valu", labels.get((s, r), ""), re.I)),
                   "anchor": False} for s, r, a in hits]
            for m in ms:
                m["score"], m["located"] = 3 * m["label_match"], m["label_match"]
            ms.sort(key=lambda m: -m["score"])
            out.append({"fact_id": f["id"], "key": f["key"], "label": f.get("label"), "value_text": f.get("value_text"),
                        "matches": ms[:5], "n": len(hits)})
        elif f.get("category") in ("conclusion", "assumption", "sensitivity") and any(
                re.search(r"\d", f.get(k) or "") for k in ("value_text", "low_text", "high_text")):  # a range too
            out.append({"fact_id": f["id"], "key": f["key"], "label": f.get("label"), "value_text": f.get("value_text"),
                        **match_fact(f, nums, labels, anchors)})
    return out


# ---- overlay -> prior client --------------------------------------------------------------------------------

def _norm_file(name: str | None) -> str:
    return re.sub(r"[^a-z0-9]+", "", re.sub(r"\.(xlsx|xlsm|xlsb|xls)$", "", (name or "").lower()))


def ext_cached_match(overlay_db: str, idx: int, client_db: str) -> dict:
    """Do the values the overlay last saw through link [idx] equal the client workbook's cells?"""
    with _ro(overlay_db) as o, _ro(client_db) as c:
        cached = o.execute("SELECT sheet, row, col, value FROM extcells WHERE idx=?", (idx,)).fetchall()
        n = ok = 0
        diffs = []
        for s, r, col, v in cached:
            got = c.execute("SELECT value FROM cells WHERE sheet=? AND row=? AND col=?", (s, r, col)).fetchone()
            n += 1
            same = got is not None and (got[0] == v or _close(got[0], v))
            ok += same
            if not same and len(diffs) < 10:
                diffs.append({"cell": f"{s}!r{r}c{col}", "cached": v, "client": got[0] if got else None})
    return {"cells": n, "matched": ok, "differences": diffs}


def _close(a, b) -> bool:
    try:
        a, b = float(a), float(b)
    except (TypeError, ValueError):
        return False
    return abs(a - b) <= 1e-6 * max(1.0, abs(a), abs(b))


def overlay_to_client(overlay: dict, client: dict | None) -> dict:
    """overlay / client: {"db_path", "source_path", "filename", "sheets" (set or None = all)}.
    Returns {"mode", "links": [{overlay row -> client row}], "books": [...]} ."""
    links, books = [], []
    same_file = client is not None and overlay["db_path"] == client["db_path"]
    with _ro(overlay["db_path"]) as o:
        o_labels = _labels(o)
        ov_sheets = overlay.get("sheets")
        if same_file:
            cl_sheets = client.get("sheets")
            live = ",".join(f"'{k}'" for k in LIVE)
            for s, r, ds, dr, kind in o.execute(f"""SELECT DISTINCT src_sheet, src_row, dst_sheet, dst_row, kind FROM edges
                                                  WHERE src_sheet <> dst_sheet AND kind IN ({live})"""):
                if (not ov_sheets or s in ov_sheets) and (cl_sheets and ds in cl_sheets):
                    links.append({"overlay": f"{s}!r{r}", "overlay_label": o_labels.get((s, r), ""),
                                  "client": f"{ds}!r{dr}", "client_sheet": ds, "client_row": dr,
                                  "client_label": o_labels.get((ds, dr), ""), "via": f"same workbook ({kind})"})
            return {"mode": "inside the client model", "links": links, "books": []}
    extlinks.ensure(overlay["source_path"], overlay["db_path"])
    books = extlinks.summary(overlay["db_path"])
    target = None
    if client:
        by_name = [b for b in books if _norm_file(b["filename"]) == _norm_file(client["filename"])]
        with _ro(client["db_path"]) as c:
            cl_sheets = {r[0] for r in c.execute("SELECT sheet FROM sheets")}
        for b in books:
            b["cached_check"] = ext_cached_match(overlay["db_path"], b["idx"], client["db_path"]) if b["cached_cells"] else None
            b["is_client"] = b in by_name
            b["sheet_share"] = len(set(b["sheets"]) & cl_sheets) / len(b["sheets"]) if b["sheets"] else 0.0
        # The link whose cached values best match the client file is the client model, whatever its name now; failing
        # both name and values (a later version of the model, renamed), the one link whose sheets are the client
        # model's sheets. Without that last step nothing would read the client model, silently: every feed would
        # give the overlay's saved values.
        scored = sorted(books, key=lambda b: (b.get("is_client", False), (b["cached_check"] or {}).get("matched", 0),
                                             b["sheet_share"]), reverse=True)
        by_sheets = [b for b in books if b["sheet_share"] >= 0.8]
        if scored and (scored[0].get("is_client") or (scored[0]["cached_check"] or {}).get("matched")):
            target = scored[0]["idx"]
            scored[0]["matched_by"] = "name" if scored[0].get("is_client") else "the values it last read"
            scored[0]["is_client"] = True
        elif len(by_sheets) == 1:
            target = by_sheets[0]["idx"]
            by_sheets[0].update(is_client=True, matched_by="its sheet names (neither its name nor the values it last "
                                                          "read match the file: probably another version)")
    c_labels = {}
    if client and target:
        with _ro(client["db_path"]) as c:
            c_labels = _labels(c)
    with _ro(overlay["db_path"]) as o:
        for s, r, idx, es, er, n in o.execute("SELECT * FROM extrefs ORDER BY sheet, row"):
            if ov_sheets and s not in ov_sheets:
                continue
            links.append({"overlay": f"{s}!r{r}", "overlay_label": o_labels.get((s, r), ""), "link": idx,
                          "client": f"{es}!r{er}" if er else es, "client_sheet": es, "client_row": er,
                          "client_label": c_labels.get((es, er), "") if idx == target else "",
                          "via": f"external link [{idx}]" + (" (the prior client model)" if idx == target else "")})
    return {"mode": "separate workbook", "links": links, "books": books, "client_link": target}


# ---- prior client -> current client -------------------------------------------------------------------------

def align_rows(prior_db: str, current_db: str, refs: list[tuple[str, int]]) -> dict[tuple[str, int], dict]:
    """Prior (sheet, row) -> the current model's row with the same sheet and label (same occurrence of the label),
    with values for the first periods both models have (matched by timeline date, as the timeline rolls forward)."""
    import diff as diffmod

    def occurrences(db):
        seen, out = defaultdict(int), {}
        for s, r, lab in db.execute("SELECT sheet, row, label FROM rows ORDER BY sheet, row"):
            k = (s, re.sub(r"\s+", " ", (lab or "").strip().lower()))
            seen[k] += 1
            out[(s, r)] = (*k, seen[k])
        return out

    import overlay as ovmod
    db = rodb.connect(current_db)
    db.execute("ATTACH DATABASE ? AS o", (rodb.uri(prior_db),))
    with _ro(prior_db) as p:
        po = occurrences(p)
    import rowfind
    pw, cw = ovmod.Workbook(prior_db), ovmod.Workbook(current_db)
    rm = rowfind.RowFinder(ovmod.RowMap(pw, cw), pw, cw)  # the same finding as the rolled-forward feed
    out, tl = {}, {}
    for ref in refs:
        key = po.get(ref) or (ref[0], "", 0)
        ex = rm.explain(*ref)
        if ex["found"] is None:
            out[ref] = {"current": None, "why": rm.why(ref[0], ref[1], pw.labels())}
            continue
        hit = ex["found"]
        sheet_o, sheet_n = ref[0], hit[0]
        if (sheet_o, sheet_n) not in tl:
            ot, nt = diffmod.timeline(db, "o", sheet_o), diffmod.timeline(db, "main", sheet_n)
            by_date = {d: col for col, d in nt.items()}
            tl[(sheet_o, sheet_n)] = sorted((d, oc, by_date[d]) for oc, d in ot.items() if d in by_date)[:3]
        periods = tl[(sheet_o, sheet_n)]
        val = lambda schema, s, r, col: (db.execute(f"SELECT value FROM {schema}.cells WHERE sheet=? AND row=? AND col=?",
                                                    (s, r, col)).fetchone() or [None])[0]
        if periods:
            pv = [val("o", sheet_o, ref[1], oc) for _, oc, _ in periods]
            cv = [val("main", sheet_n, hit[1], nc) for _, _, nc in periods]
        else:
            first = lambda schema, s, r: [v for (v,) in db.execute(
                f"SELECT value FROM {schema}.cells WHERE sheet=? AND row=? AND typeof(value) IN ('real','integer') "
                "ORDER BY col LIMIT 3", (s, r))]
            pv, cv = first("o", *ref), first("main", sheet_n, hit[1])
        out[ref] = {"current": f"{sheet_n}!r{hit[1]}", "label": key[1], "periods": [d for d, _, _ in periods],
                    "prior_values": pv, "current_values": cv, "how": ex["how"],
                    "evidence": [f"{n}: {t}" for n, t in ex["evidence"][:3]]}
    db.close()
    return out


def chain(ov_links: dict, alignment: dict) -> list[dict]:
    """Each overlay -> prior client link, extended to the current model's matching row."""
    rows = []
    for ln in ov_links["links"]:
        cur = alignment.get((ln["client_sheet"], ln["client_row"])) if ln.get("client_row") else None
        cur = cur or {}
        rows.append({**ln, "current": cur.get("current"), "periods": cur.get("periods"),
                     "prior_values": cur.get("prior_values"), "current_values": cur.get("current_values"),
                     "why": cur.get("why")})
    return rows
