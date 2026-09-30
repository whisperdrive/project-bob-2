"""The Valuation tab: every DCF in a workbook, found and recomputed from its own formulas, with no chat model.

A DCF here is a cell whose formula (directly or through cells it adds up) contains
SUMPRODUCT(cash-flow row, discount-factor row). For each one:
  - the discount factors are read back into a rate, valuation date, convention and cut-off date: the rate is
    back-solved from the factors for each candidate valuation date and convention, and kept only if it
    reproduces every factor
  - the cell's formula is split into additive terms: the SUMPRODUCT is the PV; other cells, and SUMIFS picking
    one period's amount, are the bridge (debt, cash, a valuation-date distribution, ...)
  - dcf.compute() redoes it in Python and checks it against the workbook's value
Low / high values change the discount rate only; the saved cash flows and bridge stay as they are.
"""
import hashlib
import json
import math
import os
import re
import sqlite3
import threading
from datetime import date
from pathlib import Path

import dcf
import rodb

_CACHE: dict = {}
_LOCK = threading.Lock()


# ---- splitting a formula into additive terms ----------------------------------------------------------------

def _split(expr: str, seps: str) -> list[tuple[str, str]]:
    """Top-level split of expr on the characters in seps (outside brackets and quotes) -> [(sep, part)]."""
    out, depth, quote, cur, sign = [], 0, False, "", "+"
    for ch in expr:
        if ch == '"':
            quote = not quote
        if not quote:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
            elif depth == 0 and ch in seps:
                if cur.strip():
                    out.append((sign, cur.strip()))
                elif ch == "-":  # a leading minus
                    sign = "-" if sign == "+" else "+"
                    continue
                cur, sign = "", ch
                continue
        cur += ch
    if cur.strip():
        out.append((sign, cur.strip()))
    return out


_FUNC = re.compile(r"^(?P<fn>[A-Z][A-Z0-9_.]*)\((?P<args>.*)\)$", re.I | re.S)
_WHOLE_ROW = re.compile(r"^(?:'?(?P<sheet>[^!']+)'?!)?\$?(?P<r1>\d+):\$?(?P<r2>\d+)$")


def _strip(t: str) -> str:
    t = t.strip()
    while t.startswith("(") and t.endswith(")") and _split(t[1:-1], "\0") == [("+", t[1:-1].strip())]:
        t = t[1:-1].strip()
    return t


def _row_values(db, ref_text: str, here: str) -> tuple[str, int, dict[int, object]] | None:
    """A single-row range ("$L172:$HO172", "Debt!$61:$61") -> (sheet, row, {col: value})."""
    m = _WHOLE_ROW.match(ref_text.strip())
    if m and m["r1"] == m["r2"]:
        sheet, row = (m["sheet"] or here), int(m["r1"])
        return sheet, row, dict(db.execute("SELECT col, value FROM cells WHERE sheet=? AND row=?", (sheet, row)))
    r = dcf._ref(ref_text, here)
    if r and r[1] == r[3]:
        sheet, row = r[0], r[1]
        return sheet, row, dict(db.execute("SELECT col, value FROM cells WHERE sheet=? AND row=? AND col BETWEEN ? AND ?",
                                           (sheet, row, r[2], r[4])))
    return None


def _criterion(db, text: str, here: str):
    t = text.strip()
    if t.startswith('"'):
        return t.strip('"')
    try:
        return dcf.resolve(db, t if "!" in t or db.execute("SELECT 1 FROM names WHERE lower(name)=lower(?)", (t,)).fetchone()
                           else f"{here}!{t}")[0]
    except ValueError:
        return None


class _Undecomposable(Exception):
    pass


def _reaches(db, sheet: str, row: int, col: int, pv_key, seen=None, depth: int = 5) -> bool:
    """Does this cell's formula lead (through single cells / short ranges) to the PV's SUMPRODUCT?"""
    seen = seen if seen is not None else set()
    if (sheet, row, col) in seen or depth < 0:
        return False
    seen.add((sheet, row, col))
    f = db.execute("SELECT formula FROM cells WHERE sheet=? AND row=? AND col=?", (sheet, row, col)).fetchone()
    if not f or not f[0]:
        return False
    if any(_sp_key(m, sheet) == pv_key for m in dcf._SUMPRODUCT.finditer(f[0])):
        return True
    for m in dcf._FREF.finditer(re.sub(r'"[^"]*"', "", f[0])):
        r = dcf._ref(m[0], sheet)
        if r and (r[3] - r[1] + 1) * (r[4] - r[2] + 1) <= 12:
            if any(_reaches(db, r[0], rr, cc, pv_key, seen, depth - 1)
                   for rr in range(r[1], r[3] + 1) for cc in range(r[2], r[4] + 1)):
                return True
    return False


def _sp_key(m, here):
    a, b = dcf._ref(m[1], here), dcf._ref(m[2], here)
    return (a, b) if a and b else None


def decompose(db, sheet: str, row: int, col: int, pv_key, vd: date, sign: int = 1, depth: int = 6) -> dict:
    """Split a value cell into sign * (PV + bridge items). Returns {"pv_sign": int, "items": [...]};
    each item is an adjustment for dcf.compute(). Raises _Undecomposable if a term mixes the PV with
    anything but addition or subtraction."""
    f = db.execute("SELECT formula, value FROM cells WHERE sheet=? AND row=? AND col=?", (sheet, row, col)).fetchone()
    here = f"{sheet}!{dcf._addr(col, row)}"
    if not f or not f[0]:
        v = dcf._num(f[1]) if f else None
        return {"pv_sign": 0, "items": [{"label": dcf._row_label(db, sheet, row) or here,
                                         "value": sign * (v or 0.0), "source": here}]}
    if depth < 0:
        raise _Undecomposable(here)
    pv_sign, items = 0, []
    for op, term in _split(f[0].lstrip("=").lstrip("+"), "+-"):
        s = sign * (-1 if op == "-" else 1)
        t = _strip(term.lstrip("+"))
        fm = _FUNC.match(t)
        fn = fm["fn"].upper() if fm else None
        if fn == "SUMPRODUCT":
            sp = dcf._SUMPRODUCT.fullmatch(t)
            if sp and _sp_key(sp, sheet) == pv_key:
                pv_sign += s
                continue
            raise _Undecomposable(t)
        if fn == "SUM":
            refs = []
            for _, arg in _split(fm["args"], ","):
                r = dcf._ref(arg, sheet)
                if not r:
                    raise _Undecomposable(t)
                refs += [(r[0], rr, cc) for rr in range(r[1], r[3] + 1) for cc in range(r[2], r[4] + 1)]
            if len(refs) > 50:
                raise _Undecomposable(t)
            for r in refs:
                sub = _cell_term(db, r, pv_key, vd, s, depth)
                pv_sign += sub["pv_sign"]
                items += sub["items"]
            continue
        if fn == "SUMIFS":
            args = [a for _, a in _split(fm["args"], ",")]
            if len(args) != 3:
                raise _Undecomposable(t)
            rows = _row_values(db, args[0], sheet)
            crit_rows = _row_values(db, args[1], sheet)
            crit = _criterion(db, args[2], sheet)
            if not rows or not crit_rows or crit is None:
                raise _Undecomposable(t)
            key = dcf._as_date(crit) or crit
            hits = [c for c, v in crit_rows[2].items() if (dcf._as_date(v) or v) == key]
            total = sum(dcf._num(rows[2].get(c)) or 0.0 for c in hits)
            label = dcf._row_label(db, rows[0], rows[1]) or f"{rows[0]}!r{rows[1]}"
            when = f" at {key.isoformat()}" if isinstance(key, date) else ""
            where = f"{rows[0]}!{dcf._addr(hits[0], rows[1])}" if len(hits) == 1 else f"SUMIFS of {rows[0]}!r{rows[1]}"
            items.append({"label": label + (" (valuation date)" if key == vd else when), "value": s * total,
                          "source": where})
            continue
        r = dcf._ref(t, sheet) if not fm else None
        if r and r[1] == r[3] and r[2] == r[4]:
            sub = _cell_term(db, (r[0], r[1], r[2]), pv_key, vd, s, depth)
            pv_sign += sub["pv_sign"]
            items += sub["items"]
            continue
        try:
            items.append({"label": t, "value": s * float(t), "source": "constant in formula"})
            continue
        except ValueError:
            pass
        raise _Undecomposable(t)
    return {"pv_sign": pv_sign, "items": items}


def _cell_term(db, ref, pv_key, vd, sign, depth):
    sh, r, c = ref
    if _reaches(db, sh, r, c, pv_key):
        return decompose(db, sh, r, c, pv_key, vd, sign, depth - 1)
    where = f"{sh}!{dcf._addr(c, r)}"
    lab = dcf._row_label(db, sh, r)
    v = dcf._num(dcf._cell(db, sh, r, c)) or 0.0
    return {"pv_sign": 0, "items": [{"label": re.sub(r"^(add|less)\s*:\s*", "", lab, flags=re.I) or where,
                                     "value": sign * v, "source": where}]}


# ---- reading the discount factors back into assumptions -----------------------------------------------------

def read_factors(db, df_ref, cols: list[int], theirs: dict | None = None, sheet: str | None = None,
                 dates: str | None = None) -> dict | None:
    """Rate, valuation date, convention and cut-off that reproduce the workbook's factor row exactly; or, with
    theirs, factors given by column (e.g. computed inside a formula), timed by sheet's period dates. Only the
    factors there are count: a period with no cash flow shows no factor (pv / cash flow can't be taken), and isn't
    a factor of 0."""
    sheet, row = (df_ref[0], df_ref[1]) if df_ref else (sheet, None)
    if theirs is None:
        theirs = {c: dcf._num(v) or 0.0 for c, v in db.execute(
            "SELECT col, value FROM cells WHERE sheet=? AND row=?", (sheet, row)) if c in cols}
        theirs.update({c: 0.0 for c in cols if c not in theirs})  # a factor row's blank cell is a factor of 0
    ends, ends_src = dcf.period_ends(db, sheet, cols, dates)
    live = [c for c in sorted(ends) if 0 < theirs.get(c, 0.0) < 1]
    if not live:
        return None
    last = max(live, key=lambda c: ends[c])
    after = [c for c in ends if ends[c] > ends[last]]
    td = ends[last] if after and all(c in theirs and not theirs[c] for c in after) else None  # zeros seen, not unseen
    rates = dcf.candidate_cells(db, "rate", 40)
    for vd_c in dcf.candidate_cells(db, "date", 20):
        vd = dcf._as_date(vd_c["value"])
        first = min(live, key=lambda c: ends[c])
        if not vd or ends[first] <= vd:
            continue
        for timing in dcf.TIMINGS:
            for dc in dcf.DAY_COUNTS:
                # back-solve the rate from the first live factor, then check every factor
                probe = dcf.factors(ends, vd, 0.1, timing, dc, td)
                t = math.log(probe[first]) / math.log(1 / 1.1)
                rate = theirs[first] ** (-1 / t) - 1
                ours = dcf.factors(ends, vd, rate, timing, dc, td)
                if all(abs(ours[c] - theirs[c]) < 1e-9 for c in ends if c in theirs):
                    rate_cell = next((r for r in rates if abs(dcf._num(r["value"]) - rate) < 1e-9), None)
                    return {"rate": rate_cell["ref"] if rate_cell else round(rate, 12),
                            "rate_note": None if rate_cell else f"back-solved from the factors"
                                                                + (f" in {sheet}!r{row}" if row else ""),
                            "valuation_date": vd_c["ref"], "timing": timing, "day_count": dc,
                            "terminal_date": td.isoformat() if td else None, "ends_source": ends_src,
                            "factor_row": f"{sheet}!r{row} {dcf._row_label(db, sheet, row)}".strip() if row
                            else "factors computed in the formula"}
    return None


# ---- the tab ------------------------------------------------------------------------------------------------

def _pv_cells(db) -> list[tuple[str, int, int, str, tuple]]:
    """Cells whose formula is a SUMPRODUCT of a cash-flow row and a discount-factor row."""
    out = []
    for sheet, row, col, f in db.execute("SELECT sheet, row, col, formula FROM cells WHERE formula LIKE '%SUMPRODUCT(%'"):
        for m in dcf._SUMPRODUCT.finditer(f):
            key = _sp_key(m, sheet)
            if not key or key[0][1] != key[0][3] or key[1][1] != key[1][3]:
                continue
            a, b = key
            if _is_df(db, a) and not _is_df(db, b):
                a, b = b, a
            if _is_df(db, b) and not _is_df(db, a):
                out.append((sheet, row, col, f, (a, b), key))
    return out


def _is_df(db, ref) -> bool:
    vals = [dcf._num(v) for (v,) in db.execute("SELECT value FROM cells WHERE sheet=? AND row=? AND col BETWEEN ? AND ?",
                                               (ref[0], ref[1], ref[2], ref[4]))]
    vals = [v for v in vals if v is not None]
    return bool(vals) and all(0 <= v <= 1 for v in vals) and sum(0 < v < 1 for v in vals) >= 3


def find(db) -> list[dict]:
    """Every DCF result in the workbook: cells labelled like a valuation (enterprise value, equity PV, NPV,
    total valuation, ...) that are, or add a bridge to, a SUMPRODUCT of cash flows and discount factors."""
    pvs = _pv_cells(db)
    if not pvs:
        return []
    found, seen = [], set()
    # result cells: labelled like a valuation (not every line item's PV), with a formula leading to a PV cell
    cands = []
    for line in dcf.outputs(db, limit=200):
        ref = line.split(" = ")[0]
        r = dcf._ref(ref, "")
        if r:
            cands.append((r[0], r[1], r[2]))
    for sheet, row, col in cands:
        if (sheet, row, col) in seen:
            continue
        seen.add((sheet, row, col))
        for ps, pr, pc, pf, (cf, dfr), key in pvs:
            if (sheet, row, col) != (ps, pr, pc) and not _reaches(db, sheet, row, col, key):
                continue
            v = dcf._num(dcf._cell(db, sheet, row, col))
            if v is None or v == 0:
                break
            found.append({"cell": f"{sheet}!{dcf._addr(col, row)}", "label": dcf._row_label(db, sheet, row) or "",
                          "value": v, "pv_cell": f"{ps}!{dcf._addr(pc, pr)}", "cf": cf, "df": dfr, "key": key})
            break
    return found


def build(db, v: dict) -> dict:
    """Derive the inputs for one found DCF and recompute it (no low/high yet)."""
    sheet, row, col = dcf._ref(v["cell"], "")[:3]
    cf, dfr = v["cf"], v["df"]
    cols = list(range(cf[2], cf[4] + 1))
    fx = read_factors(db, dfr, cols)
    if not fx:
        return {**_public(v), "ok": False, "reason": "couldn't read a rate and valuation date back from the discount factors"}
    vd = dcf._as_date(dcf.resolve(db, fx["valuation_date"])[0])
    try:
        parts = decompose(db, sheet, row, col, v["key"], vd)
    except _Undecomposable as e:
        return {**_public(v), "ok": False, "reason": f"its formula isn't a plain sum of the PV and other amounts ({e})"}
    if parts["pv_sign"] != 1:
        return {**_public(v), "ok": False, "reason": "the PV doesn't enter its formula exactly once with a + sign"}
    cf_range = f"{cf[0]}!{dcf._addr(cf[2], cf[1])}:{dcf._addr(cf[4], cf[1])}"
    return {**_public(v), "ok": True, "df": list(dfr), "inputs": {
        "cashflow": [cf_range], "rate": fx["rate"], "valuation_date": fx["valuation_date"],
        "timing": fx["timing"], "day_count": fx["day_count"], "terminal_date": fx["terminal_date"],
        "adjustments": parts["items"], "compare_to": v["cell"]},
        "factor_row": fx["factor_row"], "rate_note": fx["rate_note"], "ends_source": fx["ends_source"]}


def _public(v):
    return {k: v[k] for k in ("cell", "label", "value", "pv_cell")}


_SOURCE: list = []


def _source() -> str:
    """The code a catalogue comes from: a change to it makes saved catalogues stale."""
    if not _SOURCE:
        h = hashlib.sha256()
        for f in ("valuation.py", "dcf.py"):
            h.update((Path(__file__).resolve().parent / f).read_bytes())
        _SOURCE.append(h.hexdigest()[:16])
    return _SOURCE[0]


def catalogue(db_path: str) -> list[dict]:
    """find() + build() for every DCF in a workbook, cached per model.db: in memory, and saved beside it
    (catalogue.json, for that model.db's modification time and this code), so a restarted server doesn't redo it
    (on a large workbook it takes minutes)."""
    key = (db_path, os.path.getmtime(db_path))
    with _LOCK:
        if key in _CACHE:
            return _CACHE[key]
    saved = Path(db_path).parent / "catalogue.json"
    try:
        d = json.loads(saved.read_text(encoding="utf-8"))
        if d.get("mtime") == key[1] and d.get("source") == _source():
            with _LOCK:
                _CACHE[key] = d["items"]
            return d["items"]
    except (OSError, ValueError, KeyError, AttributeError):
        pass
    db = rodb.connect(db_path)
    out = []
    for v in find(db):
        b = build(db, v)
        if b["ok"]:
            r = dcf.compute(db, **b["inputs"], fix=False)
            b["matches"] = r["compare_to"] is not None and dcf._close(r["total"], r["compare_to"])
            b["rate_value"] = r["rate"]
            b["total"] = r["total"]
        out.append(b)
    out.sort(key=lambda b: (not b.get("ok"), not b.get("matches"), _rank(b["label"]), b["cell"]))
    with _LOCK:
        _CACHE[key] = out
    try:
        tmp = saved.with_name(f".catalogue.{os.getpid()}.{threading.get_ident()}.tmp")
        tmp.write_text(json.dumps({"mtime": key[1], "source": _source(), "items": out}), encoding="utf-8")
        os.replace(tmp, saved)
    except (OSError, TypeError, ValueError):
        pass
    return out


def _rank(label: str) -> int:
    for i, pat in enumerate((r"enterprise value", r"equity (value|pv)", r"\bnpv\b|present value|total valuation")):
        if re.search(pat, label or "", re.I):
            return i
    return 9


def _listing(cat):
    return [{"cell": b["cell"], "label": b["label"], "value": b["value"], "total": b.get("total"),
             "ok": b.get("ok", False), "matches": b.get("matches", False), "reason": b.get("reason"),
             "traced_from": b.get("traced_from"), "what": b.get("what")} for b in cat]


def _pick(cat, cell):
    usable = [b for b in cat if b.get("ok")]
    return next((b for b in usable if b["cell"] == cell), usable[0] if usable else None)


def _chart(db, db_path, cf_range, runs: list[tuple], cumulative: bool = False, dates: str | None = None) -> dict:
    """The workbook's cash flows and each run's present value per period; or, with cumulative, only each run's
    running total of present value (how the value builds up; the last point is the PV of cash flows), which
    shows where runs diverge. runs: [(series name, compute() result)], or (name, result, flows, ends) for a run
    with its own cash flows and period ends by column (the Python overlay's, on another feed)."""
    import chartdata
    import tools
    with tools.using(db_path):
        spec = tools.chart("Cash flows and present values", [{"range": cf_range, "name": "Cash flow (workbook)"}],
                           kind="bar")
    sheet, _, cf_cols = dcf._row_range(db, cf_range)
    ends, _ = dcf.period_ends(db, sheet, cf_cols, dates)
    cols = spec.get("columns") or []
    cfd = spec["series"][0]["data"]
    live = set()
    for name, r, *own in runs:
        flows, run_ends = (own + [None, None])[:2]
        f = dcf.factors(run_ends or ends, r["valuation_date"], r["rate"], r["timing"], r["day_count"], r["terminal_date"])
        live |= {i for i, c in enumerate(cols) if f.get(c)}
        data = [flows.get(c) for c in cols] if flows is not None else cfd
        pv = [(v * f[c]) if isinstance(v, (int, float)) and f.get(c) else None for v, c in zip(data, cols)]
        if cumulative:
            run, acc, started = [], 0.0, False
            for x in pv:
                started = started or x is not None
                acc += x or 0.0
                run.append(acc if started else None)
            pv = run
        # the source row's range keeps the chart's phase shading on that sheet
        spec["series"].append({"name": name, "range": f"{cf_range} (discounted in Python)",
                               "label": "Cumulative present value" if cumulative else "Present value",
                               "units": spec["series"][0].get("units"), "data": pv})
    if cumulative:
        spec["series"] = spec["series"][1:]
        spec.update(title="Cumulative present value", kind="line")
    spec = chartdata.enrich(spec, db)
    live = sorted(live)
    if live:
        x_end, extra = live[-1], ""
        # chart rule F1: a period more than 5x the next largest (typically a terminal value at the end) flattens
        # everything else, so leave it out of the default view and say so
        vals = sorted(((abs(cfd[i]), i) for i in live if isinstance(cfd[i], (int, float))), reverse=True)
        if not cumulative and len(vals) > 2 and vals[1][0] and vals[0][0] > 5 * vals[1][0] and vals[0][1] >= live[-1] - 1:
            i = vals[0][1]
            x_end = i - 1
            extra = (f" The {spec['period_labels'][i]} cash flow ({cfd[i]:,.0f}, {vals[0][0] / vals[1][0]:.0f}x the next "
                     f"largest) is left out of this view so the rest is readable.")
        spec["view"] = {"x_start": live[0], "x_end": x_end}
        if spec.get("annual") and len(live) > 40:
            spec["mode"] = "annual"  # a few dozen annual bars read better than 100+ quarterly ones
        spec["note"] = (f"Showing the discounted periods ({spec['period_labels'][live[0]]} to "
                        f"{spec['period_labels'][live[-1]]}).{extra} \u201cFull range\u201d shows the whole timeline."
                        + (" Each line is the running total of present value; its last point is the PV of cash flows."
                           if cumulative else ""))
    return spec


def _assumptions(pick, r) -> dict:
    iso = lambda d: d.isoformat() if d else None
    return {"rate": r["rate"], "rate_source": pick["rate_note"] or r["rate_source"],
            "valuation_date": iso(r["valuation_date"]), "valuation_date_source": r["valuation_date_source"],
            "terminal_date": iso(r["terminal_date"]), "timing": r["timing"], "day_count": r["day_count"],
            "cashflow": r["rows"], "factor_row": pick["factor_row"], "period_ends": pick["ends_source"],
            "periods": r["periods"], "first_period": iso(r["first_period"]), "last_period": iso(r["last_period"]),
            "undiscounted": r["undiscounted"], "units": r.get("units") or []}


def validation(db_path: str, cell: str | None = None, anchors: list[dict] | None = None) -> dict:
    """Step 1: can the workbook's own numbers be reproduced? Every anchor value, and for the selected one the
    approach found and a check of each step (cash flows, discount factors, PV, bridge, the anchor itself).
    anchors: the anchors to choose from (default: this workbook's catalogue)."""
    cat = catalogue(db_path) if anchors is None else anchors
    pick = _pick(cat, cell)
    if not pick:
        return {"anchors": _listing(cat), "selected": None}
    db = rodb.connect(db_path)
    r = dcf.compute(db, **pick["inputs"], fix=False)
    sheet, row, cols = dcf._row_range(db, pick["inputs"]["cashflow"][0])
    dates = pick["inputs"].get("dates")
    ends, _ = dcf.period_ends(db, sheet, cols, dates)
    checks = []

    # cash flows: the row's own total column, if it has one (e.g. =SUM(L173:HO173))
    flows = {c: dcf._num(v) or 0.0 for c, v in db.execute(
        "SELECT col, value FROM cells WHERE sheet=? AND row=? AND col BETWEEN ? AND ?", (sheet, row, cols[0], cols[-1]))}
    total_cell = None
    for c, f, v in db.execute("SELECT col, formula, value FROM cells WHERE sheet=? AND row=? AND formula LIKE '%SUM(%' "
                              "AND (col < ? OR col > ?)", (sheet, row, cols[0], cols[-1])):
        m = re.fullmatch(r"=\+?SUM\(\$?([A-Z]+)\$?\d+:\$?([A-Z]+)\$?\d+\)", (f or "").replace(" ", ""), re.I)
        if m and dcf._col(m[1]) == cols[0] and dcf._col(m[2]) == cols[-1] and dcf._num(v) is not None:
            total_cell = (f"{sheet}!{dcf._addr(c, row)}", dcf._num(v))
            break
    ours_sum = sum(flows.values())
    checks.append({"step": "Cash flows", "what": f"{len(cols)} periods read from {r['rows'][0]}",
                   "ours": ours_sum, "theirs": total_cell[1] if total_cell else None,
                   "where": total_cell[0] if total_cell else None,
                   "ok": dcf._close(ours_sum, total_cell[1]) if total_cell else None,
                   "note": None if total_cell else "the row has no total column to compare with"})

    # discount factors, period by period
    dfr = pick.get("df")
    ours = dcf.factors(ends, r["valuation_date"], r["rate"], r["timing"], r["day_count"], r["terminal_date"])
    if dfr:
        theirs = {c: dcf._num(v) or 0.0 for c, v in db.execute(
            "SELECT col, value FROM cells WHERE sheet=? AND row=?", (dfr[0], dfr[1])) if c in ends}
        worst = max((abs(ours[c] - theirs.get(c, 0.0)), c) for c in ours)
        checks.append({"step": "Discount factors", "what": f"{sum(1 for c in ours if ours[c])} non-zero factors vs "
                       f"{dfr[0]}!r{dfr[1]}, {r['timing']}-of-period, {r['day_count']}",
                       "ours": None, "theirs": None, "where": f"{dfr[0]}!r{dfr[1]}", "ok": worst[0] < 1e-9,
                       "note": f"largest difference {worst[0]:.1e} ({ends[worst[1]]})"})
    else:  # computed inside the formula (XNPV, or an expression): the present value below checks them
        checks.append({"step": "Discount factors", "what": f"{sum(1 for c in ours if ours[c])} factors, "
                       f"{r['timing']}-of-period, {r['day_count']}: {pick.get('what') or 'computed in the formula'}",
                       "ours": None, "theirs": None, "where": pick["pv_cell"], "ok": None,
                       "note": "the workbook computes them inside the formula, so the present value checks them"})

    # PV: the workbook's SUMPRODUCT cell
    pv_ref = dcf._ref(pick["pv_cell"], "")
    pv_theirs = dcf._num(dcf._cell(db, *pv_ref[:3]))
    checks.append({"step": "Present value", "what": "sum of cash flow x discount factor", "ours": r["pv"],
                   "theirs": pv_theirs, "where": pick["pv_cell"],
                   "ok": pv_theirs is not None and dcf._close(r["pv"], pv_theirs), "note": None})

    # bridge
    for b in r["bridge"]:
        checks.append({"step": "Bridge", "what": b["label"], "ours": b["value"], "theirs": None, "where": b["source"],
                       "ok": None, "note": "read from the workbook"})

    checks.append({"step": "Anchor value", "what": pick["label"] or pick["cell"], "ours": r["total"],
                   "theirs": r["compare_to"], "where": pick["cell"],
                   "ok": r["compare_to"] is not None and dcf._close(r["total"], r["compare_to"]), "note": None})
    A = _assumptions(pick, r)
    return {"anchors": _listing(cat), "selected": pick["cell"], "validated": checks[-1]["ok"],
            "assumptions": A, "checks": checks, "bridge": r["bridge"], "pv": r["pv"], "total": r["total"],
            "chart": _chart(db, db_path, pick["inputs"]["cashflow"][0], [("Present value (model)", r)], dates=dates)}


def scenario(db_path: str, cell: str | None = None, rate: float | None = None, valuation_date: str | None = None,
             timing: str | None = None, day_count: str | None = None, cutoff: str | None = "model",
             include: list[bool] | None = None, low: float | None = None, high: float | None = None) -> dict:
    """Step 2: the validated anchor recomputed under chosen assumptions and methodology. Anything not given
    stays as the model has it. cutoff: "model" (the model's cut-off), "" (none) or YYYY-MM-DD.
    include: which bridge items to keep (all by default). low / high: sensitivity rates around the scenario rate
    (default: scenario rate +/- 1pp). Cash flows and bridge amounts are the workbook's saved values."""
    cat = catalogue(db_path)
    pick = _pick(cat, cell)
    if not pick:
        return {"selected": None}
    db = rodb.connect(db_path)
    base_in = pick["inputs"]
    model = dcf.compute(db, **base_in, fix=False)
    s_in = dict(base_in)
    if rate is not None and rate == rate:
        s_in["rate"] = rate
    if valuation_date:
        s_in["valuation_date"] = valuation_date
    if timing in dcf.TIMINGS:
        s_in["timing"] = timing
    if day_count in dcf.DAY_COUNTS:
        s_in["day_count"] = day_count
    if cutoff != "model":
        s_in["terminal_date"] = cutoff or None
    items = base_in["adjustments"]
    if include is not None:
        items = [a for a, keep in zip(items, list(include) + [True] * len(items)) if keep]
    s_in["adjustments"] = items
    s_rate = dcf._num(dcf.resolve(db, s_in["rate"])[0])
    lo = s_rate + 0.01 if low is None or low != low else low
    hi = max(s_rate - 0.01, 0.0) if high is None or high != high else high
    sc = dcf.compute(db, **{**s_in, "compare_to": None}, rates=[lo, hi], fix=False)
    adj = sum(b["value"] for b in sc["bridge"])
    changes = []
    if not dcf._close(sc["rate"], model["rate"]):
        changes.append(f"discount rate {model['rate']:.2%} → {sc['rate']:.2%}")
    if sc["valuation_date"] != model["valuation_date"]:
        changes.append(f"valuation date {model['valuation_date']} → {sc['valuation_date']}")
    words = {"end": "end of period", "mid": "mid-period", "mid-year": "mid-year (half a year before each period's end)",
             "actual/actual": "actual/actual (YEARFRAC)",
             "actual/365": "actual/365 (XNPV)"}
    if sc["timing"] != model["timing"]:
        changes.append(f"discounted at {words[model['timing']]} → {words[sc['timing']]}")
    if sc["day_count"] != model["day_count"]:
        changes.append(f"time measured by {words[model['day_count']]} → {words[sc['day_count']]}")
    if sc["terminal_date"] != model["terminal_date"]:
        changes.append(f"cut-off {model['terminal_date'] or 'none'} → {sc['terminal_date'] or 'none'}")
    dropped = [b["label"] for b, keep in zip(model["bridge"], list(include or []) + [True] * len(model["bridge"])) if not keep]
    if dropped:
        changes.append("bridge without " + ", ".join(dropped))
    iso = lambda d: d.isoformat() if d else None
    return {
        "selected": pick["cell"], "label": pick["label"], "validated": bool(pick.get("matches")),
        "model": {"rate": model["rate"], "pv": model["pv"], "total": model["total"], "workbook": model["compare_to"],
                  "valuation_date": iso(model["valuation_date"]), "timing": model["timing"],
                  "day_count": model["day_count"], "terminal_date": iso(model["terminal_date"]),
                  "bridge": model["bridge"]},
        "scenario": {"rate": sc["rate"], "pv": sc["pv"], "total": sc["total"], "valuation_date": iso(sc["valuation_date"]),
                     "timing": sc["timing"], "day_count": sc["day_count"], "terminal_date": iso(sc["terminal_date"]),
                     "bridge": sc["bridge"], "periods": sc["periods"], "first_period": iso(sc["first_period"]),
                     "last_period": iso(sc["last_period"])},
        "sensitivity": [{"case": "low", "rate": lo, "total": sc["sensitivity"][0]["total"],
                         "pv": sc["sensitivity"][0]["total"] - adj},
                        {"case": "scenario", "rate": sc["rate"], "total": sc["total"], "pv": sc["pv"]},
                        {"case": "high", "rate": hi, "total": sc["sensitivity"][1]["total"],
                         "pv": sc["sensitivity"][1]["total"] - adj}],
        "changes": changes, "units": sc.get("units") or [],
        "chart": _chart(db, db_path, base_in["cashflow"][0],
                        [("Cumulative PV (model)", model)] + ([("Cumulative PV (scenario)", sc)] if changes else []),
                        cumulative=True),
    }
