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
import math
import os
import re
import sqlite3
import threading
from datetime import date

import dcf

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

def read_factors(db, df_ref, cols: list[int]) -> dict | None:
    """Rate, valuation date, convention and cut-off that reproduce the workbook's factor row exactly."""
    sheet, row = df_ref[0], df_ref[1]
    theirs = {c: dcf._num(v) or 0.0 for c, v in db.execute(
        "SELECT col, value FROM cells WHERE sheet=? AND row=?", (sheet, row)) if c in cols}
    ends, ends_src = dcf.period_ends(db, sheet, cols)
    live = [c for c in sorted(ends) if 0 < theirs.get(c, 0.0) < 1]
    if not live:
        return None
    last = max(live, key=lambda c: ends[c])
    after = [c for c in ends if ends[c] > ends[last]]
    td = ends[last] if after and all(not theirs.get(c) for c in after) else None
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
                if all(abs(ours[c] - theirs.get(c, 0.0)) < 1e-9 for c in ends):
                    rate_cell = next((r for r in rates if abs(dcf._num(r["value"]) - rate) < 1e-9), None)
                    return {"rate": rate_cell["ref"] if rate_cell else round(rate, 12),
                            "rate_note": None if rate_cell else f"back-solved from the factors in {sheet}!r{row}",
                            "valuation_date": vd_c["ref"], "timing": timing, "day_count": dc,
                            "terminal_date": td.isoformat() if td else None, "ends_source": ends_src,
                            "factor_row": f"{sheet}!r{row} {dcf._row_label(db, sheet, row)}".strip()}
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
    return {**_public(v), "ok": True, "inputs": {
        "cashflow": [cf_range], "rate": fx["rate"], "valuation_date": fx["valuation_date"],
        "timing": fx["timing"], "day_count": fx["day_count"], "terminal_date": fx["terminal_date"],
        "adjustments": parts["items"], "compare_to": v["cell"]},
        "factor_row": fx["factor_row"], "rate_note": fx["rate_note"], "ends_source": fx["ends_source"]}


def _public(v):
    return {k: v[k] for k in ("cell", "label", "value", "pv_cell")}


def catalogue(db_path: str) -> list[dict]:
    """find() + build() for every DCF in a workbook, cached per model.db."""
    key = (db_path, os.path.getmtime(db_path))
    with _LOCK:
        if key in _CACHE:
            return _CACHE[key]
    db = sqlite3.connect(f"file:{db_path}?mode=ro", uri=True)
    out = []
    for v in find(db):
        b = build(db, v)
        if b["ok"]:
            r = dcf.compute(db, **b["inputs"], fix=False)
            b["matches"] = r["compare_to"] is not None and dcf._close(r["total"], r["compare_to"])
            b["rate_value"] = r["rate"]
        out.append(b)
    out.sort(key=lambda b: (not b.get("ok"), not b.get("matches"), _rank(b["label"]), b["cell"]))
    with _LOCK:
        _CACHE[key] = out
    return out


def _rank(label: str) -> int:
    for i, pat in enumerate((r"enterprise value", r"equity (value|pv)", r"\bnpv\b|present value|total valuation")):
        if re.search(pat, label or "", re.I):
            return i
    return 9


def view(db_path: str, cell: str | None = None, low: float | None = None, high: float | None = None) -> dict:
    """Everything the Valuation tab shows for one DCF. low / high are the discount rates for the low and high
    value cases (default: the model's rate +/- 1 percentage point)."""
    import chartdata
    import tools
    cat = catalogue(db_path)
    usable = [b for b in cat if b.get("ok")]
    listing = [{"cell": b["cell"], "label": b["label"], "value": b["value"], "ok": b.get("ok", False),
                "matches": b.get("matches", False), "reason": b.get("reason")} for b in cat]
    if not usable:
        return {"valuations": listing, "selected": None}
    pick = next((b for b in usable if b["cell"] == cell), usable[0])
    db = sqlite3.connect(f"file:{db_path}?mode=ro", uri=True)
    base_rate = pick["rate_value"]
    # rates as fractions (0.1425); the low value case uses the higher rate
    low = base_rate + 0.01 if low is None or low != low else low
    high = max(base_rate - 0.01, 0.0) if high is None or high != high else high
    r = dcf.compute(db, **pick["inputs"], rates=[low, high], fix=False)
    adj = sum(b["value"] for b in r["bridge"])
    cases = {"low": {"rate": low, "total": r["sensitivity"][0]["total"]},
             "base": {"rate": r["rate"], "total": r["total"]},
             "high": {"rate": high, "total": r["sensitivity"][1]["total"]}}
    for c in cases.values():
        c["pv"] = c["total"] - adj

    # the cash-flow chart: the workbook's row plus each period's present value at the model's rate
    with tools.using(db_path):
        spec = tools.chart("Cash flows and present values", [{"range": pick["inputs"]["cashflow"][0],
                                                               "name": "Cash flow (workbook)"}], kind="bar")
    ends, _ = dcf.period_ends(db, *dcf._row_range(db, pick["inputs"]["cashflow"][0])[::2])
    f = dcf.factors(ends, r["valuation_date"], r["rate"], r["timing"], r["day_count"], r["terminal_date"])
    cols = spec.get("columns") or []
    cfd = spec["series"][0]["data"]
    pvs = [(v * f[c]) if isinstance(v, (int, float)) and f.get(c) else None for v, c in zip(cfd, cols)]
    spec["series"].append({"name": f"Present value at {r['rate']:.2%}", "range": "computed in Python",
                           "data": pvs, "label": "Present value", "units": spec["series"][0].get("units")})
    spec = chartdata.enrich(spec, db)
    live = [i for i, c in enumerate(cols) if f.get(c)]
    if live:
        x_end, extra = live[-1], ""
        # chart rule F1: a period more than 5x the next largest (typically a terminal value at the end) flattens
        # everything else, so leave it out of the default view and say so
        vals = sorted(((abs(cfd[i]), i) for i in live if isinstance(cfd[i], (int, float))), reverse=True)
        if len(vals) > 2 and vals[1][0] and vals[0][0] > 5 * vals[1][0] and vals[0][1] >= live[-1] - 1:
            i = vals[0][1]
            x_end = i - 1
            extra = (f" The {spec['period_labels'][i]} cash flow ({cfd[i]:,.0f}, {vals[0][0] / vals[1][0]:.0f}x the next "
                     f"largest) is left out of this view so the rest is readable.")
        spec["view"] = {"x_start": live[0], "x_end": x_end}
        if spec.get("annual") and len(live) > 40:
            spec["mode"] = "annual"  # a few dozen annual bars read better than 100+ quarterly ones
        spec["note"] = (f"Showing the discounted periods ({r['first_period']:%b-%Y} to {r['last_period']:%b-%Y}).{extra} "
                        "“Full range” shows the whole timeline.")
    iso = lambda d: d.isoformat() if d else None
    return {
        "valuations": listing, "selected": pick["cell"],
        "result": dcf.card(r), "cases": cases, "chart": spec,
        "assumptions": {
            "rate": r["rate"], "rate_source": pick["rate_note"] or r["rate_source"],
            "valuation_date": iso(r["valuation_date"]), "valuation_date_source": r["valuation_date_source"],
            "terminal_date": iso(r["terminal_date"]),
            "terminal_source": "last period with a non-zero discount factor" if r["terminal_date"] else None,
            "timing": r["timing"], "day_count": r["day_count"], "cashflow": r["rows"], "factor_row": pick["factor_row"],
            "period_ends": pick["ends_source"], "periods": r["periods"], "first_period": iso(r["first_period"]),
            "last_period": iso(r["last_period"]), "undiscounted": r["undiscounted"], "units": r.get("units") or [],
        },
    }
