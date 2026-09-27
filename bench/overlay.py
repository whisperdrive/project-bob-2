"""The Python overlay: last year's valuation recreated as a live Python module, then re-fed from this year's model.

build() compiles the overlay sheets (xlcompile.py) into out/overlays/e<id>/overlay.py and checks it:
  1. every formula cell recomputed from the workbook's own inputs must equal the value Excel saved
  2. fed from the prior client model instead of the saved link values, the results must not move
  3. the results must tie to the report: conclusions at the printed precision, and each sensitivity in the
     report (e.g. "WACC 7.50%, TGR 2.25%") rerun by changing the discount rate and growth levers
Then a Session runs it live: change levers (the report's assumptions, found in the overlay) or any cell, and
choose where the client-model values come from:
  workbook   the values saved in the overlay workbook (for an external link: the link's cached values)
  prior      the prior client model file
  current    the current client model, rolled forward: each overlay period moves on by the roll (default: the
             months between the two valuation dates), client values are read from the same line item (sheet +
             label) and the period with the rolled date, and the valuation date lever is set to the new date.
             Anything that can't be matched is listed, never silently zero.
From a terminal, after the engagement's overlay is built in the app (step 6):
    uv run python bench/overlay.py 1                                  # engagement 1, as saved
    uv run python bench/overlay.py 1 --mode current --set Val_Inputs!C5=0.075
"""
import json
import math
import re
import sqlite3
import sys
import threading
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor
from datetime import date
from pathlib import Path

import xlcompile
import xlruntime
from xlruntime import from_db, same, serial, to_date

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "out" / "overlays"
LEVER_KEYS = re.compile(r"discount|wacc|terminal|growth|multiple|rab|valuation_date|net_debt|cpi|tax|cost_of|gearing|"
                        r"inflation|beta|premium", re.I)

sys.setrecursionlimit(1_000_000)
_EXEC = None


def deep(fn, *args, **kw):
    """Run fn on a thread with a 1 GB stack: a timeline recurrence can nest thousands of cells deep. One thread,
    so runs don't interleave (a Session's settings belong to the run that set them)."""
    global _EXEC
    if threading.current_thread().name.startswith("overlay-eval"):
        return fn(*args, **kw)
    if _EXEC is None:
        old = threading.stack_size(1024 * 1024 * 1024)
        try:
            _EXEC = ThreadPoolExecutor(1, thread_name_prefix="overlay-eval")
            _EXEC.submit(lambda: None).result()  # create the thread while the big stack size is set
        finally:
            threading.stack_size(old)
    return _EXEC.submit(fn, *args, **kw).result()


def _ro(path) -> sqlite3.Connection:
    return sqlite3.connect(f"file:{path}?mode=ro", uri=True, check_same_thread=False)


def _a1(sheet, row, col) -> str:
    from openpyxl.utils import get_column_letter
    return f"{sheet}!{get_column_letter(col)}{row}"


def parse_a1(ref: str) -> tuple[str, int, int]:
    from openpyxl.utils import column_index_from_string
    sheet, addr = ref.rsplit("!", 1)
    m = re.match(r"^\$?([A-Z]{1,3})\$?(\d+)$", addr.upper())
    return sheet.strip("'"), int(m.group(2)), column_index_from_string(m.group(1))


def add_months(s: float, months: int) -> float:
    """Serial date moved by whole months; a month-end stays a month-end."""
    d = to_date(s)
    end = xlruntime._add_months(d, 0, end=True) == d
    return serial(xlruntime._add_months(d, months, end=end))


def months_between(a: str, b: str) -> int:
    da, db_ = date.fromisoformat(a[:10]), date.fromisoformat(b[:10])
    return (db_.year - da.year) * 12 + (db_.month - da.month)


# ---- workbook data ------------------------------------------------------------------------------------------

class Workbook:
    """Cell values of one model.db, loaded a sheet at a time, plus line-item labels and the timeline per sheet."""

    def __init__(self, db_path: str):
        self.path = db_path
        self.db = _ro(db_path)
        self.sheets = {}
        self._labels = None
        self._timeline = {}
        self._inverse = {}

    def sheet(self, s):
        if s not in self.sheets:
            self.sheets[s] = {(r, c): from_db(v) for r, c, v in
                              self.db.execute("SELECT row, col, value FROM cells WHERE sheet=?", (s,))}
        return self.sheets[s]

    def value(self, s, r, c):
        return self.sheet(s).get((r, c))

    def labels(self):
        if self._labels is None:
            self._labels = {(s, r): (lab or "") for s, r, lab in self.db.execute("SELECT sheet, row, label FROM rows")}
        return self._labels

    def timeline(self, s) -> dict[int, float]:
        """col -> period date (serial) from the sheet's timeline row."""
        if s not in self._timeline:
            lay = self.db.execute("SELECT layout FROM sheets WHERE sheet=?", (s,)).fetchone()
            hr = json.loads(lay[0] or "{}").get("header_row") if lay else None
            tl = {}
            if hr:
                for (r, c), v in self.sheet(s).items():
                    if r == hr and isinstance(v, float) and 3000 < v < 120000:
                        tl[c] = v
            self._timeline[s] = tl
        return self._timeline[s]

    def column_of(self, s, when: float):
        """Column of a period date on the sheet's timeline, or None."""
        if s not in self._inverse:
            self._inverse[s] = {v: k for k, v in self.timeline(s).items()}
        return self._inverse[s].get(when)

    def header_row(self, s):
        lay = self.db.execute("SELECT layout FROM sheets WHERE sheet=?", (s,)).fetchone()
        return json.loads(lay[0] or "{}").get("header_row") if lay else None


class RowMap:
    """Prior (sheet, row) -> current (sheet, row) by line-item label (the n-th row with that label on the sheet)."""

    def __init__(self, prior: Workbook, current: Workbook):
        def occ(wb):
            seen, out = defaultdict(int), {}
            for (s, r), lab in sorted(wb.labels().items()):
                k = (s, re.sub(r"\s+", " ", lab.strip().lower()))
                seen[k] += 1
                out[(s, r)] = (*k, seen[k])
            return out
        self.prior, self.current = occ(prior), occ(current)
        self.back = {v: k for k, v in self.current.items()}

    def row(self, s, r):
        k = self.prior.get((s, r))
        if not k or not k[1]:
            return None
        hit = self.back.get(k)
        return hit[1] if hit and hit[0] == s else None


# ---- a live session -----------------------------------------------------------------------------------------

class Session:
    """A compiled overlay wired to its inputs and a feed. All evaluation goes through deep()."""

    def __init__(self, module_path: str, overlay_db: str, sheets: list[str], prior_db: str | None = None,
                 current_db: str | None = None, client_link: int | None = None, client_sheets: list[str] | None = None):
        src = Path(module_path).read_text()
        self.g = {"__name__": "overlay"}
        exec(compile(src, str(module_path), "exec"), self.g)
        self.B = self.g["B"]
        self.sheets = sheets
        self.ov = Workbook(overlay_db)
        self.prior = Workbook(prior_db) if prior_db and prior_db != overlay_db else None
        self.current = Workbook(current_db) if current_db else None
        self.client_link = client_link
        self.client_sheets = set(client_sheets or [])  # client sheets inside the overlay workbook (combined case)
        self.ext_cached = {}
        with _ro(overlay_db) as db:
            have = {r[0] for r in db.execute("SELECT name FROM sqlite_master WHERE type='table'")}
            if "extcells" in have:
                self.ext_cached = {(i, s, r, c): from_db(v) for i, s, r, c, v in
                                   db.execute("SELECT idx, sheet, row, col, value FROM extcells")}
        B = self.B
        for s in sheets:
            with _ro(overlay_db) as db:
                for r, c, v in db.execute("SELECT row, col, value FROM cells WHERE sheet=? AND formula IS NULL "
                                          "AND value IS NOT NULL", (s,)):
                    B.inputs[(s, r, c)] = from_db(v)
        self.formula_cells = []
        with _ro(overlay_db) as db:
            for s, r, c in db.execute(f"SELECT sheet, row, col FROM cells WHERE formula IS NOT NULL AND sheet IN "
                                      f"({','.join('?' * len(sheets))})", sheets):
                self.formula_cells.append((s, r, c))
        self.formula_cells.sort(key=lambda k: (k[2], k[1]))  # column by column: early periods are memoised first
        B.cached = lambda s, r, c: self.ov.value(s, r, c)
        self.rowmap = RowMap(self.prior or self.ov, self.current) if self.current else None
        self.mode = None
        self.configure("workbook")

    # feeds
    def configure(self, mode: str, overrides: dict | None = None, shift_months: int = 0, roll_dates: bool = True):
        """mode: workbook | prior | current. overrides: {(sheet, row, col): value}. shift_months: roll-forward."""
        B = self.B
        self.mode, self.shift = mode, shift_months
        self.unmatched = {}
        B.overrides.clear()
        B.reset()
        if mode == "workbook":
            B.feed = lambda s, r, c: self.ov.value(s, r, c)
            B.ext = lambda i, s, r, c: self.ext_cached.get((i, s, r, c))
        elif mode == "prior":
            src = self.prior or self.ov
            B.feed = lambda s, r, c: src.value(s, r, c) if s in self.client_sheets else self.ov.value(s, r, c)
            B.ext = lambda i, s, r, c: (self.prior.value(s, r, c) if self.prior and i == self.client_link
                                        else self.ext_cached.get((i, s, r, c)))
        elif mode == "current":
            if not self.current:
                raise ValueError("no current client model to feed from")
            prior = self.prior or self.ov
            B.feed = lambda s, r, c: self._rolled(prior, s, r, c) if s in self.client_sheets else self.ov.value(s, r, c)
            B.ext = lambda i, s, r, c: (self._rolled(prior, s, r, c) if i == self.client_link
                                        else self.ext_cached.get((i, s, r, c)))
            if roll_dates and shift_months:
                for key, v in self.rolled_timeline(shift_months).items():
                    B.overrides[key] = v
        else:
            raise ValueError(f"unknown feed {mode}")
        for k, v in (overrides or {}).items():
            B.overrides[k] = v
        B.range_cache.clear()

    def _rolled(self, prior: Workbook, s, r, c):
        """Prior client cell -> the current model's value: same line item (label), period rolled on by the shift."""
        cur = self.current
        r2 = self.rowmap.row(s, r)
        if r2 is None:
            self.unmatched[(s, r, c)] = f"line item '{prior.labels().get((s, r), '')}' not in the current model"
            return None
        tl_p = prior.timeline(s)
        c2 = c
        if c in tl_p:
            want = add_months(tl_p[c], self.shift)
            c2 = cur.column_of(s, want)
            if c2 is None:
                self.unmatched[(s, r, c)] = f"period {to_date(want).isoformat()} not in the current model"
                return None
        return cur.value(s, r2, c2)

    def rolled_timeline(self, months: int) -> dict:
        """The overlay's own period dates moved on: constants in each overlay sheet's timeline row."""
        out = {}
        for s in self.sheets:
            hr = self.ov.header_row(s)
            if not hr:
                continue
            for (r, c), v in self.ov.sheet(s).items():
                if r == hr and (s, r, c) in self.B.inputs and isinstance(v, float) and 3000 < v < 120000:
                    out[(s, r, c)] = add_months(v, months)
        return out

    # evaluation
    def value(self, s, r, c):
        fn = self.B.rows.get((s, r))
        return fn(c) if fn else self.B.input(s, r, c)

    def values(self, cells):
        return deep(lambda: [self.value(*k) for k in cells])

    def evaluate_all(self):
        return deep(lambda: {k: self.value(*k) for k in self.formula_cells})

    def validate(self, limit: int = 60) -> dict:
        """Every formula cell vs the value Excel saved (feed = workbook)."""
        self.configure("workbook")
        xlruntime.xl.quirks.clear()
        res = self.evaluate_all()
        bad = []
        for k, v in res.items():
            w = self.ov.value(*k)
            if not same(v, w):
                bad.append((k, v, w))
        with _ro(self.ov.path) as db:
            forms = {}
            for k, _, _ in bad[:limit]:
                forms[k] = db.execute("SELECT formula FROM cells WHERE sheet=? AND row=? AND col=?", k).fetchone()[0]
        labels = self.ov.labels()
        return {"cells": len(res), "matched": len(res) - len(bad), "cycles": len(self.B.cycles),
                "unsupported": dict(self.B.unsupported), "quirks": dict(xlruntime.xl.quirks),
                "mismatches": [{"cell": _a1(*k), "label": labels.get(k[:2], ""), "python": _show(v), "workbook": _show(w),
                                "formula": forms.get(k)} for k, v, w in bad[:limit]]}

    def outputs(self, cells: list[tuple]) -> list:
        return self.values(cells)


def _show(v):
    if isinstance(v, xlruntime.XLError):
        return v.code
    if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
        return str(v)
    return v


# ---- building and checking ----------------------------------------------------------------------------------

def _num_in(text: str):
    m = re.search(r"-?\d[\d,]*(?:\.\d+)?", text or "")
    return (float(m.group(0).replace(",", "")), len(m.group(0).split(".")[1]) if "." in m.group(0) else 0) if m else None


def levers_and_outputs(overlay_db: str, sheets: list[str], facts: list[dict]) -> tuple[list, list]:
    """Levers: the report's assumptions found in the overlay (label agrees). Outputs: the overlay's DCF anchors and
    the cells holding the report's conclusions."""
    import linkmap
    import valuation
    fm = linkmap.match_facts(overlay_db, facts, set(sheets))
    by_id = {f["id"]: f for f in facts}
    with _ro(overlay_db) as db:
        formula = {(s, r, c) for s, r, c in db.execute(
            f"SELECT sheet, row, col FROM cells WHERE formula IS NOT NULL AND sheet IN ({','.join('?' * len(sheets))})", sheets)}
    from openpyxl.utils import column_index_from_string as ci
    levers, outputs, seen = [], [], set()
    for m in fm:
        f = by_id[m["fact_id"]]
        best = next((x for x in m["matches"] if x["label_match"] or x["anchor"]), None)
        if not best:
            continue
        cell = (best["sheet"], best["row"], ci(re.sub(r"\d", "", best["addr"])))
        if f["category"] in ("assumption", "identity") and LEVER_KEYS.search(f["key"] or "") and cell not in seen:
            seen.add(cell)
            levers.append({"key": f["key"], "label": f.get("label") or f["key"], "cell": _a1(*cell), "report": f.get("value_text"),
                           "unit": f.get("unit"), "scale": best.get("scale", 1.0), "formula": cell in formula,
                           "row_label": best["label"]})
        if f["category"] == "conclusion":
            outputs.append({"cell": _a1(*cell), "label": best["label"] or f.get("label"), "fact_id": f["id"],
                            "key": f["key"], "report": f.get("value_text"), "low": f.get("low_text"), "high": f.get("high_text"),
                            "unit": f.get("unit"), "scale": best.get("scale", 1.0), "sign": best.get("sign", 1)})
    try:
        for a in valuation.catalogue(overlay_db):
            if a.get("ok") and a["cell"].split("!")[0] in sheets and a["cell"] not in {o["cell"] for o in outputs}:
                outputs.append({"cell": a["cell"], "label": a["label"], "fact_id": None, "key": None, "report": None,
                                "unit": None, "scale": 1.0, "sign": 1})
    except Exception:
        pass
    return levers, outputs


def sensitivity_points(facts: list[dict], levers: list[dict], outputs: list[dict]) -> list[dict]:
    """Report sensitivities as scenarios: 'WACC 7.50%, TGR 2.25%' -> discount-rate and growth lever values."""
    lev = {}
    for l in levers:
        k = l["key"].lower()
        if re.search(r"discount|wacc", k):
            lev.setdefault("rate", l)
        elif re.search(r"terminal|growth", k):
            lev.setdefault("growth", l)
        elif re.search(r"multiple", k):
            lev.setdefault("multiple", l)
    pts = []
    for f in facts:
        if f["category"] != "sensitivity":
            continue
        basis = f.get("basis") or f.get("label") or ""
        sets, why = {}, []
        for name, pat in (("rate", r"(?:wacc|discount rate|\bdr\b)\s*(?:of|=|:)?\s*(-?[\d.]+)\s*%"),
                          ("growth", r"(?:tgr|terminal growth|growth)\s*(?:rate)?\s*(?:of|=|:)?\s*(-?[\d.]+)\s*%"),
                          ("multiple", r"(?:multiple)\s*(?:of|=|:)?\s*([\d.]+)\s*x")):
            m = re.search(pat, basis, re.I)
            if m:
                if name not in lev:
                    why.append(f"no {name} lever found in the overlay")
                    continue
                # the report prints 7.50%; the cell holds 0.075 (or 7.5 if the model keeps percentages as numbers)
                sets[lev[name]["cell"]] = float(m.group(1)) * (lev[name].get("scale") or 1.0)
        key = (f.get("key") or "") + " " + (f.get("label") or "")
        out = next((o for o in outputs if o.get("key") and (("equity" in key.lower() and "equity" in o["key"])
                                                            or ("enterprise" in key.lower() and "enterprise" in o["key"]))), None)
        if not out:
            why.append("no matching conclusion cell (approve the equity / enterprise value facts)")
        pts.append({"fact_id": f["id"], "basis": basis, "report": f.get("value_text"), "set": sets,
                    "output": out["cell"] if out else None, "scale": out["scale"] if out else 1.0,
                    "skip": "; ".join(why) if why or not sets else None})
    return pts


def tie(value, report_text, scale=1.0, sign=1) -> dict | None:
    """Does a Python value round to the report's printed figure?"""
    n = _num_in(report_text)
    if n is None or not isinstance(value, float):
        return None
    x, d = n
    shown = sign * value / (scale or 1.0)
    return {"report": report_text, "python": round(shown, d + 2), "ok": abs(round(shown, d) - x) <= 0.5 * 10 ** -d + 1e-9}


def build(out_dir: Path, overlay: dict, prior: dict | None, current: dict | None, facts: list[dict],
          title: str, client_link: int | None, prior_val_date: str | None, progress=None) -> dict:
    """overlay / prior / current: {"db_path", "filename", "sheets"}. Returns the summary saved to overlay.json."""
    progress = progress or (lambda f, m: None)
    out_dir.mkdir(parents=True, exist_ok=True)
    sheets = overlay["sheets"]
    progress(0.05, "Compiling the overlay's formulas to Python")
    src, stats = deep(xlcompile.compile_overlay, overlay["db_path"], sheets, title,
                      lambda f, m: progress(0.05 + 0.3 * f, m))
    module = out_dir / "overlay.py"
    module.write_text(src)
    progress(0.4, "Loading the module")
    same_file = prior is not None and prior["db_path"] == overlay["db_path"]
    sess = Session(str(module), overlay["db_path"], sheets, None if same_file else (prior or {}).get("db_path"),
                   (current or {}).get("db_path"), client_link, (prior or {}).get("sheets") if same_file else None)
    progress(0.5, f"Recomputing {len(sess.formula_cells):,} formula cells and checking each against Excel")
    val = sess.validate()
    progress(0.7, "Finding levers and outputs")
    levers, outputs = levers_and_outputs(overlay["db_path"], sheets, facts)
    out_cells = [parse_a1(o["cell"]) for o in outputs]
    lever_cells = [parse_a1(l["cell"]) for l in levers]
    base = sess.values(out_cells + lever_cells)
    for o, v in zip(outputs, base[:len(outputs)]):
        o["value"] = _show(v)
        o["tie"] = tie(v, o["report"], o["scale"], o["sign"]) if o.get("report") else None
    for l, v in zip(levers, base[len(outputs):]):
        l["value"] = _show(v)
    feeds = {}
    if prior and not same_file:
        progress(0.78, "Feeding it from the prior client model")
        sess.configure("prior")
        pv = sess.values(out_cells)
        feeds["prior"] = {"outputs": [_show(v) for v in pv], "same": all(same(a, b) for a, b in zip(pv, base[:len(outputs)]))}
    progress(0.85, "Rerunning the report's sensitivities")
    points = sensitivity_points(facts, levers, outputs)
    for p in points:
        if p["skip"] or not p["output"]:
            continue
        sess.configure("workbook", {parse_a1(k): v for k, v in p["set"].items()})
        v = sess.values([parse_a1(p["output"])])[0]
        p["python"] = _show(v)
        p["tie"] = tie(v, p["report"], p["scale"])
    sess.configure("workbook")
    roll = None
    if current:
        months = roll_months(sess, prior, overlay, same_file)
        vd_lever = next((l for l in levers if l["key"] == "valuation_date"), None)
        roll = {"prior_valuation_date": prior_val_date, "months": months, "valuation_date_cell": vd_lever["cell"] if vd_lever else None,
                "current_valuation_date": to_date(add_months(serial(date.fromisoformat(prior_val_date[:10])), months)).isoformat()
                if prior_val_date else None}
    summary = {"module": str(module), "stats": {k: v for k, v in stats.items() if k != "not_compiled"},
               "not_compiled": stats["not_compiled"][:50], "validation": val, "levers": levers, "outputs": outputs,
               "feeds": feeds, "sensitivities": points, "roll": roll, "sheets": sheets,
               "files": {k: (v or {}).get("filename") for k, v in (("overlay", overlay), ("prior", prior), ("current", current))},
               "client_link": client_link, "same_file": same_file}
    (out_dir / "overlay.json").write_text(json.dumps(summary, default=str, indent=1))
    progress(1.0, "Done")
    return summary, sess


def roll_months(sess: Session, prior: dict | None, overlay: dict, same_file: bool) -> int:
    """How far the client model's timeline moved: first period of the current model vs the prior one, on the
    client sheets the overlay reads (12 if they can't be compared)."""
    src = sess.prior or sess.ov
    sheets = (prior or {}).get("sheets") if same_file else None
    for s in sorted(sheets or {k[1] for k in sess.ext_cached} or []):
        a, b = src.timeline(s), sess.current.timeline(s) if sess.current else {}
        if a and b:
            return months_between(to_date(min(a.values())).isoformat(), to_date(min(b.values())).isoformat())
    return 12


def scenario(sess: Session, summary: dict, mode: str, changes: dict, valuation_date: str | None = None,
             months: int | None = None) -> dict:
    """Run the overlay with lever / cell changes ({"Sheet!A1": value}) on a feed. Returns outputs against the base
    (same feed, no changes) and the workbook, what the roll-forward changed, and feed cells it couldn't match."""
    outputs = summary["outputs"]
    cells = [parse_a1(o["cell"]) for o in outputs]
    extra = {parse_a1(k): v for k, v in changes.items()}
    roll_info, defaults = None, {}
    if mode == "current":
        roll = summary.get("roll") or {}
        months = months if months is not None else roll.get("months", 12)
        valuation_date = valuation_date or roll.get("current_valuation_date")
        vd_cell = roll.get("valuation_date_cell")
        if valuation_date and vd_cell:  # the roll-forward's own settings are the base; the person's changes go on top
            defaults[parse_a1(vd_cell)] = serial(date.fromisoformat(valuation_date[:10]))
        roll_info = {"months": months, "valuation_date": valuation_date, "valuation_date_cell": vd_cell}
    sess.configure(mode, defaults, months or 0)
    base = sess.values(cells)
    base_unmatched = dict(sess.unmatched)
    sess.configure(mode, {**defaults, **extra}, months or 0)
    got = sess.values(cells)
    unmatched = {**base_unmatched, **sess.unmatched}
    if mode == "current":
        tl = sess.rolled_timeline(months or 0)
        firsts = sorted(tl.values())
        roll_info.update(first_period=to_date(firsts[0]).isoformat() if firsts else None,
                         last_period=to_date(firsts[-1]).isoformat() if firsts else None, periods=len(set(firsts)))
    rows = []
    for o, b, v in zip(outputs, base, got):
        wb = o.get("value")
        rows.append({"cell": o["cell"], "label": o["label"], "workbook": wb, "base": _show(b), "value": _show(v),
                     "change": (v - b) if isinstance(v, float) and isinstance(b, float) else None,
                     "vs_workbook": (v - wb) if isinstance(v, float) and isinstance(wb, float) else None,
                     "report": o.get("report"), "tie": tie(v, o["report"], o["scale"], o["sign"]) if o.get("report") and mode != "current" else None})
    sess.configure("workbook")
    return {"mode": mode, "changes": changes, "outputs": rows, "roll": roll_info,
            "unmatched": [{"cell": _a1(*k), "why": w} for k, w in list(unmatched.items())[:100]],
            "n_unmatched": len(unmatched)}


def inputs_list(sess: Session, q: str = "", limit: int = 80) -> list[dict]:
    """The overlay's constants (every one can be changed), searchable by label or cell."""
    labels = sess.ov.labels()
    out = []
    ql = q.lower().strip()
    for (s, r, c), v in sorted(sess.B.inputs.items(), key=lambda kv: (sess.sheets.index(kv[0][0]), kv[0][1], kv[0][2])):
        if not isinstance(v, float):
            continue
        lab = labels.get((s, r), "")
        ref = _a1(s, r, c)
        if ql and ql not in lab.lower() and ql not in ref.lower():
            continue
        out.append({"cell": ref, "label": lab, "value": v})
        if len(out) >= limit:
            break
    return out


def trace(sess: Session, cell: str, depth: int = 2) -> dict:
    """The Python behind one cell: its row function's source, and the client-model values it reads."""
    import inspect
    s, r, c = parse_a1(cell)
    fn = sess.B.rows.get((s, r))
    src = inspect.getsource(fn.raw) if fn else None
    sess.B.feed_log = {}
    try:
        sess.B.reset()
        v = sess.values([(s, r, c)])[0]
        feeds = [{"source": "this workbook" if k[0] == "" else f"external link [{k[0]}]", "cell": _a1(k[1], k[2], k[3]),
                  "value": _show(val)} for k, val in list(sess.B.feed_log.items())[:200]]
    finally:
        sess.B.feed_log = None
    return {"cell": cell, "value": _show(v), "label": sess.B.labels.get((s, r)), "source": src,
            "client_reads": feeds, "n_client_reads": len(feeds)}


if __name__ == "__main__":
    import argparse
    sys.path.insert(0, str(Path(__file__).parent))
    ap = argparse.ArgumentParser(description="Run an engagement's Python overlay")
    ap.add_argument("engagement", type=int)
    ap.add_argument("--mode", default="workbook", choices=["workbook", "prior", "current"])
    ap.add_argument("--set", action="append", default=[], metavar="SHEET!A1=VALUE", help="change an input (repeatable)")
    ap.add_argument("--valuation-date", help="YYYY-MM-DD, for --mode current")
    a = ap.parse_args()
    import engagement
    changes = dict(x.split("=", 1) for x in a.set)
    res = engagement.overlay_run(a.engagement, a.mode, changes, a.valuation_date, None)
    if res.get("roll"):
        r = res["roll"]
        print(f"rolled forward {r['months']} months to {r['valuation_date']}: periods {r.get('first_period')} to "
              f"{r.get('last_period')} ({r.get('periods')}); {res['n_unmatched']} client value(s) unmatched")
    for o in res["outputs"]:
        ch = f"  change {o['change']:+,.2f}" if isinstance(o.get("change"), float) and changes else ""
        print(f"{o['cell']:<18} {o['label'][:40]:<40} {o['value']:>16,.2f}{ch}")
