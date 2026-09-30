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
from collections import Counter, defaultdict
from concurrent.futures import ThreadPoolExecutor
from datetime import date
from pathlib import Path

import rodb
import xlcompile
import xlruntime
from xlruntime import from_db, same, serial, to_date

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "out" / "overlays"
LEVER_KEYS = re.compile(r"discount|wacc|terminal|growth|multiple|rab|valuation_date|net_debt|cpi|tax|cost_of|gearing|"
                        r"inflation|beta|premium", re.I)

sys.setrecursionlimit(1_000_000)
# Windows accepts a thread stack strictly under 256 MB (and commits all of it up front); other systems want a whole
# number of pages. 255 MB suits both, and smaller sizes are tried if a system refuses it.
STACK = 255 * 1024 * 1024
STACK_FALLBACKS = (STACK, 192 * 1024 * 1024, 128 * 1024 * 1024, 64 * 1024 * 1024)
_EXEC = None


def _big_stack() -> int:
    """Set the next thread's stack to the largest size this system accepts; returns the size it had."""
    for size in STACK_FALLBACKS:
        try:
            return threading.stack_size(size)
        except ValueError:  # "size not valid" (Windows: 256 MB or more; some systems: not a page multiple)
            continue
    return threading.stack_size()


def deep(fn, *args, **kw):
    """Run fn on a thread with a 255 MB stack: a timeline recurrence can nest thousands of cells deep. One thread,
    so runs don't interleave: a Session's settings belong to the run that set them."""
    global _EXEC
    if threading.current_thread().name.startswith("overlay-eval"):
        return fn(*args, **kw)
    if _EXEC is None:
        old = _big_stack()
        try:
            _EXEC = ThreadPoolExecutor(1, thread_name_prefix="overlay-eval")
            _EXEC.submit(lambda: None).result()  # create the thread while the big stack size is set
        finally:
            threading.stack_size(old)
    return _EXEC.submit(fn, *args, **kw).result()


def _ro(path) -> sqlite3.Connection:
    return rodb.connect(path, check_same_thread=False)


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

    def close(self):
        self.db.close()

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
    """Prior (sheet, row) -> current (sheet, row), by line-item label:
      1. the n-th row with that label on the sheet (rows inserted above it don't matter)
      2. else the row with that label nearest its old position (the label now occurs more or fewer times)
      3. else, on a sheet laid out as before (most labelled rows still at the same row), the same row, which also
         covers rows without a label (flags, timing rows)
    why() says what failed for a row it can't map."""

    def __init__(self, prior: Workbook, current: Workbook):
        norm = lambda lab: re.sub(r"\s+", " ", (lab or "").strip().lower())

        def occ(wb):
            seen, out = defaultdict(int), {}
            for (s, r), lab in sorted(wb.labels().items()):
                k = (s, norm(lab))
                seen[k] += 1
                out[(s, r)] = (*k, seen[k])
            return out
        self.prior, self.current = occ(prior), occ(current)
        self.back = {v: k for k, v in self.current.items()}
        self.by_label = defaultdict(list)
        for (s, r), (_, lab, _) in self.current.items():
            if lab:
                self.by_label[(s, lab)].append(r)
        self.cur_sheets = {s for (s,) in current.db.execute("SELECT sheet FROM sheets")}
        cur_labels = {k: norm(v) for k, v in current.labels().items()}
        same, total = Counter(), Counter()
        for (s, r), (_, lab, _) in self.prior.items():
            if lab:
                total[s] += 1
                same[s] += cur_labels.get((s, r)) == lab
        self.same_layout = {s for s in total if s in self.cur_sheets and same[s] >= 0.8 * total[s]}

    def row(self, s, r):
        return self.match(s, r)[0]

    def match(self, s, r) -> tuple[int | None, str]:
        """(current row or None, how it was matched)."""
        k = self.prior.get((s, r))
        if k and k[1]:
            hit = self.back.get(k)
            if hit and hit[0] == s:
                return hit[1], "same label" if k[2] == 1 else f"same label, occurrence {k[2]} on the sheet"
            rows = self.by_label.get((s, k[1]))
            if rows:
                return min(rows, key=lambda x: abs(x - r)), "same label, the nearest of its rows (it occurs a different number of times now)"
        if s in self.same_layout:
            return r, "same row (no label to follow; the sheet's layout is unchanged)"
        return None, "unmatched"

    def why(self, s, r, labels: dict) -> str:
        lab = labels.get((s, r), "")
        if s not in self.cur_sheets:
            return f"sheet {s} isn't in the current model"
        if not lab:
            return f"row {r} of {s} has no line-item label, and the sheet's layout changed, so it can't be followed"
        return f"line item '{lab}' isn't on {s} in the current model"


# ---- a live session -----------------------------------------------------------------------------------------

class Session:
    """A compiled overlay wired to its inputs and a feed. All evaluation goes through deep()."""

    def __init__(self, module_path: str, overlay_db: str, sheets: list[str], prior_db: str | None = None,
                 current_db: str | None = None, client_link: int | None = None, client_sheets: list[str] | None = None):
        src = Path(module_path).read_text(encoding="utf-8")
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
        self.rowmap = None
        if self.current:  # last year's rows in this year's model: by label, history, words, neighbours, banner
            import rowfind
            base = self.prior or self.ov
            self.rowmap = rowfind.RowFinder(RowMap(base, self.current), base, self.current)
        self.stood_in = {}  # (sheet, row, col) -> last year's value, used where this year's model has no match
        self.blank = {}  # (sheet, row, col) -> this year's cell found for it, blank where last year's had a value
        self.derived = {}  # (sheet, row) -> timing rule: a period flag or date row worked out from the period dates
        self.derived_used = {}  # (sheet, row, col) -> the rule, where a value was worked out on the current feed
        self.client_reads = set()  # (sheet, row, col) of last year's model read on the current feed
        self.base_vd = None  # last year's valuation date (serial): the roll's start
        self._pshift = {}
        self.horizon_set = None  # "fixed" / "rolling": the engagement's profile says, not worked out (fixed_horizon)
        self.mode = None
        self.holds = {}  # (sheet, row, col) -> value: cells held at Excel's value on every feed (the doctor's fixes)
        self.configure("workbook")

    # feeds
    def configure(self, mode: str, overrides: dict | None = None, shift_months: int = 0, roll_dates: bool = True):
        """mode: workbook | prior | current. overrides: {(sheet, row, col): value}. shift_months: roll-forward."""
        B = self.B
        self.mode, self.shift = mode, shift_months
        self.unmatched = {}
        self.stood_in = {}
        self.blank = {}
        self.derived_used = {}
        self.client_reads = set()
        B.overrides.clear()
        B.overrides.update(self.holds)
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
        """Prior client cell -> the current model's value: the same line item (found by rowfind, wherever it is
        now), the period rolled on by the shift. Where this year's model has no match, last year's value stands
        in (its forecast for the same period), recorded in unmatched and stood_in: never a blank, which would
        read as zero. That includes a row and period that are found but blank this year where last year's cell
        had a value (a row matched on its history that stops where the history does): recorded in blank too."""
        cur = self.current
        self.client_reads.add((s, r, c))
        tl_p = prior.timeline(s)
        want = add_months(tl_p[c], self.period_shift(prior, s)) if c in tl_p else None
        rule = self.derived.get((s, r))
        if rule and want is not None:  # timing worked out from the rolled period date, not hunted for
            self.derived_used[(s, r, c)] = rule["text"]
            new_vd = add_months(self.base_vd, self.shift) if self.base_vd is not None else None
            return timing_value(rule, want, new_vd)
        hit = self.rowmap.locate(s, r)
        if hit is None:
            self.unmatched[(s, r, c)] = self.rowmap.why(s, r, prior.labels())
            return self._stand_in(prior, s, r, c, want)
        s2, r2 = hit
        c2 = c
        if want is not None:
            c2 = cur.column_of(s2, want)
            if c2 is None:
                self.unmatched[(s, r, c)] = f"period {to_date(want).isoformat()} not in the current model"
                return self._stand_in(prior, s, r, c, want)
        v = cur.value(s2, r2, c2)
        if v is None and prior.value(s, r, c) is not None:
            self.blank[(s, r, c)] = (s2, r2, c2)
            self.unmatched[(s, r, c)] = (f"found at {s2}!r{r2}, but blank there"
                                         + (f" for {to_date(want).isoformat()}" if want is not None else ""))
            return self._stand_in(prior, s, r, c, want)
        return v

    def _stand_in(self, prior: Workbook, s, r, c, want):
        """Last year's value for a client cell this year's model doesn't have: last year's forecast for the same
        period; where last year's model doesn't reach that period, its own cell (the period it read last year);
        the same cell where the row has no timeline. Blank only where last year's cell was blank too."""
        c0 = prior.column_of(s, want) if want is not None else None
        v = prior.value(s, r, c0 if c0 is not None else c)
        if v is not None:
            self.stood_in[(s, r, c)] = v
        return v

    def fixed_horizon(self) -> bool:
        """Whether rolling forward keeps the period dates: set in the engagement's profile (horizon_set), else
        worked out (horizon()). Cached until the roll is planned again."""
        if ("fixed",) not in self._pshift:
            set_ = getattr(self, "horizon_set", None)
            if set_ in ("fixed", "rolling"):
                self._pshift[("fixed",)] = set_ == "fixed"
            elif not self.current or not self.rowmap:
                self._pshift[("fixed",)] = False
            else:
                sheets = self.client_sheets or {k[1] for k in self.ext_cached}
                self._pshift[("fixed",)] = horizon(self.prior or self.ov, self.current, sheets, self.rowmap.sheet_for)[0] == "fixed"
        return self._pshift[("fixed",)]

    def period_shift(self, wb: Workbook, s: str) -> int:
        """How far a sheet's periods move when rolling forward, in months: the whole periods that ended between
        last year's valuation date and the new one. The valuation date can move three months while an annual
        sheet's periods stay put (FY2026 is still FY2026) and a quarterly sheet's move one quarter; moving an
        annual sheet by three months would land on dates it doesn't have. Nothing on a fixed horizon (this year's
        model ends where last year's did): the dates are the same dates, only the valuation date moves."""
        if not self.shift:
            return 0
        key = (wb.path, s, self.shift)
        if key not in self._pshift:
            tl = sorted(set(wb.timeline(s).values()))
            if self.fixed_horizon():
                out = 0
            elif len(tl) < 2:
                out = 0
            else:
                gaps = sorted(b - a for a, b in zip(tl, tl[1:]))
                plen = max(1, round(gaps[len(gaps) // 2] / 30.44))
                # a timeline of period starts (1 July, 1 October: all on the 1st) is counted by its periods' ends,
                # the day before the next start: a year that began in the window hasn't ended in it
                if all(to_date(d).day == 1 for d in tl):
                    ends = [add_months(d, plen) - 1 for d in tl]
                else:
                    ends = tl
                if self.base_vd is not None:
                    new = add_months(self.base_vd, self.shift)
                    out = plen * sum(1 for d in ends if self.base_vd < d <= new)
                else:
                    out = plen * (max(0, self.shift) // plen)
            self._pshift[key] = out
        return self._pshift[key]

    def rolled_timeline(self, months: int) -> dict:
        """The overlay's own period dates moved on: constants in each overlay sheet's timeline row, by whole
        periods (period_shift), not by the valuation date's move."""
        out = {}
        keep = self.shift
        self.shift = months
        try:
            for s in self.sheets:
                hr = self.ov.header_row(s)
                step = self.period_shift(self.ov, s)
                if not hr or not step:
                    continue
                for (r, c), v in self.ov.sheet(s).items():
                    if r == hr and (s, r, c) in self.B.inputs and isinstance(v, float) and 3000 < v < 120000:
                        out[(s, r, c)] = add_months(v, step)
        finally:
            self.shift = keep
        return out

    # evaluation
    def close(self):
        """Let go of the model.db files (a workbook being rebuilt deletes its model.db)."""
        for wb in (self.ov, self.prior, self.current):
            if wb:
                wb.close()

    def paths(self) -> set[str]:
        return {wb.path for wb in (self.ov, self.prior, self.current) if wb}

    def value(self, s, r, c):
        """A cell as the module sees it: overlay sheets computed, other sheets from the feed."""
        return self.B.get("", s, r, c)

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
        return {"cells": len(res), "matched": len(res) - len(bad), "cycles": len(self.B.cycles), "runtime": xlruntime.RUNTIME,
                "text": sum(isinstance(w, str) for _, _, w in bad),  # label cells: Excel saved text (a caption)
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
          title: str, client_link: int | None, prior_val_date: str | None, progress=None,
          client_sheets: list[str] | None = None) -> dict:
    """overlay / prior / current: {"db_path", "filename", "sheets"}. Returns the summary saved to overlay.json.
    client_sheets: the overlay workbook's own copy of the client model's sheets, when the overlay sits in a copy of
    a client model that is also here as its own file (prior): those sheets are then fed from that file."""
    progress = progress or (lambda f, m: None)
    out_dir.mkdir(parents=True, exist_ok=True)
    sheets = overlay["sheets"]
    progress(0.05, "Compiling the overlay's formulas to Python")
    src, stats = deep(xlcompile.compile_overlay, overlay["db_path"], sheets, title,
                      lambda f, m: progress(0.05 + 0.3 * f, m), overlay.get("source_path"))
    module = out_dir / "overlay.py"
    module.write_text(src, encoding="utf-8")
    progress(0.4, "Loading the module")
    same_file = prior is not None and prior["db_path"] == overlay["db_path"]
    sess = Session(str(module), overlay["db_path"], sheets, None if same_file else (prior or {}).get("db_path"),
                   (current or {}).get("db_path"), client_link,
                   client_sheets or ((prior or {}).get("sheets") if same_file else None))
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
        vd_lever = next((l for l in levers if l["key"] == "valuation_date"), None)
        roll = plan_roll(sess, prior, overlay, same_file, prior_val_date, (prior or {}).get("valuation_date"),
                         current.get("valuation_date"))
        roll["valuation_date_cell"] = vd_lever["cell"] if vd_lever else None
    summary = {"module": str(module), "stats": {k: v for k, v in stats.items() if k != "not_compiled"},
               "not_compiled": stats["not_compiled"][:50], "validation": val, "levers": levers, "outputs": outputs,
               "feeds": feeds, "sensitivities": points, "roll": roll, "sheets": sheets,
               "files": {k: (v or {}).get("filename") for k, v in (("overlay", overlay), ("prior", prior), ("current", current))},
               "client_link": client_link, "same_file": same_file}
    (out_dir / "overlay.json").write_text(json.dumps(summary, default=str, indent=1), encoding="utf-8")
    progress(1.0, "Done")
    return summary, sess


def horizon(prior: Workbook, current: Workbook, sheets, sheet_for=None) -> tuple[str | None, dict]:
    """How this year's client model's periods end against last year's: "fixed" where most client sheets with a
    timeline end on the same date both years (a concession or asset life with an end date: the period dates are the
    same dates, so rolling forward moves the valuation date, not the periods; moving them would read periods past
    the end and stand last year's values in for them), "rolling" where they end later, None with no sheet to tell
    by. sheet_for: last year's sheet -> this year's (rowfind), else the same name. -> (kind, {"fixed": n, "rolling": n})."""
    n = {"fixed": 0, "rolling": 0}
    for s in sorted(sheets or []):
        to = sheet_for(s) if sheet_for else s
        a, b = prior.timeline(s), current.timeline(to) if to else {}
        if len(a) >= 2 and len(b) >= 2:
            n["fixed" if max(a.values()) == max(b.values()) else "rolling"] += 1
    if not n["fixed"] and not n["rolling"]:
        return None, n
    return ("fixed" if n["fixed"] > n["rolling"] else "rolling"), n


def roll_months(sess: Session, prior: dict | None, overlay: dict, same_file: bool,
                dates: tuple = (None, None, None, None)) -> tuple[int, str, str | None]:
    """How far to roll forward: (months, how that was worked out, the new valuation date).
    dates: (last year's valuation date: the overlay's, from the report; last year's client model's; this year's
    client model's; this year's valuation date as set for the engagement).
    0. This year's valuation date set for the engagement: from last year's to it.
    1. From last year's valuation date to this year's client model's, when both are known and this year's is
       later: the new valuation date is this year's model's. (Not the move between the two client models: the
       overlay can sit on a copy of a client model of another date.)
    Where this year's model isn't dated after last year's valuation date, that date is likely the model's own,
    not this year's valuation date: nothing is rolled (0 months, "check:" in the basis) and the figures wait for
    this year's valuation date. Guessing a roll from anything else would put numbers on the page that aren't
    this year's.
    2. Else, only when last year's valuation date isn't known, the move between the two client models' dates.
    3. Else the client sheets' timelines: how far each sheet's first period moved (each of last year's sheets
       against the sheet it is this year), the most common move, the smaller on a tie; only forward, at most
       ROLL_MAX months, and from at least ROLL_SHEETS sheets.
    4. Else 12 months, flagged so the page asks for a check."""
    ov_vd, prior_vd, current_vd, this_vd = (list(dates) + [None] * 4)[:4]
    base = ov_vd or prior_vd
    if this_vd and base and this_vd[:10] > base[:10]:
        return (months_between(base[:10], this_vd[:10]),
                f"from last year's valuation date ({base[:10]}) to this year's ({this_vd[:10]}), set for the engagement",
                this_vd[:10])
    if ov_vd and current_vd and current_vd[:10] > ov_vd[:10]:
        return (months_between(ov_vd[:10], current_vd[:10]),
                f"from last year's valuation date ({ov_vd[:10]}) to this year's client model's ({current_vd[:10]})",
                current_vd[:10])
    if ov_vd and current_vd:
        return 0, (f"check: this year's model's date ({current_vd[:10]}) isn't after last year's valuation date "
                   f"({ov_vd[:10]}), so it's likely the model's own date, not this year's valuation date; nothing is "
                   "rolled until this year's valuation date is set"), ov_vd[:10]
    if not ov_vd and prior_vd and current_vd and current_vd[:10] > prior_vd[:10]:
        m = months_between(prior_vd[:10], current_vd[:10])
        return m, (f"from the valuation dates in the two client models ({prior_vd[:10]} to {current_vd[:10]}); last "
                   "year's own valuation date isn't known"), current_vd[:10]
    new = lambda m: to_date(add_months(serial(date.fromisoformat(base[:10])), m)).isoformat() if base else None
    src = sess.prior or sess.ov
    sheets = sess.client_sheets or ((prior or {}).get("sheets") if same_file else None)
    moves = Counter()
    for s in sorted(sheets or {k[1] for k in sess.ext_cached} or []):
        to = sess.rowmap.sheet_for(s) if sess.rowmap else s
        a, b = src.timeline(s), sess.current.timeline(to) if sess.current and to else {}
        if a and b:
            moves[months_between(to_date(min(a.values())).isoformat(), to_date(min(b.values())).isoformat())] += 1
    seen = sum(moves.values())
    forward = sorted(((n, -m) for m, n in moves.items() if 0 < m <= ROLL_MAX), reverse=True)
    if forward and seen >= ROLL_SHEETS:
        n, m = forward[0][0], -forward[0][1]
        return m, (f"from the client sheets' timelines (the first period moved {m} months on {n} of {seen} "
                   "sheet(s))"), new(m)
    return 12, ("assumed: the valuation dates don't give a roll and the timelines don't show one"
                + (f" (moves seen on {seen} sheet(s): {', '.join(f'{m:+d}' for m in sorted(moves))} months)" if moves else "")
                + "; check the roll-forward"), new(12)


ZERO_ROLL = (0.75, 1.33)  # this year's model at last year's date, against last year's figure: about the same
ROLL_PLAN = 4  # the rules' version: a roll planned by older rules is planned again when a session loads
ROLL_MAX = 24  # months: a move beyond it from the timelines isn't a roll-forward
ROLL_SHEETS = 5  # sheets: fewer can't show how far the timelines moved


def plan_roll(sess: Session, prior: dict | None, overlay: dict, same_file: bool, ov_vd: str | None,
              prior_vd: str | None, current_vd: str | None, this_vd: str | None = None) -> dict:
    """The roll-forward's settings (roll_months), and last year's valuation date set on the session."""
    months, basis, new_vd = roll_months(sess, prior, overlay, same_file, (ov_vd, prior_vd, current_vd, this_vd))
    sess.base_vd = serial(date.fromisoformat(ov_vd[:10])) if ov_vd else None
    sess._pshift.clear()
    return {"prior_valuation_date": ov_vd, "months": months, "months_basis": basis,
            "fixed_horizon": bool(sess.current) and sess.fixed_horizon(), "horizon_set": getattr(sess, "horizon_set", None),
            "months_assumed": "assumed:" in basis, "date_check": basis.startswith("check:"),
            "current_valuation_date": new_vd, "plan": ROLL_PLAN,
            "dates": {"overlay": ov_vd, "prior_client": prior_vd, "current_client": current_vd, "this_year": this_vd}}


def _feed(summary: dict, mode: str, valuation_date: str | None, months: int | None) -> tuple[dict, dict | None, int]:
    """A feed's own settings, the base that a person's changes go on top of: rolled forward, the months to roll
    and the new valuation date on its lever. -> (overrides, roll info or None, months)."""
    if mode != "current":
        return {}, None, 0
    roll = summary.get("roll") or {}
    chosen = valuation_date is not None  # a date chosen on the page, rather than the plan's
    valuation_date = valuation_date or roll.get("current_valuation_date")
    pvd = roll.get("prior_valuation_date")
    if months is None:  # the plan's months; a date chosen on the page moves the periods as far as it moves from last year's
        months = months_between(pvd[:10], valuation_date[:10]) if chosen and pvd else roll.get("months", 12)
    vd_cell = roll.get("valuation_date_cell")
    defaults = {parse_a1(vd_cell): serial(date.fromisoformat(valuation_date[:10]))} if valuation_date and vd_cell else {}
    return defaults, {"months": months, "valuation_date": valuation_date, "valuation_date_cell": vd_cell}, months


def scenario(sess: Session, summary: dict, mode: str, changes: dict, valuation_date: str | None = None,
             months: int | None = None) -> dict:
    """Run the overlay with lever / cell changes ({"Sheet!A1": value}) on a feed. Returns outputs against the base
    (same feed, no changes) and the workbook, what the roll-forward changed, and feed cells it couldn't match."""
    outputs = summary["outputs"]
    cells = [parse_a1(o["cell"]) for o in outputs]
    extra = {parse_a1(k): v for k, v in changes.items()}
    defaults, roll_info, months = _feed(summary, mode, valuation_date, months)
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


# ---- the Summary page: the report's summary table, rebuilt and rolled forward --------------------------------

def _range_cells(db, cell: str, texts: list, scale: float, sign: int, sheets: list[str]) -> list[str | None]:
    """The cells holding a conclusion's low and high ends: the report's printed figures at their precision, near
    the preferred value first (same area of the sheet), else anywhere on the overlay's sheets (a formula cell
    before a typed one, then the nearest sheet order)."""
    s, r, c = parse_a1(cell)
    out = []
    for t in texts:
        n = _num_in(t or "")
        if not n:
            out.append(None)
            continue
        x, d = n
        lo, hi = (x - 0.5 * 10 ** -d) * (scale or 1.0) * sign, (x + 0.5 * 10 ** -d) * (scale or 1.0) * sign
        hits = []
        for sh, rr, cc, v, f in db.execute(
                f"SELECT sheet, row, col, value, formula FROM cells WHERE sheet IN ({','.join('?' * len(sheets))}) "
                f"AND typeof(value) IN ('real','integer') AND value BETWEEN ? AND ?", (*sheets, min(lo, hi) - 1e-9, max(lo, hi) + 1e-9)):
            if (sh, rr, cc) == (s, r, c) or abs(round(sign * v / (scale or 1.0), d) - x) > 0.5 * 10 ** -d + 1e-9:
                continue
            near = sh == s and abs(rr - r) <= 6 and abs(cc - c) <= 8
            hits.append((not near, f is None, (abs(rr - r) + abs(cc - c)) if sh == s else 999, sh, rr, cc))
        out.append(_a1(*min(hits)[3:]) if hits else None)
    return out


def timing_rule(wb: Workbook, s: str, r: int, vd: float | None) -> dict | None:
    """How a timing row (period flags, period dates) follows from its sheet's period dates, if one relation holds
    for every period last year: the period's date as the header shows it, its end, its start, the year it ends
    in (plus a constant), 1 after the valuation date (or up to it) and 0 otherwise, or one constant. Then this
    year's value is worked out from the rolled period date instead of hunting for a row: the period grid is a
    property of the model."""
    tl = wb.timeline(s)
    vals = wb.sheet(s)
    pts = [(tl[c], vals.get((r, c))) for c in sorted(tl)]
    pts = [(t, float(v)) for t, v in pts if isinstance(v, (int, float)) and not isinstance(v, bool)]
    if len(pts) < 2:
        return None
    dates = sorted(set(tl.values()))
    gaps = sorted(b - a for a, b in zip(dates, dates[1:]))
    plen = max(1, round(gaps[len(gaps) // 2] / 30.44)) if gaps else 12
    starts = all(to_date(d).day == 1 for d in dates)
    base = {"plen": plen, "starts": starts}
    tries = [("timeline", "the period's date"), ("end", "the period's end"), ("start", "the period's start")]
    for kind, text in tries:
        rule = {**base, "kind": kind, "text": f"{text}, from the period dates"}
        if all(abs(timing_value(rule, t, vd) - v) < 0.5 for t, v in pts):
            return rule
    off = {v - to_date(_period_end(base, t)).year for t, v in pts}
    if len(off) == 1:
        o = off.pop()
        if abs(o) <= 1:
            return {**base, "kind": "year", "offset": o, "text": "the year the period ends in, from the period dates"}
    if vd is not None and {v for _, v in pts} <= {0.0, 1.0}:
        for kind, text in (("after", "1 for periods ending after the valuation date"),
                           ("upto", "1 for periods ending by the valuation date")):
            rule = {**base, "kind": kind, "text": f"{text}, from the period dates"}
            if all(timing_value(rule, t, vd) == v for t, v in pts):
                return rule
    if len({v for _, v in pts}) == 1:
        return {**base, "kind": "const", "value": pts[0][1], "text": f"the same in every period ({pts[0][1]:g})"}
    return None


def _period_end(rule: dict, t: float) -> float:
    return add_months(t, rule["plen"]) - 1 if rule["starts"] else t


def timing_value(rule: dict, t: float, vd: float | None) -> float:
    """A timing rule's value for the period whose header date is t, with valuation date vd."""
    end = _period_end(rule, t)
    k = rule["kind"]
    if k == "timeline":
        return t
    if k == "end":
        return end
    if k == "start":
        return t if rule["starts"] else add_months(t, -rule["plen"]) + 1
    if k == "year":
        return float(to_date(end).year + rule["offset"])
    if k == "after":
        return 1.0 if vd is not None and end > vd else 0.0
    if k == "upto":
        return 1.0 if vd is not None and end <= vd else 0.0
    return rule["value"]


def dcf_origins(sess: Session, summary: dict, cells: list[str]) -> dict:
    """For each figure, the client rows the discountings under it take their cash flows from (dcffacts.py):
    {cell: {"amounts": [(sheet, row)], "timing": [(sheet, row, why)]}}. The amounts are the rows this year's model
    must have for the figure to be this year's; the timing rows (period flags and dates) are what the discounting's
    timing depends on, listed but not asked for. Cached on the session. (Facts.run resets the feed: call it before
    reading a feed whose unmatched values are wanted.)"""
    import dcffacts
    cache = sess.__dict__.setdefault("_origins", {})
    out = {}
    for cell in cells:
        if cell not in cache:
            amounts, timing = set(), set()
            try:
                fx = dcffacts.Facts(sess, summary).run(cell)
                for c in fx["discountings"]:
                    for o in c["origins"]:
                        m = re.match(r"^(?:\[\d+\])?(.+)!r(\d+)$", o["row"])
                        if not m:
                            continue
                        if o.get("kind") == "timing":
                            timing.add((m[1], int(m[2]), o.get("kind_why") or ""))
                        else:
                            amounts.add((m[1], int(m[2])))
            except Exception:
                pass
            cache[cell] = {"amounts": sorted(amounts), "timing": sorted(timing)}
        out[cell] = cache[cell]
    return out


def summary_table(sess: Session, summary: dict, facts: list[dict], changes: dict | None = None,
                  valuation_date: str | None = None, months: int | None = None, method: dict | None = None,
                  schedule: list[dict] | None = None, outside: set | None = None) -> dict:
    """The report's summary, rebuilt and rolled forward. Rows: the report's conclusions (with their ranges), its
    assumptions and its approach. Columns: the report; as saved (the overlay's own Excel values: last year's tie to
    the report is checked here, once); rebuilt (the Python overlay fed from last year's client model, or from the
    values saved in the overlay without one: it shouldn't move from as saved, and feed_moves says where it does);
    this year, rolled forward (the current feed); and a scenario: this year (last year without a current model) with
    the person's assumption changes, valuation date and discounting method. schedule: the overlay's own outputs
    (outputs.py); its conclusions and assumptions that no report row has come after them (source "overlay"), with
    every column, so last year sits beside this year whether or not the report quoted them. They don't enter the
    gate: this year's gate (this_year_gaps) is the same with or without them, and they are held with the table.
    outside: the keys of report figures a person marked as produced outside the model (row["outside"]). The method reaches a figure through its trace (dcftrace.recompute): each discounting under it is
    redone with the new method and the formulas above carry the results up; that is checked first by
    reproducing the module's own value with the method unchanged."""
    import dcf
    import dcftrace
    w = summary["wiring"]
    path = w["overlay"]["db_path"]
    db = rodb.connect(path)
    base_feed = "prior" if w.get("prior") else "workbook"
    this_feed = "current" if w.get("current") else None
    sc_feed = this_feed or base_feed
    extra = {parse_a1(k): v for k, v in (changes or {}).items()}
    outs = {o["fact_id"]: o for o in summary.get("outputs") or [] if o.get("fact_id")}
    levers = {l["key"]: l for l in summary.get("levers") or []}
    rows, keys = [], set()
    for f in facts:
        cat = f.get("category")
        if cat == "conclusion":
            o = outs.get(f["id"])
            row = {"kind": "conclusion", "key": f.get("key"), "label": f.get("label") or f.get("key"),
                   "report": f.get("value_text"), "report_low": f.get("low_text"), "report_high": f.get("high_text"),
                   "unit": f.get("unit"), "page": f.get("page"), "cell": o["cell"] if o else None}
            if o:
                row.update(scale=o.get("scale") or 1.0, sign=o.get("sign") or 1)
                row["low_cell"], row["high_cell"] = _range_cells(db, o["cell"], [f.get("low_text"), f.get("high_text")],
                                                                 row["scale"], row["sign"], summary["sheets"]) \
                    if f.get("low_text") else (None, None)
                keys.update(parse_a1(x) for x in (o["cell"], row["low_cell"], row["high_cell"]) if x)
            rows.append(row)
        elif cat == "assumption":
            l = levers.get(f.get("key"))
            rows.append({"kind": "assumption", "key": f.get("key"), "label": f.get("label") or f.get("key"),
                         "report": f.get("value_text"), "basis": f.get("basis"), "unit": f.get("unit"),
                         "page": f.get("page"), "lever": l})
            if l:
                keys.add(parse_a1(l["cell"]))
        elif cat == "approach":
            rows.append({"kind": "approach", "key": f.get("key"), "label": f.get("label") or f.get("key"),
                         "report": f.get("value_text"), "page": f.get("page")})
    for r in rows:
        if r["kind"] in ("conclusion", "assumption") and not r.get("cell") and not r.get("lever") and r.get("key") in (outside or ()):
            r["outside"] = True
    used = {r["cell"] for r in rows if r.get("cell")} | {r["lever"]["cell"] for r in rows if r.get("lever")}
    for o in schedule or []:  # the overlay's own figures the report doesn't quote
        if o.get("kind") != "figure" or o.get("class") not in ("conclusion", "assumption") or not o.get("cell") \
                or o["cell"] in used:
            continue
        used.add(o["cell"])
        rows.append({"kind": o["class"], "key": None, "label": o["label"], "report": None, "unit": None, "page": None,
                     "cell": o["cell"], "scale": 1.0, "sign": 1, "source": "overlay"})
        keys.add(parse_a1(o["cell"]))
    keys = sorted(keys)
    ours = lambda r: r["kind"] == "conclusion" and r.get("cell") and r.get("source") != "overlay"  # the gate's figures
    by_fig = dcf_origins(sess, summary, [r["cell"] for r in rows if ours(r)]) \
        if w.get("current") and sess.rowmap else {}
    origins = sorted({k for x in by_fig.values() for k in x["amounts"]})
    timing = sorted({k for x in by_fig.values() for k in x["timing"]} - {(s_, r_, "") for s_, r_ in origins})
    # a timing row (period flags, period dates) not found well this year is worked out from the period dates,
    # where one relation held for every period last year; else it's a row like any other
    sess.derived = {}
    if w.get("current") and sess.rowmap:
        src = sess.prior or sess.ov
        for s_, r_, _w in timing:
            if not sess.rowmap.confident(s_, r_):
                rule = timing_rule(src, s_, r_, sess.base_vd)
                if rule:
                    sess.derived[(s_, r_)] = rule

    def read(feed, with_changes, vd=None, mo=None):
        defaults, roll, mo = _feed(summary, feed, vd, mo)
        sess.configure(feed, {**defaults, **(extra if with_changes else {})}, mo or 0)
        got = dict(zip(keys, sess.values(keys)))
        return got, roll, defaults, mo

    cols = {"saved": {k: sess.ov.value(*k) for k in keys}}  # Excel's own values: where last year's tie is checked
    cols["rebuilt"], _, _, _ = read(base_feed, False)
    roll = None
    gaps = None
    if this_feed:
        cols["this_year"], roll, _, _ = read(this_feed, False)
        # client values this year's model has no match for: last year's stand in; the rows, for the page
        by_row = Counter((s_, r_) for (s_, r_, _c) in sess.unmatched)
        labels = (sess.prior or sess.ov).labels()
        reads = len(sess.client_reads)
        share = 1 - len(sess.unmatched) / reads if reads else 1.0
        family = sess.rowmap.family() if sess.rowmap else None
        # the rows the discountings' cash flows come from decide: each found this year (or picked), with its
        # periods; without discountings found, at least half of all the client values the figures read
        read_by_row = Counter((s_, r_) for (s_, r_, _c) in sess.client_reads)
        # last year's values kept on purpose, by a person, or a row with nothing to find (code): the agents may keep
        # them for other rows, not for the rows the DCF's cash flows come from (keeping every hard row would pass
        # every check with last year's numbers)
        kept = lambda k: sess.rowmap.explain(*k).get("stand_in") and sess.rowmap.explain(*k).get("by", "you") in ("you", "code")
        missing = [k for k in origins if not kept(k) and (sess.rowmap.locate(*k) is None
                   or by_row.get(k, 0) > 0.5 * max(1, read_by_row.get(k, 0)))]
        # rows found this year but blank in more than a fifth of the periods read, where last year's had values:
        # the wrong row (an actuals sheet matched on its history), or the line item moved. They stand in and block
        blank_by_row, blank_at = Counter(), {}
        for (s_, r_, _c), at in sess.blank.items():
            blank_by_row[(s_, r_)] += 1
            blank_at.setdefault((s_, r_), at)
        blank_rows = [k for k, n in blank_by_row.most_common() if n > 0.2 * max(1, read_by_row.get(k, 0)) and k not in missing]
        # rows the figures read that nobody has checked and that weren't found well: not found, or matched with
        # little confidence (79% of them on a rebuilt model, some to pasted copies). A person looks at each: picks
        # the row, keeps the one found, or keeps last year's values
        timing_rows = {(s_, r_) for s_, r_, _w in timing}
        weak_rows = [k for k in sorted(read_by_row) if k not in missing and k not in blank_rows
                     and k not in timing_rows and not sess.rowmap.confident(*k)]
        # the timing the discounting depends on counts too: found well, or worked out from the period dates. The
        # agents keeping last year's timing doesn't settle it: last year's flags would discount at last year's dates
        agents_kept = lambda k: sess.rowmap.explain(*k).get("stand_in") and sess.rowmap.explain(*k).get("by") == "agent"
        timing_open = [k for k in sorted(timing_rows) if k in read_by_row and k not in sess.derived
                       and (not sess.rowmap.confident(*k) or agents_kept(k))]

        def why_missing(k):
            if sess.rowmap.explain(*k).get("stand_in"):
                return "the agents couldn't find it this year (last year's values stand in)"
            if sess.rowmap.locate(*k) is None:
                return "not found this year"
            if blank_by_row.get(k, 0) > 0.5 * by_row.get(k, 0):
                return f"found at {blank_at[k][0]}!r{blank_at[k][1]}, but blank there in most of its periods"
            return "found, but most of its periods aren't in this year's model"

        # each figure on its own basis: the rows its own discountings' cash flows come from, where it has any;
        # else the share of client values found (a figure whose discounting isn't recognised can't lean on
        # another figure's rows). A row found but blank this year holds every figure back.
        # with no valuation date for this year, nothing is rolled and no figure is this year's
        date_hold = bool((summary.get("roll") or {}).get("date_check"))
        by_cell = {}
        for row in rows:
            if not ours(row):
                continue
            mine = (by_fig.get(row["cell"]) or {}).get("amounts") or []
            gone = [k for k in mine if k in missing]
            by_cell[row["cell"]] = {"basis": "dcf" if mine else "share", "dcf_rows": len(mine),
                                    "missing": [f"{s_}!r{r_}" for s_, r_ in gone],
                                    "reliable": (not gone if mine else share >= 0.5) and not blank_rows and not weak_rows
                                    and not timing_open and not date_hold}
            for part in ("low_cell", "high_cell"):
                if row.get(part):
                    by_cell[row[part]] = by_cell[row["cell"]]
        gaps = {"values": len(sess.unmatched), "stood_in": len(sess.stood_in), "reads": reads,
                "found_share": round(share, 3), "family": family, "basis": "dcf" if origins else "share",
                "dcf_rows": len(origins), "dcf_origins": [f"{s_}!r{r_}" for s_, r_ in origins],
                "read_rows": [f"{s_}!r{r_}" for s_, r_ in sorted(read_by_row)],
                "dcf_missing": [{"row": f"{s_}!r{r_}", "label": labels.get((s_, r_), ""), "why": why_missing((s_, r_))}
                                for s_, r_ in missing],
                "blank_rows": [{"row": f"{s_}!r{r_}", "label": labels.get((s_, r_), ""),
                                "found": f"{blank_at[(s_, r_)][0]}!r{blank_at[(s_, r_)][1]}",
                                "blank": blank_by_row[(s_, r_)], "of": read_by_row.get((s_, r_), 0)} for s_, r_ in blank_rows],
                "weak_rows": [dict(zip(("row", "label", "found", "confidence", "how"), (
                    f"{s_}!r{r_}", labels.get((s_, r_), ""),
                    (lambda ex: f"{ex['found'][0]}!r{ex['found'][1]}" if ex["found"] else None)(sess.rowmap.explain(s_, r_)),
                    sess.rowmap.explain(s_, r_)["confidence"], sess.rowmap.explain(s_, r_)["how"]))) for s_, r_ in weak_rows],
                "timing": [{"row": f"{s_}!r{r_}", "label": labels.get((s_, r_), ""), "why": why,
                            "found": (lambda h: f"{h[0]}!r{h[1]}" if h else None)(sess.rowmap.locate(s_, r_)),
                            "derived": (sess.derived.get((s_, r_)) or {}).get("text"),
                            "open": (s_, r_) in timing_open}
                           for s_, r_, why in timing],
                "by_cell": by_cell, "date_check": date_hold,
                "reliable": all(x["reliable"] for x in by_cell.values()) if by_cell
                else (not missing if origins else share >= 0.5) and not blank_rows and not weak_rows and not timing_open
                and not date_hold,
                "rebuilt": family is not None and family < 0.5,
                "rows": [{"row": f"{s_}!r{r_}", "label": labels.get((s_, r_), ""), "values": n,
                          "why": sess.unmatched[next(k for k in sess.unmatched if k[:2] == (s_, r_))]}
                         for (s_, r_), n in by_row.most_common(30)]}
    if this_feed and gaps is not None:
        # the zero-roll check: this year's model at last year's valuation date, rolled by nothing, should give
        # about last year's figures (forecasts are revised, not replaced). Far off, the rows found don't carry
        # the same numbers, whatever the other tests say; rows standing in pass it trivially
        pvd = (summary.get("roll") or {}).get("prior_valuation_date")
        zero = cols["this_year"] if date_hold else (read(this_feed, False, pvd, 0)[0] if pvd else None)
        for row in rows:
            if row["kind"] != "conclusion" or not row.get("cell") or zero is None:
                continue
            k = parse_a1(row["cell"])
            z, last = zero.get(k), cols["rebuilt"].get(k)
            ratio = z / last if isinstance(z, float) and isinstance(last, float) and last else None
            ok = ratio is not None and ZERO_ROLL[0] <= ratio <= ZERO_ROLL[1]
            row["zero_roll"] = {"value": _show(z), "ratio": None if ratio is None else round(ratio, 4), "ok": ok,
                                "valuation_date": pvd}
            c = gaps["by_cell"].get(row["cell"])
            if c is not None:
                c["zero_roll"] = row["zero_roll"]
                c["reliable"] = c["reliable"] and ok
        gaps["zero_roll_off"] = [{"cell": r["cell"], "label": r["label"], **r["zero_roll"]} for r in rows
                                 if r.get("zero_roll") and not r["zero_roll"]["ok"] and ours(r)]
        if gaps["by_cell"]:
            gaps["reliable"] = all(x["reliable"] for x in gaps["by_cell"].values())
    cols["scenario"], sc_roll, sc_defaults, sc_months = read(sc_feed, True, valuation_date, months)
    val = lambda col, cell: _show(cols[col].get(parse_a1(cell))) if cell else None
    for row in rows:
        if row.get("source") == "overlay":
            row["values"] = {col: val(col, row["cell"]) for col in cols}
            a, b = cols["saved"].get(parse_a1(row["cell"])), cols["rebuilt"].get(parse_a1(row["cell"]))
            row["feed_moves"] = base_feed == "prior" and isinstance(a, float) and isinstance(b, float) and not dcf._close(a, b)
        elif row["kind"] == "conclusion" and row.get("cell"):
            for part, cell in (("values", row["cell"]), ("low", row.get("low_cell")), ("high", row.get("high_cell"))):
                if not cell:
                    continue
                typed = part != "values" and db.execute("SELECT formula FROM cells WHERE sheet=? AND row=? AND col=?",
                                                        parse_a1(cell)).fetchone()[0] is None
                row[part] = {col: val(col, cell) for col in cols if not typed or col in ("saved", "rebuilt")}
                if typed:
                    row["range_note"] = (f"the range ends ({row.get('low_cell')}, {row.get('high_cell')}) are typed into the "
                                         f"overlay, not calculated, so they don't move with this year's model")
            if row.get("values"):
                k = parse_a1(row["cell"])
                row["tie"] = tie(cols["saved"].get(k), row["report"], row["scale"], row["sign"])  # the overlay as saved
                a, b = cols["saved"].get(k), cols["rebuilt"].get(k)
                # fed from last year's client model it moves: that file isn't the version the overlay was built on
                row["feed_moves"] = base_feed == "prior" and isinstance(a, float) and isinstance(b, float) and not dcf._close(a, b)
        elif row["kind"] == "assumption" and row.get("lever"):
            row["values"] = {col: val(col, row["lever"]["cell"]) for col in cols}

    # the discounting method, through each conclusion's trace
    detected, notes = [], []
    method = {k: v for k, v in (method or {}).items() if v}
    traces = {}
    for row in rows:
        if row["kind"] != "conclusion" or not row.get("cell"):
            continue
        try:
            t = dcftrace.trace(db, row["cell"])
        except ValueError:
            continue
        cs = [c for c in dcftrace.cores(t) if c.get("inputs")]
        traces[row["cell"]] = (t, cs)
        for c in dcftrace.cores(t):
            m = c.get("method") or {}
            d = f"{c['what']}" + (f", {m['timing']} of period, {m['day_count']}" if m else "")
            if d not in detected:
                detected.append(d)
    rate_check = _rate_check(sess, db, rows, traces)
    if method and traces:
        need = set(keys)
        for t, cs in traces.values():
            for c in cs:
                need |= _dcf_cells(db, c["inputs"])

            def walk(n):
                if not n.get("again"):
                    r = dcf._ref(n["cell"], "")
                    need.add((r[0], r[1], r[2]))
                for ch in n.get("children", []):
                    walk(ch)
            walk(t)
        sess.configure(sc_feed, {**sc_defaults, **extra}, sc_months or 0)
        live_db = rodb.patched(path, _module_values(db, sess, need))
        for row in rows:
            if row.get("cell") not in traces:
                continue
            t, cs = traces[row["cell"]]
            if not cs:
                row["method_note"] = "no discounting under it that the method can be applied to"
                continue
            try:
                same = {c["cell"]: dcf.compute(live_db, **{**c["inputs"], "compare_to": None}, fix=False)["total"] for c in cs}
                redone = {c["cell"]: dcf.compute(live_db, **{**c["inputs"], "compare_to": None, **method}, fix=False)["total"]
                          for c in cs}
            except ValueError as e:
                row["method_note"] = f"couldn't redo its discounting: {e}"
                continue
            module = cols["scenario"].get(parse_a1(row["cell"]))
            again = dcftrace.recompute(live_db, t, same)
            if again is None or not isinstance(module, float) or not dcf._close(again, module):
                row["method_note"] = "the formulas above its discounting can't be recomputed here, so the method isn't applied"
                continue
            row.setdefault("values", {})["scenario"] = dcftrace.recompute(live_db, t, redone)
            row["method_applied"] = True
    sess.configure("workbook")
    return {"rows": rows, "columns": [c for c in ("saved", "rebuilt", "this_year", "scenario") if c in cols],
            "feeds": {"rebuilt": base_feed, "this_year": this_feed, "scenario": sc_feed},
            "roll": roll, "scenario_roll": sc_roll, "detected_method": detected, "method": method,
            "changes": changes or {}, "notes": notes, "this_year_gaps": gaps, "rate_check": rate_check}


def _rate_check(sess: Session, db, rows: list[dict], traces: dict) -> dict | None:
    """The discount rate three ways, where they don't agree: the report's, the lever's (the overlay cell labelled
    like it) and the rate cells the discountings under the figures use (read back from their factors). None when
    changing the lever moves one of those cells. Rates on different bases can rightly differ (a WACC, and a
    cost of equity for equity cash flows); a lever that moves none of them moves nothing, so a scenario on it
    means nothing. Said, not changed: a person decides which cell is the rate."""
    import dcf
    used = {}
    for _, cs in traces.values():
        for c in cs:
            ref = c["inputs"].get("rate")
            if isinstance(ref, str) and "!" in ref:
                used.setdefault(ref, c.get("rate_value"))
    if not used:
        return None
    dr = next((r for r in rows if r["kind"] == "assumption" and r.get("key") == "discount_rate"), None)
    lever = (dr or {}).get("lever")
    labels = sess.ov.labels()
    lab = lambda ref: labels.get(parse_a1(ref)[:2], "")
    moves = False
    if lever:
        try:
            lk = parse_a1(lever["cell"])
            v0 = sess.ov.value(*lk)
            if isinstance(v0, float):
                sess.configure("workbook", {lk: v0 + 0.01})
                moved = sess.values([parse_a1(ref) for ref in used])
                moves = any(isinstance(v, float) and isinstance(used[ref], float) and not dcf._close(v, used[ref])
                            for ref, v in zip(used, moved))  # a low and a high rate: the lever is one of them
        except (ValueError, KeyError):
            moves = False
        finally:
            sess.configure("workbook")
    if moves:
        return None
    return {"report": {"value": dr.get("report"), "basis": dr.get("basis")} if dr else None,
            "lever": {"cell": lever["cell"], "label": lever.get("label") or lab(lever["cell"]),
                      "value": sess.ov.value(*parse_a1(lever["cell"]))} if lever else None,
            "dcf": [{"cell": ref, "label": lab(ref), "value": v} for ref, v in used.items()]}


def value_bridge(sess: Session, summary: dict, facts: list[dict], changes: dict | None = None,
                 valuation_date: str | None = None, months: int | None = None, method: dict | None = None,
                 schedule: list[dict] | None = None) -> dict:
    """Last year's value to this year's, step by step, for each of the report's conclusions in the overlay:
      last year           the rebuild on last year's model (it ties to the report)
      time value          each discounting under it grows at its own rate from last year's valuation date to this
                          year's (dcftrace.recompute carries that up to the figure)
      cash flows          each discounting redone at the new valuation date, so last year's forecast cash flows up
                          to it drop out
      this year's forecast  the engine's roll-forward onto this year's client model (new cash flows, a longer
                          horizon), less the step before: last year's forecast at the new date
      your changes        the Summary page's scenario, if it has one
    Sequential, so the steps add up exactly. Where the formulas above a figure's discountings can't be recomputed
    here, the first three steps are shown as one."""
    import dcf
    import dcftrace
    w = summary["wiring"]
    if not w.get("current"):
        return {"bridges": [], "why": "assign this year's client model (Roles) and rebuild in Python"}
    table = summary_table(sess, summary, facts, changes, valuation_date, months, method, schedule)
    gaps = table.get("this_year_gaps") or {}
    held = {c for c, x in (gaps.get("by_cell") or {}).items() if not x["reliable"]}
    conc = [r for r in table["rows"] if r["kind"] == "conclusion" and r.get("cell") and r.get("values")]
    # the report's figures; the overlay's own only where none of the report's was found in it (no report)
    conc = [r for r in conc if r.get("source") != "overlay"] or conc
    figures = [r["cell"] for r in conc]
    if gaps and not gaps.get("reliable", True) and (not gaps.get("by_cell") or all(c in held for c in figures)):
        what = (f"{len(gaps['dcf_missing'])} of the {gaps['dcf_rows']} rows the DCF's cash flows come from weren't found in "
                "this year's model" if gaps.get("basis") == "dcf" else
                f"only {gaps['found_share']:.0%} of the client values the report's figures read were found in this year's model")
        return {"bridges": [], "this_year_gaps": gaps,
                "why": f"{what}, so there's no this-year value to bridge to yet: find or pick them in Rebuild in Python → "
                       "Valuation (DCF) → The facts"}
    path = w["overlay"]["db_path"]
    db = rodb.connect(path)
    base_feed = table["feeds"]["rebuilt"]
    vd1 = (table.get("roll") or {}).get("valuation_date")
    rows = [r for r in conc if r["cell"] not in held]
    withheld = [{"label": r["label"], "cell": r["cell"]} for r in table["rows"] if r.get("cell") in held]
    traced, need = {}, set()
    for r in rows:
        try:
            t = dcftrace.trace(db, r["cell"])
        except ValueError:
            continue
        cs = [c for c in dcftrace.cores(t) if c.get("inputs")]
        traced[r["cell"]] = (t, cs)
        for c in cs:
            need |= _dcf_cells(db, c["inputs"])

        def walk(n):
            if not n.get("again"):
                x = dcf._ref(n["cell"], "")
                need.add((x[0], x[1], x[2]))
            for ch in n.get("children", []):
                walk(ch)
        walk(t)
    sess.configure(base_feed)
    prior_db = rodb.patched(path, _module_values(db, sess, need)) if need else db
    sess.configure("workbook")
    out, errors = [], []
    for r in rows:
        v = r["values"]
        v0, v2, v3 = v.get("rebuilt"), v.get("this_year"), v.get("scenario")
        if not all(isinstance(x, float) for x in (v0, v2)):
            errors.append(f"{r['label']} ({v0 if not isinstance(v0, float) else v2})")
            continue
        steps = [{"key": "last_year", "label": "Last year (rebuilt; ties to the report)" if (r.get("tie") or {}).get("ok")
                  else "Last year (rebuilt)", "value": v0, "total": True}]
        note, split = None, False
        t_cs = traced.get(r["cell"])
        if t_cs and t_cs[1] and vd1:
            t, cs = t_cs
            try:
                base = {c["cell"]: dcf.compute(prior_db, **{**c["inputs"], "compare_to": None}, fix=False) for c in cs}
                again = dcftrace.recompute(prior_db, t, {k: x["total"] for k, x in base.items()})
                if again is None or not dcf._close(again, v0):
                    note = "the formulas above its discountings can't be recomputed here, so the roll-forward is one step"
                else:
                    vd1d = date.fromisoformat(vd1[:10])
                    grown = {k: x["total"] * (1 + x["rate"]) ** dcf.yearfrac(x["valuation_date"], vd1d, x["day_count"])
                             for k, x in base.items()}
                    rolled = {c["cell"]: dcf.compute(prior_db, **{**c["inputs"], "compare_to": None, "valuation_date": vd1},
                                                     fix=False)["total"] for c in cs}
                    vu, v1 = dcftrace.recompute(prior_db, t, grown), dcftrace.recompute(prior_db, t, rolled)
                    if vu is not None and v1 is not None:
                        years = dcf.yearfrac(next(iter(base.values()))["valuation_date"], vd1d, "actual/actual")
                        steps += [{"key": "time", "label": f"Time value: {years:.2f} years of unwind at the discount rate",
                                   "value": vu - v0},
                                  {"key": "cash", "label": f"Last year's forecast cash flows up to {vd1[:10]}", "value": v1 - vu},
                                  {"key": "forecast", "label": "This year's client model (new forecast, rolled forward)",
                                   "value": v2 - v1}]
                        split = True
            except (ValueError, ZeroDivisionError, OverflowError) as e:
                note = f"couldn't split the roll-forward: {e}"
        elif not t_cs or not t_cs[1]:
            note = "no discounting found under it, so the roll-forward is one step"
        if not split:
            steps.append({"key": "roll", "label": "Roll-forward onto this year's model (time, cash flows and forecast)",
                          "value": v2 - v0})
        steps.append({"key": "this_year", "label": f"This year, rolled forward{f' to {vd1[:10]}' if vd1 else ''}",
                      "value": v2, "total": True})
        if isinstance(v3, float) and not dcf._close(v3, v2):
            steps.append({"key": "changes", "label": "Your changes (the Summary page's scenario)", "value": v3 - v2})
            steps.append({"key": "scenario", "label": "Your scenario", "value": v3, "total": True})
        out.append({"label": r["label"], "cell": r["cell"], "report": r.get("report"), "scale": r.get("scale") or 1.0,
                    "sign": r.get("sign") or 1, "unit": r.get("unit"), "steps": steps, "note": note})
    why = None
    if not out:
        why = ("the report's conclusions come out as Excel errors in the Python overlay: " + "; ".join(errors[:4])
               + ". Rebuild in Python lists the formulas and names it couldn't compute.") if errors else \
              "no conclusion of the report is matched to a cell in the overlay (check the Map)"
    return {"bridges": out, "valuation_date": vd1, "feeds": table["feeds"], "why": why, "errors": errors,
            "this_year_gaps": gaps, "withheld": withheld}


# ---- the DCF on the live module ----------------------------------------------------------------------------
# The Model Desk's Valuation tab (valuation.py) on the Python overlay's own numbers. valuation.py finds each DCF's
# structure in the workbook (cash-flow rows, discount rate, valuation date, convention, cut-off, bridge); dcf.py
# then recomputes it from the cash flows, dates and bridge amounts the module computes on the chosen feed, with
# the person's changes. The module and dcf.py are separate calculations, so their agreeing on the anchor value is
# a check of both. dcf.py reads a model.db, so the module's values reach it through rodb.patched().

def _dcf_cells(db, inputs: dict) -> set[tuple]:
    """Every cell a DCF's inputs read: cash-flow rows, their period-end dates, the rate, valuation date and
    cut-off cells, bridge cells or rows, and the anchor cell itself."""
    import dcf
    cells: set[tuple] = set()

    def one(ref):
        if ref is None or isinstance(ref, (int, float)) or dcf._as_date(str(ref)):
            return
        t = str(ref).strip()
        named = db.execute("SELECT ref FROM names WHERE lower(name)=lower(?)", (t,)).fetchone()
        t = named[0] if named else t
        m = dcf._REF.match(t) or (dcf._REF.match(t.replace(" ", "")) if "'" not in t else None)
        if m and (not m["c2"] or (m["c2"], m["r2"]) == (m["c1"], m["r1"])):
            cells.add((m["sheet"].strip("'"), int(m["r1"]), dcf._col(m["c1"])))

    def row(ref):
        sheet, r, cols = dcf._row_range(db, ref)
        cells.update((sheet, r, c) for c in cols)
        _, src = dcf.period_ends(db, sheet, cols)
        m = re.match(r"^(.+?)!r(\d+)\b", src)
        if m:
            cells.update((m[1], int(m[2]), c) for c in cols)
        return sheet, cols

    for ref in inputs["cashflow"]:
        row(ref)
    if inputs.get("dates"):
        row(inputs["dates"])
    for k in ("rate", "valuation_date", "terminal_date", "compare_to"):
        one(inputs.get(k))
    for a in inputs.get("adjustments") or []:
        v = a.get("value") if isinstance(a, dict) else a
        if isinstance(a, dict) and a.get("at_valuation_date"):
            try:
                row(str(v))
                continue
            except ValueError:
                pass
        one(v)
    return cells


def _module_values(db, sess: Session, cells: set[tuple]) -> dict:
    """The module's value for each cell, stored the way model.db stores it (dates as ISO text) so dcf.py reads it."""
    keys = sorted(cells)
    got = sess.values(keys)
    out = {}
    for k, v in zip(keys, got):
        saved = db.execute("SELECT value FROM main.cells WHERE sheet=? AND row=? AND col=?", k).fetchone()
        saved = saved[0] if saved else None
        if isinstance(v, xlruntime.XLError):
            v = v.code
        elif isinstance(v, float) and isinstance(saved, str) and xlruntime._DATE.match(saved) and 0 < v < 2958466:
            v = to_date(v).isoformat()
        out[k] = v
    return out


def _flows(db, inputs: dict) -> tuple[dict, dict]:
    """(cash flow by column, summed over the cash-flow rows; period end date by column) as db holds them."""
    import dcf
    flows, sheet0, cols0 = {}, None, None
    for ref in inputs["cashflow"]:
        sheet, r, cols = dcf._row_range(db, ref)
        sheet0, cols0 = sheet0 or sheet, cols0 or cols
        for c, v in db.execute("SELECT col, value FROM cells WHERE sheet=? AND row=? AND col BETWEEN ? AND ?",
                               (sheet, r, cols[0], cols[-1])):
            if dcf._num(v) is not None:
                flows[c] = flows.get(c, 0.0) + dcf._num(v)
    return flows, dcf.period_ends(db, sheet0, cols0, inputs.get("dates"))[0]


def _brief(r: dict) -> dict:
    iso = lambda d: d.isoformat() if d else None
    return {"rate": r["rate"], "pv": r["pv"], "total": r["total"], "anchor": r["compare_to"],
            "valuation_date": iso(r["valuation_date"]), "timing": r["timing"], "day_count": r["day_count"],
            "terminal_date": iso(r["terminal_date"]), "bridge": r["bridge"], "periods": r["periods"],
            "first_period": iso(r["first_period"]), "last_period": iso(r["last_period"]), "undiscounted": r["undiscounted"]}


def trace_starts(summary: dict) -> list[dict]:
    """Where to start tracing the valuation: the cells the report's conclusions were matched to, those that tie
    to the report at its printed precision first, then label matches with no report figure to tie to. A match
    that differs from the report is never used (it's likely the wrong cell)."""
    outs = [o for o in summary.get("outputs") or [] if o.get("fact_id")]
    tied = [o for o in outs if (o.get("tie") or {}).get("ok")]
    rest = [o for o in outs if not o.get("tie") and re.search(r"value|valuation|\bnpv\b", o.get("label") or "", re.I)]
    out, seen = [], set()
    for o in tied + rest:
        if o["cell"] not in seen:
            seen.add(o["cell"])
            out.append({"cell": o["cell"], "label": o.get("label"), "report": o.get("report"), "key": o.get("key"),
                        "ties": bool((o.get("tie") or {}).get("ok")), "value": o.get("value")})
    return out


def dcf_anchors(summary: dict) -> list[dict]:
    """The DCFs valuation.py finds on the overlay sheets (all of the workbook's if none are there), then the
    discountings traced down from the report's conclusions (dcftrace.py) that those don't already cover."""
    import dcftrace
    import valuation
    path = summary["wiring"]["overlay"]["db_path"]
    cat = valuation.catalogue(path)
    mine = [a for a in cat if a["cell"].split("!")[0].strip("'") in set(summary["sheets"])] or cat
    have = {a["cell"] for a in mine} | {a.get("pv_cell") for a in mine}
    try:
        traced = dcftrace.anchors(path, [s["cell"] for s in trace_starts(summary)])
    except Exception:  # the trace is extra: the label-based ones still stand
        traced = []
    return mine + [a for a in traced if a["cell"] not in have]


FEED_WORDS = {"workbook": "the values saved in the overlay", "prior": "the prior client model",
              "current": "the current client model, rolled forward"}


def dcf_live(sess: Session, summary: dict, mode: str = "workbook", changes: dict | None = None,
             valuation_date: str | None = None, months: int | None = None, cell: str | None = None,
             rate: float | None = None, timing: str | None = None, day_count: str | None = None,
             cutoff: str | None = "model", include: list[bool] | None = None, low: float | None = None,
             high: float | None = None, dcf_valuation_date: str | None = None) -> dict:
    """One of the overlay's DCFs three ways: from the values Excel saved, from the module on a feed with the
    person's changes (checked against the module's own anchor cell), and that again under another discounting
    method (rate, valuation date, end / mid period, day count, cut-off, bridge items), with low / high rates.
    cutoff: "model", "" (none) or YYYY-MM-DD. include: which bridge items to keep."""
    import dcf
    import valuation
    cat = dcf_anchors(summary)
    pick = valuation._pick(cat, cell)
    if not pick:
        return {"anchors": valuation._listing(cat), "selected": None}
    path = summary["wiring"]["overlay"]["db_path"]
    inputs = pick["inputs"]
    db0 = rodb.connect(path)
    cells = _dcf_cells(db0, inputs)
    defaults, roll, months = _feed(summary, mode, valuation_date, months)
    extra = {parse_a1(k): v for k, v in (changes or {}).items()}
    sess.configure(mode, {**defaults, **extra}, months)
    live_vals = _module_values(db0, sess, cells)
    unmatched = len(sess.unmatched)
    sess.configure("workbook")
    live_db = rodb.patched(path, live_vals)
    saved = dcf.compute(db0, **inputs, fix=False)
    live = dcf.compute(live_db, **inputs, fix=False)

    s_in = dict(inputs)
    if rate is not None and rate == rate:
        s_in["rate"] = rate
    if dcf_valuation_date:
        s_in["valuation_date"] = dcf_valuation_date
    if timing in dcf.TIMINGS:
        s_in["timing"] = timing
    if day_count in dcf.DAY_COUNTS:
        s_in["day_count"] = day_count
    if cutoff != "model":
        s_in["terminal_date"] = cutoff or None
    keep = list(include or []) + [True] * len(inputs["adjustments"])
    s_in["adjustments"] = [a for a, k in zip(inputs["adjustments"], keep) if k]
    lo = live["rate"] + 0.01 if low is None or low != low else low
    hi = max(live["rate"] - 0.01, 0.0) if high is None or high != high else high
    sc = dcf.compute(live_db, **{**s_in, "compare_to": None}, rates=[lo, hi], fix=False)
    words = {"end": "end of period", "mid": "mid-period", "actual/actual": "actual/actual (YEARFRAC)",
             "actual/365": "actual/365 (XNPV)"}
    method = []
    if not dcf._close(sc["rate"], live["rate"]):
        method.append(f"discount rate {live['rate']:.2%} → {sc['rate']:.2%}")
    if sc["valuation_date"] != live["valuation_date"]:
        method.append(f"valuation date {live['valuation_date']} → {sc['valuation_date']}")
    if sc["timing"] != live["timing"]:
        method.append(f"discounted at {words[live['timing']]} → {words[sc['timing']]}")
    if sc["day_count"] != live["day_count"]:
        method.append(f"time measured by {words[live['day_count']]} → {words[sc['day_count']]}")
    if sc["terminal_date"] != live["terminal_date"]:
        method.append(f"cut-off {live['terminal_date'] or 'none'} → {sc['terminal_date'] or 'none'}")
    dropped = [b["label"] for b, k in zip(live["bridge"], keep) if not k]
    if dropped:
        method.append("bridge without " + ", ".join(dropped))
    adj = sum(b["value"] for b in sc["bridge"])

    label = pick["label"] or pick["cell"]
    ok_saved = saved["compare_to"] is not None and dcf._close(saved["total"], saved["compare_to"])
    ok_live = live["compare_to"] is not None and dcf._close(live["total"], live["compare_to"])
    checks = [
        {"step": "Workbook", "what": f"{label} recomputed from the values Excel saved", "ours": saved["total"],
         "theirs": saved["compare_to"], "where": pick["cell"], "ok": ok_saved},
        {"step": "Python overlay", "what": f"the module's {label} on {FEED_WORDS[mode]}"
                                           + (f", with {len(extra)} change(s)" if extra else "")
                                           + ", against the DCF recomputed from the module's own cash flows",
         "ours": live["total"], "theirs": live["compare_to"], "where": pick["cell"], "ok": ok_live}]
    f0, e0 = _flows(db0, inputs)
    f1, e1 = _flows(live_db, inputs)
    runs = [("Cumulative PV (as saved in Excel)", saved, f0, e0)]
    if mode != "workbook" or extra:
        runs.append(("Cumulative PV (Python overlay)", live, f1, e1))
    if method:
        runs.append(("Cumulative PV (scenario method)", sc, f1, e1))
    chart = valuation._chart(db0, path, inputs["cashflow"][0], runs, cumulative=True, dates=inputs.get("dates"))
    if mode == "current":
        chart["note"] = (chart.get("note") or "") + (f" Periods are labelled with last year's dates; rolled forward "
                                                     f"they each move on {months} months.")
    return {"anchors": valuation._listing(cat), "selected": pick["cell"], "label": label,
            "feed": {"mode": mode, "words": FEED_WORDS[mode], "roll": roll, "changes": changes or {}, "unmatched": unmatched},
            "workbook": _brief(saved), "live": _brief(live), "scenario": _brief(sc), "method_changes": method,
            "checks": checks, "agrees": ok_live, "units": sc.get("units") or [],
            "sensitivity": [{"case": "low", "rate": lo, "total": sc["sensitivity"][0]["total"], "pv": sc["sensitivity"][0]["total"] - adj},
                            {"case": "scenario", "rate": sc["rate"], "total": sc["total"], "pv": sc["pv"]},
                            {"case": "high", "rate": hi, "total": sc["sensitivity"][1]["total"], "pv": sc["sensitivity"][1]["total"] - adj}],
            "chart": chart}


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
