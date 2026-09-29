"""The facts behind a report figure: the discounting under it, found by following what the Python overlay reads.

Last year's figure ties (the rebuild reproduces it), so the overlay's values are the source of truth. From the
figure's cell, walk() follows the cells each formula actually reads (the runtime records them: Book.reads), so
lookups (INDEX, OFFSET, CHOOSE), names and the branch an IF took are followed as Excel took them, across sheets.
At each formula cell on the way it looks for a discounting, by its values rather than its function's name:
  cash flows x factors    it reads two rows of the same length and equals the sum of their products, one of
                          them factors between 0 and 1 (a SUMPRODUCT, however written)
  a present-value row     it reads one row and equals its sum, and each cell of that row is a cash flow (in
                          another row, any sheet, the client model's too) times a factor between 0 and 1
  NPV, XNPV, and SUMPRODUCTs with inline factors: recognised from the formula (dcftrace._core), checked by value
For each discounting found it sets down the facts: the cash-flow row and its periods, the factors and what they
are computed from (the rate, the valuation date, the period dates: the single cells the factor cells read), a
Gordon terminal value if there is one, and the path from the discounting up to the figure in line-item words (the
mid of a low and a high rate, less the distribution payable, ...). Then where the cash flows come from: the
client-model rows they read, the rows to find in this year's model.
"""
import json
import re
from collections import Counter, deque

import dcf
import dcftrace
import rodb
from overlay import _a1, _show, parse_a1
from xlruntime import XLError, to_date

WALK_LIMIT = 30000
CLOSE = 1e-6
# a client row the discounting's timing depends on, not an amount: its label says it's a flag or a period marker
TIMING = re.compile(r"\bflags?\b|\bperiod (start|end|number|no\.?|counter)\b|\b(start|end) of (the )?period\b|\btiming\b|"
                    r"\bcounter\b|\bswitch\b|\bindicator\b|\bdates?\b", re.I)


def _close(a, b) -> bool:
    return isinstance(a, float) and isinstance(b, float) and abs(a - b) <= CLOSE * max(1.0, abs(a), abs(b))


def _num(v):
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def _a1k(k) -> str:
    src, s, r, c = k
    return (f"[{src}]" if src != "" else "") + _a1(s, r, c)


class Facts:
    def __init__(self, sess, summary: dict):
        self.sess, self.B = sess, sess.B
        self.w = summary["wiring"]
        self.summary = summary
        self.db = rodb.connect(self.w["overlay"]["db_path"])
        self.names = dcftrace._names(self.db)
        self.levers = {parse_a1(l["cell"]): l for l in summary.get("levers") or []}
        self._reads, self._formula = {}, {}
        self.client_labels = (sess.prior or sess.ov).labels()

    # ---- cells -----------------------------------------------------------------------------------------------
    def val(self, k):
        v = self.B.get(*k)
        return v if not isinstance(v, bool) else float(v)

    def num(self, k):
        return _num(self.val(k))

    def kind(self, k) -> str:
        src, s, r, c = k
        if src != "":
            return "client" if src == self.w.get("client_link") else "saved"
        if s in self.B.overlay:
            return "formula" if self.B.is_formula(s, r, c) else "input"
        return "client" if s in self.sess.client_sheets else "saved"

    def reads(self, k) -> set:
        if k not in self._reads:
            self._reads[k] = self.B.reads(k[1], k[2], k[3])[0] if self.kind(k) == "formula" else set()
        return self._reads[k]

    def formula(self, k) -> str | None:
        if k not in self._formula:
            src, s, r, c = k
            row = None if src != "" else self.db.execute(
                "SELECT formula FROM cells WHERE sheet=? AND row=? AND col=?", (s, r, c)).fetchone()
            self._formula[k] = row[0] if row else None
        return self._formula[k]

    def label(self, k) -> str:
        src, s, r, c = k
        if self.kind(k) == "client":
            return self.client_labels.get((s, r), "")
        return self.sess.ov.labels().get((s, r), "")

    def words(self, k) -> str | None:
        f = self.formula(k)
        if not f:
            return None
        try:
            return dcftrace.words(self.db, dcftrace._expand(self.db, dcftrace._STR.sub('""', f), self.names), k[1])
        except Exception:
            return f

    def timeline(self, k) -> dict:
        """col -> period date (serial) of the row's sheet, in whichever workbook holds it."""
        src, s, r, c = k
        wb = (self.sess.prior or self.sess.ov) if src != "" else self.sess.ov
        return wb.timeline(s)

    # ---- rows among the cells a formula reads ----------------------------------------------------------------
    @staticmethod
    def rows_of(cells) -> dict[tuple, list]:
        """{(src, sheet, row): [cells in column order]} for rows read three cells or more."""
        by: dict[tuple, list] = {}
        for k in cells:
            by.setdefault(k[:3], []).append(k)
        return {g: sorted(ks, key=lambda k: k[3]) for g, ks in by.items() if len(ks) >= 3}

    @staticmethod
    def _discount_like(fs: list[float]) -> bool:
        live = [f for f in fs if f]
        if len(live) < 3 or not all(0 < f <= 1.0000001 for f in live):
            return False
        drops = sum(b <= a + 1e-12 for a, b in zip(live, live[1:]))
        return drops >= 0.8 * (len(live) - 1)

    # ---- a discounting at one cell ---------------------------------------------------------------------------
    def detect(self, k) -> dict | None:
        v = self.num(k)
        if v is None:
            return None
        f = self.formula(k) or ""
        # NPV, XNPV and SUMPRODUCTs with factors computed inline: from the formula, checked by value
        try:
            body = dcftrace._expand(self.db, dcftrace._STR.sub('""', f), self.names)
            whole = body.strip().lstrip("=").lstrip("+").strip()
            for call in dcftrace._calls(body):
                if call[0] not in ("NPV", "XNPV", "SUMPRODUCT"):
                    continue
                core = dcftrace._core(self.db, k[1], k[2], k[3], call, whole == call[2])
                if core and (not core["whole"] or _close(float(core["pv"]), v)):
                    sheet, row, cols = dcf._row_range(self.db, core["cashflow"])
                    cf = [("", sheet, row, c) for c in cols]
                    fac = core["factors"]
                    fcells = None
                    if core.get("factor_row"):
                        fs_, fr_, fcols = dcf._row_range(self.db, core["factor_row"])
                        fcells = [("", fs_, fr_, c) for c in fcols]
                    return self._core(k, core["kind"], core["what"], cf, [fac.get(c, 0.0) for c in cols], fcells,
                                      float(core["pv"]), whole=core["whole"], static=core)
        except Exception:
            pass
        rows = self.rows_of(self.reads(k))
        # cash flows x factors: two rows read, the value their sum of products
        groups = list(rows.items())
        for i, (ga, a) in enumerate(groups):
            for gb, b in groups[i + 1:]:
                if len(a) != len(b):
                    continue
                va, vb = [self.num(x) or 0.0 for x in a], [self.num(x) or 0.0 for x in b]
                if not _close(sum(x * y for x, y in zip(va, vb)), v):
                    continue
                for cf, fac, fv in ((a, b, vb), (b, a, va)):
                    if self._discount_like(fv):
                        return self._core(k, "cash flows x factors", "the sum of a cash-flow row times a factor row",
                                          cf, fv, fac, v)
        # a present-value row: the value is the row's sum, each cell a cash flow times a factor
        for g, ps in groups:
            pv = [self.num(x) or 0.0 for x in ps]
            if not _close(sum(pv), v):
                continue
            found = self._pv_row(ps)
            if found:
                cf, fs, factor_cells = found
                return self._core(k, "present-value row", "the sum of a row of present values, each a cash flow "
                                                          "times a factor", cf, fs, factor_cells, v, pv_row=ps)
        return None

    def _pv_row(self, ps: list) -> tuple | None:
        """The cash-flow row behind a row of present values: the row each PV cell reads whose cell, divided into
        the PV, gives a factor between 0 and 1 falling over time. -> (cash-flow cells, factors, factor cells)."""
        if not all(self.kind(p) == "formula" for p in ps):
            return None
        per_row: dict[tuple, dict] = {}
        for p in ps:
            pv = self.num(p) or 0.0
            for a in self.reads(p):
                x = self.num(a)
                if a[:3] == p[:3] or x is None:
                    continue
                if x == 0:
                    per_row.setdefault(a[:3], {})[p] = (a, 0.0)
                elif 0 < pv / x <= 1.0000001:
                    per_row.setdefault(a[:3], {})[p] = (a, pv / x)
        best = None
        for g, hits in per_row.items():
            if len(hits) < 0.8 * len(ps):
                continue
            fs = [hits[p][1] if p in hits else 0.0 for p in ps]
            if not self._discount_like(fs):
                continue
            # the factor row, if the PV cells also read one: the row whose values are the factors
            score = sum(1 for p in ps if p in hits)
            if best is None or score > best[0]:
                best = (score, [hits[p][0] if p in hits else None for p in ps], fs)
        if not best:
            return None
        _, cf, fs = best
        cf = [c for c in cf if c]
        fac_row = None
        for g, cells in self.rows_of({a for p in ps for a in self.reads(p)}).items():
            vals = [self.num(c) or 0.0 for c in cells]
            if len(cells) == len(ps) and all(abs(x - y) < 1e-9 for x, y in zip(vals, fs) if x or y):
                fac_row = cells
        return cf, fs, fac_row

    def _core(self, k, kind, what, cf, fs, factor_cells, pv, pv_row=None, whole=True, static=None) -> dict:
        cf_row = cf[0][:3]
        tl = self.timeline(cf[0])
        live = [(c, f) for c, f in zip(cf, fs) if f and (self.num(c) or 0.0)]
        dates = [tl.get(c[3]) for c, _ in live]
        core = {"cell": _a1k(k), "kind": kind, "what": what, "value": self.num(k), "pv": pv,
                "pv_recomputed": sum((self.num(c) or 0.0) * f for c, f in zip(cf, fs)),
                "cashflow": {"row": _a1k(cf[0]).rsplit("!", 1)[0] + f"!r{cf_row[2]}", "label": self.label(cf[0]),
                             "first": _a1k(cf[0]), "last": _a1k(cf[-1]), "where": self.kind(cf[0]),
                             "periods": len(live), "undiscounted": sum(self.num(c) or 0.0 for c, _ in live),
                             "first_period": to_date(dates[0]).isoformat() if dates and dates[0] else None,
                             "last_period": to_date(dates[-1]).isoformat() if dates and dates[-1] else None},
                "factors": {"first": fs[0] if fs else None, "last_live": live[-1][1] if live else None,
                            "row": (_a1k(factor_cells[0]).rsplit("!", 1)[0] + f"!r{factor_cells[0][2]}") if factor_cells else None,
                            "row_label": self.label(factor_cells[0]) if factor_cells else None},
                "whole": whole}
        core["ties"] = _close(core["pv_recomputed"], pv) if whole else None
        core["factors"]["inputs"] = self._factor_inputs(factor_cells or pv_row or [])
        if static and static.get("method"):
            core["factors"]["method"] = {k2: static["method"].get(k2) for k2 in ("rate", "valuation_date", "timing", "day_count")}
        elif static and static.get("rate") is not None:
            core["factors"]["method"] = {"rate": static.get("rate"), "valuation_date": static.get("valuation_date"),
                                         "timing": "end", "day_count": "actual/365" if static["kind"] == "xnpv" else None}
        core["terminal"] = self._terminal(cf)
        core["_cf"] = cf
        return core

    def _factor_inputs(self, cells: list) -> list[dict]:
        """What the factors are computed from: the single cells every factor cell reads (the rate, the valuation
        date, a mid-period half), and the rows they read column by column (the period dates)."""
        cells = [c for c in cells if self.kind(c) == "formula"]
        if not cells:
            return []
        pick = [cells[0], cells[len(cells) // 2], cells[-1]]
        common = set.intersection(*(self.reads(c) for c in pick))
        out = []
        for a in sorted(common, key=lambda a: (a[1], a[2], a[3])):
            v = self.num(a)
            what = ("a date" if v and 20000 < v < 80000 else "a rate" if v and 0 < v < 0.5 else "a number")
            lev = self.levers.get(a[1:]) if a[0] == "" else None
            out.append({"cell": _a1k(a), "label": self.label(a), "value": _show(self.val(a)), "what": what,
                        "date": to_date(v).isoformat() if what == "a date" else None,
                        "source": "an assumption on the overlay" if lev or self.kind(a) == "input" else
                                  "the client model" if self.kind(a) == "client" else self.kind(a)})
        varying = self.rows_of({a for c in cells for a in self.reads(c)} - common)
        for g, cs in varying.items():
            v = self.num(cs[0])
            out.append({"cell": _a1k(cs[0]).rsplit("!", 1)[0] + f"!r{g[2]}", "label": self.label(cs[0]),
                        "value": None, "what": "period dates" if v and 20000 < v < 80000 else "a row read column by column",
                        "source": self.kind(cs[0])})
        return out

    def _terminal(self, cf: list) -> dict | None:
        """A Gordon terminal value in the last cash flows (or what they read): X x (1 + g) / (r - g)."""
        for c in reversed(cf[-2:]):
            for t in [c] + sorted(self.reads(c), key=lambda a: (a[1], a[2], a[3])):
                tv = self.num(t)
                if not tv or self.kind(t) != "formula":
                    continue
                singles = [a for a in self.reads(t) if (self.num(a) is not None)]
                small = [a for a in singles if 0 <= (self.num(a) or -1) < 0.3]
                for g in small:
                    for r in small:
                        gv, rv = self.num(g), self.num(r)
                        if g == r or rv <= gv:
                            continue
                        for x in singles:
                            xv = self.num(x)
                            if x in (g, r) or not xv:
                                continue
                            if _close(xv * (1 + gv) / (rv - gv), tv) or _close(xv / (rv - gv), tv):
                                return {"cell": _a1k(t), "label": self.label(t), "value": tv,
                                        "growth": {"cell": _a1k(g), "label": self.label(g), "value": gv},
                                        "rate": {"cell": _a1k(r), "label": self.label(r), "value": rv},
                                        "from": {"cell": _a1k(x), "label": self.label(x), "value": xv}}
        return None

    # ---- the walk -----------------------------------------------------------------------------------------------
    def walk(self, start) -> tuple[list[dict], dict]:
        """From the figure's cell down what it reads; a discounting found stops the walk below it (its reads are
        its cash flows and factors). -> (discountings, parent of each cell visited)."""
        parent, seen, queue, cores = {start: None}, {start}, deque([start]), []
        while queue and len(seen) < WALK_LIMIT:
            k = queue.popleft()
            if self.kind(k) != "formula":
                continue
            core = self.detect(k)
            if core:
                cores.append(core)
                continue
            for a in sorted(self.reads(k), key=lambda a: (str(a[0]), a[1], a[2], a[3])):
                if a not in seen:
                    seen.add(a)
                    parent[a] = k
                    queue.append(a)
        return cores, parent

    def path(self, start, cell, parent) -> list[dict]:
        """The cells from the figure down to a discounting, each with its formula in words and value."""
        k = ("",) + parse_a1(cell) if "[" not in cell else None
        chain = []
        while k is not None:
            chain.append(k)
            k = parent.get(k)
        chain.reverse()
        out = []
        for i, k in enumerate(chain):
            nxt = chain[i + 1] if i + 1 < len(chain) else None
            others = [a for a in self.reads(k) if a != nxt] if nxt else []
            singles = [a for a in others if len(self.rows_of([a])) == 0]
            out.append({"cell": _a1k(k), "label": self.label(k), "value": _show(self.val(k)), "words": self.words(k),
                        "with": [{"cell": _a1k(a), "label": self.label(a), "value": _show(self.val(a))}
                                 for a in sorted(singles, key=lambda a: (a[1], a[2], a[3]))[:6]]})
        return out

    def origins(self, cf: list) -> list[dict]:
        """The client-model rows the cash flows come from: from each cash-flow cell, what it reads, down to the
        client model's cells (last year's model: the rows to find this year)."""
        if self.kind(cf[0]) == "client":
            g = cf[0][:3]
            return [{"row": _a1k(cf[0]).rsplit("!", 1)[0] + f"!r{g[2]}", "label": self.label(cf[0]), "cells": len(cf),
                     "how": "the cash flows are read straight from it"}]
        hits: dict[tuple, list] = {}
        seen, queue = set(cf), deque(cf)
        while queue and len(seen) < WALK_LIMIT:
            k = queue.popleft()
            for a in self.reads(k):
                if a in seen:
                    continue
                seen.add(a)
                if self.kind(a) == "client":
                    hits.setdefault(a[:3], []).append(a)
                elif self.kind(a) == "formula":
                    queue.append(a)
        out = []
        for g, cells in sorted(hits.items(), key=lambda kv: -len(kv[1])):
            kind, why = self.row_kind(g[1], g[2], self.label(cells[0]))
            out.append({"row": _a1k(cells[0]).rsplit("!", 1)[0] + f"!r{g[2]}", "label": self.label(cells[0]),
                        "cells": len(cells), "how": "read by the overlay's cash-flow calculation", "kind": kind,
                        "kind_why": why})
        return out[:40]

    def row_kind(self, s, r, label) -> tuple[str, str]:
        """("amount" | "timing", why): whether a client row the cash flows read carries amounts, or the timing
        the discounting depends on (period flags, period dates). By the whole row in last year's model, not the
        cells read: a single integer in the date range is an amount (a fee of 45,000), dates are several, rising."""
        if TIMING.search(label or ""):
            return "timing", "its label says it's a flag or a date"
        wb = self.sess.prior or self.sess.ov
        vals = wb.sheet(s)
        nums = [v for c in sorted(wb.timeline(s)) if (v := _num(vals.get((r, c)))) is not None]
        nz = [v for v in nums if v]
        if nz and all(v in (0.0, 1.0) for v in nums):
            return "timing", "its values are 0/1 flags"
        if len(nz) >= 2 and all(v.is_integer() and 30000 <= v <= 80000 for v in nz) and all(b > a for a, b in zip(nz, nz[1:])):
            return "timing", "its values are dates"
        return "amount", ""

    def run(self, cell: str) -> dict:
        self.sess.configure("workbook")
        start = ("",) + parse_a1(cell)
        value = self.val(start)
        saved = self.sess.ov.value(*start[1:])
        cores, parent = self.walk(start)
        for c in cores:
            c["path"] = self.path(start, c["cell"], parent)
            c["origins"] = self.origins(c.pop("_cf"))
        return {"cell": cell, "label": self.label(start), "value": _show(value), "excel": _show(saved),
                "python_equals_excel": _close(_num(value), _num(saved)) if _num(value) is not None else value == saved,
                "discountings": cores, "cells_followed": len(parent),
                "frontier": None if cores else self.frontier(parent)}

    def frontier(self, parent: dict) -> dict:
        """Where the walk stopped, when it found no discounting: the cells it reached that it can't follow further
        (the client model's, inputs, other sheets), by row, and the first formulas it went through. That says
        where the discounting must be: in the client model's sheets (not compiled, so not followed), past the
        walk's limit, or in a shape not recognised."""
        leaves = Counter()
        cells = {}
        for k in parent:
            kind = self.kind(k)
            if kind != "formula":
                g = (kind,) + k[:3]
                leaves[g] += 1
                cells.setdefault(g, k)
        rows = [{"kind": g[0], "row": _a1k(cells[g]).rsplit("!", 1)[0] + f"!r{g[3]}", "label": self.label(cells[g]),
                 "cells": n, "has_formulas": bool(self.db.execute(
                     "SELECT 1 FROM cells WHERE sheet=? AND row=? AND formula IS NOT NULL LIMIT 1", (g[2], g[3])).fetchone())
                 if g[1] == "" else None}
                for g, n in sorted(leaves.items(), key=lambda kv: (kv[0][0] != "client", -kv[1]))[:30]]
        visited = [k for k in parent if self.kind(k) == "formula"][:15]
        return {"limit_reached": len(parent) >= WALK_LIMIT, "stopped_at": rows,
                "formulas": [{"cell": _a1k(k), "label": self.label(k), "value": _show(self.val(k)), "words": self.words(k),
                              "reads": len(self.reads(k))} for k in visited]}


def facts(sess, summary: dict, cell: str) -> dict:
    """The facts behind the figure at cell (see the module doc). Runs in overlay.deep."""
    f = Facts(sess, summary)
    try:
        return f.run(cell)
    finally:
        sess.configure("workbook")


def text(fx: dict) -> str:
    """The facts as plain text, to paste into a message."""
    n = lambda v: f"{v:,.4f}" if isinstance(v, float) else str(v)
    L = [f"{fx['label'] or fx['cell']} ({fx['cell']}) = {n(fx['value'])}; Excel saved {n(fx['excel'])}"
         + ("" if fx["python_equals_excel"] else "  <-- Python differs from Excel"),
         f"{len(fx['discountings'])} discounting(s) found under it ({fx['cells_followed']:,} cells followed)"]
    fr = fx.get("frontier")
    if fr:
        L.append("No discounting found. Where the walk stopped" + (" (it reached its limit of cells)" if fr["limit_reached"] else "") + ":")
        for x in fr["stopped_at"]:
            L.append(f"  [{x['kind']}] {x['row']} {x['label']} ({x['cells']} cells" + (", formulas in the workbook" if x.get("has_formulas") else "") + ")")
        L.append("The first formulas it went through:")
        for x in fr["formulas"]:
            L.append(f"  {x['cell']} {x['label']} = {n(x['value'])} reads {x['reads']} cell(s)" + (f"   [{x['words']}]" if x.get("words") else ""))
    for i, c in enumerate(fx["discountings"], 1):
        cf, fa = c["cashflow"], c["factors"]
        L += ["", f"{i}. {c['cell']} {c['what']}: PV {n(c['pv'])}; cash flows x factors recomputed {n(c['pv_recomputed'])}"
                  + (" (ties)" if c.get("ties") else " (does NOT tie)" if c.get("ties") is False else ""),
              f"   cash flows: {cf['row']} {cf['label']} ({cf['where']}), {cf['periods']} periods "
              f"{cf['first_period']} to {cf['last_period']}, undiscounted {n(cf['undiscounted'])}",
              f"   factors: first {n(fa['first'])}, last {n(fa['last_live'])}" + (f"; row {fa['row']} {fa['row_label'] or ''}" if fa.get("row") else "")
              + (f"; method {json.dumps(fa['method'])}" if fa.get("method") else "")]
        for x in fa.get("inputs") or []:
            L.append(f"     from {x['cell']} {x['label']}: {x['what']} {x.get('date') or x['value'] if x['value'] is not None else ''} ({x['source']})")
        if c.get("terminal"):
            t = c["terminal"]
            L.append(f"   terminal value: {t['cell']} {t['label']} {n(t['value'])} = {t['from']['label'] or t['from']['cell']} "
                     f"{n(t['from']['value'])} grown at {t['growth']['value']:.4%} ({t['growth']['cell']}) over rate {t['rate']['value']:.4%} ({t['rate']['cell']})")
        L.append("   up to the figure:")
        for p in c["path"]:
            L.append(f"     {p['cell']} {p['label']} = {n(p['value'])}" + (f"   [{p['words']}]" if p.get("words") else "")
                     + (" with " + "; ".join(f"{w['label'] or w['cell']} {n(w['value'])}" for w in p["with"]) if p.get("with") else ""))
        L.append("   cash flows come from (last year's client model):")
        amounts = [o for o in c["origins"] if o.get("kind", "amount") == "amount"]
        for o in amounts or [{"row": "-", "label": "none found (the cash flows are the overlay's own)", "cells": 0, "how": ""}]:
            L.append(f"     {o['row']} {o['label']} ({o['cells']} cells; {o['how']})")
        timing = [o for o in c["origins"] if o.get("kind") == "timing"]
        if timing:
            L.append("   timing the discounting depends on (last year's client model):")
            for o in timing:
                L.append(f"     {o['row']} {o['label']} ({o['cells']} cells; {o['kind_why']})")
    return "\n".join(L)
