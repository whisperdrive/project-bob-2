"""Last year's client-model rows in this year's model, when the model changed between valuations.

Rolling a valuation forward means reading, for every client row the overlay reads, the same line item in this
year's model. Models change between valuations: rows are inserted, renamed, restructured, sheets renamed. A row
is looked for in several independent ways, and the one the evidence supports best is taken:
  label        the same label on the same sheet (its n-th occurrence, else the nearest), or, where the sheet is
               gone (renamed), the only row with that label anywhere
  history      the same values in the periods both models have as history: actual years don't change between
               versions, so a row whose actuals equal last year's is the same line item whatever it's called now;
               forecast years that stay close (revised, not replaced) count too
  words        a label on the same sheet sharing most of its words ("Free cash flow" -> "Free cash flow after
               working capital"), with the forecast staying close
  neighbours   the same rows around it in the dependency graph (the edges table): it reads rows with the same
               labels (EBITDA, capital expenditure) and is read by the same ones
  banner       a model's summary cells in the first rows of its sheets (a total, a value at the valuation date,
               often repeated on several sheets) that read the row last year: the same banner this year reads
               the row to use
A candidate of another shape counts for less: last year's row calculated (formulas) and this one typed values,
or the other way round. A reconciliation sheet of pasted values ("LINKED EBITDA") has last year's history exactly
and a full series, and would otherwise win every row it copies.
A person's pick for a row always wins: a row, or last year's values kept on purpose (STAND_IN). explain() gives what was found for a row and why, with the alternatives.
A row with nothing to find (no label, no numbers, no formulas last year: a spacer inside a range a formula reads) is
settled by code: its blanks stand in, and no model is asked about it.
"""
import bisect
import re
from collections import defaultdict

BANNER_ROWS = 10
EXACT = 1e-9
SAME_SHEET = 0.35     # a sheet of the same name is the same sheet only if it shares this much of its labels
RENAMED_SHEET = 0.5   # a sheet of another name is a renamed one if it shares this much
CONFIDENT = 0.5      # a row found with less than this needs a person's look before its figures are this year's
CHECK_GAP = 0.15     # a row carries last year's numbers when its median difference from them is within this
CHECK_PERIODS = 3    # ... over at least this many periods side by side
SHAPE = 0.6           # a candidate whose share of formulas differs from last year's row's by this much is another shape
STAND_IN = "stand-in"  # a person's pick: keep last year's values for the row
LAYOUT = 0.9         # an unlabelled row found at its own row number, on a sheet laid out as before
VERSION = 2          # bump when finding changes: the agents' picks made under another version are dropped and redone
                     # (2: unlabelled rows followed by the layout, blank rows settled by code)


def _norm(label: str) -> str:
    return re.sub(r"\s+", " ", (label or "").strip().lower())


def _words(label: str) -> set[str]:
    return set(re.findall(r"[a-z][a-z0-9]+", (label or "").lower())) - {"the", "and", "of", "for", "to", "in"}


def _jaccard(a: set, b: set) -> float:
    return len(a & b) / len(a | b) if a and b else 0.0


class RowFinder:
    def __init__(self, rowmap, prior, current, picks: dict | None = None):
        """rowmap: overlay.RowMap (the label matching); prior, current: overlay.Workbook."""
        self.rowmap, self.prior, self.current = rowmap, prior, current
        self.picks = dict(picks or {})   # {(sheet, row): (sheet, row) or STAND_IN}: a person's or the agents' choice
        self.pick_by: dict = {}          # {(sheet, row): "you" | "agent"}: whose choice it is (a person's by default)
        self._cache: dict = {}
        self._idx = None
        self._shapes: dict = {}
        self._numeric: dict = {}         # {sheet: {rows with a number other than 0}} in last year's model

    # ---- what each model has ---------------------------------------------------------------------------------
    def _index(self):
        if self._idx is not None:
            return self._idx
        cur = self.current
        sheets = [s for (s,) in cur.db.execute("SELECT sheet FROM sheets")]
        labels = cur.labels()
        by_label = defaultdict(list)
        for (s, r), lab in labels.items():
            if lab:
                by_label[_norm(lab)].append((s, r))
        dates = {s: {v: c for c, v in cur.timeline(s).items()} for s in sheets}
        by_word = defaultdict(list)
        for k, lab in labels.items():
            for w in _words(lab):
                by_word[w].append(k)
        self._idx = {"sheets": set(sheets), "labels": labels, "by_label": by_label, "dates": dates, "by_word": by_word,
                     "edges": (self._edges(self.prior), self._edges(cur)), "values": {}}
        return self._idx

    @staticmethod
    def _edges(wb) -> tuple[dict, dict]:
        """(reads, read_by): {(sheet, row): {(sheet, row)}} from the model's edges table."""
        reads, read_by = defaultdict(set), defaultdict(set)
        try:
            for s, r, ds, dr in wb.db.execute("SELECT src_sheet, src_row, dst_sheet, dst_row FROM edges "
                                              "WHERE kind IN ('direct', 'offset', 'active')"):
                if (s, r) != (ds, dr):
                    reads[(s, r)].add((ds, dr))
                    read_by[(ds, dr)].add((s, r))
        except Exception:
            pass
        return reads, read_by

    def sheet_for(self, s: str) -> str | None:
        """The sheet in this year's model that last year's sheet s is: the same name if it shares at least
        SAME_SHEET of its line-item labels (a rebuilt model can reuse a name for something else), else the sheet of
        another name sharing RENAMED_SHEET of them (renamed); None when no sheet does."""
        idx = self._index()
        idx.setdefault("sheet_map", {})
        if s in idx["sheet_map"]:
            return idx["sheet_map"][s]
        mine = {_norm(l) for (sh, _), l in self.prior.labels().items() if sh == s and l}
        labels_of = lambda sh: {_norm(l) for (x, _), l in idx["labels"].items() if x == sh and l}
        if s in idx["sheets"] and (not mine or _jaccard(mine, labels_of(s)) >= SAME_SHEET):
            idx["sheet_map"][s] = s
            return s
        best, share = None, 0.0
        for sh in idx["sheets"]:
            if sh == s or self._has_prior_sheet(sh):
                continue
            j = _jaccard(mine, labels_of(sh))
            if j > share:
                best, share = sh, j
        idx["sheet_map"][s] = best if share >= RENAMED_SHEET else None
        return idx["sheet_map"][s]

    def _has_prior_sheet(self, sheet: str) -> bool:
        return self.prior.db.execute("SELECT 1 FROM sheets WHERE sheet=?", (sheet,)).fetchone() is not None

    def _values_at(self, sheet: str, when: float) -> dict:
        """This year's model at one period date on one sheet: {rounded value: [rows]} (each sheet indexed once)."""
        idx = self._index()
        c = idx["dates"].get(sheet, {}).get(when)
        if c is None:
            return {}
        if sheet not in idx["values"]:
            by_col: dict = defaultdict(lambda: defaultdict(list))
            cols = set(idx["dates"][sheet].values())
            for (r, cc), v in self.current.sheet(sheet).items():
                if cc in cols and isinstance(v, float) and v:
                    by_col[cc][float(f"{v:.9g}")].append(r)
            idx["values"][sheet] = by_col
        return idx["values"][sheet].get(c, {})

    def _series(self, wb, sheet: str, row: int) -> dict:
        """A row's values by period date."""
        tl = wb.timeline(sheet)
        vals = wb.sheet(sheet)
        return {when: vals.get((row, c)) for c, when in tl.items() if isinstance(vals.get((row, c)), float)}

    def _shape(self, wb, k) -> tuple | None:
        """(formulas, typed values) in a row, from the rows table (each workbook read once)."""
        if id(wb) not in self._shapes:
            try:
                self._shapes[id(wb)] = {(s, r): (nf or 0, nc or 0) for s, r, nf, nc in
                                        wb.db.execute("SELECT sheet, row, n_formula, n_const FROM rows")}
            except Exception:
                self._shapes[id(wb)] = {}
        return self._shapes[id(wb)].get(k)

    def other_shape(self, s, r, k) -> str | None:
        """Why candidate k is another shape than last year's row (formulas against typed values), or None."""
        a, b = self._shape(self.prior, (s, r)), self._shape(self.current, k)
        if not a or not b or not sum(a) or not sum(b):
            return None
        fa, fb = a[0] / sum(a), b[0] / sum(b)
        if abs(fa - fb) < SHAPE:
            return None
        what = lambda x, f: f"formulas ({x[0]} of {sum(x)})" if f >= 0.5 else f"typed values ({x[1]} of {sum(x)})"
        return f"last year's row is {what(a, fa)}, this one {what(b, fb)}: another kind of row (a pasted copy?)"

    # ---- the strategies ----------------------------------------------------------------------------------------
    def _by_label(self, s, r) -> list[tuple]:
        idx = self._index()
        to = self.sheet_for(s)
        if to == s:
            r2, how = self.rowmap.match(s, r)
            if r2:
                unlabelled = not self.prior.labels().get((s, r)) and not self.current.labels().get((s, r2))
                return [((s, r2), 1.0 if how.startswith("same label") else
                         LAYOUT if how.startswith("same row") and unlabelled else 0.4, how)]
        lab = _norm(self.prior.labels().get((s, r), ""))
        if to and to != s and lab:  # the sheet was renamed: the label on the sheet it became, occurrence as before
            n = sum(1 for (sh, rr), l in self.prior.labels().items() if sh == s and rr <= r and _norm(l) == lab)
            there = sorted(rr for (sh, rr) in idx["by_label"].get(lab, []) if sh == to)
            if there:
                r2 = there[min(n, len(there)) - 1]
                return [((to, r2), 0.9, f"same label on sheet {to} (last year's {s}, renamed)")]
        hits = idx["by_label"].get(lab, []) if lab else []
        where = f"sheet {s} isn't in this year's model" if s not in idx["sheets"] else \
            f"this year's {s} isn't last year's (they share few labels)" if not to else f"not on {to}"
        if len(hits) == 1:
            return [(hits[0], 0.8, f"the only row labelled '{self.prior.labels().get((s, r))}' ({where})")]
        if 1 < len(hits) <= 5:
            return [(k, 0.5, f"one of {len(hits)} rows labelled '{self.prior.labels().get((s, r))}' ({where})") for k in hits]
        return []

    def _history(self, mine: dict, k: tuple) -> tuple | None:
        """One candidate's history against last year's row: (score, text, equal periods, differing history), or
        None when they share no period. A candidate blank in periods its sheet has dates for, where last year's
        row has values, scores less: an actuals sheet has the history exactly and no forecast, and would
        otherwise beat the row that carries both (a row with 9 of 51 periods can't outrank one with 51)."""
        theirs = self._series(self.current, *k)
        both = [w for w in mine if w in theirs and mine[w]]
        if not both:
            return None
        equal = [w for w in both if abs(theirs[w] - mine[w]) <= EXACT * max(1.0, abs(mine[w]))]
        later = [w for w in both if w not in equal]
        gaps = sorted(abs(theirs[w] / mine[w] - 1) for w in later)
        close = not gaps or gaps[len(gaps) // 2] < 0.25
        if not equal:
            return (0.0, "", 0, len(later))
        dates = self._index()["dates"].get(k[0], {})
        could = [w for w in mine if mine[w] and w in dates]
        cover = len(both) / len(could) if could else 1.0
        score = min(1.0, 0.5 * len(equal)) * (1.0 if close else 0.6) * min(1.0, cover / 0.8)
        return (score, f"{len(equal)} period(s) of history equal last year's"
                       + (f"; later years within {gaps[len(gaps) // 2]:.0%} (median)" if gaps else "")
                       + (f"; blank in {len(could) - len(both)} of the {len(could)} periods it has dates for"
                          if cover < 1 else ""), len(equal), len(later))

    def _by_history(self, s, r) -> list[tuple]:
        """Rows whose values equal last year's in the periods both models have (history), on any sheet."""
        mine = self._series(self.prior, s, r)
        if not mine:
            return []
        idx = self._index()
        exact = defaultdict(int)
        for sheet, dates in idx["dates"].items():
            for when, v in mine.items():
                if not v or when not in dates:
                    continue
                for r2 in self._values_at(sheet, when).get(float(f"{v:.9g}"), []):
                    exact[(sheet, r2)] += 1
        out = []
        for k in sorted(exact, key=lambda k: (-exact[k], k))[:40]:
            h = self._history(mine, k)
            if h and h[2] and (h[2] >= 2 or h[3] <= 20 * h[2]):
                out.append((k, h[0], h[1]))
        return out

    def _by_words(self, s, r) -> list[tuple]:
        """A label sharing most of its words, with the forecast staying close: on the sheet last year's sheet is
        this year, or anywhere when no sheet corresponds (then with more words in common)."""
        idx = self._index()
        mine_words = _words(self.prior.labels().get((s, r), ""))
        if not mine_words:
            return []
        to = self.sheet_for(s)
        floor = 0.4 if to else 0.5
        mine = self._series(self.prior, s, r)
        cands = {k for w in mine_words for k in idx["by_word"].get(w, ())}
        out = []
        for (s2, r2) in sorted(cands):
            if to and s2 != to:
                continue
            lab = idx["labels"].get((s2, r2), "")
            sim = _jaccard(mine_words, _words(lab))
            if sim < floor or _norm(lab) == _norm(self.prior.labels().get((s, r), "")):
                continue
            theirs = self._series(self.current, s2, r2)
            both = [w for w in mine if w in theirs and mine[w]]
            gaps = sorted(abs(theirs[w] / mine[w] - 1) for w in both)
            if both and gaps[len(gaps) // 2] < 0.1:
                out.append(((s2, r2), sim, f"label '{lab}' shares {sim:.0%} of its words; values within "
                                           f"{gaps[len(gaps) // 2]:.0%} (median) in {len(both)} period(s)"))
        return sorted(out, key=lambda x: (-x[1], x[0]))[:3]

    def _by_neighbours(self, s, r) -> list[tuple]:
        """Rows that read rows labelled as this one's inputs are and are read by rows labelled as its users."""
        idx = self._index()
        (p_reads, p_by), (c_reads, c_by) = idx["edges"]
        plab, clab = self.prior.labels(), idx["labels"]
        ins = {_norm(plab.get(x, "")) for x in p_reads.get((s, r), ())} - {""}
        outs = {_norm(plab.get(x, "")) for x in p_by.get((s, r), ())} - {""}
        if not ins and not outs:
            return []
        cands = set()
        for lab in ins:
            for x in idx["by_label"].get(lab, []):
                cands |= c_by.get(x, set())
        for lab in outs:
            for x in idx["by_label"].get(lab, []):
                cands |= c_reads.get(x, set())
        out = []
        for k in cands:
            ci = {_norm(clab.get(x, "")) for x in c_reads.get(k, ())} - {""}
            co = {_norm(clab.get(x, "")) for x in c_by.get(k, ())} - {""}
            parts = [j for j, a, b in ((_jaccard(ins, ci), ins, ci), (_jaccard(outs, co), outs, co)) if a]
            score = sum(parts) / len(parts) if parts else 0.0
            if score >= 0.34:
                out.append((k, score, f"reads {', '.join(sorted(ins & ci)) or 'nothing alike'}; read by "
                                      f"{', '.join(sorted(outs & co)) or 'nothing alike'}"))
        return sorted(out, key=lambda x: (-x[1], x[0]))[:3]

    def _by_banner(self, s, r) -> list[tuple]:
        """The model's banner (summary cells in its sheets' first rows) that read the row last year: the same
        banner in this year's model reads the row to use."""
        idx = self._index()
        (p_reads, p_by), (c_reads, c_by) = idx["edges"]
        plab = self.prior.labels()
        banners = [b for b in p_by.get((s, r), ()) if b[1] <= BANNER_ROWS and plab.get(b)]
        out = {}
        for b in banners:
            lab = _norm(plab[b])
            same = [x for x in idx["by_label"].get(lab, []) if x[1] <= BANNER_ROWS]
            for x in same:
                targets = [t for t in c_reads.get(x, ()) if t[1] > BANNER_ROWS]
                if len(targets) == 1:
                    k = targets[0]
                    seen = out.get(k, (k, 0.0, ""))
                    out[k] = (k, min(1.0, seen[1] + 0.7), f"the banner '{plab[b]}' ({x[0]}!r{x[1]}) reads it, as it read "
                                                          f"last year's ({b[0]}!r{b[1]})")
        return sorted(out.values(), key=lambda x: (-x[1], x[0]))

    # ---- deciding ----------------------------------------------------------------------------------------------
    WEIGHTS = {"label": 0.3, "history": 0.35, "words": 0.2, "neighbours": 0.15, "banner": 0.25, "shape": 0.0}

    def explain(self, s: str, r: int) -> dict:
        """{found: (sheet, row) or None, how, evidence: [(strategy, text)], confidence, alternatives}."""
        key = (s, r)
        if key in self._cache:
            return self._cache[key]
        if key in self.picks:
            k = self.picks[key]
            by = self.pick_by.get(key, "you")
            who = "the agents" if by == "agent" else "you"
            res = {"found": None if k == STAND_IN else k, "how": "the agents' pick" if by == "agent" else "your pick",
                   "confidence": 1.0, "alternatives": [], "in_place": False, "stand_in": k == STAND_IN, "by": by,
                   "evidence": [(who, "last year's values kept on purpose" if k == STAND_IN else f"picked by {who}")]}
            self._cache[key] = res
            return res
        found = defaultdict(dict)
        for name, fn in (("label", self._by_label), ("history", self._by_history), ("words", self._by_words),
                         ("neighbours", self._by_neighbours), ("banner", self._by_banner)):
            try:
                for k, score, text in fn(s, r):
                    if k and (k not in found or name not in found[k] or found[k][name][0] < score):
                        found[k][name] = (score, text)
            except Exception:
                continue
        mine = self._series(self.prior, s, r)
        for k, ev in found.items():  # every candidate's own history, not only the closest lookalikes'
            if "history" not in ev and mine:
                h = self._history(mine, k)
                if h and h[2]:
                    ev["history"] = (h[0], h[1])
        ranked = []
        home = self.sheet_for(s)
        odd = set()
        for k, ev in found.items():
            total = sum(self.WEIGHTS[n] * sc for n, (sc, _) in ev.items())
            if home:  # on the sheet last year's sheet became, or not
                total += 0.1 if k[0] == home else -0.1
            strong = any((n == "label" and sc >= 1.0) or (n == "history" and sc >= 0.5) or (n == "banner" and sc >= 0.7)
                         or (n == "neighbours" and sc >= 0.6) or (n == "label" and sc >= 0.8) for n, (sc, _) in ev.items())
            # the same label where it was, and last year's numbers in the periods both have: nothing beats that
            sure = "label" in ev and ev["label"][0] >= 1.0 and "history" in ev
            why = self.other_shape(s, r, k)
            if why:  # another kind of row: its evidence counts for less, and alone it isn't enough
                ev["shape"] = (0.0, why)
                total -= 0.2
                strong = sure = False
                odd.add(k)
            ranked.append((strong, total, k, ev, sure))
        ranked.sort(key=lambda x: (-x[4], -x[0], -x[1], x[2]))
        ranked = [x[:4] for x in ranked]
        res = {"found": None, "how": None, "evidence": [], "confidence": 0.0, "in_place": False,
               "alternatives": [{"row": f"{k[0]}!r{k[1]}", "label": self.current.labels().get(k, ""),
                                 "score": round(t, 2), "evidence": [f"{n}: {tx}" for n, (_, tx) in ev.items()]}
                                for _, t, k, ev in ranked[:4]]}
        if ranked:
            strong, total, k, ev = ranked[0]
            # the history outranks a label that contradicts it: the same label, but other numbers in periods
            # both models have, where another row has last year's numbers
            hist = next((x for x in ranked if "history" in x[3] and x[3]["history"][0] >= 0.5 and x[2] not in odd), None)
            if hist and hist[2] != k and "history" not in ev and mine:
                h = self._history(mine, k)
                if h and h[3] and not h[2]:
                    strong, total, k, ev = hist
            if strong or total >= 0.3:
                # the same label in the same place, of the same kind: confident even where there's no history to
                # add (a single value, no timeline)
                res["in_place"] = k[0] == s and "label" in ev and k not in odd and (
                    ev["label"][0] >= 1.0 or (ev["label"][0] >= LAYOUT and k[1] == r))
                res.update(found=k, confidence=round(max(0.0, min(1.0, total)), 2),
                           how=max(ev.items(), key=lambda kv: self.WEIGHTS[kv[0]] * kv[1][0])[0],
                           evidence=[(n, tx) for n, (_, tx) in sorted(ev.items(), key=lambda kv: -self.WEIGHTS[kv[0]] * kv[1][0])])
                res["alternatives"] = [a for a in res["alternatives"] if a["row"] != f"{k[0]}!r{k[1]}"]
        if not (res["found"] and (res["confidence"] >= CONFIDENT or res["in_place"])) and self.blank(s, r):
            res.update(found=None, how="nothing to find", confidence=1.0, in_place=False, stand_in=True, blank=True,
                       by="code", evidence=[("code", "last year's row has no label, no numbers and no formulas: nothing "
                                                     "to find, so its blanks stand in")])
        self._cache[key] = res
        return res

    def blank(self, s: str, r: int) -> bool:
        """Last year's row has nothing to find: no label, no number other than 0, no formulas."""
        if self.prior.labels().get((s, r)) or (self._shape(self.prior, (s, r)) or (0, 0))[0]:
            return False
        if s not in self._numeric:
            self._numeric[s] = {rr for (rr, _c), v in self.prior.sheet(s).items() if isinstance(v, float) and v}
        return r not in self._numeric[s]

    def family(self) -> float:
        """How alike the two models are: the share of line-item labels they have in common (a new version of a
        model shares most; a model rebuilt from the ground up, few)."""
        idx = self._index()
        if "family" not in idx:
            mine = {_norm(l) for l in self.prior.labels().values() if l}
            theirs = {_norm(l) for l in idx["labels"].values() if l}
            idx["family"] = round(_jaccard(mine, theirs), 3)
        return idx["family"]

    def locate(self, s: str, r: int) -> tuple | None:
        return self.explain(s, r)["found"]

    def why(self, s: str, r: int, labels: dict) -> str:
        if self.picks.get((s, r)) == STAND_IN:
            return "last year's values kept on purpose (your pick)"
        return self.rowmap.why(s, r, labels) + " (and no other way found it: not by its history, words, neighbours or banner)"

    def confident(self, s: str, r: int) -> bool:
        """Settled well enough to roll on without a person looking: picked (a row, or last year's values kept; a
        person's pick or the agents'), nothing to find, a confidence of CONFIDENT or more, or found in place (the
        same label, or an unlabelled row at its own row number on a sheet laid out as before)."""
        if (s, r) in self.picks:
            return True
        ex = self.explain(s, r)
        return bool(ex.get("blank")) or (ex["found"] is not None and (ex["confidence"] >= CONFIDENT or ex["in_place"]))

    def pick(self, s: str, r: int, to, by: str = "you") -> None:
        """A choice for a row: (sheet, row), STAND_IN (keep last year's values), or None (back to what's found);
        by: "you" (a person) or "agent". The agents never replace a person's choice."""
        if by == "agent" and self.pick_by.get((s, r), "you" if (s, r) in self.picks else None) == "you":
            return
        if to:
            self.picks[(s, r)] = to
            self.pick_by[(s, r)] = by
        else:
            self.picks.pop((s, r), None)
            self.pick_by.pop((s, r), None)
        self._cache.pop((s, r), None)

    # ---- checking a match by its numbers, and finding rows by them -------------------------------------------
    def check(self, s: str, r: int, k: tuple) -> dict:
        """Whether this year's row k carries last year's row's numbers: over the periods both have (no roll), the
        median gap between them, how many are equal (history), and whether it's the same kind of row. Forecasts
        are revised between valuations, not replaced: ok when the median gap is within CHECK_GAP over at least
        CHECK_PERIODS periods, the kind agrees, and the row isn't blank where last year's has values."""
        mine = self._series(self.prior, s, r)
        theirs = self._series(self.current, *k)
        both = [w for w in mine if w in theirs and mine[w]]
        dates = self._index()["dates"].get(k[0], {})
        could = [w for w in mine if mine[w] and w in dates]
        gaps = sorted(abs(theirs[w] / mine[w] - 1) for w in both)
        med = gaps[len(gaps) // 2] if gaps else None
        equal = sum(1 for g in gaps if g <= 1e-9)
        shape = self.other_shape(s, r, k)
        cover = len(both) / len(could) if could else 0.0
        ok = len(both) >= CHECK_PERIODS and med is not None and med <= CHECK_GAP and not shape and cover >= 0.8
        return {"row": f"{k[0]}!r{k[1]}", "periods": len(both), "median_gap": None if med is None else round(med, 4),
                "equal": equal, "cover": round(cover, 3), "shape": shape, "ok": ok,
                "text": (f"{len(both)} period(s) side by side, median difference {med:.1%}" if med is not None else
                         "no periods side by side") + (f", {equal} equal" if equal else "")
                        + (f"; {shape}" if shape else "") + (f"; blank in {len(could) - len(both)} of {len(could)}"
                                                              if could and cover < 1 else "")}

    def _sorted_values(self, sheet: str, col: int) -> tuple[list, list]:
        idx = self._index()
        key = ("sorted", sheet, col)
        if key not in idx:
            vals = sorted((v, rr) for (rr, cc), v in self.current.sheet(sheet).items()
                          if cc == col and isinstance(v, float) and v)
            idx[key] = ([v for v, _ in vals], [rr for _, rr in vals])
        return idx[key]

    def near(self, s: str, r: int, limit: int = 8, tol: float = CHECK_GAP) -> list[tuple]:
        """Rows of this year's model, on any sheet, whose values are close to last year's row's (within tol) in
        several of the same periods: a row found by its numbers when its label says nothing (or there's none)."""
        mine = [(w, v) for w, v in self._series(self.prior, s, r).items() if v]
        if len(mine) < CHECK_PERIODS:
            return []
        step = max(1, len(mine) // 8)
        sample = mine[::step][:8]
        hits = defaultdict(int)
        for sheet, dates in self._index()["dates"].items():
            for w, v in sample:
                c = dates.get(w)
                if c is None:
                    continue
                vals, rows = self._sorted_values(sheet, c)
                lo, hi = sorted((v * (1 - tol), v * (1 + tol)))
                for i in range(bisect.bisect_left(vals, lo), bisect.bisect_right(vals, hi)):
                    hits[(sheet, rows[i])] += 1
        need = max(CHECK_PERIODS, len(sample) // 2)
        return sorted((k for k, n in hits.items() if n >= need), key=lambda k: (-hits[k], k))[:limit]
