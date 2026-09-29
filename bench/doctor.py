"""The overlay doctor: why a figure comes out wrong in the Python overlay, and what to do about it.

It works from evidence first (examine: no model call), in four passes, then a reviewer model writes it up
(diagnose), held to that evidence.

  1. Layers: where does it break? Each of the report's figures through the layers it passes on the way to this
     year's value: the report -> the overlay as Excel saved it -> Python on the overlay's saved values (the
     workbook feed) -> Python on last year's client model (the prior feed) -> Python on this year's client
     model, rolled forward (the current feed). The first layer where a figure stops agreeing with the one before
     it is where to look: a report figure the saved overlay doesn't give points at the files; Python that
     differs from Excel on the overlay's own saved values is the compiler or the runtime (no client file is
     read there, so neither the wrong file nor the wrong row can cause it); a prior feed that differs from the
     saved values is the client model file; a current feed that errors is this year's model or the rows.
  2. Roots: why? From a figure that is wrong, follow the cells it reads (the runtime records them: Book.reads)
     while they are wrong too, to the cells where it starts: a formula whose own inputs are all fine, or a
     client value. Each root gets its cause: a formula that doesn't compile, a function Python doesn't have, a
     name the workbook doesn't define (a table reference, a LET parameter, a name on another sheet), a client
     cell that holds an error or a different value, or Python working the formula out differently from Excel.
     These all show as #NAME? on the page; here each says what it is.
  3. Files: is it the right file? The overlay's saved values tie to the report; its external links name the
     assigned client model; the client values the figures read are the ones the overlay last saw (else the
     client model assigned isn't the version the overlay was saved with); the client model is read at all.
  4. Rows: is it the right row? Each client line item the figures read, followed into this year's model: its
     history (periods up to last year's valuation date) should be the same numbers in both models, else the row
     picked this year is a different line item (or the units changed).

And what can be fixed here: a root that reads nothing from the client model and no assumption (so its value
can't change between feeds) can be held at Excel's saved value on every feed (holds). Holding them is checked
by recomputing the figures on every feed first.
"""
import inspect
import json
import re
import time
from collections import Counter
from datetime import date

import rodb
import xlcompile
import xlruntime
from overlay import _a1, _feed, _show, parse_a1, tie
from xlruntime import XLError, same, serial, to_date

WALK_LIMIT = 40000      # cells looked at per figure and feed
CLOSURE_LIMIT = 20000   # cells looked at to decide a hold is safe
HOLD_CANDIDATES = 400   # root cells checked for a hold, at most ...
HOLD_SECONDS = 90       # ... and for at most this long
# functions that fetch data from a provider's add-in: Excel's saved values are all a file holds of them
ADDIN = re.compile(r"^(CIQ|CIQ\w+|FDS\w*|BDP|BDH|BDS|BQL\w*|SPG\w*|RDP\.\w+|TR|RHISTORY|DSGRID|PITCHBOOK\w*|MSCI\w*|"
                   r"CAPIQ\w*|IQ_\w+|SNL\w*|REFINITIV\w*|EIKON\w*)$", re.I)


def _err(v) -> bool:
    return isinstance(v, XLError)


def _a1k(k) -> str:
    src, s, r, c = k
    return (f"[{src}]" if src != "" else "") + _a1(s, r, c)


class Doctor:
    def __init__(self, sess, summary: dict, progress=None):
        self.sess, self.summary, self.B = sess, summary, sess.B
        self.w = summary["wiring"]
        self.db = rodb.connect(self.w["overlay"]["db_path"])
        self.ctx = xlcompile.context(self.db, summary["sheets"], self.w["overlay"].get("source_path"))[0]
        self.labels = sess.ov.labels()
        self.client_labels = (sess.prior or sess.ov).labels()
        self.levers = {parse_a1(l["cell"]): l for l in summary.get("levers") or []}
        self.progress = progress or (lambda f, m: None)
        self.feed = None
        self._reads = {}
        self._formula = {}
        self._depends = {}  # cell -> why it can change between feeds, or None: its whole closure is clean
        self.xlsm = (self.w["overlay"].get("filename") or "").lower().endswith(".xlsm")

    # ---- cells -----------------------------------------------------------------------------------------------
    def use(self, feed: str):
        """Configure the session as the pages do for a feed (the current feed rolled forward to its date)."""
        defaults, _, months = _feed(self.summary, feed, None, None)
        self.sess.configure(feed, defaults, months or 0)
        self.feed = feed

    def val(self, k):
        return self.B.get(*k)

    def saved(self, k):
        src, s, r, c = k
        return self.sess.ov.value(s, r, c) if src == "" else self.sess.ext_cached.get((src, s, r, c))

    def kind(self, k) -> str:
        """formula / input (an overlay constant) / client (the client model's value, through the feed) /
        saved (another sheet of the overlay workbook, always read as saved)."""
        src, s, r, c = k
        if src != "":
            return "client" if src == self.w.get("client_link") else "saved"
        if s in self.B.overlay:
            return "formula" if self.B.is_formula(s, r, c) else "input"
        return "client" if s in self.sess.client_sheets else "saved"

    def reads(self, k) -> tuple[set, set]:
        key = (self.feed, k)
        if key not in self._reads:
            self._reads[key] = self.B.reads(k[1], k[2], k[3])
        return self._reads[key]

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
        return self.labels.get((s, r), "")

    # ---- the walk --------------------------------------------------------------------------------------------
    def walk(self, start, bad) -> tuple[list, int, bool]:
        """From start, follow the cells read while bad(cell) holds. A bad formula cell none of whose reads is
        bad (or that calls a function Python doesn't have) is a root; so is a bad input or client value.
        -> ([(root, unsupported functions it calls)], number of bad cells on the way, stopped at WALK_LIMIT:
        then the roots are the first ones reached, not necessarily all)."""
        memo = {}

        def is_bad(k):
            if k not in memo:
                try:
                    memo[k] = bool(bad(k))
                except Exception:
                    memo[k] = False
            return memo[k]

        seen, stack, roots, n = set(), [start], [], 0
        while stack and len(seen) < WALK_LIMIT:
            k = stack.pop()
            if k in seen:
                continue
            seen.add(k)
            if not is_bad(k):
                continue
            n += 1
            if self.kind(k) != "formula":
                roots.append((k, set()))
                continue
            reads, unknown = self.reads(k)
            kids = [x for x in reads if is_bad(x)]
            if not kids or unknown:
                roots.append((k, unknown))
            stack.extend(x for x in kids if x not in seen)
        return roots, n, bool(stack)

    # ---- a root's cause --------------------------------------------------------------------------------------
    def name_why(self, name: str, sheet: str) -> str:
        n = name.strip()
        if re.match(r"^'?\[\d+\]", n):
            return "a name defined in another workbook (through an external link), which isn't read"
        if "[" in n:
            return ("a reference into an Excel table (a structured reference like Table1[Column]); the Python "
                    "compiler doesn't read Excel tables yet")
        if n.lower().startswith("_xlpm."):
            return "a LET or LAMBDA parameter; the Python compiler doesn't support LET or LAMBDA yet"
        bare = n.split("!")[-1].lower()
        scopes = sorted(sc for (sc, nm) in self.ctx["local_names"] if nm == bare)
        if scopes:
            return f"defined only on sheet(s) {', '.join(scopes)}, not for {sheet} or the whole workbook"
        return "not defined in the workbook's list of names"

    def classify(self, k, unknown: set, feed: str) -> dict:
        src, s, r, c = k
        v, w = self.val(k), self.saved(k)
        kind = self.kind(k)
        out = {"cell": _a1k(k), "label": self.label(k), "kind": kind, "python": _show(v), "excel": _show(w)}
        if kind == "client":
            if feed == "current":
                ex = self.sess.rowmap.explain(s, r) if self.sess.rowmap else {"found": None}
                out["current_row"] = f"{ex['found'][0]}!r{ex['found'][1]}" if ex["found"] else None
                if (s, r, c) in self.sess.unmatched:
                    return {**out, "cause": "unmatched", "detail": self.sess.unmatched[(s, r, c)]
                            + ("; last year's value stands in" if (s, r, c) in self.sess.stood_in else "")}
            if _err(v):
                where = {"workbook": "the overlay's saved copy of the client value", "prior": "last year's client model",
                         "current": "this year's client model"}[feed]
                return {**out, "cause": "client_error", "detail": f"{where} holds {v.code} here as saved"}
            return {**out, "cause": "client_differs",
                    "detail": "last year's client model holds a different value here than the overlay last read"}
        if kind in ("input", "saved"):
            return {**out, "cause": "input_error" if _err(v) else "input_differs",
                    "detail": "a constant in the workbook that holds an error as saved" if _err(v) else
                              "a constant that differs from the saved value (a change or a hold on it)"}
        f = self.formula(k) or ""
        out["formula"] = f[:400]
        ex = xlcompile.explain(self.ctx, f, s, r, c) if f else {"not_compiled": None, "unknown_functions": [],
                                                                 "missing_names": [], "issues": []}
        fns = sorted(set(unknown) | set(ex["unknown_functions"]))
        if ex["not_compiled"]:
            return {**out, "cause": "not_compiled", "detail": ex["not_compiled"]}
        if fns:
            return {**out, "cause": "function", "functions": fns, "detail": self.function_why(fns)}
        if ex["missing_names"]:
            return {**out, "cause": "name", "names": ex["missing_names"],
                    "detail": "; ".join(f"{n}: {self.name_why(n, s)}" for n in ex["missing_names"][:4])}
        if _err(v) and _err(w) and v.code == w.code:
            return {**out, "cause": "excel_error", "detail": f"Excel shows {w.code} here too, in the file as saved"}
        reads, _ = self.reads(k)
        sample = []
        for x in sorted(reads, key=lambda x: (str(x[0]), x[1], x[2], x[3]))[:12]:
            sample.append({"cell": _a1k(x), "label": self.label(x), "python": _show(self.val(x)),
                           "excel": _show(self.saved(x))})
        out["reads"] = sample
        out["issues"] = ex["issues"]
        fn = self.B.rows.get((s, r))
        try:
            out["python_code"] = inspect.getsource(fn.raw)[:1800] if fn else None
        except (OSError, TypeError):
            out["python_code"] = None
        if feed == "workbook":
            return {**out, "cause": "computed_differently",
                    "detail": "Python works this formula out differently from Excel, from the same inputs"}
        return {**out, "cause": "error_on_new_values",
                "detail": f"the formula gives {_show(v)} on the {feed} feed's values (Excel's saved value is {_show(w)})"}

    def function_why(self, fns: list[str]) -> str:
        addin = [f for f in fns if ADDIN.match(f)]
        if addin:
            return (f"{', '.join(addin)}: a data provider's add-in function; the file only holds the values it last "
                    "fetched, so Python can't recompute it")
        return (f"{', '.join(fns)}: not in the Python runtime" +
                ("; the workbook has macros, so it may be a VBA function written for this model" if self.xlsm else ""))

    # ---- grouping --------------------------------------------------------------------------------------------
    @staticmethod
    def signature(root: dict) -> tuple:
        cause = root["cause"]
        s = root["cell"].split("!")[0]
        if cause == "function":
            return cause, ", ".join(root["functions"])
        if cause == "name":
            return cause, ", ".join(root["names"])
        if cause == "not_compiled":
            return cause, re.sub(r"\d+", "#", root["detail"])[:80]
        if cause in ("client_error", "client_differs", "unmatched", "input_error", "input_differs"):
            return cause, s
        return cause, re.sub(r"\d+$", "", root["cell"].rstrip("0123456789")) + "|" + root["label"]

    TITLES = {
        "not_compiled": "A formula Python couldn't compile",
        "function": "A function the Python runtime doesn't have",
        "name": "A defined name Python can't find",
        "excel_error": "Excel shows the same error in the saved file",
        "computed_differently": "Python works a formula out differently from Excel",
        "error_on_new_values": "A formula that errors on the new values",
        "client_error": "The client model holds an error",
        "client_differs": "Last year's client model differs from what the overlay last read",
        "unmatched": "A client value that couldn't be found in this year's model",
        "input_error": "A constant holding an error",
        "input_differs": "A constant that differs from the saved file",
    }

    def group(self, found: list[tuple[str, dict]]) -> list[dict]:
        """[(figure label, classified root)] -> groups by cause, the biggest first."""
        groups = {}
        for fig, root in found:
            sig = self.signature(root)
            root.setdefault("_order", _order(root["cell"]))
            g = groups.setdefault(sig, {"cause": root["cause"], "title": self.TITLES.get(root["cause"], root["cause"]),
                                        "detail": root["detail"], "cells": [], "figures": [], "example": root})
            if root["cell"] not in g["cells"]:
                g["cells"].append(root["cell"])
                g["cells"].sort(key=_order)
            if fig not in g["figures"]:
                g["figures"].append(fig)
        out = sorted(groups.values(), key=lambda g: (-len(g["figures"]), -len(g["cells"])))
        for g in out:
            g["example"].pop("_order", None)
            g["n_cells"] = len(g["cells"])
            g["all_cells"], g["cells"] = g["cells"], g["cells"][:12]
        return out

    # ---- the passes ------------------------------------------------------------------------------------------
    def figures(self) -> list[dict]:
        outs = [o for o in self.summary.get("outputs") or [] if o.get("report") or o.get("fact_id")] or \
            list(self.summary.get("outputs") or [])
        return [{"cell": o["cell"], "key": ("",) + parse_a1(o["cell"]), "label": o.get("label") or o["cell"],
                 "report": o.get("report"), "scale": o.get("scale") or 1.0, "sign": o.get("sign") or 1} for o in outs]

    def run(self) -> dict:
        w, sess, B = self.w, self.sess, self.B
        figs = self.figures()
        feeds = ["workbook"] + (["prior"] if w.get("prior") and not w.get("same_file") else []) + \
            (["current"] if w.get("current") else [])
        values, roots, logs, bad_counts, truncated = {}, {}, {}, {}, {}
        wb_memo = {}
        for i, feed in enumerate(feeds):
            self.progress(0.05 + 0.5 * i / len(feeds), f"Following each figure back on the {feed} feed")
            self.use(feed)
            B.feed_log = {}
            try:
                vals = [self.val(f["key"]) for f in figs]
                if feed == "workbook":
                    bad = lambda k: not same(self.val(k), self.saved(k))
                elif feed == "prior":
                    bad = lambda k: not same(self.val(k), wb_memo.get(k[1:], self.saved(k)) if self.kind(k) == "formula"
                                             else self.saved(k))
                else:
                    bad = lambda k: _err(self.val(k)) or (self.kind(k) == "client" and k[1:] in self.sess.unmatched)
                found = []
                n_bad = 0
                for f, v in zip(figs, vals):
                    if not bad(f["key"]):
                        continue
                    rs, n, cut = self.walk(f["key"], bad)
                    n_bad += n
                    if cut:
                        truncated.setdefault(feed, []).append(f["label"])
                    found += [(f["label"], self.classify(k, unk, feed)) for k, unk in rs]
                values[feed] = [_show(v) for v in vals]
                roots[feed] = self.group(found)
                bad_counts[feed] = n_bad
                logs[feed] = dict(B.feed_log)
                if feed == "workbook":
                    wb_memo = dict(B.memo)
                if feed == "current":
                    self.unmatched = dict(sess.unmatched)
            finally:
                B.feed_log = None
        first = {}
        for feed in feeds:
            keep = []
            for g in roots[feed]:
                k = (g["cause"], g["detail"], tuple(g["all_cells"]))
                if k in first:
                    first[k].setdefault("also_on", []).append(feed)
                else:
                    first[k] = g
                    keep.append(g)
            roots[feed] = keep
        self.progress(0.6, "Checking the files")
        files = self.files(logs.get("prior"))
        self.progress(0.7, "Checking the rows followed into this year's model")
        rows = self.rows(logs.get("prior") or logs.get("workbook"), getattr(self, "unmatched", {}))
        self.progress(0.8, "Checking which causes can be held at Excel's values")
        holds = self.holds(roots)
        for groups in roots.values():
            for g in groups:
                g.pop("all_cells", None)
        layers = self.layers(figs, values, feeds)
        v = self.summary.get("validation") or {}
        st = self.summary.get("stats") or {}
        whole = {"formula_cells": v.get("cells"), "matched_excel": v.get("matched"),
                 "unsupported_functions": v.get("unsupported") or {}, "missing_names": st.get("missing_names") or [],
                 "not_compiled": len(self.summary.get("not_compiled") or []),
                 "not_compiled_examples": (self.summary.get("not_compiled") or [])[:5],
                 "circular": v.get("cycles"), "held": len(sess.holds)}
        sess.configure("workbook")
        return {"feeds": feeds, "layers": layers, "roots": roots, "bad_cells": bad_counts, "truncated": truncated,
                "files": files,
                "rows": rows, "holds": holds, "whole_model": whole,
                "wiring": {k: (w.get(k) or {}).get("filename") for k in ("overlay", "prior", "current")}}

    def layers(self, figs, values, feeds) -> list[dict]:
        out = []
        for i, f in enumerate(figs):
            saved = self.saved(f["key"])
            row = {"cell": f["cell"], "label": f["label"], "report": f["report"], "saved": _show(saved),
                   **{feed: values[feed][i] for feed in feeds}}
            t = tie(saved, f["report"], f["scale"], f["sign"]) if f["report"] else None
            row["saved_ties"] = t["ok"] if t else None
            wbv = values["workbook"][i]
            breaks = None
            if t and not t["ok"]:
                breaks = "saved"
            elif not same(wbv, row["saved"]):
                breaks = "workbook"
            elif "prior" in feeds and not same(values["prior"][i], wbv):
                breaks = "prior"
            elif "current" in feeds and isinstance(values["current"][i], str) and values["current"][i].startswith("#"):
                breaks = "current"
            row["breaks_at"] = breaks
            out.append(row)
        return out

    def files(self, prior_log: dict | None) -> list[dict]:
        """Checks on the files, each {check, status: ok|warn|bad|info, detail, examples}."""
        w, sess = self.w, self.sess
        out = []
        have = {r[0] for r in self.db.execute("SELECT name FROM sqlite_master WHERE type='table'")}
        books = []
        if "extbooks" in have:
            books = [{"idx": i, "filename": fn, "target": t, "cached": n} for i, t, fn, _, n in
                     self.db.execute("SELECT idx, target, filename, sheets, n_cached FROM extbooks ORDER BY idx")]
        prior = w.get("prior")
        norm = lambda n: re.sub(r"[^a-z0-9]+", "", re.sub(r"\.(xlsx|xlsm|xlsb|xls)$", "", (n or "").lower()))
        if books:
            out.append({"check": "External links in the overlay", "status": "info",
                        "detail": "; ".join(f"[{b['idx']}] {b['filename'] or b['target'] or '?'} ({b['cached'] or 0} cached "
                                            f"value(s))" for b in books)})
        if prior and not w.get("same_file"):
            link = w.get("client_link")
            if not w.get("client_sheets") and link is None:
                out.append({"check": "The overlay reads last year's client model", "status": "bad",
                            "detail": f"{prior['filename']} is assigned as last year's client model, but none of the "
                                      "overlay's external links was matched to it (by name, by the values it last "
                                      "read or by its sheet names), so nothing reads it: 'Rebuilt on last year's "
                                      "model' is only the saved values again, and this year's model isn't read "
                                      "either. Check the file is the one the overlay links to."})
            link_check = None
            if link is not None:
                b = next((b for b in books if b["idx"] == link), None)
                named = b and norm(b["filename"]) == norm(prior["filename"])
                link_check = {"check": "The link the overlay reads is last year's client model", "status": "ok" if named else "warn",
                              "detail": f"link [{link}] points at {b['filename'] if b else '?'}; the file assigned is "
                                        f"{prior['filename']}" + ("" if named else " (a different name)")}
                out.append(link_check)
            if prior_log is not None and (link is not None or w.get("client_sheets")):
                reads = {k: v for k, v in prior_log.items() if self.kind(k) == "client"}
                diffs = [(k, v, self.saved(k)) for k, v in reads.items() if not same(v, self.saved(k))]
                errs = [(k, v) for k, v in reads.items() if _err(v)]
                n = len(reads)
                share = 1 - len(diffs) / n if n else 1.0
                status = "ok" if share >= 0.98 else "warn" if share >= 0.8 else "bad"
                by_row = Counter((k[1], k[2]) for k, _, _ in diffs)
                if not n:
                    detail = "the figures read no client values"
                else:
                    detail = (f"of the {n:,} client value(s) the figures read, {n - len(diffs):,} are the same in "
                              f"{prior['filename']} as the overlay last saw ({share:.0%})")
                    if status == "bad":
                        detail += (f"; the differences are on {len(by_row)} line item(s). Most values differ: this is "
                                   "probably not the version the overlay was saved with (check the file's date and name)")
                    elif status == "warn":
                        detail += (f"; the differences are on {len(by_row)} line item(s): the file was probably updated "
                                   "after the overlay was last saved")
                if link_check and link_check["status"] != "ok" and n:
                    link_check["detail"] = link_check["detail"][:-1] + (
                        ", matched by the values it last read)" if diffs and len(diffs) < n else
                        ", matched by its sheet names: none of the values it last read match)" if diffs else
                        ", matched by the values it last read)")
                out.append({"check": "Last year's client model is the version the overlay was saved with",
                            "status": status if n else "info", "detail": detail,
                            "examples": [{"cell": _a1k(k), "label": self.label(k), "overlay_saw": _show(sv),
                                          "file_has": _show(v)} for k, v, sv in diffs[:10]]})
                if errs:
                    out.append({"check": "Errors in last year's client model", "status": "bad",
                                "detail": f"{len(errs)} client value(s) the figures read hold an Excel error in "
                                          f"{prior['filename']} as saved (e.g. saved on a machine without an add-in, or "
                                          "with broken links): open it in Excel, recalculate and save, then upload again",
                                "examples": [{"cell": _a1k(k), "label": self.label(k), "file_has": _show(v)}
                                             for k, v in errs[:10]]})
        # the overlay's own saved errors
        errs = self.db.execute("SELECT sheet, row, col, value FROM cells WHERE formula IS NOT NULL AND typeof(value)='text' "
                               "AND value IN ('#NAME?','#REF!','#VALUE!','#DIV/0!','#N/A','#NUM!','#NULL!') AND sheet IN "
                               f"({','.join('?' * len(self.summary['sheets']))})", self.summary["sheets"]).fetchall()
        if errs:
            out.append({"check": "Errors in the overlay as saved", "status": "warn",
                        "detail": f"{len(errs):,} formula cell(s) on the overlay's sheets already show an Excel error in "
                                  "the saved file; Python reproducing them is correct",
                        "examples": [{"cell": _a1(s, r, c), "label": self.labels.get((s, r), ""), "file_has": v}
                                     for s, r, c, v in errs[:8]]})
        return out

    def rows(self, log: dict | None, unmatched: dict) -> dict | None:
        """Each client line item the figures read, followed into this year's model, with its history compared."""
        sess = self.sess
        if not sess.current or not sess.rowmap or not log:
            return None
        prior = sess.prior or sess.ov
        cur = sess.current
        pvd = (self.summary.get("roll") or {}).get("prior_valuation_date") or self.w.get("prior_valuation_date")
        pvd_s = serial(date.fromisoformat(pvd[:10])) if pvd else None
        items = sorted({(k[1], k[2]) for k in log if self.kind(k) == "client"})
        checked, flagged, n_ok = 0, [], 0
        cur_labels = cur.labels()
        for s, r in items:
            ex = sess.rowmap.explain(s, r)
            lab = prior.labels().get((s, r), "")
            if ex["found"] is None:
                flagged.append({"item": f"{s}!r{r}", "label": lab, "status": "missing",
                                "detail": sess.rowmap.why(s, r, prior.labels())})
                continue
            s2, r2 = ex["found"]
            how = "; ".join(f"{n}: {t}" for n, t in ex["evidence"][:2]) or ex["how"]
            lab2 = cur_labels.get((s2, r2), "")
            hist, fut = [], []
            for c, when in sorted(prior.timeline(s).items()):
                c2 = cur.column_of(s2, when)
                a, b = prior.value(s, r, c), cur.value(s2, r2, c2) if c2 else None
                if isinstance(a, float) and isinstance(b, float) and (a or b):
                    (hist if pvd_s is not None and when <= pvd_s else fut).append((to_date(when).isoformat(), a, b))
            checked += 1
            item = {"item": f"{s}!r{r} -> {s2 + '!' if s2 != s else ''}r{r2}", "label": lab, "now": lab2, "matched_by": how}
            pairs = hist or fut
            if not pairs:
                if "same row" in how and lab and lab2 and lab.strip().lower() != lab2.strip().lower():
                    flagged.append({**item, "status": "check",
                                    "detail": f"matched by position ({how}) but the label reads '{lab2}' now"})
                else:
                    n_ok += 1
                continue
            close = lambda p: abs(p[1] - p[2]) <= 0.005 * max(1.0, abs(p[1]), abs(p[2]))
            ratios = sorted(p[2] / p[1] for p in pairs if p[1])
            mid = ratios[len(ratios) // 2] if ratios else None
            flips = sum(p[1] * p[2] < 0 for p in pairs) / len(pairs)
            if hist and sum(map(close, hist)) >= 0.8 * len(hist):
                n_ok += 1
                continue
            if mid and any(abs(abs(mid) - m) <= 0.01 * m for m in (1000, 0.001, 1e6, 1e-6, 100, 0.01)):
                status, why = "units", f"this year's values are {mid:g} times last year's for the same periods: the units changed"
            elif flips >= 0.6:
                status, why = "sign", f"{flips:.0%} of the same periods have the opposite sign this year"
            elif hist:
                status, why = "history differs", (f"{sum(not close(p) for p in hist)} of {len(hist)} period(s) up to last "
                                                  "year's valuation date differ: a different line item, or history restated")
            elif mid is not None and not (1 / 3 <= abs(mid) <= 3):
                status, why = "forecast differs", (f"no history to compare; this year's forecast for the same "
                                                   f"{len(fut)} period(s) is {abs(mid):.2g} times last year's")
            else:
                n_ok += 1
                continue
            flagged.append({**item, "status": status, "detail": why,
                            "periods": [{"period": d, "last_year": a, "this_year": b} for d, a, b in pairs
                                        if not close((d, a, b))][:4]})
        um = Counter((s, r) for (s, r, c) in unmatched)
        return {"items": len(items), "checked": checked, "ok": n_ok, "flagged": flagged[:60], "n_flagged": len(flagged),
                "unmatched_values": len(unmatched),
                "unmatched_items": [{"item": f"{s}!r{r}", "label": prior.labels().get((s, r), ""), "values": n,
                                     "detail": unmatched.get(next(k for k in unmatched if k[:2] == (s, r)))}
                                    for (s, r), n in um.most_common(20)]}

    def holds(self, roots: dict) -> dict:
        """Roots Python can't compute (a function, a name, a formula it can't compile) whose value can't change
        between feeds: nothing from the client model or an assumption goes into them. Holding them at Excel's
        saved value is then right on every feed; checked by recomputing the figures with them held."""
        sess = self.sess
        cand, seen = [], set()
        for feed, groups in roots.items():
            for g in groups:
                if g["cause"] not in ("function", "name", "not_compiled", "computed_differently"):
                    continue
                for cell in g["all_cells"]:
                    k = ("",) + parse_a1(cell)
                    if k not in seen:
                        seen.add(k)
                        cand.append({"key": k, "cell": cell, "cause": g["cause"], "title": g["title"]})
        if not cand:
            return {"safe": [], "unsafe": [], "check": None}
        self.use("workbook")
        _, _, months = _feed(self.summary, "current", None, None)
        moving = set(self.levers) | (set(sess.rolled_timeline(months)) if self.w.get("current") else set())
        safe, unsafe, skipped = [], [], 0
        t0 = time.time()
        for i, c in enumerate(cand):
            if i >= HOLD_CANDIDATES or time.time() - t0 > HOLD_SECONDS:
                skipped = len(cand) - i
                break
            k = c["key"]
            w = self.saved(k)
            if w is None or _err(w):
                unsafe.append({**c, "why": f"Excel's saved value is {_show(w)}, so there's nothing to hold it at"})
                continue
            why = self.depends(k, moving)
            if why:
                unsafe.append({**c, "why": why})
            else:
                safe.append({**c, "value": w, "excel": _show(w)})
        check = None
        if safe:
            check = self.try_holds({x["key"][1:]: x["value"] for x in safe})
        for x in safe + unsafe:
            x.pop("key", None)
        return {"safe": safe, "unsafe": unsafe[:30], "n_unsafe": len(unsafe), "not_checked": skipped, "check": check}

    def depends(self, k, moving: set) -> str | None:
        """Why a cell's value can change between feeds (it reads the client model or an assumption), or None.
        Remembered: a closure walked to the end with nothing found is clean for every cell in it."""
        memo = self._depends
        if k in memo:
            return memo[k]
        seen, stack = set(), [k]
        while stack:
            if len(seen) > CLOSURE_LIMIT:
                memo[k] = "it reads too many cells to check"
                return memo[k]
            x = stack.pop()
            if x in seen:
                continue
            seen.add(x)
            if x in memo:
                if memo[x]:
                    memo[k] = memo[x]
                    return memo[k]
                continue
            kind = self.kind(x)
            why = None
            if kind == "client":
                why = f"it reads the client model ({_a1k(x)} {self.label(x)})".strip()
            elif x[0] == "" and x[1:] in moving:
                lev = self.levers.get(x[1:])
                why = f"it reads {'the assumption ' + lev['label'] if lev else 'the timeline'} ({_a1k(x)}), which changes between feeds"
            if why:
                memo[k] = why
                return why
            if kind == "formula":
                stack.extend(self.reads(x)[0])
        for x in seen:
            memo[x] = None
        return None

    def try_holds(self, held: dict) -> dict:
        """The figures on every feed with these cells held, against without."""
        sess = self.sess
        figs = self.figures()
        keep = dict(sess.holds)
        out = {}
        feeds = ["workbook"] + (["prior"] if self.w.get("prior") and not self.w.get("same_file") else []) + \
            (["current"] if self.w.get("current") else [])
        try:
            for feed in feeds:
                sess.holds = keep
                self.use(feed)
                before = [self.val(f["key"]) for f in figs]
                sess.holds = {**keep, **held}
                self.use(feed)
                after = [self.val(f["key"]) for f in figs]
                out[feed] = {"errors_before": sum(map(_err, before)), "errors_after": sum(map(_err, after)),
                             "saved_matched_after": sum(same(a, self.saved(f["key"])) for a, f in zip(after, figs))
                             if feed == "workbook" else None}
        finally:
            sess.holds = keep
            self.use("workbook")
        return out


def _order(cell: str) -> tuple:
    """'[1]Sheet!AB12' -> a key that sorts by sheet, row, then column."""
    m = re.match(r"^(.*)!\$?([A-Z]{1,3})\$?(\d+)$", cell)
    if not m:
        return (cell, 0, 0)
    col = 0
    for ch in m.group(2):
        col = col * 26 + ord(ch) - 64
    return (m.group(1), int(m.group(3)), col)


def examine(sess, summary: dict, progress=None) -> dict:
    """The evidence (see the module's doc). Runs in overlay.deep (deep formula chains)."""
    d = Doctor(sess, summary, progress)
    try:
        return d.run()
    finally:
        try:
            sess.configure("workbook")
        finally:
            d.db.close()


# ---- the write-up ----------------------------------------------------------------------------------------------

DIAGNOSE_PROMPT = """You are the doctor for a valuation overlay that has been rebuilt in Python. The app recomputes
last year's valuation (an Excel overlay workbook reading a client's financial model) in Python, checks it against
the values Excel saved and the figures in last year's report, then feeds it this year's client model.

Below is the evidence the app gathered. Work only from it: don't invent causes, cells or numbers. If the evidence
doesn't settle something, say what to check and how.

How to read it:
- layers: each report figure through: report -> Excel's saved value in the overlay ("saved") -> Python on the
  overlay's saved inputs ("workbook") -> Python on last year's client model ("prior") -> Python on this year's
  client model, rolled forward ("current"). breaks_at is the first layer where it stops agreeing with the one
  before. "workbook" breaking means the Python compiler or runtime (no client file is read there, so it can't be
  the wrong file or the wrong row). "prior" breaking means the client model file. "current" means this year's
  model or the rows followed into it. "saved" means the overlay file or the cell picked for the figure.
- roots: per feed, the cells where the wrong values start, grouped by cause, with the figures they reach.
  Causes: not_compiled, function (not in the Python runtime; add-in or VBA), name (a defined name Python can't
  find: table references, LET parameters, names on another sheet), computed_differently (same inputs, different
  answer: a runtime difference), client_error / client_differs (the client file), unmatched (a client line item
  not found this year), excel_error (Excel shows it too).
- files: checks on which files are read. rows: client line items followed into this year's model, with their
  history compared. A row flagged "history differs", "sign" or "units" is most likely the wrong line item picked
  this year (or its sign or units changed): say so plainly and name it. "forecast differs" is weaker: a check.
  "missing" means the line item wasn't found this year, so its values count as blank. holds: roots that can be held at Excel's saved value on every feed (safe) or not (why).
- truncated: per feed, figures whose walk stopped at its cell limit: the roots found are the first ones reached,
  not necessarily all of them, and bad_cells is a floor. Say so if it matters to a finding.
- whole_model: counts for all formula cells, not only the figures'.

Reply with:
- headline: one or two sentences: what is wrong and where, in plain words for a valuer.
- findings: the causes, most important first (at most 6). For each: title; what_we_saw (the evidence, with
  cells and numbers); cause; fix (concrete steps); who: "app" (the doctor can apply it here: only holds listed as
  safe), "you" (the valuer: re-upload, re-save, reassign a file, check a row) or "developer" (the Python
  compiler/runtime needs a change: say exactly what, e.g. "support Excel tables (structured references)");
  confidence: high, medium or low.
- wrong_file: your answer to "are we reading the wrong file?" from the files and layers evidence, one sentence.
- wrong_row: your answer to "are we picking up the wrong rows?" from the rows evidence, one sentence.

Evidence:
{evidence}"""

_DIAGNOSIS = {"type": "json_schema", "name": "overlay_diagnosis", "strict": True, "schema": {
    "type": "object", "additionalProperties": False, "required": ["headline", "findings", "wrong_file", "wrong_row"],
    "properties": {
        "headline": {"type": "string"},
        "findings": {"type": "array", "items": {
            "type": "object", "additionalProperties": False,
            "required": ["title", "what_we_saw", "cause", "fix", "who", "confidence"],
            "properties": {"title": {"type": "string"}, "what_we_saw": {"type": "string"}, "cause": {"type": "string"},
                           "fix": {"type": "string"}, "who": {"type": "string", "enum": ["app", "you", "developer"]},
                           "confidence": {"type": "string", "enum": ["high", "medium", "low"]}}}},
        "wrong_file": {"type": "string"}, "wrong_row": {"type": "string"}}}}


def _round(x):
    """Numbers at six significant figures: the write-up quotes them."""
    if isinstance(x, float):
        return float(f"{x:.6g}")
    if isinstance(x, dict):
        return {k: _round(v) for k, v in x.items()}
    if isinstance(x, list):
        return [_round(v) for v in x]
    return x


def _trim(ev: dict, limit: int = 60000) -> str:
    """The evidence as JSON for the prompt, the bulkiest parts cut first."""
    ev = _round(json.loads(json.dumps(ev, default=str)))
    text = json.dumps(ev, indent=1)
    for feed, groups in (ev.get("roots") or {}).items():
        if len(text) <= limit:
            break
        for g in groups:
            g["example"].pop("python_code", None)
        del groups[12:]
        text = json.dumps(ev, indent=1)
    if len(text) > limit and ev.get("rows"):
        ev["rows"]["flagged"] = ev["rows"]["flagged"][:15]
        text = json.dumps(ev, indent=1)
    return text[:limit]


def diagnose(reader, evidence: dict) -> dict:
    """The reviewer model's write-up of the evidence (docingest.Reader: text only, JSON reply)."""
    return reader._call(reader.reviewer_model, DIAGNOSE_PROMPT.format(evidence=_trim(evidence)), None, _DIAGNOSIS,
                        "overlay-doctor")


# ---- the text to paste ---------------------------------------------------------------------------------------

def _fmt(v) -> str:
    if isinstance(v, float):
        return f"{v:,.2f}" if abs(v) >= 1 else f"{v:.6g}"
    return str(v)


def report_text(res: dict) -> str:
    """The diagnosis and its evidence as plain text, to paste into a message or a ticket."""
    ev, dx = res.get("evidence") or {}, res.get("diagnosis") or {}
    L = ["OVERLAY DOCTOR", ""]
    wi = ev.get("wiring") or {}
    L.append("Files: overlay " + str(wi.get("overlay")) + "; last year's client model " + str(wi.get("prior")) +
             "; this year's " + str(wi.get("current")))
    if dx:
        L += ["", "Diagnosis: " + dx.get("headline", "")]
        for i, f in enumerate(dx.get("findings") or [], 1):
            L += [f"{i}. {f['title']} ({f['who']}, {f['confidence']})", f"   saw: {f['what_we_saw']}",
                  f"   cause: {f['cause']}", f"   fix: {f['fix']}"]
        L += [f"Wrong file? {dx.get('wrong_file', '')}", f"Wrong rows? {dx.get('wrong_row', '')}"]
    L += ["", "Figures (report | Excel saved | Python: " + " | ".join(ev.get("feeds") or []) + " | breaks at)"]
    for r in ev.get("layers") or []:
        L.append(f"- {r['label']} {r['cell']}: {r.get('report')} | {_fmt(r['saved'])} | " +
                 " | ".join(_fmt(r.get(f)) for f in ev.get("feeds") or []) + f" | {r.get('breaks_at') or '-'}")
    for feed, groups in (ev.get("roots") or {}).items():
        if not groups:
            continue
        cut = (ev.get("truncated") or {}).get(feed)
        L += ["", f"Where it starts, {feed} feed ({ev.get('bad_cells', {}).get(feed, 0):,} wrong cell(s) on the way"
                  + (f"; stopped at {WALK_LIMIT:,} cells for {', '.join(cut)}: the first causes reached, maybe not all"
                     if cut else "") + "):"]
        for g in groups[:10]:
            ex = g["example"]
            L.append(f"- {g['title']}: {g['detail']}")
            L.append(f"  {g['n_cells']} cell(s), e.g. {', '.join(g['cells'][:5])}; reaches {', '.join(g['figures'][:4])}"
                     + (f"; the same on the {' and '.join(g['also_on'])} feed" if g.get("also_on") else ""))
            if ex.get("formula"):
                L.append(f"  {ex['cell']} ({ex.get('label', '')}) = {ex['formula'][:300]}")
                L.append(f"  Python {_fmt(ex['python'])}, Excel {_fmt(ex['excel'])}")
    fl = ev.get("files") or []
    if fl:
        L += ["", "Files:"]
        for f in fl:
            L.append(f"- [{f['status']}] {f['check']}: {f['detail']}")
            for x in (f.get("examples") or [])[:4]:
                L.append("    " + ", ".join(f"{k} {_fmt(v)}" for k, v in x.items()))
    rows = ev.get("rows")
    if rows:
        L += ["", f"Rows: {rows['items']} client line item(s) read; {rows['ok']} look right in this year's model; "
                  f"{rows['n_flagged']} flagged; {rows['unmatched_values']} value(s) unmatched this year"]
        for x in rows["flagged"][:12]:
            L.append(f"- [{x['status']}] {x['item']} {x.get('label', '')}: {x['detail']}")
        for x in rows["unmatched_items"][:8]:
            L.append(f"- [unmatched] {x['item']} {x['label']}: {x['detail']} ({x['values']} value(s))")
    wm = ev.get("whole_model") or {}
    if wm:
        L += ["", f"Whole model: {wm.get('matched_excel')}/{wm.get('formula_cells')} formula cells match Excel; "
                  f"functions not supported: {', '.join(f'{k} ({v})' for k, v in (wm.get('unsupported_functions') or {}).items()) or 'none'}; "
                  f"names not found: {', '.join((wm.get('missing_names') or [])[:15]) or 'none'}; "
                  f"formulas not compiled: {wm.get('not_compiled')}"]
        for x in wm.get("not_compiled_examples") or []:
            L.append(f"- {x['cell']}: {x['reason']} :: {x['formula'][:160]}")
    h = ev.get("holds") or {}
    if h.get("safe") or h.get("unsafe"):
        L += ["", f"Holds: {len(h.get('safe') or [])} cell(s) safe to hold at Excel's value; "
                  f"{h.get('n_unsafe', len(h.get('unsafe') or []))} not safe"
                  + (f"; {h['not_checked']} not checked (too many to check in time)" if h.get("not_checked") else "")]
        if h.get("check"):
            L.append("  with them held: " + "; ".join(f"{k} errors {v['errors_before']} -> {v['errors_after']}"
                                                     for k, v in h["check"].items()))
        for x in (h.get("unsafe") or [])[:3]:
            L.append(f"  not safe: {x['cell']}: {x['why']}")
    return "\n".join(L)
