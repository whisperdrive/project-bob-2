"""How a figure the report states is built in a workbook, traced down to its discounting.

The Map step matched the report's conclusions to overlay cells; those that tie to the report are the starting
points. From one, trace() follows the cells each formula reads, level by level, until it reaches the discounting:
  SUMPRODUCT(cash flows, factors)   the factors a row of the workbook, or computed in the formula
                                    (e.g. 1/(1+rate)^((dates-valuation date)/365))
  XNPV(rate, cash flows, dates)     actual/365 from the first date, whose cash flow isn't discounted
  NPV(rate, cash flows)             one period per column, end of period
  SUM(a row of present values)      each column a cash flow times a factor
Each discounting found (a "core") is recomputed here from its own cash flows and factors, and where its rate,
valuation date and convention can be read back from the factors, it gets the inputs dcf.compute() takes, so the
Valuation tab can validate it and rerun it under another method. Its cash-flow row is traced one step further
into the rows it adds up. Every cell on the way keeps its formula written in line-item words, so the tree reads
as an explanation: equity value (ex-div) = cum-div less the distribution; cum-div = the mid value; mid = the
average of the PVs at the low and the high rate.
Rows (long ranges) are operands, not followed cell by cell; single cells and short ranges are followed.
"""
import math
import os
import re
import threading
from datetime import date, timedelta

import numpy as np

import dcf
import valuation

MAX_NODES = 400
MAX_DEPTH = 16
SHORT = 12  # a range of at most this many cells is followed cell by cell; longer single rows are operands
_STR = re.compile(r'"[^"]*"')
_CALL = re.compile(r"(?<![A-Za-z0-9_.!$'])((?:_xlfn\.)?[A-Za-z][A-Za-z0-9.]*)\(")
_TOK = re.compile(r"""\s*(?:
    (?P<ref>(?:'[^']+'|[A-Za-z_][\w.]*)!\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?|\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)(?![\w(])
  | (?P<num>\d+(?:\.\d+)?(?:[eE][+-]?\d+)?%?)
  | (?P<fn>[A-Za-z][A-Za-z0-9.]*)\(
  | (?P<op>[-+*/^(),])
)""", re.X)
EXCEL_EPOCH = date(1899, 12, 30)


def _a1(sheet: str, row: int, col: int) -> str:
    return f"{sheet}!{dcf._addr(col, row)}"


def _cell(db, sheet, row, col):
    r = db.execute("SELECT formula, value FROM cells WHERE sheet=? AND row=? AND col=?", (sheet, row, col)).fetchone()
    return (r[0], r[1]) if r else (None, None)


def _label(db, sheet: str, row: int) -> str:
    r = db.execute("SELECT label FROM rows WHERE sheet=? AND row=?", (sheet, row)).fetchone()
    return (r[0] or "").strip() if r else ""


def _names(db) -> dict[str, str]:
    try:
        return {n.lower(): ref for n, ref in db.execute("SELECT name, ref FROM names") if ref and "!" in ref}
    except Exception:
        return {}


def _expand(db, formula: str, names: dict) -> str:
    """Named ranges replaced by what they refer to, so every reference reads as Sheet!A1."""
    if not names:
        return formula
    return re.sub(r"(?<![\w.!$'])([A-Za-z_][\w.]*)(?![\w(!])",
                  lambda m: names.get(m[1].lower(), m[1]).lstrip("="), formula)


def _row(db, text: str, here: str):
    """A single-row range long enough to be an operand -> (sheet, row, c1, c2), else None."""
    r = dcf._ref(text.strip(), here)
    if r and r[1] == r[3] and r[4] - r[2] + 1 > 2:
        return r[0], r[1], r[2], r[4]
    return None


def _calls(formula: str) -> list[tuple[str, list[str], str]]:
    """Every function call in a formula, nested ones too: (NAME, [argument texts], the whole call's text)."""
    out = []
    for m in _CALL.finditer(formula):
        depth, i, quote = 1, m.end(), False
        while i < len(formula) and depth:
            ch = formula[i]
            if ch == '"':
                quote = not quote
            elif not quote:
                depth += (ch == "(") - (ch == ")")
            i += 1
        inner = formula[m.end():i - 1]
        out.append((m[1].split(".")[-1].upper(), [a.strip() for _, a in valuation._split(inner, ",")] if inner.strip() else [],
                    formula[m.start():i]))
    return out


# ---- evaluating a factor expression, column by column -------------------------------------------------------

def _num(v) -> float:
    d = dcf._as_date(v) if isinstance(v, str) else None
    if d:
        return float((d - EXCEL_EPOCH).days)
    n = dcf._num(v)
    if n is None:
        raise ValueError(f"not a number: {v!r}")
    return float(n)


def _yearfrac(a, b, basis=0):
    basis = int(basis)
    if basis not in (1, 3):
        raise ValueError("YEARFRAC basis other than 1 (actual/actual) or 3 (actual/365)")
    dc = "actual/actual" if basis == 1 else "actual/365"
    a, b = np.broadcast_arrays(np.asarray(a, dtype=float), np.asarray(b, dtype=float))
    out = [dcf.yearfrac(EXCEL_EPOCH + timedelta(days=int(x)), EXCEL_EPOCH + timedelta(days=int(y)), dc) * (1 if y >= x else -1)
           for x, y in zip(a.ravel(), b.ravel())]
    return np.array(out).reshape(a.shape) if a.shape else out[0]


_flat = lambda args: np.concatenate([np.atleast_1d(np.asarray(x, dtype=float)).ravel() for x in args])
_FUNCS = {"YEARFRAC": _yearfrac, "EXP": np.exp, "LN": np.log, "ABS": np.abs,
          "SUM": lambda *a: float(np.sum(_flat(a))), "AVERAGE": lambda *a: float(np.mean(_flat(a))),
          "MIN": lambda *a: float(np.min(_flat(a))), "MAX": lambda *a: float(np.max(_flat(a))),
          "ROUND": lambda x, n=0: float(np.round(x, int(n)))}


def evaluate(db, expr: str, here: str, given: dict | None = None):
    """An arithmetic expression over cells: a number, or an array when it reads a row (one value per column of
    that row, in order). Dates count as Excel serial numbers. Only + - * / ^, brackets and a few functions
    (YEARFRAC, EXP, LN, ABS, SUM, AVERAGE, MIN, MAX, ROUND); anything else raises ValueError. given: values to use
    for some cells instead of the workbook's ({(sheet, row, col): value})."""
    given = given or {}
    code, vals, pos = [], [], 0
    text = expr.strip().lstrip("=")
    while pos < len(text):
        m = _TOK.match(text, pos)
        if not m or m.end() == pos:
            raise ValueError(f"can't read {text[pos:pos + 20]!r}")
        pos = m.end()
        if m["ref"]:
            r = dcf._ref(m["ref"], here)
            if not r:
                raise ValueError(m["ref"])
            sheet, r1, c1, r2, c2 = r
            one = lambda rr, cc: given[(sheet, rr, cc)] if (sheet, rr, cc) in given else (
                _num(v) if (v := _cell(db, sheet, rr, cc)[1]) is not None else 0.0)
            if r1 == r2 and c1 == c2:
                vals.append(float(one(r1, c1)))
            elif r1 == r2:
                got = dict(db.execute("SELECT col, value FROM cells WHERE sheet=? AND row=? AND col BETWEEN ? AND ?",
                                      (sheet, r1, c1, c2)))
                vals.append(np.array([given.get((sheet, r1, c), _num(got.get(c)) if got.get(c) is not None else 0.0)
                                      for c in range(c1, c2 + 1)]))
            elif (r2 - r1 + 1) * (c2 - c1 + 1) <= SHORT:  # a short block (AVERAGE(F10:F11)): its values in order
                vals.append(np.array([one(rr, cc) for rr in range(r1, r2 + 1) for cc in range(c1, c2 + 1)]))
            else:
                raise ValueError("a block of several rows")
            code.append(f"_v[{len(vals) - 1}]")
        elif m["num"]:
            code.append(f"({float(m['num'].rstrip('%')) / (100 if m['num'].endswith('%') else 1)!r})")
        elif m["fn"]:
            fn = m["fn"].split(".")[-1].upper()
            if fn not in _FUNCS:
                raise ValueError(f"function {fn}")
            code.append(f"_f['{fn}'](")
        else:
            code.append("**" if m["op"] == "^" else m["op"])
    return eval("".join(code), {"__builtins__": {}}, {"_v": vals, "_f": _FUNCS})  # noqa: S307 (built from tokens only)


# ---- the discounting ----------------------------------------------------------------------------------------

def _values(db, sheet, row, c1, c2) -> dict[int, float]:
    return {c: dcf._num(v) or 0.0 for c, v in db.execute(
        "SELECT col, value FROM cells WHERE sheet=? AND row=? AND col BETWEEN ? AND ?", (sheet, row, c1, c2))}


def _range(sheet, row, c1, c2) -> str:
    return f"{sheet}!{dcf._addr(c1, row)}:{dcf._addr(c2, row)}"


def _method(db, factors: dict[int, float], sheet: str, cols: list[int], dates: str | None = None) -> dict | None:
    """Rate, valuation date, convention and cut-off that reproduce these factors (valuation.read_factors)."""
    try:
        return valuation.read_factors(db, None, cols, theirs=factors, sheet=sheet, dates=dates)
    except (ValueError, ZeroDivisionError):
        return None


def _core(db, sheet: str, row: int, col: int, call: tuple, whole: bool) -> dict | None:
    """One discounting call in a formula -> its core, or None if the call isn't one."""
    fn, args, text = call
    here = sheet
    core = {"cell": _a1(sheet, row, col), "call": text, "whole": whole}
    try:
        if fn == "SUMPRODUCT" and len(args) == 2:
            a, b = _row(db, args[0], here), _row(db, args[1], here)
            if a and b:
                ra, rb = (a[0], a[1], a[2], a[1], a[3]), (b[0], b[1], b[2], b[1], b[3])
                if valuation._is_df(db, ra) and not valuation._is_df(db, rb):
                    a, b = b, a
                elif not valuation._is_df(db, rb):
                    return None
                cf, fac_row = a, b
                fv = list(_values(db, *fac_row).values())
                fac = {c: fv[i] if i < len(fv) else 0.0 for i, c in enumerate(range(cf[2], cf[3] + 1))}
                core.update(kind="sumproduct", what="SUMPRODUCT of a cash-flow row and a discount-factor row",
                            factor_row=_range(*fac_row))
            elif a or b:
                cf, expr = (a, args[1]) if a else (b, args[0])
                vec = np.atleast_1d(evaluate(db, expr, here))
                cols = list(range(cf[2], cf[3] + 1))
                if len(vec) != len(cols) or not all(0 <= x <= 1.0000001 for x in vec if x):
                    return None
                fac = dict(zip(cols, map(float, vec)))
                core.update(kind="sumproduct", what="SUMPRODUCT of a cash-flow row and factors computed in the formula",
                            factor_expression=expr)
            else:
                return None
            flows = _values(db, *cf)
            pv = sum(flows.get(c, 0.0) * f for c, f in fac.items())
            cols = list(range(cf[2], cf[3] + 1))
            m = _method(db, fac, cf[0], cols)
            core.update(cashflow=_range(*cf), pv=pv, factors=fac, method=m)
            if m:
                core["inputs"] = {"cashflow": [_range(*cf)], "rate": m["rate"], "valuation_date": m["valuation_date"],
                                  "timing": m["timing"], "day_count": m["day_count"], "terminal_date": m["terminal_date"],
                                  "adjustments": [], "compare_to": core["cell"] if whole else None}
            return core
        if fn == "XNPV" and len(args) == 3:
            vr, dr = _row(db, args[1], here), _row(db, args[2], here)
            if not vr or not dr:
                return None
            rate = float(evaluate(db, args[0], here))
            flows = _values(db, *vr)
            dv = dict(db.execute("SELECT col, value FROM cells WHERE sheet=? AND row=? AND col BETWEEN ? AND ?",
                                 (dr[0], dr[1], dr[2], dr[3])))
            vcols, dcols = list(range(vr[2], vr[3] + 1)), list(range(dr[2], dr[3] + 1))
            dates = [dcf._as_date(dv.get(c)) for c in dcols]
            if len(vcols) != len(dcols) or not dates[0]:
                return None
            d0 = dates[0]
            fac = {c: (1 + rate) ** (-((d - d0).days / 365)) if d else 0.0 for c, d in zip(vcols, dates)}
            pv = sum(flows.get(c, 0.0) * f for c, f in fac.items())
            rate_ref = args[0].lstrip("+") if dcf._ref(args[0], here) else rate
            if isinstance(rate_ref, str) and "!" not in rate_ref:
                rate_ref = f"{here}!{rate_ref}"
            first = flows.get(vcols[0], 0.0)
            core.update(kind="xnpv", what="XNPV: actual/365 from the first date", cashflow=_range(*vr), dates=_range(*dr),
                        rate=rate, valuation_date=d0.isoformat(), pv=pv, factors=fac,
                        method={"rate": rate_ref, "valuation_date": _a1(dr[0], dr[1], dr[2]), "timing": "end",
                                "day_count": "actual/365", "terminal_date": None})
            core["inputs"] = {"cashflow": [_range(*vr)], "dates": _range(*dr), "rate": rate_ref,
                              "valuation_date": _a1(dr[0], dr[1], dr[2]), "timing": "end", "day_count": "actual/365",
                              "terminal_date": None, "compare_to": core["cell"] if whole else None,
                              "adjustments": [{"label": "First cash flow (on the XNPV start date, not discounted)",
                                               "value": _a1(vr[0], vr[1], vr[2])}] if first else []}
            return core
        if fn == "NPV" and len(args) >= 2:
            vr = _row(db, args[1], here)
            if not vr or len(args) > 2:
                return None
            rate = float(evaluate(db, args[0], here))
            flows = _values(db, *vr)
            fac = {c: (1 + rate) ** -(i + 1) for i, c in enumerate(range(vr[2], vr[3] + 1))}
            core.update(kind="npv", what="NPV: one period per column, end of period (no dates)", cashflow=_range(*vr),
                        rate=rate, pv=sum(flows.get(c, 0.0) * f for c, f in fac.items()), factors=fac, method=None,
                        note="NPV counts columns, not dates, so the valuation date and convention can't be changed")
            return core
        if fn == "SUM" and len(args) == 1:
            pr = _row(db, args[0], here)
            if not pr:
                return None
            # a present-value row: each column's formula multiplies a cash flow by a factor (or divides it)
            f0 = next((f for c, f in db.execute("SELECT col, formula FROM cells WHERE sheet=? AND row=? AND col BETWEEN ? "
                                                "AND ? AND formula IS NOT NULL ORDER BY col", (pr[0], pr[1], pr[2], pr[3]))), None)
            if not f0 or not re.search(r"[*/]", f0):
                return None
            refs = [dcf._ref(m[0], pr[0]) for m in dcf._FREF.finditer(_STR.sub("", f0))]
            same_col = [(r[0], r[1]) for r in refs if r and r[1] == r[3] and r[2] == r[4] and r[1] != pr[1] and r[0] == pr[0]]
            pvv = _values(db, *pr)
            best = None
            for s_, r_ in dict.fromkeys(same_col):
                if valuation._is_df(db, (s_, r_, pr[2], r_, pr[3])):
                    continue
                cf = _values(db, s_, r_, pr[2], pr[3])
                fac = {c: pvv.get(c, 0.0) / cf[c] for c in cf if cf.get(c)}
                if fac and all(0 < f <= 1.0000001 for f in fac.values()):
                    best = ((s_, r_, pr[2], pr[3]), cf, fac)
                    break
            if not best:
                return None
            cf_row, flows, fac = best
            cols = list(range(pr[2], pr[3] + 1))
            m = _method(db, fac, pr[0], cols)  # the factors seen: a period with no cash flow shows none
            core.update(kind="pv row", what="SUM of a present-value row (each column a cash flow times a factor)",
                        cashflow=_range(*cf_row), pv_row=_range(*pr), pv=sum(pvv.values()), factors=fac, method=m)
            if m:
                core["inputs"] = {"cashflow": [_range(*cf_row)], "rate": m["rate"], "valuation_date": m["valuation_date"],
                                  "timing": m["timing"], "day_count": m["day_count"], "terminal_date": m["terminal_date"],
                                  "adjustments": [], "compare_to": core["cell"] if whole else None}
            return core
    except (ValueError, ZeroDivisionError, OverflowError, TypeError, FloatingPointError):
        return None
    return None


def _describe(db, core: dict) -> None:
    """Words, periods and the cash flow's make-up for a core."""
    sheet, row, cols = dcf._row_range(db, core["cashflow"])
    core["cashflow_label"] = _label(db, sheet, row)
    flows = _values(db, sheet, row, cols[0], cols[-1])
    live = [c for c in cols if core["factors"].get(c) and flows.get(c)]
    core["periods"] = len(live)
    core["undiscounted"] = sum(flows.get(c, 0.0) for c in live)
    try:
        ends, _ = dcf.period_ends(db, sheet, cols, core.get("dates"))
        core["first_period"] = ends[live[0]].isoformat() if live and live[0] in ends else None
        core["last_period"] = ends[live[-1]].isoformat() if live and live[-1] in ends else None
    except ValueError:
        core["first_period"] = core["last_period"] = None
    core["parts"] = parts(db, sheet, row, cols)
    core.pop("factors", None)
    m = core.get("method") or {}
    for k, out in (("rate", "rate_value"), ("valuation_date", "valuation_date_value")):
        if m.get(k) is not None and out not in core:
            try:
                v = dcf.resolve(db, m[k])[0]
                core[out] = v.isoformat() if isinstance(v, date) else dcf._as_date(v).isoformat() if k == "valuation_date" and dcf._as_date(v) else dcf._num(v)
            except (ValueError, AttributeError):
                pass


def parts(db, sheet: str, row: int, cols: list[int], depth: int = 3) -> list[dict]:
    """The rows a row adds up, from its first formula in these columns: terms that are the same column of another
    row, with their sign and total over the columns, each broken down again (up to depth levels)."""
    if depth <= 0:
        return []
    f = next((f for c, f in db.execute("SELECT col, formula FROM cells WHERE sheet=? AND row=? AND col BETWEEN ? AND ? "
                                       "AND formula IS NOT NULL ORDER BY col", (sheet, row, cols[0], cols[-1]))), None)
    if not f:
        return []
    col = next(c for (c,) in db.execute("SELECT col FROM cells WHERE sheet=? AND row=? AND formula=? AND col BETWEEN ? AND ?",
                                         (sheet, row, f, cols[0], cols[-1])))
    out = []
    for op, term in valuation._split(_STR.sub("", f).lstrip("=").lstrip("+"), "+-"):
        r = dcf._ref(valuation._strip(term), sheet)
        if not r or r[1] != r[3] or r[2] != r[4] or r[2] != col:
            return []  # not a plain sum of same-period rows: say nothing rather than half
        s_, r_ = r[0], r[1]
        if not db.execute("SELECT 1 FROM sheets WHERE sheet=?", (s_,)).fetchone():  # another workbook, via a link
            out.append({"row": f"{s_}!r{r_}", "label": "read from a linked workbook", "sign": -1 if op == "-" else 1,
                        "total": None, "parts": [], "words": None, "external": True})
            continue
        vals = _values(db, s_, r_, cols[0], cols[-1])
        sub = parts(db, s_, r_, cols, depth - 1)
        fr = None if sub else next((x for (x,) in db.execute(
            "SELECT formula FROM cells WHERE sheet=? AND row=? AND col BETWEEN ? AND ? AND formula IS NOT NULL ORDER BY col "
            "LIMIT 1", (s_, r_, cols[0], cols[-1]))), None)
        out.append({"row": f"{s_}!r{r_}", "label": _label(db, s_, r_), "sign": -1 if op == "-" else 1,
                    "total": sum(vals.values()), "parts": sub, "words": words(db, fr, s_) if fr else None})
    return out


# ---- the tree ---------------------------------------------------------------------------------------------

def words(db, formula: str, here: str) -> str:
    """The formula with every reference replaced by its line item's name: '= Equity value (cum-div) - Distribution
    payable'. Rows read as 'the <label> row'."""
    def name(m):
        r = dcf._ref(m[0], here)
        if not r:
            return m[0]
        sheet, r1, c1, r2, c2 = r
        lab = _label(db, sheet, r1)
        if r1 == r2 and c1 == c2:
            return lab or _a1(sheet, r1, c1)
        if r1 == r2 and c2 - c1 + 1 > 2:
            return f"the {lab or f'{sheet}!r{r1}'} row"
        last = _label(db, sheet, r2)
        return f"{lab or _a1(sheet, r1, c1)} to {last or _a1(sheet, r2, c2)}"
    return dcf._FREF.sub(name, _STR.sub('""', formula)).replace("*", " × ").replace("^", "^")


def trace(db, cell: str) -> dict:
    """The tree from a cell down to its discounting. Each node: cell, label, value, formula, words, cores (the
    discounting found in its own formula), children (the single cells and short ranges it reads), on_path (a
    core is at or below it). Nodes off every path keep no children; a cell met twice is shown once."""
    names = _names(db)
    r = dcf._ref(cell, "")
    if not r:
        raise ValueError(f"not a cell: {cell!r}")
    seen: dict[tuple, dict] = {}

    def visit(sheet, row, col, depth):
        key = (sheet, row, col)
        if key in seen:
            return {"cell": _a1(sheet, row, col), "label": _label(db, sheet, row), "again": True,
                    "on_path": seen[key].get("on_path", False), "value": seen[key].get("value")}
        f, v = _cell(db, sheet, row, col)
        node = {"cell": _a1(sheet, row, col), "label": _label(db, sheet, row), "value": v, "formula": f,
                "children": [], "cores": []}
        seen[key] = node
        if not f or depth > MAX_DEPTH or len(seen) > MAX_NODES:
            return node
        body = _expand(db, _STR.sub('""', f), names)
        node["words"] = words(db, body, sheet)
        whole = body.strip().lstrip("=").lstrip("+").strip()
        for call in _calls(body):
            c = _core(db, sheet, row, col, call, whole == call[2])
            if c:
                _describe(db, c)
                node["cores"].append(c)
        for m in dcf._FREF.finditer(body):
            rr = dcf._ref(m[0], sheet)
            if not rr:
                continue
            s_, r1, c1, r2, c2 = rr
            if (r2 - r1 + 1) * (c2 - c1 + 1) > SHORT:
                continue  # a row or block: an operand, not followed cell by cell
            for rx in range(r1, r2 + 1):
                for cx in range(c1, c2 + 1):
                    if dcf._cell(db, s_, rx, cx) is None and not db.execute(
                            "SELECT 1 FROM cells WHERE sheet=? AND row=? AND col=?", (s_, rx, cx)).fetchone():
                        continue
                    node["children"].append(visit(s_, rx, cx, depth + 1))
        return node

    root = visit(r[0], r[1], r[2], 0)

    def mark(n):
        if n.get("again"):
            return n["on_path"]
        below = [mark(c) for c in n["children"]]
        n["on_path"] = bool(n["cores"]) or any(below)
        return n["on_path"]

    def prune(n):
        if not n.get("on_path"):
            n["children"] = []  # off every path: the node stays (it's a term), its own inputs don't
        for c in n["children"]:
            prune(c)

    mark(root)
    prune(root)
    return root


def cores(tree: dict) -> list[dict]:
    """Every core in the tree, depth first, each once."""
    out, seen = [], set()

    def walk(n):
        for c in n.get("cores", []):
            if (c["cell"], c["call"]) not in seen:
                seen.add((c["cell"], c["call"]))
                out.append(c)
        for ch in n.get("children", []):
            walk(ch)
    walk(tree)
    return out


def text(tree: dict, fmt=lambda v: f"{v:,.4f}" if isinstance(v, (int, float)) else str(v)) -> str:
    """The tree as indented lines, for the chat."""
    lines = []

    def walk(n, depth):
        pad = "  " * depth
        py = n.get("python")
        extra = f" (Python {fmt(py)})" if py is not None and not math.isclose(py, dcf._num(n.get("value")) or 0.0,
                                                                              rel_tol=1e-9, abs_tol=1e-9) else ""
        head = f"{pad}{n['cell']} {n.get('label') or ''} = {fmt(n.get('value'))}{extra}"
        if n.get("again"):
            lines.append(head + " (see above)")
            return
        lines.append(head + (f"   [{n['words']}]" if n.get("words") else ""))
        for c in n.get("cores", []):
            lines.append(f"{pad}  discounting: {c['what']}; cash flows {c['cashflow']} {c.get('cashflow_label', '')}; "
                         f"PV recomputed {fmt(c['pv'])}" + (f"; rate {c['method']['rate']}, valuation date "
                                                            f"{c['method']['valuation_date']}, {c['method']['timing']}-of-"
                                                            f"period, {c['method']['day_count']}" if c.get("method") else "")
                         + (f"; {c['periods']} periods {c.get('first_period')} to {c.get('last_period')}, undiscounted "
                            f"{fmt(c['undiscounted'])}" if c.get("periods") else ""))
            if c.get("parts"):
                lines.append(f"{pad}  the cash flows are made of:")

                def part(x, d):
                    lines.append(f"{pad}  {'  ' * d}{'+' if x['sign'] > 0 else '-'} {x['row']} {x['label']}"
                                 + (f" (total {fmt(x['total'])})" if x.get("total") is not None else "")
                                 + (f" [{x['words']}]" if x.get("words") else ""))
                    for y in x["parts"]:
                        part(y, d + 1)
                for x in c["parts"]:
                    part(x, 1)
        for ch in n.get("children", []):
            walk(ch, depth + 1)
    walk(tree, 0)
    return "\n".join(lines)


# ---- anchors for the Valuation tab ------------------------------------------------------------------------

_CACHE: dict = {}
_LOCK = threading.Lock()


def anchors(db_path: str, starts: list[str]) -> list[dict]:
    """The discountings under these cells, as anchors the Valuation tab's Validate and Methods steps take (the
    shape valuation.catalogue() gives): only those whose formula is the discounting alone (so the cell's value is
    its PV) and whose rate, valuation date and convention were read back. Cached per model.db."""
    import rodb
    key = (db_path, os.path.getmtime(db_path), tuple(starts))
    with _LOCK:
        if key in _CACHE:
            return _CACHE[key]
    db = rodb.connect(db_path)
    out, seen = [], set()
    for start in starts:
        try:
            tree = trace(db, start)
        except ValueError:
            continue
        for c in cores(tree):
            if not c.get("inputs") or not c["whole"] or c["cell"] in seen:
                continue
            seen.add(c["cell"])
            try:
                r = dcf.compute(db, **c["inputs"], fix=False)
            except ValueError:
                continue
            sheet, row, col = dcf._ref(c["cell"], "")[:3]
            m = c.get("method") or {}
            fr = dcf._ref(c["factor_row"], "") if c.get("factor_row") else None
            out.append({"cell": c["cell"], "label": _label(db, sheet, row) or c["cell"],
                        "value": dcf._num(_cell(db, sheet, row, col)[1]), "pv_cell": c["cell"],
                        "df": (fr[0], fr[1]) if fr else None, "ok": True, "inputs": c["inputs"],
                        "matches": r["compare_to"] is not None and dcf._close(r["total"], r["compare_to"]),
                        "rate_value": r["rate"], "total": r["total"], "rate_note": m.get("rate_note"),
                        "factor_row": m.get("factor_row") or c["what"], "ends_source": m.get("ends_source") or c.get("dates"),
                        "traced_from": start, "kind": c["kind"], "what": c["what"]})
    with _LOCK:
        _CACHE[key] = out
    return out


# ---- carrying new discounting results up to the figure --------------------------------------------------------

def recompute(db, tree: dict, pv: dict[str, float]) -> float | None:
    """The figure at the top of the tree again, with these discountings' results ({core cell: value}) in place of
    theirs: every cell on the way up re-evaluated from its formula (evaluate()), the cells off the path as db has
    them. None when a formula on the way can't be evaluated here (IF, INDEX, ...)."""
    names = _names(db)
    given: dict[tuple, float] = {}

    def value(n):
        r = dcf._ref(n["cell"], "")
        key = (r[0], r[1], r[2])
        if key in given:
            return given[key]
        if n["cell"] in pv:
            given[key] = pv[n["cell"]]
            return given[key]
        if not n.get("on_path") or n.get("again") or not n.get("formula"):
            return dcf._num(_cell(db, *key)[1])
        for c in n.get("children", []):
            v = value(c)
            if v is None:
                raise ValueError(c["cell"])
            rc = dcf._ref(c["cell"], "")
            given[(rc[0], rc[1], rc[2])] = v
        out = evaluate(db, _expand(db, _STR.sub('""', n["formula"]), names), key[0], given)
        given[key] = float(out)
        return given[key]
    try:
        return value(tree)
    except (ValueError, ZeroDivisionError, TypeError, OverflowError):
        return None
