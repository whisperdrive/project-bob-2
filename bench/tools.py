"""Tools an agent can call against out/<workbook>/model.db (built by build_map.py).

Call use(<workbook path or model.db path>) first. Each function returns a compact string - that
string is exactly what would be sent to Claude.
"""
import os
import re
import sqlite3
import threading
from collections import deque
from contextlib import contextmanager

import rodb

DB = None
_local = threading.local()  # per-thread override set by using(), so concurrent requests don't clash
_JUNK_NAME = re.compile(r"^(_|EV__|CIQ|IQ_|Cell)", re.I)


def use(path: str) -> None:
    """Point the tools at a workbook's model.db (accepts the workbook path or the db path)."""
    global DB
    if not path.endswith(".db"):
        path = os.path.join("out", os.path.splitext(os.path.basename(path))[0], "model.db")
    if not os.path.exists(path):
        raise FileNotFoundError(f"{path} not found - run build_map.py on the workbook first")
    DB = path


@contextmanager
def using(db_path: str):
    """Point tool calls in this thread at db_path for the duration of the block."""
    prev, _local.db = getattr(_local, "db", None), db_path
    try:
        yield
    finally:
        _local.db = prev


def _path() -> str:
    path = getattr(_local, "db", None) or DB
    if path is None:
        raise RuntimeError("call tools.use(<workbook>) first")
    return path


_SQUASH = re.compile(r"[\s\-_./&()']+")


def _squash(text) -> str:
    """Lower-case text with spaces and common punctuation removed ("Cash-flow (ex. hist)" -> "cashflowexhist")."""
    return _SQUASH.sub("", str(text).lower()) if text is not None else ""


def _db() -> sqlite3.Connection:
    db = rodb.connect(_path())
    db.create_function("squash", 1, _squash, deterministic=True)
    return db


def _label(db, sheet: str, row: int) -> str:
    r = db.execute("SELECT label, units FROM rows WHERE sheet=? AND row=?", (sheet, row)).fetchone()
    return f"{r[0]} [{r[1]}]" if r else "?"


def overview() -> str:
    """Level-0 map: sheets, sections, useful named ranges, model checks. Send this first (and cache it)."""
    db = _db()
    name = os.path.basename(os.path.dirname(_path()))
    out = [f"Workbook: {name}",
           "Each sheet lists its detected layout (label/units columns, timeline span). Refer to line items",
           "as Sheet!rN. Use tools find/rows/trace/cells/sql for detail.", ""]
    sheets = db.execute("""SELECT s.sheet, s.state, s.summary, COUNT(r.row), SUM(r.n_formula), SUM(r.n_const)
                           FROM sheets s LEFT JOIN rows r ON r.sheet = s.sheet
                           GROUP BY s.sheet ORDER BY MIN(s.rowid)""").fetchall()
    for sheet, state, summary, n, nf, nc in sheets:
        if not n:
            continue  # empty divider / cover sheets
        hidden = "" if state == "visible" else f" ({state})"
        out.append(f"## {sheet}{hidden}: {n} line items, {nf} formula cells, {nc} constants; {summary}")
        secs = db.execute("""SELECT section, MIN(row), MAX(row) FROM rows WHERE sheet=?
                             GROUP BY section ORDER BY MIN(row)""", (sheet,)).fetchall()
        out.append("  sections: " + "; ".join(f"{s or '-'} r{a}-{b}" for s, a, b in secs))
    names = [f"{n}={ref}" for n, ref in db.execute("SELECT name, ref FROM names ORDER BY name")
             if not _JUNK_NAME.match(n) and "#REF" not in ref and "!" in ref]
    out.append(f"\nNamed ranges ({len(names)}): " + ", ".join(names))
    return "\n".join(out)


def find(text: str, sheet: str | None = None, limit: int = 25) -> str:
    """Line items whose label or section contains text (case-insensitive; spaces and punctuation ignored,
    so "cash flow" also finds "Cashflow")."""
    db = _db()
    q = ("SELECT sheet,row,section,label,units,samples FROM rows WHERE (label LIKE ? OR section LIKE ? "
         "OR squash(label) LIKE ? OR squash(section) LIKE ?)")
    key = _squash(text)
    args: list = [f"%{text}%", f"%{text}%", f"%{key}%", f"%{key}%"]
    if sheet:
        q += " AND sheet=?"
        args.append(sheet)
    res = db.execute(q + " ORDER BY rowid LIMIT ?", (*args, limit + 1)).fetchall()
    lines = [f"{s}!r{r} {lab} [{u}] (in: {sec}) {('| eg ' + sm) if sm else ''}"
             for s, r, sec, lab, u, sm in res[:limit]]
    if len(res) > limit:
        lines.append(f"... more than {limit} matches; narrow the search")
    return "\n".join(lines) or "no matches"


def rows(sheet: str, r1: int, r2: int | None = None) -> str:
    """Full detail (formula patterns in R1C1, counts, samples) for rows r1..r2 of a sheet."""
    db = _db()
    res = db.execute("""SELECT row,label,units,n_formula,n_const,patterns,samples FROM rows
                        WHERE sheet=? AND row BETWEEN ? AND ? ORDER BY row LIMIT 60""",
                     (sheet, r1, r2 or r1)).fetchall()
    return "\n".join(f"{sheet}!r{r} {lab} [{u}] f={nf} c={nc} | {p} | eg {sm}"
                     for r, lab, u, nf, nc, p, sm in res) or "no line items in range"


def trace(sheet: str, row: int, direction: str = "up", depth: int = 2, limit: int = 60,
          include_inactive: bool = False) -> str:
    """Line-item dependency tree. up = precedents (what feeds it), down = dependents (what it feeds).
    By default follows what the current scenario uses: rows a SUMIFS / INDEX-MATCH / CHOOSE considers
    but doesn't pick are left out and counted (include_inactive=True lists them, marked "not selected")."""
    db = _db()
    has_kind = "kind" in [r[1] for r in db.execute("PRAGMA table_info(edges)")]
    k = ", kind" if has_kind else ", 'direct'"
    if direction == "up":
        q = f"SELECT dst_sheet, dst_row{k} FROM edges WHERE src_sheet=? AND src_row=?"
    else:
        q = f"SELECT src_sheet, src_row{k} FROM edges WHERE dst_sheet=? AND dst_row=?"
    out = [f"{sheet}!r{row} {_label(db, sheet, row)}"]
    seen = {(sheet, row)}
    queue = deque([(sheet, row, 0)])
    hidden = 0
    while queue and len(out) < limit:
        s, r, d = queue.popleft()
        if d >= depth:
            continue
        for ns, nr, kind in sorted(db.execute(q, (s, r)).fetchall()):
            if (ns, nr) in seen:
                continue
            if kind == "inactive" and not include_inactive:
                hidden += 1
                continue
            seen.add((ns, nr))
            tag = {"inactive": "  (not selected in the current scenario)", "active": "  (selected)",
                   "offset": "  (via OFFSET)"}.get(kind, "")
            out.append(f"{'  ' * (d + 1)}{'<-' if direction == 'up' else '->'} {ns}!r{nr} {_label(db, ns, nr)}{tag}")
            if kind != "inactive":
                queue.append((ns, nr, d + 1))
            if len(out) >= limit:
                out.append(f"... truncated at {limit}")
                break
    if hidden:
        out.append(f"(+{hidden} rows a lookup considers but doesn't select in the current scenario; "
                   f"trace(..., include_inactive=True) lists them)")
    return "\n".join(out)


def cells(sheet: str, addr_from: str, addr_to: str | None = None, limit: int = 80) -> str:
    """Raw cells (formula + cached value) in an A1 range, e.g. cells('Valuation','K100','P105')."""
    from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
    c1, r1 = coordinate_from_string(addr_from)
    c2, r2 = coordinate_from_string(addr_to or addr_from)
    db = _db()
    res = db.execute("""SELECT addr, formula, value FROM cells WHERE sheet=? AND row BETWEEN ? AND ?
                        AND col BETWEEN ? AND ? ORDER BY row, col LIMIT ?""",
                     (sheet, r1, r2, column_index_from_string(c1), column_index_from_string(c2), limit)).fetchall()
    return "\n".join(f"{a}: {f + ' -> ' if f else ''}{v}" for a, f, v in res) or "empty"


def sql(query: str, limit: int = 50) -> str:
    """Read-only SQL over tables cells(sheet,row,col,addr,formula,value), rows(...), edges(...), names(...)."""
    db = _db()
    cur = db.execute(query)
    cols = [d[0] for d in cur.description]
    res = cur.fetchmany(limit + 1)
    lines = ["\t".join(cols)] + ["\t".join(str(x) for x in r) for r in res[:limit]]
    if len(res) > limit:
        lines.append(f"... truncated at {limit} rows")
    return "\n".join(lines)


MAX_POINTS = 400
_RANGE = re.compile(r"^'?(?P<sheet>[^!]+?)'?!\$?(?P<c1>[A-Z]{1,3})\$?(?P<r1>\d+)(?::\$?(?P<c2>[A-Z]{1,3})\$?(?P<r2>\d+))?$")


def _range_cells(db, ref: str) -> tuple[str, list[tuple[int, int]]]:
    """'Valuation!L95:AO95' -> (sheet, [(row, col), ...]) for a single row or single column range."""
    from openpyxl.utils.cell import column_index_from_string
    m = _RANGE.match(ref.strip()) or _RANGE.match(ref.strip().replace(" ", ""))  # sheet names may have spaces
    if not m:
        raise ValueError(f"not a cell range: {ref!r} (use Sheet!L95:AO95)")
    sheet, r1 = m["sheet"], int(m["r1"])
    c1 = column_index_from_string(m["c1"])
    r2, c2 = (int(m["r2"]), column_index_from_string(m["c2"])) if m["c2"] else (r1, c1)
    if r1 != r2 and c1 != c2:
        raise ValueError(f"{ref}: use one row or one column per series")
    pts = [(r1, c) for c in range(c1, c2 + 1)] if r1 == r2 else [(r, c1) for r in range(r1, r2 + 1)]
    if len(pts) > MAX_POINTS:
        raise ValueError(f"{ref}: {len(pts)} points; chart at most {MAX_POINTS}")
    if not db.execute("SELECT 1 FROM sheets WHERE sheet=?", (sheet,)).fetchone():
        raise ValueError(f"no sheet named {sheet!r}")
    return sheet, pts


def _period_label(v) -> str:
    s = str(v)
    return s[:10] if re.match(r"^\d{4}-\d{2}-\d{2}", s) else s[:24]


_ROW_TOTAL = re.compile(r"^=\+?\(?SUM\(RC\[(-?\d+)\]:RC\[(-?\d+)\]\)\)?$", re.I)


def _total_columns(db, sheet: str, pts: list[tuple[int, int]]) -> set[int]:
    """Columns at either end of a row range whose formula sums the rest of that row (a "row total" column,
    e.g. K = SUM(L:HO)); charting them would add the whole row's total as a period."""
    import sys as _sys, os as _os
    _sys.path.insert(0, _os.path.dirname(__file__))
    from build_map import to_pattern
    if len(pts) < 4 or pts[0][0] != pts[-1][0]:
        return set()
    row, drop = pts[0][0], set()
    for r, c in (pts[0], pts[-1]):
        f = db.execute("SELECT formula FROM cells WHERE sheet=? AND row=? AND col=?", (sheet, r, c)).fetchone()
        if f and f[0]:
            m = _ROW_TOTAL.match(to_pattern(f[0], r, c, sheet)[0].replace(" ", ""))
            if m and abs(int(m[2]) - int(m[1])) + 1 >= 0.5 * (len(pts) - 1):
                drop.add(c)
    return drop


KINDS = ("line", "bar", "stacked", "area", "combo", "waterfall")


def chart(title: str, series: list[dict], kind: str = "line", x_range: str | None = None,
          units: str | None = None, totals: list[int] | None = None) -> dict:
    """Data for a chart the UI draws. Each series is {"range": "Sheet!L95:AO95", "name": optional, "as": "bar" |
    "line" (combo / stacked: drawn as columns or as a line on top), "axis": "right" (its own axis on the right)}.
    Values are read from the workbook; x labels come from x_range, the sheet's timeline header row (a row
    range), or the line items' labels (a column range). kind: line, bar (clustered), stacked, area (stacked),
    combo (stacked columns + lines), waterfall (one series as a bridge; totals = positions of the bars drawn
    from zero, e.g. the start, subtotals and the end; single cells as series become the bridge's steps)."""
    import json as _json
    db = _db()
    out_series, labels = [], None
    dropped: list[int] = []  # positions removed from every series (row-total columns)
    for s in series[:8]:
        sheet, pts = _range_cells(db, s["range"])
        if not dropped and out_series == []:
            tot = _total_columns(db, sheet, pts)
            dropped = [i for i, p in enumerate(pts) if p[1] in tot]
        pts = [p for i, p in enumerate(pts) if i not in dropped]
        vals = {(r, c): v for r, c, v in db.execute(
            f"SELECT row, col, value FROM cells WHERE sheet=? AND row BETWEEN ? AND ? AND col BETWEEN ? AND ?",
            (sheet, min(p[0] for p in pts), max(p[0] for p in pts), min(p[1] for p in pts), max(p[1] for p in pts)))}
        data = []
        for p in pts:
            v = vals.get(p)
            data.append(v if isinstance(v, (int, float)) and not isinstance(v, bool) else None)
        row0 = pts[0][0]
        lab = db.execute("SELECT label, units FROM rows WHERE sheet=? AND row=?", (sheet, row0)).fetchone()
        one_row = pts[0][0] == pts[-1][0]
        name = s.get("name") or (lab[0] if lab and lab[0] and one_row else s["range"])
        out_series.append({"name": name, "range": s["range"], "data": data,
                           "label": lab[0] if lab and one_row else None,
                           "units": lab[1] if lab and one_row else None,
                           **({"as": s["as"]} if s.get("as") in ("bar", "line") else {}),
                           **({"axis": "right"} if s.get("axis") == "right" else {})})
        if labels is None:
            if x_range:
                xs, xpts = _range_cells(db, x_range)
                xpts = [p for i, p in enumerate(xpts) if i not in dropped]
                xv = dict(((r, c), v) for r, c, v in db.execute(
                    "SELECT row, col, value FROM cells WHERE sheet=?", (xs,)) if (r, c) in set(xpts))
                labels = [_period_label(xv.get(p, "")) for p in xpts]
            elif pts[0][0] == pts[-1][0]:  # a row: use the timeline header for those columns
                from openpyxl.utils import get_column_letter
                lay = db.execute("SELECT layout FROM sheets WHERE sheet=?", (sheet,)).fetchone()
                hr = _json.loads(lay[0] or "{}").get("header_row") if lay else None
                hv = dict(db.execute("SELECT col, value FROM cells WHERE sheet=? AND row=?", (sheet, hr))) if hr else {}
                labels = [_period_label(hv[c]) if c in hv else get_column_letter(c) for _, c in pts]
            else:  # a column: the line items' labels
                labs = dict(db.execute(f"SELECT row, label FROM rows WHERE sheet=? AND row IN ({','.join('?' * len(pts))})",
                                       (sheet, *[r for r, _ in pts])))
                labels = [labs.get(r) or str(r) for r, _ in pts]
    n = max((len(s["data"]) for s in out_series), default=0)
    labels = (labels or [])[:n] + [""] * (n - len(labels or []))
    import chartdata
    if not units:  # all series share their rows' units (e.g. pax '000): use them
        us = {s.get("units") for s in out_series}
        if len(us) == 1 and next(iter(us)):
            units = next(iter(us))
    first = [p for i, p in enumerate(_range_cells(db, series[0]["range"])[1]) if i not in dropped] if series else []
    kind = kind if kind in KINDS else "line"
    if kind == "waterfall" and len(out_series) > 1 and all(len(x["data"]) == 1 for x in out_series):
        # single cells (e.g. enterprise value, net debt, equity value): each is one step of the bridge
        labels = [x["name"] for x in out_series]
        out_series = [{"name": title, "range": ", ".join(x["range"] for x in out_series), "label": None,
                       "units": units, "data": [x["data"][0] for x in out_series]}]
        first = []
    spec = {"title": title, "kind": kind, "units": units, "labels": labels, "series": out_series,
            "columns": [c for r, c in first] if first and first[0][0] == first[-1][0] else None}
    if kind == "waterfall":
        spec["totals"] = sorted({int(i) for i in totals or [] if 0 <= int(i) < len(labels)})
    if dropped:
        spec["dropped_total_columns"] = len(dropped)
    spec = chartdata.enrich(spec, db)
    if kind == "waterfall":  # a bridge keeps its signs: the steps down are the point
        spec["sign"] = 1
    return spec


def _fuller_rows(rng: str, limit: int = 4) -> str:
    """Rows that depend on the charted row (up to 3 levels down) and have values across more of the timeline."""
    try:
        db = _db()
        sheet, pts = _range_cells(db, rng)
        row = pts[0][0]
    except Exception:
        return ""
    seen, frontier, found = {(sheet, row)}, [(sheet, row)], []
    for _ in range(3):
        nxt = []
        for s, r in frontier:
            for ds, dr in db.execute("SELECT src_sheet, src_row FROM edges WHERE dst_sheet=? AND dst_row=?", (s, r)):
                if (ds, dr) not in seen:
                    seen.add((ds, dr))
                    nxt.append((ds, dr))
        frontier = nxt
        for ds, dr in nxt:
            n = db.execute("SELECT COUNT(*) FROM cells WHERE sheet=? AND row=? AND typeof(value) IN ('real','integer') "
                           "AND value<>0", (ds, dr)).fetchone()[0]
            found.append((n, ds, dr))
    found.sort(reverse=True)
    out = []
    for n, ds, dr in found[:limit]:
        lab = db.execute("SELECT label, units FROM rows WHERE sheet=? AND row=?", (ds, dr)).fetchone()
        out.append(f"{ds}!r{dr} {lab[0] if lab else ''} [{lab[1] if lab else ''}] ({n} non-zero periods)")
    return "; ".join(out)


def chartdata_methods(a: dict) -> list[str]:
    import chartdata
    return [chartdata.METHOD_TEXT[m] for m in a["methods"]]


def chart_note(spec: dict) -> str:
    """What the model is told after drawing: enough to describe the chart without re-reading every value."""
    shown = spec.get("period_labels") or spec["labels"]
    lines = [f"Chart shown to the user: '{spec['title']}', {len(shown)} {spec.get('frequency', '')} points "
             f"({shown[0] if shown else ''} to {shown[-1] if shown else ''}); values below are the workbook's."]
    k = spec.get("kind") or "line"
    if k != "line":
        lines.append({"bar": "Drawn as clustered columns.", "stacked": "Drawn as stacked columns (series marked line "
                      "on top as lines).", "area": "Drawn as stacked areas.", "combo": "Drawn as stacked columns with "
                      "the series marked line as lines.", "waterfall": "Drawn as a waterfall: each value moves the "
                      "running total" + (f"; bars from zero at positions {spec.get('totals')}" if spec.get("totals")
                                         else "; a Total bar is added at the end") + "."}[k])
    if any(s.get("axis") == "right" for s in spec["series"]):
        lines.append("On the right-hand axis: " + ", ".join(s["name"] for s in spec["series"] if s.get("axis") == "right") + ".")
    if spec.get("sign") == -1:
        lines.append("All values are negative in the workbook, so the chart shows them as positive (labelled).")
    if spec.get("annual"):
        a = spec["annual"]
        lines.append(f"The user can switch to annual figures by {a['basis']} "
                     f"({', '.join(chartdata_methods(a))}); years marked * are partial.")
    if spec.get("dropped_total_columns"):
        lines.append("A row-total column (a SUM of the rest of the row) was left out of the chart.")
    if spec.get("phases"):
        lines.append("Shaded on the chart: " + ", ".join(
            f"{p['name']} {shown[p['start']]} to {shown[p['end']]}" for p in spec["phases"]) + ".")
    n_periods = len(shown)
    for s in spec["series"]:
        nz = [i for i, v in enumerate(s["data"]) if isinstance(v, (int, float)) and v]
        if n_periods >= 8 and nz and (nz[-1] - nz[0] + 1) < 0.6 * n_periods:
            hint = _fuller_rows(s["range"])
            lines.append(f"WARNING: {s['name']} ({s['range']}) only has values from {shown[nz[0]]} to "
                         f"{shown[nz[-1]]} ({nz[-1] - nz[0] + 1} of {n_periods} periods). It looks like an input or "
                         f"actuals-only row. Look for the calculated row that combines actuals and forecast over the "
                         f"whole timeline and chart that instead" + (f"; rows that use this one and cover more "
                                                                     f"periods: {hint}" if hint else "") + ".")
    for s in spec["series"]:
        nums = [v for v in s["data"] if v is not None]
        if nums:
            lines.append(f"- {s['name']} ({s['range']}): first {nums[0]:.6g}, last {nums[-1]:.6g}, "
                         f"min {min(nums):.6g}, max {max(nums):.6g}, total {sum(nums):.6g}, {len(nums)} numeric")
        else:
            lines.append(f"- {s['name']} ({s['range']}): no numeric values")
    return "\n".join(lines)


def dcf(cashflow=None, rate=None, valuation_date=None, dates: str | None = None, timing: str = "end",
        day_count: str = "actual/actual", terminal_date=None, adjustments: list | None = None,
        compare_to=None, rates: list | None = None) -> str:
    """Recompute a DCF from cash-flow rows in Python (see dcf.py). Without a cash-flow row, lists the model's
    valuation result cells; without a rate or valuation date, lists the cells that look like one."""
    return dcf_result(cashflow, rate, valuation_date, dates, timing, day_count, terminal_date, adjustments,
                      compare_to, rates)[0]


def dcf_result(cashflow=None, rate=None, valuation_date=None, dates=None, timing="end", day_count="actual/actual",
               terminal_date=None, adjustments=None, compare_to=None, rates=None) -> tuple[str, dict | None]:
    """(text for the model, a card for the UI or None when only candidates were listed)."""
    import dcf as _dcf
    db = _db()
    blank = lambda v: v in (None, "", [])
    if blank(cashflow):
        found = _dcf.outputs(db)
        return ("cashflow not given. Cells that look like the model's valuation results, with their formulas "
                "(follow the formula: the row a SUMPRODUCT / XNPV / NPV discounts is the cash flow; the other "
                "terms of a SUM are the bridge):\n  " + ("\n  ".join(found) if found else "none found; use find()")), None
    missing = [w for w, v in (("rate", rate), ("valuation_date", valuation_date)) if blank(v)]
    if missing:
        out = []
        for w in missing:
            c = _dcf.candidates(db, "rate" if w == "rate" else "date")
            out.append(f"{w} not given. Candidates:\n  " + ("\n  ".join(c) if c else "none found; use find()"))
        return "\n".join(out) + "\nCall dcf again with the ones this valuation uses (trace the discount factor row if unsure).", None
    r = _dcf.compute(db, cashflow, rate, valuation_date, dates or None, timing or "end", day_count or "actual/actual",
                     terminal_date or None, adjustments or None, compare_to or None, rates or None)
    return _dcf.report(r), _dcf.card(r)
