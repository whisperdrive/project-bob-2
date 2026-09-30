"""Recompute a DCF from a workbook's cash-flow rows, in plain Python (no Excel, no formula engine).

The cash flows are the workbook's saved values; this redoes only the discounting:
    PV = sum of CF_t / (1 + rate) ^ t, t = years from the valuation date to the period end (or mid-point)
Periods ending on or before the valuation date (and after terminal_date, if given) are left out, as
DCF sheets usually do with a "post valuation date" flag.

Conventions:
  timing    "end" (default), "mid" or "mid-year": discount to the period end, to the middle of the part of the
            period after the valuation date, or half a year before the period end (the mid-year convention as
            models often write it: (end - valuation date) / 365 - 0.5)
  day_count "actual/actual" (Excel YEARFRAC basis 1, default) or "actual/365" (what XNPV uses)
  "auto"    with compare_to: tries every combination and keeps the one closest to the workbook's value

dcf() returns a compact text report for the agent; compute() returns the numbers.
"""
import re
import sqlite3
from calendar import monthrange
from datetime import date, datetime, timedelta

_DATE = re.compile(r"^(\d{4})-(\d{2})-(\d{2})")
_REF = re.compile(r"^'?(?P<sheet>[^!]+?)'?!\$?(?P<c1>[A-Z]{1,3})\$?(?P<r1>\d+)(?::\$?(?P<c2>[A-Z]{1,3})\$?(?P<r2>\d+))?$")
_ROWREF = re.compile(r"^'?(?P<sheet>[^!]+?)'?!r?(?P<row>\d+)$", re.I)
_END_LABEL = re.compile(r"\bend\b.*\bdate\b|period\s*end|end\s*of\s*period|period\s*ending", re.I)
TIMINGS = ("end", "mid", "mid-year")
DAY_COUNTS = ("actual/actual", "actual/365")


# ---- values -----------------------------------------------------------------------------------------------

def _as_date(v) -> date | None:
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    m = _DATE.match(str(v or "").strip())
    return date(int(m[1]), int(m[2]), int(m[3])) if m else None


def _num(v) -> float | None:
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def _col(letters: str) -> int:
    from openpyxl.utils.cell import column_index_from_string
    return column_index_from_string(letters)


def _addr(col: int, row: int) -> str:
    from openpyxl.utils import get_column_letter
    return f"{get_column_letter(col)}{row}"


def _check_sheet(db, sheet: str) -> str:
    if not db.execute("SELECT 1 FROM sheets WHERE sheet=?", (sheet,)).fetchone():
        raise ValueError(f"no sheet named {sheet!r}")
    return sheet


def _cell(db, sheet: str, row: int, col: int):
    r = db.execute("SELECT value FROM cells WHERE sheet=? AND row=? AND col=?", (sheet, row, col)).fetchone()
    return r[0] if r else None


def _row_label(db, sheet: str, row: int) -> str:
    r = db.execute("SELECT label, units FROM rows WHERE sheet=? AND row=?", (sheet, row)).fetchone()
    return (r[0] + (f" [{r[1]}]" if r[1] else "")) if r and r[0] else ""


def resolve(db, ref) -> tuple[object, str]:
    """A number, an ISO date, a named range, or a single cell (Sheet!A1) -> (value, where it came from)."""
    if isinstance(ref, (int, float)) and not isinstance(ref, bool):
        return ref, "given"
    s = str(ref).strip()
    if _as_date(s) and len(s) <= 19:
        return _as_date(s), "given"
    try:
        return float(s.rstrip("%")) / (100 if s.endswith("%") else 1), "given"
    except ValueError:
        pass
    named = db.execute("SELECT ref FROM names WHERE lower(name)=lower(?)", (s,)).fetchone()
    target = named[0] if named else s
    # as written first (sheet names may have spaces); then without spaces, for a typed "DCF! D12"
    m = _REF.match(target.strip()) or (_REF.match(target.replace(" ", "")) if "'" not in target else None)
    if not m or (m["c2"] and (m["c2"], m["r2"]) != (m["c1"], m["r1"])):
        raise ValueError(f"{ref!r}: give a number, a date (YYYY-MM-DD), a named range or one cell (Sheet!F435)")
    sheet = _check_sheet(db, m["sheet"])
    row, col = int(m["r1"]), _col(m["c1"])
    where = f"{sheet}!{_addr(col, row)}" + (f" ({s})" if named else "")
    lab = _row_label(db, sheet, row)
    return _cell(db, sheet, row, col), where + (f" {lab}" if lab else "")


# ---- the timeline -----------------------------------------------------------------------------------------

def _layout(db, sheet: str) -> dict:
    import json
    r = db.execute("SELECT layout FROM sheets WHERE sheet=?", (sheet,)).fetchone()
    return json.loads(r[0] or "{}") if r else {}


def _row_range(db, ref: str) -> tuple[str, int, list[int]]:
    """'Sheet!L95:HO95' or 'Sheet!r95' (the sheet's whole timeline) -> (sheet, row, [col, ...])."""
    s = ref.strip()
    m = _ROWREF.match(s)
    if m:
        sheet = _check_sheet(db, m["sheet"])
        lay = _layout(db, sheet)
        if not lay.get("tl_first"):
            raise ValueError(f"{sheet} has no detected timeline; give a range like {sheet}!L{m['row']}:HO{m['row']}")
        return sheet, int(m["row"]), list(range(lay["tl_first"], lay["tl_last"] + 1))
    m = _REF.match(s)
    if not m or not m["c2"] or m["r1"] != m["r2"]:
        raise ValueError(f"{ref!r}: give one row, e.g. Sheet!L95:HO95 or Sheet!r95 for the whole timeline")
    sheet = _check_sheet(db, m["sheet"])
    return sheet, int(m["r1"]), list(range(_col(m["c1"]), _col(m["c2"]) + 1))


def _month_end(d: date) -> date:
    return date(d.year, d.month, monthrange(d.year, d.month)[1])


def period_ends(db, sheet: str, cols: list[int], dates: str | None = None) -> tuple[dict[int, date], str]:
    """Period end date for each column, and a note on where they came from.
    Uses `dates` if given; else a row on the sheet labelled like "End date" / "Period end"; else the
    timeline header row: dates on the 1st of a month are period starts (ends = next start - 1 day)."""
    if dates:
        ds, dr, dcols = _row_range(db, dates)
        vals = {c: _as_date(v) for c, v in db.execute(
            "SELECT col, value FROM cells WHERE sheet=? AND row=?", (ds, dr))}
        return {c: vals[c] for c in cols if vals.get(c)}, f"{ds}!r{dr} {_row_label(db, ds, dr)}".strip()
    for r, lab in db.execute("SELECT row, label FROM rows WHERE sheet=? ORDER BY row", (sheet,)):
        if lab and _END_LABEL.search(lab):
            vals = {c: _as_date(v) for c, v in db.execute(
                "SELECT col, value FROM cells WHERE sheet=? AND row=?", (sheet, r))}
            got = {c: vals[c] for c in cols if vals.get(c)}
            if len(got) >= 0.8 * len(cols):
                return got, f"{sheet}!r{r} {lab}"
    hr = _layout(db, sheet).get("header_row")
    if not hr:
        raise ValueError(f"no period dates found on {sheet}; pass dates='Sheet!rN' (the period end date row)")
    vals = {c: _as_date(v) for c, v in db.execute("SELECT col, value FROM cells WHERE sheet=? AND row=?", (sheet, hr))}
    got = {c: vals[c] for c in cols if vals.get(c)}
    if not got:
        raise ValueError(f"no period dates in {sheet}!r{hr}; pass dates='Sheet!rN' (the period end date row)")
    if all(d.day == 1 for d in got.values()):  # period starts -> ends
        cs = sorted(got)
        starts = [got[c] for c in cs]
        gaps = sorted((b - a).days for a, b in zip(starts, starts[1:])) or [91]
        months = max(1, round(gaps[len(gaps) // 2] / 30.44))
        ends = {c: b - timedelta(days=1) for c, b in zip(cs, starts[1:])}
        y, m = divmod(starts[-1].month - 1 + months, 12)
        ends[cs[-1]] = date(starts[-1].year + y, m + 1, 1) - timedelta(days=1)
        return ends, f"{sheet}!r{hr} {_row_label(db, sheet, hr)} (period starts; ends derived)".strip()
    return got, f"{sheet}!r{hr} {_row_label(db, sheet, hr)}".strip()


# ---- discounting ------------------------------------------------------------------------------------------

def _leap(y: int) -> bool:
    return y % 4 == 0 and (y % 100 != 0 or y % 400 == 0)


def yearfrac(a: date, b: date, day_count: str = "actual/actual") -> float:
    """Excel YEARFRAC(a, b, 1) for actual/actual; days / 365 for actual/365."""
    if a > b:
        a, b = b, a
    if day_count == "actual/365":
        return (b - a).days / 365
    y1, y2 = a.year, b.year
    try:
        one_year = date(y1 + 1, a.month, a.day)
    except ValueError:  # 29 Feb
        one_year = date(y1 + 1, 3, 1)
    if b <= one_year:
        if y1 == y2:
            den = 366 if _leap(y1) else 365
        else:
            den = 366 if any(_leap(y) and a <= date(y, 2, 29) <= b for y in (y1, y2)) else 365
    else:
        den = sum(366 if _leap(y) else 365 for y in range(y1, y2 + 1)) / (y2 - y1 + 1)
    return (b - a).days / den


def factors(ends, val_date, rate, timing, day_count, terminal_date=None) -> dict[int, float]:
    """Discount factor per column: 0 on or before the valuation date and after terminal_date."""
    cols = sorted(ends)
    out = {}
    for i, c in enumerate(cols):
        e = ends[c]
        if e <= val_date or (terminal_date and e > terminal_date):
            out[c] = 0.0
            continue
        if timing == "mid":
            start = ends[cols[i - 1]] + timedelta(days=1) if i else e - timedelta(days=90)
            start = max(start, val_date)
            t = yearfrac(val_date, start + (e - start) / 2, day_count)
        elif timing == "mid-year":
            t = yearfrac(val_date, e, day_count) - 0.5
        else:
            t = yearfrac(val_date, e, day_count)
        out[c] = 1 / (1 + rate) ** t
    return out


def _pv(flows, ends, val_date, rate, timing, day_count, terminal_date=None) -> tuple[float, int]:
    """flows/ends: {col: value}, {col: end date}. Returns (PV, periods discounted)."""
    f = factors(ends, val_date, rate, timing, day_count, terminal_date)
    return sum((flows.get(c) or 0.0) * f[c] for c in f), sum(1 for c in f if f[c])


def _close(a: float, b: float) -> bool:
    return abs(a - b) <= 1e-6 * max(1.0, abs(b))


# ---- reading the workbook's own calculation ---------------------------------------------------------------

_FREF = re.compile(r"(?:'(?P<qs>[^']+)'|(?P<s>[A-Za-z_][\w.]*))?!?\$?(?P<c1>[A-Z]{1,3})\$?(?P<r1>\d+)"
                   r"(?::\$?(?P<c2>[A-Z]{1,3})\$?(?P<r2>\d+))?")
_ARG = r"\s*([^,()]+?)\s*"
_SUMPRODUCT = re.compile(r"SUMPRODUCT\(" + _ARG + "," + _ARG + r"\)", re.I)
_SUMIFS = re.compile(r"SUMIFS\(" + _ARG + "," + _ARG + "," + _ARG + r"\)", re.I)


def _ref(text: str, here: str):
    """'$L173:$HO173' / 'Debt!F8' -> (sheet, r1, c1, r2, c2), or None."""
    t = text.strip()
    if "!" in t:
        sh, a = t.rsplit("!", 1)
        sh = sh.strip("'")
    else:
        sh, a = here, t
    m = re.fullmatch(r"\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?", a)
    if not m:
        return None
    r1, c1 = int(m[2]), _col(m[1])
    return sh, r1, c1, int(m[4]) if m[3] else r1, _col(m[3]) if m[3] else c1


def workbook_calc(db, cell: str, depth: int = 3) -> dict:
    """Follow a value cell's formula (and the single cells / short ranges it refers to) to find what the
    workbook discounts: SUMPRODUCT(row, row) pairs and SUMIFS(row, date row, ...) terms."""
    named = db.execute("SELECT ref FROM names WHERE lower(name)=lower(?)", (cell,)).fetchone()
    start = _ref(named[0] if named else cell, "")
    found = {"sumproduct": [], "sumifs": [], "tree": []}
    if not start or not start[0]:
        return found
    seen, queue = set(), [(start[0], start[1], start[2], 0)]
    while queue and len(seen) < 25:
        sh, r, c, d = queue.pop(0)
        if (sh, r, c) in seen:
            continue
        seen.add((sh, r, c))
        row = db.execute("SELECT formula, value FROM cells WHERE sheet=? AND row=? AND col=?", (sh, r, c)).fetchone()
        if not row or not row[0]:
            continue
        f, v = row
        found["tree"].append(("  " * d) + f"{sh}!{_addr(c, r)} = {_show(v)}  {f[:120]}")
        for m in _SUMPRODUCT.finditer(f):
            a, b = _ref(m[1], sh), _ref(m[2], sh)
            if a and b and a[1] == a[3] and b[1] == b[3]:
                found["sumproduct"].append((a, b))
        for m in _SUMIFS.finditer(f):
            a, b = _ref(m[1], sh), _ref(m[2], sh)
            if a and b and a[1] == a[3]:
                found["sumifs"].append((a, b, m[3], sh))
        if d >= depth:
            continue
        for m in _FREF.finditer(re.sub(r'"[^"]*"', "", f)):
            ref = _ref(m[0], sh)
            if not ref:
                continue
            s2, r1, c1, r2, c2 = ref
            if (r2 - r1 + 1) * (c2 - c1 + 1) <= 12:  # single cells and short ranges (a SUM of bridge items)
                queue += [(s2, rr, cc, d + 1) for rr in range(r1, r2 + 1) for cc in range(c1, c2 + 1)]
    return found


def diagnose(db, calc: dict, rows_given: list[tuple[str, int]], ours: dict[int, float], ends: dict[int, date],
             bridge: list[dict], convention, vd: date) -> tuple[list[str], dict]:
    """Compare this run with what the workbook's formula does. Returns (notes, fixes): fixes holds the changes
    the workbook's own formula shows unambiguously (cashflow, terminal_date, timing/day_count, adjustments)."""
    out, fixes = [], {}
    one = len(calc["sumproduct"]) == 1
    for cf, dfr in calc["sumproduct"]:
        # which side is the discount factor row: values all within [0, 1]
        def is_df(ref):
            vals = [_num(v) for (v,) in db.execute("SELECT value FROM cells WHERE sheet=? AND row=? AND col BETWEEN ? AND ?",
                                                   (ref[0], ref[1], ref[2], ref[4]))]
            vals = [v for v in vals if v is not None]
            return vals and all(0 <= v <= 1 for v in vals) and any(0 < v < 1 for v in vals)
        if is_df(cf) and not is_df(dfr):
            cf, dfr = dfr, cf
        out.append(f"workbook discounts {cf[0]}!r{cf[1]} {_row_label(db, cf[0], cf[1])} with factors "
                   f"{dfr[0]}!r{dfr[1]} {_row_label(db, dfr[0], dfr[1])}".rstrip())
        if (cf[0], cf[1]) not in rows_given:
            out.append(f"  -> you passed {', '.join(f'{s}!r{r}' for s, r in rows_given)}; use {cf[0]}!r{cf[1]}")
            if one:
                fixes["cashflow"] = [f"{cf[0]}!{_addr(cf[2], cf[1])}:{_addr(cf[4], cf[1])}"]
        theirs = {c: _num(v) or 0.0 for c, v in db.execute(
            "SELECT col, value FROM cells WHERE sheet=? AND row=?", (dfr[0], dfr[1])) if c in ours}
        if theirs:
            diff = max(abs(ours[c] - theirs.get(c, 0.0)) for c in ours)
            if diff < 1e-9:
                out.append("  discount factors: identical to the workbook's")
            else:
                live = [c for c in sorted(ours) if theirs.get(c)]
                cut = [c for c in sorted(ours) if ours[c] > 0 and not theirs.get(c)]
                if live and cut and min(ends[c] for c in cut) > max(ends[c] for c in live):
                    out.append(f"  -> the workbook's factors are 0 after {ends[live[-1]]}: pass that date (or the "
                               f"workbook's terminal date cell) as terminal_date")
                    if one:
                        fixes["terminal_date"] = ends[live[-1]].isoformat()
                both = [c for c in live if ours.get(c)]
                if both and max(abs(ours[c] - theirs[c]) for c in both) > 1e-9:
                    fits = [(t, d) for t in TIMINGS for d in DAY_COUNTS
                            if max(abs(v - theirs[c]) for c, v in convention(t, d).items() if c in both) < 1e-9]
                    if fits:
                        out.append(f"  -> the workbook's factors match timing='{fits[0][0]}', day_count='{fits[0][1]}'")
                        if one:
                            fixes["timing"], fixes["day_count"] = fits[0]
                    else:
                        worst = max(both, key=lambda c: abs(ours[c] - theirs[c]))
                        out.append(f"  discount factors differ (e.g. {ends[worst]}: yours {ours[worst]:.6f}, workbook "
                                   f"{theirs[worst]:.6f}) and no convention fits: check the rate and valuation date")
    for rng, date_rng, crit, here in calc["sumifs"]:
        out.append(f"workbook adds SUMIFS of {rng[0]}!r{rng[1]} {_row_label(db, rng[0], rng[1])} where "
                   f"{date_rng[0]}!r{date_rng[1]} = {crit}".rstrip())
        if not any(b["source"].startswith(f"{rng[0]}!") and re.search(rf"[A-Z]+{rng[1]}\b", b["source"].split(" ")[0])
                   for b in bridge):
            out.append(f"  -> add it: adjustments=[{{'label': ..., 'value': '{rng[0]}!r{rng[1]}', 'at_valuation_date': true}}] "
                       f"if {crit} is the valuation date")
            try:
                c = crit.strip()
                cv = _as_date(resolve(db, c if "!" in c or db.execute(
                    "SELECT 1 FROM names WHERE lower(name)=lower(?)", (c,)).fetchone() else f"{here}!{c}")[0])
            except ValueError:
                cv = None
            if cv == vd and rng[0] == date_rng[0]:
                fixes.setdefault("add", []).append({"label": _row_label(db, rng[0], rng[1]) or f"{rng[0]}!r{rng[1]}",
                                                    "value": f"{rng[0]}!r{rng[1]}", "at_valuation_date": True})
    return out, fixes


def compute(db: sqlite3.Connection, cashflow, rate, valuation_date, dates: str | None = None,
            timing: str = "end", day_count: str = "actual/actual", terminal_date=None,
            adjustments: list | None = None, compare_to=None, rates: list | None = None,
            fix: bool = True) -> dict:
    """fix: if the result doesn't match compare_to and the workbook's own formula shows why (another cash-flow
    row, a cut-off date, a different convention, an amount taken on the valuation date), redo it with those
    corrections and report them."""
    ranges = [cashflow] if isinstance(cashflow, str) else list(cashflow or [])
    if not ranges:
        raise ValueError("cashflow: give at least one row range")
    rate_v, rate_src = resolve(db, rate)
    rate_v = _num(rate_v)
    if rate_v is None:
        raise ValueError(f"rate {rate!r} is not a number ({rate_src})")
    if rate_v >= 1:  # 13.25 meant as a percentage
        rate_v /= 100
    vd, vd_src = resolve(db, valuation_date)
    vd = _as_date(vd)
    if vd is None:
        raise ValueError(f"valuation_date {valuation_date!r} is not a date ({vd_src})")
    td, td_src = (None, None)
    if terminal_date:
        td, td_src = resolve(db, terminal_date)
        td = _as_date(td)

    flows: dict[int, float] = {}
    rows_used, sheet0, cols0 = [], None, None
    for ref in ranges:
        sheet, row, cols = _row_range(db, ref)
        if sheet0 is None:
            sheet0, cols0 = sheet, cols
        elif (sheet, cols[0], cols[-1]) != (sheet0, cols0[0], cols0[-1]):
            raise ValueError("all cashflow rows must be on the same sheet and span the same columns")
        vals = dict(db.execute("SELECT col, value FROM cells WHERE sheet=? AND row=? AND col BETWEEN ? AND ?",
                               (sheet, row, cols[0], cols[-1])))
        for c in cols:
            v = _num(vals.get(c))
            if v is not None:
                flows[c] = flows.get(c, 0.0) + v
        rows_used.append(f"{sheet}!{_addr(cols[0], row)}:{_addr(cols[-1], row)} {_row_label(db, sheet, row)}".strip())
    ends, ends_src = period_ends(db, sheet0, cols0, dates)
    # Row-total columns (e.g. K = SUM(L:HO)) have no period date of their own, so they drop out here.
    flows = {c: v for c, v in flows.items() if c in ends}

    target = target_src = None
    if compare_to is not None:
        target, target_src = resolve(db, compare_to)
        target = _num(target)

    adj_total, bridge = 0.0, []
    for a in adjustments or []:
        v0 = str(a.get("value")) if isinstance(a, dict) else ""
        is_row = bool(_ROWREF.match(v0.strip()) or ((m_ := _REF.match(v0.strip())) and m_["c2"] and m_["c1"] != m_["c2"]))
        if isinstance(a, dict) and a.get("at_valuation_date") and is_row:  # a row's amount in the valuation-date period
            s_, r_, cs_ = _row_range(db, str(a.get("value")))
            e_, _ = period_ends(db, s_, cs_, dates if s_ == sheet0 else None)
            hit = [c for c in cs_ if e_.get(c) == vd]
            if not hit:
                raise ValueError(f"adjustment {a.get('value')}: no period ending on the valuation date {vd}")
            a = {**a, "value": f"{s_}!{_addr(hit[0], r_)}"}
        v, src = resolve(db, a.get("value") if isinstance(a, dict) else a)
        v = _num(v)
        if v is None:
            raise ValueError(f"adjustment {a!r} is not a number ({src})")
        if isinstance(a, dict) and a.get("source"):
            src = a["source"]
        label = (a.get("label") if isinstance(a, dict) else None) or src
        bridge.append({"label": label, "value": v, "source": src})
        adj_total += v

    combos = [(t, d) for t in TIMINGS for d in DAY_COUNTS] if timing == "auto" else [(timing, day_count)]
    if timing == "auto" and target is None:
        raise ValueError("timing='auto' needs compare_to (the workbook's value to match)")
    tried = []
    for t, d in combos:
        if t not in TIMINGS or d not in DAY_COUNTS:
            raise ValueError(f"timing must be one of {TIMINGS} or 'auto'; day_count one of {DAY_COUNTS}")
        pv, n = _pv(flows, ends, vd, rate_v, t, d, td)
        tried.append({"timing": t, "day_count": d, "pv": pv, "total": pv + adj_total, "periods": n})
    matched = [x for x in tried if target is not None and _close(x["total"], target)]
    # auto keeps a convention only if it reproduces the workbook exactly; a merely "closest" one would hide
    # a wrong input (a missing cut-off date or bridge item) behind a convention change
    best = matched[0] if matched else tried[0]

    in_window = [c for c in sorted(ends) if ends[c] > vd and not (td and ends[c] > td)]
    hint = None
    if not td:  # a terminal date in the workbook that this run discounts past
        for c in candidates(db, "terminal", 3):
            d = _as_date(c.split(" = ")[1][:10])
            n_after = sum(1 for k in in_window if ends[k] > d and flows.get(k)) if d else 0
            if n_after:
                hint = f"{c} - {n_after} non-zero periods after it were included; pass terminal_date if the valuation stops there"
                break
    notes = []
    if compare_to is not None and not (target is not None and _close(best["total"], target)):
        calc = workbook_calc(db, str(compare_to))
        ours = factors(ends, vd, rate_v, best["timing"], best["day_count"], td)
        given = [(_row_range(db, x)[0], _row_range(db, x)[1]) for x in ranges]
        notes, fixes = diagnose(db, calc, given, ours, ends, bridge,
                                lambda t, d: {c: f for c, f in factors(ends, vd, rate_v, t, d, None).items() if f}, vd)
        if fix and fixes:
            again = compute(db, fixes.get("cashflow", ranges), rate, valuation_date, dates,
                            fixes.get("timing", best["timing"]), fixes.get("day_count", best["day_count"]),
                            fixes.get("terminal_date", terminal_date), [*(adjustments or []), *fixes.get("add", [])],
                            compare_to, rates, fix=False)
            if again["compare_to"] is not None and _close(again["total"], again["compare_to"]):
                done = []
                if "cashflow" in fixes:
                    done.append(f"cash flow row {', '.join(ranges)} -> {fixes['cashflow'][0]} (the row the workbook discounts)")
                if "terminal_date" in fixes:
                    done.append(f"terminal_date {fixes['terminal_date']} (the workbook's factors are 0 after it)")
                if "timing" in fixes and (fixes["timing"], fixes["day_count"]) != (timing, day_count):
                    done.append(f"convention {fixes['timing']}-of-period, {fixes['day_count']} (fits the workbook's factors)")
                for a in fixes.get("add", []):
                    done.append(f"added {a['label']} in the valuation-date period (the workbook's SUMIFS)")
                again["corrections"] = done
                again["uncorrected_total"] = best["total"]
                return again
        if calc["tree"]:
            notes += ["workbook formula (first levels):"] + calc["tree"][:8]
    units = {db.execute("SELECT units FROM rows WHERE sheet=? AND row=?", (sh, r)).fetchone() or (None,)
             for sh, r in [(_row_range(db, x)[0], _row_range(db, x)[1]) for x in ranges]}
    units = {u[0] for u in units if u and u[0]}
    sens = [{"rate": r / 100 if r >= 1 else r,
             "total": _pv(flows, ends, vd, r / 100 if r >= 1 else r, best["timing"], best["day_count"], td)[0] + adj_total}
            for r in (rates or [])]
    return {
        "rate": rate_v, "rate_source": rate_src, "valuation_date": vd, "valuation_date_source": vd_src,
        "terminal_date": td, "terminal_date_source": td_src, "rows": rows_used, "dates_source": ends_src,
        "periods": best["periods"], "first_period": ends[in_window[0]] if in_window else None,
        "last_period": ends[in_window[-1]] if in_window else None,
        "undiscounted": sum(flows.get(c, 0.0) for c in in_window),
        "excluded_before": sum(1 for c in ends if ends[c] <= vd and flows.get(c)),
        "timing": best["timing"], "day_count": best["day_count"], "pv": best["pv"],
        "bridge": bridge, "total": best["total"], "tried": tried if len(tried) > 1 else None,
        "terminal_hint": hint, "notes": notes, "matched_any": bool(matched), "units": sorted(units),
        "compare_to": target, "compare_source": target_src, "sensitivity": sens,
    }


# ---- finding inputs ---------------------------------------------------------------------------------------

_RATE_WORDS = re.compile(r"discount\s*rate|\bwacc\b|cost of (equity|capital)|hurdle|required return|\bk[ed]\b", re.I)
_VD_WORDS = re.compile(r"valuation\s*date|val\W*date|\bvd\b", re.I)
_TV_WORDS = re.compile(r"term\w*\W*(val\w*\W*)?date|terminal", re.I)


def candidates(db: sqlite3.Connection, what: str, limit: int = 12) -> list[str]:
    """Cells that look like a discount rate or a valuation date, as text lines "ref = value  (where)"."""
    return [c["text"] for c in candidate_cells(db, what, limit)]


def candidate_cells(db: sqlite3.Connection, what: str, limit: int = 12) -> list[dict]:
    """Cells that look like a discount rate ("rate"), valuation date ("date") or terminal date ("terminal"):
    named ranges first, then labelled rows (their first number or date to the right of the label).
    Each is {"ref", "value", "text"}; ref is the name or Sheet!A1."""
    is_date = lambda v: _as_date(v) is not None
    words, ok = {"rate": (_RATE_WORDS, lambda v: _num(v) is not None and 0 < _num(v) < 1),
                 "date": (_VD_WORDS, is_date), "terminal": (_TV_WORDS, is_date)}[what]
    skip = _TV_WORDS if what == "date" else None  # a "terminal value date" isn't a valuation date
    out, seen = [], set()
    for name, ref in db.execute("SELECT name, ref FROM names"):
        if words.search(name.replace("_", " ")) and not (skip and skip.search(name.replace("_", " "))):
            try:
                v, src = resolve(db, name)
            except ValueError:
                continue
            if ok(v):
                seen.add(src.split(" ")[0])
                out.append({"ref": name, "value": v, "text": f"{name} = {_show(v)}  ({src})"})
    for sheet, row, label in db.execute("SELECT sheet, row, label FROM rows WHERE label IS NOT NULL ORDER BY sheet, row"):
        if not words.search(label) or (skip and skip.search(label)):
            continue
        lay = _layout(db, sheet)
        for col, v in db.execute("SELECT col, value FROM cells WHERE sheet=? AND row=? AND col>? ORDER BY col",
                                 (sheet, row, lay.get("units_col") or lay.get("label_col") or 0)):
            if ok(v):
                ref = f"{sheet}!{_addr(col, row)}"
                if ref not in seen:  # already listed under its name
                    out.append({"ref": ref, "value": v, "text": f"{ref} = {_show(v)}  ({label})"})
                break
        if len(out) >= limit:
            break
    return out[:limit]


_OUT_WORDS = re.compile(r"equity value|enterprise value|\bev\b|\bnpv\b|\bxnpv\b|present value|\bpv\b|"
                        r"total valuation|valuation\b", re.I)


def outputs(db: sqlite3.Connection, limit: int = 25) -> list[str]:
    """Cells that look like a valuation result: labelled like "Enterprise value" / "Equity PV" / "NPV", with a
    formula and a numeric value (the first such cell right of the label), shown with that formula."""
    out = []
    for sheet, row, label in db.execute("SELECT sheet, row, label FROM rows WHERE label IS NOT NULL ORDER BY sheet, row"):
        if not _OUT_WORDS.search(label) or (re.search(r"date|rate|factor|multiple|period|/", label, re.I)
                                            and not re.search(r"value|\bn?pv\b", label, re.I)):
            continue
        lay = _layout(db, sheet)
        for col, f, v in db.execute("SELECT col, formula, value FROM cells WHERE sheet=? AND row=? AND col>? "
                                    "AND formula IS NOT NULL ORDER BY col",
                                    (sheet, row, lay.get("label_col") or 0)):  # units column too: totals sit there
            if _num(v) is not None and _num(v) != 0:
                out.append(f"{sheet}!{_addr(col, row)} = {_num(v):,.1f}  {label}  {f[:140]}")
                break
        if len(out) >= limit:
            break
    return out


def _show(v) -> str:
    d = _as_date(v)
    if d:
        return d.isoformat()
    n = _num(v)
    return f"{n:.4%}" if n is not None and 0 < abs(n) < 1 else f"{n:,.4f}" if n is not None else str(v)


def _money(v: float) -> str:
    return f"{v:,.1f}"


def report(r: dict) -> str:
    out = [f"DCF recomputed in Python from the workbook's saved cash flows (the cash flows themselves aren't recalculated).",
           f"rate {r['rate']:.4%} from {r['rate_source']}",
           f"valuation date {r['valuation_date']} from {r['valuation_date_source']}"]
    if r.get("corrections"):
        out.append(f"CORRECTED to follow the workbook's own formula (your inputs gave {_money(r['uncorrected_total'])}):")
        out += [f"  - {c}" for c in r["corrections"]]
    if r.get("terminal_hint"):
        out.append(f"note: workbook terminal date {r['terminal_hint']}")
    if r["terminal_date"]:
        out.append(f"terminal date {r['terminal_date']} from {r['terminal_date_source']} (later periods left out)")
    out += [f"cash flows: " + "; ".join(r["rows"]),
            f"period ends: {r['dates_source']}",
            f"discounted {r['periods']} periods {r['first_period']} to {r['last_period']}"
            + (f"; {r['excluded_before']} non-zero periods on or before the valuation date left out" if r["excluded_before"] else ""),
            f"convention: {r['timing']}-of-period, {r['day_count']}",
            f"undiscounted total {_money(r['undiscounted'])}",
            f"PV {_money(r['pv'])}"]
    run = r["pv"]
    for b in r["bridge"]:
        run += b["value"]
        src = "" if b["label"] == b["source"] else f" ({b['source']})"
        out.append(f"  {'+' if b['value'] >= 0 else '-'} {b['label']}: {_money(abs(b['value']))}{src} -> {_money(run)}")
    if r["bridge"]:
        out.append(f"total {_money(r['total'])}")
    if r["compare_to"] is not None:
        diff = r["total"] - r["compare_to"]
        rel = abs(diff) / abs(r["compare_to"]) if r["compare_to"] else abs(diff)
        verdict = "MATCHES" if rel < 1e-6 else "close" if rel < 1e-3 else "DOES NOT MATCH"
        out.append(f"workbook {r['compare_source']}: {_money(r['compare_to'])}; difference {diff:,.4f} -> {verdict}")
    if r["tried"] and not r["matched_any"]:
        out.append("no convention reproduces the workbook, so the inputs differ; showing end-of-period, actual/actual")
    for n in r.get("notes") or []:
        out.append(n)
    if r["tried"]:
        out.append("conventions tried: " + "; ".join(f"{t['timing']}/{t['day_count']} {_money(t['total'])}" for t in r["tried"]))
    if r["sensitivity"]:
        out.append("at other rates (same cash flows, same bridge): " +
                   "; ".join(f"{s['rate']:.2%} -> {_money(s['total'])}" for s in r["sensitivity"]))
    out.append(f"units: {', '.join(r['units'])} (from the cash-flow row)" if r.get("units")
               else "units: not labelled on the cash-flow row; don't assume any")
    return "\n".join(out)


def card(r: dict) -> dict:
    """JSON-safe summary of a compute() result for the UI."""
    iso = lambda d: d.isoformat() if d else None
    return {
        "total": r["total"], "pv": r["pv"], "rate": r["rate"], "rate_source": r["rate_source"],
        "valuation_date": iso(r["valuation_date"]), "valuation_date_source": r["valuation_date_source"],
        "terminal_date": iso(r["terminal_date"]), "rows": r["rows"], "periods": r["periods"],
        "first_period": iso(r["first_period"]), "last_period": iso(r["last_period"]),
        "timing": r["timing"], "day_count": r["day_count"], "bridge": r["bridge"],
        "compare_to": r["compare_to"], "compare_source": r["compare_source"],
        "matches": r["compare_to"] is not None and _close(r["total"], r["compare_to"]),
        "corrections": r.get("corrections") or [], "sensitivity": r["sensitivity"], "units": r.get("units") or [],
    }
