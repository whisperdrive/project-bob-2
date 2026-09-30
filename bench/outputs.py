"""The overlay's own outputs, whether or not the report quotes them: the engagement's output schedule.

The report quotes a few of the overlay's figures; the overlay computes many more. An output here is a labelled
row on the overlay sheets with a formula that is either
  - a single figure: formulas in one or two cells, not a row of periods (an equity value, a rate, a multiple, net
    debt, a bridge's line), or
  - the end of a chain: a row of periods no other row on the overlay sheets reads,
and isn't a check (check, error, integrity). Each is classified once, by rules a person confirms:
  conclusion   where a report conclusion sits (the Python overlay's outputs), a DCF (valuation.catalogue), or
               labelled like a value (equity value, enterprise value, NPV, present value, fair value, per unit)
  assumption   where a report assumption sits (the overlay's levers), or labelled like an input to the value (a
               rate, growth, a multiple, CPI, tax, gearing, a premium, a margin)
  working      the rest: an intermediate figure, a row of periods nothing reads
The report's facts annotate the outputs they sit on; a report figure that sits on none is shown apart, for a person
to mark as produced outside the model. Reads the overlay's model.db; no model calls.
"""
import json
import re
import sqlite3
from collections import defaultdict

CONCLUSION = re.compile(r"equity value|enterprise value|\bev\b|\bn?pv\b|x?npv|present value|fair value|valuation\b|"
                        r"value per|per (?:unit|share|security)|\bprice\b", re.I)
ASSUMPTION = re.compile(r"\brate\b|wacc|growth|multiple|\bcpi\b|inflation|\btax\b|gearing|premium|\bbeta\b|margin|"
                        r"cost of|yield|escalat", re.I)  # not "terminal": a terminal value is a part of a value
CHECK = re.compile(r"check|error|integrity", re.I)
CLASSES = ("conclusion", "assumption", "working")
FIGURE_CELLS = 3  # formulas in at most this many cells: a single figure (or a low / mid / high), not a row of periods


def _a1(col: int, row: int) -> str:
    s = ""
    while col:
        col, r = divmod(col - 1, 26)
        s = chr(65 + r) + s
    return f"{s}{row}"


def _num(v):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def detect(db: sqlite3.Connection, sheets: list[str], levers: list[dict] | None = None,
           outputs: list[dict] | None = None, anchors: set[str] | None = None) -> list[dict]:
    """The schedule as detected: [{"row", "cell", "label", "kind": "figure" | "series", "value", "periods",
    "class", "why", "fact"}], by sheet and row. levers / outputs: the Python overlay's (overlay.levers_and_outputs:
    where the report's assumptions and conclusions sit); anchors: the DCF cells (valuation.catalogue)."""
    sheets = list(sheets or [])
    if not sheets:
        return []
    q = ",".join("?" * len(sheets))
    label_col = {s: json.loads(lay or "{}").get("label_col") or 0
                 for s, lay in db.execute(f"SELECT sheet, layout FROM sheets WHERE sheet IN ({q})", sheets)}
    formulas = defaultdict(list)  # (sheet, row) -> [(col, value)] right of the label
    for s, r, c, v in db.execute(f"SELECT sheet, row, col, value FROM cells WHERE formula IS NOT NULL AND sheet IN ({q}) "
                                 "ORDER BY sheet, row, col", sheets):
        if c > label_col.get(s, 0):
            formulas[(s, r)].append((c, v))
    read = set()  # rows another overlay row reads
    for s, r, ds, dr in db.execute(f"SELECT src_sheet, src_row, dst_sheet, dst_row FROM edges WHERE kind IN "
                                   f"('direct', 'offset', 'active') AND src_sheet IN ({q}) AND dst_sheet IN ({q})",
                                   sheets + sheets):
        if (s, r) != (ds, dr):
            read.add((ds, dr))
    at = {}  # (sheet, row) -> the report fact sitting there, from the overlay's levers and outputs
    for o in outputs or []:
        if o.get("fact_id") is not None or o.get("key"):
            m = re.match(r"^(.+)!([A-Z]+)(\d+)$", o["cell"])
            if m:
                at.setdefault((m[1], int(m[3])), {"key": o.get("key"), "category": "conclusion", "report": o.get("report"),
                                                  "tie": (o.get("tie") or {}).get("ok"), "cell": o["cell"]})
    for lv in levers or []:
        m = re.match(r"^(.+)!([A-Z]+)(\d+)$", lv["cell"])
        if m:
            at.setdefault((m[1], int(m[3])), {"key": lv.get("key"), "category": "assumption", "report": lv.get("report"),
                                              "tie": None, "cell": lv["cell"]})
    anchors = anchors or set()
    out = []
    for s, r, label in db.execute(f"SELECT sheet, row, label FROM rows WHERE sheet IN ({q}) AND label IS NOT NULL "
                                  "AND label != '' ORDER BY sheet, row", sheets):
        label = " ".join(label.split())  # "Roll \nforward": a label broken over lines in its cell
        cells = formulas.get((s, r))
        nums = [(c, v) for c, v in cells or [] if _num(v) is not None]
        if not nums or CHECK.search(label):
            continue
        figure = len(cells) <= FIGURE_CELLS
        if not figure and (s, r) in read:
            continue  # a row of periods another row reads: a step on the way, not an output
        fact = at.get((s, r))
        cell = f"{s}!{_a1(nums[0][0], r)}"
        if fact:
            cls, why = fact["category"], f"the report's {fact['key'] or 'figure'} sits here"
        elif any(f"{s}!{_a1(c, r)}" in anchors for c, _ in nums):
            cls, why = "conclusion", "a DCF (its value recomputes from its cash flows and factors)"
        elif figure and CONCLUSION.search(label):
            cls, why = "conclusion", "labelled like a value"
        elif figure and ASSUMPTION.search(label):
            cls, why = "assumption", "labelled like an input to the value"
        else:
            cls, why = "working", ("a figure on the way to others" if figure and (s, r) in read else
                                   "a figure nothing else on the overlay sheets reads" if figure else
                                   "a row of periods nothing else on the overlay sheets reads")
        out.append({"row": f"{s}!r{r}", "cell": cell if figure else None, "label": label,
                    "kind": "figure" if figure else "series", "value": nums[0][1] if figure else None,
                    "periods": None if figure else len(nums), "read": (s, r) in read, "class": cls, "why": why,
                    "fact": fact, **({} if figure else {"range": f"{s}!{_a1(nums[0][0], r)}:{_a1(nums[-1][0], r)}",
                                                        "total": sum(v for _, v in nums)})})
    return out


def outside(facts: list[dict], schedule: list[dict], levers: list[dict] | None = None) -> list[dict]:
    """The report's conclusions and assumptions that sit neither on an output of the schedule nor on an input of the
    overlay (a lever: the discount rate is typed in, not computed): produced outside the model, or not found.
    [{"key", "label", "category", "value_text", "page"}]."""
    here = {o["fact"]["key"] for o in schedule if o.get("fact")} | {lv.get("key") for lv in levers or []}
    return [{"key": f["key"], "label": f.get("label") or f["key"], "category": f["category"],
             "value_text": f.get("value_text"), "page": f.get("page")}
            for f in facts if f.get("category") in ("conclusion", "assumption") and f.get("key") and f["key"] not in here]


def apply(schedule: list[dict], mine: dict) -> list[dict]:
    """The classes as they stand: a person's ({"classes": {"Sheet!r12": "working"}}), else those carried from last
    year's confirmed schedule (mine["carried"]["classes"]), else the detected ones. Each output says which ("source":
    "you" | "carried" | "detected") and, where a schedule was carried, whether it's new this year."""
    classes = mine.get("classes") or {}
    carried = (mine.get("carried") or {}).get("classes")
    out = []
    for o in schedule:
        set_, got = classes.get(o["row"]), (carried or {}).get(o["row"])
        cls, source = (set_, "you") if set_ in CLASSES else (got, "carried") if got in CLASSES else (o["class"], "detected")
        out.append({**o, "detected": o["class"], "class": cls, "set": source == "you", "source": source,
                    **({"new": o["row"] not in carried} if carried is not None else {})})
    return out


def carry(previous: list[dict], old_db: str, new_db: str, sheets: list[str]) -> dict:
    """Last year's confirmed schedule onto this year's overlay: each of its outputs found again by its label (the
    same label on the same sheet, its n-th occurrence, else the nearest of its rows: overlay.RowMap), with the class
    it ended with. A row found only by its place (its label gone) isn't followed: another line item may sit there
    now. Nor is one on a sheet that isn't an overlay sheet this year, or one a second row of last year's already
    took. -> {"classes": {this year's row: class}, "missing": [{row, label, class, why}]}."""
    import overlay as ovmod
    a, b = ovmod.Workbook(old_db), ovmod.Workbook(new_db)
    try:
        rm = ovmod.RowMap(a, b)
        classes, taken, missing = {}, {}, []
        for o in previous:
            m = re.match(r"^(.+)!r(\d+)$", o["row"])
            if not m:
                continue
            s, r = m[1], int(m[2])
            why = None
            if s not in sheets:
                why = f"sheet {s} isn't an overlay sheet this year"
            else:
                r2, how = rm.match(s, r)
                if not r2 or not how.startswith("same label"):
                    why = "its label isn't on the sheet this year"
                elif f"{s}!r{r2}" in taken:
                    why = f"{s}!r{r2} already took last year's {taken[f'{s}!r{r2}']}"
                else:
                    taken[f"{s}!r{r2}"] = o["row"]
                    classes[f"{s}!r{r2}"] = o["class"]
            if why:
                missing.append({"row": o["row"], "label": o.get("label"), "class": o["class"], "why": why})
        return {"classes": classes, "missing": missing}
    finally:
        a.close()
        b.close()
