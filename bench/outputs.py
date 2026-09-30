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
    """A person's classes over the detected ones ({"classes": {"Sheet!r12": "working"}})."""
    classes = mine.get("classes") or {}
    out = []
    for o in schedule:
        set_ = classes.get(o["row"])
        out.append({**o, "detected": o["class"], "class": set_ if set_ in CLASSES else o["class"], "set": set_ in CLASSES})
    return out
