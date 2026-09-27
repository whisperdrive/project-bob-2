"""Key facts from last year's valuation report: the reference used to find the valuation in the overlay model.

Three passes, each visible to the person who approves the result:
  1. extract   a model reads the report Markdown (docingest.py) and returns each fact with the page and a
               verbatim quote: target, valuation date, conclusions (ranges and preferred values), assumptions
               (discount rate and its basis, terminal growth or exit / RAB multiple, ...), approaches
  2. check     code, no model: the quote must be on the cited page, every value must be in the quote, the
               number must match the text, and a quote from a table that hasn't been verified is marked
  3. review    a second model call sees the same pages and the facts, and accepts, corrects or rejects each
               one with a reason, and lists key facts that were missed (corrections are checked like 2)
    uv run python bench/reportfacts.py out/docs/<dir>/document.md [model]
"""
import json
import re
import sys

from docingest import numbers

MAX_CHARS = 150_000  # longer reports: send the pages most likely to hold conclusions and assumptions
KEYWORDS = re.compile(r"valuation|discount|wacc|terminal|growth|multiple|rab|conclu|range|preferred|assumption|"
                      r"enterprise value|equity value|methodolog|approach|summary|cost of capital|cpi|inflation", re.I)

KEYS = """identity:   target_name, project_name, valuation_date, basis_of_value (e.g. fair market value),
            interest_valued (e.g. 100% of ordinary equity), currency_units (e.g. A$m), client
conclusion: enterprise_value, equity_value, other values the report concludes on (e.g. unit value), each with
            the preferred value in value_text and the range in low_text / high_text
assumption: discount_rate (basis: e.g. post-tax nominal WACC), terminal_growth_rate, exit_multiple,
            rab_multiple, cpi, tax_rate, forecast_period, net_debt, cost_of_equity, gearing, other key inputs
approach:   primary_approach (e.g. DCF of unlevered free cash flows), cross_check (e.g. EV/EBITDA multiple),
            terminal_value_method (Gordon growth, exit multiple, RAB multiple), discounting_convention
            (mid-period or end of period), cash_flow_basis (nominal / real, pre / post tax)
sensitivity: each figure in a sensitivity table, key e.g. equity_value_sensitivity, with the scenario in basis
            (e.g. "WACC 7.50%, TGR 2.25%"); these are test points for rebuilding the valuation"""

EXTRACT_PROMPT = """You are reading last year's final valuation report for a recurring infrastructure valuation.
Its conclusions and assumptions are the reference used to find the valuation in last year's Excel model, so
extract every datapoint below that the report states. Use these keys where they fit (add others in snake_case):
{keys}

Rules:
- value_text: exactly as printed ("7.25%", "A$2,296.7m", "30 June 2025"); low_text / high_text for a range,
  else "". value: the number in value_text (7.25 for 7.25%, 2296.7 for A$2,296.7m, 20250630 for a date as
  YYYYMMDD) or null for text. unit: "%", "x", "date", "years", "text" or the currency units ("A$m").
- basis: what the figure is on (e.g. "post-tax nominal WACC", "preferred", "real"), else "".
- page: the N of the nearest "<!-- page N -->" marker above the text you used.
- quote: copied verbatim from the document, the shortest sentence or table row that states the value
  (a table row as its cells separated by spaces, without the | characters). Never paraphrase.
- Only facts the document states. If the report gives a figure more than once, use the main statement
  (the executive summary or the assumptions table).

Document:
{doc}"""

REVIEW_PROMPT = """You are the reviewer. Another model extracted the facts below from last year's valuation report
(the document follows). Code has already checked that each quote appears on the cited page and contains the
value; those results are shown. For every fact decide:
- accept: the value, basis and page are right and it is the report's main statement of it
- correct: something is wrong; give the corrected value_text / low_text / high_text / basis / page / quote
  (quote verbatim from the document)
- reject: the document doesn't support it, or it isn't a key fact (say why)
Then list key facts that were missed (target, valuation date, conclusions and their range, discount rate and
basis, terminal growth or exit / RAB multiple, approach), with the same fields. Be strict: numbers must match
the document exactly, including units and whether a value is pre- or post-tax, nominal or real.
A quote must be one continuous piece of the document, so for a figure in a table the quote is its row and the
column it sits under goes in basis: check the column against the table, but don't "correct" a fact only to add
the table's headings to its quote.

Facts:
{facts}

Document:
{doc}"""

_S, _N, _I = {"type": "string"}, {"type": ["number", "null"]}, {"type": ["integer", "null"]}
CATEGORIES = ["identity", "conclusion", "assumption", "approach", "sensitivity"]
_FACT = {"category": {"type": "string", "enum": CATEGORIES},
         "key": _S, "label": _S, "value_text": _S, "low_text": _S, "high_text": _S, "value": _N, "unit": _S,
         "basis": _S, "page": _I, "quote": _S}
EXTRACT_SCHEMA = {"type": "json_schema", "name": "report_facts", "strict": True, "schema": {
    "type": "object", "additionalProperties": False, "required": ["facts", "notes"],
    "properties": {"facts": {"type": "array", "items": {"type": "object", "additionalProperties": False,
                                                        "required": list(_FACT), "properties": _FACT}},
                   "notes": _S}}}
_REV = {"id": {"type": "integer"}, "verdict": {"type": "string", "enum": ["accept", "correct", "reject"]},
        "reason": _S, "value_text": _S, "low_text": _S, "high_text": _S, "basis": _S, "page": _I, "quote": _S}
_MISS = {**_FACT, "why": _S}
REVIEW_SCHEMA = {"type": "json_schema", "name": "facts_review", "strict": True, "schema": {
    "type": "object", "additionalProperties": False, "required": ["reviews", "missing", "summary"],
    "properties": {
        "reviews": {"type": "array", "items": {"type": "object", "additionalProperties": False,
                                               "required": list(_REV), "properties": _REV}},
        "missing": {"type": "array", "items": {"type": "object", "additionalProperties": False,
                                               "required": list(_MISS), "properties": _MISS}},
        "summary": _S}}}


# ---- the document -------------------------------------------------------------------------------------------

def pages(markdown: str) -> dict[int, str]:
    parts = re.split(r"<!-- page (\d+) -->\n", markdown)
    return {int(parts[i]): parts[i + 1] for i in range(1, len(parts) - 1, 2)}


def select(markdown: str, limit: int = MAX_CHARS) -> str:
    """The whole document if it fits, else the first pages plus the pages with the most valuation keywords."""
    if len(markdown) <= limit:
        return markdown
    pg = pages(markdown)
    score = {n: len(KEYWORDS.findall(t)) for n, t in pg.items()}
    keep, used = set(), 0
    for n in sorted(pg, key=lambda n: (n > 3, -score[n], n)):
        if used + len(pg[n]) > limit:
            continue
        keep.add(n)
        used += len(pg[n])
    return "\n\n".join(f"<!-- page {n} -->\n{pg[n]}" for n in sorted(keep))


def _norm(s: str) -> str:
    s = re.sub(r"<!--.*?-->|[|*#`>]", " ", s or "")
    return re.sub(r"\s+", " ", s).strip().lower()


# ---- checks (no model) --------------------------------------------------------------------------------------

def check(f: dict, pg: dict[int, str]) -> dict:
    """Is the fact supported by the text it cites? Returns {"ok", "items": [(ok, message)]}."""
    items = []
    q = _norm(f.get("quote"))
    where = [n for n, t in pg.items() if q and q in _norm(t)]
    if not q:
        items.append((False, "no quote"))
    elif f.get("page") in where:
        items.append((True, f"quote found on page {f['page']}"))
    elif where:
        items.append((False, f"quote is on page {where[0]}, not page {f.get('page')}"))
    else:
        items.append((False, "quote not found in the document"))
    for k in ("value_text", "low_text", "high_text"):
        v = (f.get(k) or "").strip()
        if not v:
            continue
        nums = numbers(v)
        if nums:
            missing = [n for n in nums if n not in numbers(f.get("quote") or "")]
            items.append((not missing, f"{v} in the quote" if not missing else f"{v} is not in the quote"))
        else:
            ok = _norm(v) in q
            items.append((ok, f"'{v}' in the quote" if ok else f"'{v}' is not in the quote"))
    nums = numbers(f.get("value_text") or "")
    if f.get("value") is not None and nums and f.get("unit") != "date":
        want = float(nums[0].rstrip("%"))
        items.append((abs(want - float(f["value"])) < 1e-9 * max(1, abs(want)),
                      "number matches the text" if abs(want - float(f["value"])) < 1e-9 * max(1, abs(want))
                      else f"number {f['value']} doesn't match {f['value_text']}"))
    # A quote taken from a table inherits that table's status (it's in the page's <!-- table id (status) --> marker).
    for tid, status in source_tables(f, pg):
        if status not in ("verified", "approved", "edited"):
            items.append((False, f"from table {tid}, which is {status}: settle the table first"))
    return {"ok": all(ok for ok, _ in items), "items": [{"ok": ok, "text": t} for ok, t in items]}


def source_tables(f: dict, pg: dict[int, str]) -> list[tuple[str, str]]:
    """[(table id, status)] for tables on the cited page whose text contains the quote."""
    q = _norm(f.get("quote"))
    out = []
    for m in re.finditer(r"<!-- table (\S+) \((\w+)\) -->\n(.*?)(?=\n\n(?!\|)|\Z)", pg.get(f.get("page"), ""), re.S):
        if q and q in _norm(m.group(3)):
            out.append((m.group(1), m.group(2)))
    return out


# ---- model passes -------------------------------------------------------------------------------------------

def _call(model: str, prompt: str, schema: dict, purpose: str, on_usage) -> dict:
    from llm import client, create
    r = create(client(interactive=False), model, input=prompt, text={"format": schema}, max_output_tokens=12000)
    if r.usage and on_usage:
        on_usage(model, r.usage, purpose)
    return json.loads(r.output_text)


def extract(markdown: str, model: str, on_usage=None) -> dict:
    return _call(model, EXTRACT_PROMPT.format(keys=KEYS, doc=select(markdown)), EXTRACT_SCHEMA, "facts", on_usage)


def review(markdown: str, facts: list[dict], model: str, on_usage=None) -> dict:
    brief = [{"id": f["id"], **{k: f.get(k) for k in _FACT}, "automatic_checks": [i["text"] for i in f["check"]["items"]]}
             for f in facts]
    return _call(model, REVIEW_PROMPT.format(facts=json.dumps(brief, indent=1), doc=select(markdown)),
                 REVIEW_SCHEMA, "facts-review", on_usage)


def run(markdown: str, model: str = "gpt-6-luna", reviewer_model: str = "gpt-6-sol",
        on_usage=None, progress=None) -> dict:
    """All three passes. Returns {"facts": [...], "notes", "review_summary"}; each fact carries check, review and,
    if the reviewer corrected it, a checked suggestion."""
    progress = progress or (lambda f, m: None)
    pg = pages(markdown)

    def checked(f):
        f["check"] = check(f, pg)
        return f

    progress(0.1, f"Extracting key facts ({model})")
    ext = extract(markdown, model, on_usage)
    facts = [checked({"id": i, "origin": "extractor", **f}) for i, f in enumerate(ext["facts"], start=1)]
    progress(0.55, f"Reviewing {len(facts)} facts ({reviewer_model})")
    rev = review(markdown, facts, reviewer_model, on_usage)
    by_id = {f["id"]: f for f in facts}
    for r in rev["reviews"]:
        f = by_id.get(r["id"])
        if not f:
            continue
        f["review"] = {"verdict": r["verdict"], "reason": r["reason"]}
        if r["verdict"] == "correct":
            fix = {**{k: f.get(k) for k in _FACT}, **{k: r[k] for k in ("value_text", "low_text", "high_text", "basis",
                                                                          "page", "quote") if r[k] not in (None, "")}}
            nums = numbers(fix["value_text"])
            if nums and fix.get("unit") != "date":
                fix["value"] = float(nums[0].rstrip("%"))
            f["review"]["suggestion"] = checked(fix)
    next_id = len(facts) + 1
    for m in rev["missing"]:
        f = checked({"id": next_id, "origin": "reviewer", **{k: m[k] for k in _FACT}})
        f["review"] = {"verdict": "added", "reason": m["why"]}
        facts.append(f)
        next_id += 1
    progress(1.0, "Facts extracted and reviewed")
    return {"facts": facts, "notes": ext.get("notes"), "review_summary": rev.get("summary")}


if __name__ == "__main__":
    from pathlib import Path
    sys.path.insert(0, str(Path(__file__).parent))
    md = Path(sys.argv[1]).read_text(encoding="utf-8")
    model = sys.argv[2] if len(sys.argv) > 2 else "gpt-6-luna"
    res = run(md, model, progress=lambda f, m: print(f"{f:4.0%} {m}"))
    for f in res["facts"]:
        rv = f.get("review", {})
        print(f"{f['id']:>2} {f['category']:10s} {f['key']:24s} {f['value_text']!r:18} p{f['page']} "
              f"check={'ok' if f['check']['ok'] else [i['text'] for i in f['check']['items'] if not i['ok']]} "
              f"review={rv.get('verdict')}: {rv.get('reason', '')[:80]}")
    print("summary:", res["review_summary"])
