"""Key facts from last year's valuation report: the reference used to find the valuation in the overlay model.

Four passes, each visible to the person who approves the result:
  1. extract   a model reads the report Markdown (docingest.py) and returns each fact with the page and a
               verbatim quote: target, valuation date, conclusions (ranges and preferred values), assumptions
               (discount rate and its basis, terminal growth or exit / RAB multiple, ...), approaches
  2. check     code, no model: the quote must be on the cited page, every value must be in the quote, the
               number must match the text, and a quote from a table that hasn't been verified is marked
  3. review    a second model call sees the same pages and the facts, and accepts, corrects or rejects each
               one with a reason, and lists key facts that were missed (corrections are checked like 2)
  4. resolve   the review and remediation loop: for every fact still open (failed checks, a correction or
               rejection, a fact the reviewer added) the extractor revises, keeps or withdraws it with a reason,
               code re-checks, and the reviewer accepts or objects again; up to MAX_ROUNDS. Facts both agree on are
               "agreed" (a person still approves); what they can't settle is escalated with both positions
Every prompt carries the rules learned from earlier loops (lessons.py), and the loop's episodes feed new ones.
    uv run python bench/reportfacts.py out/docs/<dir>/document.md [model]
"""
import json
import re
import sys

import lessons
from docingest import numbers

MAX_CHARS = 150_000  # longer reports: send the pages most likely to hold conclusions and assumptions
MAX_ROUNDS = 3       # review and remediation rounds before a fact goes to a person
CHECK_VERSION = 2    # bump when check() changes: existing facts are checked again once (2: spacing-tolerant, waivers)
LOOP_CHARS = 60_000  # the loop sends only the pages the open facts cite, and their neighbours
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

FIELDS = """- value_text: exactly as printed ("7.25%", "A$2,296.7m", "30 June 2025"); low_text / high_text for a range,
  else "". value: the number in value_text (7.25 for 7.25%, 2296.7 for A$2,296.7m, 20250630 for a date as
  YYYYMMDD) or null for text. unit: "%", "x", "date", "years", "text" or the currency units ("A$m").
- basis: what the figure is on (e.g. "post-tax nominal WACC", "preferred", "real"), else "".
- page: the N of the nearest "<!-- page N -->" marker above the text you used.
- quote: copied verbatim from the document, the shortest sentence or table row that states the value
  (a table row as its cells separated by spaces, without the | characters). Never paraphrase."""

EXTRACT_PROMPT = """You are reading last year's final valuation report for a recurring infrastructure valuation.
Its conclusions and assumptions are the reference used to find the valuation in last year's Excel model, so
extract every datapoint below that the report states. Use these keys where they fit (add others in snake_case):
{keys}

Rules:
{fields}
- Only facts the document states. If the report gives a figure more than once, use the main statement
  (the executive summary or the assumptions table).

Rules learned from earlier reviews (apply where relevant):
{rules}

Document:
{doc}"""

REVIEW_PROMPT = """You are the reviewer. Another model extracted the facts below from last year's valuation report
(the document follows). Code has already checked that each quote appears on the cited page and contains the
value; those results are shown. For every fact decide:
- accept: the value, basis and page are right and it is the report's main statement of it
- correct: something is wrong; give the corrected value_text / low_text / high_text / basis / page / quote
  (quote verbatim from the document)
- reject: the document doesn't support it, or it isn't a key fact (say why)
Then list key facts that were missed, with the same fields. The reference wants every datapoint below that the
report states, identity included (the project name and the client are wanted even though they aren't valuation
figures): don't reject a fact only because it isn't a number.
{keys} Be strict: numbers must match
the document exactly, including units and whether a value is pre- or post-tax, nominal or real.
A quote must be one continuous piece of the document, so for a figure in a table the quote is its row and the
column it sits under goes in basis: check the column against the table, but don't "correct" a fact only to add
the table's headings to its quote.

Rules learned from earlier reviews (apply where relevant):
{rules}

Facts:
{facts}

Document:
{doc}"""

FIX_PROMPT = """You extracted key facts from last year's valuation report. A reviewer and automatic checks raised the
issues below. For each one decide:
- revise: the fact needs changing; give it in full, with the quote copied verbatim from the pages below
- keep: it is right as it stands; say why, pointing to the text
- withdraw: it shouldn't be in the reference (the report doesn't support it, or it duplicates another fact); every
  datapoint in the brief below is wanted, identity included (project name, client)
A fact the reviewer added becomes yours if you keep or revise it; withdraw it if you disagree. The automatic
checks need the quote to be one continuous piece of the cited page that contains the value, and the number to
match the text. The brief:
{keys}
Field conventions:
{fields}

Rules learned from earlier reviews (list the IDs you apply in rules_applied):
{rules}

Issues:
{issues}

Pages:
{doc}"""

VERIFY_PROMPT = """You are the reviewer of key facts extracted from last year's valuation report. You, or the automatic
checks, raised the issues below and the extractor has answered each one. For each item decide:
- accept: the fact as it now stands is right and a key fact (for a withdrawal: it should indeed go)
- object: it is still wrong or shouldn't go; say why, and give the corrected value_text / low_text / high_text /
  basis / page / quote (quote verbatim) if you can, else leave them ""
Be strict about numbers, units, basis (pre or post tax, nominal or real) and the page. The automatic checks on
the fact as it now stands are shown: don't accept a fact whose checks fail unless you say why the check is wrong.
The reference wants every datapoint below that the report states, identity included (project name, client), so
a withdrawal is right only when the report doesn't support the fact or it duplicates another:
{keys}

Rules learned from earlier reviews (list the IDs you apply in rules_applied):
{rules}

Items:
{items}

Pages:
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
_IDS = {"type": "array", "items": _S}
_FIXR = {"id": {"type": "integer"}, "action": {"type": "string", "enum": ["revise", "keep", "withdraw"]}, "reason": _S,
         **_FACT, "rules_applied": _IDS}
FIX_SCHEMA = {"type": "json_schema", "name": "facts_fix", "strict": True, "schema": {
    "type": "object", "additionalProperties": False, "required": ["responses"],
    "properties": {"responses": {"type": "array", "items": {"type": "object", "additionalProperties": False,
                                                            "required": list(_FIXR), "properties": _FIXR}}}}}
_VER = {"id": {"type": "integer"}, "verdict": {"type": "string", "enum": ["accept", "object"]}, "reason": _S,
        "value_text": _S, "low_text": _S, "high_text": _S, "basis": _S, "page": _I, "quote": _S, "rules_applied": _IDS}
VERIFY_SCHEMA = {"type": "json_schema", "name": "facts_verify", "strict": True, "schema": {
    "type": "object", "additionalProperties": False, "required": ["verdicts"],
    "properties": {"verdicts": {"type": "array", "items": {"type": "object", "additionalProperties": False,
                                                           "required": list(_VER), "properties": _VER}}}}}
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

def _squash(s: str) -> str:
    return re.sub(r"\s+", "", _norm(s))


def _in_squashed(num: str, text: str) -> bool:
    """A number (as numbers() gives it: "5223.0", "-70.0", "7.25%") in text with all spacing ignored, reading glued
    figures by the number's own shape: 5,223.0 is in "5 , 2 2 3 . 0 5 , 2 2 3 . 0" (cells run together)."""
    neg, pct = num.startswith("-"), num.endswith("%")
    core = num.lstrip("-").rstrip("%")
    d = len(core.split(".")[1]) if "." in core else 0
    frac = rf"\.\d{{{d}}}" if d else ""
    shape = re.compile(rf"\d{{1,3}}(?:,\d{{3}})+{frac}|\d+{frac}")
    flat = re.sub(r"\s+", "", text or "")
    for m in shape.finditer(flat):
        if abs(float(m.group(0).replace(",", "")) - float(core)) > 1e-9 * max(1, abs(float(core))):
            continue
        before, after = flat[max(0, m.start() - 1):m.start()], flat[m.end():m.end() + 1]
        if neg != (before in ("(", "-", "−", "–")):  # a bracket or minus in front: negative, both ways
            continue
        if pct and after != "%":
            continue
        return True
    return False


def check(f: dict, pg: dict[int, str]) -> dict:
    """Is the fact supported by the text it cites? Returns {"ok", "items": [{"ok", "text"}]}.
    Strict first; where a report's text layer spaces letters oddly or runs table cells together, a match that
    ignores spacing still counts, and says so. A check an arbiter waived (f["waivers"]) counts, with its note."""
    items = []
    q = _norm(f.get("quote"))
    where = [n for n, t in pg.items() if q and q in _norm(t)]
    loose = [] if where or not q else [n for n, t in pg.items() if _squash(f.get("quote")) in _squash(t)]
    if not q:
        items.append((False, "no quote"))
    elif f.get("page") in where:
        items.append((True, f"quote found on page {f['page']}"))
    elif f.get("page") in loose:
        items.append((True, f"quote found on page {f['page']} (spacing ignored)"))
    elif where or loose:
        items.append((False, f"quote is on page {(where or loose)[0]}, not page {f.get('page')}"))
    else:
        items.append((False, "quote not found in the document"))
    for k in ("value_text", "low_text", "high_text"):
        v = (f.get(k) or "").strip()
        if not v:
            continue
        nums = numbers(v)
        if nums:
            missing = [n for n in nums if n not in numbers(f.get("quote") or "")]
            if not missing:
                items.append((True, f"{v} in the quote"))
            elif all(_in_squashed(n, f.get("quote") or "") for n in missing):
                items.append((True, f"{v} in the quote (spacing ignored)"))
            else:
                items.append((False, f"{v} is not in the quote"))
        elif _norm(v) in q:
            items.append((True, f"'{v}' in the quote"))
        elif _squash(v) and _squash(v) in _squash(f.get("quote")):
            items.append((True, f"'{v}' in the quote (spacing ignored)"))
        else:
            items.append((False, f"'{v}' is not in the quote"))
    nums = numbers(f.get("value_text") or "")
    if f.get("value") is not None and nums and f.get("unit") != "date":
        want = float(nums[0].rstrip("%"))
        items.append((abs(want - float(f["value"])) < 1e-9 * max(1, abs(want)),
                      "number matches the text" if abs(want - float(f["value"])) < 1e-9 * max(1, abs(want))
                      else f"number {f['value']} doesn't match {f['value_text']}"))
    # A quote taken from a table inherits that table's status (it's in the page's <!-- table id (status) --> marker).
    for tid, status in source_tables(f, pg):
        if status == "unread":  # quoted from the PDF's own text while the table waits to be read: that text is real
            items.append((True, f"from table {tid}, still being read; the quote is the PDF's own text"))
        elif status not in ("verified", "approved", "edited", "resolved"):
            items.append((False, f"from table {tid}, which is {status}: settle the table first"))
    waived = {w["check"]: w for w in f.get("waivers") or []}
    items = [(True, f"{t} — waived by {waived[t]['by']}: {waived[t]['note']}") if not ok and t in waived else (ok, t)
             for ok, t in items]
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
    """One structured call. A reply that isn't valid JSON (a model sometimes runs on in whitespace until it hits
    the output limit) is asked for once more before the step fails."""
    from llm import client, create
    for attempt in (1, 2):
        r = create(client(interactive=False), model, input=prompt, text={"format": schema}, max_output_tokens=12000,
                   purpose=purpose if attempt == 1 else f"{purpose} (again: the first reply wasn't valid JSON)")
        if r.usage and on_usage:
            on_usage(model, r.usage, purpose)
        try:
            return json.loads(r.output_text)
        except json.JSONDecodeError:
            if attempt == 2:
                raise
    raise AssertionError("unreachable")


def extract(markdown: str, model: str, on_usage=None) -> dict:
    return _call(model, EXTRACT_PROMPT.format(keys=KEYS, fields=FIELDS, rules=lessons.rules_text("facts"),
                                              doc=select(markdown)), EXTRACT_SCHEMA, "facts", on_usage)


def review(markdown: str, facts: list[dict], model: str, on_usage=None) -> dict:
    brief = [{"id": f["id"], **{k: f.get(k) for k in _FACT}, "automatic_checks": [i["text"] for i in f["check"]["items"]]}
             for f in facts]
    return _call(model, REVIEW_PROMPT.format(facts=json.dumps(brief, indent=1), rules=lessons.rules_text("facts"), keys=KEYS,
                                             doc=select(markdown)), REVIEW_SCHEMA, "facts-review", on_usage)


# ---- the review and remediation loop --------------------------------------------------------------------------

def cited_pages(markdown: str, page_numbers, limit: int = LOOP_CHARS) -> str:
    """The pages the open facts cite, with their neighbours (then without them, if that's too long)."""
    pg = pages(markdown)
    cited = {n for n in page_numbers if n in pg}
    for want in ({m for n in cited for m in (n - 1, n, n + 1)}, cited):
        keep = sorted(n for n in want if n in pg)
        text = "\n\n".join(f"<!-- page {n} -->\n{pg[n]}" for n in keep)
        if keep and len(text) <= limit:
            return text
    return select(markdown, limit)


def _quote_pages(f: dict, pg: dict[int, str]) -> list[int]:
    """Pages where the fact's quote actually is (a wrong page number is a common slip)."""
    q = _norm(f.get("quote"))
    return [n for n, t in pg.items() if q and q in _norm(t)]


def _fields(f: dict) -> dict:
    return {k: f.get(k) for k in _FACT}


def _failed(f: dict) -> list[str]:
    return [i["text"] for i in (f.get("check") or {}).get("items", []) if not i["ok"]]


def _open_issue(f: dict) -> dict | None:
    """What's unsettled about a fact after review: failed checks, the reviewer's correction or rejection, or a fact
    the reviewer added that the extractor hasn't confirmed. None if both sides and the checks agree."""
    rv = f.get("review") or {}
    if rv.get("verdict") == "accept" and not _failed(f):
        return None
    sug = rv.get("suggestion")
    return {"verdict": rv.get("verdict"), "reason": rv.get("reason") or "",
            "correction": {k: sug.get(k) for k in ("value_text", "low_text", "high_text", "basis", "page", "quote")} if sug else None}


# ---- the arbiter: a third model for what the loop couldn't settle --------------------------------------------------

ARBITER_PROMPT = """You are an independent arbiter for the key facts taken from last year's valuation report. An
extractor and a reviewer (two other models) went back and forth on each fact below and couldn't settle it: an
automatic check kept failing, or they disagreed on a detail. Read the cited pages and decide each one:
- waive: the fact is right and a failing automatic check is wrong on a clerical point (spacing, a table row
  copied into the quote, separators, formatting). Copy the failing check's text exactly into check, and give a
  one-sentence note for the file saying why it is clerical.
- use_correction: the reviewer's latest correction is right.
- keep: the extractor's version is right as it stands.
- escalate: a person should decide: the report is ambiguous, the figure may really be wrong, or you can't see
  the value in the cited text.
Only waive or pick a version if you can see the value in the page text yourself. When unsure, escalate; a
person will see your note. The brief (what the reference wants):
{keys}

Items:
{items}

Pages:
{doc}"""
_ARB = {"id": {"type": "integer"}, "decision": {"type": "string", "enum": ["waive", "use_correction", "keep", "escalate"]},
        "check": _S, "note": _S}
ARBITER_SCHEMA = {"type": "json_schema", "name": "arbiter", "strict": True, "schema": {
    "type": "object", "additionalProperties": False, "required": ["decisions"],
    "properties": {"decisions": {"type": "array", "items": {"type": "object", "additionalProperties": False,
                                                            "required": list(_ARB), "properties": _ARB}}}}}


def arbitrate(markdown: str, facts: list[dict], model: str, on_usage=None) -> int:
    """An independent model decides facts the loop escalated (updated in place). A waiver is kept on the fact
    (f["waivers"]) so every later check honours it and shows its note; nothing is settled unless the checks then
    pass. Returns how many it settled; the rest stay escalated, with the arbiter's note."""
    if not facts:
        return 0
    pg = pages(markdown)
    items = []
    for f in facts:
        sug = (f.get("review") or {}).get("suggestion")
        items.append({"id": f["id"], "fact": _fields(f), "checks_failed": _failed(f),
                      "open_point": (f.get("agent") or {}).get("open", {}).get("reason"),
                      "reviewer_correction": {k: sug.get(k) for k in _FACT} if sug else None,
                      "rounds": [{"extractor": t.get("extractor"), "reviewer": t.get("reviewer")}
                                 for t in (f.get("agent") or {}).get("thread", [])[-3:]]})
    cites = [f.get("page") for f in facts] + [n for f in facts for n in _quote_pages(f, pg)]
    out = _call(model, ARBITER_PROMPT.format(keys=KEYS, items=json.dumps(items, indent=1, ensure_ascii=False),
                                             doc=cited_pages(markdown, [c for c in cites if c] or [1])),
                ARBITER_SCHEMA, "facts-arbiter", on_usage)
    by_id, settled = {f["id"]: f for f in facts}, 0
    for dec in out["decisions"]:
        f = by_id.get(dec["id"])
        if not f:
            continue
        a = f["agent"]
        entry = {"round": "arbiter", "arbiter": {"by": model, "decision": dec["decision"], "check": dec["check"],
                                                 "note": dec["note"]}}
        a["thread"].append(entry)
        ok = False
        if dec["decision"] == "waive" and dec["check"]:
            waived = {**f, "waivers": (f.get("waivers") or []) + [{"check": dec["check"].strip(), "by": model,
                                                                 "note": dec["note"]}]}
            chk = check(waived, pg)
            if chk["ok"]:
                f["waivers"], f["check"], ok = waived["waivers"], chk, True
        elif dec["decision"] == "use_correction":
            sug = (f.get("review") or {}).get("suggestion")
            if sug:
                chk = check({**sug, "waivers": f.get("waivers")}, pg)
                if chk["ok"]:
                    f.update({k: sug.get(k) for k in _FACT})
                    f["check"], ok = chk, True
        elif dec["decision"] == "keep":
            chk = check(f, pg)
            f["check"], ok = chk, chk["ok"]
        if ok:
            a.update(status="agreed", round="arbiter", waivers=f.get("waivers") or [])
            a.pop("open", None)
            settled += 1
        else:
            entry["arbiter"]["outcome"] = "escalated" if dec["decision"] == "escalate" else "didn't pass the checks"
            a["open"] = {**(a.get("open") or {}), "arbiter": dec["note"]}
    return settled


def resolve(markdown: str, facts: list[dict], model: str, reviewer_model: str, on_usage=None, progress=None,
            rounds: int = MAX_ROUNDS, arbiter_model: str | None = None) -> dict:
    """The loop on reviewed facts (updated in place: each gets "agent" = {status, round, thread}). A fact is
    agreed when the reviewer accepts it and the checks pass, withdrawn when the reviewer accepts the extractor's
    withdrawal, and escalated when rounds run out. Returns {"summary", "episodes"} (episodes feed lessons.py)."""
    progress = progress or (lambda f, m: None)
    pg = pages(markdown)
    by_id = {f["id"]: f for f in facts}
    issues = {}
    for f in facts:
        iss = _open_issue(f)
        rv = f.get("review") or {}
        thread = [{"round": 0, "extractor": {"action": "extract" if f.get("origin") != "reviewer" else "-"},
                   "reviewer": {"verdict": rv.get("verdict"), "reason": rv.get("reason")}, "checks_failed": _failed(f)}]
        f["agent"] = {"status": "agreed" if iss is None else "open", "round": 0, "thread": thread}
        if iss:
            issues[f["id"]] = iss
    k = 0
    while issues and k < rounds:
        k += 1
        progress((k - 1) / rounds, f"Fact review loop, round {k}: {len(issues)} fact(s) open ({model} fixes, {reviewer_model} checks)")
        cites = [by_id[i].get("page") for i in issues] + [(iss["correction"] or {}).get("page") for iss in issues.values()] \
            + [n for i in issues for n in _quote_pages(by_id[i], pg)]
        doc = cited_pages(markdown, [c for c in cites if c])
        brief = [{"id": i, "added_by_reviewer": by_id[i].get("origin") == "reviewer", "fact": _fields(by_id[i]),
                  "checks_failed": _failed(by_id[i]), "reviewer": iss} for i, iss in issues.items()]
        fix = _call(model, FIX_PROMPT.format(fields=FIELDS, keys=KEYS, rules=lessons.rules_text("facts"),
                                             issues=json.dumps(brief, indent=1, ensure_ascii=False), doc=doc),
                    FIX_SCHEMA, "facts-fix", on_usage)
        answers = {r["id"]: r for r in fix["responses"] if r["id"] in issues}
        lessons.applied([x for r in answers.values() for x in r["rules_applied"]])
        pending, items = {}, []
        for i in issues:
            f, r = by_id[i], answers.get(i) or {"action": "keep", "reason": "(no answer)", "rules_applied": []}
            cand = _fields(f)
            if r["action"] == "revise":
                cand.update({x: r[x] for x in _FACT if x in r and r[x] not in (None, "") or x in ("low_text", "high_text")
                             and x in r})
                cand["category"] = cand["category"] if cand.get("category") in CATEGORIES else f["category"]
                cand["key"] = cand.get("key") or f["key"]
                nums = numbers(cand.get("value_text") or "")
                if nums and cand.get("unit") != "date":
                    cand["value"] = float(nums[0].rstrip("%"))
            cand["check"] = check(cand, pg) if r["action"] == "revise" else f["check"]
            pending[i] = (r, cand)
            items.append({"id": i, "you_said": issues[i], "extractor": {"action": r["action"], "reason": r["reason"]},
                          "fact_now": _fields(cand), "automatic_checks_now": [x["text"] for x in cand["check"]["items"]]})
        cites = [c.get("page") for _, c in pending.values()] + [n for _, c in pending.values() for n in _quote_pages(c, pg)]
        ver = _call(reviewer_model, VERIFY_PROMPT.format(rules=lessons.rules_text("facts"), keys=KEYS, items=json.dumps(
            items, indent=1, ensure_ascii=False), doc=cited_pages(markdown, [c for c in cites if c] or [1])),
            VERIFY_SCHEMA, "facts-verify", on_usage)
        verdicts = {v["id"]: v for v in ver["verdicts"]}
        lessons.applied([x for v in verdicts.values() for x in v["rules_applied"]])
        issues = {}
        for i, (r, cand) in pending.items():
            f = by_id[i]
            v = verdicts.get(i) or {"verdict": "object", "reason": "the reviewer gave no verdict"}
            ok = cand["check"]["ok"]
            f["agent"]["thread"].append({"round": k, "extractor": {"action": r["action"], "reason": r["reason"]},
                                         "fact": _fields(cand) if r["action"] == "revise" else None,
                                         "reviewer": {"verdict": v["verdict"], "reason": v["reason"]},
                                         "checks_failed": [x["text"] for x in cand["check"]["items"] if not x["ok"]]})
            if r["action"] == "revise":  # the discussion moves on to the latest version
                f.update(_fields(cand))
                f["check"] = cand["check"]
            if v["verdict"] == "accept" and r["action"] == "withdraw":
                f["agent"].update(status="withdrawn", round=k)
            elif v["verdict"] == "accept" and ok:
                f["agent"].update(status="agreed", round=k)
            else:
                corr = {x: v.get(x) for x in ("value_text", "low_text", "high_text", "basis", "page", "quote")
                        if v.get(x) not in (None, "")}
                reason = v["reason"] if v["verdict"] == "object" else \
                    "accepted by the reviewer, but the automatic checks still fail: " + "; ".join(_failed(cand))
                issues[i] = {"verdict": "object", "reason": reason, "correction": corr or None,
                             "accepted": v["verdict"] == "accept"}
    for i, iss in issues.items():  # rounds ran out: a person decides, with the reviewer's latest correction to hand
        f = by_id[i]
        f["agent"].update(status="escalated", round=k, open=iss)
        if iss.get("correction"):
            fix = {**_fields(f), **iss["correction"]}
            nums = numbers(fix.get("value_text") or "")
            if nums and fix.get("unit") != "date":
                fix["value"] = float(nums[0].rstrip("%"))
            fix["check"] = check(fix, pg)
            f["review"] = {**(f.get("review") or {}), "suggestion": fix}
    arbitrated = 0
    if arbiter_model and issues:
        progress(0.95, f"Arbiter ({arbiter_model}) on {len(issues)} fact(s) the loop couldn't settle")
        arbitrated = arbitrate(markdown, [by_id[i] for i in issues], arbiter_model, on_usage)
    count = lambda st: sum(f["agent"]["status"] == st for f in facts)
    summary = {"rounds": k, "agreed": count("agreed"), "withdrawn": count("withdrawn"), "escalated": count("escalated"),
               "settled_in_loop": sum(f["agent"]["status"] in ("agreed", "withdrawn") and f["agent"]["round"] not in (0, None)
                                      for f in facts), "arbitrated": arbitrated}
    episodes = [{"category": f["category"], "key": f["key"], "added_by_reviewer": f.get("origin") == "reviewer",
                 "thread": f["agent"]["thread"], "outcome": f["agent"]["status"]}
                for f in facts if len(f["agent"]["thread"]) > 1]
    progress(1.0, f"Fact review loop: {summary['agreed']} agreed, {summary['withdrawn']} withdrawn, {summary['escalated']} for you")
    return {"summary": summary, "episodes": episodes}


def run(markdown: str, model: str = "gpt-6-luna", reviewer_model: str = "gpt-6-sol",
        on_usage=None, progress=None, loop_rounds: int = MAX_ROUNDS, arbiter_model: str | None = None) -> dict:
    """All four passes. Returns {"facts": [...], "notes", "review_summary", "loop"}; each fact carries check,
    review (with a checked suggestion if the reviewer corrected it) and agent (the loop's outcome and thread)."""
    progress = progress or (lambda f, m: None)
    pg = pages(markdown)

    def checked(f):
        f["check"] = check(f, pg)
        return f

    progress(0.1, f"Extracting key facts ({model})")
    ext = extract(markdown, model, on_usage)
    facts = [checked({"id": i, "origin": "extractor", **f}) for i, f in enumerate(ext["facts"], start=1)]
    progress(0.45, f"Reviewing {len(facts)} facts ({reviewer_model})")
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
    loop = resolve(markdown, facts, model, reviewer_model, on_usage,
                   lambda f, m: progress(0.75 + 0.25 * f, m), arbiter_model=arbiter_model) if loop_rounds else None
    progress(1.0, "Facts extracted and reviewed")
    return {"facts": facts, "notes": ext.get("notes"), "review_summary": rev.get("summary"), "loop": loop}


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
