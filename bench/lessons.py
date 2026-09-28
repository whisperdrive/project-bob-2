"""What the report agents learn from their review loops: method and rules, never names or figures.

Two tiers, both read fresh into the prompts on every run (extraction, review, the fixes and the table reads):
  docs/report_rules.md   curated rules (R1, R2, ...), in the repo; a person promotes a lesson into it
  out/lessons.json       lessons the agents write after each review loop (L1, L2, ...), git-ignored, shared by
                         every engagement on this machine
After a loop, the reviewer model turns what went wrong and how it was fixed into general rules about method.
Code enforces the anonymity rather than trusting the instruction: a lesson that names anything from the
engagement (target, project, client, file names, the report's proper nouns) or carries a figure is sent back once
to be rewritten without it, then dropped. A lesson the loop confirms again is reinforced rather than repeated;
the agents cite the rule IDs they apply, so each lesson shows how often it was used. A person can retire any.
    uv run python bench/lessons.py            # list them
"""
import json
import os
import re
import threading
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
RULES_FILE = ROOT / "docs" / "report_rules.md"
STORE = ROOT / "out" / "lessons.json"
SCOPES = ("facts", "tables")
MAX_IN_PROMPT = 25
_LOCK = threading.Lock()

# Words that are capitalised in reports but are method, not identity (kept out of the proper-noun check).
GENERIC = set("""
wacc capm dcf ebitda ebit ebitdar rab cpi tgr npv irr ev gst fy pdf pptx aud usd nzd gbp eur ifrs aasb gaap nta nav
ltm ntm cagr mrp erp tv fcf fcff fcfe opex capex npat eps tsr cfads dscr llcr rba asx gdp ppi kpi yearfrac xnpv
gordon excel markdown python table tables page pages figure appendix section executive summary valuation report
company company's group trust fund business asset target vendor purchaser acquirer board management directors
january february march april may june july august september october november december
monday tuesday wednesday thursday friday saturday sunday
""".split())
_WORD = re.compile(r"[A-Za-z][A-Za-z'&.-]*[A-Za-z]")
_FIGURE = re.compile(r"\d[\d,]*\.\d|\d{2,}|\d\s?%|[$€£]\s?\d")  # a single digit on its own is allowed ("1 decimal")

DISTIL_PROMPT = """You keep the rulebook for two agents that read valuation reports: a table reader that transcribes
tables from page images, and an extractor that pulls the key facts (target, valuation date, conclusions and
ranges, assumptions, approach, sensitivities) with the page and a verbatim quote. Below is what went wrong in
a review loop just now ({scope}) and how it was resolved. Write down what the agents should do differently next
time, as rules about METHOD: where to look, how to read, what to check, how to quote, pitfalls to avoid.

What you write must be anonymous and general:
- it must apply to any report, company, asset or year
- no names of companies, people, projects, assets, places, advisers or documents; no figures, percentages,
  dates or amounts; no quotes from the report. Describe things generically ("the preferred value within a
  range", "a currency prefix in a column heading", "a figure repeated in the summary and the body").
- one rule per lesson, imperative, at most 40 words, with a short "why" (at most 25 words, also generic)
- only lessons the episodes support. Don't restate an existing rule: list its ID in reinforce instead, and
  refine an existing rule's wording only if the episodes show it's incomplete. Nothing learned is a fine answer.

Existing rules:
{existing}

Episodes:
{episodes}"""

REWRITE_PROMPT = """These rules for report-reading agents mention things specific to one engagement (listed after
each). Rewrite each one so it keeps the method but names no company, person, project, asset, place or document
and carries no figure, percentage, date or amount. Keep the same IDs.

{items}"""

_S = {"type": "string"}
_NEW = {"scope": {"type": "string", "enum": list(SCOPES)}, "rule": _S, "why": _S}
_REF = {"id": _S, "rule": _S, "why": _S}
DISTIL_SCHEMA = {"type": "json_schema", "name": "lessons", "strict": True, "schema": {
    "type": "object", "additionalProperties": False, "required": ["new", "reinforce", "refine"],
    "properties": {
        "new": {"type": "array", "items": {"type": "object", "additionalProperties": False, "required": list(_NEW),
                                           "properties": _NEW}},
        "reinforce": {"type": "array", "items": _S},
        "refine": {"type": "array", "items": {"type": "object", "additionalProperties": False, "required": list(_REF),
                                              "properties": _REF}}}}}
REWRITE_SCHEMA = {"type": "json_schema", "name": "rewrites", "strict": True, "schema": {
    "type": "object", "additionalProperties": False, "required": ["items"],
    "properties": {"items": {"type": "array", "items": {"type": "object", "additionalProperties": False,
                                                        "required": list(_REF), "properties": _REF}}}}}


# ---- the store ----------------------------------------------------------------------------------------------

def load() -> dict:
    try:
        return json.loads(STORE.read_text(encoding="utf-8"))
    except (FileNotFoundError, ValueError):
        return {"next": 1, "lessons": []}


def _save(d: dict) -> None:
    STORE.parent.mkdir(exist_ok=True)
    tmp = STORE.with_suffix(".tmp")
    tmp.write_text(json.dumps(d, indent=1, ensure_ascii=False), encoding="utf-8")
    os.replace(tmp, STORE)


def curated(scope: str | None = None) -> list[dict]:
    """Rules in docs/report_rules.md: lines like "- **R3** [facts] Rule text. _Why: ..._"."""
    try:
        text = RULES_FILE.read_text(encoding="utf-8")
    except FileNotFoundError:
        return []
    out = []
    for m in re.finditer(r"^- \*\*(R\d+)\*\* \[(facts|tables)\] (.+)$", text, re.M):
        if scope is None or m[2] == scope:
            out.append({"id": m[1], "scope": m[2], "rule": m[3].strip()})
    return out


def active(scope: str) -> list[dict]:
    ls = [x for x in load()["lessons"] if x["scope"] == scope and x["status"] == "active"]
    return sorted(ls, key=lambda x: (-x["seen"], -x["last"]))


def rules_text(scope: str) -> str:
    """The rules for a prompt: curated first, then the most-confirmed learned lessons."""
    lines = [f"- {r['id']}: {r['rule']}" for r in curated(scope)]
    lines += [f"- {x['id']}: {x['rule']} (why: {x['why']})" for x in active(scope)[:MAX_IN_PROMPT]]
    return "\n".join(lines) or "(none yet)"


def applied(ids) -> None:
    """Count the lessons the agents say they applied."""
    ids = {str(i).strip() for i in ids or [] if str(i).strip().startswith("L")}
    if not ids:
        return
    with _LOCK:
        d = load()
        for x in d["lessons"]:
            if x["id"] in ids:
                x["applied"] = x.get("applied", 0) + 1
        _save(d)


def set_status(lid: str, status: str) -> dict:
    if status not in ("active", "retired"):
        raise ValueError("status must be active or retired")
    with _LOCK:
        d = load()
        x = next((x for x in d["lessons"] if x["id"] == lid), None)
        if not x:
            raise ValueError(f"no lesson {lid}")
        x["status"] = status
        _save(d)
    return x


def promote(lid: str) -> dict:
    """Copy a learned lesson into docs/report_rules.md as the next R rule, and retire the lesson (the rule
    replaces it in the prompts). The file is in the repo, so a person reviews it before committing."""
    with _LOCK:
        d = load()
        x = next((x for x in d["lessons"] if x["id"] == lid), None)
        if not x:
            raise ValueError(f"no lesson {lid}")
        bad = leaks(f"{x['rule']} {x['why']}", set())
        if bad:
            raise ValueError(f"{lid} carries figures ({', '.join(bad)}); edit it before promoting")
        text = RULES_FILE.read_text(encoding="utf-8") if RULES_FILE.exists() else _RULES_HEAD
        n = max([int(m) for m in re.findall(r"^- \*\*R(\d+)\*\*", text, re.M)] + [0]) + 1
        text = text.rstrip("\n") + f"\n- **R{n}** [{x['scope']}] {x['rule']} _Why: {x['why']}_ (from {lid})\n"
        RULES_FILE.parent.mkdir(exist_ok=True)
        RULES_FILE.write_text(text, encoding="utf-8")
        x.update(status="retired", promoted_to=f"R{n}")
        _save(d)
    return x


_RULES_HEAD = """# Report reading rules

Read by the report agents on every run (bench/lessons.py): the table reader, the fact extractor, its
reviewer, and both sides of the review loop. Each rule has an ID the agents cite. Rules are about method only:
no names, figures or quotes from any engagement. Learned lessons (L1, L2, ...) live in out/lessons.json on each
machine; promote one here from the Report reference step when it has proved itself.

"""


# ---- anonymity ----------------------------------------------------------------------------------------------

def proper_nouns(markdown: str) -> set[str]:
    """Names in the document: words capitalised in the middle of a sentence (after a lower-case word, or after
    another such name: "prepared for Acme Holdings"), never written in lower case, and not generic terms.
    Headings, table cells and sentence starts are capitalised anyway, so they don't count."""
    text = re.sub(r"<!--.*?-->", " ", markdown or "")
    lower = {w.lower() for w in _WORD.findall(text) if w.islower()}
    out = set()
    for cell in re.split(r"[|\n]", text):
        prev, prev_name = None, False
        for tk in re.findall(r"[A-Za-z][A-Za-z'&-]*|[.!?:;()]", cell):
            if not tk[0].isalpha():
                prev, prev_name = None, False
                continue
            name = False
            if tk[0].isupper() and prev is not None and (prev.islower() or prev_name):
                w = tk.lower().strip("'-")
                name = len(w) >= 3 and w not in lower and w not in GENERIC
                if name:
                    out.add(w)
            prev, prev_name = tk, name
    return out


def deny_terms(names: list[str], markdown: str = "") -> set[str]:
    """Everything a lesson mustn't mention: the document's proper nouns, and the words of the given names (target,
    project, client, engagement, file names) except ordinary words (ones the document also writes in lower case,
    or that the lesson prompts use), so "toll road" or "airport" can describe a kind of asset but a name can't."""
    text = re.sub(r"<!--.*?-->", " ", markdown or "")
    common = GENERIC | {w.lower() for w in _WORD.findall(text) if w.islower()} | \
        {w.lower() for w in _WORD.findall(DISTIL_PROMPT + REWRITE_PROMPT)}
    out = set(proper_nouns(markdown))
    for n in names:
        for w in _WORD.findall(re.sub(r"[_]+|\.(pdf|pptx|xlsx|xlsm)$", " ", str(n or ""), flags=re.I)):
            if len(w) >= 3 and w.lower() not in common:
                out.add(w.lower())
    return out


def leaks(text: str, deny: set[str]) -> list[str]:
    """What in a lesson would identify an engagement: denied words and figures."""
    found = [w for w in dict.fromkeys(_WORD.findall(text)) if w.lower() in deny]
    found += [m.group(0) for m in _FIGURE.finditer(text)]
    return found


# ---- learning -----------------------------------------------------------------------------------------------

def _norm(s: str) -> set[str]:
    return {w.lower() for w in _WORD.findall(s) if len(w) > 2}


def _similar(a: str, b: str) -> bool:
    x, y = _norm(a), _norm(b)
    return bool(x and y) and len(x & y) / len(x | y) >= 0.75


def distil(scope: str, episodes: list[dict], names: list[str], markdown: str, model: str, on_usage=None,
           engagement: int | None = None) -> dict:
    """Turn a loop's episodes into lessons. Returns {"added", "reinforced", "refined", "dropped"}."""
    out = {"added": [], "reinforced": [], "refined": [], "dropped": []}
    if not episodes:
        return out
    from llm import client, create
    llm = client(interactive=False)

    def call(prompt, schema, purpose):
        r = create(llm, model, input=prompt, text={"format": schema}, max_output_tokens=4000)
        if r.usage and on_usage:
            on_usage(model, r.usage, purpose)
        return json.loads(r.output_text)

    existing = [f"- {r['id']} [{r['scope']}] {r['rule']}" for r in curated()] + \
               [f"- {x['id']} [{x['scope']}] {x['rule']}" for x in load()["lessons"] if x["status"] == "active"]
    res = call(DISTIL_PROMPT.format(scope=scope, existing="\n".join(existing) or "(none yet)",
                                    episodes=json.dumps(episodes, indent=1, ensure_ascii=False)[:60000]),
               DISTIL_SCHEMA, "lessons")
    deny = deny_terms(names, markdown)
    proposed = [{"id": f"new{i}", "scope": n["scope"], "rule": n["rule"].strip(), "why": n["why"].strip()}
                for i, n in enumerate(res["new"])] + \
               [{"id": r["id"], "rule": r["rule"].strip(), "why": r["why"].strip(), "refine": True} for r in res["refine"]]
    bad = {p["id"]: leaks(f"{p['rule']} {p['why']}", deny) for p in proposed}
    bad = {k: v for k, v in bad.items() if v}
    if bad:  # one chance to rewrite without the specifics
        items = "\n".join(f"{p['id']}: {p['rule']} Why: {p['why']}  [mentions: {', '.join(bad[p['id']])}]"
                          for p in proposed if p["id"] in bad)
        fixed = {x["id"]: x for x in call(REWRITE_PROMPT.format(items=items), REWRITE_SCHEMA, "lessons")["items"]}
        for p in proposed:
            if p["id"] in bad and p["id"] in fixed:
                p["rule"], p["why"] = fixed[p["id"]]["rule"].strip(), fixed[p["id"]]["why"].strip()
    now = time.time()
    with _LOCK:
        d = load()
        by_id = {x["id"]: x for x in d["lessons"]}

        def seen(x):
            x["seen"] += 1
            x["last"] = now
            if engagement is not None and engagement not in x["engagements"]:
                x["engagements"].append(engagement)

        for lid in res["reinforce"]:
            if lid in by_id:
                seen(by_id[lid])
                out["reinforced"].append(lid)
        for p in proposed:
            still = leaks(f"{p['rule']} {p['why']}", deny)
            if still:
                out["dropped"].append({"rule": p["rule"], "why": "named or quoted the engagement: " + ", ".join(still)})
                continue
            if p.get("refine"):
                x = by_id.get(p["id"])
                if x and x["status"] == "active":
                    x.update(rule=p["rule"], why=p["why"])
                    seen(x)
                    out["refined"].append(x["id"])
                continue
            twin = next((x for x in d["lessons"] if x["scope"] == p["scope"] and _similar(x["rule"], p["rule"])), None)
            if twin:
                seen(twin)
                out["reinforced"].append(twin["id"])
                continue
            x = {"id": f"L{d['next']}", "scope": p["scope"], "rule": p["rule"], "why": p["why"], "seen": 1,
                 "applied": 0, "first": now, "last": now, "status": "active",
                 "engagements": [engagement] if engagement is not None else []}
            d["next"] += 1
            d["lessons"].append(x)
            by_id[x["id"]] = x
            out["added"].append(x["id"])
        _save(d)
    return out


if __name__ == "__main__":
    for r in curated():
        print(f"{r['id']:>4} [{r['scope']}] {r['rule']}")
    for x in load()["lessons"]:
        print(f"{x['id']:>4} [{x['scope']}] {x['status']:7} seen {x['seen']} applied {x.get('applied', 0)}: {x['rule']}")
