"""Identify what a workbook values (target name, project code name) and its valuation date.

Two steps: SQL over model.db collects candidate cells (named ranges, labelled rows, header text), then
one small model call picks among them and cites the cell. Nothing here is specific to one workbook.
    uv run python bench/identify.py out/<dir>/model.db
"""
import json
import re
import sqlite3
import sys
from collections import Counter

import rodb

DATE_RE = re.compile(r"^(\d{4}-\d{2}-\d{2})")
VAL_DATE_NAME = re.compile(r"val.*date|valuation|as_?at", re.I)
TARGET_NAME = re.compile(r"name|company|target|asset|entity|project|client|borrower|issuer|deal", re.I)
VAL_DATE_LABELS = ("%valuation date%", "%val. date%", "%val date%", "%valuation as at%", "%as at date%")
TARGET_LABELS = ("%model name%", "%project name%", "%company%", "%target%", "%asset name%", "%entity%",
                 "%client%", "%borrower%", "%issuer%", "%deal name%")


def _as_date(v) -> str | None:
    m = DATE_RE.match(str(v)) if v is not None else None
    return m.group(1) if m else None


def _cell(db, sheet: str, addr: str):
    return db.execute("SELECT value FROM cells WHERE sheet=? AND addr=?", (sheet, addr)).fetchone()


def _resolve(db, ref: str):
    """Named-range ref like Meta!$F$34 -> (Sheet!F34, value); None for ranges or broken refs."""
    if "!" not in ref or ":" in ref or "#REF" in ref:
        return None
    sheet, addr = ref.rsplit("!", 1)
    sheet, addr = sheet.strip("=").strip("'"), addr.replace("$", "")
    row = _cell(db, sheet, addr)
    return (f"{sheet}!{addr}", row[0]) if row else None


def candidates(db_path: str) -> dict:
    db = rodb.connect(db_path)
    dates, texts, other_dates = [], [], []
    for name, ref in db.execute("SELECT name, ref FROM names"):
        hit = _resolve(db, ref)
        if not hit or hit[1] in (None, ""):
            continue
        where, v = hit
        if VAL_DATE_NAME.search(name) and _as_date(v):
            dates.append({"value": _as_date(v), "where": where, "why": f"named range {name}"})
        elif "date" in name.lower() and _as_date(v):
            other_dates.append({"value": _as_date(v), "where": where, "why": f"named range {name}"})
        elif TARGET_NAME.search(name) and isinstance(v, str) and len(v) <= 80:
            texts.append({"value": v, "where": where, "why": f"named range {name}"})

    def labelled(patterns, want_date):
        q = " OR ".join("label LIKE ?" for _ in patterns)
        for sheet, row, label in db.execute(f"SELECT sheet, row, label FROM rows WHERE {q} LIMIT 40", patterns):
            for addr, v in db.execute("SELECT addr, value FROM cells WHERE sheet=? AND row=? ORDER BY col",
                                      (sheet, row)):
                ok = _as_date(v) if want_date else (isinstance(v, str) and 1 < len(v) <= 80 and v != label)
                if ok:
                    yield {"value": _as_date(v) if want_date else v, "where": f"{sheet}!{addr}",
                           "why": f"row labelled '{label}'"}
                    break
    dates += labelled(VAL_DATE_LABELS, True)
    texts += labelled(TARGET_LABELS, False)

    # Text repeated in sheet headers (rows 1-4) across sheets, e.g. a project name pulled into every title.
    heads = Counter()
    first_at = {}
    for sheet, addr, v in db.execute("""SELECT sheet, addr, value FROM cells WHERE row <= 4
                                        AND typeof(value)='text' AND length(value) BETWEEN 3 AND 80"""):
        heads[v] += 1
        first_at.setdefault(v, f"{sheet}!{addr}")
    for v, n in heads.most_common(6):
        if n >= 2:
            texts.append({"value": v, "where": first_at[v], "why": f"in the header of {n} sheets"})

    # Short standalone text near the top of each sheet (titles, asset names on a summary sheet).
    seen = {t["value"] for t in texts}
    for sheet, addr, v in db.execute("""
            SELECT c.sheet, c.addr, c.value FROM cells c WHERE c.row <= 20 AND c.formula IS NULL
            AND typeof(c.value)='text' AND length(c.value) BETWEEN 4 AND 60
            AND (SELECT COUNT(*) FROM cells d WHERE d.sheet=c.sheet AND d.row=c.row) = 1
            ORDER BY c.rowid LIMIT 60"""):
        if v not in seen:
            seen.add(v)
            texts.append({"value": v, "where": f"{sheet}!{addr}", "why": "standalone text near top of sheet"})
    return {"valuation_date": dates[:15], "target": texts[:50], "other_dates": other_dates[:15]}


def fallback(c: dict, filename: str) -> dict:
    """Best guess without a model: the most-cited date and the most-repeated header text."""
    votes = Counter(d["value"] for d in c["valuation_date"])
    vd = votes.most_common(1)[0][0] if votes else None
    heads = [t for t in c["target"] if t["why"].startswith("in the header")]
    tgt = (heads or c["target"] or [{"value": filename, "where": None}])[0]
    return {"target_name": tgt["value"], "target_evidence": tgt["where"], "project_name": None,
            "valuation_date": vd,
            "valuation_date_evidence": next((d["where"] for d in c["valuation_date"] if d["value"] == vd), None),
            "confidence": "low", "notes": "Picked by rules, without the model.", "method": "rules"}


PROMPT = """You are checking an Excel valuation model. From the candidate cells below, identify:
- target_name: the company or asset being valued (a real-world name, e.g. an airport or company, not a code name)
- project_name: the deal code name if there is one (e.g. "Project X"), else null
- valuation_date: the date the valuation is as at, YYYY-MM-DD (not the model date, acquisition date or a log entry)
- target_evidence / valuation_date_evidence: the cell address you used ("Sheet!A1")
- confidence: high / medium / low, and notes: one sentence on anything ambiguous (e.g. conflicting dates)
Use only the candidates; use null if none fits.

File name: {filename}
Candidates:
{cands}"""


_S, _N = {"type": "string"}, {"type": ["string", "null"]}
_FIELDS = {"target_name": _N, "target_evidence": _N, "project_name": _N, "valuation_date": _N,
           "valuation_date_evidence": _N, "confidence": {"type": "string", "enum": ["high", "medium", "low"]},
           "notes": _S}
SCHEMA = {"type": "json_schema", "name": "identity", "strict": True,
          "schema": {"type": "object", "properties": _FIELDS, "required": list(_FIELDS),
                     "additionalProperties": False}}


def identify(db_path: str, filename: str, model: str = "gpt-4o", file_id: int | None = None) -> dict:
    c = candidates(db_path)
    try:
        from llm import client, create
        r = create(client(interactive=False), model, input=PROMPT.format(filename=filename, cands=json.dumps(c, indent=1, default=str)),
            text={"format": SCHEMA}, purpose="identify", log={"workbook": file_id})
        if r.usage:
            import usage
            usage.record(model, r.usage, "identify", file_id)
        out = json.loads(r.output_text)
        out["method"] = model
    except Exception as e:  # no sign-in or model error: still return something useful
        out = fallback(c, filename)
        out["notes"] += f" ({type(e).__name__})"
    out["candidates"] = c
    return out


if __name__ == "__main__":
    sys.path.insert(0, "bench")
    res = identify(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else "")
    res.pop("candidates")
    print(json.dumps(res, indent=1))
