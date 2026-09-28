"""Every model call, kept for audit and for improving the prompts later: what was sent, what came back, how long it
took, what it cost, and what it was for (the engagement, the document or workbook, the step, the purpose).

out/calls.db (git-ignored, like everything under out/: it holds report text). Images sent with a call aren't kept,
only a note that one was sent; the sign-in never passes through here. llm.create() writes one row per call,
successful or not. What a call was for comes from tag(), set by the job running it (a context variable, so every
call made inside the job, on any thread started with carry(), is tagged), and from the purpose the caller names.
    uv run python bench/calllog.py            # the last calls
"""
import contextlib
import contextvars
import json
import sqlite3
import threading
import time
from pathlib import Path

DB = Path(__file__).resolve().parent.parent / "out" / "calls.db"
MAX_TEXT = 1_000_000  # characters kept of a request or a response
_CTX: contextvars.ContextVar[dict] = contextvars.ContextVar("calllog_tags", default={})
_LOCK = threading.Lock()
SCHEMA = """CREATE TABLE IF NOT EXISTS calls(id INTEGER PRIMARY KEY, ts REAL, secs REAL, model TEXT, purpose TEXT,
  engagement INT, document INT, workbook INT, step TEXT, session TEXT, input_tokens INT, cached_tokens INT,
  output_tokens INT, cost_usd REAL, ok INT, error TEXT, request TEXT, response TEXT);
CREATE INDEX IF NOT EXISTS ix_calls_eng ON calls(engagement, ts);
CREATE INDEX IF NOT EXISTS ix_calls_wb ON calls(workbook, ts);"""


def _conn() -> sqlite3.Connection:
    DB.parent.mkdir(exist_ok=True)
    db = sqlite3.connect(DB, timeout=30, check_same_thread=False)
    db.executescript(SCHEMA)
    return db


@contextlib.contextmanager
def tag(**tags):
    """Tag every call made inside the block (engagement=, document=, workbook=, step=, session=)."""
    token = _CTX.set({**_CTX.get(), **{k: v for k, v in tags.items() if v is not None}})
    try:
        yield
    finally:
        _CTX.reset(token)


def carry(fn):
    """fn wrapped to run with the tags of the thread that wraps it (for thread pools, which start untagged)."""
    ctx = contextvars.copy_context()
    return lambda *a, **k: ctx.copy().run(fn, *a, **k)


def _plain(x):
    """A request's input as JSON-able data, images left out."""
    if hasattr(x, "model_dump"):
        x = x.model_dump(exclude_none=True)
    if isinstance(x, dict):
        if x.get("type") == "input_image":
            return {"type": "input_image", "note": "an image was sent (not kept in the log)"}
        return {k: _plain(v) for k, v in x.items() if k != "encrypted_content"}
    if isinstance(x, (list, tuple)):
        return [_plain(v) for v in x]
    return x


def _response_text(r) -> str:
    if r is None:
        return ""
    parts = []
    try:
        if r.output_text:
            parts.append(r.output_text)
    except Exception:
        pass
    for o in getattr(r, "output", None) or []:
        if getattr(o, "type", "") == "function_call":
            parts.append(f"[tool call] {o.name}({o.arguments})")
    return "\n".join(parts)


def record(model: str, kwargs: dict, response, secs: float, error: Exception | None = None, purpose: str | None = None,
           extra: dict | None = None) -> None:
    """One row per call. Never raises: logging must not break the call it logs."""
    try:
        from pricing import cost
        tags = {**_CTX.get(), **(extra or {})}
        u = getattr(response, "usage", None)
        i = getattr(u, "input_tokens", 0) or 0
        o = getattr(u, "output_tokens", 0) or 0
        c = getattr(getattr(u, "input_tokens_details", None), "cached_tokens", 0) or 0
        fmt = ((kwargs.get("text") or {}).get("format") or {}).get("name")
        purpose = purpose or tags.get("purpose") or fmt or ("chat" if kwargs.get("tools") else "text")
        req = {"instructions": kwargs.get("instructions"), "input": _plain(kwargs.get("input")),
               "tools": [t.get("name") for t in kwargs.get("tools") or [] if isinstance(t, dict)] or None,
               "format": fmt, "max_output_tokens": kwargs.get("max_output_tokens")}
        row = (time.time(), round(secs, 3), model, purpose, tags.get("engagement"), tags.get("document"),
               tags.get("workbook"), tags.get("step"), tags.get("session"), i, c, o, cost(model, i, c, o),
               int(error is None), f"{type(error).__name__}: {error}" if error else None,
               json.dumps(req, default=str, ensure_ascii=False)[:MAX_TEXT], _response_text(response)[:MAX_TEXT])
        with _LOCK, _conn() as db:
            db.execute("INSERT INTO calls(ts, secs, model, purpose, engagement, document, workbook, step, session, "
                       "input_tokens, cached_tokens, output_tokens, cost_usd, ok, error, request, response) "
                       "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", row)
    except Exception:
        pass


# ---- reading it back ----------------------------------------------------------------------------------------------

def _q(sql: str, *args) -> list[dict]:
    if not DB.exists():
        return []
    with _conn() as db:
        cur = db.execute(sql, args)
        cols = [c[0] for c in cur.description]
        return [dict(zip(cols, r)) for r in cur]


_SUMS = ("COUNT(*) AS calls, SUM(input_tokens) AS input_tokens, SUM(cached_tokens) AS cached_tokens, "
         "SUM(output_tokens) AS output_tokens, SUM(cost_usd) AS cost_usd, SUM(secs) AS model_secs, "
         "SUM(1 - ok) AS failed")


def breakdown(engagement: int, workbooks: list[int]) -> dict:
    """An engagement's calls (its own, and those made processing its workbooks): totals, by file, by step."""
    wb = ",".join(str(int(w)) for w in workbooks) or "NULL"
    where = f"(engagement = ? OR (engagement IS NULL AND workbook IN ({wb})))"
    return {
        "total": (_q(f"SELECT {_SUMS}, MIN(ts) AS first, MAX(ts) AS last FROM calls WHERE {where}", engagement) or [{}])[0],
        "by_file": _q(f"""SELECT document, workbook, {_SUMS} FROM calls WHERE {where}
                          GROUP BY document, workbook ORDER BY cost_usd DESC""", engagement),
        "by_step": _q(f"""SELECT document, workbook, step, purpose, model, {_SUMS} FROM calls WHERE {where}
                          GROUP BY document, workbook, step, purpose, model ORDER BY MIN(ts)""", engagement),
        "recent": _q(f"""SELECT id, ts, secs, model, purpose, document, workbook, step, input_tokens, cached_tokens,
                         output_tokens, cost_usd, ok, error FROM calls WHERE {where} ORDER BY id DESC LIMIT 300""", engagement),
    }


def call(cid: int, engagement: int | None = None, workbooks: list[int] | None = None) -> dict | None:
    """One call in full; only if it belongs to the engagement asked about (when given)."""
    rows = _q("SELECT * FROM calls WHERE id=?", cid)
    if not rows:
        return None
    c = rows[0]
    if engagement is not None and c["engagement"] != engagement and c["workbook"] not in (workbooks or []):
        return None
    c["request"] = json.loads(c["request"] or "null")
    return c


if __name__ == "__main__":
    for c in _q("SELECT id, ts, secs, model, purpose, engagement, document, workbook, step, input_tokens, output_tokens, "
                "cost_usd, ok FROM calls ORDER BY id DESC LIMIT 20"):
        print(c)
