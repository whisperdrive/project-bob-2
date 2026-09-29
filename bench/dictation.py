"""Dictation in Ask: Azure Speech real-time speech to text, streamed from the browser.

The page records the microphone with Microsoft's Speech SDK and streams it straight to the Azure Speech resource;
the server only hands it a short-lived token (the key stays in .env) and the engagement's own words as a phrase
list, so names like the sheets, the report's labels and WACC come out spelled right. The Free (F0) tier gives 5
audio hours a month, one stream at a time; the page stops after a silence to save them, and reports the seconds
it streamed so the desk can show what's been used this month.

Settings in .env (git-ignored): SPEECH_KEY, SPEECH_REGION (e.g. australiaeast), optionally SPEECH_LANGUAGE (en-AU).
"""
import json
import os
import re
import sqlite3
import threading
import time
import urllib.error
import urllib.request
from pathlib import Path

from llm import _load_env

DB = Path(__file__).resolve().parent.parent / "out" / "registry.db"
FREE_SECONDS = 5 * 3600     # the F0 tier's monthly allowance
MAX_PHRASES = 500           # the service takes up to 2,000; longer lists cost accuracy and latency
TOKEN_SECONDS = 240         # STS tokens live 10 minutes; hand out none with less than 6 left
_lock = threading.Lock()
_token: dict = {}

# Words a valuation conversation uses that a general model mishears. Engagement words come from its data at runtime.
DOMAIN = ["WACC", "EBITDA", "EBIT", "NPAT", "capex", "opex", "NPV", "XNPV", "IRR", "DCF", "CPI", "FCFF", "FCFE",
          "enterprise value", "equity value", "net debt", "terminal value", "terminal growth", "Gordon growth",
          "discount rate", "discount factor", "cost of equity", "cost of debt", "gearing", "beta", "mid-period",
          "roll-forward", "roll forward", "overlay", "what-if", "sensitivities", "valuation date", "cash flows",
          "free cash flow", "working capital", "depreciation", "amortisation", "tax losses", "franking credits"]


def settings() -> dict:
    """The Speech settings, read from .env each time so a key added after startup works without a restart."""
    _load_env()
    return {"key": os.environ.get("SPEECH_KEY", "").strip(), "region": os.environ.get("SPEECH_REGION", "").strip(),
            "language": os.environ.get("SPEECH_LANGUAGE", "").strip() or "en-AU"}


def status() -> dict:
    s = settings()
    return {"configured": bool(s["key"] and s["region"]), "engine": "Azure Speech", "language": s["language"],
            "region": s["region"] or None}


def _why(code: int, body: str) -> str:
    if code == 401:
        return "Azure Speech refused the key: check SPEECH_KEY and SPEECH_REGION in .env (the key's resource must be in that region)."
    if code == 403 and "quota" in body.lower():
        return "The free tier's 5 hours of speech for this month are used up; dictation works again on the 1st."
    if code == 429:
        return "Azure Speech is busy with another stream (the free tier takes one at a time); try again in a moment."
    return f"Azure Speech didn't issue a token ({code}): {body[:200]}"


def token() -> dict:
    """A token for the page's SDK: {token, region}. Minted with the key and reused for a few minutes."""
    s = settings()
    if not (s["key"] and s["region"]):
        raise ValueError("Dictation needs an Azure Speech resource: add SPEECH_KEY and SPEECH_REGION to .env "
                         "(the Free F0 tier gives 5 hours a month).")
    with _lock:
        if _token.get("for") == (s["key"], s["region"]) and time.time() - _token["at"] < TOKEN_SECONDS:
            return {"token": _token["token"], "region": s["region"]}
        req = urllib.request.Request(f"https://{s['region']}.api.cognitive.microsoft.com/sts/v1.0/issueToken",
                                     data=b"", method="POST", headers={"Ocp-Apim-Subscription-Key": s["key"]})
        try:
            with urllib.request.urlopen(req, timeout=15) as r:
                tok = r.read().decode()
        except urllib.error.HTTPError as e:
            raise ValueError(_why(e.code, e.read().decode(errors="replace"))) from None
        except urllib.error.URLError as e:
            raise ValueError(f"Couldn't reach Azure Speech in {s['region']}: {e.reason}") from None
        _token.update({"for": (s["key"], s["region"]), "at": time.time(), "token": tok})
        return {"token": tok, "region": s["region"]}


def _phrase(text) -> str | None:
    """A label as someone would say it: no bracketed notes, underscores or bare numbers; None if nothing's left."""
    t = re.sub(r"\[[^\]]*\]|\([^)]*\)", " ", str(text or ""))           # [units], (notes)
    t = re.sub(r"[_|]+", " ", t)
    t = re.sub(r"\s+", " ", t).strip(" -:.,;")
    if len(t) < 3 or len(t) > 60 or not re.search(r"[A-Za-z]{2}", t) or re.fullmatch(r"[\d\s.,%$-]+", t):
        return None
    return t


def phrases(eid: int) -> list[str]:
    """The engagement's own words, most useful first: the report's key facts (the names in them, then their labels), the
    overlay's outputs and levers, the workbooks' sheet names, the engagement's name; then the valuation words.
    Deduplicated, at most MAX_PHRASES."""
    import engagement
    out, seen = [], set()

    def add(t):
        p = _phrase(t)
        if p and p.lower() not in seen:
            seen.add(p.lower())
            out.append(p)

    e = engagement._q("SELECT name, overlay_json FROM engagements WHERE id=?", eid)
    if not e:
        raise ValueError("no such engagement")
    ref = engagement.reference(eid)
    for f in ref:                       # the names first: the target, the project, the client, the method
        add(f.get("value_text"))
    for f in ref:
        add(f.get("label") or f.get("key"))
    summary = json.loads(e[0]["overlay_json"] or "null") or {}
    for o in summary.get("outputs") or []:
        add(o.get("label"))
    for lv in summary.get("levers") or []:
        add(lv.get("label"))
    for key in ("prior_overlay", "prior_model", "current_model"):
        w = engagement._role_wb(eid, key)
        for s in sorted((w or {}).get("sheets") or []):
            add(s)
    add(e[0]["name"])
    for t in DOMAIN:
        add(t)
    return out[:MAX_PHRASES]


def _conn() -> sqlite3.Connection:
    DB.parent.mkdir(exist_ok=True)
    db = sqlite3.connect(DB, check_same_thread=False)
    db.execute("CREATE TABLE IF NOT EXISTS dictation(ts REAL, seconds REAL)")
    return db


def used() -> float:
    """Seconds streamed from this desk since the start of the month (the free hours reset on the 1st)."""
    start = time.mktime(time.strptime(time.strftime("%Y-%m-01"), "%Y-%m-%d"))
    with _lock, _conn() as db:
        return db.execute("SELECT COALESCE(SUM(seconds), 0) FROM dictation WHERE ts >= ?", (start,)).fetchone()[0]


def record(seconds: float) -> dict:
    seconds = max(0.0, min(float(seconds), 3600.0))
    with _lock, _conn() as db:
        db.execute("INSERT INTO dictation VALUES (?, ?)", (time.time(), seconds))
    return {"used_seconds": used(), "free_seconds": FREE_SECONDS}


def start(eid: int) -> dict:
    """What the page needs to start dictating into an engagement's Ask box."""
    s = settings()
    return {**token(), "language": s["language"], "phrases": phrases(eid), "used_seconds": used(),
            "free_seconds": FREE_SECONDS}
