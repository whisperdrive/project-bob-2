"""Engagements: the four files of a recurring valuation, from upload to an approved reference, roles and a map.

  1. upload     workbooks go through the Model Desk pipeline (library.py: fingerprint, build, identify) and are
                shared with it; reports (PDF / PPTX) are read here by docingest.py, tables checked
  2. reference  reportfacts.py extracts the report's key facts, checks them in code and has a reviewer model
                review them; flagged tables and open facts go through a review and remediation loop between the
                two models (docingest.resolve_tables, reportfacts.resolve), whose lessons are kept anonymised
                for next time (lessons.py); a person approves, edits or rejects each fact
  3. roles      roles.py suggests prior report / prior client model / prior overlay / current client model
                (the overlay as a sheet list, since it can sit inside the client model); a person confirms
  4. compare    diff.py between the prior and current client models (overlay sheets left out)
  5. map        linkmap.py: report -> overlay cells, overlay -> prior client rows, prior -> current rows, and
                the overlay's DCFs recomputed in Python (valuation.py)
  6. overlay    overlay.py: the overlay sheets compiled to a Python module, checked cell by cell against Excel and
                against the report (conclusions, sensitivities), then run live on the prior or current client model

State is in out/engage.db; reports are copied to uploads/<sha12>/ and read into out/docs/<stem>__<sha8>/
(all git-ignored). Model calls use the engagement's chosen models and go to the usage log with session
"engagement-<id>". Long steps run on one background worker thread, like library.py.
"""
import json
import queue
import re
import shutil
import sqlite3
import threading
import time
import traceback
from collections import Counter
from contextlib import closing
from pathlib import Path

import diff as diffmod
import docingest
import extlinks
import library
import calllog
import lessons
import linkmap
import reportfacts
import rodb
import roles as rolesmod
import usage

ROOT = Path(__file__).resolve().parent.parent
OUT, UPLOADS = ROOT / "out", ROOT / "uploads"
DB, DOCS = OUT / "engage.db", OUT / "docs"
DEFAULT_MODEL = "gpt-6-luna"      # extraction and table reads
DEFAULT_REVIEWER = "gpt-6-sol"    # second reads and fact review: a different, stronger model than the first read
DEFAULT_ARBITER = "gpt-6-sol"     # settles what the review loop can't (the user's choice, 2026-09-28): a fresh look
                                  # with its own brief, even where it is the reviewer's model
REPORT_TYPES = (".pdf", ".pptx")

_lock = threading.Lock()
REPORT_JOBS = ("doc", "resolve_tables", "facts", "resolve_facts")


class _Lanes:
    """Two background lanes, each running its jobs one at a time: the report's (reading it, the key facts, both
    review loops) and the models' (roles, compare, map, Python overlay, charts, doctor, row agents). A review loop
    never holds up the map; the models' jobs stay one at a time because they share model.db and the Python
    session. Model calls from both share one rate limit (ratelimit.py)."""

    def __init__(self):
        self.queues: dict[str, queue.Queue] = {"report": queue.Queue(), "models": queue.Queue()}
        self.running: dict[str, tuple[str, int, float]] = {}  # lane -> (job, id, started)

    def put(self, job: tuple[str, int]) -> None:
        self.queues["report" if job[0] in REPORT_JOBS else "models"].put(job)


_jobs = _Lanes()

SCHEMA = """
CREATE TABLE IF NOT EXISTS engagements(id INTEGER PRIMARY KEY, name TEXT, created_at REAL, updated_at REAL,
  model TEXT, reviewer_model TEXT, roles_suggested TEXT,
  compare_status TEXT, compare_step TEXT, compare_error TEXT, compare_json TEXT, compare_summary TEXT,
  map_status TEXT, map_step TEXT, map_error TEXT, map_json TEXT,
  overlay_status TEXT, overlay_step TEXT, overlay_error TEXT, overlay_json TEXT);
CREATE TABLE IF NOT EXISTS eng_files(engagement_id INT, file_id INT, added_at REAL, PRIMARY KEY(engagement_id, file_id));
CREATE TABLE IF NOT EXISTS documents(id INTEGER PRIMARY KEY, engagement_id INT, sha256 TEXT, filename TEXT, kind TEXT,
  size INT, uploaded_at REAL, source_path TEXT, out_dir TEXT, status TEXT, step TEXT, pct REAL, error TEXT,
  pages INT, n_tables INT, n_flagged INT, doc_json TEXT, processed_at REAL,
  facts_status TEXT, facts_step TEXT, facts_error TEXT, facts_notes TEXT, review_summary TEXT, loop_json TEXT,
  UNIQUE(engagement_id, sha256));
CREATE TABLE IF NOT EXISTS facts(id INTEGER PRIMARY KEY, engagement_id INT, document_id INT, n INT, category TEXT,
  key TEXT, label TEXT, value_text TEXT, low_text TEXT, high_text TEXT, value REAL, unit TEXT, basis TEXT, page INT,
  quote TEXT, origin TEXT, check_json TEXT, review_json TEXT, status TEXT DEFAULT 'pending', final_json TEXT,
  updated_at REAL, agent_json TEXT, decided_by TEXT);
CREATE TABLE IF NOT EXISTS roles(engagement_id INT, role TEXT, kind TEXT, ref_id INT, sheets_json TEXT, why_json TEXT,
  confirmed INT DEFAULT 0, PRIMARY KEY(engagement_id, role));
"""
FACT_FIELDS = ("category", "key", "label", "value_text", "low_text", "high_text", "value", "unit", "basis", "page", "quote")


def _conn() -> sqlite3.Connection:
    OUT.mkdir(exist_ok=True)
    db = sqlite3.connect(DB, check_same_thread=False, timeout=30)
    db.row_factory = sqlite3.Row
    db.executescript(SCHEMA)
    have = {r[1] for r in db.execute("PRAGMA table_info(engagements)")}
    for col in ("overlay_status", "overlay_step", "overlay_error", "overlay_json"):  # databases made before step 6
        if col not in have:
            db.execute(f"ALTER TABLE engagements ADD COLUMN {col} TEXT")
    for col, kind in (("charts_status", "TEXT"), ("charts_step", "TEXT"), ("charts_error", "TEXT"),
                      ("charts_started_at", "REAL"), ("charts_secs", "REAL"),  # ... and before the Summary page
                      ("doctor_status", "TEXT"), ("doctor_step", "TEXT"), ("doctor_error", "TEXT"),
                      ("doctor_started_at", "REAL"), ("doctor_secs", "REAL"),  # ... and before the doctor
                      ("rows_status", "TEXT"), ("rows_step", "TEXT"), ("rows_error", "TEXT"),
                      ("rows_started_at", "REAL"), ("rows_secs", "REAL")):  # ... and before the row agents
        if col not in have:
            db.execute(f"ALTER TABLE engagements ADD COLUMN {col} {kind}")
    for table, col in (("facts", "agent_json"), ("documents", "loop_json"), ("facts", "decided_by")):  # ... and before the loop
        if col not in {r[1] for r in db.execute(f"PRAGMA table_info({table})")}:
            db.execute(f"ALTER TABLE {table} ADD COLUMN {col} TEXT")
    have = {r[1] for r in db.execute("PRAGMA table_info(engagements)")}
    for col in ("arbiter_model", "profile_json", "schedule_json"):  # ... and before the profile and the schedule
        if col not in have:
            db.execute(f"ALTER TABLE engagements ADD COLUMN {col} TEXT")
    db.execute("CREATE TABLE IF NOT EXISTS migrations(name TEXT PRIMARY KEY, at REAL)")
    if not db.execute("SELECT 1 FROM migrations WHERE name='arbiter-sol'").fetchone():
        # once: the arbiter's default moved from gpt-4o to gpt-6-sol, and engagements showing the old default follow
        # it (a later choice of gpt-4o in the header sticks)
        db.execute("UPDATE engagements SET arbiter_model=NULL WHERE arbiter_model='gpt-4o'")
        db.execute("INSERT INTO migrations VALUES ('arbiter-sol', ?)", (time.time(),))
    for col, kind in (("tables_status", "TEXT"), ("tables_left", "INT"), ("tables_step", "TEXT"),
                      ("tables_error", "TEXT"), ("tables_started_at", "REAL"), ("tables_secs", "REAL")):  # ... and before background reading
        if col not in {r[1] for r in db.execute("PRAGMA table_info(documents)")}:
            db.execute(f"ALTER TABLE documents ADD COLUMN {col} {kind}")
    timing = [("documents", c) for c in ("started_at", "doc_secs", "facts_started_at", "facts_secs")] + \
        [("engagements", f"{k}_{c}") for k in ("compare", "map", "overlay") for c in ("started_at", "secs")] + \
        [("engagements", "overlay_pct")]
    for table, col in timing:  # ... and before job timings
        if col not in {r[1] for r in db.execute(f"PRAGMA table_info({table})")}:
            db.execute(f"ALTER TABLE {table} ADD COLUMN {col} REAL")
    return db


def _q(sql: str, *args) -> list[dict]:
    with _conn() as db:
        return [dict(r) for r in db.execute(sql, args)]


def _exec(sql: str, *args) -> int:
    with _lock, _conn() as db:
        return db.execute(sql, args).lastrowid


def _set(table: str, rid: int, **fields) -> None:
    with _lock, _conn() as db:
        db.execute(f"UPDATE {table} SET {', '.join(f'{k}=?' for k in fields)} WHERE id=?", (*fields.values(), rid))


def _session(eid: int) -> str:
    return f"engagement-{eid}"


def _logger(eid: int):
    return lambda model, u, purpose: usage.record(model, u, purpose, None, _session(eid))


def friendly(e: Exception) -> str:
    name = type(e).__name__
    if name in ("AuthenticationRequiredError", "ClientAuthenticationError"):
        return ("Azure sign-in needed: run `uv run python bench/llm.py` in a terminal, sign in with the device code, "
                "then retry.")
    if name == "APITimeoutError":
        import llm
        return (f"the model didn't answer within {llm.READ_TIMEOUT:.0f} s, {llm.MAX_RETRIES + 1} tries "
                "(LLM_READ_TIMEOUT in .env): retry")
    return f"{name}: {e}"


# ---- engagements --------------------------------------------------------------------------------------------

def create(name: str) -> dict:
    now = time.time()
    eid = _exec("INSERT INTO engagements(name, created_at, updated_at, model, reviewer_model) VALUES (?,?,?,?,?)",
                name.strip() or "Untitled engagement", now, now, DEFAULT_MODEL, DEFAULT_REVIEWER)
    return get(eid)


def all_engagements() -> list[dict]:
    rows = _q("""SELECT e.id, e.name, e.created_at, e.updated_at,
                 (SELECT COUNT(*) FROM documents d WHERE d.engagement_id = e.id) AS n_docs,
                 (SELECT COUNT(*) FROM eng_files f WHERE f.engagement_id = e.id) AS n_workbooks,
                 (SELECT COUNT(*) FROM roles r WHERE r.engagement_id = e.id AND r.confirmed) AS n_roles
                 FROM engagements e ORDER BY e.updated_at DESC""")
    return rows


def update(eid: int, **fields) -> dict:
    fields = {k: v for k, v in fields.items() if k in ("name", "model", "reviewer_model", "arbiter_model") and v}
    if fields:
        _set("engagements", eid, **fields, updated_at=time.time())
    return get(eid)


def workbooks(eid: int) -> list[dict]:
    ids = [r["file_id"] for r in _q("SELECT file_id FROM eng_files WHERE engagement_id=? ORDER BY added_at", eid)]
    out = []
    for fid in ids:
        rec = library.get(fid, full=True)
        if rec:
            w = {k: rec.get(k) for k in ("id", "filename", "size", "uploaded_at", "status", "step", "pct", "error",
                                         "sheets", "line_items", "target_name", "project_name", "valuation_date",
                                         "identity_confirmed",
                                         "db_path", "source_path", "identity", "processed_at", "started_at",
                                         "build_secs")}
            w["sheet_names"] = _sheet_names(w) if w["status"] == "done" and w["db_path"] else []
            out.append(w)
    return out


_SHEET_NAMES: dict[tuple, list[str]] = {}


def _sheet_names(w: dict) -> list[str]:
    """A processed workbook's sheets, read once: the page polls every couple of seconds, and a model.db can be
    busy for a moment (its link tables being built); a busy file shows no sheet list rather than failing the page."""
    key = (w["db_path"], w.get("processed_at"))
    if key not in _SHEET_NAMES:
        try:
            with rodb.connect(w["db_path"], timeout=2) as m:
                _SHEET_NAMES[key] = [r[0] for r in m.execute("SELECT sheet FROM sheets ORDER BY rowid")]
        except sqlite3.Error:
            return []  # not cached: tried again on the next poll
    return _SHEET_NAMES[key]


def documents(eid: int) -> list[dict]:
    cols = ("id, engagement_id, filename, kind, size, uploaded_at, status, step, pct, error, pages, n_tables, n_flagged, "
            "processed_at, facts_status, facts_step, facts_error, facts_notes, review_summary, loop_json, started_at, "
            "doc_secs, facts_started_at, facts_secs, tables_status, tables_left, tables_step, tables_error, "
            "tables_started_at, tables_secs")
    out = _q(f"SELECT {cols} FROM documents WHERE engagement_id=? ORDER BY uploaded_at", eid)
    for d in out:
        d["loop"] = json.loads(d.pop("loop_json") or "null")
    return out


def facts(eid: int) -> list[dict]:
    out = []
    for f in _q("SELECT * FROM facts WHERE engagement_id=? ORDER BY document_id, n", eid):
        for k in ("check_json", "review_json", "final_json", "agent_json"):
            f[k.removesuffix("_json")] = json.loads(f.pop(k) or "null")
        out.append(f)
    return out


def reference(eid: int) -> list[dict]:
    """The facts to navigate by: approved ones as approved (with edits); if none approved yet, those that pass
    the code checks, marked unapproved."""
    fs = facts(eid)
    approved = [{**{k: f[k] for k in ("id", *FACT_FIELDS)}, **(f["final"] or {}), "approved": True}
                for f in fs if f["status"] == "approved"]
    if approved:
        return approved
    return [{**{k: f[k] for k in ("id", *FACT_FIELDS)}, "approved": False} for f in fs
            if f["status"] != "rejected" and (f["check"] or {}).get("ok")
            and (f["agent"] or {}).get("status") != "withdrawn"]


def roles(eid: int) -> dict:
    out = {}
    for r in _q("SELECT * FROM roles WHERE engagement_id=?", eid):
        out[r["role"]] = {"kind": r["kind"], "id": r["ref_id"], "sheets": json.loads(r["sheets_json"] or "null"),
                          "why": json.loads(r["why_json"] or "[]"), "confirmed": bool(r["confirmed"])}
    return out


def get(eid: int) -> dict | None:
    rows = _q("SELECT * FROM engagements WHERE id=?", eid)
    if not rows:
        return None
    e = rows[0]
    for k in ("roles_suggested", "compare_json", "map_json", "overlay_json", "profile_json", "schedule_json"):
        e[k.removesuffix("_json")] = json.loads(e.pop(k) or "null")
    wbs = workbooks(eid)
    if _agents_check_dates(eid, wbs):
        wbs = workbooks(eid)
    for w in wbs:
        w.pop("db_path", None)
        w.pop("source_path", None)
        ident = w.pop("identity", None) or {}
        w["identity_notes"] = ident.get("notes")
        conf = ident.get("confirmed") or {}
        w["identity_by"] = conf.get("by") or ("you" if w.get("identity_confirmed") else None)
        w["identity_why"] = conf.get("why") or []
        w["identity_check"] = ident.get("auto_check")  # the agents' check of the date, where it didn't confirm
    out = {**e, "documents": documents(eid), "workbooks": wbs, "facts": facts(eid), "roles": roles(eid),
           "session": _session(eid), "now": time.time()}
    _auto_review(out)
    _maybe_suggest(eid, out)
    out["roles_pending"] = eid in _ROLE_JOBS
    out["jobs"] = jobs_view(eid)
    return out


def delete(eid: int) -> None:
    for d in documents(eid):
        remove_document(d["id"])
    with _lock, _conn() as db:
        for t, col in (("facts", "engagement_id"), ("roles", "engagement_id"), ("eng_files", "engagement_id"),
                       ("engagements", "id")):
            db.execute(f"DELETE FROM {t} WHERE {col}=?", (eid,))


# ---- uploads ------------------------------------------------------------------------------------------------

def _model(eid: int) -> str:
    """The engagement's model, as the header's Models selector has it."""
    rows = _q("SELECT model FROM engagements WHERE id=?", eid)
    return (rows[0]["model"] if rows else None) or DEFAULT_MODEL


def add_upload(eid: int, tmp: Path, filename: str, sha: str) -> dict:
    """Workbooks -> the shared library (deduplicated across the app); reports -> this engagement's documents."""
    ext = Path(filename).suffix.lower()
    if ext in library.SUPPORTED:  # identified and summarised with the engagement's model (the Models selector)
        status, rec = library.add_upload(tmp, filename, sha, _model(eid))
        with _lock, _conn() as db:
            db.execute("INSERT OR IGNORE INTO eng_files VALUES (?,?,?)", (eid, rec["id"], time.time()))
        _touch(eid)
        return {"status": status, "kind": "workbook", "id": rec["id"], "filename": rec["filename"]}
    if ext not in REPORT_TYPES:
        tmp.unlink(missing_ok=True)
        raise ValueError(f"{filename}: upload reports as .pdf or .pptx and models as .xlsx or .xlsm "
                         "(save .xlsb / .xls / .ppt / .docx in one of those formats first)")
    have = _q("SELECT id FROM documents WHERE engagement_id=? AND sha256=?", eid, sha)
    if have:
        tmp.unlink(missing_ok=True)
        return {"status": "duplicate", "kind": "document", "id": have[0]["id"], "filename": filename}
    dest = UPLOADS / sha[:12] / Path(filename).name
    dest.parent.mkdir(parents=True, exist_ok=True)
    tmp.replace(dest)
    out_dir = DOCS / f"{Path(filename).stem}__{sha[:8]}"
    did = _exec("""INSERT INTO documents(engagement_id, sha256, filename, kind, size, uploaded_at, source_path, out_dir,
                   status, step, pct) VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
                eid, sha, filename, ext.lstrip("."), dest.stat().st_size, time.time(), str(dest), str(out_dir),
                "queued", "Waiting to start", 0)
    _touch(eid)
    _jobs.put(("doc", did))
    return {"status": "queued", "kind": "document", "id": did, "filename": filename}


# ---- this year's valuation date, checked by the agents ---------------------------------------------------------
# identify.py reads a workbook's valuation date off one cell. Before a person is asked to check it, the agents weigh
# the evidence the files already hold (no model calls): other cells labelled like it, the file name, a year on from
# last year's valuation date, the model's own financial-year end. Confirmed when two or more agree and none
# disagrees; otherwise the person is asked, and told why.

_DATE_CHECKED: set = set()  # (workbook, date, last year's): weighed in this run
_MONTH = r"(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|june?|july?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)"
_FILE_MONTH = re.compile(rf"(?i)(?<![a-z]){_MONTH}(?![a-z])[\s_.-]*(\d{{4}}|\d{{2}})(?!\d)")


def _file_months(name: str) -> set[tuple[int, int]]:
    """(year, month) a file name gives in words ("Jun 26", "Jun-26", "June 2026"): a date stamp in digits
    (20260521) is when the file was saved, not its valuation date, so it doesn't count."""
    out = set()
    for m in _FILE_MONTH.finditer(name or ""):
        y = int(m[2])
        out.add((y + 2000 if y < 100 else y, [x[:3] for x in (s.lower() for s in MONTHS)].index(m[1][:3].lower()) + 1))
    return out


def _year_on(iso: str) -> str:
    """A date a year later, a month's end kept at the month's end."""
    import calendar
    y, m, d = map(int, iso[:10].split("-"))
    last = d == calendar.monthrange(y, m)[1]
    return f"{y + 1:04d}-{m:02d}-{calendar.monthrange(y + 1, m)[1] if last else min(d, calendar.monthrange(y + 1, m)[1]):02d}"


def _date_evidence(eid: int, w: dict, cited: str | None) -> dict:
    """What agrees and what disagrees with a workbook's valuation date, besides the cell identify cited:
    {"agree": [...], "disagree": [...]}. A row whose label names another date ("roll forward to 30/9/2025") counts
    neither way; the financial-year end counts only where the model shows it or a person set it (the profile's
    fallback is last year's valuation date's month, which the year-on test already weighs)."""
    import calendar
    import chartdata
    import identify
    date = w["valuation_date"]
    y, m, d = map(int, date.split("-"))
    agree, disagree = [], []
    others = [x for x in identify.candidates(w["db_path"])["valuation_date"] if x["where"] != cited]
    same = [x for x in others if x["value"] == date]
    if same:
        agree.append(f"{same[0]['where']} ({same[0]['why']}) holds it too" + (f", and {len(same) - 1} more" if len(same) > 1 else ""))
    against = [x for x in others if x.get("rank", 1) <= 1 and x.get("value") and x["value"] != date]
    if against:
        disagree.append(f"{against[0]['where']} ({against[0]['why']}) holds {against[0]['value']}")
    months = _file_months(w["filename"])
    if (y, m) in months:
        agree.append(f"the file name says {MONTHS[m - 1][:3]} {y}")
    elif months:
        disagree.append("the file name says " + ", ".join(f"{MONTHS[mm - 1][:3]} {yy}" for yy, mm in sorted(months)))
    prior = _prior_vd(eid)
    if prior and date <= prior[:10]:
        disagree.append(f"not after last year's valuation date ({prior[:10]})")
    elif prior and date == _year_on(prior):
        agree.append(f"a year after last year's valuation date ({prior[:10]})")
    fy = _profile(eid).get("fy_end_month")
    if not fy:
        with closing(rodb.connect(w["db_path"])) as db:
            fy, _ = chartdata.fy_end_detect(db)
    if fy and m == fy and d == calendar.monthrange(y, m)[1]:
        agree.append(f"the model's financial year ends in {MONTHS[fy - 1]}")
    return {"agree": agree, "disagree": disagree}


def _agents_check_dates(eid: int, wbs: list[dict]) -> bool:
    """This year's client model's valuation date, weighed once per date and last year's date (in this run, and kept
    with the workbook's identity across a restart): confirmed by the agents where two or more signals agree and
    none disagrees. Waits for last year's valuation date: without it a model's own stale date can't be told.
    True if it confirmed one."""
    cur = roles(eid).get("current_model") or {}
    w = next((x for x in wbs if x["id"] == cur.get("id") and cur.get("kind") == "workbook"), None)
    if not w or w.get("status") != "done" or w.get("identity_confirmed") or not w.get("valuation_date") or not w.get("db_path"):
        return False
    prior = _prior_vd(eid)
    key = (w["id"], w["valuation_date"], prior)
    if not prior or key in _DATE_CHECKED:
        return False
    _DATE_CHECKED.add(key)
    ident = (library.get(w["id"], full=True) or {}).get("identity") or {}
    done = ident.get("auto_check") or {}
    if (done.get("date"), done.get("prior")) == (w["valuation_date"], prior):
        return False
    try:
        ev = _date_evidence(eid, w, ident.get("valuation_date_evidence"))
    except Exception:  # the check is a bonus: the person is asked as before
        traceback.print_exc()
        return False
    library.note_identity(w["id"], auto_check={"date": w["valuation_date"], "prior": prior, **ev, "at": time.time()})
    if len(ev["agree"]) >= 2 and not ev["disagree"]:
        library.confirm_identity(w["id"], "agents", ev["agree"])
        return True
    return False


def confirm_date(eid: int, fid: int, valuation_date: str) -> dict:
    """A person's check of a workbook's valuation date (the roll-forward runs from last year's to this year's)."""
    w = next((w for w in workbooks(eid) if w["id"] == fid), None)
    if not w:
        raise ValueError("that workbook isn't in this engagement")
    if not re.match(r"^\d{4}-\d{2}-\d{2}$", valuation_date or ""):
        raise ValueError("give the date as YYYY-MM-DD")
    if valuation_date == w.get("valuation_date"):  # right as it is: confirmed, nothing re-linked
        library.confirm_identity(fid, "you")
    else:
        library.set_identity(fid, w.get("target_name"), w.get("project_name"), valuation_date)
    _touch(eid)
    return get(eid)


def remove_workbook(eid: int, fid: int) -> None:
    """Unlink from this engagement only; the workbook stays in the shared library."""
    with _lock, _conn() as db:
        db.execute("DELETE FROM eng_files WHERE engagement_id=? AND file_id=?", (eid, fid))
        db.execute("DELETE FROM roles WHERE engagement_id=? AND kind='workbook' AND ref_id=?", (eid, fid))


def remove_document(did: int) -> None:
    d = _doc(did)
    if not d:
        return
    if d["status"] in ("queued", "processing") or d["facts_status"] == "running":
        raise ValueError("wait for this document to finish processing before removing it")
    with _lock, _conn() as db:
        db.execute("DELETE FROM facts WHERE document_id=?", (did,))
        db.execute("DELETE FROM roles WHERE kind='document' AND ref_id=?", (did,))
        db.execute("DELETE FROM documents WHERE id=?", (did,))
    out_dir, src = Path(d["out_dir"]), Path(d["source_path"])
    if out_dir.is_relative_to(DOCS):
        shutil.rmtree(out_dir, ignore_errors=True)
    others = _q("SELECT id FROM documents WHERE sha256=?", d["sha256"])
    if src.is_relative_to(UPLOADS) and not others and not library.by_sha(d["sha256"]):
        shutil.rmtree(src.parent, ignore_errors=True)


def _touch(eid: int) -> None:
    _set("engagements", eid, updated_at=time.time())


def _doc(did: int) -> dict | None:
    rows = _q("SELECT * FROM documents WHERE id=?", did)
    return rows[0] if rows else None


def document(did: int) -> dict | None:
    """Pages (Markdown) and tables (with checks) of a processed report."""
    d = _doc(did)
    if not d:
        return None
    doc = json.loads(d.pop("doc_json") or "null")
    d.pop("source_path")
    d["loop"] = json.loads(d.pop("loop_json", None) or "null")
    return {**d, "doc": doc}


def table_png(did: int, tid: str) -> Path | None:
    d = _doc(did)
    p = Path(d["out_dir"]) / "tables" / f"{tid}.png" if d else None
    return p if p and p.is_file() and p.parent.parent == Path(d["out_dir"]) else None


def settle_table(did: int, tid: str, action: str, markdown: str | None = None) -> dict:
    """A person's decision on a table: approve the first read, take the reviewer's read, or edit it."""
    with _DOC_LOCK:
        return _settle_table(did, tid, action, markdown)


def _settle_table(did: int, tid: str, action: str, markdown: str | None = None) -> dict:
    d = _doc(did)
    doc = json.loads(d["doc_json"])
    t = next((t for t in doc["tables"] if t["id"] == tid), None)
    if t is None:
        raise ValueError("no such table")
    if action == "approve":  # the text as it stands: the agents' settled or latest correction, else the first read
        t.update(status="approved", final_markdown=docingest.latest(t))
    elif action == "use_second":
        t.update(status="approved", final_markdown=(t.get("resolution") or {}).get("second_markdown")
                 or t.get("second_markdown"))
    elif action == "edit":
        t.update(status="edited", final_markdown=markdown or "")
    elif action == "reset":
        t.pop("final_markdown", None)
        t.pop("final_check", None)
        t["status"] = "verified" if (t.get("check") or t.get("second_read") or {}).get("ok") else "flagged"
    else:
        raise ValueError(f"unknown action {action}")
    if t.get("final_markdown") is not None and t.get("text_lines"):
        t["final_check"] = docingest.check_text_layer(t["final_markdown"], t["text_lines"], t.get("title"))
    _save_doc(did, doc)
    changed = _recheck_facts(did, doc)
    auto_decide(did)  # facts only a check held up, that now pass: the agents' agreement decides them
    # facts handed to you that this table bears on (they rest on its page, or their checks now read differently)
    # go back to the loop with the page as it now reads; the others stay with you, since the same loop on the same
    # text would come out the same
    pg = reportfacts.pages(doc["markdown"])
    n, after_loop = _reopen(did, lambda f: f["id"] in changed or t["page"] in reportfacts.fact_pages(f, pg))
    return {**document(did), "reopened": n, "after_loop": after_loop}


def _reopen(did: int, bears_on) -> tuple[int, bool]:
    """Facts handed to a person that something they rest on changed for (bears_on(fact row)) go back to the loop:
    now, with a loop already queued, or after the one running (it took the facts as they were). Returns how many,
    and whether they wait for the running loop."""
    reopen = [f["id"] for f in _q("SELECT * FROM facts WHERE document_id=? AND status='pending'", did)
              if (a := json.loads(f["agent_json"] or "null") or {}).get("status") == "escalated" and not a.get("held")
              and bears_on(f)]
    status = _doc(did)["facts_status"]
    if reopen and status == "running":
        _AFTER.setdefault(did, set()).update(reopen)
    elif reopen and status != "queued":
        resolve_facts(did, reopen)
    elif reopen and _REOPEN.get(did) is not None:
        _REOPEN[did] |= set(reopen)  # a loop on other reopened facts hasn't started: it takes these too
    return len(reopen), bool(reopen) and status == "running"


_DOC_LOCK = threading.RLock()  # the background reader and a person's table decisions save into the same document
_READING: set[int] = set()     # documents whose remaining tables are being read in the background


def _merge_tables(did: int, updated: list[dict], only_if) -> dict:
    """Put tables read or settled elsewhere into the stored document, each only if the stored one still qualifies
    (e.g. still unread): a person's decision made meanwhile is never overwritten. Returns the saved document."""
    with _DOC_LOCK:
        doc = json.loads(_doc(did)["doc_json"])
        by_id = {t["id"]: t for t in updated}
        doc["tables"] = [by_id[t["id"]] if t["id"] in by_id and only_if(t) else t for t in doc["tables"]]
        _save_doc(did, doc)
        return doc


def _start_rest(did: int) -> None:
    if did in _READING:
        return
    _READING.add(did)
    threading.Thread(target=_read_rest, args=(did,), daemon=True, name=f"report-tables-{did}").start()


def _read_rest(did: int) -> None:
    """Read the report's remaining tables (the key ones were read first), saving each as it's done; then the review
    loop on those it flags, and the facts checked again against the fuller text. Runs beside the worker, so the
    facts and the model steps don't wait for it."""
    d0 = _doc(did)
    with calllog.tag(engagement=d0["engagement_id"] if d0 else None, document=did, step="background tables"):
        _read_rest_tagged(did)


def _read_rest_tagged(did: int) -> None:
    try:
        d = _doc(did)
        eid = d["engagement_id"]
        e = _q("SELECT model, reviewer_model, arbiter_model FROM engagements WHERE id=?", eid)[0]
        model, reviewer = e["model"] or DEFAULT_MODEL, e["reviewer_model"] or DEFAULT_REVIEWER
        doc = json.loads(d["doc_json"])
        todo = [t["id"] for t in doc["tables"] if t.get("deferred") and t.get("status") == "unread"]
        _set("documents", did, tables_status="reading", tables_left=len(todo), tables_started_at=time.time(),
             tables_step=f"Reading {len(todo)} more tables", tables_error=None)

        def saved(t):
            cur = _merge_tables(did, [t], lambda c: c.get("status") == "unread")
            left = sum(bool(x.get("deferred")) and x.get("status") == "unread" for x in cur["tables"])
            _set("documents", did, tables_left=left, tables_step=f"{left} more tables to read")

        docingest.read_rest(doc, d["out_dir"], model, reviewer, _logger(eid), on_table=saved)
        cur = json.loads(_doc(did)["doc_json"])
        if any(t.get("status") == "flagged" and t["id"] in todo for t in cur["tables"]):
            _set("documents", did, tables_step="Table review loop on the flagged tables")
            res = docingest.resolve_tables(cur, d["out_dir"], model, reviewer, None, _logger(eid), ids=set(todo),
                                           arbiter_model=e["arbiter_model"] or DEFAULT_ARBITER)
            _merge_tables(did, [t for t in cur["tables"] if t.get("resolution") and t["id"] in todo],
                          lambda c: c.get("status") == "flagged")
            _learn(did, "tables", res["episodes"])
        else:
            res = {"resolved": 0, "escalated": 0}
        prev = (json.loads(_doc(did)["loop_json"] or "{}").get("tables") or {})
        _note_loop(did, tables={"resolved": (prev.get("resolved") or 0) + res["resolved"], "escalated": res["escalated"],
                                "at": time.time(), "version": docingest.LOOP_VERSION})  # nothing starts another
        with _DOC_LOCK:
            _recheck_facts(did, json.loads(_doc(did)["doc_json"]))
        _set("documents", did, tables_status="done", tables_left=0, tables_step="Done",
             tables_secs=round(time.time() - (_doc(did)["tables_started_at"] or time.time()), 1))
        _touch(eid)
    except Exception as ex:
        traceback.print_exc()
        _set("documents", did, tables_status="error", tables_error=friendly(ex))
    finally:
        _READING.discard(did)


def _save_doc(did: int, doc: dict) -> None:
    md = docingest.render(doc)
    d = _doc(did)
    Path(d["out_dir"], "document.md").write_text(md, encoding="utf-8")
    flagged = sum(t.get("status") in ("flagged", "error") or (t.get("status") == "unread" and not t.get("deferred"))
                  for t in doc["tables"])
    _set("documents", did, doc_json=json.dumps(doc, default=str), n_flagged=flagged,
         n_tables=sum(t.get("status") != "figure" for t in doc["tables"]), pages=len(doc["pages"]))


def _recheck_facts(did: int, doc: dict) -> set[int]:
    """Table decisions change the page text facts are checked against (no model calls). An approval the agents
    made rested on the checks passing, so if they now fail the fact goes back to a person. Returns the facts whose
    checks came out differently, or whose number did (worked out by code: reportfacts.settle_value)."""
    pg = reportfacts.pages(doc["markdown"])
    failing = lambda c: [i["text"] for i in (c or {}).get("items", []) if not i["ok"]]
    changed = set()
    for f in _q("SELECT * FROM facts WHERE document_id=?", did):
        agent = json.loads(f["agent_json"] or "null") or {}
        value = reportfacts.settle_value({k: f[k] for k in FACT_FIELDS})["value"]
        chk = reportfacts.check({**f, "value": value, "waivers": agent.get("waivers")}, pg)
        if failing(chk) != failing(json.loads(f["check_json"] or "null")) or value != f["value"]:
            changed.add(f["id"])
        rv = json.loads(f["review_json"] or "null")
        if rv and rv.get("suggestion"):
            rv["suggestion"]["check"] = reportfacts.check(rv["suggestion"], pg)
        fields = {"check_json": json.dumps(chk), "review_json": json.dumps(rv), "value": value}
        o = agent.get("open") or {}
        if f["status"] == "pending" and agent.get("status") == "escalated" and chk["ok"] and \
                (o.get("accepted") or (o.get("reason") or "").startswith("accepted by the reviewer")):
            # the reviewer had accepted it and only a check held it up: that check now passes
            agent["thread"] = agent.get("thread", []) + [{"round": "recheck", "note": "the checks pass on a fresh check"}]
            agent.update(status="agreed", round="recheck")
            agent.pop("open", None)
            fields["agent_json"] = json.dumps(agent)
        if f["decided_by"] == "agents" and f["status"] == "approved" and not chk["ok"]:
            a = json.loads(f["agent_json"] or "null") or {"thread": []}
            a.update(status="escalated", open={"verdict": "object", "reason": "a table on its page changed and the checks "
                                               "now fail: " + "; ".join(i["text"] for i in chk["items"] if not i["ok"]),
                                               "correction": None})
            fields.update(status="pending", decided_by=None, agent_json=json.dumps(a))
        _set("facts", f["id"], **fields)
    return changed


# ---- facts --------------------------------------------------------------------------------------------------

def extract_facts(did: int) -> None:
    d = _doc(did)
    if d["status"] != "done":
        raise ValueError("the document hasn't been read yet")
    _set("documents", did, facts_status="queued", facts_step="Waiting to start", facts_error=None)
    _jobs.put(("facts", did))


def set_fact(fact_id: int, action: str, fields: dict | None = None) -> dict:
    rows = _q("SELECT * FROM facts WHERE id=?", fact_id)
    if not rows:
        raise ValueError("no such fact")
    f = rows[0]
    if action not in ("approve", "use_suggestion", "edit", "reject", "reset"):
        raise ValueError(f"unknown action {action}")
    rv = json.loads(f["review_json"] or "null") or {}
    now = time.time()
    if action == "reset" and f["decided_by"] == "agents":  # a person undid the agents' decision: it's theirs now
        a = json.loads(f["agent_json"] or "null") or {}
        _set("facts", fact_id, agent_json=json.dumps({**a, "held": True}))
    _set("facts", fact_id, decided_by=None if action == "reset" else "you")
    if action == "approve":
        _set("facts", fact_id, status="approved", final_json=None, updated_at=now)
    elif action == "use_suggestion":
        s = rv.get("suggestion")
        if not s:
            raise ValueError("the reviewer made no correction to use")
        _set("facts", fact_id, status="approved", final_json=json.dumps({k: s.get(k) for k in FACT_FIELDS}), updated_at=now)
    elif action == "edit":
        final = {k: f[k] for k in FACT_FIELDS}
        final.update({k: v for k, v in (fields or {}).items() if k in FACT_FIELDS})
        if "value" not in (fields or {}):
            reportfacts.settle_value(final)
        d = _doc(f["document_id"])
        pg = reportfacts.pages(json.loads(d["doc_json"])["markdown"])
        final["check"] = reportfacts.check(final, pg)
        _set("facts", fact_id, status="approved", final_json=json.dumps(final), updated_at=now)
    elif action == "reject":
        _set("facts", fact_id, status="rejected", updated_at=now)
    elif action == "reset":
        _set("facts", fact_id, status="pending", final_json=None, updated_at=now)
    else:
        raise ValueError(f"unknown action {action}")
    _touch(f["engagement_id"])
    return next(x for x in facts(f["engagement_id"]) if x["id"] == fact_id)


def agreed(f: dict) -> bool:
    """Both models agree on the fact and the code checks pass (the review loop's outcome, or, for facts from
    before the loop, the reviewer's accept)."""
    a = f.get("agent")
    ok = (f.get("check") or {}).get("ok")
    return bool(ok and (a["status"] == "agreed" if a else (f.get("review") or {}).get("verdict") == "accept"))


def auto_decide(did: int) -> dict:
    """The agents' agreement is the decision: facts they agree on (and that pass the checks) are approved, facts
    they agree to withdraw are rejected, both marked as decided by the agents. Facts a person has decided, or
    whose agents' decision a person undid, are left alone."""
    n = {"approved": 0, "rejected": 0}
    now = time.time()
    for f in _q("SELECT id, check_json, agent_json FROM facts WHERE document_id=? AND status='pending'", did):
        a = json.loads(f["agent_json"] or "null") or {}
        if a.get("held"):
            continue
        if a.get("status") == "agreed" and (json.loads(f["check_json"] or "null") or {}).get("ok"):
            _set("facts", f["id"], status="approved", final_json=None, decided_by="agents", updated_at=now)
            n["approved"] += 1
        elif a.get("status") == "withdrawn":
            _set("facts", f["id"], status="rejected", decided_by="agents", updated_at=now)
            n["rejected"] += 1
    return n


def approve_passed(eid: int) -> int:
    """Approve every pending fact the agents agreed on and that passes the code checks."""
    n = 0
    for f in facts(eid):
        if f["status"] == "pending" and agreed(f):
            set_fact(f["id"], "approve")
            n += 1
    return n


# ---- the review loop and its lessons --------------------------------------------------------------------------

def _names(eid: int, did: int | None = None) -> list[str]:
    """What identifies this engagement, for the lessons' anonymity check: its name, file names, targets and the
    report's identity facts."""
    e = _q("SELECT name FROM engagements WHERE id=?", eid)
    out = [e[0]["name"]] if e else []
    out += [d["filename"] for d in documents(eid)]
    for w in workbooks(eid):
        out += [w["filename"], w.get("target_name"), w.get("project_name")]
    out += [f["value_text"] for f in facts(eid) if f["category"] == "identity"
            and re.search(r"name|client|target|project|asset|company|vendor|purchaser|owner", f["key"] or "")]
    return [x for x in out if x]


def _learn(did: int, scope: str, episodes: list[dict]) -> dict | None:
    """Distil a loop's episodes into lessons (reviewer model), and note the result on the document."""
    if not episodes:
        return None
    d = _doc(did)
    eid = d["engagement_id"]
    e = _q("SELECT reviewer_model FROM engagements WHERE id=?", eid)[0]
    md = json.loads(d["doc_json"] or "{}").get("markdown", "")
    try:
        got = lessons.distil(scope, episodes, _names(eid, did), md, e["reviewer_model"] or DEFAULT_REVIEWER,
                             _logger(eid), engagement=eid)
    except Exception as ex:  # learning is a bonus: never fail the step for it
        traceback.print_exc()
        got = {"error": friendly(ex)}
    _note_loop(did, **{f"lessons_{scope}": got})
    return got


def _note_loop(did: int, **fields) -> None:
    d = _q("SELECT loop_json FROM documents WHERE id=?", did)
    loop = json.loads((d[0]["loop_json"] if d else None) or "{}")
    loop.update(fields)
    _set("documents", did, loop_json=json.dumps(loop, default=str))


def resolve_tables(did: int) -> None:
    d = _doc(did)
    if d["status"] != "done":
        raise ValueError("the document hasn't been read yet")
    _set("documents", did, status="queued", step="Waiting for the review loop", pct=0)
    _jobs.put(("resolve_tables", did))


_REOPEN: dict[int, set[int]] = {}  # document -> the facts a table decision sent back to the loop (absent: all open)
_AFTER: dict[int, set[int]] = {}   # ... reopened while a loop was running: a loop of their own once it's done


def _next_loop(did: int) -> None:
    after = _AFTER.pop(did, None)
    if after:
        try:
            resolve_facts(did, list(after))
        except ValueError:  # settled meanwhile
            pass


def resolve_facts(did: int, ids: list[int] | None = None) -> None:
    """The fact review loop on the facts still open, or on just `ids`."""
    open_ids = [r["id"] for r in _q("SELECT id, agent_json FROM facts WHERE document_id=? AND status='pending'", did)
                if (json.loads(r["agent_json"] or "null") or {}).get("status") not in ("agreed", "withdrawn")
                and (ids is None or r["id"] in ids)]
    if not open_ids:
        raise ValueError("no open facts: the agents agree on every fact still waiting for a decision")
    if ids is None:
        _REOPEN.pop(did, None)
    else:
        _REOPEN[did] = set(open_ids)
    _set("documents", did, facts_status="queued", facts_error=None, facts_step="Waiting for the review loop"
         + ("" if ids is None else f" on {len(open_ids)} fact(s) the table decision bears on"))
    _jobs.put(("resolve_facts", did))


def _run_table_loop(did: int, doc: dict) -> None:
    d = _doc(did)
    eid = d["engagement_id"]
    e = _q("SELECT model, reviewer_model, arbiter_model FROM engagements WHERE id=?", eid)[0]
    prog = lambda frac, msg: _set("documents", did, pct=round(frac, 3), step=msg)
    res = docingest.resolve_tables(doc, d["out_dir"], e["model"] or DEFAULT_MODEL, e["reviewer_model"] or DEFAULT_REVIEWER,
                                   prog, _logger(eid), arbiter_model=e["arbiter_model"] or DEFAULT_ARBITER)
    _save_doc(did, doc)
    _note_loop(did, tables={"resolved": res["resolved"], "escalated": res["escalated"], "at": time.time(),
                            "version": docingest.LOOP_VERSION},
               lessons_tables=None)  # this loop's lessons replace the last one's note
    _learn(did, "tables", res["episodes"])


def _resolve_tables_job(did: int) -> None:
    d = _doc(did)
    _set("documents", did, status="processing", step="Table review loop on the flagged tables", pct=0)
    eid = d["engagement_id"]
    with _DOC_LOCK:  # today's check first: tables it now passes need no model at all
        doc = json.loads(_doc(did)["doc_json"])
        rechecked = docingest.recheck(doc)
        if rechecked:
            _save_doc(did, doc)
    e = _q("SELECT model, reviewer_model, arbiter_model FROM engagements WHERE id=?", eid)[0]
    prog = lambda frac, msg: _set("documents", did, pct=round(frac, 3), step=msg)
    res = docingest.resolve_tables(doc, d["out_dir"], e["model"] or DEFAULT_MODEL, e["reviewer_model"] or DEFAULT_REVIEWER,
                                   prog, _logger(eid), arbiter_model=e["arbiter_model"] or DEFAULT_ARBITER)
    doc = _merge_tables(did, [t for t in doc["tables"] if t.get("resolution")], lambda c: c.get("status") == "flagged")
    _note_loop(did, tables={"resolved": res["resolved"] + rechecked, "escalated": res["escalated"], "at": time.time(),
                            "version": docingest.LOOP_VERSION}, lessons_tables=None)
    _learn(did, "tables", res["episodes"])
    with _DOC_LOCK:
        _recheck_facts(did, json.loads(_doc(did)["doc_json"]))
    _set("documents", did, status="done", step="Done", pct=1.0)


def _resolve_facts_job(did: int) -> None:
    d = _doc(did)
    eid = d["engagement_id"]
    e = _q("SELECT model, reviewer_model, arbiter_model FROM engagements WHERE id=?", eid)[0]
    md = json.loads(d["doc_json"])["markdown"]
    pg = reportfacts.pages(md)
    only = _REOPEN.pop(did, None)
    rows = [r for r in _q("SELECT * FROM facts WHERE document_id=? AND status='pending' ORDER BY n", did)
            if (json.loads(r["agent_json"] or "null") or {}).get("status") not in ("agreed", "withdrawn")
            and (only is None or r["id"] in only)]
    if not rows:
        _set("documents", did, facts_status="done", facts_step="Done")
        _next_loop(did)
        return
    fs = []
    for r in rows:
        f = {"id": r["id"], "origin": r["origin"], **{k: r[k] for k in FACT_FIELDS},
             "check": json.loads(r["check_json"] or "null"), "review": json.loads(r["review_json"] or "null") or {}}
        prev = json.loads(r["agent_json"] or "null")
        f["waivers"] = (prev or {}).get("waivers") or []
        reportfacts.settle_value(f)  # the number from the text, by code (none for a name or an unpicked range)
        f["check"] = reportfacts.check(f, pg)  # today's checks (spacing-tolerant, waivers honoured)
        if prev and prev.get("status") in ("escalated", "open") and prev.get("open"):
            o = prev["open"]
            accepted = o.get("accepted") or (o.get("reason") or "").startswith("accepted by the reviewer")
            f["review"] = {**f["review"], "verdict": "accept" if accepted else (o.get("verdict") or "correct"),
                           "reason": o.get("reason")}
        fs.append(f)
    _set("documents", did, facts_status="running", facts_step="Fact review loop")
    loop = reportfacts.resolve(md, fs, e["model"] or DEFAULT_MODEL, e["reviewer_model"] or DEFAULT_REVIEWER, _logger(eid),
                               lambda frac, msg: _set("documents", did, facts_step=msg),
                               arbiter_model=e["arbiter_model"] or DEFAULT_ARBITER)
    now = time.time()
    for f, r in zip(fs, rows):
        prev = json.loads(r["agent_json"] or "null")
        if prev and prev.get("thread"):  # earlier rounds stay in the thread
            f["agent"]["thread"] = prev["thread"] + f["agent"]["thread"][1:]
        _set("facts", f["id"], **{k: f.get(k) for k in FACT_FIELDS}, check_json=json.dumps(f["check"]),
             review_json=json.dumps(f.get("review")), agent_json=json.dumps(f["agent"]), updated_at=now)
    _note_loop(did, facts={**loop["summary"], **auto_decide(did), "at": now}, lessons_facts=None)
    _set("documents", did, facts_step="Writing down what the loop taught")
    _learn(did, "facts", loop["episodes"])
    _set("documents", did, facts_status="done", facts_step="Done")
    _touch(eid)
    _next_loop(did)


def calls_view(eid: int) -> dict:
    """Tokens, cost and time for an engagement: model calls by file and by step (calllog.py), plus how long each
    job took on the clock (a job's time also counts waiting for rate-limit room and work without model calls)."""
    wbs, docs = workbooks(eid), documents(eid)
    e = _q("SELECT * FROM engagements WHERE id=?", eid)
    if not e:
        raise ValueError("no such engagement")
    e = e[0]
    jobs = []
    for d in docs:
        jobs += [{"file": d["filename"], "document": d["id"], "job": "read the report (key tables)", "secs": d["doc_secs"]},
                 {"file": d["filename"], "document": d["id"], "job": "read the other tables (background)", "secs": d["tables_secs"]},
                 {"file": d["filename"], "document": d["id"], "job": "key facts and their review loop", "secs": d["facts_secs"]}]
    for w in wbs:
        took = (w["processed_at"] - w["started_at"]) if w.get("processed_at") and w.get("started_at") and \
            w["processed_at"] > w["started_at"] else None
        jobs.append({"file": w["filename"], "workbook": w["id"], "job": "process the workbook", "secs": took,
                     "build_secs": w.get("build_secs")})
    jobs += [{"file": None, "job": label, "secs": e.get(f"{k}_secs")} for k, label in
             (("compare", "compare models"), ("map", "map"), ("overlay", "python overlay"))]
    return {"calls": calllog.breakdown(eid, [w["id"] for w in wbs]), "jobs": [j for j in jobs if j["secs"]],
            "names": {"documents": {d["id"]: d["filename"] for d in docs}, "workbooks": {w["id"]: w["filename"] for w in wbs}},
            "log_file": str(calllog.DB.relative_to(ROOT))}


def call_view(eid: int, cid: int) -> dict | None:
    return calllog.call(cid, eid, [w["id"] for w in workbooks(eid)])


def lessons_view() -> dict:
    return {"curated": lessons.curated(), "learned": lessons.load()["lessons"], "rules_file": "docs/report_rules.md"}


# ---- roles --------------------------------------------------------------------------------------------------

def _wb_inputs(eid: int) -> list[dict]:
    out = []
    for w in workbooks(eid):
        if w["status"] == "done" and w["db_path"] and Path(w["db_path"]).exists():
            extlinks.ensure(w["source_path"], w["db_path"])
            out.append(w)
    return out


_ROLE_JOBS: set[int] = set()  # engagements with a role suggestion queued
_AUTO: set[tuple[str, int]] = set()  # (job, document) queued by _auto_review, so a poll doesn't queue it twice


def _auto_review(e: dict) -> None:
    """Start the report's review work without a button, for every report that's been read: the loop on flagged
    tables if it hasn't run, then the key facts (extraction, review, the loop, the agents' approvals, lessons),
    or just the loop for facts from before it existed. Once per document: a failure is recorded and shown, not
    retried on every poll (the buttons retry)."""
    for d in e["documents"]:
        if d["status"] != "done" or busy_doc(d) or d["id"] in _READING or d.get("tables_status") in ("queued", "reading"):
            continue  # the background reader runs its own loop on the tables it reads
        loop = d.get("loop") or {}
        did = d["id"]
        if loop.get("check_version") != reportfacts.CHECK_VERSION and any(f["document_id"] == did for f in e["facts"]):
            changed = _recheck_facts(did, json.loads(_doc(did)["doc_json"]))  # checks improved: facts they held up can settle
            auto_decide(did)
            _note_loop(did, check_version=reportfacts.CHECK_VERSION)
            e["facts"] = facts(e["id"])
            if _reopen(did, lambda f: f["id"] in changed)[0]:  # handed to you, now read differently: back to the loop
                continue
        if d["n_flagged"] and (loop.get("tables") or {}).get("version") != docingest.LOOP_VERSION and not d["error"]:
            job = ("resolve_tables", did)
        elif not d["facts_status"] and not any(f["document_id"] == did for f in e["facts"]):
            job = ("facts", did)
        elif d["facts_status"] == "done" and "facts" not in loop and any(
                f["document_id"] == did and f["status"] == "pending" and not f.get("agent") for f in e["facts"]):
            job = ("resolve_facts", did)
        else:
            continue
        if job in _AUTO:
            continue
        _AUTO.add(job)
        if job[0] == "resolve_tables":
            _set("documents", did, status="queued", step="Waiting for the review loop", pct=0)
        else:
            _set("documents", did, facts_status="queued", facts_step="Waiting to start" if job[0] == "facts"
                 else "Waiting for the review loop", facts_error=None)
        _jobs.put(job)


def busy_doc(d: dict) -> bool:
    return d["status"] in ("queued", "processing") or d["facts_status"] in ("queued", "running")


def _roles_key(eid: int, wbs: list[dict], docs: list[dict]) -> str:
    """What a suggestion depends on: the processed files and the report facts it navigates by."""
    ref = [(f["id"], f["value_text"], f["page"], f["approved"]) for f in reference(eid)]
    return json.dumps([sorted(w["id"] for w in wbs if w["status"] == "done"),
                       sorted(d["id"] for d in docs if d["status"] == "done"), ref], default=str)


def _maybe_suggest(eid: int, e: dict) -> None:
    """Suggest roles by itself once the files are read, and again whenever the files or the report facts change,
    until every role is confirmed: structure first, then the report's evidence as it arrives."""
    if eid in _ROLE_JOBS or len(e["roles"]) == 4 and all(r["confirmed"] for r in e["roles"].values()):
        return
    wbs, docs = e["workbooks"], e["documents"]
    if not any(w["status"] == "done" for w in wbs) or any(busy for busy in
            [w["status"] in ("queued", "processing") for w in wbs] + [d["status"] in ("queued", "processing") for d in docs]
            + [d["facts_status"] in ("queued", "running") for d in docs]):
        return  # wait until nothing is being read
    if (e.get("roles_suggested") or {}).get("key") == _roles_key(eid, wbs, docs):
        return
    _ROLE_JOBS.add(eid)
    _jobs.put(("roles", eid))


def _suggest_job(eid: int) -> None:
    try:
        suggest_roles(eid, force=False)
    finally:
        _ROLE_JOBS.discard(eid)


_SUGGESTING = threading.Lock()  # one suggestion at a time: two at once wrote the same model.db and locked each other out


def suggest_roles(eid: int, force: bool = True) -> dict:
    """Suggest (doesn't overwrite confirmed roles). Stored so the page can show it next to what's confirmed.
    force=False (the automatic one) skips if a suggestion for the same files and facts was made meanwhile."""
    with _SUGGESTING, calllog.tag(engagement=eid, step="roles"):
        key = _roles_key(eid, workbooks(eid), documents(eid))
        prev = _q("SELECT roles_suggested FROM engagements WHERE id=?", eid)
        if not force and prev and (json.loads(prev[0]["roles_suggested"] or "null") or {}).get("key") == key:
            return get(eid)
        return _suggest_roles(eid, key)


def _suggest_roles(eid: int, key: str) -> dict:
    docs = [d for d in documents(eid) if d["status"] == "done"]
    n_facts = {d["id"]: sum(1 for f in facts(eid) if f["document_id"] == d["id"]) for d in docs}
    try:
        res = rolesmod.suggest([{**d, "n_facts": n_facts[d["id"]]} for d in docs], _wb_inputs(eid), reference(eid))
    except Exception as ex:  # kept with its key, so a failing suggestion isn't retried on every poll
        traceback.print_exc()
        prev = _q("SELECT roles_suggested FROM engagements WHERE id=?", eid)[0]["roles_suggested"]
        _set("engagements", eid, roles_suggested=json.dumps({**(json.loads(prev or "null") or {"roles": {}, "workbooks": {}}),
                                                             "key": key, "error": friendly(ex)}))
        return get(eid)
    res.update(key=key, at=time.time())
    e = _q("SELECT reviewer_model FROM engagements WHERE id=?", eid)[0]
    try:  # a second opinion from the reviewer model on the same evidence; the person decides
        docs_in = [{**d, "n_facts": n_facts[d["id"]]} for d in docs]
        res["second_opinion"] = rolesmod.second_opinion(docs_in, _wb_inputs(eid), reference(eid), res,
                                                        e["reviewer_model"] or DEFAULT_REVIEWER, _logger(eid))
    except Exception as ex:
        res["second_opinion"] = {"error": friendly(ex)}
    _set("engagements", eid, roles_suggested=json.dumps(res, default=str), updated_at=time.time())
    have = roles(eid)
    with _lock, _conn() as db:
        for role, r in res["roles"].items():
            if not have.get(role, {}).get("confirmed"):
                db.execute("INSERT OR REPLACE INTO roles VALUES (?,?,?,?,?,?,0)",
                           (eid, role, r["kind"], r["id"], json.dumps(r["sheets"]), json.dumps(r["why"])))
    return get(eid)


def confirm_roles(eid: int, assignments: dict) -> dict:
    """assignments: {role: {"kind", "id", "sheets"} or None to clear}. Everything given is confirmed."""
    have = roles(eid)
    with _lock, _conn() as db:
        for role, a in assignments.items():
            if role not in rolesmod.ROLES:
                raise ValueError(f"unknown role {role}")
            if not a:
                db.execute("DELETE FROM roles WHERE engagement_id=? AND role=?", (eid, role))
                continue
            why = have.get(role, {}).get("why") if have.get(role, {}).get("id") == a["id"] else ["set by you"]
            db.execute("INSERT OR REPLACE INTO roles VALUES (?,?,?,?,?,?,1)",
                       (eid, role, a["kind"], int(a["id"]), json.dumps(a.get("sheets")), json.dumps(why or [])))
        # the compare, the map and the Python overlay were built on the old roles
        db.execute("""UPDATE engagements SET updated_at=?, map_status=NULL, map_json=NULL, map_error=NULL,
                      compare_status=NULL, compare_json=NULL, compare_summary=NULL, compare_error=NULL,
                      overlay_status=NULL, overlay_json=NULL, overlay_error=NULL WHERE id=?""", (time.time(), eid))
    _SESSIONS.pop(eid, None)
    return get(eid)


def _role_wb(eid: int, role: str) -> dict | None:
    r = roles(eid).get(role)
    if not r or r["kind"] != "workbook":
        return None
    w = next((w for w in workbooks(eid) if w["id"] == r["id"]), None)
    if not w or w["status"] != "done":
        return None
    return {**w, "sheets": set(r["sheets"]) if r["sheets"] else None, "confirmed": r["confirmed"]}


# ---- compare and map ----------------------------------------------------------------------------------------

def start(kind: str, eid: int) -> dict:
    if kind == "compare" and not (_role_wb(eid, "prior_model") and _role_wb(eid, "current_model")):
        raise ValueError("assign the prior and current client models first")
    if kind in ("map", "overlay") and not _role_wb(eid, "prior_overlay"):
        raise ValueError("assign the prior overlay first")
    if kind == "rows":
        return start_rows(eid)
    _set("engagements", eid, **{f"{kind}_status": "queued", f"{kind}_step": "Waiting to start", f"{kind}_error": None})
    _jobs.put((kind, eid))
    return get(eid)


def _compare(eid: int) -> None:
    e = _q("SELECT * FROM engagements WHERE id=?", eid)[0]
    prior, cur = _role_wb(eid, "prior_model"), _role_wb(eid, "current_model")
    _set("engagements", eid, compare_status="running", compare_step="Comparing the client models cell by cell")
    d = diffmod.diff(prior["db_path"], cur["db_path"])
    ov = _role_wb(eid, "prior_overlay")
    left_out = []
    if prior["sheets"] and ov and ov["id"] == prior["id"]:
        # overlay inside the prior client model: its sheets aren't in the client's new model, and that's expected
        left_out = [s for s in d["sheets"]["removed"] if s not in prior["sheets"]]
        d["sheets"]["removed"] = [s for s in d["sheets"]["removed"] if s in prior["sheets"]]
    d["overlay_sheets_left_out"] = left_out
    d["previous"] = {"id": prior["id"], "filename": prior["filename"]}
    d["warnings"] = library._warnings(d, library.get(cur["id"]))
    _set("engagements", eid, compare_json=json.dumps(d, default=str), compare_step="Writing the summary")
    summary = library.summarize(d, library.get(prior["id"]), library.get(cur["id"]), model=e["model"] or DEFAULT_MODEL,
                                session=_session(eid))
    _set("engagements", eid, compare_summary=summary, compare_status="done", compare_step="Done", updated_at=time.time())


def _map(eid: int) -> None:
    import valuation
    ov, prior, cur = _role_wb(eid, "prior_overlay"), _role_wb(eid, "prior_model"), _role_wb(eid, "current_model")
    ref = reference(eid)
    step = lambda msg: _set("engagements", eid, map_status="running", map_step=msg)
    step("Finding the report's figures in the overlay")
    report_overlay = linkmap.match_facts(ov["db_path"], ref, ov["sheets"])
    step("Recomputing the overlay's DCFs in Python")
    try:
        cat = [a for a in valuation.catalogue(ov["db_path"]) if not ov["sheets"] or a["cell"].split("!")[0] in ov["sheets"]]
    except Exception:
        cat = []
        traceback.print_exc()
    anchors_at = {}
    for fm in report_overlay:
        for m in (x for x in fm["matches"] if x.get("located")):
            anchors_at.setdefault(f"{m['sheet']}!{m['addr']}", []).append(f"{fm['label'] or fm['key']} {fm['value_text']}")
    python = [{"cell": a["cell"], "label": a["label"], "value": a["value"], "python": a.get("total"),
               "reproduced": bool(a.get("matches")), "ok": a.get("ok"), "reason": a.get("reason"),
               "report": anchors_at.get(a["cell"], [])} for a in cat]
    step("Following the overlay's links into the prior client model")
    copy = _wiring(eid).get("client_sheets") if ov and prior else None
    if copy:  # the overlay reads its own copy of the client sheets; that copy's rows are the prior model's rows
        ov_client = linkmap.overlay_to_client({**ov, "sheets": set(ov["sheets"] or [])}, {**ov, "sheets": set(copy)})
        ov_client["mode"] = "inside a copy of the prior client model"
    else:
        ov_client = linkmap.overlay_to_client(ov, prior) if ov else {"links": [], "books": []}
    alignment = {}
    if prior and cur:
        step("Lining up those rows with the current client model")
        refs = [(ln["client_sheet"], ln["client_row"]) for ln in ov_client["links"] if ln.get("client_row")
                and (ov_client.get("client_link") is None or ln.get("link") == ov_client.get("client_link"))]
        alignment = linkmap.align_rows(prior["db_path"], cur["db_path"], refs)
    out = {"report_overlay": report_overlay, "python": python, "overlay_client": {**ov_client, "links": None},
           "chain": linkmap.chain(ov_client, alignment), "reference_approved": all(f["approved"] for f in ref) and bool(ref),
           "files": {k: (v["filename"] if v else None) for k, v in
                     (("overlay", ov), ("prior_model", prior), ("current_model", cur))}}
    _set("engagements", eid, map_json=json.dumps(out, default=str), map_status="done", map_step="Done",
         updated_at=time.time())


# ---- model dashboards (map step) ----------------------------------------------------------------------------

_ROW_REF = re.compile(r"^'?(.+?)'?!r(\d+)$")
_CELL_REF = re.compile(r"^'?(.+?)'?!\$?[A-Z]{1,3}\$?(\d+)$")


def _engagement_wb(eid: int, fid: int) -> dict:
    w = next((w for w in workbooks(eid) if w["id"] == fid), None)
    if not w:
        raise ValueError("that workbook isn't in this engagement")
    if w["status"] != "done" or not w["db_path"] or not Path(w["db_path"]).exists():
        raise ValueError(f"{w['filename']} hasn't finished processing")
    return w


def model_marks(eid: int, fid: int) -> dict[tuple[str, int], list[str]]:
    """What the map and the Python overlay say about this workbook's rows: report figures found there, overlay rows
    reading the client model, client rows the overlay reads, this year's matches, levers and outputs."""
    rl = roles(eid)
    has = lambda role: (rl.get(role) or {}).get("kind") == "workbook" and rl[role]["id"] == fid
    e = _q("SELECT map_json, overlay_json FROM engagements WHERE id=?", eid)[0]
    m = json.loads(e["map_json"] or "null") or {}
    o = json.loads(e["overlay_json"] or "null") or {}
    marks: dict[tuple[str, int], list[str]] = {}

    def mark(ref, pattern, text):
        hit = pattern.match(ref or "")
        if hit:
            tags = marks.setdefault((hit[1], int(hit[2])), [])
            if text not in tags:
                tags.append(text)

    chain = m.get("chain") or []
    if has("prior_overlay"):
        for x in m.get("report_overlay") or []:
            if x["matches"] and x["matches"][0].get("located", True):  # a value-only match isn't where the figure is
                b = x["matches"][0]
                mark(f"{b['sheet']}!r{b['row']}", _ROW_REF, f"report: {x.get('label') or x['key']} {x['value_text']}")
        for c in chain:
            mark(c["overlay"], _ROW_REF, f"reads client {c['client']}")
        for lv in o.get("levers") or []:
            mark(lv["cell"], _CELL_REF, f"lever: {lv['label']}")
        for x in o.get("outputs") or []:
            mark(x["cell"], _CELL_REF, f"output: {x['label']}")
    if has("prior_model"):
        for c in chain:
            mark(c["client"], _ROW_REF, f"read by overlay {c['overlay']}")
    if has("current_model"):
        for c in chain:
            if c.get("current"):
                mark(c["current"], _ROW_REF, f"this year's {c['client']} (read by overlay {c['overlay']})")
    return marks


def model_dashboard(eid: int, fid: int) -> dict:
    import modeldash
    w = _engagement_wb(eid, fid)
    rl = roles(eid)
    as_roles = [k for k, r in rl.items() if r["kind"] == "workbook" and r["id"] == fid]
    ov = rl.get("prior_overlay") or {}
    marks = model_marks(eid, fid)
    per_sheet: dict[str, int] = {}
    for s, _ in marks:
        per_sheet[s] = per_sheet.get(s, 0) + 1
    return {**modeldash.summary(w["db_path"]), "id": fid, "filename": w["filename"], "roles": as_roles,
            "target_name": w["target_name"], "valuation_date": w["valuation_date"],
            "overlay_sheets": (ov.get("sheets") or []) if ov.get("id") == fid else [],
            "marked": {"rows": len(marks), "by_sheet": per_sheet}}


def model_rows(eid: int, fid: int, sheet: str | None, q: str | None, mapped: bool, limit: int = 200) -> list[dict]:
    import modeldash
    w = _engagement_wb(eid, fid)
    marks = model_marks(eid, fid)
    out = modeldash.rows(w["db_path"], sheet or None, q or None, list(marks) if mapped else None, limit)
    for r in out:
        r["marks"] = marks.get((r["sheet"], r["row"]), [])
    return out


# ---- the Python overlay -------------------------------------------------------------------------------------

_SESSIONS: dict[int, tuple] = {}  # engagement -> (overlay.Session, summary): the compiled module, loaded and wired


def _wiring(eid: int) -> dict:
    """Which files and sheets the overlay module reads, from the confirmed (or suggested) roles."""
    ov, prior, cur = _role_wb(eid, "prior_overlay"), _role_wb(eid, "prior_model"), _role_wb(eid, "current_model")
    names = {w["id"]: w["sheet_names"] for w in workbooks(eid)}
    sheets = [s for s in names[ov["id"]] if not ov["sheets"] or s in ov["sheets"]]
    same_file = bool(prior) and prior["id"] == ov["id"]
    # the overlay in a copy of the client model, the client's own file assigned as the prior model: the overlay
    # reads its copy's sheets, which are fed from that file
    copy = [s for s in names[ov["id"]] if s not in sheets and s in set(names.get(prior["id"], []))] \
        if prior and not same_file and ov["sheets"] else []
    client_link = None
    if prior and not same_file and not copy:
        extlinks.ensure(ov["source_path"], ov["db_path"])
        client_link = linkmap.overlay_to_client({**ov, "sheets": None}, {**prior, "sheets": None}).get("client_link")
    vd = next((f for f in reference(eid) if f["key"] == "valuation_date" and f.get("value")), None)
    prior_vd = None
    if vd:
        v = int(vd["value"])
        prior_vd = f"{v // 10000:04d}-{v // 100 % 100:02d}-{v % 100:02d}"
    elif library.get(ov["id"]) and library.get(ov["id"])["valuation_date"]:
        prior_vd = library.get(ov["id"])["valuation_date"]
    return {"overlay": {"db_path": ov["db_path"], "filename": ov["filename"], "sheets": sheets,
                        "source_path": ov.get("source_path")},
            "prior": {"db_path": prior["db_path"], "filename": prior["filename"],
                      "sheets": sorted(prior["sheets"]) if prior["sheets"] else None,
                      "valuation_date": (library.get(prior["id"]) or {}).get("valuation_date")} if prior else None,
            "current": {"db_path": cur["db_path"], "filename": cur["filename"], "sheets": None,
                        "valuation_date": (library.get(cur["id"]) or {}).get("valuation_date")} if cur else None,
            "client_link": client_link, "prior_valuation_date": prior_vd, "same_file": same_file,
            "client_sheets": copy or None}


def _overlay(eid: int) -> None:
    import overlay as ovmod
    e = _q("SELECT name FROM engagements WHERE id=?", eid)[0]
    step = lambda frac, msg: _set("engagements", eid, overlay_status="running", overlay_step=msg,
                                  overlay_pct=round(frac, 3))
    step(0, "Reading the roles")
    w = _wiring(eid)
    _SESSIONS.pop(eid, None)
    summary, sess = ovmod.build(OUT / "overlays" / f"e{eid}", w["overlay"], w["prior"], w["current"], reference(eid),
                                e["name"], w["client_link"], w["prior_valuation_date"], step, client_sheets=w["client_sheets"])
    summary["wiring"] = w
    summary["reference_approved"] = all(f["approved"] for f in reference(eid)) and bool(reference(eid))
    ovmod.deep(_load_holds, eid, sess)
    _load_rowpicks(eid, sess)
    _sync_roll(eid, sess, summary)
    _SESSIONS[eid] = (sess, summary)
    _set("engagements", eid, overlay_json=json.dumps(summary, default=str), overlay_status="done", overlay_step="Done",
         updated_at=time.time())
    if summary["wiring"].get("current"):  # the row agents take it from here, without anyone starting them
        try:
            start_rows(eid)
        except ValueError:  # the build itself succeeded
            pass


def _this_year_file(eid: int) -> Path:
    return OUT / "overlays" / f"e{eid}" / "this_year.json"


def this_year_date(eid: int) -> str | None:
    """This year's valuation date as set for the engagement (a client model's own date is the model's, which a
    fresh model built for another date doesn't make this year's), or None."""
    f = _this_year_file(eid)
    try:
        return json.loads(f.read_text(encoding="utf-8")).get("valuation_date") if f.exists() else None
    except (OSError, ValueError):
        return None


def set_this_year_date(eid: int, valuation_date: str | None) -> dict:
    """Set (or clear, with None) this year's valuation date for the engagement: the roll-forward runs to it."""
    if valuation_date and not re.match(r"^\d{4}-\d{2}-\d{2}$", valuation_date):
        raise ValueError("give the date as YYYY-MM-DD")
    last = _dates(eid)["dates"]
    base = last.get("overlay") or last.get("prior_client")
    if valuation_date and base and valuation_date <= base[:10]:
        raise ValueError(f"this year's valuation date must be after last year's ({base[:10]})")
    f = _this_year_file(eid)
    f.parent.mkdir(parents=True, exist_ok=True)
    f.write_text(json.dumps({"valuation_date": valuation_date}), encoding="utf-8")
    if eid in _SESSIONS:
        sess, summary = _SESSIONS[eid]
        _sync_roll(eid, sess, summary)
    _touch(eid)
    try:  # the rows to find follow the roll: the agents look again
        start_rows(eid)
    except ValueError:
        pass
    return {"valuation_date": valuation_date}


def _dates(eid: int) -> dict:
    """The valuation dates the roll-forward uses, as the files and the report give them now: last year's (the
    report's, else the overlay file's), last year's and this year's client models', this year's as set for the
    engagement, and which are confirmed."""
    vd = next((f for f in reference(eid) if f["key"] == "valuation_date" and f.get("value")), None)
    out, confirmed = {}, {}
    for role, key in (("prior_overlay", "overlay"), ("prior_model", "prior_client"), ("current_model", "current_client")):
        w = _role_wb(eid, role)
        rec = library.get(w["id"]) if w else None
        out[key] = (rec or {}).get("valuation_date")
        confirmed[key] = bool((rec or {}).get("identity_confirmed"))
    if vd:
        v = int(vd["value"])
        out["overlay"] = f"{v // 10000:04d}-{v // 100 % 100:02d}-{v % 100:02d}"
        confirmed["overlay"] = bool(vd.get("approved"))
    out["this_year"] = this_year_date(eid)
    confirmed["this_year"] = bool(out["this_year"])
    return {"dates": out, "confirmed": confirmed}


def _sync_roll(eid: int, sess, summary: dict) -> None:
    """The roll-forward from the dates as they are now (a file's valuation date can be corrected after the build),
    and last year's valuation date set on the session."""
    import overlay as ovmod
    want = _profile(eid).get("horizon")  # set in the engagement's profile, else worked out
    if getattr(sess, "horizon_set", None) != want:
        sess.horizon_set = want
        sess._pshift.clear()
    roll = summary.get("roll")
    if not roll or not summary["wiring"].get("current"):
        return
    now = _dates(eid)
    d = now["dates"]
    if d != roll.get("dates") or roll.get("plan") != ovmod.ROLL_PLAN or roll.get("horizon_set") != want:
        w = summary["wiring"]
        fresh = ovmod.deep(ovmod.plan_roll, sess, w.get("prior"), w["overlay"], w.get("same_file"), d["overlay"],
                           d["prior_client"], d["current_client"], d.get("this_year"))
        roll.update(fresh)
    elif roll.get("prior_valuation_date"):
        sess.base_vd = ovmod.serial(ovmod.date.fromisoformat(roll["prior_valuation_date"][:10]))
    roll["confirmed"] = now["confirmed"]


def overlay_session(eid: int):
    """The live module for an engagement, loaded from its saved build after a restart."""
    import overlay as ovmod
    if eid in _SESSIONS:
        return _SESSIONS[eid]
    rows = _q("SELECT overlay_json FROM engagements WHERE id=?", eid)
    summary = json.loads(rows[0]["overlay_json"] or "null") if rows else None
    if not summary or not Path(summary["module"]).exists():
        raise ValueError("build the Python overlay first")
    w = summary["wiring"]
    sess = ovmod.Session(summary["module"], w["overlay"]["db_path"], w["overlay"]["sheets"],
                         None if w["same_file"] else (w["prior"] or {}).get("db_path"), (w["current"] or {}).get("db_path"),
                         w["client_link"], w.get("client_sheets") or ((w["prior"] or {}).get("sheets") if w["same_file"] else None))
    ovmod.deep(_load_holds, eid, sess)
    stale = _load_rowpicks(eid, sess)
    _sync_roll(eid, sess, summary)
    _SESSIONS[eid] = (sess, summary)
    if stale:  # the agents' picks from an older row finder were set aside: they look at those rows again
        try:
            start_rows(eid)
        except ValueError:
            pass
    return _SESSIONS[eid]


def _live(eid: int, mode: str, changes: dict | None) -> tuple:
    """The engagement's session and summary, after checking the feed exists; changes with dates as serials."""
    import overlay as ovmod
    sess, summary = overlay_session(eid)
    if mode not in ("workbook", "prior", "current"):
        raise ValueError("the feed must be workbook, prior or current")
    if mode == "current" and not summary["wiring"].get("current"):
        raise ValueError("assign the current client model (Roles) and rebuild in Python to roll forward")
    if mode == "prior" and not summary["wiring"].get("prior"):
        raise ValueError("assign the prior client model (Roles) to feed from it")
    clean = {}
    for cell, v in (changes or {}).items():
        if isinstance(v, str) and re.match(r"^\d{4}-\d{2}-\d{2}$", v):
            v = ovmod.serial(ovmod.date.fromisoformat(v))
        elif isinstance(v, str):
            v = float(v.replace(",", "").rstrip("%")) / (100 if v.strip().endswith("%") else 1)
        clean[cell] = v
    return sess, summary, clean


def overlay_run(eid: int, mode: str, changes: dict, valuation_date: str | None, months: int | None) -> dict:
    import overlay as ovmod
    sess, summary, clean = _live(eid, mode, changes)
    _sync_roll(eid, sess, summary)
    return ovmod.deep(ovmod.scenario, sess, summary, mode, clean, valuation_date, months)


def overlay_valuation(eid: int, cell: str | None) -> dict:
    """The Model Desk's Validate step on the overlay: a DCF found in the workbook, recomputed step by step from the
    values Excel saved."""
    import overlay as ovmod
    import valuation
    _, summary = overlay_session(eid)
    anchors = ovmod.dcf_anchors(summary)
    usable = [a for a in anchors if a.get("ok")]
    if not usable:
        return {"anchors": valuation._listing(anchors), "selected": None}
    pick = next((a for a in usable if a["cell"] == cell), usable[0])
    v = valuation.validation(summary["wiring"]["overlay"]["db_path"], pick["cell"], anchors=anchors)
    v["anchors"] = valuation._listing(anchors)
    return v


# ---- the Summary page: the report's summary table and charts, rebuilt and rolled forward ----------------------

def _charts_file(eid: int) -> Path | None:
    r = roles(eid).get("prior_report")
    d = _doc(r["id"]) if r and r.get("kind") == "document" else None
    return Path(d["out_dir"]) / "report_charts.json" if d else None


def recreate_charts(eid: int) -> dict:
    """Queue the report's charts: found, read, recreated from the prior client model, checked, redrawn on this
    year's model (reportcharts.py)."""
    r = roles(eid)
    if not r.get("prior_report") or not r.get("prior_model"):
        raise ValueError("assign last year's report and client model (Roles) first")
    _set("engagements", eid, charts_status="queued", charts_step="Waiting to start", charts_error=None)
    _jobs.put(("charts", eid))
    return {"queued": True}


def _fy_hint(eid: int):
    """The engagement's financial-year end for the charts (chartdata.fy_end_hint): the profile's if a person set it
    (it wins over what a model says), else last year's valuation date's month for a model that shows none."""
    import chartdata
    mine = _profile(eid).get("fy_end_month")
    if mine:
        return chartdata.fy_end_hint(mine, forced=True)
    vd = _prior_vd(eid)
    return chartdata.fy_end_hint(int(vd[5:7]) if vd else None)


def _prior_vd(eid: int) -> str | None:
    """Last year's valuation date (ISO): the report's fact, else the roll's from the overlay build."""
    vd = next((f for f in reference(eid) if f["key"] == "valuation_date" and f.get("value")), None)
    if vd:
        v = int(vd["value"])
        return f"{v // 10000:04d}-{v // 100 % 100:02d}-{v % 100:02d}"
    rows = _q("SELECT overlay_json FROM engagements WHERE id=?", eid)
    return ((json.loads((rows[0]["overlay_json"] if rows else None) or "null") or {}).get("roll") or {}).get("prior_valuation_date")


# ---- the engagement's profile -------------------------------------------------------------------------------
# The facts about an engagement's models that each step would otherwise guess on its own, in one place: detected,
# with how, and set by a person where the detection is wrong. The financial-year end and the horizon drive the
# charts and the roll-forward; the period frequency, the units and the discounting convention are shown (the
# convention is read back from each DCF's own factors, and the Summary's method selector changes it for a scenario).

MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October",
          "November", "December"]


def _profile(eid: int) -> dict:
    """A person's settings: {"fy_end_month": 6, "horizon": "fixed"}."""
    rows = _q("SELECT profile_json FROM engagements WHERE id=?", eid)
    return json.loads((rows[0]["profile_json"] if rows else None) or "null") or {}


def set_profile(eid: int, fields: dict) -> dict:
    """Set (or, with None, clear) the financial-year end month (1-12) or the horizon ("fixed" / "rolling")."""
    mine = _profile(eid)
    for k, v in fields.items():
        if k == "fy_end_month" and v not in (None, ""):
            if not (isinstance(v, int) and 1 <= v <= 12):
                raise ValueError("the financial year ends in a month, 1 to 12")
        elif k == "horizon" and v not in (None, ""):
            if v not in ("fixed", "rolling"):
                raise ValueError('the horizon is "fixed" or "rolling"')
        elif k not in ("fy_end_month", "horizon"):
            raise ValueError(f"{k} isn't set here: it's read from the models")
        if v in (None, ""):
            mine.pop(k, None)
        else:
            mine[k] = v
    _set("engagements", eid, profile_json=json.dumps(mine), updated_at=time.time())
    return profile_view(eid)


class _Timelines:
    """A model's period dates by sheet, read from each sheet's timeline row alone (overlay.Workbook reads whole
    sheets: on a large model, the profile would read most of the workbook)."""

    def __init__(self, path: str):
        self.db = rodb.connect(path)

    def timeline(self, s: str) -> dict[int, float]:
        from xlruntime import from_db
        lay = self.db.execute("SELECT layout FROM sheets WHERE sheet=?", (s,)).fetchone()
        hr = json.loads((lay[0] if lay else None) or "{}").get("header_row")
        out = {}
        for c, v in self.db.execute("SELECT col, value FROM cells WHERE sheet=? AND row=?", (s, hr)) if hr else []:
            x = from_db(v)
            if isinstance(x, float) and 3000 < x < 120000:
                out[c] = x
        return out

    def close(self):
        self.db.close()


def _frequency(wb) -> tuple[str | None, str]:
    """The periods' length most sheets with a timeline have: monthly, quarterly, half-yearly or annual."""
    kinds = Counter()
    for (s,) in wb.db.execute("SELECT sheet FROM sheets"):
        ends = sorted(set(wb.timeline(s).values()))
        if len(ends) < 3:
            continue
        gap = sorted(b - a for a, b in zip(ends, ends[1:]))[(len(ends) - 1) // 2]
        kinds["monthly" if gap <= 35 else "quarterly" if gap <= 100 else "half-yearly" if gap <= 200
              else "annual" if gap <= 400 else "longer than a year"] += 1
    if not kinds:
        return None, "no sheet with a timeline"
    k, n = kinds.most_common(1)[0]
    return k, f"{n} of {sum(kinds.values())} sheet(s) with a timeline" + (
        f" (also {', '.join(f'{x}: {m}' for x, m in kinds.items() if x != k)})" if len(kinds) > 1 else "")


def profile_view(eid: int) -> dict:
    """The profile: each fact detected (value, how), a person's setting where there is one (set), what it drives.
    Reads the models' own files; no model calls."""
    import chartdata
    import overlay as ovmod
    mine = _profile(eid)
    prior, cur, ovw = _role_wb(eid, "prior_model"), _role_wb(eid, "current_model"), _role_wb(eid, "prior_overlay")
    out = {}
    # the financial year: what the model shows, else last year's valuation date's month
    src = cur or prior
    month, how = None, "no client model assigned yet"
    if src:
        with closing(rodb.connect(src["db_path"])) as db:  # closed: Windows won't rebuild an open file
            month, how = chartdata.fy_end_detect(db)
        how = f"{how}, in {src['filename']}" if month else f"{src['filename']} shows none: {how}"
    vd = _prior_vd(eid)
    if not month and vd:
        month, how = int(vd[5:7]), f"last year's valuation date ({vd}); {how}"
    out["fy_end_month"] = {"label": "Financial year ends", "value": mine.get("fy_end_month") or month, "detected": month,
                           "shown": MONTHS[(mine.get("fy_end_month") or month) - 1] if (mine.get("fy_end_month") or month) else None,
                           "how": how if month else f"{how}; December is assumed", "set": "fy_end_month" in mine,
                           "drives": "the report's charts: model rows are totalled by financial year to match them"}
    # the horizon: this year's client model ends its periods where last year's did, or later
    kind, n = None, {"fixed": 0, "rolling": 0}
    live = _SESSIONS.get(eid)
    if live and live[0].current and live[0].rowmap:
        sess = live[0]
        kind, n = ovmod.horizon(sess.prior or sess.ov, sess.current, sess.client_sheets or {k[1] for k in sess.ext_cached},
                                sess.rowmap.sheet_for)
    elif (prior or ovw) and cur:
        base = prior or ovw
        a, b = _Timelines(base["db_path"]), _Timelines(cur["db_path"])
        try:
            sheets = base["sheets"] or [s for (s,) in a.db.execute("SELECT sheet FROM sheets")]
            kind, n = ovmod.horizon(a, b, sheets)
        finally:
            a.close()
            b.close()
    how = (f"this year's client model ends its periods where last year's did on {n['fixed']} sheet(s), later on "
           f"{n['rolling']}" if kind else "needs last year's and this year's client models, with timelines")
    out["horizon"] = {"label": "Horizon", "value": mine.get("horizon") or kind, "detected": kind, "how": how,
                      "set": "horizon" in mine,
                      "shown": {"fixed": "Fixed: the periods end on the same date each year",
                                "rolling": "Rolling: the periods move on each year"}.get(mine.get("horizon") or kind),
                      "drives": "the roll-forward: on a fixed horizon the periods stay and only the valuation date moves"}
    # shown, not set: the period frequency, the units, the discounting convention
    freq, how = (None, "no client model assigned yet")
    if src:
        w = _Timelines(src["db_path"])
        try:
            freq, how = _frequency(w)
        finally:
            w.close()
    out["frequency"] = {"label": "Periods", "value": freq, "shown": freq, "how": how, "set": False,
                        "drives": "shown: each sheet's own periods are used"}
    units = next((f for f in reference(eid) if f["key"] == "currency_units"), None)
    out["units"] = {"label": "Currency and units", "value": units and units.get("value_text"),
                    "shown": units and units.get("value_text"), "set": False,
                    "how": f"the report's facts (page {units.get('page')})" if units else "not among the report's facts",
                    "drives": "shown: each figure keeps its own units"}
    conv, how = _conventions(eid)
    out["discounting"] = {"label": "Discounting", "value": conv, "shown": conv, "how": how, "set": False,
                          "drives": "shown: read back from each DCF's factors; the Summary's method selector changes "
                                    "it for a scenario"}
    return {"fields": out, "settable": ["fy_end_month", "horizon"]}


# ---- the output schedule: the overlay's own outputs, whether or not the report quotes them (outputs.py) ----------

_SCHEDULES: dict[int, tuple] = {}  # engagement -> (what it was worked out from, the detected schedule)


def _schedule(eid: int) -> dict:
    """A person's part: {"classes": {"Sheet!r12": "working"}, "outside": [fact key], "confirmed_at", "changed_at"}."""
    rows = _q("SELECT schedule_json FROM engagements WHERE id=?", eid)
    return json.loads((rows[0]["schedule_json"] if rows else None) or "null") or {}


def schedule_view(eid: int) -> dict:
    """The overlay's outputs with a person's classes over the detected ones, the report's figures that sit on none
    of them (each marked, or not, as produced outside the model), and when the schedule was confirmed."""
    import outputs as outmod
    import valuation
    rows = _q("SELECT overlay_json FROM engagements WHERE id=?", eid)
    summary = json.loads((rows[0]["overlay_json"] if rows else None) or "null")
    if not summary:
        raise ValueError("build the Python overlay first: the schedule is read from the overlay it compiles")
    path = summary["wiring"]["overlay"]["db_path"]
    key = (path, Path(path).stat().st_mtime if Path(path).exists() else None, json.dumps(
        [summary.get("levers"), [(o["cell"], o.get("key")) for o in summary.get("outputs") or []]], default=str))
    if _SCHEDULES.get(eid, (None,))[0] != key:
        anchors = {a["cell"] for a in valuation.catalogue(path) if a.get("ok")}
        with closing(rodb.connect(path)) as db:
            _SCHEDULES[eid] = (key, outmod.detect(db, summary.get("sheets") or [], summary.get("levers"),
                                                  summary.get("outputs"), anchors))
    mine = _schedule(eid)
    if not any(mine.get(k) for k in ("classes", "outside", "confirmed_at", "carried")):
        src = _schedule_source(eid)  # a new schedule, and last year's engagement confirmed one: carried, once
        if src:
            try:
                mine = _carry(eid, src, summary)
            except Exception:  # a carry that fails leaves the detected schedule, and the button to try again
                traceback.print_exc()
    sched = outmod.apply(_SCHEDULES[eid][1], mine)
    marked = set(mine.get("outside") or [])
    out = outmod.outside(reference(eid), sched, summary.get("levers"))
    return {"outputs": sched, "outside": [{**f, "marked": f["key"] in marked} for f in out],
            "counts": dict(Counter(o["class"] for o in sched)), "confirmed_at": mine.get("confirmed_at"),
            "changed_at": mine.get("changed_at"), "carried": {k: v for k, v in (mine.get("carried") or {}).items() if k != "classes"} or None,
            "sources": _schedule_sources(eid)}


# ---- carrying a confirmed schedule into next year's engagement ---------------------------------------------------

def _chain(fid: int, steps: int = 5) -> list[int]:
    """A workbook and its earlier versions (library previous_id), nearest first."""
    out = [fid]
    for _ in range(steps):
        prev = (library.get(out[-1]) or {}).get("previous_id")
        if not prev or prev in out:
            break
        out.append(prev)
    return out


def _schedule_source(eid: int) -> int | None:
    """Last year's engagement: the nearest whose prior overlay is this one's overlay or an earlier version of it
    (the library links versions of one workbook), with a confirmed schedule."""
    ov = _role_wb(eid, "prior_overlay")
    for fid in _chain(ov["id"]) if ov else []:
        for r in _q("SELECT engagement_id FROM roles WHERE role='prior_overlay' AND kind='workbook' AND ref_id=? "
                    "AND engagement_id != ?", fid, eid):
            if _schedule(r["engagement_id"]).get("confirmed_at"):
                return r["engagement_id"]
    return None


def _schedule_sources(eid: int) -> dict:
    """For carrying by hand: the linked engagement, and every other one with a confirmed schedule."""
    linked = _schedule_source(eid)
    others = [{"id": r["id"], "name": r["name"]} for r in _q("SELECT id, name, schedule_json FROM engagements WHERE id != ?", eid)
              if (json.loads(r["schedule_json"] or "null") or {}).get("confirmed_at")]
    return {"linked": linked, "engagements": others}


def _carry(eid: int, src: int, summary: dict) -> dict:
    """Carry engagement src's confirmed schedule into this one (outputs.carry): its classes, as carried (a person's
    own classes here still win), and its outside-the-model marks for figures still on no output or input this year.
    Returns the schedule as it now stands; nothing is confirmed: the person looks at what's new, then confirms."""
    import outputs as outmod
    if src == eid:
        raise ValueError("an engagement can't carry its own schedule")
    theirs = _schedule(src)
    if not theirs.get("confirmed_at"):
        raise ValueError("that engagement hasn't confirmed its schedule")
    ov_src = _role_wb(src, "prior_overlay")
    if not ov_src or not ov_src.get("db_path"):
        raise ValueError("that engagement's overlay isn't available")
    got = outmod.carry(schedule_view(src)["outputs"], ov_src["db_path"], summary["wiring"]["overlay"]["db_path"],
                       summary.get("sheets") or [])
    mine = _schedule(eid)
    here = {f["key"] for f in outmod.outside(reference(eid), outmod.apply(_SCHEDULES[eid][1], mine), summary.get("levers"))}
    kept = [k for k in theirs.get("outside") or [] if k in here]
    name = (_q("SELECT name FROM engagements WHERE id=?", src) or [{}])[0].get("name")
    mine["carried"] = {"from": src, "name": name, "at": time.time(), "classes": got["classes"], "missing": got["missing"],
                       "outside": kept, "outside_dropped": [k for k in theirs.get("outside") or [] if k not in here]}
    mine["outside"] = sorted(set(mine.get("outside") or []) | set(kept))
    _set("engagements", eid, schedule_json=json.dumps(mine), updated_at=time.time())
    return mine


def carry_schedule(eid: int, src: int) -> dict:
    """Carry a confirmed schedule from another engagement by hand (again, or from one not linked)."""
    schedule_view(eid)  # this year's detected schedule, worked out
    rows = _q("SELECT overlay_json FROM engagements WHERE id=?", eid)
    _carry(eid, int(src), json.loads(rows[0]["overlay_json"]))
    return schedule_view(eid)


def _schedule_rows(eid: int) -> list[dict]:
    """The schedule's conclusions and assumptions that are single figures, for the Summary (none before a build)."""
    try:
        return [o for o in schedule_view(eid)["outputs"] if o["kind"] == "figure" and o["class"] in ("conclusion", "assumption")]
    except (ValueError, OSError):
        return []


def set_schedule(eid: int, classes: dict | None = None, outside: dict | None = None, confirm: bool = False) -> dict:
    """A person's classes ({"Sheet!r12": "working" | None to go back to the detected one}), report figures marked
    as produced outside the model ({fact key: True | False}), and confirming the schedule as it stands."""
    import outputs as outmod
    mine = _schedule(eid)
    cl = dict(mine.get("classes") or {})
    for row, c in (classes or {}).items():
        if c in (None, ""):
            cl.pop(row, None)
        elif c not in outmod.CLASSES:
            raise ValueError(f"a class is one of {', '.join(outmod.CLASSES)}")
        else:
            cl[row] = c
    out = set(mine.get("outside") or [])
    for k, v in (outside or {}).items():
        (out.add if v else out.discard)(k)
    now = time.time()
    mine.update(classes=cl, outside=sorted(out))
    if classes or outside:
        mine["changed_at"] = now
    if confirm:
        mine["confirmed_at"] = now
    _set("engagements", eid, schedule_json=json.dumps(mine), updated_at=now)
    return schedule_view(eid)


def _conventions(eid: int) -> tuple[str | None, str]:
    """How the DCFs under the overlay's figures discount (timing, day count), read back from their factors."""
    import dcftrace
    rows = _q("SELECT overlay_json FROM engagements WHERE id=?", eid)
    summary = json.loads((rows[0]["overlay_json"] if rows else None) or "null") or {}
    if not summary.get("outputs"):
        return None, "build the Python overlay: it's read from the DCFs under the figures"
    seen = Counter()
    with closing(rodb.connect(summary["wiring"]["overlay"]["db_path"])) as db:
        for o in summary["outputs"][:8]:
            try:
                for c in dcftrace.cores(dcftrace.trace(db, o["cell"])):
                    m = c.get("method")
                    if m:
                        seen[f"{m['timing']} of period, {m['day_count']}"] += 1
            except ValueError:
                continue
    if not seen:
        return None, "no DCF under the figures whose factors could be read back"
    k, n = seen.most_common(1)[0]
    return k, f"read back from the factors of {sum(seen.values())} discounting(s) under the figures" + (
        f" (others: {', '.join(x for x in seen if x != k)})" if len(seen) > 1 else "")


def _charts_job(eid: int) -> None:
    with _fy_hint(eid):
        _charts_run(eid)


def _charts_run(eid: int) -> None:
    import reportcharts
    rl = roles(eid)
    d = _doc(rl["prior_report"]["id"])
    doc = json.loads(d["doc_json"])
    e = _q("SELECT model, reviewer_model FROM engagements WHERE id=?", eid)[0]
    books = []
    for role, name in (("prior_model", "last year's client model"), ("prior_overlay", "the overlay")):
        w = _role_wb(eid, role)
        if w and w.get("db_path") and not any(b["db_path"] == w["db_path"] for b in books):
            books.append({"key": role, "name": name, "db_path": w["db_path"]})
    cur = _role_wb(eid, "current_model")
    _set("engagements", eid, charts_status="running", charts_step="Finding the charts in the report")
    reader = docingest.Reader(e["model"] or DEFAULT_MODEL, e["reviewer_model"] or DEFAULT_REVIEWER, _logger(eid))
    res = reportcharts.run(reader, doc, d["source_path"], d["out_dir"], books, (cur or {}).get("db_path"),
                           lambda f, m: _set("engagements", eid, charts_step=m))
    res["document"] = d["id"]
    _charts_file(eid).write_text(json.dumps(res, default=str), encoding="utf-8")
    _set("engagements", eid, charts_status="done", charts_step="Done")


def summary_view(eid: int, changes: dict | None = None, valuation_date: str | None = None, months: int | None = None,
                 method: dict | None = None) -> dict:
    """The Summary page: the report's summary table rebuilt on last year's model, rolled forward onto this
    year's, and as a scenario (overlay.summary_table), with the report's charts recreated (reportcharts.py). The
    charts are recreated by themselves the first time the page asks, once the overlay is built."""
    import overlay as ovmod
    rows = _q("SELECT overlay_status, charts_status, charts_step, charts_error FROM engagements WHERE id=?", eid)
    if not rows:
        raise ValueError("no such engagement")
    e = rows[0]
    out = {"charts": None, "charts_status": e.get("charts_status"), "charts_step": e.get("charts_step"),
           "charts_error": e.get("charts_error")}
    f = _charts_file(eid)
    if f and f.exists():
        out["charts"] = json.loads(f.read_text(encoding="utf-8"))
        import reportcharts
        for ch in out["charts"].get("charts") or []:  # files saved before the compare view: work it out now
            if "compare" not in ch and ch.get("read", {}).get("series") and reportcharts.time_axis(ch["read"].get("x_labels")):
                spec = ch.get("spec") or {}
                ch["compare"] = reportcharts.comparison(ch["read"], spec if spec.get("frequency") == "annual" else None)
    elif f and e.get("overlay_status") == "done" and not e.get("charts_status"):
        recreate_charts(eid)
        out.update(charts_status="queued", charts_step="Waiting to start")
    if e.get("overlay_status") != "done":
        out["table"] = None
        out["why"] = "rebuild in Python first: the summary is recomputed from the Python overlay"
        return out
    sess, summary = overlay_session(eid)
    _sync_roll(eid, sess, summary)
    clean = _live(eid, "current" if summary["wiring"].get("current") else "workbook", changes)[2]
    out["table"] = ovmod.deep(ovmod.summary_table, sess, summary, reference(eid), clean, valuation_date, months, method,
                              _schedule_rows(eid), set(_schedule(eid).get("outside") or []))
    out["identity"] = {f["key"]: f.get("value_text") for f in reference(eid) if f.get("category") == "identity"}
    out["roll_plan"] = summary.get("roll")  # as the dates are now (_sync_roll), with which are checked
    out["rows"] = rows_view(eid)  # the row agents: running, or what they decided
    return out


def bridge_view(eid: int, changes: dict | None = None, valuation_date: str | None = None, months: int | None = None,
                method: dict | None = None) -> dict:
    """The value bridge from last year's value to this year's (overlay.value_bridge), with the Summary page's
    scenario as its last step."""
    import overlay as ovmod
    rows = _q("SELECT overlay_status FROM engagements WHERE id=?", eid)
    if not rows:
        raise ValueError("no such engagement")
    if rows[0]["overlay_status"] != "done":
        return {"bridges": [], "why": "build the Python overlay first: the bridge is recomputed from it"}
    sess, summary = overlay_session(eid)
    _sync_roll(eid, sess, summary)
    clean = _live(eid, "current" if summary["wiring"].get("current") else "workbook", changes)[2]
    return ovmod.deep(ovmod.value_bridge, sess, summary, reference(eid), clean, valuation_date, months, method,
                      _schedule_rows(eid))


def chart_pick(eid: int, cid: str, series: int, rows: list[dict] | None) -> dict:
    """A person's rows for one series of one of the report's charts (rows: [{"book", "sheet", "row"}], all on one
    sheet; empty to clear it): the chart redrawn from them, checked by numbers against the reading, and redrawn on
    this year's model."""
    with _fy_hint(eid):
        return _chart_pick(eid, cid, series, rows)


def _chart_pick(eid: int, cid: str, series: int, rows: list[dict] | None) -> dict:
    import reportcharts
    f = _charts_file(eid)
    if not f or not f.exists():
        raise ValueError("recreate the report's charts first")
    data = json.loads(f.read_text(encoding="utf-8"))
    ch = next((c for c in data.get("charts") or [] if c["id"] == cid), None)
    if not ch or not ch.get("read"):
        raise ValueError("no such chart")
    read = ch["read"]
    years = reportcharts.time_axis(read.get("x_labels"))
    if not years or not 0 <= series < len(read["series"]):
        raise ValueError("that chart isn't one over years, or has no such series")
    # picks saved before the recreation was drawn from year totals have none: they are picked again
    picks = [p if p and "years" in p else None for p in (ch.get("picks") or [None] * len(read["series"]))]
    if rows:
        if len({(r["book"], r["sheet"]) for r in rows}) > 1:
            raise ValueError("pick rows from one sheet of one workbook")
        w = _role_wb(eid, rows[0]["book"])
        if not w:
            raise ValueError(f"no {rows[0]['book'].replace('_', ' ')} assigned")
        p = reportcharts.pick_rows(w["db_path"], read, series, rows[0]["sheet"], [int(r["row"]) for r in rows])
        picks[series] = {**p, "book": rows[0]["book"]}
    else:
        picks[series] = None
    title = ch.get("caption") or read.get("title") or ""
    spec = reportcharts.spec_for(read, picks, years, title)
    verdict = reportcharts.numbers_verdict(read, picks, ch.get("series_notes"))
    ch.update(spec=spec, picks=picks, matches=verdict["matches"], compare=reportcharts.comparison(read, spec))
    ch.setdefault("tries", []).append({"picks": picks, "verdict": {**verdict, "by": "numbers, with your rows"}, "png": None})
    ch.pop("problem", None)
    prior, cur = _role_wb(eid, "prior_model"), _role_wb(eid, "current_model")
    ch.pop("current_spec", None)
    if prior and cur:
        try:
            ch["current_spec"] = reportcharts.current_spec(
                prior["db_path"], cur["db_path"], read,
                [p if p and p.get("book") == "prior_model" else None for p in picks], title + " (this year)")
            ch.pop("current_problem", None)
        except Exception as e:
            ch["current_problem"] = f"{type(e).__name__}: {e}"
    f.write_text(json.dumps(data, default=str), encoding="utf-8")
    return ch


def chart_png(eid: int, name: str) -> Path | None:
    """A chart image of the Summary page (the report's crop or a recreation), inside the report's folder."""
    f = _charts_file(eid)
    if not f or not re.fullmatch(r"[\w.-]+\.png", name):
        return None
    p = f.parent / "charts" / name
    if not p.is_file():
        p = next((x for x in (f.parent / "tables").glob(name)), None) if (f.parent / "tables").is_dir() else None
    return p if p and p.is_file() else None


def overlay_value_trace(eid: int, start: str | None = None) -> dict:
    """How a report figure is built in the overlay, from the cell it was matched to down to the discounting
    (dcftrace.py), with the Python overlay's value for every cell on the way."""
    import dcftrace
    import overlay as ovmod
    import rodb
    sess, summary = overlay_session(eid)
    starts = ovmod.trace_starts(summary)
    if start and start not in {x["cell"] for x in starts}:
        starts.append({"cell": start, "label": None, "report": None, "key": None, "ties": False, "value": None})
    if not starts:
        return {"starts": [], "selected": None}
    pick = next((x for x in starts if x["cell"] == start), starts[0])
    db = rodb.connect(summary["wiring"]["overlay"]["db_path"])
    tree = dcftrace.trace(db, pick["cell"])
    nodes, sheets = [], set(summary["sheets"])

    def walk(n):
        if not n.get("again"):
            nodes.append(n)
        for c in n.get("children", []):
            walk(c)
    walk(tree)
    mine = [n for n in nodes if n["cell"].rsplit("!", 1)[0].strip("'") in sheets]
    keys = [ovmod.parse_a1(n["cell"]) for n in mine]

    def run():
        sess.configure("workbook")
        return sess.values(keys)
    for n, v in zip(mine, ovmod.deep(run)):  # numbers only: a date cell is text in model.db and a serial in Python
        n["python"] = v if isinstance(v, float) and isinstance(n.get("value"), (int, float)) else None
    return {"starts": starts, "selected": pick["cell"], "start": pick, "tree": tree, "cores": dcftrace.cores(tree),
            "text": dcftrace.text(tree)}


def overlay_dcf(eid: int, mode: str = "workbook", changes: dict | None = None, valuation_date: str | None = None,
                months: int | None = None, **method) -> dict:
    """A DCF on the live module (feed + changes) and under another discounting method: overlay.dcf_live()."""
    import overlay as ovmod
    sess, summary, clean = _live(eid, mode, changes)
    return ovmod.deep(ovmod.dcf_live, sess, summary, mode, clean, valuation_date, months, **method)


def overlay_ask(eid: int, question: str, model: str | None, history: list | None, session: str | None = None):
    """The chat on the Python overlay step (overlay_chat.py): a generator of UI events."""
    import overlay_chat
    e = _q("SELECT model FROM engagements WHERE id=?", eid)
    if not e:
        raise ValueError("no such engagement")
    log = lambda m, u: usage.record(m, u, "chat", None, _session(eid))
    yield from overlay_chat.ask(eid, question, model or e[0]["model"] or DEFAULT_MODEL, history, _session(eid), log)


def overlay_module(eid: int) -> Path:
    _, summary = overlay_session(eid)
    return Path(summary["module"])


def overlay_inputs(eid: int, q: str) -> list[dict]:
    import overlay as ovmod
    sess, _ = overlay_session(eid)
    return ovmod.inputs_list(sess, q)


def overlay_trace(eid: int, cell: str) -> dict:
    import overlay as ovmod
    sess, _ = overlay_session(eid)
    return ovmod.deep(ovmod.trace, sess, cell)


# ---- worker -------------------------------------------------------------------------------------------------

# ---- the overlay doctor (doctor.py) ------------------------------------------------------------------------

def _doctor_file(eid: int) -> Path:
    return OUT / "overlays" / f"e{eid}" / "doctor.json"


def _holds_file(eid: int) -> Path:
    return OUT / "overlays" / f"e{eid}" / "holds.json"


def _load_holds(eid: int, sess) -> None:
    """The cells a person chose to hold at Excel's saved value, while each is still the same cell: the same
    formula and the same saved value as when it was held (a rebuilt workbook that changed it drops the hold)."""
    import overlay as ovmod
    f = _holds_file(eid)
    held = json.loads(f.read_text(encoding="utf-8")) if f.exists() else {}
    keep = {}
    with rodb.connect(sess.ov.path) as db:
        for cell, h in held.items():
            k = ovmod.parse_a1(cell)
            row = db.execute("SELECT formula FROM cells WHERE sheet=? AND row=? AND col=?", k).fetchone()
            if row and row[0] == h.get("formula") and ovmod.same(sess.ov.value(*k), h.get("value")):
                keep[k] = h["value"]
    sess.holds = keep
    sess.configure("workbook")


def _rowpicks_file(eid: int) -> Path:
    return OUT / "overlays" / f"e{eid}" / "rowpicks.json"


def _load_rowpicks(eid: int, sess) -> int:
    """A person's choices of this year's row for last year's rows ({"CF!r11": "CF!r15"}), and the agents', for the
    roll-forward. The agents' picks made by another version of the row finder are left out (it may now find those
    rows itself, and better): returns how many, so the agents can look again."""
    if not sess.rowmap:
        return 0
    import rowfind
    stale = 0
    for a, b in _read_rowpicks(eid).items():
        try:  # "[1]Sheet!r9" (a row of the linked client model) is Sheet row 9, as when it was picked
            s, r = _row_ref(a)
            to, by = (b.get("to"), b.get("by", "you")) if isinstance(b, dict) else (b, "you")
            if by == "agent" and b.get("v") != rowfind.VERSION:
                stale += 1
                continue
            sess.rowmap.pick(s, r, rowfind.STAND_IN if to == "-" else _row_ref(to) if to else None, by)
        except ValueError:
            continue
    return stale


def _read_rowpicks(eid: int) -> dict:
    """{"Sheet!r9": "Sheet!r12" | "-" | {"to", "by": "you" | "agent", "why", "checked_by"}}: a person's picks were
    saved as plain strings before the agents made picks too."""
    f = _rowpicks_file(eid)
    try:
        return json.loads(f.read_text(encoding="utf-8")) if f.exists() else {}
    except (OSError, ValueError):
        return {}


def _write_rowpicks(eid: int, picks: dict) -> None:
    f = _rowpicks_file(eid)
    f.parent.mkdir(parents=True, exist_ok=True)
    f.write_text(json.dumps(picks, indent=1), encoding="utf-8")


def _row_ref(text: str) -> tuple[str, int]:
    m = re.match(r"^(?:\[\d+\])?(.+)!r(\d+)$", text or "")  # [1]Sheet!r9: a row of the linked client model
    if not m:
        raise ValueError(f"not a row: {text!r} (Sheet!rN)")
    return m[1], int(m[2])


def row_found(sess, s: str, r: int) -> dict:
    """What this year's model has for one of last year's rows: the row found, how, the evidence, the
    alternatives, and a few periods of both years' values side by side."""
    ex = sess.rowmap.explain(s, r)
    prior = sess.prior or sess.ov
    out = {"row": f"{s}!r{r}", "label": prior.labels().get((s, r), ""), "found": None, "how": ex["how"],
           "evidence": [f"{n}: {t}" for n, t in ex["evidence"]], "confidence": ex["confidence"],
           "alternatives": ex["alternatives"], "picked": (s, r) in sess.rowmap.picks,
           "stand_in": bool(ex.get("stand_in")), "confident": sess.rowmap.confident(s, r)}
    if ex["found"]:
        s2, r2 = ex["found"]
        out.update(found=f"{s2}!r{r2}", found_label=sess.current.labels().get((s2, r2), ""))
        tl_p, tl_c = prior.timeline(s), sess.current.timeline(s2)
        both = sorted(set(tl_p.values()) & set(tl_c.values()))[:6]
        inv_p, inv_c = {v: c for c, v in tl_p.items()}, {v: c for c, v in tl_c.items()}
        from xlruntime import to_date
        out["side_by_side"] = [{"period": to_date(w).isoformat(), "last_year": prior.value(s, r, inv_p[w]),
                                "this_year": sess.current.value(s2, r2, inv_c[w])} for w in both]
    return out


def overlay_facts(eid: int, start: str | None = None) -> dict:
    """The facts behind a report figure (dcffacts.py), and where this year's model has the client rows its cash
    flows come from (rowfind.py)."""
    import dcffacts
    import overlay as ovmod
    sess, summary = overlay_session(eid)
    starts = ovmod.trace_starts(summary)
    if start and start not in {x["cell"] for x in starts}:
        starts.append({"cell": start, "label": None, "report": None, "key": None, "ties": False, "value": None})
    if not starts:
        return {"starts": [], "selected": None}
    pick = next((x for x in starts if x["cell"] == start), starts[0])
    fx = ovmod.deep(dcffacts.facts, sess, summary, pick["cell"])
    if sess.rowmap:
        def found():
            seen = {}
            for c in fx["discountings"]:
                for o in c["origins"]:
                    if o["row"] not in seen:
                        seen[o["row"]] = row_found(sess, *_row_ref(o["row"]))
                    o["this_year"] = seen.get(o["row"])
        ovmod.deep(found)
    return {"starts": starts, "selected": pick["cell"], "start": pick, "facts": fx, "text": dcffacts.text(fx),
            "current": bool(sess.rowmap), "current_file": (summary["wiring"].get("current") or {}).get("filename")}


def row_pick(eid: int, prior_row: str, current_row: str | None) -> dict:
    """A person's choice of this year's row for one of last year's (None to go back to what was found)."""
    import overlay as ovmod
    sess, summary = overlay_session(eid)
    if not sess.rowmap:
        raise ValueError("assign this year's client model (Roles) and rebuild in Python first")
    import rowfind
    s, r = _row_ref(prior_row)
    keep = current_row == "-"  # keep last year's values for the row, on purpose
    to = rowfind.STAND_IN if keep else _row_ref(current_row) if current_row else None
    if to and not keep and not sess.current.db.execute("SELECT 1 FROM rows WHERE sheet=? AND row=?", to).fetchone():
        raise ValueError(f"{current_row} isn't a line item in this year's model")
    picks = _read_rowpicks(eid)
    if to:
        picks[prior_row] = {"to": current_row, "by": "you"}
    else:
        picks.pop(prior_row, None)
    _write_rowpicks(eid, picks)
    ovmod.deep(sess.rowmap.pick, s, r, to)
    return ovmod.deep(row_found, sess, s, r)


def row_info(eid: int, prior_row: str) -> dict:
    """What this year's model has for one of last year's rows (Sheet!rN), with the alternatives: for picking it."""
    import overlay as ovmod
    sess, _ = overlay_session(eid)
    if not sess.rowmap:
        raise ValueError("assign this year's client model (Roles) and rebuild in Python first")
    return ovmod.deep(row_found, sess, *_row_ref(prior_row))


def _rowagent_file(eid: int) -> Path:
    return OUT / "overlays" / f"e{eid}" / "rowagent.json"


def start_rows(eid: int) -> dict:
    """Queue the row agents: this year's row for each row the Summary is waiting on, found and checked."""
    rows = _q("SELECT overlay_status, rows_status FROM engagements WHERE id=?", eid)
    if not rows:
        raise ValueError("no such engagement")
    if rows[0]["overlay_status"] != "done":
        raise ValueError("rebuild in Python first: the row agents work on the Python overlay")
    if not _role_wb(eid, "current_model"):
        raise ValueError("assign this year's client model (Roles) first")
    if rows[0]["rows_status"] not in ("queued", "running"):
        _set("engagements", eid, rows_status="queued", rows_step="Waiting to start", rows_error=None)
        _jobs.put(("rows", eid))
    return rows_view(eid)


def _rows_job(eid: int) -> None:
    import rowagent
    import rowfind
    step = lambda msg: _set("engagements", eid, rows_status="running", rows_step=msg)
    step("Loading the Python overlay")
    sess, summary = overlay_session(eid)
    _sync_roll(eid, sess, summary)
    reader, why = None, None
    try:  # luna proposes, sol checks: the engagement's model and its reviewer
        e = _q("SELECT model, reviewer_model FROM engagements WHERE id=?", eid)[0]
        reader = docingest.Reader(e["model"] or DEFAULT_MODEL, e["reviewer_model"] or DEFAULT_REVIEWER, _logger(eid))
    except Exception as ex:
        why = friendly(ex)
    res = rowagent.run(sess, summary, reference(eid), step, reader)
    if reader is not None:
        res["models"] = {"proposes": reader.model, "checks": reader.reviewer_model}
    else:
        res["models_error"] = why
    # the agents' picks, beside a person's (a person's always win; the agents' from an earlier run are replaced)
    picks = {k: v for k, v in _read_rowpicks(eid).items() if not (isinstance(v, dict) and v.get("by") == "agent")}
    for d in res["decisions"]:
        if d.get("decision") and d["row"] not in picks:
            picks[d["row"]] = {"to": d["decision"], "by": "agent", "v": rowfind.VERSION, "why": d.get("why"),
                               "checked_by": "the numbers" if d.get("how") == "numbers" else d.get("review") or d.get("how")}
    _write_rowpicks(eid, picks)
    res["at"], res["v"] = time.time(), rowfind.VERSION
    f = _rowagent_file(eid)
    f.parent.mkdir(parents=True, exist_ok=True)
    f.write_text(json.dumps(res, default=str), encoding="utf-8")
    _set("engagements", eid, rows_status="done", rows_step="Done")


def rows_view(eid: int) -> dict:
    rows = _q("SELECT rows_status, rows_step, rows_error, rows_secs, overlay_started_at FROM engagements WHERE id=?", eid)
    if not rows:
        raise ValueError("no such engagement")
    e = rows[0]
    f = _rowagent_file(eid)
    res = json.loads(f.read_text(encoding="utf-8")) if f.exists() else None
    import rowfind
    built = max([(_role_wb(eid, k) or {}).get("processed_at") or 0 for k in ("prior_overlay", "prior_model", "current_model")]
                + [e["overlay_started_at"] or 0])
    # made before the Python overlay or a workbook was built again, or by another version of the row finder
    stale = bool(res and (res["at"] < built or res.get("v") != rowfind.VERSION))
    return {"status": e["rows_status"], "step": e["rows_step"], "error": e["rows_error"], "secs": e["rows_secs"],
            "result": res, "stale": stale}


def start_doctor(eid: int) -> dict:
    """Queue the doctor: where each figure breaks, why, and whether the files and rows are the right ones."""
    rows = _q("SELECT overlay_status, doctor_status FROM engagements WHERE id=?", eid)
    if not rows:
        raise ValueError("no such engagement")
    if rows[0]["overlay_status"] != "done":
        raise ValueError("rebuild in Python first: the doctor examines the Python overlay")
    if rows[0]["doctor_status"] in ("queued", "running"):
        return doctor_view(eid)
    _set("engagements", eid, doctor_status="queued", doctor_step="Waiting to start", doctor_error=None)
    _jobs.put(("doctor", eid))
    return doctor_view(eid)


def _doctor_job(eid: int) -> None:
    import doctor
    import overlay as ovmod
    step = lambda frac, msg: _set("engagements", eid, doctor_status="running", doctor_step=msg)
    step(0, "Loading the Python overlay")
    sess, summary = overlay_session(eid)
    evidence = ovmod.deep(doctor.examine, sess, summary, step)
    step(0.9, "Writing up the diagnosis")
    e = _q("SELECT model, reviewer_model FROM engagements WHERE id=?", eid)[0]
    res = {"at": time.time(), "evidence": evidence, "diagnosis": None, "diagnosis_error": None}
    try:
        reader = docingest.Reader(e["model"] or DEFAULT_MODEL, e["reviewer_model"] or DEFAULT_REVIEWER, _logger(eid))
        res["diagnosis"] = doctor.diagnose(reader, evidence)
        res["model"] = reader.reviewer_model
    except Exception as ex:  # the evidence stands on its own
        traceback.print_exc()
        res["diagnosis_error"] = friendly(ex)
    res["text"] = doctor.report_text(res)
    f = _doctor_file(eid)
    f.parent.mkdir(parents=True, exist_ok=True)
    f.write_text(json.dumps(res, default=str), encoding="utf-8")
    _set("engagements", eid, doctor_status="done", doctor_step="Done")


def doctor_view(eid: int) -> dict:
    rows = _q("SELECT overlay_status, overlay_started_at, doctor_status, doctor_step, doctor_error, doctor_secs "
              "FROM engagements WHERE id=?", eid)
    if not rows:
        raise ValueError("no such engagement")
    e = rows[0]
    f = _doctor_file(eid)
    res = json.loads(f.read_text(encoding="utf-8")) if f.exists() else None
    held = json.loads(_holds_file(eid).read_text(encoding="utf-8")) if _holds_file(eid).exists() else {}
    return {"status": e["doctor_status"], "step": e["doctor_step"], "error": e["doctor_error"], "secs": e["doctor_secs"],
            "overlay_status": e["overlay_status"], "result": res,
            "stale": bool(res and e["overlay_started_at"] and res["at"] < e["overlay_started_at"]),
            "held": [{"cell": c, **h} for c, h in held.items()]}


def doctor_holds(eid: int, cells: list[str] | None, release: bool = False) -> dict:
    """Hold the doctor's safe cells at Excel's saved value on every feed (cells: all of them when None), or
    release every hold. Only cells the doctor found safe can be held."""
    import overlay as ovmod
    f = _holds_file(eid)
    held = json.loads(f.read_text(encoding="utf-8")) if f.exists() else {}
    if release:
        held = {}
    else:
        res = json.loads(_doctor_file(eid).read_text(encoding="utf-8")) if _doctor_file(eid).exists() else None
        safe = {h["cell"]: h for h in ((res or {}).get("evidence", {}).get("holds") or {}).get("safe") or []}
        if not safe:
            raise ValueError("the doctor found nothing that can be held: run it first")
        pick = [c for c in (cells or list(safe)) if c in safe]
        sess, summary = overlay_session(eid)
        with rodb.connect(sess.ov.path) as db:
            for c in pick:
                k = ovmod.parse_a1(c)
                row = db.execute("SELECT formula FROM cells WHERE sheet=? AND row=? AND col=?", k).fetchone()
                held[c] = {"value": safe[c]["value"], "formula": row[0] if row else None, "why": safe[c]["title"],
                           "at": time.time()}
    f.parent.mkdir(parents=True, exist_ok=True)
    f.write_text(json.dumps(held, default=str), encoding="utf-8")
    if eid in _SESSIONS:
        import overlay as ovmod
        ovmod.deep(_load_holds, eid, _SESSIONS[eid][0])
    return doctor_view(eid)


def _process_doc(did: int) -> None:
    d = _doc(did)
    eid = d["engagement_id"]
    e = _q("SELECT model, reviewer_model FROM engagements WHERE id=?", eid)[0]
    prog = lambda frac, msg: _set("documents", did, pct=round(frac, 3), step=msg)
    _set("documents", did, status="processing", step="Opening document", pct=0, error=None)
    args = (d["source_path"], d["out_dir"], e["model"] or DEFAULT_MODEL, e["reviewer_model"] or DEFAULT_REVIEWER, prog,
            _logger(eid))
    note = None
    try:
        doc = docingest.process(*args, key_only=True)  # the tables with the key figures now, the rest afterwards
    except Exception as ex:  # not signed in: keep the text layer, leave the tables to read on retry
        if type(ex).__name__ not in ("AuthenticationRequiredError", "ClientAuthenticationError"):
            raise
        doc = docingest.process(*args, read=False)
        note = "Tables not read. " + friendly(ex)
    _save_doc(did, doc)
    if _q("SELECT 1 FROM facts WHERE document_id=?", did):  # read again: the facts are checked against the new text
        _recheck_facts(did, doc)
    if note is None and any(t.get("status") == "flagged" for t in doc["tables"]):
        _set("documents", did, step="Table review loop on the flagged tables", pct=0.95)
        try:
            _run_table_loop(did, doc)
        except Exception as ex:  # the tables stay flagged for a person
            traceback.print_exc()
            note = "The review loop on the flagged tables failed: " + friendly(ex)
    errors = [t for t in doc["tables"] if t.get("status") == "error"]
    if errors:
        note = f"{len(errors)} table(s) couldn't be read: {errors[0].get('error')}"
    left = sum(bool(t.get("deferred")) and t.get("status") == "unread" for t in doc["tables"])
    _set("documents", did, status="done", step="Done", pct=1.0, processed_at=time.time(), error=note,
         tables_status="queued" if left and note is None else None, tables_left=left, tables_error=None)
    if note is None and not _q("SELECT 1 FROM facts WHERE document_id=?", did):
        extract_facts(did)  # the key facts, their review and the review loop start without a button
    if left and note is None:
        _start_rest(did)  # the other tables, alongside


def _process_facts(did: int) -> None:
    d = _doc(did)
    eid = d["engagement_id"]
    e = _q("SELECT model, reviewer_model, arbiter_model FROM engagements WHERE id=?", eid)[0]
    doc = json.loads(d["doc_json"])
    _set("documents", did, facts_status="running", facts_step="Extracting key facts")
    res = reportfacts.run(doc["markdown"], e["model"] or DEFAULT_MODEL, e["reviewer_model"] or DEFAULT_REVIEWER, _logger(eid),
                          lambda frac, msg: _set("documents", did, facts_step=msg),
                          arbiter_model=e["arbiter_model"] or DEFAULT_ARBITER)
    now = time.time()
    with _lock, _conn() as db:
        db.execute("DELETE FROM facts WHERE document_id=?", (did,))
        for f in res["facts"]:
            db.execute(f"""INSERT INTO facts(engagement_id, document_id, n, {', '.join(FACT_FIELDS)}, origin, check_json,
                           review_json, status, updated_at, agent_json) VALUES ({', '.join('?' * (len(FACT_FIELDS) + 9))})""",
                       (eid, did, f["id"], *[f.get(k) for k in FACT_FIELDS], f["origin"], json.dumps(f["check"]),
                        json.dumps(f.get("review")), "pending", now, json.dumps(f.get("agent"))))
    _AFTER.pop(did, None)  # facts reopened during the extraction are gone: every fact is new and checked afresh
    _set("documents", did, facts_notes=res.get("notes"), review_summary=res.get("review_summary"))
    if res.get("loop"):
        _note_loop(did, facts={**res["loop"]["summary"], **auto_decide(did), "at": now}, lessons_facts=None)
        _set("documents", did, facts_step="Writing down what the loop taught")
        _learn(did, "facts", res["loop"]["episodes"])
    _set("documents", did, facts_status="done", facts_step="Done")
    _touch(eid)


_STEP = {"doc": "read the report", "resolve_tables": "table review loop", "facts": "key facts",
         "resolve_facts": "fact review loop", "compare": "compare models", "map": "map", "overlay": "python overlay",
         "roles": "roles", "charts": "report charts", "doctor": "overlay doctor", "rows": "row agents"}
# where each job's start time and duration go: (table, started column, seconds column)
_TIMED = {"doc": ("documents", "started_at", "doc_secs"), "resolve_tables": ("documents", "started_at", "doc_secs"),
          "facts": ("documents", "facts_started_at", "facts_secs"),
          "resolve_facts": ("documents", "facts_started_at", "facts_secs"),
          **{k: ("engagements", f"{k}_started_at", f"{k}_secs") for k in ("compare", "map", "overlay", "charts", "doctor", "rows")}}


def _run(kind: str, rid: int) -> None:
    t0 = time.time()
    timed = _TIMED.get(kind)
    if timed:
        _set(timed[0], rid, **{timed[1]: t0})
    doc = _doc(rid) if kind in ("doc", "resolve_tables", "facts", "resolve_facts") else None
    tags = {"engagement": doc["engagement_id"] if doc else rid, "document": rid if doc else None, "step": _STEP.get(kind, kind)}
    with calllog.tag(**tags):  # every model call in the job is logged against its engagement, document and step
        ok = _run_job(kind, rid)
    if ok and timed:  # how long it took: the page's estimate for the next run of the same job
        _set(timed[0], rid, **{timed[2]: round(time.time() - t0, 1)})


def _run_job(kind: str, rid: int) -> bool:
    """Run one job; a failure is recorded where the page shows it. True if it finished."""
    try:
        {"doc": _process_doc, "facts": _process_facts, "compare": _compare, "map": _map, "overlay": _overlay,
         "resolve_tables": _resolve_tables_job, "resolve_facts": _resolve_facts_job, "roles": _suggest_job,
         "charts": _charts_job, "doctor": _doctor_job, "rows": _rows_job}[kind](rid)
        return True
    except Exception as e:
        traceback.print_exc()
        msg = friendly(e)
        if kind == "doc":
            _set("documents", rid, status="error", step="Failed", error=msg)
        elif kind == "resolve_tables":  # the document itself is still fine
            _set("documents", rid, status="done", step="Done", pct=1.0, error="The review loop failed: " + msg)
            _note_loop(rid, tables={"error": msg, "at": time.time()})  # tried: not started again by itself
        elif kind == "roles":
            pass  # suggest_roles kept the error with the suggestion
        elif kind in ("facts", "resolve_facts"):
            _set("documents", rid, facts_status="error", facts_step="Failed", facts_error=msg)
        else:
            _set("engagements", rid, **{f"{kind}_status": "error", f"{kind}_step": "Failed", f"{kind}_error": msg})
        return False


def _worker(lane: str) -> None:
    q = _jobs.queues[lane]
    while True:
        kind, rid = q.get()
        _jobs.running[lane] = (kind, rid, time.time())
        try:
            _run(kind, rid)
        finally:
            _jobs.running.pop(lane, None)
            q.task_done()


# what a job is, for a person waiting behind it ({file}: the report's name)
_BEHIND = {"doc": "reading {file}", "resolve_tables": "the table review loop on {file}", "facts": "the key facts from {file}",
           "resolve_facts": "the fact review loop on {file}", "compare": "the compare of the client models", "map": "the map",
           "overlay": "the Python overlay", "roles": "the roles suggestion", "charts": "the report's charts",
           "doctor": "the overlay doctor", "rows": "the row agents"}


def _job_owner(kind: str, rid: int) -> tuple[int | None, str]:
    """The engagement a job belongs to, and the report's name for a report job."""
    if kind not in REPORT_JOBS:
        return rid, ""
    d = _q("SELECT engagement_id, filename FROM documents WHERE id=?", rid)
    return (d[0]["engagement_id"], d[0]["filename"]) if d else (None, "")


def jobs_view(eid: int) -> list[dict]:
    """This engagement's background jobs as they stand: running, or queued with the job running ahead of it in its
    lane (another engagement's included) and how many are ahead, so the page can say what a queued step waits for."""
    out = []
    for lane, q in _jobs.queues.items():
        with q.mutex:
            waiting = list(q.queue)
        run, ahead = _jobs.running.get(lane), None
        if run:
            owner, file = _job_owner(run[0], run[1])
            label = _BEHIND[run[0]].format(file=file) if owner == eid else f"{_STEP[run[0]]} for another engagement"
            ahead = {"label": label, "since": run[2]}
            if owner == eid:
                out.append({"lane": lane, "job": run[0], "id": run[1], "state": "running", "since": run[2]})
        for i, (kind, rid) in enumerate(waiting):
            if _job_owner(kind, rid)[0] == eid:
                out.append({"lane": lane, "job": kind, "id": rid, "state": "queued", "behind": ahead,
                            "ahead": i + bool(run)})
    return out


def rebuild_document(did: int) -> dict:
    """Read a report again from scratch (every table read again, the table loop again); its key facts are kept and
    re-checked against the new reading."""
    d = _doc(did)
    if not d:
        raise ValueError("no such document")
    if d["status"] in ("queued", "processing") or d["facts_status"] in ("queued", "running") or did in _READING:
        raise ValueError(f"{d['filename']} is already being worked on")
    _note_loop(did, tables=None, lessons_tables=None)
    retry_document(did)
    return document(did)


def rebuild_workbook(eid: int, fid: int) -> dict:
    """Process a workbook again from scratch. Its model.db is shared (other engagements, the Model Desk), so every
    live Python overlay reading it lets go of it first: the build deletes the file, which Windows won't do while
    it's open."""
    return library.rebuild(_release(eid, fid)["id"], _model(eid))


def retry_workbook(eid: int, fid: int) -> dict:
    """Process a failed workbook again. It may have failed because a live Python overlay held its model.db (Windows
    won't delete an open file), so those let go of it first, as for a rebuild."""
    return library.retry(_release(eid, fid)["id"])


def _release(eid: int, fid: int) -> dict:
    """Every live Python overlay reading the workbook's model.db lets go of it (its build deletes the file)."""
    w = next((w for w in workbooks(eid) if w["id"] == fid), None)
    if not w:
        raise ValueError("that workbook isn't in this engagement")
    for other, (sess, _) in list(_SESSIONS.items()):
        if w["db_path"] and w["db_path"] in sess.paths():
            try:
                import overlay as ovmod
                ovmod.deep(sess.close)
            finally:
                _SESSIONS.pop(other, None)
    for key in [k for k in _SHEET_NAMES if k[0] == w["db_path"]]:
        _SHEET_NAMES.pop(key, None)
    return w


def retry_document(did: int) -> None:
    _set("documents", did, status="queued", step="Waiting to start", pct=0, error=None)
    _jobs.put(("doc", did))


def start_worker() -> None:
    """Re-queue anything a restart interrupted, then start the worker."""
    for r in _q("SELECT id, step, doc_json IS NOT NULL AS read FROM documents WHERE status IN ('queued','processing')"):
        # a document already read and in its table loop only needs the loop again
        _jobs.put(("resolve_tables" if r["read"] and "review loop" in (r["step"] or "").lower() else "doc", r["id"]))
    for r in _q("SELECT id FROM documents WHERE status='done' AND tables_status IN ('queued','reading')"):
        _start_rest(r["id"])  # background table reading a restart interrupted
    for r in _q("SELECT id FROM documents WHERE facts_status IN ('queued','running')"):
        # an interrupted loop reruns as a loop (on the facts saved); an interrupted extraction starts again
        _jobs.put(("resolve_facts" if _q("SELECT 1 FROM facts WHERE document_id=?", r["id"]) else "facts", r["id"]))
    for kind in ("compare", "map", "overlay", "charts", "doctor", "rows"):
        for r in _q(f"SELECT id FROM engagements WHERE {kind}_status IN ('queued','running')"):
            _jobs.put((kind, r["id"]))
    for lane in _jobs.queues:
        threading.Thread(target=_worker, args=(lane,), daemon=True, name=f"engagement-{lane}").start()

