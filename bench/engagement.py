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
DEFAULT_ARBITER = "gpt-4o"        # settles what the review loop can't: a third model, independent of both
REPORT_TYPES = (".pdf", ".pptx")

_lock = threading.Lock()
_jobs: "queue.Queue[tuple[str, int]]" = queue.Queue()

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
    for table, col in (("facts", "agent_json"), ("documents", "loop_json"), ("facts", "decided_by")):  # ... and before the loop
        if col not in {r[1] for r in db.execute(f"PRAGMA table_info({table})")}:
            db.execute(f"ALTER TABLE {table} ADD COLUMN {col} TEXT")
    if "arbiter_model" not in {r[1] for r in db.execute("PRAGMA table_info(engagements)")}:
        db.execute("ALTER TABLE engagements ADD COLUMN arbiter_model TEXT")
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
    for k in ("roles_suggested", "compare_json", "map_json", "overlay_json"):
        e[k.removesuffix("_json")] = json.loads(e.pop(k) or "null")
    wbs = workbooks(eid)
    for w in wbs:
        w.pop("db_path", None)
        w.pop("source_path", None)
        ident = w.pop("identity", None) or {}
        w["identity_notes"] = ident.get("notes")
    out = {**e, "documents": documents(eid), "workbooks": wbs, "facts": facts(eid), "roles": roles(eid),
           "session": _session(eid), "now": time.time()}
    _auto_review(out)
    _maybe_suggest(eid, out)
    out["roles_pending"] = eid in _ROLE_JOBS
    return out


def delete(eid: int) -> None:
    for d in documents(eid):
        remove_document(d["id"])
    with _lock, _conn() as db:
        for t, col in (("facts", "engagement_id"), ("roles", "engagement_id"), ("eng_files", "engagement_id"),
                       ("engagements", "id")):
            db.execute(f"DELETE FROM {t} WHERE {col}=?", (eid,))


# ---- uploads ------------------------------------------------------------------------------------------------

def add_upload(eid: int, tmp: Path, filename: str, sha: str) -> dict:
    """Workbooks -> the shared library (deduplicated across the app); reports -> this engagement's documents."""
    ext = Path(filename).suffix.lower()
    if ext in library.SUPPORTED:
        status, rec = library.add_upload(tmp, filename, sha)
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
    if action == "approve":
        t.update(status="approved", final_markdown=t.get("markdown"))
    elif action == "use_second":
        t.update(status="approved", final_markdown=t.get("second_markdown"))
    elif action == "edit":
        t.update(status="edited", final_markdown=markdown or "")
    elif action == "reset":
        t.pop("final_markdown", None)
        t.pop("final_check", None)
        t["status"] = "verified" if (t.get("check") or t.get("second_read") or {}).get("ok") else "flagged"
    else:
        raise ValueError(f"unknown action {action}")
    if t.get("final_markdown") is not None and t.get("text_lines"):
        t["final_check"] = docingest.check_text_layer(t["final_markdown"], t["text_lines"])
    _save_doc(did, doc)
    _recheck_facts(did, doc)
    reopened = [f for f in _q("SELECT agent_json FROM facts WHERE document_id=? AND status='pending'", did)
                if (json.loads(f["agent_json"] or "null") or {}).get("status") == "escalated"]
    if reopened and _doc(did)["facts_status"] not in ("queued", "running"):
        resolve_facts(did)  # the loop takes them again with the page as it now reads
    return document(did)


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
        e = _q("SELECT model, reviewer_model FROM engagements WHERE id=?", eid)[0]
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
            res = docingest.resolve_tables(cur, d["out_dir"], model, reviewer, None, _logger(eid), ids=set(todo))
            _merge_tables(did, [t for t in cur["tables"] if t.get("resolution") and t["id"] in todo],
                          lambda c: c.get("status") == "flagged")
            _learn(did, "tables", res["episodes"])
        else:
            res = {"resolved": 0, "escalated": 0}
        prev = (json.loads(_doc(did)["loop_json"] or "{}").get("tables") or {})
        _note_loop(did, tables={"resolved": (prev.get("resolved") or 0) + res["resolved"], "escalated": res["escalated"],
                                "at": time.time()})  # this reading's loop has run: nothing starts another
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


def _recheck_facts(did: int, doc: dict) -> None:
    """Table decisions change the page text facts are checked against (no model calls). An approval the agents
    made rested on the checks passing, so if they now fail the fact goes back to a person."""
    pg = reportfacts.pages(doc["markdown"])
    for f in _q("SELECT * FROM facts WHERE document_id=?", did):
        agent = json.loads(f["agent_json"] or "null") or {}
        chk = reportfacts.check({**f, "waivers": agent.get("waivers")}, pg)
        rv = json.loads(f["review_json"] or "null")
        if rv and rv.get("suggestion"):
            rv["suggestion"]["check"] = reportfacts.check(rv["suggestion"], pg)
        fields = {"check_json": json.dumps(chk), "review_json": json.dumps(rv)}
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
        nums = docingest.numbers(final.get("value_text") or "")
        if "value" not in (fields or {}) and nums and final.get("unit") != "date":
            final["value"] = float(nums[0].rstrip("%"))
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


def resolve_facts(did: int) -> None:
    if not [r for r in _q("SELECT agent_json FROM facts WHERE document_id=? AND status='pending'", did)
            if (json.loads(r["agent_json"] or "null") or {}).get("status") not in ("agreed", "withdrawn")]:
        raise ValueError("no open facts: the agents agree on every fact still waiting for a decision")
    _set("documents", did, facts_status="queued", facts_step="Waiting for the review loop", facts_error=None)
    _jobs.put(("resolve_facts", did))


def _run_table_loop(did: int, doc: dict) -> None:
    d = _doc(did)
    eid = d["engagement_id"]
    e = _q("SELECT model, reviewer_model FROM engagements WHERE id=?", eid)[0]
    prog = lambda frac, msg: _set("documents", did, pct=round(frac, 3), step=msg)
    res = docingest.resolve_tables(doc, d["out_dir"], e["model"] or DEFAULT_MODEL, e["reviewer_model"] or DEFAULT_REVIEWER,
                                   prog, _logger(eid))
    _save_doc(did, doc)
    _note_loop(did, tables={"resolved": res["resolved"], "escalated": res["escalated"], "at": time.time()},
               lessons_tables=None)  # this loop's lessons replace the last one's note
    _learn(did, "tables", res["episodes"])


def _resolve_tables_job(did: int) -> None:
    d = _doc(did)
    _set("documents", did, status="processing", step="Table review loop on the flagged tables", pct=0)
    doc = json.loads(d["doc_json"])
    eid = d["engagement_id"]
    e = _q("SELECT model, reviewer_model FROM engagements WHERE id=?", eid)[0]
    prog = lambda frac, msg: _set("documents", did, pct=round(frac, 3), step=msg)
    res = docingest.resolve_tables(doc, d["out_dir"], e["model"] or DEFAULT_MODEL, e["reviewer_model"] or DEFAULT_REVIEWER,
                                   prog, _logger(eid))
    doc = _merge_tables(did, [t for t in doc["tables"] if t.get("resolution")], lambda c: c.get("status") == "flagged")
    _note_loop(did, tables={"resolved": res["resolved"], "escalated": res["escalated"], "at": time.time()},
               lessons_tables=None)
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
    rows = [r for r in _q("SELECT * FROM facts WHERE document_id=? AND status='pending' ORDER BY n", did)
            if (json.loads(r["agent_json"] or "null") or {}).get("status") not in ("agreed", "withdrawn")]
    if not rows:
        _set("documents", did, facts_status="done", facts_step="Done")
        return
    fs = []
    for r in rows:
        f = {"id": r["id"], "origin": r["origin"], **{k: r[k] for k in FACT_FIELDS},
             "check": json.loads(r["check_json"] or "null"), "review": json.loads(r["review_json"] or "null") or {}}
        prev = json.loads(r["agent_json"] or "null")
        f["waivers"] = (prev or {}).get("waivers") or []
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
            _recheck_facts(did, json.loads(_doc(did)["doc_json"]))  # checks improved: facts they held up can settle
            auto_decide(did)
            _note_loop(did, check_version=reportfacts.CHECK_VERSION)
            e["facts"] = facts(e["id"])
        if d["n_flagged"] and not loop.get("tables") and not d["error"]:
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
        for m in fm["matches"]:
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
            if x["matches"]:
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
    return {"overlay": {"db_path": ov["db_path"], "filename": ov["filename"], "sheets": sheets},
            "prior": {"db_path": prior["db_path"], "filename": prior["filename"],
                      "sheets": sorted(prior["sheets"]) if prior["sheets"] else None} if prior else None,
            "current": {"db_path": cur["db_path"], "filename": cur["filename"], "sheets": None} if cur else None,
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
    _SESSIONS[eid] = (sess, summary)
    _set("engagements", eid, overlay_json=json.dumps(summary, default=str), overlay_status="done", overlay_step="Done",
         updated_at=time.time())


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
    _SESSIONS[eid] = (sess, summary)
    return _SESSIONS[eid]


def _live(eid: int, mode: str, changes: dict | None) -> tuple:
    """The engagement's session and summary, after checking the feed exists; changes with dates as serials."""
    import overlay as ovmod
    sess, summary = overlay_session(eid)
    if mode not in ("workbook", "prior", "current"):
        raise ValueError("the feed must be workbook, prior or current")
    if mode == "current" and not summary["wiring"].get("current"):
        raise ValueError("assign the current client model (step 3) and rebuild to roll forward")
    if mode == "prior" and not summary["wiring"].get("prior"):
        raise ValueError("assign the prior client model (step 3) to feed from it")
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
    v = valuation.validation(summary["wiring"]["overlay"]["db_path"], pick["cell"])
    v["anchors"] = valuation._listing(anchors)
    return v


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
    _set("documents", did, facts_notes=res.get("notes"), review_summary=res.get("review_summary"))
    if res.get("loop"):
        _note_loop(did, facts={**res["loop"]["summary"], **auto_decide(did), "at": now}, lessons_facts=None)
        _set("documents", did, facts_step="Writing down what the loop taught")
        _learn(did, "facts", res["loop"]["episodes"])
    _set("documents", did, facts_status="done", facts_step="Done")
    _touch(eid)


_STEP = {"doc": "read the report", "resolve_tables": "table review loop", "facts": "key facts",
         "resolve_facts": "fact review loop", "compare": "compare models", "map": "map", "overlay": "python overlay",
         "roles": "roles"}
# where each job's start time and duration go: (table, started column, seconds column)
_TIMED = {"doc": ("documents", "started_at", "doc_secs"), "resolve_tables": ("documents", "started_at", "doc_secs"),
          "facts": ("documents", "facts_started_at", "facts_secs"),
          "resolve_facts": ("documents", "facts_started_at", "facts_secs"),
          **{k: ("engagements", f"{k}_started_at", f"{k}_secs") for k in ("compare", "map", "overlay")}}


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
         "resolve_tables": _resolve_tables_job, "resolve_facts": _resolve_facts_job, "roles": _suggest_job}[kind](rid)
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


def _worker() -> None:
    while True:
        kind, rid = _jobs.get()
        try:
            _run(kind, rid)
        finally:
            _jobs.task_done()


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
    w = next((w for w in workbooks(eid) if w["id"] == fid), None)
    if not w:
        raise ValueError("that workbook isn't in this engagement")
    for other, (sess, _) in list(_SESSIONS.items()):
        if w["db_path"] in sess.paths():
            try:
                import overlay as ovmod
                ovmod.deep(sess.close)
            finally:
                _SESSIONS.pop(other, None)
    for key in [k for k in _SHEET_NAMES if k[0] == w["db_path"]]:
        _SHEET_NAMES.pop(key, None)
    return library.rebuild(fid)


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
    for kind in ("compare", "map", "overlay"):
        for r in _q(f"SELECT id FROM engagements WHERE {kind}_status IN ('queued','running')"):
            _jobs.put((kind, r["id"]))
    threading.Thread(target=_worker, daemon=True, name="engagement-worker").start()

