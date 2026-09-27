"""Engagements: the four files of a recurring valuation, from upload to an approved reference, roles and a map.

  1. upload     workbooks go through the Model Desk pipeline (library.py: fingerprint, build, identify) and are
                shared with it; reports (PDF / PPTX) are read here by docingest.py, tables checked
  2. reference  reportfacts.py extracts the report's key facts, checks them in code and has a reviewer model
                review them; a person approves, edits or rejects each one
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
import linkmap
import reportfacts
import roles as rolesmod
import usage

ROOT = Path(__file__).resolve().parent.parent
OUT, UPLOADS = ROOT / "out", ROOT / "uploads"
DB, DOCS = OUT / "engage.db", OUT / "docs"
DEFAULT_MODEL = "gpt-6-luna"      # extraction and table reads
DEFAULT_REVIEWER = "gpt-6-sol"    # second reads and fact review: a different, stronger model than the first read
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
  facts_status TEXT, facts_step TEXT, facts_error TEXT, facts_notes TEXT, review_summary TEXT,
  UNIQUE(engagement_id, sha256));
CREATE TABLE IF NOT EXISTS facts(id INTEGER PRIMARY KEY, engagement_id INT, document_id INT, n INT, category TEXT,
  key TEXT, label TEXT, value_text TEXT, low_text TEXT, high_text TEXT, value REAL, unit TEXT, basis TEXT, page INT,
  quote TEXT, origin TEXT, check_json TEXT, review_json TEXT, status TEXT DEFAULT 'pending', final_json TEXT,
  updated_at REAL);
CREATE TABLE IF NOT EXISTS roles(engagement_id INT, role TEXT, kind TEXT, ref_id INT, sheets_json TEXT, why_json TEXT,
  confirmed INT DEFAULT 0, PRIMARY KEY(engagement_id, role));
"""
FACT_FIELDS = ("category", "key", "label", "value_text", "low_text", "high_text", "value", "unit", "basis", "page", "quote")


def _conn() -> sqlite3.Connection:
    OUT.mkdir(exist_ok=True)
    db = sqlite3.connect(DB, check_same_thread=False)
    db.row_factory = sqlite3.Row
    db.executescript(SCHEMA)
    have = {r[1] for r in db.execute("PRAGMA table_info(engagements)")}
    for col in ("overlay_status", "overlay_step", "overlay_error", "overlay_json"):  # databases made before step 6
        if col not in have:
            db.execute(f"ALTER TABLE engagements ADD COLUMN {col} TEXT")
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
    fields = {k: v for k, v in fields.items() if k in ("name", "model", "reviewer_model") and v}
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
                                         "db_path", "source_path", "identity")}
            w["sheet_names"] = []
            if w["status"] == "done" and w["db_path"] and Path(w["db_path"]).exists():
                with sqlite3.connect(f"file:{w['db_path']}?mode=ro", uri=True) as m:
                    w["sheet_names"] = [r[0] for r in m.execute("SELECT sheet FROM sheets ORDER BY rowid")]
            out.append(w)
    return out


def documents(eid: int) -> list[dict]:
    cols = ("id, engagement_id, filename, kind, size, uploaded_at, status, step, pct, error, pages, n_tables, n_flagged, "
            "processed_at, facts_status, facts_step, facts_error, facts_notes, review_summary")
    return _q(f"SELECT {cols} FROM documents WHERE engagement_id=? ORDER BY uploaded_at", eid)


def facts(eid: int) -> list[dict]:
    out = []
    for f in _q("SELECT * FROM facts WHERE engagement_id=? ORDER BY document_id, n", eid):
        for k in ("check_json", "review_json", "final_json"):
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
            if f["status"] != "rejected" and (f["check"] or {}).get("ok")]


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
    return {**e, "documents": documents(eid), "workbooks": wbs, "facts": facts(eid), "roles": roles(eid),
            "session": _session(eid)}


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
    return {**d, "doc": doc}


def table_png(did: int, tid: str) -> Path | None:
    d = _doc(did)
    p = Path(d["out_dir"]) / "tables" / f"{tid}.png" if d else None
    return p if p and p.is_file() and p.parent.parent == Path(d["out_dir"]) else None


def settle_table(did: int, tid: str, action: str, markdown: str | None = None) -> dict:
    """A person's decision on a table: approve the first read, take the reviewer's read, or edit it."""
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
    return document(did)


def _save_doc(did: int, doc: dict) -> None:
    md = docingest.render(doc)
    d = _doc(did)
    Path(d["out_dir"], "document.md").write_text(md)
    flagged = sum(t.get("status") in ("flagged", "error", "unread") for t in doc["tables"])
    _set("documents", did, doc_json=json.dumps(doc, default=str), n_flagged=flagged,
         n_tables=sum(t.get("status") != "figure" for t in doc["tables"]), pages=len(doc["pages"]))


def _recheck_facts(did: int, doc: dict) -> None:
    """Table decisions change the page text facts are checked against (no model calls)."""
    pg = reportfacts.pages(doc["markdown"])
    for f in _q("SELECT * FROM facts WHERE document_id=?", did):
        chk = reportfacts.check(f, pg)
        rv = json.loads(f["review_json"] or "null")
        if rv and rv.get("suggestion"):
            rv["suggestion"]["check"] = reportfacts.check(rv["suggestion"], pg)
        _set("facts", f["id"], check_json=json.dumps(chk), review_json=json.dumps(rv))


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
    rv = json.loads(f["review_json"] or "null") or {}
    now = time.time()
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


def approve_passed(eid: int) -> int:
    """Approve every pending fact that passed the code checks and that the reviewer accepted as extracted."""
    n = 0
    for f in facts(eid):
        if f["status"] == "pending" and (f["check"] or {}).get("ok") and (f["review"] or {}).get("verdict") == "accept":
            set_fact(f["id"], "approve")
            n += 1
    return n


# ---- roles --------------------------------------------------------------------------------------------------

def _wb_inputs(eid: int) -> list[dict]:
    out = []
    for w in workbooks(eid):
        if w["status"] == "done" and w["db_path"] and Path(w["db_path"]).exists():
            extlinks.ensure(w["source_path"], w["db_path"])
            out.append(w)
    return out


def suggest_roles(eid: int) -> dict:
    """Suggest (doesn't overwrite confirmed roles). Stored so the page can show it next to what's confirmed."""
    docs = [d for d in documents(eid) if d["status"] == "done"]
    n_facts = {d["id"]: sum(1 for f in facts(eid) if f["document_id"] == d["id"]) for d in docs}
    res = rolesmod.suggest([{**d, "n_facts": n_facts[d["id"]]} for d in docs], _wb_inputs(eid), reference(eid))
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


# ---- the Python overlay -------------------------------------------------------------------------------------

_SESSIONS: dict[int, tuple] = {}  # engagement -> (overlay.Session, summary): the compiled module, loaded and wired


def _wiring(eid: int) -> dict:
    """Which files and sheets the overlay module reads, from the confirmed (or suggested) roles."""
    ov, prior, cur = _role_wb(eid, "prior_overlay"), _role_wb(eid, "prior_model"), _role_wb(eid, "current_model")
    sheets = [s for s in next(w for w in workbooks(eid) if w["id"] == ov["id"])["sheet_names"]
              if not ov["sheets"] or s in ov["sheets"]]
    same_file = bool(prior) and prior["id"] == ov["id"]
    client_link = None
    if prior and not same_file:
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
            "client_link": client_link, "prior_valuation_date": prior_vd, "same_file": same_file}


def _overlay(eid: int) -> None:
    import overlay as ovmod
    e = _q("SELECT name FROM engagements WHERE id=?", eid)[0]
    step = lambda frac, msg: _set("engagements", eid, overlay_status="running", overlay_step=msg)
    step(0, "Reading the roles")
    w = _wiring(eid)
    _SESSIONS.pop(eid, None)
    summary, sess = ovmod.build(OUT / "overlays" / f"e{eid}", w["overlay"], w["prior"], w["current"], reference(eid),
                                e["name"], w["client_link"], w["prior_valuation_date"], step)
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
                         w["client_link"], (w["prior"] or {}).get("sheets") if w["same_file"] else None)
    _SESSIONS[eid] = (sess, summary)
    return _SESSIONS[eid]


def overlay_run(eid: int, mode: str, changes: dict, valuation_date: str | None, months: int | None) -> dict:
    import overlay as ovmod
    sess, summary = overlay_session(eid)
    if mode == "current" and not summary["wiring"].get("current"):
        raise ValueError("assign the current client model (step 3) and rebuild to roll forward")
    if mode == "prior" and not summary["wiring"].get("prior"):
        raise ValueError("assign the prior client model (step 3) to feed from it")
    clean = {}
    for cell, v in (changes or {}).items():
        if isinstance(v, str) and re.match(r"^\d{4}-\d{2}-\d{2}$", v):
            v = ovmod.serial(ovmod.date.fromisoformat(v))
        elif isinstance(v, str):
            v = float(v.replace(",", ""))
        clean[cell] = v
    return ovmod.deep(ovmod.scenario, sess, summary, mode, clean, valuation_date, months)


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
        doc = docingest.process(*args)
    except Exception as ex:  # not signed in: keep the text layer, leave the tables to read on retry
        if type(ex).__name__ not in ("AuthenticationRequiredError", "ClientAuthenticationError"):
            raise
        doc = docingest.process(*args, read=False)
        note = "Tables not read. " + friendly(ex)
    _save_doc(did, doc)
    errors = [t for t in doc["tables"] if t.get("status") == "error"]
    if errors:
        note = f"{len(errors)} table(s) couldn't be read: {errors[0].get('error')}"
    _set("documents", did, status="done", step="Done", pct=1.0, processed_at=time.time(), error=note)


def _process_facts(did: int) -> None:
    d = _doc(did)
    eid = d["engagement_id"]
    e = _q("SELECT model, reviewer_model FROM engagements WHERE id=?", eid)[0]
    doc = json.loads(d["doc_json"])
    _set("documents", did, facts_status="running", facts_step="Extracting key facts")
    res = reportfacts.run(doc["markdown"], e["model"] or DEFAULT_MODEL, e["reviewer_model"] or DEFAULT_REVIEWER, _logger(eid),
                          lambda frac, msg: _set("documents", did, facts_step=msg))
    now = time.time()
    with _lock, _conn() as db:
        db.execute("DELETE FROM facts WHERE document_id=?", (did,))
        for f in res["facts"]:
            db.execute(f"""INSERT INTO facts(engagement_id, document_id, n, {', '.join(FACT_FIELDS)}, origin, check_json,
                           review_json, status, updated_at) VALUES ({', '.join('?' * (len(FACT_FIELDS) + 8))})""",
                       (eid, did, f["id"], *[f.get(k) for k in FACT_FIELDS], f["origin"], json.dumps(f["check"]),
                        json.dumps(f.get("review")), "pending", now))
    _set("documents", did, facts_status="done", facts_step="Done", facts_notes=res.get("notes"),
         review_summary=res.get("review_summary"))
    _touch(eid)


def _run(kind: str, rid: int) -> None:
    try:
        {"doc": _process_doc, "facts": _process_facts, "compare": _compare, "map": _map, "overlay": _overlay}[kind](rid)
    except Exception as e:
        traceback.print_exc()
        msg = friendly(e)
        if kind == "doc":
            _set("documents", rid, status="error", step="Failed", error=msg)
        elif kind == "facts":
            _set("documents", rid, facts_status="error", facts_step="Failed", facts_error=msg)
        else:
            _set("engagements", rid, **{f"{kind}_status": "error", f"{kind}_step": "Failed", f"{kind}_error": msg})


def _worker() -> None:
    while True:
        kind, rid = _jobs.get()
        try:
            _run(kind, rid)
        finally:
            _jobs.task_done()


def retry_document(did: int) -> None:
    _set("documents", did, status="queued", step="Waiting to start", pct=0, error=None)
    _jobs.put(("doc", did))


def start_worker() -> None:
    """Re-queue anything a restart interrupted, then start the worker."""
    for r in _q("SELECT id FROM documents WHERE status IN ('queued','processing')"):
        _jobs.put(("doc", r["id"]))
    for r in _q("SELECT id FROM documents WHERE facts_status IN ('queued','running')"):
        _jobs.put(("facts", r["id"]))
    for kind in ("compare", "map", "overlay"):
        for r in _q(f"SELECT id FROM engagements WHERE {kind}_status IN ('queued','running')"):
            _jobs.put((kind, r["id"]))
    threading.Thread(target=_worker, daemon=True, name="engagement-worker").start()

