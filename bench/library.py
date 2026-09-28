"""Registry of uploaded workbooks: fingerprint, process once, identify, and compare with the last version.

Each file is keyed by the SHA-256 of its bytes, so the same file uploaded again (under any name) is not
processed twice. Each version gets its own folder, out/<stem>__<sha8>/, so older versions stay available
to compare against. Processing runs on one background worker thread, one file at a time.
"""
import hashlib
import json
import queue
import re
import shutil
import sqlite3
import threading
import time
import traceback
from pathlib import Path

import build_map
import diff as diffmod
import identify as identmod
import usage

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "out"
UPLOADS = ROOT / "uploads"
REGISTRY = OUT / "registry.db"
SUMMARY_MODEL = "gpt-4o"
SUPPORTED = (".xlsx", ".xlsm")

_lock = threading.Lock()
_jobs: "queue.Queue[int]" = queue.Queue()

SCHEMA = """
CREATE TABLE IF NOT EXISTS files(
  id INTEGER PRIMARY KEY, sha256 TEXT UNIQUE, filename TEXT, size INT, uploaded_at REAL, source_path TEXT,
  status TEXT, step TEXT, pct REAL, error TEXT, out_dir TEXT, db_path TEXT, processed_at REAL,
  sheets INT, line_items INT, build_secs REAL,
  target_name TEXT, project_name TEXT, valuation_date TEXT, identity_json TEXT, identity_confirmed INT DEFAULT 0,
  previous_id INT, diff_json TEXT, diff_summary TEXT);
"""
LIST_COLS = ("id, sha256, filename, size, uploaded_at, status, step, pct, error, processed_at, started_at, sheets, line_items, "
             "build_secs, target_name, project_name, valuation_date, identity_confirmed, previous_id, "
             "json_array_length(diff_json, '$.warnings') AS n_warnings")


def _conn() -> sqlite3.Connection:
    OUT.mkdir(exist_ok=True)
    db = sqlite3.connect(REGISTRY, check_same_thread=False, timeout=30)
    db.row_factory = sqlite3.Row
    db.executescript(SCHEMA)
    if "started_at" not in {r[1] for r in db.execute("PRAGMA table_info(files)")}:  # registries from before timings
        db.execute("ALTER TABLE files ADD COLUMN started_at REAL")
    return db


def _update(fid: int, **fields) -> None:
    with _lock, _conn() as db:
        db.execute(f"UPDATE files SET {', '.join(f'{k}=?' for k in fields)} WHERE id=?", (*fields.values(), fid))


def get(fid: int, full: bool = False) -> dict | None:
    with _conn() as db:
        r = db.execute(f"SELECT {'*' if full else LIST_COLS} FROM files WHERE id=?", (fid,)).fetchone()
    if not r:
        return None
    d = dict(r)
    if full:
        for k in ("identity_json", "diff_json"):
            d[k.removesuffix("_json")] = json.loads(d.pop(k) or "null")
    return d


def all_files() -> list[dict]:
    with _conn() as db:
        return [dict(r) for r in db.execute(f"SELECT {LIST_COLS} FROM files ORDER BY uploaded_at DESC")]


def by_sha(sha: str) -> dict | None:
    with _conn() as db:
        r = db.execute("SELECT id FROM files WHERE sha256=?", (sha,)).fetchone()
    return get(r["id"]) if r else None


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def add_upload(tmp_path: Path, filename: str, sha: str) -> tuple[str, dict]:
    """Register an uploaded file. Returns ("duplicate", existing) or ("queued", new record)."""
    existing = by_sha(sha)
    if existing:
        tmp_path.unlink(missing_ok=True)
        return "duplicate", existing
    if not filename.lower().endswith(SUPPORTED):
        tmp_path.unlink(missing_ok=True)
        raise ValueError(f"{filename}: only .xlsx and .xlsm can be read (formulas can't be read from .xlsb/.xls). "
                         "Save a copy as .xlsm from Excel and upload that.")
    dest = UPLOADS / sha[:12] / Path(filename).name
    dest.parent.mkdir(parents=True, exist_ok=True)
    tmp_path.replace(dest)
    return "queued", _register(dest, filename, sha)


def _register(path: Path, filename: str, sha: str, out_dir: Path | None = None) -> dict:
    out_dir = out_dir or OUT / f"{Path(filename).stem}__{sha[:8]}"
    with _lock, _conn() as db:
        cur = db.execute("""INSERT INTO files(sha256, filename, size, uploaded_at, source_path, status, step, pct,
                            out_dir, db_path) VALUES (?,?,?,?,?,?,?,?,?,?)""",
                         (sha, filename, path.stat().st_size, time.time(), str(path), "queued", "Waiting to start",
                          0, str(out_dir), str(out_dir / "model.db")))
        fid = cur.lastrowid
    _jobs.put(fid)
    return get(fid)


# ---- versions ---------------------------------------------------------------------------------------

_VERSION_BITS = re.compile(r"^copy of\b|\b(v\d+[a-z]*|vsent|final|draft|updated?|rev\d*|clean)\b|"
                           r"\b\d{4} \d{2} \d{2}\b|\b\d{6,8}\b|\(\d+\)", re.I)


def family(filename: str) -> str:
    """File name with dates, 'Copy of', v2/final/draft etc. removed: '200401 Acme Valuation model_vSent' -> 'acme valuation model'."""
    s = Path(filename).stem.replace("_", " ").replace("-", " ")
    s = _VERSION_BITS.sub(" ", s)
    return re.sub(r"\s+", " ", s).strip().lower()


def _norm(s: str | None) -> str:
    return re.sub(r"[^a-z0-9]+", " ", (s or "").lower()).strip()


def find_previous(fid: int) -> dict | None:
    """The version just before this one: same target (or, failing that, same file family), ordered by valuation
    date when both files have one, otherwise by upload time."""
    me = get(fid)
    others = [f for f in all_files() if f["id"] != fid and f["status"] == "done"]
    for key in (lambda f: _norm(f["target_name"]), lambda f: family(f["filename"])):
        mine = key(me)
        hits = [f for f in others if mine and key(f) == mine]
        if not hits:
            continue
        vd = me["valuation_date"]
        dated = [f for f in hits if f["valuation_date"]]
        if vd and dated:
            earlier = [f for f in dated if (f["valuation_date"], f["uploaded_at"]) < (vd, me["uploaded_at"])]
            return max(earlier, key=lambda f: (f["valuation_date"], f["uploaded_at"])) if earlier else None
        earlier = [f for f in hits if f["uploaded_at"] < me["uploaded_at"]]
        return max(earlier, key=lambda f: f["uploaded_at"]) if earlier else None
    return None


def _relink_later(fid: int) -> None:
    """An older version uploaded after a newer one: point the newer one at it instead."""
    for f in all_files():
        if f["id"] != fid and f["status"] == "done":
            prev = find_previous(f["id"])
            if prev and prev["id"] == fid and f["previous_id"] != fid:
                compare(f["id"], fid)


def remove(fid: int) -> None:
    """Forget a file: registry row, its model.db folder and the uploaded copy. Later versions get re-linked."""
    f = get(fid, full=True)
    if not f:
        return
    if f["status"] in ("queued", "processing"):
        raise ValueError("wait for processing to finish before removing this file")
    dependents = [x["id"] for x in all_files() if x["previous_id"] == fid]
    with _lock, _conn() as db:
        db.execute("DELETE FROM files WHERE id=?", (fid,))
    out_dir, src = Path(f["out_dir"]), Path(f["source_path"])
    if out_dir.is_relative_to(OUT) and out_dir.name.endswith(f"__{f['sha256'][:8]}"):
        shutil.rmtree(out_dir, ignore_errors=True)  # never delete folders built before the registry
    if src.is_relative_to(UPLOADS):
        shutil.rmtree(src.parent, ignore_errors=True)
    for d in dependents:
        prev = find_previous(d)
        compare(d, prev["id"] if prev else None)


# ---- processing ---------------------------------------------------------------------------------------

def _stage(fid: int, lo: float, hi: float):
    return lambda frac, msg: _update(fid, pct=round(lo + (hi - lo) * frac, 3), step=msg)


def compare(fid: int, prev_id: int | None) -> None:
    """(Re)compute changes of fid against prev_id and store them with a short written summary."""
    if prev_id is None:
        _update(fid, previous_id=None, diff_json=None, diff_summary=None)
        return
    me, prev = get(fid, full=True), get(prev_id, full=True)
    d = diffmod.diff(prev["db_path"], me["db_path"])
    d["identity"] = {k: {"old": prev[k], "new": me[k]} for k in ("target_name", "project_name", "valuation_date")
                     if (prev[k] or None) != (me[k] or None)}
    d["previous"] = {"id": prev_id, "filename": prev["filename"]}
    d["warnings"] = _warnings(d, me)
    _update(fid, previous_id=prev_id, diff_json=json.dumps(d, default=str), diff_summary=None)
    _update(fid, diff_summary=summarize(d, prev, me))


def _warnings(d: dict, me: dict) -> list[str]:
    """Signs the new file was saved without recalculating, so its cached results may be stale."""
    out, c = [], d["counts"]
    if c.get("calculated_now_blank", 0) >= max(5, c["calculated_cells_changed"] // 2):
        out.append(f"{c['calculated_now_blank']:,} formula cells have no saved result in the new file (they had one before). "
                   "It was probably written by a script or tool that doesn't calculate, so values read from it are "
                   "blank. Open it in Excel, recalculate, save, and upload again.")
    elif c["inputs_changed"] and not c["calculated_cells_changed"]:
        out.append(f"{c['inputs_changed']} input(s) changed but no calculated value moved. The workbook was probably "
                   "saved without recalculating (manual calculation mode), so its results may be stale. "
                   "Open it in Excel, press F9 or Ctrl+Alt+F9, save, and upload again.")
    for i in d["inputs"]:
        if re.search(r"valuation date|val\.? date", i["label"] or "", re.I) and me["valuation_date"] \
                and str(i["new"])[:10] != me["valuation_date"]:
            out.append(f"The valuation date input {i['ref']} is now {str(i['new'])[:10]}, but the workbook's "
                       f"calculated valuation date is still {me['valuation_date']}.")
    return out


SUMMARY_PROMPT = """Two versions of an Excel valuation model were compared cell by cell. Write the change summary
an analyst reads first: 3-6 short bullet points in Markdown, most important first. Start directly with the first
bullet: no heading or title. If there are warnings, the first bullet states them, and don't describe blank
values as "recalculated". Lead with any change to the
valuation date or target, then input changes and their effect on key outputs, then structural changes (rows,
formulas, sheets). Quote numbers old -> new with units, and cite cells as Sheet!A1. If the content is identical,
say the file was re-saved with no changes. Don't speculate beyond the data.

Previous: {old}
New: {new}
Comparison (lists truncated):
{d}"""


def summarize(d: dict, prev: dict, me: dict, model: str = SUMMARY_MODEL, session: str | None = None) -> str:
    if d["identical_content"] and not d.get("identity"):
        return "No content changes: same sheets, line items, formulas and values. The file was re-saved."
    brief = {k: (v[:25] if isinstance(v, list) else v) for k, v in d.items()}
    try:
        from llm import client, create
        r = create(client(interactive=False), model, input=SUMMARY_PROMPT.format(
            old=prev["filename"], new=me["filename"], d=json.dumps(brief, default=str)[:24000]))
        if r.usage:
            usage.record(model, r.usage, "summary", me["id"], session)
        return r.output_text
    except Exception as e:
        c = d["counts"]
        return (f"- {c['inputs_changed']} inputs changed, {c['calculated_cells_changed']} calculated cells moved, "
                f"{c['formula_rows_changed']} formula rows changed, {c['line_items_added']} line items added, "
                f"{c['line_items_removed']} removed.\n- (Written summary unavailable: {type(e).__name__}.)")


def process(fid: int) -> None:
    f = get(fid, full=True)
    try:
        _update(fid, status="processing", step="Opening workbook", pct=0, error=None, started_at=time.time())
        stats = build_map.main(f["source_path"], f["out_dir"], _stage(fid, 0.0, 0.8))
        _update(fid, sheets=stats["sheets"], line_items=stats["line_items"], build_secs=stats["secs"])
        _update(fid, step="Reading external links", pct=0.81)
        import extlinks  # now, while nothing else reads this model.db: later they'd be written under readers' feet
        extlinks.build(f["source_path"], f["db_path"])

        _update(fid, step="Identifying target and valuation date", pct=0.82)
        ident = identmod.identify(f["db_path"], f["filename"], file_id=fid)
        _update(fid, target_name=ident.get("target_name"), project_name=ident.get("project_name"),
                valuation_date=ident.get("valuation_date"), identity_json=json.dumps(ident, default=str))

        _update(fid, step="Looking for an earlier version", pct=0.9)
        prev = find_previous(fid)
        if prev:
            _update(fid, step=f"Comparing with {prev['filename']}", pct=0.92)
        compare(fid, prev["id"] if prev else None)
        _update(fid, status="done", step="Done", pct=1.0, processed_at=time.time())
        _relink_later(fid)
    except Exception as e:
        traceback.print_exc()
        _update(fid, status="error", step="Failed", error=f"{type(e).__name__}: {e}")


def rebuild(fid: int) -> dict:
    """Process a file again from scratch, done or failed (its model.db is rebuilt for everyone using it)."""
    f = get(fid)
    if not f:
        raise ValueError("no such file")
    if f["status"] in ("queued", "processing"):
        raise ValueError(f"{f['filename']} is already being processed")
    _update(fid, status="queued", step="Waiting to start", pct=0, error=None)
    _jobs.put(fid)
    return get(fid)


def retry(fid: int) -> dict:
    """Process a failed file again (uploading the same bytes again would only find the failed record)."""
    f = get(fid)
    if not f:
        raise ValueError("no such file")
    if f["status"] != "error":
        raise ValueError("only a file that failed can be retried")
    _update(fid, status="queued", step="Waiting to start", pct=0, error=None)
    _jobs.put(fid)
    return get(fid)


def set_identity(fid: int, target_name: str | None, project_name: str | None, valuation_date: str | None) -> dict:
    """User confirmation / correction. Target and date decide which file is the previous version, so re-link."""
    _update(fid, target_name=target_name or None, project_name=project_name or None,
            valuation_date=valuation_date or None, identity_confirmed=1)
    prev = find_previous(fid)
    compare(fid, prev["id"] if prev else None)
    _relink_later(fid)
    return get(fid, full=True)


def _worker() -> None:
    while True:
        fid = _jobs.get()
        try:
            process(fid)
        finally:
            _jobs.task_done()


def start() -> None:
    """Start the worker; re-queue files interrupted by a restart; register workbooks built before the registry."""
    import ratelimit
    threading.Thread(target=ratelimit.load_capacities, daemon=True, name="rate-limits").start()
    with _lock, _conn() as db:
        stuck = [r["id"] for r in db.execute("SELECT id FROM files WHERE status IN ('queued','processing')")]
    for fid in stuck:
        _jobs.put(fid)
    threading.Thread(target=_worker, daemon=True, name="library-worker").start()
    threading.Thread(target=_adopt_existing, daemon=True, name="library-adopt").start()


def _adopt_existing() -> None:
    """Workbooks in reference/ and tests/ whose out/<stem>/model.db already exists: register without rebuilding."""
    for src in sorted([*ROOT.glob("reference/*.xls[xm]"), *ROOT.glob("tests/*.xlsx")]):
        out_dir = OUT / src.stem
        if not (out_dir / "model.db").exists():
            continue
        sha = sha256_file(src)
        if by_sha(sha):
            continue
        with _lock, _conn() as db:
            cur = db.execute("""INSERT INTO files(sha256, filename, size, uploaded_at, source_path, status, step, pct,
                                out_dir, db_path, processed_at) VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
                             (sha, src.name, src.stat().st_size, src.stat().st_mtime, str(src), "processing",
                              "Identifying target and valuation date", 0.85, str(out_dir), str(out_dir / "model.db"),
                              time.time()))
            fid = cur.lastrowid
        try:
            ident = identmod.identify(str(out_dir / "model.db"), src.name, file_id=fid)
            _update(fid, target_name=ident.get("target_name"), project_name=ident.get("project_name"),
                    valuation_date=ident.get("valuation_date"), identity_json=json.dumps(ident, default=str))
            with sqlite3.connect(out_dir / "model.db") as m:
                _update(fid, sheets=m.execute("SELECT COUNT(*) FROM sheets").fetchone()[0],
                        line_items=m.execute("SELECT COUNT(*) FROM rows").fetchone()[0])
            _update(fid, status="done", step="Done", pct=1.0)
        except Exception as e:
            _update(fid, status="error", step="Failed", error=f"{type(e).__name__}: {e}")
