"""Valuation Desk: a recurring valuation engagement, from last year's four files to an approved map.
    uv run uvicorn engage.server:app --port 8002      then open http://localhost:8002

Upload the prior report (PDF / PPTX), the prior client model, the prior overlay and the current client model;
check the report's tables and key facts, confirm who's who, compare the client models and build the map.
Workbooks share the Model Desk library (app/server.py), so a model uploaded in either app is processed once.
"""
import os
import sys
import tempfile
from contextlib import asynccontextmanager
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse
from pydantic import BaseModel
from starlette.concurrency import run_in_threadpool

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "bench"))
import dictation  # noqa: E402
import engagement  # noqa: E402
import library  # noqa: E402
import usage  # noqa: E402
import xlruntime  # noqa: E402

# Model deployments offered. gpt-6-luna extracts and reads tables; gpt-6-sol reviews (second reads, fact review).
MODELS = ["gpt-6-luna", "gpt-6-sol", "gpt-4o", "gpt-4o-mini", "gpt-5-nano"]


@asynccontextmanager
async def lifespan(app):
    library.start()  # builds uploaded workbooks (shared with Model Desk)
    engagement.start_worker()
    yield


app = FastAPI(lifespan=lifespan)


def _run(fn, *args, **kw):
    """Run in a worker thread; ValueError -> 400, missing -> 404."""
    async def go():
        try:
            out = await run_in_threadpool(lambda: fn(*args, **kw))
        except ValueError as e:
            raise HTTPException(400, str(e))
        if out is None:
            raise HTTPException(404, "not found")
        return out
    return go()


@app.get("/")
def index():
    return FileResponse(Path(__file__).parent / "index.html")


@app.get("/charts.js")
def charts_js():  # the chart drawing shared with the other app (web/charts.js)
    return FileResponse(Path(__file__).resolve().parent.parent / "web" / "charts.js", media_type="text/javascript")


# ---- the logo in the header's top-left corner: your firm's, from the git-ignored brand/ folder, else the mascot ----
BRAND = ROOT / "brand"
LOGO_TYPES = {".svg": "image/svg+xml", ".png": "image/png", ".webp": "image/webp", ".jpg": "image/jpeg", ".jpeg": "image/jpeg"}
LOGO_MAX = 2 * 1024 * 1024
LOGO_SPEC = {"folder": "brand", "names": ["logo.svg", "logo.png"], "also": ["logo.webp", "logo.jpg"],
             "height_px": 144, "shown_px": 36, "max_width_px": 640, "max_bytes": LOGO_MAX,
             "note": "transparent background, light artwork (the header is dark)"}


def _logo() -> Path | None:
    return next((BRAND / f"logo{ext}" for ext in LOGO_TYPES if (BRAND / f"logo{ext}").exists()), None)


@app.get("/brand/logo")
def brand_logo():
    p = _logo()
    headers = {"Cache-Control": "no-cache", "Content-Security-Policy": "script-src 'none'"}  # an SVG never runs scripts
    if p:
        return FileResponse(p, media_type=LOGO_TYPES[p.suffix.lower()], headers=headers)
    return FileResponse(Path(__file__).parent / "mascot.svg", media_type="image/svg+xml", headers=headers)


@app.get("/api/brand")
def brand_info():
    p = _logo()
    return {"custom": bool(p), "file": f"brand/{p.name}" if p else None, "spec": LOGO_SPEC}


@app.post("/api/brand/logo")
async def brand_upload(file: UploadFile = File(...)):
    """Your firm's logo for the header: SVG, PNG, WebP or JPEG, kept in brand/ (never committed)."""
    ext = Path(file.filename or "").suffix.lower()
    data = await file.read(LOGO_MAX + 1)
    if ext not in LOGO_TYPES:
        raise HTTPException(400, "upload the logo as .svg or .png (or .webp / .jpg)")
    if len(data) > LOGO_MAX:
        raise HTTPException(400, "the logo must be 2 MB or less")
    ok = {".png": data[:8] == b"\x89PNG\r\n\x1a\n", ".jpg": data[:3] == b"\xff\xd8\xff", ".jpeg": data[:3] == b"\xff\xd8\xff",
          ".webp": data[:4] == b"RIFF" and data[8:12] == b"WEBP", ".svg": b"<svg" in data[:4096].lower()}[ext]
    if not ok:
        raise HTTPException(400, f"that file isn't a {ext[1:].upper()} image")
    BRAND.mkdir(exist_ok=True)
    for other in LOGO_TYPES:
        (BRAND / f"logo{other}").unlink(missing_ok=True)
    (BRAND / f"logo{'.jpg' if ext == '.jpeg' else ext}").write_bytes(data)
    return brand_info()


@app.delete("/api/brand/logo")
def brand_reset():
    for ext in LOGO_TYPES:
        (BRAND / f"logo{ext}").unlink(missing_ok=True)
    return brand_info()


@app.get("/api/config")
def config():
    return {"models": MODELS, "default_model": engagement.DEFAULT_MODEL, "default_reviewer": engagement.DEFAULT_REVIEWER,
            "default_arbiter": engagement.DEFAULT_ARBITER,
            "roles": engagement.rolesmod.ROLES, "dictation": dictation.status(), "runtime": xlruntime.RUNTIME}


class NewEngagement(BaseModel):
    name: str


class EngagementPatch(BaseModel):
    name: str | None = None
    model: str | None = None
    reviewer_model: str | None = None
    arbiter_model: str | None = None


@app.get("/api/engagements")
def list_engagements():
    return engagement.all_engagements()


@app.post("/api/engagements")
async def new_engagement(body: NewEngagement):
    return await _run(engagement.create, body.name)


@app.get("/api/engagements/{eid}")
async def get_engagement(eid: int):
    return await _run(engagement.get, eid)


@app.patch("/api/engagements/{eid}")
async def patch_engagement(eid: int, body: EngagementPatch):
    for m in (body.model, body.reviewer_model, body.arbiter_model):
        if m and m not in MODELS:
            raise HTTPException(400, f"unknown model {m}")
    return await _run(engagement.update, eid, **body.model_dump())


@app.delete("/api/engagements/{eid}")
async def delete_engagement(eid: int):
    await _run(lambda: engagement.delete(eid) or True)
    return {"ok": True}


@app.post("/api/engagements/{eid}/files")
async def upload(eid: int, file: UploadFile = File(...)):
    if not engagement.get(eid):
        raise HTTPException(404, "no such engagement")
    library.UPLOADS.mkdir(exist_ok=True)
    fd, tmp_name = tempfile.mkstemp(dir=library.UPLOADS, suffix=".part")
    os.close(fd)
    tmp = Path(tmp_name)
    h = library.hashlib.sha256()
    with open(tmp, "wb") as out:
        while chunk := await file.read(1 << 20):
            h.update(chunk)
            out.write(chunk)
    return await _run(engagement.add_upload, eid, tmp, file.filename, h.hexdigest())


@app.delete("/api/engagements/{eid}/workbooks/{fid}")
async def unlink_workbook(eid: int, fid: int):
    await _run(lambda: engagement.remove_workbook(eid, fid) or True)
    return {"ok": True}


class DateCheck(BaseModel):
    valuation_date: str


@app.post("/api/engagements/{eid}/workbooks/{fid}/date")
async def confirm_date(eid: int, fid: int, body: DateCheck):
    """Your check of a workbook's valuation date: the roll-forward runs between last year's and this year's."""
    return await _run(engagement.confirm_date, eid, fid, body.valuation_date)


@app.post("/api/engagements/{eid}/workbooks/{fid}/retry")
async def retry_workbook(eid: int, fid: int):
    await _run(lambda: engagement.retry_workbook(eid, fid) and True)
    return {"ok": True}


@app.post("/api/engagements/{eid}/workbooks/{fid}/rebuild")
async def rebuild_workbook(eid: int, fid: int):
    """Process a workbook again from scratch (shared: every engagement using it gets the new build)."""
    await _run(lambda: engagement.rebuild_workbook(eid, fid) and True)
    return {"ok": True}


@app.post("/api/documents/{did}/rebuild")
async def rebuild_document(did: int):
    """Read a report again from scratch; its key facts are kept and re-checked."""
    await _run(lambda: engagement.rebuild_document(did) and True)
    return {"ok": True}


@app.get("/api/documents/{did}")
async def get_document(did: int):
    return await _run(engagement.document, did)


@app.delete("/api/documents/{did}")
async def delete_document(did: int):
    await _run(lambda: engagement.remove_document(did) or True)
    return {"ok": True}


@app.post("/api/documents/{did}/retry")
async def retry_document(did: int):
    await _run(lambda: engagement.retry_document(did) or True)
    return {"ok": True}


@app.get("/api/documents/{did}/tables/{tid}.png")
def table_image(did: int, tid: str):
    p = engagement.table_png(did, tid)
    if not p:
        raise HTTPException(404, "no such table image")
    return FileResponse(p, media_type="image/png")


class TableDecision(BaseModel):
    action: str  # approve | use_second | edit | reset
    markdown: str | None = None


@app.put("/api/documents/{did}/tables/{tid}")
async def put_table(did: int, tid: str, body: TableDecision):
    return await _run(engagement.settle_table, did, tid, body.action, body.markdown)


@app.post("/api/documents/{did}/facts")
async def run_facts(did: int):
    await _run(lambda: engagement.extract_facts(did) or True)
    return {"ok": True}


@app.post("/api/documents/{did}/tables/resolve")
async def resolve_tables(did: int):
    """Run the review and remediation loop on the document's flagged tables."""
    await _run(lambda: engagement.resolve_tables(did) or True)
    return {"ok": True}


@app.post("/api/documents/{did}/facts/resolve")
async def resolve_facts(did: int):
    """Run the review and remediation loop on the facts the agents haven't settled."""
    await _run(lambda: engagement.resolve_facts(did) or True)
    return {"ok": True}


@app.get("/api/engagements/{eid}/calls")
async def get_calls(eid: int):
    """Tokens, cost and time by file and by step, and the latest calls (the call log, calllog.py)."""
    return await _run(engagement.calls_view, eid)


@app.get("/api/engagements/{eid}/calls/{cid}")
async def get_call(eid: int, cid: int):
    """One call in full: what was sent and what came back."""
    return await _run(engagement.call_view, eid, cid)


@app.get("/api/lessons")
async def get_lessons():
    """What the report agents have learned: curated rules (docs/report_rules.md) and learned lessons."""
    return await _run(engagement.lessons_view)


class LessonChange(BaseModel):
    action: str  # retire | restore | promote


@app.put("/api/lessons/{lid}")
async def put_lesson(lid: str, body: LessonChange):
    import lessons
    if body.action == "promote":
        await _run(lessons.promote, lid)
    elif body.action in ("retire", "restore"):
        await _run(lessons.set_status, lid, "retired" if body.action == "retire" else "active")
    else:
        raise HTTPException(400, "action must be retire, restore or promote")
    return await _run(engagement.lessons_view)


class FactDecision(BaseModel):
    action: str  # approve | use_suggestion | edit | reject | reset
    fields: dict | None = None


@app.put("/api/facts/{fact_id}")
async def put_fact(fact_id: int, body: FactDecision):
    return await _run(engagement.set_fact, fact_id, body.action, body.fields)


@app.post("/api/engagements/{eid}/facts/approve-passed")
async def approve_passed(eid: int):
    return {"approved": await _run(engagement.approve_passed, eid)}


@app.post("/api/engagements/{eid}/roles/suggest")
async def suggest_roles(eid: int):
    return await _run(engagement.suggest_roles, eid)


class Roles(BaseModel):
    roles: dict


@app.put("/api/engagements/{eid}/roles")
async def put_roles(eid: int, body: Roles):
    return await _run(engagement.confirm_roles, eid, body.roles)


class Profile(BaseModel):
    fields: dict  # {"fy_end_month": 6 | None, "horizon": "fixed" | "rolling" | None}


@app.get("/api/engagements/{eid}/profile")
async def get_profile(eid: int):
    """The engagement's profile: financial-year end, horizon, periods, units, discounting; detected, and set."""
    return await _run(engagement.profile_view, eid)


@app.put("/api/engagements/{eid}/profile")
async def put_profile(eid: int, body: Profile):
    return await _run(engagement.set_profile, eid, body.fields)


class Schedule(BaseModel):
    classes: dict | None = None   # {"Sheet!r12": "conclusion" | "assumption" | "working" | None}
    outside: dict | None = None   # {fact key: True | False}: produced outside the model
    confirm: bool = False


@app.get("/api/engagements/{eid}/schedule")
async def get_schedule(eid: int):
    """The overlay's own outputs (outputs.py), classified, with the report's figures that sit on none."""
    return await _run(engagement.schedule_view, eid)


@app.put("/api/engagements/{eid}/schedule")
async def put_schedule(eid: int, body: Schedule):
    return await _run(engagement.set_schedule, eid, body.classes, body.outside, body.confirm)


@app.get("/api/engagements/{eid}/models/{fid}")
async def model_dashboard(eid: int, fid: int):
    """The Map step's dashboard for one of the engagement's workbooks."""
    return await _run(engagement.model_dashboard, eid, fid)


@app.get("/api/engagements/{eid}/models/{fid}/rows")
async def model_rows(eid: int, fid: int, sheet: str | None = None, q: str | None = None, mapped: bool = False,
                     limit: int = 200):
    return await _run(engagement.model_rows, eid, fid, sheet, q, mapped, limit)


class Summary(BaseModel):
    changes: dict = {}
    valuation_date: str | None = None
    months: int | None = None
    method: dict = {}


@app.post("/api/engagements/{eid}/summary")
async def summary_view(eid: int, body: Summary):
    """The Summary page: the report's summary table rebuilt, rolled forward and as a scenario, and its charts."""
    return await _run(engagement.summary_view, eid, body.changes, body.valuation_date, body.months, body.method)


@app.post("/api/engagements/{eid}/bridge")
async def bridge_view(eid: int, body: Summary):
    """Last year's value to this year's, step by step."""
    return await _run(engagement.bridge_view, eid, body.changes, body.valuation_date, body.months, body.method)


@app.post("/api/engagements/{eid}/charts")
async def recreate_charts(eid: int):
    """Recreate the report's charts from the models again."""
    await _run(engagement.recreate_charts, eid)
    return {"ok": True}


class ChartPick(BaseModel):
    series: int
    rows: list[dict] = []


@app.post("/api/engagements/{eid}/charts/{cid}/pick")
async def chart_pick(eid: int, cid: str, body: ChartPick):
    """Your rows for one series of a report chart: redrawn from them and checked against the reading."""
    return await _run(engagement.chart_pick, eid, cid, body.series, body.rows)


@app.get("/api/engagements/{eid}/charts/{name}")
async def chart_png(eid: int, name: str):
    p = engagement.chart_png(eid, name)
    if not p:
        raise HTTPException(404, "no such chart image")
    return FileResponse(p, media_type="image/png")


@app.get("/api/engagements/{eid}/overlay/facts")
async def overlay_facts(eid: int, start: str | None = None):
    """The facts behind a report figure: its discounting, the path up to it, and where this year's model has
    the rows its cash flows come from."""
    return await _run(engagement.overlay_facts, eid, start)


class RowPick(BaseModel):
    prior: str
    current: str | None = None


@app.post("/api/engagements/{eid}/overlay/rowpick")
async def row_pick(eid: int, body: RowPick):
    """Your choice of this year's row for one of last year's rows (none: back to what was found)."""
    return await _run(engagement.row_pick, eid, body.prior, body.current)


@app.get("/api/engagements/{eid}/overlay/row")
async def overlay_row(eid: int, row: str):
    """What this year's model has for one of last year's rows (Sheet!rN), with the alternatives."""
    return await _run(engagement.row_info, eid, row)


@app.get("/api/engagements/{eid}/rows")
async def rows_view(eid: int):
    """The row agents: their status and what they decided for each row the Summary was waiting on."""
    return await _run(engagement.rows_view, eid)


@app.get("/api/engagements/{eid}/doctor")
async def doctor_view(eid: int):
    """The overlay doctor's last diagnosis, its progress, and the cells held at Excel's values."""
    return await _run(engagement.doctor_view, eid)


@app.post("/api/engagements/{eid}/doctor")
async def start_doctor(eid: int):
    """Run the doctor: where each figure breaks, why, and whether the files and rows are the right ones."""
    return await _run(engagement.start_doctor, eid)


class Holds(BaseModel):
    cells: list[str] | None = None
    release: bool = False


@app.post("/api/engagements/{eid}/doctor/holds")
async def doctor_holds(eid: int, body: Holds):
    """Hold the doctor's safe cells at Excel's saved value on every feed, or release them all."""
    return await _run(engagement.doctor_holds, eid, body.cells, body.release)


@app.post("/api/engagements/{eid}/{kind}")
async def start_step(eid: int, kind: str):
    if kind not in ("compare", "map", "overlay", "rows"):
        raise HTTPException(404, "unknown step")
    return await _run(engagement.start, kind, eid)


class OverlayRun(BaseModel):
    mode: str = "workbook"  # workbook | prior | current
    changes: dict = {}      # {"Sheet!A1": value}
    valuation_date: str | None = None
    months: int | None = None


@app.post("/api/engagements/{eid}/overlay/run")
async def overlay_run(eid: int, body: OverlayRun):
    if body.mode not in ("workbook", "prior", "current"):
        raise HTTPException(400, "mode must be workbook, prior or current")
    return await _run(engagement.overlay_run, eid, body.mode, body.changes, body.valuation_date, body.months)


class OverlayDcf(BaseModel):
    mode: str = "workbook"
    changes: dict = {}
    valuation_date: str | None = None
    months: int | None = None
    cell: str | None = None
    rate: float | None = None
    dcf_valuation_date: str | None = None
    timing: str | None = None
    day_count: str | None = None
    cutoff: str | None = "model"
    include: list[bool] | None = None
    low: float | None = None
    high: float | None = None


@app.get("/api/engagements/{eid}/overlay/valuation")
async def overlay_valuation(eid: int, cell: str | None = None):
    """The Model Desk's Validate step on the overlay's saved values."""
    return await _run(engagement.overlay_valuation, eid, cell)


@app.get("/api/engagements/{eid}/overlay/value-trace")
async def overlay_value_trace(eid: int, start: str | None = None):
    """How a report figure is built in the overlay, traced down to its discounting."""
    return await _run(engagement.overlay_value_trace, eid, start)


@app.post("/api/engagements/{eid}/overlay/dcf")
async def overlay_dcf(eid: int, body: OverlayDcf):
    """A DCF on the live module (feed + changes), checked against the module, and under another method."""
    b = body.model_dump()
    return await _run(engagement.overlay_dcf, eid, b.pop("mode"), b.pop("changes"), b.pop("valuation_date"),
                      b.pop("months"), **b)


class ThisYearDate(BaseModel):
    valuation_date: str | None = None


@app.post("/api/engagements/{eid}/roll/this_year_date")
async def this_year_date(eid: int, body: ThisYearDate):
    """This year's valuation date for the engagement (None clears it): the roll-forward runs to it."""
    return await _run(engagement.set_this_year_date, eid, body.valuation_date)


class Dictated(BaseModel):
    seconds: float


@app.get("/api/engagements/{eid}/dictation")
async def dictation_start(eid: int):  # a Speech token, the language and the engagement's words for the Ask box's microphone
    return await _run(dictation.start, eid)


@app.post("/api/dictation/usage")
async def dictation_usage(req: Dictated):
    return await _run(dictation.record, req.seconds)


class Ask(BaseModel):
    question: str
    model: str | None = None
    history: list[dict] = []
    session: str | None = None


@app.post("/api/engagements/{eid}/overlay/ask")
def overlay_ask(eid: int, req: Ask):
    """Chat about the engagement with the Model Desk's tools and the Python overlay's (NDJSON events)."""
    import json
    from fastapi.responses import StreamingResponse
    from starlette.concurrency import iterate_in_threadpool
    if req.model and req.model not in MODELS:
        raise HTTPException(400, f"unknown model {req.model}")

    def events():
        try:
            yield from engagement.overlay_ask(eid, req.question, req.model, req.history, req.session)
        except Exception as e:  # sign-in and other failures show in the chat, not as a broken stream
            yield {"type": "error", "text": engagement.friendly(e)}

    return StreamingResponse(iterate_in_threadpool(json.dumps(e, default=str) + "\n" for e in events()),
                             media_type="application/x-ndjson")


@app.get("/api/engagements/{eid}/overlay/module.py")
async def overlay_module(eid: int):
    path = await _run(engagement.overlay_module, eid)
    return FileResponse(path, media_type="text/x-python; charset=utf-8", filename=f"overlay_e{eid}.py",
                        content_disposition_type="inline")


@app.get("/api/engagements/{eid}/overlay/inputs")
async def overlay_inputs(eid: int, q: str = ""):
    return await _run(engagement.overlay_inputs, eid, q)


@app.get("/api/engagements/{eid}/overlay/trace")
async def overlay_trace(eid: int, cell: str):
    return await _run(engagement.overlay_trace, eid, cell)


@app.get("/api/usage")
def get_usage(session: str | None = None):
    s = usage.summary(session)
    return {"total": s["total"], "session": s["session"], "by_model": s["by_model"], "by_purpose": s["by_purpose"]}
