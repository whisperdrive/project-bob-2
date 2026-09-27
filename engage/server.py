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
import engagement  # noqa: E402
import library  # noqa: E402
import usage  # noqa: E402

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


@app.get("/api/config")
def config():
    return {"models": MODELS, "default_model": engagement.DEFAULT_MODEL, "default_reviewer": engagement.DEFAULT_REVIEWER,
            "roles": engagement.rolesmod.ROLES}


class NewEngagement(BaseModel):
    name: str


class EngagementPatch(BaseModel):
    name: str | None = None
    model: str | None = None
    reviewer_model: str | None = None


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
    for m in (body.model, body.reviewer_model):
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


@app.post("/api/engagements/{eid}/workbooks/{fid}/retry")
async def retry_workbook(eid: int, fid: int):
    await _run(lambda: library.retry(fid) and True)
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


@app.post("/api/engagements/{eid}/{kind}")
async def start_step(eid: int, kind: str):
    if kind not in ("compare", "map", "overlay"):
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
