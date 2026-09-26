"""Local app: upload workbooks, see them processed, check target / valuation date, compare versions, ask questions.
    uv run uvicorn app.server:app --port 8000      then open http://localhost:8000
"""
import json
import os
import sys
import tempfile
from contextlib import asynccontextmanager
from pathlib import Path

from azure.core.exceptions import ClientAuthenticationError
from azure.identity import AuthenticationRequiredError
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, StreamingResponse
from pydantic import BaseModel
from starlette.concurrency import iterate_in_threadpool, run_in_threadpool

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "bench"))
import agent  # noqa: E402
import library  # noqa: E402
import usage  # noqa: E402

# Chat deployments offered in the app (first = default). gpt-4 (gpt-4.1), o3-mini and DeepSeek-V4-Flash
# were dropped on 2026-09-25: their capacities were too low for multi-turn tool use.
MODELS = ["gpt-4o", "gpt-6-sol", "gpt-6-luna", "gpt-4o-mini", "gpt-5-nano"]


@asynccontextmanager
async def lifespan(app):
    library.start()
    yield


app = FastAPI(lifespan=lifespan)


@app.get("/")
def index():
    return FileResponse(Path(__file__).parent / "index.html")


@app.get("/api/config")
def config():
    return {"models": MODELS}


@app.get("/api/files")
def files():
    return library.all_files()


@app.get("/api/files/check")
def check(sha: str):
    """Lets the browser skip uploading a file it has already sent (it hashes the file first)."""
    return {"file": library.by_sha(sha.lower())}


@app.post("/api/files")
async def upload(file: UploadFile = File(...)):
    library.UPLOADS.mkdir(exist_ok=True)
    fd, tmp_name = tempfile.mkstemp(dir=library.UPLOADS, suffix=".part")
    os.close(fd)
    tmp = Path(tmp_name)
    h = library.hashlib.sha256()
    with open(tmp, "wb") as out:
        while chunk := await file.read(1 << 20):
            h.update(chunk)
            out.write(chunk)
    try:
        status, rec = await run_in_threadpool(library.add_upload, tmp, file.filename, h.hexdigest())
    except ValueError as e:
        raise HTTPException(400, str(e))
    return {"status": status, "file": rec}


@app.get("/api/files/{fid}")
def file_detail(fid: int):
    rec = library.get(fid, full=True)
    if not rec:
        raise HTTPException(404, "no such file")
    rec.pop("source_path", None)
    return rec


@app.delete("/api/files/{fid}")
async def delete_file(fid: int):
    try:
        await run_in_threadpool(library.remove, fid)
    except ValueError as e:
        raise HTTPException(409, str(e))
    return {"ok": True}


class Identity(BaseModel):
    target_name: str | None = None
    project_name: str | None = None
    valuation_date: str | None = None


@app.put("/api/files/{fid}/identity")
async def put_identity(fid: int, body: Identity):
    return await run_in_threadpool(library.set_identity, fid, body.target_name, body.project_name, body.valuation_date)


class CompareWith(BaseModel):
    previous_id: int | None


@app.put("/api/files/{fid}/compare")
async def put_compare(fid: int, body: CompareWith):
    if body.previous_id == fid:
        raise HTTPException(400, "pick a different file to compare with")
    await run_in_threadpool(library.compare, fid, body.previous_id)
    return library.get(fid, full=True)


@app.get("/api/files/{fid}/valuation")
async def get_valuation(fid: int, cell: str | None = None, low: float | None = None, high: float | None = None):
    """The Valuation tab: DCFs found in the workbook's formulas and recomputed in Python (no model calls)."""
    import valuation
    rec = library.get(fid, full=True)
    if not rec or rec["status"] != "done":
        raise HTTPException(404, "no processed file with that id")
    try:
        return await run_in_threadpool(valuation.view, rec["db_path"], cell, low, high)
    except Exception as e:
        raise HTTPException(500, f"{type(e).__name__}: {e}")


@app.get("/api/usage")
def get_usage(session: str | None = None):
    import ratelimit
    return {**usage.summary(session), "limits": ratelimit.status()}


class Ask(BaseModel):
    question: str
    file_id: int
    model: str = MODELS[0]
    history: list[dict] = []
    session: str | None = None


def _context(rec: dict) -> str:
    lines = [f"File: {rec['filename']}",
             f"Target: {rec['target_name'] or 'unknown'}; project: {rec['project_name'] or 'n/a'}; "
             f"valuation date: {rec['valuation_date'] or 'unknown'}"
             + (" (confirmed by the user)" if rec["identity_confirmed"] else " (identified automatically)")]
    if rec.get("diff_summary") and rec.get("diff"):
        lines.append(f"Changes versus the previous version ({rec['diff']['previous']['filename']}):\n"
                     + rec["diff_summary"])
    return "\n".join(lines)


@app.post("/api/ask")
def ask(req: Ask):
    rec = library.get(req.file_id, full=True)

    def events():
        if not rec or rec["status"] != "done":
            yield {"type": "error", "text": "That file isn't ready yet. Wait for processing to finish."}
            return
        log = lambda model, u: usage.record(model, u, "chat", req.file_id, req.session)
        for attempt in (1, 2):
            started = False
            try:
                # Never prompt from the server: a device code would only show up in the uvicorn log.
                for ev in agent.ask(req.question, rec["db_path"], req.model, req.history, interactive=False,
                                    context=_context(rec), on_usage=log, file_id=req.file_id,
                                    session=req.session):
                    started = True
                    yield ev
                return
            except AuthenticationRequiredError:
                yield {"type": "error", "text": "Your Azure sign-in has expired. Run `uv run python bench/llm.py` "
                                                "in a terminal, sign in with the device code, then ask again."}
                return
            except ClientAuthenticationError as e:
                # A long-running server can fail to refresh its cached token (e.g. the Keychain was locked
                # while the Mac slept). Retry once with a fresh credential before giving up.
                if attempt == 1 and not started:
                    continue
                yield {"type": "error", "text": "Couldn't refresh your Azure sign-in "
                       f"({e.message or 'no details'}). Run `uv run python bench/llm.py` in a terminal to "
                       "sign in again, then restart the app and ask again."}
                return
            except Exception as e:
                yield {"type": "error", "text": f"{type(e).__name__}: {e}"}
                return

    # agent.ask blocks on network calls, so run the generator in a worker thread.
    lines = iterate_in_threadpool(json.dumps(e) + "\n" for e in events())
    return StreamingResponse(lines, media_type="application/x-ndjson")
