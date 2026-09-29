"""The report's review workflow after a person settles a table (bench/engagement.py), with no model calls:

- Taking the reviewer's read of a table sends back to the fact review loop only the facts handed to a person that
  rest on that table's page; a fact on another page stays with the person (the same loop on the same text would
  come out the same), and a fact only a check held up is approved by the agents once the check passes.
- The loop takes just those facts; if a loop is already running, they get one of their own when it finishes.
- Report jobs and model jobs run in separate lanes, so a review loop never holds up the map, and a queued job says
  what it's waiting behind.
- A fact's number is worked out by code from its text: none for a name holding digits or a range without a
  preferred point. Stored facts are brought up to date once, and one handed to a person whose number changed goes
  back to the loop.
- Retrying a workbook that failed lets go of the live Python overlays reading its model.db first, as a rebuild does
  (Windows won't delete an open file, so a retry without it failed every time).

    uv run python tests/check_report_loop.py
"""
import json
import sys
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "bench"))
import engagement  # noqa: E402
import reportfacts  # noqa: E402

FIRST = "| Item | Value |\n|---|---|\n| Discount rate | 7.2% |"
SECOND = "| Item | Value |\n|---|---|\n| Discount rate | 7.0% |"


def _doc() -> dict:
    table = {"id": "p2-t1", "page": 2, "source": "text-layer table", "title": "Key assumptions", "status": "flagged",
             "markdown": FIRST, "second_markdown": SECOND, "check": {"ok": False, "method": "image"}}
    pages = [{"n": 1, "blocks": [[0, "text", "Summary of the valuation.", True]]},
             {"n": 2, "blocks": [[0, "text", "The assumptions are below.", True], [0, "table", "p2-t1"]]},
             {"n": 3, "blocks": [[0, "text", "Terminal growth of 2.5% a year.", True]]}]
    doc = {"kind": "pdf", "pages": pages, "tables": [table]}
    engagement.docingest.render(doc)
    return doc


def _fact(eid, did, n, key, value_text, page, quote, pg, open_):
    f = {"category": "assumption", "key": key, "label": key, "value_text": value_text,
         "value": float(value_text.rstrip("%")), "unit": "%", "page": page, "quote": quote}
    agent = {"status": "escalated", "round": 3, "thread": [{"round": 0}], "open": open_}
    return engagement._exec(
        """INSERT INTO facts(engagement_id, document_id, n, category, key, label, value_text, value, unit, page, quote,
           origin, check_json, review_json, status, agent_json) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        eid, did, n, f["category"], key, key, value_text, f["value"], "%", page, quote, "extractor",
        json.dumps(reportfacts.check(f, pg)), json.dumps({"verdict": "correct", "reason": "see the open point"}),
        "pending", json.dumps(agent))


def main() -> None:
    tmp = Path(tempfile.mkdtemp(prefix="report_loop_"))
    engagement.OUT, engagement.DB, engagement.DOCS = tmp, tmp / "engage.db", tmp / "docs"
    eid = engagement.create("Neutral asset, FY26")["id"]
    doc = _doc()
    out_dir = tmp / "docs" / "report"
    out_dir.mkdir(parents=True)
    did = engagement._exec("""INSERT INTO documents(engagement_id, sha256, filename, kind, status, step, out_dir,
                              doc_json, facts_status, source_path) VALUES (?,?,?,?,?,?,?,?,?,?)""",
                           eid, "0" * 64, "report.pdf", "pdf", "done", "Done", str(out_dir), json.dumps(doc), "done", "")
    pg = reportfacts.pages(doc["markdown"])
    objected = {"verdict": "object", "reason": "the rate looks wrong", "correction": None}
    on_page = _fact(eid, did, 1, "discount_rate", "7.0%", 2, "Discount rate | 7.0%", pg, objected)
    elsewhere = _fact(eid, did, 2, "terminal_growth", "2.5%", 3, "Terminal growth of 2.5% a year", pg, objected)
    held_by_check = _fact(eid, did, 3, "discount_rate_nominal", "7.0%", 2, "Discount rate | 7.0%", pg,
                          {"verdict": "object", "accepted": True, "reason": "accepted by the reviewer, but the checks fail"})

    got = engagement.settle_table(did, "p2-t1", "use_second")
    facts = {f["id"]: f for f in engagement.facts(eid)}
    assert got["reopened"] == 1, got["reopened"]
    assert facts[held_by_check]["status"] == "approved" and facts[held_by_check]["decided_by"] == "agents", \
        "a fact only the check held up is approved by the agents once it passes"
    d = engagement._doc(did)
    assert d["facts_status"] == "queued" and "1 fact(s)" in d["facts_step"], d["facts_step"]

    taken = []

    def fake_resolve(md, fs, *a, **k):  # the loop, without the models: record what it was given
        taken.extend(f["id"] for f in fs)
        for f in fs:
            f["agent"] = {"status": "escalated", "round": 1, "thread": [{"round": 0}, {"round": 1}]}
        return {"summary": {"rounds": 1, "agreed": 0, "withdrawn": 0, "escalated": len(fs), "settled_in_loop": 0,
                            "arbitrated": 0}, "episodes": []}

    engagement.reportfacts.resolve, engagement._learn = fake_resolve, lambda *a, **k: None
    engagement._resolve_facts_job(did)
    assert taken == [on_page], f"the loop takes only the fact on the settled table's page, not {taken}"
    assert facts[elsewhere]["status"] == "pending", "the fact on another page stays with the person"
    print("report loop: ok (a table decision sends back only the facts on its page; one only a check held up is "
          "approved)")

    lanes = engagement._jobs.queues
    assert [j for j in lanes["report"].queue] == [("resolve_facts", did)] and not lanes["models"].queue
    lanes["report"].get_nowait()
    engagement._jobs.running["models"] = ("rows", eid, time.time() - 600)
    engagement._jobs.put(("map", eid))
    engagement._jobs.put(("resolve_facts", did))
    view = {(j["job"], j["state"]): j for j in engagement.jobs_view(eid)}
    assert ("rows", "running") in view
    assert view[("map", "queued")]["behind"]["label"] == "the row agents" and view[("map", "queued")]["ahead"] == 1
    assert view[("resolve_facts", "queued")]["behind"] is None, "the fact loop doesn't wait behind the row agents"
    print("lanes: ok (the map waits behind the row agents and says so; the fact loop has its own lane)")

    for q in lanes.values():
        q.queue.clear()
    engagement._jobs.running.clear()
    engagement._set("documents", did, facts_status="running", facts_step="Fact review loop, round 1")
    got = engagement.settle_table(did, "p2-t1", "reset")  # undone while a loop runs: its page's facts wait for it
    assert got["reopened"] == 2 and got["after_loop"] and not lanes["report"].queue, got["reopened"]
    engagement._set("documents", did, facts_status="done", facts_step="Done")
    engagement._next_loop(did)  # the running loop finished
    assert list(lanes["report"].queue) == [("resolve_facts", did)] and engagement._REOPEN[did] == {on_page, held_by_check}
    print("mid-loop: ok (facts a decision reopens while a loop runs get a loop of their own when it finishes)")
    numbers_check(eid, tmp)
    retry_check()


def retry_check() -> None:
    import library
    import overlay as ovmod
    closed = []
    sess = type("Sess", (), {"paths": lambda self: ["/books/model.db"], "close": lambda self: closed.append(1)})()
    was = engagement.workbooks, library.retry, ovmod.deep
    engagement.workbooks = lambda eid: [{"id": 5, "db_path": "/books/model.db"}]
    library.retry = lambda fid: {"id": fid, "status": "queued"}
    ovmod.deep = lambda fn, *a: fn(*a)
    engagement._SESSIONS[42] = (sess, {})
    try:
        assert engagement.retry_workbook(42, 5)["status"] == "queued"
        assert closed == [1] and 42 not in engagement._SESSIONS, "the live overlay let go of the model.db first"
    finally:
        engagement.workbooks, library.retry, ovmod.deep = was
    print("retry: ok (a failed workbook's retry closes the live overlays reading it, as a rebuild does)")


def numbers_check(eid: int, tmp: Path) -> None:
    sv = reportfacts.settle_value
    assert sv({"category": "identity", "unit": "text", "value_text": "Project 30", "value": 30.0})["value"] is None
    assert sv({"category": "assumption", "unit": "%", "value_text": "1.25%–1.75%", "low_text": "1.25%",
               "high_text": "1.75%", "value": 1.5})["value"] is None, "a range with no preferred point has no number"
    assert sv({"category": "conclusion", "unit": "A$m", "value_text": "A$2,296.7m", "low_text": "A$2,200.0m",
               "high_text": "A$2,400.0m", "value": None})["value"] == 2296.7
    assert sv({"category": "identity", "unit": "date", "value_text": "30 June 2025", "value": 20250630.0})["value"] == 20250630.0

    # a fact handed to a person because the reviewer wanted no number on a name: once code clears it, back to the loop
    for q in engagement._jobs.queues.values():
        q.queue.clear()
    doc = _doc()
    out_dir = tmp / "docs" / "second"
    out_dir.mkdir(parents=True)
    did = engagement._exec("""INSERT INTO documents(engagement_id, sha256, filename, kind, status, step, out_dir,
                              doc_json, facts_status, source_path, loop_json) VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
                           eid, "1" * 64, "second.pdf", "pdf", "done", "Done", str(out_dir), json.dumps(doc), "done", "",
                           json.dumps({"check_version": reportfacts.CHECK_VERSION - 1, "facts": {"rounds": 3}}))
    pg = reportfacts.pages(doc["markdown"])
    name = _fact(eid, did, 1, "project_name", "30%", 1, "Summary of the valuation.", pg,
                 {"verdict": "object", "reason": "a name takes no number", "correction": None})
    engagement._set("facts", name, category="identity", unit="text", value_text="Summary 30", value=30.0)
    engagement.get(eid)  # a poll: the stored facts are brought up to date once
    f = next(x for x in engagement.facts(eid) if x["id"] == name)
    assert f["value"] is None and engagement._REOPEN.get(did) == {name}, (f["value"], engagement._REOPEN.get(did))
    assert ("resolve_facts", did) in list(engagement._jobs.queues["report"].queue)
    print("numbers: ok (none for a name or an unpicked range, by code; a fact held on it goes back to the loop)")


if __name__ == "__main__":
    main()
