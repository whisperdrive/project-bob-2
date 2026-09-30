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
- This year's valuation date is confirmed by the agents where the files agree on it (another cell labelled like it,
  the file name, a year after last year's), not where anything disagrees (the file name, a date not after last
  year's); a row naming another date counts neither way, and the profile's fallback year end isn't counted twice.
- Last year's engagement is found through the overlay's earlier versions: the nearest with a confirmed schedule,
  an unconfirmed one skipped.
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
    profile_check(eid)
    date_check(tmp)
    link_check()
    carry_check(tmp)


def carry_check(tmp: Path) -> None:
    """The carry on the engagement: last year's classes onto this year's rows, and its outside-the-model marks for
    the figures still on no output or input this year (a figure now on a lever loses its mark)."""
    import sqlite3
    import build_map
    import outputs
    import xlsxwriter

    def overlay(name, labels):
        path = tmp / f"{name}.xlsx"
        wb = xlsxwriter.Workbook(path)
        ws = wb.add_worksheet("Out")
        for i, label in enumerate(labels):
            ws.write(4 + i, 1, label)
            ws.write_formula(4 + i, 3, "=1+1", None, 2.0)
        wb.close()
        return build_map.main(str(path), str(tmp / f"{name}_db"))["db"]
    last = overlay("glue_last", ["Equity value", "Low", "Selected valuation"])
    now = overlay("glue_now", ["Opening note", "Equity value", "Low", "Selected valuation"])
    src, dst = (engagement.create(n)["id"] for n in ("Asset B, FY25", "Asset B, FY26"))
    engagement._set("engagements", src, schedule_json=json.dumps(
        {"confirmed_at": 1, "classes": {"Out!r6": "conclusion"}, "outside": ["net_debt", "ev_multiple"]}))
    engagement._SCHEDULES[dst] = (None, outputs.detect(sqlite3.connect(now), ["Out"]))
    summary = {"wiring": {"overlay": {"db_path": now}}, "sheets": ["Out"], "outputs": [],
               "levers": [{"key": "ev_multiple", "cell": "Out!D5"}]}  # this year the multiple is an input
    facts = [{"key": k, "category": "conclusion", "label": k, "value_text": "1.0"} for k in ("net_debt", "ev_multiple")]
    was = engagement._role_wb, engagement.schedule_view, engagement.reference
    engagement._role_wb = lambda eid, role: {"id": 1, "db_path": last}
    engagement.schedule_view = lambda eid: {"outputs": outputs.apply(outputs.detect(sqlite3.connect(last), ["Out"]),
                                                                     engagement._schedule(src))}
    engagement.reference = lambda eid: facts
    try:
        mine = engagement._carry(dst, src, summary)
    finally:
        engagement._role_wb, engagement.schedule_view, engagement.reference = was
    c = mine["carried"]
    assert c["from"] == src and c["classes"]["Out!r7"] == "conclusion" and not c["missing"], c
    assert mine["outside"] == ["net_debt"] and c["outside_dropped"] == ["ev_multiple"], mine
    assert not mine.get("confirmed_at"), "carried, not confirmed: the person looks at what's new first"
    try:
        engagement._carry(src, src, summary)
        raise AssertionError("carried its own schedule")
    except ValueError:
        pass
    print("carry: ok (last year's classes onto this year's rows; its outside marks for figures still outside; "
          "not confirmed until a person looks)")


def link_check() -> None:
    import library
    a, b, c = (engagement.create(n)["id"] for n in ("Asset A, FY25", "Asset A, FY26", "Asset A, FY24"))
    for eid, fid in ((a, 19), (b, 20), (c, 18)):
        engagement._exec("INSERT INTO roles VALUES (?,?,?,?,?,?,1)", eid, "prior_overlay", "workbook", fid, "null", "[]")
    was = library.get, engagement._role_wb
    chain = {20: 19, 19: 18, 18: None}  # this year's overlay, an earlier version of it, and one before that
    library.get = lambda fid, full=False: {"id": fid, "previous_id": chain.get(fid)}
    engagement._role_wb = lambda eid, role: {"id": {a: 19, b: 20, c: 18}[eid]}
    try:
        assert engagement._schedule_source(b) is None, "no confirmed schedule: nothing to carry"
        engagement._set("engagements", c, schedule_json='{"confirmed_at": 1}')
        assert engagement._schedule_source(b) == c, "the confirmed one further back"
        engagement._set("engagements", a, schedule_json='{"confirmed_at": 2}')
        assert engagement._schedule_source(b) == a, "the nearest confirmed one"
        assert {e["id"] for e in engagement._schedule_sources(b)["engagements"]} == {a, c}
        assert engagement._schedule_source(a) == c and engagement._schedule_source(c) is None
    finally:
        library.get, engagement._role_wb = was
    print("link: ok (last year's engagement through the overlay's earlier versions: the nearest confirmed schedule)")


def date_check(tmp: Path) -> None:
    import build_map
    import library
    import xlsxwriter
    from datetime import date
    path = tmp / "model.xlsx"
    wb = xlsxwriter.Workbook(path)
    dt = wb.add_format({"num_format": "dd-mmm-yy"})
    val, log = wb.add_worksheet("Val"), wb.add_worksheet("Log")
    val.write(19, 1, "Valuation Date")
    val.write_datetime(19, 3, date(2026, 6, 30), dt)
    val.write(24, 1, "Roll forward valuation date (to 30/9/2025)")  # names another date: counts neither way
    val.write_datetime(24, 3, date(2025, 9, 30), dt)
    log.write(89, 1, "Update valuation date")
    log.write_datetime(89, 3, date(2026, 6, 30), dt)
    wb.close()
    db = build_map.main(str(path), str(tmp / "model_db"))["db"]
    was = engagement._prior_vd, engagement._profile, engagement.roles, library.get, library.note_identity, library.confirm_identity
    engagement._prior_vd, engagement._profile = (lambda eid: "2025-06-30"), (lambda eid: {})
    try:
        w = {"id": 7, "status": "done", "identity_confirmed": 0, "valuation_date": "2026-06-30", "db_path": db,
             "filename": "Model_Jun 26.xlsx"}
        ev = engagement._date_evidence(1, w, "Val!D20")
        assert len(ev["agree"]) == 3 and not ev["disagree"], ev  # the other cell, the file name, a year on; no year end
        assert engagement._date_evidence(1, {**w, "filename": "Model_Dec 25.xlsx"}, "Val!D20")["disagree"]
        assert not engagement._date_evidence(1, {**w, "filename": "20260521_Model.xlsx"}, "Val!D20")["disagree"], \
            "a date stamp in digits is when the file was saved: neither way"
        engagement._prior_vd = lambda eid: "2026-06-30"
        assert any("not after" in x for x in engagement._date_evidence(1, w, "Val!D20")["disagree"])
        # the whole check: confirmed by the agents, once
        engagement._prior_vd = lambda eid: "2025-06-30"
        calls = []
        engagement.roles = lambda eid: {"current_model": {"kind": "workbook", "id": 7}}
        library.get = lambda fid, full=False: {"identity": {"valuation_date_evidence": "Val!D20"}}
        library.note_identity = lambda fid, **k: calls.append(("note", k["auto_check"]["agree"]))
        library.confirm_identity = lambda fid, by, why=None: calls.append((by, why))
        assert engagement._agents_check_dates(1, [w]) is True and calls[-1][0] == "agents" and len(calls[-1][1]) == 3
        assert engagement._agents_check_dates(1, [w]) is False, "weighed once per date"
        assert engagement._agents_check_dates(2, [{**w, "identity_confirmed": 1}]) is False, "a confirmed date is left alone"
    finally:
        engagement._prior_vd, engagement._profile, engagement.roles, library.get, library.note_identity, \
            library.confirm_identity = was
    print("date: ok (confirmed by the agents on three agreeing signals; a file name or a date not after last year's "
          "says otherwise; a row naming another date and the fallback year end count neither way)")


def profile_check(eid: int) -> None:
    import chartdata
    engagement.set_profile(eid, {"fy_end_month": 6, "horizon": "fixed"})
    assert engagement._profile(eid) == {"fy_end_month": 6, "horizon": "fixed"}
    with engagement._fy_hint(eid):
        assert chartdata._FY_HINT.get() == (6, True), "a person's year end is forced on the charts"
    for bad in ({"fy_end_month": 13}, {"horizon": "sideways"}, {"frequency": "annual"}):
        try:
            engagement.set_profile(eid, bad)
            raise AssertionError(f"accepted {bad}")
        except ValueError:
            pass
    engagement.set_profile(eid, {"horizon": None})
    assert engagement._profile(eid) == {"fy_end_month": 6}
    try:
        engagement.set_schedule(eid, classes={"Val!r3": "headline"})
        raise AssertionError("accepted a class that isn't one")
    except ValueError:
        pass
    view = engagement.profile_view(eid)["fields"]
    assert view["fy_end_month"]["set"] and view["fy_end_month"]["shown"] == "June" and not view["horizon"]["set"]
    print("profile: ok (a year end and a horizon set, checked and cleared; the charts take the year end as set)")


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
