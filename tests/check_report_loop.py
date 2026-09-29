"""The report's review workflow after a person settles a table (bench/engagement.py), with no model calls:

- Taking the reviewer's read of a table sends back to the fact review loop only the facts handed to a person that
  rest on that table's page; a fact on another page stays with the person (the same loop on the same text would
  come out the same), and a fact only a check held up is approved by the agents once the check passes.
- The loop takes just those facts.
- Report jobs and model jobs run in separate lanes, so a review loop never holds up the map, and a queued job says
  what it's waiting behind.

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


if __name__ == "__main__":
    main()
