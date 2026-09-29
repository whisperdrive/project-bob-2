"""Dictation in Ask (bench/dictation.py): the words handed to Azure Speech as hints, the messages for the token
service's refusals, and the setup message without a key. No Azure calls.

    uv run python tests/check_dictation.py
"""
import json
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "bench"))
import dictation  # noqa: E402


def phrase_check():
    p = dictation._phrase
    assert p("Other_Rev (A$m)") == "Other Rev", p("Other_Rev (A$m)")
    assert p("FY26 capex") == "FY26 capex"
    assert p("[units] 2025") is None and p("12.5%") is None and p("$1,234.5m") is None and p("") is None
    assert p("x" * 80) is None
    print("phrases: labels as they're said; numbers, notes and long text left out")


def phrases_check():
    """An engagement's words come first (the names in its facts, then their labels, its outputs, levers and sheets),
    each once, then the valuation words; never more than MAX_PHRASES."""
    class FakeEngagement:
        @staticmethod
        def _q(sql, eid):
            summary = {"outputs": [{"label": "Enterprise value"}, {"label": "North plaza revenue"}, {"label": "Plaza revenue"}],
                       "levers": [{"label": "Discount rate"}]}
            return [{"name": "Plaza test (synthetic)", "overlay_json": json.dumps(summary)}]

        @staticmethod
        def reference(eid):
            return [{"label": "Target name", "value_text": "Plaza Holdings Pty Ltd"},
                    {"label": "Discount rate", "value_text": "7.25%"},
                    {"label": "Terminal value method", "value_text": "Gordon growth method"}]

        @staticmethod
        def _role_wb(eid, key):
            return {"sheets": {"Val_Inputs", "DCF"}} if key == "prior_overlay" else None

    sys.modules["engagement"] = FakeEngagement
    try:
        got = dictation.phrases(1)
    finally:
        del sys.modules["engagement"]
    # a name whole, then its distinctive word on its own ("Plaza's" comes out right); the legal ending isn't said
    assert got[:4] == ["Plaza Holdings Pty Ltd", "Plaza", "Gordon growth method", "Gordon"], got[:5]
    # labels read as ordinary words mid-sentence, unless they start with a name
    assert got.index("target name") < got.index("north plaza revenue") < got.index("DCF") < got.index("Val Inputs")
    assert "Plaza revenue" in got and "enterprise value" in got and "Target name" not in got
    assert got.index("Plaza test") < got.index("WACC")         # "(synthetic)" is a note, left out
    assert got.count("discount rate") == 1 and "7.25%" not in got
    assert "Gordon growth" in got and "roll forward" in got and "roll-forward" not in got and len(got) <= dictation.MAX_PHRASES
    print(f"engagement words: {len(got)} hints, names first, each once")


def refusal_check():
    w = dictation._why
    assert "SPEECH_KEY" in w(401, "")
    assert "used up" in w(403, '{"error":{"code":"403","message":"Out of call volume quota. Quota will be replenished in 2.12 days."}}')
    assert "one at a time" in w(429, "")
    assert "(500)" in w(500, "boom")
    print("refusals: a bad key, the month's free hours used up and a busy stream each say what to do")


def setup_check():
    saved = {k: os.environ.pop(k, None) for k in ("SPEECH_KEY", "SPEECH_REGION", "SPEECH_ENDPOINT")}
    real = dictation._load_env
    dictation._load_env = lambda: None       # this machine's .env may hold a key
    try:
        assert dictation.status()["configured"] is False
        try:
            dictation.token()
            raise AssertionError("a token without a key")
        except ValueError as e:
            assert "SPEECH_KEY and SPEECH_ENDPOINT (or SPEECH_REGION)" in str(e)
    finally:
        dictation._load_env = real
        os.environ.update({k: v for k, v in saved.items() if v is not None})
    print("setup: without a key the desk says which lines .env needs")


if __name__ == "__main__":
    phrase_check()
    phrases_check()
    refusal_check()
    setup_check()
    print("dictation checks passed")
