"""How alike an engagement's workbooks are, and what each one looks like on its own: evidence for suggesting roles as
soon as the files are read, before the report has been (roles.py adds the report's evidence once there is some).

  signature(wb)   one workbook, from its row map (model.db) and the file itself: sheets, line items, formula
                  patterns, charts per sheet, valuation vocabulary found in labels and text, external links, the
                  file's author and company
  compare(a, b)   how alike two workbooks are: sheet names, line items (sheet + label) and formula patterns, each as
                  the share the two have in common, plus how much of each one the other contains, and the sheets
                  only one of them has
Years and period markers are ignored ("BP25 Inputs" = "BP26 Inputs", "Revenue FY25" = "Revenue FY26"), so last
year's and this year's client models come out nearly the same even though the timeline moved.

What that tells roles.py: two workbooks that are mostly the same are the client model twice (the earlier timeline
is the prior one); one that contains another plus some extra sheets is a client model with the overlay added, and
the extra sheets are the overlay; one that is unlike the others, small, full of valuation vocabulary (valuation
range, WACC, gearing, beta, terminal value, ...), with charts and links to another file, is a standalone overlay.
The adviser's own name is a strong sign too; it stays out of this public code: put it in .env as
VALUATION_DESK_OVERLAY_MARKERS (comma separated) and it is searched in the text, sheet names and file properties.
"""
import os
import re
import threading
import zipfile
from collections import Counter
from pathlib import Path, PurePosixPath
from xml.etree import ElementTree as ET

import rodb

# A term counts once per workbook (and per sheet), however often it appears.
TERMS = {
    "valuation range": r"valuation range|range of values|\blow\b.{0,40}\bhigh\b|\bpreferred\b|mid[- ]?point",
    "enterprise value": r"enterprise value|\bev\b(?!\s*/)",
    "equity value": r"equity value|value of equity",
    "WACC / discount rate": r"\bwacc\b|discount rate|cost of capital",
    "cost of equity / debt": r"cost of (equity|debt)|\bk[ed]\b",
    "gearing": r"gearing|debt\s*/\s*(ev|v|value)\b|leverage ratio",
    "time-weighted average": r"time[- ]weighted|\btwa\b",
    "beta": r"\bbeta\b|asset beta|equity beta",
    "risk premium / risk-free": r"risk premium|\bmrp\b|\berp\b|risk[- ]free",
    "terminal value": r"terminal (value|growth|year)|gordon|exit multiple|\bperpetuity\b",
    "valuation date": r"valuation date",
    "discounting": r"discount factor|mid[- ](year|period)|present value|\bnpv\b",
    "sensitivity": r"sensitivit",
    "multiples": r"ev\s*/\s*ebitda|implied multiple|trading multiple|transaction multiple|comparable",
    "net debt bridge": r"net debt|equity bridge|surplus assets",
}
_TERMS = {k: re.compile(v, re.I) for k, v in TERMS.items()}
_ROWREF = re.compile(r"R\[-?\d+\]|R\d+")
_NUMBER = re.compile(r"\d+(?:\.\d+)?")
_PERIOD = re.compile(r"\b(?:fy|cy|bp|h[12]|q[1-4])?\s?'?\d+\b", re.I)
_NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
       "p": "http://schemas.openxmlformats.org/package/2006/relationships"}
_R_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
_CACHE: dict = {}
_LOCK = threading.Lock()
MAX_TEXT_CELLS = 300_000


def norm(s) -> str:
    """Lower case, without years, period markers or punctuation: "BP25 Inputs" -> "inputs"."""
    s = _PERIOD.sub(" ", str(s or "").lower())
    return re.sub(r"[^a-z]+", " ", s).strip()


def markers() -> list[str]:
    """The adviser's names to look for (VALUATION_DESK_OVERLAY_MARKERS in .env), lower case."""
    raw = os.environ.get("VALUATION_DESK_OVERLAY_MARKERS", "")
    if not raw:
        try:
            for line in (Path(__file__).resolve().parent.parent / ".env").read_text(encoding="utf-8").splitlines():
                if line.strip().startswith("VALUATION_DESK_OVERLAY_MARKERS"):
                    raw = line.split("=", 1)[1].strip().strip('"').strip("'")
        except OSError:
            pass
    return [m.strip().lower() for m in raw.split(",") if m.strip()]


# ---- the file itself ------------------------------------------------------------------------------------------

def _rels(z: zipfile.ZipFile, part: str) -> dict[str, str]:
    p = PurePosixPath(part)
    name = str(p.parent / "_rels" / (p.name + ".rels"))
    if name not in z.namelist():
        return {}
    return {r.get("Id"): str((p.parent / r.get("Target")).as_posix()) if not r.get("Target", "").startswith("/")
            else r.get("Target").lstrip("/") for r in ET.fromstring(z.read(name)).findall("p:Relationship", _NS)}


def _resolve(path: str) -> str:
    out = []
    for part in path.split("/"):
        if part == "..":
            out and out.pop()
        elif part and part != ".":
            out.append(part)
    return "/".join(out)


def file_facts(path: str) -> dict:
    """Charts per sheet (through each sheet's drawing), and the author, last editor, company and title."""
    out = {"charts": {}, "props": {}}
    if not path or not str(path).lower().endswith((".xlsx", ".xlsm")) or not Path(path).exists():
        return out
    try:
        with zipfile.ZipFile(path) as z:
            names = set(z.namelist())
            wb = ET.fromstring(z.read("xl/workbook.xml"))
            wb_rels = _rels(z, "xl/workbook.xml")
            for sh in wb.findall("m:sheets/m:sheet", _NS):
                part = _resolve(wb_rels.get(sh.get(_R_ID), ""))
                n = 0
                for target in _rels(z, part).values():
                    target = _resolve(target)
                    if "/drawings/" in target and target in names:
                        n += sum(1 for t in _rels(z, target).values() if "/charts/" in t)
                if n:
                    out["charts"][sh.get("name")] = n
            for part, tags in (("docProps/core.xml", ("creator", "lastModifiedBy", "title")),
                               ("docProps/app.xml", ("Company", "Manager"))):
                if part in names:
                    root = ET.fromstring(z.read(part))
                    for el in root.iter():
                        tag = el.tag.split("}")[-1]
                        if tag in tags and (el.text or "").strip():
                            out["props"][tag[0].lower() + tag[1:]] = el.text.strip()
    except (zipfile.BadZipFile, KeyError, ET.ParseError):
        pass
    return out


# ---- one workbook ---------------------------------------------------------------------------------------------

def signature(wb: dict) -> dict:
    """wb: {"db_path", "source_path", "filename"}. Cached per model.db."""
    path = wb["db_path"]
    key = (path, os.path.getmtime(path), wb.get("source_path"))
    with _LOCK:
        if key in _CACHE:
            return _CACHE[key]
    db = rodb.connect(path)
    have = {r[0] for r in db.execute("SELECT name FROM sqlite_master WHERE type='table'")}
    sheets = [r[0] for r in db.execute("SELECT sheet FROM sheets ORDER BY rowid")]
    starts = [r[0] for r in db.execute("SELECT json_extract(layout, '$.period_start') FROM sheets") if r[0]]
    items, patterns = Counter(), Counter()
    per_sheet = {s: {"rows": 0, "formulas": 0, "terms": set()} for s in sheets}
    terms: dict[str, str] = {}  # term -> first place it was seen

    def seen(text, sheet, where):
        for t, rx in _TERMS.items():
            if rx.search(text or ""):
                terms.setdefault(t, where)
                if sheet in per_sheet:
                    per_sheet[sheet]["terms"].add(t)

    for s, r, label, section, nf, pats in db.execute("SELECT sheet, row, label, section, n_formula, patterns FROM rows"):
        ns, nl = norm(s), norm(label)
        if nl:
            items[(ns, nl)] += 1
        for pat in (pats or "").split("; "):
            f = pat.split(" x")[0].strip()
            if f.startswith("="):  # the formula's shape: a row inserted above shifts every R[n], and years and
                patterns[(ns, nl, _NUMBER.sub("#", _ROWREF.sub("R", f)))] += 1  # other constants move on, so both go
        if s in per_sheet:
            per_sheet[s]["rows"] += 1
            per_sheet[s]["formulas"] += nf or 0
        seen(label, s, f"{s}!r{r} {label}")
        seen(section, s, f"{s} section {section}")
    for s in sheets:
        seen(s, s, f"sheet {s}")
    mk = markers()
    hits: dict[str, str] = {}
    n = 0
    for s, addr, v in db.execute("SELECT sheet, addr, value FROM cells WHERE formula IS NULL AND typeof(value)='text' "
                                 "AND length(value) BETWEEN 3 AND 300 LIMIT ?", (MAX_TEXT_CELLS,)):
        n += 1
        seen(v, s, f"{s}!{addr} “{v[:60]}”")
        low = v.lower()
        for m in mk:
            if m in low:
                hits.setdefault(m, f"{s}!{addr}")
    ext = [r[0] for r in db.execute("SELECT filename FROM extbooks WHERE filename IS NOT NULL")] if "extbooks" in have else []
    db.close()
    ff = file_facts(wb.get("source_path"))
    for m in mk:
        for k, v in ff["props"].items():
            if m in v.lower():
                hits.setdefault(m, f"file {k}")
        for s in sheets:
            if m in s.lower():
                hits.setdefault(m, f"sheet name {s}")
        if m in (wb.get("filename") or "").lower():
            hits.setdefault(m, "file name")
    out = {"filename": wb.get("filename"), "sheets": sheets, "sheet_keys": {norm(s) or s.lower(): s for s in sheets},
           "items": items, "patterns": patterns, "per_sheet": {s: {**v, "terms": sorted(v["terms"])} for s, v in per_sheet.items()},
           "terms": terms, "charts": ff["charts"], "props": ff["props"], "markers": hits, "external_links": ext,
           "formulas": sum(v["formulas"] for v in per_sheet.values()), "line_items": sum(items.values()),
           "timeline_start": min(starts) if starts else None, "text_cells": n}
    with _LOCK:
        _CACHE[key] = out
    return out


# ---- two workbooks --------------------------------------------------------------------------------------------

def _share(a: Counter, b: Counter) -> tuple[float, float, float]:
    """(in common / either, share of a found in b, share of b found in a)."""
    both = sum((a & b).values())
    ta, tb = sum(a.values()), sum(b.values())
    either = ta + tb - both
    return (both / either if either else 0.0, both / ta if ta else 0.0, both / tb if tb else 0.0)


def compare(a: dict, b: dict) -> dict:
    sa, sb = Counter(a["sheet_keys"].keys()), Counter(b["sheet_keys"].keys())
    s_j, _, _ = _share(sa, sb)
    i_j, i_ab, i_ba = _share(a["items"], b["items"])
    f_j, f_ab, f_ba = _share(a["patterns"], b["patterns"])
    similarity = 0.2 * s_j + 0.45 * i_j + 0.35 * f_j
    only = lambda x, y: [x["sheet_keys"][k] for k in x["sheet_keys"] if k not in y["sheet_keys"]]
    return {"similarity": round(100 * similarity), "sheets": round(100 * s_j), "line_items": round(100 * i_j),
            "formulas": round(100 * f_j), "a_in_b": round(100 * i_ab), "b_in_a": round(100 * i_ba),
            "formulas_a_in_b": round(100 * f_ab), "formulas_b_in_a": round(100 * f_ba),
            "shared_sheets": sum((sa & sb).values()), "only_a": only(a, b), "only_b": only(b, a)}


def public(sig: dict) -> dict:
    """What the page shows for one workbook."""
    return {k: sig[k] for k in ("sheets", "terms", "charts", "props", "markers", "external_links", "formulas",
                                "line_items", "timeline_start")} | {
        "sheet_terms": {s: v["terms"] for s, v in sig["per_sheet"].items() if v["terms"]}}
