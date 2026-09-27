"""One workbook's model.db at a glance, for the Map step's model dashboards: sheets, formula cells and inputs,
which sheets feed which, the most-read rows, line items and named ranges. Read-only; no model calls.

Same views as the separate model dashboard (dashboard/server.py), opened with rodb so paths work on Windows.
"""
import json
import os
import threading

import rodb

LIVE = ("direct", "offset", "active")  # edge kinds the current scenario actually uses (see edges.py)
_LIVE_SQL = ",".join(f"'{k}'" for k in LIVE)
_CACHE: dict = {}
_LOCK = threading.Lock()


def _q(db, sql: str, *args) -> list[dict]:
    cur = db.execute(sql, args)
    cols = [c[0] for c in cur.description]
    return [dict(zip(cols, r)) for r in cur]


def _json(s):
    try:
        return json.loads(s) if s else None
    except ValueError:
        return None


def summary(db_path: str) -> dict:
    """Sheets with their counts and timeline, edge kinds, sheet-to-sheet flows, the most-read and most-reading
    rows, sections and named ranges. Cached per model.db (the big workbooks take a few seconds)."""
    key = (db_path, os.path.getmtime(db_path))
    with _LOCK:
        if key in _CACHE:
            return _CACHE[key]
    db = rodb.connect(db_path)
    have = {r[0] for r in db.execute("SELECT name FROM sqlite_master WHERE type='table'")}
    sheets = _q(db, "SELECT rowid AS pos, sheet, state, layout FROM sheets ORDER BY rowid")
    per = {r["sheet"]: r for r in _q(db, """SELECT sheet, COUNT(*) AS rows, SUM(n_formula) AS formulas,
                                            SUM(n_const) AS consts FROM rows GROUP BY sheet""")}
    for s in sheets:
        s["layout"] = _json(s["layout"]) or {}
        s.update({k: (per.get(s["sheet"]) or {}).get(k) or 0 for k in ("rows", "formulas", "consts")})
    edges = {"kinds": [], "flows": [], "most_read": [], "most_reading": []}
    if "edges" in have:
        edges["kinds"] = _q(db, "SELECT kind, COUNT(*) AS n FROM edges GROUP BY kind ORDER BY n DESC")
        edges["flows"] = _q(db, f"""SELECT src_sheet, dst_sheet, SUM(kind IN ({_LIVE_SQL})) AS live,
                                    SUM(kind NOT IN ({_LIVE_SQL})) AS other FROM edges
                                    WHERE src_sheet <> dst_sheet GROUP BY src_sheet, dst_sheet""")
        top = """SELECT e.{a}_sheet AS sheet, e.{a}_row AS row, r.label, r.units, COUNT(*) AS n,
                        COUNT(DISTINCT e.{b}_sheet) AS sheets
                 FROM edges e LEFT JOIN rows r ON r.sheet = e.{a}_sheet AND r.row = e.{a}_row
                 WHERE e.kind IN ({live}) GROUP BY e.{a}_sheet, e.{a}_row ORDER BY n DESC LIMIT 15"""
        edges["most_read"] = _q(db, top.format(a="dst", b="src", live=_LIVE_SQL))
        edges["most_reading"] = _q(db, top.format(a="src", b="dst", live=_LIVE_SQL))
    names = _q(db, "SELECT name, ref FROM names ORDER BY name") if "names" in have else []
    sections = _q(db, """SELECT sheet, section, COUNT(*) AS rows FROM rows WHERE section IS NOT NULL
                         AND section <> '' GROUP BY sheet, section ORDER BY sheet, MIN(row)""")
    db.close()
    out = {"sheets": sheets, "edges": edges, "names": names, "sections": sections,
           "totals": {"sheets": len(sheets), "rows": sum(s["rows"] for s in sheets),
                      "formulas": sum(s["formulas"] for s in sheets), "consts": sum(s["consts"] for s in sheets),
                      "edges": sum(k["n"] for k in edges["kinds"]), "names": len(names)}}
    with _LOCK:
        _CACHE[key] = out
    return out


def rows(db_path: str, sheet: str | None = None, q: str | None = None, only: list[tuple[str, int]] | None = None,
         limit: int = 200) -> list[dict]:
    """Line items in workbook order, filtered by sheet, a label / section search, or a list of (sheet, row); with
    how many rows each one reads and is read by."""
    where, args = [], []
    if sheet:
        where.append("r.sheet = ?")
        args.append(sheet)
    if q:
        like = "%" + q.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_") + "%"
        where.append("(r.label LIKE ? ESCAPE '\\' OR r.section LIKE ? ESCAPE '\\')")
        args += [like, like]
    db = rodb.connect(db_path)
    if only is not None:
        db.execute("CREATE TEMP TABLE pick(sheet TEXT, row INT)")
        db.executemany("INSERT INTO temp.pick VALUES (?,?)", only)
        where.append("EXISTS (SELECT 1 FROM temp.pick p WHERE p.sheet = r.sheet AND p.row = r.row)")
    has_edges = bool(db.execute("SELECT 1 FROM sqlite_master WHERE name='edges'").fetchone())
    count = (lambda a, b: f"""(SELECT COUNT(*) FROM edges e WHERE e.{a}_sheet=m.sheet AND e.{a}_row=m.row
                               AND e.kind IN ({_LIVE_SQL}))""") if has_edges else (lambda a, b: "0")
    # edges has no index, so pick the page of rows first and count edges only for those
    out = _q(db, f"""WITH m AS (SELECT * FROM rows r {"WHERE " + " AND ".join(where) if where else ""}
                                ORDER BY (SELECT rowid FROM sheets s WHERE s.sheet = r.sheet), r.row LIMIT ?)
                     SELECT m.sheet, m.row, m.section, m.label, m.units, m.n_formula, m.n_const, m.samples,
                            {count("src", "dst")} AS reads, {count("dst", "src")} AS read_by FROM m""",
             *args, min(max(limit, 1), 1000))
    db.close()
    return out
