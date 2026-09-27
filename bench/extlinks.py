"""External workbook links: which other files a workbook reads, which of their cells, and the values it last saw.

Excel stores a formula like ='C:\\models\\[Client model.xlsx]CashFlow'!D9 as =[1]CashFlow!D9, with the file
behind [1] and a cached copy of every referenced cell in xl/externalLinks/externalLinkN.xml. That is an
overlay-to-client-model map already sitting in the file, readable without the client workbook. When the
client workbook is available too, the cached values show whether it is the exact version the overlay read.

Adds three tables to model.db:
  extbooks(idx, target, filename, sheets, n_cached)      one row per linked file; idx is the n in [n]
  extcells(idx, sheet, addr, row, col, value)            the cached values
  extrefs(sheet, row, idx, ext_sheet, ext_row, n_cells)  which line items read which external rows
    uv run python bench/extlinks.py <workbook.xlsx|xlsm> <model.db>
"""
import json
import re
import sqlite3
import sys
import zipfile
from pathlib import PurePosixPath, PureWindowsPath
from urllib.parse import unquote
from xml.etree import ElementTree as ET

from openpyxl.utils import column_index_from_string

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "p": "http://schemas.openxmlformats.org/package/2006/relationships"}
R_ID = "{%s}id" % NS["r"]
# [1]Sheet!A1, '[1]My sheet'!A1:B2, [1]Sheet!$A$1 ; [1]!Name (an external defined name) is counted without a row.
REF = re.compile(r"(?:'\[(\d+)\]((?:[^']|'')*)'|\[(\d+)\]([A-Za-z0-9_.]*))!(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)?")
ADDR = re.compile(r"\$?([A-Z]{1,3})\$?(\d+)")
MAX_ROWS_PER_REF = 400


def _rels(z: zipfile.ZipFile, part: str) -> dict[str, str]:
    p = PurePosixPath(part)
    name = str(p.parent / "_rels" / (p.name + ".rels"))
    if name not in z.namelist():
        return {}
    return {r.get("Id"): r.get("Target") for r in ET.fromstring(z.read(name)).findall("p:Relationship", NS)}


def _filename(target: str) -> str:
    """'file:///C:\\x\\My%20model.xlsm' or '../x/My model.xlsm' -> 'My model.xlsm'."""
    t = unquote(target or "").replace("file:///", "").replace("file://", "")
    return PureWindowsPath(t).name if "\\" in t else PurePosixPath(t).name


def read_links(path: str) -> list[dict]:
    """[{idx, target, filename, sheets, cells: [(sheet, addr, value)]}] in [n] order."""
    with zipfile.ZipFile(path) as z:
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        wb_rels = _rels(z, "xl/workbook.xml")
        out = []
        refs = wb.find("m:externalReferences", NS)
        for idx, ref in enumerate(refs if refs is not None else [], start=1):
            part = "xl/" + wb_rels.get(ref.get(R_ID), "").lstrip("/").removeprefix("xl/")
            if part not in z.namelist():
                out.append({"idx": idx, "target": None, "filename": None, "sheets": [], "cells": []})
                continue
            root = ET.fromstring(z.read(part))
            book = root.find("m:externalBook", NS)
            if book is None:  # DDE / OLE link, not a workbook
                kind = root[0].tag.split("}")[-1] if len(root) else "unknown"
                out.append({"idx": idx, "target": kind, "filename": None, "sheets": [], "cells": []})
                continue
            target = _rels(z, part).get(book.get(R_ID), "")
            sheets = [s.get("val") for s in book.findall("m:sheetNames/m:sheetName", NS)]
            cells = []
            for sd in book.findall("m:sheetDataSet/m:sheetData", NS):
                sid = int(sd.get("sheetId", -1))
                sheet = sheets[sid] if 0 <= sid < len(sheets) else f"#{sid}"
                for c in sd.iter("{%s}cell" % NS["m"]):
                    v = c.find("m:v", NS)
                    val = v.text if v is not None else None
                    if val is not None and c.get("t") not in ("s", "str", "e", "b"):
                        try:
                            val = float(val)
                        except ValueError:
                            pass
                    cells.append((sheet, c.get("r"), val))
            out.append({"idx": idx, "target": target, "filename": _filename(target), "sheets": sheets, "cells": cells})
    return out


def build(path: str, db_path: str) -> dict:
    """(Re)write the extbooks / extcells / extrefs tables in model.db from the workbook file."""
    links = read_links(path) if path.lower().endswith((".xlsx", ".xlsm")) else []
    db = sqlite3.connect(db_path)
    db.executescript("""
        DROP TABLE IF EXISTS extbooks; DROP TABLE IF EXISTS extcells; DROP TABLE IF EXISTS extrefs;
        CREATE TABLE extbooks(idx INT, target TEXT, filename TEXT, sheets TEXT, n_cached INT);
        CREATE TABLE extcells(idx INT, sheet TEXT, addr TEXT, row INT, col INT, value);
        CREATE TABLE extrefs(sheet TEXT, row INT, idx INT, ext_sheet TEXT, ext_row INT, n_cells INT);""")
    for b in links:
        db.execute("INSERT INTO extbooks VALUES (?,?,?,?,?)",
                   (b["idx"], b["target"], b["filename"], json.dumps(b["sheets"]), len(b["cells"])))
        rows = []
        for sheet, addr, val in b["cells"]:
            m = ADDR.fullmatch(addr or "")
            if m:
                rows.append((b["idx"], sheet, addr, int(m.group(2)), column_index_from_string(m.group(1)), val))
        db.executemany("INSERT INTO extcells VALUES (?,?,?,?,?,?)", rows)

    refs: dict[tuple, int] = {}
    for sheet, row, formula in db.execute("SELECT sheet, row, formula FROM cells WHERE formula LIKE '%[%]%'"):
        for m in REF.finditer(formula):
            idx = int(m.group(1) or m.group(3))
            ext_sheet = (m.group(2) or m.group(4) or "").replace("''", "'") or None
            ext_rows = [None]
            if m.group(5):
                ends = [int(a.group(2)) for a in ADDR.finditer(m.group(5))]
                ext_rows = list(range(min(ends), min(max(ends), min(ends) + MAX_ROWS_PER_REF) + 1))
            for er in ext_rows:
                refs[(sheet, row, idx, ext_sheet, er)] = refs.get((sheet, row, idx, ext_sheet, er), 0) + 1
    db.executemany("INSERT INTO extrefs VALUES (?,?,?,?,?,?)", [(*k, n) for k, n in refs.items()])
    db.execute("CREATE INDEX IF NOT EXISTS ix_extrefs ON extrefs(sheet, row)")
    db.commit()
    stats = {"books": len(links), "cached_cells": sum(len(b["cells"]) for b in links),
             "rows_reading": len({(k[0], k[1]) for k in refs}), "refs": sum(refs.values())}
    db.close()
    return stats


def ensure(path: str, db_path: str) -> None:
    """Build the tables once per model.db."""
    with sqlite3.connect(db_path) as db:
        have = {r[0] for r in db.execute("SELECT name FROM sqlite_master WHERE type='table'")}
    if "extbooks" not in have:
        build(path, db_path)


def summary(db_path: str) -> list[dict]:
    """Per linked file: its sheets, which of this workbook's sheets read it, and how many line items."""
    db = sqlite3.connect(f"file:{db_path}?mode=ro", uri=True)
    out = []
    for idx, target, filename, sheets, n_cached in db.execute("SELECT * FROM extbooks ORDER BY idx"):
        by_sheet = db.execute("""SELECT sheet, COUNT(DISTINCT row), COUNT(DISTINCT ext_sheet) FROM extrefs
                                 WHERE idx=? GROUP BY sheet ORDER BY 2 DESC""", (idx,)).fetchall()
        out.append({"idx": idx, "filename": filename, "target": target, "sheets": json.loads(sheets or "[]"),
                    "cached_cells": n_cached, "read_by": [{"sheet": s, "rows": n, "ext_sheets": k} for s, n, k in by_sheet],
                    "rows_reading": sum(n for _, n, _ in by_sheet)})
    db.close()
    return out


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    print(build(sys.argv[1], sys.argv[2]))
    for b in summary(sys.argv[2]):
        print(f"[{b['idx']}] {b['filename']}: {len(b['sheets'])} sheets, {b['cached_cells']} cached cells, "
              f"read by {b['rows_reading']} line items in {[r['sheet'] for r in b['read_by']]}")
