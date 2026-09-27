"""Head-to-head: load time, formula support, and token cost of each reading approach.

Usage: compare.py <workbook.xlsx|.xlsm> [other copies of the same workbook, e.g. .xlsb]

Token counts use tiktoken cl100k_base as a proxy (Claude's tokenizer differs; treat as approximate).
Run build_map.py on the .xlsx/.xlsm first.
"""
import io
import json
import os
import re
import sqlite3
import sys
import time
from collections import defaultdict

import tiktoken

sys.path.insert(0, "bench")
import rodb  # noqa: E402
import tools  # noqa: E402

ENC = tiktoken.get_encoding("cl100k_base")


def tok(s: str) -> int:
    return len(ENC.encode(s, disallowed_special=()))


def m_calamine(path):
    from python_calamine import CalamineWorkbook
    t = time.time()
    wb = CalamineWorkbook.from_path(path)
    buf = io.StringIO()
    for name in wb.sheet_names:
        buf.write(f"## {name}\n")
        for i, row in enumerate(wb.get_sheet_by_name(name).to_python()):
            vals = [str(v) for v in row if v not in ("", None)]
            if vals:
                buf.write(f"{i + 1}: " + ",".join(vals) + "\n")
    return time.time() - t, buf.getvalue(), "values only"


def m_pyxlsb(path):
    from pyxlsb import open_workbook
    t = time.time()
    buf = io.StringIO()
    with open_workbook(path) as wb:
        for name in wb.sheets:
            buf.write(f"## {name}\n")
            with wb.get_sheet(name) as s:
                for row in s.rows(sparse=True):
                    vals = [str(c.v) for c in row if c.v is not None]
                    if vals:
                        buf.write(f"{row[0].r + 1}: " + ",".join(vals) + "\n")
    return time.time() - t, buf.getvalue(), "values only"


def m_openpyxl(path):
    import openpyxl
    t = time.time()
    wb = openpyxl.load_workbook(path, read_only=True, data_only=False)
    buf = io.StringIO()
    for ws in wb.worksheets:
        buf.write(f"## {ws.title}\n")
        for row in ws.iter_rows():
            for c in row:
                if c.value is not None:
                    buf.write(f"{c.coordinate}={c.value}\n")
    return time.time() - t, buf.getvalue(), "formulas (no values)"


def top_output() -> tuple[str, int, str]:
    """The line item nothing depends on with the largest upstream tree - usually the model's headline output."""
    db = rodb.connect(tools.DB)
    up = defaultdict(list)
    has_dependents = set()
    for ss, sr, ds, dr in db.execute("SELECT src_sheet, src_row, dst_sheet, dst_row FROM edges"):
        up[(ss, sr)].append((ds, dr))
        has_dependents.add((ds, dr))
    # model-check rows (sum of every error flag) have huge trees but aren't business outputs
    checks = {(s, r) for s, r, lab in db.execute("SELECT sheet, row, label FROM rows")
              if re.search(r"check|error|integrity", lab or "", re.I)}
    best, best_n = None, -1
    for node in up:
        if node in has_dependents or node in checks or any(p in checks for p in up[node]):
            continue
        seen, stack = {node}, [node]
        while stack:
            for nxt in up.get(stack.pop(), ()):
                if nxt not in seen:
                    seen.add(nxt)
                    stack.append(nxt)
        if len(seen) > best_n:
            best, best_n = node, len(seen)
    if best is None:
        raise SystemExit("no formula dependencies found - nothing to trace")
    label = db.execute("SELECT label FROM rows WHERE sheet=? AND row=?", best).fetchone()[0] or ""
    return best[0], best[1], label


def main(paths: list[str]) -> None:
    formula_file = next((p for p in paths if p.lower().endswith((".xlsx", ".xlsm"))), None)
    if not formula_file:
        sys.exit("pass at least one .xlsx/.xlsm (the one build_map.py was run on)")
    tools.use(formula_file)
    out = os.path.dirname(tools.DB)

    results = []
    for path in paths:
        ext = os.path.splitext(path)[1].lower()
        methods = [("calamine", m_calamine)]
        methods += [("pyxlsb", m_pyxlsb)] if ext == ".xlsb" else []
        methods += [("openpyxl", m_openpyxl)] if ext in (".xlsx", ".xlsm") else []
        for name, fn in methods:
            secs, text, kind = fn(path)
            results.append({"method": f"{name} ({ext}) - raw dump", "secs": round(secs, 2),
                            "content": kind, "tokens": tok(text)})
            print(json.dumps(results[-1]), flush=True)

    full_map = open(os.path.join(out, "map.txt"), encoding="utf-8").read()
    results.append({"method": "row map - full (map.txt)", "secs": "build once",
                    "content": "formulas as R1C1 patterns + sample values", "tokens": tok(full_map)})

    # A scripted session on the auto-detected headline output: "what drives <top output>?"
    sheet, row, label = top_output()
    word = next((w for w in re.findall(r"[A-Za-z]{4,}", label)), label)
    session = [
        ("overview()", tools.overview()),
        (f"find({word!r})", tools.find(word)),
        (f"rows({sheet!r},{row})", tools.rows(sheet, row)),
        (f"trace({sheet!r},{row},'up',2)", tools.trace(sheet, row, "up", 2)),
    ]
    steps = [{"call": c, "tokens": tok(r)} for c, r in session]
    results.append({"method": f"overview + tools - 'what drives {sheet}!r{row} {label}?'",
                    "secs": "<0.1 per call", "content": "formulas + values, on demand",
                    "tokens": sum(st["tokens"] for st in steps), "steps": steps})
    print(json.dumps(results[-2:], indent=1))

    with open(os.path.join(out, "compare.json"), "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    main(sys.argv[1:])
