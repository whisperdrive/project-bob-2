"""Formula census: per-sheet counts of formulas / constants, hidden state, load time.

Usage: census.py <workbook.xlsx|.xlsm>

Uses openpyxl read-only streaming so the full workbook never sits in memory.
"""
import json
import os
import sys
import time
from collections import Counter
from datetime import datetime

import openpyxl


def main(path: str) -> None:
    t0 = time.time()
    wb = openpyxl.load_workbook(path, read_only=True, data_only=False)
    t_open = time.time() - t0

    sheets = []
    for ws in wb.worksheets:
        t = time.time()
        kinds = Counter()
        date_count_by_row = Counter()
        max_r = max_c = 0
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if v is None:
                    continue
                max_r, max_c = max(max_r, c.row), max(max_c, c.column)
                if isinstance(v, str) and v.startswith("="):
                    kinds["formula"] += 1
                elif isinstance(v, datetime):
                    kinds["date"] += 1
                    date_count_by_row[c.row] += 1
                elif isinstance(v, (int, float)):
                    kinds["number"] += 1
                else:
                    kinds["text"] += 1
        header_row = date_count_by_row.most_common(1)
        sheets.append({
            "sheet": ws.title,
            "state": ws.sheet_state,
            "max_row": max_r,
            "max_col": max_c,
            **kinds,
            "date_header_row": header_row[0] if header_row else None,
            "secs": round(time.time() - t, 2),
        })
        print(json.dumps(sheets[-1]), file=sys.stderr)

    result = {
        "open_secs": round(t_open, 2),
        "total_secs": round(time.time() - t0, 2),
        "defined_names": len(wb.defined_names),
        "sheets": sheets,
    }
    out = os.path.join("out", os.path.splitext(os.path.basename(path))[0])
    os.makedirs(out, exist_ok=True)
    with open(os.path.join(out, "census.json"), "w", encoding="utf-8") as f:
        json.dump(result, f, indent=2, default=str)
    print(json.dumps({k: v for k, v in result.items() if k != "sheets"}))


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    main(sys.argv[1])
