"""Identifying a workbook's valuation date (bench/identify.py) where a large model mentions valuation dates in many
rows: a model built for this year's date that also rolls back to last year's has rows like "Roll forward
valuation date (to 30/9/2025)" holding last year's date, many of them, and one row labelled "Valuation Date"
holding its own. No model calls (the rules' pick, and the order the model is shown the candidates in).

    uv run python tests/check_identify.py
"""
import sys
import tempfile
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "bench"))
import xlsxwriter  # noqa: E402

import build_map  # noqa: E402
import identify  # noqa: E402


def main() -> None:
    out = Path(tempfile.mkdtemp(prefix="identify_"))
    path = out / "model.xlsx"
    wb = xlsxwriter.Workbook(path)
    dt = wb.add_format({"num_format": "dd-mmm-yy"})
    sens = wb.add_worksheet("Sens")  # listed first: fifty rows about last year's date
    for i in range(50):
        sens.write(i, 1, f"Roll forward valuation date (to 30/9/2025), case {i + 1}")
        sens.write_datetime(i, 3, date(2025, 9, 30), dt)
    val = wb.add_worksheet("Val")
    val.write(19, 1, "Valuation Date")
    val.write_datetime(19, 3, date(2025, 12, 31), dt)
    wb.close()
    db = build_map.main(str(path), str(out / "db"))["db"]
    c = identify.candidates(db)
    first = c["valuation_date"][0]
    assert first["where"] == "Val!D20" and first["value"] == "2025-12-31" and first["rank"] == 0, c["valuation_date"][:3]
    assert all(d["rank"] == 3 for d in c["valuation_date"][1:]), "the rows naming another date come after it"
    got = identify.fallback(c, "model.xlsx")
    assert got["valuation_date"] == "2025-12-31" and got["valuation_date_evidence"] == "Val!D20", got
    print("identify: ok (the row labelled exactly 'Valuation Date' first, and picked, over fifty rows naming another date)")


if __name__ == "__main__":
    main()
