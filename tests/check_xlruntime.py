"""Excel functions in the Python runtime (bench/xlruntime.py) where a model's labels or figures depend on them.

  TEXT      number formats, rounded half up as Excel does (1234.5 -> "1235", not Python's "1234"); date formats
            of day, month and year codes ("dd mmm yyyy"): a label like ="Valuation date: "&TEXT(F7,"dd mmm yyyy")
            reads as Excel shows it, not as the date's serial number

    uv run python tests/check_xlruntime.py
"""
import sys
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "bench"))
from xlruntime import serial, xl  # noqa: E402


def main() -> None:
    d = serial(date(2025, 6, 30))  # a Monday
    cases = [(d, "dd mmm yyyy", "30 Jun 2025"), (d, "d/mm/yy", "30/06/25"), (d, "mmmm yyyy", "June 2025"),
             (d, "yyyy-mm-dd", "2025-06-30"), (d, "dddd d mmmm", "Monday 30 June"), (d, "ddd", "Mon"),
             (d, '"FY"yy', "FY25"), (d, "mmm-yy", "Jun-25"),
             (0.0725, "0.00%", "7.25%"), (1234.5, "#,##0.0", "1,234.5"), (1234.5, "0", "1235"), (2.5, "0", "3"),
             (0.00125, "0.00%", "0.13%"), (-1234.5, "#,##0", "-1,235")]
    for v, fmt, want in cases:
        got = xl.TEXT(v, fmt)
        assert got == want, (v, fmt, got, want)
    print(f"TEXT: ok ({len(cases)} formats: dates by their day, month and year codes, numbers rounded half up "
          "as Excel does)")


if __name__ == "__main__":
    main()
