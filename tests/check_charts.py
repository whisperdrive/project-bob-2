"""The report's charts (bench/reportcharts.py) on a synthetic report laid out as real ones are: two panels a page,
each with a title over a coloured rule, a chart per panel or a chart beside a commentary column. No model calls:
readings are given as the vision model would return them.

  finding      each panel's chart found on its own (not the page's two panels as one), titled by its panel, the
               commentary column left out
  time axis    a chart of categories (valuation ranges: "FY26-FY30") isn't a chart over years; "2025-26" is FY26
  matching     a series for a site is the sum of that site's rows on a quarterly sheet; the total line is a row on
               an annual sheet; both chart together (annual, from financial-year totals); a multiple (units x)
               isn't matched to a row a thousand times bigger; a series too small to read isn't matched
  this year    the same rows followed into this year's model, the groups too
  year end     a model of quarters only has no annual timeline to show its financial year: the engagement's (last
               year's valuation date's month) is used, not December, so a June-year model's rows total by June years

    uv run python tests/check_charts.py
"""
import sys
import tempfile
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "bench"))
import matplotlib  # noqa: E402

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import xlsxwriter  # noqa: E402
from matplotlib.backends.backend_pdf import PdfPages  # noqa: E402
from matplotlib.lines import Line2D  # noqa: E402
from matplotlib.patches import Rectangle  # noqa: E402

import build_map  # noqa: E402
import reportcharts as rc  # noqa: E402
import rodb  # noqa: E402

YEARS = list(range(2026, 2046))
SITES = {"North": [("North plaza A revenue", 21.0), ("North plaza B revenue", 9.0), ("North plaza costs", -6.0)],
         "South": [("South plaza revenue", 12.0), ("South kiosk revenue", 3.0)],
         "East": [("East plaza revenue", 7.5)]}
GROWTH = 0.03


def site_quarter(base: float, k: int, fy0: int) -> float:
    """One quarter's value: a quarter of the year's, growing each year, with a step down in the tenth year."""
    y = k // 4
    return base / 4 * (1 + GROWTH) ** y * (0.6 if fy0 + y >= fy0 + 9 else 1.0)


def client(path: Path, fy0: int) -> dict:
    """A client model: quarterly site revenue (each site a section of rows) and an annual summary sheet."""
    wb = xlsxwriter.Workbook(path)
    dt = wb.add_format({"num_format": "dd-mmm-yy"})
    q = wb.add_worksheet("Sites")
    q.write(0, 0, "Site revenue by quarter")
    q.write(2, 0, "Quarter ending")
    ends = []
    for k in range(4 * len(YEARS)):
        y, m = fy0 - 1 + (6 + 3 * (k + 1) - 1) // 12, (6 + 3 * (k + 1) - 1) % 12 + 1
        e = date(y, m, [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1])
        ends.append(e)
        q.write_datetime(2, 2 + k, e, dt)
    r = 4
    totals = {fy: 0.0 for fy in range(fy0, fy0 + len(YEARS))}
    for site, rows in SITES.items():
        q.write(r, 0, site)
        r += 1
        for label, base in rows:
            q.write(r, 0, label)
            q.write(r, 1, "A$m")
            for k in range(4 * len(YEARS)):
                v = site_quarter(base, k, fy0)
                q.write_number(r, 2 + k, v)
                if "revenue" in label:
                    totals[fy0 + k // 4] += v
            r += 1
    a = wb.add_worksheet("Annual")
    a.write(0, 0, "Annual summary")
    a.write(2, 0, "Year ending")
    for k in range(len(YEARS)):
        a.write_datetime(2, 2 + k, date(fy0 + k, 6, 30), dt)
    a.write(4, 0, "Total revenue")
    a.write(4, 1, "A$m")
    a.write(5, 0, "Parking bays")  # a decoy for the multiple: about a thousand times it
    for k in range(len(YEARS)):
        a.write_number(4, 2 + k, totals[fy0 + k])
        a.write_number(5, 2 + k, 10700.0 + 100 * (k % 3))
    wb.close()
    return {"totals": totals}


def site_years(site: str, fy0: int) -> dict:
    out = {}
    for label, base in SITES[site]:
        if "revenue" in label:
            for k in range(4 * len(YEARS)):
                out[fy0 + k // 4] = out.get(fy0 + k // 4, 0.0) + site_quarter(base, k, fy0)
    return out


L0, L1, R0, R1 = 0.05, 0.488, 0.508, 0.95  # two panels, a narrow gutter between (about 17pt)


def panel(fig, x0, x1, title, chart=True, notes=()):
    """A panel as the report draws one: a bold title over a coloured rule; the chart in a white box (as an Excel
    chart pasted into the page is) straight under it; bullet notes under the chart, the panel's full width."""
    fig.text(x0, 0.855, title, fontsize=9, weight="bold")
    fig.add_artist(Line2D([x0, x1], [0.847, 0.847], transform=fig.transFigure, color="#1a9afa", linewidth=1.6))
    if chart:
        fig.add_artist(Rectangle((x0, 0.30), x1 - x0, 0.535, transform=fig.transFigure, facecolor="white",
                                 edgecolor="#d9d9d9", linewidth=0.5, zorder=-10))
    for i, ln in enumerate(notes):
        fig.text(x0, 0.26 - 0.022 * i, "- " + ln, fontsize=7)


def axes(fig, x0, x1):
    ax = fig.add_axes([x0 + 0.035, 0.40, x1 - x0 - 0.045, 0.42])
    ax.yaxis.grid(True, color="#d9d9d9", linewidth=0.5)
    ax.set_axisbelow(True)
    for side in ("top", "right"):
        ax.spines[side].set_visible(False)
    return ax


def report(path: Path, fy0: int, totals: dict) -> None:
    plt.rcParams["pdf.fonttype"] = 42
    labels = [str(y) for y in YEARS]
    xs = range(len(YEARS))
    with PdfPages(path) as pdf:
        # page 1: a stacked chart by site with the total as a line, beside a commentary column
        fig = plt.figure(figsize=(11.69, 8.27))
        fig.text(0.05, 0.9, "3.4 Key figures", fontsize=16, weight="bold")
        panel(fig, L0, L1, "Revenue", notes=["Source: synthetic model"])
        ax = axes(fig, L0, L1)
        bottom = [0.0] * len(YEARS)
        for site, colour in zip(SITES, ("#2e2e38", "#1a9afa", "#ffe600")):
            ys = [site_years(site, fy0)[y] for y in YEARS]
            ax.bar(xs, ys, bottom=bottom, color=colour, label=site)
            bottom = [b + v for b, v in zip(bottom, ys)]
        ax.plot(xs, [totals[y] for y in YEARS], color="#747480", marker="o", markersize=2, label="Total revenue")
        ax.set_xticks(list(xs)); ax.set_xticklabels(labels, rotation=90, fontsize=6)
        ax.tick_params(axis="y", labelsize=6)
        ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.13), ncol=4, fontsize=6, frameon=False)
        panel(fig, R0, R1, "Commentary", chart=False)
        for i, ln in enumerate(["Revenue comes from three sites, North the largest with two plazas.",
                                "Revenue steps down in the tenth year when a lease ends.",
                                "The total is shown as a line over the stacked sites.",
                                "Since the prior valuation, revenue over the forecast has increased by 5%."]):
            fig.text(R0, 0.82 - 0.022 * i, "- " + ln, fontsize=7)
        pdf.savefig(fig); plt.close(fig)
        # page 2: two charts side by side
        fig = plt.figure(figsize=(11.69, 8.27))
        fig.text(0.05, 0.9, "3.4 Key figures (continued)", fontsize=16, weight="bold")
        for x0, x1, title, ys in ((L0, L1, "Operating expenses", [3 + 0.1 * k for k in xs]),
                                  (R0, R1, "Net debt", [200 - 9 * k for k in xs])):
            panel(fig, x0, x1, title, notes=["Source: synthetic model",
                                             f"{title} since the prior valuation have moved with the forecast."])
            ax = axes(fig, x0, x1)
            ax.bar(xs, ys, color="#1a9afa", label=title)
            ax.plot(xs, [v * 1.05 for v in ys], color="#2e2e38", label=title + " (last year)")
            ax.set_xticks(list(xs)); ax.set_xticklabels(labels, rotation=90, fontsize=6)
            ax.tick_params(axis="y", labelsize=6)
            ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.13), ncol=2, fontsize=6, frameon=False)
        pdf.savefig(fig); plt.close(fig)


def year_end_check(out: Path) -> None:
    import sqlite3
    import build_map
    import chartdata
    from datetime import timedelta
    path = out / "quarters_only.xlsx"
    wb = xlsxwriter.Workbook(path)
    dt = wb.add_format({"num_format": "dd-mmm-yy"})
    ws = wb.add_worksheet("Ops")
    ws.write(2, 1, "Period ending")
    ws.write(4, 1, "Revenue")
    y, m = 2025, 9
    for k in range(12):  # Sep-25 .. Jun-28: three June years, each 4 x 100
        ws.write_datetime(2, 3 + k, date(y + (m == 12), m % 12 + 1, 1) - timedelta(days=1), dt)
        ws.write_number(4, 3 + k, 100.0)
        y, m = (y + 1, 3) if m == 12 else (y, m + 3)
    wb.close()
    db = sqlite3.connect(build_map.main(str(path), str(out / "quarters_db"))["db"])
    assert chartdata.fy_end_detect(db)[0] is None, "a model of quarters shows no financial year (not December)"
    assert chartdata.fy_end_month(db) == 12, "no hint: December, as before"
    with chartdata.fy_end_hint(6):
        assert chartdata.fy_end_month(db) == 6
        row = next(r for r in rc.fy_totals(db) if r["label"] == "Revenue")
    assert row["years"] == {2026: 400.0, 2027: 400.0, 2028: 400.0}, row["years"]
    # an annual model of March years shows its own year end; a hint doesn't override it, a person's setting does
    path = out / "march_years.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Ops")
    ws.write(2, 1, "Period ending")
    for k in range(5):
        ws.write_datetime(2, 3 + k, date(2026 + k, 3, 31), wb.add_format({"num_format": "dd-mmm-yy"}))
    wb.close()
    db = sqlite3.connect(build_map.main(str(path), str(out / "march_db"))["db"])
    assert chartdata.fy_end_detect(db)[0] == 3
    with chartdata.fy_end_hint(6):
        assert chartdata.fy_end_month(db) == 3
    with chartdata.fy_end_hint(6, forced=True):
        assert chartdata.fy_end_month(db) == 6
    print("year end: ok (a quarterly model's rows total by the engagement's June years, not calendar years; a model's "
          "own year end beats the hint, a person's setting beats both)")


def main() -> None:
    out = Path(tempfile.mkdtemp(prefix="charts_"))
    fy0 = YEARS[0]
    made = client(out / "client_last_year.xlsx", fy0)
    report(out / "report.pdf", fy0, made["totals"])

    # ---- finding
    found = rc.find({"tables": []}, str(out / "report.pdf"), out)
    by_page = {}
    for c in found:
        by_page.setdefault(c["page"], []).append(c)
    for c in found:
        print(f"  page {c['page']}: {c['caption']!r} bbox {c['bbox']}")
    assert [c["caption"] for c in by_page[1]] == ["Revenue"], by_page[1]
    assert by_page[1][0]["bbox"][2] < R0 * 842, "the commentary column isn't part of the chart"
    assert sorted(c["caption"] for c in by_page[2]) == ["Net debt", "Operating expenses"], by_page[2]
    left, right = sorted(by_page[2], key=lambda c: c["bbox"][0])
    assert left["bbox"][2] < right["bbox"][0], "the two panels' charts don't overlap"
    # the report reader often cuts a drawn chart out as a figure: just its plot, no axis labels or legend.
    # The drawn chart's own crop takes its place.
    (out / "tables").mkdir(exist_ok=True)
    plt.imsave(out / "tables" / "p001-t1.png", [[0.5] * 40] * 30)
    doc = {"tables": [{"id": "p001-t1", "status": "figure", "source": "text-layer table", "page": 1,
                       "bbox": [65.5, 101.2, 408.3, 363.3], "png": "tables/p001-t1.png"}]}
    again = [c for c in rc.find(doc, str(out / "report.pdf"), out) if c["page"] == 1]
    assert [(c["source"], c["caption"]) for c in again] == [("drawn in the PDF", "Revenue")], again
    doc["tables"][0]["bbox"] = [36.0, 60.0, 412.0, 420.0]  # ... or the whole panel, title and legend too
    again = [c for c in rc.find(doc, str(out / "report.pdf"), out) if c["page"] == 1]
    assert [(c["source"], c["caption"]) for c in again] == [("drawn in the PDF", "Revenue")], again
    print("finding: ok")

    # ---- time axis
    assert rc._year("FY26") == 2026 and rc._year("2026") == 2026 and rc._year("2025-26") == 2026
    assert rc._year("FY25/26") == 2026 and rc._year("Jun-26") == 2026
    assert rc._year("FY26-FY30") is None and rc._year("FY26–FY30") is None and rc._year("0.70") is None
    assert rc.time_axis(["FY26-FY30", "FY31-FY35", "Peers"]) is None
    assert rc.time_axis(["0.40", "0.45", "0.35"]) is None
    assert rc.time_axis([str(y) for y in YEARS]) == YEARS
    print("time axis: ok")

    # ---- matching, on last year's model
    db = build_map.main(str(out / "client_last_year.xlsx"), str(out / "last_db"))["db"]
    books = [{"key": "prior_model", "name": "last year's client model", "db_path": db, "rows": rc.fy_totals(rodb.connect(db))}]

    class Reader:  # no model calls: every check here is by numbers
        model = reviewer_model = "none"

    noisy = lambda ys, k: [round(v, 2) for v in ys]  # as a PPTX chart's data would have them
    read = {"is_chart": True, "title": "Revenue", "kind": "stacked column", "units": "A$m",
            "x_labels": [str(y) for y in YEARS], "series": [
                {"name": site, "drawn_as": "bar", "values": noisy([site_years(site, fy0)[y] for y in YEARS], k)}
                for k, site in enumerate(SITES)] + [
                {"name": "Total revenue", "drawn_as": "line", "values": noisy([made["totals"][y] for y in YEARS], 5)},
                {"name": "Other", "drawn_as": "bar", "values": [0.3] * len(YEARS)}]}
    res = rc.recreate(Reader(), {"id": "t1", "png": None, "caption": "Revenue"}, read, books, out)
    first = res["tries"][0]["picks"]
    print("  picks:", [(p["sheet"], p.get("rows") or p["row"], p["label"]) if p else None for p in first])
    labels = dict(rodb.connect(db).execute("SELECT row, label FROM rows WHERE sheet='Sites'").fetchall())
    north = sorted(labels[r] for r in first[0]["rows"])
    assert north == ["North plaza A revenue", "North plaza B revenue"], north
    assert sorted(labels[r] for r in first[1]["rows"]) == ["South kiosk revenue", "South plaza revenue"]
    assert first[2]["label"] == "East plaza revenue" and not first[2].get("rows")
    assert first[3]["sheet"] == "Annual" and first[3]["label"] == "Total revenue"
    assert first[4] is None and res["series_notes"] == {4: "too small to read off the chart"}
    spec = res["spec"]
    assert spec["labels"][0] == "FY2026" and len(spec["labels"]) == len(YEARS)
    assert abs(spec["series"][0]["data"][0] - site_years("North", fy0)[fy0]) < 1e-6
    rc.render(spec)  # a quarterly sheet and an annual one, charted together: no shape mismatch
    cmp_ = res["compare"]["series"]
    assert cmp_[0]["model"] and abs(cmp_[0]["diff"][0]) < 0.03 and cmp_[4]["model"] is None
    print("matching: ok")

    # ---- a multiple isn't a thousand of something; categories aren't years
    mult = {"is_chart": True, "title": "EV / EBITDA", "kind": "column", "units": "x", "x_labels": [str(y) for y in YEARS],
            "series": [{"name": "EV / EBITDA", "drawn_as": "bar", "values": [10.7 + 0.1 * (k % 3) for k in range(len(YEARS))]}]}
    r2 = rc.recreate(Reader(), {"id": "t2", "png": None, "caption": "EV / EBITDA"}, mult, books, out)
    assert r2.get("problem") == "no model row follows any of its series", r2.get("candidates")
    ff = {**mult, "x_labels": ["FY26-FY30", "FY31-FY35", "Peers"],
          "series": [{"name": "Range", "drawn_as": "bar", "values": [10.7, 10.9, 11.0]}]}
    assert rc.recreate(Reader(), {"id": "t3", "png": None}, ff, books, out)["skipped"].startswith("not a chart over years")
    print("gates: ok")

    # ---- your rows: summed, at the size of the reading (not shrunk to fit), and how far off
    mine = rc.pick_rows(db, read, 0, "Sites", [6, 7])
    assert mine["scale"] == 1.0 and mine["error"] < 0.01 and mine["manual"], mine
    wrong = rc.pick_rows(db, read, 0, "Sites", [10, 11])
    assert wrong["scale"] == 1.0 and wrong["error"] > 0.3, wrong
    assert rc.numbers_verdict(read, [mine, None, None, None, None], {4: "too small"})["matches"] is False
    print("your rows: ok")

    # ---- this year: the same rows (and groups) in this year's model
    client(out / "client_this_year.xlsx", fy0 + 1)
    cdb = build_map.main(str(out / "client_this_year.xlsx"), str(out / "this_db"))["db"]
    cur = rc.current_spec(db, cdb, read, res["picks"], "Revenue (this year)")
    assert cur["labels"][0] == f"FY{fy0 + 1}", cur["labels"][:2]
    assert abs(cur["series"][0]["data"][0] - site_years("North", fy0 + 1)[fy0 + 1]) < 1e-6
    assert [s["name"] for s in cur["series"]] == ["North", "South", "East", "Total revenue"]
    print("this year: ok")
    year_end_check(out)
    print("charts: all checks passed")


if __name__ == "__main__":
    main()
