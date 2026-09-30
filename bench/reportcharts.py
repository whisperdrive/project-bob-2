"""The prior report's charts, recreated from the models and checked by eye.

Last year's report charted last year's client model (the "current" model at the time), so drawing the same rows
from the prior client model has to reproduce each chart: that validates the rows found. The same rows in this
year's client model give this year's version of the chart.

  find(doc, pdf_path, out_dir)  the report's charts: pictures the table reader classed as figures, PPTX charts
                                (their data is in the file), and charts drawn in the PDF itself (vector paths: bars
                                sharing a baseline, or plotted lines), cropped to PNG. Code only, no model call.
  digitise(reader, png)         a vision model reads a chart: title, kind, units, x labels, each series' name,
                                how it's drawn and its values (approximate: read off the picture)
  match(series, years, books)   the model rows whose values, totalled by financial year and scaled for units and
                                sign, follow a series within MATCH_TOL; best fit first, label as tie-break
  spec_for(...)                 a chart spec from those rows over the report's years, for the page and the render
  check(reader, report, ours)   a vision model compares the report's chart with the recreation (data, not looks);
                                the next candidate rows are tried for a series it rejects, twice at most
run() does all of it for one engagement and returns what the page shows.
"""
import hashlib
import io
import json
import math
import re
from datetime import date
from pathlib import Path

import dcf
import rodb

MATCH_TOL = 0.06    # mean absolute error, as a share of the series' size: values read off a picture are approximate
MAX_TRIES = 3       # recreations checked per chart (the best rows, then the next candidates)
SCALES = (1.0, 1e-3, 1e3, 1e-6, 1e6)
DPI = 200

DIGITISE_PROMPT = """This image is cut from {where} of last year's valuation report. If it is a chart, read it:
- title: its caption or title as printed ("" if none)
- kind: column, stacked column, line, area, combo (columns and lines together), waterfall, pie or other
- units: as printed on the axis, the title or the caption (e.g. "A$m"); "" if none
- x_labels: every category along the x axis, in order, as printed. Where only every second or fifth label is
  printed, fill in the ones between from the sequence (FY26, FY27, ...), one per bar or point.
- series: one per legend entry (or one unnamed series): its name as in the legend, drawn_as (bar, line or area),
  and values: one per x label, read off the chart as accurately as you can (use printed data labels where there
  are any; null where the series has no point). Negative values below the axis are negative.
If it isn't a chart (a table, photo, logo or diagram), set is_chart false and leave the rest empty."""
CHECK_PROMPT = """Image 1 is a chart from last year's valuation report. Image 2 is our recreation of it from the
valuation model's own data. Colours, fonts and styling differ by design: judge the data only. Compare:
- the series: the same measures (by meaning), the same number of them
- the x categories: the same years, in the same order
- the values: each series' bars or points at the same heights, within about 5% (read both images; the exact
  values of the recreation and the values read from the report are below)
- the chart family: columns, stacked columns, lines, or a mix
Say whether the recreation matches, and for each series whether it matches and what differs.

Report chart, as read: {read}
Recreation, exact values: {ours}"""
_S = {"type": "string"}
_NUMS = {"type": "array", "items": {"type": ["number", "null"]}}
_DIGITISE = {"type": "json_schema", "name": "chart_read", "strict": True, "schema": {
    "type": "object", "additionalProperties": False, "required": ["is_chart", "title", "kind", "units", "x_labels", "series"],
    "properties": {"is_chart": {"type": "boolean"}, "title": _S, "kind": _S, "units": _S,
                   "x_labels": {"type": "array", "items": _S},
                   "series": {"type": "array", "items": {"type": "object", "additionalProperties": False,
                                                         "required": ["name", "drawn_as", "values"],
                                                         "properties": {"name": _S, "values": _NUMS,
                                                                        "drawn_as": {"type": "string", "enum": ["bar", "line", "area"]}}}}}}}
_CHECK = {"type": "json_schema", "name": "chart_check", "strict": True, "schema": {
    "type": "object", "additionalProperties": False, "required": ["matches", "series", "differences"],
    "properties": {"matches": {"type": "boolean"}, "differences": {"type": "array", "items": _S},
                   "series": {"type": "array", "items": {"type": "object", "additionalProperties": False,
                                                         "required": ["name", "matches", "note"],
                                                         "properties": {"name": _S, "matches": {"type": "boolean"}, "note": _S}}}}}}


# ---- finding the charts --------------------------------------------------------------------------------------

def _overlaps(a, b, pad: float = 0) -> bool:
    return not (a[2] + pad < b[0] or b[2] + pad < a[0] or a[3] + pad < b[1] or b[3] + pad < a[1])


def _share(a, b) -> float:
    """How much of box a lies inside box b."""
    w = max(0.0, min(a[2], b[2]) - max(a[0], b[0]))
    h = max(0.0, min(a[3], b[3]) - max(a[1], b[1]))
    return w * h / max(1e-9, (a[2] - a[0]) * (a[3] - a[1]))


def _union(boxes):
    return [min(b[0] for b in boxes), min(b[1] for b in boxes), max(b[2] for b in boxes), max(b[3] for b in boxes)]


BAND = (60, 45, 60, 75)  # how far round a chart's bars and lines its axis labels and legend reach: left, up, right, down
GAP = 14                 # an empty vertical strip this wide between two things puts them in different panels
UNITS_ONLY = re.compile(r"^[\s$A-Z]{0,4}[$%]?[mbk]?[\s()]*$|^[\d\s.,()%$x-]+$", re.I)


def _dist(a, b) -> float:
    dx = max(0.0, b[0] - a[2], a[0] - b[2])
    dy = max(0.0, b[1] - a[3], a[1] - b[3])
    return (dx * dx + dy * dy) ** 0.5


def _xspan(seed, items, gap: float = GAP) -> tuple[float, float]:
    """The run of x covered by the chart and what sits next to it with no empty strip of gap or more between:
    the chart's panel, whatever is beside it (another chart, a commentary column) left out."""
    ivs = sorted([(seed[0], seed[2])] + [(b[0], b[2]) for b in items])
    runs = []
    for a, b in ivs:
        if runs and a <= runs[-1][1] + gap:
            runs[-1][1] = max(runs[-1][1], b)
        else:
            runs.append([a, b])
    mid = (seed[0] + seed[2]) / 2
    return next(((a, b) for a, b in runs if a <= mid <= b), (seed[0], seed[2]))


def _rules(page) -> list[list[float]]:
    """Horizontal rules: lines, and filled rectangles thinner than 3pt."""
    box = lambda o: [float(o["x0"]), float(o["top"]), float(o["x1"]), float(o["bottom"])]
    out = [box(l) for l in page.lines if abs(float(l["bottom"]) - float(l["top"])) < 1.5]
    out += [box(r) for r in page.rects if float(r["bottom"]) - float(r["top"]) < 3 and float(r["x1"]) - float(r["x0"]) > 20]
    return out


def _drawn_charts(page, skip: list) -> list[tuple[list[float], list[float]]]:
    """Regions of a PDF page drawn as a chart: four or more filled bars sharing a baseline (or top), or plotted
    lines with several points, not inside a table region or a picture. Returns (bbox, plot): the bbox takes in
    the axis labels and legend round the plot, within BAND, but nothing across an empty strip (GAP) and nothing
    nearer another chart; a rule over the chart (a panel's header line) bounds it too."""
    W, H = float(page.width), float(page.height)
    box = lambda o: [float(o["x0"]), float(o["top"]), float(o["x1"]), float(o["bottom"])]
    free = lambda b: not any(_overlaps(b, s, -1) for s in skip)
    bars = [box(r) for r in page.rects if r.get("fill") and (r["x1"] - r["x0"]) < 0.5 * W and (r["bottom"] - r["top"]) < 0.6 * H
            and (r["x1"] - r["x0"]) > 1 and (r["bottom"] - r["top"]) >= 0.5 and free(box(r))]
    seeds = []
    for key in (3, 1):  # bars standing on a common baseline, or hanging from one (negative values)
        by: dict[int, list] = {}
        for b in bars:
            by.setdefault(round(b[key]), []).append(b)
        clusters = []
        for g in by.values():
            # the chart's own background sits on the same baseline: leave out what is far wider than a bar
            med = sorted(b[2] - b[0] for b in g)[len(g) // 2]
            g = sorted((b for b in g if b[2] - b[0] <= 2.5 * med + 1), key=lambda b: b[0])
            # two charts side by side can share a baseline: split where the gap is far wider than the spacing
            gaps = sorted(max(0.0, b[0] - a[2]) for a, b in zip(g, g[1:]))
            cut = max(3 * (gaps[len(gaps) // 2] if gaps else 0), 3 * med, GAP)
            run = g[:1]
            for a, b in zip(g, g[1:]):
                if b[0] - a[2] > cut:
                    clusters.append(run)
                    run = []
                run.append(b)
            clusters.append(run)
        for g in clusters:
            widths = sorted(round(b[2] - b[0], 1) for b in g)
            if len(g) >= 4 and widths[-1] <= 2 * widths[0] + 1:
                region = _union(g)
                grown = True
                while grown:  # the parts stacked on those bars, up (or down) to the top of each column
                    grown = False
                    for b in bars:
                        if b[0] >= region[0] - 1 and b[2] <= region[2] + 1 and _overlaps(b, region, 1.5) \
                                and not (b[1] >= region[1] and b[3] <= region[3]):
                            region = _union([region, b])
                            grown = True
                seeds.append(region)
    for c in page.curves:
        b = box(c)
        if len(c.get("pts") or []) >= 5 and (b[2] - b[0]) > 60 and free(b):
            seeds.append(b)
    regions: list[list[float]] = []
    for sd in seeds:  # merge what overlaps (bars and the line drawn over them are one chart)
        for r in regions:
            if _overlaps(r, sd, 12):
                r[:] = _union([r, sd])
                break
        else:
            regions.append(list(sd))
    for r in regions:  # merging can bring two regions together: once more
        for o in regions:
            if o is not r and o and r and _overlaps(r, o, 12):
                r[:] = _union([r, o])
                o.clear()
    regions = [r for r in regions if r and (r[2] - r[0]) >= 80 and (r[3] - r[1]) >= 40]
    words = [box(w) for w in page.extract_words()]
    rules = _rules(page)
    out = []
    for r in regions:
        others = [o for o in regions if o is not r]
        band = [r[0] - BAND[0], r[1] - BAND[1], r[2] + BAND[2], r[3] + BAND[3]]
        near = lambda b: _overlaps(b, band) and all(_dist(b, r) <= _dist(b, o) for o in others)
        wide = (r[2] - r[0]) + BAND[0] + BAND[2]
        lines = [box(l) for l in page.lines if (float(l["x1"]) - float(l["x0"])) <= wide]
        items = [b for b in lines + words if near(b)]
        x0, x1 = _xspan(r, items)
        over = [u for u in rules if u[3] <= r[1] + 1 and r[1] - u[3] < 90 and _overlaps(u, [r[0], 0, r[2], H])
                and (u[2] - u[0]) >= 0.5 * (r[2] - r[0])]
        if over:  # a panel's header rule: the panel is no wider than it
            u = max(over, key=lambda u: u[3])
            x0, x1 = max(x0, u[0] - 2), min(x1, u[2] + 2)
        items = [b for b in items if b[0] >= x0 - 0.5 and b[2] <= x1 + 0.5]
        g = _union([r] + items)
        out.append(([max(0, g[0] - 6), max(0, g[1] - 6), min(W, g[2] + 6), min(H, g[3] + 6)], r))
    return out


def _caption(page, bbox, plot=None) -> str:
    """The chart's title: the line just above a rule over it (a panel's header), else the nearest line above
    it that isn't a number or a unit, bold ones first."""
    top = (plot or bbox)[1]
    x0, x1 = bbox[0], bbox[2]
    lines: dict[int, list] = {}
    for w in page.extract_words(extra_attrs=["fontname", "size"]):
        if w["bottom"] <= top + 2 and w["bottom"] > top - 90 and w["x1"] > x0 and w["x0"] < x1:
            lines.setdefault(round(w["top"]), []).append(w)
    cands = []
    for t, ws in lines.items():
        text = " ".join(w["text"] for w in sorted(ws, key=lambda w: w["x0"]))
        if len(text) < 3 or UNITS_ONLY.match(text):
            continue
        bold = any(re.search(r"bold|black|heavy|semibold", w.get("fontname") or "", re.I) for w in ws)
        cands.append((t, text, bold, max(w["bottom"] for w in ws)))
    if not cands:
        return ""
    rules = [u for u in _rules(page) if u[3] <= top + 1 and top - u[3] < 90 and u[0] < x1 and u[2] > x0
             and (u[2] - u[0]) >= 0.4 * (x1 - x0)]
    if rules:
        u = max(rules, key=lambda u: u[3])
        above = [c for c in cands if c[3] <= u[1] + 2 and u[1] - c[3] < 24]
        if above:
            return max(above, key=lambda c: c[3])[1]
    near = [c for c in cands if top - c[3] < 50]
    bold = [c for c in near if c[2]]
    pick = max(bold or near or cands, key=lambda c: c[3])
    return pick[1]


def find(doc: dict, pdf_path: str | None, out_dir: str | Path) -> list[dict]:
    """The report's charts: [{"id", "page", "source", "png" (relative to out_dir), "caption", "data" (PPTX)}]."""
    out_dir = Path(out_dir)
    (out_dir / "charts").mkdir(parents=True, exist_ok=True)
    found = []
    for t in doc.get("tables") or []:
        if t.get("status") == "figure" and t.get("png"):
            ch = {"id": t["id"], "page": t.get("page"), "source": "picture", "png": t["png"],
                  "caption": t.get("title") or t.get("description") or ""}
            try:
                from PIL import Image
                w, h = Image.open(out_dir / t["png"]).size
                if w > 8 * h or h > 8 * w or min(w, h) < 40:
                    ch["skipped"] = "a strip cut from a picture (an axis on its own), not a whole chart"
            except Exception:
                pass
            found.append(ch)
        elif t.get("source") == "pptx chart data" and (t.get("final_markdown") or t.get("markdown")):
            found.append({"id": t["id"], "page": t.get("page"), "source": "pptx chart", "png": None,
                          "caption": t.get("title") or "", "markdown": t.get("final_markdown") or t.get("markdown")})
    if pdf_path and str(pdf_path).lower().endswith(".pdf") and Path(pdf_path).exists():
        import pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            for ch in found:  # a picture's caption as the report prints it, above it on the page
                t = next((t for t in doc.get("tables") or [] if t["id"] == ch["id"]), {})
                if t.get("bbox") and ch.get("page") and 0 < ch["page"] <= len(pdf.pages):
                    ch["caption"] = _caption(pdf.pages[ch["page"] - 1], t["bbox"]) or ch["caption"]
            for pno, page in enumerate(pdf.pages, 1):
                skip = [t["bbox"] for t in doc.get("tables") or [] if t.get("page") == pno and t.get("bbox")
                        and t.get("source") != "page image" and t.get("status") != "figure"]
                skip += [[float(i["x0"]), float(i["top"]), float(i["x1"]), float(i["bottom"])] for i in page.images]
                drawn = _drawn_charts(page, skip)
                # a figure the report reader cut out of a drawn chart is often just its plot, without the axis
                # labels and legend it needs to be read: the drawn chart's own crop replaces it
                for bb, plot in drawn:
                    for ch in [c for c in found if c["page"] == pno and c["source"] == "picture"]:
                        t = next((t for t in doc.get("tables") or [] if t["id"] == ch["id"]), {})
                        fb = t.get("bbox")
                        if fb and _overlaps(fb, plot) and max(_share(fb, plot), _share(plot, fb)) >= 0.5:
                            found.remove(ch)
                for k, (bb, plot) in enumerate(drawn, 1):
                    cid = f"p{pno:03d}-c{k}"
                    rel = f"charts/{cid}.png"
                    img = page.crop(bb).to_image(resolution=DPI).original
                    buf = io.BytesIO()
                    img.convert("RGB").save(buf, format="PNG")
                    (out_dir / rel).write_bytes(buf.getvalue())
                    found.append({"id": cid, "page": pno, "source": "drawn in the PDF", "png": rel,
                                  "caption": _caption(page, bb, plot), "bbox": [round(v, 1) for v in bb]})
    return found


# ---- reading a chart -------------------------------------------------------------------------------------------

def digitise(reader, png: bytes, where: str) -> dict:
    """The vision model's reading. A dense chart (forty years of ten series) is a long answer: it gets room for
    one, is asked again once if the answer is cut short or isn't JSON, and says which it was if that fails too."""
    import base64
    from llm import create
    content = [{"type": "input_text", "text": DIGITISE_PROMPT.format(where=where)},
               {"type": "input_image", "image_url": f"data:image/png;base64,{base64.b64encode(png).decode()}",
                "detail": "high"}]
    why = ""
    for attempt in (1, 2):
        r = create(reader.llm, reader.model, text={"format": _DIGITISE}, max_output_tokens=24000,
                   purpose="report-chart-read" if attempt == 1 else "report-chart-read (again)",
                   input=[{"role": "user", "content": content}])
        if r.usage:
            reader.on_usage(reader.model, r.usage, "report-chart-read")
        if getattr(r, "status", None) == "incomplete":
            why = "the reading ran out of room (too many values in one chart)"
            continue
        try:
            return json.loads(r.output_text)
        except json.JSONDecodeError:
            why = "the reading came back empty or cut short"
    raise ValueError(why + ", twice")


def from_markdown(md: str) -> dict:
    """A PPTX chart's data table (categories down, series across) as a digitised chart: exact values."""
    import docingest
    rows = docingest.md_rows(md)
    if len(rows) < 2:
        return {"is_chart": False}
    head, body = rows[0], rows[1:]
    series = []
    for j, name in enumerate(head[1:], 1):
        vals = []
        for r in body:
            n = docingest.numbers(r[j]) if j < len(r) else []
            vals.append(float(n[0].rstrip("%")) if n else None)
        series.append({"name": name, "drawn_as": "bar", "values": vals})
    return {"is_chart": True, "title": "", "kind": "column", "units": "", "x_labels": [r[0] for r in body],
            "series": series, "exact": True}


# ---- matching series to model rows ----------------------------------------------------------------------------

def _year(label: str) -> int | None:
    """The year a label stands for: "FY26", "2026", "Jun-26" -> 2026; "2025-26" or "FY25/26" -> 2026 (the
    financial year ending then). None for a range of years ("FY26-FY30"), a decimal ("0.70") or no year."""
    s = str(label or "").strip()
    nums = re.findall(r"(?<![\d.,])(\d{4}|\d{2})(?![\d.,])", s)
    ys = [int(n) + 2000 if len(n) == 2 else int(n) for n in nums]
    ys = [y for y in ys if 1980 <= y <= 2120]
    if not ys:
        return None
    if len(ys) >= 2:
        return ys[1] if ys[1] == ys[0] + 1 else None  # 2025-26 is a financial year; FY26-FY30 is a range
    return ys[0]


def time_axis(labels: list) -> list | None:
    """The labels' years when the axis is years (three or more, rising, a year or a few apart), else None: a
    chart of categories (valuation ranges, peers, sites) isn't a chart over time."""
    ys = [_year(x) for x in labels or []]
    got = [y for y in ys if y]
    if len(set(got)) < 3 or len(got) < 0.6 * len(ys):
        return None
    steps = sorted(b - a for a, b in zip(got, got[1:]))
    if steps[0] <= 0 or steps[len(steps) // 2] > 5:
        return None
    return ys


def fy_totals(db, fy_m: int | None = None) -> list[dict]:
    """Every timeline row of a model, totalled by financial year: [{"sheet", "row", "label", "cols", "years": {FY:
    total}}]. Periods come from each sheet's own timeline."""
    import chartdata
    fy_m = fy_m or chartdata.fy_end_month(db)
    out = []
    for sheet, lay in db.execute("SELECT sheet, layout FROM sheets"):
        lay = json.loads(lay or "{}")
        if not lay.get("tl_first") or not lay.get("header_row"):
            continue
        cols = list(range(lay["tl_first"], lay["tl_last"] + 1))
        try:
            ends, _ = dcf.period_ends(db, sheet, cols)
        except ValueError:
            continue
        fy = {c: (e.year if e.month <= fy_m else e.year + 1) for c, e in ends.items()}
        labels, sections = {}, {}
        for r, lab, sec in db.execute("SELECT row, label, section FROM rows WHERE sheet=?", (sheet,)):
            labels[r], sections[r] = lab, sec
        rows: dict[int, dict] = {}
        for r, c, v in db.execute("SELECT row, col, value FROM cells WHERE sheet=? AND col BETWEEN ? AND ?",
                                  (sheet, cols[0], cols[-1])):
            if c in fy and isinstance(v, (int, float)) and not isinstance(v, bool) and r != lay["header_row"]:
                rows.setdefault(r, {}).setdefault(fy[c], 0.0)
                rows[r][fy[c]] += v
        for r, years in rows.items():
            if len(years) >= 3 and any(years.values()):
                out.append({"sheet": sheet, "row": r, "label": labels.get(r) or "", "section": sections.get(r) or "",
                            "cols": (cols[0], cols[-1]), "years": years})
    return out


def _words(s: str) -> set[str]:
    """Words of two letters or more, digits kept (site codes, acronyms): "NW3 plaza revenue" -> {nw3, plaza, revenue}."""
    return {w for w in re.findall(r"[a-z][a-z0-9]+", (s or "").lower()) if not re.fullmatch(r"(fy|cy|jun|dec)\d*", w)}


# words that say what is measured, not whose it is: a series named by them isn't a group of rows
GENERIC = {"total", "revenue", "revenues", "income", "sales", "cost", "costs", "expense", "expenses", "opex", "capex",
           "ebitda", "ebit", "debt", "net", "cash", "flow", "flows", "free", "opening", "closing", "balance", "forecast",
           "actual", "budget", "the", "and", "of", "for", "other", "operating", "capital", "expenditure", "maintenance",
           "interest", "tax", "equity", "value", "profit", "margin", "growth", "real", "nominal"}


def groups(rows: list[dict], name: str, title: str = "") -> list[dict]:
    """Rows that together make a series named for whose it is (a site, a segment): on one sheet, the rows whose
    section or label carries the name's own words, totals left out (they would count twice); and those of them
    whose label also has the chart's measure (its title's words). Each is summed by financial year."""
    own = _words(name) - GENERIC - _words(title)
    if not own:
        return []
    measure = _words(title) & GENERIC - {"total", "net", "the", "and", "of", "for"}
    by_sheet: dict[str, list] = {}
    for r in rows:
        if own & _words(r["section"] + " " + r["label"]) and "total" not in _words(r["label"]):
            by_sheet.setdefault(r["sheet"], []).append(r)
    out = []
    for sheet, members in by_sheet.items():
        sets = [members]
        if measure:
            sets.append([m for m in members if measure & _words(m["label"])])
        for ms in sets:
            if not 2 <= len(ms) <= 60 or any(o["rows"] == [m["row"] for m in ms] for o in out):
                continue
            years: dict = {}
            for m in ms:
                for y, v in m["years"].items():
                    years[y] = years.get(y, 0.0) + v
            out.append({"sheet": sheet, "row": ms[0]["row"], "rows": [m["row"] for m in ms], "section": "",
                        "label": f"{' '.join(sorted(own))} ({len(ms)} rows)", "cols": ms[0]["cols"], "years": years,
                        "group": True})
    return out


UNSCALED = re.compile(r"^\s*(x|times|%|per ?cent|percent)\s*$|\bmultiple\b|/ ?ebitda|\bratio\b", re.I)
NO_NAME_TOL = 0.02   # a row that shares no word with the series must follow it this closely ...
NO_NAME_PTS = 6      # ... over at least this many years: numbers alone match by coincidence otherwise


MIXED = re.compile(r"[;/,]|\band\b", re.I)
CURRENCY = re.compile(r"[$€£¥]|\b(?:m|bn|k|mn|million|billion|thousands?|000s?)\b", re.I)


def _scales(units: str, name: str) -> tuple:
    """Unit scales a series may need: none for a multiple (10.9x is never 10,900 of anything), percentages as
    fractions for a %, else thousands and millions either way. A chart whose units mix a currency and a % ("A$m;
    %": columns in A$m, a line in %) doesn't say which a series is: all of them, unless its name does."""
    u = f"{units} {name}"
    pct = re.compile(r"%|per ?cent", re.I)
    if pct.search(units or "") and CURRENCY.search(units or "") and MIXED.search(units or "") and not pct.search(name or ""):
        return SCALES + (100.0, 0.01)
    if re.search(r"%|per ?cent", u, re.I):
        return (1.0, 100.0, 0.01)
    if UNSCALED.search(units or "") or UNSCALED.search(name or ""):
        return (1.0,)
    return SCALES


def _rounding(values: list) -> float:
    """How much error the read's own rounding explains, as a share of its size: where most values are read to two
    significant figures (490, 1,770, 38), half a step each; 0 for a read more precise than that."""
    vs = [abs(v) for v in values if isinstance(v, (int, float)) and v]
    if not vs:
        return 0.0
    steps = [10 ** (math.floor(math.log10(v)) - 1) for v in vs]
    if sum(abs(v / st - round(v / st)) < 1e-9 for v, st in zip(vs, steps)) < 0.8 * len(vs):
        return 0.0
    return sum(st / 2 for st in steps) / sum(vs)


STOP = {"and", "of", "the", "for", "&"}


def _acronym(name: str, label: str) -> bool:
    """A label's "O&M" or "R/M" for a series named "Operations and maintenance": its content words' initials."""
    words = [w for w in re.findall(r"[a-z]+", (name or "").lower()) if w not in STOP]
    if not 2 <= len(words) <= 4:
        return False
    initials = "".join(w[0] for w in words)
    return initials in {"".join(m) for m in re.findall(r"\b([a-z])\s*[&/]\s*([a-z])\b", (label or "").lower())}


def match(values: list, years: list, rows: list[dict], name: str = "", limit: int = 5, units: str = "",
          tol: float = MATCH_TOL) -> list[dict]:
    """Rows that follow a series: mean absolute error over its years, as a share of its size, after the best of
    the unit scales it may need and either sign; under tol, widened by what the read's own rounding explains (at
    most tol again). A row whose label and section share no word with the series' name (nor its initials: O&M for
    operations and maintenance) must follow it within NO_NAME_TOL over NO_NAME_PTS years or more. Best fit first;
    the label decides near-ties."""
    pts = [(y, v) for y, v in zip(years, values) if y is not None and isinstance(v, (int, float))]
    if len(pts) < 3:
        return []
    size = sum(abs(v) for _, v in pts) or 1.0
    want = _words(name)
    scales = _scales(units, name)
    within = tol + min(tol, _rounding([v for _, v in pts]))
    out = []
    for r in rows:
        got = [(r["years"].get(y), v) for y, v in pts]
        if sum(g is not None for g, _ in got) < 0.8 * len(pts):
            continue
        best = None
        for k in scales:
            for sgn in (1, -1):
                err = sum(abs(sgn * k * (g or 0.0) - v) for g, v in got) / size
                if best is None or err < best[0]:
                    best = (err, sgn * k)
        words = _words(r.get("section", "") + " " + r["label"])
        sim = len(want & words) / (len(want | words) or 1)
        if not sim and _acronym(name, r["label"]):
            sim = 0.5
        ok = best[0] <= within and (sim > 0 or (best[0] <= min(tol, NO_NAME_TOL) and len(pts) >= NO_NAME_PTS))
        if ok:
            out.append({**{k: r[k] for k in ("sheet", "row", "label", "cols")}, "error": round(best[0], 4),
                        "scale": best[1], "label_match": round(sim, 2), "years": r["years"],
                        **({"within_rounding": True} if best[0] > tol else {}),
                        **({"rows": r["rows"]} if r.get("rows") else {})})
    out.sort(key=lambda m: (round(m["error"] / 0.01), -m["label_match"], m["error"]))
    return out[:limit]


def source(p: dict) -> str:
    """Where a pick's numbers are in the model: its row range, or its rows for a group."""
    from openpyxl.utils import get_column_letter as L
    a, b = L(p["cols"][0]), L(p["cols"][1])
    if p.get("rows") and len(p["rows"]) > 1:
        shown = ", ".join(str(r) for r in p["rows"][:6]) + ("…" if len(p["rows"]) > 6 else "")
        return f"{p['sheet']}!{a}:{b} rows {shown} (summed)"
    return f"{p['sheet']}!{a}{p['row']}:{b}{p['row']}"


# ---- recreating and checking ----------------------------------------------------------------------------------

KIND = {"column": "bar", "stacked column": "stacked", "line": "line", "area": "area", "combo": "combo",
        "waterfall": "bar"}


def spec_for(read: dict, picks: list[dict | None], years: list, title: str | None = None) -> dict | None:
    """A chart of the picked rows (one per series, None to leave a series out) over the report's years: each
    row's financial-year totals, the numbers the match compared, scaled and signed as the report shows them.
    Annual whatever each row's own timeline is, so rows from a quarterly and an annual sheet chart together."""
    ys = [y for y in years if y]
    labels = [f"FY{y}" for y in ys]
    series = []
    for s, p in zip(read["series"], picks):
        if not p:
            continue
        fy = {int(k): v for k, v in p["years"].items()}
        series.append({"name": s["name"] or p["label"], "range": source(p),
                       "data": [fy[y] * p["scale"] if isinstance(fy.get(y), (int, float)) else None for y in ys],
                       "as": "line" if s.get("drawn_as") == "line" else "bar"})
    if not series:
        return None
    kind = KIND.get((read.get("kind") or "").lower(), "bar")
    if kind == "bar" and any(x["as"] == "line" for x in series):
        kind = "combo"
    return {"title": title or read.get("title") or "Recreated chart", "kind": kind, "units": read.get("units") or "",
            "labels": labels, "period_labels": labels, "frequency": "annual", "sign": 1, "series": series}


def exact_values(spec: dict) -> dict:
    """What the recreation shows, by series and year label."""
    return {s["name"]: {lab: (round(v, 2) if isinstance(v, (int, float)) else None) for lab, v in zip(spec["labels"], s["data"])}
            for s in spec["series"]}


def comparison(read: dict, spec: dict | None) -> dict:
    """The report's reading beside the model's numbers, year by year, for every series (the page's compare
    view): {labels, series: [{name, drawn_as, read, model, diff (model - read, as a share of the reading)}]}."""
    years = [_year(x) for x in read.get("x_labels") or []]
    mine = {s["name"]: dict(zip((_year(x) for x in spec["labels"]), s["data"])) for s in (spec or {}).get("series") or []}
    out = []
    for s in read.get("series") or []:
        m = mine.get(s["name"])
        model = [m.get(y) if m and y else None for y in years]
        diff = [(b - a) / abs(a) if isinstance(a, (int, float)) and isinstance(b, (int, float)) and a else None
                for a, b in zip(s["values"], model)]
        out.append({"name": s["name"], "drawn_as": s.get("drawn_as"), "read": s["values"], "model": model if m else None,
                    "diff": diff if m else None})
    return {"labels": read.get("x_labels") or [], "series": out}


def render(spec: dict) -> bytes:
    import chartrender
    return chartrender.render_png(spec, None, mode="periodic")


def check(reader, report_png: bytes, ours_png: bytes, read: dict, ours: dict) -> dict:
    """The vision comparison: both images, with what was read from the report and our exact values."""
    import base64
    from llm import create
    prompt = CHECK_PROMPT.format(read=json.dumps({"x_labels": read["x_labels"], "series": read["series"]})[:6000],
                                 ours=json.dumps(ours)[:6000])
    content = [{"type": "input_text", "text": prompt}]
    for png in (report_png, ours_png):
        content.append({"type": "input_image", "image_url": f"data:image/png;base64,{base64.b64encode(png).decode()}",
                        "detail": "high"})
    r = create(reader.llm, reader.reviewer_model, text={"format": _CHECK}, max_output_tokens=3000,
               purpose="report-chart-check", input=[{"role": "user", "content": content}])
    if r.usage:
        reader.on_usage(reader.reviewer_model, r.usage, "report-chart-check")
    return json.loads(r.output_text)


def _numeric_check(read: dict, ours: dict, picks: list, skip: set = frozenset()) -> dict:
    """For a chart whose values are exact (PPTX), compare numbers instead of pictures. skip: series left out on
    purpose (too small to read)."""
    diffs, per = [], []
    for i, (s, p) in enumerate(zip(read["series"], picks)):
        if i in skip:
            continue
        mine = {_year(k): v for k, v in (ours.get(s["name"]) or {}).items()}  # by year: "FY26" is "FY2026"
        bad = [(lab, v, mine.get(_year(lab))) for lab, v in zip(read["x_labels"], s["values"])
               if isinstance(v, (int, float)) and (mine.get(_year(lab)) is None
                                                   or abs(mine[_year(lab)] - v) > 0.005 * max(1, abs(v)))]
        per.append({"name": s["name"], "matches": bool(p) and not bad, "note": "" if not bad else
                    f"{len(bad)} value(s) differ, e.g. {bad[0][0]}: report {bad[0][1]}, model {bad[0][2]}"})
        diffs += [f"{s['name']} {lab}: report {v}, model {m}" for lab, v, m in bad[:3]]
    return {"matches": all(x["matches"] for x in per), "series": per, "differences": diffs, "by": "numbers"}


TINY = 0.02  # a series under this share of the chart's largest can't be read off the picture: it isn't matched


def _slim(p: dict | None) -> dict | None:
    return {k: p[k] for k in ("book", "sheet", "row", "rows", "label", "error", "scale", "cols", "years", "manual")
            if k in p} if p else None


def _rows_verdict(read: dict, picks: list, skip: set = frozenset()) -> dict:
    """The verdict without a vision call, when most series have no row: which ones are missing."""
    missing = [s["name"] or f"series {i + 1}" for i, (s, p) in enumerate(zip(read["series"], picks)) if not p and i not in skip]
    return {"matches": False, "by": "rows",
            "differences": [f"{len(missing)} of {len(read['series'])} series have no matching row: {', '.join(missing[:8])}"],
            "series": [{"name": s["name"], "matches": False, "note": "no matching row" if not p else "not checked by eye"}
                       for s, p in zip(read["series"], picks)]}


def _stack_fill(read: dict, years: list, books: list[dict], title: str, cands: list, notes: dict, tol: float) -> list:
    """A stacked chart's segments are read off coarsely (each segment is a difference of two heights), but their
    sum, the top of the stack, is read well. A segment no row follows within tol takes its best row within twice
    tol, kept only if with every segment's row the stack's total follows the read total within tol; marked as
    found by the total. Otherwise the candidates are as they were."""
    loose = {}
    for i, cs in enumerate(cands):
        if cs or i in notes:
            continue
        s, found = read["series"][i], []
        for b in books:
            rows = b["rows"] + groups(b["rows"], s["name"], title)
            found += [{**m, "book": b["key"], "by": "the stack's total"}
                      for m in match(s["values"], years, rows, s["name"], units=read.get("units") or "", tol=2 * tol)]
        if not found:
            return cands  # a segment with nothing near it: the total can't vouch for it
        loose[i] = sorted(found, key=lambda m: (round(m["error"] / 0.01), -m["label_match"], m["error"]))[0]
    if not loose:
        return cands
    picks = {i: (cs[0] if cs else loose[i]) for i, cs in enumerate(cands) if i not in notes}
    if len({p.get("book") for p in picks.values()}) > 1:  # one workbook per chart
        return cands
    err = size = 0.0
    for j, y in enumerate(years):
        vals = [s["values"][j] for i, s in enumerate(read["series"]) if i in picks and j < len(s["values"])]
        if y is None or not any(isinstance(v, (int, float)) for v in vals):
            continue
        read_total = sum(v for v in vals if isinstance(v, (int, float)))
        ours = sum(p["scale"] * (p["years"].get(y) or 0.0) for p in picks.values())
        err, size = err + abs(ours - read_total), size + abs(read_total)
    if not size or err / size > tol:
        return cands
    return [[loose[i]] if i in loose else cs for i, cs in enumerate(cands)]


def recreate(reader, chart: dict, read: dict, books: list[dict], out_dir: Path, checks: dict | None = None) -> dict:
    """Match, draw and check one chart, trying the next candidates for a series the check rejects. checks: the
    vision comparisons kept (by both pictures, the values and the reviewer model): the same recreation isn't
    compared again."""
    years = time_axis(read.get("x_labels"))
    res = {"read": read, "tries": []}
    if not years:
        res["skipped"] = "not a chart over years: its x axis is categories, not a run of years"
        return res
    res["years"] = [y for y in years if y]
    biggest = max((abs(v) for s in read["series"] for v in s["values"] if isinstance(v, (int, float))), default=0) or 1
    tol = 0.01 if read.get("exact") else MATCH_TOL
    title = chart.get("caption") or read.get("title") or ""
    cands, notes = [], {}
    for i, s in enumerate(read["series"]):
        vals = [v for v in s["values"] if isinstance(v, (int, float))]
        if vals and max(abs(v) for v in vals) < TINY * biggest:
            notes[i] = "too small to read off the chart"
            cands.append([])
            continue
        cs = []
        for b in books:
            rows = b["rows"] + groups(b["rows"], s["name"], title)
            cs += [{**m, "book": b["key"]} for m in match(s["values"], years, rows, s["name"], units=read.get("units") or "",
                                                           tol=tol)]
        cs.sort(key=lambda m: (round(m["error"] / 0.01), -m["label_match"], m["error"]))
        cands.append(cs)
    if "stack" in (read.get("kind") or "").lower():
        cands = _stack_fill(read, years, books, title, cands, notes, tol)
    res["candidates"] = [[{k: m[k] for k in ("book", "sheet", "row", "label", "error", "scale")} for m in cs[:3]] for cs in cands]
    res["series_notes"] = notes
    if not any(cands):
        res["problem"] = "no model row follows any of its series"
        res["compare"] = comparison(read, None)
        return res
    choice = [0] * len(cands)
    report_png = (out_dir / chart["png"]).read_bytes() if chart.get("png") else None
    for k in range(MAX_TRIES):
        picks = [cs[i] if i < len(cs) else None for cs, i in zip(cands, choice)]
        book = next((p["book"] for p in picks if p), None)
        same = [p if p and p["book"] == book else None for p in picks]  # one workbook per chart
        spec = spec_for(read, same, years, title)
        if not spec:
            break
        png = render(spec)
        rel = f"charts/{chart['id']}-ours-{k + 1}.png"
        (out_dir / rel).write_bytes(png)
        wanted = len(same) - len(notes)
        if read.get("exact"):
            verdict = _numeric_check(read, exact_values(spec), same, set(notes))
        elif sum(p is not None for p in same) * 2 < wanted:  # the check would only say series are missing
            verdict = _rows_verdict(read, same, set(notes))
        elif report_png:
            key = _check_key(report_png, png, read, exact_values(spec), reader)
            verdict = (checks or {}).get(key)
            if verdict is None:
                verdict = check(reader, report_png, png, read, exact_values(spec))
                if checks is not None:
                    checks[key] = {**verdict, "by": verdict.get("by") or reader.reviewer_model}
                    _save_kept(out_dir, "checks", checks)
        else:
            verdict = _numeric_check(read, exact_values(spec), same, set(notes))
        verdict.setdefault("by", reader.reviewer_model)
        res["tries"].append({"picks": [_slim(p) for p in same], "png": rel, "verdict": verdict})
        res.update(spec=spec, picks=[_slim(p) for p in same], book=book, matches=verdict["matches"], ours_png=rel)
        if verdict["matches"] or verdict.get("by") == "rows":
            break
        rejected = {(sv.get("name") or "").strip().lower() for sv in verdict.get("series") or [] if not sv.get("matches")}
        bad = [i for i, s in enumerate(read["series"]) if (s.get("name") or "").strip().lower() in rejected]
        moved = False
        for i in bad or range(len(cands)):
            if choice[i] + 1 < len(cands[i]):
                choice[i] += 1
                moved = True
        if not moved:
            break
    res["compare"] = comparison(read, res.get("spec"))
    return res


def row_years(db, sheet: str, rows: list[int]) -> dict[int, dict]:
    """Financial-year totals of some rows of one sheet (fy_totals for a few rows): {row: {FY: total}}."""
    import chartdata
    lay = json.loads((db.execute("SELECT layout FROM sheets WHERE sheet=?", (sheet,)).fetchone() or ["{}"])[0] or "{}")
    if not lay.get("tl_first"):
        return {}
    cols = list(range(lay["tl_first"], lay["tl_last"] + 1))
    ends, _ = dcf.period_ends(db, sheet, cols)
    fy_m = chartdata.fy_end_month(db)
    fy = {c: (e.year if e.month <= fy_m else e.year + 1) for c, e in ends.items()}
    out: dict[int, dict] = {r: {} for r in rows}
    q = f"SELECT row, col, value FROM cells WHERE sheet=? AND row IN ({','.join('?' * len(rows))}) AND col BETWEEN ? AND ?"
    for r, c, v in db.execute(q, (sheet, *rows, cols[0], cols[-1])):
        if c in fy and isinstance(v, (int, float)) and not isinstance(v, bool):
            out[r][fy[c]] = out[r].get(fy[c], 0.0) + v
    return out


def pick_rows(db_path: str, read: dict, i: int, sheet: str, rows: list[int]) -> dict:
    """A person's rows for series i (one sheet): their financial-year totals summed, at the unit scale and sign
    that follow the reading best, with how far off it they are."""
    db = rodb.connect(db_path)
    per = row_years(db, sheet, rows)
    years: dict = {}
    for r in rows:
        for y, v in per.get(r, {}).items():
            years[y] = years.get(y, 0.0) + v
    if not years:
        raise ValueError(f"{sheet} has no timeline to total those rows by year")
    s = read["series"][i]
    ys = time_axis(read.get("x_labels")) or []
    pts = [(y, v) for y, v in zip(ys, s["values"]) if y and isinstance(v, (int, float))]
    size = sum(abs(v) for _, v in pts) or 1.0
    # the unit scale that brings the rows to the reading's size (a best fit would shrink wrong rows to nothing),
    # then the sign that fits
    import math
    mid = lambda xs: sorted(xs)[len(xs) // 2] if xs else 0.0
    want, have = mid([abs(v) for _, v in pts]), mid([abs(years.get(y, 0.0)) for y, _ in pts])
    k = min(_scales(read.get("units") or "", s["name"]),
            key=lambda k: abs(math.log((k * have or 1e-12) / (want or 1e-12))) if have and want else (k != 1))
    best = min((sum(abs(sgn * k * years.get(y, 0.0) - v) for y, v in pts) / size, sgn * k) for sgn in (1, -1))
    labels = dict(db.execute(f"SELECT row, label FROM rows WHERE sheet=? AND row IN ({','.join('?' * len(rows))})",
                             (sheet, *rows)))
    lay = json.loads((db.execute("SELECT layout FROM sheets WHERE sheet=?", (sheet,)).fetchone() or ["{}"])[0] or "{}")
    return {"sheet": sheet, "row": rows[0], **({"rows": list(rows)} if len(rows) > 1 else {}),
            "label": labels.get(rows[0]) or "" if len(rows) == 1 else f"{len(rows)} rows: " + "; ".join(labels.get(r) or "?" for r in rows[:4]),
            "cols": (lay.get("tl_first"), lay.get("tl_last")), "years": years, "scale": best[1],
            "error": round(best[0], 4) if pts else None, "manual": True}


def numbers_verdict(read: dict, picks: list, notes: dict | None = None) -> dict:
    """The verdict from the numbers alone (a person's picks): each series' row within MATCH_TOL of the reading."""
    notes = {int(k): v for k, v in (notes or {}).items()}
    per = []
    for i, (s, p) in enumerate(zip(read["series"], picks)):
        if i in notes:
            continue
        ok = bool(p) and p.get("error") is not None and p["error"] <= MATCH_TOL
        per.append({"name": s["name"], "matches": ok,
                    "note": "no row" if not p else f"{p['error']:.1%} off the reading" if p.get("error") is not None else ""})
    return {"matches": bool(per) and all(x["matches"] for x in per), "series": per, "by": "numbers",
            "differences": [f"{x['name']}: {x['note']}" for x in per if not x["matches"]]}


def current_spec(prior_db: str, current_db: str, read: dict, picks: list, title: str) -> dict | None:
    """The same chart on this year's client model: each row (each row of a group) found by its sheet and label,
    the nearest such row, over as many years from this year's first forecast year."""
    pdb, db = rodb.connect(prior_db), rodb.connect(current_db)
    mapped = []
    for p in picks:
        if not p:
            mapped.append(None)
            continue
        members = p.get("rows") or [p["row"]]
        labels = dict(pdb.execute(f"SELECT row, label FROM rows WHERE sheet=? AND row IN ({','.join('?' * len(members))})",
                                  (p["sheet"], *members)))
        found = []
        for r in members:
            hits = [x for (x,) in db.execute("SELECT row FROM rows WHERE sheet=? AND label=?", (p["sheet"], labels.get(r)))]
            if hits:
                found.append(min(hits, key=lambda x: abs(x - r)))
        if len(found) < len(members):
            mapped.append(None)
            continue
        per = row_years(db, p["sheet"], found)
        years: dict = {}
        for r in found:
            for y, v in per.get(r, {}).items():
                years[y] = years.get(y, 0.0) + v
        lay = json.loads((db.execute("SELECT layout FROM sheets WHERE sheet=?", (p["sheet"],)).fetchone() or ["{}"])[0] or "{}")
        mapped.append({**p, "row": found[0], "rows": found if p.get("rows") else None, "years": years,
                       "cols": (lay.get("tl_first") or p["cols"][0], lay.get("tl_last") or p["cols"][1])})
    if not any(mapped):
        return None
    years = time_axis(read.get("x_labels")) or []
    first = min(y for y in years if y)
    have = sorted({y for m in mapped if m for y in m["years"]})
    start = min((y for y in have if y > first), default=first + 1)
    shifted = [y - first + start if y else None for y in years]  # as many years, from this year's first forecast year
    return spec_for(read, mapped, shifted, title)


READ_VERSION = hashlib.sha256((DIGITISE_PROMPT + json.dumps(_DIGITISE, sort_keys=True)).encode()).hexdigest()[:12]


def _read_key(png: bytes, reader) -> str:
    """A reading holds for the same picture, the same prompt and schema, and the same model."""
    return f"{hashlib.sha256(png).hexdigest()[:24]}|{READ_VERSION}|{reader.model}"


def _reads(out_dir: Path, reader) -> dict:
    """The charts' readings kept (charts/reads.json), so recreating the charts again (a matcher changed, rows
    picked) doesn't read every picture again. Seeded once from the last run's record, whose readings were made with
    this prompt."""
    f = out_dir / "charts" / "reads.json"
    try:
        reads = json.loads(f.read_text(encoding="utf-8")) if f.exists() else {}
    except (OSError, ValueError):
        reads = {}
    last = out_dir / "report_charts.json"
    if not reads and last.exists():
        try:
            for ch in json.loads(last.read_text(encoding="utf-8")).get("charts") or []:
                if ch.get("read") and ch.get("png") and (out_dir / ch["png"]).exists() and ch.get("source") != "pptx chart":
                    reads[_read_key((out_dir / ch["png"]).read_bytes(), reader)] = ch["read"]
        except (OSError, ValueError):
            pass
    return reads


def _save_reads(out_dir: Path, reads: dict) -> None:
    _save_kept(out_dir, "reads", reads)


CHECK_VERSION = hashlib.sha256((CHECK_PROMPT + json.dumps(_CHECK, sort_keys=True)).encode()).hexdigest()[:12]


def _check_key(report_png: bytes, ours_png: bytes, read: dict, ours: dict, reader) -> str:
    """A comparison holds for the same two pictures, the same values beside them, prompt and reviewer model. Our
    picture is drawn from the spec alone, so the same recreation gives the same bytes."""
    h = hashlib.sha256()
    for part in (report_png, ours_png, json.dumps([read.get("x_labels"), read.get("series"), ours], sort_keys=True,
                                                   default=str).encode()):
        h.update(hashlib.sha256(part).digest())
    return f"{h.hexdigest()[:32]}|{CHECK_VERSION}|{reader.reviewer_model}"


def _kept(out_dir: Path, name: str) -> dict:
    f = out_dir / "charts" / f"{name}.json"
    try:
        return json.loads(f.read_text(encoding="utf-8")) if f.exists() else {}
    except (OSError, ValueError):
        return {}


def _save_kept(out_dir: Path, name: str, kept: dict) -> None:
    f = out_dir / "charts" / f"{name}.json"
    f.parent.mkdir(parents=True, exist_ok=True)
    f.write_text(json.dumps(kept, default=str), encoding="utf-8")


def run(reader, doc: dict, pdf_path: str | None, out_dir: str | Path, books: list[dict], current_db: str | None,
        progress=None) -> dict:
    """Every chart in the report: found, read, recreated from the prior model, checked, and redrawn on this
    year's model. books: [{"key", "name", "db_path"}] to search (the prior client model, the overlay)."""
    progress = progress or (lambda f, m: None)
    out_dir = Path(out_dir)
    charts = find(doc, pdf_path, out_dir)
    for b in books:
        b["rows"] = fy_totals(rodb.connect(b["db_path"]))
    reads, checks = _reads(out_dir, reader), _kept(out_dir, "checks")
    out = []
    for i, ch in enumerate(charts, 1):
        if ch.get("skipped"):
            out.append(ch)
            continue
        progress((i - 1) / max(1, len(charts)), f"Chart {i} of {len(charts)}: reading it")
        try:
            if ch["source"] == "pptx chart":
                read = from_markdown(ch["markdown"])
            else:
                png = (out_dir / ch["png"]).read_bytes()
                key = _read_key(png, reader)
                read = reads.get(key)
                if read is None:
                    read = reads[key] = digitise(reader, png, f"page {ch['page']}")
                    _save_reads(out_dir, reads)
        except Exception as e:  # one chart failing doesn't stop the rest
            out.append({**ch, "problem": f"couldn't read it: {type(e).__name__}: {e}"})
            continue
        if not read.get("is_chart") or not read.get("series"):
            out.append({**ch, "read": read, "skipped": "not a chart"})
            continue
        progress((i - 0.5) / max(1, len(charts)), f"Chart {i} of {len(charts)}: finding its rows and checking the recreation")
        try:
            res = recreate(reader, ch, read, books, out_dir, checks)
        except Exception as e:
            out.append({**ch, "read": read, "problem": f"couldn't recreate it: {type(e).__name__}: {e}"})
            continue
        item = {**ch, **{k: v for k, v in res.items()}}
        if res.get("picks") and current_db and res.get("book") == "prior_model":
            try:
                item["current_spec"] = current_spec(next(b for b in books if b["key"] == "prior_model")["db_path"], current_db,
                                                    read, res["picks"], (ch.get("caption") or read.get("title") or "") + " (this year)")
            except Exception as e:
                item["current_problem"] = f"{type(e).__name__}: {e}"
        out.append(item)
    for b in books:
        b.pop("rows", None)
    progress(1.0, "Done")
    return {"charts": out, "at": date.today().isoformat()}
