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
import io
import json
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


def _union(boxes):
    return [min(b[0] for b in boxes), min(b[1] for b in boxes), max(b[2] for b in boxes), max(b[3] for b in boxes)]


def _drawn_charts(page, skip: list) -> list[list[float]]:
    """Regions of a PDF page drawn as a chart: four or more filled bars sharing a baseline (or top), or plotted
    lines with several points, not inside a table region or a picture. Returns bboxes grown to take in the axis
    labels and legend."""
    W, H = float(page.width), float(page.height)
    box = lambda o: [float(o["x0"]), float(o["top"]), float(o["x1"]), float(o["bottom"])]
    free = lambda b: not any(_overlaps(b, s, -1) for s in skip)
    bars = [box(r) for r in page.rects if r.get("fill") and (r["x1"] - r["x0"]) < 0.5 * W and (r["bottom"] - r["top"]) < 0.6 * H
            and (r["x1"] - r["x0"]) > 1 and free(box(r))]
    seeds = []
    for key in (3, 1):  # bars standing on a common baseline, or hanging from one (negative values)
        groups: dict[int, list] = {}
        for b in bars:
            groups.setdefault(round(b[key]), []).append(b)
        for g in groups.values():
            widths = sorted(round(b[2] - b[0], 1) for b in g)
            if len(g) >= 4 and widths[-1] <= 2 * widths[0] + 1:
                seeds.append(_union(g))
    for c in page.curves:
        b = box(c)
        if len(c.get("pts") or []) >= 5 and (b[2] - b[0]) > 60 and free(b):
            seeds.append(b)
    regions: list[list[float]] = []
    for s in seeds:  # merge what overlaps (bars and the line drawn over them are one chart)
        for r in regions:
            if _overlaps(r, s, 12):
                r[:] = _union([r, s])
                break
        else:
            regions.append(list(s))
    out = []
    for r in regions:
        if (r[2] - r[0]) < 80 or (r[3] - r[1]) < 40:
            continue
        # take in the axes and ticks, then the tick labels and legend around those: without the axis numbers
        # the values can't be read
        g = _union([r] + [box(l) for l in page.lines if _overlaps(box(l), r, 30)])
        words = [box(w) for w in page.extract_words()]
        for _ in range(2):
            g = _union([g] + [w for w in words if _overlaps(w, g, 30)])
        out.append([max(0, g[0] - 6), max(0, g[1] - 6), min(W, g[2] + 6), min(H, g[3] + 6)])
    return out


def _caption(page, bbox) -> str:
    """The text line just above a chart (a "Figure 1: ..." caption, usually)."""
    lines: dict[int, list] = {}
    for w in page.extract_words():
        if w["bottom"] <= bbox[1] + 2 and w["bottom"] > bbox[1] - 40 and w["x1"] > bbox[0] and w["x0"] < bbox[2]:
            lines.setdefault(round(w["top"]), []).append(w)
    if not lines:
        return ""
    top = max(lines)
    return " ".join(w["text"] for w in sorted(lines[top], key=lambda w: w["x0"]))


def find(doc: dict, pdf_path: str | None, out_dir: str | Path) -> list[dict]:
    """The report's charts: [{"id", "page", "source", "png" (relative to out_dir), "caption", "data" (PPTX)}]."""
    out_dir = Path(out_dir)
    (out_dir / "charts").mkdir(parents=True, exist_ok=True)
    found = []
    for t in doc.get("tables") or []:
        if t.get("status") == "figure" and t.get("png"):
            found.append({"id": t["id"], "page": t.get("page"), "source": "picture", "png": t["png"],
                          "caption": t.get("title") or t.get("description") or ""})
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
                        and t.get("source") != "page image"]
                skip += [[float(i["x0"]), float(i["top"]), float(i["x1"]), float(i["bottom"])] for i in page.images]
                for k, bb in enumerate(_drawn_charts(page, skip), 1):
                    cid = f"p{pno:03d}-c{k}"
                    rel = f"charts/{cid}.png"
                    img = page.crop(bb).to_image(resolution=DPI).original
                    buf = io.BytesIO()
                    img.convert("RGB").save(buf, format="PNG")
                    (out_dir / rel).write_bytes(buf.getvalue())
                    found.append({"id": cid, "page": pno, "source": "drawn in the PDF", "png": rel,
                                  "caption": _caption(page, bb), "bbox": [round(v, 1) for v in bb]})
    return found


# ---- reading a chart -------------------------------------------------------------------------------------------

def digitise(reader, png: bytes, where: str) -> dict:
    return reader._call(reader.model, DIGITISE_PROMPT.format(where=where), png, _DIGITISE, "report-chart-read")


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
    s = str(label or "").strip()
    m = re.search(r"(?:FY|CY|YE)?\s*'?(\d{4}|\d{2})(?!\d)", s, re.I)
    if not m:
        return None
    y = int(m[1])
    return y + 2000 if y < 100 else y


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
        labels = dict(db.execute("SELECT row, label FROM rows WHERE sheet=?", (sheet,)))
        rows: dict[int, dict] = {}
        for r, c, v in db.execute("SELECT row, col, value FROM cells WHERE sheet=? AND col BETWEEN ? AND ?",
                                  (sheet, cols[0], cols[-1])):
            if c in fy and isinstance(v, (int, float)) and not isinstance(v, bool) and r != lay["header_row"]:
                rows.setdefault(r, {}).setdefault(fy[c], 0.0)
                rows[r][fy[c]] += v
        for r, years in rows.items():
            if len(years) >= 3 and any(years.values()):
                out.append({"sheet": sheet, "row": r, "label": labels.get(r) or "", "cols": (cols[0], cols[-1]),
                            "years": years})
    return out


def _words(s: str) -> set[str]:
    return {w for w in re.findall(r"[a-z]+", (s or "").lower()) if len(w) > 2}


def match(values: list, years: list, rows: list[dict], name: str = "", limit: int = 5) -> list[dict]:
    """Rows that follow a series: mean absolute error over its years, as a share of its size, after the best of a
    few unit scales and either sign; under MATCH_TOL. Best fit first; the label decides near-ties."""
    pts = [(y, v) for y, v in zip(years, values) if y is not None and isinstance(v, (int, float))]
    if len(pts) < 3:
        return []
    size = sum(abs(v) for _, v in pts) or 1.0
    want = _words(name)
    out = []
    for r in rows:
        got = [(r["years"].get(y), v) for y, v in pts]
        if sum(g is not None for g, _ in got) < 0.8 * len(pts):
            continue
        best = None
        for k in SCALES:
            for sgn in (1, -1):
                err = sum(abs(sgn * k * (g or 0.0) - v) for g, v in got) / size
                if best is None or err < best[0]:
                    best = (err, sgn * k)
        if best[0] <= MATCH_TOL:
            sim = len(want & _words(r["label"])) / (len(want | _words(r["label"])) or 1)
            out.append({**{k: r[k] for k in ("sheet", "row", "label", "cols")}, "error": round(best[0], 4),
                        "scale": best[1], "label_match": round(sim, 2)})
    out.sort(key=lambda m: (round(m["error"] / 0.01), -m["label_match"], m["error"]))
    return out[:limit]


# ---- recreating and checking ----------------------------------------------------------------------------------

KIND = {"column": "bar", "stacked column": "stacked", "line": "line", "area": "area", "combo": "combo",
        "waterfall": "bar"}


def spec_for(db_path: str, read: dict, picks: list[dict | None], years: list, title: str | None = None) -> dict | None:
    """A chart spec drawing the picked rows (one per series, None to leave a series out) over the report's years,
    annual, with the report's series names; values scaled and signed as the report shows them."""
    import tools
    from openpyxl.utils import get_column_letter as L
    series = []
    for s, p in zip(read["series"], picks):
        if p:
            series.append({"range": f"{p['sheet']}!{L(p['cols'][0])}{p['row']}:{L(p['cols'][1])}{p['row']}",
                           "name": s["name"] or p["label"], "as": "line" if s.get("drawn_as") == "line" else "bar"})
    if not series:
        return None
    kind = KIND.get((read.get("kind") or "").lower(), "bar")
    if kind == "bar" and any(x["as"] == "line" for x in series):
        kind = "combo"
    with tools.using(db_path):
        spec = tools.chart(title or read.get("title") or "Recreated chart", series, kind=kind,
                           units=read.get("units") or None)
    import chartdata
    for x, p in zip(spec["series"], [p for p in picks if p]):  # as the report shows them (units, sign)
        x["data"] = [v * p["scale"] if isinstance(v, (int, float)) else v for v in x["data"]]
    spec.pop("annual", None)
    spec = chartdata.enrich(spec, rodb.connect(db_path))  # the annual view again, from the scaled values
    spec["sign"] = 1
    if spec.get("annual"):
        spec["mode"] = "annual"  # the report charts financial years
    labs = spec["annual"]["labels"] if spec.get("annual") else spec.get("period_labels") or []
    idx = [i for i, lab in enumerate(labs) if _year(str(lab).rstrip("*")) in {y for y in years if y}]
    if idx and spec.get("annual"):  # the page keeps the window in periods and shows whole years of it
        py = spec["annual"]["period_year"]
        spec["view"] = {"x_start": min(i for i, y in enumerate(py) if y == idx[0]),
                        "x_end": max(i for i, y in enumerate(py) if y == idx[-1])}
    elif idx:
        spec["view"] = {"x_start": idx[0], "x_end": idx[-1]}
    return spec


def _shown_view(spec: dict) -> dict:
    """The window in the positions of what is shown: years in the annual view."""
    v = spec.get("view") or {}
    if v and spec.get("mode") == "annual" and spec.get("annual"):
        py = spec["annual"]["period_year"]
        return {"x_start": py[v["x_start"]], "x_end": py[v["x_end"]]}
    return v


def exact_values(spec: dict) -> dict:
    """What the recreation shows, by series and year label, as the page would show it (annual if it is)."""
    import chartdata
    shown = chartdata.display(spec, spec.get("mode") or "periodic")
    v = _shown_view(spec)
    lo, hi = v.get("x_start") or 0, v.get("x_end") if v.get("x_end") is not None else len(shown["labels"]) - 1
    return {s["name"]: {str(shown["labels"][i]).rstrip("*"): (round(s["data"][i], 2) if isinstance(s["data"][i], (int, float)) else None)
                        for i in range(lo, hi + 1)} for s in shown["series"]}


def render(spec: dict) -> bytes:
    import chartrender
    return chartrender.render_png(spec, _shown_view(spec), mode=spec.get("mode") or "periodic")


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


def _numeric_check(read: dict, ours: dict, picks: list) -> dict:
    """For a chart whose values are exact (PPTX), compare numbers instead of pictures."""
    diffs, per = [], []
    for s, p in zip(read["series"], picks):
        mine = {_year(k): v for k, v in (ours.get(s["name"]) or {}).items()}  # by year: "FY26" is "FY2026"
        bad = [(lab, v, mine.get(_year(lab))) for lab, v in zip(read["x_labels"], s["values"])
               if isinstance(v, (int, float)) and (mine.get(_year(lab)) is None
                                                   or abs(mine[_year(lab)] - v) > 0.005 * max(1, abs(v)))]
        per.append({"name": s["name"], "matches": bool(p) and not bad, "note": "" if not bad else
                    f"{len(bad)} value(s) differ, e.g. {bad[0][0]}: report {bad[0][1]}, model {bad[0][2]}"})
        diffs += [f"{s['name']} {lab}: report {v}, model {m}" for lab, v, m in bad[:3]]
    return {"matches": all(x["matches"] for x in per), "series": per, "differences": diffs, "by": "numbers"}


def recreate(reader, chart: dict, read: dict, books: list[dict], out_dir: Path) -> dict:
    """Match, draw and check one chart, trying the next candidates for a series the check rejects."""
    years = [_year(x) for x in read.get("x_labels") or []]
    res = {"read": read, "years": [y for y in years if y], "tries": []}
    if sum(y is not None for y in years) < 0.5 * max(1, len(years)):
        res["problem"] = "couldn't read the years off its x axis"
        return res
    cands = []
    for s in read["series"]:
        cs = []
        for b in books:
            cs += [{**m, "book": b["key"]} for m in match(s["values"], years, b["rows"], s["name"])]
        cs.sort(key=lambda m: (round(m["error"] / 0.01), -m["label_match"], m["error"]))
        cands.append(cs)
    res["candidates"] = [[{k: m[k] for k in ("book", "sheet", "row", "label", "error", "scale")} for m in cs[:3]] for cs in cands]
    if not any(cands):
        res["problem"] = "no model row follows any of its series"
        return res
    choice = [0] * len(cands)
    report_png = (out_dir / chart["png"]).read_bytes() if chart.get("png") else None
    by_key = {b["key"]: b for b in books}
    for k in range(MAX_TRIES):
        picks = [cs[i] if i < len(cs) else None for cs, i in zip(cands, choice)]
        book = next((p["book"] for p in picks if p), None)
        same = [p if p and p["book"] == book else None for p in picks]  # one workbook per chart
        spec = spec_for(by_key[book]["db_path"], read, same, years, chart.get("caption") or read.get("title"))
        if not spec:
            break
        ours = exact_values(spec)
        png = render(spec)
        rel = f"charts/{chart['id']}-ours-{k + 1}.png"
        (out_dir / rel).write_bytes(png)
        verdict = check(reader, report_png, png, read, ours) if report_png else _numeric_check(read, ours, same)
        verdict.setdefault("by", reader.reviewer_model)
        res["tries"].append({"picks": [{k2: p[k2] for k2 in ("book", "sheet", "row", "label", "error", "scale")} if p else None
                                       for p in same], "png": rel, "verdict": verdict})
        res.update(spec=spec, picks=same, book=book, matches=verdict["matches"], ours_png=rel)
        if verdict["matches"]:
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
    return res


def current_spec(prior_book: dict, current_db: str, read: dict, picks: list, title: str) -> dict | None:
    """The same chart on this year's client model: each row found by its sheet and label (the nearest such row),
    over this year's forecast years."""
    import rodb as _r
    db = _r.connect(current_db)
    mapped = []
    for p in picks:
        if not p:
            mapped.append(None)
            continue
        hits = [r for (r,) in db.execute("SELECT row FROM rows WHERE sheet=? AND label=?", (p["sheet"], p["label"]))]
        if not hits:
            mapped.append(None)
            continue
        lay = json.loads((db.execute("SELECT layout FROM sheets WHERE sheet=?", (p["sheet"],)).fetchone() or ["{}"])[0] or "{}")
        if not lay.get("tl_first"):
            mapped.append(None)
            continue
        mapped.append({**p, "row": min(hits, key=lambda r: abs(r - p["row"])), "cols": (lay["tl_first"], lay["tl_last"])})
    if not any(mapped):
        return None
    years = [_year(x) for x in read.get("x_labels") or []]
    first = min(y for y in years if y)
    rows = fy_totals(db)
    have = sorted({y for r in rows if r["sheet"] == next(m["sheet"] for m in mapped if m) for y in r["years"]})
    start = min((y for y in have if y > first), default=first + 1)
    shifted = [y - first + start if y else None for y in years]  # as many years, from this year's first forecast year
    return spec_for(current_db, read, mapped, shifted, title)


def run(reader, doc: dict, pdf_path: str | None, out_dir: str | Path, books: list[dict], current_db: str | None,
        progress=None) -> dict:
    """Every chart in the report: found, read, recreated from the prior model, checked, and redrawn on this
    year's model. books: [{"key", "name", "db_path"}] to search (the prior client model, the overlay)."""
    progress = progress or (lambda f, m: None)
    out_dir = Path(out_dir)
    charts = find(doc, pdf_path, out_dir)
    for b in books:
        b["rows"] = fy_totals(rodb.connect(b["db_path"]))
    out = []
    for i, ch in enumerate(charts, 1):
        progress((i - 1) / max(1, len(charts)), f"Chart {i} of {len(charts)}: reading it")
        try:
            read = from_markdown(ch["markdown"]) if ch["source"] == "pptx chart" else \
                digitise(reader, (out_dir / ch["png"]).read_bytes(), f"page {ch['page']}")
        except Exception as e:  # one chart failing doesn't stop the rest
            out.append({**ch, "problem": f"couldn't read it: {type(e).__name__}: {e}"})
            continue
        if not read.get("is_chart") or not read.get("series"):
            continue  # a picture that isn't a chart
        progress((i - 0.5) / max(1, len(charts)), f"Chart {i} of {len(charts)}: finding its rows and checking the recreation")
        try:
            res = recreate(reader, ch, read, books, out_dir)
        except Exception as e:
            out.append({**ch, "read": read, "problem": f"couldn't recreate it: {type(e).__name__}: {e}"})
            continue
        item = {**ch, **{k: v for k, v in res.items()}}
        if res.get("picks") and current_db and res.get("book") == "prior_model":
            try:
                item["current_spec"] = current_spec(next(b for b in books if b["key"] == "prior_model"), current_db, read,
                                                    res["picks"], (ch.get("caption") or read.get("title") or "") + " (this year)")
            except Exception as e:
                item["current_problem"] = f"{type(e).__name__}: {e}"
        out.append(item)
    for b in books:
        b.pop("rows", None)
    progress(1.0, "Done")
    return {"charts": out, "at": date.today().isoformat()}
