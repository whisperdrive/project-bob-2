"""Report documents (PDF, PPTX) -> Markdown per page, with every table read from a picture of it and checked.

PDF: the text layer gives the prose. Each table region (ruled tables, runs of number-heavy lines, and
embedded images, which are often tables pasted as pictures) is cropped, rendered at 200 dpi and read by a
vision model into a Markdown table. Then it is checked:
  - text-layer tables: every row's numbers must appear on one line of the page's own text in that region,
    and every number in the region must be in the transcription (catches dropped or misread digits)
  - picture-only tables: a second, independent read by the reviewer model (it doesn't see the first read);
    the two reads are compared row by row, number by number
A table is "verified" when its check passes, otherwise "flagged" with the differences, for a person to settle.
Pages with no text layer (scans) are read whole from an image and checked the same way as pictures.

PPTX: text boxes, native tables and chart data come straight from the file (exact, no reading involved);
pictures go through the same read-and-check as PDF pictures.

process() returns {"kind", "pages": [{"n", "title", "md"}], "tables": [...], "markdown"} and writes table
images to out_dir/tables/. The Markdown has <!-- page N --> markers so extracted facts can cite pages.
    uv run python bench/docingest.py report.pdf out/docs/x [model]
"""
import base64
import io
import json
import re
import sys
from collections import Counter
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

DPI = 200
MIN_IMAGE_PT = (150, 60)  # smaller embedded images are logos / icons
MIN_PICTURE_PX = (300, 90)
READERS = 4  # table reads in flight at once (each call still waits for rate-limit room)

READ_PROMPT = """This image is cut from {where} of a valuation report. If it shows a table, transcribe it
exactly as one GitHub-flavoured Markdown table: every row and column, numbers exactly as printed (keep thousands
separators, decimals, %, currency symbols and brackets for negatives), empty cells left empty, a header row
first (repeat a merged header in each column it spans). Put the table's caption or title, if one is visible,
in title. If it isn't a table (a chart, photo, logo or diagram), set is_table false, markdown "" and describe it
in one sentence."""
PAGE_PROMPT = """This is {where} of a valuation report, as an image. Transcribe it to Markdown: headings as ###,
paragraphs as text, bullet lists as lists, and every table as a GitHub-flavoured Markdown table with numbers
exactly as printed. Don't summarise or add commentary. Put "" in title if there's no heading."""
_S = {"type": "string"}
_READ = {"type": "json_schema", "name": "table_read", "strict": True, "schema": {
    "type": "object", "additionalProperties": False, "required": ["is_table", "title", "markdown", "description"],
    "properties": {"is_table": {"type": "boolean"}, "title": _S, "markdown": _S, "description": _S}}}
_PAGE = {"type": "json_schema", "name": "page_read", "strict": True, "schema": {
    "type": "object", "additionalProperties": False, "required": ["title", "markdown"],
    "properties": {"title": _S, "markdown": _S}}}


# ---- numbers ------------------------------------------------------------------------------------------------

NUM = re.compile(r"\(?[-−–]?\s?(?:[A-Z]{0,2}\$|€|£)?\s?\d[\d,]*(?:\.\d+)?\s?%?\)?")


def numbers(text: str) -> list[str]:
    """Numbers as printed, normalised: '(1,234.5)' -> '-1234.5', '7.25%' -> '7.25%', 'A$850.0m' -> '850.0'."""
    out = []
    for m in NUM.finditer(text or ""):
        s = m.group(0).strip()
        digits = re.sub(r"[^\d.]", "", s).rstrip(".")
        if not digits or not re.search(r"\d", digits):
            continue
        neg = (s.startswith("(") and s.endswith(")")) or bool(re.match(r"\(?[-−–]", s))
        out.append(("-" if neg else "") + digits + ("%" if "%" in s else ""))
    return out


def md_rows(md: str) -> list[list[str]]:
    """Cells of each body row of the Markdown table(s) in md (separator rows skipped)."""
    rows = []
    for line in (md or "").splitlines():
        line = line.strip()
        if not line.startswith("|"):
            continue
        cells = [c.strip() for c in line.strip("|").split("|")]
        if all(re.fullmatch(r":?-{2,}:?", c) for c in cells if c):
            continue
        rows.append(cells)
    return rows


def _row_key(cells: list[str]) -> str:
    return re.sub(r"[^a-z0-9]+", " ", (cells[0] if cells else "").lower()).strip()


def compare_reads(a: str, b: str) -> dict:
    """Two independent transcriptions of one table, compared row by row (rows matched by first cell)."""
    ra, rb = md_rows(a), md_rows(b)
    diffs = []
    kb = {}
    for cells in rb:
        kb.setdefault(_row_key(cells), []).append(cells)
    for cells in ra:
        others = kb.get(_row_key(cells))
        other = others.pop(0) if others else None
        na = numbers(" ".join(cells[1:]))
        nb = numbers(" ".join(other[1:])) if other else None
        if nb is None:
            if na:
                diffs.append({"row": cells[0], "first": " | ".join(cells[1:]), "second": "(row not found)"})
        elif Counter(na) != Counter(nb):
            diffs.append({"row": cells[0], "first": " | ".join(cells[1:]), "second": " | ".join(other[1:])})
    for rest in kb.values():
        for cells in rest:
            if numbers(" ".join(cells[1:])):
                diffs.append({"row": cells[0], "first": "(row not found)", "second": " | ".join(cells[1:])})
    n = sum(len(numbers(" ".join(c[1:]))) for c in ra)
    return {"method": "second read", "numbers": n, "rows": len(ra), "differences": diffs, "ok": not diffs and n > 0}


def _in_order(seq: list, line: list) -> bool:
    it = iter(line)
    return all(x in it for x in seq)


def _words(text: str) -> Counter:
    return Counter(w for w in re.findall(r"[a-z]+", (text or "").lower()))


def check_text_layer(md: str, lines: list[str]) -> dict:
    """Transcription vs the PDF's own text in the table region. lines = the region's text, one entry per line.
    Numbers must match exactly; words too, so a dropped unit ("$m" for "A$m") or a misread label is caught."""
    src_all = Counter(n for ln in lines for n in numbers(ln))
    src_lines = [numbers(ln) for ln in lines]
    got_all = Counter()
    bad_rows = []
    for cells in md_rows(md):
        row = numbers(" ".join(cells[1:]))
        got_all += Counter(numbers(" ".join(cells)))
        # the row's numbers, in order, on one source line (so a swapped column or a row's figures under
        # another row's label is caught, not just a misread digit)
        if row and not any(_in_order(row, line) for line in src_lines):
            bad_rows.append({"row": cells[0], "numbers": row})
    not_in_source = list((got_all - src_all).elements())
    not_transcribed = list((src_all - got_all).elements())
    src_words, got_words = _words(" ".join(lines)), _words(md)
    words_not_in_source = sorted(set(got_words - src_words))
    words_not_transcribed = sorted(set(src_words - got_words))
    return {"method": "text layer", "numbers": sum(got_all.values()), "rows_not_on_one_line": bad_rows,
            "not_in_source": not_in_source, "not_transcribed": not_transcribed,
            "words_not_in_source": words_not_in_source, "words_not_transcribed": words_not_transcribed,
            "ok": not (bad_rows or not_in_source or not_transcribed or words_not_in_source or words_not_transcribed)
                  and sum(got_all.values()) > 0}


# ---- model calls --------------------------------------------------------------------------------------------

class Reader:
    """Vision calls through llm.create (rate-limited); usage goes to on_usage(model, usage, purpose)."""

    def __init__(self, model: str, reviewer_model: str, on_usage=None):
        from llm import client
        self.llm, self.model, self.reviewer_model = client(interactive=False), model, reviewer_model
        self.on_usage = on_usage or (lambda *a: None)

    def _call(self, model, prompt, png: bytes, schema, purpose) -> dict:
        from llm import create
        b64 = base64.b64encode(png).decode()
        r = create(self.llm, model, text={"format": schema}, max_output_tokens=6000, input=[{"role": "user", "content": [
            {"type": "input_text", "text": prompt},
            {"type": "input_image", "image_url": f"data:image/png;base64,{b64}", "detail": "high"}]}])
        if r.usage:
            self.on_usage(model, r.usage, purpose)
        return json.loads(r.output_text)

    def table(self, png: bytes, where: str, second: bool = False) -> dict:
        return self._call(self.reviewer_model if second else self.model, READ_PROMPT.format(where=where), png, _READ,
                          "doc-review" if second else "doc-table")

    def page(self, png: bytes, where: str, second: bool = False) -> dict:
        return self._call(self.reviewer_model if second else self.model, PAGE_PROMPT.format(where=where), png, _PAGE,
                          "doc-review" if second else "doc-page")


def _png(img) -> bytes:
    buf = io.BytesIO()
    img.convert("RGB").save(buf, format="PNG")
    return buf.getvalue()


def read_and_check(reader: Reader, t: dict, png: bytes, text_lines: list[str] | None) -> dict:
    """Fill in t (a table record) from the reads and checks."""
    whole_page = t["source"] == "page image"
    first = reader.page(png, t["where"]) if whole_page else reader.table(png, t["where"])
    if not whole_page and not first["is_table"]:
        return {**t, "is_table": False, "markdown": "", "description": first["description"], "status": "figure",
                "check": None}
    second = reader.page(png, t["where"], True) if whole_page else reader.table(png, t["where"], True)
    check = check_text_layer(first["markdown"], text_lines) if text_lines else None
    agree = compare_reads(first["markdown"], second["markdown"])
    ok = check["ok"] if check else agree["ok"]
    return {**t, "is_table": True, "title": first.get("title") or None, "markdown": first["markdown"],
            "second_markdown": second["markdown"], "check": check, "second_read": agree,
            "status": "verified" if ok else "flagged"}


# ---- PDF ----------------------------------------------------------------------------------------------------

def _lines(words: list[dict], tol: float = 2.5) -> list[dict]:
    """Group words into lines: [{"top", "bottom", "x0", "x1", "text", "size"}], top to bottom."""
    lines = []
    for w in sorted(words, key=lambda w: (round(w["top"]), w["x0"])):
        if lines and abs(lines[-1]["top"] - w["top"]) <= tol:
            ln = lines[-1]
            ln["words"].append(w)
            ln["bottom"] = max(ln["bottom"], w["bottom"])
        else:
            lines.append({"top": w["top"], "bottom": w["bottom"], "words": [w]})
    for ln in lines:
        ws = sorted(ln["words"], key=lambda w: w["x0"])
        ln.update(text=" ".join(w["text"] for w in ws), x0=ws[0]["x0"], x1=ws[-1]["x1"],
                  size=sum(w.get("size", 0) for w in ws) / len(ws))
        del ln["words"]
    return lines


def _inside(ln: dict, box) -> bool:
    x0, top, x1, bottom = box
    mid = (ln["top"] + ln["bottom"]) / 2
    return top - 1 <= mid <= bottom + 1 and ln["x1"] >= x0 and ln["x0"] <= x1


def _number_runs(lines: list[dict], taken: list) -> list[tuple]:
    """Runs of 3+ lines with 2+ numbers each (borderless tables), plus a header line just above."""
    boxes, run = [], []

    def flush():
        if len(run) >= 3:
            i0 = lines.index(run[0])
            head = [lines[i0 - 1]] if i0 > 0 and run[0]["top"] - lines[i0 - 1]["bottom"] < 14 else []
            rs = head + run
            boxes.append((min(r["x0"] for r in rs) - 4, rs[0]["top"] - 3, max(r["x1"] for r in rs) + 4, rs[-1]["bottom"] + 3))

    for ln in lines:
        words = ln["text"].split()
        n = sum(1 for w in words if numbers(w))
        # tabular: mostly numbers (a row label plus figures), not prose that happens to quote two amounts
        numeric = n >= 2 and n >= 0.5 * len(words) and not any(_inside(ln, b) for b in taken)
        if numeric and (not run or ln["top"] - run[-1]["bottom"] < 16):
            run.append(ln)
        else:
            flush()
            run = [ln] if numeric else []
    flush()
    return boxes


def _pdf(path: str, out_dir: Path, reader: Reader | None, progress) -> dict:
    import pdfplumber
    pages, tables, jobs = [], [], []
    with pdfplumber.open(path) as pdf:
        n_pages = len(pdf.pages)
        for pno, page in enumerate(pdf.pages, start=1):
            progress(0.05 + 0.25 * pno / n_pages, f"Reading page {pno} of {n_pages}")
            words = page.extract_words(extra_attrs=["size"])
            lines = _lines(words)
            area = float(page.width * page.height)
            imgs = [(float(i["x0"]), float(i["top"]), float(i["x1"]), float(i["bottom"])) for i in page.images
                    if i["x1"] - i["x0"] >= MIN_IMAGE_PT[0] and i["bottom"] - i["top"] >= MIN_IMAGE_PT[1]]
            blocks = []  # (top, kind, payload)
            if len(words) < 5 and any((b[2] - b[0]) * (b[3] - b[1]) > 0.4 * area for b in imgs):
                t = {"id": f"p{pno:03d}-page", "page": pno, "source": "page image", "where": f"page {pno}",
                     "bbox": [0, 0, float(page.width), float(page.height)]}
                jobs.append((t, _png(page.to_image(resolution=DPI).original), None))
                tables.append(t)
                blocks.append((0, "table", t["id"]))
            else:
                ruled = [tuple(float(v) for v in tb.bbox) for tb in page.find_tables()]
                ruled = [b for b in ruled if (b[2] - b[0]) > 60 and (b[3] - b[1]) > 20]
                regions = [(b, "text-layer table") for b in ruled]
                regions += [(b, "text-layer table") for b in _number_runs(lines, ruled + imgs)]
                regions += [(b, "picture") for b in imgs if not any(_overlap(b, r) for r, _ in regions)]
                for k, (box, source) in enumerate(sorted(regions, key=lambda r: r[0][1]), start=1):
                    box = _pad(box, page)
                    t = {"id": f"p{pno:03d}-t{k}", "page": pno, "source": source, "where": f"page {pno}",
                         "bbox": [round(v, 1) for v in box]}
                    inside = [ln["text"] for ln in lines if _inside(ln, box)]
                    if source == "text-layer table":
                        t["text_lines"] = inside  # kept so a person's edit can be checked against the page too
                    png = _png(page.crop(box).to_image(resolution=DPI).original)
                    jobs.append((t, png, inside if source == "text-layer table" and inside else None))
                    tables.append(t)
                    blocks.append((box[1], "table", t["id"]))
                sizes = sorted(ln["size"] for ln in lines) or [10]
                body = sizes[len(sizes) // 2]
                prev = None
                for ln in lines:
                    if any(_inside(ln, t["bbox"]) for t in tables if t["page"] == pno):
                        continue
                    if re.fullmatch(r"(page )?\d+( of \d+)?", ln["text"].strip(), re.I) and ln["top"] > page.height * 0.9:
                        continue  # page number in the footer
                    heading = ln["size"] >= body * 1.3
                    gap = prev is not None and ln["top"] - prev["bottom"] > 1.2 * (prev["bottom"] - prev["top"])
                    blocks.append((ln["top"], "heading" if heading else "text", ln["text"], gap))
                    prev = ln
            pages.append({"n": pno, "blocks": sorted(blocks, key=lambda b: b[0])})
    _read_all(reader, jobs, out_dir, progress)
    return {"kind": "pdf", "pages": pages, "tables": tables}


def _overlap(a, b) -> bool:
    return not (a[2] <= b[0] or b[2] <= a[0] or a[3] <= b[1] or b[3] <= a[1])


def _pad(box, page, pad: float = 6):
    return (max(0, box[0] - pad), max(0, box[1] - pad), min(float(page.width), box[2] + pad),
            min(float(page.height), box[3] + pad))


def _page_md(blocks, by_id) -> tuple[str, str | None]:
    out, title, para = [], None, []

    def end_para():
        if para:
            out.append(" ".join(para))
            para.clear()

    for b in blocks:
        if b[1] == "table":
            end_para()
            t = by_id[b[2]]
            if t.get("source") == "page image":
                title = title or t.get("title")
            out.append(table_md(t))
        elif b[1] == "heading":
            end_para()
            title = title or b[2]
            out.append(f"### {b[2]}")
        else:
            if b[3]:
                end_para()
            para.append(b[2])
    end_para()
    return "\n\n".join(x for x in out if x), title


def table_md(t: dict) -> str:
    """How a table appears in the document Markdown: the approved text if a person settled it."""
    if t.get("status") == "figure":
        return f"*[Figure: {t.get('description') or 'image'}]*"
    md = t.get("final_markdown") or t.get("markdown") or ""
    if not md:
        return f"*[Table {t['id']}: not read]*"
    if t.get("source") == "page image":
        return md
    cap = f"**{t['title']}**\n\n" if t.get("title") else ""
    return f"<!-- table {t['id']} ({t.get('status')}) -->\n{cap}{md}"


def _read_all(reader: Reader | None, jobs: list, out_dir: Path, progress) -> None:
    (out_dir / "tables").mkdir(parents=True, exist_ok=True)
    for t, png, _ in jobs:
        t["png"] = f"tables/{t['id']}.png"
        (out_dir / t["png"]).write_bytes(png)
    if not jobs:
        return
    if reader is None:
        for t, _, _ in jobs:
            t.update(status="unread", markdown="", check=None)
        return
    done = 0

    def one(job):
        t, png, lines = job
        try:
            t.update(read_and_check(reader, t, png, lines))
        except Exception as e:  # keep going; the table shows as failed with the reason
            t.update(status="error", markdown="", error=f"{type(e).__name__}: {e}")
        return t

    with ThreadPoolExecutor(READERS) as pool:
        for t in pool.map(one, jobs):
            done += 1
            progress(0.3 + 0.65 * done / len(jobs), f"Read and checked {done} of {len(jobs)} tables")


# ---- PPTX ---------------------------------------------------------------------------------------------------

def _shapes(shapes):
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    for s in shapes:
        if s.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _shapes(s.shapes)
        else:
            yield s


def _grid_md(rows: list[list[str]]) -> str:
    rows = [[(c or "").replace("|", "/").replace("\n", " ").strip() for c in r] for r in rows]
    if not rows:
        return ""
    w = max(len(r) for r in rows)
    rows = [r + [""] * (w - len(r)) for r in rows]
    return "\n".join(["| " + " | ".join(rows[0]) + " |", "|" + "---|" * w] + ["| " + " | ".join(r) + " |" for r in rows[1:]])


def _chart_md(chart) -> str:
    plot = chart.plots[0]
    cats = [str(c) for c in plot.categories]
    series = [(s.name, list(s.values)) for p in chart.plots for s in p.series]
    rows = [[""] + [name for name, _ in series]]
    for i, c in enumerate(cats):
        rows.append([c] + [("" if i >= len(v) or v[i] is None else f"{v[i]:,.6g}") for _, v in series])
    return _grid_md(rows)


def _pptx(path: str, out_dir: Path, reader: Reader | None, progress) -> dict:
    from PIL import Image
    from pptx import Presentation
    prs = Presentation(path)
    pages, tables, jobs = [], [], []
    n_slides = len(prs.slides)
    for sno, slide in enumerate(prs.slides, start=1):
        progress(0.05 + 0.25 * sno / n_slides, f"Reading slide {sno} of {n_slides}")
        title_shape = slide.shapes.title
        title = title_shape.text_frame.text.strip() if title_shape is not None and title_shape.has_text_frame else None
        blocks, k = [], 0
        for s in sorted(_shapes(slide.shapes), key=lambda s: ((s.top or 0), (s.left or 0))):
            top = s.top or 0
            if s is title_shape or (title_shape is not None and s.shape_id == title_shape.shape_id):
                continue
            if getattr(s, "has_table", False) and s.has_table:
                k += 1
                grid = [[c.text for c in r.cells] for r in s.table.rows]
                t = {"id": f"s{sno:03d}-t{k}", "page": sno, "source": "pptx table", "where": f"slide {sno}",
                     "is_table": True, "markdown": _grid_md(grid), "status": "verified",
                     "check": {"method": "native", "ok": True, "note": "read from the slide's table, not from an image"}}
                tables.append(t)
                blocks.append((top, "table", t["id"]))
            elif getattr(s, "has_chart", False) and s.has_chart:
                k += 1
                ch = s.chart
                name = ch.chart_title.text_frame.text if ch.has_title and ch.chart_title.has_text_frame else None
                t = {"id": f"s{sno:03d}-t{k}", "page": sno, "source": "pptx chart data", "where": f"slide {sno}",
                     "is_table": True, "title": name, "markdown": _chart_md(ch), "status": "verified",
                     "check": {"method": "native", "ok": True, "note": "the chart's own data, from the file"}}
                tables.append(t)
                blocks.append((top, "table", t["id"]))
            elif hasattr(s, "image"):
                try:
                    img = Image.open(io.BytesIO(s.image.blob))
                except Exception:
                    continue
                if img.width < MIN_PICTURE_PX[0] or img.height < MIN_PICTURE_PX[1]:
                    continue
                k += 1
                t = {"id": f"s{sno:03d}-t{k}", "page": sno, "source": "picture", "where": f"slide {sno}"}
                jobs.append((t, _png(img), None))
                tables.append(t)
                blocks.append((top, "table", t["id"]))
            elif s.has_text_frame and s.text_frame.text.strip():
                lines = []
                for p in s.text_frame.paragraphs:
                    txt = "".join(r.text for r in p.runs).strip()
                    if txt:
                        lines.append(("  " * max(0, p.level - 1) + "- " if p.level else "") + txt)
                blocks.append((top, "text", "\n".join(lines), True))
        if slide.has_notes_slide and slide.notes_slide.notes_text_frame.text.strip():
            blocks.append((10 ** 12, "text", "> Speaker notes: " + slide.notes_slide.notes_text_frame.text.strip(), True))
        pages.append({"n": sno, "title": title, "blocks": sorted(blocks, key=lambda b: b[0])})
    _read_all(reader, jobs, out_dir, progress)
    return {"kind": "pptx", "pages": pages, "tables": tables}


# ---- entry point --------------------------------------------------------------------------------------------

def render(doc: dict) -> str:
    """(Re)build each page's Markdown and the whole document's from the page blocks and the tables, so a
    table a person approves or edits shows up everywhere. Slides count as pages; <!-- page N --> marks each."""
    by_id = {t["id"]: t for t in doc["tables"]}
    label = "Slide" if doc["kind"] == "pptx" else "Page"
    for p in doc["pages"]:
        md, title = _page_md(p["blocks"], by_id)
        if doc["kind"] == "pptx":
            md = (f"### {p['title']}\n\n" if p.get("title") else "") + md
        else:
            p["title"] = title
        p["md"] = md
    doc["markdown"] = "\n\n".join(f"<!-- page {p['n']} -->\n## {label} {p['n']}\n\n{p['md']}" for p in doc["pages"])
    return doc["markdown"]


def process(path: str, out_dir: str | Path, model: str = "gpt-6-luna", reviewer_model: str = "gpt-6-sol",
            progress=None, on_usage=None, read: bool = True) -> dict:
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    progress = progress or (lambda frac, msg: None)
    reader = Reader(model, reviewer_model, on_usage) if read else None
    ext = Path(path).suffix.lower()
    if ext == ".pdf":
        doc = _pdf(path, out_dir, reader, progress)
    elif ext == ".pptx":
        doc = _pptx(path, out_dir, reader, progress)
    else:
        raise ValueError(f"{Path(path).name}: reports must be .pdf or .pptx (save .ppt / .docx as PDF first)")
    (out_dir / "document.md").write_text(render(doc), encoding="utf-8")
    progress(1.0, "Document read")
    return doc


if __name__ == "__main__":
    sys.path.insert(0, str(Path(__file__).parent))
    if len(sys.argv) < 3:
        sys.exit(__doc__)
    d = process(sys.argv[1], sys.argv[2], *(sys.argv[3:4] * 2), progress=lambda f, m: print(f"{f:4.0%} {m}"))
    for t in d["tables"]:
        print(t["id"], t["source"], t["status"], json.dumps(t.get("check") or t.get("second_read"), default=str)[:300])
    print(d["markdown"][:3000])
