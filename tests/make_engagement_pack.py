"""Write a synthetic recurring-valuation pack to tests/engagement_pack/ (git-ignored; re-run to regenerate).

A fictional toll road ("Riverbend Toll Road", code name Project Kestrel), valued a year ago and now being
rolled forward. Two layouts of the same prior valuation, so both ways an overlay is delivered are covered:

  Pack A (overlay standalone)                Pack B (overlay inside the client model)
    Riverbend_valuation_report_FY25.pdf        Riverbend_valuation_report_FY25.pptx
    Riverbend_BP25_client_model.xlsx           Riverbend_BP25_with_overlay.xlsx
    Kestrel_valuation_overlay_FY25.xlsx        (Val_Inputs / DCF / Summary sheets inside it)
    Riverbend_BP26_client_model.xlsx           Riverbend_BP26_client_model.xlsx

The standalone overlay reads the client model through real external links ([1]CashFlow!D10, with the
link's cached values), as Excel saves them. The PDF has text-layer tables and one table that is only an
image (a pasted picture), so both extraction paths run. Every number the report quotes is computed here
from the models, so report, overlay and client model agree.
    uv run python tests/make_engagement_pack.py
"""
import io
import re
import sys
import zipfile
from datetime import date
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import xlsxwriter  # noqa: E402
from matplotlib.backends.backend_pdf import PdfPages  # noqa: E402
from xlsxwriter.utility import xl_col_to_name as COL  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "bench"))
import dcf  # noqa: E402

OUT = ROOT / "tests" / "engagement_pack"
YEARS = 20
FIRST_COL = 3  # column D


# ---- the client model ------------------------------------------------------------------------------------

def client_numbers(fy0: int, traffic0: float, growth: float, toll0: float, cpi: float, opex_pct: float,
                   capex: float, major: float, insurance: float | None) -> dict:
    ends = [date(fy0 + i, 6, 30) for i in range(YEARS)]
    traffic = [traffic0 * (1 + growth) ** i for i in range(YEARS)]
    toll = [toll0 * (1 + cpi) ** i for i in range(YEARS)]
    revenue = [t * p for t, p in zip(traffic, toll)]
    opex = [-r * opex_pct for r in revenue]
    ins = [-(insurance or 0) * (1 + cpi) ** i for i in range(YEARS)]
    ebitda = [r + o + s for r, o, s in zip(revenue, opex, ins)]
    cap = [-(capex * (1 + cpi) ** i + (major if (fy0 + i) % 5 == 0 else 0)) for i in range(YEARS)]
    tax = [-max(0.0, (e + c) * 0.30) for e, c in zip(ebitda, cap)]
    fcf = [e + c + t for e, c, t in zip(ebitda, cap, tax)]
    return dict(ends=ends, traffic=traffic, toll=toll, revenue=revenue, opex=opex, insurance=ins, ebitda=ebitda,
                capex=cap, tax=tax, fcf=fcf)


def write_client(wb, n: dict, inputs: dict, insurance: bool) -> dict:
    """Inputs / Operations / CashFlow sheets. Returns {line item: (sheet, row)} (1-based rows)."""
    b = wb.add_format({"bold": True})
    pct, num, dt = wb.add_format({"num_format": "0.00%"}), wb.add_format({"num_format": "#,##0.0"}), \
        wb.add_format({"num_format": "dd-mmm-yy"})
    i = wb.add_worksheet("Inputs")
    i.write(0, 0, "Riverbend Toll Road - business plan inputs", b)
    at = {}
    for r, (label, value, unit) in enumerate(inputs["rows"], start=2):
        i.write(r, 0, label)
        if isinstance(value, date):
            i.write_datetime(r, 1, value, dt)
        else:
            i.write(r, 1, value, pct if unit == "%" else num)
        i.write(r, 2, unit)
        at[label] = f"Inputs!$B${r + 1}"

    def timeline(ws, title):
        ws.write(0, 0, title, b)
        ws.write(2, 0, "Period ending", b)
        for k, d in enumerate(n["ends"]):
            ws.write_datetime(2, FIRST_COL + k, d, dt)
        ws.write(3, 0, "Financial year")
        for k, d in enumerate(n["ends"]):
            ws.write(3, FIRST_COL + k, f"FY{d.year % 100:02d}")

    rows = {}
    o = wb.add_worksheet("Operations")
    timeline(o, "Operations")
    op_rows = [("Traffic", "m trips", "traffic"), ("Toll (nominal)", "A$", "toll"), ("Toll revenue", "A$m", "revenue"),
               ("Operating costs", "A$m", "opex")]
    if insurance:
        op_rows.append(("Insurance", "A$m", "insurance"))
    op_rows.append(("EBITDA", "A$m", "ebitda"))
    for k, (label, unit, key) in enumerate(op_rows):
        rows[key] = ("Operations", 6 + k)
    for key, (sheet, r) in rows.items():
        label, unit = next((lab, u) for lab, u, kk in op_rows if kk == key)
        o.write(r - 1, 0, label)
        o.write(r - 1, 1, unit)
    for k in range(YEARS):
        c, p = COL(FIRST_COL + k), COL(FIRST_COL + k - 1)
        rr = {key: r for key, (_, r) in rows.items()}
        o.write_formula(f"{c}{rr['traffic']}", f"={at['Opening traffic']}" if k == 0 else
                        f"={p}{rr['traffic']}*(1+{at['Traffic growth']})", num, n["traffic"][k])
        o.write_formula(f"{c}{rr['toll']}", f"={at['Toll at start of forecast']}" if k == 0 else
                        f"={p}{rr['toll']}*(1+{at['CPI']})", num, n["toll"][k])
        o.write_formula(f"{c}{rr['revenue']}", f"={c}{rr['traffic']}*{c}{rr['toll']}", num, n["revenue"][k])
        o.write_formula(f"{c}{rr['opex']}", f"=-{c}{rr['revenue']}*{at['Operating costs (% of revenue)']}", num,
                        n["opex"][k])
        parts = f"{c}{rr['revenue']}+{c}{rr['opex']}"
        if insurance:
            o.write_formula(f"{c}{rr['insurance']}", f"=-{at['Insurance premium']}*(1+{at['CPI']})^{k}", num,
                            n["insurance"][k])
            parts += f"+{c}{rr['insurance']}"
        o.write_formula(f"{c}{rr['ebitda']}", f"={parts}", num, n["ebitda"][k])

    cf = wb.add_worksheet("CashFlow")
    timeline(cf, "Cash flow")
    cf_rows = [("EBITDA", "ebitda"), ("Capital expenditure", "capex"), ("Tax paid", "tax"),
               ("Unlevered free cash flow", "fcf")]
    for k, (label, key) in enumerate(cf_rows):
        cf.write(5 + k, 0, label)
        cf.write(5 + k, 1, "A$m")
        rows[f"cf_{key}"] = ("CashFlow", 6 + k)
    e_row = rows["ebitda"][1]
    for k in range(YEARS):
        c = COL(FIRST_COL + k)
        fy = n["ends"][k].year
        cf.write_formula(f"{c}6", f"=Operations!{c}{e_row}", num, n["ebitda"][k])
        major = f"+IF(MOD({fy},5)=0,{at['Major maintenance (every 5 years)']},0)"
        cf.write_formula(f"{c}7", f"=-({at['Maintenance capex']}*(1+{at['CPI']})^{k}{major})", num, n["capex"][k])
        cf.write_formula(f"{c}8", f"=-MAX(0,({c}6+{c}7)*{at['Tax rate']})", num, n["tax"][k])
        cf.write_formula(f"{c}9", f"={c}6+{c}7+{c}8", num, n["fcf"][k])
    return rows


# ---- the overlay (valuation) -----------------------------------------------------------------------------

def valuation(n: dict, vd: date, rate: float, g: float, net_debt: float) -> dict:
    ends = {k: e for k, e in enumerate(n["ends"])}
    df = dcf.factors(ends, vd, rate, "end", "actual/actual")
    tv = n["fcf"][-1] * (1 + g) / (rate - g)
    flows = [f + (tv if k == YEARS - 1 else 0) for k, f in enumerate(n["fcf"])]
    ev = sum(flows[k] * df[k] for k in range(YEARS))
    return dict(df=[df[k] for k in range(YEARS)], tv=tv, flows=flows, ev=ev, equity=ev - net_debt)


def write_overlay(wb, n: dict, cf_ref, vd: date, rate: float, g: float, net_debt: float) -> dict:
    """Val_Inputs / DCF / Summary. cf_ref(col_letter) -> formula text reading the client's free cash flow."""
    b = wb.add_format({"bold": True})
    pct, num, dt = wb.add_format({"num_format": "0.00%"}), wb.add_format({"num_format": "#,##0.0"}), \
        wb.add_format({"num_format": "dd-mmm-yy"})
    v = valuation(n, vd, rate, g, net_debt)
    vi = wb.add_worksheet("Val_Inputs")
    vi.write(0, 0, "Valuation assumptions", b)
    vi.write(3, 0, "Valuation date"); vi.write_datetime(3, 2, vd, dt)
    vi.write(4, 0, "Discount rate (post-tax nominal WACC)"); vi.write(4, 2, rate, pct)
    vi.write(5, 0, "Terminal growth rate"); vi.write(5, 2, g, pct)
    vi.write(6, 0, "Net debt at valuation date"); vi.write(6, 1, "A$m"); vi.write(6, 2, net_debt, num)

    d = wb.add_worksheet("DCF")
    d.write(0, 0, "Discounted cash flow", b)
    d.write(2, 0, "Period ending", b)
    for k, e in enumerate(n["ends"]):
        d.write_datetime(2, FIRST_COL + k, e, dt)
    labels = [(5, "Unlevered free cash flow (client model)"), (6, "Terminal value"), (7, "Valuation cash flow"),
              (9, "Discount factor")]
    for r, lab in labels:
        d.write(r - 1, 0, lab)
        d.write(r - 1, 1, "" if r == 9 else "A$m")
    last = COL(FIRST_COL + YEARS - 1)
    for k in range(YEARS):
        c = COL(FIRST_COL + k)
        d.write_formula(f"{c}5", cf_ref(c), num, n["fcf"][k])
        if k == YEARS - 1:
            d.write_formula(f"{c}6", f"={c}5*(1+Val_Inputs!$C$6)/(Val_Inputs!$C$5-Val_Inputs!$C$6)", num, v["tv"])
        else:
            d.write(f"{c}6", 0, num)
        d.write_formula(f"{c}7", f"={c}5+{c}6", num, v["flows"][k])
        d.write_formula(f"{c}9", f"=1/(1+Val_Inputs!$C$5)^YEARFRAC(Val_Inputs!$C$4,{c}$3,1)",
                        wb.add_format({"num_format": "0.0000"}), v["df"][k])
    d.write(11, 0, "Enterprise value"); d.write(11, 1, "A$m")
    d.write_formula("D12", f"=SUMPRODUCT(D7:{last}7,D9:{last}9)", num, v["ev"])
    d.write(12, 0, "Less: net debt"); d.write(12, 1, "A$m")
    d.write_formula("D13", "=-Val_Inputs!$C$7", num, -net_debt)
    d.write(13, 0, "Equity value"); d.write(13, 1, "A$m")
    d.write_formula("D14", "=D12+D13", num, v["equity"])

    lo, hi = valuation(n, vd, rate + 0.0025, g, net_debt), valuation(n, vd, rate - 0.0025, g, net_debt)
    s = wb.add_worksheet("Summary")
    s.write(0, 0, "Valuation summary (A$m)", b)
    for c, h in enumerate(["", "Low", "Preferred", "High"]):
        s.write(2, c, h, b)
    s.write(3, 0, "Enterprise value"); s.write(3, 1, round(lo["ev"], 1), num)
    s.write_formula("C4", "=DCF!D12", num, v["ev"]); s.write(3, 3, round(hi["ev"], 1), num)
    s.write(4, 0, "Equity value"); s.write(4, 1, round(lo["equity"], 1), num)
    s.write_formula("C5", "=DCF!D14", num, v["equity"]); s.write(4, 3, round(hi["equity"], 1), num)
    s.write(5, 0, "Low / high: discount rate +/- 0.25% (pasted from the sensitivity run)")
    return {**v, "low": lo, "high": hi}


def add_external_link(path: Path, target: str, sheets: list[str], cached: dict[tuple[str, str], float]) -> None:
    """Add xl/externalLinks/externalLink1.xml (with cached values) so [1]Sheet!A1 in formulas resolves,
    as Excel would have saved it. XlsxWriter writes the formulas but not the link part."""
    ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    rns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    by_sheet = {s: {} for s in sheets}
    for (s, addr), val in cached.items():
        by_sheet[s][addr] = val
    data = []
    for i, s in enumerate(sheets):
        rows = {}
        for addr, val in by_sheet[s].items():
            rows.setdefault(int(re.sub(r"\D", "", addr)), []).append((addr, val))
        body = "".join(f'<row r="{r}">' + "".join(f'<cell r="{a}"><v>{v!r}</v></cell>' for a, v in sorted(cells))
                       + "</row>" for r, cells in sorted(rows.items()))
        data.append(f'<sheetData sheetId="{i}">{body}</sheetData>')
    link = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<externalLink xmlns="{ns}" xmlns:r="{rns}">'
            f'<externalBook r:id="rId1"><sheetNames>' + "".join(f'<sheetName val="{s}"/>' for s in sheets)
            + f'</sheetNames><sheetDataSet>{"".join(data)}</sheetDataSet></externalBook></externalLink>')
    link_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.'
                 'openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.'
                 'openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" '
                 f'Target="{target}" TargetMode="External"/></Relationships>')
    src = zipfile.ZipFile(path)
    parts = {name: src.read(name) for name in src.namelist()}
    src.close()
    wb_xml = parts["xl/workbook.xml"].decode()
    wb_xml = wb_xml.replace("</sheets>", '</sheets><externalReferences><externalReference r:id="rIdExt1"/>'
                                         "</externalReferences>", 1)
    parts["xl/workbook.xml"] = wb_xml.encode()
    rels = parts["xl/_rels/workbook.xml.rels"].decode().replace(
        "</Relationships>", '<Relationship Id="rIdExt1" Type="http://schemas.openxmlformats.org/officeDocument/'
                            '2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/>'
                            "</Relationships>")
    parts["xl/_rels/workbook.xml.rels"] = rels.encode()
    ct = parts["[Content_Types].xml"].decode().replace(
        "</Types>", '<Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.'
                    'openxmlformats-officedocument.spreadsheetml.externalLink+xml"/></Types>')
    parts["[Content_Types].xml"] = ct.encode()
    parts["xl/externalLinks/externalLink1.xml"] = link.encode()
    parts["xl/externalLinks/_rels/externalLink1.xml.rels"] = link_rels.encode()
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, blob in parts.items():
            z.writestr(name, blob)


# ---- the report ------------------------------------------------------------------------------------------

def m(x: float) -> str:
    return f"{x:,.1f}"


def report_content(v: dict, n: dict, vd: date, rate: float, g: float, net_debt: float) -> dict:
    sens = [[m(valuation(n, vd, rate + dr, g + dg, net_debt)["equity"]) for dg in (-0.0025, 0, 0.0025)]
            for dr in (0.0025, 0, -0.0025)]
    return {
        "title": "Riverbend Toll Road Pty Ltd",
        "subtitle": f"Independent valuation as at {vd:%-d %B %Y}  |  Project Kestrel",
        "summary": [
            "We have been engaged by Riverbend Holdings Pty Ltd to assess the fair market value of 100% of the",
            "ordinary equity in Riverbend Toll Road Pty Ltd (the Company) as at 30 June 2025 (the Valuation Date).",
            f"We have assessed the equity value to be in the range of A${m(v['low']['equity'])}m to "
            f"A${m(v['high']['equity'])}m,",
            f"with a preferred value of A${m(v['equity'])}m. This corresponds to an enterprise value of "
            f"A${m(v['ev'])}m.",
        ],
        "summary_table": [["A$m", "Low", "Preferred", "High"],
                          ["Enterprise value", m(v["low"]["ev"]), m(v["ev"]), m(v["high"]["ev"])],
                          ["Less: net debt", f"({m(net_debt)})", f"({m(net_debt)})", f"({m(net_debt)})"],
                          ["Equity value", m(v["low"]["equity"]), m(v["equity"]), m(v["high"]["equity"])]],
        "method": [
            "Our primary valuation approach is a discounted cash flow (DCF) analysis of the unlevered free cash",
            "flows in the Company's FY25 business plan model, over the forecast period to FY45.",
            "A terminal value is calculated at the end of the forecast using the Gordon growth method.",
            "Cash flows are discounted to the Valuation Date at a post-tax nominal WACC, end of period.",
            "We cross-checked the result against the implied EV / EBITDA multiple of comparable toll roads.",
            f"The implied FY26 EV / EBITDA multiple is {v['ev'] / n['ebitda'][0]:.1f}x.",
        ],
        "assumptions": [["Assumption", "Value", "Basis"],
                        ["Discount rate", f"{rate * 100:.2f}%", "Post-tax nominal WACC"],
                        ["Terminal growth rate", f"{g * 100:.2f}%", "Long-term CPI"],
                        ["Forecast period", "FY26 - FY45", "20 years"],
                        ["Corporate tax rate", "30.00%", "Statutory"],
                        ["Net debt", f"A${m(net_debt)}m", "At Valuation Date"]],
        "sensitivity": [["Equity value (A$m)", f"TGR {(g - .0025) * 100:.2f}%", f"TGR {g * 100:.2f}%",
                         f"TGR {(g + .0025) * 100:.2f}%"]] +
                       [[f"WACC {(rate + dr) * 100:.2f}%"] + row for dr, row in zip((0.0025, 0, -0.0025), sens)],
    }


def table_png(rows: list[list[str]], width: float = 7.0) -> bytes:
    plt.rcParams["text.parse_math"] = False
    fig = plt.figure(figsize=(width, 0.42 * len(rows) + 0.3), dpi=200)
    ax = fig.add_axes([0, 0, 1, 1]); ax.axis("off")
    t = ax.table(cellText=rows[1:], colLabels=rows[0], loc="center", cellLoc="left")
    t.auto_set_font_size(False); t.set_fontsize(10); t.scale(1, 1.5)
    for (r, _), cell in t.get_celld().items():
        cell.set_edgecolor("#c4c4cd")
        if r == 0:
            cell.set_facecolor("#2e2e38"); cell.get_text().set_color("white")
    buf = io.BytesIO(); fig.savefig(buf, format="png"); plt.close(fig)
    return buf.getvalue()


def write_pdf(path: Path, c: dict) -> None:
    plt.rcParams["pdf.fonttype"] = 42  # TrueType, so the text layer can be read back
    plt.rcParams["text.parse_math"] = False  # "A$2,135.4m to A$2,475.9m" is not maths
    with PdfPages(path) as pdf:
        def page(title, n):
            fig = plt.figure(figsize=(8.27, 11.69))
            fig.text(0.08, 0.93, title, fontsize=17, weight="bold")
            fig.text(0.5, 0.03, f"Page {n}", fontsize=8, ha="center", color="#747480")
            return fig

        def text(fig, lines, y, size=10.5):
            for ln in lines:
                fig.text(0.08, y, ln, fontsize=size); y -= 0.022
            return y

        def table(fig, rows, y, h):
            ax = fig.add_axes([0.08, y - h, 0.84, h]); ax.axis("off")
            t = ax.table(cellText=rows[1:], colLabels=rows[0], loc="upper left", cellLoc="right")
            t.auto_set_font_size(False); t.set_fontsize(9.5); t.scale(1, 1.4)
            for (r, col), cell in t.get_celld().items():
                cell.set_edgecolor("#c4c4cd")
                if col == 0:
                    cell._loc = "left"

        fig = plt.figure(figsize=(8.27, 11.69))
        fig.text(0.08, 0.6, c["title"], fontsize=24, weight="bold")
        fig.text(0.08, 0.56, c["subtitle"], fontsize=13)
        fig.text(0.08, 0.52, "Prepared for Riverbend Holdings Pty Ltd  |  Final report", fontsize=10, color="#747480")
        fig.text(0.08, 0.1, "Synthetic test document: every name and number in it is fictional.", fontsize=8)
        pdf.savefig(fig); plt.close(fig)

        fig = page("1. Executive summary", 2)
        y = text(fig, c["summary"], 0.88)
        fig.text(0.08, y - 0.02, "Table 1: Valuation summary", fontsize=10, weight="bold")
        table(fig, c["summary_table"], y - 0.035, 0.14)
        pdf.savefig(fig); plt.close(fig)

        fig = page("2. Valuation approach", 3)
        y = text(fig, c["method"], 0.88)
        fig.text(0.08, y - 0.02, "Table 2: Key valuation assumptions", fontsize=10, weight="bold")
        # Pasted as a picture: no text layer, so it has to be read from the image.
        from PIL import Image
        img = Image.open(io.BytesIO(table_png(c["assumptions"])))
        ax = fig.add_axes([0.08, y - 0.3, 0.84, 0.26]); ax.axis("off"); ax.imshow(img)
        pdf.savefig(fig); plt.close(fig)

        fig = page("3. Sensitivity analysis", 4)
        y = text(fig, ["Equity value under alternative discount rate and terminal growth assumptions (A$m)."], 0.88)
        table(fig, c["sensitivity"], y - 0.02, 0.14)
        pdf.savefig(fig); plt.close(fig)


def write_pptx(path: Path, c: dict, v: dict, n: dict) -> None:
    from pptx import Presentation
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.util import Inches, Pt
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.33), Inches(7.5)
    s = prs.slides.add_slide(prs.slide_layouts[0])
    s.shapes.title.text = c["title"]
    s.placeholders[1].text = c["subtitle"]

    def slide(title):
        sl = prs.slides.add_slide(prs.slide_layouts[5])
        sl.shapes.title.text = title
        return sl

    sl = slide("Executive summary")
    tb = sl.shapes.add_textbox(Inches(0.6), Inches(1.4), Inches(12), Inches(1.6)).text_frame
    tb.word_wrap = True
    tb.text = " ".join(c["summary"])
    for p in tb.paragraphs:
        p.font.size = Pt(14)
    rows = c["summary_table"]
    t = sl.shapes.add_table(len(rows), 4, Inches(0.6), Inches(3.4), Inches(9), Inches(1.6)).table
    for r, row in enumerate(rows):
        for k, val in enumerate(row):
            t.cell(r, k).text = val

    sl = slide("Valuation approach and key assumptions")
    tb = sl.shapes.add_textbox(Inches(0.6), Inches(1.3), Inches(12), Inches(1.8)).text_frame
    tb.word_wrap = True
    for i, ln in enumerate(c["method"]):
        p = tb.paragraphs[0] if i == 0 else tb.add_paragraph()
        p.text, p.font.size = ln, Pt(13)
    sl.shapes.add_picture(io.BytesIO(table_png(c["assumptions"])), Inches(0.6), Inches(3.6), width=Inches(8))

    sl = slide("Forecast unlevered free cash flow (A$m)")
    cd = CategoryChartData()
    cd.categories = [f"FY{e.year % 100:02d}" for e in n["ends"]]
    cd.add_series("Unlevered free cash flow", [round(x, 1) for x in n["fcf"]])
    sl.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.6), Inches(1.4), Inches(12), Inches(5.5), cd)
    prs.save(path)


# ---- write everything ------------------------------------------------------------------------------------

def inputs_rows(start: date, traffic0, growth, toll0, cpi, opex_pct, capex, major, insurance=None):
    rows = [("Forecast start", start, "date"), ("Opening traffic", traffic0, "m trips"), ("Traffic growth", growth, "%"),
            ("Toll at start of forecast", toll0, "A$"), ("CPI", cpi, "%"),
            ("Operating costs (% of revenue)", opex_pct, "%"), ("Maintenance capex", capex, "A$m"),
            ("Major maintenance (every 5 years)", major, "A$m"), ("Tax rate", 0.30, "%")]
    if insurance is not None:
        rows.append(("Insurance premium", insurance, "A$m"))
    return {"rows": rows}


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    vd, rate, g, net_debt = date(2025, 6, 30), 0.0725, 0.025, 850.0
    prior_in = dict(traffic0=42.0, growth=0.02, toll0=6.50, cpi=0.025, opex_pct=0.18, capex=35.0, major=120.0)
    cur_in = dict(traffic0=43.1, growth=0.021, toll0=6.70, cpi=0.03, opex_pct=0.18, capex=38.0, major=120.0)
    prior = client_numbers(2026, **prior_in, insurance=None)
    current = client_numbers(2027, **cur_in, insurance=4.0)

    files = {}
    p = OUT / "Riverbend_BP25_client_model.xlsx"
    wb = xlsxwriter.Workbook(p)
    rows = write_client(wb, prior, inputs_rows(date(2025, 7, 1), **prior_in), insurance=False)
    wb.close(); files["prior client model"] = p

    p = OUT / "Riverbend_BP26_client_model.xlsx"
    wb = xlsxwriter.Workbook(p)
    write_client(wb, current, inputs_rows(date(2026, 7, 1), **cur_in, insurance=4.0), insurance=True)
    wb.close(); files["current client model"] = p

    fcf_row = rows["cf_fcf"][1]
    p = OUT / "Kestrel_valuation_overlay_FY25.xlsx"
    wb = xlsxwriter.Workbook(p)
    v = write_overlay(wb, prior, lambda c: f"=[1]CashFlow!{c}{fcf_row}", vd, rate, g, net_debt)
    wb.close()
    cached = {("CashFlow", f"{COL(FIRST_COL + k)}{fcf_row}"): prior["fcf"][k] for k in range(YEARS)}
    add_external_link(p, "Riverbend_BP25_client_model.xlsx", ["Inputs", "Operations", "CashFlow"], cached)
    files["prior overlay (standalone)"] = p

    p = OUT / "Riverbend_BP25_with_overlay.xlsx"
    wb = xlsxwriter.Workbook(p)
    write_client(wb, prior, inputs_rows(date(2025, 7, 1), **prior_in), insurance=False)
    write_overlay(wb, prior, lambda c: f"=CashFlow!{c}{fcf_row}", vd, rate, g, net_debt)
    wb.close(); files["prior client model with overlay inside"] = p

    content = report_content(v, prior, vd, rate, g, net_debt)
    p = OUT / "Riverbend_valuation_report_FY25.pdf"
    write_pdf(p, content); files["prior report (PDF)"] = p
    p = OUT / "Riverbend_valuation_report_FY25.pptx"
    write_pptx(p, content, v, prior); files["prior report (PPTX)"] = p

    for role, f in files.items():
        print(f"{role:40s} {f.relative_to(ROOT)}")
    print(f"\nprior: EV A${m(v['ev'])}m, equity A${m(v['equity'])}m "
          f"(range {m(v['low']['equity'])} - {m(v['high']['equity'])}), rate {rate:.2%}, TGR {g:.2%}")


if __name__ == "__main__":
    main()
