"""Render a chart spec (from tools.chart) to PNG on the server, for the chart review. No browser involved.

Styled like the in-app chart (same series colours, light grid, sparse x labels) so the reviewer sees what the
user will see, and it applies the same framing (visible window, y-axis limits) when a view is given.
"""
import io

import matplotlib

matplotlib.use("Agg")  # no display
import matplotlib.pyplot as plt  # noqa: E402

COLOURS = ["#1d6b47", "#3f6fb0", "#b7791f", "#9b4d8f", "#2f8f8f", "#a33a22"]


def render_png(spec: dict, view: dict | None = None, width_px: int = 1200, height_px: int = 560,
               mode: str = "periodic") -> bytes:
    from chartdata import display
    shown = display(spec, mode)  # formatted period labels, sign presentation applied
    labels, series = shown["labels"], shown["series"]
    n = len(labels)
    fig, ax = plt.subplots(figsize=(width_px / 100, height_px / 100), dpi=100)
    kind = spec.get("kind") or "line"
    num = lambda v: v if isinstance(v, (int, float)) else float("nan")
    zero = lambda v: v if isinstance(v, (int, float)) else 0.0
    bridge = []
    if kind == "waterfall":  # one series as a bridge: floating bars from the running total
        totals, run = set(spec.get("totals") or []), 0.0
        for i, v in enumerate(zero(v) for v in (series[0]["data"] if series else [])):
            if i in totals:
                bridge.append((0.0, v, "#3f6fb0"))
                run = v
            else:
                bridge.append((run, run + v, "#1d6b47" if v >= 0 else "#a33a22"))
                run += v
        if not totals:
            bridge.append((0.0, run, "#3f6fb0"))
            labels = list(labels) + ["Total"]
        n, series = len(labels), []
    xs = list(range(n))
    span = [[] for _ in xs]  # per period: the values the y-axis has to show
    for i, (lo, hi, c) in enumerate(bridge):
        ax.bar(i, hi - lo, bottom=lo, color=c, width=0.7)
        span[i] += [lo, hi]
    left = [s for s in series if s.get("axis") != "right"]
    right = [s for s in series if s.get("axis") == "right"]
    stacking = kind in ("stacked", "combo", "area")
    as_line = lambda s: kind == "line" or (kind in ("stacked", "combo", "area") and s.get("as") == "line")
    cols = [s for s in left if not as_line(s)]
    width = 0.8 / max(1, len(cols)) if kind == "bar" else 0.8
    pos, neg = [0.0] * n, [0.0] * n
    areas = [s for s in left if kind == "area" and not as_line(s)]
    if areas:
        ys = [[zero(v) for v in s.get("data", [])] for s in areas]
        ax.stackplot(xs, *ys, colors=[COLOURS[series.index(s) % len(COLOURS)] for s in areas], alpha=0.55,
                     labels=[s.get("name") for s in areas])
        for k, t in enumerate(map(sum, zip(*ys))):
            span[k] += [0.0, t]
    for i, s in enumerate(series):
        c, vals = COLOURS[i % len(COLOURS)], s.get("data", [])
        if s.get("axis") == "right" or s in areas:
            continue
        if as_line(s):
            ys = [num(v) for v in vals]
            ax.plot(xs, ys, color=c, linewidth=1.8, label=s.get("name"), marker="o" if n <= 60 else None, markersize=3)
            for k, y in enumerate(ys):
                if y == y:
                    span[k].append(y)
        elif stacking:
            ys = [zero(v) for v in vals]
            ax.bar(xs, ys, bottom=[pos[k] if y >= 0 else neg[k] for k, y in enumerate(ys)], width=width, color=c,
                   label=s.get("name"))
            for k, y in enumerate(ys):
                pos[k], neg[k] = (pos[k] + y, neg[k]) if y >= 0 else (pos[k], neg[k] + y)
                span[k] += [neg[k], pos[k]]
        else:
            j = cols.index(s)
            ax.bar([x + (j - (len(cols) - 1) / 2) * width for x in xs], [num(v) for v in vals], width=width, color=c,
                   label=s.get("name"))
            for k, v in enumerate(vals):
                span[k] += [0.0, zero(v)]
    if right:
        ax2 = ax.twinx()
        for i, s in enumerate(series):
            if s.get("axis") == "right":
                ax2.plot(xs, [num(v) for v in s.get("data", [])], color=COLOURS[i % len(COLOURS)], linewidth=1.8,
                         linestyle="--", label=f"{s.get('name')} (right axis)")
        ax2.tick_params(axis="y", labelsize=8, colors="#5d6a64")
        for side in ("top", "left"):
            ax2.spines[side].set_visible(False)
    # Actuals / Business plan / Forecast spans, shaded like the page (clipped to the visible window;
    # labels only where the span is wide enough to read)
    vw = view or {}
    lo_v = vw.get("x_start") or 0
    hi_v = vw.get("x_end") if vw.get("x_end") is not None else n - 1
    for k, ph in enumerate(shown.get("phases") or []):
        a, b = max(ph["start"], lo_v), min(ph["end"], hi_v)
        if a > b:
            continue
        if k % 2 == 0:
            ax.axvspan(a - 0.5, b + 0.5, color="#1b2320", alpha=0.05, linewidth=0)
        if (b - a + 1) >= 0.08 * (hi_v - lo_v + 1):
            ax.text((a + b) / 2, 1.0, ph["name"], transform=ax.get_xaxis_transform(), ha="center", va="bottom",
                    fontsize=8, color="#5d6a64", clip_on=False)
    ax.set_title(spec.get("title") or "", loc="left", fontsize=13, fontweight="bold", pad=16)
    units = spec.get("units") or "units not labelled"
    if spec.get("sign") == -1:
        units += " (negative values shown as positive)"
    ax.set_ylabel(units, fontsize=9, color="#5d6a64")
    step = max(1, n // 10)
    ax.set_xticks(xs[::step])
    ax.set_xticklabels([str(labels[i]) for i in xs[::step]], fontsize=8, color="#5d6a64")
    ax.tick_params(axis="y", labelsize=8, colors="#5d6a64")
    ax.yaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda v, _: f"{v:,.0f}" if abs(v) >= 10 else f"{v:,.3g}"))
    ax.grid(axis="y", color="#d8ded9", linewidth=0.8)
    for side in ("top", "right", "left"):
        ax.spines[side].set_visible(False)
    if len(series) > 1:
        handles, names = ax.get_legend_handles_labels()
        if right:
            h2, n2 = ax2.get_legend_handles_labels()
            handles, names = handles + h2, names + n2
        ax.legend(handles, names, fontsize=8, frameon=False)
    v = view or {}
    if v.get("x_start") is not None or v.get("x_end") is not None:
        ax.set_xlim(v.get("x_start", 0) - 0.5, (v.get("x_end") if v.get("x_end") is not None else n - 1) + 0.5)
    # Like the page: the y-axis fits the visible periods, unless the view sets limits.
    lo_i = v.get("x_start") or 0
    hi_i = v.get("x_end") if v.get("x_end") is not None else n - 1
    visible = [x for k in range(lo_i, min(hi_i, n - 1) + 1) for x in span[k]]
    if visible:
        lo, hi = min(visible), max(visible)
        pad = (hi - lo) * 0.05 or abs(hi) * 0.05 or 1
        y_lo = v["y_min"] if v.get("y_min") is not None else lo - pad
        y_hi = v["y_max"] if v.get("y_max") is not None else hi + pad
        ax.set_ylim(y_lo, y_hi)
    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", facecolor="white")
    plt.close(fig)
    return buf.getvalue()
