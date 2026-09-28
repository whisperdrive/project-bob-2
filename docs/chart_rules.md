# Chart rules

The chart reviewer reads this file on every review (`bench/chartreview.py`), so edits here apply to the next
chart without restarting the app. Each rule has an ID; the reviewer cites it in every issue it raises
(e.g. "[F1] one value is 37x the next largest"). Keep rules short, testable, and within what the reviewer can
change.

**What the reviewer can change:** the title, the chart type, the default visible periods, the y-axis range, and a
note under the chart. **What it can't change:** the numbers, which series are shown, colours, or number
formats. If a rule needs something outside that list, the reviewer reports it as an issue and leaves it.

## P. Presentation applied automatically (don't change or flag these)
The app applies these before the review; the data summary tells you which apply.
- **P1 Period labels.** Monthly and quarterly periods are labelled by their end month, MMM-YYYY
  ("Sep-2017"); annual periods by year, "FY2018" when the model's financial year doesn't end in December.
- **P2 Signs.** If every value on the chart is negative (e.g. capex), values are shown as positive and the
  chart says so. If positives and negatives are mixed (e.g. sources and uses of funds), signs are kept, so
  negatives read in context with the positives.
- **P3 Periodic or annual.** The viewer can switch between the model's own periods and annual figures by
  financial year: flows are summed, balances taken at year end (opening balances at the start), rates,
  percentages and indices averaged; partial years are marked *. You review the periodic view; your visible
  range carries over to whole years in the annual view, your y-axis limits don't.

- **P4 Phases.** Where the model flags periods as Actuals, Business plan or Forecast, those spans are shaded
  and labelled on the chart. Don't frame the chart to hide the actuals.

## A. Accuracy (always applies)
- **A1 Data is never altered.** The chart shows the workbook's values. Framing may hide some periods or cut
  the y-axis, but never changes a value.
- **A2 Say what's hidden.** If the default view leaves out any period or cuts off any point, the note says
  what (value and period) and that "Full range" shows everything.
- **A3 Units only from the workbook.** State units only if the data summary gives them. If it doesn't, the
  note says "Units aren't labelled in the workbook." Never infer units (A$, A$'000, A$m) from the size of
  the numbers.
- **A4 Periods as labelled.** Refer to periods by the labels on the axis (e.g. 2067-04-01), not by guessed
  financial years or quarters.
- **A5 Actuals and forecast together.** A chart over time shows the whole timeline, actual and forecast
  periods, unless the user asked for part of it. The app shades the model's Actuals / Business plan /
  Forecast periods (P4); if a series has values for only part of the timeline, raise an issue.

## T. Title
- **T1 Say what is measured**, in plain words, using the line item's name where it's clear
  (e.g. "Valuation cash flows", "Equity cash flows vs total valuation cash flow").
- **T2 Sentence case, no more than 60 characters**, no trailing full stop.
- **T3 No dates** unless they match the visible range exactly. Prefer no dates: the axis shows them.
- **T4 No judgements** in the title ("strong growth", "worrying dip"). The title describes; the note explains.

## K. Chart type
- **K1 Line** for a series over time with more than 12 periods.
- **K2 Bar** for 12 periods or fewer, or when values are totals for discrete periods (e.g. annual figures).
- **K3 Keep the user's choice** if they asked for a specific type, unless it makes the data unreadable; then
  keep it and say why in an issue.
- **K4 Parts of a total stay stacked.** Stacked columns, stacked areas and combos (stacked columns with lines
  for a total or last year's figure) show how parts make up a whole; keep them, whatever K1 and K2 say. Raise
  an issue if a stacked series doesn't add to the others (a rate, a balance beside flows, two versions of one
  line item): those belong on a line.
- **K5 Bridges stay waterfalls.** A waterfall walks from a start value through the steps to an end value; keep
  it. Its bars drawn from zero (the start, subtotals, the end) are listed in waterfall_totals.
- **K6 A second axis** (series marked right axis) is for a rate or other units beside amounts; S2 is met when
  the units differ only between the two axes.

## F. Framing (visible range and y-axis)
- **F1 Outliers.** A value is an outlier if its absolute size is more than 5x the next largest absolute
  value in the same series. Typical in these models: a terminal value in the last years of the timeline.
  - At the start of the timeline: start the visible range just after it (x_start = its index + 1).
  - At the end of the timeline: end the visible range just before it (x_end = its index - 1).
  - Inside the timeline: set y_max (or y_min) about 10% beyond the largest remaining value.
  - Either way, the note gives its value and period (A2).
- **F2 Leading and trailing empty periods.** If more than 4 periods at the start or end are all zero or blank
  in every series, leave them out of the default view.
- **F3 Zero.** Bar charts always include zero on the y-axis. Line charts include zero whenever the data
  crosses it or sits within 20% of it.
- **F4 Don't over-crop.** The default view keeps at least half of the periods (the app enforces this). Never cut off more than one period's value per series without a strong reason
  stated in an issue. Never hide a sign change (e.g. a large negative payment) the user would need to see.
- **F5 One limit at a time.** Prefer the least change that makes the trend readable: a visible range before
  a y-axis cap; don't set both unless one alone fails.

## S. Several series
- **S1 Comparable scales.** If one series is more than 10x another over the visible range, raise an issue:
  the smaller one reads as flat. (The reviewer can't split the chart; the chat can redraw it.)
- **S2 Same units.** Series with different units shouldn't share an axis. Raise an issue if the labels
  suggest different units (e.g. % and A$).

## N. Notes
- **N1 Only when needed:** hidden data (A2), missing units (A3), an outlier (F1), or a sign convention the
  viewer must know (e.g. negatives are payments out).
- **N2 At most two sentences**, plain language, no judgement.
- **N3 Numbers readable:** thousands separators, rounded to what matters (3,691,698, not 3691698.0423;
  percentages to one decimal place).

## R. Review behaviour
- **R1 Least change.** If the chart already follows these rules, return verdict "ok" with no changes.
- **R2 Cite rules.** Every issue starts with the rule ID, e.g. "[F1] 2067-04-01 is 37x the next largest".
- **R3 Second look.** On the second pass, check the re-drawn chart against the same rules; only change what
  still breaks one.
