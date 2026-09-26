# Excel workbook reading benchmark

Tests ways to let an AI agent read a large Excel financial model without sending it whole. The reference
case is a quarterly infrastructure valuation model: 20 sheets, ~420k formulas, a 222-column timeline.
Client workbooks go in `reference/`, which git ignores.

## Run order (works on any .xlsx / .xlsm)
```bash
W="reference/valuation_model.xlsm"
uv run python bench/census.py "$W"      # per-sheet formula/constant counts
uv run python bench/build_map.py "$W"   # row map + SQLite store (~15s here, run once per workbook)
uv run python bench/compare.py "$W" "reference/valuation_model.xlsb"   # optional .xlsb copy
```
Outputs go to `out/<workbook name>/` (`census.json`, `map.txt`, `model.db`, `compare.json`).
`.xlsb`/`.xls` can be benchmarked by compare.py but build_map needs `.xlsx`/`.xlsm` for formulas.

Edges (`bench/edges.py`) link each line item to the rows it reads, with a kind: `direct` (a plain
reference), `offset` (the range an OFFSET actually points at, worked out from saved values), `active` (the
row a SUMIFS / INDEX-MATCH / CHOOSE selects in the current scenario) or `inactive` (a candidate it
considers but doesn't select). `trace` follows the active path and counts the inactive candidates.
`uv run python bench/edges.py out/<dir>/model.db` upgrades an older model.db in place.

`bench/tools.py` holds the agent tools. Call `tools.use(<workbook>)` first, then `overview()`, `find(text)`,
`rows(sheet, r1, r2)`, `trace(sheet, row, "up"|"down")`, `cells(sheet, A1, A1)`, `sql(query)`.

### How it adapts to a workbook (bench/layout.py)
Per sheet, from cached values: the **timeline** is the row with the most dates (longest run of date
columns, periodicity from the date gaps); the **label column** is the most text-heavy column left of it;
the **units column** is a column right of that with short, repeated text. Runs of 50+ same-shaped
formula-free rows are summarised as one TABLE line. compare.py picks the "headline output" itself: the
line item with the biggest upstream tree that nothing depends on (skipping model-check rows).

`bench/make_test_workbook.py` writes `tests/sample_model.xlsx`, a differently laid-out model (labels in A,
monthly timeline C:N with dates in row 3, flat data table) used to check nothing is workbook-specific.

## Model Desk app (upload, identify, compare, ask)
```bash
uv run python bench/llm.py                          # one-time Entra sign-in (device code) + smoke test
uv run uvicorn app.server:app --port 8000           # then open http://localhost:8000
```
Drop .xlsx/.xlsm files on the page. Each file is:
1. **Fingerprinted** (SHA-256 of its bytes, computed in the browser first). A file seen before, under any
   name, is not uploaded or processed again.
2. **Built** into `out/<stem>__<sha8>/model.db` by build_map.py, with per-sheet progress in the UI. Each
   version keeps its own folder, so older versions stay available.
3. **Identified** (`bench/identify.py`): SQL collects candidate cells (named ranges like `Val_date`, rows
   labelled "Valuation date", sheet-header text), then one gpt-4o call picks the target, code name and
   valuation date and cites the cell. Shown as "Check" until the user confirms or corrects it.
4. **Compared** (`bench/diff.py`) with the latest earlier file for the same target (or, failing that, the
   same file name without dates/v2/final). Line items are aligned per sheet like a text diff, so inserted,
   removed and renamed rows don't shift everything else. Reports inputs changed (with timeline period),
   key outputs that moved, formula changes, hard-code/formula swaps, sheets and named ranges, plus
   warnings when a file looks un-recalculated (inputs changed but no results moved, the valuation-date input
   disagrees with the calculated date, or formula results are blank). gpt-4o writes a short summary.

The chat (Ask tab) gets the confirmed target, valuation date and change summary as context. Tool calls are
chosen by the model on Azure but run locally (`agent._run_tool` against model.db); only their text results
are sent back. The `chart` tool takes cell ranges, reads the exact values locally and the page draws them
with Chart.js (hover, line/bar, click-to-zoom, copy data); the model only sees a first/last/min/max summary.
Before a chart is shown it's reviewed on the server: `bench/chartrender.py` draws it to PNG with matplotlib (no
browser), a vision model (`bench/chartreview.py`, gpt-4o) checks the image against an exact data summary and
suggests presentation fixes (title, line/bar, visible window, y-axis range, a note, e.g. for an outlier that
flattens the trend), and it looks once more at the re-rendered result. Titles naming years outside the visible
range are corrected in code. The data can't be changed; "Full range" undoes the framing, and the chat model is
told what was changed so its answer matches. `find` ignores spaces and punctuation ("cash flow" finds
"Cashflow"). "Copy data" falls back to a selectable panel where the browser blocks clipboard access.

DCF recompute: the `dcf` tool (`bench/dcf.py`) redoes a valuation's discounting in plain Python from the
workbook's saved cash-flow rows: PV to the valuation date, end- or mid-period, Excel `YEARFRAC` actual/actual or
actual/365 (what `XNPV` uses), an optional cut-off date, and a bridge to equity or enterprise value (cells, numbers,
or a row's amount in the valuation-date period, like a `SUMIFS` on the date row). With `compare_to` it checks
the result against the workbook's own value. When they differ, it follows that cell's formula to the model's
`SUMPRODUCT` and discount factor row. If that formula shows the cause (the wrong cash-flow row, a cut-off, the
convention, a valuation-date amount), it corrects the run and lists every correction; otherwise it reports the
mismatch. `rates` gives what-if values at other discount rates, with the cash flows and bridge held fixed; the
cash flows themselves aren't recalculated. Called without a cash-flow row, it lists cells that look like
valuation results, with their formulas; without a rate or valuation date, it lists candidate cells. The chat
shows each result as a card.

Valuation tab: finds every DCF in a workbook from its formulas and recomputes it, with no model calls
(`bench/valuation.py`, `GET /api/files/{id}/valuation`). A DCF is a cell labelled like a valuation (enterprise
value, equity value, NPV, total valuation) whose formula is, or adds amounts to, a `SUMPRODUCT` of a
cash-flow row and a discount-factor row. The assumptions are read back from the discount factors: for each
candidate valuation date and convention, the rate is back-solved from one factor and kept only if it reproduces
every factor. The cut-off is the last period with a non-zero factor, and the rate is named by the cell holding that
value. The value cell's formula is split into additive terms: the PV, other cells (debt, cash) and `SUMIFS`
picks of one period's amount. The tab shows three things:
- Low, model and high values. The low and high discount rates are editable and default to the model's rate ± 1
  percentage point.
- A chart of the cash flows and their present values over the discounted periods. A final-period outlier such as
  a terminal value is left out of the default view and named in a note.
- A conclusions table (PV, bridge, value and the difference from the model, for each case) and an assumptions
  table giving each assumption's source cell.

Valuation cells that can't be reproduced, for example a result rescaled by a units divisor or a rolling
`XNPV`, are listed with the reason. Low and high change only the discount rate, so levers that move the cash
flows (growth, CPI, terminal assumptions) need the Python rebuild.

Token usage: every model call (chat, identify, change summary) is logged to the `usage` table in
`out/registry.db` (`bench/usage.py`). The header shows this chat and all-time totals; click it for totals by
purpose, model, day and recent chats. A chat session starts with the first question and ends with "New chat".
Costs use Azure list prices (`bench/pricing.py`, from prices.azure.com, East US, 2026-09-25).

Rate limits: `bench/ratelimit.py` keeps every call 10% below each deployment's tokens/min and requests/min
(`RATE_LIMIT_BUFFER=0.05` for 5%). Limits are capacity x 1,000 TPM with Microsoft's per-model RPM ratios;
capacities are read from the Foundry project at startup. Calls that would exceed the budget wait, and the
chat shows the pause. The app offers gpt-4o (default), gpt-6-sol, gpt-6-luna, gpt-4o-mini and gpt-5-nano; gpt-4 (gpt-4.1),
o3-mini and DeepSeek-V4-Flash were dropped because their capacities (8k-20k TPM) were too low for tool use.
Registry: `out/registry.db`; uploaded originals: `uploads/<sha12>/`. `bench/library.py` runs the pipeline on
one background worker thread and re-queues anything interrupted by a restart.

## Results (tokens are tiktoken approximations, not Claude's tokenizer)
| Method | Load | Content | Tokens |
|---|---|---|---|
| 1 calamine (.xlsb) raw dump | 0.2s | values only | 3.3M |
| 2 pyxlsb (.xlsb) raw dump | 1.8s | values only | 3.4M |
| 3 openpyxl (.xlsm) raw dump | 4.2s | formulas only | 10.5M |
| 6a row map, full `map.txt` | 15s build | formulas (R1C1 patterns) + samples | 303k |
| 6b overview + 5 tool calls ("what drives equity value?") | <0.1s/call | on demand | 5.5k content |

6b is hand-scripted. Measured in the real loop (`bench/agent.py`, Foundry usage counts; every turn
re-sends the ~4.4k-token overview plus all earlier tool results):

| Question (gpt-4.1) | Model turns | Tool calls | Input tokens | Output tokens |
|---|---|---|---|---|
| "What drives equity value?" | 11 | 16 | 72.8k | 1.2k |
| "What discount rate is used, and where is it set?" | 7 | 12 | 39.5k | 0.4k |

Foundry caches the repeated prefix automatically: with gpt-4o, ~4.7-4.9k of each turn's input was
reported as cached (the overview), and a follow-up answered from chat history took 1 turn / 5k tokens.
The app shows cached tokens per question. gpt-4o answered in ~8s vs ~100s for gpt-4.1 (gpt-4.1 was more thorough).

## Not yet tested
- xlwings (live Excel): only needed for what-if recalculation.
- sheetwise / ks-xlsx-parser / ohh-my-excel: column-oriented compression built for data tables, a poor
  fit for a timeline formula model; worth one baseline run.

## Known limits
- Layout detection is heuristic: a category column can be taken for units (Bridge!E), and sheets
  with several side-by-side blocks get a single label column.
- Only one timeline per sheet; a sheet with no date row is treated as a plain list.
