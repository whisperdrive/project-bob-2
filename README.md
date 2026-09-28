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
   removed and renamed rows don't shift everything else. When the timeline has moved (next year's model), each
   period is compared with the same period, matched by the timeline row's dates, instead of the same column.
   Reports inputs changed (with timeline period), key outputs that moved, formula changes, hard-code/formula
   swaps, sheets and named ranges, plus
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

Valuation tab: works on the workbook's DCFs with no model calls (`bench/valuation.py`). A DCF is a cell labelled
like a valuation (enterprise value, equity value, NPV, total valuation) whose formula is, or adds amounts to, a
`SUMPRODUCT` of a cash-flow row and a discount-factor row. The assumptions are read back from the discount factors:
for each candidate valuation date and convention, the rate is back-solved from one factor and kept only if it
reproduces every factor. The cut-off is the last period with a non-zero factor, and the rate is named by the cell
holding that value. The value cell's formula is split into additive terms: the PV, other cells (debt, cash) and
`SUMIFS` picks of one period's amount. The tab has two steps:
1. **Validate** (`GET /api/files/{id}/valuation`): can the client's anchor values be reproduced? It lists every
   anchor with the workbook's value, Python's value and the result, and states the approach found in plain words.
   A step-by-step check compares the cash-flow total with the row's own total column, the discount factors period by
   period, the PV with the workbook's `SUMPRODUCT` cell, and the bridge items and anchor value. It also charts the
   cash flows and their present values and shows the model's assumptions with source cells. Cells that can't be
   reproduced, for example a result rescaled by a units divisor or a rolling `XNPV`, are listed with the reason.
2. **Scenarios** (`POST /api/files/{id}/valuation/scenario`, unlocked once the anchor is reproduced): choose the
   discount rate, valuation date, end- or mid-period discounting, day count (`YEARFRAC` actual/actual or
   actual/365 as in `XNPV`), cut-off (the model's, none, or another date), which bridge items to include, and low /
   high rates. It shows the validated model value, the scenario value and the difference, a model-vs-scenario table
   with each changed assumption, the rate sensitivity around the scenario, and a chart of cumulative present value
   for both.

Scenarios re-discount the saved cash flows. Levers that move the cash flows (growth, CPI, terminal assumptions)
need the Python rebuild.

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

## Valuation Desk (recurring valuation engagements)
```bash
uv run python tests/make_engagement_pack.py         # optional: a synthetic 4-file pack in tests/engagement_pack/
uv run uvicorn engage.server:app --port 8002        # then open http://localhost:8002
```
One engagement per asset and year, built from last year's final report, last year's client model, last year's
overlay (the workings that take the client model to the report's conclusions) and this year's client model. The
steps, each checked by a person before the next relies on it:
1. **Files.** Workbooks go through the Model Desk pipeline and library (`bench/library.py`), so a model uploaded in
   either app is built once. Reports (PDF, PPTX) are read by `bench/docingest.py` into Markdown with page markers.
   A PDF page is read in reading order, not line by line across the page. The page is cut into blocks (XY cut)
   down any tall, clear gutter with running text on both sides, so two columns, a sidebar, or text beside a
   picture are read one after the other instead of interleaved. Tables are cut out first and never split into
   columns, and label-value lists stay on their lines. Short bold lines are side headings. Letter-spaced text
   (glyphs set one at a time, as some exporters write tables) is rebuilt into words, so a row reads "Net financial
   debt 5,223.0", not "N e t f i n a n c i a l …". No model call and no
   extra service: Azure Document Intelligence's layout model reads layout better and is an option for later.
   Every table is cropped, rendered at 200 dpi and transcribed by a vision model, then checked. A table with a text
   layer must match it (every number on the page, each row's numbers on one line in order, the same label words,
   so "$m" for "A$m" is caught). When it does, no second read is needed. A picture-only table, or one that fails
   the check, gets a second, independent read by the reviewer model, compared number by number. The tables most
   likely to hold the key figures (up to 12, ranked by valuation words in the table and on its page) are read
   first. The report then counts as read, so its facts and the model steps start. The other tables are read in
   the background: until one is, its text comes straight from the PDF, which facts can already quote. PPTX tables and chart data are read from the file. A table that fails goes through the
   **review loop** (below); anything the loop can't settle is "Check": the page shows the image beside the
   transcription and every round, and the person approves, takes the reviewer's read or edits it. Edits are
   re-checked against the page.
2. **Report reference.** `bench/reportfacts.py` extracts the target, valuation date, conclusions (preferred value and
   range), assumptions (discount rate and basis, terminal growth or exit / RAB multiple, ...), approach and the
   sensitivity grid, each with a page and a verbatim quote. Code checks each one: the quote is on that page, the
   values are in the quote, and quotes from unsettled tables are marked. A reviewer model accepts, corrects or
   rejects each fact and lists what was missed. None of this needs a button: once a report is read, the table
   loop, the extraction, the review and the fact loop start by themselves (opening an engagement also picks up
   any report they haven't run on yet). Every fact still open then goes through the review loop. Facts the
   two models agree on, and that pass the checks, are approved by the agents; facts they agree to withdraw are
   rejected by them. Both are marked as the agents' decision, and a person can undo either (an undone decision
   stays theirs). The person decides what the agents escalated: edit, take the reviewer's latest correction, or
   reject. If a table is later changed and an agents' approval no longer passes the checks, it goes back to the
   person.

   **The review loop** (`docingest.resolve_tables`, `reportfacts.resolve`) lets the two models settle problems
   between themselves, up to three rounds. The extractor (gpt-6-luna) answers each open point: it revises, keeps
   or withdraws a fact with a reason, or corrects a table transcription. Code re-checks the answer (a table
   against the page's text layer, a fact's quote against its page), and the reviewer (gpt-6-sol) accepts or
   objects again. A fact is "agreed" only when the reviewer accepts and the checks pass. The code checks tolerate a
   report's text layer: where letters are spaced or table cells run together, a figure still counts if it's there
   once spacing is ignored, read by its own shape (so 5,223.0 is found in "5,223.05,223.0"). The sign must still
   match, and the check says it ignored spacing. Whatever the two models still can't settle goes to an **arbiter**:
   a third model (gpt-4o by default, set in the header). It can waive a check that failed on a clerical point, with
   a note kept on the fact and shown beside the check. It can also take the reviewer's correction or keep the
   extractor's version, but only if the checks then pass. Otherwise it hands the fact to the person with its
   note. When the checks improve, existing facts are checked again once (no model calls), and facts that were held
   up only by a check settle.

   **What the agents learn** (`bench/lessons.py`). After each loop the reviewer turns what went wrong and how it was
   fixed into rules about method: where to look, how to read, what to check. Code enforces the anonymity. A
   lesson that names anything from the engagement or carries a figure is sent back once to be rewritten, then
   dropped. That covers the target, project, client and file names, and the report's own names (words capitalised
   mid-sentence). A lesson the loop confirms again is reinforced rather than repeated. The agents cite the rule
   IDs they apply, so each lesson shows how often it was used. Every report prompt reads the rules fresh: the
   table reads, extraction, review and both sides of the loop. Curated rules (R1, R2, ...) live in
   `docs/report_rules.md` in the repo. Learned lessons (L1, L2, ...) stay on each machine in `out/lessons.json`.
   The page lists both, and a person can retire a lesson or promote it into the rules file.
3. **Roles** (`bench/roles.py`). The overlay is picked as a sheet list, because it sometimes sits inside a copy of
   the client model. The first suggestion comes from the workbooks alone, as soon as they're read
   (`bench/likeness.py`), which compares every pair of workbooks. It measures the sheet names, line items (sheet
   and label) and formula shapes they share, ignoring years, period labels and row shifts, so last year's and this
   year's client model come out nearly the same. It also measures how much of each workbook the other contains.
   - Two workbooks that are mostly alike are one client model twice.
   - One that holds all of another plus extra sheets carrying valuation work is a client model with the overlay
     added, and the extra sheets are the overlay.
   - One unlike the others, with valuation vocabulary (valuation range, WACC, gearing, time-weighted average,
     beta, terminal value, ...), charts, external links or the adviser's name, is a standalone overlay. The
     adviser's name is searched in the text, sheet names and file properties when it's set as
     `VALUATION_DESK_OVERLAY_MARKERS` in `.env`, so it stays out of the repo.

   Which version is earlier is decided by the identified valuation date, then the date in the file name
   ("20250523 …", "Jun 25", "FY26", "BP25"), then the timeline's first period. When the overlay sits in a copy of a
   client model that is also uploaded as its own file (v2.1 = v2.0 plus an overlay sheet), the client's own file
   is the prior client model. The overlay's copy of the client sheets is then fed from it, so feeding the prior
   model checks the copy matches the file. Each suggestion lists plain checks (✓ / ✗ / ?) and gets a second
   opinion from the reviewer model on the same evidence. The model can disagree, with reasons, and its assignment
   can be taken with one click; nothing is assigned until the person confirms. Only one suggestion runs at a
   time, and a workbook's external-link tables are built while it's processed, so suggestions don't write to a
   model.db that others are reading.

   The report's facts then add the stronger evidence: an overlay sheet holds the report's conclusions or
   valuation-only assumptions (label and value must both agree), or a DCF that `valuation.py` reproduces, or reads
   such a sheet. A DCF that another version of the model also has is the client's own, not the overlay. The prior
   client model is the file the overlay's external links point to. `bench/extlinks.py` reads `xl/externalLinks`
   and checks the link's cached values against the file. Prior vs current is decided by timeline start, within
   the versions of one model. The suggestion redoes itself whenever the files or the report facts change, until
   the person confirms. The page shows the similarity of every pair of workbooks and what each one contains.
4. **Compare.** `diff.py` between the client models, without the overlay sheets. When the timeline has rolled
   forward, each period is compared with the same period (matched by the timeline row's dates), not the same
   column.
5. **Map** (`bench/linkmap.py`). The chain runs report figure → overlay cell (matched to the printed precision,
   allowing for A$m vs A$ and sign), then overlay row → prior client row (external link or same-workbook
   reference), then the same line item in the current model, with values for the same periods. The overlay's
   DCFs are also recomputed in Python. **Inside each model** gives a dashboard for the prior overlay, the prior
   client model and the current client model (`bench/modeldash.py`, the same views as the separate model
   dashboard): formula cells and inputs by sheet, which sheets feed which, the most-read rows, sheets, searchable
   line items and named ranges. Rows the map or the Python overlay uses are tagged: report figures, levers,
   outputs, the client rows the overlay reads and this year's matches.

6. **Python overlay** (`bench/overlay.py`). The overlay sheets are compiled into a Python module, and three checks
   show it reproduces them:
   - every formula cell is recomputed from the overlay's own inputs and compared with the value Excel saved;
   - the module is fed from the prior client model file instead of the link's cached values, and the results must
     not move;
   - the results must tie to the report's conclusions at the printed precision, and each sensitivity in the report
     (for example "WACC 7.50%, TGR 2.25%") is rerun through the discount-rate and growth levers.

   The page runs the module live. Levers are the report's assumptions located in the overlay, and any other input
   can be found by search and changed. Results update as you type. There are three feeds: the workbook as saved, the
   prior client model, or the current client model **rolled forward**. Rolling forward moves the overlay's period
   dates on by the roll (by default, how far the client timeline moved) and sets the valuation-date lever to the new
   date. Client values come from the same line item (sheet and label) in the period with the rolled date. Values
   that can't be matched are listed, not zeroed. Each output shows its Python function and the client values it
   reads. The module is saved as `out/overlays/e<id>/overlay.py`; open it from the page, or run it from a terminal:
   `uv run python bench/overlay.py <id> --mode current --set Val_Inputs!C5=0.075`.

   The step has three tabs, bringing the Model Desk's tools to the live module:
   - **Run it live**: the feeds, levers and inputs above.
   - **Valuation (DCF)**: the Model Desk's Validate step on the overlay's saved values, then the same DCF on the
     module's own numbers. It uses the feed and input changes from Run it live and is checked against the
     module's anchor cell. The module and `dcf.py` are separate calculations, so their agreement checks both. It
     can also run under another discounting method: rate, valuation date, end or mid period, day count, cut-off,
     bridge items and low / high rates, with the cumulative PV chart. This adds the cash-flow levers (growth,
     CPI, the roll-forward) that the Model Desk's Scenarios step couldn't reach. `rodb.patched()` shows the
     module's values to `dcf.py` through a temporary view over model.db.
   - **Ask**: the Model Desk's chat agent (`bench/overlay_chat.py` on `bench/agent.py`). It has the workbook tools
     on any of the three workbooks, and tools that run the overlay: `overlay_run`, `overlay_chart` (a row saved
     vs recomputed, several feeds on one chart), `overlay_dcf`, `overlay_formula` and `overlay_inputs`. It knows
     the report's key facts, the levers, the outputs and the feeds. Charts go through the same chart review as
     on the Model Desk.

A status panel at the top of every step shows where the engagement is. Each stage appears as done, running (with
what the agents are doing now, step by step, the review loop animated), needing you, or failed. It also shows the
next thing for you to do, as a button that takes you there. Every running job shows how long it has been going
and, where it reports progress, roughly how long is left. Other jobs estimate from how long the last run took,
and finished jobs show how long they took.

Any processed file can be rebuilt from scratch from the Files step. A workbook's build is shared: every
engagement using it, and the Model Desk, get the new one, and live Python overlays let go of the old file first.
A report is read again: every table is re-read and re-checked, and its key facts are kept and re-checked against
the new reading. Uploading the same workbook to another engagement reuses its build. The same report uploaded to
another engagement is read again.

All model calls run on Azure Foundry through `bench/llm.py`: gpt-6-luna extracts and reads tables, gpt-6-sol reviews
(second reads of picture tables, fact review), and both can be changed per engagement in the header. Calls are logged
to the usage table with session `engagement-<id>`. State lives in `out/engage.db`, report reads in `out/docs/` and
compiled overlays in `out/overlays/` (all git-ignored).

### Excel formulas as Python (bench/xlcompile.py, bench/xlruntime.py)
`xlcompile.py` turns a workbook's formulas into one Python function per line item. Cells in a row whose formulas differ
only by a column shift share one branch, with the Excel formula and the inputs it uses written beside it.
IF / IFERROR / IFNA / CHOOSE branches are lambdas, so only the branch taken is computed. Ranges are lazy, so
`INDEX(range, MATCH(...))` evaluates one cell. Defined names, whole-row and whole-column references and ISFORMULA are
resolved at compile time. Workbook text enters the code only through `repr()`.

`xlruntime.py` has Excel's values, operators and about 100 functions: lookups, conditional sums, MMULT, OFFSET,
dates, YEARFRAC, XNPV / XIRR, text. It also follows Excel's rules for blanks, errors, comparisons across types and
implicit intersection. Evaluation runs column by column on a thread with a large stack, because a timeline
recurrence can nest thousands of cells deep. Circular references fall back to the saved value and are listed.

On the reference workbook (20 sheets, about 422k formulas) the compiled module reproduces 422,332 of 422,335 formula
cells. The other 3 are TODAY() and NOW(). Compiling takes about 11 seconds and a full recalculation about 4. One Excel
quirk is copied on purpose: XIRR with a zero first cash flow returns 2.98E-09 instead of the rate, and the workbook
carries that value.

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
