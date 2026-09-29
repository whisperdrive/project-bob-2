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
   labelled "Valuation date", sheet-header text), then one call to the selected model (the Valuation
   Desk's engagement model, or the Model Desk's chat model, when the file was uploaded) picks the target, code name and
   valuation date and cites the cell. Rows labelled exactly "Valuation date" come first, before rows that only
   mention one (a model that rolls back to last year's date has many, such as "Roll forward valuation date (to
   …)"), and a date is never read from a label's text. Shown as "Check" until the user confirms or corrects it.
4. **Compared** (`bench/diff.py`) with the latest earlier file for the same target (or, failing that, the
   same file name without dates/v2/final). Line items are aligned per sheet like a text diff, so inserted,
   removed and renamed rows don't shift everything else. When the timeline has moved (next year's model), each
   period is compared with the same period, matched by the timeline row's dates, instead of the same column.
   Reports inputs changed (with timeline period), key outputs that moved, formula changes, hard-code/formula
   swaps, sheets and named ranges, plus
   warnings when a file looks un-recalculated (inputs changed but no results moved, the valuation-date input
   disagrees with the calculated date, or formula results are blank). The same model writes a short summary.

The chat (Ask tab) gets the confirmed target, valuation date and change summary as context. Tool calls are
chosen by the model on Azure but run locally (`agent._run_tool` against model.db); only their text results
are sent back. The `chart` tool takes cell ranges, reads the exact values locally and the page draws them
with Chart.js (hover, click-to-zoom, copy data); the model only sees a first/last/min/max summary. Charts can be
lines, clustered or stacked columns, stacked areas, a combo (stacked columns for the parts, with lines for a
total or last year's figure), or a waterfall (a bridge, e.g. enterprise value to equity value, from a column of
line items or single cells). A rate can sit on its own axis on the right. Both apps draw them with the same
script (`web/charts.js`), and the viewer can switch between the kinds that fit.
Before a chart is shown it's reviewed on the server: `bench/chartrender.py` draws it to PNG with matplotlib (no
browser), a vision model (`bench/chartreview.py`, gpt-4o) checks the image against an exact data summary and
suggests presentation fixes (title, chart type, visible window, y-axis range, a note, e.g. for an outlier that
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
overlay (the workings that take the client model to the report's conclusions) and this year's client model.

The page is arranged by the valuer's questions rather than the pipeline. **Overview**: last year's value (from
the report, marked when the rebuild ties to it), this year's rolled forward, the change and the bridge between
them, how sure we are (report read and checked, conclusions tie, Python equals Excel, the report's charts
match), and one **needs-you** list of everything waiting for a person: files to retry or upload, roles to
confirm, report facts and tables the agents couldn't settle, builds to start, figures that don't tie, charts
that don't match. The sidebar counts them by page. **Doctor**: why a figure comes out wrong, and what to do
(below). **Setup**: Files and Roles. **Last year**: The report (its
tables and key facts), Rebuild in Python (step 6 below) and the Map. **This year**: Summary (step 7), the
**Value bridge** and What changed (the client models compared). Ask is a panel beside every page, the
Azure models sit behind the Models button, and a switch in the header's top-right corner picks the light or dark
theme (remembered in the browser; the system's until one is picked). The accent is blue (#1a9afa for lines and
borders; buttons and done circles are a deeper #1478d0 with white on them, 4.5:1), on the off-black and grey palette. A disclaimer at the foot of every page says the desk is experimental,
that its output must be checked by a qualified person and isn't advice, and gives the copyright and trademark
notes.

**Your firm's logo** goes in the header's top-left corner. Upload it from the page (hover over the logo, click
the pencil, then Upload a logo) or put the file in the `brand/` folder at the project root yourself and reload:
- **file:** `brand/logo.svg` (best: sharp at any size) or `brand/logo.png` (also accepted: `logo.webp`,
  `logo.jpg`); one logo at a time, 2 MB at most;
- **size:** 144 px high (it's shown 36 px high, so 4x stays sharp on high-resolution screens), up to 640 px wide
  (it's shown up to 160 px wide);
- **look:** a transparent background and light artwork: the header is off-black in both themes.
`brand/` is git-ignored, so the logo stays on the machine it was added on. Without one, the header shows the
default mascot (`engage/mascot.svg`: an original animated pixel critter in Claude's colours, still when the
system asks for reduced motion); "Use the mascot" on the same panel goes back to it. The pipeline behind the pages is below; each step is checked by a
person before the next relies on it:
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
   re-checked against the page, and so are the facts. A fact the reviewer accepted and only a check held up is
   approved by the agents once the check passes. Facts handed to you that rest on that table's page go back to the
   fact review loop with the new text (if a loop is already running, once it finishes). Facts on other pages stay with you, since the same loop on the same text
   would come out the same.
2. **Report reference.** `bench/reportfacts.py` extracts the target, valuation date, conclusions (preferred value and
   range), assumptions (discount rate and basis, terminal growth or exit / RAB multiple, ...), approach and the
   sensitivity grid, each with a page and a verbatim quote. Code checks each one: the quote is on that page, the
   values are in the quote, and quotes from unsettled tables are marked. Each fact's number is worked out by
   code from its printed text, never taken from a model: none for identity text (a name holding digits is still a
   name) or for a range the report gives without a preferred point. A reviewer model accepts, corrects or
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
   a third look (gpt-6-sol by default, set in the header; it gets its own brief and the evidence afresh). It can waive a check that failed on a clerical point, with
   a note kept on the fact and shown beside the check. It can also take the reviewer's correction or keep the
   extractor's version, but only if the checks then pass. Otherwise it hands the fact to the person with its
   note. When the checks improve, existing facts are checked again once (no model calls), and facts that were held
   up only by a check settle.

   **Tables** go through the same loop, with three more rules. First, the text-layer check forgives spacing and
   nothing else. Letter-spaced labels ("Hi gh"), numbers split at the decimal point ("3. 75%") and numbers run
   together across columns ("222" for 2 | 2 | 2) match when the characters are the same and in the same order on
   one line. A misread digit, a dropped % or bracket, or a swapped column still fails. Second, from round two, if
   the text layer still objects but the extractor's correction and the reviewer's independent read of the image
   agree on every number, the image wins. The figures in question are then looked for elsewhere in the report:
   the running text (the table's own page first) and the other tables. The reviewer judges whether each mention
   is about the same item (same measure, column, date and entity). A mention that contradicts the reads stops
   the table from settling; one that confirms them is shown with the table. Third, what still isn't settled goes
   to the arbiter. It looks at the image, both reads, the text layer and those mentions, and gives each reader
   feedback. The extractor corrects once more and the reviewer reads the image again. If the two now agree, the
   table is settled. If not, it goes to the person, with the agents' latest correction on screen. When the loop
   improves, the tables it left for a person go round once more, and any the new check now passes settle
   without a model call.

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
     `VALUATION_DESK_OVERLAY_MARKERS` in `.env`, so it stays out of the repo. That look counts for less than
     the report's figures: its points are added only when no other workbook has the report's figures on its
     overlay sheets. An engagement has one overlay, so any other workbook that looks standalone only by its
     structure (a client model rebuilt from the ground up, with valuation words and charts of its own) is typed
     a client model.
   Finding every DCF in a large workbook takes minutes (`valuation.catalogue`): it's saved beside the
   workbook's model.db, for that model.db and that version of the code, so a restarted server reads it back.

   Which version is earlier is decided by the identified valuation date, then the date in the file name
   ("20250523 …", "Jun 25", "FY26", "BP25"), then the timeline's first period. When the overlay sits in a copy of a
   client model that is also uploaded as its own file (v2.1 = v2.0 plus an overlay sheet), the client's own file
   is the prior client model, whatever its date: an adviser builds on the model the client sent and values at a
   later date, so the dates are given as reasons, not required to match. The overlay's copy of the client sheets
   is then fed from it, so feeding the prior model checks the copy matches the file. Each suggestion lists plain checks (✓ / ✗ / ?) and gets a second
   opinion from the reviewer model on the same evidence. The model can disagree, with reasons, and its assignment
   can be taken with one click; nothing is assigned until the person confirms. Only one suggestion runs at a
   time, and a workbook's external-link tables are built while it's processed, so suggestions don't write to a
   model.db that others are reading.

   The report's facts then add the stronger evidence: an overlay sheet holds the report's conclusions or
   valuation-only assumptions (label and value must both agree), or a DCF that `valuation.py` reproduces, or reads
   such a sheet. A sheet that another version of the model also has is the client's own, not the overlay, even
   when a report figure matches a cell on it (a cost of equity on the client's assumptions sheet) or its DCF is
   reproduced; and following what reads the valuation stops at such sheets and at 60% of the workbook's
   formulas (past that, a hub the whole model reads was taken for the valuation). When the reviewer model
   disagrees with confidence, the status panel and the needs-you list say so. The prior
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
   dates on and sets the valuation-date lever to the new date. The valuation date moves from last year's (the
   report's) to this year's client model's, not by the gap between the two client models (the overlay can sit
   on a copy of a model of another date); only without last year's, by the gap between the client models;
   failing those, by how far the client sheets' timelines moved; failing that, 12 months, flagged for a check.
   Where this year's model is dated on or before last year's valuation date (likely the model's own date, not
   this year's valuation date), nothing is rolled: every figure waits, and the Summary asks for this year's
   valuation date, which is set for the engagement (a client model's own date is the model's; a date set there
   comes before it). The timelines count only as a last resort: the most common forward move across at least
   five sheets, each against the sheet it is this year, the smaller on a tie, at most two years. The new date is
   always last year's moved by the months, so the plan and the feed agree. The dates
   are read from the files each time (a date corrected in Files moves the roll without a rebuild) and can be
   checked on the Files page; the Summary shows them and which are checked. Periods move separately from the
   valuation date, sheet by sheet, by the whole periods that ended in between: from September to December an
   annual sheet's FY2026 is still FY2026 while a quarterly sheet moves one quarter (moving an annual sheet by
   three months would land on dates it doesn't have). Client values come from the same line item in the period,
   found however this year's model changed (`bench/rowfind.py`). Each of last year's rows is looked for several
   ways, and the one the evidence supports best is taken:
   - **label**: the same label on the same sheet (its n-th occurrence, else the nearest), provided the sheet of
     that name shares at least 35% of its labels (a rebuilt model can reuse a name for something else); on the
     sheet a renamed sheet became (one sharing half its labels); else a label found once anywhere in the model;
   - **history**: the same values in the periods both models have as history, on any sheet (actual years don't
     change between versions), with later years close. A row blank in periods its sheet has dates for, where
     last year's has values, counts for less: an actuals sheet has the history exactly and nothing after it, and
     would otherwise win the row;
   - **words**: a label sharing most of its words, the values close, on the corresponding sheet or anywhere;
   - **neighbours**: the same rows around it in the dependency graph (it reads rows labelled as its inputs were,
     and is read by rows labelled as its users were);
   - **banner**: a summary cell in a sheet's first ten rows that read the row last year, found again by its label.
   The same label in place with last year's history comes first; a label whose history contradicts it gives way
   to the row that has last year's numbers. A candidate of another kind counts for less: last year's row
   calculated and this one typed values (from the rows table's formula and value counts), as a reconciliation
   sheet of pasted copies ("LINKED EBITDA") is, with last year's history exactly and a full series. A
   person's pick wins over all of them: a row, the row found (kept), or last year's values kept on purpose. Where nothing is found, or the row
   found is blank in a period where last year's had a value, last year's value stands in (its forecast for the
   same period, else the cell it read), never a blank (which would read as zero), and the Summary says how many
   values stood in, on which line items. Each of this year's figures on the Summary is held back (it would be
   last year's numbers under this year's name) while a row its own discountings' cash flows come from isn't
   found or mostly stands in, or, with no discounting recognised under it, while under half of the client
   values the figures read are found; all are while a row the figures read is found but blank this year in more
   than a fifth of its periods, or found with a confidence under 0.5 (or not at all) and not yet checked by a
   person (the same label in the same place counts as found well, and so does a row with no label at its own row
   number on a sheet laid out as before, where most labelled rows are where they were). A row with nothing to
   find (no label, no numbers and no formulas last year: a spacer inside a range a formula reads) is settled by
   code, its blanks standing in, and no model is asked about it. Rows of period flags and dates found no
   better are worked out from the rolled period dates where one relation held for every period last year
   (the period's date, end or start, its year, 1 after the valuation date or up to it, a constant), and
   count like any other row where none did. Last, the **zero-roll check**: this year's model at last year's
   valuation date, rolled by nothing, should give about last year's figure (0.75 to 1.33 times it:
   forecasts are revised, not replaced). A figure outside that is held whatever the other tests say (the rows
   found don't carry last year's numbers), and every figure shows its value at last year's date.

   **The row agents** (`bench/rowagent.py`) settle the rows the gate waits on without a person stopping the
   work. They run by themselves after every build in Python (and when this year's valuation date is set), as a
   background job the Summary follows. First by the numbers, with no model: for each open row, the rows
   `rowfind` has for it and rows whose values are close to last year's in several periods (which finds a row
   with no label at all). A candidate that carries last year's numbers (over the periods both have, a median
   difference within 15%, the same kind of row, not blank where last year's has values) is taken as the
   agents' pick, with the numbers as the reason. The agents' picks count as settled and show as theirs,
   including last year's values kept for a row that isn't a DCF cash-flow or timing row (last year's timing flags
   would discount at last year's dates); a person's pick always wins over them. Their picks carry the row
   finder's version: when finding changes, the agents' older picks are set aside and they look at those rows
   again, a person's picks kept. Their model calls are tagged with the engagement and the step, like every job's. The agents never keep last year's values for a row the DCF's cash
   flows come from: that figure stays held, saying they couldn't find it.
   Then, for the rows the numbers didn't settle, the models: the engagement's model (luna) gets a dossier on
   last year's row that doesn't lean on its label (sheet, section, the labelled rows around it, a formula, the
   kind of row, its values by period, the rows it reads and that read it) and the candidates with their
   numbers checked. It can search this year's model by words and inspect rows before it proposes one, or says
   the model has no such line item (last year's values are then kept, never for a DCF cash-flow row). The
   reviewer (sol) checks each proposal with the numbers and accepts or rejects it with a reason; a rejection
   goes back to luna. At most six actions and two proposals a row, three rows at a time; each decision, with
   luna's reason and sol's verdict, is on the Summary. Last, the figures again: where one still fails the
   zero-roll check, the rows the figures read are ranked by how far their numbers are from last year's (a DCF
   cash-flow row first), sol, as advisor, sees the figures, those rows and the decisions so far and names the
   rows to look at again and what to look for, and luna looks again with that; at most three rounds. A figure
   still off after that stays held, with the agents' account of what they tried rather than a list of chores;
   a person's pick on the Summary overrides any of theirs. From a wrong earlier pick on the synthetic pack, one
   round (an advice, a proposal, a review) put the right row back. Run on the synthetic pack, the renamed and moved
   distributions row took one proposal and one review: about 2,300 tokens. Rows of period flags and dates (by their labels, 0/1 values or rising dates)
   are the timing the discounting depends on: listed apart. The page says why, how alike the two
   models are, and lists each row it needs with a button to pick this year's row; the bridge leaves out the
   figures held back. The Map uses the same finding. Each output shows its Python function and the client values it
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
     The tab starts from the report: `bench/dcftrace.py` takes the cell a report conclusion was matched to in
     the Map, preferring one that ties at the report's printed precision and never one that differs. It follows
     every cell that cell's formula reads, down to the discounting: a `SUMPRODUCT` of cash flows and a factor
     row, or of cash flows and factors computed in the formula; an `XNPV`; an `NPV`; or the `SUM` of a
     present-value row. The tree shows each formula in line-item words, e.g. equity value (ex-div) = cum-div
     less the distribution payable; cum-div = the mid value; mid = the average of the PVs at the low and the
     high rate. Every cell is checked against the Python overlay. Each discounting is recomputed independently,
     with its rate, valuation date and convention read back from its factors, and its cash-flow row is broken
     into the rows it adds up. Those discountings join the anchors above, so Validate and the methods work for
     overlays whose DCF isn't a labelled `SUMPRODUCT`. The chat has the same trace as `overlay_value`.
     `tests/make_trace_workbook.py` writes a synthetic overlay with each of these shapes.
     **The facts** (`bench/dcffacts.py`), at the top of the tab, find the same thing from what the Python overlay
     actually reads (`Book.reads`), so lookups (INDEX, OFFSET, CHOOSE), names and the branch an IF took are
     followed as Excel took them, across sheets. A discounting is recognised by its numbers: a cell that reads two
     rows and equals the sum of their products, one of them factors between 0 and 1; or one that reads a row and
     equals its sum, each cell of it a cash flow (in any row, on any sheet, the client model's too) times a
     factor; or an NPV, XNPV or SUMPRODUCT with inline factors, checked by value. Each is recomputed from its cash
     flows and factors and must tie exactly. The card shows the cash-flow row and its periods, what the factors
     are computed from (the rate, the valuation date, the period dates), a Gordon terminal value, and the path up
     to the report's figure. It then lists the client-model rows the cash flows come from, and what this year's
     model has for each, with the evidence and a picker to choose another row. **Copy the facts** gives it all as
     text. `tests/rollforward_pack.py` writes an overlay inside a copy of a client model (an INDEX-picked scenario,
     the mid of a low and a high rate, present-value rows reading the client sheet) and a changed version of the
     model (rows inserted, renamed, restructured, a sheet renamed); `tests/check_rollforward.py` checks the facts,
     that every row is found, and that the figure rolled forward equals a calculation made from the numbers.
   - **Ask**: the Model Desk's chat agent (`bench/overlay_chat.py` on `bench/agent.py`). It has the workbook tools
     on any of the three workbooks, and tools that run the overlay: `overlay_run`, `overlay_chart` (a row saved
     vs recomputed, several feeds on one chart, and optionally the rows that add up to it as stacked columns
     underneath), `overlay_dcf`, `overlay_value` (the trace above), `overlay_formula` and `overlay_inputs`. When the prior client model, recomputed
     in Python, gives the values Excel saved in every period, the chart draws them as one line; where they
     differ it keeps both and says where (the overlay's saved link values are out of date there, or the Python
     differs). It knows
     the report's key facts, the levers, the outputs and the feeds. Charts go through the same chart review as
     on the Model Desk. **Dictation** (`bench/dictation.py`): the microphone beside Ask streams speech to an Azure
     Speech resource (`SPEECH_KEY` and `SPEECH_ENDPOINT` or `SPEECH_REGION` in `.env`; the Free F0 tier gives 5
     audio hours a month, one stream at a time; the page streams to the region the key belongs to). Words appear in the box as they're heard, to read and correct before asking.
     Microsoft's Speech SDK is loaded in the page on first use, pinned and checksummed. The server keeps the key.
     It hands the page a short-lived token and the engagement's own words as a phrase list: the names in the
     report's facts, their labels, the outputs, levers and sheet names, then common valuation terms. Dictation
     stops on the microphone or Esc, and after 8 seconds of silence (the free tier counts silence too). Enter
     stops it and asks. The line under the box shows the minutes dictated from this desk this month.
7. **Summary** (`overlay.summary_table`, `bench/reportcharts.py`). Last year's report, rebuilt and rolled forward.
   - **Valuation summary.** The rows are the report's conclusions (with their low and high ends), its key
     assumptions and its approach. The columns are the report; the rebuild on last year's client model (it
     should tie, and each figure says whether it does at the report's printed precision); this year, rolled
     forward onto this year's client model; and your scenario. In the scenario you can change any assumption
     that has an input in the overlay, the valuation date, and the discounting method (end or mid-period,
     actual/actual or actual/365). The method reaches a figure through its trace: each discounting under it is
     redone, and the formulas above carry the results up (e.g. the mid value as the average of the low- and
     high-rate PVs). That path is first checked by reproducing the module's own value with the method
     unchanged; where it can't be, the method isn't applied and the page says so. A range end typed into the
     overlay, rather than calculated, is shown for last year only.
   - **The report's charts.** Last year's report charted last year's client model, so the same rows drawn from
     that model must reproduce each chart. Charts are found in three places: pictures the table reader classed
     as figures, PPTX charts (whose data is exact), and charts drawn in the PDF itself (bars sharing a baseline
     or plotted lines, found from the page's drawing). A drawn chart is cropped with its axis labels and legend
     but no further: not across an empty strip (two panels side by side are two charts; a commentary column
     beside one isn't part of it), and not past the rule over a panel. It takes its title from that panel. A
     figure the table reader cut out of a drawn chart (often just the plot) gives way to the drawn chart's own
     crop. gpt-6-luna reads each one (title, kind, units, years, each series' values). A chart whose x axis
     isn't a run of years (valuation ranges like FY26-FY30, peers, a strip of an axis on its own) isn't a chart
     over years; it's listed, not recreated. Code finds the model rows whose financial-year totals follow each
     series, allowing for units and sign, within 6%. A series named for a site or segment can be the sum of
     that site's rows (its section, or rows carrying its name). A row that shares no word with the series must
     follow it within 2% over six years or more, a multiple or a percentage isn't rescaled by thousands, and a
     series too small to read off the picture isn't matched. The recreation is drawn from those financial-year
     totals, so rows from a quarterly and an annual sheet chart together. gpt-6-sol then compares the report's
     picture with the recreation (the data, not the styling); it isn't asked when most series have no row. A
     series it rejects gets the next candidate row, up to three tries. The same rows in this year's client model
     give this year's chart. **Compare** puts the report's reading beside the model's numbers, series by series
     and year by year. You can pick the rows for a series by hand (searched by label or section, added up on one
     sheet); the chart is then redrawn from them and checked by numbers. They run by themselves the first time
     the page opens after the overlay is built. `tests/make_engagement_pack.py` puts both kinds of chart in the
     synthetic report; `tests/check_charts.py` checks the finding, the time axis, the matching and the gates on
     a report laid out in panels.

**The value bridge** (`overlay.value_bridge`) goes from last year's value to this year's for each of the report's
conclusions in the overlay, one change at a time, so the steps add up exactly: last year's value (the rebuild);
a year of time value (each discounting under the figure grows at its own rate from last year's valuation date
to this year's, carried up the figure's formulas); last year's forecast cash flows up to the new valuation date
(each discounting redone at the new date); this year's client model (the engine's roll-forward, less the step
before); and your changes from the Summary page's scenario. Where the formulas above a figure's discountings
can't be recomputed, the first three steps are shown as one roll-forward step, and the page says why.

**The doctor** (`bench/doctor.py`) answers "why is this figure wrong?" from evidence, then gpt-6-sol writes it up,
held to that evidence (each finding says what was seen, the cause, the fix, and whether the fix is the app's,
the valuer's or a code change). It asks four questions:
- **Where does it break?** Each of the report's figures through the layers on its way to this year's value: the
  report, the overlay as Excel saved it, Python on the overlay's saved values, Python on last year's client model,
  Python on this year's. The first layer that disagrees with the one before is where to look. Python differing
  from Excel on the saved values can't be the wrong file or the wrong row, since no client file is read there.
- **Why?** From a wrong figure, the cells it reads are followed while they're wrong too, down to where it starts.
  The runtime records what a formula reads by running it again (`Book.reads`), so names, INDEX, OFFSET and IF
  branches are followed as Excel took them. Each starting cell gets its cause: a formula that doesn't compile, a
  function Python doesn't have (a data provider's add-in, or a macro function in an .xlsm), a name it can't find
  (an Excel table reference, a LET parameter, a name on another sheet), a client value that holds an error or
  differs from what the overlay last read, or the same inputs giving a different answer. On the pages these
  all show as #NAME?.
- **The right file?** The overlay's external links against the file assigned, and the client values the figures
  read against the values the overlay last saw (most differing means another version of the client model). An
  overlay link whose name and values both differ from the assigned file is still matched by its sheet names; if
  nothing matches, the doctor says nothing reads the client model.
- **The right rows?** Each client line item the figures read, followed into this year's model: its history (the
  periods up to last year's valuation date) should be the same numbers in both models, and without history the
  same forecast periods shouldn't be wildly different, flip sign or change units.

A cell Python can't compute whose value can't change between years (it reads nothing from the client model, no
assumption and not the timeline) can be held at Excel's saved value on every feed, from the page. Holding is
checked first by recomputing the figures on every feed, and a hold drops if the workbook is rebuilt and that cell
changed. **Copy for a message** puts the diagnosis and its evidence on the clipboard as plain text.
`tests/check_doctor.py` plants each kind of fault in the synthetic pack and checks the doctor finds it.

A status panel at the top of every page shows where the engagement is. Each stage appears as done, running (with
what the agents are doing now, step by step, the review loop animated), needing you, or failed. It also shows the
next thing for you to do, as a button that takes you there. Every running job shows how long it has been going
and, where it reports progress, roughly how long is left. Other jobs estimate from how long the last run took,
and finished jobs show how long they took.

Background jobs run in two lanes, each one job at a time: the report's (reading it, the key facts, both review
loops) and the models' (roles, compare, map, Python overlay, charts, doctor, row agents). A review loop never
holds up the map. The models' jobs stay one at a time because they share `model.db` and the Python session. Model
calls from both lanes share the same rate limit. A queued job is shown as in line, not running, with the job
ahead of it and when that job started (for example "Map: in line behind the row agents (started 14:02)"). The
row agents show under Rebuild in Python while they work.

Any processed file can be rebuilt from scratch from the Files step. A workbook's build is shared: every
engagement using it, and the Model Desk, get the new one, and live Python overlays let go of the old file first.
A report is read again: every table is re-read and re-checked, and its key facts are kept and re-checked against
the new reading. Uploading the same workbook to another engagement reuses its build. The same report uploaded to
another engagement is read again.

**The call log** (`bench/calllog.py`). Every model call is kept in `out/calls.db`, which is git-ignored because it
holds report text. Each row has what was sent (instructions, input, tools offered, the answer format; images are
noted but not kept), what came back, the model, tokens, cost at list price, how long it took, and what it was
for: the engagement, the report or workbook, the step and the purpose. Jobs tag their calls, including those
made on background threads, so the spend chip in the header opens a breakdown. It shows totals, tokens and model
time by file and by step, clock time by job, and the latest calls; click a call to read its transcript.

All model calls run on Azure Foundry through `bench/llm.py`: gpt-6-luna extracts and reads tables, gpt-6-sol reviews
(second reads of picture tables, fact review), and both can be changed per engagement in the header. Calls are logged
to the usage table with session `engagement-<id>`. State lives in `out/engage.db`, report reads in `out/docs/` and
compiled overlays in `out/overlays/` (all git-ignored).

### Excel formulas as Python (bench/xlcompile.py, bench/xlruntime.py)
`xlcompile.py` turns a workbook's formulas into one Python function per line item. Cells in a row whose formulas differ
only by a column shift share one branch, with the Excel formula and the inputs it uses written beside it.
IF / IFERROR / IFNA / CHOOSE branches are lambdas, so only the branch taken is computed. Ranges are lazy, so
`INDEX(range, MATCH(...))` evaluates one cell. Defined names (each sheet's own before the workbook's, read from
the file's workbook.xml, as openpyxl's read-only mode drops sheet-level ones), whole-row and whole-column references
and ISFORMULA are resolved at compile time. Array formulas (dynamic arrays and Ctrl+Shift+Enter ones) are kept: over
one cell as the formula itself, over a range as `INDEX(formula, i, j)` in each cell. IFS and SWITCH compile to lazy IF
chains. Names and functions it can't resolve are listed on the Rebuild in Python page. Workbook text enters the code
only through `repr()`.

`xlruntime.py` has Excel's values, operators and about 130 functions: lookups (XLOOKUP and XMATCH too), conditional
sums, MMULT, OFFSET, dates, YEARFRAC, XNPV / XIRR, PMT / PV / FV, LARGE / SMALL / MEDIAN / RANK / STDEV, text. It also follows Excel's rules for blanks, errors, comparisons across types and
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
