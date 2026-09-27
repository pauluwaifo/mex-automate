# What to build next, and why

Research notes behind the MEx Automate roadmap. Every claim here has a source; every
"we could build this" has been checked against what the Office JavaScript API can
actually do, using the type definitions in `@types/office-js` rather than hope.

Written September 2026.

---

## Part 1 — What everyday Excel users actually struggle with

### The numbers

| Finding | Source |
| --- | --- |
| ~94% of operational spreadsheets contain at least one error; average cell error rate ~5% | [Panko, via DataHub Pro](https://www.datahubpro.co.uk/spreadsheet-error-statistics) |
| 22% of operations professionals hit spreadsheet errors **every day**. Of those, 59% say manual data entry is the most common cause and 46% say formulas | [DOSS survey, 2026](https://www.inc.com/kit-eaton/the-4300-spreadsheet-mistake-everyones-making/91345013) |
| Teams spend 3.6 hours a week fixing spreadsheet errors — 22 working days per person per year | [Tier2 Systems](https://tier2systems.com/en/blog/spreadsheet-fix-time-cost/) |
| Finance staff spend 40–60% of their hours on data entry and validation | [Lido](https://www.lido.app/blog/eliminate-manual-data-entry-finance) |
| 94% of finance teams use Excel to manage the close; **50% name it as the main reason the close is slow** | [Ledge, month-end benchmarks](https://www.ledge.co/content/month-end-close-benchmarks-for-2025) |
| Cash reconciliation alone takes 20–50 hours a month for many finance teams | [Ledge](https://www.ledge.co/content/month-end-close-benchmarks-for-2025) |
| Only about **1 in 5 people use PivotTables**; roughly 95% of users touch ~5% of Excel's functions | [MrExcel / Acuity](https://www.acuitytraining.co.uk/news-tips/new-excel-facts-statistics/) |
| Unpivoting a crosstab is the standard blocker, and Power Query — the official answer — has "a steep learning curve for many Excel users" | [Excel University](https://www.excel-university.com/unpivot-excel-data/) |

### What that adds up to

Three things, and they are not the same problem:

1. **People make mistakes and don't find them.** Manual entry and formulas, daily, at a
   scale where the average sheet is wrong. This is a *detection* problem.
2. **People repeat themselves.** The close is slow not because the reporting is hard but
   because everything before it — reconciling, aligning, correcting — is done by hand
   every month. This is a *repetition* problem.
3. **The tools that would fix both exist and go unused.** PivotTables: 1 in 5. Power
   Query: the documented answer to reshaping that most people never learn. This is an
   *accessibility* problem, and it is the one that decides whether a feature lands.

The third point is the important one for us. Excel is not short of capability. It is
short of capability that an ordinary person can reach on a Tuesday afternoon without a
course. Every feature below is judged on that, not on whether it is technically clever.

### Where MEx already answers it

| Their problem | What MEx does now | Honest gap |
| --- | --- | --- |
| Formula errors (46%) | `/review` finds typed-in numbers among formulas, short SUM ranges, text numbers | Only within one sheet; no cross-sheet or whole-workbook audit |
| Manual entry errors (59%) | `/review` finds duplicate IDs, blank keys, near-duplicate spellings, outliers, impossible dates | Nothing *prevents* the next one — no data validation |
| Reconciliation (20–50 hrs/month) | `/compare` diffs two sheets with a tolerance | Two sheets only; not sheet-vs-file, not many-to-one |
| Repetition | Recipes record and replay | Only within one workbook; can't run last month's recipe on this month's *file* |
| Reshaping (Power Query too hard) | `/unpivot` | One transformation of several; no split-column, fill-down, group-and-total |
| Nobody uses PivotTables | Dashboards build charts instead | Fine — but a PivotTable is what people ask for by name |

---

## Part 2 — Ranked opportunities for Excel

Ranked by (evidence of pain) × (how many people it reaches) ÷ (effort), with feasibility
checked against the API.

### 1. Workbook health check — "why is this file slow and what's fragile in it"

**The pain.** Slow files are a constant complaint, and the causes are well documented and
*detectable*: whole-column references, volatile functions, conditional formatting applied
to thousands of cells, thousands of unused formulas, circular references, and used ranges
that stretch far beyond the real data ([PerfectXL](https://www.perfectxl.com/blog/excel-tips/excel-running-slow-10-common-causes-and-how-to-speed-it-up/),
[Microsoft's own performance guidance](https://learn.microsoft.com/en-us/office/vba/excel/concepts/excel-performance/excel-tips-for-optimizing-performance-obstructions)).
None of this is visible to the person suffering from it.

**Why us.** `explain.ts` already detects whole-column references, volatile functions,
unguarded division and positional VLOOKUP — on one formula. Running the same detectors
across every formula in the workbook, plus a few sheet-level ones, is mostly wiring.

**Feasible?** Yes. `Range.getSpecialCells` (ExcelApi 1.9) returns just the formula cells
without reading the whole sheet; `worksheet.getUsedRange()` versus the real data extent
finds bloat; `conditionalFormats` (1.6) can be counted. Gate on 1.9, fall back to reading
the used range.

**Shape.** `/health` — a report naming each problem, the sheet it's on, what it costs, and
a one-click fix where one exists (tighten a whole-column SUM to the data; delete the junk
beyond the used range). This is also the most *shareable* output we could produce: "here's
what's wrong with the model you inherited."

### 2. Stop the next mistake — data validation from what the column already holds

**The pain.** 59% of errors are manual entry. Detection after the fact is second best.

**Why us.** We already profile every column (`profileColumn`): we know a column is a
category with six distinct values, or a date, or a positive quantity. Turning that into a
rule is a short step from what we compute anyway.

**Feasible?** Yes — `Range.dataValidation` (ExcelApi 1.8): list rules, date ranges, number
bounds, plus an input message. Gate on 1.8.

**Shape.** `/protect` — "Region only accepts the six values already in it; Units must be a
positive number; Order Date must be this year." Shown as a list, applied on a yes, and
removable. This is the only feature here that changes the *future* of a sheet rather than
its past.

### 3. Reconcile a sheet against a file, and many sheets against one

**The pain.** 20–50 hours a month, and the single biggest line item in a slow close.

**Why us.** `/compare` exists and `workbookReader.ts` already parses `.xlsx` and `.csv` in
the browser. Comparing the open sheet against a file the user picks is the same engine
with a different left-hand side.

**Feasible?** Yes, entirely — no new API surface.

**Shape.** `/compare` gains: against a picked file; against several sheets at once
("which of these twelve tabs disagrees with the summary?"); and a tolerance the user can
set in words ("to the nearest pound").

### 4. The rest of Power Query's greatest hits

**The pain.** Unpivot was the top blocker and we shipped it. The same "I know what I want,
I can't get there" applies to: split one column into several, fill a value down a merged
or gappy column, group and total, and trim a table to a date range.

**Feasible?** Yes — pure grid transforms, the pattern the codebase is built around.

**Shape.** `/split Customer by comma`, `/fill down Region`, `/group by Region total
Revenue`. Each is small; together they're the reason someone stops opening Power Query.

### 5. A PivotTable, because that is what people call it

**The pain.** People ask for a pivot by name, and 4 out of 5 never make one.

**Feasible?** `pivotTables.add` plus row/column/data hierarchies is ExcelApi **1.8**, which
is a real availability cost. Our dashboards already produce the same answers, so this is
about meeting people in their own vocabulary rather than new capability. Gate on 1.8 and
fall back to the summary table we already build.

### Deliberately not doing

- **An AI/LLM layer.** It would need a backend and a per-call cost, and the product's
  promise is that nothing leaves the machine. The command assistant answers the same
  questions deterministically.
- **Competing on tool count.** [ASAP Utilities, Kutools and Tellsheet](https://tellsheet.com/best-excel-add-ins)
  ship 150–300 tools each. That race is lost and not worth entering: the whole point of
  MEx is that you don't have to find the tool.

---

## Part 3 — PowerPoint

### The case for it

| Finding | Source |
| --- | --- |
| Employees spend close to a full working day a week in PowerPoint; ~100 hours a year | [empower, The Big PowerPoint Study](https://www.empowersuite.com/hubfs/Marketing/Downloads/The-Big-PowerPoint-Study.pdf) |
| Over 40% of the time spent in PowerPoint goes on **formatting**, not content | [PresentationLoad](https://www.presentationload.com/blog/uncovered-biggest-powerpoint-time-wasters/) |
| 2 hours to build a deck from scratch; 55% of marketers say making slides is not a good use of their time | [24Slides](https://24slides.com/presentbetter/professional-powerpoint-design-cost) |
| Linked Excel charts break when the workbook is renamed, moved, or its ranges change — and users report links that simply don't update | [Microsoft Q&A](https://learn.microsoft.com/en-us/answers/questions/5025943/powerpoint-doesnt-update-linked-charts), [Rollstack](https://www.rollstack.com/articles/how-to-link-an-excel-chart-to-powerpoint) |
| Fonts not installed on the reader's machine silently reflow the whole deck | [SketchDeck](https://sketchdeck.com/blog/ppt-font-issues/) |

**The specific job to steal:** the monthly reporting deck built from an Excel file. Today
that is either copy-paste (an hour a month, wrong the moment the numbers change) or
paste-link (breaks when anything moves). Both are the same failure: the deck is coupled to
a *file path* instead of to a *question*.

MEx already knows how to answer the question. It reads a messy export, works out what the
columns hold, and decides which charts tell the story. Pointing that at a slide instead of
a sheet is a small step for us and a large one for the person doing it by hand.

### What the PowerPoint API can and cannot do

Checked against `@types/office-js` directly. This decides the design, so it is worth being
precise:

**Available, and enough to build with (PowerPointApi 1.3–1.5):**

| Capability | API | Set |
| --- | --- | --- |
| Add a slide on a chosen layout | `slides.add({layoutId, slideMasterId})` | 1.3 |
| Text boxes, shapes, lines | `shapes.addTextBox`, `addGeometricShape`, `addLine` | 1.4 |
| Position and size anything | `shape.left/top/width/height` | 1.4 |
| Fill and outline colour | `shape.fill.setSolidColor`, `lineFormat` | 1.4 |
| Fonts: name, size, bold, colour | `textFrame.textRange.font` | 1.4 |
| **Metadata on slides and shapes** | `slide.tags`, `shape.tags`, `presentation.tags` | 1.3 |
| Read the selection | `getSelectedSlides`, `getSelectedShapes` | 1.5 |
| Insert slides from another deck | `insertSlidesFromBase64` | 1.2 |

**Newer, so feature-detect rather than depend on:** native tables (`addTable`, 1.8),
images (`ShapeFill.setImage` 1.8, `Shape.setImage` 1.10), theme colours
(`themeColorScheme`, 1.10), slide masters detail (1.10).

**Not available at all: there is no chart object in the PowerPoint JavaScript API.** No
`addChart`, nothing. "Chart" appears only as shape names (`ChartX`, `ChartPlus`) and
layout names.

Microsoft's docs list **1.5 as the most recent stable set**, supported on the web and on
Windows M365 from version 2208 ([requirement sets](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/powerpoint/powerpoint-api-requirement-sets)).
So: **build on 1.4, gate everything above it.**

### Why the missing chart API is good news

The obvious workaround is to paste a picture of an Excel chart. That is the wrong answer:
a picture can't be restyled, can't be recoloured to the client's brand, goes fuzzy on a
projector, and is invisible to a screen reader.

The better answer is to **draw the chart out of native shapes** — rectangles for bars,
`addLine` for a trend, and the `Pie`, `PieWedge` and `Donut` geometric shapes for a share.
The result is ordinary PowerPoint vector shapes that anyone can nudge, recolour and
animate, in the deck's own theme. It is more work for us and a better artifact for them.
It also means the whole thing runs on 1.4.

### How it gets the data, with no backend

A PowerPoint add-in cannot read the user's Excel workbook — an add-in is scoped to its own
document. Two honest options:

1. **The user picks the .xlsx** in the pane, and we parse it in the browser.
   `workbookReader.ts` already does exactly this with `fflate`, for Merge. One click per
   refresh, no server, no file path to break.
2. Chart images handed over from the Excel add-in. Rejected: it depends on browser storage
   being shared between two different Office hosts, which is not something to promise.

Option 1 also removes the problem that breaks paste-link. The deck stores, in slide
**tags**, *what the slide shows* — "revenue by region, summed, from the sheet named Sales"
— not a path to a file. Rename the workbook, move it, add three hundred rows, reorder the
columns: **Refresh** asks for the file, re-reads it, recomputes, redraws. Nothing to
repair.

### What MEx for PowerPoint would be

Same pane, same conversation, same commands where they make sense:

- **`/deck`** — pick a workbook, and MEx builds the monthly deck: a title slide, a
  headline-numbers slide, a slide per chart with the insight sentence already written
  underneath (`insights.ts` produces those today), and a closing "what changed" slide when
  it can see last month's file too.
- **`/refresh`** — next month, one command. Re-reads, recomputes, redraws in place, keeping
  anything you added by hand. This is the feature that sells it.
- **`/tidy`** — the formatting time sink: find every font that isn't in the deck's theme,
  every off-brand colour, titles that overflow their placeholder, text below a readable
  size, and inconsistent margins. Report them, fix on a yes. (Reads `shape.textFrame`,
  `fill`, geometry — all 1.4.)
- **`/table`** — put a real table of the numbers on a slide (gated on 1.8, drawn as shapes
  below that).
- **`/explain`** — what this deck's numbers actually say, from the same engine as Excel.

Roughly 70% of the existing codebase carries over untouched: `workbookReader`, `tidy`,
`charts`, `dashboard` planning, `insights`, `commands`, `assistant`, `recipes`. What's new
is a renderer that draws to slides instead of a sheet, and a thin PowerPoint driver layer.

### Shipping shape

One repo, two manifests, one deployment. `manifest.xml` stays as it is; a second
`manifest-powerpoint.xml` points at `/addin/taskpane.html` with `<Host Name="Presentation"/>`.
The build already produces one site; this adds one file to it and one line to the
installer.

---

## What to build first

If the goal is the largest number of people helped per week of work:

1. **Workbook health check** (Excel) — biggest reach, mostly reuses `explain.ts`.
2. **`/deck` and `/refresh`** (PowerPoint) — the monthly-deck job, and the thing nobody
   else does without a server.
3. **Data validation** (Excel) — the only feature that prevents tomorrow's errors.
4. Compare-against-a-file, then the remaining reshapes.
