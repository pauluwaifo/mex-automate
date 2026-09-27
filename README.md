# MEx Automate

An Excel task pane add-in that automates repetitive spreadsheet work: cleaning data, merging
sheets and files, building charts and summary tables, filling formula patterns, and refreshing
report templates.

Everything runs client-side inside the Excel add-in sandbox. There is no backend, no account, no
AI/API cost, and no data leaves the machine — including the .xlsx and .csv files you merge, which
are parsed in the task pane itself.

---

## Requirements

- Node.js 18+ and npm
- Excel 2019 or later on Windows/Mac, Excel in Microsoft 365, or Excel on the web
- The manifest declares **ExcelApi 1.4** as its minimum. `Range.copyFrom` (1.9) is used when the
  host supports it, with a pure-TypeScript fallback when it does not.

## Getting started

```bash
npm install
npm start          # builds, trusts the dev certificate, and sideloads into Excel
```

`npm start` opens Excel with the add-in loaded. Find it on the **Home** tab under
**MEx Automate → Automate**.

| Script | What it does |
| --- | --- |
| `npm start` | Sideload into Excel with a live dev server on https://localhost:3100 |
| `npm stop` | Stop debugging and remove the sideloaded add-in |
| `npm run dev-server` | Dev server only, without sideloading |
| `npm run build` | Production bundle into `dist/` |
| `npm test` | Jest unit tests |
| `npm run typecheck` | `tsc --noEmit` |
| `npm run lint` / `lint:fix` | ESLint (flat config) + Prettier |
| `npm run validate` | Validate `manifest.xml` |
| `npm run sample` | Write sample test data into `samples/` |

---

## How you use it

The pane opens on a **chat**. Type what you want, or tap a suggestion:

```
/review                              check this sheet for mistakes
/fix                                 repair a messy export onto a new sheet
/dashboard                           up to 8 charts and headline numbers
/chart sum of Revenue by Region as pie
/compare April Book with May Book    every difference between two sheets
/unpivot                             month columns become rows
/recipe save month end               remember these steps for next month
/insights                            what the numbers show, in plain English
/explain                             read the selected formula back in English
/watch                               re-check cells as you edit them
fix 2 and 3                          act on items from the list just shown
skip totals                          leave one kind of fix out
undo                                 put back the cells from the last fix
```

Plain phrasings work too: "check this sheet", "remove duplicates", "pie of
revenue by region", "what's wrong with this?". Type `/` for the full list.

MEx is a **command assistant, not an AI**: no model, no API key, no network
call, nothing sent anywhere. Commands are parsed against a fixed list, with
forgiving matching for sheet and column names (case, spacing, plurals, small
typos). When it can't tell what you meant it says so and suggests the closest
commands, rather than guessing at something destructive.

Anything large is previewed and waits for a yes. Small, targeted clean-ups run
straight away on the cells named in the reply.

The step-by-step screens are still there behind **All tools**, for the things
with many choices, and the conversation survives the trip.

## Features

### Check for mistakes (Review)
Goes over a sheet looking for the errors that quietly produce wrong numbers,
and **marks them in Excel itself**: red for errors, amber for things worth
checking, blue for tidying.

It finds:

- **A typed-in number among formulas** - the classic broken column, where one
  row was pasted over and stopped updating
- **A total that misses rows** - `=SUM(F2:F14)` sitting under data that runs to
  row 16, with the corrected range offered
- **Numbers and dates stored as text**, so they never add up or sort
- **An ID used twice**, naming the row it clashes with
- **The same name spelled two ways**, which splits every chart and total
- **Values far outside the rest** (by median absolute deviation, so the outlier
  can't hide the outlier)
- **Dates in the future or before 1990**, usually a mistyped year
- **More than one currency in a column**, and negative counts

Each one carries the exact cell, a plain-English explanation, and - where it's
safe - a one-click fix. Anything needing judgement is reported without a fix,
so nothing changes on a guess.

Two details worth knowing:

- **Marks are borrowed, not taken.** A cell's own fill colour is remembered
  before it's highlighted, so *Clear marks* puts back exactly what was there.
- **Undo is ours, not Excel's.** Ctrl+Z doesn't reliably cover what an add-in
  writes, so every fix snapshots the cells first and *Undo last fix* restores
  them exactly.

Ignored issues are remembered inside the workbook, so a known quirk stops being
reported.

The home screen leads with the two tools most people need: **Fix messy data** and **Build a
dashboard**. The rest sit under "More tools".

### Fix messy data
Finds the real table inside a messy export and repairs it in one pass, writing a clean, typed Excel
table to a new sheet. The original is never changed.

- **Finds the table**: skips report titles, "Generated by..." lines and footnotes; joins headings
  split over two rows (a merged "Q1" over Jan/Feb/Mar becomes "Q1 Jan"); names blank and duplicate
  headings.
- **Removes rows that aren't data**: blanks, headings repeated mid-table (page breaks), and subtotal
  or total rows ("Europe Total", "Grand Total", "Sub-total") - but never a customer called
  "Total Foods Ltd".
- **Turns group headings into a column**: "Region: Europe" above a block of rows becomes a Region
  column filled in for every row.
- **Reads numbers however they were typed**: `$1,204.50`, `1.130,00 €`, `(250)`, SAP's trailing
  minus `980-`, `NGN 45,000`, `12.5%`, `1 234 567`, and `1,2O0` with a letter O for a zero. US
  vs European decimals is decided per column from the values that give it away.
- **Reads dates in any common style**, deciding day-first vs month-first per column from the
  unambiguous ones (13/04 can only be day-first).
- **Unifies spellings** of the same label - case, spacing, punctuation, one-letter typos - onto the
  most common, best-capitalised one. Codes containing digits (`SO-5001`) are never touched.
- Empties `N/A` and `-` placeholders in number and date columns, and removes duplicate rows.

Before anything changes, every problem is listed in plain words with a count and examples. Each
optional fix can be switched off.

### Build a dashboard
Suggests up to eight charts from what the columns hold, plus up to four headline numbers, and lays
them out on a 12-column grid on a new sheet: a full-width trend first, then charts in pairs.

- The suggestions aim for a rounded picture: a trend over time, a comparison, a share, a two-way
  breakdown, a trend per group, a top 10, a second measure, and a relationship between two
  measures. Each comes with a one-line reason; untick any you don't want.
- Charts are drawn from a repaired copy of the data (the same engine as Fix messy data), so a
  messy export works directly. The sheet itself is never changed.
- Prices, rates, percentages and years are never summed. IDs are never charted.
- Colours follow the value: "Europe" is the same colour in every chart it appears in. The palette
  is a fixed categorical order checked for colour-blind separation, with "Other" always grey.
- The numbers behind every chart go on a companion "... data" sheet.
- Dashboards are saved in the workbook and refer to columns **by name**, so **Refresh** rebuilds
  them from next month's export even if its columns come in a different order.

### Recipes — do this month's work again next month
Every command that does real work is recorded as you go; there is no mode to turn on. `/recipe save
month end` stores the sequence in the workbook, and `/recipe run month end` replays it.

- What is stored is **the command, not the result**: `/fix` rather than "clean A1:F250". Replaying
  re-parses each command against the workbook that is open now, which is what makes a recipe useful
  on a sheet that has grown by four hundred rows.
- Replay goes through the same code path as typing would, so there is no second implementation of
  "clean up" that could drift. Steps that normally ask "shall I?" are auto-confirmed — agreeing to
  the recipe is the agreement.
- A step naming a sheet this workbook hasn't got is **skipped and reported**, never re-pointed at a
  sheet with a similar name. Last month's `Feb Orders` is not this month's `Mar Orders`.
- Recipes live in `workbook.settings`, so they travel inside the .xlsx.

### Compare two sheets
`/compare April Book with May Book` matches rows on a key column and reports every cell that differs,
with both values and the movement, plus rows added and rows removed kept separate.

- The key column is chosen by looking for a shared column that is filled in and near-unique on both
  sides. A column of plain numbers is not taken as a key unless its name says so (`Invoice No`),
  because a quantity column is usually distinct too.
- Numbers are compared with a tolerance (half a penny by default), so rounding noise is not reported
  as change. `1000` and `"1,000.00"` match; `12` and `"twelve"` do not.
- Repeated keys are reported rather than silently compared against the wrong row.
- With no usable key it compares whole rows and **says so** — that finds rows added and removed but
  not cells edited, and the report states that limitation instead of quietly finding less.
- The differences go on a new sheet; changed cells are marked in the newer sheet through Review's own
  mark register, so **Clear marks** lifts them and restores each cell's own fill.

### Unpivot a crosstab
`/unpivot` turns `Region | Product | Jan | Feb | Mar` into one row per period — the shape charts and
PivotTables need, without Power Query.

- Headings are read as periods: `Jan`, `January 2024`, `Jan-24`, `2024-01`, `Q1 2024`, `2023`. A
  heading naming both (`Jan Units`, `Jan Revenue`) produces one row per period with a column per
  measure.
- A year merged across a row above the months is handled: the two heading rows are joined, so the
  year travels with the month.
- When every heading knows its year, the period column is written as **real dates** with a
  `mmm yyyy` format, so a trend chart works immediately.
- It always writes a new sheet; the crosstab is left exactly as it was.
- Headings that mix named measures with bare periods are refused rather than guessed at.

### Insights and explaining
- **`/insights`** — plain-English takeaways drawn from the same aggregates the charts use, so a
  sentence can never disagree with the chart above it: the last move, the overall direction from
  comparing halves, best and worst period, who leads, how concentrated the total is, any run of
  consecutive rises. They also appear under every dashboard as *What this shows*. The `Other` bucket
  is never described as a leader, because it is not a category of the business.
- **`/explain`** — reads the selected formula back in English via a real parser (precedence,
  right-associative `^`, sheet-qualified and absolute references, `""` inside strings), not regexes.
  It also names what is fragile: a rate typed into the formula, a whole-column reference, `IF`
  nested three deep, division with no guard, `VLOOKUP` finding its answer by counting columns. An
  unreadable formula is reported as unreadable rather than described wrongly.
- **`/watch`** — while the pane is open, re-checks just the cells you edit, a moment after you stop
  typing, and says if something looks wrong. It never changes anything. Needs ExcelApi 1.7; where
  that is missing it says so.

### Quick clean-ups
- **Remove duplicate rows** — compare whole rows or only chosen key columns, optionally ignoring
  case and spacing. Either blanks the tail of the range or deletes whole worksheet rows.
- **Trim whitespace** — including the invisible characters that survive a web copy-paste
  (non-breaking space, zero-width space, BOM).
- **Standardize dates** — converts text that looks like a date into a real Excel date, then applies
  one number format. Handles ISO, `d/m/y` vs `m/d/y` (with an explicit day-first switch, and
  automatic disambiguation when one part is greater than 12), named months, and 2-digit years.
- **Standardize text case** — UPPER, lower, Proper, Sentence. Formula cells are never overwritten.

Commands run on your selection. Select a **single cell** and the command runs on the whole sheet
instead — the pane always shows the exact range it is about to touch.

### Merge
Combine two or more sheets, or one or more picked files (`.xlsx`, `.xlsm`, `.csv`, `.txt`, `.tsv`),
into a single table on a new sheet.

Columns are matched by *normalized* header, so `Order Date`, `order_date` and `ORDER DATE` all land
in one column. When sources genuinely disagree, add an alias:

```
Client Name = Customer
```

Preview the planned column layout before committing, optionally tag each row with its source, sort
the result by any column, and skip blank rows. Number formats (dates, currency, percentages) carry
across from each source.

The result lands on a **new sheet** as a real Excel table with a frozen header row and auto-fitted
columns, ready to filter, chart or feed into the Reports tab.

### Charts
Build column, bar, line, area, pie, doughnut and scatter charts from any sheet, table or selection.

Raw rows rarely chart well, so the Charts tab **summarizes first** by default: group by one column,
then Sum / Average / Count / Count unique / Min / Max one or more others. That turns 4,000 order
lines into "Sum of Amount by Region" and charts *that*. Turn the summary off if your range is
already a small table of totals.

- Sort categories largest-first or alphabetically, and cap them at a **top N** with everything else
  rolled into a single "Other" slice — the difference between a readable pie and forty slivers.
- Numbers stored as text (`"1,234.50"`, `"(50)"`) are read as numbers; anything genuinely
  non-numeric is skipped and reported rather than silently counted as zero.
- **Build summary table only** produces the grouped table without a chart.
- The summary table is always written to the sheet next to the chart. An Office.js chart must bind
  to a real range, and it means the numbers behind the picture stay visible and auditable.

Charts can go on a new sheet or beside the source data, with a title, legend position and optional
value labels. Aggregated columns inherit the source column's number format, so a currency column
still totals as currency.

> Combining sheets and charting are designed to work together: **Merge** several sheets into one
> table, then point **Charts** at that merged sheet.

### Formulas
- **Fill a formula across a range** — select the cell with your formula plus the cells to fill. The
  pane shows what the last cell will become before you commit.
- **Formula library** — 13 one-click patterns across four categories: running total/average,
  % of total, % change, total by category, XLOOKUP, VLOOKUP, INDEX/MATCH, rank, count occurrences,
  join columns, year-month key, days between dates. Each renders a live preview against your actual
  selection.

### Reports
Mark parts of a report as **data zones**, map an incoming data source onto them by header name, and
refresh in one click. Titles, totals, formulas, charts and formatting around the zones stay put.

- A zone is anchored by an Excel *name* (a table name or a defined name), not an address, so it
  survives the data growing or shrinking.
- **Calculated columns are never overwritten** — any zone column whose first body row holds a
  formula is left for Excel to recompute.
- Templates are saved with `context.workbook.settings`, which stores them **inside the .xlsx**. Mail
  the workbook to a colleague and their copy of the add-in sees the same template.
- After writing, the refresh recalculates the workbook and refreshes every PivotTable (both
  optional). Charts pick up changed values automatically; a chart will only follow the data as it
  *grows* if its zone is an Excel **table**, because a chart bound to a fixed `A1:D100` range keeps
  that range. Prefer a table zone for anything a chart or pivot reads.

---

## Project layout

```
src/
  taskpane/
    index.tsx                  React entry point; follows the Office theme
    components/
      App.tsx                  Home screen of tools, and the header with a back button
      ChatPanel.tsx            The conversation: messages, chips, / autocomplete
      ReviewPanel.tsx          ...and a screen per tool, for people who prefer buttons
      TidyPanel.tsx
      DashboardPanel.tsx
      CleanPanel.tsx
      MergePanel.tsx
      ChartPanel.tsx
      FormulaPanel.tsx
      ReportPanel.tsx
      ui.tsx                   Step, ActionCard, MoreOptions, ActionButton, ChoiceGrid, Tip
      useSelection.ts          Live selection tracking
    features/
      tidy.ts                  Table detection, repairs and findings (Fix messy data)
      dashboard.ts             Column profiling, chart suggestions, layout, refresh
      dataCleaning.ts          Pure transforms + Office.js drivers
      review.ts                Mistake detectors (pure) - formulas, values, keys
      reviewSheet.ts           Marking in the sheet, fixes, snapshot undo, ignore list
      reshape.ts               Reading periods out of headings; unpivot (pure)
      reshapeSheet.ts          Finds the crosstab, writes the tidy sheet
      compare.ts               Key matching and cell-by-cell diffing (pure)
      compareSheets.ts         Reads both sheets, writes the differences report
      recipes.ts               Recording, checking and describing recipes (pure)
      recipeStore.ts           Recipes saved inside the workbook
      insights.ts              Turning aggregates into sentences (pure)
      explain.ts               Formula parser and plain-English describer (pure)
      explainSheet.ts          Reads the selected cell for the explainer
      watch.ts                 Re-checks the cells you edit, while you edit them
      commands.ts              The assistant's command language (pure)
      assistant.ts             Runs a command and answers with structured messages
      merge.ts                 Header planning + merge drivers
      charts.ts                Aggregation + chart drivers
      formulaPatterns.ts       A1 translation + template library
      reportBuilder.ts         Zone mapping + refresh engine
    shared/
      excelHelpers.ts          Grid/address helpers and Office.js utilities
      templateStore.ts         Template persistence in workbook settings
      workbookReader.ts        .xlsx / .csv parsing
      types.ts                 Shared domain types
    theme.ts                   The website's jade palette and type, as a Fluent theme
tests/                         Jest unit tests (473)
manifest.xml                   Add-in manifest
```

Each feature module is split the same way: **pure functions over plain data** at the top (unit
tested), and thin **Office.js drivers** at the bottom that read a range, call the pure function, and
write back. Every driver returns the same `OperationResult` shape, which is what the pane renders.

## Testing

```bash
npm test
```

473 unit tests cover the pure logic, including the mistake detectors and the
command parser, table detection and every kind of repair on a
deliberately nasty export, chart suggestion, time bucketing, chart data and dashboard layout, and
regression tests for bugs caught on the sample sheet. Also covered: grid transforms, dedupe keying, date parsing and Excel serial
conversion, header planning, merging and sorting, grouped aggregation and chart validation, CSV
parsing, A1 formula translation, and column mapping.

Anything that calls Office.js is verified by sideloading (`npm start`) — mocking the Office.js proxy
object model produces tests that pass while the real thing fails, so it is deliberately not mocked.

> If the test runner crashes with a native V8 error, run `npx jest --runInBand`. That is memory
> pressure from parallel workers, not a test failure.

### Trying it in Excel

```bash
npm run sample     # writes samples/ with a deliberately messy test workbook
npm start          # sideloads the add-in and opens Excel
```

The first `npm start` installs and asks you to trust a certificate for
`https://localhost` — say yes, or the task pane loads blank. Then open
`samples/MEx-Automate-Test.xlsx` and click **Home → MEx Automate → Automate**.

The sample workbook is built to exercise every tab:

| Tab | What to try | What should happen |
| --- | --- | --- |
| Review | Open **Sales Check**, then type `/review` | Six planted mistakes found and marked in the sheet: a typed-in number among formulas, a SUM that stops at row 14, text Units, a repeated ID, "acme ltd", a 2029 date |
| Review | Say `fix 1`, then `undo` | The safe fixes apply, then the cells come back exactly as they were |
| Review | Say `show 2`, then `ignore 2` | Excel jumps to the cells; ignored issues stay gone on the next check |
| Fix messy data | Open **Sales Extract**, then Fix messy data | About 250 problems listed by kind; Fix writes a clean 78-row table to "Sales Extract (clean)" |
| Build a dashboard | Point it at **Sales Extract** (the messy one) | Eight suggestions with reasons; Build lays out headline numbers and eight charts on a new sheet |
| Build a dashboard | Change some Revenue values in Sales Extract, then Refresh | The same dashboard is rebuilt with the new numbers |
| Clean | On **Jan Orders**, click one cell, then Remove duplicates | Order 4 appears twice; one copy goes |
| Clean | Trim whitespace, then Proper Case | `"  Acme Ltd "`, `"ACME LTD"` and `"Acme Ltd"` become one spelling |
| Clean | Standardize dates (day-first **on**) | `2024-01-05`, `06/01/2024`, `Jan 7, 2024` and `10-Jan-2024` all become real dates in one format |
| Merge | Tick all three order sheets → Preview columns | `order_id`/`Order ID` collapse to one column; `Channel` is flagged as partial |
| Merge | Add alias `Client Name = Customer`, sort by Amount | Mar's client column folds into Customer; result is a sorted table on a new sheet |
| Merge | Switch to Files, pick both CSVs from `samples/` | The semicolon-delimited May file is detected automatically |
| Charts | Source = the merged sheet, group by Region, Sum of Amount, Pie | A summary table plus a pie of totals by region |
| Charts | Turn on Top N = 3 | Smaller regions collapse into one "Other" slice |
| Formulas | Select `G1:G9` on Jan Orders, put `=E2*2` in G1, Fill formula | Preview shows the last cell before you commit |
| Reports | On **Sales Report**, select `A3:E6`, name a zone `SalesData`, source = Jan Orders, Refresh | Rows are replaced; the title, the `Total` row and column E's formula all survive |
| Unpivot | Open **Budget by Month**, then `/unpivot` | Reads the year merged above the months; 10 rows x 6 months become 59 tidy rows (one gap skipped) with a real date column |
| Compare | `/compare April Book with May Book` | SO-2002 changed (units and amount), SO-2005 renamed, SO-2004 gone, SO-2008 new - and SO-2003's half-penny difference correctly ignored |
| Insights | On **Sales Extract**, `/insights` | Four sentences drawn from the same numbers the charts use |
| Explain | Click the `Total` cell on **Sales Check**, then `/explain` | "The total of F2:F14", and a note that the range stops short |
| Watch | `/watch`, then type text into a number column | The edited cells are re-checked within a second and reported |
| Recipes | Run `/fix`, then `/dashboard`, then `/recipe save month end` | Both steps are remembered; `/recipe run month end` replays them on a fresh copy |

`npm stop` unregisters the add-in when you are done.

Sample files are gitignored — regenerate them any time with `npm run sample`.

---

## Design notes

**Reading external workbooks without a library.** Office.js can only see the workbook the add-in is
running in, so merging *other* files means parsing them. `workbookReader.ts` unzips the .xlsx with
[fflate](https://github.com/101arrowz/fflate) and walks the XML with the browser's own `DOMParser`.
This avoids `xlsx` (deprecated on npm, with known prototype-pollution advisories) and `exceljs`
(needs Node polyfills under webpack 5), and keeps the added bundle weight to a few KB. It reads
values, shared strings and number formats — enough for merging data — not charts or formulas.

**Letting Excel shift the references.** The formula filler prefers `Range.copyFrom`, so reference
adjustment always matches what Excel itself would do. `translateFormula` is the fallback for hosts
below ExcelApi 1.9, and powers the preview. It correctly leaves string literals, `$`-anchors, quoted
sheet names and structured references alone. Its one documented limit: a defined name shaped like a
cell reference (`Q1`) is indistinguishable from a reference without Excel's name table.

**The 1900 leap-year bug.** Excel believes 1900 was a leap year and reserves serial 60 for a
29 February that never existed. `dateToExcelSerial`/`excelSerialToDate` compensate, so serial 1 is
1900-01-01 exactly as Excel shows it.

**Undo.** Operations are batched so Excel's native Ctrl+Z keeps working. A large refresh may take
several undo steps, because writes are split into batches to stay under Office.js payload limits.

---

## Before publishing to AppSource

`manifest.xml` still contains placeholders that must be replaced. AppSource will reject the
submission otherwise:

- [ ] `<ProviderName>` — must match your Partner Center publisher name
- [ ] `<SupportUrl>` — currently `https://www.example.com/mex-automate/support`; needs a real,
      reachable HTTPS support page
- [ ] `GetStarted.LearnMoreUrl` — currently `https://www.example.com/mex-automate/help`
- [ ] `assets/*.png` — still the icons from the Yeoman scaffold; replace with your own
- [ ] Host the built `dist/` on HTTPS and set `urlProd` in `webpack.config.js` to that origin
      (the production build rewrites `https://localhost:3100/` to it)
- [ ] Add a `LICENSE` and a privacy policy URL

`<Id>` is already a unique GUID for this add-in — keep it stable across versions.

## Licence

Not yet chosen — add a `LICENSE` file before publishing.

---

## Landing page and hosting

The product website is a single self-contained file, [docs/index.html](docs/index.html), with an
animated walkthrough of the add-in. It has no build step and no dependencies beyond Google Fonts.
To preview it, open that file in a browser.

### One deployment serves both

The site and the add-in are deployed together, so every URL in the manifest sits on the same origin
as the page that explains it:

```
https://mex-automate.vercel.app/                 the landing page
https://mex-automate.vercel.app/addin/...        the built add-in (taskpane.html, bundles, assets)
https://mex-automate.vercel.app/manifest.xml     the manifest people install
https://mex-automate.vercel.app/install.ps1      the Windows installer
```

```bash
npm run build:site     # webpack --mode production, then assemble public/
```

- [scripts/make-manifest.mjs](scripts/make-manifest.mjs) derives the hosted manifest from
  [manifest.xml](manifest.xml): it rewrites the `localhost:3100` URLs to `/addin/`, swaps in a
  **separate add-in id** so the localhost copy and the installed copy can both be registered at
  once, and replaces the placeholder support links. It refuses to emit a manifest that still
  contains a placeholder URL.
- [scripts/build-site.mjs](scripts/build-site.mjs) assembles `public/`, taking the production
  domain from Vercel's own `VERCEL_PROJECT_PRODUCTION_URL` so renaming the project cannot leave the
  manifest pointing at a stale host. Override with `MEX_SITE_BASE`.
- Pushing to `main` deploys: the GitHub repo is connected to the Vercel project. `vercel.json` sets
  the build command, serves `manifest.xml` as a download, and keeps `/addin/**` uncached so an
  update reaches people without a reinstall.

`.npmrc` sets `legacy-peer-deps=true`. `babel-jest` 29 asks for Babel 7 while this project builds
with Babel 8; without it a clean `npm install` fails, including on Vercel's build container.

### Installing the hosted add-in

Excel loads an add-in from a manifest, and outside AppSource that manifest has to be pointed at by
hand. [install/install.ps1](install/install.ps1) does it for Windows: it downloads the manifest to
`%LOCALAPPDATA%\MEx Automate` and writes one value under
`HKCU\Software\Microsoft\Office\16.0\WEF\Developer`. No admin rights, nothing machine-wide, and
[install/uninstall.ps1](install/uninstall.ps1) removes exactly those two things.

The installer refuses to register a manifest whose `<Id>` is not the expected one, so a wrong or
redirected URL cannot quietly install something else.

Excel on the web takes the manifest through **Home → Add-ins → More Settings → Upload My Add-in**;
Excel for Mac reads it from the `wef` folder. Both routes, and the Trusted Add-in Catalogs route for
people who would rather not run a script, are on the site's install section.

**Before launch:** the "Get it free" buttons scroll to the install steps until the AppSource
listing exists. When it does, set `STORE_URL` near the top of the page's `<script>` to the listing
URL; every button uses it.
