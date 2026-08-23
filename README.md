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
| `npm start` | Sideload into Excel with a live dev server on https://localhost:3000 |
| `npm stop` | Stop debugging and remove the sideloaded add-in |
| `npm run dev-server` | Dev server only, without sideloading |
| `npm run build` | Production bundle into `dist/` |
| `npm test` | Jest unit tests |
| `npm run typecheck` | `tsc --noEmit` |
| `npm run lint` / `lint:fix` | ESLint (flat config) + Prettier |
| `npm run validate` | Validate `manifest.xml` |

---

## Features

### Clean
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
      App.tsx                  Tab shell (Clean / Merge / Charts / Formulas / Reports)
      CleanPanel.tsx           ...one panel per feature
      MergePanel.tsx
      ChartPanel.tsx
      FormulaPanel.tsx
      ReportPanel.tsx
      ui.tsx                   Section, ResultBanner, useActionRunner, RunButton
      useSelection.ts          Live selection tracking
    features/
      dataCleaning.ts          Pure transforms + Office.js drivers
      merge.ts                 Header planning + merge drivers
      charts.ts                Aggregation + chart drivers
      formulaPatterns.ts       A1 translation + template library
      reportBuilder.ts         Zone mapping + refresh engine
    shared/
      excelHelpers.ts          Grid/address helpers and Office.js utilities
      templateStore.ts         Template persistence in workbook settings
      workbookReader.ts        .xlsx / .csv parsing
      types.ts                 Shared domain types
tests/                         Jest unit tests (179)
manifest.xml                   Add-in manifest
```

Each feature module is split the same way: **pure functions over plain data** at the top (unit
tested), and thin **Office.js drivers** at the bottom that read a range, call the pure function, and
write back. Every driver returns the same `OperationResult` shape, which is what the pane renders.

## Testing

```bash
npm test
```

179 unit tests cover the pure logic: grid transforms, dedupe keying, date parsing and Excel serial
conversion, header planning, merging and sorting, grouped aggregation and chart validation, CSV
parsing, A1 formula translation, and column mapping.

Anything that calls Office.js is verified by sideloading (`npm start`) — mocking the Office.js proxy
object model produces tests that pass while the real thing fails, so it is deliberately not mocked.

> If the test runner crashes with a native V8 error, run `npx jest --runInBand`. That is memory
> pressure from parallel workers, not a test failure.

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
      (the production build rewrites `https://localhost:3000/` to it)
- [ ] Add a `LICENSE` and a privacy policy URL

`<Id>` is already a unique GUID for this add-in — keep it stable across versions.

## Licence

Not yet chosen — add a `LICENSE` file before publishing.
