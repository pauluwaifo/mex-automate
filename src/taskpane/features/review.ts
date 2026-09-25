/**
 * Review: look over a sheet the way a careful colleague would and list what
 * looks wrong, where, and what to do about it.
 *
 * "Fix messy data" (tidy.ts) repairs the *shape* of an export. Review is about
 * mistakes in data that already looks fine: a hardcoded number sitting in a
 * column of formulas, a SUM that misses the last row, an amount typed as text
 * so it never adds up, an ID that appears twice, a value far outside the rest.
 * These are the errors that quietly produce wrong totals.
 *
 * Every issue carries the exact cell, a plain-English explanation, and - where
 * it is safe to do so - a fix that can be applied with one click. Anything
 * needing human judgement (an outlier, a duplicate ID) is reported without a
 * fix, so nothing is changed on a guess.
 *
 * Everything in this file is pure and unit tested in tests/review.test.ts; the
 * Office.js side lives in reviewSheet.ts.
 */

import {
  buildAddress,
  columnIndexFromLetter,
  columnLetter,
  normalizeHeader,
} from "../shared/excelHelpers";
import { CellValue, Grid } from "../shared/types";
import { dateToExcelSerial, parseFlexibleDate } from "./dataCleaning";
import { translateFormula } from "./formulaPatterns";
import {
  buildHeaders,
  classifyRow,
  cleanText,
  ColumnProfile,
  isPlaceholder,
  locateTable,
  numberFormatFor,
  parseLooseNumber,
  profileColumn,
  TableLocation,
  unifySpellings,
} from "./tidy";

// ---------------------------------------------------------------------------
// The issue model
// ---------------------------------------------------------------------------

export type Severity = "error" | "warning" | "tidy";

export type IssueKind =
  | "inconsistentFormula"
  | "shortSumRange"
  | "textNumber"
  | "textDate"
  | "placeholder"
  | "outlier"
  | "duplicateId"
  | "blankKey"
  | "nearDuplicate"
  | "negativeQuantity"
  | "futureDate"
  | "ancientDate"
  | "mixedCurrency"
  | "straySpaces";

/** What applying a fix does to the cell. */
export type FixAction =
  | { type: "setValue"; value: CellValue; numberFormat?: string }
  | { type: "setFormula"; formula: string };

export interface Issue {
  /** Stable across scans of the same sheet, so "ignore" sticks. */
  id: string;
  kind: IssueKind;
  severity: Severity;
  /** Groups issues for "fix all like this": one kind in one column. */
  group: string;
  /** Short name of the problem, shared by every issue in the group. */
  title: string;
  /** What is wrong with this cell, in plain words. */
  detail: string;
  /** Zero-based sheet coordinates. */
  row: number;
  column: number;
  /** A1 reference within the sheet, e.g. "C7". */
  address: string;
  header: string;
  value: CellValue;
  /** Null when the fix needs a human decision. */
  fix: FixAction | null;
}

export interface ReviewOptions {
  /** Today, as an Excel serial. Injectable so tests don't drift. */
  today?: number;
  /** Stop after this many issues of one kind in one column. */
  perGroupLimit?: number;
}

export interface ReviewResult {
  issues: Issue[];
  columns: ColumnProfile[];
  headers: string[];
  location: TableLocation | null;
  /** Data rows examined. */
  rows: number;
  /** Rows that aren't data (titles, totals, group headings) - a hint to run /fix. */
  structuralRows: number;
}

export interface ReviewInput {
  grid: Grid;
  /** Same shape as `grid`; "" where a cell holds a constant. */
  formulas?: ReadonlyArray<ReadonlyArray<string>>;
  numberFormats?: ReadonlyArray<ReadonlyArray<string>>;
  /** Zero-based sheet position of grid[0][0], for addresses. */
  origin?: { row: number; column: number };
}

export const SEVERITY_ORDER: Severity[] = ["error", "warning", "tidy"];

const QUANTITY_HINT = /\b(qty|quantity|units?|count|stock|pieces|pcs|hours|days|weight|age)\b/i;

// ---------------------------------------------------------------------------
// Small statistics
// ---------------------------------------------------------------------------

export function median(values: readonly number[]): number {
  if (values.length === 0) return 0;
  const sorted = [...values].sort((a, b) => a - b);
  const middle = Math.floor(sorted.length / 2);
  return sorted.length % 2 === 0 ? (sorted[middle - 1] + sorted[middle]) / 2 : sorted[middle];
}

/**
 * Values far outside the rest, by median absolute deviation - which, unlike a
 * standard deviation, isn't dragged around by the very outliers it's looking
 * for. Returns the indexes of the offending values.
 */
export function outlierIndexes(values: readonly number[], threshold = 8): number[] {
  if (values.length < 12) return [];
  const mid = median(values);
  const mad = median(values.map((value) => Math.abs(value - mid)));
  if (mad === 0) return [];
  const flagged: number[] = [];
  values.forEach((value, index) => {
    if (Math.abs((0.6745 * (value - mid)) / mad) > threshold) flagged.push(index);
  });
  // A column where a tenth of the values are "outliers" is just a spread-out
  // column; only report when they really are the exception.
  return flagged.length > Math.max(3, values.length * 0.05) ? [] : flagged;
}

// ---------------------------------------------------------------------------
// Formulas
// ---------------------------------------------------------------------------

/** The formula rewritten as if it sat in the first row, so a column can be compared. */
function formulaShape(formula: string, rowOffset: number): string {
  return translateFormula(formula, -rowOffset, 0).toUpperCase().replace(/\s+/g, "");
}

const SUM_RANGE = /^=SUM\(([A-Z]+)(\d+):([A-Z]+)(\d+)\)$/i;

/**
 * A total that misses rows: "=SUM(C4:C9)" sitting under numbers that carry on
 * to row 11. `hasNumber` is asked about Excel row numbers and column letters,
 * the same ones the formula is written in. Returns the range the SUM should
 * cover, or null when it already covers everything.
 */
export function widenSumRange(
  formula: string,
  hasNumber: (excelRow: number, column: string) => boolean,
  limit: { first: number; last: number }
): string | null {
  const match = SUM_RANGE.exec(formula.replace(/\s+/g, ""));
  if (!match) return null;
  const [, startColumn, startRow, endColumn, endRow] = match;
  if (startColumn.toUpperCase() !== endColumn.toUpperCase()) return null;

  let first = Number(startRow);
  let last = Number(endRow);
  while (first - 1 >= limit.first && hasNumber(first - 1, startColumn)) first -= 1;
  while (last + 1 <= limit.last && hasNumber(last + 1, startColumn)) last += 1;
  if (first === Number(startRow) && last === Number(endRow)) return null;
  return `=SUM(${startColumn}${first}:${endColumn}${last})`;
}

// ---------------------------------------------------------------------------
// The review
// ---------------------------------------------------------------------------

interface Cursor {
  issues: Issue[];
  counts: Map<string, number>;
  limit: number;
}

function add(cursor: Cursor, issue: Omit<Issue, "id" | "address"> & { address?: string }): void {
  const seen = cursor.counts.get(issue.group) ?? 0;
  if (seen >= cursor.limit) return;
  cursor.counts.set(issue.group, seen + 1);
  const address = issue.address ?? `${columnLetter(issue.column)}${issue.row + 1}`;
  cursor.issues.push({ ...issue, address, id: `${issue.kind}:${address}` });
}

/** Look over a table and list everything that looks wrong. */
export function reviewTable(input: ReviewInput, options: ReviewOptions = {}): ReviewResult {
  const grid = input.grid;
  const origin = input.origin ?? { row: 0, column: 0 };
  const today = options.today ?? dateToExcelSerial(new Date());
  const cursor: Cursor = { issues: [], counts: new Map(), limit: options.perGroupLimit ?? 100 };
  const location = locateTable(grid);
  if (!location) {
    return { issues: [], columns: [], headers: [], location: null, rows: 0, structuralRows: 0 };
  }

  const { headerRow, headerRows, firstColumn, lastColumn, lastRow } = location;
  const slice = (row: readonly CellValue[]) => row.slice(firstColumn, lastColumn + 1);
  const width = lastColumn - firstColumn + 1;
  const upper = headerRows === 2 ? slice(grid[headerRow - 1]) : null;
  const { headers } = buildHeaders(upper, slice(grid[headerRow]));
  const headerKeys = slice(grid[headerRow]).map((cell) => normalizeHeader(cell));

  // Only real data rows are examined; titles, totals and group headings are
  // structure, which "Fix messy data" handles.
  const dataRows: number[] = [];
  let structuralRows = 0;
  for (let r = headerRow + 1; r <= lastRow; r += 1) {
    const kind = classifyRow(slice(grid[r]), headerKeys, width);
    if (kind === "data") dataRows.push(r);
    else if (kind !== "blank") structuralRows += 1;
  }

  const columnValues = (column: number) =>
    dataRows.map((r) => grid[r][firstColumn + column] ?? null);
  const columns = headers.map((header, c) =>
    profileColumn(header, columnValues(c), c, {
      formats: input.numberFormats
        ? dataRows.map((r) => input.numberFormats![r]?.[firstColumn + c])
        : undefined,
    })
  );

  const sheetRow = (r: number) => origin.row + r;
  const sheetColumn = (c: number) => origin.column + firstColumn + c;
  const cellAddress = (r: number, c: number) => `${columnLetter(sheetColumn(c))}${sheetRow(r) + 1}`;

  detectFormulaIssues(cursor, {
    grid,
    input,
    location,
    headers,
    dataRows,
    origin,
    sheetRow,
    sheetColumn,
    cellAddress,
  });

  columns.forEach((profile, c) => {
    const header = profile.header;
    const numeric =
      profile.kind === "number" || profile.kind === "currency" || profile.kind === "percent";
    const values = columnValues(c);
    const base = { column: sheetColumn(c), header };

    // --- Values that should be numbers or dates but are text --------------
    dataRows.forEach((r, i) => {
      const value = values[i];
      const address = cellAddress(r, c);
      if (typeof value !== "string" || value.trim() === "") return;

      if (isPlaceholder(value) && (numeric || profile.kind === "date")) {
        add(cursor, {
          kind: "placeholder",
          severity: "tidy",
          group: `placeholder:${c}`,
          title: "Placeholder in a number column",
          detail: `"${cleanText(value)}" in ${header} isn't a number, so it's skipped by sums and charts.`,
          row: sheetRow(r),
          ...base,
          address,
          value,
          fix: { type: "setValue", value: null },
        });
        return;
      }

      if (numeric) {
        const parsed = parseLooseNumber(value, profile.numberStyle);
        if (parsed) {
          add(cursor, {
            kind: "textNumber",
            severity: "error",
            group: `textNumber:${c}`,
            title: "Number stored as text",
            detail: `${header} holds "${cleanText(value)}" as text, so it doesn't count towards totals. It means ${parsed.value.toLocaleString("en-US")}.`,
            row: sheetRow(r),
            ...base,
            address,
            value,
            fix: {
              type: "setValue",
              value: parsed.value,
              numberFormat: numberFormatFor(profile, values),
            },
          });
        }
        return;
      }

      if (profile.kind === "date") {
        const parsed = parseFlexibleDate(cleanText(value), { dayFirst: profile.dayFirst });
        if (parsed) {
          add(cursor, {
            kind: "textDate",
            severity: "warning",
            group: `textDate:${c}`,
            title: "Date stored as text",
            detail: `${header} holds "${cleanText(value)}" as text, so it won't sort or filter by date.`,
            row: sheetRow(r),
            ...base,
            address,
            value,
            fix: { type: "setValue", value: dateToExcelSerial(parsed), numberFormat: "yyyy-mm-dd" },
          });
        }
      }
    });

    // --- Stray spaces ------------------------------------------------------
    dataRows.forEach((r, i) => {
      const value = values[i];
      if (typeof value !== "string" || value === "" || isPlaceholder(value)) return;
      const cleaned = cleanText(value);
      if (cleaned === value || cleaned === "") return;
      add(cursor, {
        kind: "straySpaces",
        severity: "tidy",
        group: `straySpaces:${c}`,
        title: "Extra spaces",
        detail: `"${value}" has spaces that stop it matching "${cleaned}" elsewhere.`,
        row: sheetRow(r),
        ...base,
        address: cellAddress(r, c),
        value,
        fix: { type: "setValue", value: cleaned },
      });
    });

    // --- Numbers that look wrong ------------------------------------------
    if (numeric) {
      const numbers: Array<{ value: number; row: number }> = [];
      dataRows.forEach((r, i) => {
        const parsed = parseLooseNumber(values[i], profile.numberStyle);
        if (parsed) numbers.push({ value: parsed.value, row: r });
      });

      outlierIndexes(numbers.map((item) => item.value)).forEach((index) => {
        const { value, row } = numbers[index];
        add(cursor, {
          kind: "outlier",
          severity: "warning",
          group: `outlier:${c}`,
          title: "Value far outside the rest",
          detail: `${value.toLocaleString("en-US")} is a long way from the typical ${header} of about ${median(numbers.map((item) => item.value)).toLocaleString("en-US")}. Worth a look - it may be a typo or an extra zero.`,
          row: sheetRow(row),
          ...base,
          address: cellAddress(row, c),
          value,
          fix: null,
        });
      });

      if (QUANTITY_HINT.test(header)) {
        numbers
          .filter((item) => item.value < 0)
          .forEach(({ value, row }) => {
            add(cursor, {
              kind: "negativeQuantity",
              severity: "warning",
              group: `negativeQuantity:${c}`,
              title: "Negative where a count is expected",
              detail: `${header} is ${value.toLocaleString("en-US")}. A count below zero usually means a return recorded the wrong way round.`,
              row: sheetRow(row),
              ...base,
              address: cellAddress(row, c),
              value,
              fix: null,
            });
          });
      }

      if (profile.kind === "currency") {
        const symbols = new Map<string, number>();
        values.forEach((value) => {
          const parsed = parseLooseNumber(value, profile.numberStyle);
          if (parsed?.currency)
            symbols.set(parsed.currency, (symbols.get(parsed.currency) ?? 0) + 1);
        });
        if (symbols.size > 1) {
          const ranked = Array.from(symbols.entries()).sort((a, b) => b[1] - a[1]);
          const odd = ranked.slice(1).map(([symbol]) => symbol);
          dataRows.forEach((r, i) => {
            const parsed = parseLooseNumber(values[i], profile.numberStyle);
            if (parsed?.currency && odd.includes(parsed.currency)) {
              add(cursor, {
                kind: "mixedCurrency",
                severity: "warning",
                group: `mixedCurrency:${c}`,
                title: "More than one currency in a column",
                detail: `${header} is mostly ${ranked[0][0]} but this is ${parsed.currency}. Totals across the column would mix currencies.`,
                row: sheetRow(r),
                ...base,
                address: cellAddress(r, c),
                value: values[i],
                fix: null,
              });
            }
          });
        }
      }
    }

    // --- Dates that look wrong --------------------------------------------
    if (profile.kind === "date") {
      const ancient = dateToExcelSerial(new Date(Date.UTC(1990, 0, 1)));
      dataRows.forEach((r, i) => {
        const raw = values[i];
        const serial =
          typeof raw === "number"
            ? raw
            : typeof raw === "string"
              ? (() => {
                  const parsed = parseFlexibleDate(cleanText(raw), { dayFirst: profile.dayFirst });
                  return parsed ? dateToExcelSerial(parsed) : null;
                })()
              : null;
        if (serial === null) return;
        if (serial > today + 365) {
          add(cursor, {
            kind: "futureDate",
            severity: "warning",
            group: `futureDate:${c}`,
            title: "Date far in the future",
            detail: `${header} is more than a year from now. Check the year - a mistyped one is the usual cause.`,
            row: sheetRow(r),
            ...base,
            address: cellAddress(r, c),
            value: raw,
            fix: null,
          });
        } else if (serial < ancient) {
          add(cursor, {
            kind: "ancientDate",
            severity: "warning",
            group: `ancientDate:${c}`,
            title: "Date before 1990",
            detail: `${header} is earlier than 1990, which usually means a two-digit year was read the wrong way.`,
            row: sheetRow(r),
            ...base,
            address: cellAddress(r, c),
            value: raw,
            fix: null,
          });
        }
      });
    }

    // --- Keys and labels ----------------------------------------------------
    if (profile.kind === "id") {
      const seen = new Map<string, number>();
      dataRows.forEach((r, i) => {
        const value = values[i];
        if (value === null || value === "" || isPlaceholder(value)) {
          add(cursor, {
            kind: "blankKey",
            severity: "warning",
            group: `blankKey:${c}`,
            title: "Missing ID",
            detail: `${header} is empty on this row, so the row can't be matched to anything else.`,
            row: sheetRow(r),
            ...base,
            address: cellAddress(r, c),
            value,
            fix: null,
          });
          return;
        }
        const key = typeof value === "string" ? cleanText(value).toLowerCase() : String(value);
        const first = seen.get(key);
        if (first === undefined) {
          seen.set(key, r);
          return;
        }
        add(cursor, {
          kind: "duplicateId",
          severity: "error",
          group: `duplicateId:${c}`,
          title: "ID used twice",
          detail: `${header} "${String(value)}" is already on row ${sheetRow(first) + 1}. One of the two rows is probably a duplicate, or the ID is wrong.`,
          row: sheetRow(r),
          ...base,
          address: cellAddress(r, c),
          value,
          fix: null,
        });
      });
    }

    if (profile.kind === "category") {
      const mapping = unifySpellings(values);
      if (mapping.size > 0) {
        dataRows.forEach((r, i) => {
          const value = values[i];
          if (typeof value !== "string") return;
          const canonical = mapping.get(cleanText(value));
          if (!canonical) return;
          add(cursor, {
            kind: "nearDuplicate",
            severity: "warning",
            group: `nearDuplicate:${c}`,
            title: "Same thing, spelled differently",
            detail: `"${cleanText(value)}" is almost certainly "${canonical}". As written they count as two different groups in every chart and total.`,
            row: sheetRow(r),
            ...base,
            address: cellAddress(r, c),
            value,
            fix: { type: "setValue", value: canonical },
          });
        });
      }
    }
  });

  cursor.issues.sort(
    (a, b) =>
      SEVERITY_ORDER.indexOf(a.severity) - SEVERITY_ORDER.indexOf(b.severity) ||
      a.row - b.row ||
      a.column - b.column
  );

  return {
    issues: cursor.issues,
    columns,
    headers,
    location,
    rows: dataRows.length,
    structuralRows,
  };
}

// ---------------------------------------------------------------------------
// Formula detectors
// ---------------------------------------------------------------------------

interface FormulaContext {
  grid: Grid;
  input: ReviewInput;
  location: TableLocation;
  headers: string[];
  dataRows: number[];
  origin: { row: number; column: number };
  sheetRow: (r: number) => number;
  sheetColumn: (c: number) => number;
  cellAddress: (r: number, c: number) => string;
}

function detectFormulaIssues(cursor: Cursor, context: FormulaContext): void {
  const formulas = context.input.formulas;
  if (!formulas) return;
  const { location, headers, dataRows, grid } = context;
  const { firstColumn } = location;

  headers.forEach((header, c) => {
    const column = firstColumn + c;
    const entries = dataRows.map((r) => ({ row: r, formula: String(formulas[r]?.[column] ?? "") }));
    const written = entries.filter((entry) => entry.formula.startsWith("="));
    if (written.length < 3) return;

    // The shape most of the column agrees on, once row references are lined up.
    const tally = new Map<string, number>();
    written.forEach((entry) => {
      const shape = formulaShape(entry.formula, entry.row);
      tally.set(shape, (tally.get(shape) ?? 0) + 1);
    });
    const [dominant, count] = Array.from(tally.entries()).sort((a, b) => b[1] - a[1])[0];
    if (count < Math.max(3, written.length * 0.6)) return;
    const template = written.find((entry) => formulaShape(entry.formula, entry.row) === dominant)!;

    entries.forEach((entry) => {
      const value = grid[entry.row][column] ?? null;
      const isBlank = value === null || value === "";
      if (isBlank) return;
      const expected = translateFormula(template.formula, entry.row - template.row, 0);
      if (entry.formula.startsWith("=")) {
        if (formulaShape(entry.formula, entry.row) === dominant) return;
        cursorAddFormula(cursor, context, c, entry.row, value, expected, header, "different");
      } else {
        cursorAddFormula(cursor, context, c, entry.row, value, expected, header, "typed");
      }
    });
  });

  // Totals that miss rows: check every SUM in and just below the table.
  const lastFormulaRow = Math.min(formulas.length - 1, location.lastRow + 3);
  for (let r = location.headerRow + 1; r <= lastFormulaRow; r += 1) {
    for (let column = location.firstColumn; column <= location.lastColumn; column += 1) {
      const formula = String(formulas[r]?.[column] ?? "");
      if (!formula.toUpperCase().startsWith("=SUM(")) continue;
      // Ask about the formula's own column, which isn't always the cell's column.
      const widened = widenSumRange(
        formula,
        (excelRow, letter) => {
          const gridRow = excelRow - 1 - context.origin.row;
          const gridColumn = columnIndexFromLetter(letter) - context.origin.column;
          return typeof grid[gridRow]?.[gridColumn] === "number" && dataRows.indexOf(gridRow) >= 0;
        },
        {
          first: location.headerRow + context.origin.row + 2,
          last: location.lastRow + context.origin.row + 1,
        }
      );
      if (!widened) continue;
      const address = context.cellAddress(r, column - location.firstColumn);
      add(cursor, {
        kind: "shortSumRange",
        severity: "error",
        group: `shortSumRange:${column}`,
        title: "Total misses rows",
        detail: `${formula} leaves out rows that sit right beside it, so this total is wrong. It should be ${widened}.`,
        row: context.sheetRow(r),
        column: context.sheetColumn(column - location.firstColumn),
        address,
        header: headers[column - location.firstColumn] ?? "",
        value: grid[r]?.[column] ?? null,
        fix: { type: "setFormula", formula: widened },
      });
    }
  }
}

function cursorAddFormula(
  cursor: Cursor,
  context: FormulaContext,
  c: number,
  row: number,
  value: CellValue,
  expected: string,
  header: string,
  reason: "typed" | "different"
): void {
  add(cursor, {
    kind: "inconsistentFormula",
    severity: "error",
    group: `inconsistentFormula:${c}`,
    title:
      reason === "typed" ? "Typed-in number among formulas" : "Formula doesn't match the column",
    detail:
      reason === "typed"
        ? `Every other row of ${header} calculates this value, but this one is typed in. It won't update when the data changes. The column's formula here is ${expected}.`
        : `This formula is different from the rest of ${header}, which uses ${expected}. One of them is wrong.`,
    row: context.sheetRow(row),
    column: context.sheetColumn(c),
    address: context.cellAddress(row, c),
    header,
    value,
    fix: { type: "setFormula", formula: expected },
  });
}

// ---------------------------------------------------------------------------
// Grouping for the UI
// ---------------------------------------------------------------------------

export interface IssueGroup {
  group: string;
  kind: IssueKind;
  severity: Severity;
  title: string;
  header: string;
  issues: Issue[];
  /** How many of them can be fixed in one click. */
  fixable: number;
}

/** One entry per kind of problem per column, worst first. */
export function groupIssues(issues: readonly Issue[]): IssueGroup[] {
  const groups = new Map<string, IssueGroup>();
  for (const issue of issues) {
    let entry = groups.get(issue.group);
    if (!entry) {
      entry = {
        group: issue.group,
        kind: issue.kind,
        severity: issue.severity,
        title: issue.title,
        header: issue.header,
        issues: [],
        fixable: 0,
      };
      groups.set(issue.group, entry);
    }
    entry.issues.push(issue);
    if (issue.fix) entry.fixable += 1;
  }
  return Array.from(groups.values()).sort(
    (a, b) =>
      SEVERITY_ORDER.indexOf(a.severity) - SEVERITY_ORDER.indexOf(b.severity) ||
      b.issues.length - a.issues.length
  );
}

/** "3 errors, 12 warnings" - the one-line verdict. */
export function summarize(issues: readonly Issue[]): string {
  const counts = { error: 0, warning: 0, tidy: 0 };
  issues.forEach((issue) => {
    counts[issue.severity] += 1;
  });
  const parts: string[] = [];
  if (counts.error) parts.push(`${counts.error} ${counts.error === 1 ? "error" : "errors"}`);
  if (counts.warning) parts.push(`${counts.warning} to check`);
  if (counts.tidy) parts.push(`${counts.tidy} tidy-up${counts.tidy === 1 ? "" : "s"}`);
  return parts.length === 0 ? "nothing to flag" : parts.join(", ");
}

/** The address of a whole group, for selecting it in one go. */
export function groupAddress(group: IssueGroup, sheetName: string): string {
  const rows = group.issues.map((issue) => issue.row);
  const column = group.issues[0].column;
  return buildAddress(
    sheetName,
    Math.min(...rows),
    column,
    Math.max(...rows) - Math.min(...rows) + 1,
    1
  );
}
