/**
 * Stopping the mess from coming back.
 *
 * Everything else in MEx finds problems after they happen. Manual entry is the
 * single biggest source of spreadsheet errors, and no amount of checking
 * afterwards is as good as the cell simply refusing the bad value at the moment
 * someone types it.
 *
 * Excel has had this for years - Data Validation - and almost nobody sets it
 * up, because doing it by hand means knowing the rule, finding the dialog, and
 * repeating it per column. But we already know the rule: the column profiler
 * has worked out that Region holds six distinct categories, that Units is a
 * positive whole number, that Order Date runs from January to June. Turning
 * what we already know into a rule is a short step.
 *
 * The rules proposed here are deliberately loose. A guard that rejects a value
 * someone legitimately needs to enter is worse than no guard at all: they will
 * turn the whole thing off, and be right to. So ranges are padded, lists are
 * only offered where the set is genuinely small and settled, and every rule is
 * a warning the user can override rather than a hard stop.
 *
 * Everything here is pure and unit tested in tests/guards.test.ts.
 */

import { excelSerialToDate } from "./dataCleaning";
import { ColumnProfile, parseLooseNumber } from "./tidy";
import { isBlankCell, plural } from "../shared/excelHelpers";
import type { CellValue, Grid } from "../shared/types";

export type GuardKind = "list" | "wholeNumber" | "decimal" | "dateRange" | "unique";

export interface Guard {
  header: string;
  /** Column index within the table. */
  columnIndex: number;
  kind: GuardKind;
  /** One line for the list in the pane. */
  title: string;
  /** What it stops, in the user's terms. */
  detail: string;
  /** Message Excel shows when someone types something the rule rejects. */
  errorMessage: string;
  /** The allowed values, for a list rule. */
  values?: string[];
  /** Bounds for a number or date rule. Dates are Excel serials. */
  min?: number;
  max?: number;
}

/** A list guard is only offered when the column is genuinely a small, settled set. */
const MAX_LIST_VALUES = 25;
const MIN_LIST_ROWS = 8;

/** A number column needs this many values before its range means anything. */
const MIN_NUMBERS = 8;

/** Numeric bounds are padded by this much, so next month's bigger order still fits. */
const NUMBER_HEADROOM = 0.5;

/** Date bounds are padded by a year either way. */
const DATE_HEADROOM_DAYS = 365;

function numbersIn(rows: Grid, index: number, profile: ColumnProfile): number[] {
  const out: number[] = [];
  for (const row of rows) {
    const value = row[index];
    if (isBlankCell(value)) continue;
    const parsed = parseLooseNumber(value, profile.numberStyle);
    if (parsed) out.push(parsed.value);
  }
  return out;
}

function distinctIn(rows: Grid, index: number): string[] {
  const seen = new Map<string, string>();
  for (const row of rows) {
    const value = row[index];
    if (isBlankCell(value)) continue;
    const text = typeof value === "string" ? value.trim() : String(value);
    if (text === "") continue;
    // Keep the first spelling seen, and treat case as the same value.
    if (!seen.has(text.toLowerCase())) seen.set(text.toLowerCase(), text);
  }
  return [...seen.values()];
}

function round(value: number): number {
  if (!Number.isFinite(value)) return 0;
  const magnitude = Math.abs(value);
  if (magnitude >= 1000) return Math.round(value / 100) * 100;
  if (magnitude >= 100) return Math.round(value / 10) * 10;
  if (magnitude >= 1) return Math.round(value);
  return Number(value.toFixed(2));
}

function formatDate(serial: number): string {
  const date = excelSerialToDate(serial);
  return date.toLocaleDateString("en-US", { day: "numeric", month: "short", year: "numeric" });
}

/**
 * Works out which columns could sensibly be guarded, and how.
 *
 * Returns nothing for a column it cannot describe confidently. A guard that has
 * to be guessed at is the kind that gets in the way later.
 */
export function suggestGuards(profiles: readonly ColumnProfile[], rows: Grid): Guard[] {
  const guards: Guard[] = [];

  for (const profile of profiles) {
    const index = profile.index;

    // A short, settled set of categories becomes a dropdown. This is the guard
    // that pays for itself: it stops "Nrth" and "north " ever being typed, which
    // is what splits a chart in two.
    if ((profile.kind === "category" || profile.kind === "text") && rows.length >= MIN_LIST_ROWS) {
      const values = distinctIn(rows, index);
      if (
        values.length >= 2 &&
        values.length <= MAX_LIST_VALUES &&
        values.length < rows.length / 2
      ) {
        guards.push({
          header: profile.header,
          columnIndex: index,
          kind: "list",
          title: `${profile.header} only accepts its ${values.length} existing values`,
          detail: `A dropdown of ${values.slice(0, 4).join(", ")}${values.length > 4 ? ", ..." : ""}. Stops a new spelling splitting your totals in two.`,
          errorMessage: `${profile.header} should be one of: ${values.join(", ")}.`,
          values,
        });
        continue;
      }
    }

    // An identifier that is unique everywhere should stay unique. Excel can
    // enforce that with a COUNTIF over the column's own data.
    if (profile.kind === "id" && rows.length >= MIN_LIST_ROWS) {
      const values = distinctIn(rows, index);
      if (values.length === rows.filter((row) => !isBlankCell(row[index])).length) {
        guards.push({
          header: profile.header,
          columnIndex: index,
          kind: "unique",
          title: `${profile.header} has to stay unique`,
          detail:
            "Every value in this column appears once today. The rule warns if an id is entered twice - the mistake that makes a row silently count double.",
          errorMessage: `That ${profile.header} is already used further up the column.`,
        });
        continue;
      }
    }

    if (profile.kind === "number" || profile.kind === "currency" || profile.kind === "percent") {
      const values = numbersIn(rows, index, profile);
      if (values.length < MIN_NUMBERS) continue;

      const low = Math.min(...values);
      const high = Math.max(...values);
      const allWhole = values.every((value) => Number.isInteger(value));
      const neverNegative = low >= 0;

      // Padding matters: next month's biggest order will be bigger than this
      // month's, and a rule that rejects it would be worse than no rule.
      const span = high - low || Math.abs(high) || 1;
      const min = neverNegative ? 0 : round(low - span * NUMBER_HEADROOM);
      const max = round(high + span * NUMBER_HEADROOM);

      guards.push({
        header: profile.header,
        columnIndex: index,
        kind: allWhole ? "wholeNumber" : "decimal",
        title: neverNegative
          ? `${profile.header} must be a ${allWhole ? "whole number" : "number"} of 0 or more`
          : `${profile.header} must be a ${allWhole ? "whole number" : "number"} between ${min.toLocaleString("en-US")} and ${max.toLocaleString("en-US")}`,
        detail: neverNegative
          ? `Today's values run ${round(low).toLocaleString("en-US")} to ${round(high).toLocaleString("en-US")}. The rule allows up to ${max.toLocaleString("en-US")}, so a bigger month still fits, but catches a stray minus sign or a value typed as text.`
          : `Today's values run ${round(low).toLocaleString("en-US")} to ${round(high).toLocaleString("en-US")}, and the rule leaves room either side.`,
        errorMessage: `${profile.header} should be a ${allWhole ? "whole number" : "number"} between ${min.toLocaleString("en-US")} and ${max.toLocaleString("en-US")}.`,
        min,
        max,
      });
      continue;
    }

    if (profile.kind === "date") {
      const serials: number[] = [];
      for (const row of rows) {
        const value = row[index];
        if (typeof value === "number" && Number.isFinite(value)) serials.push(Math.floor(value));
      }
      if (serials.length < MIN_NUMBERS) continue;

      const low = Math.min(...serials) - DATE_HEADROOM_DAYS;
      const high = Math.max(...serials) + DATE_HEADROOM_DAYS;
      guards.push({
        header: profile.header,
        columnIndex: index,
        kind: "dateRange",
        title: `${profile.header} must be a real date near the others`,
        detail: `Between ${formatDate(low)} and ${formatDate(high)}. Catches a year typed as 2204, and text that only looks like a date.`,
        errorMessage: `${profile.header} should be a date between ${formatDate(low)} and ${formatDate(high)}.`,
        min: low,
        max: high,
      });
    }
  }

  return guards;
}

/** One sentence for the chat. */
export function summarizeGuards(guards: readonly Guard[], sheet: string): string {
  if (guards.length === 0) {
    return `I couldn't see a column on "${sheet}" settled enough to put a rule on. Rules need a table with headings and a few rows of consistent data underneath.`;
  }
  return `${plural(guards.length, "rule")} I can put on "${sheet}" to stop bad values being typed in.`;
}

/** The address of a guarded column's data rows, e.g. "C2:C251". */
export function guardRange(
  columnLetterFor: (index: number) => string,
  columnIndex: number,
  firstDataRow: number,
  lastDataRow: number,
  extraRows: number
): string {
  const letter = columnLetterFor(columnIndex);
  // Guards cover some empty rows below the data too, so they apply to the rows
  // that get typed next - which is the whole point.
  return `${letter}${firstDataRow}:${letter}${lastDataRow + extraRows}`;
}

/** A COUNTIF that is true only when this cell's value appears once in the range. */
export function uniqueFormula(rangeAddress: string, firstCell: string): string {
  const absolute = rangeAddress.replace(/([A-Za-z]+)(\d+)/g, "$$$1$$$2");
  return `=COUNTIF(${absolute},${firstCell})=1`;
}

/** Guards worth applying to a cell that is currently empty, for the list in the pane. */
export function guardCount(guards: readonly Guard[], kind: GuardKind): number {
  return guards.filter((guard) => guard.kind === kind).length;
}

/** Turns one guard into the value the Excel API expects. Kept pure for testing. */
export function toValidationRule(
  guard: Guard,
  context: { rangeAddress: string; firstCell: string }
): Record<string, unknown> {
  switch (guard.kind) {
    case "list":
      return {
        list: {
          inCellDropDown: true,
          // Excel takes a comma-separated string for an inline list.
          source: (guard.values ?? []).join(","),
        },
      };
    case "wholeNumber":
      return {
        wholeNumber: { formula1: guard.min ?? 0, formula2: guard.max ?? 0, operator: "Between" },
      };
    case "decimal":
      return {
        decimal: { formula1: guard.min ?? 0, formula2: guard.max ?? 0, operator: "Between" },
      };
    case "dateRange":
      return {
        date: {
          formula1: toIsoDate(guard.min ?? 0),
          formula2: toIsoDate(guard.max ?? 0),
          operator: "Between",
        },
      };
    case "unique":
      return { custom: { formula: uniqueFormula(context.rangeAddress, context.firstCell) } };
  }
}

/** Excel's date validation takes an ISO date string, not a serial. */
export function toIsoDate(serial: number): string {
  return excelSerialToDate(serial).toISOString().slice(0, 10);
}

/** True when a value would be rejected by the guard. Used to warn before applying. */
export function wouldReject(guard: Guard, value: CellValue): boolean {
  if (isBlankCell(value)) return false;
  switch (guard.kind) {
    case "list": {
      const text =
        typeof value === "string" ? value.trim().toLowerCase() : String(value).toLowerCase();
      return !(guard.values ?? []).some((allowed) => allowed.toLowerCase() === text);
    }
    case "wholeNumber":
    case "decimal": {
      const parsed = parseLooseNumber(value);
      if (!parsed) return true;
      if (guard.kind === "wholeNumber" && !Number.isInteger(parsed.value)) return true;
      return parsed.value < (guard.min ?? -Infinity) || parsed.value > (guard.max ?? Infinity);
    }
    case "dateRange": {
      if (typeof value !== "number") return true;
      return value < (guard.min ?? -Infinity) || value > (guard.max ?? Infinity);
    }
    case "unique":
      return false;
  }
}

/** How many existing cells a guard would have rejected, so we can say so first. */
export function countRejections(guard: Guard, rows: Grid): number {
  let count = 0;
  for (const row of rows) {
    if (wouldReject(guard, row[guard.columnIndex])) count += 1;
  }
  return count;
}
