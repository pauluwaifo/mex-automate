/**
 * Data cleaning: remove duplicate rows, trim whitespace, standardize date
 * formats and standardize text case.
 *
 * The transformations themselves are pure functions over a `Grid` (unit tested
 * in tests/dataCleaning.test.ts); the exported `*Command` functions are the thin
 * Office.js drivers that read a range, apply a transformation and write back.
 */

import {
  fail,
  isBlankCell,
  isBlankRow,
  loadSelectedRange,
  loadUsedRange,
  ok,
  plural,
  rectangularize,
  runExcel,
  writeGrid,
} from "../shared/excelHelpers";
import { CellValue, Grid, OperationResult } from "../shared/types";

// ---------------------------------------------------------------------------
// Options
// ---------------------------------------------------------------------------

/** Which cells a cleaning command runs against. */
export type CleaningScope = "selection" | "usedRange";

export interface RangeScopeOptions {
  scope: CleaningScope;
}

export interface DuplicateOptions extends RangeScopeOptions {
  /** Treat the first row as headers and never remove it. */
  hasHeaderRow: boolean;
  /** Zero-based column offsets to compare. Empty means "compare the whole row". */
  keyColumns: number[];
  /** "Acme" and "ACME " collapse to the same key. */
  ignoreCaseAndSpacing: boolean;
  /** Delete whole worksheet rows instead of blanking the tail of the range. */
  deleteEntireRows: boolean;
}

export interface TrimOptions extends RangeScopeOptions {
  /** Collapse runs of internal whitespace down to a single space. */
  collapseInnerSpaces: boolean;
  /** Convert cells that are only whitespace into truly empty cells. */
  blankOutWhitespaceOnlyCells: boolean;
}

export type TextCaseMode = "upper" | "lower" | "proper" | "sentence";

export interface TextCaseOptions extends RangeScopeOptions {
  mode: TextCaseMode;
  /** Change the heading row too. Off by default, so "Order ID" stays "Order ID". */
  includeHeaderRow?: boolean;
}

/** Number-format strings offered by the date standardizer. */
export const DATE_FORMATS = [
  { label: "2024-03-09  (ISO)", value: "yyyy-mm-dd" },
  { label: "03/09/2024  (US)", value: "mm/dd/yyyy" },
  { label: "09/03/2024  (UK)", value: "dd/mm/yyyy" },
  { label: "09-Mar-2024", value: "dd-mmm-yyyy" },
  { label: "Mar 9, 2024", value: "mmm d, yyyy" },
  { label: "9 March 2024", value: "d mmmm yyyy" },
] as const;

export interface DateOptions extends RangeScopeOptions {
  /** An Excel number-format string, e.g. "yyyy-mm-dd". */
  numberFormat: string;
  /** Read ambiguous 03/04/2024 as 3 April rather than 4 March. */
  dayFirst: boolean;
}

// ---------------------------------------------------------------------------
// Pure transformations
// ---------------------------------------------------------------------------

/** Whitespace Excel users actually hit: non-breaking space, zero-width space, BOM. */
const INVISIBLE_WHITESPACE = /[\u00A0\u1680\u2000-\u200D\u202F\u205F\u3000\uFEFF]/g;

export function trimCell(value: CellValue, options: { collapseInnerSpaces: boolean }): CellValue {
  if (typeof value !== "string") {
    return value;
  }
  let text = value.replace(INVISIBLE_WHITESPACE, " ");
  if (options.collapseInnerSpaces) {
    text = text.replace(/\s+/g, " ");
  }
  return text.trim();
}

export function toProperCase(text: string): string {
  // Split on the separators, keeping them, so "o'brien-smith" capitalises each part.
  return text
    .toLowerCase()
    .replace(/[^\s\-'/]+/g, (word) => word.charAt(0).toUpperCase() + word.slice(1));
}

export function toSentenceCase(text: string): string {
  const lower = text.toLowerCase();
  return lower.replace(
    /(^\s*|[.!?]\s+)([a-z])/g,
    (_match, lead: string, letter: string) => lead + letter.toUpperCase()
  );
}

export function applyTextCase(value: CellValue, mode: TextCaseMode): CellValue {
  if (typeof value !== "string" || value.trim() === "") {
    return value;
  }
  switch (mode) {
    case "upper":
      return value.toUpperCase();
    case "lower":
      return value.toLowerCase();
    case "proper":
      return toProperCase(value);
    case "sentence":
      return toSentenceCase(value);
    default:
      return value;
  }
}

/**
 * Excel counts days from an epoch of 1899-12-30, which lines its serials up with
 * the real calendar from 1900-03-01 onwards. Before that date Excel is one ahead,
 * because it wrongly believes 1900 was a leap year and reserves serial 60 for a
 * 29 February that never existed. Serial 1 is therefore 1900-01-01, not 1900-01-02.
 */
const EXCEL_EPOCH_UTC = Date.UTC(1899, 11, 30);
const MS_PER_DAY = 86400000;
/** The first serial at which Excel and the real calendar agree (1900-03-01). */
const FIRST_ACCURATE_SERIAL = 61;
/** Serial for 9999-12-31; anything beyond is not a date Excel can hold. */
const MAX_EXCEL_SERIAL = 2958465;

export function excelSerialToDate(serial: number): Date {
  // Serial 60 is Excel's phantom 1900-02-29; it resolves to 1900-03-01.
  const adjusted = serial < FIRST_ACCURATE_SERIAL ? serial + 1 : serial;
  return new Date(EXCEL_EPOCH_UTC + Math.round(adjusted * MS_PER_DAY));
}

export function dateToExcelSerial(date: Date): number {
  const utcMidnight = Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate());
  const raw = Math.round((utcMidnight - EXCEL_EPOCH_UTC) / MS_PER_DAY);
  return raw < FIRST_ACCURATE_SERIAL ? raw - 1 : raw;
}

const MONTH_NAMES: Record<string, number> = {
  jan: 0,
  january: 0,
  feb: 1,
  february: 1,
  mar: 2,
  march: 2,
  apr: 3,
  april: 3,
  may: 4,
  jun: 5,
  june: 5,
  jul: 6,
  july: 6,
  aug: 7,
  august: 7,
  sep: 8,
  sept: 8,
  september: 8,
  oct: 9,
  october: 9,
  nov: 10,
  november: 10,
  dec: 11,
  december: 11,
};

/** Excel's own pivot year: a 2-digit year below 30 means 20xx. */
function expandTwoDigitYear(year: number): number {
  if (year >= 100) {
    return year;
  }
  return year < 30 ? 2000 + year : 1900 + year;
}

function buildUtcDate(year: number, monthIndex: number, day: number): Date | null {
  if (monthIndex < 0 || monthIndex > 11 || day < 1 || day > 31) {
    return null;
  }
  const date = new Date(Date.UTC(expandTwoDigitYear(year), monthIndex, day));
  // Rejects 2024-02-31, which JS would silently roll forward to March.
  if (date.getUTCMonth() !== monthIndex || date.getUTCDate() !== day) {
    return null;
  }
  return date;
}

/**
 * Best-effort parse of the date shapes that turn up in exported spreadsheets.
 * Returns null rather than guessing when the text is not a date.
 */
export function parseFlexibleDate(value: CellValue, options: { dayFirst: boolean }): Date | null {
  if (value === null || value === undefined || typeof value === "boolean") {
    return null;
  }

  if (typeof value === "number") {
    if (!Number.isFinite(value) || value < 1 || value > MAX_EXCEL_SERIAL) {
      return null;
    }
    return excelSerialToDate(value);
  }

  const text = value.replace(INVISIBLE_WHITESPACE, " ").trim();
  if (text === "") {
    return null;
  }

  // ISO first: unambiguous, so it never depends on the dayFirst setting.
  const iso = /^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:[T ].*)?$/.exec(text);
  if (iso) {
    return buildUtcDate(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
  }

  // Named months: "9 Mar 2024", "Mar 9, 2024", "9-March-24".
  const dayMonthName = /^(\d{1,2})[\s\-/.]+([A-Za-z]{3,9})[\s\-/.,]+(\d{2,4})$/.exec(text);
  if (dayMonthName) {
    const month = MONTH_NAMES[dayMonthName[2].toLowerCase()];
    if (month !== undefined) {
      return buildUtcDate(Number(dayMonthName[3]), month, Number(dayMonthName[1]));
    }
  }

  const monthNameDay =
    /^([A-Za-z]{3,9})[\s\-/.]+(\d{1,2})(?:st|nd|rd|th)?[\s\-/.,]+(\d{2,4})$/.exec(text);
  if (monthNameDay) {
    const month = MONTH_NAMES[monthNameDay[1].toLowerCase()];
    if (month !== undefined) {
      return buildUtcDate(Number(monthNameDay[3]), month, Number(monthNameDay[2]));
    }
  }

  // All-numeric: 03/09/2024. Ambiguous unless one part is clearly > 12.
  const numeric = /^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})$/.exec(text);
  if (numeric) {
    const first = Number(numeric[1]);
    const second = Number(numeric[2]);
    const year = Number(numeric[3]);
    const firstIsDay = first > 12 ? true : second > 12 ? false : options.dayFirst;
    return firstIsDay
      ? buildUtcDate(year, second - 1, first)
      : buildUtcDate(year, first - 1, second);
  }

  return null;
}

export interface DedupeResult {
  rows: Grid;
  removedCount: number;
  /** Zero-based offsets, within the supplied grid, of the rows that were dropped. */
  removedRowOffsets: number[];
}

/**
 * Cell values are joined with a control character no spreadsheet cell can
 * contain, so ["ab", "c"] and ["a", "bc"] cannot collide into one key.
 */
const KEY_SEPARATOR = "\u0001";

function dedupeKey(
  row: readonly CellValue[],
  keyColumns: number[],
  ignoreCaseAndSpacing: boolean
): string {
  const columns = keyColumns.length > 0 ? keyColumns : row.map((_cell, index) => index);
  return columns
    .map((column) => {
      const cell = row[column] ?? null;
      // An empty cell and an empty string are indistinguishable on the grid, so
      // they have to key the same way.
      if (cell === null || cell === "") {
        return "";
      }
      const text = String(cell);
      return ignoreCaseAndSpacing ? text.replace(/\s+/g, " ").trim().toLowerCase() : text;
    })
    .join(KEY_SEPARATOR);
}

/** Keep the first occurrence of each key; report which rows were dropped. */
export function dedupeGrid(grid: Grid, options: Omit<DuplicateOptions, "scope">): DedupeResult {
  const rows = rectangularize(grid);
  const kept: Grid = [];
  const removedRowOffsets: number[] = [];
  const seen = new Set<string>();

  rows.forEach((row, index) => {
    if (options.hasHeaderRow && index === 0) {
      kept.push(row);
      return;
    }
    // Blank rows are padding, not data - never fold them into one another.
    if (isBlankRow(row)) {
      kept.push(row);
      return;
    }
    const key = dedupeKey(row, options.keyColumns, options.ignoreCaseAndSpacing);
    if (seen.has(key)) {
      removedRowOffsets.push(index);
      return;
    }
    seen.add(key);
    kept.push(row);
  });

  return { rows: kept, removedCount: removedRowOffsets.length, removedRowOffsets };
}

/** Apply a per-cell transformation, reporting how many cells actually changed. */
export function mapGrid(
  grid: Grid,
  transform: (value: CellValue, rowIndex: number, columnIndex: number) => CellValue
): { grid: Grid; changedCount: number } {
  let changedCount = 0;
  const mapped = grid.map((row, rowIndex) =>
    row.map((value, columnIndex) => {
      const next = transform(value, rowIndex, columnIndex);
      if (next !== value) {
        changedCount += 1;
      }
      return next;
    })
  );
  return { grid: mapped, changedCount };
}

/**
 * True when the first row reads as headings: two or more labels, none of them
 * a number or a date, sitting above rows that hold values.
 */
export function looksLikeHeaderRow(grid: Grid): boolean {
  if (grid.length < 2) {
    return false;
  }
  const labels = grid[0].filter((value) => typeof value === "string" && value.trim() !== "");
  if (
    labels.length < 2 ||
    labels.length !== grid[0].filter((value) => !isBlankCell(value)).length
  ) {
    return false;
  }
  if (
    labels.some(
      (label) =>
        parseFlexibleDate(label, { dayFirst: true }) !== null ||
        /^[\s$£€]*[\d,.]+[%\s]*$/.test(String(label))
    )
  ) {
    return false;
  }
  // Something below must be a value rather than another label.
  return grid
    .slice(1)
    .some((row) =>
      row.some(
        (value) =>
          typeof value === "number" ||
          (typeof value === "string" && /^[\s$£€]*[\d,.]+[%\s]*$/.test(value))
      )
    );
}

/** Group ascending row offsets into contiguous [start, count] blocks. */
export function groupContiguous(
  offsets: readonly number[]
): Array<{ start: number; count: number }> {
  const sorted = [...offsets].sort((a, b) => a - b);
  const groups: Array<{ start: number; count: number }> = [];
  for (const offset of sorted) {
    const last = groups[groups.length - 1];
    if (last && offset === last.start + last.count) {
      last.count += 1;
    } else {
      groups.push({ start: offset, count: 1 });
    }
  }
  return groups;
}

// ---------------------------------------------------------------------------
// Office.js drivers
// ---------------------------------------------------------------------------

interface TargetRange {
  range: Excel.Range;
  sheet: Excel.Worksheet;
  values: Grid;
  rowIndex: number;
  columnIndex: number;
  rowCount: number;
  columnCount: number;
  address: string;
}

/**
 * Resolve what a command should act on. A single-cell selection is treated as
 * "the whole table" - that is what users expect when they click into their data
 * and hit a cleaning button.
 */
async function resolveTarget(
  context: Excel.RequestContext,
  scope: CleaningScope
): Promise<TargetRange | null> {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const selection = await loadSelectedRange(context);
  const wantsUsedRange =
    scope === "usedRange" || (selection.rowCount === 1 && selection.columnCount === 1);

  if (wantsUsedRange) {
    const used = await loadUsedRange(context, sheet);
    if (!used) {
      return null;
    }
    return {
      range: used,
      sheet,
      values: used.values as Grid,
      rowIndex: used.rowIndex,
      columnIndex: used.columnIndex,
      rowCount: used.rowCount,
      columnCount: used.columnCount,
      address: used.address,
    };
  }

  return {
    range: selection,
    sheet,
    values: selection.values as Grid,
    rowIndex: selection.rowIndex,
    columnIndex: selection.columnIndex,
    rowCount: selection.rowCount,
    columnCount: selection.columnCount,
    address: selection.address,
  };
}

const EMPTY_SELECTION =
  "There is no data to clean. Select a range, or add data to this sheet, and try again.";

/** Read the headers of the current target so the UI can offer key-column pickers. */
export async function readTargetHeaders(scope: CleaningScope): Promise<string[]> {
  try {
    return await Excel.run(async (context) => {
      const target = await resolveTarget(context, scope);
      if (!target || target.values.length === 0) {
        return [];
      }
      return target.values[0].map((cell, index) => {
        const label = String(cell ?? "").trim();
        return label === "" ? `Column ${index + 1}` : label;
      });
    });
  } catch {
    return [];
  }
}

export async function removeDuplicateRows(options: DuplicateOptions): Promise<OperationResult> {
  return runExcel(async (context) => {
    const target = await resolveTarget(context, options.scope);
    if (!target) {
      return fail(EMPTY_SELECTION);
    }

    const { rows, removedCount, removedRowOffsets } = dedupeGrid(target.values, options);
    if (removedCount === 0) {
      return ok(`No duplicate rows found in ${target.address}.`);
    }

    if (options.deleteEntireRows) {
      // Delete from the bottom up so earlier offsets stay valid as rows vanish.
      const groups = groupContiguous(removedRowOffsets).reverse();
      for (const group of groups) {
        target.sheet
          .getRangeByIndexes(target.rowIndex + group.start, 0, group.count, 1)
          .getEntireRow()
          .delete(Excel.DeleteShiftDirection.up);
      }
      await context.sync();
      return ok(`Removed ${plural(removedCount, "duplicate row")} from ${target.address}.`);
    }

    // Write the survivors back at the top, then blank out the tail they vacated.
    await writeGrid(context, target.sheet, target.rowIndex, target.columnIndex, rows);
    target.sheet
      .getRangeByIndexes(
        target.rowIndex + rows.length,
        target.columnIndex,
        removedCount,
        target.columnCount
      )
      .clear(Excel.ClearApplyTo.contents);
    await context.sync();

    return ok(`Removed ${plural(removedCount, "duplicate row")} from ${target.address}.`);
  });
}

export async function trimWhitespace(options: TrimOptions): Promise<OperationResult> {
  return runExcel(async (context) => {
    const target = await resolveTarget(context, options.scope);
    if (!target) {
      return fail(EMPTY_SELECTION);
    }

    const { grid, changedCount } = mapGrid(target.values, (value) => {
      const trimmed = trimCell(value, options);
      if (options.blankOutWhitespaceOnlyCells && trimmed === "") {
        return null;
      }
      return trimmed;
    });

    if (changedCount === 0) {
      return ok(`Nothing to trim in ${target.address}.`);
    }

    await writeGrid(context, target.sheet, target.rowIndex, target.columnIndex, grid);
    return ok(`Trimmed ${plural(changedCount, "cell")} in ${target.address}.`);
  });
}

export async function standardizeTextCase(options: TextCaseOptions): Promise<OperationResult> {
  return runExcel(async (context) => {
    const target = await resolveTarget(context, options.scope);
    if (!target) {
      return fail(EMPTY_SELECTION);
    }

    // Formulas must not be overwritten with their own displayed text.
    target.range.load("formulas");
    await context.sync();
    const formulas = target.range.formulas as string[][];

    // Headings are names, not prose: Title Case would turn "Order ID" into
    // "Order Id" and break anything matching on the column name.
    const skipHeaderRow = !options.includeHeaderRow && looksLikeHeaderRow(target.values);

    const { grid, changedCount } = mapGrid(target.values, (value, rowIndex, columnIndex) => {
      if (skipHeaderRow && rowIndex === 0) {
        return value;
      }
      const formula = formulas[rowIndex]?.[columnIndex];
      if (typeof formula === "string" && formula.startsWith("=")) {
        return value;
      }
      return applyTextCase(value, options.mode);
    });

    if (changedCount === 0) {
      return ok(`No text to change in ${target.address}.`);
    }

    await writeGrid(context, target.sheet, target.rowIndex, target.columnIndex, grid);
    return ok(
      `Updated the case of ${plural(changedCount, "cell")} in ${target.address}.`,
      skipHeaderRow ? ["The heading row was left as it is."] : undefined
    );
  });
}

/** Excel number formats that contain a day, month or year token. */
function looksLikeDateFormat(numberFormat: string): boolean {
  return /[ymd]/i.test(numberFormat.replace(/\[[^\]]*\]/g, "").replace(/"[^"]*"/g, ""));
}

export async function standardizeDates(options: DateOptions): Promise<OperationResult> {
  return runExcel(async (context) => {
    const target = await resolveTarget(context, options.scope);
    if (!target) {
      return fail(EMPTY_SELECTION);
    }

    target.range.load(["numberFormat", "formulas"]);
    await context.sync();
    const numberFormats = target.range.numberFormat as string[][];
    const formulas = target.range.formulas as string[][];

    const serials: Grid = [];
    const formats: string[][] = [];
    let convertedCount = 0;
    let reformattedCount = 0;
    let skippedCount = 0;

    target.values.forEach((row, rowIndex) => {
      const serialRow: CellValue[] = [];
      const formatRow: string[] = [];

      row.forEach((value, columnIndex) => {
        const existingFormat = numberFormats[rowIndex]?.[columnIndex] ?? "General";
        const formula = formulas[rowIndex]?.[columnIndex];
        const isFormula = typeof formula === "string" && formula.startsWith("=");

        // A formula that already yields a date only needs its display format changed.
        if (isFormula) {
          if (typeof value === "number" && looksLikeDateFormat(existingFormat)) {
            serialRow.push(formula);
            formatRow.push(options.numberFormat);
            reformattedCount += 1;
          } else {
            serialRow.push(formula);
            formatRow.push(existingFormat);
          }
          return;
        }

        // A number is only a date if Excel is already displaying it as one -
        // otherwise 42 is a quantity, not 1915-01-11.
        if (typeof value === "number") {
          if (looksLikeDateFormat(existingFormat)) {
            serialRow.push(value);
            formatRow.push(options.numberFormat);
            reformattedCount += 1;
          } else {
            serialRow.push(value);
            formatRow.push(existingFormat);
          }
          return;
        }

        const parsed = parseFlexibleDate(value, options);
        if (parsed) {
          serialRow.push(dateToExcelSerial(parsed));
          formatRow.push(options.numberFormat);
          convertedCount += 1;
          return;
        }

        if (typeof value === "string" && value.trim() !== "") {
          skippedCount += 1;
        }
        serialRow.push(value);
        formatRow.push(existingFormat);
      });

      serials.push(serialRow);
      formats.push(formatRow);
    });

    if (convertedCount === 0 && reformattedCount === 0) {
      return fail(
        `No dates recognised in ${target.address}.`,
        skippedCount > 0
          ? [`${plural(skippedCount, "text value")} could not be read as a date.`]
          : undefined
      );
    }

    const body = target.sheet.getRangeByIndexes(
      target.rowIndex,
      target.columnIndex,
      target.rowCount,
      target.columnCount
    );
    // Number format first, so the serials render correctly the moment they land.
    body.numberFormat = formats;
    body.formulas = serials as unknown[][];
    await context.sync();

    const details: string[] = [];
    if (convertedCount > 0) {
      details.push(`${plural(convertedCount, "text date")} converted to real dates.`);
    }
    if (reformattedCount > 0) {
      details.push(`${plural(reformattedCount, "existing date")} reformatted.`);
    }
    if (skippedCount > 0) {
      details.push(`${plural(skippedCount, "value")} left alone - not recognisable as a date.`);
    }

    return ok(`Standardized dates in ${target.address}.`, details);
  });
}

/** Address of the range a command would act on, for the UI's "acting on ..." hint. */
export async function describeTarget(scope: CleaningScope): Promise<string> {
  try {
    return await Excel.run(async (context) => {
      const target = await resolveTarget(context, scope);
      if (!target) {
        return "(no data)";
      }
      return target.address;
    });
  } catch {
    return "(unavailable)";
  }
}
