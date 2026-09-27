/**
 * Turning crosstabs back into tables.
 *
 * Plenty of real exports are shaped for reading, not for analysis: one row per
 * product and then twelve columns called Jan, Feb, Mar... That layout defeats
 * every chart, PivotTable and SUMIFS you might want to write, because the thing
 * you want on an axis - the month - isn't in a column at all, it's in the
 * headings.
 *
 * The fix is called unpivoting, and Excel's own answer is Power Query, which is
 * a lot of machinery to learn for one transformation. This module does the same
 * job by reading the headings: if they look like periods, the wide block is a
 * crosstab and each cell in it becomes a row.
 *
 *   Product  Jan  Feb        Product  Month     Value
 *   Widget   100  120   ->   Widget   Jan 2024  100
 *   Gadget    80   95        Widget   Feb 2024  120
 *                            Gadget   Jan 2024   80
 *                            Gadget   Feb 2024   95
 *
 * Headings that name both a period and a measure ("Jan Units", "Jan Revenue")
 * are handled too: those become one row per period with a column per measure,
 * which is what you actually want to chart.
 *
 * Everything here is pure and unit tested in tests/reshape.test.ts.
 */

import { dateToExcelSerial } from "./dataCleaning";
import { parseLooseNumber } from "./tidy";
import { isBlankCell, normalizeHeader } from "../shared/excelHelpers";
import type { CellValue, Grid, HeaderTable } from "../shared/types";

// ---------------------------------------------------------------------------
// Reading a period out of a column heading
// ---------------------------------------------------------------------------

export type PeriodGrain = "month" | "quarter" | "year";

export interface PeriodHeader {
  /** The heading this came from, unchanged. */
  source: string;
  /** How to label the period in the tidy sheet: "Jan 2024", "Q1 2024", "2024". */
  label: string;
  /** Sorts periods chronologically even when the columns are out of order. */
  sortKey: number;
  /** First day of the period, when the year is known. Null when it isn't. */
  date: Date | null;
  /** The leftover words in a heading like "Jan Units", or null. */
  measure: string | null;
  grain: PeriodGrain;
  /** True when the heading named a period but no year, e.g. plain "Jan". */
  yearAssumed: boolean;
}

const MONTHS: Record<string, number> = {
  jan: 0,
  january: 0,
  feb: 1,
  february: 1,
  febuary: 1,
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

const MONTH_LABELS = [
  "Jan",
  "Feb",
  "Mar",
  "Apr",
  "May",
  "Jun",
  "Jul",
  "Aug",
  "Sep",
  "Oct",
  "Nov",
  "Dec",
];

/**
 * Two-digit years: 00-69 are this century, 70-99 the last. The same split Excel
 * itself uses, so a sheet reading "Jan 98" means 1998 in both places.
 */
function fullYear(text: string): number {
  const year = Number(text);
  if (text.length === 4) return year;
  return year <= 69 ? 2000 + year : 1900 + year;
}

/**
 * Letters and digits that run together, which headings do constantly: "JAN24",
 * "2024Q1". A quarter is left whole, because "Q1" split into "Q" and "1" stops
 * looking like a quarter at all.
 */
const RUN_TOGETHER: RegExp[] = [
  /^(\d{2,4})(q[1-4])$/i,
  /^(q[1-4])(\d{2,4})$/i,
  /^([a-z]+)(\d+)$/i,
  /^(\d+)([a-z]+)$/i,
];

/** Splits a heading into words, treating punctuation as a separator. */
function words(header: string): string[] {
  const out: string[] = [];
  for (const token of header.split(/[^a-z0-9]+/i)) {
    if (token === "") continue;
    if (/^q[1-4]$/i.test(token)) {
      out.push(token);
      continue;
    }
    const split = RUN_TOGETHER.map((pattern) => pattern.exec(token)).find(
      (match) => match !== null
    );
    if (split) out.push(split[1], split[2]);
    else out.push(token);
  }
  return out;
}

/**
 * Reads a period from the front of a token list, returning how many tokens it
 * consumed. Everything left over is the measure.
 */
function readPeriod(
  tokens: string[]
): { period: Omit<PeriodHeader, "source" | "measure">; used: number } | null {
  const lower = tokens.map((token) => token.toLowerCase());

  // "2024-01", "2024 1"  (year then month number)
  if (lower.length >= 2 && /^(19|20)\d{2}$/.test(lower[0]) && /^\d{1,2}$/.test(lower[1])) {
    const month = Number(lower[1]);
    if (month >= 1 && month <= 12) {
      const year = Number(lower[0]);
      return {
        period: monthPeriod(year, month - 1, false),
        used: 2,
      };
    }
  }

  // "Q1", "Q1 2024"
  const quarter = /^q([1-4])$/.exec(lower[0]);
  if (quarter) {
    const number = Number(quarter[1]);
    const hasYear = lower.length >= 2 && /^\d{2,4}$/.test(lower[1]);
    const year = hasYear ? fullYear(lower[1]) : new Date().getUTCFullYear();
    return {
      period: {
        label: hasYear ? `Q${number} ${year}` : `Q${number}`,
        sortKey: year * 12 + (number - 1) * 3,
        date: hasYear ? new Date(Date.UTC(year, (number - 1) * 3, 1)) : null,
        grain: "quarter",
        yearAssumed: !hasYear,
      },
      used: hasYear ? 2 : 1,
    };
  }

  // "2024 Q1"
  if (lower.length >= 2 && /^(19|20)\d{2}$/.test(lower[0]) && /^q[1-4]$/.test(lower[1])) {
    const year = Number(lower[0]);
    const number = Number(lower[1].slice(1));
    return {
      period: {
        label: `Q${number} ${year}`,
        sortKey: year * 12 + (number - 1) * 3,
        date: new Date(Date.UTC(year, (number - 1) * 3, 1)),
        grain: "quarter",
        yearAssumed: false,
      },
      used: 2,
    };
  }

  // "Jan", "Jan 24", "January 2024"
  const monthIndex = MONTHS[lower[0]];
  if (monthIndex !== undefined) {
    const hasYear = lower.length >= 2 && /^\d{2,4}$/.test(lower[1]) && lower[1].length !== 1;
    const year = hasYear ? fullYear(lower[1]) : new Date().getUTCFullYear();
    return {
      period: monthPeriod(year, monthIndex, !hasYear),
      used: hasYear ? 2 : 1,
    };
  }

  // A bare year. Checked last so "2024 Q1" and "2024-01" win first.
  if (/^(19|20)\d{2}$/.test(lower[0])) {
    const year = Number(lower[0]);
    return {
      period: {
        label: String(year),
        sortKey: year * 12,
        date: new Date(Date.UTC(year, 0, 1)),
        grain: "year",
        yearAssumed: false,
      },
      used: 1,
    };
  }

  return null;
}

function monthPeriod(year: number, monthIndex: number, yearAssumed: boolean) {
  return {
    label: yearAssumed ? MONTH_LABELS[monthIndex] : `${MONTH_LABELS[monthIndex]} ${year}`,
    sortKey: year * 12 + monthIndex,
    date: yearAssumed ? null : new Date(Date.UTC(year, monthIndex, 1)),
    grain: "month" as PeriodGrain,
    yearAssumed,
  };
}

/** Words that appear alongside a period but don't name a measure. */
const NOISE = /^(total|actual|actuals|amount|value|values|sales|fy|cy|month|period|qtr|quarter)$/i;

/**
 * Reads a column heading as a period, with any leftover words treated as the
 * measure. Returns null when the heading is not a period at all, which is how
 * the id columns on the left are told apart from the wide block on the right.
 */
export function parsePeriodHeader(header: unknown): PeriodHeader | null {
  if (typeof header !== "string" && typeof header !== "number") return null;
  const source = String(header).trim();
  if (source === "") return null;

  const tokens = words(source);
  if (tokens.length === 0) return null;

  // Try the period at the front ("Jan Units"), then at the back ("Units Jan").
  let found = readPeriod(tokens);
  let leftover: string[] = [];
  if (found) {
    leftover = tokens.slice(found.used);
  } else {
    for (let start = 1; start < tokens.length; start += 1) {
      const attempt = readPeriod(tokens.slice(start));
      if (attempt && attempt.used === tokens.length - start) {
        found = attempt;
        leftover = tokens.slice(0, start);
        break;
      }
    }
  }
  if (!found) return null;

  const measureWords = leftover.filter((word) => !NOISE.test(word));

  // A measure is a word or two next to the period ("Jan Units", "Revenue Jan
  // 2024"). Anything longer is a sentence that happens to contain a month, and
  // "May" is the one month that is also an ordinary English word - so a bare
  // "May" followed by other words is prose, not a heading.
  if (measureWords.length > 2) return null;
  const monthWord = found.period.grain === "month" && found.period.yearAssumed;
  if (monthWord && measureWords.length > 0 && /^may$/i.test(tokens[0])) return null;

  return {
    ...found.period,
    source,
    measure: measureWords.length > 0 ? measureWords.join(" ") : null,
  };
}

// ---------------------------------------------------------------------------
// Spotting the crosstab
// ---------------------------------------------------------------------------

export interface CrosstabShape {
  /** Column indexes that identify the row: everything left of the wide block. */
  idColumns: number[];
  /** Column indexes of the wide block, in sheet order. */
  valueColumns: number[];
  /** Parsed heading for each value column, parallel to valueColumns. */
  periods: PeriodHeader[];
  /** Distinct measures found in the headings, in first-seen order. */
  measures: string[];
  grain: PeriodGrain;
  /** What to call the new period column: "Month", "Quarter" or "Year". */
  periodColumnName: string;
  /** True when every period knows its year, so real dates can be written. */
  datesKnown: boolean;
}

/** The minimum number of period columns before a sheet counts as a crosstab. */
const MIN_VALUE_COLUMNS = 2;

/** How much of a wide column has to be numeric before we believe it. */
const NUMERIC_SHARE = 0.5;

function isNumericColumn(rows: Grid, index: number): boolean {
  let present = 0;
  let numeric = 0;
  for (const row of rows) {
    const value = row[index];
    if (isBlankCell(value)) continue;
    present += 1;
    if (typeof value === "number" || parseLooseNumber(value) !== null) numeric += 1;
  }
  // A column that is entirely empty is still fine to unpivot - it just
  // contributes nothing - so only a column with contradicting content is
  // rejected.
  return present === 0 || numeric / present >= NUMERIC_SHARE;
}

function grainOf(periods: readonly PeriodHeader[]): PeriodGrain {
  if (periods.some((period) => period.grain === "month")) return "month";
  if (periods.some((period) => period.grain === "quarter")) return "quarter";
  return "year";
}

const GRAIN_NAMES: Record<PeriodGrain, string> = {
  month: "Month",
  quarter: "Quarter",
  year: "Year",
};

/**
 * Looks for a run of period headings at the right-hand end of the table. The
 * run has to reach the last column: a sheet whose period columns are followed
 * by a Total column is handled, because totals are dropped first by the caller.
 */
export function detectCrosstab(table: HeaderTable): CrosstabShape | null {
  const { headers, rows } = table;
  if (headers.length < MIN_VALUE_COLUMNS + 1) return null;

  const parsed = headers.map((header) => parsePeriodHeader(header));

  // Walk in from the right while the headings keep parsing as periods.
  let first = headers.length;
  while (first - 1 >= 0 && parsed[first - 1] !== null && isNumericColumn(rows, first - 1)) {
    first -= 1;
  }

  const valueColumns: number[] = [];
  for (let index = first; index < headers.length; index += 1) valueColumns.push(index);
  if (valueColumns.length < MIN_VALUE_COLUMNS) return null;

  // Something has to be left over to identify the row. A sheet that is periods
  // all the way across has no keys, so there is nothing to unpivot against.
  const idColumns: number[] = [];
  for (let index = 0; index < first; index += 1) idColumns.push(index);
  if (idColumns.length === 0) return null;

  const periods = valueColumns.map((index) => parsed[index] as PeriodHeader);

  const measures: string[] = [];
  for (const period of periods) {
    if (period.measure && !measures.includes(period.measure)) measures.push(period.measure);
  }

  // Headings that mix "Jan Units" with a bare "Feb" are ambiguous: we would not
  // know what the bare column measures. Treat that as not a crosstab rather
  // than guess.
  if (measures.length > 0 && periods.some((period) => period.measure === null)) return null;

  const grain = grainOf(periods);
  return {
    idColumns,
    valueColumns,
    periods,
    measures,
    grain,
    periodColumnName: GRAIN_NAMES[grain],
    datesKnown: periods.every((period) => period.date !== null),
  };
}

/** Forward-fills blanks, which is what a merged header row reads back as. */
export function fillForward(row: readonly CellValue[]): CellValue[] {
  const out: CellValue[] = [];
  let last: CellValue = null;
  for (const value of row) {
    if (!isBlankCell(value)) last = value;
    out.push(isBlankCell(value) ? last : value);
  }
  return out;
}

/**
 * Joins a two-row heading into one. Crosstabs often carry the year on one row,
 * merged across the months, and the month on the next: "2024" over "Jan"
 * becomes "Jan 2024".
 */
export function mergeHeaderRows(top: readonly CellValue[], bottom: readonly CellValue[]): string[] {
  const filled = fillForward(top);
  const width = Math.max(filled.length, bottom.length);
  const out: string[] = [];
  for (let index = 0; index < width; index += 1) {
    const above = filled[index];
    const below = bottom[index];
    const aboveText = isBlankCell(above) ? "" : String(above).trim();
    const belowText = isBlankCell(below) ? "" : String(below).trim();
    if (
      aboveText === "" ||
      belowText === "" ||
      normalizeHeader(aboveText) === normalizeHeader(belowText)
    ) {
      out.push(belowText !== "" ? belowText : aboveText);
    } else {
      // "Jan" under "2024" reads better as "Jan 2024" than "2024 Jan", and
      // parsePeriodHeader understands both.
      out.push(`${belowText} ${aboveText}`);
    }
  }
  return out;
}

export interface DetectedCrosstab {
  shape: CrosstabShape;
  /** Headings used, after any two-row merge. */
  headers: string[];
  rows: Grid;
  /** How many rows at the top of the grid were headings. */
  headerRows: number;
}

/**
 * Finds a crosstab in a raw grid, trying a single heading row first and then a
 * merged pair. The caller passes the grid as read from the sheet.
 */
export function detectCrosstabInGrid(grid: Grid): DetectedCrosstab | null {
  if (grid.length < 2) return null;

  const single: string[] = grid[0].map((value) => (isBlankCell(value) ? "" : String(value).trim()));
  const singleShape = detectCrosstab({ headers: single, rows: grid.slice(1) });
  if (singleShape) {
    return { shape: singleShape, headers: single, rows: grid.slice(1), headerRows: 1 };
  }

  if (grid.length < 3) return null;
  const merged = mergeHeaderRows(grid[0], grid[1]);
  const mergedShape = detectCrosstab({ headers: merged, rows: grid.slice(2) });
  if (mergedShape) {
    return { shape: mergedShape, headers: merged, rows: grid.slice(2), headerRows: 2 };
  }
  return null;
}

// ---------------------------------------------------------------------------
// Unpivoting
// ---------------------------------------------------------------------------

export interface UnpivotOptions {
  /** Keep rows whose value cell is empty. Off by default: they add nothing. */
  keepBlanks?: boolean;
  /** Name for the single value column. Ignored when the headings name measures. */
  valueColumnName?: string;
  /** Write the period as a real date when every heading knows its year. */
  periodAsDate?: boolean;
  /** Sort the output by period. Off by default, so sheet order is preserved. */
  sortByPeriod?: boolean;
}

export interface UnpivotResult {
  table: HeaderTable;
  /** Number format per output column, for the writer. Undefined means General. */
  formats: (string | undefined)[];
  rowsIn: number;
  rowsOut: number;
  /** Cells left out because they were empty. */
  skippedBlank: number;
}

const DATE_FORMATS: Record<PeriodGrain, string> = {
  month: "mmm yyyy",
  quarter: "mmm yyyy",
  year: "yyyy",
};

/**
 * Turns the wide block into rows. One output row per (source row, period),
 * carrying the id columns across unchanged.
 */
export function unpivot(
  table: HeaderTable,
  shape: CrosstabShape,
  options: UnpivotOptions = {}
): UnpivotResult {
  const { keepBlanks = false, periodAsDate = true, sortByPeriod = false } = options;
  const asDate = periodAsDate && shape.datesKnown;

  const headers = shape.idColumns.map((index) => table.headers[index] || `Column ${index + 1}`);
  headers.push(shape.periodColumnName);
  const measured = shape.measures.length > 0;
  if (measured) {
    headers.push(...shape.measures);
  } else {
    headers.push(options.valueColumnName ?? "Value");
  }

  // Periods in the order their columns appear, de-duplicated: with measures,
  // several columns share one period.
  const periodOrder: PeriodHeader[] = [];
  for (const period of shape.periods) {
    if (!periodOrder.some((seen) => seen.label === period.label)) periodOrder.push(period);
  }
  if (sortByPeriod) periodOrder.sort((a, b) => a.sortKey - b.sortKey);

  const rows: Grid = [];
  let skippedBlank = 0;

  for (const sourceRow of table.rows) {
    for (const period of periodOrder) {
      const ids = shape.idColumns.map((index) => sourceRow[index] ?? null);
      const periodCell: CellValue =
        asDate && period.date ? dateToExcelSerial(period.date) : period.label;

      if (measured) {
        const cells = shape.measures.map((measure) => {
          const at = shape.periods.findIndex(
            (candidate) => candidate.label === period.label && candidate.measure === measure
          );
          return at >= 0 ? (sourceRow[shape.valueColumns[at]] ?? null) : null;
        });
        const allBlank = cells.every((cell) => isBlankCell(cell));
        if (allBlank) {
          skippedBlank += 1;
          if (!keepBlanks) continue;
        }
        rows.push([...ids, periodCell, ...cells]);
      } else {
        const at = shape.periods.findIndex((candidate) => candidate.label === period.label);
        const value = at >= 0 ? (sourceRow[shape.valueColumns[at]] ?? null) : null;
        if (isBlankCell(value)) {
          skippedBlank += 1;
          if (!keepBlanks) continue;
        }
        rows.push([...ids, periodCell, value]);
      }
    }
  }

  const formats: (string | undefined)[] = headers.map(() => undefined);
  if (asDate) formats[shape.idColumns.length] = DATE_FORMATS[shape.grain];

  return {
    table: { headers, rows },
    formats,
    rowsIn: table.rows.length,
    rowsOut: rows.length,
    skippedBlank,
  };
}

/** One line describing what was found, for the assistant to say out loud. */
export function describeCrosstab(shape: CrosstabShape): string {
  const span =
    shape.periods.length > 0
      ? `${shape.periods[0].label} to ${shape.periods[shape.periods.length - 1].label}`
      : "";
  const idNames = shape.idColumns.length;
  const measures = shape.measures.length > 0 ? `, each with ${shape.measures.join(" and ")}` : "";
  return `${shape.periods.length} ${shape.grain} columns (${span})${measures}, keyed by ${idNames} column${idNames === 1 ? "" : "s"}`;
}
