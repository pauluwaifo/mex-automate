/**
 * Comparing two versions of the same table.
 *
 * "Last month's file and this month's file disagree by 412 and I can't find
 * where" is one of the oldest jobs in a spreadsheet, and the usual approach - a
 * screenful of VLOOKUPs, or scrolling with two windows side by side - finds the
 * difference eventually and tells you nothing about how many others there are.
 *
 * This compares them properly: rows are matched on a key column, and then every
 * cell that differs is reported with both values and the amount it moved by.
 * Rows that only exist on one side are listed separately, because a missing row
 * and a changed row need different explanations.
 *
 * Numbers are compared with a tolerance, since a value that arrives as
 * 1234.5600000000001 after a division is not a real difference, and neither is
 * a figure rounded to the penny in one export and not the other.
 *
 * Everything here is pure and unit tested in tests/compare.test.ts.
 */

import { isBlankCell, normalizeHeader, plural } from "../shared/excelHelpers";
import { parseLooseNumber } from "./tidy";
import type { CellValue, Grid, HeaderTable } from "../shared/types";

// ---------------------------------------------------------------------------
// Options and results
// ---------------------------------------------------------------------------

export interface CompareOptions {
  /**
   * Header to match rows on. When null, rows are matched by their whole
   * contents, which finds added and removed rows but cannot report changes.
   */
  key?: string | null;
  /**
   * How far apart two numbers may be and still count as equal. Defaults to half
   * a penny, which absorbs rounding without hiding real movement.
   */
  tolerance?: number;
  /** Treat "ACME" and "Acme" as the same value. On by default. */
  ignoreCase?: boolean;
  /** Only compare these headers. Defaults to every header both sides share. */
  columns?: readonly string[];
}

export interface CellChange {
  key: string;
  header: string;
  before: CellValue;
  after: CellValue;
  /** 1-based row in the first sheet's body, for going to the cell. */
  rowBefore: number;
  rowAfter: number;
  /** after - before when both are numbers, otherwise null. */
  delta: number | null;
}

export interface RowRef {
  key: string;
  /** 1-based row in that sheet's body. */
  row: number;
  /** The row's values, so the report can show what was added or removed. */
  values: CellValue[];
}

export interface Diff {
  keyColumn: string | null;
  comparedColumns: string[];
  columnsOnlyInFirst: string[];
  columnsOnlyInSecond: string[];
  /** Rows in the second sheet with no match in the first. */
  added: RowRef[];
  /** Rows in the first sheet with no match in the second. */
  removed: RowRef[];
  changed: CellChange[];
  /** Keys that appear more than once, where only the first was compared. */
  duplicateKeysFirst: string[];
  duplicateKeysSecond: string[];
  matchedRows: number;
  /** Matched rows in which nothing differed. */
  identicalRows: number;
}

export const DEFAULT_TOLERANCE = 0.005;

// ---------------------------------------------------------------------------
// Choosing the column to match on
// ---------------------------------------------------------------------------

const KEY_HINT =
  /(^|[\s_#.])(id|code|no|nr|number|num|ref|reference|sku|key|account|invoice|order)($|[\s_#.])/i;

function keyText(value: CellValue, ignoreCase: boolean): string {
  if (isBlankCell(value)) return "";
  const text = typeof value === "string" ? value.trim() : String(value);
  return ignoreCase ? text.toLowerCase() : text;
}

/**
 * How many rows a table needs before every value being distinct counts as
 * evidence of a key. In a two-row sheet the quantity column is "unique" too.
 */
const MIN_ROWS_FOR_UNHINTED_KEY = 4;

/** Share of non-blank values in a column that appear exactly once. */
function uniqueness(
  rows: Grid,
  index: number
): { unique: number; filled: number; allNumbers: boolean } {
  const seen = new Map<string, number>();
  let filled = 0;
  let numbers = 0;
  for (const row of rows) {
    const value = row[index];
    if (isBlankCell(value)) continue;
    filled += 1;
    if (parseLooseNumber(value) !== null) numbers += 1;
    const text = keyText(value, true);
    seen.set(text, (seen.get(text) ?? 0) + 1);
  }
  let unique = 0;
  for (const count of seen.values()) if (count === 1) unique += 1;
  return {
    unique: filled === 0 ? 0 : unique / filled,
    filled,
    allNumbers: filled > 0 && numbers === filled,
  };
}

/**
 * Picks the column to match rows on: one both sheets have, that is filled in and
 * near enough unique to identify a row. A column whose name sounds like an
 * identifier wins ties, and otherwise the leftmost does.
 */
export function chooseKeyColumn(first: HeaderTable, second: HeaderTable): string | null {
  const shared = first.headers.filter((header) =>
    second.headers.some((other) => normalizeHeader(other) === normalizeHeader(header))
  );

  const scored = shared
    .map((header) => {
      const a = uniqueness(first.rows, first.headers.indexOf(header));
      const indexInSecond = second.headers.findIndex(
        (other) => normalizeHeader(other) === normalizeHeader(header)
      );
      const b = uniqueness(second.rows, indexInSecond);
      const filledEnough =
        a.filled >= Math.max(1, first.rows.length * 0.9) &&
        b.filled >= Math.max(1, second.rows.length * 0.9);
      const hinted = KEY_HINT.test(header);
      // A column of plain numbers is far more likely to be a quantity than an
      // identifier, unless its name says otherwise ("Invoice No"). And in a very
      // short table, distinctness on its own proves nothing.
      const plausible =
        hinted ||
        (!a.allNumbers &&
          !b.allNumbers &&
          first.rows.length >= MIN_ROWS_FOR_UNHINTED_KEY &&
          second.rows.length >= MIN_ROWS_FOR_UNHINTED_KEY);
      return {
        header,
        score: Math.min(a.unique, b.unique),
        filledEnough,
        plausible,
        hinted,
        at: first.headers.indexOf(header),
      };
    })
    .filter((candidate) => candidate.filledEnough && candidate.plausible && candidate.score >= 0.98)
    .sort((a, b) => {
      if (a.hinted !== b.hinted) return a.hinted ? -1 : 1;
      if (b.score !== a.score) return b.score - a.score;
      return a.at - b.at;
    });

  return scored[0]?.header ?? null;
}

// ---------------------------------------------------------------------------
// Comparing values
// ---------------------------------------------------------------------------

/**
 * True when two cells say the same thing. Numbers are compared numerically, so
 * 1000 and "1,000.00" match; text is compared trimmed.
 */
export function cellsMatch(
  before: CellValue,
  after: CellValue,
  tolerance = DEFAULT_TOLERANCE,
  ignoreCase = true
): boolean {
  const beforeBlank = isBlankCell(before);
  const afterBlank = isBlankCell(after);
  if (beforeBlank || afterBlank) return beforeBlank && afterBlank;

  const a = parseLooseNumber(before);
  const b = parseLooseNumber(after);
  if (a && b) return Math.abs(a.value - b.value) <= tolerance;

  // One side numeric and the other not is a real difference, even if the text
  // happens to look similar: 12 and "twelve" are not the same cell.
  if (a || b) return false;

  const left = String(before).trim();
  const right = String(after).trim();
  return ignoreCase ? left.toLowerCase() === right.toLowerCase() : left === right;
}

function numericDelta(before: CellValue, after: CellValue): number | null {
  const a = parseLooseNumber(before);
  const b = parseLooseNumber(after);
  return a && b ? b.value - a.value : null;
}

// ---------------------------------------------------------------------------
// The comparison
// ---------------------------------------------------------------------------

interface Indexed {
  /** First row for each key, 0-based within the body. */
  rows: Map<string, number>;
  duplicates: string[];
  blanks: number;
}

function indexByKey(rows: Grid, keyIndex: number, ignoreCase: boolean): Indexed {
  const index = new Map<string, number>();
  const duplicates: string[] = [];
  let blanks = 0;
  rows.forEach((row, at) => {
    const text = keyText(row[keyIndex], ignoreCase);
    if (text === "") {
      blanks += 1;
      return;
    }
    if (index.has(text)) {
      if (!duplicates.includes(text)) duplicates.push(text);
      return;
    }
    index.set(text, at);
  });
  return { rows: index, duplicates, blanks };
}

/** A whole row reduced to one string, for matching when there is no key. */
function rowSignature(row: readonly CellValue[], ignoreCase: boolean): string {
  return row.map((value) => keyText(value, ignoreCase)).join("\u0001");
}

function compareWithoutKey(first: HeaderTable, second: HeaderTable, ignoreCase: boolean): Diff {
  const counts = new Map<string, number>();
  first.rows.forEach((row) => {
    const signature = rowSignature(row, ignoreCase);
    counts.set(signature, (counts.get(signature) ?? 0) + 1);
  });

  const added: RowRef[] = [];
  second.rows.forEach((row, at) => {
    const signature = rowSignature(row, ignoreCase);
    const remaining = counts.get(signature) ?? 0;
    if (remaining > 0) counts.set(signature, remaining - 1);
    else added.push({ key: `Row ${at + 1}`, row: at + 1, values: [...row] });
  });

  const leftover = new Map(counts);
  const removed: RowRef[] = [];
  first.rows.forEach((row, at) => {
    const signature = rowSignature(row, ignoreCase);
    const remaining = leftover.get(signature) ?? 0;
    if (remaining > 0) {
      leftover.set(signature, remaining - 1);
      removed.push({ key: `Row ${at + 1}`, row: at + 1, values: [...row] });
    }
  });

  return {
    keyColumn: null,
    comparedColumns: [],
    columnsOnlyInFirst: [],
    columnsOnlyInSecond: [],
    added,
    removed,
    changed: [],
    duplicateKeysFirst: [],
    duplicateKeysSecond: [],
    matchedRows: first.rows.length - removed.length,
    identicalRows: first.rows.length - removed.length,
  };
}

/**
 * Compares two tables. `first` is the older or expected version and `second` the
 * newer one, so "added" means present in the second only.
 */
export function compareTables(
  first: HeaderTable,
  second: HeaderTable,
  options: CompareOptions = {}
): Diff {
  const { tolerance = DEFAULT_TOLERANCE, ignoreCase = true } = options;
  const key = options.key === undefined ? chooseKeyColumn(first, second) : options.key;

  if (!key) return compareWithoutKey(first, second, ignoreCase);

  const keyInFirst = first.headers.findIndex(
    (header) => normalizeHeader(header) === normalizeHeader(key)
  );
  const keyInSecond = second.headers.findIndex(
    (header) => normalizeHeader(header) === normalizeHeader(key)
  );
  if (keyInFirst < 0 || keyInSecond < 0) return compareWithoutKey(first, second, ignoreCase);

  const shared = first.headers.filter((header) =>
    second.headers.some((other) => normalizeHeader(other) === normalizeHeader(header))
  );
  const requested = options.columns
    ? shared.filter((header) =>
        options.columns!.some((wanted) => normalizeHeader(wanted) === normalizeHeader(header))
      )
    : shared;
  // The key itself carries no information about change: matched rows share it.
  const comparedColumns = requested.filter(
    (header) => normalizeHeader(header) !== normalizeHeader(key)
  );

  const columnsOnlyInFirst = first.headers.filter(
    (header) => !second.headers.some((other) => normalizeHeader(other) === normalizeHeader(header))
  );
  const columnsOnlyInSecond = second.headers.filter(
    (header) => !first.headers.some((other) => normalizeHeader(other) === normalizeHeader(header))
  );

  const indexFirst = indexByKey(first.rows, keyInFirst, ignoreCase);
  const indexSecond = indexByKey(second.rows, keyInSecond, ignoreCase);

  const changed: CellChange[] = [];
  const removed: RowRef[] = [];
  const added: RowRef[] = [];
  let matchedRows = 0;
  let identicalRows = 0;

  for (const [text, rowInFirst] of indexFirst.rows) {
    const rowInSecond = indexSecond.rows.get(text);
    const row = first.rows[rowInFirst];
    if (rowInSecond === undefined) {
      removed.push({
        key: String(row[keyInFirst] ?? text),
        row: rowInFirst + 1,
        values: [...row],
      });
      continue;
    }

    matchedRows += 1;
    const other = second.rows[rowInSecond];
    let rowChanged = false;
    for (const header of comparedColumns) {
      const before = row[first.headers.indexOf(header)] ?? null;
      const at = second.headers.findIndex(
        (candidate) => normalizeHeader(candidate) === normalizeHeader(header)
      );
      const after = other[at] ?? null;
      if (cellsMatch(before, after, tolerance, ignoreCase)) continue;
      rowChanged = true;
      changed.push({
        key: String(row[keyInFirst] ?? text),
        header,
        before,
        after,
        rowBefore: rowInFirst + 1,
        rowAfter: rowInSecond + 1,
        delta: numericDelta(before, after),
      });
    }
    if (!rowChanged) identicalRows += 1;
  }

  for (const [text, rowInSecond] of indexSecond.rows) {
    if (indexFirst.rows.has(text)) continue;
    const row = second.rows[rowInSecond];
    added.push({ key: String(row[keyInSecond] ?? text), row: rowInSecond + 1, values: [...row] });
  }

  return {
    keyColumn: key,
    comparedColumns,
    columnsOnlyInFirst,
    columnsOnlyInSecond,
    added,
    removed,
    changed,
    duplicateKeysFirst: indexFirst.duplicates,
    duplicateKeysSecond: indexSecond.duplicates,
    matchedRows,
    identicalRows,
  };
}

// ---------------------------------------------------------------------------
// Saying what changed
// ---------------------------------------------------------------------------

/** One sentence for the chat. */
export function summarizeDiff(diff: Diff): string {
  if (diff.changed.length === 0 && diff.added.length === 0 && diff.removed.length === 0) {
    return "These two are identical.";
  }
  const parts: string[] = [];
  if (diff.changed.length > 0) {
    const rows = new Set(diff.changed.map((change) => change.key)).size;
    parts.push(`${plural(diff.changed.length, "changed cell")} across ${plural(rows, "row")}`);
  }
  if (diff.added.length > 0) parts.push(`${plural(diff.added.length, "new row")}`);
  if (diff.removed.length > 0)
    parts.push(`${diff.removed.length} row${diff.removed.length === 1 ? "" : "s"} gone`);
  return `Found ${parts.join(", ")}.`;
}

/** The biggest numeric movements first, for the summary at the top of the report. */
export function largestMovements(diff: Diff, limit = 5): CellChange[] {
  return diff.changed
    .filter((change) => change.delta !== null)
    .sort((a, b) => Math.abs(b.delta as number) - Math.abs(a.delta as number))
    .slice(0, limit);
}

/** Total movement per column, so "Revenue is up 412 overall" can be stated. */
export function movementByColumn(
  diff: Diff
): Array<{ header: string; total: number; cells: number }> {
  const totals = new Map<string, { total: number; cells: number }>();
  for (const change of diff.changed) {
    if (change.delta === null) continue;
    const entry = totals.get(change.header) ?? { total: 0, cells: 0 };
    entry.total += change.delta;
    entry.cells += 1;
    totals.set(change.header, entry);
  }
  return [...totals.entries()]
    .map(([header, entry]) => ({ header, ...entry }))
    .sort((a, b) => Math.abs(b.total) - Math.abs(a.total));
}

export interface ReportNames {
  first: string;
  second: string;
}

/**
 * Lays the whole comparison out as a grid, ready to be written to a new sheet.
 * Section headings are returned too, so the writer can make them bold without
 * having to guess which rows they are.
 */
export function diffReportGrid(
  diff: Diff,
  names: ReportNames
): { grid: Grid; headingRows: number[]; tableHeaderRows: number[] } {
  const grid: Grid = [];
  const headingRows: number[] = [];
  const tableHeaderRows: number[] = [];

  const heading = (text: string) => {
    headingRows.push(grid.length);
    grid.push([text]);
  };
  const tableHeader = (cells: CellValue[]) => {
    tableHeaderRows.push(grid.length);
    grid.push(cells);
  };
  const blank = () => grid.push([]);

  heading("Comparison");
  grid.push(["Older sheet", names.first]);
  grid.push(["Newer sheet", names.second]);
  grid.push(["Matched on", diff.keyColumn ?? "whole rows (no shared key column)"]);
  grid.push(["Rows matched", diff.matchedRows]);
  grid.push(["Rows identical", diff.identicalRows]);
  grid.push(["Cells changed", diff.changed.length]);
  grid.push(["Rows only in the newer sheet", diff.added.length]);
  grid.push(["Rows only in the older sheet", diff.removed.length]);

  const movements = movementByColumn(diff);
  if (movements.length > 0) {
    blank();
    heading("Net movement");
    tableHeader(["Column", "Change", "Cells"]);
    for (const movement of movements) {
      grid.push([movement.header, movement.total, movement.cells]);
    }
  }

  if (diff.changed.length > 0) {
    blank();
    heading("Changed cells");
    tableHeader([diff.keyColumn ?? "Row", "Column", names.first, names.second, "Change", "Row"]);
    for (const change of diff.changed) {
      grid.push([
        change.key,
        change.header,
        change.before,
        change.after,
        change.delta,
        change.rowAfter,
      ]);
    }
  }

  if (diff.added.length > 0) {
    blank();
    heading(`Only in ${names.second}`);
    tableHeader([diff.keyColumn ?? "Row", "Row"]);
    for (const row of diff.added) grid.push([row.key, row.row]);
  }

  if (diff.removed.length > 0) {
    blank();
    heading(`Only in ${names.first}`);
    tableHeader([diff.keyColumn ?? "Row", "Row"]);
    for (const row of diff.removed) grid.push([row.key, row.row]);
  }

  const notes: string[] = [];
  if (diff.columnsOnlyInFirst.length > 0) {
    notes.push(`Columns only in ${names.first}: ${diff.columnsOnlyInFirst.join(", ")}`);
  }
  if (diff.columnsOnlyInSecond.length > 0) {
    notes.push(`Columns only in ${names.second}: ${diff.columnsOnlyInSecond.join(", ")}`);
  }
  if (diff.duplicateKeysFirst.length > 0) {
    notes.push(
      `Repeated keys in ${names.first} (only the first row of each was compared): ${diff.duplicateKeysFirst.slice(0, 10).join(", ")}`
    );
  }
  if (diff.duplicateKeysSecond.length > 0) {
    notes.push(
      `Repeated keys in ${names.second} (only the first row of each was compared): ${diff.duplicateKeysSecond.slice(0, 10).join(", ")}`
    );
  }
  if (notes.length > 0) {
    blank();
    heading("Worth knowing");
    for (const note of notes) grid.push([note]);
  }

  // Every row has to be the same width for Range.values.
  const width = grid.reduce((widest, row) => Math.max(widest, row.length), 1);
  const rectangular = grid.map((row) => {
    const padded = [...row];
    while (padded.length < width) padded.push(null);
    return padded;
  });

  return { grid: rectangular, headingRows, tableHeaderRows };
}
