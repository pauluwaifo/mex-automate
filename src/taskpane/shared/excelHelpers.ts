/**
 * Reusable Office.js range/table/sheet utilities plus the pure grid helpers the
 * feature modules build on. Everything above the "Office.js" banner is free of
 * Office globals so it can be unit tested under plain Node.
 */

import { CellValue, Grid, HeaderTable, OperationResult, ParsedAddress } from "./types";

// ---------------------------------------------------------------------------
// Pure helpers
// ---------------------------------------------------------------------------

/** 0 -> "A", 25 -> "Z", 26 -> "AA". */
export function columnLetter(index: number): string {
  if (!Number.isInteger(index) || index < 0) {
    throw new RangeError(`Column index must be a non-negative integer, got ${index}`);
  }
  let result = "";
  let n = index;
  while (n >= 0) {
    result = String.fromCharCode((n % 26) + 65) + result;
    n = Math.floor(n / 26) - 1;
  }
  return result;
}

/** "A" -> 0, "AA" -> 26. Case insensitive. */
export function columnIndexFromLetter(letter: string): number {
  const upper = letter.toUpperCase();
  if (!/^[A-Z]+$/.test(upper)) {
    throw new RangeError(`Not a column letter: ${letter}`);
  }
  let index = 0;
  for (const char of upper) {
    index = index * 26 + (char.charCodeAt(0) - 64);
  }
  return index - 1;
}

/**
 * Fold a header into a comparable key so "Order Date", "order_date" and
 * " ORDER  DATE " all match each other.
 */
export function normalizeHeader(header: unknown): string {
  return String(header ?? "")
    .replace(/[\s_\-.]+/g, " ")
    .trim()
    .toLowerCase();
}

export function isBlankCell(value: CellValue): boolean {
  return (
    value === null || value === undefined || (typeof value === "string" && value.trim() === "")
  );
}

export function isBlankRow(row: readonly CellValue[]): boolean {
  return row.every(isBlankCell);
}

/** Drop trailing all-blank rows and all-blank columns from a grid. */
export function trimTrailingBlanks(grid: Grid): Grid {
  let lastRow = grid.length - 1;
  while (lastRow >= 0 && isBlankRow(grid[lastRow])) {
    lastRow -= 1;
  }
  const rows = grid.slice(0, lastRow + 1);
  if (rows.length === 0) {
    return [];
  }
  const width = Math.max(...rows.map((row) => row.length));
  let lastColumn = width - 1;
  while (lastColumn >= 0 && rows.every((row) => isBlankCell(row[lastColumn] ?? null))) {
    lastColumn -= 1;
  }
  if (lastColumn < 0) {
    return [];
  }
  return rows.map((row) => padRow(row, lastColumn + 1));
}

/** Pad (or truncate) a row to an exact width so Excel accepts it as a rectangle. */
export function padRow(row: readonly CellValue[], width: number): CellValue[] {
  const out: CellValue[] = new Array(width);
  for (let i = 0; i < width; i += 1) {
    out[i] = i < row.length ? (row[i] ?? null) : null;
  }
  return out;
}

/** Force a ragged grid into a rectangle wide enough for its widest row. */
export function rectangularize(grid: Grid): Grid {
  if (grid.length === 0) {
    return [];
  }
  const width = Math.max(...grid.map((row) => row.length));
  return grid.map((row) => padRow(row, width));
}

/** Split a grid into a header row plus body rows. */
export function toHeaderTable(grid: Grid): HeaderTable {
  if (grid.length === 0) {
    return { headers: [], rows: [] };
  }
  const headers = grid[0].map((cell, index) => {
    const label = String(cell ?? "").trim();
    return label === "" ? `Column ${columnLetter(index)}` : label;
  });
  return { headers, rows: grid.slice(1).map((row) => padRow(row, headers.length)) };
}

/** Make `name` unique against `taken` by appending " (2)", " (3)", ... */
export function uniqueName(name: string, taken: Iterable<string>): string {
  const used = new Set(Array.from(taken, (item) => item.toLowerCase()));
  if (!used.has(name.toLowerCase())) {
    return name;
  }
  let suffix = 2;
  while (used.has(`${name} (${suffix})`.toLowerCase())) {
    suffix += 1;
  }
  return `${name} (${suffix})`;
}

/** Split an array into fixed-size chunks. */
export function chunk<T>(items: readonly T[], size: number): T[][] {
  if (size <= 0) {
    throw new RangeError("Chunk size must be greater than zero");
  }
  const out: T[][] = [];
  for (let i = 0; i < items.length; i += size) {
    out.push(items.slice(i, i + size));
  }
  return out;
}

/** Quote a sheet name for use in an A1 address when it needs quoting. */
export function quoteSheetName(sheetName: string): string {
  return /^[A-Za-z0-9_]+$/.test(sheetName) ? sheetName : `'${sheetName.replace(/'/g, "''")}'`;
}

/** Build an A1 address from zero-based offsets: buildAddress("Data", 0, 0, 3, 2) -> "Data!A1:B3". */
export function buildAddress(
  sheetName: string | null,
  rowIndex: number,
  columnIndex: number,
  rowCount: number,
  columnCount: number
): string {
  const start = `${columnLetter(columnIndex)}${rowIndex + 1}`;
  const end = `${columnLetter(columnIndex + columnCount - 1)}${rowIndex + rowCount}`;
  const local = rowCount === 1 && columnCount === 1 ? start : `${start}:${end}`;
  return sheetName ? `${quoteSheetName(sheetName)}!${local}` : local;
}

const ADDRESS_PATTERN =
  /^(?:(?:'((?:[^']|'')+)'|([^'!]+))!)?\$?([A-Za-z]+)\$?(\d+)(?::\$?([A-Za-z]+)\$?(\d+))?$/;

/** Parse `Sheet1!$B$2:$D$20` (and simpler forms) into zero-based offsets. */
export function parseAddress(address: string): ParsedAddress {
  const match = ADDRESS_PATTERN.exec(address.trim());
  if (!match) {
    throw new Error(`Could not parse range address: ${address}`);
  }
  const [, quotedSheet, plainSheet, startCol, startRow, endCol, endRow] = match;
  const sheetName = quotedSheet ? quotedSheet.replace(/''/g, "'") : (plainSheet ?? null);
  const rowIndex = Number(startRow) - 1;
  const columnIndex = columnIndexFromLetter(startCol);
  const lastRowIndex = endRow ? Number(endRow) - 1 : rowIndex;
  const lastColumnIndex = endCol ? columnIndexFromLetter(endCol) : columnIndex;
  return {
    sheetName,
    rowIndex: Math.min(rowIndex, lastRowIndex),
    columnIndex: Math.min(columnIndex, lastColumnIndex),
    rowCount: Math.abs(lastRowIndex - rowIndex) + 1,
    columnCount: Math.abs(lastColumnIndex - columnIndex) + 1,
  };
}

export function ok(message: string, details?: string[]): OperationResult {
  return { ok: true, message, details };
}

export function fail(message: string, details?: string[]): OperationResult {
  return { ok: false, message, details };
}

/** Keep "1 row"/"2 rows" pluralisation in one place. */
export function plural(count: number, singular: string, pluralForm = `${singular}s`): string {
  return `${count.toLocaleString()} ${count === 1 ? singular : pluralForm}`;
}

// ---------------------------------------------------------------------------
// Office.js helpers
// ---------------------------------------------------------------------------

/**
 * Office.js rejects very large single payloads, so grid writes are split into
 * batches of roughly this many cells.
 */
const MAX_CELLS_PER_BATCH = 40000;

/** Human-readable text for anything Excel.run can throw. */
export function describeExcelError(error: unknown): string {
  if (error && typeof error === "object" && "code" in error && "message" in error) {
    const officeError = error as {
      code?: string;
      message?: string;
      debugInfo?: { errorLocation?: string };
    };
    const where = officeError.debugInfo?.errorLocation
      ? ` (${officeError.debugInfo.errorLocation})`
      : "";
    return `${officeError.message ?? "Excel reported an error"}${where}`;
  }
  return error instanceof Error ? error.message : String(error);
}

/**
 * Run an Excel batch and normalise both success and failure into an
 * `OperationResult`, so no feature module needs its own try/catch.
 */
export async function runExcel(
  action: (context: Excel.RequestContext) => Promise<OperationResult>
): Promise<OperationResult> {
  try {
    return await Excel.run(action);
  } catch (error) {
    return fail(describeExcelError(error));
  }
}

/** The user's current selection, loaded with the properties every feature needs. */
export async function loadSelectedRange(context: Excel.RequestContext): Promise<Excel.Range> {
  const range = context.workbook.getSelectedRange();
  range.load([
    "address",
    "values",
    "formulas",
    "rowIndex",
    "columnIndex",
    "rowCount",
    "columnCount",
    "worksheet/name",
  ]);
  await context.sync();
  return range;
}

/** The used range of a worksheet, or `null` when the sheet is empty. */
export async function loadUsedRange(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet
): Promise<Excel.Range | null> {
  const used = sheet.getUsedRangeOrNullObject(true);
  used.load(["address", "values", "rowIndex", "columnIndex", "rowCount", "columnCount"]);
  await context.sync();
  return used.isNullObject ? null : used;
}

/** Names of every worksheet in the workbook, in tab order. */
export async function loadSheetNames(context: Excel.RequestContext): Promise<string[]> {
  const sheets = context.workbook.worksheets;
  sheets.load("items/name");
  await context.sync();
  return sheets.items.map((sheet) => sheet.name);
}

/**
 * Write a grid starting at a zero-based offset, syncing between batches so
 * large merges do not exceed Office.js payload limits.
 */
export async function writeGrid(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  rowIndex: number,
  columnIndex: number,
  grid: Grid
): Promise<void> {
  if (grid.length === 0) {
    return;
  }
  const rectangle = rectangularize(grid);
  const width = rectangle[0].length;
  const rowsPerBatch = Math.max(1, Math.floor(MAX_CELLS_PER_BATCH / Math.max(width, 1)));
  let offset = 0;

  for (const batch of chunk(rectangle, rowsPerBatch)) {
    sheet.getRangeByIndexes(rowIndex + offset, columnIndex, batch.length, width).values =
      batch as unknown[][];
    offset += batch.length;
    // eslint-disable-next-line office-addins/no-context-sync-in-loop -- deliberate: the loop exists to keep each payload under Office's size limit, so one sync per batch is required.
    await context.sync();
  }
}

/**
 * Like `writeGrid`, but also applies a parallel grid of number-format codes so
 * merged dates, currency and percentages survive the trip.
 */
export async function writeGridWithFormats(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  rowIndex: number,
  columnIndex: number,
  grid: Grid,
  formats: string[][]
): Promise<void> {
  if (grid.length === 0) {
    return;
  }
  const rectangle = rectangularize(grid);
  const width = rectangle[0].length;
  const formatRows = rectangle.map((_row, index) => {
    const source = formats[index] ?? [];
    const out = new Array<string>(width);
    for (let i = 0; i < width; i += 1) {
      out[i] = source[i] ?? "General";
    }
    return out;
  });

  const rowsPerBatch = Math.max(1, Math.floor(MAX_CELLS_PER_BATCH / Math.max(width, 1)));
  let offset = 0;
  for (let start = 0; start < rectangle.length; start += rowsPerBatch) {
    const valueBatch = rectangle.slice(start, start + rowsPerBatch);
    const formatBatch = formatRows.slice(start, start + rowsPerBatch);
    const range = sheet.getRangeByIndexes(rowIndex + offset, columnIndex, valueBatch.length, width);
    // Format first so serial numbers render as dates the moment they land.
    range.numberFormat = formatBatch;
    range.values = valueBatch as unknown[][];
    offset += valueBatch.length;
    // eslint-disable-next-line office-addins/no-context-sync-in-loop -- deliberate: the loop exists to keep each payload under Office's size limit, so one sync per batch is required.
    await context.sync();
  }
}

/**
 * Write one value (and optionally a number format) into a single cell.
 *
 * Office.js checks the array you pass against the *whole* range, and a merged
 * range still reports every cell it covers - so assigning `[[value]]` to a
 * three-column merged title is rejected with "The number of rows or columns in
 * the input array doesn't match the size or dimensions of the range". Writing
 * through the top-left cell avoids that trap everywhere.
 */
export function writeCell(
  sheet: Excel.Worksheet,
  rowIndex: number,
  columnIndex: number,
  value: CellValue,
  numberFormat?: string
): Excel.Range {
  const cell = sheet.getRangeByIndexes(rowIndex, columnIndex, 1, 1);
  cell.values = [[value]];
  if (numberFormat) {
    cell.numberFormat = [[numberFormat]];
  }
  return cell;
}

/** Excel forbids these characters in sheet names and caps the name at 31 characters. */
export function sanitizeSheetName(desiredName: string): string {
  const cleaned = desiredName.replace(/[:\\/?*[\]]/g, "-").trim();
  return (cleaned === "" ? "Sheet" : cleaned).slice(0, 31);
}

/** Create a worksheet whose name does not collide with an existing tab. */
export async function createSheetWithUniqueName(
  context: Excel.RequestContext,
  desiredName: string
): Promise<Excel.Worksheet> {
  const existing = await loadSheetNames(context);
  const sheet = context.workbook.worksheets.add(
    uniqueName(sanitizeSheetName(desiredName), existing).slice(0, 31)
  );
  sheet.load("name");
  await context.sync();
  return sheet;
}
