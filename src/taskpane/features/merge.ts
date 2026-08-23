/**
 * Merge/combine: stack several sheets - or several picked files - into one
 * table, lining columns up by header name.
 *
 * Column matching is by *normalized* header, so "Order Date", "order_date" and
 * "ORDER DATE" land in one column. Where two sources genuinely disagree
 * ("Customer" vs "Client Name") the user supplies an alias and the planner
 * treats them as the same column.
 *
 * `planHeaders` and `buildMergedTable` are pure and unit tested; the exported
 * `merge*` functions are the Office.js drivers around them.
 */

import {
  createSheetWithUniqueName,
  fail,
  isBlankRow,
  loadSheetNames,
  normalizeHeader,
  ok,
  padRow,
  plural,
  runExcel,
  writeGridWithFormats,
} from "../shared/excelHelpers";
import { CellValue, Grid, OperationResult } from "../shared/types";
import { readWorkbookFile } from "../shared/workbookReader";

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

/** One table feeding a merge: a worksheet, or a sheet inside a picked file. */
export interface SourceTable {
  /** Shown in the UI and written into the Source column, e.g. "orders.csv - Sheet1". */
  label: string;
  headers: string[];
  rows: Grid;
  /** Number-format codes parallel to `rows`. Optional; defaults to "General". */
  numberFormats?: string[][];
}

export interface MergeOptions {
  /** Record which sheet/file each row came from in a leading "Source" column. */
  addSourceColumn: boolean;
  /** Drop rows that are entirely blank. */
  skipBlankRows: boolean;
  /**
   * Maps a normalized source header to the canonical output header it should
   * merge into. Lets the user reconcile "Client Name" with "Customer".
   */
  headerAliases: Record<string, string>;
  /** Name for the sheet the merged table is written to. */
  destinationSheetName: string;
  /** Output header to sort the merged rows by. Empty keeps source order. */
  sortByHeader: string;
  sortDescending: boolean;
}

export const DEFAULT_MERGE_OPTIONS: MergeOptions = {
  addSourceColumn: true,
  skipBlankRows: true,
  headerAliases: {},
  destinationSheetName: "Merged",
  sortByHeader: "",
  sortDescending: false,
};

/** Layout choices `buildMergedTable` needs; sorting is optional. */
export type MergeLayoutOptions = Pick<MergeOptions, "addSourceColumn" | "skipBlankRows"> &
  Partial<Pick<MergeOptions, "sortByHeader" | "sortDescending">>;

function isBlankSortValue(value: CellValue): boolean {
  return value === null || value === "";
}

/**
 * Order two cells the way a spreadsheet user expects: numbers numerically, text
 * alphabetically and case-insensitively, blanks last.
 */
export function compareCells(a: CellValue, b: CellValue): number {
  if (isBlankSortValue(a) || isBlankSortValue(b)) {
    return isBlankSortValue(a) && isBlankSortValue(b) ? 0 : isBlankSortValue(a) ? 1 : -1;
  }
  if (typeof a === "number" && typeof b === "number") {
    return a - b;
  }
  return String(a).localeCompare(String(b), undefined, { numeric: true, sensitivity: "base" });
}

/**
 * `compareCells` with a direction. Blanks stay at the bottom either way -
 * simply negating the comparison would float every empty cell to the top, which
 * is never what someone sorting a merged table wants.
 */
export function compareCellsDirected(a: CellValue, b: CellValue, descending: boolean): number {
  if (isBlankSortValue(a) || isBlankSortValue(b)) {
    return isBlankSortValue(a) && isBlankSortValue(b) ? 0 : isBlankSortValue(a) ? 1 : -1;
  }
  const result = compareCells(a, b);
  return descending ? -result : result;
}

/**
 * Parse the alias box, one rule per line, written as
 * `Client Name = Customer`. The left side is the header as it appears in a
 * source; the right side is the column it should merge into.
 */
export function parseAliasLines(text: string): Record<string, string> {
  const aliases: Record<string, string> = {};
  for (const line of text.split(/\r?\n/)) {
    const separator = line.indexOf("=");
    if (separator < 0) {
      continue;
    }
    const from = line.slice(0, separator).trim();
    const to = line.slice(separator + 1).trim();
    if (from === "" || to === "") {
      continue;
    }
    aliases[normalizeHeader(from)] = to;
  }
  return aliases;
}

export interface HeaderPlan {
  /** The unified output headers, in the order they were first encountered. */
  headers: string[];
  /** For each source, the output column index of each of its columns. */
  columnTargets: number[][];
  /** Output headers that are missing from at least one source. */
  partialHeaders: string[];
  /** Which sources contributed each output header, keyed by header. */
  contributors: Record<string, string[]>;
}

export const SOURCE_COLUMN_HEADER = "Source";

// ---------------------------------------------------------------------------
// Pure planning
// ---------------------------------------------------------------------------

/**
 * Work out the unified column list. An alias is applied before normalizing, so
 * users can point an odd header at whatever canonical name they prefer.
 */
export function planHeaders(
  sources: readonly SourceTable[],
  headerAliases: Record<string, string> = {}
): HeaderPlan {
  const headers: string[] = [];
  const indexByKey = new Map<string, number>();
  const columnTargets: number[][] = [];
  const contributors: Record<string, string[]> = {};

  for (const source of sources) {
    const targets: number[] = [];

    source.headers.forEach((header) => {
      const aliased = headerAliases[normalizeHeader(header)] ?? header;
      const key = normalizeHeader(aliased);

      if (key === "") {
        // An unnamed column has nothing to match on; keep it positionally unique.
        targets.push(-1);
        return;
      }

      let index = indexByKey.get(key);
      if (index === undefined) {
        index = headers.length;
        indexByKey.set(key, index);
        headers.push(String(aliased).trim());
        contributors[headers[index]] = [];
      }
      targets.push(index);

      const label = headers[index];
      if (!contributors[label].includes(source.label)) {
        contributors[label].push(source.label);
      }
    });

    columnTargets.push(targets);
  }

  const partialHeaders = headers.filter((header) => contributors[header].length < sources.length);

  return { headers, columnTargets, partialHeaders, contributors };
}

export interface MergedTable {
  /** Header row followed by every source's rows. */
  grid: Grid;
  numberFormats: string[][];
  /** Rows written per source, keyed by source label. */
  rowsPerSource: Record<string, number>;
  totalRows: number;
}

/** Stack every source onto the planned column layout. */
export function buildMergedTable(
  sources: readonly SourceTable[],
  plan: HeaderPlan,
  options: MergeLayoutOptions
): MergedTable {
  const offset = options.addSourceColumn ? 1 : 0;
  const width = plan.headers.length + offset;

  const headerRow: CellValue[] = options.addSourceColumn
    ? [SOURCE_COLUMN_HEADER, ...plan.headers]
    : [...plan.headers];

  const grid: Grid = [headerRow];
  const numberFormats: string[][] = [new Array<string>(width).fill("General")];
  const rowsPerSource: Record<string, number> = {};

  sources.forEach((source, sourceIndex) => {
    const targets = plan.columnTargets[sourceIndex] ?? [];
    let written = 0;

    for (let rowIndex = 0; rowIndex < source.rows.length; rowIndex += 1) {
      const row = source.rows[rowIndex];
      if (options.skipBlankRows && isBlankRow(row)) {
        continue;
      }

      const outRow = new Array<CellValue>(width).fill(null);
      const outFormats = new Array<string>(width).fill("General");
      if (options.addSourceColumn) {
        outRow[0] = source.label;
      }

      targets.forEach((target, columnIndex) => {
        if (target < 0) {
          return;
        }
        outRow[target + offset] = row[columnIndex] ?? null;
        outFormats[target + offset] = source.numberFormats?.[rowIndex]?.[columnIndex] ?? "General";
      });

      grid.push(outRow);
      numberFormats.push(outFormats);
      written += 1;
    }

    rowsPerSource[source.label] = (rowsPerSource[source.label] ?? 0) + written;
  });

  const sortColumn = options.sortByHeader ? headerRow.indexOf(options.sortByHeader) : -1;
  if (sortColumn >= 0 && grid.length > 2) {
    // Sort an index permutation so values and their number formats stay paired.
    const order = grid
      .slice(1)
      .map((_row, index) => index + 1)
      .sort((left, right) =>
        compareCellsDirected(
          grid[left][sortColumn],
          grid[right][sortColumn],
          Boolean(options.sortDescending)
        )
      );

    const sortedGrid: Grid = [headerRow, ...order.map((index) => grid[index])];
    const sortedFormats: string[][] = [
      numberFormats[0],
      ...order.map((index) => numberFormats[index]),
    ];
    grid.length = 0;
    grid.push(...sortedGrid);
    numberFormats.length = 0;
    numberFormats.push(...sortedFormats);
  }

  return { grid, numberFormats, rowsPerSource, totalRows: grid.length - 1 };
}

/** Human-readable notes about a plan, surfaced under the success message. */
export function describePlan(plan: HeaderPlan, merged: MergedTable): string[] {
  const details = Object.entries(merged.rowsPerSource).map(
    ([label, count]) => `${label}: ${plural(count, "row")}`
  );

  if (plan.partialHeaders.length > 0) {
    const shown = plan.partialHeaders.slice(0, 6).join(", ");
    const more = plan.partialHeaders.length > 6 ? `, +${plan.partialHeaders.length - 6} more` : "";
    details.push(`Columns missing from some sources (left blank there): ${shown}${more}`);
  }

  return details;
}

// ---------------------------------------------------------------------------
// Office.js drivers
// ---------------------------------------------------------------------------

/** Split a raw used-range grid into a header row plus body rows. */
function toSourceTable(label: string, grid: Grid, numberFormats?: string[][]): SourceTable | null {
  if (grid.length === 0) {
    return null;
  }
  const headers = grid[0].map((cell, index) => {
    const text = String(cell ?? "").trim();
    return text === "" ? `Column ${index + 1}` : text;
  });
  return {
    label,
    headers,
    rows: grid.slice(1).map((row) => padRow(row, headers.length)),
    numberFormats: numberFormats?.slice(1),
  };
}

/** Read the used range of the named worksheets as merge sources. */
async function readSheetSources(
  context: Excel.RequestContext,
  sheetNames: readonly string[]
): Promise<{ sources: SourceTable[]; skipped: string[] }> {
  const loaded = sheetNames.map((name) => {
    const sheet = context.workbook.worksheets.getItem(name);
    const used = sheet.getUsedRangeOrNullObject(true);
    used.load(["values", "numberFormat"]);
    return { name, used };
  });
  await context.sync();

  const sources: SourceTable[] = [];
  const skipped: string[] = [];
  for (const entry of loaded) {
    if (entry.used.isNullObject) {
      skipped.push(entry.name);
      continue;
    }
    const table = toSourceTable(
      entry.name,
      entry.used.values as Grid,
      entry.used.numberFormat as string[][]
    );
    if (table && table.rows.length > 0) {
      sources.push(table);
    } else {
      skipped.push(entry.name);
    }
  }
  return { sources, skipped };
}

/** Write a merged table to a fresh sheet and format it as an Excel table. */
async function writeMergedTable(
  context: Excel.RequestContext,
  merged: MergedTable,
  plan: HeaderPlan,
  options: MergeOptions,
  skipped: string[]
): Promise<OperationResult> {
  if (merged.totalRows === 0) {
    return fail("Nothing to merge - every source was empty.");
  }

  const sheet = await createSheetWithUniqueName(context, options.destinationSheetName);
  await writeGridWithFormats(context, sheet, 0, 0, merged.grid, merged.numberFormats);

  const width = merged.grid[0].length;
  const table = sheet.tables.add(
    sheet.getRangeByIndexes(0, 0, merged.grid.length, width),
    true /* hasHeaders */
  );
  table.name = `MExMerged_${Date.now()}`;
  sheet.getUsedRange().format.autofitColumns();
  // Keep the headers on screen while scrolling a long merged table.
  sheet.freezePanes.freezeRows(1);
  sheet.activate();
  await context.sync();

  const details = describePlan(plan, merged);
  if (skipped.length > 0) {
    details.push(`Skipped (no data): ${skipped.join(", ")}`);
  }

  return ok(
    `Merged ${plural(merged.totalRows, "row")} and ${plural(plan.headers.length, "column")} into "${sheet.name}".`,
    details
  );
}

/** Merge the used range of several worksheets in the current workbook. */
export async function mergeSheets(
  sheetNames: readonly string[],
  options: MergeOptions
): Promise<OperationResult> {
  if (sheetNames.length < 2) {
    return fail("Pick at least two sheets to merge.");
  }

  return runExcel(async (context) => {
    const { sources, skipped } = await readSheetSources(context, sheetNames);
    if (sources.length === 0) {
      return fail("None of the selected sheets contain data.");
    }
    const plan = planHeaders(sources, options.headerAliases);
    const merged = buildMergedTable(sources, plan, options);
    return writeMergedTable(context, merged, plan, options, skipped);
  });
}

/**
 * Merge files the user picked from disk into the current workbook. The files are
 * parsed in the browser; nothing is uploaded anywhere.
 */
export async function mergeFiles(
  files: readonly File[],
  options: MergeOptions
): Promise<OperationResult> {
  if (files.length === 0) {
    return fail("Choose at least one file to merge.");
  }

  const sources: SourceTable[] = [];
  const problems: string[] = [];

  for (const file of files) {
    try {
      const workbook = await readWorkbookFile(file);
      const withData = workbook.sheets.filter((sheet) => sheet.grid.length > 1);
      if (withData.length === 0) {
        problems.push(`${file.name}: no data rows found.`);
        continue;
      }
      for (const sheet of withData) {
        // Only qualify the label with the sheet name when the file has several.
        const label = withData.length > 1 ? `${file.name} - ${sheet.name}` : file.name;
        const table = toSourceTable(label, sheet.grid, sheet.numberFormats);
        if (table) {
          sources.push(table);
        }
      }
    } catch (error) {
      problems.push(error instanceof Error ? error.message : String(error));
    }
  }

  if (sources.length === 0) {
    return fail("None of the chosen files could be read.", problems);
  }

  const plan = planHeaders(sources, options.headerAliases);
  const merged = buildMergedTable(sources, plan, options);

  const result = await runExcel(async (context) =>
    writeMergedTable(context, merged, plan, options, [])
  );

  if (problems.length > 0 && result.details) {
    result.details.push(...problems);
  } else if (problems.length > 0) {
    result.details = problems;
  }
  return result;
}

/** Sheet names for the merge picker. */
export async function listSheetNames(): Promise<string[]> {
  try {
    return await Excel.run((context) => loadSheetNames(context));
  } catch {
    return [];
  }
}

/**
 * Read headers only, so the UI can show the planned column layout - and any
 * mismatches - before the user commits to the merge.
 */
export async function previewSheetMerge(
  sheetNames: readonly string[],
  headerAliases: Record<string, string>
): Promise<HeaderPlan> {
  try {
    return await Excel.run(async (context) => {
      const { sources } = await readSheetSources(context, sheetNames);
      return planHeaders(sources, headerAliases);
    });
  } catch {
    return { headers: [], columnTargets: [], partialHeaders: [], contributors: {} };
  }
}
