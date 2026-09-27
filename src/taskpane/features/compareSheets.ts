/**
 * The Excel side of comparing two sheets: read both, match the rows, write a
 * report, and colour the cells that moved.
 *
 * The report is a new sheet rather than a message in the pane, because the
 * answer to "what changed" is usually something to send to somebody, and a list
 * of forty differences does not belong in a chat bubble.
 *
 * Marking is done through Review's own mark register, so "Clear marks" lifts
 * these highlights too and puts each cell's own fill back.
 */

import {
  buildAddress,
  columnLetter,
  createSheetWithUniqueName,
  fail,
  ok,
  plural,
  runExcel,
  writeGrid,
} from "../shared/excelHelpers";
import { toHeaderTable } from "../shared/excelHelpers";
import type { Grid, HeaderTable, OperationResult } from "../shared/types";
import {
  compareTables,
  CompareOptions,
  Diff,
  diffReportGrid,
  largestMovements,
  summarizeDiff,
} from "./compare";
import { markCells } from "./reviewSheet";

/** Rows beyond this are left unread rather than freezing Excel. */
const MAX_ROWS = 20000;

/** Marking every difference in a 500-row diff makes the sheet unreadable. */
const MAX_MARKS = 200;

interface SheetTable {
  name: string;
  table: HeaderTable;
  /** Row index in the sheet where the body starts, for addressing cells. */
  bodyRow: number;
  columnIndex: number;
  truncated: boolean;
}

async function readTable(
  context: Excel.RequestContext,
  sheetName: string
): Promise<SheetTable | null> {
  const sheet = context.workbook.worksheets.getItem(sheetName);
  const used = sheet.getUsedRangeOrNullObject(true);
  used.load(["values", "rowIndex", "columnIndex", "rowCount"]);
  sheet.load("name");
  await context.sync();
  if (used.isNullObject) return null;

  const truncated = used.rowCount > MAX_ROWS;
  const values = (
    truncated ? (used.values as Grid).slice(0, MAX_ROWS) : (used.values as Grid)
  ) as Grid;
  if (values.length < 2) return null;

  return {
    name: sheet.name,
    table: toHeaderTable(values),
    bodyRow: used.rowIndex + 1,
    columnIndex: used.columnIndex,
    truncated,
  };
}

export interface ComparisonOutcome extends OperationResult {
  diff?: Diff;
  reportSheet?: string;
}

export interface CompareRequest {
  first: string;
  second: string;
  options?: CompareOptions;
  /** Write the differences to a new sheet. On by default. */
  writeReport?: boolean;
  /** Colour the changed cells in the newer sheet. On by default. */
  mark?: boolean;
}

/**
 * Compares two sheets in the open workbook. `first` is treated as the older
 * version, so "added" means present in `second` only.
 */
export async function compareSheets(request: CompareRequest): Promise<ComparisonOutcome> {
  const { writeReport = true, mark = true } = request;

  return runExcel(async (context) => {
    const first = await readTable(context, request.first);
    const second = await readTable(context, request.second);

    if (!first) return fail(`"${request.first}" has no table in it to compare.`);
    if (!second) return fail(`"${request.second}" has no table in it to compare.`);

    const diff = compareTables(first.table, second.table, request.options);

    const details: string[] = [];
    if (diff.keyColumn) {
      details.push(`Matched rows on "${diff.keyColumn}".`);
    } else {
      details.push(
        "No column was unique enough to match rows on, so whole rows were compared. That finds rows added and removed, but not cells edited within a row."
      );
    }
    if (first.truncated || second.truncated) {
      details.push(
        `Only the first ${MAX_ROWS.toLocaleString("en-US")} rows of each sheet were read.`
      );
    }
    for (const movement of largestMovements(diff, 3)) {
      details.push(
        `${movement.key} - ${movement.header} moved by ${(movement.delta as number).toLocaleString("en-US")}.`
      );
    }

    let reportSheet: string | undefined;
    if (
      writeReport &&
      (diff.changed.length > 0 || diff.added.length > 0 || diff.removed.length > 0)
    ) {
      const report = diffReportGrid(diff, { first: first.name, second: second.name });
      const sheet = await createSheetWithUniqueName(context, `${second.name} vs ${first.name}`);
      await writeGrid(context, sheet, 0, 0, report.grid);

      const width = report.grid[0]?.length ?? 1;
      for (const row of report.headingRows) {
        const range = sheet.getRangeByIndexes(row, 0, 1, width);
        range.format.font.bold = true;
        range.format.font.size = 12;
      }
      for (const row of report.tableHeaderRows) {
        const range = sheet.getRangeByIndexes(row, 0, 1, width);
        range.format.font.bold = true;
        range.format.fill.color = "#eef3f1";
      }
      sheet.getUsedRange().format.autofitColumns();
      sheet.activate();
      await context.sync();
      reportSheet = sheet.name;
    }

    if (mark && diff.changed.length > 0) {
      // Address each changed cell in the newer sheet, where the current numbers
      // are, since that is the sheet someone will go and look at.
      const addresses: string[] = [];
      for (const change of diff.changed.slice(0, MAX_MARKS)) {
        const columnAt = second.table.headers.indexOf(change.header);
        if (columnAt < 0) continue;
        addresses.push(
          `${columnLetter(second.columnIndex + columnAt)}${second.bodyRow + change.rowAfter}`
        );
      }
      if (addresses.length > 0) {
        await markCells(second.name, addresses, "warning");
        if (diff.changed.length > addresses.length) {
          details.push(
            `Marked the first ${addresses.length} changed cells in "${second.name}"; the report lists all ${diff.changed.length}.`
          );
        } else {
          details.push(`Marked ${plural(addresses.length, "changed cell")} in "${second.name}".`);
        }
      }
    }

    if (reportSheet) details.push(`Full list in "${reportSheet}".`);

    return {
      ...ok(summarizeDiff(diff), details),
      diff,
      reportSheet,
    };
  });
}

/**
 * The two sheets to compare when the user didn't say: the active sheet and the
 * one next to it, which is how monthly files are usually arranged.
 */
export async function guessSheetsToCompare(): Promise<{ first: string; second: string } | null> {
  try {
    return await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name,items/position");
      const active = context.workbook.worksheets.getActiveWorksheet();
      active.load("name,position");
      await context.sync();

      const ordered = sheets.items
        .slice()
        .sort((a, b) => a.position - b.position)
        .map((sheet) => sheet.name);
      if (ordered.length < 2) return null;

      const at = ordered.indexOf(active.name);
      if (at < 0) return { first: ordered[0], second: ordered[1] };
      // Prefer the sheet before the active one as the older version.
      if (at > 0) return { first: ordered[at - 1], second: active.name };
      return { first: active.name, second: ordered[1] };
    });
  } catch {
    return null;
  }
}

/** Address of one changed cell in the newer sheet, for a "go to" button. */
export function changedCellAddress(
  sheetName: string,
  headers: readonly string[],
  bodyRow: number,
  columnIndex: number,
  header: string,
  rowAfter: number
): string | null {
  const at = headers.indexOf(header);
  if (at < 0) return null;
  return buildAddress(sheetName, bodyRow + rowAfter - 1, columnIndex + at, 1, 1);
}
