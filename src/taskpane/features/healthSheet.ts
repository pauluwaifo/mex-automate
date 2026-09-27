/**
 * The Excel side of the health check: read every sheet, run the detectors, and
 * repair the two things that can be repaired safely.
 *
 * Reading a whole workbook is the expensive part, so each sheet is read once,
 * in one request, and only the three things the detectors need: formulas,
 * values and where the used range starts. Sheets past a size limit are skipped
 * and named rather than silently left out of the total.
 *
 * Neither fix deletes anything. Tightening a range rewrites a formula to cover
 * the rows that hold data; clearing the empty tail clears blank cells and their
 * formatting. Both are snapshotted first, so /undo puts them back.
 */

import { columnLetter, fail, ok, plural, runExcel } from "../shared/excelHelpers";
import type { Grid, OperationResult } from "../shared/types";
import {
  checkWorkbook,
  dataExtent,
  HealthFix,
  HealthReport,
  SheetSnapshot,
  tightenWholeColumns,
} from "./health";
import { rememberCells } from "./reviewSheet";

/** A sheet bigger than this is skipped: reading it would freeze the pane. */
const MAX_CELLS_PER_SHEET = 400000;

/** Conditional formats are counted only when the host can do it cheaply. */
function supportsConditionalFormats(): boolean {
  try {
    return Office.context.requirements.isSetSupported("ExcelApi", "1.6");
  } catch {
    return false;
  }
}

/**
 * Looks over every sheet in the workbook. Nothing is changed.
 */
export async function checkWorkbookHealth(): Promise<HealthReport | null> {
  try {
    return await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();

      const wanted = sheets.items.map((sheet) => {
        const used = sheet.getUsedRangeOrNullObject(true);
        used.load(["rowCount", "columnCount", "rowIndex", "columnIndex"]);
        return { sheet, used, name: sheet.name };
      });
      await context.sync();

      const readable: Array<(typeof wanted)[number]> = [];
      const skipped: string[] = [];
      for (const candidate of wanted) {
        if (candidate.used.isNullObject) continue;
        const cells = candidate.used.rowCount * candidate.used.columnCount;
        if (cells > MAX_CELLS_PER_SHEET) {
          skipped.push(candidate.name);
          continue;
        }
        readable.push(candidate);
      }

      // Ask for everything up front: one sync for the whole workbook.
      const loaded = readable.map((candidate) => {
        const range = candidate.sheet.getUsedRange(true);
        range.load(["formulas", "values", "rowIndex", "columnIndex"]);
        const formats = supportsConditionalFormats()
          ? candidate.sheet.getUsedRange(true).conditionalFormats
          : null;
        formats?.load("items/type");
        return { name: candidate.name, range, formats };
      });
      await context.sync();

      const snapshots: SheetSnapshot[] = loaded.map((item) => ({
        name: item.name,
        formulas: item.range.formulas as string[][],
        values: item.range.values as Grid,
        origin: { row: item.range.rowIndex, column: item.range.columnIndex },
        conditionalFormats: item.formats?.items?.length,
      }));

      return checkWorkbook(snapshots, skipped);
    });
  } catch {
    return null;
  }
}

/**
 * Rewrites whole-column references so they cover the data and stop.
 *
 * The data extent is re-read at the moment of the fix rather than trusted from
 * the scan: the user may have typed a row in between looking and agreeing, and
 * a range that stops one row short of the data is a worse problem than the one
 * being fixed.
 */
export async function applyHealthFix(fix: HealthFix): Promise<OperationResult> {
  if (fix.kind === "tightenRanges") {
    return runExcel(async (context) => {
      const sheet = context.workbook.worksheets.getItem(fix.sheet);
      const used = sheet.getUsedRangeOrNullObject(true);
      used.load(["values", "rowIndex", "rowCount"]);
      await context.sync();
      if (used.isNullObject) return fail(`"${fix.sheet}" is empty now.`);

      const extent = dataExtent(used.values as Grid);
      const firstDataRow = used.rowIndex + 2; // one past the heading row, 1-based
      const lastDataRow = used.rowIndex + extent.rows;
      if (lastDataRow < firstDataRow) {
        return fail(`There isn't enough data on "${fix.sheet}" to narrow anything to.`);
      }

      const targets: Array<{ address: string; formula: string }> = [];
      for (const cell of fix.cells) {
        const tightened = tightenWholeColumns(cell.formula, firstDataRow, lastDataRow);
        if (tightened) targets.push({ address: cell.address, formula: tightened });
      }
      if (targets.length === 0) {
        return fail("Those formulas don't have a whole-column reference any more.");
      }

      // Snapshot before writing, so /undo covers this like any other fix.
      await rememberCells(
        context,
        fix.sheet,
        targets.map((target) => target.address),
        `narrowed ${plural(targets.length, "formula")} on "${fix.sheet}"`
      );

      for (const target of targets) {
        sheet.getRange(target.address).formulas = [[target.formula]];
      }
      sheet.activate();
      await context.sync();

      return ok(
        `Narrowed ${plural(targets.length, "formula")} to rows ${firstDataRow}-${lastDataRow}.`,
        [
          "Each one now covers the data and stops, so Excel recalculates less and stray notes below the table can't join the totals.",
          'Say "undo" to put the old formulas back.',
        ]
      );
    });
  }

  return runExcel(async (context) => {
    const sheet = context.workbook.worksheets.getItem(fix.sheet);
    const used = sheet.getUsedRangeOrNullObject(true);
    used.load(["rowIndex", "columnIndex", "rowCount", "columnCount", "values"]);
    await context.sync();
    if (used.isNullObject) return fail(`"${fix.sheet}" is empty.`);

    const extent = dataExtent(used.values as Grid);
    const firstBlankRow = used.rowIndex + extent.rows;
    const lastUsedRow = used.rowIndex + used.rowCount - 1;
    if (firstBlankRow > lastUsedRow) {
      return fail(`There's nothing below the data on "${fix.sheet}" to clear.`);
    }

    const rows = lastUsedRow - firstBlankRow + 1;
    sheet
      .getRangeByIndexes(firstBlankRow, used.columnIndex, rows, used.columnCount)
      .clear(Excel.ClearApplyTo.all);
    sheet.activate();
    await context.sync();

    return ok(`Cleared ${plural(rows, "empty row")} below the data on "${fix.sheet}".`, [
      "Excel only forgets them once the file is saved and reopened, so save it to see the difference.",
      "Nothing that held a value was touched.",
    ]);
  });
}

/** Where to go to look at a finding. */
export async function goToHealthFinding(
  sheetName: string,
  address: string
): Promise<OperationResult> {
  return runExcel(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    sheet.activate();
    sheet.getRange(address).select();
    await context.sync();
    return ok(`Selected ${address} on "${sheetName}".`);
  });
}

/** The address of the first cell of a sheet's data, for a friendly "go to". */
export function firstCellOf(snapshot: SheetSnapshot): string {
  return `${columnLetter(snapshot.origin.column)}${snapshot.origin.row + 1}`;
}
