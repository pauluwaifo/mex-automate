/**
 * Watching a sheet while you work on it.
 *
 * Checking a sheet tells you what was wrong a moment ago. If the point is to
 * stop mistakes rather than catalogue them, the check has to happen while the
 * typing is still fresh in mind - so this listens for edits and re-checks just
 * the cells that changed.
 *
 * Three things keep it from being a nuisance:
 *
 * - **Only the edited block is re-read.** A change event carries its own
 *   address, so a keystroke costs one small range read, not a full scan.
 * - **Edits are settled before checking.** Pasting a column fires a burst of
 *   events; they are collected for a moment and handled once.
 * - **It says nothing while you are mid-thought.** A half-typed cell is not a
 *   mistake, so findings are counted and shown quietly, and nothing is ever
 *   changed in the sheet without being asked.
 *
 * Worksheet events need ExcelApi 1.7. Where that isn't available the add-in says
 * so instead of failing.
 */

import { columnLetter, describeExcelError } from "../shared/excelHelpers";
import type { Grid } from "../shared/types";
import { Issue, reviewTable } from "./review";

/** How long to wait for edits to stop before re-checking. */
const SETTLE_MS = 600;

/** Re-reading a huge paste defeats the point; above this the change is ignored. */
const MAX_CELLS = 5000;

export interface WatchUpdate {
  sheet: string;
  /** Issues found in the cells that just changed. */
  issues: Issue[];
  /** Address of the block that changed, e.g. "B4:D9". */
  address: string;
  /** True when the edit was too large to check. */
  tooBig: boolean;
}

export interface Watcher {
  sheet: string;
  stop: () => Promise<void>;
}

export function isWatchSupported(): boolean {
  try {
    return Office.context.requirements.isSetSupported("ExcelApi", "1.7");
  } catch {
    return false;
  }
}

/**
 * Starts watching a sheet. `onUpdate` is called after each settled burst of
 * edits, with whatever the changed cells now look like - an empty issue list
 * means what was just typed is fine.
 */
export async function watchSheet(
  sheetName: string | undefined,
  onUpdate: (update: WatchUpdate) => void
): Promise<Watcher | { error: string }> {
  if (!isWatchSupported()) {
    return {
      error:
        "This version of Excel can't tell an add-in when cells change. Everything else works; you'll just need to re-run the check yourself.",
    };
  }

  try {
    let timer: ReturnType<typeof setTimeout> | null = null;
    let pending: string[] = [];

    const handler = async (event: Excel.WorksheetChangedEventArgs) => {
      pending.push(event.address);
      if (timer) clearTimeout(timer);
      timer = setTimeout(() => {
        const addresses = pending;
        pending = [];
        timer = null;
        void checkChanged(sheetName, addresses, onUpdate);
      }, SETTLE_MS);
    };

    const registration = await Excel.run(async (context) => {
      const sheet = sheetName
        ? context.workbook.worksheets.getItem(sheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");
      const result = sheet.onChanged.add(handler);
      await context.sync();
      return { name: sheet.name, result };
    });

    return {
      sheet: registration.name,
      stop: async () => {
        if (timer) clearTimeout(timer);
        try {
          await Excel.run(registration.result.context, async (context) => {
            registration.result.remove();
            await context.sync();
          });
        } catch {
          // A closed workbook has already forgotten the handler.
        }
      },
    };
  } catch (error) {
    return { error: describeExcelError(error) };
  }
}

/**
 * Re-reads the changed block and checks it. The block is widened to whole
 * columns of the used range's header row, because a single number means little
 * without knowing which column it is in.
 */
async function checkChanged(
  sheetName: string | undefined,
  addresses: readonly string[],
  onUpdate: (update: WatchUpdate) => void
): Promise<void> {
  try {
    await Excel.run(async (context) => {
      const sheet = sheetName
        ? context.workbook.worksheets.getItem(sheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");

      const used = sheet.getUsedRangeOrNullObject(true);
      used.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);
      await context.sync();
      if (used.isNullObject) return;

      // The union of everything that changed, as row/column bounds.
      let top = Number.POSITIVE_INFINITY;
      let bottom = -1;
      let left = Number.POSITIVE_INFINITY;
      let right = -1;
      const blocks = addresses.map((address) => {
        const range = sheet.getRange(address.includes("!") ? address.split("!")[1] : address);
        range.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);
        return range;
      });
      await context.sync();

      for (const block of blocks) {
        top = Math.min(top, block.rowIndex);
        left = Math.min(left, block.columnIndex);
        bottom = Math.max(bottom, block.rowIndex + block.rowCount - 1);
        right = Math.max(right, block.columnIndex + block.columnCount - 1);
      }
      if (bottom < 0 || right < 0) return;

      const headerRow = used.rowIndex;
      const firstRow = Math.max(headerRow, Math.min(top, bottom));
      const lastRow = Math.min(used.rowIndex + used.rowCount - 1, bottom);
      const firstColumn = Math.max(used.columnIndex, left);
      const lastColumn = Math.min(used.columnIndex + used.columnCount - 1, right);
      const rowCount = lastRow - firstRow + 1;
      const columnCount = lastColumn - firstColumn + 1;

      if (rowCount <= 0 || columnCount <= 0) return;
      if (rowCount * columnCount > MAX_CELLS) {
        onUpdate({
          sheet: sheet.name,
          issues: [],
          address: `${columnLetter(firstColumn)}${firstRow + 1}`,
          tooBig: true,
        });
        return;
      }

      // Read the heading row along with the changed rows, so a checked cell is
      // judged against the column it belongs to.
      const headings = sheet.getRangeByIndexes(headerRow, firstColumn, 1, columnCount);
      headings.load(["values"]);
      const changed = sheet.getRangeByIndexes(firstRow, firstColumn, rowCount, columnCount);
      changed.load(["values", "formulas", "numberFormat", "address"]);
      await context.sync();

      const includesHeader = firstRow === headerRow;
      const grid: Grid = includesHeader
        ? (changed.values as Grid)
        : [(headings.values as Grid)[0], ...(changed.values as Grid)];
      const formulas: string[][] = includesHeader
        ? (changed.formulas as string[][])
        : [
            (headings.values as Grid)[0].map((value) => String(value ?? "")),
            ...(changed.formulas as string[][]),
          ];
      const numberFormats: string[][] = includesHeader
        ? (changed.numberFormat as string[][])
        : [
            (headings.values as Grid)[0].map(() => "General"),
            ...(changed.numberFormat as string[][]),
          ];

      const result = reviewTable({
        grid,
        formulas,
        numberFormats,
        // The heading row sits at headerRow; the body starts at firstRow.
        origin: { row: includesHeader ? firstRow : firstRow - 1, column: firstColumn },
      });

      const address = changed.address.includes("!")
        ? changed.address.split("!")[1]
        : changed.address;

      onUpdate({ sheet: sheet.name, issues: result.issues, address, tooBig: false });
    });
  } catch {
    // A watcher that throws would be worse than one that misses an edit.
  }
}
