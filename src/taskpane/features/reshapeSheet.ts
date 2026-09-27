/**
 * The Excel side of unpivoting: read the sheet, find the crosstab, write the
 * tidy version to a new sheet.
 *
 * It always writes somewhere new. Unpivoting is not an edit, it is a different
 * shape of the same data, and overwriting the sheet someone hands round in a
 * meeting would be the wrong kind of helpful. The original stays exactly as it
 * was, which also means the transformation can simply be repeated next month.
 */

import {
  createSheetWithUniqueName,
  fail,
  loadSheetNames,
  ok,
  plural,
  runExcel,
  writeGridWithFormats,
} from "../shared/excelHelpers";
import type { Grid, OperationResult } from "../shared/types";
import {
  CrosstabShape,
  DetectedCrosstab,
  describeCrosstab,
  detectCrosstabInGrid,
  unpivot,
  UnpivotOptions,
} from "./reshape";

/** Rows beyond this are left alone rather than freezing Excel. */
const MAX_ROWS = 20000;

export interface CrosstabLook {
  sheet: string;
  found: DetectedCrosstab | null;
  /** Set when the sheet was read but holds nothing usable. */
  why: string | null;
  rowCount: number;
  truncated: boolean;
}

/**
 * Looks at a sheet without changing anything, so the assistant can describe what
 * it would do and let the user say yes.
 */
export async function findCrosstab(sheetName?: string): Promise<CrosstabLook | null> {
  try {
    return await Excel.run(async (context) => {
      const sheet = sheetName
        ? context.workbook.worksheets.getItem(sheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      const used = sheet.getUsedRangeOrNullObject(true);
      used.load(["values", "rowCount"]);
      sheet.load("name");
      await context.sync();

      if (used.isNullObject) {
        return {
          sheet: sheet.name,
          found: null,
          why: "that sheet is empty",
          rowCount: 0,
          truncated: false,
        };
      }

      const truncated = used.rowCount > MAX_ROWS;
      const grid = (
        truncated ? (used.values as Grid).slice(0, MAX_ROWS) : (used.values as Grid)
      ) as Grid;

      const found = detectCrosstabInGrid(grid);
      return {
        sheet: sheet.name,
        found,
        why: found
          ? null
          : "I couldn't see a block of period columns in it - the headings would need to read like Jan, Feb, Mar or Q1, Q2, or 2023, 2024",
        rowCount: grid.length,
        truncated,
      };
    });
  } catch {
    return null;
  }
}

export interface UnpivotOutcome extends OperationResult {
  /** Name of the sheet that was written, when one was. */
  sheetName?: string;
  rowsOut?: number;
}

/**
 * Writes the tidy version of a crosstab to a new sheet and leaves the user
 * looking at it.
 */
export async function unpivotToNewSheet(
  sheetName: string | undefined,
  options: UnpivotOptions = {}
): Promise<UnpivotOutcome> {
  const look = await findCrosstab(sheetName);
  if (!look) return fail("I couldn't read that sheet.");
  if (!look.found) return fail(`I couldn't unpivot ${look.sheet}: ${look.why}.`);

  const found = look.found;
  const { shape, headers, rows } = found;
  const result = unpivot({ headers, rows }, shape, options);

  if (result.rowsOut === 0) {
    return fail("Every cell in the period columns was empty, so there would be nothing to write.");
  }

  return runExcel(async (context) => {
    const names = await loadSheetNames(context);
    const target = await createSheetWithUniqueName(context, `${look.sheet} (tidy)`.slice(0, 31));

    const grid: Grid = [result.table.headers, ...result.table.rows];
    const formats: string[][] = grid.map((_row, index) =>
      index === 0
        ? result.table.headers.map(() => "General")
        : result.formats.map((format) => format ?? "General")
    );
    await writeGridWithFormats(context, target, 0, 0, grid, formats);

    const headerRow = target.getRangeByIndexes(0, 0, 1, result.table.headers.length);
    headerRow.format.font.bold = true;
    headerRow.format.fill.color = "#eef3f1";
    target.getUsedRange().format.autofitColumns();

    // A frozen heading row is the difference between a table you can scroll and
    // one you lose your place in. Not every host has freezePanes, so it is
    // attempted and shrugged off.
    try {
      if (Office.context.requirements.isSetSupported("ExcelApi", "1.7")) {
        target.freezePanes.freezeRows(1);
      }
    } catch {
      // An older Excel simply doesn't freeze the row.
    }

    target.activate();
    await context.sync();

    const details = [
      `Found ${describeCrosstab(shape)}.`,
      `${plural(result.rowsIn, "row")} became ${plural(result.rowsOut, "row")}.`,
    ];
    if (result.skippedBlank > 0) {
      details.push(`${plural(result.skippedBlank, "empty cell")} left out.`);
    }
    if (look.truncated) {
      details.push(`Only the first ${MAX_ROWS.toLocaleString("en-US")} rows were read.`);
    }
    if (found.headerRows === 2) {
      details.push("The two heading rows were joined, so the year travels with the month.");
    }
    if (names.length > 0 && shape.datesKnown) {
      details.push(
        `The ${shape.periodColumnName.toLowerCase()} column holds real dates, so charts can use it.`
      );
    }

    return {
      ...ok(`Wrote ${plural(result.rowsOut, "tidy row")} to "${target.name}".`, details),
      sheetName: target.name,
      rowsOut: result.rowsOut,
    };
  });
}

/** Column headings of the crosstab, for showing the user what will move. */
export function crosstabColumns(
  shape: CrosstabShape,
  headers: readonly string[]
): {
  keys: string[];
  periods: string[];
} {
  return {
    keys: shape.idColumns.map((index) => headers[index] || `Column ${index + 1}`),
    periods: shape.periods.map((period) => period.source),
  };
}
