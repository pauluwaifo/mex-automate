/**
 * The Excel side of explaining a formula: read whatever is selected and hand it
 * to the parser.
 *
 * Only the top-left cell of the selection is read. Explaining forty cells at
 * once would produce forty sentences nobody reads, and in a filled-down column
 * they are all the same formula anyway.
 */

import { loadSelectedRange } from "../shared/excelHelpers";
import type { CellValue } from "../shared/types";

export interface SelectedFormula {
  sheet: string;
  /** Address of the cell the formula came from, e.g. "D4". */
  address: string;
  /** The formula, starting with "=", or "" when the cell holds a plain value. */
  formula: string;
  value: CellValue;
  /** How many cells were selected, so the pane can say it looked at the first. */
  selectedCells: number;
}

export async function readSelectedFormula(): Promise<SelectedFormula | null> {
  try {
    return await Excel.run(async (context) => {
      const range = await loadSelectedRange(context);
      range.load(["formulas", "values", "address", "rowCount", "columnCount", "worksheet/name"]);
      await context.sync();

      const formulas = range.formulas as unknown[][];
      const values = range.values as CellValue[][];
      const first = formulas?.[0]?.[0];
      const formula = typeof first === "string" ? first : first === undefined ? "" : String(first);

      const address = range.address.includes("!") ? range.address.split("!")[1] : range.address;
      const topLeft = address.split(":")[0];

      return {
        sheet: range.worksheet.name,
        address: topLeft,
        formula: formula.startsWith("=") ? formula : "",
        value: values?.[0]?.[0] ?? null,
        selectedCells: (range.rowCount || 1) * (range.columnCount || 1),
      };
    });
  } catch {
    return null;
  }
}
