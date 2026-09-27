/**
 * The Excel side of the guards: propose rules from what a sheet already holds,
 * and put them on the cells.
 *
 * Two decisions worth knowing about:
 *
 * - **The rules cover empty rows below the data.** A guard that only covers
 *   today's rows protects nothing, because the next value is typed underneath
 *   them. So each rule extends past the last row, which is where the typing
 *   happens.
 * - **Every rule is a warning, not a block.** Excel can refuse a value outright;
 *   we don't. A person who genuinely needs to enter something unusual gets a
 *   message explaining what the column expects and can carry on. A rule that
 *   cannot be overridden is a rule that gets switched off entirely.
 */

import { columnLetter, fail, ok, plural, runExcel } from "../shared/excelHelpers";
import type { Grid, OperationResult } from "../shared/types";
import { countRejections, Guard, guardRange, suggestGuards, toValidationRule } from "./guards";
import { locateTable, profileColumn } from "./tidy";

/** How far past the data a rule reaches, so it covers what gets typed next. */
const SPARE_ROWS = 500;

export interface GuardProposal {
  sheet: string;
  guards: Guard[];
  /** How many existing cells each rule would have flagged, keyed by header. */
  existingProblems: Record<string, number>;
  rows: number;
  /** Where the table starts, for addressing the ranges later. */
  firstDataRow: number;
  lastDataRow: number;
  firstColumn: number;
}

function supportsValidation(): boolean {
  try {
    return Office.context.requirements.isSetSupported("ExcelApi", "1.8");
  } catch {
    return false;
  }
}

/** Reads a sheet and works out which rules could go on it. Changes nothing. */
export async function proposeGuards(sheetName?: string): Promise<GuardProposal | null> {
  try {
    return await Excel.run(async (context) => {
      const sheet = sheetName
        ? context.workbook.worksheets.getItem(sheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      const used = sheet.getUsedRangeOrNullObject(true);
      used.load(["values", "numberFormat", "rowIndex", "columnIndex"]);
      sheet.load("name");
      await context.sync();
      if (used.isNullObject) return null;

      const grid = used.values as Grid;
      const table = locateTable(grid);
      if (!table) return null;

      const headers = (grid[table.headerRow] ?? []).map((value) => String(value ?? ""));
      const body = grid.slice(table.headerRow + 1, table.lastRow + 1);
      if (body.length === 0) return null;

      const formats = used.numberFormat as string[][];
      const profiles = headers.map((header, index) =>
        profileColumn(
          header,
          body.map((row) => row[index] ?? null),
          index,
          { formats: body.map((_row, at) => formats[table.headerRow + 1 + at]?.[index]) }
        )
      );

      const guards = suggestGuards(profiles, body);
      const existingProblems: Record<string, number> = {};
      for (const guard of guards) {
        existingProblems[guard.header] = countRejections(guard, body);
      }

      return {
        sheet: sheet.name,
        guards,
        existingProblems,
        rows: body.length,
        // 1-based rows, as Excel addresses them.
        firstDataRow: used.rowIndex + table.headerRow + 2,
        lastDataRow: used.rowIndex + table.lastRow + 1,
        firstColumn: used.columnIndex,
      };
    });
  } catch {
    return null;
  }
}

/** Puts the chosen rules on the sheet. */
export async function applyGuards(
  proposal: GuardProposal,
  chosen: readonly Guard[]
): Promise<OperationResult> {
  if (!supportsValidation()) {
    return fail(
      "This version of Excel can't have rules set by an add-in. Everything else works; you'd need Excel 2019 or a Microsoft 365 subscription for this one."
    );
  }
  if (chosen.length === 0) return fail("Pick at least one rule.");

  return runExcel(async (context) => {
    const sheet = context.workbook.worksheets.getItem(proposal.sheet);
    const applied: string[] = [];

    for (const guard of chosen) {
      const address = guardRange(
        (index) => columnLetter(proposal.firstColumn + index),
        guard.columnIndex,
        proposal.firstDataRow,
        proposal.lastDataRow,
        SPARE_ROWS
      );
      const firstCell = `${columnLetter(proposal.firstColumn + guard.columnIndex)}${proposal.firstDataRow}`;
      const range = sheet.getRange(address);

      range.dataValidation.clear();
      range.dataValidation.rule = toValidationRule(guard, {
        rangeAddress: `${columnLetter(proposal.firstColumn + guard.columnIndex)}${proposal.firstDataRow}:${columnLetter(proposal.firstColumn + guard.columnIndex)}${proposal.lastDataRow}`,
        firstCell,
      }) as unknown as Excel.DataValidationRule;

      // A warning, not a stop: the person can still enter what they meant to.
      range.dataValidation.errorAlert = {
        message: guard.errorMessage,
        showAlert: true,
        style: Excel.DataValidationAlertStyle.warning,
        title: `Check ${guard.header}`,
      };
      range.dataValidation.prompt = {
        message: guard.detail,
        showPrompt: false,
        title: guard.header,
      };
      applied.push(guard.header);
    }

    sheet.activate();
    await context.sync();

    return ok(`Put ${plural(applied.length, "rule")} on "${proposal.sheet}".`, [
      `Covered: ${applied.join(", ")}.`,
      `Each rule reaches ${SPARE_ROWS.toLocaleString("en-US")} rows past today's data, so it applies to what gets typed next.`,
      "They warn rather than block, so you can still enter something unusual on purpose.",
      'Say "/protect off" to take them off again.',
    ]);
  });
}

/** Removes every rule this sheet carries. */
export async function clearGuards(sheetName?: string): Promise<OperationResult> {
  return runExcel(async (context) => {
    const sheet = sheetName
      ? context.workbook.worksheets.getItem(sheetName)
      : context.workbook.worksheets.getActiveWorksheet();
    const used = sheet.getUsedRangeOrNullObject(true);
    used.load(["address", "rowIndex", "columnIndex", "rowCount", "columnCount"]);
    sheet.load("name");
    await context.sync();
    if (used.isNullObject) return fail("That sheet is empty.");

    // Rules reach below the data, so clear past the used range too.
    sheet
      .getRangeByIndexes(
        used.rowIndex,
        used.columnIndex,
        used.rowCount + SPARE_ROWS,
        used.columnCount
      )
      .dataValidation.clear();
    await context.sync();
    return ok(`Took the rules off "${sheet.name}".`);
  });
}
