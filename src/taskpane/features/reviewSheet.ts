/**
 * The Excel side of Review: scan the sheet, mark what's wrong *in the sheet*,
 * jump to any cell, apply fixes, and undo them.
 *
 * Two things matter here beyond the mechanics:
 *
 * - **Marks are borrowed, not taken.** Before highlighting a cell, its own fill
 *   colour is remembered in the workbook, so "Clear marks" puts back exactly
 *   what was there rather than blanking the user's own formatting.
 * - **Undo is ours, not Excel's.** Excel's Ctrl+Z doesn't reliably cover what
 *   an add-in writes, so every fix snapshots the cells it is about to change
 *   and "Undo last fix" writes them back.
 */

import {
  buildAddress,
  columnLetter,
  describeExcelError,
  fail,
  loadSelectedRange,
  ok,
  plural,
  runExcel,
} from "../shared/excelHelpers";
import { CellValue, Grid, OperationResult } from "../shared/types";
import { Issue, IssueGroup, reviewTable, ReviewResult, Severity } from "./review";

// ---------------------------------------------------------------------------
// Marking
// ---------------------------------------------------------------------------

/** Fills used to mark each severity, light enough to read the cell through. */
const MARK_COLOR: Record<Severity, string> = {
  error: "#fbd9d3",
  warning: "#fdecc8",
  tidy: "#dce9f7",
};

/** Cap the marking so a sheet full of problems doesn't become unreadable. */
const MAX_MARKS = 300;

const MARKS_KEY = "MExAutomate.review.marks";
const IGNORED_KEY = "MExAutomate.review.ignored";

interface MarkRecord {
  sheet: string;
  /** address -> the fill colour that was there before, "" for no fill. */
  cells: Record<string, string>;
}

async function readSetting<T>(context: Excel.RequestContext, key: string, fallback: T): Promise<T> {
  const setting = context.workbook.settings.getItemOrNullObject(key);
  setting.load(["value"]);
  await context.sync();
  if (setting.isNullObject) return fallback;
  try {
    return JSON.parse(String(setting.value)) as T;
  } catch {
    return fallback;
  }
}

function writeSetting(context: Excel.RequestContext, key: string, value: unknown): void {
  context.workbook.settings.add(key, JSON.stringify(value));
}

// ---------------------------------------------------------------------------
// Scanning
// ---------------------------------------------------------------------------

export interface SheetReview extends ReviewResult {
  sheet: string;
  /** Issues the user has told us to stop reporting. */
  ignored: number;
  /** True when the sheet was too big to scan completely. */
  truncated: boolean;
}

/** Rows beyond this are left unscanned rather than freezing Excel. */
const MAX_ROWS = 20000;

/** Look over a sheet (the active one by default) and list what looks wrong. */
export async function reviewSheet(sheetName?: string): Promise<SheetReview | null> {
  try {
    return await Excel.run(async (context) => {
      const sheet = sheetName
        ? context.workbook.worksheets.getItem(sheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      const used = sheet.getUsedRangeOrNullObject(true);
      used.load(["values", "formulas", "numberFormat", "rowIndex", "columnIndex", "rowCount"]);
      sheet.load("name");
      await context.sync();
      if (used.isNullObject) return null;

      const truncated = used.rowCount > MAX_ROWS;
      const limit = (rows: unknown[][]) => (truncated ? rows.slice(0, MAX_ROWS) : rows);
      const ignored = await readSetting<string[]>(context, IGNORED_KEY, []);

      const result = reviewTable({
        grid: limit(used.values as Grid[]) as Grid,
        formulas: limit(used.formulas as string[][]) as string[][],
        numberFormats: limit(used.numberFormat as string[][]) as string[][],
        origin: { row: used.rowIndex, column: used.columnIndex },
      });

      const ignoredSet = new Set(ignored);
      const kept = result.issues.filter((issue) => !ignoredSet.has(`${sheet.name}|${issue.id}`));
      return {
        ...result,
        issues: kept,
        sheet: sheet.name,
        ignored: result.issues.length - kept.length,
        truncated,
      };
    });
  } catch {
    return null;
  }
}

/**
 * Colour the problem cells in the sheet, remembering each cell's own fill so
 * the marks can be lifted again later.
 */
export async function markIssues(
  sheetName: string,
  issues: readonly Issue[]
): Promise<OperationResult> {
  return runExcel(async (context) => {
    await clearMarksIn(context);
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const shown = issues.slice(0, MAX_MARKS);
    if (shown.length === 0) {
      await context.sync();
      return ok("Nothing to mark.");
    }

    // Read the fills we are about to paint over, so they can be restored.
    const ranges = shown.map((issue) => {
      const range = sheet.getRangeByIndexes(issue.row, issue.column, 1, 1);
      range.load(["address", "format/fill/color"]);
      return range;
    });
    await context.sync();

    const record: MarkRecord = { sheet: sheetName, cells: {} };
    ranges.forEach((range, i) => {
      const local = range.address.includes("!") ? range.address.split("!")[1] : range.address;
      // An unfilled cell reads back as white; treat that as "no fill".
      record.cells[local] = range.format.fill.color === "#FFFFFF" ? "" : range.format.fill.color;
      range.format.fill.color = MARK_COLOR[shown[i].severity];
    });

    writeSetting(context, MARKS_KEY, record);
    sheet.activate();
    await context.sync();

    const more = issues.length - shown.length;
    return ok(
      `Marked ${plural(shown.length, "cell")} on "${sheetName}".`,
      more > 0
        ? [`${more.toLocaleString()} more aren't marked, to keep the sheet readable.`]
        : undefined
    );
  });
}

/**
 * Mark cells given by address rather than by issue, so anything that finds
 * something worth looking at - a comparison, a live check - lands in the same
 * register and is lifted by the same "Clear marks".
 */
export async function markCells(
  sheetName: string,
  addresses: readonly string[],
  severity: Severity = "warning",
  options: { replace?: boolean } = {}
): Promise<OperationResult> {
  const { replace = true } = options;
  return runExcel(async (context) => {
    const previous = replace
      ? null
      : await readSetting<MarkRecord | null>(context, MARKS_KEY, null);
    if (replace) await clearMarksIn(context);

    const sheet = context.workbook.worksheets.getItem(sheetName);
    const shown = addresses.slice(0, MAX_MARKS);
    if (shown.length === 0) {
      await context.sync();
      return ok("Nothing to mark.");
    }

    const ranges = shown.map((address) => {
      const range = sheet.getRange(address);
      range.load(["address", "format/fill/color"]);
      return range;
    });
    await context.sync();

    // Keeping the earlier marks means adding to their record, not starting a new
    // one, or the first set could never be put back.
    const record: MarkRecord =
      previous && previous.sheet === sheetName
        ? { sheet: sheetName, cells: { ...previous.cells } }
        : { sheet: sheetName, cells: {} };

    ranges.forEach((range) => {
      const local = range.address.includes("!") ? range.address.split("!")[1] : range.address;
      if (!(local in record.cells)) {
        record.cells[local] = range.format.fill.color === "#FFFFFF" ? "" : range.format.fill.color;
      }
      range.format.fill.color = MARK_COLOR[severity];
    });

    writeSetting(context, MARKS_KEY, record);
    await context.sync();

    return ok(`Marked ${plural(shown.length, "cell")} on "${sheetName}".`);
  });
}

async function clearMarksIn(context: Excel.RequestContext): Promise<number> {
  const record = await readSetting<MarkRecord | null>(context, MARKS_KEY, null);
  if (!record) return 0;
  const sheet = context.workbook.worksheets.getItemOrNullObject(record.sheet);
  await context.sync();
  if (!sheet.isNullObject) {
    for (const [address, color] of Object.entries(record.cells)) {
      const range = sheet.getRange(address);
      if (color) range.format.fill.color = color;
      else range.format.fill.clear();
    }
  }
  context.workbook.settings.add(MARKS_KEY, JSON.stringify(null));
  await context.sync();
  return Object.keys(record.cells).length;
}

/** Put back whatever fill each marked cell had before. */
export async function clearMarks(): Promise<OperationResult> {
  return runExcel(async (context) => {
    const cleared = await clearMarksIn(context);
    return ok(
      cleared === 0 ? "There are no marks to clear." : `Cleared ${plural(cleared, "mark")}.`
    );
  });
}

/** Select a cell (or a whole group of them) so Excel scrolls it into view. */
export async function goToIssue(sheetName: string, issue: Issue): Promise<OperationResult> {
  return runExcel(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    sheet.activate();
    sheet.getRangeByIndexes(issue.row, issue.column, 1, 1).select();
    await context.sync();
    return ok(`Selected ${issue.address} on "${sheetName}".`);
  });
}

export async function goToGroup(sheetName: string, group: IssueGroup): Promise<OperationResult> {
  return runExcel(async (context) => {
    const rows = group.issues.map((issue) => issue.row);
    const first = Math.min(...rows);
    const sheet = context.workbook.worksheets.getItem(sheetName);
    sheet.activate();
    sheet
      .getRangeByIndexes(first, group.issues[0].column, Math.max(...rows) - first + 1, 1)
      .select();
    await context.sync();
    return ok(`Selected ${plural(group.issues.length, "cell")} on "${sheetName}".`);
  });
}

// ---------------------------------------------------------------------------
// Fixing, with an undo we control
// ---------------------------------------------------------------------------

interface UndoEntry {
  sheet: string;
  description: string;
  cells: Array<{
    row: number;
    column: number;
    value: CellValue;
    formula: string;
    numberFormat: string;
  }>;
}

/** Fixes applied this session, newest last. Undo walks back through them. */
const undoStack: UndoEntry[] = [];

export function canUndo(): boolean {
  return undoStack.length > 0;
}

export function lastFixDescription(): string | null {
  return undoStack.length > 0 ? undoStack[undoStack.length - 1].description : null;
}

/**
 * Snapshots cells by address before something else changes them, so "undo"
 * covers that change too.
 *
 * Anything in the add-in that writes to cells the user did not type into should
 * call this first. It reads the current contents into the same undo stack the
 * Review fixes use, which is why there is one "undo" rather than one per
 * feature. The caller writes afterwards, in the same Excel.run batch.
 */
export async function rememberCells(
  context: Excel.RequestContext,
  sheetName: string,
  addresses: readonly string[],
  description: string
): Promise<void> {
  if (addresses.length === 0) return;
  const sheet = context.workbook.worksheets.getItem(sheetName);
  const ranges = addresses.map((address) => {
    const range = sheet.getRange(address);
    range.load(["values", "formulas", "numberFormat", "rowIndex", "columnIndex"]);
    return range;
  });
  await context.sync();

  const entry: UndoEntry = { sheet: sheetName, description, cells: [] };
  for (const range of ranges) {
    entry.cells.push({
      row: range.rowIndex,
      column: range.columnIndex,
      value: (range.values as CellValue[][])[0][0],
      formula: String((range.formulas as unknown[][])[0][0] ?? ""),
      numberFormat: String((range.numberFormat as unknown[][])[0][0] ?? "General"),
    });
  }
  undoStack.push(entry);
}

/**
 * Apply the fixes attached to these issues. Cells are snapshotted first, so
 * `undoLastFix` can put them back exactly as they were.
 */
export async function applyFixes(
  sheetName: string,
  issues: readonly Issue[],
  description: string
): Promise<OperationResult> {
  const fixable = issues.filter((issue) => issue.fix);
  if (fixable.length === 0) {
    return fail("Those need a decision from you - I won't change them on a guess.");
  }

  return runExcel(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);

    const before = fixable.map((issue) => {
      const range = sheet.getRangeByIndexes(issue.row, issue.column, 1, 1);
      range.load(["values", "formulas", "numberFormat"]);
      return range;
    });
    await context.sync();

    const entry: UndoEntry = { sheet: sheetName, description, cells: [] };
    fixable.forEach((issue, i) => {
      entry.cells.push({
        row: issue.row,
        column: issue.column,
        value: (before[i].values as CellValue[][])[0][0],
        formula: String((before[i].formulas as unknown[][])[0][0] ?? ""),
        numberFormat: String((before[i].numberFormat as unknown[][])[0][0] ?? "General"),
      });

      const range = sheet.getRangeByIndexes(issue.row, issue.column, 1, 1);
      const fix = issue.fix!;
      if (fix.type === "setFormula") {
        range.formulas = [[fix.formula]];
      } else {
        range.values = [[fix.value]];
        if (fix.numberFormat) range.numberFormat = [[fix.numberFormat]];
      }
    });

    await context.sync();
    undoStack.push(entry);
    return ok(`Fixed ${plural(fixable.length, "cell")} on "${sheetName}".`, [
      'Say "undo" to put them back.',
    ]);
  });
}

/** Put the last set of fixed cells back exactly as they were. */
export async function undoLastFix(): Promise<OperationResult> {
  const entry = undoStack[undoStack.length - 1];
  if (!entry) {
    return fail("There's nothing from this session to undo.");
  }
  return runExcel(async (context) => {
    const sheet = context.workbook.worksheets.getItemOrNullObject(entry.sheet);
    await context.sync();
    if (sheet.isNullObject) {
      undoStack.pop();
      return fail(`The sheet "${entry.sheet}" isn't in this workbook any more.`);
    }
    for (const cell of entry.cells) {
      const range = sheet.getRangeByIndexes(cell.row, cell.column, 1, 1);
      // Writing the formula back covers both cases: a constant is just a
      // formula-free value, and Excel stores it as typed.
      range.formulas = [[cell.formula === "" ? cell.value : cell.formula]];
      range.numberFormat = [[cell.numberFormat]];
    }
    sheet.activate();
    await context.sync();
    undoStack.pop();
    return ok(`Put back ${plural(entry.cells.length, "cell")}: ${entry.description}.`);
  });
}

// ---------------------------------------------------------------------------
// Ignoring
// ---------------------------------------------------------------------------

/** Stop reporting these issues on this sheet. Kept in the workbook. */
export async function ignoreIssues(
  sheetName: string,
  issues: readonly Issue[]
): Promise<OperationResult> {
  return runExcel(async (context) => {
    const ignored = await readSetting<string[]>(context, IGNORED_KEY, []);
    const next = Array.from(
      new Set([...ignored, ...issues.map((issue) => `${sheetName}|${issue.id}`)])
    );
    writeSetting(context, IGNORED_KEY, next);
    await context.sync();
    return ok(`Ignoring ${plural(issues.length, "issue")} from now on.`, [
      'Say "unignore all" to bring them back.',
    ]);
  });
}

export async function clearIgnored(): Promise<OperationResult> {
  return runExcel(async (context) => {
    const ignored = await readSetting<string[]>(context, IGNORED_KEY, []);
    writeSetting(context, IGNORED_KEY, []);
    await context.sync();
    return ok(
      ignored.length === 0
        ? "Nothing was being ignored."
        : `Bringing back ${plural(ignored.length, "issue")}.`
    );
  });
}

// ---------------------------------------------------------------------------
// What the user is looking at
// ---------------------------------------------------------------------------

export interface SelectionInfo {
  sheet: string;
  address: string;
  row: number;
  column: number;
  rowCount: number;
  columnCount: number;
}

/** Where the cursor is, so the pane can talk about the column you're in. */
export async function currentSelection(): Promise<SelectionInfo | null> {
  try {
    return await Excel.run(async (context) => {
      const range = await loadSelectedRange(context);
      return {
        sheet: range.worksheet.name,
        address: range.address.includes("!") ? range.address.split("!")[1] : range.address,
        row: range.rowIndex,
        column: range.columnIndex,
        rowCount: range.rowCount,
        columnCount: range.columnCount,
      };
    });
  } catch {
    return null;
  }
}

/** The issues inside the current selection, for "what's wrong with this bit?". */
export function issuesInSelection(issues: readonly Issue[], selection: SelectionInfo): Issue[] {
  // A single cell means "the column I'm in", which is what people mean when
  // they click a heading and ask what's wrong.
  const wholeColumn = selection.rowCount === 1 && selection.columnCount === 1;
  return issues.filter((issue) => {
    if (wholeColumn) return issue.column === selection.column;
    return (
      issue.row >= selection.row &&
      issue.row < selection.row + selection.rowCount &&
      issue.column >= selection.column &&
      issue.column < selection.column + selection.columnCount
    );
  });
}

/** A one-line description of a range, for messages: "Orders!C2:C40". */
export function describeSelection(selection: SelectionInfo): string {
  return buildAddress(
    selection.sheet,
    selection.row,
    selection.column,
    selection.rowCount,
    selection.columnCount
  );
}

/** The column letter a cell sits in, for messages. */
export function columnOf(issue: Issue): string {
  return columnLetter(issue.column);
}

/** Wrap an Office failure the same way the rest of the add-in does. */
export function describeFailure(error: unknown): OperationResult {
  return fail(describeExcelError(error));
}
