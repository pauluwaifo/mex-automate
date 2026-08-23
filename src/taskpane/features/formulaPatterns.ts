/**
 * Formula patterns: extend one formula across a range with its relative
 * references adjusted, plus a library of one-click formula templates.
 *
 * The fill engine prefers `Range.copyFrom`, which hands the reference-shifting
 * to Excel itself and therefore always agrees with what Excel would do. On hosts
 * too old for that API (ExcelApi < 1.9) it falls back to `translateFormula`, the
 * pure A1 translator below, which is also what powers the "last cell will read"
 * preview.
 */

import {
  columnIndexFromLetter,
  columnLetter,
  fail,
  loadSelectedRange,
  ok,
  plural,
  runExcel,
} from "../shared/excelHelpers";
import { OperationResult } from "../shared/types";

// ---------------------------------------------------------------------------
// Pure A1 translation
// ---------------------------------------------------------------------------

/** Excel's grid limits; a reference pushed outside them becomes #REF!. */
const MAX_ROWS = 1048576;
const MAX_COLUMNS = 16384;

const CELL_REFERENCE = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})/;

function isIdentifierChar(char: string | undefined): boolean {
  return char !== undefined && /[A-Za-z0-9_.]/.test(char);
}

/**
 * Shift every *relative* A1 reference in a formula by the given row/column
 * delta, leaving `$`-anchored parts, string literals, quoted sheet names and
 * structured (`Table[Column]`) references untouched.
 *
 * Caveat: a defined name shaped like a cell reference (e.g. `Q1`) is
 * indistinguishable from a reference without Excel's name table, so it would be
 * shifted. This is why `copyFrom` is preferred whenever the host supports it.
 */
export function translateFormula(formula: string, rowDelta: number, columnDelta: number): string {
  if (!formula.startsWith("=")) {
    return formula;
  }

  let out = "";
  let i = 0;

  while (i < formula.length) {
    const char = formula[i];

    // String literal - copy verbatim, honouring "" escapes.
    if (char === '"') {
      const start = i;
      i += 1;
      while (i < formula.length) {
        if (formula[i] === '"') {
          if (formula[i + 1] === '"') {
            i += 2;
            continue;
          }
          i += 1;
          break;
        }
        i += 1;
      }
      out += formula.slice(start, i);
      continue;
    }

    // Quoted sheet name - copy verbatim; the reference after "!" still shifts.
    if (char === "'") {
      const start = i;
      i += 1;
      while (i < formula.length) {
        if (formula[i] === "'") {
          if (formula[i + 1] === "'") {
            i += 2;
            continue;
          }
          i += 1;
          break;
        }
        i += 1;
      }
      out += formula.slice(start, i);
      continue;
    }

    // Structured reference body - Excel adjusts these by name, not position.
    if (char === "[") {
      const end = formula.indexOf("]", i);
      const stop = end === -1 ? formula.length : end + 1;
      out += formula.slice(i, stop);
      i = stop;
      continue;
    }

    const match = CELL_REFERENCE.exec(formula.slice(i));
    const previous = i > 0 ? formula[i - 1] : undefined;
    const startsFresh = !isIdentifierChar(previous) && previous !== "$";

    if (match && startsFresh) {
      const [text, columnAnchor, columnPart, rowAnchor, rowPart] = match;
      const after = formula[i + text.length];
      // Rule out the look-alikes: "LOG10(" is a function call, "A1B" is part of a
      // longer name, "AB1!A2" starts with a sheet name, and "Table1[Amount]" is a
      // structured reference whose head is not a cell at all.
      const isReference =
        after !== "(" && after !== "!" && after !== "[" && !isIdentifierChar(after);

      if (isReference) {
        let columnIndex = columnIndexFromLetter(columnPart);
        let rowIndex = Number(rowPart) - 1;

        if (columnAnchor !== "$") {
          columnIndex += columnDelta;
        }
        if (rowAnchor !== "$") {
          rowIndex += rowDelta;
        }

        if (columnIndex < 0 || rowIndex < 0 || columnIndex >= MAX_COLUMNS || rowIndex >= MAX_ROWS) {
          out += "#REF!";
        } else {
          out += `${columnAnchor}${columnLetter(columnIndex)}${rowAnchor}${rowIndex + 1}`;
        }
        i += text.length;
        continue;
      }
    }

    out += char;
    i += 1;
  }

  return out;
}

/** Translate one formula across a rows x columns block, as a fill would. */
export function buildFilledFormulas(
  formula: string,
  rowCount: number,
  columnCount: number
): string[][] {
  const grid: string[][] = [];
  for (let row = 0; row < rowCount; row += 1) {
    const line: string[] = [];
    for (let column = 0; column < columnCount; column += 1) {
      line.push(translateFormula(formula, row, column));
    }
    grid.push(line);
  }
  return grid;
}

// ---------------------------------------------------------------------------
// Template library
// ---------------------------------------------------------------------------

export interface FormulaInput {
  key: string;
  label: string;
  /** Hint for the UI; all values arrive as strings. */
  kind: "column" | "range" | "text";
  placeholder: string;
  /** Filled in from the selection when the user leaves the field blank. */
  defaultFromSelection?: "valueColumn" | "keyColumn";
  help?: string;
}

/** What the builder knows about where the formula is going. */
export interface FormulaTargetContext {
  /** 1-based row of the first target cell. */
  firstRow: number;
  /** Column letter of the first target cell, e.g. "D". */
  firstColumn: string;
  /** 1-based row of the last target cell. */
  lastRow: number;
  rowCount: number;
}

export interface FormulaTemplate {
  id: string;
  name: string;
  category: "Aggregate" | "Lookup" | "Ranking" | "Text & dates";
  description: string;
  inputs: FormulaInput[];
  build: (values: Record<string, string>, target: FormulaTargetContext) => string;
}

/** The column immediately left of the target, a sensible default for "value column". */
function columnLeftOf(target: FormulaTargetContext): string {
  const index = columnIndexFromLetter(target.firstColumn);
  return columnLetter(Math.max(0, index - 1));
}

function value(values: Record<string, string>, key: string, fallback: string): string {
  const raw = (values[key] ?? "").trim();
  return raw === "" ? fallback : raw;
}

export const FORMULA_TEMPLATES: FormulaTemplate[] = [
  {
    id: "running-total",
    name: "Running total",
    category: "Aggregate",
    description: "Cumulative sum down a column. Anchors the start, so each row adds the one above.",
    inputs: [
      {
        key: "column",
        label: "Value column",
        kind: "column",
        placeholder: "e.g. C",
        defaultFromSelection: "valueColumn",
      },
    ],
    build: (values, target) => {
      const column = value(values, "column", columnLeftOf(target));
      return `=SUM($${column}$${target.firstRow}:${column}${target.firstRow})`;
    },
  },
  {
    id: "running-average",
    name: "Running average",
    category: "Aggregate",
    description: "Average of every value from the top of the column down to this row.",
    inputs: [
      {
        key: "column",
        label: "Value column",
        kind: "column",
        placeholder: "e.g. C",
        defaultFromSelection: "valueColumn",
      },
    ],
    build: (values, target) => {
      const column = value(values, "column", columnLeftOf(target));
      return `=AVERAGE($${column}$${target.firstRow}:${column}${target.firstRow})`;
    },
  },
  {
    id: "percent-of-total",
    name: "% of total",
    category: "Aggregate",
    description: "Each row as a share of the column total. Format the result as a percentage.",
    inputs: [
      {
        key: "column",
        label: "Value column",
        kind: "column",
        placeholder: "e.g. C",
        defaultFromSelection: "valueColumn",
      },
    ],
    build: (values, target) => {
      const column = value(values, "column", columnLeftOf(target));
      const span = `$${column}$${target.firstRow}:$${column}$${target.lastRow}`;
      return `=IFERROR(${column}${target.firstRow}/SUM(${span}),"")`;
    },
  },
  {
    id: "percent-change",
    name: "% change vs previous row",
    category: "Aggregate",
    description:
      "Period-over-period growth. The first row has nothing to compare against, so it stays blank.",
    inputs: [
      {
        key: "column",
        label: "Value column",
        kind: "column",
        placeholder: "e.g. C",
        defaultFromSelection: "valueColumn",
      },
    ],
    build: (values, target) => {
      const column = value(values, "column", columnLeftOf(target));
      const current = `${column}${target.firstRow}`;
      const previous = `${column}${target.firstRow - 1}`;
      return `=IFERROR((${current}-${previous})/${previous},"")`;
    },
  },
  {
    id: "subtotal-by-key",
    name: "Total by category",
    category: "Aggregate",
    description: "SUMIF: total the value column for every row sharing this row's category.",
    inputs: [
      { key: "keyColumn", label: "Category column", kind: "column", placeholder: "e.g. A" },
      { key: "valueColumn", label: "Value column", kind: "column", placeholder: "e.g. C" },
    ],
    build: (values, target) => {
      const key = value(values, "keyColumn", "A");
      const amount = value(values, "valueColumn", columnLeftOf(target));
      const keySpan = `$${key}$${target.firstRow}:$${key}$${target.lastRow}`;
      const valueSpan = `$${amount}$${target.firstRow}:$${amount}$${target.lastRow}`;
      return `=SUMIF(${keySpan},${key}${target.firstRow},${valueSpan})`;
    },
  },
  {
    id: "xlookup",
    name: "XLOOKUP",
    category: "Lookup",
    description: "Modern lookup. Needs Microsoft 365 - use INDEX/MATCH on older Excel.",
    inputs: [
      { key: "keyColumn", label: "Lookup value column", kind: "column", placeholder: "e.g. A" },
      {
        key: "lookupRange",
        label: "Where to look",
        kind: "range",
        placeholder: "e.g. Prices!$A$2:$A$500",
        help: "Use absolute refs ($) so the range does not drift as the formula fills.",
      },
      {
        key: "returnRange",
        label: "What to return",
        kind: "range",
        placeholder: "e.g. Prices!$B$2:$B$500",
      },
      { key: "notFound", label: "If not found", kind: "text", placeholder: "Not found" },
    ],
    build: (values, target) => {
      const key = value(values, "keyColumn", "A");
      const lookup = value(values, "lookupRange", "$A$2:$A$100");
      const result = value(values, "returnRange", "$B$2:$B$100");
      const missing = value(values, "notFound", "Not found");
      return `=XLOOKUP(${key}${target.firstRow},${lookup},${result},"${missing}")`;
    },
  },
  {
    id: "vlookup",
    name: "VLOOKUP",
    category: "Lookup",
    description: "Classic lookup wrapped in IFERROR. Works on every version of Excel.",
    inputs: [
      { key: "keyColumn", label: "Lookup value column", kind: "column", placeholder: "e.g. A" },
      {
        key: "table",
        label: "Lookup table",
        kind: "range",
        placeholder: "e.g. Prices!$A$2:$D$500",
        help: "The lookup key must be the table's first column.",
      },
      { key: "columnIndex", label: "Return column number", kind: "text", placeholder: "2" },
      { key: "notFound", label: "If not found", kind: "text", placeholder: "Not found" },
    ],
    build: (values, target) => {
      const key = value(values, "keyColumn", "A");
      const table = value(values, "table", "$A$2:$B$100");
      const index = value(values, "columnIndex", "2");
      const missing = value(values, "notFound", "Not found");
      return `=IFERROR(VLOOKUP(${key}${target.firstRow},${table},${index},FALSE),"${missing}")`;
    },
  },
  {
    id: "index-match",
    name: "INDEX / MATCH",
    category: "Lookup",
    description: "Lookup that works in any direction and survives inserted columns.",
    inputs: [
      { key: "keyColumn", label: "Lookup value column", kind: "column", placeholder: "e.g. A" },
      {
        key: "returnRange",
        label: "What to return",
        kind: "range",
        placeholder: "e.g. Prices!$B$2:$B$500",
      },
      {
        key: "lookupRange",
        label: "Where to look",
        kind: "range",
        placeholder: "e.g. Prices!$A$2:$A$500",
      },
      { key: "notFound", label: "If not found", kind: "text", placeholder: "Not found" },
    ],
    build: (values, target) => {
      const key = value(values, "keyColumn", "A");
      const result = value(values, "returnRange", "$B$2:$B$100");
      const lookup = value(values, "lookupRange", "$A$2:$A$100");
      const missing = value(values, "notFound", "Not found");
      return `=IFERROR(INDEX(${result},MATCH(${key}${target.firstRow},${lookup},0)),"${missing}")`;
    },
  },
  {
    id: "rank",
    name: "Rank",
    category: "Ranking",
    description: "Position of each value within its column, largest first.",
    inputs: [
      {
        key: "column",
        label: "Value column",
        kind: "column",
        placeholder: "e.g. C",
        defaultFromSelection: "valueColumn",
      },
    ],
    build: (values, target) => {
      const column = value(values, "column", columnLeftOf(target));
      const span = `$${column}$${target.firstRow}:$${column}$${target.lastRow}`;
      return `=IFERROR(RANK.EQ(${column}${target.firstRow},${span}),"")`;
    },
  },
  {
    id: "count-occurrences",
    name: "Count occurrences",
    category: "Ranking",
    description:
      "How many times this row's value appears in the column - the quick way to spot duplicates.",
    inputs: [
      {
        key: "column",
        label: "Column to count in",
        kind: "column",
        placeholder: "e.g. A",
        defaultFromSelection: "keyColumn",
      },
    ],
    build: (values, target) => {
      const column = value(values, "column", columnLeftOf(target));
      const span = `$${column}$${target.firstRow}:$${column}$${target.lastRow}`;
      return `=COUNTIF(${span},${column}${target.firstRow})`;
    },
  },
  {
    id: "join-columns",
    name: "Join columns",
    category: "Text & dates",
    description: "Combine several columns into one, skipping blanks.",
    inputs: [
      { key: "columns", label: "Columns to join", kind: "text", placeholder: "A,B,C" },
      { key: "separator", label: "Separator", kind: "text", placeholder: " " },
    ],
    build: (values, target) => {
      const columns = value(values, "columns", "A,B")
        .split(/[,\s]+/)
        .filter(Boolean)
        .map((column) => `${column.toUpperCase()}${target.firstRow}`);
      const separator = values.separator ?? " ";
      return `=TEXTJOIN("${separator}",TRUE,${columns.join(",")})`;
    },
  },
  {
    id: "month-key",
    name: "Year-month key",
    category: "Text & dates",
    description: 'Turn a date into a "2024-03" label - handy for grouping and pivots.',
    inputs: [
      {
        key: "column",
        label: "Date column",
        kind: "column",
        placeholder: "e.g. A",
        defaultFromSelection: "keyColumn",
      },
    ],
    build: (values, target) => {
      const column = value(values, "column", columnLeftOf(target));
      return `=IF(${column}${target.firstRow}="","",TEXT(${column}${target.firstRow},"yyyy-mm"))`;
    },
  },
  {
    id: "days-between",
    name: "Days between two dates",
    category: "Text & dates",
    description: "Whole days from a start date to an end date.",
    inputs: [
      { key: "startColumn", label: "Start date column", kind: "column", placeholder: "e.g. A" },
      { key: "endColumn", label: "End date column", kind: "column", placeholder: "e.g. B" },
    ],
    build: (values, target) => {
      const start = value(values, "startColumn", "A");
      const end = value(values, "endColumn", "B");
      const row = target.firstRow;
      return `=IF(OR(${start}${row}="",${end}${row}=""),"",${end}${row}-${start}${row})`;
    },
  },
];

export function findTemplate(id: string): FormulaTemplate | undefined {
  return FORMULA_TEMPLATES.find((template) => template.id === id);
}

// ---------------------------------------------------------------------------
// Office.js drivers
// ---------------------------------------------------------------------------

/** `Range.copyFrom` needs ExcelApi 1.9; older hosts fall back to translateFormula. */
function supportsCopyFrom(): boolean {
  try {
    return Office.context.requirements.isSetSupported("ExcelApi", "1.9");
  } catch {
    return false;
  }
}

interface SelectionInfo {
  address: string;
  sheetName: string;
  rowIndex: number;
  columnIndex: number;
  rowCount: number;
  columnCount: number;
  firstFormula: string;
}

async function readSelection(context: Excel.RequestContext): Promise<SelectionInfo> {
  const range = await loadSelectedRange(context);
  const formulas = range.formulas as string[][];
  return {
    address: range.address,
    sheetName: range.worksheet.name,
    rowIndex: range.rowIndex,
    columnIndex: range.columnIndex,
    rowCount: range.rowCount,
    columnCount: range.columnCount,
    firstFormula: String(formulas?.[0]?.[0] ?? ""),
  };
}

function targetContextFrom(selection: SelectionInfo): FormulaTargetContext {
  return {
    firstRow: selection.rowIndex + 1,
    firstColumn: columnLetter(selection.columnIndex),
    lastRow: selection.rowIndex + selection.rowCount,
    rowCount: selection.rowCount,
  };
}

/**
 * Fill the formula in the first cell of the selection across the rest of it.
 * Excel does the reference shifting via `copyFrom` wherever that API exists.
 */
export async function fillFormulaAcrossSelection(): Promise<OperationResult> {
  return runExcel(async (context) => {
    const selection = await readSelection(context);

    if (selection.rowCount * selection.columnCount < 2) {
      return fail(
        "Select the cell holding your formula together with the cells you want it filled into."
      );
    }
    if (!selection.firstFormula.startsWith("=")) {
      return fail(
        `The first cell of ${selection.address} does not contain a formula. Put your formula in the top-left cell of the selection.`
      );
    }

    const sheet = context.workbook.worksheets.getItem(selection.sheetName);
    const source = sheet.getRangeByIndexes(selection.rowIndex, selection.columnIndex, 1, 1);
    const target = sheet.getRangeByIndexes(
      selection.rowIndex,
      selection.columnIndex,
      selection.rowCount,
      selection.columnCount
    );

    if (supportsCopyFrom()) {
      target.copyFrom(source, Excel.RangeCopyType.formulas);
    } else {
      target.formulas = buildFilledFormulas(
        selection.firstFormula,
        selection.rowCount,
        selection.columnCount
      );
    }
    await context.sync();

    const filled = selection.rowCount * selection.columnCount - 1;
    return ok(`Filled ${plural(filled, "cell")} in ${selection.address}.`, [
      `Pattern: ${selection.firstFormula}`,
    ]);
  });
}

export interface FillPreview {
  available: boolean;
  message: string;
  sourceFormula: string;
  lastCellFormula: string;
  targetAddress: string;
}

/** Show what the bottom-right cell will contain, before anything is written. */
export async function previewFill(): Promise<FillPreview> {
  const empty: FillPreview = {
    available: false,
    message: "Select the cell with your formula plus the cells to fill.",
    sourceFormula: "",
    lastCellFormula: "",
    targetAddress: "",
  };

  try {
    return await Excel.run(async (context) => {
      const selection = await readSelection(context);
      if (selection.rowCount * selection.columnCount < 2) {
        return { ...empty, targetAddress: selection.address };
      }
      if (!selection.firstFormula.startsWith("=")) {
        return {
          ...empty,
          message: "The top-left cell of the selection has no formula in it.",
          targetAddress: selection.address,
        };
      }
      return {
        available: true,
        message: `Ready to fill ${plural(selection.rowCount * selection.columnCount - 1, "cell")}.`,
        sourceFormula: selection.firstFormula,
        lastCellFormula: translateFormula(
          selection.firstFormula,
          selection.rowCount - 1,
          selection.columnCount - 1
        ),
        targetAddress: selection.address,
      };
    });
  } catch {
    return { ...empty, message: "Could not read the current selection." };
  }
}

/**
 * Build a template's formula for the top-left cell of the selection and fill it
 * across the rest.
 */
export async function applyTemplate(
  templateId: string,
  values: Record<string, string>
): Promise<OperationResult> {
  const template = findTemplate(templateId);
  if (!template) {
    return fail(`Unknown formula template "${templateId}".`);
  }

  return runExcel(async (context) => {
    const selection = await readSelection(context);
    if (selection.columnCount > 1) {
      return fail("Select a single column of cells for the formula to fill down.");
    }

    const formula = template.build(values, targetContextFrom(selection));
    const sheet = context.workbook.worksheets.getItem(selection.sheetName);
    const first = sheet.getRangeByIndexes(selection.rowIndex, selection.columnIndex, 1, 1);
    first.formulas = [[formula]];

    if (selection.rowCount > 1) {
      const target = sheet.getRangeByIndexes(
        selection.rowIndex,
        selection.columnIndex,
        selection.rowCount,
        1
      );
      if (supportsCopyFrom()) {
        await context.sync();
        target.copyFrom(first, Excel.RangeCopyType.formulas);
      } else {
        target.formulas = buildFilledFormulas(formula, selection.rowCount, 1);
      }
    }
    await context.sync();

    return ok(`Inserted "${template.name}" into ${selection.address}.`, [`First cell: ${formula}`]);
  });
}

/** Render a template's formula without writing it, for the live preview box. */
export async function previewTemplate(
  templateId: string,
  values: Record<string, string>
): Promise<string> {
  const template = findTemplate(templateId);
  if (!template) {
    return "";
  }
  try {
    return await Excel.run(async (context) => {
      const selection = await readSelection(context);
      return template.build(values, targetContextFrom(selection));
    });
  } catch {
    // No live selection (e.g. the pane just opened) - show a representative row.
    return template.build(values, { firstRow: 2, firstColumn: "D", lastRow: 100, rowCount: 99 });
  }
}
