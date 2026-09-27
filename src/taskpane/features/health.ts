/**
 * Why this workbook is slow, and what in it is fragile.
 *
 * A slow file is one of the most common complaints in Excel, and the causes are
 * well known and entirely detectable: references that cover a whole column when
 * the data stops at row 800, functions that recalculate on every keystroke,
 * conditional formatting spread over tens of thousands of cells, and a used
 * range that Excel believes stretches to row 40,000 because someone once
 * formatted that far. None of it is visible to the person suffering from it.
 *
 * The fragile half matters more. Inheriting a workbook means inheriting its
 * assumptions, and the dangerous ones are invisible too: a rate typed into a
 * formula, a lookup that finds its answer by counting columns, a figure that
 * comes from a file somebody else controls.
 *
 * Formulas are examined once per distinct formula rather than once per cell. A
 * column of 40,000 copies of the same formula is one idea repeated, not 40,000
 * problems, and treating it that way keeps a whole-workbook scan quick.
 *
 * Everything here is pure and unit tested in tests/health.test.ts.
 */

import { columnLetter, plural } from "../shared/excelHelpers";
import { explainFormula } from "./explain";
import type { Grid } from "../shared/types";

export type HealthKind =
  | "wholeColumn"
  | "volatile"
  | "hardcodedNumber"
  | "lookupByPosition"
  | "deepNesting"
  | "divideRisk"
  | "externalLink"
  | "usedRangeBloat"
  | "conditionalFormatBloat"
  | "inconsistentColumn";

export type HealthSeverity = "slow" | "fragile" | "tidy";

export interface HealthFinding {
  kind: HealthKind;
  severity: HealthSeverity;
  sheet: string;
  /** One line naming the problem, in the user's terms. */
  title: string;
  /** What it costs them, or why it is a risk. */
  detail: string;
  /** How many cells or objects are involved. */
  count: number;
  /** A few example addresses, for going to look. */
  examples: string[];
  /** Set when MEx can repair it. */
  fix?: HealthFix;
}

export type HealthFix =
  | { kind: "tightenRanges"; sheet: string; cells: Array<{ address: string; formula: string }> }
  | { kind: "clearBeyondData"; sheet: string; firstRow: number; firstColumn: number };

/** What one sheet looks like to the scanner. */
export interface SheetSnapshot {
  name: string;
  /** Formulas of the used range, row-major. Empty strings for non-formula cells. */
  formulas: string[][];
  /** Values of the used range, to find where the data really stops. */
  values: Grid;
  /** Where the used range starts in the sheet (0-based). */
  origin: { row: number; column: number };
  /** How many conditional formats the sheet carries, when the host could say. */
  conditionalFormats?: number;
}

/** A sheet Excel thinks is this much bigger than its data before we mention it. */
const BLOAT_ROWS = 200;
const BLOAT_FACTOR = 2;

/** Conditional formatting beyond this many rules starts to cost real time. */
const HEAVY_CONDITIONAL_FORMATS = 20;

/** How many example addresses to keep per finding. */
const EXAMPLES = 5;

function isBlank(value: unknown): boolean {
  return value === null || value === undefined || value === "";
}

/** The last row and column that actually hold something. */
export function dataExtent(values: Grid): { rows: number; columns: number } {
  let rows = 0;
  let columns = 0;
  values.forEach((row, rowIndex) => {
    row.forEach((cell, columnIndex) => {
      if (isBlank(cell)) return;
      if (rowIndex + 1 > rows) rows = rowIndex + 1;
      if (columnIndex + 1 > columns) columns = columnIndex + 1;
    });
  });
  return { rows, columns };
}

/** A formula with its row and column numbers stripped, so copies collapse into one. */
export function formulaShape(formula: string): string {
  return formula.replace(/\$?\b([A-Za-z]{1,3})\$?(\d{1,7})\b/g, "$1#");
}

interface FormulaUse {
  formula: string;
  addresses: string[];
  count: number;
}

/** Groups the sheet's formulas by their exact text. */
function collectFormulas(snapshot: SheetSnapshot): FormulaUse[] {
  const seen = new Map<string, FormulaUse>();
  snapshot.formulas.forEach((row, rowIndex) => {
    row.forEach((formula, columnIndex) => {
      if (typeof formula !== "string" || !formula.startsWith("=")) return;
      const address = `${columnLetter(snapshot.origin.column + columnIndex)}${snapshot.origin.row + rowIndex + 1}`;
      const existing = seen.get(formula);
      if (existing) {
        existing.count += 1;
        if (existing.addresses.length < EXAMPLES) existing.addresses.push(address);
      } else {
        seen.set(formula, { formula, addresses: [address], count: 1 });
      }
    });
  });
  return [...seen.values()];
}

/** A formula that reads another workbook looks like ='[Budget.xlsx]Sheet1'!A1 */
const EXTERNAL = /\[[^\]]+\.xls[xmb]?\]/i;

interface Bucket {
  count: number;
  examples: string[];
  cells: Array<{ address: string; formula: string }>;
}

function bucket(): Bucket {
  return { count: 0, examples: [], cells: [] };
}

function record(into: Bucket, use: FormulaUse): void {
  into.count += use.count;
  for (const address of use.addresses) {
    if (into.examples.length < EXAMPLES) into.examples.push(address);
  }
  into.cells.push({ address: use.addresses[0], formula: use.formula });
}

/**
 * Looks over one sheet. The findings are per sheet so the report can say where
 * to go, and are counted by cell so "3 formulas" and "12,000 formulas" read
 * differently - which they should, because they cost differently.
 */
export function checkSheet(snapshot: SheetSnapshot): HealthFinding[] {
  const findings: HealthFinding[] = [];
  const uses = collectFormulas(snapshot);

  const wholeColumn = bucket();
  const volatile = bucket();
  const hardcoded = bucket();
  const positional = bucket();
  const nested = bucket();
  const divide = bucket();
  const external = bucket();

  for (const use of uses) {
    if (EXTERNAL.test(use.formula)) record(external, use);
    // One parse per distinct formula, however many cells repeat it.
    for (const warning of explainFormula(use.formula).warnings) {
      switch (warning.kind) {
        case "wholeColumn":
          record(wholeColumn, use);
          break;
        case "volatile":
          record(volatile, use);
          break;
        case "hardcodedNumber":
          record(hardcoded, use);
          break;
        case "lookupByPosition":
          record(positional, use);
          break;
        case "deepNesting":
          record(nested, use);
          break;
        case "divideRisk":
          record(divide, use);
          break;
      }
    }
  }

  const extent = dataExtent(snapshot.values);

  if (wholeColumn.count > 0) {
    findings.push({
      kind: "wholeColumn",
      severity: "slow",
      sheet: snapshot.name,
      title: `${plural(wholeColumn.count, "formula")} read whole columns`,
      detail:
        extent.rows > 0
          ? `Each one looks at every row of a column when the data stops at row ${(snapshot.origin.row + extent.rows).toLocaleString("en-US")}. Excel recalculates all of it, and anything typed below the data quietly joins the total.`
          : "Each one looks at every row of a column, so Excel recalculates far more than it needs to.",
      count: wholeColumn.count,
      examples: wholeColumn.examples,
      fix: {
        kind: "tightenRanges",
        sheet: snapshot.name,
        cells: wholeColumn.cells,
      },
    });
  }

  if (volatile.count > 0) {
    findings.push({
      kind: "volatile",
      severity: "slow",
      sheet: snapshot.name,
      title: `${plural(volatile.count, "formula")} never settle`,
      detail:
        "They use TODAY, NOW, OFFSET, INDIRECT or RAND, which recalculate on every edit anywhere in the workbook - and everything depending on them recalculates too.",
      count: volatile.count,
      examples: volatile.examples,
    });
  }

  if (external.count > 0) {
    findings.push({
      kind: "externalLink",
      severity: "fragile",
      sheet: snapshot.name,
      title: `${plural(external.count, "formula")} read another workbook`,
      detail:
        "These numbers come from a file somebody else may rename, move or edit. If that happens the values here go stale or break, and nothing on this sheet will say so.",
      count: external.count,
      examples: external.examples,
    });
  }

  if (positional.count > 0) {
    findings.push({
      kind: "lookupByPosition",
      severity: "fragile",
      sheet: snapshot.name,
      title: `${plural(positional.count, "lookup")} count columns`,
      detail:
        "VLOOKUP and HLOOKUP find their answer by counting across. Insert a column in the source and every one of them silently returns the wrong field - no error, just different numbers.",
      count: positional.count,
      examples: positional.examples,
    });
  }

  if (hardcoded.count > 0) {
    findings.push({
      kind: "hardcodedNumber",
      severity: "fragile",
      sheet: snapshot.name,
      title: `${plural(hardcoded.count, "formula")} have numbers typed inside`,
      detail:
        "A rate or a price written into a formula is invisible to anyone reading the sheet, and has to be found and changed by hand when it moves. A named cell would show it.",
      count: hardcoded.count,
      examples: hardcoded.examples,
    });
  }

  if (nested.count > 0) {
    findings.push({
      kind: "deepNesting",
      severity: "tidy",
      sheet: snapshot.name,
      title: `${plural(nested.count, "formula")} nest IF three or more deep`,
      detail: "Hard to check and easy to break. IFS, or a small lookup table, would read plainly.",
      count: nested.count,
      examples: nested.examples,
    });
  }

  if (divide.count > 0) {
    findings.push({
      kind: "divideRisk",
      severity: "tidy",
      sheet: snapshot.name,
      title: `${plural(divide.count, "formula")} divide without a guard`,
      detail: "An empty or zero divisor shows #DIV/0!, which spreads to every total above it.",
      count: divide.count,
      examples: divide.examples,
    });
  }

  // Excel's used range grows when cells are formatted and does not shrink when
  // they are emptied, so a sheet can carry tens of thousands of blank rows.
  const usedRows = snapshot.values.length;
  const usedColumns = snapshot.values[0]?.length ?? 0;
  const spareRows = usedRows - extent.rows;
  const spareColumns = usedColumns - extent.columns;
  if (
    extent.rows > 0 &&
    (spareRows >= BLOAT_ROWS || usedRows >= extent.rows * BLOAT_FACTOR + BLOAT_ROWS)
  ) {
    findings.push({
      kind: "usedRangeBloat",
      severity: "slow",
      sheet: snapshot.name,
      title: `Excel thinks this sheet is ${usedRows.toLocaleString("en-US")} rows; the data stops at ${extent.rows.toLocaleString("en-US")}`,
      detail: `${spareRows.toLocaleString("en-US")} empty rows${spareColumns > 0 ? ` and ${spareColumns.toLocaleString("en-US")} empty columns` : ""} are still carried in the file. They make it bigger, slower to open, and send Ctrl+End miles past the data.`,
      count: spareRows,
      examples: [`${columnLetter(snapshot.origin.column)}${snapshot.origin.row + extent.rows + 1}`],
      fix: {
        kind: "clearBeyondData",
        sheet: snapshot.name,
        firstRow: snapshot.origin.row + extent.rows,
        firstColumn: snapshot.origin.column + extent.columns,
      },
    });
  }

  if ((snapshot.conditionalFormats ?? 0) > HEAVY_CONDITIONAL_FORMATS) {
    findings.push({
      kind: "conditionalFormatBloat",
      severity: "slow",
      sheet: snapshot.name,
      title: `${plural(snapshot.conditionalFormats ?? 0, "conditional formatting rule")}`,
      detail:
        "Every rule is re-evaluated as you scroll and as you type. Rules built by copying rows tend to multiply; merging them back into a few that cover whole ranges makes the sheet noticeably quicker.",
      count: snapshot.conditionalFormats ?? 0,
      examples: [],
    });
  }

  return findings;
}

export interface HealthReport {
  findings: HealthFinding[];
  sheetsScanned: number;
  /** Sheets left out because they were too big to read. */
  sheetsSkipped: string[];
  formulaCells: number;
}

export function checkWorkbook(
  snapshots: readonly SheetSnapshot[],
  skipped: readonly string[] = []
): HealthReport {
  const findings: HealthFinding[] = [];
  let formulaCells = 0;

  for (const snapshot of snapshots) {
    findings.push(...checkSheet(snapshot));
    for (const row of snapshot.formulas) {
      for (const formula of row) {
        if (typeof formula === "string" && formula.startsWith("=")) formulaCells += 1;
      }
    }
  }

  const order: Record<HealthSeverity, number> = { slow: 0, fragile: 1, tidy: 2 };
  findings.sort((a, b) => {
    if (order[a.severity] !== order[b.severity]) return order[a.severity] - order[b.severity];
    return b.count - a.count;
  });

  return {
    findings,
    sheetsScanned: snapshots.length,
    sheetsSkipped: [...skipped],
    formulaCells,
  };
}

/** One sentence for the chat: what is worth knowing before reading the list. */
export function summarizeHealth(report: HealthReport): string {
  if (report.findings.length === 0) {
    return report.sheetsScanned === 0
      ? "There was nothing to look at."
      : `Nothing to report across ${report.sheetsScanned} ${report.sheetsScanned === 1 ? "sheet" : "sheets"} - this workbook is in good shape.`;
  }
  const slow = report.findings.filter((finding) => finding.severity === "slow").length;
  const fragile = report.findings.filter((finding) => finding.severity === "fragile").length;
  const tidy = report.findings.filter((finding) => finding.severity === "tidy").length;

  const parts: string[] = [];
  if (slow > 0) parts.push(`${slow} making it slower`);
  if (fragile > 0) parts.push(`${fragile} that could break quietly`);
  if (tidy > 0) parts.push(`${tidy} worth tidying`);
  return `${report.findings.length} things to know about this workbook: ${parts.join(", ")}.`;
}

/**
 * Narrows a whole-column reference to the rows that actually hold data.
 *
 * `=SUM(B:B)` over 812 rows of data becomes `=SUM(B2:B812)`. The first row is
 * kept at the top of the data rather than row 1, because row 1 is usually a
 * heading and including it changes nothing except how the formula reads.
 */
export function tightenWholeColumns(
  formula: string,
  firstDataRow: number,
  lastDataRow: number
): string | null {
  if (lastDataRow < firstDataRow) return null;
  const pattern = /(\$?)([A-Za-z]{1,3})(\$?):(\$?)([A-Za-z]{1,3})(\$?)(?![0-9])/g;
  let changed = false;
  const out = formula.replace(
    pattern,
    (_match, d1: string, left: string, d2: string, d3: string, right: string) => {
      changed = true;
      return `${d1}${left}${d2}${firstDataRow}:${d3}${right}${lastDataRow}`;
    }
  );
  return changed ? out : null;
}
