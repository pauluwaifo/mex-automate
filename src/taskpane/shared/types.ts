/** Shared domain types used across every feature module. */

export type CellValue = string | number | boolean | null;

/** A rectangular block of cell values, row-major. Mirrors `Excel.Range.values`. */
export type Grid = CellValue[][];

/** Uniform result shape so the task pane can render success/failure the same way everywhere. */
export interface OperationResult {
  ok: boolean;
  message: string;
  /** Optional per-item notes: skipped columns, unmatched headers, etc. */
  details?: string[];
}

/** A grid split into its header row and its body. */
export interface HeaderTable {
  headers: string[];
  rows: Grid;
}

/** Parsed form of an A1-style address such as `Sheet1!B2:D20`. */
export interface ParsedAddress {
  sheetName: string | null;
  rowIndex: number;
  columnIndex: number;
  rowCount: number;
  columnCount: number;
}

// ---------------------------------------------------------------------------
// Report builder
// ---------------------------------------------------------------------------

/** A region of a report template that gets cleared and repopulated on refresh. */
export interface DataZone {
  /** Excel table name or defined name that anchors the zone. */
  name: string;
  kind: "table" | "namedRange";
  /** Address captured when the zone was defined, e.g. `Data!A1:F200`. */
  address: string;
  sheetName: string;
  /** Header labels captured when the zone was defined. */
  headers: string[];
  /** True when the first row of the zone is a header row that must survive a refresh. */
  hasHeaderRow: boolean;
}

/** One column of an incoming data source wired to one column of a data zone. */
export interface ColumnMapping {
  /** Header in the incoming data source. Empty string means "leave this column blank". */
  sourceHeader: string;
  /** Header of the destination data zone column. */
  targetHeader: string;
  /** True when the pairing was chosen by the user rather than matched automatically. */
  manual: boolean;
}

export interface RefreshOptions {
  /** Clear the existing body rows before writing new ones. */
  clearExistingRows: boolean;
  /** Run a full recalculation after writing. */
  recalculate: boolean;
  /** Refresh every PivotTable in the workbook after writing. */
  refreshPivotTables: boolean;
  /** Write values only, leaving number formats/fills/borders untouched. */
  preserveFormatting: boolean;
}

export const DEFAULT_REFRESH_OPTIONS: RefreshOptions = {
  clearExistingRows: true,
  recalculate: true,
  refreshPivotTables: true,
  preserveFormatting: true,
};

export interface ReportTemplate {
  name: string;
  createdAt: string;
  updatedAt: string;
  zones: DataZone[];
  /** Column mappings keyed by zone name. */
  mappings: Record<string, ColumnMapping[]>;
  options: RefreshOptions;
}
