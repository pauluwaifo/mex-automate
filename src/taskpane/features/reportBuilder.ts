/**
 * Report builder: mark parts of a report as "data zones", map an incoming data
 * source onto them by header name, and refresh the whole report in one click
 * while the surrounding structure - titles, totals, formulas, charts,
 * conditional formatting - stays exactly where it was.
 *
 * A zone is anchored by an Excel *name*: either a table name or a defined name.
 * Anchoring by name rather than by address is what lets a refresh grow or shrink
 * the data without the template losing track of it.
 *
 * Templates are saved into the workbook via `templateStore`, so a report file
 * carries its own refresh configuration.
 */

import {
  buildAddress,
  chunk,
  fail,
  normalizeHeader,
  ok,
  padRow,
  parseAddress,
  plural,
  runExcel,
} from "../shared/excelHelpers";
import { nowIso } from "../shared/templateStore";
import {
  CellValue,
  ColumnMapping,
  DataZone,
  DEFAULT_REFRESH_OPTIONS,
  Grid,
  OperationResult,
  RefreshOptions,
  ReportTemplate,
} from "../shared/types";
import { readWorkbookFile } from "../shared/workbookReader";
import { SourceTable } from "./merge";

/** Office.js payload ceiling, matching excelHelpers. */
const MAX_CELLS_PER_BATCH = 40000;

// ---------------------------------------------------------------------------
// Pure mapping logic
// ---------------------------------------------------------------------------

/** Leave a zone column untouched by mapping it to this. */
export const UNMAPPED = "";

/**
 * Pair each zone column with a source column of the same (normalized) header.
 * Any mapping the user has already set by hand is preserved.
 */
export function autoMapColumns(
  zoneHeaders: readonly string[],
  sourceHeaders: readonly string[],
  existing: readonly ColumnMapping[] = []
): ColumnMapping[] {
  const sourceByKey = new Map<string, string>();
  for (const header of sourceHeaders) {
    const key = normalizeHeader(header);
    if (key !== "" && !sourceByKey.has(key)) {
      sourceByKey.set(key, header);
    }
  }

  const manualByTarget = new Map<string, ColumnMapping>();
  for (const mapping of existing) {
    if (mapping.manual) {
      manualByTarget.set(mapping.targetHeader, mapping);
    }
  }

  return zoneHeaders.map((targetHeader) => {
    const manual = manualByTarget.get(targetHeader);
    if (manual) {
      // Only honour a manual choice while that source column still exists.
      if (manual.sourceHeader === UNMAPPED || sourceHeaders.includes(manual.sourceHeader)) {
        return manual;
      }
    }
    return {
      targetHeader,
      sourceHeader: sourceByKey.get(normalizeHeader(targetHeader)) ?? UNMAPPED,
      manual: false,
    };
  });
}

export interface ZoneGrid {
  /** One row per source row, one column per zone column. `null` where unmapped. */
  grid: Grid;
  /** Zero-based zone column offsets that carry data. */
  mappedColumns: number[];
  /** Zone columns with no source column behind them. */
  unmappedTargets: string[];
  /** Source columns that no zone column claimed. */
  unusedSources: string[];
}

/** Reshape a source table into the zone's column order. */
export function buildZoneGrid(
  zoneHeaders: readonly string[],
  mappings: readonly ColumnMapping[],
  source: SourceTable
): ZoneGrid {
  const sourceIndexByHeader = new Map<string, number>();
  source.headers.forEach((header, index) => {
    if (!sourceIndexByHeader.has(header)) {
      sourceIndexByHeader.set(header, index);
    }
  });

  const mappingByTarget = new Map(mappings.map((mapping) => [mapping.targetHeader, mapping]));

  const columnSources: number[] = [];
  const mappedColumns: number[] = [];
  const unmappedTargets: string[] = [];
  const usedSourceHeaders = new Set<string>();

  zoneHeaders.forEach((header, index) => {
    const mapping = mappingByTarget.get(header);
    const sourceIndex =
      mapping && mapping.sourceHeader !== UNMAPPED
        ? sourceIndexByHeader.get(mapping.sourceHeader)
        : undefined;

    if (sourceIndex === undefined) {
      columnSources.push(-1);
      unmappedTargets.push(header);
      return;
    }
    columnSources.push(sourceIndex);
    mappedColumns.push(index);
    usedSourceHeaders.add(mapping!.sourceHeader);
  });

  const grid: Grid = source.rows.map((row) =>
    columnSources.map((sourceIndex) => (sourceIndex < 0 ? null : (row[sourceIndex] ?? null)))
  );

  const unusedSources = source.headers.filter((header) => !usedSourceHeaders.has(header));

  return { grid, mappedColumns, unmappedTargets, unusedSources };
}

/** A template with no zones cannot refresh anything; say so before touching the workbook. */
export function validateTemplate(template: ReportTemplate): string[] {
  const problems: string[] = [];
  if (template.zones.length === 0) {
    problems.push("This template has no data zones yet.");
  }
  for (const zone of template.zones) {
    if (zone.headers.length === 0) {
      problems.push(`Zone "${zone.name}" has no columns recorded.`);
    }
  }
  return problems;
}

export function emptyTemplate(name: string): ReportTemplate {
  return {
    name,
    createdAt: nowIso(),
    updatedAt: nowIso(),
    zones: [],
    mappings: {},
    options: { ...DEFAULT_REFRESH_OPTIONS },
  };
}

// ---------------------------------------------------------------------------
// Discovering zones
// ---------------------------------------------------------------------------

export interface ZoneCandidate {
  name: string;
  kind: "table" | "namedRange";
  address: string;
  sheetName: string;
}

/** Tables and workbook-scoped named ranges the user can turn into data zones. */
export async function listZoneCandidates(): Promise<ZoneCandidate[]> {
  try {
    return await Excel.run(async (context) => {
      const tables = context.workbook.tables;
      tables.load("items/name,items/worksheet/name");
      const names = context.workbook.names;
      names.load("items/name,items/type,items/value");
      await context.sync();

      // Table addresses need a second round trip, so queue them all up first.
      const tableRanges = tables.items.map((table) => {
        const range = table.getRange();
        range.load("address");
        return { name: table.name, sheetName: table.worksheet.name, range };
      });
      await context.sync();

      const candidates: ZoneCandidate[] = tableRanges.map((entry) => ({
        name: entry.name,
        kind: "table" as const,
        address: entry.range.address,
        sheetName: entry.sheetName,
      }));

      for (const name of names.items) {
        if (name.type !== Excel.NamedItemType.range) {
          continue;
        }
        const address = String(name.value ?? "").replace(/^=/, "");
        try {
          const parsed = parseAddress(address);
          candidates.push({
            name: name.name,
            kind: "namedRange",
            address,
            sheetName: parsed.sheetName ?? "",
          });
        } catch {
          // A name pointing at a formula rather than a range is not a zone.
        }
      }

      return candidates.sort((a, b) => a.name.localeCompare(b.name));
    });
  } catch {
    return [];
  }
}

/** Read a zone's current headers and address from the workbook. */
export async function loadZone(
  name: string,
  kind: "table" | "namedRange",
  hasHeaderRow: boolean
): Promise<DataZone | null> {
  try {
    return await Excel.run(async (context) => {
      const range =
        kind === "table"
          ? context.workbook.tables.getItem(name).getRange()
          : context.workbook.names.getItem(name).getRange();
      range.load(["address", "rowCount", "columnCount", "worksheet/name"]);
      await context.sync();

      const headerRange = range.getRow(0);
      headerRange.load("values");
      await context.sync();

      const headerValues = (headerRange.values as Grid)[0] ?? [];
      const headers = headerValues.map((cell, index) => {
        const text = String(cell ?? "").trim();
        return text === "" ? `Column ${index + 1}` : text;
      });

      return {
        name,
        kind,
        address: range.address,
        sheetName: range.worksheet.name,
        headers,
        hasHeaderRow,
      };
    });
  } catch {
    return null;
  }
}

/**
 * Turn the current selection into a named data zone. The defined name is what
 * the template stores, so later refreshes find the zone even after it moves.
 */
export async function createZoneFromSelection(
  zoneName: string,
  hasHeaderRow: boolean
): Promise<{ result: OperationResult; zone: DataZone | null }> {
  const trimmed = zoneName.trim();
  if (!/^[A-Za-z_][A-Za-z0-9_.]*$/.test(trimmed)) {
    return {
      result: fail(
        "A zone name must start with a letter or underscore and contain only letters, digits, dots or underscores."
      ),
      zone: null,
    };
  }

  let created: DataZone | null = null;
  const result = await runExcel(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load(["address", "rowCount", "columnCount", "values", "worksheet/name"]);
    await context.sync();

    if (range.rowCount < (hasHeaderRow ? 2 : 1)) {
      return fail(
        hasHeaderRow
          ? "Select the header row together with at least one data row."
          : "Select at least one row of data."
      );
    }

    // Re-adding an existing name repoints it, which is what a redefinition means.
    const existing = context.workbook.names.getItemOrNullObject(trimmed);
    await context.sync();
    if (!existing.isNullObject) {
      existing.delete();
      await context.sync();
    }
    context.workbook.names.add(trimmed, range);
    await context.sync();

    const values = range.values as Grid;
    const headers = hasHeaderRow
      ? (values[0] ?? []).map((cell, index) => {
          const text = String(cell ?? "").trim();
          return text === "" ? `Column ${index + 1}` : text;
        })
      : Array.from({ length: range.columnCount }, (_unused, index) => `Column ${index + 1}`);

    created = {
      name: trimmed,
      kind: "namedRange",
      address: range.address,
      sheetName: range.worksheet.name,
      headers,
      hasHeaderRow,
    };

    return ok(
      `Zone "${trimmed}" now covers ${range.address} (${plural(headers.length, "column")}).`
    );
  });

  return { result, zone: created };
}

// ---------------------------------------------------------------------------
// Reading data sources
// ---------------------------------------------------------------------------

/** Use a worksheet in this workbook as the incoming data for a zone. */
export async function readSheetAsSource(sheetName: string): Promise<SourceTable | null> {
  try {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(sheetName);
      const used = sheet.getUsedRangeOrNullObject(true);
      used.load("values");
      await context.sync();

      if (used.isNullObject) {
        return null;
      }
      const grid = used.values as Grid;
      if (grid.length < 2) {
        return null;
      }
      const headers = grid[0].map((cell, index) => {
        const text = String(cell ?? "").trim();
        return text === "" ? `Column ${index + 1}` : text;
      });
      return {
        label: sheetName,
        headers,
        rows: grid.slice(1).map((row) => padRow(row, headers.length)),
      };
    });
  } catch {
    return null;
  }
}

/** Use a file picked from disk as the incoming data. Parsed locally, never uploaded. */
export async function readFileAsSources(file: File): Promise<SourceTable[]> {
  const workbook = await readWorkbookFile(file);
  return workbook.sheets
    .filter((sheet) => sheet.grid.length > 1)
    .map((sheet) => {
      const headers = sheet.grid[0].map((cell, index) => {
        const text = String(cell ?? "").trim();
        return text === "" ? `Column ${index + 1}` : text;
      });
      return {
        label: workbook.sheets.length > 1 ? `${file.name} - ${sheet.name}` : file.name,
        headers,
        rows: sheet.grid.slice(1).map((row) => padRow(row, headers.length)),
        numberFormats: sheet.numberFormats.slice(1),
      };
    });
}

// ---------------------------------------------------------------------------
// The refresh engine
// ---------------------------------------------------------------------------

/** Write one zone column, chunked so a big refresh stays inside payload limits. */
async function writeColumn(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  rowIndex: number,
  columnIndex: number,
  values: readonly CellValue[]
): Promise<void> {
  const rowsPerBatch = MAX_CELLS_PER_BATCH;
  let offset = 0;

  for (const batch of chunk(values as CellValue[], rowsPerBatch)) {
    sheet.getRangeByIndexes(rowIndex + offset, columnIndex, batch.length, 1).values = batch.map(
      (cellValue) => [cellValue]
    ) as unknown[][];
    offset += batch.length;
    // eslint-disable-next-line office-addins/no-context-sync-in-loop -- deliberate: the loop exists to keep each payload under Office's size limit, so one sync per batch is required.
    await context.sync();
  }
}

interface ZoneGeometry {
  sheet: Excel.Worksheet;
  sheetName: string;
  /** Zero-based row of the zone's first row (the header row, when it has one). */
  topRow: number;
  columnIndex: number;
  columnCount: number;
  /** Zero-based row where data rows start. */
  bodyRow: number;
  /** How many data rows the zone currently holds. */
  bodyRowCount: number;
}

async function readGeometry(
  context: Excel.RequestContext,
  zone: DataZone
): Promise<ZoneGeometry | null> {
  // Check the anchor exists before navigating from it: a deleted table or name
  // would otherwise throw ItemNotFound partway through the batch.
  const anchor =
    zone.kind === "table"
      ? context.workbook.tables.getItemOrNullObject(zone.name)
      : context.workbook.names.getItemOrNullObject(zone.name);
  await context.sync();

  if (anchor.isNullObject) {
    return null;
  }

  const range = anchor.getRange();
  range.load(["address", "rowIndex", "columnIndex", "rowCount", "columnCount", "worksheet/name"]);
  await context.sync();

  const headerRows = zone.hasHeaderRow ? 1 : 0;
  return {
    sheet: context.workbook.worksheets.getItem(range.worksheet.name),
    sheetName: range.worksheet.name,
    topRow: range.rowIndex,
    columnIndex: range.columnIndex,
    columnCount: range.columnCount,
    bodyRow: range.rowIndex + headerRows,
    bodyRowCount: Math.max(0, range.rowCount - headerRows),
  };
}

/** Which body columns hold formulas, so a refresh can leave them alone. */
async function readFormulaColumns(
  context: Excel.RequestContext,
  geometry: ZoneGeometry
): Promise<Set<number>> {
  if (geometry.bodyRowCount === 0) {
    return new Set();
  }
  const firstBodyRow = geometry.sheet.getRangeByIndexes(
    geometry.bodyRow,
    geometry.columnIndex,
    1,
    geometry.columnCount
  );
  firstBodyRow.load("formulas");
  await context.sync();

  const formulas = (firstBodyRow.formulas as string[][])[0] ?? [];
  const formulaColumns = new Set<number>();
  formulas.forEach((formula, index) => {
    if (typeof formula === "string" && formula.startsWith("=")) {
      formulaColumns.add(index);
    }
  });
  return formulaColumns;
}

/** Grow or shrink a zone's body to exactly `targetRows` rows. */
async function resizeBody(
  context: Excel.RequestContext,
  zone: DataZone,
  geometry: ZoneGeometry,
  targetRows: number
): Promise<void> {
  const delta = targetRows - geometry.bodyRowCount;
  if (delta === 0) {
    return;
  }

  if (delta > 0) {
    if (zone.kind === "table") {
      // Adding table rows makes Excel extend calculated columns and banding for us.
      const blank = Array.from({ length: delta }, () =>
        new Array<CellValue>(geometry.columnCount).fill(null)
      );
      const table = context.workbook.tables.getItem(zone.name);

      for (const batch of chunk(
        blank,
        Math.max(1, Math.floor(MAX_CELLS_PER_BATCH / geometry.columnCount))
      )) {
        // Excel accepts null for "leave this cell empty", but the typings only
        // admit string | number | boolean. Empty strings are not equivalent here:
        // they would make ISBLANK() false and quietly change report formulas.
        table.rows.add(undefined, batch as unknown as (string | number | boolean)[][]);
        // eslint-disable-next-line office-addins/no-context-sync-in-loop -- deliberate: the loop exists to keep each payload under Office's size limit, so one sync per batch is required.
        await context.sync();
      }
    } else {
      // A plain range grows by inserting cells, so anything below is pushed down
      // rather than overwritten.
      geometry.sheet
        .getRangeByIndexes(
          geometry.bodyRow + geometry.bodyRowCount,
          geometry.columnIndex,
          delta,
          geometry.columnCount
        )
        .insert(Excel.InsertShiftDirection.down);
      await context.sync();
    }
    return;
  }

  const surplus = -delta;
  geometry.sheet
    .getRangeByIndexes(
      geometry.bodyRow + targetRows,
      geometry.columnIndex,
      surplus,
      geometry.columnCount
    )
    .delete(Excel.DeleteShiftDirection.up);
  await context.sync();
}

/** Re-point a defined name at the zone's new extent after a resize. */
async function redefineName(
  context: Excel.RequestContext,
  zone: DataZone,
  geometry: ZoneGeometry,
  bodyRows: number
): Promise<string> {
  const totalRows = bodyRows + (zone.hasHeaderRow ? 1 : 0);
  const range = geometry.sheet.getRangeByIndexes(
    geometry.topRow,
    geometry.columnIndex,
    Math.max(totalRows, 1),
    geometry.columnCount
  );

  const existing = context.workbook.names.getItemOrNullObject(zone.name);
  await context.sync();
  if (!existing.isNullObject) {
    existing.delete();
    await context.sync();
  }
  context.workbook.names.add(zone.name, range);
  await context.sync();

  return buildAddress(
    geometry.sheetName,
    geometry.topRow,
    geometry.columnIndex,
    Math.max(totalRows, 1),
    geometry.columnCount
  );
}

export interface ZoneRefreshReport {
  zoneName: string;
  rowsWritten: number;
  previousRows: number;
  unmappedTargets: string[];
  unusedSources: string[];
  formulaColumnsPreserved: number;
  address: string;
}

/** Repopulate a single zone from a source table. */
async function refreshZoneInContext(
  context: Excel.RequestContext,
  zone: DataZone,
  mappings: readonly ColumnMapping[],
  source: SourceTable,
  options: RefreshOptions
): Promise<ZoneRefreshReport> {
  const geometry = await readGeometry(context, zone);
  if (!geometry) {
    throw new Error(
      `Data zone "${zone.name}" no longer exists in this workbook. Re-define it on the Reports tab.`
    );
  }

  const { grid, mappedColumns, unmappedTargets, unusedSources } = buildZoneGrid(
    zone.headers,
    mappings,
    source
  );

  const formulaColumns = await readFormulaColumns(context, geometry);
  const previousRows = geometry.bodyRowCount;
  const targetRows = grid.length;

  if (options.clearExistingRows && previousRows > 0) {
    // Clear values only; number formats, fills and borders stay put.
    const columnsToClear = mappedColumns.filter((column) => !formulaColumns.has(column));
    for (const column of columnsToClear) {
      geometry.sheet
        .getRangeByIndexes(geometry.bodyRow, geometry.columnIndex + column, previousRows, 1)
        .clear(options.preserveFormatting ? Excel.ClearApplyTo.contents : Excel.ClearApplyTo.all);
    }
    await context.sync();
  }

  await resizeBody(context, zone, geometry, Math.max(targetRows, 1));

  for (const column of mappedColumns) {
    if (formulaColumns.has(column)) {
      // A formula column recomputes itself; overwriting it would destroy the report.
      continue;
    }
    const columnValues = grid.map((row) => row[column] ?? null);
    await writeColumn(
      context,
      geometry.sheet,
      geometry.bodyRow,
      geometry.columnIndex + column,
      columnValues
    );
  }

  let address = zone.address;
  if (zone.kind === "namedRange") {
    address = await redefineName(context, zone, geometry, Math.max(targetRows, 1));
  } else {
    const range = context.workbook.tables.getItem(zone.name).getRange();
    range.load("address");
    await context.sync();
    address = range.address;
  }

  return {
    zoneName: zone.name,
    rowsWritten: targetRows,
    previousRows,
    unmappedTargets,
    unusedSources,
    formulaColumnsPreserved: formulaColumns.size,
    address,
  };
}

/**
 * Refresh every zone in a template, then recalculate and refresh PivotTables so
 * charts and summaries pick the new data up.
 */
export async function refreshReport(
  template: ReportTemplate,
  sourcesByZone: Record<string, SourceTable>
): Promise<OperationResult> {
  const problems = validateTemplate(template);
  if (problems.length > 0) {
    return fail("This template is not ready to refresh.", problems);
  }

  const missing = template.zones.filter((zone) => !sourcesByZone[zone.name]);
  if (missing.length === template.zones.length) {
    return fail("Choose a data source for at least one zone before refreshing.");
  }

  return runExcel(async (context) => {
    const reports: ZoneRefreshReport[] = [];

    for (const zone of template.zones) {
      const source = sourcesByZone[zone.name];
      if (!source) {
        continue;
      }
      const mappings = template.mappings[zone.name] ?? autoMapColumns(zone.headers, source.headers);
      reports.push(await refreshZoneInContext(context, zone, mappings, source, template.options));
    }

    if (template.options.refreshPivotTables) {
      try {
        const pivotTables = context.workbook.pivotTables;
        pivotTables.load("items/name");
        await context.sync();
        for (const pivotTable of pivotTables.items) {
          pivotTable.refresh();
        }
        await context.sync();
      } catch {
        // PivotTable.refresh needs ExcelApi 1.8; skipping it is not fatal.
      }
    }

    if (template.options.recalculate) {
      context.workbook.application.calculate(Excel.CalculationType.full);
      await context.sync();
    }

    const details: string[] = [];
    let totalRows = 0;
    for (const report of reports) {
      totalRows += report.rowsWritten;
      details.push(
        `${report.zoneName}: ${plural(report.rowsWritten, "row")} (was ${report.previousRows}) at ${report.address}`
      );
      if (report.unmappedTargets.length > 0) {
        details.push(`  - left unchanged (no source column): ${report.unmappedTargets.join(", ")}`);
      }
      if (report.unusedSources.length > 0) {
        details.push(`  - source columns not used: ${report.unusedSources.join(", ")}`);
      }
    }

    if (missing.length > 0) {
      details.push(`Zones with no source this run: ${missing.map((zone) => zone.name).join(", ")}`);
    }

    return ok(
      `Refreshed ${plural(reports.length, "zone")} with ${plural(totalRows, "row")}.`,
      details
    );
  });
}

/** Refresh a single zone - the common case while building a template up. */
export async function refreshSingleZone(
  zone: DataZone,
  mappings: readonly ColumnMapping[],
  source: SourceTable,
  options: RefreshOptions = DEFAULT_REFRESH_OPTIONS
): Promise<OperationResult> {
  return refreshReport(
    {
      ...emptyTemplate("(ad hoc)"),
      zones: [zone],
      mappings: { [zone.name]: [...mappings] },
      options,
    },
    { [zone.name]: source }
  );
}
