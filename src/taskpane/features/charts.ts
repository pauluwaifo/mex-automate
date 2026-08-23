/**
 * Charts and summary tables.
 *
 * Raw spreadsheet rows rarely chart well: a pie of 4,000 order lines is
 * unreadable. So this module does two related jobs, and the chart builder is
 * built on top of the summarizer:
 *
 *  1. `aggregateTable` groups rows by one column and aggregates others - the
 *     "Sum of Amount by Region" table you would otherwise build by hand.
 *  2. `createChart` optionally summarizes first, writes the resulting table to a
 *     sheet, then binds a chart to it.
 *
 * An Office.js chart must bind to a real range, so a summarized chart always
 * leaves its source table behind on the sheet. That is deliberate: the numbers
 * behind the picture stay visible and auditable.
 */

import {
  buildAddress,
  columnLetter,
  createSheetWithUniqueName,
  fail,
  isBlankRow,
  loadSelectedRange,
  loadSheetNames,
  ok,
  padRow,
  plural,
  runExcel,
  writeGridWithFormats,
} from "../shared/excelHelpers";
import { CellValue, Grid, OperationResult } from "../shared/types";

// ---------------------------------------------------------------------------
// Aggregation
// ---------------------------------------------------------------------------

export type Aggregation = "sum" | "average" | "count" | "countDistinct" | "min" | "max";

export const AGGREGATIONS: Array<{ value: Aggregation; label: string; numeric: boolean }> = [
  { value: "sum", label: "Sum", numeric: true },
  { value: "average", label: "Average", numeric: true },
  { value: "count", label: "Count", numeric: false },
  { value: "countDistinct", label: "Count unique", numeric: false },
  { value: "min", label: "Minimum", numeric: true },
  { value: "max", label: "Maximum", numeric: true },
];

export type SortOrder = "categoryAsc" | "categoryDesc" | "valueDesc" | "valueAsc" | "none";

export interface AggregateOptions {
  /** Zero-based index of the column whose values become the categories. */
  groupByColumn: number;
  /** Zero-based indexes of the columns being aggregated. */
  valueColumns: number[];
  aggregation: Aggregation;
  sort: SortOrder;
  /** Keep only this many categories, rolling the rest into one "Other" row. */
  topN: number | null;
}

export interface AggregatedTable {
  headers: string[];
  rows: Grid;
  /** How many source rows contributed. */
  sourceRows: number;
  /** How many distinct categories were found before any topN trimming. */
  categoryCount: number;
  /** Non-numeric cells ignored by a numeric aggregation such as Sum. */
  ignoredNonNumeric: number;
}

/** Label used for rows whose group-by cell is empty. */
export const BLANK_CATEGORY = "(blank)";
/** Label for the bucket that absorbs everything past topN. */
export const OTHER_CATEGORY = "Other";

function toNumber(value: CellValue): number | null {
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : null;
  }
  if (typeof value === "string") {
    // Tolerate numbers that arrived as text: "1,234.50", " 12 ", "(5)".
    const cleaned = value.replace(/[\s,]/g, "").replace(/^\((.*)\)$/, "-$1");
    if (cleaned === "" || !/^-?\d*\.?\d+(?:[eE][+-]?\d+)?$/.test(cleaned)) {
      return null;
    }
    const parsed = Number(cleaned);
    return Number.isFinite(parsed) ? parsed : null;
  }
  return null;
}

export function aggregationLabel(aggregation: Aggregation, columnHeader: string): string {
  const meta = AGGREGATIONS.find((item) => item.value === aggregation);
  return `${meta?.label ?? aggregation} of ${columnHeader}`;
}

interface Accumulator {
  sum: number;
  numericCount: number;
  nonBlankCount: number;
  min: number | null;
  max: number | null;
  distinct: Set<string>;
}

function newAccumulator(): Accumulator {
  return { sum: 0, numericCount: 0, nonBlankCount: 0, min: null, max: null, distinct: new Set() };
}

function finalize(accumulator: Accumulator, aggregation: Aggregation): CellValue {
  switch (aggregation) {
    case "sum":
      return accumulator.numericCount === 0 ? 0 : accumulator.sum;
    case "average":
      return accumulator.numericCount === 0 ? null : accumulator.sum / accumulator.numericCount;
    case "count":
      return accumulator.nonBlankCount;
    case "countDistinct":
      return accumulator.distinct.size;
    case "min":
      return accumulator.min;
    case "max":
      return accumulator.max;
    default:
      return null;
  }
}

/**
 * Group `rows` by one column and aggregate the others. Pure, so the whole
 * "Sum of Amount by Region" calculation is unit tested without Excel.
 */
export function aggregateTable(
  headers: readonly string[],
  rows: Grid,
  options: AggregateOptions
): AggregatedTable {
  const valueColumns = options.valueColumns.length > 0 ? options.valueColumns : [];
  const groups = new Map<string, Accumulator[]>();
  // Preserve the first-seen spelling of each category for display.
  const displayByKey = new Map<string, CellValue>();
  let ignoredNonNumeric = 0;
  let sourceRows = 0;

  const numeric = AGGREGATIONS.find((item) => item.value === options.aggregation)?.numeric ?? false;

  for (const row of rows) {
    if (isBlankRow(row)) {
      continue;
    }
    sourceRows += 1;

    const rawCategory = row[options.groupByColumn] ?? null;
    const key =
      rawCategory === null || String(rawCategory).trim() === ""
        ? BLANK_CATEGORY
        : String(rawCategory);

    let accumulators = groups.get(key);
    if (!accumulators) {
      accumulators = valueColumns.map(newAccumulator);
      groups.set(key, accumulators);
      displayByKey.set(key, key === BLANK_CATEGORY ? BLANK_CATEGORY : rawCategory);
    }

    valueColumns.forEach((column, index) => {
      const cell = row[column] ?? null;
      const accumulator = accumulators![index];

      if (cell !== null && String(cell).trim() !== "") {
        accumulator.nonBlankCount += 1;
        accumulator.distinct.add(String(cell));
      }

      const value = toNumber(cell);
      if (value === null) {
        if (numeric && cell !== null && String(cell).trim() !== "") {
          ignoredNonNumeric += 1;
        }
        return;
      }
      accumulator.sum += value;
      accumulator.numericCount += 1;
      accumulator.min = accumulator.min === null ? value : Math.min(accumulator.min, value);
      accumulator.max = accumulator.max === null ? value : Math.max(accumulator.max, value);
    });
  }

  const categoryCount = groups.size;

  let entries: SortableEntry[] = Array.from(groups.entries()).map(([key, accumulators]) => ({
    key,
    display: displayByKey.get(key) ?? key,
    values: accumulators.map((accumulator) => finalize(accumulator, options.aggregation)),
    accumulators,
  }));

  entries = sortEntries(entries, options.sort);

  // Roll everything past topN into a single "Other" row, so a pie stays readable.
  if (options.topN !== null && options.topN > 0 && entries.length > options.topN) {
    const kept = entries.slice(0, options.topN);
    const rest = entries.slice(options.topN);
    const merged = valueColumns.map((_column, index) => {
      const combined = newAccumulator();
      for (const entry of rest) {
        const accumulator = entry.accumulators[index];
        combined.sum += accumulator.sum;
        combined.numericCount += accumulator.numericCount;
        combined.nonBlankCount += accumulator.nonBlankCount;
        combined.min =
          combined.min === null
            ? accumulator.min
            : accumulator.min === null
              ? combined.min
              : Math.min(combined.min, accumulator.min);
        combined.max =
          combined.max === null
            ? accumulator.max
            : accumulator.max === null
              ? combined.max
              : Math.max(combined.max, accumulator.max);
        for (const item of accumulator.distinct) {
          combined.distinct.add(item);
        }
      }
      return finalize(combined, options.aggregation);
    });
    entries = [
      ...kept,
      { key: OTHER_CATEGORY, display: OTHER_CATEGORY, values: merged, accumulators: [] },
    ];
  }

  const categoryHeader = String(headers[options.groupByColumn] ?? "Category");
  const valueHeaders = valueColumns.map((column) =>
    aggregationLabel(options.aggregation, String(headers[column] ?? columnLetter(column)))
  );

  return {
    headers: [categoryHeader, ...valueHeaders],
    rows: entries.map((entry) => [entry.display, ...entry.values]),
    sourceRows,
    categoryCount,
    ignoredNonNumeric,
  };
}

interface SortableEntry {
  key: string;
  display: CellValue;
  values: CellValue[];
  accumulators: Accumulator[];
}

function sortEntries(entries: SortableEntry[], order: SortOrder): SortableEntry[] {
  if (order === "none") {
    return entries;
  }
  const sorted = [...entries];
  if (order === "categoryAsc" || order === "categoryDesc") {
    sorted.sort((a, b) =>
      String(a.display).localeCompare(String(b.display), undefined, { numeric: true })
    );
    return order === "categoryAsc" ? sorted : sorted.reverse();
  }
  // Value sorts key off the first aggregated column.
  sorted.sort((a, b) => {
    const left =
      typeof a.values[0] === "number" ? (a.values[0] as number) : Number.NEGATIVE_INFINITY;
    const right =
      typeof b.values[0] === "number" ? (b.values[0] as number) : Number.NEGATIVE_INFINITY;
    return left - right;
  });
  return order === "valueAsc" ? sorted : sorted.reverse();
}

// ---------------------------------------------------------------------------
// Chart types
// ---------------------------------------------------------------------------

export interface ChartTypeMeta {
  id: string;
  label: string;
  /** Matches an `Excel.ChartType` enum value. */
  excelType: string;
  group: "Column & bar" | "Line & area" | "Pie" | "Scatter";
  /** Pie and doughnut can only render one series. */
  singleSeriesOnly: boolean;
  /** Pie and doughnut have no category/value axes to title. */
  hasAxes: boolean;
}

export const CHART_TYPES: ChartTypeMeta[] = [
  {
    id: "columnClustered",
    label: "Clustered column",
    excelType: "ColumnClustered",
    group: "Column & bar",
    singleSeriesOnly: false,
    hasAxes: true,
  },
  {
    id: "columnStacked",
    label: "Stacked column",
    excelType: "ColumnStacked",
    group: "Column & bar",
    singleSeriesOnly: false,
    hasAxes: true,
  },
  {
    id: "barClustered",
    label: "Clustered bar",
    excelType: "BarClustered",
    group: "Column & bar",
    singleSeriesOnly: false,
    hasAxes: true,
  },
  {
    id: "barStacked",
    label: "Stacked bar",
    excelType: "BarStacked",
    group: "Column & bar",
    singleSeriesOnly: false,
    hasAxes: true,
  },
  {
    id: "line",
    label: "Line",
    excelType: "Line",
    group: "Line & area",
    singleSeriesOnly: false,
    hasAxes: true,
  },
  {
    id: "lineMarkers",
    label: "Line with markers",
    excelType: "LineMarkers",
    group: "Line & area",
    singleSeriesOnly: false,
    hasAxes: true,
  },
  {
    id: "area",
    label: "Area",
    excelType: "Area",
    group: "Line & area",
    singleSeriesOnly: false,
    hasAxes: true,
  },
  {
    id: "areaStacked",
    label: "Stacked area",
    excelType: "AreaStacked",
    group: "Line & area",
    singleSeriesOnly: false,
    hasAxes: true,
  },
  {
    id: "pie",
    label: "Pie",
    excelType: "Pie",
    group: "Pie",
    singleSeriesOnly: true,
    hasAxes: false,
  },
  {
    id: "doughnut",
    label: "Doughnut",
    excelType: "Doughnut",
    group: "Pie",
    singleSeriesOnly: true,
    hasAxes: false,
  },
  {
    id: "xyscatter",
    label: "Scatter",
    excelType: "XYScatter",
    group: "Scatter",
    singleSeriesOnly: false,
    hasAxes: true,
  },
  {
    id: "xyscatterLines",
    label: "Scatter with lines",
    excelType: "XYScatterLines",
    group: "Scatter",
    singleSeriesOnly: false,
    hasAxes: true,
  },
];

export function findChartType(id: string): ChartTypeMeta | undefined {
  return CHART_TYPES.find((type) => type.id === id);
}

export type LegendPosition = "top" | "bottom" | "left" | "right" | "none";

// ---------------------------------------------------------------------------
// Requests and validation
// ---------------------------------------------------------------------------

/** Where the rows come from. */
export interface DataSourceRef {
  kind: "selection" | "sheet" | "table";
  /** Sheet or table name. Ignored when kind is "selection". */
  name: string;
}

export interface SummarizeRequest {
  groupByColumn: number;
  valueColumns: number[];
  aggregation: Aggregation;
  sort: SortOrder;
  topN: number | null;
}

export interface ChartRequest {
  source: DataSourceRef;
  /** Null charts the source range as-is; otherwise the data is grouped first. */
  summarize: SummarizeRequest | null;
  chartTypeId: string;
  title: string;
  legendPosition: LegendPosition;
  showDataLabels: boolean;
  /** Put the chart (and any summary table) on a new sheet, or beside the data. */
  destination: "newSheet" | "sourceSheet";
  destinationSheetName: string;
}

/** Everything wrong with a request, before anything is written to the workbook. */
export function validateChartRequest(request: ChartRequest, headerCount: number): string[] {
  const problems: string[] = [];
  const chartType = findChartType(request.chartTypeId);

  if (!chartType) {
    problems.push(`Unknown chart type "${request.chartTypeId}".`);
  }

  if (request.summarize) {
    const { groupByColumn, valueColumns } = request.summarize;
    if (groupByColumn < 0 || groupByColumn >= headerCount) {
      problems.push("Choose a column to group by.");
    }
    if (valueColumns.length === 0) {
      problems.push("Choose at least one column to summarize.");
    }
    if (valueColumns.includes(groupByColumn)) {
      problems.push("The group-by column cannot also be a value column.");
    }
    if (chartType?.singleSeriesOnly && valueColumns.length > 1) {
      problems.push(
        `A ${chartType.label.toLowerCase()} chart shows one series - pick a single value column.`
      );
    }
    if (request.summarize.topN !== null && request.summarize.topN < 1) {
      problems.push("Top N must be at least 1.");
    }
  } else if (headerCount < 2) {
    problems.push("Charting a range directly needs at least two columns: labels and values.");
  }

  return problems;
}

// ---------------------------------------------------------------------------
// Reading the source
// ---------------------------------------------------------------------------

export interface SourceData {
  headers: string[];
  rows: Grid;
  /** Number format of the first data row, one per column, for the summary table. */
  numberFormats: string[];
  sheetName: string;
  address: string;
  rowIndex: number;
  columnIndex: number;
  rowCount: number;
  columnCount: number;
}

function headersFrom(grid: Grid): string[] {
  return (grid[0] ?? []).map((cell, index) => {
    const label = String(cell ?? "").trim();
    return label === "" ? `Column ${columnLetter(index)}` : label;
  });
}

async function readSource(
  context: Excel.RequestContext,
  ref: DataSourceRef
): Promise<SourceData | null> {
  let range: Excel.Range;

  if (ref.kind === "selection") {
    const selection = await loadSelectedRange(context);
    // A single cell means "the table I am standing in", as elsewhere in the add-in.
    range =
      selection.rowCount === 1 && selection.columnCount === 1
        ? context.workbook.worksheets.getActiveWorksheet().getUsedRangeOrNullObject(true)
        : selection.worksheet.getRangeByIndexes(
            selection.rowIndex,
            selection.columnIndex,
            selection.rowCount,
            selection.columnCount
          );
  } else if (ref.kind === "table") {
    range = context.workbook.tables.getItem(ref.name).getRange();
  } else {
    range = context.workbook.worksheets.getItem(ref.name).getUsedRangeOrNullObject(true);
  }

  range.load([
    "address",
    "values",
    "numberFormat",
    "rowIndex",
    "columnIndex",
    "rowCount",
    "columnCount",
    "worksheet/name",
  ]);
  await context.sync();

  if (range.isNullObject) {
    return null;
  }

  const grid = range.values as Grid;
  if (grid.length < 2) {
    return null;
  }

  const headers = headersFrom(grid);
  const formats = (range.numberFormat as string[][])[1] ?? [];

  return {
    headers,
    rows: grid.slice(1).map((row) => padRow(row, headers.length)),
    numberFormats: headers.map((_header, index) => formats[index] ?? "General"),
    sheetName: range.worksheet.name,
    address: range.address,
    rowIndex: range.rowIndex,
    columnIndex: range.columnIndex,
    rowCount: range.rowCount,
    columnCount: range.columnCount,
  };
}

/** Headers of a source, so the UI can offer group-by and value pickers. */
export async function readSourceHeaders(ref: DataSourceRef): Promise<string[]> {
  try {
    return await Excel.run(async (context) => {
      const source = await readSource(context, ref);
      return source?.headers ?? [];
    });
  } catch {
    return [];
  }
}

export interface SourceOption {
  kind: "sheet" | "table";
  name: string;
  label: string;
}

/** Sheets and tables the user can chart or summarize. */
export async function listSourceOptions(): Promise<SourceOption[]> {
  try {
    return await Excel.run(async (context) => {
      const sheetNames = await loadSheetNames(context);
      const tables = context.workbook.tables;
      tables.load("items/name");
      await context.sync();

      return [
        ...tables.items.map((table) => ({
          kind: "table" as const,
          name: table.name,
          label: `${table.name} (table)`,
        })),
        ...sheetNames.map((name) => ({ kind: "sheet" as const, name, label: `${name} (sheet)` })),
      ];
    });
  } catch {
    return [];
  }
}

// ---------------------------------------------------------------------------
// Writing summary tables
// ---------------------------------------------------------------------------

/** Number format to give an aggregated column, derived from its source column. */
function formatForAggregate(aggregation: Aggregation, sourceFormat: string): string {
  if (aggregation === "count" || aggregation === "countDistinct") {
    return "#,##0";
  }
  if (aggregation === "average") {
    // An integer format would hide the decimals an average almost always has.
    return sourceFormat === "General" || !/0/.test(sourceFormat) ? "#,##0.00" : sourceFormat;
  }
  return sourceFormat;
}

async function writeSummaryTable(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  rowIndex: number,
  columnIndex: number,
  summary: AggregatedTable,
  request: SummarizeRequest,
  sourceFormats: string[]
): Promise<{ table: Excel.Table; rowCount: number; columnCount: number }> {
  const grid: Grid = [summary.headers, ...summary.rows];

  const valueFormats = request.valueColumns.map((column) =>
    formatForAggregate(request.aggregation, sourceFormats[column] ?? "General")
  );
  const formats = grid.map((_row, index) =>
    index === 0
      ? new Array<string>(summary.headers.length).fill("General")
      : ["General", ...valueFormats]
  );

  await writeGridWithFormats(context, sheet, rowIndex, columnIndex, grid, formats);

  const table = sheet.tables.add(
    sheet.getRangeByIndexes(rowIndex, columnIndex, grid.length, summary.headers.length),
    true /* hasHeaders */
  );
  table.name = `MExSummary_${Date.now()}`;
  await context.sync();

  return { table, rowCount: grid.length, columnCount: summary.headers.length };
}

/** Build a grouped summary table on a new sheet, without charting it. */
export async function summarizeToNewSheet(
  ref: DataSourceRef,
  request: SummarizeRequest,
  destinationSheetName: string
): Promise<OperationResult> {
  return runExcel(async (context) => {
    const source = await readSource(context, ref);
    if (!source) {
      return fail("That source has no header row and at least one data row.");
    }

    const problems = validateChartRequest(
      {
        source: ref,
        summarize: request,
        chartTypeId: "columnClustered",
        title: "",
        legendPosition: "none",
        showDataLabels: false,
        destination: "newSheet",
        destinationSheetName,
      },
      source.headers.length
    );
    if (problems.length > 0) {
      return fail("Check the summary settings.", problems);
    }

    const summary = aggregateTable(source.headers, source.rows, request);
    if (summary.rows.length === 0) {
      return fail("Nothing to summarize - the source has no data rows.");
    }

    const sheet = await createSheetWithUniqueName(context, destinationSheetName);
    await writeSummaryTable(context, sheet, 0, 0, summary, request, source.numberFormats);

    sheet.getUsedRange().format.autofitColumns();
    sheet.freezePanes.freezeRows(1);
    sheet.activate();
    await context.sync();

    return ok(
      `Summarized ${plural(summary.sourceRows, "row")} into ${plural(summary.rows.length, "category", "categories")} on "${sheet.name}".`,
      summaryNotes(summary, request)
    );
  });
}

function summaryNotes(summary: AggregatedTable, request: SummarizeRequest): string[] {
  const notes: string[] = [];
  if (request.topN !== null && summary.categoryCount > request.topN) {
    notes.push(
      `${summary.categoryCount} categories found; the smallest ${summary.categoryCount - request.topN} were combined into "${OTHER_CATEGORY}".`
    );
  }
  if (summary.ignoredNonNumeric > 0) {
    notes.push(
      `${plural(summary.ignoredNonNumeric, "non-numeric cell")} ignored by the calculation.`
    );
  }
  return notes;
}

// ---------------------------------------------------------------------------
// Building the chart
// ---------------------------------------------------------------------------

function applyChartStyling(
  chart: Excel.Chart,
  request: ChartRequest,
  chartType: ChartTypeMeta,
  categoryHeader: string,
  valueHeader: string
): void {
  if (request.title.trim() !== "") {
    chart.title.text = request.title.trim();
    chart.title.visible = true;
  }

  if (request.legendPosition === "none") {
    chart.legend.visible = false;
  } else {
    chart.legend.visible = true;
    chart.legend.position = request.legendPosition as Excel.ChartLegendPosition;
  }

  if (request.showDataLabels) {
    chart.dataLabels.showValue = true;
  }

  if (chartType.hasAxes) {
    // Axis titles are unavailable on some chart types and hosts; they are a
    // nicety, so a failure here must not lose the chart itself.
    try {
      chart.axes.categoryAxis.title.text = categoryHeader;
      chart.axes.valueAxis.title.text = valueHeader;
    } catch {
      // Leave the axes untitled.
    }
  }
}

/**
 * Create a chart from a sheet, table or the current selection - summarizing the
 * data first when asked.
 */
export async function createChart(request: ChartRequest): Promise<OperationResult> {
  const chartType = findChartType(request.chartTypeId);
  if (!chartType) {
    return fail(`Unknown chart type "${request.chartTypeId}".`);
  }

  return runExcel(async (context) => {
    const source = await readSource(context, request.source);
    if (!source) {
      return fail("That source needs a header row and at least one data row.");
    }

    const problems = validateChartRequest(request, source.headers.length);
    if (problems.length > 0) {
      return fail("Check the chart settings.", problems);
    }

    const onNewSheet = request.destination === "newSheet";
    const sheet = onNewSheet
      ? await createSheetWithUniqueName(context, request.destinationSheetName)
      : context.workbook.worksheets.getItem(source.sheetName);

    let dataRange: Excel.Range;
    let categoryHeader = source.headers[0] ?? "";
    let valueHeader = source.headers[1] ?? "";
    const notes: string[] = [];
    // Where the chart itself is anchored, in sheet coordinates.
    let anchorRow = 0;
    let anchorColumn = 0;

    if (request.summarize) {
      const summary = aggregateTable(source.headers, source.rows, request.summarize);
      if (summary.rows.length === 0) {
        return fail("Nothing to chart - the source has no data rows.");
      }

      // The summary table has to live on a sheet for the chart to bind to it.
      const target = onNewSheet
        ? { row: 0, column: 0 }
        : { row: 0, column: source.columnIndex + source.columnCount + 1 };

      const written = await writeSummaryTable(
        context,
        sheet,
        target.row,
        target.column,
        summary,
        request.summarize,
        source.numberFormats
      );

      dataRange = sheet.getRangeByIndexes(
        target.row,
        target.column,
        written.rowCount,
        written.columnCount
      );
      categoryHeader = summary.headers[0];
      valueHeader = summary.headers[1] ?? "";
      anchorRow = target.row + written.rowCount + 1;
      anchorColumn = target.column;
      notes.push(
        `Summary table written to ${buildAddress(sheet.name, target.row, target.column, written.rowCount, written.columnCount)}.`
      );
      notes.push(...summaryNotes(summary, request.summarize));
    } else {
      dataRange = context.workbook.worksheets
        .getItem(source.sheetName)
        .getRangeByIndexes(
          source.rowIndex,
          source.columnIndex,
          source.rowCount,
          source.columnCount
        );
      anchorRow = onNewSheet ? 0 : source.rowIndex + source.rowCount + 1;
      anchorColumn = onNewSheet ? 0 : source.columnIndex;

      if (chartType.singleSeriesOnly && source.columnCount > 2) {
        notes.push(
          `A ${chartType.label.toLowerCase()} chart shows one series, so only "${source.headers[1]}" is plotted.`
        );
      }
    }

    const chart = sheet.charts.add(
      chartType.excelType as Excel.ChartType,
      dataRange,
      Excel.ChartSeriesBy.columns
    );
    applyChartStyling(chart, request, chartType, categoryHeader, valueHeader);

    // setPosition takes the top-left and bottom-right cells the chart covers.
    chart.setPosition(
      sheet.getRangeByIndexes(anchorRow, anchorColumn, 1, 1),
      sheet.getRangeByIndexes(anchorRow + 17, anchorColumn + 7, 1, 1)
    );
    chart.name = `MExChart_${Date.now()}`;
    await context.sync();

    if (onNewSheet) {
      sheet.getUsedRange().format.autofitColumns();
    }
    sheet.activate();
    await context.sync();

    return ok(
      `Added a ${chartType.label.toLowerCase()} chart to "${sheet.name}".`,
      notes.length > 0 ? notes : undefined
    );
  });
}
