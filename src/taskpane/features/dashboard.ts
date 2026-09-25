/**
 * Build a dashboard: up to eight charts and a row of headline numbers, laid
 * out as an organized grid on a new sheet, suggested automatically from what
 * the columns hold.
 *
 * The source goes through `tidyTable` first, so a messy export - titles,
 * subtotal rows, numbers typed as text - charts correctly without the user
 * having to fix it by hand. The source itself is never changed.
 *
 * Charts refer to columns by header name, not position, so a saved dashboard
 * still refreshes when next month's export has its columns in another order.
 *
 * Everything above the Office.js banner is pure and unit tested in
 * tests/dashboard.test.ts.
 */

import {
  createSheetWithUniqueName,
  fail,
  loadSelectedRange,
  ok,
  plural,
  runExcel,
  writeCell,
  writeGrid,
} from "../shared/excelHelpers";
import { CellValue, Grid, OperationResult } from "../shared/types";
import type { DataSourceRef } from "./charts";
import { excelSerialToDate, parseFlexibleDate, dateToExcelSerial } from "./dataCleaning";
import {
  cleanText,
  ColumnProfile,
  countProblems,
  numberFormatFor,
  parseLooseNumber,
  tidyTable,
} from "./tidy";

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

export type DashChartKind =
  "column" | "bar" | "line" | "pie" | "doughnut" | "stackedColumn" | "scatter";
export type TimeGrain = "day" | "month" | "quarter" | "year";
export type DashAggregation = "sum" | "count";

export interface DashChart {
  id: string;
  kind: DashChartKind;
  title: string;
  /** Why this chart was suggested, shown under its name. */
  reason: string;
  /** Column the chart groups by: a category or a date. */
  dimension: string | null;
  /** Second grouping, drawn as separate series (stacked bars, one line per group). */
  series: string | null;
  /** Column that is added up. Null counts rows instead. */
  measure: string | null;
  /** Second number column, for scatter charts. */
  measure2: string | null;
  aggregation: DashAggregation;
  timeGrain: TimeGrain | null;
  /** Show only the biggest N groups; the rest become "Other". */
  topN: number | null;
}

export interface Kpi {
  id: string;
  label: string;
  value: number;
  format: string;
  /** Text shown instead of the number, e.g. the name of the top region. */
  text: string | null;
  /** Small line under the value. */
  note: string | null;
}

export interface DashboardPlan {
  headers: string[];
  profiles: ColumnProfile[];
  charts: DashChart[];
  kpis: Kpi[];
  rowCount: number;
}

export const MAX_CHARTS = 8;
export const OTHER = "Other";
export const BLANK = "(blank)";

// ---------------------------------------------------------------------------
// Reading values
// ---------------------------------------------------------------------------

function readNumber(value: CellValue, profile: ColumnProfile): number | null {
  return parseLooseNumber(value, profile.numberStyle)?.value ?? null;
}

/** An Excel date serial for a cell, however the date was written. */
function readDate(value: CellValue, profile: ColumnProfile): number | null {
  if (typeof value === "number") {
    return Number.isFinite(value) ? Math.floor(value) : null;
  }
  if (typeof value !== "string") {
    return null;
  }
  const parsed = parseFlexibleDate(cleanText(value), { dayFirst: profile.dayFirst });
  return parsed ? dateToExcelSerial(parsed) : null;
}

function readLabel(value: CellValue): string {
  if (value === null || value === undefined) return BLANK;
  const text = typeof value === "string" ? cleanText(value) : String(value);
  return text === "" ? BLANK : text;
}

// ---------------------------------------------------------------------------
// Time buckets
// ---------------------------------------------------------------------------

/** Pick a time grain that gives a readable number of points. */
export function chooseGrain(firstSerial: number, lastSerial: number): TimeGrain {
  const span = lastSerial - firstSerial;
  if (span <= 45) return "day";
  if (span <= 1100) return "month";
  if (span <= 3000) return "quarter";
  return "year";
}

const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

/** A sortable key and a display label for the bucket a date falls in. */
export function bucketOf(serial: number, grain: TimeGrain): { key: number; label: string } {
  const date = excelSerialToDate(serial);
  const year = date.getUTCFullYear();
  const month = date.getUTCMonth();
  switch (grain) {
    case "day":
      return { key: serial, label: `${date.getUTCDate()} ${MONTHS[month]}` };
    case "month":
      return { key: year * 12 + month, label: `${MONTHS[month]} ${year}` };
    case "quarter": {
      const quarter = Math.floor(month / 3) + 1;
      return { key: year * 4 + quarter, label: `Q${quarter} ${year}` };
    }
    default:
      return { key: year, label: String(year) };
  }
}

const GRAIN_WORD: Record<TimeGrain, string> = {
  day: "Day",
  month: "Month",
  quarter: "Quarter",
  year: "Year",
};

// ---------------------------------------------------------------------------
// Choosing columns
// ---------------------------------------------------------------------------

const MEASURE_HINT =
  /revenue|sales|amount|total|value|cost|profit|spend|income|turnover|net|gross|fee|paid|balance/i;
/** Numbers that mean nothing added up: twenty unit prices summed is not a price. */
const NON_ADDITIVE =
  /price|rate|ratio|percent|%|margin|avg|average|mean|score|rating|\bage\b|temperature|index|rank|latitude|longitude|year/i;
const DIMENSION_HINT =
  /region|country|city|state|category|product|channel|segment|type|status|department|team|customer|client|brand|store|branch|group|source|year|quarter|month|rep|manager/i;

/** Number columns worth adding up, most important first. */
export function rankMeasures(profiles: readonly ColumnProfile[]): ColumnProfile[] {
  return profiles
    .filter((p) => (p.kind === "number" || p.kind === "currency") && !NON_ADDITIVE.test(p.header))
    .map((p) => ({
      p,
      score: (p.kind === "currency" ? 1 : 0) + (MEASURE_HINT.test(p.header) ? 2 : 0),
    }))
    .sort((a, b) => b.score - a.score || a.p.index - b.p.index)
    .map(({ p }) => p);
}

/** Columns worth grouping by, most chartable first: a handful of distinct values beats hundreds. */
export function rankDimensions(profiles: readonly ColumnProfile[]): ColumnProfile[] {
  return profiles
    .filter((p) => p.kind === "category" && p.distinct >= 2)
    .map((p) => {
      const sizeScore = p.distinct >= 3 && p.distinct <= 12 ? 2 : p.distinct <= 20 ? 1 : 0;
      return { p, score: sizeScore + (DIMENSION_HINT.test(p.header) ? 2 : 0) };
    })
    .sort((a, b) => b.score - a.score || a.p.distinct - b.p.distinct || a.p.index - b.p.index)
    .map(({ p }) => p);
}

/** Columns with too many values for a pie, but right for a ranked list (products, customers). */
function rankWideDimensions(
  profiles: readonly ColumnProfile[],
  used: ReadonlySet<string>
): ColumnProfile[] {
  return profiles
    .filter((p) => (p.kind === "text" || p.kind === "category") && !used.has(p.header))
    .filter((p) => p.distinct >= 8 && p.distinct < p.nonBlank)
    .sort(
      (a, b) =>
        Number(DIMENSION_HINT.test(b.header)) - Number(DIMENSION_HINT.test(a.header)) ||
        a.index - b.index
    );
}

// ---------------------------------------------------------------------------
// Suggestions
// ---------------------------------------------------------------------------

function measureWord(measure: ColumnProfile | null): string {
  return measure ? measure.header : "Count";
}

/** "Revenue by Region", or "Top 10 Product by Revenue" when only the top ten fit. */
function byTitle(what: string, dimension: ColumnProfile): string {
  return dimension.distinct > 10
    ? `Top 10 ${dimension.header} by ${what}`
    : `${what} by ${dimension.header}`;
}

/**
 * Suggest up to `max` charts that together give a rounded picture: a trend, a
 * comparison, a share, a breakdown, a top-10, a relationship.
 */
export function suggestCharts(
  profiles: readonly ColumnProfile[],
  rows: Grid,
  max = MAX_CHARTS
): DashChart[] {
  const measures = rankMeasures(profiles);
  const dims = rankDimensions(profiles);
  const dates = profiles.filter((p) => p.kind === "date");
  const m1 = measures[0] ?? null;
  const m2 = measures[1] ?? null;
  const date = dates[0] ?? null;
  const [c0, c1, c2] = dims;
  const what = measureWord(m1);
  const aggregation: DashAggregation = m1 ? "sum" : "count";

  const charts: DashChart[] = [];
  const add = (
    chart: Omit<
      DashChart,
      "id" | "aggregation" | "measure" | "measure2" | "series" | "timeGrain" | "topN"
    > &
      Partial<DashChart>
  ) => {
    const full: DashChart = {
      aggregation,
      measure: m1?.header ?? null,
      measure2: null,
      series: null,
      timeGrain: null,
      topN: null,
      ...chart,
      id: "",
    };
    full.id = [full.kind, full.dimension, full.series, full.measure, full.measure2].join("|");
    if (!charts.some((existing) => existing.id === full.id || existing.title === full.title))
      charts.push(full);
  };

  let grain: TimeGrain | null = null;
  if (date) {
    const serials = rows
      .map((row) => readDate(row[date.index], date))
      .filter((s): s is number => s !== null);
    if (serials.length > 0) {
      grain = chooseGrain(Math.min(...serials), Math.max(...serials));
    }
  }

  if (date && grain) {
    add({
      kind: "line",
      title: `${what} over time`,
      reason: `Trend by ${GRAIN_WORD[grain].toLowerCase()}, to spot growth and seasonality.`,
      dimension: date.header,
      timeGrain: grain,
    });
  }
  if (c0) {
    add({
      kind: c0.distinct > 8 ? "bar" : "column",
      title: byTitle(what, c0),
      reason: `Compare each ${c0.header.toLowerCase()} side by side.`,
      dimension: c0.header,
      topN: 10,
    });
  }
  if (c1) {
    if (c1.distinct <= 6) {
      add({
        kind: "doughnut",
        title: `Share of ${what.toLowerCase() === "count" ? "rows" : what} by ${c1.header}`,
        reason: `How the total splits across ${c1.header.toLowerCase()}.`,
        dimension: c1.header,
        topN: 6,
      });
    } else {
      add({
        kind: "bar",
        title: byTitle(what, c1),
        reason: `Compare each ${c1.header.toLowerCase()}, biggest first.`,
        dimension: c1.header,
        topN: 10,
      });
    }
  }
  if (c0 && c1) {
    add({
      kind: "stackedColumn",
      title: `${what} by ${c0.header} and ${c1.header}`,
      reason: `Each ${c0.header.toLowerCase()} broken down by ${c1.header.toLowerCase()}.`,
      dimension: c0.header,
      series: c1.header,
      topN: 8,
    });
  }
  if (date && grain && (c1 || c0)) {
    // One line per group reads best with few groups.
    const split = [c1, c0].filter(Boolean).sort((a, b) => a!.distinct - b!.distinct)[0]!;
    add({
      kind: "line",
      title: `${what} over time by ${split.header}`,
      reason: `Which ${split.header.toLowerCase()} is driving the trend.`,
      dimension: date.header,
      series: split.header,
      timeGrain: grain,
    });
  }
  const used = new Set([c0, c1].filter(Boolean).map((p) => p!.header));
  const wide = rankWideDimensions(profiles, used)[0] ?? null;
  if (wide) {
    const ranked = wide.distinct > 10;
    add({
      kind: "bar",
      title: byTitle(what, wide),
      reason: ranked
        ? `The biggest ${wide.header.toLowerCase()} values, with the rest grouped as Other.`
        : `Every ${wide.header.toLowerCase()}, biggest first.`,
      dimension: wide.header,
      topN: 10,
    });
  }
  if (m2 && c0) {
    add({
      kind: "column",
      title: `${m2.header} by ${c0.header}`,
      reason: `A second measure across each ${c0.header.toLowerCase()}.`,
      dimension: c0.header,
      measure: m2.header,
    });
  }
  if (m1 && m2 && rows.length >= 8) {
    add({
      kind: "scatter",
      title: `${m1.header} vs ${m2.header}`,
      reason: "See whether the two move together.",
      dimension: null,
      measure: m1.header,
      measure2: m2.header,
    });
  }
  if (c2 && c2 !== wide) {
    add({
      kind: c2.distinct <= 6 ? "pie" : "column",
      title: `${what} by ${c2.header}`,
      reason: `Compare each ${c2.header.toLowerCase()}.`,
      dimension: c2.header,
      topN: 8,
    });
  }
  if (date && grain && m1) {
    add({
      kind: "column",
      title: `Rows per ${GRAIN_WORD[grain].toLowerCase()}`,
      reason: "How much activity each period had.",
      dimension: date.header,
      measure: null,
      aggregation: "count",
      timeGrain: grain,
    });
  }
  if (c0 && m1) {
    add({
      kind: "pie",
      title: `Rows by ${c0.header}`,
      reason: `How many records each ${c0.header.toLowerCase()} has.`,
      dimension: c0.header,
      measure: null,
      aggregation: "count",
      topN: 6,
    });
  }

  return charts.slice(0, max);
}

export interface CustomChartRequest {
  kind: DashChartKind;
  /** Column to group by: a category or a date. */
  dimension: string | null;
  /** Column to add up. Null counts rows instead. */
  measure: string | null;
  /** Second grouping for stacked bars or one line per group. */
  series?: string | null;
  /** Second number column, for a scatter. */
  measure2?: string | null;
  aggregation?: DashAggregation;
  topN?: number | null;
}

/**
 * Build one chart to the user's own specification, rather than from the
 * suggestions. Returns null when the columns don't make a chart - a scatter
 * needs two numbers, everything else needs something to group by.
 */
export function customChart(
  request: CustomChartRequest,
  profiles: readonly ColumnProfile[]
): DashChart | null {
  const find = (name: string | null | undefined) =>
    name ? (profiles.find((p) => p.header === name) ?? null) : null;
  const dimension = find(request.dimension);
  const measure = find(request.measure);
  const measure2 = find(request.measure2);
  const series = find(request.series);
  const aggregation: DashAggregation = request.aggregation ?? (measure ? "sum" : "count");

  if (request.kind === "scatter") {
    if (!measure || !measure2) return null;
    return {
      id: ["scatter", null, null, measure.header, measure2.header].join("|"),
      kind: "scatter",
      title: `${measure.header} vs ${measure2.header}`,
      reason: "Your choice.",
      dimension: null,
      series: null,
      measure: measure.header,
      measure2: measure2.header,
      aggregation: "sum",
      timeGrain: null,
      topN: null,
    };
  }
  if (!dimension) return null;

  const what = measure ? measure.header : "Count";
  const grain = dimension.kind === "date" ? "month" : null;
  const title = grain
    ? `${what} over time${series ? ` by ${series.header}` : ""}`
    : series
      ? `${what} by ${dimension.header} and ${series.header}`
      : `${what} by ${dimension.header}`;

  return {
    id: [
      request.kind,
      dimension.header,
      series?.header ?? null,
      measure?.header ?? null,
      null,
    ].join("|"),
    kind: request.kind,
    title,
    reason: "Your choice.",
    dimension: dimension.header,
    series: series?.header ?? null,
    measure: measure?.header ?? null,
    measure2: null,
    aggregation,
    timeGrain: grain,
    topN: request.topN ?? (request.kind === "pie" || request.kind === "doughnut" ? 6 : null),
  };
}

/** Up to four headline numbers for the top of the dashboard. */
export function suggestKpis(profiles: readonly ColumnProfile[], rows: Grid): Kpi[] {
  const measures = rankMeasures(profiles);
  const dims = rankDimensions(profiles);
  const m1 = measures[0] ?? null;
  const m2 = measures[1] ?? null;
  const kpis: Kpi[] = [];

  const sum = (profile: ColumnProfile) =>
    rows.reduce((total, row) => total + (readNumber(row[profile.index], profile) ?? 0), 0);
  const formatOf = (profile: ColumnProfile) =>
    numberFormatFor(
      profile,
      rows.map((row) => readNumber(row[profile.index], profile))
    );

  if (m1) {
    const total = sum(m1);
    kpis.push({
      id: "total",
      label: `Total ${m1.header}`,
      value: total,
      format: formatOf(m1),
      text: null,
      note: null,
    });
  }
  kpis.push({
    id: "rows",
    label: "Records",
    value: rows.length,
    format: "#,##0",
    text: null,
    note: null,
  });
  if (m1 && rows.length > 0) {
    kpis.push({
      id: "average",
      label: `Average ${m1.header}`,
      value: sum(m1) / rows.length,
      // An average almost always has decimals; currency formats already show them.
      format: m1.kind === "currency" ? formatOf(m1) : "#,##0.00",
      text: null,
      note: "per record",
    });
  }
  if (dims[0]) {
    const totals = new Map<string, number>();
    for (const row of rows) {
      const key = readLabel(row[dims[0].index]);
      const amount = m1 ? (readNumber(row[m1.index], m1) ?? 0) : 1;
      totals.set(key, (totals.get(key) ?? 0) + amount);
    }
    const [name, value] = Array.from(totals.entries()).sort((a, b) => b[1] - a[1])[0] ?? [null, 0];
    if (name !== null) {
      kpis.push({
        id: "top",
        label: `Top ${dims[0].header}`,
        value,
        format: m1 ? formatOf(m1) : "#,##0",
        text: name,
        note: m1 ? `${m1.header}` : "records",
      });
    }
  } else if (m2) {
    kpis.push({
      id: "total2",
      label: `Total ${m2.header}`,
      value: sum(m2),
      format: formatOf(m2),
      text: null,
      note: null,
    });
  }
  return kpis.slice(0, 4);
}

// ---------------------------------------------------------------------------
// Chart data
// ---------------------------------------------------------------------------

function profileByName(
  profiles: readonly ColumnProfile[],
  name: string | null
): ColumnProfile | null {
  if (name === null) return null;
  return profiles.find((p) => p.header === name) ?? null;
}

/** Keep the biggest `topN` entries and fold the rest into "Other". */
export function topWithOther(
  entries: Array<[string, number]>,
  topN: number | null
): Array<[string, number]> {
  const sorted = entries.slice().sort((a, b) => b[1] - a[1]);
  if (topN === null || sorted.length <= topN) return sorted;
  const kept = sorted.slice(0, topN - 1);
  const rest = sorted.slice(topN - 1).reduce((total, [, value]) => total + value, 0);
  return [...kept, [OTHER, rest]];
}

/**
 * The small table a chart is drawn from: a header row, then one row per
 * category or time bucket. Returns null when the chart's columns are missing.
 */
export function buildChartData(
  chart: DashChart,
  profiles: readonly ColumnProfile[],
  rows: Grid
): Grid | null {
  const dimension = profileByName(profiles, chart.dimension);
  const series = profileByName(profiles, chart.series);
  const measure = profileByName(profiles, chart.measure);
  const measure2 = profileByName(profiles, chart.measure2);
  if ((chart.dimension && !dimension) || (chart.series && !series) || (chart.measure && !measure)) {
    return null;
  }
  const amountOf = (row: CellValue[]): number | null =>
    chart.aggregation === "count" || !measure ? 1 : readNumber(row[measure.index], measure);
  const valueName = chart.aggregation === "count" || !measure ? "Count" : measure.header;

  if (chart.kind === "scatter") {
    if (!measure || !measure2) return null;
    const points = rows
      .map((row) => [
        readNumber(row[measure.index], measure),
        readNumber(row[measure2.index], measure2),
      ])
      .filter((pair): pair is [number, number] => pair[0] !== null && pair[1] !== null);
    // Excel draws a scatter's first column along the bottom, so put measure2 there.
    const step = Math.max(1, Math.ceil(points.length / 400));
    const sampled = points.filter((_p, i) => i % step === 0);
    return [[measure2.header, measure.header], ...sampled.map(([y, x]) => [x, y])];
  }
  if (!dimension) return null;

  const keyOf = (row: CellValue[]): { key: string | number; label: string } | null => {
    if (chart.timeGrain) {
      const serial = readDate(row[dimension.index], dimension);
      if (serial === null) return null;
      const bucket = bucketOf(serial, chart.timeGrain);
      return { key: bucket.key, label: bucket.label };
    }
    const label = readLabel(row[dimension.index]);
    return { key: label, label };
  };

  if (!series) {
    const totals = new Map<string | number, { label: string; value: number }>();
    for (const row of rows) {
      const bucket = keyOf(row);
      const amount = amountOf(row);
      if (!bucket || amount === null) continue;
      const entry = totals.get(bucket.key) ?? { label: bucket.label, value: 0 };
      entry.value += amount;
      totals.set(bucket.key, entry);
    }
    const header: CellValue[] = [
      chart.timeGrain ? GRAIN_WORD[chart.timeGrain] : dimension.header,
      valueName,
    ];
    if (chart.timeGrain) {
      const ordered = Array.from(totals.entries()).sort((a, b) => Number(a[0]) - Number(b[0]));
      return [header, ...ordered.map(([, entry]) => [entry.label, entry.value])];
    }
    const entries = topWithOther(
      Array.from(totals.values()).map((entry) => [entry.label, entry.value] as [string, number]),
      chart.topN
    );
    return [header, ...entries];
  }

  // Two groupings: rows are the dimension, columns the top series groups.
  const seriesTotals = new Map<string, number>();
  for (const row of rows) {
    const amount = amountOf(row);
    if (amount === null) continue;
    const label = readLabel(row[series.index]);
    seriesTotals.set(label, (seriesTotals.get(label) ?? 0) + amount);
  }
  const seriesNames = topWithOther(Array.from(seriesTotals.entries()), 5).map(([name]) => name);
  const seriesIndex = (label: string) => {
    const index = seriesNames.indexOf(label);
    return index >= 0 ? index : seriesNames.indexOf(OTHER);
  };

  const cells = new Map<string | number, { label: string; values: number[] }>();
  for (const row of rows) {
    const bucket = keyOf(row);
    const amount = amountOf(row);
    if (!bucket || amount === null) continue;
    const entry = cells.get(bucket.key) ?? {
      label: bucket.label,
      values: seriesNames.map(() => 0),
    };
    const index = seriesIndex(readLabel(row[series.index]));
    if (index >= 0) entry.values[index] += amount;
    cells.set(bucket.key, entry);
  }

  let ordered = Array.from(cells.entries());
  if (chart.timeGrain) {
    ordered.sort((a, b) => Number(a[0]) - Number(b[0]));
  } else {
    ordered.sort(
      (a, b) => b[1].values.reduce((s, v) => s + v, 0) - a[1].values.reduce((s, v) => s + v, 0)
    );
    if (chart.topN !== null && ordered.length > chart.topN) ordered = ordered.slice(0, chart.topN);
  }

  const first = chart.timeGrain ? GRAIN_WORD[chart.timeGrain] : dimension.header;
  return [[first, ...seriesNames], ...ordered.map(([, entry]) => [entry.label, ...entry.values])];
}

/** Everything needed to show suggestions for a table. */
export function planDashboard(
  headers: string[],
  profiles: ColumnProfile[],
  rows: Grid,
  options: PlanOptions = {}
): DashboardPlan {
  // Suggestions are limited to the chosen columns, but the full profile list is
  // kept, so any column is still available to build a chart from by hand.
  const usable =
    options.useColumns && options.useColumns.length > 0
      ? profiles.filter((profile) => options.useColumns!.indexOf(profile.header) >= 0)
      : profiles;
  return {
    headers,
    profiles,
    charts: suggestCharts(usable, rows, options.max ?? MAX_CHARTS),
    kpis: suggestKpis(usable, rows),
    rowCount: rows.length,
  };
}

export interface PlanOptions {
  /** How many charts to suggest. */
  max?: number;
  /** Only suggest charts built from these columns. Empty means all of them. */
  useColumns?: readonly string[];
}

// ---------------------------------------------------------------------------
// Layout
// ---------------------------------------------------------------------------

export interface Placement {
  row: number;
  column: number;
  rows: number;
  columns: number;
}

export interface DashboardLayout {
  kpis: Placement[];
  charts: Placement[];
  /** First row after the dashboard. */
  bottom: number;
}

/** The dashboard grid: 12 columns, a title band, a KPI band, then charts two to a row. */
export const LAYOUT = {
  columns: 12,
  titleRows: 3,
  kpiRows: 4,
  chartRows: 16,
  gap: 1,
} as const;

/**
 * Place KPI tiles and charts on a 12-column grid. The first chart - the trend,
 * when there is one - runs full width as the hero; the rest sit two per row,
 * and an odd one out at the end also runs full width so nothing sits alone
 * beside a gap.
 */
export function layoutDashboard(
  chartCount: number,
  kpiCount: number,
  heroFirst: boolean
): DashboardLayout {
  const kpis: Placement[] = [];
  let row = LAYOUT.titleRows;
  if (kpiCount > 0) {
    const width = Math.floor(LAYOUT.columns / kpiCount);
    for (let i = 0; i < kpiCount; i += 1) {
      kpis.push({ row, column: i * width, rows: LAYOUT.kpiRows, columns: width });
    }
    row += LAYOUT.kpiRows + LAYOUT.gap;
  }

  const charts: Placement[] = [];
  let index = 0;
  if (heroFirst && chartCount > 0) {
    charts.push({ row, column: 0, rows: LAYOUT.chartRows, columns: LAYOUT.columns });
    row += LAYOUT.chartRows + LAYOUT.gap;
    index = 1;
  }
  while (index < chartCount) {
    const remaining = chartCount - index;
    if (remaining === 1) {
      charts.push({ row, column: 0, rows: LAYOUT.chartRows, columns: LAYOUT.columns });
      index += 1;
    } else {
      const half = LAYOUT.columns / 2;
      charts.push({ row, column: 0, rows: LAYOUT.chartRows, columns: half });
      charts.push({ row, column: half, rows: LAYOUT.chartRows, columns: half });
      index += 2;
    }
    row += LAYOUT.chartRows + LAYOUT.gap;
  }
  return { kpis, charts, bottom: row };
}

// ---------------------------------------------------------------------------
// Colours
// ---------------------------------------------------------------------------

/**
 * Categorical palette, in a fixed order validated for colour-blind separation
 * between neighbours. "Other" is always neutral grey.
 */
export const PALETTE = [
  "#2a78d6",
  "#eb6834",
  "#1baf7a",
  "#eda100",
  "#e87ba4",
  "#008300",
  "#4a3aa7",
  "#e34948",
];
export const OTHER_COLOR = "#9aa3a0";

/**
 * One colour per value of a column, ranked by size across the whole table, so
 * "Europe" is the same colour in every chart it appears in.
 */
export function colorsFor(labels: readonly string[]): Map<string, string> {
  const map = new Map<string, string>();
  let slot = 0;
  for (const label of labels) {
    if (label === OTHER) {
      map.set(label, OTHER_COLOR);
    } else if (!map.has(label)) {
      map.set(label, PALETTE[slot % PALETTE.length]);
      slot += 1;
    }
  }
  return map;
}

// ---------------------------------------------------------------------------
// Office.js drivers
// ---------------------------------------------------------------------------

export interface DashboardConfig {
  name: string;
  title: string;
  source: DataSourceRef;
  charts: DashChart[];
  includeKpis: boolean;
  sheetName: string;
  dataSheetName: string;
  updatedAt: string;
}

const KEY_PREFIX = "MExAutomate.dashboard.";

interface SourceTable {
  headers: string[];
  rows: Grid;
  profiles: ColumnProfile[];
  name: string;
  problemsFixed: number;
}

async function readSource(
  context: Excel.RequestContext,
  ref: DataSourceRef
): Promise<SourceTable | null> {
  let range: Excel.Range;
  if (ref.kind === "selection") {
    const selection = await loadSelectedRange(context);
    range =
      selection.rowCount === 1 && selection.columnCount === 1
        ? context.workbook.worksheets.getActiveWorksheet().getUsedRangeOrNullObject(true)
        : selection;
  } else if (ref.kind === "table") {
    range = context.workbook.tables.getItem(ref.name).getRange();
  } else {
    range = context.workbook.worksheets.getItem(ref.name).getUsedRangeOrNullObject(true);
  }
  range.load(["values", "numberFormat", "worksheet/name"]);
  await context.sync();
  if (range.isNullObject) return null;

  // Chart from a repaired copy of the data, so titles, totals and text numbers
  // don't distort the charts. The sheet itself is left exactly as it was.
  const tidy = tidyTable(range.values as Grid, { formats: range.numberFormat as string[][] });
  if (!tidy || tidy.rows.length === 0) return null;
  return {
    headers: tidy.headers,
    rows: tidy.rows,
    profiles: tidy.columns,
    name: ref.kind === "table" ? ref.name : range.worksheet.name,
    problemsFixed: countProblems(tidy.findings),
  };
}

export interface DashboardAnalysis {
  plan: DashboardPlan | null;
  sourceName: string;
  /** Problems smoothed over on the way (the sheet itself is unchanged). */
  problemsFixed: number;
}

/** Look at a table and suggest charts and headline numbers for it. */
export async function analyzeDashboard(
  ref: DataSourceRef,
  options: PlanOptions = {}
): Promise<DashboardAnalysis> {
  try {
    return await Excel.run(async (context) => {
      const source = await readSource(context, ref);
      if (!source) return { plan: null, sourceName: ref.name, problemsFixed: 0 };
      return {
        plan: planDashboard(source.headers, source.profiles, source.rows, options),
        sourceName: source.name,
        problemsFixed: source.problemsFixed,
      };
    });
  } catch {
    return { plan: null, sourceName: ref.name, problemsFixed: 0 };
  }
}

const EXCEL_TYPE: Record<DashChartKind, string> = {
  column: "ColumnClustered",
  bar: "BarClustered",
  line: "LineMarkers",
  pie: "Pie",
  doughnut: "Doughnut",
  stackedColumn: "ColumnStacked",
  scatter: "XYScatter",
};

/** Points, matching the column width and row height the builder sets. */
const CELL_WIDTH = 64;
const CELL_HEIGHT = 16;
const INSET = 5;

async function prepareSheet(
  context: Excel.RequestContext,
  existingName: string | null,
  desiredName: string
): Promise<Excel.Worksheet> {
  if (existingName) {
    const existing = context.workbook.worksheets.getItemOrNullObject(existingName);
    await context.sync();
    if (!existing.isNullObject) {
      existing.charts.load("items/name");
      await context.sync();
      existing.charts.items.forEach((chart) => chart.delete());
      const everything = existing.getRange();
      everything.unmerge();
      everything.clear(Excel.ClearApplyTo.all);
      await context.sync();
      return existing;
    }
  }
  return createSheetWithUniqueName(context, desiredName);
}

function styleChart(
  chart: Excel.Chart,
  spec: DashChart,
  data: Grid,
  colorMaps: Map<string, Map<string, string>>
): void {
  chart.title.text = spec.title;
  chart.title.format.font.size = 12;
  chart.title.format.font.bold = true;

  const seriesCount = data[0].length - 1;
  const isRound = spec.kind === "pie" || spec.kind === "doughnut";
  if (isRound) {
    chart.legend.visible = true;
    chart.legend.position = Excel.ChartLegendPosition.right;
    chart.dataLabels.showPercentage = true;
    chart.dataLabels.showValue = false;
  } else if (seriesCount > 1) {
    chart.legend.visible = true;
    chart.legend.position = Excel.ChartLegendPosition.bottom;
  } else {
    chart.legend.visible = false;
  }

  if (spec.kind === "bar") {
    // Excel draws bar categories bottom-up; flip so the biggest is on top.
    try {
      chart.axes.categoryAxis.reversePlotOrder = true;
    } catch {
      // Older hosts: leave the default order.
    }
  }

  const categoryColors = spec.dimension ? colorMaps.get(spec.dimension) : undefined;
  const seriesColors = spec.series ? colorMaps.get(spec.series) : undefined;
  const header = data[0].slice(1).map(String);

  for (let s = 0; s < seriesCount; s += 1) {
    const series = chart.series.getItemAt(s);
    const color = seriesColors?.get(header[s]) ?? PALETTE[s % PALETTE.length];
    if (isRound) {
      data.slice(1).forEach((row, p) => {
        const pointColor = categoryColors?.get(String(row[0])) ?? PALETTE[p % PALETTE.length];
        series.points.getItemAt(p).format.fill.setSolidColor(pointColor);
      });
    } else if (spec.kind === "line") {
      series.format.line.color = color;
    } else if (spec.kind === "scatter") {
      series.markerBackgroundColor = color;
      series.markerForegroundColor = color;
    } else {
      series.format.fill.setSolidColor(color);
    }
  }
}

/** Rank a column's values by total, for consistent colours across charts. */
function rankedLabels(profile: ColumnProfile, measure: ColumnProfile | null, rows: Grid): string[] {
  const totals = new Map<string, number>();
  for (const row of rows) {
    const label = readLabel(row[profile.index]);
    const amount = measure ? (readNumber(row[measure.index], measure) ?? 0) : 1;
    totals.set(label, (totals.get(label) ?? 0) + amount);
  }
  return Array.from(totals.entries())
    .sort((a, b) => b[1] - a[1])
    .map(([label]) => label);
}

async function drawDashboard(
  context: Excel.RequestContext,
  config: Omit<DashboardConfig, "sheetName" | "dataSheetName" | "updatedAt">,
  existing: { sheetName: string; dataSheetName: string } | null
): Promise<{ result: OperationResult; sheetName: string; dataSheetName: string }> {
  const source = await readSource(context, config.source);
  if (!source) {
    return {
      result: fail("Couldn't find a table with headings and data rows there."),
      sheetName: existing?.sheetName ?? "",
      dataSheetName: existing?.dataSheetName ?? "",
    };
  }

  const drawable = config.charts
    .map((chart) => ({ chart, data: buildChartData(chart, source.profiles, source.rows) }))
    .filter(
      (item): item is { chart: DashChart; data: Grid } => item.data !== null && item.data.length > 1
    )
    .slice(0, MAX_CHARTS);
  const skipped = config.charts.length - drawable.length;
  const kpis = config.includeKpis ? suggestKpis(source.profiles, source.rows) : [];

  const sheet = await prepareSheet(context, existing?.sheetName ?? null, config.title);
  const dataSheet = await prepareSheet(
    context,
    existing?.dataSheetName ?? null,
    `${config.title} data`
  );
  sheet.load("name");
  dataSheet.load("name");
  await context.sync();

  // --- The numbers behind every chart, on their own sheet ---------------------
  const ranges: Excel.Range[] = [];
  let dataRow = 0;
  for (const { chart, data } of drawable) {
    writeCell(dataSheet, dataRow, 0, chart.title);
    dataSheet.getRangeByIndexes(dataRow, 0, 1, 1).format.font.bold = true;
    await writeGrid(context, dataSheet, dataRow + 1, 0, data);
    dataSheet.getRangeByIndexes(dataRow + 1, 0, 1, data[0].length).format.font.bold = true;
    ranges.push(dataSheet.getRangeByIndexes(dataRow + 1, 0, data.length, data[0].length));
    dataRow += data.length + 3;
  }
  dataSheet.getUsedRange().format.autofitColumns();

  // --- The dashboard sheet ------------------------------------------------------
  const heroFirst = drawable[0]?.chart.kind === "line" && drawable[0].chart.series === null;
  const layout = layoutDashboard(drawable.length, kpis.length, heroFirst);
  sheet.getRangeByIndexes(0, 0, layout.bottom + 2, LAYOUT.columns).format.rowHeight = CELL_HEIGHT;
  sheet.getRangeByIndexes(0, 0, 1, LAYOUT.columns).format.columnWidth = CELL_WIDTH;
  try {
    sheet.showGridlines = false;
  } catch {
    // Gridlines stay on for hosts without ExcelApi 1.8.
  }

  const title = sheet.getRangeByIndexes(0, 0, 1, 1);
  writeCell(sheet, 0, 0, config.title);
  title.format.font.size = 20;
  title.format.font.bold = true;
  title.format.rowHeight = 30;
  const subtitle = sheet.getRangeByIndexes(1, 0, 1, 1);
  subtitle.values = [
    [
      `From "${source.name}" · ${plural(source.rows.length, "record")} · updated ${new Date().toLocaleDateString(undefined, { day: "numeric", month: "short", year: "numeric" })}`,
    ],
  ];
  subtitle.format.font.color = "#5b6b66";

  kpis.forEach((kpi, i) => {
    const place = layout.kpis[i];
    const tile = sheet.getRangeByIndexes(place.row, place.column, place.rows, place.columns);
    tile.format.fill.color = "#eef3f1";
    // A thick border in the sheet's own white makes a gap between tiles.
    for (const edge of [
      Excel.BorderIndex.edgeLeft,
      Excel.BorderIndex.edgeRight,
      Excel.BorderIndex.edgeTop,
      Excel.BorderIndex.edgeBottom,
    ]) {
      const border = tile.format.borders.getItem(edge);
      border.style = Excel.BorderLineStyle.continuous;
      border.weight = Excel.BorderWeight.thick;
      border.color = "#ffffff";
    }

    // Each band is merged for looks, but written through its top-left cell:
    // a merged range still reports its full width to Office.js.
    const label = sheet.getRangeByIndexes(place.row, place.column, 1, place.columns);
    label.merge();
    writeCell(sheet, place.row, place.column, kpi.label);
    label.format.font.color = "#4b5c57";
    label.format.indentLevel = 1;

    const value = sheet.getRangeByIndexes(place.row + 1, place.column, 2, place.columns);
    value.merge();
    writeCell(
      sheet,
      place.row + 1,
      place.column,
      kpi.text ?? kpi.value,
      kpi.text ? undefined : kpi.format
    );
    value.format.font.size = 20;
    value.format.font.bold = true;
    value.format.verticalAlignment = Excel.VerticalAlignment.center;
    value.format.indentLevel = 1;

    const note = sheet.getRangeByIndexes(place.row + 3, place.column, 1, place.columns);
    note.merge();
    if (kpi.text) {
      writeCell(sheet, place.row + 3, place.column, kpi.value, kpi.format);
    } else if (kpi.note) {
      writeCell(sheet, place.row + 3, place.column, kpi.note);
    }
    note.format.font.color = "#4b5c57";
    note.format.horizontalAlignment = Excel.HorizontalAlignment.left;
    note.format.indentLevel = 1;
  });

  // Colours follow the value: "Europe" is the same colour in every chart.
  const measure =
    source.profiles.find(
      (p) => p.header === drawable.find((d) => d.chart.measure)?.chart.measure
    ) ?? null;
  const colorMaps = new Map<string, Map<string, string>>();
  for (const profile of source.profiles) {
    if (profile.kind === "category" || profile.kind === "text") {
      colorMaps.set(
        profile.header,
        colorsFor([...rankedLabels(profile, measure, source.rows), OTHER])
      );
    }
  }

  drawable.forEach(({ chart: spec, data }, i) => {
    const place = layout.charts[i];
    const chart = sheet.charts.add(
      EXCEL_TYPE[spec.kind] as Excel.ChartType,
      ranges[i],
      Excel.ChartSeriesBy.columns
    );
    chart.left = place.column * CELL_WIDTH + INSET;
    chart.top = place.row * CELL_HEIGHT + 14 + INSET;
    chart.width = place.columns * CELL_WIDTH - INSET * 2;
    chart.height = place.rows * CELL_HEIGHT - INSET * 2;
    chart.name = `MExDash_${i + 1}`;
    styleChart(chart, spec, data, colorMaps);
  });

  sheet.activate();
  sheet.getRangeByIndexes(0, 0, 1, 1).select();
  await context.sync();

  const notes: string[] = [];
  if (source.problemsFixed > 0) {
    notes.push(
      `${plural(source.problemsFixed, "data problem")} smoothed over for the charts. Your sheet is unchanged.`
    );
  }
  if (skipped > 0) {
    notes.push(`${plural(skipped, "chart")} skipped: their columns aren't in the data any more.`);
  }
  notes.push(`The numbers behind each chart are on "${dataSheet.name}".`);

  return {
    result: ok(
      `Built "${sheet.name}" with ${plural(drawable.length, "chart")}${kpis.length ? ` and ${kpis.length} headline numbers` : ""}.`,
      notes
    ),
    sheetName: sheet.name,
    dataSheetName: dataSheet.name,
  };
}

async function saveConfig(context: Excel.RequestContext, config: DashboardConfig): Promise<void> {
  context.workbook.settings.add(`${KEY_PREFIX}${config.name}`, JSON.stringify(config));
  await context.sync();
}

/** Build a dashboard on new sheets and remember it so it can be refreshed later. */
export async function buildDashboard(
  source: DataSourceRef,
  charts: DashChart[],
  options: { title: string; includeKpis: boolean }
): Promise<OperationResult> {
  if (charts.length === 0 && !options.includeKpis) {
    return fail("Pick at least one chart.");
  }
  const title = options.title.trim() || "Dashboard";
  return runExcel(async (context) => {
    const drawn = await drawDashboard(
      context,
      { name: title, title, source, charts, includeKpis: options.includeKpis },
      null
    );
    if (drawn.result.ok) {
      await saveConfig(context, {
        name: drawn.sheetName,
        title,
        source,
        charts,
        includeKpis: options.includeKpis,
        sheetName: drawn.sheetName,
        dataSheetName: drawn.dataSheetName,
        updatedAt: new Date().toISOString(),
      });
    }
    return drawn.result;
  });
}

/** Every dashboard saved in this workbook. */
export async function listDashboards(): Promise<DashboardConfig[]> {
  try {
    return await Excel.run(async (context) => {
      const settings = context.workbook.settings;
      settings.load("items/key,items/value");
      await context.sync();
      const configs: DashboardConfig[] = [];
      for (const item of settings.items) {
        if (!item.key.startsWith(KEY_PREFIX)) continue;
        try {
          const parsed = JSON.parse(String(item.value)) as DashboardConfig;
          if (parsed && parsed.name && Array.isArray(parsed.charts)) configs.push(parsed);
        } catch {
          // A damaged entry is skipped rather than breaking the list.
        }
      }
      return configs.sort((a, b) => a.name.localeCompare(b.name));
    });
  } catch {
    return [];
  }
}

/** Rebuild a saved dashboard from the current contents of its source. */
export async function refreshDashboard(config: DashboardConfig): Promise<OperationResult> {
  return runExcel(async (context) => {
    const drawn = await drawDashboard(context, config, {
      sheetName: config.sheetName,
      dataSheetName: config.dataSheetName,
    });
    if (drawn.result.ok) {
      // The dashboard sheet was deleted and rebuilt under a new name: drop the old entry.
      if (config.name !== drawn.sheetName) {
        const old = context.workbook.settings.getItemOrNullObject(`${KEY_PREFIX}${config.name}`);
        await context.sync();
        if (!old.isNullObject) old.delete();
      }
      await saveConfig(context, {
        ...config,
        name: drawn.sheetName,
        sheetName: drawn.sheetName,
        dataSheetName: drawn.dataSheetName,
        updatedAt: new Date().toISOString(),
      });
    }
    return drawn.result;
  });
}
