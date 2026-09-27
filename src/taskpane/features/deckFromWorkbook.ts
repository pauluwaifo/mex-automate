/**
 * From a spreadsheet file to a planned deck, without Office being involved at
 * all.
 *
 * This is the part that makes MEx for PowerPoint worth having: it is the same
 * engine the Excel side uses. The file is repaired by `tidyTable`, its columns
 * profiled, and the charts chosen by `planDashboard` - so the deck says the
 * same things a dashboard built from the same export would say, and the
 * sentences under the charts come from the same arithmetic.
 *
 * Everything here is pure: a grid in, a deck plan out. That means the layout
 * can be tested without a copy of PowerPoint, and the same code runs whether
 * the grid came from a picked file or anywhere else.
 */

import { buildChartData, DashChart, planDashboard, planInsights, suggestKpis } from "./dashboard";
import { ChartSlideInput, DeckInput, DeckPlan, drawKindFor, planDeck, readChartData } from "./deck";
import { categoryInsights, formatValue, pickInsights, trendInsights } from "./insights";
import { tidyTable } from "./tidy";
import type { Grid } from "../shared/types";

/** More than this and the deck is longer than anybody will sit through. */
export const MAX_DECK_CHARTS = 6;

export interface DeckSource {
  /** What to call the source on the slides: the file name, and sheet if useful. */
  label: string;
  grid: Grid;
  /** Number formats parallel to the grid, when the reader supplied them. */
  formats?: readonly (readonly string[])[];
}

export interface DeckBuildOptions {
  title?: string;
  /** Cap the number of chart slides. */
  maxCharts?: number;
  /** Whether the host can rotate shapes, which decides if trends are lines. */
  canRotate?: boolean;
  /** Only use these chart ids, when the user has chosen. */
  useCharts?: string[];
}

/** One sentence about a single chart, from the same numbers the chart draws. */
export function insightForChart(chart: DashChart, data: Grid): string | null {
  const read = readChartData(data);
  if (!read || read.series.length !== 1) return null;

  const points = read.categories.map((label, index) => ({
    label,
    value: read.series[0].values[index] ?? 0,
  }));
  if (points.length < 2) return null;

  const insights = chart.timeGrain
    ? trendInsights(
        read.valueName,
        points.map((point, index) => ({ ...point, sortKey: index }))
      )
    : categoryInsights(read.valueName, chart.dimension ?? "group", points);

  return pickInsights(insights, 1)[0]?.text ?? null;
}

export interface DeckPreview {
  plan: DeckPlan;
  /** What the deck will contain, for showing before it is built. */
  slideTitles: string[];
  rowCount: number;
  problemsFixed: number;
  /** Charts the plan produced but the deck left out, because of the cap. */
  chartsLeftOut: number;
}

/**
 * Reads a spreadsheet grid and plans the whole deck from it.
 *
 * Returns null when the grid holds nothing that can be charted - a file of
 * notes, or a single column of text - rather than producing empty slides.
 */
export function planDeckFromGrid(
  source: DeckSource,
  options: DeckBuildOptions = {}
): DeckPreview | null {
  const tidy = tidyTable(source.grid, { formats: source.formats });
  if (!tidy || tidy.rows.length === 0) return null;

  // tidyTable has already profiled the cleaned columns, and those profiles are
  // what the repairs were based on, so re-deriving them here could only
  // disagree with the table it produced.
  const profiles = tidy.columns;

  const plan = planDashboard(tidy.headers, profiles, tidy.rows);
  if (plan.charts.length === 0) return null;

  const wanted = options.useCharts
    ? plan.charts.filter((chart) => options.useCharts?.includes(chart.id))
    : plan.charts;
  const cap = Math.max(1, options.maxCharts ?? MAX_DECK_CHARTS);
  const chosen = wanted.slice(0, cap);

  const charts: ChartSlideInput[] = [];
  for (const chart of chosen) {
    const grid = buildChartData(chart, profiles, tidy.rows);
    if (!grid || grid.length < 2) continue;
    const data = readChartData(grid);
    if (!data) continue;
    charts.push({
      id: chart.id,
      title: chart.title,
      insight: insightForChart(chart, grid),
      kind: drawKindFor(chart.kind, { canRotate: Boolean(options.canRotate) }),
      data,
    });
  }

  const kpis = suggestKpis(profiles, tidy.rows)
    .slice(0, 4)
    .map((kpi) => ({
      label: kpi.label,
      value: kpi.text ?? formatValue(kpi.value),
      note: kpi.note,
    }));

  const input: DeckInput = {
    title: options.title?.trim() || defaultTitle(source.label),
    source: source.label,
    rowCount: tidy.rows.length,
    kpis,
    charts,
    insights: planInsights(plan, tidy.rows, 4).map((insight) => insight.text),
  };

  const deck = planDeck(input);
  return {
    plan: deck,
    slideTitles: deck.slides.map((slide) => slide.title),
    rowCount: tidy.rows.length,
    problemsFixed: tidy.findings.reduce((total, finding) => total + finding.count, 0),
    chartsLeftOut: Math.max(0, wanted.length - charts.length),
  };
}

/** "sales-march.xlsx" becomes "Sales march", which beats "Untitled deck". */
export function defaultTitle(label: string): string {
  const base = label
    .split(/[·/\\]/)[0]
    .trim()
    .replace(/\.(xlsx|xlsm|csv|txt|tsv)$/i, "");
  const words = base.replace(/[_-]+/g, " ").replace(/\s+/g, " ").trim();
  if (words === "") return "Monthly report";
  return words.charAt(0).toUpperCase() + words.slice(1);
}
