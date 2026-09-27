/**
 * Turning a spreadsheet into slides.
 *
 * The monthly reporting deck is the job worth taking off people: build it once
 * by hand, then rebuild it every month because the numbers moved. The usual
 * answers are copy-paste (an hour a month, wrong the moment anything changes)
 * or Paste Link (breaks when the workbook is renamed, moved, or its ranges
 * shift). Both fail the same way - the deck is tied to a *file path* rather
 * than to a *question*.
 *
 * There is no chart object in the PowerPoint JavaScript API. None: you cannot
 * insert a chart, only shapes. That sounds like a problem and is actually the
 * opportunity, because the alternative everyone reaches for - pasting a picture
 * of an Excel chart - gives you something nobody can restyle, that goes soft on
 * a projector and is invisible to a screen reader.
 *
 * So charts here are drawn: rectangles for bars, line segments for trends, a
 * single stacked bar for a share. What lands on the slide is ordinary
 * PowerPoint vector shapes in the deck's own theme, which anyone can nudge,
 * recolour and animate. It also means the whole thing runs on PowerPointApi
 * 1.4, which is widely available, rather than on the newest sets.
 *
 * This module is pure geometry and text: it decides what goes where, in points,
 * and hands a list of shapes to the driver. Unit tested in tests/deck.test.ts.
 */

import { formatValue } from "./insights";
import type { Grid } from "../shared/types";

// ---------------------------------------------------------------------------
// The canvas
// ---------------------------------------------------------------------------

/**
 * A 16:9 slide is 13.333 x 7.5 inches, and the PowerPoint API measures in
 * points at 72 to the inch. Everything below is in points from the top left.
 */
export const SLIDE = { width: 960, height: 540 } as const;

export const MARGIN = 56;
const TITLE_TOP = 48;
const TITLE_HEIGHT = 44;
const SUBTITLE_HEIGHT = 22;
const BODY_TOP = 130;
const FOOTNOTE_HEIGHT = 40;

export interface Box {
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface TextStyle {
  size: number;
  bold?: boolean;
  color: string;
  align?: "left" | "center" | "right";
}

export type ShapeSpec =
  | { kind: "text"; name: string; box: Box; text: string; style: TextStyle }
  | { kind: "rect"; name: string; box: Box; fill: string; transparency?: number }
  /**
   * A horizontal rule. Kept separate from a sloped segment because a connector
   * added through the API only fills its bounding box one way - top-left to
   * bottom-right - so it can draw a flat line correctly and a rising one
   * upside down.
   */
  | {
      kind: "line";
      name: string;
      from: { x: number; y: number };
      to: { x: number; y: number };
      color: string;
      weight: number;
    }
  /**
   * A sloped segment of a trend line, drawn as a thin rotated rectangle. That
   * needs PowerPointApi 1.10, so `drawKindFor` only asks for one where the host
   * supports rotation, and falls back to columns where it does not.
   */
  | {
      kind: "segment";
      name: string;
      from: { x: number; y: number };
      to: { x: number; y: number };
      color: string;
      weight: number;
    };

export type SlideKind = "title" | "numbers" | "chart" | "insights";

export interface SlideSpec {
  kind: SlideKind;
  /** Shown in the pane's list of what will be built. */
  title: string;
  shapes: ShapeSpec[];
  /** Written into the slide's tags, so a refresh knows what to redraw. */
  tag: SlideTag;
}

/** What a generated slide remembers about itself. */
export interface SlideTag {
  /** Which chart of the plan this slide shows, by id. Empty for title/insight slides. */
  chartId: string;
  kind: SlideKind;
  /** The workbook and sheet the numbers came from, for the footnote and refresh. */
  source: string;
}

export interface DeckPlan {
  title: string;
  subtitle: string;
  slides: SlideSpec[];
}

// ---------------------------------------------------------------------------
// Ink
// ---------------------------------------------------------------------------

/** The site's palette, so a deck and a dashboard from the same data match. */
export const INK = {
  heading: "#12201C",
  body: "#4B5C57",
  faint: "#74847F",
  rule: "#DDE5E2",
  panel: "#EEF3F1",
  accent: "#0E7A5C",
} as const;

export const SERIES_COLORS = [
  "#0E7A5C",
  "#2A78D6",
  "#EB6834",
  "#EDA100",
  "#4A3AA7",
  "#1BAF7A",
  "#E87BA4",
  "#E34948",
] as const;

export const OTHER_COLOR = "#9AA3A0";

export function seriesColor(index: number, label?: string): string {
  if (label === "Other" || label === "(blank)") return OTHER_COLOR;
  return SERIES_COLORS[index % SERIES_COLORS.length];
}

// ---------------------------------------------------------------------------
// Axis scales
// ---------------------------------------------------------------------------

/**
 * A round number at or above the largest value, so the axis ends somewhere a
 * person would have chosen: 1, 2, 2.5 or 5 times a power of ten.
 */
export function niceCeiling(value: number): number {
  if (!Number.isFinite(value) || value <= 0) return 1;
  const magnitude = Math.pow(10, Math.floor(Math.log10(value)));
  const normalized = value / magnitude;
  const step =
    normalized <= 1 ? 1 : normalized <= 2 ? 2 : normalized <= 2.5 ? 2.5 : normalized <= 5 ? 5 : 10;
  return step * magnitude;
}

/** Evenly spaced axis values from zero to a round top. */
export function axisTicks(max: number, count = 4): number[] {
  const top = niceCeiling(max);
  const ticks: number[] = [];
  for (let i = 0; i <= count; i += 1) ticks.push((top / count) * i);
  return ticks;
}

// ---------------------------------------------------------------------------
// Reading the data a chart was built from
// ---------------------------------------------------------------------------

export interface Series {
  label: string;
  values: number[];
}

export interface ChartData {
  /** Category labels along the axis. */
  categories: string[];
  series: Series[];
  /** What the numbers are, for the axis title. */
  valueName: string;
}

/**
 * Reads the grid the Excel side already builds for a chart: a heading row, then
 * one row per category, with one column per series.
 */
export function readChartData(grid: Grid): ChartData | null {
  if (grid.length < 2 || grid[0].length < 2) return null;
  const headings = grid[0].map((value) => String(value ?? ""));
  const categories: string[] = [];
  const series: Series[] = headings.slice(1).map((label) => ({ label, values: [] }));

  for (const row of grid.slice(1)) {
    categories.push(String(row[0] ?? ""));
    series.forEach((entry, index) => {
      const value = row[index + 1];
      entry.values.push(typeof value === "number" ? value : Number(value) || 0);
    });
  }
  if (categories.length === 0) return null;
  return { categories, series, valueName: headings[1] ?? "Value" };
}

/** Every value across every series, for scaling the axis. */
function allValues(data: ChartData): number[] {
  return data.series.flatMap((entry) => entry.values);
}

// ---------------------------------------------------------------------------
// Drawing a chart
// ---------------------------------------------------------------------------

export type DrawKind = "column" | "bar" | "line" | "share";

/**
 * Which drawing to use for a chart the Excel side suggested.
 *
 * A trend is only drawn as a line where the host can rotate a shape, because
 * without rotation a rising segment cannot be drawn the right way up. Columns
 * over months answer the same question and are never wrong, so that is the
 * fallback rather than a line that slopes the wrong way.
 */
export function drawKindFor(
  chartKind: string,
  options: { canRotate: boolean } = { canRotate: false }
): DrawKind {
  switch (chartKind) {
    case "pie":
    case "doughnut":
      return "share";
    case "bar":
      return "bar";
    case "line":
      return options.canRotate ? "line" : "column";
    default:
      return "column";
  }
}

const AXIS_LABEL_SIZE = 10;
const CATEGORY_LABEL_SIZE = 10;

/** How much room the value labels down the left need. */
const AXIS_WIDTH = 64;
/** How much room the category labels along the bottom need. */
const CATEGORY_HEIGHT = 34;

/**
 * Lays a chart out inside a box. Returns plain shape specs; nothing here knows
 * about PowerPoint.
 */
export function drawChart(data: ChartData, kind: DrawKind, box: Box, prefix: string): ShapeSpec[] {
  switch (kind) {
    case "share":
      return drawShare(data, box, prefix);
    case "bar":
      return drawBars(data, box, prefix);
    case "line":
      return drawLine(data, box, prefix);
    default:
      return drawColumns(data, box, prefix);
  }
}

/** The plotting rectangle, once the axis and labels have taken their room. */
function plotArea(box: Box, kind: DrawKind): Box {
  if (kind === "bar") {
    return {
      left: box.left + AXIS_WIDTH + 40,
      top: box.top,
      width: box.width - AXIS_WIDTH - 40,
      height: box.height - CATEGORY_HEIGHT / 2,
    };
  }
  return {
    left: box.left + AXIS_WIDTH,
    top: box.top,
    width: box.width - AXIS_WIDTH,
    height: box.height - CATEGORY_HEIGHT,
  };
}

function gridlines(plot: Box, ticks: number[], top: number, prefix: string): ShapeSpec[] {
  const shapes: ShapeSpec[] = [];
  ticks.forEach((tick, index) => {
    const y = plot.top + plot.height - (top === 0 ? 0 : (tick / top) * plot.height);
    shapes.push({
      kind: "line",
      name: `${prefix}_grid_${index}`,
      from: { x: plot.left, y },
      to: { x: plot.left + plot.width, y },
      color: INK.rule,
      weight: index === 0 ? 1.25 : 0.75,
    });
    shapes.push({
      kind: "text",
      name: `${prefix}_axis_${index}`,
      box: { left: plot.left - AXIS_WIDTH, top: y - 9, width: AXIS_WIDTH - 8, height: 18 },
      text: formatValue(tick),
      style: { size: AXIS_LABEL_SIZE, color: INK.faint, align: "right" },
    });
  });
  return shapes;
}

function drawColumns(data: ChartData, box: Box, prefix: string): ShapeSpec[] {
  const plot = plotArea(box, "column");
  const top = niceCeiling(Math.max(0, ...allValues(data)));
  const shapes = gridlines(plot, axisTicks(top), top, prefix);

  const groups = data.categories.length;
  const seriesCount = Math.max(1, data.series.length);
  const slot = plot.width / Math.max(1, groups);
  // A quarter of each slot is breathing room, split either side of the group.
  const groupWidth = slot * 0.72;
  const barWidth = groupWidth / seriesCount;

  data.categories.forEach((category, groupIndex) => {
    const groupLeft = plot.left + slot * groupIndex + (slot - groupWidth) / 2;

    data.series.forEach((entry, seriesIndex) => {
      const value = entry.values[groupIndex] ?? 0;
      const height = top === 0 ? 0 : (Math.max(0, value) / top) * plot.height;
      shapes.push({
        kind: "rect",
        name: `${prefix}_bar_${groupIndex}_${seriesIndex}`,
        box: {
          left: groupLeft + barWidth * seriesIndex,
          top: plot.top + plot.height - height,
          width: Math.max(2, barWidth - 2),
          height: Math.max(1, height),
        },
        fill: seriesColor(seriesIndex, seriesCount === 1 ? category : entry.label),
      });
    });

    shapes.push({
      kind: "text",
      name: `${prefix}_cat_${groupIndex}`,
      box: {
        left: plot.left + slot * groupIndex,
        top: plot.top + plot.height + 6,
        width: slot,
        height: CATEGORY_HEIGHT - 6,
      },
      text: category,
      style: { size: CATEGORY_LABEL_SIZE, color: INK.body, align: "center" },
    });
  });

  return shapes;
}

function drawBars(data: ChartData, box: Box, prefix: string): ShapeSpec[] {
  const plot = plotArea(box, "bar");
  const top = niceCeiling(Math.max(0, ...allValues(data)));
  const shapes: ShapeSpec[] = [];

  const rows = data.categories.length;
  const slot = plot.height / Math.max(1, rows);
  const barHeight = Math.min(34, slot * 0.62);

  data.categories.forEach((category, index) => {
    const value = data.series[0]?.values[index] ?? 0;
    const width = top === 0 ? 0 : (Math.max(0, value) / top) * plot.width;
    const centre = plot.top + slot * index + slot / 2;

    shapes.push({
      kind: "text",
      name: `${prefix}_cat_${index}`,
      box: { left: box.left, top: centre - 10, width: AXIS_WIDTH + 32, height: 20 },
      text: category,
      style: { size: CATEGORY_LABEL_SIZE, color: INK.body, align: "right" },
    });
    shapes.push({
      kind: "rect",
      name: `${prefix}_bar_${index}`,
      box: {
        left: plot.left,
        top: centre - barHeight / 2,
        width: Math.max(1, width),
        height: barHeight,
      },
      fill: seriesColor(index, category),
    });
    // The value goes at the end of its own bar, which saves an axis entirely.
    shapes.push({
      kind: "text",
      name: `${prefix}_val_${index}`,
      box: { left: plot.left + width + 6, top: centre - 9, width: 90, height: 18 },
      text: formatValue(value),
      style: { size: AXIS_LABEL_SIZE, color: INK.faint, align: "left" },
    });
  });

  return shapes;
}

function drawLine(data: ChartData, box: Box, prefix: string): ShapeSpec[] {
  const plot = plotArea(box, "line");
  const values = allValues(data);
  const top = niceCeiling(Math.max(0, ...values));
  const shapes = gridlines(plot, axisTicks(top), top, prefix);

  const points = data.categories.length;
  const step = points > 1 ? plot.width / (points - 1) : 0;
  const yOf = (value: number) =>
    plot.top + plot.height - (top === 0 ? 0 : (Math.max(0, value) / top) * plot.height);

  data.series.forEach((entry, seriesIndex) => {
    const color = seriesColor(seriesIndex, entry.label);
    for (let index = 0; index < points - 1; index += 1) {
      shapes.push({
        kind: "segment",
        name: `${prefix}_seg_${seriesIndex}_${index}`,
        from: { x: plot.left + step * index, y: yOf(entry.values[index] ?? 0) },
        to: { x: plot.left + step * (index + 1), y: yOf(entry.values[index + 1] ?? 0) },
        color,
        weight: 2.5,
      });
    }
    // A dot at each reading, so a single point still shows and the eye can
    // follow the line's turns.
    entry.values.forEach((value, index) => {
      const size = 7;
      shapes.push({
        kind: "rect",
        name: `${prefix}_dot_${seriesIndex}_${index}`,
        box: {
          left: plot.left + step * index - size / 2,
          top: yOf(value) - size / 2,
          width: size,
          height: size,
        },
        fill: color,
      });
    });
  });

  // Too many dates along the bottom become unreadable, so only some are shown.
  const everyNth = Math.ceil(points / 8);
  data.categories.forEach((category, index) => {
    if (index % everyNth !== 0 && index !== points - 1) return;
    shapes.push({
      kind: "text",
      name: `${prefix}_cat_${index}`,
      box: {
        left: plot.left + step * index - 40,
        top: plot.top + plot.height + 6,
        width: 80,
        height: CATEGORY_HEIGHT - 6,
      },
      text: category,
      style: { size: CATEGORY_LABEL_SIZE, color: INK.body, align: "center" },
    });
  });

  return shapes;
}

/**
 * A share, drawn as one bar across the slide rather than a pie.
 *
 * PowerPoint's pie geometry needs the adjustments API to set its angles, which
 * is too new to rely on. A single stacked bar answers the same question - how
 * the total splits - and is easier to read than a pie anyway.
 */
function drawShare(data: ChartData, box: Box, prefix: string): ShapeSpec[] {
  const values = data.series[0]?.values ?? [];
  const total = values.reduce((sum, value) => sum + Math.max(0, value), 0);
  const shapes: ShapeSpec[] = [];
  if (total <= 0) return shapes;

  const barTop = box.top + 30;
  const barHeight = 58;
  let x = box.left;

  values.forEach((value, index) => {
    const share = Math.max(0, value) / total;
    const width = share * box.width;
    const color = seriesColor(index, data.categories[index]);
    shapes.push({
      kind: "rect",
      name: `${prefix}_part_${index}`,
      box: { left: x, top: barTop, width: Math.max(1, width - 2), height: barHeight },
      fill: color,
    });
    // A slice narrower than this cannot hold its own label legibly.
    if (width >= 54) {
      shapes.push({
        kind: "text",
        name: `${prefix}_pct_${index}`,
        box: { left: x, top: barTop + barHeight / 2 - 11, width: width - 2, height: 22 },
        text: `${Math.round(share * 100)}%`,
        style: { size: 13, bold: true, color: "#FFFFFF", align: "center" },
      });
    }
    x += width;
  });

  // The key sits under the bar, in the same order as the bar itself.
  let keyX = box.left;
  const keyTop = barTop + barHeight + 18;
  values.forEach((value, index) => {
    const label = `${data.categories[index]} · ${formatValue(value)}`;
    const width = Math.min(210, Math.max(110, label.length * 6.2));
    if (keyX + width > box.left + box.width) return;
    shapes.push({
      kind: "rect",
      name: `${prefix}_keydot_${index}`,
      box: { left: keyX, top: keyTop + 4, width: 10, height: 10 },
      fill: seriesColor(index, data.categories[index]),
    });
    shapes.push({
      kind: "text",
      name: `${prefix}_key_${index}`,
      box: { left: keyX + 16, top: keyTop - 2, width: width - 16, height: 20 },
      text: label,
      style: { size: 11, color: INK.body, align: "left" },
    });
    keyX += width + 12;
  });

  return shapes;
}

/** A key for a chart with more than one series. */
export function seriesKey(data: ChartData, box: Box, prefix: string): ShapeSpec[] {
  if (data.series.length < 2) return [];
  const shapes: ShapeSpec[] = [];
  let x = box.left;
  data.series.forEach((entry, index) => {
    const width = Math.min(200, Math.max(90, entry.label.length * 6.4 + 22));
    if (x + width > box.left + box.width) return;
    shapes.push({
      kind: "rect",
      name: `${prefix}_legdot_${index}`,
      box: { left: x, top: box.top + 5, width: 10, height: 10 },
      fill: seriesColor(index, entry.label),
    });
    shapes.push({
      kind: "text",
      name: `${prefix}_leg_${index}`,
      box: { left: x + 16, top: box.top, width: width - 16, height: 20 },
      text: entry.label,
      style: { size: 11, color: INK.body, align: "left" },
    });
    x += width + 10;
  });
  return shapes;
}

// ---------------------------------------------------------------------------
// Slides
// ---------------------------------------------------------------------------

function heading(title: string, subtitle: string | null, prefix: string): ShapeSpec[] {
  const shapes: ShapeSpec[] = [
    {
      kind: "text",
      name: `${prefix}_title`,
      box: { left: MARGIN, top: TITLE_TOP, width: SLIDE.width - MARGIN * 2, height: TITLE_HEIGHT },
      text: title,
      style: { size: 28, bold: true, color: INK.heading, align: "left" },
    },
  ];
  if (subtitle) {
    shapes.push({
      kind: "text",
      name: `${prefix}_subtitle`,
      box: {
        left: MARGIN,
        top: TITLE_TOP + TITLE_HEIGHT,
        width: SLIDE.width - MARGIN * 2,
        height: SUBTITLE_HEIGHT,
      },
      text: subtitle,
      style: { size: 13, color: INK.faint, align: "left" },
    });
  }
  return shapes;
}

export interface KpiSpec {
  label: string;
  value: string;
  note: string | null;
}

function titleSlide(title: string, subtitle: string, source: string): SlideSpec {
  const prefix = "MEx_title";
  return {
    kind: "title",
    title,
    tag: { chartId: "", kind: "title", source },
    shapes: [
      {
        kind: "rect",
        name: `${prefix}_rule`,
        box: { left: MARGIN, top: 214, width: 96, height: 6 },
        fill: INK.accent,
      },
      {
        kind: "text",
        name: `${prefix}_headline`,
        box: { left: MARGIN, top: 238, width: SLIDE.width - MARGIN * 2, height: 80 },
        text: title,
        style: { size: 44, bold: true, color: INK.heading, align: "left" },
      },
      {
        kind: "text",
        name: `${prefix}_sub`,
        box: { left: MARGIN, top: 322, width: SLIDE.width - MARGIN * 2, height: 30 },
        text: subtitle,
        style: { size: 15, color: INK.body, align: "left" },
      },
    ],
  };
}

function numbersSlide(kpis: readonly KpiSpec[], source: string): SlideSpec {
  const prefix = "MEx_kpi";
  const shapes = heading("The headline numbers", source, prefix);
  const count = Math.min(4, kpis.length);
  const gap = 20;
  const width = (SLIDE.width - MARGIN * 2 - gap * (count - 1)) / count;

  kpis.slice(0, count).forEach((kpi, index) => {
    const left = MARGIN + (width + gap) * index;
    shapes.push({
      kind: "rect",
      name: `${prefix}_card_${index}`,
      box: { left, top: BODY_TOP + 20, width, height: 168 },
      fill: INK.panel,
    });
    shapes.push({
      kind: "text",
      name: `${prefix}_label_${index}`,
      box: { left: left + 20, top: BODY_TOP + 44, width: width - 40, height: 22 },
      text: kpi.label,
      style: { size: 12, color: INK.body, align: "left" },
    });
    shapes.push({
      kind: "text",
      name: `${prefix}_value_${index}`,
      box: { left: left + 20, top: BODY_TOP + 70, width: width - 40, height: 58 },
      text: kpi.value,
      style: { size: 34, bold: true, color: INK.heading, align: "left" },
    });
    if (kpi.note) {
      shapes.push({
        kind: "text",
        name: `${prefix}_note_${index}`,
        box: { left: left + 20, top: BODY_TOP + 130, width: width - 40, height: 40 },
        text: kpi.note,
        style: { size: 11, color: INK.faint, align: "left" },
      });
    }
  });

  return {
    kind: "numbers",
    title: "The headline numbers",
    tag: { chartId: "", kind: "numbers", source },
    shapes,
  };
}

export interface ChartSlideInput {
  id: string;
  title: string;
  /** The sentence that goes under the chart, when there is one. */
  insight: string | null;
  kind: DrawKind;
  data: ChartData;
}

function chartSlide(input: ChartSlideInput, source: string): SlideSpec {
  const prefix = `MEx_chart_${input.id}`;
  const shapes = heading(input.title, source, prefix);

  const hasKey = input.data.series.length > 1 && input.kind !== "share";
  const keyHeight = hasKey ? 26 : 0;
  const plotBox: Box = {
    left: MARGIN,
    top: BODY_TOP + keyHeight,
    width: SLIDE.width - MARGIN * 2,
    height: SLIDE.height - BODY_TOP - FOOTNOTE_HEIGHT - 34 - keyHeight,
  };

  if (hasKey) {
    shapes.push(
      ...seriesKey(
        input.data,
        { left: MARGIN, top: BODY_TOP - 4, width: plotBox.width, height: 20 },
        prefix
      )
    );
  }
  shapes.push(...drawChart(input.data, input.kind, plotBox, prefix));

  if (input.insight) {
    shapes.push({
      kind: "text",
      name: `${prefix}_insight`,
      box: {
        left: MARGIN,
        top: SLIDE.height - FOOTNOTE_HEIGHT - 12,
        width: SLIDE.width - MARGIN * 2,
        height: FOOTNOTE_HEIGHT,
      },
      text: input.insight,
      style: { size: 13, color: INK.body, align: "left" },
    });
  }

  return {
    kind: "chart",
    title: input.title,
    tag: { chartId: input.id, kind: "chart", source },
    shapes,
  };
}

function insightsSlide(lines: readonly string[], source: string): SlideSpec {
  const prefix = "MEx_insights";
  const shapes = heading("What the numbers show", source, prefix);
  lines.slice(0, 5).forEach((line, index) => {
    const top = BODY_TOP + 20 + index * 62;
    shapes.push({
      kind: "rect",
      name: `${prefix}_dot_${index}`,
      box: { left: MARGIN, top: top + 8, width: 10, height: 10 },
      fill: INK.accent,
    });
    shapes.push({
      kind: "text",
      name: `${prefix}_line_${index}`,
      box: { left: MARGIN + 24, top, width: SLIDE.width - MARGIN * 2 - 24, height: 52 },
      text: line,
      style: { size: 16, color: INK.heading, align: "left" },
    });
  });
  return {
    kind: "insights",
    title: "What the numbers show",
    tag: { chartId: "", kind: "insights", source },
    shapes,
  };
}

export interface DeckInput {
  title: string;
  /** Where the numbers came from: file name, and sheet if it matters. */
  source: string;
  rowCount: number;
  kpis: KpiSpec[];
  charts: ChartSlideInput[];
  insights: string[];
}

/**
 * Lays out the whole deck. The order is the order someone reads a report in:
 * what this is, the numbers that matter, then the evidence, then what it means.
 */
export function planDeck(input: DeckInput): DeckPlan {
  const built = new Date().toLocaleDateString("en-US", {
    day: "numeric",
    month: "long",
    year: "numeric",
  });
  const subtitle = `${input.source} · ${input.rowCount.toLocaleString("en-US")} rows · ${built}`;

  const slides: SlideSpec[] = [titleSlide(input.title, subtitle, input.source)];
  if (input.kpis.length > 0) slides.push(numbersSlide(input.kpis, subtitle));
  for (const chart of input.charts) slides.push(chartSlide(chart, subtitle));
  if (input.insights.length > 0) slides.push(insightsSlide(input.insights, subtitle));

  return { title: input.title, subtitle, slides };
}

/** Every shape MEx draws is named with this prefix, so a refresh can find them. */
export const SHAPE_PREFIX = "MEx_";

/** The tag key a generated slide carries. */
export const SLIDE_TAG_KEY = "MEXAUTOMATE";

export function serializeTag(tag: SlideTag): string {
  return JSON.stringify(tag);
}

export function parseTag(raw: unknown): SlideTag | null {
  if (typeof raw !== "string" || raw === "") return null;
  try {
    const parsed = JSON.parse(raw) as Partial<SlideTag>;
    if (typeof parsed.kind !== "string") return null;
    return {
      chartId: typeof parsed.chartId === "string" ? parsed.chartId : "",
      kind: parsed.kind as SlideKind,
      source: typeof parsed.source === "string" ? parsed.source : "",
    };
  } catch {
    return null;
  }
}
