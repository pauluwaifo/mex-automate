/**
 * Saying what the numbers actually show.
 *
 * A dashboard answers "what does this look like"; it does not answer "so what".
 * The reader still has to notice that two regions carry three quarters of the
 * revenue, or that the last month broke a four-month climb. These are arithmetic
 * observations, not opinions, so they can be worked out and written down.
 *
 * No model is involved. Each insight is a template filled from a number the
 * caller can check, and anything that would need judgement - why revenue fell,
 * whether that is bad - is left to the reader, who knows the business.
 *
 * Everything here is pure and unit tested in tests/insights.test.ts.
 */

export type InsightKind =
  | "total"
  | "leader"
  | "concentration"
  | "spread"
  | "trend"
  | "latest"
  | "best"
  | "worst"
  | "streak"
  | "gap";

export interface Insight {
  kind: InsightKind;
  text: string;
  /** 0-1. Higher means more worth saying; the picker sorts on it. */
  importance: number;
}

/** One aggregated category: "North", 128400. */
export interface Point {
  label: string;
  value: number;
}

/** One aggregated period, with a key that sorts chronologically. */
export interface TimePoint extends Point {
  sortKey: number;
}

// ---------------------------------------------------------------------------
// Formatting
// ---------------------------------------------------------------------------

/**
 * Numbers as a person would write them in a sentence: thousands separated, and
 * big ones shortened, because "1.2m" reads better mid-sentence than
 * "1,234,567.89".
 */
export function formatValue(value: number): string {
  const magnitude = Math.abs(value);
  if (magnitude >= 1_000_000_000) return `${trim(value / 1_000_000_000)}bn`;
  if (magnitude >= 1_000_000) return `${trim(value / 1_000_000)}m`;
  if (magnitude >= 100_000) return `${trim(value / 1_000)}k`;
  if (magnitude >= 1000) return Math.round(value).toLocaleString("en-US");
  if (Number.isInteger(value)) return String(value);
  return trim(value, 2);
}

function trim(value: number, places = 1): string {
  const rounded = value.toFixed(places);
  return rounded.replace(/\.0+$/, "").replace(/(\.\d*[1-9])0+$/, "$1");
}

/** A share as a percentage, with the precision the size deserves. */
export function formatShare(share: number): string {
  const percent = share * 100;
  if (percent >= 10) return `${Math.round(percent)}%`;
  if (percent >= 1) return `${trim(percent)}%`;
  return `${trim(percent, 2)}%`;
}

/** A change as a signed percentage, or null when there is nothing to divide by. */
export function percentChange(from: number, to: number): number | null {
  if (from === 0) return null;
  return (to - from) / Math.abs(from);
}

function direction(change: number): string {
  return change >= 0 ? "up" : "down";
}

// ---------------------------------------------------------------------------
// Insights about categories
// ---------------------------------------------------------------------------

const MEANINGFUL_CONCENTRATION = 0.6;

/**
 * What the split across a dimension shows: the total, who leads, how
 * concentrated it is, and whether anything is conspicuously empty.
 */
export function categoryInsights(
  measure: string,
  dimension: string,
  points: readonly Point[]
): Insight[] {
  const positive = points.filter((point) => Number.isFinite(point.value));
  if (positive.length === 0) return [];

  const total = positive.reduce((sum, point) => sum + point.value, 0);
  const sorted = [...positive].sort((a, b) => b.value - a.value);
  const insights: Insight[] = [];

  insights.push({
    kind: "total",
    text: `${measure} totals ${formatValue(total)} across ${sorted.length} ${sorted.length === 1 ? dimension.toLowerCase() : pluralize(dimension.toLowerCase())}.`,
    importance: 0.4,
  });

  if (total !== 0 && sorted.length > 1) {
    const leader = sorted[0];
    const share = leader.value / total;
    insights.push({
      kind: "leader",
      text: `${leader.label} is the largest ${dimension.toLowerCase()} at ${formatValue(leader.value)}, ${formatShare(share)} of the total.`,
      importance: 0.55 + Math.min(0.25, share),
    });

    // How many categories it takes to reach most of the total. When a handful
    // carry the business, that is the most useful sentence on the page.
    let running = 0;
    let needed = 0;
    for (const point of sorted) {
      running += point.value;
      needed += 1;
      if (running / total >= MEANINGFUL_CONCENTRATION) break;
    }
    if (needed < sorted.length && needed <= Math.max(3, Math.ceil(sorted.length / 4))) {
      insights.push({
        kind: "concentration",
        text: `The top ${needed} of ${sorted.length} make up ${formatShare(running / total)} of ${measure.toLowerCase()}.`,
        importance: 0.75,
      });
    }

    const smallest = sorted[sorted.length - 1];
    if (smallest.value > 0 && leader.value / smallest.value >= 10) {
      insights.push({
        kind: "spread",
        text: `${leader.label} is ${trim(leader.value / smallest.value)} times ${smallest.label}, the smallest.`,
        importance: 0.5,
      });
    }

    const empty = positive.filter((point) => point.value === 0);
    if (empty.length > 0 && empty.length <= 3) {
      insights.push({
        kind: "gap",
        text: `${empty.map((point) => point.label).join(", ")} recorded no ${measure.toLowerCase()} at all.`,
        importance: 0.6,
      });
    }
  }

  return insights;
}

function pluralize(word: string): string {
  if (/(s|x|z|ch|sh)$/i.test(word)) return `${word}es`;
  if (/[^aeiou]y$/i.test(word)) return `${word.slice(0, -1)}ies`;
  return `${word}s`;
}

// ---------------------------------------------------------------------------
// Insights about time
// ---------------------------------------------------------------------------

/**
 * What a series over time shows: the last move, the overall direction, the best
 * and worst periods, and any run of consecutive rises or falls.
 *
 * The last period is often partial - a month that is only half over - so its
 * drop is reported as a fact about the data, never as a decline in the business.
 */
export function trendInsights(measure: string, points: readonly TimePoint[]): Insight[] {
  const ordered = [...points]
    .filter((point) => Number.isFinite(point.value))
    .sort((a, b) => a.sortKey - b.sortKey);
  if (ordered.length < 2) return [];

  const insights: Insight[] = [];
  const last = ordered[ordered.length - 1];
  const previous = ordered[ordered.length - 2];

  const change = percentChange(previous.value, last.value);
  if (change !== null) {
    insights.push({
      kind: "latest",
      text: `${measure} in ${last.label} is ${formatValue(last.value)}, ${direction(change)} ${formatShare(Math.abs(change))} on ${previous.label}.`,
      importance: 0.85,
    });
  } else {
    insights.push({
      kind: "latest",
      text: `${measure} in ${last.label} is ${formatValue(last.value)}, against nothing in ${previous.label}.`,
      importance: 0.8,
    });
  }

  // First half against second half: a steadier read on direction than the last
  // two points, which can swing on one order.
  if (ordered.length >= 4) {
    const half = Math.floor(ordered.length / 2);
    const early = average(ordered.slice(0, half).map((point) => point.value));
    const late = average(ordered.slice(ordered.length - half).map((point) => point.value));
    const shift = percentChange(early, late);
    if (shift !== null && Math.abs(shift) >= 0.05) {
      insights.push({
        kind: "trend",
        text: `Across ${ordered.length} periods, ${measure.toLowerCase()} is ${direction(shift)} ${formatShare(Math.abs(shift))} comparing the recent half with the earlier half.`,
        importance: 0.7,
      });
    }
  }

  const best = ordered.reduce((top, point) => (point.value > top.value ? point : top), ordered[0]);
  const worst = ordered.reduce((low, point) => (point.value < low.value ? point : low), ordered[0]);
  if (best.label !== worst.label) {
    insights.push({
      kind: "best",
      text: `${best.label} was the strongest period at ${formatValue(best.value)}; ${worst.label} the weakest at ${formatValue(worst.value)}.`,
      importance: 0.5,
    });
  }

  const run = trailingRun(ordered.map((point) => point.value));
  if (Math.abs(run) >= 3) {
    insights.push({
      kind: "streak",
      text: `${measure} has ${run > 0 ? "risen" : "fallen"} ${Math.abs(run)} periods in a row.`,
      importance: 0.8,
    });
  }

  return insights;
}

function average(values: readonly number[]): number {
  return values.length === 0 ? 0 : values.reduce((sum, value) => sum + value, 0) / values.length;
}

/**
 * Length of the run of consecutive rises (positive) or falls (negative) at the
 * end of a series. A run of 3 means the last three steps all went the same way.
 */
export function trailingRun(values: readonly number[]): number {
  if (values.length < 2) return 0;
  const rising = values[values.length - 1] > values[values.length - 2];
  let run = 0;
  for (let at = values.length - 1; at > 0; at -= 1) {
    const up = values[at] > values[at - 1];
    if (values[at] === values[at - 1] || up !== rising) break;
    run += 1;
  }
  // Without this, a flat series returns -0, which reads as a falling run.
  if (run === 0) return 0;
  return rising ? run : -run;
}

// ---------------------------------------------------------------------------
// Choosing what to say
// ---------------------------------------------------------------------------

/**
 * The most worthwhile insights, most important first, with at most one of each
 * kind so the list does not say the same thing three ways.
 */
export function pickInsights(insights: readonly Insight[], limit = 5): Insight[] {
  const seenKinds = new Set<InsightKind>();
  const seenText = new Set<string>();
  const picked: Insight[] = [];
  for (const insight of [...insights].sort((a, b) => b.importance - a.importance)) {
    if (seenKinds.has(insight.kind) || seenText.has(insight.text)) continue;
    seenKinds.add(insight.kind);
    seenText.add(insight.text);
    picked.push(insight);
    if (picked.length >= limit) break;
  }
  return picked;
}
