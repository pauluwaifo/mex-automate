import {
  axisTicks,
  Box,
  drawChart,
  drawKindFor,
  niceCeiling,
  parseTag,
  planDeck,
  readChartData,
  serializeTag,
  seriesColor,
  SHAPE_PREFIX,
  SLIDE,
  ShapeSpec,
} from "../src/taskpane/features/deck";
import { Grid } from "../src/taskpane/shared/types";

const BOX: Box = { left: 56, top: 130, width: 848, height: 330 };

const SINGLE: Grid = [
  ["Region", "Revenue"],
  ["North", 500],
  ["South", 300],
  ["East", 150],
  ["West", 50],
];

const MULTI: Grid = [
  ["Month", "Online", "Retail"],
  ["Jan", 100, 60],
  ["Feb", 120, 80],
  ["Mar", 140, 70],
];

/** Everything a shape covers, for checking nothing escapes the slide. */
function bounds(shape: ShapeSpec): { left: number; top: number; right: number; bottom: number } {
  if (shape.kind === "line" || shape.kind === "segment") {
    return {
      left: Math.min(shape.from.x, shape.to.x),
      top: Math.min(shape.from.y, shape.to.y),
      right: Math.max(shape.from.x, shape.to.x),
      bottom: Math.max(shape.from.y, shape.to.y),
    };
  }
  return {
    left: shape.box.left,
    top: shape.box.top,
    right: shape.box.left + shape.box.width,
    bottom: shape.box.top + shape.box.height,
  };
}

describe("axis scaling", () => {
  it("rounds up to a number a person would have chosen", () => {
    expect(niceCeiling(87)).toBe(100);
    expect(niceCeiling(230)).toBe(250);
    expect(niceCeiling(1)).toBe(1);
    expect(niceCeiling(4200)).toBe(5000);
    expect(niceCeiling(12)).toBe(20);
  });

  it("copes with nothing to scale", () => {
    expect(niceCeiling(0)).toBe(1);
    expect(niceCeiling(-5)).toBe(1);
    expect(niceCeiling(NaN)).toBe(1);
  });

  it("spaces ticks evenly from zero", () => {
    expect(axisTicks(87, 4)).toEqual([0, 25, 50, 75, 100]);
  });
});

describe("readChartData", () => {
  it("reads categories and one series", () => {
    const data = readChartData(SINGLE)!;
    expect(data.categories).toEqual(["North", "South", "East", "West"]);
    expect(data.series).toHaveLength(1);
    expect(data.series[0]).toEqual({ label: "Revenue", values: [500, 300, 150, 50] });
  });

  it("reads several series", () => {
    const data = readChartData(MULTI)!;
    expect(data.series.map((entry) => entry.label)).toEqual(["Online", "Retail"]);
    expect(data.series[1].values).toEqual([60, 80, 70]);
  });

  it("returns nothing for a grid with no rows or no values", () => {
    expect(readChartData([["Region", "Revenue"]])).toBeNull();
    expect(readChartData([["Region"], ["North"]])).toBeNull();
    expect(readChartData([])).toBeNull();
  });
});

describe("drawing charts", () => {
  it("draws one bar per category, inside the box", () => {
    const shapes = drawChart(readChartData(SINGLE)!, "column", BOX, "c1");
    const bars = shapes.filter((shape) => shape.name.includes("_bar_"));
    expect(bars).toHaveLength(4);
    for (const shape of shapes) {
      const area = bounds(shape);
      expect(area.left).toBeGreaterThanOrEqual(BOX.left - 1);
      expect(area.right).toBeLessThanOrEqual(BOX.left + BOX.width + 1);
      expect(area.bottom).toBeLessThanOrEqual(BOX.top + BOX.height + 1);
    }
  });

  it("makes bar height proportional to value", () => {
    const shapes = drawChart(readChartData(SINGLE)!, "column", BOX, "c1");
    const heights = shapes
      .filter((shape) => shape.name.includes("_bar_"))
      .map((shape) => (shape.kind === "rect" ? shape.box.height : 0));
    // North is 500 and South 300, so the first bar is taller in that ratio.
    expect(heights[0] / heights[1]).toBeCloseTo(500 / 300, 1);
  });

  it("puts a group of bars side by side for several series", () => {
    const shapes = drawChart(readChartData(MULTI)!, "column", BOX, "c2");
    const bars = shapes.filter((shape) => shape.name.includes("_bar_"));
    expect(bars).toHaveLength(6);
    // The two bars of January must not overlap.
    const jan = bars.filter((shape) => shape.name.includes("_bar_0_"));
    const [first, second] = jan.map((shape) => (shape.kind === "rect" ? shape.box : null));
    expect(first && second && first.left + first.width).toBeLessThanOrEqual(second!.left + 0.01);
  });

  it("draws a line as one segment between each pair of readings", () => {
    const shapes = drawChart(readChartData(MULTI)!, "line", BOX, "c3");
    const segments = shapes.filter((shape) => shape.name.includes("_seg_"));
    // Two series, three points each: two segments per series.
    expect(segments).toHaveLength(4);
    expect(segments.every((shape) => shape.kind === "segment")).toBe(true);
  });

  it("draws a share as one bar whose parts add up to the width", () => {
    const shapes = drawChart(readChartData(SINGLE)!, "share", BOX, "c4");
    const parts = shapes.filter((shape) => shape.name.includes("_part_"));
    expect(parts).toHaveLength(4);
    const total = parts.reduce(
      (sum, shape) => sum + (shape.kind === "rect" ? shape.box.width : 0),
      0
    );
    // Each part loses 2pt to the gap between them.
    expect(total).toBeGreaterThan(BOX.width - 12);
    expect(total).toBeLessThanOrEqual(BOX.width);
  });

  it("labels only the slices wide enough to hold a label", () => {
    const lopsided: Grid = [
      ["Region", "Revenue"],
      ["North", 980],
      ["South", 10],
      ["East", 10],
    ];
    const shapes = drawChart(readChartData(lopsided)!, "share", BOX, "c5");
    expect(shapes.filter((shape) => shape.name.includes("_pct_"))).toHaveLength(1);
  });

  it("thins out the dates under a long trend", () => {
    const long: Grid = [
      ["Month", "Revenue"],
      ...Array.from({ length: 24 }, (_v, i) => [`M${i + 1}`, 100 + i]),
    ];
    const shapes = drawChart(readChartData(long)!, "line", BOX, "c6");
    const labels = shapes.filter((shape) => shape.name.includes("_cat_"));
    expect(labels.length).toBeLessThan(12);
    expect(labels.length).toBeGreaterThan(2);
  });

  it("puts the value at the end of each bar in a horizontal chart", () => {
    const shapes = drawChart(readChartData(SINGLE)!, "bar", BOX, "c7");
    expect(shapes.filter((shape) => shape.name.includes("_val_"))).toHaveLength(4);
  });

  it("draws something rather than nothing when every value is zero", () => {
    const zeros: Grid = [
      ["Region", "Revenue"],
      ["North", 0],
      ["South", 0],
    ];
    const shapes = drawChart(readChartData(zeros)!, "column", BOX, "c8");
    expect(shapes.filter((shape) => shape.name.includes("_bar_"))).toHaveLength(2);
  });
});

describe("colours", () => {
  it("keeps Other neutral wherever it appears", () => {
    expect(seriesColor(0, "Other")).toBe(seriesColor(5, "Other"));
    expect(seriesColor(0, "North")).not.toBe(seriesColor(0, "Other"));
  });

  it("cycles rather than running out", () => {
    expect(seriesColor(99, "x")).toMatch(/^#[0-9A-F]{6}$/i);
  });
});

describe("planDeck", () => {
  const deck = planDeck({
    title: "March sales",
    source: "sales.xlsx",
    rowCount: 2481,
    kpis: [
      { label: "Revenue", value: "128.4k", note: "across 4 regions" },
      { label: "Orders", value: "2,481", note: null },
    ],
    charts: [
      {
        id: "c1",
        title: "Revenue by Region",
        insight: "North is the largest region at 500, 50% of the total.",
        kind: "column",
        data: readChartData(SINGLE)!,
      },
    ],
    insights: ["Revenue is up 14% on last month.", "The top 2 of 4 make up 80%."],
  });

  it("opens with a title, then the numbers, then the evidence, then the meaning", () => {
    expect(deck.slides.map((slide) => slide.kind)).toEqual([
      "title",
      "numbers",
      "chart",
      "insights",
    ]);
  });

  it("says where the numbers came from on every slide", () => {
    expect(deck.subtitle).toContain("sales.xlsx");
    expect(deck.subtitle).toContain("2,481 rows");
  });

  it("names every shape so a refresh can find and replace them", () => {
    for (const slide of deck.slides) {
      for (const shape of slide.shapes) {
        expect(shape.name.startsWith(SHAPE_PREFIX)).toBe(true);
      }
    }
  });

  it("keeps every shape on the slide", () => {
    for (const slide of deck.slides) {
      for (const shape of slide.shapes) {
        const area = bounds(shape);
        expect(area.left).toBeGreaterThanOrEqual(0);
        expect(area.top).toBeGreaterThanOrEqual(0);
        expect(area.right).toBeLessThanOrEqual(SLIDE.width + 1);
        expect(area.bottom).toBeLessThanOrEqual(SLIDE.height + 1);
      }
    }
  });

  it("puts the insight sentence under its own chart", () => {
    const chart = deck.slides.find((slide) => slide.kind === "chart")!;
    const insight = chart.shapes.find((shape) => shape.name.endsWith("_insight"));
    expect(insight && insight.kind === "text" && insight.text).toContain("North is the largest");
  });

  it("leaves out the slides it has nothing for", () => {
    const bare = planDeck({
      title: "Bare",
      source: "x.xlsx",
      rowCount: 3,
      kpis: [],
      charts: [],
      insights: [],
    });
    expect(bare.slides.map((slide) => slide.kind)).toEqual(["title"]);
  });
});

describe("slide tags", () => {
  it("survives a round trip", () => {
    const tag = { chartId: "c1", kind: "chart" as const, source: "sales.xlsx" };
    expect(parseTag(serializeTag(tag))).toEqual(tag);
  });

  it("refuses anything it cannot use", () => {
    expect(parseTag("")).toBeNull();
    expect(parseTag("not json")).toBeNull();
    expect(parseTag(JSON.stringify({ nothing: true }))).toBeNull();
    expect(parseTag(null)).toBeNull();
  });
});

describe("drawKindFor", () => {
  it("draws a share for anything round", () => {
    expect(drawKindFor("pie")).toBe("share");
    expect(drawKindFor("doughnut")).toBe("share");
  });

  it("only draws a real line where the host can rotate a shape", () => {
    // A connector fills its box one way, so without rotation a rising segment
    // would come out upside down. Columns answer the same question correctly.
    expect(drawKindFor("line", { canRotate: false })).toBe("column");
    expect(drawKindFor("line", { canRotate: true })).toBe("line");
  });

  it("falls back to columns for anything it does not know", () => {
    expect(drawKindFor("stackedColumn")).toBe("column");
    expect(drawKindFor("scatter")).toBe("column");
  });

  it("uses sloped segments only in a line chart", () => {
    const columns = drawChart(readChartData(MULTI)!, "column", BOX, "k1");
    expect(columns.some((shape) => shape.kind === "segment")).toBe(false);
    const line = drawChart(readChartData(MULTI)!, "line", BOX, "k2");
    expect(line.some((shape) => shape.kind === "segment")).toBe(true);
    // Gridlines stay flat lines, which a connector can draw correctly.
    expect(
      line
        .filter((shape) => shape.kind === "line")
        .every((shape) => shape.kind === "line" && shape.from.y === shape.to.y)
    ).toBe(true);
  });
});
