import {
  defaultTitle,
  MAX_DECK_CHARTS,
  planDeckFromGrid,
} from "../src/taskpane/features/deckFromWorkbook";
import { dateToExcelSerial } from "../src/taskpane/features/dataCleaning";
import { SLIDE } from "../src/taskpane/features/deck";
import { Grid } from "../src/taskpane/shared/types";

const serial = (month: number, day: number) =>
  dateToExcelSerial(new Date(Date.UTC(2024, month, day)));

/**
 * An export shaped like the ones people actually have: a title row, a blank
 * line, then the table - with money written as text and a stray total at the
 * bottom.
 */
const MESSY: Grid = [
  ["Sales export - March 2024", null, null, null, null],
  [null, null, null, null, null],
  ["Order ID", "Order Date", "Region", "Product", "Revenue"],
  ...Array.from({ length: 48 }, (_v, i) => [
    `SO-${2000 + i}`,
    serial(i % 6, 1 + (i % 27)),
    ["North", "South", "East", "West"][i % 4],
    ["Widget", "Gadget", "Sprocket"][i % 3],
    // Half the amounts arrive as text with a currency symbol, as they do.
    i % 2 === 0 ? 100 + i * 7 : `$${(100 + i * 7).toFixed(2)}`,
  ]),
  [null, null, null, null, null],
  ["Total", null, null, null, 9000],
];

describe("planDeckFromGrid", () => {
  const preview = planDeckFromGrid({ label: "sales-march.xlsx", grid: MESSY });

  it("builds a deck straight from a messy export", () => {
    expect(preview).not.toBeNull();
    expect(preview!.rowCount).toBe(48);
    // The title row, the blank line and the total row are not data.
    expect(preview!.problemsFixed).toBeGreaterThan(0);
  });

  it("opens with a title and ends with what the numbers show", () => {
    const kinds = preview!.plan.slides.map((slide) => slide.kind);
    expect(kinds[0]).toBe("title");
    expect(kinds).toContain("numbers");
    expect(kinds).toContain("chart");
    expect(kinds[kinds.length - 1]).toBe("insights");
  });

  it("names the deck after the file", () => {
    expect(preview!.plan.title).toBe("Sales march");
  });

  it("says how many rows and where they came from, on the slides", () => {
    expect(preview!.plan.subtitle).toContain("sales-march.xlsx");
    expect(preview!.plan.subtitle).toContain("48 rows");
  });

  it("writes a sentence under charts it can describe", () => {
    const chartSlides = preview!.plan.slides.filter((slide) => slide.kind === "chart");
    const withInsight = chartSlides.filter((slide) =>
      slide.shapes.some((shape) => shape.name.endsWith("_insight"))
    );
    expect(withInsight.length).toBeGreaterThan(0);
  });

  it("keeps the deck to a length someone will sit through", () => {
    const chartSlides = preview!.plan.slides.filter((slide) => slide.kind === "chart");
    expect(chartSlides.length).toBeLessThanOrEqual(MAX_DECK_CHARTS);
  });

  it("can be told to use fewer charts", () => {
    const small = planDeckFromGrid({ label: "sales.xlsx", grid: MESSY }, { maxCharts: 2 });
    expect(small!.plan.slides.filter((slide) => slide.kind === "chart")).toHaveLength(2);
  });

  it("takes a title when it is given one", () => {
    const named = planDeckFromGrid(
      { label: "sales.xlsx", grid: MESSY },
      { title: "Board pack, March" }
    );
    expect(named!.plan.title).toBe("Board pack, March");
  });

  it("draws trends as columns unless the host can rotate a shape", () => {
    const flat = planDeckFromGrid({ label: "s.xlsx", grid: MESSY }, { canRotate: false });
    const rotated = planDeckFromGrid({ label: "s.xlsx", grid: MESSY }, { canRotate: true });
    const segments = (preview: typeof flat) =>
      preview!.plan.slides.flatMap((slide) =>
        slide.shapes.filter((shape) => shape.kind === "segment")
      ).length;
    expect(segments(flat)).toBe(0);
    expect(segments(rotated)).toBeGreaterThan(0);
  });

  it("keeps every shape on the slide", () => {
    for (const slide of preview!.plan.slides) {
      for (const shape of slide.shapes) {
        if (shape.kind === "line" || shape.kind === "segment") {
          expect(Math.max(shape.from.x, shape.to.x)).toBeLessThanOrEqual(SLIDE.width + 1);
          expect(Math.max(shape.from.y, shape.to.y)).toBeLessThanOrEqual(SLIDE.height + 1);
          continue;
        }
        expect(shape.box.left).toBeGreaterThanOrEqual(-1);
        expect(shape.box.left + shape.box.width).toBeLessThanOrEqual(SLIDE.width + 1);
        expect(shape.box.top + shape.box.height).toBeLessThanOrEqual(SLIDE.height + 1);
      }
    }
  });

  it("gives every chart slide a tag naming what it shows", () => {
    for (const slide of preview!.plan.slides.filter((s) => s.kind === "chart")) {
      expect(slide.tag.chartId).not.toBe("");
      expect(slide.tag.source).toContain("sales-march.xlsx");
    }
  });

  it("says no rather than making empty slides", () => {
    expect(planDeckFromGrid({ label: "notes.txt", grid: [["just a note"]] })).toBeNull();
    expect(planDeckFromGrid({ label: "empty.xlsx", grid: [] })).toBeNull();
  });
});

describe("defaultTitle", () => {
  it("makes a readable title out of a file name", () => {
    expect(defaultTitle("sales-march.xlsx")).toBe("Sales march");
    expect(defaultTitle("Q1_board_pack.csv")).toBe("Q1 board pack");
    expect(defaultTitle("report.xlsx · Sheet1")).toBe("Report");
  });

  it("falls back to something sensible", () => {
    expect(defaultTitle("")).toBe("Monthly report");
    expect(defaultTitle(".xlsx")).toBe("Monthly report");
  });
});
