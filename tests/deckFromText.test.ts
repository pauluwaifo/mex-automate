import { planDeckFromText } from "../src/taskpane/features/deckFromText";
import { SLIDE } from "../src/taskpane/features/deck";

/** The kind of thing someone actually pastes: notes, then a table from Excel. */
const NOTES = `Q1 board update

Where we are
- Revenue up 14% on last quarter
- Two new enterprise clients
- Churn down to 3%

Where we are going
1. Hire two engineers
2. Open the Leeds office
    - Lease signed in June
    - Fit-out over the summer
3. Ship the API

Revenue by region
Region\tQ1\tQ2
North\t1200\t1400
South\t900\t1100
East\t600\t700`;

describe("planDeckFromText", () => {
  const preview = planDeckFromText(NOTES)!;

  it("turns a paste into a deck", () => {
    expect(preview).not.toBeNull();
    expect(preview.plan.slides.length).toBeGreaterThan(3);
  });

  it("takes its title from the first heading", () => {
    expect(preview.plan.title).toBe("Q1 board update");
    expect(preview.plan.slides[0].kind).toBe("title");
  });

  it("makes a slide per heading, in the order they were written", () => {
    const titles = preview.plan.slides.map((slide) => slide.title);
    expect(titles).toEqual(
      expect.arrayContaining(["Where we are", "Where we are going", "Revenue by region"])
    );
    expect(titles.indexOf("Where we are")).toBeLessThan(titles.indexOf("Where we are going"));
  });

  it("puts every pasted word on a slide, and invents none", () => {
    const text = preview.plan.slides
      .flatMap((slide) => slide.shapes)
      .filter((shape) => shape.kind === "text")
      .map((shape) => (shape.kind === "text" ? shape.text : ""))
      .join(" ");
    for (const phrase of [
      "Revenue up 14% on last quarter",
      "Two new enterprise clients",
      "Hire two engineers",
      "Lease signed in June",
    ]) {
      expect(text).toContain(phrase);
    }
  });

  it("keeps the numbering of a numbered list", () => {
    const slide = preview.plan.slides.find(
      (candidate) => candidate.title === "Where we are going"
    )!;
    const numbers = slide.shapes
      .filter((shape) => shape.name.includes("_num_"))
      .map((shape) => (shape.kind === "text" ? shape.text : ""));
    expect(numbers).toEqual(["1.", "2.", "3."]);
  });

  it("draws a table for the pasted table, and a chart of its numbers", () => {
    const kinds = preview.plan.slides.map((slide) => slide.kind);
    expect(kinds).toContain("table");
    expect(kinds).toContain("chart");
    expect(preview.chartsAdded).toBe(1);
  });

  it("can be told to leave the charts out", () => {
    const plain = planDeckFromText(NOTES, { chartTables: false })!;
    expect(plain.chartsAdded).toBe(0);
    expect(plain.plan.slides.map((slide) => slide.kind)).toContain("table");
  });

  it("can be told to skip the title slide", () => {
    const plain = planDeckFromText(NOTES, { titleSlide: false })!;
    expect(plain.plan.slides[0].kind).not.toBe("title");
  });

  it("takes a title when it is given one", () => {
    expect(planDeckFromText(NOTES, { title: "Board pack" })!.plan.title).toBe("Board pack");
  });

  it("keeps every shape on the slide", () => {
    for (const slide of preview.plan.slides) {
      for (const shape of slide.shapes) {
        if (shape.kind === "line" || shape.kind === "segment") {
          expect(Math.max(shape.from.x, shape.to.x)).toBeLessThanOrEqual(SLIDE.width + 1);
          expect(Math.max(shape.from.y, shape.to.y)).toBeLessThanOrEqual(SLIDE.height + 1);
          continue;
        }
        expect(shape.box.left).toBeGreaterThanOrEqual(-1);
        expect(shape.box.left + shape.box.width).toBeLessThanOrEqual(SLIDE.width + 1);
        expect(shape.box.top).toBeGreaterThanOrEqual(-1);
        expect(shape.box.top + shape.box.height).toBeLessThanOrEqual(SLIDE.height + 1);
      }
    }
  });

  it("names every shape, so a refresh can find them", () => {
    for (const slide of preview.plan.slides) {
      for (const shape of slide.shapes) {
        expect(shape.name.startsWith("MEx_")).toBe(true);
      }
    }
  });

  it("splits a list too long for one slide", () => {
    const many = Array.from({ length: 15 }, (_v, i) => `- point ${i + 1}`).join("\n");
    const long = planDeckFromText(`Everything\n${many}`)!;
    const pointSlides = long.plan.slides.filter((slide) => slide.kind === "points");
    expect(pointSlides.length).toBeGreaterThan(1);
    expect(pointSlides[1].title).toContain("(cont.)");
  });

  it("copes with a bare list and no headings at all", () => {
    const bare = planDeckFromText("- first thing\n- second thing")!;
    expect(bare.plan.slides.some((slide) => slide.kind === "points")).toBe(true);
  });

  it("says no to an empty paste", () => {
    expect(planDeckFromText("")).toBeNull();
    expect(planDeckFromText("\n\n   \n")).toBeNull();
  });

  it("summarizes what it read", () => {
    expect(preview.summary).toContain("slides");
  });
});
