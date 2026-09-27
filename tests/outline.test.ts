import {
  chartableColumns,
  describeOutline,
  MAX_BULLETS_PER_SLIDE,
  parseOutline,
  tableToGrid,
} from "../src/taskpane/features/outline";

const titles = (text: string) => parseOutline(text).map((slide) => slide.title);

describe("parseOutline", () => {
  it("reads a heading with a list under it as one slide", () => {
    const slides = parseOutline(`Next steps
- Sign the contract
- Move the servers
- Tell the team`);
    expect(slides).toHaveLength(1);
    expect(slides[0].title).toBe("Next steps");
    expect(slides[0].bullets.map((bullet) => bullet.text)).toEqual([
      "Sign the contract",
      "Move the servers",
      "Tell the team",
    ]);
  });

  it("starts a new slide at every blank-line-separated heading", () => {
    expect(
      titles(`Where we are
- Revenue up 14%
- Two new clients

Where we are going
- Hire two people
- Open in Leeds`)
    ).toEqual(["Where we are", "Where we are going"]);
  });

  it("reads markdown headings", () => {
    expect(
      titles(`# Q1 review
- Good
## Q2 plan
- Better`)
    ).toEqual(["Q1 review", "Q2 plan"]);
  });

  it("drops the trailing colon people write after a heading", () => {
    expect(titles("Risks:\n- Supply\n- Staffing")).toEqual(["Risks"]);
  });

  it("does not mistake a sentence for a heading", () => {
    const slides = parseOutline(`Revenue grew by fourteen percent this quarter.
- Driven by the north
- And by two new clients`);
    expect(slides[0].title).not.toBe("Revenue grew by fourteen percent this quarter.");
    expect(slides[0].bullets).toHaveLength(3);
  });

  it("recognises the bullet characters people actually paste", () => {
    const slides = parseOutline(`Points
- dash
* star
• round
– en dash
+ plus`);
    expect(slides[0].bullets.map((bullet) => bullet.text)).toEqual([
      "dash",
      "star",
      "round",
      "en dash",
      "plus",
    ]);
  });

  it("keeps numbered lists numbered", () => {
    const slides = parseOutline(`Process
1. Gather the data
2. Check it
3) Publish`);
    expect(slides[0].bullets.map((bullet) => bullet.number)).toEqual([1, 2, 3]);
  });

  it("reads indentation as sub-points", () => {
    const slides = parseOutline(`Plan
- Build it
    - Write the code
    - Test it
- Ship it`);
    expect(slides[0].bullets.map((bullet) => bullet.level)).toEqual([0, 1, 1, 0]);
  });

  it("splits a long list over several slides, keeping the title", () => {
    const many = Array.from({ length: 14 }, (_v, i) => `- point ${i + 1}`).join("\n");
    const slides = parseOutline(`Everything\n${many}`);
    expect(slides).toHaveLength(Math.ceil(14 / MAX_BULLETS_PER_SLIDE));
    expect(slides.every((slide) => slide.title === "Everything")).toBe(true);
    expect(slides[0].continued).toBe(false);
    expect(slides[1].continued).toBe(true);
  });

  it("makes a title out of the first point when the notes never gave one", () => {
    const slides = parseOutline(`- We should move the office to Leeds. It is cheaper.
- The lease ends in June`);
    expect(slides[0].title).toBe("We should move the office to Leeds");
  });

  it("keeps a heading with nothing under it as its own slide", () => {
    const slides = parseOutline(`Part two

Details
- something`);
    expect(slides.map((slide) => slide.title)).toEqual(["Part two", "Details"]);
    expect(slides[0].bullets).toHaveLength(0);
  });

  it("returns nothing for nothing", () => {
    expect(parseOutline("")).toEqual([]);
    expect(parseOutline("   \n\n  ")).toEqual([]);
  });

  it("handles Windows line endings", () => {
    expect(titles("Title\r\n- one\r\n- two")).toEqual(["Title"]);
  });
});

describe("tables in pasted text", () => {
  const pasted = `Revenue by region
Region\tQ1\tQ2
North\t1200\t1400
South\t900\t1100
East\t600\t700`;

  it("reads tab-separated lines as a table, not as bullets", () => {
    const slides = parseOutline(pasted);
    expect(slides[0].table).toEqual({
      headers: ["Region", "Q1", "Q2"],
      rows: [
        ["North", "1200", "1400"],
        ["South", "900", "1100"],
        ["East", "600", "700"],
      ],
    });
    expect(slides[0].bullets).toHaveLength(0);
  });

  it("reads a markdown pipe table and drops its dashes", () => {
    const slides = parseOutline(`Costs
| Item | Amount |
| --- | --- |
| Rent | 1200 |
| Power | 300 |`);
    expect(slides[0].table?.headers).toEqual(["Item", "Amount"]);
    expect(slides[0].table?.rows).toEqual([
      ["Rent", "1200"],
      ["Power", "300"],
    ]);
  });

  it("does not treat a sentence with commas as a table", () => {
    const slides = parseOutline(`Products
- Widgets, gadgets and sprockets
- Sold in Leeds, York and Hull`);
    expect(slides[0].table).toBeNull();
    expect(slides[0].bullets).toHaveLength(2);
  });

  it("keeps bullets and a table on the same block", () => {
    const slides = parseOutline(`Summary
- Revenue is up
Region\tTotal
North\t100
South\t200`);
    expect(slides[0].table).not.toBeNull();
    expect(slides[0].bullets.map((bullet) => bullet.text)).toEqual(["Revenue is up"]);
  });
});

describe("charting a pasted table", () => {
  const table = {
    headers: ["Region", "Q1", "Q2"],
    rows: [
      ["North", "1200", "1400"],
      ["South", "900", "1100"],
    ],
  };

  it("finds the label column and the numeric ones", () => {
    expect(chartableColumns(table)).toEqual({ labelColumn: 0, valueColumns: [1, 2] });
  });

  it("says no when there is nothing to chart", () => {
    expect(
      chartableColumns({
        headers: ["A", "B"],
        rows: [
          ["one", "two"],
          ["three", "four"],
        ],
      })
    ).toBeNull();
    expect(chartableColumns({ headers: ["A"], rows: [["1"]] })).toBeNull();
  });

  it("copes with money and percentages written as text", () => {
    const money = {
      headers: ["Item", "Cost"],
      rows: [
        ["Rent", "$1,200"],
        ["Power", "$300"],
      ],
    };
    expect(chartableColumns(money)).toEqual({ labelColumn: 0, valueColumns: [1] });
    expect(tableToGrid(money, { labelColumn: 0, valueColumns: [1] })).toEqual([
      ["Item", "Cost"],
      ["Rent", 1200],
      ["Power", 300],
    ]);
  });
});

describe("describeOutline", () => {
  it("says what it read", () => {
    const slides = parseOutline("Plan\n- one\n- two");
    expect(describeOutline(slides)).toBe("Read 1 slide, 2 points.");
  });

  it("is honest about reading nothing", () => {
    expect(describeOutline([])).toContain("nothing in that");
  });
});
