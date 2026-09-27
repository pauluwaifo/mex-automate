import {
  categoryInsights,
  formatShare,
  formatValue,
  percentChange,
  pickInsights,
  trailingRun,
  trendInsights,
} from "../src/taskpane/features/insights";

describe("formatting", () => {
  it("writes numbers the way a person would in a sentence", () => {
    expect(formatValue(42)).toBe("42");
    expect(formatValue(1234.56)).toBe("1,235");
    expect(formatValue(128400)).toBe("128.4k");
    expect(formatValue(1_234_567)).toBe("1.2m");
    expect(formatValue(2_000_000_000)).toBe("2bn");
    expect(formatValue(12.5)).toBe("12.5");
    expect(formatValue(-1_500_000)).toBe("-1.5m");
  });

  it("gives a share the precision it deserves", () => {
    expect(formatShare(0.78)).toBe("78%");
    expect(formatShare(0.045)).toBe("4.5%");
    expect(formatShare(0.0031)).toBe("0.31%");
  });

  it("refuses to divide by nothing", () => {
    expect(percentChange(0, 100)).toBeNull();
    expect(percentChange(100, 120)).toBeCloseTo(0.2);
    expect(percentChange(-100, -50)).toBeCloseTo(0.5);
  });
});

describe("categoryInsights", () => {
  const points = [
    { label: "North", value: 500 },
    { label: "South", value: 300 },
    { label: "East", value: 150 },
    { label: "West", value: 50 },
  ];

  it("states the total and who leads", () => {
    const insights = categoryInsights("Revenue", "Region", points);
    const text = insights.map((insight) => insight.text);
    expect(text).toContain("Revenue totals 1,000 across 4 regions.");
    expect(text).toContain("North is the largest region at 500, 50% of the total.");
  });

  it("points out when a few carry the total", () => {
    const insights = categoryInsights("Revenue", "Region", points);
    expect(insights.map((insight) => insight.text)).toContain(
      "The top 2 of 4 make up 80% of revenue."
    );
  });

  it("mentions a wide spread between biggest and smallest", () => {
    const skewed = [
      { label: "North", value: 1000 },
      { label: "South", value: 50 },
      { label: "East", value: 20 },
    ];
    const insights = categoryInsights("Revenue", "Region", skewed);
    expect(insights.some((insight) => insight.kind === "spread")).toBe(true);
  });

  it("calls out categories with nothing recorded", () => {
    const withGap = [...points, { label: "Central", value: 0 }];
    const insights = categoryInsights("Revenue", "Region", withGap);
    expect(insights.map((insight) => insight.text)).toContain(
      "Central recorded no revenue at all."
    );
  });

  it("says nothing at all about an empty series", () => {
    expect(categoryInsights("Revenue", "Region", [])).toEqual([]);
  });

  it("pluralizes the dimension properly", () => {
    const branches = categoryInsights("Sales", "Branch", [
      { label: "A", value: 1 },
      { label: "B", value: 2 },
    ]);
    expect(branches[0].text).toContain("2 branches");
    const countries = categoryInsights("Sales", "Country", [
      { label: "A", value: 1 },
      { label: "B", value: 2 },
    ]);
    expect(countries[0].text).toContain("2 countries");
  });
});

describe("trendInsights", () => {
  const rising = [
    { label: "Jan", sortKey: 1, value: 100 },
    { label: "Feb", sortKey: 2, value: 110 },
    { label: "Mar", sortKey: 3, value: 125 },
    { label: "Apr", sortKey: 4, value: 140 },
    { label: "May", sortKey: 5, value: 160 },
  ];

  it("leads with the most recent move", () => {
    const insights = trendInsights("Revenue", rising);
    expect(insights[0].text).toBe("Revenue in May is 160, up 14% on Apr.");
  });

  it("reads the overall direction from halves, not the last wobble", () => {
    const wobbly = [...rising.slice(0, 4), { label: "May", sortKey: 5, value: 138 }];
    const trend = trendInsights("Revenue", wobbly).find((insight) => insight.kind === "trend");
    expect(trend?.text).toContain("up");
  });

  it("counts a run of rises", () => {
    const streak = trendInsights("Revenue", rising).find((insight) => insight.kind === "streak");
    expect(streak?.text).toBe("Revenue has risen 4 periods in a row.");
  });

  it("names the best and worst periods", () => {
    const best = trendInsights("Revenue", rising).find((insight) => insight.kind === "best");
    expect(best?.text).toBe("May was the strongest period at 160; Jan the weakest at 100.");
  });

  it("sorts the periods before reading them", () => {
    const shuffled = [rising[2], rising[0], rising[4], rising[1], rising[3]];
    expect(trendInsights("Revenue", shuffled)[0].text).toBe(
      "Revenue in May is 160, up 14% on Apr."
    );
  });

  it("needs two periods before it says anything", () => {
    expect(trendInsights("Revenue", [rising[0]])).toEqual([]);
  });

  it("copes with a period that had nothing to compare against", () => {
    const fromZero = [
      { label: "Jan", sortKey: 1, value: 0 },
      { label: "Feb", sortKey: 2, value: 50 },
    ];
    expect(trendInsights("Revenue", fromZero)[0].text).toBe(
      "Revenue in Feb is 50, against nothing in Jan."
    );
  });
});

describe("trailingRun", () => {
  it("counts consecutive rises and falls at the end of a series", () => {
    expect(trailingRun([1, 2, 3, 4])).toBe(3);
    expect(trailingRun([4, 3, 2, 1])).toBe(-3);
    expect(trailingRun([1, 5, 4, 6])).toBe(1);
    expect(trailingRun([2, 2, 2])).toBe(0);
    expect(trailingRun([7])).toBe(0);
  });
});

describe("pickInsights", () => {
  it("keeps the most important and never repeats a kind", () => {
    const picked = pickInsights(
      [
        { kind: "total", text: "a", importance: 0.2 },
        { kind: "latest", text: "b", importance: 0.9 },
        { kind: "latest", text: "c", importance: 0.8 },
        { kind: "trend", text: "d", importance: 0.5 },
      ],
      2
    );
    expect(picked.map((insight) => insight.text)).toEqual(["b", "d"]);
  });

  it("drops an insight that says exactly the same thing", () => {
    const picked = pickInsights([
      { kind: "total", text: "same", importance: 0.9 },
      { kind: "trend", text: "same", importance: 0.8 },
    ]);
    expect(picked).toHaveLength(1);
  });
});
