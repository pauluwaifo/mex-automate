import {
  aggregateTable,
  aggregationLabel,
  BLANK_CATEGORY,
  CHART_TYPES,
  ChartRequest,
  findChartType,
  OTHER_CATEGORY,
  validateChartRequest,
} from "../src/taskpane/features/charts";
import { Grid } from "../src/taskpane/shared/types";

const headers = ["Region", "Rep", "Amount", "Units"];

const rows: Grid = [
  ["North", "Ada", 100, 2],
  ["South", "Bo", 250, 5],
  ["North", "Cy", 50, 1],
  ["East", "Ada", 25, 1],
  ["South", "Bo", 250, 3],
];

const base = {
  groupByColumn: 0,
  valueColumns: [2],
  aggregation: "sum" as const,
  sort: "none" as const,
  topN: null,
};

describe("aggregateTable", () => {
  it("groups rows and sums a value column", () => {
    const result = aggregateTable(headers, rows, base);
    expect(result.headers).toEqual(["Region", "Sum of Amount"]);
    expect(result.rows).toEqual([
      ["North", 150],
      ["South", 500],
      ["East", 25],
    ]);
    expect(result.sourceRows).toBe(5);
    expect(result.categoryCount).toBe(3);
  });

  it("aggregates several value columns at once", () => {
    const result = aggregateTable(headers, rows, { ...base, valueColumns: [2, 3] });
    expect(result.headers).toEqual(["Region", "Sum of Amount", "Sum of Units"]);
    expect(result.rows).toEqual([
      ["North", 150, 3],
      ["South", 500, 8],
      ["East", 25, 1],
    ]);
  });

  it("averages only over the numeric cells", () => {
    const result = aggregateTable(headers, rows, { ...base, aggregation: "average" });
    expect(result.headers).toEqual(["Region", "Average of Amount"]);
    expect(result.rows[0]).toEqual(["North", 75]);
    expect(result.rows[1]).toEqual(["South", 250]);
  });

  it("counts non-blank cells and distinct values separately", () => {
    const counted = aggregateTable(headers, rows, {
      ...base,
      valueColumns: [1],
      aggregation: "count",
    });
    expect(counted.rows).toEqual([
      ["North", 2],
      ["South", 2],
      ["East", 1],
    ]);

    const distinct = aggregateTable(headers, rows, {
      ...base,
      valueColumns: [1],
      aggregation: "countDistinct",
    });
    // South is Bo twice, so it has one distinct rep.
    expect(distinct.rows).toEqual([
      ["North", 2],
      ["South", 1],
      ["East", 1],
    ]);
  });

  it("finds minimum and maximum", () => {
    expect(aggregateTable(headers, rows, { ...base, aggregation: "min" }).rows[0]).toEqual([
      "North",
      50,
    ]);
    expect(aggregateTable(headers, rows, { ...base, aggregation: "max" }).rows[0]).toEqual([
      "North",
      100,
    ]);
  });

  it("sorts by value and by category", () => {
    expect(
      aggregateTable(headers, rows, { ...base, sort: "valueDesc" }).rows.map((row) => row[0])
    ).toEqual(["South", "North", "East"]);
    expect(
      aggregateTable(headers, rows, { ...base, sort: "valueAsc" }).rows.map((row) => row[0])
    ).toEqual(["East", "North", "South"]);
    expect(
      aggregateTable(headers, rows, { ...base, sort: "categoryAsc" }).rows.map((row) => row[0])
    ).toEqual(["East", "North", "South"]);
    expect(
      aggregateTable(headers, rows, { ...base, sort: "categoryDesc" }).rows.map((row) => row[0])
    ).toEqual(["South", "North", "East"]);
  });

  it("rolls everything past topN into a single Other row", () => {
    const result = aggregateTable(headers, rows, { ...base, sort: "valueDesc", topN: 2 });
    expect(result.rows).toEqual([
      ["South", 500],
      ["North", 150],
      [OTHER_CATEGORY, 25],
    ]);
    // categoryCount reports the real number, before trimming.
    expect(result.categoryCount).toBe(3);
  });

  it("leaves the table alone when topN exceeds the category count", () => {
    const result = aggregateTable(headers, rows, { ...base, topN: 10 });
    expect(result.rows.map((row) => row[0])).not.toContain(OTHER_CATEGORY);
  });

  it("labels rows with an empty category", () => {
    const result = aggregateTable(
      headers,
      [
        ["", "Ada", 10, 1],
        [null, "Bo", 5, 1],
      ],
      base
    );
    expect(result.rows).toEqual([[BLANK_CATEGORY, 15]]);
  });

  it("reads numbers that arrived as text", () => {
    const result = aggregateTable(headers, [["North", "Ada", "1,234.50", 1]], base);
    expect(result.rows).toEqual([["North", 1234.5]]);
  });

  it("reads parenthesised negatives the way accounting exports write them", () => {
    const result = aggregateTable(headers, [["North", "Ada", "(50)", 1]], base);
    expect(result.rows).toEqual([["North", -50]]);
  });

  it("counts cells a numeric aggregation had to ignore", () => {
    const result = aggregateTable(
      headers,
      [
        ["North", "Ada", "n/a", 1],
        ["North", "Bo", 10, 1],
      ],
      base
    );
    expect(result.rows).toEqual([["North", 10]]);
    expect(result.ignoredNonNumeric).toBe(1);
  });

  it("skips blank rows entirely", () => {
    const result = aggregateTable(
      headers,
      [
        ["North", "Ada", 10, 1],
        [null, null, null, null],
      ],
      base
    );
    expect(result.sourceRows).toBe(1);
    expect(result.rows).toHaveLength(1);
  });

  it("returns a sum of zero, not null, for a group with no numbers", () => {
    const result = aggregateTable(headers, [["North", "Ada", "n/a", 1]], base);
    expect(result.rows).toEqual([["North", 0]]);
  });

  it("returns an empty table for no rows", () => {
    const result = aggregateTable(headers, [], base);
    expect(result.rows).toEqual([]);
    expect(result.sourceRows).toBe(0);
  });
});

describe("aggregationLabel", () => {
  it("reads like a pivot table field", () => {
    expect(aggregationLabel("sum", "Amount")).toBe("Sum of Amount");
    expect(aggregationLabel("countDistinct", "Rep")).toBe("Count unique of Rep");
  });
});

describe("chart types", () => {
  it("gives every type a unique id and an Excel type", () => {
    const ids = CHART_TYPES.map((type) => type.id);
    expect(new Set(ids).size).toBe(ids.length);
    for (const type of CHART_TYPES) {
      expect(type.excelType).not.toBe("");
      expect(type.label).not.toBe("");
    }
  });

  it("marks pie and doughnut as single-series and axis-free", () => {
    for (const id of ["pie", "doughnut"]) {
      const type = findChartType(id);
      expect(type?.singleSeriesOnly).toBe(true);
      expect(type?.hasAxes).toBe(false);
    }
    expect(findChartType("columnClustered")?.singleSeriesOnly).toBe(false);
    expect(findChartType("columnClustered")?.hasAxes).toBe(true);
  });
});

describe("validateChartRequest", () => {
  const request = (overrides: Partial<ChartRequest> = {}): ChartRequest => ({
    source: { kind: "selection", name: "" },
    summarize: {
      groupByColumn: 0,
      valueColumns: [2],
      aggregation: "sum",
      sort: "none",
      topN: null,
    },
    chartTypeId: "columnClustered",
    title: "",
    legendPosition: "right",
    showDataLabels: false,
    destination: "newSheet",
    destinationSheetName: "Chart",
    ...overrides,
  });

  it("accepts a well-formed request", () => {
    expect(validateChartRequest(request(), 4)).toEqual([]);
  });

  it("rejects an unknown chart type", () => {
    expect(validateChartRequest(request({ chartTypeId: "nope" }), 4)).toContain(
      'Unknown chart type "nope".'
    );
  });

  it("requires a value column", () => {
    const problems = validateChartRequest(
      request({
        summarize: {
          groupByColumn: 0,
          valueColumns: [],
          aggregation: "sum",
          sort: "none",
          topN: null,
        },
      }),
      4
    );
    expect(problems).toContain("Choose at least one column to summarize.");
  });

  it("rejects grouping by a column that is also a value column", () => {
    const problems = validateChartRequest(
      request({
        summarize: {
          groupByColumn: 2,
          valueColumns: [2],
          aggregation: "sum",
          sort: "none",
          topN: null,
        },
      }),
      4
    );
    expect(problems).toContain("The group-by column cannot also be a value column.");
  });

  it("rejects a multi-series pie", () => {
    const problems = validateChartRequest(
      request({
        chartTypeId: "pie",
        summarize: {
          groupByColumn: 0,
          valueColumns: [2, 3],
          aggregation: "sum",
          sort: "none",
          topN: null,
        },
      }),
      4
    );
    expect(problems.some((problem) => problem.includes("one series"))).toBe(true);
  });

  it("rejects a topN below one", () => {
    const problems = validateChartRequest(
      request({
        summarize: {
          groupByColumn: 0,
          valueColumns: [2],
          aggregation: "sum",
          sort: "none",
          topN: 0,
        },
      }),
      4
    );
    expect(problems).toContain("Top N must be at least 1.");
  });

  it("needs two columns when charting a range without summarizing", () => {
    expect(validateChartRequest(request({ summarize: null }), 1)).toContain(
      "Charting a range directly needs at least two columns: labels and values."
    );
    expect(validateChartRequest(request({ summarize: null }), 2)).toEqual([]);
  });
});
