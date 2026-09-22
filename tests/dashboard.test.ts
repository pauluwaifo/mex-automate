import {
  bucketOf,
  buildChartData,
  chooseGrain,
  colorsFor,
  DashChart,
  layoutDashboard,
  LAYOUT,
  MAX_CHARTS,
  OTHER,
  OTHER_COLOR,
  PALETTE,
  planDashboard,
  rankDimensions,
  rankMeasures,
  suggestCharts,
  suggestKpis,
  topWithOther,
} from "../src/taskpane/features/dashboard";
import { dateToExcelSerial } from "../src/taskpane/features/dataCleaning";
import { profileColumn } from "../src/taskpane/features/tidy";
import { CellValue, Grid } from "../src/taskpane/shared/types";

const HEADERS = ["Order ID", "Order Date", "Region", "Channel", "Product", "Units", "Revenue"];
const REGIONS = ["Europe", "Africa", "Americas", "Asia"];
const CHANNELS = ["Online", "Retail", "Wholesale"];
const PRODUCTS = Array.from({ length: 15 }, (_v, i) => `Product ${String.fromCharCode(65 + i)}`);

/** 60 orders across Jan-Jun 2024, deterministic. */
const ROWS: Grid = Array.from({ length: 60 }, (_v, i) => {
  const date = dateToExcelSerial(new Date(Date.UTC(2024, i % 6, 1 + (i % 27))));
  const units = 1 + (i % 7);
  return [
    `A-${1000 + i}`,
    date,
    REGIONS[i % 4],
    CHANNELS[i % 3],
    PRODUCTS[(i * 7) % 15],
    units,
    units * (100 + (i % 5) * 25),
  ];
});

const PROFILES = HEADERS.map((header, c) =>
  profileColumn(
    header,
    ROWS.map((row) => row[c]),
    c
  )
);
const TOTAL_REVENUE = ROWS.reduce((sum, row) => sum + (row[6] as number), 0);

const sumColumn = (grid: Grid, column: number) =>
  grid.slice(1).reduce((sum, row) => sum + (row[column] as number), 0);

describe("column ranking", () => {
  it("puts revenue first among measures and never counts IDs", () => {
    const measures = rankMeasures(PROFILES).map((p) => p.header);
    expect(measures[0]).toBe("Revenue");
    expect(measures).toContain("Units");
    expect(measures).not.toContain("Order ID");
  });

  it("prefers a few distinct values for grouping", () => {
    const dims = rankDimensions(PROFILES).map((p) => p.header);
    expect(dims.slice(0, 2).sort()).toEqual(["Channel", "Region"]);
    expect(dims).not.toContain("Order ID");
  });
});

describe("suggestCharts", () => {
  const charts = suggestCharts(PROFILES, ROWS);
  const kinds = charts.map((chart) => chart.kind);

  it("suggests a full set of up to eight distinct charts", () => {
    expect(charts).toHaveLength(MAX_CHARTS);
    expect(new Set(charts.map((chart) => chart.id)).size).toBe(charts.length);
    expect(new Set(charts.map((chart) => chart.title)).size).toBe(charts.length);
  });

  it("leads with the trend over time", () => {
    expect(charts[0].kind).toBe("line");
    expect(charts[0].dimension).toBe("Order Date");
    expect(charts[0].timeGrain).toBe("month");
    expect(charts[0].title).toBe("Revenue over time");
  });

  it("covers comparison, share, breakdown, top-10 and relationship", () => {
    expect(kinds).toContain("doughnut");
    expect(kinds).toContain("stackedColumn");
    expect(kinds).toContain("scatter");
    expect(charts.some((chart) => chart.title === "Top 10 Product by Revenue")).toBe(true);
    expect(charts.some((chart) => chart.kind === "line" && chart.series !== null)).toBe(true);
  });

  it("respects a smaller limit", () => {
    expect(suggestCharts(PROFILES, ROWS, 3)).toHaveLength(3);
  });

  it("falls back to counting rows when there is nothing to add up", () => {
    const headers = ["Region", "Channel"];
    const rows: Grid = ROWS.map((row) => [row[2], row[3]]);
    const profiles = headers.map((h, c) =>
      profileColumn(
        h,
        rows.map((row) => row[c]),
        c
      )
    );
    const counted = suggestCharts(profiles, rows);
    expect(counted.length).toBeGreaterThan(0);
    expect(counted.every((chart) => chart.aggregation === "count")).toBe(true);
    expect(counted[0].title).toMatch(/^Count by /);
  });
});

describe("suggestKpis", () => {
  const kpis = suggestKpis(PROFILES, ROWS);

  it("gives total, count, average and the top group", () => {
    expect(kpis.map((kpi) => kpi.id)).toEqual(["total", "rows", "average", "top"]);
    expect(kpis[0].value).toBe(TOTAL_REVENUE);
    expect(kpis[1].value).toBe(60);
    expect(kpis[2].value).toBeCloseTo(TOTAL_REVENUE / 60);
    expect(kpis[3].text).not.toBeNull();
  });
});

describe("time buckets", () => {
  it("chooses a readable grain from the span", () => {
    expect(chooseGrain(45000, 45030)).toBe("day");
    expect(chooseGrain(45000, 45300)).toBe("month");
    expect(chooseGrain(45000, 47000)).toBe("quarter");
    expect(chooseGrain(40000, 46000)).toBe("year");
  });

  it("labels buckets the way people read them", () => {
    const serial = dateToExcelSerial(new Date(Date.UTC(2024, 3, 15)));
    expect(bucketOf(serial, "month").label).toBe("Apr 2024");
    expect(bucketOf(serial, "quarter").label).toBe("Q2 2024");
    expect(bucketOf(serial, "year").label).toBe("2024");
    expect(bucketOf(serial, "day").label).toBe("15 Apr");
  });

  it("gives keys that sort chronologically", () => {
    const dec = dateToExcelSerial(new Date(Date.UTC(2023, 11, 1)));
    const jan = dateToExcelSerial(new Date(Date.UTC(2024, 0, 1)));
    expect(bucketOf(dec, "month").key).toBeLessThan(bucketOf(jan, "month").key);
  });
});

describe("topWithOther", () => {
  it("keeps the biggest and folds the rest into Other", () => {
    const entries: Array<[string, number]> = [
      ["a", 1],
      ["b", 5],
      ["c", 3],
      ["d", 2],
    ];
    expect(topWithOther(entries, 3)).toEqual([
      ["b", 5],
      ["c", 3],
      [OTHER, 3],
    ]);
    expect(topWithOther(entries, null)).toEqual([
      ["b", 5],
      ["c", 3],
      ["d", 2],
      ["a", 1],
    ]);
    expect(topWithOther(entries, 10)).toHaveLength(4);
  });
});

describe("buildChartData", () => {
  const chart = (overrides: Partial<DashChart>): DashChart => ({
    id: "x",
    kind: "column",
    title: "t",
    reason: "",
    dimension: "Region",
    series: null,
    measure: "Revenue",
    measure2: null,
    aggregation: "sum",
    timeGrain: null,
    topN: null,
    ...overrides,
  });

  it("totals a measure by category, biggest first", () => {
    const data = buildChartData(chart({}), PROFILES, ROWS)!;
    expect(data[0]).toEqual(["Region", "Revenue"]);
    expect(data).toHaveLength(1 + REGIONS.length);
    expect(sumColumn(data, 1)).toBe(TOTAL_REVENUE);
    const values = data.slice(1).map((row) => row[1] as number);
    expect(values).toEqual(values.slice().sort((a, b) => b - a));
  });

  it("folds small groups into Other without losing any total", () => {
    const data = buildChartData(chart({ dimension: "Product", topN: 10 }), PROFILES, ROWS)!;
    expect(data).toHaveLength(11);
    expect(data[10][0]).toBe(OTHER);
    expect(sumColumn(data, 1)).toBe(TOTAL_REVENUE);
  });

  it("buckets dates by month in order", () => {
    const data = buildChartData(
      chart({ kind: "line", dimension: "Order Date", timeGrain: "month" }),
      PROFILES,
      ROWS
    )!;
    expect(data[0]).toEqual(["Month", "Revenue"]);
    expect(data.slice(1).map((row) => row[0])).toEqual([
      "Jan 2024",
      "Feb 2024",
      "Mar 2024",
      "Apr 2024",
      "May 2024",
      "Jun 2024",
    ]);
    expect(sumColumn(data, 1)).toBe(TOTAL_REVENUE);
  });

  it("counts rows when asked", () => {
    const data = buildChartData(chart({ measure: null, aggregation: "count" }), PROFILES, ROWS)!;
    expect(data[0][1]).toBe("Count");
    expect(sumColumn(data, 1)).toBe(60);
  });

  it("splits by a second column for stacked charts", () => {
    const data = buildChartData(
      chart({ kind: "stackedColumn", series: "Channel" }),
      PROFILES,
      ROWS
    )!;
    expect(data[0][0]).toBe("Region");
    expect(data[0].slice(1).sort()).toEqual(CHANNELS.slice().sort());
    const grand = data
      .slice(1)
      .reduce((sum, row) => sum + row.slice(1).reduce<number>((s, v) => s + (v as number), 0), 0);
    expect(grand).toBe(TOTAL_REVENUE);
  });

  it("draws one line per group over time", () => {
    const data = buildChartData(
      chart({ kind: "line", dimension: "Order Date", series: "Channel", timeGrain: "month" }),
      PROFILES,
      ROWS
    )!;
    expect(data).toHaveLength(7);
    expect(data[0]).toHaveLength(1 + CHANNELS.length);
  });

  it("pairs two measures for a scatter, with the second along the bottom", () => {
    const data = buildChartData(
      chart({ kind: "scatter", dimension: null, measure: "Revenue", measure2: "Units" }),
      PROFILES,
      ROWS
    )!;
    expect(data[0]).toEqual(["Units", "Revenue"]);
    expect(data).toHaveLength(61);
    expect(data[1]).toEqual([ROWS[0][5], ROWS[0][6]]);
  });

  it("returns null when a column has gone", () => {
    expect(buildChartData(chart({ dimension: "Nope" }), PROFILES, ROWS)).toBeNull();
    expect(buildChartData(chart({ measure: "Nope" }), PROFILES, ROWS)).toBeNull();
  });

  it("reads numbers and dates still stored as text", () => {
    const rows: Grid = [
      ["Europe", "$1,000.00", "2024-01-05"],
      ["Europe", "(250)", "05/02/2024"],
      ["Africa", "1.500,00 €", "2024-01-20"],
    ];
    const headers = ["Region", "Revenue", "Order Date"];
    const profiles = headers.map((h, c) =>
      profileColumn(
        h,
        rows.map((row: CellValue[]) => row[c]),
        c
      )
    );
    const byRegion = buildChartData(chart({}), profiles, rows)!;
    expect(byRegion.slice(1)).toEqual([
      ["Africa", 1500],
      ["Europe", 750],
    ]);
  });
});

describe("planDashboard", () => {
  it("bundles suggestions and headline numbers", () => {
    const plan = planDashboard(HEADERS, PROFILES, ROWS);
    expect(plan.charts).toHaveLength(8);
    expect(plan.kpis.length).toBeGreaterThan(0);
    expect(plan.rowCount).toBe(60);
  });
});

describe("layoutDashboard", () => {
  const overlaps = (
    a: { row: number; column: number; rows: number; columns: number },
    b: typeof a
  ) =>
    a.row < b.row + b.rows &&
    b.row < a.row + a.rows &&
    a.column < b.column + b.columns &&
    b.column < a.column + a.columns;

  it("runs the hero chart full width and pairs the rest", () => {
    const layout = layoutDashboard(8, 4, true);
    expect(layout.charts[0].columns).toBe(LAYOUT.columns);
    expect(layout.charts.slice(1, 7).every((place) => place.columns === LAYOUT.columns / 2)).toBe(
      true
    );
    // 7 after the hero: three pairs and a full-width last one.
    expect(layout.charts[7].columns).toBe(LAYOUT.columns);
  });

  it("never overlaps and stays inside the grid", () => {
    for (const [charts, kpis, hero] of [
      [8, 4, true],
      [7, 3, false],
      [1, 0, false],
      [0, 4, false],
    ] as const) {
      const layout = layoutDashboard(charts, kpis, hero);
      const all = [...layout.kpis, ...layout.charts];
      all.forEach((a, i) => {
        expect(a.column + a.columns).toBeLessThanOrEqual(LAYOUT.columns);
        all.slice(i + 1).forEach((b) => expect(overlaps(a, b)).toBe(false));
      });
      expect(layout.charts).toHaveLength(charts);
      expect(layout.kpis).toHaveLength(kpis);
    }
  });

  it("puts the headline numbers above the charts", () => {
    const layout = layoutDashboard(2, 4, false);
    expect(layout.kpis.every((tile) => tile.row < layout.charts[0].row)).toBe(true);
    expect(layout.kpis.map((tile) => tile.columns)).toEqual([3, 3, 3, 3]);
  });
});

describe("colorsFor", () => {
  it("assigns palette slots in order and keeps Other grey", () => {
    const colors = colorsFor(["Europe", "Africa", OTHER, "Asia"]);
    expect(colors.get("Europe")).toBe(PALETTE[0]);
    expect(colors.get("Africa")).toBe(PALETTE[1]);
    expect(colors.get("Asia")).toBe(PALETTE[2]);
    expect(colors.get(OTHER)).toBe(OTHER_COLOR);
  });
});

describe("regressions found on a real-sized export", () => {
  const headers = ["Region", "Product", "Units", "Unit Price", "Revenue"];
  const rows: Grid = ROWS.map((row) => [row[2], row[4], row[5], 12.5, row[6]]);
  const profiles = headers.map((h, c) =>
    profileColumn(
      h,
      rows.map((row) => row[c]),
      c
    )
  );

  it("never adds up prices, rates or percentages", () => {
    const measures = rankMeasures(profiles).map((p) => p.header);
    expect(measures[0]).toBe("Revenue");
    expect(measures).not.toContain("Unit Price");
  });

  it("gives the ranked list to a many-valued column, not a three-value one", () => {
    const titles = suggestCharts(profiles, rows).map((chart) => chart.title);
    expect(titles).toContain("Top 10 Product by Revenue");
    expect(titles.some((title) => /^Top 10 Region/.test(title))).toBe(false);
  });
});
