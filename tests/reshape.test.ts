import {
  describeCrosstab,
  detectCrosstab,
  detectCrosstabInGrid,
  fillForward,
  mergeHeaderRows,
  parsePeriodHeader,
  unpivot,
} from "../src/taskpane/features/reshape";
import { excelSerialToDate } from "../src/taskpane/features/dataCleaning";
import { Grid, HeaderTable } from "../src/taskpane/shared/types";

/** The shape almost every budget sheet arrives in. */
const MONTHLY: HeaderTable = {
  headers: ["Region", "Product", "Jan 2024", "Feb 2024", "Mar 2024"],
  rows: [
    ["North", "Widget", 100, 120, 140],
    ["North", "Gadget", 80, null, 95],
    ["South", "Widget", 60, 70, 75],
  ],
};

describe("parsePeriodHeader", () => {
  it("reads month names with and without a year", () => {
    expect(parsePeriodHeader("Jan 2024")).toMatchObject({ label: "Jan 2024", grain: "month" });
    expect(parsePeriodHeader("January 2024")).toMatchObject({ label: "Jan 2024" });
    expect(parsePeriodHeader("Mar")).toMatchObject({ label: "Mar", yearAssumed: true, date: null });
  });

  it("reads the separators exports actually use", () => {
    for (const header of ["Jan-24", "Jan_24", "Jan/24", "Jan 24", "JAN24"]) {
      expect(parsePeriodHeader(header)).toMatchObject({ label: "Jan 2024" });
    }
  });

  it("reads year-first and quarter headings", () => {
    expect(parsePeriodHeader("2024-01")).toMatchObject({ label: "Jan 2024", grain: "month" });
    expect(parsePeriodHeader("Q1 2024")).toMatchObject({ label: "Q1 2024", grain: "quarter" });
    expect(parsePeriodHeader("2024 Q3")).toMatchObject({ label: "Q3 2024", grain: "quarter" });
    expect(parsePeriodHeader("2023")).toMatchObject({ label: "2023", grain: "year" });
  });

  it("splits a measure off the period, whichever side it is on", () => {
    expect(parsePeriodHeader("Jan Units")).toMatchObject({ label: "Jan", measure: "Units" });
    expect(parsePeriodHeader("Revenue Jan 2024")).toMatchObject({
      label: "Jan 2024",
      measure: "Revenue",
    });
  });

  it("ignores filler words that are not measures", () => {
    expect(parsePeriodHeader("Jan Total")).toMatchObject({ label: "Jan", measure: null });
    expect(parsePeriodHeader("Jan 2024 Actual")).toMatchObject({ measure: null });
  });

  it("reads two-digit years the way Excel does", () => {
    expect(parsePeriodHeader("Jan 98")?.date?.getUTCFullYear()).toBe(1998);
    expect(parsePeriodHeader("Jan 05")?.date?.getUTCFullYear()).toBe(2005);
  });

  it("says no to headings that are not periods", () => {
    for (const header of [
      "Region",
      "Revenue",
      "Order ID",
      "",
      "Total",
      "Q5",
      "1899",
      "May Contain Nuts",
    ]) {
      expect(parsePeriodHeader(header)).toBeNull();
    }
  });

  it("sorts periods chronologically even out of order", () => {
    const dec = parsePeriodHeader("Dec 2023");
    const jan = parsePeriodHeader("Jan 2024");
    expect(dec && jan && dec.sortKey < jan.sortKey).toBe(true);
  });
});

describe("detectCrosstab", () => {
  it("finds the wide block and the columns that key it", () => {
    const shape = detectCrosstab(MONTHLY);
    expect(shape).not.toBeNull();
    expect(shape?.idColumns).toEqual([0, 1]);
    expect(shape?.valueColumns).toEqual([2, 3, 4]);
    expect(shape?.periodColumnName).toBe("Month");
    expect(shape?.datesKnown).toBe(true);
  });

  it("leaves an ordinary table alone", () => {
    const tidy: HeaderTable = {
      headers: ["Order ID", "Date", "Region", "Units", "Revenue"],
      rows: [["SO-1", 45000, "North", 5, 120]],
    };
    expect(detectCrosstab(tidy)).toBeNull();
  });

  it("needs at least two period columns", () => {
    const one: HeaderTable = { headers: ["Region", "Jan 2024"], rows: [["North", 10]] };
    expect(detectCrosstab(one)).toBeNull();
  });

  it("needs a column left over to identify the row", () => {
    const allPeriods: HeaderTable = {
      headers: ["Jan 2024", "Feb 2024", "Mar 2024"],
      rows: [[1, 2, 3]],
    };
    expect(detectCrosstab(allPeriods)).toBeNull();
  });

  it("refuses a wide block whose cells are text, not numbers", () => {
    const notNumbers: HeaderTable = {
      headers: ["Region", "Jan 2024", "Feb 2024"],
      rows: [
        ["North", "on track", "late"],
        ["South", "late", "on track"],
      ],
    };
    expect(detectCrosstab(notNumbers)).toBeNull();
  });

  it("picks up the measures when headings name them", () => {
    const shape = detectCrosstab({
      headers: ["Product", "Jan Units", "Jan Revenue", "Feb Units", "Feb Revenue"],
      rows: [["Widget", 10, 100, 12, 120]],
    });
    expect(shape?.measures).toEqual(["Units", "Revenue"]);
    expect(shape?.idColumns).toEqual([0]);
  });

  it("refuses headings that mix named measures with bare periods", () => {
    const mixed = detectCrosstab({
      headers: ["Product", "Jan Units", "Feb", "Mar Units"],
      rows: [["Widget", 10, 11, 12]],
    });
    expect(mixed).toBeNull();
  });
});

describe("merged and two-row headings", () => {
  it("forward-fills the blanks a merged row reads back as", () => {
    expect(fillForward(["2023", null, null, "2024", null])).toEqual([
      "2023",
      "2023",
      "2023",
      "2024",
      "2024",
    ]);
  });

  it("joins a year row above a month row", () => {
    expect(mergeHeaderRows(["", "2024", null, null], ["Region", "Jan", "Feb", "Mar"])).toEqual([
      "Region",
      "Jan 2024",
      "Feb 2024",
      "Mar 2024",
    ]);
  });

  it("finds a crosstab that only appears once the two heading rows are joined", () => {
    const grid: Grid = [
      [null, "2024", null, null],
      ["Region", "Jan", "Feb", "Mar"],
      ["North", 10, 20, 30],
      ["South", 40, 50, 60],
    ];
    const found = detectCrosstabInGrid(grid);
    expect(found?.headerRows).toBe(2);
    expect(found?.headers).toEqual(["Region", "Jan 2024", "Feb 2024", "Mar 2024"]);
    expect(found?.rows).toHaveLength(2);
  });

  it("prefers a single heading row when that already works", () => {
    const grid: Grid = [
      ["Region", "Jan 2024", "Feb 2024"],
      ["North", 10, 20],
    ];
    expect(detectCrosstabInGrid(grid)?.headerRows).toBe(1);
  });
});

describe("unpivot", () => {
  it("writes one row per cell of the wide block", () => {
    const shape = detectCrosstab(MONTHLY)!;
    const result = unpivot(MONTHLY, shape);
    expect(result.table.headers).toEqual(["Region", "Product", "Month", "Value"]);
    // 3 rows x 3 months, less the one empty cell.
    expect(result.rowsOut).toBe(8);
    expect(result.skippedBlank).toBe(1);
    expect(result.table.rows[0].slice(0, 2)).toEqual(["North", "Widget"]);
    expect(result.table.rows[0][3]).toBe(100);
  });

  it("writes the period as a real date, formatted for reading", () => {
    const shape = detectCrosstab(MONTHLY)!;
    const result = unpivot(MONTHLY, shape);
    const serial = result.table.rows[0][2] as number;
    expect(typeof serial).toBe("number");
    const date = excelSerialToDate(serial);
    expect([date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()]).toEqual([2024, 0, 1]);
    expect(result.formats[2]).toBe("mmm yyyy");
  });

  it("keeps text labels when the headings never said which year", () => {
    const noYear: HeaderTable = {
      headers: ["Region", "Jan", "Feb"],
      rows: [["North", 1, 2]],
    };
    const shape = detectCrosstab(noYear)!;
    const result = unpivot(noYear, shape);
    // One id column, so the period lands in column 1 and the value in column 2.
    expect(result.table.headers).toEqual(["Region", "Month", "Value"]);
    expect(result.table.rows[0][1]).toBe("Jan");
    expect(result.formats[1]).toBeUndefined();
  });

  it("keeps empty cells when asked to", () => {
    const shape = detectCrosstab(MONTHLY)!;
    expect(unpivot(MONTHLY, shape, { keepBlanks: true }).rowsOut).toBe(9);
  });

  it("keeps a zero, which is a measurement and not a gap", () => {
    const withZero: HeaderTable = {
      headers: ["Region", "Jan 2024", "Feb 2024"],
      rows: [["North", 0, 5]],
    };
    const shape = detectCrosstab(withZero)!;
    const result = unpivot(withZero, shape);
    expect(result.rowsOut).toBe(2);
    expect(result.table.rows[0][2]).toBe(0);
  });

  it("gives each measure its own column", () => {
    const table: HeaderTable = {
      headers: ["Product", "Jan Units", "Jan Revenue", "Feb Units", "Feb Revenue"],
      rows: [
        ["Widget", 10, 100, 12, 120],
        ["Gadget", 5, 50, 6, 60],
      ],
    };
    const shape = detectCrosstab(table)!;
    const result = unpivot(table, shape);
    expect(result.table.headers).toEqual(["Product", "Month", "Units", "Revenue"]);
    expect(result.table.rows).toEqual([
      ["Widget", "Jan", 10, 100],
      ["Widget", "Feb", 12, 120],
      ["Gadget", "Jan", 5, 50],
      ["Gadget", "Feb", 6, 60],
    ]);
  });

  it("can sort the output chronologically", () => {
    const outOfOrder: HeaderTable = {
      headers: ["Region", "Mar 2024", "Jan 2024", "Feb 2024"],
      rows: [["North", 3, 1, 2]],
    };
    const shape = detectCrosstab(outOfOrder)!;
    const sorted = unpivot(outOfOrder, shape, { sortByPeriod: true, periodAsDate: false });
    expect(sorted.table.rows.map((row) => row[1])).toEqual(["Jan 2024", "Feb 2024", "Mar 2024"]);
    const asIs = unpivot(outOfOrder, shape, { periodAsDate: false });
    expect(asIs.table.rows.map((row) => row[1])).toEqual(["Mar 2024", "Jan 2024", "Feb 2024"]);
  });

  it("names the value column when asked", () => {
    const shape = detectCrosstab(MONTHLY)!;
    expect(unpivot(MONTHLY, shape, { valueColumnName: "Revenue" }).table.headers[3]).toBe(
      "Revenue"
    );
  });
});

describe("describeCrosstab", () => {
  it("says what it found in one line", () => {
    const shape = detectCrosstab(MONTHLY)!;
    expect(describeCrosstab(shape)).toBe(
      "3 month columns (Jan 2024 to Mar 2024), keyed by 2 columns"
    );
  });
});
