import {
  groupAddress,
  groupIssues,
  Issue,
  IssueKind,
  median,
  outlierIndexes,
  reviewTable,
  summarize,
  widenSumRange,
} from "../src/taskpane/features/review";
import { describeSelection, issuesInSelection } from "../src/taskpane/features/reviewSheet";
import { dateToExcelSerial } from "../src/taskpane/features/dataCleaning";
import { Grid } from "../src/taskpane/shared/types";

const TODAY = dateToExcelSerial(new Date(Date.UTC(2024, 5, 30)));
const serial = (y: number, m: number, d: number) => dateToExcelSerial(new Date(Date.UTC(y, m, d)));

/** A tidy-looking sheet that quietly contains six real mistakes. */
const SHEET: Grid = [
  ["Order ID", "Order Date", "Customer", "Units", "Unit Price", "Revenue"],
  ["SO-1001", serial(2024, 0, 3), "Acme Ltd", 12, 12.5, 150],
  ["SO-1002", serial(2024, 0, 9), "Beta Trading", 8, 12.5, 100],
  ["SO-1003", serial(2024, 1, 2), "acme ltd", 20, 12.5, 250],
  ["SO-1004", serial(2024, 1, 14), "Gamma Co", 35, 12.5, 437.5],
  ["SO-1005", serial(2024, 2, 1), "Beta Trading", 7, 12.5, 87.5],
  ["SO-1003", serial(2024, 2, 8), "Delta Supplies", 5, 12.5, 62.5],
  ["SO-1007", serial(2024, 2, 19), "Gamma Co", 9, 12.5, 112.5],
  ["SO-1008", serial(2024, 3, 2), "Acme Ltd", 15, 12.5, 187.5],
  ["SO-1009", serial(2024, 3, 11), "Delta Supplies", 11, 12.5, 137.5],
  ["SO-1010", serial(2024, 3, 22), "Acme Ltd", 6, 12.5, 75],
  ["SO-1011", serial(2024, 4, 6), "Beta Trading", 14, 12.5, 175],
  ["SO-1012", serial(2024, 4, 17), "Gamma Co", "1,250.00", 12.5, 250],
  ["SO-1013", serial(2024, 4, 28), "Delta Supplies", 10, 12.5, 125],
  ["SO-1014", serial(2029, 5, 4), "Acme Ltd", 13, 12.5, 162.5],
  ["SO-1015", serial(2024, 5, 12), "  Gamma Co ", 9, 12.5, 112.5],
  [null, null, null, null, null, null],
  ["Total", null, null, null, null, 2225],
];

/** Column F calculates Units * Unit Price everywhere except row 8, which is typed in. */
const FORMULAS: string[][] = SHEET.map((row, r) => {
  if (r === 0 || r >= 16) return row.map(() => "");
  const cells = row.map(() => "");
  cells[5] = r === 8 ? "" : `=D${r + 1}*E${r + 1}`;
  return cells;
});
FORMULAS[17] = ["", "", "", "", "", "=SUM(F2:F14)"];

const review = reviewTable({ grid: SHEET, formulas: FORMULAS }, { today: TODAY });
const of = (kind: IssueKind): Issue[] => review.issues.filter((issue) => issue.kind === kind);

describe("reviewTable on a sheet with real mistakes", () => {
  it("counts the data rows and finds the table", () => {
    expect(review.rows).toBe(15);
    expect(review.headers).toEqual([
      "Order ID",
      "Order Date",
      "Customer",
      "Units",
      "Unit Price",
      "Revenue",
    ]);
  });

  it("spots the typed-in number in a column of formulas", () => {
    const issues = of("inconsistentFormula");
    expect(issues).toHaveLength(1);
    expect(issues[0].address).toBe("F9");
    expect(issues[0].severity).toBe("error");
    expect(issues[0].fix).toEqual({ type: "setFormula", formula: "=D9*E9" });
    expect(issues[0].detail).toContain("typed in");
  });

  it("spots the total that misses rows", () => {
    const issues = of("shortSumRange");
    expect(issues).toHaveLength(1);
    expect(issues[0].address).toBe("F18");
    expect(issues[0].fix).toEqual({ type: "setFormula", formula: "=SUM(F2:F16)" });
  });

  it("spots the amount typed as text", () => {
    const issues = of("textNumber");
    expect(issues).toHaveLength(1);
    expect(issues[0].address).toBe("D13");
    expect(issues[0].fix).toMatchObject({ type: "setValue", value: 1250 });
  });

  it("spots the duplicate ID, naming the row it clashes with", () => {
    const issues = of("duplicateId");
    expect(issues).toHaveLength(1);
    expect(issues[0].address).toBe("A7");
    expect(issues[0].detail).toContain("row 4");
    expect(issues[0].fix).toBeNull();
  });

  it("spots the name spelled two ways and offers the common spelling", () => {
    const issues = of("nearDuplicate");
    expect(issues.map((issue) => issue.address)).toEqual(["C4"]);
    expect(issues[0].fix).toEqual({ type: "setValue", value: "Acme Ltd" });
  });

  it("spots stray spaces", () => {
    const issues = of("straySpaces");
    expect(issues).toHaveLength(1);
    expect(issues[0].address).toBe("C16");
    expect(issues[0].fix).toEqual({ type: "setValue", value: "Gamma Co" });
  });

  it("spots the date years in the future", () => {
    const issues = of("futureDate");
    expect(issues).toHaveLength(1);
    expect(issues[0].address).toBe("B15");
    expect(issues[0].fix).toBeNull();
  });

  it("puts errors first", () => {
    expect(review.issues[0].severity).toBe("error");
    const severities = review.issues.map((issue) => issue.severity);
    expect(severities.indexOf("error")).toBeLessThan(severities.indexOf("warning"));
  });

  it("gives every issue a stable id and a plain-English detail", () => {
    const ids = review.issues.map((issue) => issue.id);
    expect(new Set(ids).size).toBe(ids.length);
    review.issues.forEach((issue) => {
      expect(issue.detail.length).toBeGreaterThan(20);
      expect(issue.title).not.toBe("");
    });
    // Scanning twice gives the same ids, so "ignore" sticks.
    const again = reviewTable({ grid: SHEET, formulas: FORMULAS }, { today: TODAY });
    expect(again.issues.map((issue) => issue.id)).toEqual(ids);
  });
});

describe("reviewTable on clean data", () => {
  it("finds nothing to flag", () => {
    const clean: Grid = [
      ["Region", "Amount"],
      ["North", 100],
      ["South", 200],
      ["East", 150],
    ];
    expect(reviewTable({ grid: clean }, { today: TODAY }).issues).toEqual([]);
  });

  it("ignores rows that aren't data, and says how many there were", () => {
    const withTotals: Grid = [
      ["Region", "Amount"],
      ["North", 100],
      ["Total", 100],
    ];
    const result = reviewTable({ grid: withTotals }, { today: TODAY });
    expect(result.rows).toBe(1);
    expect(result.structuralRows).toBe(1);
  });

  it("returns nothing for an empty sheet", () => {
    expect(reviewTable({ grid: [] }).issues).toEqual([]);
    expect(reviewTable({ grid: [] }).location).toBeNull();
  });
});

describe("addresses away from A1", () => {
  it("reports sheet addresses when the table starts further in", () => {
    const grid: Grid = [
      ["Region", "Amount"],
      ["North", "1,200"],
    ];
    const result = reviewTable({ grid, origin: { row: 4, column: 2 } }, { today: TODAY });
    // Table starts at C5, so the amount is in D6.
    expect(result.issues[0].address).toBe("D6");
  });
});

describe("widenSumRange", () => {
  const hasNumber = (rows: number[]) => (excelRow: number) => rows.includes(excelRow);

  it("extends a total that stops short", () => {
    expect(
      widenSumRange("=SUM(C4:C9)", hasNumber([4, 5, 6, 7, 8, 9, 10, 11]), { first: 4, last: 11 })
    ).toBe("=SUM(C4:C11)");
  });

  it("extends upwards too", () => {
    expect(widenSumRange("=SUM(C6:C9)", hasNumber([4, 5, 6, 7, 8, 9]), { first: 4, last: 9 })).toBe(
      "=SUM(C4:C9)"
    );
  });

  it("leaves a correct total alone", () => {
    expect(
      widenSumRange("=SUM(C4:C9)", hasNumber([4, 5, 6, 7, 8, 9]), { first: 4, last: 9 })
    ).toBeNull();
  });

  it("respects the edge of the table", () => {
    expect(
      widenSumRange("=SUM(C4:C9)", hasNumber([3, 4, 5, 6, 7, 8, 9, 10]), { first: 4, last: 9 })
    ).toBeNull();
  });

  it("ignores formulas it doesn't understand", () => {
    expect(widenSumRange("=SUM(C4:D9)", hasNumber([4]), { first: 1, last: 20 })).toBeNull();
    expect(widenSumRange("=AVERAGE(C4:C9)", hasNumber([4]), { first: 1, last: 20 })).toBeNull();
  });
});

describe("statistics", () => {
  it("takes a median of odd and even counts", () => {
    expect(median([3, 1, 2])).toBe(2);
    expect(median([4, 1, 2, 3])).toBe(2.5);
    expect(median([])).toBe(0);
  });

  it("finds a value far outside the rest", () => {
    const values = [10, 11, 9, 10, 12, 11, 10, 9, 11, 10, 12, 10, 9500];
    expect(outlierIndexes(values)).toEqual([12]);
  });

  it("stays quiet on small or evenly spread columns", () => {
    expect(outlierIndexes([1, 2, 3, 900])).toEqual([]);
    expect(outlierIndexes(Array.from({ length: 30 }, (_v, i) => i * 10))).toEqual([]);
  });
});

describe("grouping and summary", () => {
  it("groups by problem and column, worst first", () => {
    const groups = groupIssues(review.issues);
    expect(groups[0].severity).toBe("error");
    groups.forEach((group) => {
      expect(group.issues.length).toBeGreaterThan(0);
      expect(group.fixable).toBeLessThanOrEqual(group.issues.length);
    });
    const formulaGroup = groups.find((group) => group.kind === "inconsistentFormula");
    expect(formulaGroup?.fixable).toBe(1);
  });

  it("writes the one-line verdict", () => {
    expect(summarize(review.issues)).toMatch(/errors/);
    expect(summarize([])).toBe("nothing to flag");
  });

  it("gives a group one address covering its rows", () => {
    const group = groupIssues(review.issues).find((item) => item.kind === "duplicateId")!;
    expect(groupAddress(group, "Orders")).toBe("Orders!A7");
  });
});

describe("issuesInSelection", () => {
  const issue = (row: number, column: number): Issue =>
    ({ row, column, id: `x${row}${column}` }) as Issue;
  const issues = [issue(5, 2), issue(9, 2), issue(5, 4)];

  it("takes the whole column when a single cell is selected", () => {
    const picked = issuesInSelection(issues, {
      sheet: "S",
      address: "C6",
      row: 5,
      column: 2,
      rowCount: 1,
      columnCount: 1,
    });
    expect(picked).toHaveLength(2);
  });

  it("takes just the block when a range is selected", () => {
    const picked = issuesInSelection(issues, {
      sheet: "S",
      address: "C6:E7",
      row: 5,
      column: 2,
      rowCount: 2,
      columnCount: 3,
    });
    expect(picked.map((item) => item.column)).toEqual([2, 4]);
  });

  it("describes a selection with its sheet", () => {
    expect(
      describeSelection({
        sheet: "Jan Orders",
        address: "C6:C9",
        row: 5,
        column: 2,
        rowCount: 4,
        columnCount: 1,
      })
    ).toBe("'Jan Orders'!C6:C9");
  });
});
