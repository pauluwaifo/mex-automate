import {
  cellsMatch,
  chooseKeyColumn,
  compareTables,
  diffReportGrid,
  largestMovements,
  movementByColumn,
  summarizeDiff,
} from "../src/taskpane/features/compare";
import { HeaderTable } from "../src/taskpane/shared/types";

const APRIL: HeaderTable = {
  headers: ["Order ID", "Customer", "Units", "Revenue"],
  rows: [
    ["SO-1001", "Acme Ltd", 10, 1000],
    ["SO-1002", "Beta Trading", 5, 500],
    ["SO-1003", "Gamma Co", 8, 800],
    ["SO-1004", "Delta Supplies", 2, 200],
  ],
};

/** The same file a month later: one price fix, one new order, one cancelled. */
const MAY: HeaderTable = {
  headers: ["Order ID", "Customer", "Units", "Revenue"],
  rows: [
    ["SO-1001", "Acme Ltd", 10, 1000],
    ["SO-1002", "Beta Trading", 6, 620],
    ["SO-1003", "Gamma Co", 8, 800],
    ["SO-1005", "Epsilon Ltd", 3, 300],
  ],
};

describe("chooseKeyColumn", () => {
  it("picks the column that identifies a row", () => {
    expect(chooseKeyColumn(APRIL, MAY)).toBe("Order ID");
  });

  it("prefers a column that sounds like an identifier over one that merely is unique", () => {
    const first: HeaderTable = {
      headers: ["Customer", "Invoice No"],
      rows: [
        ["Acme", "INV-1"],
        ["Beta", "INV-2"],
      ],
    };
    const second: HeaderTable = {
      headers: ["Customer", "Invoice No"],
      rows: [
        ["Acme", "INV-1"],
        ["Beta", "INV-2"],
      ],
    };
    expect(chooseKeyColumn(first, second)).toBe("Invoice No");
  });

  it("returns nothing when no shared column is unique enough", () => {
    const first: HeaderTable = {
      headers: ["Region", "Units"],
      rows: [
        ["North", 1],
        ["North", 2],
      ],
    };
    expect(chooseKeyColumn(first, first)).toBeNull();
  });

  it("will not use a column with gaps in it", () => {
    const gappy: HeaderTable = {
      headers: ["Ref", "Units"],
      rows: [
        ["A", 1],
        [null, 2],
        ["C", 3],
      ],
    };
    expect(chooseKeyColumn(gappy, gappy)).toBeNull();
  });
});

describe("cellsMatch", () => {
  it("compares numbers numerically, however they were typed", () => {
    expect(cellsMatch(1000, "1,000.00")).toBe(true);
    expect(cellsMatch(1000, 1000.004)).toBe(true);
    expect(cellsMatch(1000, 1000.02)).toBe(false);
  });

  it("uses the tolerance it is given", () => {
    expect(cellsMatch(100, 100.4, 0.5)).toBe(true);
    expect(cellsMatch(100, 100.4, 0.1)).toBe(false);
  });

  it("ignores case and surrounding spaces in text by default", () => {
    expect(cellsMatch("Acme Ltd", " acme ltd ")).toBe(true);
    expect(cellsMatch("Acme Ltd", "acme ltd", 0, false)).toBe(false);
  });

  it("treats two empty cells as equal and an empty one as a change", () => {
    expect(cellsMatch(null, "")).toBe(true);
    expect(cellsMatch(null, 0)).toBe(false);
  });

  it("does not match a number against text", () => {
    expect(cellsMatch(12, "twelve")).toBe(false);
  });
});

describe("compareTables", () => {
  const diff = compareTables(APRIL, MAY);

  it("matches on the key it found", () => {
    expect(diff.keyColumn).toBe("Order ID");
    expect(diff.comparedColumns).toEqual(["Customer", "Units", "Revenue"]);
  });

  it("reports the cells that changed, with both values", () => {
    expect(diff.changed).toHaveLength(2);
    expect(diff.changed).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ key: "SO-1002", header: "Units", before: 5, after: 6, delta: 1 }),
        expect.objectContaining({ key: "SO-1002", header: "Revenue", delta: 120 }),
      ])
    );
  });

  it("separates rows that were added from rows that went away", () => {
    expect(diff.added.map((row) => row.key)).toEqual(["SO-1005"]);
    expect(diff.removed.map((row) => row.key)).toEqual(["SO-1004"]);
  });

  it("counts the rows that did not move", () => {
    expect(diff.matchedRows).toBe(3);
    expect(diff.identicalRows).toBe(2);
  });

  it("says nothing changed when nothing did", () => {
    const same = compareTables(APRIL, APRIL);
    expect(same.changed).toHaveLength(0);
    expect(same.added).toHaveLength(0);
    expect(same.removed).toHaveLength(0);
    expect(summarizeDiff(same)).toBe("These two are identical.");
  });

  it("ignores rounding noise", () => {
    const rounded: HeaderTable = {
      headers: APRIL.headers,
      rows: APRIL.rows.map((row) => [...row.slice(0, 3), (row[3] as number) + 0.004]),
    };
    expect(compareTables(APRIL, rounded).changed).toHaveLength(0);
  });

  it("notes the columns that only one side has", () => {
    const withExtra: HeaderTable = {
      headers: [...MAY.headers, "Channel"],
      rows: MAY.rows.map((row) => [...row, "Web"]),
    };
    const result = compareTables(APRIL, withExtra);
    expect(result.columnsOnlyInSecond).toEqual(["Channel"]);
    expect(result.columnsOnlyInFirst).toEqual([]);
    expect(result.comparedColumns).not.toContain("Channel");
  });

  it("can be told which columns to look at", () => {
    const result = compareTables(APRIL, MAY, { key: "Order ID", columns: ["Revenue"] });
    expect(result.comparedColumns).toEqual(["Revenue"]);
    expect(result.changed.every((change) => change.header === "Revenue")).toBe(true);
  });

  it("reports repeated keys instead of quietly comparing the wrong rows", () => {
    const repeated: HeaderTable = {
      headers: APRIL.headers,
      rows: [...APRIL.rows, ["SO-1001", "Acme Ltd", 99, 9900]],
    };
    const result = compareTables(repeated, APRIL, { key: "Order ID" });
    expect(result.duplicateKeysFirst).toEqual(["so-1001"]);
    // The first row of each key is the one compared, so the 99 is not reported.
    expect(result.changed).toHaveLength(0);
  });

  it("falls back to whole-row matching when there is no usable key", () => {
    const first: HeaderTable = {
      headers: ["Region", "Units"],
      rows: [
        ["North", 1],
        ["South", 2],
      ],
    };
    const second: HeaderTable = {
      headers: ["Region", "Units"],
      rows: [
        ["North", 1],
        ["East", 3],
      ],
    };
    const result = compareTables(first, second);
    expect(result.keyColumn).toBeNull();
    expect(result.added.map((row) => row.values)).toEqual([["East", 3]]);
    expect(result.removed.map((row) => row.values)).toEqual([["South", 2]]);
    expect(result.changed).toHaveLength(0);
  });
});

describe("movement", () => {
  const diff = compareTables(APRIL, MAY);

  it("totals the movement per column, biggest first", () => {
    expect(movementByColumn(diff)).toEqual([
      { header: "Revenue", total: 120, cells: 1 },
      { header: "Units", total: 1, cells: 1 },
    ]);
  });

  it("ranks the individual movements", () => {
    expect(largestMovements(diff, 1)).toEqual([
      expect.objectContaining({ header: "Revenue", delta: 120 }),
    ]);
  });

  it("summarizes in one sentence", () => {
    expect(summarizeDiff(diff)).toBe("Found 2 changed cells across 1 row, 1 new row, 1 row gone.");
  });
});

describe("diffReportGrid", () => {
  const diff = compareTables(APRIL, MAY);
  const { grid, headingRows, tableHeaderRows } = diffReportGrid(diff, {
    first: "April",
    second: "May",
  });

  it("is rectangular, as Range.values requires", () => {
    const widths = new Set(grid.map((row) => row.length));
    expect(widths.size).toBe(1);
  });

  it("opens with the summary", () => {
    expect(grid[0][0]).toBe("Comparison");
    expect(grid.slice(1, 9).map((row) => row[0])).toEqual([
      "Older sheet",
      "Newer sheet",
      "Matched on",
      "Rows matched",
      "Rows identical",
      "Cells changed",
      "Rows only in the newer sheet",
      "Rows only in the older sheet",
    ]);
  });

  it("marks its headings so the writer can style them", () => {
    for (const row of headingRows) expect(typeof grid[row][0]).toBe("string");
    expect(headingRows.length).toBeGreaterThanOrEqual(4);
    expect(tableHeaderRows.length).toBeGreaterThanOrEqual(3);
  });

  it("lists every changed cell with both values", () => {
    const flat = grid.map((row) => row.join("|"));
    expect(flat.some((row) => row.startsWith("SO-1002|Revenue|500|620|120"))).toBe(true);
  });

  it("mentions repeated keys where they matter", () => {
    const repeated: HeaderTable = {
      headers: APRIL.headers,
      rows: [...APRIL.rows, ["SO-1001", "Acme Ltd", 99, 9900]],
    };
    const report = diffReportGrid(compareTables(repeated, MAY, { key: "Order ID" }), {
      first: "April",
      second: "May",
    });
    const text = report.grid.map((row) => row.join(" ")).join("\n");
    expect(text).toContain("Repeated keys in April");
  });
});
