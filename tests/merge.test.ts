import {
  buildMergedTable,
  compareCells,
  compareCellsDirected,
  DEFAULT_MERGE_OPTIONS,
  describePlan,
  parseAliasLines,
  planHeaders,
  SourceTable,
  SOURCE_COLUMN_HEADER,
} from "../src/taskpane/features/merge";

const january: SourceTable = {
  label: "January",
  headers: ["Order ID", "Customer", "Amount"],
  rows: [
    [1, "Acme", 100],
    [2, "Beta", 250],
  ],
};

const february: SourceTable = {
  label: "February",
  // Same columns, different spelling and order - the classic real-world case.
  headers: ["order_id", "AMOUNT", "Customer", "Region"],
  rows: [[3, 400, "Gamma", "North"]],
};

const mergeDefaults = { addSourceColumn: false, skipBlankRows: true };

describe("planHeaders", () => {
  it("unifies headers that differ only by case, spacing or separators", () => {
    const plan = planHeaders([january, february]);
    expect(plan.headers).toEqual(["Order ID", "Customer", "Amount", "Region"]);
  });

  it("maps each source column to its output position", () => {
    const plan = planHeaders([january, february]);
    expect(plan.columnTargets[0]).toEqual([0, 1, 2]);
    // February is order_id, AMOUNT, Customer, Region.
    expect(plan.columnTargets[1]).toEqual([0, 2, 1, 3]);
  });

  it("flags columns that only some sources have", () => {
    const plan = planHeaders([january, february]);
    expect(plan.partialHeaders).toEqual(["Region"]);
    expect(plan.contributors["Order ID"]).toEqual(["January", "February"]);
    expect(plan.contributors["Region"]).toEqual(["February"]);
  });

  it("merges genuinely different names via an alias", () => {
    const withClient: SourceTable = {
      label: "March",
      headers: ["Order ID", "Client Name"],
      rows: [[4, "Delta"]],
    };
    const plan = planHeaders([january, withClient], { "client name": "Customer" });
    expect(plan.headers).toEqual(["Order ID", "Customer", "Amount"]);
    expect(plan.columnTargets[1]).toEqual([0, 1]);
    expect(plan.partialHeaders).toEqual(["Amount"]);
  });

  it("drops unnamed columns rather than colliding them", () => {
    const plan = planHeaders([{ label: "X", headers: ["A", "", "B"], rows: [] }]);
    expect(plan.headers).toEqual(["A", "B"]);
    expect(plan.columnTargets[0]).toEqual([0, -1, 1]);
  });

  it("handles a single source", () => {
    const plan = planHeaders([january]);
    expect(plan.headers).toEqual(january.headers);
    expect(plan.partialHeaders).toEqual([]);
  });
});

describe("buildMergedTable", () => {
  it("stacks rows onto the unified column layout", () => {
    const plan = planHeaders([january, february]);
    const merged = buildMergedTable([january, february], plan, mergeDefaults);

    expect(merged.grid).toEqual([
      ["Order ID", "Customer", "Amount", "Region"],
      [1, "Acme", 100, null],
      [2, "Beta", 250, null],
      [3, "Gamma", 400, "North"],
    ]);
    expect(merged.totalRows).toBe(3);
    expect(merged.rowsPerSource).toEqual({ January: 2, February: 1 });
  });

  it("adds a source column when asked", () => {
    const plan = planHeaders([january, february]);
    const merged = buildMergedTable([january, february], plan, {
      ...mergeDefaults,
      addSourceColumn: true,
    });

    expect(merged.grid[0][0]).toBe(SOURCE_COLUMN_HEADER);
    expect(merged.grid[1][0]).toBe("January");
    expect(merged.grid[3][0]).toBe("February");
    expect(merged.grid[3]).toEqual(["February", 3, "Gamma", 400, "North"]);
  });

  it("skips blank rows only when asked", () => {
    const withBlank: SourceTable = {
      label: "Blanks",
      headers: ["A"],
      rows: [["x"], [null], [""]],
    };
    const plan = planHeaders([withBlank]);

    expect(buildMergedTable([withBlank], plan, mergeDefaults).totalRows).toBe(1);
    expect(
      buildMergedTable([withBlank], plan, { ...mergeDefaults, skipBlankRows: false }).totalRows
    ).toBe(3);
  });

  it("carries each source's number formats through to the right output column", () => {
    const dated: SourceTable = {
      label: "Dated",
      headers: ["Amount", "Order Date"],
      rows: [[100, 45360]],
      numberFormats: [["#,##0.00", "yyyy-mm-dd"]],
    };
    const reversed: SourceTable = {
      label: "Reversed",
      headers: ["Order Date", "Amount"],
      rows: [[45361, 200]],
      numberFormats: [["dd/mm/yyyy", "0.00"]],
    };

    const plan = planHeaders([dated, reversed]);
    const merged = buildMergedTable([dated, reversed], plan, mergeDefaults);

    // Output columns are [Amount, Order Date]; the second source is reordered.
    expect(merged.numberFormats[1]).toEqual(["#,##0.00", "yyyy-mm-dd"]);
    expect(merged.numberFormats[2]).toEqual(["0.00", "dd/mm/yyyy"]);
  });

  it("defaults missing number formats to General", () => {
    const plan = planHeaders([january]);
    const merged = buildMergedTable([january], plan, mergeDefaults);
    expect(merged.numberFormats[1]).toEqual(["General", "General", "General"]);
  });

  it("pads a source row that is shorter than its header list", () => {
    const short: SourceTable = { label: "Short", headers: ["A", "B"], rows: [["x"]] };
    const plan = planHeaders([short]);
    const merged = buildMergedTable([short], plan, mergeDefaults);
    expect(merged.grid[1]).toEqual(["x", null]);
  });
});

describe("sorting the merged table", () => {
  const plan = planHeaders([january, february]);

  it("leaves rows in source order by default", () => {
    const merged = buildMergedTable([january, february], plan, mergeDefaults);
    expect(merged.grid.slice(1).map((row) => row[2])).toEqual([100, 250, 400]);
  });

  it("sorts by a chosen column, keeping the header first", () => {
    const merged = buildMergedTable([january, february], plan, {
      ...mergeDefaults,
      sortByHeader: "Amount",
      sortDescending: true,
    });
    expect(merged.grid[0]).toEqual(["Order ID", "Customer", "Amount", "Region"]);
    expect(merged.grid.slice(1).map((row) => row[2])).toEqual([400, 250, 100]);
  });

  it("sorts text naturally and case-insensitively", () => {
    const merged = buildMergedTable([january, february], plan, {
      ...mergeDefaults,
      sortByHeader: "Customer",
    });
    expect(merged.grid.slice(1).map((row) => row[1])).toEqual(["Acme", "Beta", "Gamma"]);
  });

  it("keeps each row's number formats attached to it when sorting", () => {
    const first: SourceTable = {
      label: "First",
      headers: ["Amount"],
      rows: [[300], [100]],
      numberFormats: [["#,##0.00"], ["0%"]],
    };
    const merged = buildMergedTable([first], planHeaders([first]), {
      ...mergeDefaults,
      sortByHeader: "Amount",
    });
    expect(merged.grid.slice(1)).toEqual([[100], [300]]);
    // The "0%" format belonged to the 100 row and must have travelled with it.
    expect(merged.numberFormats.slice(1)).toEqual([["0%"], ["#,##0.00"]]);
  });

  it("pushes blank cells to the end in both directions", () => {
    const sparse: SourceTable = {
      label: "Sparse",
      headers: ["Amount"],
      rows: [[5], [null], [1]],
    };
    const ascending = buildMergedTable([sparse], planHeaders([sparse]), {
      ...mergeDefaults,
      skipBlankRows: false,
      sortByHeader: "Amount",
    });
    expect(ascending.grid.slice(1)).toEqual([[1], [5], [null]]);

    const descending = buildMergedTable([sparse], planHeaders([sparse]), {
      ...mergeDefaults,
      skipBlankRows: false,
      sortByHeader: "Amount",
      sortDescending: true,
    });
    expect(descending.grid.slice(1)).toEqual([[5], [1], [null]]);
  });

  it("ignores a sort column that is not in the merged headers", () => {
    const merged = buildMergedTable([january, february], plan, {
      ...mergeDefaults,
      sortByHeader: "Nope",
    });
    expect(merged.grid.slice(1).map((row) => row[2])).toEqual([100, 250, 400]);
  });

  it("can sort by the Source column", () => {
    const merged = buildMergedTable([january, february], plan, {
      ...mergeDefaults,
      addSourceColumn: true,
      sortByHeader: SOURCE_COLUMN_HEADER,
    });
    expect(merged.grid.slice(1).map((row) => row[0])).toEqual(["February", "January", "January"]);
  });
});

describe("compareCells", () => {
  it("orders numbers numerically and text alphabetically", () => {
    expect(compareCells(1, 2)).toBeLessThan(0);
    expect(compareCells(10, 9)).toBeGreaterThan(0);
    expect(compareCells("apple", "banana")).toBeLessThan(0);
  });

  it("puts blanks last whichever way round they are compared", () => {
    expect(compareCells(null, 1)).toBeGreaterThan(0);
    expect(compareCells(1, null)).toBeLessThan(0);
    expect(compareCells("", 1)).toBeGreaterThan(0);
    expect(compareCells(null, "")).toBe(0);
  });

  it("compares text case-insensitively", () => {
    expect(compareCells("apple", "APPLE")).toBe(0);
  });
});

describe("compareCellsDirected", () => {
  it("reverses the order when descending", () => {
    expect(compareCellsDirected(1, 2, false)).toBeLessThan(0);
    expect(compareCellsDirected(1, 2, true)).toBeGreaterThan(0);
  });

  it("keeps blanks last in both directions", () => {
    expect(compareCellsDirected(null, 1, false)).toBeGreaterThan(0);
    expect(compareCellsDirected(null, 1, true)).toBeGreaterThan(0);
    expect(compareCellsDirected(1, null, true)).toBeLessThan(0);
  });
});

describe("describePlan", () => {
  it("reports row counts per source and the partial columns", () => {
    const plan = planHeaders([january, february]);
    const merged = buildMergedTable([january, february], plan, mergeDefaults);
    const details = describePlan(plan, merged);

    expect(details).toContain("January: 2 rows");
    expect(details).toContain("February: 1 row");
    expect(details.some((line) => line.includes("Region"))).toBe(true);
  });
});

describe("parseAliasLines", () => {
  it("reads one rule per line and normalizes the left side", () => {
    expect(parseAliasLines("Client Name = Customer\nOrder No=Order ID")).toEqual({
      "client name": "Customer",
      "order no": "Order ID",
    });
  });

  it("ignores blank lines, comments and half-written rules", () => {
    expect(parseAliasLines("\n  \nno equals sign\nClient = \n= Customer\nA = B")).toEqual({
      a: "B",
    });
  });

  it("keeps only the first = so the target can contain one", () => {
    expect(parseAliasLines("A = B = C")).toEqual({ a: "B = C" });
  });

  it("handles CRLF line endings", () => {
    expect(parseAliasLines("A = B\r\nC = D")).toEqual({ a: "B", c: "D" });
  });

  it("returns an empty map for empty input", () => {
    expect(parseAliasLines("")).toEqual({});
  });

  it("produces aliases planHeaders can actually use", () => {
    const aliases = parseAliasLines("Client Name = Customer");
    const withClient: SourceTable = {
      label: "March",
      headers: ["Order ID", "Client Name"],
      rows: [[4, "Delta"]],
    };
    expect(planHeaders([january, withClient], aliases).headers).toEqual([
      "Order ID",
      "Customer",
      "Amount",
    ]);
  });
});

describe("DEFAULT_MERGE_OPTIONS", () => {
  it("is a usable starting configuration", () => {
    expect(DEFAULT_MERGE_OPTIONS.destinationSheetName).toBe("Merged");
    expect(DEFAULT_MERGE_OPTIONS.headerAliases).toEqual({});
  });
});
