import {
  buildAddress,
  chunk,
  columnIndexFromLetter,
  columnLetter,
  isBlankRow,
  normalizeHeader,
  padRow,
  parseAddress,
  plural,
  quoteSheetName,
  rectangularize,
  toHeaderTable,
  trimTrailingBlanks,
  uniqueName,
} from "../src/taskpane/shared/excelHelpers";

describe("column letters", () => {
  it.each([
    [0, "A"],
    [25, "Z"],
    [26, "AA"],
    [51, "AZ"],
    [52, "BA"],
    [701, "ZZ"],
    [702, "AAA"],
    [16383, "XFD"],
  ])("index %i is column %s", (index, letter) => {
    expect(columnLetter(index)).toBe(letter);
    expect(columnIndexFromLetter(letter)).toBe(index);
  });

  it("rejects a negative index", () => {
    expect(() => columnLetter(-1)).toThrow(RangeError);
  });

  it("rejects text that is not a column letter", () => {
    expect(() => columnIndexFromLetter("A1")).toThrow(RangeError);
  });
});

describe("normalizeHeader", () => {
  it("folds spacing, separators and case together", () => {
    const variants = ["Order Date", "order_date", "  ORDER   DATE ", "Order-Date", "order.date"];
    const keys = new Set(variants.map(normalizeHeader));
    expect(keys.size).toBe(1);
    expect([...keys][0]).toBe("order date");
  });

  it("treats null and undefined as empty", () => {
    expect(normalizeHeader(null)).toBe("");
    expect(normalizeHeader(undefined)).toBe("");
  });
});

describe("parseAddress", () => {
  it("parses a plain range", () => {
    expect(parseAddress("B2:D20")).toEqual({
      sheetName: null,
      rowIndex: 1,
      columnIndex: 1,
      rowCount: 19,
      columnCount: 3,
    });
  });

  it("parses a sheet-qualified, dollar-anchored range", () => {
    expect(parseAddress("Sheet1!$B$2:$D$20")).toMatchObject({
      sheetName: "Sheet1",
      rowIndex: 1,
      columnIndex: 1,
      rowCount: 19,
      columnCount: 3,
    });
  });

  it("parses a quoted sheet name containing spaces and apostrophes", () => {
    expect(parseAddress("'Paul''s Data'!A1")).toMatchObject({
      sheetName: "Paul's Data",
      rowIndex: 0,
      columnIndex: 0,
      rowCount: 1,
      columnCount: 1,
    });
  });

  it("normalises a range written bottom-right first", () => {
    expect(parseAddress("D20:B2")).toMatchObject({
      rowIndex: 1,
      columnIndex: 1,
      rowCount: 19,
      columnCount: 3,
    });
  });

  it("throws on junk", () => {
    expect(() => parseAddress("not an address")).toThrow(/Could not parse/);
  });
});

describe("buildAddress", () => {
  it("builds a multi-cell address", () => {
    expect(buildAddress("Data", 0, 0, 3, 2)).toBe("Data!A1:B3");
  });

  it("collapses a single cell", () => {
    expect(buildAddress(null, 4, 2, 1, 1)).toBe("C5");
  });

  it("quotes a sheet name that needs it", () => {
    expect(buildAddress("My Data", 0, 0, 2, 2)).toBe("'My Data'!A1:B2");
  });

  it("round-trips through parseAddress", () => {
    const address = buildAddress("Sheet1", 3, 4, 10, 6);
    expect(parseAddress(address)).toMatchObject({
      sheetName: "Sheet1",
      rowIndex: 3,
      columnIndex: 4,
      rowCount: 10,
      columnCount: 6,
    });
  });
});

describe("quoteSheetName", () => {
  it("leaves a simple name alone", () => {
    expect(quoteSheetName("Sheet1")).toBe("Sheet1");
  });

  it("doubles embedded apostrophes", () => {
    expect(quoteSheetName("Paul's Data")).toBe("'Paul''s Data'");
  });
});

describe("grid shaping", () => {
  it("drops trailing blank rows and columns", () => {
    const grid = [
      ["a", "b", null, null],
      ["c", "d", null, null],
      [null, null, null, null],
      ["", "  ", null, null],
    ];
    expect(trimTrailingBlanks(grid)).toEqual([
      ["a", "b"],
      ["c", "d"],
    ]);
  });

  it("returns an empty grid when everything is blank", () => {
    expect(
      trimTrailingBlanks([
        [null, ""],
        ["  ", null],
      ])
    ).toEqual([]);
  });

  it("keeps interior blanks", () => {
    const grid = [
      ["a", null, "c"],
      [null, null, null],
      ["d", null, "f"],
    ];
    expect(trimTrailingBlanks(grid)).toEqual(grid);
  });

  it("pads ragged rows to a rectangle", () => {
    expect(rectangularize([["a"], ["b", "c", "d"]])).toEqual([
      ["a", null, null],
      ["b", "c", "d"],
    ]);
  });

  it("pads and truncates a single row to a width", () => {
    expect(padRow(["a", "b", "c"], 2)).toEqual(["a", "b"]);
    expect(padRow(["a"], 3)).toEqual(["a", null, null]);
  });

  it("names unnamed header columns after their column letter", () => {
    expect(
      toHeaderTable([
        ["Name", "", "Total"],
        ["x", 1, 2],
      ])
    ).toEqual({
      headers: ["Name", "Column B", "Total"],
      rows: [["x", 1, 2]],
    });
  });

  it("recognises blank rows including whitespace-only cells", () => {
    expect(isBlankRow([null, "", "   "])).toBe(true);
    expect(isBlankRow([null, "", "x"])).toBe(false);
    expect(isBlankRow([0])).toBe(false);
  });
});

describe("uniqueName", () => {
  it("passes an unused name through", () => {
    expect(uniqueName("Merged", ["Sheet1"])).toBe("Merged");
  });

  it("suffixes until it finds a free name, ignoring case", () => {
    expect(uniqueName("Merged", ["merged", "Merged (2)"])).toBe("Merged (3)");
  });
});

describe("chunk", () => {
  it("splits into fixed-size batches with a short tail", () => {
    expect(chunk([1, 2, 3, 4, 5], 2)).toEqual([[1, 2], [3, 4], [5]]);
  });

  it("rejects a non-positive size", () => {
    expect(() => chunk([1], 0)).toThrow(RangeError);
  });
});

describe("plural", () => {
  it("switches on the count", () => {
    expect(plural(1, "row")).toBe("1 row");
    expect(plural(2, "row")).toBe("2 rows");
  });

  it("accepts an irregular plural", () => {
    expect(plural(2, "entry", "entries")).toBe("2 entries");
  });
});
