import {
  applyTextCase,
  dateToExcelSerial,
  dedupeGrid,
  excelSerialToDate,
  groupContiguous,
  mapGrid,
  parseFlexibleDate,
  toProperCase,
  toSentenceCase,
  trimCell,
} from "../src/taskpane/features/dataCleaning";
import { Grid } from "../src/taskpane/shared/types";

const dedupeDefaults = {
  hasHeaderRow: true,
  keyColumns: [],
  ignoreCaseAndSpacing: true,
  deleteEntireRows: false,
};

describe("trimCell", () => {
  it("trims the ends", () => {
    expect(trimCell("  hello  ", { collapseInnerSpaces: false })).toBe("hello");
  });

  it("collapses inner runs only when asked", () => {
    expect(trimCell("a   b", { collapseInnerSpaces: false })).toBe("a   b");
    expect(trimCell("a   b", { collapseInnerSpaces: true })).toBe("a b");
  });

  it("removes the invisible whitespace that survives a web paste", () => {
    const nonBreaking = "\u00A0Acme\u200B Ltd\uFEFF";
    expect(trimCell(nonBreaking, { collapseInnerSpaces: true })).toBe("Acme Ltd");
  });

  it("leaves non-text values untouched", () => {
    expect(trimCell(42, { collapseInnerSpaces: true })).toBe(42);
    expect(trimCell(null, { collapseInnerSpaces: true })).toBeNull();
    expect(trimCell(true, { collapseInnerSpaces: true })).toBe(true);
  });
});

describe("text case", () => {
  it("proper-cases across separators", () => {
    expect(toProperCase("o'brien-smith ltd")).toBe("O'Brien-Smith Ltd");
    expect(toProperCase("ACME TRADING")).toBe("Acme Trading");
  });

  it("sentence-cases each sentence", () => {
    expect(toSentenceCase("first thing. second THING! third?")).toBe(
      "First thing. Second thing! Third?"
    );
  });

  it("applies the selected mode and skips blanks and non-text", () => {
    expect(applyTextCase("aBc", "upper")).toBe("ABC");
    expect(applyTextCase("aBc", "lower")).toBe("abc");
    expect(applyTextCase("   ", "upper")).toBe("   ");
    expect(applyTextCase(5, "upper")).toBe(5);
    expect(applyTextCase(null, "proper")).toBeNull();
  });
});

describe("Excel date serials", () => {
  it("matches the values Excel itself uses", () => {
    // 1900-01-01 is serial 1; 2024-03-09 is 45360. Both verified against Excel.
    expect(dateToExcelSerial(new Date(Date.UTC(1900, 0, 1)))).toBe(1);
    expect(dateToExcelSerial(new Date(Date.UTC(2024, 2, 9)))).toBe(45360);
  });

  it("straddles Excel's phantom 1900-02-29 correctly", () => {
    // Excel: 1900-02-28 is 59, the non-existent 1900-02-29 is 60, 1900-03-01 is 61.
    expect(dateToExcelSerial(new Date(Date.UTC(1900, 1, 28)))).toBe(59);
    expect(dateToExcelSerial(new Date(Date.UTC(1900, 2, 1)))).toBe(61);
    expect(excelSerialToDate(59).toISOString().slice(0, 10)).toBe("1900-02-28");
    expect(excelSerialToDate(61).toISOString().slice(0, 10)).toBe("1900-03-01");
    // Serial 60 is not a real date; it resolves forward rather than throwing.
    expect(excelSerialToDate(60).toISOString().slice(0, 10)).toBe("1900-03-01");
  });

  it("round-trips on both sides of the 1900 boundary", () => {
    for (const serial of [1, 59, 61, 1000, 45360, 2958465]) {
      expect(dateToExcelSerial(excelSerialToDate(serial))).toBe(serial);
    }
  });
});

describe("parseFlexibleDate", () => {
  const dayFirst = { dayFirst: true };
  const monthFirst = { dayFirst: false };

  const iso = (date: Date | null) => date?.toISOString().slice(0, 10);

  it("reads ISO regardless of the dayFirst setting", () => {
    expect(iso(parseFlexibleDate("2024-03-09", dayFirst))).toBe("2024-03-09");
    expect(iso(parseFlexibleDate("2024-03-09", monthFirst))).toBe("2024-03-09");
    expect(iso(parseFlexibleDate("2024-03-09T14:30:00", monthFirst))).toBe("2024-03-09");
  });

  it("uses dayFirst only when the value is genuinely ambiguous", () => {
    expect(iso(parseFlexibleDate("03/09/2024", dayFirst))).toBe("2024-09-03");
    expect(iso(parseFlexibleDate("03/09/2024", monthFirst))).toBe("2024-03-09");
  });

  it("disambiguates itself when one part cannot be a month", () => {
    expect(iso(parseFlexibleDate("25/03/2024", monthFirst))).toBe("2024-03-25");
    expect(iso(parseFlexibleDate("03/25/2024", dayFirst))).toBe("2024-03-25");
  });

  it("reads named months in either order", () => {
    expect(iso(parseFlexibleDate("9 Mar 2024", monthFirst))).toBe("2024-03-09");
    expect(iso(parseFlexibleDate("Mar 9, 2024", monthFirst))).toBe("2024-03-09");
    expect(iso(parseFlexibleDate("March 9th, 2024", monthFirst))).toBe("2024-03-09");
    expect(iso(parseFlexibleDate("9-March-24", monthFirst))).toBe("2024-03-09");
  });

  it("expands two-digit years the way Excel does", () => {
    expect(iso(parseFlexibleDate("01/01/29", monthFirst))).toBe("2029-01-01");
    expect(iso(parseFlexibleDate("01/01/30", monthFirst))).toBe("1930-01-01");
  });

  it("reads a number as an Excel serial", () => {
    expect(iso(parseFlexibleDate(45360, monthFirst))).toBe("2024-03-09");
  });

  it("rejects impossible and non-date values", () => {
    expect(parseFlexibleDate("31/02/2024", dayFirst)).toBeNull();
    expect(parseFlexibleDate("Acme Ltd", monthFirst)).toBeNull();
    expect(parseFlexibleDate("", monthFirst)).toBeNull();
    expect(parseFlexibleDate(null, monthFirst)).toBeNull();
    expect(parseFlexibleDate(true, monthFirst)).toBeNull();
    expect(parseFlexibleDate(0, monthFirst)).toBeNull();
  });
});

describe("dedupeGrid", () => {
  const grid: Grid = [
    ["Name", "City"],
    ["Acme", "Leeds"],
    ["acme ", "LEEDS"],
    ["Beta", "York"],
    ["Acme", "Leeds"],
  ];

  it("keeps the header and the first of each duplicate", () => {
    const result = dedupeGrid(grid, dedupeDefaults);
    expect(result.removedCount).toBe(2);
    expect(result.removedRowOffsets).toEqual([2, 4]);
    expect(result.rows).toEqual([
      ["Name", "City"],
      ["Acme", "Leeds"],
      ["Beta", "York"],
    ]);
  });

  it("respects case and spacing when asked to", () => {
    const result = dedupeGrid(grid, { ...dedupeDefaults, ignoreCaseAndSpacing: false });
    expect(result.removedCount).toBe(1);
    expect(result.removedRowOffsets).toEqual([4]);
  });

  it("compares only the key columns", () => {
    const byName = dedupeGrid(
      [
        ["Name", "City"],
        ["Acme", "Leeds"],
        ["Acme", "York"],
      ],
      { ...dedupeDefaults, keyColumns: [0] }
    );
    expect(byName.removedCount).toBe(1);
  });

  it("treats the first row as data when there is no header", () => {
    const result = dedupeGrid([["Acme"], ["Acme"]], {
      ...dedupeDefaults,
      hasHeaderRow: false,
    });
    expect(result.removedCount).toBe(1);
    expect(result.rows).toEqual([["Acme"]]);
  });

  it("never folds blank rows into one another", () => {
    const result = dedupeGrid([["Name"], ["Acme"], [null], [""], ["Acme"]], dedupeDefaults);
    expect(result.removedCount).toBe(1);
    expect(result.rows).toHaveLength(4);
  });

  it("does not let adjacent cells run together into the same key", () => {
    const result = dedupeGrid(
      [
        ["A", "B"],
        ["ab", "c"],
        ["a", "bc"],
      ],
      { ...dedupeDefaults, ignoreCaseAndSpacing: false }
    );
    expect(result.removedCount).toBe(0);
  });

  it("treats an empty cell and an empty string as the same value", () => {
    const result = dedupeGrid(
      [
        ["A", "B"],
        [null, "x"],
        ["", "x"],
      ],
      { ...dedupeDefaults, ignoreCaseAndSpacing: false }
    );
    // Both rows key to an empty first cell, so the second is a duplicate.
    expect(result.removedCount).toBe(1);
  });
});

describe("mapGrid", () => {
  it("counts only the cells that actually changed", () => {
    const { grid, changedCount } = mapGrid(
      [
        ["a", "B"],
        ["c", "d"],
      ],
      (value) => (typeof value === "string" ? value.toLowerCase() : value)
    );
    expect(changedCount).toBe(1);
    expect(grid).toEqual([
      ["a", "b"],
      ["c", "d"],
    ]);
  });
});

describe("groupContiguous", () => {
  it("merges runs and leaves gaps separate", () => {
    expect(groupContiguous([1, 2, 3, 7, 9, 10])).toEqual([
      { start: 1, count: 3 },
      { start: 7, count: 1 },
      { start: 9, count: 2 },
    ]);
  });

  it("sorts before grouping", () => {
    expect(groupContiguous([5, 3, 4])).toEqual([{ start: 3, count: 3 }]);
  });

  it("handles an empty list", () => {
    expect(groupContiguous([])).toEqual([]);
  });
});
