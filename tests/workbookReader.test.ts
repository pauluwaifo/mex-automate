import {
  isSupportedFile,
  parseDelimited,
  sniffDelimiter,
} from "../src/taskpane/shared/workbookReader";
import { parseTemplate } from "../src/taskpane/shared/templateStore";
import { DEFAULT_REFRESH_OPTIONS } from "../src/taskpane/shared/types";

describe("sniffDelimiter", () => {
  it("picks the delimiter that gives a consistent column count", () => {
    expect(sniffDelimiter("a,b,c\n1,2,3")).toBe(",");
    expect(sniffDelimiter("a;b;c\n1;2;3")).toBe(";");
    expect(sniffDelimiter("a\tb\tc\n1\t2\t3")).toBe("\t");
    expect(sniffDelimiter("a|b|c\n1|2|3")).toBe("|");
  });

  it("is not fooled by commas inside a semicolon-delimited file", () => {
    expect(sniffDelimiter('name;note\nAcme;"a, b, c"\nBeta;"d, e, f"')).toBe(";");
  });

  it("falls back to a comma when there is nothing to go on", () => {
    expect(sniffDelimiter("single-column")).toBe(",");
  });
});

describe("parseDelimited", () => {
  it("parses a simple CSV and coerces numbers", () => {
    expect(parseDelimited("Name,Amount\nAcme,100\nBeta,250.5")).toEqual([
      ["Name", "Amount"],
      ["Acme", 100],
      ["Beta", 250.5],
    ]);
  });

  it("keeps quoted fields as text even when they look numeric", () => {
    // Leading zeros matter: "007" is an account code, not the number 7.
    expect(parseDelimited('code,n\n"007",007')).toEqual([
      ["code", "n"],
      ["007", "007"],
    ]);
  });

  it("handles quoted fields containing the delimiter", () => {
    expect(parseDelimited('a,b\n"one, two",three')).toEqual([
      ["a", "b"],
      ["one, two", "three"],
    ]);
  });

  it("handles escaped quotes", () => {
    expect(parseDelimited('a\n"say ""hi"""')).toEqual([["a"], ['say "hi"']]);
  });

  it("handles a newline inside a quoted field", () => {
    expect(parseDelimited('a,b\n"line1\nline2",x')).toEqual([
      ["a", "b"],
      ["line1\nline2", "x"],
    ]);
  });

  it("handles CRLF line endings", () => {
    expect(parseDelimited("a,b\r\n1,2\r\n")).toEqual([
      ["a", "b"],
      [1, 2],
    ]);
  });

  it("strips a UTF-8 BOM from the first header", () => {
    const grid = parseDelimited("\uFEFFName,Amount\nAcme,1");
    expect(grid[0][0]).toBe("Name");
  });

  it("pads short rows so the grid is rectangular", () => {
    expect(parseDelimited("a,b,c\n1,2")).toEqual([
      ["a", "b", "c"],
      [1, 2, null],
    ]);
  });

  it("turns empty fields into nulls", () => {
    expect(parseDelimited("a,b\n,2")).toEqual([
      ["a", "b"],
      [null, 2],
    ]);
  });

  it("accepts an explicit delimiter", () => {
    expect(parseDelimited("a\tb\n1\t2", "\t")).toEqual([
      ["a", "b"],
      [1, 2],
    ]);
  });

  it("returns an empty grid for empty input", () => {
    expect(parseDelimited("")).toEqual([]);
    expect(parseDelimited("\n\n")).toEqual([]);
  });
});

describe("isSupportedFile", () => {
  it("accepts the formats the reader understands", () => {
    expect(isSupportedFile("data.xlsx")).toBe(true);
    expect(isSupportedFile("DATA.CSV")).toBe(true);
    expect(isSupportedFile("macro.xlsm")).toBe(true);
  });

  it("rejects everything else", () => {
    expect(isSupportedFile("old.xls")).toBe(false);
    expect(isSupportedFile("report.pdf")).toBe(false);
  });
});

describe("parseTemplate", () => {
  const valid = {
    schemaVersion: 1,
    template: {
      name: "Monthly",
      createdAt: "2024-01-01T00:00:00.000Z",
      updatedAt: "2024-02-01T00:00:00.000Z",
      zones: [],
      mappings: {},
      options: DEFAULT_REFRESH_OPTIONS,
    },
  };

  it("reads the wrapped shape from a JSON string", () => {
    expect(parseTemplate(JSON.stringify(valid))?.name).toBe("Monthly");
  });

  it("reads an already-parsed object", () => {
    expect(parseTemplate(valid)?.name).toBe("Monthly");
  });

  it("accepts a bare template written by an older build", () => {
    expect(parseTemplate(valid.template)?.name).toBe("Monthly");
  });

  it("fills in defaults for anything missing", () => {
    const parsed = parseTemplate({ name: "Sparse" });
    expect(parsed).not.toBeNull();
    expect(parsed!.zones).toEqual([]);
    expect(parsed!.mappings).toEqual({});
    expect(parsed!.options).toEqual(DEFAULT_REFRESH_OPTIONS);
    expect(Date.parse(parsed!.createdAt)).not.toBeNaN();
  });

  it("keeps options the caller did set", () => {
    const parsed = parseTemplate({ name: "X", options: { recalculate: false } });
    expect(parsed!.options.recalculate).toBe(false);
    expect(parsed!.options.clearExistingRows).toBe(true);
  });

  it("returns null for anything unusable", () => {
    expect(parseTemplate("not json")).toBeNull();
    expect(parseTemplate(null)).toBeNull();
    expect(parseTemplate(42)).toBeNull();
    expect(parseTemplate({ name: "   " })).toBeNull();
    expect(parseTemplate({ zones: [] })).toBeNull();
  });
});
