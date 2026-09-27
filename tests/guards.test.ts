import {
  countRejections,
  Guard,
  guardRange,
  suggestGuards,
  summarizeGuards,
  toIsoDate,
  toValidationRule,
  uniqueFormula,
  wouldReject,
} from "../src/taskpane/features/guards";
import { dateToExcelSerial } from "../src/taskpane/features/dataCleaning";
import { profileColumn } from "../src/taskpane/features/tidy";
import { columnLetter } from "../src/taskpane/shared/excelHelpers";
import { Grid } from "../src/taskpane/shared/types";

const HEADERS = ["Order ID", "Order Date", "Region", "Units", "Revenue", "Note"];
const serial = (m: number, d: number) => dateToExcelSerial(new Date(Date.UTC(2024, m, d)));

/** 24 tidy orders: four regions, whole-number units, money, unique ids. */
const ROWS: Grid = Array.from({ length: 24 }, (_v, i) => [
  `SO-${1000 + i}`,
  serial(i % 6, 1 + (i % 27)),
  ["North", "South", "East", "West"][i % 4],
  1 + (i % 9),
  (1 + (i % 9)) * 12.5,
  `note ${i}`,
]);

const profiles = HEADERS.map((header, index) =>
  profileColumn(
    header,
    ROWS.map((row) => row[index]),
    index
  )
);

const guards = suggestGuards(profiles, ROWS);
const byHeader = (header: string) => guards.find((guard) => guard.header === header);

describe("suggestGuards", () => {
  it("offers a dropdown for a settled set of categories", () => {
    const region = byHeader("Region");
    expect(region?.kind).toBe("list");
    expect(region?.values).toEqual(["North", "South", "East", "West"]);
  });

  it("does not offer a dropdown for a column that is nearly all distinct", () => {
    expect(byHeader("Note")?.kind).not.toBe("list");
  });

  it("keeps a unique id unique", () => {
    expect(byHeader("Order ID")?.kind).toBe("unique");
  });

  it("does not claim uniqueness for a column that already repeats", () => {
    const repeated = ROWS.map((row, i) => [i < 3 ? "SO-1000" : row[0], ...row.slice(1)]);
    const repeatedProfiles = HEADERS.map((header, index) =>
      profileColumn(
        header,
        repeated.map((row) => row[index]),
        index
      )
    );
    const found = suggestGuards(repeatedProfiles, repeated).find(
      (guard) => guard.header === "Order ID"
    );
    expect(found?.kind).not.toBe("unique");
  });

  it("reads a quantity as a whole number that is never negative", () => {
    const units = byHeader("Units");
    expect(units?.kind).toBe("wholeNumber");
    expect(units?.min).toBe(0);
  });

  it("leaves room above today's largest value", () => {
    const units = byHeader("Units")!;
    const largest = Math.max(...ROWS.map((row) => row[3] as number));
    expect(units.max).toBeGreaterThan(largest);
  });

  it("allows decimals where the column has them", () => {
    expect(byHeader("Revenue")?.kind).toBe("decimal");
  });

  it("bounds a date column either side of the dates it holds", () => {
    const date = byHeader("Order Date")!;
    expect(date.kind).toBe("dateRange");
    expect(date.min).toBeLessThan(serial(0, 1));
    expect(date.max).toBeGreaterThan(serial(5, 27));
  });

  it("says nothing about a table too short to learn anything from", () => {
    const few = ROWS.slice(0, 3);
    const fewProfiles = HEADERS.map((header, index) =>
      profileColumn(
        header,
        few.map((row) => row[index]),
        index
      )
    );
    expect(suggestGuards(fewProfiles, few)).toEqual([]);
  });

  it("explains each rule in the user's terms", () => {
    for (const guard of guards) {
      expect(guard.title.length).toBeGreaterThan(10);
      expect(guard.detail.length).toBeGreaterThan(10);
      expect(guard.errorMessage).toContain(guard.header);
    }
  });
});

describe("summarizeGuards", () => {
  it("counts what it found", () => {
    expect(summarizeGuards(guards, "Orders")).toContain(`${guards.length} rules`);
  });

  it("explains itself when it found nothing", () => {
    expect(summarizeGuards([], "Orders")).toContain("couldn't see a column");
  });
});

describe("wouldReject", () => {
  const list: Guard = {
    header: "Region",
    columnIndex: 0,
    kind: "list",
    title: "",
    detail: "",
    errorMessage: "",
    values: ["North", "South"],
  };

  it("rejects a value outside the list, ignoring case", () => {
    expect(wouldReject(list, "north")).toBe(false);
    expect(wouldReject(list, "Nrth")).toBe(true);
  });

  it("never rejects an empty cell", () => {
    expect(wouldReject(list, null)).toBe(false);
    expect(wouldReject(list, "")).toBe(false);
  });

  it("rejects numbers outside the bounds, and text that is not a number", () => {
    const number: Guard = {
      header: "Units",
      columnIndex: 0,
      kind: "wholeNumber",
      title: "",
      detail: "",
      errorMessage: "",
      min: 0,
      max: 100,
    };
    expect(wouldReject(number, 50)).toBe(false);
    expect(wouldReject(number, -1)).toBe(true);
    expect(wouldReject(number, 101)).toBe(true);
    expect(wouldReject(number, 2.5)).toBe(true);
    expect(wouldReject(number, "lots")).toBe(true);
    // A number typed as text is still a number.
    expect(wouldReject(number, "1,000")).toBe(true);
    expect(wouldReject(number, "50")).toBe(false);
  });

  it("counts how many existing cells a rule would have caught", () => {
    const rows: Grid = [["North"], ["Nrth"], ["South"], [null]];
    expect(countRejections({ ...list, columnIndex: 0 }, rows)).toBe(1);
  });
});

describe("the shape handed to Excel", () => {
  it("builds a list rule with a dropdown", () => {
    const rule = toValidationRule(
      { ...(byHeader("Region") as Guard) },
      { rangeAddress: "C2:C30", firstCell: "C2" }
    );
    expect(rule).toEqual({
      list: { inCellDropDown: true, source: "North,South,East,West" },
    });
  });

  it("builds numeric bounds", () => {
    const rule = toValidationRule(byHeader("Units") as Guard, {
      rangeAddress: "D2:D30",
      firstCell: "D2",
    }) as { wholeNumber: { operator: string } };
    expect(rule.wholeNumber.operator).toBe("Between");
  });

  it("gives dates as ISO strings, not serials", () => {
    const rule = toValidationRule(byHeader("Order Date") as Guard, {
      rangeAddress: "B2:B30",
      firstCell: "B2",
    }) as { date: { formula1: string } };
    expect(rule.date.formula1).toMatch(/^\d{4}-\d{2}-\d{2}$/);
    expect(toIsoDate(serial(0, 15))).toBe("2024-01-15");
  });

  it("builds a COUNTIF that allows exactly one of each", () => {
    expect(uniqueFormula("A2:A30", "A2")).toBe("=COUNTIF($A$2:$A$30,A2)=1");
  });

  it("covers empty rows below the data, which is where typing happens", () => {
    expect(guardRange(columnLetter, 2, 2, 25, 200)).toBe("C2:C225");
  });
});
