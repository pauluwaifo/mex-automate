import {
  checkSheet,
  checkWorkbook,
  dataExtent,
  formulaShape,
  HealthFinding,
  SheetSnapshot,
  summarizeHealth,
  tightenWholeColumns,
} from "../src/taskpane/features/health";
import { Grid } from "../src/taskpane/shared/types";

/** Builds a snapshot from parallel value/formula grids. */
function sheet(
  name: string,
  values: Grid,
  formulas: string[][],
  extra: Partial<SheetSnapshot> = {}
): SheetSnapshot {
  return { name, values, formulas, origin: { row: 0, column: 0 }, ...extra };
}

const kinds = (findings: HealthFinding[]) => findings.map((finding) => finding.kind);

describe("dataExtent", () => {
  it("finds where the data really stops", () => {
    expect(
      dataExtent([
        ["a", "b"],
        ["c", null],
        [null, null],
      ])
    ).toEqual({ rows: 2, columns: 2 });
  });

  it("treats an empty string as empty", () => {
    expect(dataExtent([["a"], [""], [""]])).toEqual({ rows: 1, columns: 1 });
  });

  it("copes with a sheet holding nothing", () => {
    expect(dataExtent([[null, null]])).toEqual({ rows: 0, columns: 0 });
  });
});

describe("formulaShape", () => {
  it("collapses copies of one formula down a column", () => {
    expect(formulaShape("=D2*E2")).toBe(formulaShape("=D900*E900"));
  });

  it("keeps different formulas apart", () => {
    expect(formulaShape("=D2*E2")).not.toBe(formulaShape("=D2+E2"));
  });
});

describe("checkSheet", () => {
  it("counts a repeated formula by cells, not by ideas", () => {
    const values: Grid = Array.from({ length: 30 }, (_v, i) => [i + 1, null]);
    const formulas = Array.from({ length: 30 }, () => ["", "=SUM(B:B)"]);
    const findings = checkSheet(sheet("Data", values, formulas));
    const wholeColumn = findings.find((finding) => finding.kind === "wholeColumn");
    expect(wholeColumn?.count).toBe(30);
    // Examples are addresses, capped so the message stays readable.
    expect(wholeColumn?.examples.length).toBeLessThanOrEqual(5);
  });

  it("finds the things that make a file slow", () => {
    const values: Grid = [[1], [2]];
    const formulas = [["=SUM(A:A)"], ["=TODAY()"]];
    expect(kinds(checkSheet(sheet("S", values, formulas)))).toEqual(
      expect.arrayContaining(["wholeColumn", "volatile"])
    );
  });

  it("finds the things that break quietly", () => {
    const values: Grid = [[1], [2]];
    const formulas = [["=VLOOKUP(A1,Data!A:C,3,FALSE)"], ["='[Budget.xlsx]Sheet1'!A1"]];
    expect(kinds(checkSheet(sheet("S", values, formulas)))).toEqual(
      expect.arrayContaining(["lookupByPosition", "externalLink"])
    );
  });

  it("sees a sheet Excel thinks is far bigger than its data", () => {
    const values: Grid = Array.from({ length: 900 }, (_v, i) => [i < 40 ? i : null]);
    const formulas = values.map(() => [""]);
    const bloat = checkSheet(sheet("Padded", values, formulas)).find(
      (finding) => finding.kind === "usedRangeBloat"
    );
    expect(bloat).toBeTruthy();
    expect(bloat?.count).toBe(860);
    expect(bloat?.fix).toMatchObject({ kind: "clearBeyondData", firstRow: 40 });
  });

  it("does not complain about a sheet that merely has a few spare rows", () => {
    const values: Grid = Array.from({ length: 50 }, (_v, i) => [i < 45 ? i : null]);
    expect(
      kinds(
        checkSheet(
          sheet(
            "Fine",
            values,
            values.map(() => [""])
          )
        )
      )
    ).not.toContain("usedRangeBloat");
  });

  it("mentions conditional formatting only when there is a lot of it", () => {
    const values: Grid = [[1]];
    const formulas = [[""]];
    expect(
      kinds(checkSheet(sheet("A", values, formulas, { conditionalFormats: 4 })))
    ).not.toContain("conditionalFormatBloat");
    expect(kinds(checkSheet(sheet("B", values, formulas, { conditionalFormats: 90 })))).toContain(
      "conditionalFormatBloat"
    );
  });

  it("offers a fix only where it has one", () => {
    const values: Grid = [[1], [2]];
    const findings = checkSheet(sheet("S", values, [["=SUM(A:A)"], ["=TODAY()"]]));
    expect(findings.find((f) => f.kind === "wholeColumn")?.fix).toBeTruthy();
    expect(findings.find((f) => f.kind === "volatile")?.fix).toBeUndefined();
  });

  it("says nothing about a plain, healthy sheet", () => {
    const values: Grid = [
      ["Region", "Revenue"],
      ["North", 100],
      ["South", 200],
    ];
    const formulas = [
      ["", ""],
      ["", ""],
      ["", "=B2*C2"],
    ];
    expect(checkSheet(sheet("Clean", values, formulas))).toEqual([]);
  });

  it("reports addresses using the sheet's real origin", () => {
    const snapshot: SheetSnapshot = {
      name: "Offset",
      values: [[1]],
      formulas: [["=SUM(C:C)"]],
      origin: { row: 4, column: 2 },
    };
    expect(checkSheet(snapshot)[0].examples[0]).toBe("C5");
  });
});

describe("checkWorkbook", () => {
  const slowSheet = sheet("Slow", [[1], [2]], [["=SUM(A:A)"], ["=SUM(B:B)"]]);
  const fragileSheet = sheet("Fragile", [[1]], [["=VLOOKUP(A1,D:F,3,FALSE)"]]);

  it("puts what makes it slow before what might break", () => {
    const report = checkWorkbook([fragileSheet, slowSheet]);
    expect(report.findings[0].severity).toBe("slow");
  });

  it("counts formula cells across the workbook", () => {
    expect(checkWorkbook([slowSheet, fragileSheet]).formulaCells).toBe(3);
  });

  it("remembers the sheets it could not read", () => {
    const report = checkWorkbook([slowSheet], ["Huge"]);
    expect(report.sheetsSkipped).toEqual(["Huge"]);
  });

  it("summarizes in one sentence", () => {
    expect(summarizeHealth(checkWorkbook([slowSheet, fragileSheet]))).toContain("making it slower");
    const clean = checkWorkbook([sheet("Clean", [["a"]], [[""]])]);
    expect(summarizeHealth(clean)).toContain("good shape");
  });
});

describe("tightenWholeColumns", () => {
  it("narrows a whole-column reference to the data", () => {
    expect(tightenWholeColumns("=SUM(B:B)", 2, 812)).toBe("=SUM(B2:B812)");
  });

  it("narrows every reference in the formula", () => {
    expect(tightenWholeColumns('=SUMIFS(C:C,A:A,"North")', 2, 500)).toBe(
      '=SUMIFS(C2:C500,A2:A500,"North")'
    );
  });

  it("keeps the dollar signs it was given", () => {
    expect(tightenWholeColumns("=SUM($B:$B)", 2, 100)).toBe("=SUM($B2:$B100)");
  });

  it("leaves a formula alone when there is nothing to narrow", () => {
    expect(tightenWholeColumns("=SUM(B2:B99)", 2, 812)).toBeNull();
    expect(tightenWholeColumns("=A1*2", 2, 812)).toBeNull();
  });

  it("refuses when the data extent makes no sense", () => {
    expect(tightenWholeColumns("=SUM(B:B)", 10, 2)).toBeNull();
  });
});
