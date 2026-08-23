import {
  buildFilledFormulas,
  findTemplate,
  FORMULA_TEMPLATES,
  translateFormula,
} from "../src/taskpane/features/formulaPatterns";

const target = { firstRow: 2, firstColumn: "D", lastRow: 100, rowCount: 99 };

describe("translateFormula", () => {
  it("shifts relative references and leaves anchored parts alone", () => {
    expect(translateFormula("=SUM($B$2:B2)", 1, 0)).toBe("=SUM($B$2:B3)");
    expect(translateFormula("=A1+B1", 2, 1)).toBe("=B3+C3");
    expect(translateFormula("=$A1", 3, 5)).toBe("=$A4");
    expect(translateFormula("=A$1", 3, 5)).toBe("=F$1");
  });

  it("leaves a non-formula untouched", () => {
    expect(translateFormula("A1", 5, 5)).toBe("A1");
    expect(translateFormula("", 1, 1)).toBe("");
  });

  it("does not touch references inside string literals", () => {
    expect(translateFormula('=IF(A1>0,"see A1","")', 1, 0)).toBe('=IF(A2>0,"see A1","")');
  });

  it("handles doubled quotes inside a string literal", () => {
    expect(translateFormula('=A1&""""&B1', 1, 0)).toBe('=A2&""""&B2');
  });

  it("keeps function names intact", () => {
    expect(translateFormula("=LOG10(A1)", 1, 0)).toBe("=LOG10(A2)");
    expect(translateFormula("=SUM(A1:A10)", 1, 0)).toBe("=SUM(A2:A11)");
  });

  it("shifts references on another sheet", () => {
    expect(translateFormula("=Sheet1!A1", 1, 0)).toBe("=Sheet1!A2");
    expect(translateFormula("='My Sheet'!A1", 1, 0)).toBe("='My Sheet'!A2");
  });

  it("does not mistake a sheet name for a cell reference", () => {
    // "AB1" looks exactly like a cell reference until you see the "!".
    expect(translateFormula("=AB1!A2", 1, 0)).toBe("=AB1!A3");
  });

  it("leaves structured table references alone", () => {
    expect(translateFormula("=SUM(Table1[Amount])", 5, 3)).toBe("=SUM(Table1[Amount])");
    expect(translateFormula("=Table1[@Price]*B2", 1, 0)).toBe("=Table1[@Price]*B3");
  });

  it("leaves a defined name that merely looks like a reference alone", () => {
    expect(translateFormula("=A1*Rate1", 1, 0)).toBe("=A2*Rate1");
  });

  it("produces #REF! when a reference is pushed off the grid", () => {
    expect(translateFormula("=A1", -1, 0)).toBe("=#REF!");
    expect(translateFormula("=A2", 0, -1)).toBe("=#REF!");
  });

  it("preserves a lookup's anchored table while shifting its key", () => {
    expect(translateFormula("=VLOOKUP(A2,Prices!$A$2:$B$100,2,FALSE)", 1, 0)).toBe(
      "=VLOOKUP(A3,Prices!$A$2:$B$100,2,FALSE)"
    );
  });
});

describe("buildFilledFormulas", () => {
  it("translates across both axes", () => {
    expect(buildFilledFormulas("=A1", 2, 2)).toEqual([
      ["=A1", "=B1"],
      ["=A2", "=B2"],
    ]);
  });

  it("returns the source formula for a 1x1 fill", () => {
    expect(buildFilledFormulas("=SUM($A$1:A1)", 1, 1)).toEqual([["=SUM($A$1:A1)"]]);
  });
});

describe("formula template library", () => {
  it("gives every template a unique id and at least one input", () => {
    const ids = FORMULA_TEMPLATES.map((template) => template.id);
    expect(new Set(ids).size).toBe(ids.length);
    for (const template of FORMULA_TEMPLATES) {
      expect(template.inputs.length).toBeGreaterThan(0);
      expect(template.name).not.toBe("");
    }
  });

  it("builds a formula starting with = for every template, even with no input", () => {
    for (const template of FORMULA_TEMPLATES) {
      const formula = template.build({}, target);
      expect(formula.startsWith("=")).toBe(true);
    }
  });

  it("anchors the start of a running total and lets the end drift", () => {
    const formula = findTemplate("running-total")!.build({ column: "C" }, target);
    expect(formula).toBe("=SUM($C$2:C2)");
    // Filling down must extend the window, not move it.
    expect(translateFormula(formula, 1, 0)).toBe("=SUM($C$2:C3)");
  });

  it("anchors both ends of a % of total so the denominator stays fixed", () => {
    const formula = findTemplate("percent-of-total")!.build({ column: "C" }, target);
    expect(formula).toBe('=IFERROR(C2/SUM($C$2:$C$100),"")');
    expect(translateFormula(formula, 1, 0)).toBe('=IFERROR(C3/SUM($C$2:$C$100),"")');
  });

  it("compares against the row above for % change", () => {
    expect(findTemplate("percent-change")!.build({ column: "C" }, target)).toBe(
      '=IFERROR((C2-C1)/C1,"")'
    );
  });

  it("defaults a value column to the column left of the target", () => {
    expect(findTemplate("running-total")!.build({}, target)).toBe("=SUM($C$2:C2)");
  });

  it("builds the lookups with the supplied ranges", () => {
    expect(
      findTemplate("xlookup")!.build(
        {
          keyColumn: "A",
          lookupRange: "Prices!$A$2:$A$500",
          returnRange: "Prices!$B$2:$B$500",
          notFound: "n/a",
        },
        target
      )
    ).toBe('=XLOOKUP(A2,Prices!$A$2:$A$500,Prices!$B$2:$B$500,"n/a")');

    expect(
      findTemplate("index-match")!.build(
        { keyColumn: "A", returnRange: "$B$2:$B$50", lookupRange: "$A$2:$A$50" },
        target
      )
    ).toBe('=IFERROR(INDEX($B$2:$B$50,MATCH(A2,$A$2:$A$50,0)),"Not found")');
  });

  it("splits a joined-column list on commas or spaces", () => {
    expect(findTemplate("join-columns")!.build({ columns: "a, b c", separator: "-" }, target)).toBe(
      '=TEXTJOIN("-",TRUE,A2,B2,C2)'
    );
  });

  it("returns undefined for an unknown template", () => {
    expect(findTemplate("nope")).toBeUndefined();
  });
});
