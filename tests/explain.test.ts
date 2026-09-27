import { explainFormula, parseFormula } from "../src/taskpane/features/explain";

const english = (formula: string) => explainFormula(formula).english;

describe("parseFormula", () => {
  it("respects precedence", () => {
    expect(parseFormula("=1+2*3")).toEqual({
      kind: "binary",
      operator: "+",
      left: { kind: "number", value: 1 },
      right: {
        kind: "binary",
        operator: "*",
        left: { kind: "number", value: 2 },
        right: { kind: "number", value: 3 },
      },
    });
  });

  it("groups powers rightwards, as Excel does", () => {
    const node = parseFormula("=2^3^2");
    expect(node).toMatchObject({ operator: "^", right: { operator: "^" } });
  });

  it("reads brackets, sheet names and absolute references", () => {
    expect(parseFormula("=('My Sheet'!$A$1+Sheet2!B2)*2")).toMatchObject({ operator: "*" });
  });

  it("does not mistake a function whose name ends in a digit for a reference", () => {
    expect(parseFormula("=LOG10(100)")).toMatchObject({ kind: "call", name: "LOG10" });
  });

  it("reads doubled quotes inside text", () => {
    expect(parseFormula('="he said ""hi"""')).toEqual({
      kind: "string",
      value: 'he said "hi"',
    });
  });
});

describe("explainFormula", () => {
  it("explains arithmetic", () => {
    expect(english("=A2*B2")).toBe("A2 times B2.");
    expect(english("=(A2-B2)/B2")).toBe("A2 minus B2 divided by B2.");
    expect(english("=-A1")).toBe("Minus A1.");
    expect(english("=A1*10%")).toBe("A1 times 10%.");
  });

  it("explains the functions people actually write", () => {
    expect(english("=SUM(B2:B10)")).toBe("The total of B2:B10.");
    expect(english("=AVERAGE(B2:B10)")).toBe("The average of B2:B10.");
    expect(english('=COUNTIF(A:A,"North")')).toBe('How many cells in A:A match "North".');
    expect(english("=ROUND(A1,2)")).toBe("A1 rounded to 2 decimal places.");
  });

  it("explains a lookup in terms of what it does", () => {
    expect(english("=VLOOKUP(A2,Sheet2!A:C,3,FALSE)")).toBe(
      "A2 looked up in the first column of Sheet2!A:C, returning column 3, requiring an exact match."
    );
    expect(english('=XLOOKUP(A2,Products!A:A,Products!C:C,"none")')).toBe(
      'A2 looked up in Products!A:A, returning the matching value from Products!C:C, or "none" when there is no match.'
    );
  });

  it("explains nested calls as one sentence", () => {
    expect(english("=IFERROR(VLOOKUP(A2,Data!A:B,2,FALSE),0)")).toBe(
      "A2 looked up in the first column of Data!A:B, returning column 2, requiring an exact match, or 0 if that gives an error."
    );
    expect(english('=IF(SUM(B2:B10)>100,"High","Low")')).toBe(
      'If the total of B2:B10 is greater than 100, then "High", otherwise "Low".'
    );
  });

  it("falls back to the function name when it has no phrase for it", () => {
    expect(english("=SQRT(A1)")).toBe("SQRT of A1.");
  });

  it("lists what the formula reads and which functions it uses", () => {
    const explanation = explainFormula("=SUM(B2:B10)+Sheet2!C1");
    expect(explanation.references).toEqual(["B2:B10", "Sheet2!C1"]);
    expect(explanation.functions).toEqual(["SUM"]);
  });

  it("says plainly when a cell is not a formula", () => {
    expect(english("")).toBe("That cell is empty.");
    expect(english("1250")).toBe("That cell holds a typed-in value, not a formula.");
  });

  it("admits when it cannot read something rather than guessing", () => {
    const broken = explainFormula("=SUM(B2:B10");
    expect(broken.english).toContain("I can't read that formula");
    expect(broken.english).toContain("bracket");
    expect(explainFormula('=CONCAT("a')?.english).toContain("closing quote");
  });
});

describe("warnings", () => {
  const kinds = (formula: string) =>
    explainFormula(formula).warnings.map((warning) => warning.kind);

  it("notices a rate typed into the formula", () => {
    expect(kinds("=B2*0.075")).toContain("hardcodedNumber");
    // Small whole numbers are usually structural, not amounts.
    expect(kinds("=ROUND(A1,2)")).not.toContain("hardcodedNumber");
  });

  it("notices a whole-column reference", () => {
    expect(kinds("=SUM(B:B)")).toContain("wholeColumn");
    expect(kinds("=SUM(B2:B99)")).not.toContain("wholeColumn");
  });

  it("notices deeply nested IFs", () => {
    expect(kinds('=IF(A1=1,"a",IF(A1=2,"b",IF(A1=3,"c","d")))')).toContain("deepNesting");
    expect(kinds('=IF(A1=1,"a","b")')).not.toContain("deepNesting");
  });

  it("notices a division with no guard, and stays quiet when there is one", () => {
    expect(kinds("=A1/B1")).toContain("divideRisk");
    expect(kinds("=IFERROR(A1/B1,0)")).not.toContain("divideRisk");
  });

  it("notices a formula that never settles", () => {
    expect(kinds("=TODAY()-A1")).toContain("volatile");
  });

  it("explains why counting columns in a VLOOKUP is fragile", () => {
    const warning = explainFormula("=VLOOKUP(A2,Data!A:C,3,FALSE)").warnings.find(
      (candidate) => candidate.kind === "lookupByPosition"
    );
    expect(warning?.text).toContain("inserting a column");
  });
});
