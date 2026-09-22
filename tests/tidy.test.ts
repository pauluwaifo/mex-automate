import {
  buildHeaders,
  classifyRow,
  cleanText,
  countProblems,
  detectDayFirst,
  detectNumberStyle,
  editDistance,
  FindingId,
  isPlaceholder,
  locateTable,
  numberFormatFor,
  parseLooseNumber,
  profileColumn,
  splitSectionLabel,
  tidyTable,
  unifySpellings,
} from "../src/taskpane/features/tidy";
import { excelSerialToDate } from "../src/taskpane/features/dataCleaning";
import { Grid } from "../src/taskpane/shared/types";

/** The kind of export that ruins an afternoon. */
const MESSY: Grid = [
  ["ACME FOODS LTD - SALES EXTRACT", null, null, null, null, null],
  ["Generated 02/04/2024 by SAP", null, null, null, null, null],
  [null, null, null, null, null, null],
  ["Order ID", "Order Date", "Customer", "Channel", "Units", "Revenue"],
  ["Region: Europe", null, null, null, null, null],
  ["A-100", "03/04/2024", "  Acme Ltd ", "Online", "3", "$1,204.50"],
  ["A-101", "2024-04-05", "ACME LTD", "Retail", 2, "1.130,00 €"],
  ["A-102", "Apr 6, 2024", "Beta Trading", "online", "N/A", "(250)"],
  ["A-102", "Apr 6, 2024", "Beta Trading", "online", "N/A", "(250)"],
  ["Europe Total", null, null, null, 5, 2084.5],
  [null, null, null, null, null, null],
  ["Region: Africa", null, null, null, null, null],
  ["A-103", "13/04/2024", "Gamma Co", "Retail", "4", "980-"],
  ["A-104", 45397, "Gamma  Co", "Wholesale", "1", "1,2O0"],
  ["Order ID", "Order Date", "Customer", "Channel", "Units", "Revenue"],
  ["A-105", "16.04.2024", "Delta Supplies", "Wholesale", "7", "NGN 45,000"],
  ["Africa Total", null, null, null, 12, 45220],
  ["Grand Total", null, null, null, 17, 47304.5],
  ["Source: SAP ERP, exported by finance for internal use only", null, null, null, null, null],
];

describe("cleanText / isPlaceholder", () => {
  it("collapses invisible and repeated whitespace", () => {
    expect(cleanText("  Acme  Ltd ")).toBe("Acme Ltd");
  });

  it("recognises the ways people write 'no value'", () => {
    for (const value of ["-", "—", "N/A", "n/a", "#N/A", "null", "?"]) {
      expect(isPlaceholder(value)).toBe(true);
    }
    expect(isPlaceholder("Nairobi")).toBe(false);
    expect(isPlaceholder(0)).toBe(false);
  });
});

describe("parseLooseNumber", () => {
  const value = (text: string, style: "us" | "eu" = "us") => parseLooseNumber(text, style)?.value;

  it("reads currency, thousands and decimals", () => {
    expect(value("$1,204.50")).toBe(1204.5);
    expect(value("NGN 45,000")).toBe(45000);
    expect(value("₦45,000.75")).toBe(45000.75);
    expect(value("£ 12")).toBe(12);
  });

  it("reads European formats", () => {
    expect(value("1.130,00 €")).toBe(1130);
    expect(value("1.234.567")).toBe(1234567);
    expect(value("12,5")).toBe(12.5);
    expect(value("1.234", "eu")).toBe(1234);
  });

  it("reads accounting and ERP negatives", () => {
    expect(value("(250)")).toBe(-250);
    expect(value("980-")).toBe(-980);
    expect(value("-$75")).toBe(-75);
    expect(value("$-75")).toBe(-75);
  });

  it("reads percentages as fractions", () => {
    const parsed = parseLooseNumber("12.5%");
    expect(parsed?.value).toBeCloseTo(0.125);
    expect(parsed?.percent).toBe(true);
  });

  it("reads spaced and apostrophe thousands", () => {
    expect(value("1 234 567")).toBe(1234567);
    expect(value("1'234")).toBe(1234);
  });

  it("repairs a letter O typed for a zero, and says so", () => {
    const parsed = parseLooseNumber("1,2O0");
    expect(parsed?.value).toBe(1200);
    expect(parsed?.repaired).toBe(true);
    expect(parseLooseNumber("1,200")?.repaired).toBe(false);
  });

  it("records the currency", () => {
    expect(parseLooseNumber("$5")?.currency).toBe("$");
    expect(parseLooseNumber("NGN 5")?.currency).toBe("NGN");
  });

  it("refuses things that are not numbers", () => {
    for (const text of [
      "Q1",
      "2024-01-05",
      "03/04/2024",
      "Acme",
      "",
      "N/A",
      "1,23,4",
      "1.2.3,4.5",
    ]) {
      expect(parseLooseNumber(text)).toBeNull();
    }
  });

  it("passes real numbers straight through", () => {
    expect(value(42 as unknown as string)).toBe(42);
  });
});

describe("detectNumberStyle", () => {
  it("spots European columns from their decimal commas", () => {
    expect(detectNumberStyle(["1.130,00", "12,50", "7"])).toBe("eu");
    expect(detectNumberStyle(["1,130.00", "12.50", "7"])).toBe("us");
  });

  it("defaults to US when nothing gives it away", () => {
    expect(detectNumberStyle(["1,234", "7"])).toBe("us");
  });
});

describe("detectDayFirst", () => {
  it("lets unambiguous dates decide", () => {
    expect(detectDayFirst(["03/04/2024", "13/04/2024"], false)).toBe(true);
    expect(detectDayFirst(["03/04/2024", "04/13/2024"], true)).toBe(false);
  });

  it("falls back when every date is ambiguous", () => {
    expect(detectDayFirst(["03/04/2024"], true)).toBe(true);
    expect(detectDayFirst(["03/04/2024"], false)).toBe(false);
  });
});

describe("profileColumn", () => {
  it("classifies the common column kinds", () => {
    expect(profileColumn("Revenue", ["$1,204.50", "$20", "$3"], 0).kind).toBe("currency");
    expect(profileColumn("Units", ["3", 2, "7"], 0).kind).toBe("number");
    expect(profileColumn("Margin", ["12%", "8.5%", "10%"], 0).kind).toBe("percent");
    expect(profileColumn("Order Date", ["03/04/2024", "2024-04-05", "Apr 6, 2024"], 0).kind).toBe(
      "date"
    );
    expect(profileColumn("Channel", ["Online", "Retail", "Online", "Retail"], 0).kind).toBe(
      "category"
    );
    expect(profileColumn("Order ID", [101, 102, 103], 0).kind).toBe("id");
    expect(profileColumn("Notes", ["a", "b", "c", "d"], 0).kind).toBe("text");
    expect(profileColumn("Empty", [null, "", "N/A"], 0).kind).toBe("empty");
  });

  it("treats a column of years as something to group by", () => {
    expect(profileColumn("Year", [2022, 2023, 2023, 2024], 0).kind).toBe("category");
  });

  it("recognises Excel date serials in a date-named column", () => {
    expect(profileColumn("Order Date", [45390, 45397], 0).kind).toBe("date");
    expect(profileColumn("Quantity", [45390, 45397], 0).kind).toBe("number");
  });

  it("recognises real dates from their number format", () => {
    expect(
      profileColumn("Shipped", [45390, 45397], 0, { formats: ["dd/mm/yyyy", "dd/mm/yyyy"] }).kind
    ).toBe("date");
  });

  it("ignores placeholders when deciding", () => {
    expect(profileColumn("Units", ["3", "N/A", "-", "4"], 0).kind).toBe("number");
  });
});

describe("numberFormatFor", () => {
  const profile = (kind: string, currency: string | null = null) =>
    ({ kind, currency }) as unknown as Parameters<typeof numberFormatFor>[0];

  it("picks a format per kind", () => {
    expect(numberFormatFor(profile("date"), [])).toBe("yyyy-mm-dd");
    expect(numberFormatFor(profile("percent"), [])).toBe("0.0%");
    expect(numberFormatFor(profile("currency", "$"), [])).toBe("$#,##0.00");
    expect(numberFormatFor(profile("currency", "NGN"), [])).toBe('"NGN "#,##0.00');
    expect(numberFormatFor(profile("currency", "₦"), [])).toBe('"₦"#,##0.00');
    expect(numberFormatFor(profile("number"), [1, 2])).toBe("#,##0");
    expect(numberFormatFor(profile("number"), [1.5, 2])).toBe("#,##0.00");
    expect(numberFormatFor(profile("text"), [])).toBe("General");
  });
});

describe("locateTable", () => {
  it("finds the headings under a title and a 'Generated by' line", () => {
    const location = locateTable(MESSY)!;
    expect(location.headerRow).toBe(3);
    expect(location.headerRows).toBe(1);
    expect(location.titleRows).toEqual([0, 1]);
  });

  it("leaves footnotes out of the table", () => {
    const location = locateTable(MESSY)!;
    expect(location.lastRow).toBe(17);
    expect(location.footerRows).toEqual([18]);
  });

  it("spans exactly the used columns", () => {
    const location = locateTable([
      [null, null, null, null],
      [null, "Name", "Amount", null],
      [null, "Ann", 5, null],
    ])!;
    expect(location.firstColumn).toBe(1);
    expect(location.lastColumn).toBe(2);
  });

  it("recognises headings split over two rows", () => {
    const location = locateTable([
      ["Region", "Q1", null, "Q2", null],
      [null, "Units", "Revenue", "Units", "Revenue"],
      ["North", 1, 100, 2, 200],
    ]);
    // "Region" sits over a blank heading, so it is not the merged pattern: the
    // first row is treated as the heading row instead.
    expect(location?.headerRows).toBe(1);

    const merged = locateTable([
      [null, "Q1", null, "Q2", null],
      ["Region", "Units", "Revenue", "Units", "Revenue"],
      ["North", 1, 100, 2, 200],
    ])!;
    expect(merged.headerRow).toBe(1);
    expect(merged.headerRows).toBe(2);
  });

  it("works on an already-clean table", () => {
    const location = locateTable([
      ["Name", "Amount"],
      ["Ann", 5],
    ])!;
    expect(location.headerRow).toBe(0);
    expect(location.titleRows).toEqual([]);
    expect(location.footerRows).toEqual([]);
  });

  it("returns null for an empty grid", () => {
    expect(locateTable([])).toBeNull();
    expect(locateTable([[null, null]])).toBeNull();
  });
});

describe("buildHeaders", () => {
  it("joins a merged upper row onto the headings below it", () => {
    const { headers } = buildHeaders(
      [null, "Q1", null, "Q2", null],
      ["Region", "Units", "Revenue", "Units", "Revenue"]
    );
    expect(headers).toEqual(["Region", "Q1 Units", "Q1 Revenue", "Q2 Units", "Q2 Revenue"]);
  });

  it("names blank headings and makes duplicates unique", () => {
    const { headers, renamed } = buildHeaders(null, ["Amount", null, "Amount"]);
    expect(headers).toEqual(["Amount", "Column B", "Amount (2)"]);
    expect(renamed).toEqual(["Column B", "Amount (2)"]);
  });
});

describe("classifyRow", () => {
  const keys = ["order id", "order date", "customer", "units"];

  it("spots each kind of non-data row", () => {
    expect(classifyRow([null, "", "N/A", null], keys, 4)).toBe("blank");
    expect(classifyRow(["Order ID", "Order Date", "Customer", "Units"], keys, 4)).toBe(
      "repeatedHeader"
    );
    expect(classifyRow(["Europe Total", null, null, 5], keys, 4)).toBe("total");
    expect(classifyRow(["Grand Total", null, null, 17], keys, 4)).toBe("total");
    expect(classifyRow(["Sub-total", null, null, 17], keys, 4)).toBe("total");
    expect(classifyRow(["TOTAL:", null, null, 17], keys, 4)).toBe("total");
    expect(classifyRow(["EUROPE", null, null, null], keys, 4)).toBe("section");
  });

  it("never mistakes a customer called Total for a total row", () => {
    expect(classifyRow(["A-1", "2024-01-01", "Total Foods Ltd", 5], keys, 4)).toBe("data");
  });

  it("does not call a lone number a section", () => {
    expect(classifyRow([null, null, null, 17], keys, 4)).toBe("data");
  });
});

describe("splitSectionLabel", () => {
  it("turns 'Name: value' into a column name and value", () => {
    expect(splitSectionLabel("Region: Europe")).toEqual({ name: "Region", value: "Europe" });
    expect(splitSectionLabel("EUROPE")).toEqual({ name: null, value: "EUROPE" });
  });
});

describe("editDistance", () => {
  it("counts single-letter edits", () => {
    expect(editDistance("gamma", "gama")).toBe(1);
    expect(editDistance("north", "north")).toBe(0);
    expect(editDistance("kitten", "sitting")).toBe(3);
  });

  it("stops early past the limit", () => {
    expect(editDistance("abcdefgh", "zzzzzzzz", 2)).toBe(3);
  });
});

describe("unifySpellings", () => {
  it("folds case, spacing and punctuation variants onto the most common spelling", () => {
    const mapping = unifySpellings([
      "North America",
      "north america",
      "NORTH  AMERICA",
      "North America",
      "Europe",
    ]);
    expect(mapping.get("north america")).toBe("North America");
    expect(mapping.get("NORTH AMERICA")).toBe("North America");
    expect(mapping.has("Europe")).toBe(false);
  });

  it("folds one-letter typos in longer words", () => {
    const mapping = unifySpellings(["Wholesale", "Wholesale", "Wholsale"]);
    expect(mapping.get("Wholsale")).toBe("Wholesale");
  });

  it("keeps short codes that differ by one letter apart", () => {
    expect(unifySpellings(["UK", "US", "UK"]).size).toBe(0);
  });

  it("prefers a properly cased spelling when counts tie", () => {
    expect(unifySpellings(["acme ltd", "Acme Ltd"]).get("acme ltd")).toBe("Acme Ltd");
  });
});

describe("tidyTable on a messy export", () => {
  const result = tidyTable(MESSY)!;
  const col = (name: string) => result.headers.indexOf(name);
  const found = (id: FindingId) => result.findings.find((finding) => finding.id === id);

  it("keeps only real data rows", () => {
    // 7 order rows in the source; one is an exact duplicate.
    expect(result.rows).toHaveLength(6);
  });

  it("moves 'Region: ...' headings into a Region column", () => {
    expect(result.headers[0]).toBe("Region");
    expect(result.rows.map((row) => row[0])).toEqual([
      "Europe",
      "Europe",
      "Europe",
      "Africa",
      "Africa",
      "Africa",
    ]);
  });

  it("turns every revenue figure into a real number", () => {
    expect(result.rows.map((row) => row[col("Revenue")])).toEqual([
      1204.5, 1130, -250, -980, 1200, 45000,
    ]);
  });

  it("empties N/A and converts text units", () => {
    expect(result.rows.map((row) => row[col("Units")])).toEqual([3, 2, null, 4, 1, 7]);
  });

  it("turns five date styles into real dates", () => {
    const dates = result.rows.map((row) => row[col("Order Date")] as number);
    expect(dates.every((value) => typeof value === "number")).toBe(true);
    expect(dates.map((serial) => excelSerialToDate(serial).toISOString().slice(0, 10))).toEqual([
      "2024-04-03",
      "2024-04-05",
      "2024-04-06",
      "2024-04-13",
      "2024-04-15",
      "2024-04-16",
    ]);
  });

  it("unifies customer and channel spellings", () => {
    expect(result.rows.map((row) => row[col("Customer")])).toEqual([
      "Acme Ltd",
      "Acme Ltd",
      "Beta Trading",
      "Gamma Co",
      "Gamma Co",
      "Delta Supplies",
    ]);
    expect(result.rows[2][col("Channel")]).toBe("Online");
  });

  it("gives each column a display format", () => {
    expect(result.numberFormats[col("Order Date")]).toBe("yyyy-mm-dd");
    expect(result.numberFormats[col("Units")]).toBe("#,##0");
  });

  it("reports every kind of problem with a count", () => {
    expect(found("titleRows")?.count).toBe(2);
    expect(found("footerRows")?.count).toBe(1);
    expect(found("totals")?.count).toBe(3);
    expect(found("sections")?.count).toBe(2);
    expect(found("repeatedHeaders")?.count).toBe(1);
    expect(found("duplicates")?.count).toBe(1);
    expect(found("blankRows")?.count).toBe(1);
    expect(found("placeholders")?.count).toBe(2);
    expect(found("textNumbers")?.count).toBeGreaterThanOrEqual(10);
    expect(found("textDates")?.label).toMatch(/formats/);
    expect(found("spellings")?.count).toBeGreaterThanOrEqual(3);
    expect(countProblems(result.findings)).toBeGreaterThan(25);
  });

  it("shows the letter-O repair among the number examples", () => {
    expect(found("textNumbers")?.examples.some((example) => example.includes("1,2O0"))).toBe(true);
  });

  it("lets the user switch individual fixes off", () => {
    const kept = tidyTable(MESSY, {
      skip: new Set<FindingId>(["totals", "duplicates", "spellings"]),
    })!;
    expect(kept.rows.length).toBe(6 + 1 + 3);
    expect(kept.rows.some((row) => row.includes("ACME LTD"))).toBe(true);
    // Findings are still reported even when their fix is off.
    expect(kept.findings.find((finding) => finding.id === "totals")?.count).toBe(3);
  });

  it("keeps group headings as rows when that fix is off", () => {
    const kept = tidyTable(MESSY, { skip: new Set<FindingId>(["sections"]) })!;
    expect(kept.headers[0]).toBe("Order ID");
  });
});

describe("tidyTable on clean data", () => {
  it("finds nothing to fix and changes nothing", () => {
    const clean: Grid = [
      ["Region", "Amount"],
      ["North", 100],
      ["South", 200],
    ];
    const result = tidyTable(clean)!;
    expect(result.findings).toEqual([]);
    expect(result.rows).toEqual([
      ["North", 100],
      ["South", 200],
    ]);
  });
});

describe("regressions found on a real-sized export", () => {
  it("never treats codes as typos of each other", () => {
    expect(unifySpellings(["SO-5001", "SO-5008", "SO-5008"]).size).toBe(0);
    expect(unifySpellings(["SKU 12", "SKU 13"]).size).toBe(0);
  });

  it("keeps an ID column as IDs even when duplicate rows repeat some values", () => {
    const ids = Array.from({ length: 20 }, (_v, i) => `SO-${5000 + i}`);
    ids.push("SO-5003");
    expect(profileColumn("Order ID", ids, 0).kind).toBe("id");
  });

  it("leaves ID values untouched when tidying", () => {
    const grid: Grid = [
      ["Order ID", "Region", "Revenue"],
      ...Array.from({ length: 12 }, (_v, i) => [
        `SO-${5001 + i}`,
        i % 2 ? "North" : "South",
        10 * i,
      ]),
    ];
    const result = tidyTable(grid)!;
    expect(result.rows.map((row) => row[0])).toEqual(grid.slice(1).map((row) => row[0]));
  });

  it("calls a column money when only some cells carry the symbol", () => {
    expect(profileColumn("Revenue", ["$10.00", "12.50", "7.25", "8.00", "9.10"], 0).kind).toBe(
      "currency"
    );
  });

  it("counts date formats by shape, not by digit count", () => {
    const result = tidyTable([
      ["Order Date", "Units"],
      ["Jan 5, 2024", 1],
      ["Jan 11, 2024", 2],
      ["2024-01-05", 3],
    ])!;
    expect(result.findings.find((finding) => finding.id === "textDates")?.label).toBe(
      "Dates typed as text, in 2 different formats"
    );
  });
});
