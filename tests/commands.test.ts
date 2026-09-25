import {
  COMMANDS,
  completeCommand,
  findSheets,
  Intent,
  matchHeader,
  parseCommand,
  skipTargets,
  suggestCommands,
} from "../src/taskpane/features/commands";

const SHEETS = ["Jan Orders", "Feb Orders", "Mar Orders", "Sales Extract", "Sales Extract (clean)"];
const parse = (text: string): Intent => parseCommand(text, { sheets: SHEETS });

describe("slash commands", () => {
  it("maps each command to its action", () => {
    expect(parse("/help").kind).toBe("help");
    expect(parse("/fix")).toEqual({ kind: "fix", sheet: null });
    expect(parse("/dashboard")).toEqual({ kind: "dashboard", sheet: null, max: null });
    expect(parse("/refresh")).toEqual({ kind: "refresh", name: null });
    expect(parse("/duplicates").kind).toBe("duplicates");
    expect(parse("/spaces").kind).toBe("spaces");
    expect(parse("/fill").kind).toBe("fill");
    expect(parse("/tools")).toEqual({ kind: "open", tool: "home" });
    expect(parse("/formulas")).toEqual({ kind: "open", tool: "formulas" });
  });

  it("is case-insensitive and accepts aliases", () => {
    expect(parse("/FIX").kind).toBe("fix");
    expect(parse("/tidy").kind).toBe("fix");
    expect(parse("/dedupe").kind).toBe("duplicates");
    expect(parse("/trim").kind).toBe("spaces");
  });

  it("picks up a sheet named after the command", () => {
    expect(parse("/fix Sales Extract")).toEqual({ kind: "fix", sheet: "Sales Extract" });
    expect(parse("/dashboard on sales extract (clean)")).toEqual({
      kind: "dashboard",
      sheet: "Sales Extract (clean)",
      max: null,
    });
  });

  it("reads date and case options", () => {
    expect(parse("/dates day-first")).toEqual({ kind: "dates", dayFirst: true });
    expect(parse("/dates US")).toEqual({ kind: "dates", dayFirst: false });
    expect(parse("/dates")).toEqual({ kind: "dates", dayFirst: null });
    expect(parse("/case title")).toEqual({ kind: "case", mode: "proper" });
    expect(parse("/case upper")).toEqual({ kind: "case", mode: "upper" });
    expect(parse("/case")).toEqual({ kind: "case", mode: null });
  });

  it("names the sheets to combine, in the order given", () => {
    expect(parse("/combine Feb Orders, Jan Orders")).toEqual({
      kind: "combine",
      sheets: ["Feb Orders", "Jan Orders"],
      all: false,
    });
    expect(parse("/combine all")).toEqual({ kind: "combine", sheets: [], all: true });
  });

  it("suggests the nearest command for a typo", () => {
    const result = parse("/dashbord");
    expect(result.kind).toBe("unknown");
    expect(result.kind === "unknown" && result.suggestions[0]).toBe("/dashboard");
  });
});

describe("charts in plain words", () => {
  it("reads aggregation, measure, dimension and chart type", () => {
    expect(parse("/chart sum of Revenue by Region as pie")).toEqual({
      kind: "chart",
      add: false,
      summaryOnly: false,
      aggregation: "sum",
      measure: "revenue",
      dimension: "region",
      chartType: "pie",
      sheet: null,
    });
  });

  it("understands everyday phrasings", () => {
    const pie = parse("make a pie of revenue by region");
    expect(pie).toMatchObject({
      kind: "chart",
      chartType: "pie",
      measure: "revenue",
      dimension: "region",
    });

    const average = parse("average units per channel as a bar chart");
    expect(average).toMatchObject({
      kind: "chart",
      aggregation: "average",
      measure: "units",
      dimension: "channel",
      chartType: "barClustered",
    });

    const count = parse("/chart count by region");
    expect(count).toMatchObject({ aggregation: "count", measure: null, dimension: "region" });
  });

  it("builds summary tables without a chart", () => {
    expect(parse("/summary total revenue by product")).toMatchObject({
      kind: "chart",
      summaryOnly: true,
      measure: "revenue",
      dimension: "product",
    });
  });

  it("notes a source sheet", () => {
    expect(parse("/chart revenue by region on Sales Extract")).toMatchObject({
      sheet: "Sales Extract",
      dimension: "region",
    });
  });
});

describe("plain-language requests", () => {
  it.each([
    ["fix this sheet", "fix"],
    ["can you tidy up this mess", "fix"],
    ["build a dashboard", "dashboard"],
    ["refresh my dashboard", "refresh"],
    ["remove duplicates", "duplicates"],
    ["get rid of extra spaces", "spaces"],
    ["fix the dates", "dates"],
    ["make the names title case", "case"],
    ["combine jan orders and feb orders", "combine"],
    ["fill the formula down", "fill"],
    ["open the formulas", "open"],
    ["hello", "hello"],
    ["what can you do", "help"],
  ])("%s -> %s", (text, kind) => {
    expect(parse(text).kind).toBe(kind);
  });

  it("finds the sheet in a sentence", () => {
    expect(parse("please fix sales extract")).toEqual({ kind: "fix", sheet: "Sales Extract" });
  });

  it("says it doesn't understand rather than guessing", () => {
    expect(parse("banana").kind).toBe("unknown");
    expect(parse("").kind).toBe("unknown");
  });
});

describe("replies to a question", () => {
  it("recognises yes and no", () => {
    for (const text of ["yes", "ok", "go ahead", "Fix it", "build", "build all"])
      expect(parse(text).kind).toBe("confirm");
    for (const text of ["no", "cancel", "never mind"]) expect(parse(text).kind).toBe("cancel");
  });

  it("reads which fixes to skip", () => {
    expect(parse("skip totals and spellings")).toEqual({
      kind: "skip",
      words: ["totals", "spellings"],
    });
    expect(parse("keep duplicates")).toEqual({ kind: "skip", words: ["duplicates"] });
  });

  it("reads which charts to build", () => {
    expect(parse("build 1, 2 and 5")).toEqual({ kind: "pick", numbers: [1, 2, 5], exclude: false });
    expect(parse("without 7 and 8")).toEqual({ kind: "pick", numbers: [7, 8], exclude: true });
    expect(parse("only 4")).toEqual({ kind: "pick", numbers: [1, 2, 3, 4], exclude: false });
    expect(parse("just three charts")).toEqual({
      kind: "pick",
      numbers: [1, 2, 3],
      exclude: false,
    });
  });
});

describe("helpers", () => {
  it("prefers the longest matching sheet name", () => {
    expect(findSheets("use sales extract (clean) please", SHEETS)).toEqual([
      "Sales Extract (clean)",
    ]);
    expect(findSheets("jan orders then mar orders", SHEETS)).toEqual(["Jan Orders", "Mar Orders"]);
  });

  it("matches typed column names to headings", () => {
    const headers = ["Order ID", "Order Date", "Region", "Revenue", "Unit Price", "Units"];
    expect(matchHeader("revenue", headers)).toBe("Revenue");
    expect(matchHeader("regions", headers)).toBe("Region");
    expect(matchHeader("the region column", headers)).toBe("Region");
    expect(matchHeader("revnue", headers)).toBe("Revenue");
    expect(matchHeader("price", headers)).toBe("Unit Price");
    expect(matchHeader("banana", headers)).toBeNull();
    expect(matchHeader(null, headers)).toBeNull();
  });

  it("autocompletes slash commands", () => {
    expect(completeCommand("/d").map((c) => c.command)).toEqual([
      "/dashboard",
      "/duplicates",
      "/dates",
    ]);
    expect(completeCommand("/fix Sales")).toEqual([]);
    expect(completeCommand("fix")).toEqual([]);
    expect(completeCommand("/")).toHaveLength(COMMANDS.length);
  });

  it("maps skip words to fixes", () => {
    expect(skipTargets(["totals", "spellings"]).sort()).toEqual(["spellings", "totals"]);
    expect(skipTargets(["group headings"])).toEqual(["sections"]);
  });

  it("suggests commands for near-misses", () => {
    expect(suggestCommands("dashbord")[0]).toBe("dashboard".replace(/^/, "/"));
  });
});

describe("review commands", () => {
  it("maps the review family", () => {
    expect(parse("/review")).toEqual({ kind: "review", sheet: null });
    expect(parse("/review Sales Extract")).toEqual({ kind: "review", sheet: "Sales Extract" });
    expect(parse("/check").kind).toBe("review");
    expect(parse("/undo").kind).toBe("undo");
    expect(parse("/marks clear").kind).toBe("clearMarks");
    expect(parse("/unignore").kind).toBe("unignore");
    expect(parse("/show 2")).toEqual({ kind: "show", numbers: [2] });
    expect(parse("/ignore 3")).toEqual({ kind: "ignore", numbers: [3] });
  });

  it("understands the everyday phrasings", () => {
    expect(parse("check this sheet for mistakes").kind).toBe("review");
    expect(parse("what's wrong with this sheet?").kind).toBe("review");
    expect(parse("any errors here").kind).toBe("review");
    expect(parse("undo that").kind).toBe("undo");
    expect(parse("clear the marks").kind).toBe("clearMarks");
    expect(parse("show me 2")).toEqual({ kind: "show", numbers: [2] });
    expect(parse("ignore 4")).toEqual({ kind: "ignore", numbers: [4] });
  });

  it("reads which numbered items to fix", () => {
    expect(parse("fix 2 and 3")).toEqual({ kind: "pick", numbers: [2, 3], exclude: false });
  });

  it("keeps 'fix this sheet' separate from 'fix 2'", () => {
    expect(parse("fix this sheet").kind).toBe("fix");
  });
});

describe("shaping a dashboard", () => {
  it("reads a chart limit", () => {
    expect(parse("/dashboard 5")).toMatchObject({ kind: "dashboard", max: 5 });
    expect(parse("build a dashboard with 4 charts")).toMatchObject({ kind: "dashboard", max: 4 });
  });

  it("reads an added chart", () => {
    expect(parse("add a pie of Units by Channel")).toMatchObject({
      kind: "chart",
      add: true,
      chartType: "pie",
      measure: "units",
      dimension: "channel",
    });
    expect(parse("/add line of Revenue by Order Date")).toMatchObject({
      add: true,
      chartType: "lineMarkers",
    });
    expect(parse("/chart sum of Revenue by Region")).toMatchObject({ add: false });
  });
});
