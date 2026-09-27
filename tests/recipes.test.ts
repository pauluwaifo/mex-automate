import {
  addStep,
  checkRecipe,
  describeRecipe,
  emptyRecipe,
  isRepeatableIntent,
  MAX_STEPS,
  parseRecipe,
  removeStep,
  serializeRecipe,
  sheetsUsedBy,
  summarizeRun,
} from "../src/taskpane/features/recipes";
import { parseCommand } from "../src/taskpane/features/commands";
import type { Recipe, RecipeStep } from "../src/taskpane/features/recipes";

const step = (typed: string, label = typed, sheets: string[] = []): RecipeStep => ({
  typed,
  label,
  sheets,
});

describe("what is worth recording", () => {
  const parse = (text: string) => parseCommand(text, { sheets: ["Jan Orders", "Feb Orders"] });

  it("records the commands that do work", () => {
    for (const text of [
      "/fix",
      "/dashboard",
      "/combine Jan Orders, Feb Orders",
      "/unpivot",
      "/duplicates",
    ]) {
      expect(isRepeatableIntent(parse(text))).toBe(true);
    }
  });

  it("does not record conversation, navigation or corrections", () => {
    for (const text of [
      "/help",
      "hello",
      "/undo",
      "show me 2",
      "/ignore 1",
      "/tools",
      "nonsense-xyz",
    ]) {
      expect(isRepeatableIntent(parse(text))).toBe(false);
    }
  });

  it("notes the sheets a step depends on", () => {
    expect(sheetsUsedBy(parse("/combine Jan Orders, Feb Orders"))).toEqual([
      "Jan Orders",
      "Feb Orders",
    ]);
    expect(sheetsUsedBy(parse("/dashboard on Jan Orders"))).toEqual(["Jan Orders"]);
    expect(sheetsUsedBy(parse("/duplicates"))).toEqual([]);
  });
});

describe("building a recipe", () => {
  it("starts empty and collects steps in order", () => {
    let recipe = emptyRecipe("Month end");
    recipe = addStep(recipe, step("/fix", "Clean up"));
    recipe = addStep(recipe, step("/dashboard", "Build a dashboard"));
    expect(recipe.steps.map((item) => item.label)).toEqual(["Clean up", "Build a dashboard"]);
  });

  it("does not record the same command twice in a row", () => {
    let recipe = emptyRecipe("Month end");
    recipe = addStep(recipe, step("/fix"));
    recipe = addStep(recipe, step("  /FIX  "));
    expect(recipe.steps).toHaveLength(1);
  });

  it("does record the same command again later in the process", () => {
    let recipe = emptyRecipe("Month end");
    recipe = addStep(recipe, step("/fix"));
    recipe = addStep(recipe, step("/combine all"));
    recipe = addStep(recipe, step("/fix"));
    expect(recipe.steps).toHaveLength(3);
  });

  it("keeps the most recent steps once it is full", () => {
    let recipe = emptyRecipe("Long");
    for (let at = 0; at < MAX_STEPS + 5; at += 1) recipe = addStep(recipe, step(`/chart ${at}`));
    expect(recipe.steps).toHaveLength(MAX_STEPS);
    expect(recipe.steps[MAX_STEPS - 1].typed).toBe(`/chart ${MAX_STEPS + 4}`);
  });

  it("can drop a step", () => {
    let recipe = emptyRecipe("Month end");
    recipe = addStep(recipe, step("/fix"));
    recipe = addStep(recipe, step("/dashboard"));
    expect(removeStep(recipe, 0).steps.map((item) => item.typed)).toEqual(["/dashboard"]);
    expect(removeStep(recipe, 9).steps).toHaveLength(2);
  });
});

describe("checkRecipe", () => {
  const recipe: Recipe = {
    ...emptyRecipe("Month end"),
    steps: [
      step("/fix", "Clean up"),
      step("/combine Jan Orders, Feb Orders", "Combine 2 sheets", ["Jan Orders", "Feb Orders"]),
    ],
  };

  it("passes every step when the sheets are all there", () => {
    const check = checkRecipe(recipe, ["Jan Orders", "Feb Orders", "Sheet1"]);
    expect(check.runnable).toBe(2);
    expect(check.blocked).toBe(0);
  });

  it("names the sheets this workbook has not got instead of guessing", () => {
    const check = checkRecipe(recipe, ["Mar Orders", "Apr Orders"]);
    expect(check.blocked).toBe(1);
    expect(check.steps[1].missingSheets).toEqual(["Jan Orders", "Feb Orders"]);
  });

  it("ignores case and stray spaces in sheet names", () => {
    const check = checkRecipe(recipe, ["jan orders ", "FEB ORDERS"]);
    expect(check.runnable).toBe(2);
  });
});

describe("describing and summarizing", () => {
  it("describes a recipe in one line", () => {
    const recipe: Recipe = {
      ...emptyRecipe("Month end"),
      steps: [step("/fix", "Clean up"), step("/dashboard", "Build a dashboard")],
    };
    expect(describeRecipe(recipe)).toBe('"Month end" - 2 steps: clean up, then build a dashboard.');
  });

  it("says when a recipe is still empty", () => {
    expect(describeRecipe(emptyRecipe("New"))).toBe('"New" has no steps yet.');
  });

  it("reports a clean run plainly", () => {
    const outcomes = [
      { step: step("/fix"), ok: true, message: "done", skipped: false },
      { step: step("/dashboard"), ok: true, message: "done", skipped: false },
    ];
    expect(summarizeRun("Month end", outcomes)).toBe('Ran "Month end" - all 2 steps done.');
  });

  it("does not hide the steps that did not run", () => {
    const outcomes = [
      { step: step("/fix"), ok: true, message: "done", skipped: false },
      { step: step("/combine X"), ok: false, message: "no sheet", skipped: true },
      { step: step("/dashboard"), ok: false, message: "failed", skipped: false },
    ];
    expect(summarizeRun("Month end", outcomes)).toBe(
      'Ran "Month end" - 1 of 3 steps done, 1 skipped, 1 failed.'
    );
  });
});

describe("storage", () => {
  it("survives a round trip", () => {
    const recipe: Recipe = {
      ...emptyRecipe("Month end"),
      steps: [step("/combine all", "Combine every sheet", [])],
    };
    const restored = parseRecipe(serializeRecipe(recipe));
    expect(restored?.name).toBe("Month end");
    expect(restored?.steps).toEqual(recipe.steps);
  });

  it("reads a bare recipe written by an older build", () => {
    const restored = parseRecipe(JSON.stringify({ name: "Old", steps: [{ typed: "/fix" }] }));
    expect(restored?.steps[0]).toEqual({ typed: "/fix", label: "/fix", sheets: [] });
  });

  it("refuses what it cannot use, rather than half-loading it", () => {
    expect(parseRecipe("not json")).toBeNull();
    expect(parseRecipe(JSON.stringify({ steps: [] }))).toBeNull();
    expect(parseRecipe(JSON.stringify({ name: "  " }))).toBeNull();
    expect(parseRecipe(null)).toBeNull();
  });

  it("drops damaged steps but keeps the rest", () => {
    const restored = parseRecipe(
      JSON.stringify({ name: "Mixed", steps: [{ typed: "/fix" }, { label: "no command" }, null] })
    );
    expect(restored?.steps).toHaveLength(1);
  });
});
