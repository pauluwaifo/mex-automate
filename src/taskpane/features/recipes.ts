/**
 * Recipes: doing this month's work again next month.
 *
 * The expensive part of a monthly report is never the thinking, it is the
 * repetition - the same clean-up, the same combine, the same dashboard, in the
 * same order, from a file with the same shape and different numbers. Power Query
 * solves this with saved steps and Excel solves it with macros, and both ask you
 * to learn a second tool before you get the benefit.
 *
 * A recipe is the same idea made out of what the user already did: MEx remembers
 * the commands that worked, and replays them.
 *
 * What gets stored is the command as it was typed, not the result of running it.
 * Replaying therefore re-parses the command against the workbook that is open
 * now, which is what makes a recipe useful on next month's file: the sheet is
 * longer, the numbers are different, and the command still means the same thing.
 * The alternative - remembering "clean A1:F250" - would quietly clean the wrong
 * range as soon as a row was added.
 *
 * Recipes live in the workbook's own settings, so they travel inside the .xlsx.
 *
 * Everything here is pure and unit tested in tests/recipes.test.ts.
 */

import type { Intent } from "./commands";
import { plural, uniqueName } from "../shared/excelHelpers";

export interface RecipeStep {
  /** The command exactly as it was typed, e.g. "/combine Jan Orders, Feb Orders". */
  typed: string;
  /** What it did, in the assistant's words, for the list in the pane. */
  label: string;
  /** Sheet names the command named, so a replay can check they still exist. */
  sheets: string[];
}

export interface Recipe {
  name: string;
  createdAt: string;
  updatedAt: string;
  steps: RecipeStep[];
}

export const MAX_STEPS = 30;

/**
 * Intents worth remembering. Everything else is either conversation ("help"),
 * navigation ("show me number 3"), or a correction of something that just
 * happened ("undo") - replaying any of those next month would be noise at best.
 */
const REPEATABLE = new Set<Intent["kind"]>([
  "fix",
  "review",
  "dashboard",
  "refresh",
  "duplicates",
  "spaces",
  "dates",
  "case",
  "combine",
  "chart",
  "unpivot",
  "compare",
  "fill",
]);

export function isRepeatableIntent(intent: Intent): boolean {
  return REPEATABLE.has(intent.kind);
}

/** Sheet names an intent refers to, so a replay can say what is missing. */
export function sheetsUsedBy(intent: Intent): string[] {
  const named: (string | null | undefined)[] = [];
  switch (intent.kind) {
    case "fix":
    case "review":
    case "dashboard":
    case "chart":
      named.push(intent.sheet);
      break;
    case "combine":
      named.push(...intent.sheets);
      break;
    case "unpivot":
      named.push(intent.sheet);
      break;
    case "compare":
      named.push(intent.first, intent.second);
      break;
    default:
      break;
  }
  return named.filter((name): name is string => typeof name === "string" && name !== "");
}

export function nowIso(): string {
  return new Date().toISOString();
}

export function emptyRecipe(name: string): Recipe {
  const created = nowIso();
  return { name, createdAt: created, updatedAt: created, steps: [] };
}

/**
 * Adds a step, unless it repeats the one before it. Running the same clean-up
 * twice in a session is usually a person checking it worked, not two steps of a
 * process.
 */
export function addStep(recipe: Recipe, step: RecipeStep): Recipe {
  const previous = recipe.steps[recipe.steps.length - 1];
  if (previous && previous.typed.trim().toLowerCase() === step.typed.trim().toLowerCase()) {
    return recipe;
  }
  const steps = [...recipe.steps, step];
  return {
    ...recipe,
    // A recipe is a process, not a transcript; past the limit the oldest steps
    // are the least likely to still matter.
    steps: steps.slice(-MAX_STEPS),
    updatedAt: nowIso(),
  };
}

export function removeStep(recipe: Recipe, index: number): Recipe {
  if (index < 0 || index >= recipe.steps.length) return recipe;
  return {
    ...recipe,
    steps: recipe.steps.filter((_, at) => at !== index),
    updatedAt: nowIso(),
  };
}

/** A name that isn't taken yet, for "save this as a recipe" with no name given. */
export function suggestRecipeName(taken: readonly string[]): string {
  const month = new Date().toLocaleString("en-US", { month: "long" });
  return uniqueName(`${month} routine`, taken);
}

// ---------------------------------------------------------------------------
// Checking a recipe against the workbook that is open now
// ---------------------------------------------------------------------------

export interface StepCheck {
  step: RecipeStep;
  /** Sheets the step names that this workbook does not have. */
  missingSheets: string[];
  runnable: boolean;
}

export interface RecipeCheck {
  steps: StepCheck[];
  runnable: number;
  blocked: number;
}

/**
 * Works out which steps can run here. A step that names a sheet this workbook
 * hasn't got is reported rather than guessed at: "Feb Orders" in last month's
 * recipe is not this month's "Mar Orders", and quietly substituting one for the
 * other would produce a report that looks right and isn't.
 */
export function checkRecipe(recipe: Recipe, sheets: readonly string[]): RecipeCheck {
  const present = new Set(sheets.map((name) => name.trim().toLowerCase()));
  const steps = recipe.steps.map((step) => {
    const missingSheets = step.sheets.filter((name) => !present.has(name.trim().toLowerCase()));
    return { step, missingSheets, runnable: missingSheets.length === 0 };
  });
  return {
    steps,
    runnable: steps.filter((step) => step.runnable).length,
    blocked: steps.filter((step) => !step.runnable).length,
  };
}

/** One line for the chat: "4 steps: clean up, combine 3 sheets, dashboard". */
export function describeRecipe(recipe: Recipe): string {
  if (recipe.steps.length === 0) return `"${recipe.name}" has no steps yet.`;
  const labels = recipe.steps.map((step) => step.label.toLowerCase());
  return `"${recipe.name}" - ${plural(recipe.steps.length, "step")}: ${labels.join(", then ")}.`;
}

// ---------------------------------------------------------------------------
// Results of a replay
// ---------------------------------------------------------------------------

export interface StepOutcome {
  step: RecipeStep;
  ok: boolean;
  message: string;
  /** True when the step was not attempted, e.g. a sheet it needs is missing. */
  skipped: boolean;
}

export function summarizeRun(name: string, outcomes: readonly StepOutcome[]): string {
  const done = outcomes.filter((outcome) => outcome.ok).length;
  const skipped = outcomes.filter((outcome) => outcome.skipped).length;
  const failed = outcomes.filter((outcome) => !outcome.ok && !outcome.skipped).length;

  if (outcomes.length === 0) return `"${name}" has no steps to run.`;
  if (failed === 0 && skipped === 0) {
    return `Ran "${name}" - all ${plural(done, "step")} done.`;
  }

  const parts = [`${done} of ${outcomes.length} steps done`];
  if (skipped > 0) parts.push(`${skipped} skipped`);
  if (failed > 0) parts.push(`${failed} failed`);
  return `Ran "${name}" - ${parts.join(", ")}.`;
}

// ---------------------------------------------------------------------------
// Storage shape
// ---------------------------------------------------------------------------

const SCHEMA_VERSION = 1;

interface StoredRecipe {
  schemaVersion: number;
  recipe: Recipe;
}

export function serializeRecipe(recipe: Recipe): string {
  const stored: StoredRecipe = { schemaVersion: SCHEMA_VERSION, recipe };
  return JSON.stringify(stored);
}

/** Rebuilds a recipe from stored JSON, tolerating anything missing or damaged. */
export function parseRecipe(raw: unknown): Recipe | null {
  let payload: unknown = raw;
  if (typeof raw === "string") {
    try {
      payload = JSON.parse(raw);
    } catch {
      return null;
    }
  }
  if (!payload || typeof payload !== "object") return null;

  const stored = payload as Partial<StoredRecipe> & Partial<Recipe>;
  const recipe = (stored.recipe ?? stored) as Partial<Recipe>;
  if (typeof recipe.name !== "string" || recipe.name.trim() === "") return null;

  const steps: RecipeStep[] = Array.isArray(recipe.steps)
    ? recipe.steps
        .map((step) => {
          if (!step || typeof step !== "object") return null;
          const candidate = step as Partial<RecipeStep>;
          if (typeof candidate.typed !== "string" || candidate.typed.trim() === "") return null;
          return {
            typed: candidate.typed,
            label: typeof candidate.label === "string" ? candidate.label : candidate.typed,
            sheets: Array.isArray(candidate.sheets)
              ? candidate.sheets.filter((name): name is string => typeof name === "string")
              : [],
          };
        })
        .filter((step): step is RecipeStep => step !== null)
    : [];

  return {
    name: recipe.name.trim(),
    createdAt: typeof recipe.createdAt === "string" ? recipe.createdAt : nowIso(),
    updatedAt: typeof recipe.updatedAt === "string" ? recipe.updatedAt : nowIso(),
    steps: steps.slice(0, MAX_STEPS),
  };
}
