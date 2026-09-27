/**
 * Where recipes live: the workbook's own settings.
 *
 * Excel saves these inside the .xlsx, so a recipe travels with the file. Send
 * next month's workbook to a colleague and their copy of MEx can run the same
 * routine, with no server, account or shared drive in it anywhere.
 */

import { fail, ok, runExcel } from "../shared/excelHelpers";
import type { OperationResult } from "../shared/types";
import { MAX_STEPS, nowIso, parseRecipe, Recipe, RecipeStep, serializeRecipe } from "./recipes";

const KEY_PREFIX = "MExAutomate.recipe.";

function keyFor(name: string): string {
  return `${KEY_PREFIX}${name}`;
}

/** Every recipe stored in this workbook, most recently updated first. */
export async function loadRecipes(): Promise<Recipe[]> {
  try {
    return await Excel.run(async (context) => {
      const settings = context.workbook.settings;
      settings.load("items/key,items/value");
      await context.sync();

      const recipes: Recipe[] = [];
      for (const setting of settings.items) {
        if (!setting.key.startsWith(KEY_PREFIX)) continue;
        const parsed = parseRecipe(setting.value);
        if (parsed) recipes.push(parsed);
      }
      return recipes.sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
    });
  } catch {
    // A locked or closing workbook should read as "no recipes", not a crash.
    return [];
  }
}

export async function getRecipe(name: string): Promise<Recipe | null> {
  const recipes = await loadRecipes();
  const wanted = name.trim().toLowerCase();
  return (
    recipes.find((recipe) => recipe.name.toLowerCase() === wanted) ??
    // A partial name is enough when it only matches one recipe.
    recipes.filter((recipe) => recipe.name.toLowerCase().includes(wanted))[0] ??
    null
  );
}

export async function saveRecipe(
  name: string,
  steps: readonly RecipeStep[]
): Promise<OperationResult> {
  const trimmed = name.trim();
  if (trimmed === "") return fail("Give the recipe a name, for example /recipe save month end.");
  if (steps.length === 0) {
    return fail("There is nothing to save yet - run a command or two first, then save.");
  }

  return runExcel(async (context) => {
    const existing = context.workbook.settings.getItemOrNullObject(keyFor(trimmed));
    existing.load("value");
    await context.sync();

    const previous = existing.isNullObject ? null : parseRecipe(existing.value);
    const recipe: Recipe = {
      name: trimmed,
      createdAt: previous?.createdAt ?? nowIso(),
      updatedAt: nowIso(),
      steps: steps.slice(-MAX_STEPS).map((step) => ({ ...step })),
    };

    context.workbook.settings.add(keyFor(trimmed), serializeRecipe(recipe));
    await context.sync();

    return ok(
      previous
        ? `Updated "${trimmed}" - ${recipe.steps.length} steps.`
        : `Saved "${trimmed}" into this workbook - ${recipe.steps.length} steps.`,
      ["It travels with the file, so next month you can just run it."]
    );
  });
}

export async function deleteRecipe(name: string): Promise<OperationResult> {
  const existing = await getRecipe(name);
  if (!existing) return fail(`I have no recipe called "${name}" in this workbook.`);

  return runExcel(async (context) => {
    const setting = context.workbook.settings.getItemOrNullObject(keyFor(existing.name));
    await context.sync();
    if (setting.isNullObject) return fail(`I have no recipe called "${existing.name}".`);
    setting.delete();
    await context.sync();
    return ok(`Deleted "${existing.name}".`);
  });
}
