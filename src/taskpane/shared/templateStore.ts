/**
 * Persistence for report templates.
 *
 * Templates live in `context.workbook.settings`, which Excel saves *inside the
 * workbook file*. That means a template travels with the .xlsx: mail the file to
 * a colleague and their copy of the add-in sees the same template, with no
 * server, account or external storage involved.
 */

import { fail, ok, runExcel } from "./excelHelpers";
import { DEFAULT_REFRESH_OPTIONS, OperationResult, ReportTemplate } from "./types";

/** Namespacing keeps our settings from colliding with other add-ins' settings. */
const KEY_PREFIX = "MExAutomate.template.";

/** Bumped if the persisted shape ever changes, so old workbooks can be migrated. */
const SCHEMA_VERSION = 1;

interface StoredTemplate {
  schemaVersion: number;
  template: ReportTemplate;
}

function keyFor(name: string): string {
  return `${KEY_PREFIX}${name}`;
}

export function nowIso(): string {
  return new Date().toISOString();
}

/**
 * Rebuild a template from persisted JSON, tolerating anything missing or
 * malformed. Returns null when the payload is too damaged to use.
 */
export function parseTemplate(raw: unknown): ReportTemplate | null {
  let payload: unknown = raw;
  if (typeof raw === "string") {
    try {
      payload = JSON.parse(raw);
    } catch {
      return null;
    }
  }
  if (!payload || typeof payload !== "object") {
    return null;
  }

  const stored = payload as Partial<StoredTemplate> & Partial<ReportTemplate>;
  // Accept both the wrapped shape and a bare template written by an older build.
  const template = (stored.template ?? stored) as Partial<ReportTemplate>;
  if (!template || typeof template.name !== "string" || template.name.trim() === "") {
    return null;
  }

  return {
    name: template.name,
    createdAt: typeof template.createdAt === "string" ? template.createdAt : nowIso(),
    updatedAt: typeof template.updatedAt === "string" ? template.updatedAt : nowIso(),
    zones: Array.isArray(template.zones) ? template.zones : [],
    mappings: template.mappings && typeof template.mappings === "object" ? template.mappings : {},
    options: { ...DEFAULT_REFRESH_OPTIONS, ...(template.options ?? {}) },
    sourceSheetName:
      typeof template.sourceSheetName === "string" && template.sourceSheetName !== ""
        ? template.sourceSheetName
        : undefined,
  };
}

function serialize(template: ReportTemplate): string {
  const stored: StoredTemplate = { schemaVersion: SCHEMA_VERSION, template };
  return JSON.stringify(stored);
}

/** Create or overwrite a template. `settings.add` upserts by key. */
export async function saveTemplate(template: ReportTemplate): Promise<OperationResult> {
  const name = template.name.trim();
  if (name === "") {
    return fail("Give the template a name before saving it.");
  }

  return runExcel(async (context) => {
    const existing = context.workbook.settings.getItemOrNullObject(keyFor(name));
    existing.load("value");
    await context.sync();

    const previous = existing.isNullObject ? null : parseTemplate(existing.value);
    const record: ReportTemplate = {
      ...template,
      name,
      createdAt: previous?.createdAt ?? template.createdAt ?? nowIso(),
      updatedAt: nowIso(),
    };

    context.workbook.settings.add(keyFor(name), serialize(record));
    await context.sync();

    return ok(
      previous ? `Updated template "${name}".` : `Saved template "${name}" into this workbook.`
    );
  });
}

/** Every template stored in the current workbook, sorted by name. */
export async function loadTemplates(): Promise<ReportTemplate[]> {
  try {
    return await Excel.run(async (context) => {
      const settings = context.workbook.settings;
      settings.load("items/key,items/value");
      await context.sync();

      const templates: ReportTemplate[] = [];
      for (const setting of settings.items) {
        if (!setting.key.startsWith(KEY_PREFIX)) {
          continue;
        }
        const parsed = parseTemplate(setting.value);
        if (parsed) {
          templates.push(parsed);
        }
      }
      return templates.sort((a, b) => a.name.localeCompare(b.name));
    });
  } catch {
    // A missing/locked workbook should surface as "no templates", not a crash.
    return [];
  }
}

export async function getTemplate(name: string): Promise<ReportTemplate | null> {
  const templates = await loadTemplates();
  return templates.find((template) => template.name === name) ?? null;
}

export async function deleteTemplate(name: string): Promise<OperationResult> {
  return runExcel(async (context) => {
    const setting = context.workbook.settings.getItemOrNullObject(keyFor(name));
    await context.sync();

    if (setting.isNullObject) {
      return fail(`No template named "${name}" is stored in this workbook.`);
    }

    setting.delete();
    await context.sync();
    return ok(`Deleted template "${name}".`);
  });
}
