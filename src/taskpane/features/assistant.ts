/**
 * The assistant: turns each message into an action using the same tested
 * features as the tool screens, and answers with structured messages the chat
 * renders (findings, chart plans, results) plus a few suggested replies.
 *
 * Anything that changes the workbook in a big way - fixing a sheet, building
 * a dashboard - is previewed first and waits for a "yes". Small, targeted
 * clean-ups run straight away, on the cells shown in the reply.
 */

import { loadSheetNames, plural } from "../shared/excelHelpers";
import { OperationResult } from "../shared/types";
import {
  aggregationLabel,
  createChart,
  readSourceHeaders,
  SummarizeRequest,
  summarizeToNewSheet,
} from "./charts";
import { COMMANDS, Intent, matchHeader, parseCommand, skipTargets, ToolName } from "./commands";
import {
  analyzeDashboard,
  buildDashboard,
  customChart,
  DashChart,
  DashChartKind,
  listDashboards,
  refreshDashboard,
} from "./dashboard";
import {
  removeDuplicateRows,
  standardizeDates,
  standardizeTextCase,
  trimWhitespace,
} from "./dataCleaning";
import { fillFormulaAcrossSelection } from "./formulaPatterns";
import { DEFAULT_MERGE_OPTIONS, mergeSheets } from "./merge";
import { groupIssues, Issue, IssueGroup, summarize } from "./review";
import {
  applyFixes,
  clearIgnored,
  clearMarks,
  goToGroup,
  ignoreIssues,
  markIssues,
  reviewSheet,
  undoLastFix,
} from "./reviewSheet";
import { countProblems, FindingId, previewTidy, TidyFinding, tidyToNewSheet } from "./tidy";
import { describeCrosstab } from "./reshape";
import { crosstabColumns, CrosstabLook, findCrosstab, unpivotToNewSheet } from "./reshapeSheet";
import { compareSheets, guessSheetsToCompare } from "./compareSheets";
import { explainFormula } from "./explain";
import { readSelectedFormula } from "./explainSheet";
import {
  addStep,
  checkRecipe,
  isRepeatableIntent,
  Recipe,
  RecipeStep,
  sheetsUsedBy,
  StepOutcome,
  summarizeRun,
  suggestRecipeName,
} from "./recipes";
import { deleteRecipe, getRecipe, loadRecipes, saveRecipe } from "./recipeStore";
import { sheetInsights } from "./dashboard";

// ---------------------------------------------------------------------------
// What the assistant says
// ---------------------------------------------------------------------------

export type BotContent =
  | { type: "text"; text: string }
  | { type: "help" }
  | {
      type: "findings";
      sheet: string;
      rows: number;
      findings: TidyFinding[];
      skip: FindingId[];
    }
  | {
      type: "charts";
      sheet: string;
      rows: number;
      charts: DashChart[];
      chosen: number[];
      kpis: string[];
      problemsFixed: number;
    }
  | { type: "result"; result: OperationResult }
  | { type: "columns"; sheet: string; headers: string[] }
  | {
      type: "issues";
      sheet: string;
      /** One entry per kind of problem per column, worst first. */
      groups: IssueGroup[];
      /** "3 errors, 12 to check". */
      summary: string;
      /** Marked in the sheet itself. */
      marked: number;
      /** Rows that aren't data - a hint that /fix is the better tool. */
      structuralRows: number;
      truncated: boolean;
    }
  | {
      type: "crosstab";
      sheet: string;
      /** Columns that identify each row and will be carried across. */
      keys: string[];
      /** The period headings that will become rows. */
      periods: string[];
      /** One line describing the shape that was found. */
      shape: string;
    }
  | { type: "insights"; sheet: string; lines: string[] }
  | {
      type: "explain";
      address: string;
      formula: string;
      english: string;
      references: string[];
      warnings: Array<{ kind: string; text: string }>;
    }
  | {
      type: "recipes";
      recipes: Array<{ name: string; steps: number; updatedAt: string }>;
      /** Steps recorded in this session, available to save. */
      recorded: number;
    };

/** A suggested reply: what the chip says, and what it sends. */
export interface Chip {
  label: string;
  send: string;
}

export interface BotReply {
  content: BotContent[];
  chips: Chip[];
  /** Open one of the full tool screens. */
  openTool?: ToolName;
  /**
   * Start or stop watching a sheet. The subscription itself lives in the pane,
   * not here, because it outlives any one message.
   */
  watch?: { action: "start" | "stop"; sheet?: string };
}

type Pending =
  | null
  | { kind: "fix"; sheet: string; rows: number; findings: TidyFinding[]; skip: FindingId[] }
  | {
      kind: "dashboard";
      sheet: string;
      rows: number;
      charts: DashChart[];
      chosen: number[];
      kpis: string[];
      problemsFixed: number;
    }
  | { kind: "review"; sheet: string; groups: IssueGroup[] }
  | { kind: "unpivot"; sheet: string; look: CrosstabLook };

export interface AssistantState {
  pending: Pending;
  /**
   * Commands that did real work this session, in order, ready to be saved as a
   * recipe. Recording always happens - there is no mode to remember to turn on -
   * and nothing leaves the session until the user names a recipe.
   */
  steps: RecipeStep[];
  /**
   * The request that set up whatever is pending, so a bare "yes" can be recorded
   * as the command it actually agreed to.
   */
  lastRequest: string | null;
}

export const INITIAL_STATE: AssistantState = { pending: null, steps: [], lastRequest: null };

const say = (text: string, chips: Chip[] = []): BotReply => ({
  content: [{ type: "text", text }],
  chips,
});
const chip = (label: string, send = label): Chip => ({ label, send });

export const STARTER_CHIPS: Chip[] = [
  chip("Check this sheet", "/review"),
  chip("Fix this sheet", "/fix"),
  chip("Build a dashboard", "/dashboard"),
  chip("What can you do?", "/help"),
];

// ---------------------------------------------------------------------------
// Workbook context
// ---------------------------------------------------------------------------

interface WorkbookContext {
  sheets: string[];
  active: string;
}

async function workbookContext(): Promise<WorkbookContext> {
  try {
    return await Excel.run(async (context) => {
      const sheets = await loadSheetNames(context);
      const active = context.workbook.worksheets.getActiveWorksheet();
      active.load("name");
      await context.sync();
      return { sheets, active: active.name };
    });
  } catch {
    return { sheets: [], active: "" };
  }
}

// ---------------------------------------------------------------------------
// Handlers
// ---------------------------------------------------------------------------

function optionalCount(findings: readonly TidyFinding[], skip: readonly FindingId[]): number {
  return countProblems(findings.filter((finding) => !skip.includes(finding.id)));
}

function fixChips(findings: readonly TidyFinding[], skip: readonly FindingId[]): Chip[] {
  const chips = [chip(`Fix ${optionalCount(findings, skip).toLocaleString()} problems`, "yes")];
  const totals = findings.find((finding) => finding.id === "totals");
  if (totals && !skip.includes("totals")) chips.push(chip("Keep the total rows", "skip totals"));
  chips.push(chip("Cancel", "cancel"));
  return chips;
}

async function startFix(sheet: string): Promise<{ reply: BotReply; pending: Pending }> {
  const preview = await previewTidy({ kind: "sheet", name: sheet });
  const result = preview.result;
  if (!result) {
    return {
      reply: say(
        `I couldn't find a table on "${sheet}". It needs a row of headings with data underneath.`
      ),
      pending: null,
    };
  }
  if (result.findings.length === 0) {
    return {
      reply: say(`"${sheet}" already looks tidy: nothing to fix.`, [
        chip("Build a dashboard from it", `/dashboard ${sheet}`),
      ]),
      pending: null,
    };
  }
  const pending: Pending = {
    kind: "fix",
    sheet,
    rows: result.rows.length,
    findings: result.findings,
    skip: [],
  };
  return {
    reply: {
      content: [
        { type: "findings", sheet, rows: result.rows.length, findings: result.findings, skip: [] },
      ],
      chips: fixChips(result.findings, []),
    },
    pending,
  };
}

function dashboardChips(chosen: readonly number[]): Chip[] {
  return [
    chip(`Build ${chosen.length} ${chosen.length === 1 ? "chart" : "charts"}`, "yes"),
    ...(chosen.length > 4 ? [chip("Only the first 4", "only 4")] : []),
    chip("Cancel", "cancel"),
  ];
}

async function startDashboard(
  sheet: string,
  max: number | null = null
): Promise<{ reply: BotReply; pending: Pending }> {
  const analysis = await analyzeDashboard({ kind: "sheet", name: sheet }, max ? { max } : {});
  const plan = analysis.plan;
  if (!plan) {
    return { reply: say(`I couldn't find a table on "${sheet}" to chart.`), pending: null };
  }
  if (plan.charts.length === 0) {
    return {
      reply: say(
        `"${sheet}" has no column to group by - a region, product, category or date - so there's nothing to chart yet.`
      ),
      pending: null,
    };
  }
  const chosen = plan.charts.map((_chart, i) => i + 1);
  const pending: Pending = {
    kind: "dashboard",
    sheet,
    rows: plan.rowCount,
    charts: plan.charts,
    chosen,
    kpis: plan.kpis.map((kpi) => kpi.label),
    problemsFixed: analysis.problemsFixed,
  };
  return { reply: chartsReply(pending), pending };
}

/** Add a chart the user asked for by name to the dashboard being planned. */
async function addChartToPlan(
  intent: Extract<Intent, { kind: "chart" }>,
  pending: Extract<Pending, { kind: "dashboard" }>
): Promise<{ reply: BotReply; pending: Pending }> {
  const headers = await readSourceHeaders({ kind: "sheet", name: pending.sheet });
  const analysis = await analyzeDashboard({ kind: "sheet", name: pending.sheet });
  const profiles = analysis.plan?.profiles ?? [];
  const dimension = matchHeader(intent.dimension, headers);
  const measure = matchHeader(intent.measure, headers);

  if (!dimension && intent.chartType !== "xyscatter") {
    return {
      reply: {
        content: [
          {
            type: "text",
            text: "Which column should it group by? For example: add a pie of Units by Channel.",
          },
          { type: "columns", sheet: pending.sheet, headers },
        ],
        chips: dashboardChips(pending.chosen),
      },
      pending,
    };
  }

  const kind = CHART_KIND_FOR[intent.chartType ?? "columnClustered"] ?? "column";
  const chart = customChart(
    {
      kind,
      dimension,
      measure,
      measure2:
        kind === "scatter" ? (measure ? (headers.find((h) => h !== measure) ?? null) : null) : null,
      aggregation: measure ? (intent.aggregation === "average" ? "sum" : "sum") : "count",
    },
    profiles
  );
  if (!chart) {
    return {
      reply: say("I couldn't make a chart from those columns.", dashboardChips(pending.chosen)),
      pending,
    };
  }
  if (pending.charts.some((existing) => existing.id === chart.id)) {
    return {
      reply: say(`"${chart.title}" is already on the list.`, dashboardChips(pending.chosen)),
      pending,
    };
  }

  const charts = [...pending.charts, chart];
  const next = { ...pending, charts, chosen: [...pending.chosen, charts.length] };
  return { reply: chartsReply(next), pending: next };
}

/** Chart-type ids the parser produces, mapped to the dashboard's own kinds. */
const CHART_KIND_FOR: Record<string, DashChartKind> = {
  columnClustered: "column",
  columnStacked: "stackedColumn",
  barClustered: "bar",
  lineMarkers: "line",
  line: "line",
  area: "line",
  pie: "pie",
  doughnut: "doughnut",
  xyscatter: "scatter",
};

function chartsReply(pending: Extract<Pending, { kind: "dashboard" }>): BotReply {
  return {
    content: [
      {
        type: "charts",
        sheet: pending.sheet,
        rows: pending.rows,
        charts: pending.charts,
        chosen: pending.chosen,
        kpis: pending.kpis,
        problemsFixed: pending.problemsFixed,
      },
    ],
    chips: dashboardChips(pending.chosen),
  };
}

function resultReply(result: OperationResult, chips: Chip[] = []): BotReply {
  return { content: [{ type: "result", result }], chips };
}

function fixableCount(groups: readonly IssueGroup[]): number {
  return groups.reduce((total, group) => total + group.fixable, 0);
}

function allIssues(groups: readonly IssueGroup[]): Issue[] {
  return groups.flatMap((group) => group.issues);
}

/** Scan a sheet, mark the problem cells in Excel, and report what was found. */
async function startReview(sheet: string | null): Promise<{ reply: BotReply; pending: Pending }> {
  const review = await reviewSheet(sheet ?? undefined);
  if (!review) {
    return { reply: say("I couldn't find a table to check there."), pending: null };
  }
  if (review.issues.length === 0) {
    const chips: Chip[] = [chip("Build a dashboard", `/dashboard ${review.sheet}`)];
    if (review.structuralRows > 0)
      chips.unshift(chip("Tidy up the layout", `/fix ${review.sheet}`));
    return {
      reply: say(
        `I went through ${review.rows.toLocaleString()} rows of "${review.sheet}" and nothing looks wrong.` +
          (review.ignored > 0 ? ` (${review.ignored} ignored.)` : ""),
        chips
      ),
      pending: null,
    };
  }

  const groups = groupIssues(review.issues);
  const marked = await markIssues(review.sheet, review.issues);
  const fixable = fixableCount(groups);
  const chips: Chip[] = [];
  if (fixable > 0) chips.push(chip(`Fix the ${fixable} safe ones`, "yes"));
  chips.push(chip("Show me #1", "show 1"));
  if (review.structuralRows > 0) chips.push(chip("Tidy the layout", `/fix ${review.sheet}`));
  chips.push(chip("Clear marks", "/marks clear"));

  return {
    reply: {
      content: [
        {
          type: "issues",
          sheet: review.sheet,
          groups,
          summary: summarize(review.issues),
          marked: marked.ok ? Math.min(review.issues.length, 300) : 0,
          structuralRows: review.structuralRows,
          truncated: review.truncated,
        },
      ],
      chips,
    },
    pending: { kind: "review", sheet: review.sheet, groups },
  };
}

/** Groups picked by number ("fix 2 and 3"), or all of them. */
function pickedGroups(
  pending: Extract<Pending, { kind: "review" }>,
  numbers: readonly number[]
): IssueGroup[] {
  if (numbers.length === 0) return pending.groups;
  return numbers.map((n) => pending.groups[n - 1]).filter(Boolean);
}

async function chart(
  intent: Extract<Intent, { kind: "chart" }>,
  active: string
): Promise<BotReply> {
  const sheet = intent.sheet ?? active;
  const headers = await readSourceHeaders({ kind: "sheet", name: sheet });
  if (headers.length === 0) return say(`"${sheet}" looks empty.`);

  const dimension = matchHeader(intent.dimension, headers);
  const measure = matchHeader(intent.measure, headers);
  const example = (text: string) => chip(text, text);
  const exampleChips =
    headers.length >= 2
      ? [example(`/chart sum of ${headers[headers.length - 1]} by ${headers[0]}`)]
      : [];

  if (!dimension) {
    return {
      content: [
        {
          type: "text",
          text: intent.dimension
            ? `I couldn't find a column called "${intent.dimension}" on "${sheet}".`
            : "Which column should I group by? Say it like: sum of Revenue by Region.",
        },
        { type: "columns", sheet, headers },
      ],
      chips: exampleChips,
    };
  }
  if (intent.measure && !measure) {
    return {
      content: [
        {
          type: "text",
          text: `I couldn't find a column called "${intent.measure}" on "${sheet}".`,
        },
        { type: "columns", sheet, headers },
      ],
      chips: exampleChips,
    };
  }

  const groupByColumn = headers.indexOf(dimension);
  // Counting needs any other column to count; the group-by column can't count itself.
  const valueColumn = measure
    ? headers.indexOf(measure)
    : headers.findIndex((_h, i) => i !== groupByColumn);
  const aggregation = measure ? intent.aggregation : "count";
  const round = intent.chartType === "pie" || intent.chartType === "doughnut";
  const request: SummarizeRequest = {
    groupByColumn,
    valueColumns: [valueColumn],
    aggregation,
    sort: "valueDesc",
    topN: round ? 8 : null,
  };

  if (intent.summaryOnly) {
    return resultReply(
      await summarizeToNewSheet({ kind: "sheet", name: sheet }, request, "Summary")
    );
  }
  const title = `${measure ? aggregationLabel(aggregation, measure) : "Count"} by ${dimension}`;
  const result = await createChart({
    source: { kind: "sheet", name: sheet },
    summarize: request,
    chartTypeId: intent.chartType ?? "columnClustered",
    title,
    legendPosition: round ? "right" : "none",
    showDataLabels: round,
    destination: "newSheet",
    destinationSheetName: "Chart",
  });
  return resultReply(result, result.ok ? [] : [chip("Fix this sheet first", `/fix ${sheet}`)]);
}

// ---------------------------------------------------------------------------
// The conversation
// ---------------------------------------------------------------------------

/** Answer one message. Never throws: failures come back as a message. */
/** True when a reply reports that something actually succeeded. */
function didWork(reply: BotReply): boolean {
  return reply.content.some((item) => item.type === "result" && item.result.ok);
}

/** How a recorded step reads in the recipe list. */
function labelForIntent(intent: Intent, typed: string): string {
  switch (intent.kind) {
    case "fix":
      return "Clean up the sheet";
    case "review":
      return "Check for mistakes";
    case "dashboard":
      return "Build a dashboard";
    case "refresh":
      return "Refresh the dashboard";
    case "duplicates":
      return "Remove duplicate rows";
    case "spaces":
      return "Remove stray spaces";
    case "dates":
      return "Turn text dates into real dates";
    case "case":
      return "Fix text capitals";
    case "combine":
      return intent.all ? "Combine every sheet" : `Combine ${intent.sheets.length} sheets`;
    case "chart":
      return "Add a chart";
    case "unpivot":
      return "Turn period columns into rows";
    case "compare":
      return "Compare two sheets";
    case "fill":
      return "Fill the formula down";
    default:
      return typed;
  }
}

export async function respond(
  text: string,
  state: AssistantState
): Promise<{ reply: BotReply; state: AssistantState }> {
  try {
    const outcome = await handle(text, state);

    // Remember what worked, so it can be saved as a recipe later. A "yes" is
    // recorded as the request it agreed to, not as the word "yes".
    const typed =
      outcome.intent.kind === "confirm"
        ? state.lastRequest
        : isRepeatableIntent(outcome.intent)
          ? text
          : null;
    if (typed && didWork(outcome.reply)) {
      const recorded = parseCommand(typed, { sheets: outcome.sheets });
      if (isRepeatableIntent(recorded)) {
        const step: RecipeStep = {
          typed,
          label: labelForIntent(recorded, typed),
          sheets: sheetsUsedBy(recorded),
        };
        const kept = addStep({ ...EMPTY_RECIPE, steps: outcome.state.steps }, step);
        return { reply: outcome.reply, state: { ...outcome.state, steps: kept.steps } };
      }
    }
    return outcome;
  } catch (error) {
    return {
      reply: resultReply({
        ok: false,
        message: error instanceof Error ? error.message : String(error),
      }),
      state,
    };
  }
}

/** A shell recipe, only ever used to reuse addStep's rules for the session log. */
const EMPTY_RECIPE: Recipe = { name: "", createdAt: "", updatedAt: "", steps: [] };

// ---------------------------------------------------------------------------
// Recipes
// ---------------------------------------------------------------------------

/**
 * Replays a recipe by feeding each stored command back through the assistant.
 *
 * Steps run through exactly the same code path as if they had been typed, which
 * is the point: there is no second implementation of "clean up" that could drift
 * from the first. Two-phase steps are auto-confirmed, since agreeing to the
 * recipe is the agreement; a step whose sheet is missing is skipped and reported
 * rather than pointed at whatever sheet happens to be nearby.
 */
async function runRecipe(
  recipe: Recipe,
  state: AssistantState,
  sheets: readonly string[]
): Promise<{ outcomes: StepOutcome[]; state: AssistantState }> {
  const check = checkRecipe(recipe, sheets);
  const outcomes: StepOutcome[] = [];
  // Replaying must not fold the replayed commands back into the session log, or
  // saving afterwards would double every step.
  let working: AssistantState = { ...state, pending: null };

  for (const { step, runnable, missingSheets } of check.steps) {
    if (!runnable) {
      outcomes.push({
        step,
        ok: false,
        skipped: true,
        message: `needs ${missingSheets.map((name) => `"${name}"`).join(" and ")}, which this workbook hasn't got`,
      });
      continue;
    }

    const attempt = await handle(step.typed, working);
    working = attempt.state;
    let reply = attempt.reply;

    // Anything that asks "shall I?" is answered yes, once.
    if (working.pending) {
      const confirmed = await handle("yes", working);
      working = confirmed.state;
      reply = confirmed.reply;
    }

    const failure = reply.content.find(
      (item): item is { type: "result"; result: OperationResult } =>
        item.type === "result" && !item.result.ok
    );
    const success = reply.content.find(
      (item): item is { type: "result"; result: OperationResult } =>
        item.type === "result" && item.result.ok
    );
    outcomes.push({
      step,
      ok: Boolean(success),
      skipped: false,
      message: (success ?? failure)?.result.message ?? "nothing came back",
    });
  }

  return { outcomes, state: { ...state, pending: null } };
}

function recipeChips(recipes: readonly Recipe[], recorded: number): Chip[] {
  const chips: Chip[] = [];
  if (recipes.length > 0)
    chips.push(chip(`Run "${recipes[0].name}"`, `/recipe run ${recipes[0].name}`));
  if (recorded > 0) chips.push(chip("Save these steps", "/recipe save"));
  chips.push(chip("What can you do?", "/help"));
  return chips;
}

async function handle(
  text: string,
  state: AssistantState
): Promise<{
  reply: BotReply;
  state: AssistantState;
  /** What the message was understood as, so respond() can record it. */
  intent: Intent;
  /** Sheet names as they were when the message was handled. */
  sheets: string[];
}> {
  const workbook = await workbookContext();
  const intent = parseCommand(text, { sheets: workbook.sheets });
  const outcome = await route(text, state, intent, workbook);
  return { ...outcome, intent, sheets: workbook.sheets };
}

/** One message, one action. Split from handle() only so the intent is visible to both. */
async function route(
  text: string,
  state: AssistantState,
  intent: Intent,
  workbook: WorkbookContext
): Promise<{ reply: BotReply; state: AssistantState }> {
  const pending = state.pending;
  const keep = (reply: BotReply) => ({ reply, state });
  const clear = (reply: BotReply) => ({ reply, state: { ...state, pending: null } });
  /** A reply that leaves something waiting for a yes, remembering what asked for it. */
  const waiting = (reply: BotReply, next: Pending) => ({
    reply,
    state: { ...state, pending: next, lastRequest: text },
  });

  switch (intent.kind) {
    case "hello":
      return keep(
        say(
          "Hi! I fix messy sheets and build dashboards. Pick one below, or type / to see every command.",
          STARTER_CHIPS
        )
      );

    case "help":
      return keep({ content: [{ type: "help" }], chips: STARTER_CHIPS });

    case "fix": {
      const started = await startFix(intent.sheet ?? workbook.active);
      return {
        reply: started.reply,
        state: { ...state, pending: started.pending, lastRequest: text },
      };
    }

    case "dashboard": {
      const started = await startDashboard(intent.sheet ?? workbook.active, intent.max);
      return {
        reply: started.reply,
        state: { ...state, pending: started.pending, lastRequest: text },
      };
    }

    case "review": {
      const started = await startReview(intent.sheet ?? workbook.active);
      return {
        reply: started.reply,
        state: { ...state, pending: started.pending, lastRequest: text },
      };
    }

    case "show": {
      if (pending?.kind !== "review") break;
      const groups = pickedGroups(pending, intent.numbers);
      if (groups.length === 0) return keep(say("I don't have a number like that in the list."));
      const result = await goToGroup(pending.sheet, groups[0]);
      const issue = groups[0].issues[0];
      return keep({
        content: [
          { type: "text", text: `${groups[0].title} - ${issue.address}. ${issue.detail}` },
          { type: "result", result },
        ],
        chips: [
          ...(groups[0].fixable > 0
            ? [chip("Fix this one", `fix ${pending.groups.indexOf(groups[0]) + 1}`)]
            : []),
          chip("Ignore it", `ignore ${pending.groups.indexOf(groups[0]) + 1}`),
        ],
      });
    }

    case "ignore": {
      if (pending?.kind !== "review") break;
      const groups = pickedGroups(pending, intent.numbers);
      const result = await ignoreIssues(pending.sheet, allIssues(groups));
      const remaining = pending.groups.filter((group) => !groups.includes(group));
      return {
        reply: resultReply(result, [chip("Check again", `/review ${pending.sheet}`)]),
        state: {
          ...state,
          pending: remaining.length > 0 ? { ...pending, groups: remaining } : null,
        },
      };
    }

    case "unignore":
      return keep(resultReply(await clearIgnored(), [chip("Check again", "/review")]));

    case "undo":
      return keep(resultReply(await undoLastFix()));

    case "clearMarks":
      return keep(resultReply(await clearMarks()));

    case "confirm": {
      if (pending?.kind === "fix") {
        const { result } = await tidyToNewSheet(
          { kind: "sheet", name: pending.sheet },
          new Set(pending.skip)
        );
        // The dashboard chip points at the original export, not the clean copy:
        // that's the sheet next month's data will land in, so Refresh keeps working.
        return clear(
          resultReply(
            result,
            result.ok ? [chip("Build a dashboard", `/dashboard ${pending.sheet}`)] : []
          )
        );
      }
      if (pending?.kind === "review") {
        const issues = allIssues(pending.groups).filter((issue) => issue.fix);
        const result = await applyFixes(
          pending.sheet,
          issues,
          `fixed ${issues.length} cells on "${pending.sheet}"`
        );
        return clear(
          resultReply(result, [
            chip("Check again", `/review ${pending.sheet}`),
            chip("Undo that", "/undo"),
          ])
        );
      }
      if (pending?.kind === "unpivot") {
        const result = await unpivotToNewSheet(pending.sheet);
        return clear(
          resultReply(
            result,
            result.ok && result.sheetName
              ? [
                  chip("Build a dashboard from it", `/dashboard ${result.sheetName}`),
                  chip("Check it", `/review ${result.sheetName}`),
                ]
              : []
          )
        );
      }
      if (pending?.kind === "dashboard") {
        const charts = pending.chosen.map((n) => pending.charts[n - 1]).filter(Boolean);
        const result = await buildDashboard({ kind: "sheet", name: pending.sheet }, charts, {
          title: `${pending.sheet} dashboard`,
          includeKpis: true,
        });
        return clear(resultReply(result, result.ok ? [chip("Refresh it later", "/refresh")] : []));
      }
      return keep(say("There's nothing waiting for a yes right now.", STARTER_CHIPS));
    }

    case "unpivot": {
      const sheet = intent.sheet ?? workbook.active;
      const look = await findCrosstab(sheet);
      if (!look) return keep(say(`I couldn't read "${sheet}".`));
      if (!look.found) {
        return keep(
          say(
            `I can't see a crosstab on "${look.sheet}" - ${look.why}. If the periods are there under different headings, rename them and ask again.`,
            [chip("Check this sheet instead", `/review ${look.sheet}`)]
          )
        );
      }
      const columns = crosstabColumns(look.found.shape, look.found.headers);
      return waiting(
        {
          content: [
            {
              type: "crosstab",
              sheet: look.sheet,
              keys: columns.keys,
              periods: columns.periods,
              shape: describeCrosstab(look.found.shape),
            },
          ],
          chips: [chip("Write the tidy version", "yes"), chip("Cancel", "cancel")],
        },
        { kind: "unpivot", sheet: look.sheet, look }
      );
    }

    case "compare": {
      let first = intent.first;
      let second = intent.second;
      if (!first || !second) {
        const guess = await guessSheetsToCompare();
        if (!guess) {
          return keep(
            say(
              "Tell me which two sheets to compare, for example /compare April with May. This workbook needs at least two sheets with tables in them."
            )
          );
        }
        first = first ?? guess.first;
        second = second ?? guess.second;
      }
      if (first === second) {
        return keep(say("Those are the same sheet. Name two different ones, oldest first."));
      }
      const outcome = await compareSheets({ first, second });
      return clear(
        resultReply(outcome, [
          chip("Clear the marks", "/marks clear"),
          chip("Compare two others", "/compare"),
        ])
      );
    }

    case "insights": {
      const sheet = intent.sheet ?? workbook.active;
      const found = await sheetInsights({ kind: "sheet", name: sheet }, 4);
      if (!found || found.lines.length === 0) {
        return keep(
          say(
            `I couldn't find enough in "${sheet}" to say anything worth saying - I need a table with a column of numbers and something to group it by.`,
            [chip("Check this sheet", `/review ${sheet}`)]
          )
        );
      }
      return keep({
        content: [{ type: "insights", sheet: found.sourceName, lines: found.lines }],
        chips: [
          chip("Build a dashboard", `/dashboard ${sheet}`),
          chip("Check for mistakes", `/review ${sheet}`),
        ],
      });
    }

    case "explain": {
      const selected = await readSelectedFormula();
      if (!selected) {
        return keep(say("Select a cell first, then ask me to explain it."));
      }
      const explanation = explainFormula(selected.formula || String(selected.value ?? ""));
      const content: BotContent[] = [
        {
          type: "explain",
          address: `${selected.sheet}!${selected.address}`,
          formula: selected.formula,
          english: explanation.english,
          references: explanation.references,
          warnings: explanation.warnings,
        },
      ];
      if (selected.selectedCells > 1) {
        content.push({
          type: "text",
          text: `You had ${selected.selectedCells.toLocaleString()} cells selected, so that's the first one.`,
        });
      }
      return keep({ content, chips: [chip("Check this sheet", "/review")] });
    }

    case "recipe": {
      const saved = await loadRecipes();
      const recorded = state.steps;

      switch (intent.action) {
        case "list":
          if (saved.length === 0 && recorded.length === 0) {
            return keep(
              say(
                "No recipes saved in this workbook yet. Run a few commands - clean up, combine, dashboard - then say /recipe save month end and I'll remember the sequence.",
                STARTER_CHIPS
              )
            );
          }
          return keep({
            content: [
              {
                type: "recipes",
                recipes: saved.map((recipe) => ({
                  name: recipe.name,
                  steps: recipe.steps.length,
                  updatedAt: recipe.updatedAt,
                })),
                recorded: recorded.length,
              },
            ],
            chips: recipeChips(saved, recorded.length),
          });

        case "record":
          return {
            reply: say(
              "Starting a fresh recording. Everything that works from here goes into the recipe; save it when you're done.",
              [chip("What can you do?", "/help")]
            ),
            state: { ...state, steps: [] },
          };

        case "stop":
          if (recorded.length === 0) {
            return keep(say("I haven't recorded anything yet - nothing has run this session."));
          }
          return keep(
            say(
              `I have ${plural(recorded.length, "step")}: ${recorded.map((step) => step.label.toLowerCase()).join(", then ")}. Name it to keep it.`,
              [chip("Save it", "/recipe save"), chip("Forget it", "/recipe record")]
            )
          );

        case "save": {
          if (recorded.length === 0) {
            return keep(
              say(
                "There's nothing to save yet. Run a command or two - I record whatever works - then save.",
                STARTER_CHIPS
              )
            );
          }
          const name = intent.name ?? suggestRecipeName(saved.map((recipe) => recipe.name));
          const result = await saveRecipe(name, recorded);
          return clear(
            resultReply(result, result.ok ? [chip(`Run "${name}" now`, `/recipe run ${name}`)] : [])
          );
        }

        case "delete": {
          if (!intent.name) {
            return keep(
              say(
                "Which one? Say /recipe delete followed by its name.",
                recipeChips(saved, recorded.length)
              )
            );
          }
          return clear(resultReply(await deleteRecipe(intent.name)));
        }

        case "run": {
          if (saved.length === 0) {
            return keep(
              say(
                "There are no recipes in this workbook yet. Run the steps once, then /recipe save gives them a name.",
                STARTER_CHIPS
              )
            );
          }
          const recipe = intent.name
            ? await getRecipe(intent.name)
            : saved.length === 1
              ? saved[0]
              : null;
          if (!recipe) {
            return keep(
              say(
                intent.name
                  ? `I have no recipe called "${intent.name}". Saved here: ${saved.map((item) => item.name).join(", ")}.`
                  : `Which one? Saved here: ${saved.map((item) => item.name).join(", ")}.`,
                saved.slice(0, 3).map((item) => chip(item.name, `/recipe run ${item.name}`))
              )
            );
          }

          const run = await runRecipe(recipe, state, workbook.sheets);
          const details = run.outcomes.map(
            (outcome) =>
              `${outcome.ok ? "Done" : outcome.skipped ? "Skipped" : "Failed"} - ${outcome.step.label}: ${outcome.message}`
          );
          const allOk = run.outcomes.every((outcome) => outcome.ok);
          return {
            reply: resultReply(
              {
                ok: allOk,
                message: summarizeRun(recipe.name, run.outcomes),
                details,
              },
              [chip("Check the result", "/review"), chip("Undo the last fix", "/undo")]
            ),
            state: run.state,
          };
        }
      }
      return keep(say("I didn't follow that. Try /recipe to see what's saved."));
    }

    case "watch": {
      const turnOn = intent.on ?? state.pending === null;
      if (!turnOn) {
        return keep({
          content: [{ type: "text", text: "Stopped watching. Nothing else has changed." }],
          chips: STARTER_CHIPS,
          watch: { action: "stop" },
        });
      }
      return keep({
        content: [
          {
            type: "text",
            text: `Watching "${workbook.active}". As you type, I'll re-check the cells you touch and say if something looks wrong. I won't change anything.`,
          },
        ],
        chips: [chip("Stop watching", "/watch off")],
        watch: { action: "start", sheet: workbook.active },
      });
    }

    case "cancel":
      return clear(
        say(pending ? "OK, cancelled. Nothing was changed." : "Nothing to cancel.", STARTER_CHIPS)
      );

    case "skip": {
      if (pending?.kind !== "fix") break;
      const present = new Set(
        pending.findings.filter((finding) => finding.optional).map((finding) => finding.id)
      );
      const targets = skipTargets(intent.words).filter((id): id is FindingId =>
        present.has(id as FindingId)
      );
      if (targets.length === 0) {
        return keep(
          say(
            "I can leave out: " +
              pending.findings
                .filter((finding) => finding.optional)
                .map((finding) => finding.label.toLowerCase())
                .join("; ") +
              ".",
            fixChips(pending.findings, pending.skip)
          )
        );
      }
      const skip = Array.from(new Set([...pending.skip, ...targets]));
      const next: Pending = { ...pending, skip };
      return {
        reply: {
          content: [
            { type: "findings", sheet: next.sheet, rows: next.rows, findings: next.findings, skip },
          ],
          chips: fixChips(next.findings, skip),
        },
        state: { ...state, pending: next },
      };
    }

    case "pick": {
      if (pending?.kind === "review") {
        const groups = pickedGroups(pending, intent.numbers);
        const issues = allIssues(groups).filter((issue) => issue.fix);
        if (issues.length === 0) {
          return keep(
            say("Those need a decision from you, so I won't change them on a guess.", [
              chip("Show me #1", "show 1"),
            ])
          );
        }
        const result = await applyFixes(
          pending.sheet,
          issues,
          `fixed ${issues.length} cells on "${pending.sheet}"`
        );
        return keep(
          resultReply(result, [
            chip("Check again", `/review ${pending.sheet}`),
            chip("Undo that", "/undo"),
          ])
        );
      }
      if (pending?.kind !== "dashboard") break;
      const all = pending.charts.map((_c, i) => i + 1);
      const valid = intent.numbers.filter((n) => n <= pending.charts.length);
      const chosen = intent.exclude
        ? pending.chosen.filter((n) => !valid.includes(n))
        : all.filter((n) => valid.includes(n));
      if (chosen.length === 0)
        return keep(
          say("That leaves no charts. Pick at least one.", dashboardChips(pending.chosen))
        );
      const next = { ...pending, chosen };
      return { reply: chartsReply(next), state: { ...state, pending: next } };
    }

    case "refresh": {
      const dashboards = await listDashboards();
      if (dashboards.length === 0) {
        return keep(
          say("There are no saved dashboards in this workbook yet.", [
            chip("Build a dashboard", "/dashboard"),
          ])
        );
      }
      const named = intent.name
        ? dashboards.find((d) => d.name.toLowerCase().includes(intent.name!.toLowerCase()))
        : dashboards.length === 1
          ? dashboards[0]
          : undefined;
      if (!named) {
        return keep(
          say(
            "Which dashboard?",
            dashboards.map((d) => chip(d.name, `/refresh ${d.name}`))
          )
        );
      }
      return clear(resultReply(await refreshDashboard(named)));
    }

    case "duplicates":
      return clear(
        resultReply(
          await removeDuplicateRows({
            scope: "selection",
            hasHeaderRow: true,
            keyColumns: [],
            ignoreCaseAndSpacing: true,
            deleteEntireRows: false,
          })
        )
      );

    case "spaces":
      return clear(
        resultReply(
          await trimWhitespace({
            scope: "selection",
            collapseInnerSpaces: true,
            blankOutWhitespaceOnlyCells: true,
          })
        )
      );

    case "dates":
      if (intent.dayFirst === null) {
        return keep(
          say("Are your dates day-first (03/04/2024 is 3 April) or month-first (March 4)?", [
            chip("Day-first", "/dates day-first"),
            chip("Month-first", "/dates month-first"),
          ])
        );
      }
      return clear(
        resultReply(
          await standardizeDates({
            scope: "selection",
            numberFormat: "yyyy-mm-dd",
            dayFirst: intent.dayFirst,
          })
        )
      );

    case "case":
      if (intent.mode === null) {
        return keep(
          say(
            "Which style? Select the cells first; with one cell selected I change the whole table, headings included.",
            [
              chip("Title Case", "/case title"),
              chip("UPPER CASE", "/case upper"),
              chip("lower case", "/case lower"),
              chip("Sentence case", "/case sentence"),
            ]
          )
        );
      }
      return clear(
        resultReply(await standardizeTextCase({ scope: "selection", mode: intent.mode }))
      );

    case "combine": {
      const sheets = intent.all ? workbook.sheets : intent.sheets;
      if (sheets.length < 2) {
        const example = workbook.sheets.slice(0, 2).join(", ");
        return keep(
          say(
            "Which sheets should I combine? Name two or more.",
            example
              ? [
                  chip(`/combine ${example}`, `/combine ${example}`),
                  chip("Combine all sheets", "/combine all"),
                ]
              : []
          )
        );
      }
      return clear(resultReply(await mergeSheets(sheets, { ...DEFAULT_MERGE_OPTIONS })));
    }

    case "chart": {
      // "add a pie of Units by Channel" while a dashboard is being planned
      // extends that plan rather than making a chart on its own.
      if (intent.add && pending?.kind === "dashboard") {
        const added = await addChartToPlan(intent, pending);
        return { reply: added.reply, state: { ...state, pending: added.pending } };
      }
      return clear(await chart(intent, workbook.active));
    }

    case "fill":
      return clear(resultReply(await fillFormulaAcrossSelection()));

    case "open": {
      const names: Record<ToolName, string> = {
        home: "the full tool list",
        review: "the Check for mistakes screen",
        formulas: "the Formulas tool",
        reports: "the Update a report tool",
        merge: "the Combine sheets tool",
        charts: "the Make a chart tool",
        clean: "Quick clean-ups",
      };
      return keep({
        content: [{ type: "text", text: `Opening ${names[intent.tool]}.` }],
        chips: [],
        openTool: intent.tool,
      });
    }

    case "unknown":
      break;
  }

  // Unrecognised, or a reply that only makes sense after a question.
  const suggestions = intent.kind === "unknown" ? intent.suggestions : [];
  return keep(
    say(
      "I didn't catch that. Try one of these, or type / to see every command.",
      [...suggestions.map((command) => chip(command, command)), ...STARTER_CHIPS].slice(0, 4)
    )
  );
}

/** The command list, grouped, for the help message. */
export function helpGroups(): Array<{ group: string; commands: typeof COMMANDS }> {
  const groups: Array<{ group: string; commands: typeof COMMANDS }> = [];
  for (const info of COMMANDS) {
    let entry = groups.find((g) => g.group === info.group);
    if (!entry) {
      entry = { group: info.group, commands: [] };
      groups.push(entry);
    }
    entry.commands.push(info);
  }
  return groups;
}
