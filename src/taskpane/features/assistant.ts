/**
 * The assistant: turns each message into an action using the same tested
 * features as the tool screens, and answers with structured messages the chat
 * renders (findings, chart plans, results) plus a few suggested replies.
 *
 * Anything that changes the workbook in a big way - fixing a sheet, building
 * a dashboard - is previewed first and waits for a "yes". Small, targeted
 * clean-ups run straight away, on the cells shown in the reply.
 */

import { loadSheetNames } from "../shared/excelHelpers";
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
  DashChart,
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
import { countProblems, FindingId, previewTidy, TidyFinding, tidyToNewSheet } from "./tidy";

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
  | { type: "columns"; sheet: string; headers: string[] };

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
    };

export interface AssistantState {
  pending: Pending;
}

export const INITIAL_STATE: AssistantState = { pending: null };

const say = (text: string, chips: Chip[] = []): BotReply => ({
  content: [{ type: "text", text }],
  chips,
});
const chip = (label: string, send = label): Chip => ({ label, send });

export const STARTER_CHIPS: Chip[] = [
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

async function startDashboard(sheet: string): Promise<{ reply: BotReply; pending: Pending }> {
  const analysis = await analyzeDashboard({ kind: "sheet", name: sheet });
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
export async function respond(
  text: string,
  state: AssistantState
): Promise<{ reply: BotReply; state: AssistantState }> {
  try {
    return await handle(text, state);
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

async function handle(
  text: string,
  state: AssistantState
): Promise<{ reply: BotReply; state: AssistantState }> {
  const workbook = await workbookContext();
  const intent = parseCommand(text, { sheets: workbook.sheets });
  const pending = state.pending;
  const keep = (reply: BotReply) => ({ reply, state });
  const clear = (reply: BotReply) => ({ reply, state: { pending: null } });

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
      return { reply: started.reply, state: { pending: started.pending } };
    }

    case "dashboard": {
      const started = await startDashboard(intent.sheet ?? workbook.active);
      return { reply: started.reply, state: { pending: started.pending } };
    }

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
        state: { pending: next },
      };
    }

    case "pick": {
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
      return { reply: chartsReply(next), state: { pending: next } };
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

    case "chart":
      return clear(await chart(intent, workbook.active));

    case "fill":
      return clear(resultReply(await fillFormulaAcrossSelection()));

    case "open": {
      const names: Record<ToolName, string> = {
        home: "the full tool list",
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
