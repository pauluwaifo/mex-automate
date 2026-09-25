/**
 * The assistant's command language.
 *
 * MEx is a command assistant, not an AI: it runs no model and makes no network
 * calls. It understands a fixed set of slash commands plus the everyday
 * phrasings people actually type ("fix this sheet", "remove duplicates",
 * "pie of revenue by region"), and when it can't tell what was meant it says
 * so and suggests the closest commands - it never guesses at a destructive
 * action.
 *
 * Everything here is pure and unit tested in tests/commands.test.ts.
 */

import { normalizeHeader } from "../shared/excelHelpers";
import type { Aggregation } from "./charts";
import type { TextCaseMode } from "./dataCleaning";
import { editDistance } from "./tidy";

// ---------------------------------------------------------------------------
// The command list (drives /help and the autocomplete)
// ---------------------------------------------------------------------------

export interface CommandInfo {
  command: string;
  /** Example shown in autocomplete and help. */
  example: string;
  description: string;
  group: "Main" | "Clean up" | "Charts" | "More";
}

export const COMMANDS: CommandInfo[] = [
  {
    command: "/review",
    example: "/review",
    description: "Check this sheet for mistakes and mark them in Excel",
    group: "Main",
  },
  {
    command: "/fix",
    example: "/fix",
    description: "Find and fix everything wrong with this sheet",
    group: "Main",
  },
  {
    command: "/dashboard",
    example: "/dashboard",
    description: "Build a dashboard of up to 8 charts",
    group: "Main",
  },
  {
    command: "/refresh",
    example: "/refresh",
    description: "Rebuild a saved dashboard with the latest data",
    group: "Main",
  },
  {
    command: "/duplicates",
    example: "/duplicates",
    description: "Remove duplicate rows",
    group: "Clean up",
  },
  {
    command: "/spaces",
    example: "/spaces",
    description: "Remove extra and invisible spaces",
    group: "Clean up",
  },
  {
    command: "/dates",
    example: "/dates day-first",
    description: "Turn dates typed as text into real dates",
    group: "Clean up",
  },
  {
    command: "/case",
    example: "/case title",
    description: "Make text UPPER, lower, Title or Sentence case",
    group: "Clean up",
  },
  {
    command: "/chart",
    example: "/chart sum of Revenue by Region as pie",
    description: "One chart, in plain words",
    group: "Charts",
  },
  {
    command: "/summary",
    example: "/summary sum of Revenue by Region",
    description: "A totals table on a new sheet",
    group: "Charts",
  },
  {
    command: "/combine",
    example: "/combine Jan Orders, Feb Orders",
    description: "Stack sheets into one table",
    group: "More",
  },
  {
    command: "/fill",
    example: "/fill",
    description: "Copy the formula at the top of your selection down",
    group: "More",
  },
  {
    command: "/undo",
    example: "/undo",
    description: "Put back the cells from the last fix",
    group: "More",
  },
  {
    command: "/marks",
    example: "/marks clear",
    description: "Clear the marks Review left in the sheet",
    group: "More",
  },
  {
    command: "/tools",
    example: "/tools",
    description: "Open the full tool screens (formulas, reports and more)",
    group: "More",
  },
  { command: "/help", example: "/help", description: "Show everything I can do", group: "More" },
];

/** Commands whose name starts with what has been typed so far, for autocomplete. */
export function completeCommand(typed: string): CommandInfo[] {
  const text = typed.trim().toLowerCase();
  if (!text.startsWith("/")) return [];
  if (text.includes(" ")) return [];
  return COMMANDS.filter((info) => info.command.startsWith(text));
}

// ---------------------------------------------------------------------------
// Intents
// ---------------------------------------------------------------------------

export type ToolName = "review" | "formulas" | "reports" | "merge" | "charts" | "clean" | "home";

export type Intent =
  | { kind: "help" }
  | { kind: "hello" }
  | { kind: "fix"; sheet: string | null }
  | { kind: "review"; sheet: string | null }
  | { kind: "undo" }
  | { kind: "clearMarks" }
  | { kind: "unignore" }
  | { kind: "show"; numbers: number[] }
  | { kind: "ignore"; numbers: number[] }
  | { kind: "dashboard"; sheet: string | null }
  | { kind: "refresh"; name: string | null }
  | { kind: "duplicates" }
  | { kind: "spaces" }
  | { kind: "dates"; dayFirst: boolean | null }
  | { kind: "case"; mode: TextCaseMode | null }
  | { kind: "combine"; sheets: string[]; all: boolean }
  | {
      kind: "chart";
      summaryOnly: boolean;
      aggregation: Aggregation;
      measure: string | null;
      dimension: string | null;
      chartType: string | null;
      sheet: string | null;
    }
  | { kind: "fill" }
  | { kind: "open"; tool: ToolName }
  | { kind: "confirm" }
  | { kind: "cancel" }
  | { kind: "skip"; words: string[] }
  | { kind: "pick"; numbers: number[]; exclude: boolean }
  | { kind: "unknown"; suggestions: string[] };

export interface ParseContext {
  /** Sheet names in the workbook, for "on Sales Extract" and /combine. */
  sheets: readonly string[];
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function squash(text: string): string {
  return text
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

/** Sheet names mentioned in the text, longest first, so "Sales Extract (clean)" beats "Sales Extract". */
export function findSheets(text: string, sheets: readonly string[]): string[] {
  const haystack = ` ${squash(text)} `;
  const found: string[] = [];
  let remaining = haystack;
  for (const sheet of [...sheets].sort((a, b) => b.length - a.length)) {
    const needle = ` ${squash(sheet)} `;
    if (needle.trim() !== "" && remaining.includes(needle)) {
      found.push(sheet);
      remaining = remaining.replace(needle, " ");
    }
  }
  // Report them in the order they were written.
  return found.sort(
    (a, b) => haystack.indexOf(` ${squash(a)} `) - haystack.indexOf(` ${squash(b)} `)
  );
}

/**
 * Match what the user typed ("revenue", "the region column") to one of the
 * sheet's headings, forgiving case, spacing, plurals and small typos.
 */
export function matchHeader(typed: string | null, headers: readonly string[]): string | null {
  if (!typed) return null;
  const want = normalizeHeader(typed.replace(/\b(the|column|field)\b/gi, " "));
  if (want === "") return null;
  const keyed = headers.map((header) => ({ header, key: normalizeHeader(header) }));

  const exact = keyed.find((item) => item.key === want);
  if (exact) return exact.header;
  const singular = want.replace(/s$/, "");
  const plural = keyed.find(
    (item) => item.key === singular || item.key.replace(/s$/, "") === singular
  );
  if (plural) return plural.header;
  const contains = keyed.filter((item) => item.key.includes(want) || want.includes(item.key));
  if (contains.length === 1) return contains[0].header;
  const close = keyed
    .map((item) => ({ ...item, distance: editDistance(item.key, want, 2) }))
    .filter((item) => item.distance <= (want.length >= 6 ? 2 : 1))
    .sort((a, b) => a.distance - b.distance);
  return close[0]?.header ?? null;
}

const CHART_WORDS: Array<[RegExp, string]> = [
  [/\b(doughnut|donut|ring)\b/, "doughnut"],
  [/\bpie\b/, "pie"],
  [/\bstacked\s+(column|bar)s?\b/, "columnStacked"],
  [/\b(bar|horizontal)\b/, "barClustered"],
  [/\b(column|vertical)\b/, "columnClustered"],
  [/\b(line|trend)\b/, "lineMarkers"],
  [/\barea\b/, "area"],
  [/\b(scatter|dots?)\b/, "xyscatter"],
];

const AGGREGATION_WORDS: Array<[RegExp, Aggregation]> = [
  [/\b(average|avg|mean)\b/, "average"],
  [/\b(unique|distinct)\b/, "countDistinct"],
  [/\b(count|number of|how many)\b/, "count"],
  [/\b(min|minimum|lowest|smallest)\b/, "min"],
  [/\b(max|maximum|highest|largest|biggest)\b/, "max"],
  [/\b(sum|total)\b/, "sum"],
];

function parseChart(body: string, summaryOnly: boolean, sheets: readonly string[]): Intent {
  let text = body.toLowerCase();

  // "... on Sales Extract" / "... from Sales Extract"
  let sheet: string | null = null;
  const sheetMatch = /\s(?:on|from|in|of sheet|using)\s+(.+)$/.exec(text);
  if (sheetMatch) {
    const named = findSheets(sheetMatch[1], sheets)[0];
    if (named) {
      sheet = named;
      text = text.slice(0, sheetMatch.index);
    }
  }

  let chartType: string | null = null;
  for (const [pattern, id] of CHART_WORDS) {
    if (pattern.test(text)) {
      chartType = id;
      break;
    }
  }

  let aggregation: Aggregation = "sum";
  for (const [pattern, value] of AGGREGATION_WORDS) {
    if (pattern.test(text)) {
      aggregation = value;
      break;
    }
  }

  // Strip the words that aren't column names, leaving "<measure> by <dimension>".
  const cleaned = text
    .replace(/\b(as|in|into)\s+(a|an)?\s*[a-z ]*chart\b/g, " ")
    .replace(/\b(as|in)\s+(a|an)\s+\w+$/g, " ")
    .replace(
      /\b(make|draw|create|show|plot|give me|build|chart|graph|summary|summari[sz]e|table|please)\b/g,
      " "
    )
    .replace(
      /\b(doughnut|donut|ring|pie|stacked|bar|horizontal|column|vertical|line|trend|area|scatter|dots?)s?\b/g,
      " "
    )
    .replace(
      /\b(average|avg|mean|unique|distinct|count|number of|how many|min|minimum|lowest|smallest|max|maximum|highest|largest|biggest|sum|total)\b/g,
      " "
    )
    .replace(/\b(a|an|the|of|me)\b/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    // "... by region as" once "pie" has gone.
    .replace(/\s+(as|in|into|with|using)$/, "");

  const split =
    /^(.*?)\s*\b(?:by|per|for each|for every|across|split by|grouped by)\b\s*(.*)$/.exec(cleaned);
  const measure = split ? split[1].trim() || null : cleaned || null;
  const dimension = split ? split[2].trim() || null : null;

  return { kind: "chart", summaryOnly, aggregation, measure, dimension, chartType, sheet };
}

const NUMBER_WORDS: Record<string, number> = {
  one: 1,
  two: 2,
  three: 3,
  four: 4,
  five: 5,
  six: 6,
  seven: 7,
  eight: 8,
};

function parseNumbers(text: string): number[] {
  const numbers: number[] = [];
  const pattern = /\b(\d+|one|two|three|four|five|six|seven|eight)\b/g;
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(text)) !== null) {
    const value = /^\d+$/.test(match[1]) ? Number(match[1]) : NUMBER_WORDS[match[1]];
    if (value >= 1 && value <= 8 && !numbers.includes(value)) numbers.push(value);
  }
  return numbers;
}

/** The closest command names to an unrecognised word, for "Did you mean /fix?" */
export function suggestCommands(text: string): string[] {
  const word = squash(text).split(" ")[0] ?? "";
  if (word === "") return ["/help"];
  return COMMANDS.map((info) => ({
    command: info.command,
    distance: editDistance(info.command.slice(1), word, 3),
  }))
    .filter((item) => item.distance <= 3)
    .sort((a, b) => a.distance - b.distance)
    .slice(0, 3)
    .map((item) => item.command);
}

// ---------------------------------------------------------------------------
// The parser
// ---------------------------------------------------------------------------

/** Work out what the user wants from one message. */
export function parseCommand(input: string, context: ParseContext): Intent {
  const raw = input.trim();
  const text = raw.toLowerCase();
  if (text === "") return { kind: "unknown", suggestions: ["/help"] };

  // Slash commands: the first word decides.
  if (text.startsWith("/")) {
    const [head, ...rest] = raw.slice(1).split(/\s+/);
    const body = rest.join(" ");
    const sheetIn = (value: string) => findSheets(value, context.sheets)[0] ?? null;
    switch (head.toLowerCase()) {
      case "help":
      case "commands":
      case "?":
        return { kind: "help" };
      case "fix":
      case "tidy":
      case "clean":
        return { kind: "fix", sheet: sheetIn(body) };
      case "review":
      case "check":
      case "audit":
        return { kind: "review", sheet: sheetIn(body) };
      case "undo":
        return { kind: "undo" };
      case "marks":
      case "unmark":
        return { kind: "clearMarks" };
      case "unignore":
        return { kind: "unignore" };
      case "show":
      case "goto":
        return { kind: "show", numbers: parseNumbers(body) };
      case "ignore":
      case "hide":
        return { kind: "ignore", numbers: parseNumbers(body) };
      case "dashboard":
      case "dash":
        return { kind: "dashboard", sheet: sheetIn(body) };
      case "refresh":
      case "update":
        return { kind: "refresh", name: body.trim() || null };
      case "duplicates":
      case "dupes":
      case "dedupe":
        return { kind: "duplicates" };
      case "spaces":
      case "trim":
        return { kind: "spaces" };
      case "dates":
      case "date":
        return parseDates(body);
      case "case":
        return parseCase(body);
      case "chart":
      case "graph":
      case "plot":
        return parseChart(body, false, context.sheets);
      case "summary":
      case "summarize":
      case "summarise":
      case "totals":
        return parseChart(body, true, context.sheets);
      case "combine":
      case "merge":
        return parseCombine(body, context.sheets);
      case "fill":
        return { kind: "fill" };
      case "tools":
        return { kind: "open", tool: "home" };
      case "formulas":
        return { kind: "open", tool: "formulas" };
      case "report":
      case "reports":
        return { kind: "open", tool: "reports" };
      default:
        return { kind: "unknown", suggestions: suggestCommands(head) };
    }
  }

  // Short replies to a question the assistant just asked.
  if (
    /^(yes|y|yeah|yep|ok|okay|sure|go|go ahead|do it|confirm|run it|fix it|fix all|fix them|build it|build all|build them|build)[.!]*$/.test(
      text
    )
  ) {
    return { kind: "confirm" };
  }
  if (/^(no|nope|cancel|stop|never ?mind|forget it)[.!]*$/.test(text)) {
    return { kind: "cancel" };
  }
  const skip = /^(?:skip|keep|leave|don'?t (?:fix|remove|touch)|without|except)\s+(.+)$/.exec(text);
  if (skip && !/\d/.test(skip[1])) {
    return {
      kind: "skip",
      words: skip[1]
        .split(/\s*(?:,|\band\b)\s*/)
        .map((word) => word.trim())
        .filter(Boolean),
    };
  }
  if (
    /^(only|just|build|fix|use|keep|charts?|without|except|drop|remove|not)\b.*\d|^(only|just|top|first)\s+(one|two|three|four|five|six|seven|eight)\b/.test(
      text
    )
  ) {
    const numbers = parseNumbers(text);
    const exclude = /\b(without|except|drop|remove|not)\b/.test(text);
    // "only 4" / "top 4" / "first 4" means the first four, not chart number 4.
    if (/^(only|just|top|first)\s+\S+(\s+charts?)?$/.test(text) && numbers.length === 1) {
      return {
        kind: "pick",
        numbers: Array.from({ length: numbers[0] }, (_v, i) => i + 1),
        exclude: false,
      };
    }
    if (numbers.length > 0) return { kind: "pick", numbers, exclude };
  }

  if (/^(help|\?|what can you do|what do you do|commands|how does this work)\b/.test(text))
    return { kind: "help" };
  if (/^(hi|hello|hey|good (morning|afternoon|evening))\b/.test(text)) return { kind: "hello" };

  // Plain-language phrasings, most specific first.
  const sheet = findSheets(raw, context.sheets)[0] ?? null;
  if (/^(undo|undo that|put (it|them) back|revert)\b/.test(text)) return { kind: "undo" };
  if (/\b(clear|remove|hide)\b.*\b(marks?|highlights?|colou?rs?)\b/.test(text))
    return { kind: "clearMarks" };
  if (/\bunignore\b|\bbring (them|it) back\b/.test(text)) return { kind: "unignore" };
  if (/^(ignore|hide|dismiss)\b/.test(text)) return { kind: "ignore", numbers: parseNumbers(text) };
  if (/^(show|go ?to|take me to)\s+(me\s+)?\d/.test(text))
    return { kind: "show", numbers: parseNumbers(text) };
  if (
    /\b(review|check|audit|proof ?read)\b/.test(text) ||
    /what'?s wrong|any (mistakes|errors|problems)|is (this|it) (right|correct)/.test(text)
  ) {
    return { kind: "review", sheet };
  }
  if (/\b(refresh|rebuild|update (the |my )?dashboard)\b/.test(text))
    return { kind: "refresh", name: null };
  if (/\bdash ?board\b/.test(text)) return { kind: "dashboard", sheet };
  if (/\b(combine|merge|stack|append)\b/.test(text))
    return parseCombine(raw.replace(/^.*?\b(combine|merge|stack|append)\b/i, ""), context.sheets);
  if (/\b(duplicates?|dupes?|dedupe|repeated rows)\b/.test(text)) return { kind: "duplicates" };
  if (/\b(spaces?|whitespace|trim)\b/.test(text)) return { kind: "spaces" };
  if (
    /\b(upper ?case|lower ?case|title ?case|proper ?case|sentence ?case|capitals?|capital letters|caps)\b/.test(
      text
    )
  )
    return parseCase(text);
  if (/\b(summary|summari[sz]e|totals? table)\b/.test(text))
    return parseChart(raw, true, context.sheets);
  if (
    /\b(chart|graph|plot|pie|doughnut|donut|scatter)\b/.test(text) ||
    /\b\w+ (by|per) \w+/.test(text)
  ) {
    return parseChart(raw, false, context.sheets);
  }
  if (/\bdates?\b/.test(text)) return parseDates(text);
  if (/\b(fill( it)? down|copy (the |this |my )?formula|fill (the |this )?formula)\b/.test(text))
    return { kind: "fill" };
  if (/\bformulas?\b/.test(text)) return { kind: "open", tool: "formulas" };
  if (/\breports?\b/.test(text)) return { kind: "open", tool: "reports" };
  if (/\b(tools|menu|settings)\b/.test(text)) return { kind: "open", tool: "home" };
  if (/\b(fix|tidy|clean|messy|repair|sort out|problems?)\b/.test(text))
    return { kind: "fix", sheet };

  return { kind: "unknown", suggestions: suggestCommands(text) };
}

function parseDates(text: string): Intent {
  const lower = text.toLowerCase();
  if (/\b(day[- ]?first|dd\/?mm|uk|european?)\b/.test(lower))
    return { kind: "dates", dayFirst: true };
  if (/\b(month[- ]?first|mm\/?dd|us|american)\b/.test(lower))
    return { kind: "dates", dayFirst: false };
  return { kind: "dates", dayFirst: null };
}

function parseCase(text: string): Intent {
  const lower = text.toLowerCase();
  if (/\bupper|caps\b|capitals/.test(lower)) return { kind: "case", mode: "upper" };
  if (/\blower/.test(lower)) return { kind: "case", mode: "lower" };
  if (/\bsentence/.test(lower)) return { kind: "case", mode: "sentence" };
  if (/\b(title|proper)/.test(lower)) return { kind: "case", mode: "proper" };
  return { kind: "case", mode: null };
}

function parseCombine(body: string, sheets: readonly string[]): Intent {
  if (/\b(all|every|each)\b/i.test(body)) return { kind: "combine", sheets: [], all: true };
  return { kind: "combine", sheets: findSheets(body, sheets), all: false };
}

/**
 * Map the words in "skip totals and spellings" onto the fixes they name.
 * Returns the finding ids as plain strings so this module stays independent
 * of the clean-up engine's types.
 */
export function skipTargets(words: readonly string[]): string[] {
  const ids = new Set<string>();
  const table: Array<[RegExp, string]> = [
    [/total|subtotal/, "totals"],
    [/spell|typo|name|capital/, "spellings"],
    [/dup/, "duplicates"],
    [/date/, "textDates"],
    [/number|amount|currenc|money/, "textNumbers"],
    [/group|section|region|heading/, "sections"],
    [/space|trim/, "spaces"],
    [/n\/?a|placeholder|blank|dash/, "placeholders"],
    [/empty col|column/, "emptyColumns"],
  ];
  for (const word of words) {
    for (const [pattern, id] of table) {
      if (pattern.test(word.toLowerCase())) ids.add(id);
    }
  }
  return Array.from(ids);
}
