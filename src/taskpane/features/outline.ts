/**
 * Reading pasted text as a deck.
 *
 * Most decks do not start in a spreadsheet. They start as notes in an email, a
 * list in a document, a few lines somebody typed in a hurry - and the work is
 * not deciding what to say, it is the half hour of making slides, one text box
 * at a time.
 *
 * So: paste it, and MEx works out the shape. That means reading structure the
 * way a person does. A short line with a list under it is a heading. A blank
 * line ends a thought. A line beginning with a dash is a bullet, and one
 * indented under it is a sub-bullet. A run of tab-separated lines is a table
 * that came out of a spreadsheet. Numbered lists stay numbered.
 *
 * None of this involves a model. They are the conventions people already use
 * when they write notes, applied consistently - which means the same paste
 * always gives the same deck, and when it reads something wrongly you can see
 * why and change one line to fix it.
 *
 * Everything here is pure and unit tested in tests/outline.test.ts.
 */

export interface OutlineBullet {
  text: string;
  /** 0 for a top-level bullet, 1 for one indented under it. */
  level: number;
  /** The number for an ordered list item, or null for a plain bullet. */
  number: number | null;
}

export interface OutlineTable {
  headers: string[];
  rows: string[][];
}

export interface OutlineSlide {
  title: string;
  bullets: OutlineBullet[];
  table: OutlineTable | null;
  /** True when this slide carries on a list that was too long for one. */
  continued: boolean;
}

/** Past this many bullets a slide is unreadable, so the list carries on. */
export const MAX_BULLETS_PER_SLIDE = 6;

/** A heading longer than this is really a sentence. */
const MAX_TITLE_LENGTH = 80;

/** A table needs at least this many rows before it is worth treating as one. */
const MIN_TABLE_ROWS = 2;

const BULLET_MARKER = /^[-*•–—○▪·+]\s+/;
const ORDERED_MARKER = /^(\d{1,2})[.)]\s+/;
const HEADING_MARKER = /^(#{1,6})\s+/;

interface Line {
  raw: string;
  text: string;
  indent: number;
  bullet: boolean;
  ordered: number | null;
  heading: number;
  blank: boolean;
  cells: string[] | null;
}

/** How deeply a line is indented, counting a tab as four spaces. */
function indentOf(raw: string): number {
  const leading = /^[ \t]*/.exec(raw)?.[0] ?? "";
  let width = 0;
  for (const character of leading) width += character === "\t" ? 4 : 1;
  return width;
}

/**
 * Splits a line into table cells, when it looks like one row of a table.
 *
 * Tabs are the reliable signal, because that is what a spreadsheet puts on the
 * clipboard. Commas are not: "Widgets, gadgets and sprockets" is one sentence,
 * not three cells, so a comma only counts when several lines agree on how many
 * fields there are.
 */
function cellsOf(raw: string): string[] | null {
  const text = raw.trim();
  if (text === "") return null;
  if (text.includes("\t")) {
    return text.split("\t").map((cell) => cell.trim());
  }
  // A pipe table, as written in Markdown and chat.
  if (/^\|.*\|$/.test(text)) {
    const cells = text
      .slice(1, -1)
      .split("|")
      .map((cell) => cell.trim());
    return cells.length >= 2 ? cells : null;
  }
  return null;
}

function readLine(raw: string): Line {
  const withoutIndent = raw.replace(/^[ \t]+/, "");
  const heading = HEADING_MARKER.exec(withoutIndent);
  const bullet = BULLET_MARKER.test(withoutIndent);
  const ordered = ORDERED_MARKER.exec(withoutIndent);

  let text = withoutIndent;
  if (heading) text = withoutIndent.slice(heading[0].length);
  else if (bullet) text = withoutIndent.replace(BULLET_MARKER, "");
  else if (ordered) text = withoutIndent.slice(ordered[0].length);

  return {
    raw,
    text: text.trim(),
    indent: indentOf(raw),
    bullet,
    ordered: ordered ? Number(ordered[1]) : null,
    heading: heading ? heading[1].length : 0,
    blank: raw.trim() === "",
    cells: cellsOf(raw),
  };
}

/**
 * True when a line reads like a heading rather than a sentence: short, not
 * punctuated like prose, and not a bullet.
 */
function looksLikeTitle(line: Line): boolean {
  if (line.blank || line.bullet || line.ordered !== null) return false;
  // A row of a table is short and unpunctuated, which is exactly what a title
  // looks like. It is never one.
  if (line.cells !== null) return false;
  if (line.heading > 0) return true;
  if (line.text.length > MAX_TITLE_LENGTH) return false;
  // "Revenue is up 14% this quarter." is a statement; "Revenue" is a heading.
  if (/[.!?]$/.test(line.text)) return false;
  // "Next steps:" is a heading with a colon, which people write constantly.
  return true;
}

/** Strips a trailing colon, which belongs to the notes and not to the slide. */
function cleanTitle(text: string): string {
  return text.replace(/\s*:\s*$/, "").trim();
}

interface Block {
  titleLine: Line | null;
  body: Line[];
}

/** Groups lines into blocks, each of which becomes one slide (or more). */
function toBlocks(lines: Line[]): Block[] {
  const blocks: Block[] = [];
  let current: Block | null = null;

  const start = (titleLine: Line | null) => {
    current = { titleLine, body: [] };
    blocks.push(current);
  };

  lines.forEach((line, index) => {
    if (line.blank) return;

    // An explicit heading always starts a slide.
    if (line.heading > 0) {
      start(line);
      return;
    }

    const next = lines.slice(index + 1).find((candidate) => !candidate.blank);
    const nextStartsList = Boolean(next && (next.bullet || next.ordered !== null || next.cells));
    const followsBlank = index > 0 && lines[index - 1].blank;
    const atStart = current === null;
    // A line standing on its own between blanks is a section marker: a slide
    // with a title and nothing on it, which is a normal thing to want.
    const standsAlone = index + 1 >= lines.length || lines[index + 1].blank;

    // A short, unpunctuated line with a list under it is what people write as
    // a slide title, whether or not they mark it as one.
    if (
      looksLikeTitle(line) &&
      (nextStartsList || standsAlone) &&
      (atStart || followsBlank || current?.body.length)
    ) {
      start(line);
      return;
    }

    if (!current) start(null);
    current!.body.push(line);
  });

  return blocks.filter((block) => block.titleLine !== null || block.body.length > 0);
}

/**
 * Reads consecutive table rows out of a block's body.
 *
 * The first row is treated as headings when the rows below it hold numbers and
 * it does not - the same rule the spreadsheet side uses, and for the same
 * reason: a heading row that gets charted as data is the classic silent error.
 */
function readTable(body: Line[]): { table: OutlineTable | null; rest: Line[] } {
  const rows = body.filter((line) => line.cells !== null);
  if (rows.length < MIN_TABLE_ROWS) return { table: null, rest: body };

  const width = Math.max(...rows.map((line) => line.cells!.length));
  // Markdown writes a row of dashes under the headings; it is punctuation.
  const cells = rows
    .map((line) => line.cells!)
    .filter((row) => !row.every((cell) => /^:?-{2,}:?$/.test(cell)))
    .map((row) => {
      const padded = [...row];
      while (padded.length < width) padded.push("");
      return padded;
    });
  if (cells.length < MIN_TABLE_ROWS) return { table: null, rest: body };

  const isNumeric = (value: string) =>
    value !== "" && !Number.isNaN(Number(value.replace(/[,%$£€\s]/g, "")));
  const headerRow = cells[0];
  const bodyRows = cells.slice(1);
  const headerLooksLikeLabels = headerRow.some((cell) => cell !== "" && !isNumeric(cell));
  const bodyHasNumbers = bodyRows.some((row) => row.some(isNumeric));

  const table: OutlineTable =
    headerLooksLikeLabels && (bodyHasNumbers || bodyRows.length > 0)
      ? { headers: headerRow, rows: bodyRows }
      : { headers: headerRow.map((_cell, index) => `Column ${index + 1}`), rows: cells };

  return { table, rest: body.filter((line) => line.cells === null) };
}

/** The indent widths used in a block, so sub-bullets can be spotted relatively. */
function baseIndent(body: Line[]): number {
  const indents = body.filter((line) => !line.blank).map((line) => line.indent);
  return indents.length > 0 ? Math.min(...indents) : 0;
}

function toBullets(body: Line[]): OutlineBullet[] {
  const base = baseIndent(body);
  return body
    .filter((line) => line.text !== "")
    .map((line) => ({
      text: line.text,
      // Anything indented past the block's own margin is a sub-point.
      level: line.indent > base + 1 ? 1 : 0,
      number: line.ordered,
    }));
}

export interface OutlineOptions {
  /** Used when the notes never name the first slide. */
  fallbackTitle?: string;
  maxBulletsPerSlide?: number;
}

/**
 * Turns pasted text into slides. Returns an empty list for text with nothing
 * in it, rather than one empty slide.
 */
export function parseOutline(text: string, options: OutlineOptions = {}): OutlineSlide[] {
  const limit = Math.max(1, options.maxBulletsPerSlide ?? MAX_BULLETS_PER_SLIDE);
  const lines = text.replace(/\r\n?/g, "\n").split("\n").map(readLine);
  if (lines.every((line) => line.blank)) return [];

  const slides: OutlineSlide[] = [];

  for (const block of toBlocks(lines)) {
    const { table, rest } = readTable(block.body);
    const bullets = toBullets(rest);
    const title = block.titleLine
      ? cleanTitle(block.titleLine.text)
      : deriveTitle(bullets, options.fallbackTitle);

    // A block with a title and nothing else is a section marker, which is a
    // perfectly good slide on its own.
    if (bullets.length === 0) {
      slides.push({ title, bullets: [], table, continued: false });
      continue;
    }

    // A table belongs with the first slide of the block; the bullets carry on
    // after it rather than being crammed alongside.
    let first = true;
    for (let at = 0; at < bullets.length; at += limit) {
      slides.push({
        title,
        bullets: bullets.slice(at, at + limit),
        table: first ? table : null,
        continued: !first,
      });
      first = false;
    }
  }

  return slides;
}

/**
 * A title for a block that never had one: the first bullet, shortened, which
 * is nearly always what the writer would have called it.
 */
function deriveTitle(bullets: readonly OutlineBullet[], fallback?: string): string {
  const first = bullets[0]?.text ?? "";
  if (first === "") return fallback ?? "Notes";
  const sentence = first.split(/(?<=[.!?])\s/)[0];
  // A heading does not end in a full stop, even when the sentence it came
  // from did.
  const trimmed = cleanTitle(sentence).replace(/[.!?]+$/, "");
  if (trimmed.length <= MAX_TITLE_LENGTH) return trimmed;
  return `${trimmed.slice(0, MAX_TITLE_LENGTH - 1).trimEnd()}...`;
}

/** One line describing what was read, for the pane. */
export function describeOutline(slides: readonly OutlineSlide[]): string {
  if (slides.length === 0) return "There was nothing in that to put on a slide.";
  const bullets = slides.reduce((total, slide) => total + slide.bullets.length, 0);
  const tables = slides.filter((slide) => slide.table).length;
  const parts = [`${slides.length} ${slides.length === 1 ? "slide" : "slides"}`];
  if (bullets > 0) parts.push(`${bullets} ${bullets === 1 ? "point" : "points"}`);
  if (tables > 0) parts.push(`${tables} ${tables === 1 ? "table" : "tables"}`);
  return `Read ${parts.join(", ")}.`;
}

/**
 * A table of labels and numbers can be charted instead of listed.
 *
 * Only offered when there is one obvious label column and at least one column
 * that is entirely numeric - anything less certain would produce a chart of
 * something nobody asked about.
 */
export function chartableColumns(
  table: OutlineTable
): { labelColumn: number; valueColumns: number[] } | null {
  if (table.rows.length < 2) return null;
  const width = table.headers.length;
  const numeric: number[] = [];
  const labels: number[] = [];

  for (let column = 0; column < width; column += 1) {
    const values = table.rows
      .map((row) => (row[column] ?? "").trim())
      .filter((value) => value !== "");
    if (values.length === 0) continue;
    const allNumbers = values.every(
      (value) => !Number.isNaN(Number(value.replace(/[,%$£€\s]/g, "")))
    );
    if (allNumbers) numeric.push(column);
    else labels.push(column);
  }

  if (numeric.length === 0 || labels.length === 0) return null;
  return { labelColumn: labels[0], valueColumns: numeric };
}

/** The numbers of a chartable table, as the grid the chart drawing expects. */
export function tableToGrid(
  table: OutlineTable,
  columns: { labelColumn: number; valueColumns: number[] }
) {
  const headers = [
    table.headers[columns.labelColumn] ?? "Item",
    ...columns.valueColumns.map((column) => table.headers[column] ?? "Value"),
  ];
  const rows = table.rows.map((row) => [
    row[columns.labelColumn] ?? "",
    ...columns.valueColumns.map(
      (column) => Number((row[column] ?? "0").replace(/[,%$£€\s]/g, "")) || 0
    ),
  ]);
  return [headers, ...rows];
}
