/**
 * Reading a formula back in English.
 *
 * Inheriting a sheet means inheriting its formulas, and a nested IF wrapped in an
 * IFERROR around a VLOOKUP is genuinely hard to read even for the person who
 * wrote it. Excel's own answer is the formula bar and Evaluate Formula, which
 * show you the same thing again, slowly.
 *
 * So: parse the formula properly and describe it. A real parser rather than a
 * pile of regexes, because formulas nest, and a description that is subtly wrong
 * about which branch does what is worse than no description at all.
 *
 * The parser covers what business sheets actually contain: numbers, text,
 * booleans, cell and range references (including sheet-qualified and absolute
 * ones), function calls, the arithmetic and comparison operators, string
 * concatenation with &, and percent and unary minus. Anything it cannot read is
 * reported as unreadable rather than guessed at.
 *
 * Everything here is pure and unit tested in tests/explain.test.ts.
 */

// ---------------------------------------------------------------------------
// Tokens
// ---------------------------------------------------------------------------

type TokenKind = "number" | "string" | "ref" | "name" | "operator" | "open" | "close" | "comma";

interface Token {
  kind: TokenKind;
  text: string;
}

/** A1, $B$7, Sheet2!A1, 'My Sheet'!A1:C9, B:B, 3:3 */
const REFERENCE =
  /^((?:'[^']+'|[A-Za-z_][A-Za-z0-9_. ]*)!)?(\$?[A-Za-z]{1,3}\$?\d{1,7}(?::\$?[A-Za-z]{1,3}\$?\d{1,7})?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d{1,7}:\$?\d{1,7})/;

const OPERATORS = ["<=", ">=", "<>", "+", "-", "*", "/", "^", "&", "=", "<", ">"];

export class FormulaError extends Error {}

function tokenize(input: string): Token[] {
  let text = input.trim();
  if (text.startsWith("=")) text = text.slice(1);

  const tokens: Token[] = [];
  let at = 0;

  while (at < text.length) {
    const character = text[at];

    if (/\s/.test(character)) {
      at += 1;
      continue;
    }

    if (character === '"') {
      let closed = false;
      let value = "";
      at += 1;
      while (at < text.length) {
        if (text[at] === '"') {
          // "" inside a string is one quote character.
          if (text[at + 1] === '"') {
            value += '"';
            at += 2;
            continue;
          }
          at += 1;
          closed = true;
          break;
        }
        value += text[at];
        at += 1;
      }
      if (!closed) throw new FormulaError("a piece of text is missing its closing quote");
      tokens.push({ kind: "string", text: value });
      continue;
    }

    if (character === "(") {
      tokens.push({ kind: "open", text: "(" });
      at += 1;
      continue;
    }
    if (character === ")") {
      tokens.push({ kind: "close", text: ")" });
      at += 1;
      continue;
    }
    if (character === "," || character === ";") {
      tokens.push({ kind: "comma", text: "," });
      at += 1;
      continue;
    }

    const rest = text.slice(at);

    // A reference has to be tried before a plain name, or "A1" reads as a name.
    const reference = REFERENCE.exec(rest);
    if (reference && reference.index === 0) {
      // "SUM(" looks like nothing of the sort, but "LOG10(" starts with a
      // letter-digit pair; only treat it as a reference when no "(" follows.
      const after = rest.slice(reference[0].length);
      if (!after.startsWith("(")) {
        tokens.push({ kind: "ref", text: reference[0] });
        at += reference[0].length;
        continue;
      }
    }

    const number = /^\d+(\.\d+)?(e[+-]?\d+)?|^\.\d+/i.exec(rest);
    if (number) {
      tokens.push({ kind: "number", text: number[0] });
      at += number[0].length;
      continue;
    }

    const name = /^[A-Za-z_][A-Za-z0-9_.]*/.exec(rest);
    if (name) {
      tokens.push({ kind: "name", text: name[0] });
      at += name[0].length;
      continue;
    }

    const operator = OPERATORS.find((candidate) => rest.startsWith(candidate));
    if (operator) {
      tokens.push({ kind: "operator", text: operator });
      at += operator.length;
      continue;
    }

    if (character === "%") {
      tokens.push({ kind: "operator", text: "%" });
      at += 1;
      continue;
    }

    throw new FormulaError(`it contains a character I can't read: ${character}`);
  }

  return tokens;
}

// ---------------------------------------------------------------------------
// The tree
// ---------------------------------------------------------------------------

export type Node =
  | { kind: "number"; value: number }
  | { kind: "string"; value: string }
  | { kind: "boolean"; value: boolean }
  | { kind: "ref"; text: string }
  | { kind: "name"; text: string }
  | { kind: "call"; name: string; args: Node[] }
  | { kind: "binary"; operator: string; left: Node; right: Node }
  | { kind: "unary"; operator: string; operand: Node }
  | { kind: "percent"; operand: Node };

/** Binding power, loosest first. Comparison binds least, ^ most. */
const PRECEDENCE: Record<string, number> = {
  "=": 1,
  "<": 1,
  ">": 1,
  "<=": 1,
  ">=": 1,
  "<>": 1,
  "&": 2,
  "+": 3,
  "-": 3,
  "*": 4,
  "/": 4,
  "^": 5,
};

class Parser {
  private at = 0;

  constructor(private readonly tokens: Token[]) {}

  parse(): Node {
    if (this.tokens.length === 0) throw new FormulaError("there is nothing in it");
    const node = this.expression(0);
    if (this.at < this.tokens.length) {
      throw new FormulaError(
        `I got as far as "${this.tokens[this.at].text}" and then lost the thread`
      );
    }
    return node;
  }

  private peek(): Token | undefined {
    return this.tokens[this.at];
  }

  private expression(minimum: number): Node {
    let left = this.unary();
    for (;;) {
      const token = this.peek();
      if (!token || token.kind !== "operator") break;
      const precedence = PRECEDENCE[token.text];
      if (precedence === undefined || precedence < minimum) break;
      this.at += 1;
      // ^ is right-associative; everything else groups leftwards.
      const right = this.expression(token.text === "^" ? precedence : precedence + 1);
      left = { kind: "binary", operator: token.text, left, right };
    }
    return left;
  }

  private unary(): Node {
    const token = this.peek();
    if (token && token.kind === "operator" && (token.text === "-" || token.text === "+")) {
      this.at += 1;
      return { kind: "unary", operator: token.text, operand: this.unary() };
    }
    return this.postfix();
  }

  private postfix(): Node {
    let node = this.primary();
    for (;;) {
      const token = this.peek();
      if (token && token.kind === "operator" && token.text === "%") {
        this.at += 1;
        node = { kind: "percent", operand: node };
        continue;
      }
      break;
    }
    return node;
  }

  private primary(): Node {
    const token = this.peek();
    if (!token) throw new FormulaError("it stops in the middle");

    if (token.kind === "number") {
      this.at += 1;
      return { kind: "number", value: Number(token.text) };
    }
    if (token.kind === "string") {
      this.at += 1;
      return { kind: "string", value: token.text };
    }
    if (token.kind === "ref") {
      this.at += 1;
      return { kind: "ref", text: token.text };
    }
    if (token.kind === "open") {
      this.at += 1;
      const inner = this.expression(0);
      this.expect("close");
      return inner;
    }
    if (token.kind === "name") {
      this.at += 1;
      const next = this.peek();
      if (next && next.kind === "open") {
        this.at += 1;
        const args: Node[] = [];
        if (this.peek()?.kind !== "close") {
          for (;;) {
            args.push(this.expression(0));
            if (this.peek()?.kind === "comma") {
              this.at += 1;
              continue;
            }
            break;
          }
        }
        this.expect("close");
        return { kind: "call", name: token.text.toUpperCase(), args };
      }
      const upper = token.text.toUpperCase();
      if (upper === "TRUE" || upper === "FALSE") {
        return { kind: "boolean", value: upper === "TRUE" };
      }
      return { kind: "name", text: token.text };
    }

    throw new FormulaError(`I did not expect "${token.text}" there`);
  }

  private expect(kind: TokenKind): void {
    if (this.peek()?.kind !== kind) {
      throw new FormulaError(
        kind === "close" ? "a bracket is not closed" : "it is missing a piece"
      );
    }
    this.at += 1;
  }
}

export function parseFormula(formula: string): Node {
  return new Parser(tokenize(formula)).parse();
}

// ---------------------------------------------------------------------------
// Describing it
// ---------------------------------------------------------------------------

const OPERATOR_WORDS: Record<string, string> = {
  "+": "plus",
  "-": "minus",
  "*": "times",
  "/": "divided by",
  "^": "to the power of",
  "&": "joined to",
  "=": "is",
  "<": "is less than",
  ">": "is greater than",
  "<=": "is at most",
  ">=": "is at least",
  "<>": "is not",
};

/** How to read a handful of functions whose plain name explains nothing. */
const FUNCTION_PHRASES: Record<string, (args: string[]) => string> = {
  SUM: (args) => `the total of ${list(args)}`,
  SUMIF: (args) =>
    `the total of ${args[2] ?? args[0]} for rows where ${args[0]} matches ${args[1]}`,
  SUMIFS: (args) => `the total of ${args[0]} for rows matching ${list(args.slice(1))}`,
  AVERAGE: (args) => `the average of ${list(args)}`,
  AVERAGEIF: (args) => `the average of ${args[2] ?? args[0]} where ${args[0]} matches ${args[1]}`,
  COUNT: (args) => `how many numbers are in ${list(args)}`,
  COUNTA: (args) => `how many cells in ${list(args)} are not empty`,
  COUNTIF: (args) => `how many cells in ${args[0]} match ${args[1]}`,
  COUNTIFS: (args) => `how many rows match ${list(args)}`,
  MIN: (args) => `the smallest of ${list(args)}`,
  MAX: (args) => `the largest of ${list(args)}`,
  ROUND: (args) => `${args[0]} rounded to ${args[1]} decimal places`,
  ROUNDUP: (args) => `${args[0]} rounded up to ${args[1]} decimal places`,
  ROUNDDOWN: (args) => `${args[0]} rounded down to ${args[1]} decimal places`,
  ABS: (args) => `${args[0]} without its sign`,
  IF: (args) => `if ${args[0]}, then ${args[1] ?? "TRUE"}, otherwise ${args[2] ?? "FALSE"}`,
  IFS: (args) => `the first of these that holds: ${list(pairUp(args))}`,
  IFERROR: (args) => `${args[0]}, or ${args[1]} if that gives an error`,
  IFNA: (args) => `${args[0]}, or ${args[1]} if that is not found`,
  AND: (args) => `${list(args, "and")} are all true`,
  OR: (args) => `any of ${list(args, "or")} is true`,
  NOT: (args) => `the opposite of ${args[0]}`,
  VLOOKUP: (args) =>
    `${args[0]} looked up in the first column of ${args[1]}, returning column ${args[2]}${exactness(args[3])}`,
  HLOOKUP: (args) =>
    `${args[0]} looked up along the first row of ${args[1]}, returning row ${args[2]}${exactness(args[3])}`,
  XLOOKUP: (args) =>
    `${args[0]} looked up in ${args[1]}, returning the matching value from ${args[2]}${
      args[3] ? `, or ${args[3]} when there is no match` : ""
    }`,
  INDEX: (args) =>
    args.length >= 3
      ? `the cell in ${args[0]} at row ${args[1]}, column ${args[2]}`
      : `item ${args[1]} of ${args[0]}`,
  MATCH: (args) => `the position of ${args[0]} within ${args[1]}`,
  TEXT: (args) => `${args[0]} written as text in the format ${args[1]}`,
  VALUE: (args) => `${args[0]} read as a number`,
  LEFT: (args) => `the first ${args[1] ?? "1"} characters of ${args[0]}`,
  RIGHT: (args) => `the last ${args[1] ?? "1"} characters of ${args[0]}`,
  MID: (args) => `${args[1]} characters of ${args[0]} starting at character ${args[1]}`,
  LEN: (args) => `how many characters are in ${args[0]}`,
  TRIM: (args) => `${args[0]} with extra spaces removed`,
  UPPER: (args) => `${args[0]} in capitals`,
  LOWER: (args) => `${args[0]} in lower case`,
  PROPER: (args) => `${args[0]} with each word capitalised`,
  CONCAT: (args) => `${list(args)} joined together`,
  CONCATENATE: (args) => `${list(args)} joined together`,
  TEXTJOIN: (args) => `${list(args.slice(2))} joined with ${args[0]}`,
  TODAY: () => "today's date",
  NOW: () => "the current date and time",
  YEAR: (args) => `the year of ${args[0]}`,
  MONTH: (args) => `the month of ${args[0]}`,
  DAY: (args) => `the day of ${args[0]}`,
  EOMONTH: (args) => `the last day of the month ${args[1]} months from ${args[0]}`,
  SUBTOTAL: (args) => `a subtotal of ${list(args.slice(1))} that skips hidden rows`,
  ROW: (args) => (args.length === 0 ? "this row's number" : `the row number of ${args[0]}`),
  COLUMN: (args) =>
    args.length === 0 ? "this column's number" : `the column number of ${args[0]}`,
  ISBLANK: (args) => `${args[0]} is empty`,
  ISNUMBER: (args) => `${args[0]} is a number`,
  ISTEXT: (args) => `${args[0]} is text`,
  ISERROR: (args) => `${args[0]} is an error`,
};

function exactness(argument: string | undefined): string {
  if (argument === undefined) return "";
  if (/^(FALSE|0)$/i.test(argument)) return ", requiring an exact match";
  if (/^(TRUE|1)$/i.test(argument)) return ", accepting the nearest match below";
  return `, matching by ${argument}`;
}

function pairUp(args: string[]): string[] {
  const out: string[] = [];
  for (let at = 0; at + 1 < args.length; at += 2) out.push(`${args[at]} then ${args[at + 1]}`);
  if (args.length % 2 === 1) out.push(args[args.length - 1]);
  return out;
}

function list(parts: string[], conjunction = "and"): string {
  if (parts.length === 0) return "nothing";
  if (parts.length === 1) return parts[0];
  return `${parts.slice(0, -1).join(", ")} ${conjunction} ${parts[parts.length - 1]}`;
}

function quote(value: string): string {
  return value === "" ? "an empty cell" : `"${value}"`;
}

/** Describes one node. Nested calls read as phrases, so the whole thing composes. */
function describe(node: Node): string {
  switch (node.kind) {
    case "number":
      return String(node.value);
    case "string":
      return quote(node.value);
    case "boolean":
      return node.value ? "true" : "false";
    case "ref":
      return node.text;
    case "name":
      return node.text;
    case "percent":
      return `${describe(node.operand)}%`;
    case "unary":
      return node.operator === "-" ? `minus ${describe(node.operand)}` : describe(node.operand);
    case "binary": {
      const word = OPERATOR_WORDS[node.operator] ?? node.operator;
      return `${describe(node.left)} ${word} ${describe(node.right)}`;
    }
    case "call": {
      const args = node.args.map(describe);
      const phrase = FUNCTION_PHRASES[node.name];
      if (phrase) return phrase(args);
      return args.length === 0 ? `${node.name}()` : `${node.name} of ${list(args)}`;
    }
  }
}

export type WarningKind =
  | "hardcodedNumber"
  | "wholeColumn"
  | "deepNesting"
  | "volatile"
  | "divideRisk"
  | "lookupByPosition";

export interface FormulaWarning {
  kind: WarningKind;
  text: string;
}

export interface Explanation {
  /** One sentence, starting with a capital and ending with a full stop. */
  english: string;
  /** Cells and ranges the formula reads, in the order they appear. */
  references: string[];
  /** Functions used, uppercased, in the order they appear. */
  functions: string[];
  /** Things worth knowing about how it is written. */
  warnings: FormulaWarning[];
}

const VOLATILE = new Set(["NOW", "TODAY", "RAND", "RANDBETWEEN", "OFFSET", "INDIRECT"]);

function walk(node: Node, visit: (node: Node, depth: number) => void, depth = 0): void {
  visit(node, depth);
  switch (node.kind) {
    case "call":
      for (const argument of node.args) walk(argument, visit, depth + 1);
      break;
    case "binary":
      walk(node.left, visit, depth + 1);
      walk(node.right, visit, depth + 1);
      break;
    case "unary":
      walk(node.operand, visit, depth + 1);
      break;
    case "percent":
      walk(node.operand, visit, depth + 1);
      break;
    default:
      break;
  }
}

function collectWarnings(root: Node): FormulaWarning[] {
  const warnings: FormulaWarning[] = [];
  const functions: string[] = [];
  let ifDepth = 0;
  let hardcoded = 0;
  const wholeColumns: string[] = [];
  let divides = false;

  walk(root, (node, depth) => {
    if (node.kind === "call") {
      functions.push(node.name);
      if (node.name === "IF") ifDepth = Math.max(ifDepth, countNested(node, "IF"));
      if (VOLATILE.has(node.name)) {
        warnings.push({
          kind: "volatile",
          text: `${node.name} changes on its own, so this cell recalculates constantly.`,
        });
      }
      if ((node.name === "VLOOKUP" || node.name === "HLOOKUP") && node.args[2]) {
        warnings.push({
          kind: "lookupByPosition",
          text: `${node.name} finds its answer by counting columns, so inserting a column silently changes the result. XLOOKUP or INDEX/MATCH name the column instead.`,
        });
      }
    }
    if (node.kind === "ref" && /^[^!]*!?\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}$/.test(node.text)) {
      wholeColumns.push(node.text);
    }
    // A number typed into a formula is a number nobody can find later. Small
    // ones are usually structural (column 3, 0 decimal places), so only flag
    // numbers that look like amounts or rates.
    if (
      node.kind === "number" &&
      depth > 0 &&
      (node.value >= 100 || !Number.isInteger(node.value))
    ) {
      hardcoded += 1;
    }
    if (node.kind === "binary" && node.operator === "/") divides = true;
  });

  if (hardcoded > 0) {
    warnings.push({
      kind: "hardcodedNumber",
      text: `${hardcoded === 1 ? "A number is" : `${hardcoded} numbers are`} typed into the formula itself. If one of those is a rate or a price, it is invisible to anyone reading the sheet.`,
    });
  }
  if (wholeColumns.length > 0) {
    warnings.push({
      kind: "wholeColumn",
      text: `${wholeColumns.join(", ")} covers a whole column, which is slower than it needs to be and will pick up anything typed below the data.`,
    });
  }
  if (ifDepth >= 3) {
    warnings.push({
      kind: "deepNesting",
      text: `IF is nested ${ifDepth} deep. IFS, or a small lookup table, would be easier to check.`,
    });
  }
  if (divides && !functions.includes("IFERROR")) {
    warnings.push({
      kind: "divideRisk",
      text: "It divides without a guard, so an empty or zero divisor shows #DIV/0!.",
    });
  }
  return warnings;
}

function countNested(node: Node, name: string): number {
  if (node.kind !== "call" || node.name !== name) return 0;
  let deepest = 0;
  for (const argument of node.args) deepest = Math.max(deepest, countNested(argument, name));
  return 1 + deepest;
}

/**
 * Explains a formula, or says why it can't. Never throws: an unreadable formula
 * comes back as an explanation that admits as much, because the caller is a chat
 * message and "I can't read that" is a perfectly good answer.
 */
export function explainFormula(formula: string): Explanation {
  const text = formula.trim();
  if (text === "" || !text.startsWith("=")) {
    return {
      english:
        text === "" ? "That cell is empty." : "That cell holds a typed-in value, not a formula.",
      references: [],
      functions: [],
      warnings: [],
    };
  }

  let root: Node;
  try {
    root = parseFormula(text);
  } catch (error) {
    const why = error instanceof FormulaError ? error.message : "I could not read it";
    return {
      english: `I can't read that formula - ${why}.`,
      references: [],
      functions: [],
      warnings: [],
    };
  }

  const references: string[] = [];
  const functions: string[] = [];
  walk(root, (node) => {
    if (node.kind === "ref" && !references.includes(node.text)) references.push(node.text);
    if (node.kind === "call" && !functions.includes(node.name)) functions.push(node.name);
  });

  const body = describe(root);
  const english = `${body.charAt(0).toUpperCase()}${body.slice(1)}.`;

  return { english, references, functions, warnings: collectWarnings(root) };
}
