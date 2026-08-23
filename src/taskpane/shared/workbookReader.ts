/**
 * Reads .xlsx / .xlsm / .csv files the user picks from disk.
 *
 * Office.js can only see the workbook the add-in is running in, so merging
 * *other* files means parsing them ourselves. An .xlsx is a zip of XML parts, so
 * this module unzips with fflate and walks the XML with the browser's own
 * DOMParser - no heavyweight spreadsheet library, no server round-trip, and
 * nothing leaves the machine.
 *
 * Values come back alongside their Excel number-format codes so a merge can
 * preserve dates, currency and percentages instead of flattening them to text.
 */

import { unzipSync } from "fflate";

import { columnIndexFromLetter, padRow } from "./excelHelpers";
import { CellValue, Grid } from "./types";

export interface SheetData {
  name: string;
  grid: Grid;
  /** Excel number-format code per cell, parallel to `grid`. "General" when unstyled. */
  numberFormats: string[][];
}

export interface WorkbookData {
  fileName: string;
  sheets: SheetData[];
}

/** Guard against a pathological file exhausting the task pane's memory. */
const MAX_ROWS_PER_SHEET = 200000;

// ---------------------------------------------------------------------------
// Number formats
// ---------------------------------------------------------------------------

/** The built-in number formats an .xlsx refers to by id instead of spelling out. */
const BUILT_IN_NUMBER_FORMATS: Record<number, string> = {
  0: "General",
  1: "0",
  2: "0.00",
  3: "#,##0",
  4: "#,##0.00",
  9: "0%",
  10: "0.00%",
  11: "0.00E+00",
  12: "# ?/?",
  13: "# ??/??",
  14: "mm-dd-yy",
  15: "d-mmm-yy",
  16: "d-mmm",
  17: "mmm-yy",
  18: "h:mm AM/PM",
  19: "h:mm:ss AM/PM",
  20: "h:mm",
  21: "h:mm:ss",
  22: "m/d/yy h:mm",
  37: "#,##0 ;(#,##0)",
  38: "#,##0 ;[Red](#,##0)",
  39: "#,##0.00;(#,##0.00)",
  40: "#,##0.00;[Red](#,##0.00)",
  45: "mm:ss",
  46: "[h]:mm:ss",
  47: "mmss.0",
  48: "##0.0E+0",
  49: "@",
};

// ---------------------------------------------------------------------------
// Delimited text
// ---------------------------------------------------------------------------

/** Convert a CSV field to a number when it unambiguously is one. */
function coerceScalar(field: string): CellValue {
  const text = field.trim();
  if (text === "") {
    return null;
  }
  // Leading zeros and "+1 555..." are identifiers, not numbers - leave them as text.
  if (/^-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?$/.test(text)) {
    const value = Number(text);
    if (Number.isFinite(value)) {
      return value;
    }
  }
  return field;
}

/** Pick the delimiter that yields the most consistent column count. */
export function sniffDelimiter(sample: string): string {
  const candidates = [",", ";", "\t", "|"];
  let best = ",";
  let bestScore = -1;
  for (const candidate of candidates) {
    const counts = sample
      .split(/\r?\n/)
      .slice(0, 20)
      .filter((line) => line.trim() !== "")
      .map((line) => line.split(candidate).length);
    if (counts.length === 0 || counts[0] < 2) {
      continue;
    }
    const consistent = counts.filter((count) => count === counts[0]).length / counts.length;
    const score = consistent * counts[0];
    if (score > bestScore) {
      bestScore = score;
      best = candidate;
    }
  }
  return best;
}

/** RFC 4180 parser: handles quoted fields, escaped quotes and embedded newlines. */
export function parseDelimited(text: string, delimiter?: string): Grid {
  const content = text.replace(/^\uFEFF/, "");
  const separator = delimiter ?? sniffDelimiter(content);

  const rows: Grid = [];
  let row: CellValue[] = [];
  let field = "";
  let inQuotes = false;
  let fieldWasQuoted = false;

  const endField = () => {
    row.push(fieldWasQuoted ? field : coerceScalar(field));
    field = "";
    fieldWasQuoted = false;
  };
  const endRow = () => {
    endField();
    rows.push(row);
    row = [];
  };

  for (let i = 0; i < content.length; i += 1) {
    const char = content[i];

    if (inQuotes) {
      if (char === '"') {
        if (content[i + 1] === '"') {
          field += '"';
          i += 1;
        } else {
          inQuotes = false;
        }
      } else {
        field += char;
      }
      continue;
    }

    if (char === '"' && field === "") {
      inQuotes = true;
      fieldWasQuoted = true;
    } else if (char === separator) {
      endField();
    } else if (char === "\n") {
      endRow();
    } else if (char === "\r") {
      // Swallow CR; the following LF ends the row.
    } else {
      field += char;
    }
  }

  // Trailing content with no final newline is still a row.
  if (field !== "" || fieldWasQuoted || row.length > 0) {
    endRow();
  }

  // Drop a trailing empty row produced by a file that ends in a newline.
  while (rows.length > 0 && rows[rows.length - 1].every((cell) => cell === null)) {
    rows.pop();
  }

  const width = rows.reduce((max, current) => Math.max(max, current.length), 0);
  return rows.map((current) => padRow(current, width));
}

// ---------------------------------------------------------------------------
// XLSX
// ---------------------------------------------------------------------------

function parseXml(xml: string): Document {
  const document = new DOMParser().parseFromString(xml, "application/xml");
  const parseError = document.getElementsByTagName("parsererror")[0];
  if (parseError) {
    throw new Error("The file contains XML this add-in could not read.");
  }
  return document;
}

/** `getElementsByTagName` but tolerant of namespace prefixes. */
function elementsByLocalName(scope: Document | Element, localName: string): Element[] {
  const direct = Array.from(scope.getElementsByTagName(localName));
  if (direct.length > 0) {
    return direct;
  }
  return Array.from(scope.getElementsByTagName("*")).filter(
    (element) => element.localName === localName
  );
}

function decodeUtf8(bytes: Uint8Array): string {
  return new TextDecoder("utf-8").decode(bytes);
}

/** Shared strings are stored once and referenced by index from every cell that uses them. */
function readSharedStrings(files: Record<string, Uint8Array>): string[] {
  const part = files["xl/sharedStrings.xml"];
  if (!part) {
    return [];
  }
  const document = parseXml(decodeUtf8(part));
  return elementsByLocalName(document, "si").map((si) => {
    // Rich text splits one string across several <t> runs.
    return elementsByLocalName(si, "t")
      .filter((t) => t.parentElement?.localName !== "rPh")
      .map((t) => t.textContent ?? "")
      .join("");
  });
}

/** Map each cell-style index to its number-format code. */
function readStyleFormats(files: Record<string, Uint8Array>): string[] {
  const part = files["xl/styles.xml"];
  if (!part) {
    return [];
  }
  const document = parseXml(decodeUtf8(part));

  const custom: Record<number, string> = {};
  for (const numFmt of elementsByLocalName(document, "numFmt")) {
    const id = Number(numFmt.getAttribute("numFmtId"));
    const code = numFmt.getAttribute("formatCode");
    if (Number.isFinite(id) && code) {
      custom[id] = code;
    }
  }

  const cellXfs = elementsByLocalName(document, "cellXfs")[0];
  if (!cellXfs) {
    return [];
  }
  return elementsByLocalName(cellXfs, "xf").map((xf) => {
    const id = Number(xf.getAttribute("numFmtId") ?? 0);
    return custom[id] ?? BUILT_IN_NUMBER_FORMATS[id] ?? "General";
  });
}

/** Resolve each sheet's display name to the zip entry that holds it. */
function readSheetIndex(files: Record<string, Uint8Array>): Array<{ name: string; path: string }> {
  const workbookPart = files["xl/workbook.xml"];
  if (!workbookPart) {
    throw new Error("This does not look like an Excel workbook (no xl/workbook.xml inside).");
  }

  const relationships: Record<string, string> = {};
  const relsPart = files["xl/_rels/workbook.xml.rels"];
  if (relsPart) {
    for (const relationship of elementsByLocalName(
      parseXml(decodeUtf8(relsPart)),
      "Relationship"
    )) {
      const id = relationship.getAttribute("Id");
      const target = relationship.getAttribute("Target");
      if (id && target) {
        relationships[id] = target.replace(/^\//, "").replace(/^xl\//, "");
      }
    }
  }

  const workbook = parseXml(decodeUtf8(workbookPart));
  const sheets: Array<{ name: string; path: string }> = [];

  elementsByLocalName(workbook, "sheet").forEach((sheet, index) => {
    const name = sheet.getAttribute("name") ?? `Sheet${index + 1}`;
    // Hidden sheets are usually scratch space, not data the user means to merge.
    if ((sheet.getAttribute("state") ?? "visible") !== "visible") {
      return;
    }
    const relationshipId =
      sheet.getAttribute("r:id") ??
      sheet.getAttributeNS(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "id"
      );
    const target = relationshipId ? relationships[relationshipId] : undefined;
    const path = `xl/${target ?? `worksheets/sheet${index + 1}.xml`}`;
    if (files[path]) {
      sheets.push({ name, path });
    }
  });

  return sheets;
}

function readSheet(
  files: Record<string, Uint8Array>,
  entry: { name: string; path: string },
  sharedStrings: string[],
  styleFormats: string[]
): SheetData {
  const document = parseXml(decodeUtf8(files[entry.path]));

  const cells: Array<{ row: number; column: number; value: CellValue; format: string }> = [];
  let maxRow = -1;
  let maxColumn = -1;

  for (const rowElement of elementsByLocalName(document, "row")) {
    for (const cell of elementsByLocalName(rowElement, "c")) {
      const reference = cell.getAttribute("r");
      if (!reference) {
        continue;
      }
      const match = /^([A-Za-z]+)(\d+)$/.exec(reference);
      if (!match) {
        continue;
      }
      const rowIndex = Number(match[2]) - 1;
      const columnIndex = columnIndexFromLetter(match[1]);
      if (rowIndex >= MAX_ROWS_PER_SHEET) {
        continue;
      }

      const type = cell.getAttribute("t") ?? "n";
      const styleIndex = Number(cell.getAttribute("s") ?? -1);
      const format = styleFormats[styleIndex] ?? "General";
      const valueElement = elementsByLocalName(cell, "v")[0];
      const rawValue = valueElement?.textContent ?? "";

      let value: CellValue;
      switch (type) {
        case "s": {
          value = sharedStrings[Number(rawValue)] ?? null;
          break;
        }
        case "inlineStr": {
          const inline = elementsByLocalName(cell, "t");
          value = inline.map((t) => t.textContent ?? "").join("") || null;
          break;
        }
        case "b":
          value = rawValue === "1";
          break;
        case "e":
          // Error values (#N/A, #REF!) carry across as their text.
          value = rawValue || null;
          break;
        case "str":
          value = rawValue || null;
          break;
        default: {
          if (rawValue === "") {
            value = null;
          } else {
            const numeric = Number(rawValue);
            value = Number.isFinite(numeric) ? numeric : rawValue;
          }
        }
      }

      if (value === null && format === "General") {
        continue;
      }

      cells.push({ row: rowIndex, column: columnIndex, value, format });
      maxRow = Math.max(maxRow, rowIndex);
      maxColumn = Math.max(maxColumn, columnIndex);
    }
  }

  if (maxRow < 0 || maxColumn < 0) {
    return { name: entry.name, grid: [], numberFormats: [] };
  }

  const width = maxColumn + 1;
  const grid: Grid = Array.from({ length: maxRow + 1 }, () =>
    new Array<CellValue>(width).fill(null)
  );
  const numberFormats: string[][] = Array.from({ length: maxRow + 1 }, () =>
    new Array<string>(width).fill("General")
  );

  for (const cell of cells) {
    grid[cell.row][cell.column] = cell.value;
    numberFormats[cell.row][cell.column] = cell.format;
  }

  return { name: entry.name, grid, numberFormats };
}

function readXlsx(fileName: string, bytes: Uint8Array): WorkbookData {
  let files: Record<string, Uint8Array>;
  try {
    files = unzipSync(bytes);
  } catch {
    throw new Error(
      `"${fileName}" could not be opened. If it is an old .xls file, re-save it as .xlsx and try again.`
    );
  }

  const sharedStrings = readSharedStrings(files);
  const styleFormats = readStyleFormats(files);
  const sheets = readSheetIndex(files).map((entry) =>
    readSheet(files, entry, sharedStrings, styleFormats)
  );

  return { fileName, sheets };
}

// ---------------------------------------------------------------------------
// Entry point
// ---------------------------------------------------------------------------

/** File extensions the picker and this reader agree on. */
export const SUPPORTED_EXTENSIONS = [".xlsx", ".xlsm", ".csv", ".txt", ".tsv"];

export function isSupportedFile(fileName: string): boolean {
  return SUPPORTED_EXTENSIONS.some((extension) => fileName.toLowerCase().endsWith(extension));
}

/** Parse one picked file into sheets of values plus number formats. */
export async function readWorkbookFile(file: File): Promise<WorkbookData> {
  const name = file.name;
  const lower = name.toLowerCase();

  if (lower.endsWith(".xls")) {
    throw new Error(`"${name}" is the legacy .xls format. Re-save it as .xlsx and try again.`);
  }

  if (lower.endsWith(".csv") || lower.endsWith(".txt") || lower.endsWith(".tsv")) {
    const text = await file.text();
    const grid = parseDelimited(text, lower.endsWith(".tsv") ? "\t" : undefined);
    const numberFormats = grid.map((row) => row.map(() => "General"));
    // A delimited file is one nameless table; use the file name as the sheet name.
    return {
      fileName: name,
      sheets: [{ name: name.replace(/\.[^.]+$/, ""), grid, numberFormats }],
    };
  }

  if (!isSupportedFile(name)) {
    throw new Error(`"${name}" is not a supported file type (${SUPPORTED_EXTENSIONS.join(", ")}).`);
  }

  const buffer = await file.arrayBuffer();
  return readXlsx(name, new Uint8Array(buffer));
}
