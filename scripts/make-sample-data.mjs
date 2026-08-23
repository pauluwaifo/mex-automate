/**
 * Generates sample data for manually testing the add-in in Excel.
 *
 * Writes samples/MEx-Automate-Test.xlsx (four sheets, deliberately messy) plus
 * two CSVs for the "merge files from my computer" path.
 *
 * The workbook is assembled straight from OOXML parts and zipped with fflate -
 * the same approach src/taskpane/shared/workbookReader.ts uses to read files, in
 * reverse. That keeps this script dependency-free beyond what the add-in already
 * ships.
 *
 *   node scripts/make-sample-data.mjs
 */

import fs from "node:fs";
import path from "node:path";
import { zipSync, strToU8 } from "fflate";

const OUT_DIR = "samples";

// ---------------------------------------------------------------------------
// Minimal .xlsx writer
// ---------------------------------------------------------------------------

function escapeXml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function columnLetter(index) {
  let result = "";
  let n = index;
  while (n >= 0) {
    result = String.fromCharCode((n % 26) + 65) + result;
    n = Math.floor(n / 26) - 1;
  }
  return result;
}

/** A cell value of `{ f: "SUM(A1:A2)" }` is written as a formula. */
function cellXml(value, ref) {
  if (value === null || value === undefined || value === "") {
    return "";
  }
  if (typeof value === "object" && typeof value.f === "string") {
    return `<c r="${ref}"><f>${escapeXml(value.f)}</f></c>`;
  }
  if (typeof value === "number" && Number.isFinite(value)) {
    return `<c r="${ref}"><v>${value}</v></c>`;
  }
  // xml:space="preserve" matters here: several sample values carry the leading
  // and trailing spaces the Clean tab is meant to strip.
  return `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${escapeXml(value)}</t></is></c>`;
}

function sheetXml(grid) {
  const rows = grid
    .map((row, rowIndex) => {
      const cells = row
        .map((value, columnIndex) => cellXml(value, `${columnLetter(columnIndex)}${rowIndex + 1}`))
        .join("");
      return `<row r="${rowIndex + 1}">${cells}</row>`;
    })
    .join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${rows}</sheetData></worksheet>`;
}

function buildXlsx(sheets) {
  const files = {};

  files["[Content_Types].xml"] = strToU8(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${sheets
      .map(
        (_sheet, index) =>
          `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`
      )
      .join("")}</Types>`
  );

  files["_rels/.rels"] = strToU8(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`
  );

  files["xl/workbook.xml"] = strToU8(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${sheets
      .map(
        (sheet, index) =>
          `<sheet name="${escapeXml(sheet.name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`
      )
      .join("")}</sheets><calcPr calcId="0" fullCalcOnLoad="1"/></workbook>`
  );

  files["xl/_rels/workbook.xml.rels"] = strToU8(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets
      .map(
        (_sheet, index) =>
          `<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`
      )
      .join("")}</Relationships>`
  );

  sheets.forEach((sheet, index) => {
    files[`xl/worksheets/sheet${index + 1}.xml`] = strToU8(sheetXml(sheet.grid));
  });

  return zipSync(files, { level: 6 });
}

// ---------------------------------------------------------------------------
// The sample data
// ---------------------------------------------------------------------------

// Sheet 1: the "normal" one, but with trailing spaces, inconsistent case, three
// different text date formats, and one exact duplicate row (order 4).
const janOrders = [
  ["Order ID", "Order Date", "Customer", "Region", "Amount", "Units"],
  [1, "2024-01-05", "  Acme Ltd ", "North", 1200, 3],
  [2, "06/01/2024", "beta trading", "South", 850, 2],
  [3, "Jan 7, 2024", "ACME LTD", "north", 400, 1],
  [4, "2024-01-08", "Gamma Co", "East", 2200, 5],
  [5, "09/01/2024", "beta trading  ", "South", 975, 2],
  [4, "2024-01-08", "Gamma Co", "East", 2200, 5],
  [6, "10-Jan-2024", "Delta Supplies", "West", 310, 1],
  [7, "2024-01-11", "Acme Ltd", "North", 1450, 4],
];

// Sheet 2: same data, different column order and header spellings, plus a
// "Channel" column the other sheets do not have. One amount is stored as text.
const febOrders = [
  ["order_id", "AMOUNT", "Customer", "region", "order_date", "Channel"],
  [20, 1500, "Acme Ltd", "North", "2024-02-02", "Direct"],
  [21, "1,050.00", "Gamma Co", "East", "2024-02-05", "Partner"],
  [22, 640, "Delta Supplies", "West", "2024-02-09", "Direct"],
  [23, 2300, "Epsilon PLC", "South", "2024-02-14", "Partner"],
  [24, 480, "beta trading", "South", "2024-02-20", "Direct"],
];

// Sheet 3: uses "Client Name" instead of "Customer" - the case the alias box
// ("Client Name = Customer") exists for.
const marOrders = [
  ["Order ID", "Order Date", "Client Name", "Region", "Amount", "Units"],
  [40, "2024-03-01", "Acme Ltd", "North", 990, 2],
  [41, "2024-03-06", "Zeta Holdings", "East", 3100, 7],
  [42, "2024-03-12", "Gamma Co", "East", 780, 2],
  [43, "2024-03-19", "Delta Supplies", "West", 1220, 3],
];

// Sheet 4: a report template for the Reports tab. Row 3 is the header row and
// rows 4-6 are the data zone; column E holds a formula that a refresh must not
// overwrite, and row 8 is a total that sits outside the zone.
const salesReport = [
  ["Monthly Sales Report"],
  [],
  ["Customer", "Region", "Amount", "Units", "Amount per Unit"],
  ["Acme Ltd", "North", 1200, 3, { f: 'IF(D4=0,"",C4/D4)' }],
  ["Gamma Co", "East", 2200, 5, { f: 'IF(D5=0,"",C5/D5)' }],
  ["Delta Supplies", "West", 310, 1, { f: 'IF(D6=0,"",C6/D6)' }],
  [],
  ["Total", "", { f: "SUM(C4:C6)" }, { f: "SUM(D4:D6)" }, ""],
];

const aprCsv = `Order ID,Order Date,Customer,Region,Amount,Units
60,2024-04-03,Acme Ltd,North,1310,3
61,2024-04-08,Epsilon PLC,South,2750,6
62,2024-04-15,Gamma Co,East,690,2
63,2024-04-22,"Beta Trading, Ltd",South,1180,3
`;

const mayCsv = `order_id;order_date;Client Name;region;AMOUNT;Units
80;2024-05-02;Delta Supplies;West;1490;4
81;2024-05-11;Acme Ltd;North;2050;5
82;2024-05-19;Zeta Holdings;East;860;2
`;

// ---------------------------------------------------------------------------

fs.mkdirSync(OUT_DIR, { recursive: true });

const workbookPath = path.join(OUT_DIR, "MEx-Automate-Test.xlsx");
fs.writeFileSync(
  workbookPath,
  buildXlsx([
    { name: "Jan Orders", grid: janOrders },
    { name: "Feb Orders", grid: febOrders },
    { name: "Mar Orders", grid: marOrders },
    { name: "Sales Report", grid: salesReport },
  ])
);

const aprPath = path.join(OUT_DIR, "apr-orders.csv");
const mayPath = path.join(OUT_DIR, "may-orders.csv");
fs.writeFileSync(aprPath, aprCsv, "utf8");
// Semicolon-delimited on purpose, to exercise the delimiter sniffing.
fs.writeFileSync(mayPath, mayCsv, "utf8");

console.log(`Wrote ${workbookPath}`);
console.log(`Wrote ${aprPath}`);
console.log(`Wrote ${mayPath} (semicolon-delimited, to test delimiter sniffing)`);
