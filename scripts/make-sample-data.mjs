/**
 * Generates sample data for manually testing the add-in in Excel.
 *
 * Writes samples/MEx-Automate-Test.xlsx (six sheets, deliberately messy) plus
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

// Sheet 5: the export that ruins an afternoon. Six months of orders from three
// regions, with every problem "Fix messy data" handles: title and "Generated by"
// rows, "Region: ..." group headings, subtotal rows, a heading row repeated
// mid-table, a grand total, a footnote, five date styles, numbers typed as text
// ($, thousands separators, European decimals, brackets, SAP's trailing minus,
// a letter O for a zero), N/A placeholders, stray spaces, one-letter typos,
// inconsistent capitals and duplicated rows. Deterministic, so the fixes and the
// dashboard are the same every time the workbook is regenerated.
export function buildSalesExtract() {
  let seed = 7;
  const rand = () => {
    seed = (seed * 16807) % 2147483647;
    return (seed - 1) / 2147483646;
  };
  const pick = (items) => items[Math.floor(rand() * items.length)];

  const products = [
    ["Jollof Rice Mix 1kg", "Staples", 12.5],
    ["Palm Oil 1L", "Staples", 9.75],
    ["Plantain Chips 150g", "Snacks", 3.2],
    ["Chin Chin 250g", "Snacks", 4.1],
    ["Suya Spice 100g", "Spices", 5.6],
    ["Pepper Soup Spice", "Spices", 4.85],
    ["Zobo Drink 500ml", "Beverages", 2.9],
    ["Kunu Drink 500ml", "Beverages", 2.75],
    ["Garri Ijebu 2kg", "Staples", 7.4],
    ["Egusi Ground 500g", "Staples", 8.9],
    ["Tiger Nut Milk 1L", "Beverages", 6.3],
    ["Puff Puff Mix 500g", "Snacks", 3.95],
  ];
  const customers = {
    Europe: [
      ["Acme Ltd", "ACME LTD", " acme ltd "],
      ["Beta Trading", "beta trading"],
      ["Epsilon PLC"],
    ],
    Africa: [["Gamma Co", "Gamma  Co"], ["Delta Supplies", "Delta Suplies"], ["Zeta Holdings"]],
    Americas: [["Harbor Foods"], ["Northwind Grocers", "NORTHWIND GROCERS"], ["Sol Markets"]],
  };
  const channels = [["Online", "online"], ["Retail"], ["Wholesale", "Wholsale"]];
  const pad = (n) => String(n).padStart(2, "0");
  const MONTH = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"];
  const serialOf = (y, m, d) => Math.round((Date.UTC(y, m, d) - Date.UTC(1899, 11, 30)) / 86400000);

  const dateText = (i, m, d) => {
    switch (i % 6) {
      case 0:
        return `2024-${pad(m + 1)}-${pad(d)}`;
      case 1:
        return `${pad(d)}/${pad(m + 1)}/2024`;
      case 2:
        return `${MONTH[m]} ${d}, 2024`;
      case 3:
        return `${d}-${MONTH[m]}-2024`;
      case 4:
        return serialOf(2024, m, d);
      default:
        return `${pad(d)}.${pad(m + 1)}.2024`;
    }
  };
  const money = (i, value) => {
    const fixed = value.toFixed(2);
    const us = Number(fixed).toLocaleString("en-US", { minimumFractionDigits: 2 });
    switch (i % 9) {
      case 0:
        return Number(fixed);
      case 1:
        return `${us}`;
      case 2:
        return us;
      case 3:
        return `${fixed.replace(".", ",").replace(/\B(?=(\d{3})+(?!\d))/g, ".")} `;
      case 4:
        return ` $ ${us}`;
      case 5:
        return Number(fixed);
      case 6:
        return `USD ${us}`;
      case 7:
        return us.replace(/(\d)0/, "$1O"); // a letter O typed for a zero
      default:
        return `${us}`;
    }
  };

  const header = [
    "Order ID",
    "Order Date",
    "Customer",
    "Product",
    "Category",
    "Channel",
    "Units",
    "Unit Price",
    "Revenue",
  ];
  const grid = [
    ["ACME FOODS LTD - SALES EXTRACT, JAN-JUN 2024"],
    ["Generated 01/07/2024 by SAP ERP (user: finance01)"],
    [],
    header,
  ];

  let id = 5000;
  let grandUnits = 0;
  let grandRevenue = 0;
  ["Europe", "Africa", "Americas"].forEach((region, r) => {
    grid.push([`Region: ${region}`]);
    let units = 0;
    let revenue = 0;
    for (let i = 0; i < 26; i += 1) {
      const n = r * 26 + i;
      const month = Math.floor((i / 26) * 6);
      const day = 1 + Math.floor(rand() * 27);
      const [product, category, price] = pick(products);
      const qty = 1 + Math.floor(rand() * 40);
      // A few returns, written the two ways accounting systems write negatives.
      const isReturn = n % 23 === 11;
      const amount = (isReturn ? -1 : 1) * qty * price;
      const customer = pick(pick(customers[region]));
      const channel = pick(pick(channels));
      let revenueCell = money(n, Math.abs(amount));
      if (isReturn)
        revenueCell =
          n % 2 ? `(${Math.abs(amount).toFixed(2)})` : `${Math.abs(amount).toFixed(2)}-`;
      const unitsCell = n % 17 === 5 ? "N/A" : n % 3 === 0 ? String(qty) : qty;
      id += 1;
      const row = [
        `SO-${id}`,
        dateText(n, month, day),
        customer,
        product,
        category,
        channel,
        unitsCell,
        price,
        revenueCell,
      ];
      grid.push(row);
      if (n % 19 === 7) grid.push([...row]);
      if (typeof unitsCell === "number" || /^\d+$/.test(String(unitsCell))) units += qty;
      revenue += amount;
      // Page breaks in the original print layout repeat the headings.
      if (r === 1 && i === 12) grid.push([...header]);
    }
    grid.push([
      `${region} Total`,
      null,
      null,
      null,
      null,
      null,
      units,
      null,
      Math.round(revenue * 100) / 100,
    ]);
    grid.push([]);
    grandUnits += units;
    grandRevenue += revenue;
  });
  grid.push([
    "Grand Total",
    null,
    null,
    null,
    null,
    null,
    grandUnits,
    null,
    Math.round(grandRevenue * 100) / 100,
  ]);
  grid.push([]);
  grid.push(["Source: SAP ERP - internal use only. Figures are unaudited and exclude VAT."]);
  return grid;
}
const salesExtract = buildSalesExtract();

// Sheet 6: a sheet that looks fine but isn't, for the Review tool. Every row
// below is plausible; six of them are the kind of mistake that quietly makes a
// report wrong. Revenue is calculated everywhere except row 9, which is typed
// in; the total stops at row 14; SO-1003 is used twice; "acme ltd" and
// "  Gamma Co " won't group with their proper spellings; row 13's Units is
// text; row 12's Units has an extra zero; row 15 is dated five years out.
const salesCheck = (() => {
  const serial = (y, m, d) => Math.round((Date.UTC(y, m, d) - Date.UTC(1899, 11, 30)) / 86400000);
  const orders = [
    ["SO-1001", serial(2024, 0, 3), "Acme Ltd", 12, 12.5],
    ["SO-1002", serial(2024, 0, 9), "Beta Trading", 8, 12.5],
    ["SO-1003", serial(2024, 1, 2), "acme ltd", 20, 12.5],
    ["SO-1004", serial(2024, 1, 14), "Gamma Co", 35, 12.5],
    ["SO-1005", serial(2024, 2, 1), "Beta Trading", 7, 12.5],
    ["SO-1003", serial(2024, 2, 8), "Delta Supplies", 5, 12.5],
    ["SO-1007", serial(2024, 2, 19), "Gamma Co", 9, 12.5],
    ["SO-1008", serial(2024, 3, 2), "Acme Ltd", 15, 12.5],
    ["SO-1009", serial(2024, 3, 11), "Delta Supplies", 11, 12.5],
    ["SO-1010", serial(2024, 3, 22), "Acme Ltd", 6, 12.5],
    ["SO-1011", serial(2024, 4, 6), "Beta Trading", 9500, 12.5],
    ["SO-1012", serial(2024, 4, 17), "Gamma Co", "1,250.00", 12.5],
    ["SO-1013", serial(2024, 4, 28), "Delta Supplies", 10, 12.5],
    ["SO-1014", serial(2029, 5, 4), "Acme Ltd", 13, 12.5],
    ["SO-1015", serial(2024, 5, 12), "  Gamma Co ", 9, 12.5],
  ];
  const grid = [["Order ID", "Order Date", "Customer", "Units", "Unit Price", "Revenue"]];
  orders.forEach((order, i) => {
    const row = i + 2;
    // Row 9 is typed in rather than calculated, like a value someone pasted.
    grid.push([...order, row === 9 ? 187.5 : { f: `D${row}*E${row}` }]);
  });
  grid.push([]);
  grid.push(["Total", null, null, null, null, { f: "SUM(F2:F14)" }]);
  return grid;
})();

// Sheet 7: a crosstab, for /unpivot. The year sits on a merged row above the
// months - which reads back as blanks, exactly as a real one does - and the
// figures are read across, which is comfortable for a human and useless for a
// chart. One gap, because real budgets have them.
const budgetByMonth = (() => {
  const grid = [
    [null, null, "2024", null, null, null, null, null],
    ["Region", "Product", "Jan", "Feb", "Mar", "Apr", "May", "Jun"],
  ];
  const rows = [
    ["North", "Widget", 12400, 13100, 12850, 14200, 15050, 15600],
    ["North", "Gadget", 8200, 8650, 9100, 8950, 9400, 9950],
    ["North", "Sprocket", 3100, 3250, null, 3600, 3720, 3810],
    ["South", "Widget", 9800, 10200, 10450, 11100, 11550, 12000],
    ["South", "Gadget", 6400, 6550, 6900, 7050, 7300, 7600],
    ["South", "Sprocket", 2250, 2300, 2410, 2500, 2580, 2650],
    ["East", "Widget", 7100, 7400, 7650, 8000, 8250, 8500],
    ["East", "Gadget", 4900, 5050, 5200, 5400, 5600, 5800],
    ["West", "Widget", 5600, 5800, 6050, 6300, 6500, 6750],
    ["West", "Gadget", 3800, 3950, 4100, 4250, 4400, 4550],
  ];
  return [...grid, ...rows];
})();

// Sheets 8 and 9: the same order book a month apart, for /compare. Between the
// two: one order's units and amount are corrected, one customer is renamed, one
// order is cancelled, one is new, and one amount moves by half a penny - which
// is rounding, not a change, and should not be reported.
const [aprilBook, mayBook] = (() => {
  const serial = (y, m, d) => Math.round((Date.UTC(y, m, d) - Date.UTC(1899, 11, 30)) / 86400000);
  const header = ["Order ID", "Order Date", "Customer", "Region", "Units", "Amount"];
  const april = [
    ["SO-2001", serial(2024, 3, 2), "Acme Ltd", "North", 10, 1250],
    ["SO-2002", serial(2024, 3, 5), "Beta Trading", "South", 4, 520],
    ["SO-2003", serial(2024, 3, 9), "Gamma Co", "East", 7, 910.25],
    ["SO-2004", serial(2024, 3, 14), "Delta Supplies", "West", 12, 1560],
    ["SO-2005", serial(2024, 3, 18), "Epsilon PLC", "North", 3, 375],
    ["SO-2006", serial(2024, 3, 23), "Zeta Holdings", "South", 9, 1170],
    ["SO-2007", serial(2024, 3, 27), "Acme Ltd", "East", 6, 780],
  ];
  const may = april
    // SO-2004 was cancelled.
    .filter((row) => row[0] !== "SO-2004")
    .map((row) => {
      const copy = [...row];
      if (copy[0] === "SO-2002") {
        copy[4] = 6;
        copy[5] = 780;
      }
      if (copy[0] === "SO-2005") copy[2] = "Epsilon Ltd";
      // Half a penny: rounding noise, and MEx should stay quiet about it.
      if (copy[0] === "SO-2003") copy[5] = 910.254;
      return copy;
    });
  may.push(["SO-2008", serial(2024, 4, 4), "Theta Systems", "West", 5, 650]);
  return [
    [header, ...april],
    [header, ...may],
  ];
})();

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
    { name: "Sales Extract", grid: salesExtract },
    { name: "Sales Check", grid: salesCheck },
    { name: "Budget by Month", grid: budgetByMonth },
    { name: "April Book", grid: aprilBook },
    { name: "May Book", grid: mayBook },
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
