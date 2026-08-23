import { SourceTable } from "../src/taskpane/features/merge";
import {
  autoMapColumns,
  buildZoneGrid,
  emptyTemplate,
  UNMAPPED,
  validateTemplate,
} from "../src/taskpane/features/reportBuilder";
import { ColumnMapping, DataZone } from "../src/taskpane/shared/types";

const zoneHeaders = ["Order ID", "Customer", "Amount", "Margin"];

const source: SourceTable = {
  label: "new-data.csv",
  headers: ["order_id", "CUSTOMER", "Amount", "Notes"],
  rows: [
    [1, "Acme", 100, "rush"],
    [2, "Beta", 250, ""],
  ],
};

describe("autoMapColumns", () => {
  it("pairs columns whose headers normalize the same", () => {
    const mappings = autoMapColumns(zoneHeaders, source.headers);
    expect(mappings).toEqual([
      { targetHeader: "Order ID", sourceHeader: "order_id", manual: false },
      { targetHeader: "Customer", sourceHeader: "CUSTOMER", manual: false },
      { targetHeader: "Amount", sourceHeader: "Amount", manual: false },
      // "Margin" is a calculated column in the report - nothing feeds it.
      { targetHeader: "Margin", sourceHeader: UNMAPPED, manual: false },
    ]);
  });

  it("keeps a manual mapping the user has already made", () => {
    const existing: ColumnMapping[] = [
      { targetHeader: "Margin", sourceHeader: "Notes", manual: true },
    ];
    const mappings = autoMapColumns(zoneHeaders, source.headers, existing);
    expect(mappings[3]).toEqual({ targetHeader: "Margin", sourceHeader: "Notes", manual: true });
  });

  it("keeps a manual decision to leave a column unmapped", () => {
    const existing: ColumnMapping[] = [
      { targetHeader: "Amount", sourceHeader: UNMAPPED, manual: true },
    ];
    expect(autoMapColumns(zoneHeaders, source.headers, existing)[2].sourceHeader).toBe(UNMAPPED);
  });

  it("falls back to auto-matching when the manual source column has disappeared", () => {
    const existing: ColumnMapping[] = [
      { targetHeader: "Amount", sourceHeader: "Old Column", manual: true },
    ];
    const mappings = autoMapColumns(zoneHeaders, source.headers, existing);
    expect(mappings[2]).toEqual({ targetHeader: "Amount", sourceHeader: "Amount", manual: false });
  });

  it("does not reuse one source column for two zone columns", () => {
    const mappings = autoMapColumns(["Amount", "amount"], ["Amount"]);
    // Both normalize to the same key, so both legitimately point at the one source column.
    expect(mappings.map((mapping) => mapping.sourceHeader)).toEqual(["Amount", "Amount"]);
  });
});

describe("buildZoneGrid", () => {
  it("reshapes source rows into the zone's column order", () => {
    const mappings = autoMapColumns(zoneHeaders, source.headers);
    const { grid, mappedColumns } = buildZoneGrid(zoneHeaders, mappings, source);

    expect(mappedColumns).toEqual([0, 1, 2]);
    expect(grid).toEqual([
      [1, "Acme", 100, null],
      [2, "Beta", 250, null],
    ]);
  });

  it("reports both unmapped zone columns and unused source columns", () => {
    const mappings = autoMapColumns(zoneHeaders, source.headers);
    const { unmappedTargets, unusedSources } = buildZoneGrid(zoneHeaders, mappings, source);

    expect(unmappedTargets).toEqual(["Margin"]);
    expect(unusedSources).toEqual(["Notes"]);
  });

  it("honours a manual mapping when reshaping", () => {
    const mappings: ColumnMapping[] = [
      { targetHeader: "Order ID", sourceHeader: "order_id", manual: false },
      { targetHeader: "Customer", sourceHeader: "CUSTOMER", manual: false },
      { targetHeader: "Amount", sourceHeader: "Amount", manual: false },
      { targetHeader: "Margin", sourceHeader: "Notes", manual: true },
    ];
    const { grid, unusedSources } = buildZoneGrid(zoneHeaders, mappings, source);
    expect(grid[0]).toEqual([1, "Acme", 100, "rush"]);
    expect(unusedSources).toEqual([]);
  });

  it("produces an empty grid when the source has no rows", () => {
    const mappings = autoMapColumns(zoneHeaders, source.headers);
    const { grid } = buildZoneGrid(zoneHeaders, mappings, { ...source, rows: [] });
    expect(grid).toEqual([]);
  });

  it("leaves every column null when nothing maps", () => {
    const mappings = autoMapColumns(zoneHeaders, ["totally", "different"]);
    const { grid, mappedColumns } = buildZoneGrid(zoneHeaders, mappings, {
      label: "x",
      headers: ["totally", "different"],
      rows: [["a", "b"]],
    });
    expect(mappedColumns).toEqual([]);
    expect(grid).toEqual([[null, null, null, null]]);
  });
});

describe("validateTemplate", () => {
  const zone: DataZone = {
    name: "SalesData",
    kind: "namedRange",
    address: "Report!A5:D100",
    sheetName: "Report",
    headers: zoneHeaders,
    hasHeaderRow: true,
  };

  it("rejects a template with no zones", () => {
    expect(validateTemplate(emptyTemplate("Monthly"))).toEqual([
      "This template has no data zones yet.",
    ]);
  });

  it("rejects a zone with no columns", () => {
    const template = { ...emptyTemplate("Monthly"), zones: [{ ...zone, headers: [] }] };
    expect(validateTemplate(template)).toEqual(['Zone "SalesData" has no columns recorded.']);
  });

  it("accepts a complete template", () => {
    const template = { ...emptyTemplate("Monthly"), zones: [zone] };
    expect(validateTemplate(template)).toEqual([]);
  });
});

describe("emptyTemplate", () => {
  it("starts with defaults and timestamps", () => {
    const template = emptyTemplate("Monthly");
    expect(template.name).toBe("Monthly");
    expect(template.zones).toEqual([]);
    expect(template.options.clearExistingRows).toBe(true);
    expect(Date.parse(template.createdAt)).not.toBeNaN();
  });
});
