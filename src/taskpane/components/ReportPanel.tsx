import * as React from "react";
import {
  Button,
  Checkbox,
  Dropdown,
  Field,
  Input,
  makeStyles,
  Option,
  Radio,
  RadioGroup,
  Text,
  tokens,
} from "@fluentui/react-components";
import {
  ArrowSync20Regular,
  ArrowUpload20Regular,
  CheckmarkCircle16Filled,
  CursorClick20Regular,
  Delete20Regular,
  Dismiss16Regular,
  Save20Regular,
  TableSimple20Regular,
  Warning16Filled,
} from "@fluentui/react-icons";

import { listSheetNames, SourceTable } from "../features/merge";
import {
  autoMapColumns,
  createZoneFromSelection,
  emptyTemplate,
  listZoneCandidates,
  loadZone,
  readFileAsSources,
  readSheetAsSource,
  refreshReport,
  UNMAPPED,
  ZoneCandidate,
} from "../features/reportBuilder";
import { deleteTemplate, loadTemplates, saveTemplate } from "../shared/templateStore";
import { DataZone, OperationResult, ReportTemplate } from "../shared/types";
import { SUPPORTED_EXTENSIONS } from "../shared/workbookReader";
import {
  ActionButton,
  friendlyAddress,
  MoreOptions,
  ResultBanner,
  Step,
  Tip,
  useActionRunner,
  useSharedStyles,
} from "./ui";

type SourceMode = "sheet" | "file";

const KEEP_AS_IS = "(keep as is)";

const useStyles = makeStyles({
  savedRow: {
    display: "flex",
    alignItems: "center",
    columnGap: "8px",
    padding: "8px 0",
    borderBottom: `1px solid ${tokens.colorNeutralStroke3}`,
    ":last-child": {
      borderBottom: "none",
    },
  },
  savedText: {
    display: "flex",
    flexDirection: "column",
    flexGrow: 1,
    minWidth: 0,
  },
  savedName: {
    fontWeight: tokens.fontWeightSemibold,
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  zoneRow: {
    display: "flex",
    alignItems: "center",
    columnGap: "8px",
    padding: "8px 10px",
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground3,
  },
  zoneText: {
    display: "flex",
    flexDirection: "column",
    flexGrow: 1,
    minWidth: 0,
  },
  mapRow: {
    display: "flex",
    flexDirection: "column",
    rowGap: "4px",
  },
  mapLabel: {
    display: "flex",
    alignItems: "center",
    columnGap: "6px",
    fontWeight: tokens.fontWeightSemibold,
  },
  ok: {
    color: tokens.colorPaletteGreenForeground1,
    flexShrink: 0,
  },
  warn: {
    color: tokens.colorPaletteMarigoldForeground1,
    flexShrink: 0,
  },
  hiddenInput: {
    display: "none",
  },
});

/** Excel defined names must be unique and start with a letter. */
function nextZoneName(taken: Iterable<string>): string {
  const used = new Set(Array.from(taken, (name) => name.toLowerCase()));
  let index = 1;
  while (used.has(`reportarea${index}`)) {
    index += 1;
  }
  return `ReportArea${index}`;
}

/** Map every zone to the same source, refreshing each zone's column mapping. */
function withMappingsFor(template: ReportTemplate, table: SourceTable): ReportTemplate {
  const mappings = { ...template.mappings };
  for (const zone of template.zones) {
    mappings[zone.name] = autoMapColumns(zone.headers, table.headers, mappings[zone.name] ?? []);
  }
  return { ...template, mappings };
}

const ReportPanel: React.FC = () => {
  const shared = useSharedStyles();
  const styles = useStyles();
  const runner = useActionRunner();

  const [template, setTemplate] = React.useState<ReportTemplate>(() => emptyTemplate(""));
  const [savedTemplates, setSavedTemplates] = React.useState<ReportTemplate[]>([]);
  const [candidates, setCandidates] = React.useState<ZoneCandidate[]>([]);
  const [sheetNames, setSheetNames] = React.useState<string[]>([]);

  const [hasHeaderRow, setHasHeaderRow] = React.useState(true);
  const [activeZoneName, setActiveZoneName] = React.useState("");
  const [zoneResult, setZoneResult] = React.useState<OperationResult | null>(null);

  const [sourceMode, setSourceMode] = React.useState<SourceMode>("sheet");
  const [sourceSheet, setSourceSheet] = React.useState("");
  const [source, setSource] = React.useState<SourceTable | null>(null);
  const [sourceError, setSourceError] = React.useState("");

  const fileInputRef = React.useRef<HTMLInputElement | null>(null);

  React.useEffect(() => {
    void loadTemplates().then(setSavedTemplates);
    void listZoneCandidates().then(setCandidates);
    void listSheetNames().then(setSheetNames);
  }, [runner.result]);

  const activeZone = template.zones.find((zone) => zone.name === activeZoneName) ?? null;
  const mappings = activeZone ? (template.mappings[activeZone.name] ?? []) : [];
  const sourceHeaders = source?.headers ?? [];
  const matchedCount = mappings.filter((mapping) => mapping.sourceHeader !== UNMAPPED).length;

  const addZone = (zone: DataZone) => {
    setTemplate((current) => ({
      ...current,
      name: current.name || zone.sheetName,
      zones: [...current.zones.filter((item) => item.name !== zone.name), zone],
      mappings: {
        ...current.mappings,
        [zone.name]: autoMapColumns(zone.headers, source?.headers ?? []),
      },
    }));
    setActiveZoneName(zone.name);
  };

  const removeZone = (name: string) => {
    setTemplate((current) => {
      const rest = { ...current.mappings };
      delete rest[name];
      return { ...current, zones: current.zones.filter((zone) => zone.name !== name), mappings: rest };
    });
    if (activeZoneName === name) {
      setActiveZoneName("");
    }
  };

  const takeSelection = async () => {
    setZoneResult(null);
    const name = nextZoneName([
      ...candidates.map((item) => item.name),
      ...template.zones.map((zone) => zone.name),
    ]);
    const { result, zone } = await createZoneFromSelection(name, hasHeaderRow);
    if (zone) {
      addZone(zone);
      void listZoneCandidates().then(setCandidates);
    } else {
      setZoneResult(result);
    }
  };

  const applySource = (table: SourceTable | null, problem: string) => {
    setSource(table);
    setSourceError(table ? "" : problem);
    if (table) {
      setTemplate((current) => withMappingsFor(current, table));
    }
  };

  const pickSourceSheet = async (name: string) => {
    setSourceSheet(name);
    applySource(await readSheetAsSource(name), `"${name}" needs a heading row and at least one row of data.`);
  };

  const pickSourceFile = async (file: File | undefined) => {
    if (!file) {
      return;
    }
    try {
      const tables = await readFileAsSources(file);
      applySource(tables[0] ?? null, `No data found in "${file.name}".`);
    } catch (error) {
      applySource(null, error instanceof Error ? error.message : String(error));
    }
  };

  const setMapping = (targetHeader: string, sourceHeader: string) => {
    if (!activeZone) {
      return;
    }
    setTemplate((current) => ({
      ...current,
      mappings: {
        ...current.mappings,
        [activeZone.name]: (current.mappings[activeZone.name] ?? []).map((mapping) =>
          mapping.targetHeader === targetHeader ? { ...mapping, sourceHeader, manual: true } : mapping
        ),
      },
    }));
  };

  const setOption = (key: keyof ReportTemplate["options"], value: boolean) =>
    setTemplate((current) => ({ ...current, options: { ...current.options, [key]: value } }));

  const openSaved = (saved: ReportTemplate) => {
    setTemplate(saved);
    setActiveZoneName(saved.zones[0]?.name ?? "");
    setSource(null);
    setSourceError("");
    setSourceMode("sheet");
    if (saved.sourceSheetName) {
      void pickSourceSheet(saved.sourceSheetName);
    } else {
      setSourceSheet("");
    }
  };

  /** Re-run a saved report against its remembered sheet, without touching the steps. */
  const updateSaved = async (saved: ReportTemplate): Promise<OperationResult> => {
    if (!saved.sourceSheetName) {
      return { ok: false, message: "This report doesn't remember where its data comes from. Open it and choose a sheet." };
    }
    const table = await readSheetAsSource(saved.sourceSheetName);
    if (!table) {
      return {
        ok: false,
        message: `The sheet "${saved.sourceSheetName}" is missing or empty.`,
      };
    }
    const ready = withMappingsFor(saved, table);
    const sources: Record<string, SourceTable> = {};
    for (const zone of ready.zones) {
      sources[zone.name] = table;
    }
    return refreshReport(ready, sources);
  };

  const templateToSave: ReportTemplate = {
    ...template,
    sourceSheetName: sourceMode === "sheet" && sourceSheet ? sourceSheet : template.sourceSheetName,
  };

  return (
    <div>
      {savedTemplates.length > 0 ? (
        <section className={shared.card}>
          <Text weight="semibold">Your saved reports</Text>
          <div>
            {savedTemplates.map((saved) => (
              <div key={saved.name} className={styles.savedRow}>
                <div className={styles.savedText}>
                  <span className={styles.savedName} title={saved.name}>
                    {saved.name}
                  </span>
                  <Text className={shared.hint}>
                    {saved.sourceSheetName ? `Data from "${saved.sourceSheetName}"` : "No data sheet saved"}
                  </Text>
                </div>
                <Button size="small" appearance="subtle" onClick={() => openSaved(saved)}>
                  Edit
                </Button>
                <Button
                  size="small"
                  appearance="subtle"
                  icon={<Delete20Regular />}
                  aria-label={`Delete ${saved.name}`}
                  title="Delete"
                  disabled={runner.busy}
                  onClick={() => void runner.run(`delete-${saved.name}`, () => deleteTemplate(saved.name))}
                />
              </div>
            ))}
          </div>
          {savedTemplates.some((saved) => saved.sourceSheetName) ? (
            <div className={shared.column}>
              {savedTemplates
                .filter((saved) => saved.sourceSheetName)
                .map((saved) => (
                  <ActionButton
                    key={saved.name}
                    runner={runner}
                    id={`update-${saved.name}`}
                    icon={<ArrowSync20Regular />}
                    label={`Update "${saved.name}" now`}
                    busyLabel="Updating..."
                    onRun={() => updateSaved(saved)}
                  />
                ))}
            </div>
          ) : null}
          {runner.result && runner.resultId?.startsWith("delete-") ? (
            <ResultBanner result={runner.result} onDismiss={runner.clear} />
          ) : null}
        </section>
      ) : (
        <div style={{ marginBottom: "10px" }}>
          <Tip>
            Use this when you have a report laid out the way you like, and each week or month you
            paste new data into it. The layout, totals and formulas are kept; only the data rows
            change.
          </Tip>
        </div>
      )}

      <Step
        number={1}
        title="Which part of the report gets new data?"
        description="Select the heading row plus the data rows in your report (for example A3:E6), then click the button."
      >
        <Button icon={<CursorClick20Regular />} onClick={() => void takeSelection()} disabled={runner.busy}>
          Use the cells I&apos;ve selected
        </Button>

        {candidates.length > 0 ? (
          <Dropdown
            placeholder="...or pick an existing table"
            value=""
            selectedOptions={[]}
            onOptionSelect={(_event, data) => {
              const candidate = candidates.find((item) => `${item.kind}:${item.name}` === data.optionValue);
              if (candidate) {
                void loadZone(candidate.name, candidate.kind, hasHeaderRow).then((zone) => {
                  if (zone) {
                    addZone(zone);
                  }
                });
              }
            }}
          >
            {candidates.map((candidate) => (
              <Option
                key={`${candidate.kind}:${candidate.name}`}
                value={`${candidate.kind}:${candidate.name}`}
                text={candidate.name}
              >
                {`${candidate.name} (${friendlyAddress(candidate.address)})`}
              </Option>
            ))}
          </Dropdown>
        ) : null}

        <MoreOptions>
          <Checkbox
            label="The first selected row is headings"
            checked={hasHeaderRow}
            onChange={(_event, data) => setHasHeaderRow(Boolean(data.checked))}
          />
        </MoreOptions>

        {zoneResult ? <ResultBanner result={zoneResult} onDismiss={() => setZoneResult(null)} /> : null}

        {template.zones.length > 0 ? (
          <RadioGroup value={activeZoneName} onChange={(_event, data) => setActiveZoneName(data.value)}>
            {template.zones.map((zone) => (
              <div key={zone.name} className={styles.zoneRow}>
                {template.zones.length > 1 ? <Radio value={zone.name} label="" /> : <TableSimple20Regular />}
                <div className={styles.zoneText}>
                  <Text weight="semibold">{friendlyAddress(zone.address)}</Text>
                  <Text className={shared.hint}>
                    {zone.headers.length} columns: {zone.headers.slice(0, 3).join(", ")}
                    {zone.headers.length > 3 ? ", ..." : ""}
                  </Text>
                </div>
                <Button
                  size="small"
                  appearance="subtle"
                  icon={<Dismiss16Regular />}
                  aria-label="Remove"
                  onClick={() => removeZone(zone.name)}
                />
              </div>
            ))}
          </RadioGroup>
        ) : null}
      </Step>

      <Step
        number={2}
        title="Where is the new data?"
        waitingFor={activeZone ? undefined : "Choose the part of the report first."}
      >
        <RadioGroup
          layout="horizontal"
          value={sourceMode}
          onChange={(_event, data) => {
            setSourceMode(data.value as SourceMode);
            setSource(null);
            setSourceError("");
          }}
        >
          <Radio value="sheet" label="A sheet" />
          <Radio value="file" label="A file" />
        </RadioGroup>

        {sourceMode === "sheet" ? (
          <Dropdown
            placeholder="Choose a sheet"
            value={sourceSheet}
            selectedOptions={sourceSheet ? [sourceSheet] : []}
            onOptionSelect={(_event, data) => void pickSourceSheet(String(data.optionValue))}
          >
            {sheetNames
              .filter((name) => name !== activeZone?.sheetName)
              .map((name) => (
                <Option key={name} value={name}>
                  {name}
                </Option>
              ))}
          </Dropdown>
        ) : (
          <>
            <input
              ref={fileInputRef}
              className={styles.hiddenInput}
              type="file"
              accept={SUPPORTED_EXTENSIONS.join(",")}
              onChange={(event) => {
                void pickSourceFile(event.target.files?.[0]);
                event.target.value = "";
              }}
            />
            <Button icon={<ArrowUpload20Regular />} onClick={() => fileInputRef.current?.click()}>
              Choose a file
            </Button>
          </>
        )}

        {sourceError ? (
          <Text className={shared.hint} style={{ color: tokens.colorPaletteRedForeground1 }}>
            {sourceError}
          </Text>
        ) : null}
        {source ? (
          <Text className={shared.hint}>
            Found {source.rows.length} rows and {source.headers.length} columns in {source.label}.
          </Text>
        ) : null}
      </Step>

      <Step
        number={3}
        title="Check the columns line up"
        waitingFor={activeZone && source ? undefined : "Choose the new data first."}
      >
        <Text className={shared.hint}>
          {matchedCount === mappings.length
            ? `All ${mappings.length} columns were matched by name.`
            : `${matchedCount} of ${mappings.length} columns matched by name. Pick where the others come from, or keep them as they are (for example, formula columns).`}
        </Text>
        {mappings.map((mapping) => {
          const matched = mapping.sourceHeader !== UNMAPPED;
          return (
            <div key={mapping.targetHeader} className={styles.mapRow}>
              <span className={styles.mapLabel}>
                {matched ? (
                  <CheckmarkCircle16Filled className={styles.ok} />
                ) : (
                  <Warning16Filled className={styles.warn} />
                )}
                {mapping.targetHeader}
              </span>
              <Dropdown
                size="small"
                value={matched ? `Gets data from "${mapping.sourceHeader}"` : KEEP_AS_IS}
                selectedOptions={[mapping.sourceHeader]}
                onOptionSelect={(_event, data) => setMapping(mapping.targetHeader, String(data.optionValue))}
              >
                <Option value={UNMAPPED} text={KEEP_AS_IS}>
                  {KEEP_AS_IS}
                </Option>
                {sourceHeaders.map((header) => (
                  <Option key={header} value={header} text={`Gets data from "${header}"`}>
                    {header}
                  </Option>
                ))}
              </Dropdown>
            </div>
          );
        })}
      </Step>

      <Step
        number={4}
        title="Update the report"
        waitingFor={activeZone && source ? undefined : "Finish the steps above first."}
      >
        <MoreOptions>
          <Checkbox
            label="Remove the old rows first"
            checked={template.options.clearExistingRows}
            onChange={(_event, data) => setOption("clearExistingRows", Boolean(data.checked))}
          />
          <Checkbox
            label="Keep the report's colours and number formats"
            checked={template.options.preserveFormatting}
            onChange={(_event, data) => setOption("preserveFormatting", Boolean(data.checked))}
          />
          <Checkbox
            label="Refresh PivotTables afterwards"
            checked={template.options.refreshPivotTables}
            onChange={(_event, data) => setOption("refreshPivotTables", Boolean(data.checked))}
          />
          <Checkbox
            label="Recalculate formulas afterwards"
            checked={template.options.recalculate}
            onChange={(_event, data) => setOption("recalculate", Boolean(data.checked))}
          />
        </MoreOptions>

        <ActionButton
          runner={runner}
          id="refresh"
          wide
          icon={<ArrowSync20Regular />}
          label="Update report"
          busyLabel="Updating..."
          disabled={!activeZone || !source}
          onRun={() => refreshReport(template, activeZone && source ? { [activeZone.name]: source } : {})}
        />

        <Field
          label="Save for next time"
          hint={
            sourceMode === "sheet"
              ? "Saved inside this workbook. Next time, one click updates it from the same sheet."
              : "Saved inside this workbook. Next time, open it and choose the new file."
          }
        >
          <div className={shared.row}>
            <Input
              style={{ flexGrow: 1 }}
              placeholder="Report name, e.g. Monthly sales"
              value={template.name}
              onChange={(_event, data) => setTemplate((current) => ({ ...current, name: data.value }))}
            />
          </div>
        </Field>
        <ActionButton
          runner={runner}
          id="save"
          appearance="secondary"
          icon={<Save20Regular />}
          label="Save"
          disabled={template.name.trim() === ""}
          onRun={() => saveTemplate(templateToSave)}
        />
      </Step>
    </div>
  );
};

export default ReportPanel;
