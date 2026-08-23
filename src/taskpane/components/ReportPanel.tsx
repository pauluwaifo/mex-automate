import * as React from "react";
import {
  Button,
  Checkbox,
  Dropdown,
  Field,
  Input,
  Option,
  Radio,
  RadioGroup,
  Text,
} from "@fluentui/react-components";

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
import {
  deleteTemplate,
  loadTemplates,
  saveTemplate,
} from "../shared/templateStore";
import { ColumnMapping, DataZone, ReportTemplate } from "../shared/types";
import { SUPPORTED_EXTENSIONS } from "../shared/workbookReader";
import { ResultBanner, RunButton, Section, useActionRunner, useSharedStyles } from "./ui";

type SourceMode = "sheet" | "file";

const ReportPanel: React.FC = () => {
  const styles = useSharedStyles();
  const runner = useActionRunner();

  const [template, setTemplate] = React.useState<ReportTemplate>(() => emptyTemplate("My report"));
  const [savedTemplates, setSavedTemplates] = React.useState<ReportTemplate[]>([]);
  const [candidates, setCandidates] = React.useState<ZoneCandidate[]>([]);
  const [sheetNames, setSheetNames] = React.useState<string[]>([]);

  const [newZoneName, setNewZoneName] = React.useState("");
  const [newZoneHasHeader, setNewZoneHasHeader] = React.useState(true);
  const [activeZoneName, setActiveZoneName] = React.useState<string>("");

  const [sourceMode, setSourceMode] = React.useState<SourceMode>("sheet");
  const [sourceSheet, setSourceSheet] = React.useState("");
  const [source, setSource] = React.useState<SourceTable | null>(null);
  const [sourceError, setSourceError] = React.useState("");

  const reload = React.useCallback(() => {
    void loadTemplates().then(setSavedTemplates);
    void listZoneCandidates().then(setCandidates);
    void listSheetNames().then(setSheetNames);
  }, []);

  React.useEffect(reload, [reload, runner.result]);

  const activeZone = template.zones.find((zone) => zone.name === activeZoneName) ?? null;
  const mappings = activeZone ? (template.mappings[activeZone.name] ?? []) : [];

  /** Keep mappings in step whenever the zone or the incoming source changes. */
  const applyAutoMapping = React.useCallback(
    (zone: DataZone, table: SourceTable | null) => {
      setTemplate((current) => ({
        ...current,
        mappings: {
          ...current.mappings,
          [zone.name]: autoMapColumns(
            zone.headers,
            table?.headers ?? [],
            current.mappings[zone.name] ?? []
          ),
        },
      }));
    },
    []
  );

  const addZone = (zone: DataZone) => {
    setTemplate((current) => ({
      ...current,
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
      const mappingsWithout = { ...current.mappings };
      delete mappingsWithout[name];
      return {
        ...current,
        zones: current.zones.filter((zone) => zone.name !== name),
        mappings: mappingsWithout,
      };
    });
    if (activeZoneName === name) {
      setActiveZoneName("");
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
          mapping.targetHeader === targetHeader
            ? { ...mapping, sourceHeader, manual: true }
            : mapping
        ),
      },
    }));
  };

  const pickSourceSheet = async (name: string) => {
    setSourceSheet(name);
    setSourceError("");
    const table = await readSheetAsSource(name);
    setSource(table);
    if (!table) {
      setSourceError(`"${name}" needs a header row and at least one data row.`);
    } else if (activeZone) {
      applyAutoMapping(activeZone, table);
    }
  };

  const pickSourceFile = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }
    setSourceError("");
    try {
      const tables = await readFileAsSources(file);
      const table = tables[0] ?? null;
      setSource(table);
      if (!table) {
        setSourceError(`No data rows found in "${file.name}".`);
      } else if (activeZone) {
        applyAutoMapping(activeZone, table);
      }
    } catch (error) {
      setSource(null);
      setSourceError(error instanceof Error ? error.message : String(error));
    }
  };

  const setOption = (key: keyof ReportTemplate["options"], value: boolean) =>
    setTemplate((current) => ({ ...current, options: { ...current.options, [key]: value } }));

  const sourceHeaders = source?.headers ?? [];

  return (
    <div>
      <ResultBanner result={runner.result} />

      <Section title="Template" description="Saved inside this workbook, so it travels with the file.">
        <Field label="Name">
          <Input
            value={template.name}
            onChange={(_event, data) => setTemplate((current) => ({ ...current, name: data.value }))}
          />
        </Field>
        <div className={styles.row}>
          <RunButton
            label="Save"
            appearance="secondary"
            busy={runner.busy}
            onClick={() => void runner.run(() => saveTemplate(template))}
          />
          <RunButton
            label="Delete"
            appearance="secondary"
            busy={runner.busy}
            disabled={!savedTemplates.some((item) => item.name === template.name)}
            onClick={() => void runner.run(() => deleteTemplate(template.name))}
          />
        </div>
        {savedTemplates.length > 0 ? (
          <Field label="Load a saved template">
            <Dropdown
              placeholder="Select a template"
              onOptionSelect={(_event, data) => {
                const found = savedTemplates.find((item) => item.name === data.optionValue);
                if (found) {
                  setTemplate(found);
                  setActiveZoneName(found.zones[0]?.name ?? "");
                }
              }}
            >
              {savedTemplates.map((item) => (
                <Option key={item.name} value={item.name}>
                  {item.name}
                </Option>
              ))}
            </Dropdown>
          </Field>
        ) : null}
      </Section>

      <Section
        title="Data zones"
        description="The parts of the report that get replaced on refresh. Everything else - titles, totals, charts - stays put."
      >
        {template.zones.length === 0 ? (
          <Text className={styles.hint}>No zones yet.</Text>
        ) : (
          <div className={styles.scrollList}>
            {template.zones.map((zone) => (
              <div key={zone.name} className={styles.row}>
                <Radio
                  checked={activeZoneName === zone.name}
                  onChange={() => setActiveZoneName(zone.name)}
                  label={`${zone.name} (${zone.headers.length} cols, ${zone.sheetName})`}
                />
                <Button size="small" appearance="subtle" onClick={() => removeZone(zone.name)}>
                  Remove
                </Button>
              </div>
            ))}
          </div>
        )}

        <Field
          label="Add a zone from the current selection"
          hint="Select the header row plus the data rows, then name the zone."
        >
          <Input
            placeholder="SalesData"
            value={newZoneName}
            onChange={(_event, data) => setNewZoneName(data.value)}
          />
        </Field>
        <Checkbox
          label="First row of the selection is a header row"
          checked={newZoneHasHeader}
          onChange={(_event, data) => setNewZoneHasHeader(Boolean(data.checked))}
        />
        <RunButton
          label="Add zone from selection"
          appearance="secondary"
          busy={runner.busy}
          disabled={newZoneName.trim() === ""}
          onClick={() =>
            void runner.run(async () => {
              const { result, zone } = await createZoneFromSelection(newZoneName, newZoneHasHeader);
              if (zone) {
                addZone(zone);
                setNewZoneName("");
              }
              return result;
            })
          }
        />

        {candidates.length > 0 ? (
          <Field label="...or use an existing table or named range">
            <Dropdown
              placeholder="Select"
              onOptionSelect={(_event, data) => {
                const candidate = candidates.find(
                  (item) => `${item.kind}:${item.name}` === data.optionValue
                );
                if (!candidate) {
                  return;
                }
                void loadZone(candidate.name, candidate.kind, true).then((zone) => {
                  if (zone) {
                    addZone(zone);
                  }
                });
              }}
            >
              {candidates.map((candidate) => (
                <Option
                  key={`${candidate.kind}:${candidate.name}`}
                  value={`${candidate.kind}:${candidate.name}`}
                  text={candidate.name}
                >
                  {`${candidate.name} - ${candidate.kind === "table" ? "table" : "named range"}`}
                </Option>
              ))}
            </Dropdown>
          </Field>
        ) : null}
      </Section>

      <Section title="New data" description="Where the fresh rows come from.">
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
            placeholder="Select a sheet"
            value={sourceSheet}
            selectedOptions={sourceSheet ? [sourceSheet] : []}
            onOptionSelect={(_event, data) => void pickSourceSheet(String(data.optionValue))}
          >
            {sheetNames.map((name) => (
              <Option key={name} value={name}>
                {name}
              </Option>
            ))}
          </Dropdown>
        ) : (
          <input type="file" accept={SUPPORTED_EXTENSIONS.join(",")} onChange={pickSourceFile} />
        )}

        {sourceError ? <Text className={styles.hint}>{sourceError}</Text> : null}
        {source ? (
          <Text className={styles.hint}>
            {source.label}: {source.rows.length} row(s), {source.headers.length} column(s)
          </Text>
        ) : null}
      </Section>

      {activeZone ? (
        <Section
          title={`Column mapping - ${activeZone.name}`}
          description="Columns are matched by name. Change any that guessed wrong; leave a column unmapped to keep whatever is already there (calculated columns are always left alone)."
        >
          {mappings.length === 0 ? (
            <Text className={styles.hint}>Pick a data source to map columns.</Text>
          ) : (
            mappings.map((mapping: ColumnMapping) => (
              <Field key={mapping.targetHeader} label={mapping.targetHeader}>
                <Dropdown
                  value={mapping.sourceHeader === UNMAPPED ? "(leave unchanged)" : mapping.sourceHeader}
                  selectedOptions={[mapping.sourceHeader]}
                  onOptionSelect={(_event, data) =>
                    setMapping(mapping.targetHeader, String(data.optionValue))
                  }
                >
                  <Option value={UNMAPPED} text="(leave unchanged)">
                    (leave unchanged)
                  </Option>
                  {sourceHeaders.map((header) => (
                    <Option key={header} value={header}>
                      {header}
                    </Option>
                  ))}
                </Dropdown>
              </Field>
            ))
          )}
        </Section>
      ) : null}

      <Section title="Refresh">
        <Checkbox
          label="Clear the old rows first"
          checked={template.options.clearExistingRows}
          onChange={(_event, data) => setOption("clearExistingRows", Boolean(data.checked))}
        />
        <Checkbox
          label="Keep existing cell formatting"
          checked={template.options.preserveFormatting}
          onChange={(_event, data) => setOption("preserveFormatting", Boolean(data.checked))}
        />
        <Checkbox
          label="Refresh PivotTables afterwards"
          checked={template.options.refreshPivotTables}
          onChange={(_event, data) => setOption("refreshPivotTables", Boolean(data.checked))}
        />
        <Checkbox
          label="Recalculate the workbook afterwards"
          checked={template.options.recalculate}
          onChange={(_event, data) => setOption("recalculate", Boolean(data.checked))}
        />
        <RunButton
          label="Refresh report"
          busy={runner.busy}
          disabled={!activeZone || !source}
          onClick={() =>
            void runner.run(() =>
              refreshReport(template, activeZone && source ? { [activeZone.name]: source } : {})
            )
          }
        />
        <Text className={styles.hint}>
          Refreshes the selected zone from the chosen source. Excel&apos;s undo (Ctrl+Z) still works,
          though a large refresh may take several steps to undo.
        </Text>
      </Section>
    </div>
  );
};

export default ReportPanel;
