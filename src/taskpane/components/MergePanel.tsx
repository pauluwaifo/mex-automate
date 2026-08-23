import * as React from "react";
import { Button, Checkbox, Field, Input, Radio, RadioGroup, Text, Textarea } from "@fluentui/react-components";

import {
  DEFAULT_MERGE_OPTIONS,
  HeaderPlan,
  listSheetNames,
  mergeFiles,
  mergeSheets,
  parseAliasLines,
  previewSheetMerge,
} from "../features/merge";
import { SUPPORTED_EXTENSIONS } from "../shared/workbookReader";
import { ResultBanner, RunButton, Section, useActionRunner, useSharedStyles } from "./ui";

type MergeMode = "sheets" | "files";

const MergePanel: React.FC = () => {
  const styles = useSharedStyles();
  const runner = useActionRunner();

  const [mode, setMode] = React.useState<MergeMode>("sheets");
  const [sheetNames, setSheetNames] = React.useState<string[]>([]);
  const [selectedSheets, setSelectedSheets] = React.useState<string[]>([]);
  const [files, setFiles] = React.useState<File[]>([]);
  const [plan, setPlan] = React.useState<HeaderPlan | null>(null);

  const [addSourceColumn, setAddSourceColumn] = React.useState(DEFAULT_MERGE_OPTIONS.addSourceColumn);
  const [skipBlankRows, setSkipBlankRows] = React.useState(DEFAULT_MERGE_OPTIONS.skipBlankRows);
  const [destinationSheetName, setDestinationSheetName] = React.useState(
    DEFAULT_MERGE_OPTIONS.destinationSheetName
  );
  const [aliasText, setAliasText] = React.useState("");

  const fileInputRef = React.useRef<HTMLInputElement | null>(null);

  React.useEffect(() => {
    void listSheetNames().then(setSheetNames);
  }, [runner.result]);

  const options = React.useMemo(
    () => ({
      addSourceColumn,
      skipBlankRows,
      destinationSheetName,
      headerAliases: parseAliasLines(aliasText),
    }),
    [addSourceColumn, skipBlankRows, destinationSheetName, aliasText]
  );

  const toggleSheet = (name: string, checked: boolean) => {
    setPlan(null);
    setSelectedSheets((current) =>
      checked ? [...current, name] : current.filter((sheet) => sheet !== name)
    );
  };

  const onPickFiles = (event: React.ChangeEvent<HTMLInputElement>) => {
    setPlan(null);
    setFiles(Array.from(event.target.files ?? []));
  };

  const canMerge = mode === "sheets" ? selectedSheets.length >= 2 : files.length > 0;

  return (
    <div>
      <Section
        title="What to combine"
        description="Columns are matched by header name, so sources can list them in any order."
      >
        <RadioGroup
          value={mode}
          onChange={(_event, data) => {
            setMode(data.value as MergeMode);
            setPlan(null);
          }}
        >
          <Radio value="sheets" label="Sheets in this workbook" />
          <Radio value="files" label="Files from my computer" />
        </RadioGroup>
      </Section>

      <ResultBanner result={runner.result} />

      {mode === "sheets" ? (
        <Section title="Sheets" description="Pick two or more. The first row of each is its header row.">
          <div className={styles.scrollList}>
            {sheetNames.length === 0 ? (
              <Text className={styles.hint}>No sheets found.</Text>
            ) : (
              sheetNames.map((name) => (
                <Checkbox
                  key={name}
                  label={name}
                  checked={selectedSheets.includes(name)}
                  onChange={(_event, data) => toggleSheet(name, Boolean(data.checked))}
                />
              ))
            )}
          </div>
          <Button
            appearance="secondary"
            disabled={selectedSheets.length < 2 || runner.busy}
            onClick={() => {
              void previewSheetMerge(selectedSheets, options.headerAliases).then(setPlan);
            }}
          >
            Preview columns
          </Button>
        </Section>
      ) : (
        <Section
          title="Files"
          description={`Read on this machine only - nothing is uploaded. Supported: ${SUPPORTED_EXTENSIONS.join(", ")}`}
        >
          <input
            ref={fileInputRef}
            type="file"
            multiple
            accept={SUPPORTED_EXTENSIONS.join(",")}
            onChange={onPickFiles}
          />
          {files.length > 0 ? (
            <div className={styles.scrollList}>
              {files.map((file) => (
                <Text key={file.name} block>
                  {file.name}
                </Text>
              ))}
            </div>
          ) : null}
        </Section>
      )}

      {plan ? (
        <Section title="Planned columns" description={`${plan.headers.length} column(s) after merging.`}>
          <div className={styles.code}>{plan.headers.join("\n")}</div>
          {plan.partialHeaders.length > 0 ? (
            <Text className={styles.hint}>
              Missing from some sheets (blank there): {plan.partialHeaders.join(", ")}
            </Text>
          ) : (
            <Text className={styles.hint}>Every sheet has every column.</Text>
          )}
        </Section>
      ) : null}

      <Section title="Options">
        <Checkbox
          label="Add a Source column recording where each row came from"
          checked={addSourceColumn}
          onChange={(_event, data) => setAddSourceColumn(Boolean(data.checked))}
        />
        <Checkbox
          label="Skip blank rows"
          checked={skipBlankRows}
          onChange={(_event, data) => setSkipBlankRows(Boolean(data.checked))}
        />
        <Field label="Put the result on a new sheet named">
          <Input
            value={destinationSheetName}
            onChange={(_event, data) => setDestinationSheetName(data.value)}
          />
        </Field>
        <Field
          label="Header aliases"
          hint="One per line, as: Client Name = Customer. Use this when sources name the same column differently."
        >
          <Textarea
            resize="vertical"
            rows={3}
            placeholder="Client Name = Customer"
            value={aliasText}
            onChange={(_event, data) => {
              setAliasText(data.value);
              setPlan(null);
            }}
          />
        </Field>
        <RunButton
          label={mode === "sheets" ? "Merge sheets" : "Merge files"}
          busy={runner.busy}
          disabled={!canMerge}
          onClick={() =>
            void runner.run(() =>
              mode === "sheets" ? mergeSheets(selectedSheets, options) : mergeFiles(files, options)
            )
          }
        />
      </Section>
    </div>
  );
};

export default MergePanel;
