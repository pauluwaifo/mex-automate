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
  Spinner,
  Switch,
  Text,
  tokens,
} from "@fluentui/react-components";
import {
  ArrowUpload20Regular,
  CheckmarkCircle16Filled,
  Dismiss16Regular,
  DocumentBulletList20Regular,
  Link16Regular,
  TableStackBelow24Regular,
  Warning16Filled,
} from "@fluentui/react-icons";

import {
  DEFAULT_MERGE_OPTIONS,
  HeaderPlan,
  listSheetNames,
  mergeFiles,
  mergeSheets,
  planHeaders,
  previewSheetMerge,
  SOURCE_COLUMN_HEADER,
  SourceTable,
} from "../features/merge";
import { readFileAsSources } from "../features/reportBuilder";
import { normalizeHeader } from "../shared/excelHelpers";
import { SUPPORTED_EXTENSIONS } from "../shared/workbookReader";
import { ActionButton, MoreOptions, Step, Tip, useActionRunner, useSharedStyles } from "./ui";

type MergeMode = "sheets" | "files";

/** "Client Name" should be treated as "Customer". */
interface ColumnMatch {
  from: string;
  to: string;
}

const useStyles = makeStyles({
  columnRow: {
    display: "flex",
    flexDirection: "column",
    rowGap: "4px",
    padding: "8px 0",
    borderBottom: `1px solid ${tokens.colorNeutralStroke3}`,
    ":last-child": {
      borderBottom: "none",
    },
  },
  columnName: {
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
  matchRow: {
    display: "flex",
    alignItems: "center",
    columnGap: "6px",
    fontSize: tokens.fontSizeBase200,
  },
  fileRow: {
    display: "flex",
    alignItems: "center",
    columnGap: "6px",
  },
  fileName: {
    flexGrow: 1,
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  hiddenInput: {
    display: "none",
  },
});

const MergePanel: React.FC = () => {
  const shared = useSharedStyles();
  const styles = useStyles();
  const runner = useActionRunner();

  const [mode, setMode] = React.useState<MergeMode>("sheets");
  const [sheetNames, setSheetNames] = React.useState<string[]>([]);
  const [selectedSheets, setSelectedSheets] = React.useState<string[]>([]);
  const [files, setFiles] = React.useState<File[]>([]);
  const [fileSources, setFileSources] = React.useState<SourceTable[]>([]);
  const [fileError, setFileError] = React.useState("");

  const [matches, setMatches] = React.useState<ColumnMatch[]>([]);
  const [plan, setPlan] = React.useState<HeaderPlan | null>(null);
  const [planning, setPlanning] = React.useState(false);

  const [destinationSheetName, setDestinationSheetName] = React.useState(
    DEFAULT_MERGE_OPTIONS.destinationSheetName
  );
  const [sortByHeader, setSortByHeader] = React.useState("");
  const [sortDescending, setSortDescending] = React.useState(false);
  const [addSourceColumn, setAddSourceColumn] = React.useState(DEFAULT_MERGE_OPTIONS.addSourceColumn);
  const [skipBlankRows, setSkipBlankRows] = React.useState(DEFAULT_MERGE_OPTIONS.skipBlankRows);

  const fileInputRef = React.useRef<HTMLInputElement | null>(null);

  React.useEffect(() => {
    void listSheetNames().then(setSheetNames);
  }, [runner.result]);

  const headerAliases = React.useMemo(() => {
    const aliases: Record<string, string> = {};
    for (const match of matches) {
      aliases[normalizeHeader(match.from)] = match.to;
    }
    return aliases;
  }, [matches]);

  const enoughChosen = mode === "sheets" ? selectedSheets.length >= 2 : fileSources.length > 0;

  // Work out the combined columns as soon as the sources are known - no
  // separate "preview" button to discover.
  React.useEffect(() => {
    if (!enoughChosen) {
      setPlan(null);
      return undefined;
    }
    if (mode === "files") {
      setPlan(planHeaders(fileSources, headerAliases));
      return undefined;
    }
    let active = true;
    setPlanning(true);
    void previewSheetMerge(selectedSheets, headerAliases).then((next) => {
      if (active) {
        setPlan(next);
        setPlanning(false);
      }
    });
    return () => {
      active = false;
    };
  }, [mode, selectedSheets, fileSources, headerAliases, enoughChosen]);

  // A sort column that no longer exists after re-planning is quietly dropped.
  React.useEffect(() => {
    if (sortByHeader && plan && sortByHeader !== SOURCE_COLUMN_HEADER && !plan.headers.includes(sortByHeader)) {
      setSortByHeader("");
    }
  }, [plan, sortByHeader]);

  const toggleSheet = (name: string, checked: boolean) =>
    setSelectedSheets((current) =>
      checked ? [...current, name] : current.filter((sheet) => sheet !== name)
    );

  const readFiles = async (picked: File[]) => {
    setFiles(picked);
    setFileError("");
    const tables: SourceTable[] = [];
    const problems: string[] = [];
    for (const file of picked) {
      try {
        tables.push(...(await readFileAsSources(file)));
      } catch (error) {
        problems.push(`${file.name}: ${error instanceof Error ? error.message : String(error)}`);
      }
    }
    setFileSources(tables);
    setFileError(problems.join(" "));
  };

  const removeFile = (name: string) => void readFiles(files.filter((file) => file.name !== name));

  const sourceCount = mode === "sheets" ? selectedSheets.length : fileSources.length;
  const partial = new Set(plan?.partialHeaders ?? []);
  const sortableHeaders = plan
    ? addSourceColumn
      ? [SOURCE_COLUMN_HEADER, ...plan.headers]
      : plan.headers
    : [];

  const options = {
    addSourceColumn,
    skipBlankRows,
    destinationSheetName,
    headerAliases,
    sortByHeader,
    sortDescending,
  };

  return (
    <div>
      <Step number={1} title="Choose what to combine">
        <RadioGroup
          value={mode}
          onChange={(_event, data) => {
            setMode(data.value as MergeMode);
            setMatches([]);
          }}
        >
          <Radio value="sheets" label="Sheets in this workbook" />
          <Radio value="files" label="Files from my computer" />
        </RadioGroup>

        {mode === "sheets" ? (
          <>
            <div className={shared.row}>
              <Text className={shared.hint}>Tick two or more sheets.</Text>
              <Button
                size="small"
                appearance="transparent"
                onClick={() =>
                  setSelectedSheets(selectedSheets.length === sheetNames.length ? [] : sheetNames)
                }
              >
                {selectedSheets.length === sheetNames.length ? "Clear all" : "Select all"}
              </Button>
            </div>
            <div className={shared.scrollList}>
              {sheetNames.length === 0 ? (
                <Text className={shared.hint}>No sheets found.</Text>
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
          </>
        ) : (
          <>
            <input
              ref={fileInputRef}
              className={styles.hiddenInput}
              type="file"
              multiple
              accept={SUPPORTED_EXTENSIONS.join(",")}
              onChange={(event) => {
                void readFiles(Array.from(event.target.files ?? []));
                event.target.value = "";
              }}
            />
            <Button icon={<ArrowUpload20Regular />} onClick={() => fileInputRef.current?.click()}>
              {files.length > 0 ? "Choose different files" : "Choose files"}
            </Button>
            <Text className={shared.hint}>
              Excel or CSV files ({SUPPORTED_EXTENSIONS.join(", ")}). They are read on this computer
              only.
            </Text>
            {files.length > 0 ? (
              <div className={shared.scrollList}>
                {files.map((file) => (
                  <div key={file.name} className={styles.fileRow}>
                    <DocumentBulletList20Regular />
                    <Text className={styles.fileName} title={file.name}>
                      {file.name}
                    </Text>
                    <Button
                      size="small"
                      appearance="subtle"
                      icon={<Dismiss16Regular />}
                      aria-label={`Remove ${file.name}`}
                      onClick={() => removeFile(file.name)}
                    />
                  </div>
                ))}
              </div>
            ) : null}
            {fileError ? (
              <Text className={shared.hint} style={{ color: tokens.colorPaletteRedForeground1 }}>
                {fileError}
              </Text>
            ) : null}
          </>
        )}
      </Step>

      <Step
        number={2}
        title="Check the columns"
        description="Columns with the same name are lined up automatically, even if the spelling or order differs."
        waitingFor={
          enoughChosen
            ? undefined
            : mode === "sheets"
              ? "Tick at least two sheets above."
              : "Choose at least one file above."
        }
      >
        {planning && !plan ? <Spinner size="tiny" label="Reading your sheets..." /> : null}
        {plan && plan.headers.length === 0 ? (
          <Text className={shared.hint}>No columns found. Check that each source has a heading row.</Text>
        ) : null}

        {plan && plan.headers.length > 0 ? (
          <>
            <Text className={shared.hint}>
              {partial.size === 0
                ? `All ${plan.headers.length} columns appear in every ${mode === "sheets" ? "sheet" : "file"}.`
                : `${plan.headers.length} columns. ${partial.size} ${partial.size === 1 ? "is" : "are"} missing from some sources. If one is the same as another column under a different name, match them up.`}
            </Text>

            <div className={shared.scrollList}>
              {plan.headers.map((header) => {
                const isPartial = partial.has(header);
                const from = plan.contributors[header] ?? [];
                return (
                  <div key={header} className={styles.columnRow}>
                    <span className={styles.columnName}>
                      {isPartial ? (
                        <Warning16Filled className={styles.warn} />
                      ) : (
                        <CheckmarkCircle16Filled className={styles.ok} />
                      )}
                      {header}
                    </span>
                    {isPartial ? (
                      <>
                        <Text className={shared.hint}>
                          Only in {from.join(", ")}. Blank for the others.
                        </Text>
                        <Dropdown
                          size="small"
                          placeholder="Same as another column?"
                          value=""
                          selectedOptions={[]}
                          onOptionSelect={(_event, data) => {
                            const to = String(data.optionValue);
                            setMatches((current) => [
                              ...current.filter((item) => item.from !== header),
                              { from: header, to },
                            ]);
                          }}
                        >
                          {plan.headers
                            .filter((other) => other !== header)
                            .map((other) => (
                              <Option key={other} value={other} text={`Same as ${other}`}>
                                Same as {other}
                              </Option>
                            ))}
                        </Dropdown>
                      </>
                    ) : null}
                  </div>
                );
              })}
            </div>

            {matches.length > 0 ? (
              <div className={shared.column}>
                <Text className={shared.hint}>Columns you matched up:</Text>
                {matches.map((match) => (
                  <div key={match.from} className={styles.matchRow}>
                    <Link16Regular />
                    <span>
                      <strong>{match.from}</strong> goes into <strong>{match.to}</strong>
                    </span>
                    <Button
                      size="small"
                      appearance="transparent"
                      onClick={() =>
                        setMatches((current) => current.filter((item) => item.from !== match.from))
                      }
                    >
                      Undo
                    </Button>
                  </div>
                ))}
              </div>
            ) : null}
          </>
        ) : null}
      </Step>

      <Step
        number={3}
        title="Create the combined table"
        waitingFor={plan && plan.headers.length > 0 ? undefined : "Finish the steps above first."}
      >
        <Field label="Put it on a new sheet called">
          <Input
            value={destinationSheetName}
            onChange={(_event, data) => setDestinationSheetName(data.value)}
          />
        </Field>

        <Field label="Sort rows by">
          <Dropdown
            value={sortByHeader === "" ? "Keep original order" : sortByHeader}
            selectedOptions={[sortByHeader]}
            onOptionSelect={(_event, data) => setSortByHeader(String(data.optionValue))}
          >
            <Option value="" text="Keep original order">
              Keep original order
            </Option>
            {sortableHeaders.map((header) => (
              <Option key={header} value={header}>
                {header}
              </Option>
            ))}
          </Dropdown>
        </Field>
        {sortByHeader ? (
          <Switch
            label={sortDescending ? "Largest / Z first" : "Smallest / A first"}
            checked={sortDescending}
            onChange={(_event, data) => setSortDescending(data.checked)}
          />
        ) : null}

        <MoreOptions>
          <Checkbox
            label={`Add a "${SOURCE_COLUMN_HEADER}" column showing where each row came from`}
            checked={addSourceColumn}
            onChange={(_event, data) => setAddSourceColumn(Boolean(data.checked))}
          />
          <Checkbox
            label="Leave out empty rows"
            checked={skipBlankRows}
            onChange={(_event, data) => setSkipBlankRows(Boolean(data.checked))}
          />
        </MoreOptions>

        <ActionButton
          runner={runner}
          id="merge"
          wide
          icon={<TableStackBelow24Regular />}
          label={`Combine ${sourceCount} ${mode === "sheets" ? "sheets" : "sources"}`}
          busyLabel="Combining..."
          disabled={destinationSheetName.trim() === ""}
          onRun={() =>
            mode === "sheets" ? mergeSheets(selectedSheets, options) : mergeFiles(files, options)
          }
        />
        <Tip>Your original sheets are not changed. The result goes on a new sheet.</Tip>
      </Step>
    </div>
  );
};

export default MergePanel;
