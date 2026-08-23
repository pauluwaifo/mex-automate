import * as React from "react";
import {
  Checkbox,
  Dropdown,
  Field,
  Option,
  Radio,
  RadioGroup,
  Text,
} from "@fluentui/react-components";

import {
  CleaningScope,
  DATE_FORMATS,
  removeDuplicateRows,
  standardizeDates,
  standardizeTextCase,
  TextCaseMode,
  trimWhitespace,
} from "../features/dataCleaning";
import { ResultBanner, RunButton, Section, useActionRunner, useSharedStyles } from "./ui";
import { useTargetAddress, useTargetHeaders } from "./useSelection";

const CASE_MODES: Array<{ value: TextCaseMode; label: string }> = [
  { value: "upper", label: "UPPERCASE" },
  { value: "lower", label: "lowercase" },
  { value: "proper", label: "Proper Case" },
  { value: "sentence", label: "Sentence case" },
];

const CleanPanel: React.FC = () => {
  const styles = useSharedStyles();
  const runner = useActionRunner();

  const [scope, setScope] = React.useState<CleaningScope>("selection");
  // Bumped after every command so the "acting on" hint re-reads the grid.
  const [refreshToken, setRefreshToken] = React.useState(0);
  const address = useTargetAddress(scope, refreshToken);

  // Duplicates
  const [hasHeaderRow, setHasHeaderRow] = React.useState(true);
  const [ignoreCaseAndSpacing, setIgnoreCaseAndSpacing] = React.useState(true);
  const [deleteEntireRows, setDeleteEntireRows] = React.useState(false);
  const [keyColumns, setKeyColumns] = React.useState<string[]>([]);
  const headers = useTargetHeaders(scope, true);

  // Trim
  const [collapseInnerSpaces, setCollapseInnerSpaces] = React.useState(true);
  const [blankOutWhitespaceOnlyCells, setBlankOutWhitespaceOnlyCells] = React.useState(true);

  // Dates
  const [numberFormat, setNumberFormat] = React.useState<string>(DATE_FORMATS[0].value);
  const [dayFirst, setDayFirst] = React.useState(false);

  // Case
  const [caseMode, setCaseMode] = React.useState<TextCaseMode>("proper");

  const run = React.useCallback(
    async (action: () => Promise<Awaited<ReturnType<typeof trimWhitespace>>>) => {
      await runner.run(action);
      setRefreshToken((token) => token + 1);
    },
    [runner]
  );

  const selectedKeyIndexes = React.useMemo(
    () => keyColumns.map((label) => headers.indexOf(label)).filter((index) => index >= 0),
    [keyColumns, headers]
  );

  const dateFormatLabel =
    DATE_FORMATS.find((format) => format.value === numberFormat)?.label ?? numberFormat;

  return (
    <div>
      <Section
        title="What to clean"
        description="Cleaning runs on your selection. Select a single cell to clean the whole sheet instead."
      >
        <RadioGroup
          layout="horizontal"
          value={scope}
          onChange={(_event, data) => setScope(data.value as CleaningScope)}
        >
          <Radio value="selection" label="Selection" />
          <Radio value="usedRange" label="Whole sheet" />
        </RadioGroup>
        <Text className={styles.hint}>Acting on: {address}</Text>
      </Section>

      <ResultBanner result={runner.result} />

      <Section title="Remove duplicate rows" description="Keeps the first of each repeated row.">
        <Checkbox
          label="First row is a header"
          checked={hasHeaderRow}
          onChange={(_event, data) => setHasHeaderRow(Boolean(data.checked))}
        />
        <Checkbox
          label="Ignore case and extra spacing when comparing"
          checked={ignoreCaseAndSpacing}
          onChange={(_event, data) => setIgnoreCaseAndSpacing(Boolean(data.checked))}
        />
        <Checkbox
          label="Delete the whole worksheet row"
          checked={deleteEntireRows}
          onChange={(_event, data) => setDeleteEntireRows(Boolean(data.checked))}
        />
        <Field
          label="Compare these columns only"
          hint={keyColumns.length === 0 ? "Leave empty to compare the entire row." : undefined}
        >
          <Dropdown
            multiselect
            placeholder="Entire row"
            selectedOptions={keyColumns}
            value={keyColumns.join(", ")}
            onOptionSelect={(_event, data) => setKeyColumns(data.selectedOptions)}
            disabled={headers.length === 0}
          >
            {headers.map((header) => (
              <Option key={header} value={header}>
                {header}
              </Option>
            ))}
          </Dropdown>
        </Field>
        <RunButton
          label="Remove duplicates"
          busy={runner.busy}
          onClick={() =>
            run(() =>
              removeDuplicateRows({
                scope,
                hasHeaderRow,
                ignoreCaseAndSpacing,
                deleteEntireRows,
                keyColumns: selectedKeyIndexes,
              })
            )
          }
        />
      </Section>

      <Section
        title="Trim whitespace"
        description="Removes leading and trailing spaces, including the invisible ones that survive a web copy-paste."
      >
        <Checkbox
          label="Also collapse repeated spaces inside the text"
          checked={collapseInnerSpaces}
          onChange={(_event, data) => setCollapseInnerSpaces(Boolean(data.checked))}
        />
        <Checkbox
          label="Empty out cells that contain only spaces"
          checked={blankOutWhitespaceOnlyCells}
          onChange={(_event, data) => setBlankOutWhitespaceOnlyCells(Boolean(data.checked))}
        />
        <RunButton
          label="Trim whitespace"
          busy={runner.busy}
          onClick={() =>
            run(() => trimWhitespace({ scope, collapseInnerSpaces, blankOutWhitespaceOnlyCells }))
          }
        />
      </Section>

      <Section
        title="Standardize dates"
        description="Converts text that looks like a date into a real date, then formats every date the same way."
      >
        <Field label="Date format">
          <Dropdown
            selectedOptions={[numberFormat]}
            value={dateFormatLabel}
            onOptionSelect={(_event, data) => setNumberFormat(String(data.optionValue))}
          >
            {DATE_FORMATS.map((format) => (
              <Option key={format.value} value={format.value}>
                {format.label}
              </Option>
            ))}
          </Dropdown>
        </Field>
        <Checkbox
          label="Read 03/04/2024 as 3 April (day first)"
          checked={dayFirst}
          onChange={(_event, data) => setDayFirst(Boolean(data.checked))}
        />
        <RunButton
          label="Standardize dates"
          busy={runner.busy}
          onClick={() => run(() => standardizeDates({ scope, numberFormat, dayFirst }))}
        />
      </Section>

      <Section title="Standardize text case" description="Formulas are left untouched.">
        <Field label="Convert to">
          <Dropdown
            selectedOptions={[caseMode]}
            value={CASE_MODES.find((mode) => mode.value === caseMode)?.label ?? ""}
            onOptionSelect={(_event, data) => setCaseMode(data.optionValue as TextCaseMode)}
          >
            {CASE_MODES.map((mode) => (
              <Option key={mode.value} value={mode.value}>
                {mode.label}
              </Option>
            ))}
          </Dropdown>
        </Field>
        <RunButton
          label="Change case"
          busy={runner.busy}
          onClick={() => run(() => standardizeTextCase({ scope, mode: caseMode }))}
        />
      </Section>
    </div>
  );
};

export default CleanPanel;
