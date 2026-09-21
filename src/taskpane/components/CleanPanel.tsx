import * as React from "react";
import {
  Button,
  Checkbox,
  Dropdown,
  Field,
  Option,
  Radio,
  RadioGroup,
  Spinner,
  Switch,
} from "@fluentui/react-components";
import {
  CalendarLtr24Regular,
  CopySelect24Regular,
  Eraser24Regular,
  TextCaseTitle24Regular,
} from "@fluentui/react-icons";

import {
  CleaningScope,
  DATE_FORMATS,
  removeDuplicateRows,
  standardizeDates,
  standardizeTextCase,
  TextCaseMode,
  trimWhitespace,
} from "../features/dataCleaning";
import {
  ActionButton,
  ActionCard,
  friendlyAddress,
  MoreOptions,
  ResultBanner,
  SelectionBar,
  Tip,
  useActionRunner,
  useSharedStyles,
} from "./ui";
import { useTargetAddress, useTargetHeaders } from "./useSelection";

type CardId = "duplicates" | "spaces" | "dates" | "case";

const CASE_BUTTONS: Array<{ mode: TextCaseMode; label: string }> = [
  { mode: "upper", label: "UPPER CASE" },
  { mode: "lower", label: "lower case" },
  { mode: "proper", label: "Title Case" },
  { mode: "sentence", label: "Sentence case" },
];

const CleanPanel: React.FC = () => {
  const styles = useSharedStyles();
  const runner = useActionRunner();

  const [openCard, setOpenCard] = React.useState<CardId | null>("duplicates");
  const [scope, setScope] = React.useState<CleaningScope>("selection");
  // Bumped after every command so the "working on" bar re-reads the grid.
  const [refreshToken, setRefreshToken] = React.useState(0);
  const address = useTargetAddress(scope, refreshToken);
  const headers = useTargetHeaders(scope, openCard === "duplicates");

  // Duplicates
  const [compareMode, setCompareMode] = React.useState<"row" | "columns">("row");
  const [keyColumns, setKeyColumns] = React.useState<string[]>([]);
  const [hasHeaderRow, setHasHeaderRow] = React.useState(true);
  const [ignoreCaseAndSpacing, setIgnoreCaseAndSpacing] = React.useState(true);
  const [deleteEntireRows, setDeleteEntireRows] = React.useState(false);

  // Spaces
  const [collapseInnerSpaces, setCollapseInnerSpaces] = React.useState(true);
  const [blankOutWhitespaceOnlyCells, setBlankOutWhitespaceOnlyCells] = React.useState(true);

  // Dates
  const [numberFormat, setNumberFormat] = React.useState<string>(DATE_FORMATS[0].value);
  const [dayFirst, setDayFirst] = React.useState(true);

  const toggle = (id: CardId) => setOpenCard((current) => (current === id ? null : id));
  const afterRun = () => setRefreshToken((token) => token + 1);

  const selectedKeyIndexes = React.useMemo(
    () => keyColumns.map((label) => headers.indexOf(label)).filter((index) => index >= 0),
    [keyColumns, headers]
  );

  const noData = address === "(no data)";
  const dateFormatLabel =
    DATE_FORMATS.find((format) => format.value === numberFormat)?.label ?? numberFormat;

  const shared = {
    runner,
    wide: true,
    disabled: noData,
    onDone: afterRun,
  };

  return (
    <div>
      <SelectionBar address={noData ? "No data found" : friendlyAddress(address)} />

      <RadioGroup
        layout="horizontal"
        value={scope}
        onChange={(_event, data) => setScope(data.value as CleaningScope)}
        aria-label="What to clean"
      >
        <Radio value="selection" label="Selected cells" />
        <Radio value="usedRange" label="Whole sheet" />
      </RadioGroup>
      <div style={{ margin: "4px 0 12px" }}>
        <Tip>Click a single cell and the whole table is cleaned.</Tip>
      </div>

      <ActionCard
        icon={<CopySelect24Regular />}
        title="Remove duplicates"
        description="Delete rows that appear more than once, keeping the first."
        open={openCard === "duplicates"}
        onToggle={() => toggle("duplicates")}
      >
        <RadioGroup
          value={compareMode}
          onChange={(_event, data) => setCompareMode(data.value as "row" | "columns")}
        >
          <Radio value="row" label="A row is a duplicate if every cell matches" />
          <Radio value="columns" label="Only compare certain columns" />
        </RadioGroup>

        {compareMode === "columns" ? (
          <Field label="Columns to compare" hint="For example, just Order ID or Email.">
            <Dropdown
              multiselect
              placeholder={headers.length === 0 ? "No columns found" : "Choose columns"}
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
        ) : null}

        <MoreOptions>
          <Checkbox
            label="The first row is headings (never remove it)"
            checked={hasHeaderRow}
            onChange={(_event, data) => setHasHeaderRow(Boolean(data.checked))}
          />
          <Checkbox
            label='Treat "ACME" and "acme " as the same'
            checked={ignoreCaseAndSpacing}
            onChange={(_event, data) => setIgnoreCaseAndSpacing(Boolean(data.checked))}
          />
          <Checkbox
            label="Delete the entire row in the sheet"
            checked={deleteEntireRows}
            onChange={(_event, data) => setDeleteEntireRows(Boolean(data.checked))}
          />
        </MoreOptions>

        <ActionButton
          {...shared}
          id="duplicates"
          label="Remove duplicates"
          disabled={noData || (compareMode === "columns" && selectedKeyIndexes.length === 0)}
          onRun={() =>
            removeDuplicateRows({
              scope,
              hasHeaderRow,
              ignoreCaseAndSpacing,
              deleteEntireRows,
              keyColumns: compareMode === "columns" ? selectedKeyIndexes : [],
            })
          }
        />
      </ActionCard>

      <ActionCard
        icon={<Eraser24Regular />}
        title="Remove extra spaces"
        description="Strip spaces from the start and end of text, including invisible ones."
        open={openCard === "spaces"}
        onToggle={() => toggle("spaces")}
      >
        <MoreOptions>
          <Checkbox
            label='Turn double spaces inside text into one ("New   York" becomes "New York")'
            checked={collapseInnerSpaces}
            onChange={(_event, data) => setCollapseInnerSpaces(Boolean(data.checked))}
          />
          <Checkbox
            label="Empty cells that contain only spaces"
            checked={blankOutWhitespaceOnlyCells}
            onChange={(_event, data) => setBlankOutWhitespaceOnlyCells(Boolean(data.checked))}
          />
        </MoreOptions>
        <ActionButton
          {...shared}
          id="spaces"
          label="Remove extra spaces"
          onRun={() => trimWhitespace({ scope, collapseInnerSpaces, blankOutWhitespaceOnlyCells })}
        />
      </ActionCard>

      <ActionCard
        icon={<CalendarLtr24Regular />}
        title="Fix dates"
        description="Turn dates typed as text into real dates, all shown the same way."
        open={openCard === "dates"}
        onToggle={() => toggle("dates")}
      >
        <Field label="Show every date as">
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
        <Switch
          label={
            dayFirst
              ? "Your dates are day first: 03/04/2024 means 3 April"
              : "Your dates are month first: 03/04/2024 means March 4"
          }
          checked={dayFirst}
          onChange={(_event, data) => setDayFirst(data.checked)}
        />
        <ActionButton
          {...shared}
          id="dates"
          label="Fix dates"
          onRun={() => standardizeDates({ scope, numberFormat, dayFirst })}
        />
      </ActionCard>

      <ActionCard
        icon={<TextCaseTitle24Regular />}
        title="Fix capital letters"
        description="Make text consistent. Formulas are never changed."
        open={openCard === "case"}
        onToggle={() => toggle("case")}
      >
        <div className={styles.grid2}>
          {CASE_BUTTONS.map((item) => (
            <Button
              key={item.mode}
              disabled={runner.busy || noData}
              icon={runner.busyId === `case-${item.mode}` ? <Spinner size="tiny" /> : undefined}
              onClick={() =>
                void runner
                  .run(`case-${item.mode}`, () => standardizeTextCase({ scope, mode: item.mode }))
                  .then(afterRun)
              }
            >
              {item.label}
            </Button>
          ))}
        </div>
        {runner.result && runner.resultId?.startsWith("case-") ? (
          // The four buttons share one result, drawn once below the grid.
          <ResultBanner result={runner.result} onDismiss={runner.clear} />
        ) : null}
      </ActionCard>
    </div>
  );
};

export default CleanPanel;
