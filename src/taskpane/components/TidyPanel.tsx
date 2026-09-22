import * as React from "react";
import {
  Button,
  Checkbox,
  Dropdown,
  makeStyles,
  Option,
  Spinner,
  Switch,
  Text,
  tokens,
} from "@fluentui/react-components";
import { ArrowRight20Regular, CheckmarkCircle20Filled, TableSparkle24Regular } from "@fluentui/react-icons";

import { DataSourceRef, listSourceOptions, SourceOption } from "../features/charts";
import { countProblems, FindingId, previewTidy, TidyPreview, tidyToNewSheet } from "../features/tidy";
import type { ToolProps } from "./App";
import { ActionButton, MoreOptions, Step, Tip, useActionRunner, useSharedStyles } from "./ui";
import { useSelectionVersion } from "./useSelection";

const SELECTION_KEY = "selection";

const useStyles = makeStyles({
  headline: {
    display: "flex",
    alignItems: "baseline",
    columnGap: "8px",
  },
  bigNumber: {
    fontSize: tokens.fontSizeHero700,
    fontWeight: tokens.fontWeightBold,
    lineHeight: tokens.lineHeightHero700,
    color: tokens.colorPaletteMarigoldForeground1,
  },
  clean: {
    color: tokens.colorPaletteGreenForeground1,
  },
  finding: {
    display: "flex",
    flexDirection: "column",
    rowGap: "2px",
    padding: "8px 0",
    borderBottom: `1px solid ${tokens.colorNeutralStroke3}`,
    ":last-child": {
      borderBottom: "none",
    },
  },
  findingHead: {
    display: "flex",
    alignItems: "flex-start",
    columnGap: "6px",
  },
  fixed: {
    flexShrink: 0,
    marginTop: "6px",
    marginLeft: "8px",
    marginRight: "6px",
    color: tokens.colorPaletteGreenForeground1,
  },
  count: {
    marginLeft: "auto",
    flexShrink: 0,
    marginTop: "6px",
    padding: "0 8px",
    borderRadius: tokens.borderRadiusCircular,
    backgroundColor: tokens.colorNeutralBackground3,
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
  },
  detail: {
    paddingLeft: "32px",
    display: "flex",
    flexDirection: "column",
    rowGap: "2px",
  },
  example: {
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
  },
});

function parseSourceKey(key: string): DataSourceRef {
  if (key === SELECTION_KEY) return { kind: "selection", name: "" };
  const separator = key.indexOf(":");
  return { kind: key.slice(0, separator) as "sheet" | "table", name: key.slice(separator + 1) };
}

const TidyPanel: React.FC<ToolProps> = ({ navigate, params }) => {
  const shared = useSharedStyles();
  const styles = useStyles();
  const runner = useActionRunner();
  const selectionVersion = useSelectionVersion();

  const [sources, setSources] = React.useState<SourceOption[]>([]);
  const [sourceKey, setSourceKey] = React.useState(params?.source ?? SELECTION_KEY);
  const [preview, setPreview] = React.useState<TidyPreview | null>(null);
  const [scanning, setScanning] = React.useState(false);
  const [skip, setSkip] = React.useState<Set<FindingId>>(new Set());
  const [dayFirst, setDayFirst] = React.useState(true);
  const [cleanSheet, setCleanSheet] = React.useState<string | null>(null);

  const sourceRef = React.useMemo(() => parseSourceKey(sourceKey), [sourceKey]);

  React.useEffect(() => {
    void listSourceOptions().then(setSources);
  }, [runner.result]);

  // Rescan whenever the source changes, or the selection moves when scanning the selection.
  const selectionDependency = sourceKey === SELECTION_KEY ? selectionVersion : 0;
  React.useEffect(() => {
    let active = true;
    setScanning(true);
    void previewTidy(sourceRef, dayFirst).then((next) => {
      if (active) {
        setPreview(next);
        setScanning(false);
      }
    });
    return () => {
      active = false;
    };
  }, [sourceRef, selectionDependency, dayFirst]);

  const findings = preview?.result?.findings ?? [];
  const total = countProblems(findings);
  const toFix = countProblems(findings.filter((finding) => !skip.has(finding.id)));
  const result = preview?.result ?? null;

  const toggle = (id: FindingId, on: boolean) =>
    setSkip((current) => {
      const next = new Set(current);
      if (on) next.delete(id);
      else next.add(id);
      return next;
    });

  const sourceLabel =
    sourceKey === SELECTION_KEY
      ? "The sheet I'm on"
      : (sources.find((item) => `${item.kind}:${item.name}` === sourceKey)?.label ?? sourceKey);

  return (
    <div>
      <Step
        number={1}
        title="Which data?"
        description="Messy exports are fine: titles, totals, notes and all. The whole sheet is scanned."
      >
        <Dropdown
          value={sourceLabel}
          selectedOptions={[sourceKey]}
          onOptionSelect={(_event, data) => {
            setSourceKey(String(data.optionValue));
            setSkip(new Set());
            setCleanSheet(null);
          }}
        >
          <Option value={SELECTION_KEY} text="The sheet I'm on">
            The sheet I&apos;m on
          </Option>
          {sources.map((item) => (
            <Option key={`${item.kind}:${item.name}`} value={`${item.kind}:${item.name}`} text={item.label}>
              {item.label}
            </Option>
          ))}
        </Dropdown>
        {sourceKey === SELECTION_KEY ? (
          <Tip>Select a range first to scan only that part of the sheet.</Tip>
        ) : null}
      </Step>

      <Step number={2} title="What we found">
        {scanning && !result ? <Spinner size="tiny" label="Scanning your data..." /> : null}

        {!scanning && !result ? (
          <Text className={shared.hint}>
            No table found there. It needs a row of headings with data underneath.
          </Text>
        ) : null}

        {result && total === 0 ? (
          <div className={styles.headline}>
            <CheckmarkCircle20Filled className={styles.clean} />
            <Text weight="semibold">This table looks tidy already. Nothing to fix.</Text>
          </div>
        ) : null}

        {result && total > 0 ? (
          <>
            <div className={styles.headline}>
              <span className={styles.bigNumber}>{total.toLocaleString()}</span>
              <Text weight="semibold">
                problems in {result.rows.length.toLocaleString()} rows of &quot;{preview?.sourceName}&quot;
              </Text>
            </div>
            <div className={shared.column} style={{ rowGap: 0 }}>
              {findings.map((finding) => (
                <div key={finding.id} className={styles.finding}>
                  <div className={styles.findingHead}>
                    {finding.optional ? (
                      <Checkbox
                        checked={!skip.has(finding.id)}
                        onChange={(_event, data) => toggle(finding.id, Boolean(data.checked))}
                        label={<Text weight="semibold">{finding.label}</Text>}
                      />
                    ) : (
                      <>
                        <CheckmarkCircle20Filled className={styles.fixed} />
                        <Text weight="semibold" style={{ paddingTop: "5px" }}>
                          {finding.label}
                        </Text>
                      </>
                    )}
                    <span className={styles.count}>{finding.count.toLocaleString()}</span>
                  </div>
                  <div className={styles.detail}>
                    <Text className={shared.hint}>{finding.fix}</Text>
                    {finding.examples.map((example) => (
                      <span key={example} className={styles.example} title={example}>
                        {example}
                      </span>
                    ))}
                  </div>
                </div>
              ))}
            </div>
            <MoreOptions>
              <Switch
                label={
                  dayFirst
                    ? "Unclear dates are day first: 03/04/2024 is 3 April"
                    : "Unclear dates are month first: 03/04/2024 is March 4"
                }
                checked={dayFirst}
                onChange={(_event, data) => setDayFirst(data.checked)}
              />
              <Text className={shared.hint}>
                Only used when a column gives no clue. A date like 13/04/2024 settles it automatically.
              </Text>
            </MoreOptions>
          </>
        ) : null}
      </Step>

      <Step
        number={3}
        title="Fix it"
        waitingFor={result && result.rows.length > 0 ? undefined : "Choose some data first."}
      >
        {total > 0 ? (
          <ActionButton
            runner={runner}
            id="tidy"
            wide
            icon={<TableSparkle24Regular />}
            label={`Fix ${toFix.toLocaleString()} problems`}
            busyLabel="Fixing..."
            disabled={toFix === 0}
            onRun={async () => {
              const { result: outcome, sheetName } = await tidyToNewSheet(sourceRef, skip, dayFirst);
              setCleanSheet(outcome.ok ? sheetName : null);
              return outcome;
            }}
          />
        ) : null}
        <Tip>Your original sheet isn&apos;t changed. The clean table goes on a new sheet.</Tip>

        {cleanSheet || (result && total === 0) ? (
          <Button
            appearance={cleanSheet ? "primary" : "secondary"}
            icon={<ArrowRight20Regular />}
            iconPosition="after"
            onClick={() =>
              navigate("dashboard", { source: cleanSheet ? `sheet:${cleanSheet}` : sourceKey })
            }
          >
            Build a dashboard from it
          </Button>
        ) : null}
      </Step>
    </div>
  );
};

export default TidyPanel;
