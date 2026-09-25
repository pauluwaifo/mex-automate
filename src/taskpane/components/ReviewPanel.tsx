import * as React from "react";
import { Button, Dropdown, makeStyles, mergeClasses, Option, Spinner, Text, tokens } from "@fluentui/react-components";
import {
  ArrowCounterclockwise20Regular,
  ArrowSync20Regular,
  CheckmarkCircle20Filled,
  ChevronDown20Regular,
  ChevronRight20Regular,
  Dismiss16Regular,
  TableSearch24Regular,
} from "@fluentui/react-icons";

import { listSheetNames } from "../features/merge";
import { groupIssues, Issue, IssueGroup, summarize } from "../features/review";
import {
  applyFixes,
  clearMarks,
  goToGroup,
  goToIssue,
  ignoreIssues,
  markIssues,
  reviewSheet,
  SheetReview,
  undoLastFix,
} from "../features/reviewSheet";
import type { ToolProps } from "./App";
import { ActionButton, ResultBanner, Tip, useActionRunner, useSharedStyles } from "./ui";

const useStyles = makeStyles({
  summary: {
    display: "flex",
    alignItems: "baseline",
    columnGap: "8px",
  },
  big: {
    fontSize: tokens.fontSizeHero700,
    lineHeight: tokens.lineHeightHero700,
    fontWeight: tokens.fontWeightBold,
    fontVariantNumeric: "tabular-nums",
  },
  clean: { color: tokens.colorPaletteGreenForeground1 },
  group: {
    borderBottom: `1px solid ${tokens.colorNeutralStroke3}`,
    paddingBottom: "8px",
    marginBottom: "8px",
    ":last-child": { borderBottom: "none", marginBottom: 0, paddingBottom: 0 },
  },
  groupHead: {
    display: "flex",
    alignItems: "flex-start",
    columnGap: "8px",
    width: "100%",
    padding: "4px 0",
    border: "none",
    background: "none",
    textAlign: "left",
    cursor: "pointer",
    color: tokens.colorNeutralForeground1,
    fontFamily: "inherit",
  },
  severity: {
    flexShrink: 0,
    width: "8px",
    height: "8px",
    borderRadius: "2px",
    marginTop: "6px",
  },
  error: { backgroundColor: tokens.colorPaletteRedForeground1 },
  warning: { backgroundColor: tokens.colorPaletteMarigoldForeground1 },
  tidy: { backgroundColor: tokens.colorBrandForeground1 },
  groupText: { display: "flex", flexDirection: "column", flexGrow: 1, minWidth: 0 },
  groupTitle: { fontWeight: tokens.fontWeightSemibold, fontSize: tokens.fontSizeBase300 },
  count: {
    flexShrink: 0,
    padding: "0 7px",
    borderRadius: tokens.borderRadiusCircular,
    backgroundColor: tokens.colorNeutralBackground3,
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
  },
  actions: { display: "flex", flexWrap: "wrap", gap: "6px", paddingLeft: "16px", marginTop: "4px" },
  issue: {
    display: "flex",
    alignItems: "flex-start",
    columnGap: "6px",
    padding: "4px 0 4px 16px",
    fontSize: tokens.fontSizeBase200,
  },
  address: {
    flexShrink: 0,
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorBrandForeground1,
    cursor: "pointer",
    background: "none",
    border: "none",
    padding: 0,
    textDecoration: "underline",
  },
  detail: { color: tokens.colorNeutralForeground2, minWidth: 0 },
  toolbar: { display: "flex", flexWrap: "wrap", gap: "6px" },
});

const ReviewPanel: React.FC<ToolProps> = ({ navigate, params }) => {
  const shared = useSharedStyles();
  const styles = useStyles();
  const runner = useActionRunner();

  const [sheets, setSheets] = React.useState<string[]>([]);
  const [sheet, setSheet] = React.useState<string>(params?.source?.replace(/^sheet:/, "") ?? "");
  const [review, setReview] = React.useState<SheetReview | null>(null);
  const [scanning, setScanning] = React.useState(true);
  const [open, setOpen] = React.useState<string | null>(null);
  const [marked, setMarked] = React.useState(false);

  React.useEffect(() => {
    void listSheetNames().then(setSheets);
  }, []);

  const scan = React.useCallback(
    async (name?: string) => {
      setScanning(true);
      const result = await reviewSheet(name || undefined);
      setReview(result);
      setSheet(result?.sheet ?? name ?? "");
      setScanning(false);
      return result;
    },
    []
  );

  React.useEffect(() => {
    void scan(params?.source?.replace(/^sheet:/, ""));
  }, [scan, params]);

  const groups = React.useMemo(() => (review ? groupIssues(review.issues) : []), [review]);
  const fixable = groups.reduce((total, group) => total + group.fixable, 0);

  const afterChange = async () => {
    const next = await scan(sheet);
    if (marked && next && next.issues.length > 0) await markIssues(next.sheet, next.issues);
    else if (marked) setMarked(false);
  };

  return (
    <div>
      <section className={shared.card}>
        <div className={shared.row}>
          <TableSearch24Regular />
          <Text weight="semibold">Check a sheet for mistakes</Text>
        </div>
        <Dropdown
          value={sheet}
          selectedOptions={sheet ? [sheet] : []}
          onOptionSelect={(_event, data) => void scan(String(data.optionValue))}
        >
          {sheets.map((name) => (
            <Option key={name} value={name}>
              {name}
            </Option>
          ))}
        </Dropdown>

        {scanning ? <Spinner size="tiny" label="Reading the sheet..." /> : null}

        {!scanning && !review ? (
          <Text className={shared.hint}>No table found there. It needs a row of headings with data underneath.</Text>
        ) : null}

        {review && review.issues.length === 0 ? (
          <div className={styles.summary}>
            <CheckmarkCircle20Filled className={styles.clean} />
            <Text weight="semibold">
              Nothing looks wrong in {review.rows.toLocaleString()} rows.
              {review.ignored > 0 ? ` (${review.ignored} ignored.)` : ""}
            </Text>
          </div>
        ) : null}

        {review && review.issues.length > 0 ? (
          <>
            <div className={styles.summary}>
              <span className={styles.big}>{review.issues.length.toLocaleString()}</span>
              <Text weight="semibold">to look at: {summarize(review.issues)}</Text>
            </div>
            <div className={styles.toolbar}>
              <ActionButton
                runner={runner}
                id="mark"
                appearance="secondary"
                label={marked ? "Re-mark in sheet" : "Mark them in the sheet"}
                busyLabel="Marking..."
                onRun={async () => {
                  const result = await markIssues(review.sheet, review.issues);
                  setMarked(result.ok);
                  return result;
                }}
              />
              {marked ? (
                <ActionButton
                  runner={runner}
                  id="unmark"
                  appearance="secondary"
                  label="Clear marks"
                  onRun={async () => {
                    const result = await clearMarks();
                    setMarked(false);
                    return result;
                  }}
                />
              ) : null}
              <Button appearance="secondary" icon={<ArrowSync20Regular />} onClick={() => void scan(sheet)}>
                Check again
              </Button>
              <ActionButton
                runner={runner}
                id="undo"
                appearance="secondary"
                icon={<ArrowCounterclockwise20Regular />}
                label="Undo last fix"
                onRun={async () => {
                  const result = await undoLastFix();
                  if (result.ok) await afterChange();
                  return result;
                }}
              />
            </div>
            {fixable > 0 ? (
              <ActionButton
                runner={runner}
                id="fixAll"
                wide
                label={`Fix the ${fixable} safe ones`}
                busyLabel="Fixing..."
                onRun={async () => {
                  const issues = groups.flatMap((group) => group.issues).filter((issue) => issue.fix);
                  const result = await applyFixes(review.sheet, issues, `fixed ${issues.length} cells`);
                  if (result.ok) await afterChange();
                  return result;
                }}
              />
            ) : null}
            <Tip>
              Red is an error that changes your numbers, amber is worth checking, blue is tidying. Nothing is changed
              until you say so, and Undo puts cells back exactly as they were.
            </Tip>
          </>
        ) : null}
      </section>

      {groups.map((group) => (
        <GroupCard
          key={group.group}
          group={group}
          sheet={review!.sheet}
          runner={runner}
          open={open === group.group}
          onToggle={() => setOpen(open === group.group ? null : group.group)}
          onChanged={afterChange}
        />
      ))}

      {review && review.structuralRows > 0 ? (
        <section className={shared.card}>
          <Text className={shared.hint}>
            This sheet also has {review.structuralRows.toLocaleString()} rows that aren&apos;t data: titles, totals or
            group headings. Fix messy data handles those.
          </Text>
          <Button appearance="secondary" onClick={() => navigate("tidy", { source: `sheet:${review.sheet}` })}>
            Open Fix messy data
          </Button>
        </section>
      ) : null}
    </div>
  );
};

const GroupCard: React.FC<{
  group: IssueGroup;
  sheet: string;
  runner: ReturnType<typeof useActionRunner>;
  open: boolean;
  onToggle: () => void;
  onChanged: () => Promise<void>;
}> = ({ group, sheet, runner, open, onToggle, onChanged }) => {
  const shared = useSharedStyles();
  const styles = useStyles();

  return (
    <section className={shared.card}>
      <button type="button" className={styles.groupHead} onClick={onToggle} aria-expanded={open}>
        <span className={mergeClasses(styles.severity, styles[group.severity])} />
        <span className={styles.groupText}>
          <span className={styles.groupTitle}>
            {group.title}
            {group.header ? ` in ${group.header}` : ""}
          </span>
          <span className={shared.hint}>
            {group.issues.length === 1 ? group.issues[0].address : `${group.issues.length} cells`}
            {group.fixable > 0 ? ` · ${group.fixable} fixable` : " · needs your call"}
          </span>
        </span>
        <span className={styles.count}>{group.issues.length}</span>
        {open ? <ChevronDown20Regular /> : <ChevronRight20Regular />}
      </button>

      <div className={styles.actions}>
        <Button size="small" appearance="secondary" onClick={() => void goToGroup(sheet, group)}>
          Go to
        </Button>
        {group.fixable > 0 ? (
          <ActionButton
            runner={runner}
            id={`fix-${group.group}`}
            appearance="primary"
            label={`Fix ${group.fixable}`}
            busyLabel="Fixing..."
            onRun={async () => {
              const issues = group.issues.filter((issue) => issue.fix);
              const result = await applyFixes(sheet, issues, `fixed ${issues.length} cells`);
              if (result.ok) await onChanged();
              return result;
            }}
          />
        ) : null}
        <ActionButton
          runner={runner}
          id={`ignore-${group.group}`}
          appearance="secondary"
          icon={<Dismiss16Regular />}
          label="Ignore"
          onRun={async () => {
            const result = await ignoreIssues(sheet, group.issues);
            if (result.ok) await onChanged();
            return result;
          }}
        />
      </div>

      {open ? (
        <div>
          {group.issues.slice(0, 40).map((issue: Issue) => (
            <div key={issue.id} className={styles.issue}>
              <button type="button" className={styles.address} onClick={() => void goToIssue(sheet, issue)}>
                {issue.address}
              </button>
              <span className={styles.detail}>{issue.detail}</span>
            </div>
          ))}
          {group.issues.length > 40 ? (
            <Text className={shared.hint} style={{ paddingLeft: "16px" }}>
              ...and {(group.issues.length - 40).toLocaleString()} more.
            </Text>
          ) : null}
        </div>
      ) : null}

      {runner.result && runner.resultId?.endsWith(group.group) ? (
        <ResultBanner result={runner.result} onDismiss={runner.clear} />
      ) : null}
    </section>
  );
};

export default ReviewPanel;
