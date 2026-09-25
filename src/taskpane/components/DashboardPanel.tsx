import * as React from "react";
import {
  Button,
  Checkbox,
  Dropdown,
  Field,
  Input,
  makeStyles,
  mergeClasses,
  Option,
  Spinner,
  Switch,
  Text,
  tokens,
} from "@fluentui/react-components";
import {
  Add20Regular,
  ArrowSync20Regular,
  Board24Regular,
  DataBarHorizontal20Regular,
  DataBarVertical20Regular,
  DataHistogram20Regular,
  DataLine20Regular,
  DataPie20Regular,
  DataScatter20Regular,
} from "@fluentui/react-icons";

import { DataSourceRef, listSourceOptions, SourceOption } from "../features/charts";
import {
  analyzeDashboard,
  buildDashboard,
  customChart,
  DashboardAnalysis,
  DashboardConfig,
  DashChart,
  DashChartKind,
  listDashboards,
  MAX_CHARTS,
  refreshDashboard,
} from "../features/dashboard";
import type { ToolProps } from "./App";
import { ActionButton, MoreOptions, Step, Tip, useActionRunner, useSharedStyles } from "./ui";
import { useSelectionVersion } from "./useSelection";

const SELECTION_KEY = "selection";

const KIND_LABEL: Record<DashChartKind, string> = {
  column: "Column",
  bar: "Bar",
  line: "Line",
  pie: "Pie",
  doughnut: "Doughnut",
  stackedColumn: "Stacked column",
  scatter: "Scatter",
};

const KIND_ICON: Record<DashChartKind, React.ReactElement> = {
  line: <DataLine20Regular />,
  column: <DataBarVertical20Regular />,
  bar: <DataBarHorizontal20Regular />,
  stackedColumn: <DataHistogram20Regular />,
  pie: <DataPie20Regular />,
  doughnut: <DataPie20Regular />,
  scatter: <DataScatter20Regular />,
};

const useStyles = makeStyles({
  suggestion: {
    display: "flex",
    alignItems: "flex-start",
    columnGap: "8px",
    padding: "7px 8px",
    borderRadius: tokens.borderRadiusMedium,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    cursor: "pointer",
    backgroundColor: tokens.colorNeutralBackground1,
    textAlign: "left",
    width: "100%",
    fontFamily: "inherit",
    color: tokens.colorNeutralForeground1,
  },
  off: {
    opacity: 0.55,
  },
  kindIcon: {
    flexShrink: 0,
    width: "30px",
    height: "30px",
    borderRadius: tokens.borderRadiusMedium,
    display: "grid",
    placeItems: "center",
    backgroundColor: tokens.colorBrandBackground2,
    color: tokens.colorBrandForeground2,
  },
  text: {
    display: "flex",
    flexDirection: "column",
    flexGrow: 1,
    minWidth: 0,
  },
  list: {
    display: "flex",
    flexDirection: "column",
    rowGap: "6px",
  },
  saved: {
    display: "flex",
    alignItems: "center",
    columnGap: "8px",
  },
  savedText: {
    display: "flex",
    flexDirection: "column",
    flexGrow: 1,
    minWidth: 0,
  },
  kpis: {
    display: "flex",
    flexWrap: "wrap",
    gap: "4px",
  },
  kpiChip: {
    padding: "2px 8px",
    borderRadius: tokens.borderRadiusCircular,
    backgroundColor: tokens.colorNeutralBackground3,
    fontSize: tokens.fontSizeBase200,
  },
});

function parseSourceKey(key: string): DataSourceRef {
  if (key === SELECTION_KEY) return { kind: "selection", name: "" };
  const separator = key.indexOf(":");
  return { kind: key.slice(0, separator) as "sheet" | "table", name: key.slice(separator + 1) };
}

const DashboardPanel: React.FC<ToolProps> = ({ navigate, params }) => {
  const shared = useSharedStyles();
  const styles = useStyles();
  const runner = useActionRunner();
  const selectionVersion = useSelectionVersion();

  const [sources, setSources] = React.useState<SourceOption[]>([]);
  const [saved, setSaved] = React.useState<DashboardConfig[]>([]);
  const [sourceKey, setSourceKey] = React.useState(params?.source ?? SELECTION_KEY);
  const [analysis, setAnalysis] = React.useState<DashboardAnalysis | null>(null);
  const [loading, setLoading] = React.useState(false);
  const [chosen, setChosen] = React.useState<Set<string>>(new Set());
  const [includeKpis, setIncludeKpis] = React.useState(true);
  // Shaping controls: which columns to use, how many charts, and charts the
  // user builds themselves.
  const [useColumns, setUseColumns] = React.useState<string[]>([]);
  const [maxCharts, setMaxCharts] = React.useState<number>(MAX_CHARTS);
  const [extra, setExtra] = React.useState<DashChart[]>([]);
  const [newKind, setNewKind] = React.useState<DashChartKind>("column");
  const [newMeasure, setNewMeasure] = React.useState<string>("");
  const [newDimension, setNewDimension] = React.useState<string>("");
  const [title, setTitle] = React.useState("");
  const [titleEdited, setTitleEdited] = React.useState(false);

  const sourceRef = React.useMemo(() => parseSourceKey(sourceKey), [sourceKey]);

  React.useEffect(() => {
    void listSourceOptions().then(setSources);
    void listDashboards().then(setSaved);
  }, [runner.result]);

  const selectionDependency = sourceKey === SELECTION_KEY ? selectionVersion : 0;
  React.useEffect(() => {
    let active = true;
    setLoading(true);
    void analyzeDashboard(sourceRef, { max: maxCharts, useColumns }).then((next) => {
      if (!active) return;
      setAnalysis(next);
      setLoading(false);
      // Every suggestion starts ticked; the user unticks what they don't want.
      setChosen(new Set(next.plan?.charts.map((chart) => chart.id) ?? []));
    });
    return () => {
      active = false;
    };
  }, [sourceRef, selectionDependency, maxCharts, useColumns]);

  React.useEffect(() => {
    if (!titleEdited && analysis?.plan) setTitle(`${analysis.sourceName} dashboard`);
  }, [analysis, titleEdited]);

  const plan = analysis?.plan ?? null;
  // The user's own charts sit after the suggestions, and are always included.
  const charts = React.useMemo(() => [...(plan?.charts ?? []), ...extra], [plan, extra]);
  const selected = charts.filter((chart) => chosen.has(chart.id));
  const profiles = plan?.profiles ?? [];
  const groupable = profiles
    .filter((profile) => profile.kind === "category" || profile.kind === "date" || profile.kind === "text")
    .map((profile) => profile.header);
  const measurable = profiles
    .filter((profile) => profile.kind === "number" || profile.kind === "currency" || profile.kind === "percent")
    .map((profile) => profile.header);

  const addChart = () => {
    const chart = customChart(
      { kind: newKind, dimension: newDimension || null, measure: newMeasure || null },
      profiles
    );
    if (!chart) return;
    if (charts.some((existing) => existing.id === chart.id)) return;
    setExtra((current) => [...current, chart]);
    setChosen((current) => new Set([...current, chart.id]));
  };

  const toggle = (id: string) =>
    setChosen((current) => {
      const next = new Set(current);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });

  const sourceLabel =
    sourceKey === SELECTION_KEY
      ? "The sheet I'm on"
      : (sources.find((item) => `${item.kind}:${item.name}` === sourceKey)?.label ?? sourceKey.split(":")[1]);

  return (
    <div>
      {saved.length > 0 ? (
        <section className={shared.card}>
          <Text weight="semibold">Your dashboards</Text>
          {saved.map((config) => (
            <div key={config.name} className={shared.column}>
              <div className={styles.saved}>
                <span className={styles.kindIcon}>
                  <Board24Regular />
                </span>
                <span className={styles.savedText}>
                  <Text weight="semibold" truncate wrap={false}>
                    {config.name}
                  </Text>
                  <Text className={shared.hint}>
                    {config.charts.length} charts · from &quot;{config.source.name || "a selection"}&quot;
                  </Text>
                </span>
              </div>
              <ActionButton
                runner={runner}
                id={`refresh-${config.name}`}
                appearance="secondary"
                icon={<ArrowSync20Regular />}
                label="Refresh with the latest data"
                busyLabel="Rebuilding..."
                onRun={() => refreshDashboard(config)}
              />
            </div>
          ))}
        </section>
      ) : null}

      <Step number={1} title="Which data?" description="Messy exports work too: titles, totals and text numbers are handled.">
        <Dropdown
          value={sourceLabel}
          selectedOptions={[sourceKey]}
          onOptionSelect={(_event, data) => setSourceKey(String(data.optionValue))}
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
        {loading ? <Spinner size="tiny" label="Reading your columns..." /> : null}
        {!loading && !plan ? (
          <Text className={shared.hint}>No table found there. It needs a row of headings with data underneath.</Text>
        ) : null}
        {plan ? (
          <Text className={shared.hint}>
            {plan.rowCount.toLocaleString()} rows, {plan.headers.length} columns.
          </Text>
        ) : null}
        {analysis && analysis.problemsFixed > 0 ? (
          <div className={shared.column}>
            <Tip>
              {analysis.problemsFixed.toLocaleString()} data problems will be smoothed over for the charts. Your sheet
              stays as it is.
            </Tip>
            <Button size="small" appearance="transparent" onClick={() => navigate("tidy", { source: sourceKey })}>
              Keep a clean copy too
            </Button>
          </div>
        ) : null}
      </Step>

      <Step
        number={2}
        title={`Pick your charts (up to ${MAX_CHARTS})`}
        description="Suggested from what your columns hold. Untick any you don't need."
        waitingFor={plan ? undefined : "Choose some data first."}
      >
        {plan && plan.charts.length === 0 ? (
          <Text className={shared.hint}>
            No charts to suggest: the table needs at least one column to group by, such as a region, product or date.
          </Text>
        ) : null}
        <div className={styles.list}>
          {charts.map((chart) => {
            const on = chosen.has(chart.id);
            return (
              <button
                key={chart.id}
                type="button"
                role="checkbox"
                aria-checked={on}
                className={mergeClasses(styles.suggestion, !on && styles.off)}
                onClick={() => toggle(chart.id)}
              >
                <Checkbox checked={on} tabIndex={-1} aria-hidden="true" />
                <span className={styles.kindIcon}>{KIND_ICON[chart.kind]}</span>
                <span className={styles.text}>
                  <Text weight="semibold">{chart.title}</Text>
                  <Text className={shared.hint}>{chart.reason}</Text>
                </span>
              </button>
            );
          })}
        </div>
        {plan ? (
          <div className={shared.column}>
            <Field label="Columns to use" hint="Leave empty to use them all.">
              <Dropdown
                multiselect
                placeholder="Every column"
                selectedOptions={useColumns}
                value={useColumns.length === 0 ? "Every column" : useColumns.join(", ")}
                onOptionSelect={(_event, data) => setUseColumns(data.selectedOptions)}
              >
                {plan.headers.map((header) => (
                  <Option key={header} value={header}>
                    {header}
                  </Option>
                ))}
              </Dropdown>
            </Field>
            <Field label="How many charts to suggest">
              <Input
                type="number"
                min={1}
                max={MAX_CHARTS}
                value={String(maxCharts)}
                onChange={(_event, data) => {
                  const value = Number(data.value);
                  if (value >= 1 && value <= MAX_CHARTS) setMaxCharts(value);
                }}
              />
            </Field>

            <MoreOptions label="Build one myself">
              <Field label="Chart">
                <Dropdown
                  value={KIND_LABEL[newKind]}
                  selectedOptions={[newKind]}
                  onOptionSelect={(_event, data) => setNewKind(data.optionValue as DashChartKind)}
                >
                  {(Object.keys(KIND_LABEL) as DashChartKind[]).map((kind) => (
                    <Option key={kind} value={kind} text={KIND_LABEL[kind]}>
                      {KIND_LABEL[kind]}
                    </Option>
                  ))}
                </Dropdown>
              </Field>
              <Field label="Of" hint="Leave empty to count rows instead.">
                <Dropdown
                  placeholder="Count of rows"
                  value={newMeasure}
                  selectedOptions={newMeasure ? [newMeasure] : []}
                  onOptionSelect={(_event, data) => setNewMeasure(String(data.optionValue))}
                >
                  <Option value="">Count of rows</Option>
                  {measurable.map((header) => (
                    <Option key={header} value={header}>
                      {header}
                    </Option>
                  ))}
                </Dropdown>
              </Field>
              <Field label="For each">
                <Dropdown
                  placeholder="Choose a column"
                  value={newDimension}
                  selectedOptions={newDimension ? [newDimension] : []}
                  onOptionSelect={(_event, data) => setNewDimension(String(data.optionValue))}
                >
                  {groupable.map((header) => (
                    <Option key={header} value={header}>
                      {header}
                    </Option>
                  ))}
                </Dropdown>
              </Field>
              <Button
                appearance="secondary"
                icon={<Add20Regular />}
                disabled={!newDimension && newKind !== "scatter"}
                onClick={addChart}
              >
                Add this chart
              </Button>
            </MoreOptions>
          </div>
        ) : null}

        {plan && plan.kpis.length > 0 ? (
          <>
            <Switch
              label="Headline numbers across the top"
              checked={includeKpis}
              onChange={(_event, data) => setIncludeKpis(data.checked)}
            />
            {includeKpis ? (
              <div className={styles.kpis}>
                {plan.kpis.map((kpi) => (
                  <span key={kpi.id} className={styles.kpiChip}>
                    {kpi.label}
                  </span>
                ))}
              </div>
            ) : null}
          </>
        ) : null}
      </Step>

      <Step number={3} title="Build it" waitingFor={plan ? undefined : "Choose some data first."}>
        <Field label="Dashboard name">
          <Input
            value={title}
            onChange={(_event, data) => {
              setTitle(data.value);
              setTitleEdited(true);
            }}
          />
        </Field>
        <ActionButton
          runner={runner}
          id="build"
          wide
          icon={<Board24Regular />}
          label={`Build dashboard with ${selected.length} ${selected.length === 1 ? "chart" : "charts"}`}
          busyLabel="Building..."
          disabled={selected.length === 0 && !includeKpis}
          onRun={() => buildDashboard(sourceRef, selected, { title, includeKpis })}
        />
        <Tip>
          It goes on a new sheet, laid out on a tidy grid. It&apos;s saved in this workbook: when new data arrives,
          click Refresh at the top of this page.
        </Tip>
      </Step>
    </div>
  );
};

export default DashboardPanel;
