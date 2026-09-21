import * as React from "react";
import {
  Dropdown,
  Field,
  Input,
  Option,
  Radio,
  RadioGroup,
  Switch,
  Text,
} from "@fluentui/react-components";
import {
  DataArea24Regular,
  DataBarHorizontal24Regular,
  DataBarVertical24Regular,
  DataLine24Regular,
  DataPie24Regular,
  DataScatter24Regular,
  TableSimple20Regular,
} from "@fluentui/react-icons";

import {
  Aggregation,
  AGGREGATIONS,
  aggregationLabel,
  ChartRequest,
  createChart,
  DataSourceRef,
  LegendPosition,
  listSourceOptions,
  readSourceHeaders,
  SortOrder,
  SourceOption,
  summarizeToNewSheet,
  SummarizeRequest,
  validateChartRequest,
} from "../features/charts";
import {
  ActionButton,
  Choice,
  ChoiceGrid,
  MoreOptions,
  Step,
  Tip,
  useActionRunner,
  useSharedStyles,
} from "./ui";
import { useSelectionVersion } from "./useSelection";

/** The chart shapes offered up front; stacking is a separate switch. */
type Shape = "column" | "bar" | "line" | "area" | "pie" | "doughnut" | "scatter";

const SHAPES: Array<Choice<Shape>> = [
  { value: "column", label: "Column", icon: <DataBarVertical24Regular /> },
  { value: "bar", label: "Bar", icon: <DataBarHorizontal24Regular /> },
  { value: "line", label: "Line", icon: <DataLine24Regular /> },
  { value: "pie", label: "Pie", icon: <DataPie24Regular /> },
  { value: "doughnut", label: "Doughnut", icon: <DataPie24Regular /> },
  { value: "area", label: "Area", icon: <DataArea24Regular /> },
  { value: "scatter", label: "Scatter", icon: <DataScatter24Regular /> },
];

const STACKABLE: ReadonlySet<Shape> = new Set(["column", "bar", "area"]);
const ONE_SERIES: ReadonlySet<Shape> = new Set(["pie", "doughnut"]);

function chartTypeIdFor(shape: Shape, stacked: boolean): string {
  switch (shape) {
    case "column":
      return stacked ? "columnStacked" : "columnClustered";
    case "bar":
      return stacked ? "barStacked" : "barClustered";
    case "area":
      return stacked ? "areaStacked" : "area";
    case "line":
      return "lineMarkers";
    case "scatter":
      return "xyscatter";
    default:
      return shape;
  }
}

const SORT_ORDERS: Array<{ value: SortOrder; label: string }> = [
  { value: "valueDesc", label: "Biggest first" },
  { value: "valueAsc", label: "Smallest first" },
  { value: "categoryAsc", label: "A to Z" },
  { value: "categoryDesc", label: "Z to A" },
  { value: "none", label: "As they appear in the data" },
];

const LEGEND_POSITIONS: Array<{ value: LegendPosition; label: string }> = [
  { value: "right", label: "Right" },
  { value: "bottom", label: "Bottom" },
  { value: "top", label: "Top" },
  { value: "left", label: "Left" },
  { value: "none", label: "Hide the legend" },
];

const SELECTION_KEY = "selection";

const ChartPanel: React.FC = () => {
  const styles = useSharedStyles();
  const runner = useActionRunner();
  const selectionVersion = useSelectionVersion();

  const [sources, setSources] = React.useState<SourceOption[]>([]);
  const [sourceKey, setSourceKey] = React.useState(SELECTION_KEY);
  const [headers, setHeaders] = React.useState<string[]>([]);

  const [summarize, setSummarize] = React.useState(true);
  const [groupBy, setGroupBy] = React.useState("");
  const [valueColumns, setValueColumns] = React.useState<string[]>([]);
  const [aggregation, setAggregation] = React.useState<Aggregation>("sum");
  const [sort, setSort] = React.useState<SortOrder>("valueDesc");
  const [limitCategories, setLimitCategories] = React.useState(false);
  const [topN, setTopN] = React.useState("8");

  const [shape, setShape] = React.useState<Shape>("column");
  const [stacked, setStacked] = React.useState(false);
  const [title, setTitle] = React.useState("");
  const [titleEdited, setTitleEdited] = React.useState(false);
  const [showDataLabels, setShowDataLabels] = React.useState(false);
  const [legendPosition, setLegendPosition] = React.useState<LegendPosition>("right");
  const [destination, setDestination] = React.useState<"newSheet" | "sourceSheet">("newSheet");
  const [destinationSheetName, setDestinationSheetName] = React.useState("Chart");

  const sourceRef: DataSourceRef = React.useMemo(() => {
    if (sourceKey === SELECTION_KEY) {
      return { kind: "selection", name: "" };
    }
    const separator = sourceKey.indexOf(":");
    return {
      kind: sourceKey.slice(0, separator) as "sheet" | "table",
      name: sourceKey.slice(separator + 1),
    };
  }, [sourceKey]);

  React.useEffect(() => {
    void listSourceOptions().then(setSources);
  }, [runner.result]);

  // Re-read the columns when the source changes - and, for "selected cells",
  // whenever the user clicks somewhere else in Excel.
  const selectionDependency = sourceKey === SELECTION_KEY ? selectionVersion : 0;
  React.useEffect(() => {
    let active = true;
    void readSourceHeaders(sourceRef).then((next) => {
      if (!active) {
        return;
      }
      setHeaders(next);
      setGroupBy((current) => (next.includes(current) ? current : (next[0] ?? "")));
      setValueColumns((current) => {
        const kept = current.filter((header) => next.includes(header));
        if (kept.length > 0) {
          return kept;
        }
        return next.length > 1 ? [next[1]] : [];
      });
    });
    return () => {
      active = false;
    };
  }, [sourceRef, selectionDependency, runner.result]);

  const singleSeries = ONE_SERIES.has(shape);

  // Offer a sensible title until the user types their own.
  const suggestedTitle =
    summarize && valueColumns.length > 0 && groupBy
      ? `${aggregationLabel(aggregation, valueColumns[0])} by ${groupBy}`
      : "";
  React.useEffect(() => {
    if (!titleEdited) {
      setTitle(suggestedTitle);
    }
  }, [suggestedTitle, titleEdited]);

  const summarizeRequest: SummarizeRequest | null = React.useMemo(() => {
    if (!summarize) {
      return null;
    }
    return {
      groupByColumn: headers.indexOf(groupBy),
      valueColumns: valueColumns.map((header) => headers.indexOf(header)).filter((i) => i >= 0),
      aggregation,
      sort,
      topN: limitCategories ? Number(topN) || null : null,
    };
  }, [summarize, headers, groupBy, valueColumns, aggregation, sort, limitCategories, topN]);

  const request: ChartRequest = {
    source: sourceRef,
    summarize: summarizeRequest,
    chartTypeId: chartTypeIdFor(shape, stacked && STACKABLE.has(shape)),
    title,
    legendPosition,
    showDataLabels,
    destination,
    destinationSheetName,
  };

  const problems = headers.length > 0 ? validateChartRequest(request, headers.length) : [];
  const hasData = headers.length > 0;
  const ready = hasData && problems.length === 0;

  const valueChoices = headers.filter((header) => !summarize || header !== groupBy);

  const pickShape = (next: Shape) => {
    setShape(next);
    // A pie can only show one set of numbers.
    if (ONE_SERIES.has(next)) {
      setValueColumns((current) => current.slice(0, 1));
    }
  };

  const sourceLabel =
    sourceKey === SELECTION_KEY
      ? "The cells I've selected"
      : (sources.find((item) => `${item.kind}:${item.name}` === sourceKey)?.label ?? "");

  return (
    <div>
      <Step number={1} title="Choose your data">
        <Dropdown
          value={sourceLabel}
          selectedOptions={[sourceKey]}
          onOptionSelect={(_event, data) => setSourceKey(String(data.optionValue))}
        >
          <Option value={SELECTION_KEY} text="The cells I've selected">
            The cells I&apos;ve selected
          </Option>
          {sources.map((item) => (
            <Option key={`${item.kind}:${item.name}`} value={`${item.kind}:${item.name}`} text={item.label}>
              {item.label}
            </Option>
          ))}
        </Dropdown>
        {hasData ? (
          <div className={styles.row}>
            <TableSimple20Regular />
            <Text className={styles.hint}>
              {headers.length} columns: {headers.slice(0, 4).join(", ")}
              {headers.length > 4 ? ", ..." : ""}
            </Text>
          </div>
        ) : (
          <Tip>
            {sourceKey === SELECTION_KEY
              ? "Click inside your table in Excel. The first row should be headings."
              : "That sheet looks empty. It needs a heading row and at least one row of data."}
          </Tip>
        )}
      </Step>

      <Step
        number={2}
        title="What should it show?"
        waitingFor={hasData ? undefined : "Choose some data first."}
      >
        <Switch
          label="Add up my data first"
          checked={summarize}
          onChange={(_event, data) => setSummarize(data.checked)}
        />
        <Text className={styles.hint}>
          {summarize
            ? "Best for a long list, like one row per sale. Rows are grouped and totalled before charting."
            : "Best when your data is already a short table of totals. The first column becomes the labels."}
        </Text>

        {summarize ? (
          <>
            <Field label="Calculate the">
              <Dropdown
                value={AGGREGATIONS.find((item) => item.value === aggregation)?.label ?? ""}
                selectedOptions={[aggregation]}
                onOptionSelect={(_event, data) => setAggregation(data.optionValue as Aggregation)}
              >
                {AGGREGATIONS.map((item) => (
                  <Option key={item.value} value={item.value}>
                    {item.label}
                  </Option>
                ))}
              </Dropdown>
            </Field>

            <Field
              label="of"
              hint={singleSeries ? "A pie chart shows one column." : "You can pick more than one."}
            >
              <Dropdown
                multiselect={!singleSeries}
                placeholder="Choose a column"
                selectedOptions={valueColumns}
                value={valueColumns.join(", ")}
                onOptionSelect={(_event, data) =>
                  setValueColumns(singleSeries ? [String(data.optionValue)] : data.selectedOptions)
                }
              >
                {valueChoices.map((header) => (
                  <Option key={header} value={header}>
                    {header}
                  </Option>
                ))}
              </Dropdown>
            </Field>

            <Field label="for each">
              <Dropdown
                value={groupBy}
                selectedOptions={[groupBy]}
                onOptionSelect={(_event, data) => {
                  const next = String(data.optionValue);
                  setGroupBy(next);
                  setValueColumns((current) => current.filter((header) => header !== next));
                }}
              >
                {headers.map((header) => (
                  <Option key={header} value={header}>
                    {header}
                  </Option>
                ))}
              </Dropdown>
            </Field>

            <MoreOptions>
              <Field label="Order">
                <Dropdown
                  value={SORT_ORDERS.find((item) => item.value === sort)?.label ?? ""}
                  selectedOptions={[sort]}
                  onOptionSelect={(_event, data) => setSort(data.optionValue as SortOrder)}
                >
                  {SORT_ORDERS.map((item) => (
                    <Option key={item.value} value={item.value}>
                      {item.label}
                    </Option>
                  ))}
                </Dropdown>
              </Field>
              <Switch
                label="Show only the biggest few, and group the rest as Other"
                checked={limitCategories}
                onChange={(_event, data) => setLimitCategories(data.checked)}
              />
              {limitCategories ? (
                <Field label="How many to show">
                  <Input
                    type="number"
                    min={1}
                    value={topN}
                    onChange={(_event, data) => setTopN(data.value)}
                  />
                </Field>
              ) : null}
            </MoreOptions>
          </>
        ) : null}
      </Step>

      <Step number={3} title="Pick a chart" waitingFor={hasData ? undefined : "Choose some data first."}>
        <ChoiceGrid label="Chart type" choices={SHAPES} value={shape} onChange={pickShape} columns={4} />
        {STACKABLE.has(shape) && valueColumns.length > 1 ? (
          <Switch
            label="Stack the columns on top of each other"
            checked={stacked}
            onChange={(_event, data) => setStacked(data.checked)}
          />
        ) : null}
      </Step>

      <Step number={4} title="Finish" waitingFor={hasData ? undefined : "Choose some data first."}>
        <Field label="Chart title">
          <Input
            value={title}
            placeholder="No title"
            onChange={(_event, data) => {
              setTitle(data.value);
              setTitleEdited(true);
            }}
          />
        </Field>

        <Switch
          label="Show the numbers on the chart"
          checked={showDataLabels}
          onChange={(_event, data) => setShowDataLabels(data.checked)}
        />

        <Field label="Put the chart">
          <RadioGroup
            value={destination}
            onChange={(_event, data) => setDestination(data.value as "newSheet" | "sourceSheet")}
          >
            <Radio value="newSheet" label="On a new sheet" />
            <Radio value="sourceSheet" label="Next to my data" />
          </RadioGroup>
        </Field>

        <MoreOptions>
          {destination === "newSheet" ? (
            <Field label="New sheet name">
              <Input
                value={destinationSheetName}
                onChange={(_event, data) => setDestinationSheetName(data.value)}
              />
            </Field>
          ) : null}
          <Field label="Legend">
            <Dropdown
              value={LEGEND_POSITIONS.find((item) => item.value === legendPosition)?.label ?? ""}
              selectedOptions={[legendPosition]}
              onOptionSelect={(_event, data) => setLegendPosition(data.optionValue as LegendPosition)}
            >
              {LEGEND_POSITIONS.map((item) => (
                <Option key={item.value} value={item.value}>
                  {item.label}
                </Option>
              ))}
            </Dropdown>
          </Field>
        </MoreOptions>

        {problems.length > 0 ? (
          <ul className={styles.detailList}>
            {problems.map((problem) => (
              <li key={problem}>{problem}</li>
            ))}
          </ul>
        ) : null}

        <ActionButton
          runner={runner}
          id="chart"
          wide
          label="Create chart"
          busyLabel="Creating chart..."
          disabled={!ready}
          onRun={() => createChart(request)}
        />
        {summarize && summarizeRequest ? (
          <ActionButton
            runner={runner}
            id="summary"
            appearance="secondary"
            label="Just make the summary table"
            disabled={!ready}
            onRun={() => summarizeToNewSheet(sourceRef, summarizeRequest, "Summary")}
          />
        ) : null}
        {summarize ? (
          <Tip>The totals behind the chart are put on the sheet too, so you can check them.</Tip>
        ) : null}
      </Step>
    </div>
  );
};

export default ChartPanel;
