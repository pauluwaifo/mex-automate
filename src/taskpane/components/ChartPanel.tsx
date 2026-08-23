import * as React from "react";
import {
  Checkbox,
  Dropdown,
  Field,
  Input,
  Option,
  OptionGroup,
  Radio,
  RadioGroup,
  Text,
} from "@fluentui/react-components";

import {
  Aggregation,
  AGGREGATIONS,
  CHART_TYPES,
  ChartRequest,
  createChart,
  DataSourceRef,
  findChartType,
  LegendPosition,
  listSourceOptions,
  readSourceHeaders,
  SortOrder,
  SourceOption,
  summarizeToNewSheet,
  SummarizeRequest,
  validateChartRequest,
} from "../features/charts";
import { ResultBanner, RunButton, Section, useActionRunner, useSharedStyles } from "./ui";

const CHART_GROUPS = ["Column & bar", "Line & area", "Pie", "Scatter"] as const;

const SORT_ORDERS: Array<{ value: SortOrder; label: string }> = [
  { value: "valueDesc", label: "Largest first" },
  { value: "valueAsc", label: "Smallest first" },
  { value: "categoryAsc", label: "Category A-Z" },
  { value: "categoryDesc", label: "Category Z-A" },
  { value: "none", label: "Source order" },
];

const LEGEND_POSITIONS: Array<{ value: LegendPosition; label: string }> = [
  { value: "right", label: "Right" },
  { value: "bottom", label: "Bottom" },
  { value: "top", label: "Top" },
  { value: "left", label: "Left" },
  { value: "none", label: "No legend" },
];

const ChartPanel: React.FC = () => {
  const styles = useSharedStyles();
  const runner = useActionRunner();

  const [sources, setSources] = React.useState<SourceOption[]>([]);
  const [sourceKey, setSourceKey] = React.useState("selection");
  const [headers, setHeaders] = React.useState<string[]>([]);

  const [summarize, setSummarize] = React.useState(true);
  const [groupBy, setGroupBy] = React.useState("");
  const [valueColumns, setValueColumns] = React.useState<string[]>([]);
  const [aggregation, setAggregation] = React.useState<Aggregation>("sum");
  const [sort, setSort] = React.useState<SortOrder>("valueDesc");
  const [limitCategories, setLimitCategories] = React.useState(false);
  const [topN, setTopN] = React.useState("10");

  const [chartTypeId, setChartTypeId] = React.useState("columnClustered");
  const [title, setTitle] = React.useState("");
  const [legendPosition, setLegendPosition] = React.useState<LegendPosition>("right");
  const [showDataLabels, setShowDataLabels] = React.useState(false);
  const [destination, setDestination] = React.useState<"newSheet" | "sourceSheet">("newSheet");
  const [destinationSheetName, setDestinationSheetName] = React.useState("Chart");

  const chartType = findChartType(chartTypeId);

  const sourceRef: DataSourceRef = React.useMemo(() => {
    if (sourceKey === "selection") {
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

  // Re-read headers whenever the chosen source changes, and reset any picks
  // that no longer exist in it.
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
        // Default to the first column that is not the category column.
        return next.length > 1 ? [next[1]] : [];
      });
    });
    return () => {
      active = false;
    };
  }, [sourceRef, runner.result]);

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

  const request: ChartRequest = React.useMemo(
    () => ({
      source: sourceRef,
      summarize: summarizeRequest,
      chartTypeId,
      title,
      legendPosition,
      showDataLabels,
      destination,
      destinationSheetName,
    }),
    [
      sourceRef,
      summarizeRequest,
      chartTypeId,
      title,
      legendPosition,
      showDataLabels,
      destination,
      destinationSheetName,
    ]
  );

  // Show the same problems the feature would report, before anything is written.
  const problems = headers.length > 0 ? validateChartRequest(request, headers.length) : [];
  const ready = headers.length > 0 && problems.length === 0;

  const valueColumnOptions = headers.filter((header) => !summarize || header !== groupBy);

  return (
    <div>
      <ResultBanner result={runner.result} />

      <Section
        title="Data"
        description="Chart the current selection, or any sheet or table in this workbook. The first row is treated as headers."
      >
        <Dropdown
          value={
            sourceKey === "selection"
              ? "Current selection"
              : (sources.find((item) => `${item.kind}:${item.name}` === sourceKey)?.label ?? "")
          }
          selectedOptions={[sourceKey]}
          onOptionSelect={(_event, data) => setSourceKey(String(data.optionValue))}
        >
          <Option value="selection">Current selection</Option>
          {sources.map((item) => (
            <Option key={`${item.kind}:${item.name}`} value={`${item.kind}:${item.name}`} text={item.label}>
              {item.label}
            </Option>
          ))}
        </Dropdown>
        <Text className={styles.hint}>
          {headers.length > 0
            ? `${headers.length} column(s): ${headers.slice(0, 5).join(", ")}${headers.length > 5 ? "..." : ""}`
            : "No data found. Select a range with a header row, or pick a sheet."}
        </Text>
      </Section>

      <Section
        title="Summarize"
        description="Group the rows before charting. Leave this on for raw data such as order lines; turn it off if your range is already a small summary table."
      >
        <Checkbox
          label="Group and total the data first"
          checked={summarize}
          onChange={(_event, data) => setSummarize(Boolean(data.checked))}
        />

        {summarize ? (
          <>
            <Field label="Group by (categories)">
              <Dropdown
                value={groupBy}
                selectedOptions={[groupBy]}
                disabled={headers.length === 0}
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

            <Field label="Calculate">
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
              label="Of these columns"
              hint={chartType?.singleSeriesOnly ? "Pie and doughnut charts show one column." : undefined}
            >
              <Dropdown
                multiselect={!chartType?.singleSeriesOnly}
                placeholder="Choose a column"
                selectedOptions={valueColumns}
                value={valueColumns.join(", ")}
                disabled={headers.length === 0}
                onOptionSelect={(_event, data) =>
                  setValueColumns(
                    chartType?.singleSeriesOnly ? [String(data.optionValue)] : data.selectedOptions
                  )
                }
              >
                {valueColumnOptions.map((header) => (
                  <Option key={header} value={header}>
                    {header}
                  </Option>
                ))}
              </Dropdown>
            </Field>

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

            <Checkbox
              label="Keep only the top categories"
              checked={limitCategories}
              onChange={(_event, data) => setLimitCategories(Boolean(data.checked))}
            />
            {limitCategories ? (
              <Field label="How many" hint='Everything else is combined into one "Other" row.'>
                <Input type="number" min={1} value={topN} onChange={(_event, data) => setTopN(data.value)} />
              </Field>
            ) : null}

            <RunButton
              label="Build summary table only"
              appearance="secondary"
              busy={runner.busy}
              disabled={!ready || !summarizeRequest}
              onClick={() =>
                void runner.run(() =>
                  summarizeToNewSheet(sourceRef, summarizeRequest!, "Summary")
                )
              }
            />
          </>
        ) : null}
      </Section>

      <Section title="Chart">
        <Field label="Type">
          <Dropdown
            value={chartType?.label ?? ""}
            selectedOptions={[chartTypeId]}
            onOptionSelect={(_event, data) => {
              const nextId = String(data.optionValue);
              setChartTypeId(nextId);
              // A pie cannot show several series, so drop back to one column.
              if (findChartType(nextId)?.singleSeriesOnly) {
                setValueColumns((current) => current.slice(0, 1));
              }
            }}
          >
            {CHART_GROUPS.map((group) => {
              const inGroup = CHART_TYPES.filter((type) => type.group === group);
              return inGroup.length === 0 ? null : (
                <OptionGroup key={group} label={group}>
                  {inGroup.map((type) => (
                    <Option key={type.id} value={type.id} text={type.label}>
                      {type.label}
                    </Option>
                  ))}
                </OptionGroup>
              );
            })}
          </Dropdown>
        </Field>

        <Field label="Title" hint="Leave blank for no title.">
          <Input value={title} onChange={(_event, data) => setTitle(data.value)} />
        </Field>

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

        <Checkbox
          label="Show values on the chart"
          checked={showDataLabels}
          onChange={(_event, data) => setShowDataLabels(Boolean(data.checked))}
        />

        <Field label="Put it">
          <RadioGroup
            value={destination}
            onChange={(_event, data) => setDestination(data.value as "newSheet" | "sourceSheet")}
          >
            <Radio value="newSheet" label="On a new sheet" />
            <Radio value="sourceSheet" label="Next to the data" />
          </RadioGroup>
        </Field>

        {destination === "newSheet" ? (
          <Field label="New sheet name">
            <Input
              value={destinationSheetName}
              onChange={(_event, data) => setDestinationSheetName(data.value)}
            />
          </Field>
        ) : null}

        {problems.length > 0 ? (
          <ul className={styles.detailList}>
            {problems.map((problem) => (
              <li key={problem}>{problem}</li>
            ))}
          </ul>
        ) : null}

        <RunButton
          label="Create chart"
          busy={runner.busy}
          disabled={!ready}
          onClick={() => void runner.run(() => createChart(request))}
        />
        {summarize ? (
          <Text className={styles.hint}>
            The summary table is written to the sheet as well, so the numbers behind the chart stay
            visible.
          </Text>
        ) : null}
      </Section>
    </div>
  );
};

export default ChartPanel;
