import * as React from "react";
import { Field, Input, makeStyles, mergeClasses, Text, tokens } from "@fluentui/react-components";
import { ArrowDown16Regular, CopySelect24Regular, Search20Regular, Wand24Regular } from "@fluentui/react-icons";

import {
  applyTemplate,
  fillFormulaAcrossSelection,
  FillPreview,
  findTemplate,
  FORMULA_TEMPLATES,
  previewFill,
  previewTemplate,
} from "../features/formulaPatterns";
import { ActionButton, ActionCard, Tip, useActionRunner, useSharedStyles } from "./ui";
import { useSelectionVersion } from "./useSelection";

type CardId = "fill" | "library";

const CATEGORIES = ["Aggregate", "Lookup", "Ranking", "Text & dates"] as const;

const CATEGORY_LABELS: Record<(typeof CATEGORIES)[number], string> = {
  Aggregate: "Totals and percentages",
  Lookup: "Look up a value",
  Ranking: "Rank and count",
  "Text & dates": "Text and dates",
};

const useStyles = makeStyles({
  steps: {
    margin: "0",
    paddingLeft: "20px",
    display: "flex",
    flexDirection: "column",
    rowGap: "4px",
    fontSize: tokens.fontSizeBase200,
  },
  arrow: {
    display: "flex",
    justifyContent: "center",
    color: tokens.colorNeutralForeground3,
  },
  list: {
    display: "flex",
    flexDirection: "column",
    maxHeight: "260px",
    overflowY: "auto",
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
  },
  groupLabel: {
    padding: "8px 10px 4px",
    fontSize: tokens.fontSizeBase100,
    fontWeight: tokens.fontWeightSemibold,
    textTransform: "uppercase",
    letterSpacing: "0.04em",
    color: tokens.colorNeutralForeground3,
  },
  item: {
    display: "flex",
    flexDirection: "column",
    rowGap: "2px",
    padding: "8px 10px",
    border: "none",
    borderLeft: "3px solid transparent",
    background: "none",
    textAlign: "left",
    cursor: "pointer",
    color: tokens.colorNeutralForeground1,
    fontFamily: "inherit",
    ":hover": {
      backgroundColor: tokens.colorNeutralBackground1Hover,
    },
  },
  itemSelected: {
    backgroundColor: tokens.colorBrandBackground2,
    borderLeft: `3px solid ${tokens.colorBrandStroke1}`,
    ":hover": {
      backgroundColor: tokens.colorBrandBackground2Hover,
    },
  },
  itemName: {
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase300,
  },
});

const FormulaPanel: React.FC = () => {
  const shared = useSharedStyles();
  const styles = useStyles();
  const runner = useActionRunner();
  const selectionVersion = useSelectionVersion();

  const [openCard, setOpenCard] = React.useState<CardId | null>("fill");
  const [fill, setFill] = React.useState<FillPreview | null>(null);

  const [search, setSearch] = React.useState("");
  const [templateId, setTemplateId] = React.useState(FORMULA_TEMPLATES[0].id);
  const [inputs, setInputs] = React.useState<Record<string, string>>({});
  const [templatePreview, setTemplatePreview] = React.useState("");

  const template = findTemplate(templateId);
  const toggle = (id: CardId) => setOpenCard((current) => (current === id ? null : id));

  // The fill preview follows the selection, so there is nothing to click to "check" it.
  React.useEffect(() => {
    if (openCard !== "fill") {
      return undefined;
    }
    let active = true;
    void previewFill().then((next) => {
      if (active) {
        setFill(next);
      }
    });
    return () => {
      active = false;
    };
  }, [openCard, selectionVersion, runner.result]);

  React.useEffect(() => {
    if (openCard !== "library") {
      return undefined;
    }
    let active = true;
    void previewTemplate(templateId, inputs).then((formula) => {
      if (active) {
        setTemplatePreview(formula);
      }
    });
    return () => {
      active = false;
    };
  }, [openCard, templateId, inputs, selectionVersion, runner.result]);

  const query = search.trim().toLowerCase();
  const matching = FORMULA_TEMPLATES.filter(
    (item) =>
      query === "" ||
      item.name.toLowerCase().includes(query) ||
      item.description.toLowerCase().includes(query)
  );

  return (
    <div>
      <ActionCard
        icon={<CopySelect24Regular />}
        title="Copy a formula down"
        description="Like dragging the fill handle, for any size of range."
        open={openCard === "fill"}
        onToggle={() => toggle("fill")}
      >
        <ol className={styles.steps}>
          <li>Type your formula in the top cell.</li>
          <li>Select from that cell down (or across) to where it should stop.</li>
          <li>Check the preview below, then click Fill.</li>
        </ol>

        {fill?.available ? (
          <>
            <Field label="Top cell">
              <div className={shared.code}>{fill.sourceFormula}</div>
            </Field>
            <div className={styles.arrow}>
              <ArrowDown16Regular />
            </div>
            <Field label="Last cell will be">
              <div className={shared.code}>{fill.lastCellFormula}</div>
            </Field>
          </>
        ) : (
          <Tip>{fill?.message ?? "Select your formula cell and the cells to fill."}</Tip>
        )}

        <ActionButton
          runner={runner}
          id="fill"
          wide
          label="Fill"
          busyLabel="Filling..."
          disabled={!fill?.available}
          onRun={fillFormulaAcrossSelection}
        />
      </ActionCard>

      <ActionCard
        icon={<Wand24Regular />}
        title="Insert a ready-made formula"
        description="Running totals, lookups, rankings and more, without typing them."
        open={openCard === "library"}
        onToggle={() => toggle("library")}
      >
        <Input
          contentBefore={<Search20Regular />}
          placeholder="Search, e.g. total, lookup, rank"
          value={search}
          onChange={(_event, data) => setSearch(data.value)}
        />

        <div className={styles.list} role="listbox" aria-label="Formulas">
          {matching.length === 0 ? (
            <Text className={shared.hint} style={{ padding: "10px" }}>
              Nothing matches &quot;{search}&quot;.
            </Text>
          ) : (
            CATEGORIES.map((category) => {
              const inCategory = matching.filter((item) => item.category === category);
              if (inCategory.length === 0) {
                return null;
              }
              return (
                <React.Fragment key={category}>
                  <div className={styles.groupLabel}>{CATEGORY_LABELS[category]}</div>
                  {inCategory.map((item) => (
                    <button
                      key={item.id}
                      type="button"
                      role="option"
                      aria-selected={item.id === templateId}
                      className={mergeClasses(
                        styles.item,
                        item.id === templateId && styles.itemSelected
                      )}
                      onClick={() => {
                        setTemplateId(item.id);
                        setInputs({});
                      }}
                    >
                      <span className={styles.itemName}>{item.name}</span>
                      <span className={shared.hint}>{item.description}</span>
                    </button>
                  ))}
                </React.Fragment>
              );
            })
          )}
        </div>

        {template ? (
          <>
            <Tip>Select the cells where the results should go, then fill in anything below.</Tip>
            {template.inputs.map((input) => (
              <Field key={input.key} label={input.label} hint={input.help}>
                <Input
                  placeholder={input.placeholder}
                  value={inputs[input.key] ?? ""}
                  onChange={(_event, data) =>
                    setInputs((current) => ({ ...current, [input.key]: data.value }))
                  }
                />
              </Field>
            ))}
            {templatePreview ? (
              <Field label="The first cell will get">
                <div className={shared.code}>{templatePreview}</div>
              </Field>
            ) : null}
            <ActionButton
              runner={runner}
              id="library"
              wide
              label={`Insert "${template.name}"`}
              busyLabel="Inserting..."
              onRun={() => applyTemplate(templateId, inputs)}
            />
          </>
        ) : null}
      </ActionCard>
    </div>
  );
};

export default FormulaPanel;
