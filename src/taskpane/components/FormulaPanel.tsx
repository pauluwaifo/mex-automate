import * as React from "react";
import { Button, Dropdown, Field, Input, Option, OptionGroup, Text } from "@fluentui/react-components";

import {
  applyTemplate,
  fillFormulaAcrossSelection,
  FillPreview,
  findTemplate,
  FORMULA_TEMPLATES,
  previewFill,
  previewTemplate,
} from "../features/formulaPatterns";
import { ResultBanner, RunButton, Section, useActionRunner, useSharedStyles } from "./ui";

const CATEGORIES = ["Aggregate", "Lookup", "Ranking", "Text & dates"] as const;

const FormulaPanel: React.FC = () => {
  const styles = useSharedStyles();
  const runner = useActionRunner();

  const [fill, setFill] = React.useState<FillPreview | null>(null);
  const [templateId, setTemplateId] = React.useState(FORMULA_TEMPLATES[0].id);
  const [inputs, setInputs] = React.useState<Record<string, string>>({});
  const [templatePreview, setTemplatePreview] = React.useState("");

  const template = findTemplate(templateId);

  const refreshFillPreview = React.useCallback(() => {
    void previewFill().then(setFill);
  }, []);

  React.useEffect(() => {
    refreshFillPreview();
  }, [refreshFillPreview, runner.result]);

  // Re-render the template preview whenever the template or its inputs change.
  React.useEffect(() => {
    let active = true;
    void previewTemplate(templateId, inputs).then((formula) => {
      if (active) {
        setTemplatePreview(formula);
      }
    });
    return () => {
      active = false;
    };
  }, [templateId, inputs, runner.result]);

  const setInput = (key: string, value: string) =>
    setInputs((current) => ({ ...current, [key]: value }));

  return (
    <div>
      <ResultBanner result={runner.result} />

      <Section
        title="Fill a formula across a range"
        description="Select the cell holding your formula together with the cells to fill. References shift the way they would if you dragged the fill handle."
      >
        <Button appearance="secondary" onClick={refreshFillPreview} disabled={runner.busy}>
          Check my selection
        </Button>
        {fill ? (
          <>
            <Text className={styles.hint}>{fill.message}</Text>
            {fill.available ? (
              <>
                <Field label="First cell">
                  <div className={styles.code}>{fill.sourceFormula}</div>
                </Field>
                <Field label="Last cell will become">
                  <div className={styles.code}>{fill.lastCellFormula}</div>
                </Field>
              </>
            ) : null}
          </>
        ) : null}
        <RunButton
          label="Fill formula"
          busy={runner.busy}
          disabled={!fill?.available}
          onClick={() => void runner.run(fillFormulaAcrossSelection)}
        />
      </Section>

      <Section
        title="Formula library"
        description="Select the column of cells the formula should fill, pick a pattern, then insert."
      >
        <Field label="Pattern">
          <Dropdown
            selectedOptions={[templateId]}
            value={template?.name ?? ""}
            onOptionSelect={(_event, data) => {
              setTemplateId(String(data.optionValue));
              setInputs({});
            }}
          >
            {CATEGORIES.map((category) => {
              const inCategory = FORMULA_TEMPLATES.filter((item) => item.category === category);
              return inCategory.length === 0 ? null : (
                <OptionGroup key={category} label={category}>
                  {inCategory.map((item) => (
                    <Option key={item.id} value={item.id} text={item.name}>
                      {item.name}
                    </Option>
                  ))}
                </OptionGroup>
              );
            })}
          </Dropdown>
        </Field>

        {template ? <Text className={styles.hint}>{template.description}</Text> : null}

        {template?.inputs.map((input) => (
          <Field key={input.key} label={input.label} hint={input.help}>
            <Input
              placeholder={input.placeholder}
              value={inputs[input.key] ?? ""}
              onChange={(_event, data) => setInput(input.key, data.value)}
            />
          </Field>
        ))}

        {templatePreview ? (
          <Field label="Preview (first cell)">
            <div className={styles.code}>{templatePreview}</div>
          </Field>
        ) : null}

        <RunButton
          label="Insert formula"
          busy={runner.busy}
          onClick={() => void runner.run(() => applyTemplate(templateId, inputs))}
        />
      </Section>
    </div>
  );
};

export default FormulaPanel;
