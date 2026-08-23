/**
 * Small shared building blocks for the task pane: a section wrapper, the
 * result banner every command reports through, and the hook that runs a command
 * while keeping the button in a sensible busy/disabled state.
 */

import * as React from "react";
import {
  Button,
  makeStyles,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  Spinner,
  Text,
  tokens,
} from "@fluentui/react-components";

import { OperationResult } from "../shared/types";

export const useSharedStyles = makeStyles({
  section: {
    display: "flex",
    flexDirection: "column",
    rowGap: "10px",
    padding: "12px",
    marginBottom: "12px",
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground1,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  sectionTitle: {
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase400,
  },
  hint: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
  },
  row: {
    display: "flex",
    alignItems: "center",
    columnGap: "8px",
    flexWrap: "wrap",
  },
  column: {
    display: "flex",
    flexDirection: "column",
    rowGap: "8px",
  },
  code: {
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: tokens.fontSizeBase200,
    backgroundColor: tokens.colorNeutralBackground3,
    padding: "6px 8px",
    borderRadius: tokens.borderRadiusSmall,
    wordBreak: "break-all",
    whiteSpace: "pre-wrap",
  },
  detailList: {
    margin: "6px 0 0 0",
    paddingLeft: "18px",
    fontSize: tokens.fontSizeBase200,
  },
  scrollList: {
    maxHeight: "180px",
    overflowY: "auto",
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusSmall,
    padding: "6px",
  },
});

export interface SectionProps {
  title: string;
  description?: string;
  children: React.ReactNode;
}

export const Section: React.FC<SectionProps> = ({ title, description, children }) => {
  const styles = useSharedStyles();
  return (
    <section className={styles.section}>
      <Text className={styles.sectionTitle}>{title}</Text>
      {description ? <Text className={styles.hint}>{description}</Text> : null}
      {children}
    </section>
  );
};

/** Render an OperationResult as a success or error banner with optional notes. */
export const ResultBanner: React.FC<{ result: OperationResult | null }> = ({ result }) => {
  const styles = useSharedStyles();
  if (!result) {
    return null;
  }
  return (
    <MessageBar intent={result.ok ? "success" : "error"}>
      <MessageBarBody>
        <MessageBarTitle>{result.ok ? "Done" : "Could not finish"}</MessageBarTitle>
        {result.message}
        {result.details && result.details.length > 0 ? (
          <ul className={styles.detailList}>
            {result.details.map((detail, index) => (
              <li key={`${index}-${detail}`}>{detail}</li>
            ))}
          </ul>
        ) : null}
      </MessageBarBody>
    </MessageBar>
  );
};

export interface ActionRunner {
  busy: boolean;
  result: OperationResult | null;
  run: (action: () => Promise<OperationResult>) => Promise<void>;
  clear: () => void;
}

/**
 * Run one command at a time, capturing whatever it returns - including anything
 * it throws, so a failure never leaves the pane stuck on a spinner.
 */
export function useActionRunner(): ActionRunner {
  const [busy, setBusy] = React.useState(false);
  const [result, setResult] = React.useState<OperationResult | null>(null);

  const run = React.useCallback(async (action: () => Promise<OperationResult>) => {
    setBusy(true);
    setResult(null);
    try {
      setResult(await action());
    } catch (error) {
      setResult({
        ok: false,
        message: error instanceof Error ? error.message : String(error),
      });
    } finally {
      setBusy(false);
    }
  }, []);

  const clear = React.useCallback(() => setResult(null), []);

  return { busy, result, run, clear };
}

export interface RunButtonProps {
  label: string;
  busy: boolean;
  disabled?: boolean;
  appearance?: "primary" | "secondary";
  onClick: () => void;
}

export const RunButton: React.FC<RunButtonProps> = ({
  label,
  busy,
  disabled,
  appearance = "primary",
  onClick,
}) => (
  <Button
    appearance={appearance}
    disabled={busy || disabled}
    onClick={onClick}
    icon={busy ? <Spinner size="tiny" /> : undefined}
  >
    {label}
  </Button>
);
