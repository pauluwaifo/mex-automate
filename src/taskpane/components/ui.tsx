/**
 * Shared building blocks for the task pane.
 *
 * The pane is narrow (about 320px) and used by people who are not thinking about
 * add-ins, so these components exist to keep every tool looking and behaving the
 * same way:
 *
 * - `Step` numbers the stages of a multi-step tool so the order is obvious.
 * - `ActionCard` folds a tool away until it is wanted, so a page of four tools
 *   is four lines, not four screens.
 * - `MoreOptions` hides settings most people never change.
 * - `ActionButton` shows its own result directly underneath itself. A result
 *   banner at the top of the pane is invisible when the button that caused it
 *   is scrolled a screen further down.
 */

import * as React from "react";
import {
  Button,
  makeStyles,
  mergeClasses,
  MessageBar,
  MessageBarActions,
  MessageBarBody,
  MessageBarTitle,
  Spinner,
  Text,
  tokens,
} from "@fluentui/react-components";
import {
  ChevronDown20Regular,
  ChevronRight20Regular,
  Dismiss20Regular,
  Lightbulb16Regular,
  TableSimple20Regular,
} from "@fluentui/react-icons";

import { OperationResult } from "../shared/types";

export const useSharedStyles = makeStyles({
  card: {
    display: "flex",
    flexDirection: "column",
    rowGap: "12px",
    padding: "14px",
    marginBottom: "10px",
    borderRadius: tokens.borderRadiusLarge,
    backgroundColor: tokens.colorNeutralBackground1,
    boxShadow: tokens.shadow2,
  },
  hint: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
    lineHeight: tokens.lineHeightBase200,
  },
  row: {
    display: "flex",
    alignItems: "center",
    columnGap: "8px",
    rowGap: "8px",
    flexWrap: "wrap",
  },
  column: {
    display: "flex",
    flexDirection: "column",
    rowGap: "8px",
  },
  grid2: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr",
    gap: "8px",
  },
  code: {
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: tokens.fontSizeBase200,
    backgroundColor: tokens.colorNeutralBackground3,
    padding: "8px 10px",
    borderRadius: tokens.borderRadiusMedium,
    wordBreak: "break-all",
    whiteSpace: "pre-wrap",
  },
  detailList: {
    margin: "6px 0 0 0",
    paddingLeft: "18px",
    fontSize: tokens.fontSizeBase200,
  },
  scrollList: {
    display: "flex",
    flexDirection: "column",
    maxHeight: "200px",
    overflowY: "auto",
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
    padding: "4px 6px",
  },
  fullWidth: {
    width: "100%",
  },
  primaryButton: {
    width: "100%",
    minHeight: "36px",
  },
  muted: {
    opacity: 0.55,
    pointerEvents: "none",
  },
});

// ---------------------------------------------------------------------------
// Steps
// ---------------------------------------------------------------------------

const useStepStyles = makeStyles({
  header: {
    display: "flex",
    alignItems: "flex-start",
    columnGap: "10px",
  },
  badge: {
    flexShrink: 0,
    width: "24px",
    height: "24px",
    borderRadius: "50%",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundOnBrand,
  },
  badgeWaiting: {
    backgroundColor: tokens.colorNeutralBackground4,
    color: tokens.colorNeutralForeground3,
  },
  titles: {
    display: "flex",
    flexDirection: "column",
    rowGap: "2px",
    paddingTop: "2px",
  },
  title: {
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase300,
  },
});

export interface StepProps {
  number: number;
  title: string;
  description?: React.ReactNode;
  /** Greys the step out and explains why, until an earlier step is finished. */
  waitingFor?: string;
  children?: React.ReactNode;
}

/** One numbered stage of a multi-step tool. */
export const Step: React.FC<StepProps> = ({ number, title, description, waitingFor, children }) => {
  const shared = useSharedStyles();
  const styles = useStepStyles();
  const waiting = Boolean(waitingFor);

  return (
    <section className={shared.card} aria-disabled={waiting}>
      <div className={styles.header}>
        <span className={mergeClasses(styles.badge, waiting && styles.badgeWaiting)}>{number}</span>
        <div className={styles.titles}>
          <Text className={styles.title}>{title}</Text>
          {waiting ? (
            <Text className={shared.hint}>{waitingFor}</Text>
          ) : description ? (
            <Text className={shared.hint}>{description}</Text>
          ) : null}
        </div>
      </div>
      {waiting ? null : children}
    </section>
  );
};

// ---------------------------------------------------------------------------
// Collapsible action cards
// ---------------------------------------------------------------------------

const useActionCardStyles = makeStyles({
  toggle: {
    display: "flex",
    alignItems: "center",
    columnGap: "12px",
    width: "100%",
    padding: "0",
    border: "none",
    background: "none",
    cursor: "pointer",
    textAlign: "left",
    color: "inherit",
    fontFamily: "inherit",
  },
  icon: {
    flexShrink: 0,
    width: "36px",
    height: "36px",
    borderRadius: tokens.borderRadiusMedium,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    backgroundColor: tokens.colorBrandBackground2,
    color: tokens.colorBrandForeground2,
  },
  text: {
    display: "flex",
    flexDirection: "column",
    rowGap: "2px",
    flexGrow: 1,
  },
  title: {
    fontWeight: tokens.fontWeightSemibold,
  },
  chevron: {
    flexShrink: 0,
    color: tokens.colorNeutralForeground3,
  },
  body: {
    display: "flex",
    flexDirection: "column",
    rowGap: "12px",
    paddingTop: "4px",
  },
});

export interface ActionCardProps {
  icon: React.ReactNode;
  title: string;
  description: string;
  open: boolean;
  onToggle: () => void;
  children: React.ReactNode;
}

/** A tool that shows as one line until clicked, then unfolds its settings. */
export const ActionCard: React.FC<ActionCardProps> = ({
  icon,
  title,
  description,
  open,
  onToggle,
  children,
}) => {
  const shared = useSharedStyles();
  const styles = useActionCardStyles();
  const bodyId = React.useId();

  return (
    <section className={shared.card}>
      <button
        type="button"
        className={styles.toggle}
        onClick={onToggle}
        aria-expanded={open}
        aria-controls={bodyId}
      >
        <span className={styles.icon}>{icon}</span>
        <span className={styles.text}>
          <Text className={styles.title}>{title}</Text>
          <Text className={shared.hint}>{description}</Text>
        </span>
        <span className={styles.chevron}>{open ? <ChevronDown20Regular /> : <ChevronRight20Regular />}</span>
      </button>
      {open ? (
        <div id={bodyId} className={styles.body}>
          {children}
        </div>
      ) : null}
    </section>
  );
};

/** Settings most people never touch, hidden behind one link. */
export const MoreOptions: React.FC<{ children: React.ReactNode; label?: string }> = ({
  children,
  label = "More options",
}) => {
  const [open, setOpen] = React.useState(false);
  const shared = useSharedStyles();

  return (
    <div className={shared.column}>
      <div>
        <Button
          appearance="transparent"
          size="small"
          icon={open ? <ChevronDown20Regular /> : <ChevronRight20Regular />}
          onClick={() => setOpen((current) => !current)}
          aria-expanded={open}
        >
          {open ? "Fewer options" : label}
        </Button>
      </div>
      {open ? <div className={shared.column}>{children}</div> : null}
    </div>
  );
};

// ---------------------------------------------------------------------------
// Small text helpers
// ---------------------------------------------------------------------------

const useTipStyles = makeStyles({
  tip: {
    display: "flex",
    alignItems: "flex-start",
    columnGap: "6px",
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
    lineHeight: tokens.lineHeightBase200,
  },
  icon: {
    flexShrink: 0,
    marginTop: "1px",
    color: tokens.colorPaletteMarigoldForeground2,
  },
});

/** A one-line piece of advice with a lightbulb. */
export const Tip: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const styles = useTipStyles();
  return (
    <div className={styles.tip}>
      <Lightbulb16Regular className={styles.icon} />
      <span>{children}</span>
    </div>
  );
};

const useSelectionBarStyles = makeStyles({
  bar: {
    display: "flex",
    alignItems: "center",
    columnGap: "8px",
    padding: "8px 12px",
    marginBottom: "10px",
    borderRadius: tokens.borderRadiusLarge,
    backgroundColor: tokens.colorBrandBackground2,
    color: tokens.colorBrandForeground2,
  },
  text: {
    display: "flex",
    flexDirection: "column",
    minWidth: 0,
  },
  label: {
    fontSize: tokens.fontSizeBase100,
    textTransform: "uppercase",
    letterSpacing: "0.04em",
  },
  address: {
    fontWeight: tokens.fontWeightSemibold,
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
});

/** Always-visible reminder of exactly which cells a tool is about to change. */
export const SelectionBar: React.FC<{ label?: string; address: string; children?: React.ReactNode }> = ({
  label = "Working on",
  address,
  children,
}) => {
  const styles = useSelectionBarStyles();
  return (
    <div className={styles.bar} role="status" aria-live="polite">
      <TableSimple20Regular />
      <div className={styles.text}>
        <span className={styles.label}>{label}</span>
        <span className={styles.address} title={address}>
          {address}
        </span>
      </div>
      {children}
    </div>
  );
};

/** Turn "'Jan Orders'!A1:F9" into "Jan Orders, A1:F9" for display. */
export function friendlyAddress(address: string): string {
  const separator = address.lastIndexOf("!");
  if (separator < 0) {
    return address;
  }
  const sheet = address.slice(0, separator).replace(/^'(.*)'$/, "$1").replace(/''/g, "'");
  return `${sheet}, ${address.slice(separator + 1).replace(/\$/g, "")}`;
}

// ---------------------------------------------------------------------------
// Running commands
// ---------------------------------------------------------------------------

export interface ActionRunner {
  /** Id of the action currently running, if any. */
  busyId: string | null;
  busy: boolean;
  result: OperationResult | null;
  /** Id of the action that produced `result`. */
  resultId: string | null;
  run: (id: string, action: () => Promise<OperationResult>) => Promise<OperationResult>;
  clear: () => void;
}

/**
 * Run one command at a time, capturing whatever it returns - including anything
 * it throws, so a failure never leaves the pane stuck on a spinner.
 */
export function useActionRunner(): ActionRunner {
  const [busyId, setBusyId] = React.useState<string | null>(null);
  const [result, setResult] = React.useState<OperationResult | null>(null);
  const [resultId, setResultId] = React.useState<string | null>(null);

  const run = React.useCallback(async (id: string, action: () => Promise<OperationResult>) => {
    setBusyId(id);
    setResult(null);
    setResultId(null);
    let outcome: OperationResult;
    try {
      outcome = await action();
    } catch (error) {
      outcome = { ok: false, message: error instanceof Error ? error.message : String(error) };
    }
    setResult(outcome);
    setResultId(id);
    setBusyId(null);
    return outcome;
  }, []);

  const clear = React.useCallback(() => {
    setResult(null);
    setResultId(null);
  }, []);

  return { busyId, busy: busyId !== null, result, resultId, run, clear };
}

/** A success or error message, with any notes, that scrolls itself into view. */
export const ResultBanner: React.FC<{ result: OperationResult; onDismiss?: () => void }> = ({
  result,
  onDismiss,
}) => {
  const styles = useSharedStyles();
  const ref = React.useRef<HTMLDivElement | null>(null);

  React.useEffect(() => {
    ref.current?.scrollIntoView?.({ block: "nearest", behavior: "smooth" });
  }, [result]);

  return (
    <div ref={ref}>
      <MessageBar intent={result.ok ? "success" : "error"} layout="multiline">
        <MessageBarBody>
          <MessageBarTitle>{result.ok ? "Done" : "That didn't work"}</MessageBarTitle>
          {result.message}
          {result.details && result.details.length > 0 ? (
            <ul className={styles.detailList}>
              {result.details.map((detail, index) => (
                <li key={`${index}-${detail}`}>{detail}</li>
              ))}
            </ul>
          ) : null}
        </MessageBarBody>
        {onDismiss ? (
          <MessageBarActions
            containerAction={
              <Button
                appearance="transparent"
                aria-label="Dismiss"
                icon={<Dismiss20Regular />}
                onClick={onDismiss}
              />
            }
          />
        ) : null}
      </MessageBar>
    </div>
  );
};

export interface ActionButtonProps {
  runner: ActionRunner;
  /** Identifies this button so only its own result appears beneath it. */
  id: string;
  label: string;
  busyLabel?: string;
  icon?: React.ReactElement;
  appearance?: "primary" | "secondary";
  disabled?: boolean;
  /** Stretch to the full width of the card - use for a step's main action. */
  wide?: boolean;
  onRun: () => Promise<OperationResult>;
  /** Called after the action finishes, whatever the outcome. */
  onDone?: (result: OperationResult) => void;
}

/** A command button that reports its own outcome right below itself. */
export const ActionButton: React.FC<ActionButtonProps> = ({
  runner,
  id,
  label,
  busyLabel = "Working...",
  icon,
  appearance = "primary",
  disabled,
  wide,
  onRun,
  onDone,
}) => {
  const styles = useSharedStyles();
  const isBusy = runner.busyId === id;

  return (
    <div className={styles.column}>
      <Button
        appearance={appearance}
        className={wide ? styles.primaryButton : undefined}
        disabled={runner.busy || disabled}
        icon={isBusy ? <Spinner size="tiny" /> : icon}
        onClick={() => {
          void runner.run(id, onRun).then((result) => onDone?.(result));
        }}
      >
        {isBusy ? busyLabel : label}
      </Button>
      {runner.result && runner.resultId === id ? (
        <ResultBanner result={runner.result} onDismiss={runner.clear} />
      ) : null}
    </div>
  );
};

// ---------------------------------------------------------------------------
// Empty and choice states
// ---------------------------------------------------------------------------

const useChoiceStyles = makeStyles({
  group: {
    display: "grid",
    gap: "6px",
  },
  choice: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    rowGap: "4px",
    minHeight: "64px",
    padding: "8px 4px",
    borderRadius: tokens.borderRadiusMedium,
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    backgroundColor: tokens.colorNeutralBackground1,
    color: tokens.colorNeutralForeground1,
    cursor: "pointer",
    fontFamily: "inherit",
    fontSize: tokens.fontSizeBase200,
    ":hover": {
      backgroundColor: tokens.colorNeutralBackground1Hover,
    },
  },
  selected: {
    border: `2px solid ${tokens.colorBrandStroke1}`,
    backgroundColor: tokens.colorBrandBackground2,
    color: tokens.colorBrandForeground2,
    ":hover": {
      backgroundColor: tokens.colorBrandBackground2Hover,
    },
  },
  disabled: {
    opacity: 0.45,
    cursor: "not-allowed",
  },
});

export interface Choice<T extends string> {
  value: T;
  label: string;
  icon?: React.ReactNode;
  disabled?: boolean;
  title?: string;
}

/** A grid of big, picture-first buttons - friendlier than a dropdown for a handful of options. */
export function ChoiceGrid<T extends string>(props: {
  choices: Array<Choice<T>>;
  value: T;
  onChange: (value: T) => void;
  columns?: number;
  label: string;
}): React.ReactElement {
  const styles = useChoiceStyles();
  const { choices, value, onChange, columns = 3, label } = props;

  return (
    <div
      role="radiogroup"
      aria-label={label}
      className={styles.group}
      style={{ gridTemplateColumns: `repeat(${columns}, 1fr)` }}
    >
      {choices.map((choice) => (
        <button
          key={choice.value}
          type="button"
          role="radio"
          aria-checked={choice.value === value}
          title={choice.title}
          disabled={choice.disabled}
          className={mergeClasses(
            styles.choice,
            choice.value === value && styles.selected,
            choice.disabled && styles.disabled
          )}
          onClick={() => onChange(choice.value)}
        >
          {choice.icon}
          <span>{choice.label}</span>
        </button>
      ))}
    </div>
  );
}
