import * as React from "react";
import { Button, makeStyles, mergeClasses, Text, tokens } from "@fluentui/react-components";
import {
  CheckmarkCircle20Filled,
  DataBarHorizontal20Regular,
  DataBarVertical20Regular,
  DataHistogram20Regular,
  DataLine20Regular,
  DataPie20Regular,
  DataScatter20Regular,
  ArrowUpload20Regular,
  ErrorCircle20Filled,
  Send20Filled,
} from "@fluentui/react-icons";

import {
  AssistantState,
  BotContent,
  compareAgainstFile,
  BotReply,
  Chip,
  helpGroups,
  INITIAL_STATE,
  respond,
  STARTER_CHIPS,
} from "../features/assistant";
import { CommandInfo, completeCommand, ToolName } from "../features/commands";
import type { DashChartKind } from "../features/dashboard";
import { groupIssues, summarize } from "../features/review";
import { SUPPORTED_EXTENSIONS } from "../shared/workbookReader";
import { watchSheet, Watcher, WatchUpdate } from "../features/watch";
import { DISPLAY_FONT } from "../theme";

interface ChatMessage {
  id: number;
  from: "user" | "bot";
  text?: string;
  reply?: BotReply;
}

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
  root: {
    display: "flex",
    flexDirection: "column",
    height: "100%",
    minHeight: 0,
  },
  log: {
    flexGrow: 1,
    overflowY: "auto",
    padding: "14px 12px 6px",
    display: "flex",
    flexDirection: "column",
    rowGap: "12px",
  },
  botRow: {
    display: "flex",
    alignItems: "flex-start",
    columnGap: "8px",
    maxWidth: "100%",
  },
  avatar: {
    flexShrink: 0,
    width: "26px",
    height: "26px",
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorBrandBackground,
    display: "grid",
    placeItems: "center",
    marginTop: "2px",
  },
  avatarMark: {
    width: "9px",
    height: "9px",
    backgroundColor: tokens.colorNeutralForegroundOnBrand,
    borderRadius: "2px",
  },
  botBody: {
    display: "flex",
    flexDirection: "column",
    rowGap: "8px",
    minWidth: 0,
    flexGrow: 1,
  },
  bubble: {
    backgroundColor: tokens.colorNeutralBackground1,
    borderRadius: "4px 12px 12px 12px",
    boxShadow: tokens.shadow2,
    padding: "10px 12px",
    display: "flex",
    flexDirection: "column",
    rowGap: "8px",
    fontSize: tokens.fontSizeBase300,
    lineHeight: tokens.lineHeightBase300,
    overflowWrap: "anywhere",
  },
  userRow: {
    display: "flex",
    justifyContent: "flex-end",
  },
  userBubble: {
    maxWidth: "85%",
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundOnBrand,
    borderRadius: "12px 4px 12px 12px",
    padding: "8px 12px",
    fontSize: tokens.fontSizeBase300,
    overflowWrap: "anywhere",
  },
  chips: {
    display: "flex",
    flexWrap: "wrap",
    gap: "6px",
  },
  chip: {
    borderRadius: tokens.borderRadiusCircular,
  },
  hint: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
    lineHeight: tokens.lineHeightBase200,
  },
  headline: {
    display: "flex",
    alignItems: "baseline",
    columnGap: "8px",
  },
  bigNumber: {
    fontFamily: DISPLAY_FONT,
    fontWeight: 700,
    fontSize: "28px",
    lineHeight: 1,
    color: tokens.colorPaletteMarigoldForeground1,
    fontVariantNumeric: "tabular-nums",
  },
  list: {
    display: "flex",
    flexDirection: "column",
    margin: 0,
    padding: 0,
    listStyle: "none",
  },
  item: {
    display: "flex",
    alignItems: "flex-start",
    columnGap: "8px",
    padding: "5px 0",
    borderBottom: `1px solid ${tokens.colorNeutralStroke3}`,
    ":last-child": { borderBottom: "none" },
  },
  itemText: {
    display: "flex",
    flexDirection: "column",
    flexGrow: 1,
    minWidth: 0,
    fontSize: tokens.fontSizeBase200,
  },
  itemTitle: {
    fontWeight: tokens.fontWeightSemibold,
  },
  off: {
    opacity: 0.5,
    textDecoration: "line-through",
  },
  count: {
    flexShrink: 0,
    padding: "0 7px",
    borderRadius: tokens.borderRadiusCircular,
    backgroundColor: tokens.colorNeutralBackground3,
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    fontVariantNumeric: "tabular-nums",
  },
  example: {
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
  },
  number: {
    flexShrink: 0,
    width: "20px",
    height: "20px",
    borderRadius: "50%",
    display: "grid",
    placeItems: "center",
    fontSize: tokens.fontSizeBase100,
    fontWeight: tokens.fontWeightSemibold,
    backgroundColor: tokens.colorBrandBackground2,
    color: tokens.colorBrandForeground2,
  },
  kindIcon: {
    flexShrink: 0,
    display: "flex",
    color: tokens.colorBrandForeground1,
  },
  severity: {
    flexShrink: 0,
    width: "8px",
    height: "8px",
    borderRadius: "2px",
    marginTop: "5px",
  },
  sevError: { backgroundColor: tokens.colorPaletteRedForeground1 },
  sevWarning: { backgroundColor: tokens.colorPaletteMarigoldForeground1 },
  sevTidy: { backgroundColor: tokens.colorBrandForeground1 },
  result: {
    display: "flex",
    alignItems: "flex-start",
    columnGap: "8px",
  },
  ok: { color: tokens.colorPaletteGreenForeground1, flexShrink: 0 },
  bad: { color: tokens.colorPaletteRedForeground1, flexShrink: 0 },
  details: {
    margin: "4px 0 0",
    paddingLeft: "16px",
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
  },
  groupLabel: {
    fontSize: tokens.fontSizeBase100,
    fontWeight: tokens.fontWeightSemibold,
    letterSpacing: "0.06em",
    textTransform: "uppercase",
    color: tokens.colorNeutralForeground3,
    marginTop: "4px",
  },
  commandButton: {
    display: "flex",
    flexDirection: "column",
    alignItems: "flex-start",
    width: "100%",
    padding: "5px 6px",
    border: "none",
    borderRadius: tokens.borderRadiusMedium,
    background: "none",
    textAlign: "left",
    cursor: "pointer",
    color: tokens.colorNeutralForeground1,
    fontFamily: "inherit",
    ":hover": { backgroundColor: tokens.colorNeutralBackground1Hover },
  },
  commandName: {
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorBrandForeground1,
  },
  columnChips: {
    display: "flex",
    flexWrap: "wrap",
    gap: "4px",
  },
  columnChip: {
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: tokens.fontSizeBase100,
    padding: "1px 6px",
    borderRadius: tokens.borderRadiusSmall,
    backgroundColor: tokens.colorNeutralBackground3,
  },
  /** A formula shown back to the reader: monospaced, and allowed to wrap. */
  formula: {
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: tokens.fontSizeBase200,
    padding: "5px 7px",
    borderRadius: tokens.borderRadiusSmall,
    backgroundColor: tokens.colorNeutralBackground3,
    overflowWrap: "anywhere",
  },
  typing: {
    display: "flex",
    columnGap: "4px",
    padding: "4px 2px",
  },
  dot: {
    width: "6px",
    height: "6px",
    borderRadius: "50%",
    backgroundColor: tokens.colorNeutralForeground3,
    animationName: {
      "0%, 80%, 100%": { opacity: 0.25, transform: "translateY(0)" },
      "40%": { opacity: 1, transform: "translateY(-3px)" },
    },
    animationDuration: "1.1s",
    animationIterationCount: "infinite",
    "@media (prefers-reduced-motion: reduce)": { animationName: "none" },
  },
  composer: {
    flexShrink: 0,
    position: "relative",
    padding: "8px 10px 10px",
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground1,
  },
  inputRow: {
    display: "flex",
    alignItems: "flex-end",
    columnGap: "6px",
    padding: "4px 4px 4px 12px",
    borderRadius: tokens.borderRadiusXLarge,
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    backgroundColor: tokens.colorNeutralBackground1,
    ":focus-within": {
      border: `1px solid ${tokens.colorBrandStroke1}`,
      boxShadow: `0 0 0 1px ${tokens.colorBrandStroke1}`,
    },
  },
  input: {
    flexGrow: 1,
    minWidth: 0,
    border: "none",
    outline: "none",
    resize: "none",
    background: "transparent",
    color: tokens.colorNeutralForeground1,
    fontFamily: "inherit",
    fontSize: tokens.fontSizeBase300,
    lineHeight: "20px",
    padding: "6px 0",
    maxHeight: "96px",
  },
  popup: {
    position: "absolute",
    left: "10px",
    right: "10px",
    bottom: "calc(100% - 2px)",
    maxHeight: "260px",
    overflowY: "auto",
    backgroundColor: tokens.colorNeutralBackground1,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusLarge,
    boxShadow: tokens.shadow16,
    padding: "4px",
    zIndex: 10,
  },
  popupItem: {
    display: "flex",
    flexDirection: "column",
    width: "100%",
    padding: "6px 8px",
    border: "none",
    borderRadius: tokens.borderRadiusMedium,
    background: "none",
    textAlign: "left",
    cursor: "pointer",
    color: tokens.colorNeutralForeground1,
    fontFamily: "inherit",
  },
  popupActive: {
    backgroundColor: tokens.colorBrandBackground2,
  },
  footnote: {
    marginTop: "5px",
    textAlign: "center",
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase100,
  },
});

const WELCOME: ChatMessage = {
  id: 0,
  from: "bot",
  reply: {
    content: [
      {
        type: "text",
        text: "Hi, I'm MEx. I fix messy spreadsheets and turn them into dashboards. Open the sheet you want to work on, then pick one below, or type / to see every command.",
      },
    ],
    chips: STARTER_CHIPS,
  },
};

export interface ChatPanelProps {
  onOpenTool: (tool: ToolName) => void;
}

const ChatPanel: React.FC<ChatPanelProps> = ({ onOpenTool }) => {
  const styles = useStyles();
  const [messages, setMessages] = React.useState<ChatMessage[]>([WELCOME]);
  const [state, setState] = React.useState<AssistantState>(INITIAL_STATE);
  const [draft, setDraft] = React.useState("");
  const [busy, setBusy] = React.useState(false);
  const [highlight, setHighlight] = React.useState(0);
  const [history, setHistory] = React.useState<string[]>([]);
  const logRef = React.useRef<HTMLDivElement | null>(null);
  const inputRef = React.useRef<HTMLTextAreaElement | null>(null);
  const nextId = React.useRef(1);

  // Autocomplete shows while typing the command word itself ("/da").
  const suggestions = busy ? [] : completeCommand(draft);
  const popupOpen = suggestions.length > 0 && !suggestions.some((s) => s.command === draft.trim());

  React.useEffect(() => setHighlight(0), [draft]);

  React.useEffect(() => {
    const log = logRef.current;
    if (log) log.scrollTo?.({ top: log.scrollHeight, behavior: "smooth" });
  }, [messages, busy]);

  // ---------------------------------------------------------------------------
  // Watching the sheet
  //
  // The subscription has to outlive the message that started it, so it lives in
  // a ref here rather than in the assistant's state. Findings are posted as
  // ordinary bot messages, and nothing in the sheet is changed.
  // ---------------------------------------------------------------------------
  const watcher = React.useRef<Watcher | null>(null);

  const postWatchUpdate = React.useCallback((update: WatchUpdate) => {
    if (update.tooBig) return;
    if (update.issues.length === 0) return;
    const groups = groupIssues(update.issues);
    setMessages((current) => [
      ...current,
      {
        id: nextId.current++,
        from: "bot",
        reply: {
          content: [
            {
              type: "issues",
              sheet: update.sheet,
              groups,
              summary: `${summarize(update.issues)} in ${update.address}, just edited`,
              marked: 0,
              structuralRows: 0,
              truncated: false,
            },
          ],
          chips: [
            { label: "Check the whole sheet", send: `/review ${update.sheet}` },
            { label: "Stop watching", send: "/watch off" },
          ],
        },
      },
    ]);
  }, []);

  const applyWatch = React.useCallback(
    async (request: { action: "start" | "stop"; sheet?: string }) => {
      if (watcher.current) {
        await watcher.current.stop();
        watcher.current = null;
      }
      if (request.action === "stop") return;

      const started = await watchSheet(request.sheet, postWatchUpdate);
      if ("error" in started) {
        setMessages((current) => [
          ...current,
          {
            id: nextId.current++,
            from: "bot",
            reply: { content: [{ type: "text", text: started.error }], chips: STARTER_CHIPS },
          },
        ]);
        return;
      }
      watcher.current = started;
    },
    [postWatchUpdate]
  );

  // Leaving the pane open with a dead handler would leak; stop on unmount.
  React.useEffect(
    () => () => {
      void watcher.current?.stop();
      watcher.current = null;
    },
    []
  );

  // ---------------------------------------------------------------------------
  // Picking a file to compare against
  //
  // An add-in can't reach the file system, so a reply that needs last month's
  // export asks the pane to open a picker. The file is read here and never
  // leaves the machine.
  // ---------------------------------------------------------------------------
  const [fileRequest, setFileRequest] = React.useState<{ sheet: string } | null>(null);
  const fileInputRef = React.useRef<HTMLInputElement | null>(null);

  const takeFile = React.useCallback(
    async (file: File | undefined) => {
      if (!file || !fileRequest) return;
      const sheet = fileRequest.sheet;
      setFileRequest(null);
      setMessages((current) => [
        ...current,
        { id: nextId.current++, from: "user", text: file.name },
      ]);
      setBusy(true);
      const reply = await compareAgainstFile(sheet, file);
      setMessages((current) => [...current, { id: nextId.current++, from: "bot", reply }]);
      setBusy(false);
    },
    [fileRequest]
  );

  const send = React.useCallback(
    async (text: string) => {
      const trimmed = text.trim();
      if (trimmed === "" || busy) return;
      setDraft("");
      setHistory((current) => [...current.filter((item) => item !== trimmed), trimmed].slice(-30));
      setMessages((current) => [...current, { id: nextId.current++, from: "user", text: trimmed }]);
      setBusy(true);
      const { reply, state: nextState } = await respond(trimmed, state);
      setState(nextState);
      setMessages((current) => [...current, { id: nextId.current++, from: "bot", reply }]);
      setBusy(false);
      if (reply.openTool) onOpenTool(reply.openTool);
      if (reply.watch) await applyWatch(reply.watch);
      if (reply.pickFile) {
        setFileRequest({ sheet: reply.pickFile.sheet });
        // Opening the picker straight away saves a click, and the button stays
        // in the composer for anyone who dismisses the dialog.
        fileInputRef.current?.click();
      }
      inputRef.current?.focus();
    },
    [busy, state, onOpenTool, applyWatch]
  );

  const choose = (info: CommandInfo) => {
    // Commands that take words get a trailing space so the user can keep typing.
    setDraft(info.example === info.command ? info.command : `${info.command} `);
    inputRef.current?.focus();
  };

  const onKeyDown = (event: React.KeyboardEvent<HTMLTextAreaElement>) => {
    if (popupOpen) {
      if (event.key === "ArrowDown") {
        event.preventDefault();
        setHighlight((i) => (i + 1) % suggestions.length);
        return;
      }
      if (event.key === "ArrowUp") {
        event.preventDefault();
        setHighlight((i) => (i - 1 + suggestions.length) % suggestions.length);
        return;
      }
      if (event.key === "Tab" || (event.key === "Enter" && !event.shiftKey)) {
        event.preventDefault();
        choose(suggestions[highlight]);
        return;
      }
      if (event.key === "Escape") {
        setDraft("");
        return;
      }
    }
    if (event.key === "Enter" && !event.shiftKey) {
      event.preventDefault();
      void send(draft);
    } else if (event.key === "ArrowUp" && draft === "" && history.length > 0) {
      event.preventDefault();
      setDraft(history[history.length - 1]);
    }
  };

  const lastBotId = [...messages].reverse().find((m) => m.from === "bot")?.id;

  return (
    <div className={styles.root}>
      <div className={styles.log} ref={logRef} role="log" aria-live="polite" aria-label="Conversation with MEx">
        {messages.map((message) =>
          message.from === "user" ? (
            <div key={message.id} className={styles.userRow}>
              <div className={styles.userBubble}>{message.text}</div>
            </div>
          ) : (
            <div key={message.id} className={styles.botRow}>
              <span className={styles.avatar} aria-hidden="true">
                <span className={styles.avatarMark} />
              </span>
              <div className={styles.botBody}>
                <div className={styles.bubble}>
                  {message.reply?.content.map((content, i) => (
                    <Content key={i} content={content} onCommand={(text) => { setDraft(text); inputRef.current?.focus(); }} />
                  ))}
                </div>
                {/* Only the latest reply offers chips, so old choices don't pile up. */}
                {message.id === lastBotId && !busy && message.reply && message.reply.chips.length > 0 ? (
                  <ChipRow chips={message.reply.chips} onPick={(c) => void send(c.send)} />
                ) : null}
              </div>
            </div>
          )
        )}
        {busy ? (
          <div className={styles.botRow} aria-label="MEx is working">
            <span className={styles.avatar} aria-hidden="true">
              <span className={styles.avatarMark} />
            </span>
            <div className={styles.bubble}>
              <div className={styles.typing}>
                {[0, 1, 2].map((i) => (
                  <span key={i} className={styles.dot} style={{ animationDelay: `${i * 0.15}s` }} />
                ))}
              </div>
            </div>
          </div>
        ) : null}
      </div>

      <div className={styles.composer}>
        <input
          ref={fileInputRef}
          type="file"
          accept={SUPPORTED_EXTENSIONS.join(",")}
          style={{ display: "none" }}
          onChange={(event) => {
            void takeFile(event.target.files?.[0]);
            event.target.value = "";
          }}
        />
        {fileRequest ? (
          <Button
            className={styles.chip}
            size="small"
            appearance="primary"
            icon={<ArrowUpload20Regular />}
            onClick={() => fileInputRef.current?.click()}
          >
            Choose the file to compare against
          </Button>
        ) : null}
        {popupOpen ? (
          <div className={styles.popup} role="listbox" aria-label="Commands">
            {suggestions.map((info, i) => (
              <button
                key={info.command}
                type="button"
                role="option"
                aria-selected={i === highlight}
                className={mergeClasses(styles.popupItem, i === highlight && styles.popupActive)}
                onMouseEnter={() => setHighlight(i)}
                onClick={() => choose(info)}
              >
                <span className={styles.commandName}>{info.example}</span>
                <span className={styles.hint}>{info.description}</span>
              </button>
            ))}
          </div>
        ) : null}
        <div className={styles.inputRow}>
          <textarea
            ref={inputRef}
            id="mex-command"
            className={styles.input}
            rows={1}
            value={draft}
            placeholder="Type a command, or / to see them all"
            aria-label="Message MEx"
            onChange={(event) => setDraft(event.target.value)}
            onKeyDown={onKeyDown}
          />
          <Button
            appearance="primary"
            shape="circular"
            icon={<Send20Filled />}
            aria-label="Send"
            disabled={busy || draft.trim() === ""}
            onClick={() => void send(draft)}
          />
        </div>
        <div className={styles.footnote}>Works on the sheet you have open. Your data never leaves Excel.</div>
      </div>
    </div>
  );
};

const ChipRow: React.FC<{ chips: Chip[]; onPick: (chip: Chip) => void }> = ({ chips, onPick }) => {
  const styles = useStyles();
  return (
    <div className={styles.chips}>
      {chips.map((c, i) => (
        <Button
          key={c.label}
          size="small"
          className={styles.chip}
          appearance={i === 0 ? "primary" : "outline"}
          onClick={() => onPick(c)}
        >
          {c.label}
        </Button>
      ))}
    </div>
  );
};

const Content: React.FC<{ content: BotContent; onCommand: (text: string) => void }> = ({ content, onCommand }) => {
  const styles = useStyles();

  switch (content.type) {
    case "text":
      return <Text>{content.text}</Text>;

    case "help":
      return (
        <>
          <Text>Here&apos;s everything I can do. Tap one to put it in the message box.</Text>
          {helpGroups().map((group) => (
            <div key={group.group}>
              <div className={styles.groupLabel}>{group.group}</div>
              {group.commands.map((info) => (
                <button key={info.command} type="button" className={styles.commandButton} onClick={() => onCommand(info.example)}>
                  <span className={styles.commandName}>{info.example}</span>
                  <span className={styles.hint}>{info.description}</span>
                </button>
              ))}
            </div>
          ))}
          <Text className={styles.hint}>You can also just write it: &quot;remove duplicates&quot;, &quot;pie of revenue by region&quot;.</Text>
        </>
      );

    case "findings": {
      const toFix = content.findings.filter((f) => !content.skip.includes(f.id)).reduce((s, f) => s + f.count, 0);
      return (
        <>
          <div className={styles.headline}>
            <span className={styles.bigNumber}>{toFix.toLocaleString()}</span>
            <Text weight="semibold">
              problems in {content.rows.toLocaleString()} rows of &quot;{content.sheet}&quot;
            </Text>
          </div>
          <ul className={styles.list}>
            {content.findings.map((finding) => {
              const skipped = content.skip.includes(finding.id);
              return (
                <li key={finding.id} className={styles.item}>
                  <span className={mergeClasses(styles.itemText, skipped && styles.off)}>
                    <span className={styles.itemTitle}>{finding.label}</span>
                    {finding.examples.slice(0, 2).map((example) => (
                      <span key={example} className={styles.example} title={example}>
                        {example}
                      </span>
                    ))}
                  </span>
                  <span className={styles.count}>{skipped ? "kept" : finding.count.toLocaleString()}</span>
                </li>
              );
            })}
          </ul>
          <Text className={styles.hint}>
            Nothing has changed yet. The clean table will go on a new sheet; &quot;{content.sheet}&quot; stays as it is. To
            leave something alone, say e.g. &quot;skip totals&quot;.
          </Text>
        </>
      );
    }

    case "charts":
      return (
        <>
          <Text>
            For &quot;{content.sheet}&quot; ({content.rows.toLocaleString()} rows) I&apos;d build these:
          </Text>
          <ul className={styles.list}>
            {content.charts.map((chart, i) => {
              const on = content.chosen.includes(i + 1);
              return (
                <li key={chart.id} className={styles.item}>
                  <span className={styles.number}>{i + 1}</span>
                  <span className={styles.kindIcon}>{KIND_ICON[chart.kind]}</span>
                  <span className={mergeClasses(styles.itemText, !on && styles.off)}>
                    <span className={styles.itemTitle}>{chart.title}</span>
                    <span className={styles.hint}>{chart.reason}</span>
                  </span>
                </li>
              );
            })}
          </ul>
          {content.kpis.length > 0 ? <Text className={styles.hint}>Plus headline numbers: {content.kpis.join(", ")}.</Text> : null}
          {content.problemsFixed > 0 ? (
            <Text className={styles.hint}>
              {content.problemsFixed.toLocaleString()} data problems will be smoothed over for the charts; the sheet itself isn&apos;t
              changed.
            </Text>
          ) : null}
          <Text className={styles.hint}>
            To change the set, say &quot;build 1, 2 and 5&quot;, &quot;without 7&quot;, or add your own: &quot;add a pie
            of Units by Channel&quot;.
          </Text>
        </>
      );

    case "result":
      return (
        <div className={styles.result}>
          {content.result.ok ? <CheckmarkCircle20Filled className={styles.ok} /> : <ErrorCircle20Filled className={styles.bad} />}
          <div>
            <Text weight="semibold">{content.result.ok ? "Done. " : "That didn't work. "}</Text>
            <Text>{content.result.message}</Text>
            {content.result.details && content.result.details.length > 0 ? (
              <ul className={styles.details}>
                {content.result.details.map((detail, i) => (
                  <li key={`${i}-${detail}`}>{detail}</li>
                ))}
              </ul>
            ) : null}
          </div>
        </div>
      );

    case "issues":
      return (
        <>
          <div className={styles.headline}>
            <span className={styles.bigNumber}>
              {content.groups.reduce((total, group) => total + group.issues.length, 0).toLocaleString()}
            </span>
            <Text weight="semibold">
              to look at on &quot;{content.sheet}&quot;: {content.summary}
            </Text>
          </div>
          <ul className={styles.list}>
            {content.groups.map((group, i) => (
              <li key={group.group} className={styles.item}>
                <span className={styles.number}>{i + 1}</span>
                <span
                  className={mergeClasses(
                    styles.severity,
                    group.severity === "error" && styles.sevError,
                    group.severity === "warning" && styles.sevWarning,
                    group.severity === "tidy" && styles.sevTidy
                  )}
                  aria-label={group.severity === "error" ? "Error" : group.severity === "warning" ? "Check" : "Tidy up"}
                />
                <span className={styles.itemText}>
                  <span className={styles.itemTitle}>
                    {group.title}
                    {group.header ? ` in ${group.header}` : ""}
                  </span>
                  <span className={styles.hint}>
                    {group.issues.length === 1 ? group.issues[0].address : `${group.issues.length} cells`}
                    {group.fixable > 0 ? ` · ${group.fixable} fixable` : " · needs your call"}
                  </span>
                </span>
                <span className={styles.count}>{group.issues.length.toLocaleString()}</span>
              </li>
            ))}
          </ul>
          <Text className={styles.hint}>
            {content.marked > 0
              ? `Marked in the sheet: red for errors, amber to check, blue for tidy-ups. Say "show 2" to jump to one, "fix 2" to fix that kind, or "ignore 2" to stop reporting it.`
              : `Say "show 2" to jump to one, "fix 2" to fix that kind, or "ignore 2" to stop reporting it.`}
          </Text>
          {content.structuralRows > 0 ? (
            <Text className={styles.hint}>
              This sheet also has {content.structuralRows.toLocaleString()} rows that aren&apos;t data (titles, totals or
              group headings). /fix handles those.
            </Text>
          ) : null}
          {content.truncated ? <Text className={styles.hint}>Only the first 20,000 rows were checked.</Text> : null}
        </>
      );

    case "columns":
      return (
        <>
          <Text className={styles.hint}>Columns on &quot;{content.sheet}&quot;:</Text>
          <div className={styles.columnChips}>
            {content.headers.map((header) => (
              <span key={header} className={styles.columnChip}>
                {header}
              </span>
            ))}
          </div>
        </>
      );

    case "crosstab":
      return (
        <>
          <Text weight="semibold">
            &quot;{content.sheet}&quot; is a crosstab: {content.shape}.
          </Text>
          <Text className={styles.hint}>These stay as they are, one per row:</Text>
          <div className={styles.columnChips}>
            {content.keys.map((key) => (
              <span key={key} className={styles.columnChip}>
                {key}
              </span>
            ))}
          </div>
          <Text className={styles.hint}>These become rows instead of columns:</Text>
          <div className={styles.columnChips}>
            {content.periods.slice(0, 14).map((period) => (
              <span key={period} className={styles.columnChip}>
                {period}
              </span>
            ))}
            {content.periods.length > 14 ? (
              <span className={styles.columnChip}>+{content.periods.length - 14} more</span>
            ) : null}
          </div>
          <Text className={styles.hint}>
            I&apos;ll write the tidy version to a new sheet. &quot;{content.sheet}&quot; is left exactly as it is.
          </Text>
        </>
      );

    case "guards":
      return (
        <>
          <Text weight="semibold">
            {content.guards.length} {content.guards.length === 1 ? "rule" : "rules"} for &quot;
            {content.sheet}&quot;, from what its {content.rows.toLocaleString()} rows already hold
          </Text>
          <ul className={styles.list}>
            {content.guards.map((guard) => (
              <li key={guard.header} className={styles.item}>
                <span className={styles.itemText}>
                  <span className={styles.itemTitle}>{guard.title}</span>
                  <span className={styles.hint}>{guard.detail}</span>
                  {content.existingProblems[guard.header] > 0 ? (
                    <span className={styles.hint}>
                      {content.existingProblems[guard.header].toLocaleString()} cells already there
                      would be flagged by this — worth a look before you add it.
                    </span>
                  ) : null}
                </span>
              </li>
            ))}
          </ul>
          <Text className={styles.hint}>
            Each rule warns rather than blocks, so you can still enter something unusual on purpose,
            and reaches past today&apos;s rows so it covers what gets typed next.
          </Text>
        </>
      );

    case "health":
      return (
        <>
          <Text weight="semibold">{content.summary}</Text>
          <ul className={styles.list}>
            {content.findings.map((finding, i) => (
              <li key={`${finding.sheet}-${finding.kind}-${i}`} className={styles.item}>
                <span
                  className={mergeClasses(
                    styles.severity,
                    finding.severity === "slow" && styles.sevWarning,
                    finding.severity === "fragile" && styles.sevError,
                    finding.severity === "tidy" && styles.sevTidy
                  )}
                  aria-label={
                    finding.severity === "slow"
                      ? "Makes it slower"
                      : finding.severity === "fragile"
                        ? "Could break quietly"
                        : "Worth tidying"
                  }
                />
                <span className={styles.itemText}>
                  <span className={styles.itemTitle}>{finding.title}</span>
                  <span className={styles.hint}>{finding.detail}</span>
                  <span className={styles.hint}>
                    {finding.sheet}
                    {finding.examples.length > 0 ? ` · ${finding.examples.join(", ")}` : ""}
                    {finding.fix ? " · fixable" : ""}
                  </span>
                </span>
              </li>
            ))}
          </ul>
          <Text className={styles.hint}>
            Looked at {content.sheetsScanned} {content.sheetsScanned === 1 ? "sheet" : "sheets"}.
            {content.sheetsSkipped.length > 0
              ? ` Too big to read: ${content.sheetsSkipped.join(", ")}.`
              : ""}{" "}
            Nothing has been changed.
          </Text>
        </>
      );

    case "insights":
      return (
        <>
          <Text weight="semibold">What &quot;{content.sheet}&quot; shows</Text>
          <ul className={styles.details}>
            {content.lines.map((line) => (
              <li key={line}>{line}</li>
            ))}
          </ul>
          <Text className={styles.hint}>
            These are straight arithmetic on your own numbers - no guessing about why.
          </Text>
        </>
      );

    case "explain":
      return (
        <>
          <Text className={styles.hint}>{content.address}</Text>
          {content.formula ? <code className={styles.formula}>{content.formula}</code> : null}
          <Text>{content.english}</Text>
          {content.references.length > 0 ? (
            <>
              <Text className={styles.hint}>It reads:</Text>
              <div className={styles.columnChips}>
                {content.references.map((reference) => (
                  <span key={reference} className={styles.columnChip}>
                    {reference}
                  </span>
                ))}
              </div>
            </>
          ) : null}
          {content.warnings.length > 0 ? (
            <ul className={styles.details}>
              {content.warnings.map((warning) => (
                <li key={warning.kind}>{warning.text}</li>
              ))}
            </ul>
          ) : null}
        </>
      );

    case "recipes":
      return (
        <>
          <Text weight="semibold">Saved in this workbook</Text>
          {content.recipes.length === 0 ? (
            <Text className={styles.hint}>Nothing saved yet.</Text>
          ) : (
            <ul className={styles.list}>
              {content.recipes.map((recipe) => (
                <li key={recipe.name} className={styles.item}>
                  <span className={styles.itemText}>
                    <span className={styles.itemTitle}>{recipe.name}</span>
                    <span className={styles.hint}>
                      {recipe.steps} step{recipe.steps === 1 ? "" : "s"}
                    </span>
                  </span>
                  <Button size="small" appearance="outline" onClick={() => onCommand(`/recipe run ${recipe.name}`)}>
                    Run
                  </Button>
                </li>
              ))}
            </ul>
          )}
          {content.recorded > 0 ? (
            <Text className={styles.hint}>
              I&apos;ve recorded {content.recorded} step{content.recorded === 1 ? "" : "s"} this session. Say
              &quot;/recipe save month end&quot; to keep them.
            </Text>
          ) : (
            <Text className={styles.hint}>
              Run a few commands and I&apos;ll remember the sequence, ready to save as a recipe.
            </Text>
          )}
        </>
      );

    default:
      return null;
  }
};

export default ChatPanel;
