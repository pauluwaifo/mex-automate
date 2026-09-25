import * as React from "react";
import { Button, makeStyles, mergeClasses, Text, tokens } from "@fluentui/react-components";
import {
  Apps20Regular,
  ArrowLeft20Regular,
  ArrowSync24Regular,
  Board24Regular,
  Broom24Regular,
  ChevronRight20Regular,
  DataPie24Regular,
  MathFormula24Regular,
  TableSparkle24Regular,
  TableStackBelow24Regular,
} from "@fluentui/react-icons";

import type { ToolName } from "../features/commands";
import { DISPLAY_FONT } from "../theme";
import ChartPanel from "./ChartPanel";
import ChatPanel from "./ChatPanel";
import CleanPanel from "./CleanPanel";
import DashboardPanel from "./DashboardPanel";
import FormulaPanel from "./FormulaPanel";
import MergePanel from "./MergePanel";
import ReportPanel from "./ReportPanel";
import TidyPanel from "./TidyPanel";
import { Tip } from "./ui";

export type ToolId = "tidy" | "dashboard" | "clean" | "merge" | "charts" | "formulas" | "reports";

/** Lets one tool hand the user on to another, e.g. "Build a dashboard from it". */
export interface ToolParams {
  /** A data source key such as "sheet:Sales (clean)". */
  source?: string;
}

export interface ToolProps {
  navigate: (tool: ToolId, params?: ToolParams) => void;
  params?: ToolParams;
}

interface Tool {
  id: ToolId;
  title: string;
  description: string;
  icon: React.ReactElement;
  panel: React.FC<ToolProps>;
  /** Shown large at the top of the home screen. */
  featured?: boolean;
}

/** The home screen, in the order most people reach for them. */
const TOOLS: Tool[] = [
  {
    id: "tidy",
    title: "Fix messy data",
    description: "Titles, totals, numbers stored as text, mixed dates, typos: found and fixed in one click.",
    icon: <TableSparkle24Regular />,
    panel: TidyPanel,
    featured: true,
  },
  {
    id: "dashboard",
    title: "Build a dashboard",
    description: "Up to 8 charts and headline numbers, chosen and laid out for you.",
    icon: <Board24Regular />,
    panel: DashboardPanel,
    featured: true,
  },
  {
    id: "clean",
    title: "Quick clean-ups",
    description: "One fix at a time: duplicates, spaces, dates, capital letters.",
    icon: <Broom24Regular />,
    panel: CleanPanel,
  },
  {
    id: "merge",
    title: "Combine sheets",
    description: "Stack several sheets or files into one tidy table.",
    icon: <TableStackBelow24Regular />,
    panel: MergePanel,
  },
  {
    id: "charts",
    title: "Make a chart",
    description: "Turn your data into a chart or a summary table.",
    icon: <DataPie24Regular />,
    panel: ChartPanel,
  },
  {
    id: "formulas",
    title: "Formulas",
    description: "Copy a formula down a column, or insert a ready-made one.",
    icon: <MathFormula24Regular />,
    panel: FormulaPanel,
  },
  {
    id: "reports",
    title: "Update a report",
    description: "Put new data into an existing report without breaking its layout.",
    icon: <ArrowSync24Regular />,
    panel: ReportPanel,
  },
];

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    height: "100vh",
    backgroundColor: tokens.colorNeutralBackground2,
  },
  header: {
    flexShrink: 0,
    display: "flex",
    alignItems: "center",
    columnGap: "8px",
    minHeight: "48px",
    padding: "0 12px",
    backgroundColor: tokens.colorNeutralBackground1,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  headerIcon: {
    display: "flex",
    color: tokens.colorBrandForeground1,
  },
  headerTitle: {
    fontFamily: DISPLAY_FONT,
    fontSize: tokens.fontSizeBase400,
    fontWeight: 700,
    letterSpacing: "-0.01em",
  },
  headerAction: {
    marginLeft: "auto",
  },
  logo: {
    width: "24px",
    height: "24px",
    borderRadius: tokens.borderRadiusMedium,
    border: `2px solid ${tokens.colorNeutralForeground1}`,
    display: "grid",
    placeItems: "center",
  },
  logoMark: {
    width: "8px",
    height: "8px",
    backgroundColor: tokens.colorBrandBackground,
  },
  chatWrap: {
    flexGrow: 1,
    minHeight: 0,
  },
  body: {
    flexGrow: 1,
    overflowY: "auto",
    padding: "12px",
  },
  intro: {
    display: "flex",
    flexDirection: "column",
    rowGap: "4px",
    margin: "4px 2px 14px",
  },
  greeting: {
    fontSize: tokens.fontSizeBase500,
    fontWeight: tokens.fontWeightSemibold,
  },
  subtle: {
    color: tokens.colorNeutralForeground3,
  },
  toolList: {
    display: "flex",
    flexDirection: "column",
    rowGap: "8px",
    marginBottom: "16px",
  },
  toolButton: {
    display: "flex",
    alignItems: "center",
    columnGap: "12px",
    width: "100%",
    padding: "12px",
    textAlign: "left",
    border: `1px solid transparent`,
    borderRadius: tokens.borderRadiusLarge,
    backgroundColor: tokens.colorNeutralBackground1,
    boxShadow: tokens.shadow2,
    color: tokens.colorNeutralForeground1,
    cursor: "pointer",
    fontFamily: "inherit",
    ":hover": {
      backgroundColor: tokens.colorNeutralBackground1Hover,
      border: `1px solid ${tokens.colorBrandStroke2}`,
    },
    ":focus-visible": {
      outline: `2px solid ${tokens.colorStrokeFocus2}`,
    },
  },
  toolIcon: {
    flexShrink: 0,
    width: "40px",
    height: "40px",
    borderRadius: tokens.borderRadiusMedium,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    backgroundColor: tokens.colorBrandBackground2,
    color: tokens.colorBrandForeground2,
  },
  toolText: {
    display: "flex",
    flexDirection: "column",
    rowGap: "2px",
    flexGrow: 1,
  },
  toolTitle: {
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase300,
  },
  toolDescription: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
    lineHeight: tokens.lineHeightBase200,
  },
  chevron: {
    flexShrink: 0,
    color: tokens.colorNeutralForeground3,
  },
  featured: {
    padding: "16px 14px",
    border: `1px solid ${tokens.colorBrandStroke2}`,
  },
  featuredIcon: {
    width: "44px",
    height: "44px",
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundOnBrand,
  },
  featuredTitle: {
    fontSize: tokens.fontSizeBase400,
  },
  sectionLabel: {
    display: "block",
    margin: "4px 2px 8px",
    fontSize: tokens.fontSizeBase100,
    fontWeight: tokens.fontWeightSemibold,
    letterSpacing: "0.06em",
    textTransform: "uppercase",
    color: tokens.colorNeutralForeground3,
  },
});

type View = "chat" | "tools" | ToolId;

/** Tool screens the assistant can open by name. */
const TOOL_FOR: Record<ToolName, View> = {
  home: "tools",
  formulas: "formulas",
  reports: "reports",
  merge: "merge",
  charts: "charts",
  clean: "clean",
};

const App: React.FC = () => {
  const styles = useStyles();
  // The chat is home; the tool screens are one tap away for anything with lots of choices.
  const [view, setView] = React.useState<View>("chat");
  const [cameFromChat, setCameFromChat] = React.useState(false);
  const [params, setParams] = React.useState<ToolParams | undefined>(undefined);
  // Bumped on every navigation so a tool reopened with new params starts fresh.
  const [visit, setVisit] = React.useState(0);
  const tool = TOOLS.find((item) => item.id === view) ?? null;
  const bodyRef = React.useRef<HTMLDivElement | null>(null);

  const navigate = React.useCallback((next: ToolId | null, nextParams?: ToolParams) => {
    setView(next ?? "tools");
    setParams(nextParams);
    setCameFromChat(false);
    setVisit((count) => count + 1);
  }, []);

  const openFromChat = React.useCallback((name: ToolName) => {
    setView(TOOL_FOR[name]);
    setParams(undefined);
    setCameFromChat(true);
    setVisit((count) => count + 1);
  }, []);

  // Each tool starts at the top, not wherever the previous one was scrolled to.
  React.useEffect(() => {
    bodyRef.current?.scrollTo?.({ top: 0 });
  }, [visit]);

  const featured = TOOLS.filter((item) => item.featured);
  const others = TOOLS.filter((item) => !item.featured);
  const back = () => {
    if (view === "tools" || cameFromChat) setView("chat");
    else setView("tools");
  };

  const toolButton = (item: Tool) => (
    <button
      key={item.id}
      type="button"
      className={mergeClasses(styles.toolButton, item.featured && styles.featured)}
      onClick={() => navigate(item.id)}
    >
      <span className={mergeClasses(styles.toolIcon, item.featured && styles.featuredIcon)}>{item.icon}</span>
      <span className={styles.toolText}>
        <span className={mergeClasses(styles.toolTitle, item.featured && styles.featuredTitle)}>{item.title}</span>
        <span className={styles.toolDescription}>{item.description}</span>
      </span>
      <ChevronRight20Regular className={styles.chevron} />
    </button>
  );

  return (
    <div className={styles.root}>
      <header className={styles.header}>
        {view === "chat" ? (
          <>
            <span className={styles.logo} aria-hidden="true">
              <span className={styles.logoMark} />
            </span>
            <Text className={styles.headerTitle}>MEx Automate</Text>
            <Button
              className={styles.headerAction}
              appearance="subtle"
              icon={<Apps20Regular />}
              onClick={() => setView("tools")}
            >
              All tools
            </Button>
          </>
        ) : (
          <>
            <Button
              appearance="subtle"
              icon={<ArrowLeft20Regular />}
              aria-label={view === "tools" || cameFromChat ? "Back to the chat" : "Back to all tools"}
              title={view === "tools" || cameFromChat ? "Back to the chat" : "Back to all tools"}
              onClick={back}
            />
            {tool ? <span className={styles.headerIcon}>{tool.icon}</span> : null}
            <Text className={styles.headerTitle}>{tool ? tool.title : "All tools"}</Text>
          </>
        )}
      </header>

      {/* The chat stays mounted while a tool is open, so the conversation survives the trip. */}
      <div className={styles.chatWrap} hidden={view !== "chat"}>
        <ChatPanel onOpenTool={openFromChat} />
      </div>

      {view !== "chat" ? (
        <main className={styles.body} ref={bodyRef}>
          {tool ? (
            // Keyed so each visit starts fresh and re-reads the workbook.
            <tool.panel key={`${tool.id}-${visit}`} navigate={navigate} params={params} />
          ) : (
            <>
              <div className={styles.intro}>
                <Text className={styles.greeting}>All tools</Text>
                <Text className={styles.subtle}>
                  The same things the chat does, as step-by-step screens, plus tools with more choices.
                </Text>
              </div>

              <nav className={styles.toolList} aria-label="Main tools">
                {featured.map(toolButton)}
              </nav>

              <Text className={styles.sectionLabel}>More tools</Text>
              <nav className={styles.toolList} aria-label="More tools">
                {others.map(toolButton)}
              </nav>

              <Tip>
                Everything runs inside Excel on your computer and nothing is uploaded. For data that
                matters, save a copy of the workbook before making big changes.
              </Tip>
            </>
          )}
        </main>
      ) : null}
    </div>
  );
};

export default App;
