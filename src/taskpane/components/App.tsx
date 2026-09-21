import * as React from "react";
import { Button, makeStyles, Text, tokens } from "@fluentui/react-components";
import {
  ArrowLeft20Regular,
  ArrowSync24Regular,
  Broom24Regular,
  ChevronRight20Regular,
  DataPie24Regular,
  MathFormula24Regular,
  TableStackBelow24Regular,
} from "@fluentui/react-icons";

import ChartPanel from "./ChartPanel";
import CleanPanel from "./CleanPanel";
import FormulaPanel from "./FormulaPanel";
import MergePanel from "./MergePanel";
import ReportPanel from "./ReportPanel";
import { Tip } from "./ui";

type ToolId = "clean" | "merge" | "charts" | "formulas" | "reports";

interface Tool {
  id: ToolId;
  title: string;
  description: string;
  icon: React.ReactElement;
  panel: React.FC;
}

/** The home screen, in the order most people reach for them. */
const TOOLS: Tool[] = [
  {
    id: "clean",
    title: "Clean up data",
    description: "Remove duplicates and extra spaces, fix dates and capital letters.",
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
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
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
});

const App: React.FC = () => {
  const styles = useStyles();
  const [toolId, setToolId] = React.useState<ToolId | null>(null);
  const tool = TOOLS.find((item) => item.id === toolId) ?? null;
  const bodyRef = React.useRef<HTMLDivElement | null>(null);

  // Each tool starts at the top, not wherever the previous one was scrolled to.
  React.useEffect(() => {
    bodyRef.current?.scrollTo?.({ top: 0 });
  }, [toolId]);

  return (
    <div className={styles.root}>
      <header className={styles.header}>
        {tool ? (
          <>
            <Button
              appearance="subtle"
              icon={<ArrowLeft20Regular />}
              aria-label="Back to all tools"
              title="Back to all tools"
              onClick={() => setToolId(null)}
            />
            <span className={styles.headerIcon}>{tool.icon}</span>
            <Text className={styles.headerTitle}>{tool.title}</Text>
          </>
        ) : (
          <Text className={styles.headerTitle}>MEx Automate</Text>
        )}
      </header>

      <main className={styles.body} ref={bodyRef}>
        {tool ? (
          // Keyed so each visit starts fresh and re-reads the workbook.
          <tool.panel key={tool.id} />
        ) : (
          <>
            <div className={styles.intro}>
              <Text className={styles.greeting}>What would you like to do?</Text>
              <Text className={styles.subtle}>Pick a tool. You can always come back here.</Text>
            </div>

            <nav className={styles.toolList} aria-label="Tools">
              {TOOLS.map((item) => (
                <button
                  key={item.id}
                  type="button"
                  className={styles.toolButton}
                  onClick={() => setToolId(item.id)}
                >
                  <span className={styles.toolIcon}>{item.icon}</span>
                  <span className={styles.toolText}>
                    <span className={styles.toolTitle}>{item.title}</span>
                    <span className={styles.toolDescription}>{item.description}</span>
                  </span>
                  <ChevronRight20Regular className={styles.chevron} />
                </button>
              ))}
            </nav>

            <Tip>
              Everything runs inside Excel on your computer and nothing is uploaded. For data that
              matters, save a copy of the workbook before making big changes.
            </Tip>
          </>
        )}
      </main>
    </div>
  );
};

export default App;
