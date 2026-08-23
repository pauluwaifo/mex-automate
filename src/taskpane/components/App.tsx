import * as React from "react";
import { makeStyles, Tab, TabList, tokens } from "@fluentui/react-components";

import ChartPanel from "./ChartPanel";
import CleanPanel from "./CleanPanel";
import FormulaPanel from "./FormulaPanel";
import MergePanel from "./MergePanel";
import ReportPanel from "./ReportPanel";

type TabId = "clean" | "merge" | "charts" | "formulas" | "reports";

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    height: "100vh",
    backgroundColor: tokens.colorNeutralBackground2,
  },
  tabs: {
    flexShrink: 0,
    backgroundColor: tokens.colorNeutralBackground1,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  body: {
    flexGrow: 1,
    overflowY: "auto",
    padding: "12px",
  },
});

const App: React.FC = () => {
  const styles = useStyles();
  const [tab, setTab] = React.useState<TabId>("clean");

  return (
    <div className={styles.root}>
      <TabList
        className={styles.tabs}
        selectedValue={tab}
        onTabSelect={(_event, data) => setTab(data.value as TabId)}
        size="small"
      >
        <Tab value="clean">Clean</Tab>
        <Tab value="merge">Merge</Tab>
        <Tab value="charts">Charts</Tab>
        <Tab value="formulas">Formulas</Tab>
        <Tab value="reports">Reports</Tab>
      </TabList>

      {/* Each panel keeps its own state, so remounting on tab change is deliberate:
          it re-reads the workbook rather than showing a stale selection. */}
      <div className={styles.body}>
        {tab === "clean" ? <CleanPanel /> : null}
        {tab === "merge" ? <MergePanel /> : null}
        {tab === "charts" ? <ChartPanel /> : null}
        {tab === "formulas" ? <FormulaPanel /> : null}
        {tab === "reports" ? <ReportPanel /> : null}
      </div>
    </div>
  );
};

export default App;
