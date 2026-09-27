import * as React from "react";
import { createRoot, Root } from "react-dom/client";
import { FluentProvider } from "@fluentui/react-components";

import App from "./components/App";
import DeckPanel from "./components/DeckPanel";
import { darkTheme, lightTheme } from "./theme";

const rootElement: HTMLElement | null = document.getElementById("container");
const root: Root | undefined = rootElement ? createRoot(rootElement) : undefined;

/** Follow Office: a dark Office theme gets the dark MEx theme. */
function currentTheme() {
  try {
    const background = Office.context.officeTheme?.bodyBackgroundColor ?? "#ffffff";
    const hex = background.replace("#", "");
    const [r, g, b] = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16));
    // Relative luminance below the midpoint means Office is showing a dark theme.
    return 0.2126 * r + 0.7152 * g + 0.0722 * b < 128 ? darkTheme : lightTheme;
  } catch {
    return lightTheme;
  }
}

/**
 * One bundle serves both add-ins, so the host decides which pane to show.
 *
 * The two have different jobs. In Excel the add-in can read the data beside it,
 * so the pane is a conversation. In PowerPoint it can see the slides but not
 * the numbers, so every useful action starts by asking for the file - which a
 * screen says better than a chat would.
 */
function paneForHost() {
  try {
    if (Office.context.host === Office.HostType.PowerPoint) return <DeckPanel />;
  } catch {
    // An older host that cannot say what it is gets the Excel pane, which is
    // what the manifest would have loaded anyway.
  }
  return <App />;
}

function render() {
  root?.render(<FluentProvider theme={currentTheme()}>{paneForHost()}</FluentProvider>);
}

Office.onReady(() => {
  render();
});

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    render();
  });
}
