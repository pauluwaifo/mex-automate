import * as React from "react";
import { createRoot, Root } from "react-dom/client";
import { FluentProvider } from "@fluentui/react-components";

import App from "./components/App";
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

function render() {
  root?.render(
    <FluentProvider theme={currentTheme()}>
      <App />
    </FluentProvider>
  );
}

Office.onReady(() => {
  render();
});

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    render();
  });
}
