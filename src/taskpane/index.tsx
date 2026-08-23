import * as React from "react";
import { createRoot, Root } from "react-dom/client";
import { FluentProvider, webDarkTheme, webLightTheme } from "@fluentui/react-components";

import App from "./components/App";

const rootElement: HTMLElement | null = document.getElementById("container");
const root: Root | undefined = rootElement ? createRoot(rootElement) : undefined;

/** Follow whatever theme Office is using, falling back to light. */
function currentTheme() {
  try {
    return Office.context.officeTheme?.bodyBackgroundColor === "#000000" ? webDarkTheme : webLightTheme;
  } catch {
    return webLightTheme;
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
