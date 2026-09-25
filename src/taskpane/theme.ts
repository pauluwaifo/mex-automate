/**
 * The MEx look: the landing page's jade accent, tinted neutral grounds and
 * IBM Plex type, applied to Fluent UI so the add-in matches the website.
 */

import {
  BrandVariants,
  createDarkTheme,
  createLightTheme,
  Theme,
} from "@fluentui/react-components";

/** Jade ramp; 80 is the light-theme brand colour (#0e7a5c), 110 reads well on dark. */
const jade: BrandVariants = {
  10: "#031510",
  20: "#05211a",
  30: "#072e24",
  40: "#093b2e",
  50: "#0a4838",
  60: "#0b5542",
  70: "#0c644e",
  80: "#0e7a5c",
  90: "#1f8a6c",
  100: "#35a07f",
  110: "#35c595",
  120: "#5bd6ac",
  130: "#80e0bf",
  140: "#a6ead2",
  150: "#cbf3e4",
  160: "#e8faf3",
};

const fonts = {
  fontFamilyBase: '"IBM Plex Sans", -apple-system, "Segoe UI", system-ui, sans-serif',
  fontFamilyMonospace: '"IBM Plex Mono", ui-monospace, "Cascadia Mono", Consolas, monospace',
};

export const lightTheme: Theme = {
  ...createLightTheme(jade),
  ...fonts,
  colorNeutralBackground2: "#f3f6f5",
  colorNeutralBackground3: "#eef3f1",
  colorNeutralStroke2: "#dde5e2",
  colorNeutralStroke3: "#e8eeec",
};

export const darkTheme: Theme = {
  ...createDarkTheme(jade),
  ...fonts,
  colorBrandForeground1: jade[110],
  colorBrandForeground2: jade[120],
  colorNeutralBackground1: "#17201d",
  colorNeutralBackground2: "#0d1412",
  colorNeutralBackground3: "#1d2825",
  colorNeutralStroke2: "#2a3733",
  colorNeutralStroke3: "#22302b",
};

/** The display face for headings, matching the website. */
export const DISPLAY_FONT =
  '"Bricolage Grotesque", "Arial Narrow", "Segoe UI", system-ui, sans-serif';
