/**
 * From pasted notes to a finished deck.
 *
 * The spreadsheet route answers "make the monthly report". This one answers the
 * other half of the job: someone has the content already - in an email, in a
 * document, in their head and now in a text box - and what they do not have is
 * half an hour to make slides out of it.
 *
 * What makes it worth using rather than typing into PowerPoint directly:
 *
 * - The structure is read, not asked for. A heading with a list under it
 *   becomes a slide; a long list becomes two; an indented line becomes a
 *   sub-point; a run of tab-separated lines becomes a table.
 * - Numbers get charted. If a pasted table has a column of labels and a column
 *   of numbers, the deck gets the table *and* a chart of it, drawn with the
 *   same code the spreadsheet decks use.
 * - Nothing is invented. Every word on every slide is a word that was pasted.
 *   No model is involved, so the same paste always gives the same deck, and
 *   there is nothing to fact-check.
 *
 * Everything here is pure and unit tested in tests/deckFromText.test.ts.
 */

import {
  bulletSlide,
  DeckPlan,
  drawChart,
  MARGIN,
  readChartData,
  SLIDE,
  SlideSpec,
  tableSlide,
} from "./deck";
import {
  chartableColumns,
  describeOutline,
  OutlineSlide,
  parseOutline,
  tableToGrid,
} from "./outline";

export interface TextDeckOptions {
  /** Shown on the title slide. Taken from the first heading when not given. */
  title?: string;
  /** Put a title slide in front. On by default. */
  titleSlide?: boolean;
  /** Draw a chart for any pasted table that has numbers in it. On by default. */
  chartTables?: boolean;
  /** A line under each slide's heading, e.g. who it is for. */
  subtitle?: string;
}

export interface TextDeckPreview {
  plan: DeckPlan;
  slideTitles: string[];
  /** What the parser made of the paste, for showing before anything is built. */
  outline: OutlineSlide[];
  summary: string;
  chartsAdded: number;
}

const BODY_TOP = 130;
const FOOTNOTE = 40;

/**
 * Builds the chart slide that goes with a pasted table, when its numbers allow
 * one. A table of text only gets the table.
 */
function chartSlideFor(slide: OutlineSlide, index: number): SlideSpec | null {
  if (!slide.table) return null;
  const columns = chartableColumns(slide.table);
  if (!columns) return null;

  const data = readChartData(tableToGrid(slide.table, columns));
  if (!data || data.categories.length < 2) return null;

  const prefix = `MEx_txtchart_${index}`;
  const shapes = [
    {
      kind: "text" as const,
      name: `${prefix}_title`,
      box: { left: MARGIN, top: 48, width: SLIDE.width - MARGIN * 2, height: 44 },
      text: slide.title,
      style: { size: 28, bold: true, color: "#12201C", align: "left" as const },
    },
    ...drawChart(
      data,
      // More than one number column per row compares like with like, which a
      // grouped column chart shows and a share cannot.
      data.series.length > 1 ? "column" : "bar",
      {
        left: MARGIN,
        top: BODY_TOP,
        width: SLIDE.width - MARGIN * 2,
        height: SLIDE.height - BODY_TOP - FOOTNOTE,
      },
      prefix
    ),
  ];

  return {
    kind: "chart",
    title: `${slide.title} — chart`,
    tag: { chartId: `text-${index}`, kind: "chart", source: "pasted notes" },
    shapes,
  };
}

/** The deck's own title: the first heading, or whatever the caller said. */
function deckTitle(outline: readonly OutlineSlide[], given?: string): string {
  const trimmed = given?.trim();
  if (trimmed) return trimmed;
  return outline[0]?.title ?? "Notes";
}

/**
 * Reads pasted text and plans the whole deck. Returns null when the text holds
 * nothing that could go on a slide.
 */
export function planDeckFromText(
  text: string,
  options: TextDeckOptions = {}
): TextDeckPreview | null {
  const outline = parseOutline(text);
  if (outline.length === 0) return null;

  const { titleSlide: wantTitle = true, chartTables = true } = options;
  const title = deckTitle(outline, options.title);
  const subtitle = options.subtitle ?? null;

  const slides: SlideSpec[] = [];
  let chartsAdded = 0;

  if (wantTitle) {
    slides.push({
      kind: "title",
      title,
      tag: { chartId: "", kind: "title", source: "pasted notes" },
      shapes: [
        {
          kind: "rect",
          name: "MEx_txttitle_rule",
          box: { left: MARGIN, top: 214, width: 96, height: 6 },
          fill: "#0E7A5C",
        },
        {
          kind: "text",
          name: "MEx_txttitle_headline",
          box: { left: MARGIN, top: 238, width: SLIDE.width - MARGIN * 2, height: 80 },
          text: title,
          style: { size: 44, bold: true, color: "#12201C", align: "left" },
        },
        ...(subtitle
          ? [
              {
                kind: "text" as const,
                name: "MEx_txttitle_sub",
                box: { left: MARGIN, top: 322, width: SLIDE.width - MARGIN * 2, height: 30 },
                text: subtitle,
                style: { size: 15, color: "#4B5C57", align: "left" as const },
              },
            ]
          : []),
      ],
    });
  }

  outline.forEach((slide, index) => {
    // A block can carry both: the points that were written about it, and the
    // table that was pasted under them. They get a slide each rather than
    // being squeezed together.
    if (slide.bullets.length > 0) {
      slides.push(
        bulletSlide(slide.title, slide.bullets, {
          prefix: `MEx_pts_${index}`,
          subtitle,
          continued: slide.continued,
        })
      );
    }

    if (slide.table) {
      slides.push(tableSlide(slide.title, slide.table, { prefix: `MEx_tbl_${index}`, subtitle }));
      if (chartTables) {
        const chart = chartSlideFor(slide, index);
        if (chart) {
          slides.push(chart);
          chartsAdded += 1;
        }
      }
    }

    // A heading with neither points nor a table is a section divider.
    if (slide.bullets.length === 0 && !slide.table) {
      slides.push(bulletSlide(slide.title, [], { prefix: `MEx_sec_${index}`, subtitle }));
    }
  });

  return {
    plan: { title, subtitle: subtitle ?? "", slides },
    slideTitles: slides.map((slide) => slide.title),
    outline,
    summary: describeOutline(outline),
    chartsAdded,
  };
}
