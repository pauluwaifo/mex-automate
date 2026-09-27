/**
 * The PowerPoint pane: build a deck from a spreadsheet, and rebuild it next
 * month.
 *
 * Deliberately not a chat. In Excel the pane sits beside data the add-in can
 * already read, so a conversation makes sense - there is always something to
 * talk about. Here the add-in can see the slides but not the numbers, so every
 * useful action starts with "which file?". A screen with the file at the top
 * and the deck laid out underneath says that plainly; a chat would only ask the
 * same question in more words.
 */

import * as React from "react";
import {
  Button,
  Input,
  makeStyles,
  Spinner,
  Text,
  tokens,
} from "@fluentui/react-components";
import {
  ArrowSync20Regular,
  ArrowUpload20Regular,
  CheckmarkCircle20Filled,
  Delete20Regular,
  ErrorCircle20Filled,
  SlideAdd20Regular,
} from "@fluentui/react-icons";

import { DeckPreview, planDeckFromGrid } from "../features/deckFromWorkbook";
import {
  buildDeck,
  canRotateShapes,
  findBuiltSlides,
  isDeckSupported,
  refreshDeck,
  removeBuiltSlides,
} from "../features/deckSlides";
import { readWorkbookFile, SUPPORTED_EXTENSIONS } from "../shared/workbookReader";
import type { Grid, OperationResult } from "../shared/types";
import { DISPLAY_FONT } from "../theme";

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    rowGap: "14px",
    padding: "14px",
    height: "100%",
    boxSizing: "border-box",
    overflowY: "auto",
  },
  lead: {
    fontFamily: DISPLAY_FONT,
    fontSize: tokens.fontSizeBase500,
    fontWeight: tokens.fontWeightSemibold,
    lineHeight: tokens.lineHeightBase500,
  },
  hint: { color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200 },
  card: {
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
    padding: "12px",
    display: "flex",
    flexDirection: "column",
    rowGap: "8px",
  },
  fileRow: { display: "flex", alignItems: "center", columnGap: "8px" },
  fileName: { fontWeight: tokens.fontWeightSemibold, overflowWrap: "anywhere" },
  slides: { display: "flex", flexDirection: "column", rowGap: "4px", margin: 0, padding: 0 },
  slide: {
    display: "flex",
    alignItems: "center",
    columnGap: "8px",
    padding: "6px 8px",
    borderRadius: tokens.borderRadiusSmall,
    backgroundColor: tokens.colorNeutralBackground2,
    fontSize: tokens.fontSizeBase200,
  },
  number: {
    flexShrink: 0,
    width: "18px",
    height: "18px",
    borderRadius: "50%",
    backgroundColor: tokens.colorBrandBackground2,
    color: tokens.colorBrandForeground1,
    display: "grid",
    placeItems: "center",
    fontSize: "10px",
    fontWeight: tokens.fontWeightSemibold,
  },
  actions: { display: "flex", columnGap: "8px", flexWrap: "wrap" },
  result: { display: "flex", columnGap: "8px", alignItems: "flex-start" },
  ok: { color: tokens.colorPaletteGreenForeground1, flexShrink: 0 },
  bad: { color: tokens.colorPaletteRedForeground1, flexShrink: 0 },
  details: {
    margin: "4px 0 0",
    paddingLeft: "16px",
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
  },
  hidden: { display: "none" },
});

interface Picked {
  label: string;
  /** Kept so the deck can be re-planned when the title or options change. */
  grid: Grid;
  formats?: string[][];
  preview: DeckPreview;
}

const DeckPanel: React.FC = () => {
  const styles = useStyles();
  const fileInput = React.useRef<HTMLInputElement | null>(null);

  const [picked, setPicked] = React.useState<Picked | null>(null);
  const [title, setTitle] = React.useState("");
  const [busy, setBusy] = React.useState<string | null>(null);
  const [result, setResult] = React.useState<OperationResult | null>(null);
  const [builtSlides, setBuiltSlides] = React.useState<number>(0);

  const supported = isDeckSupported();

  const countBuilt = React.useCallback(async () => {
    setBuiltSlides((await findBuiltSlides()).length);
  }, []);

  React.useEffect(() => {
    void countBuilt();
  }, [countBuilt]);

  const take = async (file: File | undefined) => {
    if (!file) return;
    setBusy("Reading the file...");
    setResult(null);
    try {
      const workbook = await readWorkbookFile(file);
      const usable = workbook.sheets
        .filter((sheet) => sheet.grid.length >= 2)
        .sort((a, b) => b.grid.length - a.grid.length);
      if (usable.length === 0) {
        setResult({ ok: false, message: `"${file.name}" has no sheet with data in it.` });
        setPicked(null);
        return;
      }
      const sheet = usable[0];
      const label = workbook.sheets.length > 1 ? `${file.name} · ${sheet.name}` : file.name;
      const preview = planDeckFromGrid(
        { label, grid: sheet.grid, formats: sheet.numberFormats },
        { title, canRotate: canRotateShapes() }
      );
      if (!preview) {
        setResult({
          ok: false,
          message: `I couldn't find anything in "${file.name}" worth charting. It needs a table with a heading row, a column of numbers, and something to group them by.`,
        });
        setPicked(null);
        return;
      }
      setPicked({ label, grid: sheet.grid, formats: sheet.numberFormats, preview });
      if (title.trim() === "") setTitle(preview.plan.title);
    } catch (error) {
      setResult({
        ok: false,
        message: `I couldn't read that file: ${error instanceof Error ? error.message : "unknown problem"}.`,
      });
    } finally {
      setBusy(null);
    }
  };

  /** Re-plans from the grid, so an edited title reaches the slides. */
  const planNow = (): DeckPreview | null => {
    if (!picked) return null;
    if (title.trim() === "" || title.trim() === picked.preview.plan.title) return picked.preview;
    return (
      planDeckFromGrid(
        { label: picked.label, grid: picked.grid, formats: picked.formats },
        { title, canRotate: canRotateShapes() }
      ) ?? picked.preview
    );
  };

  const build = async () => {
    const preview = planNow();
    if (!preview) return;
    setBusy("Building the slides...");
    setResult(await buildDeck(preview.plan));
    setBusy(null);
    void countBuilt();
  };

  const refresh = async () => {
    const preview = planNow();
    if (!preview) return;
    setBusy("Redrawing from the new numbers...");
    setResult(await refreshDeck(preview.plan));
    setBusy(null);
    void countBuilt();
  };

  const remove = async () => {
    setBusy("Removing the slides I built...");
    setResult(await removeBuiltSlides());
    setBusy(null);
    void countBuilt();
  };

  return (
    <div className={styles.root}>
      <div>
        <Text as="h1" className={styles.lead}>
          Build the deck from your spreadsheet
        </Text>
        <Text className={styles.hint}>
          Pick the file, and MEx cleans it, works out what is worth showing, and lays out the
          slides. Charts are drawn as ordinary PowerPoint shapes, so you can recolour and move
          anything afterwards. The file is read on this computer and never uploaded.
        </Text>
      </div>

      {!supported ? (
        <div className={styles.card}>
          <Text weight="semibold">This version of PowerPoint can&apos;t do it</Text>
          <Text className={styles.hint}>
            Adding shapes from an add-in needs PowerPoint on the web, or Microsoft 365 on Windows
            from version 2208. Nothing here will work until then.
          </Text>
        </div>
      ) : null}

      <input
        ref={fileInput}
        type="file"
        className={styles.hidden}
        accept={SUPPORTED_EXTENSIONS.join(",")}
        onChange={(event) => {
          void take(event.target.files?.[0]);
          event.target.value = "";
        }}
      />

      <div className={styles.card}>
        <div className={styles.fileRow}>
          <Button
            appearance={picked ? "outline" : "primary"}
            icon={<ArrowUpload20Regular />}
            disabled={!supported || busy !== null}
            onClick={() => fileInput.current?.click()}
          >
            {picked ? "Choose a different file" : "Choose a spreadsheet"}
          </Button>
        </div>
        {picked ? (
          <>
            <Text className={styles.fileName}>{picked.label}</Text>
            <Text className={styles.hint}>
              {picked.preview.rowCount.toLocaleString()} rows
              {picked.preview.problemsFixed > 0
                ? ` · ${picked.preview.problemsFixed.toLocaleString()} problems smoothed over for the charts`
                : ""}
            </Text>
          </>
        ) : (
          <Text className={styles.hint}>
            Excel or CSV ({SUPPORTED_EXTENSIONS.join(", ")}). A messy export is fine — titles,
            totals and numbers stored as text are handled.
          </Text>
        )}
      </div>

      {picked ? (
        <div className={styles.card}>
          <Text weight="semibold">Deck title</Text>
          <Input value={title} onChange={(_event, data) => setTitle(data.value)} />
          <Text weight="semibold">
            {picked.preview.plan.slides.length} slides, in this order
          </Text>
          <ul className={styles.slides}>
            {picked.preview.plan.slides.map((slide, index) => (
              <li key={`${slide.kind}-${index}`} className={styles.slide}>
                <span className={styles.number}>{index + 1}</span>
                <span>{slide.title}</span>
              </li>
            ))}
          </ul>
          {picked.preview.chartsLeftOut > 0 ? (
            <Text className={styles.hint}>
              {picked.preview.chartsLeftOut} more chart
              {picked.preview.chartsLeftOut === 1 ? " was" : "s were"} suggested and left out to
              keep the deck short.
            </Text>
          ) : null}
        </div>
      ) : null}

      <div className={styles.actions}>
        <Button
          appearance="primary"
          icon={<SlideAdd20Regular />}
          disabled={!picked || !supported || busy !== null}
          onClick={() => void build()}
        >
          Add the slides
        </Button>
        <Button
          icon={<ArrowSync20Regular />}
          disabled={!picked || builtSlides === 0 || busy !== null}
          onClick={() => void refresh()}
        >
          Refresh {builtSlides > 0 ? `${builtSlides} slides` : ""}
        </Button>
        {builtSlides > 0 ? (
          <Button
            icon={<Delete20Regular />}
            disabled={busy !== null}
            onClick={() => void remove()}
          >
            Remove them
          </Button>
        ) : null}
      </div>

      {builtSlides > 0 && !picked ? (
        <Text className={styles.hint}>
          This deck already has {builtSlides} slides I built. Pick the new file and choose Refresh
          to redraw them — they remember what they show, so nothing has to be relinked.
        </Text>
      ) : null}

      {busy ? <Spinner size="tiny" label={busy} labelPosition="after" /> : null}

      {result ? (
        <div className={styles.result}>
          {result.ok ? (
            <CheckmarkCircle20Filled className={styles.ok} />
          ) : (
            <ErrorCircle20Filled className={styles.bad} />
          )}
          <div>
            <Text>{result.message}</Text>
            {result.details && result.details.length > 0 ? (
              <ul className={styles.details}>
                {result.details.map((detail, index) => (
                  <li key={`${index}-${detail}`}>{detail}</li>
                ))}
              </ul>
            ) : null}
          </div>
        </div>
      ) : null}
    </div>
  );
};

export default DeckPanel;
