/**
 * The PowerPoint side: put the planned slides into the open presentation, and
 * redraw them next month from the tags they carry.
 *
 * What makes refresh work is what each slide remembers. A slide does not store
 * a path to a workbook - that is the thing that breaks when the file is renamed
 * or moved. It stores *what it shows*: which chart of the plan, and where the
 * numbers came from, as a tag. Refresh reads those tags, asks for the file
 * again, recomputes, and replaces only the shapes MEx drew. Anything added by
 * hand on those slides stays.
 *
 * Every shape is named with the MEx_ prefix for exactly that reason: it is how
 * a redraw knows what is ours to remove.
 */

import { describeExcelError, fail, ok, plural } from "../shared/excelHelpers";
import type { OperationResult } from "../shared/types";
import {
  DeckPlan,
  parseTag,
  serializeTag,
  ShapeSpec,
  SHAPE_PREFIX,
  SlideSpec,
  SLIDE_TAG_KEY,
  SlideTag,
} from "./deck";

/** The set that gives shapes, text frames, fills and fonts. Everything needs this. */
export function isDeckSupported(): boolean {
  try {
    return Office.context.requirements.isSetSupported("PowerPointApi", "1.4");
  } catch {
    return false;
  }
}

/** Rotation arrived in 1.10; without it a sloped trend line cannot be drawn. */
export function canRotateShapes(): boolean {
  try {
    return Office.context.requirements.isSetSupported("PowerPointApi", "1.10");
  } catch {
    return false;
  }
}

function alignmentOf(
  align: "left" | "center" | "right" | undefined
): PowerPoint.ParagraphHorizontalAlignment {
  if (align === "center") return PowerPoint.ParagraphHorizontalAlignment.center;
  if (align === "right") return PowerPoint.ParagraphHorizontalAlignment.right;
  return PowerPoint.ParagraphHorizontalAlignment.left;
}

/** Draws one planned shape onto a slide. */
function draw(slide: PowerPoint.Slide, spec: ShapeSpec): void {
  if (spec.kind === "text") {
    const shape = slide.shapes.addTextBox(spec.text, {
      left: spec.box.left,
      top: spec.box.top,
      width: spec.box.width,
      height: spec.box.height,
    });
    shape.name = spec.name;
    // A text box with a fill or outline looks like a box; these should read as
    // text sitting on the slide.
    shape.fill.clear();
    shape.lineFormat.visible = false;
    const font = shape.textFrame.textRange.font;
    font.size = spec.style.size;
    font.bold = Boolean(spec.style.bold);
    font.color = spec.style.color;
    shape.textFrame.textRange.paragraphFormat.horizontalAlignment = alignmentOf(spec.style.align);
    return;
  }

  if (spec.kind === "rect") {
    const shape = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
      left: spec.box.left,
      top: spec.box.top,
      width: spec.box.width,
      height: spec.box.height,
    });
    shape.name = spec.name;
    shape.fill.setSolidColor(spec.fill);
    if (spec.transparency !== undefined) shape.fill.transparency = spec.transparency;
    shape.lineFormat.visible = false;
    return;
  }

  if (spec.kind === "line") {
    // Flat rules only: a connector fills its bounding box from top-left to
    // bottom-right, which is correct for a horizontal line and wrong for a
    // rising one.
    const left = Math.min(spec.from.x, spec.to.x);
    const top = Math.min(spec.from.y, spec.to.y);
    const shape = slide.shapes.addLine(PowerPoint.ConnectorType.straight, {
      left,
      top,
      width: Math.abs(spec.to.x - spec.from.x),
      height: Math.abs(spec.to.y - spec.from.y),
    });
    shape.name = spec.name;
    shape.lineFormat.color = spec.color;
    shape.lineFormat.weight = spec.weight;
    return;
  }

  // A sloped segment: a thin rectangle turned to match the slope. The planner
  // only asks for these where rotation is available.
  const dx = spec.to.x - spec.from.x;
  const dy = spec.to.y - spec.from.y;
  const length = Math.sqrt(dx * dx + dy * dy);
  const midX = (spec.from.x + spec.to.x) / 2;
  const midY = (spec.from.y + spec.to.y) / 2;
  const shape = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
    left: midX - length / 2,
    top: midY - spec.weight / 2,
    width: Math.max(1, length),
    height: spec.weight,
  });
  shape.name = spec.name;
  shape.fill.setSolidColor(spec.color);
  shape.lineFormat.visible = false;
  shape.rotation = (Math.atan2(dy, dx) * 180) / Math.PI;
}

/** Adds one slide and everything on it. */
async function addSlide(
  context: PowerPoint.RequestContext,
  spec: SlideSpec,
  at: number
): Promise<void> {
  context.presentation.slides.add();
  await context.sync();

  const slide = context.presentation.slides.getItemAt(at);
  for (const shape of spec.shapes) draw(slide, shape);
  slide.tags.add(SLIDE_TAG_KEY, serializeTag(spec.tag));
  await context.sync();
}

export interface DeckOutcome extends OperationResult {
  slidesAdded?: number;
}

/**
 * Builds the deck at the end of the presentation.
 *
 * Slides are appended rather than replacing what is there: the deck usually has
 * a cover and an agenda somebody wrote, and deleting those would be a rude
 * surprise.
 */
export async function buildDeck(plan: DeckPlan): Promise<DeckOutcome> {
  if (!isDeckSupported()) {
    return fail(
      "This version of PowerPoint can't have shapes added by an add-in. It needs PowerPoint on the web, or Microsoft 365 on Windows from version 2208."
    );
  }
  if (plan.slides.length === 0) return fail("There was nothing to put on a slide.");

  try {
    return await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items/id");
      await context.sync();
      const startedWith = slides.items.length;

      for (let index = 0; index < plan.slides.length; index += 1) {
        // Sequential on purpose: each slide has to exist before it can be drawn on.

        await addSlide(context, plan.slides[index], startedWith + index);
      }

      return {
        ...ok(`Built ${plural(plan.slides.length, "slide")} at the end of this deck.`, [
          plan.subtitle,
          "Every chart is made of ordinary PowerPoint shapes, so you can recolour, move and animate any of it.",
          'Next month, say "refresh" and pick the new file - the slides know what they show.',
        ]),
        slidesAdded: plan.slides.length,
      };
    });
  } catch (error) {
    return fail(describeExcelError(error));
  }
}

export interface TaggedSlide {
  index: number;
  tag: SlideTag;
}

/** The slides MEx built, in the order they appear. */
export async function findBuiltSlides(): Promise<TaggedSlide[]> {
  try {
    return await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items/id");
      await context.sync();

      const tags = slides.items.map((slide) => {
        const tag = slide.tags.getItemOrNullObject(SLIDE_TAG_KEY);
        tag.load("value");
        return tag;
      });
      await context.sync();

      const found: TaggedSlide[] = [];
      tags.forEach((tag, index) => {
        if (tag.isNullObject) return;
        const parsed = parseTag(tag.value);
        if (parsed) found.push({ index, tag: parsed });
      });
      return found;
    });
  } catch {
    return [];
  }
}

/**
 * Redraws the slides MEx built, using a freshly planned deck.
 *
 * Only shapes named with the MEx_ prefix are removed, so a note or an arrow
 * somebody added to a slide survives the refresh. Slides that the new plan has
 * no replacement for are left exactly as they are rather than deleted - if last
 * month had a chart this month's data cannot produce, that is worth seeing, not
 * hiding.
 */
export async function refreshDeck(plan: DeckPlan): Promise<DeckOutcome> {
  if (!isDeckSupported()) {
    return fail("This version of PowerPoint can't have shapes changed by an add-in.");
  }

  const built = await findBuiltSlides();
  if (built.length === 0) {
    return fail(
      "I can't see any slides in this deck that I built. Use /deck to make one, and then refreshing it will work."
    );
  }

  try {
    return await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items/id");
      await context.sync();

      let redrawn = 0;
      const leftAlone: string[] = [];

      for (const target of built) {
        const replacement =
          plan.slides.find(
            (candidate) =>
              candidate.tag.kind === target.tag.kind && candidate.tag.chartId === target.tag.chartId
          ) ?? null;
        if (!replacement) {
          leftAlone.push(target.tag.chartId || target.tag.kind);
          continue;
        }

        const slide = slides.getItemAt(target.index);
        slide.shapes.load("items/name");

        // One sync per slide is the point: each slide's existing shapes have to
        // be read before its own can be removed and redrawn.
        // eslint-disable-next-line office-addins/no-context-sync-in-loop -- deliberate, see above.
        await context.sync();

        // eslint-disable-next-line office-addins/load-object-before-read -- the load is two lines up, on slide.shapes; the rule does not follow it through the collection.
        for (const shape of slide.shapes.items) {
          if (shape.name.startsWith(SHAPE_PREFIX)) shape.delete();
        }
        for (const shape of replacement.shapes) draw(slide, shape);
        slide.tags.add(SLIDE_TAG_KEY, serializeTag(replacement.tag));

        // eslint-disable-next-line office-addins/no-context-sync-in-loop -- deliberate: the deletions must land before the next slide is read.
        await context.sync();
        redrawn += 1;
      }

      const details = [plan.subtitle];
      if (leftAlone.length > 0) {
        details.push(
          `${plural(leftAlone.length, "slide")} left untouched: this month's data doesn't produce ${leftAlone.length === 1 ? "that chart" : "those charts"}.`
        );
      }
      details.push("Anything you added to these slides by hand is still there.");

      return {
        ...ok(`Redrew ${plural(redrawn, "slide")} from the new numbers.`, details),
        slidesAdded: 0,
      };
    });
  } catch (error) {
    return fail(describeExcelError(error));
  }
}

/** Deletes the slides MEx built, for starting again. */
export async function removeBuiltSlides(): Promise<OperationResult> {
  const built = await findBuiltSlides();
  if (built.length === 0) return fail("There are no slides here that I built.");

  try {
    return await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items/id");
      await context.sync();

      // Back to front, so removing one does not shift the index of the next.
      for (const target of [...built].sort((a, b) => b.index - a.index)) {
        slides.getItemAt(target.index).delete();
      }
      await context.sync();
      return ok(`Removed ${plural(built.length, "slide")}. The rest of the deck is untouched.`);
    });
  } catch (error) {
    return fail(describeExcelError(error));
  }
}
