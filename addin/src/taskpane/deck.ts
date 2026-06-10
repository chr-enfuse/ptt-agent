/* global PowerPoint Office */

/**
 * Office.js executors for the client-side agent tools. The backend tool-use
 * loop relays these calls to the task pane over SSE (`client_tool` events);
 * Chat.tsx dispatches them here and POSTs the outcome back.
 */

export interface DeckShape {
  name: string;
  /** Current text, or null when the shape has no readable text frame. */
  text: string | null;
}

export interface DeckSlide {
  id: string;
  /** 0-based position in the deck. */
  index: number;
  shapes: DeckShape[];
}

export interface FillSpec {
  shape_name: string;
  text: string;
  /** Provenance ref ({file_id, sheet, cell}) — recorded for the M5 verifier. */
  source?: unknown;
}

export interface FillResult {
  shape_name: string;
  ok: boolean;
  error?: string;
}

function ensureApiSupport(): void {
  if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.4")) {
    throw new Error(
      "this PowerPoint build does not support PowerPointApi 1.4 — update Microsoft 365 to use the deck tools"
    );
  }
}

/** Slides, shape names, and current shape text for the open deck. */
export async function getDeckState(): Promise<{ slides: DeckSlide[] }> {
  ensureApiSupport();
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id,items/shapes/items/name,items/shapes/items/textFrame/textRange/text");
    await context.sync();

    return {
      slides: slides.items.map((slide, index) => ({
        id: slide.id,
        index,
        shapes: slide.shapes.items.map((shape) => {
          let text: string | null = null;
          try {
            text = shape.textFrame.textRange.text;
          } catch {
            // Shape without a text frame (picture, line, …) — name only.
          }
          return { name: shape.name, text };
        }),
      })),
    };
  });
}

/** Append a template's slides to the deck; returns the newly inserted slides. */
export async function insertTemplateSlides(
  base64: string
): Promise<{ insertedSlides: { id: string; index: number }[]; totalSlides: number }> {
  ensureApiSupport();
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();

    const before = new Set(slides.items.map((s) => s.id));
    const last = slides.items[slides.items.length - 1];
    // Without targetSlideId Office inserts at the BEGINNING — anchor on the last
    // slide to append. No return value, so new slides are found by diffing ids.
    context.presentation.insertSlidesFromBase64(base64, {
      formatting: PowerPoint.InsertSlideFormatting.keepSourceFormatting,
      ...(last ? { targetSlideId: last.id } : {}),
    });
    await context.sync();

    slides.load("items/id");
    await context.sync();

    const insertedSlides = slides.items
      .map((s, index) => ({ id: s.id, index }))
      .filter((s) => !before.has(s.id));
    return { insertedSlides, totalSlides: slides.items.length };
  });
}

/** Write text into named placeholder shapes on one slide; per-fill outcome. */
export async function applySlideContent(
  slideId: string,
  fills: FillSpec[]
): Promise<{ slideId: string; results: FillResult[] }> {
  ensureApiSupport();
  return PowerPoint.run(async (context) => {
    const slide = context.presentation.slides.getItemOrNullObject(slideId);
    await context.sync();
    if (slide.isNullObject) {
      throw new Error(`no slide with id "${slideId}" — call get_deck_state for current slide ids`);
    }

    const shapes = slide.shapes;
    shapes.load("items/name");
    await context.sync();

    const shapesByName = new Map(shapes.items.map((s) => [s.name, s]));
    const results: FillResult[] = [];
    for (const fill of fills) {
      const shape = shapesByName.get(fill.shape_name);
      if (!shape) {
        results.push({
          shape_name: fill.shape_name,
          ok: false,
          error: "no shape with this name on the slide",
        });
        continue;
      }
      shape.textFrame.textRange.text = fill.text;
      results.push({ shape_name: fill.shape_name, ok: true });
    }
    await context.sync();
    return { slideId, results };
  });
}

/** Route a relayed client tool call to its Office.js executor. */
export async function executeDeckTool(
  name: string,
  payload: Record<string, unknown>
): Promise<unknown> {
  switch (name) {
    case "get_deck_state":
      return getDeckState();
    case "insert_template_slides":
      return insertTemplateSlides(String(payload.base64 ?? ""));
    case "apply_slide_content":
      return applySlideContent(String(payload.slide_id ?? ""), (payload.fills as FillSpec[]) ?? []);
    default:
      throw new Error(`unknown client tool "${name}"`);
  }
}
