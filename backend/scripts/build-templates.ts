/**
 * Renders the branded .pptx templates into `backend/templates/` from the
 * definitions in `src/template-defs.ts`. Every placeholder is created with
 * `objectName` set to its `shape` name so Office.js can find it by `shape.name`.
 *
 * Run with: npm run build:templates   (the generated .pptx files are committed)
 */
import PptxGenJSImport from "pptxgenjs";
// tsx CJS/ESM interop: the constructor can land on `.default` depending on loader.
const PptxGenJS =
  (PptxGenJSImport as unknown as { default?: typeof PptxGenJSImport }).default ?? PptxGenJSImport;
import { mkdirSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { TEMPLATES } from "../src/template-defs.js";

const TEMPLATES_DIR = path.join(path.dirname(fileURLToPath(import.meta.url)), "..", "templates");

// Brand palette (generic "navy/gold" house style).
const NAVY = "1F3864";
const NAVY_LIGHT = "2E5395";
const GOLD = "C9A227";
const GRAY_TEXT = "595959";
const CARD_BG = "F2F4F8";
const FONT = "Segoe UI";

async function buildQuarterlyReview(file: string) {
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_16x9"; // 13.33in × 7.5in

  // ---- Slide 1: Title ----------------------------------------------------
  const title = pptx.addSlide();
  title.background = { color: NAVY };
  // Static brand accents.
  title.addShape("rect", { x: 0, y: 0, w: 13.33, h: 0.25, fill: { color: GOLD } });
  title.addShape("rect", { x: 0.9, y: 2.5, w: 2.2, h: 0.07, fill: { color: GOLD } });
  title.addText("Presentation Title", {
    objectName: "title_text",
    x: 0.9, y: 2.7, w: 11.5, h: 1.4,
    fontFace: FONT, fontSize: 44, bold: true, color: "FFFFFF", align: "left",
  });
  title.addText("Client / engagement subtitle", {
    objectName: "subtitle_text",
    x: 0.9, y: 4.1, w: 11.5, h: 0.8,
    fontFace: FONT, fontSize: 22, color: "D6DCE5", align: "left",
  });
  title.addText("Date", {
    objectName: "date_text",
    x: 0.9, y: 6.5, w: 6, h: 0.5,
    fontFace: FONT, fontSize: 14, color: GOLD, align: "left",
  });

  // ---- Slide 2: KPI Summary ----------------------------------------------
  const kpi = pptx.addSlide();
  kpi.background = { color: "FFFFFF" };
  kpi.addShape("rect", { x: 0, y: 0, w: 13.33, h: 1.1, fill: { color: NAVY } });
  kpi.addShape("rect", { x: 0, y: 1.1, w: 13.33, h: 0.06, fill: { color: GOLD } });
  kpi.addText("Financial Highlights", {
    objectName: "heading_text",
    x: 0.6, y: 0.2, w: 12, h: 0.7,
    fontFace: FONT, fontSize: 26, bold: true, color: "FFFFFF", align: "left",
  });

  // Four KPI cards in a row.
  const cardW = 2.9;
  const gap = 0.25;
  const rowX = (13.33 - (4 * cardW + 3 * gap)) / 2;
  for (let i = 1; i <= 4; i++) {
    const x = rowX + (i - 1) * (cardW + gap);
    kpi.addShape("roundRect", {
      x, y: 2.2, w: cardW, h: 2.6,
      fill: { color: CARD_BG }, line: { color: NAVY_LIGHT, width: 0.75 }, rectRadius: 0.08,
    });
    kpi.addText(`KPI ${i} label`, {
      objectName: `kpi${i}_label`,
      x: x + 0.15, y: 2.5, w: cardW - 0.3, h: 0.5,
      fontFace: FONT, fontSize: 14, bold: true, color: GRAY_TEXT, align: "center",
    });
    kpi.addText("—", {
      objectName: `kpi${i}_value`,
      x: x + 0.15, y: 3.2, w: cardW - 0.3, h: 1.1,
      fontFace: FONT, fontSize: 30, bold: true, color: NAVY, align: "center",
    });
  }

  kpi.addText("Source: workbook reference", {
    objectName: "source_note",
    x: 0.6, y: 6.9, w: 12, h: 0.4,
    fontFace: FONT, fontSize: 10, italic: true, color: GRAY_TEXT, align: "left",
  });

  await pptx.writeFile({ fileName: path.join(TEMPLATES_DIR, file) });
}

const builders: Record<string, (file: string) => Promise<void>> = {
  "quarterly-review": buildQuarterlyReview,
};

mkdirSync(TEMPLATES_DIR, { recursive: true });
for (const def of TEMPLATES) {
  const build = builders[def.id];
  if (!build) throw new Error(`no builder for template "${def.id}"`);
  await build(def.file);
  console.log(`built templates/${def.file} (${def.slides.length} slides)`);
}
