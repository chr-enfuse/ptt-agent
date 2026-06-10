import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { TEMPLATES, TemplateDef } from "./template-defs.js";

/**
 * Template store backing the `list_templates` / `load_template` agent tools.
 * The placeholder inventory goes to the model; the .pptx base64 never does —
 * it travels only in the relayed `insert_template_slides` payload to the pane.
 */

// Works from both src/ (tsx) and dist/ (tsc build) — templates/ is a sibling.
const TEMPLATES_DIR = path.join(path.dirname(fileURLToPath(import.meta.url)), "..", "templates");

export class TemplateToolError extends Error {}

function requireTemplate(templateId: string): TemplateDef {
  const def = TEMPLATES.find((t) => t.id === templateId);
  if (!def) {
    const ids = TEMPLATES.map((t) => t.id).join(", ");
    throw new TemplateToolError(`unknown template_id "${templateId}" — available templates: ${ids}`);
  }
  return def;
}

export function listTemplates() {
  return {
    templates: TEMPLATES.map((t) => ({
      id: t.id,
      name: t.name,
      description: t.description,
      slideCount: t.slides.length,
    })),
  };
}

/** Full placeholder inventory the model plans `apply_slide_content` fills against. */
export function loadTemplate(templateId: string) {
  const { id, name, description, slides } = requireTemplate(templateId);
  return { id, name, description, slides };
}

/** The .pptx itself — relayed to the task pane for insertSlidesFromBase64. */
export function getTemplateBase64(templateId: string): string {
  const def = requireTemplate(templateId);
  try {
    return readFileSync(path.join(TEMPLATES_DIR, def.file)).toString("base64");
  } catch {
    throw new TemplateToolError(
      `template file "${def.file}" is missing — run \`npm run build:templates\` in backend/`,
    );
  }
}
