/**
 * Branded-template definitions — the single source of truth shared by the
 * pptxgenjs generator (`scripts/build-templates.ts`), which creates each
 * placeholder as a shape whose name (OOXML `objectName`) is `shape`, and the
 * template store (`templates.ts`), which serves this inventory to the agent so
 * it can plan `apply_slide_content` fills against stable shape names.
 */

export interface PlaceholderDef {
  /** Shape name inside the .pptx — what Office.js reads back as `shape.name`. */
  shape: string;
  description: string;
  /** `number` placeholders must be filled from workbook cells with a source ref. */
  kind: "text" | "number";
}

export interface TemplateSlideDef {
  name: string;
  description: string;
  placeholders: PlaceholderDef[];
}

export interface TemplateDef {
  id: string;
  name: string;
  description: string;
  /** Filename within `backend/templates/`. */
  file: string;
  slides: TemplateSlideDef[];
}

export const TEMPLATES: TemplateDef[] = [
  {
    id: "quarterly-review",
    name: "Quarterly Business Review",
    description:
      "Two-slide branded deck (navy/gold): a title slide and a four-KPI summary slide. " +
      "Each KPI has a label placeholder and a value placeholder.",
    file: "quarterly-review.pptx",
    slides: [
      {
        name: "Title",
        description: "Cover slide with presentation title, subtitle, and date.",
        placeholders: [
          { shape: "title_text", description: "Main presentation title", kind: "text" },
          { shape: "subtitle_text", description: "Subtitle, e.g. client or engagement name", kind: "text" },
          { shape: "date_text", description: "Presentation date or reporting period", kind: "text" },
        ],
      },
      {
        name: "KPI Summary",
        description:
          "Headline metrics: four KPI cards, each with a short label and a prominent value, " +
          "plus a source note at the bottom.",
        placeholders: [
          { shape: "heading_text", description: "Slide heading, e.g. 'FY24 Financial Highlights'", kind: "text" },
          { shape: "kpi1_label", description: "KPI 1 label, e.g. 'Revenue'", kind: "text" },
          { shape: "kpi1_value", description: "KPI 1 value (from a workbook cell)", kind: "number" },
          { shape: "kpi2_label", description: "KPI 2 label", kind: "text" },
          { shape: "kpi2_value", description: "KPI 2 value (from a workbook cell)", kind: "number" },
          { shape: "kpi3_label", description: "KPI 3 label", kind: "text" },
          { shape: "kpi3_value", description: "KPI 3 value (from a workbook cell)", kind: "number" },
          { shape: "kpi4_label", description: "KPI 4 label", kind: "text" },
          { shape: "kpi4_value", description: "KPI 4 value (from a workbook cell)", kind: "number" },
          { shape: "source_note", description: "Data-source note, e.g. workbook name and sheet", kind: "text" },
        ],
      },
    ],
  },
];
