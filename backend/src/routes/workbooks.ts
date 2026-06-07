import { Router } from "express";
import multer from "multer";
import { randomUUID } from "node:crypto";
import ExcelJS from "exceljs";
import { describeStructure } from "../excel.js";

export const workbooksRouter = Router();

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 25 * 1024 * 1024 }, // 25 MB
});

// In-memory store for v1 scaffolding. Numbers are always re-read from here by the
// deterministic verifier — they never round-trip through model text.
const workbooks = new Map<string, ExcelJS.Workbook>();

export function getWorkbook(fileId: string): ExcelJS.Workbook | undefined {
  return workbooks.get(fileId);
}

/**
 * POST /api/workbooks  (multipart/form-data, field name: "file")
 * Parses the uploaded .xlsx deterministically and returns { fileId, sheets }.
 */
workbooksRouter.post("/", upload.single("file"), async (req, res) => {
  if (!req.file) {
    res.status(400).json({ error: "multipart field `file` is required (.xlsx)" });
    return;
  }
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(req.file.buffer as unknown as ArrayBuffer);

    const fileId = randomUUID();
    workbooks.set(fileId, workbook);

    res.json({
      fileId,
      filename: req.file.originalname,
      sheets: workbook.worksheets.map((ws) => ({
        name: ws.name,
        rowCount: ws.rowCount,
        columnCount: ws.columnCount,
      })),
    });
  } catch (err) {
    res.status(422).json({ error: `failed to parse workbook: ${String(err)}` });
  }
});

/**
 * GET /api/workbooks/:fileId/structure
 * Sheets + used ranges + first-row headers (no bulk values) — feeds the
 * `list_workbook_structure` tool in milestone 3.
 */
workbooksRouter.get("/:fileId/structure", (req, res) => {
  const workbook = workbooks.get(req.params.fileId);
  if (!workbook) {
    res.status(404).json({ error: "unknown fileId" });
    return;
  }
  res.json(describeStructure(workbook));
});
