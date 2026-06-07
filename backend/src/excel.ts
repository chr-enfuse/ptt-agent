import ExcelJS from "exceljs";
import { getWorkbook } from "./routes/workbooks.js";

/**
 * Deterministic Excel readers backing the `list_workbook_structure` and
 * `read_excel_range` agent tools. Numbers flow from here as structured data —
 * the model selects cells but never re-types their values.
 */

/** Hard cap per read_excel_range call so one tool result can't flood the context. */
export const MAX_CELLS_PER_READ = 400;

export class ExcelToolError extends Error {}

export interface CellData {
  addr: string;
  /** Exact stored value (for formula cells, the cached computed result). */
  raw: unknown;
  formula?: string;
  numberFormat: string | null;
  /** The value as Excel would display it (best-effort string form). */
  text: string;
}

export function describeStructure(workbook: ExcelJS.Workbook) {
  return {
    sheets: workbook.worksheets.map((ws) => {
      const headers: (string | null)[] = [];
      ws.getRow(1).eachCell({ includeEmpty: false }, (cell) => {
        headers.push(cell.text || null);
      });
      return {
        name: ws.name,
        rowCount: ws.rowCount,
        columnCount: ws.columnCount,
        headers,
      };
    }),
  };
}

function requireWorkbook(fileId: string): ExcelJS.Workbook {
  const workbook = getWorkbook(fileId);
  if (!workbook) {
    throw new ExcelToolError(
      `unknown file_id "${fileId}" — the workbook may not have been uploaded in this session`,
    );
  }
  return workbook;
}

export function listWorkbookStructure(fileId: string) {
  return describeStructure(requireWorkbook(fileId));
}

/** Parse "B7" or "A1:C10" (case-insensitive) into 1-based bounds. */
function parseA1Range(range: string): { top: number; left: number; bottom: number; right: number } {
  const m = /^([A-Za-z]{1,3})([0-9]+)(?::([A-Za-z]{1,3})([0-9]+))?$/.exec(range.trim());
  if (!m) {
    throw new ExcelToolError(`invalid range "${range}" — use A1 notation like "B7" or "A1:C10"`);
  }
  const colToNum = (col: string) =>
    col.toUpperCase().split("").reduce((n, ch) => n * 26 + (ch.charCodeAt(0) - 64), 0);
  const left = colToNum(m[1]!);
  const top = Number(m[2]!);
  const right = m[3] ? colToNum(m[3]) : left;
  const bottom = m[4] ? Number(m[4]) : top;
  if (top < 1 || left < 1 || bottom < top || right < left) {
    throw new ExcelToolError(`invalid range "${range}" — start must be top-left, end bottom-right`);
  }
  return { top, left, bottom, right };
}

function cellData(cell: ExcelJS.Cell): CellData {
  const value = cell.value;
  let raw: unknown = value;
  let formula: string | undefined;

  if (value !== null && typeof value === "object" && !(value instanceof Date)) {
    if ("formula" in value || "sharedFormula" in value) {
      const fv = value as ExcelJS.CellFormulaValue;
      formula = fv.formula;
      raw = fv.result ?? null;
    } else if ("richText" in value) {
      raw = cell.text;
    } else if ("hyperlink" in value) {
      raw = (value as ExcelJS.CellHyperlinkValue).text ?? cell.text;
    } else if ("error" in value) {
      raw = (value as ExcelJS.CellErrorValue).error;
    }
  }

  return {
    addr: cell.address,
    raw,
    ...(formula !== undefined ? { formula } : {}),
    numberFormat: cell.numFmt ?? null,
    text: cell.text,
  };
}

export function readExcelRange(fileId: string, sheetName: string, range: string) {
  const workbook = requireWorkbook(fileId);
  const sheet = workbook.getWorksheet(sheetName);
  if (!sheet) {
    const names = workbook.worksheets.map((ws) => ws.name).join(", ");
    throw new ExcelToolError(`unknown sheet "${sheetName}" — available sheets: ${names}`);
  }

  const { top, left, bottom, right } = parseA1Range(range);
  const cellCount = (bottom - top + 1) * (right - left + 1);
  if (cellCount > MAX_CELLS_PER_READ) {
    throw new ExcelToolError(
      `range "${range}" spans ${cellCount} cells (max ${MAX_CELLS_PER_READ} per call) — read it in smaller chunks`,
    );
  }

  const cells: CellData[] = [];
  for (let r = top; r <= bottom; r++) {
    const row = sheet.getRow(r);
    for (let c = left; c <= right; c++) {
      cells.push(cellData(row.getCell(c)));
    }
  }
  return { sheet: sheet.name, range: range.toUpperCase(), cells };
}
