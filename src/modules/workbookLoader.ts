/**
 * Module 1: Workbook Loader
 * Load and parse Excel files with intelligent bounds detection
 */

import * as XLSX from "xlsx";
import { WorkbookData, CellValue } from "../types";

export interface WorkbookLoaderConfig {
  maxRowsInMemory?: number;
  maxColsInMemory?: number;
  includeEmptyCells?: boolean;
}

/**
 * Check if a cell is truly populated (raw value or formatted text)
 */
function isCellPopulated(cell: any): boolean {
  if (!cell) return false;
  if (cell.v !== undefined && cell.v !== null && cell.v !== "") return true;
  if (cell.w !== undefined && cell.w !== null && cell.w.trim() !== "") return true;
  return false;
}

/**
 * Detect tight bounds of actual populated cells
 * Prevents rendering massive empty space (e.g., 2000 empty rows)
 */
function detectTightBounds(sheet: any): { maxRow: number; maxCol: number } {
  let maxRow = 0;
  let maxCol = 0;

  for (const key in sheet) {
    if (key[0] === "!") continue; // skip metadata
    const parsed = XLSX.utils.decode_cell(key);
    const cell = sheet[key];
    if (isCellPopulated(cell)) {
      if (parsed.r > maxRow) maxRow = parsed.r;
      if (parsed.c > maxCol) maxCol = parsed.c;
    }
  }

  return { maxRow, maxCol };
}

/**
 * Parse Excel file and extract workbook data
 */
export async function loadWorkbookFromFile(
  file: File,
  config: WorkbookLoaderConfig = {}
): Promise<WorkbookData> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        if (!data) throw new Error("File empty or unreadable");

        const workbook = XLSX.read(data, { type: "binary" });
        const wbData: WorkbookData = {
          fileName: file.name,
          sheets: {}
        };

        workbook.SheetNames.forEach((sheetName) => {
          const sheet = workbook.Sheets[sheetName];
          const bounds = detectTightBounds(sheet);

          // Expand merged cells (Module 3 integration)
          if (sheet["!merges"]) {
            sheet["!merges"].forEach((merge) => {
              const startCellRef = XLSX.utils.encode_cell({ r: merge.s.r, c: merge.s.c });
              const startCellValue = sheet[startCellRef];
              if (startCellValue) {
                const endRow = Math.min(merge.e.r, bounds.maxRow);
                const endCol = Math.min(merge.e.c, bounds.maxCol);

                for (let r = merge.s.r; r <= endRow; r++) {
                  for (let c = merge.s.c; c <= endCol; c++) {
                    const currentRef = XLSX.utils.encode_cell({ r, c });
                    if (currentRef !== startCellRef) {
                      sheet[currentRef] = { ...startCellValue };
                    }
                  }
                }
              }
            });
          }

          // Extract cells within bounds
          const cells: Record<string, CellValue> = {};
          for (let r = 0; r <= bounds.maxRow; r++) {
            for (let c = 0; c <= bounds.maxCol; c++) {
              const cellRef = XLSX.utils.encode_cell({ r, c });
              const cell = sheet[cellRef];
              if (isCellPopulated(cell)) {
                const rawVal = cell.v !== undefined && cell.v !== null && cell.v !== "" ? cell.v : (cell.w || "");
                cells[`${r},${c}`] = {
                  raw: rawVal,
                  formatted: cell.w || String(rawVal),
                  normalized: String(rawVal),
                  type: typeof rawVal === "number" ? "number" : "string"
                };
              }
            }
          }

          wbData.sheets[sheetName] = {
            name: sheetName,
            maxRow: bounds.maxRow,
            maxCol: bounds.maxCol,
            cells
          };
        });

        resolve(wbData);
      } catch (err: any) {
        reject(err);
      }
    };

    reader.onerror = () => reject(new Error("File read error"));
    reader.readAsBinaryString(file);
  });
}
