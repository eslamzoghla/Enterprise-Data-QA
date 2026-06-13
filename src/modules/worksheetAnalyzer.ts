/**
 * Module 2: Worksheet Analyzer
 * Analyze worksheet structure and metadata
 */

import { WorkbookData } from "../types";

export interface WorksheetMetadata {
  name: string;
  totalCells: number;
  totalRows: number;
  totalColumns: number;
  populatedCells: number;
  density: number; // percentage of populated cells
  hasHeaders: boolean;
  estimatedHeaderRow: number;
}

/**
 * Analyze a single worksheet
 */
export function analyzeWorksheet(
  workbookData: WorkbookData,
  sheetName: string
): WorksheetMetadata {
  const sheet = workbookData.sheets[sheetName];
  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found`);
  }

  const totalRows = sheet.maxRow + 1;
  const totalColumns = sheet.maxCol + 1;
  const totalCells = totalRows * totalColumns;
  const populatedCells = Object.keys(sheet.cells).length;
  const density = (populatedCells / totalCells) * 100;

  // Detect headers: typically row 0 has different structure/content
  const hasHeaders = populatedCells > 0;
  const estimatedHeaderRow = 0;

  return {
    name: sheetName,
    totalCells,
    totalRows,
    totalColumns,
    populatedCells,
    density,
    hasHeaders,
    estimatedHeaderRow
  };
}

/**
 * Analyze all worksheets in workbook
 */
export function analyzeAllWorksheets(workbookData: WorkbookData): WorksheetMetadata[] {
  return Object.keys(workbookData.sheets).map((sheetName) =>
    analyzeWorksheet(workbookData, sheetName)
  );
}

/**
 * Get header row from sheet
 */
export function extractHeaderRow(
  workbookData: WorkbookData,
  sheetName: string
): Record<number, string> {
  const sheet = workbookData.sheets[sheetName];
  if (!sheet) return {};

  const headers: Record<number, string> = {};
  for (let c = 0; c <= sheet.maxCol; c++) {
    const cellKey = `0,${c}`;
    const cell = sheet.cells[cellKey];
    if (cell) {
      headers[c] = cell.formatted || String(cell.raw);
    }
  }
  return headers;
}
