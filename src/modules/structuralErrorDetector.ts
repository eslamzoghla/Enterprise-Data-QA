/**
 * Module 6: Structural Error Detector
 * Detect all 12 structural error types
 */

import { WorkbookData, ErrorType } from "../types";
import { validateStructures } from "./structureValidator";

export interface StructuralError {
  type: ErrorType;
  sheet?: string;
  description: string;
  affectedCells: number;
}

/**
 * Detect structural errors between two workbooks
 */
export function detectStructuralErrors(
  employeeWb: WorkbookData,
  reviewerWb: WorkbookData
): StructuralError[] {
  const errors: StructuralError[] = [];
  const validation = validateStructures(employeeWb, reviewerWb);

  // Missing Sheets
  validation.missingSheets.forEach((sheet) => {
    errors.push({
      type: ErrorType.MissingSheet,
      sheet,
      description: `Sheet "${sheet}" exists in Reviewer but not in Employee submission`,
      affectedCells: 0
    });
  });

  // Extra Sheets
  validation.extraSheets.forEach((sheet) => {
    errors.push({
      type: ErrorType.ExtraSheet,
      sheet,
      description: `Sheet "${sheet}" exists in Employee submission but not in Reviewer file`,
      affectedCells: Object.keys(employeeWb.sheets[sheet]?.cells || {}).length
    });
  });

  // Check shared sheets for column/row mismatches
  const empSheets = Object.keys(employeeWb.sheets);
  const revSheets = Object.keys(reviewerWb.sheets);

  for (const sheet of empSheets) {
    if (!revSheets.includes(sheet)) continue;

    const empSheet = employeeWb.sheets[sheet];
    const revSheet = reviewerWb.sheets[sheet];

    // Missing Columns
    if (empSheet.maxCol < revSheet.maxCol) {
      errors.push({
        type: ErrorType.MissingColumn,
        sheet,
        description: `Missing column(s) in Employee sheet (Employee: ${empSheet.maxCol + 1} cols, Reviewer: ${revSheet.maxCol + 1} cols)`,
        affectedCells: (revSheet.maxCol - empSheet.maxCol) * (empSheet.maxRow + 1)
      });
    }

    // Extra Columns
    if (empSheet.maxCol > revSheet.maxCol) {
      errors.push({
        type: ErrorType.ExtraColumn,
        sheet,
        description: `Extra column(s) in Employee sheet (Employee: ${empSheet.maxCol + 1} cols, Reviewer: ${revSheet.maxCol + 1} cols)`,
        affectedCells: (empSheet.maxCol - revSheet.maxCol) * (empSheet.maxRow + 1)
      });
    }

    // Missing Rows
    if (empSheet.maxRow < revSheet.maxRow) {
      errors.push({
        type: ErrorType.MissingRow,
        sheet,
        description: `Missing row(s) in Employee sheet (Employee: ${empSheet.maxRow + 1} rows, Reviewer: ${revSheet.maxRow + 1} rows)`,
        affectedCells: (revSheet.maxRow - empSheet.maxRow) * (empSheet.maxCol + 1)
      });
    }

    // Extra Rows
    if (empSheet.maxRow > revSheet.maxRow) {
      errors.push({
        type: ErrorType.ExtraRow,
        sheet,
        description: `Extra row(s) in Employee sheet (Employee: ${empSheet.maxRow + 1} rows, Reviewer: ${revSheet.maxRow + 1} rows)`,
        affectedCells: (empSheet.maxRow - revSheet.maxRow) * (empSheet.maxCol + 1)
      });
    }
  }

  return errors;
}
