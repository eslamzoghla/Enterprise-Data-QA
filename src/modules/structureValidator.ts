/**
 * Module 4: Structure Validator
 * Validate table structure integrity
 */

import { WorkbookData } from "../types";

export interface StructureValidationResult {
  isValid: boolean;
  errors: string[];
  warnings: string[];
  missingSheets: string[];
  extraSheets: string[];
}

/**
 * Compare structures of two workbooks
 */
export function validateStructures(
  employeeWb: WorkbookData,
  reviewerWb: WorkbookData
): StructureValidationResult {
  const result: StructureValidationResult = {
    isValid: true,
    errors: [],
    warnings: [],
    missingSheets: [],
    extraSheets: []
  };

  const empSheets = new Set(Object.keys(employeeWb.sheets));
  const revSheets = new Set(Object.keys(reviewerWb.sheets));

  // Detect missing sheets
  for (const sheet of revSheets) {
    if (!empSheets.has(sheet)) {
      result.missingSheets.push(sheet);
      result.errors.push(`Missing sheet: "${sheet}"`);
      result.isValid = false;
    }
  }

  // Detect extra sheets
  for (const sheet of empSheets) {
    if (!revSheets.has(sheet)) {
      result.extraSheets.push(sheet);
      result.warnings.push(`Extra sheet: "${sheet}"`);
    }
  }

  // Validate shared sheets have content
  for (const sheet of empSheets) {
    if (revSheets.has(sheet)) {
      const empSheet = employeeWb.sheets[sheet];
      const revSheet = reviewerWb.sheets[sheet];

      if (Object.keys(empSheet.cells).length === 0) {
        result.warnings.push(`Employee sheet "${sheet}" is empty`);
      }
      if (Object.keys(revSheet.cells).length === 0) {
        result.warnings.push(`Reviewer sheet "${sheet}" is empty`);
      }
    }
  }

  return result;
}
