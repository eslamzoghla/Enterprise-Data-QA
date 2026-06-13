/**
 * Module 8: Cell Comparison Engine
 * Compare cell values with context and normalization
 */

import { CellValue, ErrorType } from "../types";
import { normalizeWithLanguageDetection } from "./arabicNormalizationEngine";
import { parseNumericValue, areNumericEquivalent } from "./numericNormalizationEngine";
import { levenshteinDistance, calculateSimilarity } from "./alignmentRecoveryEngine";

export interface CellComparisonResult {
  match: boolean;
  similarity: number;
  errorType?: ErrorType;
  employeeNormalized: string;
  reviewerNormalized: string;
}

/**
 * Compare two cell values
 */
export function compareCells(
  employeeCell: CellValue | undefined,
  reviewerCell: CellValue | undefined,
  config: {
    arabicComparisonMode: "STANDARD" | "NONE";
    numericTolerance: number;
    numericToleranceMode: "ABSOLUTE" | "PERCENTAGE";
  }
): CellComparisonResult {
  const empVal = employeeCell ? String(employeeCell.raw) : "";
  const revVal = reviewerCell ? String(reviewerCell.raw) : "";

  // Exact match
  if (empVal === revVal) {
    return {
      match: true,
      similarity: 100,
      employeeNormalized: empVal,
      reviewerNormalized: revVal
    };
  }

  // Try numeric comparison
  if (isNumericContent(empVal) && isNumericContent(revVal)) {
    const numMatch = areNumericEquivalent(
      empVal,
      revVal,
      config.numericTolerance,
      config.numericToleranceMode
    );
    if (numMatch) {
      return {
        match: true,
        similarity: 100,
        employeeNormalized: empVal,
        reviewerNormalized: revVal
      };
    }
  }

  // Normalize based on content type
  let empNormalized = empVal;
  let revNormalized = revVal;

  if (config.arabicComparisonMode === "STANDARD") {
    empNormalized = normalizeWithLanguageDetection(empVal);
    revNormalized = normalizeWithLanguageDetection(revVal);

    if (empNormalized === revNormalized) {
      return {
        match: true,
        similarity: 100,
        employeeNormalized: empNormalized,
        reviewerNormalized: revNormalized
      };
    }
  }

  // Calculate similarity
  const similarity = calculateSimilarity(empNormalized, revNormalized);

  return {
    match: false,
    similarity,
    employeeNormalized: empNormalized,
    reviewerNormalized: revNormalized
  };
}

/**
 * Check if content is numeric
 */
function isNumericContent(text: string): boolean {
  const parsed = parseNumericValue(text);
  return parsed !== null;
}
