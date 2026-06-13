/**
 * Module 11: Error Classifier
 * Classify 12+ error types
 */

import { ErrorType, Severity } from "../types";
import { CellComparisonResult } from "./cellComparisonEngine";
import { levenshteinDistance } from "./alignmentRecoveryEngine";

export interface ClassificationResult {
  errorType: ErrorType;
  severity: Severity;
  confidence: number;
  description: string;
}

/**
 * Classify error based on comparison result
 */
export function classifyError(
  employeeValue: string,
  reviewerValue: string,
  comparisonResult: CellComparisonResult,
  isHeaderRow: boolean = false
): ClassificationResult {
  // Header row errors
  if (isHeaderRow) {
    return {
      errorType: ErrorType.HeaderMismatch,
      severity: Severity.Critical,
      confidence: 1.0,
      description: "Header text mismatch"
    };
  }

  // Empty vs non-empty
  if (employeeValue === "" && reviewerValue !== "") {
    return {
      errorType: ErrorType.MissingValue,
      severity: Severity.High,
      confidence: 1.0,
      description: "Employee left cell empty"
    };
  }
  if (employeeValue !== "" && reviewerValue === "") {
    return {
      errorType: ErrorType.ExtraValue,
      severity: Severity.Medium,
      confidence: 1.0,
      description: "Employee entered extra data"
    };
  }

  // Numeric errors
  if (isNumeric(employeeValue) && isNumeric(reviewerValue)) {
    const empNum = parseFloat(employeeValue);
    const revNum = parseFloat(reviewerValue);

    // Check for digit substitution/transposition
    if (Math.abs(empNum - revNum) === 1) {
      return {
        errorType: ErrorType.DigitSubstitution,
        severity: Severity.High,
        confidence: 0.95,
        description: "Single digit substitution"
      };
    }

    // Check if digits are transposed
    if (isTransposition(employeeValue, reviewerValue)) {
      return {
        errorType: ErrorType.DigitTransposition,
        severity: Severity.High,
        confidence: 0.9,
        description: "Digit transposition detected"
      };
    }

    // Numeric difference
    return {
      errorType: ErrorType.NumericDifference,
      severity: Severity.High,
      confidence: 0.8,
      description: `Numeric difference: ${empNum} vs ${revNum}`
    };
  }

  // Text errors
  const distance = levenshteinDistance(employeeValue, reviewerValue);
  const similarity = comparisonResult.similarity;

  if (similarity > 85) {
    return {
      errorType: ErrorType.TextTypo,
      severity: Severity.Low,
      confidence: 0.9,
      description: "Minor text difference (typo)"
    };
  } else if (similarity > 50) {
    return {
      errorType: ErrorType.TextDifference,
      severity: Severity.Medium,
      confidence: 0.8,
      description: "Text difference"
    };
  } else {
    return {
      errorType: ErrorType.MajorTextDifference,
      severity: Severity.Critical,
      confidence: 0.7,
      description: "Major text difference"
    };
  }
}

/**
 * Check if string is numeric
 */
function isNumeric(value: string): boolean {
  return !isNaN(parseFloat(value)) && isFinite(Number(value));
}

/**
 * Check if digits are transposed
 */
function isTransposition(val1: string, val2: string): boolean {
  // Simple check: same digits, different order
  const arr1 = val1.split("").sort();
  const arr2 = val2.split("").sort();
  return arr1.join("") === arr2.join("") && val1 !== val2;
}
