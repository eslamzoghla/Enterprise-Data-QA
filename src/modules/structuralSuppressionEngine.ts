/**
 * Module 7: Structural Suppression Engine
 * Suppress dependent errors based on root causes
 */

import { ErrorType } from "../types";

export interface SuppressionRule {
  cause: ErrorType;
  suppressedTypes: ErrorType[];
  reason: string;
}

/**
 * Define suppression rules: root cause -> dependent errors
 */
const SUPPRESSION_RULES: SuppressionRule[] = [
  {
    cause: ErrorType.MissingSheet,
    suppressedTypes: [
      ErrorType.MissingRow,
      ErrorType.ExtraRow,
      ErrorType.MissingColumn,
      ErrorType.ExtraColumn,
      ErrorType.MissingCell,
      ErrorType.ExtraCell,
      ErrorType.TextTypo,
      ErrorType.NumericDifference,
      ErrorType.TextDifference
    ],
    reason: "Entire sheet missing - all cell errors suppressed"
  },
  {
    cause: ErrorType.ExtraSheet,
    suppressedTypes: [
      ErrorType.MissingRow,
      ErrorType.ExtraRow,
      ErrorType.MissingColumn,
      ErrorType.ExtraColumn,
      ErrorType.TextTypo,
      ErrorType.NumericDifference
    ],
    reason: "Extra sheet has no corresponding data to compare"
  },
  {
    cause: ErrorType.MissingColumn,
    suppressedTypes: [
      ErrorType.MissingValue,
      ErrorType.ExtraValue,
      ErrorType.TextTypo,
      ErrorType.NumericDifference,
      ErrorType.TextDifference
    ],
    reason: "Missing column explains all cell mismatches in that column"
  },
  {
    cause: ErrorType.ExtraColumn,
    suppressedTypes: [
      ErrorType.ExtraValue,
      ErrorType.TextTypo,
      ErrorType.NumericDifference
    ],
    reason: "Extra column explains all value mismatches"
  },
  {
    cause: ErrorType.MissingRow,
    suppressedTypes: [
      ErrorType.MissingValue,
      ErrorType.TextTypo,
      ErrorType.NumericDifference,
      ErrorType.TextDifference
    ],
    reason: "Missing row explains all cell mismatches in that row"
  },
  {
    cause: ErrorType.ExtraRow,
    suppressedTypes: [
      ErrorType.ExtraValue,
      ErrorType.TextTypo,
      ErrorType.NumericDifference
    ],
    reason: "Extra row explains all value mismatches"
  },
  {
    cause: ErrorType.RowShift,
    suppressedTypes: [
      ErrorType.TextTypo,
      ErrorType.NumericDifference,
      ErrorType.TextDifference,
      ErrorType.MissingValue,
      ErrorType.ExtraValue
    ],
    reason: "Row shift explains apparent cell mismatches"
  },
  {
    cause: ErrorType.ColumnShift,
    suppressedTypes: [
      ErrorType.TextTypo,
      ErrorType.NumericDifference,
      ErrorType.TextDifference,
      ErrorType.MissingValue,
      ErrorType.ExtraValue
    ],
    reason: "Column shift explains apparent cell mismatches"
  }
];

/**
 * Apply suppression rules to error log
 */
export function applySuppressionRules(
  errorLog: any[],
  structuralErrors: any[] = []
): any[] {
  if (errorLog.length === 0) return errorLog;

  // Track which errors should be suppressed
  const suppressed = new Set<number>();

  // Find all root cause errors
  const rootCauses = structuralErrors.filter(
    (err) => err.type && SUPPRESSION_RULES.some((rule) => rule.cause === err.type)
  );

  // Apply suppression for each root cause
  rootCauses.forEach((rootCause) => {
    const rule = SUPPRESSION_RULES.find((r) => r.cause === rootCause.type);
    if (rule) {
      // Find all dependent errors to suppress
      errorLog.forEach((error, index) => {
        // Check if error is related to root cause (e.g., same sheet/region)
        if (
          rootCause.sheet &&
          error.sheet === rootCause.sheet &&
          rule.suppressedTypes.includes(error.errorType)
        ) {
          suppressed.add(index);
        }
      });
    }
  });

  // Return non-suppressed errors
  return errorLog.map((error, index) => ({
    ...error,
    suppressed: suppressed.has(index),
    suppressionReason: suppressed.has(index) ? "Suppressed by root cause analysis" : null
  }));
}

/**
 * Get suppression rule for error type
 */
export function getSuppressionRule(errorType: ErrorType): SuppressionRule | undefined {
  return SUPPRESSION_RULES.find((rule) => rule.cause === errorType);
}

/**
 * Count non-suppressed errors
 */
export function countNonSuppressedErrors(errorLog: any[]): number {
  return errorLog.filter((error) => !error.suppressed).length;
}
