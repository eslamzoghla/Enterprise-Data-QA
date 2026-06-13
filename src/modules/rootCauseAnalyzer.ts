/**
 * Module 13: Root Cause Analyzer
 * Identify root causes and aggregate error impact
 */

import { ErrorType } from "../types";

export interface RootCauseAnalysis {
  missingValuesPct: number;
  numericErrorsPct: number;
  textErrorsPct: number;
  rangeErrorsPct: number;
  shiftErrorsPct: number;
  headerErrorsPct: number;
  structuralEventCount: number;
  patterns: Record<string, number>;
}

/**
 * Analyze root causes in error log
 */
export function analyzeRootCauses(errorLog: any[]): RootCauseAnalysis {
  if (errorLog.length === 0) {
    return {
      missingValuesPct: 0,
      numericErrorsPct: 0,
      textErrorsPct: 0,
      rangeErrorsPct: 0,
      shiftErrorsPct: 0,
      headerErrorsPct: 0,
      structuralEventCount: 0,
      patterns: {}
    };
  }

  const nonSuppressed = errorLog.filter((e) => !e.suppressed);
  const total = nonSuppressed.length;

  const counts = {
    missingValues: 0,
    numericErrors: 0,
    textErrors: 0,
    rangeErrors: 0,
    shiftErrors: 0,
    headerErrors: 0,
    structural: 0
  };

  nonSuppressed.forEach((error) => {
    const notes = error.notes || "";

    if (notes.includes("[Header Row Error]")) {
      counts.headerErrors++;
    } else if (
      error.errorType === ErrorType.MissingValue ||
      error.errorType === ErrorType.ExtraValue
    ) {
      counts.missingValues++;
    } else if (
      error.errorType === ErrorType.NumericDifference ||
      error.errorType === ErrorType.DigitSubstitution ||
      error.errorType === ErrorType.DigitTransposition ||
      error.errorType === ErrorType.MissingDigit ||
      error.errorType === ErrorType.ExtraDigit ||
      error.errorType === ErrorType.MajorNumericError
    ) {
      counts.numericErrors++;
    } else if (
      error.errorType === ErrorType.TextTypo ||
      error.errorType === ErrorType.TextDifference ||
      error.errorType === ErrorType.MajorTextDifference
    ) {
      counts.textErrors++;
    } else if (
      error.errorType === ErrorType.RangeInversionError ||
      error.errorType === ErrorType.RangeBoundaryError ||
      error.errorType === ErrorType.RangeRepresentationError
    ) {
      counts.rangeErrors++;
    } else if (
      error.errorType === ErrorType.RowShift ||
      error.errorType === ErrorType.ColumnShift ||
      error.errorType === ErrorType.LocalRowMisalignment ||
      error.errorType === ErrorType.LocalColumnMisalignment
    ) {
      counts.shiftErrors++;
    }

    if (
      error.errorType === ErrorType.TableMerge ||
      error.errorType === ErrorType.TableSplit ||
      error.errorType === ErrorType.MissingSheet ||
      error.errorType === ErrorType.ExtraSheet
    ) {
      counts.structural++;
    }
  });

  return {
    missingValuesPct: Math.round((counts.missingValues / total) * 100),
    numericErrorsPct: Math.round((counts.numericErrors / total) * 100),
    textErrorsPct: Math.round((counts.textErrors / total) * 100),
    rangeErrorsPct: Math.round((counts.rangeErrors / total) * 100),
    shiftErrorsPct: Math.round((counts.shiftErrors / total) * 100),
    headerErrorsPct: Math.round((counts.headerErrors / total) * 100),
    structuralEventCount: counts.structural,
    patterns: {
      missingValues: counts.missingValues,
      numericErrors: counts.numericErrors,
      textErrors: counts.textErrors,
      rangeErrors: counts.rangeErrors,
      shiftErrors: counts.shiftErrors,
      headerErrors: counts.headerErrors
    }
  };
}

/**
 * Identify predominant root cause
 */
export function identifyPredominantCause(
  analysis: RootCauseAnalysis
): string {
  const causes = [
    { name: "Missing/Extra Values", value: analysis.missingValuesPct },
    { name: "Numeric Errors", value: analysis.numericErrorsPct },
    { name: "Text Errors", value: analysis.textErrorsPct },
    { name: "Range Errors", value: analysis.rangeErrorsPct },
    { name: "Shift Events", value: analysis.shiftErrorsPct },
    { name: "Header Errors", value: analysis.headerErrorsPct }
  ];

  const predominant = causes.reduce((prev, current) =>
    current.value > prev.value ? current : prev
  );

  return predominant.name;
}
