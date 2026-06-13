/**
 * Module 16: Grid Inspector Engine
 * Display 8+ metrics for selected table
 */

import { ErrorType, Severity } from "../types";

export interface GridMetrics {
  comparedCells: number;
  totalErrors: number;
  accuracy: number;
  structuralErrors: number;
  shiftErrors: number;
  rangeErrors: number;
  numericErrors: number;
  textErrors: number;
  headerErrors: number;
}

/**
 * Calculate grid metrics for selected sheet
 */
export function calculateGridMetrics(errorLog: any[]): GridMetrics {
  const nonSuppressed = errorLog.filter((e) => !e.suppressed);

  let structural = 0;
  let shift = 0;
  let range = 0;
  let numeric = 0;
  let text = 0;
  let header = 0;

  nonSuppressed.forEach((error) => {
    if (error.notes?.includes("[Header Row Error]")) {
      header++;
      return;
    }

    switch (error.errorType) {
      case ErrorType.MissingDigit:
      case ErrorType.ExtraDigit:
      case ErrorType.DigitSubstitution:
      case ErrorType.DigitTransposition:
      case ErrorType.NumericDifference:
      case ErrorType.MajorNumericError:
        numeric++;
        break;
      case ErrorType.MissingValue:
      case ErrorType.ExtraValue:
      case ErrorType.MissingRow:
      case ErrorType.ExtraRow:
      case ErrorType.MissingColumn:
      case ErrorType.ExtraColumn:
      case ErrorType.TableMerge:
      case ErrorType.TableSplit:
        structural++;
        break;
      case ErrorType.RowShift:
      case ErrorType.ColumnShift:
      case ErrorType.LocalColumnMisalignment:
      case ErrorType.LocalRowMisalignment:
        shift++;
        break;
      case ErrorType.TextTypo:
      case ErrorType.MajorTextDifference:
      case ErrorType.TextDifference:
        text++;
        break;
      case ErrorType.RangeInversionError:
      case ErrorType.RangeBoundaryError:
      case ErrorType.RangeRepresentationError:
        range++;
        break;
    }
  });

  return {
    comparedCells: nonSuppressed.length > 0 ? 1000 : 0, // Placeholder
    totalErrors: nonSuppressed.length,
    accuracy: nonSuppressed.length > 0 ? 95 : 100, // Placeholder
    structuralErrors: structural,
    shiftErrors: shift,
    rangeErrors: range,
    numericErrors: numeric,
    textErrors: text,
    headerErrors: header
  };
}

/**
 * Format metrics for display
 */
export function formatGridMetrics(metrics: GridMetrics): Record<string, string> {
  return {
    "Compared Cells": metrics.comparedCells.toLocaleString(),
    "Total Errors": metrics.totalErrors.toString(),
    "Accuracy": metrics.accuracy.toFixed(2) + "%",
    "Structural Errors": metrics.structuralErrors.toString(),
    "Shift Errors": metrics.shiftErrors.toString(),
    "Range Errors": metrics.rangeErrors.toString(),
    "Numeric Errors": metrics.numericErrors.toString(),
    "Text Errors": metrics.textErrors.toString(),
    "Header Errors": metrics.headerErrors.toString()
  };
}
