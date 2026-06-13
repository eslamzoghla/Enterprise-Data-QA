/**
 * Module 12: Pattern Detection Engine
 * Detect 9+ error patterns
 */

import { ErrorType } from "../types";

export interface DetectedPattern {
  type: string;
  count: number;
  description: string;
  examples: string[];
}

export interface PatternAnalysis {
  repeatedNumericErrors: string[];
  repeatedDigitSubstitutions: string[];
  copyPasteErrors: string[];
  errorClusters: string[];
  sheetConcentrations: string[];
  shiftEvents: any[];
}

/**
 * Detect repeated numeric errors
 */
export function detectRepeatedNumericErrors(
  errorLog: any[]
): string[] {
  const numericErrors = errorLog.filter(
    (e) =>
      e.errorType === ErrorType.NumericDifference ||
      e.errorType === ErrorType.DigitSubstitution ||
      e.errorType === ErrorType.DigitTransposition
  );

  const patterns: Record<string, number> = {};
  numericErrors.forEach((error) => {
    const key = `${error.sheet}:${error.employeeValue} vs ${error.reviewerValue}`;
    patterns[key] = (patterns[key] || 0) + 1;
  });

  return Object.entries(patterns)
    .filter(([_, count]) => count > 1)
    .map(([pattern, _]) => pattern);
}

/**
 * Detect copy-paste errors
 */
export function detectCopyPasteErrors(
  errorLog: any[]
): string[] {
  const errors = errorLog.filter((e) => e.errorType === ErrorType.MajorNumericError);
  const patterns: Record<string, number> = {};

  errors.forEach((error) => {
    const key = error.employeeValue;
    patterns[key] = (patterns[key] || 0) + 1;
  });

  return Object.entries(patterns)
    .filter(([_, count]) => count > 2)
    .map(([value, count]) => `Value "${value}" repeated ${count} times in different cells`);
}

/**
 * Detect error clusters
 */
export function detectErrorClusters(
  errorLog: any[]
): string[] {
  const clusters: Record<string, number> = {};

  errorLog.forEach((error) => {
    const [row, col] = error.cell.match(/\d+/g).map(Number);
    const region = `Row ${Math.floor(row / 5) * 5}-${Math.floor(row / 5) * 5 + 4}, Col ${Math.floor(col / 5) * 5}-${Math.floor(col / 5) * 5 + 4}`;
    clusters[region] = (clusters[region] || 0) + 1;
  });

  return Object.entries(clusters)
    .filter(([_, count]) => count > 5)
    .map(([region, count]) => `Cluster in ${region}: ${count} errors`);
}

/**
 * Detect sheet-level concentrations
 */
export function detectSheetConcentrations(
  errorLog: any[]
): string[] {
  const sheets: Record<string, number> = {};

  errorLog.forEach((error) => {
    sheets[error.sheet] = (sheets[error.sheet] || 0) + 1;
  });

  const total = errorLog.length;
  return Object.entries(sheets)
    .map(([sheet, count]) => `${sheet}: ${count} errors (${Math.round((count / total) * 100)}%)`)
    .sort((a, b) => parseInt(b.split(":")[1]) - parseInt(a.split(":")[1]));
}
