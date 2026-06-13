/**
 * Module 14: Scoring Engine
 * Calculate performance scores and grades with safety rules
 */

import { ErrorType } from "../types";

export type AccuracyGrade =
  | "Outstanding"
  | "Excellent"
  | "Very Good"
  | "Good"
  | "Fair"
  | "Needs Improvement"
  | "Poor";

export interface ScoringResult {
  comparedCells: number;
  totalErrors: number;
  baseAccuracy: number;
  finalGrade: AccuracyGrade;
  totalPenaltyPoints: number;
  errorRatePer10k: number;
  reviewerWorkloadIndex: number;
}

/**
 * Calculate accuracy percentage
 */
export function calculateAccuracy(
  comparedCells: number,
  totalErrors: number
): number {
  if (comparedCells === 0) return 0;
  const accuracy = ((comparedCells - totalErrors) / comparedCells) * 100;
  return Math.round(accuracy * 100) / 100; // 2 decimal places
}

/**
 * Assign grade based on accuracy
 */
export function assignGrade(
  accuracy: number,
  hasRowShift: boolean = false,
  hasColumnShift: boolean = false,
  criticalErrorCount: number = 0,
  majorErrorCount: number = 0
): AccuracyGrade {
  // Safety rule: shifts cap grade at "Needs Improvement"
  if (hasRowShift || hasColumnShift) {
    if (accuracy >= 95) return "Needs Improvement";
    if (accuracy >= 85) return "Needs Improvement";
    if (accuracy >= 75) return "Poor";
    return "Poor";
  }

  // Safety rule: excessive critical errors cap grade
  if (criticalErrorCount > 10) {
    return "Needs Improvement";
  }

  // Safety rule: excessive major errors
  if (majorErrorCount > 50) {
    return "Fair";
  }

  // Standard grading
  if (accuracy >= 99.9) return "Outstanding";
  if (accuracy >= 99.0) return "Excellent";
  if (accuracy >= 97.0) return "Very Good";
  if (accuracy >= 95.0) return "Good";
  if (accuracy >= 90.0) return "Fair";
  if (accuracy >= 80.0) return "Needs Improvement";
  return "Poor";
}

/**
 * Calculate penalty points
 */
export function calculatePenaltyPoints(errorLog: any[]): number {
  let totalPenalty = 0;

  errorLog.forEach((error) => {
    if (error.suppressed) return; // Skip suppressed errors

    // Penalty based on severity
    const severityMultiplier = {
      Critical: 5,
      High: 2,
      Medium: 1,
      Low: 0.5
    };

    const multiplier = severityMultiplier[error.severity] || 1;
    totalPenalty += error.penalty * multiplier;
  });

  return Math.round(totalPenalty * 100) / 100;
}

/**
 * Calculate error rate per 10,000 cells
 */
export function calculateErrorRatePer10k(
  totalErrors: number,
  comparedCells: number
): number {
  if (comparedCells === 0) return 0;
  return Math.round((totalErrors / comparedCells) * 10000);
}

/**
 * Estimate reviewer workload (hours)
 */
export function estimateReviewerWorkload(
  errorLog: any[],
  avgMinutesPerError: number = 5
): number {
  const nonSuppressed = errorLog.filter((e) => !e.suppressed);
  const totalMinutes = nonSuppressed.length * avgMinutesPerError;
  return Math.round((totalMinutes / 60) * 10) / 10; // Convert to hours, 1 decimal place
}

/**
 * Comprehensive scoring
 */
export function score(
  errorLog: any[],
  comparedCells: number,
  hasShiftEvents: boolean = false,
  criticalCount: number = 0,
  majorCount: number = 0
): ScoringResult {
  const nonSuppressed = errorLog.filter((e) => !e.suppressed);
  const totalErrors = nonSuppressed.length;
  const baseAccuracy = calculateAccuracy(comparedCells, totalErrors);
  const finalGrade = assignGrade(
    baseAccuracy,
    hasShiftEvents,
    hasShiftEvents,
    criticalCount,
    majorCount
  );
  const totalPenaltyPoints = calculatePenaltyPoints(errorLog);
  const errorRatePer10k = calculateErrorRatePer10k(totalErrors, comparedCells);
  const reviewerWorkloadIndex = estimateReviewerWorkload(errorLog);

  return {
    comparedCells,
    totalErrors,
    baseAccuracy,
    finalGrade,
    totalPenaltyPoints,
    errorRatePer10k,
    reviewerWorkloadIndex
  };
}
