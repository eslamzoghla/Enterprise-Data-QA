/**
 * Module 15: Reporting Engine
 * Generate audit-grade reports with 6 sections
 */

import { RootCauseAnalysis } from "./rootCauseAnalyzer";
import { ScoringResult } from "./scoringEngine";

export interface AuditReport {
  executiveSummary: string;
  structuralDefectsSummary: string;
  rootCauseAnalysis: string;
  detailedErrorLog: any[];
  patternFindings: string;
  coachingRecommendations: string[];
  generatedAt: string;
}

/**
 * Generate executive summary
 */
export function generateExecutiveSummary(
  scoring: ScoringResult,
  employeeName: string,
  projectName: string,
  evaluationDate: string
): string {
  return `
## Executive Summary

**Employee:** ${employeeName}  
**Project:** ${projectName}  
**Evaluation Date:** ${evaluationDate}  

### Performance Overview
- **Final Grade:** ${scoring.finalGrade}
- **Accuracy Score:** ${scoring.baseAccuracy}%
- **Total Errors Detected:** ${scoring.totalErrors} out of ${scoring.comparedCells} compared cells
- **Penalty Points:** ${scoring.totalPenaltyPoints}
- **Error Rate:** ${scoring.errorRatePer10k} errors per 10,000 cells
- **Estimated Reviewer Workload:** ${scoring.reviewerWorkloadIndex} hours

### Key Finding
The employee achieved an accuracy of **${scoring.baseAccuracy}%**, classified as **${scoring.finalGrade}** performance.
`;
}

/**
 * Generate structural defects summary
 */
export function generateStructuralDefectsSummary(
  structuralErrors: any[]
): string {
  if (structuralErrors.length === 0) {
    return "## Structural Defects Summary\n\nNo structural defects detected. Sheet structures match perfectly.";
  }

  let summary = "## Structural Defects Summary\n\n";
  const grouped = structuralErrors.reduce((acc, err) => {
    acc[err.type] = (acc[err.type] || 0) + 1;
    return acc;
  }, {} as Record<string, number>);

  for (const [type, count] of Object.entries(grouped)) {
    summary += `- **${type}:** ${count} occurrence(s)\n`;
  }

  return summary;
}

/**
 * Generate root cause analysis
 */
export function generateRootCauseAnalysis(
  analysis: RootCauseAnalysis
): string {
  return `
## Root Cause Analysis

Based on error classification, the following root causes were identified:

| Root Cause | Percentage | Count |
|------------|-----------|-------|
| Missing/Extra Values | ${analysis.missingValuesPct}% | ${analysis.patterns.missingValues} |
| Numeric Errors | ${analysis.numericErrorsPct}% | ${analysis.patterns.numericErrors} |
| Text Errors | ${analysis.textErrorsPct}% | ${analysis.patterns.textErrors} |
| Range Errors | ${analysis.rangeErrorsPct}% | ${analysis.patterns.rangeErrors} |
| Shift Events | ${analysis.shiftErrorsPct}% | ${analysis.patterns.shiftErrors} |
| Header Errors | ${analysis.headerErrorsPct}% | ${analysis.patterns.headerErrors} |

**Analysis:** The predominant error source appears to be **${
    analysis.missingValuesPct > analysis.numericErrorsPct ? "omissions" : "numeric issues"
  }**, suggesting that employees should focus on $({
    analysis.missingValuesPct > analysis.numericErrorsPct
      ? "data entry completeness"
      : "numeric accuracy and precision"
  }).
`;
}

/**
 * Generate error log table
 */
export function generateDetailedErrorLog(
  errorLog: any[],
  maxErrors: number = 50
): any[] {
  const nonSuppressed = errorLog.filter((e) => !e.suppressed);
  const truncated = nonSuppressed.slice(0, maxErrors);

  return truncated.map((error) => ({
    sheet: error.sheet,
    cell: error.cell,
    employeeValue: error.employeeValue === "" ? "[empty]" : error.employeeValue,
    reviewerValue: error.reviewerValue === "" ? "[empty]" : error.reviewerValue,
    normalizedEmployee: error.normalizedEmployeeValue,
    normalizedReviewer: error.normalizedReviewerValue,
    similarity: error.similarity + "%",
    errorType: error.errorType,
    severity: error.severity,
    penalty: error.penalty,
    notes: error.notes
  }));
}

/**
 * Generate pattern findings
 */
export function generatePatternFindings(patterns: any): string {
  let findings = "## Pattern Findings\n\n";

  if (patterns.repeatedNumericErrors?.length > 0) {
    findings += `### Repeated Numeric Errors (${patterns.repeatedNumericErrors.length})\n`;
    patterns.repeatedNumericErrors.slice(0, 5).forEach((pattern: string) => {
      findings += `- ${pattern}\n`;
    });
  }

  if (patterns.copyPasteErrors?.length > 0) {
    findings += `\n### Copy-Paste Anomalies (${patterns.copyPasteErrors.length})\n`;
    patterns.copyPasteErrors.slice(0, 5).forEach((pattern: string) => {
      findings += `- ${pattern}\n`;
    });
  }

  if (patterns.errorClusters?.length > 0) {
    findings += `\n### Error Clusters (${patterns.errorClusters.length})\n`;
    patterns.errorClusters.slice(0, 5).forEach((cluster: string) => {
      findings += `- ${cluster}\n`;
    });
  }

  if (patterns.shiftEvents?.length > 0) {
    findings += `\n### Alignment Shift Events (${patterns.shiftEvents.length})\n`;
    patterns.shiftEvents.slice(0, 3).forEach((event: any) => {
      findings += `- ${event.type}: ${event.detail}\n`;
    });
  }

  return findings || "No significant patterns detected.";
}

/**
 * Generate coaching recommendations
 */
export function generateCoachingRecommendations(
  errorLog: any[],
  analysis: RootCauseAnalysis
): string[] {
  const recommendations: string[] = [];

  if (analysis.missingValuesPct > 30) {
    recommendations.push(
      "Focus on data entry completeness - ensure all required fields are filled before submission."
    );
  }

  if (analysis.numericErrorsPct > 25) {
    recommendations.push(
      "Double-check numeric values before entry. Consider using a calculator to verify calculations."
    );
  }

  if (analysis.textErrorsPct > 20) {
    recommendations.push(
      "Review spelling and text entry carefully. Use spell-check tools before submitting."
    );
  }

  if (analysis.shiftErrorsPct > 10) {
    recommendations.push(
      "Pay attention to row and column alignment. Ensure data is entered in correct cells."
    );
  }

  if (analysis.headerErrorsPct > 5) {
    recommendations.push(
      "Verify column headers match the expected format. Headers should be entered exactly as specified."
    );
  }

  if (recommendations.length === 0) {
    recommendations.push(
      "Performance is good! Continue maintaining current data entry standards."
    );
  }

  return recommendations;
}

/**
 * Generate complete audit report
 */
export function generateCompleteReport(
  scoring: ScoringResult,
  errorLog: any[],
  structuralErrors: any[],
  analysis: RootCauseAnalysis,
  patterns: any,
  employeeName: string,
  projectName: string,
  evaluationDate: string
): AuditReport {
  const coachingRecommendations = generateCoachingRecommendations(errorLog, analysis);

  return {
    executiveSummary: generateExecutiveSummary(
      scoring,
      employeeName,
      projectName,
      evaluationDate
    ),
    structuralDefectsSummary: generateStructuralDefectsSummary(structuralErrors),
    rootCauseAnalysis: generateRootCauseAnalysis(analysis),
    detailedErrorLog: generateDetailedErrorLog(errorLog),
    patternFindings: generatePatternFindings(patterns),
    coachingRecommendations,
    generatedAt: new Date().toISOString()
  };
}
