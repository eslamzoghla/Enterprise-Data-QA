/**
 * Module 17: Export Engine
 * Export results in multiple formats
 */

import { AuditReport } from "./reportingEngine";

export interface ExportOptions {
  format: "JSON" | "CSV" | "HTML" | "PDF";
  includeErrorLog: boolean;
  includePatterns: boolean;
  filename?: string;
}

/**
 * Export report as JSON
 */
export function exportAsJSON(report: AuditReport): string {
  return JSON.stringify(report, null, 2);
}

/**
 * Export error log as CSV
 */
export function exportErrorLogAsCSV(errorLog: any[]): string {
  if (errorLog.length === 0) return "";

  const headers = [
    "Sheet",
    "Cell",
    "Employee Value",
    "Reviewer Value",
    "Normalized Employee",
    "Normalized Reviewer",
    "Similarity %",
    "Error Type",
    "Severity",
    "Penalty",
    "Notes"
  ];

  const rows = errorLog
    .filter((e) => !e.suppressed)
    .map((error) => [
      escapeCSV(error.sheet),
      escapeCSV(error.cell),
      escapeCSV(error.employeeValue),
      escapeCSV(error.reviewerValue),
      escapeCSV(error.normalizedEmployeeValue),
      escapeCSV(error.normalizedReviewerValue),
      error.similarity,
      escapeCSV(error.errorType),
      escapeCSV(error.severity),
      error.penalty,
      escapeCSV(error.notes)
    ]);

  return [
    headers.map(escapeCSV).join(","),
    ...rows.map((row) => row.join(","))
  ].join("\n");
}

/**
 * Export report as HTML
 */
export function exportAsHTML(report: AuditReport): string {
  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Enterprise QA Engine - Audit Report</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 40px; }
    h1 { color: #333; border-bottom: 3px solid #0066cc; padding-bottom: 10px; }
    h2 { color: #0066cc; margin-top: 30px; }
    table { border-collapse: collapse; width: 100%; margin: 20px 0; }
    th, td { border: 1px solid #ddd; padding: 12px; text-align: left; }
    th { background-color: #0066cc; color: white; }
    tr:nth-child(even) { background-color: #f9f9f9; }
    .metric { display: inline-block; margin-right: 30px; }
    .timestamp { color: #666; font-size: 0.9em; }
  </style>
</head>
<body>
  <h1>Enterprise Excel QA Engine - Audit Report</h1>
  <p class="timestamp">Generated: ${report.generatedAt}</p>
  
  ${report.executiveSummary.replace(/\n/g, "<br>")}
  ${report.structuralDefectsSummary.replace(/\n/g, "<br>")}
  ${report.rootCauseAnalysis.replace(/\n/g, "<br>")}
  ${report.patternFindings.replace(/\n/g, "<br>")}
  
  <h2>Coaching Recommendations</h2>
  <ul>
    ${report.coachingRecommendations.map((rec) => `<li>${rec}</li>`).join("")}
  </ul>
</body>
</html>
`;
}

/**
 * Escape CSV special characters
 */
function escapeCSV(value: any): string {
  const str = String(value || "");
  if (str.includes(",") || str.includes('"') || str.includes("\n")) {
    return `"${str.replace(/"/g, '""')}"`;
  }
  return str;
}

/**
 * Download file
 */
export function downloadFile(
  content: string,
  filename: string,
  mimeType: string = "text/plain"
): void {
  if (typeof window === "undefined") return; // Server-side check

  const blob = new Blob([content], { type: mimeType });
  const url = window.URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  window.URL.revokeObjectURL(url);
}

/**
 * Export complete report
 */
export function exportReport(
  report: AuditReport,
  errorLog: any[],
  options: ExportOptions
): void {
  let content = "";
  let filename = options.filename || "audit-report";
  let mimeType = "text/plain";

  switch (options.format) {
    case "JSON":
      content = exportAsJSON(report);
      filename += ".json";
      mimeType = "application/json";
      break;
    case "CSV":
      content = exportErrorLogAsCSV(errorLog);
      filename += ".csv";
      mimeType = "text/csv";
      break;
    case "HTML":
      content = exportAsHTML(report);
      filename += ".html";
      mimeType = "text/html";
      break;
    case "PDF":
      // PDF export would require a PDF library like pdfkit or jspdf
      console.warn("PDF export requires additional setup");
      return;
  }

  if (typeof window !== "undefined") {
    downloadFile(content, filename, mimeType);
  }
}
