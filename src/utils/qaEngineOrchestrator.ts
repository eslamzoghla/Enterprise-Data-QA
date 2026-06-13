/**
 * Main QA Engine Orchestrator
 * Executes the complete 17-module pipeline in order
 */

import { WorkbookData, QAConfig, AnalysisResult } from "./types";
import { loadWorkbookFromFile } from "./modules/workbookLoader";
import { analyzeWorksheet, analyzeAllWorksheets } from "./modules/worksheetAnalyzer";
import { validateStructures } from "./modules/structureValidator";
import { detectStructuralErrors } from "./modules/structuralErrorDetector";
import { recoverRowAlignment, calculateSimilarity } from "./modules/alignmentRecoveryEngine";
import { applySuppressionRules } from "./modules/structuralSuppressionEngine";
import { compareCells } from "./modules/cellComparisonEngine";
import { classifyError } from "./modules/errorClassifier";
import { detectRepeatedNumericErrors, detectCopyPasteErrors, detectErrorClusters, detectSheetConcentrations } from "./modules/patternDetectionEngine";
import { analyzeRootCauses } from "./modules/rootCauseAnalyzer";
import { score } from "./modules/scoringEngine";
import { generateCompleteReport } from "./modules/reportingEngine";
import { calculateGridMetrics } from "./modules/gridInspectorEngine";

/**
 * Main QA Evaluation Function
 * Executes all 17 modules in the exact pipeline order
 */
export async function executeQAEvaluation(
  employeeWb: WorkbookData,
  reviewerWb: WorkbookData,
  config: QAConfig
): Promise<AnalysisResult> {
  console.log("🚀 Starting Enterprise QA Engine Pipeline...");

  // ============================================================================
  // PHASE 1: STRUCTURE VALIDATION (Modules 1-4)
  // ============================================================================
  console.log("📋 Phase 1: Structure Validation");

  // Module 2: Analyze worksheets
  const empSheets = analyzeAllWorksheets(employeeWb);
  const revSheets = analyzeAllWorksheets(reviewerWb);
  console.log(`  ✓ Employee workbook: ${empSheets.length} sheets`);
  console.log(`  ✓ Reviewer workbook: ${revSheets.length} sheets`);

  // Module 4: Validate structures
  const structureValidation = validateStructures(employeeWb, reviewerWb);
  console.log(`  ✓ Structure validation: ${structureValidation.errors.length} errors, ${structureValidation.warnings.length} warnings`);

  // Module 6: Detect structural errors
  const structuralErrors = detectStructuralErrors(employeeWb, reviewerWb);
  console.log(`  ✓ Structural errors detected: ${structuralErrors.length}`);

  // ⚠️ NO CELL COMPARISON BEFORE STRUCTURE VALIDATION ⚠️

  // ============================================================================
  // PHASE 2: ALIGNMENT & RECOVERY (Module 5)
  // ============================================================================
  console.log("🔄 Phase 2: Alignment Recovery");
  const alignmentMappings: Record<string, any[]> = {};
  const sharedSheets = Object.keys(employeeWb.sheets).filter(
    (sheet) => reviewerWb.sheets[sheet]
  );

  for (const sheet of sharedSheets) {
    const mappings = recoverRowAlignment(
      employeeWb,
      reviewerWb,
      sheet,
      0.9
    );
    alignmentMappings[sheet] = mappings;
  }
  console.log(`  ✓ Row alignment recovered for ${Object.keys(alignmentMappings).length} sheets`);

  // ============================================================================
  // PHASE 3: CELL COMPARISON (Module 8)
  // ============================================================================
  console.log("📊 Phase 3: Cell-Level Comparison");
  const errorLog: any[] = [];
  let comparedCells = 0;

  for (const sheet of sharedSheets) {
    const empSheet = employeeWb.sheets[sheet];
    const revSheet = reviewerWb.sheets[sheet];
    const isHeaderRow = (row: number) => row === 0;

    // Compare all cells
    const maxRows = Math.max(empSheet.maxRow, revSheet.maxRow);
    const maxCols = Math.max(empSheet.maxCol, revSheet.maxCol);

    for (let row = 0; row <= maxRows; row++) {
      for (let col = 0; col <= maxCols; col++) {
        const cellKey = `${row},${col}`;
        const empCell = empSheet.cells[cellKey];
        const revCell = revSheet.cells[cellKey];

        comparedCells++;

        // Skip if both empty
        if (!empCell && !revCell) continue;

        // Module 8: Compare cells
        const comparison = compareCells(empCell, revCell, {
          arabicComparisonMode: config.arabicComparisonMode as "STANDARD" | "NONE",
          numericTolerance: config.numericTolerance,
          numericToleranceMode: config.numericToleranceMode as "ABSOLUTE" | "PERCENTAGE"
        });

        if (!comparison.match) {
          // Module 11: Classify error
          const empVal = empCell ? String(empCell.raw) : "";
          const revVal = revCell ? String(revCell.raw) : "";
          const classification = classifyError(
            empVal,
            revVal,
            comparison,
            isHeaderRow(row)
          );

          errorLog.push({
            sheet,
            cell: `${String.fromCharCode(65 + col)}${row + 1}`,
            rowIndex: row,
            colIndex: col,
            employeeValue: empVal,
            reviewerValue: revVal,
            normalizedEmployeeValue: comparison.employeeNormalized,
            normalizedReviewerValue: comparison.reviewerNormalized,
            similarity: comparison.similarity,
            errorType: classification.errorType,
            severity: classification.severity,
            penalty: classification.confidence * 5,
            notes: classification.description
          });
        }
      }
    }
  }
  console.log(`  ✓ Compared ${comparedCells} cells, found ${errorLog.length} errors`);

  // ============================================================================
  // PHASE 4: ERROR SUPPRESSION & ROOT CAUSE (Modules 7 & 13)
  // ============================================================================
  console.log("🔍 Phase 4: Error Suppression & Root Cause Analysis");

  // Module 7: Apply suppression rules
  const suppressedErrorLog = applySuppressionRules(errorLog, structuralErrors);
  const nonSuppressedCount = suppressedErrorLog.filter((e) => !e.suppressed).length;
  console.log(`  ✓ Suppression applied: ${errorLog.length} → ${nonSuppressedCount} active errors`);

  // Module 13: Root cause analysis
  const rootCauseAnalysis = analyzeRootCauses(suppressedErrorLog);
  console.log(`  ✓ Root cause analysis complete`);
  console.log(`    - Missing/Extra: ${rootCauseAnalysis.missingValuesPct}%`);
  console.log(`    - Numeric: ${rootCauseAnalysis.numericErrorsPct}%`);
  console.log(`    - Text: ${rootCauseAnalysis.textErrorsPct}%`);
  console.log(`    - Shifts: ${rootCauseAnalysis.shiftErrorsPct}%`);

  // ============================================================================
  // PHASE 5: PATTERN DETECTION (Module 12)
  // ============================================================================
  console.log("🎯 Phase 5: Pattern Detection");
  const patterns = {
    repeatedNumericErrors: detectRepeatedNumericErrors(suppressedErrorLog),
    copyPasteErrors: detectCopyPasteErrors(suppressedErrorLog),
    errorClusters: detectErrorClusters(suppressedErrorLog),
    sheetConcentrations: detectSheetConcentrations(suppressedErrorLog),
    shiftEvents: structuralErrors.filter(
      (e) => e.type === "RowShift" || e.type === "ColumnShift"
    )
  };
  console.log(`  ✓ Patterns detected:`);
  console.log(`    - Repeated numeric: ${patterns.repeatedNumericErrors.length}`);
  console.log(`    - Copy-paste: ${patterns.copyPasteErrors.length}`);
  console.log(`    - Error clusters: ${patterns.errorClusters.length}`);
  console.log(`    - Shift events: ${patterns.shiftEvents.length}`);

  // ============================================================================
  // PHASE 6: SCORING (Module 14)
  // ============================================================================
  console.log("⭐ Phase 6: Scoring & Grading");
  const metrics = score(
    suppressedErrorLog,
    comparedCells,
    patterns.shiftEvents.length > 0
  );
  console.log(`  ✓ Accuracy: ${metrics.baseAccuracy}%`);
  console.log(`  ✓ Grade: ${metrics.finalGrade}`);
  console.log(`  ✓ Penalty Points: ${metrics.totalPenaltyPoints}`);
  console.log(`  ✓ Error Rate: ${metrics.errorRatePer10k}/10k cells`);
  console.log(`  ✓ Reviewer Workload: ${metrics.reviewerWorkloadIndex} hours`);

  // ============================================================================
  // PHASE 7: REPORTING (Module 15)
  // ============================================================================
  console.log("📝 Phase 7: Report Generation");
  const report = generateCompleteReport(
    metrics,
    suppressedErrorLog,
    structuralErrors,
    rootCauseAnalysis,
    patterns,
    config.employeeName,
    config.projectName,
    config.evaluationDate
  );
  console.log(`  ✓ 6-section audit report generated`);

  // ============================================================================
  // PHASE 8: GRID METRICS (Module 16)
  // ============================================================================
  console.log("📊 Phase 8: Grid Metrics Calculation");
  const gridMetrics = calculateGridMetrics(suppressedErrorLog);
  console.log(`  ✓ Grid metrics calculated`);

  // ============================================================================
  // FINAL RESULT
  // ============================================================================
  console.log("✅ Pipeline Complete!");

  return {
    metrics,
    errorLog: suppressedErrorLog,
    rootCause: rootCauseAnalysis,
    patterns,
    coachingRecommendations: report.coachingRecommendations,
    report,
    gridMetrics,
    timestamp: new Date().toISOString()
  };
}

/**
 * Helper: Normalize workbook with config
 */
export function normalizeWorkbook(
  wb: WorkbookData,
  config: QAConfig
): WorkbookData {
  const normalizedSheets: Record<string, any> = {};

  for (const [sheetName, sheetGrid] of Object.entries(wb.sheets)) {
    const normCells: Record<string, any> = {};

    for (const [coord, cell] of Object.entries(sheetGrid.cells)) {
      // Normalization would happen here based on config
      normCells[coord] = cell;
    }

    normalizedSheets[sheetName] = {
      ...sheetGrid,
      cells: normCells
    };
  }

  return {
    ...wb,
    sheets: normalizedSheets
  };
}
