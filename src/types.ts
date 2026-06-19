/**
 * Types and interfaces for the Enterprise Excel QA Engine.
 */

export enum ErrorType {
  // Structural Errors
  MissingValue = "Missing Value",
  ExtraValue = "Extra Value",
  RowShift = "Row Shift",
  ColumnShift = "Column Shift",
  MissingRow = "Missing Row",
  ExtraRow = "Extra Row",
  MissingColumn = "Missing Column",
  ExtraColumn = "Extra Column",
  MissingCell = "Missing Cell",
  ExtraCell = "Extra Cell",
  LocalColumnMisalignment = "Local Column Misalignment",
  LocalRowMisalignment = "Local Row Misalignment",
  TableMerge = "Table Merge",
  TableSplit = "Table Split",

  // New Structuring Categories
  ExtraTable = "Extra Table",
  MissingTable = "Missing Table",
  ExtraColumns = "Extra Columns",
  MissingColumns = "Missing Columns",

  // Numeric Errors
  MissingDigit = "Missing Digit",
  ExtraDigit = "Extra Digit",
  DigitSubstitution = "Digit Substitution",
  DigitTransposition = "Digit Transposition",
  NumericDifference = "Numeric Difference", // All remaining numeric mismatches
  MajorNumericError = "Major Numeric Error", // Override based on variance > 20% or absolute threshold

  // Text Errors
  TextTypo = "Text Typo", // Similarity >= 90% after Arabic normalization
  MajorTextDifference = "Major Text Difference", // Similarity < 90%
  TextDifference = "Text Difference", // Fallback text mismatch

  // Sequence / Range Errors
  RangeInversionError = "Range Inversion Error",
  RangeBoundaryError = "Range Boundary Error",
  RangeRepresentationError = "Range Representation Error",
}

export enum Severity {
  Critical = "Critical",
  High = "High",
  Medium = "Medium",
  Low = "Low",
}

export interface QAConfig {
  numericMajorVarianceThreshold: number; // e.g., 0.20 for 20%
  numericMajorAbsoluteThreshold: number; // e.g., 5.0
  arabicComparisonMode: "STANDARD" | "NONE";
  minimumShiftCells: number; // e.g., 20
  shiftDetectionThreshold: number; // e.g., 0.80 for 80%
  employeeName: string;
  projectName: string;
  evaluationDate: string;
  numericTolerance: number; // e.g., 0.01 for 1%
  numericToleranceMode: "PERCENTAGE" | "ABSOLUTE";
  shiftConfidenceScore: boolean;
  headerPenalty: number; // e.g., 3
  strictMode: "AUTO" | "ON" | "OFF";

  // New Penalty Coefficients
  extraTableCoefficient: number;
  missingTableCoefficient: number;
  extraColumnCoefficient: number;
  missingColumnCoefficient: number;
  extraRowCoefficient: number;
  missingRowCoefficient: number;
  numericDifferenceCoefficient: number;
  textDifferenceCoefficient: number;
  emptyCellDifferenceCoefficient: number;
}

export interface CellValue {
  raw: any;
  formatted: string;
  normalized: string;
  type: "string" | "number" | "boolean" | "date" | "empty";
}

export interface VirtualSegment {
  originalSheetName: string;
  virtualStartRow: number;
  virtualEndRow: number;
  originalStartRow: number;
}

export interface SheetGrid {
  name: string;
  maxRow: number;
  maxCol: number;
  // Map of "row,col" -> CellValue
  cells: Record<string, CellValue>;
  virtualSegments?: VirtualSegment[];
}

export interface WorkbookData {
  fileName: string;
  sheets: Record<string, SheetGrid>;
}

export interface ErrorLogEntry {
  sheet: string;
  cell: string; // e.g., "A1" or "C5" (can be zero-indexed row/col indicator, let's use readable coordinate like "B5" or "Row 5, Col B")
  rowIndex: number;
  colIndex: number;
  employeeValue: string;
  reviewerValue: string;
  normalizedEmployeeValue: string;
  normalizedReviewerValue: string;
  similarity: number; // percentage, or 100 for match, or 0
  errorType: ErrorType;
  severity: Severity;
  penalty: number;
  notes: string;
}

export interface ShiftEvent {
  sheetName: string;
  type: "row" | "column";
  offset: number;
  spanStart: number;
  spanEnd: number;
  detail: string;
  affectedCellsCount: number;
}

export interface QAMetrics {
  comparedCells: number;
  totalErrors: number;
  totalPenaltyPoints: number;
  baseAccuracy: number; // %
  weightedAccuracy: number; // %
  errorRatePer10k: number;
  reviewerWorkloadIndex: number; // correction burden metric
  finalGrade: string;

  // New Audit Scoring fields
  structuralPenalty: number;
  dataPenalty: number;
  totalPenalty: number;
  structuralScore: number;
  dataScore: number;
  finalAuditScore: number;

  // Audit Category Counts and contributions
  extraTablesCount: number;
  missingTablesCount: number;
  extraColumnsCount: number;
  missingColumnsCount: number;
  extraRowsCount: number;
  missingRowsCount: number;
  numericDifferencesCount: number;
  textDifferencesCount: number;
  emptyCellDifferencesCount: number;
}

export interface RootCauseStats {
  numericErrorsPct: number;
  missingValuesPct: number;
  textErrorsPct: number;
  shiftErrorsPct: number;
  rangeErrorsPct: number;
  headerErrorsPct: number;
}

export interface PatternFindings {
  repeatedNumericErrors: string[];
  copyPasteErrors: string[];
  errorClusters: string[];
  sheetConcentrations: string[];
  shiftEvents: ShiftEvent[];
}

export interface AnalysisResult {
  config: QAConfig;
  metrics: QAMetrics;
  rootCause: RootCauseStats;
  errorLog: ErrorLogEntry[];
  patterns: PatternFindings;
  coachingRecommendations: string[];
  summarizedAIReview?: string; // Optional field populated by Gemini server-side route
  virtualSheets?: Record<string, SheetGrid>;
}

export interface TableMergeSplitResult {
    detected: boolean;
    employeeTables: string[];
    reviewerTable: string;
    similarity: number;
    type: "MERGE" | "SPLIT";
}