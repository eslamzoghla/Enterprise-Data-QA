/**
 * Complete Type Definitions for Enterprise QA Engine
 */

export enum ErrorType {
  // Structural Errors (12 types)
  MissingSheet = "MissingSheet",
  ExtraSheet = "ExtraSheet",
  MissingColumn = "MissingColumn",
  ExtraColumn = "ExtraColumn",
  MissingRow = "MissingRow",
  ExtraRow = "ExtraRow",
  TableMerge = "TableMerge",
  TableSplit = "TableSplit",
  LocalRowMisalignment = "LocalRowMisalignment",
  LocalColumnMisalignment = "LocalColumnMisalignment",
  RowShift = "RowShift",
  ColumnShift = "ColumnShift",

  // Cell-Level Errors
  MissingValue = "MissingValue",
  ExtraValue = "ExtraValue",
  MissingCell = "MissingCell",
  ExtraCell = "ExtraCell",

  // Range Errors (3 types)
  RangeInversionError = "RangeInversionError",
  RangeBoundaryError = "RangeBoundaryError",
  RangeRepresentationError = "RangeRepresentationError",

  // Digit Errors (4 types)
  MissingDigit = "MissingDigit",
  ExtraDigit = "ExtraDigit",
  DigitSubstitution = "DigitSubstitution",
  DigitTransposition = "DigitTransposition",

  // Numeric Errors (2 types)
  NumericDifference = "NumericDifference",
  MajorNumericError = "MajorNumericError",

  // Text Errors (3 types)
  TextTypo = "TextTypo",
  MajorTextDifference = "MajorTextDifference",
  TextDifference = "TextDifference",

  // Header Error
  HeaderMismatch = "HeaderMismatch"
}

export enum Severity {
  Critical = "Critical",
  High = "High",
  Medium = "Medium",
  Low = "Low"
}

export interface CellValue {
  raw: any;
  formatted: string;
  normalized: string;
  type: "string" | "number";
}

export interface SheetData {
  name: string;
  maxRow: number;
  maxCol: number;
  cells: Record<string, CellValue>;
}

export interface WorkbookData {
  fileName: string;
  sheets: Record<string, SheetData>;
}

export interface QAConfig {
  numericMajorVarianceThreshold: number;
  numericMajorAbsoluteThreshold: number;
  arabicComparisonMode: string;
  minimumShiftCells: number;
  shiftDetectionThreshold: number;
  employeeName: string;
  projectName: string;
  evaluationDate: string;
  numericTolerance: number;
  numericToleranceMode: string;
  shiftConfidenceScore: boolean;
  headerPenalty: number;
  strictMode: string;
}

export interface ErrorLogEntry {
  sheet: string;
  cell: string;
  rowIndex: number;
  colIndex: number;
  employeeValue: string;
  reviewerValue: string;
  normalizedEmployeeValue: string;
  normalizedReviewerValue: string;
  similarity: number;
  errorType: ErrorType;
  severity: Severity;
  penalty: number;
  suppressed?: boolean;
  suppressionReason?: string | null;
  notes: string;
}

export interface ShiftEvent {
  type: string;
  sheet: string;
  spanStart: number;
  spanEnd: number;
  detail: string;
  affectedCellsCount: number;
}

export interface RootCauseData {
  missingValuesPct: number;
  numericErrorsPct: number;
  textErrorsPct: number;
  rangeErrorsPct: number;
  shiftErrorsPct: number;
  headerErrorsPct: number;
}

export interface PatternData {
  repeatedNumericErrors: string[];
  copyPasteErrors: string[];
  errorClusters: string[];
  sheetConcentrations: string[];
  shiftEvents: ShiftEvent[];
}

export interface MetricsData {
  comparedCells: number;
  totalErrors: number;
  baseAccuracy: number;
  finalGrade: string;
  totalPenaltyPoints: number;
  errorRatePer10k: number;
  reviewerWorkloadIndex: number;
}

export interface AnalysisResult {
  metrics: MetricsData;
  errorLog: ErrorLogEntry[];
  rootCause: RootCauseData;
  patterns: PatternData;
  coachingRecommendations: string[];
  report: any;
  gridMetrics: any;
  timestamp: string;
}
