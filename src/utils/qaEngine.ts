/**
 * Upgraded Enterprise Excel QA Deterministic Comparison & Shift Detection Engine.
 * Tailored precisely to audit-grade requirements.
 */

import {
  ErrorType,
  Severity,
  CellValue,
  SheetGrid,
  WorkbookData,
  ErrorLogEntry,
  ShiftEvent,
  QAMetrics,
  RootCauseStats,
  PatternFindings,
  QAConfig,
  AnalysisResult
} from "../types.ts";

/**
 * Normalizes leading/trailing/multiple internal spaces and line breaks.
 */
export function normalizeTextSpaces(text: string): string {
  if (!text) return "";
  return text
    .replace(/\r?\n/g, " ") // Normalize line breaks to a space
    .replace(/\s+/g, " ")   // Replace multiple spaces with a single space
    .trim();                // Trim leading/trailing whitespace
}

/**
 * Standard Arabic Text Normalization (if ArabicComparisonMode = STANDARD)
 */
export function normalizeArabicText(text: string): string {
  if (!text) return "";
  let norm = text;

  // Remove Arabic diacritics (Harakat)
  norm = norm.replace(/[\u064B-\u0652]/g, "");

  // Remove Tatweel/Kashida (ـ)
  norm = norm.replace(/\u0640/g, "");

  // Convert Alef variations (أ, إ, آ) -> plain Alef (ا)
  norm = norm.replace(/[أإآ]/g, "ا");

  // Convert Dotless Ya (ى) -> Ya (ي)
  norm = norm.replace(/ى/g, "ي");

  // Convert Waw with Hamza (ؤ) -> Waw (و)
  norm = norm.replace(/ؤ/g, "و");

  // Convert Ya with Hamza (ئ) -> Ya (ي)
  norm = norm.replace(/ئ/g, "ي");

  // Replace Te Marbuta (ة) with Heh (ه) to simplify spelling checks
  norm = norm.replace(/ة/g, "ه");

  return norm;
}

/**
 * Clean numeric string helpers (removes thousand commas, handles leading zeros)
 */
export function normalizeNumericString(val: string | number): { isNumeric: boolean; value: number | null; cleanedStr: string } {
  if (typeof val === "number") {
    return { isNumeric: true, value: val, cleanedStr: String(val) };
  }
  const trimmed = val.trim();
  if (!trimmed) {
    return { isNumeric: false, value: null, cleanedStr: "" };
  }

  // Strip standard grouping commas (e.g. 1,000 -> 1000)
  const noCommas = trimmed.replace(/,/g, "");

  // Check if it matches a standard float/int representation
  const numVal = Number(noCommas);
  if (!isNaN(numVal) && noCommas.match(/^-?\d+(\.\d+)?$/)) {
    return { isNumeric: true, value: numVal, cleanedStr: String(numVal) };
  }

  return { isNumeric: false, value: null, cleanedStr: trimmed };
}

/**
 * Standardizes common Date representations into YYYY-MM-DD
 */
export function normalizeDate(val: string): { isDate: boolean; formatted: string } {
  const cleaned = val.trim();
  const timestamp = Date.parse(cleaned);
  if (!isNaN(timestamp)) {
    const d = new Date(timestamp);
    const year = d.getFullYear();
    const month = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    
    // Dates usually have delimiters like /, -, or spaces, prevent raw large numbers being false dates
    if (cleaned.includes("/") || cleaned.includes("-") || cleaned.includes(",")) {
      return { isDate: true, formatted: `${year}-${month}-${day}` };
    }
  }
  return { isDate: false, formatted: cleaned };
}

/**
 * Standardizes time strings (e.g. 14:30:00 vs 14:30)
 */
export function normalizeTime(val: string): { isTime: boolean; formatted: string } {
  const cleaned = val.trim();
  const timeRegex = /^([0-1]?[0-9]|2[0-3]):([0-5][0-9])(:([0-5][0-9]))?\s*(AM|PM)?$/i;
  if (timeRegex.test(cleaned)) {
    return { isTime: true, formatted: cleaned.toLowerCase() };
  }
  return { isTime: false, formatted: cleaned };
}

/**
 * Computes Levenshtein Distance between s1 and s2
 */
export function levenshteinDistance(s1: string, s2: string): number {
  if (s1.length === 0) return s2.length;
  if (s2.length === 0) return s1.length;

  const matrix: number[][] = Array.from({ length: s2.length + 1 }, (_, i) => [i]);
  for (let j = 1; j <= s1.length; j++) {
    matrix[0][j] = j;
  }

  for (let i = 1; i <= s2.length; i++) {
    for (let j = 1; j <= s1.length; j++) {
      if (s2[i - 1] === s1[j - 1]) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(
          matrix[i - 1][j - 1] + 1, // substitution
          matrix[i][j - 1] + 1,     // insertion
          matrix[i - 1][j] + 1      // deletion
        );
      }
    }
  }
  return matrix[s2.length][s1.length];
}

/**
 * Calculates similarity metric as a percentage (0 to 100)
 */
export function calculateTextSimilarity(s1: string, s2: string): number {
  const maxLength = Math.max(s1.length, s2.length);
  if (maxLength === 0) return 100;
  const distance = levenshteinDistance(s1, s2);
  return parseFloat(((1 - distance / maxLength) * 100).toFixed(2));
}

/**
 * Full normalization of any value based on configurations
 */
export function getNormalizedValue(raw: any, config: QAConfig): CellValue {
  if (raw === undefined || raw === null || raw === "") {
    return { raw: null, formatted: "", normalized: "", type: "empty" };
  }

  const str = String(raw);
  const spacedNorm = normalizeTextSpaces(str);

  // Check numeric
  const numNorm = normalizeNumericString(spacedNorm);
  if (numNorm.isNumeric) {
    return {
      raw,
      formatted: spacedNorm,
      normalized: numNorm.cleanedStr,
      type: "number"
    };
  }

  // Check DateTime
  const dateNorm = normalizeDate(spacedNorm);
  if (dateNorm.isDate) {
    return {
      raw,
      formatted: spacedNorm,
      normalized: dateNorm.formatted,
      type: "date"
    };
  }

  const timeNorm = normalizeTime(spacedNorm);
  if (timeNorm.isTime) {
    return {
      raw,
      formatted: spacedNorm,
      normalized: timeNorm.formatted,
      type: "date"
    };
  }

  // Treat as string comparison
  let textNormalized = spacedNorm;
  if (config.arabicComparisonMode === "STANDARD") {
    const hasArabic = /[\u0600-\u06FF]/.test(spacedNorm);
    if (hasArabic) {
      textNormalized = normalizeArabicText(spacedNorm);
    }
  }

  return {
    raw,
    formatted: spacedNorm,
    normalized: textNormalized,
    type: "string"
  };
}

/**
 * Auto-detect sheets exclusions
 */
export function shouldExcludeSheet(sheetName: string, maxRow: number, cellsCount: number): boolean {
  // Do not automatically exclude sheets by name keyword or low row counts. Evaluate everything that is loaded.
  return false;
}

/**
 * Auto-detect header rows by row index, styling candidate heuristics or string dominance.
 */
export function detectHeaderRows(sheet: SheetGrid): Set<number> {
  const headers = new Set<number>();
  if (sheet.maxRow < 0) return headers;

  const maxInspect = Math.min(4, sheet.maxRow);
  for (let r = 0; r <= maxInspect; r++) {
    let stringCount = 0;
    let totalCount = 0;

    for (let c = 0; c <= sheet.maxCol; c++) {
      const cell = sheet.cells[`${r},${c}`];
      if (cell) {
        totalCount++;
        if (cell.type === "string") {
          stringCount++;
        }
      }
    }

    // Default row 0 is header. For rows 1-4, if dominantly string context, consider as sub-header.
    if (r === 0) {
      headers.add(r);
    } else if (totalCount > 0 && stringCount / totalCount >= 0.6) {
      headers.add(r);
    }
  }
  return headers;
}

/**
 * Auto-detect if strict evaluation mode is active based on file or sheet meta indicators
 */
export function isStrictModeActive(workbook: WorkbookData, config: QAConfig): boolean {
  if (config.strictMode === "ON") return true;
  if (config.strictMode === "OFF") return false;

  const fName = (workbook.fileName || "").toLowerCase();
  if (
    fName.includes("census") ||
    fName.includes("financial") ||
    fName.includes("budget") ||
    fName.includes("demographic") ||
    fName.includes("survey") ||
    fName.includes("stat") ||
    fName.includes("tax") ||
    fName.includes("audit")
  ) {
    return true;
  }

  for (const sheetName of Object.keys(workbook.sheets)) {
    const sName = sheetName.toLowerCase();
    if (
      sName.includes("census") ||
      sName.includes("financial") ||
      sName.includes("budget") ||
      sName.includes("demographic") ||
      sName.includes("survey") ||
      sName.includes("stat")
    ) {
      return true;
    }
  }

  return false;
}

/**
 * Robust parsing of Ranges/Sequences e.g., "64/65", "1990-1995"
 */
interface RangeParse {
  isValid: boolean;
  left: string;
  right: string;
  delimiter: string;
}

export function parseRangeOrSequence(val: string): RangeParse {
  const cleaned = val.trim();
  // Support delimiters: slash, dash, en-dash, em-dash
  const regex = /^([a-zA-Z0-9]+)([\/\-–—])([a-zA-Z0-9]+)$/;
  const match = cleaned.match(regex);
  if (match) {
    return {
      isValid: true,
      left: match[1],
      right: match[3],
      delimiter: match[2]
    };
  }
  return { isValid: false, left: "", right: "", delimiter: "" };
}

/**
 * Check if two ranges represent matching coordinates but representational differences
 */
export function isEquivalentRangeRepresentational(r: RangeParse, e: RangeParse): boolean {
  if (!r.isValid || !e.isValid) return false;

  const cleanNum = (str: string) => {
    const parsed = parseInt(str);
    return isNaN(parsed) ? str : (parsed % 100);
  };

  const rL_num = cleanNum(r.left);
  const rR_num = cleanNum(r.right);
  const eL_num = cleanNum(e.left);
  const eR_num = cleanNum(e.right);

  // If underlying modulo 100 values are the same, but the formatting/literals or delimiter differ
  if (rL_num === eL_num && rR_num === eR_num) {
    return true;
  }
  return false;
}

/**
 * High quality digit error checks for numbers or alpha-numeric extracted components
 */
export function classifyDigitError(empDigits: string, revDigits: string): ErrorType | null {
  if (empDigits === revDigits) return null;
  if (!empDigits || !revDigits) return null;

  const lenE = empDigits.length;
  const lenR = revDigits.length;

  // 1. Missing Digit (employee missing 1 digit)
  if (lenE === lenR - 1) {
    let j = 0;
    for (let i = 0; i < lenR && j < lenE; i++) {
      if (revDigits[i] === empDigits[j]) j++;
    }
    if (j === lenE) return ErrorType.MissingDigit;
  }

  // 2. Extra Digit (employee has 1 extra digit)
  if (lenE === lenR + 1) {
    let j = 0;
    for (let i = 0; i < lenE && j < lenR; i++) {
      if (empDigits[i] === revDigits[j]) j++;
    }
    if (j === lenR) return ErrorType.ExtraDigit;
  }

  // 3. Digit Transposition (exact same digits adjacent transpose)
  if (lenE === lenR) {
    let mismatches: number[] = [];
    for (let i = 0; i < lenE; i++) {
      if (empDigits[i] !== revDigits[i]) {
        mismatches.push(i);
      }
    }
    if (mismatches.length === 2 && mismatches[1] - mismatches[0] === 1) {
      const idx1 = mismatches[0];
      const idx2 = mismatches[1];
      if (
        empDigits[idx1] === revDigits[idx2] &&
        empDigits[idx2] === revDigits[idx1]
      ) {
        return ErrorType.DigitTransposition;
      }
    }

    // 4. Digit Substitution (exact length, exactly 1 mismatch)
    if (mismatches.length === 1) {
      return ErrorType.DigitSubstitution;
    }
  }

  return ErrorType.NumericDifference;
}

/**
 * Extracts and compares numeric digits from mixed alpha-numeric strings
 */
export function analyzeMixedAlphaNumeric(
  empVal: string,
  revVal: string
): { digitError: ErrorType | null; similarity: number } {
  const empDigits = empVal.replace(/\D/g, "");
  const revDigits = revVal.replace(/\D/g, "");

  const dError = classifyDigitError(empDigits, revDigits);
  const similarity = calculateTextSimilarity(empVal, revVal);

  return { digitError: dError, similarity };
}

/**
 * Converts zero-based column index to Excel column string (e.g., 0 -> A, 27 -> AB)
 */
export function getColLetter(colIndex: number): string {
  let temp = colIndex;
  let letter = "";
  while (temp >= 0) {
    letter = String.fromCharCode((temp % 26) + 65) + letter;
    temp = Math.floor(temp / 26) - 1;
  }
  return letter;
}

/**
 * Executes high fidelity Column Shift Detection on grids with confidence level
 */
export function detectShifts(
  empGrid: SheetGrid,
  revGrid: SheetGrid,
  config: QAConfig
): {
  shiftedCoords: Record<string, "row" | "column">;
  events: ShiftEvent[];
} {
  const shiftedCoords: Record<string, "row" | "column"> = {};
  const events: ShiftEvent[] = [];

  const maxRow = Math.max(empGrid.maxRow, revGrid.maxRow);
  const maxCol = Math.max(empGrid.maxCol, revGrid.maxCol);

  const minShiftCells = config.minimumShiftCells;
  const shiftThreshold = config.shiftDetectionThreshold;

  // Search standard offsets from -3 to +3, excluding 0
  const offsets = [-3, -2, -1, 1, 2, 3];

  const getNorm = (grid: SheetGrid, r: number, c: number) => {
    return grid.cells[`${r},${c}`]?.normalized || "";
  };

  const getCellType = (grid: SheetGrid, r: number, c: number) => {
    return grid.cells[`${r},${c}`]?.type || "empty";
  };

  // --- 2. Detect Column Shift ---
  for (const dc of offsets) {
    let contiguousCols: number[] = [];

    for (let c = 0; c <= maxCol; c++) {
      let matches = 0;
      let compared = 0;

      for (let r = 0; r <= maxRow; r++) {
        const empVal = getNorm(empGrid, r, c);
        const revVal = getNorm(revGrid, r, c + dc);

        const empType = getCellType(empGrid, r, c);
        const revType = getCellType(revGrid, r, c + dc);

        if (empType !== "empty" || revType !== "empty") {
          compared++;
          if (empVal === revVal) {
            matches++;
          }
        }
      }

      if (compared > 0 && matches / compared >= shiftThreshold) {
        contiguousCols.push(c);
      } else {
        if (contiguousCols.length > 0) {
          evaluateColGroup(contiguousCols, dc);
          contiguousCols = [];
        }
      }
    }
    if (contiguousCols.length > 0) {
      evaluateColGroup(contiguousCols, dc);
    }

    function evaluateColGroup(cols: number[], offset_c: number) {
      let totalSegmentCells = 0;
      let matchesCount = 0;
      for (const c of cols) {
        for (let r = 0; r <= maxRow; r++) {
          if (getCellType(empGrid, r, c) !== "empty" || getCellType(revGrid, r, c + offset_c) !== "empty") {
            totalSegmentCells++;
            if (getNorm(empGrid, r, c) === getNorm(revGrid, r, c + offset_c)) {
              matchesCount++;
            }
          }
        }
      }

      if (totalSegmentCells >= minShiftCells) {
        const start = cols[0];
        const end = cols[cols.length - 1];

        const startLetter = getColLetter(start);
        const endLetter = getColLetter(end);

        const confidenceScore = totalSegmentCells > 0 ? Math.round((matchesCount / totalSegmentCells) * 100) : 100;

        events.push({
          sheetName: empGrid.name,
          type: "column",
          offset: offset_c,
          spanStart: start,
          spanEnd: end,
          detail: `Column Alignment Shift: Columns ${startLetter} to ${endLetter} shifted horizontally by ${offset_c} space(s) [Confidence: ${confidenceScore}%, cells affected: ${totalSegmentCells}]`,
          affectedCellsCount: totalSegmentCells
        });

        for (const c of cols) {
          for (let r = 0; r <= maxRow; r++) {
            shiftedCoords[`${r},${c}`] = "column";
          }
        }
      }
    }
  }

  return { shiftedCoords, events };
}

export interface AlignmentStep {
  empRow: number | null;
  revRow: number | null;
  type: "match" | "missing_row" | "extra_row";
}

function getRowMatchScoreValue(
  empRowCells: (CellValue | undefined)[],
  revRowCells: (CellValue | undefined)[],
  maxCol: number
): number {
  let matches = 0;
  let totalCompared = 0;
  let bothEmpty = true;

  for (let c = 0; c <= maxCol; c++) {
    const e = empRowCells[c];
    const r = revRowCells[c];

    if (!e && !r) continue;

    bothEmpty = false;
    totalCompared++;
    if (e && r && e.normalized === r.normalized) {
      matches++;
    }
  }

  if (bothEmpty) {
    return 4.0; // Two empty rows align perfectly
  }

  const matchRatio = matches / totalCompared;
  
  if (matchRatio >= 0.5) {
    // Highly similar row pair
    return matchRatio * 15.0 - 5.0; // scales from +2.5 to +10.0
  } else {
    // Dissimilar row pair
    return matchRatio * 15.0 - 8.0; // scales from -8.0 to -0.5
  }
}

export function alignSheetRows(
  empSheet: SheetGrid,
  revSheet: SheetGrid,
  maxCol: number
): AlignmentStep[] {
  const N = empSheet.maxRow + 1;
  const M = revSheet.maxRow + 1;

  if (N === 0) {
    return Array.from({ length: M }, (_, j) => ({
      empRow: null,
      revRow: j,
      type: "missing_row" as const
    }));
  }
  if (M === 0) {
    return Array.from({ length: N }, (_, i) => ({
      empRow: i,
      revRow: null,
      type: "extra_row" as const
    }));
  }

  // Pre-extract cell reference values for O(1) array access without string template interpolation inside DP
  const empRowCells: (CellValue | undefined)[][] = Array.from({ length: N }, (_, r) =>
    Array.from({ length: maxCol + 1 }, (_, c) => empSheet.cells[`${r},${c}`])
  );
  const revRowCells: (CellValue | undefined)[][] = Array.from({ length: M }, (_, r) =>
    Array.from({ length: maxCol + 1 }, (_, c) => revSheet.cells[`${r},${c}`])
  );

  const GAP_PENALTY = -4.0;
  const BAND_WIDTH = 100; // Generous diagonal constraint of 100 rows for O(N * BW) linear time instead of O(N^2)

  const dp: number[][] = Array.from({ length: N + 1 }, () => new Array(M + 1).fill(-1e9));
  const parent: Array<Array<{ i: number; j: number; op: "match" | "extra" | "missing" } | null>> = Array.from(
    { length: N + 1 },
    () => new Array(M + 1).fill(null)
  );

  dp[0][0] = 0;

  for (let i = 1; i <= N; i++) {
    if (i <= BAND_WIDTH) {
      dp[i][0] = dp[i - 1][0] + GAP_PENALTY;
      parent[i][0] = { i: i - 1, j: 0, op: "extra" };
    }
  }
  for (let j = 1; j <= M; j++) {
    if (j <= BAND_WIDTH) {
      dp[0][j] = dp[0][j - 1] + GAP_PENALTY;
      parent[0][j] = { i: 0, j: j - 1, op: "missing" };
    }
  }

  for (let i = 1; i <= N; i++) {
    const minJ = Math.max(1, i - BAND_WIDTH);
    const maxJ = Math.min(M, i + BAND_WIDTH);

    for (let j = minJ; j <= maxJ; j++) {
      const matchScore = getRowMatchScoreValue(empRowCells[i - 1], revRowCells[j - 1], maxCol);
      const scoreDiag = dp[i - 1][j - 1] + matchScore;
      const scoreExtra = dp[i - 1][j] + GAP_PENALTY;
      const scoreMissing = dp[i][j - 1] + GAP_PENALTY;

      let maxVal = scoreDiag;
      let op: "match" | "extra" | "missing" = "match";
      let pi = i - 1, pj = j - 1;

      if (scoreExtra > maxVal) {
        maxVal = scoreExtra;
        op = "extra";
        pi = i - 1;
        pj = j;
      }
      if (scoreMissing > maxVal) {
        maxVal = scoreMissing;
        op = "missing";
        pi = i;
        pj = j - 1;
      }

      dp[i][j] = maxVal;
      parent[i][j] = { i: pi, j: pj, op };
    }
  }

  let iVal = N;
  let jVal = M;
  const steps: AlignmentStep[] = [];

  while (iVal > 0 || jVal > 0) {
    const p = parent[iVal][jVal];
    if (!p) {
      if (iVal > 0 && jVal > 0) {
        steps.push({ empRow: iVal - 1, revRow: jVal - 1, type: "match" });
        iVal--; jVal--;
      } else if (iVal > 0) {
        steps.push({ empRow: iVal - 1, revRow: null, type: "extra_row" });
        iVal--;
      } else {
        steps.push({ empRow: null, revRow: jVal - 1, type: "missing_row" });
        jVal--;
      }
      continue;
    }

    if (p.op === "match") {
      steps.push({ empRow: iVal - 1, revRow: jVal - 1, type: "match" });
    } else if (p.op === "extra") {
      steps.push({ empRow: iVal - 1, revRow: null, type: "extra_row" });
    } else {
      steps.push({ empRow: null, revRow: jVal - 1, type: "missing_row" });
    }
    iVal = p.i;
    jVal = p.j;
  }

  steps.reverse();
  return steps;
}

export interface ColumnAlignmentStep {
  empCol: number | null;
  revCol: number | null;
  type: "match" | "missing_column" | "extra_column";
}

function getColumnMatchScoreValue(
  empColCells: (CellValue | undefined)[],
  revColCells: (CellValue | undefined)[],
  maxRow: number
): number {
  let matches = 0;
  let totalCompared = 0;
  let bothEmpty = true;

  for (let r = 0; r <= maxRow; r++) {
    const e = empColCells[r];
    const rCell = revColCells[r];

    if (!e && !rCell) continue;

    bothEmpty = false;
    totalCompared++;
    if (e && rCell && e.normalized === rCell.normalized) {
      matches++;
    }
  }

  if (bothEmpty) {
    return 4.0; // Two empty columns align perfectly
  }

  const matchRatio = matches / totalCompared;
  
  if (matchRatio >= 0.5) {
    // Highly similar column pair
    return matchRatio * 15.0 - 5.0; // scales from +2.5 to +10.0
  } else {
    // Dissimilar column pair
    return matchRatio * 15.0 - 8.0; // scales from -8.0 to -0.5
  }
}

export function alignSheetColumns(
  empSheet: SheetGrid,
  revSheet: SheetGrid,
  maxRow: number
): ColumnAlignmentStep[] {
  const N = empSheet.maxCol + 1;
  const M = revSheet.maxCol + 1;

  if (N === 0) {
    return Array.from({ length: M }, (_, j) => ({
      empCol: null,
      revCol: j,
      type: "missing_column" as const
    }));
  }
  if (M === 0) {
    return Array.from({ length: N }, (_, i) => ({
      empCol: i,
      revCol: null,
      type: "extra_column" as const
    }));
  }

  // Pre-extract cell reference values for O(1) array access without string template interpolation inside DP
  const empColCells: (CellValue | undefined)[][] = Array.from({ length: N }, (_, c) =>
    Array.from({ length: maxRow + 1 }, (_, r) => empSheet.cells[`${r},${c}`])
  );
  const revColCells: (CellValue | undefined)[][] = Array.from({ length: M }, (_, c) =>
    Array.from({ length: maxRow + 1 }, (_, r) => revSheet.cells[`${r},${c}`])
  );

  const GAP_PENALTY = -4.0;
  const BAND_WIDTH = 50; // Column shift limit to speed up columns grid matching

  const dp: number[][] = Array.from({ length: N + 1 }, () => new Array(M + 1).fill(-1e9));
  const parent: Array<Array<{ i: number; j: number; op: "match" | "extra" | "missing" } | null>> = Array.from(
    { length: N + 1 },
    () => new Array(M + 1).fill(null)
  );

  dp[0][0] = 0;

  for (let i = 1; i <= N; i++) {
    if (i <= BAND_WIDTH) {
      dp[i][0] = dp[i - 1][0] + GAP_PENALTY;
      parent[i][0] = { i: i - 1, j: 0, op: "extra" };
    }
  }
  for (let j = 1; j <= M; j++) {
    if (j <= BAND_WIDTH) {
      dp[0][j] = dp[0][j - 1] + GAP_PENALTY;
      parent[0][j] = { i: 0, j: j - 1, op: "missing" };
    }
  }

  for (let i = 1; i <= N; i++) {
    const minJ = Math.max(1, i - BAND_WIDTH);
    const maxJ = Math.min(M, i + BAND_WIDTH);

    for (let j = minJ; j <= maxJ; j++) {
      const matchScore = getColumnMatchScoreValue(empColCells[i - 1], revColCells[j - 1], maxRow);
      const scoreDiag = dp[i - 1][j - 1] + matchScore;
      const scoreExtra = dp[i - 1][j] + GAP_PENALTY;
      const scoreMissing = dp[i][j - 1] + GAP_PENALTY;

      let maxVal = scoreDiag;
      let op: "match" | "extra" | "missing" = "match";
      let pi = i - 1, pj = j - 1;

      if (scoreExtra > maxVal) {
        maxVal = scoreExtra;
        op = "extra";
        pi = i - 1;
        pj = j;
      }
      if (scoreMissing > maxVal) {
        maxVal = scoreMissing;
        op = "missing";
        pi = i;
        pj = j - 1;
      }

      dp[i][j] = maxVal;
      parent[i][j] = { i: pi, j: pj, op };
    }
  }

  let iVal = N;
  let jVal = M;
  const steps: ColumnAlignmentStep[] = [];

  while (iVal > 0 || jVal > 0) {
    const p = parent[iVal][jVal];
    if (!p) {
      if (iVal > 0 && jVal > 0) {
        steps.push({ empCol: iVal - 1, revCol: jVal - 1, type: "match" });
        iVal--; jVal--;
      } else if (iVal > 0) {
        steps.push({ empCol: iVal - 1, revCol: null, type: "extra_column" });
        iVal--;
      } else {
        steps.push({ empCol: null, revCol: jVal - 1, type: "missing_column" });
        jVal--;
      }
      continue;
    }

    if (p.op === "match") {
      steps.push({ empCol: iVal - 1, revCol: jVal - 1, type: "match" });
    } else if (p.op === "extra") {
      steps.push({ empCol: iVal - 1, revCol: null, type: "extra_column" });
    } else {
      steps.push({ empCol: null, revCol: jVal - 1, type: "missing_column" });
    }
    iVal = p.i;
    jVal = p.j;
  }

  steps.reverse();
  return steps;
}

/**
 * Consolidates sequential row errors (e.g., MissingRow, ExtraRow, RowShift)
 * on consecutive row indices so that they are calculated as 1 single error instead of N.
 */
function consolidateSequentialRowErrors(errors: ErrorLogEntry[]): ErrorLogEntry[] {
  const result: ErrorLogEntry[] = [];
  
  // Group errors by sheet name
  const sheetGroups: Record<string, ErrorLogEntry[]> = {};
  for (const err of errors) {
    if (!sheetGroups[err.sheet]) {
      sheetGroups[err.sheet] = [];
    }
    sheetGroups[err.sheet].push(err);
  }
  
  for (const sheetName in sheetGroups) {
    const sheetErrors = sheetGroups[sheetName];
    
    // Separate row-level errors and non-row-level errors on this sheet
    const rowLevelErrors: ErrorLogEntry[] = [];
    const otherErrors: ErrorLogEntry[] = [];
    
    for (const err of sheetErrors) {
      const isRowLevel = 
        err.cell.startsWith("Row ") || 
        err.errorType === ErrorType.MissingRow || 
        err.errorType === ErrorType.ExtraRow || 
        err.errorType === ErrorType.RowShift;
      
      if (isRowLevel) {
        rowLevelErrors.push(err);
      } else {
        otherErrors.push(err);
      }
    }
    
    // Sort row-level errors by rowIndex
    rowLevelErrors.sort((a, b) => a.rowIndex - b.rowIndex);
    
    // Group consecutive row-level errors of the same errorType
    const groupedRowErrors: ErrorLogEntry[][] = [];
    let currentGroup: ErrorLogEntry[] = [];
    
    for (const err of rowLevelErrors) {
      if (currentGroup.length === 0) {
        currentGroup.push(err);
      } else {
        const lastErr = currentGroup[currentGroup.length - 1];
        const isConsecutive = err.rowIndex === lastErr.rowIndex + 1;
        const isSameType = err.errorType === lastErr.errorType;
        
        if (isConsecutive && isSameType) {
          currentGroup.push(err);
        } else {
          groupedRowErrors.push(currentGroup);
          currentGroup = [err];
        }
      }
    }
    if (currentGroup.length > 0) {
      groupedRowErrors.push(currentGroup);
    }
    
    // For each group, consolidate into a single ErrorLogEntry
    for (const group of groupedRowErrors) {
      if (group.length === 1) {
        result.push(group[0]);
      } else {
        // Consolidate consecutive rows
        const first = group[0];
        const last = group[group.length - 1];
        
        const startRow = first.rowIndex + 1;
        const endRow = last.rowIndex + 1;
        
        const cellRef = `Rows ${startRow}-${endRow}`;
        
        let combinedEmpVal = "";
        let combinedRevVal = "";
        if (first.errorType === ErrorType.ExtraRow) {
          combinedEmpVal = group.map(g => g.employeeValue).filter(Boolean).join("; ");
        } else if (first.errorType === ErrorType.MissingRow) {
          combinedRevVal = group.map(g => g.reviewerValue).filter(Boolean).join("; ");
        } else {
          combinedEmpVal = group.map(g => g.employeeValue).filter(Boolean).join("; ");
          combinedRevVal = group.map(g => g.reviewerValue).filter(Boolean).join("; ");
        }
        
        // Single error penalty and original severity
        const penalty = first.penalty;
        const severity = first.severity;
        
        const typeLabel = first.errorType === ErrorType.RowShift 
          ? "Consecutive Row Shift" 
          : first.errorType === ErrorType.MissingRow 
            ? "Omitted Rows Block" 
            : "Extraneous Rows Block";
            
        const notes = `${typeLabel}: Continuous sequential rows (${startRow} to ${endRow}) grouped as 1 single consolidated error.`;
        
        result.push({
          sheet: sheetName,
          cell: cellRef,
          rowIndex: first.rowIndex,
          colIndex: first.colIndex,
          employeeValue: combinedEmpVal,
          reviewerValue: combinedRevVal,
          normalizedEmployeeValue: combinedEmpVal,
          normalizedReviewerValue: combinedRevVal,
          similarity: 0,
          errorType: first.errorType,
          severity: severity,
          penalty: penalty,
          notes: notes
        });
      }
    }
    
    // Add other non-row-level errors
    for (const err of otherErrors) {
      result.push(err);
    }
  }
  
  return result;
}

/**
 * Resolves a sheet name to its normalized base name (e.g. "Table 17" or "Table 17 (2)" -> "table 17").
 */
export function getNormalizedBaseName(name: string): string {
  let s = name.toLowerCase().trim();
  // Remove copy suffixes
  s = s.replace(/\s*[-_]?\s*copy\s*$/, "");
  // Remove parentheses or brackets and contents within them, e.g. (2), [2]
  s = s.replace(/\s*[([].*?[\])]\s*$/, "");
  // Remove part suffixes (e.g. " part 1", " part a", " pt 1", " pt. b", " part1")
  s = s.replace(/\s*[-_]?\s*(part|pt\.?)\s*[a-zA-Z0-9]+\s*$/, "");
  // Remove trailing alphabetic character if preceded by space, dash or underscore (e.g. "table 17 a" -> "table 17")
  // EXCEPT when it's part of a multi-digit number (to avoid turning "table 17" into "table 1")
  s = s.replace(/\s*[-_]?\s*[a-z]\s*$/, "");
  return s.trim();
}

/**
 * Checks if two sheet names are related (e.g. they share numbers like Table 17 vs Table 17 A).
 */
export function isNameRelated(name1: string, name2: string): boolean {
  const n1 = name1.toLowerCase().trim();
  const n2 = name2.toLowerCase().trim();
  if (n1 === n2) return true;
  if (n1.startsWith(n2) || n2.startsWith(n1)) return true;
  
  const d1 = n1.match(/\d+/g);
  const d2 = n2.match(/\d+/g);
  if (d1 && d2) {
    for (const num1 of d1) {
      if (d2.includes(num1)) return true;
    }
  }
  return false;
}

/**
 * Concatenates multiple SheetGrids vertically.
 */
export function concatenateSheets(sheets: SheetGrid[]): SheetGrid {
  const combined: SheetGrid = {
    name: "Combined",
    maxRow: -1,
    maxCol: -1,
    cells: {},
    virtualSegments: []
  };
  let currentRowOffset = 0;
  for (const sh of sheets) {
    const startRow = currentRowOffset;
    if (sh.maxCol > combined.maxCol) {
      combined.maxCol = sh.maxCol;
    }
    for (let r = 0; r <= sh.maxRow; r++) {
      for (let c = 0; c <= sh.maxCol; c++) {
        const cellVal = sh.cells[`${r},${c}`];
        if (cellVal) {
          combined.cells[`${currentRowOffset},${c}`] = cellVal;
        }
      }
      currentRowOffset++;
    }
    const endRow = currentRowOffset - 1;
    combined.virtualSegments!.push({
      originalSheetName: sh.name,
      virtualStartRow: startRow,
      virtualEndRow: endRow,
      originalStartRow: 0
    });
  }
  combined.maxRow = currentRowOffset - 1;
  return combined;
}

/**
 * Computes frequency-based cell content similarity between two SheetGrids.
 */
export function computeContentSimilarity(sheet1: SheetGrid, sheet2: SheetGrid): number {
  const vals1: string[] = [];
  for (const cell of Object.values(sheet1.cells)) {
    if (cell && cell.normalized !== undefined && cell.normalized !== null) {
      const norm = cell.normalized.trim().toLowerCase();
      if (norm !== "") {
        vals1.push(norm);
      }
    }
  }
  const vals2: string[] = [];
  for (const cell of Object.values(sheet2.cells)) {
    if (cell && cell.normalized !== undefined && cell.normalized !== null) {
      const norm = cell.normalized.trim().toLowerCase();
      if (norm !== "") {
        vals2.push(norm);
      }
    }
  }

  if (vals1.length === 0 && vals2.length === 0) return 100.0;
  if (vals1.length === 0 || vals2.length === 0) return 0.0;

  const freq1: Record<string, number> = {};
  for (const v of vals1) {
    freq1[v] = (freq1[v] || 0) + 1;
  }
  const freq2: Record<string, number> = {};
  for (const v of vals2) {
    freq2[v] = (freq2[v] || 0) + 1;
  }

  let intersection = 0;
  for (const [v, f1] of Object.entries(freq1)) {
    if (freq2[v]) {
      intersection += Math.min(f1, freq2[v]);
    }
  }

  const totalElements = Math.max(vals1.length, vals2.length);
  return parseFloat(((intersection / totalElements) * 100).toFixed(2));
}

/**
 * Slices a SheetGrid to only contain rows between startRow and endRow (inclusive).
 * Shifts row coordinates to start at 0.
 */
export function sliceSheet(sourceSheet: SheetGrid, startRow: number, endRow: number): SheetGrid {
  const sliced: SheetGrid = {
    name: `${sourceSheet.name} (Sliced)`,
    maxRow: endRow - startRow,
    maxCol: sourceSheet.maxCol,
    cells: {}
  };
  for (let r = startRow; r <= endRow; r++) {
    const targetRow = r - startRow;
    for (let c = 0; c <= sourceSheet.maxCol; c++) {
      const cellVal = sourceSheet.cells[`${r},${c}`];
      if (cellVal) {
        sliced.cells[`${targetRow},${c}`] = cellVal;
      }
    }
  }
  return sliced;
}

/**
 * Main Enterprise QA evaluation runner. Calculates final accurate metrics,
 * enforces safety overrides, categorizes errors strictly, and detects systematic patterns.
 */
export function executeQAEvaluation(
  employeeData: WorkbookData,
  reviewerData: WorkbookData,
  config: QAConfig
): AnalysisResult {
  const errorLog: ErrorLogEntry[] = [];
  const allShiftEvents: ShiftEvent[] = [];
  const virtualSheets: Record<string, SheetGrid> = {};
  
  let totalComparedCellsCount = 0;
  let totalErrorsCount = 0;
  let penaltyPointsTotal = 0;

  // Determine standard Strict Mode
  const strictModeActive = isStrictModeActive(employeeData, config);
  const effectiveTolerance = strictModeActive ? 0.0 : config.numericTolerance;

  // Identify non-excluded sheets
  const allSheetNames = Array.from(
    new Set([
      ...Object.keys(employeeData.sheets),
      ...Object.keys(reviewerData.sheets)
    ])
  ).filter(sheetName => {
    const empSheet = employeeData.sheets[sheetName];
    const revSheet = reviewerData.sheets[sheetName];
    const activeSheet = empSheet || revSheet;
    if (!activeSheet) return false;
    return !shouldExcludeSheet(sheetName, activeSheet.maxRow, Object.keys(activeSheet.cells).length);
  });

  // Track merged/split processed sheets so we can suppress baseline omissions errors
  const processedReviewerSheets = new Set<string>();
  const processedEmployeeSheets = new Set<string>();

  interface SheetComparisonPair {
    sheetName: string;
    empSheet?: SheetGrid;
    revSheet?: SheetGrid;
    isVirtualComparison: boolean;
    comparisonType: "split" | "merge" | "standard";
  }

  const comparisonPairs: SheetComparisonPair[] = [];

  // Extract missing sheets candidates
  const missingReviewerSheets = Object.keys(reviewerData.sheets).filter(name => {
    const sheet = reviewerData.sheets[name];
    if (shouldExcludeSheet(name, sheet.maxRow, Object.keys(sheet.cells).length)) return false;
    return !employeeData.sheets[name];
  });

  const extraEmployeeSheets = Object.keys(employeeData.sheets).filter(name => {
    const sheet = employeeData.sheets[name];
    if (shouldExcludeSheet(name, sheet.maxRow, Object.keys(sheet.cells).length)) return false;
    return !reviewerData.sheets[name];
  });

  // Group employee sheets by normalized base name:
  const empGroups: Record<string, string[]> = {};
  for (const name of Object.keys(employeeData.sheets)) {
    const sheet = employeeData.sheets[name];
    if (shouldExcludeSheet(name, sheet.maxRow, Object.keys(sheet.cells).length)) continue;
    const base = getNormalizedBaseName(name);
    if (!empGroups[base]) empGroups[base] = [];
    empGroups[base].push(name);
  }

  // Group reviewer sheets by normalized base name:
  const revGroups: Record<string, string[]> = {};
  for (const name of Object.keys(reviewerData.sheets)) {
    const sheet = reviewerData.sheets[name];
    if (shouldExcludeSheet(name, sheet.maxRow, Object.keys(sheet.cells).length)) continue;
    const base = getNormalizedBaseName(name);
    if (!revGroups[base]) revGroups[base] = [];
    revGroups[base].push(name);
  }

  // --- 1. PRIMARY BASE-NAME GROUPED TABLE SPLIT DETECTION ---
  for (const base of Object.keys(empGroups)) {
    const empSheets = empGroups[base];
    const revSheets = revGroups[base] || [];
    if (empSheets.length >= 2 && revSheets.length === 1) {
      const revName = revSheets[0];
      if (processedReviewerSheets.has(revName) || empSheets.some(name => processedEmployeeSheets.has(name))) {
        continue;
      }

      empSheets.sort((a, b) => a.localeCompare(b));
      const combined = concatenateSheets(empSheets.map(name => employeeData.sheets[name]));
      const sim = computeContentSimilarity(reviewerData.sheets[revName], combined);
      if (sim >= 95.0) {
        processedReviewerSheets.add(revName);
        empSheets.forEach(name => processedEmployeeSheets.add(name));

        totalErrorsCount++;
        penaltyPointsTotal += 10;
        
        // Count reviewer sheet cell coverage
        totalComparedCellsCount += Object.keys(reviewerData.sheets[revName].cells).length;

        const noteMsg = `Table Split Event: Reviewer sheet '${revName}' was split into employee sheets [${empSheets.join(", ")}]. Content similarity is ${sim}%.`;
        errorLog.push({
          sheet: revName,
          cell: "Sheet Layout",
          rowIndex: 0,
          colIndex: 0,
          employeeValue: empSheets.join("; "),
          reviewerValue: revName,
          normalizedEmployeeValue: "",
          normalizedReviewerValue: "",
          similarity: sim,
          errorType: ErrorType.TableSplit,
          severity: Severity.Critical,
          penalty: 10,
          notes: noteMsg
        });

        // Populates virtual sliced reviewer sheets so they align with partitioned employee ones
        if (combined.virtualSegments) {
          for (const segment of combined.virtualSegments) {
            const slicedRev = sliceSheet(reviewerData.sheets[revName], segment.virtualStartRow, segment.virtualEndRow);
            slicedRev.name = segment.originalSheetName;
            virtualSheets[segment.originalSheetName] = slicedRev;
          }
        }

        // Add the virtual comparison pair!
        comparisonPairs.push({
          sheetName: revName,
          empSheet: combined,
          revSheet: reviewerData.sheets[revName],
          isVirtualComparison: true,
          comparisonType: "split"
        });
      }
    }
  }

  // --- 2. PRIMARY BASE-NAME GROUPED TABLE MERGE DETECTION ---
  for (const base of Object.keys(revGroups)) {
    const revSheets = revGroups[base];
    const empSheets = empGroups[base] || [];
    if (revSheets.length >= 2 && empSheets.length === 1) {
      const empName = empSheets[0];
      if (processedEmployeeSheets.has(empName) || revSheets.some(name => processedReviewerSheets.has(name))) {
        continue;
      }

      revSheets.sort((a, b) => a.localeCompare(b));
      const combined = concatenateSheets(revSheets.map(name => reviewerData.sheets[name]));
      const sim = computeContentSimilarity(employeeData.sheets[empName], combined);
      if (sim >= 95.0) {
        processedEmployeeSheets.add(empName);
        revSheets.forEach(name => processedReviewerSheets.add(name));

        totalErrorsCount++;
        penaltyPointsTotal += 10;

        // Count employee sheet cell coverage
        totalComparedCellsCount += Object.keys(employeeData.sheets[empName].cells).length;

        const noteMsg = `Table Merge Event: Reviewer sheets [${revSheets.join(", ")}] were merged into employee sheet '${empName}'. Content similarity is ${sim}%.`;
        errorLog.push({
          sheet: empName,
          cell: "Sheet Layout",
          rowIndex: 0,
          colIndex: 0,
          employeeValue: empName,
          reviewerValue: revSheets.join("; "),
          normalizedEmployeeValue: "",
          normalizedReviewerValue: "",
          similarity: sim,
          errorType: ErrorType.TableMerge,
          severity: Severity.Critical,
          penalty: 10,
          notes: noteMsg
        });

        // Store the virtual merged reviewer sheet for visual alignment
        virtualSheets[empName] = combined;

        // Add the virtual comparison pair!
        comparisonPairs.push({
          sheetName: empName,
          empSheet: employeeData.sheets[empName],
          revSheet: combined,
          isVirtualComparison: true,
          comparisonType: "merge"
        });
      }
    }
  }

  // --- 3. FALLBACK TABLE SPLIT EVENTS ---
  for (const revName of missingReviewerSheets) {
    if (processedReviewerSheets.has(revName)) continue;
    const candidates = extraEmployeeSheets.filter(
      empName => !processedEmployeeSheets.has(empName) && isNameRelated(revName, empName)
    );
    if (candidates.length >= 2) {
      candidates.sort((a, b) => a.localeCompare(b));
      const combined = concatenateSheets(candidates.map(name => employeeData.sheets[name]));
      const sim = computeContentSimilarity(reviewerData.sheets[revName], combined);
      if (sim >= 95.0) {
        processedReviewerSheets.add(revName);
        candidates.forEach(name => processedEmployeeSheets.add(name));

        totalErrorsCount++;
        penaltyPointsTotal += 10;
        
        // Count reviewer sheet cell coverage
        totalComparedCellsCount += Object.keys(reviewerData.sheets[revName].cells).length;

        const noteMsg = `Table Split Event: Reviewer sheet '${revName}' was split into employee sheets [${candidates.join(", ")}]. Content similarity is ${sim}%.`;
        errorLog.push({
          sheet: revName,
          cell: "Sheet Layout",
          rowIndex: 0,
          colIndex: 0,
          employeeValue: candidates.join("; "),
          reviewerValue: revName,
          normalizedEmployeeValue: "",
          normalizedReviewerValue: "",
          similarity: sim,
          errorType: ErrorType.TableSplit,
          severity: Severity.Critical,
          penalty: 10,
          notes: noteMsg
        });

        // Populates virtual sliced reviewer sheets so they align with partitioned employee ones
        if (combined.virtualSegments) {
          for (const segment of combined.virtualSegments) {
            const slicedRev = sliceSheet(reviewerData.sheets[revName], segment.virtualStartRow, segment.virtualEndRow);
            slicedRev.name = segment.originalSheetName;
            virtualSheets[segment.originalSheetName] = slicedRev;
          }
        }

        // Add the virtual comparison pair!
        comparisonPairs.push({
          sheetName: revName,
          empSheet: combined,
          revSheet: reviewerData.sheets[revName],
          isVirtualComparison: true,
          comparisonType: "split"
        });
      }
    }
  }

  // --- 4. FALLBACK TABLE MERGE EVENTS ---
  for (const empName of extraEmployeeSheets) {
    if (processedEmployeeSheets.has(empName)) continue;
    
    const candidates = missingReviewerSheets.filter(
      revName => !processedReviewerSheets.has(revName) && isNameRelated(empName, revName)
    );
    if (candidates.length >= 2) {
      candidates.sort((a, b) => a.localeCompare(b));
      const combined = concatenateSheets(candidates.map(name => reviewerData.sheets[name]));
      const sim = computeContentSimilarity(employeeData.sheets[empName], combined);
      if (sim >= 95.0) {
        processedEmployeeSheets.add(empName);
        candidates.forEach(name => processedReviewerSheets.add(name));

        totalErrorsCount++;
        penaltyPointsTotal += 10;

        // Count employee sheet cell coverage
        totalComparedCellsCount += Object.keys(employeeData.sheets[empName].cells).length;

        const noteMsg = `Table Merge Event: Reviewer sheets [${candidates.join(", ")}] were merged into employee sheet '${empName}'. Content similarity is ${sim}%.`;
        errorLog.push({
          sheet: empName,
          cell: "Sheet Layout",
          rowIndex: 0,
          colIndex: 0,
          employeeValue: empName,
          reviewerValue: candidates.join("; "),
          normalizedEmployeeValue: "",
          normalizedReviewerValue: "",
          similarity: sim,
          errorType: ErrorType.TableMerge,
          severity: Severity.Critical,
          penalty: 10,
          notes: noteMsg
        });

        // Store the virtual merged reviewer sheet for visual alignment
        virtualSheets[empName] = combined;

        // Add the virtual comparison pair!
        comparisonPairs.push({
          sheetName: empName,
          empSheet: employeeData.sheets[empName],
          revSheet: combined,
          isVirtualComparison: true,
          comparisonType: "merge"
        });
      }
    }
  }

  // Add standard sheets comparison pairs
  for (const sheetName of allSheetNames) {
    if (!processedReviewerSheets.has(sheetName) && !processedEmployeeSheets.has(sheetName)) {
      comparisonPairs.push({
        sheetName,
        empSheet: employeeData.sheets[sheetName],
        revSheet: reviewerData.sheets[sheetName],
        isVirtualComparison: false,
        comparisonType: "standard"
      });
    }
  }

  // Normalization counters for categories partitions
  let missingValuesCount = 0;
  let extraValuesCount = 0;
  let shiftErrorsCount = 0;
  let numericErrorsCount = 0;
  let textErrorsCount = 0;
  let rangeErrorsCount = 0;
  let headerErrorsCount = 0;

  // Heuristics maps
  const numericSubstitutions: Record<string, number> = {};
  const copyPasteValues: Record<string, number> = {};
  const sheetErrorDensities: Record<string, number> = {};
  const clustersBySheet: Record<string, number> = {};

  for (const pair of comparisonPairs) {
    const sheetName = pair.sheetName;
    const empSheet = pair.empSheet;
    const revSheet = pair.revSheet;
    const isVirtualComparison = pair.isVirtualComparison;

    // Missing sheet or extra sheet triggers full coordinate missing/extra values
    if (!empSheet || !revSheet) {
      const activeSheet = empSheet || revSheet;
      const isEmployeeOnly = !!empSheet;
      
      for (const [coord, cell] of Object.entries(activeSheet.cells)) {
        const [rStr, cStr] = coord.split(",");
        const rIndex = parseInt(rStr);
        const cIndex = parseInt(cStr);
        const cellRef = `${getColLetter(cIndex)}${rIndex + 1}`;
        
        totalComparedCellsCount++;
        totalErrorsCount++;

        const errType = isEmployeeOnly ? ErrorType.ExtraValue : ErrorType.MissingValue;
        const sev = Severity.High;
        const penalty = 5;

        if (errType === ErrorType.MissingValue) missingValuesCount++;
        else extraValuesCount++;

        penaltyPointsTotal += penalty;

        errorLog.push({
          sheet: sheetName,
          cell: cellRef,
          rowIndex: rIndex,
          colIndex: cIndex,
          employeeValue: isEmployeeOnly ? cell.formatted : "",
          reviewerValue: isEmployeeOnly ? "" : cell.formatted,
          normalizedEmployeeValue: isEmployeeOnly ? cell.normalized : "",
          normalizedReviewerValue: isEmployeeOnly ? "" : cell.normalized,
          similarity: 0,
          errorType: errType,
          severity: sev,
          penalty,
          notes: isEmployeeOnly 
            ? `Extraneous cell present in worker submission but ground truth sheet '${sheetName}' is missing.`
            : `Cell expected in ground truth sheet but entirely omitted by worker in sheet '${sheetName}'.`
         });
      }
      continue;
    }

    // --- RUN SHIFT AND ALIGNMENT RECOVERY ---
    const maxCol = Math.max(empSheet.maxCol, revSheet.maxCol);
    const maxRow = Math.max(empSheet.maxRow, revSheet.maxRow);

    // 1. Perform optimal DP Row alignment to determine matched rows, extra rows, or skipped rows
    const steps = alignSheetRows(empSheet, revSheet, maxCol);

    // 2. Perform optimal DP Column alignment to determine matched columns, extra columns, or skipped columns
    const colSteps = alignSheetColumns(empSheet, revSheet, maxRow);

    // 3. Header row detection for reviewer sheet
    const headerRows = detectHeaderRows(revSheet);

    // 4. Group adjacent row operations to detect Missing Rows vs Extra Rows vs Row Shifts
    interface StepBlock {
      type: "match" | "missing_row" | "extra_row";
      startIndex: number;
      endIndex: number;
      count: number;
    }

    const stepBlocks: StepBlock[] = [];
    if (steps.length > 0) {
      let currentType = steps[0].type;
      let startIndex = 0;

      for (let idx = 1; idx <= steps.length; idx++) {
        const isEnd = idx === steps.length;
        const itemType = isEnd ? null : steps[idx].type;

        if (isEnd || itemType !== currentType) {
          stepBlocks.push({
            type: currentType,
            startIndex,
            endIndex: idx - 1,
            count: idx - startIndex
          });
          if (!isEnd) {
            currentType = itemType!;
            startIndex = idx;
          }
        }
      }
    }

    // 5. Group adjacent column operations to detect Missing Columns vs Extra Columns vs Column Shifts
    interface ColStepBlock {
      type: "match" | "missing_column" | "extra_column";
      startIndex: number;
      endIndex: number;
      count: number;
    }

    const colStepBlocks: ColStepBlock[] = [];
    if (colSteps.length > 0) {
      let currentType = colSteps[0].type;
      let startIndex = 0;

      for (let idx = 1; idx <= colSteps.length; idx++) {
        const isEnd = idx === colSteps.length;
        const itemType = isEnd ? null : colSteps[idx].type;

        if (isEnd || itemType !== currentType) {
          colStepBlocks.push({
            type: currentType,
            startIndex,
            endIndex: idx - 1,
            count: idx - startIndex
          });
          if (!isEnd) {
            currentType = itemType!;
            startIndex = idx;
          }
        }
      }
    }

    // 6. Record columns errors once per sheet (Missing/Extra Column or Column Shift)
    for (const colBlock of colStepBlocks) {
      if (isVirtualComparison) continue;
      if (colBlock.type === "missing_column") {
        const isSmallGap = colBlock.count <= 5;
        if (isSmallGap) {
          for (let idx = colBlock.startIndex; idx <= colBlock.endIndex; idx++) {
            const step = colSteps[idx];
            const revCol = step.revCol!;
            totalComparedCellsCount++;
            totalErrorsCount++;
            missingValuesCount++;
            const penalty = 5;
            penaltyPointsTotal += penalty;

            const summaryParts: string[] = [];
            for (let r = 0; r <= Math.min(4, empSheet.maxRow); r++) {
              const cell = revSheet.cells[`${r},${revCol}`];
              if (cell && cell.formatted) {
                summaryParts.push(`R${r + 1}: ${cell.formatted}`);
              }
            }
            if (empSheet.maxRow > 4) summaryParts.push("...");
            const colText = summaryParts.join(", ");

            errorLog.push({
              sheet: sheetName,
              cell: `Col ${getColLetter(revCol)}`,
              rowIndex: 0,
              colIndex: revCol,
              employeeValue: "",
              reviewerValue: colText || "(Empty Column)",
              normalizedEmployeeValue: "",
              normalizedReviewerValue: colText,
              similarity: 0,
              errorType: ErrorType.MissingColumn,
              severity: Severity.High,
              penalty,
              notes: `Missing Column: This entire data entry column was accidentally omitted by the employee. Aligned comparison recovered immediately on subsequent columns.`
            });
          }
        } else {
          // Large Block → Column Shift
          const startStep = colSteps[colBlock.startIndex];
          const endStep = colSteps[colBlock.endIndex];
          const startCol = startStep.revCol!;
          const endCol = endStep.revCol!;
          const affectedCells = colBlock.count * (empSheet.maxRow + 1);

          allShiftEvents.push({
            sheetName,
            type: "column",
            offset: -1,
            spanStart: startCol,
            spanEnd: endCol,
            detail: `Column Block Shift: Columns ${getColLetter(startCol)} to ${getColLetter(endCol)} suffer large scale mismatch block of ${colBlock.count} consecutive omitted columns [Column Shift].`,
            affectedCellsCount: affectedCells
          });

          for (let idx = colBlock.startIndex; idx <= colBlock.endIndex; idx++) {
            const step = colSteps[idx];
            const revCol = step.revCol!;
            totalComparedCellsCount++;
            totalErrorsCount++;
            shiftErrorsCount++;
            const penalty = 10;
            penaltyPointsTotal += penalty;

            errorLog.push({
              sheet: sheetName,
              cell: `Col ${getColLetter(revCol)}`,
              rowIndex: 0,
              colIndex: revCol,
              employeeValue: "",
              reviewerValue: `Block Shift (Col ${getColLetter(revCol)})`,
              normalizedEmployeeValue: "",
              normalizedReviewerValue: `Block Shift`,
              similarity: 0,
              errorType: ErrorType.ColumnShift,
              severity: Severity.Critical,
              penalty,
              notes: `Confirmed Column Shift: Consecutively omitted substantial block (${colBlock.count} columns) with no local recovery alignment.`
            });
          }
        }
      } else if (colBlock.type === "extra_column") {
        const isSmallGap = colBlock.count <= 5;
        if (isSmallGap) {
          for (let idx = colBlock.startIndex; idx <= colBlock.endIndex; idx++) {
            const step = colSteps[idx];
            const empCol = step.empCol!;
            totalComparedCellsCount++;
            totalErrorsCount++;
            extraValuesCount++;
            const penalty = 5;
            penaltyPointsTotal += penalty;

            const summaryParts: string[] = [];
            for (let r = 0; r <= Math.min(4, empSheet.maxRow); r++) {
              const cell = empSheet.cells[`${r},${empCol}`];
              if (cell && cell.formatted) {
                summaryParts.push(`R${r + 1}: ${cell.formatted}`);
              }
            }
            if (empSheet.maxRow > 4) summaryParts.push("...");
            const colText = summaryParts.join(", ");

            errorLog.push({
              sheet: sheetName,
              cell: `Col ${getColLetter(empCol)}`,
              rowIndex: 0,
              colIndex: empCol,
              employeeValue: colText || "(Empty Column)",
              reviewerValue: "",
              normalizedEmployeeValue: colText,
              normalizedReviewerValue: "",
              similarity: 0,
              errorType: ErrorType.ExtraColumn,
              severity: Severity.High,
              penalty,
              notes: `Extra Column: An extraneous data entry column was inserted in submission. Aligned comparison recovered immediately on subsequent columns.`
            });
          }
        } else {
          // Large Block → Column Shift
          const startStep = colSteps[colBlock.startIndex];
          const endStep = colSteps[colBlock.endIndex];
          const startCol = startStep.empCol!;
          const endCol = endStep.empCol!;
          const affectedCells = colBlock.count * (empSheet.maxRow + 1);

          allShiftEvents.push({
            sheetName,
            type: "column",
            offset: 1,
            spanStart: startCol,
            spanEnd: endCol,
            detail: `Column Block Shift: Extraneous columns ${getColLetter(startCol)} to ${getColLetter(endCol)} suffer large scale mismatch block of ${colBlock.count} consecutive inserted columns [Column Shift].`,
            affectedCellsCount: affectedCells
          });

          for (let idx = colBlock.startIndex; idx <= colBlock.endIndex; idx++) {
            const step = colSteps[idx];
            const empCol = step.empCol!;
            totalComparedCellsCount++;
            totalErrorsCount++;
            shiftErrorsCount++;
            const penalty = 10;
            penaltyPointsTotal += penalty;

            errorLog.push({
              sheet: sheetName,
              cell: `Col ${getColLetter(empCol)}`,
              rowIndex: 0,
              colIndex: empCol,
              employeeValue: `Block Shift (Col ${getColLetter(empCol)})`,
              reviewerValue: "",
              normalizedEmployeeValue: "Block Shift",
              normalizedReviewerValue: "",
              similarity: 0,
              errorType: ErrorType.ColumnShift,
              severity: Severity.Critical,
              penalty,
              notes: `Confirmed Column Shift: Consecutively inserted extraneous substantial block (${colBlock.count} columns) with no local recovery alignment.`
            });
          }
        }
      }
    }

    // Filter to columns that are matched
    const matchedColSteps = colSteps.filter(s => s.type === "match" && s.empCol !== null && s.revCol !== null);

    // 7. Evaluate blocks with strict alignment recovery context
    for (const block of stepBlocks) {
      if (block.type === "match") {
        for (let idx = block.startIndex; idx <= block.endIndex; idx++) {
          const step = steps[idx];
          const empRow = step.empRow!;
          const revRow = step.revRow!;

          let colIdxEmp = 0;
          let colIdxRev = 0;

          while (colIdxEmp < matchedColSteps.length || colIdxRev < matchedColSteps.length) {
            if (colIdxEmp >= matchedColSteps.length) {
              const stepR = matchedColSteps[colIdxRev];
              const revCol = stepR.revCol!;
              const revCell = revSheet.cells[`${revRow},${revCol}`];
              if (revCell && revCell.type !== "empty") {
                totalComparedCellsCount++;
                totalErrorsCount++;
                missingValuesCount++;
                penaltyPointsTotal += 5;
                errorLog.push({
                  sheet: sheetName,
                  cell: `${getColLetter(revCol)}${revRow + 1}`,
                  rowIndex: revRow,
                  colIndex: revCol,
                  employeeValue: "",
                  reviewerValue: revCell.formatted,
                  normalizedEmployeeValue: "",
                  normalizedReviewerValue: revCell.normalized,
                  similarity: 0,
                  errorType: ErrorType.MissingValue,
                  severity: Severity.High,
                  penalty: 5,
                  notes: `Omitted Cell at end of row: expected '${revCell.formatted}'.`
                });
              }
              colIdxRev++;
              continue;
            }

            if (colIdxRev >= matchedColSteps.length) {
              const stepE = matchedColSteps[colIdxEmp];
              const empCol = stepE.empCol!;
              const empCell = empSheet.cells[`${empRow},${empCol}`];
              if (empCell && empCell.type !== "empty") {
                totalComparedCellsCount++;
                totalErrorsCount++;
                extraValuesCount++;
                penaltyPointsTotal += 5;
                errorLog.push({
                  sheet: sheetName,
                  cell: `${getColLetter(empCol)}${revRow + 1}`,
                  rowIndex: revRow,
                  colIndex: empCol,
                  employeeValue: empCell.formatted,
                  reviewerValue: "",
                  normalizedEmployeeValue: empCell.normalized,
                  normalizedReviewerValue: "",
                  similarity: 0,
                  errorType: ErrorType.ExtraValue,
                  severity: Severity.High,
                  penalty: 5,
                  notes: `Extraneous Cell at end of row: of value '${empCell.formatted}'.`
                });
              }
              colIdxEmp++;
              continue;
            }

            const stepE = matchedColSteps[colIdxEmp];
            const stepR = matchedColSteps[colIdxRev];

            const empCol = stepE.empCol!;
            const revCol = stepR.revCol!;

            const empCell = empSheet.cells[`${empRow},${empCol}`];
            const revCell = revSheet.cells[`${revRow},${revCol}`];

            if (!empCell && !revCell) {
              colIdxEmp++;
              colIdxRev++;
              continue;
            }

            const eNormalized = empCell ? empCell.normalized : "";
            const rNormalized = revCell ? revCell.normalized : "";
            const eType = empCell ? empCell.type : "empty";
            const rType = revCell ? revCell.type : "empty";

            if (eNormalized === rNormalized && eType === rType) {
              totalComparedCellsCount++;
              colIdxEmp++;
              colIdxRev++;
              continue;
            }

            // Local Sequence Alignment Attempt
            let foundLocal = false;
            const offsets = [1, -1, 2, -2];

            for (const dc of offsets) {
              let localCompared = 0;
              let localMatches = 0;
              const K = 5;

              for (let k = 0; k < K; k++) {
                const idxE = colIdxEmp + (dc < 0 ? -dc : 0) + k;
                const idxR = colIdxRev + (dc > 0 ? dc : 0) + k;

                if (idxE >= matchedColSteps.length || idxR >= matchedColSteps.length) {
                  break;
                }

                const tempE = matchedColSteps[idxE].empCol!;
                const tempR = matchedColSteps[idxR].revCol!;

                const cellE = empSheet.cells[`${empRow},${tempE}`];
                const cellR = revSheet.cells[`${revRow},${tempR}`];

                if (!cellE && !cellR) continue;

                localCompared++;
                if (cellE && cellR && cellE.normalized === cellR.normalized && cellE.type === cellR.type) {
                  localMatches++;
                }
              }

              if (localCompared > 0 && (localMatches / localCompared) >= 0.80) {
                foundLocal = true;
                const cellRef = `${getColLetter(revCol)}${revRow + 1}`;
                const eFormatted = empCell ? empCell.formatted : "";
                const rFormatted = revCell ? revCell.formatted : "";

                totalComparedCellsCount++;

                if (isVirtualComparison) {
                  // Just advance indices, do NOT log any error or penalty!
                  if (dc === 1) colIdxRev += 1;
                  else if (dc === -1) colIdxEmp += 1;
                  else if (dc === 2) colIdxRev += 2;
                  else colIdxEmp += 2;
                  break;
                }

                totalErrorsCount++;

                let errType: ErrorType;
                let notes: string;
                let penalty = 5;
                let severity = Severity.High;

                if (dc === 1) {
                  errType = ErrorType.MissingCell;
                  notes = `Local Alignment Recovery: Missing Cell detected. Subsequent cells aligned after skipping reviewer column '${getColLetter(revCol)}'.`;
                  missingValuesCount++;
                  colIdxRev += 1;
                } else if (dc === -1) {
                  errType = ErrorType.ExtraCell;
                  notes = `Local Alignment Recovery: Extraneous Cell detected. Subsequent cells aligned after skipping employee column '${getColLetter(empCol)}'.`;
                  extraValuesCount++;
                  colIdxEmp += 1;
                } else if (dc === 2) {
                  errType = ErrorType.LocalColumnMisalignment;
                  notes = `Local Alignment Recovery: Local Column Misalignment (2 cells shift) detected. Subsequent cells aligned after skipping 2 reviewer values.`;
                  shiftErrorsCount++;
                  colIdxRev += 2;
                } else {
                  errType = ErrorType.LocalColumnMisalignment;
                  notes = `Local Alignment Recovery: Local Column Misalignment (2 cells shift) detected. Subsequent cells aligned after skipping 2 employee values.`;
                  shiftErrorsCount++;
                  colIdxEmp += 2;
                }

                penaltyPointsTotal += penalty;

                errorLog.push({
                  sheet: sheetName,
                  cell: cellRef,
                  rowIndex: revRow,
                  colIndex: revCol,
                  employeeValue: eFormatted,
                  reviewerValue: rFormatted,
                  normalizedEmployeeValue: eNormalized,
                  normalizedReviewerValue: rNormalized,
                  similarity: 0,
                  errorType: errType,
                  severity,
                  penalty,
                  notes
                });

                break;
              }
            }

            if (foundLocal) {
              continue;
            }

            // Normal cell mismatch classifications
            totalComparedCellsCount++;
            totalErrorsCount++;
            let errType: ErrorType = ErrorType.TextDifference;
            let severity = Severity.Medium;
            let penalty = 2;
            let notes = "";
            let sim = 0;

            const cellRef = `${getColLetter(revCol)}${revRow + 1}`;
            const eFormatted = empCell ? empCell.formatted : "";
            const rFormatted = revCell ? revCell.formatted : "";

            if (eType === "empty" && rType !== "empty") {
              errType = ErrorType.MissingValue;
              severity = Severity.High;
              penalty = 5;
              missingValuesCount++;
              notes = `Omitted Cell: expected '${rFormatted}' but worker cell is empty.`;
            } else if (eType !== "empty" && rType === "empty") {
              errType = ErrorType.ExtraValue;
              severity = Severity.High;
              penalty = 5;
              extraValuesCount++;
              notes = `Extraneous Cell: cell is filled with '${eFormatted}' but expected empty.`;
            } else if (
              parseRangeOrSequence(rFormatted).isValid ||
              parseRangeOrSequence(eFormatted).isValid
            ) {
              const rRange = parseRangeOrSequence(rFormatted);
              const eRange = parseRangeOrSequence(eFormatted);

              if (rRange.isValid && eRange.isValid) {
                if (rRange.left === eRange.right && rRange.right === eRange.left) {
                  errType = ErrorType.RangeInversionError;
                  severity = Severity.Medium;
                  penalty = 2;
                  notes = "Sequence order reversed — directionally incorrect.";
                  rangeErrorsCount++;
                } else if (
                  (rRange.left === eRange.left && rRange.right !== eRange.right) ||
                  (rRange.left !== eRange.left && rRange.right === eRange.right)
                ) {
                  errType = ErrorType.RangeBoundaryError;
                  severity = Severity.Medium;
                  penalty = 2;
                  notes = `Sequence boundary mismatch: expected '${rFormatted}' but submitted '${eFormatted}'.`;
                  rangeErrorsCount++;
                } else if (isEquivalentRangeRepresentational(rRange, eRange)) {
                  errType = ErrorType.RangeRepresentationError;
                  severity = Severity.High;
                  penalty = 5;
                  notes = "Employee failed to reproduce the source representation exactly, which violates audit-grade data-entry requirements.";
                  rangeErrorsCount++;
                } else {
                  errType = ErrorType.RangeRepresentationError;
                  severity = Severity.High;
                  penalty = 5;
                  notes = `Range Representation mismatch: expected '${rFormatted}' but worker entered '${eFormatted}'.`;
                  rangeErrorsCount++;
                }
              } else {
                errType = ErrorType.RangeRepresentationError;
                severity = Severity.High;
                penalty = 5;
                notes = `Invalid range structure: worker entered '${eFormatted}' where range pattern '${rFormatted}' was expected.`;
                rangeErrorsCount++;
              }
            } else if (eType === "number" || rType === "number") {
              const eNum = Number(eNormalized);
              const rNum = Number(rNormalized);

              let matchesTolerance = false;
              if (effectiveTolerance > 0) {
                if (config.numericToleranceMode === "PERCENTAGE" && rNum !== 0) {
                  matchesTolerance = (Math.abs(eNum - rNum) / Math.abs(rNum)) < effectiveTolerance;
                } else {
                  matchesTolerance = Math.abs(eNum - rNum) < effectiveTolerance;
                }
              }

              if (matchesTolerance) {
                totalErrorsCount--;
                colIdxEmp++;
                colIdxRev++;
                continue;
              }

              const numErr = classifyDigitError(eNormalized, rNormalized);
              if (numErr) {
                errType = numErr;
                if (numErr === ErrorType.DigitSubstitution) {
                  const pair = `${rNormalized}→${eNormalized}`;
                  numericSubstitutions[pair] = (numericSubstitutions[pair] || 0) + 1;
                }
              } else {
                errType = ErrorType.NumericDifference;
              }

              severity = Severity.Medium;
              penalty = 2;

              let variance = 0;
              if (rNum !== 0) {
                variance = Math.abs((eNum - rNum) / rNum);
              } else if (eNum !== 0) {
                variance = 1.0;
              }

              const absoluteDiff = Math.abs(eNum - rNum);
              const isMajorVariance = variance > config.numericMajorVarianceThreshold;
              const isMajorAbsolute = absoluteDiff > config.numericMajorAbsoluteThreshold;

              if (isMajorVariance || isMajorAbsolute) {
                errType = ErrorType.MajorNumericError;
                severity = Severity.High;
                penalty = 5;
                notes = `Major Numeric Error: Variance is ${(variance * 100).toFixed(1)}% (Threshold: ${config.numericMajorVarianceThreshold * 100}%) and Absolute Diff is ${absoluteDiff.toFixed(2)}`;
              } else {
                notes = `Arithmetic discrepancy: Absolute diff is ${absoluteDiff.toFixed(2)}, variance is ${(variance * 100).toFixed(1)}%.`;
              }

              numericErrorsCount++;
            } else {
              const empStr = eFormatted;
              const revStr = rFormatted;

              const hasDigits = /\d/.test(empStr) || /\d/.test(revStr);
              if (hasDigits) {
                const mixedResult = analyzeMixedAlphaNumeric(empStr, revStr);
                sim = mixedResult.similarity;

                if (mixedResult.digitError) {
                  errType = mixedResult.digitError;
                  severity = Severity.Medium;
                  penalty = 2;
                  notes = `Mixed word digit error: ${mixedResult.digitError}. Context mismatch.`;
                  numericErrorsCount++;
                } else {
                  if (sim >= 90) {
                    errType = ErrorType.TextTypo;
                    severity = Severity.Medium;
                    penalty = 2;
                    notes = `Minor typographical mismatch (spelling overlap ${sim.toFixed(1)}%).`;
                  } else {
                    errType = ErrorType.MajorTextDifference;
                    severity = Severity.High;
                    penalty = 5;
                    notes = `Major alphanumeric structure mismatch (spelling overlap ${sim.toFixed(1)}%).`;
                  }
                  textErrorsCount++;
                }
              } else {
                sim = calculateTextSimilarity(eNormalized, rNormalized);
                if (sim >= 90) {
                  errType = ErrorType.TextTypo;
                  severity = Severity.Medium;
                  penalty = 2;
                  notes = `Minor spelling variation (similarity ${sim.toFixed(1)}%).`;
                } else {
                  errType = ErrorType.MajorTextDifference;
                  severity = Severity.High;
                  penalty = 5;
                  notes = `Substantive spelling divergence: expected '${rFormatted}' but worker entered '${eFormatted}'.`;
                }
                textErrorsCount++;
              }

              const pairKey = `${revStr}→${empStr}`;
              copyPasteValues[pairKey] = (copyPasteValues[pairKey] || 0) + 1;
            }

            // Header Row override checking
            const isHeaderErr = headerRows.has(revRow);
            if (isHeaderErr) {
              severity = Severity.High;
              penalty = config.headerPenalty;
              notes = `[Header Row Error] May disrupt whole column interpretation. Notes: ${notes}`;
              headerErrorsCount++;
            }

            penaltyPointsTotal += penalty;

            let targetSheetName = sheetName;
            let targetRowIndex = revRow;
            let targetCellRef = cellRef;

            if (isVirtualComparison) {
              if (pair.comparisonType === "split" && empSheet?.virtualSegments) {
                const segment = empSheet.virtualSegments.find(
                  (seg) => empRow >= seg.virtualStartRow && empRow <= seg.virtualEndRow
                );
                if (segment) {
                  targetSheetName = segment.originalSheetName;
                  const localEmpRow = empRow - segment.virtualStartRow + segment.originalStartRow;
                  targetRowIndex = localEmpRow;
                  targetCellRef = `${getColLetter(revCol)}${localEmpRow + 1}`;
                }
              } else if (pair.comparisonType === "merge") {
                targetRowIndex = empRow;
                targetCellRef = `${getColLetter(empCol)}${empRow + 1}`;
              }
            }

            errorLog.push({
              sheet: targetSheetName,
              cell: targetCellRef,
              rowIndex: targetRowIndex,
              colIndex: revCol,
              employeeValue: eFormatted,
              reviewerValue: rFormatted,
              normalizedEmployeeValue: eNormalized,
              normalizedReviewerValue: rNormalized,
              similarity: sim,
              errorType: errType,
              severity,
              penalty,
              notes
            });

            const neighborIdx = Math.floor(targetRowIndex / 5);
            const clusterKey = `${targetSheetName}-RowNeighborhood-${neighborIdx}`;
            clustersBySheet[clusterKey] = (clustersBySheet[clusterKey] || 0) + 1;

            colIdxEmp++;
            colIdxRev++;
          }
        }
      } else if (block.type === "missing_row") {
        if (isVirtualComparison) continue;
        const isSmallGap = block.count <= 5;

        if (isSmallGap) {
          for (let idx = block.startIndex; idx <= block.endIndex; idx++) {
            const step = steps[idx];
            const revRow = step.revRow!;
            const cellRef = `Row ${revRow + 1}`;

            totalComparedCellsCount++;
            totalErrorsCount++;
            missingValuesCount++;
            const penalty = 5;
            penaltyPointsTotal += penalty;

            const summaryParts: string[] = [];
            for (let c = 0; c <= Math.min(4, maxCol); c++) {
              const cell = revSheet.cells[`${revRow},${c}`];
              if (cell && cell.formatted) {
                summaryParts.push(`${getColLetter(c)}: ${cell.formatted}`);
              }
            }
            if (maxCol > 4) summaryParts.push("...");
            const rowText = summaryParts.join(", ");

            errorLog.push({
              sheet: sheetName,
              cell: cellRef,
              rowIndex: revRow,
              colIndex: 0,
              employeeValue: "",
              reviewerValue: rowText || "(Empty Row)",
              normalizedEmployeeValue: "",
              normalizedReviewerValue: rowText,
              similarity: 0,
              errorType: ErrorType.MissingRow,
              severity: Severity.High,
              penalty,
              notes: `Missing Row: This entire data entry row was accidentally omitted by the employee. Aligned comparison recovered immediately on subsequent rows.`
            });
          }
        } else {
          // Large Block → Row Shift
          const startStep = steps[block.startIndex];
          const endStep = steps[block.endIndex];
          const startRow = startStep.revRow!;
          const endRow = endStep.revRow!;
          const affectedCells = block.count * (maxCol + 1);

          allShiftEvents.push({
            sheetName,
            type: "row",
            offset: -1,
            spanStart: startRow,
            spanEnd: endRow,
            detail: `Row Block Shift: Rows ${startRow + 1} to ${endRow + 1} suffer large scale mismatch block of ${block.count} consecutive omitted rows [Row Shift].`,
            affectedCellsCount: affectedCells
          });

          for (let idx = block.startIndex; idx <= block.endIndex; idx++) {
            const step = steps[idx];
            const revRow = step.revRow!;
            const cellRef = `Row ${revRow + 1}`;

            totalComparedCellsCount++;
            totalErrorsCount++;
            shiftErrorsCount++;
            const penalty = 10;
            penaltyPointsTotal += penalty;

            errorLog.push({
              sheet: sheetName,
              cell: cellRef,
              rowIndex: revRow,
              colIndex: 0,
              employeeValue: "",
              reviewerValue: `Block Shift (Row ${revRow + 1})`,
              normalizedEmployeeValue: "",
              normalizedReviewerValue: `Block Shift`,
              similarity: 0,
              errorType: ErrorType.RowShift,
              severity: Severity.Critical,
              penalty,
              notes: `Confirmed Row Shift: Consecutively omitted substantial block (${block.count} rows) with no local recovery alignment.`
            });
          }
        }
      } else if (block.type === "extra_row") {
        if (isVirtualComparison) continue;
        const isSmallGap = block.count <= 5;

        if (isSmallGap) {
          for (let idx = block.startIndex; idx <= block.endIndex; idx++) {
            const step = steps[idx];
            const empRow = step.empRow!;
            const cellRef = `Row ${empRow + 1}`;

            totalComparedCellsCount++;
            totalErrorsCount++;
            extraValuesCount++;
            const penalty = 5;
            penaltyPointsTotal += penalty;

            const summaryParts: string[] = [];
            for (let c = 0; c <= Math.min(4, maxCol); c++) {
              const cell = empSheet.cells[`${empRow},${c}`];
              if (cell && cell.formatted) {
                summaryParts.push(`${getColLetter(c)}: ${cell.formatted}`);
              }
            }
            if (maxCol > 4) summaryParts.push("...");
            const rowText = summaryParts.join(", ");

            errorLog.push({
              sheet: sheetName,
              cell: cellRef,
              rowIndex: empRow,
              colIndex: 0,
              employeeValue: rowText || "(Empty Row)",
              reviewerValue: "",
              normalizedEmployeeValue: rowText,
              normalizedReviewerValue: "",
              similarity: 0,
              errorType: ErrorType.ExtraRow,
              severity: Severity.High,
              penalty,
              notes: `Extra Row: An extraneous data entry row was inserted in submission. Aligned comparison recovered immediately on subsequent rows.`
            });
          }
        } else {
          // Large Block → Row Shift
          const startStep = steps[block.startIndex];
          const endStep = steps[block.endIndex];
          const startRow = startStep.empRow!;
          const endRow = endStep.empRow!;
          const affectedCells = block.count * (maxCol + 1);

          allShiftEvents.push({
            sheetName,
            type: "row",
            offset: 1,
            spanStart: startRow,
            spanEnd: endRow,
            detail: `Row Block Shift: Extraneous rows ${startRow + 1} to ${endRow + 1} suffer large scale mismatch block of ${block.count} consecutive inserted rows [Row Shift].`,
            affectedCellsCount: affectedCells
          });

          for (let idx = block.startIndex; idx <= block.endIndex; idx++) {
            const step = steps[idx];
            const empRow = step.empRow!;
            const cellRef = `Row ${empRow + 1}`;

            totalComparedCellsCount++;
            totalErrorsCount++;
            shiftErrorsCount++;
            const penalty = 10;
            penaltyPointsTotal += penalty;

            errorLog.push({
              sheet: sheetName,
              cell: cellRef,
              rowIndex: empRow,
              colIndex: 0,
              employeeValue: `Block Shift (Row ${empRow + 1})`,
              reviewerValue: "",
              normalizedEmployeeValue: "Block Shift",
              normalizedReviewerValue: "",
              similarity: 0,
              errorType: ErrorType.RowShift,
              severity: Severity.Critical,
              penalty,
              notes: `Confirmed Row Shift: Consecutively inserted extraneous substantial block (${block.count} rows) with no local recovery alignment.`
            });
          }
        }
      }
    }

    sheetErrorDensities[sheetName] = totalErrorsCount;
  }

  // Consolidate sequential row errors so that consecutive row differences are 1 error instead of N
  const consolidatedErrorLog = consolidateSequentialRowErrors(errorLog);

  // Recalculate category counts and penalty points based on consolidated log
  numericErrorsCount = 0;
  missingValuesCount = 0;
  extraValuesCount = 0;
  shiftErrorsCount = 0;
  textErrorsCount = 0;
  rangeErrorsCount = 0;
  headerErrorsCount = 0;
  let recalculatedPenaltyPoints = 0;

  for (const err of consolidatedErrorLog) {
    const isHeaderErr = err.notes.includes("[Header Row Error]");
    if (isHeaderErr) {
      headerErrorsCount++;
    }
    
    recalculatedPenaltyPoints += err.penalty;

    switch (err.errorType) {
      case ErrorType.MissingDigit:
      case ErrorType.ExtraDigit:
      case ErrorType.DigitSubstitution:
      case ErrorType.DigitTransposition:
      case ErrorType.NumericDifference:
      case ErrorType.MajorNumericError:
        numericErrorsCount++;
        break;
      case ErrorType.MissingValue:
      case ErrorType.MissingRow:
      case ErrorType.MissingCell:
      case ErrorType.TableMerge:
        missingValuesCount++;
        break;
      case ErrorType.ExtraValue:
      case ErrorType.ExtraRow:
      case ErrorType.ExtraCell:
      case ErrorType.TableSplit:
        extraValuesCount++;
        break;
      case ErrorType.RowShift:
      case ErrorType.ColumnShift:
      case ErrorType.LocalColumnMisalignment:
      case ErrorType.LocalRowMisalignment:
        shiftErrorsCount++;
        break;
      case ErrorType.TextTypo:
      case ErrorType.MajorTextDifference:
      case ErrorType.TextDifference:
        textErrorsCount++;
        break;
      case ErrorType.RangeInversionError:
      case ErrorType.RangeBoundaryError:
      case ErrorType.RangeRepresentationError:
        rangeErrorsCount++;
        break;
    }
  }

  // Compile final aggregated accuracy metrics including consolidation
  const comparedCells = totalComparedCellsCount;
  const totalErrors = consolidatedErrorLog.length;
  const totalPenaltyPoints = recalculatedPenaltyPoints;

  // Base Accuracy (percentage)
  const baseAccuracy = comparedCells > 0 
    ? parseFloat((((comparedCells - totalErrors) / comparedCells) * 100).toFixed(2))
    : 100.0;

  // Weighted Accuracy aligns with the standard unweighted accuracy
  const weightedAccuracy = baseAccuracy;

  // Error Rate per 10,000 cells
  const errorRatePer10k = comparedCells > 0
    ? parseFloat(((totalErrors / comparedCells) * 10000).toFixed(2))
    : 0.0;

  // Reviewer Workload Index
  // Critical errors = 4.0h, High = 1.0h, Medium = 0.25h, Low = 0.05h
  const criticalCount = consolidatedErrorLog.filter(x => x.severity === Severity.Critical).length;
  const highCount = consolidatedErrorLog.filter(x => x.severity === Severity.High).length;
  const mediumCount = consolidatedErrorLog.filter(x => x.severity === Severity.Medium).length;
  const lowCount = consolidatedErrorLog.filter(x => x.severity === Severity.Low).length;
  const reviewerWorkloadIndex = (criticalCount * 4.00) + (highCount * 1.00) + (mediumCount * 0.25) + (lowCount * 0.05);

  // Performance scoring grade rubric
  let computedGrade = "Outstanding";
  if (weightedAccuracy >= 99.90) computedGrade = "Outstanding";
  else if (weightedAccuracy >= 99.00) computedGrade = "Excellent";
  else if (weightedAccuracy >= 97.00) computedGrade = "Very Good";
  else if (weightedAccuracy >= 95.00) computedGrade = "Good";
  else if (weightedAccuracy >= 90.00) computedGrade = "Fair";
  else if (weightedAccuracy >= 80.00) computedGrade = "Needs Improvement";
  else computedGrade = "Poor";

  // Enforce automatic Safety Overrides
  let finalGrade = computedGrade;
  const hasShifts = allShiftEvents.length > 0;
  if (hasShifts) {
    if (finalGrade !== "Needs Improvement" && finalGrade !== "Poor") {
      finalGrade = "Needs Improvement";
    }
  }

  if (criticalCount > 5) {
    if (finalGrade !== "Needs Improvement" && finalGrade !== "Poor" && finalGrade !== "Fair") {
      finalGrade = "Fair";
    }
  }

  const metrics: QAMetrics = {
    comparedCells,
    totalErrors,
    totalPenaltyPoints,
    baseAccuracy,
    weightedAccuracy,
    errorRatePer10k,
    reviewerWorkloadIndex,
    finalGrade
  };

  // Compile standard Root Cause percentages partition
  const totalClassifiedErrors = totalErrors || 1;
  const rootCause: RootCauseStats = {
    numericErrorsPct: Math.round((numericErrorsCount / totalClassifiedErrors) * 100),
    missingValuesPct: Math.round(((missingValuesCount + extraValuesCount) / totalClassifiedErrors) * 100),
    textErrorsPct: Math.round((textErrorsCount / totalClassifiedErrors) * 100),
    shiftErrorsPct: Math.round((shiftErrorsCount / totalClassifiedErrors) * 100),
    rangeErrorsPct: Math.round((rangeErrorsCount / totalClassifiedErrors) * 100),
    headerErrorsPct: Math.round((headerErrorsCount / totalClassifiedErrors) * 100),
  };

  // Build systematic findings
  const repeatedNumericErrors = Object.entries(numericSubstitutions)
    .filter(([_, count]) => count >= 3)
    .map(([pair, count]) => `Repeated digit mismatch of [${pair}] (occurred ${count} times)`);

  const copyPasteErrors = Object.entries(copyPasteValues)
    .filter(([_, count]) => count >= 3)
    .map(([pair, count]) => `Possible Copy/Paste Error: pattern [${pair}] repeats recursively (occurred ${count} times)`);

  const errorClusters = Object.entries(clustersBySheet)
    .filter(([_, count]) => count >= 5)
    .map(([key, count]) => {
      const parts = key.split("-RowNeighborhood-");
      const sName = parts[0];
      const neighIdx = parseInt(parts[1]) * 5;
      return `Error Cluster: rows ${neighIdx + 1}–${neighIdx + 5} on Sheet '${sName}' contains high density defects (${count} mismatches)`;
    });

  const sheetConcentrations = Object.entries(sheetErrorDensities)
    .filter(([_, count]) => count > 0)
    .map(([name, count]) => `Sheet '${name}' contains ${count} mismatches (${Math.round((count / totalClassifiedErrors) * 100)}% density)`);

  // Include Table Split/Merge events in Pattern Findings
  for (const err of consolidatedErrorLog) {
    if (err.errorType === ErrorType.TableSplit) {
      sheetConcentrations.push(`Table Split Pattern Detected on '${err.sheet}': Split into worker sub-sheets (${err.employeeValue}).`);
    } else if (err.errorType === ErrorType.TableMerge) {
      sheetConcentrations.push(`Table Merge Pattern Detected on '${err.sheet}': Combined multiple reviewer sheets (${err.reviewerValue}).`);
    }
  }

  const patterns: PatternFindings = {
    repeatedNumericErrors,
    copyPasteErrors,
    errorClusters,
    sheetConcentrations,
    shiftEvents: allShiftEvents
  };

  // Actions and coaching recommendations
  const coachingRecommendations: string[] = [];
  if (missingValuesCount > 0 && rootCause.missingValuesPct >= 20) {
    coachingRecommendations.push(
      "Enforce Multi-Pass entry reconciliation. The employee omitted multiple data rows. Adding row-count validations prevents premature submission of truncated tables."
    );
  }
  if (rangeErrorsCount > 0 && rootCause.rangeErrorsPct >= 15) {
    coachingRecommendations.push(
      "Review year-range entry conventions: X/Y format must be entered chronologically and literal delimiters must match exactly to comply with data-entry specifications."
    );
  }
  if (numericErrorsCount > 0 && rootCause.numericErrorsPct >= 25) {
    coachingRecommendations.push(
      "Conduct training on numeric keypad standards with selective self-checking calculations. Encourage double-entry checksums for dense tables."
    );
  }
  if (textErrorsCount > 0 && rootCause.textErrorsPct >= 25) {
    coachingRecommendations.push(
      "Implement spelling check rules. Arabic transcription errors (such as final Heh/Teh Marbuta confusion) can be resolved by enforcing browser spelling plugins."
    );
  }
  if (hasShifts) {
    coachingRecommendations.push(
      "Apply Grid Alignment Validation procedures first. Alignment shift detected due to row skips. Reinforce locking layout anchors (using row header freezes)."
    );
  }

  if (coachingRecommendations.length === 0) {
    coachingRecommendations.push(
      "Demonstrated pristine data-entry compliance standards. Ready for senior auditing delegation."
    );
  }

  return {
    config,
    metrics,
    rootCause,
    errorLog: consolidatedErrorLog,
    patterns,
    coachingRecommendations,
    virtualSheets
  };
}
