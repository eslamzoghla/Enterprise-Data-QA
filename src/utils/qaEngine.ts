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
  AnalysisResult,
  SheetMatch,
  RootCauseCluster
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
 * Converts Arabic numerals (٠١٢٣٤٥٦٧٨٩) to standard Eastern-Arabic/Latin digits (0123456789).
 */
export function convertArabicDigits(str: string): string {
  if (!str) return "";
  const arabicDigitsMap: Record<string, string> = {
    "٠": "0", "١": "1", "٢": "2", "٣": "3", "٤": "4",
    "٥": "5", "٦": "6", "٧": "7", "٨": "8", "٩": "9"
  };
  return str.replace(/[٠١٢٣٤٥٦٧٨٩]/g, (char) => arabicDigitsMap[char] || char);
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

  // NOTE: Te Marbuta (ة) conversion to Heh (ه) is strictly disabled in compliance with the specification
  // as it changes word meanings and results in incorrect fuzzy matching.

  return norm;
}

/**
 * Clean numeric string helpers (removes thousand commas, handles leading zeros)
 */
export function normalizeNumericString(val: string | number): { isNumeric: boolean; value: number | null; cleanedStr: string } {
  if (typeof val === "number") {
    return { isNumeric: true, value: val, cleanedStr: String(val) };
  }
  let trimmed = val.trim();
  // Normalize Arabic digits to Latin numerals first
  trimmed = convertArabicDigits(trimmed);
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
 * Safely parses a string into a normalized ISO date YYYY-MM-DD without Date.parse to maintain browser consistency.
 * Supports formats: YYYY-MM-DD, DD/MM/YYYY, MM/DD/YYYY, and dots/spaces separation.
 */
export function parseDateDeterministic(val: string): { isDate: boolean; formatted: string } {
  const cleaned = convertArabicDigits(val.trim());
  if (!cleaned) return { isDate: false, formatted: "" };

  // 1. Match YYYY-MM-DD pattern (or with slash, dot, space delimiters)
  const ymdRegex = /^(\d{4})[\/\.\-\s](\d{1,2})[\/\.\-\s](\d{1,2})$/;
  let match = cleaned.match(ymdRegex);
  if (match) {
    const year = parseInt(match[1]);
    const month = parseInt(match[2]);
    const day = parseInt(match[3]);
    if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      const formatted = `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
      return { isDate: true, formatted };
    }
  }

  // 2. Match DD/MM/YYYY or MM/DD/YYYY patterns with 4-digit years
  const dmyRegex = /^(\d{1,2})[\/\.\-\s](\d{1,2})[\/\.\-\s](\d{4})$/;
  match = cleaned.match(dmyRegex);
  if (match) {
    const p1 = parseInt(match[1]);
    const p2 = parseInt(match[2]);
    const year = parseInt(match[3]);

    let day = p1;
    let month = p2;

    if (p1 > 12) {
      // Must be DD/MM/YYYY since month cannot exceed 12
      day = p1;
      month = p2;
    } else if (p2 > 12) {
      // Must be MM/DD/YYYY since day exceeds 12
      day = p2;
      month = p1;
    } else {
      // Ambiguous case (both <= 12). Default to DD/MM/YYYY as standard behavior
      day = p1;
      month = p2;
    }

    if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      const formatted = `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
      return { isDate: true, formatted };
    }
  }

  return { isDate: false, formatted: cleaned };
}

/**
 * Standardizes common Date representations into YYYY-MM-DD
 */
export function normalizeDate(val: string): { isDate: boolean; formatted: string } {
  return parseDateDeterministic(val);
}

/**
 * Standardizes time strings (e.g. 14:30:00 vs 14:30) with Eastern-Arabic digit conversion
 */
export function normalizeTime(val: string): { isTime: boolean; formatted: string } {
  const cleaned = convertArabicDigits(val.trim());
  const timeRegex = /^([0-1]?[0-9]|2[0-3]):([0-5][0-9])(:([0-5][0-9]))?\s*(AM|PM)?$/i;
  if (timeRegex.test(cleaned)) {
    return { isTime: true, formatted: cleaned.toLowerCase() };
  }
  return { isTime: false, formatted: cleaned };
}

const levenshteinCache = new Map<string, number>();

/**
 * Computes Levenshtein Distance between s1 and s2 with memoization and length-difference early exit
 */
export function levenshteinDistance(s1: string, s2: string): number {
  const cacheKey = s1 + "|||" + s2;
  if (levenshteinCache.has(cacheKey)) {
    return levenshteinCache.get(cacheKey)!;
  }
  const revCacheKey = s2 + "|||" + s1;
  if (levenshteinCache.has(revCacheKey)) {
    return levenshteinCache.get(revCacheKey)!;
  }

  // Early-exit for highly dissimilar lengths to optimize performance
  if (Math.abs(s1.length - s2.length) > 10) {
    const minPossibleDistance = Math.abs(s1.length - s2.length);
    levenshteinCache.set(cacheKey, minPossibleDistance);
    return minPossibleDistance;
  }

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
  const result = matrix[s2.length][s1.length];
  levenshteinCache.set(cacheKey, result);
  return result;
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
 * Computes a robust match confidence score (0 to 100) based on name similarity, headers, structure, and content.
 */
export function computeMatchConfidence(
  reviewerSheet: SheetGrid,
  employeeSheet: SheetGrid
): {
  confidence: number;
  factors: {
    nameSimilarity: number;
    headerSimilarity: number;
    structuralSimilarity: number;
    dataOverlap: number;
  }
} {
  // 1. Name Similarity
  const name1 = reviewerSheet.name;
  const name2 = employeeSheet.name;
  const nameSim = calculateTextSimilarity(name1, name2);

  // 2. Header Similarity
  const headers1Set = detectHeaderRows(reviewerSheet);
  const headers2Set = detectHeaderRows(employeeSheet);
  
  const h1Vals: string[] = [];
  headers1Set.forEach(row => {
    for (let c = 0; c <= reviewerSheet.maxCol; c++) {
      const cell = reviewerSheet.cells[`${row},${c}`];
      if (cell && cell.normalized) {
        h1Vals.push(cell.normalized.trim().toLowerCase());
      }
    }
  });

  const h2Vals: string[] = [];
  headers2Set.forEach(row => {
    for (let c = 0; c <= employeeSheet.maxCol; c++) {
      const cell = employeeSheet.cells[`${row},${c}`];
      if (cell && cell.normalized) {
        h2Vals.push(cell.normalized.trim().toLowerCase());
      }
    }
  });

  let headerSim = 0;
  if (h1Vals.length === 0 && h2Vals.length === 0) {
    headerSim = 100.0;
  } else if (h1Vals.length === 0 || h2Vals.length === 0) {
    headerSim = 20.0;
  } else {
    const s1 = new Set(h1Vals);
    const s2 = new Set(h2Vals);
    let common = 0;
    s1.forEach(v => {
      if (s2.has(v)) common++;
    });
    headerSim = (common / Math.max(s1.size, s2.size)) * 100;
  }

  // 3. Structural Similarity (diff of row/col bounds)
  const rDiff = Math.abs(reviewerSheet.maxRow - employeeSheet.maxRow);
  const cDiff = Math.abs(reviewerSheet.maxCol - employeeSheet.maxCol);
  const rMax = Math.max(reviewerSheet.maxRow, employeeSheet.maxRow, 1);
  const cMax = Math.max(reviewerSheet.maxCol, employeeSheet.maxCol, 1);
  const structSim = Math.max(0, 100 * (1 - (rDiff / rMax) * 0.5 - (cDiff / cMax) * 0.5));

  // 4. Data Content Overlap
  const dataOverlap = computeContentSimilarity(reviewerSheet, employeeSheet);

  // Weighted Combination: 20% name, 20% header, 20% struct, 40% content
  const confidence = parseFloat((
    (nameSim * 0.20) +
    (headerSim * 0.20) +
    (structSim * 0.20) +
    (dataOverlap * 0.40)
  ).toFixed(1));

  return {
    confidence: Math.max(0, Math.min(100, confidence)),
    factors: {
      nameSimilarity: parseFloat(nameSim.toFixed(1)),
      headerSimilarity: parseFloat(headerSim.toFixed(1)),
      structuralSimilarity: parseFloat(structSim.toFixed(1)),
      dataOverlap: parseFloat(dataOverlap.toFixed(1))
    }
  };
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

  const timeNorm = normalizeTime(spacedNorm);
  if (timeNorm.isTime) {
    return {
      raw,
      formatted: spacedNorm,
      normalized: timeNorm.formatted,
      type: "time"
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
export function shouldExcludeSheet(
  sheetName: string, 
  maxRow: number, 
  cellsCount: number,
  autoIgnoreEnabled?: boolean,
  customIgnorePatterns?: string
): boolean {
  if (!autoIgnoreEnabled || !customIgnorePatterns) {
    return false;
  }

  const normalizedPatterns = customIgnorePatterns
    .split(",")
    .map(p => p.trim())
    .filter(p => p.length > 0);

  for (const pat of normalizedPatterns) {
    // Check if it's a regex-style string wrapped in slashes
    if (pat.startsWith("/") && pat.endsWith("/")) {
      try {
        const regexStr = pat.slice(1, -1);
        const regex = new RegExp(regexStr, "i");
        if (regex.test(sheetName)) {
          return true;
        }
      } catch (err) {
        // Fallback
      }
    }

    // Support standard regex symbols without explicit slashes
    if (pat.includes("*") || pat.includes("?") || pat.includes("^") || pat.includes("$")) {
      try {
        const regex = new RegExp(pat, "i");
        if (regex.test(sheetName)) {
          return true;
        }
      } catch (err) {
        // Fallback
      }
    }

    // Exact case-insensitive match
    if (sheetName.toLowerCase() === pat.toLowerCase()) {
      return true;
    }

    // Substring match
    if (sheetName.toLowerCase().includes(pat.toLowerCase())) {
      return true;
    }
  }

  return false;
}

/**
 * Auto-detect header rows by row index, styling candidate heuristics or string dominance.
 */
export function detectHeaderRows(sheet: SheetGrid): Set<number> {
  const headers = new Set<number>();
  if (sheet.maxRow < 0) return headers;

  const rowScores: { row: number; score: number; nonEmpty: number }[] = [];
  const maxInspect = Math.min(10, sheet.maxRow); // Inspect up to top 10 rows

  for (let r = 0; r <= maxInspect; r++) {
    let stringCount = 0;
    let numberCount = 0;
    let emptyCount = 0;
    let totalCount = 0;
    const values: string[] = [];

    for (let c = 0; c <= sheet.maxCol; c++) {
      const cell = sheet.cells[`${r},${c}`];
      totalCount++;
      if (!cell || cell.type === "empty" || !cell.normalized.trim()) {
        emptyCount++;
      } else {
        values.push(cell.normalized.trim());
        if (cell.type === "string") {
          stringCount++;
        } else if (cell.type === "number") {
          numberCount++;
        }
      }
    }

    const nonEmptyCount = values.length;
    if (nonEmptyCount <= 1) {
      // Either blank row (0 cells) or title row (1 cell), highly unlikely to be the main table header
      rowScores.push({ row: r, score: -100, nonEmpty: nonEmptyCount });
      continue;
    }

    // 1. Uniqueness score (0.0 to 1.0)
    const uniqueVals = new Set(values);
    const uniqueness = uniqueVals.size / nonEmptyCount;

    // 2. Text percentage (0.0 to 1.0)
    const textPct = stringCount / nonEmptyCount;

    // 3. Density / Column fill
    const fillRate = nonEmptyCount / totalCount;

    // Calculate a base score
    let score = 0;
    score += textPct * 40;       // Up to 40 points for being text-dominant
    score += uniqueness * 30;    // Up to 30 points for having unique values
    score += fillRate * 20;      // Up to 20 points for being fully populated

    // Penalty for numbers
    const numberPct = numberCount / nonEmptyCount;
    score -= numberPct * 40;     // Huge penalty if it contains mostly numbers (typical data rows)

    // 4. Check neighborhood signal: style/uniqueness contrast with the next row (r+1)
    if (r < sheet.maxRow) {
      let nextRowStringCount = 0;
      let nextRowNumberCount = 0;
      let nextRowNonEmpty = 0;
      for (let c = 0; c <= sheet.maxCol; c++) {
        const nextCell = sheet.cells[`${r+1},${c}`];
        if (nextCell && nextCell.type !== "empty" && nextCell.normalized.trim()) {
          nextRowNonEmpty++;
          if (nextCell.type === "string") {
            nextRowStringCount++;
          } else if (nextCell.type === "number") {
            nextRowNumberCount++;
          }
        }
      }
      if (nextRowNonEmpty > 0) {
        const nextRowNumberPct = nextRowNumberCount / nextRowNonEmpty;
        // If the row below is full of numbers, then this row r has a higher chance of being the header
        score += nextRowNumberPct * 20;
      }
    }

    rowScores.push({ row: r, score, nonEmpty: nonEmptyCount });
  }

  // Find the row with the maximum score
  let bestRow = 0;
  let maxScore = -999;
  for (const item of rowScores) {
    if (item.score > maxScore && item.nonEmpty > 1) {
      maxScore = item.score;
      bestRow = item.row;
    }
  }

  // If we found a clear header row, put it in the headers Set
  if (maxScore > 0) {
    headers.add(bestRow);
    
    // Check if there are immediately following rows that are ALSO headers (sub-headers)
    // They must have predominantly text cells and decent uniqueness
    for (let r = bestRow + 1; r <= Math.min(bestRow + 2, sheet.maxRow); r++) {
      const scoreObj = rowScores.find(item => item.row === r);
      if (scoreObj && scoreObj.score > 25) {
        headers.add(r);
      }
    }
  } else {
    // Fallback: row 0 is header
    headers.add(0);
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
  const normalizedText = convertArabicDigits(val.trim());
  // Unicode regex with /u flag allowing letters (\p{L}) and numbers (\p{N}) across all language blocks
  const regex = /^([\p{L}\p{N}]+)([\/\-–—])([\p{L}\p{N}]+)$/u;
  const match = normalizedText.match(regex);
  if (match) {
    let left = match[1];
    let right = match[3];

    // Normalize Arabic letters if present (e.g. أ1 -> ا1) for consistent boundaries
    const leftHasArabic = /[\u0600-\u06FF]/.test(left);
    if (leftHasArabic) {
      left = normalizeArabicText(left);
    }
    const rightHasArabic = /[\u0600-\u06FF]/.test(right);
    if (rightHasArabic) {
      right = normalizeArabicText(right);
    }

    return {
      isValid: true,
      left,
      right,
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
/**
 * Creates an efficient 2D array representation of sheet cells for fast lookup.
 */
export function getSheetCellsMatrix(sheet: SheetGrid): (CellValue | undefined)[][] {
  const matrix: (CellValue | undefined)[][] = [];
  for (let r = 0; r <= sheet.maxRow; r++) {
    const rowList: (CellValue | undefined)[] = [];
    for (let c = 0; c <= sheet.maxCol; c++) {
      rowList.push(sheet.cells[`${r},${c}`]);
    }
    matrix.push(rowList);
  }
  return matrix;
}

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

  // Access via cached 2D matrix to prevent string key memory leak and performance bottleneck
  const empMatrix = getSheetCellsMatrix(empGrid);
  const revMatrix = getSheetCellsMatrix(revGrid);

  const getNorm = (matrix: (CellValue | undefined)[][], r: number, c: number) => {
    return matrix[r]?.[c]?.normalized || "";
  };

  const getCellType = (matrix: (CellValue | undefined)[][], r: number, c: number) => {
    return matrix[r]?.[c]?.type || "empty";
  };

  // --- Dynamic Column Shift Offset Range up to 20 (Requirement 4) ---
  const offsets: number[] = [];
  const maxColOffset = Math.min(20, maxCol);
  for (let i = 1; i <= maxColOffset; i++) {
    offsets.push(i);
    offsets.push(-i);
  }

  // Detect Column Shifts
  for (const dc of offsets) {
    let contiguousCols: number[] = [];

    for (let c = 0; c <= maxCol; c++) {
      let matches = 0;
      let compared = 0;

      for (let r = 0; r <= maxRow; r++) {
        const empVal = getNorm(empMatrix, r, c);
        const revVal = getNorm(revMatrix, r, c + dc);

        const empType = getCellType(empMatrix, r, c);
        const revType = getCellType(revMatrix, r, c + dc);

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
          if (getCellType(empMatrix, r, c) !== "empty" || getCellType(revMatrix, r, c + offset_c) !== "empty") {
            totalSegmentCells++;
            if (getNorm(empMatrix, r, c) === getNorm(revMatrix, r, c + offset_c)) {
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

  // --- Dynamic Row Shift Offset Range up to 20 (Requirement 5) ---
  const rowOffsets: number[] = [];
  const maxRowOffset = Math.min(20, maxRow);
  for (let i = 1; i <= maxRowOffset; i++) {
    rowOffsets.push(i);
    rowOffsets.push(-i);
  }

  // Detect Row Shifts
  for (const dr of rowOffsets) {
    let contiguousRows: number[] = [];

    for (let r = 0; r <= maxRow; r++) {
      let matches = 0;
      let compared = 0;

      for (let c = 0; c <= maxCol; c++) {
        const empVal = getNorm(empMatrix, r, c);
        const revVal = getNorm(revMatrix, r + dr, c);

        const empType = getCellType(empMatrix, r, c);
        const revType = getCellType(revMatrix, r + dr, c);

        if (empType !== "empty" || revType !== "empty") {
          compared++;
          if (empVal === revVal) {
            matches++;
          }
        }
      }

      if (compared > 0 && matches / compared >= shiftThreshold) {
        contiguousRows.push(r);
      } else {
        if (contiguousRows.length > 0) {
          evaluateRowGroup(contiguousRows, dr);
          contiguousRows = [];
        }
      }
    }
    if (contiguousRows.length > 0) {
      evaluateRowGroup(contiguousRows, dr);
    }

    function evaluateRowGroup(rows: number[], offset_r: number) {
      let totalSegmentCells = 0;
      let matchesCount = 0;
      for (const r of rows) {
        for (let c = 0; c <= maxCol; c++) {
          if (getCellType(empMatrix, r, c) !== "empty" || getCellType(revMatrix, r + offset_r, c) !== "empty") {
            totalSegmentCells++;
            if (getNorm(empMatrix, r, c) === getNorm(revMatrix, r + offset_r, c)) {
              matchesCount++;
            }
          }
        }
      }

      if (totalSegmentCells >= minShiftCells) {
        const start = rows[0];
        const end = rows[rows.length - 1];

        const confidenceScore = totalSegmentCells > 0 ? Math.round((matchesCount / totalSegmentCells) * 100) : 100;

        events.push({
          sheetName: empGrid.name,
          type: "row",
          offset: offset_r,
          spanStart: start,
          spanEnd: end,
          detail: `Row Alignment Shift: Rows ${start + 1} to ${end + 1} shifted vertically by ${offset_r} space(s) [Confidence: ${confidenceScore}%, cells affected: ${totalSegmentCells}]`,
          affectedCellsCount: totalSegmentCells
        });

        for (const r of rows) {
          for (let c = 0; c <= maxCol; c++) {
            shiftedCoords[`${r},${c}`] = "row";
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
  // FIX: Only strip trailing single letter if it's preceded by a DIGIT (e.g. "table 17 a" -> "table 17")
  // Do NOT strip letters from words like "table" -> "tabl"
  s = s.replace(/(\d)\s*[-_]?\s*[a-z]\s*$/, "$1");
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
 * Concatenates multiple SheetGrids HORIZONTALLY (side by side, same rows).
 * Used to detect when an employee merges two reviewer tables into one wide sheet.
 */
export function horizontalConcatenateSheets(sheets: SheetGrid[]): SheetGrid {
  const combined: SheetGrid = {
    name: "HorizontalCombined",
    maxRow: -1,
    maxCol: -1,
    cells: {},
    virtualSegments: []
  };
  let currentColOffset = 0;
  for (const sh of sheets) {
    const startCol = currentColOffset;
    if (sh.maxRow > combined.maxRow) {
      combined.maxRow = sh.maxRow;
    }
    for (let r = 0; r <= sh.maxRow; r++) {
      for (let c = 0; c <= sh.maxCol; c++) {
        const cellVal = sh.cells[`${r},${c}`];
        if (cellVal) {
          combined.cells[`${r},${currentColOffset + c}`] = cellVal;
        }
      }
    }
    currentColOffset += sh.maxCol + 1;
    combined.virtualSegments!.push({
      originalSheetName: sh.name,
      virtualStartRow: 0,
      virtualEndRow: sh.maxRow,
      originalStartRow: startCol  // reusing field to store col offset
    });
  }
  combined.maxCol = currentColOffset - 1;
  return combined;
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

  // New Audit Structural Counters
  let extraTablesCount = 0;
  let missingTablesCount = 0;
  let extraColumnsCount = 0;
  let missingColumnsCount = 0;
  let extraRowsCount = 0;
  let missingRowsCount = 0;

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
    return !shouldExcludeSheet(
      sheetName, 
      activeSheet.maxRow, 
      Object.keys(activeSheet.cells).length,
      config.autoIgnoreEnabled,
      config.customIgnorePatterns
    );
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
    if (shouldExcludeSheet(
      name, 
      sheet.maxRow, 
      Object.keys(sheet.cells).length,
      config.autoIgnoreEnabled,
      config.customIgnorePatterns
    )) return false;
    return !employeeData.sheets[name];
  });

  const extraEmployeeSheets = Object.keys(employeeData.sheets).filter(name => {
    const sheet = employeeData.sheets[name];
    if (shouldExcludeSheet(
      name, 
      sheet.maxRow, 
      Object.keys(sheet.cells).length,
      config.autoIgnoreEnabled,
      config.customIgnorePatterns
    )) return false;
    return !reviewerData.sheets[name];
  });

  // Group employee sheets by normalized base name:
  const empGroups: Record<string, string[]> = {};
  for (const name of Object.keys(employeeData.sheets)) {
    const sheet = employeeData.sheets[name];
    if (shouldExcludeSheet(
      name, 
      sheet.maxRow, 
      Object.keys(sheet.cells).length,
      config.autoIgnoreEnabled,
      config.customIgnorePatterns
    )) continue;
    const base = getNormalizedBaseName(name);
    if (!empGroups[base]) empGroups[base] = [];
    empGroups[base].push(name);
  }

  // Group reviewer sheets by normalized base name:
  const revGroups: Record<string, string[]> = {};
  for (const name of Object.keys(reviewerData.sheets)) {
    const sheet = reviewerData.sheets[name];
    if (shouldExcludeSheet(
      name, 
      sheet.maxRow, 
      Object.keys(sheet.cells).length,
      config.autoIgnoreEnabled,
      config.customIgnorePatterns
    )) continue;
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

      // Try vertical concatenation first, then horizontal
      const combinedVertical = concatenateSheets(revSheets.map(name => reviewerData.sheets[name]));
      const combinedHorizontal = horizontalConcatenateSheets(revSheets.map(name => reviewerData.sheets[name]));
      const simVertical = computeContentSimilarity(employeeData.sheets[empName], combinedVertical);
      const simHorizontal = computeContentSimilarity(employeeData.sheets[empName], combinedHorizontal);
      const sim = Math.max(simVertical, simHorizontal);
      const combined = simHorizontal >= simVertical ? combinedHorizontal : combinedVertical;
      const mergeOrientation = simHorizontal >= simVertical ? "horizontal" : "vertical";

      if (sim >= 95.0) {
        processedEmployeeSheets.add(empName);
        revSheets.forEach(name => processedReviewerSheets.add(name));

        totalErrorsCount++;
        penaltyPointsTotal += 10;

        // Count employee sheet cell coverage
        totalComparedCellsCount += Object.keys(employeeData.sheets[empName].cells).length;

        const noteMsg = `Table Merge Event (${mergeOrientation}): Reviewer sheets [${revSheets.join(", ")}] were merged into employee sheet '${empName}'. Content similarity is ${sim}%.`;
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

      // Try vertical then horizontal concatenation
      const combinedV = concatenateSheets(candidates.map(name => reviewerData.sheets[name]));
      const combinedH = horizontalConcatenateSheets(candidates.map(name => reviewerData.sheets[name]));
      const simV = computeContentSimilarity(employeeData.sheets[empName], combinedV);
      const simH = computeContentSimilarity(employeeData.sheets[empName], combinedH);
      const sim = Math.max(simV, simH);
      const combined = simH >= simV ? combinedH : combinedV;
      const mergeOrientation = simH >= simV ? "horizontal" : "vertical";

      if (sim >= 95.0) {
        processedEmployeeSheets.add(empName);
        candidates.forEach(name => processedReviewerSheets.add(name));

        totalErrorsCount++;
        penaltyPointsTotal += 10;

        // Count employee sheet cell coverage
        totalComparedCellsCount += Object.keys(employeeData.sheets[empName].cells).length;

        const noteMsg = `Table Merge Event (${mergeOrientation}): Reviewer sheets [${candidates.join(", ")}] were merged into employee sheet '${empName}'. Content similarity is ${sim}%.`;
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
      // REQUIREMENT 5: Structural suppression guarantee
      const isRelatedToSplitOrMerge = processedReviewerSheets.has(sheetName) || processedEmployeeSheets.has(sheetName);
      if (isRelatedToSplitOrMerge) {
        continue;
      }

      const isEmployeeOnly = !!empSheet;
      const penalty = isEmployeeOnly ? config.extraTableCoefficient : config.missingTableCoefficient;
      const errType = isEmployeeOnly ? ErrorType.ExtraTable : ErrorType.MissingTable;
      const notes = isEmployeeOnly
        ? `Extra Table: Employee table '${sheetName}' has no legitimate counterpart in the Reviewer workbook.`
        : `Missing Table: Reviewer table '${sheetName}' is omitted in the Employee workbook.`;

      if (isEmployeeOnly) {
        extraTablesCount++;
      } else {
        missingTablesCount++;
      }

      errorLog.push({
        sheet: sheetName,
        cell: "Sheet Layout",
        rowIndex: 0,
        colIndex: 0,
        employeeValue: isEmployeeOnly ? sheetName : "",
        normalizedEmployeeValue: isEmployeeOnly ? sheetName : "",
        reviewerValue: isEmployeeOnly ? "" : sheetName,
        normalizedReviewerValue: isEmployeeOnly ? "" : sheetName,
        similarity: 0,
        errorType: errType,
        severity: Severity.High,
        penalty,
        notes
      });
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

    if (!isVirtualComparison) {
      let NumberOfExtraColumns = 0;
      let NumberOfMissingColumns = 0;
      for (const colBlock of colStepBlocks) {
        if (colBlock.type === "extra_column") {
          NumberOfExtraColumns += colBlock.count;
        } else if (colBlock.type === "missing_column") {
          NumberOfMissingColumns += colBlock.count;
        }
      }

      if (NumberOfExtraColumns > 0) {
        extraColumnsCount += NumberOfExtraColumns;
        errorLog.push({
          sheet: sheetName,
          cell: "Columns Layout",
          rowIndex: 0,
          colIndex: 0,
          employeeValue: `${NumberOfExtraColumns} extra columns`,
          reviewerValue: "",
          normalizedEmployeeValue: `${NumberOfExtraColumns}`,
          normalizedReviewerValue: "",
          similarity: 0,
          errorType: ErrorType.ExtraColumns,
          severity: Severity.High,
          penalty: NumberOfExtraColumns * config.extraColumnCoefficient,
          notes: `Extra Columns: Table contains ${NumberOfExtraColumns} additional column(s) not present in Reviewer template.`
        });
      }

      if (NumberOfMissingColumns > 0) {
        missingColumnsCount += NumberOfMissingColumns;
        errorLog.push({
          sheet: sheetName,
          cell: "Columns Layout",
          rowIndex: 0,
          colIndex: 0,
          employeeValue: "",
          reviewerValue: `${NumberOfMissingColumns} missing columns`,
          normalizedEmployeeValue: "",
          normalizedReviewerValue: `${NumberOfMissingColumns}`,
          similarity: 0,
          errorType: ErrorType.MissingColumns,
          severity: Severity.High,
          penalty: NumberOfMissingColumns * config.missingColumnCoefficient,
          notes: `Missing Columns: Table is missing ${NumberOfMissingColumns} required column(s).`
        });
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

    if (!isVirtualComparison) {
      let NumberOfExtraRows = 0;
      let NumberOfMissingRows = 0;
      for (const block of stepBlocks) {
        if (block.type === "extra_row") {
          NumberOfExtraRows += block.count;
        } else if (block.type === "missing_row") {
          NumberOfMissingRows += block.count;
        }
      }
      extraRowsCount += NumberOfExtraRows;
      missingRowsCount += NumberOfMissingRows;
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

  // Compute Data Category differences from consolidated error log
  let numericDifferencesCount = 0;
  let textDifferencesCount = 0;
  let emptyCellDifferencesCount = 0;

  for (const err of consolidatedErrorLog) {
    if (
      err.errorType === ErrorType.ExtraTable ||
      err.errorType === ErrorType.MissingTable ||
      err.errorType === ErrorType.ExtraColumns ||
      err.errorType === ErrorType.MissingColumns
    ) {
      continue;
    }

    switch (err.errorType) {
      case ErrorType.MissingValue:
      case ErrorType.ExtraValue:
      case ErrorType.MissingCell:
      case ErrorType.ExtraCell:
        emptyCellDifferencesCount++;
        break;
      case ErrorType.MissingDigit:
      case ErrorType.ExtraDigit:
      case ErrorType.DigitSubstitution:
      case ErrorType.DigitTransposition:
      case ErrorType.NumericDifference:
      case ErrorType.MajorNumericError:
        numericDifferencesCount++;
        break;
      case ErrorType.TextDifference:
      case ErrorType.TextTypo:
      case ErrorType.MajorTextDifference:
      case ErrorType.RangeInversionError:
      case ErrorType.RangeBoundaryError:
      case ErrorType.RangeRepresentationError:
        textDifferencesCount++;
        break;
    }
  }

  const structuralPenalty = 
    (extraTablesCount * config.extraTableCoefficient) +
    (missingTablesCount * config.missingTableCoefficient) +
    (extraColumnsCount * config.extraColumnCoefficient) +
    (missingColumnsCount * config.missingColumnCoefficient) +
    (extraRowsCount * config.extraRowCoefficient) +
    (missingRowsCount * config.missingRowCoefficient);

  const dataPenalty = 
    (numericDifferencesCount * config.numericDifferenceCoefficient) +
    (textDifferencesCount * config.textDifferenceCoefficient) +
    (emptyCellDifferencesCount * config.emptyCellDifferenceCoefficient);

  const totalPenalty = structuralPenalty + dataPenalty;

  const structuralScore = Math.max(0, Math.min(100, 100 - structuralPenalty));
  const dataScore = Math.max(0, Math.min(100, 100 - dataPenalty));
  const finalAuditScore = (structuralScore * 0.40) + (dataScore * 0.60);

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
    finalGrade,

    structuralPenalty,
    dataPenalty,
    totalPenalty,
    structuralScore,
    dataScore,
    finalAuditScore,

    extraTablesCount,
    missingTablesCount,
    extraColumnsCount,
    missingColumnsCount,
    extraRowsCount,
    missingRowsCount,
    numericDifferencesCount,
    textDifferencesCount,
    emptyCellDifferencesCount,

    // Add compatibility properties of metrics to resolve dashboard errors
    accuracyRatio: baseAccuracy,
    totalStructuralPenalty: totalPenalty,
    grade: finalGrade
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

  // REQUIREMENT 1 & 2: Compile mapping and confidence scores
  const sheetMatchesFromRun: SheetMatch[] = [];

  for (const revName of Object.keys(reviewerData.sheets)) {
    const revSheet = reviewerData.sheets[revName];
    if (shouldExcludeSheet(
      revName, 
      revSheet.maxRow, 
      Object.keys(revSheet.cells).length,
      config.autoIgnoreEnabled,
      config.customIgnorePatterns
    )) {
      continue;
    }

    // 1. Table Split
    // Check if was processed as a TABLE SPLIT:
    const revBase = getNormalizedBaseName(revName);
    const relatedEmps = empGroups[revBase] || [];
    const splitMatch = comparisonPairs.find(p => p.sheetName === revName && p.comparisonType === "split");
    if (splitMatch && relatedEmps.length >= 2) {
      const combEmp = splitMatch.empSheet!;
      const confData = computeMatchConfidence(revSheet, combEmp);
      sheetMatchesFromRun.push({
        reviewerSheetName: revName,
        matchedEmployeeSheets: relatedEmps,
        matchType: "Split Match",
        matchConfidence: confData.confidence,
        factors: confData.factors
      });
      continue;
    }

    // 2. Table Merge
    const mergeMatch = comparisonPairs.find(p => p.comparisonType === "merge" && p.revSheet && Object.keys(employeeData.sheets).some(empKey => {
      const empS = employeeData.sheets[empKey];
      if (!empS) return false;
      const rGroupsList = revGroups[getNormalizedBaseName(empKey)] || [];
      return rGroupsList.includes(revName);
    }));

    if (mergeMatch) {
      const empName = mergeMatch.sheetName;
      const empSheet = employeeData.sheets[empName];
      const confData = computeMatchConfidence(revSheet, empSheet);
      sheetMatchesFromRun.push({
        reviewerSheetName: revName,
        matchedEmployeeSheets: [empName],
        matchType: "Merge Match",
        matchConfidence: confData.confidence,
        factors: confData.factors
      });
      continue;
    }

    // 3. Direct / Fuzzy standard match
    const empSheet = employeeData.sheets[revName];
    if (empSheet) {
      const confData = computeMatchConfidence(revSheet, empSheet);
      const isExact = revName.toLowerCase() === empSheet.name.toLowerCase() && confData.confidence >= 95.0;
      sheetMatchesFromRun.push({
        reviewerSheetName: revName,
        matchedEmployeeSheets: [revName],
        matchType: isExact ? "Direct Match" : "Fuzzy Match",
        matchConfidence: confData.confidence,
        factors: confData.factors
      });
      continue;
    }

    // 4. Other fuzzy matches in employee workbook
    let bestCandidate: string | null = null;
    let bestConf = 0;
    let bestFactors = { nameSimilarity: 0, headerSimilarity: 0, structuralSimilarity: 0, dataOverlap: 0 };

    for (const empKey of Object.keys(employeeData.sheets)) {
      const eSheet = employeeData.sheets[empKey];
      if (shouldExcludeSheet(
        empKey, 
        eSheet.maxRow, 
        Object.keys(eSheet.cells).length, 
        config.autoIgnoreEnabled, 
        config.customIgnorePatterns
      )) {
        continue;
      }
      const confData = computeMatchConfidence(revSheet, eSheet);
      if (confData.confidence > bestConf && confData.confidence >= 25.0) {
        bestConf = confData.confidence;
        bestCandidate = empKey;
        bestFactors = confData.factors;
      }
    }

    if (bestCandidate) {
      sheetMatchesFromRun.push({
        reviewerSheetName: revName,
        matchedEmployeeSheets: [bestCandidate],
        matchType: "Fuzzy Match",
        matchConfidence: bestConf,
        factors: bestFactors
      });
    } else {
      // Unmatched
      sheetMatchesFromRun.push({
        reviewerSheetName: revName,
        matchedEmployeeSheets: [],
        matchType: "Fuzzy Match",
        matchConfidence: 0,
        factors: {
          nameSimilarity: 0,
          headerSimilarity: 0,
          structuralSimilarity: 0,
          dataOverlap: 0
        }
      });
    }
  }

  // REQUIREMENT 4: Compile root cause clusters
  const compiledClusters = compileRootCauseClusters(consolidatedErrorLog, allShiftEvents, patterns);

  return {
    config,
    metrics,
    rootCause,
    errorLog: consolidatedErrorLog,
    patterns,
    coachingRecommendations,
    virtualSheets,
    sheetMatches: sheetMatchesFromRun,
    rootCauseClusters: compiledClusters
  };
}

/**
 * Groups and analyzes the error log and patterns into root cause clusters
 */
export function compileRootCauseClusters(
  errorLog: ErrorLogEntry[],
  shiftEvents: ShiftEvent[],
  patterns: PatternFindings
): RootCauseCluster[] {
  const clusters: RootCauseCluster[] = [];
  let clusterIdCounter = 1;

  // 1. Row shifts clustering
  const rowShifts = shiftEvents.filter(s => s.type === "row");
  for (const s of rowShifts) {
    clusters.push({
      id: `cluster-${clusterIdCounter++}`,
      type: "row_shift",
      sheetName: s.sheetName,
      title: "Data Row Shift Phenomenon Detected",
      description: `Likely Root Cause: Data shifted by ${s.offset > 0 ? "+" : ""}${s.offset} rows starting around Row ${s.spanStart + 1}. This offset cascade affects ${s.affectedCellsCount} cells before recovery or end-of-grid.`,
      symptomCount: s.affectedCellsCount,
      severity: "Critical"
    });
  }

  // 2. Column shifts clustering
  const colShifts = shiftEvents.filter(s => s.type === "column");
  for (const s of colShifts) {
    // Determine column letter from zero-indexed col index
    let colLet = "?";
    try {
      const charCode = 65 + (s.spanStart % 26);
      colLet = String.fromCharCode(charCode);
    } catch(e) {}

    clusters.push({
      id: `cluster-${clusterIdCounter++}`,
      type: "column_shift",
      sheetName: s.sheetName,
      title: "Column Insertion or Layout Shift Detected",
      description: `Likely Root Cause: Column insertion or misalignment detected around column index ${colLet}. Auditor recommendation: verify if a column was inserted or deleted.`,
      symptomCount: s.affectedCellsCount,
      severity: "Critical"
    });
  }

  // 3. Repeated values / Copy-paste errors
  if (patterns.copyPasteErrors && patterns.copyPasteErrors.length > 0) {
    for (const cp of patterns.copyPasteErrors) {
      clusters.push({
        id: `cluster-${clusterIdCounter++}`,
        type: "copy_paste",
        sheetName: "All Workbook",
        title: "Systematic Copy-Paste Mismatch",
        description: `Likely Root Cause: Repeating block copying or formula dragging error. Pattern occurrences: ${cp}.`,
        symptomCount: 5,
        severity: "High"
      });
    }
  }

  // 4. Repeated numeric substitutions
  if (patterns.repeatedNumericErrors && patterns.repeatedNumericErrors.length > 0) {
    for (const d of patterns.repeatedNumericErrors) {
      clusters.push({
        id: `cluster-${clusterIdCounter++}`,
        type: "value_substitution",
        sheetName: "All Workbook",
        title: "Keyboard Numerical Entry Defect Pattern",
        description: `Likely Root Cause: Numpad transposition or repeated digit entry oversight. Pattern: ${d}.`,
        symptomCount: 3,
        severity: "Medium"
      });
    }
  }

  // 5. Layout splits/merges
  const splitsMet = errorLog.filter(e => e.errorType === ErrorType.TableSplit);
  for (const s of splitsMet) {
    clusters.push({
      id: `cluster-${clusterIdCounter++}`,
      type: "structural_mismatch",
      sheetName: s.sheet,
      title: "Multi-Sheet Table Split Event",
      description: `Likely Root Cause: The reviewer table was structured into multiple physical sheets: [${s.employeeValue}] in the worker file. Suppressed baseline missing/extra table penalties.`,
      symptomCount: 1,
      severity: "High"
    });
  }

  const mergesMet = errorLog.filter(e => e.errorType === ErrorType.TableMerge);
  for (const m of mergesMet) {
    clusters.push({
      id: `cluster-${clusterIdCounter++}`,
      type: "structural_mismatch",
      sheetName: m.sheet,
      title: "Workbook Structural Table Merge Event",
      description: `Likely Root Cause: Multiple physical reviewer sheets were combined into a singular sheet: '${m.sheet}' in the worker file. Standard standard missing table row/column penalties suppressed.`,
      symptomCount: 1,
      severity: "High"
    });
  }

  // 6. Template Cell Misalignment Empty Cell Clusters
  const emptyCellErrors = errorLog.filter(e => e.errorType === ErrorType.MissingValue || e.errorType === ErrorType.ExtraValue || e.errorType === ErrorType.MissingCell || e.errorType === ErrorType.ExtraCell);
  const sheetsEmptyCounts: Record<string, number> = {};
  for (const e of emptyCellErrors) {
    sheetsEmptyCounts[e.sheet] = (sheetsEmptyCounts[e.sheet] || 0) + 1;
  }
  for (const [sName, count] of Object.entries(sheetsEmptyCounts)) {
    if (count >= 15) {
      clusters.push({
        id: `cluster-${clusterIdCounter++}`,
        type: "empty_cell_dense",
        sheetName: sName,
        title: "Systemic Empty Cell Density Hotspot",
        description: `Likely Root Cause: Excessive blank/populated cell disagreement (${count} symptoms). This is commonly caused by an offset template draft or incomplete data entry in dense regions.`,
        symptomCount: count,
        severity: "High"
      });
    }
  }

  // Fallback
  if (clusters.length === 0 && errorLog.length > 0) {
    clusters.push({
      id: `cluster-${clusterIdCounter++}`,
      type: "general",
      sheetName: "All Workbook",
      title: "Localized Random Defects",
      description: "Likely Root Cause: Scattered isolated typos with no systematic row or column alignment propagation trends.",
      symptomCount: errorLog.length,
      severity: "Low"
    });
  }

  return clusters;
}