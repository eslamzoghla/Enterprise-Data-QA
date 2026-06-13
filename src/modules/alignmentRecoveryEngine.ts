/**
 * Module 5: Alignment Recovery Engine
 * Recover row/column alignment using LCS and fuzzy matching
 */

import { WorkbookData } from "../types";

export interface AlignmentMapping {
  employeeRow: number;
  reviewerRow: number;
  confidence: number;
}

/**
 * Longest Common Subsequence calculation
 */
function calculateLCS(seq1: string[], seq2: string[]): number {
  const m = seq1.length;
  const n = seq2.length;
  const dp = Array(m + 1).fill(null).map(() => Array(n + 1).fill(0));

  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      if (seq1[i - 1] === seq2[j - 1]) {
        dp[i][j] = dp[i - 1][j - 1] + 1;
      } else {
        dp[i][j] = Math.max(dp[i - 1][j], dp[i][j - 1]);
      }
    }
  }
  return dp[m][n];
}

/**
 * Levenshtein distance for fuzzy matching
 */
export function levenshteinDistance(str1: string, str2: string): number {
  const m = str1.length;
  const n = str2.length;
  const dp = Array(m + 1).fill(null).map(() => Array(n + 1).fill(0));

  for (let i = 0; i <= m; i++) dp[i][0] = i;
  for (let j = 0; j <= n; j++) dp[0][j] = j;

  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      if (str1[i - 1] === str2[j - 1]) {
        dp[i][j] = dp[i - 1][j - 1];
      } else {
        dp[i][j] = 1 + Math.min(dp[i - 1][j], dp[i][j - 1], dp[i - 1][j - 1]);
      }
    }
  }
  return dp[m][n];
}

/**
 * Calculate similarity score (0-100%)
 */
export function calculateSimilarity(str1: string, str2: string): number {
  if (str1 === str2) return 100;
  const maxLen = Math.max(str1.length, str2.length);
  if (maxLen === 0) return 100;
  const distance = levenshteinDistance(str1, str2);
  return Math.round((1 - distance / maxLen) * 100);
}

/**
 * Recover row alignment using LCS and similarity matching
 */
export function recoverRowAlignment(
  employeeWb: WorkbookData,
  reviewerWb: WorkbookData,
  sheetName: string,
  confidenceThreshold: number = 0.9
): AlignmentMapping[] {
  const empSheet = employeeWb.sheets[sheetName];
  const revSheet = reviewerWb.sheets[sheetName];

  if (!empSheet || !revSheet) return [];

  const mappings: AlignmentMapping[] = [];

  // Extract row content for both sheets
  const empRowContents: string[][] = [];
  const revRowContents: string[][] = [];

  for (let r = 0; r <= empSheet.maxRow; r++) {
    const row: string[] = [];
    for (let c = 0; c <= empSheet.maxCol; c++) {
      const cell = empSheet.cells[`${r},${c}`];
      row.push(cell ? String(cell.raw) : "");
    }
    empRowContents.push(row);
  }

  for (let r = 0; r <= revSheet.maxRow; r++) {
    const row: string[] = [];
    for (let c = 0; c <= revSheet.maxCol; c++) {
      const cell = revSheet.cells[`${r},${c}`];
      row.push(cell ? String(cell.raw) : "");
    }
    revRowContents.push(row);
  }

  // Find best matches
  for (let empR = 0; empR < empRowContents.length; empR++) {
    let bestMatch = -1;
    let bestConfidence = 0;

    for (let revR = 0; revR < revRowContents.length; revR++) {
      const lcsLength = calculateLCS(empRowContents[empR], revRowContents[revR]);
      const maxLen = Math.max(empRowContents[empR].length, revRowContents[revR].length);
      const confidence = maxLen > 0 ? lcsLength / maxLen : 0;

      if (confidence > bestConfidence && confidence >= confidenceThreshold) {
        bestConfidence = confidence;
        bestMatch = revR;
      }
    }

    if (bestMatch >= 0) {
      mappings.push({
        employeeRow: empR,
        reviewerRow: bestMatch,
        confidence: bestConfidence
      });
    }
  }

  return mappings;
}
