/**
 * Module 10: Numeric Normalization Engine
 * Normalize numeric values and handle representations
 */

/**
 * Convert Arabic numerals to Latin
 */
function arabicToLatin(text: string): string {
  const arabicNumerals = ["٠", "١", "٢", "٣", "٤", "٥", "٦", "٧", "٨", "٩"];
  const latinNumerals = ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9"];

  let result = text;
  for (let i = 0; i < arabicNumerals.length; i++) {
    result = result.replace(new RegExp(arabicNumerals[i], "g"), latinNumerals[i]);
  }
  return result;
}

/**
 * Normalize numeric string representation
 */
export function normalizeNumericString(text: string): string {
  // Convert Arabic numerals
  let normalized = arabicToLatin(text);

  // Remove thousand separators
  normalized = normalized.replace(/,/g, "");

  // Remove leading zeros (but preserve "0.x" format)
  if (!normalized.includes(".")) {
    normalized = String(parseInt(normalized) || 0);
  } else {
    const parts = normalized.split(".");
    parts[0] = String(parseInt(parts[0]) || 0);
    normalized = parts.join(".");
  }

  // Normalize negative zero
  if (normalized === "-0") {
    normalized = "0";
  }

  return normalized;
}

/**
 * Parse numeric value
 */
export function parseNumericValue(value: string | number): number | null {
  if (typeof value === "number") {
    return isNaN(value) ? null : value;
  }

  const normalized = normalizeNumericString(String(value));
  const parsed = parseFloat(normalized);
  return isNaN(parsed) ? null : parsed;
}

/**
 * Check if two numbers are equivalent within tolerance
 */
export function areNumericEquivalent(
  val1: string | number,
  val2: string | number,
  tolerance: number = 0.01,
  toleranceMode: "ABSOLUTE" | "PERCENTAGE" = "ABSOLUTE"
): boolean {
  const num1 = parseNumericValue(val1);
  const num2 = parseNumericValue(val2);

  if (num1 === null || num2 === null) {
    return false;
  }

  if (num1 === num2) return true;

  const diff = Math.abs(num1 - num2);

  if (toleranceMode === "ABSOLUTE") {
    return diff <= tolerance;
  } else {
    // PERCENTAGE
    const percentage = (diff / Math.max(Math.abs(num1), Math.abs(num2))) * 100;
    return percentage <= tolerance;
  }
}
