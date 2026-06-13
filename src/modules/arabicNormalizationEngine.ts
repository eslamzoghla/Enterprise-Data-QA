/**
 * Module 9: Arabic Normalization Engine
 * Normalize Arabic text with Unicode mappings
 */

/**
 * Arabic normalization mapping
 */
const ARABIC_REPLACEMENTS: Record<string, string> = {
  "\u0623": "\u0627", // أ → ا
  "\u0625": "\u0627", // إ → ا
  "\u0622": "\u0627", // آ → ا
  "\u0649": "\u064A", // ى → ي
  "\u0624": "\u0648", // ؤ → و
  "\u0626": "\u064A"  // ئ → ي
};

/**
 * Arabic diacritics to remove
 */
const ARABIC_DIACRITICS = /[\u064B-\u0652]/g; // Fatha, Damma, Kasra, Sukun, Tanwin, Shadda, Maddah

/**
 * Tatweel character
 */
const TATWEEL = "\u0640";

/**
 * Normalize Arabic text
 */
export function normalizeArabicText(text: string): string {
  if (!text) return text;

  // Remove diacritics
  let normalized = text.replace(ARABIC_DIACRITICS, "");

  // Remove Tatweel
  normalized = normalized.replace(new RegExp(TATWEEL, "g"), "");

  // Apply replacements
  for (const [from, to] of Object.entries(ARABIC_REPLACEMENTS)) {
    normalized = normalized.replace(new RegExp(from, "g"), to);
  }

  // Preserve ة (Ta Marbuta) as-is

  return normalized;
}

/**
 * Check if text contains Arabic characters
 */
export function isArabicText(text: string): boolean {
  return /[\u0600-\u06FF]/.test(text);
}

/**
 * Normalize text with language detection
 */
export function normalizeWithLanguageDetection(text: string): string {
  if (isArabicText(text)) {
    return normalizeArabicText(text);
  }
  return text;
}
