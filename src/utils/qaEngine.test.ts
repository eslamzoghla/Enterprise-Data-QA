import { 
  normalizeArabicText, 
  convertArabicDigits, 
  normalizeNumericString, 
  parseDateDeterministic, 
  getNormalizedValue, 
  detectHeaderRows, 
  parseRangeOrSequence, 
  isEquivalentRangeRepresentational, 
  levenshteinDistance, 
  getSheetCellsMatrix,
  detectShifts 
} from "./qaEngine";
import { SheetGrid, CellValue, QAConfig } from "../types";

// Standard mock configuration
const testConfig: QAConfig = {
  numericMajorVarianceThreshold: 0.20,
  numericMajorAbsoluteThreshold: 5.0,
  arabicComparisonMode: "STANDARD",
  minimumShiftCells: 1,
  shiftDetectionThreshold: 0.80,
  employeeName: "Test worker",
  projectName: "Test Proj",
  evaluationDate: "2026-06-19",
  numericTolerance: 0.01,
  numericToleranceMode: "PERCENTAGE",
  shiftConfidenceScore: true,
  headerPenalty: 3,
  strictMode: "AUTO",
  extraTableCoefficient: 10,
  missingTableCoefficient: 10,
  extraColumnCoefficient: 5,
  missingColumnCoefficient: 5,
  extraRowCoefficient: 5,
  missingRowCoefficient: 5,
  numericDifferenceCoefficient: 2,
  textDifferenceCoefficient: 2,
  emptyCellDifferenceCoefficient: 1,
  autoIgnoreEnabled: false,
  customIgnorePatterns: ""
};

function runTestSuite() {
  console.log("=========================================");
  console.log("RUNNING QA ENGINE COMPLIANCE UNIT TESTS");
  console.log("=========================================");
  
  let passedCount = 0;
  let failedCount = 0;

  function assert(name: string, condition: boolean, message?: string) {
    if (condition) {
      console.log(`[PASS] ${name}`);
      passedCount++;
    } else {
      console.error(`[FAIL] ${name}${message ? " - " + message : ""}`);
      failedCount++;
    }
  }

  // Helper mock sheet builder
  function buildMockSheet(name: string, maxRow: number, maxCol: number, cellData: Record<string, any>): SheetGrid {
    const cells: Record<string, CellValue> = {};
    for (let r = 0; r <= maxRow; r++) {
      for (let c = 0; c <= maxCol; c++) {
        const val = cellData[`${r},${c}`];
        if (val !== undefined) {
          cells[`${r},${c}`] = getNormalizedValue(val, testConfig);
        } else {
          cells[`${r},${c}`] = { raw: null, formatted: "", normalized: "", type: "empty" };
        }
      }
    }
    return { name, maxRow, maxCol, cells };
  }

  // 1. Arabic Normalization checks (Rule 1: No Te Marbuta conversion)
  try {
    const testText = "مكتبة"; // contains ة
    const normalized = normalizeArabicText(testText);
    assert(
      "1. Arabic Normalization: Te Marbuta is NOT altered (ة -> ه is disabled)",
      !normalized.includes("مكتبه") && normalized.includes("ة")
    );
    
    const harakatText = "كِتَابٌ";
    const diacriticsRemoved = normalizeArabicText(harakatText);
    assert(
      "1b. Arabic Normalization: Diacritics are successfully removed",
      diacriticsRemoved === "كتاب"
    );
  } catch (e: any) {
    assert("1. Arabic Normalization", false, e.message);
  }

  // 2. Deterministic Date Parsing checks (Rule 2)
  try {
    const d1 = parseDateDeterministic("2026-06-19");
    const d2 = parseDateDeterministic("19/06/2026");
    const d3 = parseDateDeterministic("06/19/2026"); // MM/DD/YYYY where day > 12
    const dAny = parseDateDeterministic("١٩/٠٦/٢٠٢٦"); // Arabic numbers date

    assert("2. Date Parser: YYYY-MM-DD", d1.isDate && d1.formatted === "2026-06-19");
    assert("2b. Date Parser: DD/MM/YYYY", d2.isDate && d2.formatted === "2026-06-19");
    assert("2c. Date Parser: MM/DD/YYYY", d3.isDate && d3.formatted === "2026-06-19");
    assert("2d. Date Parser: Arabic digits support", dAny.isDate && dAny.formatted === "2026-06-19");
  } catch (e: any) {
    assert("2. Deterministic Date Parsing", false, e.message);
  }

  // 3. Time Type Handling (Rule 3)
  try {
    const normValue = getNormalizedValue("14:30:15", testConfig);
    const normValueArabic = getNormalizedValue("١٤:٣٠", testConfig);
    assert(
      "3. Time Classification: Treated as time rather than date",
      normValue.type === "time" && normValue.normalized === "14:30:15"
    );
    assert(
      "3b. Time Classification: Arabic digits parsed to time",
      normValueArabic.type === "time" && normValueArabic.normalized === "14:30"
    );
  } catch (e: any) {
    assert("3. Time Type Handling", false, e.message);
  }

  // 4. Column Shift Detection & Expanded Limits (Rule 4)
  try {
    const empDataColShift: Record<string, any> = {
      // Columns shifted right by 1
      "0,1": "ID", "0,2": "Name", "0,3": "Value",
      "1,1": "101", "1,2": "Alice", "1,3": "High",
      "2,1": "102", "2,2": "Bob", "2,3": "Low"
    };
    const revDataColShift: Record<string, any> = {
      // Standard layout
      "0,0": "ID", "0,1": "Name", "0,2": "Value",
      "1,0": "101", "1,1": "Alice", "1,2": "High",
      "2,0": "102", "2,1": "Bob", "2,2": "Low"
    };

    const empSheet = buildMockSheet("emp", 2, 4, empDataColShift);
    const revSheet = buildMockSheet("rev", 2, 4, revDataColShift);

    const shiftResult = detectShifts(empSheet, revSheet, testConfig);
    const colShiftEvent = shiftResult.events.find(ev => ev.type === "column" && ev.offset === -1);
    
    assert(
      "4. Column Shift Detection: Correctly identifies horizontal offsets and records shiftedCoords",
      colShiftEvent !== undefined && shiftResult.shiftedCoords["1,1"] === "column"
    );
  } catch (e: any) {
    assert("4. Column Shift Detection", false, e.message);
  }

  // 5. Row Shift Detection (Rule 5)
  try {
    const empDataRowShift: Record<string, any> = {
      // Row shifted down by 1 (Row 0 is empty/inserted, Row 1 starts data)
      "1,0": "A", "1,1": "B",
      "2,0": "C", "2,1": "D"
    };
    const revDataRowShift: Record<string, any> = {
      "0,0": "A", "0,1": "B",
      "1,0": "C", "1,1": "D"
    };

    const empSheet = buildMockSheet("empRow", 2, 1, empDataRowShift);
    const revSheet = buildMockSheet("revRow", 2, 1, revDataRowShift);

    const shiftResult = detectShifts(empSheet, revSheet, testConfig);
    const rowShiftEvent = shiftResult.events.find(ev => ev.type === "row" && ev.offset === -1);

    assert(
      "5. Row Shift Detection: Identifies vertical aligned block displacement",
      rowShiftEvent !== undefined && shiftResult.shiftedCoords["1,0"] === "row"
    );
  } catch (e: any) {
    assert("5. Row Shift Detection", false, e.message);
  }

  // 6. Score-Based Header Detection (Rule 6)
  try {
    // Sheet structure:
    // Row 0: Title row (has single cell)
    // Row 1: Blank row
    // Row 2: Actual Header row with distinct string headers
    // Row 3: Numeric Data rows
    const headerMockData: Record<string, any> = {
      "0,0": "MEDIC ALERTS TABLE CLINICAL DIRECTORY", // Title Cell
      "2,0": "Patient ID", "2,1": "Age Group", "2,2": "BP Category", // Headers
      "3,0": "101", "3,1": "45", "3,2": "120" // Numeric Data Row
    };
    const headerSheet = buildMockSheet("HeaderTest", 4, 2, headerMockData);
    const detectedHeaders = detectHeaderRows(headerSheet);

    assert(
      "6. Header Detection: Skips title & blank rows, correctly identifying Row 2 as actual header",
      detectedHeaders.has(2) && !detectedHeaders.has(0) && !detectedHeaders.has(1)
    );
  } catch (e: any) {
    assert("6. Header Detection", false, e.message);
  }

  // 7. Arabic Digit Support (Rule 7)
  try {
    const arabicNumText = "١٩٩٥.٥٠";
    const result = normalizeNumericString(arabicNumText);
    assert(
      "7. Arabic Numerals: Successfully normalized to Latin digits and numeric content validated",
      result.isNumeric && result.value === 1995.50 && result.cleanedStr === "1995.5"
    );
  } catch (e: any) {
    assert("7. Arabic Digit Support", false, e.message);
  }

  // 8. Range Parsing Improvements (Rule 8)
  try {
    const r1 = parseRangeOrSequence("١٩٩٠-١٩٩٥");
    const r2 = parseRangeOrSequence("أ1-أ5");

    assert(
      "8. Range Parser: Handles pure Arabic digit ranges",
      r1.isValid && r1.left === "1990" && r1.right === "1995" && r1.delimiter === "-"
    );
    assert(
      "8b. Range Parser: Handles hybrid Alphanumeric ranges like أ1-أ5",
      r2.isValid && r2.left === "ا1" && r2.right === "ا5" && r2.delimiter === "-"
    );
  } catch (e: any) {
    assert("8. Range Parsing", false, e.message);
  }

  // 9. Levenshtein Performance (Rule 9)
  try {
    const startTime = Date.now();
    const d1 = levenshteinDistance("very_long_mismatching_string_alpha_omega_12345", "entirely_different_and_unrelated_set_of_characters_here_67890");
    const duration = Date.now() - startTime;
    
    assert(
      "9. Levenshtein Performance: Highly dissimilar strings trigger early exit swiftly",
      duration <= 5
    );
  } catch (e: any) {
    assert("9. Levenshtein Performance", false, e.message);
  }

  // 10. Cell Access Performance (Rule 10)
  try {
    const sampleData: Record<string, any> = {
      "0,0": "A", "0,1": "B",
      "1,0": "C", "1,1": "D"
    };
    const sheet = buildMockSheet("LookupPerf", 1, 1, sampleData);
    const matrix = getSheetCellsMatrix(sheet);

    assert(
      "10. Cell Access performance: Correct 2D array representation built",
      matrix && matrix[0] && matrix[0][1]?.raw === "B" && matrix[1][0]?.raw === "C"
    );
  } catch (e: any) {
    assert("10. Cell Access Performance", false, e.message);
  }

  console.log("=========================================");
  console.log(`UNIT TESTING SUMMARY: ${passedCount} PASSED, ${failedCount} FAILED.`);
  console.log("=========================================");

  if (failedCount > 0) {
    process.exit(1);
  } else {
    process.exit(0);
  }
}

runTestSuite();
