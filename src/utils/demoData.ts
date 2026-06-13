/**
 * High-fidelity preloaded demo structures for Enterprise Excel QA testing.
 */

import { WorkbookData, CellValue } from "../types.ts";

export function getDemoEmployeeData(): WorkbookData {
  const cells: Record<string, CellValue> = {};

  // Define simple helper to set cell
  const setCell = (r: number, c: number, raw: any, formatted: string) => {
    cells[`${r},${c}`] = {
      raw,
      formatted,
      normalized: "",
      type: typeof raw === "number" ? "number" : "string"
    };
  };

  // Row 0: Merged Header in Reviewer, worker just submitted value on A1, leaving B1-D1 blank
  // Virtual expansion propagates A1 into B1-D1, resulting in perfect match!
  setCell(0, 0, "DEMOGRAPHIC AUDIT REPORT", "DEMOGRAPHIC AUDIT REPORT");
  // B1, C1, D1 left blank by employee

  // Row 1: Headers
  setCell(1, 0, "Employee ID", "Employee ID");
  setCell(1, 1, "Region / Location", "Region / Location");
  setCell(1, 2, "Age Bracket", "Age Bracket");
  setCell(1, 3, "Compliance Score", "Compliance Score");
  setCell(1, 4, "Arabic City Reference", "Arabic City Reference");

  // Row 2: Standard Match with Arabic normalizations (القاهرة with Heh 'ه' vs Reviewer with 'ة')
  setCell(2, 0, "EMP-001", "EMP-001");
  setCell(2, 1, "Middle East", "Middle East");
  setCell(2, 2, "20 سنة فأكثر", "20 سنة فأكثر");
  setCell(2, 3, 1000, "1,000.0"); // Numeric formatting equivalence
  setCell(2, 4, "القاهره", "القاهره"); // Normalizes to match "القاهرة"

  // Row 3: Arabic Alef-Hamza normalization equivalence
  setCell(3, 0, "EMP-002", "EMP-002");
  setCell(3, 1, "Middle East", "Middle East");
  setCell(3, 2, "أقل من 30", "أقل من 30");
  setCell(3, 3, 125, "125");
  setCell(3, 4, "الاسكندرية", "الاسكندرية"); // Normalizes to match "الإسكندرية"

  // Row 4: Minor text matching (diacritics removal equivalence)
  setCell(4, 0, "EMP-003", "EMP-003");
  setCell(4, 1, "North Region", "North Region");
  setCell(4, 2, "15-19 سنة", "15-19 سنة");
  setCell(4, 3, 350, "350");
  setCell(4, 4, "رِياض الجميل", "رِياض الجميل"); // With diacritics

  // Row 5: Ya/Te-marbuta Hamza variations (equivalence)
  setCell(5, 0, "EMP-004", "EMP-004");
  setCell(5, 1, "East District", "East District");
  setCell(5, 2, "75 فأكثر", "75 فأكثر");
  setCell(5, 3, 450, "450");
  setCell(5, 4, "جميل عبد الرءوف", "جميل عبد الرءوف"); // Hamza standard representation

  // Row 6: Spelling difference (Major Text Difference < 90%)
  setCell(6, 0, "EMP-005", "EMP-005");
  setCell(6, 1, "Levant", "Levant");
  setCell(6, 2, "20-24 سنة", "20-24 سنة");
  setCell(6, 3, 620, "620");
  setCell(6, 4, "سوريا", "سوريا"); // vs Reviewer "سورية" (Major text error)

  // Row 7: Dotless Ya / Ya-Hamza (equivalence)
  setCell(7, 0, "EMP-006", "EMP-006");
  setCell(7, 1, "West District", "West District");
  setCell(7, 2, "أنثي", "أنثي"); // Ends with dotless or standard 'ي' -> matches 'أنثى'
  setCell(7, 3, 720, "720");
  setCell(7, 4, "الجيزة", "الجيزة");

  // Row 8: Pure Numeric matching (equivalence check)
  setCell(8, 0, "EMP-007", "EMP-007");
  setCell(8, 1, "Central", "Central");
  setCell(8, 2, "50 فأكثر", "50 فأكثر");
  setCell(8, 3, 125, "125"); // 00125 in reviewer
  setCell(8, 4, "أسوان", "أسوان");

  // Row 9: Digit Substitution (Medium severity error)
  setCell(9, 0, "EMP-008", "EMP-008");
  setCell(9, 1, "South Region", "South Region");
  setCell(9, 2, "30-34 سنة", "30-34 سنة");
  setCell(9, 3, 2947, "2,947"); // vs Reviewer 2147
  setCell(9, 4, "الاقصر", "الاقصر");

  // Row 10: Missing Digit (Medium severity)
  setCell(10, 0, "EMP-009", "EMP-009");
  setCell(10, 1, "North", "North");
  setCell(10, 2, "40-44 سنة", "40-44 سنة");
  setCell(10, 3, 369, "369"); // vs Reviewer 9369
  setCell(10, 4, "الغردقة", "الغردقة");

  // Row 11: Extra Digit (Medium severity)
  setCell(11, 0, "EMP-010", "EMP-010");
  setCell(11, 1, "Marina", "Marina");
  setCell(11, 2, "45-49 سنة", "45-49 سنة");
  setCell(11, 3, 9369, "9,369"); // vs Reviewer 369
  setCell(11, 4, "شرم الشيخ", "شرم الشيخ");

  // Row 12: Digit Transposition (Medium severity)
  setCell(12, 0, "EMP-011", "EMP-011");
  setCell(12, 1, "Coastal", "Coastal");
  setCell(12, 2, "15-19 سنة", "15-19 سنة");
  setCell(12, 3, 1324, "1,324"); // vs Reviewer 1234
  setCell(12, 4, "دمياط", "دمياط");

  // Row 13: Major Numeric Error (High severity override)
  setCell(13, 0, "EMP-012", "EMP-012");
  setCell(13, 1, "Highlands", "Highlands");
  setCell(13, 2, "25-29 سنة", "25-29 سنة");
  setCell(13, 3, 500, "500"); // vs Reviewer 100 (Variance 400% > 20%)
  setCell(13, 4, "بورسعيد", "بورسعيد");

  // Row 14: Mixed Alphanumeric Extra Digit
  setCell(14, 0, "EMP-013", "EMP-013");
  setCell(14, 1, "Suburbs", "Suburbs");
  setCell(14, 2, "أقل من 50", "أقل من 50"); // vs Reviewer "أقل من 5" -> Extra digit '0'!
  setCell(14, 3, 600, "600");
  setCell(14, 4, "السويس", "السويس");

  // --- VERTICAL ROW SHIFT EVENT ---
  // Worker skips Row 15 completely, and inserts value, shifting subsequent rows up.
  // Reviewer Row 15: EMP-014, "Oasis", "35-39 سنة", 850, "الفيوم"
  // Employee Row 15 contains Reviewer Row 16 instead!
  // Reviewer Row 16: EMP-015, "Delta", "55-59 سنة", 950, "المنصورة"
  // Employee Row 16 contains Reviewer Row 17!
  // Reviewer Row 17: EMP-016, "Sina", "60-64 سنة", 120, "العريش"
  // Employee Row 17 contains Reviewer Row 18!
  // Reviewer Row 18: EMP-017, "Red Sea", "18-20 سنة", 130, "القصير"
  // This systematic row-mismatch shift continues for exactly 21 cells, satisfying shifted cells >= 20!

  // Row 15: Employee has EMP-015 (matches Reviewer 16!)
  setCell(15, 0, "EMP-015", "EMP-015");
  setCell(15, 1, "Delta", "Delta");
  setCell(15, 2, "55-59 سنة", "55-59 سنة");
  setCell(15, 3, 950, "950");
  setCell(15, 4, "المنصورة", "المنصورة");

  // Row 16: Employee has EMP-016 (matches Reviewer 17!)
  setCell(16, 0, "EMP-016", "EMP-016");
  setCell(16, 1, "Sina", "Sina");
  setCell(16, 2, "60-64 سنة", "60-64 سنة");
  setCell(16, 3, 120, "120");
  setCell(16, 4, "العريش", "العريش");

  // Row 17: Employee has EMP-017 (matches Reviewer 18!)
  setCell(17, 0, "EMP-017", "EMP-017");
  setCell(17, 1, "Red Sea", "Red Sea");
  setCell(17, 2, "18-20 سنة", "18-20 سنة");
  setCell(17, 3, 130, "130");
  setCell(17, 4, "القصير", "القصير");

  // Row 18: Employee has EMP-018 (matches Reviewer 19!)
  setCell(18, 0, "EMP-018", "EMP-018");
  setCell(18, 1, "Frontier", "Frontier");
  setCell(18, 2, "30-34 سنة", "30-34 سنة");
  setCell(18, 3, 190, "190");
  setCell(18, 4, "مطروح", "مطروح");

  // Row 19: Employee has EMP-019 (matches Reviewer 20!)
  setCell(19, 0, "EMP-019", "EMP-019");
  setCell(19, 1, "Canal", "Canal");
  setCell(19, 2, "20-22 سنة", "20-22 سنة");
  setCell(19, 3, 210, "210");
  setCell(19, 4, "الإسماعيلية", "الإسماعيلية");

  // Row 20: Employee has EMP-020 (matches Reviewer 21!)
  setCell(20, 0, "EMP-020", "EMP-020");
  setCell(20, 1, "Upper Egypt", "Upper Egypt");
  setCell(20, 2, "25-29 سنة", "25-29 سنة");
  setCell(20, 3, 310, "310");
  setCell(20, 4, "المنيا", "المنيا");

  // Row 21: Employee has EMP-021 (matches Reviewer 22!)
  setCell(21, 0, "EMP-021", "EMP-021");
  setCell(21, 1, "Nubia", "Nubia");
  setCell(21, 2, "35-39 سنة", "35-39 سنة");
  setCell(21, 3, 410, "410");
  setCell(21, 4, "نوبة", "نوبة");

  // Row 22: Back to alignment! Standard match
  setCell(22, 0, "EMP-022", "EMP-022");
  setCell(22, 1, "Gulf", "Gulf");
  setCell(22, 2, "40-44 سنة", "40-44 سنة");
  setCell(22, 3, 550, "550");
  setCell(22, 4, "طنطا", "طنطا");

  return {
    fileName: "Employee_DemographicSubmission.xlsx",
    sheets: {
      "Demographic Data": {
        name: "Demographic Data",
        maxRow: 22,
        maxCol: 4,
        cells
      }
    }
  };
}

export function getDemoReviewerData(): WorkbookData {
  const cells: Record<string, CellValue> = {};

  const setCell = (r: number, c: number, raw: any, formatted: string) => {
    cells[`${r},${c}`] = {
      raw,
      formatted,
      normalized: "",
      type: typeof raw === "number" ? "number" : "string"
    };
  };

  // Row 0: Merged cell range A1:D1 in ground truth Reviewer
  setCell(0, 0, "DEMOGRAPHIC AUDIT REPORT", "DEMOGRAPHIC AUDIT REPORT");
  setCell(0, 1, "DEMOGRAPHIC AUDIT REPORT", "DEMOGRAPHIC AUDIT REPORT");
  setCell(0, 2, "DEMOGRAPHIC AUDIT REPORT", "DEMOGRAPHIC AUDIT REPORT");
  setCell(0, 3, "DEMOGRAPHIC AUDIT REPORT", "DEMOGRAPHIC AUDIT REPORT");

  // Row 1: Headers
  setCell(1, 0, "Employee ID", "Employee ID");
  setCell(1, 1, "Region / Location", "Region / Location");
  setCell(1, 2, "Age Bracket", "Age Bracket");
  setCell(1, 3, "Compliance Score", "Compliance Score");
  setCell(1, 4, "Arabic City Reference", "Arabic City Reference");

  // Row 2: Match Cairo with Te Marbuta 'ة'
  setCell(2, 0, "EMP-001", "EMP-001");
  setCell(2, 1, "Middle East", "Middle East");
  setCell(2, 2, "20 سنة فأكثر", "20 سنة فأكثر");
  setCell(2, 3, 1000, "1000"); // 1000 vs 1,000.0 (equivalent)
  setCell(2, 4, "القاهرة", "القاهرة");

  // Row 3: Alef Hamza
  setCell(3, 0, "EMP-002", "EMP-002");
  setCell(3, 1, "Middle East", "Middle East");
  setCell(3, 2, "أقل من 30", "أقل من 30");
  setCell(3, 3, 125, "125");
  setCell(3, 4, "الإسكندرية", "الإسكندرية");

  // Row 4: Riyadh (stripped diacritics match)
  setCell(4, 0, "EMP-003", "EMP-003");
  setCell(4, 1, "North Region", "North Region");
  setCell(4, 2, "15-19 سنة", "15-19 سنة");
  setCell(4, 3, 350, "350");
  setCell(4, 4, "رياض الجميل", "رياض الجميل");

  // Row 5: Ya-Hamza equivalence
  setCell(5, 0, "EMP-004", "EMP-004");
  setCell(5, 1, "East District", "East District");
  setCell(5, 2, "75 فأكثر", "75 فأكثر");
  setCell(5, 3, 450, "450");
  setCell(5, 4, "جميل عبد الرؤوف", "جميل عبد الرؤوف");

  // Row 6: Spelling discrepancy
  setCell(6, 0, "EMP-005", "EMP-005");
  setCell(6, 1, "Levant", "Levant");
  setCell(6, 2, "20-24 سنة", "20-24 سنة");
  setCell(6, 3, 620, "620");
  setCell(6, 4, "سورية", "سورية"); // vs employee "سوريا"

  // Row 7: Dotless Ya equivalent match
  setCell(7, 0, "EMP-006", "EMP-006");
  setCell(7, 1, "West District", "West District");
  setCell(7, 2, "أنثى", "أنثى");
  setCell(7, 3, 720, "720");
  setCell(7, 4, "الجيزة", "الجيزة");

  // Row 8: Leading numeric zero sequence equal
  setCell(8, 0, "EMP-007", "EMP-007");
  setCell(8, 1, "Central", "Central");
  setCell(8, 2, "50 فأكثر", "50 فأكثر");
  setCell(8, 3, 125, "00125"); // 00125 vs 125 equivalent
  setCell(8, 4, "أسوان", "أسوان");

  // Row 9: Digit Substitution
  setCell(9, 0, "EMP-008", "EMP-008");
  setCell(9, 1, "South Region", "South Region");
  setCell(9, 2, "30-34 سنة", "30-34 سنة");
  setCell(9, 3, 2147, "2,147");
  setCell(9, 4, "الاقصر", "الاقصر");

  // Row 10: Missing Digit
  setCell(10, 0, "EMP-009", "EMP-009");
  setCell(10, 1, "North", "North");
  setCell(10, 2, "40-44 سنة", "40-44 سنة");
  setCell(10, 3, 9369, "9,369");
  setCell(10, 4, "الغردقة", "الغردقة");

  // Row 11: Extra Digit
  setCell(11, 0, "EMP-010", "EMP-010");
  setCell(11, 1, "Marina", "Marina");
  setCell(11, 2, "45-49 سنة", "45-49 سنة");
  setCell(11, 3, 369, "369");
  setCell(11, 4, "شرم الشيخ", "شرم الشيخ");

  // Row 12: Digit Transposition
  setCell(12, 0, "EMP-011", "EMP-011");
  setCell(12, 1, "Coastal", "Coastal");
  setCell(12, 2, "15-19 سنة", "15-19 سنة");
  setCell(12, 3, 1234, "1,234");
  setCell(12, 4, "دمياط", "دمياط");

  // Row 13: Major Numeric
  setCell(13, 0, "EMP-012", "EMP-012");
  setCell(13, 1, "Highlands", "Highlands");
  setCell(13, 2, "25-29 سنة", "25-29 سنة");
  setCell(13, 3, 100, "100");
  setCell(13, 4, "بورسعيد", "بورسعيد");

  // Row 14: Mixed Alphanumeric Extra digit
  setCell(14, 0, "EMP-013", "EMP-013");
  setCell(14, 1, "Suburbs", "Suburbs");
  setCell(14, 2, "أقل من 5", "أقل من 5");
  setCell(14, 3, 600, "600");
  setCell(14, 4, "السويس", "السويس");

  // --- REVIEWER EXPLICIT SECTIONS WHICH EMPLOYEE DISPLACED UPWARD ---
  // Row 15: Reviewer Row 15 - Worker omitted Row 15 completely in submission
  setCell(15, 0, "EMP-014", "EMP-014");
  setCell(15, 1, "Oasis", "Oasis");
  setCell(15, 2, "35-39 سنة", "35-39 سنة");
  setCell(15, 3, 850, "850");
  setCell(15, 4, "الفيوم", "الفيوم");

  // Row 16: Worker row 15 matches this
  setCell(16, 0, "EMP-015", "EMP-015");
  setCell(16, 1, "Delta", "Delta");
  setCell(16, 2, "55-59 سنة", "55-59 سنة");
  setCell(16, 3, 950, "950");
  setCell(16, 4, "المنصورة", "المنصورة");

  // Row 17: Worker row 16 matches this
  setCell(17, 0, "EMP-016", "EMP-016");
  setCell(17, 1, "Sina", "Sina");
  setCell(17, 2, "60-64 سنة", "60-64 سنة");
  setCell(17, 3, 120, "120");
  setCell(17, 4, "العريش", "العريش");

  // Row 18: Worker row 17 matches this
  setCell(18, 0, "EMP-017", "EMP-017");
  setCell(18, 1, "Red Sea", "Red Sea");
  setCell(18, 2, "18-20 سنة", "18-20 سنة");
  setCell(18, 3, 130, "130");
  setCell(18, 4, "القصير", "القصير");

  // Row 19: Worker row 18 matches this
  setCell(19, 0, "EMP-018", "EMP-018");
  setCell(19, 1, "Frontier", "Frontier");
  setCell(19, 2, "30-34 سنة", "30-34 سنة");
  setCell(19, 3, 190, "190");
  setCell(19, 4, "مطروح", "مطروح");

  // Row 20: Worker row 19 matches this
  setCell(20, 0, "EMP-019", "EMP-019");
  setCell(20, 1, "Canal", "Canal");
  setCell(20, 2, "20-22 سنة", "20-22 سنة");
  setCell(20, 3, 210, "210");
  setCell(20, 4, "الإسماعيلية", "الإسماعيلية");

  // Row 21: Worker row 20 matches this
  setCell(21, 0, "EMP-020", "EMP-020");
  setCell(21, 1, "Upper Egypt", "Upper Egypt");
  setCell(21, 2, "25-29 سنة", "25-29 سنة");
  setCell(21, 3, 310, "310");
  setCell(21, 4, "المنيا", "المنيا");

  // Row 22: Worker row 21 matches this
  setCell(22, 0, "EMP-021", "EMP-021");
  setCell(22, 1, "Nubia", "Nubia");
  setCell(22, 2, "35-39 سنة", "35-39 سنة");
  setCell(22, 3, 410, "410");
  setCell(22, 4, "نوبة", "نوبة");

  // Row 23: Reviewer row 23 (Employee has no row 23, results in aligned end of shift at row 22)
  setCell(23, 0, "EMP-022", "EMP-022");
  setCell(23, 1, "Gulf", "Gulf");
  setCell(23, 2, "40-44 سنة", "40-44 سنة");
  setCell(23, 3, 550, "550");
  setCell(23, 4, "طنطا", "طنطا");

  return {
    fileName: "GroundTruth_ReviewerDemographics.xlsx",
    sheets: {
      "Demographic Data": {
        name: "Demographic Data",
        maxRow: 23,
        maxCol: 4,
        cells
      }
    }
  };
}
