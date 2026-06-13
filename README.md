<div align="center">
<img width="1200" height="475" alt="Enterprise QA Banner" src="https://images.unsplash.com/photo-1552664730-d307ca884978?w=1200&h=475&fit=crop" />
</div>

# 🏢 Enterprise Excel QA Engine — Audit Edition

> **Production-grade data quality auditing system** for Excel/CSV workbooks. Detects 30+ error types using a 17-module pipeline with root cause suppression, Arabic normalization, and enterprise-grade reporting.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![TypeScript](https://img.shields.io/badge/TypeScript-4.9+-blue.svg)](https://www.typescriptlang.org/)
[![React](https://img.shields.io/badge/React-18+-blue.svg)](https://react.dev/)
![Status](https://img.shields.io/badge/Status-Production%20Ready-green.svg)

---

## 🎯 Core Features

### ✅ **Literal Accuracy Over Semantics**
- Exact value comparison (e.g., `1965/64` ≠ `65/64`)
- No business logic interpretation
- Deterministic shift precedence engine

### ✅ **30+ Error Classifications**
- **Structural:** Missing/Extra Sheets, Columns, Rows, Merges, Splits, Shifts
- **Digit-Level:** Substitution, Transposition, Missing/Extra digits
- **Numeric:** Differences, Major errors, Range inversions
- **Text:** Typos, Major differences
- **Headers:** Mismatch detection

### ✅ **Root Cause > Symptoms**
- Single structural event suppresses dependent cell errors
- Prevents double-counting mismatches
- Accuracy metrics reflect true data quality issues

### ✅ **Arabic Text Intelligence**
- Unicode normalization: ا, ي, و replacements
- Diacritic removal (تشكيل)
- Tatweel (ـ) stripping
- Ta Marbuta (ة) preservation

### ✅ **Enterprise Reporting**
- 6-section audit reports (Executive Summary, Structural Defects, Root Cause, Error Log, Patterns, Coaching)
- Multi-format export (JSON, CSV, HTML)
- Grid Inspector with 8+ metrics
- Reviewer workload estimation

### ✅ **Alignment Recovery**
- Longest Common Subsequence (LCS)
- Levenshtein distance fuzzy matching
- Intelligent row/column recovery

---

## 🏗️ Architecture: 17-Module Pipeline

```
┌─────────────────────────────────────────────────────────────┐
│ 1. WORKBOOK LOADER                                          │
│    Parse Excel/CSV with bounds detection                   │
└────────────────────┬────────────────────────────────────────┘
                     ↓
┌─────────────────────────────────────────────────────────────┐
│ 2. WORKSHEET ANALYZER                                       │
│    Extract structure & metadata                             │
└────────────────────┬────────────────────────────────────────┘
                     ↓
┌─────────────────────────────────────────────────────────────┐
│ 3. MERGED CELL EXPANDER                                     │
│    Expand within populated bounds                           │
└────────────────────┬────────────────────────────────────────┘
                     ↓
┌─────────────────────────────────────────────────────────────┐
│ 4. STRUCTURE VALIDATOR                                      │
│    Validate sheet integrity ⚠️ NO CELL COMPARISON YET      │
└────────────────────┬────────────────────────────────────────┘
                     ↓
┌─────────────────────────────────────────────────────────────┐
│ 5. ALIGNMENT RECOVERY ENGINE                                │
│    LCS + Fuzzy matching for row recovery                   │
└────────────────────┬────────────────────────────────────────┘
                     ↓
┌─────────────────────────────────────────────────────────────┐
│ 6. STRUCTURAL ERROR DETECTOR                                │
│    Detect 12 structural error types                        │
└────────────────────┬────────────────────────────────────────┘
                     ↓
┌─────────────────────────────────────────────────────────────┐
│ ✅ NOW: CELL COMPARISON ALLOWED (Structure validated)      │
└────────────────────┬────────────────────────────────────────┘
                     ↓
┌─────────────────────────────────────────────────────────────┐
│ 8. CELL COMPARISON ENGINE                                   │
│    Compare values with normalization                       │
└────────────────────┬────────────────────────────────────────┘
                     ↓
┌─────────────────────────────────────────────────────────────┐
│ 9. ARABIC NORMALIZATION ENGINE                              │
│    Unicode mappings + diacritic removal                    │
└────────────────────┬────────────────────────────────────────┘
                     ↓
┌─────────────────────────────────────────────────────────────┐
│ 10. NUMERIC NORMALIZATION ENGINE                            │
│     Tolerance modes, thousand separators, Arabic digits    │
└────────────────────┬────────────────────────────────────────┘
                     ↓
┌─────────────────────────────────────────────────────────────┐
│ 11. ERROR CLASSIFIER                                        │
│     Classify 30+ error types with severity                 │
└────────────────────┬────────────────────────────────────────┘
                     ↓
┌─────────────────────────────────────────────────────────────┐
│ 7. STRUCTURAL SUPPRESSION ENGINE                            │
│    Root cause suppresses dependent errors                  │
└────────────────────┬────────────────────────────────────────┘
                     ↓
┌─────────────────────────────────────────────────────────────┐
│ 12. PATTERN DETECTION ENGINE                                │
│     Detect repeated errors, clusters, shifts               │
└────────────────────┬────────────────────────────────────────┘
                     ↓
┌─────────────────────────────────────────────────────────────┐
│ 13. ROOT CAUSE ANALYZER                                     │
│     Aggregate error impact by type                         │
└────────────────────┬────────────────────────────────────────┘
                     ↓
┌─────────────────────────────────────────────────────────────┐
│ 14. SCORING ENGINE                                          │
│     Calculate grades with safety rules                     │
└────────────────────┬────────────────────────────────────────┘
                     ↓
┌─────────────────────────────────────────────────────────────┐
│ 15. REPORTING ENGINE                                        │
│     Generate 6-section audit reports                       │
└────────────────────┬────────────────────────────────────────┘
                     ↓
┌─────────────────────────────────────────────────────────────┐
│ 16. GRID INSPECTOR ENGINE                                   │
│     Display 8+ performance metrics                         │
└────────────────────┬────────────────────────────────────────┘
                     ↓
┌─────────────────────────────────────────────────────────────┐
│ 17. EXPORT ENGINE                                           │
│     JSON, CSV, HTML export formats                         │
└─────────────────────────────────────────────────────────────┘
```

---

## 🚀 Quick Start

### Installation

```bash
# Clone repository
git clone https://github.com/eslamzoghla/Enterprise-Data-QA.git
cd Enterprise-Data-QA

# Install dependencies
npm install

# Set up environment
cp .env.example .env.local
echo "GEMINI_API_KEY=your_key_here" >> .env.local
```

### Run Locally

```bash
# Development server
npm run dev

# Build for production
npm run build

# Start production server
npm start
```

Open [http://localhost:5173](http://localhost:5173) in your browser.

---

## 📊 Usage Example

### 1. Upload Files
- **File A:** Employee submission (Excel/CSV)
- **File B:** Reviewer ground truth (Excel/CSV)

### 2. Configure QA Parameters

```typescript
const config: QAConfig = {
  employeeName: "Ahmed Mansour",
  projectName: "Q2 Demographic Survey",
  evaluationDate: "2026-06-13",
  numericMajorVarianceThreshold: 0.20,      // 20%
  numericTolerance: 0.01,                   // 1%
  numericToleranceMode: "PERCENTAGE",
  arabicComparisonMode: "STANDARD",
  minimumShiftCells: 20,
  shiftDetectionThreshold: 0.80,            // 80%
  headerPenalty: 3,
  strictMode: "AUTO"
};
```

### 3. Execute QA Analysis

```typescript
import { executeQAEvaluation } from "./utils/qaEngineOrchestrator";

const result = await executeQAEvaluation(
  employeeWorkbook,
  reviewerWorkbook,
  config
);

console.log(result);
// {
//   metrics: { accuracy: 97.5, grade: "Very Good", ... },
//   errorLog: [ ... 50 errors ... ],
//   rootCause: { missingValues: 30%, numeric: 45%, ... },
//   patterns: { repeatedErrors: [...], shifts: [...] },
//   coachingRecommendations: [...],
//   report: { ... 6-section audit report ... }
// }
```

---

## 📋 Error Classification Reference

### Structural Errors (12)
| Error | Description |
|-------|-------------|
| `MissingSheet` | Sheet exists in Reviewer but not Employee |
| `ExtraSheet` | Sheet exists in Employee but not Reviewer |
| `MissingColumn` | Column(s) missing in Employee |
| `ExtraColumn` | Column(s) extra in Employee |
| `MissingRow` | Row(s) missing in Employee |
| `ExtraRow` | Row(s) extra in Employee |
| `TableMerge` | Separate tables merged in Employee |
| `TableSplit` | Single table split in Employee |
| `LocalRowMisalignment` | Local row offset detected |
| `LocalColumnMisalignment` | Local column offset detected |
| `RowShift` | Systematic row shift (>20 cells) |
| `ColumnShift` | Systematic column shift (>20 cells) |

### Cell-Level Errors (18+)
- **Numeric (6):** `MissingDigit`, `ExtraDigit`, `DigitSubstitution`, `DigitTransposition`, `NumericDifference`, `MajorNumericError`
- **Text (3):** `TextTypo`, `TextDifference`, `MajorTextDifference`
- **Range (3):** `RangeInversionError`, `RangeBoundaryError`, `RangeRepresentationError`
- **Values (2):** `MissingValue`, `ExtraValue`
- **Headers (1):** `HeaderMismatch`

---

## 🎨 Performance Grading

```
99.90%+        → Outstanding ⭐⭐⭐⭐⭐
99.00-99.89%   → Excellent   ⭐⭐⭐⭐
97.00-98.99%   → Very Good   ⭐⭐⭐
95.00-96.99%   → Good        ⭐⭐
90.00-94.99%   → Fair        ⭐
80.00-89.99%   → Needs Improvement 🔴
<80%           → Poor        ❌
```

**Safety Rules:**
- Row/Column Shift detected → Grade capped at "Needs Improvement"
- >10 Critical errors → Grade capped at "Needs Improvement"
- >50 Major errors → Grade capped at "Fair"

---

## 📊 Report Structure

### Executive Summary
- Employee name, project, evaluation date
- Performance metrics (Accuracy, Grade, Error count, Penalty points)
- Reviewer workload estimate

### Structural Defects Summary
- List of sheet/row/column mismatches
- Merge/split events
- Shift detections

### Root Cause Analysis
- Error distribution by category
- Predominant issue identification
- Actionable insights

### Detailed Error Log (First 50)
- Sheet, Cell, Employee Value, Reviewer Value
- Normalized representations
- Error type, Severity, Penalty points
- Root cause notes
- `[Truncated: X remaining errors]` if >50

### Pattern Findings
- Repeated numeric errors
- Copy-paste anomalies
- Error clusters by region
- Sheet-level concentrations

### Coaching Recommendations
- Data entry best practices
- Specific improvement areas
- Personalized training suggestions

---

## 🔧 Configuration Reference

```typescript
interface QAConfig {
  // Identity
  employeeName: string;              // "Ahmed Mansour"
  projectName: string;               // "Q2 Survey Ingress"
  evaluationDate: string;            // "2026-06-13"

  // Numeric Tolerance
  numericTolerance: number;          // 0.01 (1%)
  numericToleranceMode: string;      // "PERCENTAGE" | "ABSOLUTE"
  numericMajorVarianceThreshold: number; // 0.20 (20%)
  numericMajorAbsoluteThreshold: number; // 5.0

  // Text Processing
  arabicComparisonMode: string;      // "STANDARD" | "NONE"

  // Shift Detection
  minimumShiftCells: number;         // 20 (minimum cells for shift detection)
  shiftDetectionThreshold: number;   // 0.80 (80% pattern density)
  shiftConfidenceScore: boolean;     // true

  // Grading
  headerPenalty: number;             // 3 (penalty points per header error)
  strictMode: string;                // "AUTO" | "ON" | "OFF"
}
```

---

## 🌍 Arabic Support

### Automatic Normalization

| Original | Normalized | Note |
|----------|------------|------|
| `أحمد` | `احمد` | Alef variants unified |
| `يَ` | `ي` | Diacritics removed |
| `ـ` | (removed) | Tatweel stripped |
| `ة` | `ة` | Ta Marbuta preserved |
| `ئ`, `ؤ` | `ي`, `و` | Hamza variants normalized |

---

## 📈 Metrics & Analytics

### Grid Inspector Displays
1. **Compared Cells** - Total cells evaluated
2. **Total Errors** - Non-suppressed count
3. **Accuracy %** - Base accuracy score
4. **Structural Errors** - Sheet/Row/Column issues
5. **Shift Errors** - Alignment shift events
6. **Range Errors** - Range representation issues
7. **Numeric Errors** - Digit/numeric errors
8. **Text Errors** - Spelling/typo errors
9. **Header Errors** - Column label mismatches

### Workload Estimation
```
Reviewer Hours = (Error Count × 5 minutes per error) / 60
```

---

## 🔒 Data Privacy

- ✅ All spreadsheet processing occurs **locally in-browser** (client-side)
- ✅ No data leaves your device unless explicitly exported
- ✅ Optional: Upload summary reports to Google Sheets (your choice)
- ✅ Supports private deployment

---

## 🛠️ Technology Stack

- **Frontend:** React 18+ with TypeScript
- **UI Framework:** Tailwind CSS + Lucide Icons
- **Spreadsheet:** XLSX.js for Excel/CSV parsing
- **Animations:** Motion/Framer Motion
- **Backend (Optional):** Express.js + Gemini API
- **Deployment:** Vite + Node.js

---

## 📚 Module Documentation

Each module is independently testable:

```typescript
// Module 1: Workbook Loader
import { loadWorkbookFromFile } from "./modules/workbookLoader";
const wb = await loadWorkbookFromFile(file);

// Module 9: Arabic Normalization
import { normalizeArabicText } from "./modules/arabicNormalizationEngine";
const normalized = normalizeArabicText("أحمد"); // → "احمد"

// Module 14: Scoring
import { score } from "./modules/scoringEngine";
const result = score(errorLog, 1000);
```

---

## 🤝 Contributing

Contributions welcome! Please:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit changes (`git commit -m 'Add amazing feature'`)
4. Push to branch (`git push origin feature/amazing-feature`)
5. Open Pull Request

---

## 📄 License

MIT License - see [LICENSE](LICENSE) file for details

---

## 🙋 Support

- **Issues:** [GitHub Issues](https://github.com/eslamzoghla/Enterprise-Data-QA/issues)
- **Discussions:** [GitHub Discussions](https://github.com/eslamzoghla/Enterprise-Data-QA/discussions)
- **Email:** contact@example.com

---

## 📞 Contact

**Eslam Zoghla** - [@eslamzoghla](https://github.com/eslamzoghla)

Project Link: [https://github.com/eslamzoghla/Enterprise-Data-QA](https://github.com/eslamzoghla/Enterprise-Data-QA)

---

## 🙏 Acknowledgments

- Excel.js for spreadsheet parsing
- Inspired by enterprise QA standards
- Arabic NLP best practices
- Open-source community

---

<div align="center">

**[⬆ back to top](#-enterprise-excel-qa-engine--audit-edition)**

Built with ❤️ for Data Quality Excellence

</div>
