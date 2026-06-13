# 🏗️ Architecture Documentation

## System Design Overview

The Enterprise QA Engine implements a **deterministic 17-module pipeline** for auditing Excel data quality.

### Core Principles

1. **Literal Accuracy > Semantic Accuracy**
   - `1965/64` ≠ `65/64` (different representations)
   - No business logic interpretation
   - Exact value matching (with configurable tolerance)

2. **Root Cause > Symptoms**
   - One structural event explains multiple cell mismatches
   - Suppress dependent errors to avoid double-counting
   - Report smallest number of errors explaining largest number of mismatches

3. **No Cell Comparison Before Structure Validation**
   - Validate sheets exist, rows/columns match
   - Recover alignment using LCS + fuzzy matching
   - THEN compare cell contents

---

## Module Dependency Graph

```
Modules 1-6: Data Loading & Structural Validation
    ↓
[GATE: Structure must be valid before cell comparison]
    ↓
Modules 8-11: Cell Comparison & Classification
    ↓
Module 7: Error Suppression (with root cause)
    ↓
Modules 12-13: Pattern Detection & Root Cause Analysis
    ↓
Modules 14-17: Scoring, Reporting, Metrics, Export
```

---

## Detailed Module Specifications

### Module 1: Workbook Loader
**Input:** File (Excel/CSV)  
**Output:** WorkbookData object

- Parses XLSX using `xlsx` library
- Expands merged cells within populated bounds
- Detects "tight bounds" to prevent 2000-row phantom grids
- Extracts raw, formatted, and type metadata for each cell

**Key Optimization:** Memory-efficient bounds detection prevents DOM lag

---

### Module 2: Worksheet Analyzer
**Input:** WorkbookData  
**Output:** WorksheetMetadata[]

- Analyzes each sheet independently
- Calculates density (populated cells %)
- Estimates header row position
- Validates sheet non-emptiness

---

### Module 3: Merged Cell Expander
**Input:** WorkbookData  
**Output:** Expanded WorkbookData

- Expands `!merges` array
- Copies starting cell value to all merged coordinates
- Respects tight bounds (max row/col) to prevent infinite loops

---

### Module 4: Structure Validator
**Input:** Employee WorkbookData, Reviewer WorkbookData  
**Output:** ValidationResult { isValid, errors, warnings }

- Compares sheet names
- Flags missing sheets (critical error)
- Flags extra sheets (warning)
- Validates shared sheets have content

---

### Module 5: Alignment Recovery Engine
**Input:** Two WorkbookData objects, sheet name  
**Output:** AlignmentMapping[]

**Algorithm:**
1. Extract row content sequences from both workbooks
2. Calculate LCS (Longest Common Subsequence) for each row pair
3. Score matches using `LCS_length / max_length`
4. Return mappings with confidence ≥ 90%

**Fallback:** Use fixed ±2 row offset if LCS fails

**Similarity Function:**
```typescript
similarity = 1 - (levenshteinDistance / maxLength)
```

---

### Module 6: Structural Error Detector
**Input:** Two validated WorkbookData objects  
**Output:** StructuralError[]

**Detects 12 types:**
- `MissingSheet`, `ExtraSheet`
- `MissingColumn`, `ExtraColumn` (count vs count)
- `MissingRow`, `ExtraRow` (count vs count)
- `TableMerge`, `TableSplit`, `RowShift`, `ColumnShift`, etc.

**Affected Cells Calculation:**
```
MissingColumns: (reviewerMaxCol - empMaxCol) × (empMaxRow + 1)
```

---

### Module 7: Structural Suppression Engine
**Input:** ErrorLog[], StructuralErrors[]  
**Output:** ErrorLog[] with suppression flags

**Suppression Rules (8):**

| Root Cause | Suppresses | Reason |
|-----------|-----------|--------|
| MissingSheet | All cell errors in sheet | Sheet doesn't exist |
| MissingColumn | Cell errors in column | Column doesn't exist |
| RowShift | Text/Numeric in affected range | Alignment explains mismatches |
| ColumnShift | Text/Numeric in affected range | Alignment explains mismatches |
| ExtraRow | Cell errors in row | Extra row mismatches are expected |
| TableMerge/Split | Dependent cell errors | Table structure changed |

**Algorithm:**
1. Find all structural errors
2. For each error, mark all related cell errors as suppressed
3. Add `suppressionReason` field to suppressed errors
4. Mark `suppressed: true` flag

---

### Module 8: Cell Comparison Engine
**Input:** CellValue, CellValue, config  
**Output:** CellComparisonResult

**Logic:**
1. Check exact match: `empVal === revVal` → match=true
2. If numeric, check tolerance:
   ```typescript
   diff = |empNum - revNum|
   if (mode === PERCENTAGE) {
     pct = (diff / max(empNum, revNum)) × 100
     match = pct ≤ tolerance
   } else { // ABSOLUTE
     match = diff ≤ tolerance
   }
   ```
3. If Arabic mode=STANDARD, normalize and retry
4. Calculate Levenshtein similarity %

---

### Module 9: Arabic Normalization Engine
**Input:** Text string  
**Output:** Normalized string

**Replacements (6):**
```typescript
أ → ا  (Alef with Hamza above)
إ → ا  (Alef with Hamza below)
آ → ا  (Alef with Maddah)
ى → ي  (Alef Maksura)
ؤ → و  (Waw with Hamza)
ئ → ي  (Yeh with Hamza)
```

**Removals:**
- All diacritics: َ ُ ِ ْ ً ٌ ٍ (Unicode U+064B-U+0652)
- Tatweel: ـ (U+0640)

**Preservation:**
- Ta Marbuta: ة (NEVER normalize)

---

### Module 10: Numeric Normalization Engine
**Input:** String or number  
**Output:** Normalized numeric value

**Transformations:**
1. Convert Arabic numerals (٠-٩) → Latin (0-9)
2. Remove thousand separators (`,`)
3. Remove leading zeros: `00125` → `125`
4. Normalize `-0` → `0`
5. Preserve decimal places: `100.0` → `100`

**Equivalence Check:**
```typescript
areNumericEquivalent(val1, val2, tolerance, mode) {
  num1 = parseNumericValue(val1)
  num2 = parseNumericValue(val2)
  if (num1 === null || num2 === null) return false
  
  diff = |num1 - num2|
  if (mode === ABSOLUTE)
    return diff ≤ tolerance
  else
    return (diff / max(|num1|, |num2|)) × 100 ≤ tolerance
}
```

---

### Module 11: Error Classifier
**Input:** EmployeeValue, ReviewerValue, ComparisonResult, isHeaderRow  
**Output:** ClassificationResult

**Decision Tree:**
```
1. If header row → HeaderMismatch (CRITICAL)
2. If one empty, one not → MissingValue or ExtraValue (HIGH)
3. If both numeric:
   a. |emp - rev| = 1 → DigitSubstitution (HIGH)
   b. Digits transposed → DigitTransposition (HIGH)
   c. Else → NumericDifference (HIGH)
4. If text:
   a. similarity > 85% → TextTypo (LOW)
   b. similarity > 50% → TextDifference (MEDIUM)
   c. Else → MajorTextDifference (CRITICAL)
```

---

### Module 12: Pattern Detection Engine
**Input:** ErrorLog[]  
**Output:** PatternAnalysis

**Detects 5 patterns:**

1. **Repeated Numeric Errors**
   - Group errors by `${sheet}:${empValue} vs ${revValue}`
   - Count occurrences > 1

2. **Repeated Digit Substitutions**
   - Specific digit pairs substituted multiple times

3. **Copy-Paste Errors**
   - Same employee value in 3+ cells with different reviewer values

4. **Error Clusters**
   - Divide sheet into 5×5 regions
   - Flag regions with >5 errors

5. **Sheet-Level Concentrations**
   - Calculate % of total errors per sheet
   - Flag sheets with disproportionate error rates

---

### Module 13: Root Cause Analyzer
**Input:** ErrorLog[] (non-suppressed only)  
**Output:** RootCauseAnalysis

**Aggregation:**
```typescript
for each error {
  if (notes.includes("[Header Row Error]")) headerErrors++
  if (errorType in [MissingValue, ExtraValue]) missingValues++
  if (errorType in [NumericDifference, ...]) numericErrors++
  if (errorType in [TextTypo, ...]) textErrors++
  if (errorType in [RangeInversion, ...]) rangeErrors++
  if (errorType in [RowShift, ...]) shiftErrors++
}

return {
  missingValuesPct: (missingValues / total) × 100,
  numericErrorsPct: (numericErrors / total) × 100,
  ...
}
```

---

### Module 14: Scoring Engine
**Input:** ErrorLog[], comparedCells, hasShiftEvents  
**Output:** ScoringResult

**Accuracy Calculation:**
```typescript
baseAccuracy = ((comparedCells - totalErrors) / comparedCells) × 100
```

**Grade Assignment:**
```
99.90%+ → Outstanding
99.00%+ → Excellent
97.00%+ → Very Good
95.00%+ → Good
90.00%+ → Fair
80.00%+ → Needs Improvement
<80%   → Poor
```

**Safety Rules:**
- Row/Column Shift detected → max grade "Needs Improvement"
- >10 Critical errors → max grade "Needs Improvement"
- >50 Major errors → max grade "Fair"

**Penalty Points:**
```typescript
totalPenalty = Σ(error.penalty × severityMultiplier)
// Critical=5x, High=2x, Medium=1x, Low=0.5x
```

**Workload Estimation:**
```typescript
reviewerHours = (nonSuppressedErrorCount × 5 min/error) / 60
```

**Error Rate:**
```typescript
errorRatePer10k = (totalErrors / comparedCells) × 10000
```

---

### Module 15: Reporting Engine
**Input:** Scoring, ErrorLog, RootCause, Patterns, Config  
**Output:** AuditReport

**6 Sections:**

1. **Executive Summary**
   - Employee, Project, Date
   - Grade, Accuracy %, Penalty Points
   - Error count, Reviewer workload

2. **Structural Defects Summary**
   - List missing/extra sheets/rows/columns
   - Table merges/splits
   - Counts per defect type

3. **Root Cause Analysis**
   - Error distribution table
   - Predominant issue identification
   - Recommendations for focus area

4. **Detailed Error Log**
   - First 50 errors (non-suppressed)
   - Columns: Sheet, Cell, Employee Value, Reviewer Value, Similarity %, Type, Severity, Penalty, Notes
   - Truncation disclaimer if >50 total

5. **Pattern Findings**
   - Repeated numeric errors (top 5)
   - Copy-paste anomalies (top 5)
   - Error clusters (top 3)
   - Shift events (top 3)

6. **Coaching Recommendations**
   - Conditional on error distribution
   - E.g., "Focus on data entry completeness" if missing values > 30%

---

### Module 16: Grid Inspector Engine
**Input:** ErrorLog[] (selected sheet)  
**Output:** GridMetrics

**8+ Metrics:**
1. Compared Cells
2. Total Errors
3. Accuracy %
4. Structural Errors
5. Shift Errors
6. Range Errors
7. Numeric Errors
8. Text Errors
9. Header Errors

---

### Module 17: Export Engine
**Input:** AuditReport, ErrorLog[], format  
**Output:** File content or download

**Formats:**
- **JSON:** Complete report object
- **CSV:** Error log table (with escaping)
- **HTML:** Formatted report with styles
- **PDF:** (Requires external library)

**CSV Escaping:**
```typescript
if (value contains "," or '"' or "\n")
  return `"${value.replace(/"/g, '""')}"`
else
  return value
```

---

## Error Suppression Example

**Scenario:**
```
Employee has EXTRA COLUMN D
Employee Column D cells: E1("X"), E2("Y"), E3("Z")
Reviewer has no column D (stops at C)
```

**Errors Before Suppression:**
```
E1: MissingValue (reviewer empty, emp "X")
E2: MissingValue (reviewer empty, emp "Y")
E3: MissingValue (reviewer empty, emp "Z")
+ 1 Structural: ExtraColumn
= 4 errors total
```

**After Suppression:**
```
1 Structural: ExtraColumn (ACTIVE)
3 Cell errors: SUPPRESSED (reason: "Suppressed by ExtraColumn")
= 1 error total (for metrics)
```

**Accuracy Impact:**
Before: 997/1000 = 99.7%  
After: 999/1000 = 99.9% ← Actual data quality!

---

## Performance Considerations

### Time Complexity
- **Module 1-6:** O(n) where n = total cells
- **Module 5:** O(r² × c) row alignment (r=rows, c=cols per row)
- **Module 8-11:** O(n) cell comparison
- **Module 12:** O(n log n) pattern grouping
- **Module 14:** O(n) scoring
- **Total:** O(n log n) dominated by pattern detection

### Space Complexity
- **ErrorLog:** O(e) where e = error count (typically << n)
- **Alignment mappings:** O(r²)
- **Total:** O(n) for workbook + O(e) for errors

### Optimization Tips
1. Use tight bounds detection (Module 1)
2. Batch similar errors for pattern detection
3. Apply suppression early to reduce error log size
4. Cache similarity calculations

---

## Testing Strategy

### Unit Tests (Per Module)
```typescript
// Module 9: Arabic Normalization
assert(normalizeArabicText("أحمد") === "احمد")
assert(normalizeArabicText("يَ") === "ي")
assert(normalizeArabicText("ة") === "ة") // NOT normalized

// Module 14: Scoring
assert(assignGrade(99.95) === "Outstanding")
assert(assignGrade(97.5, hasShift=true) === "Needs Improvement")
```

### Integration Tests
```typescript
// Full pipeline
const result = await executeQAEvaluation(empWb, revWb, config)
assert(result.metrics.baseAccuracy >= 0 && <= 100)
assert(result.errorLog.length <= result.metrics.totalErrors)
assert(result.coachingRecommendations.length > 0)
```

### Regression Tests
- Compare output across versions
- Validate error counts and grades remain consistent
- Benchmark performance improvements

---

## Future Enhancements

1. **Fuzzy table matching** for TableMerge/TableSplit detection
2. **Machine learning** for anomaly detection
3. **Multi-language support** (Spanish, French, etc.)
4. **Batch processing** for 1000s of audits
5. **Real-time dashboards** with WebSocket updates
6. **API for external integrations** (SAP, Oracle)
7. **PDF export** with charts and graphs

