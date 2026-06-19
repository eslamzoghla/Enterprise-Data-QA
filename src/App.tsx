import React, { useState, useMemo, useEffect } from "react";
import * as XLSX from "xlsx";
import { motion, AnimatePresence } from "motion/react";
import { jsPDF } from "jspdf";
import html2canvas from "html2canvas";

if (typeof window !== "undefined") {
  (window as any).html2canvas = html2canvas;
}
import {
  Upload,
  CheckCircle,
  AlertTriangle,
  FileSpreadsheet,
  Play,
  SlidersHorizontal,
  Calendar,
  AlertCircle,
  Sparkles,
  Layers,
  Table,
  Filter,
  ArrowUpDown,
  Download,
  BookOpen,
  User,
  Activity,
  ChevronRight,
  RefreshCw,
  Printer,
  Plus,
  ExternalLink,
  Award,
  TrendingUp,
  TrendingDown,
  Copy,
  FileText,
  Check,
  FileJson,
  ShieldAlert
} from "lucide-react";

declare global {
  interface Window {
    google?: any;
  }
}

import {
  ErrorType,
  Severity,
  QAConfig,
  CellValue,
  WorkbookData,
  AnalysisResult,
  ErrorLogEntry,
  ShiftEvent
} from "./types.ts";

import {
  getNormalizedValue,
  executeQAEvaluation,
  getColLetter
} from "./utils/qaEngine.ts";

import {
  getDemoEmployeeData,
  getDemoReviewerData
} from "./utils/demoData.ts";

// Helper to normalize the raw workbook data with current configurations on the fly
function normalizeWorkbook(wb: WorkbookData, config: QAConfig): WorkbookData {
  const normalizedSheets: Record<string, any> = {};
  for (const [sheetName, sheetGrid] of Object.entries(wb.sheets)) {
    const normCells: Record<string, CellValue> = {};
    for (const [coord, cell] of Object.entries(sheetGrid.cells)) {
      normCells[coord] = getNormalizedValue(cell.raw, config);
    }
    normalizedSheets[sheetName] = {
      ...sheetGrid,
      cells: normCells
    };
  }
  return {
    ...wb,
    sheets: normalizedSheets
  };
}

const isErrorOfType = (err: any, filter: string): boolean => {
  if (filter === "ALL") return true;
  const isHeader = err.notes.includes("[Header Row Error]");
  if (isHeader) return filter === "HEADER";
  if (filter === "HEADER") return false;

  switch (err.errorType) {
    case ErrorType.MissingDigit:
    case ErrorType.ExtraDigit:
    case ErrorType.DigitSubstitution:
    case ErrorType.DigitTransposition:
    case ErrorType.NumericDifference:
    case ErrorType.MajorNumericError:
      return filter === "NUMERIC";
    case ErrorType.MissingValue:
    case ErrorType.ExtraValue:
    case ErrorType.MissingRow:
    case ErrorType.ExtraRow:
    case ErrorType.MissingColumn:
    case ErrorType.ExtraColumn:
    case ErrorType.MissingCell:
    case ErrorType.ExtraCell:
    case ErrorType.TableMerge:
    case ErrorType.TableSplit:
      return filter === "STRUCTURAL";
    case ErrorType.RowShift:
    case ErrorType.ColumnShift:
    case ErrorType.LocalColumnMisalignment:
    case ErrorType.LocalRowMisalignment:
      return filter === "SHIFT";
    case ErrorType.TextTypo:
    case ErrorType.MajorTextDifference:
    case ErrorType.TextDifference:
      return filter === "TEXT";
    case ErrorType.RangeInversionError:
    case ErrorType.RangeBoundaryError:
    case ErrorType.RangeRepresentationError:
      return filter === "RANGE";
  }
  return false;
};

export default function App() {
  // Config state
  const [config, setConfig] = useState<QAConfig>({
    numericMajorVarianceThreshold: 0.20, // 20%
    numericMajorAbsoluteThreshold: 5.0,
    arabicComparisonMode: "STANDARD",
    minimumShiftCells: 20,
    shiftDetectionThreshold: 0.80, // 80%
    employeeName: "Employee Name",
    projectName: "Q2 Demographic Survey Ingress",
    evaluationDate: new Date().toISOString().split("T")[0],
    numericTolerance: 0.01,
    numericToleranceMode: "PERCENTAGE",
    shiftConfidenceScore: true,
    headerPenalty: 3,
    strictMode: "AUTO",
    extraTableCoefficient: 50,
    missingTableCoefficient: 100,
    extraColumnCoefficient: 5,
    missingColumnCoefficient: 10,
    extraRowCoefficient: 1,
    missingRowCoefficient: 2,
    numericDifferenceCoefficient: 0.1,
    textDifferenceCoefficient: 0.1,
    emptyCellDifferenceCoefficient: 0.05,
    autoIgnoreEnabled: true,
    customIgnorePatterns: "Cover, Index, Notes, Metadata, Instructions"
  });

  // Upload/Data state
  const [employeeWb, setEmployeeWb] = useState<WorkbookData | null>(null);
  const [reviewerWb, setReviewerWb] = useState<WorkbookData | null>(null);
  const [dragActiveE, setDragActiveE] = useState(false);
  const [dragActiveR, setDragActiveR] = useState(false);
  const [errorMsg, setErrorMsg] = useState<string | null>(null);
  const [isComparingStarted, setIsComparingStarted] = useState(false);

  // Layout navigation state
  const [activeTab, setActiveTab] = useState<"dashboard" | "errorLog" | "sheetExplorer" | "patterns" | "aiAuditor" | "mappingExplorer" | "benchmark">("dashboard");
  const [selectedSheetExplorer, setSelectedSheetExplorer] = useState<string>("");
  const [explorerViewMode, setExplorerViewMode] = useState<"sideBySide" | "grid">("sideBySide");
  const [activeErrorFilter, setActiveErrorFilter] = useState<"ALL" | "STRUCTURAL" | "SHIFT" | "RANGE" | "NUMERIC" | "TEXT" | "HEADER">("ALL");
  const [forceShowAllErrors, setForceShowAllErrors] = useState(false);

  // New Heatmap Overlay and Theme customization
  const [showHeatmap, setShowHeatmap] = useState<boolean>(false);
  const [heatmapTheme, setHeatmapTheme] = useState<"coral" | "ruby" | "emerald">("coral");

  // Benchmark and Regression Testing Baseline Run
  const [benchmarkBaseline, setBenchmarkBaseline] = useState<any | null>(null);
  const [baselineFileName, setBaselineFileName] = useState<string>("");
  const [baselineUploadError, setBaselineUploadError] = useState<string | null>(null);
  
  // Custom Print error handles for sandboxed browser iframes
  const [showPrintModal, setShowPrintModal] = useState<boolean>(false);
  const [printError, setPrintError] = useState<string | null>(null);

  // Sorting and Filtering for detailed error logs
  const [logFilterSeverity, setLogFilterSeverity] = useState<string>("ALL");
  const [logFilterType, setLogFilterType] = useState<string>("ALL");
  const [logSearch, setLogSearch] = useState<string>("");
  const [logSortField, setLogSortField] = useState<"severity" | "penalty" | "sheet" | "cell">("severity");

  // AI Auditor backend analysis state
  const [aiResult, setAiResult] = useState<{ summary: string; coaching: string[] } | null>(null);
  const [aiLoading, setAiLoading] = useState(false);

  // Google Sheets integration state
  const [googleClientId, setGoogleClientId] = useState<string>("");
  const [envKeys, setEnvKeys] = useState<string[]>([]);
  const [gsiLoaded, setGsiLoaded] = useState(false);
  const [googleAccessToken, setGoogleAccessToken] = useState<string | null>(() => {
    return localStorage.getItem("g_sheets_token") || null;
  });
  const [sheetsUser, setSheetsUser] = useState<{ name?: string; email?: string } | null>(null);
  const [spreadsheetId, setSpreadsheetId] = useState<string>(() => {
    return localStorage.getItem("g_sheets_spreadsheet_id") || "";
  });
  const [sheetsLoading, setSheetsLoading] = useState(false);
  const [sheetsMessage, setSheetsMessage] = useState<{ text: string; type: "success" | "error" | "info" } | null>(null);

  const isEmployeeNameValid = useMemo(() => {
    if (!config.employeeName) return false;
    const nameStr = config.employeeName.trim().toLowerCase();
    return nameStr !== "" && nameStr !== "employee name";
  }, [config.employeeName]);

  // Excel Parser
  const parseExcelFile = async (file: File): Promise<WorkbookData> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = e.target?.result;
          if (!data) throw new Error("File empty or unreadable");
          const workbook = XLSX.read(data, { type: "binary" });
          
          const wbData: WorkbookData = {
            fileName: file.name,
            sheets: {}
          };

          workbook.SheetNames.forEach((sheetName) => {
            const sheet = workbook.Sheets[sheetName];
            
            // Helper to check if a cell is truly populated (raw value or formatted text)
            const isCellPopulated = (c: any) => {
              if (!c) return false;
              if (c.v !== undefined && c.v !== null && c.v !== "") return true;
              if (c.w !== undefined && c.w !== null && c.w.trim() !== "") return true;
              return false;
            };

            // 1. Detect tight bounds of actual populated cells first to prevent extreme lag on phantom grids
            let tightMaxRow = 0;
            let tightMaxCol = 0;
            let hasAnyPopulatedCell = false;

            for (const key in sheet) {
              if (key[0] === "!") continue; // skip '!ref', '!merges', '!margins', etc.
              const parsed = XLSX.utils.decode_cell(key);
              const cell = sheet[key];
              if (isCellPopulated(cell)) {
                hasAnyPopulatedCell = true;
                if (parsed.r > tightMaxRow) tightMaxRow = parsed.r;
                if (parsed.c > tightMaxCol) tightMaxCol = parsed.c;
              }
            }

            if (!hasAnyPopulatedCell) {
              tightMaxRow = 0;
              tightMaxCol = 0;
            }

            // 2. Expand merged cells virtually (Section 4) - ONLY within the tight populated bounds to avoid infinite loop / memory freeze on full-column/row merges.
            if (sheet["!merges"]) {
              sheet["!merges"].forEach((merge) => {
                const startCellRef = XLSX.utils.encode_cell({ r: merge.s.r, c: merge.s.c });
                const startCellValue = sheet[startCellRef];
                if (startCellValue) {
                  const endRow = Math.min(merge.e.r, tightMaxRow);
                  const endCol = Math.min(merge.e.c, tightMaxCol);

                  for (let r = merge.s.r; r <= endRow; r++) {
                    for (let c = merge.s.c; c <= endCol; c++) {
                      const currentRef = XLSX.utils.encode_cell({ r, c });
                      if (currentRef !== startCellRef) {
                        sheet[currentRef] = { ...startCellValue };
                      }
                    }
                  }
                }
              });
            }

            const cells: Record<string, CellValue> = {};

            for (let r = 0; r <= tightMaxRow; r++) {
              for (let c = 0; c <= tightMaxCol; c++) {
                const cellRef = XLSX.utils.encode_cell({ r, c });
                const cell = sheet[cellRef];
                
                if (isCellPopulated(cell)) {
                  const rawVal = cell.v !== undefined && cell.v !== null && cell.v !== "" ? cell.v : (cell.w || "");
                  cells[`${r},${c}`] = {
                    raw: rawVal,
                    formatted: cell.w || String(rawVal),
                    normalized: String(rawVal),
                    type: typeof rawVal === "number" ? "number" : "string"
                  };
                }
              }
            }

            wbData.sheets[sheetName] = {
              name: sheetName,
              maxRow: tightMaxRow,
              maxCol: tightMaxCol,
              cells
            };
          });

          resolve(wbData);
        } catch (err: any) {
          reject(err);
        }
      };
      reader.onerror = () => reject(new Error("File read error"));
      reader.readAsBinaryString(file);
    });
  };

  const handleDrag = (e: React.DragEvent, section: "emp" | "rev") => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === "dragenter" || e.type === "dragover") {
      if (section === "emp") setDragActiveE(true);
      else setDragActiveR(true);
    } else if (e.type === "dragleave") {
      if (section === "emp") setDragActiveE(false);
      else setDragActiveR(false);
    }
  };

  const handleDrop = async (e: React.DragEvent, section: "emp" | "rev") => {
    e.preventDefault();
    e.stopPropagation();
    if (section === "emp") setDragActiveE(false);
    else setDragActiveR(false);

    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      const file = e.dataTransfer.files[0];
      try {
        const parsed = await parseExcelFile(file);
        if (section === "emp") setEmployeeWb(parsed);
        else setReviewerWb(parsed);
        setErrorMsg(null);
        setIsComparingStarted(false);
      } catch (err: any) {
        setErrorMsg(`Failed parsing '${file.name}': ${err.message}`);
      }
    }
  };

  const handleFileSelect = async (e: React.ChangeEvent<HTMLInputElement>, section: "emp" | "rev") => {
    if (e.target.files && e.target.files[0]) {
      const file = e.target.files[0];
      try {
        const parsed = await parseExcelFile(file);
        if (section === "emp") setEmployeeWb(parsed);
        else setReviewerWb(parsed);
        setErrorMsg(null);
        setIsComparingStarted(false);
      } catch (err: any) {
        setErrorMsg(`Failed parsing '${file.name}': ${err.message}`);
      }
    }
  };

  // Re-run evaluation dynamically whenever config or workbooks change
  const qaAnalysis = useMemo<AnalysisResult | null>(() => {
    if (!employeeWb || !reviewerWb || !isComparingStarted) return null;
    
    // Normalize workbooks with current dynamic configurations
    const normEmployee = normalizeWorkbook(employeeWb, config);
    const normReviewer = normalizeWorkbook(reviewerWb, config);
    
    return executeQAEvaluation(normEmployee, normReviewer, config);
  }, [
    employeeWb,
    reviewerWb,
    isComparingStarted,
    config.numericMajorVarianceThreshold,
    config.numericMajorAbsoluteThreshold,
    config.arabicComparisonMode,
    config.minimumShiftCells,
    config.shiftDetectionThreshold,
    config.numericTolerance,
    config.numericToleranceMode,
    config.shiftConfidenceScore,
    config.headerPenalty,
    config.strictMode,
    config.extraTableCoefficient,
    config.missingTableCoefficient,
    config.extraColumnCoefficient,
    config.missingColumnCoefficient,
    config.extraRowCoefficient,
    config.missingRowCoefficient,
    config.numericDifferenceCoefficient,
    config.textDifferenceCoefficient,
    config.emptyCellDifferenceCoefficient,
    config.autoIgnoreEnabled,
    config.customIgnorePatterns
  ]);

  // Fast O(1) error rendering lookup map for currently selected sheet
  const sheetErrorMap = useMemo(() => {
    const map = new Map<string, any>();
    if (qaAnalysis && selectedSheetExplorer) {
      qaAnalysis.errorLog.forEach((err) => {
        if (err.sheet === selectedSheetExplorer) {
          map.set(`${err.rowIndex},${err.colIndex}`, err);
        }
      });
    }
    return map;
  }, [qaAnalysis, selectedSheetExplorer]);

  const computeSheetErrorsBreakdown = (sheetName: string) => {
    const sheetErrors = qaAnalysis?.errorLog?.filter(x => x.sheet === sheetName) || [];
    
    let structural = 0;
    let shift = 0;
    let range = 0;
    let numeric = 0;
    let text = 0;
    let header = 0;

    for (const err of sheetErrors) {
      const isHeaderErr = err.notes.includes("[Header Row Error]");
      if (isHeaderErr) {
        header++;
        continue;
      }

      switch (err.errorType) {
        case ErrorType.MissingDigit:
        case ErrorType.ExtraDigit:
        case ErrorType.DigitSubstitution:
        case ErrorType.DigitTransposition:
        case ErrorType.NumericDifference:
        case ErrorType.MajorNumericError:
          numeric++;
          break;
        case ErrorType.MissingValue:
        case ErrorType.ExtraValue:
        case ErrorType.MissingRow:
        case ErrorType.ExtraRow:
        case ErrorType.MissingColumn:
        case ErrorType.ExtraColumn:
        case ErrorType.MissingCell:
        case ErrorType.ExtraCell:
        case ErrorType.TableMerge:
        case ErrorType.TableSplit:
          structural++;
          break;
        case ErrorType.RowShift:
        case ErrorType.ColumnShift:
        case ErrorType.LocalColumnMisalignment:
        case ErrorType.LocalRowMisalignment:
          shift++;
          break;
        case ErrorType.TextTypo:
        case ErrorType.MajorTextDifference:
        case ErrorType.TextDifference:
          text++;
          break;
        case ErrorType.RangeInversionError:
        case ErrorType.RangeBoundaryError:
        case ErrorType.RangeRepresentationError:
          range++;
          break;
      }
    }

    return {
      total: sheetErrors.length,
      structural,
      shift,
      range,
      numeric,
      text,
      header
    };
  };

  // Calculate bounds of populated cells of selected sheet to prevent rendering massive empty space (e.g., 2000 empty rows)
  const visualGridBounds = useMemo(() => {
    if (!selectedSheetExplorer) return { maxRow: 0, maxCol: 0, realMaxRow: 0, realMaxCol: 0 };
    const empSheet = employeeWb?.sheets[selectedSheetExplorer];
    const revSheet = reviewerWb?.sheets[selectedSheetExplorer];
    
    let maxR = 0;
    let maxC = 0;

    if (empSheet) {
      Object.keys(empSheet.cells).forEach(coord => {
        const [r, c] = coord.split(",").map(Number);
        if (r > maxR) maxR = r;
        if (c > maxC) maxC = c;
      });
    }
    if (revSheet) {
      Object.keys(revSheet.cells).forEach(coord => {
        const [r, c] = coord.split(",").map(Number);
        if (r > maxR) maxR = r;
        if (c > maxC) maxC = c;
      });
    }

    // Give a buffer of a few cells for context
    return {
      maxRow: Math.min(Math.max(15, maxR + 2), 150), // Cap visual grid row rendering at 150 to keep DOM fast
      maxCol: Math.min(Math.max(8, maxC + 1), 35),   // Cap visual grid col rendering at 35 to keep layout clean
      realMaxRow: Math.max(empSheet?.maxRow || 0, revSheet?.maxRow || 0),
      realMaxCol: Math.max(empSheet?.maxCol || 0, revSheet?.maxCol || 0)
    };
  }, [employeeWb, reviewerWb, selectedSheetExplorer]);

  // Set default sheet in visual explorer when sheet list populates
  useEffect(() => {
    if (qaAnalysis && !selectedSheetExplorer) {
      const sheets = Object.keys(employeeWb?.sheets || {});
      if (sheets.length > 0) {
        setSelectedSheetExplorer(sheets[0]);
      }
    }
  }, [qaAnalysis]);

  // Reset AI states when submission files change
  useEffect(() => {
    setAiResult(null);
  }, [employeeWb, reviewerWb]);

  // Fetch Google Client ID and other backend configs on mount
  useEffect(() => {
    fetch("/api/config")
      .then((res) => res.json())
      .then((data) => {
        if (data.googleClientId) {
          setGoogleClientId(data.googleClientId);
        }
        if (data.envKeys) {
          setEnvKeys(data.envKeys);
        }
      })
      .catch((err) => console.error("Config fetch error:", err));
  }, []);

  // Dynamically load Google Identity Services GSI script for OAuth client-side implicit flow 
  useEffect(() => {
    const script = document.createElement("script");
    script.src = "https://accounts.google.com/gsi/client";
    script.async = true;
    script.defer = true;
    script.onload = () => setGsiLoaded(true);
    document.body.appendChild(script);
    return () => {
      // Clean up script
      try {
        document.body.removeChild(script);
      } catch (e) {}
    };
  }, []);

  // Load profile details whenever the Access Token updates
  useEffect(() => {
    if (googleAccessToken) {
      fetchGoogleUserProfile(googleAccessToken);
    }
  }, [googleAccessToken]);

  const fetchGoogleUserProfile = async (token: string) => {
    try {
      const res = await fetch("https://www.googleapis.com/oauth2/v3/userinfo", {
        headers: { Authorization: `Bearer ${token}` }
      });
      if (res.ok) {
        const data = await res.json();
        setSheetsUser({
          name: data.name,
          email: data.email
        });
      } else {
        // Token has likely expired, clear it
        setGoogleAccessToken(null);
        localStorage.removeItem("g_sheets_token");
        setSheetsUser(null);
      }
    } catch (err) {
      console.error("Error loading Google account details:", err);
    }
  };

  const handleGoogleSignIn = () => {
    if (!window.google) {
      alert("Google Identity Client script is loading. Please reload or wait a few seconds.");
      return;
    }

    // Try finding Client ID from environment variables, or fallback to custom saved client ID, or ask the user
    const activeClientId = googleClientId.trim() || localStorage.getItem("g_custom_client_id") || "";

    if (!activeClientId) {
      const inputId = window.prompt(
        "To save to Google Sheets, enter your Google OAuth 2.0 Client ID:\n(You can create one in console.cloud.google.com with origin " + window.location.origin + ")"
      );
      if (!inputId) return;
      localStorage.setItem("g_custom_client_id", inputId.trim());
      setGoogleClientId(inputId.trim());
      triggerGsiAuth(inputId.trim());
    } else {
      triggerGsiAuth(activeClientId);
    }
  };

  const triggerGsiAuth = (cid: string) => {
    try {
      const tokenClient = window.google.accounts.oauth2.initTokenClient({
        client_id: cid,
        scope: "https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/userinfo.profile https://www.googleapis.com/auth/userinfo.email",
        callback: (resp: any) => {
          if (resp.access_token) {
            setGoogleAccessToken(resp.access_token);
            localStorage.setItem("g_sheets_token", resp.access_token);
            fetchGoogleUserProfile(resp.access_token);
          }
        },
      });
      tokenClient.requestAccessToken();
    } catch (err) {
      console.error("GSI token client init error:", err);
      alert("Failed to initialize Google Sign-In. Verify that your Client ID is correct.");
    }
  };

  const handleGoogleSignOut = () => {
    setGoogleAccessToken(null);
    setSheetsUser(null);
    localStorage.removeItem("g_sheets_token");
    setSheetsMessage({ text: "Signed out of Google Account.", type: "info" });
  };

  const handleSpreadsheetIdChange = (id: string) => {
    let cleanId = id.trim();
    if (cleanId.includes("docs.google.com/spreadsheets")) {
      const match = cleanId.match(/\/d\/([a-zA-Z0-9-_]+)/);
      if (match && match[1]) {
        cleanId = match[1];
      }
    }
    setSpreadsheetId(cleanId);
    localStorage.setItem("g_sheets_spreadsheet_id", cleanId);
  };

  const handleCreateNewSpreadsheet = async () => {
    if (!googleAccessToken) return;
    setSheetsLoading(true);
    setSheetsMessage({ text: "Creating a new Google Spreadsheet on your Drive...", type: "info" });
    try {
      const response = await fetch("https://sheets.googleapis.com/v4/spreadsheets", {
        method: "POST",
        headers: {
          "Authorization": `Bearer ${googleAccessToken}`,
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          properties: {
            title: `Employee QA Evaluation Reports - ${config.projectName || "General"}`
          },
          sheets: [
            {
              properties: {
                title: "Evaluation Summary Logs"
              }
            }
          ]
        })
      });

      if (!response.ok) {
        throw new Error("Failed to create spreadsheet (HTTP " + response.status + ")");
      }

      const data = await response.json();
      const newId = data.spreadsheetId;
      handleSpreadsheetIdChange(newId);

      // Initialize headers automatically
      await initSpreadsheetHeaders(newId);

      setSheetsMessage({ text: `Success! Created new Spreadsheet. Excel headers initialized.`, type: "success" });
    } catch (err: any) {
      console.error("Create spreadsheet error:", err);
      setSheetsMessage({ text: "Failed to create new Spreadsheet: " + err.message, type: "error" });
    } finally {
      setSheetsLoading(false);
    }
  };

  const initSpreadsheetHeaders = async (sid: string) => {
    try {
      const headers = [
        "Audit Date",
        "Employee Name",
        "Project Name",
        "Accuracy Grade",
        "Accuracy",
        "Checked Coordinates",
        "Total Error Count",
        "Penalty Accumulation",
        "Reviewer Burden Index (Hrs)",
        "Major Error Categories",
        "Source Client Url"
      ];
      await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${sid}/values/${encodeURIComponent("Evaluation Summary Logs!A1")}:append?valueInputOption=USER_ENTERED`, {
        method: "POST",
        headers: {
          "Authorization": `Bearer ${googleAccessToken}`,
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          values: [headers]
        })
      });
    } catch (err) {
      console.error("Header auto-initialization error:", err);
    }
  };

  const handleSaveReportToSheets = async () => {
    if (!qaAnalysis) return;

    if (!isEmployeeNameValid) {
      setSheetsMessage({ text: "Cannot save report: Employee Name is required. Please write a valid name in core configurations.", type: "error" });
      return;
    }

    if (!googleAccessToken) {
      setSheetsMessage({ text: "Sign in with Google is required to connect to Google Sheets.", type: "error" });
      return;
    }

    if (!spreadsheetId.trim()) {
      setSheetsMessage({ text: "Please provide a valid Spreadsheet ID or click 'Create New Spreadsheet'.", type: "error" });
      return;
    }

    setSheetsLoading(true);
    setSheetsMessage({ text: "Saving evaluation logs... Appending row to Google Sheet...", type: "info" });

    try {
      const patternSummary = [
        qaAnalysis.rootCause.missingValuesPct > 0 ? "Omissions/Extras" : "",
        qaAnalysis.rootCause.numericErrorsPct > 0 ? "Numeric Typos" : "",
        qaAnalysis.rootCause.textErrorsPct > 0 ? "Spelling Typos" : "",
        qaAnalysis.patterns.shiftEvents.length > 0 ? "Alignment Shifts" : ""
      ].filter(Boolean).join(", ") || "None";

      const rowValues = [
        config.evaluationDate,
        config.employeeName,
        config.projectName,
        qaAnalysis.metrics.finalGrade,
        `${qaAnalysis.metrics.baseAccuracy}%`,
        qaAnalysis.metrics.comparedCells,
        qaAnalysis.metrics.totalErrors,
        qaAnalysis.metrics.totalPenaltyPoints,
        qaAnalysis.metrics.reviewerWorkloadIndex.toFixed(2),
        patternSummary,
        window.location.href
      ];

      const tryAppend = async (range: string) => {
        return fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(range)}:append?valueInputOption=USER_ENTERED`, {
          method: "POST",
          headers: {
            "Authorization": `Bearer ${googleAccessToken}`,
            "Content-Type": "application/json"
          },
          body: JSON.stringify({
            values: [rowValues]
          })
        });
      };

      // Try Evaluation Summary Logs first, fallback to Sheet1 or any A:L
      let res = await tryAppend("Evaluation Summary Logs!A:L");
      if (!res.ok) {
        res = await tryAppend("A:L");
      }

      if (!res.ok) {
        throw new Error("Append request failed with HTTP " + res.status);
      }

      setSheetsMessage({ text: `Report row saved successfully! Report for "${config.employeeName}" has been appended to Google Sheet.`, type: "success" });
    } catch (err: any) {
      console.error("Append row error:", err);
      setSheetsMessage({ text: "Failed to append logs: " + err.message, type: "error" });
    } finally {
      setSheetsLoading(false);
    }
  };

  const loadDemoState = () => {
    const demWorker = getDemoEmployeeData();
    const demRev = getDemoReviewerData();
    setEmployeeWb(demWorker);
    setReviewerWb(demRev);
    setErrorMsg(null);
    setConfig(c => ({
      ...c,
      employeeName: "Farid Al-Mansour"
    }));
    setIsComparingStarted(false);
  };

  const handleBaselineJSONUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setBaselineFileName(file.name);
    setBaselineUploadError(null);
    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const json = JSON.parse(event.target?.result as string);
        if (json && (json.metrics || json.accuracy !== undefined)) {
          const normalizedBaseline = {
            accuracy: json.metrics?.accuracyRatio ?? json.accuracy ?? 1,
            totalErrors: json.errorLog?.length ?? json.totalErrors ?? 0,
            penaltyScore: json.metrics?.totalStructuralPenalty ?? json.penaltyScore ?? 0,
            grade: json.metrics?.grade ?? json.grade ?? 'A',
            date: json.config?.evaluationDate ?? json.date ?? new Date().toLocaleDateString(),
            projectName: json.config?.projectName ?? json.projectName ?? "Imported Run"
          };
          setBenchmarkBaseline(normalizedBaseline);
        } else {
          throw new Error("Invalid baseline schema. Missing quality metrics.");
        }
      } catch (err: any) {
        setBaselineUploadError("Failed to parse JSON baseline: " + err.message);
        setBenchmarkBaseline(null);
      }
    };
    reader.readAsText(file);
  };

  const saveCurrentAsBaseline = () => {
    if (!qaAnalysis) return;
    const currentBaseline = {
      accuracy: qaAnalysis.metrics.accuracyRatio,
      totalErrors: qaAnalysis.errorLog.length,
      penaltyScore: qaAnalysis.metrics.totalStructuralPenalty || 0,
      grade: qaAnalysis.metrics.grade || 'C',
      date: config.evaluationDate || new Date().toLocaleDateString(),
      projectName: config.projectName || "Active Run"
    };
    setBenchmarkBaseline(currentBaseline);
    setBaselineFileName("Active Session Snapshot");
    setBaselineUploadError(null);
  };

  const exportCurrentRunJSON = () => {
    if (!qaAnalysis) return;
    const blob = new Blob([JSON.stringify(qaAnalysis, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `QA_Compliance_Report_${config.projectName.replace(/\s+/g, '_')}_${new Date().toISOString().slice(0, 10)}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  // Trigger Gemini AI professional assessor route
  const triggerAIAudit = async () => {
    if (!qaAnalysis) return;
    setAiLoading(true);
    setAiResult(null);

    // Filter down sample raw errors to prevent token overloads (Top 10 severe samples only)
    const sortedSample = [...qaAnalysis.errorLog]
      .sort((a, b) => b.penalty - a.penalty)
      .slice(0, 10)
      .map(x => ({
        sheet: x.sheet,
        coordinate: x.cell,
        workerInput: x.employeeValue,
        correctValue: x.reviewerValue,
        normalizedWorker: x.normalizedEmployeeValue,
        normalizedReviewer: x.normalizedReviewerValue,
        similarityRatio: x.similarity,
        classificationGroup: x.errorType,
        penaltyWeight: x.penalty
      }));

    try {
      const res = await fetch("/api/coach", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          workerName: config.employeeName,
          projectName: config.projectName,
          date: config.evaluationDate,
          metrics: qaAnalysis.metrics,
          rootCause: qaAnalysis.rootCause,
          topErrors: sortedSample,
          patterns: {
            repeatedNumericErrors: qaAnalysis.patterns.repeatedNumericErrors,
            copyPasteErrors: qaAnalysis.patterns.copyPasteErrors,
            errorClusters: qaAnalysis.patterns.errorClusters,
            shiftEvents: qaAnalysis.patterns.shiftEvents
          }
        })
      });

      if (!res.ok) throw new Error("Route connection problem");
      const auditorData = await res.json();
      setAiResult(auditorData);
      setActiveTab("aiAuditor");
    } catch (err: any) {
      console.error("Failed AI route fetch:", err);
      // Fallback local UI mock notice
      setAiResult({
        summary: `### Audit Analysis Output Error\n\nCould not access server-side AI evaluation routes directly. Verify local development port mappings.`,
        coaching: [
          "Double check your Express Server terminal logging output to verify route ingress status on Port 3000."
        ]
      });
      setActiveTab("aiAuditor");
    } finally {
      setAiLoading(false);
    }
  };

  // Processing detailed error logs sorting & filter
  const processedErrorLogs = useMemo(() => {
    if (!qaAnalysis) return [];
    let logs = [...qaAnalysis.errorLog];

    // Search filter
    if (logSearch) {
      const query = logSearch.toLowerCase();
      logs = logs.filter(
        x =>
          x.sheet.toLowerCase().includes(query) ||
          x.cell.toLowerCase().includes(query) ||
          x.employeeValue.toLowerCase().includes(query) ||
          x.reviewerValue.toLowerCase().includes(query) ||
          x.errorType.toLowerCase().includes(query) ||
          x.notes.toLowerCase().includes(query)
      );
    }

    // Severity filter
    if (logFilterSeverity !== "ALL") {
      logs = logs.filter(x => x.severity.toUpperCase() === logFilterSeverity);
    }

    // Type filter
    if (logFilterType !== "ALL") {
      logs = logs.filter(x => x.errorType === logFilterType);
    }

    // Sort logs: Severity highest -> Penalty highest -> Location coordinate
    logs.sort((a, b) => {
      if (logSortField === "severity") {
        const sevWeight = { [Severity.Critical]: 4, [Severity.High]: 3, [Severity.Medium]: 2, [Severity.Low]: 1 };
        return sevWeight[b.severity] - sevWeight[a.severity];
      }
      if (logSortField === "penalty") {
        return b.penalty - a.penalty;
      }
      if (logSortField === "sheet") {
        return a.sheet.localeCompare(b.sheet);
      }
      if (logSortField === "cell") {
        return a.cell.localeCompare(b.cell);
      }
      return 0;
    });

    return logs;
  }, [qaAnalysis, logSearch, logFilterSeverity, logFilterType, logSortField]);

  // Is display truncated for output 3 displays? (Total Errors > 50)
  const isTruncated = qaAnalysis ? qaAnalysis.errorLog.length > 50 && !forceShowAllErrors : false;
  const displayLogs = isTruncated ? processedErrorLogs.slice(0, 50) : processedErrorLogs;

  const totalErrorsCount = qaAnalysis ? qaAnalysis.errorLog.length : 0;

  const [copiedUrl, setCopiedUrl] = useState<boolean>(false);

  const copySandboxUrlToClipboard = () => {
    navigator.clipboard.writeText(window.location.href);
    setCopiedUrl(true);
    setTimeout(() => setCopiedUrl(false), 2000);
  };

  const generatePDFCertificate = async () => {
    if (!qaAnalysis) return;
    const projectName = config.projectName ? config.projectName.replace(/\s+/g, '_') : "Audit";
    
    try {
      const element = document.getElementById("pdf-certificate-content");
      if (!element) {
        throw new Error("PDF Certificate content element ('pdf-certificate-content') was not found in the DOM.");
      }

      const pdf = new jsPDF({
        orientation: "portrait",
        unit: "pt",
        format: "a4"
      });

      // Render the offscreen HTML template to custom pages
      await pdf.html(element, {
        html2canvas: {
          scale: 1.5, // Enhances typography crispness and prevents blur
          useCORS: true,
          logging: false
        },
        x: 40,
        y: 40,
        width: 515.28,    // 595.28 (A4 width) - 80 (margins left + right)
        windowWidth: 800, // Forces responsive styles viewport width
        margin: [40, 40, 60, 40], // Margins: [top, left, bottom, right]
        autoPaging: "text",
        callback: function (doc) {
          try {
            const totalPages = doc.getNumberOfPages();
            const timestamp = new Date().toLocaleString("en-US", {
              year: "numeric",
              month: "2-digit",
              day: "2-digit",
              hour: "2-digit",
              minute: "2-digit",
              second: "2-digit",
              hour12: false
            });

            // Post-process each page to add standardized page numbers and generation timestamp
            for (let i = 1; i <= totalPages; i++) {
              doc.setPage(i);
              doc.setFontSize(8);
              doc.setTextColor(148, 163, 184); // slate-400 color
              
              // Draw footer lines
              const footerY = 810; // A4 height is 841.89
              doc.text(`Generated on: ${timestamp} | Reference: ${config.evaluationDate}`, 40, footerY);
              doc.text(`Page ${i} of ${totalPages}`, 555, footerY, { align: "right" });
            }

            doc.save(`Compliance_Certificate_${projectName}.pdf`);
          } catch (internalErr: any) {
            console.error("Internal PDF styling finalizing failed:", internalErr);
            alert("PDF generation failed during final page-numbering step: " + internalErr.message);
          }
        }
      });
    } catch (error: any) {
      console.error("Critical PDF Generation Error:", error);
      alert("Failed to compile and export the requested PDF Certificate: " + error.message);
    }
  };

  // Print friendly audit generator
  const triggerPrint = () => {
    const isIframe = window.self !== window.top;
    if (isIframe) {
      console.warn("Iframe Sandbox Mode Restriction Checked - standard print triggered option fallback.");
      setPrintError("Direct browser printing triggers are limited inside secure embedded iframe environments.");
      setShowPrintModal(true);
      return;
    }
    try {
      if (typeof window.print !== "function") {
        throw new Error("Local print API is not supported in this frame.");
      }
      window.print();
    } catch (err: any) {
      console.warn("Standard printing failed or was blocked: ", err);
      setPrintError(err?.message || "Iframe Sandbox Mode Restriction Detected.");
      setShowPrintModal(true);
    }
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 font-sans selection:bg-indigo-100 flex flex-col">
      
      {/* 🚀 Top Premium Header */}
      <header className="flex items-center justify-between px-6 py-4 bg-white border-b border-slate-200 sticky top-0 z-40 shadow-xs print:hidden">
        <div className="flex items-center space-x-3">
          <div className="w-8 h-8 bg-indigo-600 rounded flex items-center justify-center flex-shrink-0">
            <FileSpreadsheet className="w-5 h-5 text-white" />
          </div>
          <div>
            <h1 className="text-base sm:text-lg font-bold text-slate-800 tracking-tight">Enterprise QA Engine</h1>
            <p className="text-2xs sm:text-xs text-slate-500 font-medium whitespace-nowrap overflow-hidden text-ellipsis">Project: {config.projectName || "Demo"} • Evaluated: {config.evaluationDate}</p>
          </div>
        </div>

        <div className="flex items-center space-x-2">
          {qaAnalysis && qaAnalysis.patterns.shiftEvents.length > 0 && (
            <div className="hidden lg:flex px-3 py-1 bg-amber-50 border border-amber-200 text-amber-700 text-3xs sm:text-2xs font-bold rounded items-center">
              <span className="w-2 h-2 bg-amber-500 rounded-full mr-2 animate-pulse"></span>
              CRITICAL SHIFT DETECTED
            </div>
          )}

          {!qaAnalysis ? (
            <button
              onClick={loadDemoState}
              className="px-4 py-1.5 bg-indigo-650 hover:bg-indigo-700 text-white text-xs font-semibold rounded cursor-pointer transition shadow-xs"
            >
              Load Preset Audit Demo
            </button>
          ) : (
            <div className="flex items-center gap-2">
              <button
                onClick={triggerPrint}
                className="px-4 py-1.5 bg-slate-800 text-white text-xs font-semibold rounded hover:bg-slate-700 cursor-pointer transition shadow-xs"
              >
                Print Compliance PDF
              </button>
            </div>
          )}
        </div>
      </header>

      {/* 📁 Upload Panel with File selection controls */}
      {!qaAnalysis && (
        <section className="max-w-7xl mx-auto px-6 py-12">
          <div className="text-center max-w-2xl mx-auto mb-10">
            <span className="text-[10px] font-mono bg-indigo-50 text-indigo-700 px-3 py-1 rounded-full uppercase tracking-wider font-bold border border-indigo-100">
              Deterministic Shift Precedence Engine
            </span>
            <h2 className="text-2xl font-extrabold tracking-tight text-slate-800 mt-4 sm:text-3xl">
              Evaluate Excel & CSV Data Authenticity
            </h2>
            <p className="text-sm text-slate-500 mt-2 leading-relaxed">
              Securely analyze participant worksheets against source reviewers. This sandbox automatically isolates layout offsets, Arabic letter spelling permutations, dates, and transposed mathematical digits.
            </p>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-8 mb-8">
            {/* File A: Worker File */}
            <div className="bg-white rounded-xl border border-slate-200 p-6 flex flex-col justify-between shadow-sm">
              <div>
                <div className="flex justify-between items-center mb-4">
                  <h3 className="text-xs font-bold tracking-wider text-slate-650 uppercase font-mono">
                    File A — Employee Submission
                  </h3>
                  {employeeWb && (
                    <span className="bg-indigo-50 text-indigo-700 border border-indigo-150 px-2.5 py-0.5 rounded-md text-2xs font-semibold">
                      Uploaded
                    </span>
                  )}
                </div>
                
                <label
                  onDragEnter={(e) => handleDrag(e, "emp")}
                  onDragOver={(e) => handleDrag(e, "emp")}
                  onDragLeave={(e) => handleDrag(e, "emp")}
                  onDrop={(e) => handleDrop(e, "emp")}
                  className={`border-2 border-dashed rounded-lg p-8 flex flex-col items-center justify-center transition cursor-pointer select-none ${
                    dragActiveE ? "border-indigo-500 bg-indigo-50/50" : "border-slate-200 bg-slate-50/70 hover:bg-slate-50"
                  }`}
                >
                  <Upload className="w-8 h-8 text-indigo-400 mb-3" />
                  <p className="text-xs font-medium text-slate-750 text-center pointer-events-none">
                    Drag and drop employee file here, or{" "}
                    <span className="text-indigo-600 hover:text-indigo-700 underline font-semibold">
                      browse
                    </span>
                  </p>
                  <input
                    type="file"
                    accept=".xlsx,.xls,.csv"
                    className="hidden"
                    onChange={(e) => handleFileSelect(e, "emp")}
                  />
                  <p className="text-3xs text-slate-400 mt-1 uppercase font-mono tracking-wider pointer-events-none">
                    XLSX, XLS, CSV format supported
                  </p>
                </label>
              </div>

              {employeeWb ? (
                <div className="mt-4 p-3 bg-slate-50 border border-slate-100 rounded-lg flex items-center justify-between">
                  <span className="text-xs font-semibold truncate max-w-[200px] text-slate-700 font-mono">
                    {employeeWb.fileName}
                  </span>
                  <button
                    onClick={() => {
                      setEmployeeWb(null);
                      setIsComparingStarted(false);
                    }}
                    className="text-2xs text-red-600 underline font-semibold hover:text-red-700 cursor-pointer"
                  >
                    Remove
                  </button>
                </div>
              ) : (
                <div className="mt-4 text-center text-xs text-slate-400 italic">No file loaded</div>
              )}
            </div>

            {/* File B: Reviewer File */}
            <div className="bg-white rounded-xl border border-slate-200 p-6 flex flex-col justify-between shadow-sm">
              <div>
                <div className="flex justify-between items-center mb-4">
                  <h3 className="text-xs font-bold tracking-wider text-slate-655 uppercase font-mono">
                    File B — Reviewer / Ground Truth
                  </h3>
                  {reviewerWb && (
                    <span className="bg-indigo-50 text-indigo-700 border border-indigo-150 px-2.5 py-0.5 rounded-md text-2xs font-semibold">
                      Uploaded
                    </span>
                  )}
                </div>

                <label
                  onDragEnter={(e) => handleDrag(e, "rev")}
                  onDragOver={(e) => handleDrag(e, "rev")}
                  onDragLeave={(e) => handleDrag(e, "rev")}
                  onDrop={(e) => handleDrop(e, "rev")}
                  className={`border-2 border-dashed rounded-lg p-8 flex flex-col items-center justify-center transition cursor-pointer select-none ${
                    dragActiveR ? "border-indigo-500 bg-indigo-50/50" : "border-slate-200 bg-slate-50/70 hover:bg-slate-50"
                  }`}
                >
                  <Upload className="w-8 h-8 text-indigo-400 mb-3" />
                  <p className="text-xs font-medium text-slate-750 text-center pointer-events-none">
                    Drag and drop reviewer file here, or{" "}
                    <span className="text-indigo-600 hover:text-indigo-700 underline font-semibold">
                      browse
                    </span>
                  </p>
                  <input
                    type="file"
                    accept=".xlsx,.xls,.csv"
                    className="hidden"
                    onChange={(e) => handleFileSelect(e, "rev")}
                  />
                  <p className="text-3xs text-slate-400 mt-1 uppercase font-mono tracking-wider pointer-events-none">
                    XLSX, XLS, CSV format supported
                  </p>
                </label>
              </div>

              {reviewerWb ? (
                <div className="mt-4 p-3 bg-slate-50 border border-slate-100 rounded-lg flex items-center justify-between">
                  <span className="text-xs font-semibold truncate max-w-[200px] text-slate-700 font-mono">
                    {reviewerWb.fileName}
                  </span>
                  <button
                    onClick={() => {
                      setReviewerWb(null);
                      setIsComparingStarted(false);
                    }}
                    className="text-2xs text-red-600 underline font-semibold hover:text-red-700 cursor-pointer"
                  >
                    Remove
                  </button>
                </div>
              ) : (
                <div className="mt-4 text-center text-xs text-slate-400 italic">No file loaded</div>
              )}
            </div>
          </div>

          {errorMsg && (
            <div className="mb-6 p-4 bg-red-50 text-red-800 rounded-lg border border-red-200 text-xs flex items-center gap-2">
              <AlertCircle className="w-4 h-4 flex-shrink-0" />
              <span>{errorMsg}</span>
            </div>
          )}

          {/* 📊 Start Comparison Audit Section (Only visible when 2 files are uploaded and comparison is not started) */}
          {employeeWb && reviewerWb && !isComparingStarted && (
            <div className="bg-white border-2 border-indigo-200 rounded-xl p-6 shadow-md mb-8 flex flex-col md:flex-row items-center justify-between gap-6">
              <div className="flex items-start gap-4">
                <div className="p-3 bg-indigo-50 text-indigo-700 rounded-lg shrink-0">
                  <Play className="w-6 h-6 animate-pulse-slow fill-indigo-100" />
                </div>
                <div>
                  <h3 className="font-extrabold text-sm text-slate-800">Files Are Ready For Audit Evaluation</h3>
                  <p className="text-2xs sm:text-xs text-slate-500 mt-1 leading-relaxed max-w-xl">
                    Both sheets are parsed successfully. Please enter a genuine <strong className="text-slate-700">Employee Name</strong> in the configurations card below to enable and start comparing submissions.
                  </p>
                </div>
              </div>
              
              <div className="flex flex-col items-stretch sm:items-end w-full md:w-auto gap-2 shrink-0">
                <button
                  type="button"
                  id="btn-start-evaluation"
                  disabled={!isEmployeeNameValid}
                  onClick={() => setIsComparingStarted(true)}
                  className={`px-6 py-3 rounded-xl border flex items-center justify-center gap-2 font-extrabold select-none text-xs transition duration-150 ${
                    !isEmployeeNameValid
                      ? "bg-slate-100 text-slate-400 border-slate-200 cursor-not-allowed"
                      : "bg-indigo-600 hover:bg-indigo-700 text-white border-indigo-700 hover:border-indigo-800 cursor-pointer shadow-md"
                  }`}
                >
                  <Sparkles className="w-4 h-4 shrink-0" />
                  Start Comparison Audit
                </button>
                {!isEmployeeNameValid && (
                  <p className="text-[10px] text-amber-600 font-medium text-center sm:text-right flex items-center justify-center sm:justify-end gap-1">
                    <AlertCircle className="w-3.5 h-3.5 text-amber-500 shrink-0 animate-bounce" />
                    Employee name required to start comparing
                  </p>
                )}
              </div>
            </div>
          )}

          {/* Configuration drawer settings for QA */}
          <div className="bg-white border border-slate-200 rounded-xl p-6 shadow-sm">
            <div className="flex items-center gap-2 mb-4 border-b border-slate-100 pb-3">
              <SlidersHorizontal className="w-4 h-4 text-indigo-600" />
              <h3 className="font-bold text-xs uppercase tracking-wider font-mono text-slate-700">QA Engine Core Threshold Configurations</h3>
            </div>
            
            <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-6 text-xs text-slate-750">
              <div>
                <label className="block text-slate-600 font-semibold mb-1">Employee Name <span className="text-red-500 font-bold">*</span></label>
                <input
                  type="text"
                  value={config.employeeName}
                  onChange={(e) => setConfig({ ...config, employeeName: e.target.value })}
                  className={`w-full px-3 py-2 border rounded-md focus:outline-none focus:ring-2 transition duration-150 ${
                    !isEmployeeNameValid
                      ? "border-amber-400 bg-amber-50/10 focus:ring-amber-500/20 focus:border-amber-500 font-semibold text-amber-800"
                      : "border-slate-200 focus:ring-indigo-505/20 focus:border-indigo-500"
                  }`}
                  placeholder="e.g. Ahmed"
                />
                {!isEmployeeNameValid && (
                  <p className="text-[10px] text-amber-600 mt-1.5 font-medium flex items-center gap-1">
                    <AlertCircle className="w-3.5 h-3.5 text-amber-500 shrink-0" />
                    Genuine employee name is required instead of "Employee Name" default.
                  </p>
                )}
              </div>
              <div>
                <label className="block text-slate-600 font-semibold mb-1">Project Identifier (Optional)</label>
                <input
                  type="text"
                  value={config.projectName}
                  onChange={(e) => setConfig({ ...config, projectName: e.target.value })}
                  className="w-full px-3 py-2 border border-slate-200 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-505/20 focus:border-indigo-500 transition duration-150"
                  placeholder="e.g. Ingress Survey"
                />
              </div>
              <div>
                <label className="block text-slate-600 font-semibold mb-1">Audit Evaluation Date</label>
                <input
                  type="date"
                  value={config.evaluationDate}
                  onChange={(e) => setConfig({ ...config, evaluationDate: e.target.value })}
                  className="w-full px-3 py-2 border border-slate-200 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-550/20 focus:border-indigo-550 transition duration-150"
                />
              </div>

              <div>
                <label className="block text-slate-600 font-semibold mb-1">Numeric Major Variance Threshold (%)</label>
                <input
                  type="number"
                  step="0.05"
                  value={config.numericMajorVarianceThreshold}
                  onChange={(e) => setConfig({ ...config, numericMajorVarianceThreshold: parseFloat(e.target.value) })}
                  className="w-full px-3 py-2 border border-slate-200 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-550/20 focus:border-indigo-550 transition duration-150"
                />
                <span className="text-[9px] text-slate-400 block mt-0.5 font-mono uppercase tracking-wider">Discrepancy trigger: default 20%</span>
              </div>
              <div>
                <label className="block text-slate-600 font-semibold mb-1">Numeric Major Absolute Variance</label>
                <input
                  type="number"
                  value={config.numericMajorAbsoluteThreshold}
                  onChange={(e) => setConfig({ ...config, numericMajorAbsoluteThreshold: parseFloat(e.target.value) })}
                  className="w-full px-3 py-2 border border-slate-200 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-550/20 focus:border-indigo-550 transition duration-150"
                />
                <span className="text-[9px] text-slate-400 block mt-0.5 font-mono uppercase tracking-wider">Absolute margin override: default 5.0</span>
              </div>
              <div>
                <label className="block text-slate-600 font-semibold mb-1">Arabic Text Comparison Mode</label>
                <select
                  value={config.arabicComparisonMode}
                  onChange={(e) => setConfig({ ...config, arabicComparisonMode: e.target.value as "STANDARD" | "NONE" })}
                  className="w-full px-3 py-2 border border-slate-200 bg-white rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-550/20 focus:border-indigo-550 transition duration-150 cursor-pointer"
                >
                  <option value="STANDARD">STANDARD Normalization</option>
                  <option value="NONE">NONE Exact Check</option>
                </select>
                <span className="text-[9px] text-slate-400 block mt-0.5 font-mono uppercase tracking-wider">Removes diacritics, joins Te Marbuta</span>
              </div>

              <div>
                <label className="block text-slate-650 font-semibold mb-1">Minimum Shift Cell Span (Cells)</label>
                <input
                  type="number"
                  value={config.minimumShiftCells}
                  onChange={(e) => setConfig({ ...config, minimumShiftCells: parseInt(e.target.value) })}
                  className="w-full px-3 py-2 border border-slate-200 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-550/20 focus:border-indigo-550 transition duration-150"
                />
                <span className="text-[9px] text-slate-400 block mt-0.5 font-mono uppercase tracking-wider">Size bound for Shift Trigger: default 20</span>
              </div>
              <div>
                <label className="block text-slate-650 font-semibold mb-1">Shift Consistency Alignment Ratio (%)</label>
                <input
                  type="number"
                  step="0.05"
                  value={config.shiftDetectionThreshold}
                  onChange={(e) => setConfig({ ...config, shiftDetectionThreshold: parseFloat(e.target.value) })}
                  className="w-full px-3 py-2 border border-slate-200 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-550/20 focus:border-indigo-550 transition duration-150"
                />
                <span className="text-[9px] text-slate-400 block mt-0.5 font-mono uppercase tracking-wider">Pattern density score: default 80%</span>
              </div>
              <div>
                <label className="block text-slate-600 font-semibold mb-1">Numeric Tolerance Mode</label>
                <div className="flex gap-2">
                  <input
                    type="number"
                    step="0.001"
                    value={config.numericTolerance}
                    onChange={(e) => setConfig({ ...config, numericTolerance: parseFloat(e.target.value) || 0 })}
                    className="w-2/3 px-3 py-2 border border-slate-200 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-550/20 focus:border-indigo-550 transition duration-150"
                  />
                  <select
                    value={config.numericToleranceMode}
                    onChange={(e) => setConfig({ ...config, numericToleranceMode: e.target.value as "PERCENTAGE" | "ABSOLUTE" })}
                    className="w-1/3 px-2 py-2 border border-slate-200 bg-white rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-550/20 focus:border-indigo-550 transition duration-150 cursor-pointer text-[11px]"
                  >
                    <option value="PERCENTAGE">%</option>
                    <option value="ABSOLUTE">ABS</option>
                  </select>
                </div>
                <span className="text-[9px] text-slate-400 block mt-0.5 font-mono uppercase tracking-wider">Tolerance buffer (Default 0.01 = 1%)</span>
              </div>

              <div>
                <label className="block text-slate-600 font-semibold mb-1">Audit Grade Strict Mode</label>
                <select
                  value={config.strictMode}
                  onChange={(e) => setConfig({ ...config, strictMode: e.target.value as "AUTO" | "ON" | "OFF" })}
                  className="w-full px-3 py-2 border border-slate-200 bg-white rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-550/20 focus:border-indigo-550 transition duration-150 cursor-pointer"
                >
                  <option value="AUTO">AUTO Detect (Financial/Census)</option>
                  <option value="ON">ON (Strict Zero Tolerance)</option>
                  <option value="OFF">OFF (Normal Tolerance)</option>
                </select>
                <span className="text-[9px] text-slate-400 block mt-0.5 font-mono uppercase tracking-wider">Enforces absolute correctness in critical surveys</span>
              </div>

              <div>
                <label className="block text-slate-600 font-semibold mb-1">Header Row Error Penalty</label>
                <input
                  type="number"
                  value={config.headerPenalty}
                  onChange={(e) => setConfig({ ...config, headerPenalty: parseInt(e.target.value) || 3 })}
                  className="w-full px-3 py-2 border border-slate-200 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-550/20 focus:border-indigo-550 transition duration-150"
                />
                <span className="text-[9px] text-slate-400 block mt-0.5 font-mono uppercase tracking-wider">Penalizes label inaccuracies (Default: 3)</span>
              </div>
            </div>

            <div className="mt-6 border-t border-slate-200 pt-5 pr-1 pl-1">
              <div className="flex items-center gap-2 mb-4">
                <SlidersHorizontal className="w-4 h-4 text-emerald-600" />
                <h4 className="font-bold text-xs uppercase tracking-wider font-mono text-slate-700">Audit Score Penalty Weight Coefficients</h4>
              </div>
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-3 gap-4 text-xs">
                {/* Extra Table Coefficient */}
                <div className="bg-slate-50/50 p-3 rounded-lg border border-slate-100">
                  <label className="block text-slate-700 font-bold mb-1">Extra Table Coefficient</label>
                  <input
                    type="number"
                    value={config.extraTableCoefficient}
                    onChange={(e) => setConfig({ ...config, extraTableCoefficient: parseFloat(e.target.value) || 0 })}
                    className="w-full px-3 py-1.5 border border-slate-200 bg-white rounded-md focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 transition duration-150"
                  />
                  <span className="text-[9px] text-slate-400 block mt-1">Defect penalty (Default: 50)</span>
                </div>

                {/* Missing Table Coefficient */}
                <div className="bg-slate-50/50 p-3 rounded-lg border border-slate-100">
                  <label className="block text-slate-700 font-bold mb-1">Missing Table Coefficient</label>
                  <input
                    type="number"
                    value={config.missingTableCoefficient}
                    onChange={(e) => setConfig({ ...config, missingTableCoefficient: parseFloat(e.target.value) || 0 })}
                    className="w-full px-3 py-1.5 border border-slate-200 bg-white rounded-md focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 transition duration-150"
                  />
                  <span className="text-[9px] text-slate-400 block mt-1">Omission penalty (Default: 100)</span>
                </div>

                {/* Extra Column Coefficient */}
                <div className="bg-slate-50/50 p-3 rounded-lg border border-slate-100">
                  <label className="block text-slate-700 font-bold mb-1">Extra Column Coefficient</label>
                  <input
                    type="number"
                    value={config.extraColumnCoefficient}
                    onChange={(e) => setConfig({ ...config, extraColumnCoefficient: parseFloat(e.target.value) || 0 })}
                    className="w-full px-3 py-1.5 border border-slate-200 bg-white rounded-md focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 transition duration-150"
                  />
                  <span className="text-[9px] text-slate-400 block mt-1">Defect penalty (Default: 5)</span>
                </div>

                {/* Missing Column Coefficient */}
                <div className="bg-slate-50/50 p-3 rounded-lg border border-slate-100">
                  <label className="block text-slate-700 font-bold mb-1">Missing Column Coefficient</label>
                  <input
                    type="number"
                    value={config.missingColumnCoefficient}
                    onChange={(e) => setConfig({ ...config, missingColumnCoefficient: parseFloat(e.target.value) || 0 })}
                    className="w-full px-3 py-1.5 border border-slate-200 bg-white rounded-md focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 transition duration-150"
                  />
                  <span className="text-[9px] text-slate-400 block mt-1">Omission penalty (Default: 10)</span>
                </div>

                {/* Extra Row Coefficient */}
                <div className="bg-slate-50/50 p-3 rounded-lg border border-slate-100">
                  <label className="block text-slate-700 font-bold mb-1">Extra Row Coefficient</label>
                  <input
                    type="number"
                    value={config.extraRowCoefficient}
                    onChange={(e) => setConfig({ ...config, extraRowCoefficient: parseFloat(e.target.value) || 0 })}
                    className="w-full px-3 py-1.5 border border-slate-200 bg-white rounded-md focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 transition duration-150"
                  />
                  <span className="text-[9px] text-slate-400 block mt-1">Defect penalty (Default: 1)</span>
                </div>

                {/* Missing Row Coefficient */}
                <div className="bg-slate-50/50 p-3 rounded-lg border border-slate-100">
                  <label className="block text-slate-700 font-bold mb-1">Missing Row Coefficient</label>
                  <input
                    type="number"
                    value={config.missingRowCoefficient}
                    onChange={(e) => setConfig({ ...config, missingRowCoefficient: parseFloat(e.target.value) || 0 })}
                    className="w-full px-3 py-1.5 border border-slate-200 bg-white rounded-md focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 transition duration-150"
                  />
                  <span className="text-[9px] text-slate-400 block mt-1">Omission penalty (Default: 2)</span>
                </div>

                {/* Numeric Difference Coefficient */}
                <div className="bg-slate-50/50 p-3 rounded-lg border border-slate-100">
                  <label className="block text-slate-700 font-bold mb-1">Numeric Difference Coeff.</label>
                  <input
                    type="number"
                    step="0.01"
                    value={config.numericDifferenceCoefficient}
                    onChange={(e) => setConfig({ ...config, numericDifferenceCoefficient: parseFloat(e.target.value) || 0 })}
                    className="w-full px-3 py-1.5 border border-slate-200 bg-white rounded-md focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 transition duration-150"
                  />
                  <span className="text-[9px] text-slate-400 block mt-1">Data divergence (Default: 0.1)</span>
                </div>

                {/* Text Difference Coefficient */}
                <div className="bg-slate-50/50 p-3 rounded-lg border border-slate-100">
                  <label className="block text-slate-700 font-bold mb-1">Text Difference Coeff.</label>
                  <input
                    type="number"
                    step="0.01"
                    value={config.textDifferenceCoefficient}
                    onChange={(e) => setConfig({ ...config, textDifferenceCoefficient: parseFloat(e.target.value) || 0 })}
                    className="w-full px-3 py-1.5 border border-slate-200 bg-white rounded-md focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 transition duration-150"
                  />
                  <span className="text-[9px] text-slate-400 block mt-1">Data spelling (Default: 0.1)</span>
                </div>

                {/* Empty Cell Difference Coefficient */}
                <div className="bg-slate-50/50 p-3 rounded-lg border border-slate-100 sm:col-span-2 lg:col-span-3">
                  <label className="block text-slate-700 font-bold mb-1">Empty Cell Difference Coefficient</label>
                  <input
                    type="number"
                    step="0.01"
                    value={config.emptyCellDifferenceCoefficient}
                    onChange={(e) => setConfig({ ...config, emptyCellDifferenceCoefficient: parseFloat(e.target.value) || 0 })}
                    className="w-full px-3 py-1.5 border border-slate-200 bg-white rounded-md focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 transition duration-150"
                  />
                  <span className="text-[9px] text-slate-400 block mt-1">Data omission (Default: 0.05)</span>
                </div>
              </div>
            </div>

            {/* Auto Ignore Sheets System Control */}
            <div className="mt-6 border-t border-slate-200 pt-5 pr-1 pl-1">
              <div className="flex items-center gap-2 mb-4">
                <Filter className="w-4 h-4 text-slate-600" />
                <h4 className="font-bold text-xs uppercase tracking-wider font-mono text-slate-700">Auto Ignore Sheet Patterns System</h4>
              </div>
              <div className="bg-slate-50/50 p-4 rounded-xl border border-slate-150 space-y-4">
                <div className="flex items-center gap-3">
                  <input
                    id="autoIgnoreEnabled"
                    type="checkbox"
                    checked={config.autoIgnoreEnabled}
                    onChange={(e) => setConfig({ ...config, autoIgnoreEnabled: e.target.checked })}
                    className="w-4 h-4 text-indigo-650 bg-white border-slate-300 rounded focus:ring-indigo-500 cursor-pointer pointer-events-auto"
                  />
                  <label htmlFor="autoIgnoreEnabled" className="text-xs text-slate-750 font-bold select-none cursor-pointer">
                    Enable Auto Ignore Rules
                  </label>
                </div>
                <div>
                  <label className="block text-slate-600 text-xs font-semibold mb-1">
                    Ignore Patterns (Comma-separated exact names, wildcard sub-strings, or `/RegEx/i`)
                  </label>
                  <input
                    type="text"
                    value={config.customIgnorePatterns}
                    disabled={!config.autoIgnoreEnabled}
                    onChange={(e) => setConfig({ ...config, customIgnorePatterns: e.target.value })}
                    className="w-full px-3 py-1.5 text-xs border border-slate-200 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-505/20 focus:border-indigo-500 transition duration-150 bg-white disabled:bg-slate-100 disabled:text-slate-400 font-mono"
                    placeholder="e.g. Cover, Index, Notes, /Metadata/i, ^Instructions$"
                  />
                  <p className="text-[10px] text-slate-400 mt-1 leading-normal">
                    Matches cover pages, instruction lists, or metadata grids and ignores them dynamically during evaluation to avoid false omit alerts.
                  </p>
                </div>
              </div>
            </div>
            
            <div className="mt-8 pt-4 border-t border-slate-100 flex items-center justify-center">
              <button
                onClick={loadDemoState}
                className="inline-flex items-center gap-2 px-6 py-3 rounded-lg text-xs font-bold bg-indigo-600 text-white hover:bg-indigo-700 transition shadow-sm cursor-pointer"
              >
                <Play className="w-4 h-4 text-white fill-white" />
                Initialize Quality Auditor with Preset Demo Dataset
              </button>
            </div>
          </div>
        </section>
      )}

      {/* 📊 Main Quality Assurance Auditor Dashboard */}
      {qaAnalysis && (
        <main className="max-w-7xl mx-auto px-4 py-8 sm:px-6 lg:px-8">
          
          {/* Metadata banner and Quick Statistics bar */}
          <div className="bg-white border border-gray-200 rounded-xl p-6 shadow-2xs mb-8 flex flex-col md:flex-row md:items-center justify-between gap-6">
            <div className="flex items-start gap-4">
              <div className="bg-slate-100 p-3 rounded-lg flex items-center justify-center">
                <User className="w-5 h-5 text-slate-700" />
              </div>
              <div>
                <div className="flex items-center gap-2 flex-wrap">
                  <h2 className="text-md font-bold text-gray-800">{config.employeeName}</h2>
                  <span className="text-3xs font-semibold bg-gray-100 text-gray-700 px-2 py-0.5 rounded-sm uppercase tracking-wider font-mono">
                    Grade: {qaAnalysis.metrics.finalGrade}
                  </span>
                </div>
                <p className="text-xs text-gray-500 mt-1">
                  Project: <strong className="text-gray-700">{config.projectName}</strong> • Inspected: <strong className="text-gray-700">{config.evaluationDate}</strong>
                </p>
                <div className="flex gap-4 mt-2 font-mono text-3xs text-gray-400 uppercase tracking-widest">
                  <span>File A: {employeeWb?.fileName}</span>
                  <span>File B: {reviewerWb?.fileName}</span>
                </div>
              </div>
            </div>

            <div className="flex gap-3 flex-wrap items-center">
              <button
                onClick={triggerAIAudit}
                disabled={aiLoading}
                className={`inline-flex items-center gap-2 px-4 py-2.5 rounded-lg text-xs font-semibold select-none transition ${
                  aiLoading 
                    ? "bg-slate-100 text-slate-400 cursor-not-allowed" 
                    : "bg-gradient-to-r from-violet-600 to-indigo-600 hover:from-violet-700 hover:to-indigo-700 text-white shadow-xs cursor-pointer"
                }`}
              >
                <Sparkles className="w-4 h-4 animate-pulse text-indigo-200" />
                {aiLoading ? "Consulting AI Auditor..." : "Generate AI Executive Audit Report"}
              </button>

              <button
                onClick={() => {
                  setEmployeeWb(null);
                  setReviewerWb(null);
                  setAiResult(null);
                  setIsComparingStarted(false);
                }}
                className="px-4 py-2.5 rounded-lg text-xs font-semibold border border-gray-300 hover:bg-gray-150 transition text-gray-700 "
              >
                Upload New Files
              </button>
            </div>
          </div>

          {/* Tab Selection Row */}
          <div className="flex items-center gap-1 border-b border-slate-200 mb-6 font-mono text-xs overflow-x-auto print:hidden">
            <button
              onClick={() => setActiveTab("dashboard")}
              className={`px-4 py-2.5 border-b-2 font-bold transition whitespace-nowrap cursor-pointer ${
                activeTab === "dashboard"
                  ? "border-indigo-600 text-indigo-700 bg-indigo-50/30 rounded-t-lg"
                  : "border-transparent text-slate-500 hover:text-slate-800 hover:bg-slate-50 rounded-t-lg"
              }`}
            >
              Dashboard Summary
            </button>
            <button
              onClick={() => setActiveTab("errorLog")}
              className={`px-4 py-2.5 border-b-2 font-bold transition whitespace-nowrap flex items-center gap-1.5 cursor-pointer ${
                activeTab === "errorLog"
                  ? "border-indigo-600 text-indigo-700 bg-indigo-50/30 rounded-t-lg"
                  : "border-transparent text-slate-500 hover:text-slate-800 hover:bg-slate-50 rounded-t-lg"
              }`}
            >
              Detailed Error Log
              <span className="bg-slate-200 text-slate-800 text-[10px] font-bold px-2 py-0.5 rounded-full font-sans">
                {totalErrorsCount}
              </span>
            </button>
            <button
              onClick={() => setActiveTab("sheetExplorer")}
              className={`px-4 py-2.5 border-b-2 font-bold transition whitespace-nowrap flex items-center gap-1.5 cursor-pointer ${
                activeTab === "sheetExplorer"
                  ? "border-indigo-600 text-indigo-700 bg-indigo-50/30 rounded-t-lg"
                  : "border-transparent text-slate-500 hover:text-slate-800 hover:bg-slate-50 rounded-t-lg"
              }`}
            >
              Grid Inspector
              <span className="bg-indigo-50 text-indigo-700 text-[10px] font-semibold px-2 py-0.5 rounded-full font-sans">
                Visual
              </span>
            </button>
            <button
              onClick={() => setActiveTab("patterns")}
              className={`px-4 py-2.5 border-b-2 font-bold transition whitespace-nowrap flex items-center gap-1.5 cursor-pointer ${
                activeTab === "patterns"
                  ? "border-indigo-600 text-indigo-705 bg-indigo-50/30 rounded-t-lg"
                  : "border-transparent text-slate-500 hover:text-slate-800 hover:bg-slate-50 rounded-t-lg"
              }`}
            >
              Shift Patterns
              {qaAnalysis.patterns.shiftEvents.length > 0 && (
                <span className="bg-amber-100 text-amber-900 text-[10px] font-bold px-1.5 py-0.5 rounded-full animate-bounce">
                  {qaAnalysis.patterns.shiftEvents.length}
                </span>
              )}
            </button>
            <button
              onClick={() => setActiveTab("aiAuditor")}
              className={`px-4 py-2.5 border-b-2 font-bold transition whitespace-nowrap flex items-center gap-1.5 cursor-pointer ${
                activeTab === "aiAuditor"
                  ? "border-indigo-600 text-indigo-700 bg-indigo-50/30 rounded-t-lg"
                  : "border-transparent text-slate-500 hover:text-slate-800 hover:bg-slate-50 rounded-t-lg"
              }`}
            >
              AI Executive Audit
              <Sparkles className="w-3.5 h-3.5 text-indigo-600 animate-pulse" />
            </button>
            <button
              onClick={() => setActiveTab("mappingExplorer")}
              className={`px-4 py-2.5 border-b-2 font-bold transition whitespace-nowrap flex items-center gap-1.5 cursor-pointer ${
                activeTab === "mappingExplorer"
                  ? "border-indigo-600 text-indigo-700 bg-indigo-50/30 rounded-t-lg"
                  : "border-transparent text-slate-500 hover:text-slate-800 hover:bg-slate-50 rounded-t-lg"
              }`}
            >
              Mapping Explorer
              <Layers className="w-3.5 h-3.5 text-indigo-500" />
            </button>
            <button
              onClick={() => setActiveTab("benchmark")}
              className={`px-4 py-2.5 border-b-2 font-bold transition whitespace-nowrap flex items-center gap-1.5 cursor-pointer ${
                activeTab === "benchmark"
                  ? "border-indigo-600 text-indigo-700 bg-indigo-50/30 rounded-t-lg"
                  : "border-transparent text-slate-500 hover:text-slate-800 hover:bg-slate-50 rounded-t-lg"
              }`}
            >
              Regressions & Benchmark
              <Activity className="w-3.5 h-3.5 text-emerald-500" />
            </button>
          </div>

          {/* TAB CONTENTS */}
          <AnimatePresence mode="wait">
            {activeTab === "dashboard" && (
              <motion.div
                key="dashboard"
                initial={{ opacity: 0, y: 5 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -5 }}
                transition={{ duration: 0.15 }}
                className="space-y-8"
              >
                
                {/* Visual Override Alerts if shifts compromise performance */}
                {qaAnalysis.patterns.shiftEvents.length > 0 && (
                  <div className="p-4 bg-amber-50 text-amber-900 border border-amber-200 rounded-xl flex items-start gap-3">
                    <AlertTriangle className="w-5 h-5 text-amber-600 flex-shrink-0 mt-0.5" />
                    <div>
                      <h4 className="font-semibold text-sm">Critical Quality Restriction Precedence Event Registered</h4>
                      <p className="text-xs text-amber-700 mt-1">
                        We isolated <strong>{qaAnalysis.patterns.shiftEvents.length} distinct structural grid alignment shift(s)</strong> in your file submission patterns. Regardless of calculated raw precision percentages, enterprise rules force-restricted the compliance grade to <strong>'Needs Improvement'</strong>.
                      </p>
                    </div>
                  </div>
                )}

                 {/* KPI Metrics Dashboard Columns */}
                 <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-4 px-6 py-5 bg-white border border-slate-200 rounded-xl shadow-xs">
                   
                   {/* Employee Performance */}
                   <div className="flex flex-col justify-between">
                     <span className="text-[10px] uppercase font-bold text-slate-400 tracking-widest mb-1">
                       Employee Performance
                     </span>
                     <div className="flex items-baseline space-x-2 mt-2">
                       <span className="text-[22px] font-black text-slate-800 italic leading-snug">
                         {config.employeeName || "Ahmed Mansour"}
                       </span>
                     </div>
                     <div className={`mt-2 px-3 py-1 text-white text-[10px] uppercase font-bold w-fit rounded ${
                       qaAnalysis.metrics.finalGrade.toLowerCase().includes("a") || 
                       qaAnalysis.metrics.finalGrade.toLowerCase().includes("outstanding") || 
                       qaAnalysis.metrics.finalGrade.toLowerCase().includes("excellent") || 
                       qaAnalysis.metrics.finalGrade.toLowerCase().includes("good")
                         ? "bg-emerald-600"
                         : "bg-red-600"
                     }`}>
                       {qaAnalysis.metrics.finalGrade}
                     </div>
                   </div>

                   {/* Accuracy */}
                   <div className="border-t sm:border-t-0 sm:border-l border-slate-100 pt-3 sm:pt-0 sm:pl-4 flex flex-col justify-between">
                     <span className="text-[10px] uppercase font-bold text-slate-400 tracking-widest mb-1">
                       Accuracy
                     </span>
                     <div className="text-3xl font-mono font-bold text-indigo-600 mt-1">
                       {qaAnalysis.metrics.baseAccuracy}%
                     </div>
                     <div className="text-[10px] text-slate-500 mt-1">
                       Standard unweighted cells accuracy
                     </div>
                   </div>

                   {/* Penalty Load */}
                   <div className="border-t md:border-t-0 md:border-l border-slate-100 pt-3 md:pt-0 md:pl-4 flex flex-col justify-between">
                     <span className="text-[10px] uppercase font-bold text-slate-400 tracking-widest mb-1">
                       Penalty Load
                     </span>
                     <div className="text-3xl font-mono font-bold text-slate-800 mt-1">
                       {qaAnalysis.metrics.totalPenaltyPoints} <span className="text-sm font-normal text-slate-400">pts</span>
                     </div>
                     <div className="text-[10px] text-slate-500 mt-1">
                       {qaAnalysis.metrics.totalErrors} Total Errors Identified
                     </div>
                   </div>

                   {/* Coverage Metrics */}
                   <div className="border-t md:border-t-0 md:border-l border-slate-100 pt-3 md:pt-0 md:pl-4 flex flex-col justify-between">
                     <span className="text-[10px] uppercase font-bold text-slate-400 tracking-widest mb-1">
                       Coverage Metrics
                     </span>
                     <div className="text-3xl font-mono font-bold text-slate-800 uppercase mt-1">
                       {qaAnalysis.metrics.comparedCells.toLocaleString()}
                     </div>
                     <div className="text-[10px] text-slate-500 mt-1">
                       Compared Cells ({Object.keys(employeeWb?.sheets || {}).length} Sheets)
                     </div>
                   </div>

                 </div>

                  {/* Unified Audit Score & Enterprise Penalties Section */}
                  <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                    {/* Scores Bento Column */}
                    <div className="lg:col-span-1 space-y-6 flex flex-col">
                      <div className="bg-slate-900 text-white rounded-xl p-6 border border-slate-950 flex-1 flex flex-col justify-between shadow-md">
                        <div>
                          <span className="text-[10px] font-mono uppercase font-bold tracking-widest text-[#00E5FF]/80">
                            Enterprise QA Score
                          </span>
                          <h3 className="text-2xl font-black tracking-tight text-white mt-1">
                            Weighted Audit
                          </h3>
                        </div>
                        
                        <div className="my-6 flex flex-col items-center">
                          <div className="text-6xl font-mono font-black text-[#00E5FF] tracking-tighter">
                            {qaAnalysis.metrics.finalAuditScore.toFixed(1)}%
                          </div>
                          <span className="text-[10px] tracking-wider uppercase text-slate-400 mt-2 font-semibold font-mono">
                            (40% STRUCTURAL + 60% DATA)
                          </span>
                        </div>

                        <div className="space-y-4 border-t border-slate-800 pt-4">
                          {/* Structural Score Item */}
                          <div>
                            <div className="flex justify-between items-center text-xs font-semibold text-slate-350">
                              <span>Structural Quality Score</span>
                              <span className="font-mono text-white">{qaAnalysis.metrics.structuralScore.toFixed(1)} / 100</span>
                            </div>
                            <div className="w-full bg-slate-800 h-1.5 rounded-full mt-1.5">
                              <div 
                                className="bg-[#00E5FF] h-1.5 rounded-full transition-all duration-500" 
                                style={{ width: `${qaAnalysis.metrics.structuralScore}%` }}
                              />
                            </div>
                          </div>

                          {/* Data Score Item */}
                          <div>
                            <div className="flex justify-between items-center text-xs font-semibold text-slate-355">
                              <span>Data Integrity Score</span>
                              <span className="font-mono text-white">{qaAnalysis.metrics.dataScore.toFixed(1)} / 100</span>
                            </div>
                            <div className="w-full bg-slate-800 h-1.5 rounded-full mt-1.5">
                              <div 
                                className="bg-emerald-400 h-1.5 rounded-full transition-all duration-500" 
                                style={{ width: `${qaAnalysis.metrics.dataScore}%` }}
                              />
                            </div>
                          </div>
                        </div>
                      </div>
                    </div>

                    {/* Audit Summary Details Bento Column */}
                    <div className="lg:col-span-2 bg-white border border-slate-200 rounded-xl p-6 shadow-sm flex flex-col justify-between">
                      <div>
                        <div className="flex items-center gap-2 mb-4 border-b border-slate-100 pb-3 font-mono">
                          <SlidersHorizontal className="w-4 h-4 text-indigo-500" />
                          <h3 className="font-bold text-xs uppercase tracking-wider text-slate-800">
                            Enterprise Audit Score Penalty Breakdown
                          </h3>
                        </div>

                        <div className="overflow-x-auto">
                          <table className="w-full text-left border-collapse text-xs">
                            <thead>
                              <tr className="border-b border-slate-200 bg-slate-50/50 text-slate-500 font-mono text-[10px] uppercase">
                                <th className="py-2.5 px-3 font-semibold">Quality Dimension / Defect Category</th>
                                <th className="py-2.5 px-3 font-semibold text-center">Identified Count</th>
                                <th className="py-2.5 px-3 font-semibold text-center">Coefficient (Weight)</th>
                                <th className="py-2.5 px-3 font-semibold text-right">Penalty Contribution</th>
                              </tr>
                            </thead>
                            <tbody className="divide-y divide-slate-150">
                              {/* Extra Tables Row */}
                              <tr className="hover:bg-slate-50/50">
                                <td className="py-2 px-3 font-medium text-slate-805">
                                  Extra Tables
                                  <span className="block text-[10px] font-normal text-slate-400">Worker workbook contains extraneous tables</span>
                                </td>
                                <td className="py-2 px-3 text-center font-mono font-bold text-slate-707">
                                  {qaAnalysis.metrics.extraTablesCount}
                                </td>
                                <td className="py-2 px-3 text-center font-mono text-slate-505">
                                  {config.extraTableCoefficient}
                                </td>
                                <td className="py-2 px-3 text-right font-mono font-bold text-rose-600">
                                  -{(qaAnalysis.metrics.extraTablesCount * config.extraTableCoefficient).toFixed(2)}
                                </td>
                              </tr>

                              {/* Missing Tables Row */}
                              <tr className="hover:bg-slate-50/50">
                                <td className="py-2 px-3 font-medium text-slate-805">
                                  Missing Tables
                                  <span className="block text-[10px] font-normal text-slate-400">Ground truth tables completely omitted in worker file</span>
                                </td>
                                <td className="py-2 px-3 text-center font-mono font-bold text-slate-707">
                                  {qaAnalysis.metrics.missingTablesCount}
                                </td>
                                <td className="py-2 px-3 text-center font-mono text-slate-555">
                                  {config.missingTableCoefficient}
                                </td>
                                <td className="py-2 px-3 text-right font-mono font-bold text-rose-600">
                                  -{(qaAnalysis.metrics.missingTablesCount * config.missingTableCoefficient).toFixed(2)}
                                </td>
                              </tr>

                              {/* Extra Columns Row */}
                              <tr className="hover:bg-slate-50/50">
                                <td className="py-2 px-3 font-medium text-slate-805">
                                  Extra Columns
                                  <span className="block text-[10px] font-normal text-slate-400">Additional layout columns present in worker survey grid</span>
                                </td>
                                <td className="py-2 px-3 text-center font-mono font-bold text-slate-707">
                                  {qaAnalysis.metrics.extraColumnsCount}
                                </td>
                                <td className="py-2 px-3 text-center font-mono text-slate-555">
                                  {config.extraColumnCoefficient}
                                </td>
                                <td className="py-2 px-3 text-right font-mono font-bold text-rose-600">
                                  -{(qaAnalysis.metrics.extraColumnsCount * config.extraColumnCoefficient).toFixed(2)}
                                </td>
                              </tr>

                              {/* Missing Columns Row */}
                              <tr className="hover:bg-slate-50/50">
                                <td className="py-2 px-3 font-medium text-slate-805">
                                  Missing Columns
                                  <span className="block text-[10px] font-normal text-slate-400">Required survey columns omitted by worker</span>
                                </td>
                                <td className="py-2 px-3 text-center font-mono font-bold text-slate-707">
                                  {qaAnalysis.metrics.missingColumnsCount}
                                </td>
                                <td className="py-2 px-3 text-center font-mono text-slate-555">
                                  {config.missingColumnCoefficient}
                                </td>
                                <td className="py-2 px-3 text-right font-mono font-bold text-rose-600">
                                  -{(qaAnalysis.metrics.missingColumnsCount * config.missingColumnCoefficient).toFixed(2)}
                                </td>
                              </tr>

                              {/* Extra Rows Row */}
                              <tr className="hover:bg-slate-50/50">
                                <td className="py-2 px-3 font-medium text-slate-805">
                                  Extra Rows
                                  <span className="block text-[10px] font-normal text-slate-400">Extraneous records inserted into the grid</span>
                                </td>
                                <td className="py-2 px-3 text-center font-mono font-bold text-slate-707">
                                  {qaAnalysis.metrics.extraRowsCount}
                                </td>
                                <td className="py-2 px-3 text-center font-mono text-slate-555">
                                  {config.extraRowCoefficient}
                                </td>
                                <td className="py-2 px-3 text-right font-mono font-bold text-rose-600">
                                  -{(qaAnalysis.metrics.extraRowsCount * config.extraRowCoefficient).toFixed(2)}
                                </td>
                              </tr>

                              {/* Missing Rows Row */}
                              <tr className="hover:bg-slate-50/50">
                                <td className="py-2 px-3 font-medium text-slate-805">
                                  Missing Rows
                                  <span className="block text-[10px] font-normal text-slate-400">Mandatory data rows totally omitted inside grid</span>
                                </td>
                                <td className="py-2 px-3 text-center font-mono font-bold text-slate-707">
                                  {qaAnalysis.metrics.missingRowsCount}
                                </td>
                                <td className="py-2 px-3 text-center font-mono text-slate-555">
                                  {config.missingRowCoefficient}
                                </td>
                                <td className="py-2 px-3 text-right font-mono font-bold text-rose-600">
                                  -{(qaAnalysis.metrics.missingRowsCount * config.missingRowCoefficient).toFixed(2)}
                                </td>
                              </tr>

                              {/* Numeric Differences Row */}
                              <tr className="hover:bg-slate-50/50">
                                <td className="py-2 px-3 font-medium text-slate-855">
                                  Numeric Differences
                                  <span className="block text-[10px] font-normal text-slate-400">Value or arithmetic divergences</span>
                                </td>
                                <td className="py-2 px-3 text-center font-mono font-bold text-slate-707">
                                  {qaAnalysis.metrics.numericDifferencesCount}
                                </td>
                                <td className="py-2 px-3 text-center font-mono text-slate-555">
                                  {config.numericDifferenceCoefficient}
                                </td>
                                <td className="py-2 px-3 text-right font-mono font-bold text-rose-600">
                                  -{(qaAnalysis.metrics.numericDifferencesCount * config.numericDifferenceCoefficient).toFixed(2)}
                                </td>
                              </tr>

                              {/* Text Differences Row */}
                              <tr className="hover:bg-slate-50/50">
                                <td className="py-2 px-3 font-medium text-slate-855">
                                  Text Differences
                                  <span className="block text-[10px] font-normal text-slate-400">Spelling variations or major descriptive changes</span>
                                </td>
                                <td className="py-2 px-3 text-center font-mono font-bold text-slate-707">
                                  {qaAnalysis.metrics.textDifferencesCount}
                                </td>
                                <td className="py-2 px-3 text-center font-mono text-slate-555">
                                  {config.textDifferenceCoefficient}
                                </td>
                                <td className="py-2 px-3 text-right font-mono font-bold text-rose-600">
                                  -{(qaAnalysis.metrics.textDifferencesCount * config.textDifferenceCoefficient).toFixed(2)}
                                </td>
                              </tr>

                              {/* Empty Cell Differences Row */}
                              <tr className="hover:bg-slate-50/50">
                                <td className="py-2 px-3 font-medium text-slate-855">
                                  Empty Cell Differences
                                  <span className="block text-[10px] font-normal text-slate-400">Individual blank vs. populated cell differences</span>
                                </td>
                                <td className="py-2 px-3 text-center font-mono font-bold text-slate-707">
                                  {qaAnalysis.metrics.emptyCellDifferencesCount}
                                </td>
                                <td className="py-2 px-3 text-center font-mono text-slate-555">
                                  {config.emptyCellDifferenceCoefficient}
                                </td>
                                <td className="py-2 px-3 text-right font-mono font-bold text-rose-600">
                                  -{(qaAnalysis.metrics.emptyCellDifferencesCount * config.emptyCellDifferenceCoefficient).toFixed(2)}
                                </td>
                              </tr>
                            </tbody>
                          </table>
                        </div>
                      </div>

                      <div className="grid grid-cols-2 sm:grid-cols-3 gap-2.5 border-t border-slate-100 pt-4 mt-4 font-mono text-[10px] bg-slate-50 p-3 rounded-lg uppercase">
                        <div>
                          <span className="text-slate-400 font-bold block">Structural Penalty</span>
                          <span className="text-slate-800 font-black text-xs text-rose-600">
                            {qaAnalysis.metrics.structuralPenalty.toFixed(2)} pts
                          </span>
                        </div>
                        <div>
                          <span className="text-slate-400 font-bold block">Data Penalty</span>
                          <span className="text-slate-800 font-black text-xs text-rose-600">
                            {qaAnalysis.metrics.dataPenalty.toFixed(2)} pts
                          </span>
                        </div>
                        <div className="col-span-2 sm:col-span-1 text-right">
                          <span className="text-slate-400 font-bold block">Cumulative Penalty</span>
                          <span className="text-indigo-600 font-black text-xs">
                            {qaAnalysis.metrics.totalPenalty.toFixed(2)} pts
                          </span>
                        </div>
                      </div>
                    </div>
                  </div>

                <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                  {/* Left Column: Root Cause Analytics bar charts */}
                  <div className="bg-white border border-slate-200 rounded-xl p-6 shadow-sm md:col-span-2">
                    <h3 className="font-bold text-xs uppercase tracking-wider font-mono text-slate-700 mb-6 flex items-center gap-1.5 border-b border-slate-100 pb-2">
                      <Activity className="w-4 h-4 text-indigo-600" />
                      Root Cause Failure Analysis
                    </h3>

                    <div className="space-y-4">
                      {/* Missing & Extraneous Values */}
                      <div>
                        <div className="flex justify-between items-center text-xs mb-1">
                          <span className="text-slate-600 font-semibold">Missing & Extraneous Values (Omissions)</span>
                          <span className="font-mono font-bold text-slate-800">{qaAnalysis.rootCause.missingValuesPct}%</span>
                        </div>
                        <div className="h-1.5 bg-slate-100 rounded-full overflow-hidden">
                          <div
                            className="h-full bg-amber-500 rounded-full"
                            style={{ width: `${qaAnalysis.rootCause.missingValuesPct}%` }}
                          />
                        </div>
                      </div>

                      {/* Numeric Digit Errors */}
                      <div>
                        <div className="flex justify-between items-center text-xs mb-1">
                          <span className="text-slate-600 font-semibold">Numeric Digit Errors (Keypad Transpositions)</span>
                          <span className="font-mono font-bold text-slate-800">{qaAnalysis.rootCause.numericErrorsPct}%</span>
                        </div>
                        <div className="h-1.5 bg-slate-100 rounded-full overflow-hidden">
                          <div
                            className="h-full bg-indigo-500 rounded-full"
                            style={{ width: `${qaAnalysis.rootCause.numericErrorsPct}%` }}
                          />
                        </div>
                      </div>

                      {/* Text Spelling Typos */}
                      <div>
                        <div className="flex justify-between items-center text-xs mb-1">
                          <span className="text-slate-600 font-semibold">Text & Arabic Spelling Typos</span>
                          <span className="font-mono font-bold text-slate-800">{qaAnalysis.rootCause.textErrorsPct}%</span>
                        </div>
                        <div className="h-1.5 bg-slate-100 rounded-full overflow-hidden">
                          <div
                            className="h-full bg-emerald-500 rounded-full"
                            style={{ width: `${qaAnalysis.rootCause.textErrorsPct}%` }}
                          />
                        </div>
                      </div>

                      {/* Range & Sequence Errors */}
                      <div>
                        <div className="flex justify-between items-center text-xs mb-1">
                          <span className="text-slate-600 font-semibold">Academic/Fiscal Year & Range Inconsistencies</span>
                          <span className="font-mono font-bold text-slate-800">{qaAnalysis.rootCause.rangeErrorsPct || 0}%</span>
                        </div>
                        <div className="h-1.5 bg-slate-100 rounded-full overflow-hidden">
                          <div
                            className="h-full bg-cyan-500 rounded-full"
                            style={{ width: `${qaAnalysis.rootCause.rangeErrorsPct || 0}%` }}
                          />
                        </div>
                      </div>

                      {/* Grid Alignment Structural Shifts */}
                      <div>
                        <div className="flex justify-between items-center text-xs mb-1">
                          <span className="text-slate-600 font-semibold">Systematic Column/Row Alignment Shifts</span>
                          <span className="font-mono font-bold text-slate-800">{qaAnalysis.rootCause.shiftErrorsPct}%</span>
                        </div>
                        <div className="h-1.5 bg-slate-100 rounded-full overflow-hidden">
                          <div
                            className="h-full bg-red-600 rounded-full"
                            style={{ width: `${qaAnalysis.rootCause.shiftErrorsPct}%` }}
                          />
                        </div>
                      </div>

                      {/* Header Interpretation Errors */}
                      <div>
                        <div className="flex justify-between items-center text-xs mb-1">
                          <span className="text-slate-600 font-semibold">Header Row Identification & Metadata Inaccuracies</span>
                          <span className="font-mono font-bold text-slate-800">{qaAnalysis.rootCause.headerErrorsPct || 0}%</span>
                        </div>
                        <div className="h-1.5 bg-slate-100 rounded-full overflow-hidden">
                          <div
                            className="h-full bg-rose-600 rounded-full"
                            style={{ width: `${qaAnalysis.rootCause.headerErrorsPct || 0}%` }}
                          />
                        </div>
                      </div>
                    </div>

                    <div className="mt-8 pt-4 border-t border-slate-150 grid grid-cols-2 gap-4 text-center">
                      <div className="p-3 bg-slate-50 border border-slate-100 rounded-lg">
                        <span className="text-[9px] font-mono text-slate-400 block uppercase font-bold tracking-wider">Error Rate (Per 10k Cells)</span>
                        <strong className="text-lg font-bold text-slate-800 mt-1 block font-mono">
                          {qaAnalysis.metrics.errorRatePer10k.toLocaleString()}
                        </strong>
                      </div>
                      <div className="p-3 bg-slate-50 border border-slate-100 rounded-lg flex flex-col justify-center items-center">
                        <span className="text-[9px] font-mono text-slate-400 block uppercase font-bold tracking-wider mb-1">Reviewer Workload Burden</span>
                        <strong className="text-xs font-bold text-slate-800 font-sans leading-tight">
                          Estimated {qaAnalysis.metrics.reviewerWorkloadIndex.toFixed(2)} hours of reviewer correction work generated
                        </strong>
                      </div>
                    </div>
                  </div>

                  {/* Right Column: Key Diagnostic Overview info */}
                  <div className="bg-white border border-slate-200 rounded-xl p-6 shadow-sm flex flex-col justify-between">
                    <div>
                      <h3 className="font-bold text-xs uppercase tracking-wider font-mono text-slate-700 mb-4 flex items-center gap-1.5 border-b border-slate-100 pb-2">
                        <BookOpen className="w-4 h-4 text-indigo-600" />
                        Executive Accuracy Rubric
                      </h3>
                      <p className="text-xs text-slate-500 mb-4 leading-relaxed font-sans">
                        Data accuracy tier brackets based strictly on compiled enterprise quality assurance compliance standards:
                      </p>

                      <div className="space-y-1.5 font-mono text-2xs">
                        <div className="flex justify-between items-center py-1 border-b border-slate-50">
                          <span className="text-emerald-700 font-bold uppercase tracking-wider">Outstanding</span>
                          <span className="font-bold">99.90%+</span>
                        </div>
                        <div className="flex justify-between items-center py-1 border-b border-slate-50">
                          <span className="text-emerald-600 font-semibold uppercase tracking-wider">Excellent</span>
                          <span className="font-medium">99.00% – 99.89%</span>
                        </div>
                        <div className="flex justify-between items-center py-1 border-b border-slate-50">
                          <span className="text-teal-600 font-semibold uppercase tracking-wider">Very Good</span>
                          <span className="font-medium">97.00% – 98.99%</span>
                        </div>
                        <div className="flex justify-between items-center py-1 border-b border-slate-50">
                          <span className="text-blue-600 font-semibold uppercase tracking-wider">Good</span>
                          <span className="font-medium">95.00% – 96.99%</span>
                        </div>
                        <div className="flex justify-between items-center py-1 border-b border-slate-50">
                          <span className="text-amber-600 font-semibold uppercase tracking-wider">Fair</span>
                          <span className="font-medium">90.00% – 94.99%</span>
                        </div>
                        <div className="flex justify-between items-center py-1 border-b border-slate-50">
                          <span className="text-red-500 font-semibold uppercase tracking-wider">Needs Care</span>
                          <span className="font-medium">80.00% – 89.99%</span>
                        </div>
                        <div className="flex justify-between items-center py-1">
                          <span className="text-red-700 font-bold uppercase tracking-wider">Unacceptable</span>
                          <span className="font-bold">Below 80%</span>
                        </div>
                      </div>
                    </div>

                    <div className="mt-6 pt-4 border-t border-slate-100 bg-indigo-50/50 p-4 rounded-lg border border-indigo-100 text-xs">
                      <div className="font-bold text-indigo-800 flex items-center gap-1.5 mb-1.5 font-mono text-[10px] uppercase tracking-wider">
                        <CheckCircle className="w-4 h-4 text-indigo-600" />
                        Pre-Normalized Parsing
                      </div>
                      <p className="text-[11px] text-indigo-750 font-sans leading-relaxed">
                        Leading and trailing whitespace, duplicated spaces, diacritics, and Te Marbuta variants are dynamically cleansed prior to alignment shift and transposition checks.
                      </p>
                    </div>
                  </div>
                </div>

                {/* Coaching Recommendations */}
                <div className="bg-slate-800 rounded-xl p-6 text-white shadow-lg border border-slate-705">
                  <h3 className="font-bold text-xs uppercase tracking-wider font-mono text-indigo-305 mb-5 flex items-center gap-2 border-b border-slate-700 pb-2">
                    <SlidersHorizontal className="w-4 h-4 text-indigo-400" />
                    Actionable Employee Coaching Recommendations
                  </h3>

                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    {qaAnalysis.coachingRecommendations.map((recommendation, idx) => (
                      <div key={idx} className="flex gap-2 items-start text-[11px] leading-relaxed">
                        <span className="text-indigo-400 font-extrabold font-mono text-xs">{String(idx + 1).padStart(2, '0')}.</span>
                        <span className="text-slate-150 font-sans">{recommendation}</span>
                      </div>
                    ))}
                  </div>
                </div>

                {/* 📊 Google Sheets Connection & Report Saving Panel */}
                <div className="bg-white border border-slate-200 rounded-xl p-6 shadow-sm mt-6">
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-5 border-b border-slate-150 pb-3 flex-wrap">
                    <div className="flex items-center gap-2">
                      <div className="p-2 bg-emerald-50 text-emerald-600 rounded">
                        <FileSpreadsheet className="w-5 h-5 flex-shrink-0 animate-pulse-slow" />
                      </div>
                      <div>
                        <h3 className="font-extrabold text-sm text-slate-800">Connect & Save to Google Sheets</h3>
                        <p className="text-2xs text-slate-500 font-sans mt-0.5">Maintain a running log of employee audit reports directly in your cloud workspace spreadsheets.</p>
                      </div>
                    </div>

                    <div className="flex items-center gap-2 flex-wrap">
                      {googleAccessToken ? (
                        <div className="flex items-center gap-3 bg-slate-50 border border-slate-200 px-3 py-1.5 rounded-lg text-2xs">
                          <div className="w-2.5 h-2.5 bg-emerald-500 rounded-full animate-pulse"></div>
                          <div>
                            <span className="font-semibold text-slate-700 block max-w-[150px] truncate">{sheetsUser?.name || "Connected User"}</span>
                            <span className="text-slate-400 block max-w-[150px] truncate">{sheetsUser?.email || "Google Connected"}</span>
                          </div>
                          <button
                            onClick={handleGoogleSignOut}
                            className="bg-white hover:bg-slate-100 text-slate-650 hover:text-red-500 border border-slate-200 ml-1 px-2 py-0.5 rounded transition font-bold"
                          >
                            Sign Out
                          </button>
                        </div>
                      ) : (
                        <button
                          onClick={handleGoogleSignIn}
                          className="gsi-material-button inline-flex items-center gap-2 px-3.5 py-1.5 bg-white hover:bg-slate-50 text-slate-755 border border-slate-200 rounded-lg cursor-pointer text-xs font-bold transition shadow-xs"
                        >
                          <svg version="1.1" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 48 48" className="w-4 h-4 flex-shrink-0">
                            <path fill="#EA4335" d="M24 9.5c3.54 0 6.71 1.22 9.21 3.6l6.85-6.85C35.9 2.38 30.47 0 24 0 14.62 0 6.51 5.38 2.56 13.22l7.98 6.19C12.43 13.72 17.74 9.5 24 9.5z"></path>
                            <path fill="#4285F4" d="M46.98 24.55c0-1.57-.15-3.09-.38-4.55H24v9.02h12.94c-.58 2.96-2.26 5.48-4.78 7.18l7.73 6c4.51-4.18 7.09-10.36 7.09-17.65z"></path>
                            <path fill="#FBBC05" d="M10.53 28.59c-.48-1.45-.76-2.99-.76-4.59s.27-3.14.76-4.59l-7.98-6.19C.92 16.46 0 20.12 0 24c0 3.88.92 7.54 2.56 10.78l7.97-6.19z"></path>
                            <path fill="#34A853" d="M24 48c6.48 0 11.93-2.13 15.89-5.81l-7.73-6c-2.15 1.45-4.92 2.3-8.16 2.3-6.26 0-11.57-4.22-13.47-9.91l-7.98 6.19C6.51 42.62 14.62 48 24 48z"></path>
                          </svg>
                          Sign In with Google
                        </button>
                      )}
                    </div>
                  </div>

                  {!isEmployeeNameValid && (
                    <div className="mb-4 p-3 bg-amber-50 text-amber-950 border border-amber-200 rounded-lg flex items-start gap-2.5 text-2xs">
                      <AlertTriangle className="w-4 h-4 text-amber-605 flex-shrink-0 mt-0.5" />
                      <div>
                        <strong className="font-bold">Google Sheets Log Disabled</strong>
                        <p className="text-amber-800 mt-0.5">The employee name field remains set to the default placeholder "Employee Name". Please type in the employee's genuine name under the Threshold Configurations card above to activate cloud saving capabilities.</p>
                      </div>
                    </div>
                  )}

                  <div className="grid grid-cols-1 md:grid-cols-2 gap-6 text-xs mt-2">
                    <div>
                      <label className="block text-slate-650 font-bold mb-1.5">Target Spreadsheet (Google Sheets URL or ID)</label>
                      <div className="flex gap-2">
                        <input
                          type="text"
                          value={spreadsheetId}
                          onChange={(e) => handleSpreadsheetIdChange(e.target.value)}
                          placeholder="e.g. 1a2b3c4d... or paste Spreadsheet link"
                          className="w-full px-3 py-2 border border-slate-200 rounded-md text-xs font-mono focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 transition duration-150 bg-slate-50/50"
                        />
                        {googleAccessToken && (
                          <button
                            type="button"
                            onClick={handleCreateNewSpreadsheet}
                            disabled={sheetsLoading}
                            className="bg-slate-100 hover:bg-slate-200 border border-slate-250 text-slate-750 font-semibold px-3 py-2 rounded-md hover:border-slate-350 transition duration-150 flex items-center gap-1.5 shrink-0 cursor-pointer disabled:opacity-50 text-xs"
                          >
                            {sheetsLoading ? <RefreshCw className="w-3.5 h-3.5 animate-spin" /> : <Plus className="w-3.5 h-3.5 text-slate-500" />}
                            Create New
                          </button>
                        )}
                      </div>
                      <p className="text-[10px] text-slate-450 font-sans mt-1.5 leading-relaxed">
                        Specify where evaluation logs are appended. All actions append formatted rows under a new tab named <code className="bg-slate-50 border border-slate-200 px-1 rounded text-2xs font-mono font-semibold">Evaluation Summary Logs</code>.
                      </p>
                    </div>

                    <div className="flex flex-col justify-end">
                      <div className="flex flex-col sm:flex-row gap-3">
                        <button
                          type="button"
                          onClick={handleSaveReportToSheets}
                          disabled={sheetsLoading || !googleAccessToken || !spreadsheetId.trim() || !isEmployeeNameValid}
                          className={`w-full py-2.5 rounded-lg border flex items-center justify-center gap-2 font-bold select-none cursor-pointer text-xs shadow-xs transition duration-150 ${
                            sheetsLoading || !googleAccessToken || !spreadsheetId.trim() || !isEmployeeNameValid
                              ? "bg-slate-100 text-slate-400 border-slate-200 cursor-not-allowed"
                              : "bg-emerald-600 hover:bg-emerald-700 text-white border-emerald-700 hover:border-emerald-800"
                          }`}
                        >
                          {sheetsLoading ? (
                            <RefreshCw className="w-4 h-4 animate-spin" />
                          ) : (
                            <CheckCircle className="w-4 h-4" />
                          )}
                          Save Active Compliance Report
                        </button>
                        {spreadsheetId.trim() && (
                          <a
                            href={`https://docs.google.com/spreadsheets/d/${spreadsheetId}`}
                            target="_blank"
                            rel="noopener noreferrer"
                            className="inline-flex items-center justify-center gap-1.5 px-4 py-2.5 bg-slate-50 hover:bg-slate-100 text-slate-700 border border-slate-205 rounded-lg text-xs font-semibold select-none transition"
                          >
                            Open Sheet
                            <ExternalLink className="w-3.5 h-3.5 text-slate-505" />
                          </a>
                        )}
                      </div>

                      {sheetsMessage && (
                        <div className={`mt-3 p-3 rounded-lg border text-2xs leading-snug font-medium flex items-start gap-2 ${
                          sheetsMessage.type === "success"
                            ? "bg-emerald-50 text-emerald-850 border-emerald-200"
                            : sheetsMessage.type === "error"
                              ? "bg-red-50 text-red-850 border-red-200"
                              : "bg-indigo-50 text-indigo-900 border-indigo-200"
                        }`}>
                          <div className={`w-1.5 h-1.5 rounded-full mt-1.5 shrink-0 ${
                            sheetsMessage.type === "success"
                              ? "bg-emerald-500"
                              : sheetsMessage.type === "error"
                                ? "bg-red-500"
                                : "bg-indigo-500"
                          }`} />
                          <span className="flex-1">{sheetsMessage.text}</span>
                        </div>
                      )}
                    </div>
                  </div>
                </div>

              </motion.div>
            )}

            {activeTab === "errorLog" && (
              <motion.div
                key="errorLog"
                initial={{ opacity: 0, y: 5 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -5 }}
                transition={{ duration: 0.15 }}
                className="space-y-6"
              >
                
                {/* Visual Truncation Disclaimer */}
                {isTruncated && (
                  <div className="p-4 bg-slate-900 text-white rounded-xl border border-slate-700 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                    <div className="flex items-start gap-3">
                      <AlertCircle className="w-5 h-5 text-emerald-400 flex-shrink-0 mt-0.5" />
                      <div>
                        <h4 className="font-semibold text-sm">Output Truncation Rule Triggered</h4>
                        <p className="text-xs text-slate-400 mt-1">
                          [Truncated: Total of <strong>{qaAnalysis.errorLog.length} severe mismatch coordinates found</strong>. Displaying only the Top 50 severe errors below for brevity, available in raw data stream.]
                        </p>
                      </div>
                    </div>
                    <button
                      onClick={() => setForceShowAllErrors(true)}
                      className="px-3.5 py-1.5 bg-slate-800 text-emerald-400 border border-slate-700 rounded-lg text-2xs font-semibold hover:bg-slate-700 transition flex-shrink-0 cursor-pointer"
                    >
                      Bypass Truncation (Show All)
                    </button>
                  </div>
                )}

                {/* Filter and Search Bar controls */}
                <div className="bg-white border border-slate-200 rounded-xl p-4 shadow-sm flex flex-wrap items-center justify-between gap-4">
                  
                  {/* Left: Input Search */}
                  <div className="flex items-center gap-3 flex-1 min-w-[240px]">
                    <div className="relative w-full">
                      <input
                        type="text"
                        placeholder="Search coordinates, values, cell error logs..."
                        value={logSearch}
                        onChange={(e) => setLogSearch(e.target.value)}
                        className="w-full pl-9 pr-3 py-1.5 bg-white border border-slate-250 rounded-lg text-xs font-sans text-slate-800 placeholder-slate-400 focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 transition duration-150"
                      />
                      <span className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400">
                        <Filter className="w-3.5 h-3.5" />
                      </span>
                    </div>
                  </div>

                  {/* Right: Select Drops */}
                  <div className="flex items-center gap-3 text-xs w-full sm:w-auto overflow-x-auto py-1">
                    <div>
                      <select
                        value={logFilterSeverity}
                        onChange={(e) => setLogFilterSeverity(e.target.value)}
                        className="px-3 py-1.5 bg-white border border-slate-250 rounded-lg text-[10px] font-bold font-sans text-slate-700 focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-550 transition duration-150 cursor-pointer"
                      >
                        <option value="ALL">ALL SEVERITIES</option>
                        <option value="CRITICAL">CRITICAL</option>
                        <option value="HIGH">HIGH</option>
                        <option value="MEDIUM">MEDIUM</option>
                        <option value="LOW">LOW</option>
                      </select>
                    </div>

                    <div>
                      <select
                        value={logFilterType}
                        onChange={(e) => setLogFilterType(e.target.value)}
                        className="px-3 py-1.5 bg-white border border-slate-250 rounded-lg text-[10px] font-bold font-sans text-slate-700 truncate max-w-[140px] focus:outline-none focus:ring-2 focus:ring-indigo-550/20 focus:border-indigo-550 transition duration-150 cursor-pointer"
                      >
                        <option value="ALL">ALL ERROR TYPES</option>
                        {Object.values(ErrorType).map((typ, idx) => (
                          <option key={idx} value={typ}>
                            {typ}
                          </option>
                        ))}
                      </select>
                    </div>

                    <div>
                      <select
                        value={logSortField}
                        onChange={(e) => setLogSortField(e.target.value as any)}
                        className="px-3 py-1.5 bg-white border border-slate-250 rounded-lg text-[10px] font-bold font-sans text-slate-700 focus:outline-none focus:ring-2 focus:ring-indigo-550/20 focus:border-indigo-550 transition duration-150 cursor-pointer"
                      >
                        <option value="severity">SORT BY SEVERITY</option>
                        <option value="penalty">SORT BY PENALTY</option>
                        <option value="sheet">SORT BY SHEET</option>
                        <option value="cell">SORT BY COORD</option>
                      </select>
                    </div>
                  </div>
                </div>

                {/* Table list */}
                <div className="bg-white border border-slate-200 rounded-xl overflow-hidden shadow-xs">
                  <div className="overflow-x-auto">
                    <table className="w-full text-left border-collapse text-xs">
                      <thead>
                        <tr className="bg-slate-50/70 border-b border-slate-200 text-[10px] font-bold text-slate-400 uppercase tracking-widest select-none">
                          <th className="py-3 px-4">Sheet Name</th>
                          <th className="py-3 px-4">Cell</th>
                          <th className="py-3 px-4">Employee Value</th>
                          <th className="py-3 px-4">Reviewer Value</th>
                          <th className="py-3 px-4">Error Category</th>
                          <th className="py-3 px-4">Sev</th>
                          <th className="py-3 px-4 text-center">Weight</th>
                          <th className="py-3 px-4">Diagnostic Audit Note</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100 font-sans text-slate-700">
                        {displayLogs.length > 0 ? (
                          displayLogs.map((log, idx) => {
                            const sevClass = {
                              [Severity.Critical]: "bg-red-50 text-red-700 border-red-150",
                              [Severity.High]: "bg-orange-50 text-orange-850 border-orange-150",
                              [Severity.Medium]: "bg-amber-50 text-amber-800 border-amber-150",
                              [Severity.Low]: "bg-slate-50 text-slate-600 border-slate-150"
                            }[log.severity];

                            const errorCategoryMark = {
                              [ErrorType.RowShift]: "text-red-700 bg-red-50/40 border-red-100",
                               [ErrorType.MissingRow]: "text-pink-700 bg-pink-50/40 border-pink-100",
                               [ErrorType.ExtraRow]: "text-indigo-700 bg-indigo-50/40 border-indigo-100",
                              [ErrorType.ColumnShift]: "text-red-700 bg-red-50/40 border-red-100",
                              [ErrorType.MissingValue]: "text-amber-700 bg-amber-50/40 border-amber-100",
                              [ErrorType.ExtraValue]: "text-sky-700 bg-sky-50/40 border-sky-100",
                              [ErrorType.MajorNumericError]: "text-orange-750 bg-orange-50/40 border-orange-100",
                              [ErrorType.TextTypo]: "text-emerald-700 bg-emerald-50/40 border-emerald-100",
                              [ErrorType.TableMerge]: "text-purple-700 bg-purple-50/45 border-purple-100",
                              [ErrorType.TableSplit]: "text-fuchsia-700 bg-fuchsia-50/45 border-fuchsia-100"
                            }[log.errorType] || "text-slate-750 bg-slate-50 border-slate-150";

                            return (
                              <tr key={idx} className="hover:bg-slate-50/70 border-b border-slate-50/60 transition text-[11px]">
                                <td className="py-3 px-4 font-semibold text-slate-600 truncate max-w-[120px]">{log.sheet}</td>
                                <td className="py-3 px-4 font-mono font-bold text-indigo-700">{log.cell}</td>
                                
                                <td className="py-3 px-4 font-mono text-red-650">
                                  {log.employeeValue === "" ? (
                                    <span className="text-slate-400 italic font-sans text-3xs">empty</span>
                                  ) : (
                                    log.employeeValue
                                  )}
                                </td>
                                
                                <td className="py-3 px-4 font-mono text-emerald-750">
                                  {log.reviewerValue === "" ? (
                                    <span className="text-slate-400 italic font-sans text-3xs">empty</span>
                                  ) : (
                                    log.reviewerValue
                                  )}
                                </td>

                                <td className="py-3 px-4">
                                  <span className={`px-2 py-0.5 rounded text-[9px] font-bold tracking-wider border uppercase font-mono ${errorCategoryMark}`}>
                                    {log.errorType}
                                  </span>
                                </td>

                                <td className="py-3 px-4">
                                  <span className={`px-2 py-0.5 rounded border text-[9px] font-bold uppercase font-sans ${sevClass}`}>
                                    {log.severity}
                                  </span>
                                </td>

                                <td className="py-3 px-4 text-center font-mono font-bold text-slate-800">{log.penalty}</td>
                                <td className="py-3 px-4 text-slate-500 max-w-[300px] leading-relaxed font-sans">{log.notes}</td>
                              </tr>
                            );
                          })
                        ) : (
                          <tr>
                            <td colSpan={8} className="py-12 text-center text-gray-400 italic font-sans">
                              No discrepancies found matching the selected search parameters. All coordinates conform!
                            </td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>
              </motion.div>
            )}

            {activeTab === "sheetExplorer" && (
              <motion.div
                key="sheetExplorer"
                initial={{ opacity: 0, y: 5 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -5 }}
                transition={{ duration: 0.15 }}
                className="space-y-6"
              >
                {/* Sheet Tabs explorer selection */}
                <div className="flex flex-wrap items-center justify-between gap-4 border-b border-slate-200 pb-4">
                  <div className="flex items-center gap-2">
                    <Table className="w-4 h-4 text-slate-700" />
                    <span className="text-[10px] font-mono uppercase tracking-widest text-slate-550 font-bold">
                      Grid Coordinate Selector:
                    </span>
                    <div className="flex gap-1 overflow-x-auto max-w-[400px]">
                      {Object.keys(employeeWb?.sheets || {}).map((sName, index) => (
                        <button
                          key={index}
                          onClick={() => {
                            setSelectedSheetExplorer(sName);
                            setActiveErrorFilter("ALL");
                          }}
                          className={`px-3 py-1 text-[10px] font-bold rounded font-sans transition pointer-events-auto cursor-pointer ${
                            selectedSheetExplorer === sName
                              ? "bg-indigo-600 text-white"
                              : "bg-slate-100 hover:bg-slate-200/80 text-slate-650"
                          }`}
                        >
                          {sName}
                        </button>
                      ))}
                    </div>
                  </div>

                  <div className="flex flex-wrap items-center gap-3">
                    <span className="text-[10px] font-mono uppercase text-slate-400 font-bold tracking-wider">View Options:</span>
                    <button
                      onClick={() => setExplorerViewMode("sideBySide")}
                      className={`px-3 py-1 text-[10px] font-bold rounded transition pointer-events-auto cursor-pointer ${
                        explorerViewMode === "sideBySide" ? "bg-indigo-50 text-indigo-700 border border-indigo-200" : "bg-slate-100 text-slate-500 hover:bg-slate-200/80"
                      }`}
                    >
                      Side by Side Comparison
                    </button>
                    <button
                      onClick={() => setExplorerViewMode("grid")}
                      className={`px-3 py-1 text-[10px] font-bold rounded transition pointer-events-auto cursor-pointer ${
                        explorerViewMode === "grid" ? "bg-indigo-50 text-indigo-700 border border-indigo-200" : "bg-slate-100 text-slate-500 hover:bg-slate-200/80"
                      }`}
                    >
                      Coord Overlay Audit Grid
                    </button>

                    {/* Interactive Heatmap Overlay Controls */}
                    <div className="flex items-center gap-4 bg-slate-50 border border-slate-200/70 p-1 px-3 rounded-lg text-xs">
                      <div className="flex items-center gap-2">
                        <input
                          id="showHeatmapCheck"
                          type="checkbox"
                          checked={showHeatmap}
                          onChange={(e) => setShowHeatmap(e.target.checked)}
                          className="w-3.5 h-3.5 text-indigo-650 bg-white border-slate-300 rounded focus:ring-indigo-505 cursor-pointer pointer-events-auto"
                        />
                        <label htmlFor="showHeatmapCheck" className="text-[10px] font-mono text-slate-700 uppercase font-bold select-none cursor-pointer">
                          Heatmap Overlay
                        </label>
                      </div>

                      {showHeatmap && (
                        <div className="flex items-center gap-1.5">
                          <span className="text-[9px] font-mono text-slate-400 font-bold uppercase">Theme:</span>
                          <select
                            value={heatmapTheme}
                            onChange={(e) => setHeatmapTheme(e.target.value as any)}
                            className="bg-white border border-slate-250 py-0.5 px-2 rounded text-[10px] font-mono font-bold text-slate-750 focus:outline-none focus:ring-1 focus:ring-indigo-500 pointer-events-auto"
                          >
                            <option value="coral">Classic Coral</option>
                            <option value="ruby">Volcanic Ruby</option>
                            <option value="emerald">Emerald Field</option>
                          </select>
                        </div>
                      )}
                    </div>
                  </div>
                </div>

                {/* VISUAL LAYOUT EXPLORATION GRID */}
                {selectedSheetExplorer && employeeWb?.sheets[selectedSheetExplorer] && (
                  <div className="bg-white border border-slate-200 rounded-xl p-6 shadow-sm overflow-hidden font-sans">
                    <div className="mb-4">
                      <h3 className="font-bold text-xs font-mono uppercase tracking-widest text-slate-750 mb-1">
                        Workbook Structural Map Sheet: '{selectedSheetExplorer}'
                      </h3>
                      <p className="text-[10px] text-slate-400 leading-relaxed font-mono uppercase tracking-wider flex gap-4 select-none pb-2 border-b border-slate-100 flex-wrap">
                        <span>● MATCH (White)</span>
                        <span className="text-red-500">● DISCREPANCY MISMATCH (Red)</span>
                        <span className="text-amber-600">● ALIGNMENT SHIFT EVENT (Amber)</span>
                        {visualGridBounds.realMaxRow > visualGridBounds.maxRow && (
                          <span className="text-indigo-600 bg-indigo-50 px-1.5 py-0.5 rounded font-bold normal-case">
                            ● Grid display cropped to 150 Rows / 35 Columns for smooth performance
                          </span>
                        )}
                      </p>
                    </div>
                    {/* Grid Inspector Error Summary Panel */}
                    {(() => {
                      const breakdown = computeSheetErrorsBreakdown(selectedSheetExplorer);
                      return (
                        <div className="mb-6 bg-slate-50/50 border border-slate-200 p-4 rounded-xl shadow-2xs">
                          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 mb-4 border-b border-slate-200 pb-3">
                            <div>
                              <span className="text-[10px] font-mono uppercase tracking-widest text-slate-400 font-bold">
                                Evaluation Statistics
                              </span>
                              <h4 className="text-sm font-bold text-slate-800 flex items-center gap-1.5 mt-0.5">
                                <span>Table: {selectedSheetExplorer} Summary</span>
                                {breakdown.total === 0 ? (
                                  <span className="bg-emerald-100 text-emerald-800 text-[10px] font-extrabold px-2.5 py-0.5 rounded-full select-none animate-pulse">
                                    ✔ No Errors Found
                                  </span>
                                ) : (
                                  <span className="bg-red-100 text-red-800 text-[10px] font-extrabold px-2.5 py-0.5 rounded-full select-none">
                                    {breakdown.total} Total Errors
                                  </span>
                                )}
                              </h4>
                            </div>
                            <div className="text-[11px] text-slate-500 leading-normal font-sans bg-amber-50 border border-amber-100 p-1.5 px-3 rounded-lg max-w-sm">
                              <span className="font-semibold text-amber-900">Pro-tip:</span> Click any category card below to filter the visual grid and display only those specific discrepancies.
                            </div>
                          </div>

                          {breakdown.total === 0 ? (
                            <div className="p-4 bg-emerald-50/80 rounded-lg border border-emerald-200/60 flex items-center gap-3">
                              <div className="w-8 h-8 rounded-full bg-emerald-100 flex items-center justify-center text-emerald-700 flex-shrink-0">
                                <svg className="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
                                </svg>
                              </div>
                              <div>
                                <p className="text-xs font-bold text-emerald-950">No Errors Found</p>
                                <p className="text-[11px] text-emerald-750 mt-0.5">This sheet has been checked perfectly. All entered coordinates correspond exactly to the ground truth ledger!</p>
                              </div>
                            </div>
                          ) : (
                            <div className="grid grid-cols-2 md:grid-cols-7 gap-2">
                              {/* Total Errors Card */}
                              <button
                                onClick={() => setActiveErrorFilter("ALL")}
                                className={`flex flex-col items-center justify-center p-3 rounded-xl border text-center transition duration-150 cursor-pointer pointer-events-auto ${
                                  activeErrorFilter === "ALL"
                                    ? "bg-indigo-600 border-indigo-600 text-white shadow-sm ring-2 ring-indigo-200"
                                    : "bg-white border-slate-200 hover:bg-slate-50 text-slate-700"
                                }`}
                              >
                                <span className={`text-[10px] uppercase font-bold tracking-wider ${activeErrorFilter === "ALL" ? "text-indigo-100" : "text-slate-400"}`}>
                                  Total
                                </span>
                                <span className="text-xl font-bold font-mono mt-1">
                                  {breakdown.total}
                                </span>
                              </button>

                              {/* Structural Errors Card */}
                              <button
                                onClick={() => setActiveErrorFilter(activeErrorFilter === "STRUCTURAL" ? "ALL" : "STRUCTURAL")}
                                className={`flex flex-col items-center justify-center p-3 rounded-xl border text-center transition duration-150 cursor-pointer pointer-events-auto ${
                                  activeErrorFilter === "STRUCTURAL"
                                    ? "bg-indigo-600 border-indigo-600 text-white shadow-sm ring-2 ring-indigo-200"
                                    : "bg-white border-slate-200 hover:bg-slate-50 text-slate-700"
                                }`}
                              >
                                <span className={`text-[10px] uppercase font-bold tracking-wider ${activeErrorFilter === "STRUCTURAL" ? "text-indigo-100" : "text-slate-400"}`}>
                                  Structural
                                </span>
                                <span className="text-xl font-bold font-mono mt-1">
                                  {breakdown.structural}
                                </span>
                              </button>

                              {/* Shift Errors Card */}
                              <button
                                onClick={() => setActiveErrorFilter(activeErrorFilter === "SHIFT" ? "ALL" : "SHIFT")}
                                className={`flex flex-col items-center justify-center p-3 rounded-xl border text-center transition duration-150 cursor-pointer pointer-events-auto ${
                                  activeErrorFilter === "SHIFT"
                                    ? "bg-indigo-600 border-indigo-600 text-white shadow-sm ring-2 ring-indigo-200"
                                    : "bg-white border-slate-200 hover:bg-slate-50 text-slate-700"
                                }`}
                              >
                                <span className={`text-[10px] uppercase font-bold tracking-wider ${activeErrorFilter === "SHIFT" ? "text-indigo-100" : "text-slate-400"}`}>
                                  Shift
                                </span>
                                <span className="text-xl font-bold font-mono mt-1">
                                  {breakdown.shift}
                                </span>
                              </button>

                              {/* Range Errors Card */}
                              <button
                                onClick={() => setActiveErrorFilter(activeErrorFilter === "RANGE" ? "ALL" : "RANGE")}
                                className={`flex flex-col items-center justify-center p-3 rounded-xl border text-center transition duration-150 cursor-pointer pointer-events-auto ${
                                  activeErrorFilter === "RANGE"
                                    ? "bg-indigo-600 border-indigo-600 text-white shadow-sm ring-2 ring-indigo-200"
                                    : "bg-white border-slate-200 hover:bg-slate-50 text-slate-700"
                                }`}
                              >
                                <span className={`text-[10px] uppercase font-bold tracking-wider ${activeErrorFilter === "RANGE" ? "text-indigo-100" : "text-slate-400"}`}>
                                  Range
                                </span>
                                <span className="text-xl font-bold font-mono mt-1">
                                  {breakdown.range}
                                </span>
                              </button>

                              {/* Numeric Errors Card */}
                              <button
                                onClick={() => setActiveErrorFilter(activeErrorFilter === "NUMERIC" ? "ALL" : "NUMERIC")}
                                className={`flex flex-col items-center justify-center p-3 rounded-xl border text-center transition duration-150 cursor-pointer pointer-events-auto ${
                                  activeErrorFilter === "NUMERIC"
                                    ? "bg-indigo-600 border-indigo-600 text-white shadow-sm ring-2 ring-indigo-200"
                                    : "bg-white border-slate-200 hover:bg-slate-50 text-slate-700"
                                }`}
                              >
                                <span className={`text-[10px] uppercase font-bold tracking-wider ${activeErrorFilter === "NUMERIC" ? "text-indigo-100" : "text-slate-400"}`}>
                                  Numeric
                                </span>
                                <span className="text-xl font-bold font-mono mt-1">
                                  {breakdown.numeric}
                                </span>
                              </button>

                              {/* Text Errors Card */}
                              <button
                                onClick={() => setActiveErrorFilter(activeErrorFilter === "TEXT" ? "ALL" : "TEXT")}
                                className={`flex flex-col items-center justify-center p-3 rounded-xl border text-center transition duration-150 cursor-pointer pointer-events-auto ${
                                  activeErrorFilter === "TEXT"
                                    ? "bg-indigo-600 border-indigo-600 text-white shadow-sm ring-2 ring-indigo-200"
                                    : "bg-white border-slate-200 hover:bg-slate-50 text-slate-700"
                                }`}
                              >
                                <span className={`text-[10px] uppercase font-bold tracking-wider ${activeErrorFilter === "TEXT" ? "text-indigo-100" : "text-slate-400"}`}>
                                  Text
                                </span>
                                <span className="text-xl font-bold font-mono mt-1">
                                  {breakdown.text}
                                </span>
                              </button>

                              {/* Header Errors Card */}
                              <button
                                onClick={() => setActiveErrorFilter(activeErrorFilter === "HEADER" ? "ALL" : "HEADER")}
                                className={`flex flex-col items-center justify-center p-3 rounded-xl border text-center transition duration-150 cursor-pointer pointer-events-auto ${
                                  activeErrorFilter === "HEADER"
                                    ? "bg-indigo-600 border-indigo-600 text-white shadow-sm ring-2 ring-indigo-200"
                                    : "bg-white border-slate-200 hover:bg-slate-50 text-slate-700"
                                }`}
                              >
                                <span className={`text-[10px] uppercase font-bold tracking-wider ${activeErrorFilter === "HEADER" ? "text-indigo-100" : "text-slate-400"}`}>
                                  Header
                                </span>
                                <span className="text-xl font-bold font-mono mt-1">
                                  {breakdown.header}
                                </span>
                              </button>
                            </div>
                          )}
                        </div>
                      );
                    })()}

                    <div className="space-y-4">
                      {explorerViewMode === "sideBySide" ? (
                        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 p-4 bg-slate-50/55 rounded-xl border border-slate-200">
                          
                          {/* File A Grid */}
                          <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs flex flex-col min-w-0">
                            <h4 className="text-[11px] uppercase tracking-wider font-bold font-mono text-slate-600 mb-3 border-b border-slate-100 pb-2 flex items-center justify-between flex-wrap gap-2">
                              <span className="flex items-center gap-1.5">
                                <span className="w-2 h-2 rounded-full bg-indigo-500"></span>
                                File A: Worker Submission
                              </span>
                              <span className="text-3xs font-mono font-normal normal-case text-slate-500">
                                Scale: {(employeeWb?.sheets[selectedSheetExplorer]?.maxRow ?? 0) + 1}R x {(employeeWb?.sheets[selectedSheetExplorer]?.maxCol ?? 0) + 1}C
                              </span>
                            </h4>
                            
                            <div className="overflow-auto max-h-[480px] border border-slate-150 rounded-lg shadow-2xs bg-slate-50/20">
                              <table className="w-full text-left border-collapse font-mono text-3xs min-w-max">
                                <thead>
                                  <tr className="bg-slate-100 text-slate-500">
                                    <th className="p-1 px-2 border text-center font-bold sticky top-0 bg-slate-100 z-10 shadow-xs">#</th>
                                    {Array.from({ length: visualGridBounds.maxCol + 1 }).map((_, colIdx) => (
                                      <th key={colIdx} className="p-1 border text-center font-bold text-slate-700 sticky top-0 bg-slate-100 z-10 shadow-xs">
                                        {getColLetter(colIdx)}
                                      </th>
                                    ))}
                                  </tr>
                                </thead>
                                <tbody>
                                  {Array.from({ length: visualGridBounds.maxRow + 1 }).map((_, rowIdx) => (
                                    <tr key={rowIdx}>
                                      <td className="p-1 border bg-slate-50 text-center font-bold text-slate-500 text-3xs sticky left-0 bg-slate-50 z-2">{rowIdx + 1}</td>
                                      {Array.from({ length: visualGridBounds.maxCol + 1 }).map((_, colIdx) => {
                                        const key = `${rowIdx},${colIdx}`;
                                        const logEntry = sheetErrorMap.get(key);
                                        const hasErr = logEntry && isErrorOfType(logEntry, activeErrorFilter);
                                        const cell = employeeWb?.sheets[selectedSheetExplorer]?.cells[key];

                                        let bgCol = "bg-white";
                                        const heatmapCount = showHeatmap ? (() => {
                                          if (!qaAnalysis?.errorLog) return 0;
                                          let count = 0;
                                          for (const err of qaAnalysis.errorLog) {
                                            if (err.sheet === selectedSheetExplorer && err.rowIndex !== undefined && err.colIndex !== undefined) {
                                              if (Math.abs(err.rowIndex - rowIdx) <= 1 && Math.abs(err.colIndex - colIdx) <= 1) {
                                                count++;
                                              }
                                            }
                                          }
                                          return count;
                                        })() : 0;

                                        if (showHeatmap) {
                                          if (heatmapCount === 0) {
                                            bgCol = "bg-slate-55 text-slate-350 border-slate-100";
                                          } else if (heatmapTheme === "coral") {
                                            if (heatmapCount === 1) bgCol = "bg-amber-100 text-amber-955 font-medium border-amber-200";
                                            else if (heatmapCount === 2) bgCol = "bg-orange-120 text-orange-955 font-semibold border-orange-200";
                                            else if (heatmapCount <= 4) bgCol = "bg-orange-300 text-orange-955 font-bold border-orange-400 ring-1 ring-inset ring-orange-400/20";
                                            else bgCol = "bg-red-400 text-white font-bold border-red-500 ring-1 ring-inset ring-red-650 animate-pulse";
                                          } else if (heatmapTheme === "ruby") {
                                            if (heatmapCount === 1) bgCol = "bg-rose-100 text-rose-955 font-medium border-rose-250";
                                            else if (heatmapCount === 2) bgCol = "bg-rose-205 text-rose-955 font-semibold border-rose-300";
                                            else if (heatmapCount <= 4) bgCol = "bg-rose-400 text-white font-bold border-rose-500 ring-1 ring-inset ring-rose-550/20";
                                            else bgCol = "bg-pink-705 text-white font-bold border-pink-850 ring-1 ring-inset ring-pink-905 animate-pulse";
                                          } else {
                                            if (heatmapCount === 1) bgCol = "bg-lime-50 text-lime-900 font-medium border-lime-200";
                                            else if (heatmapCount === 2) bgCol = "bg-emerald-100 text-emerald-955 font-semibold border-emerald-250";
                                            else if (heatmapCount <= 4) bgCol = "bg-emerald-305 text-emerald-955 font-bold border-emerald-450";
                                            else bgCol = "bg-emerald-600 text-white font-bold border-emerald-750 ring-1 animate-pulse";
                                          }
                                        } else if (hasErr) {
                                          if (logEntry.errorType === ErrorType.RowShift || logEntry.errorType === ErrorType.ColumnShift || logEntry.errorType === ErrorType.MissingRow || logEntry.errorType === ErrorType.ExtraRow) {
                                            bgCol = "bg-amber-100/90 text-amber-950 font-semibold";
                                          } else {
                                            bgCol = "bg-red-50/90 text-red-950 border-red-200 font-semibold";
                                          }
                                        }

                                        return (
                                          <td key={colIdx} className={`p-1 px-2 border max-w-[120px] truncate ${bgCol}`} title={cell?.formatted || ""}>
                                            {cell ? cell.formatted : ""}
                                          </td>
                                        );
                                      })}
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            </div>
                          </div>

                          {/* File B Grid */}
                          <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs flex flex-col min-w-0">
                            {(() => {
                              const inspectorRevSheet = qaAnalysis?.virtualSheets?.[selectedSheetExplorer] || reviewerWb?.sheets[selectedSheetExplorer];
                              return (
                                <>
                                  <h4 className="text-[11px] uppercase tracking-wider font-bold font-mono text-slate-600 mb-3 border-b border-slate-100 pb-2 flex items-center justify-between flex-wrap gap-2">
                                    <span className="flex items-center gap-1.5">
                                      <span className="w-2 h-2 rounded-full bg-emerald-500"></span>
                                      File B: Reviewer Ground Truth
                                    </span>
                                    <span className="text-3xs text-slate-500 font-mono font-normal normal-case">
                                      Scale: {(inspectorRevSheet?.maxRow ?? 0) + 1}R x {(inspectorRevSheet?.maxCol ?? 0) + 1}C
                                    </span>
                                  </h4>

                                  {!inspectorRevSheet ? (
                                    <div className="p-6 text-center text-xs text-red-500 italic bg-red-50 rounded-lg border border-red-100 font-sans my-auto">
                                      Warning: Sheet '{selectedSheetExplorer}' does not exist in the Reviewer verification file!
                                    </div>
                                  ) : (
                                    <div className="overflow-auto max-h-[480px] border border-slate-150 rounded-lg shadow-2xs bg-slate-50/20">
                                      <table className="w-full text-left border-collapse font-mono text-3xs min-w-max">
                                        <thead>
                                          <tr className="bg-slate-100 text-slate-500">
                                            <th className="p-1 px-2 border text-center font-bold sticky top-0 bg-slate-100 z-10 shadow-xs">#</th>
                                            {Array.from({ length: visualGridBounds.maxCol + 1 }).map((_, colIdx) => (
                                              <th key={colIdx} className="p-1 border text-center font-bold text-slate-700 sticky top-0 bg-slate-100 z-10 shadow-xs">
                                                {getColLetter(colIdx)}
                                              </th>
                                            ))}
                                          </tr>
                                        </thead>
                                        <tbody>
                                          {Array.from({ length: visualGridBounds.maxRow + 1 }).map((_, rowIdx) => (
                                            <tr key={rowIdx}>
                                              <td className="p-1 border bg-slate-50 text-center font-bold text-slate-500 sticky left-0 bg-slate-50 z-2">{rowIdx + 1}</td>
                                              {Array.from({ length: visualGridBounds.maxCol + 1 }).map((_, colIdx) => {
                                                const key = `${rowIdx},${colIdx}`;
                                                const logEntry = sheetErrorMap.get(key);
                                                const hasErr = logEntry && isErrorOfType(logEntry, activeErrorFilter);
                                                const cell = inspectorRevSheet?.cells[key];

                                                let bgCol = "bg-white";
                                                const heatmapCount = showHeatmap ? (() => {
                                                  if (!qaAnalysis?.errorLog) return 0;
                                                  let count = 0;
                                                  for (const err of qaAnalysis.errorLog) {
                                                    if (err.sheet === selectedSheetExplorer && err.rowIndex !== undefined && err.colIndex !== undefined) {
                                                      if (Math.abs(err.rowIndex - rowIdx) <= 1 && Math.abs(err.colIndex - colIdx) <= 1) {
                                                        count++;
                                                      }
                                                    }
                                                  }
                                                  return count;
                                                })() : 0;

                                                if (showHeatmap) {
                                                  if (heatmapCount === 0) {
                                                    bgCol = "bg-slate-55 text-slate-350 border-slate-100";
                                                  } else if (heatmapTheme === "coral") {
                                                    if (heatmapCount === 1) bgCol = "bg-amber-100 text-amber-955 font-medium border-amber-200";
                                                    else if (heatmapCount === 2) bgCol = "bg-orange-120 text-orange-955 font-semibold border-orange-200";
                                                    else if (heatmapCount <= 4) bgCol = "bg-orange-300 text-orange-955 font-bold border-orange-400 ring-1 ring-inset ring-orange-400/20";
                                                    else bgCol = "bg-red-400 text-white font-bold border-red-500 ring-1 ring-inset ring-red-650 animate-pulse";
                                                  } else if (heatmapTheme === "ruby") {
                                                    if (heatmapCount === 1) bgCol = "bg-rose-100 text-rose-955 font-medium border-rose-250";
                                                    else if (heatmapCount === 2) bgCol = "bg-rose-205 text-rose-955 font-semibold border-rose-300";
                                                    else if (heatmapCount <= 4) bgCol = "bg-rose-400 text-white font-bold border-rose-500 ring-1 ring-inset ring-rose-550/20";
                                                    else bgCol = "bg-pink-705 text-white font-bold border-pink-850 ring-1 ring-inset ring-pink-905 animate-pulse";
                                                  } else {
                                                    if (heatmapCount === 1) bgCol = "bg-lime-50 text-lime-900 font-medium border-lime-200";
                                                    else if (heatmapCount === 2) bgCol = "bg-emerald-100 text-emerald-955 font-semibold border-emerald-250";
                                                    else if (heatmapCount <= 4) bgCol = "bg-emerald-305 text-emerald-955 font-bold border-emerald-450";
                                                    else bgCol = "bg-emerald-600 text-white font-bold border-emerald-750 ring-1 animate-pulse";
                                                  }
                                                } else if (hasErr) {
                                                  if (logEntry.errorType === ErrorType.RowShift || logEntry.errorType === ErrorType.ColumnShift || logEntry.errorType === ErrorType.MissingRow || logEntry.errorType === ErrorType.ExtraRow) {
                                                    bgCol = "bg-amber-100/95 text-amber-950 font-semibold";
                                                  } else {
                                                    bgCol = "bg-emerald-50 text-emerald-950 font-semibold";
                                                  }
                                                }

                                                return (
                                                  <td key={colIdx} className={`p-1 px-2 border max-w-[120px] truncate ${bgCol}`} title={cell?.formatted || ""}>
                                                    {cell ? cell.formatted : ""}
                                                  </td>
                                                );
                                              })}
                                            </tr>
                                          ))}
                                        </tbody>
                                      </table>
                                    </div>
                                  )}
                                </>
                              );
                            })()}
                          </div>

                        </div>
                      ) : (
                        <div className="overflow-auto max-h-[500px] border border-slate-200 rounded-lg p-4 bg-gray-50">
                          {/* Grid layout with interactive overlay detailing errors */}
                          <table className="w-full text-left border-collapse font-mono text-3xs">
                            <thead>
                              <tr className="bg-gray-100 text-gray-500">
                                <th className="p-1.5 px-3 border text-center font-bold">Coord Index</th>
                                <th className="p-1.5 px-3 border">Employee Submit Value</th>
                                <th className="p-1.5 px-3 border">Reviewer Verified Value</th>
                                <th className="p-1.5 px-3 border">Discrepancy Category Flag</th>
                                <th className="p-1.5 px-3 border">Auditor Diagnostics Report Notes</th>
                              </tr>
                            </thead>
                            <tbody>
                              {qaAnalysis.errorLog
                                .filter(x => x.sheet === selectedSheetExplorer && isErrorOfType(x, activeErrorFilter))
                                .map((err, idx) => (
                                  <tr key={idx} className="bg-white hover:bg-gray-50">
                                    <td className="p-2 border font-bold text-center text-slate-800 bg-gray-50">{err.cell}</td>
                                    <td className="p-2 border text-red-650 font-semibold">{err.employeeValue || <span className="text-gray-400 italic font-sans font-normal">empty space</span>}</td>
                                    <td className="p-2 border text-emerald-800 font-semibold">{err.reviewerValue || <span className="text-gray-400 italic font-sans font-normal">empty space</span>}</td>
                                    <td className="p-2 border">
                                      <span className="bg-amber-100 text-amber-900 p-1 rounded text-4xs font-bold uppercase">
                                        {err.errorType}
                                      </span>
                                    </td>
                                    <td className="p-2 border text-gray-500 leading-normal">{err.notes}</td>
                                  </tr>
                                ))}
                              {qaAnalysis.errorLog.filter(x => x.sheet === selectedSheetExplorer && isErrorOfType(x, activeErrorFilter)).length === 0 && (
                                <tr>
                                  <td colSpan={5} className="py-8 text-center text-gray-400 italic bg-white">
                                    {activeErrorFilter === "ALL"
                                      ? "This sheet lists exactly 0 errors. All coordinates are fully verified!"
                                      : `This sheet lists exactly 0 errors matching the '${activeErrorFilter}' filter category.`
                                    }
                                  </td>
                                </tr>
                              )}
                            </tbody>
                          </table>
                        </div>
                      )}
                    </div>
                  </div>
                )}
              </motion.div>
            )}

            {activeTab === "patterns" && (
              <motion.div
                key="patterns"
                initial={{ opacity: 0, y: 5 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -5 }}
                transition={{ duration: 0.15 }}
                className="space-y-6 animate-fade-in"
              >
                
                {/* Visual grid layout shift event warnings lists */}
                <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                  
                  {/* Left panel: Row Shifts */}
                  <div className="bg-white border border-slate-200 rounded-xl p-6 shadow-sm">
                    <h3 className="font-bold text-xs uppercase tracking-widest font-mono text-slate-750 mb-4 border-b border-slate-100 pb-2 flex items-center justify-between">
                      <span>Shift Event Audit Logs</span>
                      <span className="bg-red-50 text-red-700 text-[10px] font-bold px-2 py-0.5 rounded-full border border-red-155">
                        {qaAnalysis.patterns.shiftEvents.length} Events Detected
                      </span>
                    </h3>

                    {qaAnalysis.patterns.shiftEvents.length > 0 ? (
                      <div className="space-y-4">
                        {qaAnalysis.patterns.shiftEvents.map((evt, idx) => {
                          return (
                            <div key={idx} className="p-4 bg-amber-50/70 border border-amber-200 rounded-lg">
                              <div className="flex justify-between items-center mb-1.5">
                                <span className="text-[10px] font-mono font-bold uppercase tracking-widest block bg-amber-900 text-white px-2 py-0.5 rounded-sm">
                                  {evt.type.toUpperCase()} TRANSMISSION SHIFT
                                </span>
                                <span className="text-[9px] font-mono text-amber-800 font-bold uppercase">
                                  Range index: {evt.spanStart + 1} to {evt.spanEnd + 1}
                                </span>
                              </div>
                              <p className="text-xs text-amber-900 leading-relaxed font-semibold">
                                {evt.detail}
                              </p>
                              <p className="text-[10px] text-amber-805 mt-2 font-mono">
                                By-passes {evt.affectedCellsCount} coordinate evaluations to avoid double-counting cell discrepancies.
                              </p>
                            </div>
                          );
                        })}
                      </div>
                    ) : (
                      <div className="py-12 text-center text-slate-400 italic text-xs font-sans">
                        No vertical Row or horizontal Column alignment shifts isolated in submission logs.
                      </div>
                    )}
                  </div>

                  {/* Right panel: Static diagnostic patterns */}
                  <div className="bg-white border border-slate-200 rounded-xl p-6 shadow-sm space-y-6">
                    <div>
                      <h4 className="font-bold text-[11px] uppercase tracking-widest font-sans text-slate-700 mb-3 block">
                        Repeated Numeric Substitutions (Keypad Errors)
                      </h4>
                      {qaAnalysis.patterns.repeatedNumericErrors.length > 0 ? (
                        <ul className="space-y-2 text-xs">
                          {qaAnalysis.patterns.repeatedNumericErrors.map((err, idx) => {
                            return (
                              <li key={idx} className="p-2.5 bg-slate-50 rounded-lg border border-slate-150 flex items-center gap-2">
                                <AlertCircle className="w-3.5 h-3.5 text-slate-400" />
                                <span className="text-slate-700 font-sans">{err}</span>
                              </li>
                            );
                          })}
                        </ul>
                      ) : (
                        <p className="text-xs text-slate-400 italic">No recurring coordinate digit substitutions detected.</p>
                      )}
                    </div>

                    <div>
                      <h4 className="font-bold text-[11px] uppercase tracking-widest font-sans text-slate-700 mb-3 block">
                        Copy-Paste Value Replication Anomalies
                      </h4>
                      {qaAnalysis.patterns.copyPasteErrors.length > 0 ? (
                        <ul className="space-y-2 text-xs">
                          {qaAnalysis.patterns.copyPasteErrors.map((err, idx) => {
                            return (
                              <li key={idx} className="p-2.5 bg-slate-50 border border-slate-100 rounded-lg flex items-center gap-2">
                                <AlertCircle className="w-3.5 h-3.5 text-slate-400" />
                                <span className="text-slate-700 font-sans">{err}</span>
                              </li>
                            );
                          })}
                        </ul>
                      ) : (
                        <p className="text-xs text-slate-400 italic font-sans text-xs">No systematic copy-paste errors detected.</p>
                      )}
                    </div>
                  </div>
                </div>

                <div className="bg-white border border-slate-200 rounded-xl p-6 shadow-sm">
                  <h3 className="font-bold text-xs uppercase tracking-widest font-mono text-slate-750 mb-4 border-b border-slate-100 pb-2">
                    Sheet-Level Location Concentrations & Error Clusters
                  </h3>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    <div>
                      <h4 className="text-[11px] font-bold uppercase tracking-widest font-sans text-slate-500 mb-3">Sheet Error Densities:</h4>
                      <ul className="space-y-2 text-xs">
                        {qaAnalysis.patterns.sheetConcentrations.map((sh, idx) => {
                          return (
                            <li key={idx} className="p-2.5 bg-slate-50 border border-slate-100 rounded-lg text-slate-700 font-mono text-2xs truncate">
                              {sh}
                            </li>
                          );
                        })}
                      </ul>
                    </div>
                    <div>
                      <h4 className="text-[11px] font-bold uppercase tracking-widest font-sans text-slate-500 mb-3">High Density Neighborhood Clusters:</h4>
                      {qaAnalysis.patterns.errorClusters.length > 0 ? (
                        <ul className="space-y-2 text-xs">
                          {qaAnalysis.patterns.errorClusters.map((cl, idx) => {
                            return (
                              <li key={idx} className="p-2.5 bg-amber-50/70 border border-amber-150 rounded-lg text-amber-900 font-bold font-sans text-xs">
                                {cl}
                              </li>
                            );
                          })}
                        </ul>
                      ) : (
                        <p className="text-xs text-slate-400 italic font-mono uppercase tracking-wider text-2xs">No dense neighborhood error clusters isolated on sheet coordinates.</p>
                      )}
                    </div>
                  </div>
                </div>

              </motion.div>
            )}

            {activeTab === "aiAuditor" && (
              <motion.div
                key="aiAuditor"
                initial={{ opacity: 0, y: 5 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -5 }}
                transition={{ duration: 0.15 }}
                className="space-y-6"
              >
                
                {/* Call Out banner to activate Gemini */}
                {!aiResult && (
                  <div className="bg-slate-950 text-white rounded-2xl p-10 border border-slate-800 text-center max-w-2xl mx-auto space-y-6 shadow-2xl">
                    <div className="bg-gradient-to-r from-indigo-650 to-violet-600 inline-flex p-4 rounded-xl items-center justify-center mx-auto mb-2 shadow-inner">
                      <Sparkles className="w-8 h-8 text-white animate-pulse" />
                    </div>
                    <h2 className="text-lg font-bold font-sans tracking-tight text-slate-100">
                      Connect to server-side Gemini AI Auditor
                    </h2>
                    <p className="text-xs text-slate-400 leading-relaxed font-sans max-w-lg mx-auto">
                      Send your calculated error metrics, spelling distribution matrices, layout shifts, error logs, and coordinates directly to Gemini. This compiles a senior QA specialist executive audit summary, failure drivers analysis, and tailored employee coaching plans.
                    </p>
                    
                    <button
                      onClick={triggerAIAudit}
                      className="inline-flex items-center gap-2 px-6 py-3 rounded-lg text-xs font-bold bg-indigo-600 hover:bg-indigo-700 text-white transition-all shadow-lg hover:shadow-indigo-500/10 border border-indigo-500 pointer-events-auto cursor-pointer"
                    >
                      <Play className="w-4 h-4 text-white bg-transparent" />
                      Generate Advanced AI Audit Evaluation Report
                    </button>
                    
                    <span className="text-[10px] text-slate-500 block font-mono">
                      * Uses server-side process.env.GEMINI_API_KEY. Bypasses securely with premium mockups if not configured.
                    </span>
                  </div>
                )}

                {aiResult && (
                  <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                    
                    {/* Left & Mid Column: Narrative assessment Markdown representation */}
                    <div className="bg-white border border-slate-200 rounded-xl p-6 shadow-sm md:col-span-2 space-y-6">
                      <div className="flex items-center gap-2 border-b border-gray-100 pb-3">
                        <Sparkles className="w-5 h-5 text-violet-600" />
                        <h3 className="font-bold text-sm uppercase tracking-wider font-mono text-gray-800">
                          Senior auditor executive analysis & failure diagnosis
                        </h3>
                      </div>

                      <div className="prose prose-xs max-w-none text-xs text-gray-700 leading-relaxed font-sans space-y-4">
                        {aiResult.summary.split("\n").map((line, lineIdx) => {
                          if (line.startsWith("###")) {
                            return <h4 key={lineIdx} className="text-sm font-bold text-gray-800 mt-6 font-mono border-b pb-1">{line.replace("###", "")}</h4>;
                          }
                          if (line.startsWith("##")) {
                            return <h3 key={lineIdx} className="text-md font-bold text-gray-900 mt-8 font-mono">{line.replace("##", "")}</h3>;
                          }
                          if (line.startsWith("* **")) {
                            const bulletContent = line.replace("* **", "").split("**: ");
                            return (
                              <div key={lineIdx} className="flex gap-2 items-start text-xs mt-2 pl-2">
                                <span className="text-violet-600 font-bold">•</span>
                                <p>
                                  <strong className="text-gray-800">{bulletContent[0]}</strong>: {bulletContent[1]}
                                </p>
                              </div>
                            );
                          }
                          if (line.startsWith("- **")) {
                            const bulletContent = line.replace("- **", "").split("**: ");
                            return (
                              <div key={lineIdx} className="flex gap-2 items-start text-xs mt-2 pl-2">
                                <span className="text-violet-600 font-bold">•</span>
                                <p>
                                  <strong className="text-gray-800">{bulletContent[0]}</strong>: {bulletContent[1]}
                                </p>
                              </div>
                            );
                          }
                          if (line.startsWith("* ")) {
                            return (
                              <li key={lineIdx} className="list-disc pl-5 mt-1 font-sans">
                                {line.replace("* ", "")}
                              </li>
                            );
                          }
                          return <p key={lineIdx} className="font-sans">{line}</p>;
                        })}
                      </div>

                      <div className="mt-8 pt-4 border-t border-slate-100 flex items-center justify-between text-[10px] font-mono text-slate-400">
                        <span>Evaluation Signature Token: QA-SECURE-GEN-AUDIT</span>
                        <span>Auditor ID: GEMINI-3.5-FLASH-ENGINE</span>
                      </div>
                    </div>

                    {/* Right column: Interactive coaching targets list from AI API */}
                    <div className="bg-white border border-slate-200 rounded-xl p-6 shadow-sm space-y-6">
                      <div className="flex items-center gap-2 border-b border-slate-100 pb-3">
                        <SlidersHorizontal className="w-5 h-5 text-indigo-600" />
                        <h3 className="font-bold text-xs uppercase tracking-widest font-mono text-slate-800">
                          Interactive training guide
                        </h3>
                      </div>

                      <p className="text-xs text-slate-500 leading-relaxed font-sans">
                        The AI Auditor synthesized compliance tasks directly mapped to correct Employee performance error sequences:
                      </p>

                      <div className="space-y-4 font-sans">
                        {aiResult.coaching.map((coachText, idx) => (
                          <div key={idx} className="p-4 bg-indigo-50/45 border border-indigo-150 rounded-lg space-y-2">
                            <span className="text-[10px] font-mono font-bold bg-indigo-600 text-white px-2.5 py-0.5 rounded uppercase">
                              Target Action #{idx + 1}
                            </span>
                            <p className="text-xs text-slate-700 leading-relaxed font-sans mt-2">
                              {coachText}
                            </p>
                          </div>
                        ))}
                      </div>
                    </div>
                    
                  </div>
                )}
              </motion.div>
            )}

            {activeTab === "mappingExplorer" && (
              <motion.div
                key="mappingExplorer"
                initial={{ opacity: 0, y: 5 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -5 }}
                transition={{ duration: 0.15 }}
                className="space-y-6"
              >
                {/* Visual Mapping HUD Banner */}
                <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
                  <div className="bg-white p-5 border border-slate-200 rounded-xl shadow-xs">
                    <span className="text-[10px] font-mono text-slate-400 uppercase tracking-widest font-bold">Matched Sheet Alignment</span>
                    <h4 className="text-2xl font-black text-indigo-700 mt-1 font-sans">
                      {qaAnalysis.sheetMatches?.length || 0}
                    </h4>
                    <p className="text-3xs text-slate-550 font-medium mt-1">Unique workbook sheets matched successfully</p>
                  </div>

                  <div className="bg-white p-5 border border-slate-200 rounded-xl shadow-xs">
                    <span className="text-[10px] font-mono text-slate-400 uppercase tracking-widest font-bold">Fuzzy Alignments</span>
                    <h4 className="text-2xl font-black text-amber-600 mt-1 font-sans">
                      {qaAnalysis.sheetMatches?.filter((m: any) => m.matchType === "Fuzzy Match").length || 0}
                    </h4>
                    <p className="text-3xs text-slate-550 font-medium mt-1">Needs translation / minor typo tolerance</p>
                  </div>

                  <div className="bg-white p-5 border border-slate-200 rounded-xl shadow-xs">
                    <span className="text-[10px] font-mono text-slate-400 uppercase tracking-widest font-bold">Merge/Split Groups</span>
                    <h4 className="text-2xl font-black text-purple-605 mt-1 font-sans">
                      {qaAnalysis.sheetMatches?.filter((m: any) => m.matchType === "Split Match" || m.matchType === "Merge Match").length || 0}
                    </h4>
                    <p className="text-3xs text-slate-555 font-medium mt-1">Complex layout variations isolated</p>
                  </div>

                  <div className="bg-white p-5 border border-slate-200 rounded-xl shadow-xs">
                    <span className="text-[10px] font-mono text-slate-400 uppercase tracking-widest font-bold">Average Match Precision</span>
                    <h4 className="text-2xl font-black text-emerald-600 mt-1 font-sans">
                      {(() => {
                        if (!qaAnalysis.sheetMatches || qaAnalysis.sheetMatches.length === 0) return "100%";
                        const sum = qaAnalysis.sheetMatches.reduce((acc: number, val: any) => acc + val.matchConfidence, 0);
                        return Math.round(sum / qaAnalysis.sheetMatches.length) + "%";
                      })()}
                    </h4>
                    <p className="text-3xs text-slate-555 font-medium mt-1">Cross-matching similarity confidence weight</p>
                  </div>
                </div>

                {/* Subtitle list header */}
                <div className="flex items-center justify-between pointer-events-auto">
                  <div>
                    <h3 className="text-sm font-bold text-slate-800 uppercase tracking-wider font-mono flex items-center gap-2">
                      <Layers className="w-4 h-4 text-indigo-505" /> Ground Truth Structural Pairings Explorer
                    </h3>
                    <p className="text-xs text-slate-550 font-sans mt-1">See precisely how the deterministic shift precedence engine maps incoming reviewer sheets to corresponding worker worksheets.</p>
                  </div>
                </div>

                {/* Grid items */}
                <div className="grid grid-cols-1 gap-4 font-sans">
                  {(!qaAnalysis.sheetMatches || qaAnalysis.sheetMatches.length === 0) ? (
                    <div className="text-center p-12 bg-white rounded-xl border border-dashed border-slate-195">
                      <Layers className="w-8 h-8 text-slate-300 mx-auto mb-2" />
                      <p className="text-xs text-slate-450 font-mono uppercase tracking-wider font-semibold">No sheet mappings loaded for evaluation.</p>
                    </div>
                  ) : (
                    qaAnalysis.sheetMatches.map((m: any, idx: number) => {
                      const confidenceVal = Math.round(m.matchConfidence < 1.1 ? m.matchConfidence * 100 : m.matchConfidence);
                      
                      return (
                        <div key={idx} className="bg-white border border-slate-200 rounded-xl p-5 shadow-xs transition hover:shadow-sm">
                          <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
                            {/* Visual Linking Card Section */}
                            <div className="flex items-center gap-3 flex-1">
                              <div className="p-2.5 bg-slate-50 border border-slate-200 rounded-lg min-w-[160px]">
                                <span className="text-[9px] font-mono text-slate-400 block font-bold">SOURCE REVIEWER</span>
                                <span className="text-xs font-black font-sans text-slate-700 truncate block max-w-[200px]" title={m.reviewerSheetName}>
                                  {m.reviewerSheetName}
                                </span>
                              </div>
                              
                              <div className="flex flex-col items-center flex-shrink-0 px-2 font-sans">
                                <span className="text-[10px] font-bold text-indigo-700 bg-indigo-50 border border-indigo-100 px-2 py-0.5 rounded">
                                  {m.matchType}
                                </span>
                                <div className="h-0.5 w-12 border-t border-dashed border-slate-300 my-1"></div>
                              </div>

                              <div className="p-2.5 bg-indigo-50/20 border border-indigo-102 rounded-lg min-w-[160px]">
                                <span className="text-[9px] font-mono text-indigo-500 block font-bold font-semibold">TARGET WORKERS</span>
                                <span className="text-xs font-semibold font-sans text-indigo-955 truncate block max-w-[200px]" title={m.matchedEmployeeSheets.join(", ")}>
                                  {m.matchedEmployeeSheets.join(", ") || "(No Sheets Matched)"}
                                </span>
                              </div>
                            </div>

                            {/* Confidence Progress */}
                            <div className="flex items-center gap-6">
                              <div className="text-right">
                                <span className="text-[10px] text-slate-400 font-mono uppercase font-bold block">Match Confidence</span>
                                <div className="flex items-center gap-2 mt-1">
                                  <div className="w-24 bg-slate-100 rounded-full h-1.5 overflow-hidden">
                                    <div 
                                      className={`h-full rounded-full ${
                                        confidenceVal >= 90 ? 'bg-emerald-500' : confidenceVal >= 75 ? 'bg-yellow-500' : 'bg-rose-500'
                                      }`}
                                      style={{ width: `${confidenceVal}%` }}
                                    ></div>
                                  </div>
                                  <span className={`text-xs font-black font-sans ${
                                    confidenceVal >= 90 ? 'text-emerald-600' : confidenceVal >= 75 ? 'text-yellow-600' : 'text-rose-600'
                                  }`}>
                                    {confidenceVal}%
                                  </span>
                                </div>
                              </div>
                            </div>
                          </div>

                          {/* Factors Grid Block */}
                          <div className="mt-4 pt-4 border-t border-slate-100 grid grid-cols-2 md:grid-cols-4 gap-4">
                            <div className="p-2 bg-slate-50 border border-slate-150 rounded space-y-1">
                              <span className="text-[9px] font-mono text-slate-400 block uppercase font-bold">Name Similarity</span>
                              <div className="flex items-center justify-between">
                                <span className="text-xs font-bold font-sans text-slate-700">
                                  {Math.round((m.factors?.nameSimilarity ?? 1) * 100)}%
                                </span>
                              </div>
                            </div>

                            <div className="p-2 bg-slate-50 border border-slate-150 rounded space-y-1">
                              <span className="text-[9px] font-mono text-slate-400 block uppercase font-bold">Header Similarity</span>
                              <div className="flex items-center justify-between">
                                <span className="text-xs font-bold font-sans text-slate-700">
                                  {Math.round((m.factors?.headerSimilarity ?? 1) * 100)}%
                                </span>
                              </div>
                            </div>

                            <div className="p-2 bg-slate-50 border border-slate-150 rounded space-y-1">
                              <span className="text-[9px] font-mono text-slate-400 block uppercase font-bold">Structural Dimensions</span>
                              <div className="flex items-center justify-between">
                                <span className="text-xs font-bold font-sans text-slate-700">
                                  {Math.round((m.factors?.structuralSimilarity ?? 1) * 100)}%
                                </span>
                              </div>
                            </div>

                            <div className="p-2 bg-slate-50 border border-slate-150 rounded space-y-1">
                              <span className="text-[9px] font-mono text-slate-400 block uppercase font-bold">Cell Data Overlap</span>
                              <div className="flex items-center justify-between">
                                <span className="text-xs font-bold font-sans text-slate-700">
                                  {Math.round((m.factors?.dataOverlap ?? 1) * 100)}%
                                </span>
                              </div>
                            </div>
                          </div>
                        </div>
                      );
                    })
                  )}
                </div>
              </motion.div>
            )}

            {activeTab === "benchmark" && (
              <motion.div
                key="benchmark"
                initial={{ opacity: 0, y: 5 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -5 }}
                transition={{ duration: 0.15 }}
                className="space-y-6"
              >
                {/* Upper description */}
                <div className="bg-white border border-slate-200 rounded-xl p-5 shadow-xs flex flex-col md:flex-row md:items-center justify-between gap-4">
                  <div className="space-y-1">
                    <h3 className="text-sm font-bold text-slate-800 uppercase tracking-wider font-mono flex items-center gap-2">
                      <Activity className="text-emerald-500 w-4 h-4" /> Compliance Regression Testing HUD
                    </h3>
                    <p className="text-xs text-slate-500 font-sans mt-1">
                      Prevent employee error rates from inflating across revisions. Establish a benchmark reference baseline to cross-audit changes.
                    </p>
                  </div>
                  
                  {benchmarkBaseline && (
                    <button
                      onClick={() => {
                        setBenchmarkBaseline(null);
                        setBaselineFileName("");
                        setBaselineUploadError(null);
                      }}
                      className="px-3 py-1.5 bg-slate-150 hover:bg-slate-200 text-slate-700 text-3xs font-semibold uppercase tracking-wider font-mono border border-slate-205 rounded cursor-pointer transition flex items-center gap-1 self-start md:self-auto"
                    >
                      Clear Active Baseline
                    </button>
                  )}
                </div>

                {/* Main section: Either prompt calibration OR show the full matrix comparison view */}
                {!benchmarkBaseline ? (
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-6 font-sans">
                    {/* Left Panel: Seal active assessment as baseline */}
                    <div className="bg-white border border-slate-200 rounded-xl p-6 shadow-xs flex flex-col justify-between space-y-5">
                      <div className="space-y-3">
                        <div className="w-10 h-10 bg-indigo-50 border border-indigo-100 rounded-lg flex items-center justify-center">
                          <CheckCircle className="w-5 h-5 text-indigo-600" />
                        </div>
                        <h4 className="text-sm font-bold text-slate-800 font-sans">Lock Current Assessment as Baseline</h4>
                        <p className="text-xs text-slate-500 leading-relaxed font-sans mt-1">
                          Use the current employee analysis metrics as the active reference checkpoint. The system will mathematically cross-examine future assessments against these exact scores:
                        </p>
                        
                        <div className="p-3.5 bg-slate-50 rounded-lg border border-slate-150 space-y-2 font-mono text-[11px] mt-2">
                          <div className="flex justify-between">
                            <span className="text-slate-400">Project Preset:</span>
                            <span className="font-bold text-slate-700">{config.projectName || "Default Audit"}</span>
                          </div>
                          <div className="flex justify-between">
                            <span className="text-slate-400">Total Score:</span>
                            <span className="font-bold text-emerald-600">{qaAnalysis.metrics.accuracyRatio.toFixed(2)}%</span>
                          </div>
                          <div className="flex justify-between">
                            <span className="text-slate-400">Error Registry:</span>
                            <span className="font-bold text-slate-700">{qaAnalysis.errorLog.length} errors</span>
                          </div>
                          <div className="flex justify-between">
                            <span className="text-slate-400 font-medium">Layout Grade:</span>
                            <span className="font-extrabold text-indigo-600 bg-indigo-50 px-1.5 py-0.2 rounded border border-indigo-100">{qaAnalysis.metrics.grade || "C"}</span>
                          </div>
                        </div>
                      </div>

                      <button
                        onClick={saveCurrentAsBaseline}
                        className="w-full text-center py-2 bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-bold rounded transition shadow-xs cursor-pointer"
                      >
                        Capture Assessment Checklist
                      </button>
                    </div>

                    {/* Right Panel: Import JSON file */}
                    <div className="bg-white border border-slate-200 rounded-xl p-6 shadow-xs flex flex-col justify-between space-y-5">
                      <div className="space-y-3">
                        <div className="w-10 h-10 bg-amber-50 border border-amber-100 rounded-lg flex items-center justify-center">
                          <Upload className="w-5 h-5 text-amber-600" />
                        </div>
                        <h4 className="text-sm font-bold text-slate-800 font-sans">Upload Benchmark Blueprint (JSON)</h4>
                        <p className="text-xs text-slate-500 leading-relaxed font-sans mt-1">
                          Import any previously exported compliance result file to calibrate regression testing. Uploading a prior report lets you execute comparisons offline.
                        </p>

                        <div className="border border-dashed border-slate-300 rounded-lg p-6 bg-slate-50 text-center relative hover:bg-slate-100/50 transition">
                          <input
                            type="file"
                            accept=".json"
                            onChange={handleBaselineJSONUpload}
                            className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
                          />
                          <FileJson className="w-8 h-8 text-slate-400 mx-auto mb-2 animate-pulse" />
                          <span className="text-xs font-sans text-slate-500 font-medium block">
                            Click or drag JSON blueprint here to import
                          </span>
                        </div>
                        
                        {baselineUploadError && (
                          <p className="text-3xs font-mono font-bold text-rose-500 bg-rose-50 p-2 rounded border border-rose-100 mt-2">
                            {baselineUploadError}
                          </p>
                        )}
                      </div>

                      <span className="text-[10px] text-slate-400 italic block text-center font-mono uppercase tracking-wider font-semibold">
                        * Accepts ground-truth JSON analysis outputs only.
                      </span>
                    </div>
                  </div>
                ) : (
                  // Benchmark active results matrix comparison HUD
                  <div className="space-y-6">
                    
                    {/* Active Baseline HUD Header Badge */}
                    <div className="bg-emerald-50 border border-emerald-200 rounded-xl px-5 py-4 flex items-center justify-between font-sans">
                      <div className="flex items-center gap-3">
                        <div className="w-8 h-8 bg-emerald-100 border border-emerald-200 rounded flex items-center justify-center">
                          <Award className="w-4.5 h-4.5 text-emerald-700 font-semibold" />
                        </div>
                        <div>
                          <p className="text-[10px] font-mono text-emerald-600 uppercase tracking-wider font-bold">Calibration Reference Blueprint Locked</p>
                          <h4 className="text-xs font-black text-emerald-900">
                            {baselineFileName} <span className="font-mono text-[10px] text-emerald-600/70">({benchmarkBaseline.projectName} • {benchmarkBaseline.date})</span>
                          </h4>
                        </div>
                      </div>

                      <span className="text-[10px] font-mono bg-emerald-600 text-white px-2.5 py-1 rounded font-bold uppercase">
                        ACTIVE BENCHMARK
                      </span>
                    </div>

                    {/* HUD Matrices */}
                    {(() => {
                      const accuracyDelta = qaAnalysis.metrics.accuracyRatio - benchmarkBaseline.accuracy;
                      const errorsDelta = qaAnalysis.errorLog.length - benchmarkBaseline.totalErrors;
                      const penaltyDelta = (qaAnalysis.metrics.totalStructuralPenalty || 0) - benchmarkBaseline.penaltyScore;
                      const regressed = accuracyDelta < -0.01 || errorsDelta > 0 || penaltyDelta > 0;

                      return (
                        <div className="space-y-6 font-sans">
                          {/* Comparative tiles */}
                          <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                            
                            {/* Score Tile */}
                            <div className="bg-white border border-slate-200 p-5 rounded-xl shadow-xs space-y-3">
                              <span className="text-[10px] font-mono text-slate-400 uppercase tracking-widest font-bold block">Compliance Accuracy</span>
                              <div className="flex items-baseline justify-between">
                                <div>
                                  <span className="text-2xl font-black text-slate-800">{qaAnalysis.metrics.accuracyRatio.toFixed(2)}%</span>
                                  <span className="text-3xs text-slate-400 block mt-1">Baseline: {benchmarkBaseline.accuracy.toFixed(2)}%</span>
                                </div>
                                
                                {accuracyDelta > 0.01 ? (
                                  <span className="inline-flex items-center gap-1 text-[10px] font-mono font-bold bg-emerald-50 text-emerald-700 px-2 py-0.5 rounded-full border border-emerald-100">
                                    <TrendingUp className="w-3 h-3" /> +{accuracyDelta.toFixed(2)}%
                                  </span>
                                ) : accuracyDelta < -0.01 ? (
                                  <span className="inline-flex items-center gap-1 text-[10px] font-mono font-bold bg-rose-50 text-rose-700 px-2 py-0.5 rounded-full border border-rose-100">
                                    <TrendingDown className="w-3 h-3" /> {accuracyDelta.toFixed(2)}%
                                  </span>
                                ) : (
                                  <span className="inline-flex items-center text-[10px] font-mono bg-slate-100 text-slate-500 px-2 py-0.5 rounded-full border border-slate-200 font-bold">
                                    STABLE
                                  </span>
                                )}
                              </div>
                            </div>

                            {/* Total Errors count */}
                            <div className="bg-white border border-slate-200 p-5 rounded-xl shadow-xs space-y-3">
                              <span className="text-[10px] font-mono text-slate-400 uppercase tracking-widest font-bold block">Spelling & Data Errors</span>
                              <div className="flex items-baseline justify-between">
                                <div>
                                  <span className="text-2xl font-black text-slate-800">{qaAnalysis.errorLog.length}</span>
                                  <span className="text-3xs text-slate-400 block mt-1">Baseline: {benchmarkBaseline.totalErrors}</span>
                                </div>

                                {errorsDelta > 0 ? (
                                  <span className="inline-flex items-center gap-1 text-[10px] font-mono font-bold bg-rose-50 text-rose-700 px-2 py-0.5 rounded-full border border-rose-100 font-bold">
                                    +{errorsDelta} regression
                                  </span>
                                ) : errorsDelta < 0 ? (
                                  <span className="inline-flex items-center gap-1 text-[10px] font-mono font-bold bg-emerald-50 text-emerald-705 px-2 py-0.5 rounded-full border border-emerald-100 font-bold">
                                    {errorsDelta} improved
                                  </span>
                                ) : (
                                  <span className="inline-flex items-center text-[10px] font-mono bg-slate-100 text-slate-500 px-2 py-0.5 rounded-full border border-slate-200 font-bold">
                                    STABLE
                                  </span>
                                )}
                              </div>
                            </div>

                            {/* Structural Penalty weights */}
                            <div className="bg-white border border-slate-200 p-5 rounded-xl shadow-xs space-y-3">
                              <span className="text-[10px] font-mono text-slate-400 uppercase tracking-widest font-bold block">Penalty Weight Accumulation</span>
                              <div className="flex items-baseline justify-between">
                                <div>
                                  <span className="text-2xl font-black text-slate-800">{qaAnalysis.metrics.totalStructuralPenalty || 0} pts</span>
                                  <span className="text-3xs text-slate-400 block mt-1">Baseline: {benchmarkBaseline.penaltyScore} pts</span>
                                </div>

                                {penaltyDelta > 0 ? (
                                  <span className="inline-flex items-center gap-1 text-[10px] font-mono font-bold bg-rose-50 text-rose-700 px-2 py-0.5 rounded-full border border-rose-100">
                                    +{penaltyDelta} regression
                                  </span>
                                ) : penaltyDelta < 0 ? (
                                  <span className="inline-flex items-center gap-1 text-[10px] font-mono font-bold bg-emerald-50 text-emerald-700 px-2 py-0.5 rounded-full border border-emerald-100 font-bold">
                                    {penaltyDelta} points
                                  </span>
                                ) : (
                                  <span className="inline-flex items-center text-[10px] font-mono bg-slate-100 text-slate-500 px-2 py-0.5 rounded-full border border-slate-200 font-bold">
                                    STABLE
                                  </span>
                                )}
                              </div>
                            </div>

                          </div>

                          {/* Diagnosis feedback block */}
                          <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                            
                            {/* Detailed Diagnosis Column */}
                            <div className="bg-white border border-slate-200 rounded-xl p-6 shadow-xs md:col-span-2 space-y-4">
                              <h4 className="text-xs font-bold uppercase tracking-widest font-mono text-slate-500 border-b border-slate-100 pb-2">Regression Diagnosis & Performance Drift</h4>
                              
                              {regressed ? (
                                <div className="space-y-4">
                                  <div className="p-4 bg-rose-50 border border-rose-200 rounded-lg flex gap-3 text-rose-950">
                                    <ShieldAlert className="w-5 h-5 text-rose-600 flex-shrink-0 mt-0.5" />
                                    <div>
                                      <h5 className="font-bold text-xs">Compliance Standards Drift Warning</h5>
                                      <p className="text-xs text-rose-700 mt-1 leading-relaxed">
                                        Calculated error counts or penalty multipliers exceed baseline drift parameters. This indicates the spreadsheet revision introduces new employee spelling variances, data-truncations, or missing structural headers. Complete localized training in the executive audit tab to resolve errors.
                                      </p>
                                    </div>
                                  </div>
                                </div>
                              ) : (
                                <div className="space-y-4">
                                  <div className="p-4 bg-emerald-50 border border-emerald-200 rounded-lg flex gap-3 text-emerald-950">
                                    <CheckCircle className="w-5 h-5 text-emerald-600 flex-shrink-0 mt-0.5" />
                                    <div>
                                      <h5 className="font-bold text-xs text-emerald-950">Assessment Compliance Safe</h5>
                                      <p className="text-xs text-emerald-700 mt-1 leading-relaxed">
                                        Perfect! All calculated quality metrics on this worksheet match or outpace standard baseline reference values. No structural regressions, typos, or decimal transcription deviations were introduced.
                                      </p>
                                    </div>
                                  </div>
                                </div>
                              )}

                              {/* Backup operations */}
                              <div className="pt-4 border-t border-slate-100 flex items-center justify-between">
                                <span className="text-[10px] text-slate-400 font-mono">Reference token: DIFF_ENGINE_ACTIVE</span>
                                <button
                                  onClick={exportCurrentRunJSON}
                                  className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-slate-800 hover:bg-slate-700 text-white rounded text-3xs font-semibold cursor-pointer transition uppercase font-mono tracking-wider shadow-xs"
                                >
                                  <Download className="w-3.5 h-3.5" /> Export Compliance Blueprint (.json)
                                </button>
                              </div>
                            </div>

                            {/* Standard limits column */}
                            <div className="bg-white border border-slate-200 rounded-xl p-6 shadow-xs space-y-4">
                              <h4 className="text-xs font-bold uppercase tracking-widest font-mono text-slate-500 border-b border-slate-150 pb-2">Benchmark Limits</h4>
                              <p className="text-xs text-slate-500 leading-relaxed font-sans">
                                Continuous inspection ensures spreadsheets align to the following reference rules:
                              </p>
                              
                              <ul className="space-y-2.5 text-xs font-sans">
                                <li className="flex items-center justify-between p-2 bg-slate-50 rounded border border-slate-100">
                                  <span className="font-mono text-[10px] text-slate-450 uppercase font-bold">Accuracy Buffer</span>
                                  <span className="font-bold text-slate-705">±0.50% max dev</span>
                                </li>
                                <li className="flex items-center justify-between p-2 bg-slate-50 rounded border border-slate-100">
                                  <span className="font-mono text-[10px] text-slate-450 uppercase font-bold">Structural shift limit</span>
                                  <span className="font-bold text-slate-705">0 rows shift</span>
                                </li>
                                <li className="flex items-center justify-between p-2 bg-slate-50 rounded border border-slate-100 font-semibold">
                                  <span className="font-mono text-[10px] text-slate-450 uppercase font-bold text-slate-400">Typograph check count</span>
                                  <span className="font-bold text-slate-705 text-emerald-600 font-bold">Auto Active</span>
                                </li>
                              </ul>
                            </div>

                          </div>
                        </div>
                      );
                    })()}

                  </div>
                )}
              </motion.div>
            )}
          </AnimatePresence>
        </main>
      )}

      {/* 🧾 Document footer audit metadata info */}
      <footer className="mt-20 border-t border-slate-200 bg-slate-50/60 py-12 text-center text-xs text-slate-500 print:hidden font-sans">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 space-y-2">
          <p className="font-mono text-[10px] uppercase tracking-widest text-slate-400 font-bold">
            Enterprise Excel QA Engine • Ground Truth Auditor Portal
          </p>
          <p className="text-[10px] text-slate-400 leading-relaxed">
            Powered by Swiss Accuracy Precedence, Arabic Morphological Cleansers, and Gemini API Workspace secrets.
          </p>
          <p className="text-[10px] text-slate-400">
            All spreadsheet coordinate comparisons and merged-cell propagation activities are structured entirely in the local container sandbox.
          </p>
        </div>
      </footer>

      {/* 🧾 Standalone certificate shown ONLY on actual hardcopy paper prints (invisible during standard screen previews) */}
      {qaAnalysis && (
        <>
          <div className="hidden print:block p-10 bg-white max-w-4xl mx-auto space-y-8 font-sans">
            <div className="flex justify-between items-center border-b pb-6">
              <div>
                <h1 className="text-2xl font-black text-slate-900 tracking-tight">Enterprise Compliance Certificate</h1>
                <p className="text-xs text-slate-550 mt-1">Ground Truth Spreadsheet Quality Alignment Service</p>
              </div>
              <div className="text-right text-xs text-slate-400 font-mono space-y-1">
                <div>Document Ref: QA-PRINT-AUTHENTICATED</div>
                <div>Evaluated Worker: {config.employeeName || "Farid Al-Mansour"}</div>
                <div>Project Preset: {config.projectName || "Standard Audit"}</div>
                <div>Assessment Date: {config.evaluationDate}</div>
              </div>
            </div>

            <div className="bg-purple-50/50 border border-purple-100 rounded-xl p-6 flex justify-between items-center text-slate-900">
              <div>
                <span className="text-[10px] uppercase font-bold text-indigo-700 tracking-widest font-mono block">Final Quality Precedence Grade</span>
                <h2 className="text-lg font-black text-slate-900 mt-1">PERFORMANCE GRADE: {qaAnalysis.metrics.grade || "C"}</h2>
                <p className="text-xs text-slate-500 mt-2 max-w-lg leading-relaxed">
                  This certificate is mathematically authenticated to seal the quality metrics evaluated on files compared side-by-side with ground truth.
                </p>
              </div>
              <div className="bg-indigo-600 text-white rounded-lg p-5 text-center min-w-[120px] font-mono shadow-sm">
                <span className="text-[10px] block font-bold text-indigo-150 uppercase tracking-widest">Score Ratio</span>
                <span className="text-2xl font-black">{qaAnalysis.metrics.accuracyRatio.toFixed(2)}%</span>
              </div>
            </div>

            <div className="grid grid-cols-4 gap-4">
              <div className="border border-slate-200 rounded-lg p-4 bg-slate-50/20">
                <span className="text-[10px] font-mono text-slate-400 block font-bold uppercase">Accuracy</span>
                <p className="text-lg font-bold text-slate-800 mt-1">{qaAnalysis.metrics.accuracyRatio.toFixed(2)}%</p>
              </div>
              <div className="border border-slate-200 rounded-lg p-4 bg-slate-50/20">
                <span className="text-[10px] font-mono text-slate-400 block font-bold uppercase">Worksheets Compared</span>
                <p className="text-lg font-bold text-slate-800 mt-1">{qaAnalysis.sheetMatches?.length || 0}</p>
              </div>
              <div className="border border-slate-200 rounded-lg p-4 bg-slate-50/20">
                <span className="text-[10px] font-mono text-slate-400 block font-bold uppercase">Data Errors</span>
                <p className="text-lg font-bold text-slate-800 mt-1">{qaAnalysis.errorLog.length} errors</p>
              </div>
              <div className="border border-slate-200 rounded-lg p-4 bg-slate-50/20">
                <span className="text-[10px] font-mono text-slate-400 block font-bold uppercase">Penalty Points</span>
                <p className="text-lg font-bold text-slate-800 mt-1">{qaAnalysis.metrics.totalStructuralPenalty || 0} pts</p>
              </div>
            </div>

            <div className="space-y-4">
              <span className="text-xs font-bold uppercase tracking-widest font-mono text-slate-500 block border-b pb-2">Top Logged Transcription Errors</span>
              {qaAnalysis.errorLog.length === 0 ? (
                <p className="text-xs text-slate-400 italic">No errors logged during this assessment cycle.</p>
              ) : (
                qaAnalysis.errorLog.slice(0, 15).map((e: any, idx: number) => (
                  <div key={idx} className="border border-slate-100 rounded bg-slate-50/30 p-3 text-xs flex justify-between items-center">
                    <div>
                      <span className="font-bold text-slate-800">Sheet: {e.sheet} | Cell: {e.cell}</span> {e.errorType}
                      <div className="text-slate-500 mt-1 text-[11px]">Given: "{e.employeeValue}" &bull; Replaced by target: "{e.reviewerValue}"</div>
                    </div>
                    <span className="font-mono text-rose-600 bg-rose-50 px-2 py-0.5 rounded border border-rose-100/60 font-semibold">- {e.penalty} pts</span>
                  </div>
                ))
              )}
            </div>

            <div className="flex justify-between mt-20 pt-10 text-xs">
              <div className="w-[200px] border-t border-slate-300 text-center pt-2 text-slate-550 font-semibold">
                <div className="h-10"></div>
                Reviewing Compliance Supervisor
              </div>
              <div className="w-[200px] border-t border-slate-300 text-center pt-2 text-slate-550 font-semibold">
                <div className="h-10"></div>
                Ground Truth Verification Officer
              </div>
              <div className="w-[200px] border-t border-slate-300 text-center pt-2 text-slate-550 font-semibold">
                <div className="h-4 text-indigo-500 font-serif text-sm italic font-bold">Verified & Signed</div>
                <div className="h-6"></div>
                Audit Authentication Stamp
              </div>
            </div>
          </div>

          {/* Off-screen high fidelity container for PDF download generation (html2canvas) */}
          <div
            id="pdf-certificate-content"
            style={{
              position: "absolute",
              left: "-9999px",
              top: "-9999px",
              width: "800px",
              backgroundColor: "#ffffff",
              padding: "40px",
              boxSizing: "border-box",
              fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif",
              color: "#334155"
            }}
          >
            {/* Header */}
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", borderBottom: "2px solid #e2e8f0", paddingBottom: "24px", marginBottom: "32px" }}>
              <div>
                <h1 style={{ fontWeight: 900, fontSize: "24px", color: "#0f172a", margin: 0, letterSpacing: "-0.025em" }}>Enterprise Compliance Certificate</h1>
                <p style={{ fontSize: "12px", color: "#64748b", margin: "4px 0 0 0" }}>Ground Truth Spreadsheet Quality Alignment Service</p>
              </div>
              <div style={{ textAlign: "right", fontSize: "11px", fontFamily: "monospace", color: "#64748b", lineHeight: "1.5" }}>
                <div>Document Ref: QA-PRINT-AUTHENTICATED</div>
                <div>Evaluated Worker: {config.employeeName || "Farid Al-Mansour"}</div>
                <div>Project Preset: {config.projectName || "Standard Audit"}</div>
                <div>Assessment Date: {config.evaluationDate}</div>
              </div>
            </div>

            {/* Performance Grade Box */}
            <div style={{ backgroundColor: "#faf5ff", border: "1px solid #f3e8ff", borderRadius: "12px", padding: "24px", display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "32px", boxSizing: "border-box" }}>
              <div style={{ flex: 1, paddingRight: "24px" }}>
                <span style={{ textTransform: "uppercase", fontWeight: "bold", fontSize: "10px", color: "#4f46e5", letterSpacing: "1px", display: "block" }}>Final Quality Precedence Grade</span>
                <h2 style={{ fontWeight: 900, fontSize: "20px", color: "#0f172a", marginTop: "4px", marginBottom: "8px" }}>PERFORMANCE GRADE: {qaAnalysis.metrics.grade || "C"}</h2>
                <p style={{ color: "#64748b", lineHeight: "1.5", fontSize: "12px", margin: 0 }}>
                  This certificate is mathematically authenticated to seal the quality metrics evaluated on files compared side-by-side with ground truth.
                </p>
              </div>
              <div style={{ backgroundColor: "#4f46e5", color: "#ffffff", borderRadius: "8px", padding: "16px 24px", textAlign: "center", minWidth: "120px", boxSizing: "border-box", flexShrink: 0 }}>
                <span style={{ display: "block", fontSize: "10px", fontWeight: "bold", color: "#c7d2fe", textTransform: "uppercase", letterSpacing: "1px", marginBottom: "4px" }}>Score Ratio</span>
                <span style={{ fontSize: "24px", fontWeight: 900, fontFamily: "monospace" }}>{qaAnalysis.metrics.accuracyRatio.toFixed(2)}%</span>
              </div>
            </div>

            {/* Critical Quality KPIs */}
            <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: "16px", marginBottom: "32px" }}>
              <div style={{ border: "1px solid #cbd5e1", borderRadius: "8px", padding: "16px", backgroundColor: "#f8fafc" }}>
                <span style={{ display: "block", fontSize: "10px", color: "#94a3b8", textTransform: "uppercase", fontWeight: "bold", fontFamily: "monospace" }}>Accuracy</span>
                <p style={{ fontSize: "18px", fontWeight: "bold", color: "#1e293b", margin: "4px 0 0 0", fontFamily: "monospace" }}>{qaAnalysis.metrics.accuracyRatio.toFixed(2)}%</p>
              </div>
              <div style={{ border: "1px solid #cbd5e1", borderRadius: "8px", padding: "16px", backgroundColor: "#f8fafc" }}>
                <span style={{ display: "block", fontSize: "10px", color: "#94a3b8", textTransform: "uppercase", fontWeight: "bold", fontFamily: "monospace" }}>Worksheets Compared</span>
                <p style={{ fontSize: "18px", fontWeight: "bold", color: "#1e293b", margin: "4px 0 0 0", fontFamily: "monospace" }}>{qaAnalysis.sheetMatches?.length || 0}</p>
              </div>
              <div style={{ border: "1px solid #cbd5e1", borderRadius: "8px", padding: "16px", backgroundColor: "#f8fafc" }}>
                <span style={{ display: "block", fontSize: "10px", color: "#94a3b8", textTransform: "uppercase", fontWeight: "bold", fontFamily: "monospace" }}>Data Errors</span>
                <p style={{ fontSize: "18px", fontWeight: "bold", color: "#1e293b", margin: "4px 0 0 0", fontFamily: "monospace" }}>{qaAnalysis.errorLog.length} errors</p>
              </div>
              <div style={{ border: "1px solid #cbd5e1", borderRadius: "8px", padding: "16px", backgroundColor: "#f8fafc" }}>
                <span style={{ display: "block", fontSize: "10px", color: "#94a3b8", textTransform: "uppercase", fontWeight: "bold", fontFamily: "monospace" }}>Penalty Points</span>
                <p style={{ fontSize: "18px", fontWeight: "bold", color: "#1e293b", margin: "4px 0 0 0", fontFamily: "monospace" }}>{qaAnalysis.metrics.totalStructuralPenalty || 0} pts</p>
              </div>
            </div>

            {/* Error logs */}
            <div style={{ marginBottom: "32px" }}>
              <span style={{ display: "block", fontSize: "12px", fontWeight: "bold", color: "#64748b", textTransform: "uppercase", borderBottom: "1px solid #cbd5e1", paddingBottom: "8px", marginBottom: "16px", fontFamily: "monospace", letterSpacing: "1px" }}>Top Logged Transcription Errors</span>
              {qaAnalysis.errorLog.length === 0 ? (
                <p style={{ fontSize: "12px", color: "#94a3b8", fontStyle: "italic", margin: 0 }}>No errors logged during this assessment cycle.</p>
              ) : (
                <div style={{ display: "flex", flexDirection: "column", gap: "8px" }}>
                  {qaAnalysis.errorLog.slice(0, 15).map((e: any, idx: number) => (
                    <div key={idx} style={{ border: "1px solid #e2e8f0", borderRadius: "6px", backgroundColor: "#f8fafc", padding: "12px", display: "flex", justifyContent: "space-between", alignItems: "center", boxSizing: "border-box" }}>
                      <div style={{ flex: 1, paddingRight: "16px" }}>
                        <span style={{ fontWeight: "bold", color: "#334155", fontSize: "12px" }}>Sheet: {e.sheet} | Cell: {e.cell}</span> <span style={{ fontSize: "12px", color: "#64748b" }}>- {e.errorType}</span>
                        <div style={{ color: "#64748b", fontSize: "11px", marginTop: "4px" }}>Given: "{e.employeeValue}" &bull; Replaced by target: "{e.reviewerValue}"</div>
                      </div>
                      <span style={{ color: "#e11d48", backgroundColor: "#fff1f2", border: "1px solid #fecdd3", padding: "2px 8px", borderRadius: "4px", fontSize: "11px", fontFamily: "monospace", fontWeight: "bold", flexShrink: 0, display: "inline-block" }}>- {e.penalty} pts</span>
                    </div>
                  ))}
                </div>
              )}
            </div>

            {/* Signatures */}
            <div style={{ display: "flex", justifyContent: "space-between", marginTop: "80px", borderTop: "1px solid #cbd5e1", paddingTop: "20px" }}>
              <div style={{ width: "200px", textAlign: "center", fontSize: "12px", color: "#475569", fontWeight: "600" }}>
                <div style={{ height: "40px" }}></div>
                Reviewing Compliance Supervisor
              </div>
              <div style={{ width: "200px", textAlign: "center", fontSize: "12px", color: "#475569", fontWeight: "600" }}>
                <div style={{ height: "40px" }}></div>
                Ground Truth Verification Officer
              </div>
              <div style={{ width: "200px", textAlign: "center", fontSize: "12px", color: "#475569", fontWeight: "600" }}>
                <div style={{ color: "#4f46e5", fontStyle: "italic", fontWeight: "bold", fontSize: "14px", height: "16px", marginBottom: "24px" }}>Verified & Signed</div>
                Audit Authentication Stamp
              </div>
            </div>
          </div>
        </>
      )}

      {/* ⚠️ Beautiful custom Print Dialog for Sandboxed Frame Fallbacks */}
      {showPrintModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 z-50 animate-fade-in font-sans">
          <div className="bg-white border border-slate-200 rounded-2xl w-full max-w-xl shadow-2xl overflow-hidden animate-scale-up">
            
            {/* Modal title header */}
            <div className="p-6 bg-slate-50 border-b border-slate-150 flex items-center justify-between">
              <div className="flex items-center gap-2.5">
                <div className="p-2 bg-indigo-50 rounded bg-indigo-100/60">
                  <Printer className="w-5 h-5 text-indigo-600" />
                </div>
                <div>
                  <h3 className="font-bold text-slate-900 font-sans text-sm">Print Compliance Framework</h3>
                  <p className="text-3xs text-slate-500 font-mono uppercase tracking-wider mt-0.5">Dual Sandbox Mitigation Protocol Active</p>
                </div>
              </div>
              <button 
                onClick={() => setShowPrintModal(false)}
                className="w-7 h-7 flex items-center justify-center rounded-full bg-slate-100 hover:bg-slate-200 text-slate-500 hover:text-slate-700 cursor-pointer transition text-xs font-bold"
              >
                ✕
              </button>
            </div>

            {/* Modal Body */}
            <div className="p-6 space-y-6">
              
              <div className="space-y-1">
                <span className="text-[9px] font-bold text-indigo-600 font-mono bg-indigo-55 bg-indigo-50 px-2 py-0.5 rounded border border-indigo-100 uppercase">
                  Sandbox Active Frame Notice
                </span>
                <p className="text-xs text-slate-600 leading-relaxed font-sans mt-2">
                  Direct browser printing triggers are limited inside secure embedded iframe environments. To ensure your digital record is legally certified, please choose from the options below:
                </p>
              </div>

              {/* Option 2 Card */}
              <div className="border border-indigo-100 bg-indigo-50/25 p-5 rounded-xl space-y-3 transition hover:bg-indigo-50/50">
                <div className="flex items-center gap-2">
                  <span className="w-5 h-5 rounded-full bg-indigo-600 text-white font-mono text-[10px] flex items-center justify-center font-bold">1</span>
                  <h4 className="text-xs font-bold text-slate-800 font-sans">Download Certified compliance Report as offline PDF</h4>
                </div>
                <p className="text-xs text-slate-550 leading-relaxed pl-7">
                  Click below to immediately compile, render, and download a multi-page PDF document conforming to quality and alignment review standards.
                </p>
                <div className="pl-7 pt-2">
                  <button
                    onClick={generatePDFCertificate}
                    className="inline-flex items-center gap-1.5 px-4 py-2 bg-indigo-650 hover:bg-indigo-700 text-white text-[10px] font-semibold uppercase tracking-wider font-mono rounded cursor-pointer transition shadow-xs"
                  >
                    <Download className="w-3.5 h-3.5" /> Generate printable PDF Document (.pdf)
                  </button>
                </div>
              </div>

              {/* Option 1 Card */}
              <div className="border border-slate-200 bg-slate-50/50 p-5 rounded-xl space-y-3 hover:bg-slate-50 transition">
                <div className="flex items-center gap-2">
                  <span className="w-5 h-5 rounded-full bg-slate-600 text-white font-mono text-[10px] flex items-center justify-center font-bold">2</span>
                  <h4 className="text-xs font-bold text-slate-800 pr-2">Launch Workspace app in clean tab</h4>
                </div>
                <p className="text-xs text-slate-500 leading-relaxed pl-7">
                  Copy this sandbox address and paste it inside your browser address bar. Your browser will run the portal natively, resolving all iframe print issues:
                </p>
                <div className="pl-7 space-y-2">
                  <div className="flex gap-2">
                    <input
                      type="text"
                      readOnly
                      value={window.location.href}
                      className="flex-1 bg-white border border-slate-200 px-3 py-1.5 rounded text-3xs font-mono text-slate-500 outline-none select-all"
                    />
                    <button
                      onClick={copySandboxUrlToClipboard}
                      className="px-3.5 py-1.5 bg-slate-850 hover:bg-slate-800 text-white text-3xs font-semibold uppercase font-mono tracking-wider rounded cursor-pointer transition whitespace-nowrap min-w-[90px] text-center"
                    >
                      {copiedUrl ? "Copied!" : "Copy Link"}
                    </button>
                  </div>
                </div>
              </div>

            </div>

            {/* Modal Footer */}
            <div className="px-6 py-4 bg-slate-50 border-t border-slate-150 flex justify-end gap-2.5">
              <button
                onClick={() => setShowPrintModal(false)}
                className="px-4 py-2 bg-slate-200 hover:bg-slate-300 text-slate-705 text-3xs font-bold uppercase tracking-wider font-mono rounded cursor-pointer transition"
              >
                Dismiss Modal
              </button>
            </div>

          </div>
        </div>
      )}

    </div>
  );
}
