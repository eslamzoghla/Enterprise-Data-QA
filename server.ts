import express from "express";
import path from "path";
import { createServer as createViteServer } from "vite";
import { GoogleGenAI, Type } from "@google/genai";
import dotenv from "dotenv";

dotenv.config();

const app = express();
const PORT = 3000;

// Enable JSON bodies with higher limits for Excel payloads
app.use(express.json({ limit: "50mb" }));

// Endpoint to find Google Client ID or other config from env
app.get("/api/config", (req, res) => {
  res.json({
    googleClientId: process.env.GOOGLE_CLIENT_ID || process.env.OAUTH_CLIENT_ID || process.env.CLIENT_ID || "",
    envKeys: Object.keys(process.env).filter(k => !k.includes("KEY") && !k.includes("SECRET") && !k.includes("PASSWORD"))
  });
});

// Secure API endpoint for Gemini-powered evaluation analysis
app.post("/api/coach", async (req, res) => {
  try {
    const { metrics, rootCause, topErrors, patterns, workerName, projectName, date } = req.body;

    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey) {
      return res.status(200).json({
        summary: `### Executive Audit Summary\n\nEmployee **${workerName || "N/A"}** has been evaluated for **${projectName || "General Data Ingress Operations"}** on **${date || "Current Session"}**.\n\n* **Final Grade**: \`${metrics.finalGrade}\`\n* **Accuracy**: \`${metrics.baseAccuracy}%\`\n* **Estimated Workload Burden**: \`${metrics.reviewerWorkloadIndex}\` points.\n\n*Auditor Cabin Notice: Server-side \`GEMINI_API_KEY\` is not configured in the workspace secrets. Executive Audit summaries are generated locally using the deterministic evaluation rules. Enter your Gemini API key in the Secrets Panel to activate the automated narrative assessment.*`,
        coaching: [
          "Review all cell alignments and coordinate indexes prior to final database submittals to avoid structural displacements.",
          "Activate standard spelling autocomplete and verify double-letter Arabic spelling rules for keys such as 'ة' vs 'ه'.",
          "Conduct high-precision visual scanning on floating-decimals and values containing variable numbers of digits."
        ]
      });
    }

    const ai = new GoogleGenAI({
      apiKey,
      httpOptions: {
        headers: {
          'User-Agent': 'aistudio-build',
        }
      }
    });

    const prompt = `
You are an Senior Enterprise Data Quality Assurance Auditor. Your task is to evaluate a data entry employee's performance based on their QA comparison report.

Review standard details:
- Employee Name: ${workerName || "Unnamed Employee"}
- Project Name: ${projectName || "Unnamed Project"}
- Evaluation Date: ${date}

Accuracy Metrics:
- Cell Coordinates Compared: ${metrics.comparedCells}
- Distinct Error Mismatches: ${metrics.totalErrors}
- Penalty Points Accumulation: ${metrics.totalPenaltyPoints}
- Accuracy: ${metrics.baseAccuracy}%
- Error Rate (per 10k cells): ${metrics.errorRatePer10k}
- Reviewer Correction Burden Score (RWI): ${metrics.reviewerWorkloadIndex}
- Final Earned Quality Grade: ${metrics.finalGrade}

Root Cause Attribution Distribution:
- Missing Value Errors (omissions): ${rootCause.missingValuesPct}%
- Extraneous Cell Entries (extras): ${rootCause.extraValuesPct}%
- Numeric Digit Discrepancies (typos, substitutions, major numeric errors): ${rootCause.numericErrorsPct}%
- Text / Spelling / Similarity Typos: ${rootCause.textErrorsPct}%
- Row or Column Layout Shift Events: ${rootCause.shiftErrorsPct}%

Discrepancy Patterns & Structural Findings:
- Repeated numeric substitutions: ${JSON.stringify(patterns.repeatedNumericErrors)}
- Copy-paste repetition anomalies: ${JSON.stringify(patterns.copyPasteErrors)}
- Dense error clusters: ${JSON.stringify(patterns.errorClusters)}
- Registered Shift-Precedence events: ${JSON.stringify(patterns.shiftEvents)}

Representative Sample Log entries:
${JSON.stringify(topErrors)}

Formulate a highly sophisticated professional auditor assessment. The overview must details exactly why the employee made these errors (e.g. shift misalignment, confusion in Arabic characters, missing numeric digits) and what the high-impact failures imply for business integrity. Detail how structural shift events (if any) significantly impacted their final grade according to enterprise QA constraints. 

Provide:
1. "summary" (written in pristine markdown with clean sections, crisp points, bullet lists, no marketing jargon or self-praise, keeping the voice serious, authoritative, and helpful).
2. "coaching" (a list of precise, measurable, and highly logical actions targeting their largest failure vectors).
`;

    const response = await ai.models.generateContent({
      model: "gemini-3.5-flash",
      contents: prompt,
      config: {
        responseMimeType: "application/json",
        responseSchema: {
          type: Type.OBJECT,
          properties: {
            summary: { type: Type.STRING, description: "Markdown text analyzing worker performance and root-cause drivers." },
            coaching: {
              type: Type.ARRAY,
              items: { type: Type.STRING },
              description: "Actionable concrete training steps to improve accuracy."
            }
          },
          required: ["summary", "coaching"]
        }
      }
    });

    const parsedData = JSON.parse(response.text || "{}");
    res.json(parsedData);
  } catch (error: any) {
    console.error("Gemini route error:", error);
    res.status(500).json({ error: "Failed to generate AI audit analytics", details: error.message });
  }
});

// Configure Vite integration or serve static bundles
async function startServer() {
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Enterprise QA server running on port ${PORT}`);
  });
}

startServer();
