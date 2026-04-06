const DEFAULT_MODEL = process.env.GEMINI_MODEL || "gemini-2.0-flash";

async function handler(req, res) {
  if (req.method === "GET") {
    const apiKey = process.env.GEMINI_API_KEY || "";
    return res.status(200).json({
      ok: true,
      model: DEFAULT_MODEL,
      hasApiKey: Boolean(apiKey),
      apiKeySuffix: apiKey ? apiKey.slice(-4) : "",
      timestamp: new Date().toISOString()
    });
  }

  if (req.method !== "POST") {
    res.setHeader("Allow", "GET, POST");
    return res.status(405).json({ error: "Method not allowed." });
  }

  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    return res.status(500).json({
      error: "Missing GEMINI_API_KEY. Add it as a Vercel environment variable before using this endpoint."
    });
  }

  const requestBody = req && req.body ? req.body : {};
  const sourceCsv = typeof requestBody.sourceCsv === "string" ? requestBody.sourceCsv.trim() : "";
  const inputDocuments = Array.isArray(requestBody.documents) ? requestBody.documents : [];

  if (!sourceCsv || !inputDocuments.length) {
    return res.status(400).json({
      error: "Missing source CSV payload or document list."
    });
  }

  try {
    const geminiResponse = await fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/${DEFAULT_MODEL}:generateContent?key=${encodeURIComponent(apiKey)}`,
      {
        method: "POST",
        headers: {
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          generationConfig: {
            temperature: 0.1,
            responseMimeType: "application/json"
          },
          contents: [
            {
              role: "user",
              parts: [
                {
                  text: buildPrompt(sourceCsv)
                }
              ]
            }
          ]
        })
      }
    );

    if (!geminiResponse.ok) {
      const errorText = await geminiResponse.text();
      return res.status(502).json({
        error: `Gemini request failed: ${errorText || geminiResponse.statusText}`
      });
    }

    const payload = await geminiResponse.json();
    const responseText = extractGeminiText(payload);

    if (!responseText) {
      return res.status(502).json({
        error: "Gemini returned an empty response."
      });
    }

    const parsed = parseJsonResponse(responseText);
    const normalizedRows = normalizeRows(parsed && parsed.rows, inputDocuments);
    const csv = buildCsv(normalizedRows);
    const summary =
      typeof (parsed && parsed.summary) === "string" && parsed.summary.trim()
        ? parsed.summary.trim()
        : fallbackSummary(normalizedRows);

    return res.status(200).json({
      summary,
      rows: normalizedRows,
      csv
    });
  } catch (error) {
    console.error("Invoice extraction endpoint failed.", error);
    return res.status(500).json({
      error: error.message || "The invoice extraction endpoint failed."
    });
  }
}

module.exports = handler;
module.exports.config = {
  api: {
    bodyParser: {
      sizeLimit: "8mb"
    }
  }
};

function buildPrompt(sourceCsv) {
  return [
    "You transform developer invoice text into structured CSV-ready rows.",
    "The input is a CSV with these columns: Document Name, Content.",
    "Return JSON only with two top-level keys: summary and rows.",
    'summary must be one or two short English sentences.',
    'rows must be an array of objects with exactly these keys: "Document Name", "User Name", "Hours Billed".',
    "Keep one row per input document.",
    "Use the original document name from the input whenever possible.",
    "Hours Billed must be numeric, not a string with units.",
    "If a document does not clearly show a billed user, leave User Name as an empty string.",
    "If a document does not clearly show billed hours, use 0.",
    "Do not invent documents that are not present in the input.",
    "",
    "Input CSV:",
    sourceCsv
  ].join("\n");
}

function extractGeminiText(payload) {
  const candidate = payload && payload.candidates ? payload.candidates[0] : null;
  const parts = candidate && candidate.content ? candidate.content.parts : null;

  if (!Array.isArray(parts)) {
    return "";
  }

  return parts
    .map((part) => (part && typeof part.text === "string" ? part.text : ""))
    .join("")
    .trim();
}

function parseJsonResponse(text) {
  const cleaned = text
    .replace(/^```json\s*/i, "")
    .replace(/^```\s*/i, "")
    .replace(/\s*```$/i, "")
    .trim();

  return JSON.parse(cleaned);
}

function normalizeRows(rows, inputDocuments) {
  const safeRows = Array.isArray(rows) ? rows : [];
  const inputNames = inputDocuments.map((documentData) =>
    String(documentData && documentData.documentName ? documentData.documentName : "").trim()
  );
  const rowMap = new Map();

  safeRows.forEach((row) => {
    const documentName = String(
      (row && row["Document Name"]) || (row && row.documentName) || ""
    ).trim();
    if (!documentName) {
      return;
    }

    rowMap.set(normalizeKey(documentName), {
      "Document Name": documentName,
      "User Name": String((row && row["User Name"]) || (row && row.userName) || "").trim(),
      "Hours Billed": toNumber(
        row && row["Hours Billed"] !== undefined ? row["Hours Billed"] : row && row.hoursBilled
      )
    });
  });

  return inputNames.map((inputName) => {
    const matched = rowMap.get(normalizeKey(inputName));
    return {
      "Document Name": inputName,
      "User Name": (matched && matched["User Name"]) || "",
      "Hours Billed": matched && matched["Hours Billed"] !== undefined ? matched["Hours Billed"] : 0
    };
  });
}

function normalizeKey(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/\.[^.]+$/, "")
    .replace(/[^a-z0-9]+/g, "");
}

function toNumber(value) {
  if (typeof value === "number" && Number.isFinite(value)) {
    return Number(value.toFixed(2));
  }

  const parsed = Number(
    String(value === undefined || value === null ? "" : value).replace(",", ".").replace(/[^\d.-]/g, "")
  );
  return Number.isFinite(parsed) ? Number(parsed.toFixed(2)) : 0;
}

function buildCsv(rows) {
  if (!rows.length) {
    return "Document Name,User Name,Hours Billed";
  }

  const headers = ["Document Name", "User Name", "Hours Billed"];
  const lines = [
    headers.map(escapeCsvCell).join(","),
    ...rows.map((row) => headers.map((header) => escapeCsvCell(row[header])).join(","))
  ];

  return lines.join("\n");
}

function escapeCsvCell(value) {
  const stringValue = String(value === undefined || value === null ? "" : value);
  if (/[",\n]/.test(stringValue)) {
    return `"${stringValue.replace(/"/g, '""')}"`;
  }

  return stringValue;
}

function fallbackSummary(rows) {
  const totalHours = rows.reduce((sum, row) => sum + toNumber(row["Hours Billed"]), 0);
  return `Gemini processed ${rows.length} invoice ${rows.length === 1 ? "row" : "rows"} and detected ${totalHours.toFixed(1)} billed hours in total.`;
}
