const state = {
  files: [],
  documents: [],
  sourceCsv: "",
  resultCsv: "",
  resultRows: [],
  activeTypingRun: 0,
  progressTimer: null
};

const elements = {
  dropzone: document.getElementById("dropzone"),
  pdfInput: document.getElementById("pdfInput"),
  processButton: document.getElementById("processBtn"),
  debugButton: document.getElementById("debugBtn"),
  downloadSourceButton: document.getElementById("downloadSourceBtn"),
  downloadResultButton: document.getElementById("downloadResultBtn"),
  progressLabel: document.getElementById("progressLabel"),
  progressValue: document.getElementById("progressValue"),
  progressFill: document.getElementById("progressFill"),
  status: document.getElementById("status"),
  fileList: document.getElementById("fileList"),
  selectionBadge: document.getElementById("selectionBadge"),
  summaryBox: document.getElementById("summaryBox"),
  docCount: document.getElementById("docCount"),
  sourceRowCount: document.getElementById("sourceRowCount"),
  resultRowCount: document.getElementById("resultRowCount"),
  totalHours: document.getElementById("totalHours"),
  sourcePreview: document.getElementById("sourcePreview"),
  resultPreview: document.getElementById("resultPreview")
};

const pdfLibrary = window["pdfjs-dist/build/pdf"];

if (pdfLibrary) {
  pdfLibrary.GlobalWorkerOptions.workerSrc =
    "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
}

elements.pdfInput.addEventListener("change", (event) => {
  setSelectedFiles(Array.from(event.target.files || []));
});
elements.processButton.addEventListener("click", processInvoices);
elements.debugButton.addEventListener("click", checkBackendConfig);
elements.downloadSourceButton.addEventListener("click", () => {
  downloadCsv(state.sourceCsv, "invoice-source.csv");
});
elements.downloadResultButton.addEventListener("click", () => {
  downloadCsv(state.resultCsv, "invoice-hours.csv");
});

["dragenter", "dragover"].forEach((eventName) => {
  elements.dropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    elements.dropzone.classList.add("is-dragover");
  });
});

["dragleave", "drop"].forEach((eventName) => {
  elements.dropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    elements.dropzone.classList.remove("is-dragover");
  });
});

elements.dropzone.addEventListener("drop", (event) => {
  const transferFiles = event.dataTransfer && event.dataTransfer.files ? event.dataTransfer.files : [];
  const droppedFiles = Array.from(transferFiles).filter(
    (file) => file.type === "application/pdf" || file.name.toLowerCase().endsWith(".pdf")
  );
  if (droppedFiles.length) {
    setSelectedFiles(droppedFiles);
    syncInputFiles(droppedFiles);
  }
});

elements.dropzone.addEventListener("keydown", (event) => {
  if (event.key === "Enter" || event.key === " ") {
    event.preventDefault();
    elements.pdfInput.click();
  }
});

renderSelection();
renderMetrics();

function setSelectedFiles(files) {
  state.files = files;
  state.documents = [];
  state.sourceCsv = "";
  state.resultCsv = "";
  state.resultRows = [];
  stopProgressSimulation();
  resetOutputViews();
  renderSelection();
  renderMetrics();
  setProgress(0, "Ready to process invoice PDFs.");

  if (!files.length) {
    setStatus("No PDFs selected yet. The intermediate CSV is generated in the browser before the secure Gemini request runs.");
    return;
  }

  const totalSize = files.reduce((sum, file) => sum + file.size, 0);
  setStatus(
    `${files.length} PDF${files.length === 1 ? "" : "s"} selected. Total size: ${formatBytes(totalSize)}.`
  );
  animateInvoiceFields([
    elements.status,
    elements.selectionBadge,
    elements.docCount,
    elements.sourceRowCount,
    elements.resultRowCount,
    elements.totalHours
  ]);
}

function syncInputFiles(files) {
  const dataTransfer = new DataTransfer();
  files.forEach((file) => dataTransfer.items.add(file));
  elements.pdfInput.files = dataTransfer.files;
}

function resetOutputViews() {
  elements.summaryBox.textContent =
    "The final summary will appear here after Gemini processes the source CSV.";
  elements.summaryBox.classList.remove("has-result", "typewriter-caret");
  elements.sourcePreview.className = "empty";
  elements.sourcePreview.textContent =
    "The source CSV preview will appear after text extraction starts.";
  elements.resultPreview.className = "empty";
  elements.resultPreview.textContent =
    "The final CSV preview will appear here after Gemini returns the structured rows.";
  elements.downloadSourceButton.disabled = true;
  elements.downloadResultButton.disabled = true;
}

function renderSelection() {
  const files = state.files;
  elements.selectionBadge.textContent = `${files.length} file${files.length === 1 ? "" : "s"} selected`;
  elements.processButton.disabled = !files.length;
  elements.fileList.className = files.length ? "file-list" : "file-list empty";

  if (!files.length) {
    elements.fileList.textContent = "Add invoice PDFs to see the processing queue.";
    return;
  }

  const fragment = document.createDocumentFragment();

  files.forEach((file) => {
    const item = document.createElement("article");
    item.className = "file-card";

    const copy = document.createElement("div");
    const title = document.createElement("p");
    title.className = "file-name";
    title.textContent = file.name;

    const meta = document.createElement("p");
    meta.className = "file-meta";
    meta.textContent = `${formatBytes(file.size)} • ${file.type || "application/pdf"}`;

    copy.append(title, meta);

    const badge = document.createElement("p");
    badge.className = "badge";
    badge.textContent = "Queued";

    item.append(copy, badge);
    fragment.append(item);
  });

  elements.fileList.replaceChildren(fragment);
}

function renderMetrics() {
  elements.docCount.textContent = String(state.files.length);
  elements.sourceRowCount.textContent = state.documents.length ? String(state.documents.length) : "0";
  elements.resultRowCount.textContent = state.resultRows.length ? String(state.resultRows.length) : "0";
  const totalHours = state.resultRows.reduce((sum, row) => sum + getHoursNumber(row["Hours Billed"]), 0);
  elements.totalHours.textContent = totalHours.toFixed(1);
}

async function processInvoices() {
  if (!state.files.length) {
    return;
  }

  if (window.location.protocol === "file:") {
    setStatus(
      "The Gemini step is unavailable in file preview mode. Open this page through Vercel or another local server with the /api/invoices-to-excel endpoint enabled."
    );
    return;
  }

  if (!pdfLibrary) {
    setStatus("PDF.js could not be loaded. Refresh the page and try again.");
    return;
  }

  elements.processButton.disabled = true;
  elements.downloadSourceButton.disabled = true;
  elements.downloadResultButton.disabled = true;
  setStatus("Extracting text from uploaded invoice PDFs...");

  try {
    const documents = [];

    for (let index = 0; index < state.files.length; index += 1) {
      const file = state.files[index];
      setProgress(
        Math.max(5, Math.round(((index + 0.2) / state.files.length) * 68)),
        `Reading ${file.name} (${index + 1} of ${state.files.length})`
      );

      const documentData = await extractPdfDocument(file);
      documents.push(documentData);
      state.documents = documents.slice();
      renderSourcePreview();
      renderMetrics();
    }

    state.sourceCsv = buildCsv(
      state.documents.map((documentData) => ({
        "Document Name": documentData.documentName,
        Content: documentData.content
      }))
    );

    renderSourcePreview();
    renderMetrics();
    animateInvoiceFields([
      elements.sourceRowCount,
      elements.status
    ]);
    elements.downloadSourceButton.disabled = false;

    setProgress(78, "Source CSV created. Sending it to Gemini...");
    setStatus("The source CSV is ready. Gemini is now structuring user names and billed hours.");
    startProgressSimulation(78, 96);

    const response = await fetch("./api/invoices-to-excel", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        sourceCsv: state.sourceCsv,
        documents: state.documents.map((documentData) => ({
          documentName: documentData.documentName
        }))
      })
    });

    stopProgressSimulation();

    if (!response.ok) {
      const errorPayload = await readJsonSafely(response);
      throw new Error((errorPayload && errorPayload.error) || "The Gemini request failed.");
    }

    const payload = await response.json();
    state.resultRows = Array.isArray(payload.rows) ? payload.rows : [];
    state.resultCsv = typeof payload.csv === "string" ? payload.csv : "";
    renderResultPreview();
    renderMetrics();

    elements.downloadResultButton.disabled = !state.resultCsv;
    setProgress(100, "Final CSV ready.");
    setStatus(
      `Completed successfully. ${state.resultRows.length} structured row${
        state.resultRows.length === 1 ? "" : "s"
      } ready for preview and download.`
    );
    typewriteInto(elements.summaryBox, payload.summary || "Gemini finished processing the invoice batch.");
    elements.summaryBox.classList.add("has-result");
    animateInvoiceFields([
      elements.status,
      elements.resultRowCount,
      elements.totalHours
    ]);
  } catch (error) {
    console.error(error);
    stopProgressSimulation();
    setProgress(0, "Ready to process invoice PDFs.");

    if (error && error.message === "Failed to fetch") {
      setStatus(
        "The invoice batch could not be processed because the Gemini backend is not reachable. Run this page in Vercel or another server environment where /api/invoices-to-excel is available."
      );
      return;
    }

    setStatus(`The invoice batch could not be processed. ${error.message || "Please try again."}`);
  } finally {
    elements.processButton.disabled = !state.files.length;
  }
}

function animateInvoiceFields(targets) {
  const runId = Date.now();
  state.activeTypingRun = runId;

  const uniqueTargets = targets
    .filter(Boolean)
    .filter((element, index, list) => list.indexOf(element) === index)
    .filter((element) => element.textContent.trim() !== "");

  uniqueTargets.forEach((element) => {
    element.dataset.typewriterText = element.textContent;
    element.classList.add("typewriter-pending");
  });

  if (window.matchMedia("(prefers-reduced-motion: reduce)").matches) {
    uniqueTargets.forEach((element) => {
      element.classList.remove("typewriter-pending");
    });
    return;
  }

  let delay = 0;
  uniqueTargets.forEach((element) => {
    const lengthFactor = Math.min(element.dataset.typewriterText.length * 8, 320);
    window.setTimeout(() => {
      typewriteElementWithRun(element, element.dataset.typewriterText, runId);
    }, delay);
    delay += Math.min(28 + lengthFactor * 0.05, 70);
  });
}

function typewriteElementWithRun(element, text, runId) {
  if (state.activeTypingRun !== runId) {
    return;
  }

  element.classList.remove("typewriter-pending");
  element.textContent = "";
  element.classList.add("typewriter-caret");

  const characters = Array.from(text);
  const step = () => {
    if (state.activeTypingRun !== runId) {
      element.classList.remove("typewriter-pending");
      element.classList.remove("typewriter-caret");
      element.textContent = text;
      return;
    }

    const nextCharacter = characters.shift();
    if (nextCharacter === undefined) {
      window.setTimeout(() => {
        if (state.activeTypingRun === runId) {
          element.classList.remove("typewriter-caret");
        }
      }, 120);
      return;
    }

    element.textContent += nextCharacter;
    window.setTimeout(step, nextCharacter === " " ? 4 : 5);
  };

  step();
}

async function checkBackendConfig() {
  if (window.location.protocol === "file:") {
    setStatus(
      "Backend config cannot be checked in file preview mode. Open this page through Vercel or another environment where /api/invoices-to-excel is available."
    );
    return;
  }

  elements.debugButton.disabled = true;

  try {
    const response = await fetch("./api/invoices-to-excel", {
      method: "GET",
      headers: {
        Accept: "application/json"
      }
    });

    const payload = await readJsonSafely(response);

    if (!response.ok) {
      throw new Error((payload && payload.error) || "Could not read backend config.");
    }

    setStatus(
      `Backend OK. Model: ${payload.model || "unknown"}. API key present: ${
        payload.hasApiKey ? "yes" : "no"
      }. Key suffix: ${payload.apiKeySuffix || "n/a"}.`
    );
  } catch (error) {
    console.error(error);
    setStatus(`Backend config check failed. ${error.message || "Please try again."}`);
  } finally {
    elements.debugButton.disabled = false;
  }
}

async function extractPdfDocument(file) {
  const buffer = await file.arrayBuffer();
  const loadingTask = pdfLibrary.getDocument({ data: new Uint8Array(buffer) });
  const pdf = await loadingTask.promise;
  const pages = [];

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const textContent = await page.getTextContent();
    const pageText = textContent.items.map((item) => item.str).join(" ").trim();
    pages.push(pageText);
  }

  const content = pages
    .join("\n")
    .replace(/\s+\n/g, "\n")
    .replace(/\n\s+/g, "\n")
    .replace(/[ \t]{2,}/g, " ")
    .trim();

  return {
    documentName: file.name,
    content: content || "No machine-readable text detected in this PDF."
  };
}

function renderSourcePreview() {
  if (!state.documents.length) {
    elements.sourcePreview.className = "empty";
    elements.sourcePreview.textContent =
      "The source CSV preview will appear after text extraction starts.";
    return;
  }

  const table = document.createElement("table");
  const thead = document.createElement("thead");
  const tbody = document.createElement("tbody");
  const headerRow = document.createElement("tr");

  ["Document Name", "Content Preview"].forEach((label) => {
    const th = document.createElement("th");
    th.textContent = label;
    headerRow.append(th);
  });

  thead.append(headerRow);

  state.documents.forEach((documentData) => {
    const row = document.createElement("tr");
    const nameCell = document.createElement("td");
    nameCell.textContent = documentData.documentName;

    const contentCell = document.createElement("td");
    contentCell.className = "truncate-cell";
    const contentSpan = document.createElement("span");
    contentSpan.title = documentData.content;
    contentSpan.textContent = documentData.content;
    contentCell.append(contentSpan);

    row.append(nameCell, contentCell);
    tbody.append(row);
  });

  table.append(thead, tbody);
  elements.sourcePreview.className = "";
  elements.sourcePreview.replaceChildren(table);
}

function renderResultPreview() {
  if (!state.resultRows.length) {
    elements.resultPreview.className = "empty";
    elements.resultPreview.textContent =
      "The final CSV preview will appear here after Gemini returns the structured rows.";
    return;
  }

  const table = document.createElement("table");
  const thead = document.createElement("thead");
  const tbody = document.createElement("tbody");
  const headerRow = document.createElement("tr");
  const columns = ["Document Name", "User Name", "Hours Billed"];

  columns.forEach((label) => {
    const th = document.createElement("th");
    th.textContent = label;
    headerRow.append(th);
  });

  thead.append(headerRow);

  state.resultRows.forEach((resultRow) => {
    const row = document.createElement("tr");

    columns.forEach((column) => {
      const cell = document.createElement("td");
      const value = Object.prototype.hasOwnProperty.call(resultRow, column) ? resultRow[column] : "";
      cell.textContent = String(value);
      row.append(cell);
    });

    tbody.append(row);
  });

  table.append(thead, tbody);
  elements.resultPreview.className = "";
  elements.resultPreview.replaceChildren(table);
}

function setProgress(value, label) {
  const boundedValue = Math.max(0, Math.min(100, Math.round(value)));
  elements.progressFill.style.width = `${boundedValue}%`;
  elements.progressValue.textContent = `${boundedValue}%`;
  elements.progressLabel.textContent = label;
}

function startProgressSimulation(from, to) {
  stopProgressSimulation();
  let currentValue = from;
  state.progressTimer = window.setInterval(() => {
    currentValue = Math.min(to, currentValue + Math.random() * 2.4);
    setProgress(currentValue, "Gemini is reviewing invoice text and preparing the final CSV...");
    if (currentValue >= to) {
      stopProgressSimulation();
    }
  }, 220);
}

function stopProgressSimulation() {
  if (!state.progressTimer) {
    return;
  }

  window.clearInterval(state.progressTimer);
  state.progressTimer = null;
}

function setStatus(message) {
  elements.status.textContent = message;
}

function typewriteInto(element, text) {
  state.activeTypingRun += 1;
  const runId = state.activeTypingRun;
  element.textContent = "";
  element.classList.add("typewriter-caret");

  let index = 0;

  const tick = () => {
    if (runId !== state.activeTypingRun) {
      return;
    }

    if (index >= text.length) {
      element.classList.remove("typewriter-caret");
      return;
    }

    element.textContent += text.charAt(index);
    index += 1;
    window.setTimeout(tick, 14);
  };

  tick();
}

function buildCsv(rows) {
  if (!rows.length) {
    return "";
  }

  const headers = Object.keys(rows[0]);
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

function getHoursNumber(value) {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }

  const parsed = Number(
    String(value === undefined || value === null ? "" : value).replace(",", ".").replace(/[^\d.-]/g, "")
  );
  return Number.isFinite(parsed) ? parsed : 0;
}

function downloadCsv(content, fileName) {
  if (!content) {
    return;
  }

  const blob = new Blob([content], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.append(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function formatBytes(bytes) {
  if (!Number.isFinite(bytes) || bytes <= 0) {
    return "0 B";
  }

  const units = ["B", "KB", "MB", "GB"];
  const unitIndex = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
  const value = bytes / 1024 ** unitIndex;
  return `${value.toFixed(value >= 10 || unitIndex === 0 ? 0 : 1)} ${units[unitIndex]}`;
}

async function readJsonSafely(response) {
  try {
    return await response.json();
  } catch {
    return null;
  }
}
