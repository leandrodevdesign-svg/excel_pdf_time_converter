const state = {
  kimaiFile: null,
  invoiceFile: null,
  kimaiUsers: [],
  invoiceGroups: [],
  pairedRows: []
};

const elements = {
  kimaiInput: document.getElementById("kimaiInput"),
  invoiceInput: document.getElementById("invoiceInput"),
  kimaiFilePill: document.getElementById("kimaiFilePill"),
  invoiceFilePill: document.getElementById("invoiceFilePill"),
  compareButton: document.getElementById("compareBtn"),
  status: document.getElementById("status"),
  kimaiUserCount: document.getElementById("kimaiUserCount"),
  invoiceUserCount: document.getElementById("invoiceUserCount"),
  matchedGroupCount: document.getElementById("matchedGroupCount"),
  totalDelta: document.getElementById("totalDelta"),
  comparisonBadge: document.getElementById("comparisonBadge"),
  kimaiTableState: document.getElementById("kimaiTableState"),
  invoiceTableState: document.getElementById("invoiceTableState")
};

elements.kimaiInput.addEventListener("change", handleKimaiFileChange);
elements.invoiceInput.addEventListener("change", handleInvoiceFileChange);
elements.compareButton.addEventListener("click", compareFiles);

renderMetrics();

function handleKimaiFileChange(event) {
  state.kimaiFile = event.target.files && event.target.files[0] ? event.target.files[0] : null;
  state.kimaiUsers = [];
  state.pairedRows = [];
  elements.kimaiFilePill.textContent = state.kimaiFile ? state.kimaiFile.name : "No Kimai file selected";
  updateCompareButton();
  resetTables("Upload both files to render the comparison.");
  renderMetrics();
  setStatus(state.kimaiFile ? "Kimai file loaded into the queue. Add the invoice CSV to compare both sources." : "Waiting for both files. The comparison runs fully in the browser.");
}

function handleInvoiceFileChange(event) {
  state.invoiceFile = event.target.files && event.target.files[0] ? event.target.files[0] : null;
  state.invoiceGroups = [];
  state.pairedRows = [];
  elements.invoiceFilePill.textContent = state.invoiceFile ? state.invoiceFile.name : "No invoice CSV selected";
  updateCompareButton();
  resetTables("Upload both files to render the comparison.");
  renderMetrics();
  setStatus(state.invoiceFile ? "Invoice CSV loaded into the queue. Add the Kimai export to compare both sources." : "Waiting for both files. The comparison runs fully in the browser.");
}

function updateCompareButton() {
  elements.compareButton.disabled = !(state.kimaiFile && state.invoiceFile);
}

async function compareFiles() {
  if (!state.kimaiFile || !state.invoiceFile) {
    return;
  }

  elements.compareButton.disabled = true;
  setStatus("Parsing the Kimai export and the invoice CSV...");

  try {
    const kimaiUsers = await parseKimaiWorkbook(state.kimaiFile);
    const invoiceGroups = await parseInvoiceCsv(state.invoiceFile);
    const comparison = buildComparisonRows(kimaiUsers, invoiceGroups);

    state.kimaiUsers = kimaiUsers;
    state.invoiceGroups = invoiceGroups;
    state.pairedRows = comparison.rows;

    renderMetrics(comparison);
    renderMirrorTables(comparison.rows);
    elements.comparisonBadge.textContent = `${comparison.groupCount} mirrored group${comparison.groupCount === 1 ? "" : "s"}`;
    setStatus(
      `Comparison ready. ${comparison.groupCount} mirrored groups generated and ordered alphabetically by normalized name.`
    );
  } catch (error) {
    console.error(error);
    resetTables("The comparison could not be generated. Check the file formats and try again.");
    setStatus(`The comparison could not be generated. ${error.message || "Please try again."}`);
  } finally {
    updateCompareButton();
  }
}

async function parseKimaiWorkbook(file) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
  const firstSheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[firstSheetName];
  const matrix = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    raw: true,
    defval: ""
  });

  if (!matrix.length) {
    throw new Error("The Kimai file is empty.");
  }

  const headers = matrix[0].map(normalizeHeader);
  const usernameColumnIndex = headers.findIndex((header) => header === "username" || header === "user");
  const hoursColumnIndex = headers.findIndex(
    (header) => header === "time" || header === "hours" || header === "total_hours" || header === "duration"
  );

  if (usernameColumnIndex === -1 || hoursColumnIndex === -1) {
    throw new Error("The Kimai file must include Username and Time columns.");
  }

  const groups = new Map();

  matrix.slice(1).forEach((row) => {
    const rawUsername = String(row[usernameColumnIndex] || "").trim();
    const rawHours = row[hoursColumnIndex];

    if (!rawUsername || /^total$/i.test(rawUsername)) {
      return;
    }

    const normalizedKey = normalizePersonName(rawUsername);
    const hours = getNumber(rawHours);

    if (!normalizedKey || !Number.isFinite(hours) || hours <= 0) {
      return;
    }

    const existing = groups.get(normalizedKey) || {
      normalizedKey,
      displayName: formatDisplayNameFromKimai(rawUsername),
      sourceName: rawUsername,
      totalHours: 0
    };

    existing.totalHours += hours;
    groups.set(normalizedKey, existing);
  });

  return Array.from(groups.values())
    .map((entry) => ({
      normalizedKey: entry.normalizedKey,
      displayName: entry.displayName,
      sourceName: entry.sourceName,
      totalHours: Number(entry.totalHours.toFixed(2))
    }))
    .sort((left, right) => left.displayName.localeCompare(right.displayName));
}

async function parseInvoiceCsv(file) {
  const text = await file.text();
  const rows = parseCsv(text);

  if (!rows.length) {
    throw new Error("The invoice CSV is empty.");
  }

  const headers = rows[0].map(normalizeHeader);
  const userColumnIndex = headers.findIndex((header) => header === "user_name" || header === "username");
  const hoursColumnIndex = headers.findIndex((header) => header === "hours_billed" || header === "hours");
  const documentColumnIndex = headers.findIndex((header) => header === "document_name" || header === "document");

  if (userColumnIndex === -1 || hoursColumnIndex === -1 || documentColumnIndex === -1) {
    throw new Error("The invoice CSV must include Document Name, User Name, and Hours Billed.");
  }

  const groups = new Map();

  rows.slice(1).forEach((row) => {
    const userName = String(row[userColumnIndex] || "").trim();
    const documentName = String(row[documentColumnIndex] || "").trim();
    const hours = getNumber(row[hoursColumnIndex]);
    const normalizedKey = normalizePersonName(userName);

    if (!normalizedKey) {
      return;
    }

    const existing = groups.get(normalizedKey) || {
      normalizedKey,
      displayName: formatDisplayNameFromInvoice(userName),
      sourceName: userName,
      totalHours: 0,
      rows: []
    };

    existing.totalHours += hours;
    existing.rows.push({
      userName,
      documentName: documentName || "Unknown document",
      hoursBilled: Number(hours.toFixed(2))
    });
    groups.set(normalizedKey, existing);
  });

  return Array.from(groups.values())
    .map((entry) => ({
      normalizedKey: entry.normalizedKey,
      displayName: entry.displayName,
      sourceName: entry.sourceName,
      totalHours: Number(entry.totalHours.toFixed(2)),
      rows: entry.rows.sort((left, right) => left.documentName.localeCompare(right.documentName))
    }))
    .sort((left, right) => left.displayName.localeCompare(right.displayName));
}

function buildComparisonRows(kimaiUsers, invoiceGroups) {
  const kimaiMap = new Map(kimaiUsers.map((entry) => [entry.normalizedKey, entry]));
  const invoiceMap = new Map(invoiceGroups.map((entry) => [entry.normalizedKey, entry]));
  const keys = Array.from(new Set(kimaiUsers.map((entry) => entry.normalizedKey).concat(invoiceGroups.map((entry) => entry.normalizedKey)))).sort(
    (left, right) => {
      const leftLabel = getGroupLabel(left, kimaiMap, invoiceMap);
      const rightLabel = getGroupLabel(right, kimaiMap, invoiceMap);
      return leftLabel.localeCompare(rightLabel);
    }
  );

  const rows = [];
  let matchedGroupCount = 0;
  let totalDelta = 0;

  keys.forEach((key) => {
    const kimaiEntry = kimaiMap.get(key) || null;
    const invoiceEntry = invoiceMap.get(key) || null;
    const invoiceRows = invoiceEntry ? invoiceEntry.rows : [];
    const invoiceTotal = invoiceEntry ? invoiceEntry.totalHours : 0;
    const kimaiTotal = kimaiEntry ? kimaiEntry.totalHours : 0;
    const delta = Number((kimaiTotal - invoiceTotal).toFixed(2));
    const groupStatus = getGroupStatus(kimaiEntry, invoiceEntry, delta);
    const rowCount = Math.max(1, invoiceRows.length);

    if (kimaiEntry && invoiceEntry) {
      matchedGroupCount += 1;
    }

    totalDelta += delta;

    for (let index = 0; index < rowCount; index += 1) {
      const invoiceRow = invoiceRows[index] || null;
      rows.push({
        normalizedKey: key,
        groupStart: index === 0,
        groupStatus,
        kimaiName: index === 0 && kimaiEntry ? kimaiEntry.displayName : "",
        kimaiSourceName: kimaiEntry ? kimaiEntry.sourceName : "",
        kimaiHours: index === 0 && kimaiEntry ? formatHours(kimaiEntry.totalHours) : "",
        deltaLabel: index === 0 ? formatSignedHours(delta) : "",
        invoiceName: invoiceRow && index === 0 ? formatDisplayNameFromInvoice(invoiceRow.userName) : "",
        invoiceSourceName: invoiceRow ? invoiceRow.userName : "",
        invoiceDocument: invoiceRow ? invoiceRow.documentName : "",
        invoiceHours: invoiceRow ? formatHours(invoiceRow.hoursBilled) : ""
      });
    }
  });

  return {
    rows,
    groupCount: keys.length,
    matchedGroupCount,
    totalDelta: Number(totalDelta.toFixed(2))
  };
}

function renderMetrics(comparison) {
  const currentComparison = comparison || {
    matchedGroupCount: 0,
    totalDelta: 0
  };

  elements.kimaiUserCount.textContent = String(state.kimaiUsers.length);
  elements.invoiceUserCount.textContent = String(state.invoiceGroups.length);
  elements.matchedGroupCount.textContent = String(currentComparison.matchedGroupCount);
  elements.totalDelta.textContent = `${currentComparison.totalDelta.toFixed(1)} h`;
}

function renderMirrorTables(rows) {
  if (!rows.length) {
    resetTables("No comparable rows were found in the uploaded files.");
    return;
  }

  const kimaiTable = document.createElement("table");
  kimaiTable.append(
    buildKimaiHead(),
    buildKimaiBody(rows)
  );

  const invoiceTable = document.createElement("table");
  invoiceTable.append(
    buildInvoiceHead(),
    buildInvoiceBody(rows)
  );

  elements.kimaiTableState.className = "";
  elements.kimaiTableState.replaceChildren(kimaiTable);
  elements.invoiceTableState.className = "";
  elements.invoiceTableState.replaceChildren(invoiceTable);
}

function buildKimaiHead() {
  const thead = document.createElement("thead");
  const row = document.createElement("tr");
  ["User", "Kimai Hours", "Delta"].forEach((label) => {
    const th = document.createElement("th");
    th.textContent = label;
    row.append(th);
  });
  thead.append(row);
  return thead;
}

function buildInvoiceHead() {
  const thead = document.createElement("thead");
  const row = document.createElement("tr");
  ["User", "Document", "Invoice Hours"].forEach((label) => {
    const th = document.createElement("th");
    th.textContent = label;
    row.append(th);
  });
  thead.append(row);
  return thead;
}

function buildKimaiBody(rows) {
  const tbody = document.createElement("tbody");
  rows.forEach((rowData) => {
    const row = document.createElement("tr");
    applyRowClasses(row, rowData);

    const nameCell = document.createElement("td");
    nameCell.textContent = rowData.kimaiName;
    nameCell.title = rowData.kimaiSourceName || rowData.kimaiName;

    const hoursCell = document.createElement("td");
    hoursCell.textContent = rowData.kimaiHours;

    const deltaCell = document.createElement("td");
    deltaCell.textContent = rowData.deltaLabel;
    deltaCell.className = `delta-cell ${rowData.groupStatus}`;

    row.append(nameCell, hoursCell, deltaCell);
    tbody.append(row);
  });
  return tbody;
}

function buildInvoiceBody(rows) {
  const tbody = document.createElement("tbody");
  rows.forEach((rowData) => {
    const row = document.createElement("tr");
    applyRowClasses(row, rowData);

    const nameCell = document.createElement("td");
    nameCell.textContent = rowData.invoiceName;
    nameCell.title = rowData.invoiceSourceName || rowData.invoiceName;

    const documentCell = document.createElement("td");
    documentCell.textContent = rowData.invoiceDocument;
    documentCell.title = rowData.invoiceDocument;

    const hoursCell = document.createElement("td");
    hoursCell.textContent = rowData.invoiceHours;

    row.append(nameCell, documentCell, hoursCell);
    tbody.append(row);
  });
  return tbody;
}

function applyRowClasses(row, rowData) {
  if (rowData.groupStart) {
    row.classList.add("row-group-start");
  }
  row.classList.add(`row-status-${rowData.groupStatus}`);
}

function resetTables(message) {
  elements.kimaiTableState.className = "empty";
  elements.kimaiTableState.textContent = message;
  elements.invoiceTableState.className = "empty";
  elements.invoiceTableState.textContent = message;
  elements.comparisonBadge.textContent = "No comparison generated yet";
}

function setStatus(message) {
  elements.status.textContent = message;
}

function getGroupLabel(key, kimaiMap, invoiceMap) {
  const kimaiEntry = kimaiMap.get(key);
  const invoiceEntry = invoiceMap.get(key);
  if (kimaiEntry) {
    return kimaiEntry.displayName;
  }
  if (invoiceEntry) {
    return invoiceEntry.displayName;
  }
  return key;
}

function getGroupStatus(kimaiEntry, invoiceEntry, delta) {
  if (!kimaiEntry || !invoiceEntry) {
    return "missing";
  }
  return Math.abs(delta) < 0.01 ? "match" : "warning";
}

function normalizeHeader(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_|_$/g, "");
}

function normalizePersonName(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[._-]+/g, " ")
    .replace(/\s+/g, " ")
    .replace(/[^a-z0-9 ]+/g, "")
    .trim();
}

function formatDisplayNameFromKimai(value) {
  return toTitleCase(
    String(value || "")
      .replace(/[._-]+/g, " ")
      .trim()
  );
}

function formatDisplayNameFromInvoice(value) {
  return toTitleCase(String(value || "").trim());
}

function toTitleCase(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/\b\w/g, (character) => character.toUpperCase());
}

function getNumber(value) {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }

  const parsed = Number(
    String(value === undefined || value === null ? "" : value).replace(",", ".").replace(/[^\d.-]/g, "")
  );

  return Number.isFinite(parsed) ? parsed : 0;
}

function formatHours(value) {
  return `${value.toFixed(1)} h`;
}

function formatSignedHours(value) {
  const sign = value > 0 ? "+" : "";
  return `${sign}${value.toFixed(1)} h`;
}

function parseCsv(text) {
  const rows = [];
  let row = [];
  let value = "";
  let inQuotes = false;

  for (let index = 0; index < text.length; index += 1) {
    const character = text[index];
    const nextCharacter = text[index + 1];

    if (character === '"') {
      if (inQuotes && nextCharacter === '"') {
        value += '"';
        index += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (character === "," && !inQuotes) {
      row.push(value);
      value = "";
      continue;
    }

    if ((character === "\n" || character === "\r") && !inQuotes) {
      if (character === "\r" && nextCharacter === "\n") {
        index += 1;
      }
      row.push(value);
      rows.push(row);
      row = [];
      value = "";
      continue;
    }

    value += character;
  }

  if (value.length || row.length) {
    row.push(value);
    rows.push(row);
  }

  return rows.filter((currentRow) => currentRow.some((cell) => String(cell || "").trim() !== ""));
}
