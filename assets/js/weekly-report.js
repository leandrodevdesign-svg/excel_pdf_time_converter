const state = {
  allRows: [],
  fileNames: [],
  selectedUsers: new Set(),
  weekStart: "",
  weekEnd: ""
};

const elements = {
  pageLoader: document.getElementById("pageLoader"),
  excelInput: document.getElementById("excelInput"),
  pdfButton: document.getElementById("pdfBtn"),
  tablePdfButton: document.getElementById("tablePdfBtn"),
  downloadCsvButton: document.getElementById("downloadCsvBtn"),
  copyMissingUsersButton: document.getElementById("copyMissingUsersBtn"),
  status: document.getElementById("status"),
  periodBadge: document.getElementById("periodBadge"),
  weekStartInput: document.getElementById("weekStartInput"),
  weekEndInput: document.getElementById("weekEndInput"),
  weekStatus: document.getElementById("weekStatus"),
  userFilterList: document.getElementById("userFilterList"),
  selectAllUsersButton: document.getElementById("selectAllUsersBtn"),
  clearUsersButton: document.getElementById("clearUsersBtn"),
  kpiTotalHours: document.getElementById("kpiTotalHours"),
  kpiTotalHoursSub: document.getElementById("kpiTotalHoursSub"),
  kpiUsers: document.getElementById("kpiUsers"),
  kpiUsersSub: document.getElementById("kpiUsersSub"),
  kpiBestDay: document.getElementById("kpiBestDay"),
  kpiBestDaySub: document.getElementById("kpiBestDaySub"),
  kpiAvgUser: document.getElementById("kpiAvgUser"),
  kpiAvgUserSub: document.getElementById("kpiAvgUserSub"),
  tableContainer: document.getElementById("tableContainer"),
  weeklyTableSection: document.getElementById("weeklyTableSection"),
  weeklyReport: document.getElementById("weeklyReport")
};

elements.excelInput.addEventListener("change", handleFileUpload);
elements.pdfButton.addEventListener("click", exportPdf);
elements.tablePdfButton.addEventListener("click", exportTablePdf);
elements.downloadCsvButton.addEventListener("click", downloadWeeklyCsv);
elements.copyMissingUsersButton.addEventListener("click", copyUsersMissingHours);
elements.selectAllUsersButton.addEventListener("click", selectAllUsers);
elements.clearUsersButton.addEventListener("click", clearSelectedUsers);
elements.weekStartInput.addEventListener("change", () => handleWeekInputChange("start"));
elements.weekEndInput.addEventListener("change", () => handleWeekInputChange("end"));
window.addEventListener("load", hideInitialLoader, { once: true });

function hideInitialLoader() {
  if ("scrollRestoration" in window.history) {
    window.history.scrollRestoration = "manual";
  }

  window.scrollTo({ top: 0, left: 0, behavior: "auto" });

  window.setTimeout(() => {
    elements.pageLoader.classList.add("is-hidden");
    document.body.classList.remove("is-loading");
    window.scrollTo({ top: 0, left: 0, behavior: "auto" });
  }, 550);
}

function setStatus(message) {
  elements.status.textContent = message;
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

function excelDateToJsDate(value) {
  if (value instanceof Date) {
    return value;
  }

  if (typeof value === "number" && Number.isFinite(value)) {
    const utcDays = Math.floor(value - 25569);
    const utcValue = utcDays * 86400;
    const dateInfo = new Date(utcValue * 1000);
    const fractionalDay = value - Math.floor(value) + 0.0000001;

    let totalSeconds = Math.floor(86400 * fractionalDay);
    const seconds = totalSeconds % 60;
    totalSeconds -= seconds;

    const hours = Math.floor(totalSeconds / 3600);
    const minutes = Math.floor(totalSeconds / 60) % 60;

    return new Date(
      dateInfo.getFullYear(),
      dateInfo.getMonth(),
      dateInfo.getDate(),
      hours,
      minutes,
      seconds
    );
  }

  if (typeof value === "string" && value.trim()) {
    const parsed = new Date(value);
    if (!Number.isNaN(parsed.getTime())) {
      return parsed;
    }
  }

  return null;
}

function getNumber(value) {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }

  if (typeof value === "string") {
    const cleaned = value.replace(",", ".").replace(/[^\d.-]/g, "");
    const number = Number(cleaned);
    return Number.isFinite(number) ? number : 0;
  }

  return 0;
}

function getDurationHours(value, fallbackHours = 0) {
  if (value instanceof Date) {
    return value.getHours() + value.getMinutes() / 60 + value.getSeconds() / 3600;
  }

  if (typeof value === "number" && Number.isFinite(value)) {
    if (value > 0 && value < 1 && fallbackHours > 0) {
      const excelStyleHours = value * 24;
      return Math.abs(excelStyleHours - fallbackHours) < Math.abs(value - fallbackHours)
        ? excelStyleHours
        : value;
    }

    return value;
  }

  if (typeof value === "string") {
    const trimmed = value.trim();
    if (!trimmed) {
      return 0;
    }

    const durationMatch = trimmed.match(/^(\d{1,3}):(\d{2})(?::(\d{2}))?$/);
    if (durationMatch) {
      const hours = Number(durationMatch[1]);
      const minutes = Number(durationMatch[2]);
      const seconds = Number(durationMatch[3] || 0);
      return hours + minutes / 60 + seconds / 3600;
    }

    return getNumber(trimmed);
  }

  return 0;
}

function safeValue(value, fallback = "No data") {
  return value === undefined || value === null || String(value).trim() === ""
    ? fallback
    : String(value);
}

function calculateHoursFromDates(start, end) {
  if (!start || !end) {
    return 0;
  }

  const differenceInHours = (end - start) / 36e5;
  return differenceInHours > 0 ? differenceInHours : 0;
}

function getDateKey(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function formatHours(value) {
  return `${value.toFixed(1)} h`;
}

function formatDuration(value) {
  const totalMinutes = Math.max(0, Math.round((Number(value) || 0) * 60));
  const hours = Math.floor(totalMinutes / 60);
  const minutes = totalMinutes % 60;
  return `${hours}:${String(minutes).padStart(2, "0")}`;
}

function formatDate(date) {
  return new Intl.DateTimeFormat("en-US", {
    day: "2-digit",
    month: "short",
    year: "numeric"
  }).format(date);
}

function formatDayHeader(date) {
  return new Intl.DateTimeFormat("en-US", {
    weekday: "short",
    month: "2-digit",
    day: "2-digit"
  }).format(date);
}

function parseDateKey(dateKey) {
  if (!dateKey) {
    return null;
  }

  const [year, month, day] = dateKey.split("-").map(Number);
  if (!year || !month || !day) {
    return null;
  }

  return new Date(year, month - 1, day);
}

function addDays(date, amount) {
  const next = new Date(date);
  next.setDate(next.getDate() + amount);
  return next;
}

function getDateRange(startKey, endKey) {
  const start = parseDateKey(startKey);
  const end = parseDateKey(endKey);

  if (!start || !end || start > end) {
    return [];
  }

  const dates = [];
  let current = new Date(start);
  while (current <= end) {
    dates.push(new Date(current));
    current = addDays(current, 1);
  }

  return dates;
}

async function handleFileUpload(event) {
  const files = Array.from(event.target.files || []);
  if (!files.length) {
    return;
  }

  setStatus(
    files.length === 1
      ? `Processing ${files[0].name}...`
      : `Processing ${files.length} files...`
  );

  try {
    const fileRowGroups = await Promise.all(files.map(parseWorkbookFile));
    const rows = fileRowGroups
      .flat()
      .sort((left, right) => left.dateObj - right.dateObj || left.inTimestamp - right.inTimestamp);

    if (!rows.length) {
      throw new Error("No valid records were found in the uploaded file set.");
    }

    state.allRows = rows;
    state.fileNames = files.map((file) => file.name);
    syncSelectedUsers(true);
    initializeWeekRange();

    setStatus(
      files.length === 1
        ? `Loaded ${files[0].name} successfully. Processed ${rows.length} records.`
        : `Loaded ${files.length} files successfully. Processed ${rows.length} records.`
    );

    renderCurrentView();
  } catch (error) {
    console.error(error);
    setStatus(`The file could not be processed. ${error.message || "Check the format and try again."}`);
  }
}

async function parseWorkbookFile(file) {
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
    throw new Error(`${file.name} is empty.`);
  }

  const headers = matrix[0].map(normalizeHeader);
  return matrix
    .slice(1)
    .filter((row) => row.some((cell) => String(cell === null || cell === undefined ? "" : cell).trim() !== ""))
    .map((row) =>
      row.reduce((record, cell, index) => {
        record[headers[index] || `col_${index}`] = cell;
        return record;
      }, {})
    )
    .map((row) => parseRow(row, file.name))
    .filter(Boolean);
}

function parseRow(row, sourceFile) {
  const dateObj = excelDateToJsDate(row.date || row.fecha || row.day);
  const inObj = excelDateToJsDate(row.in || row.entrada || row.start || row.start_time);
  const outObj = excelDateToJsDate(row.out || row.salida || row.end || row.end_time);
  const calculatedHours = calculateHoursFromDates(inObj, outObj);
  const loggedHours = getDurationHours(
    row.time || row.horas || row.hours || row.total || row.total_hours,
    calculatedHours
  );
  const effectiveHours = loggedHours > 0 ? loggedHours : calculatedHours;

  if (!dateObj || !Number.isFinite(effectiveHours) || effectiveHours <= 0) {
    return null;
  }

  return {
    dateObj,
    dateKey: getDateKey(dateObj),
    inTimestamp: inObj instanceof Date ? inObj.getTime() : 0,
    hours: effectiveHours,
    username: safeValue(row.username || row.usuario || row.user, "-"),
    sourceFile
  };
}

function getAvailableUsers() {
  return Array.from(new Set(state.allRows.map((row) => row.username))).sort((left, right) =>
    left.localeCompare(right)
  );
}

function syncSelectedUsers(forceSelectAll = false) {
  const availableUsers = getAvailableUsers();

  if (forceSelectAll || !state.selectedUsers.size) {
    state.selectedUsers = new Set(availableUsers);
    return;
  }

  state.selectedUsers = new Set(
    availableUsers.filter((username) => state.selectedUsers.has(username))
  );
}

function getSelectedRows() {
  if (!state.allRows.length || !state.selectedUsers.size) {
    return [];
  }

  return state.allRows.filter((row) => state.selectedUsers.has(row.username));
}

function getWeeklyRows() {
  const rows = getSelectedRows();
  if (!rows.length || !state.weekStart || !state.weekEnd) {
    return [];
  }

  return rows.filter((row) => row.dateKey >= state.weekStart && row.dateKey <= state.weekEnd);
}

function getDateBounds() {
  if (!state.allRows.length) {
    return null;
  }

  const sortedDates = state.allRows
    .map((row) => row.dateKey)
    .sort((left, right) => left.localeCompare(right));

  return {
    min: sortedDates[0],
    max: sortedDates[sortedDates.length - 1]
  };
}

function initializeWeekRange() {
  const bounds = getDateBounds();
  if (!bounds) {
    state.weekStart = "";
    state.weekEnd = "";
    return;
  }

  const end = parseDateKey(bounds.max);
  const defaultStart = addDays(end, -6);
  const min = parseDateKey(bounds.min);
  const start = defaultStart < min ? min : defaultStart;

  state.weekStart = getDateKey(start);
  state.weekEnd = bounds.max;
}

function syncWeekInputs() {
  const bounds = getDateBounds();
  if (!bounds) {
    elements.weekStartInput.value = "";
    elements.weekEndInput.value = "";
    elements.weekStartInput.disabled = true;
    elements.weekEndInput.disabled = true;
    return;
  }

  elements.weekStartInput.disabled = false;
  elements.weekEndInput.disabled = false;
  elements.weekStartInput.min = bounds.min;
  elements.weekStartInput.max = bounds.max;
  elements.weekEndInput.min = bounds.min;
  elements.weekEndInput.max = bounds.max;
  elements.weekStartInput.value = state.weekStart;
  elements.weekEndInput.value = state.weekEnd;
}

function handleWeekInputChange(changedField) {
  if (!state.allRows.length) {
    return;
  }

  const bounds = getDateBounds();
  const min = parseDateKey(bounds.min);
  const max = parseDateKey(bounds.max);

  let start = parseDateKey(state.weekStart);
  let end = parseDateKey(state.weekEnd);

  if (changedField === "start" && elements.weekStartInput.value) {
    start = parseDateKey(elements.weekStartInput.value);
    end = addDays(start, 6);
    if (end > max) {
      end = max;
    }
  }

  if (changedField === "end" && elements.weekEndInput.value) {
    end = parseDateKey(elements.weekEndInput.value);
    start = addDays(end, -6);
    if (start < min) {
      start = min;
    }
  }

  if (start > end) {
    start = new Date(end);
  }

  state.weekStart = getDateKey(start);
  state.weekEnd = getDateKey(end);
  renderCurrentView();
}

function renderCurrentView() {
  renderUserFilters();
  syncWeekInputs();

  const weeklyRows = getWeeklyRows();
  const weekDates = getDateRange(state.weekStart, state.weekEnd);
  renderSummary(weeklyRows, weekDates);
  renderTable(getWeeklyTotalsByUser(weekDates, weeklyRows));
}

function renderUserFilters() {
  const availableUsers = getAvailableUsers();
  elements.userFilterList.replaceChildren();

  if (!availableUsers.length) {
    return;
  }

  availableUsers.forEach((username) => {
    const label = document.createElement("label");
    label.className = "user-chip";

    const input = document.createElement("input");
    input.type = "checkbox";
    input.value = username;
    input.checked = state.selectedUsers.has(username);
    input.addEventListener("change", () => toggleUserSelection(username, input.checked));

    const text = document.createElement("span");
    text.textContent = username;

    label.append(input, text);
    elements.userFilterList.appendChild(label);
  });
}

function toggleUserSelection(username, isSelected) {
  if (isSelected) {
    state.selectedUsers.add(username);
  } else {
    state.selectedUsers.delete(username);
  }

  renderCurrentView();
  setStatus(
    state.selectedUsers.size
      ? `Showing ${state.selectedUsers.size} selected user(s).`
      : "No users selected. Choose at least one user to display results."
  );
}

function selectAllUsers() {
  state.selectedUsers = new Set(getAvailableUsers());
  renderCurrentView();
  setStatus(`Showing all ${state.selectedUsers.size} user(s).`);
}

function clearSelectedUsers() {
  state.selectedUsers = new Set();
  renderCurrentView();
  setStatus("No users selected. Choose at least one user to display results.");
}

function buildPeriodBadge(weekDates) {
  if (!weekDates.length) {
    return "No data loaded yet";
  }

  const first = weekDates[0];
  const last = weekDates[weekDates.length - 1];
  const selectedUsers = state.selectedUsers.size;
  const totalUsers = getAvailableUsers().length;
  return `${formatDate(first)} -> ${formatDate(last)} · ${state.fileNames.length} file(s) · ${selectedUsers}/${totalUsers} user(s)`;
}

function getWeeklyTotalsByUser(weekDates, weeklyRows) {
  const selectedUserList = Array.from(state.selectedUsers).sort((left, right) => left.localeCompare(right));
  const totalsByUser = new Map();

  selectedUserList.forEach((username) => {
    const dateTotals = {};
    weekDates.forEach((date) => {
      dateTotals[getDateKey(date)] = 0;
    });

    totalsByUser.set(username, {
      username,
      dateTotals,
      totalHours: 0
    });
  });

  weeklyRows.forEach((row) => {
    const userRecord = totalsByUser.get(row.username);
    if (!userRecord) {
      return;
    }

    userRecord.dateTotals[row.dateKey] = (userRecord.dateTotals[row.dateKey] || 0) + row.hours;
    userRecord.totalHours += row.hours;
  });

  return {
    weekDates,
    selectedUserList,
    totalsByUser
  };
}

function renderSummary(weeklyRows, weekDates) {
  if (!state.allRows.length) {
    resetSummary();
    return;
  }

  elements.periodBadge.textContent = buildPeriodBadge(weekDates);

  if (!state.selectedUsers.size) {
    elements.weekStatus.textContent = "No users selected. Choose at least one user to build the weekly table.";
    resetSummary();
    return;
  }

  if (!weekDates.length) {
    elements.weekStatus.textContent = "Choose a valid weekly range to render the report.";
    resetSummary();
    return;
  }

  const dayTotals = new Map();
  weekDates.forEach((date) => {
    dayTotals.set(getDateKey(date), 0);
  });

  weeklyRows.forEach((row) => {
    dayTotals.set(row.dateKey, (dayTotals.get(row.dateKey) || 0) + row.hours);
  });

  const totalHours = weeklyRows.reduce((sum, row) => sum + row.hours, 0);
  const activeUsers = new Set(weeklyRows.map((row) => row.username)).size;
  const busiestDayEntry = Array.from(dayTotals.entries()).reduce(
    (best, [dateKey, hours]) => (hours > best.hours ? { dateKey, hours } : best),
    { dateKey: weekDates[0] ? getDateKey(weekDates[0]) : "", hours: 0 }
  );

  elements.weekStatus.textContent = `${weekDates.length} day(s) selected. The weekly table updates automatically when you change dates or users.`;
  elements.kpiTotalHours.textContent = formatDuration(totalHours);
  elements.kpiTotalHoursSub.textContent = `${weeklyRows.length} record(s) matched the current filters`;
  elements.kpiUsers.textContent = String(activeUsers);
  elements.kpiUsersSub.textContent = `${state.selectedUsers.size} user(s) currently selected`;
  elements.kpiBestDay.textContent = formatDuration(busiestDayEntry.hours);
  elements.kpiBestDaySub.textContent = totalHours > 0 && busiestDayEntry.dateKey
    ? `${formatDate(parseDateKey(busiestDayEntry.dateKey))} carried the highest load`
    : "No weekly activity yet";
  elements.kpiAvgUser.textContent = formatDuration(activeUsers ? totalHours / activeUsers : 0);
  elements.kpiAvgUserSub.textContent = activeUsers
    ? "Average calculated only on users with weekly activity"
    : "No user logged hours in the selected week";
}

function resetSummary() {
  elements.kpiTotalHours.textContent = "0:00";
  elements.kpiTotalHoursSub.textContent = "Select a week to summarize logged hours";
  elements.kpiUsers.textContent = "0";
  elements.kpiUsersSub.textContent = "Weekly active users will appear here";
  elements.kpiBestDay.textContent = "-";
  elements.kpiBestDaySub.textContent = "The strongest day of the selected week";
  elements.kpiAvgUser.textContent = "0:00";
  elements.kpiAvgUserSub.textContent = "Calculated on users with weekly activity";
}

function renderTable(summary) {
  const { weekDates, selectedUserList, totalsByUser } = summary;

  if (!state.allRows.length) {
    elements.tableContainer.className = "empty";
    elements.tableContainer.textContent = "There are no weekly records to display yet.";
    return;
  }

  if (!state.selectedUsers.size) {
    elements.tableContainer.className = "empty";
    elements.tableContainer.textContent = "Select at least one user to display the weekly report.";
    return;
  }

  if (!weekDates.length) {
    elements.tableContainer.className = "empty";
    elements.tableContainer.textContent = "Choose a valid weekly range to display the report.";
    return;
  }

  const table = document.createElement("table");
  const thead = document.createElement("thead");
  const headRow = document.createElement("tr");

  const userHeader = document.createElement("th");
  userHeader.textContent = "User";
  headRow.appendChild(userHeader);

  weekDates.forEach((date) => {
    const th = document.createElement("th");
    th.textContent = formatDayHeader(date);
    headRow.appendChild(th);
  });

  thead.appendChild(headRow);

  const tbody = document.createElement("tbody");
  selectedUserList.forEach((username) => {
    const tr = document.createElement("tr");
    const nameCell = document.createElement("td");
    nameCell.textContent = username;
    tr.appendChild(nameCell);

    weekDates.forEach((date) => {
      const td = document.createElement("td");
      const hours = totalsByUser.get(username).dateTotals[getDateKey(date)] || 0;
      td.textContent = formatDuration(hours);
      tr.appendChild(td);
    });

    tbody.appendChild(tr);
  });

  table.append(thead, tbody);
  elements.tableContainer.className = "table-wrap";
  elements.tableContainer.replaceChildren(table);
}

function buildPdfFilename() {
  const start = state.weekStart || "week-start";
  const end = state.weekEnd || "week-end";
  return `weekly-report-${start}-to-${end}.pdf`;
}

function buildCsvFilename() {
  const start = state.weekStart || "week-start";
  const end = state.weekEnd || "week-end";
  return `kimai-weekly-hours-${start}-to-${end}.csv`;
}

function getActiveWeeklySummary() {
  const weekDates = getDateRange(state.weekStart, state.weekEnd);
  const weeklyRows = getWeeklyRows();
  return getWeeklyTotalsByUser(weekDates, weeklyRows);
}

function escapeCsvValue(value) {
  const text = String(value === undefined || value === null ? "" : value);
  if (/[",\n]/.test(text)) {
    return `"${text.replace(/"/g, "\"\"")}"`;
  }

  return text;
}

function downloadTextFile(filename, content, mimeType) {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 0);
}

async function copyTextToClipboard(text) {
  if (navigator.clipboard && window.isSecureContext) {
    await navigator.clipboard.writeText(text);
    return;
  }

  const textArea = document.createElement("textarea");
  textArea.value = text;
  textArea.setAttribute("readonly", "");
  textArea.style.position = "fixed";
  textArea.style.opacity = "0";
  document.body.appendChild(textArea);
  textArea.select();
  document.execCommand("copy");
  textArea.remove();
}

async function exportPdf() {
  if (!state.allRows.length) {
    setStatus("Upload files before exporting the weekly report to PDF.");
    return;
  }

  if (!state.selectedUsers.size) {
    setStatus("Select at least one user before exporting the weekly report to PDF.");
    return;
  }

  setStatus("Generating a print-friendly PDF...");

  document.documentElement.classList.add("pdf-mode");
  document.body.classList.add("pdf-mode");
  elements.weeklyReport.classList.add("pdf-mode");

  const options = {
    margin: [8, 8, 8, 8],
    filename: buildPdfFilename(),
    image: { type: "jpeg", quality: 0.98 },
    html2canvas: {
      scale: 2,
      useCORS: true,
      backgroundColor: "#ffffff",
      scrollX: 0,
      scrollY: 0
    },
    jsPDF: {
      unit: "mm",
      format: "a4",
      orientation: "landscape"
    },
    pagebreak: {
      mode: ["avoid-all", "css", "legacy"],
      avoid: [".card", ".kpi", ".panel", "tr"]
    }
  };

  try {
    await new Promise((resolve) => window.setTimeout(resolve, 180));
    await html2pdf().set(options).from(elements.weeklyReport).save();
    setStatus("The PDF was exported successfully.");
  } catch (error) {
    console.error(error);
    setStatus("The PDF export failed.");
  } finally {
    document.documentElement.classList.remove("pdf-mode");
    document.body.classList.remove("pdf-mode");
    elements.weeklyReport.classList.remove("pdf-mode");
  }
}

async function exportTablePdf() {
  if (!state.allRows.length) {
    setStatus("Upload files before exporting the weekly table to PDF.");
    return;
  }

  if (!state.selectedUsers.size) {
    setStatus("Select at least one user before exporting the weekly table to PDF.");
    return;
  }

  if (!getDateRange(state.weekStart, state.weekEnd).length) {
    setStatus("Choose a valid weekly range before exporting the table to PDF.");
    return;
  }

  setStatus("Generating the weekly table PDF...");

  document.documentElement.classList.add("pdf-mode");
  document.body.classList.add("pdf-mode");
  elements.weeklyTableSection.classList.add("pdf-mode");

  const options = {
    margin: [8, 8, 8, 8],
    filename: buildPdfFilename(),
    image: { type: "jpeg", quality: 0.98 },
    html2canvas: {
      scale: 2,
      useCORS: true,
      backgroundColor: "#ffffff",
      scrollX: 0,
      scrollY: 0
    },
    jsPDF: {
      unit: "mm",
      format: "a4",
      orientation: "landscape"
    },
    pagebreak: {
      mode: ["avoid-all", "css", "legacy"],
      avoid: [".card", ".panel", "tr"]
    }
  };

  try {
    await new Promise((resolve) => window.setTimeout(resolve, 180));
    await html2pdf().set(options).from(elements.weeklyTableSection).save();
    setStatus("The weekly table PDF was exported successfully.");
  } catch (error) {
    console.error(error);
    setStatus("The weekly table PDF export failed.");
  } finally {
    document.documentElement.classList.remove("pdf-mode");
    document.body.classList.remove("pdf-mode");
    elements.weeklyTableSection.classList.remove("pdf-mode");
  }
}

function downloadWeeklyCsv() {
  if (!state.allRows.length) {
    setStatus("Upload files before downloading the weekly CSV.");
    return;
  }

  if (!state.selectedUsers.size) {
    setStatus("Select at least one user before downloading the weekly CSV.");
    return;
  }

  const summary = getActiveWeeklySummary();
  if (!summary.weekDates.length) {
    setStatus("Choose a valid weekly range before downloading the weekly CSV.");
    return;
  }

  const headers = ["User", ...summary.weekDates.map((date) => formatDayHeader(date)), "Total Week"];
  const lines = [headers.map(escapeCsvValue).join(",")];

  summary.selectedUserList.forEach((username) => {
    const userSummary = summary.totalsByUser.get(username);
    const row = [
      username,
      ...summary.weekDates.map((date) => formatDuration(userSummary.dateTotals[getDateKey(date)] || 0)),
      formatDuration(userSummary.totalHours)
    ];
    lines.push(row.map(escapeCsvValue).join(","));
  });

  downloadTextFile(buildCsvFilename(), lines.join("\n"), "text/csv;charset=utf-8");
  setStatus("The weekly CSV was downloaded successfully.");
}

async function copyUsersMissingHours() {
  if (!state.allRows.length) {
    setStatus("Upload files before copying users with missing hours.");
    return;
  }

  if (!state.selectedUsers.size) {
    setStatus("Select at least one user before copying users with missing hours.");
    return;
  }

  const summary = getActiveWeeklySummary();
  if (!summary.weekDates.length) {
    setStatus("Choose a valid weekly range before copying users with missing hours.");
    return;
  }

  const missingUsers = summary.selectedUserList.filter((username) => {
    const userSummary = summary.totalsByUser.get(username);
    return userSummary.totalHours < 24;
  });

  if (!missingUsers.length) {
    setStatus("No selected users are below 24:00 for the current week.");
    return;
  }

  try {
    await copyTextToClipboard(missingUsers.join(", "));
    setStatus(`Copied ${missingUsers.length} user(s) missing weekly hours to the clipboard.`);
  } catch (error) {
    console.error(error);
    setStatus("The user list could not be copied to the clipboard.");
  }
}
