const COPY = {
  single: {
    eyebrow: "Time Analytics",
    description:
      "Upload a single Excel file with columns such as Date, In, Out, Time, Customer, Project, and Activity. The application restructures the raw data into an executive dashboard with metrics, workload distribution, quick insights, and an audit-ready detail table.",
    uploadLabel: "Upload Excel File",
    waitingStatus: "Waiting for a file. Use the same structure exported by your time-tracking system."
  },
  batch: {
    eyebrow: "Batch Analytics",
    description:
      "Upload one Excel file with multiple users or several Excel files at once. Review the combined result, filter the output by one or many users, and export a PDF containing only the selected users.",
    uploadLabel: "Upload One or More Excel Files",
    waitingStatus: "Waiting for one or more files. Batch mode supports multiple users and multi-file uploads."
  }
};

const CLIENT_OVERVIEW_STORAGE_KEY = "excel_pdf_time_converter_client_overview";

// This application keeps all state in memory because the dashboard is meant
// to be a lightweight client-side tool with no backend dependencies.
const state = {
  mode: "single",
  allRows: [],
  fileNames: [],
  selectedUsers: new Set(),
  activeDailyMonthKey: "",
  currentTablePage: 1,
  activeTypingRun: 0
};

const elements = {
  pageLoader: document.getElementById("pageLoader"),
  modeLinks: Array.from(document.querySelectorAll("[data-mode-link]")),
  heroEyebrow: document.getElementById("heroEyebrow"),
  heroDescription: document.getElementById("heroDescription"),
  uploadLabel: document.getElementById("uploadLabel"),
  excelInput: document.getElementById("excelInput"),
  pdfButton: document.getElementById("pdfBtn"),
  status: document.getElementById("status"),
  periodBadge: document.getElementById("periodBadge"),
  batchFilters: document.getElementById("batchFilters"),
  userFilterList: document.getElementById("userFilterList"),
  selectAllUsersButton: document.getElementById("selectAllUsersBtn"),
  clearUsersButton: document.getElementById("clearUsersBtn"),
  dailyPrevMonthButton: document.getElementById("dailyPrevMonthBtn"),
  dailyNextMonthButton: document.getElementById("dailyNextMonthBtn"),
  dailyMonthLabel: document.getElementById("dailyMonthLabel"),
  kpiTotal: document.getElementById("kpiTotal"),
  kpiTotalSub: document.getElementById("kpiTotalSub"),
  kpiAvg: document.getElementById("kpiAvg"),
  kpiAvgSub: document.getElementById("kpiAvgSub"),
  kpiBest: document.getElementById("kpiBest"),
  kpiBestSub: document.getElementById("kpiBestSub"),
  kpiBillable: document.getElementById("kpiBillable"),
  kpiBillableSub: document.getElementById("kpiBillableSub"),
  dailyChart: document.getElementById("dailyChart"),
  projectChart: document.getElementById("projectChart"),
  activityChart: document.getElementById("activityChart"),
  quickStats: document.getElementById("quickStats"),
  tableContainer: document.getElementById("tableContainer"),
  tablePagination: document.getElementById("tablePagination"),
  dashboard: document.getElementById("dashboard")
};

elements.excelInput.addEventListener("change", handleFileUpload);
elements.pdfButton.addEventListener("click", exportPdf);
elements.selectAllUsersButton.addEventListener("click", selectAllUsers);
elements.clearUsersButton.addEventListener("click", clearSelectedUsers);
elements.dailyPrevMonthButton.addEventListener("click", () => shiftDailyMonth(-1));
elements.dailyNextMonthButton.addEventListener("click", () => shiftDailyMonth(1));
elements.modeLinks.forEach((link) => {
  link.addEventListener("click", (event) => {
    event.preventDefault();
    setMode(link.dataset.modeLink);
  });
});
window.addEventListener("load", hideInitialLoader, { once: true });
window.addEventListener("hashchange", handleHashModeChange);

setMode(getModeFromHash(), { updateHash: false });

function setStatus(message) {
  elements.status.textContent = message;
}

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

function setMode(mode, options) {
  const settings = options || {};
  state.mode = mode === "batch" ? "batch" : "single";
  state.currentTablePage = 1;
  const copy = COPY[state.mode];

  elements.modeLinks.forEach((link) => {
    const isActive = link.dataset.modeLink === state.mode;
    link.classList.toggle("active", isActive);
    if (isActive) {
      link.setAttribute("aria-current", "page");
    } else {
      link.removeAttribute("aria-current");
    }
  });

  elements.heroEyebrow.textContent = copy.eyebrow;
  elements.heroDescription.textContent = copy.description;
  elements.uploadLabel.textContent = copy.uploadLabel;
  elements.excelInput.toggleAttribute("multiple", state.mode === "batch");
  elements.excelInput.value = "";

  if (settings.updateHash !== false) {
    window.history.replaceState(null, "", `#${state.mode}`);
  }

  if (!state.allRows.length) {
    setStatus(copy.waitingStatus);
    elements.periodBadge.textContent = "No data loaded yet";
  }

  syncSelectedUsers();
  renderCurrentView({ animate: false });
}

function getModeFromHash() {
  return window.location.hash === "#batch" ? "batch" : "single";
}

function handleHashModeChange() {
  const modeFromHash = getModeFromHash();
  if (modeFromHash !== state.mode) {
    setMode(modeFromHash, { updateHash: false });
  }
}

// Spreadsheet exports often vary only slightly in their column names.
// This normalizer lets the parser support multiple header spellings.
function normalizeHeader(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_|_$/g, "");
}

// Excel stores dates as serial numbers. This helper converts those serials,
// as well as normal JavaScript Date objects and parseable strings.
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

function formatHours(value) {
  return `${value.toFixed(1)} h`;
}

function formatPercent(value) {
  return `${value.toFixed(1)}%`;
}

function formatDate(date) {
  return new Intl.DateTimeFormat("en-US", {
    day: "2-digit",
    month: "short",
    year: "numeric"
  }).format(date);
}

function formatDateShort(date) {
  return new Intl.DateTimeFormat("en-US", {
    day: "2-digit",
    month: "2-digit"
  }).format(date);
}

function formatDateForFilename(date) {
  return new Intl.DateTimeFormat("en-US", {
    month: "short",
    day: "numeric"
  }).format(date);
}

function formatMonthLabel(date) {
  return new Intl.DateTimeFormat("en-US", {
    month: "long",
    year: "numeric"
  }).format(date);
}

function formatTime(date) {
  if (!date) {
    return "-";
  }

  return new Intl.DateTimeFormat("en-US", {
    hour: "2-digit",
    minute: "2-digit",
    hour12: false
  }).format(date);
}

function compareByValueDesc(a, b) {
  return b.value - a.value;
}

function safeValue(value, fallback = "No data") {
  return value === undefined || value === null || String(value).trim() === ""
    ? fallback
    : String(value);
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
    state.currentTablePage = 1;
    syncSelectedUsers(true);

    setStatus(
      files.length === 1
        ? `Loaded ${files[0].name} successfully. Processed ${rows.length} records.`
        : `Loaded ${files.length} files successfully. Processed ${rows.length} records.`
    );

    renderCurrentView({ animate: true });
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

// The parser accepts a few alternate field names so the app remains useful
// even when spreadsheet exports differ between systems or locales.
function parseRow(row, sourceFile) {
  const dateObj = excelDateToJsDate(row.date || row.fecha || row.day);
  const inObj = excelDateToJsDate(row.in || row.entrada || row.start || row.start_time);
  const outObj = excelDateToJsDate(row.out || row.salida || row.end || row.end_time);
  const loggedHours = getNumber(row.time || row.horas || row.hours || row.total || row.total_hours);
  const billableRaw = getNumber(row.billable || row.facturable || row.billable_hours);
  const effectiveHours = loggedHours > 0 ? loggedHours : calculateHoursFromDates(inObj, outObj);

  if (!dateObj || !Number.isFinite(effectiveHours) || effectiveHours <= 0) {
    return null;
  }

  return {
    dateObj,
    dateKey: getDateKey(dateObj),
    inObj,
    outObj,
    inTimestamp: inObj instanceof Date ? inObj.getTime() : 0,
    hours: effectiveHours,
    billableHours: Math.max(0, Math.min(billableRaw, effectiveHours)),
    customer: safeValue(row.customer || row.cliente),
    project: safeValue(row.project || row.proyecto),
    activity: safeValue(row.activity || row.actividad),
    description: safeValue(row.description || row.descripcion, "-"),
    username: safeValue(row.username || row.usuario || row.user, "-"),
    sourceFile
  };
}

function getDateKey(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function getMonthKey(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
}

function calculateHoursFromDates(start, end) {
  if (!start || !end) {
    return 0;
  }

  const differenceInHours = (end - start) / 36e5;
  return differenceInHours > 0 ? differenceInHours : 0;
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

function getVisibleRows() {
  if (!state.allRows.length) {
    return [];
  }

  if (state.mode !== "batch") {
    return state.allRows;
  }

  if (!state.selectedUsers.size) {
    return [];
  }

  return state.allRows.filter((row) => state.selectedUsers.has(row.username));
}

function buildPdfFilename(rows) {
  const sortedDates = rows
    .map((row) => row.dateObj)
    .sort((left, right) => left - right);

  const startDate = sortedDates[0];
  const endDate = sortedDates[sortedDates.length - 1];
  const dateRange = `from-${formatDateForFilename(startDate)}-to-${formatDateForFilename(endDate)}`
    .toLowerCase()
    .replace(/[^a-z0-9\s-]/g, "")
    .trim()
    .replace(/\s+/g, "-");

  const users = Array.from(new Set(rows.map((row) => row.username).filter((value) => value && value !== "-")));
  const rawPrefix = users.length === 1 ? users[0] : users.length > 1 ? `${users.length}-users` : "user";
  const sanitizedPrefix = rawPrefix
    .toLowerCase()
    .replace(/[^a-z0-9\s-]/g, "")
    .trim()
    .replace(/\s+/g, "-") || "user";

  return `${sanitizedPrefix}-${dateRange}.pdf`;
}

function groupSum(rows, key) {
  const groups = new Map();

  rows.forEach((row) => {
    const groupKey = safeValue(row[key]);
    groups.set(groupKey, (groups.get(groupKey) || 0) + row.hours);
  });

  return Array.from(groups, ([label, value]) => ({ label, value })).sort(compareByValueDesc);
}

function renderCurrentView({ animate }) {
  renderUserFilters();

  const visibleRows = getVisibleRows();
  renderDashboard(visibleRows);
  persistClientOverviewSnapshot(visibleRows);

  if (animate && visibleRows.length) {
    animateTopSectionAfterUpload();
  }
}

function persistClientOverviewSnapshot(visibleRows) {
  if (state.mode !== "batch" || !state.allRows.length) {
    return;
  }

  const payload = {
    savedAt: new Date().toISOString(),
    fileNames: state.fileNames,
    selectedUsers: Array.from(state.selectedUsers),
    rows: visibleRows.map((row) => ({
      dateKey: row.dateKey,
      dateObj: row.dateObj instanceof Date ? row.dateObj.toISOString() : null,
      hours: row.hours,
      billableHours: row.billableHours,
      customer: row.customer,
      project: row.project,
      activity: row.activity,
      description: row.description,
      username: row.username,
      sourceFile: row.sourceFile
    }))
  };

  try {
    window.sessionStorage.setItem(CLIENT_OVERVIEW_STORAGE_KEY, JSON.stringify(payload));
  } catch (error) {
    console.error("Failed to persist client overview snapshot.", error);
  }
}

function renderUserFilters() {
  const availableUsers = getAvailableUsers();
  const shouldShow = state.mode === "batch" && availableUsers.length > 0;

  elements.batchFilters.classList.toggle("is-hidden", !shouldShow);
  elements.userFilterList.replaceChildren();

  if (!shouldShow) {
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

  renderCurrentView({ animate: false });
  state.currentTablePage = 1;
  setStatus(
    state.selectedUsers.size
      ? `Showing ${state.selectedUsers.size} selected user(s).`
      : "No users selected. Choose at least one user to display results."
  );
}

function selectAllUsers() {
  state.selectedUsers = new Set(getAvailableUsers());
  state.currentTablePage = 1;
  renderCurrentView({ animate: false });
  setStatus(`Showing all ${state.selectedUsers.size} user(s).`);
}

function clearSelectedUsers() {
  state.selectedUsers = new Set();
  state.currentTablePage = 1;
  renderCurrentView({ animate: false });
  setStatus("No users selected. Choose at least one user to display results.");
}

function renderDashboard(rows) {
  if (!rows.length) {
    const emptyReason = state.allRows.length
      ? "There are no records for the current user selection."
      : "There are no records to display yet.";
    resetDashboard(emptyReason);
    return;
  }

  const totalHours = rows.reduce((sum, row) => sum + row.hours, 0);
  const totalBillable = rows.reduce((sum, row) => sum + row.billableHours, 0);
  const byDate = new Map();

  rows.forEach((row) => {
    if (!byDate.has(row.dateKey)) {
      byDate.set(row.dateKey, { dateObj: row.dateObj, hours: 0, entries: 0 });
    }

    const dateRecord = byDate.get(row.dateKey);
    dateRecord.hours += row.hours;
    dateRecord.entries += 1;
  });

  const daily = Array.from(byDate.values()).sort((left, right) => left.dateObj - right.dateObj);
  const averagePerDay = totalHours / daily.length;
  const bestDay = daily.reduce((max, item) => (item.hours > max.hours ? item : max), daily[0]);
  const billablePercentage = totalHours ? (totalBillable / totalHours) * 100 : 0;
  const byProject = groupSum(rows, "project").slice(0, 8);
  const byActivity = groupSum(rows, "activity").slice(0, 8);
  const byCustomer = groupSum(rows, "customer");
  const uniqueProjects = new Set(rows.map((row) => row.project)).size;
  const uniqueCustomers = new Set(rows.map((row) => row.customer)).size;
  const averageEntry = totalHours / rows.length;
  const longestEntry = rows.reduce((max, item) => (item.hours > max.hours ? item : max), rows[0]);
  const firstDate = daily.length ? daily[0].dateObj : null;
  const lastDate = daily.length ? daily[daily.length - 1].dateObj : null;

  elements.periodBadge.textContent = buildPeriodBadge(firstDate, lastDate, rows);

  elements.kpiTotal.textContent = formatHours(totalHours);
  elements.kpiTotalSub.textContent = `${rows.length} processed records`;
  elements.kpiAvg.textContent = formatHours(averagePerDay);
  elements.kpiAvgSub.textContent = `${daily.length} days with logged activity`;
  elements.kpiBest.textContent = formatHours(bestDay.hours);
  elements.kpiBestSub.textContent = `${formatDate(bestDay.dateObj)} was the busiest day`;
  elements.kpiBillable.textContent = formatPercent(billablePercentage);
  elements.kpiBillableSub.textContent = `${formatHours(totalBillable)} out of ${formatHours(totalHours)} were billable`;

  renderDailyDistribution(daily);
  renderBars(elements.projectChart, byProject, 8);
  renderBars(elements.activityChart, byActivity, 8);

  renderQuickStats([
    {
      title: state.mode === "batch" ? "Lead User" : "Main Project",
      big: state.mode === "batch"
        ? getTopUserLabel(rows)
        : byProject[0]
          ? byProject[0].label
          : "-",
      small: state.mode === "batch"
        ? `${new Set(rows.map((row) => row.username)).size} user(s) currently included.`
        : byProject[0]
          ? `${formatHours(byProject[0].value)} accumulated.`
          : "Not enough data."
    },
    {
      title: "Top Client",
      big: byCustomer[0] ? byCustomer[0].label : "-",
      small: `${uniqueCustomers} customer(s) in the current selection.`
    },
    {
      title: "Average Record Duration",
      big: formatHours(averageEntry),
      small: `Each individual record averaged ${averageEntry.toFixed(2)} hours.`
    },
    {
      title: "Longest Record",
      big: formatHours(longestEntry.hours),
      small: `${formatDate(longestEntry.dateObj)} · ${longestEntry.project} · ${longestEntry.activity}`
    },
    {
      title: "Active Portfolio",
      big: String(uniqueProjects),
      small: "Distinct projects detected in the current selection."
    }
  ]);

  renderTable(rows);
}

function buildPeriodBadge(firstDate, lastDate, rows) {
  const rangeLabel = `${formatDate(firstDate)} -> ${formatDate(lastDate)}`;

  if (state.mode !== "batch") {
    return `${rangeLabel} · ${state.fileNames[0] || "Single file"}`;
  }

  const totalUsers = getAvailableUsers().length;
  const selectedUsers = new Set(rows.map((row) => row.username)).size;
  return `${rangeLabel} · ${state.fileNames.length} file(s) · ${selectedUsers}/${totalUsers} user(s)`;
}

function getTopUserLabel(rows) {
  const totals = groupSum(rows, "username");
  return totals[0] ? totals[0].label : "-";
}

function resetDashboard(emptyTableMessage) {
  elements.periodBadge.textContent = state.allRows.length ? "No users selected" : "No data loaded yet";
  elements.kpiTotal.textContent = "0 h";
  elements.kpiTotalSub.textContent = "Upload a file to calculate this indicator";
  elements.kpiAvg.textContent = "0 h";
  elements.kpiAvgSub.textContent = "Based on days with logged activity";
  elements.kpiBest.textContent = "-";
  elements.kpiBestSub.textContent = "The date and total load will appear here";
  elements.kpiBillable.textContent = "0%";
  elements.kpiBillableSub.textContent = "Calculated from the Billable column";
  state.activeDailyMonthKey = "";
  renderDailyDistribution([]);
  renderBars(elements.projectChart, [], 8);
  renderBars(elements.activityChart, [], 8);
  renderQuickStats([]);
  renderTable([], emptyTableMessage);
}

function renderDailyDistribution(daily, options) {
  const settings = options || {};
  const months = getDailyMonths(daily);
  const hasMonths = months.length > 0;

  elements.dailyPrevMonthButton.disabled = !hasMonths;
  elements.dailyNextMonthButton.disabled = !hasMonths;

  if (!hasMonths) {
    elements.dailyMonthLabel.textContent = "No month selected";
    renderBars(elements.dailyChart, [], 31);
    return;
  }

  if (settings.showAllMonths) {
    elements.dailyMonthLabel.textContent = `${months.length} month(s) included`;
    elements.dailyPrevMonthButton.disabled = true;
    elements.dailyNextMonthButton.disabled = true;

    const allRows = daily.map((item) => ({
      label: formatDateShort(item.dateObj),
      value: item.hours
    }));

    renderBars(elements.dailyChart, allRows, allRows.length);
    return;
  }

  syncDailyMonthSelection(months);
  const monthIndex = months.findIndex((month) => month.key === state.activeDailyMonthKey);
  const currentMonth = months[monthIndex];
  const monthRows = daily
    .filter((item) => getMonthKey(item.dateObj) === currentMonth.key)
    .map((item) => ({ label: formatDateShort(item.dateObj), value: item.hours }));

  elements.dailyMonthLabel.textContent = currentMonth.label;
  elements.dailyPrevMonthButton.disabled = monthIndex <= 0;
  elements.dailyNextMonthButton.disabled = monthIndex >= months.length - 1;
  renderBars(elements.dailyChart, monthRows, monthRows.length);
}

function getDailyMonths(daily) {
  const seen = new Map();

  daily.forEach((item) => {
    const key = getMonthKey(item.dateObj);
    if (!seen.has(key)) {
      seen.set(key, {
        key,
        label: formatMonthLabel(item.dateObj)
      });
    }
  });

  return Array.from(seen.values()).sort((left, right) => left.key.localeCompare(right.key));
}

function syncDailyMonthSelection(months) {
  const hasCurrent = months.some((month) => month.key === state.activeDailyMonthKey);
  if (!hasCurrent) {
    state.activeDailyMonthKey = months[months.length - 1].key;
  }
}

function shiftDailyMonth(direction) {
  const daily = getVisibleDailyRows();
  const months = getDailyMonths(daily);
  if (!months.length) {
    return;
  }

  syncDailyMonthSelection(months);
  const currentIndex = months.findIndex((month) => month.key === state.activeDailyMonthKey);
  const nextIndex = Math.min(Math.max(currentIndex + direction, 0), months.length - 1);

  if (nextIndex === currentIndex) {
    return;
  }

  state.activeDailyMonthKey = months[nextIndex].key;
  renderDailyDistribution(daily);
}

function getVisibleDailyRows() {
  const byDate = new Map();
  const rows = getVisibleRows();

  rows.forEach((row) => {
    if (!byDate.has(row.dateKey)) {
      byDate.set(row.dateKey, { dateObj: row.dateObj, hours: 0 });
    }

    byDate.get(row.dateKey).hours += row.hours;
  });

  return Array.from(byDate.values()).sort((left, right) => left.dateObj - right.dateObj);
}

function renderDailyDistributionForPdf() {
  const daily = getVisibleDailyRows();
  renderDailyDistribution(daily, { showAllMonths: true });
}

function renderBars(container, items, limit) {
  container.replaceChildren();
  const dataset = items.slice(0, limit);

  if (!dataset.length) {
    const empty = document.createElement("div");
    empty.className = "empty";
    empty.textContent = "There is not enough data to render this chart.";
    container.appendChild(empty);
    return;
  }

  const maxValue = Math.max(...dataset.map((item) => item.value), 1);

  dataset.forEach((item) => {
    const row = document.createElement("div");
    row.className = "bar-row";

    const label = document.createElement("div");
    label.className = "bar-label";
    label.title = item.label;
    label.textContent = item.label;

    const track = document.createElement("div");
    track.className = "bar-track";

    const fill = document.createElement("div");
    fill.className = "bar-fill";
    fill.style.width = `${(item.value / maxValue) * 100}%`;
    track.appendChild(fill);

    const value = document.createElement("div");
    value.className = "bar-value";
    value.textContent = `${item.value.toFixed(1)} h`;

    row.append(label, track, value);
    container.appendChild(row);
  });
}

function renderQuickStats(items) {
  elements.quickStats.replaceChildren();

  if (!items.length) {
    const empty = document.createElement("div");
    empty.className = "empty";
    empty.textContent = "There is not enough data to generate quick reads.";
    elements.quickStats.appendChild(empty);
    return;
  }

  items.forEach((item) => {
    const card = document.createElement("article");
    card.className = "stat-item";

    const title = document.createElement("div");
    title.className = "title";
    title.textContent = item.title;

    const big = document.createElement("div");
    big.className = "big";
    big.textContent = item.big;

    const small = document.createElement("div");
    small.className = "small";
    small.textContent = item.small;

    card.append(title, big, small);
    elements.quickStats.appendChild(card);
  });
}

// The detail table is built with DOM nodes instead of template strings so
// uploaded spreadsheet content is rendered as text, not executable markup.
function renderTable(rows, emptyMessage = "There are no records to display yet.") {
  if (!rows.length) {
    elements.tableContainer.className = "empty";
    elements.tableContainer.textContent = emptyMessage;
    renderTablePagination(0, 0, 0);
    return;
  }

  const headers = ["Date", "In", "Out", "Hours", "Customer", "Project", "Activity", "Description", "User"];
  const sorted = [...rows].sort((left, right) => right.dateObj - left.dateObj || right.inTimestamp - left.inTimestamp);
  const pageSize = 100;
  const shouldPaginate = state.mode === "batch" && sorted.length > pageSize;
  const totalPages = shouldPaginate ? Math.ceil(sorted.length / pageSize) : 1;
  state.currentTablePage = Math.min(Math.max(state.currentTablePage, 1), totalPages);
  const paginatedRows = shouldPaginate
    ? sorted.slice((state.currentTablePage - 1) * pageSize, state.currentTablePage * pageSize)
    : sorted;
  const table = document.createElement("table");
  const thead = document.createElement("thead");
  const headerRow = document.createElement("tr");

  headers.forEach((header) => {
    const th = document.createElement("th");
    th.textContent = header;
    headerRow.appendChild(th);
  });

  thead.appendChild(headerRow);

  const tbody = document.createElement("tbody");
  paginatedRows.forEach((row) => {
    const tr = document.createElement("tr");
    const cells = [
      formatDate(row.dateObj),
      formatTime(row.inObj),
      formatTime(row.outObj),
      `${row.hours.toFixed(2)} h`,
      row.customer,
      row.project,
      row.activity,
      row.description,
      row.username
    ];

    cells.forEach((cellValue) => {
      const td = document.createElement("td");
      td.textContent = cellValue;
      tr.appendChild(td);
    });

    tbody.appendChild(tr);
  });

  table.append(thead, tbody);
  elements.tableContainer.className = "table-wrap";
  elements.tableContainer.replaceChildren(table);
  renderTablePagination(sorted.length, totalPages, shouldPaginate ? state.currentTablePage : 1);
}

function renderTablePagination(totalRows, totalPages, currentPage) {
  const shouldShow = state.mode === "batch" && totalRows > 100;
  elements.tablePagination.classList.toggle("is-hidden", !shouldShow);
  elements.tablePagination.replaceChildren();

  if (!shouldShow) {
    return;
  }

  const summary = document.createElement("p");
  summary.className = "table-pagination__summary";

  const start = (currentPage - 1) * 100 + 1;
  const end = Math.min(currentPage * 100, totalRows);
  summary.textContent = `Showing ${start}-${end} of ${totalRows} records`;

  const actions = document.createElement("div");
  actions.className = "table-pagination__actions";

  const prev = document.createElement("button");
  prev.type = "button";
  prev.className = "button secondary small";
  prev.textContent = "Previous";
  prev.disabled = currentPage <= 1;
  prev.addEventListener("click", () => {
    if (state.currentTablePage <= 1) {
      return;
    }
    state.currentTablePage -= 1;
    renderCurrentView({ animate: false });
  });

  const next = document.createElement("button");
  next.type = "button";
  next.className = "button secondary small";
  next.textContent = "Next";
  next.disabled = currentPage >= totalPages;
  next.addEventListener("click", () => {
    if (state.currentTablePage >= totalPages) {
      return;
    }
    state.currentTablePage += 1;
    renderCurrentView({ animate: false });
  });

  actions.append(prev, next);
  elements.tablePagination.append(summary, actions);
}

function collectTypewriterTargets() {
  const staticTargets = Array.from(document.querySelectorAll("[data-typed-static='true']"));
  const dynamicTargets = [
    elements.status,
    elements.periodBadge,
    elements.kpiTotal,
    elements.kpiTotalSub,
    elements.kpiAvg,
    elements.kpiAvgSub,
    elements.kpiBest,
    elements.kpiBestSub,
    elements.kpiBillable,
    elements.kpiBillableSub,
    ...Array.from(elements.quickStats.querySelectorAll(".title, .big, .small"))
  ];

  return [...staticTargets, ...dynamicTargets];
}

function animateTopSectionAfterUpload() {
  const runId = Date.now();
  state.activeTypingRun = runId;

  const targets = collectTypewriterTargets()
    .filter(Boolean)
    .filter((element, index, list) => list.indexOf(element) === index)
    .filter((element) => element.textContent.trim() !== "");

  targets.forEach((element) => {
    element.dataset.typewriterText = element.textContent;
    element.classList.add("typewriter-pending");
  });

  const reducedMotion = window.matchMedia("(prefers-reduced-motion: reduce)").matches;
  if (reducedMotion) {
    targets.forEach((element) => {
      element.classList.remove("typewriter-pending");
    });
    return;
  }

  let delay = 0;
  targets.forEach((element) => {
    const lengthFactor = Math.min(element.dataset.typewriterText.length * 8, 320);
    window.setTimeout(() => {
      typewriteElement(element, element.dataset.typewriterText, runId);
    }, delay);
    delay += Math.min(28 + lengthFactor * 0.05, 70);
  });
}

function typewriteElement(element, text, runId) {
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
    const charDelay = nextCharacter === " " ? 4 : 5;
    window.setTimeout(step, charDelay);
  };

  step();
}

async function exportPdf() {
  const rows = getVisibleRows();
  if (!rows.length) {
    setStatus(
      state.mode === "batch"
        ? "Select at least one user with available records before exporting to PDF."
        : "Upload a file before exporting the dashboard to PDF."
    );
    return;
  }

  setStatus("Generating a print-friendly PDF...");

  document.documentElement.classList.add("pdf-mode");
  document.body.classList.add("pdf-mode");
  elements.dashboard.classList.add("pdf-mode");
  renderDailyDistributionForPdf();

  const options = {
    margin: [8, 8, 8, 8],
    filename: buildPdfFilename(rows),
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
      orientation: "portrait"
    },
    pagebreak: {
      mode: ["avoid-all", "css", "legacy"],
      avoid: [".card", ".kpi", ".panel", ".stat-item", ".bar-row", "tr"]
    }
  };

  try {
    await new Promise((resolve) => window.setTimeout(resolve, 180));
    await html2pdf().set(options).from(elements.dashboard).save();
    setStatus("The PDF was exported successfully.");
  } catch (error) {
    console.error(error);
    setStatus("The PDF export failed.");
  } finally {
    renderDailyDistribution(getVisibleDailyRows());
    document.documentElement.classList.remove("pdf-mode");
    document.body.classList.remove("pdf-mode");
    elements.dashboard.classList.remove("pdf-mode");
  }
}
