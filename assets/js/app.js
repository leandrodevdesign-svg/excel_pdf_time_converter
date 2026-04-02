// This application keeps all state in memory because the dashboard is meant
// to be a lightweight client-side tool with no backend dependencies.
const state = {
  rows: [],
  activeTypingRun: 0
};

// Cache DOM references once so the render functions stay small and predictable.
const elements = {
  pageLoader: document.getElementById("pageLoader"),
  excelInput: document.getElementById("excelInput"),
  pdfButton: document.getElementById("pdfBtn"),
  status: document.getElementById("status"),
  periodBadge: document.getElementById("periodBadge"),
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
  dashboard: document.getElementById("dashboard")
};

elements.excelInput.addEventListener("change", handleFileUpload);
elements.pdfButton.addEventListener("click", exportPdf);
window.addEventListener("load", hideInitialLoader, { once: true });

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
  const file = event.target.files?.[0];
  if (!file) {
    return;
  }

  setStatus(`Processing ${file.name}...`);

  try {
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
      throw new Error("The file is empty.");
    }

    const headers = matrix[0].map(normalizeHeader);
    const rows = matrix
      .slice(1)
      .filter((row) => row.some((cell) => String(cell ?? "").trim() !== ""))
      .map((row) =>
        row.reduce((record, cell, index) => {
          record[headers[index] || `col_${index}`] = cell;
          return record;
        }, {})
      )
      .map(parseRow)
      .filter(Boolean)
      .sort((left, right) => left.dateObj - right.dateObj || left.inTimestamp - right.inTimestamp);

    if (!rows.length) {
      throw new Error("No valid records were found in the uploaded file.");
    }

    state.rows = rows;
    renderDashboard(rows, file.name);
    setStatus(`Loaded ${file.name} successfully. Processed ${rows.length} records.`);
    animateTopSectionAfterUpload();
  } catch (error) {
    console.error(error);
    setStatus(`The file could not be processed. ${error.message || "Check the format and try again."}`);
  }
}

// The parser accepts a few alternate field names so the app remains useful
// even when spreadsheet exports differ between systems or locales.
function parseRow(row) {
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
    username: safeValue(row.username || row.usuario || row.user, "-")
  };
}

function getDateKey(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function calculateHoursFromDates(start, end) {
  if (!start || !end) {
    return 0;
  }

  const differenceInHours = (end - start) / 36e5;
  return differenceInHours > 0 ? differenceInHours : 0;
}

function buildPdfFilename(rows) {
  const usernames = rows
    .map((row) => row.username)
    .filter((value) => value && value !== "-");

  const primaryUsername = usernames[0] || "user";
  const sanitizedUsername = primaryUsername
    .toLowerCase()
    .replace(/[^a-z0-9\s-]/g, "")
    .trim()
    .replace(/\s+/g, "-") || "user";

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

  return `${sanitizedUsername}-${dateRange}.pdf`;
}

function groupSum(rows, key) {
  const groups = new Map();

  rows.forEach((row) => {
    const groupKey = safeValue(row[key]);
    groups.set(groupKey, (groups.get(groupKey) || 0) + row.hours);
  });

  return Array.from(groups, ([label, value]) => ({ label, value })).sort(compareByValueDesc);
}

function renderDashboard(rows, fileName) {
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
  const firstDate = daily[0]?.dateObj;
  const lastDate = daily[daily.length - 1]?.dateObj;

  elements.periodBadge.textContent = `${formatDate(firstDate)} -> ${formatDate(lastDate)} · ${fileName}`;

  elements.kpiTotal.textContent = formatHours(totalHours);
  elements.kpiTotalSub.textContent = `${rows.length} processed records`;
  elements.kpiAvg.textContent = formatHours(averagePerDay);
  elements.kpiAvgSub.textContent = `${daily.length} days with logged activity`;
  elements.kpiBest.textContent = formatHours(bestDay.hours);
  elements.kpiBestSub.textContent = `${formatDate(bestDay.dateObj)} was the busiest day`;
  elements.kpiBillable.textContent = formatPercent(billablePercentage);
  elements.kpiBillableSub.textContent = `${formatHours(totalBillable)} out of ${formatHours(totalHours)} were billable`;

  renderBars(elements.dailyChart, daily.map((item) => ({ label: formatDateShort(item.dateObj), value: item.hours })), 12);
  renderBars(elements.projectChart, byProject, 8);
  renderBars(elements.activityChart, byActivity, 8);

  renderQuickStats([
    {
      title: "Main Project",
      big: byProject[0] ? byProject[0].label : "-",
      small: byProject[0] ? `${formatHours(byProject[0].value)} accumulated.` : "Not enough data."
    },
    {
      title: "Top Customer",
      big: byCustomer[0] ? byCustomer[0].label : "-",
      small: `${uniqueCustomers} customer(s) in the selected period.`
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
      small: "Distinct projects detected in the uploaded file."
    }
  ]);

  renderTable(rows);
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
function renderTable(rows) {
  if (!rows.length) {
    elements.tableContainer.className = "empty";
    elements.tableContainer.textContent = "There are no records to display yet.";
    return;
  }

  const headers = ["Date", "In", "Out", "Hours", "Customer", "Project", "Activity", "Description", "User"];
  const sorted = [...rows].sort((left, right) => right.dateObj - left.dateObj || right.inTimestamp - left.inTimestamp);
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
  sorted.forEach((row) => {
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
  if (!state.rows.length) {
    setStatus("Upload a file before exporting the dashboard to PDF.");
    return;
  }

  setStatus("Generating a print-friendly PDF...");

  document.documentElement.classList.add("pdf-mode");
  document.body.classList.add("pdf-mode");
  elements.dashboard.classList.add("pdf-mode");

  const options = {
    margin: [8, 8, 8, 8],
    filename: buildPdfFilename(state.rows),
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
    // Small delay to ensure the PDF-specific styles are fully painted.
    await new Promise((resolve) => window.setTimeout(resolve, 180));
    await html2pdf().set(options).from(elements.dashboard).save();
    setStatus("The PDF was exported successfully.");
  } catch (error) {
    console.error(error);
    setStatus("The PDF export failed.");
  } finally {
    document.documentElement.classList.remove("pdf-mode");
    document.body.classList.remove("pdf-mode");
    elements.dashboard.classList.remove("pdf-mode");
  }
}
