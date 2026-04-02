const CLIENT_OVERVIEW_STORAGE_KEY = "excel_pdf_time_converter_client_overview";
const CLIENT_COLOR_PALETTE = [
  "#ff6b6b",
  "#4ecdc4",
  "#ffd166",
  "#5dade2",
  "#f78fb3",
  "#2ed573",
  "#ff9f43",
  "#7d5fff",
  "#3dc1d3",
  "#ff7f50"
];
const PROJECT_COLOR_PALETTE = [
  "#f6d6d6",
  "#d8ecff",
  "#dff4dd",
  "#f9ebc8",
  "#eadcff",
  "#d9f1f0",
  "#ffe2cc",
  "#e6e8ff",
  "#fbe4ef",
  "#dff0f7"
];

const state = {
  rows: [],
  topClientsMode: "3",
  metrics: null,
  colorMapByClient: new Map(),
  colorMapByProject: new Map(),
  charts: []
};

const elements = {
  overviewStatus: document.getElementById("overviewStatus"),
  overviewBadge: document.getElementById("overviewBadge"),
  overviewDescription: document.getElementById("overviewDescription"),
  overviewEmptyState: document.getElementById("overviewEmptyState"),
  topClientsControl: document.getElementById("topClientsControl"),
  summaryTableContainer: document.getElementById("summaryTableContainer"),
  kpiTotalHours: document.getElementById("kpiTotalHours"),
  kpiTotalHoursSub: document.getElementById("kpiTotalHoursSub"),
  kpiTotalClients: document.getElementById("kpiTotalClients"),
  kpiTotalClientsSub: document.getElementById("kpiTotalClientsSub"),
  kpiActiveDays: document.getElementById("kpiActiveDays"),
  kpiActiveDaysSub: document.getElementById("kpiActiveDaysSub"),
  kpiAvgDailyLoad: document.getElementById("kpiAvgDailyLoad"),
  kpiAvgDailyLoadSub: document.getElementById("kpiAvgDailyLoadSub"),
  kpiTopClient: document.getElementById("kpiTopClient"),
  kpiTopClientSub: document.getElementById("kpiTopClientSub"),
  kpiTop3Concentration: document.getElementById("kpiTop3Concentration"),
  kpiTop3ConcentrationSub: document.getElementById("kpiTop3ConcentrationSub"),
  chartTotalHoursByClient: document.getElementById("chartTotalHoursByClient"),
  chartHoursDistributionByClient: document.getElementById("chartHoursDistributionByClient"),
  chartTotalWorkloadTrend: document.getElementById("chartTotalWorkloadTrend"),
  chartDailyHoursByTopClients: document.getElementById("chartDailyHoursByTopClients"),
  chartActivityDistributionByClient: document.getElementById("chartActivityDistributionByClient"),
  chartGlobalHoursByActivity: document.getElementById("chartGlobalHoursByActivity"),
  chartAvgHoursPerDayByClient: document.getElementById("chartAvgHoursPerDayByClient"),
  chartActiveDaysByClient: document.getElementById("chartActiveDaysByClient"),
  chartAvgRecordDurationByClient: document.getElementById("chartAvgRecordDurationByClient")
};

init();

function init() {
  bindEvents();
  loadSnapshot();
}

function bindEvents() {
  elements.topClientsControl.addEventListener("click", (event) => {
    const button = event.target.closest("[data-top-clients]");
    if (!button) {
      return;
    }

    state.topClientsMode = button.dataset.topClients;
    syncTopClientsButtons();
    renderTopClientsTrend(state.metrics);
  });
}

function loadSnapshot() {
  let snapshot = null;

  try {
    snapshot = JSON.parse(window.sessionStorage.getItem(CLIENT_OVERVIEW_STORAGE_KEY) || "null");
  } catch (error) {
    console.error("Failed to read client overview snapshot.", error);
  }

  if (!snapshot || !Array.isArray(snapshot.rows) || !snapshot.rows.length) {
    showEmptyState();
    return;
  }

  state.rows = snapshot.rows
    .map((row) => ({
      dateKey: row.dateKey,
      dateObj: row.dateObj ? new Date(row.dateObj) : null,
      hours: Number(row.hours) || 0,
      billableHours: Number(row.billableHours) || 0,
      customer: safeValue(row.customer, "No client"),
      project: safeValue(row.project, "No project"),
      activity: safeValue(row.activity, "No activity"),
      description: safeValue(row.description, "-"),
      username: safeValue(row.username, "-"),
      sourceFile: safeValue(row.sourceFile, "-")
    }))
    .filter((row) => row.dateObj instanceof Date && !Number.isNaN(row.dateObj.getTime()) && row.hours > 0)
    .sort((left, right) => left.dateObj - right.dateObj);

  if (!state.rows.length) {
    showEmptyState();
    return;
  }

  hideEmptyState();
  renderOverview(snapshot);
}

function renderOverview(snapshot) {
  const metrics = buildMetrics(state.rows);
  state.metrics = metrics;
  state.colorMapByClient = buildColorMap(metrics.clientSummaries.map((item) => item.client), CLIENT_COLOR_PALETTE);
  state.colorMapByProject = buildColorMap(getProjectLabels(state.rows), PROJECT_COLOR_PALETTE);
  const selectedUsersLabel = snapshot.selectedUsers && snapshot.selectedUsers.length
    ? `${snapshot.selectedUsers.length} user(s) selected`
    : "All available users";

  elements.overviewStatus.textContent = `Snapshot loaded from Batch. ${snapshot.fileNames.length} file(s) and ${state.rows.length} visible records included.`;
  elements.overviewBadge.textContent = `${metrics.firstDateLabel} -> ${metrics.lastDateLabel}`;
  elements.overviewDescription.textContent = `This view reflects the current Batch selection. ${selectedUsersLabel}. Saved at ${formatTimestamp(snapshot.savedAt)}.`;

  elements.kpiTotalHours.textContent = formatHours(metrics.totalHours);
  elements.kpiTotalHoursSub.textContent = `${metrics.totalRecords} visible records`;
  elements.kpiTotalClients.textContent = String(metrics.clientSummaries.length);
  elements.kpiTotalClientsSub.textContent = "Clients with recorded activity";
  elements.kpiActiveDays.textContent = String(metrics.activeDays);
  elements.kpiActiveDaysSub.textContent = `${metrics.dailyTotals.length} daily observations`;
  elements.kpiAvgDailyLoad.textContent = formatHours(metrics.averageDailyLoad);
  elements.kpiAvgDailyLoadSub.textContent = "Average hours per active day";
  elements.kpiTopClient.textContent = metrics.topClient ? metrics.topClient.client : "-";
  elements.kpiTopClientSub.textContent = metrics.topClient ? `${formatHours(metrics.topClient.totalHours)} accumulated.` : "No client data available.";
  elements.kpiTop3Concentration.textContent = formatPercent(metrics.top3Concentration);
  elements.kpiTop3ConcentrationSub.textContent = "Share of hours in the top 3 clients";

  renderAllCharts(metrics);
  renderSummaryTable(metrics.clientSummaries);
  syncTopClientsButtons();
}

function buildMetrics(rows) {
  const clientMap = new Map();
  const activityMap = new Map();
  const dailyTotalsMap = new Map();
  const clientDailyMap = new Map();
  const clientActivityMap = new Map();

  rows.forEach((row) => {
    const clientKey = row.customer;

    if (!clientMap.has(clientKey)) {
      clientMap.set(clientKey, {
        client: clientKey,
        totalHours: 0,
        recordCount: 0,
        activeDays: new Set(),
        activityHours: new Map()
      });
    }

    const client = clientMap.get(clientKey);
    client.totalHours += row.hours;
    client.recordCount += 1;
    client.activeDays.add(row.dateKey);
    client.activityHours.set(row.activity, (client.activityHours.get(row.activity) || 0) + row.hours);

    activityMap.set(row.activity, (activityMap.get(row.activity) || 0) + row.hours);
    const existingDailyTotal = dailyTotalsMap.get(row.dateKey);
    dailyTotalsMap.set(row.dateKey, {
      dateKey: row.dateKey,
      dateObj: row.dateObj,
      hours: (existingDailyTotal ? existingDailyTotal.hours : 0) + row.hours
    });

    if (!clientDailyMap.has(clientKey)) {
      clientDailyMap.set(clientKey, new Map());
    }
    const clientDaily = clientDailyMap.get(clientKey);
    const existingClientDaily = clientDaily.get(row.dateKey);
    clientDaily.set(row.dateKey, {
      dateKey: row.dateKey,
      dateObj: row.dateObj,
      hours: (existingClientDaily ? existingClientDaily.hours : 0) + row.hours
    });

    if (!clientActivityMap.has(clientKey)) {
      clientActivityMap.set(clientKey, new Map());
    }
    const clientActivities = clientActivityMap.get(clientKey);
    clientActivities.set(row.activity, (clientActivities.get(row.activity) || 0) + row.hours);
  });

  const clientSummaries = Array.from(clientMap.values())
    .map((client) => ({
      client: client.client,
      totalHours: client.totalHours,
      shareOfTotal: 0,
      activeDays: client.activeDays.size,
      avgHoursPerDay: client.activeDays.size ? client.totalHours / client.activeDays.size : 0,
      avgHoursPerRecord: client.recordCount ? client.totalHours / client.recordCount : 0,
      recordCount: client.recordCount,
      topActivity: getTopMapEntry(client.activityHours)
        ? getTopMapEntry(client.activityHours).label
        : "-"
    }))
    .sort((left, right) => right.totalHours - left.totalHours);

  const totalHours = clientSummaries.reduce((sum, client) => sum + client.totalHours, 0);
  clientSummaries.forEach((client) => {
    client.shareOfTotal = totalHours ? (client.totalHours / totalHours) * 100 : 0;
  });

  const dailyTotals = Array.from(dailyTotalsMap.values()).sort((left, right) => left.dateObj - right.dateObj);
  const activities = Array.from(activityMap, ([activity, hours]) => ({ activity, hours }))
    .sort((left, right) => right.hours - left.hours);

  const topClient = clientSummaries[0] || null;
  const top3Concentration = totalHours
    ? clientSummaries.slice(0, 3).reduce((sum, client) => sum + client.totalHours, 0) / totalHours * 100
    : 0;

  return {
    totalHours,
    totalRecords: rows.length,
    clientSummaries,
    topClient,
    top3Concentration,
    activeDays: dailyTotals.length,
    averageDailyLoad: dailyTotals.length ? totalHours / dailyTotals.length : 0,
    firstDateLabel: dailyTotals.length ? formatDate(dailyTotals[0].dateObj) : "-",
    lastDateLabel: dailyTotals.length ? formatDate(dailyTotals[dailyTotals.length - 1].dateObj) : "-",
    dailyTotals,
    clientDailyMap,
    clientActivityMap,
    activities
  };
}

function renderAllCharts(metrics) {
  destroyCharts();

  const topClients = metrics.clientSummaries.slice(0, 8);
  const clientLabels = topClients.map((item) => item.client);
  const clientHours = topClients.map((item) => round(item.totalHours));
  const clientColors = clientLabels.map((label) => getClientColor(label));

  state.charts.push(
    createBarChart(elements.chartTotalHoursByClient, {
      horizontal: true,
      categories: clientLabels,
      series: [{ name: "Hours", data: clientHours }],
      colors: clientColors
    })
  );

  state.charts.push(
    createDonutChart(elements.chartHoursDistributionByClient, {
      labels: topClients.map((item) => item.client),
      series: topClients.map((item) => round(item.totalHours)),
      colors: clientColors
    })
  );

  state.charts.push(
    createLineChart(elements.chartTotalWorkloadTrend, {
      categories: metrics.dailyTotals.map((item) => formatDateShort(item.dateObj)),
      series: [{ name: "Total Hours", data: metrics.dailyTotals.map((item) => round(item.hours)) }],
      colors: ["#4ecdc4"]
    })
  );

  renderTopClientsTrend(metrics);

  state.charts.push(
    createStackedBarChart(elements.chartActivityDistributionByClient, buildActivityDistributionSeries(metrics))
  );

  state.charts.push(
    createBarChart(elements.chartGlobalHoursByActivity, {
      horizontal: true,
      categories: metrics.activities.slice(0, 10).map((item) => item.activity),
      series: [{ name: "Hours", data: metrics.activities.slice(0, 10).map((item) => round(item.hours)) }],
      colors: PROJECT_COLOR_PALETTE
    })
  );

  const performanceTop = metrics.clientSummaries.slice(0, 8);
  state.charts.push(
    createBarChart(elements.chartAvgHoursPerDayByClient, {
      categories: performanceTop.map((item) => item.client),
      series: [{ name: "Avg Hours / Day", data: performanceTop.map((item) => round(item.avgHoursPerDay)) }],
      colors: performanceTop.map((item) => getClientColor(item.client))
    })
  );

  state.charts.push(
    createBarChart(elements.chartActiveDaysByClient, {
      horizontal: true,
      categories: performanceTop.map((item) => item.client),
      series: [{ name: "Active Days", data: performanceTop.map((item) => item.activeDays) }],
      colors: performanceTop.map((item) => getClientColor(item.client))
    })
  );

  state.charts.push(
    createBarChart(elements.chartAvgRecordDurationByClient, {
      categories: performanceTop.map((item) => item.client),
      series: [{ name: "Avg Hours / Record", data: performanceTop.map((item) => round(item.avgHoursPerRecord)) }],
      colors: performanceTop.map((item) => getClientColor(item.client))
    })
  );
}

function renderTopClientsTrend(metrics) {
  if (!metrics) {
    return;
  }

  const existing = state.charts.find((chart) => chart.el === elements.chartDailyHoursByTopClients);
  if (existing) {
    existing.destroy();
    state.charts = state.charts.filter((chart) => chart !== existing);
  }

  const series = buildTopClientsTrendSeries(metrics);
  state.charts.push(
    createLineChart(elements.chartDailyHoursByTopClients, {
      categories: metrics.dailyTotals.map((item) => formatDateShort(item.dateObj)),
      series,
      colors: series.map((item) => getClientColor(item.name))
    })
  );
}

function buildTopClientsTrendSeries(metrics) {
  const sortedClients = [...metrics.clientSummaries];
  let selectedClients = [];

  if (state.topClientsMode === "all") {
    selectedClients = sortedClients.slice(0, 8);
  } else {
    selectedClients = sortedClients.slice(0, Number(state.topClientsMode));
  }

  return selectedClients.map((client) => {
    const clientDaily = metrics.clientDailyMap.get(client.client) || new Map();
    return {
      name: client.client,
      data: metrics.dailyTotals.map((day) => {
        const dayRecord = clientDaily.get(day.dateKey);
        return round(dayRecord ? dayRecord.hours : 0);
      })
    };
  });
}

function buildActivityDistributionSeries(metrics) {
  const topClients = metrics.clientSummaries.slice(0, 6).map((item) => item.client);
  const activities = metrics.activities.slice(0, 5).map((item) => item.activity);

  return {
    categories: topClients,
    series: activities.map((activity) => ({
      name: activity,
      data: topClients.map((client) => {
        const clientActivities = metrics.clientActivityMap.get(client) || new Map();
        return round(clientActivities.get(activity) || 0);
      })
    })),
    colors: activities.map((_, index) => PROJECT_COLOR_PALETTE[index % PROJECT_COLOR_PALETTE.length])
  };
}

function renderSummaryTable(clientSummaries) {
  const table = document.createElement("table");
  const thead = document.createElement("thead");
  const headerRow = document.createElement("tr");
  const headers = [
    "Client",
    "Total Hours",
    "% of Total",
    "Active Days",
    "Avg Hours / Day",
    "Avg Hours / Record",
    "Record Count",
    "Top Activity"
  ];

  headers.forEach((header) => {
    const th = document.createElement("th");
    th.textContent = header;
    headerRow.appendChild(th);
  });

  thead.appendChild(headerRow);

  const tbody = document.createElement("tbody");
  clientSummaries.forEach((client) => {
    const tr = document.createElement("tr");
    [
      client.client,
      formatHours(client.totalHours),
      formatPercent(client.shareOfTotal),
      String(client.activeDays),
      formatHours(client.avgHoursPerDay),
      formatHours(client.avgHoursPerRecord),
      String(client.recordCount),
      client.topActivity
    ].forEach((value) => {
      const td = document.createElement("td");
      td.textContent = value;
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  table.append(thead, tbody);
  elements.summaryTableContainer.classList.add("summary-table-wrap");
  elements.summaryTableContainer.replaceChildren(table);
}

function destroyCharts() {
  state.charts.forEach((chart) => chart.destroy());
  state.charts = [];
}

function createBaseOptions() {
  return {
    chart: {
      background: "transparent",
      zoom: {
        enabled: false
      },
      selection: {
        enabled: false
      },
      toolbar: { show: false },
      foreColor: "#b2b2b8"
    },
    dataLabels: { enabled: false },
    stroke: { width: 3, curve: "smooth" },
    legend: {
      position: "bottom",
      labels: { colors: "#b2b2b8" }
    },
    grid: {
      borderColor: "rgba(255, 255, 255, 0.08)"
    },
    tooltip: {
      theme: "dark"
    },
    colors: ["#ffffff", "#d0d0d5", "#9fa4af", "#6f7682", "#4f5561", "#f2c94c", "#87d37c", "#56ccf2"]
  };
}

function createBarChart(element, config) {
  const options = createBaseOptions();
  const chart = new ApexCharts(element, {
    ...options,
    chart: {
      ...options.chart,
      type: "bar",
      height: "100%"
    },
    colors: config.colors || options.colors,
    plotOptions: {
      bar: {
        horizontal: Boolean(config.horizontal),
        distributed: Array.isArray(config.colors) && config.series.length === 1,
        borderRadius: 6,
        columnWidth: "48%"
      }
    },
    series: config.series,
    xaxis: {
      categories: config.categories,
      labels: { style: { colors: "#8a8a92" } }
    },
    yaxis: {
      labels: { style: { colors: "#8a8a92" } }
    }
  });

  chart.render();
  chart.el = element;
  return chart;
}

function createLineChart(element, config) {
  const options = createBaseOptions();
  const chart = new ApexCharts(element, {
    ...options,
    chart: {
      ...options.chart,
      type: "line",
      height: "100%"
    },
    colors: config.colors || options.colors,
    series: config.series,
    xaxis: {
      categories: config.categories,
      labels: {
        style: { colors: "#8a8a92" },
        rotate: -45
      }
    },
    yaxis: {
      labels: { style: { colors: "#8a8a92" } }
    }
  });

  chart.render();
  chart.el = element;
  return chart;
}

function createDonutChart(element, config) {
  const options = createBaseOptions();
  const chart = new ApexCharts(element, {
    ...options,
    chart: {
      ...options.chart,
      type: "donut",
      height: "100%"
    },
    colors: config.colors || options.colors,
    labels: config.labels,
    series: config.series,
    stroke: { colors: ["transparent"] },
    dataLabels: { enabled: false },
    plotOptions: {
      pie: {
        donut: {
          labels: {
            show: true,
            value: {
              color: "#f5f5f5"
            },
            total: {
              show: true,
              color: "#b2b2b8",
              label: "Total Hours"
            }
          }
        }
      }
    },
    legend: {
      position: "bottom",
      labels: { colors: "#b2b2b8" }
    }
  });

  chart.render();
  chart.el = element;
  return chart;
}

function createStackedBarChart(element, config) {
  const options = createBaseOptions();
  const chart = new ApexCharts(element, {
    ...options,
    chart: {
      ...options.chart,
      type: "bar",
      stacked: true,
      height: "100%"
    },
    colors: config.colors || options.colors,
    plotOptions: {
      bar: {
        horizontal: false,
        borderRadius: 4,
        columnWidth: "52%"
      }
    },
    series: config.series,
    xaxis: {
      categories: config.categories,
      labels: { style: { colors: "#8a8a92" } }
    },
    yaxis: {
      labels: { style: { colors: "#8a8a92" } }
    }
  });

  chart.render();
  chart.el = element;
  return chart;
}

function syncTopClientsButtons() {
  Array.from(elements.topClientsControl.querySelectorAll("[data-top-clients]")).forEach((button) => {
    button.classList.toggle("active", button.dataset.topClients === state.topClientsMode);
  });
}

function showEmptyState() {
  elements.overviewEmptyState.classList.remove("is-hidden");
  elements.overviewStatus.textContent = "Waiting for batch data. Open this page from Batch after loading Excel files.";
  elements.overviewBadge.textContent = "No batch data available";
}

function hideEmptyState() {
  elements.overviewEmptyState.classList.add("is-hidden");
}

function getTopMapEntry(map) {
  return Array.from(map, ([label, value]) => ({ label, value }))
    .sort((left, right) => right.value - left.value)[0] || null;
}

function formatHours(value) {
  return `${round(value).toFixed(1)} h`;
}

function formatPercent(value) {
  return `${round(value).toFixed(1)}%`;
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
    month: "short",
    day: "2-digit"
  }).format(date);
}

function formatTimestamp(value) {
  if (!value) {
    return "an unknown time";
  }

  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return "an unknown time";
  }

  return new Intl.DateTimeFormat("en-US", {
    month: "short",
    day: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit"
  }).format(date);
}

function safeValue(value, fallback) {
  return value === undefined || value === null || String(value).trim() === "" ? fallback : String(value);
}

function round(value) {
  return Math.round(value * 10) / 10;
}

function buildColorMap(labels, palette) {
  const uniqueLabels = Array.from(new Set(labels));
  const map = new Map();

  uniqueLabels.forEach((label, index) => {
    map.set(label, palette[index % palette.length]);
  });

  return map;
}

function getClientColor(label) {
  return state.colorMapByClient.get(label) || CLIENT_COLOR_PALETTE[0];
}

function getProjectLabels(rows) {
  return rows.map((row) => row.project).filter(Boolean);
}
